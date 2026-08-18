using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authentication;
using System.Security.Claims;

Environment.SetEnvironmentVariable("DOTNET_USE_POLLING_FILE_WATCHER", "1");
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var builder = WebApplication.CreateBuilder(args);

builder.Configuration.Sources.Clear();
builder.Configuration.AddEnvironmentVariables();

builder.Services.AddRazorPages();
builder.Services.AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme)
    .AddCookie(options => {
        options.LoginPath = "/Login";
        options.AccessDeniedPath = "/AccessDenied";
        options.Cookie.Name = "AuthCookie";
    });

var connectionString = Environment.GetEnvironmentVariable("DATABASE_URL");
builder.Services.AddDbContext<AppDbContext>(options => options.UseNpgsql(connectionString));

var app = builder.Build();

app.UseStaticFiles();
app.UseRouting();
app.UseAuthentication();
app.UseAuthorization();
app.MapRazorPages();

// --- API XỬ LÝ ĐĂNG NHẬP ---
app.MapPost("/api/auth/login", async (AppDbContext db, HttpContext ctx, LoginRequest req) =>
{
    try
    {
        // 1. Tự động khởi tạo bảng TaiKhoans nếu chưa có
        string sqlCreateTaiKhoan = @"
            CREATE TABLE IF NOT EXISTS ""TaiKhoans"" (
                ""Id"" SERIAL PRIMARY KEY,
                ""MaTruong"" TEXT NOT NULL,
                ""TenTruong"" TEXT NOT NULL,
                ""MatKhau"" TEXT NOT NULL,
                ""Role"" TEXT NOT NULL
            );";
        await db.Database.ExecuteSqlRawAsync(sqlCreateTaiKhoan);

        // 2. Tự động khởi tạo bảng ThiSinhs nếu chưa có
        string sqlCreateThiSinh = @"
            CREATE TABLE IF NOT EXISTS ""ThiSinhs"" (
                ""Id"" SERIAL PRIMARY KEY,
                ""HoTen"" TEXT NOT NULL,
                ""CCCD"" TEXT NOT NULL,
                ""TenTruong"" TEXT NOT NULL,
                ""MaTruong"" TEXT NOT NULL,
                ""MonThi"" TEXT NOT NULL,
                ""NgayDangKy"" TIMESTAMP WITH TIME ZONE NOT NULL
            );";
        await db.Database.ExecuteSqlRawAsync(sqlCreateThiSinh);

        // 3. Khởi tạo tài khoản Sở mặc định nếu chưa có
        if (!await db.TaiKhoans.AnyAsync(x => x.Role == "So"))
        {
            db.TaiKhoans.Add(new TaiKhoan { MaTruong = "SO_GD", TenTruong = "Sở Giáo Dục & Đào Tạo", MatKhau = "so@123456", Role = "So" });
            await db.SaveChangesAsync();
        }

        var acc = await db.TaiKhoans.FirstOrDefaultAsync(x => x.MaTruong == req.MaTruong && x.MatKhau == req.MatKhau);
        if (acc == null) return Results.BadRequest(new { message = "Mã tài khoản hoặc mật khẩu không chính xác!" });

        var claims = new List<Claim>
        {
            new Claim(ClaimTypes.Name, acc.MaTruong),
            new Claim("TenTruong", acc.TenTruong),
            new Claim(ClaimTypes.Role, acc.Role)
        };
        var identity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
        await ctx.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, new ClaimsPrincipal(identity));

        string redirectUrl = acc.Role == "So" ? "/AdminSo" : "/AdminTruong";
        return Results.Ok(new { redirectUrl });
    }
    catch (Exception ex)
    {
        return Results.Problem("Lỗi hệ thống: " + ex.Message);
    }
});
// --- API SỞ: XUẤT EXCEL ---
app.MapGet("/api/so/export-excel", async (AppDbContext db, HttpContext ctx) =>
{
    if (!ctx.User.IsInRole("So")) return Results.Unauthorized();

    var list = await db.ThiSinhs.OrderBy(x => x.TenTruong).ThenBy(x => x.HoTen).ToListAsync();
    using var package = new ExcelPackage();
    var ws = package.Workbook.Worksheets.Add("DanhSachThiSinh");

    ws.Cells[1, 1].Value = "STT";
    ws.Cells[1, 2].Value = "Họ và Tên";
    ws.Cells[1, 3].Value = "Số CCCD/Định Danh";
    ws.Cells[1, 4].Value = "Trường";
    ws.Cells[1, 5].Value = "Môn Thi";
    ws.Cells[1, 6].Value = "Ngày Đăng Ký";

    for (int i = 0; i < list.Count; i++)
    {
        ws.Cells[i + 2, 1].Value = i + 1;
        ws.Cells[i + 2, 2].Value = list[i].HoTen;
        ws.Cells[i + 2, 3].Value = list[i].CCCD;
        ws.Cells[i + 2, 4].Value = list[i].TenTruong;
        ws.Cells[i + 2, 5].Value = list[i].MonThi;
        ws.Cells[i + 2, 6].Value = list[i].NgayDangKy.ToString("dd/MM/yyyy HH:mm");
    }

    var bytes = await package.GetAsByteArrayAsync();
    return Results.File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DanhSach_ThiSinh_ToanTinh.xlsx");
});

// --- API TRƯỜNG: UPLOAD EXCEL ---
app.MapPost("/api/truong/upload-excel", async (IFormFile file, AppDbContext db, HttpContext ctx) =>
{
    if (!ctx.User.IsInRole("Truong")) return Results.Unauthorized();
    var maTruong = ctx.User.FindFirstValue(ClaimTypes.Name);
    var truongInfo = await db.TaiKhoans.FirstOrDefaultAsync(x => x.MaTruong == maTruong);

    if (file == null || file.Length == 0) return Results.BadRequest(new { message = "Vui lòng chọn file Excel!" });

    using var stream = new MemoryStream();
    await file.CopyToAsync(stream);
    using var package = new ExcelPackage(stream);
    
    if (package.Workbook.Worksheets.Count == 0) 
        return Results.BadRequest(new { message = "File Excel không có dữ liệu!" });

    var ws = package.Workbook.Worksheets[0];
    int rowCount = ws.Dimension?.Rows ?? 0;
    int countSuccess = 0;

    for (int row = 2; row <= rowCount; row++)
    {
        var hoTen = ws.Cells[row, 1].Value?.ToString()?.Trim();
        var cccd = ws.Cells[row, 2].Value?.ToString()?.Trim();
        var monThi = ws.Cells[row, 3].Value?.ToString()?.Trim();

        if (string.IsNullOrEmpty(hoTen) || string.IsNullOrEmpty(cccd)) continue;

        var exists = await db.ThiSinhs.AnyAsync(x => x.CCCD == cccd);
        if (!exists)
        {
            db.ThiSinhs.Add(new ThiSinh
            {
                HoTen = hoTen,
                CCCD = cccd,
                MonThi = monThi ?? "Chưa chọn",
                TenTruong = truongInfo?.TenTruong ?? maTruong,
                MaTruong = maTruong,
                NgayDangKy = DateTime.UtcNow
            });
            countSuccess++;
        }
    }

    await db.SaveChangesAsync();
    return Results.Ok(new { message = $"Đã nhập thành công {countSuccess} học sinh!" });
}).DisableAntiforgery();

app.Run();

// --- DATA MODELS ---
public class AppDbContext : DbContext
{
    public AppDbContext(DbContextOptions<AppDbContext> options) : base(options) { }
    public DbSet<ThiSinh> ThiSinhs { get; set; }
    public DbSet<TaiKhoan> TaiKhoans { get; set; }
}

public class LoginRequest
{
    public string MaTruong { get; set; } = "";
    public string MatKhau { get; set; } = "";
}

public class TaiKhoan
{
    public int Id { get; set; }
    public string MaTruong { get; set; } = "";
    public string TenTruong { get; set; } = "";
    public string MatKhau { get; set; } = "";
    public string Role { get; set; } = "Truong";
}

public class ThiSinh
{
    public int Id { get; set; }
    public string HoTen { get; set; } = "";
    public string CCCD { get; set; } = "";
    public string TenTruong { get; set; } = "";
    public string MaTruong { get; set; } = "";
    public string MonThi { get; set; } = "";
    public DateTime NgayDangKy { get; set; }
}
