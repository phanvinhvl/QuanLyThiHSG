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
builder.Services.Configure<RouteOptions>(options => options.LowercaseUrls = true);
builder.Services.AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme)
    .AddCookie(options => {
        options.LoginPath = "/login";
        options.AccessDeniedPath = "/accessdenied";
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

// --- API AUTH ---
app.MapPost("/api/auth/login", async (AppDbContext db, HttpContext ctx, LoginRequest req) =>
{
    try
    {
        string sqlCreateTaiKhoan = @"CREATE TABLE IF NOT EXISTS ""TaiKhoans"" (""Id"" SERIAL PRIMARY KEY, ""MaTruong"" TEXT NOT NULL, ""TenTruong"" TEXT NOT NULL, ""MatKhau"" TEXT NOT NULL, ""Role"" TEXT NOT NULL);";
        await db.Database.ExecuteSqlRawAsync(sqlCreateTaiKhoan);

        string sqlCreateThiSinh = @"CREATE TABLE IF NOT EXISTS ""ThiSinhs"" (""Id"" SERIAL PRIMARY KEY, ""HoTen"" TEXT NOT NULL, ""CCCD"" TEXT NOT NULL, ""TenTruong"" TEXT NOT NULL, ""MaTruong"" TEXT NOT NULL, ""MonThi"" TEXT NOT NULL, ""NgayDangKy"" TIMESTAMP WITH TIME ZONE NOT NULL, ""SBD"" TEXT, ""PhongThi"" TEXT);";
        await db.Database.ExecuteSqlRawAsync(sqlCreateThiSinh);

        // Bổ sung cột SBD và PhongThi nếu chưa có
        try { await db.Database.ExecuteSqlRawAsync(@"ALTER TABLE ""ThiSinhs"" ADD COLUMN IF NOT EXISTS ""SBD"" TEXT;"); } catch {}
        try { await db.Database.ExecuteSqlRawAsync(@"ALTER TABLE ""ThiSinhs"" ADD COLUMN IF NOT EXISTS ""PhongThi"" TEXT;"); } catch {}

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

        string redirectUrl = acc.Role == "So" ? "/adminso" : "/admintruong";
        return Results.Ok(new { redirectUrl });
    }
    catch (Exception ex)
    {
        return Results.Problem("Lỗi hệ thống: " + ex.Message);
    }
});

// --- API 1: SỞ IMPORT DANH SÁCH TÀI KHOẢN TRƯỜNG TỪ EXCEL ---
app.MapPost("/api/so/import-truong", async (IFormFile file, AppDbContext db, HttpContext ctx) =>
{
    if (!ctx.User.IsInRole("So")) return Results.Unauthorized();
    if (file == null || file.Length == 0) return Results.BadRequest(new { message = "Vui lòng chọn file Excel!" });

    using var stream = new MemoryStream();
    await file.CopyToAsync(stream);
    using var package = new ExcelPackage(stream);
    if (package.Workbook.Worksheets.Count == 0) return Results.BadRequest(new { message = "File Excel không có dữ liệu!" });

    var ws = package.Workbook.Worksheets[0];
    int rowCount = ws.Dimension?.Rows ?? 0;
    int countSuccess = 0;

    for (int row = 2; row <= rowCount; row++)
    {
        var maTruong = ws.Cells[row, 1].Value?.ToString()?.Trim();
        var tenTruong = ws.Cells[row, 2].Value?.ToString()?.Trim();
        var matKhau = ws.Cells[row, 3].Value?.ToString()?.Trim();

        if (string.IsNullOrEmpty(maTruong) || string.IsNullOrEmpty(tenTruong) || string.IsNullOrEmpty(matKhau)) continue;

        var exists = await db.TaiKhoans.AnyAsync(x => x.MaTruong == maTruong);
        if (!exists)
        {
            db.TaiKhoans.Add(new TaiKhoan { MaTruong = maTruong, TenTruong = tenTruong, MatKhau = matKhau, Role = "Truong" });
            countSuccess++;
        }
    }

    await db.SaveChangesAsync();
    return Results.Ok(new { message = $"Đã tạo thành công {countSuccess} tài khoản trường!" });
}).DisableAntiforgery();

// --- API 2: TỰ ĐỘNG ĐÁNH SBD VÀ CHIA PHÒNG THI (SỞ) ---
app.MapPost("/api/so/chia-phong", async (AppDbContext db, HttpContext ctx, ChiaPhongRequest req) =>
{
    if (!ctx.User.IsInRole("So")) return Results.Unauthorized();

    int hsPerPhong = req.SoHocSinhMoiPhong > 0 ? req.SoHocSinhMoiPhong : 24;
    var monThiList = await db.ThiSinhs.Select(x => x.MonThi).Distinct().ToListAsync();

    foreach (var mon in monThiList)
    {
        var thiSinhs = await db.ThiSinhs.Where(x => x.MonThi == mon).OrderBy(x => x.HoTen).ToListAsync();
        int phongIdx = 1;
        int inPhongCount = 0;

        for (int i = 0; i < thiSinhs.Count; i++)
        {
            if (inPhongCount >= hsPerPhong)
            {
                phongIdx++;
                inPhongCount = 0;
            }

            inPhongCount++;
            thiSinhs[i].SBD = $"{mon.Substring(0, Math.Min(3, mon.Length)).ToUpper()}{(i + 1):D3}";
            thiSinhs[i].PhongThi = $"P.{mon}-{phongIdx:D2}";
        }
    }

    await db.SaveChangesAsync();
    return Results.Ok(new { message = "Đã đánh Số Báo Danh và Chia Phòng Thi thành công!" });
}).DisableAntiforgery();

// --- API 3: XUẤT FILE EXCEL TỔNG HỢP (SỞ) ---
app.MapGet("/api/so/export-excel", async (AppDbContext db, HttpContext ctx) =>
{
    if (!ctx.User.IsInRole("So")) return Results.Unauthorized();

    var list = await db.ThiSinhs.OrderBy(x => x.MonThi).ThenBy(x => x.SBD).ThenBy(x => x.HoTen).ToListAsync();
    using var package = new ExcelPackage();
    var ws = package.Workbook.Worksheets.Add("DanhSachThiSinh");

    ws.Cells[1, 1].Value = "STT";
    ws.Cells[1, 2].Value = "SBD";
    ws.Cells[1, 3].Value = "Phòng Thi";
    ws.Cells[1, 4].Value = "Họ và Tên";
    ws.Cells[1, 5].Value = "Số CCCD/Định Danh";
    ws.Cells[1, 6].Value = "Trường";
    ws.Cells[1, 7].Value = "Môn Thi";

    for (int i = 0; i < list.Count; i++)
    {
        ws.Cells[i + 2, 1].Value = i + 1;
        ws.Cells[i + 2, 2].Value = list[i].SBD ?? "Chưa chia";
        ws.Cells[i + 2, 3].Value = list[i].PhongThi ?? "Chưa chia";
        ws.Cells[i + 2, 4].Value = list[i].HoTen;
        ws.Cells[i + 2, 5].Value = list[i].CCCD;
        ws.Cells[i + 2, 6].Value = list[i].TenTruong;
        ws.Cells[i + 2, 7].Value = list[i].MonThi;
    }

    var bytes = await package.GetAsByteArrayAsync();
    return Results.File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DanhSach_ThiSinh_SBD_PhongThi.xlsx");
});

// --- API 4: TẢI FILE EXCEL MẪU DÀNH CHO TRƯỜNG ---
app.MapGet("/api/truong/download-mau", async () =>
{
    using var package = new ExcelPackage();
    var ws = package.Workbook.Worksheets.Add("MauDangKy");

    ws.Cells[1, 1].Value = "Họ và Tên";
    ws.Cells[1, 2].Value = "Số CCCD/Định Danh";
    ws.Cells[1, 3].Value = "Môn Dự Thi";

    ws.Cells[2, 1].Value = "Nguyễn Văn A";
    ws.Cells[2, 2].Value = "038200123456";
    ws.Cells[2, 3].Value = "Toán";

    ws.Cells[3, 1].Value = "Trần Thị B";
    ws.Cells[3, 2].Value = "038200654321";
    ws.Cells[3, 3].Value = "Ngữ Văn";

    var bytes = await package.GetAsByteArrayAsync();
    return Results.File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Mau_Dang_Ky_HSG.xlsx");
});

// --- API 5: TRƯỜNG UPLOAD EXCEL ---
app.MapPost("/api/truong/upload-excel", async (IFormFile file, AppDbContext db, HttpContext ctx) =>
{
    if (!ctx.User.IsInRole("Truong")) return Results.Unauthorized();
    var maTruong = ctx.User.FindFirstValue(ClaimTypes.Name);
    var truongInfo = await db.TaiKhoans.FirstOrDefaultAsync(x => x.MaTruong == maTruong);

    if (file == null || file.Length == 0) return Results.BadRequest(new { message = "Vui lòng chọn file Excel!" });

    using var stream = new MemoryStream();
    await file.CopyToAsync(stream);
    using var package = new ExcelPackage(stream);
    if (package.Workbook.Worksheets.Count == 0) return Results.BadRequest(new { message = "File Excel không có dữ liệu!" });

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

// --- API 6: SỬA & XÓA THÍ SINH (CẤP TRƯỜNG) ---
app.MapPut("/api/truong/sua-thisinh/{id}", async (int id, AppDbContext db, HttpContext ctx, UpdateThiSinhReq req) =>
{
    if (!ctx.User.IsInRole("Truong")) return Results.Unauthorized();
    var maTruong = ctx.User.FindFirstValue(ClaimTypes.Name);

    var ts = await db.ThiSinhs.FirstOrDefaultAsync(x => x.Id == id && x.MaTruong == maTruong);
    if (ts == null) return Results.NotFound(new { message = "Không tìm thấy học sinh!" });

    ts.HoTen = req.HoTen;
    ts.CCCD = req.CCCD;
    ts.MonThi = req.MonThi;
    await db.SaveChangesAsync();

    return Results.Ok(new { message = "Cập nhật thành công!" });
}).DisableAntiforgery();

app.MapDelete("/api/truong/xoa-thisinh/{id}", async (int id, AppDbContext db, HttpContext ctx) =>
{
    if (!ctx.User.IsInRole("Truong")) return Results.Unauthorized();
    var maTruong = ctx.User.FindFirstValue(ClaimTypes.Name);

    var ts = await db.ThiSinhs.FirstOrDefaultAsync(x => x.Id == id && x.MaTruong == maTruong);
    if (ts == null) return Results.NotFound(new { message = "Không tìm thấy học sinh!" });

    db.ThiSinhs.Remove(ts);
    await db.SaveChangesAsync();

    return Results.Ok(new { message = "Đã xóa học sinh!" });
}).DisableAntiforgery();
// --- API XUẤT FILE EXCEL DANH SÁCH PHÒNG THI (MỖI PHÒNG 1 SHEET) ---
app.MapGet("/api/so/export-phong-thi", async (AppDbContext db, HttpContext ctx) =>
{
    var listPhong = await db.ThiSinhs
        .Where(x => !string.IsNullOrEmpty(x.PhongThi))
        .Select(x => x.PhongThi)
        .Distinct()
        .OrderBy(x => x)
        .ToListAsync();

    if (listPhong.Count == 0) 
        return Results.BadRequest(new { message = "Chưa có dữ liệu phòng thi! Vui lòng thực hiện Chia Phòng Thi trước." });

    using var package = new ExcelPackage();

    foreach (var phong in listPhong)
    {
        // Tên Sheet khống chế tối đa 31 ký tự theo chuẩn Excel
        string sheetName = phong.Replace("/", "-").Replace(":", "-");
        if (sheetName.Length > 30) sheetName = sheetName.Substring(0, 30);

        var ws = package.Workbook.Worksheets.Add(sheetName);

        // Tiêu đề Trang
        ws.Cells[1, 1, 1, 7].Merge = true;
        ws.Cells[1, 1].Value = $"DANH SÁCH THÍ SINH DỰ THI - PHÒNG THI: {phong}";
        ws.Cells[1, 1].Style.Font.Bold = true;
        ws.Cells[1, 1].Style.Font.Size = 14;
        ws.Cells[1, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

        // Tiêu đề Bảng
        ws.Cells[3, 1].Value = "STT";
        ws.Cells[3, 2].Value = "SBD";
        ws.Cells[3, 3].Value = "Họ và Tên";
        ws.Cells[3, 4].Value = "Số CCCD/Định danh";
        ws.Cells[3, 5].Value = "Trường THPT/THCS";
        ws.Cells[3, 6].Value = "Môn Thi";
        ws.Cells[3, 7].Value = "Chữ ký thí sinh";

        using (var range = ws.Cells[3, 1, 3, 7])
        {
            range.Style.Font.Bold = true;
            range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
        }

        var thiSinhs = await db.ThiSinhs
            .Where(x => x.PhongThi == phong)
            .OrderBy(x => x.SBD)
            .ToListAsync();

        for (int i = 0; i < thiSinhs.Count; i++)
        {
            int row = i + 4;
            ws.Cells[row, 1].Value = i + 1;
            ws.Cells[row, 2].Value = thiSinhs[i].SBD;
            ws.Cells[row, 3].Value = thiSinhs[i].HoTen;
            ws.Cells[row, 4].Value = thiSinhs[i].CCCD;
            ws.Cells[row, 5].Value = thiSinhs[i].TenTruong;
            ws.Cells[row, 6].Value = thiSinhs[i].MonThi;
            ws.Cells[row, 7].Value = ""; // Ô ký tên
        }

        ws.Cells[ws.Dimension.Address].AutoFitColumns();
    }

    var bytes = await package.GetAsByteArrayAsync();
    return Results.File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DanhSach_TheoPhongThi.xlsx");
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

public class ChiaPhongRequest
{
    public int SoHocSinhMoiPhong { get; set; } = 24;
}

public class UpdateThiSinhReq
{
    public string HoTen { get; set; } = "";
    public string CCCD { get; set; } = "";
    public string MonThi { get; set; } = "";
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
    public string? SBD { get; set; }
    public string? PhongThi { get; set; }
    public DateTime NgayDangKy { get; set; }
}
