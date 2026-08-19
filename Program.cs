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
        await db.Database.ExecuteSqlRawAsync(@"CREATE TABLE IF NOT EXISTS ""TaiKhoans"" (""Id"" SERIAL PRIMARY KEY, ""MaTruong"" TEXT NOT NULL, ""TenTruong"" TEXT NOT NULL, ""MatKhau"" TEXT NOT NULL, ""Role"" TEXT NOT NULL);");
        await db.Database.ExecuteSqlRawAsync(@"CREATE TABLE IF NOT EXISTS ""ThiSinhs"" (""Id"" SERIAL PRIMARY KEY, ""HoTen"" TEXT NOT NULL, ""CCCD"" TEXT NOT NULL, ""TenTruong"" TEXT NOT NULL, ""MaTruong"" TEXT NOT NULL, ""MonThi"" TEXT NOT NULL, ""NgayDangKy"" TIMESTAMP WITH TIME ZONE NOT NULL, ""SBD"" TEXT, ""PhongThi"" TEXT);");
        await db.Database.ExecuteSqlRawAsync(@"CREATE TABLE IF NOT EXISTS ""DanhMucMonThis"" (""Id"" SERIAL PRIMARY KEY, ""TenMon"" TEXT NOT NULL);");

        if (!await db.TaiKhoans.AnyAsync(x => x.Role == "So"))
        {
            db.TaiKhoans.Add(new TaiKhoan { MaTruong = "SO_GD", TenTruong = "Sở Giáo Dục & Đào Tạo", MatKhau = "so@123456", Role = "So" });
            await db.SaveChangesAsync();
        }

        // Tạo danh mục môn thi mặc định nếu chưa có
        if (!await db.DanhMucMonThis.AnyAsync())
        {
            var defaultMons = new[] { "Toán", "Ngữ Văn", "Tiếng Anh", "Vật Lý", "Hóa Học", "Sinh Học", "Tin Học", "Lịch Sử", "Địa Lý" };
            foreach (var m in defaultMons) db.DanhMucMonThis.Add(new DanhMucMonThi { TenMon = m });
            await db.SaveChangesAsync();
        }

        var acc = await db.TaiKhoans.FirstOrDefaultAsync(x => x.MaTruong == req.MaTruong && x.MatKhau == req.MatKhau);
        if (acc == null) return Results.BadRequest(new { message = "Mã tài khoản hoặc mật khẩu không chính xác!" });

        var claims = new List<Claim> { new Claim(ClaimTypes.Name, acc.MaTruong), new Claim("TenTruong", acc.TenTruong), new Claim(ClaimTypes.Role, acc.Role) };
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

// --- API QUẢN LÝ DANH MỤC MÔN THI (SỞ) ---
app.MapPost("/api/so/them-mon", async (AppDbContext db, HttpContext ctx, DanhMucMonThi req) =>
{
    if (!ctx.User.IsInRole("So")) return Results.Unauthorized();
    if (string.IsNullOrWhiteSpace(req.TenMon)) return Results.BadRequest(new { message = "Tên môn không được để trống!" });

    var tenMonClean = req.TenMon.Trim();
    if (await db.DanhMucMonThis.AnyAsync(x => x.TenMon.ToLower() == tenMonClean.ToLower()))
        return Results.BadRequest(new { message = "Môn thi này đã tồn tại!" });

    db.DanhMucMonThis.Add(new DanhMucMonThi { TenMon = tenMonClean });
    await db.SaveChangesAsync();
    return Results.Ok(new { message = "Đã thêm môn thi mới!" });
}).DisableAntiforgery();

// --- API SỬA TÊN MÔN THI (SỞ) ---
app.MapPut("/api/so/sua-mon/{id}", async (int id, AppDbContext db, HttpContext ctx, DanhMucMonThi req) =>
{
    if (!ctx.User.IsInRole("So")) return Results.Unauthorized();
    if (string.IsNullOrWhiteSpace(req.TenMon)) return Results.BadRequest(new { message = "Tên môn không được để trống!" });

    var item = await db.DanhMucMonThis.FindAsync(id);
    if (item == null) return Results.NotFound(new { message = "Không tìm thấy môn thi!" });

    var tenMonClean = req.TenMon.Trim();
    
    // Kiểm tra trùng tên với môn khác
    if (await db.DanhMucMonThis.AnyAsync(x => x.Id != id && x.TenMon.ToLower() == tenMonClean.ToLower()))
        return Results.BadRequest(new { message = "Tên môn thi này đã tồn tại!" });

    string tenMonCu = item.TenMon;
    item.TenMon = tenMonClean;

    // Tự động cập nhật lại tên môn mới cho toàn bộ thí sinh đã đăng ký môn cũ
    var listThiSinh = await db.ThiSinhs.Where(x => x.MonThi == tenMonCu).ToListAsync();
    foreach (var ts in listThiSinh)
    {
        ts.MonThi = tenMonClean;
    }

    await db.SaveChangesAsync();
    return Results.Ok(new { message = $"Đã đổi tên môn từ '{tenMonCu}' thành '{tenMonClean}'!" });
}).DisableAntiforgery();

app.MapDelete("/api/so/xoa-mon/{id}", async (int id, AppDbContext db, HttpContext ctx) =>
{
    if (!ctx.User.IsInRole("So")) return Results.Unauthorized();
    var item = await db.DanhMucMonThis.FindAsync(id);
    if (item == null) return Results.NotFound();

    db.DanhMucMonThis.Remove(item);
    await db.SaveChangesAsync();
    return Results.Ok(new { message = "Đã xóa môn thi!" });
}).DisableAntiforgery();

// --- API IMPORT EXCEL TRƯỜNG (CÓ CHECK MÔN THI SỞ CẤP) ---
app.MapPost("/api/truong/upload-excel", async (IFormFile file, AppDbContext db, HttpContext ctx) =>
{
    if (!ctx.User.IsInRole("Truong")) return Results.Unauthorized();
    var maTruong = ctx.User.FindFirstValue(ClaimTypes.Name);
    var truongInfo = await db.TaiKhoans.FirstOrDefaultAsync(x => x.MaTruong == maTruong);

    if (file == null || file.Length == 0) return Results.BadRequest(new { message = "Vui lòng chọn file Excel!" });

    var dsMonValid = await db.DanhMucMonThis.Select(x => x.TenMon.ToLower()).ToListAsync();

    using var stream = new MemoryStream();
    await file.CopyToAsync(stream);
    using var package = new ExcelPackage(stream);
    var ws = package.Workbook.Worksheets[0];
    int rowCount = ws.Dimension?.Rows ?? 0;

    int countSuccess = 0;
    List<string> loiMon = new List<string>();

    for (int row = 2; row <= rowCount; row++)
    {
        var hoTen = ws.Cells[row, 1].Value?.ToString()?.Trim();
        var cccd = ws.Cells[row, 2].Value?.ToString()?.Trim();
        var monThi = ws.Cells[row, 3].Value?.ToString()?.Trim();

        if (string.IsNullOrEmpty(hoTen) || string.IsNullOrEmpty(cccd)) continue;

        // KIỂM TRA MÔN THI CÓ TRONG DANH MỤC KHÔNG
        if (string.IsNullOrEmpty(monThi) || !dsMonValid.Contains(monThi.ToLower()))
        {
            loiMon.Add($"Dòng {row}: Môn '{monThi}' không nằm trong danh mục Sở quy định!");
            continue;
        }

        var exists = await db.ThiSinhs.AnyAsync(x => x.CCCD == cccd);
        if (!exists)
        {
            // Lấy đúng tên môn thi chuẩn danh mục
            var monChuan = await db.DanhMucMonThis.FirstAsync(x => x.TenMon.ToLower() == monThi.ToLower());
            db.ThiSinhs.Add(new ThiSinh
            {
                HoTen = hoTen,
                CCCD = cccd,
                MonThi = monChuan.TenMon,
                TenTruong = truongInfo?.TenTruong ?? maTruong,
                MaTruong = maTruong,
                NgayDangKy = DateTime.UtcNow
            });
            countSuccess++;
        }
    }

    await db.SaveChangesAsync();

    string msg = $"Đã nhập thành công {countSuccess} học sinh!";
    if (loiMon.Count > 0) msg += $"\nCó {loiMon.Count} dòng bị bỏ qua do sai môn thi.";

    return Results.Ok(new { message = msg, loiMon });
}).DisableAntiforgery();

// --- API SỬA THÍ SINH (CHECK MÔN THI SỞ CẤP) ---
app.MapPut("/api/truong/sua-thisinh/{id}", async (int id, AppDbContext db, HttpContext ctx, UpdateThiSinhReq req) =>
{
    if (!ctx.User.IsInRole("Truong")) return Results.Unauthorized();
    var maTruong = ctx.User.FindFirstValue(ClaimTypes.Name);

    var ts = await db.ThiSinhs.FirstOrDefaultAsync(x => x.Id == id && x.MaTruong == maTruong);
    if (ts == null) return Results.NotFound(new { message = "Không tìm thấy học sinh!" });

    var monChuan = await db.DanhMucMonThis.FirstOrDefaultAsync(x => x.TenMon.ToLower() == req.MonThi.ToLower());
    if (monChuan == null)
        return Results.BadRequest(new { message = $"Môn '{req.MonThi}' không hợp lệ! Vui lòng chọn môn trong danh mục." });

    ts.HoTen = req.HoTen?.Trim() ?? ts.HoTen;
    ts.CCCD = req.CCCD?.Trim() ?? ts.CCCD;
    ts.MonThi = monChuan.TenMon;
    if (!string.IsNullOrEmpty(req.TenTruong)) ts.TenTruong = req.TenTruong.Trim();

    await db.SaveChangesAsync();
    return Results.Ok(new { message = "Cập nhật thành công!" });
}).DisableAntiforgery();

// Các API khác giữ nguyên (so/import-truong, so/chia-phong, so/export-excel, truong/xoa-thisinh...)
app.MapPost("/api/so/import-truong", async (IFormFile file, AppDbContext db, HttpContext ctx) =>
{
    if (!ctx.User.IsInRole("So")) return Results.Unauthorized();
    if (file == null || file.Length == 0) return Results.BadRequest(new { message = "Vui lòng chọn file Excel!" });

    using var stream = new MemoryStream();
    await file.CopyToAsync(stream);
    using var package = new ExcelPackage(stream);
    var ws = package.Workbook.Worksheets[0];
    int rowCount = ws.Dimension?.Rows ?? 0;
    int countSuccess = 0;

    for (int row = 2; row <= rowCount; row++)
    {
        var maTruong = ws.Cells[row, 1].Value?.ToString()?.Trim();
        var tenTruong = ws.Cells[row, 2].Value?.ToString()?.Trim();
        var matKhau = ws.Cells[row, 3].Value?.ToString()?.Trim();
        if (string.IsNullOrEmpty(maTruong) || string.IsNullOrEmpty(tenTruong) || string.IsNullOrEmpty(matKhau)) continue;

        if (!await db.TaiKhoans.AnyAsync(x => x.MaTruong == maTruong))
        {
            db.TaiKhoans.Add(new TaiKhoan { MaTruong = maTruong, TenTruong = tenTruong, MatKhau = matKhau, Role = "Truong" });
            countSuccess++;
        }
    }
    await db.SaveChangesAsync();
    return Results.Ok(new { message = $"Đã tạo thành công {countSuccess} tài khoản trường!" });
}).DisableAntiforgery();

app.MapPost("/api/so/chia-phong", async (AppDbContext db, HttpContext ctx, ChiaPhongRequest req) =>
{
    if (!ctx.User.IsInRole("So")) return Results.Unauthorized();
    int hsPerPhong = req.SoHocSinhMoiPhong > 0 ? req.SoHocSinhMoiPhong : 24;
    var monThiList = await db.ThiSinhs.Select(x => x.MonThi).Distinct().ToListAsync();

    foreach (var mon in monThiList)
    {
        var thiSinhs = await db.ThiSinhs.Where(x => x.MonThi == mon).OrderBy(x => x.HoTen).ToListAsync();
        int phongIdx = 1, inPhongCount = 0;

        for (int i = 0; i < thiSinhs.Count; i++)
        {
            if (inPhongCount >= hsPerPhong) { phongIdx++; inPhongCount = 0; }
            inPhongCount++;
            thiSinhs[i].SBD = $"{mon.Substring(0, Math.Min(3, mon.Length)).ToUpper()}{(i + 1):D3}";
            thiSinhs[i].PhongThi = $"P.{mon}-{phongIdx:D2}";
        }
    }
    await db.SaveChangesAsync();
    return Results.Ok(new { message = "Đã đánh SBD và Chia Phòng Thi thành công!" });
}).DisableAntiforgery();

app.MapGet("/api/so/export-excel", async (AppDbContext db, HttpContext ctx) =>
{
    if (!ctx.User.IsInRole("So")) return Results.Unauthorized();
    var list = await db.ThiSinhs.OrderBy(x => x.MonThi).ThenBy(x => x.SBD).ThenBy(x => x.HoTen).ToListAsync();
    using var package = new ExcelPackage();
    var ws = package.Workbook.Worksheets.Add("DanhSachThiSinh");
    ws.Cells[1, 1].Value = "STT"; ws.Cells[1, 2].Value = "SBD"; ws.Cells[1, 3].Value = "Phòng Thi"; ws.Cells[1, 4].Value = "Họ và Tên"; ws.Cells[1, 5].Value = "Số CCCD"; ws.Cells[1, 6].Value = "Trường"; ws.Cells[1, 7].Value = "Môn Thi";

    for (int i = 0; i < list.Count; i++)
    {
        ws.Cells[i + 2, 1].Value = i + 1; ws.Cells[i + 2, 2].Value = list[i].SBD ?? "Chưa chia"; ws.Cells[i + 2, 3].Value = list[i].PhongThi ?? "Chưa chia"; ws.Cells[i + 2, 4].Value = list[i].HoTen; ws.Cells[i + 2, 5].Value = list[i].CCCD; ws.Cells[i + 2, 6].Value = list[i].TenTruong; ws.Cells[i + 2, 7].Value = list[i].MonThi;
    }
    var bytes = await package.GetAsByteArrayAsync();
    return Results.File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DanhSach_ThiSinh_SBD_PhongThi.xlsx");
});

app.MapGet("/api/truong/download-mau", async (AppDbContext db) =>
{
    using var package = new ExcelPackage();
    var ws = package.Workbook.Worksheets.Add("MauDangKy");
    ws.Cells[1, 1].Value = "Họ và Tên"; ws.Cells[1, 2].Value = "Số CCCD/Định Danh"; ws.Cells[1, 3].Value = "Môn Dự Thi";
    
    var mons = await db.DanhMucMonThis.Select(x => x.TenMon).ToListAsync();
    ws.Cells[2, 1].Value = "Nguyễn Văn A"; ws.Cells[2, 2].Value = "038200123456"; ws.Cells[2, 3].Value = mons.FirstOrDefault() ?? "Toán";

    // Tạo thêm Sheet Hướng dẫn danh mục môn thi
    var wsGuide = package.Workbook.Worksheets.Add("DanhMucMonThiChuan");
    wsGuide.Cells[1, 1].Value = "STT"; wsGuide.Cells[1, 2].Value = "Danh Mục Môn Thi Cho Phép";
    for (int i = 0; i < mons.Count; i++)
    {
        wsGuide.Cells[i + 2, 1].Value = i + 1;
        wsGuide.Cells[i + 2, 2].Value = mons[i];
    }

    var bytes = await package.GetAsByteArrayAsync();
    return Results.File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Mau_Dang_Ky_HSG.xlsx");
});

app.MapDelete("/api/truong/xoa-thisinh/{id}", async (int id, AppDbContext db, HttpContext ctx) =>
{
    if (!ctx.User.IsInRole("Truong")) return Results.Unauthorized();
    var maTruong = ctx.User.FindFirstValue(ClaimTypes.Name);
    var ts = await db.ThiSinhs.FirstOrDefaultAsync(x => x.Id == id && x.MaTruong == maTruong);
    if (ts == null) return Results.NotFound();
    db.ThiSinhs.Remove(ts);
    await db.SaveChangesAsync();
    return Results.Ok(new { message = "Đã xóa học sinh!" });
}).DisableAntiforgery();

app.Run();

public class AppDbContext : DbContext
{
    public AppDbContext(DbContextOptions<AppDbContext> options) : base(options) { }
    public DbSet<ThiSinh> ThiSinhs { get; set; }
    public DbSet<TaiKhoan> TaiKhoans { get; set; }
    public DbSet<DanhMucMonThi> DanhMucMonThis { get; set; }
}

public class DanhMucMonThi
{
    public int Id { get; set; }
    public string TenMon { get; set; } = "";
}

public class LoginRequest { public string MaTruong { get; set; } = ""; public string MatKhau { get; set; } = ""; }
public class ChiaPhongRequest { public int SoHocSinhMoiPhong { get; set; } = 24; }
public class UpdateThiSinhReq { public string HoTen { get; set; } = ""; public string CCCD { get; set; } = ""; public string MonThi { get; set; } = ""; public string TenTruong { get; set; } = ""; }
public class TaiKhoan { public int Id { get; set; } public string MaTruong { get; set; } = ""; public string TenTruong { get; set; } = ""; public string MatKhau { get; set; } = ""; public string Role { get; set; } = "Truong"; }
public class ThiSinh { public int Id { get; set; } public string HoTen { get; set; } = ""; public string CCCD { get; set; } = ""; public string TenTruong { get; set; } = ""; public string MaTruong { get; set; } = ""; public string MonThi { get; set; } = ""; public string? SBD { get; set; } public string? PhongThi { get; set; } public DateTime NgayDangKy { get; set; } }
