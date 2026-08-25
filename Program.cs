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

app.MapGet("/", () => Results.Redirect("/login"));

// --- API AUTH ---
app.MapPost("/api/auth/login", async (AppDbContext db, HttpContext ctx, LoginRequest req) =>
{
    try
    {
        await db.Database.ExecuteSqlRawAsync(@"CREATE TABLE IF NOT EXISTS ""TaiKhoans"" (""Id"" SERIAL PRIMARY KEY, ""MaTruong"" TEXT NOT NULL, ""TenTruong"" TEXT NOT NULL, ""MatKhau"" TEXT NOT NULL, ""Role"" TEXT NOT NULL);");
        await db.Database.ExecuteSqlRawAsync(@"CREATE TABLE IF NOT EXISTS ""ThiSinhs"" (""Id"" SERIAL PRIMARY KEY, ""HoTen"" TEXT NOT NULL, ""CCCD"" TEXT NOT NULL, ""TenTruong"" TEXT NOT NULL, ""MaTruong"" TEXT NOT NULL, ""MonThi"" TEXT NOT NULL, ""NgayDangKy"" TIMESTAMP WITH TIME ZONE NOT NULL, ""SBD"" TEXT, ""PhongThi"" TEXT, ""GiamThi1"" TEXT, ""GiamThi2"" TEXT);");
        await db.Database.ExecuteSqlRawAsync(@"CREATE TABLE IF NOT EXISTS ""DanhMucMonThis"" (""Id"" SERIAL PRIMARY KEY, ""MaMon"" TEXT, ""TenMon"" TEXT NOT NULL);");
        await db.Database.ExecuteSqlRawAsync(@"CREATE TABLE IF NOT EXISTS ""CanBoCoiThis"" (""Id"" SERIAL PRIMARY KEY, ""HoTen"" TEXT NOT NULL, ""DonViCongTac"" TEXT NOT NULL, ""ChucVu"" TEXT, ""Sdt"" TEXT);");

        try { await db.Database.ExecuteSqlRawAsync(@"ALTER TABLE ""DanhMucMonThis"" ADD COLUMN IF NOT EXISTS ""MaMon"" TEXT;"); } catch {}
        try { await db.Database.ExecuteSqlRawAsync(@"ALTER TABLE ""ThiSinhs"" ADD COLUMN IF NOT EXISTS ""GiamThi1"" TEXT;"); } catch {}
        try { await db.Database.ExecuteSqlRawAsync(@"ALTER TABLE ""ThiSinhs"" ADD COLUMN IF NOT EXISTS ""GiamThi2"" TEXT;"); } catch {}

        if (!await db.TaiKhoans.AnyAsync(x => x.Role == "So"))
        {
            db.TaiKhoans.Add(new TaiKhoan { MaTruong = "SO_GD", TenTruong = "Sở Giáo Dục & Đào Tạo", MatKhau = "so@123456", Role = "So" });
            await db.SaveChangesAsync();
        }

        if (!await db.DanhMucMonThis.AnyAsync())
        {
            var defaultMons = new[] { "Toán", "Ngữ Văn", "Tiếng Anh", "Vật Lý", "Hóa Học", "Sinh Học", "Tin Học", "Lịch Sử", "Địa Lý" };
            for (int i = 0; i < defaultMons.Length; i++)
                db.DanhMucMonThis.Add(new DanhMucMonThi { MaMon = (i + 1).ToString("D2"), TenMon = defaultMons[i] });
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

// --- API QUẢN LÝ CÁN BỘ COI THI (SỞ) ---
app.MapPost("/api/so/them-can-bo", async (AppDbContext db, HttpContext ctx, CanBoCoiThi req) =>
{
    if (!ctx.User.IsInRole("So")) return Results.Unauthorized();
    if (string.IsNullOrWhiteSpace(req.HoTen)) return Results.BadRequest(new { message = "Họ tên không được trống!" });

    db.CanBoCoiThis.Add(req);
    await db.SaveChangesAsync();
    return Results.Ok(new { message = "Thêm cán bộ coi thi thành công!" });
}).DisableAntiforgery();

app.MapPost("/api/so/phan-cong-giam-thi", async (AppDbContext db, HttpContext ctx) =>
{
    if (!ctx.User.IsInRole("So")) return Results.Unauthorized();

    var dsPhong = await db.ThiSinhs.Where(x => !string.IsNullOrEmpty(x.PhongThi))
                                   .Select(x => x.PhongThi)
                                   .Distinct().ToListAsync();

    var dsCanBo = await db.CanBoCoiThis.ToListAsync();

    if (dsCanBo.Count < dsPhong.Count * 2)
        return Results.BadRequest(new { message = $"Cần tối thiểu {dsPhong.Count * 2} cán bộ coi thi (hiện có {dsCanBo.Count}). Vui lòng bổ sung thêm!" });

    int cbIdx = 0;
    foreach (var phong in dsPhong)
    {
        var gt1 = dsCanBo[cbIdx++].HoTen;
        var gt2 = dsCanBo[cbIdx++].HoTen;

        var thiSinhsInPhong = await db.ThiSinhs.Where(x => x.PhongThi == phong).ToListAsync();
        foreach (var ts in thiSinhsInPhong)
        {
            ts.GiamThi1 = gt1;
            ts.GiamThi2 = gt2;
        }
    }

    await db.SaveChangesAsync();
    return Results.Ok(new { message = $"Đã phân công thành công Giám thị cho {dsPhong.Count} phòng thi!" });
}).DisableAntiforgery();

// --- API SỞ / TRƯỜNG KHÁC GIỮ NGUYÊN ---
app.MapPost("/api/so/them-mon", async (AppDbContext db, HttpContext ctx, DanhMucMonThi req) =>
{
    if (!ctx.User.IsInRole("So")) return Results.Unauthorized();
    var tenMonClean = req.TenMon.Trim();
    int count = await db.DanhMucMonThis.CountAsync();
    db.DanhMucMonThis.Add(new DanhMucMonThi { MaMon = (count + 1).ToString("D2"), TenMon = tenMonClean });
    await db.SaveChangesAsync();
    return Results.Ok(new { message = "Thêm môn thi thành công!" });
}).DisableAntiforgery();

app.MapPut("/api/so/sua-mon/{id}", async (int id, AppDbContext db, HttpContext ctx, DanhMucMonThi req) =>
{
    if (!ctx.User.IsInRole("So")) return Results.Unauthorized();
    var item = await db.DanhMucMonThis.FindAsync(id);
    if (item == null) return Results.NotFound();
    item.TenMon = req.TenMon.Trim();
    await db.SaveChangesAsync();
    return Results.Ok(new { message = "Cập nhật tên môn thành công!" });
}).DisableAntiforgery();

app.MapDelete("/api/so/xoa-mon/{id}", async (int id, AppDbContext db, HttpContext ctx) =>
{
    if (!ctx.User.IsInRole("So")) return Results.Unauthorized();
    var item = await db.DanhMucMonThis.FindAsync(id);
    if (item != null) { db.DanhMucMonThis.Remove(item); await db.SaveChangesAsync(); }
    return Results.Ok(new { message = "Đã xóa môn thi!" });
}).DisableAntiforgery();

app.MapPost("/api/so/import-truong", async (IFormFile file, AppDbContext db, HttpContext ctx) =>
{
    if (!ctx.User.IsInRole("So")) return Results.Unauthorized();
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
        if (string.IsNullOrEmpty(maTruong) || string.IsNullOrEmpty(tenTruong)) continue;
        if (!await db.TaiKhoans.AnyAsync(x => x.MaTruong == maTruong))
        {
            db.TaiKhoans.Add(new TaiKhoan { MaTruong = maTruong, TenTruong = tenTruong, MatKhau = matKhau ?? "123456", Role = "Truong" });
            countSuccess++;
        }
    }
    await db.SaveChangesAsync();
    return Results.Ok(new { message = $"Đã tạo {countSuccess} tài khoản trường!" });
}).DisableAntiforgery();

app.MapPost("/api/so/chia-phong", async (AppDbContext db, HttpContext ctx, ChiaPhongRequest req) =>
{
    if (!ctx.User.IsInRole("So")) return Results.Unauthorized();
    int hsPerPhong = req.SoHocSinhMoiPhong > 0 ? req.SoHocSinhMoiPhong : 24;
    var monThiList = await db.DanhMucMonThis.OrderBy(x => x.MaMon).Select(x => x.TenMon).ToListAsync();

    int globalSBDIndex = 1, globalPhongIndex = 1;
    foreach (var mon in monThiList)
    {
        var thiSinhs = await db.ThiSinhs.Where(x => x.MonThi == mon).OrderBy(x => x.MaTruong).ThenBy(x => x.HoTen).ToListAsync();
        if (thiSinhs.Count == 0) continue;
        int inPhongCount = 0;
        for (int i = 0; i < thiSinhs.Count; i++)
        {
            if (inPhongCount >= hsPerPhong) { globalPhongIndex++; inPhongCount = 0; }
            inPhongCount++;
            string maTruongPrefix = thiSinhs[i].MaTruong.PadLeft(2, '0');
            if (maTruongPrefix.Length > 2) maTruongPrefix = maTruongPrefix.Substring(0, 2);
            thiSinhs[i].SBD = $"{maTruongPrefix}{globalSBDIndex:D4}";
            thiSinhs[i].PhongThi = globalPhongIndex.ToString("D3");
            globalSBDIndex++;
        }
        globalPhongIndex++;
    }
    await db.SaveChangesAsync();
    return Results.Ok(new { message = $"Đã chia thành công {globalPhongIndex - 1} phòng thi và đánh SBD!" });
}).DisableAntiforgery();

app.MapGet("/api/so/export-excel", async (AppDbContext db, HttpContext ctx) =>
{
    if (!ctx.User.IsInRole("So")) return Results.Unauthorized();
    var list = await db.ThiSinhs.OrderBy(x => x.MonThi).ThenBy(x => x.SBD).ToListAsync();
    using var package = new ExcelPackage();
    var ws = package.Workbook.Worksheets.Add("DanhSachThiSinh");
    ws.Cells[1, 1].Value = "STT"; ws.Cells[1, 2].Value = "SBD"; ws.Cells[1, 3].Value = "Phòng Thi"; ws.Cells[1, 4].Value = "Họ và Tên"; ws.Cells[1, 5].Value = "Số CCCD"; ws.Cells[1, 6].Value = "Trường"; ws.Cells[1, 7].Value = "Môn Thi";
    for (int i = 0; i < list.Count; i++)
    {
        ws.Cells[i + 2, 1].Value = i + 1; ws.Cells[i + 2, 2].Value = list[i].SBD; ws.Cells[i + 2, 3].Value = list[i].PhongThi; ws.Cells[i + 2, 4].Value = list[i].HoTen; ws.Cells[i + 2, 5].Value = list[i].CCCD; ws.Cells[i + 2, 6].Value = list[i].TenTruong; ws.Cells[i + 2, 7].Value = list[i].MonThi;
    }
    var bytes = await package.GetAsByteArrayAsync();
    return Results.File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DanhSach_ThiSinh_TongHop.xlsx");
});

app.MapGet("/api/so/export-phong-thi", async (AppDbContext db, HttpContext ctx) =>
{
    var listPhong = await db.ThiSinhs.Where(x => !string.IsNullOrEmpty(x.PhongThi)).Select(x => x.PhongThi).Distinct().OrderBy(x => x).ToListAsync();
    if (listPhong.Count == 0) return Results.BadRequest(new { message = "Chưa có dữ liệu phòng thi!" });

    using var package = new ExcelPackage();
    foreach (var phong in listPhong)
    {
        if (string.IsNullOrEmpty(phong)) continue;
        string sheetName = phong.Replace("/", "-");
        if (sheetName.Length > 30) sheetName = sheetName.Substring(0, 30);

        var ws = package.Workbook.Worksheets.Add(sheetName);
        var sampleTS = await db.ThiSinhs.FirstOrDefaultAsync(x => x.PhongThi == phong);
        string monThi = sampleTS?.MonThi ?? "";

        ws.Cells[1, 1, 1, 3].Merge = true; ws.Cells[1, 1].Value = "SỞ GIÁO DỤC VÀ ĐÀO TẠO VĨNH LONG"; ws.Cells[1, 1].Style.Font.Bold = true;
        ws.Cells[2, 1, 2, 3].Merge = true; ws.Cells[2, 1].Value = "KỲ THI HỌC SINH GIỎI THPT CẤP TỈNH"; ws.Cells[2, 1].Style.Font.Bold = true;
        ws.Cells[3, 1, 3, 3].Merge = true; ws.Cells[3, 1].Value = $"Khóa ngày {DateTime.Now:dd/MM/yyyy}"; ws.Cells[3, 1].Style.Font.UnderLine = true;

        ws.Cells[1, 4, 1, 5].Merge = true; ws.Cells[1, 4].Value = $"DANH SÁCH PHÒNG THI SỐ: {phong}"; ws.Cells[1, 4].Style.Font.Bold = true;
        ws.Cells[2, 4, 2, 5].Merge = true; ws.Cells[2, 4].Value = $"MÔN: {monThi.ToUpper()}"; ws.Cells[2, 4].Style.Font.Bold = true;

        ws.Cells[5, 1].Value = "Số Báo Danh"; ws.Cells[5, 2].Value = "Họ và Tên"; ws.Cells[5, 3].Value = "Ngày Sinh"; ws.Cells[5, 4].Value = "Học Sinh Trường"; ws.Cells[5, 5].Value = "Ghi Chú";

        var thiSinhs = await db.ThiSinhs.Where(x => x.PhongThi == phong).OrderBy(x => x.SBD).ToListAsync();
        for (int i = 0; i < thiSinhs.Count; i++)
        {
            int row = i + 6;
            ws.Cells[row, 1].Value = thiSinhs[i].SBD;
            ws.Cells[row, 2].Value = thiSinhs[i].HoTen;
            ws.Cells[row, 3].Value = thiSinhs[i].NgayDangKy.ToString("dd/MM/yyyy");
            ws.Cells[row, 4].Value = thiSinhs[i].TenTruong;
            ws.Cells[row, 5].Value = "";
        }

        int lastRow = thiSinhs.Count + 8;
        ws.Cells[lastRow, 4, lastRow, 5].Merge = true; ws.Cells[lastRow, 4].Value = "CHỦ TỊCH HỘI ĐỒNG/BAN COI THI"; ws.Cells[lastRow, 4].Style.Font.Bold = true;
        if (ws.Dimension != null) ws.Cells[ws.Dimension.Address].AutoFitColumns();
    }

    var bytes = await package.GetAsByteArrayAsync();
    return Results.File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DanhSach_PhongThi_VinhLong.xlsx");
}).DisableAntiforgery();

app.MapGet("/api/truong/download-mau", async (AppDbContext db) =>
{
    using var package = new ExcelPackage();
    var ws = package.Workbook.Worksheets.Add("MauDangKy");
    ws.Cells[1, 1].Value = "Họ và Tên"; ws.Cells[1, 2].Value = "Số CCCD/Định Danh"; ws.Cells[1, 3].Value = "Môn Dự Thi";
    var mons = await db.DanhMucMonThis.Select(x => x.TenMon).ToListAsync();
    ws.Cells[2, 1].Value = "Nguyễn Văn A"; ws.Cells[2, 2].Value = "038200123456"; ws.Cells[2, 3].Value = mons.FirstOrDefault() ?? "Toán";
    var bytes = await package.GetAsByteArrayAsync();
    return Results.File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Mau_Dang_Ky_HSG.xlsx");
});

app.MapPost("/api/truong/upload-excel", async (IFormFile file, AppDbContext db, HttpContext ctx) =>
{
    if (!ctx.User.IsInRole("Truong")) return Results.Unauthorized();
    var maTruong = ctx.User.FindFirstValue(ClaimTypes.Name);
    var truongInfo = await db.TaiKhoans.FirstOrDefaultAsync(x => x.MaTruong == maTruong);

    if (file == null || file.Length == 0) return Results.BadRequest(new { message = "Chọn file Excel!" });

    using var stream = new MemoryStream();
    await file.CopyToAsync(stream);
    using var package = new ExcelPackage(stream);
    var ws = package.Workbook.Worksheets[0];
    int rowCount = ws.Dimension?.Rows ?? 0;
    int countSuccess = 0;

    for (int row = 2; row <= rowCount; row++)
    {
        var hoTen = ws.Cells[row, 1].Value?.ToString()?.Trim();
        var cccd = ws.Cells[row, 2].Value?.ToString()?.Trim();
        var monThi = ws.Cells[row, 3].Value?.ToString()?.Trim();
        if (string.IsNullOrEmpty(hoTen) || string.IsNullOrEmpty(cccd)) continue;

        if (!await db.ThiSinhs.AnyAsync(x => x.CCCD == cccd))
        {
            db.ThiSinhs.Add(new ThiSinh
            {
                HoTen = hoTen, CCCD = cccd, MonThi = monThi ?? "Toán", TenTruong = truongInfo?.TenTruong ?? maTruong, MaTruong = maTruong, NgayDangKy = DateTime.UtcNow
            });
            countSuccess++;
        }
    }
    await db.SaveChangesAsync();
    return Results.Ok(new { message = $"Đã nhập thành công {countSuccess} học sinh!" });
}).DisableAntiforgery();

app.MapPut("/api/truong/sua-thisinh/{id}", async (int id, AppDbContext db, HttpContext ctx, UpdateThiSinhReq req) =>
{
    if (!ctx.User.IsInRole("Truong")) return Results.Unauthorized();
    var maTruong = ctx.User.FindFirstValue(ClaimTypes.Name);
    var ts = await db.ThiSinhs.FirstOrDefaultAsync(x => x.Id == id && x.MaTruong == maTruong);
    if (ts == null) return Results.NotFound();

    ts.HoTen = req.HoTen?.Trim() ?? ts.HoTen; ts.CCCD = req.CCCD?.Trim() ?? ts.CCCD; ts.MonThi = req.MonThi;
    if (!string.IsNullOrEmpty(req.TenTruong)) ts.TenTruong = req.TenTruong.Trim();
    await db.SaveChangesAsync();
    return Results.Ok(new { message = "Cập nhật thành công!" });
}).DisableAntiforgery();

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

// --- DATA MODELS ---
public class AppDbContext : DbContext
{
    public AppDbContext(DbContextOptions<AppDbContext> options) : base(options) { }
    public DbSet<ThiSinh> ThiSinhs { get; set; }
    public DbSet<TaiKhoan> TaiKhoans { get; set; }
    public DbSet<DanhMucMonThi> DanhMucMonThis { get; set; }
    public DbSet<CanBoCoiThi> CanBoCoiThis { get; set; }
}

public class CanBoCoiThi
{
    public int Id { get; set; }
    public string HoTen { get; set; } = "";
    public string DonViCongTac { get; set; } = "";
    public string? ChucVu { get; set; }
    public string? Sdt { get; set; }
}

public class DanhMucMonThi { public int Id { get; set; } public string? MaMon { get; set; } public string TenMon { get; set; } = ""; }
public class LoginRequest { public string MaTruong { get; set; } = ""; public string MatKhau { get; set; } = ""; }
public class ChiaPhongRequest { public int SoHocSinhMoiPhong { get; set; } = 24; }
public class UpdateThiSinhReq { public string HoTen { get; set; } = ""; public string CCCD { get; set; } = ""; public string MonThi { get; set; } = ""; public string TenTruong { get; set; } = ""; }
public class TaiKhoan { public int Id { get; set; } public string MaTruong { get; set; } = ""; public string TenTruong { get; set; } = ""; public string MatKhau { get; set; } = ""; public string Role { get; set; } = "Truong"; }
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
    public string? GiamThi1 { get; set; }
    public string? GiamThi2 { get; set; }
    public DateTime NgayDangKy { get; set; }
}
