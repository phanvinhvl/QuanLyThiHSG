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

// --- TỰ ĐỘNG CẬP NHẬT DATABASE KHI APP KHỞI ĐỘNG (CHỐNG LỖI 500 TOÀN HỆ THỐNG) ---
using (var scope = app.Services.CreateScope())
{
    var db = scope.ServiceProvider.GetRequiredService<AppDbContext>();
    try
    {
        // 1. Tạo bảng nếu chưa có
        db.Database.ExecuteSqlRaw(@"CREATE TABLE IF NOT EXISTS ""TaiKhoans"" (""Id"" SERIAL PRIMARY KEY, ""MaTruong"" TEXT NOT NULL, ""TenTruong"" TEXT NOT NULL, ""MatKhau"" TEXT NOT NULL, ""Role"" TEXT NOT NULL);");
        db.Database.ExecuteSqlRaw(@"CREATE TABLE IF NOT EXISTS ""ThiSinhs"" (""Id"" SERIAL PRIMARY KEY, ""HoTen"" TEXT NOT NULL, ""NgaySinh"" TEXT, ""CCCD"" TEXT NOT NULL, ""TenTruong"" TEXT NOT NULL, ""MaTruong"" TEXT NOT NULL, ""MonThi"" TEXT NOT NULL, ""DiemTbmMon"" DOUBLE PRECISION NOT NULL DEFAULT 0, ""KetQuaHocTap"" TEXT, ""NgayDangKy"" TIMESTAMP WITH TIME ZONE NOT NULL, ""SBD"" TEXT, ""PhongThi"" TEXT);");
        db.Database.ExecuteSqlRaw(@"CREATE TABLE IF NOT EXISTS ""DanhMucMonThis"" (""Id"" SERIAL PRIMARY KEY, ""MaMon"" TEXT, ""TenMon"" TEXT NOT NULL);");
        db.Database.ExecuteSqlRaw(@"CREATE TABLE IF NOT EXISTS ""CauHinhKyThis"" (""Id"" SERIAL PRIMARY KEY, ""TenKyThi"" TEXT NOT NULL, ""KhoaNgay"" TEXT NOT NULL);");

        // 2. Ép buộc bổ sung toàn bộ các cột mới nếu CSDL cũ còn thiếu
        try { db.Database.ExecuteSqlRaw(@"ALTER TABLE ""ThiSinhs"" ADD COLUMN IF NOT EXISTS ""NgaySinh"" TEXT;"); } catch {}
        try { db.Database.ExecuteSqlRaw(@"ALTER TABLE ""ThiSinhs"" ADD COLUMN IF NOT EXISTS ""DiemTbmMon"" DOUBLE PRECISION NOT NULL DEFAULT 0;"); } catch {}
        try { db.Database.ExecuteSqlRaw(@"ALTER TABLE ""ThiSinhs"" ADD COLUMN IF NOT EXISTS ""KetQuaHocTap"" TEXT;"); } catch {}
        try { db.Database.ExecuteSqlRaw(@"ALTER TABLE ""DanhMucMonThis"" ADD COLUMN IF NOT EXISTS ""MaMon"" TEXT;"); } catch {}

        // 3. Khởi tạo tài khoản Sở mặc định
        if (!db.TaiKhoans.Any(x => x.Role == "So"))
        {
            db.TaiKhoans.Add(new TaiKhoan { MaTruong = "SO_GD", TenTruong = "Sở Giáo Dục & Đào Tạo", MatKhau = "so@123456", Role = "So" });
            db.SaveChanges();
        }

        // 4. Khởi tạo Cấu hình Kỳ thi mặc định
        if (!db.CauHinhKyThis.Any())
        {
            db.CauHinhKyThis.Add(new CauHinhKyThi { TenKyThi = "KỲ THI HỌC SINH GIỎI THPT CẤP TỈNH", KhoaNgay = DateTime.Now.ToString("dd/MM/yyyy") });
            db.SaveChanges();
        }

        // 5. Khởi tạo Danh mục môn mặc định
        if (!db.DanhMucMonThis.Any())
        {
            var defaultMons = new[] { "Toán", "Ngữ Văn", "Tiếng Anh", "Vật Lý", "Hóa Học", "Sinh Học", "Tin Học", "Lịch Sử", "Địa Lý" };
            for (int i = 0; i < defaultMons.Length; i++)
                db.DanhMucMonThis.Add(new DanhMucMonThi { MaMon = (i + 1).ToString("D2"), TenMon = defaultMons[i] });
            db.SaveChanges();
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine("Lỗi khởi tạo DB: " + ex.Message);
    }
}

app.UseStaticFiles();
app.UseRouting();
app.UseAuthentication();
app.UseAuthorization();
app.MapRazorPages();

app.MapGet("/", () => Results.Redirect("/login"));

// --- API AUTH ---
app.MapPost("/api/auth/login", async (AppDbContext db, HttpContext ctx, LoginRequest req) =>
{
    var acc = await db.TaiKhoans.FirstOrDefaultAsync(x => x.MaTruong == req.MaTruong && x.MatKhau == req.MatKhau);
    if (acc == null) return Results.BadRequest(new { message = "Mã tài khoản hoặc mật khẩu không chính xác!" });

    var claims = new List<Claim> { new Claim(ClaimTypes.Name, acc.MaTruong), new Claim("TenTruong", acc.TenTruong), new Claim(ClaimTypes.Role, acc.Role) };
    var identity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
    await ctx.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, new ClaimsPrincipal(identity));

    string redirectUrl = acc.Role == "So" ? "/adminso" : "/admintruong";
    return Results.Ok(new { redirectUrl });
});

// --- API TẢI FILE EXCEL MẪU ---
app.MapGet("/api/truong/download-mau", async (AppDbContext db) =>
{
    using var package = new ExcelPackage();
    var ws = package.Workbook.Worksheets.Add("MauDangKy");
    
    ws.Cells[1, 1].Value = "Họ và Tên";
    ws.Cells[1, 2].Value = "Ngày Sinh (dd/MM/yyyy)";
    ws.Cells[1, 3].Value = "Số CCCD/Định Danh";
    ws.Cells[1, 4].Value = "Môn Dự Thi";
    ws.Cells[1, 5].Value = "Điểm TBM Dự Thi";
    ws.Cells[1, 6].Value = "Kết Quả Học Tập";

    using (var range = ws.Cells[1, 1, 1, 6])
    {
        range.Style.Font.Bold = true;
        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
    }

    var mons = await db.DanhMucMonThis.Select(x => x.TenMon).ToListAsync();
    ws.Cells[2, 1].Value = "Nguyễn Văn A";
    ws.Cells[2, 2].Value = "15/05/2008";
    ws.Cells[2, 3].Value = "038200123456";
    ws.Cells[2, 4].Value = mons.FirstOrDefault() ?? "Toán";
    ws.Cells[2, 5].Value = 8.5;
    ws.Cells[2, 6].Value = "Tốt";

    ws.Cells[ws.Dimension.Address].AutoFitColumns();

    var bytes = await package.GetAsByteArrayAsync();
    return Results.File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Mau_Dang_Ky_HSG.xlsx");
});

// --- API UPLOAD EXCEL ĐĂNG KÝ HỌC SINH ---
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
    List<string> dsLoi = new List<string>();

    for (int row = 2; row <= rowCount; row++)
    {
        var hoTen = ws.Cells[row, 1].Value?.ToString()?.Trim();
        var ngaySinh = ws.Cells[row, 2].Value?.ToString()?.Trim();
        var cccd = ws.Cells[row, 3].Value?.ToString()?.Trim();
        var monThi = ws.Cells[row, 4].Value?.ToString()?.Trim();
        var diemStr = ws.Cells[row, 5].Value?.ToString()?.Trim();
        var kqHocTap = ws.Cells[row, 6].Value?.ToString()?.Trim();

        if (string.IsNullOrEmpty(hoTen) && string.IsNullOrEmpty(cccd)) continue;

        if (string.IsNullOrEmpty(hoTen)) { dsLoi.Add($"Dòng {row}: Thiếu Họ và Tên!"); continue; }
        if (string.IsNullOrEmpty(cccd)) { dsLoi.Add($"Dòng {row}: Thiếu Số CCCD!"); continue; }
        if (string.IsNullOrEmpty(monThi) || !dsMonValid.Contains(monThi.ToLower())) { dsLoi.Add($"Dòng {row}: Môn '{monThi}' không có trong danh mục của Sở!"); continue; }

        if (!double.TryParse(diemStr, out double diemTbm))
        {
            dsLoi.Add($"Dòng {row} ({hoTen}): Điểm TBM môn không hợp lệ!");
            continue;
        }

        string kqNorm = (kqHocTap ?? "").ToLower();
        if (kqNorm != "khá" && kqNorm != "kha" && kqNorm != "tốt" && kqNorm != "tot" && kqNorm != "giỏi" && kqNorm != "gioi")
        {
            dsLoi.Add($"Dòng {row} ({hoTen}): Bị loại do Kết quả học tập '{kqHocTap}' (Yêu cầu phải từ Khá trở lên)!");
            continue;
        }

        bool isNgatVan = monThi.ToLower().Contains("ngữ văn") || monThi.ToLower().Contains("ngu van") || monThi.ToLower().Contains("văn");
        double minDiem = isNgatVan ? 7.5 : 8.0;

        if (diemTbm < minDiem)
        {
            dsLoi.Add($"Dòng {row} ({hoTen}): Bị loại do Điểm TBM {monThi} đạt {diemTbm} (Yêu cầu môn {monThi} phải từ {minDiem} trở lên)!");
            continue;
        }

        var exists = await db.ThiSinhs.AnyAsync(x => x.CCCD == cccd);
        if (!exists)
        {
            var monChuan = await db.DanhMucMonThis.FirstAsync(x => x.TenMon.ToLower() == monThi.ToLower());
            db.ThiSinhs.Add(new ThiSinh
            {
                HoTen = hoTen,
                NgaySinh = ngaySinh ?? "",
                CCCD = cccd,
                MonThi = monChuan.TenMon,
                DiemTbmMon = diemTbm,
                KetQuaHocTap = kqHocTap ?? "Khá",
                TenTruong = truongInfo?.TenTruong ?? maTruong,
                MaTruong = maTruong,
                NgayDangKy = DateTime.UtcNow
            });
            countSuccess++;
        }
        else
        {
            dsLoi.Add($"Dòng {row} ({hoTen}): Số CCCD '{cccd}' đã tồn tại trong hệ thống!");
        }
    }

    await db.SaveChangesAsync();
    string msg = $"Đăng ký thành công {countSuccess} học sinh đủ điều kiện!";
    return Results.Ok(new { message = msg, dsLoi });
}).DisableAntiforgery();

app.MapPut("/api/truong/sua-thisinh/{id}", async (int id, AppDbContext db, HttpContext ctx, UpdateThiSinhReq req) =>
{
    if (!ctx.User.IsInRole("Truong")) return Results.Unauthorized();
    var maTruong = ctx.User.FindFirstValue(ClaimTypes.Name);

    var ts = await db.ThiSinhs.FirstOrDefaultAsync(x => x.Id == id && x.MaTruong == maTruong);
    if (ts == null) return Results.NotFound(new { message = "Không tìm thấy học sinh!" });

    var monChuan = await db.DanhMucMonThis.FirstOrDefaultAsync(x => x.TenMon.ToLower() == req.MonThi.ToLower());
    if (monChuan == null) return Results.BadRequest(new { message = "Môn thi không thuộc danh mục quy định!" });

    string kqNorm = (req.KetQuaHocTap ?? "").ToLower();
    if (kqNorm != "khá" && kqNorm != "kha" && kqNorm != "tốt" && kqNorm != "tot" && kqNorm != "giỏi" && kqNorm != "gioi")
        return Results.BadRequest(new { message = "Kết quả học tập phải đạt từ Khá trở lên!" });

    bool isNgatVan = req.MonThi.ToLower().Contains("ngữ văn") || req.MonThi.ToLower().Contains("văn");
    double minDiem = isNgatVan ? 7.5 : 8.0;

    if (req.DiemTbmMon < minDiem)
        return Results.BadRequest(new { message = $"Điểm TBM môn {req.MonThi} đạt {req.DiemTbmMon} không đủ điều kiện (Yêu cầu >= {minDiem})!" });

    ts.HoTen = req.HoTen?.Trim() ?? ts.HoTen;
    ts.NgaySinh = req.NgaySinh?.Trim() ?? ts.NgaySinh;
    ts.CCCD = req.CCCD?.Trim() ?? ts.CCCD;
    ts.MonThi = monChuan.TenMon;
    ts.DiemTbmMon = req.DiemTbmMon;
    ts.KetQuaHocTap = req.KetQuaHocTap?.Trim() ?? ts.KetQuaHocTap;
    if (!string.IsNullOrEmpty(req.TenTruong)) ts.TenTruong = req.TenTruong.Trim();

    await db.SaveChangesAsync();
    return Results.Ok(new { message = "Cập nhật thông tin học sinh thành công!" });
}).DisableAntiforgery();

app.MapPost("/api/so/luu-cau-hinh", async (AppDbContext db, HttpContext ctx, CauHinhKyThi req) =>
{
    if (!ctx.User.IsInRole("So")) return Results.Unauthorized();
    var config = await db.CauHinhKyThis.FirstOrDefaultAsync();
    if (config == null) db.CauHinhKyThis.Add(new CauHinhKyThi { TenKyThi = req.TenKyThi.Trim(), KhoaNgay = req.KhoaNgay.Trim() });
    else { config.TenKyThi = req.TenKyThi.Trim(); config.KhoaNgay = req.KhoaNgay.Trim(); }
    await db.SaveChangesAsync();
    return Results.Ok(new { message = "Đã lưu cấu hình!" });
}).DisableAntiforgery();

app.MapPut("/api/so/sua-truong/{id}", async (int id, AppDbContext db, HttpContext ctx, UpdateTruongReq req) =>
{
    if (!ctx.User.IsInRole("So")) return Results.Unauthorized();
    var truong = await db.TaiKhoans.FirstOrDefaultAsync(x => x.Id == id && x.Role == "Truong");
    if (truong == null) return Results.NotFound();
    var maCu = truong.MaTruong;
    truong.MaTruong = req.MaTruong.Trim(); truong.TenTruong = req.TenTruong.Trim();
    if (!string.IsNullOrWhiteSpace(req.MatKhau)) truong.MatKhau = req.MatKhau.Trim();
    if (maCu != truong.MaTruong || truong.TenTruong != req.TenTruong)
    {
        var listTS = await db.ThiSinhs.Where(x => x.MaTruong == maCu).ToListAsync();
        foreach (var ts in listTS) { ts.MaTruong = truong.MaTruong; ts.TenTruong = truong.TenTruong; }
    }
    await db.SaveChangesAsync();
    return Results.Ok(new { message = "Cập nhật thành công!" });
}).DisableAntiforgery();

app.MapDelete("/api/so/xoa-truong/{id}", async (int id, AppDbContext db, HttpContext ctx) =>
{
    if (!ctx.User.IsInRole("So")) return Results.Unauthorized();
    var truong = await db.TaiKhoans.FirstOrDefaultAsync(x => x.Id == id && x.Role == "Truong");
    if (truong == null) return Results.NotFound();
    var listTS = await db.ThiSinhs.Where(x => x.MaTruong == truong.MaTruong).ToListAsync();
    db.ThiSinhs.RemoveRange(listTS);
    db.TaiKhoans.Remove(truong);
    await db.SaveChangesAsync();
    return Results.Ok(new { message = "Đã xóa trường!" });
}).DisableAntiforgery();

app.MapPost("/api/so/them-mon", async (AppDbContext db, HttpContext ctx, DanhMucMonThi req) =>
{
    if (!ctx.User.IsInRole("So")) return Results.Unauthorized();
    int count = await db.DanhMucMonThis.CountAsync();
    db.DanhMucMonThis.Add(new DanhMucMonThi { MaMon = (count + 1).ToString("D2"), TenMon = req.TenMon.Trim() });
    await db.SaveChangesAsync();
    return Results.Ok(new { message = "Thêm môn thành công!" });
}).DisableAntiforgery();

app.MapPut("/api/so/sua-mon/{id}", async (int id, AppDbContext db, HttpContext ctx, DanhMucMonThi req) =>
{
    if (!ctx.User.IsInRole("So")) return Results.Unauthorized();
    var item = await db.DanhMucMonThis.FindAsync(id);
    if (item == null) return Results.NotFound();
    item.TenMon = req.TenMon.Trim();
    await db.SaveChangesAsync();
    return Results.Ok(new { message = "Cập nhật môn thành công!" });
}).DisableAntiforgery();

app.MapDelete("/api/so/xoa-mon/{id}", async (int id, AppDbContext db, HttpContext ctx) =>
{
    if (!ctx.User.IsInRole("So")) return Results.Unauthorized();
    var item = await db.DanhMucMonThis.FindAsync(id);
    if (item != null) { db.DanhMucMonThis.Remove(item); await db.SaveChangesAsync(); }
    return Results.Ok(new { message = "Đã xóa môn!" });
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
    
    ws.Cells[1, 1].Value = "STT";
    ws.Cells[1, 2].Value = "SBD";
    ws.Cells[1, 3].Value = "Phòng Thi";
    ws.Cells[1, 4].Value = "Họ và Tên";
    ws.Cells[1, 5].Value = "Ngày Sinh";
    ws.Cells[1, 6].Value = "Số CCCD";
    ws.Cells[1, 7].Value = "Trường";
    ws.Cells[1, 8].Value = "Môn Thi";
    ws.Cells[1, 9].Value = "Điểm TBM";
    ws.Cells[1, 10].Value = "KQ Học Tập";

    for (int i = 0; i < list.Count; i++)
    {
        ws.Cells[i + 2, 1].Value = i + 1;
        ws.Cells[i + 2, 2].Value = list[i].SBD;
        ws.Cells[i + 2, 3].Value = list[i].PhongThi;
        ws.Cells[i + 2, 4].Value = list[i].HoTen;
        ws.Cells[i + 2, 5].Value = list[i].NgaySinh;
        ws.Cells[i + 2, 6].Value = list[i].CCCD;
        ws.Cells[i + 2, 7].Value = list[i].TenTruong;
        ws.Cells[i + 2, 8].Value = list[i].MonThi;
        ws.Cells[i + 2, 9].Value = list[i].DiemTbmMon;
        ws.Cells[i + 2, 10].Value = list[i].KetQuaHocTap;
    }
    var bytes = await package.GetAsByteArrayAsync();
    return Results.File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DanhSach_ThiSinh_TongHop.xlsx");
});

app.MapGet("/api/so/export-phong-thi", async (AppDbContext db, HttpContext ctx) =>
{
    var config = await db.CauHinhKyThis.FirstOrDefaultAsync();
    string tenKyThi = config?.TenKyThi ?? "KỲ THI HỌC SINH GIỎI THPT CẤP TỈNH";
    string khoaNgay = config?.KhoaNgay ?? DateTime.Now.ToString("dd/MM/yyyy");

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
        ws.Cells[2, 1, 2, 3].Merge = true; ws.Cells[2, 1].Value = tenKyThi.ToUpper(); ws.Cells[2, 1].Style.Font.Bold = true;
        ws.Cells[3, 1, 3, 3].Merge = true; ws.Cells[3, 1].Value = $"Khóa ngày {khoaNgay}"; ws.Cells[3, 1].Style.Font.UnderLine = true;

        ws.Cells[1, 4, 1, 5].Merge = true; ws.Cells[1, 4].Value = $"DANH SÁCH PHÒNG THI SỐ: {phong}"; ws.Cells[1, 4].Style.Font.Bold = true;
        ws.Cells[2, 4, 2, 5].Merge = true; ws.Cells[2, 4].Value = $"MÔN: {monThi.ToUpper()}"; ws.Cells[2, 4].Style.Font.Bold = true;

        ws.Cells[5, 1].Value = "Số Báo Danh"; ws.Cells[5, 2].Value = "Họ và Tên"; ws.Cells[5, 3].Value = "Ngày Sinh"; ws.Cells[5, 4].Value = "Học Sinh Trường"; ws.Cells[5, 5].Value = "Ghi Chú";

        var thiSinhs = await db.ThiSinhs.Where(x => x.PhongThi == phong).OrderBy(x => x.SBD).ToListAsync();
        for (int i = 0; i < thiSinhs.Count; i++)
        {
            int row = i + 6;
            ws.Cells[row, 1].Value = thiSinhs[i].SBD;
            ws.Cells[row, 2].Value = thiSinhs[i].HoTen;
            ws.Cells[row, 3].Value = thiSinhs[i].NgaySinh;
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
    public DbSet<CauHinhKyThi> CauHinhKyThis { get; set; }
}

public class CauHinhKyThi { public int Id { get; set; } public string TenKyThi { get; set; } = ""; public string KhoaNgay { get; set; } = ""; }
public class UpdateTruongReq { public string MaTruong { get; set; } = ""; public string TenTruong { get; set; } = ""; public string? MatKhau { get; set; } }
public class DanhMucMonThi { public int Id { get; set; } public string? MaMon { get; set; } public string TenMon { get; set; } = ""; }
public class LoginRequest { public string MaTruong { get; set; } = ""; public string MatKhau { get; set; } = ""; }
public class ChiaPhongRequest { public int SoHocSinhMoiPhong { get; set; } = 24; }
public class UpdateThiSinhReq
{
    public string HoTen { get; set; } = "";
    public string NgaySinh { get; set; } = "";
    public string CCCD { get; set; } = "";
    public string MonThi { get; set; } = "";
    public double DiemTbmMon { get; set; }
    public string KetQuaHocTap { get; set; } = "";
    public string TenTruong { get; set; } = "";
}
public class TaiKhoan { public int Id { get; set; } public string MaTruong { get; set; } = ""; public string TenTruong { get; set; } = ""; public string MatKhau { get; set; } = ""; public string Role { get; set; } = "Truong"; }
public class ThiSinh
{
    public int Id { get; set; }
    public string HoTen { get; set; } = "";
    public string NgaySinh { get; set; } = "";
    public string CCCD { get; set; } = "";
    public string TenTruong { get; set; } = "";
    public string MaTruong { get; set; } = "";
    public string MonThi { get; set; } = "";
    public double DiemTbmMon { get; set; }
    public string KetQuaHocTap { get; set; } = "Khá";
    public string? SBD { get; set; }
    public string? PhongThi { get; set; }
    public DateTime NgayDangKy { get; set; }
}
