using Microsoft.EntityFrameworkCore;

Environment.SetEnvironmentVariable("DOTNET_USE_POLLING_FILE_WATCHER", "1");

var builder = WebApplication.CreateBuilder(args);

builder.Configuration.Sources.Clear();
builder.Configuration.AddEnvironmentVariables();

builder.Services.AddRazorPages();

// Đọc chuỗi kết nối trực tiếp từ Render
var connectionString = Environment.GetEnvironmentVariable("DATABASE_URL");

builder.Services.AddDbContext<AppDbContext>(options =>
    options.UseNpgsql(connectionString));

var app = builder.Build();

// Tự động khởi tạo bảng khi ứng dụng nhận request đầu tiên (tránh crash lúc startup)
app.Use(async (context, next) =>
{
    using (var scope = app.Services.CreateScope())
    {
        var db = scope.ServiceProvider.GetRequiredService<AppDbContext>();
        await db.Database.EnsureCreatedAsync();
    }
    await next();
});

app.UseStaticFiles();
app.UseRouting();
app.MapRazorPages();

app.MapPost("/api/dang-ky", async (AppDbContext db, ThiSinh model) =>
{
    if (string.IsNullOrEmpty(model.HoTen) || string.IsNullOrEmpty(model.CCCD))
        return Results.BadRequest(new { message = "Vui lòng điền đầy đủ thông tin!" });

    var exists = await db.ThiSinhs.AnyAsync(x => x.CCCD == model.CCCD);
    if (exists)
        return Results.BadRequest(new { message = "Số CCCD/Định danh này đã được đăng ký!" });

    model.NgayDangKy = DateTime.UtcNow;
    model.TrangThai = "Chờ duyệt";
    db.ThiSinhs.Add(model);
    await db.SaveChangesAsync();

    return Results.Ok(new { message = "Đăng ký thành công!" });
});

app.MapPost("/api/duyet/{id}", async (int id, AppDbContext db) =>
{
    var ts = await db.ThiSinhs.FindAsync(id);
    if (ts == null) return Results.NotFound();
    ts.TrangThai = "Đã duyệt";
    await db.SaveChangesAsync();
    return Results.Ok();
});

app.Run();

public class AppDbContext : DbContext
{
    public AppDbContext(DbContextOptions<AppDbContext> options) : base(options) { }
    public DbSet<ThiSinh> ThiSinhs { get; set; }
}

public class ThiSinh
{
    public int Id { get; set; }
    public string HoTen { get; set; } = "";
    public string CCCD { get; set; } = "";
    public string TenTruong { get; set; } = "";
    public string MonThi { get; set; } = "";
    public string TrangThai { get; set; } = "Chờ duyệt";
    public DateTime NgayDangKy { get; set; }
}
