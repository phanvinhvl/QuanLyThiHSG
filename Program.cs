using Microsoft.EntityFrameworkCore;
var builder = WebApplication.CreateBuilder(new WebApplicationOptions
{
    Args = args
});
builder.Host.ConfigureAppConfiguration((hostingContext, config) =>
{
    config.Sources.Clear();
    config.AddJsonFile("appsettings.json", optional: true, reloadOnChange: false);
    config.AddEnvironmentVariables();
});

// Cấu hình Razor Pages & Database
builder.Services.AddRazorPages();
var connectionString = builder.Configuration.GetConnectionString("DefaultConnection") 
    ?? Environment.GetEnvironmentVariable("DATABASE_URL");

builder.Services.AddDbContext<AppDbContext>(options =>
    options.UseNpgsql(connectionString));

var app = builder.Build();

// Tự động tạo bảng Database khi khởi chạy
using (var scope = app.Services.CreateScope())
{
    var db = scope.ServiceProvider.GetRequiredService<AppDbContext>();
    db.Database.EnsureCreated();
}

app.UseStaticFiles();
app.UseRouting();
app.MapRazorPages();

// API tiếp nhận đăng ký từ Form
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

// API duyệt hồ sơ cho Quản trị viên
app.MapPost("/api/duyet/{id}", async (int id, AppDbContext db) =>
{
    var ts = await db.ThiSinhs.FindAsync(id);
    if (ts == null) return Results.NotFound();
    ts.TrangThai = "Đã duyệt";
    await db.SaveChangesAsync();
    return Results.Ok();
});

app.Run();

// Data Models
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
