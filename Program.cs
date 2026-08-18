using Microsoft.EntityFrameworkCore;
using Npgsql;

// Chuyển cơ chế theo dõi file sang Polling để tránh lỗi inotify trên Render
Environment.SetEnvironmentVariable("DOTNET_USE_POLLING_FILE_WATCHER", "1");

var builder = WebApplication.CreateBuilder(args);

// Xóa nguồn theo dõi file appsettings để tránh kích hoạt inotify
builder.Configuration.Sources.Clear();
builder.Configuration.AddEnvironmentVariables();

builder.Services.AddRazorPages();

// Đọc chuỗi kết nối từ biến môi trường
var rawConnectionString = Environment.GetEnvironmentVariable("DATABASE_URL") 
    ?? builder.Configuration.GetConnectionString("DefaultConnection");

// Tự động phân tích URL dạng postgresql:// sang định dạng Npgsql chuẩn
string connectionString = ParsePostgresUrl(rawConnectionString);

builder.Services.AddDbContext<AppDbContext>(options =>
    options.UseNpgsql(connectionString));

var app = builder.Build();

using (var scope = app.Services.CreateScope())
{
    var db = scope.ServiceProvider.GetRequiredService<AppDbContext>();
    db.Database.EnsureCreated();
}

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

// Hàm hỗ trợ chuyển đổi URL PostgreSQL sang định dạng ADO.NET Npgsql
static string ParsePostgresUrl(string inputUrl)
{
    if (string.IsNullOrEmpty(inputUrl)) return "";
    if (!inputUrl.StartsWith("postgres://") && !inputUrl.StartsWith("postgresql://"))
    {
        return inputUrl; // Nếu đã ở dạng Host=...;Password=... thì giữ nguyên
    }

    var uri = new Uri(inputUrl);
    var userInfo = uri.UserInfo.Split(':');
    var builder = new NpgsqlConnectionStringBuilder
    {
        Host = uri.Host,
        Port = uri.Port > 0 ? uri.Port : 5432,
        Username = userInfo[0],
        Password = userInfo.Length > 1 ? Uri.UnescapeDataString(userInfo[1]) : "",
        Database = uri.AbsolutePath.TrimStart('/'),
        SslMode = SslMode.Require,
        TrustServerCertificate = true
    };
    return builder.ConnectionString;
}

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
