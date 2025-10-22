using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;
using QuanLy.Data;
using QuanLy.Models;

var builder = WebApplication.CreateBuilder(args);

// Cấu hình kết nối CSDL
builder.Services.AddDbContext<ApplicationDbContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection")));

// Cấu hình Identity
builder.Services.AddIdentity<NguoiDung, IdentityRole>()
    .AddEntityFrameworkStores<ApplicationDbContext>()
    .AddDefaultTokenProviders();

// Thêm MVC
builder.Services.AddControllersWithViews();

var app = builder.Build();

// Gọi Seed dữ liệu (vai trò, admin mẫu)
using (var scope = app.Services.CreateScope())
{
    var services = scope.ServiceProvider;
    await DbInitializer.SeedRolesAndAdminAsync(services);
}

// Cấu hình pipeline
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles(); // Đừng quên dòng này

app.UseRouting();

app.UseAuthentication(); // Quan trọng
app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.Run();
