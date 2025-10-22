using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore;
using QuanLy.Models;

namespace QuanLy.Data
{
    public class ApplicationDbContext : IdentityDbContext<NguoiDung>
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
            : base(options) { }

        public DbSet<PhongBan> PhongBans { get; set; }
        public DbSet<BaoCaoTuan> BaoCaoTuans { get; set; }
        public DbSet<NoiDungBaoCao> NoiDungBaoCaos { get; set; }

        protected override void OnModelCreating(ModelBuilder builder)
        {
            base.OnModelCreating(builder);

            builder.Entity<PhongBan>()
                .HasKey(p => p.MaPhongBan);

            builder.Entity<NguoiDung>()
                .HasOne(u => u.PhongBan)
                .WithMany(p => p.NguoiDungs)
                .HasForeignKey(u => u.MaPhongBan);

            // === Cấu hình khóa ngoại cho BaoCaoTuan để tránh multiple cascade paths ===
            builder.Entity<BaoCaoTuan>()
                .HasOne(b => b.NguoiBaoCao)
                .WithMany()
                .HasForeignKey(b => b.NguoiBaoCaoId)
                .OnDelete(DeleteBehavior.Restrict);

            builder.Entity<BaoCaoTuan>()
                .HasOne(b => b.BaoCaoCho)
                .WithMany()
                .HasForeignKey(b => b.BaoCaoChoId)
                .OnDelete(DeleteBehavior.Restrict);

        }
    }
}
