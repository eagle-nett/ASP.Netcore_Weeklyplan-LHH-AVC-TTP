using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;
using QuanLy.Models;

namespace QuanLy.Data
{
    public static class DbInitializer
    {
        public static async Task SeedRolesAndAdminAsync(IServiceProvider serviceProvider)
        {
            var roleManager = serviceProvider.GetRequiredService<RoleManager<IdentityRole>>();
            var userManager = serviceProvider.GetRequiredService<UserManager<NguoiDung>>();
            var context = serviceProvider.GetRequiredService<ApplicationDbContext>();

            // 1. Tạo các vai trò nếu chưa có
            string[] roles = { "Admin", "GiamDoc", "TruongPhong", "NhanVien" };
            foreach (var role in roles)
            {
                if (!await roleManager.RoleExistsAsync(role))
                    await roleManager.CreateAsync(new IdentityRole(role));
            }

            // 2. Thêm phòng ban mẫu nếu chưa có
            var phongBans = new[]
            {
                new PhongBan { MaPhongBan = "LBGD", TenPhongBan = "Ban Giám Đốc", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "LBHT", TenPhongBan = "Bộ phận Hoàn Thiện", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "LKBK", TenPhongBan = "Bộ phận Kho BTP Khung, Sơn, NVL", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "LKBD", TenPhongBan = "Bộ phận kho BTP Đan", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "LKTP", TenPhongBan = "Bộ phận Kho Thành phẩm", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "LKVT", TenPhongBan = "Bộ phận Kho Vật tư", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "LBKH", TenPhongBan = "Bộ phận Khung", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "LBQA", TenPhongBan = "Bộ phận QA", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "LBQC", TenPhongBan = "Bộ phận QC", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "LRnD", TenPhongBan = "Bộ phận R&D - Giá thành", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "LBSO", TenPhongBan = "Bộ phận Sơn", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "LXNK", TenPhongBan = "Bộ phận XNK", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "LBDA", TenPhongBan = "Bộ Phận Đan", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "LBDG", TenPhongBan = "Bộ Phận Đóng gói", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "LGCK", TenPhongBan = "Phòng Gia Công Khung", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "LGCD", TenPhongBan = "Phòng Gia Công Đan, Sơn", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "LPHC", TenPhongBan = "Phòng Nhân Sự Hành Chính", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "LPIT", TenPhongBan = "Phòng IT", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "LPKD", TenPhongBan = "Phòng Kinh doanh - Marketing", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "LPKH", TenPhongBan = "Phòng Kế hoạch", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "LPKT", TenPhongBan = "Phòng Kỹ thuật", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "LPMH", TenPhongBan = "Phòng Mua hàng", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "LPTC", TenPhongBan = "Phòng TCKT", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "LTXE", TenPhongBan = "Tài xế", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "LTCM", TenPhongBan = "Tổ Cement, gỗ", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "LTDA", TenPhongBan = "Tổ chạy dây - bắn tấm", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "LTMA", TenPhongBan = "Tổ mẫu, cải tiến", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "TBGD", TenPhongBan = "Ban Giám Đốc", TenCongTy = "Công Ty TTP" },
                new PhongBan { MaPhongBan = "TBBS", TenPhongBan = "Bộ phận Bán sứ", TenCongTy = "Công Ty TTP" },
                new PhongBan { MaPhongBan = "TBKH", TenPhongBan = "Bộ phận Kho TVH", TenCongTy = "Công Ty TTP" },
                new PhongBan { MaPhongBan = "TBLB", TenPhongBan = "Bộ phận Lò Bao", TenCongTy = "Công Ty TTP" },
                new PhongBan { MaPhongBan = "TBLG", TenPhongBan = "Bộ phận Lò Gas", TenCongTy = "Công Ty TTP" },
                new PhongBan { MaPhongBan = "TBQC", TenPhongBan = "Bộ phận QA-QC", TenCongTy = "Công Ty TTP" },
                new PhongBan { MaPhongBan = "TRnD", TenPhongBan = "Bộ phận R&D - Giá thành", TenCongTy = "Công Ty TTP" },
                new PhongBan { MaPhongBan = "TXNK", TenPhongBan = "Bộ phận XNK", TenCongTy = "Công Ty TTP" },
                new PhongBan { MaPhongBan = "TPGC", TenPhongBan = "Phòng Gia Công", TenCongTy = "Công Ty TTP" },
                new PhongBan { MaPhongBan = "TPHC", TenPhongBan = "Phòng HCNS", TenCongTy = "Công Ty TTP" },
                new PhongBan { MaPhongBan = "TPIT", TenPhongBan = "Phòng IT", TenCongTy = "Công Ty TTP" },
                new PhongBan { MaPhongBan = "TPKD", TenPhongBan = "Phòng Kinh doanh - Marketing", TenCongTy = "Công Ty TTP" },
                new PhongBan { MaPhongBan = "TPKH", TenPhongBan = "Phòng Kế hoạch", TenCongTy = "Công Ty TTP" },
                new PhongBan { MaPhongBan = "TPKT", TenPhongBan = "Phòng Kỹ thuật", TenCongTy = "Công Ty TTP" },
                new PhongBan { MaPhongBan = "TPMH", TenPhongBan = "Phòng Mua hàng", TenCongTy = "Công Ty TTP" },
                new PhongBan { MaPhongBan = "TPTC", TenPhongBan = "Phòng TCKT", TenCongTy = "Công Ty TTP" },
                new PhongBan { MaPhongBan = "TTXE", TenPhongBan = "Bộ phận Sản Xuất", TenCongTy = "Công Ty TTP" },
                new PhongBan { MaPhongBan = "KSNB", TenPhongBan = "Kiểm Soát Nội Bộ", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "BaoTri", TenPhongBan = "Bảo Trì", TenCongTy = "Công Ty LHH" },
                new PhongBan { MaPhongBan = "CoKhi", TenPhongBan = "Cơ Khí", TenCongTy = "Công Ty LHH" }

            };

            foreach (var phongBan in phongBans)
            {
                if (!await context.PhongBans.AnyAsync(p => p.MaPhongBan == phongBan.MaPhongBan))
                {
                    context.PhongBans.Add(phongBan);
                }
            }


            // 4. Tạo tài khoản Admin mặc định nếu chưa có
            var adminEmail = "admin@quanly.com";
            var admin = await userManager.FindByEmailAsync(adminEmail);
            if (admin == null)
            {
                var newAdmin = new NguoiDung
                {
                    UserName = adminEmail,
                    Email = adminEmail,
                    EmailConfirmed = true,
                    HoTen = "Admin Hệ Thống",
                    TenCongTy = "LHH",
                    MaNV = "L22180",
                    ChucVu = "Admin",
                    MaPhongBan = "LPIT"
                };

                var result = await userManager.CreateAsync(newAdmin, "Admin@123");
                if (result.Succeeded)
                {
                    await userManager.AddToRoleAsync(newAdmin, "Admin");
                }
            }
        }
    }
}