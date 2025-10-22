using Microsoft.AspNetCore.Identity;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Microsoft.AspNetCore.Mvc.ModelBinding.Validation;


namespace QuanLy.Models
{
    public class NguoiDung : IdentityUser
    {
        [Required]
        public string MaNV { get; set; }    // Mã NV

        [Required]
        public string HoTen { get; set; }   // Tên NV

        public string? DienThoai { get; set; } // Số điện thoại (nếu có)

        [Required]
        public string ChucVu { get; set; }  // Giám đốc / Trưởng phòng / Nhân viên

        public string? ChucDanh { get; set; } // Chức danh (nullable)

        [Required]
        public string TenCongTy { get; set; } // LHH hoặc TTP

        public string? GhiChu { get; set; }  // Ghi chú (nullable)

        [Required]
        public string MaPhongBan { get; set; } // Mã phòng ban

        [ForeignKey("MaPhongBan")]
        [ValidateNever]
        public virtual PhongBan PhongBan { get; set; } // Navigation property

    }
}
