using System.ComponentModel.DataAnnotations;

namespace QuanLy.ViewModels
{
    public class NoiDungBaoCaoViewModel
    {
        public int? Id { get; set; } // null nếu đang thêm mới

        [Required(ErrorMessage = "Vui lòng nhập nội dung báo cáo")]
        public string NoiDung { get; set; } = "";

        public DateTime? NgayHoanThanh { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập mảng công việc")]
        public string TrachNhiemChinh { get; set; } = "";

        public string? TrachNhiemHoTro { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập mục tiêu")]
        public string MucDoUuTien { get; set; } = ""; // A/B/C

        [Required(ErrorMessage = "Vui lòng chọn tiến độ")]
        public string TienDo { get; set; } = ""; // "Đã hoàn thành" / "Chưa hoàn thành"

        public string? GhiChu { get; set; }

        // ✅ Trường mới
        public string? LyDoChuaHoanThanh { get; set; }
        public string? HuongGiaiQuyet { get; set; }
        public string? KetQuaDatDuoc { get; set; }
    }
}
