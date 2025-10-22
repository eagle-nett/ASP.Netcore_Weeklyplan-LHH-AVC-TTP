using System.ComponentModel.DataAnnotations;

namespace QuanLy.Models
{
    public class NoiDungBaoCao
    {
        public int Id { get; set; }

        public int BaoCaoTuanId { get; set; }
        public BaoCaoTuan BaoCaoTuan { get; set; }

        [Required]
        public string NoiDung { get; set; } = "";

        public DateTime? NgayHoanThanh { get; set; }

        [Required]
        public string TrachNhiemChinh { get; set; } = "";

        public string? TrachNhiemHoTro { get; set; }

        public string MucDoUuTien { get; set; } = ""; // A, B, C

        [Required]
        public string TienDo { get; set; } = ""; // "Đã hoàn thành" hoặc "Chưa hoàn thành"

        public string? GhiChu { get; set; }

        // ✅ Trường mới
        public string? LyDoChuaHoanThanh { get; set; }

        public string? HuongGiaiQuyet { get; set; }

        public string? KetQuaDatDuoc { get; set; }

    }
}
