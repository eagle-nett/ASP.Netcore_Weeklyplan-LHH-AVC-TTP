using System.ComponentModel.DataAnnotations;

namespace QuanLy.Models
{
    public class BaoCaoTuan
    {
        public int Id { get; set; }

        [Required]
        public string NguoiBaoCaoId { get; set; } = string.Empty; // Người đang đăng nhập
        public NguoiDung? NguoiBaoCao { get; set; }

        [Required]
        public string BaoCaoChoId { get; set; } = string.Empty; // Người nhận báo cáo
        public NguoiDung? BaoCaoCho { get; set; }

        [Required]
        public string Tuan { get; set; } = ""; // ví dụ: Y25W28

        public DateTime TuNgay { get; set; }
        public DateTime DenNgay { get; set; }

        public ICollection<NoiDungBaoCao> NoiDungs { get; set; } = new List<NoiDungBaoCao>();

        public DateTime NgayTao { get; set; } = DateTime.Now;
        public string TrangThai { get; set; } = ""; // VD: "Chờ duyệt", "Đã duyệt", "Bị từ chối"
        public string GhiChuCuaCapTren { get; set; } = ""; // Nếu cấp trên ghi chú gì

    }
}
