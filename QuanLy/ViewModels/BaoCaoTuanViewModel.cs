using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Mvc.Rendering;

namespace QuanLy.ViewModels
{
    public class BaoCaoTuanViewModel
    {
        // Người gửi - từ tài khoản đang đăng nhập
        public int? Id { get; set; } //dùng cho edit baocao
        public string NguoiBaoCaoId { get; set; } = "";
        public string HoTenNguoiBaoCao { get; set; } = ""; // dùng để hiển thị

        // Chọn người nhận
        [Required(ErrorMessage = "Vui lòng chọn người nhận báo cáo")]
        public string BaoCaoChoId { get; set; } = "";

        // Dropdown người nhận sau khi chọn mã phòng ban
        public List<SelectListItem> DanhSachNguoiNhan { get; set; } = new();

        // Mã phòng ban để lọc người nhận
        [Required(ErrorMessage = "Vui lòng chọn mã phòng ban")]
        public string MaPhongBan { get; set; } = "";

        public List<SelectListItem> DanhSachPhongBan { get; set; } = new();

        // Chọn tuần
        [Required(ErrorMessage = "Vui lòng chọn tuần báo cáo")]
        public string Tuan { get; set; } = "";
        public List<SelectListItem> TuanOptions { get; set; } = new(); // <-- thêm

        public DateTime TuNgay { get; set; }
        public DateTime DenNgay { get; set; }

        // Danh sách dòng báo cáo
        public List<NoiDungBaoCaoViewModel> NoiDungs { get; set; } = new();

        public DateTime NgayTao { get; set; } = DateTime.Now;
    }
}
