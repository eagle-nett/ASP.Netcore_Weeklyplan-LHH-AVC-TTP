namespace QuanLy.Models

{
    public class XoaTaiKhoanViewModel
    {
        public string Id { get; set; }
        public string Email { get; set; }
        public string HoTen { get; set; }

        // Thuộc tính để bind từ form nhập mã xác thực
        public string SecurityCode { get; set; }

        // Dùng để thông báo có cần mã bảo mật không
        public bool IsAdmin { get; set; }
    }
}
