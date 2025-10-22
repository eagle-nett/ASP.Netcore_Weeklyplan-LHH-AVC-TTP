// ViewModels/TrangThaiViewModel.cs
using System;

namespace QuanLy.ViewModels
{
    public class TrangThaiViewModel
    {
        public int STT { get; set; }
        public string Tuan { get; set; } = "";
        public DateTime TuNgay { get; set; }
        public DateTime DenNgay { get; set; }

        public string MaPhongBan { get; set; } = "";
        public string TenPhongBan { get; set; } = "";

        public string MaNV { get; set; } = "";
        public string TenNV { get; set; } = "";

        public string TenQuanLy { get; set; } = ""; // tên quản lý trực tiếp

        public DateTime? ThoiGianGui { get; set; } // thời gian gửi (NgayTao)
        public bool DaGui { get; set; } = false;
    }
}
