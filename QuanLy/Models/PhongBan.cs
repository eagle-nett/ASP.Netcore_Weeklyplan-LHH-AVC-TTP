using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace QuanLy.Models
{
    public class PhongBan
    {
        public PhongBan()
        {
            NguoiDungs = new HashSet<NguoiDung>();
        }

        [Key]
        public string MaPhongBan { get; set; }

        [Required]
        public string TenPhongBan { get; set; }

        [Required]
        public string TenCongTy { get; set; }

        public virtual ICollection<NguoiDung> NguoiDungs { get; set; }
    }

}
