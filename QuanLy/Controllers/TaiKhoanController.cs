using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using QuanLy.Data;
using QuanLy.Models;
using System.Linq;
using System.Threading.Tasks;

namespace QuanLy.Controllers
{
    public class TaiKhoanController : Controller
    {
        private readonly SignInManager<NguoiDung> _signInManager;
        private readonly UserManager<NguoiDung> _userManager;
        private readonly ApplicationDbContext _context;
        private readonly IConfiguration _configuration; 

        public TaiKhoanController(SignInManager<NguoiDung> signInManager, UserManager<NguoiDung> userManager, ApplicationDbContext context, IConfiguration configuration)
        {
            _signInManager = signInManager;
            _userManager = userManager;
            _context = context;
            _configuration = configuration;
        }

        // GET: /TaiKhoan/DangNhap
        [HttpGet]
        public IActionResult DangNhap(string returnUrl = null)
        {
            ViewData["ReturnUrl"] = returnUrl;
            return View();
        }

        // POST: /TaiKhoan/DangNhap
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DangNhap(string email, string password, bool rememberMe, string returnUrl = null)
        {
            if (!ModelState.IsValid)
            {
                return View();
            }

            var result = await _signInManager.PasswordSignInAsync(email, password, rememberMe, lockoutOnFailure: false);
            if (result.Succeeded)
            {
                return RedirectToLocal(returnUrl);
            }

            ModelState.AddModelError(string.Empty, "Email hoặc mật khẩu không đúng.");
            return View();
        }

        // POST: /TaiKhoan/DangXuat
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DangXuat()
        {
            await _signInManager.SignOutAsync();
            return RedirectToAction("Index", "Home");
        }

        // GET: /TaiKhoan/DangKy
        [HttpGet]
        [Authorize(Roles = "Admin, TruongPhong, GiamDoc")]
        public async Task<IActionResult> DangKy()
        {
            ViewBag.PhongBans = await _context.PhongBans.ToListAsync();
            ViewBag.CongTys = new[] { "LHH", "TTP" };
            ViewBag.Roles = new[] { "GiamDoc", "TruongPhong", "NhanVien" };

            return View();
        }

        // POST: /TaiKhoan/DangKy
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Authorize(Roles = "Admin, TruongPhong, GiamDoc")]
        public async Task<IActionResult> DangKy(NguoiDung model, string password, string role)
        {
            if (string.IsNullOrEmpty(password))
            {
                ModelState.AddModelError("password", "Mật khẩu không được để trống.");
            }

            if (string.IsNullOrEmpty(role))
            {
                ModelState.AddModelError("role", "Phải chọn phân quyền cho tài khoản.");
            }

            if (!ModelState.IsValid)
            {
                ViewBag.PhongBans = await _context.PhongBans.ToListAsync();
                ViewBag.CongTys = new[] { "LHH", "TTP" };
                ViewBag.Roles = new[] { "GiamDoc", "TruongPhong", "NhanVien" };

                return View(model);
            }

            var existingUser = await _userManager.FindByNameAsync(model.UserName);
            if (existingUser != null)
            {
                ModelState.AddModelError("UserName", "Tài khoản đã tồn tại. Vui lòng chọn tài khoản khác.");

                ViewBag.PhongBans = await _context.PhongBans.ToListAsync();
                ViewBag.CongTys = new[] { "LHH", "TTP" };
                ViewBag.Roles = new[] { "GiamDoc", "TruongPhong", "NhanVien" };

                return View(model);
            }

            var user = new NguoiDung
            {
                UserName = model.UserName,
                Email = model.UserName + "@fakeemail.com",
                MaNV = model.MaNV,
                HoTen = model.HoTen,
                DienThoai = model.DienThoai,
                ChucVu = model.ChucVu,
                ChucDanh = model.ChucDanh,
                TenCongTy = model.TenCongTy,
                GhiChu = model.GhiChu,
                MaPhongBan = model.MaPhongBan,
                EmailConfirmed = true
            };

            var result = await _userManager.CreateAsync(user, password);
            if (result.Succeeded)
            {
                await _userManager.AddToRoleAsync(user, role);
                TempData["Success"] = "Tạo tài khoản thành công!";
                return RedirectToAction("DanhSach");
            }

            foreach (var error in result.Errors)
            {
                ModelState.AddModelError(string.Empty, error.Description);
            }

            ViewBag.PhongBans = await _context.PhongBans.ToListAsync();
            ViewBag.CongTys = new[] { "LHH", "TTP" };
            ViewBag.Roles = new[] { "GiamDoc", "TruongPhong", "NhanVien" };

            return View(model);
        }


        // GET: /TaiKhoan/DanhSach
        [HttpGet]
        [Authorize(Roles = "Admin, TruongPhong, GiamDoc")]
        public async Task<IActionResult> DanhSach()
        {
            var currentUser = User;
            var isAdmin = currentUser.IsInRole("Admin");

            var users = await _userManager.Users
                .Include(u => u.PhongBan)
                .ToListAsync();

            var userList = new List<dynamic>();

            foreach (var u in users)
            {
                var roles = await _userManager.GetRolesAsync(u);

                // Nếu không phải Admin và user có role Admin → bỏ qua
                if (!isAdmin && roles.Contains("Admin"))
                    continue;

                userList.Add(new
                {
                    u.Id,
                    u.MaNV,
                    u.HoTen,
                    u.DienThoai,
                    u.ChucVu,
                    u.ChucDanh,
                    u.UserName,
                    u.TenCongTy,
                    u.GhiChu,
                    u.MaPhongBan,
                    TenPhongBan = u.PhongBan != null ? u.PhongBan.TenPhongBan : "Chưa gán",
                    Roles = string.Join(", ", roles)
                });
            }

            ViewBag.Users = userList;
            return View();
        }

        // GET: /TaiKhoan/TaoAdmin
        [HttpGet]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> TaoAdmin()
        {
            ViewBag.PhongBans = await _context.PhongBans.ToListAsync();
            ViewBag.CongTys = new[] { "LHH", "TTP" };
            return View();
        }

        // POST: /TaiKhoan/TaoAdmin
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> TaoAdmin(NguoiDung model, string password)
        {
            if (string.IsNullOrEmpty(password))
                ModelState.AddModelError("password", "Mật khẩu không được để trống.");

            if (!ModelState.IsValid)
            {
                ViewBag.PhongBans = await _context.PhongBans.ToListAsync();
                return View(model);
            }

            var existingUser = await _userManager.FindByNameAsync(model.UserName);
            if (existingUser != null)
            {
                ModelState.AddModelError("UserName", "Tài khoản đã tồn tại.");
                ViewBag.PhongBans = await _context.PhongBans.ToListAsync();
                return View(model);
            }

            var user = new NguoiDung
            {
                UserName = model.UserName,
                Email = model.UserName + "@gmail.com",
                HoTen = model.HoTen,
                MaNV = model.MaNV,
                MaPhongBan = model.MaPhongBan,
                EmailConfirmed = true,
                ChucVu = "Admin",
                ChucDanh = "Quản trị hệ thống",
                TenCongTy = "LHH", // hoặc mặc định công ty quản trị
                GhiChu = "Tài khoản admin được tạo bởi hệ thống"
            };

            var result = await _userManager.CreateAsync(user, password);
            if (result.Succeeded)
            {
                await _userManager.AddToRoleAsync(user, "Admin");
                TempData["Success"] = "Tạo tài khoản admin thành công.";
                return RedirectToAction("DanhSach");
            }

            foreach (var error in result.Errors)
            {
                ModelState.AddModelError(string.Empty, error.Description);
            }

            ViewBag.PhongBans = await _context.PhongBans.ToListAsync();
            return View(model);
        }

        // GET: /TaiKhoan/ChinhSua/{id}
        [HttpGet]
        [Authorize(Roles = "Admin, TruongPhong, GiamDoc")]
        public async Task<IActionResult> ChinhSua(string id)
        {
            var user = await _userManager.Users
                .Include(u => u.PhongBan)
                .FirstOrDefaultAsync(u => u.Id == id);

            if (user == null)
                return NotFound();

            var roles = await _userManager.GetRolesAsync(user);
            var role = roles.FirstOrDefault();

            // Nếu đang chỉnh tài khoản admin khác
            if (role == "Admin" && user.UserName != User.Identity.Name)
            {
                // Check xem có trong TempData mật khẩu xác thực chưa
                if (TempData["AdminAuthPassed"]?.ToString() != "true")
                {
                    TempData["EditingAdminId"] = id;
                    return RedirectToAction("XacThucAdmin");
                }
            }

            ViewBag.PhongBans = await _context.PhongBans.ToListAsync();
            ViewBag.CongTys = new[] { "LHH", "TTP" };
            ViewBag.Roles = new[] { "GiamDoc", "TruongPhong", "NhanVien" };

            var model = new NguoiDung
            {
                Id = user.Id,
                MaNV = user.MaNV,
                HoTen = user.HoTen,
                DienThoai = user.DienThoai,
                ChucVu = user.ChucVu,
                ChucDanh = user.ChucDanh,
                UserName = user.UserName,
                TenCongTy = user.TenCongTy,
                GhiChu = user.GhiChu,
                MaPhongBan = user.MaPhongBan
            };

            ViewBag.CurrentRole = role;

            return View(model);
        }


        // POST: /TaiKhoan/ChinhSua/{id}
        [HttpPost]
        [ActionName("ChinhSuaPost")]
        [ValidateAntiForgeryToken]
        [Authorize(Roles = "Admin, TruongPhong, GiamDoc")]
        public async Task<IActionResult> ChinhSua(string id, NguoiDung model, string password, string role)
        {
            // Xóa lỗi password nếu có
            ModelState.Remove("password");

            if (string.IsNullOrEmpty(role))
            {
                ModelState.AddModelError("role", "Phải chọn phân quyền cho tài khoản.");
            }

            if (!ModelState.IsValid)
            {
                ViewBag.PhongBans = await _context.PhongBans.ToListAsync();
                ViewBag.CongTys = new[] { "LHH", "TTP" };
                ViewBag.Roles = new[] { "GiamDoc", "TruongPhong", "NhanVien" };
                ViewBag.CurrentRole = role;

                return View("ChinhSua", model);
            }

            var user = await _userManager.FindByIdAsync(id);
            if (user == null)
            {
                return NotFound();
            }

            // Nếu đổi UserName → check trùng
            if (user.UserName != model.UserName)
            {
                var existingUser = await _userManager.FindByNameAsync(model.UserName);
                if (existingUser != null)
                {
                    ModelState.AddModelError("UserName", "Tài khoản đã tồn tại. Vui lòng chọn tài khoản khác.");

                    ViewBag.PhongBans = await _context.PhongBans.ToListAsync();
                    ViewBag.CongTys = new[] { "LHH", "TTP" };
                    ViewBag.Roles = new[] { "GiamDoc", "TruongPhong", "NhanVien" };
                    ViewBag.CurrentRole = role;

                    return View("ChinhSua", model);
                }
            }

            user.UserName = model.UserName;
            user.Email = model.UserName + "@fakeemail.com";
            user.MaNV = model.MaNV;
            user.HoTen = model.HoTen;
            user.DienThoai = model.DienThoai;
            user.ChucVu = model.ChucVu;
            user.ChucDanh = model.ChucDanh ?? "";
            user.TenCongTy = model.TenCongTy;
            user.GhiChu = model.GhiChu ?? "";
            user.MaPhongBan = model.MaPhongBan;

            var result = await _userManager.UpdateAsync(user);
            if (!result.Succeeded)
            {
                foreach (var error in result.Errors)
                {
                    ModelState.AddModelError(string.Empty, error.Description);
                }

                ViewBag.PhongBans = await _context.PhongBans.ToListAsync();
                ViewBag.CongTys = new[] { "LHH", "TTP" };
                ViewBag.Roles = new[] { "GiamDoc", "TruongPhong", "NhanVien" };
                ViewBag.CurrentRole = role;

                return View("ChinhSua", model);
            }

            // Xóa role cũ → gán role mới
            var currentRoles = await _userManager.GetRolesAsync(user);
            if (currentRoles.Any())
            {
                await _userManager.RemoveFromRolesAsync(user, currentRoles);
            }
            await _userManager.AddToRoleAsync(user, role);

            // Nếu đổi mật khẩu
            if (!string.IsNullOrEmpty(password))
            {
                var token = await _userManager.GeneratePasswordResetTokenAsync(user);
                var pwdResult = await _userManager.ResetPasswordAsync(user, token, password);
                if (!pwdResult.Succeeded)
                {
                    foreach (var error in pwdResult.Errors)
                    {
                        ModelState.AddModelError(string.Empty, error.Description);
                    }

                    ViewBag.PhongBans = await _context.PhongBans.ToListAsync();
                    ViewBag.CongTys = new[] { "LHH", "TTP" };
                    ViewBag.Roles = new[] { "GiamDoc", "TruongPhong", "NhanVien" };
                    ViewBag.CurrentRole = role;

                    return View("ChinhSua", model);
                }
            }

            TempData["Success"] = "Cập nhật người dùng thành công.";
            return RedirectToAction("DanhSach");
        }

        // GET: /TaiKhoan/XacthucAdmin/
        [HttpGet]
        [Authorize(Roles = "Admin")]
        public IActionResult XacThucAdmin()
        {
            return View("XacThucAdmin");
        }

        // POST: /TaiKhoan/XacthucAdmin/
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Authorize(Roles = "Admin")]
        public IActionResult XacThucAdmin(string password)
        {
            if (password == "6868")
            {
                TempData["AdminAuthPassed"] = "true";
                var id = TempData["EditingAdminId"]?.ToString();
                return RedirectToAction("ChinhSua", new { id });
            }

            ModelState.AddModelError("", "Mật khẩu xác minh không đúng.");
            return View("XacThucAdmin");
        }

        // GET: /TaiKhoan/Xoa/{id}
        [HttpGet]
        [Authorize(Roles = "Admin, TruongPhong, GiamDoc")]
        public async Task<IActionResult> Xoa(string id)
        {
            var user = await _userManager.FindByIdAsync(id);
            if (user == null)
            {
                return NotFound();
            }

            var roles = await _userManager.GetRolesAsync(user);
            bool isAdmin = roles.Contains("Admin");

            var model = new XoaTaiKhoanViewModel
            {
                Id = user.Id,
                Email = user.Email,
                HoTen = user.HoTen,
                IsAdmin = isAdmin
            };

            return View(model);
        }


        // POST: /TaiKhoan/Xoa/{id}
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Authorize(Roles = "Admin, TruongPhong, GiamDoc")]
        public async Task<IActionResult> Xoa(string id, XoaTaiKhoanViewModel viewModel)
        {
            var user = await _userManager.FindByIdAsync(id);
            if (user == null)
            {
                return NotFound();
            }

            var roles = await _userManager.GetRolesAsync(user);
            bool isAdmin = roles.Contains("Admin");

            if (isAdmin)
            {
                // Tìm tất cả Admins
                var allAdmins = await _userManager.GetUsersInRoleAsync("Admin");

                if (allAdmins.Count <= 1)
                {
                    TempData["Error"] = "Phải còn ít nhất 1 Admin trong hệ thống. Không thể xóa Admin cuối cùng!";
                    return RedirectToAction("DanhSach");
                }

                // BẮT NHẬP MÃ BẢO MẬT NẾU LÀ ADMIN
                //if (string.IsNullOrEmpty(model.SecurityCode) || model.SecurityCode != "5202 <-")
                var securityCode = _configuration["SecuritySettings:AdminDeleteCode"];
                if (viewModel.SecurityCode != securityCode)
                {
                    ModelState.AddModelError("SecurityCode", "Mã bảo mật không đúng. Vui lòng nhập đúng để xóa Admin.");

                    // Load lại view model để hiển thị lại view
                    viewModel.Email = user.Email;
                    viewModel.HoTen = user.HoTen;
                    viewModel.IsAdmin = true;

                    return View(viewModel);
                }
            }

            var result = await _userManager.DeleteAsync(user);
            if (result.Succeeded)
            {
                TempData["Success"] = "Xóa tài khoản thành công.";
                return RedirectToAction("DanhSach");
            }

            foreach (var error in result.Errors)
            {
                ModelState.AddModelError(string.Empty, error.Description);
            }

            // Trường hợp lỗi, trả về View kèm dữ liệu
            viewModel.Email = user.Email;
            viewModel.HoTen = user.HoTen;
            viewModel.IsAdmin = isAdmin;

            return View(viewModel);
        }

        // GET: /TaiKhoan/DoiMatKhau
        [HttpGet]
        [Authorize]
        public IActionResult DoiMatKhau()
        {
            return View();
        }

        // POST: /TaiKhoan/DoiMatKhau
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Authorize]
        public async Task<IActionResult> DoiMatKhau(ChangePasswordViewModel model)
        {
            if (!ModelState.IsValid)
            {
                return View(model);
            }

            var user = await _userManager.GetUserAsync(User);
            if (user == null)
            {
                return RedirectToAction("DangNhap");
            }

            var result = await _userManager.ChangePasswordAsync(user, model.CurrentPassword, model.NewPassword);

            if (result.Succeeded)
            {
                TempData["Success"] = "Đổi mật khẩu thành công!";
                await _signInManager.RefreshSignInAsync(user); // cập nhật lại session
                return RedirectToAction(nameof(DoiMatKhau)); // quay lại form và hiện thông báo
            }

            // hiển thị tất cả lỗi
            foreach (var error in result.Errors)
            {
                ModelState.AddModelError(string.Empty, error.Description);
            }

            return View(model);
        }

        // GET: /TaiKhoan/ResetPassword/{id}
        [HttpGet]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> ResetPassword(string id)
        {
            if (string.IsNullOrEmpty(id)) return NotFound();

            var user = await _userManager.FindByIdAsync(id);
            if (user == null) return NotFound();

            var model = new ResetPasswordViewModel
            {
                UserId = user.Id
            };
            return View(model);
        }

        // POST: /TaiKhoan/ResetPassword
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> ResetPassword(ResetPasswordViewModel model)
        {
            if (!ModelState.IsValid)
            {
                return View(model);
            }

            var user = await _userManager.FindByIdAsync(model.UserId);
            if (user == null)
            {
                return NotFound();
            }

            // Xóa mật khẩu cũ (nếu có) rồi set mật khẩu mới
            var removeResult = await _userManager.RemovePasswordAsync(user);
            if (!removeResult.Succeeded)
            {
                foreach (var error in removeResult.Errors)
                {
                    ModelState.AddModelError(string.Empty, error.Description);
                }
                return View(model);
            }

            var addResult = await _userManager.AddPasswordAsync(user, model.NewPassword);
            if (addResult.Succeeded)
            {
                TempData["Success"] = $"Reset mật khẩu cho {user.UserName} thành công!";
                return RedirectToAction("DanhSach");
            }

            foreach (var error in addResult.Errors)
            {
                ModelState.AddModelError(string.Empty, error.Description);
            }

            return View(model);
        }



        private IActionResult RedirectToLocal(string returnUrl)
        {
            if (Url.IsLocalUrl(returnUrl))
            {
                return Redirect(returnUrl);
            }
            return RedirectToAction("Index", "Home");
        }
    }
}