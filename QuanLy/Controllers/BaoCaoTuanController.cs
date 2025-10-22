using iTextSharp.text; // NuGet: iTextSharp.LGPLv2.Core
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml; // NuGet: EPPlus
using OfficeOpenXml.Style;
using QuanLy.Data;
using QuanLy.Helpers;
using QuanLy.Models;
using QuanLy.ViewModels;
using System.Globalization;
using System.IO;
using Microsoft.AspNetCore.Authorization;
using ClosedXML.Excel;


namespace QuanLy.Controllers
{
    public class BaoCaoTuanController : Controller
    {
        private readonly ApplicationDbContext _context;
        private readonly UserManager<NguoiDung> _userManager;

        public BaoCaoTuanController(ApplicationDbContext context, UserManager<NguoiDung> userManager)
        {
            _context = context;
            _userManager = userManager;
        }

        // GET: /BaoCao
        public async Task<IActionResult> Index()
        {
            var user = await _userManager.GetUserAsync(User);
            var list = _context.BaoCaoTuans
                .Include(b => b.NguoiBaoCao)
                .Where(b => b.NguoiBaoCaoId == user.Id)
                .ToList();
            return View(list);
        }

        // GET: /BaoCao/Create?maPhongBan=...
        [HttpGet]
        public async Task<IActionResult> Tao(string? maPhongBan)
        {
            var user = await _userManager.GetUserAsync(User);
            if (user == null) return Challenge();

            // --- Khởi tạo ViewModel ---
            var vm = new BaoCaoTuanViewModel
            {
                NguoiBaoCaoId = user.Id,
                HoTenNguoiBaoCao = user.HoTen,
                MaPhongBan = maPhongBan ?? "",
                NoiDungs = new List<NoiDungBaoCaoViewModel> { new NoiDungBaoCaoViewModel() }
            };

            // --- Nạp danh sách dropdown (giống bên POST) ---
            ReloadDropdowns(vm, user.Id); // khúc này thay cho đoạn sau

            // --- Nếu có tuần → tự động tính Từ ngày / Đến ngày ---
            if (!string.IsNullOrEmpty(vm.Tuan))
            {
                var match = System.Text.RegularExpressions.Regex.Match(vm.Tuan, @"Y(?<year>\d{2})W(?<week>\d{1,2})");
                if (match.Success)
                {
                    int year = 2000 + int.Parse(match.Groups["year"].Value);
                    int week = int.Parse(match.Groups["week"].Value);

                    var jan1 = new DateTime(year, 1, 1);
                    int daysOffset = DayOfWeek.Thursday - jan1.DayOfWeek;
                    var firstThursday = jan1.AddDays(daysOffset);

                    var cal = CultureInfo.CurrentCulture.Calendar;
                    var firstWeek = cal.GetWeekOfYear(firstThursday, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
                    int delta = week - (firstWeek <= 1 ? 1 : 0);

                    var weekStart = firstThursday.AddDays(delta * 7).AddDays(-3);
                    if (weekStart.DayOfWeek != DayOfWeek.Monday)
                        weekStart = weekStart.AddDays((int)DayOfWeek.Monday - (int)weekStart.DayOfWeek);

                    vm.TuNgay = weekStart;
                    vm.DenNgay = weekStart.AddDays(6);
                }
            }

            return View(vm);
        }

        // POST: /BaoCao/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Tao(BaoCaoTuanViewModel vm)
        {
            var user = await _userManager.GetUserAsync(User);
            if (user == null) return Challenge();

            // 🛑 Bỏ qua validate nếu người dùng chỉ bấm Thêm hoặc Xóa dòng
            if (Request.Form.ContainsKey("addRow") || Request.Form.ContainsKey("removeRow"))
            {
                ModelState.Clear();
            }

            // --- Xử lý nút "Thêm dòng" ---
            if (Request.Form.ContainsKey("addRow"))
            {
                if (vm.NoiDungs == null) vm.NoiDungs = new List<NoiDungBaoCaoViewModel>();
                vm.NoiDungs.Add(new NoiDungBaoCaoViewModel());
                ReloadDropdowns(vm, user.Id); // 🔁 nạp lại dropdowns
                ModelState.Clear(); // <-- thêm dòng này để ko bị nhân đôi chữ abc, abc
                return View(vm);
            }

            // --- Xử lý nút "Xóa dòng" ---
            if (Request.Form.ContainsKey("removeRow"))
            {
                var value = Request.Form["removeRow"].ToString();
                if (int.TryParse(value, out int idx))
                {
                    if (vm.NoiDungs != null && idx >= 0 && idx < vm.NoiDungs.Count)
                        vm.NoiDungs.RemoveAt(idx);
                }
                if (vm.NoiDungs == null || vm.NoiDungs.Count == 0)
                    vm.NoiDungs = new List<NoiDungBaoCaoViewModel> { new NoiDungBaoCaoViewModel() };

                ReloadDropdowns(vm, user.Id); // 🔁 nạp lại dropdowns
                ModelState.Clear(); // <-- thêm dòng này để ko bị nhân đôi chữ abc, abc
                return View(vm);
            }

            // --- Khi người dùng nhấn "Gửi báo cáo" ---
            if (!ModelState.IsValid)
            {
                // 🔥 Xóa toàn bộ giá trị nhập cũ của danh sách nội dung để tránh nhân đôi
                foreach (var key in ModelState.Keys.Where(k => k.StartsWith("NoiDungs")))
                {
                    ModelState.Remove(key);
                }

                ReloadDropdowns(vm, user.Id);
                return View(vm);
            }


            // --- Kiểm tra trùng tuần ---
            bool daTonTai = await _context.BaoCaoTuans
                .AnyAsync(b => b.NguoiBaoCaoId == user.Id && b.Tuan == vm.Tuan);

            if (daTonTai)
            {
                ViewData["ThongBaoTrungTuan"] = "Bạn đã tạo báo cáo cho tuần này rồi.";
                ModelState.AddModelError("Tuan", "Bạn đã tạo báo cáo cho tuần này rồi.");
                ModelState.Clear(); // tránh lỗi trùng nội dung
                ReloadDropdowns(vm, user.Id);
                return View(vm);
            }

            // --- Map sang entity để lưu ---
            var entity = new BaoCaoTuan
            {
                NguoiBaoCaoId = user.Id,
                BaoCaoChoId = vm.BaoCaoChoId,
                Tuan = vm.Tuan,
                TuNgay = vm.TuNgay,
                DenNgay = vm.DenNgay,
                NgayTao = DateTime.Now
            };

            if (vm.NoiDungs != null)
            {
                foreach (var nd in vm.NoiDungs)
                {
                    if (string.IsNullOrWhiteSpace(nd.NoiDung) && string.IsNullOrWhiteSpace(nd.TrachNhiemChinh))
                        continue;

                    entity.NoiDungs.Add(new NoiDungBaoCao
                    {
                        NoiDung = nd.NoiDung,
                        NgayHoanThanh = nd.NgayHoanThanh,
                        TrachNhiemChinh = nd.TrachNhiemChinh,
                        TrachNhiemHoTro = nd.TrachNhiemHoTro,
                        MucDoUuTien = nd.MucDoUuTien,
                        TienDo = nd.TienDo,
                        GhiChu = nd.GhiChu,
                        LyDoChuaHoanThanh = nd.LyDoChuaHoanThanh,
                        KetQuaDatDuoc = nd.KetQuaDatDuoc,
                        HuongGiaiQuyet = nd.HuongGiaiQuyet
                    });
                }
            }

            _context.BaoCaoTuans.Add(entity);
            await _context.SaveChangesAsync();

            return RedirectToAction("Index", "Home");
        }

        // 💡 Hàm phụ: nạp lại dropdowns. Dùng userId (string) cho nhẹ và tránh phụ thuộc kiểu user
        private void ReloadDropdowns(BaoCaoTuanViewModel vm, string currentUserId)
        {
            // Danh sách phòng ban (text hiển thị gồm tên phòng + tên công ty)
            vm.DanhSachPhongBan = _context.PhongBans
                .AsNoTracking()
                .Select(pb => new SelectListItem(
                    $"{pb.TenPhongBan} ({pb.TenCongTy})",
                    pb.MaPhongBan
                ))
                .ToList();

            // Danh sách tuần
            int yearNow = DateTime.Now.Year;
            //vm.TuanOptions = GenerateWeekDropdown(yearNow - 1, yearNow + 1); // (2024->2026) vì (2025-1) -> (2025+1)
            vm.TuanOptions = GenerateWeekDropdown(yearNow, yearNow); //chỉ năm hiện tại
            //vm.TuanOptions = GenerateWeekDropdown(yearNow - 2, yearNow + 2);//5 năm gần nhất 


            // Nếu có MaPhongBan -> nạp danh sách người nhận (loại trừ chính user hiện tại)
            if (!string.IsNullOrEmpty(vm.MaPhongBan))
            {
                vm.DanhSachNguoiNhan = _context.Users
                    .Where(u => u.MaPhongBan == vm.MaPhongBan && u.Id != currentUserId)
                    .Select(u => new SelectListItem(u.HoTen + " – " + u.ChucVu, u.Id))
                    .ToList();
            }
        }


        // Helper: sinh dropdown tuần ISO (YxxWyy)
        private List<SelectListItem> GenerateWeekDropdown(int startYear, int endYear)
        {
            var items = new List<SelectListItem>();
            var cal = CultureInfo.CurrentCulture.Calendar;
            var weekRule = CalendarWeekRule.FirstFourDayWeek;
            var firstDay = DayOfWeek.Monday;

            for (int year = startYear; year <= endYear; year++)
            {
                // Tính số tuần trong năm
                int weeksInYear = cal.GetWeekOfYear(
                    new DateTime(year, 12, 31),
                    weekRule,
                    firstDay
                );

                for (int week = 1; week <= weeksInYear; week++)
                {
                    // Tìm ngày thứ Năm của tuần đầu tiên
                    var jan1 = new DateTime(year, 1, 1);
                    int daysOffset = DayOfWeek.Thursday - jan1.DayOfWeek;
                    var firstThursday = jan1.AddDays(daysOffset);
                    int firstWeek = cal.GetWeekOfYear(firstThursday, weekRule, firstDay);
                    int delta = week - (firstWeek <= 1 ? 1 : 0);

                    // Tính ngày thứ Hai của tuần hiện tại
                    var weekStart = firstThursday
                        .AddDays(delta * 7)
                        .AddDays(-3);
                    if (weekStart.DayOfWeek != DayOfWeek.Monday)
                        weekStart = weekStart.AddDays(
                            (int)DayOfWeek.Monday - (int)weekStart.DayOfWeek
                        );

                    string code = $"Y{year % 100:D2}W{week:D2}";
                    string label = $"{code} ({weekStart:dd/MM/yyyy} – {weekStart.AddDays(6):dd/MM/yyyy})";
                    items.Add(new SelectListItem(label, code));
                }
            }

            return items;
        }

        // GET: BaoCaoTuan/XemLai
        [HttpGet]
        public async Task<IActionResult> XemLai(string tuan, string nguoiNhan, string trangThai)
        {
            var nguoiDungId = _userManager.GetUserId(User);

            // Lấy danh sách tất cả báo cáo của người dùng hiện tại
            var query = _context.BaoCaoTuans
                .Where(b => b.NguoiBaoCaoId == nguoiDungId)
                .Include(b => b.BaoCaoCho)
                .Include(b => b.NoiDungs)
                .AsQueryable();

            // Lọc theo tuần nếu có
            if (!string.IsNullOrEmpty(tuan))
            {
                query = query.Where(b => b.Tuan == tuan);
            }

            // Lọc theo người nhận báo cáo
            if (!string.IsNullOrEmpty(nguoiNhan))
            {
                query = query.Where(b => b.BaoCaoChoId == nguoiNhan);
            }

            // Lọc theo trạng thái duyệt
            if (!string.IsNullOrEmpty(trangThai))
            {
                if (trangThai == "ChoDuyet")
                {
                    query = query.Where(b => b.TrangThai == null || b.TrangThai == "");
                }
                else
                {
                    query = query.Where(b => b.TrangThai == trangThai);
                }
            }

            var baoCaoList = await query
                .OrderByDescending(b => b.Tuan)
                .ToListAsync();

            // Truyền danh sách tuần có sẵn để lọc
            var tuanOptions = await _context.BaoCaoTuans
                .Where(b => b.NguoiBaoCaoId == nguoiDungId)
                .Select(b => b.Tuan)
                .Distinct()
                .OrderByDescending(t => t)
                .ToListAsync();

            ViewBag.TuanOptions = tuanOptions
                .Select(t => new SelectListItem { Value = t, Text = t })
                .ToList();

            // Truyền danh sách người nhận từng báo cáo để lọc
            var nguoiNhanOptions = await _context.BaoCaoTuans
                .Where(b => b.NguoiBaoCaoId == nguoiDungId)
                .Include(b => b.BaoCaoCho)
                .Select(b => new { b.BaoCaoChoId, b.BaoCaoCho.HoTen })
                .Distinct()
                .ToListAsync();

            ViewBag.DanhSachNguoiNhan = nguoiNhanOptions
                .Select(x => new SelectListItem
                {
                    Value = x.BaoCaoChoId,
                    Text = x.HoTen
                })
                .ToList();

            return View(baoCaoList);
        }

        // GET: BaoCaoTuan/ChiTiet
        [HttpGet]
        public async Task<IActionResult> ChiTiet(int id)
        {
            var baoCao = await _context.BaoCaoTuans
                .Include(b => b.NoiDungs)
                .Include(b => b.BaoCaoCho)
                .Include(b => b.NguoiBaoCao)
                .FirstOrDefaultAsync(b => b.Id == id);

            if (baoCao == null)
            {
                return NotFound();
            }

            return View(baoCao);
        }

        // xuất excel view ChiTiet & ThongKe
        [HttpGet("BaoCaoTuan/ExportExcelChiTiet/{id}")]
        public IActionResult ExportExcel(int id)
        {
            var baoCao = _context.BaoCaoTuans
                .Include(b => b.NoiDungs)
                .Include(b => b.BaoCaoCho)
                .Include(b => b.NguoiBaoCao)
                .FirstOrDefault(b => b.Id == id);

            if (baoCao == null)
                return NotFound();

            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Báo cáo tuần");

            int row = 1;

            // ==== PHẦN 1: LOGO + TIÊU ĐỀ ====
            // ✅ Thêm logo (đặt logo.png trong wwwroot/images/)
            var logoPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo.png");
            if (System.IO.File.Exists(logoPath))
            {
                var picture = ws.Drawings.AddPicture("Logo", new FileInfo(logoPath));
                picture.SetPosition(0, 0, 0, 0); // hàng 1, cột 1
                picture.SetSize(220, 40); // kích thước logo
            }

            // Tiêu đề
            ws.Cells[row, 1, row, 11].Merge = true;
            ws.Cells[row, 1].Value = $"BÁO CÁO TUẦN {baoCao.Tuan}";
            ws.Cells[row, 1].Style.Font.Size = 18;
            ws.Cells[row, 1].Style.Font.Bold = true;
            ws.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Row(row).Height = 35;
            row += 2;

            // ==== PHẦN 2: THÔNG TIN CHUNG ====
            ws.Cells[row, 1].Value = " Người báo cáo:";
            ws.Cells[row, 2].Value = $"{baoCao.NguoiBaoCao?.HoTen} ({baoCao.NguoiBaoCao?.MaNV})";
            ws.Cells[row, 4].Value = "Báo cáo cho:";
            ws.Cells[row, 5].Value = $"{baoCao.BaoCaoCho?.HoTen} ({baoCao.BaoCaoCho?.MaNV})";
            row++;

            ws.Cells[row, 1].Value = "Tuần:";
            ws.Cells[row, 2].Value = baoCao.Tuan;
            ws.Cells[row, 4].Value = "Thời gian:";
            ws.Cells[row, 5].Value = $"{baoCao.TuNgay:dd/MM/yyyy} - {baoCao.DenNgay:dd/MM/yyyy}";
            row++;

            ws.Cells[row, 1].Value = "📌 Trạng thái:";
            ws.Cells[row, 2].Value = baoCao.TrangThai switch
            {
                "DaDuyet" => "✔ Đã duyệt",
                "TuChoi" => "✖ Từ chối",
                _ => "⏳ Chờ duyệt"
            };
            row++;

            if (!string.IsNullOrWhiteSpace(baoCao.GhiChuCuaCapTren))
            {
                ws.Cells[row, 1].Value = "📝 Ghi chú của cấp trên:";
                ws.Cells[row, 2, row, 8].Merge = true;
                ws.Cells[row, 2].Value = baoCao.GhiChuCuaCapTren;
                ws.Cells[row, 2].Style.WrapText = true;
                row += 2;
            }
            else
            {
                row++;
            }

            // Định dạng phần thông tin chung
            ws.Cells[3, 1, row - 1, 8].Style.Font.Size = 12;
            ws.Cells[3, 1, row - 1, 8].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            // ==== PHẦN 3: NỘI DUNG CHI TIẾT ====
            ws.Cells[row, 1].Value = "STT";
            ws.Cells[row, 2].Value = "Mảng việc";
            ws.Cells[row, 3].Value = "Nội dung";
            ws.Cells[row, 4].Value = "Ngày hoàn thành";
            ws.Cells[row, 5].Value = "Mục tiêu";
            ws.Cells[row, 6].Value = "Tiến độ";
            ws.Cells[row, 7].Value = "Kết quả đạt được";
            ws.Cells[row, 8].Value = "Người hỗ trợ";
            ws.Cells[row, 9].Value = "Lý do chưa hoàn thành";
            ws.Cells[row, 10].Value = "Hướng giải quyết";
            ws.Cells[row, 11].Value = "Ghi chú";

            using (var headerRange = ws.Cells[row, 1, row, 11])
            {
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightSteelBlue);
                headerRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                headerRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                headerRange.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            }
            row++;

            int stt = 1;
            foreach (var nd in baoCao.NoiDungs)
            {
                ws.Cells[row, 1].Value = stt++;
                ws.Cells[row, 2].Value = nd.TrachNhiemChinh;
                ws.Cells[row, 3].Value = nd.NoiDung;
                ws.Cells[row, 4].Value = nd.NgayHoanThanh?.ToString("dd/MM/yyyy");
                ws.Cells[row, 5].Value = nd.MucDoUuTien;
                ws.Cells[row, 6].Value = nd.TienDo;
                ws.Cells[row, 7].Value = nd.KetQuaDatDuoc;
                ws.Cells[row, 8].Value = nd.TrachNhiemHoTro;
                ws.Cells[row, 9].Value = nd.LyDoChuaHoanThanh;
                ws.Cells[row, 10].Value = nd.HuongGiaiQuyet;
                ws.Cells[row, 11].Value = nd.GhiChu;

                using (var rowRange = ws.Cells[row, 1, row, 11])
                {
                    rowRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    rowRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    rowRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    rowRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    rowRange.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    rowRange.Style.WrapText = true;
                }

                row++;
            }

            // ==== PHẦN 4: TỐI ƯU GIAO DIỆN ====
            ws.Column(1).Width = 5;   // STT
            ws.Column(2).Width = 15;  // Mảng việc
            ws.Column(3).Width = 45;  // ✅ Nội dung (cột chính, rộng hơn)
            ws.Column(4).Width = 15;  // Ngày hoàn thành
            ws.Column(5).Width = 25;  // Mức độ ưu tiên
            ws.Column(6).Width = 20;  // Tiến độ
            ws.Column(7).Width = 25;  // Kết quả đạt được
            ws.Column(8).Width = 10;  // Người hỗ trợ
            ws.Column(9).Width = 25;  // Lý do chưa hoàn thành
            ws.Column(10).Width = 25; // Hướng giải quyết
            ws.Column(11).Width = 25; // Ghi chú

            ws.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Column(5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Column(6).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            // ==== PHẦN 5: XUẤT FILE ====
            var stream = new MemoryStream(package.GetAsByteArray());
            string fileName = $"BaoCaoTuan_{baoCao.Tuan}_{baoCao.NguoiBaoCao?.MaNV}.xlsx";
            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        // GET: /BaoCao/Edit/5
        [HttpGet]
        public async Task<IActionResult> Edit(int id, string? maPhongBan)
        {
            var user = await _userManager.GetUserAsync(User);
            if (user == null) return Challenge();

            // --- Tìm báo cáo theo ID ---
            var baoCao = await _context.BaoCaoTuans
                .Include(b => b.NoiDungs)
                .FirstOrDefaultAsync(b => b.Id == id && b.NguoiBaoCaoId == user.Id);

            if (baoCao == null) return NotFound();

            // --- Khởi tạo ViewModel từ entity ---
            var vm = new BaoCaoTuanViewModel
            {
                Id = baoCao.Id,
                NguoiBaoCaoId = baoCao.NguoiBaoCaoId,
                HoTenNguoiBaoCao = user.HoTen,
                BaoCaoChoId = baoCao.BaoCaoChoId,
                MaPhongBan = maPhongBan ?? user.MaPhongBan ?? "",
                Tuan = baoCao.Tuan,
                TuNgay = baoCao.TuNgay,
                DenNgay = baoCao.DenNgay,
                NgayTao = baoCao.NgayTao,
                NoiDungs = baoCao.NoiDungs.Select(nd => new NoiDungBaoCaoViewModel
                {
                    Id = nd.Id,
                    NoiDung = nd.NoiDung,
                    NgayHoanThanh = nd.NgayHoanThanh,
                    TrachNhiemChinh = nd.TrachNhiemChinh,
                    TrachNhiemHoTro = nd.TrachNhiemHoTro,
                    MucDoUuTien = nd.MucDoUuTien,
                    TienDo = nd.TienDo,
                    GhiChu = nd.GhiChu,
                    LyDoChuaHoanThanh = nd.LyDoChuaHoanThanh,
                    HuongGiaiQuyet = nd.HuongGiaiQuyet,
                    KetQuaDatDuoc = nd.KetQuaDatDuoc
                }).ToList()
            };

            // --- Nạp dropdowns ---
            ReloadDropdowns(vm, user.Id);

            return View(vm);
        }

        // POST: /BaoCao/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(BaoCaoTuanViewModel vm)
        {
            var user = await _userManager.GetUserAsync(User);
            if (user == null) return Challenge();

            // 🛑 Bỏ qua validate nếu người dùng chỉ bấm Thêm hoặc Xóa dòng
            if (Request.Form.ContainsKey("addRow") || Request.Form.ContainsKey("removeRow"))
            {
                ModelState.Clear();
            }

            // --- Xử lý nút "Thêm dòng" ---
            if (Request.Form.ContainsKey("addRow"))
            {
                if (vm.NoiDungs == null) vm.NoiDungs = new List<NoiDungBaoCaoViewModel>();
                vm.NoiDungs.Add(new NoiDungBaoCaoViewModel());
                ReloadDropdowns(vm, user.Id); // nạp lại dropdowns như bên Tao
                ModelState.Clear(); // tránh lỗi nhân đôi giá trị trong ModelState
                return View(vm);
            }

            // --- Xử lý nút "Xóa dòng" ---
            if (Request.Form.ContainsKey("removeRow"))
            {
                var value = Request.Form["removeRow"].ToString();
                if (int.TryParse(value, out int idx))
                {
                    if (vm.NoiDungs != null && idx >= 0 && idx < vm.NoiDungs.Count)
                        vm.NoiDungs.RemoveAt(idx);
                }

                if (vm.NoiDungs == null || vm.NoiDungs.Count == 0)
                    vm.NoiDungs = new List<NoiDungBaoCaoViewModel> { new NoiDungBaoCaoViewModel() };

                ReloadDropdowns(vm, user.Id); // nạp lại dropdowns
                ModelState.Clear();
                return View(vm);
            }

            // --- Khi nhấn Lưu chỉnh sửa ---
            if (!ModelState.IsValid)
            {
                // Tránh nhân đôi dữ liệu (abc → abc,abc)
                foreach (var key in ModelState.Keys.Where(k => k.StartsWith("NoiDungs")))
                    ModelState.Remove(key);

                ReloadDropdowns(vm, user.Id);
                return View(vm);
            }

            // --- Tìm lại báo cáo trong DB ---
            var entity = await _context.BaoCaoTuans
                .Include(b => b.NoiDungs)
                .FirstOrDefaultAsync(b => b.Id == vm.Id && b.NguoiBaoCaoId == user.Id);

            if (entity == null)
                return NotFound();

            // --- Kiểm tra trùng tuần ---
            bool daTonTai = await _context.BaoCaoTuans
                .AnyAsync(b => b.NguoiBaoCaoId == user.Id && b.Tuan == vm.Tuan && b.Id != vm.Id);

            if (daTonTai)
            {
                ViewData["ThongBaoTrungTuan"] = "Bạn đã có báo cáo cho tuần này rồi.";
                ModelState.AddModelError("Tuan", "Bạn đã có báo cáo cho tuần này rồi.");
                // Giữ hành vi giống Tao: nạp dropdowns rồi clear modelstate để tránh lỗi trùng nội dung
                ReloadDropdowns(vm, user.Id);
                ModelState.Clear();
                return View(vm);
            }

            // --- Cập nhật thông tin chính ---
            entity.BaoCaoChoId = vm.BaoCaoChoId;
            entity.Tuan = vm.Tuan;
            entity.TuNgay = vm.TuNgay;
            entity.DenNgay = vm.DenNgay;

            // --- Cập nhật danh sách nội dung ---
            entity.NoiDungs.Clear(); // xóa cũ, thêm mới lại từ ViewModel
            foreach (var nd in vm.NoiDungs)
            {
                if (string.IsNullOrWhiteSpace(nd.NoiDung) && string.IsNullOrWhiteSpace(nd.TrachNhiemChinh))
                    continue;

                entity.NoiDungs.Add(new NoiDungBaoCao
                {
                    NoiDung = nd.NoiDung,
                    NgayHoanThanh = nd.NgayHoanThanh,
                    TrachNhiemChinh = nd.TrachNhiemChinh,
                    TrachNhiemHoTro = nd.TrachNhiemHoTro,
                    MucDoUuTien = nd.MucDoUuTien,
                    TienDo = nd.TienDo,
                    GhiChu = nd.GhiChu,
                    LyDoChuaHoanThanh = nd.LyDoChuaHoanThanh,
                    HuongGiaiQuyet = nd.HuongGiaiQuyet,
                    KetQuaDatDuoc = nd.KetQuaDatDuoc
                });
            }

            _context.Update(entity);
            await _context.SaveChangesAsync();

            TempData["ThongBao"] = "Cập nhật báo cáo tuần thành công!";
            return RedirectToAction("XemLai");
        }


        // GET: /BaoCao/Delete/5
        [HttpGet]
        public async Task<IActionResult> Delete(int id)
        {
            var baoCao = await _context.BaoCaoTuans
                .Include(b => b.NguoiBaoCao)
                .FirstOrDefaultAsync(b => b.Id == id);

            if (baoCao == null)
                return NotFound();

            // Xác nhận xóa trước khi thực hiện
            var vm = new BaoCaoTuanViewModel
            {
                Id = baoCao.Id,
                HoTenNguoiBaoCao = baoCao.NguoiBaoCao?.HoTen ?? "",
                Tuan = baoCao.Tuan,
                TuNgay = baoCao.TuNgay,
                DenNgay = baoCao.DenNgay
            };

            return View("Delete", vm);
        }

        // POST: /BaoCao/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var baoCao = await _context.BaoCaoTuans
                .Include(b => b.NoiDungs)
                .FirstOrDefaultAsync(b => b.Id == id);

            if (baoCao == null)
                return NotFound();

            // Xóa hết nội dung trước để tránh lỗi ràng buộc
            _context.NoiDungBaoCaos.RemoveRange(baoCao.NoiDungs);
            _context.BaoCaoTuans.Remove(baoCao);
            await _context.SaveChangesAsync();

            return RedirectToAction("XemLai");
        }

        // GET: /BaoCaoTuan/ThongKe
        [HttpGet]
        public IActionResult ThongKe(string? tuan, string? maNhanVien)
        {
            // Lấy danh sách tuần
            int yearNow = DateTime.Now.Year;
            ViewBag.TuanOptions = GenerateWeekDropdown(yearNow, yearNow);

            // Gắn lại giá trị để view hiển thị lại
            ViewBag.SelectedTuan = tuan;
            ViewBag.MaNhanVien = maNhanVien;

            // 🔒 Chỉ truy vấn khi cả tuần và mã nhân viên đều có
            if (string.IsNullOrEmpty(tuan) || string.IsNullOrEmpty(maNhanVien))
            {
                // Trả về view với danh sách rỗng (chưa lọc)
                return View(new List<BaoCaoTuan>());
            }

            // Lấy dữ liệu báo cáo khi đã đủ điều kiện
            var ds = _context.BaoCaoTuans
                .Include(b => b.NguoiBaoCao)
                .Include(b => b.BaoCaoCho)
                .Include(b => b.NoiDungs)
                .Where(b => b.Tuan == tuan && b.NguoiBaoCao.MaNV == maNhanVien)
                .OrderBy(b => b.Tuan)
                .ThenBy(b => b.NguoiBaoCao.HoTen)
                .ToList();

            return View(ds);
        }

        // Excel ThongKe
        [HttpGet("BaoCaoTuan/ExportExcelThongKe")]
        public IActionResult ExportExcel(string tuan, string maNhanVien)
        {
            var query = _context.BaoCaoTuans
                .Include(b => b.NguoiBaoCao)
                .Include(b => b.BaoCaoCho)
                .Include(b => b.NoiDungs)
                .AsQueryable();

            if (!string.IsNullOrEmpty(tuan))
                query = query.Where(b => b.Tuan == tuan);

            if (!string.IsNullOrEmpty(maNhanVien))
                query = query.Where(b => b.NguoiBaoCao.MaNV == maNhanVien);

            var ds = query.OrderBy(b => b.Tuan).ThenBy(b => b.NguoiBaoCao.HoTen).ToList();

            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Báo cáo tuần");

            int row = 1;
            int stt = 1;

            foreach (var bc in ds)
            {
                // ==== PHẦN 1: TIÊU ĐỀ ====
                ws.Cells[row, 1, row, 11].Merge = true;
                ws.Cells[row, 1].Value = $"BÁO CÁO TUẦN {bc.Tuan}";
                ws.Cells[row, 1].Style.Font.Bold = true;
                ws.Cells[row, 1].Style.Font.Size = 16;
                ws.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                row += 2;

                // ==== PHẦN 2: THÔNG TIN CHUNG ====
                ws.Cells[row, 1].Value = "Người báo cáo:";
                ws.Cells[row, 2].Value = $"{bc.NguoiBaoCao?.HoTen} ({bc.NguoiBaoCao?.MaNV})";
                ws.Cells[row, 4].Value = "Báo cáo cho:";
                ws.Cells[row, 5].Value = $"{bc.BaoCaoCho?.HoTen} ({bc.BaoCaoCho?.MaNV})";
                row++;

                ws.Cells[row, 1].Value = "Tuần:";
                ws.Cells[row, 2].Value = bc.Tuan;
                ws.Cells[row, 4].Value = "Thời gian:";
                ws.Cells[row, 5].Value = $"{bc.TuNgay:dd/MM/yyyy} - {bc.DenNgay:dd/MM/yyyy}";
                row++;

                ws.Cells[row, 1].Value = "Trạng thái:";
                ws.Cells[row, 2].Value = bc.TrangThai switch
                {
                    "DaDuyet" => "✔ Đã duyệt",
                    "TuChoi" => "✖ Từ chối",
                    _ => "⏳ Chờ duyệt"
                };
                row++;

                if (!string.IsNullOrWhiteSpace(bc.GhiChuCuaCapTren))
                {
                    ws.Cells[row, 1].Value = "Ghi chú của cấp trên:";
                    ws.Cells[row, 2, row, 8].Merge = true;
                    ws.Cells[row, 2].Value = bc.GhiChuCuaCapTren;
                    ws.Cells[row, 2].Style.WrapText = true;
                    row += 2;
                }
                else row++;

                // ==== PHẦN 3: NỘI DUNG CHI TIẾT ====
                ws.Cells[row, 1].Value = "STT";
                ws.Cells[row, 2].Value = "Mảng việc";
                ws.Cells[row, 3].Value = "Nội dung";
                ws.Cells[row, 4].Value = "Ngày hoàn thành";
                ws.Cells[row, 5].Value = "Mục tiêu";
                ws.Cells[row, 6].Value = "Tiến độ";
                ws.Cells[row, 7].Value = "Kết quả đạt được";
                ws.Cells[row, 8].Value = "Người hỗ trợ";
                ws.Cells[row, 9].Value = "Lý do chưa hoàn thành";
                ws.Cells[row, 10].Value = "Hướng giải quyết";
                ws.Cells[row, 11].Value = "Ghi chú";

                using (var headerRange = ws.Cells[row, 1, row, 11])
                {
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightSteelBlue);
                    headerRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    headerRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    headerRange.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                }
                row++;

                int sttNoiDung = 1;
                foreach (var nd in bc.NoiDungs)
                {
                    ws.Cells[row, 1].Value = sttNoiDung++;
                    ws.Cells[row, 2].Value = nd.TrachNhiemChinh;
                    ws.Cells[row, 3].Value = nd.NoiDung;
                    ws.Cells[row, 4].Value = nd.NgayHoanThanh?.ToString("dd/MM/yyyy");
                    ws.Cells[row, 5].Value = nd.MucDoUuTien;
                    ws.Cells[row, 6].Value = nd.TienDo;
                    ws.Cells[row, 7].Value = nd.KetQuaDatDuoc;
                    ws.Cells[row, 8].Value = nd.TrachNhiemHoTro;
                    ws.Cells[row, 9].Value = nd.LyDoChuaHoanThanh;
                    ws.Cells[row, 10].Value = nd.HuongGiaiQuyet;
                    ws.Cells[row, 11].Value = nd.GhiChu;

                    using (var rowRange = ws.Cells[row, 1, row, 11])
                    {
                        rowRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        rowRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        rowRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        rowRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        rowRange.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                        rowRange.Style.WrapText = true;
                    }
                    row++;
                }

                // ==== PHẦN 4: CĂN LỀ VÀ KHOẢNG TRẮNG ====
                ws.Column(1).Width = 5;
                ws.Column(2).Width = 15;
                ws.Column(3).Width = 45;  // ✅ Nội dung rộng nhất
                ws.Column(4).Width = 15;
                ws.Column(5).Width = 25;
                ws.Column(6).Width = 20;
                ws.Column(7).Width = 25;
                ws.Column(8).Width = 10;
                ws.Column(9).Width = 25;
                ws.Column(10).Width = 25;
                ws.Column(11).Width = 25;

                row += 3; // khoảng cách giữa các báo cáo
                stt++;
            }

            var stream = new MemoryStream(package.GetAsByteArray());
            string fileName = $"TongHopBaoCaoTuan_{tuan}_{maNhanVien}.xlsx";
            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        // GET: /BaoCaoTuan/ThongKeNangCao
        [Authorize(Roles = "Admin,GiamDoc")]
        public IActionResult ThongKeNangCao(
            string tuan,
            string maPhongBan,
            string maNhanVien,
            string tienDo,
            int page = 1,
            int pageSize = 100)
        {
            // --- Truy vấn cơ bản ---
            var baseQuery = _context.BaoCaoTuans
                .Include(b => b.NguoiBaoCao).ThenInclude(nv => nv.PhongBan)
                .Include(b => b.BaoCaoCho).ThenInclude(nd => nd.PhongBan)
                .Include(b => b.NoiDungs)
                .AsNoTracking()
                .AsQueryable();

            // --- Bộ lọc ---
            if (!string.IsNullOrEmpty(tuan))
                baseQuery = baseQuery.Where(x => x.Tuan == tuan);

            if (!string.IsNullOrEmpty(maPhongBan))
                baseQuery = baseQuery.Where(x => x.NguoiBaoCao.MaPhongBan == maPhongBan);

            if (!string.IsNullOrEmpty(maNhanVien))
                baseQuery = baseQuery.Where(x =>
                    x.NguoiBaoCao.MaNV.Contains(maNhanVien) ||
                    x.NguoiBaoCao.HoTen.Contains(maNhanVien));

            if (!string.IsNullOrEmpty(tienDo))
                baseQuery = baseQuery.Where(x => x.NoiDungs.Any(nd => nd.TienDo == tienDo));

            // --- Trải phẳng dữ liệu (mỗi dòng là 1 nội dung thực tế) ---
            var flatList = baseQuery
                .ToList() // EF không hỗ trợ SelectMany trực tiếp với Include nên cần ToList()
                .SelectMany(bc => bc.NoiDungs.Select(nd => new
                {
                    BaoCao = bc,
                    NoiDung = nd
                }));

            // --- Lọc thêm theo tiến độ nếu cần ---
            if (!string.IsNullOrEmpty(tienDo))
                flatList = flatList.Where(x => x.NoiDung.TienDo == tienDo);

            // --- Tổng số dòng thực tế ---
            int totalCount = flatList.Count();
            int totalPages = (int)Math.Ceiling(totalCount / (double)pageSize);

            // --- Lấy trang hiện tại ---
            var pagedData = flatList
                .OrderByDescending(x => x.BaoCao.Tuan)
                .Skip((page - 1) * pageSize)
                .Take(pageSize)
                .ToList();

            // --- Gộp lại để gửi ra View theo cấu trúc cũ ---
            var grouped = pagedData
                .GroupBy(x => x.BaoCao)
                .Select(g =>
                {
                    var bc = g.Key;
                    bc.NoiDungs = g.Select(x => x.NoiDung).ToList();
                    return bc;
                })
                .ToList();

            // --- Dropdown chọn tuần ---
            ViewBag.TuanOptions = new SelectList(
                _context.BaoCaoTuans
                    .Select(x => x.Tuan)
                    .Distinct()
                    .OrderBy(x => x)
                    .ToList()
            );

            // --- Dropdown chọn phòng ban ---
            ViewBag.PhongBanOptions = new SelectList(
                _context.Users
                    .Where(x => x.MaPhongBan != null)
                    .Select(x => x.MaPhongBan)
                    .Distinct()
                    .OrderBy(x => x)
                    .ToList()
            );

            // --- Dropdown chọn nhân viên ---
            ViewBag.NhanVienOptions = new SelectList(
                _context.Users
                    .OrderBy(u => u.HoTen)
                    .Select(u => new
                    {
                        Value = u.MaNV,
                        Text = $"{u.HoTen} ({u.MaNV})"
                    })
                    .ToList(),
                "Value", "Text"
            );

            // --- Gửi thông tin phân trang ra View ---
            ViewBag.CurrentPage = page;
            ViewBag.TotalPages = totalPages;
            ViewBag.PageSize = pageSize;
            ViewBag.TotalCount = totalCount;

            return View(grouped);
        }


        // Xuất Excel cho Thống kê nâng cao
        [Authorize(Roles = "Admin,GiamDoc")]
        public IActionResult ExportExcelNangCao(string tuan, string maPhongBan, string maNhanVien, string tienDo)
        {
            var query = _context.BaoCaoTuans
                .Include(b => b.NguoiBaoCao).ThenInclude(nd => nd.PhongBan)
                .Include(b => b.BaoCaoCho).ThenInclude(nd => nd.PhongBan)
                .Include(b => b.NoiDungs)
                .AsQueryable();

            if (!string.IsNullOrEmpty(tuan))
                query = query.Where(b => b.Tuan == tuan);

            if (!string.IsNullOrEmpty(maPhongBan))
                query = query.Where(b => b.NguoiBaoCao.MaPhongBan == maPhongBan);

            if (!string.IsNullOrEmpty(maNhanVien))
                query = query.Where(b => b.NguoiBaoCao.MaNV.Contains(maNhanVien)
                                      || b.NguoiBaoCao.HoTen.Contains(maNhanVien));

            if (!string.IsNullOrEmpty(tienDo))
                query = query.Where(b => b.NoiDungs.Any(nd => nd.TienDo == tienDo));

            var ds = query.OrderBy(b => b.Tuan).ToList();

            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("ThongKeNangCao");

            // === LOGO & TIÊU ĐỀ BÁO CÁO ===
            var logoPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo.png");
            if (System.IO.File.Exists(logoPath))
            {
                var picture = ws.Drawings.AddPicture("Logo", new FileInfo(logoPath));
                picture.SetPosition(0, 0, 0, 0); // hàng 1, cột 1
                picture.SetSize(220, 40); // chỉnh kích thước phù hợp
            }

            // Tiêu đề chính
            ws.Cells["C1:V1"].Merge = true;
            ws.Cells["C1"].Value = "BÁO CÁO THỐNG KÊ NÂNG CAO - TOÀN CÔNG TY LÂM HIỆP HƯNG";
            ws.Cells["C1"].Style.Font.Bold = true;
            ws.Cells["C1"].Style.Font.Size = 18;
            ws.Cells["C1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["C1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            // Dòng phụ đề (ví dụ: ngày xuất)
            ws.Cells["C2:V2"].Merge = true;
            ws.Cells["C2"].Value = $"Ngày xuất báo cáo: {DateTime.Now:dd/MM/yyyy HH:mm}";
            ws.Cells["C2"].Style.Font.Italic = true;
            ws.Cells["C2"].Style.Font.Size = 11;
            ws.Cells["C2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["C2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            // === HEADER DỮ LIỆU ===
            int startRow = 4;
            ws.Cells[startRow, 1].Value = "STT";
            ws.Cells[startRow, 2].Value = "Tuần";
            ws.Cells[startRow, 3].Value = "Từ ngày";
            ws.Cells[startRow, 4].Value = "Đến ngày";
            ws.Cells[startRow, 5].Value = "Mã NV";
            ws.Cells[startRow, 6].Value = "Tên nhân viên";
            ws.Cells[startRow, 7].Value = "Mã phòng ban";
            ws.Cells[startRow, 8].Value = "Tên phòng ban";
            ws.Cells[startRow, 9].Value = "Báo cáo cho";
            ws.Cells[startRow, 10].Value = "Phòng ban QLTT";
            ws.Cells[startRow, 11].Value = "Mảng công việc";
            ws.Cells[startRow, 12].Value = "Nội dung";
            ws.Cells[startRow, 13].Value = "Ngày hoàn thành";
            ws.Cells[startRow, 14].Value = "Mục tiêu";
            ws.Cells[startRow, 15].Value = "Tiến độ";
            ws.Cells[startRow, 16].Value = "Người hỗ trợ";
            ws.Cells[startRow, 17].Value = "Kết quả đạt được";
            ws.Cells[startRow, 18].Value = "Lí do chưa hoàn thành";
            ws.Cells[startRow, 19].Value = "Hướng giải quyết";
            ws.Cells[startRow, 20].Value = "Ghi chú";
            ws.Cells[startRow, 21].Value = "Thời gian nộp";

            using (var range = ws.Cells[startRow, 1, startRow, 21])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(220, 230, 241));
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                range.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            }

            // === DỮ LIỆU CHI TIẾT ===
            int row = startRow + 1;
            int stt = 1;
            foreach (var bc in ds)
            {
                foreach (var nd in bc.NoiDungs)
                {
                    ws.Cells[row, 1].Value = stt++;
                    ws.Cells[row, 2].Value = bc.Tuan;
                    ws.Cells[row, 3].Value = bc.TuNgay.ToString("dd/MM/yyyy");
                    ws.Cells[row, 4].Value = bc.DenNgay.ToString("dd/MM/yyyy");
                    ws.Cells[row, 5].Value = bc.NguoiBaoCao?.MaNV;
                    ws.Cells[row, 6].Value = bc.NguoiBaoCao?.HoTen;
                    ws.Cells[row, 7].Value = bc.NguoiBaoCao?.MaPhongBan;
                    ws.Cells[row, 8].Value = bc.NguoiBaoCao?.PhongBan?.TenPhongBan;
                    ws.Cells[row, 9].Value = $"{bc.BaoCaoCho?.HoTen} ({bc.BaoCaoCho?.MaNV})";
                    ws.Cells[row, 10].Value = $"{bc.BaoCaoCho?.PhongBan?.MaPhongBan} - {bc.BaoCaoCho?.PhongBan?.TenPhongBan}";
                    ws.Cells[row, 11].Value = nd.TrachNhiemChinh;
                    ws.Cells[row, 12].Value = nd.NoiDung;
                    ws.Cells[row, 13].Value = nd.NgayHoanThanh?.ToString("dd/MM/yyyy");
                    ws.Cells[row, 14].Value = nd.MucDoUuTien;
                    ws.Cells[row, 15].Value = nd.TienDo;
                    ws.Cells[row, 16].Value = nd.TrachNhiemHoTro;
                    ws.Cells[row, 17].Value = nd.KetQuaDatDuoc;
                    ws.Cells[row, 18].Value = nd.LyDoChuaHoanThanh;
                    ws.Cells[row, 19].Value = nd.HuongGiaiQuyet;
                    ws.Cells[row, 20].Value = nd.GhiChu;
                    ws.Cells[row, 21].Value = bc.NgayTao.ToString("HH:mm dd/MM/yyyy");
                    row++;
                }
            }

            // Viền & căn chỉnh
            using (var range = ws.Cells[startRow, 1, row - 1, 21])
            {
                range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            }

            ws.Cells.AutoFitColumns();

            var stream = new MemoryStream(package.GetAsByteArray());
            return File(stream,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"ThongKeNangCao_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
        }


        // GET: /BaoCaoTuan/TrangThai
        [HttpGet]
        public async Task<IActionResult> TrangThai(string? tuan, string? maPhongBan, bool chiHienThiChuaGui = false, bool chiHienThiDaGui = false)
        {
            // Dropdown tuần (dùng helper có sẵn)
            int yearNow = DateTime.Now.Year;
            var weekOptions = GenerateWeekDropdown(yearNow, yearNow)
                   .Select(i => new SelectListItem
                   {
                       Value = i.Value,
                       Text = i.Text,
                       Selected = (i.Value == tuan) // ✅ Giữ trạng thái tuần đã chọn
                   })
                   .ToList();
            ViewBag.TuanOptions = weekOptions;
            // Phòng ban dropdown
            ViewBag.PhongBanOptions = _context.PhongBans
                .AsEnumerable() // ép EF Core load xong rồi mới xử lý ở memory
                .Select(p => new SelectListItem
                {
                    Value = p.MaPhongBan,
                    Text = $"{p.TenCongTy} - {p.TenPhongBan}",
                    Selected = (p.MaPhongBan == maPhongBan) // ✅ Đánh dấu phòng ban đang chọn
                })
                .OrderBy(p => p.Text)
                .ToList();

            ViewBag.SelectedPhongBan = maPhongBan;

            // Gắn lại giá trị đã chọn để view hiển thị
            ViewBag.SelectedTuan = tuan;
            ViewBag.SelectedPhongBan = maPhongBan;
            ViewBag.ChiHienThiChuaGui = chiHienThiChuaGui;
            ViewBag.ChiHienThiDaGui = chiHienThiDaGui;

            // Nếu chưa chọn tuần thì trả view rỗng (bảo mật / tiện dùng)
            if (string.IsNullOrEmpty(tuan))
            {
                return View(new List<TrangThaiViewModel>());
            }

            // Lấy employees theo phòng ban (hoặc tất cả nếu không chọn)
            var employeesQuery = _context.Users.AsQueryable();
            if (!string.IsNullOrEmpty(maPhongBan))
                employeesQuery = employeesQuery.Where(u => u.MaPhongBan == maPhongBan);

            var employees = await employeesQuery
                .OrderBy(u => u.MaNV)
                .ToListAsync();

            // Lấy báo cáo tuần đó
            var baoCaos = await _context.BaoCaoTuans
                .Include(b => b.BaoCaoCho)
                .Where(b => b.Tuan == tuan)
                .ToListAsync();

            // Map báo cáo theo userId (nếu có nhiều bản ghi cho 1 user, lấy bản đầu)
            var baocaoByUser = baoCaos
                .GroupBy(b => b.NguoiBaoCaoId)
                .ToDictionary(g => g.Key, g => g.First());

            // Lấy tên phòng ban map để tránh nhiều query
            var pbDict = await _context.PhongBans
                .ToDictionaryAsync(p => p.MaPhongBan, p => p.TenPhongBan);

            // Compute week range helper (dùng same logic với controller)
            (DateTime weekStart, DateTime weekEnd) = GetWeekRangeFromCode(tuan);

            var list = new List<TrangThaiViewModel>();
            int idx = 1;
            foreach (var nv in employees)
            {
                baocaoByUser.TryGetValue(nv.Id, out var bc);

                list.Add(new TrangThaiViewModel
                {
                    STT = idx++,
                    Tuan = tuan,
                    TuNgay = bc?.TuNgay ?? weekStart,
                    DenNgay = bc?.DenNgay ?? weekEnd,
                    MaPhongBan = nv.MaPhongBan,
                    TenPhongBan = pbDict.TryGetValue(nv.MaPhongBan, out var tpb) ? tpb : "",
                    MaNV = nv.MaNV,
                    TenNV = nv.HoTen,
                    TenQuanLy = bc?.BaoCaoCho?.HoTen ?? "-",
                    ThoiGianGui = bc?.NgayTao,
                    DaGui = bc != null
                });

                //list.Add(item);
            }

            // ✅ Lọc theo 2 checkbox
            // Nếu muốn chỉ hiển thị chưa gửi:
            if (chiHienThiChuaGui && !chiHienThiDaGui)
                list = list.Where(x => !x.DaGui).ToList();
            // Nếu muốn chỉ hiển thị đã gửi
            if (chiHienThiDaGui && !chiHienThiChuaGui)
                list = list.Where(x => x.DaGui).ToList();

            // Optional: thống kê tổng
            ViewBag.Total = list.Count;
            ViewBag.Sent = list.Count(x => x.DaGui);
            ViewBag.NotSent = list.Count(x => !x.DaGui);

            return View(list);
        }

        // Helper private (bỏ vào cuối controller)
        private (DateTime weekStart, DateTime weekEnd) GetWeekRangeFromCode(string tuanCode)
        {
            // Expect Y25W28 or Y2025W28?
            var match = System.Text.RegularExpressions.Regex.Match(tuanCode, @"Y(?<year>\d{2,4})W(?<week>\d{1,2})");
            if (!match.Success)
                return (DateTime.MinValue, DateTime.MinValue);

            int year = int.Parse(match.Groups["year"].Value);
            if (year < 100) year += 2000;
            int week = int.Parse(match.Groups["week"].Value);

            var cal = CultureInfo.CurrentCulture.Calendar;
            var weekRule = CalendarWeekRule.FirstFourDayWeek;
            var firstDay = DayOfWeek.Monday;

            // Find week start (Monday)
            var jan1 = new DateTime(year, 1, 1);
            int daysOffset = DayOfWeek.Thursday - jan1.DayOfWeek;
            var firstThursday = jan1.AddDays(daysOffset);
            int firstWeek = cal.GetWeekOfYear(firstThursday, weekRule, firstDay);
            int delta = week - (firstWeek <= 1 ? 1 : 0);
            var weekStart = firstThursday.AddDays(delta * 7).AddDays(-3);
            if (weekStart.DayOfWeek != DayOfWeek.Monday)
                weekStart = weekStart.AddDays((int)DayOfWeek.Monday - (int)weekStart.DayOfWeek);

            return (weekStart.Date, weekStart.AddDays(6).Date);
        }
    }

}

