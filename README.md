# 🗓️ HỆ THỐNG BÁO CÁO TUẦN NỘI BỘ CÔNG TY

## 🧩 Giới thiệu

Hệ thống **Báo cáo tuần nội bộ** được xây dựng nhằm giúp công ty quản lý, theo dõi và duyệt báo cáo công việc hằng tuần của toàn bộ nhân viên theo quy trình:

> **Nhân viên → Trưởng phòng → Giám đốc → Admin**

Hệ thống cho phép nhân viên tạo báo cáo nhiều dòng công việc, thêm/xóa dòng linh hoạt, chọn tuần báo cáo theo mã (`Y25W28`), tự động tính **Từ ngày – Đến ngày**, và phân quyền rõ ràng theo vai trò.

---

## ⚙️ Công nghệ sử dụng

| Thành phần | Công nghệ |
|-------------|------------|
| **Ngôn ngữ** | C# |
| **Framework** | ASP.NET Core MVC (Razor View) |
| **ORM** | Entity Framework Core |
| **CSDL** | SQL Server |
| **Xác thực người dùng** | ASP.NET Identity |
| **Frontend** | HTML + Razor Helper + Bootstrap |
| **Export (dự kiến)** | PDF / Excel dành cho cấp quản lý |
| **Triết lý thiết kế** | Thuần server-side (không dùng JavaScript động để tránh lỗi dữ liệu) |

---

## 🏗️ Cấu trúc thư mục

 /Controllers
├── BaoCaoTuanController.cs ← Tạo, Chỉnh sửa, Xem, Duyệt báo cáo
├── AccountController.cs ← Quản lý tài khoản, đăng nhập
└── HomeController.cs ← Trang chính

/Models
├── BaoCaoTuan.cs ← Entity chính (1-n với NoiDungBaoCao)
├── NoiDungBaoCao.cs ← Chi tiết từng dòng báo cáo
├── BaoCaoTuanViewModel.cs ← ViewModel trung gian giữa View & Controller
└── ApplicationUser.cs ← Mở rộng ASP.NET IdentityUser (HoTen, MaPhongBan,...)

/Views
├── BaoCaoTuan/
│ ├── Tao.cshtml ← View tạo báo cáo (desktop & mobile)
│ ├── ChinhSua.cshtml ← View chỉnh sửa báo cáo
│ ├── Xem.cshtml ← Xem dữ liệu đã báo cáo
│ └── Duyet.cshtml ← Duyệt báo cáo (dành cho cấp quản lý)
├── Shared/
│ ├── _Layout.cshtml
│ └── _ValidationScriptsPartial.cshtml

<img width="765" height="559" alt="image" src="https://github.com/user-attachments/assets/2f815b74-72c4-49e6-88dc-3cc7b181c0c4" />


---

## 🧠 Luồng hoạt động chính

### 📝 1. Tạo báo cáo tuần
- Người dùng đăng nhập và truy cập `/BaoCaoTuan/Tao`
- Chọn **phòng ban**, **người nhận báo cáo**, **tuần báo cáo (Y25W02)**  
- Hệ thống tự động tính **Từ ngày – Đến ngày** tương ứng với tuần đó  
- Có thể **thêm/xóa dòng công việc** tùy ý  
- Khi nhấn **Gửi báo cáo**, hệ thống lưu vào CSDL

> Tính năng `ReloadDropdowns()` được dùng để nạp lại danh sách dropdown sau mỗi thao tác thêm/xóa dòng.

---

### 🛠️ 2. Chỉnh sửa báo cáo
- Hiển thị lại dữ liệu cũ để chỉnh sửa.  
- Các trường **Tuần báo cáo**, **Mã phòng ban**, **Người nhận báo cáo** được **khóa** không cho chỉnh.  
- Giữ nguyên dữ liệu **Từ ngày – Đến ngày** theo tuần cũ (khắc phục lỗi 01/01/0001).  
- Cho phép cập nhật, thêm, xóa dòng công việc rồi lưu lại.

---

### 👀 3. Xem dữ liệu báo cáo
- Nhân viên chỉ xem **báo cáo của chính mình**
- Trưởng phòng xem **tất cả nhân viên trong phòng**
- Giám đốc và Admin xem **toàn bộ công ty**
- Có thể **lọc theo tuần, phòng ban, người báo cáo**
- Giao diện dạng bảng tổng hợp, dễ tra cứu

---

### 🔐 4. Phân quyền người dùng

| Vai trò | Quyền hạn |
|----------|-----------|
| **Nhân viên** | Tạo, xem, chỉnh sửa báo cáo của mình |
| **Trưởng phòng** | Xem, duyệt, export báo cáo của phòng mình |
| **Giám đốc** | Xem, duyệt, export toàn công ty |
| **Admin** | Toàn quyền quản lý người dùng & dữ liệu |

---

## 💡 Đặc điểm nổi bật

- **Không dùng JavaScript/AJAX**, reload thuần server-side để tránh mất dữ liệu khi thêm dòng.  
- **Binding danh sách động** qua `List<NoiDungBaoCaoViewModel>`  
- **Tự động tính tuần – ngày** dựa theo mã tuần `Y25Wxx` theo chuẩn ISOWeek  
- **Tương thích đa nền tảng**:  
  - Desktop: dạng bảng  
  - Mobile: form xếp dọc dễ nhập liệu  
- **Kiến trúc mở rộng**: có thể tích hợp module Duyệt, Ghi chú, Xuất PDF/Excel trong tương lai.

---

## 📅 Roadmap (dự kiến)
- [x] Tạo báo cáo tuần  
- [x] Chỉnh sửa báo cáo  
- [x] Xem dữ liệu đã gửi  
- [ ] Duyệt báo cáo (Trưởng phòng, Giám đốc)  
- [ ] Ghi chú và phản hồi cấp trên  
- [ ] Xuất Excel/PDF  
- [ ] Thống kê tiến độ toàn công ty

---

## 🧱 Kiến trúc tổng quan
┌────────────────────────┐
│ Giao diện UI │
│ (Razor View .cshtml) │
└──────────┬─────────────┘
│
▼
┌────────────────────────┐
│ Controller (MVC) │
│ BaoCaoTuanController │
└──────────┬─────────────┘
│
▼
┌────────────────────────┐
│ ViewModel / Model lớp │
│ BaoCaoTuanViewModel │
│ NoiDungBaoCao │
└──────────┬─────────────┘
│
▼
┌────────────────────────┐
│ Entity Framework │
│ (SQL Server) │
└────────────────────────┘

<img width="345" height="634" alt="image" src="https://github.com/user-attachments/assets/fbd47db1-75d3-42eb-a90c-55f84fb6a8cb" />

---

## 👤 Tác giả & quản lý hệ thống

**Người đảm nhiệm:** _(TienDat – Quản trị & Phát triển hệ thống nội bộ)_  
**Mục tiêu:** Tối ưu quy trình báo cáo tuần tự động, an toàn và dễ bảo trì cho doanh nghiệp.

---

> 💬 _“Hệ thống báo cáo tuần được xây dựng theo hướng thuần server-side, chú trọng ổn định và bảo toàn dữ liệu, phù hợp cho môi trường nội bộ doanh nghiệp có yêu cầu cao về tính chính xác và kiểm soát phân quyền.”_

