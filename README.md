# GT35 WEB V4

Phiên bản này đọc đúng cấu trúc file **GT35 Theo doi du lieu Dashboard3.xlsx**:

- 133 cột dữ liệu.
- 16 nhóm cột.
- Nhập theo từng nhóm giống sheet `02 INPUT DATA`.
- Tự tính lại các KPI chính.
- Nhập trực tiếp từ file Excel gốc.
- Trang **Quản lý trại** cho phép thêm/sửa/khóa trại, đổi khu vực, quy mô và quản lý trại.
- Dữ liệu lưu SQLite khi chạy thử hoặc Supabase khi chạy online.

## Chạy thử
Bấm `CHAY_WEB.bat`.

Tài khoản:
- admin@gt35.local
- admin123

## Supabase
1. Tạo project Supabase.
2. Chạy `sql/supabase_schema.sql`.
3. Tạo user trong Authentication.
4. Cập nhật `profiles.role` và `profiles.farm`.
5. Đưa mã lên GitHub và khai báo Secrets trên Streamlit Cloud.


## Điểm mới V4
- Trang Nhập từ Excel hiển thị toàn bộ 133 cột theo từng tab/nhóm, không còn chỉ xem 3 cột Năm–Tuần–Trại.
- Có bước xác nhận trước khi nhập toàn bộ dữ liệu vào Supabase.
