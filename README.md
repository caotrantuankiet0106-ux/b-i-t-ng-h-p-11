# Phần 1 – Hoàn thiện SRS bằng Heading Styles và Auto TOC
Cấu trúc 3 cấp:

1. Giới thiệu dự án (H1) 
→
→ 1.1. Mục tiêu (H2), 1.2. Đối tượng sử dụng (H2).

2. Chức năng hệ thống (H1) 
→
→ 2.1. Quản lý mượn sách (H2) 
→
→ 2.1.1. Luồng mượn sách qua thẻ sinh viên (H3); 2.2. Quản lý trả sách (H2).

3. Yêu cầu bảo mật (H1).

Auto TOC: Gán đúng style 
→
→ References 
→
→ Table of Contents. Cập nhật trang tự động bằng cách click phải bảng mục lục 
→
→ Update Field 
→
→ Update entire table.

# Phần 2 – Chuẩn hóa dữ liệu và AutoFilter
Bảng lỗi vi phạm 5 Quy tắc Vàng:

Lỗi 1 (Hàng 1): Vi phạm quy tắc không gộp ô (Merge cells).

Lỗi 2 (Hàng 5): Vi phạm quy tắc không để hàng trống.

Lỗi 3 (Hàng 6): Vi phạm quy tắc không để dòng tổng cộng nằm lẫn trong vùng dữ liệu.

Record và Field: Field là tên cột (VD: Mã Sách), Record là một dòng dữ liệu hoàn chỉnh của một đối tượng (VD: dòng thông tin SV001).

AutoFilter: Bôi đen bảng 
→
→ Nhấn Ctrl + Shift + L 
→
→ Lọc cột Tình Trạng chọn "Quá hạn".

# Phần 3 – Nối dữ liệu bằng VLOOKUP
Công thức VLOOKUP cơ bản:

Excel
=VLOOKUP(C3, DanhMucSach!$A$2:$C$4, 2, FALSE)
Xử lý lỗi S099: Trả về lỗi #N/A do mã S099 không tồn tại trong danh mục. Dùng IFERROR để khắc phục:

Excel
=IFERROR(VLOOKUP(C3, DanhMucSach!$A$2:$C$4, 2, FALSE), "Không tìm thấy sách")
# Phần 4 – Slide báo cáo tiến độ
Assertion Title: "Dữ liệu mượn sách đã chuẩn hóa thành công và sẵn sàng tích hợp tra cứu tự động"

Con số nổi bật: Làm sạch 100% dữ liệu (loại bỏ ô gộp, hàng trống) và xử lý triệt để lỗi tra cứu bằng hàm.

Biểu đồ khuyên dùng: Donut Chart (Biểu đồ vành khuyên) để minh họa tỷ lệ mượn theo các trạng thái (Đang mượn, Quá hạn, Đã trả).

Rule 10/20/30: Là Slide số 2/10, thời gian trình bày dự kiến 2 phút.
