---
name: lap-bao-gia
description: Lập bảng báo giá cho Công ty CP Thương mại Tổng hợp Minh Phát theo đúng mẫu Excel của công ty (có logo, thông tin công ty, bảng hàng hóa, thuế GTGT, số tiền bằng chữ, điều khoản, chỗ ký) rồi xuất file Excel và PDF để gửi khách qua Zalo/Email. Dùng skill này mỗi khi người dùng muốn làm, soạn, lập, gửi, sửa hoặc in báo giá / bảng giá / phiếu báo giá / chào giá cho khách hàng, kể cả khi họ chỉ liệt kê tên khách và vài mặt hàng kèm giá (ví dụ "báo giá cho cty ABC 50 ram giấy A4 giá 75k"), hoặc dán danh sách hàng từ tin nhắn, ảnh chụp, file Excel khác.
---

# Lập báo giá Minh Phát

Skill này điền mẫu báo giá chuẩn của công ty (`assets/Mau-bao-gia-Minh-Phat.xlsx`) bằng script `scripts/tao_bao_gia.py`. Script tự làm những việc dễ sai khi làm tay: đánh STT, tính thành tiền, cộng tiền, tính thuế GTGT, **đọc số tiền bằng chữ tiếng Việt**, thêm dòng khi có hơn 12 mặt hàng, ẩn dòng trống, và xuất PDF chỉ gồm trang báo giá.

Người dùng thường là nhân viên văn phòng/kế toán, không rành kỹ thuật. Hãy trả lời bằng tiếng Việt, ngắn gọn, không nói về JSON hay script trừ khi họ hỏi.

## Quy trình

### 1. Thu thập thông tin

Lấy từ tin nhắn của người dùng (hoặc ảnh, file họ gửi):

| Thông tin | Bắt buộc? | Nếu thiếu |
|---|---|---|
| Tên khách hàng | Có | Hỏi lại |
| Danh sách hàng: tên, ĐVT, số lượng, đơn giá | Có | Hỏi lại phần còn thiếu (thường là đơn giá) |
| Người liên hệ, điện thoại, địa chỉ, email khách | Không | Để trống |
| Số báo giá | Không | Script tự đặt `BG-<năm>/<thángngày>`; báo cho người dùng biết |
| Ngày báo giá | Không | Hôm nay |
| Thuế suất GTGT | Không | 10%. Hàng được giảm thuế thì người dùng sẽ nói 8% |
| Điều khoản giao hàng, thanh toán, hiệu lực | Không | Dùng mặc định của mẫu |

Chỉ hỏi khi thiếu thông tin bắt buộc. Gom các câu hỏi vào **một** lần hỏi. Đừng hỏi từng chi tiết không bắt buộc, vì người dùng muốn có báo giá nhanh.

Cách hiểu cách viết tắt hay gặp:
- "75k" = 75.000; "1tr2" hoặc "1,2tr" = 1.200.000.
- Dấu chấm/phẩy là phân cách hàng nghìn kiểu Việt Nam: "1.250.000" = 1250000.
- **Đơn giá trong mẫu là giá chưa thuế.** Nếu người dùng đưa giá *đã gồm VAT*, đổi ra giá chưa thuế (giá ÷ 1,1 hoặc ÷ 1,08, làm tròn đến đồng) và nói rõ cho họ biết đã quy đổi.
- Không có ĐVT thì suy ra hợp lý (giấy → Ram, mực → Hộp, thùng nước → Thùng...). Không chắc thì để "Cái".

### 2. Ghi dữ liệu ra file JSON

```json
{
  "so_bao_gia": "BG-2026/015",
  "ngay": "2026-10-10",
  "khach_hang": {
    "ten": "Công ty TNHH ABC",
    "nguoi_lien_he": "Anh Nam",
    "dien_thoai": "0912 345 678",
    "dia_chi": "12 Lê Lợi, Q.1, TP.HCM",
    "email": "nam@abc.vn"
  },
  "hang_hoa": [
    {"ten": "Giấy in A4 Double A 70gsm", "dvt": "Ram", "so_luong": 100, "don_gia": 72000, "ghi_chu": "Hàng có sẵn"}
  ],
  "vat": 10,
  "dieu_khoan": {
    "hieu_luc": "30 ngày kể từ ngày báo giá",
    "giao_hang": "trong vòng 03 ngày kể từ khi Quý khách xác nhận đặt hàng",
    "dia_diem": "tại kho của Quý khách",
    "thanh_toan": "chuyển khoản, 100% khi nhận hàng",
    "thue": "Đơn giá trên chưa bao gồm thuế GTGT; thuế GTGT được tính riêng như bảng trên."
  },
  "cong_ty": {
    "dia_chi": "...", "mst": "...", "dien_thoai": "...", "email": "...", "website": "...",
    "so_tai_khoan": "...", "ngan_hang": "...", "nguoi_ky": "...", "chuc_vu": "Giám đốc"
  }
}
```

- `so_luong` và `don_gia` phải là **số** (72000), không phải chuỗi ("72.000đ"), vì Excel cần số để tính.
- `ngay` nhận `2026-10-10` hoặc `10/10/2026`.
- `vat` nhận `10` hoặc `0.1`.
- Mọi trường trong `dieu_khoan` và `cong_ty` đều tùy chọn: chỉ ghi những gì người dùng nêu. Các điều khoản được nối sau tiêu đề sẵn có (ví dụ "3. Thời gian giao hàng: " + giá trị), nên không cần lặp lại tiêu đề hay dấu chấm cuối câu.

### 3. Chạy script

```bash
python scripts/tao_bao_gia.py du_lieu.json --out "<thư mục đầu ra>/Bao-gia-<Ten-khach>-<so>.xlsx" --pdf
```

Cần `openpyxl` (`pip install openpyxl` nếu thiếu). Có LibreOffice thì xuất được PDF; không có thì chỉ có file Excel, khi đó báo người dùng tự mở Excel và chọn File → Save As → PDF.

Đặt tên file không dấu, không khoảng trắng, ví dụ `Bao-gia-Cong-ty-ABC-BG-2026-015.xlsx`, để gửi qua Zalo/Email không bị lỗi tên.

Script in ra kết quả gồm: tổng tiền hàng, tiền thuế, tổng cộng, số tiền bằng chữ, và `thong_tin_cong_ty_con_trong` (các ô thông tin công ty vẫn còn chữ mẫu trong ngoặc vuông `[ ]`).

### 4. Kiểm tra và gửi kết quả

- Nếu có PDF, xem nhanh trang PDF (chuyển sang ảnh nếu cần) để chắc chắn bố cục ổn, không tràn trang.
- Gửi file Excel và PDF cho người dùng, kèm tóm tắt ngắn:
  - Số báo giá, tên khách, số mặt hàng
  - Cộng tiền hàng / Thuế GTGT / **Tổng cộng** và số tiền bằng chữ
  - Những gì đã tự giả định (số báo giá tự đặt, ĐVT tự suy ra, giá đã quy đổi về chưa VAT...)
- **Nếu `thong_tin_cong_ty_con_trong` không rỗng**, nhắc người dùng rằng địa chỉ, MST, số tài khoản... của công ty vẫn là chữ mẫu, khách sẽ thấy điều đó trên báo giá. Mời họ gửi thông tin thật để điền vào (bằng mục `cong_ty`). Đây là lỗi dễ khiến báo giá gửi khách trông thiếu chuyên nghiệp, nên đừng bỏ qua.

## Ghi chú

- Mẫu có sẵn 12 dòng hàng. Nhiều hơn thì script tự chèn thêm dòng; quá khoảng 25 mặt hàng thì báo giá tự chia nhiều trang và lặp lại dòng tiêu đề bảng.
- Logo trong mẫu là logo tạm. Nếu người dùng gửi logo thật, thay ảnh trong `assets/Mau-bao-gia-Minh-Phat.xlsx` (trang "Báo giá", góc trên bên trái, ô A1:A4).
- Khi người dùng muốn **sửa báo giá vừa làm** (đổi giá, thêm hàng...), sửa file JSON rồi chạy lại script. Cách này nhanh hơn và ít lỗi hơn sửa trực tiếp file Excel.
