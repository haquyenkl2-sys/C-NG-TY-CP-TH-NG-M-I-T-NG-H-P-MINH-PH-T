---
name: lap-bao-gia
description: Lập bảng báo giá Excel trình bày chuyên nghiệp cho các công ty người dùng quản lý (hiện có Công ty CP Thương mại Tổng hợp Minh Phát và Công ty TNHH Kỹ thuật Công trình BLT), mỗi công ty có tên, logo, MST, tài khoản riêng (có logo, thông tin công ty, bảng hàng hóa, thuế GTGT, số tiền bằng chữ, điều khoản, chỗ ký) rồi xuất file Excel và PDF để gửi khách qua Zalo/Email. Dùng skill này mỗi khi người dùng muốn làm, soạn, lập, gửi, sửa hoặc in báo giá / bảng giá / phiếu báo giá / chào giá cho khách hàng, kể cả khi họ chỉ liệt kê tên khách và vài mặt hàng kèm giá (ví dụ "báo giá cho cty ABC 50 ram giấy A4 giá 75k"), hoặc dán danh sách hàng từ tin nhắn, ảnh chụp, file Excel khác. Làm được cả bản song ngữ Việt – Trung (报价单) khi khách là công ty Trung Quốc hoặc người dùng xin "bản song ngữ", "tiếng Trung", "中文". Cũng dùng khi người dùng muốn thêm một công ty mới để làm báo giá, hoặc tạo mẫu báo giá trống cho một công ty.
---

# Lập báo giá

Script `scripts/tao_bao_gia.py` dựng bảng báo giá Excel từ đầu theo một bố cục chuyên nghiệp cố định: đầu trang có logo bên trái và thông tin công ty bên phải trên đường kẻ hai màu, tiêu đề lớn, hai khung "Thông tin khách hàng" và "Thông tin báo giá", dải "Công trình", bảng hàng có tiêu đề nền màu, dòng xen màu và nhóm hàng (I, II...), khối tổng tiền nổi bật, điều khoản thương mại, chỗ ký hai bên, chân trang có số trang. Màu chủ đạo lấy theo hồ sơ công ty.

Script cũng tự làm những việc dễ sai khi làm tay: đánh STT, công thức thành tiền/tổng/thuế (sửa số trong Excel là tự tính lại), **đọc số tiền bằng chữ**, giãn dòng theo nội dung, ngắt trang không cắt ngang nhóm hàng/khối tổng/chữ ký, và xuất PDF.

Người dùng thường là nhân viên văn phòng/kế toán, không rành kỹ thuật. Hãy trả lời bằng tiếng Việt, ngắn gọn, không nói về JSON hay script trừ khi họ hỏi.

## Các công ty đã lưu

Mỗi công ty là một file hồ sơ trong `assets/cong-ty/` (logo ở `assets/logo/`). Xem danh sách bằng:

```bash
python scripts/tao_bao_gia.py --danh-sach
```

| Mã | Công ty | Người dùng hay gọi |
|---|---|---|
| `minh-phat` | CÔNG TY CỔ PHẦN THƯƠNG MẠI TỔNG HỢP MINH PHÁT | Minh Phát, MP |
| `blt` | CÔNG TY TNHH KỸ THUẬT CÔNG TRÌNH BLT | BLT |

Nếu người dùng không nói rõ công ty nào báo giá mà đang có nhiều công ty, hỏi lại một câu ("Báo giá của Minh Phát hay BLT ạ?"), vì sai tên công ty trên báo giá là lỗi nghiêm trọng. Trong cùng cuộc trò chuyện, nếu họ đã chọn một công ty thì dùng tiếp công ty đó.

### Thêm công ty mới

Khi người dùng gửi thông tin một công ty mới (gõ tay, ảnh danh thiếp, ảnh giấy đăng ký kinh doanh...), tạo file `assets/cong-ty/<ma>.json`:

```json
{
  "ma": "blt",
  "ten_goi": ["BLT", "Kỹ thuật Công trình BLT"],
  "ten": "CÔNG TY TNHH KỸ THUẬT CÔNG TRÌNH BLT",
  "dia_chi": "...", "mst": "...", "dien_thoai": "0984 586 089", "email": "...", "website": "",
  "so_tai_khoan": "...", "ngan_hang": "...",
  "chu_tai_khoan": "CT TNHH KY THUAT CONG TRINH BLT",
  "nguoi_ky": "HỒ VĂN BÌNH", "chuc_vu": "Giám đốc",
  "logo": "blt.png", "logo_tam": true,
  "mau": {"chinh": "0D2B66", "nhan": "F7931E"}
}
```

- `ma`: chữ thường, không dấu, ngắn. `ten_goi`: các cách người dùng hay gọi công ty.
- `website` để `""` nếu không có, khi đó đầu trang chỉ ghi email.
- `chu_tai_khoan`: ghi đúng tên chủ tài khoản ngân hàng nếu khác tên công ty (thường là tên không dấu, viết tắt). Bỏ trống thì dùng tên công ty.
- Viết hoa chữ cái đầu cho tên đường, phường, tỉnh; nhóm số điện thoại cho dễ đọc. Giữ nguyên số MST và số tài khoản.
- `mau`: màu chủ đạo (`chinh`: nền tiêu đề bảng, khối tổng, chữ tiêu đề) và màu nhấn (`nhan`: đường kẻ mảnh) dạng mã hex, nên lấy theo màu logo. Bỏ trống thì dùng xanh đậm + vàng đồng. Chọn `chinh` đủ đậm để chữ trắng trên nền đọc rõ.
- Logo: nếu người dùng gửi ảnh logo thì lưu vào `assets/logo/<ma>.png` và đặt `"logo_tam": false`. Nếu chưa có, tạo logo tạm hình tròn có chữ viết tắt (theo kiểu `logo/minh-phat.png`) và đặt `"logo_tam": true`. **Không được để logo của công ty khác** trên báo giá.

Sau đó tạo mẫu trống để người dùng xem trước:

```bash
python scripts/tao_bao_gia.py --mau <ma> --out "<thư mục>/Mau-bao-gia-<MA>.xlsx" --pdf
```

## Quy trình lập báo giá

### 1. Thu thập thông tin

Lấy từ tin nhắn của người dùng (hoặc ảnh, file họ gửi):

| Thông tin | Bắt buộc? | Nếu thiếu |
|---|---|---|
| Công ty báo giá (Minh Phát, BLT...) | Có, khi có nhiều công ty | Hỏi lại |
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
- **Đơn giá mặc định là giá chưa thuế**, bảng tự cộng dòng thuế GTGT. Nếu người dùng đưa giá *đã gồm VAT* hoặc yêu cầu tính giá theo kiểu "giá có thuế rồi nhân thêm X%", đặt `"gia_da_gom_vat": true` và `"vat": <thuế suất đã gồm>`: khối tổng ghi "Cộng tiền hàng (đã bao gồm thuế GTGT 8%)", không cộng thuế lần nữa, điều khoản 1 tự ghi "Đơn giá trên đã bao gồm thuế GTGT 8%". Đừng để khách bị tính thuế hai lần.
- Tính giá từ hóa đơn đầu vào (giá mua chưa thuế × (1 + thuế suất) × (1 + % lợi nhuận)...): tính bằng code, làm tròn đơn giá đến đồng, và ghi rõ cách tính cho người dùng.
- Không có ĐVT thì suy ra hợp lý (giấy → Ram, mực → Hộp, thùng nước → Thùng...). Không chắc thì để "Cái".

### 2. Ghi dữ liệu ra file JSON

```json
{
  "cong_ty": "blt",
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
  }
}
```

- `cong_trinh` (tùy chọn): khi báo giá cho một công trình/dự án, ghi tên hạng mục + công trình; nó hiện thành dải "CÔNG TRÌNH" dưới hai khung thông tin.
- `nhom` (tùy chọn, trong từng mặt hàng): tên nhóm hàng, ví dụ "Tủ điện, tủ điều khiển", "Cáp điện". Các mặt hàng liền nhau cùng nhóm được gom dưới một dòng tiêu đề nhóm đánh số La Mã. Nên chia nhóm khi có từ khoảng 10 mặt hàng trở lên thuộc nhiều loại, vì bảng dễ đọc hơn hẳn.
- `hieu_luc_ngay` (tùy chọn, mặc định 30): số ngày hiệu lực, hiện trong khung "Thông tin báo giá" và điều khoản 2.
- `dieu_khoan.them` (tùy chọn): danh sách điều khoản bổ sung, ví dụ `["Bảo hành 12 tháng theo tiêu chuẩn nhà sản xuất"]`.
- `cong_ty`: mã hoặc tên gọi của công ty báo giá (`"blt"`, `"Minh Phát"`). Muốn đổi tạm một trường cho riêng báo giá này thì dùng dạng `{"ma": "blt", "nguoi_ky": "..."}`.

- `so_luong` và `don_gia` phải là **số** (72000), không phải chuỗi ("72.000đ"), vì Excel cần số để tính.
- `ngay` nhận `2026-10-10` hoặc `10/10/2026`.
- `vat` nhận `10` hoặc `0.1`.
- Mọi trường trong `dieu_khoan` đều tùy chọn: chỉ ghi những gì người dùng nêu. Các điều khoản được nối sau tiêu đề sẵn có (ví dụ "Thời gian giao hàng: " + giá trị), nên không cần lặp lại tiêu đề hay dấu chấm cuối câu. Điều khoản cuối luôn là thông tin chuyển khoản lấy từ hồ sơ công ty.

### 3. Chạy script

```bash
python scripts/tao_bao_gia.py du_lieu.json --out "<thư mục đầu ra>/Bao-gia-<Ten-khach>-<so>.xlsx" --pdf
```

Cần `openpyxl` (`pip install openpyxl` nếu thiếu). Có LibreOffice thì xuất được PDF; không có thì chỉ có file Excel, khi đó báo người dùng tự mở Excel và chọn File → Save As → PDF.

Đặt tên file không dấu, không khoảng trắng, ví dụ `Bao-gia-Cong-ty-ABC-BG-2026-015.xlsx`, để gửi qua Zalo/Email không bị lỗi tên.

Script in ra kết quả gồm: tên công ty báo giá, `logo_tam`, tổng tiền hàng, tiền thuế, tổng cộng, số tiền bằng chữ, và `thong_tin_cong_ty_con_trong` (các trường hồ sơ công ty còn trống hoặc còn chữ mẫu `[ ]`).

### 4. Kiểm tra và gửi kết quả

- Nếu có PDF, xem nhanh trang PDF (chuyển sang ảnh nếu cần) để chắc chắn bố cục ổn, không tràn trang.
- Gửi file Excel và PDF cho người dùng, kèm tóm tắt ngắn:
  - Công ty báo giá, số báo giá, tên khách, số mặt hàng
  - Cộng tiền hàng / Thuế GTGT / **Tổng cộng** và số tiền bằng chữ
  - Những gì đã tự giả định (số báo giá tự đặt, ĐVT tự suy ra, giá đã quy đổi về chưa VAT...)
- **Nếu `thong_tin_cong_ty_con_trong` không rỗng**, nhắc người dùng rằng địa chỉ, MST, số tài khoản... của công ty còn thiếu, khách sẽ thấy điều đó trên báo giá. Mời họ gửi thông tin thật để điền vào (bằng mục `cong_ty`). Đây là lỗi dễ khiến báo giá gửi khách trông thiếu chuyên nghiệp, nên đừng bỏ qua. Tương tự, nếu `logo_tam` là true thì nhắc một câu rằng logo đang là logo tạm và mời họ gửi logo thật.

## Bản song ngữ Việt – Trung

Khi người dùng cần báo giá song ngữ (khách Trung Quốc, BOQ song ngữ, "bản tiếng Trung"...), thêm vào JSON:

```json
{
  "song_ngu": true,
  "cong_trinh_zh": "MEGA一期MS1一楼&BP2空调安装工程——空调电气系统部分",
  "dieu_khoan_zh": {"thue": "...", "hieu_luc": "...", "giao_hang": "...", "dia_diem": "...", "thanh_toan": "..."},
  "hang_hoa": [{"ten": "Cáp điện động lực lõi đồng YJV-3×70+1×35", "ten_zh": "电力电缆 铜芯 YJV-3×70+1×35",
                "dvt": "m", "dvt_zh": "米", "ghi_chu": "Hàng trong nước", "ghi_chu_zh": "国产优质", ...}]
}
```

- Script tự đổi mọi nhãn (tiêu đề, cột bảng, tổng tiền, điều khoản, chỗ ký) sang hai dòng Việt/Trung và thêm số tiền bằng chữ tiếng Trung (大写金额, ví dụ 壹仟壹佰万越南盾整).
- Mỗi điều khoản đã nêu bằng tiếng Việt cần có bản dịch trong `dieu_khoan_zh`; điều khoản để mặc định thì script có sẵn bản dịch.
- Nếu file của khách đã có tên tiếng Trung (BOQ song ngữ), **dùng đúng thuật ngữ gốc của khách** cho `ten_zh` thay vì tự dịch, vì khách sẽ đối chiếu với bảng của họ.
- Tên hãng (ABB, Schneider, LS...) chỉ ghi bằng chữ Latin trong `ghi_chu`, không cần `ghi_chu_zh`, vì hai bên đều đọc được.
- Tên công ty, địa chỉ giữ nguyên tiếng Việt (tên pháp lý). Chức vụ người ký tự dịch (Giám đốc → 总经理); muốn khác thì thêm `chuc_vu_zh` vào hồ sơ công ty.
- Bản song ngữ dài hơn. Script tự thu nhỏ nhẹ khi trang cuối quá ít nội dung; muốn ép số trang thì đặt `"so_trang": <số trang>`, rồi xem lại PDF để chắc chữ vẫn đọc rõ.

## Báo giá từ bảng khối lượng (BOQ) của khách

Khách công trình thường gửi file Excel bảng khối lượng (có khi song ngữ Trung–Việt) với cột khối lượng, đơn giá vật tư, đơn giá nhân công, đơn giá tổng hợp. Khi người dùng muốn làm báo giá từ file đó:
- Chỉ lấy các dòng **đã có đơn giá**; dòng giá trống/0 thì bỏ và báo lại cho người dùng.
- Đơn giá = đơn giá tổng hợp (vật tư + nhân công); ghi rõ trong điều khoản 1 rằng giá đã gồm vật tư và nhân công.
- Viết tên hàng bằng tiếng Việt, gọn: tên + quy cách chính (mã hiệu, kích thước, tiết diện cáp). Thương hiệu đưa vào Ghi chú.
- Lấy số liệu (khối lượng, đơn giá) bằng code từ file, đừng gõ tay, rồi đối chiếu tổng tiền với tổng trong file.
- Ghi `cong_trinh` và điều khoản phù hợp công trình (khối lượng quyết toán theo thực tế, giao hàng theo tiến độ...).

## Ghi chú

- Số dòng hàng không giới hạn. Bảng dài tự chia nhiều trang, lặp lại dòng tiêu đề bảng, và tự ngắt trang để tiêu đề nhóm, khối tổng tiền, khối chữ ký không bị tách.
- `--mau <ma>` tạo mẫu trống (15 dòng hàng có sẵn công thức, chỗ điền bằng dấu chấm, kèm trang "Hướng dẫn") cho người dùng tự điền trong Excel; thêm `--song-ngu` để có mẫu song ngữ.
- Khi người dùng gửi logo thật, lưu đè vào `assets/logo/<ma>.png` và đổi `logo_tam` thành false trong hồ sơ công ty. Ảnh chụp màn hình thường có viền tối và nhiều nền thừa: cắt sát nội dung logo (chừa lề nhỏ) để logo hiện to, rõ. Script tự giữ đúng tỉ lệ ảnh (cao ~92px ở góc trên trái) và tự thu nhỏ chữ dòng địa chỉ nếu quá dài; vẫn nên xem lại PDF.
- Người dùng chê mẫu "chưa đẹp": hỏi cụ thể muốn đổi gì (màu, cỡ chữ, bố cục) rồi sửa trong lớp `BangBaoGia` của script, để mọi báo giá sau đều được hưởng, thay vì sửa tay một file Excel.
- Khi người dùng muốn **sửa báo giá vừa làm** (đổi giá, thêm hàng...), sửa file JSON rồi chạy lại script. Cách này nhanh hơn và ít lỗi hơn sửa trực tiếp file Excel.
