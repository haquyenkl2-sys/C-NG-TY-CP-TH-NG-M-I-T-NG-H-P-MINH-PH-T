# Trợ lý soạn văn bản – Công ty BLT (xây dựng – thi công)

Trợ lý giúp soạn nhanh các văn bản thường dùng của BLT, đúng văn phong công ty, tự tính tiền và đọc số tiền bằng chữ.

## Trợ lý soạn được gì

| # | Văn bản | File mẫu |
|---|---|---|
| 01 | Thư giới thiệu năng lực / chào thầu | `mau/01-thu-gioi-thieu-nang-luc.md` |
| 02 | Email gửi báo giá thi công | `mau/02-email-gui-bao-gia.md` |
| 03 | Tin nhắn Zalo (hẹn khảo sát, gửi giá, báo tiến độ, mời nghiệm thu, nhắc tiền, báo phát sinh) | `mau/03-tin-nhan-zalo.md` |
| 04 | Hợp đồng thi công xây dựng | `mau/04-hop-dong-thi-cong.md` |
| 05 | Biên bản nghiệm thu | `mau/05-bien-ban-nghiem-thu.md` |
| 06 | Biên bản xác nhận khối lượng phát sinh | `mau/06-bien-ban-phat-sinh.md` |
| 07 | Giấy đề nghị thanh toán đợt | `mau/07-de-nghi-thanh-toan.md` |
| 08 | Email / công văn nhắc thanh toán | `mau/08-cong-van-nhac-thanh-toan.md` |
| 09 | Biên bản bàn giao công trình | `mau/09-bien-ban-ban-giao.md` |

Ngoài các mẫu trên, có thể nhờ soạn bất kỳ văn bản nào khác — trợ lý sẽ theo văn phong của công ty.

## Bước 1 – Điền thông tin công ty (làm 1 lần)

Mở `thong-tin-cong-ty.md`, thay toàn bộ phần trong `[ngoặc vuông]` bằng thông tin thật (tên đầy đủ, địa chỉ, MST, tài khoản, người đại diện, dịch vụ, công trình tiêu biểu…). Chỗ nào còn để trống, trợ lý sẽ giữ nguyên `[...]` chứ không tự bịa.

## Bước 2 – Chọn cách dùng

### Cách A: Dùng trên Claude.ai (web / điện thoại) — dễ nhất

1. Vào claude.ai → **Projects** → **Create project**, đặt tên *Trợ lý BLT*.
2. Mục **Instructions**: dán nội dung trong khung của file `chi-dan-claude-ai.md`.
3. Mục **Project knowledge / Files**: tải lên `thong-tin-cong-ty.md`, `phong-cach-viet.md` và 9 file trong thư mục `mau/`.
4. Mở chat mới trong Project và nhờ soạn, ví dụ:
   - *"Soạn tin Zalo hẹn anh Tuấn khảo sát nhà ở 12 Lê Lợi, 9h sáng thứ Bảy 11/10."*
   - *"Soạn hợp đồng thi công phần thô nhà 3 tầng cho chị Lan, giá 1.250.000.000 đồng chưa VAT, tạm ứng 30%, làm trong 120 ngày."*
   - *"Viết công văn nhắc thanh toán lần 2, đợt 2 hợp đồng 05/2026/HĐTC/BLT, 320 triệu, quá hạn 15 ngày."*

> Khi đổi thông tin công ty, nhớ tải lại `thong-tin-cong-ty.md` lên Project.

### Cách B: Dùng trong Claude Code (với kho này)

Kho đã có sẵn skill `tro-ly-blt` (`.claude/skills/tro-ly-blt/SKILL.md`). Mở Claude Code trong kho này rồi nhờ soạn như trên (hoặc gõ `/tro-ly-blt`). Văn bản dài được lưu vào `tro-ly-blt/van-ban/`.

## Công cụ đọc số tiền bằng chữ

```
python3 doc-so-tien.py 1250000000
→ Một tỷ hai trăm năm mươi triệu đồng chẵn.
```

## Lưu ý

- Các mẫu hợp đồng, biên bản là **mẫu tham khảo** cho công trình dân dụng, nhà xưởng quy mô vừa và nhỏ. Công trình vốn nhà nước hoặc giá trị lớn cần theo quy định riêng — luôn kiểm tra lại với kế toán / luật sư trước khi ký.
- Muốn sửa văn phong (cách xưng hô, chữ ký, điều kiện bảo hành…): sửa `phong-cach-viet.md` hoặc `thong-tin-cong-ty.md`, trợ lý sẽ làm theo.
