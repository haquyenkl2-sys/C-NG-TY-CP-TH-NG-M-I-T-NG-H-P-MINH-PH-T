#!/usr/bin/env python3
"""Điền mẫu báo giá Minh Phát từ một file JSON và xuất ra Excel (và PDF nếu có LibreOffice).

Cách dùng:
    python tao_bao_gia.py du_lieu.json --out Bao-gia-ABC.xlsx [--pdf]

Cấu trúc JSON xem trong SKILL.md.
"""
import argparse
import copy
import datetime as dt
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl

MAU = Path(__file__).resolve().parent.parent / "assets" / "Mau-bao-gia-Minh-Phat.xlsx"

DONG_HANG_DAU = 18
SO_DONG_MAU = 12  # dòng 18..29 trong mẫu
DONG_CUOI_MAU = 51

# Ô trên trang "Thông tin công ty"
O_CONG_TY = {
    "ten": "C4",
    "dia_chi": "C5",
    "mst": "C6",
    "dien_thoai": "C7",
    "email": "C8",
    "website": "C9",
    "so_tai_khoan": "C10",
    "ngan_hang": "C11",
    "nguoi_ky": "C12",
    "chuc_vu": "C13",
}

# ---------------------------------------------------------------------------
# Đọc số tiền bằng chữ tiếng Việt
# ---------------------------------------------------------------------------
CHU_SO = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]


def _doc_ba_so(n, day_du):
    """Đọc một nhóm 3 chữ số. day_du=True thì luôn đọc hàng trăm (kể cả 'không trăm')."""
    tram, chuc, dv = n // 100, (n // 10) % 10, n % 10
    tu = []
    if day_du or tram:
        tu += [CHU_SO[tram], "trăm"]
    if chuc == 0:
        if dv:
            if tu:
                tu.append("lẻ")
            tu.append(CHU_SO[dv])
    elif chuc == 1:
        tu.append("mười")
        if dv == 5:
            tu.append("lăm")
        elif dv:
            tu.append(CHU_SO[dv])
    else:
        tu += [CHU_SO[chuc], "mươi"]
        if dv == 1:
            tu.append("mốt")
        elif dv == 5:
            tu.append("lăm")
        elif dv:
            tu.append(CHU_SO[dv])
    return tu


def _doc_duoi_ty(n, day_du):
    """Đọc số 0..999 999 999."""
    nhom = [(n // 1_000_000) % 1000, (n // 1000) % 1000, n % 1000]
    don_vi = ["triệu", "nghìn", ""]
    tu = []
    for g, dvi in zip(nhom, don_vi):
        if g == 0:
            continue
        tu += _doc_ba_so(g, day_du or bool(tu))
        if dvi:
            tu.append(dvi)
    return tu


def _doc_so(n, day_du=False):
    if n >= 1_000_000_000:
        tren, duoi = divmod(n, 1_000_000_000)
        tu = _doc_so(tren, day_du) + ["tỷ"]
        if duoi:
            tu += _doc_duoi_ty(duoi, True)
        return tu
    return _doc_duoi_ty(n, day_du)


def so_thanh_chu(so_tien):
    n = int(round(so_tien))
    if n == 0:
        return "Không đồng."
    am = n < 0
    tu = _doc_so(abs(n))
    cau = " ".join(tu)
    if am:
        cau = "âm " + cau
    cau = cau[0].upper() + cau[1:] + " đồng"
    if n % 1000 == 0:
        cau += " chẵn"
    return cau + "."


# ---------------------------------------------------------------------------
# Thêm dòng khi có hơn 12 mặt hàng
# ---------------------------------------------------------------------------
def them_dong(ws, so_dong_them):
    if so_dong_them <= 0:
        return
    dau_khoi = DONG_HANG_DAU + SO_DONG_MAU  # dòng 30: "Cộng tiền hàng"
    cuoi_khoi = DONG_CUOI_MAU

    gop = [r for r in ws.merged_cells.ranges if r.min_row >= dau_khoi]
    toa_do_gop = [(r.min_col, r.min_row, r.max_col, r.max_row) for r in gop]
    for r in gop:
        ws.unmerge_cells(str(r))

    ws.move_range(f"A{dau_khoi}:G{cuoi_khoi}", rows=so_dong_them, translate=True)

    for c1, r1, c2, r2 in toa_do_gop:
        ws.merge_cells(start_row=r1 + so_dong_them, start_column=c1,
                       end_row=r2 + so_dong_them, end_column=c2)

    for r in range(cuoi_khoi, dau_khoi - 1, -1):
        ws.row_dimensions[r + so_dong_them].height = ws.row_dimensions[r].height

    dong_mau = dau_khoi - 1
    for r in range(dau_khoi, dau_khoi + so_dong_them):
        ws.row_dimensions[r].height = ws.row_dimensions[dong_mau].height
        for col in "ABCDEFG":
            nguon, dich = ws[f"{col}{dong_mau}"], ws[f"{col}{r}"]
            dich._style = copy.copy(nguon._style)
        ws[f"A{r}"] = f'=IF(B{r}="","",COUNTA($B${DONG_HANG_DAU}:B{r}))'
        ws[f"F{r}"] = f'=IF(OR(D{r}="",E{r}=""),"",D{r}*E{r})'

    dong_cong = dau_khoi + so_dong_them
    ws[f"F{dong_cong}"] = f"=SUM(F{DONG_HANG_DAU}:F{dong_cong - 1})"
    ws.print_area = f"A1:G{cuoi_khoi + so_dong_them}"
    if so_dong_them > 13:
        # Quá dài cho 1 trang: cho phép nhiều trang, lặp lại dòng tiêu đề bảng
        ws.page_setup.fitToHeight = 0
        ws.print_title_rows = "17:17"


# ---------------------------------------------------------------------------
def doc_ngay(gia_tri):
    if not gia_tri:
        return dt.datetime.combine(dt.date.today(), dt.time())
    for dinh_dang in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return dt.datetime.strptime(str(gia_tri), dinh_dang)
        except ValueError:
            pass
    raise SystemExit(f"Không hiểu ngày: {gia_tri!r} (dùng dạng 2026-10-10 hoặc 10/10/2026)")


def tao(du_lieu, duong_dan_ra):
    wb = openpyxl.load_workbook(MAU)
    ws = wb["Báo giá"]
    ws_ct = wb["Thông tin công ty"]

    # Thông tin công ty (chỉ ghi đè những trường được cung cấp)
    for khoa, gia_tri in (du_lieu.get("cong_ty") or {}).items():
        if khoa in O_CONG_TY and gia_tri:
            ws_ct[O_CONG_TY[khoa]] = gia_tri
    con_trong = [k for k, o in O_CONG_TY.items() if "[" in str(ws_ct[o].value or "")]

    hang = du_lieu.get("hang_hoa") or []
    if not hang:
        raise SystemExit("Cần ít nhất 1 mặt hàng trong 'hang_hoa'.")
    so_dong_them = max(0, len(hang) - SO_DONG_MAU)
    them_dong(ws, so_dong_them)
    d = so_dong_them  # độ lệch của phần dưới bảng

    # Khách hàng
    kh = du_lieu.get("khach_hang") or {}
    ws["C10"] = kh.get("ten") or ""
    ws["C11"] = kh.get("nguoi_lien_he") or ""
    ws["F11"] = kh.get("dien_thoai") or ""
    ws["C12"] = kh.get("dia_chi") or ""
    ws["C13"] = kh.get("email") or ""

    ngay = doc_ngay(du_lieu.get("ngay"))
    ws["F12"] = du_lieu.get("so_bao_gia") or f"BG-{ngay:%Y}/{ngay:%m%d}"
    ws["F13"] = ngay

    # Hàng hóa: xóa 2 dòng ví dụ rồi ghi hàng thật
    tong_hang = 0
    for i in range(SO_DONG_MAU + so_dong_them):
        r = DONG_HANG_DAU + i
        for col in "BCDEG":
            ws[f"{col}{r}"] = None
    for i, mh in enumerate(hang):
        r = DONG_HANG_DAU + i
        sl = mh.get("so_luong")
        dg = mh.get("don_gia")
        ws[f"B{r}"] = mh.get("ten") or ""
        ws[f"C{r}"] = mh.get("dvt") or ""
        ws[f"D{r}"] = sl
        ws[f"E{r}"] = dg
        ws[f"G{r}"] = mh.get("ghi_chu") or ""
        if sl is not None and dg is not None:
            tong_hang += sl * dg
    # Ẩn các dòng trống còn thừa cho bảng gọn (mở lại bằng Unhide nếu cần thêm hàng)
    for r in range(DONG_HANG_DAU + len(hang), DONG_HANG_DAU + SO_DONG_MAU):
        ws.row_dimensions[r].hidden = True

    # Thuế GTGT: nhận 10 hoặc 0.1
    vat = du_lieu.get("vat", 0.1)
    if vat is None:
        vat = 0.1
    vat = float(vat)
    if vat > 1:
        vat = vat / 100
    ws[f"E{31 + d}"] = vat
    tien_thue = round(tong_hang * vat)
    tong_cong = tong_hang + tien_thue

    ws[f"A{33 + d}"] = "Số tiền bằng chữ: " + so_thanh_chu(tong_cong)

    # Điều khoản (dòng 36..40 của mẫu)
    dk = du_lieu.get("dieu_khoan") or {}
    ws[f"A{36 + d}"] = "1. " + dk.get(
        "thue", "Đơn giá trên chưa bao gồm thuế GTGT; thuế GTGT được tính riêng như bảng trên.")
    ws[f"A{37 + d}"] = "2. Hiệu lực báo giá: " + dk.get("hieu_luc", "30 ngày kể từ ngày báo giá") + "."
    ws[f"A{38 + d}"] = "3. Thời gian giao hàng: " + dk.get(
        "giao_hang", "trong vòng 03 ngày kể từ khi Quý khách xác nhận đặt hàng") + "."
    ws[f"A{39 + d}"] = "4. Địa điểm giao hàng: " + dk.get("dia_diem", "tại kho của Quý khách") + "."
    ws[f"A{40 + d}"] = "5. Thanh toán: " + dk.get("thanh_toan", "chuyển khoản, 100% khi nhận hàng") + "."

    wb.calculation.fullCalcOnLoad = True
    duong_dan_ra = Path(duong_dan_ra)
    duong_dan_ra.parent.mkdir(parents=True, exist_ok=True)
    wb.save(duong_dan_ra)
    return {
        "file_excel": str(duong_dan_ra),
        "so_bao_gia": ws["F12"].value,
        "so_mat_hang": len(hang),
        "cong_tien_hang": tong_hang,
        "thue_suat": vat,
        "tien_thue": tien_thue,
        "tong_cong": tong_cong,
        "bang_chu": so_thanh_chu(tong_cong),
        "thong_tin_cong_ty_con_trong": con_trong,
    }


def xuat_pdf(file_excel):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    file_excel = Path(file_excel)
    pdf = file_excel.with_suffix(".pdf")
    with tempfile.TemporaryDirectory() as tam:
        # Bản tạm chỉ hiện trang "Báo giá" để PDF không in kèm trang hướng dẫn
        ban_tam = Path(tam) / (file_excel.stem + ".xlsx")
        wb = openpyxl.load_workbook(file_excel)
        for ws in wb.worksheets:
            if ws.title != "Báo giá":
                ws.sheet_state = "hidden"
        wb.active = wb.worksheets.index(wb["Báo giá"])
        wb.save(ban_tam)
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", tam, str(ban_tam)],
                       check=True, capture_output=True, timeout=180)
        pdf_tam = ban_tam.with_suffix(".pdf")
        if not pdf_tam.exists():
            return None
        shutil.move(str(pdf_tam), pdf)
    return str(pdf)


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("du_lieu", help="file JSON chứa thông tin báo giá")
    ap.add_argument("--out", required=True, help="đường dẫn file Excel đầu ra")
    ap.add_argument("--pdf", action="store_true", help="xuất thêm PDF (cần LibreOffice)")
    args = ap.parse_args()

    with open(args.du_lieu, encoding="utf-8") as f:
        du_lieu = json.load(f)
    ket_qua = tao(du_lieu, args.out)
    if args.pdf:
        ket_qua["file_pdf"] = xuat_pdf(ket_qua["file_excel"])
    json.dump(ket_qua, sys.stdout, ensure_ascii=False, indent=2)
    print()


if __name__ == "__main__":
    main()
