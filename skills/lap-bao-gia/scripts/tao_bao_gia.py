#!/usr/bin/env python3
"""Điền mẫu báo giá từ một file JSON và xuất ra Excel (và PDF nếu có LibreOffice).

Cách dùng:
    python tao_bao_gia.py du_lieu.json --out Bao-gia-ABC.xlsx [--cong-ty blt] [--pdf]
    python tao_bao_gia.py --mau blt --out Mau-bao-gia-BLT.xlsx [--pdf]   # chỉ tạo mẫu trống của công ty
    python tao_bao_gia.py --danh-sach                                     # liệt kê các công ty đã lưu

Cấu trúc JSON xem trong SKILL.md.
"""
import argparse
import copy
import datetime as dt
import json
import math
import shutil
import subprocess
import sys
import unicodedata
import tempfile
from pathlib import Path

import openpyxl
from openpyxl.drawing.image import Image as AnhExcel
from openpyxl.styles import Alignment

ASSETS = Path(__file__).resolve().parent.parent / "assets"
MAU = ASSETS / "Mau-bao-gia-goc.xlsx"
THU_MUC_CONG_TY = ASSETS / "cong-ty"
THU_MUC_LOGO = ASSETS / "logo"
CONG_TY_MAC_DINH = "minh-phat"

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
# Hồ sơ công ty (assets/cong-ty/*.json)
# ---------------------------------------------------------------------------
def _khong_dau(chu):
    chu = unicodedata.normalize("NFD", str(chu)).replace("đ", "d").replace("Đ", "D")
    return "".join(c for c in chu if unicodedata.category(c) != "Mn").lower().strip()


def danh_sach_cong_ty():
    return [json.loads(p.read_text(encoding="utf-8")) for p in sorted(THU_MUC_CONG_TY.glob("*.json"))]


def tim_cong_ty(ten):
    """Tìm hồ sơ theo mã (blt), tên gọi (BLT, Minh Phát) hoặc tên đầy đủ, không phân biệt dấu/hoa thường."""
    can_tim = _khong_dau(ten)
    for ct in danh_sach_cong_ty():
        cac_ten = [ct.get("ma", ""), ct.get("ten", "")] + list(ct.get("ten_goi", []))
        if can_tim in [_khong_dau(t) for t in cac_ten if t]:
            return ct
    ma = ", ".join(ct.get("ma", "") for ct in danh_sach_cong_ty())
    raise SystemExit(f"Không tìm thấy công ty {ten!r}. Các công ty đã lưu: {ma}")


def nap_cong_ty(gia_tri):
    """gia_tri: mã/tên công ty (chuỗi), hoặc dict có 'ma' (+ trường ghi đè), hoặc None (mặc định)."""
    if gia_tri is None:
        return tim_cong_ty(CONG_TY_MAC_DINH)
    if isinstance(gia_tri, str):
        return tim_cong_ty(gia_tri)
    ct = dict(tim_cong_ty(gia_tri.get("ma") or CONG_TY_MAC_DINH))
    ct.update({k: v for k, v in gia_tri.items() if v})
    return ct


def ap_dung_cong_ty(wb, ct):
    """Ghi thông tin công ty vào trang 'Thông tin công ty' và thay logo. Gọi trước khi chèn dòng."""
    ws = wb["Báo giá"]
    ws_ct = wb["Thông tin công ty"]
    for khoa, o in O_CONG_TY.items():
        if ct.get(khoa):
            ws_ct[o] = ct[khoa]
    if "website" in ct and not ct["website"]:
        # Không có website: đầu trang chỉ ghi email
        ws_ct["C9"] = ""
        ws["B4"] = "=\"Email: \"&'Thông tin công ty'!$C$8"
    if ct.get("chu_tai_khoan"):
        ws["A41"] = ("=\"6. Tài khoản: \"&'Thông tin công ty'!$C$10&\" – \"&'Thông tin công ty'!$C$11"
                     f"&\" – Chủ tài khoản: {ct['chu_tai_khoan']}\"")
    if ct.get("logo"):
        duong_dan_logo = THU_MUC_LOGO / ct["logo"]
        if duong_dan_logo.exists() and ws._images:
            anh = AnhExcel(str(duong_dan_logo))
            anh.anchor = ws._images[0].anchor
            ws._images = [anh]
    return [k for k, o in O_CONG_TY.items() if "[" in str(ws_ct[o].value or "")]


def tao_mau(ct, duong_dan_ra):
    """Tạo mẫu báo giá trống (giữ 2 dòng ví dụ) mang thông tin của một công ty."""
    wb = openpyxl.load_workbook(MAU)
    con_trong = ap_dung_cong_ty(wb, ct)
    wb.calculation.fullCalcOnLoad = True
    duong_dan_ra = Path(duong_dan_ra)
    duong_dan_ra.parent.mkdir(parents=True, exist_ok=True)
    wb.save(duong_dan_ra)
    return {"file_excel": str(duong_dan_ra), "cong_ty": ct.get("ten"),
            "logo_tam": bool(ct.get("logo_tam")), "thong_tin_cong_ty_con_trong": con_trong}

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


SO_HAN = "零壹贰叁肆伍陆柒捌玖"


def _doc_nhom_han(n):
    """Đọc nhóm 4 chữ số bằng chữ số viết hoa tiếng Trung (大写)."""
    tu, can_khong = "", False
    for i, dvi in zip(range(3, -1, -1), ["仟", "佰", "拾", ""]):
        so = (n // 10 ** i) % 10
        if so == 0:
            can_khong = bool(tu)
        else:
            if can_khong:
                tu += "零"
                can_khong = False
            tu += SO_HAN[so] + dvi
    return tu


def so_thanh_chu_trung(so_tien):
    """Số tiền bằng chữ tiếng Trung, ví dụ 11000000 -> 壹仟壹佰万越南盾整."""
    n = int(round(abs(so_tien)))
    if n == 0:
        return "零越南盾整"
    nhom = []
    while n:
        nhom.append(n % 10000)
        n //= 10000
    don_vi = ["", "万", "亿", "万亿"]
    tu = ""
    for i in range(len(nhom) - 1, -1, -1):
        g = nhom[i]
        if g == 0:
            continue
        if tu and (g < 1000 or nhom[i + 1] == 0):
            tu += "零"
        tu += _doc_nhom_han(g) + don_vi[i]
    return ("负" if so_tien < 0 else "") + tu + "越南盾整"


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


def _do_dai(chu):
    """Độ rộng hiển thị: chữ Hán/Trung chiếm gấp đôi chữ Latin."""
    return sum(1.7 if unicodedata.east_asian_width(c) in "WF" else 1 for c in chu)


def _so_dong_chu(chu, do_rong_cot):
    """Ước lượng số dòng khi ô tự xuống dòng (Arial 10)."""
    if not chu:
        return 1
    ky_tu_moi_dong = max(4, do_rong_cot * 1.15 - 1)
    tong = 0
    for doan in str(chu).split("\n"):
        # Ngắt theo từ như Excel: từ không vừa thì xuống dòng mới
        dong, dai = 1, 0
        for tu in doan.split():
            dai_tu = _do_dai(tu)
            if dai and dai + 1 + dai_tu > ky_tu_moi_dong:
                dong, dai = dong + 1, dai_tu
            else:
                dai += (1 if dai else 0) + dai_tu
            while dai > ky_tu_moi_dong:
                dong, dai = dong + 1, dai - ky_tu_moi_dong
        tong += dong
    return tong


def chinh_chieu_cao(ws, dong_dau, so_dong):
    """Cho dòng hàng cao lên vừa nội dung dài; trả về tổng chiều cao bảng hàng (pt)."""
    tong = 0
    for r in range(dong_dau, dong_dau + so_dong):
        so_dong_max = max(_so_dong_chu(ws[f"{col}{r}"].value, ws.column_dimensions[col].width or 10)
                          for col in "BCG")
        cao = max(19.5, 11.5 * so_dong_max + 6)
        ws.row_dimensions[r].height = cao
        tong += cao
    return tong


# ---------------------------------------------------------------------------
def _hai_dong(viet, trung):
    viet = viet or ""
    if not trung:
        return viet
    return f"{viet}\n{trung}" if viet else trung


def _xuong_dong(o):
    a = o.alignment
    o.alignment = Alignment(horizontal=a.horizontal, vertical=a.vertical, indent=a.indent,
                            text_rotation=a.text_rotation, shrink_to_fit=a.shrink_to_fit, wrap_text=True)


DK_ZH_MAC_DINH = {
    "thue": "以上单价不含增值税，增值税按上表另行计算。",
    "hieu_luc": "自报价之日起30天",
    "giao_hang": "自贵方确认订货之日起3天内",
    "dia_diem": "贵方仓库",
    "thanh_toan": "银行转账，收货时付清100%",
}
CHUC_VU_ZH = {"giám đốc": "总经理", "tổng giám đốc": "总经理", "phó giám đốc": "副总经理",
              "trưởng phòng kinh doanh": "销售经理", "kế toán trưởng": "会计主管"}


def ap_dung_song_ngu(ws, d, du_lieu, ct, tong_cong):
    """Đổi các nhãn của mẫu sang song ngữ Việt – Trung (d: số dòng đã chèn thêm)."""
    CT = "'Thông tin công ty'!"
    ws["B2"] = f'="Địa chỉ / 地址: "&{CT}$C$5'
    ws["B3"] = f'="MST / 税号: "&{CT}$C$6&"   |   Điện thoại / 电话: "&{CT}$C$7'
    if ct.get("website"):
        ws["B4"] = f'="Email / 邮箱: "&{CT}$C$8&"   |   Website / 网站: "&{CT}$C$9'
    else:
        ws["B4"] = f'="Email / 邮箱: "&{CT}$C$8'
    ws["A7"] = "BẢNG BÁO GIÁ / 报价单"
    ws["A8"] = '="Số / 编号: "&F12&"  –  Ngày / 日期: "&TEXT(F13,"dd/mm/yyyy")'
    for o, nhan in {"B10": "Kính gửi / 致:", "B11": "Người liên hệ / 联系人:", "B12": "Địa chỉ / 地址:",
                    "B13": "Email / 邮箱:", "E11": "Điện thoại\n电话:", "E12": "Số báo giá\n报价编号:",
                    "E13": "Ngày báo giá\n报价日期:"}.items():
        ws[o] = nhan
        _xuong_dong(ws[o])
    for r in range(10, 14):
        ws.row_dimensions[r].height = 27

    ten_ct = ct.get("ten", "")
    cong_trinh = du_lieu.get("cong_trinh")
    cong_trinh_zh = du_lieu.get("cong_trinh_zh")
    viet = (f"{ten_ct} trân trọng cảm ơn Quý khách đã quan tâm. Chúng tôi xin gửi tới Quý khách bảng báo giá "
            f"{cong_trinh} như sau:" if cong_trinh else
            f"{ten_ct} trân trọng cảm ơn Quý khách đã quan tâm đến sản phẩm của chúng tôi. "
            "Chúng tôi xin gửi tới Quý khách bảng báo giá như sau:")
    trung = (f"我司衷心感谢贵方的关注，现将{cong_trinh_zh}报价如下：" if cong_trinh_zh else
             "我司衷心感谢贵方对我司产品的关注，现将报价如下：")
    ws["A15"] = viet + "\n" + trung
    ws.row_dimensions[15].height = 13 * _so_dong_chu(ws["A15"].value, 112) + 6

    for o, nhan in {"A17": "STT\n序号", "B17": "Tên hàng hóa / Quy cách\n货物名称 / 规格", "C17": "ĐVT\n单位",
                    "D17": "Số lượng\n数量", "E17": "Đơn giá (VNĐ)\n单价(越南盾)",
                    "F17": "Thành tiền (VNĐ)\n金额(越南盾)", "G17": "Ghi chú\n备注"}.items():
        ws[o] = nhan
        _xuong_dong(ws[o])
    ws.row_dimensions[17].height = 34

    ws[f"A{30 + d}"] = "Cộng tiền hàng / 货款合计"
    ws[f"A{31 + d}"] = "Thuế GTGT (VAT) – thuế suất / 增值税税率:"
    ws[f"A{32 + d}"] = "TỔNG CỘNG THANH TOÁN / 付款总额"
    ws[f"A{33 + d}"] = ("Số tiền bằng chữ: " + so_thanh_chu(tong_cong) +
                        "\n大写金额：" + so_thanh_chu_trung(tong_cong))
    ws.row_dimensions[33 + d].height = 13 * _so_dong_chu(ws[f"A{33 + d}"].value, 112) + 6

    ws[f"A{35 + d}"] = "ĐIỀU KHOẢN BÁO GIÁ / 报价条款"
    dk_zh = {**DK_ZH_MAC_DINH, **(du_lieu.get("dieu_khoan_zh") or {})}
    nhan_zh = {"thue": "1. ", "hieu_luc": "2. 报价有效期：", "giao_hang": "3. 交货时间：",
               "dia_diem": "4. 交货地点：", "thanh_toan": "5. 付款方式："}
    for i, khoa in enumerate(["thue", "hieu_luc", "giao_hang", "dia_diem", "thanh_toan"]):
        o = ws[f"A{36 + d + i}"]
        o.value = f"{o.value}\n{nhan_zh[khoa]}{dk_zh[khoa].rstrip('。')}。"
        _xuong_dong(o)
    chu_tk = ct.get("chu_tai_khoan")
    chu_tk_ct = f'"{chu_tk}"' if chu_tk else f"{CT}$C$4"
    ws[f"A{41 + d}"] = (f'="6. Tài khoản / 银行账户: "&{CT}$C$10&" – "&{CT}$C$11'
                        f'&" – Chủ tài khoản / 户名: "&{chu_tk_ct}')

    ws[f"A{43 + d}"] = "Rất mong nhận được sự hợp tác của Quý khách. Trân trọng!\n期待与贵方合作，此致敬礼！"
    _xuong_dong(ws[f"A{43 + d}"])
    ws.row_dimensions[43 + d].height = 30
    chuc_vu = ct.get("chuc_vu") or "Giám đốc"
    chuc_vu_zh = ct.get("chuc_vu_zh") or CHUC_VU_ZH.get(chuc_vu.lower(), "")
    ws[f"A{45 + d}"] = "KHÁCH HÀNG XÁC NHẬN\n客户确认"
    ws[f"E{45 + d}"] = chuc_vu.upper() + (f"\n{chuc_vu_zh}" if chuc_vu_zh else "")
    ws[f"A{46 + d}"] = "(Ký, ghi rõ họ tên / 签字并注明姓名)"
    ws[f"E{46 + d}"] = "(Ký, đóng dấu, ghi rõ họ tên / 签字盖章并注明姓名)"
    for o in (f"A{45 + d}", f"E{45 + d}"):
        _xuong_dong(ws[o])
    ws.row_dimensions[45 + d].height = 30


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

    ct = nap_cong_ty(du_lieu.get("cong_ty"))
    con_trong = ap_dung_cong_ty(wb, ct)

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
    song_ngu = bool(du_lieu.get("song_ngu"))
    if du_lieu.get("cong_trinh"):
        # Báo giá cho một công trình/dự án: nêu rõ trong lời mở đầu
        ws["A15"] = ("='Thông tin công ty'!$C$4&\" trân trọng cảm ơn Quý khách đã quan tâm. "
                     f"Chúng tôi xin gửi tới Quý khách bảng báo giá {du_lieu['cong_trinh']} như sau:\"")
        ws.row_dimensions[15].height = max(30, 13 * _so_dong_chu(
            "x" * 60 + du_lieu["cong_trinh"] + "x" * 80, 112) + 4)
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
        ws[f"B{r}"] = _hai_dong(mh.get("ten"), mh.get("ten_zh") if song_ngu else None)
        ws[f"C{r}"] = _hai_dong(mh.get("dvt"), mh.get("dvt_zh") if song_ngu else None)
        ws[f"D{r}"] = sl
        ws[f"E{r}"] = dg
        ws[f"G{r}"] = _hai_dong(mh.get("ghi_chu"), mh.get("ghi_chu_zh") if song_ngu else None)
        if isinstance(sl, (int, float)):
            ws[f"D{r}"].number_format = "#,##0" if float(sl).is_integer() else "#,##0.##"
        if sl is not None and dg is not None:
            tong_hang += sl * dg
    cao_bang = chinh_chieu_cao(ws, DONG_HANG_DAU, len(hang))
    if cao_bang > 330:
        # Quá dài cho 1 trang A4: cho phép nhiều trang, lặp lại dòng tiêu đề bảng
        ws.page_setup.fitToHeight = 0
        ws.print_title_rows = "17:17"
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
    rong_ca_trang = sum(ws.column_dimensions[c].width or 10 for c in "ABCDEFG")
    if song_ngu:
        ap_dung_song_ngu(ws, d, du_lieu, ct, tong_cong)
    for r in range(36 + d, 41 + d):
        ws.row_dimensions[r].height = max(18, 13 * _so_dong_chu(ws[f"A{r}"].value, rong_ca_trang) + 4)

    if du_lieu.get("so_trang"):
        # Thu nhỏ vừa đúng số trang mong muốn (khi trang cuối chỉ còn vài dòng)
        ws.page_setup.fitToHeight = int(du_lieu["so_trang"])

    wb.calculation.fullCalcOnLoad = True
    duong_dan_ra = Path(duong_dan_ra)
    duong_dan_ra.parent.mkdir(parents=True, exist_ok=True)
    wb.save(duong_dan_ra)
    return {
        "file_excel": str(duong_dan_ra),
        "cong_ty": ct.get("ten"),
        "logo_tam": bool(ct.get("logo_tam")),
        "so_bao_gia": ws["F12"].value,
        "so_mat_hang": len(hang),
        "cong_tien_hang": tong_hang,
        "thue_suat": vat,
        "tien_thue": tien_thue,
        "tong_cong": tong_cong,
        "bang_chu": so_thanh_chu(tong_cong),
        "bang_chu_zh": so_thanh_chu_trung(tong_cong) if song_ngu else None,
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
    ap.add_argument("du_lieu", nargs="?", help="file JSON chứa thông tin báo giá")
    ap.add_argument("--out", help="đường dẫn file Excel đầu ra")
    ap.add_argument("--cong-ty", help="mã hoặc tên công ty báo giá (ghi đè trường cong_ty trong JSON)")
    ap.add_argument("--mau", metavar="CONG_TY", help="chỉ tạo mẫu báo giá trống cho công ty này")
    ap.add_argument("--danh-sach", action="store_true", help="liệt kê các công ty đã lưu")
    ap.add_argument("--pdf", action="store_true", help="xuất thêm PDF (cần LibreOffice)")
    args = ap.parse_args()

    if args.danh_sach:
        for ct in danh_sach_cong_ty():
            print(f"{ct.get('ma')}: {ct.get('ten')}  (gọi là: {', '.join(ct.get('ten_goi', []))})")
        return
    if not args.out:
        ap.error("cần --out")
    if args.mau:
        ket_qua = tao_mau(tim_cong_ty(args.mau), args.out)
    else:
        if not args.du_lieu:
            ap.error("cần file JSON dữ liệu báo giá (hoặc dùng --mau)")
        with open(args.du_lieu, encoding="utf-8") as f:
            du_lieu = json.load(f)
        if args.cong_ty:
            du_lieu["cong_ty"] = args.cong_ty
        ket_qua = tao(du_lieu, args.out)
    if args.pdf:
        ket_qua["file_pdf"] = xuat_pdf(ket_qua["file_excel"])
    json.dump(ket_qua, sys.stdout, ensure_ascii=False, indent=2)
    print()


if __name__ == "__main__":
    main()
