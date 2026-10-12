#!/usr/bin/env python3
"""Lập bảng báo giá Excel trình bày chuyên nghiệp từ một file JSON (và xuất PDF nếu có LibreOffice).

Cách dùng:
    python tao_bao_gia.py du_lieu.json --out Bao-gia-ABC.xlsx [--cong-ty blt] [--pdf]
    python tao_bao_gia.py --mau blt --out Mau-bao-gia-BLT.xlsx [--pdf]   # mẫu trống để tự điền trong Excel
    python tao_bao_gia.py --danh-sach                                     # liệt kê các công ty đã lưu

Cấu trúc JSON xem trong SKILL.md.
"""
import argparse
import datetime as dt
import json
import shutil
import subprocess
import sys
import tempfile
import unicodedata
from pathlib import Path

import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.drawing.image import Image as AnhExcel
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.worksheet.pagebreak import Break

ASSETS = Path(__file__).resolve().parent.parent / "assets"
THU_MUC_CONG_TY = ASSETS / "cong-ty"
THU_MUC_LOGO = ASSETS / "logo"
CONG_TY_MAC_DINH = "minh-phat"

PHONG = "Arial"
DO_RONG_COT = {"A": 5.5, "B": 40, "C": 8, "D": 9, "E": 14.5, "F": 16.5, "G": 14}
MAU_MAC_DINH = {"chinh": "1F3864", "nhan": "C9A227"}  # xanh đậm + vàng đồng
SO_DONG_TRONG_MAU = 15  # số dòng hàng trống khi tạo mẫu (--mau)

# Nhãn song ngữ (tiếng Việt, tiếng Trung)
NHAN = {
    "tieu_de": ("BẢNG BÁO GIÁ", "报价单"),
    "khach_hang": ("THÔNG TIN KHÁCH HÀNG", "客户信息"),
    "bao_gia": ("THÔNG TIN BÁO GIÁ", "报价信息"),
    "kinh_gui": ("Kính gửi", "致"),
    "nguoi_lien_he": ("Người liên hệ", "联系人"),
    "dien_thoai": ("Điện thoại", "电话"),
    "email": ("Email", "邮箱"),
    "dia_chi": ("Địa chỉ", "地址"),
    "cong_trinh": ("Công trình", "工程"),
    "so": ("Số báo giá", "报价编号"),
    "ngay": ("Ngày báo giá", "报价日期"),
    "hieu_luc": ("Hiệu lực", "有效期"),
    "phu_trach": ("Người phụ trách", "负责人"),
    "stt": ("STT", "序号"),
    "ten": ("Tên hàng hóa / Quy cách kỹ thuật", "货物名称 / 技术规格"),
    "dvt": ("ĐVT", "单位"),
    "sl": ("Số lượng", "数量"),
    "don_gia": ("Đơn giá\n(VNĐ)", "单价(越南盾)"),
    "thanh_tien": ("Thành tiền\n(VNĐ)", "金额(越南盾)"),
    "ghi_chu": ("Ghi chú", "备注"),
    "cong": ("Cộng tiền hàng (chưa VAT)", "货款合计（不含税）"),
    "cong_gom_vat": ("Cộng tiền hàng (đã bao gồm thuế GTGT {p})", "货款合计（含增值税{p}）"),
    "vat": ("Thuế GTGT (VAT)", "增值税"),
    "tong": ("TỔNG CỘNG THANH TOÁN", "付款总额"),
    "bang_chu": ("Bằng chữ", "大写金额"),
    "dieu_khoan": ("ĐIỀU KHOẢN THƯƠNG MẠI", "商务条款"),
    "xac_nhan": ("XÁC NHẬN CỦA KHÁCH HÀNG", "客户确认"),
    "ky_kh": ("(Ký, ghi rõ họ tên)", "（签字并注明姓名）"),
    "ky_ct": ("(Ký, đóng dấu, ghi rõ họ tên)", "（签字盖章并注明姓名）"),
    "loi_ket": ("Rất mong nhận được sự hợp tác của Quý khách. Trân trọng!", "期待与贵方合作，此致敬礼！"),
}
DIEU_KHOAN = [  # (khóa, nhãn Việt, nhãn Trung, mặc định Việt, mặc định Trung)
    ("thue", "", "", "Đơn giá trên chưa bao gồm thuế GTGT; thuế GTGT được tính riêng như bảng trên",
     "以上单价不含增值税，增值税按上表另行计算"),
    ("hieu_luc", "Hiệu lực báo giá: ", "报价有效期：", "{n} ngày kể từ ngày báo giá", "自报价之日起{n}天"),
    ("giao_hang", "Thời gian giao hàng: ", "交货时间：",
     "trong vòng 03 ngày kể từ khi Quý khách xác nhận đặt hàng", "自贵方确认订货之日起3天内"),
    ("dia_diem", "Địa điểm giao hàng: ", "交货地点：", "tại kho của Quý khách", "贵方仓库"),
    ("thanh_toan", "Thanh toán: ", "付款方式：", "chuyển khoản, 100% khi nhận hàng", "银行转账，收货时付清100%"),
]
CHUC_VU_ZH = {"giám đốc": "总经理", "tổng giám đốc": "总经理", "phó giám đốc": "副总经理",
              "trưởng phòng kinh doanh": "销售经理", "kế toán trưởng": "会计主管"}


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


def thong_tin_con_trong(ct):
    """Các trường hồ sơ công ty còn để trống hoặc còn chữ mẫu trong ngoặc vuông."""
    can = ["dia_chi", "mst", "dien_thoai", "email", "so_tai_khoan", "ngan_hang", "nguoi_ky"]
    return [k for k in can if not ct.get(k) or "[" in str(ct.get(k))]


# ---------------------------------------------------------------------------
# Đọc số tiền bằng chữ
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
    tu = []
    for g, dvi in zip(nhom, ["triệu", "nghìn", ""]):
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
    cau = " ".join(_doc_so(abs(n)))
    if n < 0:
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


def so_la_ma(n):
    bang = [(10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]
    kq = ""
    for gt, ky in bang:
        while n >= gt:
            kq, n = kq + ky, n - gt
    return kq


# ---------------------------------------------------------------------------
# Ước lượng chiều cao dòng khi chữ tự xuống dòng
# ---------------------------------------------------------------------------
def _do_dai(chu):
    """Độ rộng hiển thị: chữ Hán/Trung chiếm gấp ~1,7 chữ Latin."""
    return sum(1.7 if unicodedata.east_asian_width(c) in "WF" else 1 for c in chu)


def _so_dong_chu(chu, do_rong_cot, co_chu=10):
    """Ước lượng số dòng của ô tự xuống dòng (Arial, cỡ co_chu)."""
    if not chu:
        return 1
    ky_tu_moi_dong = max(4, do_rong_cot * 1.15 * 10 / co_chu - 1)
    tong = 0
    for doan in str(chu).split("\n"):
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


def _rong(c1, c2):
    cot = "ABCDEFG"
    return sum(DO_RONG_COT[c] for c in cot[cot.index(c1):cot.index(c2) + 1])


# ---------------------------------------------------------------------------
# Vẽ bảng báo giá
# ---------------------------------------------------------------------------
class BangBaoGia:
    def __init__(self, ws, ct, song_ngu=False):
        self.ws, self.ct, self.song_ngu = ws, ct, song_ngu
        mau = {**MAU_MAC_DINH, **(ct.get("mau") or {})}
        self.chinh, self.mau_nhan = mau["chinh"], mau["nhan"]
        self.nen_nhat = mau.get("nen_nhat", "EEF2F8")
        self.nen_xen = mau.get("nen_xen", "F7F9FC")
        self.vien = Side(style="thin", color="C9D1DC")
        self.r = 1
        self.khoi = []  # các khối (dòng đầu, dòng cuối) không nên bị cắt ngang khi sang trang
        for cot, rong in DO_RONG_COT.items():
            ws.column_dimensions[cot].width = rong
        ws.sheet_view.showGridLines = False

    # --- tiện ích ---
    def nhan(self, khoa, noi=" / "):
        viet, trung = NHAN[khoa]
        return f"{viet}{noi}{trung}" if self.song_ngu else viet

    def font(self, co=10, dam=False, nghieng=False, mau="000000"):
        return Font(name=PHONG, size=co, bold=dam, italic=nghieng, color=mau)

    def o(self, vung, gia_tri=None, font=None, can="left", doc="center", nen=None, vien=None,
          xuong_dong=True, thut=0, dinh_dang=None):
        """Ghi một ô hoặc một vùng gộp (vd 'A5:D5'), áp kiểu cho mọi ô trong vùng."""
        ws = self.ws
        if ":" in vung:
            ws.merge_cells(vung)
        dau = vung.split(":")[0]
        if gia_tri is not None:
            ws[dau] = gia_tri
        for hang in ws[vung] if ":" in vung else [[ws[vung]]]:
            for x in hang:
                x.alignment = Alignment(horizontal=can, vertical=doc, wrap_text=xuong_dong, indent=thut)
                if font is not None:
                    x.font = font
                if nen:
                    x.fill = PatternFill("solid", fgColor=nen)
                if vien is not None:
                    x.border = vien
                if dinh_dang:
                    x.number_format = dinh_dang
        return ws[dau]

    def cao(self, chieu_cao, dong=None):
        self.ws.row_dimensions[dong or self.r].height = chieu_cao

    def cach(self, chieu_cao=8):
        self.cao(chieu_cao)
        self.r += 1

    def chu_giau(self, cap):
        """cap: [(chữ, InlineFont)] -> CellRichText."""
        return CellRichText([TextBlock(f, t) for t, f in cap if t])

    def if_nhan(self, dam=False):
        return InlineFont(rFont=PHONG, sz=9.5, color="595959", b=dam)

    def if_gia_tri(self, dam=False, co=9.5):
        return InlineFont(rFont=PHONG, sz=co, color="000000", b=dam)

    # --- các phần ---
    def dau_trang(self):
        ws, ct = self.ws, self.ct
        logo = THU_MUC_LOGO / ct["logo"] if ct.get("logo") else None
        if logo and logo.exists():
            anh = AnhExcel(str(logo))
            cao = 92
            rong = round(cao * anh.width / anh.height)
            anh.anchor = OneCellAnchor(
                _from=AnchorMarker(col=0, colOff=pixels_to_EMU(2), row=0, rowOff=pixels_to_EMU(3)),
                ext=XDRPositiveSize2D(pixels_to_EMU(rong), pixels_to_EMU(cao)))
            ws.add_image(anh)

        ten = ct.get("ten", "")
        co_ten = 16 if _do_dai(ten) <= 48 else max(12, round(16 * 48 / _do_dai(ten)))
        self.o("B1:G1", ten, self.font(co_ten, dam=True, mau=self.chinh), can="right", xuong_dong=False)
        self.cao(28)
        dong = []
        dong.append(f"{self.nhan('dia_chi')}: {ct.get('dia_chi', '')}")
        phan = [f"MST{' / 税号' if self.song_ngu else ''}: {ct.get('mst', '')}",
                f"{self.nhan('dien_thoai')}: {ct.get('dien_thoai', '')}"]
        dong.append("   |   ".join(phan))
        phan = [f"{self.nhan('email')}: {ct.get('email', '')}"] if ct.get("email") else []
        if ct.get("website"):
            phan.append(f"Website: {ct['website']}")
        dong.append("   |   ".join(phan))
        dai_nhat = max(_do_dai(d) for d in dong)
        co = 9 if dai_nhat <= 110 else max(7.5, round(9 * 110 / dai_nhat, 1))
        for i, chu in enumerate(dong):
            self.o(f"B{2 + i}:G{2 + i}", chu, self.font(co, mau="404040"), can="right", xuong_dong=False)
            self.cao(15, 2 + i)
        self.cao(10, 5)
        self.cao(4, 6)
        for c in "ABCDEFG":
            ws[f"{c}5"].border = Border(bottom=Side(style="medium", color=self.chinh))
            ws[f"{c}6"].border = Border(bottom=Side(style="thin", color=self.mau_nhan))
        self.r = 7

    def tieu_de(self):
        self.cach(10)
        viet, trung = NHAN["tieu_de"]
        self.o(f"A{self.r}:G{self.r}", viet, self.font(22, dam=True, mau=self.chinh), can="center",
               xuong_dong=False)
        self.cao(34)
        self.r += 1
        phu = trung if self.song_ngu else "QUOTATION"
        self.o(f"A{self.r}:G{self.r}", phu, self.font(11 if self.song_ngu else 9, mau="7F7F7F",
                                                         dam=self.song_ngu), can="center", xuong_dong=False)
        self.cao(16)
        self.r += 1
        self.cach(10)

    def hai_khung(self, trai, phai):
        """trai/phai: danh sách (nhãn, giá trị, đậm?). Khung khách hàng A:C, khung báo giá E:G."""
        r0 = self.r
        trang = self.font(9.5, dam=True, mau="FFFFFF")
        self.o(f"A{r0}:C{r0}", self.nhan("khach_hang"), trang, nen=self.chinh, thut=1)
        self.o(f"E{r0}:G{r0}", self.nhan("bao_gia"), trang, nen=self.chinh, thut=1)
        self.cao(20, r0)
        so_dong = max(len(trai), len(phai))
        for i in range(so_dong):
            r = r0 + 1 + i
            cao = 17
            for (vung, ds, rong) in ((f"A{r}:C{r}", trai, _rong("A", "C")), (f"E{r}:G{r}", phai, _rong("E", "G"))):
                if i < len(ds):
                    nhan, gia_tri, dam = ds[i]
                    noi_dung = self.chu_giau([(f"{nhan}: ", self.if_nhan()),
                                              (gia_tri or "", self.if_gia_tri(dam))])
                    cao = max(cao, 12.5 * _so_dong_chu(f"{nhan}: {gia_tri}", rong - 2, 9.5) + 5)
                else:
                    noi_dung = None
                self.o(vung, noi_dung, nen=self.nen_xen, thut=1)
            self.cao(cao, r)
        r_cuoi = r0 + so_dong
        # Viền ngoài cho hai khung
        for c1, c2 in (("A", "C"), ("E", "G")):
            cot = "ABCDEFG"[ "ABCDEFG".index(c1): "ABCDEFG".index(c2) + 1]
            for r in range(r0, r_cuoi + 1):
                for c in cot:
                    x = self.ws[f"{c}{r}"]
                    x.border = Border(left=self.vien if c == c1 else None, right=self.vien if c == c2 else None,
                                      top=self.vien if r == r0 else None, bottom=self.vien if r == r_cuoi else None)
        self.r = r_cuoi + 1
        self.cach(10)

    def dong_cong_trinh(self, ten, ten_zh=None):
        """Dải 'Công trình' chạy hết chiều ngang, dưới hai khung thông tin."""
        r = self.r
        ten = ten[0].upper() + ten[1:]
        cap = [(f"{self.nhan('cong_trinh').upper()}:  ", InlineFont(rFont=PHONG, sz=9.5, b=True, color=self.chinh)),
               (ten, self.if_gia_tri(True))]
        chu = f"CÔNG TRÌNH: {ten}"
        if self.song_ngu and ten_zh:
            cap.append((f"\n{ten_zh}", self.if_gia_tri(False)))
            chu += f"\n{ten_zh}"
        vien = Border(left=Side(style="thick", color=self.mau_nhan))
        self.o(f"A{r}:G{r}", self.chu_giau(cap), nen=self.nen_nhat, thut=1, vien=vien)
        self.cao(max(22, 13 * _so_dong_chu(chu, _rong("A", "G") - 3, 9.5) + 8))
        self.r += 1
        self.cach(10)

    def loi_mo_dau(self, cong_trinh=None, cong_trinh_zh=None):
        ten = self.ct.get("ten", "")
        viet = (f"{ten} trân trọng cảm ơn Quý khách đã quan tâm. Chúng tôi xin gửi tới Quý khách bảng báo giá "
                "cho công trình nêu trên như sau:" if cong_trinh else
                f"{ten} trân trọng cảm ơn Quý khách đã quan tâm đến sản phẩm, dịch vụ của chúng tôi. "
                "Chúng tôi xin gửi tới Quý khách bảng báo giá như sau:")
        chu = viet
        if self.song_ngu:
            chu += "\n" + ("我司衷心感谢贵方的关注，现将上述工程报价如下：" if cong_trinh else
                           "我司衷心感谢贵方对我司产品及服务的关注，现将报价如下：")
        self.o(f"A{self.r}:G{self.r}", chu, self.font(10, nghieng=True, mau="262626"))
        self.cao(13 * _so_dong_chu(chu, _rong("A", "G")) + 5)
        self.r += 1
        self.cach(8)

    def tieu_de_bang(self):
        r = self.r
        trang = self.font(9.5, dam=True, mau="FFFFFF")
        vien_trang = Border(left=Side(style="thin", color="FFFFFF"), right=Side(style="thin", color="FFFFFF"))
        for c, khoa in zip("ABCDEFG", ["stt", "ten", "dvt", "sl", "don_gia", "thanh_tien", "ghi_chu"]):
            viet, trung = NHAN[khoa]
            chu = f"{viet.replace(chr(10), ' ')}\n{trung}" if self.song_ngu else viet
            self.o(f"{c}{r}", chu, trang, can="center", nen=self.chinh, vien=vien_trang)
        self.cao(36 if self.song_ngu else 32)
        self.dong_tieu_de_bang = r
        self.r += 1

    def dong_nhom(self, so_thu_tu, ten, ten_zh=None):
        r = self.r
        chu = f"{ten}\n{ten_zh}" if (self.song_ngu and ten_zh) else ten
        f = self.font(9.5, dam=True, mau=self.chinh)
        vien = Border(top=self.vien, bottom=self.vien, left=self.vien, right=self.vien)
        self.o(f"A{r}", so_la_ma(so_thu_tu), f, can="center", nen=self.nen_nhat, vien=vien)
        self.o(f"B{r}:G{r}", chu.upper() if not self.song_ngu else chu, f, nen=self.nen_nhat, vien=vien)
        self.cao(max(18, 12.5 * _so_dong_chu(chu, _rong("B", "G"), 9.5) + 5))
        self.khoi.append((r, r + 1))  # tiêu đề nhóm luôn đi cùng dòng hàng đầu tiên của nhóm
        self.r += 1

    def dong_hang(self, stt, mh, xen):
        """Ghi một dòng hàng. mh=None: dòng trống (mẫu) vẫn có công thức thành tiền."""
        r = self.r
        nen = self.nen_xen if xen else "FFFFFF"
        vien = Border(top=self.vien, bottom=self.vien, left=self.vien, right=self.vien)
        mh = mh or {}

        def hai(viet, trung):
            viet = viet or ""
            if not (self.song_ngu and trung):
                return viet
            return f"{viet}\n{trung}" if viet else trung

        ten = hai(mh.get("ten"), mh.get("ten_zh"))
        dvt = hai(mh.get("dvt"), mh.get("dvt_zh"))
        ghi_chu = hai(mh.get("ghi_chu"), mh.get("ghi_chu_zh"))
        sl, dg = mh.get("so_luong"), mh.get("don_gia")
        so = self.font(9.5)
        self.o(f"A{r}", stt if mh else None, so, can="center", nen=nen, vien=vien)
        self.o(f"B{r}", ten, so, nen=nen, vien=vien)
        self.o(f"C{r}", dvt, so, can="center", nen=nen, vien=vien)
        dd_sl = "#,##0" if sl is None or float(sl).is_integer() else "#,##0.##"
        self.o(f"D{r}", sl, so, can="right", nen=nen, vien=vien, dinh_dang=dd_sl)
        self.o(f"E{r}", dg, so, can="right", nen=nen, vien=vien, dinh_dang="#,##0")
        self.o(f"F{r}", f'=IF(OR(D{r}="",E{r}=""),"",ROUND(D{r}*E{r},0))', self.font(9.5, dam=True),
               can="right", nen=nen, vien=vien, dinh_dang="#,##0")
        self.o(f"G{r}", ghi_chu, self.font(8.5, mau="404040"), nen=nen, vien=vien)
        so_dong = max(_so_dong_chu(ten, DO_RONG_COT["B"], 9.5), _so_dong_chu(dvt, DO_RONG_COT["C"], 9.5),
                      _so_dong_chu(ghi_chu, DO_RONG_COT["G"], 8.5))
        self.cao(max(19, 12 * so_dong + 6))
        self.r += 1
        return r

    def tong_tien(self, dong_dau, dong_cuoi, vat, tong_cong=None, da_gom_vat=False):
        """da_gom_vat=True: đơn giá đã gồm thuế GTGT, nên không cộng thêm dòng thuế."""
        ws = self.ws
        tren = Side(style="medium", color=self.chinh)
        r1 = self.r
        dam = self.font(10, dam=True)
        nhan_cong = self.nhan("cong_gom_vat" if da_gom_vat else "cong").replace("{p}", f"{vat:.0%}")
        self.o(f"A{r1}:E{r1}", nhan_cong, dam, can="right", vien=Border(top=tren))
        self.o(f"F{r1}", f"=SUM(F{dong_dau}:F{dong_cuoi})", dam, can="right", vien=Border(top=tren),
               dinh_dang="#,##0")
        self.o(f"G{r1}", None, vien=Border(top=tren))
        self.cao(20, r1)
        if da_gom_vat:
            r2 = None
            r3 = r1 + 1
            cong_thuc_tong = f"=F{r1}"
        else:
            r2 = r1 + 1
            self.o(f"A{r2}:D{r2}", self.nhan("vat"), self.font(10), can="right")
            self.o(f"E{r2}", vat, self.font(10), can="center", dinh_dang="0%")
            self.o(f"F{r2}", f"=ROUND(F{r1}*E{r2},0)", self.font(10), can="right", dinh_dang="#,##0")
            self.cao(20, r2)
            r3 = r2 + 1
            cong_thuc_tong = f"=F{r1}+F{r2}"
        trang = self.font(11.5, dam=True, mau="FFFFFF")
        self.o(f"A{r3}:E{r3}", self.nhan("tong"), trang, can="right", nen=self.chinh)
        self.o(f"F{r3}", cong_thuc_tong, trang, can="right", nen=self.chinh, dinh_dang="#,##0")
        self.o(f"G{r3}", None, nen=self.chinh)
        self.cao(24, r3)
        r4 = r3 + 1
        if tong_cong is None:
            chu = f"{self.nhan('bang_chu')}: ...................................................................."
        else:
            chu = f"{NHAN['bang_chu'][0]}: {so_thanh_chu(tong_cong)}"
            if self.song_ngu:
                chu += f"\n{NHAN['bang_chu'][1]}：{so_thanh_chu_trung(tong_cong)}"
        self.o(f"A{r4}:G{r4}", chu, self.font(10, dam=True, nghieng=True, mau=self.chinh), nen=self.nen_nhat,
               thut=1, vien=Border(bottom=Side(style="thin", color=self.mau_nhan)))
        self.cao(max(22, 13.5 * _so_dong_chu(chu, _rong("A", "G") - 3) + 7), r4)
        self.khoi.append((r1, r4))
        self.r = r4 + 1
        self.cach(12)
        return r1, r2, r3

    def dieu_khoan(self, dk, dk_zh, hieu_luc_ngay):
        ws = self.ws
        self.khoi.append((self.r, self.r + 1))
        self.o(f"A{self.r}:G{self.r}", self.nhan("dieu_khoan"), self.font(10.5, dam=True, mau=self.chinh),
               vien=Border(bottom=Side(style="thin", color=self.mau_nhan)), xuong_dong=False)
        self.cao(20)
        self.r += 1
        ds = []
        for khoa, nhan_vi, nhan_zh, md_vi, md_zh in DIEU_KHOAN:
            viet = (dk.get(khoa) or md_vi.format(n=hieu_luc_ngay)).rstrip(".")
            trung = (dk_zh.get(khoa) or md_zh.format(n=hieu_luc_ngay)).rstrip("。")
            ds.append((nhan_vi + viet + ".", nhan_zh + trung + "。"))
        for them in dk.get("them", []):
            ds.append((them.rstrip(".") + ".", ""))
        ct = self.ct
        chu_tk = ct.get("chu_tai_khoan") or ct.get("ten", "")
        tk = (f"Thông tin chuyển khoản: Số TK {ct.get('so_tai_khoan', '')} – {ct.get('ngan_hang', '')} – "
              f"Chủ tài khoản: {chu_tk}.")
        ds.append((tk, f"银行账户：{ct.get('so_tai_khoan', '')}，户名：{chu_tk}。"))
        for i, (viet, trung) in enumerate(ds, 1):
            chu = f"{i}. {viet}"
            if self.song_ngu and trung:
                chu += f"\n    {trung}"
            self.o(f"A{self.r}:G{self.r}", chu, self.font(9.5, mau="262626"), doc="top")
            self.cao(12.5 * _so_dong_chu(chu, _rong("A", "G"), 9.5) + 5)
            self.r += 1
        self.cach(6)
        self.dong_loi_ket = self.r
        chu = self.nhan("loi_ket", "\n")
        self.o(f"A{self.r}:G{self.r}", chu, self.font(10, nghieng=True, mau="262626"))
        self.cao(30 if self.song_ngu else 18)
        self.r += 1
        self.cach(12)

    def cho_ky(self, khach_hang_ten=""):
        ct = self.ct
        chuc_vu = (ct.get("chuc_vu") or "Giám đốc")
        chuc_vu_zh = ct.get("chuc_vu_zh") or CHUC_VU_ZH.get(chuc_vu.lower(), "")
        r = self.r
        dam = self.font(10.5, dam=True, mau=self.chinh)
        trai = self.nhan("xac_nhan", "\n")
        phai = chuc_vu.upper() + (f"\n{chuc_vu_zh}" if self.song_ngu and chuc_vu_zh else "")
        self.o(f"A{r}:C{r}", trai, dam, can="center")
        self.o(f"E{r}:G{r}", phai, dam, can="center")
        self.cao(30 if self.song_ngu else 18, r)
        r += 1
        nghieng = self.font(8.5, nghieng=True, mau="595959")
        self.o(f"A{r}:C{r}", self.nhan("ky_kh", " "), nghieng, can="center")
        self.o(f"E{r}:G{r}", self.nhan("ky_ct", " "), nghieng, can="center")
        self.cao(15, r)
        for k in range(1, 5):
            self.cao(16, r + k)
        r += 5
        self.o(f"E{r}:G{r}", ct.get("nguoi_ky", ""), self.font(11, dam=True), can="center")
        self.cao(20, r)
        self.khoi.append((getattr(self, "dong_loi_ket", self.r), r))  # lời kết + chữ ký cùng một trang
        self.r = r + 1

    def _mo_phong_trang(self, ty_le, ngat=True):
        """Ước lượng cách chia trang (A4 dọc). Trả về (số trang, tỉ lệ lấp đầy trang cuối, các dòng cần ngắt)."""
        ws = self.ws

        def cao(r):
            return ws.row_dimensions[r].height or 15

        suc_chua = (11.69 - 0.45 - 0.6) * 72 / ty_le  # đã đối chiếu với bản in PDF thực tế
        hd = self.dong_tieu_de_bang
        bat_dau = dict(self.khoi)
        y, trang, ngat_dong = 0.0, 1, []
        for r in range(1, self.r):
            h = cao(r)
            dau_trang = cao(hd) if (trang > 1 and r > hd) else 0
            if ngat and r in bat_dau:
                h_khoi = sum(cao(x) for x in range(r, bat_dau[r] + 1))
                if y + h_khoi > suc_chua and h_khoi < 0.8 * suc_chua and y > dau_trang + 1:
                    ngat_dong.append(r - 1)
                    trang += 1
                    y = cao(hd) if r > hd else 0
            if y + h > suc_chua:
                trang += 1
                y = cao(hd) if r > hd else 0
            y += h
        return trang, y / suc_chua, ngat_dong

    def hoan_thien(self, so_trang=None):
        ws, ct = self.ws, self.ct
        ws.print_area = f"A1:G{self.r - 1}"
        ws.print_title_rows = f"{self.dong_tieu_de_bang}:{self.dong_tieu_de_bang}"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = "portrait"
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        rong_pt = sum(w * 7 + 5 for w in DO_RONG_COT.values()) * 0.75
        ty_le = min(1.0, (8.27 - 0.8) * 72 / rong_pt)
        if so_trang:
            ws.page_setup.fitToHeight = int(so_trang)
        else:
            trang, lap_day, ngat_dong = self._mo_phong_trang(ty_le)
            if trang > 1 and lap_day < 0.3:
                # Trang cuối chỉ còn ít nội dung: thu nhỏ nhẹ cho vừa ít trang hơn
                ws.page_setup.fitToHeight = trang - 1
            else:
                ws.page_setup.fitToHeight = 0
                for d in ngat_dong:
                    ws.row_breaks.append(Break(id=d))
        ws.print_options.horizontalCentered = True
        ws.page_margins.left = ws.page_margins.right = 0.4
        ws.page_margins.top = 0.45
        ws.page_margins.bottom = 0.6
        ws.page_margins.footer = 0.25
        chan = f"{ct.get('ten', '')}  •  {ct.get('dien_thoai', '')}  •  Trang &P/&N"
        ws.oddFooter.center.text = chan
        ws.oddFooter.center.size = 8
        ws.oddFooter.center.font = f"{PHONG},Regular"
        ws.oddFooter.center.color = "7F7F7F"


# ---------------------------------------------------------------------------
def doc_ngay(gia_tri):
    if not gia_tri:
        return dt.date.today()
    for dinh_dang in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return dt.datetime.strptime(str(gia_tri), dinh_dang).date()
        except ValueError:
            pass
    raise SystemExit(f"Không hiểu ngày: {gia_tri!r} (dùng dạng 2026-10-10 hoặc 10/10/2026)")


def _ty_le_vat(vat):
    vat = 0.1 if vat is None else float(vat)
    return vat / 100 if vat > 1 else vat


def ve_bao_gia(du_lieu, ct, mau_trong=False):
    """Tạo workbook báo giá. mau_trong=True: mẫu trống để người dùng tự điền trong Excel."""
    song_ngu = bool(du_lieu.get("song_ngu"))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Báo giá"
    b = BangBaoGia(ws, ct, song_ngu)
    b.dau_trang()
    b.tieu_de()

    kh = du_lieu.get("khach_hang") or {}
    ngay = doc_ngay(du_lieu.get("ngay"))
    so_bg = du_lieu.get("so_bao_gia") or f"BG-{ngay:%Y}/{ngay:%m%d}"
    hieu_luc_ngay = du_lieu.get("hieu_luc_ngay", 30)
    cham = "................................"
    trai = [(b.nhan("kinh_gui"), kh.get("ten") or cham, True),
            (b.nhan("nguoi_lien_he"), kh.get("nguoi_lien_he") or (cham if mau_trong else ""), False),
            (b.nhan("dien_thoai"), kh.get("dien_thoai") or (cham if mau_trong else ""), False),
            (b.nhan("email"), kh.get("email") or (cham if mau_trong else ""), False),
            (b.nhan("dia_chi"), kh.get("dia_chi") or (cham if mau_trong else ""), False)]
    if not mau_trong:
        # Giữ đủ nhãn để khung có cấu trúc (khách có thể ghi tay); bỏ Email nếu trống cho gọn
        trai = [x for x in trai if x[1] or x is not trai[3]]
    hl = f"{hieu_luc_ngay} ngày" + (f" / {hieu_luc_ngay}天" if song_ngu else "")
    phai = [(b.nhan("so"), so_bg if not mau_trong else "BG-....../......", True),
            (b.nhan("ngay"), ngay.strftime("%d/%m/%Y") if not mau_trong else "...../...../........", False),
            (b.nhan("hieu_luc"), hl, False)]
    if ct.get("nguoi_ky"):
        phai.append((b.nhan("phu_trach"), f"{ct['nguoi_ky']} – {ct.get('dien_thoai', '')}".rstrip(" –"), False))
    b.hai_khung(trai, phai)
    if du_lieu.get("cong_trinh"):
        b.dong_cong_trinh(du_lieu["cong_trinh"], du_lieu.get("cong_trinh_zh"))
    b.loi_mo_dau(du_lieu.get("cong_trinh"), du_lieu.get("cong_trinh_zh"))
    b.tieu_de_bang()

    hang = du_lieu.get("hang_hoa") or []
    dong_dau = b.r
    tong_hang = 0
    if mau_trong:
        for i in range(SO_DONG_TRONG_MAU):
            b.dong_hang(None, None, i % 2 == 1)
    else:
        if not hang:
            raise SystemExit("Cần ít nhất 1 mặt hàng trong 'hang_hoa'.")
        nhom_hien_tai, so_nhom = object(), 0
        co_nhom = any(mh.get("nhom") for mh in hang)
        for i, mh in enumerate(hang, 1):
            if co_nhom and mh.get("nhom") != nhom_hien_tai:
                nhom_hien_tai = mh.get("nhom")
                so_nhom += 1
                b.dong_nhom(so_nhom, nhom_hien_tai or "Khác", mh.get("nhom_zh"))
            b.dong_hang(i, mh, i % 2 == 0)
            if mh.get("so_luong") is not None and mh.get("don_gia") is not None:
                tong_hang += round(mh["so_luong"] * mh["don_gia"])
    dong_cuoi = b.r - 1

    vat = _ty_le_vat(du_lieu.get("vat", 0.1))
    da_gom_vat = bool(du_lieu.get("gia_da_gom_vat"))
    tien_thue = 0 if da_gom_vat else round(tong_hang * vat)
    tong_cong = tong_hang + tien_thue
    b.tong_tien(dong_dau, dong_cuoi, vat, None if mau_trong else tong_cong, da_gom_vat)
    dk = dict(du_lieu.get("dieu_khoan") or {})
    dk_zh = dict(du_lieu.get("dieu_khoan_zh") or {})
    if da_gom_vat:
        dk.setdefault("thue", f"Đơn giá trên đã bao gồm thuế GTGT {vat:.0%}")
        dk_zh.setdefault("thue", f"以上单价已含{vat:.0%}增值税")
    b.dieu_khoan(dk, dk_zh, hieu_luc_ngay)
    b.cho_ky()
    b.hoan_thien(du_lieu.get("so_trang"))

    if mau_trong:
        hd = wb.create_sheet("Hướng dẫn")
        hd.column_dimensions["A"].width = 100
        for i, dong in enumerate([
            "HƯỚNG DẪN DÙNG MẪU BÁO GIÁ",
            "",
            "1. Trang 'Báo giá': điền thông tin khách hàng, số báo giá, ngày báo giá vào chỗ dấu chấm.",
            "2. Bảng hàng: nhập Tên hàng (cột B), ĐVT (C), Số lượng (D), Đơn giá (E), Ghi chú (G).",
            "   Cột Thành tiền, Cộng tiền hàng, Thuế GTGT và Tổng cộng TỰ ĐỘNG tính.",
            "3. Thuế suất GTGT ở dòng 'Thuế GTGT (VAT)' (mặc định 10%), đổi thành 8% hoặc 0% nếu cần.",
            "4. Gõ lại dòng 'Bằng chữ' theo số tổng cộng.",
            "5. Cần thêm dòng hàng: chọn một dòng hàng → chuột phải → Insert (Chèn), rồi copy dòng phía trên xuống.",
            "6. Gửi khách: File → Save As → PDF. Nên Save As thành file mới mỗi lần báo giá để giữ mẫu gốc.",
        ], 1):
            hd[f"A{i}"] = dong
            hd[f"A{i}"].font = Font(name=PHONG, size=13 if i == 1 else 10.5, bold=(i == 1))
        wb.active = 0

    tom_tat = {
        "cong_ty": ct.get("ten"),
        "logo_tam": bool(ct.get("logo_tam")),
        "thong_tin_cong_ty_con_trong": thong_tin_con_trong(ct),
    }
    if not mau_trong:
        tom_tat.update({
            "so_bao_gia": so_bg,
            "so_mat_hang": len(hang),
            "cong_tien_hang": tong_hang,
            "thue_suat": vat,
            "gia_da_gom_vat": da_gom_vat,
            "tien_thue": tien_thue,
            "tong_cong": tong_cong,
            "bang_chu": so_thanh_chu(tong_cong),
            "bang_chu_zh": so_thanh_chu_trung(tong_cong) if song_ngu else None,
        })
    return wb, tom_tat


def luu(wb, duong_dan_ra):
    duong_dan_ra = Path(duong_dan_ra)
    duong_dan_ra.parent.mkdir(parents=True, exist_ok=True)
    wb.calculation.fullCalcOnLoad = True
    wb.save(duong_dan_ra)
    return str(duong_dan_ra)


def tao(du_lieu, duong_dan_ra):
    ct = nap_cong_ty(du_lieu.get("cong_ty"))
    wb, tom_tat = ve_bao_gia(du_lieu, ct)
    return {"file_excel": luu(wb, duong_dan_ra), **tom_tat}


def tao_mau(ct, duong_dan_ra, song_ngu=False):
    wb, tom_tat = ve_bao_gia({"song_ngu": song_ngu}, ct, mau_trong=True)
    return {"file_excel": luu(wb, duong_dan_ra), **tom_tat}


def xuat_pdf(file_excel):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    file_excel = Path(file_excel)
    pdf = file_excel.with_suffix(".pdf")
    with tempfile.TemporaryDirectory() as tam:
        # Bản tạm chỉ hiện trang "Báo giá" để PDF không in kèm trang hướng dẫn
        ban_tam = Path(tam) / (file_excel.stem + ".xlsx")
        wb = openpyxl.load_workbook(file_excel, rich_text=True)
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
    ap.add_argument("--song-ngu", action="store_true", help="dùng với --mau: mẫu song ngữ Việt – Trung")
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
        ket_qua = tao_mau(tim_cong_ty(args.mau), args.out, args.song_ngu)
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
