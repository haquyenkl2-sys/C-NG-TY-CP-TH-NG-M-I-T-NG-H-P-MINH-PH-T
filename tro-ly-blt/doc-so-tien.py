#!/usr/bin/env python3
"""Đọc số tiền VNĐ thành chữ tiếng Việt.

Cách dùng:
    python3 doc-so-tien.py 11000000        -> Mười một triệu đồng chẵn.
    python3 doc-so-tien.py 1.250.500       -> Một triệu hai trăm năm mươi nghìn năm trăm đồng.
"""
import sys

CHU_SO = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]


def doc_ba_so(n, day_du):
    """Đọc nhóm 3 chữ số. day_du=True khi phía trước còn nhóm khác (phải đọc 'không trăm')."""
    tram, chuc, dv = n // 100, (n // 10) % 10, n % 10
    phan = []
    if tram or day_du:
        phan.append(CHU_SO[tram] + " trăm")
    if chuc == 0:
        if dv and (tram or day_du):
            phan.append("linh")
    elif chuc == 1:
        phan.append("mười")
    else:
        phan.append(CHU_SO[chuc] + " mươi")
    if dv:
        if dv == 1 and chuc >= 2:
            phan.append("mốt")
        elif dv == 4 and chuc >= 2:
            phan.append("tư")
        elif dv == 5 and chuc >= 1:
            phan.append("lăm")
        else:
            phan.append(CHU_SO[dv])
    return " ".join(phan)


def doc_so(n):
    if n == 0:
        return "không"
    if n >= 10**9:
        # Phần tỷ có thể lớn hơn 999 -> đọc đệ quy: "một nghìn tỷ", ...
        ty, con_lai = divmod(n, 10**9)
        chu = doc_so(ty) + " tỷ"
        if con_lai:
            chu += " " + doc_so_duoi_ty(con_lai, day_du=True)
        return chu
    return doc_so_duoi_ty(n, day_du=False)


def doc_so_duoi_ty(n, day_du):
    nhom = [(n // 10**6) % 1000, (n // 10**3) % 1000, n % 1000]
    don_vi = [" triệu", " nghìn", ""]
    phan = []
    for i, g in enumerate(nhom):
        if g == 0:
            continue
        phan.append(doc_ba_so(g, day_du or bool(phan)) + don_vi[i])
    return " ".join(phan)


def doc_tien(so):
    chu = doc_so(so)
    chu = chu[0].upper() + chu[1:] + " đồng"
    if so and so % 1000 == 0:
        chu += " chẵn"
    return chu + "."


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    so = sys.argv[1].replace(".", "").replace(",", "").replace(" ", "")
    if not so.isdigit():
        sys.exit("Số tiền không hợp lệ: " + sys.argv[1])
    print(doc_tien(int(so)))
