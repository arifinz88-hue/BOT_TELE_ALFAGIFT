import re

PRODUCT_RE = re.compile(r"Produk=\s*(.*?)\s*Qty=\s*(\d+)", re.I)
OID_RE = re.compile(r"O-(\d{6})-([A-Z0-9]+)", re.I)

def parse_line(line):

    m = OID_RE.search(line)

    if not m:
        return None

    tanggal = m.group(1)
    oid = f"O-{tanggal}-{m.group(2)}"

    nama = line.split(":",1)[0].strip() if ":" in line else "-"

    toko = "UNKNOWN"

    if "|:" in line:

        meta = line.split("|:",1)[1].split(":")

        if len(meta) >= 3:

            toko = f"{meta[1].strip()} - {meta[2].strip()}"

    products = []

    for p in PRODUCT_RE.finditer(line):

        produk = p.group(1).strip()
        qty = int(p.group(2))

        products.append((oid,tanggal,toko,nama,produk,qty))

    return products