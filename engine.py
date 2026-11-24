# engine.py - BANGOR PACKING LIST ENGINE (FINAL - Detail DO/Gabungan + NO KOLI kosong)
from collections import OrderedDict
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side
import datetime

ENGINE_VERSION = "ENGINE-FINAL-DETAILDO-02"

JUNK = ("PT BANGOR", "DELIVERY ORDER DETAIL", "RINCIAN PEMINDAHAN",
        "RINCIAN PENGIRIMAN PESANAN", "ACCURATE ACCOUNTING", "PRINTED ON",
        "PAGE BREAK", "FROM ", "DARI ", "BRANCH :", "CABANG", "GUDANG ASAL",
        "FILTERED BY", "ITEM NAME", "NAMA BARANG", "TOTAL", "OF 34",
        "DETAIL DO", "KODE #", "CODE#")


class PackingListEngine:
    def __init__(self, input_file, output_file, master_data=None,
                 progress_callback=None, master_wb=None):
        self.input_file = input_file
        self.output_file = output_file
        self.progress = progress_callback or (lambda m: None)
        self.master_wb = master_wb
        self.wb = None
        self.MasterData = dict(master_data) if master_data else {}
        self.DOGroups = OrderedDict()
        self.TransformResult = OrderedDict()
        self.FinalResult = {}
        self.format_type = None
        self.target_sheet = ""
        self.logs = []
        self.stats = {"do_count": 0, "it_count": 0, "item_count": 0}

    def log(self, m):
        self.logs.append(m)
        self.progress(m)

    def Txt(self, v):
        return "" if v is None else str(v).strip().upper()

    def Raw(self, v):
        return "" if v is None else str(v).strip()

    def Norm(self, v):
        t = self.Txt(v)
        for ch in [" ", ".", "_", "-", "/", "\\", "#", ":", "?"]:
            t = t.replace(ch, "")
        return t

    def Num(self, v):
        try:
            if v is None or v == "":
                return 0.0
            return float(str(v).replace(",", "").strip())
        except Exception:
            return 0.0

    def SafeSheetName(self, n):
        if not n:
            return "Sheet"
        for c in ['\\', '/', '*', '?', ':', '[', ']']:
            n = n.replace(c, " ")
        return n[:31].strip() or "Sheet"

    def NormUnit(self, u):
        u = self.Txt(u)
        if u in ("DOZ", "DOZEN"):
            return "DUS"
        if u == "PCK":
            return "PACK"
        return u

    def is_junk(self, nama):
        t = self.Txt(nama)
        return any(t.startswith(j) or t == j for j in JUNK)

    def label_value(self, ws, row, label, maxc):
        for c in range(1, maxc + 1):
            if self.Norm(ws.cell(row, c).value) == label:
                for c2 in range(c + 1, maxc + 1):
                    v = ws.cell(row, c2).value
                    if v is None:
                        continue
                    t = self.Raw(v)
                    if not t or t == ":":
                        continue
                    if self.Norm(t) == label:
                        continue
                    return v
        return None

    def find_label_row(self, ws, start, end, labels, maxc):
        for r in range(start, end):
            for c in range(1, maxc + 1):
                if self.Norm(ws.cell(r, c).value) in labels:
                    return r
        return None

    def load_master_data(self):
        if self.MasterData:
            self.log(f"📚 Master Data dari database: {len(self.MasterData)} item")
            return
        sh = next((s for s in self.wb.sheetnames if "Master" in s), None)
        if not sh:
            self.log("⚠️ Master Data kosong (DB & sheet tidak ada).")
            return
        ws = self.wb[sh]
        for r in range(6, (ws.max_row or 300) + 1):
            nama = self.Txt(ws.cell(r, 2).value)
            if not nama or nama in self.MasterData:
                continue
            self.MasterData[nama] = {
                "Nama": nama,
                "QtyMax": self.Num(ws.cell(r, 3).value),
                "Unit": self.Txt(ws.cell(r, 5).value),
                "Kategori": self.Txt(ws.cell(r, 6).value),
            }
        self.log(f"📚 Master Data dari sheet: {len(self.MasterData)} item")

    def detect_format(self):
        sheets = self.wb.sheetnames
        it = next((s for s in sheets if "Rincian" in s and "Pemindahan" in s), None)
        dod = next((s for s in sheets if "Delivery Order" in s), None)
        gab = next((s for s in sheets if "Gabungan" in s or "Detail" in s), None)
        peng = next((s for s in sheets if "Pengiriman" in s), None)
        if it:
            self.format_type, self.target_sheet = "Rincian", it
        elif dod:
            self.format_type, self.target_sheet = "Blok", dod
        elif gab:
            self.format_type, self.target_sheet = "Gabungan", gab
        elif peng:
            self.format_type, self.target_sheet = "Blok", peng
        else:
            raise Exception(f"Format tidak dikenali. Sheet: {sheets}")
        self.log(f"📄 Format: {self.format_type} | Sheet: '{self.target_sheet}'")

    # ---------- IT: Rincian Pemindahan Barang ----------
    def parse_rincian(self):
        ws = self.wb[self.target_sheet]
        maxr, maxc = ws.max_row or 1, ws.max_column or 12
        keys = {
            "DO": ["NOPEMINDAHAN", "NOMORPEMINDAHAN", "NOTRANSAKSI"],
            "TGL": ["TANGGAL", "DATE", "TGL"],
            "TUJUAN": ["GUDANGTUJUANDARI", "GUDANGTUJUAN"],
            "GUDANG": ["GUDANG"],
            "NAMA": ["NAMABARANG", "NAMAITEM"],
            "UNIT": ["SATUAN", "UNIT"],
            "QTY": ["KUANTITAS", "QUANTITY", "JUMLAH"],
        }
        hrow, cmap = None, {}
        for r in range(1, min(maxr, 30) + 1):
            cand = {}
            for c in range(1, maxc + 1):
                n = self.Norm(ws.cell(r, c).value)
                if not n:
                    continue
                for f, kws in keys.items():
                    if n in kws:
                        cand[f] = c
            if "NAMA" in cand and "QTY" in cand:
                hrow, cmap = r, cand
                break
        if not hrow:
            raise Exception("Header sheet IT tidak ditemukan.")
        out_col = cmap.get("TUJUAN") or cmap.get("GUDANG")
        last_do = last_tgl = last_out = None
        for r in range(hrow + 1, maxr + 1):
            def g(f):
                return ws.cell(r, cmap[f]).value if f in cmap else None
            do = self.Raw(g("DO")) or last_do
            tgl = g("TGL") if g("TGL") is not None else last_tgl
            out = (self.Raw(ws.cell(r, out_col).value) if out_col else "") or last_out
            nama = self.Raw(g("NAMA"))
            qty = self.Num(g("QTY"))
            if not do or not nama or qty <= 0 or self.is_junk(nama):
                continue
            last_do, last_tgl, last_out = do, tgl, out
            if do not in self.DOGroups:
                self.DOGroups[do] = {"Outlet": out, "Tanggal": tgl, "ShipVia": "", "Items": []}
            self.DOGroups[do]["Items"].append({
                "Nama": nama, "Qty": qty,
                "Unit": self.NormUnit(self.Raw(g("UNIT"))), "Notes": ""})
        self.stats["it_count"] = len(self.DOGroups)
        self.log(f"✅ IT terbaca: {len(self.DOGroups)} dokumen")

    # ---------- Gabungan / Detail DO / Delivery Order Detail ----------
    def parse_blok(self):
        ws = self.wb[self.target_sheet]
        maxr, maxc = ws.max_row or 1, ws.max_column or 20
        is_gab = (self.format_type == "Gabungan")

        starts = []
        for r in range(1, maxr + 1):
            for c in range(1, maxc + 1):
                if self.Norm(ws.cell(r, c).value) in ("NUMBER", "NOMOR"):
                    starts.append(r)
                    break
        if not starts:
            raise Exception(f"Label 'Number/Nomor #' tidak ditemukan di '{self.target_sheet}'")
        self.log(f"🔎 Ditemukan {len(starts)} blok DO")

        for i, r in enumerate(starts):
            end = starts[i + 1] if i + 1 < len(starts) else maxr + 1

            nomor = self.Raw(self.label_value(ws, r, "NUMBER", maxc)) or \
                    self.Raw(self.label_value(ws, r, "NOMOR", maxc))
            if not nomor:
                continue
            outlet = self.Raw(self.label_value(ws, r, "CUSTOMER", maxc)) or \
                     self.Raw(self.label_value(ws, r, "PELANGGAN", maxc))

            tanggal = None
            for tr in range(r, min(r + 5, end)):
                tanggal = self.label_value(ws, tr, "DATE", maxc)
                if tanggal is None:
                    tanggal = self.label_value(ws, tr, "TANGGAL", maxc)
                if tanggal is not None:
                    break

            if nomor not in self.DOGroups:
                self.DOGroups[nomor] = {"Outlet": outlet, "Tanggal": tanggal,
                                        "ShipVia": "", "Items": []}
            else:
                if outlet and not self.DOGroups[nomor]["Outlet"]:
                    self.DOGroups[nomor]["Outlet"] = outlet
                if tanggal and not self.DOGroups[nomor]["Tanggal"]:
                    self.DOGroups[nomor]["Tanggal"] = tanggal

            hr = self.find_label_row(ws, r, min(r + 10, end), ("CODE", "KODE"), maxc)
            if hr is None:
                continue
            cc = nc = qc = uc = None
            for c in range(1, maxc + 1):
                n = self.Norm(ws.cell(hr, c).value)
                if n in ("CODE", "KODE") and cc is None:
                    cc = c
                elif n in ("ITEMNAME", "NAMABARANG", "NAMAITEM") and nc is None:
                    nc = c
                elif n in ("QUANTITY", "KUANTITAS", "QTY", "JUMLAH") and qc is None:
                    qc = c
                elif n in ("UNIT", "SATUAN") and uc is None:
                    uc = c
            if cc is None or nc is None or qc is None:
                continue

            cnt = 0
            bbt = 0
            for ir in range(hr + 1, end):
                code = self.Raw(ws.cell(ir, cc).value)
                nama = self.Raw(ws.cell(ir, nc).value)
                qty = self.Num(ws.cell(ir, qc).value)
                unit = self.Raw(ws.cell(ir, uc).value) if uc else ""
                if not nama or qty <= 0 or self.is_junk(nama):
                    continue
                if self.Norm(code) in ("CODE", "KODE"):
                    continue

                notes = ""
                if is_gab:
                    ck = code[:-2] if code.endswith(".0") else code
                    if ck.isdigit():
                        notes = "BBT"
                        bbt += 1

                self.DOGroups[nomor]["Items"].append({
                    "Nama": nama, "Qty": qty,
                    "Unit": self.NormUnit(unit), "Notes": notes})
                cnt += 1

            self.log(f"   📦 {nomor} | {cnt} item ({bbt} BBT) | {outlet} | Tgl: {tanggal or 'KOSONG'}")

        self.stats["do_count"] = len(self.DOGroups)
        self.log(f"✅ Total DO: {len(self.DOGroups)}")

    # ---------- transform ----------
    def rank_kategori(self, kat, receh):
        k = self.Txt(kat)
        if k in ("HAMPERS", "GRILL BOX", "INNER", "CRAZY BUCKET", "MINYAK", "PACKAGING"):
            k = "LAIN-LAIN"
        table = {"DAGING": ("Daging", 1), "FROZEN": ("Frozen", 3), "KEJU": ("Keju", 5),
                 "KENTANG": ("Kentang", 7), "SAUS": ("Saus", 9), "BUTTER": ("Butter", 11),
                 "DRY": ("Dry", 12), "LAIN-LAIN": ("Lain-lain", 14),
                 "BUN": ("Bun", 16), "ROTI": ("Bun", 16)}
        if k in table:
            nm, rk = table[k]
            if k in ("BUTTER", "BUN", "ROTI"):
                return nm, rk
            return (nm + " Receh", rk + 1) if receh else (nm, rk)
        return (k + " Receh" if receh else k), 99

    def transform(self):
        missing = set()
        for do_no, d in self.DOGroups.items():
            hasil = []
            for it in d["Items"]:
                nama = self.Raw(it["Nama"])
                key = self.Txt(nama)
                notes = it.get("Notes", "")
                m = self.MasterData.get(key)
                if not m:
                    missing.add(nama)
                    continue
                if "DUS" in key:
                    continue
                qmax = self.Num(m["QtyMax"])
                um = self.Txt(m["Unit"])
                kat = m["Kategori"]
                ud = self.Txt(it["Unit"])
                q = self.Num(it["Qty"])
                if qmax <= 0 or q <= 0:
                    continue
                if um == "DUS":
                    real, kmax, uo = q, qmax, "DUS"
                elif um == "PACK":
                    uo = "PACK"
                    if "DUS" in ud:
                        if "BEEF PATTY SMALL" in key or "BEEF PATTY LARGE" in key:
                            isi = 18
                        elif "THOUSAND ISLAND" in key:
                            isi = 20
                        else:
                            isi = qmax
                        real = q * isi
                    else:
                        real = q
                    kmax = 20 if "THOUSAND ISLAND" in key else qmax
                elif um == "BKS":
                    uo = "BKS"
                    real = q * qmax if "DUS" in ud else q
                    kmax = qmax
                else:
                    uo, real, kmax = um, q, qmax
                if kmax <= 0:
                    kmax = qmax
                while real > 0.0001:
                    koli = min(real, kmax)
                    receh = koli < (kmax - 0.0001)
                    okat, rk = self.rank_kategori(kat, receh)
                    hasil.append({"NamaBarang": nama, "Qty": koli, "Unit": uo,
                                  "Kategori": okat, "Rank": rk, "Notes": notes})
                    real -= koli
                    self.stats["item_count"] += 1
            hasil.sort(key=lambda x: (x["Rank"], x["NamaBarang"].upper()))
            self.TransformResult[do_no] = {"Outlet": d["Outlet"], "Tanggal": d["Tanggal"],
                                           "ShipVia": d["ShipVia"], "Items": hasil}
        if missing:
            self.log(f"⚠️ {len(missing)} item tidak ada di Master: "
                     + ", ".join(sorted(missing)[:6]))
        self.log(f"✅ Transform: {len(self.TransformResult)} DO / {self.stats['item_count']} baris")

    def group_receh(self):
        for k, v in self.TransformResult.items():
            self.FinalResult[k] = v

    # ---------- export ----------
    def export(self):
        wb = self.master_wb if self.master_wb is not None else Workbook()
        if self.master_wb is None:
            wb.remove(wb.active)
        if not self.FinalResult:
            if self.master_wb is None:
                ws = wb.create_sheet("Info")
                ws["A1"] = "Tidak ada data."
                wb.save(self.output_file)
            return wb
        thin, med = Side(style="thin"), Side(style="medium")
        bt = Border(left=thin, right=thin, top=thin, bottom=thin)
        bm = Border(left=med, right=med, top=med, bottom=med)
        for do_no, d in self.FinalResult.items():
            base = self.SafeSheetName(d["Outlet"])
            nm, n = base, 2
            while nm in wb.sheetnames:
                nm = f"{base[:28]}_{n}"
                n += 1
            ws = wb.create_sheet(nm)
            for col, w in zip("ABCDE", (10, 38, 12, 14, 25)):
                ws.column_dimensions[col].width = w
            ws.merge_cells("A2:E2")
            ws["A2"] = "PACKING LIST"
            ws["A2"].font = Font(bold=True, size=18)
            ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[2].height = 30
            ws["A4"] = "DELIVERY\nORDER :"
            ws["B4"] = do_no
            ws["D4"] = "Ship Via:"
            ws["E4"] = d.get("ShipVia", "")
            ws["A5"] = "OUTLET :"
            ws["B5"] = d["Outlet"]
            ws["A6"] = "DELIVERY\nDATE :"
            tgl = d.get("Tanggal", "")
            ws["B6"] = tgl
            if isinstance(tgl, (datetime.datetime, datetime.date)):
                ws["B6"].number_format = "dd-mm-yyyy"
            for r in range(4, 7):
                ws.cell(r, 1).font = Font(bold=True, size=12)
                ws.cell(r, 1).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                ws.cell(r, 2).font = Font(size=12)
                ws.cell(r, 2).alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[4].height = 28
            ws.row_dimensions[5].height = 20
            ws.row_dimensions[6].height = 28
            for c, h in enumerate(["NO KOLI", "NAMA BARANG", "QTY", "UNIT", "NOTES"], 1):
                cell = ws.cell(9, c, h)
                cell.font = Font(bold=True, size=12)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = bm
            ws.row_dimensions[9].height = 25
            r = 10
            for it in d["Items"]:
                # kolom 1 (NO KOLI) sengaja dibiarkan KOSONG
                ws.cell(r, 2, it["NamaBarang"])
                ws.cell(r, 3, it["Qty"])
                ws.cell(r, 4, it["Unit"])
                if it.get("Notes"):
                    ws.cell(r, 5, it["Notes"])
                for c in range(1, 6):
                    cell = ws.cell(r, c)
                    cell.border = bt
                    cell.font = Font(size=12)
                for c in (1, 3, 4, 5):
                    ws.cell(r, c).alignment = Alignment(horizontal="center")
                r += 1
            ws.page_setup.orientation = "portrait"
            ws.page_setup.fitToWidth = 1
            ws.sheet_properties.pageSetUpPr.fitToPage = True
        if self.master_wb is None:
            wb.save(self.output_file)
        self.log("✅ Export selesai")
        return wb

    def run(self):
        self.logs = []
        self.stats = {"do_count": 0, "it_count": 0, "item_count": 0}
        self.log(f"=== {ENGINE_VERSION} ===")
        self.wb = load_workbook(self.input_file, data_only=True)
        self.load_master_data()
        self.detect_format()
        if self.format_type == "Rincian":
            self.parse_rincian()
        else:
            self.parse_blok()
        self.transform()
        self.group_receh()
