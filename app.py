import streamlit as st
import os
import tempfile
import sqlite3
import pandas as pd
from collections import OrderedDict
from openpyxl import load_workbook
from engine import PackingListEngine

st.set_page_config(page_title="Bangor - Packing List Generator",
                   page_icon="📦", layout="wide", initial_sidebar_state="collapsed")

DB = "bangor_master.db"


def init_db():
    con = sqlite3.connect(DB)
    con.execute("""CREATE TABLE IF NOT EXISTS master_barang(
        nama TEXT PRIMARY KEY, qty_max REAL, unit TEXT, kategori TEXT)""")
    con.commit()
    con.close()


def load_db():
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    d = {}
    for r in con.execute("SELECT * FROM master_barang"):
        d[r["nama"].upper()] = {"Nama": r["nama"], "QtyMax": r["qty_max"],
                                "Unit": r["unit"], "Kategori": r["kategori"]}
    con.close()
    return d


def save_db(d):
    con = sqlite3.connect(DB)
    con.execute("DELETE FROM master_barang")
    for k, v in d.items():
        con.execute("INSERT INTO master_barang VALUES (?,?,?,?)",
                    (k, v["QtyMax"], v["Unit"], v["Kategori"]))
    con.commit()
    con.close()


def parse_master(f):
    wb = load_workbook(f, read_only=True, data_only=True)
    sh = next((s for s in wb.sheetnames if "Master" in s), None)
    if not sh:
        wb.close()
        raise Exception(f"Sheet Master tidak ada: {wb.sheetnames}")
    ws = wb[sh]
    out = {}
    for row in ws.iter_rows(min_row=6, max_col=6, values_only=True):
        nm = row[1]
        if not nm:
            continue
        s = str(nm).strip()
        k = s.upper()
        if not s or k in out:
            continue
        try:
            qm = float(str(row[2] or 0).replace(",", ""))
        except Exception:
            qm = 0
        out[k] = {"Nama": s, "QtyMax": qm,
                  "Unit": str(row[4] or "").strip().upper(),
                  "Kategori": str(row[5] or "").strip().upper()}
    wb.close()
    return out


init_db()

st.markdown("<h1 style='text-align:center'>📦 Generator Packinglist By-Logistik</h1>",
            unsafe_allow_html=True)

tab1, tab2 = st.tabs(["🚀 Generate Packing List", "🗄️ Master Data"])

with tab2:
    st.subheader("Database Master Barang")
    db = load_db()
    st.info(f"Tersimpan: **{len(db)} item**")
    up = st.file_uploader("Import Excel Master (baris 6+: B=Nama, C=QtyMax, E=Unit, F=Kategori)",
                          type=["xlsx", "xlsm"], key="mup")
    if up:
        try:
            d = parse_master(up)
            save_db(d)
            st.success(f"✅ {len(d)} item disimpan.")
            st.rerun()
        except Exception as e:
            st.error(f"❌ {e}")
    st.divider()
    rows = [{"Nama Barang": v["Nama"], "Qty Max": v["QtyMax"],
             "Unit": v["Unit"], "Kategori": v["Kategori"]} for v in db.values()]
    df = pd.DataFrame(rows) if rows else pd.DataFrame(
        columns=["Nama Barang", "Qty Max", "Unit", "Kategori"])
    ed = st.data_editor(df, num_rows="dynamic", use_container_width=True, height=420, key="med")
    c1, c2 = st.columns(2)
    if c1.button("💾 Simpan", type="primary", use_container_width=True):
        nd = {}
        for _, r in ed.iterrows():
            nm = str(r["Nama Barang"]).strip()
            if nm and nm.lower() != "nan":
                nd[nm.upper()] = {
                    "Nama": nm,
                    "QtyMax": float(r["Qty Max"]) if pd.notna(r["Qty Max"]) else 0,
                    "Unit": str(r["Unit"]).strip().upper() if pd.notna(r["Unit"]) else "",
                    "Kategori": str(r["Kategori"]).strip().upper() if pd.notna(r["Kategori"]) else ""}
        save_db(nd)
        st.success(f"✅ {len(nd)} item disimpan.")
        st.rerun()
    if c2.button("🗑️ Kosongkan DB", use_container_width=True):
        save_db({})
        st.rerun()

with tab1:
    st.subheader("Upload DO / IT")
    ups = st.file_uploader("Pilih file Excel", type=["xlsx", "xlsm"],
                           accept_multiple_files=True, key="dup")
    if ups:
        uniq, seen = [], set()
        for f in ups:
            k = (f.name, f.size)
            if k not in seen:
                uniq.append(f)
                seen.add(k)
        ups = uniq
        st.success(f"✅ {len(ups)} file unik")
    dbm = load_db()
    ready = bool(ups) and bool(dbm)
    if not dbm:
        st.error("⚠️ Master Data kosong. Isi dulu di tab Master Data.")
    go = st.button("🚀 GENERATE PACKING LIST", type="primary",
                   use_container_width=True, disabled=not ready)
    if ready and go:
        master_wb = None
        tdo = tit = titem = 0
        seen_do = set()
        bar = st.progress(0.0)
        stat = st.empty()
        logb = st.code("", height=200)
        logs = []

        def cb(m):
            logs.append(m)
            logb.text("\n".join(logs[-14:]))

        tmps = []
        try:
            for i, f in enumerate(ups):
                bar.progress((i + 1) / len(ups))
                stat.text(f"Memproses {i+1}/{len(ups)}: {f.name}")
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as t:
                    t.write(f.getvalue())
                    p = t.name
                tmps.append(p)
                eng = PackingListEngine(p, p.replace(".xlsx", "_o.xlsx"),
                                        master_data=dbm, progress_callback=cb,
                                        master_wb=master_wb)
                eng.run()
                tdo += eng.stats["do_count"]
                tit += eng.stats["it_count"]
                titem += eng.stats["item_count"]
                uf = OrderedDict()
                for k, v in eng.FinalResult.items():
                    if k not in seen_do:
                        uf[k] = v
                        seen_do.add(k)
                eng.FinalResult = uf
                master_wb = eng.export()
            out = tempfile.mktemp(suffix=".xlsx")
            tmps.append(out)
            master_wb.save(out)
            with open(out, "rb") as fh:
                data = fh.read()
            bar.progress(1.0)
            stat.text("✅ Selesai")
            c1, c2, c3 = st.columns(3)
            c1.metric("DO", tdo)
            c2.metric("IT", tit)
            c3.metric("Baris Item", titem)
            st.download_button("⬇️ Download Packinglist (.xlsx)", data,
                               file_name=f"Packinglist_{len(ups)}FILE.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True, type="primary")
        except Exception as e:
            st.error(f"**GAGAL:** {e}")
        finally:
            for p in tmps:
                if os.path.exists(p):
                    try:
                        os.remove(p)
                    except Exception:
                        pass