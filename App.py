import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
from io import BytesIO

st.set_page_config(page_title="Tilbakemelding Verktøy", page_icon="⚙️", layout="wide")

# ── Fiori-inspirert styling ──────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif !important;
}

/* Hide default Streamlit header */
header[data-testid="stHeader"] { display: none; }
#MainMenu { display: none; }
footer { display: none; }

/* Shell bar */
.shellbar {
    background: #354a5e;
    padding: 0 32px;
    height: 48px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin: -1rem -1rem 0 -1rem;
    box-shadow: 0 2px 6px rgba(0,0,0,0.2);
}
.shellbar-title {
    font-size: 15px;
    font-weight: 600;
    color: #fff;
    letter-spacing: 0.01em;
}
.shellbar-by {
    font-size: 12px;
    color: rgba(255,255,255,0.4);
    letter-spacing: 0.03em;
}

/* Info / warning strips */
.info-strip {
    background: #e8f2ff;
    border-left: 3px solid #0070f2;
    border-radius: 4px;
    padding: 9px 14px;
    font-size: 13px;
    color: #0040b0;
    margin: 16px 0 10px 0;
}
.warning-strip {
    background: #fff8e6;
    border-left: 3px solid #e9a800;
    border-radius: 4px;
    padding: 9px 14px;
    font-size: 13px;
    color: #7a5000;
    margin-bottom: 20px;
}

/* Cards */
.stCard {
    background: #fff;
    border: 1px solid #e5e5e5;
    border-radius: 4px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.07);
    padding: 20px;
    margin-bottom: 16px;
}

/* Tables */
thead tr th {
    background: #f5f6f7 !important;
    font-size: 11px !important;
    font-weight: 600 !important;
    letter-spacing: 0.07em !important;
    text-transform: uppercase !important;
    color: #a5a5a5 !important;
    padding: 9px 16px !important;
    border-bottom: 1px solid #e5e5e5 !important;
}
thead th:nth-child(2) { text-align: center !important; }
thead th:nth-child(3) { text-align: right !important; }

tbody td {
    padding: 10px 16px !important;
    font-size: 13.5px !important;
    border-bottom: 1px solid #f2f2f2 !important;
    color: #32363a !important;
}
tbody tr:hover td { background: #f0f7ff !important; }
tbody tr:last-child td { border-bottom: none !important; }

/* Buttons */
.stButton > button {
    background: #0070f2 !important;
    color: #fff !important;
    border: none !important;
    border-radius: 4px !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    padding: 8px 20px !important;
    font-family: 'IBM Plex Sans', sans-serif !important;
}
.stButton > button:hover {
    background: #0040b0 !important;
}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {
    background: #223548;
    padding: 0 16px;
    gap: 0;
}
.stTabs [data-baseweb="tab"] {
    color: rgba(255,255,255,0.55) !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    padding: 10px 18px !important;
    border-bottom: 2px solid transparent !important;
    background: transparent !important;
}
.stTabs [aria-selected="true"] {
    color: #fff !important;
    border-bottom-color: #0070f2 !important;
}
.stTabs [data-baseweb="tab-highlight"] { display: none; }
.stTabs [data-baseweb="tab-border"] { display: none; }

/* Download button */
.stDownloadButton > button {
    background: #fff !important;
    color: #0070f2 !important;
    border: 1px solid #0070f2 !important;
    border-radius: 4px !important;
    font-size: 13px !important;
    font-weight: 500 !important;
}
.stDownloadButton > button:hover {
    background: #e8f2ff !important;
}

table { width: 100% !important; border-collapse: collapse; }
</style>

<div class="shellbar">
    <span class="shellbar-title">Tilbakemelding Verktøy</span>
    <span class="shellbar-by">Utviklet av Teodoro</span>
</div>
<div class="info-strip">
    🔒 PDF-en og Excel-filer behandles kun for analyse og lagres ikke. Ingen data deles med tredjepart.
</div>
""", unsafe_allow_html=True)

# Tabs
tab1, tab2 = st.tabs(["📄 Faktura Analyse", "🏗️ iSEKK/Kran"])


# ─────────────────────────────────────────────
# TAB 1: FAKTURA ANALYSE (PDF)
# ─────────────────────────────────────────────
with tab1:
    st.markdown('<div class="warning-strip">Faktura Analyse støtter kun Humlekjær Ødegaard enn så lenge.</div>', unsafe_allow_html=True)

    def extract_invoice_data(pdf_bytes):
        items = {}
        sum_fritt = None

        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue

                fs_match = re.search(r"Sum fritt[:\s]+([\d.,]+)", text)
                if fs_match:
                    try:
                        sum_fritt = float(fs_match.group(1).replace(".", "").replace(",", "."))
                    except ValueError:
                        pass

                for line in text.split("\n"):
                    line = line.strip()
                    if not line.endswith("*"):
                        continue
                    if not re.match(r"\d{2}\.\d{2}\.\d{4}\s+\d{7}", line):
                        continue

                    line_clean = line[:-1].strip()
                    amount_match = re.search(r"(-?[\d.]+,\d{2})\s*$", line_clean)
                    if not amount_match:
                        continue
                    raw_amount = amount_match.group(1)

                    line_body = re.sub(r"^\d{2}\.\d{2}\.\d{4}\s+\d{7}\s+", "", line_clean[:amount_match.start()].strip())
                    name_match = re.match(r"^(.*?)\s+\d[\d ]*,\d{2}", line_body)
                    name = name_match.group(1).strip() if name_match else line_body.strip()
                    name = " ".join(name.split())

                    if re.match(r"^\d+$", name) or not name:
                        dm = re.match(r"(\d{2}\.\d{2}\.\d{4})\s+(\d{7})", line_clean)
                        name = f"[{dm.group(1)} / {dm.group(2)}] {line_body[:50].strip()}" if dm else f"[Ukjent] {line_body[:50].strip()}"

                    name = name.title()

                    try:
                        amount = float(raw_amount.replace(".", "").replace(",", "."))
                    except ValueError:
                        continue

                    items.setdefault(name, []).append(amount)

        summary = {k: (len(v), round(sum(v), 2)) for k, v in items.items()}
        total = round(sum(v[1] for v in summary.values()), 2)
        return summary, total, sum_fritt

    uploaded_pdf = st.file_uploader("Velg PDF-faktura", type=["pdf"], key="pdf")

    if uploaded_pdf:
        pdf_bytes = uploaded_pdf.read()
        if st.button("Analyser faktura", type="primary"):
            with st.spinner("Analyserer faktura..."):
                try:
                    summary, total, sum_fritt = extract_invoice_data(pdf_bytes)

                    if summary:
                        st.success(f"Fant {len(summary)} poster")

                        import pandas as _pd
                        rows = []
                        for name in sorted(summary.keys()):
                            count, summation = summary[name]
                            formatted = f"{summation:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                            is_unknown = name.startswith("[")
                            rows.append({"Navn": name, "Antall": count, "Sum (kr)": formatted, "_ukjent": is_unknown})
                        df_display = _pd.DataFrame(rows)

                        def highlight_ukjent(row):
                            if row["_ukjent"]:
                                return ["background-color: #fff4e0; color: #b85c00; font-style: italic; border-left: 3px solid #f0a500"] * len(row)
                            return [""] * len(row)

                        styled = (
                            df_display[["Navn", "Antall", "Sum (kr)", "_ukjent"]].style
                            .apply(highlight_ukjent, axis=1)
                            .hide(axis="index")
                            .hide(subset=["_ukjent"], axis="columns")
                            .set_table_styles([{"selector": "table", "props": [("width", "100%")]}])
                        )
                        st.write(styled.to_html(), unsafe_allow_html=True)

                        ukjente = [n for n in summary.keys() if n.startswith("[")]
                        if ukjente:
                            st.caption(f"{len(ukjente)} post(er) markert for manuell gjennomgang.")

                        total_fmt = f"{total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

                        if sum_fritt is not None:
                            fs_fmt = f"{sum_fritt:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                            diff = round(abs(total - sum_fritt), 2)
                            if diff <= 1.0:
                                st.markdown(f"""
                                <div style="margin-top:16px; padding:14px 20px; border-radius:4px;
                                            background:#f1fdf6; border:1px solid #a8d8bc;
                                            display:flex; justify-content:space-between; align-items:center;">
                                    <span style="font-size:11px; font-weight:600; letter-spacing:0.07em; text-transform:uppercase; color:#107e3e;">Totalsum</span>
                                    <span style="font-size:22px; font-weight:600; color:#107e3e; font-family:'IBM Plex Mono',monospace;">{total_fmt} kr</span>
                                    <span style="font-size:12px; color:#a5a5a5;">Matcher Sum på fakturaen</span>
                                </div>""", unsafe_allow_html=True)
                            else:
                                diff_fmt = f"{diff:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                                st.markdown(f"""
                                <div style="margin-top:16px; padding:14px 20px; border-radius:4px;
                                            background:#fff0f0; border:1px solid #f5b2b2;
                                            display:flex; justify-content:space-between; align-items:center;">
                                    <span style="font-size:11px; font-weight:600; letter-spacing:0.07em; text-transform:uppercase; color:#bb0000;">Avvik oppdaget</span>
                                    <span style="font-size:22px; font-weight:600; color:#bb0000; font-family:'IBM Plex Mono',monospace;">{total_fmt} kr</span>
                                    <span style="font-size:12px; color:#a5a5a5;">Sum fritt: {fs_fmt} kr &nbsp;|&nbsp; Avvik: {diff_fmt} kr</span>
                                </div>""", unsafe_allow_html=True)
                        else:
                            st.markdown(f"""
                            <div style="margin-top:16px; padding:14px 20px; border-radius:4px;
                                        background:#f5f6f7; border:1px solid #e5e5e5;
                                        display:flex; justify-content:space-between; align-items:center;">
                                <span style="font-size:11px; font-weight:600; letter-spacing:0.07em; text-transform:uppercase; color:#6a6d70;">Totalsum</span>
                                <span style="font-size:22px; font-weight:600; color:#32363a; font-family:'IBM Plex Mono',monospace;">{total_fmt} kr</span>
                                <span style="font-size:12px; color:#a5a5a5;">Sjekk mot faktura manuelt</span>
                            </div>""", unsafe_allow_html=True)
                    else:
                        st.warning("Ingen poster funnet.")
                except Exception as e:
                    st.error(f"Feil: {e}")
    else:
        st.info("Last opp en PDF-faktura for å starte analysen.")


# ─────────────────────────────────────────────
# TAB 2: iSEKK/KRAN (EXCEL)
# ─────────────────────────────────────────────
with tab2:
    st.markdown("Last opp én eller flere Excel-filer fra SAP for å få samlet tilbakemelding for iSEKK/Kran.")

    BAD_UNIT_LABELS = {"", "NAN", "NA", "NONE", "NULL", "TOTAL", "SUM"}

    def clean_unit(u):
        if pd.isna(u):
            return None
        s = str(u).strip().upper()
        return None if s in BAD_UNIT_LABELS else s

    def fmt_number(val):
        if pd.isna(val):
            return ""
        f = float(val)
        if f.is_integer():
            return str(int(f))
        return str(f)

    uploaded_excels = st.file_uploader(
        "Last opp Excel-fil(er)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="excel"
    )

    if uploaded_excels:
        dfs = []
        feil_filer = []

        for uf in uploaded_excels:
            try:
                tmp = pd.read_excel(uf, engine="openpyxl")
                required_cols = ["Betegnelse", "Materialkorttekst", "Målkvantum", "KE.1", "Delsum 3"]
                missing = [c for c in required_cols if c not in tmp.columns]
                if missing:
                    feil_filer.append(f"{uf.name} (mangler kolonner: {missing})")
                    continue
                tmp["Kildefil"] = uf.name
                dfs.append(tmp)
            except Exception as e:
                feil_filer.append(f"{uf.name} (feil ved lesing: {e})")

        if feil_filer:
            st.warning("Noen filer ble hoppet over:")
            for msg in feil_filer:
                st.write("-", msg)

        if not dfs:
            st.error("Ingen av filene kunne brukes. Sjekk at de har riktige kolonner.")
        else:
            df = pd.concat(dfs, ignore_index=True)

            df["Fraksjon"] = df.apply(
                lambda x: x["Materialkorttekst"]
                if str(x["Betegnelse"]) == "Kranbil Isekk - Avfallstype"
                else x["Betegnelse"],
                axis=1,
            )

            df["Målkvantum"] = pd.to_numeric(df["Målkvantum"], errors="coerce")
            df["Delsum 3"] = pd.to_numeric(df["Delsum 3"], errors="coerce")
            df["KE.1"] = df["KE.1"].map(clean_unit)
            df = df[df["KE.1"].notna()].copy()

            if df.empty:
                st.error("Ingen gyldige rader etter rensing.")
            else:
                units_found = sorted(df["KE.1"].unique().tolist())
                units_order = (["KG"] if "KG" in units_found else []) + [u for u in units_found if u != "KG"]

                pivot_vals = df.pivot_table(index="Fraksjon", columns="KE.1", values="Målkvantum", aggfunc="sum")
                pivot_has = df.pivot_table(index="Fraksjon", columns="KE.1", values="Målkvantum", aggfunc=lambda s: s.notna().any())

                safe_cols = [c for c in pivot_vals.columns if (pd.notna(c) and str(c).strip() and str(c).strip().upper() not in BAD_UNIT_LABELS)]
                pivot_vals = pivot_vals.loc[:, safe_cols]
                pivot_has = pivot_has.loc[:, safe_cols]
                pivot_vals = pivot_vals.reindex(columns=units_order)
                pivot_has = pivot_has.reindex(columns=units_order)
                pivot = pivot_vals.where(pivot_has, pd.NA)

                price_pivot = df.pivot_table(index="Fraksjon", values="Delsum 3", aggfunc="sum")

                sum_row = pivot.fillna(0).sum(axis=0).to_frame().T
                sum_row.index = ["SUM"]
                result = pd.concat([pivot, sum_row], axis=0)

                data = result.loc[result.index != "SUM"]
                kg_vals = data["KG"].fillna(0) if "KG" in data.columns else pd.Series(0, index=data.index)
                st_vals = data["ST"].fillna(0) if "ST" in data.columns else pd.Series(0, index=data.index)
                grp = (0 * (kg_vals > 0) + 1 * ((kg_vals <= 0) & (st_vals > 0)) + 2 * ((kg_vals <= 0) & (st_vals <= 0)))
                order = pd.DataFrame({"_g": grp, "_kg": kg_vals, "_st": st_vals}, index=data.index).sort_values(by=["_g", "_kg", "_st"], ascending=[True, False, False]).index
                result = pd.concat([data.loc[order], result.loc[["SUM"]]], axis=0)

                price_series = price_pivot["Delsum 3"] if isinstance(price_pivot, pd.DataFrame) else price_pivot
                prices = price_series.reindex(result.index)
                prices.loc["SUM"] = price_series.sum()
                result["Delsum 3"] = prices

                units_label = " / ".join(units_order)
                mengde_col = units_label

                def build_mengde(row):
                    for unit in units_order:
                        if unit in result.columns:
                            val = row.get(unit, None)
                            if pd.notna(val) and val != 0:
                                return f"{fmt_number(val)} {unit}"
                    return ""

                disp = result.copy()
                disp[mengde_col] = disp.apply(build_mengde, axis=1)
                disp["Delsum 3"] = disp["Delsum 3"].apply(fmt_number)
                disp = disp[[mengde_col, "Delsum 3"]].reset_index()
                if disp.columns[0] != "Fraksjon":
                    disp = disp.rename(columns={disp.columns[0]: "Fraksjon"})

                st.markdown("""
                <style>
                tbody tr:last-child td { background-color: #f5f6f7 !important; font-weight: 700 !important; border-top: 2px solid #e5e5e5 !important; }
                td:nth-child(2) { text-align: right !important; font-weight: 600 !important; }
                td:nth-child(3) { text-align: right !important; font-weight: 600 !important; color: #107e3e !important; }
                </style>""", unsafe_allow_html=True)

                styled = (
                    disp.style
                    .hide(axis="index")
                    .set_table_styles([{"selector": "table", "props": [("width", "100%")]}])
                )

                st.subheader(f"Resultat ({len(uploaded_excels)} fil{'er' if len(uploaded_excels) > 1 else ''} samlet)")
                st.write(styled.to_html(), unsafe_allow_html=True)

                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as w:
                    export_df = result.copy()
                    for col in export_df.columns:
                        export_df[col] = export_df[col].where(~export_df[col].isna(), "")
                    export_df.to_excel(w, sheet_name="Fraksjonsoversikt", index=True)
                output.seek(0)
                st.markdown("<br>", unsafe_allow_html=True)
                st.download_button("Last ned som Excel", output.getvalue(), file_name="fraksjonsoversikt_samlet.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("Last opp én eller flere Excel-filer for å starte.")
