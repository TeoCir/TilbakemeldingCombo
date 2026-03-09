import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
from io import BytesIO

st.set_page_config(page_title="Tilbakemelding Verktøy", page_icon="⚙️", layout="wide")

# Header
st.title("⚙️ Tilbakemelding Verktøy")
st.markdown(
    "<div style='margin-bottom: 0.5em; color: #888; font-size: 0.9em;'>"
    "Utviklet av <strong>Teodoro</strong>"
    "</div>",
    unsafe_allow_html=True
)
st.info("🔒 PDF-en og Excel-filer behandles kun for analyse og lagres ikke. Ingen data deles med tredjepart.")

# Tabs
tab1, tab2 = st.tabs(["📄 Faktura Analyse", "🏗️ iSEKK/Kran"])


# ─────────────────────────────────────────────
# TAB 1: FAKTURA ANALYSE (PDF)
# ─────────────────────────────────────────────
with tab1:
    st.subheader("📄 Faktura Analyse")
    st.markdown("Last opp en faktura-PDF og få automatisk gruppert oversikt over alle poster.")
    st.warning("⚠️ Faktura Analyse støtter kun **Humlekjær Ødegaard** enn så lenge.")

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
        if st.button("🔍 Analyser faktura", type="primary"):
            with st.spinner("Analyserer faktura..."):
                try:
                    summary, total, sum_fritt = extract_invoice_data(pdf_bytes)

                    if summary:
                        st.success(f"✅ Fant {len(summary)} poster")

                        col1, col2, col3 = st.columns([3, 1, 2])
                        col1.markdown("**Navn**")
                        col2.markdown("**Antall**")
                        col3.markdown("**Sum**")
                        st.divider()

                        for name in sorted(summary.keys()):
                            count, summation = summary[name]
                            col1, col2, col3 = st.columns([3, 1, 2])
                            col1.write(name)
                            col2.write(str(count))
                            formatted = f"{summation:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                            col3.write(formatted)

                        st.divider()
                        total_fmt = f"{total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                        st.markdown(f"**BEREGNET TOTALSUM: {total_fmt}**")

                        if sum_fritt is not None:
                            fs_fmt = f"{sum_fritt:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                            diff = round(abs(total - sum_fritt), 2)
                            if diff <= 1.0:
                                st.success(f"✅ Matcher Sum fritt: {fs_fmt} (diff: {diff:.2f})")
                            else:
                                diff_fmt = f"{diff:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                                st.error(f"⚠️ Sum fritt fra PDF: {fs_fmt} — avvik: {diff_fmt} kr. Sjekk [Ukjent]-poster.")
                        else:
                            st.info("ℹ️ Kunne ikke lese Sum fritt automatisk. Sjekk manuelt.")
                    else:
                        st.warning("⚠️ Ingen poster funnet.")
                except Exception as e:
                    st.error(f"❌ Feil: {e}")
    else:
        st.info("👆 Last opp en PDF-faktura for å starte analysen.")


# ─────────────────────────────────────────────
# TAB 2: FRAKSJONSOVERSIKT (EXCEL)
# ─────────────────────────────────────────────
with tab2:
    st.subheader("🏗️ iSEKK/Kran")
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
        "Last opp en eller flere Excel-filer — husk å fjern SD Dokument etter bruk!",
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

                disp = result.copy()
                for col in disp.columns:
                    disp[col] = disp[col].apply(fmt_number)
                disp = disp.reset_index()
                if disp.columns[0] != "Fraksjon":
                    disp = disp.rename(columns={disp.columns[0]: "Fraksjon"})

                styled = (
                    disp.style
                    .set_properties(subset=["Fraksjon"], **{"font-weight": "bold", "font-size": "15px"})
                    .set_properties(**{"font-size": "14px"})
                    .set_table_styles([
                        {"selector": "table", "props": [("width", "100%"), ("table-layout", "auto")]},
                        {"selector": "th", "props": [("text-align", "left"), ("white-space", "nowrap")]},
                        {"selector": "td", "props": [("white-space", "nowrap")]},
                    ])
                )

                st.markdown("""
                <style>
                [data-testid="stHorizontalBlock"] { width: 100% !important; }
                .stDataFrame, .element-container { width: 100% !important; }
                table { width: 100% !important; }
                </style>""", unsafe_allow_html=True)

                st.subheader("Resultat (alle filer samlet)")
                st.write(styled.to_html(), unsafe_allow_html=True)

                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as w:
                    export_df = result.copy()
                    for col in export_df.columns:
                        export_df[col] = export_df[col].where(~export_df[col].isna(), "")
                    export_df.to_excel(w, sheet_name="Fraksjonsoversikt", index=True)
                output.seek(0)
                st.download_button("📥 Last ned som Excel", output.getvalue(), file_name="fraksjonsoversikt_samlet.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("👆 Last opp en eller flere Excel-filer for å starte.")


# Footer
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: grey; font-size: 0.85em;'>"
    "⚙️ Utviklet av <strong>Teodoro</strong>"
    "</div>",
    unsafe_allow_html=True
)
