
import math
from datetime import timedelta
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Temperatur – 3-Punkte-Vergleich (App6)", layout="wide")
st.title("Temperatur – Vergleich gegen Referenz (30 °C / 0 °C / −30 °C) – App6")

# ---------------- Sidebar ----------------
st.sidebar.header("Messdatei (IST)")
uploaded_mess = st.sidebar.file_uploader("Mess‑Excel (.xlsx)", type=["xlsx"], key="mess")
mess_sheet = st.sidebar.text_input("Sheet‑Name Messwerte", value="Messwerte")
start_row = st.sidebar.number_input("(Optional) Ab Zeile einlesen – wird automatisch erkannt, wenn 0", min_value=0, value=0, step=1)

st.sidebar.header("Referenzdatei (Soll)")
uploaded_ref = st.sidebar.file_uploader("Referenz‑Excel (.xlsx)", type=["xlsx"], key="ref")
ref_sheet = st.sidebar.text_input("Sheet‑Name Referenz", value="Sheet1")

st.sidebar.header("Vergleichs‑Einstellungen")
window_n = st.sidebar.number_input("Werte je Zielpunkt (n)", min_value=2, value=10, step=1)
ref_tol_min = st.sidebar.number_input("Zeit‑Toleranz beim Matching [Min]", min_value=0, value=5, step=1)
val_tol = st.sidebar.number_input("Wert‑Toleranz |IST−RV| ≤ [°C]", min_value=0.0, value=0.10, step=0.01, format="%.3f")

TARGETS = [30.0, 0.0, -30.0]

# ---------------- Helpers ----------------

def _find_header_row_for_logger(sheet_bytes, sheet_name):
    """Scanne die ersten ~150 Zeilen ohne Header und finde die Zeile,
    die die wiederkehrenden Logger‑Spalten enthält (mindestens 'Datum' und 'Wert').
    """
    df0 = pd.read_excel(sheet_bytes, sheet_name=sheet_name, header=None, engine="openpyxl")
    max_scan = min(150, len(df0))
    header_row = None
    for r in range(max_scan):
        row_vals = [str(x).strip().lower() for x in df0.iloc[r].tolist()]
        if not any(row_vals):
            continue
        # Prüfe, ob in dieser Zeile mindestens je 1x 'datum' und 'wert' vorkommen
        has_datum = any(v.startswith("datum") for v in row_vals)
        has_wert  = any(v.startswith("wert") for v in row_vals)
        # akzeptiere alternativ Englisch
        has_date  = any(v.startswith("date") for v in row_vals)
        has_value = any(v.startswith("value") for v in row_vals)
        cond = (has_datum and has_wert) or (has_date and has_value)
        if cond:
            header_row = r
            break
    return header_row


def read_mess(file_like, sheet, start_excel_row):
    """Robustes Einlesen von ZKL/Logger‑Exports.
    1) Versuche automatische Header‑Erkennung (Datum*/Wert* in einer Zeile)
    2) Fallback: benutze start_excel_row, wenn >0
    Rückgabe: DataFrame mit Spalten Timestamp, Temperatur
    """
    # 1) Autodetektion
    header_row = _find_header_row_for_logger(file_like, sheet)
    if header_row is not None:
        df = pd.read_excel(file_like, sheet_name=sheet, header=header_row, engine="openpyxl")
    else:
        # 2) Fallback: start_excel_row benutzen
        if start_excel_row and start_excel_row > 0:
            df = pd.read_excel(file_like, sheet_name=sheet, skiprows=start_excel_row-1, header=0, engine="openpyxl")
        else:
            raise ValueError("Messdatei: Konnte Kopfzeile nicht automatisch erkennen – bitte 'Ab Zeile' setzen.")

    # Leere Spalten/Zeilen entfernen
    df = df.dropna(how="all").dropna(axis=1, how="all")

    # Kandidaten sammeln (Deutsch & Englisch)
    datum_candidates = [c for c in df.columns if str(c).strip().lower().startswith(("datum","date"))]
    wert_candidates  = [c for c in df.columns if str(c).strip().lower().startswith(("wert","value","temperatur","temperature"))]

    if len(datum_candidates) == 0 or len(wert_candidates) == 0:
        # Hilfsdiagnose ausgeben
        raise ValueError(
            f"Messdatei: Konnte keine Datum*/Wert*‑Spalten finden. Gefundene Spalten: {list(df.columns)}"
        )

    # Paare in gleicher Reihenfolge bilden
    pairs = []
    for d, w in zip(datum_candidates, wert_candidates):
        block = df[[d, w]].copy()
        block.columns = ["Timestamp_raw", "Temperatur_raw"]
        pairs.append(block)

    if not pairs:
        raise ValueError("Messdatei: Es konnten keine (Datum, Wert)‑Paare gebildet werden.")

    df_all = pd.concat(pairs, ignore_index=True)

    # Konvertieren
    df_all["Timestamp"] = pd.to_datetime(df_all["Timestamp_raw"], dayfirst=True, errors="coerce")
    df_all["Temperatur"] = (
        df_all["Temperatur_raw"].astype(str).str.replace(",", ".", regex=False).astype(float)
    )

    df_all = df_all.dropna(subset=["Timestamp", "Temperatur"]).sort_values("Timestamp")
    return df_all[["Timestamp", "Temperatur"]]


def read_reference_file(file_like, sheet):
    df_raw = pd.read_excel(file_like, sheet_name=sheet, header=None, engine="openpyxl")
    start_time_raw = df_raw.iat[0, 1] if df_raw.shape[1] > 1 else None
    start_time = pd.to_datetime(start_time_raw, dayfirst=True, errors="coerce")
    if pd.isna(start_time):
        raise ValueError(f"Referenz: Startzeit unlesbar in B1: {start_time_raw}")

    df = pd.read_excel(file_like, sheet_name=sheet, header=1, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    if not {"time", "RV"}.issubset(set(df.columns)):
        raise ValueError(f"Referenz: Erwartete Spalten 'time' und 'RV' fehlen. Gefunden: {list(df.columns)}")

    df["time"] = pd.to_timedelta(df["time"].astype(str))
    df["Timestamp"] = start_time + df["time"]
    df["RV"] = df["RV"].astype(str).str.replace(",", ".", regex=False).astype(float)

    out = df[["Timestamp", "RV"]].dropna().sort_values("Timestamp").reset_index(drop=True)
    return out


def pick_ref_block(df_ref, target, n):
    if df_ref.empty:
        return df_ref
    idx = (df_ref["RV"] - target).abs().idxmin()
    start = max(0, idx - n // 2)
    end = min(len(df_ref), start + n)
    start = max(0, end - n)
    return df_ref.iloc[start:end].copy()

# ---------------- Main ----------------

if uploaded_mess is None or uploaded_ref is None:
    st.info("Bitte Mess- und Referenzdatei hochladen.")
else:
    try:
        df_mess = read_mess(uploaded_mess, mess_sheet, start_row)
        df_ref = read_reference_file(uploaded_ref, ref_sheet)

        st.success("Dateien erfolgreich eingelesen.")
        st.caption(f"Messwerte: {len(df_mess)} Zeilen • Referenz: {len(df_ref)} Zeilen")

        all_results = []
        for target in TARGETS:
            ref_block = pick_ref_block(df_ref, target, int(window_n))
            cmp = pd.merge_asof(
                ref_block.sort_values("Timestamp"),
                df_mess.sort_values("Timestamp"),
                on="Timestamp",
                direction="nearest",
                tolerance=timedelta(minutes=int(ref_tol_min)),
            )

            cmp["Abweichung"] = cmp["Temperatur"] - cmp["RV"]
            cmp["OK"] = cmp["Abweichung"].abs() <= float(val_tol)

            ok_count = int(cmp["OK"].sum())
            total = len(cmp)

            # --- Statistik ---
            mean_ist = cmp["Temperatur"].mean()
            mean_rv = cmp["RV"].mean()
            mean_abw = cmp["Abweichung"].mean()
            s = cmp["Abweichung"].std(ddof=1)
            u_A = s / (len(cmp) ** 0.5)

            st.subheader(f"Zielpunkt {target:+.0f} °C — Ergebnis: {ok_count}/{total} OK")
            st.markdown(f"""
**Statistik:**
- Mittelwert IST: **{mean_ist:.3f} °C**
- Mittelwert RV: **{mean_rv:.3f} °C**
- Mittelwert Abweichung: **{mean_abw:.4f} °C**
- Standardabweichung s: **{s:.4f} °C**
- Messunsicherheit uₐ = s / √n: **{u_A:.4f} °C**
""")

            st.dataframe(cmp, use_container_width=True)
            all_results.append((target, cmp))

        # Export
        export_frames = []
        for target, cmp in all_results:
            t = cmp.copy()
            t.insert(0, "Zielpunkt", target)
            export_frames.append(t)

        if export_frames:
            out_csv = pd.concat(export_frames, ignore_index=True).to_csv(index=False).encode("utf-8")
            st.download_button("Ergebnisse als CSV herunterladen", out_csv, "vergleich_3_punkte_app6.csv", mime="text/csv")

    except Exception as e:
        st.error(f"Fehler: {e}")
