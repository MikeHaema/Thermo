
# App15: Summary table across targets + collapsible top-10 + multi-sensor, k=2, Seriennummer
import math
from datetime import timedelta
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Temperatur – Vergleich (App15)", layout="wide")
st.title("Temperatur – Vergleich gegen Referenz (App15)")

# ----------------------- Minimal UI ---------------------------
uploaded_mess_files = st.file_uploader(
    "Messdateien (IST) – eine oder mehrere Excel‑Dateien (.xlsx)",
    type=["xlsx"],
    key="mess",
    accept_multiple_files=True,
)
uploaded_ref  = st.file_uploader("Referenzdatei (Soll) – Excel (.xlsx)", type=["xlsx"], key="ref")

# feste Parameter
mess_sheet_data = "Messwerte"    # Reiter mit den Loggerdaten
mess_sheet_over = "Übersicht"    # Reiter mit Seriennummer (G12)
ref_sheet = "Sheet1"
window_n = 10
ref_tol_min = 5
RV_start = 0.05
RV_expand_step = 0.01
RV_expand_max = 2.0
TARGETS = [30.0, 0.0, -30.0]
ROW_ORDER = ["+30 °C", "0 °C", "-30 °C"]

# --------------------------------------------------
# Helpers
# --------------------------------------------------

def _find_header_row(sheet_bytes, sheet_name):
    df0 = pd.read_excel(sheet_bytes, sheet_name=sheet_name, header=None, engine="openpyxl")
    for r in range(min(150, len(df0))):
        row = [str(x).strip().lower() for x in df0.iloc[r].tolist()]
        if any(x.startswith(("datum","date")) for x in row) and any(x.startswith(("wert","value","temperatur","temperature")) for x in row):
            return r
    return None

# Seriennummer aus Übersicht!G12

def read_serial_number(file_like, overview_sheet_name="Übersicht"):
    try:
        df = pd.read_excel(file_like, sheet_name=overview_sheet_name, header=None, engine="openpyxl")
        val = df.iat[11, 6]
        return None if pd.isna(val) else str(val).strip()
    except Exception:
        return None

# Messdaten (breiter Logger‑Export)

def read_mess(file_like, sheet):
    hr = _find_header_row(file_like, sheet)
    if hr is None:
        raise ValueError("Messdatei: Kopfzeile nicht gefunden.")
    df = pd.read_excel(file_like, sheet_name=sheet, header=hr, engine="openpyxl")
    df = df.dropna(how="all").dropna(axis=1, how="all")
    dcols=[c for c in df.columns if str(c).lower().startswith(("datum","date"))]
    wcols=[c for c in df.columns if str(c).lower().startswith(("wert","value","temperatur","temperature"))]
    parts=[]
    for d,w in zip(dcols,wcols):
        b=df[[d,w]].copy(); b.columns=["Timestamp_raw","Temperatur_raw"]; parts.append(b)
    df_all=pd.concat(parts, ignore_index=True)
    df_all["Timestamp"]=pd.to_datetime(df_all["Timestamp_raw"], dayfirst=True, errors="coerce")
    df_all["Temperatur"]=df_all["Temperatur_raw"].astype(str).str.replace(",",".").astype(float)
    return df_all.dropna(subset=["Timestamp","Temperatur"]).sort_values("Timestamp")[["Timestamp","Temperatur"]]

# Referenzdaten (Thermoscan Sheet1: Start in B1, time/RV ab Zeile 2)

def read_reference_file(file_like, sheet):
    df_raw=pd.read_excel(file_like, sheet_name=sheet, header=None, engine="openpyxl")
    start_raw=df_raw.iat[0,1]
    start_ts=pd.to_datetime(start_raw, dayfirst=True, errors="coerce")
    df=pd.read_excel(file_like, sheet_name=sheet, header=1, engine="openpyxl")
    df.columns=[str(c).strip() for c in df.columns]
    df["time"]=pd.to_timedelta(df["time"].astype(str))
    df["Timestamp"]=start_ts+df["time"]
    df["RV"]=df["RV"].astype(str).str.replace(",",".").astype(float)
    return df[["Timestamp","RV"]].dropna().sort_values("Timestamp")

# Block um Zielwert

def pick_ref_block(df_ref, target, secs):
    idx=(df_ref["RV"]-target).abs().idxmin()
    start=max(0, idx - secs//2); end=min(len(df_ref), start+secs)
    start=max(0, end-secs)
    return df_ref.iloc[start:end].copy()

# --------------------------------------------------
# Main
# --------------------------------------------------
if uploaded_ref is None or not uploaded_mess_files:
    st.info("Bitte mindestens eine Messdatei und die Referenzdatei hochladen.")
else:
    try:
        df_ref = read_reference_file(uploaded_ref, ref_sheet)
        st.success(f"Referenzdatei eingelesen. Messdateien: {len(uploaded_mess_files)}")

        export_frames=[]

        # -- Für jede Messdatei (Sensor) --
        for file_ix, mess_file in enumerate(uploaded_mess_files, start=1):
            with st.container(border=True):
                st.markdown(f"### Sensor {file_ix}")

                serial = read_serial_number(mess_file, overview_sheet_name=mess_sheet_over)
                if serial:
                    st.caption(f"Seriennummer: **{serial}** (aus {mess_sheet_over}!G12)")
                else:
                    st.caption("Seriennummer: — (nicht gefunden)")

                # Messdaten lesen
                df_mess = read_mess(mess_file, mess_sheet_data)

                # Sammle Kennzahlen pro Zielpunkt für die Zusammenfassung
                summary_rows = []

                # Ergebnisse je Zielpunkt
                for target in TARGETS:
                    sec_block=pick_ref_block(df_ref, target, window_n*60)

                    ref_min=(sec_block.assign(Timestamp_min=lambda x:x["Timestamp"].dt.floor("min"))
                             .sort_values("Timestamp")
                             .drop_duplicates("Timestamp_min", keep="first"))

                    mess_min=(df_mess.assign(Timestamp_min=lambda x:x["Timestamp"].dt.floor("min"))
                              .groupby("Timestamp_min", as_index=False)
                              .agg(Temperatur=("Temperatur","mean")))

                    cmp=pd.merge_asof(
                        ref_min.sort_values("Timestamp_min"),
                        mess_min.sort_values("Timestamp_min"),
                        on="Timestamp_min",
                        direction="nearest",
                        tolerance=pd.Timedelta(minutes=ref_tol_min)
                    )

                    cmp["Abweichung"]=cmp["Temperatur"]-cmp["RV"]
                    cmp=cmp.dropna(subset=["Temperatur","RV"])  # safety

                    # Dynamische RV-Fenster-Erweiterung
                    cmp["RV_Dist"]=(cmp["RV"]-target).abs()
                    rv_limit=RV_start
                    while True:
                        cmp_try=cmp[cmp["RV_Dist"]<=rv_limit]
                        if len(cmp_try)>=window_n:
                            cmp=cmp_try; break
                        rv_limit+=RV_expand_step
                        if rv_limit>RV_expand_max:
                            cmp=cmp_try; break

                    if cmp.empty:
                        st.warning(f"Keine geeigneten RV-Punkte bei {target} °C für Sensor {file_ix} gefunden.")
                        continue

                    # Ranking nach RV-Nähe und IST-Abweichung
                    cmp["IST_Dist"]=(cmp["Temperatur"]-cmp["RV"]).abs()
                    cmp_best=(cmp.sort_values(by=["RV_Dist","IST_Dist"], ascending=[True,True])
                                .head(window_n).reset_index(drop=True))

                    # Statistik + k=2
                    n=len(cmp_best)
                    s=cmp_best["Abweichung"].std(ddof=1)
                    uA=s/(n**0.5)
                    U=2*uA
                    mean_ist=cmp_best["Temperatur"].mean()
                    mean_rv=cmp_best["RV"].mean()
                    mean_abw=cmp_best["Abweichung"].mean()

                    # --- Zusammenfassung je Zielpunkt sammeln ---
                    row_name = f"{'+30' if target==30.0 else ('0' if target==0.0 else '-30')} °C"
                    summary_rows.append({
                        "Zielpunkt": row_name,
                        "Mittelwert RV [°C]": round(mean_rv, 3),
                        "Mittelwert IST [°C]": round(mean_ist, 3),
                        "Mittelwert Abweichung [°C]": round(mean_abw, 3),
                        "U (k=2) [°C]": round(U, 3),
                    })

                    # --- Kompakter Block mit einklappbarer Top‑10 ---
                    st.subheader(f"Zielpunkt {target:+.0f} °C — beste {window_n} Minuten (k=2)")
                    st.markdown(
                        f"**Statistik:**  Mittelwert IST **{mean_ist:.3f} °C**,  Mittelwert RV **{mean_rv:.3f} °C**,  "
                        f"Mittelwert Abweichung **{mean_abw:.3f} °C**,  s **{s:.3f} °C**,  uₐ **{uA:.3f} °C**,  U(k=2) **{U:.3f} °C**"
                    )

                    with st.expander("Top‑10 Werte anzeigen/ausblenden", expanded=False):
                        st.dataframe(
                            cmp_best.drop(columns=["RV_Dist","IST_Dist"]),
                            column_config={
                                "RV": st.column_config.NumberColumn(format="%.3f"),
                                "Temperatur": st.column_config.NumberColumn(format="%.3f"),
                                "Abweichung": st.column_config.NumberColumn(format="%.3f"),
                            },
                            use_container_width=True,
                        )

                    # Export vorbereiten (mit Seriennummer & Sensorindex)
                    ex = cmp_best.drop(columns=["RV_Dist","IST_Dist"]).copy()
                    ex.insert(0, "Zielpunkt", row_name)
                    if serial:
                        ex.insert(0, "Seriennummer", serial)
                    ex.insert(0, "Sensor", file_ix)
                    export_frames.append(ex)

                # ---- Zusammenfassungstabelle unterhalb aller Zielpunkte je Sensor ----
                if summary_rows:
                    # Gewünschte Reihenfolge der Zeilen
                    df_sum = pd.DataFrame(summary_rows)
                    df_sum.index = df_sum["Zielpunkt"]
                    df_sum = df_sum.loc[[r for r in ROW_ORDER if r in df_sum.index]]
                    df_sum = df_sum.drop(columns=["Zielpunkt"])  # Zielpunkt steht als Index (links)

                    st.markdown("#### Zusammenfassung (alle Zielpunkte)")
                    st.dataframe(
                        df_sum,
                        use_container_width=True,
                    )

        # Gesamtexport (alle Sensoren, alle Zielpunkte)
        if export_frames:
            out = pd.concat(export_frames, ignore_index=True).to_csv(index=False).encode("utf-8")
            st.download_button("CSV Export (App15, alle Sensoren)", out, "vergleich_app15_all.csv", mime="text/csv")

    except Exception as e:
        st.error(f"Fehler: {e}")
