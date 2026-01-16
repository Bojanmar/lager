import streamlit as st
import pandas as pd
import plotly.express as px

from obracun import (
    procesiraj_obracun,
    rules_to_df,
    RULE_TYPES,
    export_to_excel,
    material_rules
)

st.set_page_config(page_title="Obračun zaliha", layout="wide")
st.title("📦 Obračun zaliha – Laser Lux")

# ---------------- Sidebar ----------------
st.sidebar.header("Ulazni podaci")
lager_file = st.sidebar.file_uploader("Lager Excel (ULAZ)", type=["xlsx"])
fakture_file = st.sidebar.file_uploader("Fakture Excel (IZLAZ)", type=["xlsx"])
magacin_file = st.sidebar.file_uploader("Stvarno stanje magacina (31.12)", type=["xlsx"])
st.sidebar.markdown("---")

# ======================================================
# Manual mapping (IZLAZ -> LAGER)
# ======================================================
st.sidebar.subheader("🔁 Ručno mapiranje (IZLAZ → LAGER)")

if "manual_map" not in st.session_state:
    st.session_state["manual_map"] = pd.DataFrame(columns=["Materijal_izlaz", "Materijal_lager"])

lager_options = []
if lager_file is not None:
    try:
        lager_file.seek(0)
        df_hdr = pd.read_excel(lager_file, nrows=0)
        lager_options = sorted(df_hdr.columns.tolist())
    except Exception:
        lager_options = []

mm = st.session_state["manual_map"].copy()

# Sidebar editor: user može da doda/izmeni mapping i pre obračuna
edited_mm = st.sidebar.data_editor(
    mm,
    use_container_width=True,
    num_rows="dynamic",
    height=220,
    column_config={
        "Materijal_izlaz": st.column_config.TextColumn("Materijal_izlaz"),
        "Materijal_lager": st.column_config.SelectboxColumn("Materijal_lager", options=[""] + lager_options)
    }
)

mc1, mc2 = st.sidebar.columns(2)
with mc1:
    if st.sidebar.button("💾 Sačuvaj mapping"):
        st.session_state["manual_map"] = edited_mm
        st.sidebar.success("Sačuvano ✅")
with mc2:
    if st.sidebar.button("↩ Reset mapping"):
        st.session_state["manual_map"] = pd.DataFrame(columns=["Materijal_izlaz", "Materijal_lager"])
        st.sidebar.info("Resetovano.")

st.sidebar.markdown("---")

# ============================
# Pravila (koeficijenti)
# ============================
st.sidebar.subheader("🧮 Pravila (koeficijenti)")

if "rules_df" not in st.session_state:
    st.session_state["rules_df"] = rules_to_df(material_rules)

rules_df = st.session_state["rules_df"]

edited_rules_df = st.sidebar.data_editor(
    rules_df,
    use_container_width=True,
    num_rows="dynamic",
    height=420,
    column_config={
        "rule_type": st.column_config.SelectboxColumn("rule_type", options=RULE_TYPES),
        "enabled": st.column_config.CheckboxColumn("enabled"),
        "factor": st.column_config.NumberColumn("factor"),
        "extra": st.column_config.NumberColumn("extra"),
    }
)

c1, c2 = st.sidebar.columns(2)
with c1:
    if st.button("💾 Sačuvaj pravila"):
        st.session_state["rules_df"] = edited_rules_df
        st.success("Sačuvano ✅ (primeniće se na sledeći obračun)")
with c2:
    if st.button("↩ Reset pravila"):
        st.session_state["rules_df"] = rules_to_df(material_rules)
        st.info("Vraćeno na podrazumevano.")

st.sidebar.markdown("---")
run_calc = st.sidebar.button("🚀 Obračunaj zalihe")

# --------------- Main area ---------------
if run_calc:
    if not lager_file or not fakture_file:
        st.error("❌ Morate da uploadujete **lager** i **fakture**.")
    else:
        rules_df_for_run = st.session_state.get("rules_df", edited_rules_df)
        manual_map_df = st.session_state.get("manual_map")

        with st.spinner("Računam..."):
            uporedba, df_fakture_posle, rules_used_df, kal_ekstremi, injekt_debug, audit_map, sus_bad = procesiraj_obracun(
                lager_file,
                fakture_file,
                magacin_file=magacin_file,
                edited_rules_df=rules_df_for_run,
                manual_map_df=manual_map_df
            )

        st.success("✔ Obračun je završen.")

        tab1, tab2, tab3, tab_map, tab_audit, tab4, tab5, tab6 = st.tabs(
            [
                "📌 Uporedba",
                "⚠ Ekstremi kalibracije",
                "📄 Fakture – obračun",
                "🔎 Mapiranje IZLAZ↔LAGER",
                "🧾 Audit (JM iste ≠ 1)",
                "🧮 Pravila (primenjena)",
                "📊 Grafikon",
                "🧪 Injekt debug"
            ]
        )

        with tab1:
            st.subheader("Uporedba (ključne kolone + novi koef)")
            cols = [
                "Materijal",
                "Ukupna_potrošnja",
                "Stanje_na_lageru",
                "Razlika_pre",
                "Stanje_na_magacinu",
                "Koef_novi",
                "Finalna_potrošnja",
            ]
            show = uporedba.copy()
            for c in ["Ukupna_potrošnja","Stanje_na_lageru","Razlika_pre","Stanje_na_magacinu","Finalna_potrošnja","Koef_novi"]:
                if c in show.columns:
                    show[c] = pd.to_numeric(show[c], errors="coerce")
            st.dataframe(show[cols], use_container_width=True)

        with tab2:
            st.subheader("Ekstremi kalibracije (Koef_novi van očekivanog opsega)")
            st.dataframe(kal_ekstremi, use_container_width=True)

        with tab3:
            st.subheader("Fakture – obračun količina za skidanje sa lagera")
            st.dataframe(df_fakture_posle, use_container_width=True)

        # ======================================================
        # TAB: Mapiranje IZLAZ↔LAGER + filteri + auto-popunjavanje mapping tabele
        # ======================================================
        with tab_map:
            st.subheader("🔎 Provera mapiranja materijala i jedinica (IZLAZ ↔ LAGER)")

            if audit_map is None or audit_map.empty:
                st.info("Nema audit podataka.")
            else:
                # ---- FILTERI ----
                fc1, fc2, fc3 = st.columns([1, 1, 2])
                with fc1:
                    only_missing = st.checkbox("Samo nematchovani", value=False)
                with fc2:
                    only_jm_mismatch = st.checkbox("Samo JM ne odgovara", value=False)
                with fc3:
                    q = st.text_input("Pretraga (materijal)", value="")

                view = audit_map.copy()

                if "Match_u_lageru" in view.columns and only_missing:
                    view = view[view["Match_u_lageru"] == False]

                if only_jm_mismatch and "JM_iste" in view.columns:
                    view = view[(view["JM_iste"] == False) & (view["JM_fakt"].fillna("") != "")]

                if q.strip():
                    qq = q.strip().lower()
                    col_lager = view["Materijal_lager"] if "Materijal_lager" in view.columns else ""
                    view = view[
                        view["Materijal"].astype(str).str.lower().str.contains(qq, na=False) |
                        col_lager.astype(str).str.lower().str.contains(qq, na=False)
                    ]

                cols = [
                    "Materijal_original",
                    "Materijal", "Materijal_lager", "Match_u_lageru",
                    "Jedinica mere za fakturisanje", "Jedinica mere za lager - skidanje količine",
                    "JM_fakt", "JM_lager_mapirana", "JM_iste",
                    "Količina za fakturisanje", "Količina za skidanje sa lagera",
                    "Napomena konverzije"
                ]
                cols = [c for c in cols if c in view.columns]
                st.dataframe(view[cols], use_container_width=True)

                st.markdown("---")
                st.subheader("🔁 Ručno mapiranje nematchovanih (klikni Sačuvaj, pa ponovo Obračunaj)")

                missing_unique = audit_map[audit_map["Match_u_lageru"] == False][["Materijal"]].drop_duplicates()
                missing_list = missing_unique["Materijal"].tolist()

                if len(missing_list) == 0:
                    st.success("Nema nematchovanih materijala ✅")
                else:
                    # Lager opcije: koristi iz kolone audit_map Materijal_lager, fallback na sidebar options
                    local_lager_options = sorted(audit_map["Materijal_lager"].dropna().unique().tolist()) if "Materijal_lager" in audit_map.columns else []
                    if not local_lager_options:
                        local_lager_options = lager_options

                    mm2 = st.session_state["manual_map"].copy()
                    existing = set(mm2["Materijal_izlaz"].astype(str).tolist()) if not mm2.empty and "Materijal_izlaz" in mm2.columns else set()

                    # auto-dodaj missing u mapping tabelu
                    rows_to_add = []
                    for m in missing_list:
                        if str(m) not in existing:
                            rows_to_add.append({"Materijal_izlaz": str(m), "Materijal_lager": ""})
                    if rows_to_add:
                        mm2 = pd.concat([mm2, pd.DataFrame(rows_to_add)], ignore_index=True)

                    edited_mm2 = st.data_editor(
                        mm2,
                        use_container_width=True,
                        num_rows="dynamic",
                        column_config={
                            "Materijal_izlaz": st.column_config.TextColumn("Materijal_izlaz", disabled=True),
                            "Materijal_lager": st.column_config.SelectboxColumn("Materijal_lager", options=[""] + local_lager_options)
                        }
                    )

                    sm1, sm2 = st.columns(2)
                    with sm1:
                        if st.button("💾 Sačuvaj ručno mapiranje (iz ovog taba)"):
                            st.session_state["manual_map"] = edited_mm2
                            st.success("Sačuvano ✅ — sada klikni opet 'Obračunaj zalihe'.")
                    with sm2:
                        st.info("Tip: mapiraj na tačan naziv kolone iz lager fajla.")

        # ======================================================
        # TAB: Audit JM iste, koef != 1
        # ======================================================
        with tab_audit:
            st.subheader("🧾 Audit: Jedinice su iste, a konverzija nije 1.0 (ne bi smelo)")

            if sus_bad is None or sus_bad.empty:
                st.success("Nema sumnjivih slučajeva ✅")
            else:
                ac1, ac2 = st.columns([1, 2])
                with ac1:
                    min_diff = st.number_input("Min odstupanje (|koef-1|)", value=0.001, step=0.001)
                with ac2:
                    qq = st.text_input("Pretraga (materijal) ", value="")

                view = sus_bad.copy()
                if "Koef_konverzije" in view.columns:
                    view = view[(view["Koef_konverzije"] - 1.0).abs() >= float(min_diff)]

                if qq.strip():
                    qqq = qq.strip().lower()
                    view = view[view["Materijal"].astype(str).str.lower().str.contains(qqq, na=False)]

                cols = [
                    "Materijal",
                    "Jedinica mere za fakturisanje",
                    "Jedinica mere za lager - skidanje količine",
                    "Količina za fakturisanje",
                    "Količina za skidanje sa lagera",
                    "Koef_konverzije",
                    "Napomena konverzije"
                ]
                cols = [c for c in cols if c in view.columns]
                st.dataframe(view[cols], use_container_width=True)

        with tab4:
            st.subheader("Pravila konverzije (primenjena u ovom obračunu)")
            st.dataframe(rules_used_df, use_container_width=True)

        with tab5:
            st.subheader("Poređenje stanja na lageru i ukupne potrošnje po materijalu")
            chart_df = uporedba.dropna(subset=["Materijal", "Stanje_na_lageru", "Ukupna_potrošnja"]).copy()
            if len(chart_df) > 0:
                max_items = st.slider(
                    "Broj materijala na grafikonu (po razlici, apsolutna vrednost):",
                    min_value=5,
                    max_value=min(100, len(chart_df)),
                    value=min(30, len(chart_df)),
                    step=1
                )
                chart_df["Abs_diff"] = (chart_df["Stanje_na_lageru"] - chart_df["Ukupna_potrošnja"]).abs()
                chart_df = chart_df.sort_values("Abs_diff", ascending=False).head(max_items)

                fig = px.bar(
                    chart_df,
                    x="Materijal",
                    y=["Stanje_na_lageru", "Ukupna_potrošnja"],
                    barmode="group",
                    title="Stanje na lageru vs ukupna potrošnja",
                )
                fig.update_layout(xaxis_tickangle=45, height=650, legend_title_text="Količina")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Nema dovoljno podataka za grafikon.")

        with tab6:
            st.subheader("Injekt debug (paketi: 1.0–1.8x da se uklopi na 31.12)")
            if injekt_debug is None or injekt_debug.empty:
                st.info("Nema injekt računa ili nema magacin fajla.")
            else:
                st.dataframe(injekt_debug, use_container_width=True)

        st.subheader("📎 Preuzimanje Excel rezultata")
        buffer = export_to_excel(
            uporedba,
            df_fakture_posle,
            injekt_debug=injekt_debug,
            audit_map=audit_map,
            sus_bad=sus_bad
        )
        st.download_button(
            label="💾 Preuzmi Excel rezultat",
            data=buffer,
            file_name="obracun_zaliha.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("⬅ Uploaduj fajlove i klikni na **'Obračunaj zalihe'**.")
