import streamlit as st
import pandas as pd
import plotly.express as px

from obracun import (
    procesiraj_obracun,
    rules_to_df,
    RULE_TYPES,
    export_to_excel,
    material_rules,

    # ✅ eksport po računu
    generate_word_for_racun,
    generate_word_zip_all_racuni,
    generate_excel_for_racun,
    generate_excel_zip_all_racuni
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

edited_mm = st.sidebar.data_editor(
    mm,
    use_container_width=True,
    num_rows="dynamic",
    height=220,
    column_config={
        "Materijal_izlaz": st.column_config.TextColumn("Materijal_izlaz"),
        "Materijal_lager": st.column_config.SelectboxColumn("Materijal_lager", options=[""] + lager_options)
    },
    key="sidebar_manual_map_editor"
)

mc1, mc2 = st.sidebar.columns(2)
with mc1:
    if st.sidebar.button("💾 Sačuvaj mapping", key="btn_save_mapping_sidebar"):
        st.session_state["manual_map"] = edited_mm
        st.sidebar.success("Sačuvano ✅")
with mc2:
    if st.sidebar.button("↩ Reset mapping", key="btn_reset_mapping_sidebar"):
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
    },
    key="sidebar_rules_editor"
)

c1, c2 = st.sidebar.columns(2)
with c1:
    if st.sidebar.button("💾 Sačuvaj pravila", key="btn_save_rules"):
        st.session_state["rules_df"] = edited_rules_df
        st.sidebar.success("Sačuvano ✅ (primeniće se na sledeći obračun)")
with c2:
    if st.sidebar.button("↩ Reset pravila", key="btn_reset_rules"):
        st.session_state["rules_df"] = rules_to_df(material_rules)
        st.sidebar.info("Vraćeno na podrazumevano.")

st.sidebar.markdown("---")

# Dugmad za obračun + brisanje rezultata
run_calc = st.sidebar.button("🚀 Obračunaj zalihe", key="btn_run_calc")

if st.sidebar.button("🧹 Obriši rezultat", key="btn_clear_results"):
    st.session_state.pop("calc_results", None)
    st.sidebar.success("Rezultat obrisan ✅")
    st.rerun()

# ======================================================
# POKRETANJE OBRAČUNA (rezultat čuvamo u session_state)
# ======================================================
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

        st.session_state["calc_results"] = {
            "uporedba": uporedba,
            "df_fakture_posle": df_fakture_posle,
            "rules_used_df": rules_used_df,
            "kal_ekstremi": kal_ekstremi,
            "injekt_debug": injekt_debug,
            "audit_map": audit_map,
            "sus_bad": sus_bad,
        }

        st.success("✔ Obračun je završen.")

# ======================================================
# PRIKAZ UI: ako postoji rezultat u session_state, uvek ga prikazuj
# ======================================================
if "calc_results" not in st.session_state:
    st.info("⬅ Uploaduj fajlove i klikni na **'Obračunaj zalihe'**.")
    st.stop()

R = st.session_state["calc_results"]
uporedba = R["uporedba"]
df_fakture_posle = R["df_fakture_posle"]
rules_used_df = R["rules_used_df"]
kal_ekstremi = R["kal_ekstremi"]
injekt_debug = R["injekt_debug"]
audit_map = R["audit_map"]
sus_bad = R["sus_bad"]

# ======================================================
# ✅ "TABOVI" preko RADIO (ne skače na prvi tab)
# ======================================================
tab_labels = [
    "📌 Uporedba",
    "⚠ Ekstremi kalibracije",
    "📄 Fakture – obračun",
    "🔎 Mapiranje IZLAZ↔LAGER",
    "🧾 Audit (JM iste ≠ 1)",
    "🧮 Pravila (primenjena)",
    "📊 Grafikon",
    "🧪 Injekt debug",
    "📄 Izvoz po računu"
]

if "active_tab" not in st.session_state:
    st.session_state["active_tab"] = tab_labels[0]

active = st.radio(
    "Navigacija",
    tab_labels,
    index=tab_labels.index(st.session_state["active_tab"]),
    horizontal=True,
    key="radio_tabs"
)
st.session_state["active_tab"] = active

st.markdown("---")

# ======================================================
# TAB: Uporedba
# ======================================================
if active == "📌 Uporedba":
    st.subheader("Uporedba (ključne kolone + novi koef)")

    show = uporedba.copy()
    for c in ["Ukupna_potrošnja", "Stanje_na_lageru", "Razlika_pre", "Stanje_na_magacinu", "Finalna_potrošnja", "Koef_novi"]:
        if c in show.columns:
            show[c] = pd.to_numeric(show[c], errors="coerce")

    if "show_only_unmatched" not in st.session_state:
        st.session_state["show_only_unmatched"] = False

    if "Match_u_lageru" in show.columns:
        unmatched_mask = (show["Match_u_lageru"] == False)
    else:
        unmatched_mask = show["Stanje_na_lageru"].isna() | show["Ukupna_potrošnja"].isna()

    unmatched_count = int(unmatched_mask.sum())
    total_count = int(len(show))

    colA, colB, colC, colD = st.columns([2.0, 2.3, 2.2, 1.5])

    with colA:
        st.markdown(
            f"""
            <div style="display:inline-block;padding:6px 10px;border-radius:14px;
                        background:#eef2ff;border:1px solid #c7d2fe;font-weight:600;">
                📦 Ukupno: {total_count}
            </div>
            """,
            unsafe_allow_html=True
        )

    with colB:
        st.markdown(
            f"""
            <div style="display:inline-block;padding:6px 10px;border-radius:14px;
                        background:#fff7ed;border:1px solid #fed7aa;font-weight:600;">
                ⚠ Nematchovani: {unmatched_count}
            </div>
            """,
            unsafe_allow_html=True
        )

    with colC:
        st.session_state["show_only_unmatched"] = st.checkbox(
            "🔍 Samo nematchovani",
            value=st.session_state["show_only_unmatched"],
            key="chk_only_unmatched_uporedba"
        )

    with colD:
        if st.session_state["show_only_unmatched"]:
            if st.button("❌ Isključi", key="btn_disable_filter_uporedba"):
                st.session_state["show_only_unmatched"] = False
                st.rerun()

    filtered_df = show.copy()
    if st.session_state["show_only_unmatched"]:
        filtered_df = filtered_df[unmatched_mask].copy()
        if filtered_df.empty:
            st.warning("ℹ️ Nema nematchovanih stavki. Isključi filter da vidiš sve podatke.")

    cols = [
        "Materijal",
        "Ukupna_potrošnja",
        "Stanje_na_lageru",
        "Razlika_pre",
        "Stanje_na_magacinu",
        "Koef_novi",
        "Finalna_potrošnja",
    ]
    cols = [c for c in cols if c in filtered_df.columns]
    st.dataframe(filtered_df[cols], use_container_width=True)

# ======================================================
# TAB: Ekstremi
# ======================================================
elif active == "⚠ Ekstremi kalibracije":
    st.subheader("Ekstremi kalibracije (Koef_novi van očekivanog opsega)")
    st.dataframe(kal_ekstremi, use_container_width=True)

# ======================================================
# TAB: Fakture
# ======================================================
elif active == "📄 Fakture – obračun":
    st.subheader("Fakture – obračun količina za skidanje sa lagera")

    drop_cols = [
        "Napomena konverzije",
        "_racun_key", "_mat_key", "_jm_lager", "_uf", "_ul",
        "Koef_konverzije",
        "_koef_novi_mat",
    ]
    view_fakture = df_fakture_posle.drop(columns=drop_cols, errors="ignore")
    st.dataframe(view_fakture, use_container_width=True)

# ======================================================
# TAB: Mapiranje
# ======================================================
elif active == "🔎 Mapiranje IZLAZ↔LAGER":
    st.subheader("🔎 Provera mapiranja materijala i jedinica (IZLAZ ↔ LAGER)")

    if audit_map is None or audit_map.empty:
        st.info("Nema audit podataka.")
    else:
        fc1, fc2, fc3 = st.columns([1, 1, 2])
        with fc1:
            only_missing = st.checkbox("Samo nematchovani", value=False, key="chk_map_only_missing")
        with fc2:
            only_jm_mismatch = st.checkbox("Samo JM ne odgovara", value=False, key="chk_map_only_jm_mismatch")
        with fc3:
            q = st.text_input("Pretraga (materijal)", value="", key="txt_map_search")

        view = audit_map.copy()

        if "Match_u_lageru" in view.columns and only_missing:
            view = view[view["Match_u_lageru"] == False]

        if only_jm_mismatch and "JM_iste" in view.columns:
            view = view[(view["JM_iste"] == False) & (view["JM_fakt"].fillna("") != "")]

        if q.strip():
            qq = q.strip().lower()
            if "Materijal_lager" in view.columns:
                view = view[
                    view["Materijal"].astype(str).str.lower().str.contains(qq, na=False) |
                    view["Materijal_lager"].astype(str).str.lower().str.contains(qq, na=False)
                ]
            else:
                view = view[view["Materijal"].astype(str).str.lower().str.contains(qq, na=False)]

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
        st.subheader("🔁 Ručno mapiranje nematchovanih (sačuvaj, pa ponovo Obračunaj)")

        if "Match_u_lageru" in audit_map.columns:
            missing_unique = audit_map[audit_map["Match_u_lageru"] == False][["Materijal"]].drop_duplicates()
            missing_list = missing_unique["Materijal"].astype(str).tolist()
        else:
            missing_list = []

        if len(missing_list) == 0:
            st.success("Nema nematchovanih materijala ✅")
        else:
            local_lager_options = sorted(audit_map["Materijal_lager"].dropna().unique().tolist()) if "Materijal_lager" in audit_map.columns else []
            if not local_lager_options:
                local_lager_options = lager_options

            mm2 = st.session_state["manual_map"].copy()
            if mm2.empty:
                mm2 = pd.DataFrame(columns=["Materijal_izlaz", "Materijal_lager"])

            existing = set(mm2["Materijal_izlaz"].astype(str).tolist()) if "Materijal_izlaz" in mm2.columns else set()

            rows_to_add = []
            for m in missing_list:
                if m not in existing:
                    rows_to_add.append({"Materijal_izlaz": m, "Materijal_lager": ""})
            if rows_to_add:
                mm2 = pd.concat([mm2, pd.DataFrame(rows_to_add)], ignore_index=True)

            edited_mm2 = st.data_editor(
                mm2,
                use_container_width=True,
                num_rows="dynamic",
                column_config={
                    "Materijal_izlaz": st.column_config.TextColumn("Materijal_izlaz", disabled=True),
                    "Materijal_lager": st.column_config.SelectboxColumn("Materijal_lager", options=[""] + local_lager_options)
                },
                key="tab_map_manual_editor"
            )

            sm1, sm2 = st.columns(2)
            with sm1:
                if st.button("💾 Sačuvaj ručno mapiranje (iz ovog taba)", key="btn_save_mapping_tab"):
                    st.session_state["manual_map"] = edited_mm2
                    st.success("Sačuvano ✅ — sada klikni opet 'Obračunaj zalihe'.")
            with sm2:
                st.info("Tip: mapiraj na tačan naziv kolone iz lager fajla.")

# ======================================================
# TAB: Audit
# ======================================================
elif active == "🧾 Audit (JM iste ≠ 1)":
    st.subheader("🧾 Audit: Jedinice su iste, a konverzija nije 1.0 (ne bi smelo)")

    if sus_bad is None or sus_bad.empty:
        st.success("Nema sumnjivih slučajeva ✅")
    else:
        ac1, ac2 = st.columns([1, 2])
        with ac1:
            min_diff = st.number_input("Min odstupanje (|koef-1|)", value=0.001, step=0.001, key="audit_min_diff")
        with ac2:
            qq = st.text_input("Pretraga (materijal)", value="", key="audit_search")

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

# ======================================================
# TAB: Pravila (primenjena)
# ======================================================
elif active == "🧮 Pravila (primenjena)":
    st.subheader("Pravila konverzije (primenjena u ovom obračunu)")
    st.dataframe(rules_used_df, use_container_width=True)

# ======================================================
# TAB: Grafikon
# ======================================================
elif active == "📊 Grafikon":
    st.subheader("Poređenje stanja na lageru i ukupne potrošnje po materijalu")
    chart_df = uporedba.dropna(subset=["Materijal", "Stanje_na_lageru", "Ukupna_potrošnja"]).copy()
    if len(chart_df) > 0:
        max_items = st.slider(
            "Broj materijala na grafikonu (po razlici, apsolutna vrednost):",
            min_value=5,
            max_value=min(100, len(chart_df)),
            value=min(30, len(chart_df)),
            step=1,
            key="slider_max_items"
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

# ======================================================
# TAB: Injekt debug
# ======================================================
elif active == "🧪 Injekt debug":
    st.subheader("Injekt debug (paketi: 1.0–1.8x da se uklopi na 31.12)")
    if injekt_debug is None or injekt_debug.empty:
        st.info("Nema injekt računa ili nema magacin fajla.")
    else:
        st.dataframe(injekt_debug, use_container_width=True)

# ======================================================
# TAB: Izvoz po računu
# ======================================================
elif active == "📄 Izvoz po računu":
    st.subheader("📄 Izvoz dokumenata po računu")

    racuni = sorted(df_fakture_posle["Broj računa"].dropna().unique())
    selected_racun = st.selectbox("Izaberi broj računa", racuni, key="select_racun")

    c1, c2 = st.columns(2)

    with c1:
        word_buf = generate_word_for_racun(df_fakture_posle, selected_racun)
        st.download_button(
            "📄 Preuzmi Word (ovaj račun)",
            data=word_buf,
            file_name=f"racun_{selected_racun}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="btn_word_one"
        )

    with c2:
        xls_buf = generate_excel_for_racun(df_fakture_posle, selected_racun)
        st.download_button(
            "📊 Preuzmi Excel (ovaj račun)",
            data=xls_buf,
            file_name=f"racun_{selected_racun}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_excel_one"
        )

    st.markdown("---")
    st.subheader("📁 Grupni izvoz")

    c3, c4 = st.columns(2)

    with c3:
        zip_word = generate_word_zip_all_racuni(df_fakture_posle)
        st.download_button(
            "📁 Word za SVE račune (ZIP)",
            data=zip_word,
            file_name="svi_racuni_word.zip",
            mime="application/zip",
            key="btn_word_zip"
        )

    with c4:
        zip_excel = generate_excel_zip_all_racuni(df_fakture_posle)
        st.download_button(
            "📁 Excel za SVE račune (ZIP)",
            data=zip_excel,
            file_name="svi_racuni_excel.zip",
            mime="application/zip",
            key="btn_excel_zip"
        )

# ======================================================
# Download (Excel rezultat) – uvek na dnu
# ======================================================
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
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    key="btn_export_main_excel"
)
