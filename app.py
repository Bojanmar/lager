# app.py
import streamlit as st
import pandas as pd
import plotly.express as px

from obracun import (
    procesiraj_obracun,
    rules_to_df,
    RULE_TYPES,
    export_to_excel,
    material_rules,
    build_calibration_table_and_apply,
)

st.set_page_config(page_title="Obračun zaliha", layout="wide")
st.title("📦 Obračun zaliha – Laser Lux")

# ---------------- Sidebar ----------------
st.sidebar.header("Ulazni podaci")
lager_file = st.sidebar.file_uploader("Lager Excel (ULAZ)", type=["xlsx"])
fakture_file = st.sidebar.file_uploader("Fakture Excel (IZLAZ)", type=["xlsx"])
magacin_file = st.sidebar.file_uploader("Stvarno stanje magacina (POPIS)", type=["xlsx"])
st.sidebar.markdown("---")

# ============================
# 🛠 Editor koeficijenata (pravila)
# ============================
st.sidebar.subheader("🧮 Pravila (koeficijenti)")

if "rules_df" not in st.session_state:
    st.session_state["rules_df"] = rules_to_df(material_rules)

if "rules_df_base" not in st.session_state:
    # “inicijalna” kopija za ekstrem uporedbe
    st.session_state["rules_df_base"] = st.session_state["rules_df"].copy()

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
    if st.button("↩ Reset"):
        st.session_state["rules_df"] = rules_to_df(material_rules)
        st.session_state["rules_df_base"] = st.session_state["rules_df"].copy()
        st.info("Vraćeno na podrazumevano.")

st.sidebar.markdown("---")
run_calc = st.sidebar.button("🚀 Obračunaj zalihe")

# --------------- Main area ---------------
if run_calc:
    if not lager_file or not fakture_file:
        st.error("❌ Morate da uploadujete **lager** i **fakture**.")
    else:
        rules_df_for_run = st.session_state.get("rules_df", edited_rules_df)

        with st.spinner("Računam, molim sačekaj..."):
            uporedba, ekstremni, df_fakture_posle, rules_used_df = procesiraj_obracun(
                lager_file,
                fakture_file,
                magacin_file=magacin_file,
                edited_rules_df=rules_df_for_run
            )

        st.success("✔ Obračun je završen.")

        # ---- Kalibracija (Opcija B) – napravi tabelu, ali NE primenjuj automatski
        extreme_low = st.sidebar.number_input("Kalibracija ekstrem: min ratio", value=0.5, step=0.1)
        extreme_high = st.sidebar.number_input("Kalibracija ekstrem: max ratio", value=2.0, step=0.1)

        rules_df_after_calib, calib_df, calib_extremes_df = build_calibration_table_and_apply(
            st.session_state["rules_df"],
            uporedba,
            extreme_low=float(extreme_low),
            extreme_high=float(extreme_high)
        )

        tab1, tab2, tab3, tab4, tab5 = st.tabs(
            ["📌 Uporedba", "⚠ Ekstremi", "📄 Fakture – obračun", "🧮 Pravila (run)", "🧩 Kalibracija"]
        )

        with tab1:
            st.subheader("Uporedba (ključne kolone + popis)")
            cols = [
                "Materijal",
                "Ukupna_potrošnja",
                "Stanje_na_lageru",
                "Razlika_pre",
                "Stanje_na_magacinu",
                "Koef_popis",
                "Finalna_potrošnja",
            ]
            show = uporedba.copy()
            for c in cols:
                if c not in show.columns:
                    show[c] = None
            st.dataframe(show[cols], use_container_width=True)

        with tab2:
            st.subheader("Ekstremne vrednosti (stari koeficijent lager/potrošnja)")
            cols_to_show = ["Materijal", "Ukupna_potrošnja", "Stanje_na_lageru", "Razlika_pre", "Koeficijent", "Napomena_coef"]
            df_show = ekstremni[cols_to_show].copy() if not ekstremni.empty else ekstremni
            st.dataframe(df_show, use_container_width=True)

        with tab3:
            st.subheader("Fakture – obračun količina za skidanje sa lagera")
            st.dataframe(df_fakture_posle, use_container_width=True)

        with tab4:
            st.subheader("Pravila konverzije (primenjena u ovom obračunu)")
            st.dataframe(rules_used_df, use_container_width=True)

        with tab5:
            st.subheader("Kalibracija (Opcija B) — predlog izmena pravila prema popisu")

            if calib_df is None or calib_df.empty:
                st.info("Nema dovoljno podataka za kalibraciju (treba Ukupna_potrošnja, Stanje_na_lageru i Stanje_na_magacinu).")
            else:
                st.write("Ovo je tabela šta bi se promenilo u pravilima (menja se samo primarno pravilo po materijalu).")
                st.dataframe(calib_df, use_container_width=True)

                st.markdown("### ⚠ Ekstremi kalibracije (prevelika promena faktora)")
                if calib_extremes_df is None or calib_extremes_df.empty:
                    st.info("Nema ekstremnih promena po zadatim granicama.")
                else:
                    st.dataframe(calib_extremes_df, use_container_width=True)

                colA, colB = st.columns([1, 2])
                with colA:
                    if st.button("✅ Primeni kalibraciju u pravila"):
                        st.session_state["rules_df"] = rules_df_after_calib
                        st.success("Pravila su ažurirana kalibracijom. Pokreni obračun ponovo.")
                        st.rerun()
                with colB:
                    st.caption("Ovo menja faktore u tabeli pravila (Opcija B). Sledeći obračun koristi nove faktore.")

        # --- Side-by-side bar chart ---
        st.markdown("---")
        st.subheader("📊 Poređenje stanja na lageru i ukupne potrošnje po materijalu")

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

        # --- download ---
        st.subheader("📎 Preuzimanje Excel rezultata")
        buffer = export_to_excel(
            uporedba,
            ekstremni,
            df_fakture_posle,
            calibration_df=calib_df,
            ekstremi_kalibracije=calib_extremes_df,
            rules_used_df=rules_used_df
        )
        st.download_button(
            label="💾 Preuzmi Excel rezultat",
            data=buffer,
            file_name="obracun_zaliha.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("⬅ Uploaduj fajlove u levom meniju i klikni na **'Obračunaj zalihe'**.")
