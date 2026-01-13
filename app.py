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
stvarno_file = st.sidebar.file_uploader("Stvarno stanje magacina (POPIS)", type=["xlsx"])
st.sidebar.markdown("---")

# ============================
# 🛠 Editor koeficijenata (PRAVI)
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
    if st.button("↩ Reset"):
        st.session_state["rules_df"] = rules_to_df(material_rules)
        st.info("Vraćeno na podrazumevano.")

st.sidebar.markdown("---")
run_calc = st.sidebar.button("🚀 Obračunaj zalihe")

# --------------- Main area ---------------
if run_calc:
    if not lager_file or not fakture_file:
        st.error("❌ Morate da uploadujete **oba** Excel fajla (lager i fakture).")
    else:
        rules_df_for_run = st.session_state.get("rules_df", edited_rules_df)

        with st.spinner("Računam, molim sačekaj..."):
            uporedba, ekstremni, df_fakture_posle, rules_used_df = procesiraj_obracun(
                lager_file,
                fakture_file,
                edited_rules_df=rules_df_for_run,
                stvarno_file=stvarno_file
            )

        st.success("✔ Obračun je završen.")

        tab1, tab2, tab3, tab4 = st.tabs(
            ["📌 Koeficijenti i uporedba", "⚠ Ekstremi", "📄 Fakture – obračun", "🧮 Pravila konverzije"]
        )

        # ✅ Tab 1: samo kolone koje želiš
        with tab1:
            st.subheader("Uporedba (potrošnja, lager, popis)")

            cols = [
                "Materijal",
                "Ukupna_potrošnja",
                "Stanje_na_lageru",
                "Razlika_pre",
                "Stanje_na_magacinu",
                "Koef_popis",
                "Finalna_potrošnja",
            ]
            show_df = uporedba.copy()
            for c in cols:
                if c not in show_df.columns:
                    show_df[c] = pd.NA

            st.dataframe(show_df[cols], use_container_width=True)

        # ✅ Tab 2: ekstremi + iste kolone
        with tab2:
            st.subheader("Ekstremne vrednosti (pregled)")
            cols = [
                "Materijal",
                "Ukupna_potrošnja",
                "Stanje_na_lageru",
                "Razlika_pre",
                "Stanje_na_magacinu",
                "Koef_popis",
                "Finalna_potrošnja",
            ]
            if ekstremni is None or ekstremni.empty:
                st.info("Nema ekstremnih vrednosti po trenutno zadatim granicama.")
            else:
                tmp = ekstremni.copy()
                for c in cols:
                    if c not in tmp.columns:
                        tmp[c] = pd.NA
                st.dataframe(tmp[cols], use_container_width=True)

        with tab3:
            st.subheader("Fakture – obračun količina za skidanje sa lagera")
            st.dataframe(df_fakture_posle, use_container_width=True)

        with tab4:
            st.subheader("Pravila konverzije (primenjena u ovom obračunu)")
            st.dataframe(rules_used_df, use_container_width=True)

        # --- chart ---
        st.markdown("---")
        st.subheader("📊 Poređenje stanja na lageru i potrošnje po materijalu")

        use_final = st.checkbox("Koristi FINALNU potrošnju (po popisu)", value=True)

        chart_df = uporedba.dropna(subset=["Materijal", "Stanje_na_lageru", "Ukupna_potrošnja"]).copy()
        if len(chart_df) > 0:
            max_items = st.slider(
                "Broj materijala na grafikonu (po razlici, apsolutna vrednost):",
                min_value=5,
                max_value=min(100, len(chart_df)),
                value=min(30, len(chart_df)),
                step=1
            )

            y_col = "Ukupna_potrošnja"
            if use_final and "Finalna_potrošnja" in chart_df.columns and chart_df["Finalna_potrošnja"].notna().any():
                y_col = "Finalna_potrošnja"

            chart_df["Abs_diff"] = (chart_df["Stanje_na_lageru"] - chart_df[y_col]).abs()
            chart_df = chart_df.sort_values("Abs_diff", ascending=False).head(max_items)

            fig = px.bar(
                chart_df,
                x="Materijal",
                y=["Stanje_na_lageru", y_col],
                barmode="group",
                title=f"Stanje na lageru vs {y_col}",
            )
            fig.update_layout(xaxis_tickangle=45, height=650, legend_title_text="Količina")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Nema dovoljno podataka za grafikon.")

        # --- download ---
        st.subheader("📎 Preuzimanje Excel rezultata")
        buffer = export_to_excel(uporedba, ekstremni, df_fakture_posle)
        st.download_button(
            label="💾 Preuzmi Excel rezultat",
            data=buffer,
            file_name="obracun_zaliha.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("⬅ Uploaduj Excel fajlove u levom meniju i klikni na **'Obračunaj zalihe'**.")
