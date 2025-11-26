import streamlit as st
import pandas as pd
import plotly.express as px


from obracun import procesiraj_obracun, get_rules_overview, rules_editor_ui, export_to_excel


st.set_page_config(
    page_title="Obračun zaliha",
    layout="wide"
)

st.title("📦 Obračun zaliha – Laser Lux")

# ---------------- Sidebar ----------------
st.sidebar.header("Ulazni podaci")

lager_file = st.sidebar.file_uploader("Lager Excel (ULAZ)", type=["xlsx"])
fakture_file = st.sidebar.file_uploader("Fakture Excel (IZLAZ)", type=["xlsx"])

st.sidebar.markdown("---")

# ============================
# 🛠 Editor koeficijenata
# ============================
from obracun import rules_editor_ui






st.sidebar.markdown("---")
run_calc = st.sidebar.button("🚀 Obračunaj zalihe")

# --------------- Main area ---------------
if run_calc:
    if not lager_file or not fakture_file:
        st.error("❌ Morate da uploadujete **oba** Excel fajla (lager i fakture).")
    else:
        with st.spinner("Računam, molim sačekaj..."):
            uporedba, ekstremni, df_fakture_posle = procesiraj_obracun(
                lager_file,
                fakture_file
            )

        st.success("✔ Obračun je završen.")

        # --- Tabele ---
        tab1, tab2, tab3, tab4 = st.tabs(
            ["📌 Koeficijenti i uporedba",
            "⚠ Ekstremi",
            "📄 Fakture – obračun",
            "🧮 Pravila konverzije"]
        )


        with tab1:
            st.subheader("Koeficijenti i uporedba (lager vs potrošnja)")
            st.dataframe(uporedba, use_container_width=True)

        with tab2:
            st.subheader("Ekstremne vrednosti (pregled potrošnje)")
            cols_to_show = ["Materijal", "Ukupna_potrošnja", "Stanje_na_lageru", "Razlika_pre"]
            df_show = ekstremni[cols_to_show].copy()
            st.dataframe(df_show, use_container_width=True)



        with tab3:
            st.subheader("Fakture – obračun količina za skidanje sa lagera")
            st.dataframe(df_fakture_posle, use_container_width=True)
        from obracun import rules_editor_ui

        with tab4:
            st.subheader("Pravila konverzije (koeficijenti)")
            st.write("Ovde možeš direktno da menjaš pravila kao u sidebaru.")
            rules_editor_ui()



        # --- Side-by-side bar chart ---
        st.markdown("---")
        st.subheader("📊 Poređenje stanja na lageru i ukupne potrošnje po materijalu")

        chart_df = uporedba.dropna(
            subset=["Materijal", "Stanje_na_lageru", "Ukupna_potrošnja"]
        )

        # po želji možeš da ograničiš broj materijala u grafiku
        max_items = st.slider(
            "Broj materijala na grafikonu (po razlici, apsolutna vrednost):",
            min_value=5,
            max_value=min(100, len(chart_df)),
            value=min(30, len(chart_df)),
            step=1
        )

        chart_df["Abs_diff"] = (chart_df["Stanje_na_lageru"] - chart_df["Ukupna_potrošnja"]).abs()
        chart_df = chart_df.sort_values("Abs_diff", ascending=False).head(max_items)

        if not chart_df.empty:
            fig = px.bar(
                chart_df,
                x="Materijal",
                y=["Stanje_na_lageru", "Ukupna_potrošnja"],
                barmode="group",
                title="Stanje na lageru vs ukupna potrošnja",
            )
            fig.update_layout(
                xaxis_tickangle=45,
                height=650,
                legend_title_text="Količina"
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Nema dovoljno podataka za grafikon.")
            # ============================
        # 📥 DOWNLOAD EXCEL
        # ============================
        import io
        from obracun import export_to_excel

        buffer = export_to_excel(uporedba, ekstremni, df_fakture_posle)

        st.subheader("📎 Preuzimanje Excel rezultata")
        st.download_button(
            label="💾 Preuzmi Excel rezultat",
            data=buffer,
            file_name="obracun_zaliha.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("⬅ Uploaduj Excel fajlove u levom meniju i klikni na **'Obračunaj zalihe'**.")
