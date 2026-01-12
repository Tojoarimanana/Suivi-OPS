import streamlit as st
import pandas as pd
import geopandas as gpd
import folium
from streamlit_folium import st_folium
import os
import tempfile
import zipfile
import io
from docx import Document
from fpdf import FPDF

# --- Fonctions d'export ---

def export_to_excel(df_dict):
    with io.BytesIO() as output:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, df in df_dict.items():
                df.to_excel(writer, sheet_name=sheet_name[:31])
            writer.save()
        data = output.getvalue()
    return data

def export_to_word(df_dict):
    doc = Document()
    for titre, df in df_dict.items():
        doc.add_heading(titre, level=2)
        if df.empty:
            doc.add_paragraph("Aucune donnée sélectionnée.")
        else:
            table = doc.add_table(rows=1, cols=len(df.columns))
            hdr_cells = table.rows[0].cells
            for i, col_name in enumerate(df.columns):
                hdr_cells[i].text = str(col_name)
            for _, row in df.iterrows():
                row_cells = table.add_row().cells
                for i, val in enumerate(row):
                    row_cells[i].text = str(val)
        doc.add_paragraph()
    f = io.BytesIO()
    doc.save(f)
    f.seek(0)
    return f.read()

def export_to_pdf(df_dict):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    for titre, df in df_dict.items():
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(0, 10, titre, ln=True)
        pdf.set_font("Arial", size=10)
        if df.empty:
            pdf.cell(0, 10, "Aucune donnée sélectionnée.", ln=True)
        else:
            col_width = pdf.w / (len(df.columns) + 1)
            row_height = pdf.font_size * 1.5
            # Header
            for col_name in df.columns:
                pdf.cell(col_width, row_height, str(col_name), border=1)
            pdf.ln(row_height)
            # Rows
            for _, row in df.iterrows():
                for val in row:
                    pdf.cell(col_width, row_height, str(val), border=1)
                pdf.ln(row_height)
        pdf.ln(10)
    return pdf.output(dest='S').encode('latin1')

# --- Fonctions principales ---

def load_data(file_path):
    return pd.read_excel(file_path)

def afficher_carte(df, gdf):
    noms_filtrés = df['Nom'].unique()
    gdf_filtered = gdf[gdf['Nom'].isin(noms_filtrés)]

    if gdf_filtered.empty:
        st.warning("Aucun bloc géographique ne correspond au filtre sélectionné.")
        return

    centroid = gdf_filtered.geometry.centroid
    map_center = [centroid.y.mean(), centroid.x.mean()]
    m = folium.Map(location=map_center, zoom_start=6)

    for _, row in gdf_filtered.iterrows():
        info = df[df['Nom'] == row['Nom']].iloc[0]

        popup_content = f"""
        <b>Bloc :</b> {row.get('Nom', 'N/A')}<br>
        <b>Compagnie :</b> {info.get('Compagnie', 'N/A')}<br>
        <b>Phase :</b> {info.get('Phases_actuelle', 'N/A')}<br>
        <b>Commentaires :</b> {info.get('Commentaires1', 'N/A')}
        """
        folium.GeoJson(
            row['geometry'],
            tooltip=row.get('Nom', 'N/A'),
            popup=folium.Popup(popup_content, max_width=300)
        ).add_to(m)

    st.subheader("Carte des blocs pétroliers")
    st_folium(m, width=800, height=600)

def afficher_table(df, colonnes, titre):
    st.subheader(titre)
    st.dataframe(df[colonnes], use_container_width=True)

# --- App principale ---

def main():
    st.set_page_config(page_title="Suivi des Compagnies Pétrolières", layout="wide")
    st.title('Tableau de Bord - Suivi des Compagnies Pétrolières')

    file_path = st.file_uploader("Télécharger votre fichier Excel", type=["xlsx"])
    shapefile_zip = st.file_uploader("Téléchargez un shapefile zippé (.zip)", type=["zip"])

    if file_path is not None and shapefile_zip is not None:
        df = load_data(file_path)

        compagnies = df['Compagnie'].unique()
        selected_compagnie = st.sidebar.selectbox('Filtrer par compagnie', ['Tous'] + list(compagnies))
        if selected_compagnie != 'Tous':
            df = df[df['Compagnie'] == selected_compagnie]

        with tempfile.TemporaryDirectory() as tmpdir:
            zip_path = os.path.join(tmpdir, "shapefile.zip")
            with open(zip_path, "wb") as f:
                f.write(shapefile_zip.getvalue())
            with zipfile.ZipFile(zip_path, "r") as zip_ref:
                zip_ref.extractall(tmpdir)

            shp_files = [os.path.join(root, file) for root, dirs, files in os.walk(tmpdir) for file in files if file.endswith(".shp")]

            if shp_files:
                shp_path = shp_files[0]
                gdf = gpd.read_file(shp_path)

                onglets = st.tabs([
                    "Carte",
                    "Compagnie",
                    "Situation Actuelle",
                    "Termes Commerciaux",
                    "Obligations Contractuelles",
                    "MCM/TCM",
                    "Obligations Financières",
                    "Avenants",
                    "Rapport"
                ])

                # Colonnes par onglet
                groupes = {
                    "Compagnie": ['Compagnie', 'Nom', 'ID', 'Coordonée_X', 'Coordonée_Y', 'Date_de_signature_de_contrats', 'Date_d_entrée_en_vigeur'],
                    "Situation Actuelle": ['Phases_actuelle', 'Date_de_debut_de_la_phase', 'Date_de_la_fin_de_la_phase', 'Situation_et_Activités_en_cours', 'Travaux_déjà_réalisés', 'Commentaires1'],
                    "Termes Commerciaux": ['Cost_Recovery_Limit_(%)', 'Overhead_(%)', 'Frais_d_Administration_(M_$)', 'Frais_de_Formation_(M_$)', 'Bonus_de_Production_(M_$)', 'Partage_de_Production_Pétrole_(Part_du_Gouvernement)', 'Partage_de_Production_Gaz_(Part_du_Gouvernement)'],
                    "Obligations Contractuelles": ['Obligation_de_Travaux', 'Obligation_de_Rendu_(%)', 'Obligation_de_Banque_Garantie_(M_$)', 'Travaux_réalisées', 'Rendu_réalisé_(%)', 'Banque_Garantie_déposées_(M_$)', 'Commentaires2'],
                    "MCM/TCM": ['Date_du_dernier_MCM', 'Lieu', 'Motifs', 'Résolution', 'PTA_&_Budget', 'Réalisation_budgetaire', 'Commentaires3'],
                    "Obligations Financières": ['Frais_de_Formation', 'Dernier_Paiement_de_frais_de_Formation', 'Frais_d_Administration', 'Dernier_Paiement_de_frais_d_Administration', 'Garantie_Bancaire', 'Dernier_Dépôt', 'Observations'],
                    "Avenants": ['Dernier_Avenant', 'Date_de_Signature', 'Motifs_Avenant', 'Statut']
                }

                # Pour stocker la sélection des colonnes dans chaque onglet
                colonnes_selectionnees = {}

                with onglets[0]:
                    afficher_carte(df, gdf)

                for i, (titre, colonnes) in enumerate(list(groupes.items()), start=1):
                    with onglets[i]:
                        st.subheader(f"Filtres - {titre}")
                        selection = st.multiselect(f"Sélectionner colonnes à afficher dans {titre} :", options=colonnes, default=colonnes, key=f"multi_{titre}")
                        colonnes_selectionnees[titre] = selection

                        if selection:
                            afficher_table(df, selection, titre)
                        else:
                            st.info("Veuillez sélectionner au moins une colonne.")

                with onglets[-1]:
                    st.subheader("Rapport récapitulatif dynamique")

                    if not colonnes_selectionnees:
                        st.info("Sélectionnez d'abord les colonnes dans les onglets pour générer le rapport.")
                    else:
                        for titre, cols_sel in colonnes_selectionnees.items():
                            st.markdown(f"### {titre}")
                            if cols_sel:
                                st.dataframe(df[cols_sel], use_container_width=True)
                            else:
                                st.info("Aucune colonne sélectionnée dans cet onglet.")

                        format_export = st.selectbox("Choisissez le format d'export :", ["Excel", "Word", "PDF"], key="format_export")

                        if st.button("Télécharger le rapport"):
                            if not any(colonnes_selectionnees.values()):
                                st.warning("Aucune donnée sélectionnée à exporter.")
                            else:
                                dfs_export = {titre: df[cols_sel] for titre, cols_sel in colonnes_selectionnees.items() if cols_sel}

                                if format_export == "Excel":
                                    data = export_to_excel(dfs_export)
                                    st.download_button(
                                        label="Télécharger le rapport Excel",
                                        data=data,
                                        file_name="rapport_compagnies.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                                elif format_export == "Word":
                                    data = export_to_word(dfs_export)
                                    st.download_button(
                                        label="Télécharger le rapport Word",
                                        data=data,
                                        file_name="rapport_compagnies.docx",
                                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                    )
                                else:
                                    data = export_to_pdf(dfs_export)
                                    st.download_button(
                                        label="Télécharger le rapport PDF",
                                        data=data,
                                        file_name="rapport_compagnies.pdf",
                                        mime="application/pdf"
                                    )
            else:
                st.warning("Aucun fichier .shp trouvé dans l'archive.")

if __name__ == "__main__":
    main()


































































































































































































































































































































































































































































































































































































































































































































































































































































































































































































































































































