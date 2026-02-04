import streamlit as st
import pandas as pd
import geopandas as gpd
import folium
from streamlit_folium import st_folium
import os
import tempfile
import zipfile
import io
import re
from docx import Document

# Fonction pour formater les dates en français
def format_date_fr(date):
    if pd.isna(date) or date is None:
        return ""
    if isinstance(date, str):
        date = pd.to_datetime(date)
    months = {
        1: 'janvier', 2: 'février', 3: 'mars', 4: 'avril',
        5: 'mai', 6: 'juin', 7: 'juillet', 8: 'août',
        9: 'septembre', 10: 'octobre', 11: 'novembre', 12: 'décembre'
    }
    return f"{date.day} {months[date.month]} {date.year}"

# --- Fonctions d'export ---

def clean_sheet_name(name):
    """Nettoie le nom des onglets Excel"""
    name = re.sub(r'[\\/*?:\[\]]', '_', str(name))
    return name[:31] if name else "Sheet"

def export_to_excel(df_dict):
    with io.BytesIO() as output:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, df in df_dict.items():
                safe_name = clean_sheet_name(sheet_name)
                df.to_excel(writer, sheet_name=safe_name, index=False)
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

# --- Fonctions principales ---

def load_data(file_path):
    df = pd.read_excel(file_path)
    date_columns = [
        'Date_de_signature_de_contrats', 'Date_d_entrée_en_vigeur',
        'Date_de_debut_de_la_phase', 'Date_de_la_fin_de_la_phase',
        'Date_du_dernier_MCM', 'Dernier_Paiement_de_frais_de_Formation',
        'Dernier_Paiement_de_frais_d_Administration', 'Dernier_Dépôt',
        'Date_de_Signature'
    ]
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    return df

def load_shapefile(shapefile_zip):
    with tempfile.TemporaryDirectory() as tmpdir:
        zip_path = os.path.join(tmpdir, "shapefile.zip")
        with open(zip_path, "wb") as f:
            f.write(shapefile_zip.getvalue())
        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            zip_ref.extractall(tmpdir)

        shp_files = [os.path.join(root, file) for root, dirs, files in os.walk(tmpdir)
                     for file in files if file.endswith(".shp")]

        if shp_files:
            return gpd.read_file(shp_files[0])
        else:
            raise ValueError("Aucun fichier .shp trouvé dans l'archive.")

def format_df_for_display(df):
    df_formatted = df.copy()
    date_columns = [
        'Date_de_signature_de_contrats', 'Date_d_entrée_en_vigeur',
        'Date_de_debut_de_la_phase', 'Date_de_la_fin_de_la_phase',
        'Date_du_dernier_MCM', 'Dernier_Paiement_de_frais_de_Formation',
        'Dernier_Paiement_de_frais_d_Administration', 'Dernier_Dépôt',
        'Date_de_Signature'
    ]
    for col in date_columns:
        if col in df_formatted.columns:
            df_formatted[col] = df_formatted[col].apply(format_date_fr)
    return df_formatted

def format_df_for_export(df):
    return format_df_for_display(df)  # même formatage pour l'export

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

        date_fields = [
            'Date_de_signature_de_contrats', 'Date_d_entrée_en_vigeur',
            'Date_de_debut_de_la_phase', 'Date_de_la_fin_de_la_phase'
        ]
        formatted_dates = {}
        for field in date_fields:
            if field in info and pd.notna(info[field]):
                formatted_dates[field] = format_date_fr(info[field])
            else:
                formatted_dates[field] = 'N/A'

        popup_content = f"""
        <b>Bloc :</b> {row.get('Nom', 'N/A')}<br>
        <b>Compagnie :</b> {info.get('Compagnie', 'N/A')}<br>
        <b>Phase :</b> {info.get('Phases_actuelle', 'N/A')}<br>
        <b>Date de signature de contrats :</b> {formatted_dates.get('Date_de_signature_de_contrats', 'N/A')}<br>
        <b>Date d'entrée en vigueur :</b> {formatted_dates.get('Date_d_entrée_en_vigeur', 'N/A')}<br>
        <b>Date de début de la phase :</b> {formatted_dates.get('Date_de_debut_de_la_phase', 'N/A')}<br>
        <b>Date de fin de la phase :</b> {formatted_dates.get('Date_de_la_fin_de_la_phase', 'N/A')}<br>
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
    df_formatted = format_df_for_display(df)
    st.dataframe(df_formatted[colonnes], use_container_width=True)

# ────────────────────────────────────────────────
#                   APPLICATION
# ────────────────────────────────────────────────

def main():
    st.set_page_config(page_title="Suivi des Compagnies Pétrolières", layout="wide")

    # Fond d'écran (optionnel - remplacez l'URL si besoin)
    page_bg_img = '''
    <style>
    body {
    background-image: url("https://www.exemple.com/mon_image.jpg");
    background-size: cover;
    background-repeat: no-repeat;
    background-attachment: fixed;
    }
    </style>
    '''
    st.markdown(page_bg_img, unsafe_allow_html=True)

    st.title('OMNIS - Suivi des Compagnies Pétrolières')

    # Chargement des données
    if 'raw_df' not in st.session_state or 'gdf' not in st.session_state:
        col1, col2 = st.columns(2)
        with col1:
            file_path = st.file_uploader("Télécharger votre fichier Excel", type=["xlsx"], key='excel_uploader')
        with col2:
            shapefile_zip = st.file_uploader("Téléchargez un shapefile zippé (.zip)", type=["zip"], key='shapefile_uploader')

        if file_path is not None and 'raw_df' not in st.session_state:
            try:
                st.session_state.raw_df = load_data(file_path)
                st.rerun()
            except Exception as e:
                st.error(f"Erreur lors du chargement Excel : {e}")

        if shapefile_zip is not None and 'gdf' not in st.session_state:
            try:
                st.session_state.gdf = load_shapefile(shapefile_zip)
                st.rerun()
            except Exception as e:
                st.error(f"Erreur shapefile : {e}")
    else:
        st.success("Fichiers chargés avec succès !")

        raw_df = st.session_state.raw_df
        gdf = st.session_state.gdf

        # Filtre compagnie
        compagnies = sorted(raw_df['Compagnie'].dropna().unique())
        selected_compagnie = st.sidebar.selectbox(
            'Filtrer par compagnie',
            ['Tous'] + list(compagnies)
        )

        df = raw_df.copy()
        if selected_compagnie != 'Tous':
            df = df[df['Compagnie'] == selected_compagnie]

        # Onglets
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

        groupes = {
            "Compagnie": ['Compagnie', 'Nom', 'Bloc', 'Coordonée_X', 'Coordonée_Y',
                          'Date_de_signature_de_contrats', 'Date_d_entrée_en_vigeur'],
            "Situation Actuelle": ['Phases_actuelle', 'Date_de_debut_de_la_phase',
                                   'Date_de_la_fin_de_la_phase', 'Situation_et_Activités_en_cours',
                                   'Travaux_déjà_réalisés', 'Commentaires1'],
            "Termes Commerciaux": ['Cost_Recovery_Limit_(%)', 'Overhead_(%)',
                                   'Frais_d_Administration_(M_$)', 'Frais_de_Formation_(M_$)',
                                   'Bonus_de_Production_(M_$)',
                                   'Partage_de_Production_Pétrole_(Part_du_Gouvernement)',
                                   'Partage_de_Production_Gaz_(Part_du_Gouvernement)'],
            "Obligations Contractuelles": ['Obligation_de_Travaux', 'Obligation_de_Rendu_(%)',
                                           'Obligation_de_Banque_Garantie_(M_$)', 'Travaux_réalisées',
                                           'Rendu_réalisé_(%)', 'Banque_Garantie_déposées_(M_$)',
                                           'Commentaires2'],
            "MCM/TCM": ['Date_du_dernier_MCM', 'Lieu', 'Motifs', 'Résolution',
                        'PTA_&_Budget', 'Réalisation_budgetaire', 'Commentaires3'],
            "Obligations Financières": ['Frais_de_Formation', 'Dernier_Paiement_de_frais_de_Formation',
                                        'Frais_d_Administration', 'Dernier_Paiement_de_frais_d_Administration',
                                        'Garantie_Bancaire', 'Dernier_Dépôt', 'Observations'],
            "Avenants": ['Dernier_Avenant', 'Date_de_Signature', 'Motifs_Avenant', 'Statut']
        }

        colonnes_selectionnees = {}

        with onglets[0]:
            afficher_carte(df, gdf)

        for i, (titre, colonnes) in enumerate(list(groupes.items()), start=1):
            with onglets[i]:
                st.subheader(f"Filtres - {titre}")
                options = colonnes[:]
                if 'Bloc' not in options and 'Nom' not in options:
                    options.insert(0, 'Nom' if 'Nom' in df.columns else 'Bloc')

                selection = st.multiselect(
                    f"Colonnes à afficher dans {titre} :",
                    options=options,
                    default=options,
                    key=f"multi_{titre}"
                )

                if 'Nom' in selection or 'Bloc' in selection:
                    selection = ['Nom' if 'Nom' in df.columns else 'Bloc'] + \
                                [c for c in selection if c not in ['Nom', 'Bloc']]

                colonnes_selectionnees[titre] = selection

                if selection:
                    afficher_table(df, selection, titre)
                else:
                    st.info("Sélectionnez au moins une colonne.")

        with onglets[-1]:
            st.subheader("Rapport récapitulatif dynamique")

            if not any(colonnes_selectionnees.values()):
                st.info("Sélectionnez d'abord des colonnes dans les onglets précédents.")
            else:
                for titre, cols_sel in colonnes_selectionnees.items():
                    if cols_sel:
                        st.markdown(f"### {titre}")
                        st.dataframe(format_df_for_display(df)[cols_sel], use_container_width=True)

                format_export = st.selectbox("Format d'export :", ["Excel", "Word"])

                if st.button("Générer et télécharger le rapport"):
                    if not any(colonnes_selectionnees.values()):
                        st.warning("Aucune colonne sélectionnée.")
                    else:
                        dfs_export = {
                            titre: format_df_for_export(df[cols_sel])
                            for titre, cols_sel in colonnes_selectionnees.items()
                            if cols_sel
                        }

                        if format_export == "Excel":
                            data = export_to_excel(dfs_export)
                            st.download_button(
                                "Télécharger rapport.xlsx",
                                data,
                                file_name="rapport_compagnies_pétrolières.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            data = export_to_word(dfs_export)
                            st.download_button(
                                "Télécharger rapport.docx",
                                data,
                                file_name="rapport_compagnies_pétrolières.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )

    # Footer
    st.markdown(
        """
        <style>
        .footer {
            position: fixed;
            left: 0;
            bottom: 0;
            width: 100%;
            background-color: #f1f1f1;
            color: #333;
            text-align: center;
            padding: 8px;
            font-size: 14px;
            border-top: 1px solid #ddd;
        }
        </style>
        <div class="footer">
            Conçu par <b>RANAIVOSOA Tojoarimanana Hiratriniala</b> — Tél : +261 33 51 880 19
        </div>
        """,
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
