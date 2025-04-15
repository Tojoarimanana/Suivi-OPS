import streamlit as st
import pandas as pd
import geopandas as gpd
import folium
from streamlit_folium import st_folium
import os
import tempfile
import zipfile

# Fonction pour charger les données Excel
def load_data(file_path):
    return pd.read_excel(file_path)

# Streamlit App
def main():
    st.set_page_config(page_title="Suivi des Compagnies Pétrolières", layout="wide")
    st.title('Tableau de Bord - Suivi des Compagnies Pétrolières')

    file_path = st.file_uploader("Télécharger votre fichier Excel", type=["xlsx"])
    shapefile_zip = st.file_uploader("Téléchargez un shapefile zippé (.zip)", type=["zip"])

    if file_path is not None and shapefile_zip is not None:
        df = load_data(file_path)

        compagnies = df['Compagnie'].unique()
        selected_compagnie = st.selectbox('Sélectionner une compagnie', ['Tous'] + list(compagnies))

        if selected_compagnie != 'Tous':
            df = df[df['Compagnie'] == selected_compagnie]

        # Traitement du shapefile à partir de l'archive zip
        with tempfile.TemporaryDirectory() as tmpdir:
            zip_path = os.path.join(tmpdir, "shapefile.zip")
            with open(zip_path, "wb") as f:
                f.write(shapefile_zip.getvalue())

            with zipfile.ZipFile(zip_path, "r") as zip_ref:
                zip_ref.extractall(tmpdir)

            shp_files = [f for f in os.listdir(tmpdir) if f.endswith(".shp")]

            if shp_files:
                shp_path = os.path.join(tmpdir, shp_files[0])
                gdf = gpd.read_file(shp_path)

                gdf_merged = gdf.merge(df, left_on="Nom", right_on="Nom", how="left")

                # Disposition en colonnes : tableau à gauche, carte à droite
                col1, col2 = st.columns([2, 2])

                with col1:
                    st.sidebar.header('Sélectionner les colonnes à afficher')

                    group_1 = ['Compagnie', 'Nom','ID', 'Coordonée_X', 'Coordonée_Y', 'Date_de_signature_de_contrats', 'Date_d_entrée_en_vigeur']
                    group_2 = ['Phases_actuelle', 'Date_de_debut_de_la_phase', 'Date_de_la_fin_de_la_phase', 'Situation_et_Activités_en_cours', 'Travaux_déjà_réalisés', 'Commentaires1']
                    group_3 = ['Cost_Recovery_Limit_(%)', 'Overhead_(%)', 'Frais_d_Administration_(M_$)', 'Frais_de_Formation_(M_$)', 'Bonus_de_Production_(M_$)', 'Partage_de_Production_Pétrole_(Part_du_Gouvernement)', 'Partage_de_Production_Gaz_(Part_du_Gouvernement)']
                    group_4 = ['Obligation_de_Travaux', 'Obligation_de_Rendu_(%)', 'Obligation_de_Banque_Garantie_(M_$)', 'Travaux_réalisées', 'Rendu_réalisé_(%)', 'Banque_Garantie_déposées_(M_$)', 'Commentaires2']
                    group_5 = ['Date_du_dernier_MCM', 'Lieu', 'Motifs', 'Résolution', 'PTA_&_Budget', 'Réalisation_budgetaire', 'Commentaires3']
                    group_6 = ['Frais_de_Formation', 'Dernier_Paiement_de_frais_de_Formation', 'Frais_d_Administration', 'Dernier_Paiement_de_frais_d_Administration', 'Garantie_Bancaire', 'Dernier_Dépôt', 'Observations']
                    group_7 = ['Dernier_Avenant', 'Date_de_Signature', 'Motifs_Avenant', 'Statut']

                    selected_columns = []

                    with st.sidebar.expander('A propos de la compagnie'):
                        selected_columns += st.multiselect('Sélectionnez les colonnes à afficher', group_1, default=group_1)

                    with st.sidebar.expander('Situation Actuelle'):
                        selected_columns += st.multiselect('Sélectionnez les colonnes à afficher', group_2, default=group_2)

                    with st.sidebar.expander('Termes Commerciaux dans le contrat'):
                        selected_columns += st.multiselect('Sélectionnez les colonnes à afficher', group_3, default=group_3)

                    with st.sidebar.expander('Obligations Contractuelles'):
                        selected_columns += st.multiselect('Sélectionnez les colonnes à afficher', group_4, default=group_4)

                    with st.sidebar.expander('MCM et TCM'):
                        selected_columns += st.multiselect('Sélectionnez les colonnes à afficher', group_5, default=group_5)

                    with st.sidebar.expander('Obligations Financières'):
                        selected_columns += st.multiselect('Sélectionnez les colonnes à afficher', group_6, default=group_6)

                    with st.sidebar.expander('Avenants'):
                        selected_columns += st.multiselect('Sélectionnez les colonnes à afficher', group_7, default=group_7)

                    if selected_columns:
                        st.subheader('Tableau des données filtrées')
                        st.markdown("""
                            <style>
                                .stDataFrame tbody tr:hover td {
                                    background-color: #f0f0f0;
                                    cursor: pointer;
                                }
                                .stDataFrame tbody tr:hover td::after {
                                    content: attr(data-value);
                                    display: block;
                                    color: black;
                                    font-size: 14px;
                                    padding-top: 5px;
                                }
                            </style>
                        """, unsafe_allow_html=True)

                        st.dataframe(df[selected_columns], use_container_width=True)
                    else:
                        st.warning("Veuillez sélectionner au moins une colonne à afficher.")

                with col2:
                    centroid = gdf.geometry.centroid
                    map_center = [centroid.y.mean(), centroid.x.mean()]

                    m = folium.Map(location=map_center, zoom_start=6)
                    df['ID'] = df['ID'].astype(str)

                    for _, row in gdf_merged.iterrows():
                        if selected_compagnie == 'Tous' or row['Compagnie'] == selected_compagnie:
                            popup_content = f"""
                            <b>Bloc :</b> {row.get('Nom', 'N/A')}<br>
                            <b>Compagnie :</b> {row.get('Compagnie', 'N/A')}<br>
                            <b>Phase :</b> {row.get('Phases_actuelle', 'N/A')}<br>
                            <b>Commentaires :</b> {row.get('Commentaires1', 'N/A')}
                            """
                            folium.GeoJson(
                                row['geometry'],
                                tooltip=row.get('Nom', 'ID'),
                                popup=folium.Popup(popup_content, max_width=300)
                            ).add_to(m)

                    st.subheader("Carte des blocs pétroliers")
                    st_folium(m, width=800, height=600)
            else:
                st.warning("Aucun fichier .shp trouvé dans l'archive.")

# Lancer l'application
if __name__ == "__main__":
    main()
