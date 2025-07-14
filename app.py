import streamlit as st
import pandas as pd
import requests
import plotly.graph_objects as go
from io import BytesIO
from datetime import datetime
import os
import time
import sqlite3
from json import JSONDecodeError
import numpy as np  # Added for floating-point comparison

# Configuration de la page
st.set_page_config(
    page_title="Onacc Climate Forecast Analysis",
    page_icon="⛅",
    layout="wide"
)

# Style CSS personnalisé
st.markdown("""
    <style>
    .main {background-color: #fff;}
    h1 {color: #1f77b4;}
    h2 {color: #2ca02c; border-bottom: 2px solid #1f77b4;}
    .stTextArea textarea {height: 150px;}
    .stDownloadButton button {width: 100%;}
    .logo {text-align: center;}
    .dataframe {margin-bottom: 20px;}
    .doc-section {margin: 20px 0; padding: 15px; background: #fff; border-radius: 10px;}
    code {background: #f0f0f0; padding: 2px 5px; border-radius: 3px;}
    .login-container {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        height: 100vh;
        background-color: #f0f2f5;
    }
    .login-form {
        background: white;
        padding: 2rem;
        border-radius: 10px;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        width: 300px;
    }
    .login-form h2 {
        margin-bottom: 1.5rem;
        text-align: center;
        color: #333;
    }
    .login-form input {
        width: 100%;
        padding: 0.75rem;
        margin-bottom: 1rem;
        border: 1px solid #ccc;
        border-radius: 5px;
        font-size: 1rem;
    }
    .login-form button {
        width: 100%;
        padding: 0.75rem;
        background-color: #1f77b4;
        color: white “‘”;
        border: none;
        border-radius: 5px;
        font-size: 1rem;
        cursor: pointer;
    }
    .login-form button:hover {
        background-color: #155a8a;
    }
    </style>
    """, unsafe_allow_html=True)

# Créer ou se connecter à la base de données SQLite
conn = sqlite3.connect('users.db')
c = conn.cursor()
c.execute('''
    CREATE TABLE IF NOT EXISTS users (
        username TEXT PRIMARY KEY,
        password TEXT
    )
''')
c.execute("INSERT OR IGNORE INTO users (username, password) VALUES ('admin', 'admin')")
c.execute("INSERT OR IGNORE INTO users (username, password) VALUES ('membre', 'client')")
conn.commit()
conn.close()

# Fonction d'authentification
def authenticate(username, password):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("SELECT * FROM users WHERE username = ? AND password = ?", (username, password))
    user = c.fetchone()
    conn.close()
    return user is not None

# Page de login
def login_page():
    with st.form("login_form"):
        st.markdown('<h2>Connexion</h2>', unsafe_allow_html=True)
        username = st.text_input("Nom d'utilisateur")
        password = st.text_input("Mot de passe", type="password")
        submit = st.form_submit_button("Se connecter")
    if submit:
        if authenticate(username, password):
            st.session_state.logged_in = True
            st.success("Connexion réussie !")
            st.rerun()
        else:
            st.error("Nom d'utilisateur ou mot de passe incorrect.")

# Gestion de l'état de connexion
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False

# Afficher la page de login si non connecté
if not st.session_state.logged_in:
    login_page()
else:
    # Navigation
    pages = {
        "Application": "app",
        "Documentation": "docs"
    }

    st.sidebar.title("Navigation")
    selected_page = st.sidebar.radio("Sélectionnez une page", list(pages.keys()))

    # Page de documentation
    if selected_page == "Documentation":
        st.title("📚 Documentation - Onacc Climate Forecast")
        
        with st.expander("## 🌟 Présentation générale", expanded=True):
            st.markdown("""
            # 🚀 Objectif de l'application
            Onacc Climate Forecast Analysis est une application interactive permettant d'obtenir des **prévisions météorologiques et climatiques** pour des localités spécifiques. L'application utilise des **coordonnées GPS**, des **fichiers Excel** ou une **saisie manuelle** pour générer des prévisions précises et interactives en exploitant l'API **ONACC-MC**.
            
            # 🌍 Fonctionnalités
            ### 📂 Importation de Données
            - Importation d’un **fichier Excel** contenant des localités (latitude, longitude, altitude, région, pays).
            - Sélection et **filtrage avancé** des localités par région et pays.
            - Visualisation interactive des localités sélectionnées.

            ### 🔎 Types de Prévisions
            L'utilisateur peut choisir entre plusieurs types de prévisions :
            1. **Prévisions décadaires (1 à 14 jours)**
            - Température maximale / minimale (°C)
            - Précipitations (mm)
            - Sélection de la période (jours fixes ou plage personnalisée)
            2. **Prévisions saisonnières (45 jours à 9 mois)**
            - Analyse des tendances climatiques sur des périodes prolongées
            3. **Projections climatiques (jusqu’à 2050)**
            - Simulation des changements climatiques à long terme
            - Sélection du **modèle climatique** utilisé

            ### 📊 Visualisation Interactive
            - **Graphiques dynamiques** pour afficher :
            - Température maximale/minimale sous forme de courbes.
            - Précipitations sous forme d’histogramme.
            - **Affichage personnalisé** selon le type de prévision.

            ### 📤 Exportation des Données
            - Sauvegarde des prévisions sous **Excel**.
            - Génération d'un fichier unique par région avec blocs ordonnés اگر plus de 300 localités.
            """)

        with st.expander("## 🛠 Guide d'utilisation", expanded=False):
            st.markdown("""
            ### Workflow principal
            1. **Importation des données** (Section 📤 Importer un fichier)
            2. **Filtrage des localités** (Région/Pays)
            3. **Sélection des coordonnées**
            4. **Configuration des paramètres**
            5. **Génération des prévisions**
            6. **Export des résultats**

            ### Préparation des données
            Format du fichier Excel requis :
            ```csv
            localite,latitude,longitude,altitude,region,country
            Yaoundé,3.8480,11.5021,726,Centre,Cameroun
            Douala,4.0511,9.7679,13,Littoral,Cameroun
            ```

            ### Formats d'export
            - Excel : Un fichier par région dans un dossier "Result/[Type_Date_Heure]", avec blocs ordonnés si > 300 localités.
            """)

        with st.expander("## 🚀 Déploiement", expanded=False):
            st.markdown("""
            ### Prérequis
            - Python 3.8+
            - Librairies : `streamlit pandas requests plotly openpyxl xlsxwriter`

            ### Installation
            ```bash
            git clone https://github.com/Dev-Onacc/onacc_climate_forecast_analysis.git
            pip install -r requirements.txt
            streamlit run app.py
            ```

            ### Déploiement en production
            1. **Docker** :
            ```dockerfile
            FROM python:3.9-slim
            COPY . /app
            WORKDIR /app
            RUN pip install -r requirements.txt
            CMD ["streamlit", "run", "app.py", "--server.port=8501"]
            ```
            """)

        with st.expander("## 🆘 Support technique", expanded=False):
            st.markdown("""
            ### Problèmes courants
            | Symptôme | Solution |
            |----------|----------|
            | Erreur API | Vérifier la connexion internet |
            | Format de fichier invalide | Valider les colonnes obligatoires |

            ### Contact support
            **Équipe Onacc - DSI**  
            📧 poum.bimbar@onacc.org  
            """)

        st.stop()

    # ================= FONCTIONS UTILITAIRES =================
    def create_dataframe(data, lat, lon, localite, mode):
        """Crée le DataFrame à partir des données API avec validation"""
        try:
            time_data = data["daily"]["time"]
        except KeyError:
            raise ValueError("Données temporelles manquantes dans la réponse API")
        
        df = pd.DataFrame({
            "Localite": localite,
            "Date": pd.to_datetime(time_data),
            "Latitude": lat,
            "Longitude": lon
        })

        param_mapping = {
            "temperature_2m_max": "Température max (°C)",
            "temperature_2m_min": "Température min (°C)",
            "precipitation_sum": "Précipitations (mm)"
        }

        for api_param, df_column in param_mapping.items():
            df[df_column] = data["daily"].get(api_param, None)

        return df

    def get_api_error(response_data, status_code):
        """Gestion améliorée des erreurs API"""
        error_message = f"Erreur HTTP {status_code}"
        error_mapping = {
            400: "Requête invalide - Vérifiez les paramètres",
            401: "Authentification requise",
            403: "Accès refusé",
            404: "Endpoint introuvable",
            500: "Erreur serveur",
            429: "Trop de requêtes - Veuillez réessayer plus tard"
        }
        
        if isinstance(response_data, dict):
            return f"{error_mapping.get(status_code, error_message)} : {response_data.get('reason', 'Erreur inconnue')}"
        
        return error_mapping.get(status_code, error_message)

    def add_metadata(df, mode, params):
        """Ajoute les métadonnées de prévision"""
        df["Type de prévision"] = mode
        if mode == "Projections climatiques":
            df["Modèle climatique"] = params["model"]
        elif mode == "Prévisions saisonnières":
            df["Durée prévision"] = params["duration"]
        return df

    def create_visualization(df, mode):
        """Crée la visualisation adaptée au type de prévision"""
        fig = go.Figure()
        
        if "Température max (°C)" in df.columns:
            fig.add_trace(go.Scatter(
                x=df["Date"], 
                y=df["Température max (°C)"],
                name="Température max",
                line=dict(color='#FF5733', width=2),
                yaxis='y1'
            ))
        
        if "Température min (°C)" in df.columns:
            fig.add_trace(go.Scatter(
                x=df["Date"], 
                y=df["Température min (°C)"],
                name="Température min",
                line=dict(color='#3380FF', width=2),
                yaxis='y1'
            ))
        
        if "Précipitations (mm)" in df.columns:
            fig.add_trace(go.Bar(
                x=df["Date"], 
                y=df["Précipitations (mm)"],
                name="Précipitations",
                marker=dict(color='#33FF47', opacity=0.6),
                yaxis='y2'
            ))
        
        layout_config = {
            "title": f"Prévisions {mode} - Onacc",
            "xaxis": dict(title="Date", gridcolor='lightgray'),
            "yaxis": dict(
                title=dict(text="Température (°C)", font=dict(color='#1f77b4')),
                tickfont=dict(color='#1f77b4'),
                gridcolor='lightgray',
                side='left'
            ),
            "yaxis2": dict(
                title=dict(text="Précipitations (mm)", font=dict(color='#2ca02c')),
                tickfont=dict(color='#2ca02c'),
                overlaying='y',
                side='right'
            ),
            "legend": dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            ),
            "plot_bgcolor": 'rgba(255,255,255,0.9)',
            "hovermode": "x unified"
        }
        
        if mode == "Prévisions saisonnières":
            layout_config.update({
                "shapes": [dict(
                    type="rect",
                    xref="paper",
                    yref="paper",
                    x0=0, y0=0,
                    x1=1, y1=1,
                    fillcolor="rgba(200,200,200,0.2)",
                    line={"width": 0}
                )]
            })
        
        fig.update_layout(**layout_config)
        return fig

    # ================= INTERFACE UTILISATEUR =================
    col1, col2 = st.columns([1, 4])
    with col1:
        st.image("./logo.png", width=100)
    with col2:
        st.title("Onacc Climate Forecast Analysis")
        st.caption("Powered by Onacc")

    # Initialisation des variables de session
    if 'coordinates_set' not in st.session_state:
        st.session_state.coordinates_set = set()
    if 'selected_locations' not in st.session_state:
        st.session_state.selected_locations = pd.DataFrame(columns=['localite', 'latitude', 'longitude', 'region', 'source'])
    if 'coordinates' not in st.session_state:
        st.session_state.coordinates = ""

    # Fonction pour mettre à jour les coordonnées
    def update_coordinates():
        coords_list = [f"{row['latitude']},{row['longitude']}" for _, row in st.session_state.selected_locations.iterrows()]
        st.session_state.coordinates = ", ".join(coords_list)

    # Entrée de coordonnées brutes
    with st.expander("📝 Entrer des coordonnées brutes", expanded=True):
        with st.form("raw_coords_form"):
            localite = st.text_input("Nom de la localité")
            latitude = st.text_input("Latitude")
            longitude = st.text_input("Longitude")
            region = st.text_input("Région")
            submit_raw = st.form_submit_button("Ajouter la localité")

        if submit_raw:
            if localite and latitude and longitude and region:
                try:
                    lat = float(latitude)
                    lon = float(longitude)
                    new_row = pd.DataFrame({
                        'localite': [localite],
                        'latitude': [lat],
                        'longitude': [lon],
                        'region': [region],
                        'source': ['manuel']
                    })
                    st.session_state.selected_locations = pd.concat(
                        [st.session_state.selected_locations, new_row],
                        ignore_index=True
                    ).drop_duplicates(subset='localite')
                    st.session_state.coordinates_set.add((lat, lon))
                    update_coordinates()
                    st.success(f"Localité {localite} ajoutée avec succès.")
                except ValueError:
                    st.error("Latitude et longitude doivent être des nombres.")
            else:
                st.error("Veuillez remplir tous les champs.")

    # Importation du fichier Excel
    with st.expander("📤 Importer un fichier de localités", expanded=True):
        uploaded_file = st.file_uploader(
            "Télécharger un fichier Excel (format requis : localite,latitude,longitude,altitude,region,country)",
            type=["xlsx"]
        )

        if uploaded_file:
            try:
                df_locations = pd.read_excel(uploaded_file)
                required_columns = ['localite', 'latitude', 'longitude', 'region', 'country']
                
                if not all(col in df_locations.columns for col in required_columns):
                    st.error("Structure de fichier incorrecte! Les colonnes requises sont : localite, latitude, longitude, region, country")
                else:
                    df_locations = df_locations.dropna(subset=['latitude', 'longitude'])
                    df_locations['latitude'] = pd.to_numeric(df_locations['latitude'], errors='coerce')
                    df_locations['longitude'] = pd.to_numeric(df_locations['longitude'], errors='coerce')
                    df_locations = df_locations.dropna(subset=['latitude', 'longitude'])
                    
                    # Option de recherche
                    search_query = st.text_input("Rechercher une localité :", "")
                    if search_query:
                        df_locations = df_locations[df_locations['localite'].str.contains(search_query, case=False)]
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        selected_regions = st.multiselect(
                            "Filtrer par région:",
                            options=df_locations['region'].unique()
                        )
                    with col2:
                        selected_countries = st.multiselect(
                            "Filtrer par pays:",
                            options=df_locations['country'].unique()
                        )

                    if selected_regions:
                        df_locations = df_locations[df_locations['region'].isin(selected_regions)]
                    if selected_countries:
                        df_locations = df_locations[df_locations['country'].isin(selected_countries)]

                    total_localites = len(df_locations)
                    st.write(f"Nombre total de localités disponibles : {total_localites}")

                    # Option de sélection individuelle
                    selected_localite = st.selectbox("Sélectionnez une localité à ajouter :", df_locations['localite'].unique())
                    if selected_localite:
                        selected_row = df_locations[df_locations['localite'] == selected_localite].iloc[0]
                        selected_coord = (selected_row['latitude'], selected_row['longitude'])
                        if st.button("Ajouter aux coordonnées sélectionnées"):
                            if selected_coord not in st.session_state.coordinates_set:
                                st.session_state.coordinates_set.add(selected_coord)
                                new_row = selected_row.to_frame().T
                                new_row['source'] = 'excel'
                                st.session_state.selected_locations = pd.concat(
                                    [st.session_state.selected_locations, new_row],
                                    ignore_index=True
                                ).drop_duplicates(subset='localite')
                                update_coordinates()
                                st.success(f"Coordonnées de {selected_localite} ajoutées.")
                            else:
                                st.warning(f"Coordonnées de {selected_localite} déjà sélectionnées.")

                    # Bouton pour sélectionner toutes les localités d'une région
                    with st.expander("Sélectionner toutes les localités d'une région"):
                        for region in df_locations['region'].unique():
                            if st.button(f"Sélectionner toutes les localités de {region}"):
                                region_locations = df_locations[df_locations['region'] == region]
                                for _, row in region_locations.iterrows():
                                    selected_coord = (row['latitude'], row['longitude'])
                                    if selected_coord not in st.session_state.coordinates_set:
                                        st.session_state.coordinates_set.add(selected_coord)
                                        new_row = row.to_frame().T
                                        new_row['source'] = 'excel'
                                        st.session_state.selected_locations = pd.concat(
                                            [st.session_state.selected_locations, new_row],
                                            ignore_index=True
                                        ).drop_duplicates(subset='localite')
                                update_coordinates()
                                st.success(f"Toutes les localités de {region} ont été ajoutées.")

            except Exception as e:
                st.error(f"Erreur de lecture du fichier : {str(e)}")

    # Gestion des localités sélectionnées
    st.write("Localités sélectionnées :")
    if not st.session_state.selected_locations.empty:
        # Barre de recherche
        search_term = st.text_input("Rechercher une localité dans la sélection :")
        if search_term:
            filtered_locations = st.session_state.selected_locations[
                st.session_state.selected_locations['localite'].str.contains(search_term, case=False)
            ]
        else:
            filtered_locations = st.session_state.selected_locations

        # Afficher le tableau
        st.dataframe(filtered_locations[['localite', 'latitude', 'longitude', 'region']])

        # Boutons de suppression individuelle
        for index, row in filtered_locations.iterrows():
            if st.button(f"Supprimer {row['localite']}", key=f"delete_{row['localite']}"):
                st.session_state.selected_locations = st.session_state.selected_locations[
                    st.session_state.selected_locations['localite'] != row['localite']
                ]
                st.session_state.coordinates_set.remove((row['latitude'], row['longitude']))
                update_coordinates()
                st.success(f"Localité {row['localite']} supprimée.")

        # Bouton pour supprimer toutes les localités
        if st.button("Supprimer toutes les localités sélectionnées"):
            st.session_state.selected_locations = pd.DataFrame(columns=['localite', 'latitude', 'longitude', 'region', 'source'])
            st.session_state.coordinates_set = set()
            update_coordinates()
            st.success("Toutes les localités ont été supprimées.")
    else:
        st.write("Aucune localité sélectionnée.")

    # Formulaire de prévision
    with st.form("input_form"):
        coords = st.text_area(
            "Coordonnées sélectionnées (latitude,longitude):",
            value=st.session_state.coordinates,
            help="Format requis : 6.8399,13.2509, 6.4606,13.1184, ..."
        )

        forecast_mode = st.radio(
            "Type de prévision :",
            options=["Prévisions décadaires", "Prévisions saisonnières", "Projections climatiques"],
            horizontal=True
        )

        # Options spécifiques au type de prévision
        if forecast_mode == "Prévisions décadaires":
            forecast_type = st.radio(
                "Période de prévision :",
                ["Jours fixes", "Plage personnalisée"]
            )
            if forecast_type == "Jours fixes":
                forecast_days = st.selectbox("Durée :", [1, 3, 7, 10, 14], index=2)
            else:
                date_range = st.date_input("Plage de dates :", [])
                if len(date_range) == 2:
                    forecast_days = (date_range[1] - date_range[0]).days + 1
                else:
                    forecast_days = 1
            forecast_length = None
            start_date = None
            end_date = None
            model = None
        elif forecast_mode == "Prévisions saisonnières":
            forecast_length = st.selectbox(
                "Durée :",
                ["45 days", "3 months", "6 months", "9 months"]
            )
            forecast_days = None
            start_date = None
            end_date = None
            model = None
        elif forecast_mode == "Projections climatiques":
            col1, col2 = st.columns(2)
            with col1:
                start_date = st.date_input(
                    "Début :",
                    value=datetime(2020, 1, 1),
                    min_value=datetime(1950, 1, 1),
                    max_value=datetime(2050, 12, 31)
                )
            with col2:
                end_date = st.date_input(
                    "Fin :", 
                    value=datetime(2040, 1, 1),
                    min_value=start_date,
                    max_value=datetime(2050, 12, 31)
                )
            model = st.selectbox(
                "Modèle :",
                options=["MRI_AGCM3_2_S", "FGOALS_f3_H", "CMCC_CM2_VHR4"]
            )
            forecast_days = None
            forecast_length = None

        st.write("Paramètres à prédire :")
        temp_max = st.checkbox("Température maximale (2m)", True)
        temp_min = st.checkbox("Température minimale (2m)", True)
        precipitation = st.checkbox("Précipitations", True)
        
        submitted = st.form_submit_button("Générer la prévision")

    if submitted:
        try:
            if st.session_state.selected_locations.empty:
                st.error("Aucune localité sélectionnée. Veuillez sélectionner au moins une localité.")
            else:
                # Créer le dossier de résultats
                now = datetime.now()
                folder_name = f"{forecast_mode.replace(' ', '_')}_{now.strftime('%Y%m%d_%H%M%S')}"
                result_dir = os.path.join("Result", folder_name)
                os.makedirs(result_dir, exist_ok=True)

                # Séparer les localités par source
                excel_locations = st.session_state.selected_locations[st.session_state.selected_locations['source'] == 'excel']
                manual_locations = st.session_state.selected_locations[st.session_state.selected_locations['source'] == 'manuel']

                # Fonction pour traiter les données et générer des fichiers Excel
                def process_locations(locations, source_name):
                    if locations.empty:
                        return []

                    # Collecter tous les blocs
                    all_blocks = []
                    for region in locations['region'].unique():
                        region_locations = locations[locations['region'] == region]
                        total_localites_region = len(region_locations)
                        if total_localites_region > 200:
                            block_size = 150
                            blocks = [region_locations.iloc[i:i + block_size] for i in range(0, total_localites_region, block_size)]
                        else:
                            blocks = [region_locations]
                        for idx, block in enumerate(blocks, start=1):
                            block_name = f"{region}_{idx}"
                            all_blocks.append((region, block, block_name))

                    # Traiter les blocs par lots
                    batch_size = 8
                    all_df_blocks = []

                    for i in range(0, len(all_blocks), batch_size):
                        batch = all_blocks[i:i + batch_size]
                        for region, block, block_name in batch:
                            st.write(f"Traitement du bloc : {block_name} ({len(block)} localités)")

                            # Générer les coordonnées pour ce bloc
                            selected_coords = [
                                f"{row['latitude']},{row['longitude']}"
                                for _, row in block.iterrows()
                            ]
                            coords_list = []
                            for pair in selected_coords:
                                parts = pair.split(",")
                                coords_list.extend([parts[0].strip(), parts[1].strip()])

                            # Configuration API
                            base_params = {
                                "latitude": coords_list[::2],
                                "longitude": coords_list[1::2],
                                "daily": []
                            }

                            if forecast_mode == "Prévisions décadaires":
                                endpoint = "https://api.open-meteo.com/v1/forecast"
                                base_params.update({
                                    "forecast_days": forecast_days,
                                    "timezone": "auto",
                                    "daily": ["temperature_2m_max", "temperature_2m_min", "precipitation_sum"]
                                })
                            elif forecast_mode == "Prévisions saisonnières":
                                endpoint = "https://api.open-meteo.com/v1/forecast"
                                base_params.update({
                                    "daily": ["temperature_2m_max", "temperature_2m_min", "precipitation_sum"],
                                    "forecast_days": int(forecast_length.split()[0]) if "days" in forecast_length else int(forecast_length.split()[0]) * 30
                                })
                            elif forecast_mode == "Projections climatiques":
                                endpoint = "https://climate-api.open-meteo.com/v1/climate"
                                base_params.update({
                                    "start_date": start_date.strftime("%Y-%m-%d"),
                                    "end_date": end_date.strftime("%Y-%m-%d"),
                                    "models": model,
                                    "daily": ["temperature_2m_max", "temperature_2m_min", "precipitation_sum"]
                                })

                            # Appel API avec gestion des erreurs
                            max_retries = 3
                            retry_delay = 60  # en secondes
                            for attempt in range(max_retries):
                                response = requests.get(endpoint, params=base_params)
                                if response.status_code == 200:
                                    try:
                                        data = response.json()
                                        break
                                    except JSONDecodeError:
                                        st.error("Réponse API invalide (non JSON)")
                                        break
                                elif response.status_code == 429:  # Too Many Requests
                                    st.write("En attente d'une retransmission, veuillez patienter...")
                                    time.sleep(retry_delay)
                                else:
                                    st.error(get_api_error(response.json(), response.status_code))
                                    break
                            else:
                                st.error("Nombre maximal de tentatives atteint. Veuillez réessayer plus tard.")
                                return []

                            dfs = []
                            # Normaliser la réponse API en liste pour une gestion cohérente
                            if not isinstance(data, list):
                                data = [data]

                            # Vérifier la correspondance entre prévisions et coordonnées
                            if len(data) != len(base_params["latitude"]):
                                st.error(f"Nombre de prévisions ({len(data)}) ne correspond pas au nombre de coordonnées ({len(base_params['latitude'])})")
                                continue

                            for idx_data, (lat, lon) in enumerate(zip(base_params["latitude"], base_params["longitude"])):
                                if idx_data >= len(data):
                                    break
                                forecast = data[idx_data]
                                
                                if not isinstance(forecast, dict) or 'daily' not in forecast:
                                    st.warning(f"Données invalides pour {lat},{lon}")
                                    continue
                                
                                # Convertir lat et lon en float pour la comparaison
                                lat = float(lat)
                                lon = float(lon)
                                # Utiliser np.isclose pour une correspondance précise des coordonnées
                                matching_rows = block[
                                    np.isclose(block['latitude'], lat, atol=1e-5) & 
                                    np.isclose(block['longitude'], lon, atol=1e-5)
                                ]
                                if len(matching_rows) == 0:
                                    st.error(f"Aucune localité trouvée pour les coordonnées {lat},{lon}")
                                    continue
                                localite = matching_rows['localite'].values[0]
                                df_coord = create_dataframe(forecast, lat, lon, localite, forecast_mode)
                                dfs.append(df_coord)
                            
                            if dfs:  # S'assurer qu'il y a des données avant de concaténer
                                df_block = pd.concat(dfs, ignore_index=True)
                                df_block = add_metadata(df_block, forecast_mode, {
                                    "model": model if forecast_mode == "Projections climatiques" else None,
                                    "duration": forecast_length if forecast_mode == "Prévisions saisonnières" else None
                                })
                                df_block["Région"] = region
                                df_block["Bloc"] = block_name
                                all_df_blocks.append(df_block)
                        
                        if i + batch_size < len(all_blocks):
                            st.write("Attente d'une minute avant de continuer...")
                            time.sleep(60)

                    return all_df_blocks

                # Traiter les localités Excel et manuelles
                excel_blocks = process_locations(excel_locations, "excel")
                manual_blocks = process_locations(manual_locations, "manuel")

                # Générer les fichiers Excel séparés
                from collections import defaultdict
                def export_by_region(df_blocks, source_name):
                    if not df_blocks:
                        return
                    df_by_region = defaultdict(list)
                    for df_block in df_blocks:
                        region = df_block["Région"].iloc[0]
                        df_by_region[region].append(df_block)

                    for region, df_list in df_by_region.items():
                        df_region_all = pd.concat(df_list, ignore_index=True)
                        df_region_all = df_region_all.sort_values(by=["Bloc", "Localite"])

                        # Visualisation
                        fig = create_visualization(df_region_all, forecast_mode)
                        st.subheader(f"Prévisions pour la région : {region} ({source_name})")
                        st.plotly_chart(fig, use_container_width=True)
                        
                        # Exporter en Excel
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            df_region_all.to_excel(writer, index=False, sheet_name=region)
                        file_name = f"{region}_{source_name}.xlsx"
                        file_path = os.path.join(result_dir, file_name)
                        with open(file_path, "wb") as f:
                            f.write(output.getvalue())
                        
                        # Lien de téléchargement
                        st.download_button(
                            label=f"Télécharger {file_name}",
                            data=output.getvalue(),
                            file_name=file_name,
                            mime="application/vnd.ms-excel"
                        )

                export_by_region(excel_blocks, "excel")
                export_by_region(manual_blocks, "manuel")

        except Exception as e:
            st.error(f"Erreur : {str(e)}")