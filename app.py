import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Calculateur d'occupation", layout="wide")

st.title("Calculateur de taux d'occupation des bungalows")

# Téléchargement du fichier
uploaded_file = st.file_uploader("Téléchargez votre fichier Excel", type=['xlsx', 'xls'])

if uploaded_file:
    try:
        # Lecture du fichier
        df = pd.read_excel(uploaded_file)
        
        # Sélection des colonnes
        try:
            df = df.iloc[:, [5, 6, 7]]  # Colonnes F, G, H
            df.columns = ['Date de début', 'Date de sortie', 'Bungalow']
            
            # Conversion des dates
            df['Date de début'] = pd.to_datetime(df['Date de début'])
            df['Date de sortie'] = pd.to_datetime(df['Date de sortie'])
            
            # Sélection de la période
            st.sidebar.header("Période d'analyse")
            annee = st.sidebar.number_input("Année", min_value=2000, max_value=2100, value=datetime.now().year)
            mois_debut = st.sidebar.number_input("Mois de début", min_value=1, max_value=12, value=1)
            mois_fin = st.sidebar.number_input("Mois de fin", min_value=1, max_value=12, value=12)
            
            if st.sidebar.button("Calculer le taux d'occupation"):
                # Calculs
                date_debut = datetime(annee, mois_debut, 1)
                if mois_fin == 12:
                    date_fin = datetime(annee + 1, 1, 1)
                else:
                    date_fin = datetime(annee, mois_fin + 1, 1)
                
                # Filtrage des réservations
                df_periode = df[
                    (df['Date de sortie'] > date_debut) & 
                    (df['Date de début'] < date_fin)
                ].copy()
                
                # Calcul du taux d'occupation
                jours_totaux = (date_fin - date_debut).days
                resultats = []
                
                for bungalow in df['Bungalow'].unique():
                    df_bungalow = df_periode[df_periode['Bungalow'] == bungalow]
                    jours_occupes = 0
                    
                    for _, row in df_bungalow.iterrows():
                        debut = max(row['Date de début'], date_debut)
                        fin = min(row['Date de sortie'], date_fin)
                        jours_occupes += max(0, (fin - debut).days)
                    
                    taux_occupation = (jours_occupes / jours_totaux) * 100 if jours_totaux > 0 else 0
                    resultats.append({
                        'Bungalow': bungalow,
                        'Jours occupés': jours_occupes,
                        'Taux d\'occupation (%)': round(taux_occupation, 2)
                    })
                
                # Affichage des résultats
                df_resultat = pd.DataFrame(resultats)
                df_resultat = df_resultat.sort_values('Taux d\'occupation (%)', ascending=False)
                
                st.header("Résultats")
                st.write(f"Période du {date_debut.strftime('%d/%m/%Y')} au {date_fin.strftime('%d/%m/%Y')}")
                st.write(f"Nombre total de jours dans la période: {jours_totaux}")
                
                # Affichage du tableau
                st.dataframe(df_resultat)
                
                # Bouton de téléchargement
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_resultat.to_excel(writer, index=False, sheet_name='Resultats')
                    writer.close()
                    processed_data = output.getvalue()
                
                st.download_button(
                    label="Télécharger les résultats en Excel",
                    data=processed_data,
                    file_name=f"taux_occupation_{mois_debut}_{mois_fin}_{annee}.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
                
        except Exception as e:
            st.error(f"Erreur lors du traitement des données : {str(e)}")
    
    except Exception as e:
        st.error(f"Erreur lors de la lecture du fichier : {str(e)}")