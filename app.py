import pandas as pd
import streamlit as st
from datetime import datetime
import os
import re

REMISSE_FILE = 'last_remise.txt'

def get_last_remise():
    if os.path.exists(REMISSE_FILE):
        with open(REMISSE_FILE, 'r') as f:
            return int(f.read().strip())
    else:
        return 0

def save_last_remise(n_remise):
    with open(REMISSE_FILE, 'w') as f:
        f.write(str(n_remise))

def generate_output_filename(cod_emet, cod_dest, n_remise):
    now = datetime.now()
    dat_gen = now.strftime("%Y%m%d")  # Format YYYYMMDD pour l'année et le mois de la génération
    numeric_cod_emet = extract_numeric_part(cod_emet)
    numeric_cod_dest = extract_numeric_part(cod_dest)
    return f"bvov{numeric_cod_emet}{numeric_cod_dest}{dat_gen}{n_remise}.unl"

def extract_numeric_part(code):
    # Utilisation de regex pour extraire la partie numérique du code
    numeric_part = re.search(r'\d+', code)
    if numeric_part:
        return numeric_part.group()
    else:
        return ''

def excel_to_text(input_file, output_file, cod_emet, cod_dest, n_remise, has_header, utilisateur):
    # Read Excel file
    numeric_cod_emet = extract_numeric_part(cod_emet)
    if has_header:
        df = pd.read_excel(input_file, header=None)
    else:
        df = pd.read_excel(input_file)
    
    # Generate header information
    now = datetime.now()
    year_gen = now.strftime("%Y")  # Année de la génération au format YYYY
    month_gen = now.strftime("%m")  # Mois de la génération au format MM
    
    # Write to text file
    with open(output_file, 'w') as f:
        # Write header
        f.write(f"@nom_fic:{output_file}\n")
        f.write("@des_fic:\n")
        f.write(f"@dat_gen:{year_gen}{month_gen}\n")  # Concaténation de l'année et du mois
        f.write(f"@heur_gen:{now.strftime('%H%M')}\n")
        f.write(f"@cod_emet:{cod_emet}\n")
        f.write(f"@cod_dest:{cod_dest}\n")
        f.write(f"@N_remise:{n_remise}\n")
        f.write(f"@nbr_enr:{len(df)}\n")
        # Place holder for @taille
        f.write("@taille(octets):\n")
        f.write(f"@utilisateur:{utilisateur}\n")
        
        # Iterate through each row of the DataFrame
        for index, row in df.iterrows():
            # Convert row to string with pipe separator
            line = f"{numeric_cod_emet}|{year_gen}|{month_gen}|{n_remise}|{'|'.join(map(str, row))}"
            # Write line to text file
            f.write(line + '\n')
    
    # Calculate file size after writing
    taille_octets = os.path.getsize(output_file)
    
    # Rewrite @taille with actual file size
    with open(output_file, 'r+') as f:
        content = f.read()
        f.seek(0)
        f.write(content.replace("@taille(octets):\n", f"@taille(octets):{taille_octets}\n"))
        f.truncate()

# Streamlit Interface en français
st.title('Conversion de fichier Excel en fichier Texte')

uploaded_file = st.file_uploader("Choisissez un fichier Excel", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.write("Fichier téléchargé:", uploaded_file.name)
    
    cod_emet = st.text_input("Code de l'émetteur (obligatoire)", value="1000(CNRST)")
    cod_dest = st.text_input("Code du destinataire (obligatoire)", value="46(TR DE RABAT)")
    last_remise = get_last_remise()
    n_remise = st.number_input("Numéro de remise (obligatoire)", value=last_remise + 1, min_value=1)
    utilisateur = st.text_input("Nom d'utilisateur (obligatoire)")
    has_header = st.checkbox("Le fichier Excel contient-il un en-tête ?", value=True)
    
    if st.button('Convertir'):
        with st.spinner('Conversion en cours...'):
            output_file = generate_output_filename(cod_emet, cod_dest, n_remise)
            excel_to_text(uploaded_file, output_file, cod_emet, cod_dest, n_remise, has_header, utilisateur)
            save_last_remise(n_remise)
        st.success(f'Conversion terminée. Le fichier a été enregistré sous le nom {output_file}.')
        
        # Bouton de téléchargement pour le fichier de sortie
        st.download_button(
            label="Télécharger le fichier de sortie",
            data=open(output_file, 'rb').read(),
            file_name=output_file,
            mime="text/plain"
        )
