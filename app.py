import pandas as pd
import streamlit as st
from datetime import datetime
import os

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
    dat_gen = now.strftime("%y%m%d")
    return f"bvov{cod_emet}{cod_dest}0000{dat_gen}{n_remise}.unl"

def excel_to_text(input_file, output_file, cod_emet, cod_dest, n_remise):
    # Lire le fichier Excel
    df = pd.read_excel(input_file)
    
    # Générer les informations pour l'en-tête
    now = datetime.now()
    dat_gen = now.strftime("%d%m%y")
    heur_gen = now.strftime("%H%M")
    
    # Ouvrir le fichier texte en mode écriture
    with open(output_file, 'w') as f:
        # Écrire l'en-tête
        f.write(f"@nom_fic:{output_file}\n")
        f.write("@des_fic:\n")
        f.write(f"@dat_gen:{dat_gen}\n")
        f.write(f"@heur_gen:{heur_gen}\n")
        f.write(f"@cod_emet:{cod_emet}\n")
        f.write(f"@cod_dest:{cod_dest}\n")
        f.write(f"@N_remise:{n_remise}\n")
        
        # Parcourir chaque ligne du DataFrame
        for index, row in df.iterrows():
            # Convertir la ligne en chaîne de caractères avec séparation par des pipes
            line = '|'.join(map(str, row.values))
            # Écrire la ligne dans le fichier texte
            f.write(line + '\n')

# Interface Streamlit
st.title('Transformation de fichier Excel en fichier Texte')

uploaded_file = st.file_uploader("Choisissez un fichier Excel", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.write("Fichier téléchargé:", uploaded_file.name)
    
    cod_emet = st.text_input("Code de l'émetteur (obligatoire)", value="1000(CNRST)")
    cod_dest = st.text_input("Code du destinataire (obligatoire)", value="46(TR DE RABAT)")
    last_remise = get_last_remise()
    n_remise = st.number_input("Numéro de remise (obligatoire)", value=last_remise + 1, min_value=1)
    
    if st.button('Transformer'):
        with st.spinner('Transformation en cours...'):
            output_file = generate_output_filename(cod_emet, cod_dest, n_remise)
            excel_to_text(uploaded_file, output_file, cod_emet, cod_dest, n_remise)
            save_last_remise(n_remise)
        st.success(f'Transformation terminée. Le fichier a été enregistré sous le nom {output_file}.')

        if st.button("Afficher le fichier texte"):
            with open(output_file, 'r') as f:
                st.text(f.read())
