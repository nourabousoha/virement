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
    # Read Excel file
    df = pd.read_excel(input_file)
    
    # Generate header information
    now = datetime.now()
    dat_gen = now.strftime("%d%m%y")
    heur_gen = now.strftime("%H%M")
    
    # Write to text file
    with open(output_file, 'w') as f:
        # Write header
        f.write(f"@nom_fic:{output_file}\n")
        f.write("@des_fic:\n")
        f.write(f"@dat_gen:{dat_gen}\n")
        f.write(f"@heur_gen:{heur_gen}\n")
        f.write(f"@cod_emet:{cod_emet}\n")
        f.write(f"@cod_dest:{cod_dest}\n")
        f.write(f"@N_remise:{n_remise}\n")
        
        # Iterate through each row of the DataFrame
        for index, row in df.iterrows():
            # Convert row to string with pipe separator
            line = '|'.join(map(str, row.values))
            # Write line to text file
            f.write(line + '\n')

# Streamlit Interface
st.title('Excel to Text File Conversion')

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.write("File uploaded:", uploaded_file.name)
    
    cod_emet = st.text_input("Sender code (required)", value="1000(CNRST)")
    cod_dest = st.text_input("Recipient code (required)", value="46(TR DE RABAT)")
    last_remise = get_last_remise()
    n_remise = st.number_input("Remittance number (required)", value=last_remise + 1, min_value=1)
    
    if st.button('Convert'):
        with st.spinner('Converting...'):
            output_file = generate_output_filename(cod_emet, cod_dest, n_remise)
            excel_to_text(uploaded_file, output_file, cod_emet, cod_dest, n_remise)
            save_last_remise(n_remise)
        st.success(f'Conversion completed. File saved as {output_file}.')
        
        # Download button for the output file
        st.download_button(
            label="Download Output File",
            data=open(output_file, 'rb').read(),
            file_name=output_file,
            mime="text/plain"
        )
