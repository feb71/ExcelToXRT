import streamlit as st
import pandas as pd

# Funksjon for å konvertere Excel-data til XRT-format
def excel_to_xrt_conditional(excel_df, output_file_path):
    with open(output_file_path, 'w', encoding='latin-1') as xrt_file:
        # Skrive header-informasjonen
        xrt_file.write('[Xref_Info]\n')
        xrt_file.write('version=3\n\n')
        
        # Løper gjennom hver rad i DataFrame og skriver relevant XRT-data
        for index, row in excel_df.iterrows():
            xrt_file.write(f'[Xref{index + 1}]\n')
            xrt_file.write('descr=\n')
            xrt_file.write('otypes=P\n')
            xrt_file.write('operation=0\n')
            xrt_file.write('import=0\n')
            xrt_file.write(f'match1=name:"Point_Code" val:"{row["Point_Code"]}" opr:"0"\n')
            
            # Sjekker hvert attributt og skriver kun ikke-tomme verdier
            for i, col in enumerate(["S_OBJID", "S_FCODE", "Bruk", "Dimensjon", "Rørdel", "Type", "Type_bend", "Applag"], start=1):
                if pd.notna(row[col]):
                    attribute_name = col if col != "Applag" else "$DISTRIBUTION_ATTR$"
                    xrt_file.write(f'operation{i}=name:"{attribute_name}" val:"{row[col]}"\n')
            xrt_file.write('\n')

# Streamlit-grensesnittet
st.title("Excel til XRT Konvertering")

# Last opp Excel-filen
uploaded_file = st.file_uploader("Last opp Excel-filen din", type=["xlsx"])

if uploaded_file is not None:
    # Leser Excel-filen
    excel_data = pd.read_excel(uploaded_file)
    
    # Viser en forhåndsvisning av dataene
    st.write("Forhåndsvisning av Excel-data:")
    st.dataframe(excel_data.head())

    # Opprett XRT-fil
    output_xrt_path = "output.xrt"
    excel_to_xrt_conditional(excel_data, output_xrt_path)
    
    # Tilby nedlasting av XRT-filen
    with open(output_xrt_path, "rb") as file:
        btn = st.download_button(
            label="Last ned XRT-filen",
            data=file,
            file_name="output.xrt",
            mime="application/octet-stream"
        )
