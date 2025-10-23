import streamlit as st
import pandas as pd
import pdfplumber
import re
from datetime import datetime
from io import BytesIO

def process_pdf(factuur_file, matching_file='matching_table.xlsx'):
    matching = pd.read_excel(matching_file)
    
    # --- Normaliseer matching ---
    matching['Kenteken_stripped'] = matching['Kenteken_stripped'].str.upper().str.replace("-", "", regex=False)
    
    data = []

    with pdfplumber.open(factuur_file) as pdf:
        factuurnummer = ""
        factuurdatum = ""
        boekjaar = ""
        periode = ""

        for i, page in enumerate(pdf.pages):
            text = page.extract_text()

            # --- factuurgegevens alleen op de eerste pagina ---
            if i == 0:
                match = re.search(r"Factuur nr\..*?\n(\d+)\s+(\d{2}/\d{2}/\d{2})", text)
                if match:
                    factuurnummer = match.group(1)
                    factuurdatum_raw = match.group(2)
                    factuurdatum_dt = datetime.strptime(factuurdatum_raw, "%d/%m/%y")
                    boekjaar = factuurdatum_dt.year
                    periode = factuurdatum_dt.month
                    factuurdatum = factuurdatum_dt.strftime("%d-%m-%Y")

            # --- Per blok "Termijnbetaling" ---
            termijn_blokken = re.split(r"Termijnbetaling voor periode ", text)[1:]
            for blok in termijn_blokken:
                periode_match = re.match(r"(\d{2}/\d{2}/\d{2}) t/m (\d{2}/\d{2}/\d{2})", blok)
                if not periode_match:
                    continue

                start_periode = datetime.strptime(periode_match.group(1), "%d/%m/%y").strftime("%d-%m-%Y")
                eind_periode = datetime.strptime(periode_match.group(2), "%d/%m/%y").strftime("%d-%m-%Y")

                rij_matches = re.findall(
                    r"\d+\s+([A-Z0-9]{1,3}-[A-Z0-9]{1,3}-[A-Z0-9]{1,3})\s+\d+\s+[\d,]+\s+[\d,]+\s+([\d.,]+)",
                    blok
                )

                for kenteken, bedrag in rij_matches:
                    bedrag = bedrag.replace(".", "").replace(",", ".")
                    try:
                        bedrag = float(bedrag)
                    except ValueError:
                        bedrag = None

                    data.append({
                        "Factuurnummer": factuurnummer,
                        "Factuurdatum": factuurdatum,
                        "Boekjaar": boekjaar,
                        "Periode": periode,
                        "Kenteken": kenteken,
                        "Start periode": start_periode,
                        "Eind periode": eind_periode,
                        "Bedrag": bedrag
                    })

    factuur = pd.DataFrame(data)
    factuur['Omschrijving'] = factuur['Factuurnummer'].astype(str) + ' / SCANIA'
    factuur['Kenteken_stripped'] = factuur['Kenteken'].str.upper().str.replace("-", "", regex=False)

    # --- Merge met matching ---
    factuur = factuur.merge(
        matching[['Kenteken', 'Code', 'Kenteken_stripped']],
        left_on='Kenteken_stripped',
        right_on='Kenteken_stripped',
        how='left'
    )

    factuur = factuur.rename(columns={'Kenteken_x': 'Kenteken_factuur', 'Kenteken_y': 'Kenteken'})
    factuur = factuur.drop(columns=['Kenteken_stripped'])

    # --- Format voor import ---
    columns = [
        'Dagboek: Code', 'Boekjaar', 'Periode', 'Boekstuknummer', 'Omschrijving: Kopregel',
        'Factuurdatum', 'Vervaldatum', 'Valuta', 'Wisselkoers', 'Betalingsvoorwaarde: Code',
        'Ordernummer', 'Uw ref.', 'Betalingsreferentie', 'Relatiecode', 'Naam', 'Grootboekrekening',
        'Omschrijving', 'BTW-code', 'BTW-percentage', 'Bedrag', 'Aantal', 'BTW-bedrag',
        'Opmerkingen', 'Project', 'Van', 'Naar', 'Kostenplaats: Code', 'Kostenplaats: Omschrijving',
        'Kostendrager: Code', 'Kostendrager: Omschrijving'
    ]

    import_df = pd.DataFrame(columns=columns)
    import_df['Boekjaar'] = factuur['Boekjaar']
    import_df['Dagboek: Code'] = 60
    import_df['Periode'] = factuur['Periode']
    import_df['Factuurdatum'] = factuur['Factuurdatum']
    import_df['Omschrijving: Kopregel'] = factuur['Omschrijving']
    import_df['Omschrijving'] = factuur['Omschrijving']
    import_df['Uw ref.'] = factuur['Factuurnummer']
    import_df['Relatiecode'] = 272
    import_df['Grootboekrekening'] = 1313
    import_df['Bedrag'] = factuur['Bedrag']
    import_df['Van'] = factuur['Start periode']
    import_df['Naar'] = factuur['Eind periode']
    import_df['Kostenplaats: Code'] = factuur['Code']
    import_df['Kostenplaats: Omschrijving'] = factuur['Kenteken_factuur']

    # Extra eerste rij
    new_row = import_df.iloc[0].copy()
    new_row['Bedrag'] = pd.NA
    new_row['Kostenplaats: Code'] = pd.NA
    new_row['Kostenplaats: Omschrijving'] = ''
    import_df = pd.concat([pd.DataFrame([new_row]), import_df], ignore_index=True)

    return import_df


def main():
    st.title("Import Scania factuur")

    factuur_file = st.file_uploader("Upload je PDF-factuur", type=['pdf'])

    if factuur_file:
        processed_file = process_pdf(factuur_file)
        st.write("Verwerkte factuur:", processed_file.head())

        output = BytesIO()
        processed_file.to_excel(output, index=False)
        output.seek(0)

        st.download_button(
            label="Download verwerkte factuur",
            data=output,
            file_name="Scania_factuur.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.markdown(
            "<p style='font-size:12px; color:gray;'>Let op: sla het bestand op als Excel 97-2003 om te kunnen importeren in Exact.</p>",
            unsafe_allow_html=True
        )


if __name__ == "__main__":
    main()
