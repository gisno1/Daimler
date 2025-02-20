import streamlit as st
import pdfplumber
import pandas as pd
import io

def process_pdf(file):

    data = []

    with pdfplumber.open(file) as pdf:

        page = pdf.pages[1]  
        text = page.extract_text()
        lines = text.split('\n')

        for line in lines:
            if 'AFSCHRIJVINGSTABEL DOSSIER' in line:
                Boekstuknummer = line.split()[-1]
            if 'Start overeenkomst' in line:
                Factuurdatum = line.split()[-1]
                Factuurdatum = pd.to_datetime(Factuurdatum, dayfirst=True, errors='coerce')
                Boekjaar = pd.to_datetime(Factuurdatum, dayfirst=True).year
                Periode = pd.to_datetime(Factuurdatum, dayfirst=True).month
            if 'Kenteken' in line:
                Kenteken = line.split()[-1]

        for line in lines:
            parts = line.split()
            if len(parts) == 7 and parts[0].isdigit():
                periode, van, naar, leasetermijn, afschrijving, interest, restant = parts
                naar = naar.lstrip('-')
                data.append([van, naar, leasetermijn, afschrijving, interest, restant])

    columns = ['Van', 'Naar', 'Leasetermijn', 'Afschrijving', 'Interest', 'Restant']
    df = pd.DataFrame(data, columns=columns)

    df = df.melt(id_vars=['Van', 'Naar'], value_vars=['Afschrijving', 'Interest'], var_name='Type', value_name='Bedrag')
    df['Grootboekrekening'] = df['Type'].map({'Afschrijving': '1302', 'Interest': '4701'})

    df['Boekstuknummer'] = Boekstuknummer
    df['Boekjaar'] = Boekjaar
    df['Periode'] = Periode
    df['Factuurdatum'] = Factuurdatum
    df['Kenteken'] = Kenteken

    df.drop(columns=['Type'], inplace=True)

    df['Bedrag'] = (
        df['Bedrag']
        .str.replace('.', '', regex=False)
        .str.replace(',', '.', regex=False)
        .astype(float)
    )

    last_row = pd.DataFrame({
        'Van': [df.iloc[0]['Van']],
        'Naar': [df.iloc[-1]['Naar']],
        'Bedrag': [-df['Bedrag'].sum()],
        'Grootboekrekening': '1311',
        'Boekjaar': [Boekjaar],
        'Periode': [Periode],
        'Factuurdatum': [Factuurdatum],
        'Kenteken': [Kenteken],
        'Boekstuknummer': [Boekstuknummer]
    })

    amortisatie = pd.concat([df, last_row], ignore_index=True)
    amortisatie['Kenteken_strippedA'] = amortisatie['Kenteken'].str.replace('-', '', regex=False)
    
    matching = pd.read_excel('matching_table.xlsx')
    
    amortisatie = amortisatie.merge(matching[['Code', 'Kenteken_stripped']], 
                            left_on='Kenteken_strippedA',  
                            right_on='Kenteken_stripped',  
                            how='left')

    columns = [
        'Dagboek: Code', 'Boekjaar', 'Periode', 'Boekstuknummer', 'Omschrijving: Kopregel', 
        'Factuurdatum', 'Vervaldatum', 'Valuta', 'Wisselkoers', 'Betalingsconditie: Code', 
        'Ordernummer', 'Uw ref.', 'Betalingsreferentie', 'Code', 'Naam', 'Grootboekrekening', 
        'Omschrijving', 'BTW-code', 'BTW-percentage', 'Bedrag', 'Aantal', 'BTW-bedrag', 
        'Opmerkingen', 'Project', 'Van', 'Naar', '1099', 'Kostenplaats: Code', 'Kostenplaats: Omschrijving',
        'Kostendrager: Code', 'Kostendrager: Omschrijving'
    ]

    import_df = pd.DataFrame(columns=columns)
    import_df['Boekjaar'] = amortisatie['Boekjaar']
    import_df['Dagboek: Code'] = '01'
    import_df['Periode'] = amortisatie['Periode']
    import_df['Boekstuknummer'] = amortisatie['Boekstuknummer']
    import_df['Factuurdatum'] = amortisatie['Factuurdatum']
    import_df['Valuta'] = 'EUR'
    import_df['Betalingsconditie: Code'] = '0'
    import_df['Code'] = 201435
    import_df['Naam'] = "Kuijpers Fleet (tbv amortisatieschema's)"
    import_df['Grootboekrekening'] = amortisatie['Grootboekrekening']
    import_df['Bedrag'] = amortisatie['Bedrag']
    import_df['Van'] = amortisatie['Van']
    import_df['Naar'] = amortisatie['Naar']
    import_df['Kostenplaats: Code'] = amortisatie['Code']
    import_df['Kostenplaats: Omschrijving'] = amortisatie['Kenteken']

    new_row = import_df.iloc[0].copy()
    new_row['Bedrag'] = ''
    new_row['Van'] = ''
    new_row['Naar'] = ''
    new_row['Kostenplaats: Code'] = ''
    new_row['Kostenplaats: Omschrijving'] = ''
    import_df = pd.concat([pd.DataFrame([new_row]), import_df], ignore_index=True)

    import_df['Factuurdatum'] = pd.to_datetime(import_df['Factuurdatum'], format='mixed').dt.strftime('%d-%m-%Y')

    return import_df, Kenteken

def main():

    st.title('PDF naar Excel: Daimler amortisatieschema ')

    uploaded_file = st.file_uploader('Upload een PDF bestand', type='pdf')
    if uploaded_file is not None:
        df, kenteken = process_pdf(uploaded_file)
        
        # Toon de data
        st.write('Amortisatie schema', df.head())

        # Laat de gebruiker het bestand downloaden
        output = io.BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)

        st.download_button(label='Download het Excel bestand', data=output, file_name=f'{kenteken} 1.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    main()


