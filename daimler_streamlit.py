import streamlit as st
import pandas as pd
from io import BytesIO

def process_file(factuur_file):

    factuur = pd.read_excel(factuur_file)
    factuur = factuur.iloc[:-1]
    matching = pd.read_excel('matching_table.xlsx')
    
    factuur = factuur.merge(matching[['Kenteken', 'Code', 'Kenteken_stripped']],
                            left_on='Kenteken',
                            right_on='Kenteken_stripped',
                            how='left')
    
    factuur = factuur.rename(columns={'Kenteken_x': 'Kenteken_factuur'})
    factuur = factuur.rename(columns={'Kenteken_y': 'Kenteken'})
    factuur = factuur.drop(columns=['Kenteken_stripped'])
    
    factuur['Boekjaar'] = factuur['Factuurdatum'].dt.year
    factuur['Periode'] = factuur['Factuurdatum'].dt.month
    
    columns = [
        'Dagboek: Code', 'Boekjaar', 'Periode', 'Boekstuknummer', 'Omschrijving: Kopregel',
        'Factuurdatum', 'Vervaldatum', 'Valuta', 'Wisselkoers', 'Betalingsvoorwaarde: Code',
        'Ordernummer', 'Uw ref.', 'Betalingsreferentie', 'Code', 'Naam', 'Grootboekrekening',
        'Omschrijving', 'BTW-code', 'BTW-percentage', 'Bedrag', 'Aantal', 'BTW-bedrag',
        'Opmerkingen', 'Project', 'Van', 'Naar', 'Kostenplaats: Code', 'Kostenplaats: Omschrijving',
        'Kostendrager: Code', 'Kostendrager: Omschrijving'
    ]
    
    import_df = pd.DataFrame(columns=columns)
    import_df['Boekjaar'] = factuur['Boekjaar']
    import_df['Dagboek: Code'] = 60
    import_df['Periode'] = factuur['Periode']
    import_df['Factuurdatum'] = factuur['Factuurdatum']
    import_df['Uw ref.'] = factuur['Factuurnr']
    import_df['Code'] = 201404
    import_df['Grootboekrekening'] = 1311
    import_df['Bedrag'] = factuur['Bedrag excl']
    import_df['Van'] = factuur['Begin Periode']
    import_df['Naar'] = factuur['Eind Periode']
    import_df['Kostenplaats: Code'] = factuur['Code']
    import_df['Kostenplaats: Omschrijving'] = factuur['Kenteken_factuur']
    import_df['Factuurdatum'] = import_df['Factuurdatum'].dt.strftime('%d-%m-%Y')
    import_df['Van'] = import_df['Van'].dt.strftime('%d-%m-%Y')
    import_df['Naar'] = import_df['Naar'].dt.strftime('%d-%m-%Y')
    
    import_df = import_df.sort_values(by='Kostenplaats: Code', ascending=True, na_position='first')
    
    new_row = import_df.iloc[0].copy()
    new_row['Bedrag'] = ''
    new_row['Kostenplaats: Code'] = ''
    new_row['Kostenplaats: Omschrijving'] = ''
    import_df = pd.concat([pd.DataFrame([new_row]), import_df], ignore_index=True)
    
    return import_df


def main():

    st.title('Import Daimler factuur')
    
    factuur_file = st.file_uploader('Upload het Excel-factuurbestand', type=['xlsx'])
    
    if factuur_file:

        processed_file = process_file(factuur_file)
        st.write('Verwerkte factuur', processed_file.head())
                
        output = BytesIO()
        processed_file.to_excel(output, index=False)
        output.seek(0)

        st.download_button(label='Download verwerkte factuur',
                            data=output,
                            file_name='verwerkte_factuur.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    main()