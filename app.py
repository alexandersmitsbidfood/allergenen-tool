
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
import io

st.set_page_config(page_title="Allergenen Systeem Generator", layout="centered")

st.title("🍽️ Allergenen Excel Generator")
st.write("Upload de groothandel CSV van je klant en genereer direct een werkend NVWA-proof Excel systeem.")

uploaded_file = st.file_uploader("Upload de grossierslijst (.csv)", type="csv")

if uploaded_file is not None:
    try:
        # Lees de CSV in (sla de eerste 14 regels met leveranciersinfo over)
        df = pd.read_csv(uploaded_file, sep=';', skiprows=14)
        df = df.fillna('')
        
        allergenen_mapping = {
            'Gluten': '🌾 Gluten', 'Schaaldieren': '🦐 Schaaldieren', 'Ei': '🥚 Ei',
            'Vis': '🐟 Vis', "Aardnoten (pinda's)": '🥜 Pinda', 'Soja': '🌱 Soja',
            'Melk': '🥛 Melk', 'Noten': '🌰 Noten', 'Selderij': '🌿 Selderij',
            'Mosterd': '🌼 Mosterd', 'Sesam': '🪴 Sesam', 'Zwaveldioxide en sulfieten': '🍷 Sulfiet',
            'Lupine': '🪻 Lupine', 'Weekdieren': '🦑 Weekdieren'
        }
        
        cols_to_keep = ['Artikelomschrijving'] + list(allergenen_mapping.keys())
        df_db = df[[c for c in cols_to_keep if c in df.columns]].copy()
        df_db = df_db[df_db['Artikelomschrijving'] != '']
        df_db = df_db.drop_duplicates(subset=['Artikelomschrijving'])
        
        # Maak Excel in het geheugen aan
        wb = openpyxl.Workbook()
        header_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, name="Arial")
        bold_font = Font(bold=True, name="Arial")
        center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin_border = Border(left=Side(style='thin', color="BDC3C7"), right=Side(style='thin', color="BDC3C7"), top=Side(style='thin', color="BDC3C7"), bottom=Side(style='thin', color="BDC3C7"))

        # --- Tab 1: Menu Beheer ---
        ws_menu = wb.active
        ws_menu.title = "1. Menu Beheer"
        ws_menu.merge_cells('A1:P2')
        ws_menu['A1'] = "MENU BEHEER\nKies een gerecht, voer de componenten in en selecteer het exacte ingrediënt."
        ws_menu['A1'].font = Font(bold=True, size=12)
        
        headers_menu = ['Gerecht', 'Component', 'Ingrediënt Database'] + list(allergenen_mapping.values())
        ws_menu.append([])
        ws_menu.append(headers_menu)
        
        for col_num, header in enumerate(headers_menu, 1):
            cell = ws_menu.cell(row=4, column=col_num)
            cell.fill, cell.font, cell.alignment, cell.border = header_fill, header_font, center_align, thin_border
            
        ws_menu.column_dimensions['A'].width = 25
        ws_menu.column_dimensions['B'].width = 25
        ws_menu.column_dimensions['C'].width = 45
        for i in range(4, 18): ws_menu.column_dimensions[get_column_letter(i)].width = 12

        dv = DataValidation(type="list", formula1="'3. Database'!$A$2:$A$5000", allow_blank=True)
        ws_menu.add_data_validation(dv)
        dv.add('C5:C100')

        # Dummy structuur invoegen
        for row_idx in range(5, 55):
            for col_idx in range(4, 18):
                db_col = col_idx - 2
                formula = f'=IF($C{row_idx}="","",IFERROR(VLOOKUP($C{row_idx},\'3. Database\'!$A:$O,{db_col},0),""))'
                cell = ws_menu.cell(row=row_idx, column=col_idx, value=formula)
                cell.alignment, cell.border = center_align, thin_border

        # --- Tab 2: Klantenkaart ---
        ws_klant = wb.create_sheet("2. Klantenkaart (Print)", 1)
        ws_klant.merge_cells('A1:P2')
        ws_klant['A1'] = "ALLERGENEN OVERZICHT"
        ws_klant['A1'].font, ws_klant['A1'].alignment = Font(bold=True, size=16), center_align
        ws_klant.append([])
        ws_klant.append(['Gerecht', 'Onderdeel'] + list(allergenen_mapping.values()))
        
        for col_num, header in enumerate(ws_klant[4], 1):
            header.fill, header.font, header.alignment, header.border = header_fill, header_font, center_align, thin_border
            
        ws_klant.column_dimensions['A'].width = ws_klant.column_dimensions['B'].width = 25
        for i in range(3, 17): ws_klant.column_dimensions[get_column_letter(i)].width = 12

        for row_idx in range(5, 55):
            ws_klant.cell(row=row_idx, column=1, value=f'=IF(\'1. Menu Beheer\'!A{row_idx}="","",\'1. Menu Beheer\'!A{row_idx})').font = bold_font
            ws_klant.cell(row=row_idx, column=2, value=f'=IF(\'1. Menu Beheer\'!B{row_idx}="","",\'1. Menu Beheer\'!B{row_idx})')
            for col_idx in range(3, 17):
                menu_col = get_column_letter(col_idx + 1)
                formula = f'=IF(\'1. Menu Beheer\'!{menu_col}{row_idx}="Bevat","⚠️",IF(\'1. Menu Beheer\'!{menu_col}{row_idx}="Kan bevatten","❓",""))'
                ws_klant.cell(row=row_idx, column=col_idx, value=formula).alignment = center_align

        # --- Tab 3: Database ---
        ws_db = wb.create_sheet("3. Database", 2)
        db_headers = ['Artikelomschrijving'] + list(allergenen_mapping.keys())
        ws_db.append(db_headers)
        for r in dataframe_to_rows(df_db, index=False, header=False): ws_db.append(r)
        
        # Bestand opslaan in geheugen voor download
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.success("✅ Bestand succesvol gegenereerd!")
        st.download_button(
            label="📥 Download Excel Template",
            data=output,
            file_name="Allergenen_Systeem_Klant.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Er ging iets mis bij het inlezen. Weet je zeker dat dit de juiste CSV is? Foutmelding: {e}")