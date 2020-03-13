import json, requests, sys, pandas, datetime, openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment, Protection
from openpyxl.styles import NamedStyle, Font, Border, Side
import config as cfg


def format_date(p_date):
    data = datetime.datetime.strptime(p_date, '%Y-%m-%dT%H:%M:%S')
    new_date = data.date().strftime("%d/%m/%Y")
    return new_date


def format_title(title):
    new_title = title.replace(';', ' ')
    return new_title


def save_xlsx_file(wb):
    wb.save('output.xlsx')


def load_xlsx_file():
    return load_workbook('output.xlsx')


def create_xlsx_headers(ws):
    count = 0
    for i in range(1, 6):
        if ws.cell(row=1, column=i).value is None:
            break
        count += 1

    if count == 0:
        print('populate headers')
        ws['A1'] = 'Data de publicação'
        ws.column_dimensions['A'].width = 25
        ws['B1'] = 'Titulo'
        ws.column_dimensions['B'].width = 40
        ws['C1'] = 'Publicação'
        ws.column_dimensions['C'].width = 40
        ws['D1'] = 'Prazo'
        ws.column_dimensions['D'].width = 25
        ws['E1'] = 'Data de Prazo'
        ws.column_dimensions['E'].width = 25


def request_url():
    url = 'https://cadastroapi.aasp.org.br/api/Associado/intimacao?chave='+cfg.api_key+'&data=11-03-2020&diferencial=false'
    response = requests.get(url)
    response.raise_for_status()
    return response.text


def add_formulae_to_deadline_column():
    for row_number in range(2, len(worksheet['E'])):
        new_value = '=A' + str(row_number) + '+D' + str(row_number)
        cell = worksheet.cell(row=row_number, column=5, value=new_value)
        cell.style = date_style


def add_cells_to_worksheet(date, title, text, row_count):
    date_cell = worksheet.cell(row=row_count, column=1, value=date)
    date_cell.style = date_style
    title_cell = worksheet.cell(row=row_count, column=2, value=title)
    text_cell = worksheet.cell(row=row_count, column=3, value=text[0:500])
    text_cell.alignment = Alignment(horizontal='general', vertical='top', text_rotation=0, wrap_text=True,
                                    shrink_to_fit=False, indent=0)
def parse_data_to_cells():
    number_of_entries = len(worksheet['A'])
    for idx, d in enumerate(data['intimacoes']):
        row_count = idx + number_of_entries + 1
        date = format_date(d['jornal']['dataDisponibilizacao_Publicacao'])
        title = format_title(d['titulo'])
        text = format_title(d['textoPublicacao'])
        add_cells_to_worksheet(date, title, text,row_count)

# TODO: make a separete class to handle .xlsx
try:
    workbook = load_xlsx_file()
    date_style = 'new_datetime'

except FileNotFoundError:
    workbook = Workbook()
    date_style = NamedStyle(name='new_datetime', number_format='DD/MM/YYYY')

worksheet = workbook.active
create_xlsx_headers(worksheet)


# Main code for pulling data from AASP
data = pandas.read_json(request_url())





parse_data_to_cells()

add_formulae_to_deadline_column()

save_xlsx_file(workbook)
