
import pandas as pd
from pandas.io.formats import excel
excel.ExcelFormatter.header_style = None
import openpyxl
from functions import set_column_autowidth, is_float, getValue

FOLDER = "data"
INPUT_FILE = "report_Loads_2025_06_17_11_25_54.csv"
OUTPUT_FILE = "output.xlsx"

table = pd.read_csv(f"{FOLDER}/{INPUT_FILE}")
table = table.fillna('')
col = table.columns

table['Count'] = int

for index, row in table.iterrows():
    table.at[index,'CURRENT'] = "" if getValue(table.at[index,'CURRENT']) == 0 else getValue(table.at[index,'CURRENT'])

for index, row in table.iterrows():
    count = 1
    for index2, row2 in table.iterrows():
        if index2 > index:
            if (row.drop(col[0]) == row2.drop(col[0])).all():
                count += 1
                table.at[index, col[0]] = table.at[index, col[0]] +", "+ table.at[index2, col[0]]
                table = table.drop(index=index2)
    table.at[index, 'Count'] = count

table = table.dropna()

voltages = []
temp_array = []

for index, row in table.iterrows():
    temp_row = row
    line = row["PowerNets"]
    lines = line.split(" ")
    for l in lines:
        if l != "" and "+" in l:
            if l.find("_") == -1:
                l_short = l
            else:
                l_short = l[0:l.find("_")]
            if is_float(l_short.replace("V", ".")):
                voltageDigit = float(l_short.replace("V", "."))
            else:
                voltageDigit = -1
            if not([l, voltageDigit] in voltages):
                voltages.append([l, voltageDigit])
            
            temp_row["PowerNets"] = l
            temp_array.append(temp_row.values.tolist())
   
table_new = pd.DataFrame(temp_array)
table_new = table_new.drop_duplicates()

table_new = table_new.drop([2, 4], axis=1)
table_new[8] = ""
table_new[9] = ""
table_new = table_new.reindex(columns=[0, 1, 6, 3, 5, 8, 7, 9])


voltages_table = pd.DataFrame(voltages)
voltages_table = voltages_table.sort_values(1, ascending=False)

header = ["Обозначение", "Потребитель", "NetName", "Ток (CIP)", "Ток (Схема)", "Ток (Вручную)", "Количество", "Итоговый ток"]

with pd.ExcelWriter(f'{FOLDER}/{OUTPUT_FILE}', engine='xlsxwriter') as writer:
    voltages_table.to_excel(writer, sheet_name='Напряжения', index=False, header=["NetName", "Напряжение, В"])
    table_new.to_excel(writer, sheet_name='Потребители', index=False, header=header)

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from openpyxl.formatting.rule import CellIsRule

# STYLES FOR EXCEL
# HEADER_HEIGHT = 25
ROW_HEIGHT = 20
ALIGNMNET_STYLE_CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
THIN_BORDER = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)
GRAY_FILL = PatternFill(start_color='E2E2E2',
                        end_color='E2E2E2',
                        fill_type='solid')
WHITE_FILL = PatternFill(start_color='FFFFFF',
                        end_color='FFFFFF',
                        fill_type='solid')
RED_FILL = PatternFill(bgColor='FFC7CE')
RULE_LESS_THAN_ZERO = CellIsRule(
    operator='lessThan',
    formula=['0'],
    fill=RED_FILL
)
RULE_NOT_EMPTY = CellIsRule(
    operator='notEqual',
    formula=['""'],  
    border=THIN_BORDER
)
RULE_NOT_CURRENT = FormulaRule(
    formula=['AND(B$1 <> "", $A2 <> "")'],  
    border=THIN_BORDER
)

workbook = openpyxl.load_workbook("data/test.xlsx", data_only=False)

QUANTITY_OF_LOAD_ROWS =  len(table_new[1].unique()) + 5 + 1
MAX_READ_COLUMN = 100
MAX_READ_ROW = 100

# EDITING OF FIREST PAGE
sheet = workbook["Напряжения"]

for column in ['A', 'B', 'C']:
    sheet.column_dimensions[column].fill = GRAY_FILL
sheet.column_dimensions['C'].width = 500

for cell in sheet[1]:  # первая строка - шапка
    cell.font = Font(bold=True)
    cell.alignment = ALIGNMNET_STYLE_CENTER
    cell.border = THIN_BORDER
    cell.fill = GRAY_FILL

for row in sheet.iter_rows(min_row=2):
    for cell in row:
        current_column = cell.column
        cell.alignment = ALIGNMNET_STYLE_CENTER
        cell.border = THIN_BORDER
        if current_column == 1:
            cell.fill = GRAY_FILL

for row in range(2, 100):
    sheet.row_dimensions[row].height = ROW_HEIGHT
    for column in range(1, 3):
        cell = sheet.cell(row=row, column=column)
        cell.alignment = ALIGNMNET_STYLE_CENTER
        if row > len(voltages)+1:
            cell.fill = GRAY_FILL

set_column_autowidth(sheet, ['A', 'B'], 1.7)

sheet.conditional_formatting.add(f'B2:B{len(voltages)+1}', RULE_LESS_THAN_ZERO)
sheet.freeze_panes = 'A2'

# EDITING OF SECOND PAGE
sheet = workbook["Потребители"]

for column in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']:
    sheet.column_dimensions[column].fill = GRAY_FILL
sheet.column_dimensions['I'].width = 500

set_column_autowidth(sheet, ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'], 1.2)

for cell in sheet[1]:  # первая строка - шапка
    cell.font = Font(bold=True)
    cell.alignment = ALIGNMNET_STYLE_CENTER
    cell.border = THIN_BORDER
    cell.fill = GRAY_FILL

for row in sheet.iter_rows(min_row=2):
    current_row = row[0].row
    for cell in row:
        current_column = cell.column

        cell.alignment = ALIGNMNET_STYLE_CENTER
        cell.border = THIN_BORDER
        match current_column:
            case 1:
                cell.fill = GRAY_FILL
            case 8:
                cell.value = f'=IF(F{current_row}<>"",F{current_row}*G{current_row},IF(E{current_row}<>"",E{current_row}*G{current_row},IF(D{current_row}<>"",D{current_row}*G{current_row},-1)))'
        
for row in range(2, 100):
    sheet.row_dimensions[row].height = ROW_HEIGHT
    for column in range(1, 9):
        cell = sheet.cell(row=row, column=column)
        cell.alignment = ALIGNMNET_STYLE_CENTER
        if row > len(table_new)+1:
            cell.fill = GRAY_FILL

rule_formula = FormulaRule(
    formula=['$H2 < 0'],  
    fill=RED_FILL
)
sheet.conditional_formatting.add(f'F2:F{len(table_new)+1}', rule_formula)
sheet.freeze_panes = 'A2' 
workbook.create_sheet(title="Сводная таблица")

# EDITING OF THIRD PAGE
sheet = workbook["Сводная таблица"]

from openpyxl.worksheet.formula import ArrayFormula

sheet['A2'] = ArrayFormula(f"A2:A{QUANTITY_OF_LOAD_ROWS}", f'=IFERROR(IF(_xlfn.UNIQUE(Потребители!B2:B{MAX_READ_ROW}) <> "", _xlfn.UNIQUE(Потребители!B2:B{MAX_READ_ROW}), ""), "")')
sheet['B1'] = ArrayFormula(f"B1:{get_column_letter(MAX_READ_ROW)}1", f'=IF(TRANSPOSE(Напряжения!$A$2:$A${MAX_READ_ROW})<>"", TRANSPOSE(Напряжения!$A$2:$A${MAX_READ_ROW}),"")')

for row in range(2, QUANTITY_OF_LOAD_ROWS+1):
    for column in range(2,MAX_READ_COLUMN):
        column_letter = get_column_letter(column)
        value = ArrayFormula(f"{column_letter}{row}:{column_letter}{row}", f'=IF(IFERROR(INDEX(Потребители!$H$2:$H$100, MATCH(1, (Потребители!$B$2:$B$100=$A{row})*(Потребители!$C$2:$C$100={column_letter}$1), 0)), "")<>0, IFERROR(INDEX(Потребители!$H$2:$H$100, MATCH(1, (Потребители!$B$2:$B$100=$A{row})*(Потребители!$C$2:$C$100={column_letter}$1), 0)), ""), "")')
        sheet.cell(row=row, column=column, value=value)


for cell in sheet[1]:
    cell.font = Font(bold=True)
    cell.alignment = ALIGNMNET_STYLE_CENTER
    cell.alignment = Alignment(textRotation=45)

sheet.row_dimensions[1].height = 100
for column in range(2,MAX_READ_COLUMN):
    sheet.column_dimensions[get_column_letter(column)].width = 6

sheet_temp = workbook["Потребители"]
sheet.column_dimensions['A'].width = sheet_temp.column_dimensions['B'].width

for row in sheet.iter_rows(min_row=2):
    current_row = row[0].row
    sheet.row_dimensions[row[0].row].height = ROW_HEIGHT
    for cell in row:
        cell.alignment = ALIGNMNET_STYLE_CENTER

sheet.conditional_formatting.add(f'B2:{get_column_letter(MAX_READ_COLUMN)}{QUANTITY_OF_LOAD_ROWS}', RULE_NOT_CURRENT)
sheet.conditional_formatting.add(f'A1:A{MAX_READ_ROW}', RULE_NOT_EMPTY) 
sheet.conditional_formatting.add(f'A1:{get_column_letter(MAX_READ_COLUMN)}1', RULE_NOT_EMPTY) 
sheet.conditional_formatting.add(f'B2:{get_column_letter(MAX_READ_COLUMN)}{MAX_READ_ROW}', RULE_LESS_THAN_ZERO)  

row = QUANTITY_OF_LOAD_ROWS + 2

sheet.row_dimensions[row].height = ROW_HEIGHT
cell = sheet.cell(row=row, column=1, value="Напряжение, В")
cell.alignment = ALIGNMNET_STYLE_CENTER
cell.border = THIN_BORDER

sheet[f'B{row}'] = ArrayFormula(f"B{row}:{get_column_letter(MAX_READ_ROW)}{row}", f'=IF(TRANSPOSE(Напряжения!$B$2:$B${MAX_READ_ROW})<>"", TRANSPOSE(Напряжения!$B$2:$B${MAX_READ_ROW}),"")')

for column in range(2, MAX_READ_COLUMN):
    cell = sheet.cell(row=row, column=column)
    cell.alignment = ALIGNMNET_STYLE_CENTER

row += 1

sheet.row_dimensions[row].height = ROW_HEIGHT
cell = sheet.cell(row=row, column=1, value="Суммарный ток, А")
cell.alignment = ALIGNMNET_STYLE_CENTER
cell.border = THIN_BORDER

for column in range(2, MAX_READ_COLUMN):
    sheet.cell(row=row, column=column, value= f'=IF(SUM({get_column_letter(column)}2:{get_column_letter(column)}{row-2})>0,SUM({get_column_letter(column)}2:{get_column_letter(column)}{row-2}),"")')
    cell = sheet.cell(row=row, column=column)
    
    cell.alignment = ALIGNMNET_STYLE_CENTER

row += 1

sheet.row_dimensions[row].height = ROW_HEIGHT
cell = sheet.cell(row=row, column=1, value="Мощность, Вт")
cell.alignment = ALIGNMNET_STYLE_CENTER
cell.border = THIN_BORDER

for column in range(2, MAX_READ_COLUMN):
    sheet.cell(row=row, column=column, value= f'=IF(IFERROR({get_column_letter(column)}{row-2}*{get_column_letter(column)}{row-1}, "") > 0, IFERROR({get_column_letter(column)}{row-2}*{get_column_letter(column)}{row-1}, ""), "")')
    cell = sheet.cell(row=row, column=column)
    cell.alignment = ALIGNMNET_STYLE_CENTER

sheet.conditional_formatting.add(f'B{row-2}:{get_column_letter(MAX_READ_COLUMN)}{row}', RULE_NOT_CURRENT)

row += 1
sheet.row_dimensions[row].height = ROW_HEIGHT
cell = sheet.cell(row=row, column=1, value="Суммарная мощность, Вт")
cell.alignment = ALIGNMNET_STYLE_CENTER
cell.border = THIN_BORDER
cell = sheet.cell(row=row, column=2)
cell.alignment = ALIGNMNET_STYLE_CENTER
cell.border = THIN_BORDER
cell.value = f'=SUM(B{row-1}:{get_column_letter(MAX_READ_COLUMN)}{row-1})'

sheet.freeze_panes = 'A2' 

workbook.save(f"{FOLDER}/{OUTPUT_FILE}")



