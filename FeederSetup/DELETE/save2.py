import pandas as pd
import PySimpleGUI as sg

# File paths
feeder_setup_file = r'D:\NX_BACKWORK\Feeder Setup_PROCESS\#Output\Verified\FeederSetup.xlsx'
bom_file = r'D:\NX_BACKWORK\Feeder Setup_PROCESS\#Output\Verified\BOM_List_OP.xlsx'

# Read data from Excel files
feeder_data = pd.read_excel(feeder_setup_file, sheet_name='FeederSetup')
bom_data = pd.read_excel(bom_file, sheet_name='BOM')

# Define column headers for tables
FEEDER_COLUMNS = ['Location', 'F_Part_No', 'QTY', 'Side', 'F_Ref_List']
BOM_COLUMNS = ['PartNumber', 'Long Des', 'Qty', 'RefList']

# Merge or compare data based on F_Part_No and PartNumber
merged_data = pd.merge(feeder_data[FEEDER_COLUMNS], bom_data[BOM_COLUMNS], how='left',
                       left_on='F_Part_No', right_on='PartNumber', suffixes=('_Feeder', '_BOM'))

# Merge or compare data based on F_Part_No and PartNumber
merged_data1 = pd.merge(bom_data[BOM_COLUMNS],feeder_data[FEEDER_COLUMNS], how='left',
                       left_on='PartNumber', right_on='F_Part_No', suffixes=('_BOM', '_Feeder'))

# Determine highlighting based on match
merged_data['Highlight'] = ''
merged_data.loc[merged_data['F_Part_No'] == merged_data['PartNumber'], 'Highlight'] = 'green'
merged_data.loc[merged_data['F_Part_No'] != merged_data['PartNumber'], 'Highlight'] = 'red'

# Determine highlighting based on match
merged_data1['Highlight'] = ''
merged_data1.loc[merged_data1['PartNumber'] == merged_data1['F_Part_No'], 'Highlight'] = 'green'
merged_data1.loc[merged_data1['PartNumber'] != merged_data1['F_Part_No'], 'Highlight'] = 'red'

# Function to create tab layouts with custom heading color
def create_tab_layout(data, headers):
    # Convert DataFrame to list of lists for sg.Table
    table_data = data.values.tolist()

    # Determine column widths based on content length
    col_widths = [max(len(str(row[i])) for row in table_data) + 5 for i in range(len(headers))]

    layout = [
        [sg.Table(values=table_data, headings=headers, display_row_numbers=False,
                  auto_size_columns=False, num_rows=min(25, len(data)), key='-TABLE-',
                  col_widths=col_widths, enable_events=True, text_color='black',
                  justification='left', vertical_scroll_only=False, alternating_row_color='lightyellow',
                  row_height=35, header_text_color='Black', font='Helvetica 10', size=(None, 300))]
    ]
    return layout

# Create separate data frames for BOT, TOP, and BOM
bot_data = merged_data[merged_data['Side'] == 'BOT']
top_data = merged_data[merged_data['Side'] == 'TOP']
bom_data_all = merged_data1

# Define layouts for each tab with custom heading color
bot_layout = create_tab_layout(bot_data[FEEDER_COLUMNS + BOM_COLUMNS + ['Highlight']], FEEDER_COLUMNS + BOM_COLUMNS)
top_layout = create_tab_layout(top_data[FEEDER_COLUMNS + BOM_COLUMNS + ['Highlight']], FEEDER_COLUMNS + BOM_COLUMNS)
#bom_layout = create_tab_layout(bom_data_all[FEEDER_COLUMNS + BOM_COLUMNS + ['Highlight']], FEEDER_COLUMNS + BOM_COLUMNS)
bom_layout = create_tab_layout(bom_data_all[BOM_COLUMNS + FEEDER_COLUMNS + ['Highlight']], BOM_COLUMNS + FEEDER_COLUMNS)

# Define the main layout with tabs
layout = [
    [sg.TabGroup([
        [sg.Tab('BOT', bot_layout, element_justification='center')],
        [sg.Tab('TOP', top_layout, element_justification='center')],
        [sg.Tab('BOM', bom_layout, element_justification='center')]
    ], tab_background_color='lightgray', tab_location='topleft', expand_x=True, expand_y=True)]
]

# Create the window with resizable option
window = sg.Window('Feeder Setup Comparison', layout, resizable=True, size=(1000, 800))

# Event loop
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break

window.close()
