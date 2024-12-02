import sys
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTableView, QVBoxLayout, QWidget, QLineEdit, QHBoxLayout, 
                             QPushButton, QAction, QHeaderView, QFileDialog, QTabWidget, QLabel, QComboBox, QMessageBox, 
                             QGridLayout)
from PyQt5.QtCore import QAbstractTableModel, Qt

# Program D: LoadingList Verification interface_GUI/J0124-89P13
def program_D():

    class PandasModel(QAbstractTableModel):
        def __init__(self, data):
            QAbstractTableModel.__init__(self)
            self._data = data
            self._original_data = data.copy()

        def rowCount(self, parent=None):
            return self._data.shape[0]

        def columnCount(self, parent=None):
            return self._data.shape[1]

        def data(self, index, role=Qt.DisplayRole):
            if index.isValid():
                if role == Qt.DisplayRole:
                    return str(self._data.iloc[index.row(), index.column()])
            return None

        def headerData(self, section, orientation, role=Qt.DisplayRole):
            if orientation == Qt.Horizontal and role == Qt.DisplayRole:
                return self._data.columns[section]
            return None

        def sort(self, column, order):
            colname = self._data.columns.tolist()[column]
            self.layoutAboutToBeChanged.emit()
            self._data.sort_values(colname, ascending=order == Qt.AscendingOrder, inplace=True)
            self.layoutChanged.emit()

        def filter(self, column, text):
            if text:
                self._data = self._original_data[self._original_data[column].astype(str).str.contains(text, case=False, na=False)]
            else:
                self._data = self._original_data.copy()
            self.layoutChanged.emit()

    class MainWindow(QMainWindow):
        def __init__(self):
            super().__init__()

            # Load data
            feeder_setup_file = r'D:\NX_BACKWORK\Feeder Setup_PROCESS\#Output\Verified\FeederSetup.xlsx'
            bom_file = r'D:\NX_BACKWORK\Feeder Setup_PROCESS\#Output\Verified\BOM_List_OP.xlsx'

            try:
                feeder_data = pd.read_excel(feeder_setup_file, sheet_name='FeederSetup')
                bom_data = pd.read_excel(bom_file, sheet_name='BOM')
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load Excel files: {e}")
                return
            
            print("Feeder Data Loaded:")
            print(feeder_data.head())
            print("BOM Data Loaded:")
            print(bom_data.head())

            # Define column headers for tables
            FEEDER_COLUMNS = ['Location', 'F_Part_No', 'QTY', 'Side', 'ModelName', 'F_Ref_List']
            BOM_COLUMNS = ['PartNumber', 'Group', 'Priority', 'Long Des', 'Qty', 'RefList']
            
            merged_data = pd.merge(feeder_data[FEEDER_COLUMNS], bom_data[BOM_COLUMNS], how='left',
                                left_on='F_Part_No', right_on='PartNumber', suffixes=('_Feeder', '_BOM'))
            merged_data1 = pd.merge(bom_data[BOM_COLUMNS], feeder_data[FEEDER_COLUMNS], how='left',
                                    left_on='PartNumber', right_on='F_Part_No', suffixes=('_BOM', '_Feeder'))
            
            print("Merged Data:")
            print(merged_data.head())
            print("Merged Data 1:")
            print(merged_data1.head())

            # Determine highlighting based on match
            merged_data['Highlight'] = ''
            merged_data.loc[merged_data['F_Part_No'] == merged_data['PartNumber'], 'Highlight'] = 'green'
            merged_data.loc[merged_data['F_Part_No'] != merged_data['PartNumber'], 'Highlight'] = 'red'
            
            merged_data1['Highlight'] = ''
            merged_data1.loc[merged_data1['PartNumber'] == merged_data1['F_Part_No'], 'Highlight'] = 'green'
            merged_data1.loc[merged_data1['PartNumber'] != merged_data1['F_Part_No'], 'Highlight'] = 'red'
            
            # Separate data for tabs
            bot_data = merged_data[merged_data['Side'] == 'BOT']
            top_data = merged_data[merged_data['Side'] == 'TOP']
            bom_data_all = merged_data1
            
            print("BOT Data:")
            print(bot_data.head())
            print("TOP Data:")
            print(top_data.head())
            print("BOM Data All:")
            print(bom_data_all.head())

            # Calculate percentages
            self.bot_percentage_green = self.calculate_percentage(bot_data, 'green')
            self.bot_percentage_red = self.calculate_percentage(bot_data, 'red')
            self.top_percentage_green = self.calculate_percentage(top_data, 'green')
            self.top_percentage_red = self.calculate_percentage(top_data, 'red')
            self.bom_percentage_green = self.calculate_percentage(bom_data_all, 'green')
            self.bom_percentage_red = self.calculate_percentage(bom_data_all, 'red')

            # Calculate counts
            self.bot_count_green, self.bot_count_red = self.calculate_counts(bot_data)
            self.top_count_green, self.top_count_red = self.calculate_counts(top_data)
            self.bom_count_green, self.bom_count_red = self.calculate_counts(bom_data_all)

            # Create table views
            self.bot_view = QTableView()
            self.top_view = QTableView()
            self.bom_view = QTableView()
            self.compare_view = QTableView()
            
            self.bot_model = PandasModel(bot_data)
            self.top_model = PandasModel(top_data)
            self.bom_model = PandasModel(bom_data_all)
            self.compare_model = PandasModel(pd.DataFrame(columns=bot_data.columns))
            
            self.bot_view.setModel(self.bot_model)
            self.top_view.setModel(self.top_model)
            self.bom_view.setModel(self.bom_model)
            self.compare_view.setModel(self.compare_model)
            
            # Enable sorting
            self.bot_view.setSortingEnabled(True)
            self.top_view.setSortingEnabled(True)
            self.bom_view.setSortingEnabled(True)
            self.compare_view.setSortingEnabled(True)
            
            # Connect double-click event to the custom slot
            self.bot_view.doubleClicked.connect(self.show_cell_content)
            self.top_view.doubleClicked.connect(self.show_cell_content)
            self.bom_view.doubleClicked.connect(self.show_cell_content)
            self.compare_view.doubleClicked.connect(self.show_cell_content)

            # Create filter input
            self.bot_filter_column = QComboBox()
            self.bot_filter_column.addItems(bot_data.columns)
            self.bot_filter_text = QLineEdit()
            self.bot_filter_text.setPlaceholderText('Filter BOT...')
            self.bot_filter_text.textChanged.connect(lambda text: self.bot_model.filter(self.bot_filter_column.currentText(), text))
            
            self.top_filter_column = QComboBox()
            self.top_filter_column.addItems(top_data.columns)
            self.top_filter_text = QLineEdit()
            self.top_filter_text.setPlaceholderText('Filter TOP...')
            self.top_filter_text.textChanged.connect(lambda text: self.top_model.filter(self.top_filter_column.currentText(), text))
            
            self.bom_filter_column = QComboBox()
            self.bom_filter_column.addItems(bom_data_all.columns)
            self.bom_filter_text = QLineEdit()
            self.bom_filter_text.setPlaceholderText('Filter BOM...')
            self.bom_filter_text.textChanged.connect(lambda text: self.bom_model.filter(self.bom_filter_column.currentText(), text))

            self.compare_filter_highlight = QComboBox()
            self.compare_filter_highlight.addItems(['', 'green', 'red'])
            self.compare_filter_highlight.currentTextChanged.connect(self.update_compare_table)
            
            self.compare_filter_part_no = QLineEdit()
            self.compare_filter_part_no.setPlaceholderText('Filter Part No...')
            self.compare_filter_part_no.textChanged.connect(self.update_compare_table)
            
            # Layout
            bot_layout = QVBoxLayout()
            bot_filter_layout = QHBoxLayout()
            bot_filter_layout.addWidget(QLabel("Filter by:"))
            bot_filter_layout.addWidget(self.bot_filter_column)
            bot_filter_layout.addWidget(self.bot_filter_text)
            bot_layout.addLayout(bot_filter_layout)
            bot_layout.addWidget(self.bot_view)
            
            top_layout = QVBoxLayout()
            top_filter_layout = QHBoxLayout()
            top_filter_layout.addWidget(QLabel("Filter by:"))
            top_filter_layout.addWidget(self.top_filter_column)
            top_filter_layout.addWidget(self.top_filter_text)
            top_layout.addLayout(top_filter_layout)
            top_layout.addWidget(self.top_view)
            
            bom_layout = QVBoxLayout()
            bom_filter_layout = QHBoxLayout()
            bom_filter_layout.addWidget(QLabel("Filter by:"))
            bom_filter_layout.addWidget(self.bom_filter_column)
            bom_filter_layout.addWidget(self.bom_filter_text)
            bom_layout.addLayout(bom_filter_layout)
            bom_layout.addWidget(self.bom_view)
            
            compare_layout = QVBoxLayout()
            compare_filter_layout = QHBoxLayout()
            compare_filter_layout.addWidget(QLabel("Filter by Highlight:"))
            compare_filter_layout.addWidget(self.compare_filter_highlight)
            compare_filter_layout.addWidget(QLabel("Part No:"))
            compare_filter_layout.addWidget(self.compare_filter_part_no)
            compare_layout.addLayout(compare_filter_layout)
            compare_layout.addWidget(self.compare_view)
            
            self.bot_widget = QWidget()
            self.bot_widget.setLayout(bot_layout)
            
            self.top_widget = QWidget()
            self.top_widget.setLayout(top_layout)
            
            self.bom_widget = QWidget()
            self.bom_widget.setLayout(bom_layout)
            
            self.compare_widget = QWidget()
            self.compare_widget.setLayout(compare_layout)

            # Summary table
            self.summary_view = QTableView()
            summary_data = pd.DataFrame({
                'Category': ['BOT', 'TOP', 'BOM'],
                'Green %': [self.bot_percentage_green, self.top_percentage_green, self.bom_percentage_green],
                'Red %': [self.bot_percentage_red, self.top_percentage_red, self.bom_percentage_red],
                'Green Count': [self.bot_count_green, self.top_count_green, self.bom_count_green],
                'Red Count': [self.bot_count_red, self.top_count_red, self.bom_count_red]
            })
            self.summary_model = PandasModel(summary_data)
            self.summary_view.setModel(self.summary_model)

            compare_layout.addWidget(QLabel("Summary:"))
            compare_layout.addWidget(self.summary_view)

            # Create tabs
            self.tabs = QTabWidget()
            self.tabs.addTab(self.bot_widget, f"BOT (Green: {self.bot_percentage_green:.2f}%, Red: {self.bot_percentage_red:.2f}%)")
            self.tabs.addTab(self.top_widget, f"TOP (Green: {self.top_percentage_green:.2f}%, Red: {self.top_percentage_red:.2f}%)")
            self.tabs.addTab(self.bom_widget, f"BOM (Green: {self.bom_percentage_green:.2f}%, Red: {self.bom_percentage_red:.2f}%)")
            self.tabs.addTab(self.compare_widget, "Compare")
            
            self.setCentralWidget(self.tabs)
            
            # Menu
            self.create_menu()

        def create_menu(self):
            menubar = self.menuBar()
            file_menu = menubar.addMenu('File')
            
            print_action = QAction('Export to Excel', self)
            print_action.triggered.connect(self.export_to_excel)
            
            file_menu.addAction(print_action)

        def export_to_excel(self):
            file_path, _ = QFileDialog.getSaveFileName(self, "Save Excel File", "", "Excel Files (*.xlsx);;All Files (*)")
            if file_path:
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    self.bot_model._data.to_excel(writer, sheet_name='BOT', index=False)
                    self.top_model._data.to_excel(writer, sheet_name='TOP', index=False)
                    self.bom_model._data.to_excel(writer, sheet_name='BOM', index=False)
                    self.compare_model._data.to_excel(writer, sheet_name='Compare', index=False)
                QMessageBox.information(self, "Export Successful", f"Data exported to {file_path}")

        def update_compare_table(self):
            highlight_filter = self.compare_filter_highlight.currentText()
            part_no_filter = self.compare_filter_part_no.text()
            
            # Filter based on Highlight
            if highlight_filter:
                filtered_data = self.bom_model._original_data[self.bom_model._original_data['Highlight'] == highlight_filter]
            else:
                filtered_data = self.bom_model._original_data.copy()
            
            # Filter based on Part No
            if part_no_filter:
                filtered_data = filtered_data[filtered_data['PartNumber'].astype(str).str.contains(part_no_filter, case=False, na=False)]
            
            # Update Compare Model
            self.compare_model._data = filtered_data
            self.compare_model.layoutChanged.emit()

        def show_cell_content(self, index):
            cell_content = self.sender().model().data(index)
            QMessageBox.information(self, "Cell Content", cell_content)

        def calculate_percentage(self, data, highlight_color):
            total_count = len(data)
            if total_count == 0:
                return 0.0
            highlight_count = len(data[data['Highlight'] == highlight_color])
            return (highlight_count / total_count) * 100

        def calculate_counts(self, data):
            green_count = len(data[data['Highlight'] == 'green'])
            red_count = len(data[data['Highlight'] == 'red'])
            return green_count, red_count

    app = QApplication(sys.argv)
    window = MainWindow()
    window.setWindowTitle('Feeder Setup Comparison')
    window.setGeometry(100, 100, 1200, 800)
    window.show()
    #sys.exit(app.exec_())
    app.exec_()

program_D()
