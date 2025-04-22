import sys
import pandas as pd
from PyQt5.QtCore import Qt, QAbstractTableModel, QTimer
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableView, QPushButton, QHBoxLayout, QVBoxLayout, QWidget, \
    QLabel, QHeaderView, QStyledItemDelegate, QSpinBox, QAction, QTimeEdit, QComboBox

from settings import *

class SpinBoxDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        editor = QSpinBox(parent)
        editor.setMinimum(0)
        editor.setMaximum(100)
        editor.setAlignment(Qt.AlignCenter)
        return editor

    def setEditorData(self, editor, index):
        value = index.model().data(index, Qt.EditRole)
        editor.setValue(int(value))

    def setModelData(self, editor, model, index):
        value = editor.value()
        model.setData(index, value, Qt.EditRole)


class PandasModel(QAbstractTableModel):
    def __init__(self, data: pd.DataFrame):
        super().__init__()
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        col = index.column()
        row = index.row()

        if col == 0:
            if role == Qt.CheckStateRole:
                return Qt.Checked if self._data.iat[row, col] else Qt.Unchecked
            return None

        if role in (Qt.DisplayRole, Qt.EditRole):
            return str(self._data.iat[row, col])
        return None

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                colName = str(self._data.columns[section])
                if colName == 'required':
                    return ''
                return colName
        return None

    def flags(self, index):
        flags = super().flags(index)
        if index.column() == 0:
            flags |= Qt.ItemIsUserCheckable
        elif index.column() == 1:
            flags |= Qt.ItemIsEditable
        return flags

    def setData(self, index, value, role):
        col = index.column()
        row = index.row()

        if col == 0 and role == Qt.CheckStateRole:
            self._data.iat[row, col] = (value == Qt.Checked)
            self.dataChanged.emit(index, index)
            return True
        elif col == 1 and role == Qt.EditRole:
            try:
                int_value = int(value)
                if 0 <= int_value <= 100:
                    self._data.iat[row, col] = int_value
                    self.dataChanged.emit(index, index)
                    return True
            except ValueError:
                pass

        return False


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Starter')
        self.setWindowIcon(QIcon('icon/plane.png'))
        self.load_data()
        self.init_ui()

    def load_data(self):
        self.level_1_sheets = pd.read_excel(level_1_excel, sheet_name=None, dtype={'id': str})
        self.level_2_sheets = pd.read_excel(level_2_excel, sheet_name=None, dtype={'id': str})
        self.level_3_sheets = pd.read_excel(level_3_excel, sheet_name=None, dtype={'id': str})
        self.level_4_sheets = pd.read_excel(level_4_excel, sheet_name=None, dtype={'id': str})

        for level in [self.level_1_sheets, self.level_2_sheets,
                      self.level_3_sheets, self.level_4_sheets]:
            for name, df in level.items():
                df['id'] = df['id'].ffill()
                if 'Question' in df.columns:
                    df = (
                        df.groupby('id', as_index=False)
                        .agg({
                            'Question': ' '.join
                        })
                    )

                df.insert(0, 'required', True)
                df.insert(1, 'score', 10)
                level[name] = df

    def init_ui(self):
        self.toolbar = self.addToolBar('')
        self.toolbar.setMovable(False)
        self.toolbar.setContextMenuPolicy(Qt.PreventContextMenu)

        self.create_action('Environment', 0)
        self.create_action('Level1', 1)
        self.create_action('Level2', 2)
        self.create_action('Level3', 3)
        self.create_action('Level4', 4)

        self.env_page()

    def create_action(self, name, level):
        if level == 0:
            action = QAction(QIcon('icon/cloud.png'), 'Environment', self)
            action.triggered.connect(lambda: self.env_page())
            self.toolbar.addAction(action)
            return
        elif level == 1:
            action = QAction(QIcon('icon/one.png'), 'Level1', self)
        elif level == 2:
            action = QAction(QIcon('icon/two.png'), 'Level2', self)
        elif level == 3:
            action = QAction(QIcon('icon/three.png'), 'Level3', self)
        elif level == 4:
            action = QAction(QIcon('icon/four.png'), 'Level4', self)
        action.triggered.connect(lambda: self.level_page(level))
        self.toolbar.addAction(action)

    def ok_button_clicked(self):
        print(self.weatherPicker.currentIndex())

    def env_page(self):
        envLabel=QLabel(level_0_name)
        envLabel.setStyleSheet("font-weight:bold; margin-left:0px; margin-top:10px; margin-bottom:10px")

        timehBox = QHBoxLayout()
        timeLabel = QLabel('时间：')
        self.timePicker = QTimeEdit()
        timehBox.addWidget(timeLabel)
        timehBox.addWidget(self.timePicker)
        timehBox.addStretch(2)

        weatherhBox = QHBoxLayout()
        weatherLabel = QLabel('天气：')
        self.weatherPicker = QComboBox()
        self.weatherPicker.addItems(['晴','多云','雾','打雷','雷暴','雪','雨'])
        weatherhBox.addWidget(weatherLabel)
        weatherhBox.addWidget(self.weatherPicker)
        weatherhBox.addStretch(2)

        okButton = QPushButton("OK")
        okButton.clicked.connect(self.ok_button_clicked)
        bottomhBox = QHBoxLayout()
        bottomhBox.addStretch(1)
        bottomhBox.addWidget(okButton)

        vbox = QVBoxLayout()
        vbox.addWidget(envLabel)
        vbox.addLayout(timehBox)
        vbox.addLayout(weatherhBox)
        vbox.addStretch(3)
        vbox.addLayout(bottomhBox)

        central_widget = QWidget()
        central_widget.setLayout(vbox)
        self.setCentralWidget(central_widget)

    def level_page(self, level):
        if level == 1:
            sheets = self.level_1_sheets
        elif level == 2:
            sheets = self.level_2_sheets
        elif level == 3:
            sheets = self.level_3_sheets
        elif level == 4:
            sheets = self.level_4_sheets

        vbox = QVBoxLayout()

        if level == 1:
            levelLabel = QLabel(level_1_name)
        elif level == 2:
            levelLabel = QLabel(level_2_name)
        elif level == 3:
            levelLabel = QLabel(level_3_name)
        elif level == 4:
            levelLabel = QLabel(level_4_name)
        levelLabel.setStyleSheet("font-weight:bold; margin-left:0px; margin-top:10px; margin-bottom:10px")
        vbox.addWidget(levelLabel)

        for sheet_name, df in sheets.items():
            if not df.empty:
                # 标题
                label = QLabel(f"{sheet_name}")
                vbox.addWidget(label)

                # 表格
                model = PandasModel(df)
                table = QTableView()
                table.setModel(model)
                table.resizeColumnsToContents()

                # 设置委托
                table.setItemDelegateForColumn(1, SpinBoxDelegate())

                for col in range(model.columnCount()):
                    header = model.headerData(col, Qt.Horizontal)
                    if header not in ('required', 'score', 'id', 'Question', 'Content'):
                        table.hideColumn(col)

                # 调整列宽
                table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
                table.showColumn(0)
                table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Fixed)
                table.setColumnWidth(1, 80)
                table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Fixed)
                table.setColumnWidth(2, 80)
                table.horizontalHeader().setStretchLastSection(True)

                vbox.addWidget(table)


        okButton = QPushButton("OK")
        hbox = QHBoxLayout()
        hbox.addStretch(1)
        hbox.addWidget(okButton)
        vbox.addLayout(hbox)

        central_widget = QWidget()
        central_widget.setLayout(vbox)
        self.setCentralWidget(central_widget)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.resize(800, 1200)
    window.show()
    sys.exit(app.exec_())
