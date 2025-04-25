import csv
import os.path
import struct
import subprocess
import sys
import json
from importlib.metadata import metadata

import pandas as pd
from PyQt5.QtCore import Qt, QAbstractTableModel, QTime, QDir
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableView, QPushButton, QHBoxLayout, QVBoxLayout, QWidget, \
    QLabel, QHeaderView, QStyledItemDelegate, QSpinBox, QAction, QTimeEdit, QComboBox, QStackedWidget, QStyle, \
    QLineEdit, QFileSystemModel, QTreeView, QDialog, QTreeWidget, QTreeWidgetItem

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
        self.level1Sheets = None
        self.level2Sheets = None
        self.level3Sheets = None
        self.level4Sheets = None
        self.timePicker = None
        self.weatherPicker = None
        self.trainTimeMinuteInput = None
        self.trainTimeSecondInput = None
        self.nameInput = None
        self.loadData()
        self.initUi()

    def loadData(self):
        if os.path.exists(LEVEL_1_EXCEL):
            self.level1Sheets = pd.read_excel(LEVEL_1_EXCEL, sheet_name=None, dtype={'id': str})
        if os.path.exists(LEVEL_2_EXCEL):
            self.level2Sheets = pd.read_excel(LEVEL_2_EXCEL, sheet_name=None, dtype={'id': str})
        if os.path.exists(LEVEL_3_EXCEL):
            self.level3Sheets = pd.read_excel(LEVEL_3_EXCEL, sheet_name=None, dtype={'id': str})
        if os.path.exists(LEVEL_4_EXCEL):
            self.level4Sheets = pd.read_excel(LEVEL_4_EXCEL, sheet_name=None, dtype={'id': str})

        for level in [self.level1Sheets, self.level2Sheets,
                      self.level3Sheets, self.level4Sheets]:
            if level is None:
                continue
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

    def createAction(self, name, level):
        if level == 0:
            action = QAction(QIcon('icon/cloud.png'), 'Environment', self)
        elif level == 1:
            action = QAction(QIcon('icon/one.png'), 'Level1', self)
        elif level == 2:
            action = QAction(QIcon('icon/two.png'), 'Level2', self)
        elif level == 3:
            action = QAction(QIcon('icon/three.png'), 'Level3', self)
        elif level == 4:
            action = QAction(QIcon('icon/four.png'), 'Level4', self)
        elif level == 5:
            action = QAction(QIcon('icon/result.png'), 'Result', self)
        action.triggered.connect(lambda _, i=level: self.stacked.setCurrentIndex(i))
        self.toolbar.addAction(action)

    def createEnvPage(self):
        widget = QWidget()
        envLabel = QLabel(PAGE_0_NAME)
        envLabel.setStyleSheet("font-weight:bold; margin-left:0px; margin-top:10px; margin-bottom:10px")

        def setTimeByIndex(index):
            if index == 0:
                self.timePicker.setTime(QTime(6, 45))
            elif index == 1:
                self.timePicker.setTime(QTime(17, 30))
            elif index == 2:
                self.timePicker.setTime(QTime(12, 0))
            else:
                self.timePicker.setTime(QTime(0, 0))

        timehBox = QHBoxLayout()
        timeLabel = QLabel('时间：')
        self.timePicker = QTimeEdit()
        self.timePicker.setTime(QTime(6, 45))
        timeComboBox = QComboBox()
        timeComboBox.addItems(['黎明', '黄昏', '白天', '夜晚'])
        timeComboBox.currentIndexChanged.connect(lambda: setTimeByIndex(timeComboBox.currentIndex()))
        timehBox.addWidget(timeLabel)
        timehBox.addWidget(self.timePicker)
        timehBox.addWidget(timeComboBox)
        timehBox.addStretch(3)

        weatherhBox = QHBoxLayout()
        weatherLabel = QLabel('天气：')
        self.weatherPicker = QComboBox()
        self.weatherPicker.addItems(['晴', '多云', '雾', '雷', '雷雨', '雪', '雨', '霜'])
        weatherhBox.addWidget(weatherLabel)
        weatherhBox.addWidget(self.weatherPicker)
        weatherhBox.addStretch(2)

        trainTimehBox = QHBoxLayout()
        trainTimeLabel = QLabel('训练时长：')
        self.trainTimeMinuteInput = QSpinBox()
        trainTimeMinuteLabel = QLabel('分')
        self.trainTimeSecondInput = QSpinBox()
        trainTimeSecondLabel = QLabel('秒')
        trainTimehBox.addWidget(trainTimeLabel)
        trainTimehBox.addWidget(self.trainTimeMinuteInput)
        trainTimehBox.addWidget(trainTimeMinuteLabel)
        trainTimehBox.addWidget(self.trainTimeSecondInput)
        trainTimehBox.addWidget(trainTimeSecondLabel)
        trainTimehBox.addStretch(5)

        namehBox = QHBoxLayout()
        nameLabel = QLabel('姓名：')
        self.nameInput = QLineEdit()
        namehBox.addWidget(nameLabel)
        namehBox.addWidget(self.nameInput)
        namehBox.addStretch(2)

        vbox = QVBoxLayout(widget)
        vbox.addWidget(envLabel)
        vbox.addLayout(timehBox)
        vbox.addLayout(weatherhBox)
        vbox.addLayout(trainTimehBox)
        vbox.addLayout(namehBox)
        vbox.addStretch(5)

        return widget

    def createLevelPage(self, level):
        widget = QWidget()

        if level == 1:
            sheets = self.level1Sheets
        elif level == 2:
            sheets = self.level2Sheets
        elif level == 3:
            sheets = self.level3Sheets
        elif level == 4:
            sheets = self.level4Sheets

        vbox = QVBoxLayout(widget)

        if level == 1:
            levelLabel = QLabel(PAGE_1_NAME)
        elif level == 2:
            levelLabel = QLabel(PAGE_2_NAME)
        elif level == 3:
            levelLabel = QLabel(PAGE_3_NAME)
        elif level == 4:
            levelLabel = QLabel(PAGE_4_NAME)
        levelLabel.setStyleSheet("font-weight:bold; margin-left:0px; margin-top:10px; margin-bottom:10px")
        vbox.addWidget(levelLabel)

        if sheets is None:
            errorLabel = QLabel('本关卡题目配置读取失败，请确认路径是否正确！')
            vbox.addWidget(errorLabel)
            vbox.addStretch(2)
            return widget

        stackedWidget = QStackedWidget()
        sheethBox = QHBoxLayout()
        sheetIndex = 0
        for sheet_name, df in sheets.items():
            if not df.empty:
                sheetBtn = QPushButton(sheet_name)
                sheetBtn.clicked.connect(lambda _, i=sheetIndex: stackedWidget.setCurrentIndex(i))
                sheethBox.addWidget(sheetBtn)
                sheetIndex += 1

                model = PandasModel(df)
                table = QTableView()
                table.verticalHeader().hide()
                table.setModel(model)

                table.setItemDelegateForColumn(1, SpinBoxDelegate())

                for col in range(model.columnCount()):
                    header = model.headerData(col, Qt.Horizontal)
                    if header not in ('required', 'score', 'id', 'Question', 'Content'):
                        table.hideColumn(col)

                # 调整列宽
                table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
                table.showColumn(0)
                table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
                table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
                table.horizontalHeader().setStretchLastSection(True)
                stackedWidget.addWidget(table)
        sheethBox.addStretch(sheetIndex)
        vbox.addLayout(sheethBox)
        vbox.addWidget(stackedWidget)

        return widget

    def createResultPage(self):
        widget = QWidget()
        vbox = QVBoxLayout(widget)

        resultPageLabel = QLabel(PAGE_5_NAME)
        resultPageLabel.setStyleSheet("font-weight:bold; margin-left:0px; margin-top:10px; margin-bottom:10px")
        vbox.addWidget(resultPageLabel)

        model = QFileSystemModel()
        model.setFilter(QDir.Files | QDir.NoDotAndDotDot)
        model.setNameFilters(["*.json"])
        model.setNameFilterDisables(False)
        model.setRootPath(RESULT_FOLDER)

        def jsonFileDoubleClicked(index):
            path = model.filePath(index)
            if os.path.exists(path):
                dialog = ResultDetailDialog(path, self)
                dialog.exec_()

        table = QTableView()
        table.resizeRowsToContents()
        table.setModel(model)
        table.setRootIndex(model.index(RESULT_FOLDER))
        for col in [1, 2]:
            table.hideColumn(col)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        table.horizontalHeader().setStretchLastSection(True)
        table.horizontalHeader().hide()
        table.verticalHeader().hide()
        table.doubleClicked.connect(lambda index:jsonFileDoubleClicked(index))

        vbox.addWidget(table)

        return widget

    def initUi(self):
        self.toolbar = self.addToolBar('')
        self.toolbar.setMovable(False)
        self.toolbar.setContextMenuPolicy(Qt.PreventContextMenu)
        self.createAction('Environment', 0)
        self.toolbar.addSeparator()
        self.createAction('Level1', 1)
        self.createAction('Level2', 2)
        self.createAction('Level3', 3)
        self.createAction('Level4', 4)
        self.toolbar.addSeparator()
        self.createAction('Result',5)

        self.stacked = QStackedWidget()
        self.stacked.addWidget(self.createEnvPage())
        for i in (1, 2, 3, 4):
            self.stacked.addWidget(self.createLevelPage(i))
        self.stacked.addWidget(self.createResultPage())

        okBtn = QPushButton("OK")
        okBtn.clicked.connect(self.okButtonClicked)

        style = QApplication.style()
        defaultLeft = style.pixelMetric(QStyle.PM_LayoutLeftMargin)
        defaultTop = style.pixelMetric(QStyle.PM_LayoutTopMargin)
        defaultRight = style.pixelMetric(QStyle.PM_LayoutRightMargin)
        defaultBottom = style.pixelMetric(QStyle.PM_LayoutBottomMargin)

        central = QWidget()
        layout = QVBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.stacked)
        bottom = QHBoxLayout()
        bottom.setContentsMargins(defaultLeft, defaultTop, defaultRight, defaultBottom)
        bottom.addStretch(1)
        bottom.addWidget(okBtn)
        layout.addLayout(bottom)
        self.setCentralWidget(central)

    def okButtonClicked(self):
        time = self.timePicker.time().hour() * 100 + int(self.timePicker.time().minute() * 100 / 60)
        weather = self.weatherPicker.currentIndex()
        trainTime = self.trainTimeMinuteInput.value() * 60 + self.trainTimeSecondInput.value()
        name = self.nameInput.text()
        with open(BINARY_OUTPUT, 'wb') as binaryFile:
            packed = struct.pack('<iii', time, weather, trainTime)
            binaryFile.write(packed)
        binaryFile.close()
        with open(NAME_OUTPUT, 'w', encoding='UTF-8') as nameFile:
            nameFile.write(name)
        nameFile.close()

        with open(QUESTION1_OUTPUT, 'w', newline='') as csvFile:
            writer = csv.writer(csvFile)
            for level in (self.level1Sheets, self.level2Sheets, self.level3Sheets, self.level4Sheets):
                if level is None:
                    continue
                for df in level.values():
                    for index, row in df.iterrows():
                        if row['required']:
                            writer.writerow([str(row['id'])] + [str(row['score'])])
        csvFile.close()

        if os.path.exists(TRACKER_APPLICATION):
            subprocess.Popen(TRACKER_APPLICATION)
        if os.path.exists(UNREAL_APPLICATION):
            subprocess.Popen(UNREAL_APPLICATION)


class ResultDetailDialog(QDialog):
    def __init__(self, filePath, parent=None):
        super().__init__(parent)
        self.setWindowTitle("训练结果")
        self.resize(1200, 800)
        vbox = QVBoxLayout(self)
        tree = QTreeWidget()
        tree.setHeaderLabels(['题目','描述','用户答案','正确答案', '是否正确'])
        vbox.addWidget(tree)

        with open(filePath, 'r', encoding='UTF-16') as jsonFile:
            data = json.load(jsonFile)

        metadata = data.get('metadata', '')
        name = metadata.get('trainer', 'UnKnown')

        groups = {}
        for rec in data.get("answers", []):
            t = rec.get("type", "Unknown")
            groups.setdefault(t, []).append(rec)

        tree.clear()
        for type, items in groups.items():
            if type == 'SingleChoice':
                parent = QTreeWidgetItem(['单选题'])
            elif type == 'MultipleChoice':
                parent = QTreeWidgetItem(['多选题'])
            elif type == 'SceneTraining':
                parent = QTreeWidgetItem(['场景题'])
            elif type == 'TrueFalse':
                parent = QTreeWidgetItem(['判断题'])
            tree.addTopLevelItem(parent)
            for rec in items:
                id = rec.get('question_id', '')
                description = rec.get('description', '')
                userAnswer = rec.get('user_answer', '')
                correctAnswer = rec.get('correct_answer', rec.get('correct_answers', ''))
                questionContent = rec.get('question_content', None)

                isCorrect = False
                if type == "SingleChoice":
                    isCorrect = (userAnswer == correctAnswer)
                elif type == "MultipleChoice":
                    isCorrect = isinstance(userAnswer, list) and isinstance(correctAnswer, list) and set(
                        userAnswer) == set(correctAnswer)
                elif type == "SceneTraining":
                    isCorrect = (userAnswer == "True")
                elif type == "TrueFalse":
                    isCorrect = (userAnswer == correctAnswer)

                questionItem = QTreeWidgetItem([
                    id, description, str(userAnswer), str(correctAnswer),
                    "✓" if isCorrect else "✗"
                ])
                parent.addChild(questionItem)

                if isinstance(questionContent, list):
                    for index, option in enumerate(questionContent, 0):
                        optionItem = QTreeWidgetItem([f'选项{index}：', str(option)])
                        questionItem.addChild(optionItem)
                elif questionContent is not None:
                    questionItem.setText(1, f"描述：{description}\n内容：{questionContent}")
            parent.setExpanded(True)
        tree.header().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        tree.header().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        tree.header().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        tree.header().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        tree.header().setStretchLastSection(True)

        score, possible = self.computeScore(data.get('answers', []))
        footer = QHBoxLayout()
        scoreLabel = QLabel(f'姓名：{name}  得分：{score}/{possible}')
        footer.addWidget(scoreLabel)
        footer.addStretch(1)
        closeBtn = QPushButton('关闭')
        closeBtn.clicked.connect(lambda :self.close())
        footer.addWidget(closeBtn)

        vbox.addLayout(footer)

    def computeScore(self, answers):
        totalScore = 0  # 用户实际得分
        totalPossible = 0  # 所有题目分值之和

        for rec in answers:
            type = rec.get("type")
            weight = rec.get("score", 0)
            totalPossible += weight

            ua = rec.get("user_answer")
            ca = rec.get("correct_answer", rec.get("correct_answers"))

            earned = 0
            if type == "SingleChoice":
                if ua == ca:
                    earned = weight

            elif type == "MultipleChoice":
                if isinstance(ua, list) and isinstance(ca, list) and set(ua) == set(ca):
                    earned = weight

            elif type == "SceneTraining":
                if ua == "True":
                    earned = weight

            elif type == "TrueFalse":
                if ua == ca:
                    earned = weight

            totalScore += earned

        return totalScore, totalPossible


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.resize(1200, 800)
    window.show()
    sys.exit(app.exec_())
