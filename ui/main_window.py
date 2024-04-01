import sys
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QWidget, QPushButton, QApplication,
    QGridLayout, QLabel, QLineEdit, QVBoxLayout, QHBoxLayout, QSizePolicy,
    QCheckBox, QGroupBox, QRadioButton, QTextEdit, QPlainTextEdit,
    QCalendarWidget, QComboBox, QSpinBox, QButtonGroup, QProgressBar, QSpacerItem,
    QFrame
)
from PyQt6 import QtWidgets, QtCore
from PyQt6.QtGui import QFont, QFontDatabase, QTextCursor, QPixmap

from datetime import datetime

from settings import VERSION


class AgroMainWindow(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle(f'Agro Flow v{VERSION}')
        self.setContentsMargins(20, 20, 20, 20)
        self.resize(810, 790)

        main_layout = QVBoxLayout()
        # main_layout.setContentsMargins(20, 20, 20, 20)
        self.setLayout(main_layout)

        self.layout_logo = QVBoxLayout()
        self.layout_grid = QGridLayout()
        self.layout_processing = QVBoxLayout()
        self.layout_progres = QVBoxLayout()
        self.layout_logs = QVBoxLayout()

        main_layout.addLayout(self.layout_logo)
        main_layout.addLayout(self.layout_grid)
        main_layout.addLayout(self.layout_processing)
        main_layout.addLayout(self.layout_progres)
        main_layout.addLayout(self.layout_logs)

        self.button_group_culture = QButtonGroup()
        self.button_group_file_operations = QButtonGroup()

        self.font = QFont()
        self.font.setPointSize(8)

        self.progressBar = QProgressBar()
        self.logs_edit = QTextEdit()

        self.spacer_item = QSpacerItem(
            20,
            40,
            QtWidgets.QSizePolicy.Policy.Minimum,
            QtWidgets.QSizePolicy.Policy.Expanding
        )

        self.spacer_item2 = QtWidgets.QSpacerItem(
            40,
            20,
            QtWidgets.QSizePolicy.Policy.Expanding,
            QtWidgets.QSizePolicy.Policy.Minimum
        )

        self.fill_layout_logo()
        self.fill_layout_grid()
        self.fill_layout_processing()
        self.fill_layout_progres()
        self.fill_layout_logs()

    def get_frame(self):
        frame = QtWidgets.QFrame()
        frame.setFrameShape(QtWidgets.QFrame.Shape.StyledPanel)
        frame.setFrameShadow(QtWidgets.QFrame.Shadow.Raised)
        return frame

    def get_frame_layout(self):
        frame_layout = QVBoxLayout()
        frame_layout.setContentsMargins(20, 10, 0, 0)
        return frame_layout

    def fill_layout_logo(self):
        frame_logo = QFrame()
        frame_logo.setFixedSize(190, 60)
        frame_logo.setFrameShape(QtWidgets.QFrame.Shape.StyledPanel)
        frame_logo.setFrameShadow(QtWidgets.QFrame.Shadow.Raised)
        frame_logo.setProperty('class', 'logo')
        self.layout_logo.addWidget(frame_logo)

    def fill_frame_culture(self):
        frame = self.get_frame()

        frame_layout = self.get_frame_layout()

        label = QLabel('Культура')
        label.setProperty('class', 'radio_label')
        radio_sunflower = QRadioButton('Соняшник')
        radio_sunflower.setChecked(True)
        radio_rapeseed = QRadioButton('Ріпак')
        radio_rapeseed.setProperty('class', 'last_radio')
        radio_corn = QRadioButton('Кукурудза')
        radio_corn.setProperty('class', 'radio_corn')

        culture_group = QButtonGroup()
        culture_group.addButton(radio_sunflower)
        culture_group.addButton(radio_rapeseed)
        culture_group.addButton(radio_corn)

        frame_layout.addWidget(label)

        layout_culture_radio = QGridLayout()
        layout_culture_radio.setContentsMargins(3, 0, 0, 0)
        layout_culture_radio.addWidget(radio_sunflower, 0, 0)
        layout_culture_radio.addWidget(radio_rapeseed, 0, 1)
        layout_culture_radio.addWidget(radio_corn, 1, 0)
        layout_culture_radio.addItem(self.spacer_item2)

        frame_layout.addLayout(layout_culture_radio)
        frame_layout.addItem(self.spacer_item)

        frame.setLayout(frame_layout)

        self.layout_grid.addWidget(frame, 0, 0)

    def fill_frame_operations(self):
        frame = self.get_frame()
        frame_layout = self.get_frame_layout()

        layout_operations = QHBoxLayout()
        layout_operations.setContentsMargins(3, 0, 0, 0)

        title = QLabel('Дії з файлами')
        title.setProperty('class', 'radio_label')
        radio_new_file = QRadioButton('Новий файл')
        radio_new_file.setChecked(True)
        radio_exist_file = QRadioButton('Існуючий файл')
        radio_exist_file.setProperty('class', 'last_radio')
        button_operation = QPushButton('Обрати папку')
        button_operation.setFixedSize(121, 35)
        button_operation.setProperty('class', 'button_operation')
        label = QLabel('/home/maks_gv')
        label.setFont(self.font)
        label.setProperty('class', 'label_file')

        operation_group = QButtonGroup()
        operation_group.addButton(radio_new_file)
        operation_group.addButton(radio_exist_file)

        frame_layout.addWidget(title)

        layout_operations.addWidget(radio_new_file)
        layout_operations.addWidget(radio_exist_file)
        layout_operations.addItem(self.spacer_item2)

        frame_layout.addLayout(layout_operations)
        frame_layout.addWidget(button_operation)
        frame_layout.addWidget(label)

        frame_layout.addItem(self.spacer_item)

        frame.setLayout(frame_layout)
        self.layout_grid.addWidget(frame, 0, 1)

    def fill_input_button(self):
        frame = self.get_frame()
        frame_layout = self.get_frame_layout()

        title = QLabel('Вхідні дані')
        button = QPushButton('Обрати')

        frame_layout.addWidget(title)
        frame_layout.addWidget(button)

        frame.setLayout(frame_layout)
        self.layout_grid.addWidget(frame, 1, 0)

    def fill_input_files(self):
        frame = self.get_frame()
        frame_layout = QVBoxLayout()
        frame_layout.setContentsMargins(0, 0, 0, 0)

        text_edit = QTextEdit()
        # text_edit.setReadOnly(True)
        text_edit.setFont(self.font)
        text_edit.setText('Файлів не вибрано')

        frame_layout.addWidget(text_edit)

        frame.setLayout(frame_layout)
        self.layout_grid.addWidget(frame, 1, 1)

    def fill_layout_grid(self):
        self.fill_frame_culture()
        self.fill_frame_operations()
        self.fill_input_button()
        self.fill_input_files()

        # Setting the same stretch ratio for both speakers
        self.layout_grid.setColumnStretch(0, 1)  # The first column
        self.layout_grid.setColumnStretch(1, 1)  # The second column

        self.layout_grid.setRowStretch(0, 4)
        self.layout_grid.setRowStretch(1, 3)

    def fill_layout_processing(self):
        button = QPushButton('Обробка даних')
        self.layout_processing.addWidget(button)

    def fill_layout_progres(self):
        self.progressBar.setProperty('value', 24)
        self.progressBar.setTextVisible(False)
        self.layout_progres.addWidget(self.progressBar)

    def fill_layout_logs(self):
        frame = self.get_frame()
        frame_layout = QVBoxLayout()
        frame_layout.setContentsMargins(0, 0, 0, 0)

        # self.logs_edit.setReadOnly(True)
        self.logs_edit.setFont(self.font)
        self.logs_edit.setText('Файлів не вибрано')

        frame_layout.addWidget(self.logs_edit)

        frame.setLayout(frame_layout)
        self.layout_logs.addWidget(frame)


# class Ui_MainWindow(object):
#     def setupUi(self, MainWindow):
#         MainWindow.setWindowTitle(f'Agro Flow v{VERSION}')
#         MainWindow.setContentsMargins(20, 20, 20, 20)
#         MainWindow.resize(810, 790)
#
#         self.centralwidget = QtWidgets.QWidget(parent=MainWindow)
#         self.centralwidget.setObjectName("centralwidget")
#
#         self.main_layout = QVBoxLayout()
#         self.centralwidget.setLayout(self.main_layout)
#
#         self.frame_6 = QtWidgets.QFrame()
#         self.frame_6.setGeometry(QtCore.QRect(30, 10, 190, 60))
#         self.frame_6.setStyleSheet("")
#         self.frame_6.setFrameShape(QtWidgets.QFrame.Shape.StyledPanel)
#         self.frame_6.setFrameShadow(QtWidgets.QFrame.Shadow.Raised)
#         self.frame_6.setObjectName("frame_6")
#         self.frame_6.setProperty('class', 'logo')
#
#         self.main_layout.addWidget(self.frame_6)
#
#         MainWindow.setCentralWidget(self.centralwidget)
