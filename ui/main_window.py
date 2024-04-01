import sys
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QWidget, QPushButton, QApplication,
    QGridLayout, QLabel, QLineEdit, QVBoxLayout, QHBoxLayout, QSizePolicy,
    QCheckBox, QGroupBox, QRadioButton, QTextEdit, QPlainTextEdit,
    QCalendarWidget, QComboBox, QSpinBox, QButtonGroup, QProgressBar
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
        main_layout.setContentsMargins(20, 20, 20, 20)
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

    def fill_layout_logo(self):
        frame_logo = QtWidgets.QFrame()
        frame_logo.setFixedSize(190, 60)
        frame_logo.setFrameShape(QtWidgets.QFrame.Shape.StyledPanel)
        frame_logo.setFrameShadow(QtWidgets.QFrame.Shadow.Raised)
        frame_logo.setProperty('class', 'logo')
        self.layout_logo.addWidget(frame_logo)

    def fill_frame_culture(self):
        frame = self.get_frame()

        frame_layout = QVBoxLayout()

        label = QLabel('Культура')
        radio_sunflower = QRadioButton('Соняшник')
        radio_rapeseed = QRadioButton('Ріпак')
        radio_corn = QRadioButton('Кукурудза')

        frame_layout.addWidget(label)

        layout_culture_radio = QGridLayout()
        layout_culture_radio.addWidget(radio_sunflower, 0, 0)
        layout_culture_radio.addWidget(radio_rapeseed, 0, 1)
        layout_culture_radio.addWidget(radio_corn, 1, 0)

        frame_layout.addLayout(layout_culture_radio)

        frame.setLayout(frame_layout)

        self.layout_grid.addWidget(frame, 0, 0)

    def fill_frame_operations(self):
        frame = self.get_frame()
        frame_layout = QVBoxLayout()

        layout_operations = QHBoxLayout()

        title = QLabel('Дії з файлами')
        radio_new_file = QRadioButton('Новий файл')
        radio_exist_file = QRadioButton('Існуючий файл')
        button_operation = QPushButton('Обрати папку')
        label = QLabel('/home/maks_gv')

        frame_layout.addWidget(title)
        layout_operations.addWidget(radio_new_file)
        layout_operations.addWidget(radio_exist_file)
        frame_layout.addLayout(layout_operations)
        frame_layout.addWidget(button_operation)
        frame_layout.addWidget(label)

        frame.setLayout(frame_layout)
        self.layout_grid.addWidget(frame, 0, 1)

    def fill_input_button(self):
        frame = self.get_frame()
        frame_layout = QVBoxLayout()

        title = QLabel('Вхідні дані')
        button = QPushButton('Обрати')

        frame_layout.addWidget(title)
        frame_layout.addWidget(button)

        frame.setLayout(frame_layout)
        self.layout_grid.addWidget(frame, 1, 0)

    def fill_input_files(self):
        frame = self.get_frame()
        frame_layout = QVBoxLayout()

        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
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

        self.logs_edit.setReadOnly(True)
        self.logs_edit.setFont(self.font)
        self.logs_edit.setText('Файлів не вибрано')

        frame_layout.addWidget(self.logs_edit)

        frame.setLayout(frame_layout)
        self.layout_logs.addWidget(frame)


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setWindowTitle(f'Agro Flow v{VERSION}')
        MainWindow.setContentsMargins(20, 20, 20, 20)
        MainWindow.resize(810, 790)

        self.centralwidget = QtWidgets.QWidget(parent=MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.main_layout = QVBoxLayout()
        self.centralwidget.setLayout(self.main_layout)

        self.frame_6 = QtWidgets.QFrame()
        self.frame_6.setGeometry(QtCore.QRect(30, 10, 190, 60))
        self.frame_6.setStyleSheet("")
        self.frame_6.setFrameShape(QtWidgets.QFrame.Shape.StyledPanel)
        self.frame_6.setFrameShadow(QtWidgets.QFrame.Shadow.Raised)
        self.frame_6.setObjectName("frame_6")
        self.frame_6.setProperty('class', 'logo')

        self.main_layout.addWidget(self.frame_6)

        MainWindow.setCentralWidget(self.centralwidget)
