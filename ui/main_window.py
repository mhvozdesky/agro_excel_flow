import sys
import os.path

import resources_rc

from enum import Enum

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QWidget, QPushButton, QGridLayout, QLabel, QVBoxLayout, QHBoxLayout,
    QSizePolicy, QRadioButton, QTextEdit, QButtonGroup, QProgressBar,
    QSpacerItem, QFrame
)
from PyQt6 import QtWidgets, QtGui, QtCore
from PyQt6.QtGui import QFont
from PyQt6.QtCore import QThread, pyqtSignal
from PyQt6.sip import wrappertype

from settings import VERSION


NO_FOLDER_SELECTED = 'Папку не вибрано'
NO_FILE_SELECTED = 'Файл не вибрано'
SELECT_DIR = 'Обрати папку'
SELECT_FILE = 'Обрати файл'
TEXT_BUTTON_DATA_PROCESSING = 'Обробка даних'
UPLOAD_FILES = 'Зачекайте завантаження файлів'

STD_OUT = sys.stdout


class CropsRadio(Enum):
    sunflower = 'Соняшник'
    rapeseed = 'Ріпак'
    corn = 'Кукурудза'


class OperationsRadio(Enum):
    new_file = 'Новий файл'
    exist_file = 'Існуючий файл'


class SingletonMeta(type):
    _instances = {}

    def __call__(cls, *args, **kwargs):
        if cls not in cls._instances:
            cls._instances[cls] = super(SingletonMeta, cls).__call__(*args, **kwargs)
        return cls._instances[cls]


class SingletonQThreadMeta(wrappertype):
    _instances = {}

    def __call__(cls, *args, **kwargs):
        if cls not in cls._instances:
            cls._instances[cls] = super(SingletonQThreadMeta, cls).__call__(*args, **kwargs)
        return cls._instances[cls]


class StreamProxy:
    def __init__(self, target_instance, type_msg):
        self.target_instance = target_instance
        self.type_msg = type_msg

    def write(self, message):
        self.target_instance.send_msg(message, self.type_msg)

    def flush(self):
        pass


class DataProcessor(QThread, metaclass=SingletonQThreadMeta):
    progress_updated = pyqtSignal(int)
    msg_updated = pyqtSignal(str, str)
    processing_complete = pyqtSignal(bool)
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.input_file_list = None
        self.dir_path = None
        self.file_path = None
        self.crops = None
        self.operation = None

        sys.stdout = StreamProxy(self, 'normal')
        sys.stderr = StreamProxy(self, 'error')

    def run(self):
        from main import main

        try:
            main(
                file_list=self.input_file_list,
                file_path=self.file_path,
                dir_path=self.dir_path,
                crops=self.crops
            )

            self.send_msg('Обробка завершена', type_msg='success')
        except Exception as e:
            self.send_msg(str(e), 'error')
        finally:
            self.send_processing_complete()

    def send_msg(self, value, type_msg):
        self.msg_updated.emit(type_msg, value)

    def send_processing_complete(self):
        self.processing_complete.emit(True)

    def update_progress(self, value):
        self.progress_updated.emit(value)


class StreamRedirector:
    def __init__(self, text_widget, color=QtGui.QColor(20, 51, 43)):
        self.text_widget = text_widget
        self.color = color

    def write(self, message):
        self.display_text(message, self.color)

    def display_text(self, message, color):
        self.text_widget.setTextColor(color)
        self.text_widget.insertPlainText(message)
        self.text_widget.verticalScrollBar().setValue(
            self.text_widget.verticalScrollBar().maximum()
        )

    def write_special_text(self, message, color):
        self.display_text(message, color)

    def flush(self):
        pass


class AgroMainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.progress_top_margin = 30

        self.setWindowTitle(f'Agro Flow v{VERSION}')
        self.setContentsMargins(20, 0, 20, 20)
        self.resize(810, 614)
        self.setWindowIcon(QtGui.QIcon(':/ui/images/icon3.ico'))

        main_layout = QVBoxLayout()
        self.setLayout(main_layout)

        self.layout_logo = QVBoxLayout()
        self.layout_grid = QGridLayout()
        self.layout_processing = QHBoxLayout()
        self.layout_progres = QVBoxLayout()
        self.layout_logs = QVBoxLayout()

        main_layout.addLayout(self.layout_logo)
        main_layout.addLayout(self.layout_grid)
        main_layout.addLayout(self.layout_processing)
        main_layout.addLayout(self.layout_progres)
        main_layout.addLayout(self.layout_logs)

        self.button_group_crops = QButtonGroup()
        self.button_group_file_operations = QButtonGroup()
        self.button_operation = QPushButton('Обрати папку')
        self.crops_group = QButtonGroup()
        self.operation_group = QButtonGroup()
        self.label_dir_file = QLabel('Папку не вибрано')
        self.input_files_text_edit = QTextEdit()
        self.processing_button = QPushButton(TEXT_BUTTON_DATA_PROCESSING)

        self.font = QFont()
        self.font.setPointSize(8)

        self.progressBar = QProgressBar()
        self.logs_edit = QTextEdit()

        self.fill_layout_logo()
        self.fill_layout_grid()
        self.fill_layout_processing()
        self.fill_layout_progres()
        self.fill_layout_logs()

        self.input_file_list = None
        self.dir_path = None
        self.file_path = None
        self.crops = None
        self.operation = None

        self.set_default_params()

        self.logs_stream = StreamRedirector(self.logs_edit)

        self.data_processor = DataProcessor()
        self.init_data_processor()

    def get_spacer_item_v(self):
        spacer_item_v = QSpacerItem(
            20,
            40,
            QSizePolicy.Policy.Minimum,
            QSizePolicy.Policy.Expanding
        )
        return spacer_item_v

    def get_spacer_item_h(self):
        spacer_item_h = QtWidgets.QSpacerItem(
            40,
            20,
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Minimum
        )
        return spacer_item_h

    def init_data_processor(self):
        self.data_processor.progress_updated.connect(self.update_progress)
        self.data_processor.msg_updated.connect(self.send_text)
        self.data_processor.processing_complete.connect(self.on_processing_complete)

    def get_radio_group_value(self, radio_group):
        checked_button = radio_group.checkedButton()
        if checked_button is not None:
            return checked_button.property('name')
        return None

    def set_default_params(self):
        self.input_file_list = []
        self.crops = self.get_radio_group_value(self.crops_group)
        self.operation = self.get_radio_group_value(self.operation_group)

    def resizeEvent(self, event):
        super().resizeEvent(event)

        if self.height() <= 650:
            delta = 650 - self.height()
            new_value = max(self.progress_top_margin - delta, 0)
            self.layout_progres.setContentsMargins(0, new_value, 0, 0)
        else:
            self.layout_progres.setContentsMargins(0, self.progress_top_margin, 0, 0)

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

    def fill_frame_crops(self):
        frame = self.get_frame()
        frame.setFixedHeight(151)

        frame_layout = self.get_frame_layout()

        label = QLabel('Культура')
        label.setProperty('class', 'radio_label')

        radio_sunflower = QRadioButton(CropsRadio.sunflower.value)
        radio_sunflower.setProperty('name', CropsRadio.sunflower.name)
        radio_sunflower.setChecked(True)

        radio_rapeseed = QRadioButton(CropsRadio.rapeseed.value)
        radio_rapeseed.setProperty('name', CropsRadio.rapeseed.name)
        radio_rapeseed.setProperty('class', 'last_radio')

        radio_corn = QRadioButton(CropsRadio.corn.value)
        radio_corn.setProperty('name', CropsRadio.corn.name)
        radio_corn.setProperty('class', 'radio_corn')
        radio_corn.setDisabled(True)

        self.crops_group.addButton(radio_sunflower)
        self.crops_group.addButton(radio_rapeseed)
        self.crops_group.addButton(radio_corn)

        frame_layout.addWidget(label)

        layout_crops_radio = QGridLayout()
        layout_crops_radio.setContentsMargins(3, 0, 0, 0)
        layout_crops_radio.addWidget(radio_sunflower, 0, 0)
        layout_crops_radio.addWidget(radio_rapeseed, 0, 1)
        layout_crops_radio.addWidget(radio_corn, 1, 0)
        layout_crops_radio.addItem(self.get_spacer_item_h())

        frame_layout.addLayout(layout_crops_radio)
        frame_layout.addItem(self.get_spacer_item_v())

        frame.setLayout(frame_layout)

        self.layout_grid.addWidget(frame, 0, 0)

        self.crops_group.buttonClicked.connect(self.on_crops_changed)

    def on_crops_changed(self, button):
        self.crops = button.property('name')

    def fill_frame_operations(self):
        frame = self.get_frame()
        frame.setFixedHeight(151)
        frame_layout = self.get_frame_layout()

        layout_operations = QHBoxLayout()
        layout_operations.setContentsMargins(3, 0, 0, 0)

        title = QLabel('Дії з файлами')
        title.setProperty('class', 'radio_label')

        radio_new_file = QRadioButton(OperationsRadio.new_file.value)
        radio_new_file.setChecked(True)
        radio_new_file.setProperty('name', OperationsRadio.new_file.name)

        radio_exist_file = QRadioButton(OperationsRadio.exist_file.value)
        radio_exist_file.setProperty('class', 'last_radio')
        radio_exist_file.setProperty('name', OperationsRadio.exist_file.name)

        self.button_operation.setFixedSize(121, 43)
        self.button_operation.setProperty('class', 'button_operation')

        self.label_dir_file.setFont(self.font)
        self.label_dir_file.setProperty('class', 'label_file')
        self.label_dir_file.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Expanding,
            QtWidgets.QSizePolicy.Policy.Preferred
        )
        self.label_dir_file.setStyleSheet("font-size: 12px;")

        self.operation_group.addButton(radio_new_file)
        self.operation_group.addButton(radio_exist_file)

        frame_layout.addWidget(title)

        layout_operations.addWidget(radio_new_file)
        layout_operations.addWidget(radio_exist_file)
        layout_operations.addItem(self.get_spacer_item_h())

        frame_layout.addLayout(layout_operations)
        frame_layout.addWidget(self.button_operation)
        frame_layout.addWidget(self.label_dir_file)

        frame_layout.addItem(self.get_spacer_item_v())

        frame.setLayout(frame_layout)
        self.layout_grid.addWidget(frame, 0, 1)

        self.operation_group.buttonClicked.connect(self.on_operation_changed)
        self.button_operation.clicked.connect(self.select_dir_or_file)

    def set_operation_path(self, value, operation):
        if operation == OperationsRadio.new_file.name:
            self.dir_path = value if value != NO_FOLDER_SELECTED else None
            self.file_path = None
        elif operation == OperationsRadio.exist_file.name:
            self.file_path = value if value != NO_FILE_SELECTED else None
            self.dir_path = None
        self.set_elided_folder_path(value)

    def on_operation_changed(self, button):
        self.operation = button.property('name')
        if self.operation == OperationsRadio.exist_file.name:
            self.set_operation_path(NO_FILE_SELECTED, self.operation)
            self.button_operation.setText(SELECT_FILE)
        elif self.operation == OperationsRadio.new_file.name:
            self.set_operation_path(NO_FOLDER_SELECTED, self.operation)
            self.button_operation.setText(SELECT_DIR)

    def set_elided_folder_path(self, full_path):
        # get the width of the label
        label_width = self.label_dir_file.width() - 5

        # create a QFontMetrics object with a label font
        font_metrics = QtGui.QFontMetrics(self.label_dir_file.font())

        # cut the text so that it fits in the label,
        # keeping the final part of the path
        elided_text = font_metrics.elidedText(
            full_path,
            QtCore.Qt.TextElideMode.ElideLeft,
            label_width
        )

        # set the truncated text for the label
        self.label_dir_file.setText(elided_text)

    def select_dir_or_file(self):
        operation = self.get_radio_group_value(self.operation_group)

        if operation == OperationsRadio.new_file.name:
            folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Обрати папку")
            if folder:
                self.set_operation_path(folder, operation)
        elif operation == OperationsRadio.exist_file.name:
            file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
                self,
                'Відкрити файл',
                '',
                'Excel Files (*.xlsx *.xls);;All Files (*)'
            )
            if file_path:
                self.set_operation_path(file_path, operation)

    def fill_input_button(self):
        frame = self.get_frame()
        frame.setFixedHeight(101)
        frame_layout = self.get_frame_layout()

        title = QLabel('Вхідні дані')
        button = QPushButton('Обрати')
        button.setFixedSize(89, 43)
        button.setProperty('class', 'button_input')

        spacer_item_v = self.get_spacer_item_v()

        frame_layout.addWidget(title)
        frame_layout.addWidget(button)
        frame_layout.addItem(spacer_item_v)

        frame.setLayout(frame_layout)

        self.layout_grid.addWidget(frame, 1, 0)

        button.clicked.connect(self.select_input_files)

    def display_input_files(self, file_paths):
        self.input_files_text_edit.setText('')

        name_list = [os.path.split(file_path)[1] for file_path in file_paths]
        name_for_display = '\n'.join(name_list)
        self.input_files_text_edit.setPlainText(name_for_display)

    def select_input_files(self):
        self.processing_button.setDisabled(True)
        self.processing_button.setFixedSize(251, 41)
        self.processing_button.setText(UPLOAD_FILES)

        file_paths, _ = QtWidgets.QFileDialog.getOpenFileNames(
            self,
            'Відкрити файли',
            '',
            'Excel Files (*.xlsx *.xls);;All Files (*)'
        )
        if file_paths:
            self.input_file_list = file_paths
            self.display_input_files(file_paths)
        self.processing_button.setDisabled(False)
        self.processing_button.setText(TEXT_BUTTON_DATA_PROCESSING)
        self.processing_button.setFixedSize(151, 41)

    def fill_input_files(self):
        frame = self.get_frame()
        frame.setFixedHeight(101)
        frame_layout = QVBoxLayout()
        frame_layout.setContentsMargins(0, 0, 0, 0)

        self.input_files_text_edit.setReadOnly(True)
        self.input_files_text_edit.setFont(self.font)
        self.input_files_text_edit.setText('Файлів не вибрано')
        self.input_files_text_edit.setStyleSheet("font-size: 12px;")

        frame_layout.addWidget(self.input_files_text_edit)

        frame.setLayout(frame_layout)
        self.layout_grid.addWidget(frame, 1, 1)

    def fill_layout_grid(self):
        self.fill_frame_crops()
        self.fill_frame_operations()
        self.fill_input_button()
        self.fill_input_files()

        # Setting the same stretch ratio for both speakers
        self.layout_grid.setColumnStretch(0, 1)  # The first column
        self.layout_grid.setColumnStretch(1, 1)  # The second column

        self.layout_grid.setHorizontalSpacing(19)
        self.layout_grid.setVerticalSpacing(19)

    def fill_layout_processing(self):
        self.processing_button.setFixedSize(151, 41)

        self.layout_processing.addWidget(
            self.processing_button, alignment=Qt.AlignmentFlag.AlignRight
        )

        self.layout_processing.setContentsMargins(0, 10, 0, 0)

        self.processing_button.clicked.connect(self.processing_data)

    def processing_data(self):
        self.logs_edit.clear()
        self.progressBar.setValue(0)

        self.data_processor.input_file_list = self.input_file_list
        self.data_processor.dir_path = self.dir_path
        self.data_processor.file_path = self.file_path
        self.data_processor.crops = self.crops
        self.data_processor.operation = self.operation

        try:
            self.processing_button.setDisabled(True)
            self.data_processor.start()

        except Exception as e:
            error_message = str(e)
            self.logs_stream.write_special_text(
                error_message + '\n',
                QtGui.QColor(209, 30, 30)
            )

    def on_processing_complete(self):
        self.processing_button.setDisabled(False)

    def send_text(self, type_msg, value):
        if type_msg == 'normal':
            self.logs_stream.write(value)
        elif type_msg == 'error':
            self.logs_stream.write_special_text(value, QtGui.QColor(209, 30, 30))
        elif type_msg == 'success':
            self.logs_stream.write_special_text(value, QtGui.QColor(28, 119, 39))

    def update_progress(self, value):
        self.progressBar.setValue(value)

    def fill_layout_progres(self):
        self.progressBar.setProperty('value', 0)
        self.progressBar.setProperty('class', 'progress_bar')
        self.progressBar.setTextVisible(False)
        self.layout_progres.addWidget(self.progressBar)

        self.layout_progres.setContentsMargins(0, 30, 0, 0)

    def fill_layout_logs(self):
        frame = self.get_frame()
        frame_layout = QVBoxLayout()
        frame_layout.setContentsMargins(0, 0, 0, 0)

        self.logs_edit.setReadOnly(True)
        self.logs_edit.setFont(self.font)
        self.logs_edit.setStyleSheet("font-size: 12px;")

        frame_layout.addWidget(self.logs_edit)

        frame.setLayout(frame_layout)
        self.layout_logs.addWidget(frame)
