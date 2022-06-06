import os
import hashlib

from PySide2 import QtWidgets, QtCore, QtGui, QtSvg
from package.api.file_handler import load_wb, get_col_names
from package.api.constants import DESKTOP_DIR, OUTPUT_SUFFIX


class Worker(QtCore.QObject):
    finished = QtCore.Signal()
    file_saved = QtCore.Signal(str, bool)
    row_hashed = QtCore.Signal(int, bool)
    on_error = QtCore.Signal(str)

    def __init__(self, filename, wb, columns_name, column_chosen):
        super().__init__()
        self.filename = filename
        self.wb = wb
        self.columns_name = columns_name
        self.column_chosen = column_chosen
        self.runs = True

    def hash_column(self):
        column = filter(lambda x: x['value'] == self.column_chosen, self.columns_name)
        column = list(column)
        column_letter = column[0]['letter']
        sheet = self.wb.active
        data = sheet[column_letter][1::]
        for i in range(len(data)):
            if self.runs:
                try:
                    row = data[i]
                    encoded_value = row.value.encode()
                    hashed_value = hashlib.md5(encoded_value).hexdigest()
                    row.value = hashed_value
                    self.row_hashed.emit(i + 1, True)
                except Exception as e:
                    self.finished.emit()
                    self.on_error.emit(str(e))
                    exit(1)
        path = os.path.join(DESKTOP_DIR, self.filename + OUTPUT_SUFFIX)
        self.wb.save(path)
        self.file_saved.emit(path, os.path.exists(path))
        self.finished.emit()


class MainWindow(QtWidgets.QWidget):
    def __init__(self, ctx):
        super().__init__()
        self.ctx = ctx
        self.app_name = QtCore.QCoreApplication.applicationName()
        self.app_version = QtCore.QCoreApplication.applicationVersion()
        self.back_icon = self.ctx.get_resource("back-icon.png")
        self.open_file_icon = self.ctx.get_resource("folder-icon.svg")
        css_file = self.ctx.get_resource("style.css")
        with open(css_file, "r") as f:
            self.setStyleSheet(f.read())

        self.create_main_layout()
        self.create_description_layout()
        self.create_application_box_label()
        #self.create_description_label()
        self.create_description_text()
        self.create_file_dialog_button()

    '''
    Hash column when user clicks on a column in the list and confirms the choice
    '''
    def process(self):
        max_rows = self.wb.active.max_row - 1                    # do not count header
        self.thread = QtCore.QThread(self)                       # create thread
        self.worker = Worker(filename=self.file_name, wb=self.wb, columns_name=self.col_names, column_chosen=self.chosen_item)
        self.worker.moveToThread(self.thread)                    # move worker to thread
        self.thread.started.connect(self.worker.hash_column)     # on started, call hash_column method
        self.worker.file_saved.connect(self.on_file_saved)       # connect signal to method on_file_saved (see below)
        self.worker.row_hashed.connect(self.on_row_hashed)       # connect signal to method on_row_hashed (see below)
        self.worker.finished.connect(self.thread.quit)           # connect finished signal to quit thread
        self.worker.on_error.connect(self.on_error)              # on error signal, show error message and exit, the process will not continue
        self.thread.start()                                      # start thread

        # if no error occurs, show progress bar
        self.progress_dialog = QtWidgets.QProgressDialog(
            labelText="Hashing column...",
            cancelButtonText="Cancel",
            minimum=0,
            maximum=max_rows,
            parent=self)
        self.progress_dialog.setAutoClose(False)
        self.progress_dialog.setAutoReset(False)
        self.progress_dialog.setWindowModality(QtCore.Qt.WindowModal) # make dialog modal to main window. This is needed to make it block other windows
        self.progress_dialog.canceled.connect(self.on_progress_dialog_canceled) # if user cancels, quit thread
        self.progress_dialog.show()

    '''
    when button is clicked, open file dialog and load file.
    Then add list widget with sheet column names
    '''
    def on_open_file_dialog_button_clicked(self):
        input_file_dialog = QtWidgets.QFileDialog()
        input_file_dialog.setDirectoryUrl(QtCore.QUrl.fromLocalFile(DESKTOP_DIR))
        input_file_dialog.setFileMode(QtWidgets.QFileDialog.ExistingFile)
        input_file_dialog.setNameFilter("Excel files (*.xlsx)")
        input_file_dialog.setViewMode(QtWidgets.QFileDialog.Detail)
        input_file_dialog.setLabelText(QtWidgets.QFileDialog.Accept, "Open")
        input_file_dialog.setLabelText(QtWidgets.QFileDialog.Reject, "Cancel")
        input_file_dialog.setWindowTitle("Open File")

        if input_file_dialog.exec_():         # if user clicked open
            files = input_file_dialog.selectedFiles()
            selected_file = files[0]          # get first file in list because only one file can be selected
            self.file_name = selected_file.split("/")[-1].split(".")[0]
            try:
                self.wb = load_wb(selected_file)
                self.ws = self.wb.active
                self.col_names = get_col_names(self.ws)
                self.create_list()
                self.btn_open_file_dialog.hide()   # hide button after first click
                self.create_back_button()          # add back button to load another file
            except Exception as e:
                self.on_error_loading_file(str(e)) # if error loading file, show error message

        else:
            return False

    '''
    if the column cannot be hashed, show error message and quit thread
    '''
    def on_error(self, message):
        self.progress_dialog.hide()
        self.progress_dialog.deleteLater()
        QtWidgets.QMessageBox.critical(self, "Error", message)
        self.worker.runs = False
        self.thread.quit()
        return False

    def on_error_loading_file(self, message):
        QtWidgets.QMessageBox.critical(self, "Error", message)
        return False

    '''
    display message box confirmation, if user clicks yes, hash column, if no, let user choose another column
    '''
    def on_list_item_clicked(self):
        self.chosen_item = self.list.currentItem().text()
        if self.chosen_item:
            self.confirmation_modal = QtWidgets.QMessageBox()
            self.confirmation_modal.setText("Are you sure you want to hash this column?")
            self.confirmation_modal.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
            self.confirmation_modal.setDefaultButton(QtWidgets.QMessageBox.No)
            self.confirmation_modal.setWindowTitle("Confirmation")
            ret = self.confirmation_modal.exec_()
            if ret == QtWidgets.QMessageBox.Yes:
                self.process()
            elif ret == QtWidgets.QMessageBox.No:
                return False

    def on_row_hashed(self, iteration):
        if iteration:
            self.progress_dialog.setValue(iteration)

    def on_file_saved(self, file_saved):
        if file_saved:
            self.progress_dialog.setLabelText("Column has been hashed, you can close this app.")

    def on_back_button_clicked(self):
        self.btn_open_file_dialog.show()
        self.list.hide()
        self.btn_back.hide()

    def on_progress_dialog_canceled(self):
        self.worker.runs = False
        self.thread.quit()


    '''
    create view >>>
    '''

    def create_main_layout(self):
        self.main_layout = QtWidgets.QVBoxLayout()
        self.setLayout(self.main_layout)

    def create_description_layout(self):
        self.description_layout = QtWidgets.QHBoxLayout()
        self.main_layout.addLayout(self.description_layout)

    def create_description_label(self):
        self.description_label = QtWidgets.QLabel(self.app_name)
        self.description_label.setObjectName("descriptionLabel")
        self.description_layout.addWidget(self.description_label)

    def create_application_box_label(self):
        self.box = QtWidgets.QGroupBox("version" + ' ' + self.app_version)
        self.box.setObjectName("box")
        self.app_title_label = QtWidgets.QLabel(self.app_name)
        self.app_title_label.setObjectName("appTitleLabel")
        self.app_title_label.setAlignment(QtCore.Qt.AlignCenter)
        self.app_title_label.setWordWrap(True)
        self.app_title_label.setMinimumWidth(300)
        self.app_title_label.setSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        self.box.setLayout(QtWidgets.QVBoxLayout())
        self.box.layout().addWidget(self.app_title_label)
        self.description_layout.addWidget(self.box)

    def create_description_text(self):
        self.description_text = QtWidgets.QLabel("This app hashes Excel columns.\n\n"
                                                 "> Select a column in the list and confirm the selection in the message box.\n\n"
                                                 "> The column will be hashed and a new file will be create.\n\n"
                                                 "> Et voilà!")
        self.description_text.setObjectName("descriptionText")
        self.description_layout.addWidget(self.description_text)

    def create_file_dialog_button(self):
        self.btn_open_file_dialog = QtWidgets.QPushButton("Open File")
        self.btn_open_file_dialog.setFixedWidth(250)
        self.btn_open_file_dialog.setMinimumWidth(250)
        self.btn_open_file_dialog.setObjectName("openFileButton")
        self.btn_open_file_dialog.setCursor(QtCore.Qt.PointingHandCursor)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(self.open_file_icon))
        self.btn_open_file_dialog.setIcon(icon)
        self.btn_open_file_dialog.setIconSize(QtCore.QSize(50, 50))
        self.btn_open_file_dialog.clicked.connect(self.on_open_file_dialog_button_clicked)
        self.main_layout.addWidget(self.btn_open_file_dialog, 1, QtCore.Qt.AlignCenter)

    '''
    add list widget with sheet column names
    '''
    def create_list(self):
        self.list = QtWidgets.QListWidget()
        for column in self.col_names:
            self.list.addItem(column['value'])
        self.layout().addWidget(self.list)
        self.list.itemClicked.connect(self.on_list_item_clicked)

    def create_back_button(self):
        self.btn_back = QtWidgets.QPushButton("Back")
        self.btn_back.clicked.connect(self.on_back_button_clicked)
        self.btn_back.setObjectName("backButton")
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(self.back_icon))
        self.btn_back.setIcon(icon)
        self.btn_back.setIconSize(QtCore.QSize(50, 50))
        self.btn_back.setCursor(QtCore.Qt.PointingHandCursor)
        self.layout().addWidget(self.btn_back)
