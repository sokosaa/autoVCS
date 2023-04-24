import sys, os
from PyQt5 import QtWidgets, QtGui
from PyQt5.QtWidgets import QApplication, QWidget, QGridLayout, QPushButton, QFileDialog, QLabel, QLineEdit, QTextEdit, QMainWindow, QSizePolicy, QMessageBox, QFrame, QDialog, QVBoxLayout
from PyQt5.QtCore import Qt, QSettings
from PyQt5.QtGui import QPainter, QPen, QColor, QIcon
from vcstools import text2folder, rename_images, create_ppt_function, PPTtoPDF, move2focus, RetractFromFocus, openDDS, enter_info


if __name__ == '__main__':


    # Set the program window parameters
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon('icon.ico'))
    window = QWidget()
    window.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
    grid = QGridLayout(window)
    grid.setRowStretch(1, 0)
    grid.setColumnStretch(1, 0)
    spacing = 5 # default is 5
    grid.setSpacing(spacing)
    grid.setContentsMargins(spacing,spacing,spacing,spacing)


    # Ask user for their working path
    working_path_label = QLabel(window)
    working_path_label.setText("Paste your working path:")
    grid.addWidget(working_path_label,0,0)

    working_path = QLineEdit(window)
    grid.addWidget(working_path,0,1)


    # Ask user for their part number(s)
    part_numbers_label = QLabel(window)
    part_numbers_label.setText('Part Number(s):')
    grid.addWidget(part_numbers_label,1,0)

    part_numbers = QTextEdit(window)
    grid.addWidget(part_numbers, 1, 1)


    # Make a button that creates folders. (one folder for each of part, all in their working directory)
    create_folders = QPushButton('Create Folders')
    create_folders.clicked.connect(lambda: text2folder(working_path,part_numbers))
    grid.addWidget(create_folders,2,1)


    # Remind the user to put or make sure images are in the folder(s)
    put_images_label = QLabel(window)
    put_images_label.setText("Put images in the folder(s).")
    grid.addWidget(put_images_label,3,1)


    # Ask user for a custom .pptx template
    custom_template_label = QLabel(window)
    custom_template_label.setText('Custom Template (optional):')
    grid.addWidget(custom_template_label,4,0)

    browse_button = QPushButton("Browse")
    full_file_path = ''
    file_name = ''
    feedback_label = QLabel("or drop a file here")
    feedback_label.setAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
    feedback_label.setAcceptDrops(True)
    feedback_label.setMinimumSize(200, 50)
    feedback_label.setFrameShape(QFrame.StyledPanel)
    feedback_label.setFrameShadow(QFrame.Raised)

    def handle_file_chooser():
        file_paths = QFileDialog.getOpenFileNames()[0]
        if file_paths:
            global full_file_path, file_name
            full_file_path = file_paths[0]
            file_name = os.path.basename(full_file_path)
            feedback_label.setText(f'Selected file: {file_name}')
            feedback_label.setFrameShape(QFrame.Box)

    def handle_file_drop(event):
        file_paths = [u.toLocalFile() for u in event.mimeData().urls()]
        if file_paths:
            global full_file_path, file_name
            full_file_path = file_paths[0]
            file_name = os.path.basename(full_file_path)
            feedback_label.setText(f'Selected file: {file_name}')
            feedback_label.setFrameShape(QFrame.Box)

    def handle_drag_enter(event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    feedback_label.dropEvent = handle_file_drop
    feedback_label.dragEnterEvent = handle_drag_enter
    browse_button.clicked.connect(handle_file_chooser)
    grid.addWidget(browse_button, 4, 1)
    grid.addWidget(feedback_label, 5, 1)


    # Set the default template file for if no custom template is provided by the user
    default_template_file = 'default_VCS_template.pptx'


    # Make a button to make .pptx file(s) from their images [and custom template]
    create_ppt = QPushButton('Create PPTX')
    create_ppt.clicked.connect(lambda: (rename_images(working_path,part_numbers),create_ppt_function(working_path,part_numbers,full_file_path or default_template_file)))
    grid.addWidget(create_ppt,6,1)


    # Make a button to make .pdf file(s) from their .pptx file(s)
    create_ppt = QPushButton('PPTX to PDF')
    create_ppt.clicked.connect(lambda: PPTtoPDF(working_path,part_numbers))
    grid.addWidget(create_ppt,7,1)


    # Ask user for the vendor code (can only use this program once with parts that all have the same vendor code)
    vendor_code_label = QLabel(window)
    vendor_code_label.setText("Vendor Code:")
    grid.addWidget(vendor_code_label,8,0)

    vendor_code = QLineEdit(window)
    grid.addWidget(vendor_code,8,1)


    # Make a button to move the .pdf file(s) to their respective focus folder(s)
    create_ppt = QPushButton('Move to Focus')
    create_ppt.clicked.connect(lambda: move2focus(working_path,part_numbers,vendor_code.text(),destinations_edit))
    grid.addWidget(create_ppt,9,1)


    # Show the user the file paths of where the files were moved
    # Create a QPlainTextEdit widget to display the destinations
    destinations_edit = QTextEdit()
    grid.addWidget(destinations_edit, 10, 0, 1, 2)



    # Initalize program window
    window.setLayout(grid)
    window.show()
    sys.exit(app.exec_())