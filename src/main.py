import sys
from PyQt5.QtWidgets import (QApplication, QWidget, QGridLayout, QPushButton, QFileDialog, QLabel, 
                             QProgressBar, QSizePolicy)
from PyQt5.QtCore import Qt, QUrl, QThread, pyqtSignal
from PyQt5.QtGui import QDragEnterEvent, QDropEvent
import logging
from processing import generar_archivo_combinado

# Initialize the logger
logging.basicConfig(filename='app.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class FileProcessorThread(QThread):
    update_progress = pyqtSignal(int)
    
    def __init__(self, contactos_file, egresados_file, output_file):
        super().__init__()
        self.contactos_file = contactos_file
        self.egresados_file = egresados_file
        self.output_file = output_file

    def run(self):
        # Simulate processing and update the progress bar (example logic, adapt as needed)
        try:
            generar_archivo_combinado(self.contactos_file, self.egresados_file, self.output_file, self.update_progress)
        except Exception as e:
            logging.error("An error occurred while processing files.", exc_info=True)

class FileProcessorUI(QWidget):
    def __init__(self):
        super().__init__()
        self.contactos_file = None
        self.egresados_file = None
        self.output_file = None
        self.initUI()

    def initUI(self):
        layout = QGridLayout()

        # Labels
        self.contactos_label = QLabel("Contactos File:")
        layout.addWidget(self.contactos_label, 0, 0)

        self.egresados_label = QLabel("Egresados File:")
        layout.addWidget(self.egresados_label, 1, 0)

        self.output_label = QLabel("Output File:")
        layout.addWidget(self.output_label, 2, 0)

        # Buttons
        self.contactos_button = QPushButton("Browse Contactos")
        self.contactos_button.clicked.connect(self.select_contactos_file)
        layout.addWidget(self.contactos_button, 0, 1)

        self.egresados_button = QPushButton("Browse Egresados")
        self.egresados_button.clicked.connect(self.select_egresados_file)
        layout.addWidget(self.egresados_button, 1, 1)

        self.output_button = QPushButton("Select Output File")
        self.output_button.clicked.connect(self.select_output_file)
        layout.addWidget(self.output_button, 2, 1)

        # Drag and Drop labels
        self.contactos_drag_label = QLabel("or Drag and Drop here")
        self.contactos_drag_label.setAcceptDrops(True)
        self.contactos_drag_label.setAlignment(Qt.AlignCenter)
        self.contactos_drag_label.setStyleSheet("border: 2px dashed #aaa; padding: 10px;")
        self.contactos_drag_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.contactos_drag_label.dragEnterEvent = self.dragEnterEvent
        self.contactos_drag_label.dropEvent = self.dropEventContactos
        layout.addWidget(self.contactos_drag_label, 0, 2, 1, 2)

        self.egresados_drag_label = QLabel("or Drag and Drop here")
        self.egresados_drag_label.setAcceptDrops(True)
        self.egresados_drag_label.setAlignment(Qt.AlignCenter)
        self.egresados_drag_label.setStyleSheet("border: 2px dashed #aaa; padding: 10px;")
        self.egresados_drag_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.egresados_drag_label.dragEnterEvent = self.dragEnterEvent
        self.egresados_drag_label.dropEvent = self.dropEventEgresados
        layout.addWidget(self.egresados_drag_label, 1, 2, 1, 2)

        # Process button
        self.process_button = QPushButton("Process Files")
        self.process_button.clicked.connect(self.process_files)
        layout.addWidget(self.process_button, 3, 1, 1, 2)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        layout.addWidget(self.progress_bar, 4, 0, 1, 4)

        # Status label
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label, 5, 0, 1, 4)

        self.setLayout(layout)
        self.setWindowTitle('Excel File Processor')
        self.setGeometry(100, 100, 600, 300)  # Set window size

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEventContactos(self, event: QDropEvent):
        urls = event.mimeData().urls()
        if urls:
            self.contactos_file = urls[0].toLocalFile()
            self.contactos_drag_label.setText(f"Selected: {self.contactos_file}")

    def dropEventEgresados(self, event: QDropEvent):
        urls = event.mimeData().urls()
        if urls:
            self.egresados_file = urls[0].toLocalFile()
            self.egresados_drag_label.setText(f"Selected: {self.egresados_file}")

    def select_contactos_file(self):
        contactos_file, _ = QFileDialog.getOpenFileName(self, "Select Contactos File", "", "Excel Files (*.xlsx)")
        if contactos_file:
            self.contactos_drag_label.setText(f"Selected: {contactos_file}")
            self.contactos_file = contactos_file

    def select_egresados_file(self):
        egresados_file, _ = QFileDialog.getOpenFileName(self, "Select Egresados File", "", "Excel Files (*.xlsx)")
        if egresados_file:
            self.egresados_drag_label.setText(f"Selected: {egresados_file}")
            self.egresados_file = egresados_file

    def select_output_file(self):
        output_file, _ = QFileDialog.getSaveFileName(self, "Select Output File", "", "Excel Files (*.xlsx)")
        if output_file:
            if not output_file.endswith('.xlsx'):
                output_file += '.xlsx'
            self.output_label.setText(f"Selected: {output_file}")
            self.output_file = output_file

    def process_files(self):
        try:
            if not (self.contactos_file and self.egresados_file and self.output_file):
                self.status_label.setText("Please select all files.")
                self.status_label.setStyleSheet("color: red;")
                return
            
            self.status_label.setText("Processing...")
            self.status_label.setStyleSheet("color: orange;")
            self.progress_bar.setValue(0)

            # Run the processing in a separate thread
            self.thread = FileProcessorThread(self.contactos_file, self.egresados_file, self.output_file)
            self.thread.update_progress.connect(self.progress_bar.setValue)
            self.thread.finished.connect(self.on_processing_finished)
            self.thread.start()

        except Exception as e:
            logging.error("An error occurred while processing files.", exc_info=True)
            self.status_label.setText(f"Error: {str(e)}")
            self.status_label.setAlignment(Qt.AlignCenter)
            self.status_label.setStyleSheet("color: red;")

    def on_processing_finished(self):
        self.status_label.setText("Processing completed successfully!")
        self.status_label.setStyleSheet("color: green;")

def main():
    app = QApplication(sys.argv)
    ex = FileProcessorUI()
    ex.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
