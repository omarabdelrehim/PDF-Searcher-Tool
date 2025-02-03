from PyQt5.QtWidgets import QApplication
from pdf_processor import process_pdfs_in_folder,process_pdfs_in_folder_without_OCR
from ui import PDFProcessorUI
import pandas as pd
from PyQt5.QtWidgets import QApplication, QFileDialog,QVBoxLayout, QPushButton, QLabel, QFileDialog, QTableWidget, QTableWidgetItem, QWidget,QHeaderView, QApplication, QWidget, QVBoxLayout, QLineEdit, QPushButton
import os
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtWidgets import QApplication, QMessageBox
import os
from PyPDF2 import PdfReader

class PDFProcessorApp(PDFProcessorUI):
    def __init__(self):
        super().__init__()
        self.folder_path = ""
        self.results = []

        # Connect buttons to actions
        self.select_folder_button.clicked.connect(self.select_folder)
        self.export_button.clicked.connect(self.export_results)
        self.checkbox.stateChanged.connect(self.checkbox_changed)
        
        

       

    def select_folder(self,others_word):
        """
        
        Open a folder selection dialog and process the selected folder.
    
        """
        
        others_word=self.Get_Other()
        Page_No=self.Get_Page_no()
        Page_end=self.Get_Page_end()
        checkbox=self.checked()
        
        folder_path = QFileDialog.getExistingDirectory(self, "Select Folder")
        if others_word == [""]:
                msg_box = QMessageBox()
                msg_box.setIcon(QMessageBox.Warning)
                msg_box.setWindowTitle("Input Error")
                msg_box.setText("You haven't entered a word to search for!")
                msg_box.setStandardButtons(QMessageBox.Ok)
                msg_box.exec_()  # Show the message box
                return
        
        if folder_path:
            self.folder_path = folder_path
            checkbox=self.checked()
            if checkbox==True:
               self.results = process_pdfs_in_folder(folder_path,others_word,Page_No,Page_end)
               self.update_results_table(self.results)
               self.enable_export_button(True)
               
            else :
                
                self.results = process_pdfs_in_folder_without_OCR(folder_path,others_word,Page_No,Page_end)
                self.update_results_table(self.results)
                self.enable_export_button(True)
                
            
            

    def export_results(self):
        """
        Export the processing results to an Excel file.
        """
        if self.results:
            file_path, _ = QFileDialog.getSaveFileName(self, "Save Results", "", "Excel Files (*.xlsx)")
            if file_path:
                df = pd.DataFrame(self.results)
                df.to_excel(file_path, index=False)
                self.label.setText("Results exported successfully!")


if __name__ == "__main__":
    app = QApplication([])
    window = PDFProcessorApp()
    window.show()
    app.exec_()
