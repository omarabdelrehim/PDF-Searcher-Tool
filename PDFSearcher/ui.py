from PyQt5.QtWidgets import ( 
    QVBoxLayout, QPushButton, QLabel, QFileDialog, QTableWidget, QTableWidgetItem, QWidget,QHeaderView,
    QApplication, QWidget, QVBoxLayout, QLineEdit, QPushButton,QHBoxLayout, QCheckBox
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIntValidator


class PDFProcessorUI(QWidget):
    """
    Class to define the user interface for the PDF Processor application.
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Search Tool")
        self.resize(600, 400)
        
        # Main layout
        self.layout = QVBoxLayout()
        self.Hlayout=QHBoxLayout()
        
        
        # Instruction Label
        self.label = QLabel("Step3: Select a folder containing PDF files to process  :")
        self.label2=QLabel(" Step 1  : Write the target words with a separator ',' ")
        self.Disclaimair  = QLabel("***Disclaimer: The accuracy of this application is contingent upon the positioning of target words within the PDF.")
        self.Input_words=QLineEdit(self)
        self.label3=QLabel("Step 2 : Start Search from pages ")
        self.Input_No=QLineEdit(self)
        self.Input_end=QLineEdit(self)
        self.label4=QLabel("to")
        validator = QIntValidator(0, 10000, self)
        self.Input_No.setValidator(validator)
        self.Input_end.setValidator(validator)
        self.checkbox = QCheckBox("Use OCR")
    
        
        self.layout.addWidget(self.label2)
        self.layout.addWidget(self.Input_words)
        self.Hlayout.addWidget(self.label3)
        self.Hlayout.addWidget(self.Input_No)
        self.Hlayout.addWidget(self.label4)
        self.Hlayout.addWidget(self.Input_end)
        self.Hlayout.addWidget(self.checkbox)
        
        self.layout.addLayout(self.Hlayout)
        self.layout.addWidget(self.label)
        
        
        
        # Select Folder Button
        self.select_folder_button = QPushButton("Select Folder")
        self.layout.addWidget(self.select_folder_button)
        
        # Results Table
        self.results_table = QTableWidget()
        otherword= self.Get_Other()
        
        
        
        basic=["File Name","Check"]
        basic.extend(otherword)
        self.results_table.setColumnCount(len(basic)-1)
     
        
        self.results_table.setHorizontalHeaderLabels(basic)
        self.layout.addWidget(self.results_table)
        header = self.results_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Stretch)
        
        self.results_table.setStyleSheet("""
            QTableWidget {
                border: 2px solid black;
                border-radius: 5px;
                gridline-color: black;
            }
            QTableWidget::item {
                border: 1px solid black;
            }
            QHeaderView::section {
                border: 1px solid black;
            }
        """)
        
        
        # Export Results Button
        self.export_button = QPushButton("Export Results to Excel")
        self.export_button.setEnabled(False)  # Initially disabled
        self.checkbox.setChecked(False)
        self.layout.addWidget(self.export_button)
        self.layout.addWidget(self.Disclaimair)
        # Set the layout
        self.setLayout(self.layout)

    def update_results_table(self, results):
        """
        Updates the results table with processing results.

        Args:
            results (list[dict]): List of dictionaries with file names and status.
            
        """
        
        otherword= self.Get_Other()
        
        basic=["File Name","Check"]
        basic.extend(otherword)
     
        self.results_table.setColumnCount(len(basic))
        self.results_table.setHorizontalHeaderLabels(basic)
        self.layout.addWidget(self.results_table)
        header = self.results_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Stretch)
        
        
        
        
        self.results_table.setRowCount(len(results))
        otherword=self.Get_Other()
        print(otherword)
        Page_No=self.Get_Page_no()
        for row, result in enumerate(results):
            self.results_table.setItem(row, 0, QTableWidgetItem(result["File Name"]))
            self.results_table.setItem(row, 1, QTableWidgetItem(result["Check"]))
            self.results_table.item(row, 0).setTextAlignment(Qt.AlignCenter)
            self.results_table.item(row, 1).setTextAlignment(Qt.AlignCenter)
            for i in otherword:
                x=otherword.index(i)
                col=x+2
                v=str(i)
                self.results_table.setItem(row,col, QTableWidgetItem(result[v]))
                self.results_table.item(row, col).setTextAlignment(Qt.AlignCenter)
                
                
        self.results_table.setStyleSheet("""
            QTableWidget {
                border: 2px solid black;
                border-radius: 5px;
                gridline-color: black;
            }
            QTableWidget::item {
                border: 1px solid black;
            }
            QHeaderView::section {
                border: 1px solid black;
            }
        """)
        self.layout.addWidget(self.Disclaimair)
    
            

    def enable_export_button(self, enabled=True):
        """
        Enables or disables the export button.

        Args:
            enabled (bool): True to enable, False to disable.
        """
        self.export_button.setEnabled(enabled)
        
    def Get_Other(self):
        text= self.Input_words.text()
        word_list=text.split(',')
        
        
        return word_list
    
    def Get_Page_no(self):
        if self.Input_No.text() =="":
            Page_No=0
        else: 
            Page_No=int(self.Input_No.text())-1
        return Page_No
    
    def Get_Page_end(self):
        if self.Input_end.text() =="":
            Page_end=100
        else: 
            Page_end=int(self.Input_end.text())+1
        return Page_end
    
    
        
    
    def checkbox_changed(self, state):
        
        if state == 2:  # 2 means checked (QCheckBox is checked)
            print("checked")
           
        else:  # 0 means unchecked
            print("unchecked")
        
        
    def checked(self):
        if self.checkbox.isChecked():
            Boxischecked=True
        else:
            Boxischecked=False
        return Boxischecked 
            
    
            
       
    
    
    
    
    
        
    

    