import sys
import InputConfigParser as ICF
import os
import json
from PyQt5.QtGui import QPixmap, QDesktopServices
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QWizard, QWizardPage, QPushButton, QLabel, QLineEdit, \
    QProgressBar, QHBoxLayout, QSizePolicy,QFileDialog
from PyQt5.QtCore import QTimer, Qt, QUrl
from PC_Campange_without_searchlogic import campagne_Ver_main
from PC_Campange_without_searchlogic import progUpdate


class Page1(QWizardPage):
    def __init__(self):
        super().__init__()
        self.setTitle('Enter User Info : ')
        layout = QHBoxLayout()

        # Left side image
        self.image_label = QLabel()
        self.image_label.setPixmap(QPixmap("Expleo_Logo.png"))  # Set your image path here
        self.image_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # Fill available space
        layout.addWidget(self.image_label)

        # Right side form
        form_layout = QVBoxLayout()

        self.username_label = QLabel('Enter the Docinfo username:')
        self.username_input = QLineEdit()
        form_layout.addWidget(self.username_label)
        form_layout.addWidget(self.username_input)

        self.password_label = QLabel('Enter the Docinfo password:')
        self.password_input = QLineEdit()
        form_layout.addWidget(self.password_label)
        form_layout.addWidget(self.password_input)

        hlayout1 = QHBoxLayout()
        self.TestPlanMacroFolder_label = QLabel('Enter the TestPlanMacro_folder:')
        self.TestPlanMacroFolder_input = QLineEdit()
        self.btn1 = QPushButton("File select")
        self.btn1.clicked.connect(lambda : self.getfile(0,self.TestPlanMacroFolder_input))
        hlayout1.addWidget(self.TestPlanMacroFolder_label)
        hlayout1.addWidget(self.TestPlanMacroFolder_input)
        hlayout1.addWidget(self.btn1)
        form_layout.addLayout(hlayout1)

        hlayout2 = QHBoxLayout()
        self.InputFolder_label = QLabel('Enter the InputFolder:')
        self.InputFolder_input = QLineEdit()
        self.btn2 = QPushButton("Folder select")
        self.btn2.clicked.connect(lambda: self.getfile(1,self.InputFolder_input))
        hlayout2.addWidget(self.InputFolder_label)
        hlayout2.addWidget(self.InputFolder_input)
        hlayout2.addWidget(self.btn2)
        form_layout.addLayout(hlayout2)

        hlayout3 = QHBoxLayout()
        self.OutputFolder_label = QLabel('Enter the OutputFolder:')
        self.OutputFolder_input = QLineEdit()
        self.btn3 = QPushButton("Folder select")
        self.btn3.clicked.connect(lambda: self.getfile(1,self.OutputFolder_input))
        hlayout3.addWidget(self.OutputFolder_label)
        hlayout3.addWidget(self.OutputFolder_input)
        hlayout3.addWidget(self.btn3)
        form_layout.addLayout(hlayout3)

        hlayout4 = QHBoxLayout()
        self.DownloadFolder_label = QLabel('Enter the DownloadFolder:')
        self.DownloadFolder_input = QLineEdit()
        self.btn4 = QPushButton("Folder select")
        self.btn4.clicked.connect(lambda: self.getfile(1, self.DownloadFolder_input))
        hlayout4.addWidget(self.DownloadFolder_label)
        hlayout4.addWidget(self.DownloadFolder_input)
        hlayout4.addWidget(self.btn4)
        form_layout.addLayout(hlayout4)

        self.name_label = QLabel('Enter the Name:')
        self.name_input = QLineEdit()
        form_layout.addWidget(self.name_label)
        form_layout.addWidget(self.name_input)

        self.start_date_label = QLabel('Enter the Start Date:')
        self.start_date_input = QLineEdit()
        form_layout.addWidget(self.start_date_label)
        form_layout.addWidget(self.start_date_input)

        self.task_ID_label = QLabel('Enter the task_ID :')
        self.task_ID_input = QLineEdit()
        form_layout.addWidget(self.task_ID_label)
        form_layout.addWidget(self.task_ID_input)

        self.reference_label = QLabel('Enter the GAELE Reference (separated by comma):')
        self.reference_input = QLineEdit()
        form_layout.addWidget(self.reference_label)
        form_layout.addWidget(self.reference_input)

        self.projects_label = QLabel('Enter the number of projects:')
        self.projects_input = QLineEdit()
        form_layout.addWidget(self.projects_label)
        form_layout.addWidget(self.projects_input)

        layout.addLayout(form_layout)
        self.setLayout(layout)

    def getfile(self,n,x):
       # dlg = QFileDialog.getOpenFileName(self, 'Open file',
       #                                     'c:\\')
        #dlg.setFileMode(QFileDialog.AnyFile)
        #dlg.setFilter("Text files (*.txt)")
        #if dlg.exec_():
        #filenames = dlg.selectedFiles()
       # self.TestPlanMacroFolder_input.setText(dlg[0])
        if n==0:
            dlg = QFileDialog()
            dlg.setFileMode(QFileDialog.AnyFile)
            # dlg.setFilter("Folder")

            if dlg.exec_():
                filenames = dlg.selectedFiles()
                x.setText(filenames[0])
        elif n==1:
            dlg = QFileDialog()
            dlg.setFileMode(QFileDialog.DirectoryOnly)
           # dlg.setFilter("Folder")

            if dlg.exec_():
                filenames = dlg.selectedFiles()
                x.setText(filenames[0])

    def initializePage(self):
        # if os.path.isfile('../user_input/UserInput.json'):
        #     with open('../user_input/UserInput.json', "r") as f_userInput:

                # userInput = json.load(f_userInput)

        self.username_input.setText(ICF.userInput['docInfo']['username'])
        self.password_input.setText(ICF.userInput['docInfo']['password'])
        self.TestPlanMacroFolder_input.setText(ICF.userInput['toolsPath']['TestPlanMacro'])
        self.InputFolder_input.setText(ICF.userInput['toolsPath']['InputFolder'])
        self.OutputFolder_input.setText(ICF.userInput['toolsPath']['OutputFolder'])
        self.DownloadFolder_input.setText(ICF.userInput['toolsPath']['DownloadFolder'])

        self.name_input.clear()
        self.start_date_input.clear()
        self.task_ID_input.clear()
        self.reference_input.clear()
        self.projects_input.clear()

    def get_data(self):
        return {
            'username': self.username_input.text(),
            'password': self.password_input.text(),
            'TestPlanMacro': self.TestPlanMacroFolder_input.text(),
            'InputFolder': self.InputFolder_input.text(),
            'OutputFolder': self.OutputFolder_input.text(),
            'DownloadFolder': self.DownloadFolder_input.text(),
            'name': self.name_input.text(),
            'start_date': self.start_date_input.text(),
            'task_ID': self.task_ID_input.text(),
            'reference': self.reference_input.text(),
            'projects_count': int(self.projects_input.text())
        }


class Page2(QWizardPage):
    def __init__(self):
        super().__init__()
        self.setTitle('Enter Project Details : ')
        self.layout = QVBoxLayout()
        self.setLayout(self.layout)
        self.project_widgets = []

    def initializePage(self):
        self.clearLayout(self.layout)
        projects_count = int(self.wizard().page(0).projects_input.text())
        self.project_widgets.clear()
        for i in range(projects_count):
            sheet_label = QLabel(f'Enter the desired sheet name for project {i + 1}:')
            sheet_input = QLineEdit()
            architecture_label = QLabel(f'Enter the desired architecture for project {i + 1}:')
            architecture_input = QLineEdit()
            self.layout.addWidget(sheet_label)
            self.layout.addWidget(sheet_input)
            self.layout.addWidget(architecture_label)
            self.layout.addWidget(architecture_input)
            self.project_widgets.append((sheet_input, architecture_input))

    def clearLayout(self, layout):
        while layout.count():
            child = layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

    def get_data(self):
        project_data = []
        for sheet_input, architecture_input in self.project_widgets:
            project_data.append({
                'sheet_name': sheet_input.text(),
                'architecture': architecture_input.text()
            })
        return project_data


class Page3(QWizardPage):
    def __init__(self):
        super().__init__()
        self.setTitle('Progress:')
        layout = QVBoxLayout()
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)

        # Link to open the output file
        self.output_link = QLabel(f'<a href="#">Open Output File</a>')
        self.output_link.linkActivated.connect(self.openOutputFile)
        self.output_link.hide()  # Hide the link initially
        layout.addWidget(self.output_link)

        self.setLayout(layout)

    def initializePage(self):
        self.timer = QTimer(self)
        self.timer.setInterval(100)
        self.timer.timeout.connect(self.updateProgress)
        self.progress = 0
        #self.timer.start()
        self.wizard().button(QWizard.FinishButton).setEnabled(False)  # Disable Finish button initially

    def updateProgress(self,progs):
        self.progress += progs
        if self.progress > 100:
            #self.timer.stop()
            self.progress_bar.setValue(100)
            self.completeProgress()
        else:
            self.progress_bar.setValue(self.progress)

    def completeProgress(self):
        self.setTitle('Progress Completed Successfully')
        self.wizard().button(QWizard.FinishButton).setEnabled(True)  # Enable Finish button
        self.output_link.show()  # Show the link when progress completes
        self.output_link.linkActivated.connect(self.openOutputFile)
        wizard = self.wizard()
        wizard.activateWindow()
        wizard.raise_()

    def openOutputFile(self):
        filepath = ICF.getOutputFiles() + "\\Output_document.docx" # Absolute path to output.docx file
        try:
            os.startfile(filepath)  # Open file with default associated program
        except Exception as e:
            print("Error opening file:", e)
        output_dir = ICF.getOutputFiles()
        try:
            # List all files in the output directory
            files = os.listdir(output_dir)

            # Filter Excel files
            excel_files = [file for file in files if file.endswith('.xlsm')]

            if excel_files:
                # Open the first Excel file found
                filepath = os.path.join(output_dir, excel_files[0])
                os.startfile(filepath)  # Open file with default associated program
            else:
                print("No Excel files found in the output directory.")
        except Exception as e:
            print("Error opening file:", e)


class Wizard(QWizard):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('VSM_PT')
        self.setWizardStyle(QWizard.WizardStyle.ClassicStyle)  # Set classic style for larger appearance
        self.addPage(Page1())
        self.addPage(Page2())
        self.addPage(Page3())
        self.button(QWizard.FinishButton).setEnabled(False)  # Disable Finish button initially
        self.button(QWizard.FinishButton).clicked.connect(self.finishClicked)
        self.currentIdChanged.connect(self.pageChanged)

    def finishClicked(self):
        print("Wizard Finished")

    def pageChanged(self, pageId):
        if pageId == 1:

            if os.path.isfile('../user_input/UserInput.json'):
                with open('../user_input/UserInput.json', "r") as f_userInput:

                    ICF.userInput = json.load(f_userInput)
            page1_data = self.page(0).get_data()

            ICF.userInput['docInfo']['username'] = page1_data['username']
            ICF.userInput['docInfo']['password'] = page1_data['password']

            ICF.userInput['toolsPath']['TestPlanMacro'] = page1_data['TestPlanMacro']
            ICF.userInput['toolsPath']['InputFolder'] = page1_data['InputFolder']
            ICF.userInput['toolsPath']['OutputFolder'] = page1_data['OutputFolder']
            ICF.userInput['toolsPath']['DownloadFolder'] = page1_data['DownloadFolder']


            os.remove('../user_input/UserInput.json')
            with open('../user_input/UserInput.json', 'w') as f:
                json.dump(ICF.userInput, f)
        if pageId == 2:  # Page 3 ID
            # Collect data from Page 1
            page1_data = self.page(0).get_data()
            # Collect data from Page 2
            page2_data = self.page(1).get_data()
            # Call campagne_Ver_main function with collected data

            progUpdate(self.page(2).updateProgress)

            try:
                campagne_Ver_main(
                    name=page1_data['name'],
                    start_date=page1_data['start_date'],
                    Task_ID=page1_data['task_ID'],
                    reference=page1_data['reference'],
                    projects=page1_data['projects_count'],
                    project_details=page2_data
                )
            except Exception as ex:
                print(ex)


def main():
    try:
        ICF.loadConfig()
        app = QApplication(sys.argv)
        wizard = Wizard()
        wizard.show()
        res = app.exec_()
    except Exception as e:
        print(e)


if __name__ == '__main__':
    main()
