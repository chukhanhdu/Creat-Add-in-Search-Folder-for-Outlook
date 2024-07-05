import win32com.client
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLineEdit, QCompleter
from PyQt5.QtCore import QStringListModel, Qt

class CustomCompleter(QCompleter):
    def __init__(self, originalStringList=None, parent=None):
        super(CustomCompleter, self).__init__(parent)
        self.setCaseSensitivity(Qt.CaseInsensitive)
        self.setFilterMode(Qt.MatchContains)
        
        self.originalStringList = originalStringList if originalStringList is not None else []
        self.stringListModel = QStringListModel(self.originalStringList)
        self.setModel(self.stringListModel)

class Outlook365openfolder(QWidget):
    def __init__(self):
        super().__init__()
        self.outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
        self.initUI()

    def listfolder(self):
        inbox = self.outlook.GetDefaultFolder(6)  # 6 inbox
        folders = inbox.Folders
        listfolder = [folder.Name for folder in folders]
        return listfolder
    
    def openFolderByName(self, folderName):
        inbox = self.outlook.GetDefaultFolder(6)
        folders = inbox.Folders
        for folder in folders:
            if folder.Name == folderName:
                folder.Display()
                break
            
    def initUI(self):
        self.lineEdit = QLineEdit(self)
        keywords = self.listfolder()
        
        # Khởi tạo CustomCompleter với danh sách các folder
        completer = CustomCompleter(keywords, self)
        self.lineEdit.setCompleter(completer)
        
        # Kết nối sự kiện returnPressed của lineEdit để mở folder
        self.lineEdit.returnPressed.connect(self.onReturnPressed)

        layout = QVBoxLayout()
        layout.addWidget(self.lineEdit)
        self.setLayout(layout)
        self.setWindowTitle('Outlook365 Folder Search')
        self.setGeometry(100, 100, 300, 100)
        
    def onReturnPressed(self):
        folderName = self.lineEdit.text()
        self.openFolderByName(folderName)

