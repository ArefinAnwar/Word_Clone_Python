import sys
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtPrintSupport import *
from PyQt5 import QtCore, QtGui
from PyQt5.QtGui import * 
from PyQt5.QtCore import * 
from PyQt5.QtWidgets import QMessageBox

class Word(QMainWindow):
    def __init__(self):
        super(Word, self).__init__()
        
        self.editor = QTextEdit()
        self.editor.setFontPointSize(20)  #Font Size
        
        
        
        self.setCentralWidget(self.editor)
        
        self.setMinimumWidth(1000)  #Width
        self.setMinimumHeight(600)  #Height
        
        self.title = "Word"
        self.setWindowTitle(self.title)  #Title

        self.font_size_box = QSpinBox()
        
        
        
        self.menu_bar()
        
        self.tool_bar()

        self.path = ''
    
    def menu_bar(self):
        menu_bar = QMenuBar()
        
        file_menu = QMenu('File', self)
        menu_bar.addMenu(file_menu)

        save_action = QAction(QIcon('save.png'), 'Save', self)
        save_action.triggered.connect(self.save_file)
        file_menu.addAction(save_action)

        open_action = QAction(QIcon('045-open file.png'), 'Open', self)
        open_action.triggered.connect(self.open_file)
        file_menu.addAction(open_action)

        save_as_pdf = QAction(QIcon('pdf.png'), 'Export PDF', self)
        save_as_pdf.triggered.connect(self.export_pdf)
        file_menu.addAction(save_as_pdf)
        



        edit_menu = QMenu('Edit', self)
        menu_bar.addMenu(edit_menu)

        undo_action = QAction(QIcon("undo.png"), 'Undo', self)
        undo_action.triggered.connect(self.editor.undo)
        edit_menu.addAction(undo_action)
        
        redo_action = QAction(QIcon("redo.png"), 'Redo', self)
        redo_action.triggered.connect(self.editor.redo)
        edit_menu.addAction(redo_action)

        cut_action = QAction(QIcon("cut.png"), 'Cut', self)
        cut_action.triggered.connect(self.editor.cut)
        edit_menu.addAction(cut_action)

        copy_action = QAction(QIcon("copy.png"), 'Copy', self)
        copy_action.triggered.connect(self.editor.copy)
        edit_menu.addAction(copy_action)

        paste_action = QAction(QIcon("paste.png"), 'Paste', self)
        paste_action.triggered.connect(self.editor.paste)
        edit_menu.addAction(paste_action)
       
    


        view_menu = QMenu('View', self)
        menu_bar.addMenu(view_menu)

        full_screen = QAction(QIcon('full_screen.png'), 'Full Screen', self)
        full_screen.triggered.connect(self.showMaximized)
        view_menu.addAction(full_screen)

        normal_view = QAction(QIcon('exit_full_screen.png'), 'Normal View', self)
        normal_view.triggered.connect(self.showNormal)
        view_menu.addAction(normal_view)
        
        self.setMenuBar(menu_bar)
    
    #?########### ? Save File #?###########
    def save_file(self):
        if(self.path == ''):
            self.save_file_as()  #if file not previously saved, goto save_file_as
        else:
            text = self.editor.toPlainText() #concerting text to string

            try:
                with open(self.path, 'w') as f:
                    f.write(text)   # Creating file
                    self.update_title()
                
            except Exception as e:
                print(e)



    #?########### ? Save File As #?###########

    def save_file_as(self):
        self.path, _ = QFileDialog.getSaveFileName(self, "Save file", "", "text documents (*.txt);;Text documents (*.text);;All files (*.*)")

        if(self.path == ''):
            return      # if file path '' then return
        
        text = self.editor.toPlainText()

        try:
            with open(self.path, 'w') as f:  #opening file
                f.write(text)
                self.update_title()
                
        except Exception as e:
            print(e)



    #?########### ? Open File #?###########

    def open_file(self):
        self.path, _ = QFileDialog.getOpenFileName(self, "Save file", "", "text documents (*.txt);;Text documents (*.text);;All files (*.*)")

        try:
            with open(self.path, 'r') as f:
                text = f.read()
                self.editor.setText(text)
                self.update_title()

        except Exception as e:
            print(e)



    #?########### ? Export as PDF #?###########

    def export_pdf(self):
        self.path, _ = QFileDialog.getSaveFileName(self, "ExportPDF", "", "PDF files (*.pdf)")

        printer = QPrinter(QPrinter.PrinterMode.HighResolution)

        printer.OutputFormat(QPrinter.OutputFormat.PdfFormat)
        printer.setOutputFileName(self.path)
        self.editor.document().print_(printer)
    
    def update_title(self):
        self.editor.setWindowTitle(self.title + '' + self.path)

#?#################################################################
#?#?##################### Tool Bar Starts #?#?#####################
#?#################################################################


    def tool_bar(self):
        toolbar = QToolBar()
        
        #?########### Undo Action #?###########
        undo_action = QAction(QIcon("undo.png"), 'undo', self)
        undo_action.triggered.connect(self.editor.undo)
        toolbar.addAction(undo_action)          #Undo Action
        
        toolbar.addSeparator()   #Adding separator
        toolbar.addSeparator()
        
        #?########### End of Undo Action #?###########
        
        
        
        #?########### Redo Action #?###########
        redo_action = QAction(QIcon("redo.png"), 'redo', self)
        redo_action.triggered.connect(self.editor.redo)
        toolbar.addAction(redo_action)    #Redo action
        
        toolbar.addSeparator()
        toolbar.addSeparator()
        
        #?########### End of Redo Action #?###########
        
        
        
        #?########### Cut Action #?###########
        cut_action = QAction(QIcon("cut.png"), 'cut', self)
        cut_action.triggered.connect(self.editor.cut)
        toolbar.addAction(cut_action)       #Cut action
        
        toolbar.addSeparator()
        toolbar.addSeparator()
        
        #?########### End of Cut Action#?###########
        
        
        
        #?########### Copy Action #?###########
        copy_action = QAction(QIcon("copy.png"), 'copy', self)
        copy_action.triggered.connect(self.editor.copy)
        toolbar.addAction(copy_action)      #Copy action
        
        toolbar.addSeparator()
        toolbar.addSeparator()
        #?########### End of Copy Action #?###########
        
        
        
        
        #?########### Paste Action #?###########
        paste_action = QAction(QIcon("paste.png"), 'paste', self)
        paste_action.triggered.connect(self.editor.paste)
        toolbar.addAction(paste_action)     #Paste action
        
        #?########### End of Paste Action #?###########
        
        
    
        #?########### Text Italic#?###########
        italic_action = QAction(QIcon("italic.png"), 'italic', self)
        italic_action.triggered.connect(self.italic_text)
        toolbar.addAction(italic_action)   

        toolbar.addSeparator()
        toolbar.addSeparator()
        
        #?########### End of Text Italic#?###########

        #?########### Text Bold #?###########
        bold_action = QAction(QIcon("bold.png"), 'bold', self)
        bold_action.triggered.connect(self.bold_text)
        toolbar.addAction(bold_action)   
      

        toolbar.addSeparator()
        toolbar.addSeparator()
        
        #?########### End of Text bold#?###########

        #?########### Text Underline #?###########

        underline_action = QAction(QIcon("underline.png"), 'underline', self)
        underline_action.triggered.connect(self.underline_text)
        toolbar.addAction(underline_action)  

        toolbar.addSeparator()
        toolbar.addSeparator()
        
        #?########### End of Text Underline #?###########
        


        #?########### Font Size Box #?###########

        self.font_size_box.setValue(20)
        self.font_size_box.valueChanged.connect(self.set_font_size)
        toolbar.addWidget(self.font_size_box)

        toolbar.addSeparator()
        toolbar.addSeparator()
        
        #?########### End of Font Size Box #?###########
        


        #?########### Save Action #?###########

        save_action = QAction(QIcon('save.png'), 'Save', self)
        save_action.triggered.connect(self.save_file)
        toolbar.addAction(save_action)
        
        #?########### End of Save Action #?###########
        
        #?########### Arefin, don't change this! #?###########
        toolbar.setStyleSheet("QToolBar{spacing:5px;}")
        toolbar.setMovable(False)
        self.addToolBar(toolbar)
        #?########### Arefin, don't change this! #?###########
        
        
        
        
    #?########### Font Size Logic #?###########
        
    def set_font_size(self):
        value = self.font_size_box.value()
        self.editor.setFontPointSize(value)
        
    
    def bold_text(self):
        if (self.editor.fontWeight() != QFont.Bold):
            self.editor.setFontWeight(QFont.Bold)
            return
        else:
            self.editor.setFontWeight(QFont.Normal)

        
    #?###########  Italic Text Logic #?###########
    
    def italic_text(self):    
        if(self.editor.fontItalic()):
            self.editor.setFontItalic(False)
        else:    
            self.editor.setFontItalic(True)


    #?########### Font Underline #?###########
    
    def underline_text(self):
        state = self.editor.fontUnderline()
        self.editor.setFontUnderline(not(state))
        

    
 
######Executing######
app = QApplication(sys.argv)
window = Word()
window.show()
sys.exit(app.exec_())
