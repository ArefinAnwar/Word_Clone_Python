import sys
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtPrintSupport import *
from PyQt5 import QtCore, QtGui
from PyQt5.QtGui import * 
from PyQt5.QtCore import * 
from PyQt5.QtWidgets import QMessageBox
from darktheme.widget_template import DarkPalette
import pyttsx3


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

        self.font_size_box = QSpinBox()  #Font size Changing box
        
        self.menu_bar()  
        
        self.tool_bar()

        self.path = '' #Stores path

        self.flag = 0  #Realated to changing theme
        
    def closeEvent(self, event):

        if(self.path == ''):
            quit_msg = "The file is not saved. Are you sure you want to exit Word?"
            reply = QMessageBox.question(self, 'Warning! File not saved', quit_msg, QMessageBox.Yes, QMessageBox.No)

            if(reply == QMessageBox.Yes):
                event.accept()
            else:
                event.ignore()
        else:
            quit_msg = "Are you sure you want to exit the program Word?"
            reply = QMessageBox.question(self, 'Exit Word', quit_msg, QMessageBox.Yes, QMessageBox.No)

            if(reply == QMessageBox.Yes):
                event.accept()
            else:
                event.ignore()

        
    
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

        about_menu = QMenu('About', self)
        about_menu.triggered.connect(self.about)
        menu_bar.addMenu(about_menu)

        about = QAction(QIcon('about.png'), 'About', self)
        about.triggered.connect(self.about)
        about_menu.addAction(about)
        
        self.setMenuBar(menu_bar)
    def about(self):
        quit_msg = """ 
        Word Editor
        This apllication is made by Arefin Anwar. 
        Copyright ??2021 Arefin Anwar. All rights reserved.

        Version: 1.2"""
        reply = QMessageBox.question(self, 'About', quit_msg, QMessageBox.Ok)

    #?########### ? Save File #?###########
    def save_file(self):
        if(self.path == ''):
            self.save_file_as()  #if file not previously saved, goto save_file_as
        else:
            text = self.editor.toPlainText() #concerting text to string

            try:
                with open(self.path, 'w') as f:
                    f.write(text)   # Creating file
                    
                
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
                
                
        except Exception as e:
            print(e)



    #?########### ? Open File #?###########

    def open_file(self):
        self.path, _ = QFileDialog.getOpenFileName(self, "Save file", "", "text documents (*.txt);;Text documents (*.text);;All files (*.*)")

        try:
            with open(self.path, 'r') as f:
                text = f.read()
                self.editor.setText(text)
                if(self.flag == 2):
                    self.editor.selectAll() #Sellect all
                    self.white_color = QColor(255, 255, 255)
                    self.editor.setTextColor(self.white_color)  #Set white colour
                    value = self.font_size_box.value()
                    self.editor.setFontPointSize(value)
                else:
                    value = self.font_size_box.value()
                    self.editor.setFontPointSize(value)

        except Exception as e:
            print(e)



    #?########### ? Export as PDF #?###########

    def export_pdf(self):
        if(self.flag == 2):
            self.editor.selectAll() #Sellect all
            self.black_color = QColor(0, 0, 0) #Set black colour
            self.editor.setTextColor(self.black_color)
            self.path, _ = QFileDialog.getSaveFileName(self, "ExportPDF", "", "PDF files (*.pdf)")

            printer = QPrinter(QPrinter.PrinterMode.HighResolution)

            printer.OutputFormat(QPrinter.OutputFormat.PdfFormat)
            printer.setOutputFileName(self.path)
            self.editor.document().print_(printer)
            
            self.editor.selectAll() #Sellect all
            self.white_color = QColor(255, 255, 255)
            self.editor.setTextColor(self.white_color)  #Set white colour
        else:
            self.path, _ = QFileDialog.getSaveFileName(self, "ExportPDF", "", "PDF files (*.pdf)")

            printer = QPrinter(QPrinter.PrinterMode.HighResolution)

            printer.OutputFormat(QPrinter.OutputFormat.PdfFormat)
            printer.setOutputFileName(self.path)
            self.editor.document().print_(printer)    


#?#################################################################
#?#?##################### Tool Bar Starts #?#?#####################
#?#################################################################


    def tool_bar(self):
        toolbar = QToolBar()
        
        #?########### Undo Action #?###########
        undo_action = QAction(QIcon("undo.png"), 'undo', self)
        undo_action.triggered.connect(self.custom_undo)
        toolbar.addAction(undo_action)          #Undo Action

        #toolbar.addSeparator()   #Adding separator
        #toolbar.addSeparator()
        
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
        
        #toolbar.addSeparator()
        #toolbar.addSeparator()

        
        
        #?########### End of Cut Action#?###########
        
        
        
        #?########### Copy Action #?###########
        copy_action = QAction(QIcon("copy.png"), 'copy', self)
        copy_action.triggered.connect(self.editor.copy)
        toolbar.addAction(copy_action)      #Copy action
        
        #toolbar.addSeparator()
        #toolbar.addSeparator()
        #?########### End of Copy Action #?###########
        
        
        
        
        #?########### Paste Action #?###########
        paste_action = QAction(QIcon("paste.png"), 'paste', self)
        paste_action.triggered.connect(self.editor.paste)
        toolbar.addAction(paste_action)     #Paste action
        
        toolbar.addSeparator()
        toolbar.addSeparator()

        #?########### End of Paste Action #?###########
                                    
        

        #?########### Text Bold #?###########
        bold_action = QAction(QIcon("bold.png"), 'bold', self)
        bold_action.triggered.connect(self.bold_text)
        toolbar.addAction(bold_action)   
      
        #toolbar.addSeparator()
        #toolbar.addSeparator()
        
        #?########### End of Text bold#?###########



        #?########### Text Italic#?###########

        italic_action = QAction(QIcon("italic.png"), 'italic', self)
        italic_action.triggered.connect(self.italic_text)
        toolbar.addAction(italic_action)   

        #toolbar.addSeparator()
        #toolbar.addSeparator()
        
        #?########### End of Text Italic #?###########



        #?########### Text Underline #?###########

        underline_action = QAction(QIcon("underline.png"), 'underline', self)
        underline_action.triggered.connect(self.underline_text)
        toolbar.addAction(underline_action)  

        #toolbar.addSeparator()
        toolbar.addSeparator()
        
        #?########### End of Text Underline #?###########
        


        #?########### Font Size Box #?###########

        self.font_size_box.setValue(20)
        self.font_size_box.valueChanged.connect(self.set_font_size)
        toolbar.addWidget(self.font_size_box)

        toolbar.addSeparator()
        toolbar.addSeparator()
        
        #?########### End of Font Size Box #?###########
        

        #?########### Center Align #?###########

        center_align_action = QAction(QIcon('center-align.png'), 'Center Align', self)
        center_align_action.triggered.connect(self.center_align)
        toolbar.addAction(center_align_action)

        #toolbar.addSeparator()
        #toolbar.addSeparator()
        
        #?########### End of Center Align  #?###########



        #?########### Left Align #?###########

        left_align_action = QAction(QIcon('left-align.png'), 'Left Align', self)
        left_align_action.triggered.connect(self.left_align)
        toolbar.addAction(left_align_action)

        #toolbar.addSeparator()
        #toolbar.addSeparator()
        
        #?########### End of Left Align  #?###########
        


        #?########### Right Align #?###########

        right_align_action = QAction(QIcon('right-align.png'), 'Right Align', self)
        right_align_action.triggered.connect(self.right_align)
        toolbar.addAction(right_align_action)

        #toolbar.addSeparator()
        #toolbar.addSeparator()
        
        #?########### End of Right Align  #?###########



        #?########### Right Align #?###########

        justification_align_action = QAction(QIcon('justification.png'), 'Justification Align', self)
        justification_align_action.triggered.connect(self.justification_align)
        toolbar.addAction(justification_align_action)

        toolbar.addSeparator()
        toolbar.addSeparator()
        #?########### End of Right Align  #?###########
        #?########### Change Theme #?###########

        
        change_theme_action = QAction(QIcon('white.png'), 'Change Theme', self)
        change_theme_action.triggered.connect(self.change_theme)
        toolbar.addAction(change_theme_action)
        
        toolbar.addSeparator()
        toolbar.addSeparator()
        
        #?########### End of Change Theme #?###########
        #?########### Save Action #?###########

        save_action = QAction(QIcon('save.png'), 'Save', self)
        save_action.triggered.connect(self.save_file)
        toolbar.addAction(save_action)

        toolbar.addSeparator()
        toolbar.addSeparator()
        
        #?########### End of Save Action #?###########

        text_to_speech = QAction(QIcon('speech.png'), 'Speech', self)
        text_to_speech.triggered.connect(self.text_to_speech)
        toolbar.addAction(text_to_speech)

        

        #?########### Arefin, don't change this! #?###########
        toolbar.setStyleSheet("QToolBar{spacing:3px;}")
        toolbar.setMovable(False)
        self.addToolBar(toolbar)
        #?########### Arefin, don't change this! #?###########
        
        
        
    def text_to_speech(self):
        engine = pyttsx3.init()
        text = self.editor.toPlainText()
        engine.say(text)
        engine.runAndWait()
    #?########### Custom Undo Function #?###########

    def custom_undo(self):
        self.editor.undo()
        value = self.font_size_box.value()
        self.editor.setFontPointSize(value)
        #Arefin, for adding the line above you had to write a function
        #self.editor.setAlignment('Qt::AlignLeft')


    #?########### Font Size Logic #?###########
        
    def set_font_size(self):
        value = self.font_size_box.value()
        self.editor.setFontPointSize(value)
        


    #?########### Bold Text Function #?###########

    def bold_text(self):
        if (self.editor.fontWeight() != QFont.Bold):
            self.editor.setFontWeight(QFont.Bold)
            return
        else:
            self.editor.setFontWeight(QFont.Normal)


        
    #?###########  Italic Text Function #?###########
    
    def italic_text(self):    
        if(self.editor.fontItalic()):
            self.editor.setFontItalic(False)
        else:    
            self.editor.setFontItalic(True)


    #?########### Font Underline Function #?###########
    
    def underline_text(self):
        state = self.editor.fontUnderline()
        self.editor.setFontUnderline(not(state))
        


    #?########### Center Align Function #?###########

    def center_align(self):
        self.editor.setAlignment(Qt.AlignCenter)



    #?########### Left Align Function #?###########

    def left_align(self):
        self.editor.setAlignment(Qt.AlignLeft)



    #?########### Right Align Function #?###########
    def right_align(self):
        self.editor.setAlignment(Qt.AlignRight)



    #?########### Justification Align Function #?###########
    def justification_align(self):
        self.editor.setAlignment(Qt.AlignJustify)



    #?########### Change Theme Function #?###########
    def change_theme(self):
        if(self.flag == 0):
            self.editor.selectAll() #Sellect all
            self.white_color = QColor(255, 255, 255)
            self.editor.setTextColor(self.white_color)  #Set white colour
            value = self.font_size_box.value()
            self.editor.setFontPointSize(value) #Font Size
            self.editor.setStyleSheet("background-color: rgb(28, 28, 28);") #Background colour
            self.flag = 2
            
        else:
            self.editor.selectAll() #Sellect all
            self.black_color = QColor(0, 0, 0) #Set black colour
            self.editor.setTextColor(self.black_color)
            value = self.font_size_box.value()
            self.editor.setFontPointSize(value) #Font Size
            self.editor.setStyleSheet("background-color: rgb(255, 255, 255);")#Background colour
            self.flag = 0
            
 
######Executing######
app = QApplication(sys.argv)
window = Word()
window.show()
sys.exit(app.exec_())
