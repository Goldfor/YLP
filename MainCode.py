import mammoth
import pypandoc
from PyQt5 import Qt, QtGui, QtCore, QtWidgets
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QMainWindow, QAction, QApplication, QFileDialog


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.style = ['QPushButton {background-color: #eeeeec}', \
        'QPushButton {background-color: #babdb6}']

        self.fileURL = ''

        self.sbFontSize = Qt.QSpinBox()
        self.sbFontSize.setRange(5, 40)
        self.sbFontSize.setValue(12)
        self.sbFontSize.valueChanged.connect(self.onFontSizeChanged)

        self.button = []
        self.button.append(Qt.QPushButton('Подчеркнуть', self))
        self.button[0].clicked.connect(self.onFontUnderlineChanged)
        self.button.append(Qt.QPushButton('Курсив', self))
        self.button[1].clicked.connect(self.onFontItalicChanged)
        self.button.append(Qt.QPushButton('Жирный', self))
        self.button[2].clicked.connect(self.onFontBoldChanged)

        self.statusBar()
        self.setFocus()

        self.combo = QtWidgets.QComboBox(self)
        self.combo.setEditable(True)
        self.combo.editTextChanged.connect(self.onFontFamilyChanged)
        self.combo.addItems(QtGui.QFontDatabase().families())

        self.textEdit = Qt.QTextEdit('')
        self.textEdit.setFontWeight(12)

        if 'Arial' in QtGui.QFontDatabase().families():
            self.textEdit.setFontFamily('Arial')

        else:
            self.textEdit.setFontFamily('Ubuntu')

        self.textEdit.cursorPositionChanged.connect(self.cursorPositionChanged)

        layout = Qt.QHBoxLayout()
        layout.addWidget(self.sbFontSize)

        for i in self.button[0:3]:
            i.setStyleSheet(self.style[0])
            layout.addWidget(i)

        layout.addWidget(self.combo)

        edit = Qt.QWidget()
        edit.setLayout(layout)

        layout1 = Qt.QHBoxLayout()
        layout1.addWidget(self.sbFontSize)

        for i in self.button[3:]:
            i.setStyleSheet(self.style[0])
            layout1.addWidget(i)

        redact = Qt.QWidget()
        redact.setLayout(layout1)


        layout2 = Qt.QVBoxLayout()
        layout2.addWidget(redact)
        layout2.addWidget(edit)
        layout2.addWidget(self.textEdit)

        central_widget = Qt.QWidget()
        central_widget.setLayout(layout2)

        self.cursorPositionChanged()

        self.setCentralWidget(central_widget)

        Open = QAction('Open', self)
        Open.setShortcut('Ctrl+O')
        Open.setStatusTip('Open new File')
        Open.triggered.connect(self.openFile)

        Save = QAction('Save', self)
        Save.setShortcut('Ctrl+S')
        Save.setStatusTip('Save File')
        Save.triggered.connect(self.saveFile)

        New = QAction('New', self)
        New.setShortcut('Ctrl+N')
        New.setStatusTip('New File')
        New.triggered.connect(self.newFile)

        menubar = self.menuBar()
        file = menubar.addMenu('&File')
        file.addAction(Open)
        file.addAction(Save)
        file.addAction(New)

    def cursorPositionChanged(self):
        font = self.textCursorSelect()
        ListBool = [font.fontUnderline(), \
            font.fontItalic(), \
            (font.fontWeight() == QFont.Bold)]

        for i in range(3):
            if ListBool[i]:
                self.button[i].setStyleSheet(self.style[1])
            else:
                self.button[i].setStyleSheet(self.style[0])
        try:
            if font.fontPointSize() != 0:
                self.sbFontSize.setValue(font.fontPointSize())
            else:
                self.sbFontSize.setValue(12)

            if font.fontFamily() == '':
                if 'Arial' in QtGui.QFontDatabase().families():
                    self.textEdit.setFontFamily('Arial')
                    self.combo.setCurrentText('Arial')


                else:
                    self.textEdit.setFontFamily('Ubuntu')
                    self.combo.setCurrentText('Ubuntu')

            else:
                self.combo.setCurrentText(font.fontFamily())

        except:
            pass

    def onFontFamilyChanged(self, select):
        index = self.combo.findText(select)
        if index > -1:
            self.combo.setCurrentIndex(index)
            text_char_format = Qt.QTextCharFormat()
            text_char_format.setFontFamily(select)

            self.mergeFormatOnWordOrSelection(text_char_format)

    def onFontSizeChanged(self, value):
        text_char_format = Qt.QTextCharFormat()
        text_char_format.setFontPointSize(value)

        self.mergeFormatOnWordOrSelection(text_char_format)

    def onFontUnderlineChanged(self):
        font = self.textCursorSelect().fontUnderline()
        text_char_format = Qt.QTextCharFormat()
        text_char_format.setFontUnderline(not font)

        self.changeButtonStyle(font, 0)

        self.mergeFormatOnWordOrSelection(text_char_format)

    def onFontItalicChanged(self):
        font = self.textCursorSelect().fontItalic()
        text_char_format = Qt.QTextCharFormat()
        text_char_format.setFontItalic(not font)

        self.changeButtonStyle(font, 1)

        self.mergeFormatOnWordOrSelection(text_char_format)

    def onFontBoldChanged(self):
        font = self.textCursorSelect().fontWeight()
        text_char_format = Qt.QTextCharFormat()

        if font != QFont.Bold:
            text_char_format.setFontWeight(QFont.Bold)

        else:
            text_char_format.setFontWeight(QFont.Normal)

        self.changeButtonStyle(font == QFont.Bold, 2)

        self.mergeFormatOnWordOrSelection(text_char_format)

    def textCursorSelect(self):
        cursor = self.textEdit.textCursor()

        return cursor.charFormat()

    def mergeFormatOnWordOrSelection(self, text_char_format):
        cursor = self.textEdit.textCursor()

        cursor.mergeCharFormat(text_char_format)

    def changeButtonStyle(self, ifBool, i):
        if ifBool:
            self.button[i].setStyleSheet(self.style[1])
            
        else:
            self.button[i].setStyleSheet(self.style[0])

    def saveFile(self):
        name = self.fileURL
        if name == '':
            name = QFileDialog.getSaveFileName(self, 'Выбрать докумет', '', \
            "Documents (*.docx)")[0]
        print(name)
        self.fileURL = name
        try:
            S = open(str(name[:-4]+'html'), 'w')
            S.write(self.textEdit.toHtml())
            S.close()
            output = pypandoc.convert(source=str(name[:-4]+'html'), format='html', \
                to='docx', outputfile=f'{name}', extra_args=['-RTS'])
        except:
            pass

    def openFile(self):
        if self.fileURL != '':
            self.saveFile()
        name = QFileDialog.getOpenFileName(self, 'Выбрать докумет', '', \
            "Documents (*.docx)")[0]
        I = open(name, 'rb')
        B = open(f'{name}~~', 'wb')
        document = mammoth.convert_to_html(I)
        B.write(document.value.encode('utf8'))
        I.close()
        B.close()
        O = open(f'{name}~~', 'r')
        self.fileURL = name
        out = O.read()
        O.close()
        self.textEdit.setText(out)
        print(out)

    def newFile(self):
        self.saveFile()
        self.fileURL = ''
        self.textEdit.setText('')


if __name__ == '__main__':
    app = QApplication([])

    mW = MainWindow()
    mW.show()

    app.exec()