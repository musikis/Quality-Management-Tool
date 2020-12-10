import sys
import os
from PySide6 import QtUiTools, QtGui
from PySide6.QtWidgets import QApplication, QMainWindow, QTableWidgetItem, QMessageBox, QFileDialog
import random
import numpy
from openpyxl import Workbook,styles
from openpyxl.styles import Alignment,Border,Side

# 초기치 설정

nominal = 5 # 설계치
n = 3 #샘플 개수
lowtol = -0.2 # 하한공차
uptol = 0.2 # 상한공차
low=4.9 #난수 하한값
high=5.1 #난수 상한값
dirName=os.getcwd()
a=[] # 테이블 1열
b=[] # 테이블 2열
wb = Workbook() # 엑셀 현행 통합문서
ws = wb.active #엑셀 현행 워크시트


class MainView(QMainWindow):    #몰라 이하는 QT 쓰려면 무조건 있어야 한대

    def __init__(self):

        super().__init__()

        self.setupUI()



 

    def setupUI(self):  # UI 불러오기

        global UI_set

        # ui파일 경로 지정

        UI_set = QtUiTools.QUiLoader().load(resource_path("form.ui"))

        # 테이블 위젯 초기 세팅
        
        UI_set.tableWidget.setRowCount(n+6)
        UI_set.tableWidget.setHorizontalHeaderLabels(['Label','value'])

        #푸시버튼

        UI_set.pushButton.clicked.connect(self.start)
        UI_set.pushButton_2.clicked.connect(self.button2_clicked)
        UI_set.doubleSpinBox.valueChanged.connect(self.dsb1_clicked)
        UI_set.doubleSpinBox.setValue(nominal)
        UI_set.doubleSpinBox_2.valueChanged.connect(self.dsb2_clicked)
        UI_set.doubleSpinBox_2.setValue(uptol)
        UI_set.doubleSpinBox_3.valueChanged.connect(self.dsb3_clicked)
        UI_set.doubleSpinBox_3.setValue(lowtol)
        UI_set.spinBox.valueChanged.connect(self.sb_clicked)
        UI_set.spinBox.setValue(n)
        UI_set.doubleSpinBox_4.valueChanged.connect(self.dsb4_clicked)
        UI_set.doubleSpinBox_4.setValue(high)
        UI_set.doubleSpinBox_5.valueChanged.connect(self.dsb5_clicked)
        UI_set.doubleSpinBox_5.setValue(low)
        UI_set.pushButton_3.clicked.connect(self.browse_dest_path)
        UI_set.lineEdit.setText(dirName)
        UI_set.lineEdit.textChanged.connect(self.dest_path)
        UI_set.pushButton_4.clicked.connect(self.save_excel)
        
        self.setCentralWidget(UI_set)

        self.setWindowTitle("Random Value Maker")

        self.setWindowIcon(QtGui.QPixmap(resource_path("./images/jbmpa.png")))

        self.resize(330,630)

        self.show()


    def save_excel(self):
        global a,b
        global dirName
        msgBox = QMessageBox()
        

        if not a:
            msgBox.setWindowTitle("Alert Window") # 메세지창의 상단 제목
            msgBox.setWindowIcon(QtGui.QPixmap("info.png")) # 메세지창의 상단 icon 설정
            msgBox.setIcon(QMessageBox.Warning) # 메세지창 내부에 표시될 아이콘
            msgBox.setInformativeText("먼저 실행 버튼을 눌러주세요.") # 메세지 내용
            msgBox.setStandardButtons(QMessageBox.Ok)
            msgBox.exec_()
            return

        THIN_BORDER = Border(Side('thin'),Side('thin'),Side('thin'),Side('thin'))

        for i in range(0,n+10):
            ws ['A%d'%(i+1)] = a[i]
            ws ['A%d'%(i+1)].alignment = Alignment('center', 'center')
            ws ['A%d'%(i+1)].border = THIN_BORDER
            ws ['B%d'%(i+1)] = b[i]
            ws ['B%d'%(i+1)].number_format = '0.000'
            ws ['B%d'%(i+1)].border = THIN_BORDER
            
            
        #fileName='random.xlsx'
        #fullPath=(dirName+'\\random.xlsx')
        fullPath=os.path.join(dirName , 'random.xlsx')
        wb.save(fullPath)

        msgBox.setWindowTitle("Alert Window") # 메세지창의 상단 제목
        msgBox.setWindowIcon(QtGui.QPixmap("info.png")) # 메세지창의 상단 icon 설정
        msgBox.setIcon(QMessageBox.Warning) # 메세지창 내부에 표시될 아이콘
        msgBox.setInformativeText("저장이 완료되었습니다.") # 메세지 내용
        msgBox.setStandardButtons(QMessageBox.Ok)
        msgBox.exec_()

    def dest_path(self):
        global dirName
        dirName=UI_set.lineEdit.text()

    def browse_dest_path(self):
        global dirName
        dirName = QFileDialog.getExistingDirectory(self, self.tr("Open Data files"), dirName , QFileDialog.ShowDirsOnly)
        #dirName = QFileDialog.getOpenFileName(self, 'Open file', "","All Files(*);; Python Files(*.xls,*xlsx)", dirName)
        
        if dirName == "": # 사용자가 취소를 누를 때
            print("폴더 선택 취소")
            return
    
        UI_set.lineEdit.setText(dirName)

    def start(self):
        msgBox = QMessageBox()
        global a,b
        del a[:],b[:]
        if lowtol>uptol:
            msgBox.setWindowTitle("Alert Window") # 메세지창의 상단 제목
            msgBox.setWindowIcon(QtGui.QPixmap("info.png")) # 메세지창의 상단 icon 설정
            msgBox.setIcon(QMessageBox.Warning) # 메세지창 내부에 표시될 아이콘
            msgBox.setInformativeText("하한공차가 상한공차보다 큽니다") # 메세지 내용
            msgBox.setStandardButtons(QMessageBox.Ok)
            msgBox.exec_()
            return

        if low>high :
            msgBox.setWindowTitle("Alert Window") # 메세지창의 상단 제목
            msgBox.setWindowIcon(QtGui.QPixmap("info.png")) # 메세지창의 상단 icon 설정
            msgBox.setIcon(QMessageBox.Warning) # 메세지창 내부에 표시될 아이콘
            msgBox.setInformativeText("하한값이 상한값보다 큽니다") # 메세지 내용
            msgBox.setStandardButtons(QMessageBox.Ok)
            msgBox.exec_()
            return
        
        UI_set.tableWidget.setRowCount(n+10)

        for i in range(0,n):
            a.insert(i , i+1)

        for i in range(0,n):
            b.insert(i , random.uniform(low,high))

        for i in range(0,n):            
            UI_set.tableWidget.setItem(i , 0 , QTableWidgetItem('%d'%a[i]))    
            UI_set.tableWidget.setItem(i , 1 , QTableWidgetItem('%.3f'%b[i])) 

        mx=max(b)
        mn=min(b)
        sigma=numpy.std(b)
        USL=uptol+nominal
        LSL=lowtol+nominal
        cp=(USL-LSL)/6/sigma
        R=mx-mn
        xbar=numpy.mean(b)
        cpk = min(((USL-xbar)/(3*sigma)),((xbar-LSL)/(3*sigma)))
        

        a += ['Nominal','USL','LSL','Max','Min','XBar','R','std','cp','cpk']
        b += [nominal,USL,LSL,mx,mn,xbar,R,sigma,cp,cpk]

              
        for i in range (0,10):
            UI_set.tableWidget.setItem(n+i , 0 , QTableWidgetItem(a[n+i])) 
            UI_set.tableWidget.setItem(n+i , 1 , QTableWidgetItem('%.3f'%b[n+i])) 


    def button2_clicked(self):
        reply = QMessageBox.question(self, '종료', '종료하시겠습니까?',
                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            sys.exit()
        else:
            return
        

    def dsb1_clicked(self):
        global nominal
        nominal=UI_set.doubleSpinBox.value()
       

    def dsb2_clicked(self):
        global uptol
        uptol=UI_set.doubleSpinBox_2.value()
        

    def dsb3_clicked(self):
        global lowtol
        lowtol=UI_set.doubleSpinBox_3.value()
       

    def sb_clicked(self):
        global n
        n=UI_set.spinBox.value()
        

    def dsb4_clicked(self):
        global high
        high=UI_set.doubleSpinBox_4.value()

    def dsb5_clicked(self):
        global low
        low=UI_set.doubleSpinBox_5.value()


 

#파일 경로

#pyinstaller로 원파일로 압축할때 경로 필요함

def resource_path(relative_path):

    if hasattr(sys, '_MEIPASS'):

        return os.path.join(sys._MEIPASS, relative_path)

    return os.path.join(os.path.abspath("."), relative_path)

 

if __name__ == '__main__':

    app = QApplication(sys.argv)
        
    main = MainView()

    #main.show()

    sys.exit(app.exec_())
