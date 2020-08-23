## Jock 1.0v
## Made by Lee.Down
## L.DOWNCOMP@gamil.com

import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import QIcon

from tkinter import Tk

import win32com.client as win32
import time
import win32gui
import win32con
import os
import shutil
import pickle
import ctypes
import time

subjects=[]


class MyApp(QMainWindow):
    def __init__(self, parent=None):           
        super(MyApp,self).__init__(parent)

        ##### 메인화면 꾸미기 함수 실행
        self.initUi()

        ##### 위젯클래스 메인윈도우의 센터에 셋팅
        self.mywidget = Widgets(self)
        self.setCentralWidget(self.mywidget)

    ##### 메인화면 꾸미기
    def initUi(self):
        global state
        self.setWindowTitle("Auto족_1.0v")
        self.setGeometry(100,100,1150,800)

        menubar = self.menuBar()
        aboutAction = QAction('about', self)
        aboutAction.triggered.connect(self.aboutmethod)

        refreshAction = QAction('Refresh', self)
        refreshAction.triggered.connect(self.refreshall)
        menubar.addAction(refreshAction)
        menubar.addMenu('help').addAction(aboutAction)

        state = self.statusBar()
        self.setStatusBar(state)
        state.showMessage('Ready')


    def closeEvent(self, QCloseEvent):
        re = QMessageBox.question(self, "종료 확인", "종료 하시겠습니까?", QMessageBox.Yes|QMessageBox.No)

        if re == QMessageBox.Yes:
            QCloseEvent.accept()
        else:
            QCloseEvent.ignore()  

    def aboutmethod(self):
        QMessageBox.about(self, "about","2020.07\nmade by L_Down\nL.DOWNCOMP@gamil.com   \n" )
    
    def refreshall(self):
        self.mywidget.hide()
        if os.path.isfile(os.path.join( os.getcwd(),"data/savedata.p") ):
            os.remove(os.path.join( os.getcwd(),"data/savedata.p"))
        self.mywidget = Widgets(self)
        self.setCentralWidget(self.mywidget)



class Widgets(QWidget):
    def __init__(self, parent):           
        super(Widgets,self).__init__(parent)

        ##### 위젯 함수 실행
        self.initWidget(parent)
      

    ##### 위젯셋팅
    def initWidget(self, parent):

        # 레이아웃
        self.mainlayout = QVBoxLayout()
        uplayout = QHBoxLayout()
        self.midlayout = QHBoxLayout()
        downlayout = QHBoxLayout()

        
              
        
        self.subject_ = QComboBox(self)
        self.subject_mf = QComboBox(self)
        self.subject_.setFixedSize(300,25)
        self.subject_mf.setFixedSize(150,25)
        
        self.current_name=""

        

        # 과목 ComboBox 채우기
        
        # if os.path.exists(os.path.join(os.getcwd(),"data/savedata.p")): #파일 로드
        #     with open(os.path.join(os.getcwd(),"data/savedata.p"), 'rb') as file:
        #         subjects = pickle.load(file)
            
        #     self.subject_.addItem("과목 선택")
        #     for i in range(0, len(subjects)):
        #         self.subject_.addItem(subjects[i].name)

        # else :
        #     self.subject_.addItem("갱신 버튼을 눌려")
        self.refresh()
        
        self.subject_mf.addItems(["중간고사","기말고사"])
        self.close_mf = QComboBox(self)
        self.close_mf.hide()
        self.close_mf.setFixedSize(150,25)
        

        # 테이블 추가
        self.tables = []
        for i in range(0, len(subjects)):

            if subjects[i].name == "기생충학":
         
                self.tables.append([QTableWidget(self), QTableWidget(self)])
           
                mid_label=["단원명"]
                f_label=["단원명"]

                for l in range( subjects[i].mid_end, subjects[i].mid_start -1, -1):
                    mid_label.append(str(l)+"년도")
                mid_label.append("전체")

                self.tables[i][0].setRowCount( subjects[i].mid_num )
                self.tables[i][0].setColumnCount(subjects[i].mid_end -subjects[i].mid_start +2 +1 )
                self.tables[i][0].setHorizontalHeaderLabels(mid_label)
                self.tables[i][0].resizeColumnsToContents()
                self.tables[i][0].resizeRowsToContents()
                self.tables[i][0].setEditTriggers(QAbstractItemView.NoEditTriggers)
                self.tables[i][0].setColumnWidth(0,550)
                for l in range(0, subjects[i].mid_num):
                    self.tables[i][0].setItem(l,0,QTableWidgetItem(subjects[i].mid_list[l]) )
                    for k in range(1, subjects[i].mid_end-subjects[i].mid_start +2 +1):
                        self.tables[i][0].setCellWidget(l,k, QCheckBox())
                        if k != subjects[i].mid_end-subjects[i].mid_start +2:
                            self.tables[i][0].cellWidget(l,k).stateChanged.connect(lambda state, row=l, k=subjects[i].mid_end-subjects[i].mid_start +2  : self.job_check(state, row, k))
                        else:
                            self.tables[i][0].cellWidget(l,k).stateChanged.connect(lambda state, row=l, k=subjects[i].mid_end-subjects[i].mid_start +2  : self.all_check(state, row, k))
                self.tables[i][0].hide()
                self.midlayout.addWidget(self.tables[i][0])





                for l in range( subjects[i].f_end, subjects[i].f_start -1, -1 ):
                    f_label.append(str(l)+"년도")
                f_label.append("전체")
                self.tables[i][1].setRowCount( subjects[i].f_num )
                self.tables[i][1].setColumnCount(subjects[i].f_end -subjects[i].f_start +2 +1)
                self.tables[i][1].setHorizontalHeaderLabels(f_label)
                self.tables[i][1].resizeColumnsToContents()
                self.tables[i][1].resizeRowsToContents()
                self.tables[i][1].setEditTriggers(QAbstractItemView.NoEditTriggers)
                self.tables[i][1].setColumnWidth(0,550)
                for l in range(0, subjects[i].f_num):
                    self.tables[i][1].setItem(l,0,QTableWidgetItem(subjects[i].f_list[l]) )
                    for k in range(1, subjects[i].f_end-subjects[i].f_start +2 +1):
                        self.tables[i][1].setCellWidget(l,k, QCheckBox())
                        if k != subjects[i].f_end-subjects[i].f_start +2:
                            self.tables[i][1].cellWidget(l,k).stateChanged.connect(lambda state, row=l, k=subjects[i].f_end-subjects[i].f_start +2  : self.job_check(state, row, k))
                        else:
                            self.tables[i][1].cellWidget(l,k).stateChanged.connect(lambda state, row=l, k=subjects[i].f_end-subjects[i].f_start +2  : self.all_check(state, row, k))
                self.tables[i][1].hide()
                self.midlayout.addWidget(self.tables[i][1])
            
            else :

                self.tables.append(QTableWidget(self))
                label=["단원명"]

                for l in range( subjects[i].end, subjects[i].start -1, -1):
                    label.append(str(l)+"년도")
                label.append("전체")

                self.tables[i].setRowCount( subjects[i].num )
                self.tables[i].setColumnCount(subjects[i].end -subjects[i].start +2 +1)
                self.tables[i].setHorizontalHeaderLabels(label)
                self.tables[i].resizeColumnsToContents()
                self.tables[i].resizeRowsToContents()
                self.tables[i].setEditTriggers(QAbstractItemView.NoEditTriggers)
                self.tables[i].setColumnWidth(0,550)
                for l in range(0, subjects[i].num):
                    self.tables[i].setItem(l,0,QTableWidgetItem(subjects[i].list[l]) )
                    for k in range(1, subjects[i].end-subjects[i].start +2 +1):
                        self.tables[i].setCellWidget(l,k, QCheckBox())
                        if k != subjects[i].end-subjects[i].start +2:
                            self.tables[i].cellWidget(l,k).stateChanged.connect(lambda state, row=l, k=subjects[i].end-subjects[i].start +2  : self.job_check(state, row, k))
                        else:
                            self.tables[i].cellWidget(l,k).stateChanged.connect(lambda state, row=l, k=subjects[i].end-subjects[i].start +2  : self.all_check(state, row, k))
              

                self.tables[i].hide()
                self.midlayout.addWidget(self.tables[i])

        

            
            


        


        table_ = QTableWidget(self)
        table_.setRowCount(20)
        table_.setColumnCount(7)
        table_.setHorizontalHeaderLabels(["단원명", "19년도","18년도","17년도","16년도","15년도","14년도"])
        table_.resizeColumnsToContents()
        table_.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table_.setColumnWidth(0,550)
        self.midlayout.addWidget(table_)
        self.current = table_
        

       


        


  

        

        filename_ = QLabel("파일명 : ", self)
        self.filename_edit = QLineEdit(self) 
        trans_ = QPushButton('변환', self)

       
        trans_.clicked.connect(self.trans_clicked)

        self.subject_.currentTextChanged.connect(self.unit_change)
        self.subject_mf.currentTextChanged.connect(self.unit_change)
 

        

        #업
        uplayout.addWidget(self.subject_)
        uplayout.addWidget(self.subject_mf)
        uplayout.addWidget(self.close_mf)

        

        

        downlayout.addWidget(filename_)
        downlayout.addWidget(self.filename_edit)
        downlayout.addWidget(trans_)
        
        self.mainlayout.addLayout(uplayout)
        self.mainlayout.addLayout(self.midlayout)
        self.mainlayout.addLayout(downlayout)
        self.setLayout(self.mainlayout)

    def trans_clicked(self):
        global state
        state.showMessage('진행 중......')
        state.repaint()


        if self.subject_.currentText() == "과목 선택":
            QMessageBox.about(self, "Warning","\n과목을 선택해주세요\n" )
            state.showMessage('Ready')
            return
        
        goal_name = self.filename_edit.text().strip()+".hwp"
        if goal_name == ".hwp":
            QMessageBox.about(self, "Warning","\n파일명을 입력해 주세요.\n" )
            state.showMessage('Ready')
            return
        if os.path.exists( os.path.join( os.getcwd(),"output/"+goal_name) ):
            QMessageBox.about(self, "Warning","\n파일명이 이미 존재합니다.\n" )
            state.showMessage('Ready')
            return
        
        shutil.copy(os.path.join(os.getcwd(),"양식/양식.hwp"), os.path.join(os.getcwd(), "output/"+goal_name))

        if self.current_name == "기생충학":
            if self.subject_mf.currentIndex() ==0:
                hwp_name = self.current_name+"_중간고사.hwp"
            else :
                hwp_name = self.current_name+"_기말고사.hwp"
        else:
            hwp_name = self.current_name+".hwp"

        

        hwp = win32.Dispatch("HWPFrame.HwpObject")
        h1 = win32gui.FindWindow(None, "빈 문서 1 - 한글")
        hwp2 = win32.Dispatch("HWPFrame.HwpObject")
        h2 = win32gui.FindWindow(None, "빈 문서 1 - 한글")
        hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule')
        hwp2.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule')
        
        hwp.Open(os.path.join( os.getcwd() , "hwp/"+hwp_name),hwp,hwp)
        hwp2.Open(os.path.join( os.getcwd() , "output/"+goal_name),hwp2,hwp2)

        win32gui.ShowWindow(h1, win32con.SW_HIDE)
        win32gui.ShowWindow(h2, win32con.SW_HIDE)
        time.sleep(0.5)

        

        hwp.SetMessageBoxMode(0x00010001)
        hwp2.MovePos(1,1,0)
        a = hwp2.CreateAction("ShapeCopyPaste")
        b = a.CreateSet()
        a.GetDefault(b)
        b.SetItem("Type", 2)
        a.Execute(b) # 모양 복사

        
        

        if self.current_name == "기생충학":
            if self.subject_mf.currentIndex() == 0:
                u_num = subjects[self.subject_.currentIndex()-1].mid_num    #단원 수
                u_list = subjects[self.subject_.currentIndex()-1].mid_list   # 단원 이름 리스트
                u_start = subjects[self.subject_.currentIndex()-1].mid_start # 단원 시작년도
                u_end = subjects[self.subject_.currentIndex()-1].mid_end     # " 끝 년도
                u_quiz = subjects[self.subject_.currentIndex()-1].mid_quiz   # 문제번호 배열
            else :
                u_num = subjects[self.subject_.currentIndex()-1].f_num    #단원 수
                u_list = subjects[self.subject_.currentIndex()-1].f_list   # 단원 이름 리스트
                u_start = subjects[self.subject_.currentIndex()-1].f_start # 단원 시작년도
                u_end = subjects[self.subject_.currentIndex()-1].f_end     # " 끝 년도
                u_quiz = subjects[self.subject_.currentIndex()-1].f_quiz   # 문제번호 배열


            quiz_num =1
            missing =""

            statenum =0
            for i in range(0, u_num):
                for k in range(1, u_end - u_start+2): #연도
                    if self.current.cellWidget(i,k).isChecked() == True:
                        statenum +=1
        
            now =0
            state.showMessage('진행 중......' +str(now) +"/"+str(statenum))
            state.repaint()

            for i in range(0, u_num): #단원

                
                

                # write_act =  hwp.CreateAction("InsertText")
                # write_set = write_act.CreateSet()
                # write_act.GetDefault(write_set)
                # write_set.SetItem("InsertText","1111")
                # write_act.Execute(write_set)
                hwp2_flag = 0


                

                find_act = hwp.CreateAction("ForwardFind")
                find_set = find_act.CreateSet()
                find_act.GetDefault(find_set)

                for k in range(1, u_end - u_start+2): #연도

                    if self.current.cellWidget(i,k).isChecked() == False:
                        continue

                    now +=1
                    state.showMessage('진행 중......' +str(now) +"/"+str(statenum))
                    state.repaint()
                    
                    u_temp = u_quiz[i][k-1].split(',')
                    if(u_temp[0] in ["None", ""]):
                        continue

                    #단원명 블럭 만들기
                    if hwp2_flag == 0 :
                        hwp2_flag = 1
                        hwp2.SelectText(0,48,0,56)
                        hwp2.Run("Copy")
                        hwp2.MovePos(3,1,1)
                        hwp2.Run("BreakPara")
                        hwp2.Run("DeleteBack")
                        hwp2.Run("BreakPara")
                        hwp2.Run("Paste")
                        hwp2.Run("MoveLeft")
                        hwp2.HAction.GetDefault("InsertText", hwp2.HParameterSet.HInsertText.HSet)
                        hwp2.HParameterSet.HInsertText.Text = u_list[i]
                        hwp2.HAction.Execute("InsertText", hwp2.HParameterSet.HInsertText.HSet)
                        hwp2.MovePos(3,1,1)
                        hwp2.Run("BreakPara")
                        hwp2.Run("BreakPara")
                        hwp2.Run("DeleteBack")

                
                

                    for l in range (0, len(u_temp)): #한 체크박스
                        hwp.MovePos(2,1,1)
                        num_name = '#'+ str(u_end-k+1)+'.'+str(self.subject_mf.currentIndex()+1)+'.'+u_temp[l]

                        find_set.SetItem("FindString",num_name)
                        find_set.SetItem("WholeWordOnly", True)
                        if find_act.Execute(find_set) == False:
                            missing += num_name+"  "
                            continue

                        start_pos = hwp.GetPosBySet().Item("Para") + 1
                        find_set.SetItem("WholeWordOnly", False)


                        if l == len(u_temp)-1:
                            find_set.SetItem("FindString","#")
                            find_act.Execute(find_set)

                            end_pos = hwp.GetPosBySet().Item("Para")

                        else:
                            find_set.SetItem("FindString","#"+str(u_end-k+1)+'.'+str(self.subject_mf.currentIndex()+1))
                            find_act.Execute(find_set)

                            end_pos = hwp.GetPosBySet().Item("Para")

                        hwp.SelectText(start_pos,0,end_pos,0)
                        hwp.Run('Copy')

                        if not '~' in u_temp[l] or not '-' in u_temp[l] : #단일문제

                            # #문제 번호 입력....
                            # hwp2.HAction.GetDefault("InsertText", hwp2.HParameterSet.HInsertText.HSet)
                            # hwp2.HParameterSet.HInsertText.Text = str(quiz_num)+" "
                            quiz_num += 1
                            # hwp2.HAction.Execute("InsertText", hwp2.HParameterSet.HInsertText.HSet)

                            # hwp2.Run("MoveLineBegin")
                            # hwp2.Run("Select")
                            # hwp2.Run("MoveSelNextWord")
                            # hwp2.HAction.GetDefault("CharShape", hwp2.HParameterSet.HCharShape.HSet)
                            # hwp2.HParameterSet.HCharShape.TextColor = 0x000000
                            # hwp2.HParameterSet.HCharShape.Height = 1600
                            # hwp2.HParameterSet.HCharShape.Bold = 1
                            # hwp2.HAction.Execute("CharShape", hwp2.HParameterSet.HCharShape.HSet)
                            # hwp2.Run("Cancel")
                            # hwp2.MovePos(3,1,1)
                            
                            #문제 출저 입력...
                            hwp2.HAction.GetDefault("InsertText", hwp2.HParameterSet.HInsertText.HSet)
                            hwp2.HParameterSet.HInsertText.Text = num_name
                            hwp2.HAction.Execute("InsertText", hwp2.HParameterSet.HInsertText.HSet)

                            hwp2.Run("Select")
                            hwp2.Run("MoveSelPrevWord")
                            # hwp2.HAction.GetDefault("CharShape", hwp2.HParameterSet.HCharShape.HSet)
                            # hwp2.HParameterSet.HCharShape.TextColor = 0x8B8B8B
                            # hwp2.HParameterSet.HCharShape.Height = 900
                            # hwp2.HParameterSet.HCharShape.FaceNameHangul ="바탕체"
                            # hwp2.HAction.Execute("CharShape", hwp2.HParameterSet.HCharShape.HSet)
                            # hwp2.Run("Cancel")
                            a.Execute(b)
                            hwp2.MovePos(3,1,1)
                            hwp2.Run("BreakPara")
                            hwp2.Run("DeleteBack")

                            hwp2.Run("Paste")
                        
                        else : #복수문제
                            if '~' in u_temp[l]:
                                if '주' in u_temp[l]: #주관식
                                    num_temp = int(u_temp[l].split('~')[1]) - int(u_temp[l].split('~')[0].split('주')[1]) + 1
                                    start_temp = int(u_temp[l].split('~')[0].split('주')[1])
                                    so = 1
                                else :
                                    num_temp = int(u_temp[l].split('~')[1]) - int(u_temp[l].split('~')[0]) + 1
                                    start_temp = int(u_temp[l].split('~')[0])
                                    so = 0
                            else :
                                if '주' in u_temp[l]: #주관식
                                    num_temp = int(u_temp[l].split('-')[1]) - int(u_temp[l].split('-')[0].split('주')[1]) + 1
                                    start_temp = int(u_temp[l].split('-')[0].split('주')[1])
                                    so = 1
                                else :
                                    num_temp = int(u_temp[l].split('-')[1]) - int(u_temp[l].split('-')[0]) + 1
                                    start_temp = int(u_temp[l].split('-')[0])
                                    so = 0
                                

                            #문제 번호 입력....
                            hwp2.HAction.GetDefault("InsertText", hwp2.HParameterSet.HInsertText.HSet)
                            hwp2.HParameterSet.HInsertText.Text = "["+str(quiz_num)+"~"+str(quiz_num + num_temp -1)+"] "
                            hwp2.HAction.Execute("InsertText", hwp2.HParameterSet.HInsertText.HSet)

                            hwp2.Run("MoveLineBegin")
                            hwp2.Run("Select")
                            hwp2.Run("MoveSelNextWord")
                            hwp2.HAction.GetDefault("CharShape", hwp2.HParameterSet.HCharShape.HSet)
                            hwp2.HParameterSet.HCharShape.TextColor = 0x000000
                            hwp2.HParameterSet.HCharShape.Height = 1000
                            hwp2.HParameterSet.HCharShape.Bold = 1
                            hwp2.HParameterSet.HCharShape.FaceNameHangul = "바탕체"
                            hwp2.HAction.Execute("CharShape", hwp2.HParameterSet.HCharShape.HSet)
                            hwp2.Run("Cancel")
                            hwp2.MovePos(3,1,1)
                            
                            #문제 출저 입력...
                            hwp2.HAction.GetDefault("InsertText", hwp2.HParameterSet.HInsertText.HSet)
                            hwp2.HParameterSet.HInsertText.Text = num_name
                            hwp2.HAction.Execute("InsertText", hwp2.HParameterSet.HInsertText.HSet)

                            hwp2.Run("Select")
                            hwp2.Run("MoveSelPrevWord")
                            hwp2.HAction.GetDefault("CharShape", hwp2.HParameterSet.HCharShape.HSet)
                            hwp2.HParameterSet.HCharShape.TextColor = 0x8B8B8B
                            hwp2.HParameterSet.HCharShape.Height = 900
                            hwp2.HParameterSet.HCharShape.FaceNameHangul ="바탕체"
                            hwp2.HAction.Execute("CharShape", hwp2.HParameterSet.HCharShape.HSet)
                            hwp2.Run("Cancel")
                            hwp2.MovePos(3,1,1)
                            hwp2.Run("BreakLine")

                            hwp2.Run("Paste")

                            for u in range(0, num_temp):
                                hwp.MovePos(2,1,1)
                                if so == 1:
                                    name_t = "주"+ str(start_temp+u)
                                else :
                                    name_t = str(start_temp+u)

                                num_name = '#'+ str(u_end-k+1)+'.'+str(self.subject_mf.currentIndex()+1)+'.'+name_t

                                find_set.SetItem("FindString",num_name)
                                find_set.SetItem("WholeWordOnly", True)
                                if find_act.Execute(find_set) == False:
                                    missing += num_name+"  "
                                    continue
                                if u == 0: # 처음에는 2번 검색
                                    if find_act.Execute(find_set) == False:
                                        missing += num_name+"  "
                                        continue

                                start_pos = hwp.GetPosBySet().Item("Para") + 1
                                find_set.SetItem("WholeWordOnly", False)

                                find_set.SetItem("FindString","#"+str(u_end-k+1)+'.'+str(self.subject_mf.currentIndex()+1))
                                find_act.Execute(find_set)

                                if l == len(u_temp)-1 and u == num_temp-1:
                                    find_set.SetItem("FindString","#")
                                    find_act.Execute(find_set)
                                    end_pos = hwp.GetPosBySet().Item("Para")
                                else:
                                    find_set.SetItem("FindString","#"+str(u_end-k+1)+'.'+str(self.subject_mf.currentIndex()+1))
                                    find_act.Execute(find_set)
                                    end_pos = hwp.GetPosBySet().Item("Para")

                                hwp.SelectText(start_pos,0,end_pos,0)
                                hwp.Run('Copy')

                                # #문제 번호 입력....
                                # hwp2.HAction.GetDefault("InsertText", hwp2.HParameterSet.HInsertText.HSet)
                                # hwp2.HParameterSet.HInsertText.Text = str(quiz_num)+" "
                                quiz_num += 1
                                # hwp2.HAction.Execute("InsertText", hwp2.HParameterSet.HInsertText.HSet)

                                # hwp2.Run("MoveLineBegin")
                                # hwp2.Run("Select")
                                # hwp2.Run("MoveSelNextWord")
                                # hwp2.HAction.GetDefault("CharShape", hwp2.HParameterSet.HCharShape.HSet)
                                # hwp2.HParameterSet.HCharShape.TextColor = 0xFF0000
                                # hwp2.HParameterSet.HCharShape.Height = 1600
                                # hwp2.HParameterSet.HCharShape.Bold = 1
                                # hwp2.HAction.Execute("CharShape", hwp2.HParameterSet.HCharShape.HSet)
                                # hwp2.Run("Cancel")
                                # hwp2.MovePos(3,1,1)
                                # hwp2.Run("BreakLine")
                                hwp2.HAction.GetDefault("InsertText", hwp2.HParameterSet.HInsertText.HSet)
                                hwp2.HParameterSet.HInsertText.Text = num_name
                                hwp2.HAction.Execute("InsertText", hwp2.HParameterSet.HInsertText.HSet)

                                hwp2.Run("Select")
                                hwp2.Run("MoveSelPrevWord")
                                a.Execute(b)
                                hwp2.MovePos(3,1,1)
                                hwp2.Run("BreakPara")
                                hwp2.Run("DeleteBack")

                                hwp2.Run("Paste")

        else:
          
            u_num = subjects[self.subject_.currentIndex()-1].num    #단원 수
            u_list = subjects[self.subject_.currentIndex()-1].list   # 단원 이름 리스트
            u_start = subjects[self.subject_.currentIndex()-1].start # 단원 시작년도
            u_end = subjects[self.subject_.currentIndex()-1].end     # " 끝 년도
            u_quiz = subjects[self.subject_.currentIndex()-1].quiz   # 문제번호 배열
        


            quiz_num =1
            missing =""


            statenum = 0
            for i in range(0, u_num): #단원
                for k in range(1, u_end - u_start+2): #연도
                    if self.current.cellWidget(i,k).isChecked() == True:
                        statenum +=1
            now =0
            state.showMessage('진행 중......' +str(now) +"/"+str(statenum))
            state.repaint()


            for i in range(0, u_num): #단원

                
                

                # write_act =  hwp.CreateAction("InsertText")
                # write_set = write_act.CreateSet()
                # write_act.GetDefault(write_set)
                # write_set.SetItem("InsertText","1111")
                # write_act.Execute(write_set)
                hwp2_flag = 0


                

                find_act = hwp.CreateAction("ForwardFind")
                find_set = find_act.CreateSet()
                find_act.GetDefault(find_set)

                for k in range(1, u_end - u_start+2): #연도

                    if self.current.cellWidget(i,k).isChecked() == False:
                        continue
                    
                    now += 1
                    state.showMessage('진행 중......' +str(now) +"/"+str(statenum))
                    state.repaint()

                    #단원명 블럭 만들기
                    if hwp2_flag == 0 :
                        hwp2_flag = 1
                        hwp2.SelectText(0,48,0,56)
                        hwp2.Run("Copy")
                        hwp2.MovePos(3,1,1)
                        hwp2.Run("BreakPara")
                        hwp2.Run("DeleteBack")
                        hwp2.Run("BreakPara")
                        hwp2.Run("Paste")
                        hwp2.Run("MoveLeft")
                        hwp2.HAction.GetDefault("InsertText", hwp2.HParameterSet.HInsertText.HSet)
                        hwp2.HParameterSet.HInsertText.Text = u_list[i]
                        hwp2.HAction.Execute("InsertText", hwp2.HParameterSet.HInsertText.HSet)
                        hwp2.MovePos(3,1,1)
                        hwp2.Run("BreakPara")
                        hwp2.Run("BreakPara")
                        hwp2.Run("DeleteBack")

                
                

                    for l in range (0, len(u_quiz[k-1][i])): #한 체크박스
                        hwp.MovePos(2,1,1)
                        num_name = '#' + str(u_end-k+1)+'.' + u_quiz[k-1][i][l].name

                        find_set.SetItem("FindString",num_name+"~")

                        if find_act.Execute(find_set) == True: # ~묶음 문제~
                            hwp.Run("Select")
                            hwp.Run("Copy")
                            root = Tk()
                            root.withdraw()
                            full = str(root.clipboard_get()).split(".")[2]

                            if '주' in full.split('~')[0]:
                                gae = int(full.split('~')[1]) - int(full.split('~')[0][1:])
                            else:
                                gae = int(full.split('~')[1]) - int(full.split('~')[0])
                            
                            start_pos = hwp.GetPosBySet().Item("Para") + 1

                            find_set.SetItem("FindString","#"+str(u_end-k+1)+'.')
                            find_act.Execute(find_set)

                            end_pos = hwp.GetPosBySet().Item("Para")
                            hwp.SelectText(start_pos,0,end_pos,0)
                            hwp.Run('Copy')
                            #문제 번호 입력....
                            hwp2.HAction.GetDefault("InsertText", hwp2.HParameterSet.HInsertText.HSet)
                            hwp2.HParameterSet.HInsertText.Text = "["+str(quiz_num)+"~"+str(quiz_num + gae)+"]"
                            hwp2.HAction.Execute("InsertText", hwp2.HParameterSet.HInsertText.HSet)

                            hwp2.Run("MoveLineBegin")
                            hwp2.Run("Select")
                            hwp2.Run("MoveSelNextWord")
                            hwp2.HAction.GetDefault("CharShape", hwp2.HParameterSet.HCharShape.HSet)
                            hwp2.HParameterSet.HCharShape.TextColor = 0x000000
                            hwp2.HParameterSet.HCharShape.Height = 1000
                            hwp2.HParameterSet.HCharShape.Bold = 1
                            hwp2.HParameterSet.HCharShape.FaceNameHangul = "바탕체"
                            hwp2.HAction.Execute("CharShape", hwp2.HParameterSet.HCharShape.HSet)
                            hwp2.Run("Cancel")
                            hwp2.MovePos(3,1,1)
                            hwp2.Run("BreakLine")

                            hwp2.Run("Paste")
                        else :
                            hwp.MovePos(2,1,1)
                            find_set.SetItem("FindString",num_name+"-")
                            if find_act.Execute(find_set) == True: # -묶음 문제~
                                hwp.Run("Select")
                                hwp.Run("Copy")
                                root = Tk()
                                root.withdraw()
                                full = str(root.clipboard_get()).split(".")[2]

                                if '주' in full.split('-')[0]:
                                    gae = int(full.split('-')[1]) - int(full.split('-')[0][1:])
                                else:
                                    gae = int(full.split('-')[1]) - int(full.split('-')[0])
                                
                                start_pos = hwp.GetPosBySet().Item("Para") + 1

                                find_set.SetItem("FindString","#"+str(u_end-k+1)+'.')
                                find_act.Execute(find_set)

                                end_pos = hwp.GetPosBySet().Item("Para")
                                hwp.SelectText(start_pos,0,end_pos,0)
                                hwp.Run('Copy')
                                #문제 번호 입력....
                                hwp2.HAction.GetDefault("InsertText", hwp2.HParameterSet.HInsertText.HSet)
                                hwp2.HParameterSet.HInsertText.Text = "["+str(quiz_num)+"~"+str(quiz_num + gae)+"]"
                                hwp2.HAction.Execute("InsertText", hwp2.HParameterSet.HInsertText.HSet)

                                hwp2.Run("MoveLineBegin")
                                hwp2.Run("Select")
                                hwp2.Run("MoveSelNextWord")
                                hwp2.HAction.GetDefault("CharShape", hwp2.HParameterSet.HCharShape.HSet)
                                hwp2.HParameterSet.HCharShape.TextColor = 0x000000
                                hwp2.HParameterSet.HCharShape.Height = 1000
                                hwp2.HParameterSet.HCharShape.Bold = 1
                                hwp2.HParameterSet.HCharShape.FaceNameHangul = "바탕체"
                                hwp2.HAction.Execute("CharShape", hwp2.HParameterSet.HCharShape.HSet)
                                hwp2.Run("Cancel")
                                hwp2.MovePos(3,1,1)
                                hwp2.Run("BreakLine")

                                hwp2.Run("Paste")
                            hwp.MovePos(2,1,1)







                        
                        find_set.SetItem("FindString",num_name)
                        find_set.SetItem("WholeWordOnly", True)
                        if find_act.Execute(find_set) == False:
                            missing += num_name+"  "
                            continue
                        
                        start_pos = hwp.GetPosBySet().Item("Para") + 1
                        find_set.SetItem("WholeWordOnly", False)


                        if l == len(u_quiz[k-1][i])-1:
                            find_set.SetItem("FindString","#")
                            find_act.Execute(find_set)

                            end_pos = hwp.GetPosBySet().Item("Para")

                        else:

                            find_set.SetItem("FindString","#"+str(u_end-k+1)+'.')
                            find_act.Execute(find_set)

                            end_pos = hwp.GetPosBySet().Item("Para")

                        hwp.SelectText(start_pos,0,end_pos,0)
                        hwp.Run('Copy')


                        #문제 번호 입력....
                        # hwp2.HAction.GetDefault("InsertText", hwp2.HParameterSet.HInsertText.HSet)
                        # hwp2.HParameterSet.HInsertText.Text = str(quiz_num)+" "
                        quiz_num += 1
                        # hwp2.HAction.Execute("InsertText", hwp2.HParameterSet.HInsertText.HSet)

                        # hwp2.Run("MoveLineBegin")
                        # hwp2.Run("Select")
                        # hwp2.Run("MoveSelNextWord")
                        # hwp2.HAction.GetDefault("CharShape", hwp2.HParameterSet.HCharShape.HSet)
                        # hwp2.HParameterSet.HCharShape.TextColor = 0x000000
                        # hwp2.HParameterSet.HCharShape.Height = 1600
                        # hwp2.HParameterSet.HCharShape.Bold = 1
                        # hwp2.HAction.Execute("CharShape", hwp2.HParameterSet.HCharShape.HSet)
                        # hwp2.Run("Cancel")
                        # hwp2.MovePos(3,1,1)
                        
                        #문제 출저 입력...
                        hwp2.HAction.GetDefault("InsertText", hwp2.HParameterSet.HInsertText.HSet)
                        hwp2.HParameterSet.HInsertText.Text = num_name
                        hwp2.HAction.Execute("InsertText", hwp2.HParameterSet.HInsertText.HSet)

                        hwp2.Run("Select")
                        hwp2.Run("MoveSelPrevWord")
                        # hwp2.HAction.GetDefault("CharShape", hwp2.HParameterSet.HCharShape.HSet)
                        # hwp2.HParameterSet.HCharShape.TextColor = 0x8B8B8B
                        # hwp2.HParameterSet.HCharShape.Height = 900
                        # hwp2.HParameterSet.HCharShape.FaceNameHangul ="바탕체"
                        # hwp2.HAction.Execute("CharShape", hwp2.HParameterSet.HCharShape.HSet)
                        # hwp2.Run("Cancel")
                        a.Execute(b)
                        hwp2.MovePos(3,1,1)
                            

                        hwp2.HAction.GetDefault("InsertText", hwp2.HParameterSet.HInsertText.HSet)
                        hwp2.HParameterSet.HInsertText.Text = " "+u_quiz[k-1][i][l].pro
                        hwp2.HAction.Execute("InsertText", hwp2.HParameterSet.HInsertText.HSet)

                        # hwp2.Run("Select")
                        # hwp2.Run("MoveSelPrevWord")
                        # hwp2.HAction.GetDefault("CharShape", hwp2.HParameterSet.HCharShape.HSet)
                        # hwp2.HParameterSet.HCharShape.TextColor = 0x8B8B8B
                        # hwp2.HParameterSet.HCharShape.Height = 900
                        # hwp2.HParameterSet.HCharShape.FaceNameHangul ="바탕체"
                        # hwp2.HAction.Execute("CharShape", hwp2.HParameterSet.HCharShape.HSet)
                        # hwp2.Run("Cancel")
                        # hwp2.MovePos(3,1,1)
                        hwp2.Run("BreakPara")
                        hwp2.Run("DeleteBack")

                        hwp2.Run("Paste")
                        

                        

                            



        hwp2.MovePos(3,1,1)
        hwp2.Run("BreakPage")
        hwp2.SelectText(0,48,1,1)
        hwp2.Run("Delete")
        hwp2.Run("Delete")
        hwp2.Run("DeleteBack")

        
        win32gui.ShowWindow(h2, win32con.SW_SHOW)
        hwp.Quit()
        state.showMessage('Ready')

        if missing != "":
            hwp2.MovePos(3,1,1)
            hwp2.HAction.GetDefault("InsertText", hwp2.HParameterSet.HInsertText.HSet)
            hwp2.HParameterSet.HInsertText.Text = missing+"존재하지 않습니다."
            hwp2.HAction.Execute("InsertText", hwp2.HParameterSet.HInsertText.HSet)
            
                




        



       




    def refresh(self):
        global subjects
        global state

        if os.path.isfile(os.path.join( os.getcwd(),"data/savedata.p") ):
            del subjects[:]
            with open(os.path.join(os.getcwd(),"data/savedata.p"), 'rb') as file:    # 저장하기
                subjects = pickle.load(file)
            self.subject_.clear()
            self.subject_.addItem("과목 선택")
            for i in range(0, len(subjects)):
                self.subject_.addItem(subjects[i].name)
            
            return
            


        file_list1 = os.listdir(os.path.join( os.getcwd(),"excel"))

        
        file_list = [v for v in file_list1 if not v.startswith('~$')]
        print (file_list)

        file_num = len(file_list) # 엑셀 파일 갯수
        del subjects[:] # 과목리스트 초기화

        now =0
        state.showMessage('진행 중......' +str(now) +"/"+str(file_num))
        state.repaint()

        excel = win32.Dispatch("Excel.Application")

        for i in range(0, file_num):
            now +=1
            state.showMessage('진행 중......' +str(now) +"/"+str(file_num))
            state.repaint()

            starttime = time.time()

            file_name = str(file_list[i]) 
            subjects.append(subject(file_name))

            excelfile = excel.Workbooks.Open(os.path.join( os.getcwd() ,"excel/"+file_name))
            #####

            sheetnum = excelfile.Worksheets.Count
            type2 =0
            for h in range(1, sheetnum+1):
                if "단원명"==str(excelfile.Worksheets(h).Name).replace(" ",""):
                    type2 =1
                    break
            
            if type2 != 1: #기생충학 etc
                subjects[i].type = 1

                work1 = excelfile.Worksheets('중간고사')
                work2 = excelfile.Worksheets('기말고사')

                # 중간고사
                subjects[i].mid_end = int(work1.Cells(1,3))%100
                temp = 3
                while True:
                    temp += 1
                    if(str(work1.Cells(1,temp)).strip() in ["None", ""]):
                        subjects[i].mid_start = int(work1.Cells(1,temp-1))%100
                        break
                

                temp = 2
                while True:
                    if(str(work1.Cells(temp,2)).strip() in ["None", ""]): # 리스트 끝
                        subjects[i].mid_num = temp-2
                        break
                    subjects[i].mid_num += 1
                    subjects[i].mid_list.append(str(work1.Cells(temp,2)).strip()) #단원 이름 추가
                    temp_list=[]
                    for l in range(3, subjects[i].mid_end - subjects[i].mid_start+4):
                        if type(work1.Cells(temp,l).Value) is not float: 
                            temp_list.append(str(work1.Cells(temp,l)).replace(" ",""))
                        else :
                            temp_list.append(str(int(work1.Cells(temp,l).Value)))
                        

                    subjects[i].mid_quiz.append(temp_list)
                    temp += 1
                    

                # 기말고사
                subjects[i].f_end = int(work2.Cells(1,3))%100
                temp = 3
                while True:
                    temp += 1
                    if(str(work2.Cells(1,temp)).strip() in ["None", ""]):
                        subjects[i].f_start = int(work2.Cells(1,temp-1))%100
                        break
                

                temp = 2
                while True:
                    if(str(work2.Cells(temp,2)).strip() in ["None", ""]): # 리스트 끝
                        subjects[i].f_num = temp-2
                        break
                    subjects[i].f_num += 1
                    subjects[i].f_list.append(str(work2.Cells(temp,2)).strip()) #단원 이름 추가
                    temp_list=[]
                    for l in range(3, subjects[i].f_end - subjects[i].f_start+4):
                        if type(work2.Cells(temp,l).Value) is not float: 
                            temp_list.append(str(work2.Cells(temp,l)).replace(" ",""))
                        else :
                            temp_list.append(str(int(work2.Cells(temp,l).Value)))
                    subjects[i].f_quiz.append(temp_list)
                    temp += 1
            
            else :
                
                subjects[i].type = 2
                # 여기부터 다름
                work1 = excelfile.Worksheets('문제')
                work2 = excelfile.Worksheets('단원명')

                #단원 시트
                for l in range (1, 1000):
                    if(str(work2.Cells(l,1)).strip() in ["None", ""]):
                        subjects[i].num = l -1
                        break
                    
                    subjects[i].list.append(str(work2.Cells(l,1)).replace(" ",""))
                
                #문제 시트
                for l in range (2, 100000, 2): # l은 열
                    year_c = str(work1.Cells(1,l)).strip()
                    if(year_c in ["None", ""]):
                        break
                    
                    year_ = year_c.split(" ")[0]
                   
                    if l == 2:
                        subjects[i].end = int(year_)
                        subjects[i].start = int(year_)

                    if subjects[i].start > int(year_):
                        subjects[i].start = int(year_)

                subjects[i].quiz = [[[] for col in range(subjects[i].num)] for row in range( subjects[i].end -subjects[i].start +1 )]
            
                

                for l in range (2, 100000, 2): # l은 열
                    year_c = str(work1.Cells(1,l)).strip()
                    if(year_c in ["None", ""]):
                        break

                    year_ = int(year_c.split(" ")[0])
                    cha = year_c.split(" ")[1][0]


                    for k in range(3, 10000): # k는 행
                        if(str(work1.Cells(k,l+1)).strip() in ["None", ""]):
                            if k <201:
                                for kk in range(201, 10000):# 주관식
                                    if(str(work1.Cells(kk,l+1)).strip() in ["None", ""]):
                                        break
                                    if str(work1.Cells(kk,l+1)).replace(" ","") not in subjects[i].list:
                                        if l+ 65 > 90 :
                                            string = file_name+"의 "+ str(chr(int(l/26) + 64))+str(chr(l%26 + 65))+str(kk)+" 셀을 확인하세요!\n단원목록에 없습니다."
                                        else:
                                            string = file_name+"의 "+str(chr(l+ 65))+str(kk)+" 셀을 확인하세요!\n단원목록에 없습니다."
                                        excel.Quit()
                                        ctypes.windll.user32.MessageBoxW(None, string, "Warning!", 0)
                                        sys.exit(0)
                                    index = subjects[i].list.index(str(work1.Cells(kk,l+1)).replace(" ",""))
                                    subjects[i].quiz[subjects[i].end - year_ ][index].append(quiz_(cha+".주"+str(kk-200) ,str(work1.Cells(kk,l)).strip()))
                                    
                            
                                break
                 
                        if str(work1.Cells(k,l+1)).replace(" ","") not in subjects[i].list:
                            if l+ 65 > 90 :
                                string = file_name+"의 "+ str(chr(int(l/26) + 64))+str(chr(l%26 + 65))+str(k)+" 셀을 확인하세요!\n단원목록에 없습니다."
                            else:
                                string = file_name+"의 "+str(chr(l+ 65))+str(k)+" 셀을 확인하세요!\n단원목록에 없습니다."
                            
                            excel.Quit()
                            ctypes.windll.user32.MessageBoxW(None, string, "Warning!", 0)
                            sys.exit(0)

                        index = subjects[i].list.index(str(work1.Cells(k,l+1)).replace(" ",""))
                        subjects[i].quiz[subjects[i].end - year_ ][index].append(quiz_(cha+"."+str(k-2) ,str(work1.Cells(k,l)).strip()))
            print("time : ",time.time()-starttime)           
        excel.Quit()
        state.showMessage('Ready')
        state.repaint()

        with open(os.path.join(os.getcwd(),"data/savedata.p"), 'wb') as file:    # 저장하기
            pickle.dump(subjects, file)
        
        self.subject_.clear()
        self.subject_.addItem("과목 선택")
        for i in range(0, len(subjects)):
            self.subject_.addItem(subjects[i].name)
        
        


    def unit_change(self):
        if self.subject_.currentText() == "과목 선택":
            return

        if self.subject_.currentText() == "기생충학":
            self.current.hide()
            self.current = self.tables[self.subject_.currentIndex()-1][self.subject_mf.currentIndex()]
            self.current_name = subjects[self.subject_.currentIndex()-1].name
            self.current.show()
            self.subject_mf.show()
            self.close_mf.hide()
        
        else:
            self.current.hide()
            self.current = self.tables[self.subject_.currentIndex()-1]
            self.current_name = subjects[self.subject_.currentIndex()-1].name
            self.current.show()
            self.subject_mf.hide()
            self.close_mf.show()
    
    def job_check(self,state,row, k ):
        
        if state ==0:
            self.current.cellWidget(row,k).setCheckState(0)

    
    def all_check(self, state,row, k):
        
        if state ==2:
            for i in range(1, k):
                self.current.cellWidget(row, i).setCheckState(2)
        
        elif state ==0:
            for i in range(1, k):
                if self.current.cellWidget(row, i).checkState() !=2:
                    return
            for i in range(1, k):
                self.current.cellWidget(row, i).setCheckState(0)
    
    def refreshall(self):
        
        if os.path.isfile(os.path.join( os.getcwd(),"data/savedata.p") ):
            os.remove(os.path.join( os.getcwd(),"data/savedata.p"))
        self.refresh()

            


        

class subject:
    def __init__(self,name):
        self.name = name.split('.')[0]

        self.mid_num=0 #단원 수
        self.mid_list=[] # 단원 이름 리스트
        self.mid_start=0 # 단원 시작년도
        self.mid_end=0   # " 끝 년도
        self.mid_quiz=[] # 문제번호 배열

        self.type=0


        self.f_num=0
        self.f_list=[]
        self.f_start=0
        self.f_end=0
        self.f_quiz=[]

        self.num=0
        self.list=[]
        self.start=0
        self.end =0
        self.quiz =[]

class quiz_:
    def __init__(self,name,pro ):
        self.name = name
        self.pro = pro
        





if __name__ == '__main__':
    
    

    app = QApplication(sys.argv)
    ex = MyApp()
    ex.show()
    sys.exit(app.exec_())








