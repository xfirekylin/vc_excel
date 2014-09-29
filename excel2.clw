; CLW file contains information for the MFC ClassWizard

[General Info]
Version=1
LastClass=CPara2
LastTemplate=CDialog
NewFileInclude1=#include "stdafx.h"
NewFileInclude2=#include "excel2.h"

ClassCount=6
Class1=CExcel2App
Class2=CExcel2Dlg
Class3=CAboutDlg

ResourceCount=5
Resource1=IDD_EXCEL2_DIALOG
Resource2=IDR_MAINFRAME
Class4=ExcelTab
Resource3=IDD_PARA1
Class5=CPara1
Resource4=IDD_ABOUTBOX
Class6=CPara2
Resource5=IDD_PARA2

[CLS:CExcel2App]
Type=0
HeaderFile=excel2.h
ImplementationFile=excel2.cpp
Filter=N

[CLS:CExcel2Dlg]
Type=0
HeaderFile=excel2Dlg.h
ImplementationFile=excel2Dlg.cpp
Filter=D
BaseClass=CDialog
VirtualFilter=dWC
LastObject=IDC_TAB1

[CLS:CAboutDlg]
Type=0
HeaderFile=excel2Dlg.h
ImplementationFile=excel2Dlg.cpp
Filter=D

[DLG:IDD_ABOUTBOX]
Type=1
Class=CAboutDlg
ControlCount=4
Control1=IDC_STATIC,static,1342177283
Control2=IDC_STATIC,static,1342308480
Control3=IDC_STATIC,static,1342308352
Control4=IDOK,button,1342373889

[DLG:IDD_EXCEL2_DIALOG]
Type=1
Class=CExcel2Dlg
ControlCount=1
Control1=IDC_TAB1,SysTabControl32,1342177280

[CLS:ExcelTab]
Type=0
HeaderFile=ExcelTab.h
ImplementationFile=ExcelTab.cpp
BaseClass=CDialog
Filter=D
LastObject=ExcelTab

[DLG:IDD_PARA1]
Type=1
Class=CPara1
ControlCount=13
Control1=IDOK,button,1342242817
Control2=IDC_STATIC,static,1342308352
Control3=IDC_EDIT1,edit,1350631552
Control4=IDC_STATIC,static,1342308352
Control5=IDC_EDIT2,edit,1350639744
Control6=IDC_STATIC,static,1342308352
Control7=IDC_STATIC,static,1342308352
Control8=IDC_EDIT3,edit,1350639744
Control9=IDC_EDIT4,edit,1350639744
Control10=IDC_STATIC,static,1342308352
Control11=IDC_EDIT5,edit,1350631552
Control12=IDC_STATIC,static,1342308352
Control13=IDC_EDIT6,edit,1350631552

[CLS:CPara1]
Type=0
HeaderFile=Para1.h
ImplementationFile=Para1.cpp
BaseClass=CDialog
Filter=D
VirtualFilter=dWC
LastObject=IDC_EDIT1

[DLG:IDD_PARA2]
Type=1
Class=CPara2
ControlCount=33
Control1=IDC_EDIT1,edit,1350631552
Control2=IDC_EDIT2,edit,1350631552
Control3=IDC_EDIT3,edit,1350631552
Control4=IDC_EDIT4,edit,1350631552
Control5=IDC_EDIT5,edit,1350631552
Control6=IDC_EDIT6,edit,1350631552
Control7=IDC_EDIT7,edit,1350631552
Control8=IDC_EDIT8,edit,1350631552
Control9=IDC_EDIT13,edit,1350631552
Control10=IDC_EDIT14,edit,1350631552
Control11=IDC_EDIT9,edit,1350631552
Control12=IDC_EDIT10,edit,1350631552
Control13=IDC_EDIT11,edit,1350631552
Control14=IDC_EDIT12,edit,1350631552
Control15=IDC_CHECK1,button,1342242819
Control16=IDC_CHECK2,button,1342242819
Control17=IDOK,button,1342242817
Control18=IDC_STATIC,static,1342308352
Control19=IDC_STATIC,static,1342308352
Control20=IDC_STATIC,static,1342308352
Control21=IDC_STATIC,static,1342308352
Control22=IDC_STATIC,static,1342308352
Control23=IDC_STATIC,static,1342308352
Control24=IDC_STATIC,static,1342308352
Control25=IDC_STATIC,static,1342308352
Control26=IDC_STATIC,static,1342308352
Control27=IDC_STATIC,static,1342308352
Control28=IDC_STATIC,static,1342308352
Control29=IDC_STATIC,static,1342308352
Control30=IDC_STATIC,static,1342308352
Control31=IDC_STATIC,static,1342308352
Control32=IDC_STATIC,static,1342308352
Control33=IDC_STATIC,static,1342308352

[CLS:CPara2]
Type=0
HeaderFile=Para2.h
ImplementationFile=Para2.cpp
BaseClass=CDialog
Filter=D
VirtualFilter=dWC
LastObject=IDC_EDIT1

