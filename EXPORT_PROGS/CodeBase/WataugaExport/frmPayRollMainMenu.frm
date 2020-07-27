VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmPayrollMainMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "v 2.05 Payroll Main Menu"
   ClientHeight    =   8880
   ClientLeft      =   30
   ClientTop       =   690
   ClientWidth     =   11655
   FillColor       =   &H8000000B&
   Icon            =   "frmPayRollMainMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   360
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   360
      Top             =   360
   End
   Begin fpBtnAtlLibCtl.fpBtn controlFileMainCmmd 
      Height          =   495
      Left            =   4005
      TabIndex        =   4
      Top             =   4995
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmPayRollMainMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn reportsProcCmmd 
      Height          =   495
      Left            =   4005
      TabIndex        =   3
      Top             =   4230
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmPayRollMainMenu.frx":0AB6
   End
   Begin fpBtnAtlLibCtl.fpBtn payRollProcCmmd 
      Height          =   495
      Left            =   4005
      TabIndex        =   2
      Top             =   3465
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmPayRollMainMenu.frx":0C9C
   End
   Begin fpBtnAtlLibCtl.fpBtn EmpFileMaintCmmd 
      Height          =   495
      Left            =   4005
      TabIndex        =   1
      Top             =   2700
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmPayRollMainMenu.frx":0E82
   End
   Begin fpBtnAtlLibCtl.fpBtn W2ProcessingCmmd 
      Height          =   495
      Left            =   4005
      TabIndex        =   5
      Top             =   5745
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmPayRollMainMenu.frx":106F
   End
   Begin fpBtnAtlLibCtl.fpBtn exitCmd 
      Height          =   495
      Left            =   4005
      TabIndex        =   6
      Top             =   6480
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmPayRollMainMenu.frx":1254
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1098
      Index           =   1
      Left            =   1500
      Top             =   897
      Width           =   8655
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   2101
      Top             =   2103
      Width           =   971
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8710.757
      X2              =   8710.757
      Y1              =   2151.243
      Y2              =   7892.757
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   8593
      Top             =   2101
      Width           =   971
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8710.757
      X2              =   9412.576
      Y1              =   7884.973
      Y2              =   7884.973
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2205.432
      X2              =   2920.248
      Y1              =   7884.973
      Y2              =   7884.973
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2220.428
      X2              =   2220.428
      Y1              =   2151.243
      Y2              =   7880.108
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5910
      Index           =   0
      Left            =   2220
      Top             =   2201
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8592
      Top             =   1971
      Width           =   972
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PAYROLL MAIN MENU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2820
      TabIndex        =   0
      Top             =   1248
      Width           =   6012
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1214
      Left            =   1502
      Top             =   771
      Width           =   8652
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5910
      Index           =   1
      Left            =   8713
      Top             =   2201
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   2101
      Top             =   1970
      Width           =   975
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Sosoft Options"
      Begin VB.Menu mnuRelink 
         Caption         =   "Relink Transactions"
      End
      Begin VB.Menu mnuReindex 
         Caption         =   "Reindex Employees"
      End
   End
End
Attribute VB_Name = "frmPayrollMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class

'Private Sub cmdReindex_Click()
'  Call MakeEmpIndexs
'  MsgBox "Re-indexing has completed successfully."
'End Sub
'
'Private Sub cmdRelink_Click()
'  frmPRRelinkTransHist.Show
'  DoEvents
'  Unload frmPayrollMainMenu
'End Sub

Private Sub controlFileMainCmmd_Click()
'  Dim SysRec As RegDSysFileRecType
'  Dim SysCnt As Integer
'  Dim SysHandle As Integer
'
'  OpenSysFile SysHandle
'  SysCnt = LOF(SysHandle) \ Len(SysRec)
'  Get SysHandle, 1, SysRec
'  Close SysHandle
  'this is an alert to the user that there are
  'problems with the current citipak path
'  If SysCnt <> 0 Then
'    CurrCitiPath = QPTrim$(SysRec.CITIDIR)
'    If QPTrim$(SysRec.CITIDIR) <> "" Then '11/27/02 for users with no GL
'      If CheckCitiDir(QPTrim$(SysRec.CITIDIR)) = 0 Then
'        frmWarnOpenWNoDir.Show vbModal, Me
'      End If
'    End If
'  End If
  
  If LevelPass = 0 Then
    MsgBox "Currently you are denied access to this section. Exit from payroll and log in again through the Citipak login screen."
    Exit Sub
  ElseIf LevelPass = 1 Then
    frmControlFileMaint.Show
    DoEvents
    Unload frmPayrollMainMenu
  Else
    MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
  End If
    
End Sub

Private Sub EmpFileMaintCmmd_Click()
'  Dim SysRec As RegDSysFileRecType
'  Dim SysCnt As Integer
'  Dim SysHandle As Integer
'
'  OpenSysFile SysHandle
'  SysCnt = LOF(SysHandle) \ Len(SysRec)
'  Get SysHandle, 1, SysRec
'  Close SysHandle
'  'this is an alert to the user that there are
'  'problems with the current citipak path
'  If SysCnt <> 0 Then
'    CurrCitiPath = QPTrim$(SysRec.CITIDIR)
'    If QPTrim$(SysRec.CITIDIR) <> "" Then '11/27/02 for users with no GL
'      If CheckCitiDir(QPTrim$(SysRec.CITIDIR)) = 0 Then
'        frmWarnOpenWNoDir.Show vbModal, Me
'      End If
'    End If
'  End If
  
  If LevelPass = 0 Then
    MsgBox "Currently you are denied access to this section. Exit from payroll and log in again through the Citipak login screen."
    Exit Sub
  ElseIf LevelPass = 1 Then
    frmEmployeeMaintMenu.Show
    DoEvents
    Unload frmPayrollMainMenu
  Else
    MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
  End If
  
End Sub

Private Sub exitCmd_Click()
  If Exist("prmain.dat") Then
    KillFile "prmain.dat"
  End If
  
  MainLog ("Payroll.exe terminated via normal exit in Payroll Main Menu.")
  DoEvents
  Call Ready4others(PWcnt)
  DoEvents
  If Exist(QPTrim$(StartPath) + "\" + "Citipak.exe") Then
    Shell QPTrim$(StartPath) + "\" + "Citipak.exe", vbMaximizedFocus
  End If
  
  Timer1.Enabled = True
  
End Sub

Private Sub Form_Load()
  Dim FirstThru As Boolean
  Dim cnt&, dl&
  Dim SysRec As RegDSysFileRecType
  Dim SysHandle As Integer
  Dim SysCnt
  Dim UnitRec As UnitFileRecType
  Dim UHandle As Integer
  Dim ThisDir$
  Dim PayRate As PayRateType
  Dim PHandle As Integer
  Dim NumOfPayRate As Integer
  Dim One As Integer
  Dim AHandle As Integer
  Dim CitiPassFile As Integer
  Dim CitiPass As CitiPassType
  Dim NumPassRecs As Integer
  
'  Call FixTransEmpPins
'  Call FixElkton
  ResetAllToPrintAdYes
  Me.HelpContextID = hlpPRMain
  One = 1
  AHandle = FreeFile
  Open "prmain.dat" For Output As AHandle
  Print #AHandle, One
  Close AHandle
  
  If Exist("taxmain.dat") Then
    FromPR = True
    W2ProcessingCmmd.Enabled = False 'Timer2 re-enables it automatically after 2 seconds
  ElseIf FromPR = False Then
    FromPR = True
    exitCmd.Enabled = False 'Timer2 re-enables it automatically after 2 seconds
  End If
  
  CurrCitiPath = App.Path
  If Mid(CurrCitiPath, Len(CurrCitiPath), 1) <> "\" Then
    CurrCitiPath = CurrCitiPath + "\"
  End If
  
  If PWcnt > 0 Then
    OpenCitiPassFile CitiPassFile, NumPassRecs
    Get CitiPassFile, PWcnt, CitiPass
    Close CitiPassFile
  End If
  
  Clipboard.Clear
  If App.PrevInstance Then
    ActivatePrevInstance 'don't want two payroll
    'programs open at once
  End If
  '1) checks to see if prunit.dat exists
  '2) if it does then if the file length is 381 then
  '   this file has been converted and is ready to go
  '3) if it doesn't exist then the program continues to
  '   convertdone because this might be the first time
  '   the program has been accessed
  '4) if the file length is not 381 then this data is
  '   either windows compatible but has not been converted
  '   to the latest version or this data is still in the
  '   original dos format
  '5) checks to see if prsys.dat exists
  '6) if it does then if the file size is 337 then this
  '   data is in the original dos format.
  '7) if it does and the file size is 340 then this
  '   data has been converted to windows but not the latest
  '   version
  '8) if prsys.dat does not exist the program continues
  '   as if conversion were done because this might be
  '   the first time payroll has been accessed
  '9) if this data is dos data then the program moves to
  '   dostowin and opens a shell to the program that
  '   converts that data to the newest version of windows
  '10) if this data is windows but not the latest version
  '   then the program moves to wintowin and opens a shell
  '   to the program that converts this type of data to
  '   the latest version of windows
  '11) if the latest conversion has already taken place then
  '   the program moves to convertdone and kills the two
  '   conversion .exe programs so they cannot be accessed
  '   anymore...converting data that has already been
  '   converted will destroy the data
  '12) Inserted for Fall 04 update...looks for the correct
  '   size of the prunit.dat file. If it is 381 then that means
  '   prdata has not been updated. The program proceeds to that
  '   code that will update the files.
  
  If Exist("PRDATA\PRUNIT.DAT") Then '1)
    If FileLen("PRDATA\PRUNIT.DAT") = 398 Then ' changed to 398 with Fall 04 update '2)
      GoTo ConvertDone '11)
    Else '4)
      If FileLen("PRDATA\PRUNIT.DAT") = 381 Then '12
         GoTo Fall04Update
      ElseIf Exist("PRDATA\PRSYS.DAT") Then '5)
        If FileLen("PRDATA\PRSYS.DAT") = 337 Then '6)
          GoTo DosToWin '9)
        ElseIf FileLen("PRDATA\PRSYS.DAT") = 340 Then '7)
          GoTo WinToWin '11)
        End If
      Else '8)
        frmMessage.Label1.Caption = "This PRData is not recognizable. Loading will continue but data will be suspect."
        frmMessage.Label1.Top = 850
        frmMessage.Show vbModal
        MainLog ("***WARNING***User warned that this PRData could not be recognized.")
        GoTo ConvertDone
      End If
    End If
  Else '3)
    frmMessage.Label1.Caption = "This PRData is not recognizable. Loading will continue but data will be suspect."
    frmMessage.Label1.Top = 850
    frmMessage.Show vbModal
    MainLog ("***WARNING***User warned that this PRData could not be recognized.")
    GoTo ConvertDone
  End If
  
DosToWin:
  If Exist("PRConvertDos2Win.exe") Then
    ClearInUsePRReg PWcnt  'added 11/02/2002
    Shell "PRConvertDos2Win.exe", vbMaximizedFocus
    DoEvents
    Unload frmPayrollMainMenu
    End
  Else
    MsgBox "Please load the Dos2Win conversion exe."
    Terminate
  End If
  
Fall04Update:
  If Exist("CnvtPRFall04.exe") Then
    Close
    ClearInUsePRReg PWcnt
    Shell "CnvtPRFall04.exe", vbMaximizedFocus
    DoEvents
    Unload frmPayrollMainMenu
    End
  Else
    frmMessage.Label1.Caption = "You are attempting to run the Fall 2004 Update version of payroll but conversion has not yet taken place. Please place a copy of 'CnvtPRFall04.exe' in the Citipak directory and restart payroll."
    frmMessage.Label1.Top = 700
    frmMessage.Show vbModal
    Call Terminate
  End If
  
WinToWin:
  If Exist("PRConvertWin2Win.exe") Then
    ClearInUsePRReg PWcnt 'added 11/02/2002
    Shell "PRConvertWin2Win.exe", vbMaximizedFocus
    DoEvents
    Unload frmPayrollMainMenu
    End
  End If

ConvertDone:
  'if we are coming from the Check Print program then
  'we want to skip opening the Payroll main menu and
  'jump directly to Payroll Processing Menu so we do
  'this by creating a file called "fromchkprnt.dat" when
  'we exit the Payroll Check program...so if it exists
  'it could only be there if we just left Payroll Check...
  'it it exists then we know to go to Payroll Processing
  'and we know we must delete it because it is no longer
  'needed
  
  'added 1/19/04 so that if a new customer is starting
  'up then since they will not have prunit.dat saved yet
  'there is no need to look for the UnitRec.FileVer field
'  If Exist("prdata/prunit.dat") Then
'    OpenUnitFile UHandle
'    Get UHandle, 1, UnitRec
'    Close UHandle
'    If QPTrim$(UnitRec.FileVer) <> "Fall04" Then
'      MsgBox "This version of payroll requires a conversion to the Fall 2004 update."
'      Close
'      Call Terminate
'    End If
'  End If
  
  If Exist("PRConvertDos2Win.exe") Then Kill "PRConvertDos2Win.exe" '11)
      
  If Exist("PRConvertWin2Win.exe") Then Kill "PRConvertWin2Win.exe" '11)
      
  If Exist("CnvtPRFall04.exe") Then Kill "CnvtPRFall04.exe"
  
  'the next series of code is used to get the
  'identity of the current clerk using payroll
  'and recorded anytime MainLog is accessed
  cnt& = 199
  ComputerName$ = String$(200, 0)
  dl& = GetUserName(ComputerName$, cnt)
  ComputerName$ = QPTrim$(ComputerName$)
  
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  
  'this saves the current path
  StartPath = App.Path
  If Right$(StartPath, 1) = "\" Then
    StartPath = Mid$(StartPath, 1, Len(StartPath) - 1)
  End If
  
  ThisDir = StartPath + "\PRData" 'not case sensitive
  
  If Not DirExists(ThisDir) Then
    frmMessage.Label1.Caption = "The directory 'PRData' could not be located in the Citipak directory. Payroll cannot operate without the PRData folder. Loading aborted."
    frmMessage.Label1.Top = 700
    frmMessage.Show vbModal
    
    If Exist("prmain.dat") Then
      KillFile "prmain.dat"
    End If
    
    MainLog ("Payroll.exe terminated because the PRData directory could not be found.")
    DoEvents
    Call Ready4others(PWcnt)
    DoEvents
    If Exist(QPTrim$(StartPath) + "\" + "Citipak.exe") Then
      Shell QPTrim$(StartPath) + "\" + "Citipak.exe", vbMaximizedFocus
    End If
    KillFile "taxmain.dat"
    KillFile "paycheckmain.dat"
    KillFile "prmain.dat"
    Call Terminate2Shell 'closes all forms but does not clear password data
  End If
  
  ThisDir = StartPath + "\PRRPTS"
  
  If Not DirExists(ThisDir) Then
    frmMessageWOpts.Label1.Caption = "The directory 'PRRPTS' could not be located in the Citipak directory. Without the 'PRRPTS' directory graphics report printing is not possible. If you wish to create the 'PRRPTS' directory then press F10. Otherwise press ESC and call Southern Software @ 1-800-842-8190 for support."
    frmMessageWOpts.Label1.Top = 500
    frmMessageWOpts.cmdCont.Text = "F10 Make PRRPTS"
    frmMessageWOpts.cmdExit.Text = "ESC Escape"
    frmMessageWOpts.Show vbModal
    If frmMessageWOpts.fptxtChoice.Text = "continue" Then
      Unload frmMessageWOpts
      MkDir StartPath + "\PRRPTS"
    Else
      Unload frmMessageWOpts
    End If
  End If
  
  ThisDir = StartPath + "\PRRDF"
  
  If Not DirExists(ThisDir) Then
    frmMessageWOpts.Label1.Caption = "The directory 'PRRDF' could not be located in the Citipak directory. Without the 'PRRDF' directory graphics reports reprints are not possible. If you wish to create the 'PRRDF' directory then press F10. Otherwise press ESC and call Southern Software @ 1-800-842-8190 for support."
    frmMessageWOpts.Label1.Top = 500
    frmMessageWOpts.cmdCont.Text = "F10 Make PRRDF"
    frmMessageWOpts.cmdExit.Text = "ESC Escape"
    frmMessageWOpts.Show vbModal
    If frmMessageWOpts.fptxtChoice.Text = "continue" Then
      Unload frmMessageWOpts
      MkDir StartPath + "\PRRDF"
    Else
      Unload frmMessageWOpts
    End If
  End If
  
  OpenPayRateFile PHandle
  NumOfPayRate = LOF(PHandle) / Len(PayRate)
  If NumOfPayRate = 0 Then
    frmLoadingRpt.Label1.Caption = "Creating Initial Pay Rate Records..."
    frmLoadingRpt.Label1.FontSize = 10
    frmLoadingRpt.Show
    DoEvents
    Call UpdatePayRate("None", "None", 0, 0, "None", 0, False)
    Unload frmLoadingRpt
  End If
  Close PHandle
  
  If CitiPass.Administ = True Or PWcnt = -3 Then
    mnuOptions.Visible = True
  Else
    mnuOptions.Visible = False
  End If
  'only use these next two lines when working in the environment
  'comment out the rest of the time
  LevelPass = 1
  PWcnt = 6
  
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      If exitCmd.Enabled = True Then
        SendKeys "%E"
        Call exitCmd_Click
        KeyCode = 0
      End If
    Case Else:
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  KillFile "taxmain.dat"

End Sub

Private Sub mnuReindex_Click()
  Call MakeEmpIndexs
  MsgBox "Re-indexing has completed successfully."
End Sub

Private Sub mnuRelink_Click()
  frmPRRelinkTransHist.Show
  DoEvents
  Unload frmPayrollMainMenu
End Sub

Private Sub payRollProcCmmd_Click()
  Dim UHandle As Integer
  Dim UnitRec As UnitFileRecType
  Dim PRDraftFileHandle As Integer
  Dim PRDraftFileRec As DraftInfoFileName
  Dim PRLen As Integer
'  Dim SysRec As RegDSysFileRecType
'  Dim SysCnt As Integer
'  Dim SysHandle As Integer
  Dim ThisYear$
  Dim LastYear$
  Dim TYHandle As Integer
  Dim One As Integer
  Dim SumTotal As Double
  Dim Emp3Rec As EmpData3Type
  Dim EHandle As Integer
  Dim NumOfEmp3Recs As Integer
  Dim x As Integer
  Dim DateLen As Integer
  
'  OpenSysFile SysHandle
'  SysCnt = LOF(SysHandle) \ Len(SysRec)
'  Get SysHandle, 1, SysRec
'  Close SysHandle
'  'this is an alert to the user that there are
'  'problems with the current citipak path
'  If SysCnt <> 0 Then
'    CurrCitiPath = QPTrim$(SysRec.CITIDIR)
'    If QPTrim$(SysRec.CITIDIR) <> "" Then '11/27/02 for users with no GL
'      If CheckCitiDir(QPTrim$(SysRec.CITIDIR)) = 0 Then
'        frmWarnOpenWNoDir.Show vbModal, Me
'      End If
'    End If
'  End If
  'Year end initialization check
  DateLen = Len(Date)
  DateLen = DateLen - 3
  ThisYear = Mid(Date, DateLen, 4) 'assign variable with current year
  LastYear$ = CStr(CInt(ThisYear - 1)) 'assign variable with last year
  If Not Exist(ThisYear + ".dat") Then 'if this is the first time this
  'code has been activated then go ahead and create the .dat file that
  'flags the current year
    One = 1
    TYHandle = FreeFile
    Open ThisYear + ".dat" For Output As TYHandle
    Print #TYHandle, One 'current year flag created here
    Close TYHandle
  End If
  If Exist(LastYear$ + ".dat") Then 'if a variable assigned with last year
  'exists then this is not the first time this code has been activated and it
  'when it was last activated was not in the current year
    KillFile (LastYear$ + ".dat") 'don't need this file anymore...it has
    'served its purpose
    SumTotal = 0
    OpenEmpData3File EHandle 'now check to see if the year end initialization
    'has taken place by checking the file that is cleared in that process...
    'Emp3
    NumOfEmp3Recs = LOF(EHandle) / Len(Emp3Rec)
    For x = 1 To NumOfEmp3Recs
      Get EHandle, x, Emp3Rec
      SumTotal = SumTotal + Emp3Rec.YTDFedGrossPay 'if the value for this field
      'is more than 0 then it hasn't been cleared...this can only happen before the
      'first payroll for the new year is posted
    Next x
    Close EHandle
    If SumTotal > 0 Then 'OK...we know the new year was not initialized
      frmWarnInitialize.Show vbModal 'so warn the user
      If frmWarnInitialize.fptxtChoice.Text = "continue" Then
        MainLog ("User warned to initialize for the year " + LastYear + " and the user jumped to the initialization screen.")
        Unload frmWarnInitialize
        frmWarningYearEnd.Show
        DoEvents
        Unload Me
        Exit Sub
      Else
        Unload frmWarnInitialize
        MainLog ("User warned to initialize for the year " + LastYear + " but elected to exit warning without jumping to the initialization screen.")
      End If
    End If
  End If
  OpenPRDraftFile PRDraftFileHandle
  PRLen = LOF(PRDraftFileHandle) / Len(PRDraftFileRec)
  Get PRDraftFileHandle, 1, PRDraftFileRec
  Close PRDraftFileHandle
  
  OpenUnitFile UHandle
  Get UHandle, 1, UnitRec
  Close UHandle
  'if the ACH option is Y but the ACH control data is
  'empty the user cannot enter Payroll Processing because
  'the payroll process will use the "Y" as the OK
  'to proceed with pulling data from the ACH control
  'file which doesn't exist
  If QPTrim$(UnitRec.BankDraft) = "Y" Then
    If PRLen > 0 Then
      GoTo LetEmIn
    Else
      frmWarningACH.Show vbModal, Me
      Exit Sub
    End If
  End If
LetEmIn:
  If LevelPass = 0 Then
    MsgBox "Currently you are denied access to this section. Exit from payroll and log in again through the Citipak login screen."
    Exit Sub
  ElseIf LevelPass = 1 Then
    frmPayrollProcessingMenu.Show
    DoEvents
    Unload frmPayrollMainMenu
  Else
    MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
  End If

End Sub

Private Sub reportsProcCmmd_Click()
'  Dim SysRec As RegDSysFileRecType
'  Dim SysCnt As Integer
'  Dim SysHandle As Integer
'
'  OpenSysFile SysHandle
'  SysCnt = LOF(SysHandle) \ Len(SysRec)
'  Get SysHandle, 1, SysRec
'  Close SysHandle
'  'this is an alert to the user that there are
'  'problems with the current citipak path
'  If SysCnt <> 0 Then
'    CurrCitiPath = QPTrim$(SysRec.CITIDIR)
'    If QPTrim$(SysRec.CITIDIR) <> "" Then '11/27/02 for users with no GL
'      If CheckCitiDir(QPTrim$(SysRec.CITIDIR)) = 0 Then
'        frmWarnOpenWNoDir.Show vbModal, Me
'      End If
'    End If
'  End If
  If LevelPass = 0 Then
    MsgBox "Currently you are denied access to this section. Exit from payroll and log in again through the Citipak login screen."
    Exit Sub
  Else
    frmReportsProcessing.Show
    DoEvents
    Unload frmPayrollMainMenu
  End If
End Sub

Private Sub Timer1_Timer()
  KillFile "taxmain.dat"
  KillFile "paycheckmain.dat"
  KillFile "prmain.dat"
  Call Terminate2Shell 'closes all forms but does not clear password data
'  Unload Me
End Sub

Private Sub W2ProcessingCmmd_Click()
'  Dim SysRec As RegDSysFileRecType
'  Dim SysCnt As Integer
'  Dim SysHandle As Integer
  Dim ShellHandle As Integer
  Dim XHandle As Integer
  Dim CitiPassFile As Integer
  Dim CitiPass As CitiPassType
  Dim NumPassRecs As Integer
  Dim One As Integer
  Dim SSHandle As Integer
  
'  OpenSysFile SysHandle
'  SysCnt = LOF(SysHandle) \ Len(SysRec)
'  Get SysHandle, 1, SysRec
'  Close SysHandle
'  'this is an alert to the user that there are
'  'problems with the current citipak path
'  If SysCnt <> 0 Then
'    CurrCitiPath = QPTrim$(SysRec.CITIDIR)
'    If QPTrim$(SysRec.CITIDIR) <> "" Then '11/27/02 for users with no GL
'      If CheckCitiDir(QPTrim$(SysRec.CITIDIR)) = 0 Then
'        frmWarnOpenWNoDir.Show vbModal, Me
'      End If
'    End If
'  End If
  InFileNames(1) = "PRDATA\PRDEDCOD.DAT"
  InFileNames(2) = "PRDATA\PRUNIT.DAT"
  InFileNames(3) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 3) = False Then
    Close
    Exit Sub
  End If
  
  If PWcnt = -3 Then
    One = 1
    SSHandle = FreeFile
    Open "sosoftpw.dat" For Output As SSHandle
    Print #SSHandle, One
    Close SSHandle
    Shell "W2Processing.exe", vbMaximizedFocus
    Timer1.Enabled = True
'    Call Terminate2Shell 'closes all forms but does not clear password data
'    Close
    Exit Sub
  End If
    
  If LevelPass = 0 Then
    MsgBox "Currently you are denied access to this section. Exit from payroll and log in again through the Citipak login screen."
    Exit Sub
  ElseIf LevelPass = 1 Then 'access denied to LevelPass 2
    OpenCitiPassFile CitiPassFile, NumPassRecs
    Get CitiPassFile, PWcnt, CitiPass 'we are exiting so Flag2 is set to 3
    'which tells the W2.exe that a user is logged in there...if W2 is
    'terminated through QueryUnload the program looks for who is in
    'W2 and clears their data (even if a user is on another machine
    '...they will have to sign in again when they return...there is not
    'a mechanism for determining individuals in W2, just that someone
    'is logged in
    CitiPass.Flag2 = PWcnt
    Put CitiPassFile, PWcnt, CitiPass
    Close CitiPassFile
    Shell "W2Processing.exe", vbMaximizedFocus
    Timer1.Enabled = True
'    Call Terminate2Shell 'closes all forms but does not clear password data
'    Close
'    End
  Else
    MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
  End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If exitCmd.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmPayrollMainMenu.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub Timer2_Timer()
  If Exist("taxmain.dat") Then
    W2ProcessingCmmd.Enabled = True
    KillFile "taxmain.dat"
  Else
    exitCmd.Enabled = True
  End If
End Sub

Private Sub FixTransEmpPins()
  Dim EmpRec As EmpData2Type
  Dim EHandle As Integer
  Dim NumOfERecs As Long
  Dim x As Long
  Dim TransRec As TransRecType
  Dim THandle As Integer
  Dim NextRec As Long
  Dim cnt As Integer
  OpenEmpData2File EHandle
  NumOfERecs = LOF(EHandle) / Len(EmpRec)
  OpenTransHistFile THandle
  
  For x = 1 To NumOfERecs
    Get EHandle, x, EmpRec
    If x = 222 Then Stop
    NextRec = EmpRec.LastTransRec
    Do While NextRec > 0
      Get THandle, NextRec, TransRec
      If TransRec.EmpPin = 0 Then
        TransRec.EmpPin = EmpRec.EmpPin
        Put THandle, NextRec, TransRec
        cnt = cnt + 1
      End If
      NextRec = TransRec.PrevTransRec
    Loop
  Next x
  
  Close
  MsgBox "A total of " + CStr(cnt) + " transactions were updated."
    
  
  
End Sub

Private Sub FixElkton()
  Dim x As Long
  Dim TransRec As TransRecType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim OldDate As Integer
  Dim NewDate As Integer
  Dim cnt As Integer
  Dim Amt As Double
  
  OldDate = Date2Num("12/18/2006")
  NewDate = Date2Num("10/18/2006")
  OpenTransHistFile THandle
  NumOfTRecs = LOF(THandle) / Len(TransRec)
  For x = 1 To NumOfTRecs
    Get THandle, x, TransRec
'    If MakeRegDate(TransRec.CheckDate) = "12/18/2006" Then Stop
    If TransRec.CheckDate = OldDate Then
      TransRec.CheckDate = NewDate
      Put THandle, x, TransRec
      cnt = cnt + 1
      Amt = Amt + TransRec.NetPay
    End If
  Next x
  
  Close
  MsgBox ("Finished with " + CStr(cnt) + " transactions changed equaling " + Using$("$##,###.##", Amt) + ".")
  
    
End Sub

Private Sub ResetAllToPrintAdYes()
  Dim EmpRec As EmpData2Type
  Dim EHandle As Integer
  Dim NumOfERecs As Long
  Dim x As Long
  
  OpenEmpData2File EHandle
  NumOfERecs = LOF(EHandle) / Len(EmpRec)
  For x = 1 To NumOfERecs
    Get EHandle, x, EmpRec
    EmpRec.PRENOTED = "N"
    Put EHandle, x, EmpRec
  Next x
  Close EHandle
  
  MsgBox ("Completed successfully.")

End Sub
