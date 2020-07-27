VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmW2Processing 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Processing Menu vs 2.05"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmW2Processing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   360
      Top             =   840
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   360
      Top             =   360
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMedOnly 
      Height          =   435
      Left            =   4005
      TabIndex        =   2
      Top             =   2940
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmW2Processing.frx":08CA
   End
   Begin VB.Timer Timer1 
      Interval        =   355
      Left            =   5664
      Top             =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmd941Only 
      Height          =   435
      Left            =   4005
      TabIndex        =   1
      Top             =   2430
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmW2Processing.frx":0AB9
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEditW2 
      Height          =   435
      Left            =   4005
      TabIndex        =   3
      Top             =   3465
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmW2Processing.frx":0CA4
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintReport 
      Height          =   435
      Left            =   4005
      TabIndex        =   4
      Top             =   3975
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmW2Processing.frx":0E92
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintForms 
      Height          =   435
      Left            =   4005
      TabIndex        =   5
      Top             =   4485
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmW2Processing.frx":1076
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdReprint 
      Height          =   435
      Left            =   4005
      TabIndex        =   6
      Top             =   5010
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmW2Processing.frx":1259
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdW2ESub 
      Height          =   435
      Left            =   4005
      TabIndex        =   7
      Top             =   5520
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmW2Processing.frx":143E
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   435
      Left            =   4005
      TabIndex        =   9
      Top             =   7065
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmW2Processing.frx":162B
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdW3 
      Height          =   435
      Left            =   4005
      TabIndex        =   8
      Top             =   6030
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmW2Processing.frx":1815
   End
   Begin fpBtnAtlLibCtl.fpBtn cmd941 
      Height          =   435
      Left            =   4005
      TabIndex        =   10
      Top             =   6555
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmW2Processing.frx":19F7
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   2101
      Top             =   2103
      Width           =   971
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   8593
      Top             =   2103
      Width           =   971
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2220.428
      X2              =   2220.428
      Y1              =   2150.985
      Y2              =   7889.543
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1095
      Index           =   1
      Left            =   1500
      Top             =   900
      Width           =   8655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Payroll Tax Forms Menu"
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
      Top             =   1250
      Width           =   6012
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1214
      Left            =   1500
      Top             =   770
      Width           =   8652
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8710.757
      X2              =   8710.757
      Y1              =   2153.909
      Y2              =   7889.543
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2205.432
      X2              =   2919.248
      Y1              =   7894.417
      Y2              =   7894.417
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8710.757
      X2              =   9412.576
      Y1              =   7894.417
      Y2              =   7894.417
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      DrawMode        =   12  'Nop
      X1              =   3756.033
      X2              =   4475.848
      Y1              =   8710.173
      Y2              =   8710.173
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5900
      Index           =   0
      Left            =   2220
      Top             =   2201
      Width           =   732
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5900
      Index           =   1
      Left            =   8712
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
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2100
      Top             =   1971
      Width           =   972
   End
End
Attribute VB_Name = "frmW2Processing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmd941_Click()
  frmLoadingW2Rpt.fpBtn1.Text = "Loading Screen"
  frmLoadingW2Rpt.Show
  DoEvents
  frm941.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmd941Only_Click()
  InFileNames(1) = "PRDATA\PRUNIT.DAT" '7/20
  InFileNames(2) = "PRDATA\PREMP2.DAT" '7/20
  InFileNames(3) = "PRDATA\PRDEDCOD.DAT" '7/20
  InFileNames(4) = "PRDATA\PRTRANSH.DAT" '7/20
'  InFileNames(5) = "PRDATA\PRW2SETU.DAT" '7/20
  
  If FilesROK(Me, InFileNames(), OutFileNames(), 4) = False Then '7/20
    Close '7/20
    Exit Sub '7/20
  End If '7/20
  
  frm914W2Setup.Show
  DoEvents
  Unload frmW2Processing

End Sub

Private Sub cmdPrintForms_Click()
  InFileNames(1) = "PRDATA\PRUNIT.DAT" '7/20
  InFileNames(2) = "PRDATA\PREMP2.DAT" '7/20
  InFileNames(3) = "PRDATA\PREMPL.IDX" '7/20
  InFileNames(4) = "PRDATA\PRW2INFO.DAT" '7/20
  If FilesROK(Me, InFileNames(), OutFileNames(), 4) = False Then '7/20
    Close '7/20
    Exit Sub '7/20
  End If '7/20
  
  frmW2FormsPrinting.Show
  DoEvents
  Unload frmW2Processing
End Sub

Private Sub cmdPrintReport_Click()
  InFileNames(1) = "PRDATA\PRUNIT.DAT" '7/20
  InFileNames(2) = "PRDATA\PREMPL.IDX" '7/20
  InFileNames(3) = "PRDATA\PRW2INFO.DAT" '7/20
  InFileNames(4) = "PRDATA\PREMP2.DAT" '7/20
  If FilesROK(Me, InFileNames(), OutFileNames(), 4) = False Then '7/20
    Close '7/20
    Exit Sub '7/20
  End If '7/20
  frmReportOpt.Show vbModal
  If RptOpt = 2 Then
    Call W2ReportT
    Exit Sub
  ElseIf RptOpt = 1 Then
    Call W2ReportG
  Else
    Exit Sub
  End If
'  Call W2Report
End Sub

Private Sub cmdReprint_Click()
  InFileNames(1) = "PRDATA\PRUNIT.DAT" '7/20
  InFileNames(2) = "PRDATA\PREMP2.DAT" '7/20
  InFileNames(3) = "PRDATA\PREMPL.IDX" '7/20
  InFileNames(4) = "PRDATA\PRW2INFO.DAT" '7/20
  If FilesROK(Me, InFileNames(), OutFileNames(), 4) = False Then '7/20
    Close '7/20
    Exit Sub '7/20
  End If '7/20
  
  If Not Exist("PRDATA\W2RPNIDX.DAT") Then
    MsgBox "No reprints possible until tractor fed or dot matrix W2 forms have been printed."
    Exit Sub
  End If
  
  frmW2FormsReprinting.Show
  DoEvents
  Unload frmW2Processing
End Sub

Private Sub cmdW2ESub_Click()
  frmW2ElecSub.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdW3_Click()
  frmW3.Show
  DoEvents
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      If cmdExit.Enabled = True Then
        SendKeys "%x"
        Call cmdExit_Click
        KeyCode = 0
      End If
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Dim cnt&, dl& '7/20
  Dim ShellHandle As Integer
  Dim Cnt2$
  Dim CitiPassFile As Integer
  Dim CitiPass As CitiPassType
  Dim NumPassRecs As Integer
  Dim x As Integer
  
  App.HelpFile = "helpfiles\PAYROLL.hlp"

  Me.HelpContextID = hlpPayrollTaxForms
  If Exist("taxmain.dat") Then
    KillFile "taxmain.dat"
  End If
  
  If Exist("prmain.dat") Then
    KillFile "prmain.dat"
    cmdExit.Enabled = False 'Timer2 re-enables it
  End If
  RegExit = False
  If App.PrevInstance Then
    ActivatePrevInstance
  End If
  
  cnt& = 199
  ComputerName$ = String$(200, 0) '7/20
  dl& = GetUserName(ComputerName$, cnt) '7/20
  ComputerName$ = QPTrim$(ComputerName$) '7/20
  
  StartPath = App.Path
  If Right$(StartPath, 1) = "\" Then
    StartPath = Mid$(StartPath, 1, Len(StartPath) - 1)
  End If
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me

  If Exist("sosoftpw.dat") Then
    KillFile "sosoftpw.dat"
    PWcnt = -3
    Exit Sub
  End If

  If PWcnt > 0 Or PWcnt = -3 Then Exit Sub 'opening from another form
  OpenCitiPassFile CitiPassFile, NumPassRecs
  For x = 1 To NumPassRecs 'this procedure prevents someone
  'from trying to access W2.exe through the W2.exe
  'instead of through payroll.exe
    Get CitiPassFile, x, CitiPass
    If CitiPass.Flag2 > 0 Then
      PWcnt = CitiPass.Flag2 'had to have come from payroll.exe if > 0
      CitiPass.Flag2 = 0 'clear to zero in case another valid user
      'has to come thru this same procedure
      Put CitiPassFile, x, CitiPass
      Close CitiPassFile
      Exit Sub
    End If
  Next x
  Close CitiPassFile
  
'  Call cmdExit_Click

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub cmdExit_Click()
  Dim XHandle As Integer
  Dim PauseTime, Start, Finish, TotalTime
  Dim cnt$
  Dim CitiPassFile As Integer
  Dim CitiPass As CitiPassType
  Dim NumPassRecs As Integer
  Dim x As Integer
'  Dim One As Integer
'  Dim AHandle As Integer
'
'  One = 1
'  AHandle = FreeFile
'  Open "taxmain.dat" For Output As AHandle
'  Print #AHandle, One
'  Close AHandle

'  PauseTime = 1#   ' Set duration.
'  Start = Timer   ' Set start time.
'  Do While Timer < Start + PauseTime
'     DoEvents   ' Yield to other processes.
'  Loop
  If PWcnt = 0 Or PWcnt = -3 Then GoTo ExitW2
  
  OpenCitiPassFile CitiPassFile, NumPassRecs 'reassign all globals
  Get CitiPassFile, PWcnt, CitiPass
  'needs to be changed to -1 so that payroll.exe knows that this
  '.exe has just been accessed
  CitiPass.Flag2 = -2
  Put CitiPassFile, PWcnt, CitiPass
  Close CitiPassFile
  
ExitW2:
  Shell "payroll.exe", vbMaximizedFocus
  Timer3.Enabled = True
'  Close
'  RegExit = True
'  Call UnloadAllFormsAndOpn(RegExit) 'If RegExit is false then
'  'UnloadAllFormsAndOpn clears the password file and leaves it
'  'alone if it is true...we want to preserve the password data
'  'for use in the loading procedure in payroll.exe if this is
'  'a routine exit
'  End
End Sub

Private Sub cmdMedOnly_Click()
  InFileNames(1) = "PRDATA\PRUNIT.DAT" '7/20
  InFileNames(2) = "PRDATA\PREMP2.DAT" '7/20
  InFileNames(3) = "PRDATA\PRDEDCOD.DAT" '7/20
  InFileNames(4) = "PRDATA\PRTRANSH.DAT" '7/20
'  InFileNames(5) = "PRDATA\PRW2SETU.DAT" '7/20
  
  If FilesROK(Me, InFileNames(), OutFileNames(), 4) = False Then '7/20
    Close '7/20
    Exit Sub '7/20
  End If '7/20

  frmMedW2Setup.Show
  DoEvents
  Unload frmW2Processing
End Sub

Private Sub cmdEditW2_Click()
  
  InFileNames(1) = "PRDATA\PREMP1.DAT" '7/20
  InFileNames(2) = "PRDATA\PREMP2.DAT" '7/20
  InFileNames(3) = "PRDATA\PREMPL.IDX" '7/20
  If FilesROK(Me, InFileNames, OutFileNames, 3) = False Then '7/20
    Close '7/20
    Exit Sub '7/20
  End If '7/20
  
  frmEditReviewEmpW2.Show
  DoEvents
  Unload frmW2Processing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      Call UnloadAllFormsAndOpn(RegExit)
      MainLog ("Payroll.exe terminated via menu bar on frmW2Processing.")
      End
    End If
  End If
End Sub

Private Sub Timer2_Timer()
  cmdExit.Enabled = True
End Sub

Private Sub Timer3_Timer()
'  Unload Me
  Close
  RegExit = True
  Call UnloadAllFormsAndOpn(RegExit) 'If RegExit is false then
  'UnloadAllFormsAndOpn clears the password file and leaves it
  'alone if it is true...we want to preserve the password data
  'for use in the loading procedure in payroll.exe if this is
  'a routine exit
  End
End Sub
