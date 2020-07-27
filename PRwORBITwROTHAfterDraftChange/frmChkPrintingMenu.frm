VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmChkPrintingMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check Printing Menu vs 2.05"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   FillColor       =   &H8000000B&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChkPrintingMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Interval        =   2000
      Left            =   360
      Top             =   720
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   360
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Interval        =   355
      Left            =   5520
      Top             =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintPRChks 
      Height          =   495
      Left            =   4004
      TabIndex        =   1
      Top             =   3312
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
      ButtonDesigner  =   "frmChkPrintingMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdReprint 
      Height          =   495
      Left            =   4004
      TabIndex        =   2
      Top             =   4152
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
      ButtonDesigner  =   "frmChkPrintingMenu.frx":0AB2
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintChkReg 
      Height          =   495
      Left            =   4004
      TabIndex        =   3
      Top             =   4968
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
      ButtonDesigner  =   "frmChkPrintingMenu.frx":0C9D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
      Height          =   495
      Left            =   4004
      TabIndex        =   4
      Top             =   5760
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
      ButtonDesigner  =   "frmChkPrintingMenu.frx":0E85
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   123
      Left            =   8591
      Top             =   2050
      Width           =   971
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   123
      Left            =   2100
      Top             =   2050
      Width           =   971
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Height          =   1065
      Left            =   1500
      Top             =   880
      Width           =   8655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Check Printing Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   363
      Left            =   2819
      TabIndex        =   0
      Top             =   1218
      Width           =   6010
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   2205
      X2              =   2919
      Y1              =   7894
      Y2              =   7894
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   2220
      X2              =   2220
      Y1              =   2151
      Y2              =   7892
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   8711
      X2              =   8711
      Y1              =   2154
      Y2              =   7894
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   8711
      X2              =   9413
      Y1              =   7894
      Y2              =   7894
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5750
      Index           =   0
      Left            =   2219
      Top             =   2145
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   246
      Index           =   0
      Left            =   2100
      Top             =   1921
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5750
      Index           =   1
      Left            =   8710
      Top             =   2145
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   246
      Index           =   2
      Left            =   8590
      Top             =   1921
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1183
      Left            =   1500
      Top             =   750
      Width           =   8650
   End
End
Attribute VB_Name = "frmChkPrintingMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdEscape_Click()
  Dim ChkPrint As String
  Dim XHandle As Integer
  Dim cnt$
  Dim PauseTime, Start, Finish, TotalTime
  Dim CitiPassFile As Integer
  Dim CitiPass As CitiPassType
  Dim NumPassRecs As Integer
  Dim x As Integer
  Dim One As Integer
  Dim AHandle As Integer
  
  One = 1
  AHandle = FreeFile
  Open "paycheckmain.dat" For Output As AHandle
  Print #AHandle, One
  Close AHandle
'  PauseTime = 1#   ' Set duration.
'  Start = Timer   ' Set start time.
'  Do While Timer < Start + PauseTime
'    DoEvents   ' Yield to other processes.
'  Loop
  If PWcnt = 0 Or PWcnt = -3 Then GoTo ExitCheck
  OpenCitiPassFile CitiPassFile, NumPassRecs 'reassign all globals
  Get CitiPassFile, PWcnt, CitiPass
  'needs to be changed to -1 so that payroll.exe knows that this
  '.exe has just been accessed
  CitiPass.Flag2 = -1
  Put CitiPassFile, PWcnt, CitiPass
  Close CitiPassFile
  
ExitCheck:

  Shell StartPath + "\payroll.exe", vbMaximizedFocus
  Timer2.Enabled = True
'  DoEvents
'  RegExit = True 'if true than UnloadAllFormsAndOpn does not clear the
'  'password file...we need this preserved so we can use it in payroll.exe
'  Call UnloadAllFormsAndOpn(RegExit)
'  MainLog ("Check printing menu terminated.")
'  End
End Sub

Private Sub cmdPrintChkReg_Click()
  InFileNames(1) = "PRDATA\PRUNIT.DAT" '7/20
  InFileNames(2) = "PRDATA\PREMPN.IDX" '7/20
  InFileNames(3) = "PRDATA\PRCHECKS.DAT" '7/20
  
  If FilesROK(Me, InFileNames(), OutFileNames(), 3) = False Then '7/20
    Close '7/20
    Exit Sub '7/20
  End If '7/20
  frmReportOpt.Show vbModal
  If RptOpt = 2 Then
    Call CreateCheckRegisterT
    Exit Sub
  ElseIf RptOpt = 1 Then
    Call CreateCheckRegisterG
  Else
    Exit Sub
  End If
  
End Sub

Private Sub cmdPrintPRChks_Click()
  Dim RPTPitch As Integer
  Dim PrntType As PRNSetupRecType
  Dim x As Integer
  Dim PHandle As Integer
  Dim SHandle As Integer
  Dim SysRec As RegDSysFileRecType
  
  OpenPrinterSetupFile PHandle
  Get PHandle, 1, PrntType
  Close PHandle
  
  If PrntType.RPT(15) = 0 Then
    OpenSysFile SHandle
    Get SHandle, 1, SysRec
    Close SHandle
    If SysRec.CheckStyle = 1 Or SysRec.CheckStyle = 2 Or SysRec.CheckStyle = 3 Or SysRec.CheckStyle = 4 Then
      MsgBox "ERROR: The Payroll Checks printer pitch saved on the Printer Control screen is set to zero. This will cause the program to crash. Please save a valid Payroll Check pitch or change your check style to a laser check."
      Close
      Exit Sub
    End If
  End If
  
  frmChkPrintInfo.Show
  DoEvents
  Unload frmChkPrintingMenu
  
End Sub

Private Sub cmdReprint_Click()
  frmChkReprintInfo.Show
  DoEvents
  Unload frmChkPrintingMenu
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
      If cmdEscape.Enabled = True Then
        SendKeys "%X"
        Call cmdEscape_Click
        KeyCode = 0
      End If
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Dim cnt&, dl&
  Dim FileHandle As Integer
  Dim ShellHandle As Integer
  Dim Cnt2$
  Dim CitiPassFile As Integer
  Dim CitiPass As CitiPassType
  Dim NumPassRecs As Integer
  Dim x As Integer
  
  App.HelpFile = "helpfiles\PAYROLL.hlp"

  Me.HelpContextID = hlpCheckPrinting
  If Exist("paycheckmain.dat") Then
    KillFile "paycheck.dat"
  End If
  If Exist("prmain.dat") Then
    KillFile "prmain.dat"
    cmdEscape.Enabled = False 'timer3 re-enables it
  End If
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  
  RegExit = False
  cnt& = 199
  ComputerName$ = String$(200, 0)
  dl& = GetUserName(ComputerName$, cnt)
  ComputerName$ = QPTrim$(ComputerName$)
  StartPath = App.Path
  If Right$(StartPath, 1) = "\" Then
    StartPath = Mid$(StartPath, 1, Len(StartPath) - 1)
  End If
  If App.PrevInstance Then
    ActivatePrevInstance
  End If
  
  If Exist("sosoftpw.dat") Then
    KillFile "sosoftpw.dat"
    PWcnt = -3
    Exit Sub
  End If
  
  If PWcnt > 0 Or PWcnt = -3 Then Exit Sub 'opening from another form requires this
  OpenCitiPassFile CitiPassFile, NumPassRecs
  For x = 1 To NumPassRecs 'this procedure prevents someone
  'from trying to access payrollcheck.exe through the payrollcheck.exe
  'instead of through payroll.exe
    Get CitiPassFile, x, CitiPass
    If CitiPass.Flag2 > 0 Then
      PWcnt = CitiPass.Flag2 'had to have come from payroll.exe if > 0
      CitiPass.Flag2 = 0 'clear to zero in case another valid user
      'has to come thru this same procedure
      Put CitiPassFile, x, CitiPass
      Close CitiPassFile
      Exit Sub 'we're cleared now so exit past "Call cmdEscape_Click"
    End If
  Next x
  Close CitiPassFile
  
  Call cmdEscape_Click 'Flag2 can only be > 0 if it came from
  'payroll.exe so if it = 0 it sends the user back
  'to payroll.exe to go through the password procedure
  

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
'    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdEscape.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      Call UnloadAllFormsAndOpn(RegExit)
      MainLog ("Payroll.exe terminated via menu bar on frmChkPrintingMenu.")
      End
    End If
  End If
End Sub


Private Sub Timer2_Timer()
  DoEvents
  RegExit = True 'if true than UnloadAllFormsAndOpn does not clear the
  'password file...we need this preserved so we can use it in payroll.exe
  Call UnloadAllFormsAndOpn(RegExit)
  MainLog ("Check printing menu terminated.")
  End

End Sub

Private Sub Timer3_Timer()
  cmdEscape.Enabled = True
End Sub
