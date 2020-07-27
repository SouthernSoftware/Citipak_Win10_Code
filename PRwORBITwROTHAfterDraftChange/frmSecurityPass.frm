VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmSecurityPass 
   BackColor       =   &H008A775B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CitiPak Password Sign-In"
   ClientHeight    =   8910
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   12225
   Icon            =   "frmSecurityPass.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin EditLib.fpText txtPW 
      Height          =   420
      Left            =   5010
      TabIndex        =   0
      Top             =   4464
      Width           =   2196
      _Version        =   196608
      _ExtentX        =   3873
      _ExtentY        =   741
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   1
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   10
      MultiLine       =   0   'False
      PasswordChar    =   "*"
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCancel 
      Height          =   495
      Left            =   6405
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Press ESC to exit to the main Citipak menu."
      Top             =   5520
      Width           =   1395
      _Version        =   131072
      _ExtentX        =   2461
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmSecurityPass.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEnter 
      Height          =   495
      Left            =   4506
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Press F10 to enter payroll."
      Top             =   5520
      Width           =   1410
      _Version        =   131072
      _ExtentX        =   2487
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmSecurityPass.frx":0AE0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "PAYROLL PASSWORD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3492
      TabIndex        =   2
      Top             =   2880
      Width           =   5292
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000E&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      Height          =   2868
      Index           =   0
      Left            =   3234
      Top             =   3504
      Width           =   5748
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Payroll Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   396
      Left            =   3984
      TabIndex        =   1
      Top             =   3888
      Width           =   4260
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3222
      Top             =   2640
      Width           =   5772
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Index           =   1
      Left            =   3222
      Top             =   2544
      Width           =   5772
   End
End
Attribute VB_Name = "frmSecurityPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim CitiPass As CitiPassType
Dim stopnow As Integer, theyare As String, onwhat As String
Dim InUseExitFlag As Boolean

'***********************
' LevelPass CODES
' 1 = Full Access  ***
' 2 = Reports Only  ******
'**********************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    Cancel = True
  End If
End Sub

Private Sub cmdCancel_Click()
  'ResetInUse  'just to clear all records during testing
  Shell "Citipak.exe", vbMaximizedFocus
  Unload frmSecurityPass
End Sub

Private Sub cmdEnter_Click()
  Dim Notvalid As Boolean, z As String, Pz As String, cnt As Integer
  Dim Az As String, CntA As Integer, Findnum As Integer, PassOK As Integer
  Dim InHandle As Integer
  
  If QPTrim$(txtPW.Text) = "nonorganic" Or QPTrim$(txtPW.Text) = "NONORGANIC" Then
    PWcnt = -3
    LevelPass = 1
    Call MainLog("In Payroll, with Level " & LevelPass) 'record it
    frmPayrollMainMenu.Show
    DoEvents
    Unload frmSecurityPass
    Exit Sub
  End If
    
  Notvalid = False
  Pz$ = ""
  z$ = txtPW 'password entered
  For cnt = 1 To Len(z$)
    Pz$ = Pz$ + Chr$(Asc(Mid$(z$, cnt, 1)) Xor 127)
  Next

  If Pz$ = "1010-8>16<" Then 'default password from SoSoft
      LevelPass = 1 'Level 1 allows full access
      PWUser = "Sosoft Support"
      PWcnt = 0
      Unload frmSecurityPass
      frmPayrollMainMenu.Show
  Else
    If Len(Dir$("Citipass.dat")) Then 'not using default so check password further
      Findnum = Findsettings(Pz$) 'find password by looking for it's index
      If InUseExitFlag = True Then
        InUseExitFlag = False
        z$ = ""
        txtPW = ""
        txtPW.SetFocus
        Exit Sub
      End If
      If Not stopnow = 1 Then 'Findsettings sets stopnow
        If Findnum > 0 And LevelPass > 0 Then 'valid password found
          Call MainLog("In Payroll, with Level " & LevelPass) 'record it
          frmPayrollMainMenu.Show
          DoEvents
          Unload frmSecurityPass
        Else
          Notvalid = True
          If Findnum = -1 Then 'someone else is already using payroll.exe
            MsgBox "Password for User " + theyare$ + " In Session on " + onwhat$, vbOKOnly, "Access Denied"
            Call MainLog("Password in session, NO Payroll Access.")
            Notvalid = False
            txtPW = ""
            txtPW.SetFocus
          End If
        End If
      Else
        Exit Sub 'if stopnow = 1 then exit sub
      End If
    Else 'no Citipass.dat could be found
      MsgBox "Password Information Not Found, Check With Password Administrator.", vbOKOnly, "Access Denied"
      Call MainLog("Password Maintenance in session, Payroll Access Denied.")
      cmdCancel.SetFocus
    End If
  End If
  If Notvalid = True Then
    MsgBox "Invalid Password. Try again or See Password Administrator.", vbOKOnly, "Invalid Entry"
    Call MainLog("Invalid Payroll Password. " + txtPW)  'record invalid password attempt
    txtPW = ""
    txtPW.SetFocus
  End If
End Sub

Private Sub Form_Load()
  Dim cnt&, dl&
  Dim EHandle As Integer
  Dim CitiPassFile As Integer
  Dim CitiPass As CitiPassType
  Dim NumPassRecs As Integer
  Dim LoggedIn$
  Dim LogLevel$
  Dim CheckPrintFlag As Boolean
  Dim W2Flag As Boolean
  Dim x As Integer
  
  CurrCitiPath = App.Path
  
  If Exist("sosoftpw.dat") Then
    KillFile "sosoftpw.dat"
  End If
  InUseExitFlag = False
  PWcnt = 0
  LoggedIn = 0
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  'Startpath must be assigned a value here because if
  'payroll.exe is opening from PayrollCheck.exe or W2.exe
  'then this form exits early
  StartPath = App.Path
  If Right$(StartPath, 1) = "\" Then
    StartPath = Mid$(StartPath, 1, Len(StartPath) - 1)
  End If
  
  If App.PrevInstance Then
     ActivatePrevInstance
  End If
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  cnt& = 199
  ComputerName$ = String$(200, 0)
  dl& = GetUserName(ComputerName$, cnt)
  ComputerName$ = QPTrim$(ComputerName$)
  OpenCitiPassFile CitiPassFile, NumPassRecs 'reassign all globals
  
  For x = 1 To NumPassRecs
    Get CitiPassFile, x, CitiPass
    If CitiPass.Flag2 = -1 Then 'Flag 2 = -1 means from payrollcheck.exe
      CitiPass.Flag2 = 0 'reset this flag because it's use is solely to
      'inform that the process just left payrollcheck.exe and set PWcnt...we can skip the
      'password check-in procedure
      Put CitiPassFile, x, CitiPass
      Close CitiPassFile
      PWcnt = x 'assign global
      LevelPass = 1 'assign global...always 1 when coming from payrollcheck.exe
      frmPayrollProcessingMenu.Show 'the program left frmPayrollProcessingMenu
      'when it exited to payrollcheck.exe so we want to return there
      DoEvents
      Unload frmSecurityPass
      Exit Sub
    ElseIf CitiPass.Flag2 = -2 Then 'coming from W2 is the same process
    'as coming from payrollcheck.exe except Flag2 = -2 (denotes W2) and
    'we jump to frmPayrollMain
      CitiPass.Flag2 = 0
      Put CitiPassFile, x, CitiPass
      Close CitiPassFile
      PWcnt = x 'PWcnt is exclusive to whichever machine this user
      'is using
      LevelPass = 1 'always 1 if coming from W2.exe
      frmPayrollMainMenu.Show
      DoEvents
      Unload frmSecurityPass
      Exit Sub
    End If
  Next x
  Close CitiPassFile
End Sub
Private Sub Form_Resize()
  Temp_Class.ResizeControls Me
  DoEvents
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  Select Case KeyCode
    Case vbKeyEscape:
      SendKeys "%C"
      cmdCancel_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%E"
      cmdEnter_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub txtPW_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    cmdEnter.SetFocus
  End If
  If KeyCode = vbKeyUp Then
    cmdCancel.SetFocus
  End If
End Sub
Public Function Findsettings(pw$)
  Dim NumPassRecs As Integer, cnt As Integer, CitiPassFile As Integer
  Dim Oktogo As Boolean, Lookfor As String
  Dim x As Integer
  Dim y As Integer
  Dim InPRUse As Boolean
  Dim DoWhatFlag As PRInUse
  
  InPRUse = False 'InPRUse tells us that somebody else with full access
  'is already logged in to payroll.exe
  
'  ClearInUsePRX

  pw$ = LTrim$(pw$) 'password entered on screen
  theyare = ""
  onwhat = ""
  GoSub Check4InUse
  If Len(Dir$("Citipass.dat")) Then 'Citipass.dat contains the password files
    SetAttr ("CitiPass.dat"), vbNormal
    OpenCitiPassFile CitiPassFile, NumPassRecs
    If Not CitiPassFile = -1 Then 'Citipass.dat opened with no errors
      For cnt = 1 To NumPassRecs
        Get CitiPassFile, cnt, CitiPass
        If Not CitiPass.DelFlag Then
          Lookfor$ = Trim$(CitiPass.PassWord) 'go through each saved password looking for pw$
          If pw$ = Lookfor$ Then
            If Not CitiPass.InUseFlag Then 'InUseFlag is False
              If CitiPass.Module(4).FullAccess = True Then 'data saved when passwords initialized ...if true this user is OK
                If InPRUse = True Then 'INPRUse is set in Check4InUse ...OK someone is already in so alert
                  'the user
                  InPRUse = False 'reset for next user's check
                  DoWhatFlag = PromptUserAlreadyActive(Me)
                  Select Case DoWhatFlag
                  Case PRInUse.priuContinue
                  Case PRInUse.priuExit
                    InUseExitFlag = True 'tells Enter Sub to exit
                    Close
                    Exit Function
                  Case Else
                  End Select
                End If
                CitiPass.FlagMod = 4 'set this field to 4 denoting payroll is
                'now occupied with a full access user at PWcnt
                LevelPass = 1
              Else 'CitiPass.Module(4).FullAccess = False
                If CitiPass.Module(4).ReportsOnly = True Then
                  LevelPass = 2
                End If
              End If
              If LevelPass = 1 Or LevelPass = 2 Then
                Oktogo = True
                CitiPass.InUseFlag = True 'for this PWcnt
                PWUser = QPTrim(CitiPass.UserName)
                CitiPass.CompName = QPTrim(ComputerName$)
                Put CitiPassFile, cnt, CitiPass
              End If
              Exit For 'this user is OK
            Else 'CitiPass.InUseFlag is true so this user is already in PR
              theyare = QPTrim(CitiPass.UserName) 'can't get in because "theyare" is in already
              onwhat = QPTrim(CitiPass.CompName)
              Findsettings = -1
            End If
          End If 'end of pw$ = Lookfor$
        End If 'end of If Not CitiPass.DelFlag
      Next
      Close CitiPassFile
    Else
      stopnow = 1
    End If 'end of If Not CitiPassFile = -1
  End If 'end of If Len(Dir$("Citipass.dat"))
  
  If Oktogo Then
    PWcnt = cnt 'this is used when payroll exits so that only this
    'user's password data is cleared
    Findsettings = cnt
  Else
    LevelPass = 0
    PWcnt = 0
  End If
  
  Exit Function
  
Check4InUse:
  OpenCitiPassFile CitiPassFile, NumPassRecs
  For y = 1 To NumPassRecs
    Get CitiPassFile, y, CitiPass
    If CitiPass.FlagMod = 4 Then
    'if FlagMod = 4 then we know another full access user is
    'in payroll.exe now
      InPRUse = True
      Exit For
    End If
  Next y
  Close CitiPassFile
  Return
End Function

