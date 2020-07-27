VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmMainPassWord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Citipak Password Maintenance"
   ClientHeight    =   2505
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   5580
   Icon            =   "frmMainPassWord.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   5580
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Esc &Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3024
      TabIndex        =   2
      Top             =   1512
      Width           =   1404
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "F10 &Enter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1128
      TabIndex        =   1
      Top             =   1512
      Width           =   1404
   End
   Begin EditLib.fpText txtPW 
      Height          =   420
      Left            =   1692
      TabIndex        =   0
      Top             =   768
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Administrator Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   1176
      TabIndex        =   3
      Top             =   264
      Width           =   3228
   End
End
Attribute VB_Name = "frmMainPassWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim vWidth%, vHeight%, vTop%, vLeft%
Dim CitiPass As CitiPassType
Dim stopnow As Integer, theyare As String, onwhat As String
'***********************
' LevelPass CODES
' 1 = Support  ***
' 2 = Administrator  ******
'**********************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    Cancel = True
  End If
End Sub

Private Sub cmdCancel_Click()
  Unload frmMainPassWord
End Sub

Private Sub cmdEnter_Click()
  Dim Notvalid As Boolean, Z As String, Pz As String, cnt As Integer
  Dim Az As String, CntA As Integer, Findnum As Integer
  Notvalid = False
  Pz$ = ""
  Z$ = txtPW
  For cnt = 1 To Len(Z$)
    Pz$ = Pz$ + Chr$(Asc(Mid$(Z$, cnt, 1)) Xor 127)
  Next
  
  If Pz$ = "1010-8>16<" Then
      LevelPass = 1
      If Chk4file = 1 Then
        SetFile
        PWUser = "Software S."
        MainLog "In PW Entry/Edit"
        Unload frmMainPassWord
        frmUserSelect.Show
        Unload frmMainMenu
      Else
        Unload frmMainPassWord
        If stopnow = 1 Then
          Exit Sub
        End If
      End If
  Else
    Findnum = FindAdmin(Pz$)
    If stopnow = 1 Then
      Unload frmMainPassWord
      Exit Sub
    End If
    If Findnum > 0 Then
      LevelPass = 2
      If Chk4file = 1 Then
        SetFile
        Unload frmMainPassWord
        MainLog "In PW Entry/Edit"
        frmUserSelect.Show
        Unload frmMainMenu
      Else
        Unload frmMainPassWord
        Close CPAdminhand
      End If
    Else
      Notvalid = True
      If Findnum = -1 Then
        If MsgBox(theyare$ + " In Session on- " + onwhat$ + " Select YES to Exit, Or if this is in error, Select NO to Clear Setting and then retry Password. Everyone Must Exit CitiPak First.", vbYesNo, "Warning!!") = vbNo Then
          MainLog "Reset PassInUse Codes By Admin"
          ClearInUse -1
          txtPW = ""
          txtPW.SetFocus
        End If
        Notvalid = False
      End If
      LevelPass = 0
    End If
  End If
  If Notvalid = True Then
    MsgBox "Invalid Password. Try again or Call Software Support.", vbOKOnly, "Invalid Entry"
    MainLog "Invalid Password in PWMaint"
    txtPW = ""
    txtPW.SetFocus
  End If
End Sub

Private Sub Form_Initialize()
  vWidth = Screen.Width * 0.5      ' Set width of form.
  vHeight = Screen.Height * 0.33  ' Set height of form.
  vLeft = (Screen.Width - vWidth) \ 2   ' Center form horizontally.
  vTop = ((Screen.Height - vHeight) \ 2) + 10  ' Center form vertically.
End Sub
Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Me.Width = vWidth
  Me.Height = vHeight
  Me.Left = vLeft
  Me.Top = vTop
  stopnow = 0
End Sub

'

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
'    Case vbKeyDown, vbKeyReturn:
'      SendKeys "{Tab}"
'      KeyCode = 0
'    Case vbKeyUp:
'      SendKeys "+{Tab}"
'      KeyCode = 0
    Case vbKeyEscape:
      cmdCancel_Click
      KeyCode = 0
    Case vbKeyF10:
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
Public Function FindAdmin(pw$)
  Dim NumPassRecs As Integer, cnt As Integer, CitiPassFile As Integer
  Dim Oktogo As Boolean, LookFor As String
  pw$ = LTrim$(pw$)
  'If Exist("CitiPass.dat") Then
  OpenCitiPassFile CitiPassFile, NumPassRecs
  If Not CitiPassFile = -1 Then
  
  For cnt = 1 To NumPassRecs
  Get CitiPassFile, cnt, CitiPass
  If Not CitiPass.DelFlag Then
  LookFor$ = Trim$(CitiPass.PassWord)
  If CitiPass.Administ = True Then
    If pw$ = LookFor$ Then
      If Not CitiPass.InUseFlag Then
        Oktogo = True
        'CitiPass.InUseFlag = True
        PWUser = QPTrim(CitiPass.UserName)
        'Put CitiPassFile, cnt, CitiPass
        Exit For
      Else
        theyare = QPTrim(CitiPass.UserName)
        onwhat = QPTrim(CitiPass.CompName)
        FindAdmin = -1
      End If
    End If
  End If
  End If
  Next
  Close CitiPassFile
  'SetAttr ("CitiPass.dat"), vbReadOnly
  Else
    stopnow = 1
  End If
  If Oktogo Then
    PWcnt = cnt
    FindAdmin = cnt
  Else
    PWcnt = 0
  End If
End Function
Private Function Chk4file()
  Dim CitiPass As CitiPassType
  Dim NumPassRecs As Integer, cnt As Integer, CitiPassFile As Integer
  Dim somebodyin As Integer
  somebodyin = 0
  theyare = ""
  'If Exist("CitiPass.dat") Then
  OpenCitiPassFile CitiPassFile, NumPassRecs
  If Not CitiPassFile = -1 Then
    For cnt = 1 To NumPassRecs
    Get CitiPassFile, cnt, CitiPass
    If Not CitiPass.DelFlag Then
      If CitiPass.InUseFlag Then
        somebodyin = somebodyin + 1
        theyare = QPTrim(CitiPass.UserName)
        onwhat = QPTrim(CitiPass.CompName)
        Exit For
      End If
    End If
    Next
    Close CitiPassFile
    If somebodyin > 0 Then
      If MsgBox("User " + theyare$ + " Currently Signed On " + onwhat$ + ", Select YES to Allow Users to Exit Citipak or NO to Clear File Setting Due To Error. Everyone Must Exit CitiPak First...", vbYesNo, "Warning!!") = vbNo Then
        MainLog "Reset PassInUse Code"
        ClearInUse -1
        Chk4file = 1
      Else
        Chk4file = 2
      End If
    Else
      Chk4file = 1
    End If
  Else
    Chk4file = 2
    stopnow = 1
  End If
End Function
Private Sub SetFile()
'Lock the file so nobody can use during maintenance
  Dim PassRecLen As Integer, NumPassRecs As Integer

 ' On Local Error GoTo PassError
  'If Exist("CitiPass.dat") Then
    PassRecLen = Len(CitiPass)
    CPAdminhand = FreeFile
    Open "CitiPass.dat" For Random Lock Read Write As CPAdminhand Len = PassRecLen
    NumPassRecs = LOF(CPAdminhand) \ PassRecLen
  'End If
  Exit Sub
PassError:
  CPAdminhand = -1
  MsgBox "Password Maintenance Already Open.", vbOKOnly, "Access Denied"
  MainLog "PW File already locked"
End Sub

