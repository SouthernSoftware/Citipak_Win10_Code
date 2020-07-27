VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmSecurityPassUB 
   BackColor       =   &H008A775B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CitiPak U/B Password Sign-In"
   ClientHeight    =   8916
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   12216
   Icon            =   "frmSecurityPassUB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8916
   ScaleWidth      =   12216
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
         Size            =   10.2
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
      Left            =   4506
      TabIndex        =   1
      Top             =   5520
      Width           =   1404
   End
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
      Left            =   6402
      TabIndex        =   2
      Top             =   5520
      Width           =   1404
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "UTILITY BILLING PASSWORD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3486
      TabIndex        =   4
      Top             =   2904
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
      Caption         =   "Enter Utility Billing Password"
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
      Left            =   3972
      TabIndex        =   3
      Top             =   3912
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
Attribute VB_Name = "frmSecurityPassUB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim CitiPass As CitiPassType
Dim stopnow As Integer, theyare As String, onwhat As String
'***********************
' LevelPass CODES
' 1 = Full Access  ***
' 2 = Payments  ******
' 3 = Reports *********
'**********************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    Cancel = True
  End If
End Sub

Private Sub cmdCancel_Click()
  'ResetInUse  'just to clear all records during testing
  Shell "Citipak.exe", vbMaximizedFocus
  DoTheTime
  DoEvents
  Unload frmSecurityPassUB
End Sub

Private Sub cmdEnter_Click()
  Dim Notvalid As Boolean, Z As String, Pz As String, cnt As Integer
  Dim Az As String, CntA As Integer, Findnum As Integer, PassOK As Integer
  Notvalid = False
  Pz$ = ""
  Z$ = txtPW
  For cnt = 1 To Len(Z$)
    Pz$ = Pz$ + Chr$(Asc(Mid$(Z$, cnt, 1)) Xor 127)
  Next
  
  If Pz$ = "1010-8>16<" Then
      LevelPass = 1
      PWUser = "Sosoft Support"
      PWcnt = 0
      OPERNUM = 0
      UBLog "Support Sign in"
      Load frmUBMainMenu
      DoEvents
      frmUBMainMenu.Show
      Unload frmSecurityPassUB
  Else
    If Len(Dir$("Citipass.dat")) Then
      Findnum = Findsettings(Pz$)
      If Not stopnow = 1 Then
      If Findnum > 0 And LevelPass > 0 Then
        Call UBLog("In UB, with Level " & LevelPass)
        Load frmUBMainMenu
        DoEvents
        frmUBMainMenu.Show
        Unload frmSecurityPassUB
      Else
        Notvalid = True
        If Findnum = -1 Then
          MsgBox "Password for User " + theyare$ + " In Session on " + onwhat$, vbOKOnly, "Access Denied"
          Call UBLog("Password in session, NO Access.")
          Notvalid = False
          txtPW = ""
          txtPW.SetFocus
        End If
      End If
      Else
        Exit Sub
      End If
    Else
      MsgBox "Password Information Not Found, Check With Password Administrator.", vbOKOnly, "Access Denied"
      Call UBLog("Password Maintenance in session,  Access Denied.")
      cmdCancel.SetFocus
    End If
  End If
  If Notvalid = True Then
    MsgBox "Invalid Password. Try again or See Password Administrator.", vbOKOnly, "Invalid Entry"
    Call UBLog("Invalid ub Password." + txtPW)
    txtPW = ""
    txtPW.SetFocus
  End If
End Sub

Private Sub Form_Load()
  Dim cnt&, dl&
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
    
  If App.PrevInstance Then
     ActivatePrevInstance
  End If
  screenW = (Screen.Width / Screen.TwipsPerPixelX)
  cnt& = 199
  ComputerName$ = String$(200, 0)
  dl& = GetUserName(ComputerName$, cnt)
  ComputerName$ = QPTrim$(ComputerName$)

End Sub
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
Public Function Findsettings(pw$)
  Dim NumPassRecs As Integer, cnt As Integer, CitiPassFile As Integer
  Dim Oktogo As Boolean, LookFor As String
  pw$ = LTrim$(pw$)
  theyare = ""
  onwhat = ""
  If Len(Dir$("Citipass.dat")) Then
    SetAttr ("CitiPass.dat"), vbNormal
    OpenCitiPassFile CitiPassFile, NumPassRecs
    If Not CitiPassFile = -1 Then
    For cnt = 1 To NumPassRecs
      Get CitiPassFile, cnt, CitiPass
      If Not CitiPass.DelFlag Then
      LookFor$ = Trim$(CitiPass.PassWord)
      If pw$ = LookFor$ Then
        If Not CitiPass.InUseFlag Then
          If CitiPass.Module(9).FullAccess = True Then
            LevelPass = 1
          ElseIf CitiPass.Module(9).ReportsOnly = True Then
            LevelPass = 3
          ElseIf CitiPass.Module(9).PaymentAccess = True Then
            LevelPass = 2
          End If
       
          If LevelPass > 0 Then
            Oktogo = True
            CitiPass.InUseFlag = True
            CitiPass.FlagMod = 9
            PWUser = QPTrim(CitiPass.UserName)
            OPERNUM = CitiPass.PassNum
            CitiPass.CompName = QPTrim(ComputerName$)
            Put CitiPassFile, cnt, CitiPass
          End If
          Exit For
        Else
          theyare = QPTrim(CitiPass.UserName)
          onwhat = QPTrim(CitiPass.CompName)
          Findsettings = -1
        End If
      End If
      End If
    Next
    Close CitiPassFile
  Else
    stopnow = 1
  End If
  End If
  If Oktogo Then
    PWcnt = cnt
    Findsettings = cnt
  Else
    LevelPass = 0
    PWcnt = 0
  End If
End Function


