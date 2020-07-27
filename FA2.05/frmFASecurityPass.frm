VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "EDT32X30.OCX"
Begin VB.Form frmFASecurityPass 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CitiPak Password Sign-In"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11652
   Icon            =   "frmFASecurityPass.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      Left            =   6120
      TabIndex        =   2
      Top             =   5496
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
      Left            =   4224
      TabIndex        =   1
      Top             =   5496
      Width           =   1404
   End
   Begin EditLib.fpText txtPW 
      Height          =   420
      Left            =   4728
      TabIndex        =   0
      Top             =   4440
      Width           =   2196
      _Version        =   196608
      _ExtentX        =   3873
      _ExtentY        =   741
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   2940
      Top             =   2616
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Fixed Assets Password"
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
      Left            =   3702
      TabIndex        =   4
      Top             =   3864
      Width           =   4260
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000E&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      Height          =   2868
      Index           =   0
      Left            =   2952
      Top             =   3480
      Width           =   5748
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "FIXED ASSETS PASSWORD"
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
      Left            =   3210
      TabIndex        =   3
      Top             =   2856
      Width           =   5292
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Index           =   1
      Left            =   2940
      Top             =   2520
      Width           =   5772
   End
End
Attribute VB_Name = "frmFASecurityPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsFATextBoxOverRider
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
  Unload frmFASecurityPass
End Sub

Private Sub cmdEnter_Click()
  Dim Notvalid As Boolean, z As String, Pz As String, cnt As Integer
  Dim Az As String, CntA As Integer, Findnum As Integer, PassOK As Integer
  Dim InHandle As Integer
  
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
      Unload frmFASecurityPass
      frmFAMainMenu.Show
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
          Call MainLog("In Fixed Assets with Level " & LevelPass) 'record it
          frmFAMainMenu.Show
          DoEvents
          Unload frmFASecurityPass
        Else
          Notvalid = True
          If Findnum = -1 Then 'someone else is already using payroll.exe
            MsgBox "Password for User " + theyare$ + " In Session on " + onwhat$, vbOKOnly, "Access Denied"
            Call MainLog("Password in session, NO Fixed Asset Access.")
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
      Call MainLog("Password Maintenance in session, FA Access Denied.")
      cmdCancel.SetFocus
    End If
  End If
  If Notvalid = True Then
    MsgBox "Invalid Password. Try again or See Password Administrator.", vbOKOnly, "Invalid Entry"
    Call MainLog("Invalid FA Password. " + txtPW)  'record invalid password attempt
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
  
  InUseExitFlag = False
  PWcnt = 0
  LoggedIn = 0
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
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
End Sub
Private Sub Form_Resize()
  Temp_Class.ResizeControls Me
  DoEvents
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
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
  Dim Oktogo As Boolean, Lookfor As String
  Dim x As Integer
  Dim Y As Integer
  Dim InFAUse As Boolean
  
  InFAUse = False 'InFAUse tells us that somebody else with full access
  'is already logged in to fixedassets.exe
  
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
              If CitiPass.Module(5).FullAccess = True Then 'data saved when passwords initialized ...if true this user is OK
                If InFAUse = True Then 'InFAUse is set in Check4InUse ...OK someone is already in so alert
                  'the user
                  InFAUse = False 'reset for next user's check
'                  DoWhatFlag = PromptUserAlreadyActive(Me)
                  frmFAWarnInUse.Show vbModal
                  
                  Select Case frmFAWarnInUse.fptxtHide.Text
                  
                  Case "Continue"
                    Unload frmFAWarnInUse
                  Case "Exit"
                    InUseExitFlag = True 'tells Enter Sub to exit
                    Unload frmFAWarnInUse
                    Close
                    Exit Function
                  Case Else
                  End Select
                End If
                CitiPass.FlagMod = 5 'set this field to 5 denoting FA is
                'now occupied with a full access user at PWcnt
                LevelPass = 1
              Else 'CitiPass.Module(5).FullAccess = False
                If CitiPass.Module(5).ReportsOnly = True Then
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
            Else 'CitiPass.InUseFlag is true so this user is already in fixed assets
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
    PWcnt = cnt 'this is used when fixed assets exits so that only this
    'user's password data is cleared
    Findsettings = cnt
  Else
    LevelPass = 0
    PWcnt = 0
  End If
  
  Exit Function
  
Check4InUse:
  OpenCitiPassFile CitiPassFile, NumPassRecs
  For Y = 1 To NumPassRecs
    Get CitiPassFile, Y, CitiPass
    If CitiPass.FlagMod = 5 Then '5 indicates fixed assets
    'if FlagMod = 5 then we know another full access user is
    'in fixedassets.exe now
      InFAUse = True
      Exit For
    End If
  Next Y
  Close CitiPassFile
  Return
End Function


