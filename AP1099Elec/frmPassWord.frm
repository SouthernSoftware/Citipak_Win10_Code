VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmPassWord 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3036
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   6408
   ForeColor       =   &H00000000&
   Icon            =   "frmPassWord.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3036
   ScaleWidth      =   6408
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEnter 
      BackColor       =   &H00D0D0D0&
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
      Left            =   1541
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1932
      Width           =   1404
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00D0D0D0&
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
      Left            =   3464
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1932
      Width           =   1404
   End
   Begin EditLib.fpText txtPW 
      Height          =   420
      Left            =   3060
      TabIndex        =   0
      Top             =   1128
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
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "For Access Contact Software Support with User Code:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   540
      Left            =   576
      TabIndex        =   5
      Top             =   288
      Visible         =   0   'False
      Width           =   3012
   End
   Begin VB.Label lblcode 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3588
      TabIndex        =   4
      Top             =   288
      Visible         =   0   'False
      Width           =   2148
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password"
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
      Left            =   1164
      TabIndex        =   3
      Top             =   1224
      Width           =   1692
   End
End
Attribute VB_Name = "frmPassWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim vWidth%, vHeight%, vTop%, vLeft%
Public Callingfrm As Integer, tmp As String
'***********************
' CALLINGFRM CODES
' 1 = GLUTIL MENU FROM CONFIGMENU *** SOSOFT
' 2 = GLCLOSING FROM SETUPMENU ****** CLOSE
'**********************

Private Sub cmdCancel_Click()
  Unload frmPassWord
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    Cancel = True
  End If
End Sub

Private Sub cmdEnter_Click()
  Dim Notvalid As Boolean, Pz As String, Z As String, cnt As Integer
  Dim FileHandle As Integer, WhosOnFirst As String
  Notvalid = False
  Pz$ = ""
  Z$ = txtPW
  If Callingfrm = 1 Then
    For cnt = 1 To Len(Z$)
      Pz$ = Pz$ + Chr$(Asc(Mid$(Z$, cnt, 1)) Xor 127)
    Next
    
    If Pz$ = "1010-8>16<" Then
      Call MainLog("Support Opened GL Util Menu")
      Unload frmPassWord
      frmGLUtilMenu.Show
      Unload frmGLConfigUtilMenu
  
    ElseIf Z$ = tmp$ Then
      Call MainLog("Opened GL Util Menu with " + tmp$)
      Unload frmPassWord
      frmGLUtilMenu.Show
      Unload frmGLConfigUtilMenu
    Else
      Notvalid = True
    End If
'  ElseIf Callingfrm = 2 Then
'    If txtPW = "CLOSE" Then
'        If Exist("FClose.opn") Then
'          FileHandle = FreeFile
'          Open "FClose.opn" For Input As FileHandle
'          Line Input #FileHandle, WhosOnFirst$
'          Close FileHandle
'          MsgBox "The Close Out Menu Has Been Opened By: " + WhosOnFirst$, vbOKOnly, "Menu Not Accessible"
'          Call MainLog("Close Year, Access Denied.")
'        Else
'          FileHandle = FreeFile
'          Open "FClose.opn" For Output As FileHandle
'          Print #FileHandle, ComputerName$
'          Close FileHandle
'          Call MainLog("Opened Close Year Menu.")
'          Unload frmPassWord
'          frmGLClosingOpMenu.Show
'          Unload frmGLSetupMenu
'        End If
'     Else
'        Notvalid = True
'      End If
  End If
 If Notvalid = True Then
    Call MainLog("Invalid Password : " + txtPW + "from " + Str(Callingfrm))
    MsgBox "Invalid Password. Try again or Call Software Support.", vbOKOnly, "Invalid Entry"
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
  Dim cnt&, tmpstr$
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Me.Width = vWidth
  Me.Height = vHeight
  Me.Left = vLeft
  Me.Top = vTop
  tmp$ = Mid(Timer, 1, 8)
  For cnt = 1 To Len(tmp$)
    tmpstr$ = tmpstr$ + Chr$(Asc(Mid$(tmp$, cnt, 1)) Xor 127)
  Next
  lblcode.Caption = tmpstr$
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
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
      SendKeys "%C"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%E"
      KeyCode = 0
    Case Else:
  End Select
End Sub

