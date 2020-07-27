VERSION 5.00
Begin VB.Form frmPassWord 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2484
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   5568
   Icon            =   "frmPassWord.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2484
   ScaleWidth      =   5568
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPW 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      IMEMode         =   3  'DISABLE
      Left            =   3018
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   612
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
      Left            =   1152
      TabIndex        =   1
      Top             =   1380
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
      Left            =   3066
      TabIndex        =   2
      Top             =   1380
      Width           =   1404
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
      Left            =   1152
      TabIndex        =   3
      Top             =   708
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
Public Callingfrm As Integer
'***********************
' CALLINGFRM CODES
' 1 = Void Payment
' 2 = Adjustment
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
  Dim Notvalid As Boolean, CMSetuplen As Integer, cnt As Integer
  Dim FileHandle As Integer, Pz As String, z As String
  ReDim CMSetup(1) As CMSetupType
  CMSetuplen = Len(CMSetup(1))
  LoadCMSetUpFile CMSetup(), CMSetuplen

  Notvalid = False
  'Select Case txtPW
  Pz$ = ""
  z$ = txtPW
  For cnt = 1 To Len(z$)
    Pz$ = Pz$ + Chr$(Asc(Mid$(z$, cnt, 1)) Xor 127)
  Next
  
  If Pz$ = "1010-8>16<" Then
      'Software Support
      PWcnt = 0
      OperNum = 0
      CMLog "Support Sign in Void"
    If Callingfrm = 1 Then
      Load frmVoidSearch
      Unload Me
      frmVoidSearch.Show
    End If
    If Callingfrm = 2 Then
      'do the adjustment thang for util
      Load frmUBAdjustmentEntry
      Unload Me
      frmUBAdjustmentEntry.Show
    End If
    If Callingfrm = 3 Then
      Load frmBLAdjustBal
      Unload Me
      frmBLAdjustBal.Show
    End If
    If Callingfrm = 4 Then
      Load frmTaxAdjustments
      Unload Me
      frmTaxAdjustments.Show
    End If
    If Callingfrm = 5 Then
      Load frmVATaxAdjustments
      Unload Me
      frmVATaxAdjustments.Show
    End If
    If Callingfrm = 6 Then
      Load frmVATaxPAdjustments
      Unload Me
      frmVATaxPAdjustments.Show
    End If
  Else
    If Callingfrm = 1 Then
      If Pz$ = QPTrim(CMSetup(1).VoidPW) Then
        Load frmVoidSearch
        Unload Me
        frmVoidSearch.Show
        Unload frmCMPaySource
      Else
        Notvalid = True
      End If
    ElseIf Callingfrm = 2 Then
      If Pz$ = QPTrim(CMSetup(1).AdjPW) Then
          'do the adjustment
        Load frmUBAdjustmentEntry
        Unload Me
        frmUBAdjustmentEntry.Show
      Else
        Notvalid = True
      End If
    ElseIf Callingfrm = 3 Then
      If Pz$ = QPTrim(CMSetup(1).AdjPW) Then
          'do the adjustment
        Load frmBLAdjustBal
        Unload Me
        frmBLAdjustBal.Show
      Else
        Notvalid = True
      End If
    ElseIf Callingfrm = 4 Then
      If Pz$ = QPTrim(CMSetup(1).AdjPW) Then
          'do the adjustment
        Load frmTaxAdjustments
        Unload Me
        frmTaxAdjustments.Show
      Else
        Notvalid = True
      End If
    ElseIf Callingfrm = 5 Then
      If Pz$ = QPTrim(CMSetup(1).AdjPW) Then
        Load frmVATaxAdjustments
        Unload Me
        frmVATaxAdjustments.Show
      Else
        Notvalid = True
      End If
    ElseIf Callingfrm = 6 Then
      If Pz$ = QPTrim(CMSetup(1).AdjPW) Then
        Load frmVATaxPAdjustments
        Unload Me
        frmVATaxPAdjustments.Show
      Else
        Notvalid = True
      End If
    End If
  End If
  Erase CMSetup
  If Notvalid = True Then
    Call CMLog("Invalid Password : " + txtPW)
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
 
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Me.Width = vWidth
  Me.Height = vHeight
  Me.Left = vLeft
  Me.Top = vTop
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

