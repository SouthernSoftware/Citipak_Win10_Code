VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCMMainMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "v2.05 Citipak Cash Management"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   1890
   ClientWidth     =   12210
   ClipControls    =   0   'False
   Icon            =   "frmCMMainMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   960
      Top             =   3384
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   960
      Top             =   2952
   End
   Begin VB.CommandButton cmdCMSetUpMenu 
      BackColor       =   &H008F8265&
      Caption         =   "Setup/Maintenance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3846
      MaskColor       =   &H8000000F&
      TabIndex        =   2
      Top             =   4888
      Width           =   4524
   End
   Begin VB.CommandButton cmdReportsMenu 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Reports"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3846
      TabIndex        =   1
      Top             =   4088
      Width           =   4524
   End
   Begin VB.CommandButton cmdPaymentMenu 
      Caption         =   "Enter &Payments/Deposits"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3864
      TabIndex        =   0
      Top             =   3288
      Width           =   4524
   End
   Begin VB.CommandButton cmdExitCM 
      Caption         =   "E&XIT Cash Management"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3846
      TabIndex        =   3
      Top             =   5688
      Width           =   4524
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "2:23 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "5/14/2018"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "CASH MANAGEMENT MAIN MENU"
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
      Left            =   3348
      TabIndex        =   4
      Top             =   1176
      Width           =   5292
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1788
      Top             =   744
      Width           =   8652
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1788
      Top             =   624
      Width           =   8652
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   9828
      X2              =   9828
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   8868
      X2              =   8868
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   3
      X1              =   8868
      X2              =   9828
      Y1              =   1944
      Y2              =   1944
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   8868
      X2              =   9828
      Y1              =   2064
      Y2              =   2064
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   8868
      Top             =   1824
      Width           =   972
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   8988
      X2              =   9708
      Y1              =   7944
      Y2              =   7944
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   8988
      X2              =   8988
      Y1              =   2064
      Y2              =   7944
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8988
      Top             =   2064
      Width           =   732
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   3348
      X2              =   3348
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2388
      X2              =   2388
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   2388
      X2              =   3348
      Y1              =   1944
      Y2              =   1944
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2388
      X2              =   3348
      Y1              =   2064
      Y2              =   2064
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2388
      Top             =   1824
      Width           =   972
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2508
      X2              =   3228
      Y1              =   7944
      Y2              =   7944
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2508
      X2              =   2508
      Y1              =   2064
      Y2              =   7944
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2508
      Top             =   2064
      Width           =   732
   End
End
Attribute VB_Name = "frmCMMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
'Dim FormOver As clsFormOverRider
Private Temp_Class As Resize_Class
'LevelPass 1 is Full Access, 2 is Payments, 3 is Reports Only

Private Sub cmdPaymentMenu_Click()
  Dim FntSize As Integer, RecpPort As String
  Dim RP As Integer, lenRP As Integer
  Dim RcptPrnFile As ReceiptPRNType
  ReDim MsgText(0 To 5) As String
  If LevelPass < 3 Then
    frmInfo.Label1 = "Verifying Receipt Printer..."
    frmInfo.Show
    DoEvents
    If Not Exist(UBPath$ + "CMSetTown.DAT") Then
      Unload frmInfo
      frmMsgDialog.RetLabel = "-2"
      CMLog "ERROR: NO Setup Info"
      FntSize = frmMsgDialog.Label(3).FontSize
      frmMsgDialog.Label(1).FontSize = (FntSize + 2)
      frmMsgDialog.Label(3).FontSize = (FntSize + 2)
      MsgText(0) = "ERROR:"
      MsgText(1) = ""
      MsgText(2) = "NO SETUP INFORMATION!"
      MsgText(3) = "Please Complete Setup First."
      MsgText(4) = ""
      MsgText(5) = ""
      GetOKorNot MsgText(), True
      Exit Sub
    End If
    If Not Exist(RcptFileName$) Then
      Unload frmInfo
      ReDim MsgText(0 To 5) As String
      FntSize = frmMsgDialog.Label(1).FontSize
      frmMsgDialog.Label(1).FontSize = (FntSize + 2)
      MsgText(0) = "WARNING:"
      MsgText(1) = ""
      MsgText(2) = "RECEIPT SETUP FILE NOT FOUND!"
      MsgText(3) = "If you continue receipt printing"
      MsgText(4) = "will be disabled."
      MsgText(5) = "Receipt setup option is on CitiPak Main Menu."
      If GetOKorNot(MsgText()) Then
        UBLog "USER WANTS TO CONTINUE!"
      Else
        UBLog "USER ABORTED."
        Exit Sub
      End If
    Else
      RP = FreeFile
      lenRP = Len(RcptPrnFile)
      Open RcptFileName$ For Random Shared As RP Len = lenRP
      Get RP, 1, RcptPrnFile
      RecpPort = QPTrim(RcptPrnFile.RcpPort)
      Close
      If RcptPrnFile.PrnDefYN = 1 Then
        On Local Error GoTo noprnfound
        Open RecpPort For Output As RP
        Close RP
      End If
    End If
    Unload frmInfo
    Load frmCMPaySource
    Unload Me
    frmCMPaySource.Show
    DoEvents
  Else
    MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
  End If
Exit Sub
noprnfound:
        Unload frmInfo
        ReDim MsgText(0 To 5) As String
        FntSize = frmMsgDialog.Label(1).FontSize
        frmMsgDialog.Label(1).FontSize = (FntSize + 2)
        MsgText(0) = "WARNING:"
        MsgText(1) = ""
        MsgText(2) = "RECEIPT PRINTER NOT FOUND!"
        MsgText(3) = "If you continue receipt printing"
        MsgText(4) = "will be disabled."
        MsgText(5) = "Receipt setup option is on CitiPak Main Menu."
        If GetOKorNot(MsgText()) Then
          UBLog "USER WANTS TO CONTINUE!"
          Load frmCMPaySource
          Unload Me
          frmCMPaySource.Show
          DoEvents
        Else
          UBLog "USER ABORTED."
          Exit Sub
        End If
End Sub

Private Sub cmdCMSetUpMenu_Click()
  If LevelPass = 1 Then
    Load frmCMSetupMenu
    DoEvents
    frmCMSetupMenu.Show
    Unload Me
  Else
    MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
  End If
End Sub

Private Sub cmdReportsMenu_Click()
  Load frmCMReportMenu
  DoEvents
  frmCMReportMenu.Show
  Unload Me
End Sub

Private Sub Form_Load()
  On Error GoTo Cancel
10:  Set Temp_Class = New Resize_Class
11:  Temp_Class.InitResizeClass Me
12:  Set Over = New clsTextBoxOverRider
13:  Over.OverRide Me
14:  StatusBar1.Panels.Item(1).Text = TownName$
15:  If DelayExit = True Then
16:    DelayExit = False
17:    Timer2.Enabled = True
18:  Else
19:    cmdExitCM.Enabled = True
20:  End If
Cancel:
   If Err > 0 Then
    MsgBox "Error Code Was " + Err.Description + Str$(Err) + " (CMMenu Load - Line:" & Erl & ")"
  End If
  Close
  Exit Sub
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    If cmdExitCM.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        Call CMLog("Close via Main CM Menu" + PWUser$)
        CitiTerminate
      End If
    End If
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ''' Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub


Private Sub cmdExitCM_Click()
  Call CMTerminate
  frmCMMainMenu.Enabled = False
  Timer1.Enabled = True
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyHome
      cmdPaymentMenu.SetFocus
    Case vbKeyEnd
      cmdExitCM.SetFocus
    Case Else:
  End Select
End Sub
Private Sub Timer1_Timer()
  Unload Me
End Sub

Private Sub Timer2_Timer()
  cmdExitCM.Enabled = True
End Sub

