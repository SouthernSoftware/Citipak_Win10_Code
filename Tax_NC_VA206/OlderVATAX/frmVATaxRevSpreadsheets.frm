VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Begin VB.Form frmVATaxRevSpreadsheets 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Revenue Setup Spreadsheets"
   ClientHeight    =   8670
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   10665
   Icon            =   "frmVATaxRevSpreadsheets.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   10665
   StartUpPosition =   1  'CenterOwner
   Begin FPSpread.vaSpread vaSpread2 
      Height          =   3180
      Left            =   1680
      TabIndex        =   3
      Top             =   4440
      Width           =   7260
      _Version        =   196613
      _ExtentX        =   12806
      _ExtentY        =   5609
      _StockProps     =   64
      ColsFrozen      =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   3
      MaxRows         =   10
      SpreadDesigner  =   "frmVATaxRevSpreadsheets.frx":08CA
      VisibleCols     =   3
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   636
      Left            =   858
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7920
      Width           =   4152
      _Version        =   131072
      _ExtentX        =   7324
      _ExtentY        =   1122
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
      ButtonDesigner  =   "frmVATaxRevSpreadsheets.frx":0CE7
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdReset 
      Height          =   636
      Left            =   5646
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7920
      Width           =   4152
      _Version        =   131072
      _ExtentX        =   7324
      _ExtentY        =   1122
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
      ButtonDesigner  =   "frmVATaxRevSpreadsheets.frx":0EDE
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3132
      Left            =   2400
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Width           =   5892
      _Version        =   196613
      _ExtentX        =   10393
      _ExtentY        =   5525
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      ArrowsExitEditMode=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   4
      MaxRows         =   8
      SpreadDesigner  =   "frmVATaxRevSpreadsheets.frx":10D4
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   3372
      Left            =   1542
      Top             =   4320
      Width           =   7572
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "Personal Revenue"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   252
      Left            =   4128
      TabIndex        =   2
      Top             =   3960
      Width           =   2412
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   3372
      Left            =   2262
      Top             =   480
      Width           =   6132
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "Real Revenue"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   252
      Left            =   4128
      TabIndex        =   1
      Top             =   120
      Width           =   2412
   End
End
Attribute VB_Name = "frmVATaxRevSpreadsheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class
  Dim StrEmpty As Boolean
  Dim PenIdx As Integer
Private Sub cmdExit_Click()
  Me.Hide
End Sub

Private Sub cmdReset_Click()
  Dim x As Integer
  
  For x = 1 To 8
    vaSpread1.Col = 1
    vaSpread1.Row = x
    vaSpread1.Text = Sprd1Col1(x)
    vaSpread1.Col = 2
    vaSpread1.Row = x
    vaSpread1.Text = Sprd1Col2(x)
  Next x
  
  For x = 1 To 10
    vaSpread2.Col = 1
    vaSpread2.Row = x
    vaSpread2.Text = Sprd2Col1(x)
    vaSpread2.Col = 2
    vaSpread2.Row = x
    vaSpread2.Text = Sprd2Col2(x)
    vaSpread2.Col = 3
    vaSpread2.Row = x
    vaSpread2.Text = Sprd2Col3(x)
  Next x
  
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  MainLog ("User opened frmVATaxRevSpreadsheets.")
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    'Temp_Class.ResizeControls Me
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
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF3:
      SendKeys "%R"
      Call cmdReset_Click
      KeyCode = 0
  End Select
End Sub

Private Sub vaSpread1_Change(ByVal Col As Long, ByVal Row As Long)
  Dim PenCnt As Integer
  Dim x As Integer
  Dim CntPens As Integer
  Dim Thisx As Integer
  
  On Error GoTo ERRORSTUFF
  
  StrEmpty = False
  vaSpread1.Col = 1
  vaSpread1.Row = Row
  If QPTrim$(vaSpread1.Text) = "" Then
    StrEmpty = True
  End If
  vaSpread1.Col = 2
  If vaSpread1.Text = "1" And StrEmpty = True Then
    Call TaxMsg(800, "This row contains an unused optional revenue. Setting interest is not allowed for this row.")
    vaSpread1.Text = "0"
    vaSpread1.SetFocus
    vaSpread1.SetActiveCell 2, Row
    Exit Sub
  End If
  
  If Col <> 3 Then Exit Sub
  CntPens = 0
  Thisx = 0
  vaSpread1.Col = 3
  For x = 5 To 7
    vaSpread1.Row = x
    If vaSpread1.Text = "1" Then
      CntPens = CntPens + 1
      Thisx = x
    End If
  Next x
  
  If CntPens > 1 Then
    Call TaxMsg(800, "ERROR: Only one optional revenue can be earmarked as the penalty revenue. Please review your penalty selections and select only one penalty revenue.")
    vaSpread1.SetActiveCell 3, Thisx
    Exit Sub
  End If
  
  vaSpread1.Col = Col
  vaSpread1.Row = Row
    
  If vaSpread1.Text = "1" And PenIdx <> Row Then
    vaSpread1.Col = 1
    If QPTrim$(vaSpread1.Text) = "" Then
      Call TaxMsg(800, "The penalty revenue source has been assigned to row " + CStr(Row) + ". Please enter a penalty description.")
      vaSpread1.SetActiveCell 1, Row
    End If
    vaSpread1.Col = Col
    If PenIdx > 0 Then
      vaSpread1.Row = PenIdx
      vaSpread1.Value = "0"
    End If
    PenIdx = Row
  End If
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxSystemSetup", "vaSpread1_Change", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate

End Sub

Private Sub vaSpread1_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub vaSpread2_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub
