VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{A18D4668-91EF-101C-84A6-BA990A365A4E}#3.0#0"; "mem32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmViewPrint 
   BackColor       =   &H80000016&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Print"
   ClientHeight    =   8400
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   10512
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmViewPrint.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   10512
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAlignment 
      Caption         =   "F5 &Align"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2736
      TabIndex        =   5
      Top             =   7896
      Width           =   1140
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F8 &Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5844
      TabIndex        =   4
      Top             =   7896
      Width           =   1188
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "ESC E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1248
      TabIndex        =   3
      Top             =   7896
      Width           =   1116
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "F10 &Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7500
      TabIndex        =   2
      Top             =   7896
      Width           =   1308
   End
   Begin VB.CommandButton cmdPrnScn 
      Caption         =   "F7 Prin&t Screen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3948
      TabIndex        =   1
      Top             =   7896
      Width           =   1836
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8880
      Top             =   7872
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      DialogTitle     =   "Print"
   End
   Begin MemoLib.fpMemo fpMemo1 
      CausesValidation=   0   'False
      Height          =   7740
      Left            =   24
      TabIndex        =   0
      Top             =   -48
      Width           =   10452
      _Version        =   196608
      _ExtentX        =   18436
      _ExtentY        =   13652
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
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
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   0
      CaretOverWrite  =   3
      HideSelection   =   -1  'True
      NullColor       =   -2147483637
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   1
      Text            =   "fpMemo1"
      WordWrap        =   0   'False
      ShowEOL         =   0   'False
      SelMode         =   0
      LineLimit       =   2147483647
      ScrollBars      =   3
      PageWidth       =   0
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ProcessTab      =   0   'False
      TabLength       =   0
      AutoMenu        =   0   'False
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn btnPgUp 
      Height          =   384
      Left            =   9336
      TabIndex        =   6
      Top             =   7872
      Width           =   444
      _Version        =   131072
      _ExtentX        =   783
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmViewPrint.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn btnPgDn 
      Height          =   384
      Left            =   9864
      TabIndex        =   7
      Top             =   7872
      Width           =   444
      _Version        =   131072
      _ExtentX        =   783
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmViewPrint.frx":2C84
   End
End
Attribute VB_Name = "frmViewPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim strReportFile As String
Dim alnRpt As String
Public PgNum As Integer
Public NoPbox As Boolean
Public thePrn As String
'Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim vWidth%, vHeight%, vTop%, vLeft%
Dim doAlign As Boolean
'''Public Declare Function PrinterProperties Lib "winspool.drv" (ByVal hwnd As Long, ByVal hPrinter As Long) As Long

Private Sub btnPgDn_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 1 Then
    DoEvents
    SendKeys "{PgDn}", True
  End If
End Sub

Private Sub btnPgUp_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 1 Then
    DoEvents
    SendKeys "{PgUp}", True
  End If
End Sub

Private Sub cmdPrint_Click()
  If NoPbox Then
    PrintWSet thePrn, 1
  Else
    frmPrint.Show 1
  End If
End Sub

Private Sub cmdAlignment_Click()
  doAlign = True
  frmPrint.Show 1
End Sub

Private Sub cmdExit_Click()
  On Local Error Resume Next
  Unload frmViewPrint
End Sub
Public Sub PrintWSet(DefPrinter As String, Copies As Integer)
  Dim RptHandle As Integer, LPTHandle As Integer
  Dim RptA As Integer, LPTA As Integer, ToPrintA As String
  Dim ToPrint As String, CopyLoop As Integer
  On Error GoTo Cancel
  'Printer.Print
'''  to strReportFile DefPrinter'[ADDITIVE] | PortName]
10:
  If doAlign = True Then
    GoSub PrintAlignMask
    Exit Sub
  End If
20:
 ' MsgBox "Printer -" + DefPrinter, vbOKOnly
'    LPTHandle = FreeFile
'21: Open DefPrinter For Output As LPTHandle Len = 254
'22: Print #LPTHandle, "Hello"
'23: Close LPTHandle
  For CopyLoop = 1 To Copies
    LPTHandle = FreeFile
    Open DefPrinter For Output As LPTHandle Len = 254
    RptHandle = FreeFile
30:
    Open strReportFile For Input As RptHandle
40:
    Do
      If frmPrint.cmdCancel = False Then
45:
        Line Input #RptHandle, ToPrint$
46:
        ToPrint$ = RTrim$(ToPrint$)
47: '    SmallPause

48:     Print #LPTHandle, ToPrint$
49:   Else
50:
        Exit Do
        'Printer.EndDoc
      End If
    Loop Until eof(RptHandle)
60:
    Close RptHandle
62:
    Close LPTHandle
65:
    Next CopyLoop
68:
 Printer.EndDoc
70:
  If strReportFile = "APCHECK.PRN" Then
    MsgBox "Check Printing Complete", vbOKOnly, "Procedure Complete"
    Unload frmViewPrint
    Exit Sub
  End If
80:
  Exit Sub
PrintAlignMask:
    LPTA = FreeFile
81:
    Open DefPrinter For Output As LPTA
    RptA = FreeFile
82:
    Open alnRpt For Input As RptA
    Do
83:
      If frmPrint.cmdCancel = False Then
        Line Input #RptA, ToPrintA$
84:
        ToPrintA$ = RTrim$(ToPrintA$)
        Print #LPTA, ToPrintA$
85:
      Else
        Exit Do
        Printer.EndDoc
      End If
86:
    Loop Until eof(RptA)
87:
    Close LPTA, RptA
    Printer.EndDoc
88:
    If MsgBox("Do You Wish to Print Another Mask?", vbYesNo, "Print Mask") = vbYes Then
      GoSub PrintAlignMask
    End If
    doAlign = False
Cancel:
'  If Erl = 60 Or Erl = 62 Then
'    Resume Next
'  End If
  If Err > 0 Then
    MsgBox "Error Code Was " + DefPrinter + Err.Description + Str$(Err) + " (PrintWSet - Line:" & Erl & ")"
  End If
  Close
  Exit Sub
End Sub

Property Get ReportName() As String
  ReportName = strReportFile
End Property

Property Let ReportName(ByVal strNewReportName As String)
  strReportFile = strNewReportName
End Property

Property Get AlignRpt() As String
  AlignRpt = alnRpt
End Property

Property Let AlignRpt(ByVal alnRptName As String)
  alnRpt = alnRptName
End Property

Private Sub cmdPrnScn_Click()
  PrintForm
End Sub

Private Sub cmdSave_Click()
  Dim newrpt As String, newlen As Integer
  newlen = (Len(strReportFile) - 3)
  newrpt = Mid$(strReportFile, 1, newlen) + "txt"
  If MsgBox("Do You Wish to Save this Report - " & strReportFile, vbYesNo, "Save Report") = vbYes Then
    fpMemo1.SaveFile newrpt
    'CpyRptFile strReportFile
    MsgBox "The Report was saved in the Citipak Directory as " & newrpt, vbOKOnly, "Report Saved"
  End If
End Sub

Private Sub Form_Initialize()
  vWidth = Screen.Width * 0.9    ' Set width of form.
  vHeight = Screen.Height * 0.85   ' Set height of form.
  vLeft = (Screen.Width - vWidth) \ 2   ' Center form horizontally.
  vTop = (Screen.Height - vHeight) \ 2   ' Center form vertically.
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  DoEvents
  Me.Width = vWidth
  Me.Height = vHeight
  Me.Left = vLeft
  Me.Top = vTop
  DoEvents
  Me.fpMemo1.LoadFile strReportFile
  DoEvents
  'StatusBar1.Panels.Item(1).Text = GLUserName
  doAlign = False
End Sub
Private Sub Form_Resize()
  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      SendKeys "%X"
      'OhStop = True
      KeyCode = 0
    Case vbKeyF8:
      SendKeys "%P"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      KeyCode = 0
    Case vbKeyF7
      SendKeys "%t"
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%A"
      KeyCode = 0
    Case Else:
  End Select
End Sub


'Private Sub fpMemo1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  If Button = 2 Then
'    Call cmdExit_Click
'  End If
'End Sub

'Private Sub fpMemo1_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeyHome Then
'      fpMemo1.LineIndex = 0
'      KeyCode = 0
'  ElseIf KeyCode = vbKeyEnd Then
'      fpMemo1.LineIndex = fpMemo1.LineCount
'      KeyCode = 0
'  End If
'End Sub

'Private Sub fpMemo1_KeyUp(KeyCode As Integer, Shift As Integer)
'
'End Sub

