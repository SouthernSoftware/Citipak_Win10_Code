VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A18D4668-91EF-101C-84A6-BA990A365A4E}#3.0#0"; "mem32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmViewPrint 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Print"
   ClientHeight    =   8424
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   10476
   ForeColor       =   &H00000000&
   Icon            =   "frmViewPrint.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8424
   ScaleWidth      =   10476
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAlignment 
      BackColor       =   &H00D0D0D0&
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
      Height          =   396
      Left            =   420
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7584
      Width           =   1428
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00D0D0D0&
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
      Height          =   396
      Left            =   4620
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7584
      Width           =   1596
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Esc E&xit"
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
      Left            =   8604
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7584
      Width           =   1452
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00D0D0D0&
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
      Height          =   396
      Left            =   6612
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7584
      Width           =   1596
   End
   Begin VB.CommandButton cmdPrnScn 
      BackColor       =   &H00D0D0D0&
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
      Height          =   396
      Left            =   2244
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7584
      Width           =   1980
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   7800
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      DialogTitle     =   "Print"
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   1
      Top             =   8172
      Width           =   10476
      _ExtentX        =   18479
      _ExtentY        =   445
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6117
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   6117
            TextSave        =   "1:58 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   6117
            TextSave        =   "10/6/2004"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MemoLib.fpMemo fpMemo1 
      Height          =   7452
      Left            =   24
      TabIndex        =   0
      Top             =   -48
      Width           =   10452
      _Version        =   196608
      _ExtentX        =   18436
      _ExtentY        =   13144
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
      BackColor       =   16777215
      ForeColor       =   0
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   8421504
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
      BorderWidth     =   1
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      CaretInsert     =   0
      CaretOverWrite  =   3
      HideSelection   =   -1  'True
      NullColor       =   16777215
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
      SelBackColor    =   12632256
      SelForeColor    =   0
      BorderGrayAreaColor=   12632256
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   8421504
      BorderDropShadowWidth=   3
      ProcessTab      =   0   'False
      TabLength       =   0
      AutoMenu        =   0   'False
      OLEDropMode     =   0
      OLEDragMode     =   0
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
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'  If ((UnloadMode = vbFormControlMenu)) Then
'    Cancel = True
'  End If
'End Sub

Private Sub cmdAlignment_Click()
  doAlign = True
  frmPrint.Show 1
End Sub

Private Sub cmdExit_Click()
  Unload frmViewPrint
End Sub
Private Sub cmdPrint_Click()
If NoPbox Then
  PrintWSet thePrn, 1
Else
  frmPrint.Show 1
End If
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
  
  For CopyLoop = 1 To Copies
    LPTHandle = FreeFile
    Open DefPrinter For Output As LPTHandle
    RptHandle = FreeFile
30:
    Open strReportFile For Input As RptHandle
40:
    Do
      If frmPrint.cmdCancel = False Then
45:
        Line Input #RptHandle, ToPrint$
        
        ToPrint$ = RTrim$(ToPrint$)
        Print #LPTHandle, ToPrint$
      Else
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
    Open DefPrinter For Output As LPTA
    RptA = FreeFile
    Open alnRpt For Input As RptA
    Do
      If frmPrint.cmdCancel = False Then
        Line Input #RptA, ToPrintA$
        
        ToPrintA$ = RTrim$(ToPrintA$)
        Print #LPTA, ToPrintA$

      Else
        Exit Do
        Printer.EndDoc
      End If
    Loop Until eof(RptA)
    Close LPTA, RptA
    Printer.EndDoc
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

'Private Sub Command1_Click()
'  Dim RptHandle As Integer, LPTHandle As Integer
'  Dim ToPrint As String, CopyLoop As Integer
'  LPTHandle = FreeFile
'  On Error GoTo Cancel
'  CommonDialog1.PrinterDefault = True
'  'CommonDialog1.Min = 1
'  'CommonDialog1.Max = PgNum
'  'Printer.Orientation
'
'  CommonDialog1.ShowPrinter
'  For CopyLoop = 1 To CommonDialog1.Copies
'    Open Printer.Port For Output As LPTHandle
'    RptHandle = FreeFile
'    Open strReportFile For Input As RptHandle
'    Do
'      If CommonDialog1.CancelError = False Then
'        Line Input #RptHandle, ToPrint$
'        ToPrint$ = RTrim$(ToPrint$)
'        Print #LPTHandle, ToPrint$
'      Else
'        Exit Do
'        Printer.EndDoc
'      End If
'    Loop Until EOF(RptHandle)
'    Close LPTHandle, RptHandle
'    Next CopyLoop
'  Printer.EndDoc
'
'Cancel:
'End Sub

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
  Me.Width = vWidth
  Me.Height = vHeight
  Me.Left = vLeft
  Me.Top = vTop
  Me.fpMemo1.LoadFile strReportFile
 ' GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  'StatusBar1.Panels.Item(1).Text = GLUserName
  doAlign = False
  Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  Select Case KeyCode
    Case vbKeyEscape:
      cmdExit_Click
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

