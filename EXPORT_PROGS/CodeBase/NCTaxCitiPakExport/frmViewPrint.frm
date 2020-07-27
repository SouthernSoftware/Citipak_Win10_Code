VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmViewPrint 
   BackColor       =   &H00D0D0D0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Print"
   ClientHeight    =   8415
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   10485
   Icon            =   "frmViewPrint.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   96
      Top             =   7632
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      DialogTitle     =   "Print"
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   8160
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6112
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   6112
            TextSave        =   "9:44 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   6112
            TextSave        =   "2/5/2008"
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
   Begin VB.PictureBox fpMemo1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7452
      Left            =   0
      ScaleHeight     =   7455
      ScaleWidth      =   10455
      TabIndex        =   0
      Top             =   -48
      Width           =   10452
   End
   Begin VB.PictureBox cmdAlignment 
      Height          =   390
      Left            =   576
      ScaleHeight     =   330
      ScaleWidth      =   1365
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7584
      Width           =   1425
   End
   Begin VB.PictureBox cmdPrnScn 
      Height          =   396
      Left            =   2352
      ScaleHeight     =   330
      ScaleWidth      =   1920
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7584
      Width           =   1980
   End
   Begin VB.PictureBox cmdPrint 
      Height          =   390
      Left            =   4656
      ScaleHeight     =   330
      ScaleWidth      =   1365
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7584
      Width           =   1425
   End
   Begin VB.PictureBox cmdSave 
      Height          =   390
      Left            =   6624
      ScaleHeight     =   330
      ScaleWidth      =   1365
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7584
      Width           =   1425
   End
   Begin VB.PictureBox cmdExit 
      Height          =   390
      Left            =   8640
      ScaleHeight     =   330
      ScaleWidth      =   1365
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7584
      Width           =   1425
   End
End
Attribute VB_Name = "frmViewPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim strReportFile As String
Public PgNum As Integer
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim vWidth%, vHeight%, vTop%, vLeft%
'''Public Declare Function PrinterProperties Lib "winspool.drv" (ByVal hwnd As Long, ByVal hPrinter As Long) As Long

Private Sub cmdAlignment_Click()
  frmPrint.Show 1
End Sub

Private Sub cmdExit_Click()
  Unload frmViewPrint
  DoEvents
End Sub
Private Sub cmdPrint_Click()
  frmPrint.Show 1
  DoEvents
  Unload frmViewPrint
  DoEvents
End Sub
Public Sub PrintWSet(DefPrinter As String, Copies As Integer)
  Dim RptHandle As Integer, LPTHandle As Integer
  Dim RptA As Integer, LPTA As Integer, ToPrintA As String
  Dim ToPrint As String, CopyLoop As Integer
  On Error GoTo Cancel
  'Printer.Print
'''  to strReportFile DefPrinter'[ADDITIVE] | PortName]
  If doAlign = True Then
    NumOfAligns = 1
    GoSub PrintAlignMask
  End If
  
  LPTHandle = FreeFile
  For CopyLoop = 1 To Copies
    Open DefPrinter For Output As LPTHandle
    RptHandle = FreeFile
    Open strReportFile For Input As RptHandle
    Do
      If frmPrint.cmdCancel = False Then
        Line Input #RptHandle, ToPrint$
        ToPrint$ = RTrim$(ToPrint$)
        Print #LPTHandle, ToPrint$
      Else
        Exit Do
        Printer.EndDoc
      End If
    Loop Until EOF(RptHandle)
    Close LPTHandle, RptHandle
    Next CopyLoop
  Printer.EndDoc
  Exit Sub
PrintAlignMask:
    LPTA = FreeFile
    Open DefPrinter For Output As LPTA
    RptA = FreeFile
    If Exist(alnRpt) Then
      Open alnRpt For Input As RptA
    Else
      MsgBox "The mask file needed for the alignment test cannot be found."
      BadMaskFlag = True
      Close LPTA, RptA
      Exit Sub
    End If
    Do Until EOF(RptA)
      If frmPrint.cmdCancel = False Then
        Line Input #RptA, ToPrintA$
        ToPrintA$ = RTrim$(ToPrintA$)
        Print #LPTA, ToPrintA$
        If InStr(ToPrintA$, "BOTTOM OF") Then Exit Do
      
      Else
        Exit Do
        Printer.EndDoc
      End If
    Loop
    Close LPTA, RptA
    Printer.EndDoc
    If MsgBox("Do You Wish to Print Another Mask?", vbYesNo, "Print Mask") = vbYes Then
      NumOfAligns = NumOfAligns + 1
      GoSub PrintAlignMask
    End If
    
    Exit Sub
    
Cancel:
  MsgBox "Could not open " + DefPrinter + ". Printing aborted."
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
  Me.Width = vWidth
  Me.Height = vHeight
  Me.Left = vLeft
  Me.Top = vTop
  Me.fpMemo1.LoadFile strReportFile
  StatusBar1.Panels.Item(1).Text = GLUserName
End Sub

Private Sub Form_Resize()
  Temp_Class.ResizeControls Me
  DoEvents
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  Select Case KeyCode
    Case vbKeyEscape:
'      SendKeys "%X"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF8:
'      SendKeys "%P"
      Call cmdPrint_Click
      KeyCode = 0
    Case vbKeyF10:
'      SendKeys "%S"
      Call cmdSave_Click
      KeyCode = 0
    Case vbKeyF7
'      SendKeys "%t"
      Call cmdPrnScn_Click
      KeyCode = 0
    Case vbKeyF5:
'      SendKeys "%P"
      Call cmdAlignment_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    Unload frmViewPrint
    DoEvents
  End If
End Sub

