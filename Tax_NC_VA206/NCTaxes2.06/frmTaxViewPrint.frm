VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{A18D4668-91EF-101C-84A6-BA990A365A4E}#3.0#0"; "mem32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTaxViewPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ViewPrint"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10485
   Icon            =   "frmTaxViewPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   96
      Top             =   7680
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      DialogTitle     =   "Print"
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
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
            TextSave        =   "2:34 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   6112
            TextSave        =   "8/29/2018"
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
      Left            =   18
      TabIndex        =   1
      Top             =   0
      Width           =   10452
      _Version        =   196608
      _ExtentX        =   18436
      _ExtentY        =   13144
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
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
      AutoAdvance     =   0   'False
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
   Begin fpBtnAtlLibCtl.fpBtn cmdAlignment 
      Height          =   396
      Left            =   2366
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7632
      Width           =   1428
      _Version        =   131072
      _ExtentX        =   2519
      _ExtentY        =   698
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
      ButtonDesigner  =   "frmTaxViewPrint.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrnScn 
      Height          =   396
      Left            =   4254
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7632
      Width           =   1980
      _Version        =   131072
      _ExtentX        =   3492
      _ExtentY        =   698
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
      ButtonDesigner  =   "frmTaxViewPrint.frx":0AA5
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrint 
      Height          =   396
      Left            =   6694
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7632
      Width           =   1428
      _Version        =   131072
      _ExtentX        =   2519
      _ExtentY        =   698
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
      ButtonDesigner  =   "frmTaxViewPrint.frx":0C87
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   396
      Left            =   8582
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7632
      Width           =   1428
      _Version        =   131072
      _ExtentX        =   2519
      _ExtentY        =   698
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
      ButtonDesigner  =   "frmTaxViewPrint.frx":0E62
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   396
      Left            =   478
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7632
      Width           =   1428
      _Version        =   131072
      _ExtentX        =   2519
      _ExtentY        =   698
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
      ButtonDesigner  =   "frmTaxViewPrint.frx":103D
   End
End
Attribute VB_Name = "frmTaxViewPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim strReportFile As String
Public PgNum As Integer
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim vWidth%, vHeight%, vTop%, vLeft%
'''Public Declare Function PrinterProperties Lib "winspool.drv" (ByVal hwnd As Long, ByVal hPrinter As Long) As Long

Private Sub cmdAlignment_Click()
  frmTaxPrint.Show 1
End Sub

Private Sub cmdExit_Click()
  Unload frmTaxViewPrint
  DoEvents
End Sub
Private Sub cmdPrint_Click()
  frmTaxPrint.Show 1
  DoEvents
  Unload frmTaxViewPrint
  DoEvents
End Sub
Public Sub PrintWSet(DefPrinter As String, Copies As Integer)
  Dim RptHandle As Integer, LPTHandle As Integer
  Dim RptA As Integer, LPTA As Integer, ToPrintA As String
  Dim ToPrint As String, CopyLoop As Integer
  'on error goto Cancel
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
      If frmTaxPrint.cmdCancel = False Then
        Line Input #RptHandle, ToPrint$
        ToPrint$ = RTrim$(ToPrint$)
        Print #LPTHandle, ToPrint$
      Else
        Exit Do
        Printer.EndDoc
      End If
    Loop Until eof(RptHandle)
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
    Do Until eof(RptA)
      If frmTaxPrint.cmdCancel = False Then
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
    Unload frmTaxViewPrint
    DoEvents
  End If
End Sub


