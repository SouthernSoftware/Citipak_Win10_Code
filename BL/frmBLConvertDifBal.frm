VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "EDT32X30.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmBLConvertDifBal 
   BackColor       =   &H000000C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8508
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10488
   Icon            =   "frmBLConvertDifBal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8508
   ScaleWidth      =   10488
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin LpLib.fpList fpList 
      Height          =   2748
      Left            =   288
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4416
      Width           =   9900
      _Version        =   196608
      _ExtentX        =   17462
      _ExtentY        =   4847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Columns         =   5
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   0
      WrapList        =   0   'False
      WrapWidth       =   0
      SelMax          =   -1
      AutoSearch      =   1
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483627
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ScrollHScale    =   2
      ScrollHInc      =   0
      ColsFrozen      =   0
      ScrollBarV      =   1
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   0
      AllowColDragDrop=   0
      ReadOnly        =   -1  'True
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   2
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   1
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   -2147483637
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   -1  'True
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   -2147483637
      ExtendRow       =   0
      DataField       =   ""
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      ColDesigner     =   "frmBLConvertDifBal.frx":08CA
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "ESC EXIT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1494
      TabIndex        =   4
      Top             =   7584
      Width           =   2028
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "F10 CONTINUE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   6966
      TabIndex        =   3
      Top             =   7584
      Width           =   2028
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5 PRINT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4230
      TabIndex        =   1
      Top             =   7584
      Width           =   2028
   End
   Begin EditLib.fpText fptxtChoice 
      Height          =   300
      Left            =   9264
      TabIndex        =   2
      Top             =   2112
      Width           =   972
      _Version        =   196608
      _ExtentX        =   1714
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   0
      CaretOverWrite  =   3
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
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "fpText1"
      CharValidationText=   ""
      MaxLength       =   255
      MultiLine       =   0   'False
      PasswordChar    =   ""
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
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE: ABNORMAL ACCOUNT BALANCES WILL NOT APPEAR ON THIS REPORT."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   636
      Left            =   1776
      TabIndex        =   7
      Top             =   3456
      Width           =   6924
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmBLConvertDifBal.frx":0CB5
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
      Height          =   1500
      Left            =   2316
      TabIndex        =   6
      Top             =   1776
      Width           =   5868
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NOTICE: ACCOUNT BALANCE DISCREPENCIES DETECTED"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   732
      Left            =   3084
      TabIndex        =   5
      Top             =   672
      Width           =   4428
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   8172
      Left            =   144
      Top             =   192
      Width           =   10188
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   1212
      Left            =   2868
      Top             =   420
      Width           =   4764
   End
End
Attribute VB_Name = "frmBLConvertDifBal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdContinue_Click()
  fptxtChoice.Text = "continue"
  Me.Hide
End Sub

Private Sub cmdExit_Click()
  fptxtChoice.Text = "exit"
  Me.Hide
End Sub

Private Sub PrintText()
  Dim x As Integer
  Dim RptFile$
  Dim RptHandle As Integer
  Dim DosCustRec As DosARCustRecType
  Dim DosCustHandle As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim FF$
  Dim PntCnt As Integer
  
  On Error Resume Next
  FF$ = Chr$(12)
  MaxLines = 57
  LineCnt = 0
  RptFile = "DifBal.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  GoSub PrintHeader
  OpenDosCustFile DosCustHandle
  For x = 1 To NumOfDifs
    Get DosCustHandle, DifNums(x), DosCustRec
    If LineCnt > MaxLines Then GoSub PrintHeader
    Print #RptHandle, QPTrim$(DosCustRec.CustName); Tab(37); Using$("##########", DosCustRec.CUSTNUMB); Tab(48); Using$("$#,##0.00", DosCustRec.PenBal); Tab(58); Using$("$#,##0.00", DosCustRec.LicBal); Tab(69); Using$("$###,##0.00", DosCustRec.AcctBal)
    PntCnt = PntCnt + 1
    LineCnt = LineCnt + 1
SkipIt:
  Next x
  Print #RptHandle, FF$
  Close
  
  ViewPrint RptFile, "Account Balance Report", True

  KillFile "DifBal.PRN"
  Exit Sub
  
PrintHeader:
  Print #RptHandle, FF$
  Print #RptHandle, "Discrepant Balances Report"
  Print #RptHandle, Date
  Print #RptHandle,
  Print #RptHandle, Tab(3); "Business Name"; Tab(39); "Acct Num"; Tab(50); "Pen Bal"; Tab(60); "Lic Bal"; Tab(72); "Acct Bal"
  Print #RptHandle, String$(80, "=")
  LineCnt = 5
Return
End Sub

Private Sub cmdPrint_Click()
  Dim PrintType$
  
  frmBLReportOpt.Show vbModal 'opens small screen from which the
  'user selects the printing method
  PrintType$ = frmBLReportOpt.fptxtPrintType
  Select Case PrintType$
    Case "Graphical"
      Call PrintGraphics
    Case "Text"
      MsgBox "Pitch 10 is recommended for this report."
      Call PrintText
    Case "Exit"
  End Select
  Unload frmBLReportOpt

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      Call cmdContinue_Click
      KeyCode = 0
    Case vbKeyF5:
      Call cmdPrint_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Call LoadMe
End Sub

Private Sub LoadMe()
  Dim x As Integer
  Dim DosNumOfCustRecs As Integer
  Dim DosCustRec As DosARCustRecType
  Dim DosCustHandle As Integer
  
  DoEvents
  fptxtChoice.Visible = False
  OpenDosCustFile DosCustHandle
  For x = 1 To NumOfDifs
    Get DosCustHandle, DifNums(x), DosCustRec
    fpList.AddItem QPTrim$(DosCustRec.CustName) + Chr(9) + QPTrim$(DosCustRec.CUSTNUMB) + Chr(9) + Using("$#,##0.00", DosCustRec.PenBal) + Chr(9) + Using("$#,##0.00", DosCustRec.LicBal) + Chr(9) + Using("$#,##0.00", DosCustRec.AcctBal)
SkipIt:
  Next x

End Sub

Private Sub PrintGraphics()
  Dim x As Integer
  Dim RptFile$
  Dim RptHandle As Integer
  Dim DosCustRec As DosARCustRecType
  Dim DosCustHandle As Integer
  Dim dlm$
  
  On Error Resume Next
  dlm$ = "~"
  RptFile = "BLRPTS\DIFBAL.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  OpenDosCustFile DosCustHandle
  For x = 1 To NumOfDifs
    Get DosCustHandle, DifNums(x), DosCustRec
    Print #RptHandle, QPTrim$(DosCustRec.CustName); dlm; QPTrim$(DosCustRec.CUSTNUMB); dlm; CStr(DosCustRec.PenBal); dlm; CStr(DosCustRec.LicBal); dlm; CStr(DosCustRec.AcctBal)
  Next x
  Close
  arBLConvertDifBal.Show vbModal

End Sub

