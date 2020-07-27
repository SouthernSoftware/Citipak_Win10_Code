VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmTaxAdjustBillList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Bill List"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9900
   Icon            =   "frmTaxAdjustBillList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   9900
   StartUpPosition =   1  'CenterOwner
   Begin LpLib.fpList fpList1 
      Height          =   1770
      Left            =   900
      TabIndex        =   0
      Top             =   1440
      Width           =   8115
      _Version        =   196608
      _ExtentX        =   14314
      _ExtentY        =   3122
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
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
      Columns         =   8
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   0
      WrapList        =   0   'False
      WrapWidth       =   0
      SelMax          =   -1
      AutoSearch      =   2
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
      ReadOnly        =   0   'False
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
      ColDesigner     =   "frmTaxAdjustBillList.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClose 
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1575
      _Version        =   131072
      _ExtentX        =   2778
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmTaxAdjustBillList.frx":0D2D
   End
   Begin EditLib.fpText fptxtName 
      Height          =   375
      Left            =   2903
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   900
      Width           =   4095
      _Version        =   196608
      _ExtentX        =   7223
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
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
      ControlType     =   1
      Text            =   ""
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
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "*Property No Longer Owned By This Customer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   336
      Left            =   552
      TabIndex        =   4
      Top             =   3264
      Visible         =   0   'False
      Width           =   6252
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   510
      Left            =   2948
      Top             =   285
      Width           =   4050
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Bill List"
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
      Height          =   330
      Left            =   3203
      TabIndex        =   2
      Top             =   360
      Width           =   3510
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   4155
      Left            =   150
      Top             =   120
      Width           =   9615
   End
End
Attribute VB_Name = "frmTaxAdjustBillList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim Balance#

Private Sub cmdClose_Click()
  frmTaxAdjustments.BillRec = -1
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    Call fpList1_DblClick
    KeyCode = 0
  End If
  
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdClose_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Dim TaxRec As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim LastTrans As Long
'  Dim Balance#
  Dim TransRecord&
  Dim BillTCnt As Integer
  Dim BillNum As Double
  Dim CustPin As Long
  Dim TestBillNum As String
  On Error GoTo ERRORSTUFF
  
  CustPin = frmTaxAdjustments.fpLongAcctNum.Value
  
  OpenTaxCustFile CHandle, NumOfCRecs
  Get CHandle, GCustNum, TaxRec
  Close CHandle
  TransRecord& = TaxRec.LastTrans
  
  fptxtName.Text = QPTrim$(frmTaxAdjustments.fptxtName)
  OpenTaxTransFile THandle, NumOfTRecs
  
  Do While TransRecord& <> 0
    Get THandle, TransRecord&, TaxTrans
    If TaxTrans.TranType = 1 Then
      Balance# = OldRound#(TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      Balance# = OldRound#(Balance# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      Balance# = OldRound#(Balance# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      Balance# = OldRound#(Balance# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      Balance# = OldRound#(Balance# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt)) 'added 8/9/06 per Bob
      TestBillNum$ = ParseBillNum(TaxTrans.Description)
      If QPTrim$(TestBillNum$) <> "" Then
        BillNum# = CLng(ParseBillNum(TaxTrans.Description))
      Else
        BillNum# = 0
      End If
      If IsCurrentOwner(TaxTrans.RealPin, CustPin) = True Then
        fpList1.InsertRow = MakeRegDate(TaxTrans.TransDate) + Chr(9) + Using$("####", TaxTrans.TaxYear) + Chr(9) + Using("#########0", BillNum#) + Chr(9) + Using$("$#,###,##0.00", TaxTrans.Amount) + Chr(9) + Using$("$#,###,##0.00", Balance#) + Chr(9) + Using("#####0", TransRecord) + Chr(9) + QPTrim$(TaxTrans.RealPin) + Chr(9) + QPTrim$(TaxTrans.PersPin)
      Else
        Label2.Visible = True
        fpList1.InsertRow = MakeRegDate(TaxTrans.TransDate) + Chr(9) + Using$("####", TaxTrans.TaxYear) + Chr(9) + Using("#########0", BillNum#) + Chr(9) + Using$("$#,###,##0.00", TaxTrans.Amount) + Chr(9) + Using$("$#,###,##0.00", Balance#) + "*" + Chr(9) + Using("#####0", TransRecord) + Chr(9) + QPTrim$(TaxTrans.RealPin) + Chr(9) + QPTrim$(TaxTrans.PersPin)
      End If
    End If
Again:
    TransRecord& = TaxTrans.LastTrans
  Loop
  Close THandle
  
  fpList1.ListIndex = 0
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxAdjustBillList", "FormLoad", Erl)
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

End Sub

Private Sub fpList1_DblClick()
  Dim ThisBill As Long, xxamt As Double
  If fpList1.ListCount > 0 Then
    fpList1.Row = fpList1.ListIndex
    fpList1.col = 5
    frmTaxAdjustments.BillRec = CDbl(fpList1.ColText)
    fpList1.col = 2
    ThisBill = CLng(fpList1.ColText)
    frmTaxAdjustments.ThisBillNum = ThisBill
    frmTaxAdjustments.fpLngIntBill = ThisBill
    fpList1.col = 4
    If Mid$(fpList1.ColText, Len(fpList1.ColText), 1) = "*" Then
      xxamt = Mid$(fpList1.ColText, 1, Len(fpList1.ColText) - 1)
    Else
    xxamt = CDbl(fpList1.ColText)
    End If
    frmTaxAdjustments.ThisBillBal = xxamt
    fpList1.col = 6
    frmTaxAdjustments.RealPin = fpList1.ColText
    fpList1.col = 7
    frmTaxAdjustments.PersPin = fpList1.ColText
    Unload Me
  End If
End Sub

Private Sub fptxtName_Change()
  fptxtName.Text = QPTrim$(frmTaxAdjustments.fptxtName)
End Sub
