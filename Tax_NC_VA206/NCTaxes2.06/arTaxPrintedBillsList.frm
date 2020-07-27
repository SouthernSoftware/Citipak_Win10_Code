VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmTaxPrintedBillsList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printed Tax Bills List"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8910
   Icon            =   "arTaxPrintedBillsList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   480
      Left            =   1800
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   "Press this button to exit back to the main business license menu."
      Top             =   6120
      Width           =   2160
      _Version        =   131072
      _ExtentX        =   3810
      _ExtentY        =   847
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
      BackStyle       =   0
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
      ButtonDesigner  =   "arTaxPrintedBillsList.frx":08CA
   End
   Begin LpLib.fpList fpList1 
      Height          =   3180
      Left            =   735
      TabIndex        =   2
      Top             =   2220
      Width           =   7455
      _Version        =   196608
      _ExtentX        =   13150
      _ExtentY        =   5609
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Columns         =   3
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
      ColumnHeaderShow=   0   'False
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
      ColDesigner     =   "arTaxPrintedBillsList.frx":0AE1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H008F8265&
      Caption         =   "Reprint Numbers"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1290
      Left            =   1290
      TabIndex        =   4
      Top             =   240
      Width           =   6330
      Begin VB.OptionButton optFirst 
         BackColor       =   &H008F8265&
         Caption         =   "First Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   252
         Left            =   1080
         TabIndex        =   0
         ToolTipText     =   "Press F3 to bring up assistance for this field."
         Top             =   360
         Width           =   1890
      End
      Begin VB.OptionButton optSecond 
         BackColor       =   &H008F8265&
         Caption         =   "Last Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   252
         Left            =   3480
         TabIndex        =   1
         Top             =   360
         Width           =   1785
      End
      Begin EditLib.fpText fptxtFirst 
         Height          =   420
         Left            =   1005
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Click on the 'First Number' button then make a selection in the list. That bill number will appear in the First Number box."
         Top             =   645
         Width           =   1890
         _Version        =   196608
         _ExtentX        =   3334
         _ExtentY        =   741
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   10.5
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
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
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
      Begin EditLib.fpText fptxtSecond 
         Height          =   420
         Left            =   3435
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Click on the 'Last Number' button then make a selection in the list. That bill number will appear in the Last Number box."
         Top             =   645
         Width           =   1890
         _Version        =   196608
         _ExtentX        =   3334
         _ExtentY        =   741
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   10.5
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
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
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
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdApply 
      Height          =   480
      Left            =   4935
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6120
      Width           =   2160
      _Version        =   131072
      _ExtentX        =   3810
      _ExtentY        =   847
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
      ButtonDesigner  =   "arTaxPrintedBillsList.frx":0E61
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Amt Owed"
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
      Height          =   375
      Left            =   6360
      TabIndex        =   10
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cust Name"
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
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bill #"
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
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   1800
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   3855
      Left            =   495
      Top             =   1740
      Width           =   7920
   End
End
Attribute VB_Name = "frmTaxPrintedBillsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
  If QPTrim$(fptxtFirst.Text) = "" And QPTrim$(fptxtSecond.Text) = "" Then
    Call TaxMsg(900, "No selection has been made. Nothing to apply.")
    Exit Sub
  End If
  
  If CDbl(fptxtFirst.Text) > CDbl(fptxtSecond.Text) Then
    Call TaxMsg(900, "Please make sure the value of the first number is less than the value of the last number.")
    Exit Sub
  End If
  
  frmTaxBillReprinting.fpDblSnglFirstBill = CDbl(fptxtFirst.Text)
  frmTaxBillReprinting.fpDblSnglLastBill = CDbl(fptxtSecond.Text)
  Unload Me
  DoEvents
End Sub

Private Sub cmdExit_Click()
  Unload Me
  DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyEscape:
    SendKeys "%C"
    Call cmdExit_Click
    KeyCode = 0
  Case vbKeyF10:
    SendKeys "%A"
    Call cmdApply_Click
    KeyCode = 0
  Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Dim TBRec As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long
  Dim ZipRec As BillPrintZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  Dim WhatRec&
  
  NumOfMRRecs = 0
  NumOfZRecs = 0
  OpenTaxBillFile TBHandle, NumOfTBRecs
  If Exist("MORTIDX.DAT") Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf Exist("ZIPIDX.DAT") Then '12/6/06
    OpenZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  For x = 1 To NumOfTBRecs
    If NumOfMRRecs > 0 Then '12/6/06
      Get MRHandle, x, MortRec
      WhatRec& = MortRec.TaxBillRec
    ElseIf NumOfZRecs > 0 Then '12/6/06
      Get ZHandle, x, ZipRec
      WhatRec& = ZipRec.TaxBillRec
    Else
      WhatRec& = x '12/6/06
    End If
    Get TBHandle, WhatRec&, TBRec
    If TBRec.TotalBillDue > 0 Then
      If TBRec.BillNumber > 0 Then
        fpList1.InsertRow = Using$("########0", TBRec.BillNumber) + Chr(9) + QPTrim$(TBRec.CustName) + Chr(9) + Using$("$###,###,##0.00", TBRec.TotalBillDue - TBRec.OverPayAmt)
      End If
    End If
  Next x
  Close TBHandle
  Close ZHandle
  Close MRHandle
  
End Sub

Private Sub fpList1_Click()
'  If Exist("MORTIDX.DAT") Then
'    fpList1.Col = 0
'  ElseIf Exist("ZIPIDX.DAT") Then
'    fpList1.Col = 0
'  Else
'    fpList1.Col = 1
'  End If
  
  If optFirst.Value = True Then
    fptxtFirst.Text = QPTrim$(fpList1.ColText)
  ElseIf optSecond.Value = True Then
    fptxtSecond.Text = QPTrim$(fpList1.ColText)
  End If
End Sub

Private Sub optFirst_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
    optSecond.SetFocus
  End If
End Sub

Private Sub optSecond_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
    optFirst.SetFocus
  End If
End Sub
