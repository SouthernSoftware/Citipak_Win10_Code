VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxDeletePayment 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Payment Deletion"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxDeletePayment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpListRPay 
      Height          =   1950
      Left            =   720
      TabIndex        =   6
      Top             =   2400
      Width           =   10095
      _Version        =   196608
      _ExtentX        =   17806
      _ExtentY        =   3440
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      Columns         =   5
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   1
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
      ColDesigner     =   "frmVATaxDeletePayment.frx":08CA
   End
   Begin LpLib.fpList fpListPPay 
      Height          =   1950
      Left            =   735
      TabIndex        =   7
      Top             =   5205
      Width           =   10095
      _Version        =   196608
      _ExtentX        =   17806
      _ExtentY        =   3440
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      Columns         =   5
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   1
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
      ColDesigner     =   "frmVATaxDeletePayment.frx":0C6A
   End
   Begin EditLib.fpText fptxtOperator 
      Height          =   372
      Left            =   4320
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2892
      _Version        =   196608
      _ExtentX        =   5106
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
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
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   540
      Left            =   3696
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7680
      Width           =   2052
      _Version        =   131072
      _ExtentX        =   3619
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmVATaxDeletePayment.frx":100A
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDelete 
      Height          =   540
      Left            =   6336
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7680
      Width           =   2052
      _Version        =   131072
      _ExtentX        =   3619
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmVATaxDeletePayment.frx":11E8
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClearAll 
      Height          =   540
      Left            =   708
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7680
      Width           =   2424
      _Version        =   131072
      _ExtentX        =   4276
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmVATaxDeletePayment.frx":13C6
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSelectAll 
      Height          =   540
      Left            =   8868
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7680
      Width           =   2064
      _Version        =   131072
      _ExtentX        =   3641
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmVATaxDeletePayment.frx":15AD
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Transactions"
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
      Height          =   372
      Left            =   600
      TabIndex        =   9
      Top             =   4680
      Width           =   3132
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Real Transactions"
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
      Height          =   372
      Left            =   600
      TabIndex        =   8
      Top             =   1880
      Width           =   2172
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2292
      Left            =   600
      Top             =   5040
      Width           =   10452
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1488
      Top             =   432
      Width           =   8652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Payment List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   3132
      TabIndex        =   0
      Top             =   588
      Width           =   5292
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2292
      Left            =   588
      Top             =   2232
      Width           =   10452
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1488
      Top             =   312
      Width           =   8652
   End
End
Attribute VB_Name = "frmVATaxDeletePayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class
  Dim OpPayRecs() As Integer
  Dim RDeleteThese() As Integer
  Dim RDeleteAccts() As Long
  Dim rdcnt As Integer
  Dim PDeleteThese() As Integer
  Dim PDeleteAccts() As Long
  Dim pdcnt As Integer
  Public ThisBillType$

Private Sub cmdClearAll_Click()
  If fpListRPay.SelCount > 0 Then
    fpListRPay.Action = ActionDeselectAll
  ElseIf fpListPPay.SelCount > 0 Then
    fpListPPay.Action = ActionDeselectAll
  End If
End Sub

Private Sub cmdDelete_Click()
  'OpenTempPayFile is the same as open TaxCPRFileName
  'OpenPayListFile is the same as open TaxLOPFileName
  Dim RPayRec As TaxPaymentRecType
  Dim RPayRecNew() As TaxPaymentRecType
  Dim RPayHandle As Integer
  Dim RPayHandleNew As Integer
  Dim NumOfRRecs As Integer
  Dim NumOfRRecsNew As Integer
  Dim PPayRec As TaxPaymentRecType
  Dim PPayRecNew() As TaxPaymentRecType
  Dim PPayHandle As Integer
  Dim PPayHandleNew As Integer
  Dim NumOfPRecs As Integer
  Dim NumOfPRecsNew As Integer
  Dim x As Integer, y As Integer
  Dim RListRec As RealPayListType
  Dim PListRec As PersPayListType
  Dim NumOfRListRecs As Integer
  Dim NumOfPListRecs As Integer
  Dim RListRecNew() As RealPayListType
  Dim PListRecNew() As PersPayListType
  Dim RListHandle As Integer
  Dim RListHandleNew As Integer
  Dim PListHandle As Integer
  Dim PListHandleNew As Integer
  Dim RNewCnt As Integer
  Dim PNewCnt As Integer
  Dim Operator$
  Dim RMatchCnt As Integer
  Dim PMatchCnt As Integer
  
  On Error GoTo ERRORSTUFF
  
  Operator = CStr(OperNum)
  rdcnt = 0
  ReDim RDeleteThese(1 To 1) As Integer
  ReDim RDeleteAccts(1 To 1) As Long
  fpListRPay.Col = 0
  For x = 0 To fpListRPay.ListCount - 1
    fpListRPay.Row = x
    If fpListRPay.Selected = True Then
      rdcnt = rdcnt + 1
      ReDim Preserve RDeleteThese(1 To rdcnt) As Integer
      ReDim Preserve RDeleteAccts(1 To rdcnt) As Long
      fpListRPay.ListIndex = fpListRPay.Row
      fpListRPay.Col = 4
      RDeleteThese(rdcnt) = CInt(fpListRPay.ColText) 'x
      fpListRPay.Col = 0
'      fpListRPay.ListIndex = fpListRPay.Row
      RDeleteAccts(rdcnt) = CLng(fpListRPay.ColText)
    End If
  Next x
  
  pdcnt = 0
  ReDim PDeleteThese(1 To 1) As Integer
  ReDim PDeleteAccts(1 To 1) As Long
  fpListPPay.Col = 0
  For x = 0 To fpListPPay.ListCount - 1
    fpListPPay.Row = x
    If fpListPPay.Selected = True Then
      pdcnt = pdcnt + 1
      ReDim Preserve PDeleteThese(1 To pdcnt) As Integer
      ReDim Preserve PDeleteAccts(1 To pdcnt) As Long
      fpListPPay.Col = 4
      fpListPPay.ListIndex = fpListPPay.Row '#1 changed on 6/29/06 go to on 6/30/06
      PDeleteThese(pdcnt) = fpListPPay.ColText 'x
      fpListPPay.Col = 0
'      fpListPPay.ListIndex = fpListPPay.Row '#2
      PDeleteAccts(pdcnt) = CLng(fpListPPay.ColText)
    End If
  Next x
  
  If rdcnt = 0 And pdcnt = 0 Then
    frmVATaxMsg.Label1.Caption = "No payments have been selected. Deletion attempt aborted."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Close
    Exit Sub
  End If
  
  If fpListRPay.SelCount > 0 Then
    OpenRealPayListFile RListHandle, OperNum
    NumOfRListRecs = LOF(RListHandle) / Len(RListRec)
  Else
    OpenPersPayListFile PListHandle, OperNum
    NumOfPListRecs = LOF(PListHandle) / Len(PListRec)
  End If
  If NumOfRListRecs = 0 And NumOfPListRecs = 0 Then
    frmVATaxMsg.Label1.Caption = "No bills have been tagged for payment. Delete attempt aborted."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Close
    Exit Sub
  End If
  
  If TaxMsgWOpts(900, "Are you sure you want to delete this transaction? Press F10 to delete. Otherwise, press ESC to abort.", "F10 Delete", "ESC Abort") = "abort" Then
    Close
    Exit Sub
  End If
  
  RNewCnt = 0
  PNewCnt = 0
  If fpListRPay.SelCount > 0 Then
    For x = 1 To NumOfRListRecs
      Get RListHandle, x, RListRec
      For y = 1 To rdcnt
        If RListRec.CustRec = RDeleteAccts(y) Then
          RListRec.PrevListRec = -1
          Put RListHandle, x, RListRec
'          Exit For
        End If
      Next y
'      If y > rdcnt Then
'        RNewCnt = RNewCnt + 1
'        ReDim Preserve RListRecNew(1 To RNewCnt) As RealPayListType
'        RListRecNew(RNewCnt) = RListRec
'      End If
    Next x
    Close RListHandle
  ElseIf fpListPPay.SelCount > 0 Then
    For x = 1 To NumOfPListRecs
      Get PListHandle, x, PListRec
      PListRec.PrevListRec = PListRec.PrevListRec
      For y = 1 To pdcnt
        If PListRec.CustRec = PDeleteAccts(y) Then
          PListRec.PrevListRec = -1
          Put PListHandle, x, PListRec
        End If
      Next y
    Next x
    Close PListHandle
  End If
  
RNoMore1:
  OpenTempRealPayFile RPayHandle, OperNum
  NumOfRRecs = LOF(RPayHandle) / Len(RPayRec)
  ReDim RPayRecNew(1 To 1) As TaxPaymentRecType
  RNewCnt = 0
  For x = 1 To NumOfRRecs
    Get RPayHandle, x, RPayRec
    For y = 1 To rdcnt
      If RDeleteThese(y) = x Then ' - 1 Then
        RPayRec.LastPayRec = 0
        Put RPayHandle, x, RPayRec
        MainLog ("Payment for real acct # " + CStr(RPayRec.CustAcct) + " for " + QPTrim$(Using$("$###,##0.00", RPayRec.TotPaid)) + " was deleted successfully.")
'        Exit For
      End If
    Next y
'    If y > rdcnt Then
'      RNewCnt = RNewCnt + 1
'      ReDim Preserve RPayRecNew(1 To RNewCnt) As TaxPaymentRecType
'      RPayRecNew(RNewCnt) = RPayRec
'    End If
  Next x
  Close RPayHandle
  
  RMatchCnt = 0
  If NumOfRRecs = 1 Then
    KillFile "TAXRCPR" + Operator$ + ".DAT"
    KillFile "TAXRLOP" + Operator$ + ".DAT"
  Else
    OpenTempRealPayFile RPayHandle, OperNum
    NumOfRRecs = LOF(RPayHandle) / Len(RPayRec)
    For x = 1 To NumOfRRecs
      Get RPayHandle, x, RPayRec
      If RPayRec.LastPayRec = 0 Then
        RMatchCnt = RMatchCnt + 1
      End If
    Next x
    Close
    If RMatchCnt = NumOfRRecs Then
      KillFile "TAXRLOP" + Operator$ + ".DAT"
      KillFile "TAXRCPR" + Operator$ + ".DAT"
    End If
  End If
  
'  KillFile "TAXRCPR" + Operator$ + ".DAT"
'  If RNewCnt = 0 Then GoTo RNoMore2:
  
'  OpenTempRealPayFile RPayHandleNew, OperNum
'
'  For x = 1 To RNewCnt
'    RPayRec = RPayRecNew(x)
'    RPayRec.LastPayRec = x
'    Put RPayHandleNew, x, RPayRec
'  Next x
'
'  Close RPayHandleNew
RNoMore2:
  If rdcnt > 1 Then
    frmVATaxMsg.Label1.Caption = CStr(rdcnt) + " real payments have been deleted successfully."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
'    MainLog (CStr(rdcnt) + " real payments deleted successfully.")
  ElseIf rdcnt = 1 Then
    frmVATaxMsg.Label1.Caption = CStr(rdcnt) + " real payment has been deleted successfully."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
  End If
  
  fpListRPay.Action = ActionClear
  
PNoMore1:
  OpenTempPersPayFile PPayHandle, OperNum
  NumOfPRecs = LOF(PPayHandle) / Len(PPayRec)
  ReDim PPayRecNew(1 To 1) As TaxPaymentRecType
  PNewCnt = 0
  For x = 1 To NumOfPRecs
    Get PPayHandle, x, PPayRec
    For y = 1 To pdcnt
      If PDeleteThese(y) = x Then ' - 1 Then
        PPayRec.LastPayRec = 0
        Put PPayHandle, x, PPayRec
        MainLog ("Payment for personal acct # " + CStr(PPayRec.CustAcct) + " for " + QPTrim$(Using$("$###,##0.00", PPayRec.TotPaid)) + " was deleted successfully.")
      End If
    Next y
  Next x
  Close PPayHandle
  
  PMatchCnt = 0
  If NumOfPRecs = 1 Then
    KillFile "TAXPCPR" + Operator$ + ".DAT"
    KillFile "TAXPLOP" + Operator$ + ".DAT"
  Else
    OpenTempPersPayFile PPayHandle, OperNum
    NumOfPRecs = LOF(PPayHandle) / Len(PPayRec)
    For x = 1 To NumOfPRecs
      Get PPayHandle, x, PPayRec
'      PPayRec.CustAcct = PPayRec.CustAcct
      If PPayRec.LastPayRec = 0 Then
        PMatchCnt = PMatchCnt + 1
      End If
    Next x
    Close
    If PMatchCnt = NumOfPRecs Then
      KillFile "TAXPLOP" + Operator$ + ".DAT"
      KillFile "TAXPCPR" + Operator$ + ".DAT"
    End If
  End If

'  KillFile "TAXPCPR" + Operator$ + ".DAT"
'  If PNewCnt = 0 Then GoTo PNoMore2:
'
'  OpenTempPersPayFile PPayHandleNew, OperNum
'
'  For x = 1 To PNewCnt
'    PPayRec = PPayRecNew(x)
'    PPayRec.LastPayRec = x
'    Put PPayHandleNew, x, PPayRec
'  Next x
'  Close PPayHandleNew

PNoMore2:
  If pdcnt > 1 Then
    frmVATaxMsg.Label1.Caption = CStr(pdcnt) + " personal payments have been deleted successfully."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
'    MainLog (CStr(pdcnt) + " personal payments deleted successfully.")
  ElseIf pdcnt = 1 Then
    frmVATaxMsg.Label1.Caption = CStr(pdcnt) + " personal payment has been deleted successfully."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
'    MainLog ("Personal payment deleted successfully.")
  End If
  
  fpListPPay.Action = ActionClear
  '---------------------------------------------------------------
  Call ReloadMe
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxDeletePayment", "cmdDelete_Click", Erl)
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

Private Sub cmdExit_Click()
  Unload Me
  DoEvents
  frmVATaxPayMenu.Show
End Sub

Private Sub cmdSelectAll_Click()
  frmVATaxBillPostOpt.Show vbModal
  If frmVATaxBillPostOpt.fptxtPostType.Text = "Real" Then
    Unload frmVATaxBillPostOpt
    fpListRPay.Action = ActionSelectAll
    fpListPPay.Action = ActionDeselectAll
  ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Personal" Then
    Unload frmVATaxBillPostOpt
    fpListPPay.Action = ActionSelectAll
    fpListRPay.Action = ActionDeselectAll
  ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Exit" Then
    DoEvents
    Unload frmVATaxBillPostOpt
    Exit Sub
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
    Case vbKeyF10:
      SendKeys "%D"
      Call cmdDelete_Click
      KeyCode = 0
    Case vbKeyF3:
      SendKeys "%C"
      Call cmdClearAll_Click
      KeyCode = 0
    Case vbKeyF7:
      SendKeys "%S"
      Call cmdSelectAll_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpDeleteTax
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxDeletePayment.")
      Call Terminate
      End
    End If
  End If
End Sub

'Private Sub Form_Resize()
'  If Me.WindowState <> vbMinimized Then
'    Me.Visible = False
'    'Temp_Class.ResizeControls Me
'    Me.Visible = True
'    Me.SetFocus
'    DoEvents
'  End If
'End Sub

Private Sub LoadMe()
  Dim PPayRec As TaxPaymentRecType
  Dim PPayHandle As Integer
  Dim RPayRec As TaxPaymentRecType
  Dim RPayHandle As Integer
  Dim NumOfPRecs As Integer
  Dim NumOfRRecs As Integer
  Dim x As Integer
  Dim PListRec As PersPayListType
  Dim RListRec As RealPayListType
  Dim ListHandle As Integer
  Dim OHandle As Integer
  Dim NumOfOperRecs As Integer
  Dim OperRec As CitiPassType
  Dim NumOfOperPayRecs As Integer
  
  fptxtOperator.Text = "Operator # " + CStr(OperNum)
  If Exist("TAXRCPR" + CStr(OperNum) + ".DAT") Then
    OpenTempRealPayFile RPayHandle, OperNum
    NumOfRRecs = LOF(RPayHandle) / Len(RPayRec)
    If NumOfRRecs = 0 Then
      frmVATaxMsg.Label1.Caption = "There are no real payment records saved for operator #" + CStr(OperNum) + "."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      Close
      Exit Sub
    Else
      For x = 1 To NumOfRRecs
        Get RPayHandle, x, RPayRec
        If RPayRec.LastPayRec = 0 Then GoTo MoveOnR
        fpListRPay.InsertRow = CStr(RPayRec.CustAcct) + Chr(9) + QPTrim$(RPayRec.CustName) + Chr(9) + QPTrim$(Using$("$###,##0.00", RPayRec.AmtPaid)) + Chr(9) + QPTrim$(Using$("$###,##0.00", RPayRec.AmtOwed)) + Chr(9) + CStr(x)
        DoEvents
MoveOnR:
      Next x
    End If
  End If
  
  If Exist("TAXPCPR" + CStr(OperNum) + ".DAT") Then
    OpenTempPersPayFile PPayHandle, OperNum
    NumOfPRecs = LOF(PPayHandle) / Len(PPayRec)
    If NumOfPRecs = 0 Then
      frmVATaxMsg.Label1.Caption = "There are no personal payment records saved for operator #" + CStr(OperNum) + "."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      Close
      Exit Sub
    Else
      For x = 1 To NumOfPRecs
        Get PPayHandle, x, PPayRec
        If PPayRec.LastPayRec = 0 Then GoTo MoveOnP
        fpListPPay.InsertRow = CStr(PPayRec.CustAcct) + Chr(9) + QPTrim$(PPayRec.CustName) + Chr(9) + QPTrim$(Using$("$###,##0.00", PPayRec.AmtPaid)) + Chr(9) + QPTrim$(Using$("$###,##0.00", PPayRec.AmtOwed)) + Chr(9) + CStr(x)
        DoEvents
MoveOnP:
      Next x
    End If
  End If
  
  Close
  
'  If NumOfRRecs > 0 Then
'    fpListRPay.Row = 0
'    fpListRPay.ListIndex = 0
'    fpListRPay.Selected = True
'  ElseIf NumOfPRecs > 0 Then
'    fpListPPay.Row = 0
'    fpListPPay.ListIndex = 0
'    fpListRPay.Selected = True
'  End If
  
End Sub

Private Sub ReloadMe()
  Dim PPayRec As TaxPaymentRecType
  Dim PPayHandle As Integer
  Dim RPayRec As TaxPaymentRecType
  Dim RPayHandle As Integer
  Dim NumOfPRecs As Integer
  Dim NumOfRRecs As Integer
  Dim x As Integer
  Dim PListRec As PersPayListType
  Dim RListRec As RealPayListType
  Dim ListHandle As Integer
  Dim OHandle As Integer
  Dim NumOfOperRecs As Integer
  Dim OperRec As CitiPassType
  Dim NumOfOperPayRecs As Integer
  
  fptxtOperator.Text = "Operator # " + CStr(OperNum)
  If Exist("TAXRCPR" + CStr(OperNum) + ".DAT") Then
    OpenTempRealPayFile RPayHandle, OperNum
    NumOfRRecs = LOF(RPayHandle) / Len(RPayRec)
    If NumOfRRecs > 0 Then
      For x = 1 To NumOfRRecs
        Get RPayHandle, x, RPayRec
        If RPayRec.LastPayRec = 0 Then GoTo MoveOnR
        fpListRPay.InsertRow = CStr(RPayRec.CustAcct) + Chr(9) + QPTrim$(RPayRec.CustName) + Chr(9) + QPTrim$(Using$("$###,##0.00", RPayRec.AmtPaid)) + Chr(9) + QPTrim$(Using$("$###,##0.00", RPayRec.AmtOwed)) + Chr(9) + CStr(x)
        DoEvents
MoveOnR:
      Next x
    End If
  End If
  
  If Exist("TAXPCPR" + CStr(OperNum) + ".DAT") Then
    OpenTempPersPayFile PPayHandle, OperNum
    NumOfPRecs = LOF(PPayHandle) / Len(PPayRec)
    If NumOfPRecs > 0 Then
      For x = 1 To NumOfPRecs
        Get PPayHandle, x, PPayRec
        If PPayRec.LastPayRec = 0 Then GoTo MoveOnP
        fpListPPay.InsertRow = CStr(PPayRec.CustAcct) + Chr(9) + QPTrim$(PPayRec.CustName) + Chr(9) + QPTrim$(Using$("$###,##0.00", PPayRec.AmtPaid)) + Chr(9) + QPTrim$(Using$("$###,##0.00", PPayRec.AmtOwed)) + Chr(9) + CStr(x)
MoveOnP:
      Next x
    End If
  End If
  
  Close
  
'  OpenPersPayListFile ListHandle, OperNum
'  NumOfPRecs = LOF(ListHandle) / Len(PListRec)
'  Close
  
  fpListRPay.Refresh
  fpListPPay.Refresh
  If NumOfRRecs > 0 Then
    fpListRPay.ListIndex = -1
  ElseIf NumOfPRecs > 0 Then
    fpListPPay.ListIndex = -1
  End If

End Sub

Private Sub fpListPPay_Click()
  If fpListPPay.SelCount > 0 Then
    fpListRPay.Action = ActionDeselectAll
  End If
End Sub

Private Sub fpListRPay_Click()
  If fpListRPay.SelCount > 0 Then
    fpListPPay.Action = ActionDeselectAll
  End If

End Sub
