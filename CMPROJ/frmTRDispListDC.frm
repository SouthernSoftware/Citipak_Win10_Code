VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmTRDispListDC 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction List"
   ClientHeight    =   6276
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   9072
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6276
   ScaleWidth      =   9072
   StartUpPosition =   2  'CenterScreen
   Begin LpLib.fpList fpTRList 
      Height          =   3936
      Left            =   288
      TabIndex        =   0
      Top             =   1464
      Width           =   8532
      _Version        =   196608
      _ExtentX        =   15049
      _ExtentY        =   6943
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   10.8
         Charset         =   255
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
      Sorted          =   2
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   0
      WrapList        =   0   'False
      WrapWidth       =   0
      SelMax          =   1
      AutoSearch      =   1
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   0
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
      ScrollBarV      =   0
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   2
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   0   'False
      DataAutoSizeCols=   0
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   3
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
      ColDesigner     =   "frmTRDispListDC.frx":0000
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdOK 
      Height          =   480
      Left            =   6408
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5664
      Width           =   1164
      _Version        =   131072
      _ExtentX        =   2053
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmTRDispListDC.frx":03F4
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdExit 
      Height          =   480
      Left            =   7656
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5664
      Width           =   1236
      _Version        =   131072
      _ExtentX        =   2180
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmTRDispListDC.frx":05CD
   End
   Begin fpBtnAtlLibCtl.fpBtn fpcmdPrint 
      Height          =   480
      Left            =   5280
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5664
      Width           =   1068
      _Version        =   131072
      _ExtentX        =   1884
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmTRDispListDC.frx":07A8
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   5976
      TabIndex        =   8
      Top             =   168
      Width           =   2052
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Name/Acct"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Left            =   312
      TabIndex        =   6
      Top             =   144
      Width           =   5412
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   792
      TabIndex        =   5
      Top             =   600
      Width           =   2052
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Trans Date                        Description                         Trans Amt               Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   408
      TabIndex        =   2
      Top             =   1128
      Width           =   8412
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Item or Highlight and Double-Click for Details."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   516
      Left            =   432
      TabIndex        =   1
      Top             =   5544
      Width           =   4404
   End
End
Attribute VB_Name = "frmTRDispListDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BeenDone As Boolean

Private Sub fpCmdExit_Click()
  SearchRec = 0
  BeenDone = False
  Unload frmTRDispListDC
End Sub

Private Sub Form_Activate()
  SearchRec& = 0
  If Not BeenDone Then
    BeenDone = True
    Me.fpTRList.ListIndex = 0
  End If
End Sub


Private Sub fpCmdOk_Click()
  If fpTRList.SelCount > 0 Then
    Call fpTRList_DblClick
  End If
End Sub

Private Sub fpCmdPrint_Click()
  Dim ReportFile As String, DCRpt As Integer, cnt As Integer, go2line As Integer
  Dim gofrom As Integer
  ReportFile$ = UBPath$ + "DCtrlist.RPT"
  DCRpt = FreeFile
  Open ReportFile$ For Output As DCRpt
  Print #DCRpt, ""
  Print #DCRpt, Now
  Print #DCRpt, Tab(2); "Customer Transaction List"
  Print #DCRpt, Tab(2); QPTrim$(Label5.Caption); Tab(53); Label7.Caption
 ' Print #DCRpt, Tab(2); Label2.Caption; Tab(24); Label3.Caption; Tab(53); Label4.Caption
  Print #DCRpt, "-------------------------------------------------------------------------"
  Print #DCRpt, Tab(2); "TransDate"; Tab(22); "Description"; Tab(47); "Trans Amt"; Tab(62); "Balance"
  Print #DCRpt, Tab(2); "---------"; Tab(22); "-----------"; Tab(47); "---------"; Tab(62); "-------"
  If fpTRList.ListCount >= fpTRList.ListIndex + 18 Then
    go2line = fpTRList.ListIndex + 18
  Else
    go2line = fpTRList.ListCount - 1
  End If
  gofrom = fpTRList.ListIndex
  For cnt = gofrom To go2line
    fpTRList.ListIndex = cnt
    fpTRList.col = 1
    Print #DCRpt, fpTRList.ColText
  Next
  Close #DCRpt
  SearchRec = 0
  BeenDone = False
  PrintTRListScreenDC
End Sub

Private Sub fpTRList_DblClick()
  ReDim DCTranRec(1) As DCTransRecType
  Dim DCTranRecLen As Integer
  Dim DCFile As Integer, WhatRec As Long
  Dim TType As String, TDesc As String
  fpTRList.col = 2                      'switch to the hidden RecNo. column
  WhatRec = Val(fpTRList.ColText)     'get customer recno
'  Stop
  DCTranRecLen = Len(DCTranRec(1))
  DCFile = FreeFile
  Open UBPath + "DCTRANS.DAT" For Random Shared As DCFile Len = DCTranRecLen
  Get DCFile, WhatRec, DCTranRec(1)
  Close DCFile
  Select Case DCTranRec(1).TransType
    Case 1 'Charge
      TType$ = "Decal Charge"
    Case 2 'Payment
      TType$ = "Decal Payment"
    Case 3  'Charge Void
      TType$ = "Void Charge"
    Case 4  'Payment Void
      TType$ = "Void Payment"
    Case Else
      TType$ = Str$(DCTranRec(1).TransType) + " ???"
  End Select
  frmTRDetailDC.Caption = "Transaction Detail - " + Label5.Caption + " " + Label7.Caption
  frmTRDetailDC.Label3 = Num2Date(DCTranRec(1).TransDate)
  frmTRDetailDC.Label4 = Using("#####.##", DCTranRec(1).TransAmount, True)
  frmTRDetailDC.Make = QPTrim$(DCTranRec(1).makemodel)
  frmTRDetailDC.Label5 = QPTrim$(DCTranRec(1).TRVinDesc)
  frmTRDetailDC.Oper = Str(PWcnt)
  frmTRDetailDC.Bal = Using("#####.##", DCTranRec(1).BalanceAfterTrans, True)
  frmTRDetailDC.Label6 = TType$
  frmTRDetailDC.Label6b = QPTrim$(DCTranRec(1).ExtraDesc)
  frmTRDetailDC.Cash = Using("#####.##", DCTranRec(1).CashAmount, True)
  If DCTranRec(1).TransTender = 4 Then
    frmTRDetailDC.Charge = Using("#####.##", DCTranRec(1).ChkAmount, True)
    frmTRDetailDC.Chk = 0
  Else
    frmTRDetailDC.Chk = Using("#####.##", DCTranRec(1).ChkAmount, True)
    frmTRDetailDC.Charge = 0
  End If
  frmTRDetailDC.State = QPTrim$(DCTranRec(1).StateTag)
  frmTRDetailDC.Sticker = QPTrim$(DCTranRec(1).Sticker)
  frmTRDetailDC.Expire = Num2Date(DCTranRec(1).ExpireDate)
  frmTRDetailDC.Cat = QPTrim$(DCTranRec(1).DecalCat)
  DoEvents
  'preload stuff here
  frmTRDetailDC.Show vbModal
  
'  Unload frmTRDispListDC
End Sub

Private Sub fpTRList_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      KeyCode = 0
      Call fpCmdExit_Click
    Case vbKeyReturn
      KeyCode = 0
      DoEvents
      Call fpTRList_DblClick  'fpCmdOK_Click
    Case vbKeyTab
      KeyCode = 0
      DoEvents
      Call fpCmdExit_Click
    Case Else:
  End Select
End Sub
