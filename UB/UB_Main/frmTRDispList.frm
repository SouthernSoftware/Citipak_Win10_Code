VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmTRDispList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction List"
   ClientHeight    =   6270
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   9300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin LpLib.fpList fpTRList 
      Height          =   3660
      Left            =   120
      TabIndex        =   0
      Top             =   1470
      Width           =   9090
      _Version        =   196608
      _ExtentX        =   16034
      _ExtentY        =   6456
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   9
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
      Columns         =   2
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
      ColDesigner     =   "frmTRDispList.frx":0000
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdOK 
      Height          =   480
      Left            =   6696
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
      ButtonDesigner  =   "frmTRDispList.frx":0354
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdExit 
      Height          =   480
      Left            =   7944
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
      ButtonDesigner  =   "frmTRDispList.frx":052D
   End
   Begin fpBtnAtlLibCtl.fpBtn fpcmdPrint 
      Height          =   480
      Left            =   5568
      TabIndex        =   9
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
      ButtonDesigner  =   "frmTRDispList.frx":0708
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   6864
      TabIndex        =   10
      Top             =   168
      Width           =   2052
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Name/Acct"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Left            =   312
      TabIndex        =   8
      Top             =   144
      Width           =   6348
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   6885
      TabIndex        =   7
      Top             =   600
      Width           =   2265
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   3690
      TabIndex        =   6
      Top             =   600
      Width           =   3075
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   345
      TabIndex        =   5
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Trans Date                           Description                            Trans Amt                Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   264
      TabIndex        =   2
      Top             =   1152
      Width           =   8796
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Item or Highlight and Double-Click for Details."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
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
Attribute VB_Name = "frmTRDispList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BeenDone As Boolean
Dim RCnt As Integer, NumofRevs As Integer
Dim RevText$(1 To MaxRevsCnt)
Dim Metered(1 To MaxRevsCnt) As Boolean

Private Sub fpCmdExit_Click()
  SearchRec = 0
  BeenDone = False
  Unload frmTRDispList
End Sub

Private Sub Form_Activate()
  SearchRec& = 0
  If Not BeenDone Then
    BeenDone = True
    Me.fpTRList.ListIndex = 0
    ReDim UBSetUpRec(1) As UBSetupRecType
    LoadUBSetUpFile UBSetUpRec(), 0
    NumofRevs = MaxRevsCnt
    For RCnt = 1 To MaxRevsCnt
      If RCnt < 10 Then
        RevText$(RCnt) = " " + Str$(RCnt) + ") " + Left$(QPTrim$(UBSetUpRec(1).Revenues(RCnt).RevName), 14)
      Else
        RevText$(RCnt) = Str$(RCnt) + ") " + Left$(QPTrim$(UBSetUpRec(1).Revenues(RCnt).RevName), 14)
      End If
      If Len(QPTrim$(RevText$(RCnt))) = 0 Then
        NumofRevs = RCnt - 1
        Exit For
      End If
      If UBSetUpRec(1).Revenues(RCnt).UseMtr = "Y" Then
        Metered(RCnt) = True
      End If
    Next
  End If
End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'  Select Case KeyCode
'    Case vbKeyEscape:
'      KeyCode = 0
'      Call fpCmdExit_Click
'    Case vbKeyReturn
'      KeyCode = 0
'      DoEvents
'      Call fpTRList_DblClick  'fpCmdOK_Click
'    Case Else:
'  End Select
'End Sub

Private Sub fpCmdOk_Click()
  If fpTRList.SelCount > 0 Then
    Call fpTRList_DblClick
  End If
End Sub

Private Sub fpcmdPrint_Click()
  Dim ReportFile As String, UBRpt As Integer, cnt As Integer, go2line As Integer
  Dim gofrom As Integer
  ReportFile$ = UBPath$ + "UBtrlist.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
  Print #UBRpt, ""
  Print #UBRpt, Now
  Print #UBRpt, Tab(2); "Customer Transaction List"
  Print #UBRpt, Tab(2); QPTrim$(Label5.Caption); Tab(53); Label7.Caption
  Print #UBRpt, Tab(2); Label2.Caption; Tab(24); Label3.Caption; Tab(53); Label4.Caption
  Print #UBRpt, "-------------------------------------------------------------------------"
  Print #UBRpt, Tab(2); "TransDate"; Tab(22); "Description"; Tab(50); "Trans Amt"; Tab(65); "Balance"
  Print #UBRpt, Tab(2); "---------"; Tab(22); "-----------"; Tab(50); "---------"; Tab(65); "-------"
  If fpTRList.ListCount >= fpTRList.ListIndex + 18 Then
    go2line = fpTRList.ListIndex + 18
  Else
    go2line = fpTRList.ListCount - 1
  End If
  gofrom = fpTRList.ListIndex
  For cnt = gofrom To go2line
    fpTRList.ListIndex = cnt
    fpTRList.col = 0
    Print #UBRpt, fpTRList.ColText
  Next
  Close #UBRpt
  SearchRec = 0
  BeenDone = False
  PrintTRListScreen
End Sub

Private Sub fpTRList_DblClick()
  ReDim UBTranRec(1) As UBTransRecType
  Dim UBTranRecLen As Integer
  Dim UBFile As Integer, WhatRec As Long
  Dim TType As String, TDesc As String
  fpTRList.col = 1                       'switch to the hidden RecNo. column
  WhatRec = Val(fpTRList.ColText)     'get customer recno
'  Stop
  UBTranRecLen = Len(UBTranRec(1))
  UBFile = FreeFile
  Open UBPath + "UBTRANS.DAT" For Random Shared As UBFile Len = UBTranRecLen
  Get UBFile, WhatRec, UBTranRec(1)
  Close UBFile
  Select Case UBTranRec(1).TransType
    Case TranUtilityBill, TranUtilityBill + 100
      TType$ = "Utility Bill "
    Case TranLateCharge, TranReconnectFee, TranLateCharge + 100, TranReconnectFee + 100
      TType$ = "Penalty, Reconnect Fee"
    Case TranBillPayment, TranBillPayment + 100
      TDesc$ = QPTrim$(UBTranRec(1).TransDesc)
'      If InStr(UBTranRec(1).TransDesc, "PAYMENT") = 0 And Len(TDesc$) > 0 Then
'        TType$ = "Utility Payment " + Left$(QPTrim$(UBTranRec(1).TransDesc), 10)
'      Else
        TType$ = "Utility Payment"
'      End If
  '  Case TranPenaltyPayment
  '    TType$ = "Penalty Payment"
    Case TranPenaltyCharge
      TType$ = "Penalty/Late Fee"
    Case TranAppliedDeposit
      TType$ = "Applied Deposit"
    Case TranDepositPayment, TranDepositPayment + 100
      TDesc$ = QPTrim$(UBTranRec(1).TransDesc)
'      If InStr(UBTranRec(1).TransDesc, "DEPOSIT") = 0 And Len(TDesc$) > 0 Then
'        TType$ = "Deposit Payment " + Left$(QPTrim$(UBTranRec(1).TransDesc), 10)
'      Else
        TType$ = "Deposit Payment"
'      End If
    Case TranDraftPayment
      TType$ = "Draft Payment"
    Case TranBeginBalance, TranBeginBalance + 100
      TType$ = "Beginning Balance"
    Case 9
      TType$ = "Deposit Refund"
    Case TranUpwardAdjustment
      TType$ = "Upward Adjustment"
    Case TranDownwardAdjustment
      TType$ = "Downward Adjustment"
    Case TranOverPayAdjustment
      TType$ = "Payment Adjustment"
    Case TranDepCreditRemoval
      TType$ = "Deposit Credit Removal"
    Case TranDepPaymentVoid
      TType$ = "Deposit Payment Void"
    Case Else
      TType$ = Str$(UBTranRec(1).TransType) + " ???"
  End Select
  frmTRDetail.Caption = "Transaction Detail - " + Label5.Caption
  frmTRDetail.Label3 = Num2Date(UBTranRec(1).TransDate)
  frmTRDetail.Label4 = Using("########.##", UBTranRec(1).Transamt, True)
'  If Len(QPTrim$(UBTranRec(1).BillMsg)) > 0 Then
'    frmTRDetail.Label5 = QPTrim$(UBTranRec(1).BillMsg)
'  Else
    frmTRDetail.Label5 = UBTranRec(1).TransDesc
'  End If
  frmTRDetail.Label6 = TType$
  frmTRDetail.Label6b = QPTrim$(UBTranRec(1).BillMsg)
  For RCnt = 1 To MaxRevsCnt
    frmTRDetail.Label1(RCnt - 1) = RevText$(RCnt)
   ' frmTRDetail.Label1(RCnt - 1).Caption
    frmTRDetail.Revs(RCnt - 1) = Using("########.##", UBTranRec(1).RevAmt(RCnt), True)
    frmTRDetail.Tax(RCnt - 1) = Using("#####.##", UBTranRec(1).TaxAmt(RCnt), True)
  Next
  For RCnt = 1 To 7
  If Metered(RCnt) Then
    frmTRDetail.CurRead(RCnt - 1) = Using$("#########", UBTranRec(1).CurRead(RCnt))
    frmTRDetail.PreRead(RCnt - 1) = Using$("#########", UBTranRec(1).PrevRead(RCnt))
    If QPTrim$(UBTranRec(1).ESTREAD(RCnt)) = "" Then
      frmTRDetail.ReadEst(RCnt - 1) = "N"
    Else
      frmTRDetail.ReadEst(RCnt - 1) = "Y"
    End If
  End If
  Next

  DoEvents
  'preload stuff here
  frmTRDetail.Show vbModal
  
'  Unload frmTRDispList
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
