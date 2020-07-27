VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmTRDispList2 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction List 2"
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
      Height          =   3840
      Left            =   90
      TabIndex        =   0
      Top             =   1470
      Width           =   9090
      _Version        =   196608
      _ExtentX        =   16034
      _ExtentY        =   6773
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
      ColDesigner     =   "frmTRDispList2.frx":0000
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdOK 
      Height          =   480
      Left            =   6336
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5664
      Width           =   1284
      _Version        =   131072
      _ExtentX        =   2265
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
      ButtonDesigner  =   "frmTRDispList2.frx":0354
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdExit 
      Height          =   480
      Left            =   7728
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5664
      Width           =   1284
      _Version        =   131072
      _ExtentX        =   2265
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
      ButtonDesigner  =   "frmTRDispList2.frx":052B
   End
   Begin EditLib.fpText fpCustRecNo 
      Height          =   324
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1764
      _Version        =   196608
      _ExtentX        =   3111
      _ExtentY        =   572
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      NoSpecialKeys   =   3
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
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Left            =   1560
      TabIndex        =   8
      Top             =   144
      Width           =   6180
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
      Left            =   6888
      TabIndex        =   7
      Top             =   600
      Width           =   2052
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
      Left            =   3696
      TabIndex        =   6
      Top             =   600
      Width           =   2052
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
      Left            =   792
      TabIndex        =   5
      Top             =   600
      Width           =   2052
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
      Height          =   252
      Left            =   264
      TabIndex        =   1
      Top             =   5784
      Width           =   5604
   End
End
Attribute VB_Name = "frmTRDispList2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BeenDone As Boolean
Dim RCnt As Integer, NumofRevs As Integer
Dim RevText$(1 To MaxRevsCnt)
Dim Metered(1 To MaxRevsCnt) As Boolean
Dim fromform As Form, toform As Form, codeopt As Integer
Public Sub Wheretogo(xfrm As Form, tfrm As Form, Optional opt As Integer)
  Set fromform = xfrm
  Set toform = tfrm
  If opt <> 0 Then
    codeopt = opt
  Else
    codeopt = 0
  End If
  
  'If TransRec > 0 Then
   ' DeActivateControls Me
    DisTransList Val(fpCustRecNo)
  'End If
End Sub

Private Sub fpCmdExit_Click()
  SearchRec = 0
  BeenDone = False
  Load frmUBEditMenu
  DoEvents
  frmUBEditMenu.Show
  Unload frmTRDispList2
  Unload frmCustEditLookUP
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
Private Sub fpTRList_DblClick()
  ReDim UBTranRec(1) As UBTransRecType
  Dim UBTranRecLen As Integer
  Dim UBFile As Integer, WhatRec As Long
  Dim TType As String, TDesc As String
  fpTRList.col = 1                       'switch to the hidden RecNo. column
  WhatRec = Val(fpTRList.ColText)  'get customer recno
  If WhatRec > 0 Then
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
  frmUBTransactionEdit.Caption = "Transaction Detail - " + Label5.Caption
  frmUBTransactionEdit.txtDate = Num2Date(UBTranRec(1).TransDate)
  frmUBTransactionEdit.fpTotal = UBTranRec(1).Transamt
  frmUBTransactionEdit.fpTransDesc = UBTranRec(1).TransDesc
  frmUBTransactionEdit.fpTransType = TType$
  frmUBTransactionEdit.fpRunBal = UBTranRec(1).RunBalance
  For RCnt = 1 To 15
    frmUBTransactionEdit.fpRevSource(RCnt - 1) = RevText$(RCnt)
    frmUBTransactionEdit.fpAmount(RCnt - 1) = UBTranRec(1).RevAmt(RCnt)
    'frmTRDetail.Tax(RCnt - 1) = Using("#####.##", UBTranRec(1).TaxAmt(RCnt), True)
  Next
 

  DoEvents
  'preload stuff here
  frmUBTransactionEdit.fpTransNum = WhatRec
  frmUBTransactionEdit.fpCustRecNo = fpCustRecNo
  frmUBTransactionEdit.Show
  Unload frmTRDispList2
  Unload frmDisplayList
  Unload frmCustEditLookUP
  End If
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
Private Sub DisTransList(CustRec As Long)
  ReDim UBTranRec(1) As UBTransRecType
  ReDim UBCustRec(1) As NewUBCustRecType
  Dim UBCustRecLen As Integer, UBTranRecLen As Integer
  Dim PrevTranRec As Long
  Dim UBFile As Integer, dcnt As Integer
  Dim Build As String * 80
  Dim TType As String, TDesc As String
  Dim CurBal As Double, PreBal As Double
  
  frmInfo.Label1 = "Loading. . ."
  frmInfo.Show
  DoEvents

  UBCustRecLen = Len(UBCustRec(1))
  UBTranRecLen = Len(UBTranRec(1))
  
  UBFile = FreeFile
  Open UBCustFile For Random Shared As UBFile Len = UBCustRecLen
  Get UBFile, CustRec&, UBCustRec(1)
  Close UBFile

  CurBal# = UBCustRec(1).CurrBalance
  PreBal# = UBCustRec(1).PrevBalance
'
Top:
'
  UBFile = FreeFile
  Open UBPath + "UBTRANS.DAT" For Random Shared As UBFile Len = UBTranRecLen
  
  PrevTranRec& = UBCustRec(1).LastTrans
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      dcnt = dcnt + 1
      Get UBFile, PrevTranRec&, UBTranRec(1)
      LSet Build = " " + Num2Date(UBTranRec(1).TransDate)
      GoSub GetTransType
      Mid$(Build, 20) = TType$
      Mid$(Build, 48) = Using("#####.##", UBTranRec(1).Transamt, True)
'      'this will show th actual trans number in the list
'      'MID$(MChoice(DCnt).V, 50) = FUsing(STR$(PrevTranRec&), "######")
'      Mid$(Build, 55) = Str$(PrevTranRec&)
      Mid$(Build, 63) = Using("#####.##", UBTranRec(1).RunBalance, True)
      Mid$(Build$, 71) = Chr9$ + Str$(PrevTranRec&) 'Mid$(MChoice(DCnt).V, 61) = MKL$(PrevTranRec&)
      frmTRDispList2.fpTRList.AddItem Build$
      PrevTranRec& = UBTranRec(1).PrevTrans
    Loop
  Else
    frmTRDispList2.fpTRList.AddItem "NO TRANSACTIONS"
  End If
  
  Close UBFile
  frmTRDispList2.Label5.Caption = QPTrim(UBCustRec(1).CustName)
  frmTRDispList2.Label2 = "Balance: " + Using("#####.##", CurBal# + PreBal#, True)
  frmTRDispList2.Label3 = "Current:  " + Using("#####.##", CurBal#, True)
  frmTRDispList2.Label4 = "Previous:  " + Using("#####.##", PreBal#, True)
  Unload frmInfo
  DoEvents
  'frmTRDispList2.Show vbModal
  Erase UBTranRec, UBCustRec
  
Exit Sub

GetTransType:
'
  Select Case UBTranRec(1).TransType
  Case TranUtilityBill, TranUtilityBill + 100
    TType$ = "Utility Bill "
  Case TranLateCharge, TranReconnectFee, TranLateCharge + 100, TranReconnectFee + 100
    TType$ = "Penalty, Reconnect Fee"
  Case TranBillPayment, TranBillPayment + 100
    TDesc$ = QPTrim$(UBTranRec(1).TransDesc)
    If InStr(UBTranRec(1).TransDesc, "PAYMENT") = 0 And Len(TDesc$) > 0 Then
      TType$ = "Utility Payment " + Left$(QPTrim$(UBTranRec(1).TransDesc), 10)
    Else
      TType$ = "Utility Payment"
    End If
'  Case TranPenaltyPayment
'    TType$ = "Penalty Payment"
  Case TranPenaltyCharge
    TType$ = "Penalty/Late Fee"
  Case TranAppliedDeposit
    TType$ = "Applied Deposit"
  Case TranDepositPayment, TranDepositPayment + 100
    TDesc$ = QPTrim$(UBTranRec(1).TransDesc)
    If InStr(UBTranRec(1).TransDesc, "DEPOSIT") = 0 And Len(TDesc$) > 0 Then
      TType$ = "Deposit Payment " + Left$(QPTrim$(UBTranRec(1).TransDesc), 10)
    Else
      TType$ = "Deposit Payment"
    End If
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
    TType$ = "DepCrRemvl " + Left$(QPTrim$(UBTranRec(1).BillMsg), 10)
  Case TranDepPaymentVoid
    TType$ = "DepPayVoid " + Left$(QPTrim$(UBTranRec(1).BillMsg), 10)
  Case Else
    TType$ = Str$(UBTranRec(1).TransType) + " ???"
  End Select

Return

End Sub

