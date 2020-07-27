VERSION 5.00
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmDepListing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deposit Payment Listing"
   ClientHeight    =   2076
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   9588
   Icon            =   "frmDepListing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2076
   ScaleWidth      =   9588
   StartUpPosition =   2  'CenterScreen
   Begin LpLib.fpList fpTRList 
      Height          =   912
      Left            =   384
      TabIndex        =   6
      Top             =   384
      Width           =   8796
      _Version        =   196608
      _ExtentX        =   15515
      _ExtentY        =   1609
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
      ColDesigner     =   "frmDepListing.frx":08CA
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7776
      TabIndex        =   1
      Top             =   1440
      Width           =   1092
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6456
      TabIndex        =   0
      Top             =   1440
      Width           =   1092
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To Select Item to Void, Double-Click Item or Highlight and Click Ok."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   216
      TabIndex        =   5
      Top             =   1512
      Width           =   6132
   End
   Begin VB.Line Line1 
      X1              =   336
      X2              =   9216
      Y1              =   288
      Y2              =   288
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   7488
      TabIndex        =   4
      Top             =   48
      Width           =   852
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   3672
      TabIndex        =   3
      Top             =   48
      Width           =   1644
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   672
      TabIndex        =   2
      Top             =   48
      Width           =   492
   End
End
Attribute VB_Name = "frmDepListing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BeenDone As Boolean
Dim RCnt As Integer, NumofRevs As Integer
Dim RevText$(1 To MaxRevsCnt)
Dim Metered(1 To MaxRevsCnt) As Boolean

Private Sub cmdExit_Click()
  SearchRec = 0
  BeenDone = False
  Unload Me
  frmDepositVoid.fpCmdVoid.Enabled = False
  frmDepositVoid.fpCmdMsg.Enabled = False
  frmDepositVoid.fpCmdTranHist.Enabled = False
End Sub

Private Sub cmdOk_Click()
  If fpTRList.SelCount > 0 Then
    Call fpTRList_DblClick
  End If
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
      RevText$(RCnt) = Left$(QPTrim$(UBSetUpRec(1).Revenues(RCnt).RevName), 14)
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      KeyCode = 0
      Call cmdExit_Click
    Case vbKeyReturn
      KeyCode = 0
      DoEvents
      Call fpTRList_DblClick  'fpCmdOK_Click
    Case Else:
  End Select
End Sub
Private Sub fpTRList_DblClick()
  ReDim UBTranRec(1) As UBTransRecType
  Dim UBTranRecLen As Integer
  Dim UBFile As Integer, WhatRec As Long, newcnt As Integer
  Dim TType As String, TDesc As String, NumofRevs As Integer
  fpTRList.col = 1                       'switch to the hidden RecNo. column
  WhatRec = Val(fpTRList.ColText)     'get recno
  UBTranRecLen = Len(UBTranRec(1))
  UBFile = FreeFile
  Open UBPath + "UBTRANS.DAT" For Random Shared As UBFile Len = UBTranRecLen
  Get UBFile, WhatRec, UBTranRec(1)
  Close UBFile
  If UBTranRec(1).TransType = TranDepositPayment Or UBTranRec(1).TransType = TranDepositPayment + 100 Then
      TDesc$ = QPTrim$(UBTranRec(1).TransDesc)
      If InStr(UBTranRec(1).TransDesc, "DEPOSIT") = 0 And Len(TDesc$) > 0 Then
        TType$ = "Deposit Payment " + Left$(QPTrim$(UBTranRec(1).TransDesc), 10)
      Else
        TType$ = "Deposit Payment"
      End If
  End If
  frmDepositVoid.fpTranNum = WhatRec
  frmDepositVoid.Label3 = Num2Date(UBTranRec(1).TransDate)
  frmDepositVoid.Label4 = Using("#####.##", UBTranRec(1).Transamt, True)
  frmDepositVoid.Label5 = UBTranRec(1).TransDesc
  frmDepositVoid.Label6 = TType$
  frmDepositVoid.lblDepOper = UBTranRec(1).OperatorNumber
  For RCnt = 1 To MaxRevsCnt
    frmDepositVoid.Label1(RCnt - 1) = RevText$(RCnt)
    frmDepositVoid.Revs(RCnt - 1) = Using("#####.##", UBTranRec(1).RevAmt(RCnt), True)
    If Len(RevText$(RCnt)) = 0 Then
      NumofRevs = RCnt - 1
      Exit For
    End If
  Next
  For newcnt = NumofRevs To 14
    frmDepositVoid.Label1(newcnt).Visible = False
    frmDepositVoid.Revs(newcnt).Visible = False
  Next

  DoEvents
  SearchRec = 0
  BeenDone = False
  Unload Me
End Sub

Private Sub fpTRList_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      KeyCode = 0
      Call cmdExit_Click
    Case vbKeyReturn
      KeyCode = 0
      DoEvents
      Call fpTRList_DblClick  'fpCmdOK_Click
    Case vbKeyTab
      KeyCode = 0
      DoEvents
      Call cmdExit_Click
    Case Else:
  End Select
End Sub

