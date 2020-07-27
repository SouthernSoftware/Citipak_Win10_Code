VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmTaxBillListold 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Outstanding Bills"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9900
   Icon            =   "frmTaxBillListold.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   9900
   StartUpPosition =   1  'CenterOwner
   Begin LpLib.fpList fpList1 
      Height          =   1485
      Left            =   1020
      TabIndex        =   0
      Top             =   960
      Width           =   7875
      _Version        =   196608
      _ExtentX        =   13891
      _ExtentY        =   2619
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
      Columns         =   6
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
      ColDesigner     =   "frmTaxBillListold.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClose 
      Height          =   495
      Left            =   8040
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
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
      ButtonDesigner  =   "frmTaxBillListold.frx":0D89
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdTag 
      Height          =   495
      Left            =   210
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3120
      Width           =   2055
      _Version        =   131072
      _ExtentX        =   3625
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
      ButtonDesigner  =   "frmTaxBillListold.frx":0F65
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   495
      Left            =   7200
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3120
      Width           =   2415
      _Version        =   131072
      _ExtentX        =   4260
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
      ButtonDesigner  =   "frmTaxBillListold.frx":1147
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClear 
      Height          =   495
      Left            =   4050
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3120
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
      ButtonDesigner  =   "frmTaxBillListold.frx":132F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdTagAll 
      Height          =   495
      Left            =   2370
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3120
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
      ButtonDesigner  =   "frmTaxBillListold.frx":150A
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdReset 
      Height          =   495
      Left            =   5760
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1335
      _Version        =   131072
      _ExtentX        =   2355
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
      ButtonDesigner  =   "frmTaxBillListold.frx":16E7
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDetail 
      Height          =   495
      Left            =   360
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   240
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
      ButtonDesigner  =   "frmTaxBillListold.frx":18C2
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   3675
      Left            =   120
      Top             =   120
      Width           =   9615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Outstanding Tax Bill List"
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
      Left            =   3180
      TabIndex        =   2
      Top             =   360
      Width           =   3510
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   510
      Left            =   2925
      Top             =   285
      Width           =   4050
   End
End
Attribute VB_Name = "frmTaxBillListold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim FirstTime As Boolean
  Dim Processing As Boolean
  Dim EditMode As Boolean
  Dim XRecs() As Long
  Dim XCnt As Integer
  Dim ResetFlag As Boolean
  Dim TagAllFlag As Boolean
  Dim ClearFlag As Boolean
  Public PropRec As Long

Private Sub cmdClear_Click()
  Dim x As Integer
  Dim ThisIndex As Integer
  Dim y As Integer
  
  ClearFlag = True
  For y = 0 To fpList1.ListCount - 1
    fpList1.Row = y
    For x = 1 To XCnt
      fpList1.col = 2
      fpList1.Selected(y) = True
      If XRecs(y + 1) = CLng(fpList1.ColText) Then
        fpList1.col = 5
        If QPTrim$(fpList1.ColText) = "X" Then
          frmTaxMsg.Label1.Caption = "ERROR: The bill on line " + CStr(y + 1) + " has already been processed for payment. To clear this bill please use the 'Delete' procedure located on the payment menu."
          frmTaxMsg.Label1.Top = 800
          frmTaxMsg.Show vbModal
'          GoTo Again
          Exit Sub
        End If
      Else
        fpList1.col = 5
        fpList1.Selected(y) = True
        fpList1.ColText = ""
      End If
    Next x
Again:
  Next y
  
  For x = 0 To fpList1.ListCount - 1
    If fpList1.Selected = True Then
      ThisIndex = x
      Exit For
    End If
  Next x
  
  For x = 0 To fpList1.ListCount - 1
    fpList1.Row = x
    fpList1.col = 5
    fpList1.Selected = True
    fpList1.ColText = ""
  Next x
  
  KillFile TempBillRecs
  BillCnt = 0
  ReDim BillTrans(0 To 0) As Long
  
  fpList1.ListIndex = ThisIndex
  ClearFlag = False
End Sub

Private Sub cmdClose_Click()
'  Call cmdProcess_Click
  KillFile "C:\CPWork\billlist.dat"
  Unload Me
  DoEvents
End Sub

Private Sub cmdDetail_Click()
  Dim One As Integer
  Dim AHandle As Integer
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim RealRec As PropertyRecType
  Dim RRHandle As Integer
  Dim NumOfRRREcs As Long
  Dim ThisRPin$
  Dim ThisPPin$
  Dim x As Long
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  
  If fpList1.ListCount = 0 Then Exit Sub
  If fpList1.ListIndex = -1 Then
    Call TaxMsg(900, "Please make a selection.")
    Exit Sub
  End If
  One = 1
  AHandle = FreeFile
  Open "C:\CPWork\billlist.dat" For Output As AHandle
  Print #AHandle, One
  Close AHandle
  
  fpList1.Row = fpList1.ListIndex
  fpList1.col = 2
  PropRec = CLng(fpList1.ColText)
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Get TTHandle, PropRec, TaxTrans
  Close TTHandle
  ThisRPin = QPTrim$(TaxTrans.RealPin)
  ThisPPin = QPTrim$(TaxTrans.PersPin)
  
  If IsNumeric(ThisRPin) = True And IsNumeric(ThisPPin) = True Then
    If Val(ThisRPin) = 0 And Val(ThisPPin) > 0 Then
      GoTo TryPers
    ElseIf Val(ThisRPin) > 0 And Val(ThisPPin) = 0 Then
      GoTo TryReal
    Else
      Call TaxMsg(900, "This property detail is not available.")
      Exit Sub
    End If
  End If
  
  If QPTrim$(ThisRPin) = "" And QPTrim$(ThisPPin) = "" Then
    Call TaxMsg(900, "This property detail is not available.")
    Exit Sub
  End If
  
  If QPTrim$(ThisRPin) = "" And QPTrim$(ThisPPin) <> "" Then GoTo TryPers

TryReal:
  OpenRealPropFile RRHandle, NumOfRRREcs
  For x = 1 To NumOfRRREcs
    Get RRHandle, x, RealRec
    If QPTrim$(RealRec.RealPin) = ThisRPin Then
      PropRec = x
      Exit For
    End If
  Next x
  Close RRHandle
  
  If x > NumOfRRREcs Then
    If ThisPPin > 0 Then
      GoTo TryPers
    Else
      Call TaxMsg(900, "This property detail could not be found.")
      KillFile "C:\CPWork\billlist.dat"
      Exit Sub
    End If
  End If
  
  frmTaxRealDetail.Show vbModal
  KillFile "C:\CPWork\billlist.dat"
  Exit Sub
  
TryPers:
  OpenPersPropFile PHandle, NumOfPersRecs
  For x = 1 To NumOfPersRecs
    Get PHandle, x, PersRec
    If QPTrim$(PersRec.PropPin) = ThisPPin Then
       PropRec = x
       Exit For
    End If
  Next x
  Close PHandle
  
  If x > NumOfPersRecs Then
    Call TaxMsg(900, "This property detail could not be found.")
  Else
    frmTaxPersDetail.Show vbModal
  End If
  
  KillFile "C:\CPWork\billlist.dat"
  
End Sub

Private Sub cmdProcess_Click()
  Dim x As Integer, y As Integer
  Dim TPayRec As txPayListType
  Dim PayRec As txPayListType
  Dim NumOfPRecs As Integer
  Dim PHandle As Integer
  Dim TempHandle As Integer
  Dim NewCnt As Integer
  Dim Operator$
  Dim TempRec As TempPayList
  Dim NumOfTemps As Integer
  Dim Nextx As Integer
  
  KillFile TempBillRecs
  
  Operator$ = CStr(OperNum)
  Operator$ = QPTrim$(Operator$)
  Processing = True
  BillCnt = 0
  ReDim BillTrans(1 To 1) As Long
  OpenTempBillRecs TempHandle, NumOfTemps
  For x = 0 To fpList1.ListCount - 1
    fpList1.Row = x
    fpList1.col = 5
    fpList1.Selected = True
    If QPTrim$(fpList1.ColText) = "X" Then
      fpList1.Row = x
      fpList1.col = 2
      BillCnt = BillCnt + 1
      ReDim Preserve BillTrans(1 To BillCnt) As Long
      BillTrans(BillCnt) = CLng(fpList1.ColText)
      'save this bill's data collected from the load procedure
      fpList1.Selected(x) = True
      TempRec.BillRec = 1
      TempRec.CustRec = GCustNum
      'bill ptr
      TempRec.BillPtr = CDbl(fpList1.ColText)
      
      fpList1.col = 0 '2/14/06
      TempRec.BillDate = Date2Num(fpList1.ColText) '2/14/06

      Put TempHandle, x + 1, TempRec
    Else
      TempRec.BillRec = -1
      TempRec.CustRec = GCustNum
      fpList1.col = 2
      TempRec.BillPtr = CDbl(fpList1.ColText)

      fpList1.col = 0 '2/14/06
      TempRec.BillDate = Date2Num(fpList1.ColText) '2/14/06
      Put TempHandle, x + 1, TempRec
    End If
AlreadySaved:
  Next x
  
  Close TempHandle
  If BillCnt = 0 Then
    ReDim BillTrans(0 To 0) As Long
'    If Exist(TempBillRecs) Then KillFile TempBillRecs 'happens if
    KillFile TempBillRecs 'happens if
    'the use clears all tags and processes
  End If
  
AllDone:
  Unload Me
  DoEvents

End Sub

Private Sub cmdReset_Click()
  Dim x As Integer
  Dim y As Integer
  
  ResetFlag = True
  For y = 0 To fpList1.ListCount - 1
    For x = 1 To XCnt
      fpList1.col = 2
      fpList1.Row = y
      fpList1.Selected(y) = True
      If XRecs(x) = fpList1.ColText Then
        fpList1.col = 5
        fpList1.ColText = "   X"
      Else
        fpList1.col = 5
        fpList1.ColText = ""
      End If
    Next x
  Next y
  ResetFlag = False
End Sub

Private Sub cmdTag_Click()
  Call fpList1_DblClick
End Sub

Private Sub cmdTagAll_Click()
  Dim x As Integer
  Dim ThisIndex As Integer
  
  TagAllFlag = True
  For x = 0 To fpList1.ListCount - 1
    If fpList1.Selected = True Then
      ThisIndex = x
      Exit For
    End If
  Next x
  
  For x = 0 To fpList1.ListCount - 1
    fpList1.Row = x
    fpList1.col = 5
    fpList1.Selected = True
    fpList1.ColText = "   X"
  Next x
  
  fpList1.Row = ThisIndex
  fpList1.col = 5
  fpList1.Selected = True
  fpList1.ColText = "   X"
  fpList1.ListIndex = ThisIndex
  TagAllFlag = False
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    Call cmdProcess_Click
    KeyCode = 0
    Exit Sub
  End If

  Select Case KeyCode
    Case vbKeyDown:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      Call cmdClose_Click
      KeyCode = 0
    Case vbKeySpace:
      Call cmdTag_Click
      KeyCode = 0
    Case vbKeyF2:
      Call cmdDetail_Click
      KeyCode = 0
    Case vbKeyF3:
      Call cmdTagAll_Click
      KeyCode = 0
    Case vbKeyF7:
      Call cmdClear_Click
      KeyCode = 0
    Case vbKeyF10:
      Call cmdProcess_Click
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
  Dim Balance#
  Dim TransRecord&
  Dim BillTCnt As Integer
  Dim ListRec As txPayListType
  Dim LHandle As Integer
  Dim Operator$
  Dim NumOfLRecs As Integer
  Dim x As Integer
  Dim BillRecCnt As Integer
  Dim TempRec As TempPayList
  Dim NumOfTemps As Integer
  Dim TempHandle As Integer
  Dim LoadTagAll As Boolean
  
  LoadTagAll = False
  ResetFlag = False
  EditMode = False
  TagAllFlag = False
  ClearFlag = False
  If frmTaxPaymentEntry.EditFlag = True Then
    EditMode = True
    Call GetXRecs
  ElseIf BillCnt > 0 Then
    Call GetXRecs
  Else
    XCnt = 0
    ReDim XRecs(0 To 0) As Long
    LoadTagAll = True
  End If
  
  NumOfLRecs = 0
  Operator$ = CStr(OperNum)
  If Exist("CMXLOP" + Operator$ + ".DAT") Then
    OpenPayListFile LHandle, OperNum
    NumOfLRecs = LOF(LHandle) / Len(ListRec)
    ReDim BillRecs(1 To NumOfLRecs) As Long
    For x = 1 To NumOfLRecs
      Get LHandle, x, ListRec
      If ListRec.CustRec = GCustNum Then
        BillRecCnt = BillRecCnt + 1
        BillRecs(BillRecCnt) = ListRec.BillRec
      End If
    Next x
  End If
  
  FirstTime = True
  Processing = False
'  BillCnt = 0
'  ReDim BillTrans(0 To 0) As Long
  OpenTaxCustFile CHandle, NumOfCRecs
  Get CHandle, GCustNum, TaxRec
  Close CHandle
  TransRecord& = TaxRec.LastTrans
  
  OpenTaxTransFile THandle, NumOfTRecs
  
  Do While TransRecord& <> 0
    Get THandle, TransRecord&, TaxTrans
    If TaxTrans.TranType = 1 Then
      Balance# = OldRound#(TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      Balance# = OldRound#(Balance# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      Balance# = OldRound#(Balance# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      Balance# = OldRound#(Balance# - (TaxTrans.DiscAmt + TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd - TaxTrans.Revenue.LateListPd))
      Balance# = OldRound#(Balance# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd))
      
      If Balance# > 0 Then
        BillTCnt = BillTCnt + 1
        'collect data for each bill to be used to determine how any payments
        'will be distributed...oldest bill first
        If NumOfLRecs > 0 Then
          For x = 1 To BillRecCnt
            If BillRecs(x) = TransRecord& Then
              fpList1.InsertRow = MakeRegDate(TaxTrans.TransDate) + Chr(9) + Using$("####", TaxTrans.TaxYear) + Chr(9) + Using("#####0", TransRecord) + Chr(9) + Using$("$#,###,##0.00", TaxTrans.Amount) + Chr(9) + Using$("$#,###,##0.00", Balance#) + Chr(9) + "   X"
              GoTo Again
            End If
          Next x
          If x > NumOfLRecs Then
            fpList1.InsertRow = MakeRegDate(TaxTrans.TransDate) + Chr(9) + Using$("####", TaxTrans.TaxYear) + Chr(9) + Using("#####0", TransRecord) + Chr(9) + Using$("$#,###,##0.00", TaxTrans.Amount) + Chr(9) + Using$("$#,###,##0.00", Balance#) + Chr(9) + ""
            GoTo Again
          End If
        End If
        fpList1.InsertRow = MakeRegDate(TaxTrans.TransDate) + Chr(9) + Using$("####", TaxTrans.TaxYear) + Chr(9) + Using("#####0", TransRecord) + Chr(9) + Using$("$#,###,##0.00", TaxTrans.Amount) + Chr(9) + Using$("$#,###,##0.00", Balance#) + Chr(9) + ""
      End If
    End If
Again:
    TransRecord& = TaxTrans.LastTrans
  Loop
  Close THandle
  fpList1.ListIndex = 0
  
  If Exist(TempBillRecs) Then
    OpenTempBillRecs TempHandle, NumOfTemps
    For x = 1 To NumOfTemps
      Get TempHandle, x, TempRec
      fpList1.col = 5
      If TempRec.BillRec > 0 Then
        fpList1.Row = x - 1
        fpList1.Selected(x - 1) = True
        fpList1.ColText = "   X"
      End If
    Next x
    Close TempHandle
  End If
  FirstTime = False
  
  If LoadTagAll = True Then
    Call cmdTagAll_Click
    LoadTagAll = False
  End If
 
End Sub

Private Sub fpList1_DblClick()
  Dim x As Integer
  
  If FirstTime = True Then Exit Sub
  If Processing = True Then Exit Sub
  If ResetFlag = True Then Exit Sub
  If TagAllFlag = True Then Exit Sub
  If ClearFlag = True Then Exit Sub
  fpList1.col = 5
  fpList1.Row = fpList1.ListIndex
  
  For x = 1 To XCnt
    fpList1.col = 2
    If XRecs(x) = CLng(fpList1.ColText) Then
      fpList1.col = 5
      If QPTrim$(fpList1.ColText) = "X" Then
        frmTaxMsg.Label1.Caption = "ERROR: The bill on line " + CStr(fpList1.Row + 1) + " has already been processed for payment. To clear this bill please use the 'Delete' procedure located on the payment menu."
        frmTaxMsg.Label1.Top = 800
        frmTaxMsg.Show vbModal
        Exit Sub
      End If
    End If
  Next x
  
  fpList1.col = 5
  If QPTrim$(fpList1.ColText) = "" Then
    fpList1.ColText = "   X"
  Else
    fpList1.ColText = ""
  End If
End Sub

Private Sub GetXRecs()
  Dim ListRec As txPayListType
  Dim LHandle As Integer
  Dim NumOfLRecs As Integer
  Dim x As Integer
  Dim NewX As Integer
  'get these so that if a user attempts to untag one that has
  'already been saved in the payment process they will be alerted
  'to delete the payment instead of untagging here
  NewX = 0
  ReDim XRecs(1 To 1) As Long
  OpenPayListFile LHandle, OperNum
  NumOfLRecs = LOF(LHandle) / Len(ListRec)
  For x = 1 To NumOfLRecs
    Get LHandle, x, ListRec
      If ListRec.CustRec = GCustNum Then
        NewX = NewX + 1
        ReDim Preserve XRecs(1 To NewX) As Long
        XRecs(NewX) = ListRec.BillRec
      End If
   Next x
   XCnt = NewX
   
   Close LHandle
    
End Sub

'Private Sub fpList1_KeyDown(KeyCode As Integer, Shift As Integer)
'  Select Case KeyCode
'    Case vbKeyEscape:
'      KeyCode = 0
'      Call cmdClose_Click
'    Case vbKeyReturn
'      KeyCode = 0
'      DoEvents
'      Call cmdProcess_Click   'fpCmdOK_Click
'    Case Else:
'  End Select
'End Sub
