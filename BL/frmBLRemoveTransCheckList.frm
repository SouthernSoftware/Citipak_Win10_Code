VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmBLRemoveTransCheckList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Transaction Removal CheckList"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11652
   Icon            =   "frmBLRemoveTransCheckList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8868
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpListSearch 
      Height          =   3444
      Left            =   1788
      TabIndex        =   0
      Top             =   2232
      Width           =   8088
      _Version        =   196608
      _ExtentX        =   14266
      _ExtentY        =   6075
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
      Columns         =   5
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   0
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   1
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
      BorderStyle     =   1
      BorderColor     =   8454143
      BorderWidth     =   2
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ScrollHScale    =   2
      ScrollHInc      =   0
      ColsFrozen      =   0
      ScrollBarV      =   1
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   2
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   3
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
      ColDesigner     =   "frmBLRemoveTransCheckList.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   684
      Left            =   3612
      TabIndex        =   3
      Top             =   6672
      Width           =   1836
      _Version        =   131072
      _ExtentX        =   3238
      _ExtentY        =   1206
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmBLRemoveTransCheckList.frx":0D10
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   684
      Left            =   6204
      TabIndex        =   4
      Top             =   6672
      Width           =   1836
      _Version        =   131072
      _ExtentX        =   3238
      _ExtentY        =   1206
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmBLRemoveTransCheckList.frx":0EEE
   End
   Begin VB.Label lblCustName 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   3588
      TabIndex        =   2
      Top             =   1584
      Width           =   4476
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   750
      Index           =   1
      Left            =   1407
      Top             =   585
      Width           =   8655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Removal Checklist"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2847
      TabIndex        =   1
      Top             =   720
      Width           =   6015
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1407
      Top             =   525
      Width           =   8655
   End
End
Attribute VB_Name = "frmBLRemoveTransCheckList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class
  Dim DCnt As Integer

Private Sub cmdExit_Click()
  frmBLRemoveTransCustList.Show
  DoEvents
  MainLog ("Remove Transaction Checklist screen exited.")
  Unload frmBLRemoveTransCheckList
End Sub

Private Sub cmdProcess_Click()
  Dim CHandle As Integer
  Dim THandle As Integer
  Dim ZZ As Integer
  Dim CustRec As ARCustRecType
  Dim TRec As ARTransRecType
  Dim TransRec&
  Dim TRType$
  Dim RecCnt As Integer
  Dim RecList() As Long
  Dim ColVal$
  
  On Error GoTo ERRORSTUFF
  
  For ZZ = 0 To fpListSearch.ListCount - 1
    fpListSearch.Col = 4
    fpListSearch.Row = ZZ
    If fpListSearch.ColList = "X" Then
      RecCnt = RecCnt + 1
      ReDim Preserve RecList(0 To RecCnt) As Long
      fpListSearch.Col = 2
      fpListSearch.Row = ZZ
      RecList(RecCnt) = CLng(fpListSearch.ColList)
    End If
  Next ZZ
  
  If RecCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No transactions selected."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  Else
    OpenCustFile CHandle
    Get CHandle, GCustNum, CustRec
    frmBLWarnRemoveTrans.lblCust = QPTrim$(CustRec.CustName)
    frmBLWarnRemoveTrans.lblNum = RecCnt
    frmBLWarnRemoveTrans.Show vbModal
    If frmBLWarnRemoveTrans.fptxtChoice = "exit" Then
      Close
      Unload frmBLWarnRemoveTrans
      Exit Sub
    End If
  End If
  
  Unload frmBLWarnRemoveTrans
  
  OpenTransFile THandle
  For ZZ = 1 To RecCnt
    TransRec& = RecList(ZZ)
    Get THandle, TransRec&, TRec
    If TRec.DetailTransType = 0 Then
      Select Case TRec.TransType
        Case 1   '"Charge"
          CustRec.AcctBal = OldRound#(CustRec.AcctBal - TRec.TransAmount)
          TRType$ = " Charge"
        Case 2   '"Payment"
          CustRec.AcctBal = OldRound#(CustRec.AcctBal + TRec.TransAmount)
          TRType$ = "Payment"
        Case 6 'penalty
          CustRec.AcctBal = OldRound#(CustRec.AcctBal - TRec.TransAmount)
          CustRec.PenBal = OldRound#(CustRec.PenBal - TRec.TransAmount)
          TRType$ = "Penalty"
        Case 100 'undo Adjust down
          CustRec.AcctBal = OldRound#(CustRec.AcctBal + TRec.TransAmount)
          CustRec.PenBal = OldRound#(CustRec.PenBal + TRec.TransAmount)
          TRType$ = "Adjust Penalty Down"
        Case 101 'undo adjust up
          CustRec.AcctBal = OldRound#(CustRec.AcctBal - TRec.TransAmount)
          CustRec.PenBal = OldRound#(CustRec.PenBal - TRec.TransAmount)
          TRType$ = "Adjust Penalty Up"
        Case 102 'undo Adjust down
          CustRec.AcctBal = OldRound#(CustRec.AcctBal + TRec.TransAmount)
          CustRec.LicBal = OldRound#(CustRec.LicBal + TRec.TransAmount)
          TRType$ = "Adjust License Down"
        Case 103 'undo adjust up
          CustRec.AcctBal = OldRound#(CustRec.AcctBal - TRec.TransAmount)
          CustRec.LicBal = OldRound#(CustRec.LicBal - TRec.TransAmount)
          TRType$ = "Adjust License Up"
        Case Else
          frmBLMessageBoxJr.Label1.Caption = "This transaction's type could not be matched. The transaction removal was NOT successful."
          frmBLMessageBoxJr.Label1.Top = 700
          frmBLMessageBoxJr.Show vbModal
          Close
          Exit Sub
      End Select
    Else
'      Select Case TRec.DetailTransType
'        Case 101
          
    End If
    LSet TRec.CustomerNumber = "0"
    Put THandle, TransRec&, TRec
    MainLog ("Transaction # " + CStr(RecList(ZZ)) + " for customer " + QPTrim$(CustRec.CustName) + " dated on " + MakeRegDate(TRec.TransDate) + " for " + Using("$##,###,##0.00", TRec.TransAmount) + " was removed.")
  Next
  Put CHandle, GCustNum, CustRec
  Close
  
  Call ReLinkTransactions(Me)
  MsgBox "Transactions removed successfully."
  
  frmBLRemoveTransCustList.Show
  DoEvents
  Unload frmBLRemoveTransCheckList
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLRemoveTransCheckList", "cmdProcess_Click", Erl)
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

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      Call cmdExit_Click
      SendKeys "%C"
      KeyCode = 0
    Case vbKeyF10:
      Call cmdProcess_Click
      SendKeys "%P"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      ClearInUse PWcnt
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLRemoveTransCheckList.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim TranRec As ARTransRecType
  Dim THandle As Integer
  Dim CHandle As Integer
  Dim CustRec As ARCustRecType
  Dim PrevTranRec&
  Dim TRType$
  Dim TCnt As Integer
  
  On Error Resume Next
  OpenCustFile CHandle
  Get CHandle, GCustNum, CustRec
  Close CHandle
  lblCustName.Caption = QPTrim$(CustRec.BillName)
  OpenTransFile THandle
  PrevTranRec& = CustRec.FirstTrans

  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      DCnt = DCnt + 1
      Get THandle, PrevTranRec&, TranRec
      Select Case TranRec.TransType
      Case 1
        TRType$ = " Charge"
      Case 2
        TRType$ = "Payment"
      Case 6
        TRType$ = "Penalty"
      Case 9
        TRType$ = "Beg Bal"
      Case 100
        TRType$ = "DN Pen Adj."
      Case 101
        TRType$ = "UP Pen Adj."
      Case 102
        TRType$ = "DN Lic Adj."
      Case 103
        TRType$ = "UP Lic Adj."
        
      End Select
      fpListSearch.InsertRow = "  " & MakeRegDate(TranRec.TransDate) & Chr(9) & Left$(TRType$, 15) & Chr(9) & Str$(PrevTranRec&) & Chr(9) & Using("$##,##0.00", TranRec.TransAmount)
      PrevTranRec& = TranRec.NextTrans
    Loop
  End If
  
  Close THandle
  MainLog ("Remove Transaction Checklist screen opened.")
End Sub

Private Sub fpListSearch_Click()
  fpListSearch.Row = -1
  fpListSearch.Col = 4
  
  If fpListSearch.ColText = "" Then
    fpListSearch.ColText = "X"
    
  ElseIf fpListSearch.ColText = "X" Then
    fpListSearch.ColText = ""
  End If
    
End Sub

