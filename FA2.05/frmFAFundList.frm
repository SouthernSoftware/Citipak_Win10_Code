VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmFAFundList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fixed Assets Fund List"
   ClientHeight    =   6645
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   6990
   Icon            =   "frmFAFundList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin LpLib.fpList fpList1 
      Height          =   4050
      Left            =   840
      TabIndex        =   0
      ToolTipText     =   "Double click one of the fund numbers to activate it."
      Top             =   1230
      Width           =   5430
      _Version        =   196608
      _ExtentX        =   9578
      _ExtentY        =   7144
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
      Columns         =   2
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   0
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
      ColDesigner     =   "frmFAFundList.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   492
      Left            =   2022
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5790
      Width           =   1356
      _Version        =   131072
      _ExtentX        =   2392
      _ExtentY        =   868
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmFAFundList.frx":0BE6
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClose 
      Height          =   492
      Left            =   3774
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Click this button after selecting a fund number and it will appear on the edit screen."
      Top             =   5790
      Width           =   1356
      _Version        =   131072
      _ExtentX        =   2392
      _ExtentY        =   868
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmFAFundList.frx":0DF9
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   6588
      Left            =   96
      Top             =   36
      Width           =   6876
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Fixed Assets Funds Lookup"
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
      Height          =   444
      Left            =   1632
      TabIndex        =   1
      Top             =   468
      Width           =   3900
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   636
      Left            =   1548
      Top             =   324
      Width           =   4044
   End
End
Attribute VB_Name = "frmFAFundList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsFATextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdClose_Click()
   Unload frmFAFundList
   DoEvents
End Sub

Private Sub cmdHelp_Click()
  MsgBox "Double click a fund or highlight a fund and press enter and the fund will appear in the correct field."
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    If fpList1.ListIndex <> -1 Then GoTo FundAlreadySelected '8/6
    KeyCode = 0
    Exit Sub
FundAlreadySelected:
    fpList1.Col = 1
    If QPTrim$(fpList1.ColText) = "" Then
      MsgBox "No fund has been selected"
      Exit Sub
    Else
      Call fpList1_DblClick
      KeyCode = 0
      Exit Sub
    End If
  End If
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdClose_Click
      KeyCode = 0
    Case vbKeyF1:
      SendKeys "%H"
      Call cmdHelp_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Dim FHandle As Integer
  Dim FundRec As FAFundCodeType
  Dim x As Integer
  Dim NumOfFunds As Integer
  Dim FundIdx As FundNumbSortIdxType
  Dim FIdxHandle As Integer
  Dim FIdxRecNums As Integer
  
  On Error GoTo ERRORSTUFF
  
  OpenFundIdxFile FIdxHandle
  FIdxRecNums = LOF(FIdxHandle) \ Len(FundIdx)
  If FIdxRecNums = 0 Then
    MsgBox "No funds saved in index."
    Close
    Exit Sub
  End If
  
  ReDim FIdx(1 To FIdxRecNums) As Integer
  
  For x = 1 To FIdxRecNums
    Get FIdxHandle, x, FundIdx
    FIdx(x) = FundIdx.FundRecNum 'load array with
    'fund numbers in numeric order
  Next x
    
  Close FIdxHandle
  
  OpenFAFundCodeFile FHandle
  NumOfFunds = LOF(FHandle) / Len(FundRec)
  
  If NumOfFunds = 0 Then
    MsgBox "No records on file"
  End If
  If Exist("assetbyfundrpt.dat") Then
    fpList1.InsertRow = "ALL"
  End If
  
  For x = 1 To NumOfFunds
    Get FHandle, FIdx(x), FundRec 'populate list box
    fpList1.InsertRow = FundRec.FundNum & "  " & Chr(9) & " " & FundRec.FundDesc
  Next x
  Close FHandle
  fpList1.Row = 0
  fpList1.Selected = True 'put focus on first row
  
  Exit Sub
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAFundList", "frmFAItemEdit", Erl)
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
    ClearInUse (PWcnt)
    Terminate
    Unload Me
End Sub

Private Sub fpList1_DblClick()
  Dim ThisOne$
  Dim FHandle As Integer
  Dim NumOfRecs As Integer
  Dim FundRec As FAFundCodeType
  Dim x As Long
  Dim FundNum As Integer
  Dim Found As Boolean
  
  On Error GoTo ERRORSTUFF

  If Exist("editfundopen.dat") Then
    GoTo EditFundOpen 'called from edit fund screen
  ElseIf Exist("edititemopen.dat") Then 'called from edit item screen
    fpList1.Row = -1
    fpList1.Col = 0
    ThisOne = QPTrim$(fpList1.ColText)
    frmFAEditItemWTabs.vaTabPro1.ActiveTab = 0
    frmFAEditItemWTabs.fptxtFundNum.Text = ThisOne
    Unload frmFAFundList
    frmFAEditItemWTabs.fptxtFundNum.SetFocus
    Exit Sub
  ElseIf Exist("assetbyfundrpt.dat") Then
    fpList1.Row = -1
    fpList1.Col = 0
    ThisOne = QPTrim$(fpList1.ColText)
    frmFAAssByFundRpt.fptxtFundNum.Text = ThisOne
    Unload frmFAFundList
    frmFAAssByFundRpt.fptxtFundNum.SetFocus
    Exit Sub
  End If
  
  Exit Sub
  
EditFundOpen:
  
  fpList1.Col = 0
  If QPTrim$(fpList1.ColText) = "" Then
    MsgBox "The fund selection is not valid"
    Exit Sub
  Else
    FundNum = Val(QPTrim$(fpList1.ColText))
  End If
  'need to find the record number for the selected fund
  OpenFAFundCodeFile FHandle
  NumOfRecs = LOF(FHandle) \ Len(FundRec)
  For x = 1 To NumOfRecs
    Get FHandle, x, FundRec
    If FundRec.FundNum = FundNum Then
      Found = True
      fpList1.Row = -1
      GFundNum = x 'found a match so assign global and exit loop
      Exit For
    Else
      Found = False 'keep going...no match yet
      GoTo NotAMatch
    End If
      
NotAMatch:
  Next x
  
  Close FHandle
  
  Call frmFAEditFundCodes.LoadMe
  DoEvents
  Unload frmFAFundList
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAFundList", "fpList1_DblClick", Erl)
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
    ClearInUse (PWcnt)
    Terminate
    Unload Me
  
End Sub







