VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmFADeptList 
   BackColor       =   &H008F8265&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Fixed Assets Department List"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   6960
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin LpLib.fpList fpList1 
      Height          =   4050
      Left            =   765
      TabIndex        =   0
      Top             =   1200
      Width           =   5490
      _Version        =   196608
      _ExtentX        =   9684
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
      ColDesigner     =   "frmFADeptList.frx":0000
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   495
      Left            =   2010
      TabIndex        =   2
      Top             =   5802
      Width           =   1350
      _Version        =   131072
      _ExtentX        =   2381
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmFADeptList.frx":0354
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClose 
      Height          =   495
      Left            =   3762
      TabIndex        =   3
      Top             =   5802
      Width           =   1350
      _Version        =   131072
      _ExtentX        =   2381
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmFADeptList.frx":0567
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   636
      Left            =   1482
      Top             =   330
      Width           =   4044
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Fixed Assets Departments Lookup"
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
      Left            =   1608
      TabIndex        =   1
      Top             =   480
      Width           =   3900
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   6588
      Left            =   42
      Top             =   42
      Width           =   6876
   End
End
Attribute VB_Name = "frmFADeptList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsFATextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdClose_Click()
   Unload frmFADeptList
   DoEvents
End Sub

Private Sub cmdHelp_Click()
  MsgBox "Double click an item or highlight an item and press enter and the item will appear in the correct field."
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then 'can't select a line with nothing on it
    If fpList1.ListIndex <> -1 Then GoTo DeptAlreadySelected
    KeyCode = 0
    Exit Sub
DeptAlreadySelected:
    fpList1.Col = 0
    If QPTrim$(fpList1.ColText) = "" Then
      MsgBox "No department has been selected"
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
  Dim DHandle As Integer
  Dim DeptRec As FADeptCodeType
  Dim x As Integer
  Dim NumOfDepts As Integer
  Dim DeptIdx As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxRecNums As Integer
  
  On Error GoTo ERRORSTUFF
  
  OpenDeptIdxFile DIdxHandle
  DIdxRecNums = LOF(DIdxHandle) \ Len(DeptIdx)
  If DIdxRecNums = 0 Then
    MsgBox "No departments saved in index."
    Close
    Exit Sub
  End If
  
  ReDim DIdx(1 To DIdxRecNums) As Integer
  
  For x = 1 To DIdxRecNums
    Get DIdxHandle, x, DeptIdx
    DIdx(x) = DeptIdx.DeptRecNum 'load array with record number pointers
  Next x
    
  Close DIdxHandle
  
  OpenFADeptCodeFile DHandle
  NumOfDepts = LOF(DHandle) / Len(DeptRec)
  
  If NumOfDepts = 0 Then
    Close
    MsgBox "No records on file"
    Exit Sub
  End If
  
  If Not Exist("edititemopen.dat") And Not Exist("editdeptopen.dat") Then
    fpList1.InsertRow = "ALL" 'only want this to appear in screens that
    'use ALL as a way of processing all departments...some screens (frmFAEditItemWTabs)
    'need this list from which to select a specific department where ALL is
    'not valid
  End If
  
AssByCodeRpt:
  For x = 1 To NumOfDepts
    Get DHandle, DIdx(x), DeptRec 'retrieve data in order of dept numbers
    fpList1.InsertRow = DeptRec.DeptNum & "  " & Chr(9) & " " & DeptRec.DeptDesc
  Next x
  Close DHandle
  fpList1.Row = 0
  fpList1.Selected = True 'set focus to first line of list
  
  Exit Sub
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFADeptList", "Load", Erl)
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
    Unload Me
End Sub

Private Sub fpList1_DblClick()
  Dim ThisOne$
  Dim DHandle As Integer
  Dim NumOfRecs As Integer
  Dim DeptRec As FADeptCodeType
  Dim x As Long
  Dim DeptNum As Integer
  Dim Found As Boolean
  
  On Error GoTo ERRORSTUFF
  'this department list is an option in several screens...
  'the following looks to see what screen is active and returns
  'the selected dept to that screen...the calling screens create
  'temporary .dat files so they can be tracked
  If Exist("editdeptopen.dat") Then 'department edit screen
    GoTo EditDeptOpen
  ElseIf Exist("edititemopen.dat") Then 'item edit screen
    fpList1.Row = -1
    fpList1.Col = 0
    ThisOne = QPTrim$(fpList1.ColText)
    frmFAEditItemWTabs.fptxtDeptNum.Text = ThisOne
    Unload frmFADeptList
    frmFAEditItemWTabs.vaTabPro1.ActiveTab = 0
    frmFAEditItemWTabs.fptxtDeptNum.SetFocus
    Exit Sub
  ElseIf Exist("newadddelrptopen.dat") Then 'add/delete report
    fpList1.Row = -1
    fpList1.Col = 0
    ThisOne = QPTrim$(fpList1.ColText)
    frmFAItemsAddDelOptRpt.fptxtDeptNum.Text = ThisOne
    Unload frmFADeptList
    frmFAItemsAddDelOptRpt.fptxtDeptNum.SetFocus
    Exit Sub
  ElseIf Exist("masteritemlistopen.dat") Then 'master item list report
    fpList1.Row = -1
    fpList1.Col = 0
    ThisOne = QPTrim$(fpList1.ColText)
    frmFAMasterItemListing.fptxtDeptNum.Text = ThisOne
    Unload frmFADeptList
    frmFAMasterItemListing.fptxtDeptNum.SetFocus
    Exit Sub
  ElseIf Exist("itemchecklist.dat") Then 'item check list report
    fpList1.Row = -1
    fpList1.Col = 0
    ThisOne = QPTrim$(fpList1.ColText)
    frmFAItemCheckList.fptxtDeptNum.Text = ThisOne
    Unload frmFADeptList
    frmFAItemCheckList.fptxtDeptNum.SetFocus
    Exit Sub
  ElseIf Exist("valrpt.dat") Then 'item value report
    fpList1.Row = -1
    fpList1.Col = 0
    ThisOne = QPTrim$(fpList1.ColText)
    frmFAValueRange.fptxtDeptNum.Text = ThisOne
    Unload frmFADeptList
    frmFAValueRange.fptxtDeptNum.SetFocus
    Exit Sub
  ElseIf Exist("dprhistrpt.dat") Then 'depreciation history by year report
    fpList1.Row = -1
    fpList1.Col = 0
    ThisOne = QPTrim$(fpList1.ColText)
    frmFADprHistRpt.fptxtDeptNum.Text = ThisOne
    Unload frmFADeptList
    frmFADprHistRpt.fptxtDeptNum.SetFocus
    Exit Sub
  ElseIf Exist("Wrntyrpt.dat") Then 'warranty report
    fpList1.Row = -1
    fpList1.Col = 0
    ThisOne = QPTrim$(fpList1.ColText)
    frmWarrantyRpt.fptxtDeptNum.Text = ThisOne
    Unload frmFADeptList
    frmWarrantyRpt.fptxtDeptNum.SetFocus
    Exit Sub
  ElseIf Exist("assetbycoderpt.dat") Then
    fpList1.Row = -1
    fpList1.Col = 0
    ThisOne = QPTrim$(fpList1.ColText)
    frmFAAssByCodeRpt.fptxtDeptNum.Text = ThisOne
    Unload frmFADeptList
    frmFAAssByCodeRpt.fptxtDeptNum.SetFocus
    Exit Sub
  ElseIf Exist("assetbyfundrpt.dat") Then
    fpList1.Row = -1
    fpList1.Col = 0
    ThisOne = QPTrim$(fpList1.ColText)
    frmFAAssByFundRpt.fptxtDeptNum.Text = ThisOne
    Unload frmFADeptList
    frmFAAssByFundRpt.fptxtDeptNum.SetFocus
    Exit Sub
  End If
  
  Exit Sub
  
EditDeptOpen:
  'the dept edit screen requires the record number for this dept
  'so that the fields on that screen can be populated properly
  fpList1.Col = 0
  If QPTrim$(fpList1.ColText) = "" Then
    MsgBox "The department selection is not valid"
    Exit Sub
  Else
    DeptNum = Val(QPTrim$(fpList1.ColText)) 'capture selected data
  End If
   
  OpenFADeptCodeFile DHandle
  NumOfRecs = LOF(DHandle) \ Len(DeptRec) 'now look for it in
  'the dept records
  For x = 1 To NumOfRecs
    Get DHandle, x, DeptRec
  
    If DeptRec.DeptNum = DeptNum Then
      Found = True
      fpList1.Row = -1
      GDeptNum = x 'found it so assign global variable it's record number
      Exit For 'no need to look anymore
    Else
      Found = False
      GoTo NotAMatch
    End If
      
NotAMatch:
  Next x
  
  Close DHandle
  
  Call frmFAEditDeptCodes.LoadMe 'populate this screen with GDeptNum pointer data
  DoEvents
  Unload frmFADeptList
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFADeptList", "fpList1_DblClick", Erl)
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
    Unload Me
  
End Sub


