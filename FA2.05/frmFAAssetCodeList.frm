VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmFAAssetCodeList 
   BackColor       =   &H008F8265&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Fixed Assets Code List"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   6960
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6672
   ScaleMode       =   0  'User
   ScaleWidth      =   12938.46
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin LpLib.fpList fpList1 
      Height          =   3720
      Left            =   630
      TabIndex        =   0
      Top             =   1215
      Width           =   5730
      _Version        =   196608
      _ExtentX        =   10107
      _ExtentY        =   6562
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
      Columns         =   3
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
      ColDesigner     =   "frmFAAssetCodeList.frx":0000
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   495
      Left            =   2010
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to bring up a list of all current fund codes."
      Top             =   5802
      Width           =   1350
      _Version        =   131072
      _ExtentX        =   2381
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmFAAssetCodeList.frx":0380
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClose 
      Height          =   495
      Left            =   3762
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to bring up a list of all current fund codes."
      Top             =   5802
      Width           =   1350
      _Version        =   131072
      _ExtentX        =   2381
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmFAAssetCodeList.frx":0593
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   6588
      Left            =   48
      Top             =   96
      Width           =   6876
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Fixed Assets Code Numbers Lookup"
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   636
      Left            =   1482
      Top             =   330
      Width           =   4044
   End
End
Attribute VB_Name = "frmFAAssetCodeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsFATextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdClose_Click()
   Unload frmFAAssetCodeList 'this form is brough up modally
   DoEvents
End Sub

Private Sub cmdHelp_Click()
  MsgBox "Double click an asset or highlight an asset and press enter and the asset will appear in the correct field."
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then 'user presses the return/enter key to process
    If fpList1.ListIndex <> -1 Then GoTo AssetAlreadySelected 'only happens when
    'no row is highlighted
    KeyCode = 0
    Exit Sub
AssetAlreadySelected:
    fpList1.Col = 0
    If QPTrim$(fpList1.ColText) = "" Then 'if a row is selected
    'with no asset number then it is assumed no asset has been selected
      MsgBox "No asset has been selected"
      Exit Sub
    Else
      Call fpList1_DblClick 'otherwise treat this selected row has valid
      'and begin processing using the code in DblClick
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
   Dim CodeRec As FAAssetCodeRecType
   Dim CodeIdxRec As ACNumbSortIdxType
   Dim CodeIdxHandle As Integer
   Dim CodeIdxRecNum As Integer
   Dim ACHandle As Integer
   Dim TotalAccts As Integer
   Dim x As Integer, cnt As Integer
   Dim DoWhatFlag As BadFACodeNumOption
   Dim n As Integer
   Dim Nextx As Integer
   Dim Y As Integer
   Dim ThisText$, CodeRecNo As Integer
   Dim FAAssCnt As Integer
   
   On Error GoTo ERRORSTUFF
   If Not Exist("FAASSIDX.DAT") Then
     MsgBox "No Asset Codes saved in index."
     Exit Sub
   End If
   
   OpenAssIdxFile CodeIdxHandle
   CodeIdxRecNum = LOF(CodeIdxHandle) \ Len(CodeIdxRec)
   If CodeIdxRecNum = 0 Then 'file has been opened but nothing has
   'been saved
     MsgBox "No Asset Codes saved in index."
     Close
     Exit Sub
   End If
   
   ReDim ACIdx(1 To CodeIdxRecNum) As Integer
   For x = 1 To CodeIdxRecNum
     Get CodeIdxHandle, x, CodeIdxRec
       ACIdx(x) = CodeIdxRec.AssRecNum 'load up array with index references
   Next x
   
   If Exist("assetbycoderpt.dat") Then
     fpList1.InsertRow = "ALL"
   End If
   
   Close CodeIdxHandle
   OpenFACodeNameFile ACHandle
   FAAssCnt = LOF(ACHandle) \ Len(CodeRec)
   For x = 1 To FAAssCnt
     Get ACHandle, ACIdx(x), CodeRec
     If Len(QPTrim(CodeRec.ASSETCODE)) = 0 Then GoTo BadCode 'if the asset code is blank then
     'this is an invalid entry...should not happen because when asset codes are saved they
     'must be saved with a value or the save routine is blocked
     fpList1.InsertRow = QPTrim$(CodeRec.ASSETCODE) & " " & Chr$(9) & QPTrim$(CodeRec.AssetDesc) & " " & Chr$(9) & QPTrim$(CodeRec.AssetStatus)
BadCode:
   Next x
   Close ACHandle
   fpList1.Row = 0
   fpList1.Selected = True 'sets the focus on the first row
   Exit Sub
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAAssetCodeList", "Form Load", Erl)
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
  Dim AHandle As Integer
  Dim NumOfRecs As Integer
  Dim DeptRec As FAAssetCodeRecType
  Dim x As Long
  Dim DeptNum$
  Dim Found As Boolean
  On Error GoTo ERRORSTUFF
  'Since double clicking a row sends the data on that row to
  'different screens depending on which screen this screen is called from
  'the way this routine knows where data must be sent is by looking
  'for the temporary .dat file created by the sending screen.
  If Exist("editassetopen.dat") Then 'if the asset edit screen is
  'calling this form then jump to EditAssetOpen
    GoTo EditAssetOpen
  ElseIf Exist("edititemopen.dat") Then 'if the item edit screen is calling
  'this form then process the user's selection and send the results
  'to the appropriate place on the item edit screen then unload this form
    fpList1.Row = -1
    fpList1.Col = 0
    ThisOne = QPTrim$(fpList1.ColText) 'collect data here
    frmFAEditItemWTabs.vaTabPro1.ActiveTab = 0
    frmFAEditItemWTabs.fptxtGroupCode = ThisOne 'place data here
    Unload frmFAAssetCodeList 'close this form...the other one is still running
    frmFAEditItemWTabs.fptxtGroupCode.SetFocus
    Exit Sub
  ElseIf Exist("assetbycoderpt.dat") Then
    fpList1.Row = -1
    fpList1.Col = 0
    ThisOne = QPTrim$(fpList1.ColText) 'collect data here
    frmFAAssByCodeRpt.fptxtCodeNum.Text = ThisOne 'close this form...the other one is still running
    Unload frmFAAssetCodeList
    frmFAAssByCodeRpt.fptxtCodeNum.SetFocus
    Exit Sub
  End If
  
  Exit Sub
  
EditAssetOpen:
  
  fpList1.Col = 0
  If QPTrim$(fpList1.ColText) = "" Then 'no code saved
    MsgBox "The asset code selection is not valid."
    Exit Sub
  Else
    DeptNum = Val(QPTrim$(fpList1.ColText)) 'collect data in DeptNum
  End If
   
  OpenFACodeNameFile AHandle
  NumOfRecs = LOF(AHandle) \ Len(DeptRec)
  For x = 1 To NumOfRecs 'now find the selected code in the list
  'and assign the global variable, GCodeNum, the record number
    Get AHandle, x, DeptRec
    If InStr(DeptRec.ASSETCODE, DeptNum) > 0 Then
      Found = True
      fpList1.Row = -1
      GCodeNum = x
      Exit For
    Else
      Found = False
      GoTo NotAMatch
    End If
      
NotAMatch:
  Next x 'keep iterating until found...it has to be there because
  'the list was created from the same file
  Close AHandle
  
  'the caption for frmFAEditAssetCode will indicate that this
  'entry is for editing an existing code, not creating a new code
  frmFAEditAssetCode.Caption = "Fixed Asset Edit Asset Codes"
  frmFAEditAssetCode.Label2 = "Fixed Asset Edit Asset Codes"
  
  Call frmFAEditAssetCode.LoadMe
  DoEvents
  Unload frmFAAssetCodeList
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAAssetCodeList", "fpList1_DblClick", Erl)
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




