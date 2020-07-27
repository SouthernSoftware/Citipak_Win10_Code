VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmFATagList 
   BackColor       =   &H008F8265&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Tag List"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   8220
   FillColor       =   &H00C0FFFF&
   Icon            =   "frmFATagList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin LpLib.fpList fpList1 
      Height          =   2910
      Left            =   600
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2520
      Width           =   7125
      _Version        =   196608
      _ExtentX        =   12568
      _ExtentY        =   5133
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
      SelMax          =   2
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
      ColDesigner     =   "frmFATagList.frx":08CA
   End
   Begin VB.CheckBox optLast 
      BackColor       =   &H008F8265&
      Caption         =   "Last Tag Number"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1370
      Width           =   1815
   End
   Begin VB.CheckBox optFirst 
      BackColor       =   &H008F8265&
      Caption         =   "First Tag Number"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1370
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H008F8265&
      Caption         =   "Select either First or Last Tag Number then select a number from the list"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1305
      Left            =   690
      TabIndex        =   6
      Top             =   1056
      Width           =   6855
      Begin EditLib.fpText fptxtFirst 
         Height          =   420
         Left            =   645
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   720
         Width           =   2265
         _Version        =   196608
         _ExtentX        =   3995
         _ExtentY        =   741
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   10.5
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
         AutoAdvance     =   -1  'True
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
      Begin EditLib.fpText fptxtLast 
         Height          =   420
         Left            =   3720
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   720
         Width           =   2265
         _Version        =   196608
         _ExtentX        =   3995
         _ExtentY        =   741
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   10.5
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
         AutoAdvance     =   -1  'True
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
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdApply 
      Height          =   495
      Left            =   5160
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to commit the data entered above to memory."
      Top             =   5808
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
      ButtonDesigner  =   "frmFATagList.frx":0CA1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClose 
      Height          =   495
      Left            =   3432
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to commit the data entered above to memory."
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
      ButtonDesigner  =   "frmFATagList.frx":0EB6
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   495
      Left            =   1704
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to commit the data entered above to memory."
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
      ButtonDesigner  =   "frmFATagList.frx":10CA
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   636
      Left            =   2160
      Top             =   336
      Width           =   4044
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Tag Numbers Lookup"
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
      Left            =   2280
      TabIndex        =   5
      Top             =   480
      Width           =   3900
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   6588
      Left            =   56
      Top             =   48
      Width           =   8124
   End
End
Attribute VB_Name = "frmFATagList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Over As clsFATextBoxOverRider
Private Temp_Class As Resize_Class
Dim ClickFirst As Boolean
Dim ClickLast As Boolean

Private Sub cmdApply_Click()
  Call fpList1_DblClick
End Sub

Private Sub cmdClose_Click()
  Unload frmFATagList
  DoEvents
End Sub

Private Sub cmdHelp_Click()
  'the first part of this process is used solely with the depreciation history
  'report by item and populates the tag parameters according to the
  'user's wishes
  'the second part works for all other calling forms
  
  If Exist("dprhistbyitemrpt.dat") Then
    MsgBox "Highlight 'First Tag Number' and then select a GL number. The number selected populates the 'First Tag Number' field. Repeat the process for 'Last Tag Number'. Then press 'Apply' to populate the appropriate fields on the parent screen."
  Else
    MsgBox "Select a tag number then press F3 to populate the Item Edit screen with that tag number's data."
  End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyF10 Then
    If fpList1.ListIndex <> -1 Then GoTo TagAlreadySelected '8/6
    KeyCode = 0
    Exit Sub
TagAlreadySelected:
    fpList1.Col = 1
    If QPTrim$(fpList1.ColText) = "" Then
      MsgBox "No tag number has been selected"
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
    Case vbKeyF10:
      SendKeys "%A"
      Call fpList1_DblClick
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%x"
      Call cmdClose_Click
      KeyCode = 0
    Case vbKeyF1:
      SendKeys "%F"
      Call cmdHelp_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
   Dim TagRec As FAItemRecType
   Dim THandle As Integer
   Dim TotalAccts As Integer
   Dim x As Integer
   Dim TagIdx As TagNumbSortIdxType
   Dim TagIdxHandle As Integer
   
   On Error GoTo ERRORSTUFF
   'this procedure simply loads up an array with the
   'asset numbers in numeric order and loads the list
   'using this index
   If Exist("dprhistbyitemrpt.dat") Then
     fpList1.Height = 3200
     fpList1.Top = 2420
     optFirst.Value = 1
     fptxtFirst.Text = QPTrim$(frmFADprHistByItem.fptxtFirst.Text)
     fptxtLast.Text = QPTrim$(frmFADprHistByItem.fptxtLast.Text)
   Else
     fpList1.Top = 1500
     fpList1.Height = 3500
     cmdApply.Top = 5300
     cmdClose.Top = 5300
     cmdHelp.Top = 5300
     Frame1.Visible = False
     optFirst.Visible = False
     optLast.Visible = False
   End If
   
   OpenTagIdxFile TagIdxHandle
   TotalAccts = LOF(TagIdxHandle) \ Len(TagIdx)
   If TotalAccts = 0 Then Exit Sub
   
   ReDim TagIdxRecs(1 To TotalAccts) As Integer
   
   For x = 1 To TotalAccts
     Get TagIdxHandle, x, TagIdx
     TagIdxRecs(x) = TagIdx.DataRecNum
   Next x
   Close TagIdxHandle
   
   If Not Exist("FAITEMS.DAT") Then
     MsgBox "Path to FAITEMS.DAT could not be found"
     Exit Sub
   End If

   OpenFAItemFile THandle
   
   For x = 1 To TotalAccts
     Get THandle, TagIdxRecs(x), TagRec
     fpList1.InsertRow = QPTrim$(TagRec.ItemTag) & "  " & Chr(9) & " " & TagRec.IDESC1
   Next x
   Close THandle
   fpList1.Selected(0) = True
   Exit Sub
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFATagList", "Form Load", Erl)
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

Private Sub fpList1_Click()
  fpList1.Col = 0
  If optFirst.Value = 1 Or optFirst.BackColor = &HC0FFFF Then
    fptxtFirst.Text = QPTrim$(fpList1.ColText)
  Else
    fptxtLast.Text = QPTrim$(fpList1.ColText)
  End If
End Sub

Private Sub fpList1_DblClick()
  Dim FAHandle As Integer
  Dim NumOfRecs As Integer
  Dim FAItemRec As FAItemRecType
  Dim x As Long
  Dim TagNum$
  Dim Desc$
  Dim SerialNum$
  Dim PrintDesc$
  Dim Found As Boolean
  Dim One As Integer
  Dim FileHandle As Integer
  Dim MsgRslt As Integer
  
  On Error GoTo ERRORSTUFF
  
  If Exist("dprhistbyitemrpt.dat") Then
    frmFADprHistByItem.fptxtFirst.Text = QPTrim$(fptxtFirst.Text)
    frmFADprHistByItem.fptxtLast.Text = QPTrim$(fptxtLast.Text)
    Unload frmFATagList
    DoEvents
    Exit Sub
  End If
  
  fpList1.Col = 0
  'this procedure
  One = 1
  FileHandle = FreeFile
  'taglistopen.dat is used to tell the FAItemEdit report that
  'the user wants to change whatever data is currently there
  'to the data related to the asset selection here...that form goes through
  'an exit check process first when it sees this .dat file...It is also
  'used in the Depreciation History Report by Item
  
  Open "taglistopen.dat" For Output As FileHandle Len = 2
  
  Print #FileHandle, One
  Close FileHandle
  
  If Exist("edititemopen.dat") Then 'if this form is called
  'from the item edit form then this form hides until the data
  'that form needs is collected
    AddItemFlag = False
    frmFATagList.Hide
    Call frmFAEditItemWTabs.cmdExit_Click 'go through exit to take
    'advantage of the field checks
    If ItemChangeFlag = True Then 'set in edit item screen and
    'informs this sub that the user changed something and
    'wants to handle that change before loading that screen with new data
      ItemChangeFlag = False
      Unload frmFATagList
      Exit Sub
    End If
    'if no change was made then the program can continue to load
    'the new tag data to the edit screen
  End If
     
  fpList1.Col = 0
  'find the record number for the selected asset
  If QPTrim$(fpList1.ColText) = "" Then
    MsgBox "The tag number selection is not valid."
    Exit Sub
  Else
    TagNum$ = ReplaceString(fpList1.ColText, "-", "")
  End If
  
  'match up the data and assign the global GRecNum with
  'the appropriate record number
  OpenFAItemFile FAHandle
  NumOfRecs = LOF(FAHandle) \ Len(FAItemRec)
  For x = 1 To NumOfRecs
    Get FAHandle, x, FAItemRec
  
    If InStr(UCase$(ReplaceString(FAItemRec.ItemTag, "-", "")), TagNum$) > 0 Then 'And InStr(UCase$(PrintDesc$), Desc$) > 0 And InStr(FAItemRec.SERIALNO, SerialNum$) >= 0
      Found = True
      fpList1.Row = -1
      GRecNum = x
      Exit For
    Else
      Found = False
      GoTo NotAMatch
    End If
      
NotAMatch:
  Next x
  
  Close FAHandle
  
  frmFAEditItemWTabs.Caption = "Fixed Asset Edit Item"
  frmFAEditItemWTabs.Label2 = "Fixed Asset Edit Item"
  Call frmFAEditItemWTabs.LoadMe
  DoEvents
  Unload frmFATagList
  
SkipEditForm:

  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFATagList", "fpList1_DblClick", Erl)
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

Private Sub optFirst_GotFocus()
  optLast.Value = 0
  optFirst.BackColor = &HC0FFFF
  optFirst.ForeColor = &H0&
  optLast.BackColor = &H8F8265
  optLast.ForeColor = &H80000005
  optFirst.Value = 1

End Sub


Private Sub optLast_GotFocus()
  optFirst.Value = 0
  optLast.BackColor = &HC0FFFF
  optLast.ForeColor = &H0&
  optFirst.BackColor = &H8F8265
  optFirst.ForeColor = &H80000005
  optLast.Value = 1

End Sub
