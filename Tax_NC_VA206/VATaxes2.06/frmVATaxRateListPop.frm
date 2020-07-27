VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxRateListPop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Rate Table to Edit."
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8790
   Icon            =   "frmVATaxRateListPop.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   8790
   StartUpPosition =   1  'CenterOwner
   Begin LpLib.fpList fpListR 
      Height          =   1770
      Left            =   1200
      TabIndex        =   6
      Top             =   1200
      Width           =   6480
      _Version        =   196608
      _ExtentX        =   11430
      _ExtentY        =   3122
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
      Columns         =   4
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
      SelMax          =   -1
      AutoSearch      =   1
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
      ColDesigner     =   "frmVATaxRateListPop.frx":08CA
   End
   Begin LpLib.fpList fpListP 
      Height          =   1770
      Left            =   1200
      TabIndex        =   7
      Top             =   3720
      Width           =   6480
      _Version        =   196608
      _ExtentX        =   11430
      _ExtentY        =   3122
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
      Columns         =   4
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
      SelMax          =   -1
      AutoSearch      =   1
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
      ColDesigner     =   "frmVATaxRateListPop.frx":0C06
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   420
      Left            =   6600
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1572
      _Version        =   131072
      _ExtentX        =   2773
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmVATaxRateListPop.frx":0F42
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdOK 
      Height          =   420
      Left            =   6600
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6720
      Width           =   1572
      _Version        =   131072
      _ExtentX        =   2773
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmVATaxRateListPop.frx":1120
   End
   Begin VB.Label Label6 
      Caption         =   "Personal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1320
      TabIndex        =   9
      Top             =   3480
      Width           =   1572
   End
   Begin VB.Label Label5 
      Caption         =   "Real"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1320
      TabIndex        =   8
      Top             =   840
      Width           =   1572
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Method"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6060
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3660
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Opt Revenue"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1380
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "To Select Double-Click Item or Highlight and Click OK."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   732
      TabIndex        =   0
      Top             =   6480
      Width           =   5400
   End
End
Attribute VB_Name = "frmVATaxRateListPop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdOk_Click()
  If fpListR.SelCount > 0 Then
    Call fpListR_DblClick
  ElseIf fpListP.SelCount > 0 Then
    Call fpListP_DblClick
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%O"
      Call cmdOk_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Dim TblRec As OptRevRateTablesType
  Dim TRHandle As Integer
  Dim NumOfTRRecs As Integer
  Dim x As Integer
  Dim Method As String * 16
  Dim ThisRec As String * 14
  Dim ThisDesc As String * 20
  
  RateTblRec = 0
  OpenTaxRateTables TRHandle, NumOfTRRecs
  For x = 1 To NumOfTRRecs
    Get TRHandle, x, TblRec
    If TblRec.Deleted = True Then GoTo Deleted
    RSet ThisRec = CStr(TblRec.OptRevNum)
    RSet ThisDesc = QPTrim$(TblRec.Desc)
    If TblRec.Type = "F" Then
      RSet Method = "FLAT"
    ElseIf TblRec.Type = "S" Then
      RSet Method = "STEP FLAT"
    ElseIf TblRec.Type = "P" Then
      RSet Method = "STEP PCT"
    End If
    If TblRec.RevType = "R" Then
      fpListR.AddItem ThisRec + Chr(9) + ThisDesc + Chr(9) + Method + Chr(9) + CStr(x)
    ElseIf TblRec.RevType = "P" Then
      fpListP.AddItem ThisRec + Chr(9) + ThisDesc + Chr(9) + Method + Chr(9) + CStr(x)
    End If
Deleted:
  Next x
  Close TRHandle
  If fpListR.ListCount > 0 Then
    fpListR.ListIndex = 0
  ElseIf fpListP.ListCount > 0 Then
    fpListP.ListIndex = 0
  End If
 
End Sub

Private Sub fpListP_Click()
  fpListR.Action = ActionDeselectAll
End Sub

Private Sub fpListP_DblClick()
  Dim ThisIndex As Integer
  
  ThisIndex = fpListP.ListIndex
  If ThisIndex = -1 Then
    Call TaxMsg(900, "Please make a selection from the personal list.")
    Exit Sub
  End If
  
  fpListP.Row = ThisIndex
  fpListP.Col = 3
        
  If QPTrim$(fpListP.ColText) = "" Then
    RateTblRec = 0
  Else
    RateTblRec = CInt(fpListP.ColText)
  End If
  
  If RateTblRec = 0 Then
    Call TaxMsg(800, "ERROR: The index for the personal rate code selection could not be found. Please call Southern Software at 1-800-842-8190.")
    Exit Sub
  End If
  
  frmVATaxPRateTableFlatOnly.Show
  Call frmVATaxPRateTableFlatOnly.LoadMeEdit
  DoEvents
  Unload Me

End Sub

Private Sub fpListR_Click()
  fpListP.Action = ActionDeselectAll
End Sub

Private Sub fpListR_DblClick()
  Dim ThisIndex As Integer
  
  ThisIndex = fpListR.ListIndex
  If ThisIndex = -1 Then
    Call TaxMsg(900, "Please make a selection from the real list.")
    Exit Sub
  End If
  
  fpListR.Row = ThisIndex
  fpListR.Col = 3
  fpListR.Text = fpListR.Text
        
  If QPTrim$(fpListR.ColText) = "" Then
    RateTblRec = 0
  Else
    RateTblRec = CInt(fpListR.ColText)
  End If
  
  If RateTblRec = 0 Then
    Call TaxMsg(800, "ERROR: The index for the real rate code selection could not be found. Please call Southern Software at 1-800-842-8190.")
    Exit Sub
  End If
  
  frmVATaxRateTables.Show
  Call frmVATaxRateTables.LoadMeEdit
  DoEvents
  Unload Me
  
'ChangesTrue:
'
End Sub

Public Sub Update()
  Dim TblRec As OptRevRateTablesType
  Dim TRHandle As Integer
  Dim NumOfTRRecs As Integer
  Dim x As Integer
  Dim Method As String * 16
  Dim ThisRec As String * 14
  Dim ThisDesc As String * 20
  
  On Error GoTo ERRORSTUFF
  
  fpListR.Clear
  fpListP.Clear
  RateTblRec = 0
  OpenTaxRateTables TRHandle, NumOfTRRecs
  For x = 1 To NumOfTRRecs
    Get TRHandle, x, TblRec
    If TblRec.Deleted = True Then GoTo Deleted
    RSet ThisRec = CStr(TblRec.OptRevNum)
    RSet ThisDesc = QPTrim$(TblRec.Desc)
    If TblRec.Type = "F" Then
      RSet Method = "FLAT"
    ElseIf TblRec.Type = "S" Then
      RSet Method = "STEP FLAT"
    ElseIf TblRec.Type = "P" Then
      RSet Method = "STEP PCT"
    End If
    If TblRec.RevType = "R" Then
      fpListR.AddItem ThisRec + Chr(9) + ThisDesc + Chr(9) + Method + Chr(9) + CStr(x)
    ElseIf TblRec.RevType = "P" Then
      fpListP.AddItem ThisRec + Chr(9) + ThisDesc + Chr(9) + Method + Chr(9) + CStr(x)
    End If
Deleted:
  Next x
  Close TRHandle
  fpListR.ListIndex = 0
  fpListR.Action = ActionForceUpdate
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxRateListPop", "Update", Erl)
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
