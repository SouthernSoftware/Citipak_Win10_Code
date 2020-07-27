VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmPRTPrevEmpLookUp 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Transactions | Previous Employee:"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmPRTPrevEmpLookUp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpList 
      Height          =   3255
      Left            =   1350
      TabIndex        =   4
      Top             =   4950
      Width           =   8940
      _Version        =   196608
      _ExtentX        =   15769
      _ExtentY        =   5741
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
      ColDesigner     =   "frmPRTPrevEmpLookUp.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPickList 
      Height          =   615
      Left            =   6960
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Press to populate the list box below with employee data."
      Top             =   3360
      Width           =   1935
      _Version        =   131072
      _ExtentX        =   3413
      _ExtentY        =   1085
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
      ButtonDesigner  =   "frmPRTPrevEmpLookUp.frx":0D93
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   615
      Left            =   4845
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Press to begin the payroll editing process for the employee whose number is entered above."
      Top             =   3360
      Width           =   1935
      _Version        =   131072
      _ExtentX        =   3413
      _ExtentY        =   1085
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
      ButtonDesigner  =   "frmPRTPrevEmpLookUp.frx":0FAB
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
      Height          =   615
      Left            =   2730
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Press ESC to exit this screen."
      Top             =   3360
      Width           =   1935
      _Version        =   131072
      _ExtentX        =   3413
      _ExtentY        =   1085
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
      ButtonDesigner  =   "frmPRTPrevEmpLookUp.frx":11C2
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin EditLib.fpText fptxtEmpNum 
      Height          =   444
      Left            =   5724
      TabIndex        =   0
      Top             =   2352
      Width           =   2556
      _Version        =   196608
      _ExtentX        =   4508
      _ExtentY        =   783
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483637
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
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 "
      MaxLength       =   10
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fpText5 
      Height          =   375
      Left            =   3270
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2427
      Width           =   2130
      _Version        =   196608
      _ExtentX        =   3746
      _ExtentY        =   656
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   13684944
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483637
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   12632256
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
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   2
      AlignTextV      =   1
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   " Employee Number:"
      CharValidationText=   ""
      MaxLength       =   255
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fpText2 
      Height          =   732
      Left            =   3108
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1152
      Width           =   5412
      _Version        =   196608
      _ExtentX        =   9546
      _ExtentY        =   1291
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   13684944
      ForeColor       =   65535
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483637
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   1
      BorderColor     =   -2147483643
      BorderWidth     =   3
      ButtonDisable   =   0   'False
      ButtonHide      =   -1  'True
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   1
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   0
      ControlType     =   1
      Text            =   " Transaction Entry / Edit"
      CharValidationText=   ""
      MaxLength       =   255
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483643
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtNothing 
      Height          =   3900
      Left            =   2064
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   576
      Width           =   7428
      _Version        =   196608
      _ExtentX        =   13102
      _ExtentY        =   6879
      Enabled         =   0   'False
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
      BackColor       =   13684944
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483637
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   3
      ThreeDOutsideStyle=   2
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   3
      ThreeDFrameWidth=   0
      BorderStyle     =   1
      BorderColor     =   -2147483630
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
      ThreeDText      =   4
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483640
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   -1  'True
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   1
      BorderDropShadowColor=   -2147483634
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   4128
      Left            =   1944
      Top             =   480
      Width           =   7692
   End
End
Attribute VB_Name = "frmPRTPrevEmpLookUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
'this process must know if this employee is paid
'hourly or by salary because the transaction editing
'is different for each
Public Enum PRTTypeOfPay
  PRTtopInvalidOption = 0
  PRTtopHourly
  PRTtopSalary
End Enum
Private m_topOption As PRTTypeOfPay
'// Create a property to get the Selection value.
'   NOTE: A Read-Only property has a Property Get but
'         no Property Let or Property Set
Property Get Selection() As PRTTypeOfPay
  Selection = m_topOption
End Property

Private Sub cmdEscape_Click()
  frmPayrollProcessingMenu.Show
  DoEvents
  Unload frmPRTPrevEmpLookUp

End Sub

Public Sub cmdPickList_Click()
  Dim NHandle As Integer, EHandle As Integer
  Dim IdxRecLen As Long, IdxFileSize&
  Dim NumOfRecs& '7/20
  Dim EmpIdxNNameRec As NumbSortIdxType
  Dim x&, Emp2Rec As EmpData2Type
  Dim THandle As Integer
  Dim TRec As TransRecType
  Dim TFlag$
  Dim FocusHere As Integer '7/20
  Dim cnt As Integer '7/20
  
  IdxRecLen = 2
  IdxFileSize& = FileSize(PRData + EmpIdxNName)
  NumOfRecs = IdxFileSize& \ IdxRecLen
  OpenEmpIdxNNameFile NHandle
  If NumOfRecs = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    Get NHandle, x, IdxBuff(x)
  Next x
  Close NHandle
  fpList.Clear
  OpenEmpData2File EHandle
  OpenTransWorkFile THandle

  For x = 1 To NumOfRecs
    Get EHandle, IdxBuff(x), Emp2Rec
    Get THandle, IdxBuff(x), TRec
    If TRec.TActive = 0 Then
      TFlag = "N"
    Else
      TFlag = "Y"
    End If
'*******New in 2002: If an employee has been terminated he/she will
'*******no longer show up in the pick list
  If Emp2Rec.Deleted = -1 Or Emp2Rec.EMPTDATE <> 0 Then
    GoTo SkipIt
  ElseIf Emp2Rec.EMPTDATE <> 0 Then
    GoTo SkipIt
  End If
  fpList.InsertRow = "    " & QPTrim$(Emp2Rec.EmpNo) & Chr(9) & "  " & Emp2Rec.EmpLName & Chr(9) & "  " & Emp2Rec.EmpFName & Chr(9) & "      " & TFlag
  
  'this if statement was inserted solely to refocus the
  'employee list back to where the user left it before
  'the transaction edit procedure began...this keeps the
  'user from having to scroll back to where they were
  '...NewListFlag is set to True in frmPRDCalcScr3
  '...this if statement locates the position of the last
  'employee selected
  If NewListFlag = True Then '7/20
    If IdxBuff(x) = RecNum Then '7/20
      FocusHere = cnt '7/20
    End If '7/20
  End If '7/20
  cnt = cnt + 1 '7/20
SkipIt:
  Next x
  Close THandle
  Close EHandle
  'this if statement brings the focus to the last employee selected and
  'resets NewListFlag to false
  If NewListFlag = True Then '7/20
    fpList.Selected(FocusHere) = True '7/20
    fpList.TopIndex = FocusHere
    NewListFlag = False '7/20
  Else
    fpList.SetFocus '5/27/04
    fpList.ListIndex = 0 '5/27/04
  End If '7/20
  
End Sub

Private Sub cmdProcess_Click()
  Dim EmpData2FileHandle As Integer
  Dim EmpData2FileRec As EmpData2Type
  Dim NumEmpRec As Integer, x As Long
  Dim EmployeeNumber As String, Found As Boolean
  Dim PayType As String
  Dim TransRec(1) As TransRecType
  Dim TDate$
  
  On Error GoTo ERRORSTUFF
  
  If QPTrim$(fptxtEmpNum.Text) = "" Then
    MsgBox "There must be a valid employee number entered in order to Process"
    fptxtEmpNum.SetFocus
    Exit Sub
  End If
  frmLoadingPRTransForm.Show
  DoEvents
  Call DeActivateControls 'the process of gathering
  'data for this employee and then to load the
  'next screen can take enough time so that a user
  'could conceivably try to Terminate the program so
  'we temporarily shut off the termination option
  EmployeeNumber = QPTrim$(fptxtEmpNum.Text)
  OpenEmpData2File EmpData2FileHandle
  NumEmpRec = LOF(EmpData2FileHandle) / Len(EmpData2FileRec)
  For x = 1 To NumEmpRec
     Get EmpData2FileHandle, x, EmpData2FileRec
       If InStr(EmpData2FileRec.EmpNo, EmployeeNumber) > 0 And _
       Len(QPTrim$(EmpData2FileRec.EmpNo)) = Len(QPTrim$(EmployeeNumber)) Then 'added
       'Len = Len on 8/23 because if an employee had #125 and another employee
       'had #12 and their names were exactly the same then
       'this process would stop when it saw #12 even
       'if you were trying to process #125...
         'check to see if further processing is unnecessary
         'because this employee has been terminated
         OpenTransWorkFile TRHandle
         Get TRHandle, x, TransRec(1)
           If TransRec(1).TActive = 0 Then
'********UPDATE: The Termination check is changing so that
'********no terminated employee is allowed past this point
'********plus no terminated employee will be allowed to show
'********up in the pick list
             If EmpData2FileRec.EMPTDATE <> 0 Then
               TDate = MakeRegDate(EmpData2FileRec.EMPTDATE)
               MsgBox "This employee was terminated on " & TDate
               Close
               Unload frmLoadingPRTransForm
               Call ActivateControls
               Exit Sub
             Else
               CreateEmpTransRecs x
               Get TRHandle, x, TransRec(1)
             End If
           End If
         Close TRHandle
         PayType = QPTrim$(EmpData2FileRec.EMPPTYPE)
         PayType = UCase$(PayType)
         If PayType = UCase$("Hourly") Then
           m_topOption = PRTtopHourly
         Else
           m_topOption = PRTtopSalary
         End If
         Found = True
         fpList.Row = -1
         RecNum = x
         Exit For
       Else
         Found = False
         GoTo NotAMatch
       End If
NotAMatch:
  Next x
  Close EmpData2FileHandle
  Close
  If Found = False Then
    MsgBox "No matching number found"
    fptxtEmpNum.SetFocus
    Call ActivateControls
    Exit Sub
  End If
  
  fpList.Clear
  fptxtEmpNum.Text = ""
  
  If m_topOption = PRTtopSalary Then
    frmSPRTThisEmp.Show
    DoEvents
  ElseIf m_topOption = PRTtopHourly Then
    frmHPRTThisEmp.Show
    DoEvents
  End If
  Unload frmPRTPrevEmpLookUp '7/20
  DoEvents
  Unload frmLoadingPRTransForm
  DoEvents
  Call ActivateControls
  DoEvents
  Exit Sub
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmPRTPrevEmpLookUp", "cmdProcess_Click", Erl)
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
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyReturn:
      If Len(fptxtEmpNum.Text) > 0 Then
        Call cmdProcess_Click
        KeyCode = 0
        Exit Sub
      End If
      fpList.Col = 1
      If QPTrim$(fpList.ColText) = "" Then
        MsgBox "No employee has been selected"
        Exit Sub
      Else
        Call fpList_DblClick
        KeyCode = 0
        Exit Sub
      End If
    Case vbKeyDown:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%M"
      Call cmdEscape_Click
      KeyCode = 0
    Case vbKeyF7:
      SendKeys "%F"
      Call cmdPickList_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      fpList.Col = 1
      If QPTrim$(fpList.ColText) = "" And QPTrim$(fptxtEmpNum.Text) = "" Then 'added
      'the second "And" on 8/26
        MsgBox "No employee has been selected"
        KeyCode = 0
        Exit Sub
      ElseIf Len(QPTrim$(fptxtEmpNum.Text)) <> 0 Then 'added 8/26
        Call cmdProcess_Click 'added 8/26
        KeyCode = 0 'added 8/26
        Exit Sub 'added 8/26
      Else
        Call fpList_DblClick
        KeyCode = 0
        Exit Sub
      End If
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  'NewListFlag is set to True in frmPRDCalcScr3 and
  'is designed to return the user to whichever employee
  'he last selected...we want only the updated list
  'loaded and the focus put on the last selection
  If NewListFlag = True Then Call cmdPickList_Click
  Me.HelpContextID = hlpTransactionEntry
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub fpList_DblClick()
  Dim EmpData2FileHandle As Integer
  Dim EmpData2FileRec As EmpData2Type
  Dim NumEmpRec As Integer, x As Integer
  Dim EmployeeLastName As String
  Dim EmployeeFirstName As String
  Dim EmployeeNumber As String, Found As Boolean
  Dim PayType As String
  Dim TransRec(1) As TransRecType
  Dim TDate$
  frmLoadingPRTransForm.Show
  Call DeActivateControls
  DoEvents
  
  fpList.Col = 0
  EmployeeNumber = Val(fpList.ColText)
  fpList.Col = 2
  EmployeeFirstName = QPTrim$(fpList.ColText)
  fpList.Col = 1
  EmployeeLastName = QPTrim$(fpList.ColText)
  OpenEmpData2File EmpData2FileHandle
  
  NumEmpRec = LOF(EmpData2FileHandle) / Len(EmpData2FileRec)
  For x = 1 To NumEmpRec
    Get EmpData2FileHandle, x, EmpData2FileRec
    If InStr(UCase$(EmpData2FileRec.EmpLName), EmployeeLastName) > 0 And InStr(UCase$(EmpData2FileRec.EmpFName), EmployeeFirstName) > 0 _
    And InStr(EmpData2FileRec.EmpNo, EmployeeNumber) > 0 And Len(QPTrim$(EmployeeNumber)) = Len(QPTrim$(EmpData2FileRec.EmpNo)) Then 'added
    'this Len = Len because when an employee has #12 and another
    'employee has #125 and both names were exactly the same then
    'if you were looking for #125 the process would not get
    'past #12
      PayType = QPTrim$(EmpData2FileRec.EMPPTYPE)
      PayType = UCase$(PayType)
      If PayType = UCase$("Hourly") Then
        m_topOption = PRTtopHourly
      Else
        m_topOption = PRTtopSalary
      End If
      Found = True
      fpList.Row = -1
      RecNum = x 'OK we found it so set RecNum at this position
      Exit For 'step out because there's no need to go on
    Else
      Found = False
      GoTo NotAMatch
    End If
      
NotAMatch:
  Next x
  
  If RecNum = 0 Then
    Unload frmLoadingPRTransForm
    MsgBox "Please make a valid selection from the Employee list."
    Call ActivateControls
    Exit Sub
  End If
  
  Close EmpData2FileHandle
  
  OpenTransWorkFile TRHandle
  Get TRHandle, RecNum, TransRec(1)
  'OK we got a good number so if this one is still
  'working here and he hasn't been activated for payroll
  'transaction yet then get him started in the payroll
  'transaction process
  If TransRec(1).TActive = 0 Then 'if TActive = -1 Then
  'CreateEmpTransRecs has already been processed
    If EmpData2FileRec.EMPTDATE <> 0 Then
      TDate = MakeRegDate(Emp2Rec(1).EMPTDATE)
      MsgBox "This employee was terminated on " & TDate
      Close TRHandle
      Call ActivateControls
      Exit Sub
    Else
      If Exist("TEMPIF.DAT") Then '8/26/04
        KillFile "TEMPIF.DAT" '8/26/04
      End If '8/26/04
    
      CreateEmpTransRecs RecNum 'CreateEmpTransRecs takes
      'the data saved in this employee's file and compares
      'that data with the data that has been saved in PRDefaults
      'to determine which data will be processed and which
      'won't
      Get TRHandle, RecNum, TransRec(1) 'TransRec(1) now
      'has the file for this employee which now has
      'TransRec(1).TActive = -1
    End If
  End If
  Close TRHandle
  
  fpList.Clear
  fptxtEmpNum.Text = ""
  If m_topOption = PRTtopSalary Then
    frmSPRTThisEmp.Show
    DoEvents
    Unload frmPRTPrevEmpLookUp
    Unload frmLoadingPRTransForm
    Call ActivateControls
  ElseIf m_topOption = PRTtopHourly Then
    frmHPRTThisEmp.Show
    DoEvents
    Unload frmPRTPrevEmpLookUp
    Unload frmLoadingPRTransForm
    Call ActivateControls
  End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdEscape.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmPRTPrevEmpLookUp.")
      Call Terminate
      End
    End If
  End If
End Sub
Private Sub DeActivateControls()
  Dim cnt As Integer
  Dim x As Control
  Dim cmdButton As CommandButton

  cmdEscape.Enabled = False
  cmdProcess.Enabled = False
  cmdPickList.Enabled = False
  
  For cnt = 0 To Me.Count - 1
    Set x = Me.Controls.Item(cnt)
      If TypeOf x Is CommandButton Then
        x.Enabled = False
      End If
  Next cnt
  
  EnableCloseButton Me.hwnd, False
     
End Sub

Private Sub ActivateControls()
  Dim cmdButton As CommandButton
  Dim x As Control
  Dim cnt As Integer
  
  cmdEscape.Enabled = True
  cmdProcess.Enabled = True
  cmdPickList.Enabled = True
  
  For cnt = 0 To Me.Count - 1
    Set x = Me.Controls.Item(cnt)
      If TypeOf x Is CommandButton Then
        x.Enabled = True
      End If
  Next cnt
  EnableCloseButton Me.hwnd, True
     
End Sub



