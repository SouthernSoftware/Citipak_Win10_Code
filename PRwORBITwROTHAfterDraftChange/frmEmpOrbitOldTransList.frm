VERSION 5.00
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmEmpOrbitOldTransList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ORBIT Prior Transaction List"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10605
   Icon            =   "frmEmpOrbitOldTransList.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   10605
   StartUpPosition =   1  'CenterOwner
   Begin LpLib.fpList fpListT 
      Height          =   2625
      Left            =   735
      TabIndex        =   0
      ToolTipText     =   "Double click to select."
      Top             =   3360
      Width           =   9135
      _Version        =   196608
      _ExtentX        =   16113
      _ExtentY        =   4630
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
      Columns         =   5
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
      ColDesigner     =   "frmEmpOrbitOldTransList.frx":08CA
   End
   Begin LpLib.fpList fpListE 
      Height          =   2910
      Left            =   1635
      TabIndex        =   2
      ToolTipText     =   "Double click to select"
      Top             =   360
      Width           =   7335
      _Version        =   196608
      _ExtentX        =   12938
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
      ColDesigner     =   "frmEmpOrbitOldTransList.frx":0C7B
   End
   Begin VB.CommandButton cmdEscape 
      Caption         =   "ESC &Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3975
      TabIndex        =   1
      Top             =   6600
      Width           =   2655
   End
End
Attribute VB_Name = "frmEmpOrbitOldTransList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Public ThisORec As Long
Public ThisTRec As Long

Private Sub cmdEscape_Click()
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdEscape_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub
Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Call LoadMe
End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub LoadMe()
  Dim x As Long
  Dim OERec As OrbitEmpData
  Dim OEHandle As Integer
  Dim NumOfOERecs As Integer
  
  OpenOrbEmpData OEHandle, NumOfOERecs
  For x = 1 To NumOfOERecs
    Get OEHandle, x, OERec
    If OERec.Deleted = True Then GoTo Deleted
    fpListE.InsertRow = QPTrim$(OERec.EmpNum) & Chr(9) & QPTrim$(OERec.LastName) & ", " & QPTrim$(OERec.FirstName) & Chr(9) & CStr(OERec.EmpRecNum) & Chr(9) & CStr(x)
Deleted:
  Next x
  Close OEHandle
  
End Sub

Private Sub fpListE_DblClick()
  Dim TransRec As TransRecType
  Dim x As Long
  Dim THandle As Integer
  Dim ThisRec As Long
  Dim EmpRec As EmpData2Type
  Dim EHandle As Integer
  Dim NextRec As Long
  Dim OHRec As OrbitHeader
  Dim OHandle As Integer
  Dim BegDate As Integer
  Dim EndDate As Integer
  Dim OERec As OrbitEmpData
  Dim OEHandle As Integer
  Dim NumOfOERecs As Integer
  Dim ThisSDate As String
  
  fpListT.Clear
  fpListE.Row = fpListE.SearchIndex
  fpListE.Col = 2
  ThisRec = fpListE.ColText
  fpListE.Col = 3
  ThisORec = CInt(fpListE.ColText)
  OpenOrbHeader OHandle
  Get OHandle, 1, OHRec
  Close OHandle
  BegDate = OHRec.PayPrdBeginDate
  EndDate = OHRec.PayPrdEndDate
  
  OpenEmpData2File EHandle
  Get EHandle, ThisRec, EmpRec
  Close EHandle
  NextRec = EmpRec.LastTransRec
  
  OpenTransHistFile THandle
  
  Do While NextRec > 0
    Get THandle, NextRec, TransRec
    ThisSDate = MakeRegDate(TransRec.CheckDate)
    If TransRec.CheckDate >= BegDate And TransRec.CheckDate <= EndDate Then
      GoTo NextOne
    End If
    fpListT.InsertRow = QPTrim$(Using$("#########", TransRec.CheckNum)) & Chr(9) & MakeRegDate(TransRec.PayPdStart) & Chr(9) & MakeRegDate(TransRec.PayPdEnd) & Chr(9) & Using$("$##,###.##", TransRec.RetGrossPay) & Chr(9) & CStr(NextRec)
NextOne:
    NextRec = TransRec.PrevTransRec
  Loop
  fpListT.ListIndex = 0
  Close THandle
  
End Sub

Private Sub fpListT_DblClick()
  fpListT.Row = fpListT.SearchIndex
  fpListT.Col = 4
  ThisTRec = CLng(fpListT.ColText)
  Call frmEmpORBITEdit.LoadMeEmpFromOld(ThisTRec, ThisORec)
  Unload Me
End Sub
