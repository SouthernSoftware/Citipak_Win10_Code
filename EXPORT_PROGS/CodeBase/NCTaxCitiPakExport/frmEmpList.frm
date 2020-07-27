VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmEmpList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee List"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7920
   Icon            =   "frmEmpList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   7920
   StartUpPosition =   1  'CenterOwner
   Begin LpLib.fpList fpList1 
      Height          =   3840
      Left            =   450
      TabIndex        =   0
      ToolTipText     =   "Double click a selection to populate the appropriate field with your selection."
      Top             =   1140
      Width           =   6780
      _Version        =   196608
      _ExtentX        =   11959
      _ExtentY        =   6773
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      ColDesigner     =   "frmEmpList.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClose 
      Height          =   495
      Left            =   2775
      TabIndex        =   1
      Top             =   5400
      Width           =   2370
      _Version        =   131072
      _ExtentX        =   4180
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
      ButtonDesigner  =   "frmEmpList.frx":0BDA
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   6684
      Left            =   0
      Top             =   0
      Width           =   7644
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee List"
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
      Left            =   1920
      TabIndex        =   2
      Top             =   486
      Width           =   3900
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   636
      Left            =   1794
      Top             =   336
      Width           =   4044
   End
End
Attribute VB_Name = "frmEmpList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Dim NumOfEmpRecs As Integer
  Private Temp_Class As Resize_Class

Private Sub cmdClose_Click()
  
  Unload Me
End Sub

'Private Sub cmdHelp_Click()
'  If InStr(cmdHelp.Text, "On") Then
'    cmdHelp.Text = "F1 &Turn Help Off"
'    Label2.Visible = True
'    Line1.Visible = True
'  Else
'    cmdHelp.Text = "F1 &Turn Help On"
'    Label2.Visible = False
'    Line1.Visible = False
'  End If
'End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call LoadMe
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyReturn
      Call fpList1_DblClick
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      Call cmdClose_Click
      SendKeys "%C"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub LoadMe()
  Dim EmpRec As EmpData2Type
  Dim EHandle As Integer
  Dim IdxRec As NameSortIdxType
  Dim XHandle As Integer
  Dim x As Integer
  
  On Error GoTo ERRORSTUFF
  
  OpenEmpIdxLNameFile XHandle
  NumOfEmpRecs = LOF(XHandle) \ 2
  
  If NumOfEmpRecs = 0 Then 'file is there but there is nothing in it
    MsgBox "No employee index built. No employee list available."
    Close
    Exit Sub
  End If
   
  ReDim EmpIdx(1 To NumOfEmpRecs) As Integer
  For x = 1 To NumOfEmpRecs
    Get XHandle, x, IdxRec.DataRecNum
    EmpIdx(x) = IdxRec.DataRecNum
  Next x
  Close XHandle
  
  If Exist(PRData + EmpData2Name) Then
    OpenEmpData2File EHandle
  Else
    MsgBox "No employee records have been saved."
    Close
    Exit Sub
  End If
  
  For x = 1 To NumOfEmpRecs
    Get EHandle, EmpIdx(x), EmpRec
    If EmpRec.Deleted = -1 Then GoTo BadEmp
    fpList1.InsertRow = QPTrim$(EmpRec.EmpLName) & ", " & QPTrim$(EmpRec.EmpFName) & Chr$(9) & QPTrim$(EmpRec.EmpNo) & Chr(9) & CStr(EmpIdx(x))
BadEmp:
  Next x
  Close EHandle
  fpList1.Row = 0
  fpList1.Selected = True 'set focus to first line
ZeroText:
  Exit Sub
   

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmEmpList", "LoadMe", Erl)
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
  Dim EmpRec As EmpData2Type
  Dim EHandle As Integer
  Dim NumOfEmpRecs As Integer
  Dim IdxRec As NameSortIdxType
  Dim XHandle As Integer
  Dim x As Integer
  Dim Number$
  Dim Name$
  Dim Found As Boolean
  Dim ThisRow  As Integer
  On Error GoTo ERRORSTUFF
   
  ThisRow = fpList1.ListIndex
  If Exist("payraterpt.dat") Then
    If ThisRow = 0 Then
'      frmPayRateRpt.fptxtEmpName.Text = "ALL"
      GEmpNum = 0
    Else
'      frmPayRateRpt.fptxtEmpName.Text = QPTrim$(fpList1.ColText)
      fpList1.Col = 2
      GEmpNum = CInt(fpList1.ColText)
    End If
  End If
  Unload Me
  
  Exit Sub

ERRORSTUFF:
   Unload FrmShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmEmpList", "fpList1_DblClick", Erl)
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
  



