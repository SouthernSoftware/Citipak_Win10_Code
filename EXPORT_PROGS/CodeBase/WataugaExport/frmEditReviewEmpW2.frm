VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmEditReviewEmpW2 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit/Review Employee W2s"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmEditReviewEmpW2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8600
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpList1 
      Height          =   5295
      Left            =   1830
      TabIndex        =   1
      Top             =   2010
      Width           =   8160
      _Version        =   196608
      _ExtentX        =   14393
      _ExtentY        =   9340
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
      BorderColor     =   65535
      BorderWidth     =   2
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   1
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
      ColDesigner     =   "frmEditReviewEmpW2.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   525
      Left            =   8640
      TabIndex        =   3
      Top             =   7590
      Width           =   1335
      _Version        =   131072
      _ExtentX        =   2355
      _ExtentY        =   926
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
      DrawFocusRect   =   4
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
      ButtonDesigner  =   "frmEditReviewEmpW2.frx":0C62
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   525
      Left            =   6390
      TabIndex        =   2
      Top             =   7590
      Width           =   2010
      _Version        =   131072
      _ExtentX        =   3545
      _ExtentY        =   926
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
      DrawFocusRect   =   4
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
      ButtonDesigner  =   "frmEditReviewEmpW2.frx":0E3E
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "List of Employees"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2808
      TabIndex        =   0
      Top             =   696
      Width           =   6012
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   972
      Index           =   1
      Left            =   1512
      Top             =   456
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1092
      Left            =   1512
      Top             =   336
      Width           =   8652
   End
End
Attribute VB_Name = "frmEditReviewEmpW2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
   frmW2Processing.Show
   DoEvents
   Unload frmEditReviewEmpW2
End Sub

Private Sub LoadNames()
  Dim EmpIdxLNameHandle As Integer
  Dim EmpIdxRec As NameSortIdxType
  Dim LNameIdx() As Integer, IdxRecLen As Integer
  Dim RecNum As Long, x As Integer
  Dim Emp2Handle As Integer
  Dim Emp2Rec As EmpData2Type
  Dim LName As String * 20
  Dim FName As String * 20
  Dim ENo As String
  IdxRecLen = 2
  OpenEmpIdxLNameFile EmpIdxLNameHandle
  RecNum = LOF(EmpIdxLNameHandle) \ IdxRecLen
  If RecNum = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If
  ReDim LNameIdx(RecNum)
  For x = 1 To RecNum
     Get EmpIdxLNameHandle, x, LNameIdx(x)
  Next x
  Close EmpIdxLNameHandle
  'this procedure creates an alphabetical index by last name
  OpenEmpData2File Emp2Handle
  RecNum = LOF(Emp2Handle) \ Len(Emp2Rec)
  
  For x = 1 To RecNum
     Get Emp2Handle, LNameIdx(x), Emp2Rec
'     If Emp2Rec.EMPTDATE > 0 Then GoTo NotThisOne
     If Emp2Rec.Deleted = -1 Then GoTo NotThisOne
     fpList1.Row = -1
     LName = Emp2Rec.EmpLName
     FName = Emp2Rec.EmpFName
     ENo = Emp2Rec.EmpNo
     fpList1.InsertRow = "    " & QPTrim$(Emp2Rec.EmpNo) & Chr$(9) & "  " & QPTrim$(Emp2Rec.EmpLName) & Chr$(9) & "   " & Emp2Rec.EmpFName
NotThisOne:
  Next x
  Close Emp2Handle
  fpList1.ListIndex = 0 'added 5/28/2004
End Sub

Private Sub cmdProcess_Click()
  Call fpList1_DblClick
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
   Case vbKeyReturn:
      Call fpList1_DblClick
   Case vbKeyDown:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Call LoadNames
  Me.HelpContextID = hlpEditReview
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub fpList1_DblClick()
  Dim EmpData2FileHandle As Integer
  Dim EmpData2FileRec As EmpData2Type
  Dim EmpData1FileHandle As Integer
  Dim EmpData1FileRec As EmpData1Type
  Dim NumEmpRec As Integer, x As Integer
  Dim EmployeeLastName As String
  Dim EmployeeFirstName As String
  Dim EmployeeNumber As String, Found As Boolean
  
  RecNum = 0
  fpList1.Col = 1
  EmployeeLastName = QPTrim$(fpList1.ColText)
  If Len(QPTrim$(EmployeeLastName)) = 0 Then
    MsgBox "Please select an employee to process"
    Exit Sub
  End If
  fpList1.Col = 2
  EmployeeFirstName = QPTrim$(fpList1.ColText)
  fpList1.Col = 0
  EmployeeNumber = QPTrim$(fpList1.ColText)
  OpenEmpData2File EmpData2FileHandle
  OpenEmpData1File EmpData1FileHandle
  Call DeActivateControls
  
  NumEmpRec = LOF(EmpData2FileHandle) / Len(EmpData2FileRec)
  For x = 1 To NumEmpRec
     Get EmpData2FileHandle, x, EmpData2FileRec
     Get EmpData1FileHandle, x, EmpData1FileRec
       If InStr(UCase$(EmpData2FileRec.EmpLName), EmployeeLastName) > 0 And InStr(UCase$(EmpData2FileRec.EmpFName), EmployeeFirstName) > 0 And InStr(EmpData2FileRec.EmpNo, EmployeeNumber) > 0 Then
         Found = True
         fpList1.Row = -1
         RecNum = x
         Exit For
       Else
         Found = False
         GoTo NotAMatch
       End If
      
NotAMatch:
  Next x
'  fpList1.Clear
  Close EmpData2FileHandle
  Close EmpData1FileHandle
  Close

  frmW2EmpInfo.Show
  DoEvents
  frmEditReviewEmpW2.Hide
  
'  Unload frmEditReviewEmpW2
  Call ActivateControls
  MainLog ("W2 employee data processed.")
End Sub

Private Sub fpList1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then Call fpList1_DblClick
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      Call UnloadAllFormsAndOpn(RegExit)
      MainLog ("Payroll.exe terminated via menu bar on frmEditReviewEmpW2.")
      End
    End If
  End If
End Sub

Private Sub DeActivateControls()
  Dim cnt As Integer
  Dim x As Control
  Dim cmdButton As CommandButton

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
  
  For cnt = 0 To Me.Count - 1
    Set x = Me.Controls.Item(cnt)
      If TypeOf x Is CommandButton Then
        x.Enabled = True
      End If
  Next cnt
  EnableCloseButton Me.hwnd, True
     
End Sub

