VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmPRVoidGLError 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Void Check GL Number Error"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   6930
   StartUpPosition =   2  'CenterScreen
   Begin LpLib.fpList fpList1 
      Height          =   2355
      Left            =   360
      TabIndex        =   0
      Top             =   2910
      Width           =   6195
      _Version        =   196608
      _ExtentX        =   10927
      _ExtentY        =   4154
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      Columns         =   0
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
      ColDesigner     =   "frmPRVoidGLError.frx":0000
   End
   Begin VB.TextBox fptxtChoice 
      Height          =   288
      Left            =   3264
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   5501
      Width           =   492
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   540
      Left            =   795
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5685
      Width           =   2280
      _Version        =   131072
      _ExtentX        =   4022
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmPRVoidGLError.frx":03A8
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrint 
      Height          =   540
      Left            =   3915
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   $"frmPRVoidGLError.frx":05BC
      Top             =   5685
      Width           =   2280
      _Version        =   131072
      _ExtentX        =   4022
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmPRVoidGLError.frx":0667
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmPRVoidGLError.frx":087B
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1950
      Left            =   510
      TabIndex        =   4
      Top             =   510
      Width           =   5910
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2490
      Left            =   360
      Top             =   240
      Width           =   6210
   End
End
Attribute VB_Name = "frmPRVoidGLError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
Private Sub cmdExit_Click()
  fptxtChoice.Text = "abort"
  Me.Hide
End Sub

Private Sub cmdPrint_Click()
  Dim RptHandle As Integer
  Dim ReportFile$
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim x As Integer
  Dim FF$
  
  FF$ = Chr$(12)

  MaxLines = 50
  If ErrCnt > 0 Then
    ReportFile$ = "PRVoidGL.PRN"
    RptHandle = FreeFile
    Open ReportFile$ For Output As #RptHandle
    GoSub PrintHeader
    For x = 1 To ErrCnt
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintHeader
      End If
      LineCnt = LineCnt + 1
      Print #RptHandle, Tab(5); QPTrim$(ErrAcct(x)); Tab(26); Using$("$##,##0.00", ErrAmt(x)); Tab(50); QPTrim$(ErrType(x))
    Next x
    Print #RptHandle, FF$
    Close RptHandle
    ViewPrint ReportFile$, "Void Check GL Error Listing", True
    KillFile ReportFile$
  
  End If
  MainLog ("User printed a list of all erroneus GL numbers for voided check # " + GlobalCheckNum + ".")
  
  Exit Sub
  
PrintHeader:
  Print #RptHandle, "Unmatched General Ledger Numbers For Check # " + GlobalCheckNum$
  Print #RptHandle, Date
  Print #RptHandle, "Employee Number: " + ErrEmpNum
  Print #RptHandle, Tab(5); "GL Number"; Tab(30); "Amount"; Tab(44); "Debit/Credit"
  Print #RptHandle, String$(55, "=")
  LineCnt = LineCnt + 5
  Return
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      Call cmdExit_Click
      SendKeys "%E"
      KeyCode = 0
    Case vbKeyF3:
      Call cmdPrint_Click
      SendKeys "%P"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Dim x As Integer
  
  fptxtChoice.Visible = False

  For x = 1 To ErrCnt
    fpList1.AddItem QPTrim(ErrAcct(x)) + "   " + Using$("$##,##0.00", ErrAmt(x)) + "   " + QPTrim$(ErrType(x))
  Next x

End Sub

