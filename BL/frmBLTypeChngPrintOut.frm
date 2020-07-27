VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLTypeChngPrintOut 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License List Of Customers Affected By Type Change"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmBLTypeChngPrintOut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   1  'CenterOwner
   Begin LpLib.fpList fpList1 
      Height          =   2610
      Left            =   1410
      TabIndex        =   6
      Top             =   4320
      Width           =   8835
      _Version        =   196608
      _ExtentX        =   15584
      _ExtentY        =   4604
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
      ColDesigner     =   "frmBLTypeChngPrintOut.frx":08CA
   End
   Begin VB.TextBox fptxtChoice 
      Height          =   288
      Left            =   1392
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   7344
      Width           =   492
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   636
      Left            =   2028
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "Press 'Cancel' to exit this screen and return to the 'Business License Reports' menu."
      Top             =   7344
      Width           =   1884
      _Version        =   131072
      _ExtentX        =   3323
      _ExtentY        =   1122
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
      ButtonDesigner  =   "frmBLTypeChngPrintOut.frx":0C12
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   636
      Left            =   7740
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   $"frmBLTypeChngPrintOut.frx":0DF0
      Top             =   7344
      Width           =   1884
      _Version        =   131072
      _ExtentX        =   3323
      _ExtentY        =   1122
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
      ButtonDesigner  =   "frmBLTypeChngPrintOut.frx":0E9B
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrint 
      Height          =   636
      Left            =   4860
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   $"frmBLTypeChngPrintOut.frx":107B
      Top             =   7344
      Width           =   1884
      _Version        =   131072
      _ExtentX        =   3323
      _ExtentY        =   1122
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
      ButtonDesigner  =   "frmBLTypeChngPrintOut.frx":1126
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Only values applicable to the new type will be saved. All other values will be zeroed out."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   396
      Left            =   396
      TabIndex        =   8
      Top             =   3792
      Width           =   11052
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmBLTypeChngPrintOut.frx":1302
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1356
      Left            =   816
      TabIndex        =   4
      Top             =   2352
      Width           =   10236
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmBLTypeChngPrintOut.frx":149A
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1164
      Left            =   624
      TabIndex        =   3
      Top             =   1104
      Width           =   10236
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   684
      Left            =   2220
      Top             =   240
      Width           =   7212
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Customers Affected By Category Type Change"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   396
      Left            =   2448
      TabIndex        =   0
      Top             =   384
      Width           =   6780
   End
End
Attribute VB_Name = "frmBLTypeChngPrintOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class
  Dim CustArr() As Integer
  Dim CustRev() As String
  Dim Nextx As Integer
  Dim ThisCode$, ThisType$, ThisDesc$

Private Sub cmdPrint_Click()
  Dim PrintType$
  
  frmBLReportOpt.Show vbModal 'opens small screen from which the
  'user selects the printing method
  PrintType$ = frmBLReportOpt.fptxtPrintType
  Select Case PrintType$
    Case "Graphical"
      Call PrintGraphics
    Case "Text"
      frmBLMessageBoxJr.Label1.Caption = "Pitch 10 is recommended for this report."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Call PrintText
    Case "Exit"
  End Select
  Unload frmBLReportOpt

End Sub

Private Sub PrintText()
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim FF$
  Dim ReportFile$
  Dim RptHandle As Integer
  Dim x As Integer
  Dim Page As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  
  ReportFile$ = "ARTypChg.PRN"
  FF$ = Chr$(12)
  MaxLines = 58
  LineCnt = 0
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  GoSub PrintHeader
  OpenCustFile CustHandle
  
  For x = 1 To Nextx
    Get CustHandle, CustArr(x), CustRec
    Print #RptHandle, Tab(5); QPTrim$(CustRec.CustName); Tab(40); Using$("#####", CustRec.CustNumb); Tab(67); CustRev(x)
    LineCnt = LineCnt + 1
    If LineCnt > MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
  Next x
  Print #RptHandle, FF$
  Close
  
  ViewPrint ReportFile$, "Customers To Be Modified Listing", True
  KillFile ReportFile$
  MainLog ("User printed a list of all customers affected by the change in category # " + ThisCode + "'s type.")
  
  Exit Sub
  
PrintHeader:
  Page = Page + 1
  Print #RptHandle, "List of Business License Customers Needing Base Value Modifications"
  Print #RptHandle, "Report Date: "; Date$; Tab(65); "Page #"; Page
  Print #RptHandle,
  Print #RptHandle, "For Code Number/Description: " + ThisCode + "/"; ThisDesc
  Print #RptHandle, Tab(5); "Customer Name"; Tab(40); "Customer #"; Tab(63); "Current Base Value"
  Print #RptHandle, String$(80, "-")
  LineCnt = LineCnt + 6
  Return
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
      SendKeys "%C"
      Call cmdProcess_Click
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%P"
      Call cmdPrint_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Call LoadMe
End Sub

Private Sub LoadMe()
  Dim CodeRec As ARNewCatCodeRecType
  Dim CHandle As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim x As Integer
  Dim BaseValue$
  
  fptxtChoice.Visible = False
  OpenCatCodeFile CHandle
  Get CHandle, GCatNum, CodeRec
  ThisCode = QPTrim$(CodeRec.CatCode)
  ThisType = CodeRec.CodeType
  ThisDesc = QPTrim$(CodeRec.CODEDESC)
  Close CHandle
  
  OpenCustFile CustHandle
  NumOfCustRecs = LOF(CustHandle) / Len(CustRec)
  
  ReDim CustArr(1 To 1) As Integer
  ReDim CustRev(1 To 1) As String
  Nextx = 0
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
    BaseValue = "0"
    If CustRec.Deleted = "Y" Or QPTrim(CustRec.SortName) = "DELETED" Then GoTo DoneHere
    If QPTrim$(CustRec.BILLCAT1) = ThisCode Then
      If ThisType = "M" Or ThisType = "F" Then
        BaseValue = Using("#####", CustRec.REV1)
      ElseIf ThisType = "S" Then
        BaseValue = Using$("$##,###,##0.00", CustRec.REV1)
      End If
      Nextx = Nextx + 1
      ReDim Preserve CustArr(1 To Nextx) As Integer
      ReDim Preserve CustRev(1 To Nextx) As String
      CustArr(Nextx) = x
      CustRev(Nextx) = BaseValue
    ElseIf QPTrim$(CustRec.BILLCAT2) = ThisCode Then
      If ThisType = "M" Or ThisType = "F" Then
        BaseValue = Using("#####", CustRec.REV2)
      ElseIf ThisType = "S" Then
        BaseValue = Using$("$##,###,##0.00", CustRec.REV2)
      End If
      Nextx = Nextx + 1
      ReDim Preserve CustArr(1 To Nextx) As Integer
      ReDim Preserve CustRev(1 To Nextx) As String
      CustArr(Nextx) = x
      CustRev(Nextx) = BaseValue
    ElseIf QPTrim$(CustRec.BILLCAT3) = ThisCode Then
      If ThisType = "M" Or ThisType = "F" Then
        BaseValue = Using("#####", CustRec.REV3)
      ElseIf ThisType = "S" Then
        BaseValue = Using$("$##,###,##0.00", CustRec.REV3)
      End If
      Nextx = Nextx + 1
      ReDim Preserve CustArr(1 To Nextx) As Integer
      ReDim Preserve CustRev(1 To Nextx) As String
      CustArr(Nextx) = x
      CustRev(Nextx) = BaseValue
    ElseIf QPTrim$(CustRec.BILLCAT4) = ThisCode Then
      If ThisType = "M" Or ThisType = "F" Then
        BaseValue = Using("#####", CustRec.REV4)
      ElseIf ThisType = "S" Then
        BaseValue = Using$("$##,###,##0.00", CustRec.REV4)
      End If
      Nextx = Nextx + 1
      ReDim Preserve CustArr(1 To Nextx) As Integer
      ReDim Preserve CustRev(1 To Nextx) As String
      CustArr(Nextx) = x
      CustRev(Nextx) = BaseValue
    ElseIf QPTrim$(CustRec.BILLCAT5) = ThisCode Then
      If ThisType = "M" Or ThisType = "F" Then
        BaseValue = Using("#####", CustRec.REV5)
      ElseIf ThisType = "S" Then
        BaseValue = Using$("$##,###,##0.00", CustRec.REV5)
      End If
      Nextx = Nextx + 1
      ReDim Preserve CustArr(1 To Nextx) As Integer
      ReDim Preserve CustRev(1 To Nextx) As String
      CustArr(Nextx) = x
      CustRev(Nextx) = BaseValue
    End If
DoneHere:
  Next x
  
  For x = 1 To Nextx
    Get CustHandle, CustArr(x), CustRec
    fpList1.Row = x
    fpList1.InsertRow = QPTrim$(CustRec.CustName) + Chr(9) + QPTrim$(CustRec.CustNumb) + Chr(9) + CustRev(x)
  Next x
  
  Close
  
End Sub
Private Sub cmdExit_Click()
  Me.Hide
  fptxtChoice = "abort"
End Sub

Private Sub cmdProcess_Click()
  Me.Hide
  fptxtChoice = "continue"
End Sub
Private Sub PrintGraphics()
  Dim ReportFile$
  Dim RptHandle As Integer
  Dim x As Integer
  Dim Page As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim dlm$
  
  dlm$ = "~"
  ReportFile$ = "BLRPTS\CATCHNG.RPT"
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  OpenCustFile CustHandle
  
  For x = 1 To Nextx
    Get CustHandle, CustArr(x), CustRec
    '                              0                        1                    2                              3
    Print #RptHandle, QPTrim$(CustRec.CustName); dlm; CustRec.CustNumb; dlm; CustRev(x); dlm; "For Code Number/Description: " + ThisCode + "/" + ThisDesc
  Next x
  
  Close
  
  arBLCatChngCustPrntOut.Show vbModal
  
  MainLog ("User printed a (graphics) list of all customers affected by the change in category # " + ThisCode + "'s type.")
  
End Sub


