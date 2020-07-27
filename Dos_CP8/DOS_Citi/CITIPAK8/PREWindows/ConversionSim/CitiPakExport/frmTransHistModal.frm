VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmTransHistModal 
   BackColor       =   &H008F8265&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8490
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin LpLib.fpList fpList1 
      Height          =   3195
      Left            =   720
      TabIndex        =   0
      Top             =   435
      Width           =   7065
      _Version        =   196608
      _ExtentX        =   12462
      _ExtentY        =   5636
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
      ColDesigner     =   "frmTransHistModal.frx":0000
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdOK 
      Height          =   444
      Left            =   2658
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4005
      Width           =   3180
      _Version        =   131072
      _ExtentX        =   5609
      _ExtentY        =   783
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
      ButtonDesigner  =   "frmTransHistModal.frx":04B5
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   4524
      Left            =   210
      Top             =   240
      Width           =   8076
   End
End
Attribute VB_Name = "frmTransHistModal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LoadMe()
  Dim TransRec(1) As TransRecType
  Dim THandle As Integer
  Dim TransCnt As Long
  Dim Emp2Rec As EmpData2Type
  Dim Emp2Handle As Integer
  Dim x As Long
  Dim LastEmpTransRec As Long
  Dim NextTransRec As Long
  Dim CheckNum As Long
  Dim CheckDate As Integer
  Dim PeriodEnd As Integer
  Dim NetPay As Double
  
  Call FixRes
  
  If RecNum = 0 Then
    MsgBox "No records on file for this employee"
    Exit Sub
  End If
  'RecNum is a global that is set in the Employee Lookup
  'screen before the Employee Maintenance screen is loaded
  OpenEmpData2File Emp2Handle
  Get Emp2Handle, RecNum, Emp2Rec
  Close Emp2Handle
  
  LastEmpTransRec = Emp2Rec.LastTransRec
  If LastEmpTransRec > 0 Then
    NextTransRec = LastEmpTransRec
  Else
    MsgBox "No transaction records on file for this employee."
    Exit Sub
  End If
  
  OpenTransHistFile THandle
  TransCnt = LOF(THandle) / Len(TransRec(1))
  For x = 1 To TransCnt
    Get THandle, NextTransRec, TransRec(1)
    CheckNum = TransRec(1).CheckNum
    CheckDate = TransRec(1).CheckDate
    PeriodEnd = TransRec(1).PayPdEnd
    NetPay = TransRec(1).NetPay
    fpList1.InsertRow = "    " & CheckNum & Chr(9) & "      " & MakeRegDate(CheckDate) & Chr(9) & "   " & MakeRegDate(PeriodEnd) & Chr(9) & Using$("$###,##0.00", NetPay)
    NextTransRec = TransRec(1).PrevTransRec
    If NextTransRec = 0 Then GoTo AllDone
  Next x
AllDone:
  Close THandle

End Sub

Private Sub cmdOk_Click()
  Unload frmTransHistModal
  DoEvents
 
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      SendKeys "%E"
      Unload frmTransHistModal
      KeyCode = 0
  End Select
End Sub

Private Sub Form_Load()
  Call LoadMe
End Sub
Private Sub FixRes()

    Select Case ScreenW
      Case 1280
      If Screen.TwipsPerPixelX <> 12 Then
         fpList1.FontSize = 14
      Else
         fpList1.FontSize = 12
      End If
      Case 1152
      If Screen.TwipsPerPixelX <> 12 Then
         fpList1.FontSize = 12
      Else
         fpList1.FontSize = 12
      End If
      Case 1024
      If Screen.TwipsPerPixelX <> 12 Then
         fpList1.FontSize = 12
      Else
         fpList1.FontSize = 12
      End If
      Case 800
         fpList1.FontSize = 12
      Case Else
       
    End Select

End Sub


