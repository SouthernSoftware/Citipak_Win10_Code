VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmRptWrkOrdHist 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Work Order History"
   ClientHeight    =   4185
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   8760
   ControlBox      =   0   'False
   Icon            =   "frmRptWrkOrdHist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin LpLib.fpList fpWRKList 
      Height          =   2040
      Left            =   480
      TabIndex        =   0
      Top             =   885
      Width           =   7785
      _Version        =   196608
      _ExtentX        =   13732
      _ExtentY        =   3598
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Columns         =   2
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
      SelMax          =   1
      AutoSearch      =   1
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   0
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
      ScrollBarV      =   0
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   2
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   0   'False
      DataAutoSizeCols=   0
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   3
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
      ColDesigner     =   "frmRptWrkOrdHist.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdOK 
      Height          =   480
      Left            =   5736
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3528
      Width           =   1284
      _Version        =   131072
      _ExtentX        =   2265
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmRptWrkOrdHist.frx":0BAE
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdExit 
      Height          =   480
      Left            =   7230
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3510
      Width           =   1275
      _Version        =   131072
      _ExtentX        =   2249
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmRptWrkOrdHist.frx":0D87
   End
   Begin VB.Label LabelTot 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
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
      Height          =   324
      Left            =   1920
      TabIndex        =   10
      Top             =   3696
      Width           =   1044
   End
   Begin VB.Label Label42 
      BackStyle       =   0  'Transparent
      Caption         =   "Completed"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   6552
      TabIndex        =   9
      Top             =   408
      Width           =   1356
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   6864
      TabIndex        =   8
      Top             =   624
      Width           =   876
   End
   Begin VB.Label Labe56 
      BackStyle       =   0  'Transparent
      Caption         =   "Complete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   4752
      TabIndex        =   7
      Top             =   408
      Width           =   1116
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   3072
      TabIndex        =   6
      Top             =   624
      Width           =   804
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   1056
      TabIndex        =   5
      Top             =   624
      Width           =   1284
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Transactions:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   432
      TabIndex        =   4
      Top             =   3696
      Width           =   1980
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Order"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   912
      TabIndex        =   3
      Top             =   408
      Width           =   1260
   End
   Begin VB.Label Labe54 
      BackStyle       =   0  'Transparent
      Caption         =   "Entry"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   3072
      TabIndex        =   2
      Top             =   408
      Width           =   828
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "By Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   4872
      TabIndex        =   1
      Top             =   624
      Width           =   972
   End
End
Attribute VB_Name = "frmRptWrkOrdHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Custlook As Long
Dim WOlook As Long
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
    KeyCode = 0
    Call fpCmdExit_Click
  End If
  If KeyCode = vbKeyReturn Then
    KeyCode = 0
    DoEvents
    Call fpWRKList_DblClick  'fpCmdOK_Click
  End If
  If KeyCode = vbKeyTab Then
    KeyCode = 0
    DoEvents
    Call fpCmdExit_Click
  End If
End Sub

Private Sub fpCmdExit_Click()
  DoEvents
  Unload frmRptWrkOrdHist
End Sub

Private Sub fpCmdExit_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Button = 0
  Call fpCmdExit_Click
End Sub

Public Sub ShowWrkOrdHistory(CustRec&)
  Dim WorkOrderRecLen As Integer, dcnt As Integer
  Dim UBCustRecLen As Integer, UBFile As Integer
  Dim UBWrkOrd As Integer, PrevTranRec As Long
  Dim Build As String * 80
  ReDim WorkOrderRec(1) As WorkOrderRecType
  WorkOrderRecLen = Len(WorkOrderRec(1))
  Custlook = CustRec
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  UBFile = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
  Get UBFile, CustRec&, UBCustRec(1)
  Close UBFile
  UBWrkOrd = FreeFile
  Open UBPath$ + "UBWRKORD.DAT" For Random Shared As UBWrkOrd Len = WorkOrderRecLen

  PrevTranRec& = UBCustRec(1).WOLastTrans

  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      dcnt = dcnt + 1
      Build$ = ""
      'ReDim Preserve MChoice(1 To DCnt) As FLen2
      Get UBWrkOrd, PrevTranRec&, WorkOrderRec(1)
      LSet Build$ = "  " + Str$(PrevTranRec&)
      Mid$(Build$, 20) = Num2Date(WorkOrderRec(1).ENTRYDATE)
      Mid$(Build$, 35) = Num2Date(WorkOrderRec(1).CompleteByDate)
      If WorkOrderRec(1).CompletedDate <= 0 Then
        Mid$(Build$, 50) = "Open"
      Else
        Mid$(Build$, 50) = Num2Date(WorkOrderRec(1).CompletedDate)
      End If
      Mid$(Build$, 71) = Chr$(9) + Str$(PrevTranRec&)
      If Len(QPTrim(Build$)) > 0 Then
        frmRptWrkOrdHist.fpWRKList.AddItem Build$
      End If
      PrevTranRec& = WorkOrderRec(1).PrevTransRec
    Loop
    Close UBWrkOrd
    LabelTot.Caption = Str$(dcnt)
    Me.Show 1

WOTop:

  Else
    Close UBWrkOrd
    MsgBox "No WorkOrder Transactions", vbOKOnly, "No Transactions"
  End If

  
  Erase UBCustRec, WorkOrderRec

  Exit Sub

'WOShowDetail:

End Sub

Private Sub fpCmdOk_Click()
  If fpWRKList.SelCount > 0 Then
    Call fpWRKList_DblClick
  End If
End Sub

Private Sub fpWRKList_DblClick()
  Dim TDate As String, cnt As Integer, TransRecNum As Long
  Dim UBWrkOrd As Integer, WorkOrderRecLen As Integer
  ReDim WorkOrderRec(1) As WorkOrderRecType
  WorkOrderRecLen = Len(WorkOrderRec(1))
  If fpWRKList.SelCount <> 0 Then
  fpWRKList.col = 1     'switch to the hidden RecNo. column
  TransRecNum& = Val(fpWRKList.ColText)     'get customer recno
  WOlook = TransRecNum&
  UBWrkOrd = FreeFile
  Open UBPath$ + "UBWRKORD.DAT" For Random Shared As UBWrkOrd Len = WorkOrderRecLen
  If TransRecNum& > 0 Then
    Get UBWrkOrd, TransRecNum&, WorkOrderRec(1)
  End If
  Close UBWrkOrd
'
'
  frmRptWrkOrdDetail.LabelWONum.Caption = Str$(TransRecNum&)
  If WorkOrderRec(1).CompletedDate <= 0 Then
    TDate$ = "Open"
  Else
    TDate$ = Num2Date$(WorkOrderRec(1).CompletedDate)
  End If
  frmRptWrkOrdDetail.LabelEntryDate.Caption = Num2Date$(WorkOrderRec(1).ENTRYDATE)
  frmRptWrkOrdDetail.LabelCompDate.Caption = TDate$
  frmRptWrkOrdDetail.LabelCompBy.Caption = Num2Date$(WorkOrderRec(1).CompleteByDate)
  frmRptWrkOrdDetail.LabelI1.Caption = QPTrim(WorkOrderRec(1).OrdersText.Text(1))
  frmRptWrkOrdDetail.LabelI2.Caption = QPTrim(WorkOrderRec(1).OrdersText.Text(2))
  frmRptWrkOrdDetail.LabelI3.Caption = QPTrim(WorkOrderRec(1).OrdersText.Text(3))
  frmRptWrkOrdDetail.LabelI4.Caption = QPTrim(WorkOrderRec(1).OrdersText.Text(4))
  frmRptWrkOrdDetail.LabelI5.Caption = QPTrim(WorkOrderRec(1).OrdersText.Text(5))
  frmRptWrkOrdDetail.LabelI6.Caption = QPTrim(WorkOrderRec(1).OrdersText.Text(6))
  frmRptWrkOrdDetail.LabelR1.Caption = QPTrim(WorkOrderRec(1).RepliesText.Text(1))
  frmRptWrkOrdDetail.LabelR2.Caption = QPTrim(WorkOrderRec(1).RepliesText.Text(2))
  frmRptWrkOrdDetail.LabelR3.Caption = QPTrim(WorkOrderRec(1).RepliesText.Text(3))
  frmRptWrkOrdDetail.LabelR4.Caption = QPTrim(WorkOrderRec(1).RepliesText.Text(4))
  frmRptWrkOrdDetail.LabelR5.Caption = QPTrim(WorkOrderRec(1).RepliesText.Text(5))
  frmRptWrkOrdDetail.LabelR6.Caption = QPTrim(WorkOrderRec(1).RepliesText.Text(6))
  frmRptWrkOrdDetail.Show 1
End If
End Sub


Private Sub fpWRKList_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      KeyCode = 0
      Call fpCmdExit_Click
    Case vbKeyReturn
      KeyCode = 0
      DoEvents
      Call fpWRKList_DblClick  'fpCmdOK_Click
    Case Else:
  End Select
    
End Sub
Public Sub PrintWO()
  Unload frmRptWrkOrdDetail
  PrnOneWO Custlook&, WOlook&
  
End Sub
