VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmVATaxPPTRAMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Billing PPTRA Maintenance Menu"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "frmVATaxPPTRAMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdPost 
      Height          =   444
      Left            =   4020
      TabIndex        =   0
      Top             =   5148
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   783
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
      ButtonDesigner  =   "frmVATaxPPTRAMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintRpt 
      Height          =   432
      Left            =   4020
      TabIndex        =   1
      Top             =   4512
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   762
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
      ButtonDesigner  =   "frmVATaxPPTRAMenu.frx":0AB0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdReverse 
      Height          =   432
      Left            =   4020
      TabIndex        =   2
      Top             =   3876
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   762
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
      ButtonDesigner  =   "frmVATaxPPTRAMenu.frx":0C98
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   432
      Left            =   4020
      TabIndex        =   3
      Top             =   5832
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   762
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
      ButtonDesigner  =   "frmVATaxPPTRAMenu.frx":0E82
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1104
      Index           =   1
      Left            =   1500
      Top             =   840
      Width           =   8652
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   132
      Left            =   2100
      Top             =   2052
      Width           =   972
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8712
      X2              =   8712
      Y1              =   2160
      Y2              =   8061
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   132
      Left            =   8592
      Top             =   2052
      Width           =   972
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8712
      X2              =   9414
      Y1              =   8052
      Y2              =   8052
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2208
      X2              =   2923
      Y1              =   8052
      Y2              =   8052
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2220
      X2              =   2220
      Y1              =   2160
      Y2              =   8048
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PPTRA MAINTENANCE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2820
      TabIndex        =   4
      Top             =   1200
      Width           =   6012
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1500
      Top             =   720
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   3
      Left            =   2100
      Top             =   1920
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5916
      Index           =   0
      Left            =   2220
      Top             =   2148
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8592
      Top             =   1920
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5916
      Index           =   1
      Left            =   8712
      Top             =   2148
      Width           =   732
   End
End
Attribute VB_Name = "frmVATaxPPTRAMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim TownName$
Private Sub cmdExit_Click()
  frmVATaxMainMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPost_Click()
  If Not Exist(PPTRARemovalFile) Then
    Call TaxMsg(900, "No PPTRA Removal Files have been created. Posting access denied.")
    Exit Sub
  End If
  frmVATaxPostPPTRARmvl.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPrintRpt_Click()
  If Not Exist(PPTRARemovalFile) Then
    Call TaxMsg(900, "No PPTRA Removal Files have been created. Report access denied.")
    Exit Sub
  End If
  
  frmVATaxReportOpt.Show vbModal
  If frmVATaxReportOpt.fptxtPrintType.Text = "Graphical" Then
    Unload frmVATaxReportOpt
    Call PrintGraphics
  ElseIf frmVATaxReportOpt.fptxtPrintType.Text = "Text" Then
'    frmVATaxMsg.Label1.Caption = "Pitch 10 is recommended for this report."
'    frmVATaxMsg.Label1.Top = 900
'    frmVATaxMsg.Show vbModal
    Unload frmVATaxReportOpt
    Call PrintText
  End If

End Sub

Private Sub cmdReverse_Click()
  frmVATaxPPTRARemoval.Show
  DoEvents
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  TownName = QPTrim$(TaxMasterRec.Name)
  
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpPPTRARemovalMenu
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxPPTRAMenu.")
      Call Terminate
      End
    End If
  End If

End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub PrintGraphics()
  Dim x As Long
  Dim RmvlRec As TaxPPTRARemovalType
  Dim RHandle As Integer
  Dim NumOfRmvlRecs As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim ThisName$
  Dim ThatName$
  Dim NewTotal As Double
  Dim SubNewTotal As Double
  Dim GNewTotal As Double
  Dim SubOldTotal As Double
  Dim GOldTotal As Double
  Dim SubPPTRAAmt As Double
  Dim GPPTRAAmt As Double
  Dim SubCnt As Long
  Dim GCnt As Long
  Dim dlm$, PostDate$
  dlm$ = "~"
  RptFile$ = "TAXRPTS\RMVBYBILL.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  ThisName = ""
  
  SubNewTotal = 0
  SubOldTotal = 0
  SubPPTRAAmt = 0
  SubCnt = 0
  GNewTotal = 0
  GOldTotal = 0
  GPPTRAAmt = 0
  GCnt = 0
  OpenPPTRARmvlFile RHandle, NumOfRmvlRecs
  For x = 1 To NumOfRmvlRecs
    Get RHandle, x, RmvlRec
    ThisName = QPTrim$(RmvlRec.RmvlFile)
    If ThisName <> ThatName Then
      ThatName = ThisName
      SubNewTotal = 0
      SubOldTotal = 0
      SubPPTRAAmt = 0
      SubCnt = 0
    End If
    PostDate = MakeRegDate(RmvlRec.BillDate)
    NewTotal = OldRound(RmvlRec.TaxAmount + RmvlRec.PPTRADisc)
    SubNewTotal = OldRound(SubNewTotal + NewTotal)
    SubOldTotal = OldRound(SubOldTotal + RmvlRec.TaxAmount)
    SubPPTRAAmt = OldRound(SubPPTRAAmt + RmvlRec.PPTRADisc)
    SubCnt = SubCnt + 1
    GNewTotal = OldRound(GNewTotal + NewTotal)
    GOldTotal = OldRound(GOldTotal + RmvlRec.TaxAmount)
    GPPTRAAmt = OldRound(GPPTRAAmt + RmvlRec.PPTRADisc)
    GCnt = GCnt + 1
    GoSub PrintIt
  Next x
  
  Close
  
  arVATaxPPTRARmvlByBill.Show
  
  Exit Sub

PrintIt:
    '                    0                    1                             2
    Print #RptHandle, TownName; dlm; QPTrim$(RmvlRec.RmvlFile); dlm; RmvlRec.CustAcct; dlm;
    '                            3                           4
    Print #RptHandle, QPTrim$(RmvlRec.CustName); dlm; RmvlRec.BillNum; dlm;
    '                         5                      6                   7
    Print #RptHandle, RmvlRec.PPTRADisc; dlm; RmvlRec.TaxAmount; dlm; NewTotal; dlm;
    '                     8                  9                10
    Print #RptHandle, SubNewTotal; dlm; SubOldTotal; dlm; SubPPTRAAmt; dlm;
    '                   11             12             13              14            15
    Print #RptHandle, SubCnt; dlm; GNewTotal; dlm; GOldTotal; dlm; GPPTRAAmt; dlm; GCnt; dlm;
    '                    16
    Print #RptHandle, PostDate

  Return

End Sub

Private Sub PrintText()
  Dim x As Long
  Dim RmvlRec As TaxPPTRARemovalType
  Dim RHandle As Integer
  Dim NumOfRmvlRecs As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim MaxLines As Integer
  Dim FF$
  Dim LineCnt As Integer
  Dim Page As Integer
  Dim ThisName$
  Dim ThatName$
  Dim NewTotal As Double
  Dim SubNewTotal As Double
  Dim GNewTotal As Double
  Dim SubOldTotal As Double
  Dim GOldTotal As Double
  Dim SubPPTRAAmt As Double
  Dim GPPTRAAmt As Double
  Dim SubCnt As Long
  Dim GCnt As Long
  Dim PostDate$
  
  MaxLines = 58
  FF$ = Chr(12)
  
  RptFile$ = "TAXRPTS\RMVBYBILL.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  ThisName = ""
  GoSub PrintHeader
  SubNewTotal = 0
  SubOldTotal = 0
  SubPPTRAAmt = 0
  SubCnt = 0
  GNewTotal = 0
  GOldTotal = 0
  GPPTRAAmt = 0
  GCnt = 0
  OpenPPTRARmvlFile RHandle, NumOfRmvlRecs
  For x = 1 To NumOfRmvlRecs
    Get RHandle, x, RmvlRec
    ThisName = QPTrim$(RmvlRec.RmvlFile)
    If ThisName <> ThatName Then
      LineCnt = LineCnt + 1
      If LineCnt > MaxLines - 8 Then
        Print #RptHandle, FF$
        GoSub PrintHeader
      End If
      If SubCnt > 0 Then GoSub PrintSummary
      ThatName = ThisName
      SubNewTotal = 0
      SubOldTotal = 0
      SubPPTRAAmt = 0
      SubCnt = 0
      If LineCnt > MaxLines - 4 Then
        Print #RptHandle, FF$
        GoSub PrintHeader
      End If
      GoSub PrintSubHeader
    End If
    PostDate = MakeRegDate(RmvlRec.BillDate)
    NewTotal = OldRound(RmvlRec.TaxAmount + RmvlRec.PPTRADisc)
    SubNewTotal = OldRound(SubNewTotal + NewTotal)
    SubOldTotal = OldRound(SubOldTotal + RmvlRec.TaxAmount)
    SubPPTRAAmt = OldRound(SubPPTRAAmt + RmvlRec.PPTRADisc)
    SubCnt = SubCnt + 1
    GNewTotal = OldRound(GNewTotal + NewTotal)
    GOldTotal = OldRound(GOldTotal + RmvlRec.TaxAmount)
    GPPTRAAmt = OldRound(GPPTRAAmt + RmvlRec.PPTRADisc)
    GCnt = GCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
      GoSub PrintSubHeader
    End If
    Print #RptHandle, Using$("#####0", RmvlRec.CustAcct); Tab(12); QPTrim$(RmvlRec.CustName); Tab(53); Using("######", RmvlRec.BillNum); Tab(65); Using$("$##,##0.00", RmvlRec.PPTRADisc); Tab(77); Using$("$##,##0.00", RmvlRec.TaxAmount); Tab(93); Using$("$##,##0.00", NewTotal)
    LineCnt = LineCnt + 1
  Next x
  
  If LineCnt > MaxLines - 9 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
    GoSub PrintSubHeader
  End If
  GoSub PrintSummary
  
  If LineCnt > MaxLines - 7 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  
  GoSub PrintTotals
  Close
  
  ViewPrint RptFile$, "Printing PPTRA Removal Report", True
  
  Exit Sub
  
PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(40); "PPTRA Removal Report"
  Print #RptHandle, "Date: "; CStr(Date); Tab(83); "Page #"; CStr(Page)
  Print #RptHandle,
  Print #RptHandle, "Cust #"; Tab(12); "Customer Name"; Tab(53); "Bill #"; Tab(63); "PPTRA Amount"; Tab(77); "Bill Total"; Tab(89); "New Bill Total"
  Print #RptHandle, String$(102, "-")
  LineCnt = 5
  
  Return
  
PrintSubHeader:
  Print #RptHandle, "FileName: " + ThisName; Tab(60); "Billing Date: " + PostDate
  Print #RptHandle, String(102, "-")
  LineCnt = LineCnt + 3
  
  Return
  
PrintSummary:
  Print #RptHandle, "Summary for FileName: " + ThatName
  Print #RptHandle, String(102, "-")
  Print #RptHandle, "Transaction Count: " + Using$("##,##0", SubCnt)
  Print #RptHandle, "Total PPTRA Discounts: " + Using$("$###,###,##0.00", SubPPTRAAmt)
  Print #RptHandle, "Bill Totals:           " + Using$("$###,###,##0.00", SubOldTotal)
  Print #RptHandle, "New Bill Totals:       " + Using$("$###,###,##0.00", SubNewTotal)
  Print #RptHandle, String(102, "-")
  Print #RptHandle,
  Print #RptHandle,
  LineCnt = LineCnt + 8

  Return
  
PrintTotals:
  Print #RptHandle,
  Print #RptHandle, String$(102, "-")
  Print #RptHandle, "Grand Totals"
  Print #RptHandle, "Grand Transaction Count: " + Using$("##,##0", GCnt)
  Print #RptHandle, "Grand Total PPTRA Discounts: " + Using$("$###,###,##0.00", GPPTRAAmt)
  Print #RptHandle, "Bill Grand Totals:           " + Using$("$###,###,##0.00", GOldTotal)
  Print #RptHandle, "New Bill Grand Totals:       " + Using$("$###,###,##0.00", GNewTotal)
  Print #RptHandle, String(102, "-")
  LineCnt = LineCnt + 8

  Return

End Sub



