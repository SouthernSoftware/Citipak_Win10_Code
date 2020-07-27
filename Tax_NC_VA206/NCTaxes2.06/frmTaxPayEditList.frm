VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmTaxPayEditList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Edit Transaction List"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTaxPayEditList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpListPay 
      Height          =   3912
      Left            =   780
      TabIndex        =   0
      Top             =   2460
      Width           =   10092
      _Version        =   196608
      _ExtentX        =   17801
      _ExtentY        =   6900
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
      Object.TabStop         =   -1  'True
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
      ColDesigner     =   "frmTaxPayEditList.frx":08CA
   End
   Begin EditLib.fpText fptxtOperator 
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1620
      Width           =   2895
      _Version        =   196608
      _ExtentX        =   5106
      _ExtentY        =   661
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
      ThreeDInsideHighlightColor=   -2147483633
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
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
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
      ControlType     =   1
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   540
      Left            =   3656
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7095
      Width           =   2055
      _Version        =   131072
      _ExtentX        =   3625
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
      ButtonDesigner  =   "frmTaxPayEditList.frx":0D22
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEdit 
      Height          =   540
      Left            =   5929
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7095
      Width           =   2055
      _Version        =   131072
      _ExtentX        =   3625
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
      ButtonDesigner  =   "frmTaxPayEditList.frx":0F00
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrint 
      Height          =   540
      Left            =   4140
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7800
      Width           =   3375
      _Version        =   131072
      _ExtentX        =   5953
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
      ButtonDesigner  =   "frmTaxPayEditList.frx":10DC
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   4455
      Left            =   600
      Top             =   2295
      Width           =   10455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Payment List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3135
      TabIndex        =   4
      Top             =   885
      Width           =   5295
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1485
      Top             =   735
      Width           =   8655
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1485
      Top             =   615
      Width           =   8655
   End
End
Attribute VB_Name = "frmTaxPayEditList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim PrintType As String

Private Sub cmdEdit_Click()
  Call Edit
End Sub

Private Sub cmdExit_Click()
  KillFile "C:\CPWork\editpyment.dat"
  GCustNum = 0
  Unload Me
  DoEvents
  frmTaxPayMenu.Show
End Sub

Private Sub cmdPrint_Click()
  frmTaxRptOptForPayEdit.Show vbModal
  If frmTaxRptOptForPayEdit.fptxtPrintType.Text = "Graphical Name" Then
    Unload frmTaxRptOptForPayEdit
    PrintType = "N"
    Call PrintGraphics
  ElseIf frmTaxRptOptForPayEdit.fptxtPrintType.Text = "Graphical Entry" Then
    Unload frmTaxRptOptForPayEdit
    PrintType = "E"
    Call PrintGraphics
  ElseIf frmTaxRptOptForPayEdit.fptxtPrintType.Text = "Text Name" Then
    PrintType = "N"
    frmTaxMsg.Label1.Caption = "Pitch 17 is recommended for this report."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Unload frmTaxRptOptForPayEdit
    Call PrintText
  ElseIf frmTaxRptOptForPayEdit.fptxtPrintType.Text = "Text Entry" Then
    PrintType = "E"
    frmTaxMsg.Label1.Caption = "Pitch 17 is recommended for this report."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Unload frmTaxRptOptForPayEdit
    Call PrintText
  End If

End Sub
Private Sub PrintGraphics()
  Dim PayRec As TaxPaymentRecType
  Dim PayHandle As Integer
  Dim NumOfPRecs As Integer
  Dim x As Integer
  Dim OHandle As Integer
  Dim NumOfOperRecs As Integer
  Dim OperRec As CitiPassType
  Dim NumOfOperPayRecs As Integer
  Dim dlm$
  Dim TaxMasterRec As TaxMasterType
  Dim MHandle As Integer
  Dim Town$
  Dim RptFile$
  Dim RptHandle As Integer
  Dim SubRptFile1$
  Dim SubRptHandle1 As Integer
  Dim SubRptFile2$
  Dim SubRptHandle2 As Integer
  Dim SubRptFile3$
  Dim SubRptHandle3 As Integer
  Dim TempPayRec As PayListType
  Dim THandle As Integer
  Dim NumOfTRecs As Integer
  Dim GPrinc As Double
  Dim GInt As Double
  Dim GAdvCol As Double
  Dim GLateList As Double
  Dim GRev1 As Double
  Dim GRev2 As Double
  Dim GRev3 As Double
  Dim GTot As Double
  Dim GOverPay As Double
  Dim Operator$
  Dim TotalPaid#
  Dim GDisc As Double
  Dim YearCnt As Integer
  Dim y As Integer
  Dim ThisYear As Integer
  Dim PrincByYrTot As Double
  Dim IntByYrTot As Double
  Dim AdvColByYrTot As Double
  Dim LateListByYrTot As Double
  Dim Rev1ByYrTot As Double
  Dim Rev2ByYrTot As Double
  Dim Rev3ByYrTot As Double
  Dim DiscByYrTot As Double
  Dim TotPaidByYrTot As Double
  Dim OverPayByYrTot As Double
  Dim CheckCnt As Integer
  Dim HoldPrincByYr As Double
  Dim HoldIntByYr As Double
  Dim HoldAdvColByYr As Double
  Dim HoldLateListByYr As Double
  Dim HoldRev1ByYr As Double
  Dim HoldRev2ByYr As Double
  Dim HoldRev3ByYr As Double
  Dim HoldDiscByYr As Double
  Dim HoldOverPayByYr As Double
  Dim HoldTotPaidByYr As Double
  Dim Thisx As Integer
  Dim LilYear As Integer
  Dim Nextx As Integer
  Dim HoldYears As Integer
  Dim Done As Boolean
  Dim OPAmt As Double
  Dim BigName$
  Dim LilName$
  Dim HoldName$
  Dim HoldRec As Integer
  Dim NextOne As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
'  Dim Tot As Double
  
  'on error goto ERRORSTUFF
  
  OpenTaxSetUpFile MHandle
  Get MHandle, 1, TaxMasterRec
  Close MHandle
  Town = QPTrim$(TaxMasterRec.Name)
  
  Operator$ = "Operator # " + CStr(OperNum) + " " + PWUser
  fptxtOperator.Text = "Operator # " + CStr(OperNum)
  OpenTempPayFile PayHandle, OperNum
  NumOfPRecs = LOF(PayHandle) / Len(PayRec)
  If NumOfPRecs = 0 Then
    frmTaxMsg.Label1.Caption = "There are no payment records saved for operator #" + CStr(OperNum) + "."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Close
    Exit Sub
  End If
  
  dlm = "~"
  RptFile$ = "TAXRPTS\TaxEdPay.RPT"
  RptHandle = FreeFile
  CheckCnt = 0
  OPAmt = 0
  TotalPaid = 0
  Open RptFile$ For Output As #RptHandle
  ReDim CustNArr(1 To NumOfPRecs) As String
  ReDim CustRArr(1 To NumOfPRecs) As Integer
  GoSub SortCustomers
  For x = 1 To NumOfPRecs
    If PrintType = "N" Then
      Get PayHandle, CustRArr(x), PayRec
    Else
      Get PayHandle, x, PayRec
    End If
      PayRec.CustAcct = PayRec.CustAcct
      If PayRec.LastPayRec = 0 Then GoTo MoveOn
      If PayRec.ChkAmt > 0 Then
        CheckCnt = CheckCnt + 1
      End If
      TotalPaid = OldRound(PayRec.CashAmt + PayRec.ChkAmt + PayRec.ChrgAmt + PayRec.DiscAmt - PayRec.Change)
'      Tot = Tot + TotalPaid
'      Debug.Print Using$("$###0.00", TotalPaid) + " " + CStr(PayRec.CustAcct)
      OPAmt = OldRound(OPAmt + PayRec.PrePayAmt)
      '                   0                    1                               2
      Print #RptHandle, Town; dlm; MakeRegDate(PayRec.PayDate); dlm; CStr(PayRec.CustAcct); dlm;
      '                           3                          4                    5                   6
      Print #RptHandle, QPTrim$(PayRec.CustName); dlm; PayRec.CashAmt; dlm; PayRec.ChkAmt; dlm; PayRec.ChrgAmt; dlm;
      '                       7                 8                  9                10               11
      Print #RptHandle, PayRec.DiscAmt; dlm; TotalPaid; dlm; PayRec.Change; dlm; Operator; dlm; CStr(CheckCnt); dlm; OPAmt
MoveOn:
  Next x
  Close RptHandle
  
  If OPAmt > 0 Then
    SubRptFile3$ = "TAXRPTS\SubEdPay3.RPT"
    SubRptHandle3 = FreeFile
    Open SubRptFile3$ For Output As #SubRptHandle3
    For x = 1 To NumOfPRecs
      If PrintType = "N" Then
        Get PayHandle, CustRArr(x), PayRec
      Else
        Get PayHandle, x, PayRec
      End If
'      Get PayHandle, CustRArr(x), PayRec
      If PayRec.LastPayRec = 0 Then GoTo Deleted
      If PayRec.PrePayAmt > 0 Then
        Print #SubRptHandle3, PayRec.CustAcct; dlm; QPTrim$(PayRec.CustName); dlm; PayRec.TotPaid; dlm; PayRec.PrePayAmt
      End If
Deleted:
    Next x
  End If
  Close SubRptHandle3
  
  Close PayHandle
    
  GPrinc = 0
  GInt = 0
  GAdvCol = 0
  GLateList = 0
  GRev1 = 0
  GRev2 = 0
  GRev3 = 0
  GTot = 0
  GDisc = 0
  GOverPay = 0
  
  YearCnt = 0
  ReDim Years(1 To 1) As Integer
  
  SubRptFile1$ = "TAXRPTS\SubEdPay1.RPT"
  SubRptHandle1 = FreeFile
  Open SubRptFile1$ For Output As #SubRptHandle1
  OpenPayListFile THandle, OperNum
  NumOfTRecs = LOF(THandle) / Len(TempPayRec)
  YearCnt = 0
  For x = 1 To NumOfTRecs
    Get THandle, x, TempPayRec
    ThisYear = TempPayRec.TaxYear
    For y = 1 To YearCnt
      If y <> x Then
        If TempPayRec.TaxYear = Years(y) Then
          Exit For
        End If
      End If
    Next y
    If y > YearCnt Then
      YearCnt = YearCnt + 1
      ReDim Preserve Years(1 To YearCnt) As Integer
      Years(YearCnt) = ThisYear
    End If
  Next x
  
  ReDim PrincByYr(1 To YearCnt) As Double
  ReDim IntByYr(1 To YearCnt) As Double
  ReDim AdvColByYr(1 To YearCnt) As Double
  ReDim LateListByYr(1 To YearCnt) As Double
  ReDim Rev1ByYr(1 To YearCnt) As Double
  ReDim Rev2ByYr(1 To YearCnt) As Double
  ReDim Rev3ByYr(1 To YearCnt) As Double
  ReDim DiscByYr(1 To YearCnt) As Double
  ReDim TotPaidByYr(1 To YearCnt) As Double
  ReDim OverPayByYr(1 To YearCnt) As Double
  
  For x = 1 To NumOfTRecs
    Get THandle, x, TempPayRec
'       If TempPayRec.CustRec = 4382 Then Stop
'       If TempPayRec.CustRec = 11397 Then Stop
       If TempPayRec.PrevListRec < 0 Then GoTo MoveIt 'deleted transaction
       TempPayRec.CustRec = TempPayRec.CustRec
       GPrinc = OldRound(GPrinc + TempPayRec.Principle1)
'       Debug.Print Using$("$##0.00", TempPayRec.Principle1) + " " + CStr(TempPayRec.CustRec)

       GTot = OldRound(GTot + TempPayRec.Principle1)
       GInt = OldRound(GInt + TempPayRec.Interest1)
'       Debug.Print Using$("$##0.00", TempPayRec.Interest1)
       GTot = OldRound(GTot + TempPayRec.Interest1)
       GAdvCol = OldRound(GAdvCol + TempPayRec.Collection)
'       Debug.Print Using$("$##0.00", TempPayRec.Collection)
'       Debug.Print ""
       GTot = OldRound(GTot + TempPayRec.Collection)
       GLateList = OldRound(GLateList + TempPayRec.LateList)
       GTot = OldRound(GTot + TempPayRec.LateList)
       GRev1 = OldRound(GRev1 + TempPayRec.OptRev1)
       GTot = OldRound(GTot + TempPayRec.OptRev1)
       GRev2 = OldRound(GRev2 + TempPayRec.OptRev2)
       GTot = OldRound(GTot + TempPayRec.OptRev2)
       GRev3 = OldRound(GRev3 + TempPayRec.OptRev3)
       GTot = OldRound(GTot + TempPayRec.OptRev3)
       GDisc = OldRound(GDisc + TempPayRec.DiscAmt)
       GTot = OldRound(GTot + TempPayRec.DiscAmt)
       GOverPay = OldRound(GOverPay + TempPayRec.PrePayAmt)
       GTot = OldRound(GTot + TempPayRec.PrePayAmt)
       For y = 1 To YearCnt
         If Years(y) = TempPayRec.TaxYear Then
           PrincByYr(y) = OldRound(PrincByYr(y) + TempPayRec.Principle1)
           IntByYr(y) = OldRound(IntByYr(y) + TempPayRec.Interest1)
           AdvColByYr(y) = OldRound(AdvColByYr(y) + TempPayRec.Collection)
           LateListByYr(y) = OldRound(LateListByYr(y) + TempPayRec.LateList)
           Rev1ByYr(y) = OldRound(Rev1ByYr(y) + TempPayRec.OptRev1)
           Rev2ByYr(y) = OldRound(Rev2ByYr(y) + TempPayRec.OptRev2)
           Rev3ByYr(y) = OldRound(Rev3ByYr(y) + TempPayRec.OptRev3)
           DiscByYr(y) = OldRound(DiscByYr(y) + TempPayRec.DiscAmt)
           TotPaidByYr(y) = OldRound(TotPaidByYr(y) + TempPayRec.TotPaid)
           OverPayByYr(y) = OldRound(OverPayByYr(y) + TempPayRec.PrePayAmt)
           Exit For
         End If
       Next y
MoveIt:
  Next x
  '                        0                            1           2              3
  Print #SubRptHandle1, OldRound(GPrinc + GDisc); dlm; GInt; dlm; GAdvCol; dlm; GLateList; dlm;
  '                       4           5           6           7           8           9           10
  Print #SubRptHandle1, GRev1; dlm; GRev2; dlm; GRev3; dlm; GTot; dlm; GPrinc; dlm; GDisc; dlm; GOverPay; dlm;
  '                                11                                   12                                    13
  Print #SubRptHandle1, QPTrim$(TaxMasterRec.OptRev1); dlm; QPTrim$(TaxMasterRec.OptRev2); dlm; QPTrim$(TaxMasterRec.OptRev3)
  Close SubRptHandle1
  Close THandle
    
  SubRptFile2$ = "TAXRPTS\SubEdPay2.RPT"
  SubRptHandle2 = FreeFile
  Open SubRptFile2$ For Output As #SubRptHandle2
  Done = False
  GoSub SortYears
  For x = 1 To YearCnt
    If x = YearCnt Then Done = True
    PrincByYrTot = OldRound(PrincByYrTot + PrincByYr(x) + DiscByYr(x))
    IntByYrTot = OldRound(IntByYrTot + IntByYr(x))
    AdvColByYrTot = OldRound(AdvColByYrTot + AdvColByYr(x))
    LateListByYrTot = OldRound(LateListByYrTot + LateListByYr(x))
    Rev1ByYrTot = OldRound(Rev1ByYrTot + Rev1ByYr(x))
    Rev2ByYrTot = OldRound(Rev2ByYrTot + Rev2ByYr(x))
    Rev3ByYrTot = OldRound(Rev3ByYrTot + Rev3ByYr(x))
    DiscByYrTot = OldRound(DiscByYrTot + DiscByYr(x))
    OverPayByYrTot = OldRound(OverPayByYrTot + OverPayByYr(x))
    TotPaidByYrTot = OldRound(TotPaidByYrTot + TotPaidByYr(x))
    '                        0                                 1                       2                 3
    Print #SubRptHandle2, Years(x); dlm; OldRound(PrincByYr(x) + DiscByYr(x)); dlm; IntByYr(x); dlm; AdvColByYr(x); dlm;
    '                           4                   5                 6                 7
    Print #SubRptHandle2, LateListByYr(x); dlm; Rev1ByYr(x); dlm; Rev2ByYr(x); dlm; Rev3ByYr(x); dlm;
    '                          8                  9                10                 11
    Print #SubRptHandle2, DiscByYr(x); dlm; PrincByYrTot; dlm; IntByYrTot; dlm; AdvColByYrTot; dlm;
    '                           12                  13                14                 15
    Print #SubRptHandle2, LateListByYrTot; dlm; Rev1ByYrTot; dlm; Rev2ByYrTot; dlm; Rev3ByYrTot; dlm;
    '                         16                        17                               18                                    19
    Print #SubRptHandle2, DiscByYrTot; dlm; QPTrim$(TaxMasterRec.OptRev1); dlm; QPTrim$(TaxMasterRec.OptRev2); dlm; QPTrim$(TaxMasterRec.OptRev3); dlm;
    '                      20             21                    22
    Print #SubRptHandle2, Done; dlm; OverPayByYr(x); dlm; OverPayByYrTot
  
  Next x
  
  Close SubRptHandle2
  
  arTaxPayTransRpt.Show
  
  Exit Sub
  
SortYears:
  LilYear = 1900
  Nextx = 1
  Do
    For x = Nextx To YearCnt
      If Years(x) > LilYear Then
        LilYear = Years(x)
        Thisx = x
      End If
    Next x
    HoldYears = Years(Nextx)
    HoldPrincByYr = PrincByYr(Nextx)
    HoldIntByYr = IntByYr(Nextx)
    HoldAdvColByYr = AdvColByYr(Nextx)
    HoldLateListByYr = LateListByYr(Nextx)
    HoldRev1ByYr = Rev1ByYr(Nextx)
    HoldRev2ByYr = Rev2ByYr(Nextx)
    HoldRev3ByYr = Rev3ByYr(Nextx)
    HoldDiscByYr = DiscByYr(Nextx)
    HoldTotPaidByYr = TotPaidByYr(Nextx)
    Years(Nextx) = Years(Thisx)
    PrincByYr(Nextx) = PrincByYr(Thisx)
    IntByYr(Nextx) = IntByYr(Thisx)
    AdvColByYr(Nextx) = AdvColByYr(Thisx)
    LateListByYr(Nextx) = LateListByYr(Thisx)
    Rev1ByYr(Nextx) = Rev1ByYr(Thisx)
    Rev2ByYr(Nextx) = Rev2ByYr(Thisx)
    Rev3ByYr(Nextx) = Rev3ByYr(Thisx)
    DiscByYr(Nextx) = DiscByYr(Thisx)
    TotPaidByYr(Nextx) = TotPaidByYr(Thisx)
    Years(Thisx) = HoldYears
    PrincByYr(Thisx) = HoldPrincByYr
    IntByYr(Thisx) = HoldIntByYr
    AdvColByYr(Thisx) = HoldAdvColByYr
    LateListByYr(Thisx) = HoldLateListByYr
    Rev1ByYr(Thisx) = HoldRev1ByYr
    Rev2ByYr(Thisx) = HoldRev2ByYr
    Rev3ByYr(Thisx) = HoldRev3ByYr
    DiscByYr(Thisx) = HoldDiscByYr
    TotPaidByYr(Thisx) = HoldTotPaidByYr
    LilYear = 1900
    Nextx = Nextx + 1
    If Nextx > YearCnt Then Exit Do
  Loop
  Return
  
SortCustomers:
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  For x = 1 To NumOfPRecs
    Get PayHandle, x, PayRec
    Get TCHandle, PayRec.CustAcct, TaxCust
'    CustNArr(x) = QPTrim$(PayRec.CustName)
    CustNArr(x) = QPTrim$(TaxCust.SName)
    CustRArr(x) = x
  Next x
  Close TCHandle
  
  BigName$ = ""
  For x = 1 To NumOfPRecs
    If CustNArr(x) > BigName Then
      BigName = CustNArr(x)
    End If
  Next x
  
  LilName = BigName + "z"
  NextOne = 1
  
  Do
    For x = NextOne To NumOfPRecs
      If CustNArr(x) < LilName Then
        LilName = CustNArr(x)
        Thisx = x
      End If
    Next x
    HoldName = CustNArr(NextOne)
    HoldRec = CustRArr(NextOne)
    CustNArr(NextOne) = CustNArr(Thisx)
    CustRArr(NextOne) = CustRArr(Thisx)
    CustNArr(Thisx) = HoldName
    CustRArr(Thisx) = HoldRec
    NextOne = NextOne + 1
    LilName = BigName + "z"
    If NextOne > NumOfPRecs Then Exit Do
  Loop
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPayEditList", "PrintGraphics", Erl)
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%d"
      Call cmdEdit_Click
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%P"
      Call cmdPrint_Click
      KeyCode = 0
    Case vbKeyReturn:
      Call Edit
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpPrintTransaction
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      KillFile "C:\CPWork\editpyment.dat"
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxPayEditList.")
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

Private Sub LoadMe()
  Dim PayRec As TaxPaymentRecType
  Dim PayHandle As Integer
  Dim NumOfPRecs As Integer
  Dim x As Integer
  Dim OHandle As Integer
  Dim NumOfOperRecs As Integer
  Dim OperRec As CitiPassType
  Dim NumOfOperPayRecs As Integer
  Dim One As Integer
  Dim ThisFile As Integer
  Dim FileName$
  
  'on error goto ERRORSTUFF
  
  FileName = "C:\CPWork\editpyment.dat" 'used when using the transaction history report
  ThisFile = FreeFile
  Open FileName For Output As ThisFile
  One = 1
  Print #ThisFile, One
  Close ThisFile
  
'  OPERNUM = 1
  fptxtOperator.Text = "Operator # " + CStr(OperNum)
  OpenTempPayFile PayHandle, OperNum
  NumOfPRecs = LOF(PayHandle) / Len(PayRec)
  If NumOfPRecs = 0 Then
    frmTaxMsg.Label1.Caption = "There are no payment records saved for operator #" + CStr(OperNum) + "."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Close
    Exit Sub
  End If
  
  For x = 1 To NumOfPRecs
  Get PayHandle, x, PayRec
    If PayRec.LastPayRec <> 0 Then
      fpListPay.InsertRow = CStr(PayRec.CustAcct) + Chr(9) + QPTrim$(PayRec.CustName) + Chr(9) + QPTrim$(Using$("$###,##0.00", PayRec.AmtPaid)) + Chr(9) + QPTrim$(Using$("$###,##0.00", PayRec.AmtOwed))
      DoEvents
    End If
  Next x
  Close PayHandle
  fpListPay.ListIndex = 0
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPayEditList", "LoadMe", Erl)
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

Private Sub Edit()
  Dim PayRec As TaxPaymentRecType
  Dim PayHandle As Integer
  Dim NumOfPRecs As Integer
  Dim x As Integer
  Dim Operator$
  Dim AcctNum As Long
  
  'on error goto ERRORSTUFF
  
  If fpListPay.ListIndex = -1 Then
    Call TaxMsg(900, "Please make a selection from the list.")
    Exit Sub
  End If
  
  fpListPay.Col = 0
  fpListPay.Row = fpListPay.ListIndex
  AcctNum = CLng(QPTrim$(fpListPay.ColText))
  
  Operator = CStr(OperNum)
  GCustNum = 0
  GPayNum = 0
  OpenTempPayFile PayHandle, OperNum
  NumOfPRecs = LOF(PayHandle) / Len(PayRec)
  For x = 1 To NumOfPRecs
    Get PayHandle, x, PayRec
    If PayRec.CustAcct = AcctNum Then
      GCustNum = AcctNum
      GPayNum = x
      Exit For
    End If
  Next x
  Close PayHandle
  
  If x > NumOfPRecs Then
    Call TaxMsg(900, "The transaction selected could not be found.")
    Exit Sub
  End If
  
  frmTaxPaymentEntry.fpLongAcctNum = GCustNum
  frmTaxPaymentEntry.Show
  DoEvents
  Me.Hide
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPayEditList", "Edit", Erl)
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

Private Sub fpListPay_DblClick()
  Call Edit
End Sub

Private Sub PrintText()
  Dim PayRec As TaxPaymentRecType
  Dim PayHandle As Integer
  Dim NumOfPRecs As Integer
  Dim x As Integer
  Dim OHandle As Integer
  Dim NumOfOperRecs As Integer
  Dim OperRec As CitiPassType
  Dim NumOfOperPayRecs As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim MHandle As Integer
  Dim Town$
  Dim RptFile$
  Dim RptHandle As Integer
  Dim TempPayRec As PayListType
  Dim THandle As Integer
  Dim NumOfTRecs As Integer
  Dim GPrinc As Double
  Dim GInt As Double
  Dim GAdvCol As Double
  Dim GLateList As Double
  Dim GRev1 As Double
  Dim GRev2 As Double
  Dim GRev3 As Double
  Dim GTot As Double
  Dim GOverPay As Double
  Dim Operator$
  Dim TotalPaid#
  Dim GDisc As Double
  Dim YearCnt As Integer
  Dim y As Integer
  Dim ThisYear As Integer
  Dim PrincByYrTot As Double
  Dim IntByYrTot As Double
  Dim AdvColByYrTot As Double
  Dim LateListByYrTot As Double
  Dim Rev1ByYrTot As Double
  Dim Rev2ByYrTot As Double
  Dim Rev3ByYrTot As Double
  Dim DiscByYrTot As Double
  Dim TotPaidByYrTot As Double
  Dim OverPayByYrTot As Double
  Dim CheckCnt As Integer
  Dim HoldPrincByYr As Double
  Dim HoldIntByYr As Double
  Dim HoldAdvColByYr As Double
  Dim HoldLateListByYr As Double
  Dim HoldRev1ByYr As Double
  Dim HoldRev2ByYr As Double
  Dim HoldRev3ByYr As Double
  Dim HoldDiscByYr As Double
  Dim HoldOverPayByYr As Double
  Dim HoldTotPaidByYr As Double
  Dim Thisx As Integer
  Dim LilYear As Integer
  Dim Nextx As Integer
  Dim HoldYears As Integer
  Dim Done As Boolean
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim FF$
  Dim Page As Integer
  Dim sLine$
  Dim DLine$
  Dim UseThis$
  Dim TotCash As Double
  Dim TotCheck As Double
  Dim TotCharge As Double
  Dim TotDisc As Double
  Dim TotCredit As Double
  Dim TotChange As Double
  Dim OperLen As Integer
  Dim ThisTab As Integer
  Dim NoD$
  Dim RevTrunc1 As String * 12
  Dim RevTrunc2 As String * 12
  Dim RevTrunc3 As String * 12
  Dim GTotal As Double
  Dim YTotal As Double
  Dim OPAmt As Double
  Dim OPToOwed As Double
  Dim OPPaid As Double
  Dim OPCnt As Integer
  Dim BigName$
  Dim LilName$
  Dim HoldName$
  Dim HoldRec As Integer
  Dim NextOne As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim ReceiptCnt As Integer
  
  'on error goto ERRORSTUFF
  
  NoD = "###,##0.00"
  UseThis$ = "$###,##0.00"
  DLine$ = String(109, "=")
  sLine$ = String(109, "-")
  FF$ = Chr(12)
  MaxLines = 56
  LineCnt = 0
  
  OpenTaxSetUpFile MHandle
  Get MHandle, 1, TaxMasterRec
  Close MHandle
  Town = QPTrim$(TaxMasterRec.Name)
  
  Operator$ = "Operator # " + CStr(OperNum) + " " + PWUser
  fptxtOperator.Text = "Operator # " + CStr(OperNum)
  OpenTempPayFile PayHandle, OperNum
  NumOfPRecs = LOF(PayHandle) / Len(PayRec)
  If NumOfPRecs = 0 Then
    frmTaxMsg.Label1.Caption = "There are no payment records saved for operator #" + CStr(OperNum) + "."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Close
    Exit Sub
  End If
  
  RptFile$ = "TAXRPTS\TaxEdPay.RPT"
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  
  GoSub PrintHeader
  CheckCnt = 0
  TotCash = 0
  TotCheck = 0
  TotCharge = 0
  TotDisc = 0
  TotCredit = 0
  TotChange = 0
  OPAmt = 0
  ReDim CustNArr(1 To NumOfPRecs) As String
  ReDim CustRArr(1 To NumOfPRecs) As Integer
  GoSub SortCustomers
  ReceiptCnt = 0
  For x = 1 To NumOfPRecs
    If PrintType = "N" Then
      Get PayHandle, CustRArr(x), PayRec
    Else
      Get PayHandle, x, PayRec
    End If
'    Get PayHandle, CustRArr(x), PayRec
      If PayRec.LastPayRec = 0 Then GoTo MoveOn
      If PayRec.ChkAmt > 0 Then
        CheckCnt = CheckCnt + 1
      End If
      If PayRec.PrePayAmt > 0 Then
        OPAmt = OPAmt + 1
      End If
      TotCash = OldRound(TotCash + PayRec.CashAmt)
      TotCheck = OldRound(TotCheck + PayRec.ChkAmt)
      TotCharge = OldRound(TotCharge + PayRec.ChrgAmt)
      TotDisc = OldRound(TotDisc + PayRec.DiscAmt)
      TotalPaid = OldRound(PayRec.CashAmt + PayRec.ChkAmt + PayRec.ChrgAmt + PayRec.DiscAmt - PayRec.Change)
      TotCredit = OldRound(TotCredit + TotalPaid)
      TotChange = OldRound(TotChange + PayRec.Change)
      Print #RptHandle, MakeRegDate(PayRec.PayDate); Tab(16); CStr(PayRec.CustAcct); Tab(26);
      Print #RptHandle, QPTrim$(PayRec.CustName); Tab(50); Using(NoD, PayRec.CashAmt); Tab(60); Using(NoD, PayRec.ChkAmt); Tab(70); Using(NoD, PayRec.ChrgAmt); Tab(80);
      Print #RptHandle, Using(NoD, PayRec.DiscAmt); Tab(90); Using(NoD, TotalPaid); Tab(100); Using(NoD, PayRec.Change)
      ReceiptCnt = ReceiptCnt + 1
      LineCnt = LineCnt + 1
      If x >= NumOfPRecs - 6 Then
        If LineCnt > MaxLines - 6 Then
          Print #RptHandle, FF$
          GoSub PrintHeader
        End If
      Else
        If LineCnt > MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintHeader
        End If
      End If
MoveOn:
  Next x
  
  Print #RptHandle, sLine
  Print #RptHandle, Tab(21); "Totals"; Tab(50); Using(NoD, TotCash); Tab(60); Using(NoD, TotCheck); Tab(70); Using(NoD, TotCharge); Tab(80); Using(NoD, TotDisc); Tab(90); Using(NoD, TotCredit); Tab(100); Using(NoD, TotChange)
  Print #RptHandle, "Total Number of Receipts: " + CStr(ReceiptCnt)
  Print #RptHandle, "Total Number of Checks: " + CStr(CheckCnt)
  LineCnt = LineCnt + 4
  
  If OPAmt > 0 Then
     If LineCnt >= MaxLines - 7 Then
       Print #RptHandle, FF$
       GoSub PrintOPHeader
       GoTo NewPage
     End If
     Print #RptHandle,
     Print #RptHandle,
     Print #RptHandle, String(89, "-")
     Print #RptHandle, "Over Payment Summary"
     Print #RptHandle, Tab(38); "Payment Applied To"; Tab(60); "Over Payment"; Tab(78); "Total Amount"
     Print #RptHandle, "Cust Num"; Tab(12); "Customer"; Tab(45); "Amount Owed"; Tab(66); "Amount"; Tab(86); "Paid"
     Print #RptHandle, String(89, "-")
     LineCnt = LineCnt + 7
NewPage:
     OPToOwed = 0
     OPPaid = 0
     OPCnt = 0
     For x = 1 To NumOfPRecs
       If PrintType = "N" Then
         Get PayHandle, CustRArr(x), PayRec
       Else
         Get PayHandle, x, PayRec
       End If
'       Get PayHandle, CustRArr(x), PayRec
       If PayRec.LastPayRec = 0 Then GoTo Deleted
       If PayRec.PrePayAmt > 0 Then
         OPToOwed = OldRound(OPToOwed + PayRec.TotPaid)
         OPPaid = OldRound(OPPaid + PayRec.PrePayAmt)
         OPCnt = OPCnt + 1
         Print #RptHandle, PayRec.CustAcct; Tab(12); QPTrim$(PayRec.CustName); Tab(45); Using(UseThis, PayRec.TotPaid); Tab(61); Using(UseThis, PayRec.PrePayAmt); Tab(79); Using(UseThis, OldRound(PayRec.TotPaid + PayRec.PrePayAmt))
         LineCnt = LineCnt + 1
       End If
       If LineCnt >= MaxLines - 7 Then
         Print #RptHandle, FF$
         GoSub PrintOPHeader
       End If
Deleted:
     Next x
     If LineCnt >= MaxLines - 7 Then
       Print #RptHandle, FF$
       GoSub PrintOPHeader
     End If
     If LineCnt >= MaxLines - 2 Then
       Print #RptHandle, FF$
       GoSub PrintOPHeader
     End If
     Print #RptHandle, String(89, "-")
     Print #RptHandle, "Totals"; Tab(12); "# Over Payments: " + CStr(OPCnt); Tab(45); Using(UseThis, OPToOwed); Tab(61); Using(UseThis, OPPaid); Tab(79); Using(UseThis, OldRound(OPToOwed + OPPaid))
  Else
    Print #RptHandle,
    Print #RptHandle, "No Over Payment Activity"
  End If
  
  Close PayHandle
  GPrinc = 0
  GInt = 0
  GAdvCol = 0
  GLateList = 0
  GRev1 = 0
  GRev2 = 0
  GRev3 = 0
  GTot = 0
  GDisc = 0
  GOverPay = 0
  
  YearCnt = 0
  ReDim Years(1 To 1) As Integer
  
  OpenPayListFile THandle, OperNum
  NumOfTRecs = LOF(THandle) / Len(TempPayRec)
  YearCnt = 0
  For x = 1 To NumOfTRecs
    Get THandle, x, TempPayRec
    ThisYear = TempPayRec.TaxYear
    For y = 1 To YearCnt
      If y <> x Then
        If TempPayRec.TaxYear = Years(y) Then
          Exit For
        End If
      End If
    Next y
    If y > YearCnt Then
      YearCnt = YearCnt + 1
      ReDim Preserve Years(1 To YearCnt) As Integer
      Years(YearCnt) = ThisYear
    End If
  Next x
  
  ReDim PrincByYr(1 To YearCnt) As Double
  ReDim IntByYr(1 To YearCnt) As Double
  ReDim AdvColByYr(1 To YearCnt) As Double
  ReDim LateListByYr(1 To YearCnt) As Double
  ReDim Rev1ByYr(1 To YearCnt) As Double
  ReDim Rev2ByYr(1 To YearCnt) As Double
  ReDim Rev3ByYr(1 To YearCnt) As Double
  ReDim DiscByYr(1 To YearCnt) As Double
  ReDim TotPaidByYr(1 To YearCnt) As Double
  ReDim OverPayByYr(1 To YearCnt) As Double
  
  For x = 1 To NumOfTRecs
    Get THandle, x, TempPayRec
      If TempPayRec.PrevListRec < 0 Then GoTo NextOne
       GPrinc = OldRound(GPrinc + TempPayRec.Principle1)
       GTot = OldRound(GTot + TempPayRec.Principle1)
       GInt = OldRound(GInt + TempPayRec.Interest1)
       GTot = OldRound(GTot + TempPayRec.Interest1)
       GAdvCol = OldRound(GAdvCol + TempPayRec.Collection)
       GTot = OldRound(GTot + TempPayRec.Collection)
       GLateList = OldRound(GLateList + TempPayRec.LateList)
       GTot = OldRound(GTot + TempPayRec.LateList)
       GRev1 = OldRound(GRev1 + TempPayRec.OptRev1)
       GTot = OldRound(GTot + TempPayRec.OptRev1)
       GRev2 = OldRound(GRev2 + TempPayRec.OptRev2)
       GTot = OldRound(GTot + TempPayRec.OptRev2)
       GRev3 = OldRound(GRev3 + TempPayRec.OptRev3)
       GTot = OldRound(GTot + TempPayRec.OptRev3)
       GDisc = OldRound(GDisc + TempPayRec.DiscAmt)
       GTot = OldRound(GTot + TempPayRec.DiscAmt)
       GOverPay = OldRound(GOverPay + TempPayRec.PrePayAmt)
       GTot = OldRound(GTot + TempPayRec.PrePayAmt)
       For y = 1 To YearCnt
         If Years(y) = TempPayRec.TaxYear Then
           PrincByYr(y) = OldRound(PrincByYr(y) + TempPayRec.Principle1) ' + TempPayRec.DiscAmt)
           IntByYr(y) = OldRound(IntByYr(y) + TempPayRec.Interest1)
           AdvColByYr(y) = OldRound(AdvColByYr(y) + TempPayRec.Collection)
           LateListByYr(y) = OldRound(LateListByYr(y) + TempPayRec.LateList)
           Rev1ByYr(y) = OldRound(Rev1ByYr(y) + TempPayRec.OptRev1)
           Rev2ByYr(y) = OldRound(Rev2ByYr(y) + TempPayRec.OptRev2)
           Rev3ByYr(y) = OldRound(Rev3ByYr(y) + TempPayRec.OptRev3)
           DiscByYr(y) = OldRound(DiscByYr(y) + TempPayRec.DiscAmt)
           TotPaidByYr(y) = OldRound(TotPaidByYr(y) + TempPayRec.TotPaid)
           OverPayByYr(y) = OldRound(OverPayByYr(y) + TempPayRec.PrePayAmt)
           Exit For
         End If
       Next y
NextOne:
  Next x
  Print #RptHandle, FF$
  GoSub PrintSubHeader
  
  Print #RptHandle, "Source Summary"
  Print #RptHandle, Tab(20); "Revenue Source"; Tab(55); "Amount"
  Print #RptHandle, Tab(20); String(41, "-")
  Print #RptHandle, Tab(20); "Tax Principle + Discount:"; Tab(50); Using(UseThis, OldRound(GPrinc + GDisc))
  Print #RptHandle, Tab(20); String(41, ".")
  Print #RptHandle, Tab(20); "Tax Principle:"; Tab(50); Using(UseThis, GPrinc)
  Print #RptHandle, Tab(20); "Discount:"; Tab(50); Using(UseThis, GDisc)
  Print #RptHandle, Tab(20); "Interest:"; Tab(50); Using(UseThis, GInt)
  Print #RptHandle, Tab(20); "Advertising/Collections"; Tab(50); Using(UseThis, GAdvCol)
  Print #RptHandle, Tab(20); "Late Listing:"; Tab(50); Using(UseThis, GLateList)
  Print #RptHandle, Tab(20); QPTrim$(TaxMasterRec.OptRev1); Tab(50); Using(UseThis, GRev1)
  Print #RptHandle, Tab(20); QPTrim$(TaxMasterRec.OptRev2); Tab(50); Using(UseThis, GRev2)
  Print #RptHandle, Tab(20); QPTrim$(TaxMasterRec.OptRev3); Tab(50); Using(UseThis, GRev3)
  Print #RptHandle, Tab(20); "Over Payment:"; Tab(50); Using(UseThis, GOverPay)
  Print #RptHandle, Tab(20); Tab(20); String(41, "-")
  Print #RptHandle, Tab(20); "Grand Totals:"; Tab(50); Using(UseThis, GTot)
  
  Close THandle
  RevTrunc1 = QPTrim$(TaxMasterRec.OptRev1)
  RevTrunc2 = QPTrim$(TaxMasterRec.OptRev2)
  RevTrunc3 = QPTrim$(TaxMasterRec.OptRev3)
  RSet RevTrunc1 = QPTrim$(RevTrunc1)
  RSet RevTrunc2 = QPTrim$(RevTrunc2)
  RSet RevTrunc3 = QPTrim$(RevTrunc3)
  
  GoSub SortYears
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, "Breakdown By Year"
  Print #RptHandle, Tab(2); "Tax Year"; Tab(14); "Principle Paid"; Tab(34); "Adv/Col Paid"; Tab(52); RevTrunc1; Tab(72); RevTrunc3; Tab(90); "Discounts Allowed"
  Print #RptHandle, Tab(15); "Interest Paid"; Tab(32); "Late List Paid"; Tab(52); RevTrunc2; Tab(72); "Over Payment"; Tab(92); "Payments - Disc"
  Print #RptHandle, String(106, "-")
  
  For x = 1 To YearCnt
    If x = YearCnt Then Done = True
    PrincByYrTot = OldRound(PrincByYrTot + PrincByYr(x)) ' + DiscByYr(x))
    IntByYrTot = OldRound(IntByYrTot + IntByYr(x))
    AdvColByYrTot = OldRound(AdvColByYrTot + AdvColByYr(x))
    LateListByYrTot = OldRound(LateListByYrTot + LateListByYr(x))
    Rev1ByYrTot = OldRound(Rev1ByYrTot + Rev1ByYr(x))
    Rev2ByYrTot = OldRound(Rev2ByYrTot + Rev2ByYr(x))
    Rev3ByYrTot = OldRound(Rev3ByYrTot + Rev3ByYr(x))
    DiscByYrTot = OldRound(DiscByYrTot + DiscByYr(x))
    OverPayByYrTot = OldRound(OverPayByYrTot + OverPayByYr(x))
    TotPaidByYrTot = OldRound(TotPaidByYrTot + TotPaidByYr(x))
    YTotal = OldRound(PrincByYr(x) + IntByYr(x) + AdvColByYr(x) + LateListByYr(x) + Rev1ByYr(x) + Rev2ByYr(x) + Rev3ByYr(x) + OverPayByYr(x))
    Print #RptHandle, Tab(4); CStr(Years(x)); Tab(17); Using(UseThis, OldRound(PrincByYr(x) + DiscByYr(x))); Tab(35); Using(UseThis, AdvColByYr(x));
    Print #RptHandle, Tab(53); Using(UseThis, Rev1ByYr(x)); Tab(73); Using(UseThis, Rev3ByYr(x)); Tab(96); Using(UseThis, DiscByYr(x))
    LineCnt = LineCnt + 1
    If x >= LineCnt - 4 Then
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintSubHeader
        Print #RptHandle, "BreakDown By Year"
      End If
    Else
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintSubHeader
        Print #RptHandle, "BreakDown By Year"
      End If
    End If
    Print #RptHandle, Tab(17); Using(UseThis, IntByYr(x)); Tab(35); Using(UseThis, LateListByYr(x)); Tab(53); Using(UseThis, Rev2ByYr(x));
    Print #RptHandle, Tab(73); Using(UseThis, OverPayByYr(x)); Tab(96); Using(UseThis, YTotal)
    LineCnt = LineCnt + 1
    If x >= LineCnt - 4 Then
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintSubHeader
        Print #RptHandle, "BreakDown By Year"
      End If
    Else
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintSubHeader
        Print #RptHandle, "BreakDown By Year"
      End If
    End If
    Print #RptHandle, String(106, "-")
    LineCnt = LineCnt + 1
    If x >= LineCnt - 4 Then
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintSubHeader
        Print #RptHandle, "BreakDown By Year"
      End If
    Else
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintSubHeader
        Print #RptHandle, "BreakDown By Year"
      End If
    End If
  Next x
  GTotal = OldRound(PrincByYrTot + IntByYrTot + AdvColByYrTot + LateListByYrTot + Rev1ByYrTot + Rev2ByYrTot + Rev3ByYrTot + OverPayByYrTot) ' - DiscByYrTot)
'  Print #RptHandle, Tab(4); "Totals:"; Tab(17); Using(UseThis, OldRound(PrincByYrTot)); Tab(35); Using(UseThis, AdvColByYrTot);
  Print #RptHandle, Tab(4); "Totals:"; Tab(17); Using(UseThis, OldRound(PrincByYrTot + DiscByYrTot)); Tab(35); Using(UseThis, AdvColByYrTot);
  Print #RptHandle, Tab(53); Using(UseThis, Rev1ByYrTot); Tab(73); Using(UseThis, Rev3ByYrTot); Tab(96); Using(UseThis, DiscByYrTot)
  Print #RptHandle, Tab(17); Using(UseThis, IntByYrTot); Tab(35); Using(UseThis, LateListByYrTot); Tab(53); Using(UseThis, Rev2ByYrTot);
  Print #RptHandle, Tab(73); Using(UseThis, OverPayByYrTot); Tab(96); Using(UseThis, GTotal)
  
  Print #RptHandle, FF$
  Close RptHandle
  
  ViewPrint RptFile$, "Tax Payment Transaction Edit Journal", True
  
  KillFile RptFile$
  
  Exit Sub
  
SortYears:
  
  LilYear = 1900
  Nextx = 1
  Do
    For x = Nextx To YearCnt
      If Years(x) > LilYear Then
        LilYear = Years(x)
        Thisx = x
      End If
    Next x
    HoldYears = Years(Nextx)
    HoldPrincByYr = PrincByYr(Nextx)
    HoldIntByYr = IntByYr(Nextx)
    HoldAdvColByYr = AdvColByYr(Nextx)
    HoldLateListByYr = LateListByYr(Nextx)
    HoldRev1ByYr = Rev1ByYr(Nextx)
    HoldRev2ByYr = Rev2ByYr(Nextx)
    HoldRev3ByYr = Rev3ByYr(Nextx)
    HoldDiscByYr = DiscByYr(Nextx)
    HoldTotPaidByYr = TotPaidByYr(Nextx)
    Years(Nextx) = Years(Thisx)
    PrincByYr(Nextx) = PrincByYr(Thisx)
    IntByYr(Nextx) = IntByYr(Thisx)
    AdvColByYr(Nextx) = AdvColByYr(Thisx)
    LateListByYr(Nextx) = LateListByYr(Thisx)
    Rev1ByYr(Nextx) = Rev1ByYr(Thisx)
    Rev2ByYr(Nextx) = Rev2ByYr(Thisx)
    Rev3ByYr(Nextx) = Rev3ByYr(Thisx)
    DiscByYr(Nextx) = DiscByYr(Thisx)
    TotPaidByYr(Nextx) = TotPaidByYr(Thisx)
    Years(Thisx) = HoldYears
    PrincByYr(Thisx) = HoldPrincByYr
    IntByYr(Thisx) = HoldIntByYr
    AdvColByYr(Thisx) = HoldAdvColByYr
    LateListByYr(Thisx) = HoldLateListByYr
    Rev1ByYr(Thisx) = HoldRev1ByYr
    Rev2ByYr(Thisx) = HoldRev2ByYr
    Rev3ByYr(Thisx) = HoldRev3ByYr
    DiscByYr(Thisx) = HoldDiscByYr
    TotPaidByYr(Thisx) = HoldTotPaidByYr
    LilYear = 1900
    Nextx = Nextx + 1
    If Nextx > YearCnt Then Exit Do
  Loop
  Return

PrintHeader:
  OperLen = Len("Operator # " + CStr(OperNum) + " " + PWUser)
  ThisTab = OperLen / 2
  ThisTab = ThisTab + 45
'  PWUser = "Bob"
  Page = Page + 1
  Print #RptHandle, Tab(45); "Tax Payment Transaction Journal"
  Print #RptHandle, Tab(ThisTab); "Operator # " + CStr(OperNum) + " " + PWUser
  Print #RptHandle, Town; Tab(100); "Page #" + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, Tab(95); "Total"
  Print #RptHandle, "Date"; Tab(13); "Cust Num"; Tab(26); "Customer"; Tab(56); "Cash"; Tab(65); "Check"; Tab(74); "Charge"; Tab(82); "Discount"; Tab(92); "Credited"; Tab(104); "Change"
  Print #RptHandle, sLine
  LineCnt = 7
  
  Return
 
PrintSubHeader:
  OperLen = Len("Operator # " + CStr(OperNum) + " " + PWUser)
  ThisTab = OperLen / 2
  ThisTab = ThisTab + 45
  
  Page = Page + 1
  Print #RptHandle, Tab(45); "Tax Payment Transaction Journal"
  Print #RptHandle, Tab(ThisTab); "Operator # " + CStr(OperNum) + " " + PWUser
  Print #RptHandle, Town; Tab(100); "Page #" + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, sLine
  LineCnt = 6
  
  Return
  
PrintOPHeader:
  OperLen = Len("Operator # " + CStr(OperNum) + " " + PWUser)
  ThisTab = OperLen / 2
  ThisTab = ThisTab + 45
  
  Page = Page + 1
  Print #RptHandle, Tab(45); "Tax Payment Transaction Journal"
  Print #RptHandle, Tab(ThisTab); "Operator # " + CStr(OperNum) + " " + PWUser
  Print #RptHandle, Town; Tab(100); "Page #" + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, String(89, "-")
  Print #RptHandle, "Over Payment Summary"
  Print #RptHandle, Tab(38); "Payment Applied To"; Tab(60); "Over Payment"; Tab(78); "Total Amount"
  Print #RptHandle, "Cust Num"; Tab(12); "Customer"; Tab(45); "Amount Owed"; Tab(66); "Amount"; Tab(86); "Paid"
  Print #RptHandle, String(89, "-")
  LineCnt = 6
  Return

SortCustomers:
  OpenTaxCustFile TCHandle, NumOfTCRecs
  For x = 1 To NumOfPRecs
    Get PayHandle, x, PayRec
    Get TCHandle, PayRec.CustAcct, TaxCust
'    CustNArr(x) = QPTrim$(PayRec.CustName)
    CustNArr(x) = QPTrim$(TaxCust.SName)
    CustRArr(x) = x
  Next x
  Close TCHandle
  
  BigName$ = ""
  For x = 1 To NumOfPRecs
    If CustNArr(x) > BigName Then
      BigName = CustNArr(x)
    End If
  Next x
  
  LilName = BigName + "z"
  NextOne = 1
  
  Do
    For x = NextOne To NumOfPRecs
      If CustNArr(x) < LilName Then
        LilName = CustNArr(x)
        Thisx = x
      End If
    Next x
    HoldName = CustNArr(NextOne)
    HoldRec = CustRArr(NextOne)
    CustNArr(NextOne) = CustNArr(Thisx)
    CustRArr(NextOne) = CustRArr(Thisx)
    CustNArr(Thisx) = HoldName
    CustRArr(Thisx) = HoldRec
    NextOne = NextOne + 1
    LilName = BigName + "z"
    If NextOne > NumOfPRecs Then Exit Do
  Loop
  
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPayEditList", "PrintText", Erl)
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
