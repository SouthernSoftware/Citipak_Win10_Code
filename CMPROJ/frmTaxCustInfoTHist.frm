VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmTaxCustInfoTHist 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Customer Information"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTaxCustInfoTHist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpList1 
      Height          =   4596
      Left            =   1260
      TabIndex        =   5
      Top             =   2040
      Width           =   9132
      _Version        =   196608
      _ExtentX        =   16108
      _ExtentY        =   8107
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
      ColDesigner     =   "frmTaxCustInfoTHist.frx":08CA
   End
   Begin EditLib.fpCurrency fpCurrBalance 
      Height          =   375
      Left            =   5813
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1935
      _Version        =   196608
      _ExtentX        =   3413
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
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
      Left            =   6180
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1935
      _Version        =   131072
      _ExtentX        =   3413
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
      ButtonDesigner  =   "frmTaxCustInfoTHist.frx":0C37
   End
   Begin EditLib.fpText fptxtThisCust 
      Height          =   390
      Left            =   2873
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "Enter the official name of your town here. For example, 'Town Of Washington'."
      Top             =   720
      Width           =   6015
      _Version        =   196608
      _ExtentX        =   10610
      _ExtentY        =   688
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
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
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   1
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
      MaxLength       =   50
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
   Begin fpBtnAtlLibCtl.fpBtn cmdDetail 
      Height          =   540
      Left            =   3540
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1935
      _Version        =   131072
      _ExtentX        =   3413
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
      ButtonDesigner  =   "frmTaxCustInfoTHist.frx":0E15
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Count:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   7200
      Width           =   4215
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   5175
      Left            =   1080
      Top             =   1920
      Width           =   9495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Balance:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3893
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Customer Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2933
      TabIndex        =   2
      Top             =   360
      Width           =   6015
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   900
      Index           =   1
      Left            =   1493
      Top             =   300
      Width           =   8655
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   960
      Left            =   1493
      Top             =   240
      Width           =   8655
   End
End
Attribute VB_Name = "frmTaxCustInfoTHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdDetail_Click()
  Dim TRHandle As Integer
  Dim NumOfTaxTRecs As Long
  Dim TaxTRec As TaxTransactionType
  Dim ThisRec As Long
  
  frmTaxCustInfoTHist.fpList1.col = 3
  frmTaxCustInfoTHist.fpList1.Row = frmTaxCustInfoTHist.fpList1.ListIndex
  ThisRec = CLng(frmTaxCustInfoTHist.fpList1.ColText)
  
  If ThisRec > 0 Then
    OpenTaxTransFile TRHandle, NumOfTaxTRecs
    Get TRHandle, ThisRec, TaxTRec
    Close
    If TaxTRec.TranType = 1 Then
      frmTaxTransDetail.Show vbModal
    Else
      frmTaxTransDetailNotBill.Show vbModal
    End If
  Else
    frmTaxTransDetailNotBill.Show vbModal
  End If
End Sub

Private Sub cmdExit_Click()
'  If Exist("txadjust.dat") Then
'    frmTaxAdjustments.Show
'    DoEvents
'  Else
  If Exist("C:\CPWork\custinq.dat") Then
    frmTaxCustInq.Show
    DoEvents
  End If
  Unload Me
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
    Case vbKeyF5:
      SendKeys "%D"
      Call cmdDetail_Click
      KeyCode = 0
  End Select

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call LoadMe
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      TXLog ("cm.exe terminated via menu bar on frmTaxCustInfoTHist.")
      Call CMTerminate
      End
    End If
  End If
End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ''' Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub LoadMe()
  Dim TaxRec As TaxCustType
  Dim THandle As Integer
  Dim NumOfCustRecs As Long
  Dim x As Long
  Dim TaxTRec As TaxTransactionType
  Dim TRHandle As Integer
  Dim NumOfTaxTRecs As Long
  Dim PrevTranRec&
  Dim ThisDate$
  Dim ThisType$
  Dim ThisAmt#
  Dim ThisRec&
  Dim BillType$
  Dim TType$, Interest#
  Dim TaxYear$, BillNum$
  Dim TCnt As Integer
  
  TCnt = 0
  OpenTaxCustFile THandle, NumOfCustRecs
  Get THandle, GCustNum, TaxRec
  Close THandle
  
  fptxtThisCust.Text = QPTrim$(TaxRec.CustName)
  
  If TaxRec.LastTrans > 0 Then
    fpCurrBalance = GetCustBalance(GCustNum, -1)
  Else
    fpCurrBalance = 0
  End If
  
  OpenTaxTransFile TRHandle, NumOfTaxTRecs
  PrevTranRec& = TaxRec.LastTrans
  
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      Get TRHandle, PrevTranRec&, TaxTRec
      TCnt = TCnt + 1
      ThisDate = MakeRegDate(TaxTRec.TransDate)
      GoSub GetTransType
      ThisType = BillType
      ThisAmt = OldRound(TaxTRec.Amount + TaxTRec.DiscAmt)
      If ThisType = "Credit Applied at Billing" Then
        ThisAmt = TaxTRec.Revenue.PrePaidUsed
      End If
      ThisRec = PrevTranRec&
      fpList1.InsertRow = ThisDate & Chr(9) & ThisType & Chr(9) & Using$("$##,###,##0.00", ThisAmt) & Chr(9) & Using$("##########", ThisRec)
      PrevTranRec& = TaxTRec.LastTrans
    Loop
  End If
  Close
  fpList1.ListIndex = 0
  
  Label3.Caption = "Transaction Count: " + CStr(TCnt)
  
  Exit Sub
  
GetTransType:
'  TType$ = QPTrim$(TaxTRec.Description)
  Select Case TaxTRec.TranType
  Case 1
    BillNum$ = ParseBillNum$(TaxTRec.Description)
    Select Case TaxTRec.BillType
    Case "R"
      BillType$ = "Real-Estate Bill: #" + BillNum
    Case "P"
      BillType$ = "Personal Property Bill: #" + BillNum
    Case "C"
      BillType$ = "Combined Bill: #" + BillNum
    Case "M"
      BillType$ = "Manual Bill: #" + BillNum
    End Select
    TaxYear$ = QPTrim$(Str$(TaxTRec.TaxYear))
  Case 2
    BillNum$ = ParseBillNum$(TaxTRec.Description)
    If Len(BillNum$) = 0 Then
      If QPTrim$(TaxTRec.Description) = "Prepay" Then
        BillType = "Prepayment"
      Else
        BillType$ = "Payment ??? "
      End If
    Else
      If TaxTRec.Revenue.PrePaidAmt > 0 Then
        BillType = "Pre/Payment on: "
      Else
        BillType$ = "Payment on: "
      End If
    End If
    BillType$ = BillType$ + BillNum$
  Case 3
    BillNum$ = ParseBillNum$(TaxTRec.Description)
    BillType$ = "Release on: " + BillNum
  Case 4
    BillNum$ = ParseBillNum$(TaxTRec.Description)
    BillType$ = "Interest on: " + BillNum$
    Interest# = TaxTRec.Revenue.Interest#
  Case 6
    BillNum$ = ParseBillNum$(TaxTRec.Description)
    BillType$ = "Collection/Ad Cost on: " + BillNum$
  Case 7
    BillNum$ = ParseBillNum$(TaxTRec.Description)
    BillType$ = "Adjust Paid Down on: " + BillNum$
  Case 9
    BillType$ = "Credit Applied at Billing"
  Case 13
    BillNum$ = ParseBillNum$(TaxTRec.Description)
    BillType$ = "Adjust Bill Down on: " + BillNum$
  Case 14
    BillNum$ = ParseBillNum$(TaxTRec.Description)
    BillType$ = "Adjust Bill Up on: " + BillNum$
  Case 21
    BillNum$ = ParseBillNum$(TaxTRec.Description)
    BillType$ = "Paid Bill Plus Prepay on: " + BillNum$
  Case 22
    BillType$ = "Prepayment"
  Case 10
    BillNum$ = ParseBillNum$(TaxTRec.Description)
    BillType = "Adj Pay Down Affecting Credit on: " + BillNum$
  Case 11
    BillType = "Adjust Prepay Down"
  Case 12
    BillType = "Refund Prepay"
  Case 24
    BillNum$ = ParseBillNum$(TaxTRec.Description)
    BillType = "Adjust Bill Up Affecting Credit on: " + BillNum$
  Case Else
    BillType$ = Str$(TaxTRec.TranType) + "??"
    
  End Select
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxCustInfoTHist", "LoadMe", Erl)
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

Private Sub fpList1_DblClick()
  Dim TRHandle As Integer
  Dim NumOfTaxTRecs As Long
  Dim TaxTRec As TaxTransactionType
  Dim ThisRec As Long
  On Error GoTo ERRORSTUFF
  
  frmTaxCustInfoTHist.fpList1.col = 3
  frmTaxCustInfoTHist.fpList1.Row = frmTaxCustInfoTHist.fpList1.ListIndex
  ThisRec = CLng(frmTaxCustInfoTHist.fpList1.ColText)
  
  If ThisRec > 0 Then
    OpenTaxTransFile TRHandle, NumOfTaxTRecs
    Get TRHandle, ThisRec, TaxTRec
    Close
    If TaxTRec.TranType = 1 Then
      frmTaxTransDetail.Show vbModal
    Else
      frmTaxTransDetailNotBill.Show vbModal
    End If
  Else
    frmTaxTransDetailNotBill.Show vbModal
  End If
Exit Sub
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxCustInfoTHist", "fpList1_DblClick", Erl)
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


