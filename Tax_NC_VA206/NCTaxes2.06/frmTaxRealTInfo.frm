VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmTaxRealTInfo 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Real Property Transaction History"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTaxRealTInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpList1 
      Height          =   4050
      Left            =   1260
      TabIndex        =   0
      Top             =   2655
      Width           =   9135
      _Version        =   196608
      _ExtentX        =   16113
      _ExtentY        =   7144
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
      ColDesigner     =   "frmTaxRealTInfo.frx":08CA
   End
   Begin EditLib.fpCurrency fpCurrBalance 
      Height          =   375
      Left            =   5691
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1815
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
      AlignTextH      =   1
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
      Appearance      =   0
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
      Left            =   3528
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7812
      Width           =   1944
      _Version        =   131072
      _ExtentX        =   3429
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
      ButtonDesigner  =   "frmTaxRealTInfo.frx":0D1B
   End
   Begin EditLib.fpText fptxtThisPin 
      Height          =   390
      Left            =   2865
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   "Enter the official name of your town here. For example, 'Town Of Washington'."
      Top             =   735
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
      Left            =   6168
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7812
      Width           =   1944
      _Version        =   131072
      _ExtentX        =   3429
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
      ButtonDesigner  =   "frmTaxRealTInfo.frx":0EF9
   End
   Begin EditLib.fpText fptxtMessage 
      Height          =   390
      Left            =   593
      TabIndex        =   8
      TabStop         =   0   'False
      Tag             =   "Enter the official name of your town here. For example, 'Town Of Washington'."
      Top             =   1320
      Visible         =   0   'False
      Width           =   10455
      _Version        =   196608
      _ExtentX        =   18441
      _ExtentY        =   688
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8454143
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
      MaxLength       =   150
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
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   900
      Index           =   1
      Left            =   1485
      Top             =   315
      Width           =   8655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Real Property Information"
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
      Left            =   2925
      TabIndex        =   7
      Top             =   375
      Width           =   6015
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
      Height          =   372
      Left            =   4014
      TabIndex        =   6
      Top             =   1932
      Width           =   1452
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   4815
      Left            =   1080
      Top             =   2415
      Width           =   9495
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
      Left            =   1073
      TabIndex        =   5
      Top             =   7335
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   960
      Left            =   1485
      Top             =   255
      Width           =   8655
   End
End
Attribute VB_Name = "frmTaxRealTInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim PropPin$

Private Sub cmdDetail_Click()
  Dim TRHandle As Integer
  Dim NumOfTaxTRecs As Long
  Dim TaxTRec As TaxTransactionType
  Dim ThisRec As Long
  
  frmTaxRealTInfo.fpList1.Col = 3
  frmTaxRealTInfo.fpList1.Row = frmTaxRealTInfo.fpList1.ListIndex
  ThisRec = CLng(frmTaxRealTInfo.fpList1.ColText)
  
  If ThisRec > 0 Then
    OpenTaxTransFile TRHandle, NumOfTaxTRecs
    Get TRHandle, ThisRec, TaxTRec
    Close
    If TaxTRec.TranType = 1 Then
      frmTaxRealTransDetail.Show vbModal
    Else
      frmTaxRealTransDetailNotBill.Show vbModal
    End If
  Else
    frmTaxRealTransDetailNotBill.Show vbModal
  End If
  
End Sub

Private Sub cmdExit_Click()
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
  DoEvents
  Call LoadMe
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxRealTInfo.")
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
  Dim CDateStr$
  Dim AHandle As Integer
  
  'on error goto ERRORSTUFF
  
  DoEvents
  If Exist("cnvtdate.dat") Then
    AHandle = FreeFile
    Open "cnvtdate.dat" For Input As AHandle
    Input #AHandle, CDateStr$
    Close AHandle
    CDateStr$ = MakeRegDate(CInt(CDateStr$))
    fptxtMessage.Visible = True
    fptxtMessage.Text = "Only transactions affecting tax bills occurring on or after " + CDateStr$ + " are reported."
  End If
  
  PropPin = QPTrim$(frmTaxRealProp.fptxtRealPin.Text)
  fptxtThisPin = PropPin
  TCnt = 0
  fpCurrBalance = GetRealBalance(PropPin)
  OpenTaxTransFile TRHandle, NumOfTaxTRecs
  
  For x = NumOfTaxTRecs To 1 Step -1
    Get TRHandle, x, TaxTRec
    If QPTrim$(TaxTRec.RealPin) = PropPin Then
      TCnt = TCnt + 1
'      TaxTRec.RealPin = "37777" Yadkinville
'      Put TRHandle, x, TaxTRec
      ThisDate = MakeRegDate(TaxTRec.TransDate)
      GoSub GetTransType
      ThisType = BillType
'      ThisAmt = OldRound(TaxTRec.Amount + TaxTRec.DiscAmt)'8/11/08
'      If ThisType = "Credit Applied at Billing" Then
'        ThisAmt = TaxTRec.Revenue.PrePaidUsed
'      End If
      If ThisType = "Credit Applied at Billing" Then '8/11/08
        ThisAmt = TaxTRec.Revenue.PrePaidUsed
      ElseIf TaxTRec.TranType = 1 Then
        ThisAmt = OldRound(TaxTRec.Amount)
      Else
        ThisAmt = OldRound(TaxTRec.Amount + TaxTRec.DiscAmt)
      End If
      fpList1.InsertRow = ThisDate & Chr(9) & ThisType & Chr(9) & Using$("$##,###,##0.00", ThisAmt) & Chr(9) & Using$("##########", x)
    End If
  Next x
  Close
  fpList1.ListIndex = 0
  
  Label3.Caption = "Transaction Count: " + CStr(TCnt)
  Unload frmLoadingRpt
  
  Exit Sub
  
GetTransType:
  Select Case TaxTRec.TranType
  Case 1
    Select Case TaxTRec.BillType
    Case "R"
      BillType$ = "Real-Estate Bill"
    Case "P"
      BillType$ = "Personal Property Bill"
    Case "C"
      BillType$ = "Combined Bill"
    Case "M"
      BillType$ = "Manual Bill"
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
    BillType$ = "Release"
  Case 4
    BillType$ = "Interest"
    Interest# = TaxTRec.Revenue.Interest#
  Case 6
    BillType$ = "Collection/Ad Cost"
  Case 7
    BillType$ = "Adjust Paid Down"
  Case 9
    BillType$ = "Credit Applied at Billing"
  Case 13
    BillType$ = "Adjust Bill Down"
  Case 14
    BillType$ = "Adjust Bill Up"
  Case 21
    BillNum$ = ParseBillNum$(TaxTRec.Description)
    BillType$ = "Paid Bill Plus Prepay"
  Case 22
    BillType$ = "Prepayment"
  Case 10
    BillType = "Adjust Pay Dwn Affecting Credit"
  Case 24
    BillType = "Adjust Bill Up Affecting Credit"
  Case 11
    BillType = "Adjust Prepay Down" 'added 1/29/08
  Case Else
    BillType$ = Str$(TaxTRec.TranType) + "??"
    
  End Select

  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxRealTInfo", "LoadMe", Erl)
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

Private Sub fpList1_DblClick()
  Dim TRHandle As Integer
  Dim NumOfTaxTRecs As Long
  Dim TaxTRec As TaxTransactionType
  Dim ThisRec As Long
  
  'on error goto ERRORSTUFF
  
  frmTaxRealTInfo.fpList1.Col = 3
  frmTaxRealTInfo.fpList1.Row = frmTaxRealTInfo.fpList1.ListIndex
  ThisRec = CLng(frmTaxRealTInfo.fpList1.ColText)
  
  If ThisRec > 0 Then
    OpenTaxTransFile TRHandle, NumOfTaxTRecs
    Get TRHandle, ThisRec, TaxTRec
    Close
    If TaxTRec.TranType = 1 Then
      frmTaxRealTransDetail.Show vbModal
    Else
      frmTaxRealTransDetailNotBill.Show vbModal
    End If
  Else
    frmTaxRealTransDetailNotBill.Show vbModal
  End If
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxRealTInfo", "fpList1_DblClick", Erl)
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
