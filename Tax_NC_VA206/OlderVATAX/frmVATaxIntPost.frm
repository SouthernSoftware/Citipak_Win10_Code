VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxInterestPost 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Interest Calculation Post"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxIntPost.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbType 
      Height          =   384
      Left            =   4134
      TabIndex        =   6
      Top             =   2160
      Width           =   3372
      _Version        =   196608
      _ExtentX        =   5948
      _ExtentY        =   677
      Text            =   ""
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
      Text            =   ""
      Columns         =   0
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
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
      DataFieldList   =   ""
      ColumnEdit      =   -1
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
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
      ListPosition    =   0
      ButtonThreeDAppearance=   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      AutoSearchFill  =   0   'False
      AutoSearchFillDelay=   500
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmVATaxIntPost.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPost 
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7320
      Width           =   1800
      _Version        =   131072
      _ExtentX        =   3175
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmVATaxIntPost.frx":0BC1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7320
      Width           =   1800
      _Version        =   131072
      _ExtentX        =   3175
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmVATaxIntPost.frx":0D9D
   End
   Begin EditLib.fpDateTime fptxtRealYr 
      Height          =   372
      Left            =   6480
      TabIndex        =   7
      Top             =   3120
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
      _ExtentY        =   656
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
      ThreeDInsideHighlightColor=   -2147483637
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
      ButtonStyle     =   2
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
      InvalidColor    =   -2147483643
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483643
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   1
      Text            =   "2018"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19800101"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "20010101"
      TimeDefault     =   "000000"
      TimeStyle       =   0
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      PopUpType       =   1
      DateCalcY2KSplit=   60
      CaretPosition   =   0
      IncYear         =   1
      IncMonth        =   1
      IncDay          =   1
      IncHour         =   1
      IncMinute       =   1
      IncSecond       =   1
      ButtonColor     =   13684944
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpDateTime fptxtPersYr 
      Height          =   372
      Left            =   6480
      TabIndex        =   8
      Top             =   3720
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
      _ExtentY        =   656
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
      ThreeDInsideHighlightColor=   -2147483637
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
      ButtonStyle     =   2
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
      InvalidColor    =   -2147483643
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483643
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   1
      Text            =   "2018"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19800101"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "20010101"
      TimeDefault     =   "000000"
      TimeStyle       =   0
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      PopUpType       =   1
      DateCalcY2KSplit=   60
      CaretPosition   =   0
      IncYear         =   1
      IncMonth        =   1
      IncDay          =   1
      IncHour         =   1
      IncMinute       =   1
      IncSecond       =   1
      ButtonColor     =   13684944
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Make A Post Selection"
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
      Height          =   276
      Left            =   4614
      TabIndex        =   11
      Top             =   1800
      Width           =   2340
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1572
      Left            =   3774
      Top             =   2880
      Width           =   4092
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Year For Real:"
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
      Height          =   276
      Left            =   4440
      TabIndex        =   10
      Top             =   3228
      Width           =   1860
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Year For Personal:"
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
      Height          =   276
      Left            =   3960
      TabIndex        =   9
      Top             =   3828
      Width           =   2340
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   840
      Left            =   2304
      Top             =   708
      Width           =   7020
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Interest Calculations Post"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3828
      TabIndex        =   3
      Top             =   948
      Width           =   4020
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   2052
      Left            =   2040
      Top             =   4896
      Width           =   7572
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Press ESC To Exit."
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
      Height          =   372
      Left            =   5808
      TabIndex        =   1
      Top             =   6216
      Width           =   3132
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Press F10 To Post."
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
      Height          =   372
      Left            =   2688
      TabIndex        =   0
      Top             =   6216
      Width           =   3132
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   996
      Left            =   2316
      Top             =   588
      Width           =   7020
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   $"frmVATaxIntPost.frx":0F79
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   2052
      Left            =   2040
      TabIndex        =   2
      Top             =   4896
      Width           =   7572
   End
End
Attribute VB_Name = "frmVATaxInterestPost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class
  Dim IncReal As Boolean
  Dim IncPers As Boolean
  Dim CurTaxYear As Integer

Private Sub cmdExit_Click()
  frmVATaxInterestMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPost_Click()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim cnt As Long, Previous&
  Dim DidSome As Long
  Dim TaxTrans As TaxTransactionType
  Dim NewTaxTrans As TaxTransactionType
  Dim ClearVATaxTrans As TaxTransactionType
  Dim TaxIntRec As InterestRecType
  Dim IntTrans As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, NextRecord&
  Dim IntDateRec As TaxInterestDateType
  Dim IDHandle As Integer
  
  On Error GoTo ERRORSTUFF
  If TaxMsgWOpts(900, "If you are sure you are ready to post then press F10 to continue. Otherwise, press ESC to abort the post attempt.", "F10 Continue", "ESC Abort") = "abort" Then
    Unload frmVATaxMsgWOpts
    Call TaxMsg(900, "Post attempt aborted.")
    Close
    Exit Sub
  Else
    Unload frmVATaxMsgWOpts
    MainLog ("Interest calculations posted.")
  End If
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
'  KillFile "TAXINTCK.DAT"
  OpenTxIntTickFile IDHandle
  If fpcmbType.Text = "REAL ONLY" Or fpcmbType.Text = "POST BOTH" Then
    frmVATaxShowPctComp.Label1 = "Posting Real Interest"
    frmVATaxShowPctComp.Show , Me
    frmVATaxShowPctComp.cmdCancel.Visible = False
    cmdPost.Enabled = False
    cmdExit.Enabled = False
    OpenRInterestRecFile IRHandle, NumOfIRRecs
    For cnt& = 1 To NumOfIRRecs
      Get IRHandle, cnt&, TaxIntRec
      If TaxIntRec.DelFlag = 0 Then
        'Update the Bill transaction first
       'TaxIntRec(1).BillRec
        Get TTHandle, TaxIntRec.BillRec, TaxTrans 'get bill trans
        If TaxIntRec.Amount = 0 Then GoTo SkipIt 'edited to zero
        TaxTrans.Revenue.Interest = OldRound#(TaxTrans.Revenue.Interest + TaxIntRec.Amount)
        Put #TTHandle, TaxIntRec.BillRec, TaxTrans 'put it back
      'Now make a new clean transaction
        NewTaxTrans = ClearVATaxTrans
        NewTaxTrans.TransDate = Date2Num%(Date$)
        NewTaxTrans.TaxYear = TaxIntRec.TaxYear
        NewTaxTrans.TranType = 4       '4=Interest
        NewTaxTrans.BillType = "R"     'R=Real P=Personal Property C=Combined (NC/GA)
        NewTaxTrans.Amount = TaxIntRec.Amount  'Total Transaction Amount
        NewTaxTrans.Revenue.Interest = TaxIntRec.Amount
        NewTaxTrans.Description = "Tax Int on Bill# " + QPTrim$(TaxIntRec.BillNumber)
        NewTaxTrans.Posted2GL = "N"
        NewTaxTrans.CustomerRec = TaxIntRec.CustRec
        NewTaxTrans.CustPin = TaxIntRec.CustPin
        NewTaxTrans.RealPin = TaxIntRec.RealPin
        NewTaxTrans.PersPin = 0
        NewTaxTrans.LastTrans = 0
        NewTaxTrans.BelongTo = TaxIntRec.BillRec
        NewTaxTrans.Revenue.PrePaidAmt = 0
        NewTaxTrans.Revenue.PrePaidBal = OldRound(GetOverPayBalance(TaxIntRec.CustRec, "R"))
        NewTaxTrans.Revenue.PrePaidUsed = 0
        NewTaxTrans.OperNum = OperNum
        LSet NewTaxTrans.Padding = ""
      'Increment Transaction File Record Count
        NextRecord& = (LOF(TTHandle) / Len(NewTaxTrans)) + 1
        Put TTHandle, NextRecord&, NewTaxTrans
      'Update the Customer Pointers Now
        Get TCHandle, TaxIntRec.CustRec, TaxCust
      
        If TaxCust.LastTrans = 0 Then
          TaxCust.LastTrans = NextRecord&
          Put TCHandle, TaxIntRec.CustRec, TaxCust
        Else
          Previous& = TaxCust.LastTrans
          TaxCust.LastTrans = NextRecord&
          Put TCHandle, TaxIntRec.CustRec, TaxCust
          Get TTHandle, NextRecord&, NewTaxTrans
          NewTaxTrans.LastTrans = Previous&
          Put TTHandle, NextRecord&, NewTaxTrans
        End If
      End If
SkipIt:
      frmVATaxShowPctComp.ShowPctComp cnt, NumOfIRRecs
    Next cnt
    Close IRHandle
    KillFile TaxRIntFile
    Get IDHandle, 1, IntDateRec
    IntDateRec.RInterestDate = Date2Num%(Date$)
'    IntDateRec.PInterestDate = 0
    Put IDHandle, 1, IntDateRec
  End If
  
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdPost.Enabled = True
  cmdExit.Enabled = True
  
  If fpcmbType.Text = "PERSONAL ONLY" Or fpcmbType.Text = "POST BOTH" Then
    frmVATaxShowPctComp.Label1 = "Posting Personal Interest"
    frmVATaxShowPctComp.Show , Me
    frmVATaxShowPctComp.cmdCancel.Visible = False
    cmdPost.Enabled = False
    cmdExit.Enabled = False
    OpenPInterestRecFile IRHandle, NumOfIRRecs
    For cnt& = 1 To NumOfIRRecs
      Get IRHandle, cnt&, TaxIntRec
      If TaxIntRec.DelFlag = 0 Then
        'Update the Bill transaction first
       'TaxIntRec(1).BillRec
        Get TTHandle, TaxIntRec.BillRec, TaxTrans 'get bill trans
        TaxTrans.Revenue.Interest = OldRound#(TaxTrans.Revenue.Interest + TaxIntRec.Amount)
        Put #TTHandle, TaxIntRec.BillRec, TaxTrans 'put it back
      'Now make a new clean transaction
        NewTaxTrans = ClearVATaxTrans
        NewTaxTrans.TransDate = Date2Num%(Date$)
        NewTaxTrans.TaxYear = TaxIntRec.TaxYear
        NewTaxTrans.TranType = 4       '4=Interest
        NewTaxTrans.BillType = "P"     'R=Real P=Personal Property C=Combined (NC/GA)
        NewTaxTrans.Amount = TaxIntRec.Amount  'Total Transaction Amount
        NewTaxTrans.Revenue.Interest = TaxIntRec.Amount
        NewTaxTrans.Description = "Tax Int on Bill# " + QPTrim$(TaxIntRec.BillNumber)
        NewTaxTrans.Posted2GL = "N"
        NewTaxTrans.CustomerRec = TaxIntRec.CustRec
        NewTaxTrans.CustPin = TaxIntRec.CustPin
        NewTaxTrans.RealPin = 0
        NewTaxTrans.PersPin = TaxIntRec.PersPin
        NewTaxTrans.LastTrans = 0
        NewTaxTrans.BelongTo = TaxIntRec.BillRec
        NewTaxTrans.Revenue.PrePaidAmt = 0
        NewTaxTrans.Revenue.PrePaidBal = OldRound(GetOverPayBalance(TaxIntRec.CustRec, "P"))
        NewTaxTrans.Revenue.PrePaidUsed = 0
        NewTaxTrans.OperNum = OperNum
        LSet NewTaxTrans.Padding = ""
      'Increment Transaction File Record Count
        NextRecord& = (LOF(TTHandle) / Len(NewTaxTrans)) + 1
        Put TTHandle, NextRecord&, NewTaxTrans
      'Update the Customer Pointers Now
        Get TCHandle, TaxIntRec.CustRec, TaxCust
      
        If TaxCust.LastTrans = 0 Then
          TaxCust.LastTrans = NextRecord&
          Put TCHandle, TaxIntRec.CustRec, TaxCust
        Else
          Previous& = TaxCust.LastTrans
          TaxCust.LastTrans = NextRecord&
          Put TCHandle, TaxIntRec.CustRec, TaxCust
          Get TTHandle, NextRecord&, NewTaxTrans
          NewTaxTrans.LastTrans = Previous&
          Put TTHandle, NextRecord&, NewTaxTrans
        End If
      End If
      frmVATaxShowPctComp.ShowPctComp cnt, NumOfIRRecs
    Next cnt
    Close IRHandle
    KillFile TaxPIntFile
    Get IDHandle, 1, IntDateRec
'    IntDateRec.RInterestDate = 0
    IntDateRec.PInterestDate = Date2Num%(Date$)
    Put IDHandle, 1, IntDateRec
  End If
  
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdPost.Enabled = True
  cmdExit.Enabled = True
  Close
  
  'Now Delete the Tax Bill File so Duplicate's Cannot Be Reproduced
  
  If fpcmbType.Text = "REAL ONLY" Then
    Call Savemsg(900, "The real interest calculations have been posted successfully.")
  ElseIf fpcmbType.Text = "PERSONAL ONLY" Then
    Call Savemsg(900, "The personal interest calculations have been posted successfully.")
  ElseIf fpcmbType.Text = "POST BOTH" Then
    Call Savemsg(900, "Both real and personal interest calculations have been posted successfully.")
  End If
  
  Call cmdExit_Click
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxInterestPost", "cmdPost_Click", Erl)
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

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
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
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxInterestPost.")
      Call Terminate
      End
    End If
  End If

End Sub
'Private Sub Form_Resize()
'  If Me.WindowState <> vbMinimized Then
'    Me.Visible = False
'    'Temp_Class.ResizeControls Me
'    Me.Visible = True
'    Me.SetFocus
'    DoEvents
'  End If
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdPost_Click
      KeyCode = 0
  End Select

End Sub

Private Sub LoadMe()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim IntTrans As InterestRecType
  Dim NumOfRICRecs As Long
  Dim RICHandle As Integer
  Dim NumOfPICRecs As Long
  Dim PICHandle As Integer
  Dim x As Long, y As Long
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  NumOfRICRecs = 0
  NumOfPICRecs = 0
  IncReal = False
  IncPers = False
  If Exist(TaxRIntFile) Then
    OpenRInterestRecFile RICHandle, NumOfRICRecs
    IncReal = True
  End If
  If Exist(TaxPIntFile) Then
    OpenPInterestRecFile PICHandle, NumOfPICRecs
    IncPers = True
  End If
  
  For x = 1 To NumOfRICRecs
    Get RICHandle, x, IntTrans
    If IntTrans.DelFlag = False Then
      fptxtRealYr.Text = CStr(IntTrans.CurYear)
      Exit For
    End If
  Next x
  For y = 1 To NumOfPICRecs
    Get PICHandle, y, IntTrans
    If IntTrans.DelFlag = False Then
      fptxtPersYr.Text = CStr(IntTrans.CurYear)
      Exit For
    End If
  Next y
  
  If NumOfRICRecs > 0 And NumOfPICRecs > 0 Then
    If x <= NumOfRICRecs And y <= NumOfPICRecs Then
      fpcmbType.Text = "REAL ONLY"
      fpcmbType.AddItem "REAL ONLY"
      fpcmbType.AddItem "PERSONAL ONLY"
      fpcmbType.AddItem "POST BOTH"
      fptxtPersYr.Text = CStr(TaxMasterRec.PTaxYear)
      fptxtRealYr.Text = CStr(TaxMasterRec.RTaxYear)
    ElseIf x <= NumOfRICRecs And y > NumOfPICRecs Then
      fpcmbType.Text = "REAL ONLY"
      fpcmbType.AddItem "REAL ONLY"
      fpcmbType.AddItem "NO PERSONAL"
      fptxtPersYr.Text = "NA"
      fptxtRealYr.Text = CStr(TaxMasterRec.RTaxYear)
      IncPers = False
    ElseIf x > NumOfRICRecs And y <= NumOfPICRecs Then
      fpcmbType.Text = "PERSONAL ONLY"
      fpcmbType.AddItem "NO REAL"
      fpcmbType.AddItem "PERSONAL ONLY"
      fptxtRealYr.Text = "NA"
      fptxtPersYr.Text = CStr(TaxMasterRec.PTaxYear)
      IncReal = False
    End If
  ElseIf NumOfRICRecs > 0 And NumOfPICRecs = 0 Then
    If x <= NumOfRICRecs Then
      fpcmbType.Text = "REAL ONLY"
      fpcmbType.AddItem "REAL ONLY"
      fpcmbType.AddItem "NO PERSONAL"
      fptxtPersYr.Text = "NA"
      IncPers = False
      fptxtRealYr.Text = CStr(TaxMasterRec.RTaxYear)
    End If
  ElseIf NumOfRICRecs = 0 And NumOfPICRecs > 0 Then
    If y <= NumOfPICRecs Then
      fpcmbType.Text = "PERSONAL ONLY"
      fpcmbType.AddItem "NO REAL"
      fpcmbType.AddItem "PERSONAL ONLY"
      fptxtRealYr.Text = "NA"
      fptxtPersYr.Text = CStr(TaxMasterRec.PTaxYear)
      IncReal = False
    End If
  End If

  Close RICHandle
  Close PICHandle
  
'  fptxtCurrYear.Text = CStr(TaxMasterRec.RTaxYear)
'  CurTaxYear = TaxMasterRec.RTaxYear
  
End Sub

