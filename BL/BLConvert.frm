VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLConvertMain 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Conversion"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "BLConvert.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbVersion 
      Height          =   405
      Left            =   3930
      TabIndex        =   1
      Top             =   5160
      Width           =   2895
      _Version        =   196608
      _ExtentX        =   5106
      _ExtentY        =   714
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
      Style           =   0
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
      ColDesigner     =   "BLConvert.frx":08CA
   End
   Begin LpLib.fpCombo fpcmbCatVersion 
      Height          =   405
      Left            =   3930
      TabIndex        =   0
      Top             =   2070
      Width           =   2895
      _Version        =   196608
      _ExtentX        =   5106
      _ExtentY        =   714
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
      Style           =   0
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
      ColDesigner     =   "BLConvert.frx":0BC1
   End
   Begin VB.CommandButton cmdChangeCat2Nums 
      Caption         =   "F6 Change Category Codes To &Numbers"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   2688
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2730
      Width           =   6108
   End
   Begin VB.CheckBox chkCategories 
      BackColor       =   &H00800000&
      Caption         =   "Convert Category Data "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   492
      Left            =   4560
      TabIndex        =   2
      Top             =   4125
      Width           =   2748
   End
   Begin VB.CommandButton cmdCatVersion 
      Caption         =   "F5 Print Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   6960
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2070
      Width           =   3564
   End
   Begin VB.CommandButton cmdPrintVersion 
      Caption         =   "F4 Print Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   6960
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5160
      Width           =   3564
   End
   Begin VB.CommandButton cmdCustOnly 
      Caption         =   "F3 Help For Customer C&onversion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Left            =   3936
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5730
      Width           =   3804
   End
   Begin VB.CommandButton cmdHelpCustCat 
      Caption         =   "F2 &Help For Category Conversion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Left            =   3936
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3405
      Width           =   3804
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "ESC &Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1200
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7605
      Width           =   2364
   End
   Begin VB.CommandButton cmdBegin 
      Caption         =   "F10 &Begin"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5160
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7605
      Width           =   5370
   End
   Begin VB.CheckBox chkCustomers 
      BackColor       =   &H00800000&
      Caption         =   "Convert Customer Data (Zeros All Customer Balances)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   492
      Left            =   2784
      TabIndex        =   3
      Top             =   6450
      Width           =   6108
   End
   Begin EditLib.fpText fptxtMarque 
      Height          =   690
      Left            =   2115
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1125
      Width           =   7350
      _Version        =   196608
      _ExtentX        =   12975
      _ExtentY        =   1206
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
      BackColor       =   192
      ForeColor       =   -2147483643
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
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2025
      Left            =   870
      Top             =   5070
      Width           =   9855
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2940
      Left            =   870
      Top             =   1965
      Width           =   9855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Category Version"
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
      Height          =   345
      Left            =   720
      TabIndex        =   13
      Top             =   2160
      Width           =   2985
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Customer Version"
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
      Height          =   345
      Left            =   720
      TabIndex        =   10
      Top             =   5250
      Width           =   2985
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   636
      Left            =   2892
      Top             =   336
      Width           =   5868
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BUSINESS LICENSE CONVERSION"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   348
      Left            =   3468
      TabIndex        =   7
      Top             =   432
      Width           =   4908
   End
End
Attribute VB_Name = "frmBLConvertMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class
  Dim MakeCBalTBalFlag As Boolean
  Dim CheckCustFlag As Boolean
  Dim CheckCatFlag As Boolean
  Dim NOGLACCT As Boolean

Private Sub ConvertCust1()
  Dim DosNumOfCustRecs2 As Integer
  Dim DosCustRec2 As DOSARCustRecType2
  Dim DosCustHandle2 As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim DosNumOfCustRecs As Integer
  Dim DosCustRec As DosARCustRecType
  Dim DosCustHandle As Integer
  Dim x As Integer, y As Integer
  Dim Nextx As Integer
  Dim DosCatRec As DosARNewCatCodeRecType
  Dim DosCatHandle As Integer
  Dim NumOfCatRecs As Integer
  Dim CatRec As ARNewCatCodeRecType
  Dim CatHandle As Integer
  Dim PctDone As Integer
  
  If Not Exist("arcust.dat") Then
    MsgBox "You have elected to convert customer data. However, the file 'arcust.dat' could not be found. Conversion aborted."
    Exit Sub
  End If
  
  cmdBegin.Enabled = False
  cmdExit.Enabled = False
  
  fptxtMarque.Text = "Converting Customer Data"
  DoEvents

  OpenDosCustFile DosCustHandle
  DosNumOfCustRecs = LOF(DosCustHandle) / Len(DosCustRec)
  If DosNumOfCustRecs = 0 Then
    MsgBox "No customers on file."
    Close
    Exit Sub
  End If

  ReDim TempCUSTNUMB(1 To DosNumOfCustRecs) As String * 10
  ReDim TempSORTNAME(1 To DosNumOfCustRecs) As String * 10
  ReDim TempBILLNAME(1 To DosNumOfCustRecs) As String * 35
  ReDim TempADDRESS1(1 To DosNumOfCustRecs) As String * 35
  ReDim TempADDRESS2(1 To DosNumOfCustRecs) As String * 35
  ReDim TempCITY(1 To DosNumOfCustRecs) As String * 20
  ReDim TempSTATE(1 To DosNumOfCustRecs) As String * 2
  ReDim TempZIPCODE(1 To DosNumOfCustRecs) As String * 10
  ReDim TempCustName(1 To DosNumOfCustRecs) As String * 35
  ReDim TempContact(1 To DosNumOfCustRecs) As String * 30

  ReDim TempBILLCAT1(1 To DosNumOfCustRecs) As String * 5
  ReDim TempDESC1(1 To DosNumOfCustRecs) As String * 15
  ReDim TempREV1(1 To DosNumOfCustRecs) As Long
  ReDim TempFee1(1 To DosNumOfCustRecs) As Double
  ReDim TempBILLCAT2(1 To DosNumOfCustRecs) As String * 5
  ReDim TempDESC2(1 To DosNumOfCustRecs) As String * 15
  ReDim TempREV2(1 To DosNumOfCustRecs) As Long
  ReDim TempFee2(1 To DosNumOfCustRecs) As Double
  ReDim TempBILLCAT3(1 To DosNumOfCustRecs) As String * 5
  ReDim TempDESC3(1 To DosNumOfCustRecs) As String * 15
  ReDim TempREV3(1 To DosNumOfCustRecs) As Long
  ReDim TempFee3(1 To DosNumOfCustRecs) As Double
  ReDim TempBILLCAT4(1 To DosNumOfCustRecs) As String * 5
  ReDim TempDESC4(1 To DosNumOfCustRecs) As String * 15
  ReDim TempREV4(1 To DosNumOfCustRecs) As Long
  ReDim TempFee4(1 To DosNumOfCustRecs) As Double
  ReDim TempBILLCAT5(1 To DosNumOfCustRecs) As String * 5
  ReDim TempDESC5(1 To DosNumOfCustRecs) As String * 15
  ReDim TempREV5(1 To DosNumOfCustRecs) As Long
  ReDim TempFee5(1 To DosNumOfCustRecs) As Double

'************
  ReDim TempIssuanceFee(1 To DosNumOfCustRecs) As Double
  ReDim TempCustLocation(1 To DosNumOfCustRecs) As String * 1
  ReDim TempWPHONE(1 To DosNumOfCustRecs) As String * 14
  ReDim TempFeeAmt(1 To DosNumOfCustRecs) As Double
  ReDim TempLICENSE(1 To DosNumOfCustRecs) As String * 12
  ReDim TempVALID(1 To DosNumOfCustRecs) As Integer
  ReDim TempInactive(1 To DosNumOfCustRecs) As String * 1        '"Y" if account is inactive
  ReDim TempProrate(1 To DosNumOfCustRecs) As Integer            'prorate percentage
  ReDim TempAcctBal(1 To DosNumOfCustRecs) As Double
  ReDim TempIssueLicense(1 To DosNumOfCustRecs) As String * 1    'y/n
  ReDim TempDeleted(1 To DosNumOfCustRecs) As String * 1         '(yY)=deleted, anything else isn't
  ReDim TempFirstTrans(1 To DosNumOfCustRecs) As Long
  ReDim TempLastTrans(1 To DosNumOfCustRecs) As Long

  ReDim TempLicBal(1 To DosNumOfCustRecs) As Double
  ReDim TempFeeBal(1 To DosNumOfCustRecs) As Double
  ReDim TempPenBal(1 To DosNumOfCustRecs) As Double
  ReDim TempRoomtoGrow(1 To DosNumOfCustRecs) As String * 136
  ReDim TempChkByte(1 To DosNumOfCustRecs) As String * 1

  For x = 1 To DosNumOfCustRecs
    Get DosCustHandle, x, DosCustRec
    TempCUSTNUMB(x) = DosCustRec.CUSTNUMB
    TempSORTNAME(x) = DosCustRec.SORTNAME
    TempBILLNAME(x) = DosCustRec.BILLNAME
    TempADDRESS1(x) = DosCustRec.ADDRESS1
    TempADDRESS2(x) = DosCustRec.ADDRESS2
    TempCITY(x) = DosCustRec.CITY
    TempSTATE(x) = DosCustRec.STATE
    TempZIPCODE(x) = DosCustRec.ZIPCODE
    TempCustName(x) = DosCustRec.CustName
    TempContact(x) = DosCustRec.Contact
    TempBILLCAT1(x) = DosCustRec.BILLCAT1
    TempDESC1(x) = DosCustRec.DESC1
    TempREV1(x) = DosCustRec.REV1
    TempFee1(x) = DosCustRec.Fee1
    TempBILLCAT2(x) = DosCustRec.BILLCAT2
    TempDESC2(x) = DosCustRec.DESC2
    TempREV2(x) = DosCustRec.REV2
    TempFee2(x) = DosCustRec.Fee2
    TempBILLCAT3(x) = DosCustRec.BILLCAT3
    TempDESC3(x) = DosCustRec.DESC3
    TempREV3(x) = DosCustRec.REV3
    TempFee3(x) = DosCustRec.Fee3
    TempBILLCAT4(x) = DosCustRec.BILLCAT4
    TempDESC4(x) = DosCustRec.DESC4
    TempREV4(x) = DosCustRec.REV4
    TempFee4(x) = DosCustRec.Fee4
    TempBILLCAT5(x) = DosCustRec.BILLCAT5
    TempDESC5(x) = DosCustRec.DESC5
    TempREV5(x) = DosCustRec.REV5
    TempFee5(x) = DosCustRec.Fee5
    TempIssuanceFee(x) = DosCustRec.IssuanceFee
    TempCustLocation(x) = DosCustRec.CustLocation
    TempWPHONE(x) = DosCustRec.WPHONE
    TempFeeAmt(x) = DosCustRec.FeeAmt
    TempLICENSE(x) = DosCustRec.LICENSE
    TempVALID(x) = DosCustRec.VALID
    If QPTrim$(DosCustRec.Inactive) = "Y" Then
      TempInactive(x) = DosCustRec.Inactive
    Else
      TempInactive(x) = "N"
    End If
    TempProrate(x) = DosCustRec.Prorate
    If InStr(DosCustRec.AcctBal, "E") Then DosCustRec.AcctBal = 0
    TempAcctBal(x) = DosCustRec.AcctBal
    TempIssueLicense(x) = DosCustRec.IssueLicense
    If QPTrim$(DosCustRec.Deleted) <> "Y" Then DosCustRec.Deleted = "N"
    TempDeleted(x) = DosCustRec.Deleted
    TempFirstTrans(x) = DosCustRec.FirstTrans
    TempLastTrans(x) = DosCustRec.LastTrans
    TempLicBal(x) = DosCustRec.LicBal
    TempFeeBal(x) = DosCustRec.FeeBal
    TempPenBal(x) = DosCustRec.PenBal
    TempRoomtoGrow(x) = DosCustRec.RoomtoGrow
    TempChkByte(x) = DosCustRec.ChkByte
    DoEvents
    PctDone = (x / DosNumOfCustRecs) * 100
    fptxtMarque.Text = "Collecting old customer data is " + CStr(PctDone) + "% completed."
  Next x
  Close DosCustHandle

  OpenCustFile CustHandle

  For x = 1 To DosNumOfCustRecs
    Get CustHandle, x, CustRec
    CustRec.CUSTNUMB = QPTrim(TempCUSTNUMB(x))
    CustRec.SORTNAME = QPTrim(TempSORTNAME(x))
    CustRec.BILLNAME = QPTrim(TempBILLNAME(x))
    CustRec.ADDRESS1 = QPTrim(TempADDRESS1(x))
    CustRec.ADDRESS2 = QPTrim(TempADDRESS2(x))
    CustRec.CITY = QPTrim(TempCITY(x))
    CustRec.STATE = QPTrim(TempSTATE(x))
    CustRec.ZIPCODE = QPTrim(TempZIPCODE(x))
    CustRec.CustName = QPTrim(TempCustName(x))
    CustRec.Contact = QPTrim(TempContact(x))
    CustRec.ServAdd = ""
    CustRec.SSNFID = ""
    If chkCategories.Value = 1 Then
      CustRec.BILLCAT1 = QPTrim(TempBILLCAT1(x))
      CustRec.DESC1 = QPTrim(TempDESC1(x))
      CustRec.REV1 = TempREV1(x)
      CustRec.BILLCAT2 = QPTrim(TempBILLCAT2(x))
      CustRec.DESC2 = QPTrim(TempDESC2(x))
      CustRec.REV2 = TempREV2(x)
      CustRec.BILLCAT3 = QPTrim(TempBILLCAT3(x))
      CustRec.DESC3 = QPTrim(TempDESC3(x))
      CustRec.REV3 = TempREV3(x)
      CustRec.BILLCAT4 = QPTrim(TempBILLCAT4(x))
      CustRec.DESC4 = QPTrim(TempDESC4(x))
      CustRec.REV4 = TempREV4(x)
      CustRec.BILLCAT5 = QPTrim(TempBILLCAT5(x))
      CustRec.DESC5 = QPTrim(TempDESC5(x))
      CustRec.REV5 = TempREV5(x)
    Else
      CustRec.BILLCAT1 = ""
      CustRec.DESC1 = ""
      CustRec.REV1 = 0
      CustRec.BILLCAT2 = ""
      CustRec.DESC2 = ""
      CustRec.REV2 = 0
      CustRec.BILLCAT3 = ""
      CustRec.DESC3 = ""
      CustRec.REV3 = 0
      CustRec.BILLCAT4 = ""
      CustRec.DESC4 = ""
      CustRec.REV4 = 0
      CustRec.BILLCAT5 = ""
      CustRec.DESC5 = ""
      CustRec.REV5 = 0
    End If
    CustRec.Fee1 = 0
    'The Dos version does not have balances for each of the
    'five potential categories so in the conversion the Lic Bal
    'is automatically dumped into the first category just to have
    'a post conversion license category balance
    CustRec.FeeLicBal1 = 0
    CustRec.FeeLicPay1 = 0
    CustRec.Fee2 = 0
    CustRec.FeeLicBal2 = 0
    CustRec.FeeLicPay2 = 0
    CustRec.Fee3 = 0
    CustRec.FeeLicBal3 = 0
    CustRec.FeeLicPay3 = 0
    CustRec.Fee4 = 0
    CustRec.FeeLicBal4 = 0
    CustRec.FeeLicPay4 = 0
    CustRec.Fee5 = 0
    CustRec.FeeLicBal5 = 0
    CustRec.FeeLicPay5 = 0
    CustRec.IssuanceFee = 0
    CustRec.CustLocation = QPTrim(TempCustLocation(x))
    CustRec.WPHONE = QPTrim(TempWPHONE(x))
    CustRec.FeeAmt = 0
    CustRec.LICENSE = QPTrim(TempLICENSE(x))
    CustRec.VALID = TempVALID(x)
    CustRec.Inactive = QPTrim(TempInactive(x))
    CustRec.Prorate = TempProrate(x)
    CustRec.AcctBal = 0
    CustRec.IssueLicense = QPTrim(TempIssueLicense(x))
    CustRec.Deleted = QPTrim(TempDeleted(x))
    CustRec.FirstTrans = 0
    CustRec.LastTrans = 0
    CustRec.LicBal = 0
    CustRec.FeeBal = 0
    CustRec.PenBal = 0
    CustRec.RoomtoGrow = QPTrim(TempRoomtoGrow(x))
    CustRec.ChkByte = QPTrim(TempChkByte(x))
    CustRec.IssuanceBal = 0
    CustRec.IssuancePay = 0
    Put CustHandle, x, CustRec
    DoEvents
    PctDone = (x / DosNumOfCustRecs) * 100
    fptxtMarque.Text = "Converting customer data is " + CStr(PctDone) + "% completed."
  Next x

  On Error Resume Next
  Nextx = 1
  For x = 1 To DosNumOfCustRecs
  Get CustHandle, x, CustRec
    If QPTrim$(CustRec.BILLNAME) = "" Then
      CustRec.BILLNAME = "INVALID" + CStr(x)
      CustRec.SORTNAME = "INV" + CStr(x)
      CustRec.CustName = "INVALID" + CStr(x)
      CustRec.Contact = "INVALID" + CStr(x)
      CustRec.CITY = "INVALID"
      CustRec.CUSTNUMB = CStr(x)
      GoTo BadData
    End If

    If QPTrim$(CustRec.Inactive) <> "Y" And QPTrim$(CustRec.Inactive) <> "N" Then
      CustRec.Inactive = "N"
    End If

    If CustRec.Prorate <= 0 Or CustRec.Prorate > 100 Then
      CustRec.Prorate = 100
    End If

    If QPTrim$(CustRec.IssueLicense) <> "Y" And QPTrim$(CustRec.IssueLicense) <> "N" Then
      CustRec.IssueLicense = "N"
    End If
    
    If CustRec.VALID < 0 Or CustRec.VALID > 10000 Then
      CustRec.VALID = 0
    End If
BadData:
  Put CustHandle, x, CustRec
  Next x

  Close CustHandle

  fptxtMarque.Text = "Creating Indices"
  DoEvents

  Call CreateCustNumIdx
  Call CreateCustSearchNameIdx
  Call CreateCustNameIdx
  Call CreateLicNumIdx

End Sub

Private Sub cmdBegin_Click()
    
  If CheckData = False Then
    fptxtMarque.Text = "Conversion stopped."
    Close
    Exit Sub
  End If
  
  If chkCategories.Value = 1 Then
    If Not Exist("GLACCT.DAT") Then
      frmBLMessageBoxJrWOpts.Label1.Caption = "GLACCT.DAT could not be found. Continuing will make all GL numbers in category codes equal zero. This is because the save method for GL numbers is now based on their record numbers instead of the numbers themselves. Without the 'GLACCT.DAT' file there is no way to accurately convert category code GL numbers. Do you wish to continue?"
      frmBLMessageBoxJrWOpts.Label1.Top = 500
      frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
      frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Abort"
      frmBLMessageBoxJrWOpts.Show vbModal
      If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
        Unload frmBLMessageBoxJrWOpts
        NOGLACCT = True
        Close
        Exit Sub
      Else
        Unload frmBLMessageBoxJrWOpts
      End If
    End If
  End If
  
  frmBLMessageBoxJrWOpts.Label1.Caption = "REMINDER: ALL PRE-CONVERSION TESTS HAVE BEEN PASSED. CONVERSION CAN NOW TAKE PLACE. HOWEVER, THIS CONVERSION MAKES ALL BALANCES ZERO. ADDITIONALLY, NO TRANSACTION RELINKING TAKES PLACE. DO YOU WISH TO CONTINUE ANYWAY?"
  frmBLMessageBoxJrWOpts.Label1.Top = 500
  frmBLMessageBoxJrWOpts.Label1.Height = 1400
  frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
  frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Abort"
  frmBLMessageBoxJrWOpts.Show vbModal
  If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
    Unload frmBLMessageBoxJrWOpts
    Close
    Exit Sub
  Else
    Unload frmBLMessageBoxJrWOpts
  End If
    
  If chkCategories.Value = 1 Then
    If InStr(fpcmbCatVersion.Text, "1") Then
      Call ConvertCategory1
    ElseIf InStr(fpcmbCatVersion.Text, "2") Then
      Call ConvertCategory2
    End If
  End If
  
  If chkCustomers.Value = 1 Then
    If InStr(fpcmbVersion.Text, "1") Then
      Call ConvertCust1
    ElseIf InStr(fpcmbVersion.Text, "2") Then
      Call ConvertCust2
    End If
  End If
  
  fptxtMarque.FontSize = 14
  
  If NOGLACCT = False Then
    fptxtMarque.Text = "SUCCESS: Business License Conversion has completed."
  Else
    fptxtMarque.Text = "CONVERSION ABORTED: NO GLACCT.DAT FILE"
  End If
  
  cmdExit.Enabled = True
  
End Sub

Private Sub cmdCatVersion_Click()
  
  If InStr(fpcmbCatVersion.Text, "1") Then
    Call PrintCatVersion1
  ElseIf InStr(fpcmbCatVersion.Text, "2") Then
    Call PrintCatVersion2
  End If
End Sub

Private Sub cmdChangeCat2Nums_Click()
  
  If Clear4CatCodeProblems > 0 Then
    If Clear4CatCodeProblems = 3 Then
      frmBLMessageBoxJr.Label1.Caption = "Please examine your category code data. There are blank category codes and blank category descriptions that need attention."
      frmBLMessageBoxJr.Label1.Top = 700
      frmBLMessageBoxJr.Show vbModal
      Close
      Exit Sub
    ElseIf Clear4CatCodeProblems = 1 Then
      frmBLMessageBoxJr.Label1.Caption = "Please examine your category code data. There are blank category codes that need attention."
      frmBLMessageBoxJr.Label1.Top = 700
      frmBLMessageBoxJr.Show vbModal
      Close
      Exit Sub
    ElseIf Clear4CatCodeProblems = 2 Then
      frmBLMessageBoxJr.Label1.Caption = "Please examine your category code data. There are blank category descriptions that need attention."
      frmBLMessageBoxJr.Label1.Top = 700
      frmBLMessageBoxJr.Show vbModal
      Close
      Exit Sub
    End If
  End If
  DupCatCnt = 0
  
  Call Check4DupCats
  
  If DupCatCnt > 0 Then
    frmBLCheck4DupCats.Show vbModal
    If frmBLCheck4DupCats.fptxtChoice.Text = "abort" Then
      Unload frmBLCheck4DupCats
      Close
      fpcmbCatVersion.SetFocus
      fptxtMarque.Text = "Begin Conversion"
      Exit Sub
    Else
      Unload frmBLCheck4DupCats
    End If
  End If
  
  Call ChangeCat2Nums
End Sub

Private Sub cmdCustOnly_Click()
  If InStr(fpcmbVersion.Text, "1") Then
    frmBLCnvtHelpCustCat.Show vbModal
  ElseIf InStr(fpcmbVersion.Text, "2") Then
    frmBLCnvtHelpCustVs2.Show vbModal
  End If
End Sub

Private Sub cmdExit_Click()
  If Exist("CatVrs1.PRN") Then
    KillFile "CatVrs1.PRN"
  End If
  If Exist("CatVrs2.PRN") Then
    KillFile "CatVrs2.PRN"
  End If
  If Exist("Version1.PRN") Then
    KillFile "Version1.PRN"
  End If
  If Exist("Version2.PRN") Then
    KillFile "Version2.PRN"
  End If
  
  Unload frmBLConvertMain
End Sub

Private Sub cmdHelpCustCat_Click()
  If InStr(fpcmbCatVersion.Text, "1") Then
    frmBLCatCnvrHelpVs1.Show vbModal
  ElseIf InStr(fpcmbCatVersion.Text, "2") Then
    frmBLCatCnvrHelpVs2.Show vbModal
  End If
End Sub

Private Sub cmdPrintVersion_Click()
  If InStr(fpcmbVersion.Text, "1") Then
    Call PrintVersion1
  Else
    Call PrintVersion2
  End If
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      Call cmdExit_Click
      SendKeys "%C"
      KeyCode = 0
    Case vbKeyF10:
      Call cmdBegin_Click
      SendKeys "%B"
      KeyCode = 0
    Case vbKeyF2:
      cmdHelpCustCat_Click
      SendKeys "%H"
      KeyCode = 0
    Case vbKeyF3:
      Call cmdCustOnly_Click
      SendKeys "%o"
      KeyCode = 0
    Case vbKeyF4:
      Call cmdPrintVersion_Click
      SendKeys "%s"
      KeyCode = 0
    Case vbKeyF5:
      Call cmdCatVersion_Click
      SendKeys "%V"
      KeyCode = 0
    Case vbKeyF6:
      Call cmdChangeCat2Nums_Click
      SendKeys "%N"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub LoadMe()
  cmdBegin.Caption = "F10 Begin Conversion Processing"
  NOGLACCT = False
  chkCustomers.Value = 0
  chkCategories.Value = 0
  StartPath = App.Path
  fptxtMarque.Text = "Begin Conversion"
  fpcmbVersion.Text = "Version 1"
  fpcmbVersion.AddItem "Version 1"
  fpcmbVersion.AddItem "Version 2"
  fpcmbCatVersion.Text = "Version 1"
  fpcmbCatVersion.AddItem "Version 1"
  fpcmbCatVersion.AddItem "Version 2"
  Version = 1
  If Not Exist("arcode.dat") Then
    CatVersion = 0
  Else
    CatVersion = 1
  End If
  CheckCustFlag = False
  CheckCatFlag = False
End Sub

Private Sub ConvertCust2()
  Dim DosNumOfCustRecs2 As Integer
  Dim DosCustRec2 As DOSARCustRecType2
  Dim DosCustHandle2 As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim DosNumOfCustRecs As Integer
  Dim DosCustRec As DosARCustRecType
  Dim DosCustHandle As Integer
  Dim x As Integer, y As Integer
  Dim Nextx As Integer
  Dim DosCatRec As DosARNewCatCodeRecType
  Dim DosCatHandle As Integer
  Dim NumOfCatRecs As Integer
  Dim CatRec As ARNewCatCodeRecType
  Dim CatHandle As Integer
  Dim PctDone As Integer
  
  If Not Exist("arcust.dat") Then
    MsgBox "You have elected to convert customer data. However, the file 'arcust.dat' could not be found. Conversion aborted."
    Exit Sub
  End If
  
  cmdBegin.Enabled = False
  cmdExit.Enabled = False
  
  DoEvents
  fptxtMarque.Text = "Converting Customer Data"
  DoEvents
  
  OpenDosCustFile2 DosCustHandle2
  DosNumOfCustRecs2 = LOF(DosCustHandle2) / Len(DosCustRec2)
  If DosNumOfCustRecs2 = 0 Then
    MsgBox "No customers on file."
    Close
    Exit Sub
  End If

  ReDim TempCUSTNUMB(1 To DosNumOfCustRecs2) As String * 10
  ReDim TempSORTNAME(1 To DosNumOfCustRecs2) As String * 10
  ReDim TempBILLNAME(1 To DosNumOfCustRecs2) As String * 35
  ReDim TempADDRESS1(1 To DosNumOfCustRecs2) As String * 35
  ReDim TempADDRESS2(1 To DosNumOfCustRecs2) As String * 35
  ReDim TempCITY(1 To DosNumOfCustRecs2) As String * 20
  ReDim TempSTATE(1 To DosNumOfCustRecs2) As String * 2
  ReDim TempZIPCODE(1 To DosNumOfCustRecs2) As String * 10
  ReDim TempCustName(1 To DosNumOfCustRecs2) As String * 35
  ReDim TempBILLCAT(1 To DosNumOfCustRecs2) As String * 5
  ReDim TempSOSEC(1 To DosNumOfCustRecs2) As String * 11
  ReDim TempDRVLIC(1 To DosNumOfCustRecs2) As String * 12
  ReDim TempDATEOPED(1 To DosNumOfCustRecs2) As Integer
  ReDim TempBILLCMT(1 To DosNumOfCustRecs2) As String * 20
  ReDim TempPAYCMT(1 To DosNumOfCustRecs2) As String * 20
  ReDim TempCASHONLY(1 To DosNumOfCustRecs2) As String * 1
  ReDim TempAPPNUMB(1 To DosNumOfCustRecs2) As Integer
  ReDim TempBILLFORM(1 To DosNumOfCustRecs2) As Integer
  ReDim TempHPHONE(1 To DosNumOfCustRecs2) As String * 14
  ReDim TempWPHONE(1 To DosNumOfCustRecs2) As String * 14
  ReDim TempFeeAmt(1 To DosNumOfCustRecs2) As Double
  ReDim TempLICENSE(1 To DosNumOfCustRecs2) As String * 12
  ReDim TempVALID(1 To DosNumOfCustRecs2) As Integer
  ReDim TempAcctBal(1 To DosNumOfCustRecs2) As Double
  ReDim TempOldFirstTrans(1 To DosNumOfCustRecs2) As Integer
  ReDim TempOldLastTrans(1 To DosNumOfCustRecs2) As Integer
  ReDim TempDeleted(1 To DosNumOfCustRecs2) As String * 1         '(yY)=deleted, anything else isn't
  ReDim TempFirstTrans(1 To DosNumOfCustRecs2) As Long
  ReDim TempLastTrans(1 To DosNumOfCustRecs2) As Long
  ReDim TempIssueLicense(1 To DosNumOfCustRecs2) As String * 1
  ReDim TempIssuanceFee(1 To DosNumOfCustRecs2) As Currency
  ReDim TempCustLocation(1 To DosNumOfCustRecs2) As String * 1
  ReDim TempRoomtoGrow(1 To DosNumOfCustRecs2) As String * 164
  
  For x = 1 To DosNumOfCustRecs2
    Get DosCustHandle2, x, DosCustRec2
    TempCUSTNUMB(x) = DosCustRec2.CUSTNUMB
    TempSORTNAME(x) = DosCustRec2.SORTNAME
    TempBILLNAME(x) = DosCustRec2.BILLNAME
    TempADDRESS1(x) = DosCustRec2.ADDRESS1
    TempADDRESS2(x) = DosCustRec2.ADDRESS2
    TempCITY(x) = DosCustRec2.CITY
    TempSTATE(x) = DosCustRec2.STATE
    TempZIPCODE(x) = DosCustRec2.ZIPCODE
    TempCustName(x) = DosCustRec2.CustName
    TempBILLCAT(x) = DosCustRec2.BILLCAT
    TempSOSEC(x) = DosCustRec2.SOSEC
    TempDRVLIC(x) = DosCustRec2.DRVLIC
    TempDATEOPED(x) = DosCustRec2.DATEOPED
    TempBILLCMT(x) = DosCustRec2.BILLCMT
    TempPAYCMT(x) = DosCustRec2.PAYCMT
    TempCASHONLY(x) = DosCustRec2.CASHONLY
    TempAPPNUMB(x) = DosCustRec2.APPNUMB
    TempBILLFORM(x) = DosCustRec2.BILLFORM
    TempHPHONE(x) = DosCustRec2.HPHONE
    TempWPHONE(x) = DosCustRec2.WPHONE
    TempFeeAmt(x) = DosCustRec2.FeeAmt
    TempLICENSE(x) = DosCustRec2.LICENSE
    TempVALID(x) = DosCustRec2.VALID
    If InStr(DosCustRec.AcctBal, "E") Then DosCustRec.AcctBal = 0
    TempAcctBal(x) = DosCustRec2.AcctBal
    TempOldFirstTrans(x) = DosCustRec2.OldFirstTrans
    TempOldLastTrans(x) = DosCustRec2.OldLastTrans
    If QPTrim$(DosCustRec2.Deleted) <> "Y" Then DosCustRec2.Deleted = "N"
    TempDeleted(x) = DosCustRec2.Deleted
    TempFirstTrans(x) = DosCustRec2.FirstTrans
    TempLastTrans(x) = DosCustRec2.LastTrans
    TempIssueLicense(x) = DosCustRec2.IssueLicense
    TempIssuanceFee(x) = DosCustRec2.IssuanceFee
    TempCustLocation(x) = DosCustRec2.CustLocation
    TempRoomtoGrow(x) = DosCustRec2.RoomtoGrow
    DoEvents
    If DosNumOfCustRecs > 0 Then
      PctDone = (x / DosNumOfCustRecs) * 100
      fptxtMarque.Text = "Collecting old customer data is " + CStr(PctDone) + "% completed."
    End If
  Next x
  Close DosCustHandle2
  
  OpenCustFile CustHandle
  
  For x = 1 To DosNumOfCustRecs2
    Get CustHandle, x, CustRec
    CustRec.CUSTNUMB = QPTrim(TempCUSTNUMB(x))
    CustRec.SORTNAME = QPTrim(TempSORTNAME(x))
    CustRec.BILLNAME = QPTrim(TempBILLNAME(x))
    CustRec.ADDRESS1 = QPTrim(TempADDRESS1(x))
    CustRec.ADDRESS2 = QPTrim(TempADDRESS2(x))
    CustRec.CITY = QPTrim(TempCITY(x))
    CustRec.STATE = QPTrim(TempSTATE(x))
    CustRec.ZIPCODE = QPTrim(TempZIPCODE(x))
    CustRec.CustName = QPTrim(TempCustName(x))
    CustRec.Contact = ""
    CustRec.ServAdd = ""
    CustRec.SSNFID = ""
    CustRec.BILLCAT1 = QPTrim(TempBILLCAT(x))
    If CatVersion <> 0 Then
      CustRec.DESC1 = GetCatDesc(QPTrim$(CustRec.BILLCAT1))
    Else
      CustRec.DESC1 = "UnKnown"
    End If
    CustRec.REV1 = 0
    CustRec.Fee1 = 0 'TempFeeAmt(x)
    'The Dos version does not have balances for each of the
    'five potential categories so in the conversion the Lic Bal
    'is automatically dumped into the first category just to have
    'a post conversion license category balance
    CustRec.FeeLicBal1 = 0
    CustRec.FeeLicPay1 = 0
    CustRec.BILLCAT2 = ""
    CustRec.DESC2 = ""
    CustRec.REV2 = 0
    CustRec.Fee2 = 0
    CustRec.FeeLicBal2 = 0
    CustRec.FeeLicPay2 = 0
    CustRec.BILLCAT3 = ""
    CustRec.DESC3 = ""
    CustRec.REV3 = 0
    CustRec.Fee3 = 0
    CustRec.FeeLicBal3 = 0
    CustRec.FeeLicPay3 = 0
    CustRec.BILLCAT4 = ""
    CustRec.DESC4 = ""
    CustRec.REV4 = 0
    CustRec.Fee4 = 0
    CustRec.FeeLicBal4 = 0
    CustRec.FeeLicPay4 = 0
    CustRec.BILLCAT5 = ""
    CustRec.DESC5 = ""
    CustRec.REV5 = 0
    CustRec.Fee5 = 0
    CustRec.FeeLicBal5 = 0
    CustRec.FeeLicPay5 = 0
    CustRec.IssuanceFee = 0
    CustRec.CustLocation = QPTrim(TempCustLocation(x))
    CustRec.WPHONE = QPTrim(TempWPHONE(x))
    CustRec.FeeAmt = 0
    CustRec.LICENSE = QPTrim(TempLICENSE(x))
    CustRec.VALID = TempVALID(x)
    CustRec.Inactive = "N"
    CustRec.Prorate = 100
    CustRec.AcctBal = 0
    CustRec.IssueLicense = QPTrim(TempIssueLicense(x))
    CustRec.Deleted = QPTrim(TempDeleted(x))
    CustRec.FirstTrans = 0
    CustRec.LastTrans = 0
    CustRec.LicBal = 0
    CustRec.FeeBal = 0
    CustRec.PenBal = 0
    CustRec.RoomtoGrow = QPTrim(TempRoomtoGrow(x))
    CustRec.ChkByte = ""
    CustRec.IssuanceBal = 0
    CustRec.IssuancePay = 0
    Put CustHandle, x, CustRec
    DoEvents
    If DosNumOfCustRecs > 0 Then
      PctDone = (x / DosNumOfCustRecs) * 100
      fptxtMarque.Text = "Converting customer data is " + CStr(PctDone) + "% completed."
    End If
  Next x
  
  On Error Resume Next
  Nextx = 1
  For x = 1 To DosNumOfCustRecs2
  Get CustHandle, x, CustRec
    If QPTrim$(CustRec.BILLNAME) = "" Then
      CustRec.BILLNAME = "INVALID" + CStr(x)
      CustRec.SORTNAME = "INV" + CStr(x)
      CustRec.CustName = "INVALID" + CStr(x)
      CustRec.Contact = "INVALID" + CStr(x)
      CustRec.CITY = "INVALID"
      CustRec.CUSTNUMB = CStr(x)
      GoTo BadData
    End If
    
    If QPTrim$(CustRec.Inactive) <> "Y" And QPTrim$(CustRec.Inactive) <> "N" Then
      CustRec.Inactive = "N"
    End If
    
    If CustRec.Prorate <= 0 Or CustRec.Prorate > 100 Then
      CustRec.Prorate = 100
    End If
    
    If QPTrim$(CustRec.IssueLicense) <> "Y" And QPTrim$(CustRec.IssueLicense) <> "N" Then
      CustRec.IssueLicense = "N"
    End If
    
    If CustRec.VALID < 0 Or CustRec.VALID > 10000 Then
      CustRec.VALID = 0
    End If
    
BadData:
  Put CustHandle, x, CustRec
  Next x
  
  Close CustHandle
  
  fptxtMarque.Text = "Creating Indices"
  DoEvents
  
  Call CreateCustNumIdx
  Call CreateCustSearchNameIdx
  Call CreateCustNameIdx
  Call CreateLicNumIdx
  
  fptxtMarque.FontSize = 14
  fptxtMarque.Text = "Business License Conversion has completed"
  
  cmdExit.Enabled = True
  
End Sub

Private Sub fpcmbCatVersion_Change()
  If QPTrim$(fpcmbCatVersion.Text) = "" Then
    fpcmbCatVersion.Text = "F5 Cat &Version #1 Check"
    CatVersion = 1
  End If
  
  If InStr(fpcmbCatVersion.Text, "1") Then
    cmdCatVersion.Caption = "F5 Cat &Version #1 Check"
    CatVersion = 1
  ElseIf InStr(fpcmbCatVersion.Text, "2") Then
    cmdCatVersion.Caption = "F5 Cat &Version #2 Check"
    CatVersion = 2
  Else
    cmdCatVersion.Caption = "F5 Cat &Version #1 Check"
    CatVersion = 1
  End If
End Sub

Private Sub fpcmbVersion_Change()
  If QPTrim$(fpcmbVersion.Text) = "" Then
    cmdPrintVersion.Caption = "F4 Cu&st Version #1 Check"
    Version = 1
  End If
  
  If InStr(fpcmbVersion.Text, "1") Then
    cmdPrintVersion.Caption = "F4 Cu&st Version #1 Check"
    Version = 1
  ElseIf InStr(fpcmbVersion.Text, "2") Then
    cmdPrintVersion.Caption = "F4 Cu&st Version #2 Check"
    Version = 2
  Else
    cmdPrintVersion.Caption = "F4 Cu&st Version #1 Check"
    Version = 1
  End If
End Sub

Private Sub PrintVersion1()
  Dim DosNumOfCustRecs As Integer
  Dim DosCustRec As DosARCustRecType
  Dim DosCustHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim x As Integer
  Dim ReportFile$
  Dim RptHandle As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim FF$
  
  MaxLines = 55
  LineCnt = 0
  FF$ = Chr$(12)
  ReportFile$ = "Version1.PRN"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  OpenDosCustFile DosCustHandle
  NumOfCustRecs = LOF(DosCustHandle) / Len(DosCustRec)
  GoSub PrintHeader
  For x = 1 To NumOfCustRecs
    Get DosCustHandle, x, DosCustRec
    If QPTrim$(DosCustRec.Deleted) = "Y" Or QPTrim$(DosCustRec.SORTNAME) = "DELETED" Then GoTo SkipIt
    If LineCnt > MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
    Print #RptHandle, QPTrim$(DosCustRec.CustName); Tab(37); QPTrim$(DosCustRec.CUSTNUMB); Tab(49); QPTrim$(DosCustRec.CITY); Tab(71); QPTrim$(DosCustRec.BILLCAT1)
    LineCnt = LineCnt + 1
SkipIt:
  Next x
  
  Print #RptHandle, FF$
  
  Close
  
  ViewPrint ReportFile, "Version 1", True
  CheckCustFlag = True
  
  Exit Sub
  
PrintHeader:
  Print #RptHandle, "Version #1"
  Print #RptHandle,
  Print #RptHandle, "Number of Customers On File: " + CStr(NumOfCustRecs)
  Print #RptHandle, "Cust Name"; Tab(34); "CustNum"; Tab(53); "City"; Tab(71); "Cat #1"
  Print #RptHandle, String$(80, "=")
  LineCnt = 5
  Return

End Sub
Private Sub PrintVersion2()
  Dim DosNumOfCustRecs2 As Integer
  Dim DosCustRec2 As DOSARCustRecType2
  Dim DosCustHandle2 As Integer
  Dim NumOfCustRecs As Integer
  Dim x As Integer
  Dim ReportFile$
  Dim RptHandle As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim FF$
  
  MaxLines = 55
  LineCnt = 0
  FF$ = Chr$(12)
  
  ReportFile$ = "Version2.PRN"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  OpenDosCustFile2 DosCustHandle2
  NumOfCustRecs = LOF(DosCustHandle2) / Len(DosCustRec2)
  GoSub PrintHeader
  For x = 1 To NumOfCustRecs
    Get DosCustHandle2, x, DosCustRec2
    If QPTrim$(DosCustRec2.Deleted) = "Y" Or QPTrim$(DosCustRec2.SORTNAME) = "DELETED" Then GoTo SkipIt
    If LineCnt > MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
    Print #RptHandle, QPTrim$(DosCustRec2.CustName); Tab(37); QPTrim$(DosCustRec2.CUSTNUMB); Tab(49); QPTrim$(DosCustRec2.CITY); Tab(71); QPTrim$(DosCustRec2.LICENSE)
    LineCnt = LineCnt + 1
SkipIt:
  Next x
  
  Print #RptHandle, FF$
  Close
  
  ViewPrint ReportFile, "Version 2", True
  
  CheckCustFlag = True
  
  Exit Sub
  
PrintHeader:
  Print #RptHandle, "Version #2"
  Print #RptHandle,
  Print #RptHandle, "Number of Customers On File: " + CStr(NumOfCustRecs)
  Print #RptHandle, "Cust Name"; Tab(37); "CustNum"; Tab(49); "City"; Tab(71); "License #"
  Print #RptHandle, String$(80, "=")
  LineCnt = 5
  Return

End Sub

Private Sub PrintCatVersion2()
  Dim DosCatHandle2 As Integer
  Dim DosCatRec2 As DosARNewCatCodeRecType2
  Dim NumOfCatRecs As Integer
  Dim x As Integer
  Dim ReportFile$
  Dim RptHandle As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim FF$
  Dim y As Integer
  Dim ActiveFlag As Boolean
  
  MaxLines = 55
  LineCnt = 0
  FF$ = Chr$(12)
  ReportFile$ = "CatVrs2.PRN"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  OpenDosCatFile2 DosCatHandle2
  NumOfCatRecs = LOF(DosCatHandle2) / Len(DosCatRec2)

  GoSub PrintHeader

  For x = 1 To NumOfCatRecs
    ActiveFlag = True
    Get DosCatHandle2, x, DosCatRec2
    For y = 1 To DupCatCnt
      If x = DupCats(y) Then
        ActiveFlag = False
        Exit For
      End If
    Next y
    If LineCnt > MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
    If ActiveFlag = True Then
      Print #RptHandle, QPTrim$(DosCatRec2.CATCODE); Tab(49); QPTrim$(DosCatRec2.CODEDESC)
    Else
      Print #RptHandle, QPTrim$(DosCatRec2.CATCODE) + " INACTIVE"; Tab(49); QPTrim$(DosCatRec2.CODEDESC)
    End If
    LineCnt = LineCnt + 1
  Next x

  Print #RptHandle, FF$
  
  Close
  
  ViewPrint ReportFile, "Category Version 2", True
  CheckCatFlag = True

  Exit Sub

PrintHeader:
  Print #RptHandle, "Category Version #2"
  Print #RptHandle, "Categories marked as 'INACTIVE' should be edited after conversion completes"
  Print #RptHandle,
  Print #RptHandle, "Number of Categories On File: " + CStr(NumOfCatRecs)
  Print #RptHandle, "Cat Code"; Tab(49); "Description"
  Print #RptHandle, String$(80, "=")
  LineCnt = 6
  Return

End Sub

Private Sub PrintCatVersion1()
  Dim DosCatHandle As Integer
  Dim DosCatRec As DosARNewCatCodeRecType
  Dim NumOfCatRecs As Integer
  Dim x As Integer
  Dim ReportFile$
  Dim RptHandle As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim FF$
  Dim y As Integer
  Dim ActiveFlag As Boolean
  
  MaxLines = 55
  LineCnt = 0
  FF$ = Chr$(12)
  ReportFile$ = "CatVrs1.PRN"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  OpenDosCatFile DosCatHandle
  NumOfCatRecs = LOF(DosCatHandle) / Len(DosCatRec)
  
  GoSub PrintHeader
  
  For x = 1 To NumOfCatRecs
    ActiveFlag = True
    Get DosCatHandle, x, DosCatRec
    For y = 1 To DupCatCnt
      If x = DupCats(y) Then
        ActiveFlag = False
        Exit For
      End If
    Next y
    If LineCnt > MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
    If ActiveFlag = True Then
      Print #RptHandle, QPTrim$(DosCatRec.CATCODE); Tab(49); QPTrim$(DosCatRec.CODEDESC)
    Else
      Print #RptHandle, QPTrim$(DosCatRec.CATCODE) + " **INACTIVE**"; Tab(49); QPTrim$(DosCatRec.CODEDESC)
    End If
    LineCnt = LineCnt + 1
  Next x
  Print #RptHandle, FF$
  
  Close
  
  ViewPrint ReportFile, "Category Version 1", True
  CheckCatFlag = True
  
  Exit Sub
  
PrintHeader:
  Print #RptHandle, "Category Version #1"
  Print #RptHandle,
  Print #RptHandle, "Number of Categories On File: " + CStr(NumOfCatRecs)
  Print #RptHandle, "Cat Code"; Tab(49); "Description"
  Print #RptHandle, String$(80, "=")
  LineCnt = 5
  
  Return

End Sub

Private Sub ConvertCategory1()
  Dim DosCatRec As DosARNewCatCodeRecType
  Dim DosCatHandle As Integer
  Dim NumOfCatRecs As Integer
  Dim CatRec As ARNewCatCodeRecType
  Dim CatHandle As Integer
  Dim x As Integer
  Dim PctDone As Integer
  
'  If Not Exist("GLACCT.DAT") Then
'    frmBLMessageBoxJrWOpts.Label1.Caption = "GLACCT.DAT could not be found. Continuing will make all GL numbers in category codes equal zero. This is because the save method for GL numbers is now based on their record numbers instead of the numbers themselves. Without the 'GLACCT.DAT' file there is no way to accurately convert category code GL numbers. Do you wish to continue?"
'    frmBLMessageBoxJrWOpts.Label1.Top = 500
'    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
'    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Abort"
'    frmBLMessageBoxJrWOpts.Show vbModal
'    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
'      Unload frmBLMessageBoxJrWOpts
'      NOGLACCT = True
'      Close
'      Exit Sub
'    Else
'      Unload frmBLMessageBoxJrWOpts
'    End If
'  End If
  
  fptxtMarque.Text = "Converting Category Code Data"
  DoEvents
  
  OpenDosCatFile DosCatHandle
  NumOfCatRecs = LOF(DosCatHandle) / Len(DosCatRec)
  
  ReDim TempCATCODE(1 To NumOfCatRecs) As String * 5
  ReDim TempCodeType(1 To NumOfCatRecs) As String * 1
  ReDim TempCODEDESC(1 To NumOfCatRecs) As String * 35
  ReDim TempFee(1 To NumOfCatRecs) As Single
  ReDim TempBaseAmt1(1 To NumOfCatRecs) As Single
  ReDim TempRecpt1(1 To NumOfCatRecs) As Double
  ReDim TempPercent1(1 To NumOfCatRecs) As Single
  ReDim TempMaximum1(1 To NumOfCatRecs) As Double
  ReDim TempBaseAmt2(1 To NumOfCatRecs) As Single
  ReDim TempRecpt2(1 To NumOfCatRecs) As Double
  ReDim TempPercent2(1 To NumOfCatRecs) As Single
  ReDim TempMaximum2(1 To NumOfCatRecs) As Double
  ReDim TempBaseAmt3(1 To NumOfCatRecs) As Single
  ReDim TempRecpt3(1 To NumOfCatRecs) As Double
  ReDim TempPercent3(1 To NumOfCatRecs) As Single
  ReDim TempMaximum3(1 To NumOfCatRecs) As Double
  ReDim TempBaseAmt4(1 To NumOfCatRecs) As Single
  ReDim TempRecpt4(1 To NumOfCatRecs) As Double
  ReDim TempPercent4(1 To NumOfCatRecs) As Single
  ReDim TempMaximum4(1 To NumOfCatRecs) As Double
  ReDim TempBaseAmt5(1 To NumOfCatRecs) As Single
  ReDim TempRecpt5(1 To NumOfCatRecs) As Double
  ReDim TempPercent5(1 To NumOfCatRecs) As Single
  ReDim TempMaximum5(1 To NumOfCatRecs) As Double
  ReDim TempREVGLNUM(1 To NumOfCatRecs) As String * 14
  ReDim TempCASHACCT(1 To NumOfCatRecs) As String * 14
  ReDim TempARGLACCT(1 To NumOfCatRecs) As String * 14
  ReDim TempBaseAmt6(1 To NumOfCatRecs) As Single
  ReDim TempRecpt6(1 To NumOfCatRecs) As Double
  ReDim TempPercent6(1 To NumOfCatRecs) As Single
  ReDim TempMaximum6(1 To NumOfCatRecs) As Double
  ReDim TempRateStep(1 To NumOfCatRecs) As Long
  ReDim TempExtra(1 To NumOfCatRecs) As String * 36
  
  For x = 1 To NumOfCatRecs
    Get DosCatHandle, x, DosCatRec
      TempCATCODE(x) = DosCatRec.CATCODE
      TempCodeType(x) = DosCatRec.CodeType
      TempCODEDESC(x) = DosCatRec.CODEDESC
      TempFee(x) = DosCatRec.Fee
      TempBaseAmt1(x) = DosCatRec.BaseAmt1
      TempRecpt1(x) = DosCatRec.Recpt1
      TempPercent1(x) = DosCatRec.Percent1
      TempMaximum1(x) = DosCatRec.Maximum1
      TempBaseAmt2(x) = DosCatRec.BaseAmt2
      TempRecpt2(x) = DosCatRec.Recpt2
      TempPercent2(x) = DosCatRec.Percent2
      TempMaximum2(x) = DosCatRec.Maximum2
      TempBaseAmt3(x) = DosCatRec.BaseAmt3
      TempRecpt3(x) = DosCatRec.Recpt3
      TempPercent3(x) = DosCatRec.Percent3
      TempMaximum3(x) = DosCatRec.Maximum3
      TempBaseAmt4(x) = DosCatRec.BaseAmt4
      TempRecpt4(x) = DosCatRec.Recpt4
      TempPercent4(x) = DosCatRec.Percent4
      TempMaximum4(x) = DosCatRec.Maximum4
      TempBaseAmt5(x) = DosCatRec.BaseAmt5
      TempRecpt5(x) = DosCatRec.Recpt5
      TempPercent5(x) = DosCatRec.Percent5
      TempMaximum5(x) = DosCatRec.Maximum5
      TempREVGLNUM(x) = DosCatRec.REVGLNUM
      TempCASHACCT(x) = DosCatRec.CASHACCT
      TempARGLACCT(x) = DosCatRec.ARGLACCT
    
      TempBaseAmt6(x) = DosCatRec.BaseAmt6
      TempRecpt6(x) = DosCatRec.Recpt6
      TempPercent6(x) = DosCatRec.Percent6
      TempMaximum6(x) = DosCatRec.Maximum6
      TempRateStep(x) = DosCatRec.RateStep
      TempExtra(x) = DosCatRec.Extra
      DoEvents
      PctDone = (x / NumOfCatRecs) * 100
      fptxtMarque.Text = "Collecting old category data is " + CStr(PctDone) + "% completed."
  Next x
  Close DosCatHandle
  
  DoEvents
  KillFile "OldARCode.dat"
  Name "ARCODE.DAT" As "OldARCode.dat"
  'old is bigger than the new so if you don't kill it it
  'will retain some garbage we don't need
  OpenCatCodeFile CatHandle
  
  For x = 1 To NumOfCatRecs
    CatRec.CATCODE = TempCATCODE(x)
    CatRec.CodeType = TempCodeType(x)
    CatRec.CODEDESC = TempCODEDESC(x)
    CatRec.Fee = TempFee(x)
    CatRec.BaseAmt1 = TempBaseAmt1(x)
    CatRec.Recpt1 = TempRecpt1(x)
    CatRec.Percent1 = TempPercent1(x)
    CatRec.Maximum1 = TempMaximum1(x)
    CatRec.BaseAmt2 = TempBaseAmt2(x)
    CatRec.Recpt2 = TempRecpt2(x)
    CatRec.Percent2 = TempPercent2(x)
    CatRec.Maximum2 = TempMaximum2(x)
    CatRec.BaseAmt3 = TempBaseAmt3(x)
    CatRec.Recpt3 = TempRecpt3(x)
    CatRec.Percent3 = TempPercent3(x)
    CatRec.Maximum3 = TempMaximum3(x)
    CatRec.BaseAmt4 = TempBaseAmt4(x)
    CatRec.Recpt4 = TempRecpt4(x)
    CatRec.Percent4 = TempPercent4(x)
    CatRec.Maximum4 = TempMaximum4(x)
    CatRec.BaseAmt5 = TempBaseAmt5(x)
    CatRec.Recpt5 = TempRecpt5(x)
    CatRec.Percent5 = TempPercent5(x)
    CatRec.Maximum5 = TempMaximum5(x)
'    If Not Exist("GLACCT.DAT") Then
    If Exist("GLACCT.DAT") Then
      CatRec.REVGLNUM = GetGLRecNum(QPTrim$(TempREVGLNUM(x)))
      CatRec.CASHACCT = GetGLRecNum(QPTrim$(TempCASHACCT(x)))
      CatRec.ARGLACCT = GetGLRecNum(QPTrim$(TempARGLACCT(x)))
    Else
      CatRec.REVGLNUM = 0
      CatRec.CASHACCT = 0
      CatRec.ARGLACCT = 0
    End If
    CatRec.BaseAmt6 = TempBaseAmt6(x)
    CatRec.Recpt6 = TempRecpt6(x)
    CatRec.Percent6 = TempPercent6(x)
    CatRec.Maximum6 = TempMaximum6(x)
    CatRec.RateStep = TempRateStep(x)
    CatRec.Extra = TempExtra(x)
    Put CatHandle, x, CatRec
    DoEvents
    PctDone = (x / NumOfCatRecs) * 100
    fptxtMarque.Text = "Converting category data is " + CStr(PctDone) + "% completed."
  Next x
  Close CatHandle
  
  Call CreateCatCodeIdx

End Sub

Private Sub ConvertCategory2()
  Dim DosCatRec2 As DosARNewCatCodeRecType2
  Dim DosCatHandle2 As Integer
  Dim NumOfCatRecs As Integer
  Dim CatRec As ARNewCatCodeRecType
  Dim CatHandle As Integer
  Dim x As Integer
  Dim PctDone As Integer
  
'  If Not Exist("GLACCT.DAT") Then
'    frmBLMessageBoxJrWOpts.Label1.Caption = "GLACCT.DAT could not be found. Continuing will make all GL numbers in category codes equal zero. This is because the save method for GL numbers is now based on their record numbers instead of the numbers themselves. Without the 'GLACCT.DAT' file there is no way to accurately convert category code GL numbers. Do you wish to continue?"
'    frmBLMessageBoxJrWOpts.Label1.Top = 500
'    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
'    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Abort"
'    frmBLMessageBoxJrWOpts.Show vbModal
'    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
'      Unload frmBLMessageBoxJrWOpts
'      NOGLACCT = True
'      Close
'      Exit Sub
'    Else
'      Unload frmBLMessageBoxJrWOpts
'    End If
'  End If
  
  frmBLMessageBoxJr.Label1.Caption = "Reminder: Version #2 category conversion is unable to convert category types or category rates."
  frmBLMessageBoxJr.Label1.Top = 700
  frmBLMessageBoxJr.Show vbModal
  
  fptxtMarque.Text = "Converting Category Code Data"
  DoEvents
  
  OpenDosCatFile2 DosCatHandle2
  NumOfCatRecs = LOF(DosCatHandle2) / Len(DosCatRec2)
 
  ReDim TempCATCODE(1 To NumOfCatRecs) As String * 3
  ReDim TempCODEDESC(1 To NumOfCatRecs) As String * 35
  ReDim TempAPPNUMB(1 To NumOfCatRecs) As Integer
  ReDim TempBILLCODE(1 To NumOfCatRecs) As Integer
  ReDim TempREVGLNUM(1 To NumOfCatRecs) As String * 14
  ReDim TempCASHACCT(1 To NumOfCatRecs) As String * 14
  ReDim TempALCATCODE(1 To NumOfCatRecs) As String * 5
  ReDim TempARGLACCT(1 To NumOfCatRecs) As String * 14
  ReDim TempExtra(1 To NumOfCatRecs) As String * 39
  
  For x = 1 To NumOfCatRecs
    Get DosCatHandle2, x, DosCatRec2
      TempCATCODE(x) = DosCatRec2.CATCODE
      TempCODEDESC(x) = DosCatRec2.CODEDESC
      TempAPPNUMB(x) = DosCatRec2.APPNUMB
      TempBILLCODE(x) = DosCatRec2.BILLCODE
      TempREVGLNUM(x) = DosCatRec2.REVGLNUM
      TempCASHACCT(x) = DosCatRec2.CASHACCT
      TempALCATCODE(x) = DosCatRec2.ALCATCODE
      TempARGLACCT(x) = DosCatRec2.ARGLACCT
      TempExtra(x) = DosCatRec2.Extra
      DoEvents
      PctDone = (x / NumOfCatRecs) * 100
      fptxtMarque.Text = "Collecting old category data is " + CStr(PctDone) + "% completed."
  Next x
  Close DosCatHandle2
  
  DoEvents
  KillFile "OldARCode.dat"
  Name "ARCODE.DAT" As "OldARCode.dat"
  'old is bigger than the new so if you don't kill it it
  'will retain some garbage we don't need
  OpenCatCodeFile CatHandle
  
  For x = 1 To NumOfCatRecs
    CatRec.CATCODE = TempCATCODE(x)
    CatRec.CodeType = "N"
    CatRec.CODEDESC = TempCODEDESC(x)
    CatRec.Fee = 0
    CatRec.BaseAmt1 = 0
    CatRec.Recpt1 = 0
    CatRec.Percent1 = 0
    CatRec.Maximum1 = 0
    CatRec.BaseAmt2 = 0
    CatRec.Recpt2 = 0
    CatRec.Percent2 = 0
    CatRec.Maximum2 = 0
    CatRec.BaseAmt3 = 0
    CatRec.Recpt3 = 0
    CatRec.Percent3 = 0
    CatRec.Maximum3 = 0
    CatRec.BaseAmt4 = 0
    CatRec.Recpt4 = 0
    CatRec.Percent4 = 0
    CatRec.Maximum4 = 0
    CatRec.BaseAmt5 = 0
    CatRec.Recpt5 = 0
    CatRec.Percent5 = 0
    CatRec.Maximum5 = 0
'    If Not Exist("GLACCT.DAT") Then
    If Exist("GLACCT.DAT") Then
      CatRec.REVGLNUM = GetGLRecNum(QPTrim$(TempREVGLNUM(x)))
      CatRec.CASHACCT = GetGLRecNum(QPTrim$(TempCASHACCT(x)))
      CatRec.ARGLACCT = GetGLRecNum(QPTrim$(TempARGLACCT(x)))
    Else
      CatRec.REVGLNUM = 0
      CatRec.CASHACCT = 0
      CatRec.ARGLACCT = 0
    End If
    CatRec.BaseAmt6 = 0
    CatRec.Recpt6 = 0
    CatRec.Percent6 = 0
    CatRec.Maximum6 = 0
    CatRec.RateStep = 0
    CatRec.Extra = TempExtra(x)
    Put CatHandle, x, CatRec
    DoEvents
    PctDone = (x / NumOfCatRecs) * 100
    fptxtMarque.Text = "Converting category data is " + CStr(PctDone) + "% completed."
  Next x
  Close CatHandle
  
  Call CreateCatCodeIdx
  
End Sub

Private Function CheckData() As Boolean

  fptxtMarque.Text = "Checking for pre-conversion data problems."

  CheckData = True
  
  If chkCustomers.Value = 0 And chkCategories.Value = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No conversion selections have been made. Conversion attempt aborted."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
    CheckData = False
    Exit Function
  End If
  
  If chkCategories.Value = 1 Then
    If Not Exist("arcode.dat") Then
      frmBLMessageBoxJr.Label1.Caption = "You have elected to convert category data. However, the file 'arcode.dat' needed to convert category data could not be found. Conversion aborted."
      frmBLMessageBoxJr.Label1.Top = 700
      frmBLMessageBoxJr.Show vbModal
      CheckData = False
      Exit Function
    End If
  End If
  
  If chkCategories.Value = 0 Then
    If Exist("arcode.dat") Then
      frmBLMessageBoxJr.Label1.Caption = "You have elected not to convert category data. However, the file 'arcode.dat' can be found in the current directory. Because this file will not be compatible in the new business license program it needs to be removed. Conversion aborted."
      frmBLMessageBoxJr.Label1.Top = 600
      frmBLMessageBoxJr.Show vbModal
      CheckData = False
      Exit Function
    End If
  End If
  
  If Not Exist("arcode.dat") Then GoTo NoARCode
  
  If Clear4CatCodeProblems > 0 Then
    If Clear4CatCodeProblems = 3 Then
      frmBLMessageBoxJr.Label1.Caption = "Please examine your category code data. There are blank category codes and blank category descriptions that need attention."
      frmBLMessageBoxJr.Label1.Top = 700
      frmBLMessageBoxJr.Show vbModal
      Close
      CheckData = False
      Exit Function
    ElseIf Clear4CatCodeProblems = 1 Then
      frmBLMessageBoxJr.Label1.Caption = "Please examine your category code data. There are blank category codes that need attention."
      frmBLMessageBoxJr.Label1.Top = 700
      frmBLMessageBoxJr.Show vbModal
      Close
      CheckData = False
      Exit Function
    ElseIf Clear4CatCodeProblems = 2 Then
      frmBLMessageBoxJr.Label1.Caption = "Please examine your category code data. There are blank category descriptions that need attention."
      frmBLMessageBoxJr.Label1.Top = 700
      frmBLMessageBoxJr.Show vbModal
      Close
      CheckData = False
      Exit Function
    End If
  End If
  
NoARCode:
  If chkCustomers.Value = 1 Then
    If Not Exist("arcust.dat") Then
      frmBLMessageBoxJr.Label1.Caption = "You have elected to convert customer data. However, the file 'arcust.dat' could not be found. Conversion aborted."
      frmBLMessageBoxJr.Label1.Top = 700
      frmBLMessageBoxJr.Show vbModal
      Close
      CheckData = False
      Exit Function
    End If
  End If
  
  If chkCategories.Value = 0 Then
    If Exist("arcode.dat") Then
      frmBLMessageBoxJr.Label1.Caption = "You have elected not to convert categories but the file 'arcode.dat' still exists in this directory. Please delete 'arcode.dat' before continuing."
      frmBLMessageBoxJr.Label1.Top = 700
      frmBLMessageBoxJr.Show vbModal
      Close
      CheckData = False
      Exit Function
    End If
  End If
  
  If chkCategories.Value = 1 Then
    If CheckCatFlag = False Then
      frmBLMessageBoxJr.Label1.Caption = "To insure that conversion is converting the correct category code file please check the category version before continuing. Conversion attempt aborted."
      frmBLMessageBoxJr.Label1.Top = 700
      frmBLMessageBoxJr.Show vbModal
      Close
      CheckData = False
      Exit Function
    End If
  End If
  
  If chkCustomers.Value = 1 Then
    If CheckCustFlag = False Then
      frmBLMessageBoxJr.Label1.Caption = "To insure that conversion is converting the correct customer code file please check the customer version before continuing. Conversion attempt aborted."
      frmBLMessageBoxJr.Label1.Top = 700
      frmBLMessageBoxJr.Show vbModal
      Close
      CheckData = False
      Exit Function
    End If
  End If
  
  If Not Exist("arcode.dat") Then GoTo NoARCode2
  
  If Version = 1 Then
    fptxtMarque.Text = "Examining category code data for problems."
    DoEvents
    If Clear4CatCodeProblems > 0 Then
      If Clear4CatCodeProblems = 3 Then
        frmBLMessageBoxJr.Label1.Caption = "Please examine your category code data. There are blank category codes and blank category descriptions that need attention."
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Show vbModal
        Close
        CheckData = False
        Exit Function
      ElseIf Clear4CatCodeProblems = 1 Then
        frmBLMessageBoxJr.Label1.Caption = "Please examine your category code data. There are blank category codes that need attention."
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Show vbModal
        Close
        CheckData = False
        Exit Function
      ElseIf Clear4CatCodeProblems = 2 Then
        frmBLMessageBoxJr.Label1.Caption = "Please examine your category code data. There are blank category descriptions that need attention."
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Show vbModal
        Close
        CheckData = False
        Exit Function
      End If
    End If
  End If

NoARCode2:
  fptxtMarque.Text = "Examining category descriptions for each customer."
  
  If chkCategories.Value = 1 And Exist("arcode.dat") Then
    If Version = 1 Then
      If Check4BlankCatDesc() = True Then
        DoEvents
        frmBLCheck4BlankCatDesc.Show vbModal
        If frmBLCheck4BlankCatDesc.fptxtChoice.Text = "exit" Then
          cmdBegin.Enabled = True
          cmdExit.Enabled = True
          Close
          CheckData = False
          Unload frmBLCheck4BlankCatDesc
          Exit Function
        Else
          Unload frmBLCheck4BlankCatDesc
        End If
      End If
    End If
  End If
  
  DoEvents
  
  fptxtMarque.Text = "Examining license numbers. "
  
  If Check4BlankLicNums() = True Then
    frmBLCheck4BlankLicNums.Show vbModal
    DoEvents
    If frmBLCheck4BlankLicNums.fptxtChoice.Text = "exit" Then
      cmdBegin.Enabled = True
      cmdExit.Enabled = True
      Close
      CheckData = False
      Unload frmBLCheck4BlankLicNums
      Exit Function
    Else
      Unload frmBLCheck4BlankLicNums
    End If
  End If
  
  DoEvents
  
  If Check4DupLicNums() = True Then
    frmBLConvertDupLicNums.Show vbModal
    DoEvents
    If frmBLConvertDupLicNums.fptxtChoice.Text = "exit" Then
     cmdBegin.Enabled = True
      cmdExit.Enabled = True
      Close
      CheckData = False
      Unload frmBLConvertDupLicNums
      Exit Function
    Else
      Unload frmBLConvertDupLicNums
    End If
  End If

  DoEvents
  
  fptxtMarque.Text = "Examining customer license numbers. "
  DoEvents
  
  If Check4NonNums() = True Then
    frmBLConvertNonNumbLic.Show vbModal
    DoEvents
    If frmBLConvertNonNumbLic.fptxtChoice.Text = "exit" Then
      cmdBegin.Enabled = True
      cmdExit.Enabled = True
      Close
      CheckData = False
      Unload frmBLConvertNonNumbLic
      Exit Function
    Else
      Unload frmBLConvertNonNumbLic
    End If
  End If
  
  If Version = 1 Then
    fptxtMarque.Text = "Looking for customers with duplicate category codes. "
    DoEvents
    If Check4CustDupCats = True Then
      frmBLCnvtDupCustCats.Show vbModal
      DoEvents
      If frmBLCnvtDupCustCats.fptxtChoice.Text = "exit" Then
        cmdBegin.Enabled = True
        cmdExit.Enabled = True
        Close
        CheckData = False
        Unload frmBLCnvtDupCustCats
        Exit Function
      Else
        Unload frmBLCnvtDupCustCats
        
      End If
    End If
  End If
  
  fptxtMarque.Text = "Pre-conversion data check-up has completed successfully. "
  DoEvents

End Function
