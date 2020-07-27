VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmVoidPurchase 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Void Decal Purchase"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   Icon            =   "frmPaymentDelete1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CheckBox OtherDate 
      BackColor       =   &H008F8265&
      Caption         =   "Use this date instead of orignial entry date."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2304
      TabIndex        =   17
      Top             =   2064
      Width           =   5244
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Esc E&xit"
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
      Left            =   8310
      TabIndex        =   1
      Top             =   6960
      Width           =   1596
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "F10 &Void"
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
      Left            =   6582
      TabIndex        =   0
      Top             =   6960
      Width           =   1596
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21548
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7154
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "12:23 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "3/7/2006"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin EditLib.fpLongInteger fpCustRecNo 
      Height          =   300
      Left            =   1200
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1344
      Visible         =   0   'False
      Width           =   684
      _Version        =   196608
      _ExtentX        =   1206
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   1
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   -1  'True
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
      ControlType     =   0
      Text            =   ""
      MaxValue        =   "2147483647"
      MinValue        =   "-2147483648"
      NegFormat       =   1
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   1
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   4956
      Left            =   2172
      TabIndex        =   6
      Top             =   2544
      Width           =   7884
      _Version        =   196609
      _ExtentX        =   13906
      _ExtentY        =   8742
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmPaymentDelete1.frx":08CA
      Begin LpLib.fpList fpListTrans 
         Height          =   2124
         Left            =   216
         TabIndex        =   7
         Top             =   1488
         Width           =   7452
         _Version        =   196608
         _ExtentX        =   13144
         _ExtentY        =   3746
         TextAlias       =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   10.8
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
         Sorted          =   2
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
         ColDesigner     =   "frmPaymentDelete1.frx":08E6
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Date "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   0
         Left            =   312
         TabIndex        =   14
         Top             =   1080
         Width           =   1668
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Spacebar or Click to Toggle, F10 to Continue. "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   468
         Left            =   144
         TabIndex        =   13
         Top             =   4488
         Width           =   4140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type/Desc "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   7
         Left            =   2880
         TabIndex        =   12
         Top             =   1080
         Width           =   2724
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   1
         Left            =   6360
         TabIndex        =   11
         Top             =   1080
         Width           =   1044
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   2940
         Left            =   120
         Top             =   1320
         Width           =   7668
      End
      Begin VB.Label LblAcct 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Left            =   360
         TabIndex        =   10
         Top             =   168
         Width           =   4068
      End
      Begin VB.Label LblName 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   360
         TabIndex        =   9
         Top             =   432
         Width           =   5628
      End
      Begin VB.Label LblSSN 
         BackStyle       =   0  'Transparent
         Caption         =   "SSN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Left            =   384
         TabIndex        =   8
         Top             =   696
         Width           =   4284
      End
   End
   Begin EditLib.fpDateTime txtUseDate 
      Height          =   324
      Left            =   8328
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1548
      _Version        =   196608
      _ExtentX        =   2730
      _ExtentY        =   572
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      ControlType     =   0
      Text            =   "10/03/2001"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
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
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Index           =   1
      Left            =   7416
      TabIndex        =   16
      Top             =   2088
      Width           =   816
   End
   Begin VB.Label LabelHead 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Decal Transaction to Void"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3084
      TabIndex        =   4
      Top             =   1104
      Width           =   6012
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   852
      Left            =   2580
      Top             =   840
      Width           =   7020
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Invoices For Payment"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12
      Left            =   3684
      TabIndex        =   3
      Top             =   1080
      Width           =   4836
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   996
      Left            =   2592
      Top             =   720
      Width           =   7020
   End
End
Attribute VB_Name = "frmVoidPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim CashFlag As Boolean, uselook As Boolean, CustAcct As Long
Dim EditFlag As Boolean, TempAmtRecv As Double, Answer As Integer
Dim ChkOKFlag As Boolean, BeenDone As Boolean
Dim codeopt As Integer, noreset As Boolean
Dim Oper As String
Dim fromform As Form, toform As Form
Dim dontdoit As Boolean
Public Sub Wheretogo(xfrm As Form, tfrm As Form, Optional opt As Integer)
  Set fromform = xfrm
  Set toform = tfrm
  If opt <> 0 Then
    codeopt = opt
  Else
    codeopt = 0
  End If
End Sub


Private Sub Form_Activate()
  If Val(frmVoidPurchase.fpCustRecNo) > 0 And Not BeenDone Then
    BeenDone = True
    LoadCustInfo
    DoEvents
  End If
End Sub
Private Sub LoadCustInfo()
  Dim NumOfDCRecs As Long, DCFile As Integer, AccountRecord As Long
  ReDim DCCustRec(1) As DCCustRecType
  OpenDCCustFile NumOfDCRecs, DCFile
  AccountRecord = fpCustRecNo
  Get DCFile, AccountRecord, DCCustRec(1)
  Close DCFile
  LblAcct.Caption = "Account Number - " + RTrim$(DCCustRec(1).CUSTNUMB)
  LblName.Caption = "Customer Name - " + QPTrim$(DCCustRec(1).BILLNAME)
  LblSSN.Caption = "Social Security Number - " + DCCustRec(1).SOSEC
  'If DCCustRec(1).FirstTrans > 0 Then
    ListCustTrans
 ' Else
  '  MsgBox "No trans found", vbOKOnly
 ' End If
 Me.HelpContextID = hlpVoidDecal
End Sub

Private Sub cmdExit_Click()
  DCLog "OUT: Decal Void" + " Oper:" + Oper$
  Exitvoid
  Unload Me
  DoEvents
End Sub
Private Sub Exitvoid()
On Local Error Resume Next
  DoEvents
  BeenDone = False
  fpCustRecNo = 0
  DoEvents
  If codeopt = 1 Then
    ActivateControls frmCustEditLookUP
  ElseIf codeopt = 2 Then
    ActivateControls frmDisplayList
  End If
  If codeopt = 0 Then
    frmDCCustomerMenu.Show
  End If
  DCLog PWUser + " Exit VoidDecal"
  Unload Me
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        DCLog "Closed via Decal Void by " + PWUser$ + " operator-" + Oper$
        CitiTerminate
      End If
    End If
  End If

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  Select Case KeyCode
'    Case vbKeyDown, vbKeyReturn:
'      SendKeys "{Tab}"
'      KeyCode = 0
'    Case vbKeyUp:
'      SendKeys "+{Tab}"
'      KeyCode = 0
    Case vbKeyEscape:
      KeyCode = 0
      DoEvents
      Call cmdExit_Click
    Case vbKeyF10:
      KeyCode = 0
      DoEvents
      'Call cmdDelete_Click
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  DCLog " IN Oper " + Oper$ + ": Decal Void"
  txtUseDate = Format(Now, "mm/dd/yyyy")
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ' Me.Visible = False
    Temp_Class.ResizeControls Me
   ' Me.Visible = True
   ' Me.SetFocus
  End If
  DoEvents
End Sub
Private Sub ListCustTrans()
  ReDim DCTranRec(1) As DCTransRecType
  ReDim DCCustRec(1) As DCCustRecType
  Dim DCCustRecLen As Integer, DCTranRecLen As Integer
  Dim PrevTranRec As Long
  Dim DCFile As Integer, dcnt As Integer
  Dim Build As String * 80
  Dim TType As String, TDesc As String
  Dim CurBal As Double
  
  frmInfo.Label1 = "Loading. . ."
  frmInfo.Show
  DoEvents

  DCCustRecLen = Len(DCCustRec(1))
  DCTranRecLen = Len(DCTranRec(1))
  
  DCFile = FreeFile
  Open DCPath + "DCCust.dat" For Random Shared As DCFile Len = DCCustRecLen
  Get DCFile, fpCustRecNo, DCCustRec(1)
  Close DCFile

  CurBal# = DCCustRec(1).AcctBal
'
Top:
'
  DCFile = FreeFile
  Open DCPath + "DCTRANS.DAT" For Random Shared As DCFile Len = DCTranRecLen
  
  PrevTranRec& = DCCustRec(1).FirstTrans
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
    Get DCFile, PrevTranRec&, DCTranRec(1)
    '''If DCTranRec(1).TransType < 1 Then Stop

      If DCTranRec(1).TransType = 2 And DCTranRec(1).VoidFlag <> "Y" Then
        dcnt = dcnt + 1
        
        LSet Build = Str(DCTranRec(1).TransDate) + Chr9$ + " " + Num2Date(DCTranRec(1).TransDate)
        GoSub GetTransType
        Mid$(Build, 22) = TType$
        Mid$(Build, 42) = TDesc$
        Mid$(Build, 64) = Chr9$ + Using("#####.##", DCTranRec(1).TransAmount, True)
        Mid$(Build$, 73) = Chr9$ + Str$(PrevTranRec&) 'Mid$(MChoice(DCnt).V, 61) = MKL$(PrevTranRec&)
        fpListTrans.AddItem Build$
      End If
      PrevTranRec& = DCTranRec(1).NextTrans
    Loop
  End If
  Close DCFile

Exit Sub

GetTransType:
  Select Case DCTranRec(1).TransType
  Case 1 'Charge
    TType$ = "Decal Charge"
  Case 2 'Payment
    TType$ = "Decal Payment"
  Case 3  'Charge Void
    TType$ = "Void Charge"
  Case 4  'Payment Void
    TType$ = "Void Payment"
  Case Else
    TType$ = Str$(DCTranRec(1).TransType) + " ???"
  End Select
  TDesc$ = QPTrim$(DCTranRec(1).TRVinDesc)
Return

End Sub
Private Sub cmdDelete_Click() ''VOID
  Dim FntSize As Integer, cnt As Long
  Dim PCnt As Integer, NumPicked As Integer
  ReDim MsgText(0 To 5) As String
  On Local Error GoTo ERRORSTUFF
  For PCnt = 0 To fpListTrans.ListCount - 1
    If fpListTrans.Selected(PCnt) Then
      NumPicked = NumPicked + 1
    End If
  Next
  If Not NumPicked > 0 Then
    frmMsgDialog.RetLabel = "-2"
    FntSize = frmMsgDialog.Label(2).FontSize
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = ""
    MsgText(3) = "NO TRANSACTIONS SELECTED!"
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
    Exit Sub
  ElseIf NumPicked > 0 Then
    If MsgBox("Are You Sure You Wish to Continue With Void?", vbYesNo, "Void") = vbYes Then
      For PCnt = 0 To fpListTrans.ListCount - 1
        If fpListTrans.Selected(PCnt) Then
          fpListTrans.col = 3
          fpListTrans.Row = PCnt
          cnt = QPTrim(fpListTrans.ColList)
          Exit For
        End If
      Next
      DCLog "Void Decal " + Str(cnt) + " Decal" + " Oper:" + Oper$
    Else
      Exit Sub
    End If
  End If
  If cnt > 0 Then
    PostVoid (cnt)
    MsgBox "Void Complete", vbOKOnly, "Complete"
  End If
  cmdExit_Click
Exit Sub
ERRORSTUFF:
  DCLog PWUser + " Error " + Str(Err.Number) + " DCVoidPurchase, cmdDelete"
  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "DCVoidPurchase", "cmdDelete", Erl)
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
    Unload Me

End Sub

Private Sub PostVoid(cnt As Long)
  Dim DCFile As Integer, NumOfDCRecs As Long
  Dim DCTransRecLen As Integer
  Dim DCTransFile As Integer, NumOfTransRecs As Long, NextTransRec As Long
  Dim Prev As Long, DCVehReclen As Integer, DCvFile As Integer
  Dim NumOfVRecs As Long, VehRecord As Long, UseDate As Integer
  ReDim DCCustRec(1) As DCCustRecType
  OpenDCCustFile NumOfDCRecs, DCFile
  ReDim DCVRec(1) As DCVehType
  ReDim DCTransRec(1 To 2) As DCTransRecType
  If OtherDate.Value = 1 Then
    UseDate = Date2Num(txtUseDate)
  End If
  
  DCTransRecLen = Len(DCTransRec(1))
  DCTransFile = FreeFile
  Open "DCTrans.DAT" For Random Access Read Write Shared As DCTransFile Len = DCTransRecLen
  NumOfTransRecs = LOF(DCTransFile) \ DCTransRecLen
  NextTransRec = NumOfTransRecs + 1

    Get DCTransFile, cnt, DCTransRec(1)
    If DCTransRec(1).TransAmount >= 0 And Val(DCTransRec(1).CustomerNumber) > 0 Then
      DCTransRec(1).VoidFlag = "Y"
      'here is the problem !##(*@&($&(#&(@#&(@#&(*#&@(#&@(
      '\/  NO NO Change the file structure to contain the veh record in each trans 9/23/05
      VehRecord = DCTransRec(1).VehRecord
      Put DCTransFile, cnt, DCTransRec(1)
      If VehRecord > 0 Then GoSub UpdateVehRecord
     ' Close DCTransFile
     ' LSet DCTransRec(2) = DCTransRec(1)

      Get DCFile, Val(DCTransRec(1).CustomerNumber), DCCustRec(1)
      ' Post Void Charge First to Offset Void Payment of Decal

      DCTransRec(2).CustomerNumber = DCTransRec(1).CustomerNumber
      If OtherDate.Value = 1 Then
        DCTransRec(2).TransDate = UseDate
      Else
        DCTransRec(2).TransDate = DCTransRec(1).TransDate
      End If
      DCTransRec(2).TransAmount = DCTransRec(1).TransAmount
      DCTransRec(2).TransType = 3               ' Type 3 = Void Charge
      DCTransRec(2).TRVinDesc = DCTransRec(1).TRVinDesc
      DCTransRec(2).TransTender = DCTransRec(1).TransTender
      DCTransRec(2).CashAmount = DCTransRec(1).CashAmount
      DCTransRec(2).ChkAmount = DCTransRec(1).ChkAmount
      DCTransRec(2).BalanceAfterTrans = DCCustRec(1).AcctBal - DCTransRec(1).TransAmount
      DCTransRec(2).makemodel = DCTransRec(1).makemodel
      DCTransRec(2).StateTag = DCTransRec(1).StateTag
      DCTransRec(2).Sticker = DCTransRec(1).Sticker
      DCTransRec(2).ExpireDate = DCTransRec(1).ExpireDate
      DCTransRec(2).ExtraDesc = "DC-Void Charge"
      DCTransRec(2).ExtraRoom = ""
      DCTransRec(2).NextTrans = 0
      DCTransRec(2).OperNum = Val(Oper$)
      DCTransRec(2).GLInterfaced = "Y"
      DCTransRec(2).DecalCat = DCTransRec(1).DecalCat
      DCTransRec(2).ChkByte = Chr$(1)
      DCTransRec(2).VoidFlag = "N"
      DCTransRec(2).VehRecord = DCTransRec(1).VehRecord
'      DCTransFile = FreeFile
'      Open "DCTrans.DAT" For Random Access Read Write Shared As DCTransFile Len = DCTransRecLen
'      NumOfTransRecs = LOF(DCTransFile) \ DCTransRecLen
'      NextTransRec = NumOfTransRecs + 1

      Put DCTransFile, NextTransRec, DCTransRec(2)
      
      Get DCFile, Val(DCTransRec(1).CustomerNumber), DCCustRec(1)
      DCCustRec(1).AcctBal = DCCustRec(1).AcctBal - DCTransRec(1).TransAmount
      Put DCFile, Val(DCTransRec(1).CustomerNumber), DCCustRec(1)
      If DCCustRec(1).FirstTrans = 0 Then
        DCCustRec(1).FirstTrans = NextTransRec
        DCCustRec(1).LastTrans = NextTransRec
        Put DCFile, Val(DCTransRec(1).CustomerNumber), DCCustRec(1)
      Else
        Prev = DCCustRec(1).LastTrans
        DCCustRec(1).LastTrans = NextTransRec
        Put DCFile, Val(DCTransRec(1).CustomerNumber), DCCustRec(1)
        Get DCTransFile, Prev, DCTransRec(1)
        DCTransRec(1).NextTrans = NextTransRec
        Put DCTransFile, Prev, DCTransRec(1)
      End If
      Close DCTransFile
      DCTransFile = FreeFile
      Open "DCTrans.DAT" For Random Access Read Write Shared As DCTransFile Len = DCTransRecLen
      NumOfTransRecs = LOF(DCTransFile) \ DCTransRecLen
      NextTransRec = NumOfTransRecs + 1
      Get DCTransFile, NumOfTransRecs, DCTransRec(1)
      ' Post Transaction Record First
      DCTransRec(2).CustomerNumber = DCTransRec(1).CustomerNumber
      If OtherDate.Value = 1 Then
        DCTransRec(2).TransDate = UseDate
      Else
        DCTransRec(2).TransDate = DCTransRec(1).TransDate
      End If
      DCTransRec(2).TransAmount = DCTransRec(1).TransAmount
      DCTransRec(2).TransType = 4               ' Type 4 = Void Payment
      DCTransRec(2).TRVinDesc = DCTransRec(1).TRVinDesc
      DCTransRec(2).TransTender = DCTransRec(1).TransTender
      DCTransRec(2).CashAmount = DCTransRec(1).CashAmount
      DCTransRec(2).ChkAmount = DCTransRec(1).ChkAmount
      DCTransRec(2).BalanceAfterTrans = DCCustRec(1).AcctBal + DCTransRec(1).TransAmount
      DCTransRec(2).ExtraDesc = "DC-Void Payment"
      DCTransRec(2).ExtraRoom = ""
      DCTransRec(2).NextTrans = 0
      DCTransRec(2).GLInterfaced = "N"
      DCTransRec(2).OperNum = Val(Oper$)
      DCTransRec(2).DecalCat = DCTransRec(1).DecalCat
      DCTransRec(2).ChkByte = Chr$(1)
      DCTransRec(2).VoidFlag = "N"
      DCTransRec(2).VehRecord = DCTransRec(1).VehRecord
      Put DCTransFile, NextTransRec, DCTransRec(2)
      
      Get DCFile, Val(DCTransRec(1).CustomerNumber), DCCustRec(1)
      DCCustRec(1).AcctBal = DCCustRec(1).AcctBal + DCTransRec(1).TransAmount
      DCCustRec(1).LICENSE = ""
      Put DCFile, Val(DCTransRec(1).CustomerNumber), DCCustRec(1)

      If DCCustRec(1).FirstTrans = 0 Then
        DCCustRec(1).FirstTrans = NextTransRec
        DCCustRec(1).LastTrans = NextTransRec
        Put DCFile, Val(DCTransRec(1).CustomerNumber), DCCustRec(1)
      Else
        Prev = DCCustRec(1).LastTrans
        DCCustRec(1).LastTrans = NextTransRec
        Put DCFile, Val(DCTransRec(1).CustomerNumber), DCCustRec(1)
        Get DCTransFile, Prev, DCTransRec(1)
        DCTransRec(1).NextTrans = NextTransRec
        Put DCTransFile, Prev, DCTransRec(1)
      End If
      Close DCTransFile
   
    End If

  Close
  ' Show All Posted
  DCLog "Voided:" + Str$(cnt)
 ' MsgBox "Void Complete", vbOKOnly, "Complete"
  Close
  Exit Sub

UpdateVehRecord:
  DCVehReclen = Len(DCVRec(1))
  DCvFile = FreeFile
  Open "DCVEH.DAT" For Random Access Read Write Shared As DCvFile Len = DCVehReclen
  NumOfVRecs = LOF(DCvFile) \ DCVehReclen
  If VehRecord <= 0 Or VehRecord > NumOfVRecs Then Close DCvFile: Return
  Get DCvFile, VehRecord, DCVRec(1)
  DCVRec(1).ExpireDate = Date2Num("01/01/1980")
  DCVRec(1).Sticker = "VOID"
  Put DCvFile, VehRecord, DCVRec(1)
  Close DCvFile
Return

End Sub

