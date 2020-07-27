VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmDepositRefund 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   2172
   ClientWidth     =   12216
   Icon            =   "frmDepositRefund.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn fpCmdExit 
      Height          =   480
      Left            =   4254
      TabIndex        =   0
      Top             =   5736
      Width           =   1548
      _Version        =   131072
      _ExtentX        =   2730
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmDepositRefund.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdRefund 
      Height          =   480
      Left            =   6432
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5736
      Width           =   1548
      _Version        =   131072
      _ExtentX        =   2730
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
      DrawFocusRect   =   1
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
      ButtonDesigner  =   "frmDepositRefund.frx":0AA8
   End
   Begin EditLib.fpText fpCustName 
      CausesValidation=   0   'False
      Height          =   324
      Left            =   4320
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4392
      Width           =   4308
      _Version        =   196608
      _ExtentX        =   7599
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
      AlignTextH      =   0
      AlignTextV      =   2
      AllowNull       =   0   'False
      NoSpecialKeys   =   3
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
      MarginTop       =   0
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
      MaxLength       =   35
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
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
   Begin EditLib.fpDoubleSingle fpDeposit 
      Height          =   324
      Left            =   4320
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4848
      Width           =   1308
      _Version        =   196608
      _ExtentX        =   2307
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
      NoSpecialKeys   =   3
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   2
      CaretOverWrite  =   2
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
      OnFocusPosition =   1
      ControlType     =   1
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
      NegFormat       =   1
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   8
      Top             =   8568
      Width           =   12216
      _ExtentX        =   21548
      _ExtentY        =   529
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
            TextSave        =   "12:30 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "6/23/2005"
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
   Begin EditLib.fpText fpCustRecNo 
      Height          =   324
      Left            =   744
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2376
      Visible         =   0   'False
      Width           =   1764
      _Version        =   196608
      _ExtentX        =   3111
      _ExtentY        =   572
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   3
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
      Text            =   "fpText1"
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
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   732
      Left            =   2880
      Top             =   1008
      Width           =   6468
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Refund Customer Deposit"
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
      Left            =   3708
      TabIndex        =   9
      Top             =   1200
      Width           =   4812
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      BorderWidth     =   3
      X1              =   2832
      X2              =   9360
      Y1              =   2376
      Y2              =   2376
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      BorderWidth     =   3
      X1              =   2832
      X2              =   2832
      Y1              =   2376
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2832
      X2              =   9360
      Y1              =   5448
      Y2              =   5448
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LAST CHANCE.   ARE YOU SURE YOU WANT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   348
      Index           =   1
      Left            =   3120
      TabIndex        =   5
      Top             =   3048
      Width           =   6000
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TO REFUND THIS CUSTOMERS DEPOSIT?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   348
      Index           =   2
      Left            =   3144
      TabIndex        =   4
      Top             =   3480
      Width           =   6000
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NAME:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   372
      Index           =   3
      Left            =   3336
      TabIndex        =   3
      Top             =   4440
      Width           =   816
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DEPOSIT:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   324
      Index           =   6
      Left            =   2928
      TabIndex        =   2
      Top             =   4872
      Width           =   1224
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   4116
      Left            =   2826
      Top             =   2376
      Width           =   6564
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   852
      Left            =   2868
      Top             =   912
      Width           =   6492
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuPrnScn 
         Caption         =   "Prin&t Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmDepositRefund"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim CustAcct As Long
Dim BeenDone As Boolean
Dim fromform As Form, toform As Form, codeopt As Integer
Dim uselook As Boolean, Answer As Integer, CredOKFlag As Boolean
Dim CleveFlag As Boolean, TotalBalance As Double, TempDepAmt As Double

Public Sub Wheretogo(xfrm As Form, tfrm As Form, Optional opt As Integer)
  Set fromform = xfrm
  Set toform = tfrm
  If opt <> 0 Then
    codeopt = opt
  Else
    codeopt = 0
  End If
  uselook = True
End Sub

Private Sub cmdExit_Click()
  
  CustAcct = 0
  fpCustRecNo = 0
  BeenDone = False
  If codeopt = 1 Then
    ActivateControls frmCustEditLookUP
  ElseIf codeopt = 2 Then
    ActivateControls frmDisplayList
  End If
  If codeopt = 0 Then
    Load frmUBDepositMenu
    DoEvents
    frmUBDepositMenu.Show
  End If

  UBLog "OUT: UTIL DepRefund"
  Unload Me
  DoEvents
End Sub
Private Sub Form_Activate()
  If Val(fpCustRecNo) > 0 And Not BeenDone Then
    BeenDone = True
    loadCustrec
    DoEvents
  End If
  
End Sub

'Private Sub fpCmdSave_Click()
'  CalcBALFlds
'  CheckApplyInfo
'  If CredOKFlag Then
'    If MsgBox("Are you sure you wish to save this transaction?", vbYesNo, "Save Transaction") = vbYes Then
'      SaveTransaction
'      CustAcct = 0
'      fpCustRecNo = 0
'      BeenDone = False
'      If codeopt = 1 Then
'        ActivateControls frmCustEditLookUP
'      ElseIf codeopt = 2 Then
'        ActivateControls frmDisplayList
'      End If
'      If codeopt = 0 Then
'        Load frmUBDepositMenu
'        DoEvents
'        frmUBDepositMenu.Show
'      End If
'
'      UBLog "OUT: UTIL DepCredRem"
'      Unload Me
'      DoEvents
'    End If
'  End If
'End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If fpCmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via DepRefund by " + PWUser$
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
      Call fpCmdRefund_Click
    Case Else:
  End Select
End Sub


Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  UBLog " IN: UTIL DepRefund"
  Me.HelpContextID = hlpRefundCustomer
End Sub

Private Sub Form_Resize()
  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
  End If
End Sub

Private Sub loadCustrec()
  Dim UBCustRecLen As Integer, NumOfCustRecs As Long
  Dim CustFile As Integer, cnt As Integer
  ReDim UBCustRec(1) As NewUBCustRecType
  Dim NumofRevs As Integer, RevCnt As Integer
  NumofRevs = MaxRevsCnt
  UBCustRecLen = Len(UBCustRec(1))
  CustAcct = fpCustRecNo
  NumOfCustRecs& = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen
  CustFile = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CustFile Len = UBCustRecLen
  Get CustFile, CustAcct&, UBCustRec(1)
  Close CustFile
  fpCustName = UBCustRec(1).CustName
  
  fpDeposit = UBCustRec(1).DepositAmt
  
End Sub
Private Sub fpCmdRefund_Click()
  DoDepositRefund
  MsgBox "REFUND procedure complete.", vbOKOnly, "Completed"
  cmdExit_Click
End Sub
   

Private Sub DoDepositRefund()
  Dim UBTransRecLen As Integer, NextTranRecs As Long
  Dim TransDate As Integer, Transamt As Double, CustChCnt As Integer
  Dim UBCustRecLen As Integer, NumOfCustRecs As Long
  Dim CustFile  As Integer, cnt As Integer, RevCnt As Integer
  Dim UBTran As Integer, NumOfTranRecs As Long, PrevLastTrans As Long
  Dim TotalDepAmt As Double, LastTran As Long
  ReDim RevAmts(1 To 15) As Double
  ReDim UBCustRec(1 To 2) As NewUBCustRecType
  CustFile = FreeFile
  UBCustRecLen = Len(UBCustRec(1))
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CustFile Len = UBCustRecLen
  Get CustFile, CustAcct&, UBCustRec(1)
  Close CustFile

  
  GoSub GetDepRevAmts
  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))

  TransDate = Date2Num(Date$)
  Transamt# = -UBCustRec(1).DepositAmt
  UBTransRec(1).OperatorNumber = OPERNUM
  UBTransRec(1).TransDate = TransDate
  'UBTransRec(1)CustLocation = RecNo&
  UBTransRec(1).CustStatus = UBCustRec(1).Status
  UBTransRec(1).CustAcctNo = CustAcct&
  UBTransRec(1).Transamt = Transamt#
  UBTransRec(1).TransDesc = "Refunded Deposit"
  UBTransRec(1).VoidFlag = False
  UBTransRec(1).FromCMFlag = False
  UBTransRec(1).TransType = TranRefundDeposit
  UBTransRec(1).RunBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
  UBCustRec(1).DepositAmt = 0

  For RevCnt = 1 To 15
    UBTransRec(1).RevAmt(RevCnt) = RevAmts(RevCnt)
  Next

  CustFile = FreeFile
  Open "UBCUST.DAT" For Random Shared As CustFile Len = UBCustRecLen
  UBTran = FreeFile
  Open "UBTRANS.DAT" For Random Shared As UBTran Len = UBTransRecLen

  NextTranRecs& = (LOF(UBTran) \ UBTransRecLen) + 1
  PrevLastTrans& = UBCustRec(1).LastTrans
  UBTransRec(1).PrevTrans = PrevLastTrans&
  UBCustRec(1).LastTrans = NextTranRecs&

'remark these
'*******************************
'''  UBTransRec(1).TransDate = Date2Num("01-31-2001")
'''  UBTransRec(1).TransDesc = "Applied Deposit"
'''  UBTransRec(1).TransType = TranAppliedDeposit
'''  UBTransRec(1).Transamt = TotalDepAmt#
'*******************************

  Put CustFile, CustAcct&, UBCustRec(1)
  Put UBTran, NextTranRecs&, UBTransRec(1)
  Close UBTran, CustFile

Exit Sub

GetDepRevAmts:
  TotalDepAmt# = 0
  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))
  LastTran& = UBCustRec(1).LastTrans
  If LastTran& > 0 Then
    UBTran = FreeFile
    Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTran Len = UBTransRecLen
    Do
      Get #UBTran, LastTran&, UBTransRec(1)
      If UBTransRec(1).TransType = TranDepositPayment Then
        For RevCnt = 1 To 15
          If UBTransRec(1).RevAmt(RevCnt) > 0 Then
            RevAmts(RevCnt) = Round#(RevAmts(RevCnt) + UBTransRec(1).RevAmt(RevCnt))
            TotalDepAmt# = Round#(TotalDepAmt# + UBTransRec(1).RevAmt(RevCnt))
          End If
        Next
      ElseIf (UBTransRec(1).TransType = TranAppliedDeposit) Or (UBTransRec(1).TransType = TranRefundDeposit) Or (UBTransRec(1).TransType = TranDepPaymentVoid) Then
        For RevCnt = 1 To 15
          If UBTransRec(1).RevAmt(RevCnt) > 0 Then
            RevAmts(RevCnt) = Round#(RevAmts(RevCnt) - UBTransRec(1).RevAmt(RevCnt))
            TotalDepAmt# = Round#(TotalDepAmt# - UBTransRec(1).RevAmt(RevCnt))
          End If
        Next
      End If
      LastTran& = UBTransRec(1).PrevTrans
    Loop While LastTran& > 0
    Close UBTran
  End If

Return
End Sub

