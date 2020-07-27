VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmVATaxPRateTableFlatOnly 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personal Optional Revenue Rate Setup"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "frmVATaxPRateTableFlatOnly.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H008F8265&
      Height          =   1455
      Left            =   3239
      TabIndex        =   4
      Top             =   2736
      Width           =   5175
      Begin VB.OptionButton OptRev2 
         BackColor       =   &H008F8265&
         Caption         =   "Option2"
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
         Height          =   285
         Left            =   360
         TabIndex        =   7
         Top             =   600
         Width           =   4695
      End
      Begin VB.OptionButton OptRev3 
         BackColor       =   &H008F8265&
         Caption         =   "Option3"
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
         Height          =   285
         Left            =   360
         TabIndex        =   6
         Top             =   960
         Width           =   4695
      End
      Begin VB.OptionButton OptRev1 
         BackColor       =   &H008F8265&
         Caption         =   "Option1"
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
         Height          =   315
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   4575
      End
   End
   Begin EditLib.fpText fptxtDescription 
      Height          =   372
      Left            =   5100
      TabIndex        =   0
      Top             =   4416
      Width           =   3132
      _Version        =   196608
      _ExtentX        =   5524
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
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      ControlType     =   1
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   20
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
   Begin EditLib.fpCurrency fpCurrFlat 
      Height          =   372
      Left            =   5640
      TabIndex        =   2
      Top             =   4896
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
      _ExtentY        =   656
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
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
      Height          =   420
      Left            =   2490
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6930
      Width           =   1935
      _Version        =   131072
      _ExtentX        =   3413
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmVATaxPRateTableFlatOnly.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   420
      Left            =   7230
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6930
      Width           =   1935
      _Version        =   131072
      _ExtentX        =   3413
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmVATaxPRateTableFlatOnly.frx":0AA8
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEdit 
      Height          =   420
      Left            =   4860
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6930
      Width           =   1935
      _Version        =   131072
      _ExtentX        =   3413
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmVATaxPRateTableFlatOnly.frx":0C84
   End
   Begin EditLib.fpText fptxtComment 
      Height          =   372
      Left            =   3840
      TabIndex        =   3
      Top             =   5400
      Width           =   5052
      _Version        =   196608
      _ExtentX        =   8911
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
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      MaxLength       =   35
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comment:"
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
      Height          =   372
      Left            =   2640
      TabIndex        =   13
      Top             =   5520
      Width           =   1092
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   3372
      Left            =   2580
      Top             =   2616
      Width           =   6492
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Flat Rate:"
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
      Height          =   252
      Left            =   4440
      TabIndex        =   9
      Top             =   4968
      Width           =   1092
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Rate Description:"
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
      Height          =   372
      Left            =   3180
      TabIndex        =   8
      Top             =   4476
      Width           =   1812
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Optional Personal Revenue Rate Tables"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   2940
      TabIndex        =   1
      Top             =   1650
      Width           =   6012
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1500
      Top             =   1470
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   720
      Left            =   1500
      Top             =   1410
      Width           =   8652
   End
End
Attribute VB_Name = "frmVATaxPRateTableFlatOnly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class
  Dim EditMode As Boolean
  Dim TempDesc$
  Dim TempComment$
  Dim TempOptRevNum As Integer
  Dim TempType$
  Dim TempStepType$
  Dim TempFromAmt() As Double
  Dim TempToAmt() As Double
  Dim TempTaxFAmt() As Double
  Dim TempTaxPAmt() As Double
  Dim TempFlatAmt As Double
  Dim SaveHere As Integer
  Dim LoadNewRate As Boolean
  Dim ExitMode As Boolean
  Dim FirstLoad As Boolean
  
Private Sub cmdEdit_Click()
  LoadNewRate = True
  If Check4Changes = True Then
    LoadNewRate = False
    Exit Sub
  End If
  frmVATaxRateListPop.Show
End Sub

Private Sub cmdExit_Click()
  ExitMode = True
  If Check4Changes = True Then
    ExitMode = False
    Exit Sub
  End If
  KillFile "addptbl.dat"
  Unload frmVATaxRateListPop
  frmVATaxRateMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdSave_Click()
  Dim TblRec As OptRevRateTablesType
  Dim TRHandle As Integer
  Dim NumOfTRRecs As Integer
  Dim x As Integer
  
  On Error GoTo ERRORSTUFF
  
  If CheckReqFields = False Then
    Exit Sub
  End If
  
'  If Check4DupDescs = True Then Exit Sub
  
  OpenTaxRateTables TRHandle, NumOfTRRecs
  If EditMode = False Then
    SaveHere = NumOfTRRecs + 1
  Else
    SaveHere = RateTblRec
    Get TRHandle, SaveHere, TblRec
  End If
  
  TblRec.Cushion = ""
  TblRec.Deleted = False
  TblRec.Desc = QPTrim$(fptxtDescription.Text)
  TblRec.Comment = QPTrim$(fptxtComment.Text)
  TblRec.FlatAmt = CDbl(fpCurrFlat.Value)
  If OptRev1.Value = True Then
    TblRec.OptRevNum = 4
  ElseIf OptRev2.Value = True Then
    TblRec.OptRevNum = 5
  ElseIf OptRev3.Value = True Then
    TblRec.OptRevNum = 6
  End If
  TblRec.Type = "F"
  TblRec.RevType = "P"
  For x = 0 To 9
    TblRec.FromAmt(x + 1) = 0
    TblRec.ToAmt(x + 1) = 0
    TblRec.TaxFAmt(x + 1) = 0
    TblRec.TaxPAmt(x + 1) = 0
  Next x
  TblRec.StepType = "N"
  
  Put TRHandle, SaveHere, TblRec
  Close TRHandle
  
  Call Savemsg(900, "The personal tax rate was saved successfully.")
  If EditMode = False Then
    KillFile "addptbl.dat"
    frmVATaxRateMenu.Show
    DoEvents
    Unload Me
  Else
    frmVATaxRateListPop.Show
  End If
      
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxRateTables", "cmdSave_Click", Erl)
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
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      Call cmdSave_Click
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%L"
      Call cmdEdit_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  FirstLoad = True
  Call LoadMe
  FirstLoad = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      KillFile "addptbl.dat"
      Unload frmVATaxRateListPop
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxRateTables.")
      Call Terminate
      End
    End If
  End If

End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    'Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub LoadMe()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim ThisRev$
  Dim TblRec As OptRevRateTablesType
  Dim TRHandle As Integer
  Dim NumOfTRRecs As Integer
  Dim x As Integer
  Dim Opt1 As Boolean, Opt2 As Boolean, Opt3 As Boolean
  
  On Error GoTo ERRORSTUFF
  Me.HelpContextID = hlpAddNewRateCode
  
  Call EnableOpts
  SaveHere = 0
  LoadNewRate = False
  ExitMode = False
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  ThisRev = QPTrim$(TaxMasterRec.POptRev3)
  If ThisRev <> "" Then
    OptRev3.Caption = ThisRev
  Else
    OptRev3.Caption = "NOT BEING USED"
    OptRev3.Enabled = False
  End If
  
  ThisRev = QPTrim$(TaxMasterRec.POptRev2)
  If ThisRev <> "" Then
    OptRev2.Caption = ThisRev
  Else
    OptRev2.Caption = "NOT BEING USED"
    OptRev2.Enabled = False
  End If

  ThisRev = QPTrim$(TaxMasterRec.POptRev1)
  If ThisRev <> "" Then
    OptRev1.Caption = ThisRev
  Else
    OptRev1.Caption = "NOT BEING USED"
    OptRev1.Enabled = False
  End If
  
  Call LoadDesc
  
  If Exist("addptbl.dat") Then
    EditMode = False
    cmdEdit.Enabled = False
    GoTo AddNew
  Else
    cmdEdit.Enabled = True
    EditMode = True
  End If
  If Exist(TaxRateTableFile) Then
    OpenTaxRateTables TRHandle, NumOfTRRecs
    For x = 1 To NumOfTRRecs
      Get TRHandle, x, TblRec
      If TblRec.Deleted = False And TblRec.RevType = "P" Then
        Exit For
      End If
    Next x
    Close TRHandle
    
    If x > NumOfTRRecs Then
      Call TaxMsg(900, "There are no active personal rates on file.")
      GoTo AddNew
    End If
    fptxtDescription.Text = QPTrim$(TblRec.Desc)
    TempDesc$ = QPTrim$(TblRec.Desc)
    Select Case TblRec.OptRevNum
      Case 4:
        OptRev1.Value = True
        OptRev2.Value = False
        OptRev3.Value = False
        TempOptRevNum = 4
        OptRev1.Enabled = False
      Case 5:
        OptRev1.Value = False
        OptRev2.Value = True
        OptRev3.Value = False
        TempOptRevNum = 5
        OptRev2.Enabled = False
      Case 6:
        OptRev1.Value = False
        OptRev2.Value = False
        OptRev3.Value = True
        TempOptRevNum = 6
        OptRev3.Enabled = False
      End Select
    Select Case TblRec.Type
      Case "F":
        fpCurrFlat.Enabled = True
        fpCurrFlat = TblRec.FlatAmt
        TempFlatAmt = TblRec.FlatAmt
        TempType$ = "F"
    End Select
  Else
    If OptRev1.Enabled = True Then
      OptRev1.Value = True
    ElseIf OptRev2.Enabled = True Then
      OptRev2.Value = True
    ElseIf OptRev3.Enabled = True Then
      OptRev3.Value = True
    End If
  End If
  
  Unload frmVATaxRateListPop
  DoEvents

AddNew:
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxRateTables", "LoadMe", Erl)
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

Private Sub fpCurrFlat_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    If OptRev1.Enabled = True Then
      OptRev1.SetFocus
    ElseIf OptRev2.Enabled = True Then
      OptRev2.SetFocus
    ElseIf OptRev3.Enabled = True Then
      OptRev3.SetFocus
    Else
    fptxtDescription.SetFocus
    End If
  ElseIf KeyCode = vbKeyUp Then
    fptxtDescription.SetFocus
  End If
End Sub

Private Sub fptxtComment_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fptxtDescription.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    fpCurrFlat.SetFocus
  End If
End Sub

Private Sub fptxtDescription_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fpCurrFlat.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    fptxtComment.SetFocus
  End If
End Sub

Private Sub OptRev1_Click()
  If FirstLoad = True Then Exit Sub
  If OptRev1.Enabled = True Then
    fptxtDescription.Text = OptRev1.Caption
  End If
End Sub
Private Sub LoadDesc()
  Dim Opt1 As Boolean
  Dim Opt2 As Boolean
  Dim Opt3 As Boolean
  
  Opt1 = False
  Opt2 = False
  Opt3 = False
  
'  If OptRev1.Enabled = True Then
'    fptxtDescription.Text = OptRev1.Caption
'    Opt1 = True
'  ElseIf OptRev2.Enabled = True Then
'    fptxtDescription.Text = OptRev2.Caption
'    Opt2 = True
'  ElseIf OptRev3.Enabled = True Then
'    fptxtDescription.Text = OptRev3.Caption
'    Opt3 = True
'  End If
  
  If Opt1 = True And Opt2 = True And Opt3 = True Then
    OptRev1.Value = True
  ElseIf Opt1 = True And Opt2 = False And Opt3 = False Then
    OptRev1.Value = True
  ElseIf Opt1 = False And Opt2 = True And Opt3 = False Then
    OptRev2.Value = True
  ElseIf Opt1 = False And Opt2 = False And Opt3 = True Then
    OptRev3.Value = True
  ElseIf Opt1 = False And Opt2 = True And Opt3 = True Then
    OptRev2.Value = True
  ElseIf Opt1 = True And Opt2 = True And Opt3 = False Then
    OptRev1.Value = True
  ElseIf Opt1 = True And Opt2 = False And Opt3 = False Then
    OptRev1.Value = True
  End If
  
  
End Sub
Private Sub OptRev2_Click()
  If FirstLoad = True Then Exit Sub
  If OptRev2.Enabled = True Then
    fptxtDescription.Text = OptRev2.Caption
  End If

End Sub

Private Sub OptRev3_Click()
  If FirstLoad = True Then Exit Sub
  If OptRev3.Enabled = True Then
    fptxtDescription.Text = OptRev3.Caption
  End If
End Sub


Public Sub LoadMeEdit()
  Dim TblRec As OptRevRateTablesType
  Dim TRHandle As Integer
  Dim NumOfTRRecs As Integer
  Dim x As Integer
  
  On Error GoTo ERRORSTUFF
  Me.HelpContextID = hlpEditExistingRate

  SaveHere = 0
  LoadNewRate = False
  ExitMode = False
  OptRev1.Enabled = False
  OptRev2.Enabled = False
  OptRev3.Enabled = False
  If RateTblRec > 0 Then
    OpenTaxRateTables TRHandle, NumOfTRRecs
    Get TRHandle, RateTblRec, TblRec
    Close TRHandle
  End If
  If TblRec.OptRevNum = 4 Then
    TempOptRevNum = 4
    OptRev1.Value = True
    OptRev2.Value = False
    OptRev3.Value = False
  ElseIf TblRec.OptRevNum = 5 Then
    TempOptRevNum = 5
    OptRev1.Value = False
    OptRev2.Value = True
    OptRev3.Value = False
  ElseIf TblRec.OptRevNum = 6 Then
    TempOptRevNum = 6
    OptRev1.Value = False
    OptRev2.Value = False
    OptRev3.Value = True
  End If
   
  fptxtDescription.Text = QPTrim$(TblRec.Desc)
  TempDesc = QPTrim$(TblRec.Desc)
  fptxtComment.Text = QPTrim$(TblRec.Comment)
  TempComment = QPTrim$(TblRec.Comment)
  fpCurrFlat = TblRec.FlatAmt
  TempFlatAmt = TblRec.FlatAmt
  If TblRec.Type = "F" Then
    TempType = "F"
    TempStepType = "N"
    fpCurrFlat.Enabled = True
  End If
  If fpCurrFlat.Enabled = True Then
    fpCurrFlat.SetFocus
  End If
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxRateTables", "LoadMeEdit", Erl)
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

Private Sub ClearForm()
  Dim x As Integer
  
  OptRev1.Value = False
  OptRev2.Value = False
  OptRev3.Value = False
  fpCurrFlat.Value = 0
  fptxtDescription.Text = ""
  fptxtComment.Text = ""
End Sub

Public Function Check4Changes() As Boolean
  Dim TblRec As OptRevRateTablesType
  Dim TRHandle As Integer
  Dim x As Integer
  Dim NumOfTRRecs As Integer
  Dim Operator$
  Dim choice As String
  Dim ThisControl As Control
  Dim ThisDesc As String
  Dim ThatDesc As String
  Dim ThisText As String
  Dim ThisDbl As Double
  Dim ThatDbl As Double
  Dim ThisInt As Integer
  Dim ThatInt As Integer
  Dim Message$
  Dim CmdF10$
  Dim CmdESC$
  Dim CmdF5$
  
  On Error GoTo ERRORSTUFF
  Check4Changes = False
  If EditMode = False Then Exit Function
  If RateTblRec = 0 Then Exit Function
  
  CmdF10 = "F10 Save"
  CmdESC = "ESC Don't Save"
  CmdF5 = "F5 Review"
  
  OpenTaxRateTables TRHandle, NumOfTRRecs
  Get TRHandle, RateTblRec, TblRec
  
  If TempOptRevNum = 4 Then
    Set ThisControl = OptRev1
    ThatInt = 4
  ElseIf TempOptRevNum = 5 Then
    Set ThisControl = OptRev2
    ThatInt = 5
  ElseIf TempOptRevNum = 6 Then
    Set ThisControl = OptRev3
    ThatInt = 6
  End If
  If OptRev1.Value = True Then
    ThisInt = 4
  ElseIf OptRev2.Value = True Then
    ThisInt = 5
  ElseIf OptRev3.Value = True Then
    ThisInt = 6
  End If
  If ThisInt <> ThatInt Then
    Message = "A change has been made in the optional revenue from " + CStr(ThatInt) + " to " + CStr(ThisInt) + ". Press F10 to save, press ESC to abandon or press F5 to review."
    choice = TaxMsgW3Opts(800, Message, CmdF5, CmdF10, CmdESC)
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      If CheckReqFields = False Then
        Check4Changes = True
        Exit Function
      End If
      TblRec.OptRevNum = ThisInt
      Put TRHandle, RateTblRec, TblRec
    Else
      GoSub HandleChoice
    End If
  End If
  
  ThisDesc = "F"
  ThatDesc = TempType
  If ThatDesc <> ThisDesc Then
    Select Case ThisDesc
      Case "F"
        ThisDesc = "Flat Rate"
    End Select
    Select Case ThatDesc
      Case "F"
        ThatDesc = "Flat Rate"
    End Select
    Message = "A change has been made in the rate type from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save, press ESC to abandon or press F5 to review."
    choice = TaxMsgW3Opts(800, Message, CmdF5, CmdF10, CmdESC)
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      If CheckReqFields = False Then
        Check4Changes = True
        Exit Function
      End If
      Select Case ThisDesc
        Case "Flat Rate"
          TblRec.Type = "F"
          TblRec.StepType = "N"
      End Select
      Put TRHandle, RateTblRec, TblRec
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtComment
  ThisDesc = QPTrim$(fptxtComment.Text)
  ThatDesc = TempComment
  If QPTrim$(ThisDesc) <> QPTrim$(ThatDesc) Then
    Message = "A change has been made in the comment from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save, press ESC to abandon or press F5 to review."
    choice = TaxMsgW3Opts(800, Message, CmdF5, CmdF10, CmdESC)
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      If CheckReqFields = False Then
        Exit Function
      End If
      TblRec.Comment = ThisControl
      Put TRHandle, RateTblRec, TblRec
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtDescription
  ThisDesc = QPTrim$(fptxtDescription.Text)
  ThatDesc = TempDesc
  If QPTrim$(ThisDesc) <> QPTrim$(ThatDesc) Then
    Message = "A change has been made in the description from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save, press ESC to abandon or press F5 to review."
    choice = TaxMsgW3Opts(800, Message, CmdF5, CmdF10, CmdESC)
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      If CheckReqFields = False Then
        Exit Function
      End If
      TblRec.Desc = ThisControl
      Put TRHandle, RateTblRec, TblRec
    Else
      GoSub HandleChoice
    End If
  End If
    
  If fpCurrFlat.Enabled = True Then
    Set ThisControl = fpCurrFlat
    ThisDbl = CDbl(fpCurrFlat.Value)
    ThatDbl = TempFlatAmt
    If ThisDbl <> ThatDbl Then
      Message = "A change has been made in the flat rate from " + QPTrim$(Using("$##,##0.00", ThatDbl)) + " to " + QPTrim$(Using("$##,##0.00", ThisDbl)) + ". Press F10 to save, press ESC to abandon or press F5 to review."
      choice = TaxMsgW3Opts(800, Message, CmdF5, CmdF10, CmdESC)
      Unload frmVATaxMsgW3Opts
      If choice = "continue" Then
        If CheckReqFields = False Then
          Check4Changes = True
          Exit Function
        End If
        TblRec.FlatAmt = ThisControl
        Put TRHandle, RateTblRec, TblRec
      Else
        GoSub HandleChoice
      End If
    End If
  End If
  
  
  Exit Function

HandleChoice:
  Close TRHandle
  Select Case choice
    Case "abort" 'don't save
      If LoadNewRate = True Then 'switching to a new rate
        LoadNewRate = False
        Exit Function
      Else
        If TaxMsgWOpts(800, "If you wish to continue checking for unsaved changes then press F10. Otherwise, press ESC to exit to the Tax Rate Menu.", "F10 Keep Checking", "ESC Exit") = "abort" Then
          Unload frmVATaxMsgWOpts
          Unload frmVATaxRateListPop
          DoEvents
          frmVATaxRateMenu.Show
          DoEvents
          Unload Me
          Exit Function
        Else
          Unload frmVATaxMsgWOpts
          Unload frmVATaxRateListPop
        End If
      End If
    Case "option" 'review
      If ThisControl.Enabled = True Then
        ThisControl.SetFocus
      End If
      Check4Changes = True
      Exit Function
    Case Else
  End Select

  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxRateTables", "Check4Changes", Erl)
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

End Function

Private Function CheckReqFields() As Boolean
  Dim x As Integer
  
  On Error GoTo ERRORSTUFF
  
  CheckReqFields = True
  
  If QPTrim$(fptxtDescription.Text) = "" Then
    Call TaxMsg(900, "Please enter a tax rate description.")
    fptxtDescription.SetFocus
    CheckReqFields = False
    Exit Function
  End If
  
  Exit Function

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxRateTables", "CheckReqFields", Erl)
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

End Function

Private Function Check4DupDescs() As Boolean
  Dim TblRec As OptRevRateTablesType
  Dim TRHandle As Integer
  Dim NumOfTRRecs As Integer
  Dim x As Integer
  Dim ThisDesc$
  
  On Error GoTo ERRORSTUFF
  
  Check4DupDescs = False
  ThisDesc = QPTrim$(fptxtDescription.Text)
  OpenTaxRateTables TRHandle, NumOfTRRecs
  For x = 1 To NumOfTRRecs
    Get TRHandle, x, TblRec
    If x = RateTblRec Then GoTo SkipIt
    If TblRec.RevType = "R" Then GoTo SkipIt
    If TblRec.Deleted = True Then GoTo SkipIt
    If QPTrim$(TblRec.Desc) = ThisDesc Then
      Exit For
    End If
SkipIt:
  Next x
  If x <= NumOfTRRecs Then
    Call TaxMsg(900, "The description entered has already been used. Please make the description unique.")
    fptxtDescription.SetFocus
    Check4DupDescs = True
  End If
  
  Exit Function

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxRateTables", "Check4DupDescs", Erl)
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
    
End Function

Private Sub EnableOpts()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TblRec As OptRevRateTablesType
  Dim TRHandle As Integer
  Dim NumOfTRRecs As Integer
  Dim x As Integer
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  OpenTaxRateTables TRHandle, NumOfTRRecs
  For x = 1 To NumOfTRRecs
    Get TRHandle, x, TblRec
      If TblRec.Deleted = True Or TblRec.RevType = "R" Then GoTo SkipIt
      If QPTrim$(TaxMasterRec.POptRev1) = QPTrim$(TblRec.Desc) Then
        OptRev1.Enabled = False
      ElseIf QPTrim$(TaxMasterRec.POptRev2) = QPTrim$(TblRec.Desc) Then
        OptRev2.Enabled = False
      ElseIf QPTrim$(TaxMasterRec.POptRev3) = QPTrim$(TblRec.Desc) Then
        OptRev3.Enabled = False
      End If
SkipIt:
  Next x
  Close
End Sub

