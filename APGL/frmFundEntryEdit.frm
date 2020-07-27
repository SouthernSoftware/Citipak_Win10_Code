VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmFundEntryEdit 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fund Entry/Edit"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12225
   ClipControls    =   0   'False
   Icon            =   "frmFundEntryEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin EditLib.fpText txtFundNum 
      Height          =   420
      Left            =   6240
      TabIndex        =   0
      Top             =   3336
      Width           =   900
      _Version        =   196608
      _ExtentX        =   1587
      _ExtentY        =   741
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
      BackColor       =   16777215
      ForeColor       =   0
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   8421504
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   8421504
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
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
      ThreeDTextHighlightColor=   14737632
      ThreeDTextShadowColor=   8421504
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   16777215
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   16777215
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "0123456789"
      MaxLength       =   4
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   12632256
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   0
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   8421504
      BorderDropShadowWidth=   3
      ButtonColor     =   12632256
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText txtTitle 
      Height          =   372
      Left            =   6240
      TabIndex        =   1
      Top             =   4080
      Width           =   3372
      _Version        =   196608
      _ExtentX        =   5948
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
      BackColor       =   16777215
      ForeColor       =   0
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   8421504
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
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
      ThreeDTextHighlightColor=   14737632
      ThreeDTextShadowColor=   8421504
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   1
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   16777215
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   16777215
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   30
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   12632256
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   4210752
      BorderDropShadowWidth=   3
      ButtonColor     =   12632256
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F3 &Delete"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   8064
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7488
      Width           =   1356
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D0D0D0&
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
      Height          =   492
      Left            =   9600
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7488
      Width           =   1356
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 &Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   6528
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7488
      Width           =   1356
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   8385
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "5/14/2018"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "1:54 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblEditFund 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Edit Existing Fund"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5280
      TabIndex        =   10
      Top             =   5280
      Visible         =   0   'False
      Width           =   2052
   End
   Begin VB.Label lblNewFund 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "New Fund"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5280
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      Height          =   2052
      Left            =   2400
      Top             =   2880
      Width           =   7452
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Fund Name / Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   2760
      TabIndex        =   7
      Top             =   4080
      Width           =   3372
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fund Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   4080
      TabIndex        =   6
      Top             =   3360
      Width           =   1812
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fund Entry/Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4320
      TabIndex        =   5
      Top             =   1320
      Width           =   3612
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3240
      Top             =   1080
      Width           =   5772
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      Height          =   972
      Left            =   3240
      Top             =   960
      Width           =   5772
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuPrnScrn 
         Caption         =   "&Print Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmFundEntryEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim GLFund As GLFundRecType
Dim GLAcct As GLAcctRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Public RecordNum As Integer
Private Temp_Class As Resize_Class
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      ClearInUse PWcnt
    End If
  End If
End Sub


Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
Private Sub cmdDelete_Click()
'Only allow delete if not in use by accounts or saved previously
  If RecordNum > 0 Then
    If FindAcct(txtFundNum, GLFundLen) > 0 Then
      MsgBox "This Fund May Not Be Deleted", vbOKOnly, "Deletion Denied"
      Exit Sub
    Else
      If MsgBox("Are You Sure You Wish to Delete This Fund, OK to Delete, Cancel to Abort Deletion.", vbOKCancel, "Delete Fund") = vbOK Then
        DeleteFund
        txtFundNum = ""
        txtTitle = ""
        txtFundNum.SetFocus
        lblEditFund.Visible = False
      Else
        Exit Sub
      End If
    End If
  Else
    MsgBox "This Fund Has Not Been Saved and Does Not Need To Be Deleted", vbOKOnly, "Deletion Denied"
  End If
End Sub
Private Sub cmdExit_Click()
  frmFundMaintMenu.Show
  Unload frmFundEntryEdit
End Sub
Private Sub cmdSave_Click()
'Do not save if blank fields
  If txtFundNum = "" Or txtTitle = "" Then
    MsgBox "A Blank Field May Not Be Saved.", vbOKOnly
  Else
    Call SaveFund
    lblNewFund.Visible = False
  End If
'set focus back to fund number after save or after message
  txtFundNum.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      KeyCode = 0
    Case vbKeyF3:
      SendKeys "%D"
      KeyCode = 0
    
    Case Else:
  End Select
End Sub
Private Function FundSearch()
  Dim FoundFund As Boolean
  FoundFund = False 'assume we can't find it
  If Len(txtFundNum) <> GLFundLen Then
    MsgBox "Invalid Fund Code.", vbOKOnly, "Invalid Data!"
    FoundFund = False
  Else
    RecordNum = FindFund(txtFundNum)
    If RecordNum > 0 Then
      FoundFund = True
      GetFund RecordNum
      lblEditFund.Visible = True
      lblNewFund.Visible = False
      cmdDelete.Enabled = True
    Else
      FoundFund = True
      lblNewFund.Visible = True
      cmdDelete.Enabled = False
      lblEditFund.Visible = False
      txtTitle = ""
    End If
  End If
  FundSearch = FoundFund
End Function
Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  Me.HelpContextID = hlpAddChangeDelete
End Sub
Private Sub GetFund(RecordNum As Integer)
  Dim FundFile As Integer, NumFunds As Integer
  OpenFundFile FundFile, NumFunds
  Get FundFile, RecordNum, GLFund
  txtFundNum = GLFund.FundNum
  txtTitle = Trim(GLFund.Title)
  Close FundFile
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScrn_Click()
''  Set Picture1.Picture = capturescreen()
''  PrintPictureToFitPage Printer, Picture1.Picture
''  Printer.EndDoc
''      ' Clear out the picture box.
''  Set Picture1.Picture = Nothing
  PrintForm
End Sub

Private Sub txtFundNum_LostFocus()
  If Len(Trim(txtFundNum)) > 0 Then
    If FundSearch = False Then
      txtFundNum = ""
      txtTitle = ""
      lblNewFund.Visible = False
      lblEditFund.Visible = False
      cmdDelete.Enabled = False
      txtFundNum.SetFocus
    Else
      txtTitle.SetFocus
    End If
  Else
    txtTitle = ""
    lblNewFund.Visible = False
    lblEditFund.Visible = False
    cmdDelete.Enabled = False
  End If
End Sub
Private Sub SaveFund()
  Dim FundFile As Integer
  Dim NumFunds As Integer
  GLFund.Deleted = 0
  GLFund.FundNum = txtFundNum
  GLFund.Title = Trim(frmFundEntryEdit.txtTitle)
  OpenFundFile FundFile, NumFunds
  If RecordNum = 0 Then
    RecordNum = (LOF(FundFile) / Len(GLFund)) + 1
  End If
  Put FundFile, RecordNum, GLFund
  Close FundFile
  Call MainLog("Fund: " + txtFundNum + " Saved.")
  SortFundIndex
  Call MainLog("Funds sorted via Enter/Edit.")
  MsgBox "Your Information has been saved.", vbOKOnly
  txtFundNum = ""
  txtTitle = ""
  lblNewFund.Visible = False
  lblEditFund.Visible = False
  cmdDelete.Enabled = False
  RecordNum = 0
  txtFundNum.SetFocus
End Sub
Private Sub DeleteFund()
  Dim FundFile As Integer
  Dim NumFunds As Integer
  OpenFundFile FundFile, NumFunds
  GLFund.Deleted = -1
  Put FundFile, RecordNum, GLFund
  Close FundFile
  Call MainLog("Fund: " + txtFundNum + " Deleted.")
  SortFundIndex
  Call MainLog("Funds sorted via Enter/Edit.")
  GLFund.Deleted = 0
End Sub
