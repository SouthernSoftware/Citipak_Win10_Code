VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmFunctionEdit 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Function Entry/Edit"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12225
   Icon            =   "frmFunctionEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin EditLib.fpText txtFunctionNum 
      Height          =   420
      Left            =   6240
      TabIndex        =   0
      Top             =   3336
      Width           =   1356
      _Version        =   196608
      _ExtentX        =   2392
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
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483639
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
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
      ButtonStyle     =   0
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
      AutoCase        =   1
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483643
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
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
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
   Begin VB.Label lblEditFunction 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Edit Existing Function"
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
      Left            =   4758
      TabIndex        =   10
      Top             =   5280
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Label lblNewFunction 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "New Function"
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
      Left            =   5202
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
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Function Name / Desc."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   2616
      TabIndex        =   7
      Top             =   4080
      Width           =   3372
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Function Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   3096
      TabIndex        =   6
      Top             =   3360
      Width           =   2796
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Function Entry/Edit"
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
Attribute VB_Name = "frmFunctionEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim GLFNCT As GLFNCTRecType
Dim GLAcct As GLAcctRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Public RecordNum As Long
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
    If FindAcctFnct(RecordNum) > 0 Then
      MsgBox "This Function May Not Be Deleted", vbOKOnly, "Deletion Denied"
      Exit Sub
    Else
      If MsgBox("Are You Sure You Wish to Delete This Fund, OK to Delete, Cancel to Abort Deletion.", vbOKCancel, "Delete Fund") = vbOK Then
        DeleteFunction
        txtFunctionNum = ""
        txtTitle = ""
        txtFunctionNum.SetFocus
        lblEditFunction.Visible = False
      Else
        Exit Sub
      End If
    End If
  Else
    MsgBox "This Function Has Not Been Saved and Does Not Need To Be Deleted", vbOKOnly, "Deletion Denied"
  End If
End Sub
Private Sub cmdExit_Click()
  frmFunctionMenu.Show
  Unload frmFunctionEdit
End Sub
Private Sub cmdSave_Click()
'Do not save if blank fields
  If txtFunctionNum = "" Or txtTitle = "" Then
    MsgBox "A Blank Field May Not Be Saved.", vbOKOnly
  Else
    Call SaveFunction
    lblNewFunction.Visible = False
  End If
'set focus back to fund number after save or after message
  If txtFunctionNum.Enabled = True Then
    txtFunctionNum.SetFocus
  End If
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
Private Function FunctionSearch()
  Dim FoundFunction As Boolean
  FoundFunction = False 'assume we can't find it
  If Len(txtFunctionNum) = 0 Then
    MsgBox "Invalid Function Code.", vbOKOnly, "Invalid Data!"
    FoundFunction = False
  Else
    RecordNum = FindFnct(txtFunctionNum)
    If RecordNum > 0 Then
      FoundFunction = True
      GetFunction RecordNum
      lblEditFunction.Visible = True
      lblNewFunction.Visible = False
      cmdDelete.Enabled = True
    Else
      FoundFunction = False
      lblNewFunction.Visible = True
      cmdDelete.Enabled = False
      lblEditFunction.Visible = False
      txtTitle = ""
    End If
  End If
  FunctionSearch = FoundFunction
End Function
Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  txtFunctionNum.Enabled = True
  Me.HelpContextID = hlpFunctionEntryEdit
End Sub
Private Sub GetFunction(RecordNum As Long)
  Dim FnctFile As Integer, NumFncts As Integer
  OpenFnctFile FnctFile, NumFncts
  Get FnctFile, RecordNum, GLFNCT
  txtFunctionNum = GLFNCT.FnctNum
  txtTitle = Trim(GLFNCT.Title)
  txtFunctionNum.Enabled = False
  Close FnctFile
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

Private Sub txtFunctionNum_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Then
    txtTitle.SetFocus
  End If
End Sub

Private Sub txtFunctionNum_LostFocus()
  If Len(QPTrim(txtFunctionNum)) > 0 Then
    If FunctionSearch = False Then
      'txtFunctionNum = ""
      'txtTitle.SetFocus
      lblNewFunction.Visible = True
      lblEditFunction.Visible = False
      cmdDelete.Enabled = False
      'txtFunctionNum.SetFocus
    Else
      txtTitle.SetFocus
    End If
  Else
    txtTitle = ""
    lblNewFunction.Visible = True
    lblEditFunction.Visible = False
    cmdDelete.Enabled = False
  End If
End Sub
Private Sub SaveFunction()
  Dim FnctFile As Integer
  Dim NumFncts As Integer
  GLFNCT.Deleted = 0
  GLFNCT.FnctNum = txtFunctionNum
  GLFNCT.Title = Trim(txtTitle)
  OpenFnctFile FnctFile, NumFncts
  If RecordNum = 0 Then
    RecordNum = (LOF(FnctFile) / Len(GLFNCT)) + 1
  End If
  Put FnctFile, RecordNum, GLFNCT
  Close FnctFile
  Call MainLog("Function: " + txtFunctionNum + " Saved.")
  SortFNCTIndex
  Call MainLog("Functions sorted via Enter/Edit.")
  MsgBox "Your Information has been saved.", vbOKOnly
  txtFunctionNum = ""
  txtTitle = ""
  lblNewFunction.Visible = False
  lblEditFunction.Visible = False
  cmdDelete.Enabled = False
  RecordNum = 0
  txtFunctionNum.Enabled = True
  txtFunctionNum.SetFocus
End Sub
Private Sub DeleteFunction()
  Dim FnctFile As Integer
  Dim NumFncts As Integer
  OpenFnctFile FnctFile, NumFncts
  GLFNCT.Deleted = -1
  Put FnctFile, RecordNum, GLFNCT
  Close FnctFile
  Call MainLog("Function: " + txtFunctionNum + " Deleted.")
  SortFNCTIndex
  Call MainLog("Functions sorted via Enter/Edit.")
  GLFNCT.Deleted = 0
  txtFunctionNum.Enabled = True
End Sub
