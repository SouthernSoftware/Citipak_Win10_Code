VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "EDT32X30.OCX"
Begin VB.Form frmEarningsCodeMaint 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Additional Earnings Codes"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11652
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   36263.11
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ComboBox comboRET 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   7536
      TabIndex        =   7
      ToolTipText     =   "Is this deduction federal tax deferred?"
      Top             =   3744
      Width           =   852
   End
   Begin VB.ComboBox comboFWT 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   1344
      TabIndex        =   3
      ToolTipText     =   "Is this deduction federal tax deferred?"
      Top             =   3744
      Width           =   852
   End
   Begin VB.ComboBox comboSWT 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   2880
      TabIndex        =   4
      ToolTipText     =   "Is this deduction State tax deferred?"
      Top             =   3744
      Width           =   852
   End
   Begin VB.ComboBox comboSOC 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   4464
      TabIndex        =   5
      ToolTipText     =   "Is this deduction Social Security Tax deferred?"
      Top             =   3744
      Width           =   852
   End
   Begin VB.ComboBox comboMED 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   6000
      TabIndex        =   6
      ToolTipText     =   "Is this deduction Medicare Tax deferred?"
      Top             =   3744
      Width           =   852
   End
   Begin VB.CommandButton cmdSaveContinue 
      Caption         =   "F10 &Save and Continue"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   9072
      TabIndex        =   10
      Top             =   3120
      Width           =   1692
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Esc E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   9072
      TabIndex        =   9
      Top             =   2448
      Width           =   1692
   End
   Begin VB.CommandButton cmdSaveExit 
      Caption         =   "F11 Sa&ve and Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   9072
      TabIndex        =   11
      Top             =   3816
      Width           =   1692
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Data Entry Fields"
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
      Left            =   2448
      TabIndex        =   13
      Top             =   4752
      Width           =   3036
   End
   Begin EditLib.fpText fpDescription 
      Height          =   372
      Left            =   2640
      TabIndex        =   2
      ToolTipText     =   "Enter a description for this deduction."
      Top             =   2952
      Width           =   5532
      _Version        =   196608
      _ExtentX        =   9758
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
      ControlType     =   0
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
   Begin FPSpread.vaSpread vaSpreadDeductionCodes 
      Height          =   3192
      Left            =   2016
      TabIndex        =   8
      Top             =   5232
      Width           =   8280
      _Version        =   196613
      _ExtentX        =   14605
      _ExtentY        =   5630
      _StockProps     =   64
      AutoSize        =   -1  'True
      ButtonDrawMode  =   4
      ColsFrozen      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   6
      MaxRows         =   498
      ProcessTab      =   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "frmNewEarnCode.frx":0000
      StartingColNumber=   0
      VisibleCols     =   6
      VisibleRows     =   10
      ScrollBarTrack  =   1
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "RET"
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
      Left            =   7008
      TabIndex        =   19
      Top             =   3840
      Width           =   492
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   1092
      Index           =   1
      Left            =   1536
      Top             =   696
      Width           =   8652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Additional Earnings Codes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2760
      TabIndex        =   16
      Top             =   1056
      Width           =   6012
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Data Entry"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   2592
      TabIndex        =   0
      Top             =   2136
      Width           =   3732
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "FWT"
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
      Left            =   816
      TabIndex        =   18
      Top             =   3840
      Width           =   492
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "Description"
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
      Left            =   1008
      TabIndex        =   17
      Top             =   3072
      Width           =   1572
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1536
      Top             =   576
      Width           =   8652
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Withholding on Earnings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   5952
      TabIndex        =   1
      Top             =   4752
      Width           =   3732
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   372
      Left            =   5952
      Top             =   4752
      Width           =   3732
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2292
      Left            =   576
      Top             =   2256
      Width           =   10644
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "SWT"
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
      Left            =   2352
      TabIndex        =   15
      Top             =   3840
      Width           =   492
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "SOC"
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
      Left            =   3888
      TabIndex        =   14
      Top             =   3840
      Width           =   492
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "MED"
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
      Left            =   5424
      TabIndex        =   12
      Top             =   3840
      Width           =   492
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   26889.23
      X2              =   26889.23
      Y1              =   2244.763
      Y2              =   4419.377
   End
End
Attribute VB_Name = "frmEarningsCodeMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this was a very complicated routine...

Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim ClickFlag As Integer
Dim RowFlag As Integer
Dim PriorRowNum As Integer
Dim ContinueFlag As Integer
Dim SaveAndExitFlag As Integer
Dim BadDataFlag As Integer
Dim NoDataFlag As Integer
Dim changeFlag As Boolean
Dim PriorDesc As String
Dim PriorLibNum As String
Dim PriorFWT As String
Dim PriorSWT As String
Dim PriorSOC As String
Dim ClearFieldsFlag As Boolean
Dim PriorMED As String
Dim PriorRET As String

Private Sub cmdClear_Click()
'This sub was designed to allow the user to clear all fields
'after having been editing existing entries...if the user
'has been editing then in order to enter a brand new entry
'they would have to highlight the last empty row and then
'enter new data...this way the program knows this is a new entry
'and automatically saves new data in the next empty row
  
'before entering new data we must check to see if the last
'entry was saved properly and if not give the user the option
'to save or not to save
  PriorDesc = QPTrim$(fpDescription.Text)
  PriorFWT = QPTrim$(comboFWT.Text)
  PriorSWT = QPTrim$(comboSWT.Text)
  PriorSOC = QPTrim$(comboSOC.Text)
  PriorMED = QPTrim$(comboMED.Text)
  PriorRET = QPTrim$(comboRET.Text)
  Call CheckForChanges
  If changeFlag = True Then
    If MsgBox("Your last edit was not saved. Do you want to save it?", vbYesNo) = vbYes Then
      Call cmdSaveContinue_Click
      changeFlag = False
    End If
  End If
  fpDescription.Text = ""
  comboFWT.Text = ""
  comboSWT.Text = ""
  comboSOC.Text = ""
  comboMED.Text = ""
  comboRET.Text = ""
  ClearFieldsFlag = True
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      KeyCode = 0
    Case vbKeyF11:
      SendKeys "%v"
      KeyCode = 0
    Case Else:
  End Select

End Sub
Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  LoadUnitFile
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub cmdExit_Click()
frmControlFileMaint.Show
Unload frmDeductionCodes
End Sub

Private Sub cmdSaveContinue_Click()
   Dim A$, B$, C$, D$, E$, F$
   Dim TempRowFlag As Integer
   A = Len(QPTrim$(fpDescription.Text))
   B = Len(QPTrim$(comboRET.Text))
   C = Len(QPTrim$(comboFWT.Text))
   D = Len(QPTrim$(comboSWT.Text))
   E = Len(QPTrim$(comboSOC.Text))
   F = Len(QPTrim$(comboMED.Text))
   'If the user wants to save and then exit the screen
   'we do not turn on the ContinueFlag
   If SaveAndExitFlag = 1 Then
      ContinueFlag = 0
   Else
      ContinueFlag = 1
   End If
   Dim ErnCodeFileHandle As Integer, x As Integer, FileLen As Integer
   Dim ErnCodeFileRec As ErnCodeRecType
   Dim RowCount As Integer
   'because more than one field needs to have the focus set each
   'If statement below must handle code individually instead of sending
   'error traps to a goto line as done in the next series of If
   'statements
   If A <> 0 Then
      If B = 0 Then
         comboRET.SetFocus
         MsgBox "All fields must be filled out if the Description field is filled out."
         BadDataFlag = 1
         'reset the screen to where it was when it was last valid
         Call Form_Load
         GoTo ExitTran
      End If
      If C = 0 Then
         comboFWT.SetFocus
         MsgBox "All fields must be filled out if the Description field is filled out."
         BadDataFlag = 1
         Call Form_Load
         GoTo ExitTran
      End If
      If D = 0 Then
         comboSWT.SetFocus
         MsgBox "All fields must be filled out if the Description field is filled out."
         BadDataFlag = 1
         Call Form_Load
         GoTo ExitTran
      End If
      If E = 0 Then
         comboSOC.SetFocus
         MsgBox "All fields must be filled out if the Description field is filled out."
         BadDataFlag = 1
         Call Form_Load
         GoTo ExitTran
      End If
      If F = 0 Then
         comboMED.SetFocus
         MsgBox "All fields must be filled out if the Description field is filled out."
         BadDataFlag = 1
         Call Form_Load
         GoTo ExitTran
      End If
   End If
   'If nothing has been entered and the user tries to save
   'a message box alerts them to this because we do not want
   'to save empty fields
   If A + B + C + D + E + F = 0 Then
      'NoDataFlag is set to 1 so if the user wanted to
      'exit the screen after this save then that procedure
      'will behave with the save response
      NoDataFlag = 1
      If SaveAndExitFlag = 0 Then
         MsgBox "No new or edited data to save"
      End If
      Exit Sub
   End If
   'if the description field is empty and any other field
   'is not empty this if statement traps this error
   If A = 0 Then
      If B <> 0 Then GoTo BadDataEntry
      If C <> 0 Then GoTo BadDataEntry
      If D <> 0 Then GoTo BadDataEntry
      If E <> 0 Then GoTo BadDataEntry
      If F <> 0 Then GoTo BadDataEntry
      Else: GoTo EntryDataOK
BadDataEntry:
       MsgBox "Please complete the Description field, double click an existing account to edit or delete all fields to continue."
       BadDataFlag = 1
       Call Form_Load
       fpDescription.SetFocus
       GoTo ExitTran
   End If
EntryDataOK:
   OpenErnCodeFile ErnCodeFileHandle
   FileLen = LOF(ErnCodeFileHandle) / Len(ErnCodeFileRec)
   If FileLen = 0 Then 'first save
      vaSpreadDeductionCodes.Col = 1
      vaSpreadDeductionCodes.Row = 1
      ErnCodeFileRec.ERNCODE1 = QPTrim$(fpDescription.Text)
      vaSpreadDeductionCodes.Col = 2
      vaSpreadDeductionCodes.Row = 1
      ErnCodeFileRec.ERNFWT1 = QPTrim$(comboFWT.Text)
      vaSpreadDeductionCodes.Col = 3
      vaSpreadDeductionCodes.Row = 1
      ErnCodeFileRec.ERNSWT1 = QPTrim$(comboSWT.Text)
      vaSpreadDeductionCodes.Col = 4
      vaSpreadDeductionCodes.Row = 1
      ErnCodeFileRec.ERNSOC1 = QPTrim$(comboSOC.Text)
      vaSpreadDeductionCodes.Col = 5
      vaSpreadDeductionCodes.Row = 1
      ErnCodeFileRec.ERNMED1 = QPTrim$(comboMED.Text)
      vaSpreadDeductionCodes.Col = 6
      vaSpreadDeductionCodes.Row = 1
      ErnCodeFileRec.ERNRET1 = QPTrim$(comboRET.Text)
      Put ErnCodeFileHandle, 1, ErnCodeFileRec
      GoTo ClickSave
   End If
   'ClickFlag denotes we are here because the user double clicked a row
   'to edit it
   'If RowFlag is not revalued to the row that was just changed
   'the change takes place but in the row that is now in focus
   'causing data to be saved to the wrong row
   If changeFlag = True Then
      TempRowFlag = RowFlag 'save current row setting
      RowFlag = PriorRowNum 'reset row to the one that was changed
   End If
   'if ClearFieldsFlag is true then we do not want to save anything
   'until we find the first empty row
   If ClearFieldsFlag = True Then GoTo NonEditEntry
   If ClickFlag = 1 Then 'save row that was double clicked for edit
      Get ErnCodeFileHandle, RowFlag, ErnCodeFileRec
      vaSpreadDeductionCodes.Col = 1
      vaSpreadDeductionCodes.Row = RowFlag
      ErnCodeFileRec.ERNCODE1 = QPTrim$(fpDescription.Text)
      vaSpreadDeductionCodes.Col = 2
      vaSpreadDeductionCodes.Row = RowFlag
      ErnCodeFileRec.ERNFWT1 = QPTrim$(comboFWT.Text)
      vaSpreadDeductionCodes.Col = 3
      vaSpreadDeductionCodes.Row = RowFlag
      ErnCodeFileRec.ERNSWT1 = QPTrim$(comboSWT.Text)
      vaSpreadDeductionCodes.Col = 4
      vaSpreadDeductionCodes.Row = RowFlag
      ErnCodeFileRec.ERNSOC1 = QPTrim$(comboSOC.Text)
      vaSpreadDeductionCodes.Col = 5
      vaSpreadDeductionCodes.Row = RowFlag
      ErnCodeFileRec.ERNMED1 = QPTrim$(comboMED.Text)
      vaSpreadDeductionCodes.Col = 6
      vaSpreadDeductionCodes.Row = RowFlag
      ErnCodeFileRec.ERNRET1 = QPTrim$(comboRET.Text)
      ClickFlag = 0
      Put ErnCodeFileHandle, RowFlag, ErnCodeFileRec
      'change RowFlag back to original value
      If changeFlag = True Then
          changeFlag = False
          RowFlag = TempRowFlag
      Else
          RowFlag = 0
      End If
      GoTo ClickSave
   End If
   'save data from fields at top of form
NonEditEntry:
   For x = 1 To 498
      vaSpreadDeductionCodes.Col = 1
      vaSpreadDeductionCodes.Row = x
      If Len(QPTrim$(vaSpreadDeductionCodes.Value)) = 0 Then
      'save in the next empty row
         RowCount = x
         Exit For
      End If
   Next
   If x > 498 Then MsgBox "You have reached the maximum allowable deductions"
   
   Get ErnCodeFileHandle, RowCount, ErnCodeFileRec
   vaSpreadDeductionCodes.Col = 1
   vaSpreadDeductionCodes.Row = RowCount
   ErnCodeFileRec.ERNCODE1 = QPTrim$(fpDescription.Text)
   vaSpreadDeductionCodes.Col = 2
   vaSpreadDeductionCodes.Row = RowCount
   ErnCodeFileRec.ERNFWT1 = QPTrim$(comboFWT.Text)
   vaSpreadDeductionCodes.Col = 3
   vaSpreadDeductionCodes.Row = RowCount
   ErnCodeFileRec.ERNSWT1 = QPTrim$(comboSWT.Text)
   vaSpreadDeductionCodes.Col = 4
   vaSpreadDeductionCodes.Row = RowCount
   ErnCodeFileRec.ERNSOC1 = QPTrim$(comboSOC.Text)
   vaSpreadDeductionCodes.Col = 5
   vaSpreadDeductionCodes.Row = RowCount
   ErnCodeFileRec.ERNMED1 = QPTrim$(comboMED.Text)
   vaSpreadDeductionCodes.Col = 6
   vaSpreadDeductionCodes.Row = RowCount
   ErnCodeFileRec.ERNRET1 = QPTrim$(comboRET.Text)
   Put ErnCodeFileHandle, RowCount, ErnCodeFileRec
   Close ErnCodeFileHandle
   'Save And Exit command button so we don't need anything between here
   'and ExitTran
ClickSave:
   BadDataFlag = 0
   If SaveAndExitFlag = 1 Then GoTo ExitTran ' this save is coming from the
   'the exit and save routine that has already performed everything
   'from here to ExitTran
   MsgBox "Your Information has been saved.", vbOKOnly
   Call Form_Load
   fpDescription.SetFocus
ExitTran:

End Sub

Private Sub cmdSaveExit_Click()
   SaveAndExitFlag = 1
   Call cmdSaveContinue_Click
   If BadDataFlag = 1 Then
      Call Form_Load
      GoTo ExitTran
   End If
      If NoDataFlag = 1 Then
      MsgBox "No new or edited data to save"
      Call Form_Load
      GoTo ExitTran
   End If
   MsgBox "Your Information has been saved.", vbOKOnly
   SaveAndExitFlag = 0
   frmControlFileMaint.Show
   Unload frmDeductionCodes
ExitTran:

End Sub

Private Sub LoadUnitFile()
   'all fields in the upper block must be cleared for the
   'ClickFlag to work properly
   NoDataFlag = 0
   If BadDataFlag = 1 Then GoTo ContinueFlagOn
   fpDescription.Text = ""
   comboFWT.Text = ""
   comboSWT.Text = ""
   comboSOC.Text = ""
   comboMED.Text = ""
   comboRET.Text = ""
   Dim ErnCodeFileHandle As Integer, x As Integer, FileLen As Integer
   Dim ErnCodeFileRec As ErnCodeRecType
   If ContinueFlag = 1 Then
      ContinueFlag = 0
      GoTo ContinueFlagOn
   End If
   'load the combo boxes in the upper block..if we are reloading from
   'the Save and Continue button then we don't need to reload the combo
   'boxes because the form was never unloaded
'   If BadDataFlag = 1 Then GoTo ContinueFlagOn
   comboFWT.AddItem ("Y")
   comboFWT.AddItem ("N")
   comboSWT.AddItem ("Y")
   comboSWT.AddItem ("N")
   comboSOC.AddItem ("Y")
   comboSOC.AddItem ("N")
   comboMED.AddItem ("Y")
   comboMED.AddItem ("N")
   comboRET.AddItem ("Y")
   comboRET.AddItem ("N")
ContinueFlagOn:
   OpenErnCodeFile ErnCodeFileHandle
   FileLen = LOF(ErnCodeFileHandle) / Len(ErnCodeFileRec)
   'This for loop loads all data stored on file plus it loads "N" in
   'the FWT, SWT, SOC ,MED and RET fields if no description is on that row
   For x = 1 To FileLen
      Get ErnCodeFileHandle, x, ErnCodeFileRec
   'load form info
      vaSpreadDeductionCodes.Col = 1
      vaSpreadDeductionCodes.Row = x
      vaSpreadDeductionCodes.Text = QPTrim$(ErnCodeFileRec.ERNCODE1)
      vaSpreadDeductionCodes.Col = 2
      vaSpreadDeductionCodes.Row = x
      vaSpreadDeductionCodes.Text = QPTrim$(ErnCodeFileRec.ERNFWT1)
      vaSpreadDeductionCodes.Col = 3
      vaSpreadDeductionCodes.Row = x
      vaSpreadDeductionCodes.Text = QPTrim$(ErnCodeFileRec.ERNSWT1)
      vaSpreadDeductionCodes.Col = 4
      vaSpreadDeductionCodes.Row = x
      vaSpreadDeductionCodes.Text = QPTrim$(ErnCodeFileRec.ERNSOC1)
      vaSpreadDeductionCodes.Col = 5
      vaSpreadDeductionCodes.Text = x
      vaSpreadDeductionCodes.Value = QPTrim$(ErnCodeFileRec.ERNMED1)
      vaSpreadDeductionCodes.Col = 6
      vaSpreadDeductionCodes.Text = x
      vaSpreadDeductionCodes.Value = QPTrim$(ErnCodeFileRec.ERNRET1)
   Next
   Close ErnCodeFileHandle
End Sub

Private Sub vaSpreadDeductionCodes_DblClick(ByVal Col As Long, ByVal Row As Long)
  'save all data before the doubleclick removed them
  RowFlag = Row
  PriorDesc = QPTrim$(fpDescription.Text)
  PriorFWT = QPTrim$(comboFWT.Text)
  PriorSWT = QPTrim$(comboSWT.Text)
  PriorSOC = QPTrim$(comboSOC.Text)
  PriorMED = QPTrim$(comboMED.Text)
  PriorRET = QPTrim$(comboRET.Text)
  'if ClearFieldsFlag is true we've already checked for changes
  If ClearFieldsFlag = True Then
     ClearFieldsFlag = False
     GoTo NoChangeCheck
  End If
  If ClickFlag = 1 Then
    Call CheckForChanges
      If changeFlag = True Then
        If MsgBox("Your last edit was not saved. Do you want to save it?", vbYesNo) = vbYes Then
          Call cmdSaveContinue_Click
          changeFlag = False
        End If
      End If
   End If
'This routine allows the user to double click a specific row
'that places that row's data in the edit fields
NoChangeCheck:
   ClickFlag = 1
   'load the fields in the upper block with the data for
   'the file numbered as Row
   vaSpreadDeductionCodes.Col = 1
   vaSpreadDeductionCodes.Row = Row
   fpDescription.Text = QPTrim$(vaSpreadDeductionCodes.Value)
   vaSpreadDeductionCodes.Col = 2
   vaSpreadDeductionCodes.Row = Row
   comboFWT.Text = QPTrim$(vaSpreadDeductionCodes.Text)
   vaSpreadDeductionCodes.Col = 3
   vaSpreadDeductionCodes.Row = Row
   comboSWT.Text = QPTrim$(vaSpreadDeductionCodes.Text)
   vaSpreadDeductionCodes.Col = 4
   vaSpreadDeductionCodes.Row = Row
   comboSOC.Text = QPTrim$(vaSpreadDeductionCodes.Text)
   vaSpreadDeductionCodes.Col = 5
   vaSpreadDeductionCodes.Row = Row
   comboMED.Text = QPTrim$(vaSpreadDeductionCodes.Text)
   vaSpreadDeductionCodes.Col = 6
   vaSpreadDeductionCodes.Row = Row
   comboRET.Text = QPTrim$(vaSpreadDeductionCodes.Text)
   PriorRowNum = RowFlag

End Sub

Private Sub CheckForChanges()
'This routine compares data in the row that just lost focus with the data
'that is in the appropriate row in the spreadsheet...if a change
'has been made it will be detected here
   vaSpreadDeductionCodes.Col = 1
   vaSpreadDeductionCodes.Row = PriorRowNum
   If QPTrim$(vaSpreadDeductionCodes.Text) <> QPTrim$(PriorDesc) Then 'QPTrim$(ErnCodeFileRec.ERNCODE1) Then
     changeFlag = True
     vaSpreadDeductionCodes.SetFocus
   End If
   vaSpreadDeductionCodes.Col = 2
   vaSpreadDeductionCodes.Row = PriorRowNum
   If QPTrim$(vaSpreadDeductionCodes.Text) <> QPTrim$(PriorFWT) Then
     changeFlag = True
     vaSpreadDeductionCodes.SetFocus
   End If
      
   vaSpreadDeductionCodes.Col = 3
   vaSpreadDeductionCodes.Row = PriorRowNum
   If QPTrim$(vaSpreadDeductionCodes.Text) <> QPTrim$(PriorSWT) Then
     changeFlag = True
     vaSpreadDeductionCodes.SetFocus
   End If
   vaSpreadDeductionCodes.Col = 4
   vaSpreadDeductionCodes.Row = PriorRowNum
   If QPTrim$(vaSpreadDeductionCodes.Text) <> QPTrim$(PriorSOC) Then
     changeFlag = True
     vaSpreadDeductionCodes.SetFocus
   End If
      
   vaSpreadDeductionCodes.Col = 5
   vaSpreadDeductionCodes.Row = PriorRowNum
   If QPTrim$(vaSpreadDeductionCodes.Text) <> QPTrim$(PriorMED) Then
     changeFlag = True
     vaSpreadDeductionCodes.SetFocus
   End If

   vaSpreadDeductionCodes.Col = 6
   vaSpreadDeductionCodes.Row = PriorRowNum
   If QPTrim$(vaSpreadDeductionCodes.Text) <> QPTrim$(PriorRET) Then
     changeFlag = True
     vaSpreadDeductionCodes.SetFocus
   End If
End Sub

