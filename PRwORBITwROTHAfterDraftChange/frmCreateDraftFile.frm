VERSION 5.00
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmCreateDraftFile 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Draft File"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmCreateDraftFile.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   36263.11
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   5250
      Left            =   2340
      TabIndex        =   0
      Top             =   1800
      Width           =   6975
      _Version        =   196609
      _ExtentX        =   12298
      _ExtentY        =   9250
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   192
      Caption         =   ""
      FrameColor      =   192
      FrameThreeDHighlightColor=   8454143
      FrameThreeDShadowColor=   8454143
      FrameThreeDStyle=   2
      FrameThreeDWidth=   4
      FrameWidth      =   8
      Picture         =   "frmCreateDraftFile.frx":08CA
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00D0D0D0&
         Caption         =   "F5 &OK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   444
         Left            =   2736
         TabIndex        =   8
         Top             =   4380
         Width           =   1836
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00D0D0D0&
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
         Height          =   444
         Left            =   1104
         TabIndex        =   7
         Top             =   2310
         Width           =   1836
      End
      Begin VB.CommandButton cmdProcess 
         BackColor       =   &H00D0D0D0&
         Caption         =   "F10 &Process"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   444
         Left            =   4272
         TabIndex        =   2
         Top             =   2310
         Width           =   1836
      End
      Begin EditLib.fpDateTime fptxtDraftDate 
         Height          =   375
         Left            =   3450
         TabIndex        =   1
         Top             =   915
         Width           =   1695
         _Version        =   196608
         _ExtentX        =   2984
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
         Text            =   "10-01-2001"
         DateCalcMethod  =   0
         DateTimeFormat  =   5
         UserDefinedFormat=   "mm-dd-yyyy"
         DateMax         =   "20350101"
         DateMin         =   "19800101"
         TimeMax         =   "000000"
         TimeMin         =   "000000"
         TimeString1159  =   ""
         TimeString2359  =   ""
         DateDefault     =   "19800101"
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
      Begin EditLib.fpText fptxtFileName 
         Height          =   375
         Left            =   3510
         TabIndex        =   9
         Top             =   3465
         Width           =   1740
         _Version        =   196608
         _ExtentX        =   3069
         _ExtentY        =   667
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
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Press ""F5"" to Exit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   390
         Left            =   1965
         TabIndex        =   12
         Top             =   3900
         Width           =   3270
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "File Name is:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   390
         Left            =   1590
         TabIndex        =   11
         Top             =   3510
         Width           =   1935
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Processing Completed!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   495
         Left            =   870
         TabIndex        =   10
         Top             =   2895
         Width           =   5295
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Draft Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   390
         Left            =   1635
         TabIndex        =   6
         Top             =   975
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ready to Create Draft File?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   345
         Left            =   840
         TabIndex        =   5
         Top             =   450
         Width           =   5295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Press ""ESC"" to Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   390
         Left            =   1830
         TabIndex        =   4
         Top             =   1455
         Width           =   3270
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Press ""F10"" to Process"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   390
         Left            =   1875
         TabIndex        =   3
         Top             =   1830
         Width           =   3225
      End
   End
End
Attribute VB_Name = "frmCreateDraftFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private PRDrftCode As String
Private PRLineOnenumber As String
Private PRLineFivenumber As String
Private PRBankAcctNum As String
Private PRRouteNum As String
Private Sub cmdCancel_Click()
  frmACHBankDraftMenu.Show
  DoEvents
  Unload frmCreateDraftFile
End Sub

Private Sub cmdOk_Click()
  frmACHBankDraftMenu.Show
  DoEvents
  Unload frmCreateDraftFile
  MainLog ("Create Draft File exited.")
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
      SendKeys "%C"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%O"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  fptxtDraftDate = Date
  Label7.Visible = False
  Label4.Visible = False
  cmdOK.Visible = False
  fptxtFileName.Text = "Not Processed"
  If Exist("PRDrftCode.dat") Then
    OpenDrftCode
  Else
    PRDrftCode$ = "200"
    PRLineOnenumber$ = ""
    PRLineFivenumber$ = ""
    PRBankAcctNum = ""
    PRRouteNum$ = ""
  End If
End Sub
Private Sub OpenDrftCode()
  Dim Tempfile As Integer, lentemp As Integer
  Dim CodeTemp As CitiPRDraftCodeType
  Tempfile = FreeFile
  Open "PRDrftCode.dat" For Random Shared As Tempfile Len = Len(CodeTemp) ' Len = lentemp
  Get Tempfile, 1, CodeTemp
  PRDrftCode$ = QPTrim(CodeTemp.DraftCode)
  If Val(PRDrftCode$) = 0 Then PRDrftCode$ = "200"
  PRLineOnenumber$ = CodeTemp.Line1number
  PRLineFivenumber$ = CodeTemp.Line5number
  PRBankAcctNum$ = CodeTemp.BankAcctNum
  PRRouteNum$ = CodeTemp.RountNum
  If Val(PRLineOnenumber$) = 0 Then PRLineOnenumber$ = ""
  If Val(PRLineFivenumber$) = 0 Then PRLineFivenumber$ = ""
  If Val(PRRouteNum$) > 0 Then QPTrim (PRRouteNum$)
  If Val(PRBankAcctNum$) > 0 Then QPTrim (PRBankAcctNum$)
  Close
End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub cmdProcess_Click()
  Dim PPDFLen As Integer
  Dim PPDFFile As Integer
  Dim NumOfRecs As Long
  Dim TDraftDate$
  Dim Employer$
  Dim Unit As UnitFileRecType
  Dim UHandle As Integer
  Dim NetPay#, cnt&
  Dim NetAmt$, NetTotal#
  Dim AcctNumber$, nme$
  Dim BankAcct$, hashD#
  Dim Number As Integer, hash$
  Dim TotalAmount$
  Dim DraftDate$
  Dim DraftFile As Integer
  Dim NumOfLines As Integer
  Dim TotSize#, BlockSize!, BlockSizeS$
  Dim FillSize!, outfile As Integer
  Dim TotalAmounts$
  Dim Counter As Integer
  Dim EmpFile As Integer
  Dim Emp2Len As Integer
  Dim DraftFileNum As Integer
  Dim TraceS$
  Dim Trace As Integer
  Dim Today$
  Dim blnPlymouth As Boolean
  Dim sp As String, Trac As Integer
  
  OpenUnitFile UHandle
  Get UHandle, 1, Unit
  Close UHandle
  
  ReDim PPDFInfo(1) As PRPPDraftInfoType
  PPDFLen = Len(PPDFInfo(1))

  OpenPPDraftInfo PPDFFile
  NumOfRecs& = LOF(PPDFFile) \ PPDFLen
  Get PPDFFile, 1, PPDFInfo(1)
  Close

  TDraftDate$ = MakeRegDate(PPDFInfo(1).DraftDate)
  Today$ = Right$(Date$, 2) + Left$(Date$, 2) + Mid$(Date$, 4, 2)
  Employer$ = QPTrim$(Unit.UFEMPR)

ProcessDraft:
  DraftDate$ = fptxtDraftDate.Text
  DraftDate$ = Left$(DraftDate$, 2) + Mid$(DraftDate$, 4, 2) + Right$(DraftDate$, 2)
  If Val(DraftDate$) = 0 Then Exit Sub

  OpenPPDraftInfo PPDFFile
  NumOfRecs& = LOF(PPDFFile) \ Len(PPDFInfo(1))
  

  ' Process Record Type 1
  GoSub DraftProcessStep1
  GoSub DraftProcessStep2
  GoSub DraftProcessStep3
  GoSub DraftProcessStep4
  GoSub DraftProcessStep5

  Close
  
  Label7.Visible = True
  Label4.Visible = True
  cmdOK.Visible = True
  fptxtFileName.Text = "DD" + DraftDate$ + ".ACH"
  MainLog ("Create Draft File processed.")
  
  Exit Sub
  
OpenMainDraftInformation:
  ReDim PRDraftRec(1) As PRDraftRecType
  OpenPRDraftFile DraftFile
  Get DraftFile, 1, PRDraftRec(1)
  Close DraftFile
  If Val(PRRouteNum$) = 0 Then PRRouteNum$ = QPTrim$(PRDraftRec(1).BANKORIG)
Return

DraftProcessStep1:
  GoSub OpenMainDraftInformation
  ReDim PRDraftRecord1(1) As PRDraftRecord1Type
  DraftFileNum = FreeFile
  Open Draft1FileName For Random Shared As #DraftFileNum Len = Len(PRDraftRecord1(1))
  PRDraftRecord1(1).Field1 = "1"
  PRDraftRecord1(1).Field2 = "01"
  PRDraftRecord1(1).Field3 = " " + QPTrim$(PRDraftRec(1).BANKDEST)
  If Len(PRLineOnenumber$) > 0 Then
    PRDraftRecord1(1).Field4 = PRLineOnenumber$
  Else
    PRDraftRecord1(1).Field4 = " " + QPTrim$(PRDraftRec(1).BANKORIG)
  End If
  PRDraftRecord1(1).Field5 = Today$
  PRDraftRecord1(1).Field6 = Left$(Time$, 2) + Mid$(Time$, 4, 2)
  PRDraftRecord1(1).Field7 = "A"
  PRDraftRecord1(1).Field8 = "094"
  PRDraftRecord1(1).Field9 = "10"
  PRDraftRecord1(1).Field10 = "1"
  PRDraftRecord1(1).Field11 = QPTrim$(UCase$(PRDraftRec(1).BankName))
  PRDraftRecord1(1).Field12 = QPTrim$(UCase$(PRDraftRec(1).BANKLOC))
  PRDraftRecord1(1).Field13 = "        "        'Must = 8 Spaces
  Put DraftFileNum, 1, PRDraftRecord1(1)
  Close DraftFileNum
Return

DraftProcessStep2:
  ReDim PRDraftRecord5(1) As PRDraftRecord5Type
  DraftFileNum = FreeFile
  Open Draft5FileName For Random Shared As #DraftFileNum Len = Len(PRDraftRecord5(1))
  PRDraftRecord5(1).Field1 = "5"
  PRDraftRecord5(1).Field2 = PRDrftCode$
  PRDraftRecord5(1).Field3 = Left$(Employer$, 16)
  PRDraftRecord5(1).Field4 = "                    "
  If Len(PRLineFivenumber$) > 0 Then
    PRDraftRecord5(1).Field5 = PRLineFivenumber$
  Else
    PRDraftRecord5(1).Field5 = PRDraftRec(1).FEDPREFX + QPTrim$(PRDraftRec(1).FEDID)
  End If
  PRDraftRecord5(1).Field6 = "PPD"
  PRDraftRecord5(1).Field7 = "PAYROLL  "
  PRDraftRecord5(1).Field8 = Right$(DraftDate$, 2) + Left$(DraftDate$, 2) + Mid$(DraftDate$, 3, 2)
  PRDraftRecord5(1).Field9 = Right$(DraftDate$, 2) + Left$(DraftDate$, 2) + Mid$(DraftDate$, 3, 2)
  PRDraftRecord5(1).Field10 = "   "             'Reserved w/3 blanks
  PRDraftRecord5(1).Field11 = "1"
  PRDraftRecord5(1).Field12 = Left$(PRDraftRec(1).BANKORIG, 8)
  PRDraftRecord5(1).Field13 = "0000001"
  Put DraftFileNum, 1, PRDraftRecord5(1)
  Close DraftFileNum
Return

DraftProcessStep3:
  Counter = 0
  Emp2Len = Len(Emp2Rec(1))

  OpenEmpData2File EmpFile
  ReDim PRDraftRecord6(1) As PRDraftRecord6Type
  KillFile Draft6FileName
  DraftFileNum = FreeFile
  Open Draft6FileName For Random Shared As #DraftFileNum Len = Len(PRDraftRecord6(1))
  
  For cnt& = 1 To NumOfRecs&

    Get PPDFFile, cnt&, PPDFInfo(1)
    If PPDFInfo(1).EmpRec = 0 Then GoTo SkipIt
    Get EmpFile, PPDFInfo(1).EmpRec, Emp2Rec(1)

    NetPay# = OldRound#(PPDFInfo(1).NetPay)

    If NetPay# > 0 Then ' Process Employee's Here
      NetAmt$ = QPTrim$(Str$(NetPay# * 100))    'Remove the Decimals
      If Len(NetAmt$) < 10 Then NetAmt$ = String$(10 - Len(NetAmt$), "0") + NetAmt$
      'Get Running Net Amount
      NetTotal# = OldRound#(NetTotal# + (NetPay# * 100))

      Counter = Counter + 1

      AcctNumber$ = QPTrim$(Str$(cnt&))
      If Len(AcctNumber$) < 15 Then
        AcctNumber$ = AcctNumber$ + String$(15 - Len(AcctNumber$), 32)
      End If

      nme$ = QPTrim$(Emp2Rec(1).EmpFName) + " " + QPTrim$(Emp2Rec(1).EmpLName)
      If Len(nme$) < 22 Then
        nme$ = nme$ + String$(22 - Len(nme$), 32)
      Else
        nme$ = Left$(nme$, 22)
      End If

      BankAcct$ = QPTrim$(Emp2Rec(1).EMPDDACC)

      If Len(BankAcct$) < 17 Then BankAcct$ = BankAcct$ + String$(17 - Len(BankAcct$), 32)
      Trace = Trace + 1
      TraceS = QPTrim$(Str$(Trace))
      If Len(TraceS) < 7 Then
        TraceS = String$(7 - Len(TraceS), "0") + TraceS
      End If

      PRDraftRecord6(1).Field1 = "6"
      If Emp2Rec(1).DRAFTCOD = "C" Then
        PRDraftRecord6(1).Field2 = "22"           'Designates Credit Checking
      ElseIf Emp2Rec(1).DRAFTCOD = "S" Then
        PRDraftRecord6(1).Field2 = "32"           'Designates Credit Checking
      End If

      PRDraftRecord6(1).Field3 = Left$(Emp2Rec(1).TRANSIT, 8)
      PRDraftRecord6(1).Field4 = Right$(Emp2Rec(1).TRANSIT, 1)
      PRDraftRecord6(1).Field5 = Left$(BankAcct$, 17)
      PRDraftRecord6(1).Field6 = NetAmt$        'The amt to credit
      PRDraftRecord6(1).Field7 = AcctNumber$
      PRDraftRecord6(1).Field8 = UCase$(nme$)
      PRDraftRecord6(1).Field9 = "  "
      PRDraftRecord6(1).Field10 = "0"
      PRDraftRecord6(1).Field11 = Left$(PRDraftRec(1).BANKORIG, 8) + TraceS
      Put DraftFileNum, Counter, PRDraftRecord6(1)
      hashD# = hashD# + Val(Left$(Emp2Rec(1).TRANSIT, 8))
      Number = Number + 1
    End If
SkipIt:
  Next

 If Val(PRBankAcctNum$) > 0 Then
   
    TotalAmount$ = QPTrim$(Str$(NetTotal#))     'Remove the Decimals
    TotalAmount$ = LTrim$(TotalAmount$)
    If Len(TotalAmount$) < 10 Then TotalAmount$ = String$(10 - Len(TotalAmount$), "0") + TotalAmount$
    'Get Running Billed Amount
    Counter = Counter + 1

    AcctNumber$ = "0"   'STR$(Cnt&)
    AcctNumber$ = Right$(AcctNumber$, Len(AcctNumber$) - 1)
    If Len(AcctNumber$) < 15 Then
      AcctNumber$ = AcctNumber$ + String$(15 - Len(AcctNumber$), 32)
    End If
    nme$ = Employer$
    If Len(nme$) < 22 Then
      nme$ = nme$ + String$(22 - Len(nme$), 32)
    Else
      nme$ = Left$(nme$, 22)
    End If

    sp = InStr(PRBankAcctNum$, " ")
    If sp > 0 Then
      PRBankAcctNum$ = Left$(PRBankAcctNum$, sp - 1) + Right$(PRBankAcctNum$, Len(PRBankAcctNum$) - sp)
    End If

    If Len(PRBankAcctNum$) < 17 Then
      PRBankAcctNum$ = PRBankAcctNum$ + String$(17 - Len(PRBankAcctNum$), 32)
    End If
      Trace = Trace + 1
      TraceS = QPTrim$(Str$(Trace))
      If Len(TraceS) < 7 Then
        TraceS = String$(7 - Len(TraceS), "0") + TraceS
      End If
 
    PRDraftRecord6(1).Field1 = "6"
    PRDraftRecord6(1).Field2 = "27"       ' Designates credit checking
    PRDraftRecord6(1).Field3 = Left$(PRRouteNum$, 8)
    PRDraftRecord6(1).Field4 = Right$(PRRouteNum$, 1)
    PRDraftRecord6(1).Field5 = Left$(PRBankAcctNum$, 17)
    PRDraftRecord6(1).Field6 = TotalAmount$   ' All zero's for Prenote
    PRDraftRecord6(1).Field7 = AcctNumber$
    PRDraftRecord6(1).Field8 = UCase$(nme$)
    PRDraftRecord6(1).Field9 = "  "
    PRDraftRecord6(1).Field10 = "0"
    PRDraftRecord6(1).Field11 = Left$(PRDraftRec(1).BANKORIG, 8) + TraceS$

    Put DraftFileNum, Counter, PRDraftRecord6(1)
    hashD# = hashD# + Val(Left$(PRRouteNum$, 8))
    Number = Number + 1
  End If

  Close DraftFileNum
Return

DraftProcessStep4:
  hash$ = QPTrim$(CStr(hashD#))

  If Len(hash$) < 10 Then
    hash$ = String$(10 - Len(hash$), "0") + hash$
  End If
  If Len(hash$) > 10 Then
    hash$ = Right$(hash$, 10)
  End If
  
  If Len(TraceS) > 6 Then TraceS = Right$(TraceS, 6)
  TotalAmount$ = QPTrim$(Str$(NetTotal#))
  If Len(TotalAmount$) < 12 Then TotalAmount$ = String$(12 - Len(TotalAmount$), "0") + TotalAmount$

  ReDim PRDraftRecord8(1) As PRDraftRecord8Type
  DraftFileNum = FreeFile
  Open Draft8FileName For Random Shared As #DraftFileNum Len = Len(PRDraftRecord8(1))
  PRDraftRecord8(1).Field1 = "8"
  PRDraftRecord8(1).Field2 = PRDrftCode$
  PRDraftRecord8(1).Field3 = TraceS
  PRDraftRecord8(1).Field4 = hash$
  If Val(PRBankAcctNum$) > 0 Then           'Balancing entry debits and credits
    PRDraftRecord8(1).Field5 = TotalAmount$
    PRDraftRecord8(1).Field6 = TotalAmount$
  Else      'credits only
    PRDraftRecord8(1).Field5 = "000000000000"
    PRDraftRecord8(1).Field6 = TotalAmount$
  End If
  PRDraftRecord8(1).Field7 = PRDraftRec(1).FEDPREFX + PRDraftRec(1).FEDID
  PRDraftRecord8(1).Field8 = String$(19, 32)    ' Reserved
  PRDraftRecord8(1).Field9 = String$(6, 32)     ' Reserved for Federal Reserve Use
  PRDraftRecord8(1).Field10 = Left$(PRDraftRec(1).BANKORIG, 8)
  PRDraftRecord8(1).Field11 = "0000001"
  Put DraftFileNum, 1, PRDraftRecord8(1)
  Close DraftFileNum
Return

DraftProcessStep5:
  
  TotSize# = Val(TraceS) + 4    ' Total Records= Trace + 4 control records
  TotSize# = TotSize# * 94      ' Total Bytes = 94 per record
  BlockSize! = TotSize# / 940   ' Rem Blocks Consist of Batchs of 10 Records

  If BlockSize! <> Int(BlockSize!) Then
    BlockSize! = Int(BlockSize!) + 1
    FillSize! = 940 - (TotSize# - (940 * (BlockSize! - 1)))
  Else
    FillSize! = 0
  End If

  BlockSizeS = QPTrim$(Str$(BlockSize!))

  If Len(BlockSizeS) < 6 Then BlockSizeS = String$(6 - Len(BlockSizeS), "0") + BlockSizeS
  If Len(TraceS) < 8 Then TraceS = String$(8 - Len(TraceS), "0") + TraceS

  ReDim PRDraftRecord9(1) As PRDraftRecord9Type
  DraftFileNum = FreeFile
  Open Draft9FileName For Random Shared As #DraftFileNum Len = Len(PRDraftRecord9(1))
  PRDraftRecord9(1).Field1 = "9"
  PRDraftRecord9(1).Field2 = "000001"           ' only 1 batch
  PRDraftRecord9(1).Field3 = BlockSizeS
  PRDraftRecord9(1).Field4 = TraceS
  PRDraftRecord9(1).Field5 = hash$
  If Val(PRBankAcctNum$) > 0 Then          'Balancing entry debits and credits
    PRDraftRecord9(1).Field6 = TotalAmount$
    PRDraftRecord9(1).Field7 = TotalAmount$
  Else        'credits only
    PRDraftRecord9(1).Field6 = "000000000000"
    PRDraftRecord9(1).Field7 = TotalAmount$
  End If
  PRDraftRecord9(1).Field8 = String$(39, 32)    ' Reserved
  Put DraftFileNum, 1, PRDraftRecord9(1)
  Close DraftFileNum

  ' Now Put Them Together In File Name UBDFNOTE
  outfile = FreeFile
  Open "DD" + DraftDate$ + ".ACH" For Output Shared As outfile 'len = 255

  ReDim PRDraftRecord1(1) As PRDraftRecord1Type
  DraftFileNum = FreeFile
  Open Draft1FileName For Random Shared As #DraftFileNum Len = Len(PRDraftRecord1(1))
  Get DraftFileNum, 1, PRDraftRecord1(1)
  Print #outfile, PRDraftRecord1(1).Field1;
  Print #outfile, PRDraftRecord1(1).Field2;
  Print #outfile, PRDraftRecord1(1).Field3;
  Print #outfile, PRDraftRecord1(1).Field4;
  Print #outfile, PRDraftRecord1(1).Field5;
  Print #outfile, PRDraftRecord1(1).Field6;
  Print #outfile, PRDraftRecord1(1).Field7;
  Print #outfile, PRDraftRecord1(1).Field8;
  Print #outfile, PRDraftRecord1(1).Field9;
  Print #outfile, PRDraftRecord1(1).Field10;
  Print #outfile, PRDraftRecord1(1).Field11;
  Print #outfile, PRDraftRecord1(1).Field12;
  Print #outfile, PRDraftRecord1(1).Field13
  Close DraftFileNum

  ReDim PRDraftRecord5(1) As PRDraftRecord5Type
  DraftFileNum = FreeFile
  Open Draft5FileName For Random Shared As #DraftFileNum Len = Len(PRDraftRecord5(1))
  Get DraftFileNum, 1, PRDraftRecord5(1)
  Print #outfile, PRDraftRecord5(1).Field1;
  Print #outfile, PRDraftRecord5(1).Field2;
  Print #outfile, PRDraftRecord5(1).Field3;
  Print #outfile, PRDraftRecord5(1).Field4;
  Print #outfile, PRDraftRecord5(1).Field5;
  Print #outfile, PRDraftRecord5(1).Field6;
  Print #outfile, PRDraftRecord5(1).Field7;
  Print #outfile, PRDraftRecord5(1).Field8;
  Print #outfile, PRDraftRecord5(1).Field9;
  Print #outfile, PRDraftRecord5(1).Field10;
  Print #outfile, PRDraftRecord5(1).Field11;
  Print #outfile, PRDraftRecord5(1).Field12;
  Print #outfile, PRDraftRecord5(1).Field13
  Close DraftFileNum

  ReDim PRDraftRecord6(1) As PRDraftRecord6Type
  DraftFileNum = FreeFile
  Open Draft6FileName For Random Shared As #DraftFileNum Len = Len(PRDraftRecord6(1))
  NumOfLines = LOF(DraftFileNum) / 94

  For cnt = 1 To NumOfLines
    Get DraftFileNum, cnt, PRDraftRecord6(1)
    Print #outfile, PRDraftRecord6(1).Field1;
    Print #outfile, PRDraftRecord6(1).Field2;
    Print #outfile, PRDraftRecord6(1).Field3;
    Print #outfile, PRDraftRecord6(1).Field4;
    Print #outfile, PRDraftRecord6(1).Field5;
    Print #outfile, PRDraftRecord6(1).Field6;
    Print #outfile, PRDraftRecord6(1).Field7;
    Print #outfile, PRDraftRecord6(1).Field8;
    Print #outfile, PRDraftRecord6(1).Field9;
    Print #outfile, PRDraftRecord6(1).Field10;
    Print #outfile, PRDraftRecord6(1).Field11
  Next cnt
  Close DraftFileNum

  ReDim PRDraftRecord8(1) As PRDraftRecord8Type
  DraftFileNum = FreeFile
  Open Draft8FileName For Random Shared As #DraftFileNum Len = Len(PRDraftRecord8(1))
  Get DraftFileNum, 1, PRDraftRecord8(1)
  Print #outfile, PRDraftRecord8(1).Field1;
  Print #outfile, PRDraftRecord8(1).Field2;
  Print #outfile, PRDraftRecord8(1).Field3;
  Print #outfile, PRDraftRecord8(1).Field4;
  Print #outfile, PRDraftRecord8(1).Field5;
  Print #outfile, PRDraftRecord8(1).Field6;
  Print #outfile, PRDraftRecord8(1).Field7;
  Print #outfile, PRDraftRecord8(1).Field8;
  Print #outfile, PRDraftRecord8(1).Field9;
  Print #outfile, PRDraftRecord8(1).Field10;
  Print #outfile, PRDraftRecord8(1).Field11
  Close DraftFileNum

  ReDim PRDraftRecord9(1) As PRDraftRecord9Type
  DraftFileNum = FreeFile
  Open Draft9FileName For Random Shared As #DraftFileNum Len = Len(PRDraftRecord9(1))
  Get DraftFileNum, 1, PRDraftRecord9(1)
  Print #outfile, PRDraftRecord9(1).Field1;
  Print #outfile, PRDraftRecord9(1).Field2;
  Print #outfile, PRDraftRecord9(1).Field3;
  Print #outfile, PRDraftRecord9(1).Field4;
  Print #outfile, PRDraftRecord9(1).Field5;
  Print #outfile, PRDraftRecord9(1).Field6;
  Print #outfile, PRDraftRecord9(1).Field7;
  Print #outfile, PRDraftRecord9(1).Field8
  Close DraftFileNum

  For cnt = 1 To FillSize! / 94
    Print #outfile, String$(94, "9")
  Next cnt
  Close

Return

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdCancel.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmCreateDraftFile.")
      Call Terminate
      End
    End If
  End If
End Sub

