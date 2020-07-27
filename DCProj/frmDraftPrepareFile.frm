VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmDraftPrepareFile 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Draft Transmission File"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   2172
   ClientWidth     =   12216
   Icon            =   "frmDraftPrepareFile.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdOk 
      Caption         =   "F10 &Ok"
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
      Left            =   7776
      TabIndex        =   2
      Top             =   7368
      Width           =   1332
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
      Height          =   492
      Left            =   9456
      TabIndex        =   3
      Top             =   7368
      Width           =   1332
   End
   Begin EditLib.fpText fptxtCycleSel 
      Height          =   348
      Left            =   5904
      TabIndex        =   1
      Top             =   3312
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
      ThreeDOutsideStyle=   2
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
      InvalidColor    =   -2147483637
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
      MaxLength       =   2
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
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
            TextSave        =   "5:23 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "3/15/2005"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin EditLib.fpText fptxtcycle 
      Height          =   372
      Left            =   3000
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5568
      Width           =   6252
      _Version        =   196608
      _ExtentX        =   11028
      _ExtentY        =   656
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
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
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
   Begin EditLib.fpDateTime txtDate1 
      Height          =   348
      Left            =   5904
      TabIndex        =   0
      Top             =   2808
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
      _ExtentY        =   614
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
      Text            =   "07/09/2004"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
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
   Begin VB.Label LabelB1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Cycle:"
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
      Height          =   372
      Left            =   3840
      TabIndex        =   11
      Top             =   3360
      Width           =   1932
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3228
      Top             =   960
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Prepare Draft Transmission File"
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
      Left            =   3636
      TabIndex        =   10
      Top             =   1200
      Width           =   5004
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE:  Enter a '0' for all Cycles, or leave blank if do not bill by cycle."
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
      Height          =   324
      Left            =   3000
      TabIndex        =   9
      Top             =   3984
      Width           =   6588
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1524
      Left            =   2424
      Top             =   4872
      Width           =   7284
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Height          =   2460
      Left            =   2424
      Top             =   2424
      Width           =   7284
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Cycles"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   0
      Left            =   2520
      TabIndex        =   8
      Top             =   5136
      Width           =   2076
   End
   Begin VB.Line Line1 
      X1              =   3984
      X2              =   9648
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Press F10 to process selections."
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
      Height          =   372
      Left            =   3648
      TabIndex        =   7
      Top             =   4200
      Width           =   3468
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Draft Date:"
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
      Height          =   372
      Index           =   7
      Left            =   4056
      TabIndex        =   6
      Top             =   2856
      Width           =   1716
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3228
      Top             =   840
      Width           =   5772
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
Attribute VB_Name = "frmDraftPrepareFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim BegRoute As String, EndRoute As String
Dim UseCycle As Boolean
Dim Grpt As Boolean, CycleCnt As Integer
Dim Cycle(1 To 16) As Integer

Private Sub cmdExit_Click()
  frmUBDraftMenu.Show
  Unload frmDraftPrepareFile
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via DraftPrepareFile by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub
Private Sub fptxtCycleSel_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtCycleSel_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim cnt As Integer
  If KeyCode = vbKeyReturn Then
    If Len(fptxtCycleSel.Text) <> 0 Then
      getcyclelist
    Else
      cmdOk.SetFocus
    End If
  End If
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
Private Sub cmdOk_Click()
  DeActivateControls Me, True
  UBBuildTransmitFile
  ActivateControls Me, True
  CycleCnt = 0
  fptxtCycleSel.Text = ""
  fptxtcycle.Text = ""
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
      cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      cmdOk_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Dim UBSetupreclen As Integer
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  txtDate1.Text = Format(Now, "mm/dd/yyyy")
  CycleCnt = 0
  fptxtCycleSel.Text = ""
  Erase Cycle
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub getcyclelist()
Dim TCyc As String, ThisCycle As Integer, cnt As Integer
  TCyc$ = QPTrim$(fptxtCycleSel.Text)
  If TCyc$ = "0" Then
    fptxtcycle.Text = ""
    Erase Cycle
    cmdOk.SetFocus
  Else
    If Len(TCyc$) > 0 Then
      ThisCycle = Val(fptxtCycleSel.Text)
      For cnt = 1 To 16
        If ThisCycle = Cycle(cnt) Then
          GoTo DupeExit
        End If
      Next
      CycleCnt = CycleCnt + 1
      Cycle(CycleCnt) = ThisCycle
      fptxtcycle.Text = ""
      For cnt = 1 To CycleCnt
        If cnt = CycleCnt Then
          fptxtcycle.Text = fptxtcycle.Text & Cycle(cnt)
        Else
          fptxtcycle.Text = fptxtcycle.Text & Cycle(cnt) & ","
        End If
      Next
    End If
  End If
DupeExit:
  fptxtCycleSel.Text = ""
End Sub

Private Sub UBBuildTransmitFile()
  Dim UBSetupLen As Integer, IndexName As String, WarrFlag As Boolean
  Dim UBCustRecLen As Integer, UBCust As Integer, NWoodFlag As Boolean
  Dim NumOfRecs As Long, cnt As Long, BalRecFlag As Boolean
  Dim CustCycle As Integer, CustOk As Boolean, CCnt As Integer
  Dim CstCnt As Long, llow As Long, hhigh As Long, BankCnt As Integer
  Dim PrevBank As String, GTotal As Double, CompanyAcct As String
  Dim DraftFile As Integer, PlyFlag As Boolean, GATot As Double
  Dim PayRecLen As Integer, PayFileName As String, Make99Flag As Boolean
  Dim BDate As Integer, Done As Boolean, DraftDate As String
  Dim NewDraftFile As String, Step1 As Boolean, Step2 As Boolean
  Dim Step3 As Boolean, Step4 As Boolean, Step5 As Boolean, CustBal As Double
  Dim FedIDNum As String, DraftFileNum As Integer, Counter As Integer
  Dim BillAmt As String, TotalAmountn As Double, AcctNumber As String
  Dim nme As String, BankAcct As String, sp As String, Trac As Integer
  Dim Trace As String, hashh As Double, Number As Integer, Pay99File As Integer
  Dim hash As String, TotalAmount As String, TotSize As Double
  Dim BlockSize As Single, BlockSizeS As String, FillSize As Single
  Dim outfile As Integer, FasFlag As Boolean, NumOfLines As Integer
  Dim ZZCnt As Integer
 UBLog " IN: Draft Build Transmit File"


  '*****************
  
  ReDim DFTRec(1) As DraftRptType
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  ReDim UBDraftRec(1) As UBDraftRecType
  DraftFile = FreeFile
  Open UBPath$ + "UBSDRAFT.dat" For Random Access Read Shared As #DraftFile Len = Len(UBDraftRec(1))
  Get DraftFile, 1, UBDraftRec(1)
  Close

  CompanyAcct$ = QPTrim$(UBDraftRec(1).COMPACCT)
  If Val(CompanyAcct$) > 0 Then
    BalRecFlag = True
  End If

'load setup file
  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen

  TOWNNAME$ = UCase$(UBSetUpRec(1).UTILNAME)

  If InStr(TOWNNAME$, "WARRENTON") > 0 Then
    WarrFlag = True
  End If

  If InStr(TOWNNAME$, "NORWOOD") > 0 Then
    NWoodFlag = True
  End If
  If InStr(TOWNNAME$, "LEE") > 0 Then
    NWoodFlag = True
  End If
  If InStr(TOWNNAME$, "PLYMOUTH") > 0 Then
    PlyFlag = True
  End If

'  IF INSTR(TownName$, "FAISON") > 0 THEN
'    'STOP
'    FasFlag = True
'  END IF

  ReDim PayRec(1) As UBPaymentRecType
  PayRecLen = Len(PayRec(1))
  PayFileName$ = UBPath$ + "UBPAY99.DAT"

  If UBSetUpRec(1).Make99File = "Y" Then
    If Not Check99File Then
      GoTo DraftExit
    End If
    Make99Flag = True
  End If

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen

  NumOfRecs& = LOF(UBCust) \ UBCustRecLen

'      TCyc$ = QPTrim$(Form$(2, 0))
'      If Len(TCyc$) > 0 Then
'        ThisCycle = Value#(Form$(2, 0), ECode)
'        For cnt = 1 To CycleCnt
'          If ThisCycle = Cycle(cnt) Then
'            GoTo NotDoneHere
'          End If
'        Next
'        CycleCnt = CycleCnt + 1
'        Cycle(CycleCnt) = ThisCycle
'        LSet Form$(CycleCnt + 2, 0) = QPTrim$(Str$(ThisCycle))
'        Action = 1
'      End If
      BDate = Date2Num(txtDate1.Text)
      If BDate > 0 Then
        GoSub ProcessDraft
        Done = True
      Else
        MsgBox "Invalid Draft Date! Please correct and try again.", vbOKOnly, "Invalid Date"
      End If

ProcessDraft:

  KillACHFiles

  DraftDate$ = txtDate1
  DraftDate$ = Left$(DraftDate$, 2) + Mid$(DraftDate$, 4, 2) + Right$(DraftDate$, 2)

  If Val(DraftDate$) = 0 Then Exit Sub

  NewDraftFile$ = QPTrim$(UBDraftRec(1).FileName)

  If Len(NewDraftFile$) = 0 Then
    NewDraftFile$ = UBPath$ + "DS" + DraftDate$ + ".ACH"
  End If
  Load frmDraftMsg
  frmDraftMsg.Label(0).Caption = "Building Record Type 1"
  frmDraftMsg.Show
  DoEvents
  Do
DraftFormLoop:
    
    ' Process Record Type 1
    If Not Step1 Then
      GoSub DraftProcessStep1
      frmDraftMsg.Label(0).Caption = "Building Record Type 1 ..Done!"
      'frmDraftMsg.Refresh
      DoEvents
      frmDraftMsg.Label(1).Caption = "Building Record Type 5"
      'frmDraftMsg.Show 1
      DoEvents
      GoTo DraftFormLoop
    End If

    If Not Step2 Then
      GoSub DraftProcessStep2
      frmDraftMsg.Label(1).Caption = "Building Record Type 5 ..Done!"
      'frmDraftMsg.Show 1
      DoEvents
      frmDraftMsg.Label(2).Caption = "Building Record Type 6 "
      'frmDraftMsg.Show 1
      DoEvents
      GoTo DraftFormLoop
    End If

    If Not Step3 Then
      GoSub DraftProcessStep3
      frmDraftMsg.Label(2).Caption = "Building Record Type 6 ..Done!"
      'frmDraftMsg.Show 1
      DoEvents
      frmDraftMsg.Label(3).Caption = "Building Record Type 8 "
      'frmDraftMsg.Show 1
      DoEvents
      GoTo DraftFormLoop
    End If

    If Not Step4 Then
      GoSub DraftProcessStep4
      frmDraftMsg.Label(3).Caption = "Building Record Type 8 ..Done!"
      'frmDraftMsg.Show 1
      DoEvents
      frmDraftMsg.Label(4).Caption = "Building Record Type 9"
      'frmDraftMsg.Show 1
      DoEvents
      GoTo DraftFormLoop
    End If

    If Not Step5 Then
      GoSub DraftProcessStep5
      frmDraftMsg.Label(4).Caption = "Building Record Type 9 ..Done!"
      If Make99Flag Then
        frmDraftMsg.Label(5).Caption = "Building Payment File  ..Done!"
        'frmDraftMsg.Show 1
        DoEvents
        frmDraftMsg.Label(6).Caption = "File Name Is: " + NewDraftFile$
        'frmDraftMsg.Show 1
      Else
        frmDraftMsg.Label(6).Caption = "File Name Is: " + NewDraftFile$
        'frmDraftMsg.Show 1
      End If
      DoEvents
      GoTo DraftFormLoop
    End If


  Loop Until Done

DraftExit:
  Close

  KillFile UBPath$ + "UBDRAFT1.dat"
  KillFile UBPath$ + "UBDRAFT5.dat"
  KillFile UBPath$ + "UBDRAFT6.dat"
  KillFile UBPath$ + "UBDRAFT9.dat"
  KillFile UBPath$ + "UBDRAFT8.dat"


Exit Sub

'RETURN


DraftProcessStep1:

  FedIDNum$ = QPTrim$(UBDraftRec(1).FEDPREFX + UBDraftRec(1).FEDID) + "00"

  ReDim UBDraftRecord1(1) As UBDraftRecord1Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT1.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord1(1))
  UBDraftRecord1(1).Field1 = "1"
  UBDraftRecord1(1).Field2 = "01"
  UBDraftRecord1(1).Field3 = " " + QPTrim$(UBDraftRec(1).BANKDEST)
  UBDraftRecord1(1).Field4 = " " + QPTrim$(UBDraftRec(1).BANKORIG)
  UBDraftRecord1(1).Field5 = Right$(Date$, 2) + Left$(Date$, 2) + Mid$(Date$, 4, 2)
  UBDraftRecord1(1).Field6 = Left$(Time$, 2) + Mid$(Time$, 4, 2)
  UBDraftRecord1(1).Field7 = "A"
  UBDraftRecord1(1).Field8 = "094"
  UBDraftRecord1(1).Field9 = "10"
  UBDraftRecord1(1).Field10 = "1"
  UBDraftRecord1(1).Field11 = QPTrim$(UCase$(UBDraftRec(1).BankName))
  UBDraftRecord1(1).Field12 = QPTrim$(UCase$(UBDraftRec(1).BANKLOC))
  UBDraftRecord1(1).Field13 = "        "        'Must = 8 Spaces
  Put DraftFileNum, 1, UBDraftRecord1(1)
  Close DraftFileNum
  Step1 = True
Return

DraftProcessStep2:
  ReDim UBDraftRecord5(1) As UBDraftRecord5Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT5.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord5(1))
  UBDraftRecord5(1).Field1 = "5"
  UBDraftRecord5(1).Field2 = "200"
  UBDraftRecord5(1).Field3 = Left$(TOWNNAME$, 16)
  UBDraftRecord5(1).Field4 = "                    "

  UBDraftRecord5(1).Field5 = FedIDNum$
  UBDraftRecord5(1).Field6 = "PPD"
  UBDraftRecord5(1).Field7 = "UTIL BILL"
  UBDraftRecord5(1).Field8 = Right$(DraftDate$, 2) + Left$(DraftDate$, 2) + Mid$(DraftDate$, 3, 2)
  UBDraftRecord5(1).Field9 = Right$(DraftDate$, 2) + Left$(DraftDate$, 2) + Mid$(DraftDate$, 3, 2)
  UBDraftRecord5(1).Field10 = "   "             'Reserved w/3 blanks
  UBDraftRecord5(1).Field11 = "1"
  UBDraftRecord5(1).Field12 = Left$(UBDraftRec(1).BANKORIG, 8)
  UBDraftRecord5(1).Field13 = "0000001"
  Put DraftFileNum, 1, UBDraftRecord5(1)
  Close DraftFileNum
  Step2 = True
Return

DraftProcessStep3:
  Counter = 0

  ReDim UBDraftRecord6(1) As UBDraftRecord6Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT6.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord6(1))
  Close DraftFileNum
  Kill UBPath$ + "UBDRAFT6.DAT"
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT6.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord6(1))

  'GO THRU DATA FILE HERE
  If Make99Flag Then
    GoSub Setup99File
  End If

  For cnt& = 1 To NumOfRecs&
    Get UBCust, cnt&, UBCustRec(1)
    CustCycle = UBCustRec(1).BILLCYCL
    CustOk = False
    If CycleCnt > 0 Then
      For CCnt = 1 To CycleCnt
        If Cycle(CCnt) = 0 Then
          CustOk = True
        ElseIf CustCycle = Cycle(CCnt) Then
          CustOk = True
          Exit For
        End If
      Next
    Else
      CustOk = True
    End If

    If CustOk Then
      If UBCustRec(1).Status = "A" Or UBCustRec(1).Status = "B" Then
'changed the line below to only look at usedraft flag 10-28-04  ps
        If (UBCustRec(1).USEDRAFT = "Y") Then ' or (Len(QPTrim$(UBCustRec(1).BankName)) > 0) Then
          CustBal# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
          If CustBal# > 0 Then
            ' Process Customer Here
            BillAmt$ = (Str$(CustBal# * 100))     'Remove the Decimals
            BillAmt$ = LTrim$(BillAmt$)
            If Len(BillAmt$) < 10 Then BillAmt$ = String$(10 - Len(BillAmt$), "0") + BillAmt$
            'Get Running Billed Amount
            TotalAmountn# = TotalAmountn# + (CustBal# * 100)
            Counter = Counter + 1
            AcctNumber$ = Str$(cnt&)
            AcctNumber$ = Right$(AcctNumber$, Len(AcctNumber$) - 1)
            If Len(AcctNumber$) < 15 Then
              AcctNumber$ = AcctNumber$ + String$(15 - Len(AcctNumber$), 32)
            End If
            nme$ = UBCustRec(1).CustName
            If Len(nme$) < 22 Then
              nme$ = nme$ + String$(22 - Len(nme$), 32)
            Else
              nme$ = Left$(nme$, 22)
            End If

            'Check for Spaces WithIn Bank Account Numbered as Entered by Custo
            BankAcct$ = QPTrim$(UBCustRec(1).BankAcct)
            sp = InStr(BankAcct$, " ")
            If sp > 0 Then
              BankAcct$ = Left$(BankAcct$, sp - 1) + Right$(BankAcct$, Len(BankAcct$) - sp)
            End If
            If Len(BankAcct$) < 17 Then
              BankAcct$ = BankAcct$ + String$(17 - Len(BankAcct$), 32)
            End If
            Trac = Trac + 1
            Trace$ = Str$(Trac)
            Trace$ = Right$(Trace$, Len(Trace$) - 1)

            If Len(Trace$) < 7 Then
              Trace$ = String$(7 - Len(Trace$), "0") + Trace$
            End If
            UBDraftRecord6(1).Field1 = "6"

'HERE
            Select Case UBCustRec(1).AcctType
            Case "S"
              UBDraftRecord6(1).Field2 = "37"       ' Designates Savings Acct
            Case Else
              UBDraftRecord6(1).Field2 = "27"       ' Designates Checking Acct
            End Select
            UBDraftRecord6(1).Field3 = Left$(UBCustRec(1).TRANSIT, 8)
            UBDraftRecord6(1).Field4 = Right$(UBCustRec(1).TRANSIT, 1)
            UBDraftRecord6(1).Field5 = Left$(BankAcct$, 17)
            UBDraftRecord6(1).Field6 = BillAmt$   ' All zero's for Prenote
            UBDraftRecord6(1).Field7 = AcctNumber$
            UBDraftRecord6(1).Field8 = UCase$(nme$)
            UBDraftRecord6(1).Field9 = "  "
            UBDraftRecord6(1).Field10 = "0"
            UBDraftRecord6(1).Field11 = Left$(UBDraftRec(1).BANKORIG, 8) + Trace$
            Put DraftFileNum, Counter, UBDraftRecord6(1)
            hashh# = hashh# + Val(Left$(UBCustRec(1).TRANSIT, 8))
            Number = Number + 1
            If Make99Flag Then
              GoSub Make99Rec
            End If
          End If
        End If
      End If
    End If
  Next

  If BalRecFlag Then
    ' Process Customer Here
    BillAmt$ = Str$(TotalAmountn#)      'Remove the Decimals
    BillAmt$ = LTrim$(BillAmt$)
    If Len(BillAmt$) < 10 Then BillAmt$ = String$(10 - Len(BillAmt$), "0") + BillAmt$
    'Get Running Billed Amount
    Counter = Counter + 1

    AcctNumber$ = "0"   'STR$(Cnt&)
    AcctNumber$ = Right$(AcctNumber$, Len(AcctNumber$) - 1)
    If Len(AcctNumber$) < 15 Then
      AcctNumber$ = AcctNumber$ + String$(15 - Len(AcctNumber$), 32)
    End If
    nme$ = TOWNNAME$
    If Len(nme$) < 22 Then
      nme$ = nme$ + String$(22 - Len(nme$), 32)
    Else
      nme$ = Left$(nme$, 22)
    End If

    BankAcct$ = CompanyAcct$
    sp = InStr(BankAcct$, " ")
    If sp > 0 Then
      BankAcct$ = Left$(BankAcct$, sp - 1) + Right$(BankAcct$, Len(BankAcct$) - sp)
    End If

    If Len(BankAcct$) < 17 Then
      BankAcct$ = BankAcct$ + String$(17 - Len(BankAcct$), 32)
    End If
    Trac = Trac + 1
    Trace$ = Str$(Trac)
    Trace$ = Right$(Trace$, Len(Trace$) - 1)

    If Len(Trace$) < 7 Then
      Trace$ = String$(7 - Len(Trace$), "0") + Trace$
    End If

    UBDraftRecord6(1).Field1 = "6"
    UBDraftRecord6(1).Field2 = "22"       ' Designates credit checking
    UBDraftRecord6(1).Field3 = Left$(UBDraftRec(1).BANKDEST, 8)
    UBDraftRecord6(1).Field4 = Right$(UBDraftRec(1).BANKDEST, 1)
    UBDraftRecord6(1).Field5 = Left$(BankAcct$, 17)
    UBDraftRecord6(1).Field6 = BillAmt$   ' All zero's for Prenote
    UBDraftRecord6(1).Field7 = AcctNumber$
    UBDraftRecord6(1).Field8 = UCase$(nme$)
    UBDraftRecord6(1).Field9 = "  "
    UBDraftRecord6(1).Field10 = "0"
    UBDraftRecord6(1).Field11 = Left$(UBDraftRec(1).BANKORIG, 8) + Trace$

    Put DraftFileNum, Counter, UBDraftRecord6(1)
    hashh# = hashh# + Val(Left$(UBDraftRec(1).BANKDEST, 8))
    Number = Number + 1
  End If

  Close DraftFileNum

  If Make99Flag Then
    Close Pay99File
  End If
  Step3 = True

Return

DraftProcessStep4:
  hash$ = Str$(hashh#)
  hash$ = Right$(hash$, Len(hash$) - 1)
  If Len(hash$) < 10 Then
    hash$ = String$(10 - Len(hash$), "0") + hash$
  End If
  If Len(hash$) > 10 Then
    hash$ = Right$(hash$, 10)
  End If

  If Len(Trace$) > 6 Then Trace$ = Right$(Trace$, 6)
  TotalAmount$ = Str$(TotalAmountn#)
  TotalAmount$ = Right$(TotalAmount$, Len(TotalAmount$) - 1)
  If Len(TotalAmount$) < 12 Then TotalAmount$ = String$(12 - Len(TotalAmount$), "0") + TotalAmount$
  ReDim UBDraftRecord8(1) As UBDraftRecord8Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT8.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord8(1))
  UBDraftRecord8(1).Field1 = "8"
  UBDraftRecord8(1).Field2 = "200"
  UBDraftRecord8(1).Field3 = Trace$
  UBDraftRecord8(1).Field4 = hash$
  UBDraftRecord8(1).Field5 = TotalAmount$       ' zero for prenote
  If BalRecFlag Then
    UBDraftRecord8(1).Field6 = TotalAmount$       ' zero for prenote
  Else
    UBDraftRecord8(1).Field6 = "000000000000"     ' zero for prenote
  End If

  UBDraftRecord8(1).Field7 = FedIDNum$
  UBDraftRecord8(1).Field8 = String$(19, 32)    ' Reserved
  UBDraftRecord8(1).Field9 = String$(6, 32)     ' Reserved for Federal Reserve use
  UBDraftRecord8(1).Field10 = Left$(UBDraftRec(1).BANKORIG, 8)
  UBDraftRecord8(1).Field11 = "0000001"
  Put DraftFileNum, 1, UBDraftRecord8(1)
  Close DraftFileNum
  Step4 = True
Return

DraftProcessStep5:
  TotSize# = Val(Trace$) + 4    ' Total Records= Trace + 4 control records
  TotSize# = TotSize# * 94      ' Total Bytes = 94 per record
  BlockSize! = TotSize# / 940   ' Rem Blocks Consist of Batchs of 10 Records
  If BlockSize! <> Int(BlockSize!) Then
    BlockSize! = Int(BlockSize!) + 1
    FillSize! = 940 - (TotSize# - (940 * (BlockSize! - 1)))
  Else
    FillSize! = 0
  End If

  BlockSizeS$ = Str$(BlockSize!)
  BlockSizeS$ = Right$(BlockSizeS$, Len(BlockSizeS$) - 1)
  If Len(BlockSizeS$) < 6 Then BlockSizeS$ = String$(6 - Len(BlockSizeS$), "0") + BlockSizeS$
  If Len(Trace$) < 8 Then Trace$ = String$(8 - Len(Trace$), "0") + Trace$

  ReDim UBDraftRecord9(1) As UBDraftRecord9Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT9.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord9(1))
  UBDraftRecord9(1).Field1 = "9"
  UBDraftRecord9(1).Field2 = "000001"           ' only 1 batch
  UBDraftRecord9(1).Field3 = BlockSizeS$
  UBDraftRecord9(1).Field4 = Trace$
  UBDraftRecord9(1).Field5 = hash$
  UBDraftRecord9(1).Field6 = TotalAmount$       ' zero for prenote
  If BalRecFlag Then
    UBDraftRecord9(1).Field7 = TotalAmount$       ' zero for prenote
  Else
    UBDraftRecord9(1).Field7 = "000000000000"     ' zero for prenote
  End If

  UBDraftRecord9(1).Field8 = String$(39, 32)    ' Reserved

  Put DraftFileNum, 1, UBDraftRecord9(1)
  Close DraftFileNum

  ' Now Put Them Together In File Name UBDFNOTE
  outfile = FreeFile
  Open NewDraftFile$ For Output As outfile

  'Width #outfile, 255

  ReDim UBDraftRecord1(1) As UBDraftRecord1Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT1.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord1(1))
  Get DraftFileNum, 1, UBDraftRecord1(1)
  Print #outfile, UBDraftRecord1(1).Field1;
  Print #outfile, UBDraftRecord1(1).Field2;
  Print #outfile, UBDraftRecord1(1).Field3;
  Print #outfile, UBDraftRecord1(1).Field4;
  Print #outfile, UBDraftRecord1(1).Field5;
  Print #outfile, UBDraftRecord1(1).Field6;
  Print #outfile, UBDraftRecord1(1).Field7;
  Print #outfile, UBDraftRecord1(1).Field8;
  Print #outfile, UBDraftRecord1(1).Field9;
  Print #outfile, UBDraftRecord1(1).Field10;
  Print #outfile, UBDraftRecord1(1).Field11;
  Print #outfile, UBDraftRecord1(1).Field12;
  Print #outfile, UBDraftRecord1(1).Field13;
  If NWoodFlag = 0 And FasFlag = 0 Then
    Print #outfile,
  ElseIf FasFlag Then
    Print #outfile, Chr$(13);
  End If

  Close DraftFileNum
  ReDim UBDraftRecord5(1) As UBDraftRecord5Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT5.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord5(1))
  Get DraftFileNum, 1, UBDraftRecord5(1)
  Print #outfile, UBDraftRecord5(1).Field1;
  Print #outfile, UBDraftRecord5(1).Field2;
  Print #outfile, UBDraftRecord5(1).Field3;
  Print #outfile, UBDraftRecord5(1).Field4;
  Print #outfile, UBDraftRecord5(1).Field5;
  Print #outfile, UBDraftRecord5(1).Field6;
  Print #outfile, UBDraftRecord5(1).Field7;
  Print #outfile, UBDraftRecord5(1).Field8;
  Print #outfile, UBDraftRecord5(1).Field9;
  Print #outfile, UBDraftRecord5(1).Field10;
  Print #outfile, UBDraftRecord5(1).Field11;
  Print #outfile, UBDraftRecord5(1).Field12;
  Print #outfile, UBDraftRecord5(1).Field13;
  If NWoodFlag = 0 And FasFlag = 0 Then
    Print #outfile,
  ElseIf FasFlag Then
    Print #outfile, Chr$(13);
  End If

  Close DraftFileNum

  ReDim UBDraftRecord6(1) As UBDraftRecord6Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT6.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord6(1))
  NumOfLines = LOF(DraftFileNum) / 94

  For cnt = 1 To NumOfLines
    Get DraftFileNum, cnt, UBDraftRecord6(1)
    Print #outfile, UBDraftRecord6(1).Field1;
    Print #outfile, UBDraftRecord6(1).Field2;
    Print #outfile, UBDraftRecord6(1).Field3;
    Print #outfile, UBDraftRecord6(1).Field4;
    Print #outfile, UBDraftRecord6(1).Field5;
    Print #outfile, UBDraftRecord6(1).Field6;
    Print #outfile, UBDraftRecord6(1).Field7;
    Print #outfile, UBDraftRecord6(1).Field8;
    Print #outfile, UBDraftRecord6(1).Field9;
    Print #outfile, UBDraftRecord6(1).Field10;
    Print #outfile, UBDraftRecord6(1).Field11;
    If NWoodFlag = 0 And FasFlag = 0 Then
      Print #outfile,
    ElseIf FasFlag Then
      Print #outfile, Chr$(13);
    End If

  Next cnt
  Close DraftFileNum

  ReDim UBDraftRecord8(1) As UBDraftRecord8Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT8.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord8(1))
  Get DraftFileNum, 1, UBDraftRecord8(1)
  Print #outfile, UBDraftRecord8(1).Field1;
  Print #outfile, UBDraftRecord8(1).Field2;
  Print #outfile, UBDraftRecord8(1).Field3;
  Print #outfile, UBDraftRecord8(1).Field4;
  Print #outfile, UBDraftRecord8(1).Field5;
  Print #outfile, UBDraftRecord8(1).Field6;
  Print #outfile, UBDraftRecord8(1).Field7;
  Print #outfile, UBDraftRecord8(1).Field8;
  Print #outfile, UBDraftRecord8(1).Field9;
  Print #outfile, UBDraftRecord8(1).Field10;
  Print #outfile, UBDraftRecord8(1).Field11;
  If NWoodFlag = 0 And FasFlag = 0 Then
    Print #outfile,
  ElseIf FasFlag Then
    Print #outfile, Chr$(13);
  End If
  Close DraftFileNum

  ReDim UBDraftRecord9(1) As UBDraftRecord9Type
  DraftFileNum = FreeFile
  Open UBPath$ + "UBDRAFT9.dat" For Random Access Read Write Shared As #DraftFileNum Len = Len(UBDraftRecord9(1))
  Get DraftFileNum, 1, UBDraftRecord9(1)
  Print #outfile, UBDraftRecord9(1).Field1;
  Print #outfile, UBDraftRecord9(1).Field2;
  Print #outfile, UBDraftRecord9(1).Field3;
  Print #outfile, UBDraftRecord9(1).Field4;
  Print #outfile, UBDraftRecord9(1).Field5;
  Print #outfile, UBDraftRecord9(1).Field6;
  Print #outfile, UBDraftRecord9(1).Field7;
  Print #outfile, UBDraftRecord9(1).Field8;
  If NWoodFlag = 0 And FasFlag = 0 Then
    Print #outfile,
  ElseIf FasFlag Then
    Print #outfile, Chr$(13);
  End If

  Close DraftFileNum
  If NWoodFlag = 0 Then
    For cnt = 1 To FillSize! / 94
      Print #outfile, String$(94, "9")
    Next cnt
  End If
  Close
  Step5 = True
  Done = True
Return

Make99Rec:

  ReDim PayRec(1) As UBPaymentRecType
  PayRec(1).OPERNUM = 99
  PayRec(1).PAYDATE = BDate
  PayRec(1).CustAcct = cnt&
  PayRec(1).CustName = UBCustRec(1).CustName
  PayRec(1).CUSTADDR = UBCustRec(1).ADDR1
  'PayRec(1).CUSTCMNT= UBBillRec(1).
  PayRec(1).AMTOWED = CustBal#
  PayRec(1).TENDERTY = "BANK DRAFT"
  PayRec(1).CASHAMT = PayRec(1).AMTOWED  'was setting 0
  PayRec(1).CHKAMT = 0
  PayRec(1).AMTRECD = PayRec(1).AMTOWED
  PayRec(1).Change = 0
  PayRec(1).Desc = "DRAFT PAYMENT TRANS"

  For ZZCnt = 1 To 15
    PayRec(1).PaidOwed(ZZCnt).AMTOWE1 = UBCustRec(1).CurrRevAmts(ZZCnt)
    PayRec(1).PaidOwed(ZZCnt).AMTPD1 = PayRec(1).PaidOwed(ZZCnt).AMTOWE1
  Next

  PayRec(1).TOTOWED = PayRec(1).AMTOWED
  PayRec(1).AMTPAID = PayRec(1).AMTOWED

  Put #Pay99File, , PayRec(1)

Return

Setup99File:

  Pay99File = FreeFile
  Open PayFileName$ For Output As #Pay99File
  Close Pay99File

  Pay99File = FreeFile
  Open PayFileName$ For Random Shared As Pay99File Len = PayRecLen

Return

End Sub

Private Sub txtDate1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
    fptxtCycleSel.SetFocus
  End If
End Sub
