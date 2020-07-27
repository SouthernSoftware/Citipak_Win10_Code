VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmGrabTrans 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grab Transactions"
   ClientHeight    =   8895
   ClientLeft      =   30
   ClientTop       =   495
   ClientWidth     =   12195
   Icon            =   "frmGrabTrans.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboModule 
      Height          =   405
      Left            =   6105
      TabIndex        =   0
      Top             =   3480
      Width           =   2850
      _Version        =   196608
      _ExtentX        =   5027
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   2
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   2
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   16777215
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   16777215
      ThreeDOutsideShadowColor=   12632256
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   8421504
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   8421504
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
      GrayAreaColor   =   12632256
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   14737632
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   12632256
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
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmGrabTrans.frx":08CA
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Esc &Cancel"
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
      Left            =   10032
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7488
      Width           =   1332
   End
   Begin VB.CommandButton cmdOk 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
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
      Left            =   8256
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7488
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   8532
      Width           =   12192
      _ExtentX        =   21511
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7117
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "10:59 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "11/20/2009"
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
   Begin EditLib.fpDateTime fpDate 
      Height          =   372
      Left            =   6102
      TabIndex        =   1
      Top             =   4224
      Width           =   1692
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
      BackColor       =   16777215
      ForeColor       =   0
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   16777215
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   16777215
      ThreeDOutsideShadowColor=   12632256
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
      ButtonStyle     =   2
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   14737632
      ThreeDTextShadowColor=   12632256
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
      Text            =   "11/06/2001"
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
      BorderGrayAreaColor=   12632256
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   4210752
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
      ButtonColor     =   12632256
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
      Caption         =   "Transactions Dated Thru:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   3102
      TabIndex        =   7
      Top             =   4296
      Width           =   2868
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   852
      Left            =   3216
      Top             =   1416
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Grab Transactions"
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
      Left            =   3984
      TabIndex        =   6
      Top             =   1656
      Width           =   4332
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   2124
      Left            =   2682
      Top             =   3024
      Width           =   6828
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Grab Transactions From:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   2958
      TabIndex        =   5
      Top             =   3576
      Width           =   3012
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      Height          =   972
      Left            =   3216
      Top             =   1296
      Width           =   5772
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuPrnScn 
         Caption         =   "Prin&t Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmGrabTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim GLAcct As GLAcctRecType
Dim GLAcctidx As GLAcctIndexType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim MCodeRec(1) As MiscCodeRecType
Dim GJRec(1) As TrEditRecType
Dim GJEdit As TrEditRecType
Dim IFEdit As TrEditRecType
'Dim ARSetUpRec(1) As TownSetUpType
Dim LPDate As Integer, HPDate As Integer, TempDay As Integer, BadTxAcct As Long
Dim CDActive As String, CashAcct As String, CDCash As String, CDDue As String
Public EPFN As String
Private Sub cmdOk_Click()
  Dim msgtogo As String, Trans As String
  If DateOk = True Then
    If Not Exist("GLTRXED.DAT") Then
      If fpcboModule.ListIndex <> -1 Then
        fpcboModule.Col = 0
        Trans = QPTrim$(fpcboModule.ColText)
        If Trans = "CM" Then msgtogo = "All Users Must Exit The Cash Management System."
        If Trans = "UB" Then
          If Not Exist(StartPath + "\" + "GLUBTran.DAT") Then
            msgtogo = "All Users Must Exit The Utility System. "
          Else
            MsgBox "Prior Utility detail report MUST be printed before continuing.", vbOKOnly, "Option Canceled."
            Exit Sub
          End If
        End If
        If Trans = "NCTX" Then msgtogo = "All Users Must Exit The Tax System."
        If Trans = "BL" Then msgtogo = "All Users Must Exit The Business License System."
        If Trans = "DC" Then msgtogo = "All Users Must Exit The Decal System."
        If Trans = "VATX" Then msgtogo = "All Users Must Exit The Tax System."
        If Trans = "EP" Then
          EPFN = ""
          msgtogo = "Enter the Name of the Import File."
          frmExtractMsg.txtFileName.Visible = True
        End If
        frmExtractMsg.Label1 = msgtogo
        frmExtractMsg.Show 1, Me
        If frmExtractMsg.Exout <> 1 Then
          DoExtract
        Else
          Exit Sub
        End If
      Else
        MsgBox "Invalid Transaction Type.", vbOKOnly, "Invalid Type"
        fpcboModule.SetFocus
      End If
    Else
      MsgBox "An UnPosted Interface File Already Exits, Print Report, and either Transfer to General Journal or Post. Call Software Support If Questions.", vbOKOnly, "Previous Grab Incomplete"
    End If
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
      SendKeys "%C"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%O"
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub fpcboModule_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboModule.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboModule.ListIndex = -1
    fpcboModule.Action = ActionClearSearchBuffer
  End If
  If fpcboModule.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  GetCentDep CDActive, CashAcct, CDCash, CDDue
  StatusBar1.Panels.Item(1).Text = GLUserName
  Me.HelpContextID = hlpGrabTransactions
  fpDate.Text = Format(Now, "mm/dd/yyyy")
  GetPostDates LPDate, HPDate
  fpcboModule.AddItem "CM" & Chr(9) & "-Misc Cash Receipts"
  fpcboModule.AddItem "UB" & Chr(9) & "-Utility Billing"
  fpcboModule.AddItem "BL" & Chr(9) & "-Business License"
  If Exist("CitiTaxes.EXE") Then
    fpcboModule.AddItem "NCTX" & Chr(9) & "-Tax Billing"
  End If
  If Exist("VACitiTax.EXE") Then
    fpcboModule.AddItem "VATX" & Chr(9) & "-VA Tax Billing"
  End If
  If Exist("DCCust.DAT") Then
    If Exist("DC.EXE") Then
      fpcboModule.AddItem "DC" & Chr(9) & "-VA Vehicle Decals"
    End If
  End If
  fpcboModule.AddItem "EP" & Chr(9) & "-Import Third Party"
End Sub
Private Sub cmdCancel_Click()
  frmGetDistMenu.Show
  Unload frmGrabTrans
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdCancel.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        ClearInUse PWcnt
      End If
    End If
  End If
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub

Private Function DateOk()
  TempDay = DateDiff("d", "12/31/1979", fpDate)
  If TempDay >= LPDate And TempDay <= HPDate Then
    DateOk = True
  Else
    MsgBox "Date Is Not Within Allowable Posting Range, Please Correct and Try Again.", vbOKOnly, "Date Invalid"
    fpDate.SetFocus
    DateOk = False
  End If
End Function
Private Sub DoExtract()
  Dim Trans As String
  TempDay = (DateDiff("d", "12/31/1979", fpDate))
      FrmShowPctComp.Label1 = "Searching For Interface Transactions"
      FrmShowPctComp.cmdCancel.Enabled = False
      FrmShowPctComp.Show , Me
      DoEvents
      DeActivateControls frmGrabTrans, True
      fpcboModule.Col = 0
      Trans = QPTrim$(fpcboModule.ColText)
      If Trans = "CM" Then ExtractCM (TempDay)
      If Trans = "UB" Then ExtractUB (TempDay) 'ExtractUB (TempDay)
      If Trans = "NCTX" Then
        If Exist(StartPath + "\" + "CitiTaxes.EXE") Then
          ExtractNTX (TempDay)
        Else
          Unload FrmShowPctComp
          ActivateControls frmGrabTrans, True
          MsgBox "Invalid Tax Version, contact software support.", vbOKOnly, "Invalid Version"
        End If
      End If
      If Trans = "VATX" Then
        If Exist(StartPath + "\" + "VACitiTax.EXE") Then
          ExtractVATX (TempDay)
        Else
          Unload FrmShowPctComp
          ActivateControls frmGrabTrans, True
          MsgBox "Invalid Tax Version, contact software support.", vbOKOnly, "Invalid Version"
        End If
      End If

      If Trans = "BL" Then ExtractBL (TempDay)
      If Trans = "DC" Then
        If Exist(StartPath + "\" + "DC.EXE") Then
          ExtractDC (TempDay)
        Else
          Unload FrmShowPctComp
          ActivateControls frmGrabTrans, True
          MsgBox "Invalid Selection, try again or contact software support.", vbOKOnly, "Invalid Selection"
        End If
      End If
      If Trans = "EP" Then ExtractEP (TempDay)
      ActivateControls frmGrabTrans, True
      cmdCancel_Click
End Sub
'General Text Third Party File Import
Private Sub ExtractEP(ThruDate%)
  
  Dim Today As String, Ref As String, Dash80 As String, P2S As String
  Dim GJReclen As Integer, RptFile As Integer, EPTransRecLen As Integer
  Dim EPTran As Integer, NumOfTRecs As Long, TCnt As Long, TempD As Integer
  Dim FoundCnt As Long, NGCnt As Integer, GJFile As Integer
  Dim NumEdTrans As Integer, EPFile As Integer, cnt As Long
  Dim FirstTran As Long, ThisDate As Integer, WorkDate As Integer
  Dim DayCount As Integer, MCnt As Integer, MiscDebAmt As Double
  Dim FindCount As Long, FundCnt As Integer, Process As Integer
  Dim Acct As String, AcctName As String, AcctR As Integer
  Dim FoundFund As Integer, FCnt As Integer, Cash As Integer
  Dim SetUpRecLen As Integer, SetupFile As Integer, CAcct As String
  Dim Handle As Integer, MiscCrdAmt As Double, TDate As String
  Dim totDeb As Double, totCred As Double
  On Local Error GoTo DoerrStuff
  If Len(EPFN) > 0 Then
    If Not Exist(EPFN) Then
      FrmShowPctComp.ShowPctComp 1, 1
      MsgBox "Import File Does Not Exist", vbOKOnly, "Procedure Cancelled"
      Exit Sub
    End If
  Else
    FrmShowPctComp.ShowPctComp 1, 1
    MsgBox "Invalid File Name", vbOKOnly, "Procedure Cancelled"
    Exit Sub
  End If
  Dim MiscRec#(500), MiscDAmt#(500), MiscCAmt#(500), Fund$(100), FundAmt#(100)

  ReDim GLSetUpRec(1) As GLSetupRecType
  SetUpRecLen = Len(GLSetUpRec(1))
  SetupFile = FreeFile
  Open "GLSETUP.DAT" For Random Access Read Write Shared As SetupFile Len = SetUpRecLen
  Get SetupFile, 1, GLSetUpRec(1)
  Close SetupFile
  CAcct = QPTrim(GLSetUpRec(1).CRCashAcct)
  
  Erase GLSetUpRec

  Today$ = Date$
  Ref$ = "EP" + Left$(Today$, 2) + Mid$(Today$, 4, 2) + Right$(Today$, 2)

  ReDim TranInfo(1) As TranRecInfoType

  Dash80$ = String$(80, "-")
  P2S$ = Space$(4)

  GJReclen = Len(GJRec(1))

  RptFile = FreeFile
  Open "GLEPTRX.RPT" For Output As RptFile

  'ClearBox
  'QPrintRC "Searching Cash Transactions.", 11, 26, 126
  'QPrintRC "New Transactions:", 13, 29, Cnf.HiLite

  ReDim EPTransRec(1) As EPTransRecType
  EPTransRecLen = Len(EPTransRec(1))

  EPTran = FreeFile
  Open EPFN For Random Shared As EPTran Len = EPTransRecLen
  NumOfTRecs& = LOF(EPTran) \ EPTransRecLen
  'Lock #CMTran

  For TCnt& = 1 To NumOfTRecs&  'To 1 Step -1
    Get #EPTran, TCnt&, EPTransRec(1)
    'If Len(QPTrim$(CMTransRec(1).Trans2GL)) = 0 Or QPTrim$(CMTransRec(1).Trans2GL) = "N" Then
      'Store trans rec numbers and dates in array
      TDate$ = EPTransRec(1).EPMonth + "/" + EPTransRec(1).EPDay + "/" + EPTransRec(1).EPYear
      TempD% = DateDiff("d", "12/31/1979", TDate$)
      If TempD% <= ThruDate% Then
        FoundCnt = FoundCnt + 1
        totDeb = Round#(totDeb + CDbl(EPTransRec(1).EPDebit))
        totCred = Round#(totCred + CDbl(EPTransRec(1).EPCredit))
        ReDim Preserve TranInfo(FoundCnt) As TranRecInfoType
        TranInfo(FoundCnt).TranDate = TempD%
        TranInfo(FoundCnt).TranRecNo = TCnt&
      End If
'    Else
'      NGCnt = NGCnt + 1
'    End If
    'RSet P2S$ = Str$(FoundCnt)
    'QPrintRC P2S$, 13, 47, Cnf.HiLite
    'SmallPause
    'Allow 250 Bad Dates Before Exiting
   'If NGCnt >= 250 Then Exit For
  Next

  If FoundCnt = 0 Then
    FrmShowPctComp.ShowPctComp 1, 1
    Close
    'ClearBox
    'Print Chr$(7);
    Call MainLog("NO EP to Grab for " + fpDate)
    MsgBox "No Transactions Found To InterFace", vbOKOnly, "No Transactions"
    'SLEEP 4
    GoTo SendExit
  End If
 ' Get #EPTran, NumOfTRecs&, EPTransRec(1)
 ' If NumOfTRecs& = Val(EPTransRec(1).EPAcct) Then
    If totDeb# <> totCred Then
      FrmShowPctComp.ShowPctComp 1, 1
      'Close
      Call MainLog("Invalid Totals for EP " + fpDate)
      MsgBox "Transactions Totals Are Not Balanced And Will Need Attention.", vbOKOnly, "Invalid Totals"
      'GoTo SendExit
    End If
'  Else
'    FrmShowPctComp.ShowPctComp 1, 1
'    Close
'    Call MainLog("Invalid Totals for EP " + fpDate)
'    MsgBox "Transactions Totals Are Not Balanced. Procedure Has Terminated.", vbOKOnly, "Invalid Totals"
'    GoTo SendExit
'  End If
 ' QSortTRec TranInfo(), FoundCnt     'sort'em by date. oldest first
  'Array (1), NumElem, Dir, StructSize, MemOff, MemSize
  GJFile = FreeFile
  Open "GLTRXED.DAT" For Random Access Read Write Shared As GJFile Len = GJReclen
  NumEdTrans = LOF(GJFile) \ GJReclen


  FirstTran = 1
'  ThisDate = TranInfo(1).TranDate
'  WorkDate = ThisDate

  For cnt = 1 To FoundCnt
    FrmShowPctComp.ShowPctComp cnt, FoundCnt
'    If ThisDate <> TranInfo(cnt).TranDate Then
'      ThisDate = TranInfo(cnt).TranDate
'      GoSub ProcessThisBunch
'      DayCount = 0
'      WorkDate = ThisDate
'    End If

    Get #EPTran, TranInfo(cnt).TranRecNo, EPTransRec(1)
      TDate$ = EPTransRec(1).EPMonth + "/" + EPTransRec(1).EPDay + "/" + EPTransRec(1).EPYear
      TempD% = DateDiff("d", "12/31/1979", TDate$)

      MiscDebAmt# = CDbl(EPTransRec(1).EPDebit)
      MiscDebAmt# = Round#(MiscDebAmt#)
      MiscCrdAmt# = CDbl(EPTransRec(1).EPCredit)
      MiscCrdAmt# = Round#(MiscCrdAmt#)
      Acct$ = EPTransRec(1).EPAcct
      Acct$ = QPStrip$(Acct$)
      Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
      AcctR = AcctFind(Acct$)
      If AcctR > 0 Then
        AcctName$ = GetAcctTitle(AcctR)
      Else
        AcctName$ = "UnDefined"
      End If
      GJRec(1).AcctRec = 0
      GJRec(1).AcctNum = QPTrim$(Acct$)
      GJRec(1).AcctName = AcctName$
      GJRec(1).TRDATE = TempD%
      GJRec(1).Ref = Ref$
      If MiscCrdAmt# <> 0 And MiscDebAmt# = 0 Then
        GJRec(1).CrAmt = MiscCrdAmt#
        GJRec(1).DrAmt = MiscDebAmt#
        GJRec(1).EType = "C"
        GJRec(1).Desc = "FROM External Import"
        GJRec(1).LDesc = QPTrim$(EPTransRec(1).EPDesc)
        GJRec(1).Src = "EP"
        Put #GJFile, , GJRec(1)
      ElseIf MiscCrdAmt# = 0 And MiscDebAmt# <> 0 Then
        GJRec(1).CrAmt = MiscCrdAmt#
        GJRec(1).DrAmt = MiscDebAmt#
        GJRec(1).EType = "D"
        GJRec(1).Desc = "FROM External Import"
        GJRec(1).LDesc = QPTrim$(EPTransRec(1).EPDesc)
        GJRec(1).Src = "EP"
        Put #GJFile, , GJRec(1)
      ElseIf MiscCrdAmt# <> 0 And MiscDebAmt# <> 0 Then
        GJRec(1).CrAmt = MiscCrdAmt#
        GJRec(1).DrAmt = 0
        GJRec(1).EType = "C"
        GJRec(1).Desc = "FROM External Import"
        GJRec(1).LDesc = QPTrim$(EPTransRec(1).EPDesc)
        GJRec(1).Src = "EP"
        Put #GJFile, , GJRec(1)
        GJRec(1).CrAmt = 0
        GJRec(1).DrAmt = MiscDebAmt#
        GJRec(1).EType = "D"
        GJRec(1).Desc = "FROM External Import"
        GJRec(1).LDesc = QPTrim$(EPTransRec(1).EPDesc)
        GJRec(1).Src = "EP"
        Put #GJFile, , GJRec(1)
      End If

    'If CMTransRec(1).TransSource = 1 Or CMTransRec(1).TransSource = 201 Then

'      If DayCount = 0 Then
'        'If Val(EPTransRec(1).EPDebit) > 0 Then
'
'          MiscDebAmt# = Val(EPTransRec(1).EPDebit)
'          MiscDebAmt# = Round#(MiscDebAmt#)
'          'If MiscRevAmt# <> 0 Then
'          MiscCrdAmt# = Val(EPTransRec(1).EPCredit)
'          MiscCrdAmt# = Round#(MiscCrdAmt#)
'          DayCount = DayCount + 1
'
'              MiscRec#(DayCount) = TranInfo(cnt).TranRecNo
'              MiscDAmt#(DayCount) = MiscDebAmt#
'              MiscCAmt#(DayCount) = MiscCrdAmt#
'              MiscDebAmt# = 0
'              MiscCrdAmt# = 0
'
'
'      Else
'          MiscDebAmt# = Val(EPTransRec(1).EPDebit)
'          MiscDebAmt# = Round#(MiscDebAmt#)
'          MiscCrdAmt# = Val(EPTransRec(1).EPCredit)
'          MiscCrdAmt# = Round#(MiscCrdAmt#)
'            DayCount = DayCount + 1
'            MiscRec#(DayCount) = TranInfo(cnt).TranRecNo
'              MiscDAmt#(DayCount) = MiscDebAmt#
'              MiscCAmt#(DayCount) = MiscCrdAmt#
'            MiscDebAmt# = 0
'            MiscCrdAmt# = 0
'
'      End If
      
    
  Next cnt
  'GoSub ProcessThisBunch

  'transactions As interfaced
'  For cnt = 1 To FoundCnt
'    Get #EPTran, TranInfo(cnt).TranRecNo, EPTransRec(1)
'    CMTransRec(1).Trans2GL = "Y"
'    Put #CMTran, TranInfo(cnt).TranRecNo, CMTransRec(1)
'  Next
'  Close
  Call MainLog("Completed EP Grab " + Str$(FoundCnt) + " for " + fpDate)
  MsgBox "Transaction Grab Complete.", vbOKOnly, "Complete"
SendExit:
  Exit Sub

'ProcessThisBunch:
'  ' Must Combine By Date and Then Do Cash Debit Entry For Total by Fund
'
' ' If DayCount <= 0 Then Return
'
'  FundCnt = 0   ' Set Funds Used to Zero
'  'ReDim EPTransRec(1) As EPTransRecType
'  For Process = 1 To DayCount
'    If MiscRec#(Process) <> 0 Then
'      Get #EPTran, MiscRec#(Process), EPTransRec(1)
'      Acct$ = EPTransRec(1).EPAcct
'      Acct$ = QPStrip$(Acct$)
'      Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
'      AcctR = AcctFind(Acct$)
'      If AcctR > 0 Then
'        AcctName$ = GetAcctTitle(AcctR)
'      Else
'        AcctName$ = "UnDefined"
'      End If
'      GJRec(1).AcctRec = 0
'      GJRec(1).AcctNum = QPTrim$(Acct$)
'      GJRec(1).AcctName = AcctName$
'      GJRec(1).TRDATE = WorkDate
'      GJRec(1).Ref = Ref$
'      If MiscCAmt#(Process) <> 0 And MiscDAmt#(Process) = 0 Then
'        GJRec(1).CrAmt = MiscCAmt#(Process)
'        GJRec(1).DrAmt = MiscDAmt#(Process)
'        GJRec(1).EType = "C"
'        GJRec(1).Desc = "FROM External Import"
'        GJRec(1).Src = "EP"
'        Put #GJFile, , GJRec(1)
'      ElseIf MiscCAmt#(Process) = 0 And MiscDAmt#(Process) <> 0 Then
'        GJRec(1).CrAmt = MiscCAmt#(Process)
'        GJRec(1).DrAmt = MiscDAmt#(Process)
'        GJRec(1).EType = "D"
'        GJRec(1).Desc = "FROM External Import"
'        GJRec(1).Src = "EP"
'        Put #GJFile, , GJRec(1)
'      ElseIf MiscDAmt#(Process) <> 0 And MiscCAmt#(Process) <> 0 Then
'        GJRec(1).CrAmt = MiscCAmt#(Process)
'        GJRec(1).DrAmt = 0
'        GJRec(1).EType = "C"
'        GJRec(1).Desc = "FROM External Import"
'        GJRec(1).Src = "EP"
'        Put #GJFile, , GJRec(1)
'        GJRec(1).CrAmt = 0
'        GJRec(1).DrAmt = MiscDAmt#(Process)
'        GJRec(1).EType = "D"
'        GJRec(1).Desc = "FROM External Import"
'        GJRec(1).Src = "EP"
'        Put #GJFile, , GJRec(1)
'      End If
''      'Add Up Fund Total Here for Cash Credit Entry
''      If FundCnt = 0 Then
''        FundCnt = 1
''        Fund$(FundCnt) = Left$(Acct$, GLFundLen%)
''        FundAmt#(FundCnt) = MiscAmt#(Process)
''      Else
''        FoundFund = 0
''        For FCnt = 1 To FundCnt
''          If Fund$(FCnt) = Left$(Acct$, GLFundLen%) Then
''            FoundFund = 1
''            FundAmt#(FCnt) = FundAmt#(FCnt) + MiscAmt#(Process)
''          End If
''        Next FCnt
''        If FoundFund = 0 Then
''          FundCnt = FundCnt + 1
''          Fund$(FundCnt) = Left$(Acct$, GLFundLen%)
''          FundAmt#(FundCnt) = MiscAmt#(Process)
''        End If
''      End If
'    End If
'  Next Process

' ' Now Make Matching Debit Entries to Cash Account
'
'  For Cash = 1 To FundCnt
'    Acct$ = Fund$(Cash) + CAcct$ 'CashAcct$
'    Acct$ = QPTrim$(Acct$)
'    Acct$ = QPStrip$(Acct$)
'    Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
'    AcctR = AcctFind(Acct$)
'    If AcctR > 0 Then
'      AcctName$ = GetAcctTitle(AcctR)
'    Else
'      AcctName$ = "UnDefined"
'    End If
'
'    GJRec(1).AcctRec = 0
'    GJRec(1).AcctNum = Acct$
'    GJRec(1).AcctName = AcctName$
'    GJRec(1).TRDATE = WorkDate
'    GJRec(1).Ref = Ref$
'    GJRec(1).DrAmt = FundAmt#(Cash)
'    GJRec(1).CrAmt = 0
'    GJRec(1).EType = "D"
'    GJRec(1).Desc = "FROM CASH MGMT"
'    GJRec(1).Src = "CR"
'    Put #GJFile, , GJRec(1)
'  Next Cash
BunchReturn:
  Return
DoerrStuff:
  If Err > 0 Then
  Unload FrmShowPctComp
  MsgBox "Error Code Was " + Err.Description + Str$(Err) + " ExtractEP CANCELED"
  End If
  Close
  Exit Sub
End Sub

'Cash Management Misc Transactions
Private Sub ExtractCM(ThruDate%)
  Dim Today As String, Ref As String, Dash80 As String, P2S As String
  Dim GJReclen As Integer, RptFile As Integer, CMTransRecLen As Integer
  Dim CMTran As Integer, NumOfTRecs As Long, TCnt As Long
  Dim FoundCnt As Long, NGCnt As Integer, GJFile As Integer
  Dim NumEdTrans As Integer, MCFile As Integer, cnt As Long
  Dim FirstTran As Long, ThisDate As Integer, WorkDate As Integer
  Dim DayCount As Integer, MCnt As Integer, MiscRevAmt As Double
  Dim FindCount As Long, FundCnt As Integer, Process As Integer
  Dim Acct As String, AcctName As String, AcctR As Integer
  Dim FoundFund As Integer, FCnt As Integer, Cash As Integer
  Dim SetUpRecLen As Integer, SetupFile As Integer, CAcct As String
  Dim CMSetuplen As Integer, Handle As Integer, BadAcct As Long
  BadAcct = 0
  If Exist("CMSetTown.DAT") Then
    ReDim CMSetUp(1) As CMSetupType
    CMSetuplen = Len(CMSetUp(1))
    Handle = FreeFile
    Open "CMSetTown.dat" For Random Shared As Handle Len = CMSetuplen    'open data file
    If LOF(Handle) > 0 Then
      Get #Handle, 1, CMSetUp(1)
    End If
    Close Handle
    If Not QPTrim(CMSetUp(1).GLInterface) = "Y" Then
      Unload FrmShowPctComp
      MsgBox "Cash Management Set NOT to Interface with GL", vbOKOnly, "Procedure Cancelled"
      Exit Sub
    End If
  Else
    Unload FrmShowPctComp
    MsgBox "Cash Management SetUp File Does Not Exist", vbOKOnly, "Procedure Cancelled"
    Exit Sub
  End If
  ReDim GLSetUpRec(1) As GLSetupRecType
  SetUpRecLen = Len(GLSetUpRec(1))
  SetupFile = FreeFile
  Open "GLSETUP.DAT" For Random Access Read Write Shared As SetupFile Len = SetUpRecLen
  Get SetupFile, 1, GLSetUpRec(1)
  Close SetupFile
  CAcct = QPTrim(GLSetUpRec(1).CRCashAcct)
  
  Erase GLSetUpRec

  Today$ = Date$
  Ref$ = "CM" + Left$(Today$, 2) + Mid$(Today$, 4, 2) + Right$(Today$, 2)

  ReDim TranInfo(1) As TranRecInfoType
  Dim MiscRec#(500), MiscAmt#(500), Fund$(100), FundAmt#(100)

  Dash80$ = String$(80, "-")
  P2S$ = Space$(4)

  GJReclen = Len(GJRec(1))

  RptFile = FreeFile
  Open "GLCMTRX.RPT" For Output As RptFile

  'ClearBox
  'QPrintRC "Searching Cash Transactions.", 11, 26, 126
  'QPrintRC "New Transactions:", 13, 29, Cnf.HiLite

  ReDim CMTransRec(1) As CMTransRecType
  CMTransRecLen = Len(CMTransRec(1))

  CMTran = FreeFile
  Open "CMTRANS.DAT" For Random Shared As CMTran Len = CMTransRecLen
  NumOfTRecs& = LOF(CMTran) \ CMTransRecLen
  Lock #CMTran

  For TCnt& = NumOfTRecs& To 1 Step -1
    Get #CMTran, TCnt&, CMTransRec(1)
    If Len(QPTrim$(CMTransRec(1).Trans2GL)) = 0 Or QPTrim$(CMTransRec(1).Trans2GL) = "N" Then
      'Store trans rec numbers and dates in array
      If CMTransRec(1).TransDate <= ThruDate% Then
       '''' If CMTransRec(1).TransDate = 0 Then Stop
        FoundCnt = FoundCnt + 1
        ReDim Preserve TranInfo(FoundCnt) As TranRecInfoType
        
        TranInfo(FoundCnt).TranDate = CMTransRec(1).TransDate
        TranInfo(FoundCnt).TranRecNo = TCnt&
      End If
    Else
      NGCnt = NGCnt + 1
    End If
    'RSet P2S$ = Str$(FoundCnt)
    'QPrintRC P2S$, 13, 47, Cnf.HiLite
    'SmallPause
    'Allow 250 Bad Dates Before Exiting
    If NGCnt >= 250 Then Exit For
  Next

  If FoundCnt = 0 Then
    FrmShowPctComp.ShowPctComp 1, 1
    Close
    'ClearBox
    'Print Chr$(7);
    Call MainLog("NO CM to Grab for " + fpDate)
    MsgBox "No Transactions Found To InterFace", vbOKOnly, "No Transactions"
    'SLEEP 4
    GoTo SendExit
  End If
  
  QSortTRec TranInfo(), FoundCnt     'sort'em by date. oldest first
  'Array (1), NumElem, Dir, StructSize, MemOff, MemSize
  GJFile = FreeFile
  Open "GLTRXED.DAT" For Random Access Read Write Shared As GJFile Len = GJReclen
  NumEdTrans = LOF(GJFile) \ GJReclen

  MCFile = FreeFile
  Open "CMMISCCD.DAT" For Random Shared As MCFile Len = Len(MCodeRec(1))

  FirstTran = 1
  ThisDate = TranInfo(1).TranDate
  WorkDate = ThisDate

  For cnt = 1 To FoundCnt
    FrmShowPctComp.ShowPctComp cnt, FoundCnt
    If ThisDate <> TranInfo(cnt).TranDate Then
      ThisDate = TranInfo(cnt).TranDate
      GoSub ProcessThisBunch
      DayCount = 0
      WorkDate = ThisDate
    End If
    If TranInfo(cnt).TranRecNo <> 0 Then
    Get #CMTran, TranInfo(cnt).TranRecNo, CMTransRec(1)
    If CMTransRec(1).TransSource = 1 Or CMTransRec(1).TransSource = 201 Then

      If DayCount = 0 Then

        For MCnt = 1 To 5
          MiscRevAmt# = (CMTransRec(1).TransRevAmt(MCnt))
          MiscRevAmt# = Round#(MiscRevAmt#)
          If MiscRevAmt# <> 0 Then
            'If There Is an Amount in Misc Rev 1-5 then get code record number
            If CMTransRec(1).TransRevAmt(MCnt + 5) <> 0 Then
              DayCount = DayCount + 1
              MiscRec#(DayCount) = CMTransRec(1).TransRevAmt(MCnt + 5)
              MiscAmt#(DayCount) = MiscRevAmt#
            End If
          End If

        Next MCnt

      Else
        For MCnt = 1 To 5
          MiscRevAmt# = (CMTransRec(1).TransRevAmt(MCnt))
          MiscRevAmt# = Round#(MiscRevAmt#)
          Do While MiscRevAmt# <> 0
            For FindCount = 1 To DayCount
              If MiscRec#(FindCount) = CMTransRec(1).TransRevAmt(MCnt + 5) Then
                MiscAmt#(FindCount) = MiscAmt#(FindCount) + MiscRevAmt#
                MiscRevAmt# = 0
                Exit Do
              End If
            Next FindCount
            DayCount = DayCount + 1
            MiscRec#(DayCount) = CMTransRec(1).TransRevAmt(MCnt + 5)
            MiscAmt#(DayCount) = MiscRevAmt#
            MiscRevAmt# = 0
          Loop
        Next MCnt
      End If
      
    End If
    End If
  Next cnt
  GoSub ProcessThisBunch
  If BadAcct > 0 Then
    Close
    Unload FrmShowPctComp
    Call MainLog("Error CM Grab not Created.")
    MsgBox "Errors Found, Interface File Not Created. Please Review Report.", vbOKOnly, "Errors"
    frmReportOpt.Show 1
    If rptopt = 1 Then
      frmGetDistMenu.PrnEditList 2
    ElseIf rptopt = 2 Then
      frmGetDistMenu.PrnEditList2 2
    End If

    KillFileD "GLTRXED.DAT"
    Exit Sub
  End If
  'transactions As interfaced
  
  For cnt = 1 To FoundCnt
  If TranInfo(cnt).TranRecNo <> 0 Then
    Get #CMTran, TranInfo(cnt).TranRecNo, CMTransRec(1)
    CMTransRec(1).Trans2GL = "Y"
    Put #CMTran, TranInfo(cnt).TranRecNo, CMTransRec(1)
   End If
  Next
  Close
  Call MainLog("Completed CM Grab " + Str$(FoundCnt) + " for " + fpDate)
  MsgBox "Transaction Grab Complete.", vbOKOnly, "Complete"
SendExit:
  Exit Sub

ProcessThisBunch:
  ' Must Combine By Date and Then Do Cash Debit Entry For Total by Fund

  If DayCount <= 0 Then Return

  FundCnt = 0   ' Set Funds Used to Zero

  For Process = 1 To DayCount
    If MiscRec#(Process) <> 0 Then
      Get #MCFile, MiscRec#(Process), MCodeRec(1)
      Acct$ = MCodeRec(1).GlAcctNumb
      Acct$ = QPStrip$(Acct$)
      Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
      AcctR = AcctFind(Acct$)
      If AcctR > 0 Then
        AcctName$ = GetAcctTitle(AcctR)
      Else
        BadAcct = BadAcct + 1
        AcctName$ = "UnDefined"
      End If
      GJRec(1).AcctRec = 0
      GJRec(1).AcctNum = QPTrim$(Acct$)
      GJRec(1).AcctName = AcctName$
      GJRec(1).TRDATE = WorkDate
      GJRec(1).Ref = Ref$
      GJRec(1).CrAmt = MiscAmt#(Process)
      GJRec(1).DrAmt = 0
      GJRec(1).EType = "C"
      GJRec(1).Desc = "FROM CASH MGMT"
      GJRec(1).LDesc = "CM-Miscellaneous"
      GJRec(1).Src = "CM"
      Put #GJFile, , GJRec(1)

      'Add Up Fund Total Here for Cash Credit Entry
      If FundCnt = 0 Then
        FundCnt = 1
        Fund$(FundCnt) = Left$(Acct$, GLFundLen%)
        FundAmt#(FundCnt) = MiscAmt#(Process)
      Else
        FoundFund = 0
        For FCnt = 1 To FundCnt
          If Fund$(FCnt) = Left$(Acct$, GLFundLen%) Then
            FoundFund = 1
            FundAmt#(FCnt) = FundAmt#(FCnt) + MiscAmt#(Process)
          End If
        Next FCnt
        If FoundFund = 0 Then
          FundCnt = FundCnt + 1
          Fund$(FundCnt) = Left$(Acct$, GLFundLen%)
          FundAmt#(FundCnt) = MiscAmt#(Process)
        End If
      End If
    End If
  Next Process

 ' Now Make Matching Debit Entries to Cash Account

  For Cash = 1 To FundCnt
    Acct$ = Fund$(Cash) + CAcct$ 'CashAcct$
    Acct$ = QPTrim$(Acct$)
    Acct$ = QPStrip$(Acct$)
    Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
    AcctR = AcctFind(Acct$)
    If AcctR > 0 Then
      AcctName$ = GetAcctTitle(AcctR)
    Else
      BadAcct = BadAcct + 1
      AcctName$ = "UnDefined"
    End If
    GJRec(1).AcctRec = 0
    GJRec(1).AcctNum = Acct$
    GJRec(1).AcctName = AcctName$
    GJRec(1).TRDATE = WorkDate
    GJRec(1).Ref = Ref$
    GJRec(1).DrAmt = FundAmt#(Cash)
    GJRec(1).CrAmt = 0
    GJRec(1).EType = "D"
    GJRec(1).Desc = "FROM CASH MGMT"
    GJRec(1).LDesc = "CM-Miscellaneous"
    GJRec(1).Src = "CM"
    Put #GJFile, , GJRec(1)
  Next Cash
BunchReturn:
  Return

End Sub
'This is new one with creation of temp util detail...
Private Sub ExtractUB(ThruDate%)
  Dim Today As String, Ref As String, Dash80 As String, P2S As String
  Dim GJReclen As Integer, RptFile As Integer, UBTransRecLen As Integer
  Dim UBTran As Integer, NumOfTRecs As Long, TCnt As Long, PageNo As Integer
  Dim FoundCnt As Long, NGCnt As Integer, GJFile As Integer
  Dim NumEdTrans As Integer, MCFile As Integer, cnt As Long
  Dim FirstTran As Long, ThisDate As Integer, WorkDate As Integer
  Dim DayCount As Integer, MCnt As Integer, MiscRevAmt As Double
  Dim FindCount As Integer, FundCnt As Integer, Process As Integer
  Dim Acct As String, AcctName As String, ThisAcct As Integer
  Dim FoundFund As Integer, FCnt As Integer, Cash As Integer
  Dim UBSetUpFileNum As Integer, UBSetUpLen As Integer
  Dim AcctMeth As String, InterfaceMethod As Integer, RevCnt As Integer
  Dim TempRev As String, NumOfRevs As Integer, BadAcct As Integer
  Dim LastTran As Long, NumPrinted As Integer, BadCAcct As String
  Dim ActT As Integer, ActPg As String, BadDAcct As String, PCnt As Long
  Dim GLUBTransRecLen As Integer, GLUBTran2 As Integer, DD As String
  Dim UBCustRecLen As Integer, UBCust As Integer
  Today$ = Date$
  
  Ref$ = "UB" + Left$(Today$, 2) + Mid$(Today$, 4, 2) + Right$(Today$, 2)
  ReDim TranInfo(1) As TranRecInfoType
  Dash80$ = String$(80, "-")
  P2S$ = Space$(4)
  ReDim GJRec1(1 To 2) As TrEditRecType
  GJReclen = Len(GJRec1(1))
  GJFile = FreeFile
  Dim GJInfo() As GJXferRecType
  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpFileNum, UBSetUpLen
  Get UBSetUpFileNum, 1, UBSetUpRec(1)
  AcctMeth$ = QPTrim$(UBSetUpRec(1).MethAcct)
  If (Len(AcctMeth$) = 0) Then
    Unload FrmShowPctComp
    MsgBox "The Utility Account Method Is Not Setup", vbOKOnly, "Invalid Setup Info"
    GoTo SendExitUB
  End If

  Select Case AcctMeth$
  Case "C"
    InterfaceMethod = 1
  Case "A"
    InterfaceMethod = 2
  Case Else
    Unload FrmShowPctComp
    GoTo SendExitUB
  End Select

  RptFile = FreeFile
  Open "UBNOTFND.RPT" For Output As RptFile
  GoSub NotFoundHeader
  
  'ShowProcessingScrn "Verifying GL Transfer Accounts"
  FrmShowPctComp.ShowPctComp 20, 100

  For RevCnt = 1 To MaxRevsCnt
    TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(RevCnt).RevName)
    If Len(TempRev$) = 0 Then
      NumOfRevs = RevCnt - 1
      Exit For
    Else
      ReDim Preserve GJInfo(1 To RevCnt) As GJXferRecType
      GJInfo(RevCnt).RevText = TempRev$
      GJInfo(RevCnt).BAcctInfo.DAcctNo = UBSetUpRec(1).BillAcct(RevCnt).DebitAcct
      GJInfo(RevCnt).BAcctInfo.CAcctNo = UBSetUpRec(1).BillAcct(RevCnt).CreditAcct
      GJInfo(RevCnt).PAcctInfo.DAcctNo = UBSetUpRec(1).PayAcct(RevCnt).DebitAcct
      GJInfo(RevCnt).PAcctInfo.CAcctNo = UBSetUpRec(1).PayAcct(RevCnt).CreditAcct
      If UBSetUpRec(1).Revenues(RevCnt).UseDep = "Y" Then
        GJInfo(RevCnt).DAcctInfo.DAcctNo = UBSetUpRec(1).DepAcct(RevCnt).DebitAcct
        GJInfo(RevCnt).DAcctInfo.CAcctNo = UBSetUpRec(1).DepAcct(RevCnt).CreditAcct
      End If
    End If
  Next
  FrmShowPctComp.ShowPctComp 75, 100

  'check to see if they are valid GL accounts
  GoSub ValidateGLAccounts
  FrmShowPctComp.ShowPctComp 80, 100
  If BadAcct Then
    Unload FrmShowPctComp
    GoTo SendExitUB
  End If

'  ClearBox
'  QPrintRC "Searching Cash Transactions.", 11, 26, 126
'  QPrintRC "New Transactions:", 13, 29, Cnf.HiLite
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  UBCust = FreeFile
  Open "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
 
  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))
  UBTran = FreeFile
  Open "UBTRANS.DAT" For Random Shared As UBTran Len = UBTransRecLen
  NumOfTRecs& = LOF(UBTran) \ UBTransRecLen
  
  For TCnt& = NumOfTRecs& To 1 Step -1    '1 To NumOfTRecs&
    Get #UBTran, TCnt&, UBTransRec(1)
    If UBTransRec(1).CustAcctNo > 0 Then    'so don't get whacked trans
    If Len(QPTrim$(UBTransRec(1).Posted2GL)) = 0 Or QPTrim$(UBTransRec(1).Posted2GL) = "N" Then
     If UBTransRec(1).TransDate <= ThruDate% Then
        If UBTransRec(1).TransDate <= 0 Then 'Exit For
          FrmShowPctComp.ShowPctComp 100, 100
          Close
          GoSub Getout
        End If
      'Store trans rec numbers and dates in array
'          UBTransRec(1).TransDate = 1
'        End If
        FoundCnt = FoundCnt + 1
        ReDim Preserve TranInfo(FoundCnt) As TranRecInfoType
        TranInfo(FoundCnt).TranDate = UBTransRec(1).TransDate
        TranInfo(FoundCnt).TranRecNo = TCnt&
        'If FoundCnt = 30000 Then Exit For
        
      End If
    Else
      NGCnt = NGCnt + 1
    End If
    'RSet P2S$ = Str$(FoundCnt)
    'QPrintRC P2S$, 13, 47, Cnf.HiLite
    'SmallPause
    If NGCnt >= 2500 Then
      FrmShowPctComp.ShowPctComp 1, 1
      Exit For
    End If
    End If
    'FrmShowPctComp.ShowPctComp TCnt&, NumOfTRecs&
  Next
  'FrmShowPctComp.ShowPctComp 1, 1
  FrmShowPctComp.Label1 = "Sorting Interface Trans File"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me

  If FoundCnt = 0 Then
    Unload FrmShowPctComp
    Close
    Call MainLog("No Trans UB Grab for " + fpDate)
    MsgBox "No Transactions Found To Interface.", vbOKOnly, "No Trans"
    GoTo SendExitUB
  End If
  FrmShowPctComp.ShowPctComp 25, 100
  
  QSortTRec TranInfo(), FoundCnt
  'sort'em by date. oldest first
  'Array(1), NumElem, Dir, StructSize, MemOff, MemSize
  FrmShowPctComp.ShowPctComp 50, 100
  GJFile = FreeFile
  Open "GLTRXED.DAT" For Random Access Read Write Shared As GJFile Len = GJReclen
  ReDim GLUBTran(1) As GLUBTempRecType
  GLUBTransRecLen = Len(GLUBTran(1))
  GLUBTran2 = FreeFile
  Open "GLUBTran.DAT" For Random Shared As GLUBTran2 Len = GLUBTransRecLen
  'Get a unique name for field ?????
      
  DD$ = Mid$(Date$, 1, 5) + Right$(Str(Timer), 3)
  FirstTran = 1
  ThisDate = TranInfo(1).TranDate
  For cnt = 1 To FoundCnt
    If ThisDate <> TranInfo(cnt).TranDate Then
      ThisDate = TranInfo(cnt).TranDate
      LastTran = cnt - 1
      GoSub ProcessThisBunchUB
      FirstTran = cnt
    End If
    FrmShowPctComp.ShowPctComp cnt, FoundCnt
  Next
  FrmShowPctComp.Label1 = "Tag Interface Transactions"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me

  LastTran = FoundCnt
  GoSub ProcessThisBunchUB

  'transactions as interfaced
  For cnt = 1 To FoundCnt
    FrmShowPctComp.ShowPctComp cnt, FoundCnt
    Get #UBTran, TranInfo(cnt).TranRecNo, UBTransRec(1)
    UBTransRec(1).Posted2GL = "Y"
    Put #UBTran, TranInfo(cnt).TranRecNo, UBTransRec(1)
  Next
  Close
  Call MainLog("Completed UB Grab " + Str$(FoundCnt) + " for " + fpDate)
  MsgBox "Transaction Grab Complete.", vbOKOnly, "Complete"
  'SLEEP 2
SendExitUB:
  Exit Sub
NotFoundHeader:
  PageNo = PageNo + 1
  Print #RptFile, "Utility Billing GL Transfer Invalid Account Listing."; Tab(70); "Page:"; PageNo

  Print #RptFile, QPTrim(GLUserName$)
  Print #RptFile, "Report Date: "; Date$
  Print #RptFile, "Revenue           Acct. Type              Debit Acct."
  Print #RptFile, Dash80$
  NumPrinted = 0
  Return

PrintBadAcct:
  If Len(QPTrim$(BadCAcct$)) = 0 Then
    BadCAcct$ = "Undefined"
  End If

  Print #RptFile, GJInfo(RevCnt).RevText;

  Select Case ActT
  Case 1
    ActPg$ = "Billing"
  Case 2
    ActPg$ = "Payment"
  Case 3
    ActPg$ = "Deposit"
  End Select
  Print #RptFile, Tab(22); ActPg$;
  Print #RptFile, Tab(43); BadDAcct$; Tab(64); BadCAcct$
  Return

ProcessThisBunchUB:
  For RevCnt = 1 To NumOfRevs
    GJInfo(RevCnt).BAcctInfo.CreditAmt = 0
    GJInfo(RevCnt).BAcctInfo.DebitAmt = 0
    GJInfo(RevCnt).PAcctInfo.CreditAmt = 0
    GJInfo(RevCnt).PAcctInfo.DebitAmt = 0
    GJInfo(RevCnt).DAcctInfo.CreditAmt = 0
    GJInfo(RevCnt).DAcctInfo.DebitAmt = 0
  Next

  For PCnt = FirstTran To LastTran
    If PCnt = FirstTran Then
      WorkDate = TranInfo(PCnt).TranDate
    End If
    Get #UBTran, TranInfo(PCnt).TranRecNo, UBTransRec(1)

    Select Case InterfaceMethod
    Case 1      'Cash Central
      Select Case UBTransRec(1).TransType
      Case TranUtilityBill      ' 1=Utility bill
        'no action
      Case TranLateCharge       ' 2=late charge
        'no action
      Case TranReconnectFee     ' 3=reconnect fee
        'no action
      Case TranBillPayment      ' 4=Bill Payment
        For RevCnt = 1 To NumOfRevs
          GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))

          GJInfo(RevCnt).PAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).PAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))

        Next
        GoSub UpdateTempFile
      Case TranAppliedDeposit   ' 5=Applied Deposit
        'no action
      Case TranPenaltyCharge    ' 6=Penalty Charge
        'no action
      Case TranDepositPayment   ' 7=Deposit Payment
        For RevCnt = 1 To NumOfRevs
          GJInfo(RevCnt).DAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).DAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))

          GJInfo(RevCnt).DAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).DAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))

        Next
        GoSub UpdateTempFile
      Case TranDraftPayment     ' 8=Draft Payment
        For RevCnt = 1 To NumOfRevs
          GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))

          GJInfo(RevCnt).PAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).DAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))

        Next
        GoSub UpdateTempFile
      Case TranRefundDeposit    ' 9=Refund Deposit
        'no action
      Case TranBeginBalance     '10=Beginning Balance
        'no action
      Case TranUpwardAdjustment '11=Upward Adjustments
        'no action
      Case TranDownwardAdjustment  '12=Downward Adjustments
        'no action
      Case TranOverPayAdjustment   '33=OverPayment Adjustments
        For RevCnt = 1 To NumOfRevs
          GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
          GJInfo(RevCnt).PAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).PAcctInfo.DebitAmt - UBTransRec(1).RevAmt(RevCnt))
        Next
        GoSub UpdateTempFile
      Case TranDepCreditRemoval    '37=Deposit Credit Removal Not to Interface w/GL
        'No Action !!!
      Case TranDepPaymentVoid         ' 39=Deposit Void
        For RevCnt = 1 To NumOfRevs
          GJInfo(RevCnt).DAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).DAcctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
          GJInfo(RevCnt).DAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).DAcctInfo.DebitAmt - UBTransRec(1).RevAmt(RevCnt))
        Next
        GoSub UpdateTempFile
      End Select
    
    Case 2      'Accrual
      Select Case UBTransRec(1).TransType
      Case TranUtilityBill      ' 1=Utility bill
        For RevCnt = 1 To NumOfRevs
          GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
          GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
        Next
        GoSub UpdateTempFile
      Case TranLateCharge       ' 2=late charge
        For RevCnt = 1 To NumOfRevs
          GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
          GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
        Next
        GoSub UpdateTempFile
      Case TranReconnectFee     ' 3=reconnect fee
        For RevCnt = 1 To NumOfRevs
          GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
          GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
        Next
        GoSub UpdateTempFile
      Case TranBillPayment      ' 4=Bill Payment
        For RevCnt = 1 To NumOfRevs
          GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
          GJInfo(RevCnt).PAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).PAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
        Next
        GoSub UpdateTempFile
      Case TranAppliedDeposit   ' 5=Applied Deposit
        'no action
        'FOR RevCnt = 1 TO NumOfRevs
        '  GJInfo(RevCnt).dacctInfo.CreditAmt = Round#(GJInfo(RevCnt).dacctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
        '  GJInfo(RevCnt).pAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).pAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
        'NEXT

      Case TranPenaltyCharge    ' 6=Penalty Charge
        For RevCnt = 1 To NumOfRevs
          GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
          GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
        Next
        GoSub UpdateTempFile
      Case TranDepositPayment   ' 7=Deposit Payment
        For RevCnt = 1 To NumOfRevs
          GJInfo(RevCnt).DAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).DAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
          GJInfo(RevCnt).DAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).DAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
        Next
        GoSub UpdateTempFile
      Case TranDraftPayment     ' 8=Draft Payment
        For RevCnt = 1 To NumOfRevs
          GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
          GJInfo(RevCnt).PAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).PAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
        Next
        GoSub UpdateTempFile
      Case TranRefundDeposit    ' 9=Refund Deposit
        'no action
        '  FOR RevCnt = 1 TO NumOfRevs
        '    GJInfo(RevCnt).dacctInfo.CreditAmt = Round#(GJInfo(RevCnt).dacctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
        '    GJInfo(RevCnt).dacctInfo.DebitAmt = Round#(GJInfo(RevCnt).dacctInfo.DebitAmt - UBTransRec(1).RevAmt(RevCnt))
        '  NEXT
      Case TranBeginBalance     '10=Beginning Balance
        'no action
      Case TranUpwardAdjustment '11=Upward Adjustments
        For RevCnt = 1 To NumOfRevs
          GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
          GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
        Next
        GoSub UpdateTempFile
      Case TranDownwardAdjustment               '12=Downward Adjustments
        For RevCnt = 1 To NumOfRevs
          GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
          GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt - UBTransRec(1).RevAmt(RevCnt))
        Next
        GoSub UpdateTempFile
      Case TranOverPayAdjustment   '33=OverPayment Adjustments
        For RevCnt = 1 To NumOfRevs
          GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
          GJInfo(RevCnt).PAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).PAcctInfo.DebitAmt - UBTransRec(1).RevAmt(RevCnt))
        Next
        GoSub UpdateTempFile
      Case TranDepCreditRemoval    '37=Deposit Credit Removal Not to Interface w/GL
        'No Action !!!
      Case TranDepPaymentVoid         ' 39=Deposit Void
        For RevCnt = 1 To NumOfRevs
          GJInfo(RevCnt).DAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).DAcctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
          GJInfo(RevCnt).DAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).DAcctInfo.DebitAmt - UBTransRec(1).RevAmt(RevCnt))
        Next
        GoSub UpdateTempFile
      End Select
    End Select
    'SmallPause

  Next


  'NOTE: Journal Rec 1 is the credit, Rec 2 is the debit
  For RevCnt = 1 To NumOfRevs
    ReDim GJRec1(1 To 2) As TrEditRecType
    If GJInfo(RevCnt).BAcctInfo.CreditAmt <> 0 Then
      GJRec1(1).AcctRec = GJInfo(RevCnt).BAcctInfo.CRecNo
      GJRec1(1).AcctNum = GJInfo(RevCnt).BAcctInfo.CAcctNo
      GJRec1(1).AcctName = GJInfo(RevCnt).BAcctInfo.CTitle
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = GJInfo(RevCnt).BAcctInfo.CreditAmt
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FROM UTILITIES"
      GJRec1(1).LDesc = "UB Interface"
      GJRec1(1).Src = "UB"
      Put #GJFile, , GJRec1(1)
    End If
    If GJInfo(RevCnt).BAcctInfo.DebitAmt <> 0 Then
      GJRec1(2).AcctRec = GJInfo(RevCnt).BAcctInfo.DRecNo
      GJRec1(2).AcctNum = GJInfo(RevCnt).BAcctInfo.DAcctNo
      GJRec1(2).AcctName = GJInfo(RevCnt).BAcctInfo.DTitle
      GJRec1(2).TRDATE = WorkDate
      GJRec1(2).Ref = Ref$
      GJRec1(2).DrAmt = GJInfo(RevCnt).BAcctInfo.DebitAmt
      GJRec1(2).EType = "D"
      GJRec1(2).Desc = "FROM UTILITIES"
      GJRec1(2).LDesc = "UB Interface"
      GJRec1(2).Src = "UB"
      Put #GJFile, , GJRec1(2)
    End If
  Next

  For RevCnt = 1 To NumOfRevs
    ReDim GJRec1(1 To 2) As TrEditRecType
    If GJInfo(RevCnt).PAcctInfo.CreditAmt <> 0 Then
      GJRec1(1).AcctRec = GJInfo(RevCnt).PAcctInfo.CRecNo
      GJRec1(1).AcctNum = GJInfo(RevCnt).PAcctInfo.CAcctNo
      GJRec1(1).AcctName = GJInfo(RevCnt).PAcctInfo.CTitle
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = GJInfo(RevCnt).PAcctInfo.CreditAmt
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FROM UTILITIES"
      GJRec1(1).LDesc = "UB Interface"
      GJRec1(1).Src = "UB"
      Put #GJFile, , GJRec1(1)
    End If
    If GJInfo(RevCnt).PAcctInfo.DebitAmt <> 0 Then
      GJRec1(2).AcctRec = GJInfo(RevCnt).PAcctInfo.DRecNo
      GJRec1(2).AcctNum = GJInfo(RevCnt).PAcctInfo.DAcctNo
      GJRec1(2).AcctName = GJInfo(RevCnt).PAcctInfo.DTitle
      GJRec1(2).TRDATE = WorkDate
      GJRec1(2).Ref = Ref$
      GJRec1(2).DrAmt = GJInfo(RevCnt).PAcctInfo.DebitAmt
      GJRec1(2).EType = "D"
      GJRec1(2).Desc = "FROM UTILITIES"
      GJRec1(2).LDesc = "UB Interface"
      GJRec1(2).Src = "UB"
      Put #GJFile, , GJRec1(2)
    End If
  Next

  For RevCnt = 1 To NumOfRevs
    ReDim GJRec1(1 To 2) As TrEditRecType
    If GJInfo(RevCnt).DAcctInfo.CreditAmt <> 0 Then
      GJRec1(1).AcctRec = GJInfo(RevCnt).DAcctInfo.CRecNo
      GJRec1(1).AcctNum = GJInfo(RevCnt).DAcctInfo.CAcctNo
      GJRec1(1).AcctName = GJInfo(RevCnt).DAcctInfo.CTitle
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = GJInfo(RevCnt).DAcctInfo.CreditAmt
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FROM UTILITIES"
      GJRec1(1).LDesc = "UB Interface"
      GJRec1(1).Src = "UB"
      Put #GJFile, , GJRec1(1)
    End If
    If GJInfo(RevCnt).DAcctInfo.DebitAmt <> 0 Then
      GJRec1(2).AcctRec = GJInfo(RevCnt).DAcctInfo.DRecNo
      GJRec1(2).AcctNum = GJInfo(RevCnt).DAcctInfo.DAcctNo
      GJRec1(2).AcctName = GJInfo(RevCnt).DAcctInfo.DTitle
      GJRec1(2).TRDATE = WorkDate
      GJRec1(2).Ref = Ref$
      GJRec1(2).DrAmt = GJInfo(RevCnt).DAcctInfo.DebitAmt
      GJRec1(2).EType = "D"
      GJRec1(2).Desc = "FROM UTILITIES"
      GJRec1(2).LDesc = "UB Interface"
      GJRec1(2).Src = "UB"
      Put #GJFile, , GJRec1(2)
    End If
  Next

UBBunchReturn:
  Return

ValidateGLAccounts:
  BadAcct = False
  For RevCnt = 1 To NumOfRevs
    'Billing Accounts
    If InterfaceMethod = 2 Then
      'NOTE: We Only check billing accounts if Accural method
      ActT = 1
      
      Acct$ = GJInfo(RevCnt).BAcctInfo.DAcctNo
      Acct$ = QPStrip$(Acct$)
      Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
      Acct$ = QPTrim$(Acct$)
       'AcctR = AcctFind(Acct$)
      ThisAcct = AcctFind(Acct$) 'AcctFind(GJInfo(RevCnt).BAcctInfo.DAcctNo)
      If ThisAcct <= 0 Then
        BadDAcct$ = GJInfo(RevCnt).BAcctInfo.DAcctNo
        BadAcct = True
      Else
        GJInfo(RevCnt).BAcctInfo.DRecNo = ThisAcct
        GJInfo(RevCnt).BAcctInfo.DAcctNo = Acct$
        GJInfo(RevCnt).BAcctInfo.DTitle = GetAcctTitle$(ThisAcct)
        BadDAcct$ = "     OK"
      End If
      
      Acct$ = GJInfo(RevCnt).BAcctInfo.CAcctNo
      Acct$ = QPStrip$(Acct$)
      Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
      Acct$ = QPTrim$(Acct$)
      ThisAcct = AcctFind(Acct$)
      'ThisAcct = AcctFind(GJInfo(RevCnt).BAcctInfo.CAcctNo)
      If ThisAcct <= 0 Then
        BadCAcct$ = GJInfo(RevCnt).BAcctInfo.CAcctNo
        BadAcct = True
      Else
        GJInfo(RevCnt).BAcctInfo.CRecNo = ThisAcct
        GJInfo(RevCnt).BAcctInfo.CAcctNo = Acct$
        GJInfo(RevCnt).BAcctInfo.CTitle = GetAcctTitle$(ThisAcct)
        BadCAcct$ = "     OK"
      End If
      GoSub PrintBadAcct
    End If

    'Payment Accounts
    ActT = 2
      Acct$ = GJInfo(RevCnt).PAcctInfo.DAcctNo
      Acct$ = QPStrip$(Acct$)
      Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
      Acct$ = QPTrim$(Acct$)
      ThisAcct = AcctFind(Acct$)

    'ThisAcct = AcctFind(GJInfo(RevCnt).PAcctInfo.DAcctNo)
    If ThisAcct <= 0 Then
      BadDAcct$ = GJInfo(RevCnt).PAcctInfo.DAcctNo
      BadAcct = True
    Else
      GJInfo(RevCnt).PAcctInfo.DRecNo = ThisAcct
      GJInfo(RevCnt).PAcctInfo.DAcctNo = Acct$
      GJInfo(RevCnt).PAcctInfo.DTitle = GetAcctTitle$(ThisAcct)
      BadDAcct$ = "     OK"
    End If
    
      Acct$ = GJInfo(RevCnt).PAcctInfo.CAcctNo
      Acct$ = QPStrip$(Acct$)
      Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
      Acct$ = QPTrim$(Acct$)
      ThisAcct = AcctFind(Acct$)

    'ThisAcct = AcctFind(GJInfo(RevCnt).PAcctInfo.CAcctNo)
    If ThisAcct <= 0 Then
      BadCAcct$ = GJInfo(RevCnt).PAcctInfo.CAcctNo
      BadAcct = True
    Else
      GJInfo(RevCnt).PAcctInfo.CRecNo = ThisAcct
       GJInfo(RevCnt).PAcctInfo.CAcctNo = Acct$
      GJInfo(RevCnt).PAcctInfo.CTitle = GetAcctTitle$(ThisAcct)
      BadCAcct$ = "     OK"
    End If
    GoSub PrintBadAcct

    'Deposit Accounts
    ActT = 3
    If UBSetUpRec(1).Revenues(RevCnt).UseDep = "Y" Then
    
      Acct$ = GJInfo(RevCnt).DAcctInfo.DAcctNo
      Acct$ = QPStrip$(Acct$)
      Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
      Acct$ = QPTrim$(Acct$)
      ThisAcct = AcctFind(Acct$)

      'ThisAcct = AcctFind(GJInfo(RevCnt).DAcctInfo.DAcctNo)
      If ThisAcct <= 0 Then
        BadDAcct$ = GJInfo(RevCnt).DAcctInfo.DAcctNo
        BadAcct = True
      Else
        GJInfo(RevCnt).DAcctInfo.DRecNo = ThisAcct
        GJInfo(RevCnt).DAcctInfo.DAcctNo = Acct$
        GJInfo(RevCnt).DAcctInfo.DTitle = GetAcctTitle$(ThisAcct)
        BadDAcct$ = "     OK"
      End If
    Else
      BadDAcct$ = "    N/A"
    End If
    If UBSetUpRec(1).Revenues(RevCnt).UseDep = "Y" Then
      Acct$ = GJInfo(RevCnt).DAcctInfo.CAcctNo
      Acct$ = QPStrip$(Acct$)
      Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
      Acct$ = QPTrim$(Acct$)
      ThisAcct = AcctFind(Acct$)
 
      'ThisAcct = AcctFind(GJInfo(RevCnt).DAcctInfo.CAcctNo)
      If ThisAcct <= 0 Then
        BadCAcct$ = GJInfo(RevCnt).DAcctInfo.CAcctNo
        BadAcct = True
      Else
        GJInfo(RevCnt).DAcctInfo.CRecNo = ThisAcct
        GJInfo(RevCnt).DAcctInfo.CAcctNo = Acct$
        GJInfo(RevCnt).DAcctInfo.CTitle = GetAcctTitle$(ThisAcct)
        BadCAcct$ = "     OK"
      End If
    Else
      BadCAcct$ = "    N/A"
    End If
    GoSub PrintBadAcct
  Next
  Close RptFile

  If BadAcct Then
    Unload FrmShowPctComp
    MsgBox "Invalid Account(s)Found, Interface File Was Not Created.", vbOKOnly, "Invalid"
    Call MainLog("UB Grab - NOgo Invalid Accts.")
    ViewPrint "UBNOTFND.RPT", "GL Transfer Invalid Account List."
  End If
  Kill "UBNOTFND.RPT"
Return
UpdateTempFile:
   GLUBTran(1).Grabbatch = DD$        'As String * 8  'this will be date and num of batch that day
   GLUBTran(1).TransDate = UBTransRec(1).TransDate
   GLUBTran(1).TransType = UBTransRec(1).TransType
   GLUBTran(1).TransDesc = UBTransRec(1).TransDesc
   GLUBTran(1).Transamt = UBTransRec(1).Transamt
   For RevCnt = 1 To 15
    GLUBTran(1).RevAmt(RevCnt) = UBTransRec(1).RevAmt(RevCnt)
    GLUBTran(1).TaxAmt(RevCnt) = UBTransRec(1).TaxAmt(RevCnt)
   Next
   GLUBTran(1).CustStatus = UBTransRec(1).CustStatus
   GLUBTran(1).CustAcctNo = UBTransRec(1).CustAcctNo
   Get UBCust, UBTransRec(1).CustAcctNo, UBCustRec(1)
   GLUBTran(1).CustName = UBCustRec(1).CustName
   GLUBTran(1).OperatorNumber = UBTransRec(1).OperatorNumber
   Put #GLUBTran2, , GLUBTran(1)
Return
Getout:
  Call MainLog("EXIT Grab via error with invalid dates in UB.")
  'frmCitiCancel.Label1.Caption = "Invalid Dates/Utility Trans Call Software Support."
  'frmCitiCancel.Show 1
  MsgBox "Invalid Dates in Utility Trans, Please Call Software Support.", vbCritical, "Warning!!!!"
  frmGetDistMenu.Show
  Unload frmGrabTrans
End Sub

'Business License Interface
Private Sub ExtractBL(ThruDate)
  Dim Today As String, Ref As String, Dash80 As String, P2S As String
  Dim GJReclen As Integer, RptFile As Integer, ARTransRecLen As Integer
  Dim ARTransFile As Integer, NumOfTRecs As Long, TCnt As Long
  Dim FoundCnt As Integer, NGCnt As Integer, GJFile As Integer
  Dim NumEdTrans As Integer, MCFile As Integer, cnt As Integer
  Dim FirstTran As Integer, ThisDate As Integer, WorkDate As Integer
  Dim DayCount As Integer, MCnt As Integer, MRevAmt As Double
  Dim FindCount As Integer, FundCnt As Integer, Process As Integer
  Dim Acct As String, AcctName As String, AcctR As Integer, CshAcct As String
  Dim FoundFund As Integer, FCnt As Integer, Cash As Integer
  Dim ARCatcodereclen As Integer, ARCatFile As Integer, ACMeth As String
  Dim NumOFARCatRecs As Integer, MiddleRec As Integer, CatCodeRecord As Integer
  Dim PCnum As Long, PRnum As Long, PAnum As Long, PRevAmt As Double
  Dim PCAcct As String, PRAcct As String, PAAcct As String
  Dim TownRecHandle As Integer, BadAcct As Long
  BadAcct = 0
  ACMeth$ = ""
  Today$ = Date$
  Ref$ = "BL" + Left$(Today$, 2) + Mid$(Today$, 4, 2) + Right$(Today$, 2)
  Dim TownRec As TownSetUpType
  Dim TownRecLen As Integer
  If Exist("artownsu.dat") Then
    TownRecLen = Len(TownRec)
    TownRecHandle = FreeFile
    Open "arTownsu.dat" For Random Shared As TownRecHandle Len = TownRecLen
    Get TownRecHandle, 1, TownRec
    ACMeth$ = QPTrim(TownRec.AcctMeth)
    PCnum = TownRec.PENCASHACCT
    PAnum = TownRec.PENRECGLNUM
    PRnum = TownRec.PENREVGLNUM
    Close TownRecHandle
  Else
    Unload FrmShowPctComp
    MsgBox "Business License Information NOT Found.", vbOKOnly, "Error"
    Exit Sub
  End If
  If ACMeth$ = "N" Then
    Unload FrmShowPctComp
    MsgBox "Business License NOT Set With Accounting Type.", vbOKOnly, "Error"
    Exit Sub
  Else
    If PCnum > 0 Then PCAcct$ = GetAcctNum(PCnum)
    If PAnum > 0 Then PAAcct$ = GetAcctNum(PAnum)
    If PRnum > 0 Then PRAcct$ = GetAcctNum(PRnum)
  End If
  ReDim TranInfo(1) As TranRecInfoType
  Dim Rec#(5000), DMRevAmt#(5000), DPRevAmt#(5000), ttype%(5000), Fund$(500), FundAmt#(500)

  Dash80$ = String$(80, "-")
  P2S$ = Space$(4)

  GJReclen = Len(GJRec(1))

  RptFile = FreeFile
  Open "GLCMTRX.RPT" For Output As RptFile

  'QPrintRC "Searching Cash Transactions.", 11, 26, 126
  'QPrintRC "New Transactions:", 13, 29, Cnf.HiLite

  ReDim ARTransRec(1) As ARTransRecType
  ARTransRecLen = Len(ARTransRec(1))
  If ARTransRecLen <> 252 Then
    Close
    Unload FrmShowPctComp
    MsgBox "Business License Information NOT Correct Format.", vbOKOnly, "Error"
    Exit Sub
  End If
  ARTransFile = FreeFile
  Open "ARTrans.DAT" For Random Access Read Write Shared As ARTransFile Len = ARTransRecLen
  NumOfTRecs& = LOF(ARTransFile) \ ARTransRecLen
  Lock #ARTransFile
  FrmShowPctComp.ShowPctComp 10, 100
  For TCnt& = NumOfTRecs& To 1 Step -1
    Get #ARTransFile, TCnt&, ARTransRec(1)
    If ARTransRec(1).TransDate <= ThruDate Then
      If Len(QPTrim$(ARTransRec(1).Posted2GL)) = 0 Or QPTrim$(ARTransRec(1).Posted2GL) = "N" Then
        'Store trans rec numbers and dates in array
        FoundCnt = FoundCnt + 1
        ReDim Preserve TranInfo(FoundCnt) As TranRecInfoType
        TranInfo(FoundCnt).TranDate = ARTransRec(1).TransDate
        TranInfo(FoundCnt).TranRecNo = TCnt&
     ' ELSE
     '   NGCnt = NGCnt + 1
      End If
    End If
    RSet P2S$ = Str$(FoundCnt)
    'QPrintRC P2S$, 13, 47, Cnf.HiLite
    'SmallPause
    'Allow 500 Bad Entries Before Exiting
    'IF NGCnt >= 500 THEN EXIT FOR
  Next
  '
  If FoundCnt = 0 Then
    FrmShowPctComp.ShowPctComp 1, 1
    Close
    'ClearBox
    Print Chr$(7);
    Call MainLog("No BL to Grab for " + fpDate)
    MsgBox "No Transactions Found To InterFace", vbOKOnly, "No Trans"
    'SLEEP 4
    GoTo BLSendExit
  End If

  'Array (1), NumElem, Dir, StructSize, MemOff, MemSize
  SortTRec TranInfo(), FoundCnt    'sort'em by date. oldest first

  'Open GL InterFace File
  GJFile = FreeFile
  Open "GLTRXED.DAT" For Random Access Read Write Shared As GJFile Len = GJReclen
  NumEdTrans = LOF(GJFile) \ GJReclen
  Lock #GJFile

  'OPEN Business License CODES Here

  ReDim ARCatCodeRec(1) As ARNewCatCodeRecType

  ARCatcodereclen = Len(ARCatCodeRec(1))
  ARCatFile = FreeFile
  Open "ARCODE.DAT" For Random Access Read Write Shared As ARCatFile Len = ARCatcodereclen
  NumOFARCatRecs = LOF(ARCatFile) \ ARCatcodereclen

  If NumOFARCatRecs > 1 Then
    MiddleRec = NumOFARCatRecs \ 2
  End If

  FirstTran = 1
  ThisDate = TranInfo(1).TranDate
  WorkDate = ThisDate
  FrmShowPctComp.ShowPctComp 30, 100
  For cnt = 1 To FoundCnt
    
    If ThisDate <> TranInfo(cnt).TranDate Then
      ThisDate = TranInfo(cnt).TranDate
      GoSub BLProcessThisBunch
      DayCount = 0
      WorkDate = ThisDate
    End If

    Get #ARTransFile, TranInfo(cnt).TranRecNo, ARTransRec(1)

    'Find Catagory Record Number So We Can Pull G/L Revenue Account
    CatCodeRecord = 0
    CatCodeRecord = ARTransRec(1).CatCodeRec1

    If CatCodeRecord = 0 Then
      CatCodeRecord = MiddleRec
    End If

    If CatCodeRecord > 0 Then
      Get ARCatFile, CatCodeRecord, ARCatCodeRec(1)
      If DayCount = 0 Then
        Select Case ACMeth$
         Case "A"
          PRevAmt# = Round#(ARTransRec(1).PenAmt)
          MRevAmt# = Round#(ARTransRec(1).TransAmount - ARTransRec(1).PenAmt)
         Case "C"
          Select Case ARTransRec(1).TransType
           Case 2, 13
             PRevAmt# = Round#(ARTransRec(1).PenAmt)
             MRevAmt# = Round#(ARTransRec(1).TransAmount - ARTransRec(1).PenAmt)
           Case Else
             PRevAmt# = 0
             MRevAmt# = 0
           End Select
          Case Else
         End Select
         MRevAmt# = Round#(MRevAmt#)
        If ARTransRec(1).TransAmount <> 0 Then
          'If There Is an Amount get catagory code record
          DayCount = DayCount + 1
          Rec#(DayCount) = CatCodeRecord
          DMRevAmt#(DayCount) = MRevAmt#
          DPRevAmt#(DayCount) = PRevAmt#
          ttype%(DayCount) = ARTransRec(1).TransType
         End If
      Else
        Select Case ACMeth$
         Case "A"
          PRevAmt# = Round#(ARTransRec(1).PenAmt)
          MRevAmt# = Round#(ARTransRec(1).TransAmount - ARTransRec(1).PenAmt)
         Case "C"
          Select Case ARTransRec(1).TransType
           Case 2, 13
             PRevAmt# = Round#(ARTransRec(1).PenAmt)
             MRevAmt# = Round#(ARTransRec(1).TransAmount - ARTransRec(1).PenAmt)
           Case Else
             PRevAmt# = 0
             MRevAmt# = 0
           End Select
          Case Else
         End Select
        MRevAmt# = Round#(MRevAmt#)
        If ARTransRec(1).TransAmount <> 0 Then
'          For FindCount = 1 To DayCount
'            If Rec#(FindCount) = CatCodeRecord Then
'              DMRevAmt#(FindCount) = Round#(DMRevAmt#(FindCount) + MRevAmt#)
'              'MRevAmt# = 0
'              DPRevAmt#(FindCount) = Round#(DPRevAmt#(FindCount) + PRevAmt#)
'              'PRevAmt# = 0
'              Exit For
'            End If
'          Next FindCount
         
          DayCount = DayCount + 1
          Rec#(DayCount) = CatCodeRecord
          DMRevAmt#(DayCount) = MRevAmt#
          DPRevAmt#(DayCount) = PRevAmt#
          ttype%(DayCount) = ARTransRec(1).TransType
          MRevAmt# = 0
          PRevAmt# = 0
        End If
      End If
      

    End If

  Next cnt
  FrmShowPctComp.ShowPctComp 60, 100
  GoSub BLProcessThisBunch
  If BadAcct > 0 Then
    Close
    Unload FrmShowPctComp
    Call MainLog("Error BL Grab not Created.")
    MsgBox "Errors Found, Interface File Not Created. Please Review Report.", vbOKOnly, "Errors"
    frmReportOpt.Show 1
    If rptopt = 1 Then
      frmGetDistMenu.PrnEditList 2
    ElseIf rptopt = 2 Then
      frmGetDistMenu.PrnEditList2 2
    End If
    KillFileD "GLTRXED.DAT"
    Exit Sub
  End If

  'Mark Transactions as interfaced
  For cnt = 1 To FoundCnt
    Get #ARTransFile, TranInfo(cnt).TranRecNo, ARTransRec(1)
    ARTransRec(1).Posted2GL = "Y"
    Put #ARTransFile, TranInfo(cnt).TranRecNo, ARTransRec(1)
  Next
  Close
  FrmShowPctComp.ShowPctComp 100, 100
  Call MainLog("Completed BL Grab " + Str$(FoundCnt) + " for " + fpDate)
  MsgBox "Transaction Grab Complete.", vbOKOnly, "Complete"
BLSendExit:
  Exit Sub


BLProcessThisBunch:
  ' Must Combine By Date and Then Do Cash Debit Entry For Total by Fund

  If DayCount <= 0 Then Return

  FundCnt = 0   ' Set Funds Used to Zero

  'Process Payments First Type=2

  For Process = 1 To DayCount
    If ttype%(Process) = 2 Then
      If DMRevAmt#(Process) > 0 Then
        Get #ARCatFile, Rec#(Process), ARCatCodeRec(1)
        If ACMeth$ = "A" Then
          AcctR = ARCatCodeRec(1).ARGLACCT
          Acct$ = GetAcctNum(AcctR)
        Else
          AcctR = ARCatCodeRec(1).REVGLNUM
          Acct$ = GetAcctNum(AcctR)
        End If
        If AcctR > 0 Then
          AcctName$ = GetAcctTitle(AcctR)
        Else
          BadAcct = BadAcct + 1
          AcctName$ = "UnDefined"
        End If
        GJRec(1).AcctRec = 0
        GJRec(1).AcctNum = QPTrim$(Acct$)
        GJRec(1).AcctName = AcctName$
        GJRec(1).TRDATE = WorkDate
        GJRec(1).Ref = Ref$
        GJRec(1).CrAmt = DMRevAmt#(Process)
        GJRec(1).DrAmt = 0
        GJRec(1).EType = "C"
        GJRec(1).Desc = "FROM BL"
        GJRec(1).LDesc = "BL Interface"
        GJRec(1).Src = "BL"
        Put #GJFile, , GJRec(1)
        'Now Make Matching Debit Entries to Cash Account
        AcctR = ARCatCodeRec(1).CashAcct
        Acct$ = GetAcctNum(AcctR)
        Acct$ = QPTrim$(Acct$)
        If AcctR > 0 Then
          AcctName$ = GetAcctTitle(AcctR)
        Else
          BadAcct = BadAcct + 1
          AcctName$ = "UnDefined"
        End If
        GJRec(1).AcctRec = 0
        GJRec(1).AcctNum = Acct$
        GJRec(1).AcctName = AcctName$
        GJRec(1).TRDATE = WorkDate
        GJRec(1).Ref = Ref$
        GJRec(1).DrAmt = DMRevAmt#(Process)
        GJRec(1).CrAmt = 0
        GJRec(1).EType = "D"
        GJRec(1).Desc = "FROM BL"
        GJRec(1).LDesc = "BL Interface"
        GJRec(1).Src = "BL"
        Put #GJFile, , GJRec(1)
      End If
      'If had a penalty amt
      If DPRevAmt#(Process) > 0 Then
        If ACMeth$ = "A" Then
          AcctR = PAnum
          Acct$ = PAAcct$
        Else
          AcctR = PRnum
          Acct$ = PRAcct$
        End If
        If AcctR > 0 Then
          AcctName$ = GetAcctTitle(AcctR)
        Else
          BadAcct = BadAcct + 1
          AcctName$ = "UnDefined"
        End If
        GJRec(1).AcctRec = 0
        GJRec(1).AcctNum = QPTrim$(Acct$)
        GJRec(1).AcctName = AcctName$
        GJRec(1).TRDATE = WorkDate
        GJRec(1).Ref = Ref$
        GJRec(1).CrAmt = DPRevAmt#(Process)
        GJRec(1).DrAmt = 0
        GJRec(1).EType = "C"
        GJRec(1).Desc = "FROM BL"
        GJRec(1).LDesc = "BL Interface"
        GJRec(1).Src = "BL"
        Put #GJFile, , GJRec(1)
        'Now Make Matching Debit Entries to Cash Account
        AcctR = PCnum
        Acct$ = PCAcct
        Acct$ = QPTrim$(Acct$)
        If AcctR > 0 Then
          AcctName$ = GetAcctTitle(AcctR)
        Else
          BadAcct = BadAcct + 1
          AcctName$ = "UnDefined"
        End If
        GJRec(1).AcctRec = 0
        GJRec(1).AcctNum = Acct$
        GJRec(1).AcctName = AcctName$
        GJRec(1).TRDATE = WorkDate
        GJRec(1).Ref = Ref$
        GJRec(1).DrAmt = DPRevAmt#(Process)
        GJRec(1).CrAmt = 0
        GJRec(1).EType = "D"
        GJRec(1).Desc = "FROM BL"
        GJRec(1).LDesc = "BL Interface"
        GJRec(1).Src = "BL"
        Put #GJFile, , GJRec(1)
      End If
    End If
  Next Process

  FundCnt = 0   ' Set Funds Used to Zero

  'Process Payment Adj Type=13

  For Process = 1 To DayCount
    If ttype%(Process) = 13 Then
      If DMRevAmt#(Process) > 0 Then
        Get #ARCatFile, Rec#(Process), ARCatCodeRec(1)
        If ACMeth$ = "A" Then
          AcctR = ARCatCodeRec(1).ARGLACCT
          Acct$ = GetAcctNum(AcctR)
        Else
          AcctR = ARCatCodeRec(1).REVGLNUM
          Acct$ = GetAcctNum(AcctR)
        End If
        If AcctR > 0 Then
          AcctName$ = GetAcctTitle(AcctR)
        Else
          BadAcct = BadAcct + 1
          AcctName$ = "UnDefined"
        End If
        'make debit to rev or rec depending on acctMeth
        GJRec(1).AcctRec = 0
        GJRec(1).AcctNum = QPTrim$(Acct$)
        GJRec(1).AcctName = AcctName$
        GJRec(1).TRDATE = WorkDate
        GJRec(1).Ref = Ref$
        GJRec(1).CrAmt = 0
        GJRec(1).DrAmt = DMRevAmt#(Process)
        GJRec(1).EType = "D"
        GJRec(1).Desc = "FROM BL"
        GJRec(1).LDesc = "BL Interface"
        GJRec(1).Src = "BL"
        Put #GJFile, , GJRec(1)
        'Now Make Matching credit Entries to Cash Account
        AcctR = ARCatCodeRec(1).CashAcct
        Acct$ = GetAcctNum(AcctR)
        Acct$ = QPTrim$(Acct$)
        If AcctR > 0 Then
          AcctName$ = GetAcctTitle(AcctR)
        Else
          BadAcct = BadAcct + 1
          AcctName$ = "UnDefined"
        End If
        GJRec(1).AcctRec = 0
        GJRec(1).AcctNum = Acct$
        GJRec(1).AcctName = AcctName$
        GJRec(1).TRDATE = WorkDate
        GJRec(1).Ref = Ref$
        GJRec(1).DrAmt = 0
        GJRec(1).CrAmt = DMRevAmt#(Process)
        GJRec(1).EType = "C"
        GJRec(1).Desc = "FROM BL"
        GJRec(1).LDesc = "BL Interface"
        GJRec(1).Src = "BL"
        Put #GJFile, , GJRec(1)
      End If
      'If had a penalty amt
      If DPRevAmt#(Process) > 0 Then
        If ACMeth$ = "A" Then
          AcctR = PAnum
          Acct$ = PAAcct$
        Else
          AcctR = PRnum
          Acct$ = PRAcct$
        End If
        If AcctR > 0 Then
          AcctName$ = GetAcctTitle(AcctR)
        Else
          BadAcct = BadAcct + 1
          AcctName$ = "UnDefined"
        End If
        'Make debit to revenue or accrual depending on acctmeth
        GJRec(1).AcctRec = 0
        GJRec(1).AcctNum = QPTrim$(Acct$)
        GJRec(1).AcctName = AcctName$
        GJRec(1).TRDATE = WorkDate
        GJRec(1).Ref = Ref$
        GJRec(1).CrAmt = 0
        GJRec(1).DrAmt = DPRevAmt#(Process)
        GJRec(1).EType = "D"
        GJRec(1).Desc = "FROM BL"
        GJRec(1).LDesc = "BL Interface"
        GJRec(1).Src = "BL"
        Put #GJFile, , GJRec(1)
        'Now Make Matching Credit Entries to Cash Account
        AcctR = PCnum
        Acct$ = PCAcct
        Acct$ = QPTrim$(Acct$)
        If AcctR > 0 Then
          AcctName$ = GetAcctTitle(AcctR)
        Else
          BadAcct = BadAcct + 1
          AcctName$ = "UnDefined"
        End If
        GJRec(1).AcctRec = 0
        GJRec(1).AcctNum = Acct$
        GJRec(1).AcctName = AcctName$
        GJRec(1).TRDATE = WorkDate
        GJRec(1).Ref = Ref$
        GJRec(1).DrAmt = 0
        GJRec(1).CrAmt = DPRevAmt#(Process)
        GJRec(1).EType = "C"
        GJRec(1).Desc = "FROM BL"
        GJRec(1).LDesc = "BL Interface"
        GJRec(1).Src = "BL"
        Put #GJFile, , GJRec(1)
      End If
    End If
  Next Process
    
  'Process Charges if Needed Type=1 For Accrual Only

  FundCnt = 0

  For Process = 1 To DayCount
    If ttype%(Process) = 1 Or ttype%(Process) = 24 Or ttype%(Process) = 6 Then
      If DMRevAmt#(Process) > 0 Then
        Get #ARCatFile, Rec#(Process), ARCatCodeRec(1)
        If ACMeth$ = "A" Then
  '        Acct$ = ARCatCodeRec(1).REVGLNUM
  '        Acct$ = QPStrip$(Acct$)
  '        Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
          AcctR = ARCatCodeRec(1).REVGLNUM 'AcctFind(Acct$)
          If AcctR > 0 Then
            Acct$ = GetAcctNum(AcctR)
            AcctName$ = GetAcctTitle(AcctR)
          Else
            BadAcct = BadAcct + 1
            AcctName$ = "UnDefined"
          End If
          GJRec(1).AcctRec = 0
          GJRec(1).AcctNum = QPTrim$(Acct$)
          GJRec(1).AcctName = AcctName$
          GJRec(1).TRDATE = WorkDate
          GJRec(1).Ref = Ref$
          GJRec(1).CrAmt = DMRevAmt#(Process)
          GJRec(1).DrAmt = 0
          GJRec(1).EType = "C"
          GJRec(1).Desc = "FROM BL"
          GJRec(1).LDesc = "BL Interface"
          GJRec(1).Src = "BL"
          Put #GJFile, , GJRec(1)
  
          'Now Make Matching Debit Entries to receivables
  '        Acct$ = ARCatCodeRec(1).ARGLACCT
  '        Acct$ = QPStrip$(Acct$)
  '        Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
          AcctR = ARCatCodeRec(1).ARGLACCT 'AcctFind(Acct$)
          If AcctR > 0 Then
            Acct$ = GetAcctNum(AcctR)
            AcctName$ = GetAcctTitle(AcctR)
          Else
            BadAcct = BadAcct + 1
            AcctName$ = "UnDefined"
          End If
          GJRec(1).AcctRec = 0
          GJRec(1).AcctNum = Acct$
          GJRec(1).AcctName = AcctName$
          GJRec(1).TRDATE = WorkDate
          GJRec(1).Ref = Ref$
          GJRec(1).DrAmt = DMRevAmt#(Process)
          GJRec(1).CrAmt = 0
          GJRec(1).EType = "D"
          GJRec(1).Desc = "FROM BL"
          GJRec(1).LDesc = "BL Interface"
          GJRec(1).Src = "BL"
          Put #GJFile, , GJRec(1)
        End If
      End If
      If DPRevAmt#(Process) > 0 Then
        If ACMeth$ = "A" Then
          AcctR = PRnum
          If AcctR > 0 Then
            Acct$ = PRAcct$
            AcctName$ = GetAcctTitle(AcctR)
          Else
            BadAcct = BadAcct + 1
            AcctName$ = "UnDefined"
          End If
          GJRec(1).AcctRec = 0
          GJRec(1).AcctNum = QPTrim$(Acct$)
          GJRec(1).AcctName = AcctName$
          GJRec(1).TRDATE = WorkDate
          GJRec(1).Ref = Ref$
          GJRec(1).CrAmt = DPRevAmt#(Process)
          GJRec(1).DrAmt = 0
          GJRec(1).EType = "C"
          GJRec(1).Desc = "FROM BL"
          GJRec(1).LDesc = "BL Interface"
          GJRec(1).Src = "BL"
          Put #GJFile, , GJRec(1)
  
          'Now Make Matching Debit Entries to Receivables
  '        Acct$ = ARCatCodeRec(1).ARGLACCT
  '        Acct$ = QPStrip$(Acct$)
  '        Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
          AcctR = PAnum
          If AcctR > 0 Then
            Acct$ = PAAcct$
            AcctName$ = GetAcctTitle(AcctR)
          Else
            BadAcct = BadAcct + 1
            AcctName$ = "UnDefined"
          End If
          GJRec(1).AcctRec = 0
          GJRec(1).AcctNum = Acct$
          GJRec(1).AcctName = AcctName$
          GJRec(1).TRDATE = WorkDate
          GJRec(1).Ref = Ref$
          GJRec(1).DrAmt = DPRevAmt#(Process)
          GJRec(1).CrAmt = 0
          GJRec(1).EType = "D"
          GJRec(1).Desc = "FROM BL"
          GJRec(1).LDesc = "BL Interface"
          GJRec(1).Src = "BL"
          Put #GJFile, , GJRec(1)
        End If
      End If
    End If
  Next Process
  
  'Process Down Adj Type=23
  FundCnt = 0

  For Process = 1 To DayCount
    If ttype%(Process) = 23 Then
      If DMRevAmt#(Process) > 0 Then
        Get #ARCatFile, Rec#(Process), ARCatCodeRec(1)
        If ACMeth$ = "A" Then
  '        Acct$ = ARCatCodeRec(1).REVGLNUM
  '        Acct$ = QPStrip$(Acct$)
  '        Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
          AcctR = ARCatCodeRec(1).REVGLNUM 'AcctFind(Acct$)
          If AcctR > 0 Then
            Acct$ = GetAcctNum(AcctR)
            AcctName$ = GetAcctTitle(AcctR)
          Else
            BadAcct = BadAcct + 1
            AcctName$ = "UnDefined"
          End If
          GJRec(1).AcctRec = 0
          GJRec(1).AcctNum = QPTrim$(Acct$)
          GJRec(1).AcctName = AcctName$
          GJRec(1).TRDATE = WorkDate
          GJRec(1).Ref = Ref$
          GJRec(1).CrAmt = 0
          GJRec(1).DrAmt = DMRevAmt#(Process)
          GJRec(1).EType = "D"
          GJRec(1).Desc = "FROM BL"
          GJRec(1).LDesc = "BL Interface"
          GJRec(1).Src = "BL"
          Put #GJFile, , GJRec(1)
  
          'Now Make Matching Debit Entries to Cash Account
  '        Acct$ = ARCatCodeRec(1).ARGLACCT
  '        Acct$ = QPStrip$(Acct$)
  '        Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
          AcctR = ARCatCodeRec(1).ARGLACCT 'AcctFind(Acct$)
          If AcctR > 0 Then
            Acct$ = GetAcctNum(AcctR)
            AcctName$ = GetAcctTitle(AcctR)
          Else
            BadAcct = BadAcct + 1
            AcctName$ = "UnDefined"
          End If
          GJRec(1).AcctRec = 0
          GJRec(1).AcctNum = Acct$
          GJRec(1).AcctName = AcctName$
          GJRec(1).TRDATE = WorkDate
          GJRec(1).Ref = Ref$
          GJRec(1).DrAmt = 0
          GJRec(1).CrAmt = DMRevAmt#(Process)
          GJRec(1).EType = "C"
          GJRec(1).Desc = "FROM BL"
          GJRec(1).LDesc = "BL Interface"
          GJRec(1).Src = "BL"
          Put #GJFile, , GJRec(1)
        End If
      End If
      If DPRevAmt#(Process) > 0 Then
        If ACMeth$ = "A" Then
          AcctR = PRnum
          If AcctR > 0 Then
            Acct$ = PRAcct$
            AcctName$ = GetAcctTitle(AcctR)
          Else
            BadAcct = BadAcct + 1
            AcctName$ = "UnDefined"
          End If
          GJRec(1).AcctRec = 0
          GJRec(1).AcctNum = QPTrim$(Acct$)
          GJRec(1).AcctName = AcctName$
          GJRec(1).TRDATE = WorkDate
          GJRec(1).Ref = Ref$
          GJRec(1).CrAmt = 0
          GJRec(1).DrAmt = DPRevAmt#(Process)
          GJRec(1).EType = "D"
          GJRec(1).Desc = "FROM BL"
          GJRec(1).LDesc = "BL Interface"
          GJRec(1).Src = "BL"
          Put #GJFile, , GJRec(1)
  
          'Now Make Matching Debit Entries to Receivables
  '        Acct$ = ARCatCodeRec(1).ARGLACCT
  '        Acct$ = QPStrip$(Acct$)
  '        Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
          AcctR = PAnum
          If AcctR > 0 Then
            Acct$ = PAAcct$
            AcctName$ = GetAcctTitle(AcctR)
          Else
            BadAcct = BadAcct + 1
            AcctName$ = "UnDefined"
          End If
          GJRec(1).AcctRec = 0
          GJRec(1).AcctNum = Acct$
          GJRec(1).AcctName = AcctName$
          GJRec(1).TRDATE = WorkDate
          GJRec(1).Ref = Ref$
          GJRec(1).DrAmt = 0
          GJRec(1).CrAmt = DPRevAmt#(Process)
          GJRec(1).EType = "C"
          GJRec(1).Desc = "FROM BL"
          GJRec(1).LDesc = "BL Interface"
          GJRec(1).Src = "BL"
          Put #GJFile, , GJRec(1)
        End If
      End If
    End If
  Next Process

  
BLBunchReturn:
  Return

End Sub


'New Tax Billing Interface
Private Sub ExtractNTX(ThruDate%)
  Dim Ref As String, Dash80 As String, P2S As String, TXGLFile As Integer
  Dim GJReclen As Integer, RptFile As Integer, TaxAcctRecLen As Integer
  Dim TaxTranRecLen As Integer, NumOfTRecs As Long, TCnt As Long
  Dim FoundCnt As Integer, NGCnt As Integer, GJFile As Integer, CDCashAcct As String
  Dim NumEdTrans As Integer, MCFile As Integer, cnt As Integer, CDCashAcctName As String
  Dim FirstTran As Integer, ThisDate As Integer, WorkDate As Integer
  Dim DayCount As Integer, LastTran As Integer, RevCnt As Integer, FundDue As String
  Dim FindCount As Integer, FundCnt As Integer, ThisAcct As Integer
  Dim Acct As String, AcctName As String, T As Integer, BadAcct As Integer
  Dim FoundFund As Integer, PCnt As Integer, Cash As Integer, CDCashRec As Long
  Dim txfile As Integer, InterfaceMethod As Integer, AcctMeth As String
  Dim TaxYear As Integer, MiddleRec As Integer, TranFile As Integer, DetPad As String
  Dim DebitAcctRecord As Integer, DebitAcctNumber As String, ppcnt As Integer
  Dim DebitAcctName As String, DebitAmt As Double, CreditAcctRecord As Integer
  Dim CreditAcctName As String, CreditAmt As Double, CreditAcctNumber As String
  Dim CDDueAcct As String, CDDueRec As Long, CDDueName As String, PadChars As Integer
  Dim GJInfo() As GJXferRecType
  Ref$ = "TX" + Left$(Date$, 2) + Mid$(Date$, 4, 2) + Right$(Date$, 2)
  ReDim TranInfo(1) As TranRecInfoType
  Dim TPrinciple#(15, 51), TInterest#(15, 51), TCollection#(15, 51)
  Dim TPrinciplePd#(15, 51), TInterestPd#(15, 51), TCollectionPd#(15, 51)
  Dim TRevOpt1#(15, 51), TRevOpt1Pd#(15, 51), TRevOpt2#(15, 51), TRevOpt2Pd#(15, 51), TRevOpt3#(15, 51)
  Dim TRevOpt3Pd#(15, 51), TLateList#(15, 51), TLateListPd#(15, 51), TPrePaidAmt#(15, 51)
  Dim TPrePaidUsed#(15, 51), Dsc As String, TrType As Integer, y As Integer, Curyr As Integer
  ReDim Preserve GJInfo(1 To 3) As GJXferRecType
  Dim tttest As Long
  Dim ttcust As Long
  ReDim TaxGLAccts(1) As TaxAcctsType
  TaxAcctRecLen = Len(TaxGLAccts(1))

  Dim TaxTrans(1) As TaxTransactionType
  TaxTranRecLen = Len(TaxTrans(1))
  BadAcct = 0
  ReDim TaxSetuprec(1) As TaxMasterType
  txfile = FreeFile
  Open "TAXSETUP.DAT" For Random As #txfile Len = Len(TaxSetuprec(1))
  If LOF(txfile) > 0 Then
    Get txfile, 1, TaxSetuprec(1)
  Else
    Unload FrmShowPctComp
    MsgBox "No Tax Setup File Information.", vbOKOnly, "No Setup"
    GoTo TaxEnd
  End If
  'If Central Depository used then will need detail for acct #
    CDActive$ = QPTrim$(TaxSetuprec(1).CntrlDepYN)
    If CDActive$ = "Y" Then
      PadChars = GLDetLen - GLFundLen
      If PadChars > 0 Then
        DetPad$ = String(PadChars, "0")
      End If
      CDCashAcct$ = TaxSetuprec(1).CDCashGL
      CDCashAcct$ = QPStrip$(CDCashAcct$)
      CDCashAcct$ = FmtAcct$(CDCashAcct$, GLFundLen%, GLAcctLen%, GLDetLen%)
      CDCashAcct$ = QPTrim$(CDCashAcct$)
      CDDueAcct$ = QPTrim$(TaxSetuprec(1).CDSubGL)
      CDDueAcct$ = QPStrip$(CDDueAcct$)
      
      CDCashRec = AcctFind(CDCashAcct$)
      If CDCashRec <= 0 Then
        Unload FrmShowPctComp
        MsgBox "The Account for Central Cash Is Not Setup, Please Correct Before Trying Again.", vbOKOnly, "Invalid Tax Setup"
        GoTo TaxEnd
      Else
        CDCashAcctName$ = GetAcctTitle$(CDCashRec)
      End If
    End If
  FrmShowPctComp.ShowPctComp 10, 100
  AcctMeth$ = QPTrim$(TaxSetuprec(1).AcctgMethod)
  Curyr = Right$(Num2Date$(TaxSetuprec(1).TaxYear), 4)
  If (Len(AcctMeth$) = 0) Then
    Unload FrmShowPctComp
    MsgBox "The Accounting Method Is Not Setup, Please Correct Before Trying Again.", vbOKOnly, "Invalid Tax Setup"
    GoTo TaxEnd
  End If
  Select Case AcctMeth$
  Case "C"
    InterfaceMethod = 1
  Case "A"
    InterfaceMethod = 2
  Case "M"
    InterfaceMethod = 3
  Case Else
    Unload FrmShowPctComp
    MsgBox "The Accounting Method Is Invalid, Please Correct Before Trying Again.", vbOKOnly, "Invalid Tax Setup"
    GoTo EndTax
  End Select
  Close txfile
  
  GJReclen = Len(GJRec(1))
  
  If Exist("TAXGLACT.DAT") Then
    TXGLFile = FreeFile
    Open "TAXGLACT.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
    Get TXGLFile, 1, TaxGLAccts(1)
  Else
    Unload FrmShowPctComp
    MsgBox "Tax Accounts Not Setup,Interface File Not Created.", vbOKOnly, "Tax Acct Setup Invalid"
    GoTo EndTax
  Close
  End If
  FrmShowPctComp.Label1 = "Sorting Interface Trans File"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me

  'Now Process the Transactions
  'ClearBox

 ' QPrintRC "Searching Cash Transactions.", 11, 26, 126
 ' QPrintRC "New Transactions:", 13, 29, Cnf.HiLite
  FrmShowPctComp.ShowPctComp 25, 100
  TranFile = FreeFile
  Open "TAXTRANS.DAT" For Random Shared As TranFile Len = TaxTranRecLen
  NumOfTRecs& = LOF(TranFile) \ TaxTranRecLen
  For TCnt& = NumOfTRecs& To 1 Step -1
    Get #TranFile, TCnt&, TaxTrans(1)
    If Len(QPTrim$(TaxTrans(1).Posted2GL)) = 0 Or QPTrim$(TaxTrans(1).Posted2GL) = "N" Then
      'Store trans rec numbers and dates in array
      If TaxTrans(1).TransDate <= ThruDate% Then
        FoundCnt = FoundCnt + 1
        ReDim Preserve TranInfo(FoundCnt) As TranRecInfoType
        TranInfo(FoundCnt).TranDate = TaxTrans(1).TransDate
        TranInfo(FoundCnt).TranRecNo = TCnt&
      End If
    Else
      NGCnt = NGCnt + 1
    End If
    P2S$ = Str$(FoundCnt)
    'QPrintRC P2S$, 13, 47, Cnf.HiLite
    'SmallPause
    If NGCnt >= 2500 Then Exit For
  Next
  'FrmShowPctComp.ShowPctComp 40, 100
  If FoundCnt = 0 Then
    Close
    Unload FrmShowPctComp
    Call MainLog("No Tx to Grab " + Str$(FoundCnt) + " for " + fpDate)
    MsgBox "No Transactions Found to Interface.", vbOKOnly, "No Trans"
    GoTo EndTax
  End If
  FrmShowPctComp.Label1 = "Sorting Interface Trans File"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  FrmShowPctComp.ShowPctComp 15, 100
  SortTRec TranInfo(), FoundCnt      'sort'em by date. oldest first

  GJFile = FreeFile
  Open "GLTRXED.DAT" For Random Access Read Write Shared As GJFile Len = GJReclen
  FrmShowPctComp.ShowPctComp 35, 100
  FirstTran = 1
  ThisDate = TranInfo(1).TranDate
  For cnt = 1 To FoundCnt
    FrmShowPctComp.ShowPctComp cnt, FoundCnt
    If ThisDate <> TranInfo(cnt).TranDate Then
      ThisDate = TranInfo(cnt).TranDate
      LastTran = cnt - 1
      GoSub ProcessThisBunchTX
      FirstTran = cnt
    End If
  Next cnt
  FrmShowPctComp.Label1 = "Tag Interface Transactions"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me

  LastTran = FoundCnt
  GoSub ProcessThisBunchTX
  If BadAcct > 0 Then
    Close
    Unload FrmShowPctComp
    Call MainLog("Error - TX Grab Not Created Due to invalid or missing accts.")
    MsgBox "Errors Found, Interface File Not Created. Please Review Report.", vbOKOnly, "Errors"
    frmReportOpt.Show 1
    If rptopt = 1 Then
      frmGetDistMenu.PrnEditList 2
    ElseIf rptopt = 2 Then
      frmGetDistMenu.PrnEditList2 2
    End If
    KillFileD "GLTRXED.DAT"
    Exit Sub
  End If
  'transactions as interfaced
  
  For cnt = 1 To FoundCnt
    FrmShowPctComp.ShowPctComp cnt, FoundCnt
    Get #TranFile, TranInfo(cnt).TranRecNo, TaxTrans(1)
    TaxTrans(1).Posted2GL = "Y"
    Put #TranFile, TranInfo(cnt).TranRecNo, TaxTrans(1)
  Next cnt
  Close
  'SLEEP 2
  Call MainLog("TX Grab Complete for " + fpDate)
  MsgBox "Transaction Grab Complete.", vbOKOnly, "Complete"
  GoTo EndTax

EndTax:
  Unload FrmShowPctComp
  Exit Sub

ProcessThisBunchTX:      'Initialize for This Set
GoSub Clearouttots

  For PCnt = FirstTran To LastTran
    If PCnt = FirstTran Then
      WorkDate = TranInfo(PCnt).TranDate
    End If
    Get #TranFile, TranInfo(PCnt).TranRecNo, TaxTrans(1)
    'Now Decipher by Type and Year
    Select Case TaxTrans(1).TranType
      Case 1:
        TrType = 1
      Case 2:
        TrType = 2
      Case 3:
        TrType = 3
      Case 4:
        TrType = 4
      Case 6:
        TrType = 6
      Case 7:
        TrType = 7
      Case 9:
        TrType = 9
      Case 10:
        TrType = 10
      Case 11:
        TrType = 11
      Case 12:
        TrType = 12
      Case 13:
        TrType = 13
      Case 14:
        TrType = 14
      Case 21:
        TrType = 5
      Case 22:
        TrType = 8
      Case 24:
        TrType = 15
      End Select
      TaxYear = TaxTrans(1).TaxYear
      If TaxYear < 1 Then TaxYear = Curyr
      TaxYear = TaxYear - 1979             'Reduce Based on 1980 being = 1
'To test prob with Whitelake invalid trans created when interf taxes  8/24/09
      'ttcust = TaxTrans(1).CustomerRec
      'tttest = TaxTrans(1).CustPin
      'If TaxTrans(1).Revenue.Principle1Pd <> 0 Then Stop
      'If TaxTrans(1).Revenue.LateListPd <> 0 Then Stop
      
      TPrinciple#(TrType, TaxYear) = TPrinciple#(TrType, TaxYear) + TaxTrans(1).Revenue.Principle1
      TInterest#(TrType, TaxYear) = TInterest#(TrType, TaxYear) + TaxTrans(1).Revenue.Interest
      TCollection#(TrType, TaxYear) = TCollection#(TrType, TaxYear) + TaxTrans(1).Revenue.Collection
      TLateList#(TrType, TaxYear) = TLateList#(TrType, TaxYear) + TaxTrans(1).Revenue.LateList
      TRevOpt1#(TrType, TaxYear) = TRevOpt1#(TrType, TaxYear) + TaxTrans(1).Revenue.RevOpt1
      TRevOpt2#(TrType, TaxYear) = TRevOpt2#(TrType, TaxYear) + TaxTrans(1).Revenue.RevOpt2
      TRevOpt3#(TrType, TaxYear) = TRevOpt3#(TrType, TaxYear) + TaxTrans(1).Revenue.RevOpt3
      TPrinciplePd#(TrType, TaxYear) = TPrinciplePd#(TrType, TaxYear) + TaxTrans(1).Revenue.Principle1Pd
      TInterestPd#(TrType, TaxYear) = TInterestPd#(TrType, TaxYear) + TaxTrans(1).Revenue.InterestPd
      TCollectionPd#(TrType, TaxYear) = TCollectionPd#(TrType, TaxYear) + TaxTrans(1).Revenue.CollectionPd
      TLateListPd#(TrType, TaxYear) = TLateListPd#(TrType, TaxYear) + TaxTrans(1).Revenue.LateListPd
      TRevOpt1Pd#(TrType, TaxYear) = TRevOpt1Pd#(TrType, TaxYear) + TaxTrans(1).Revenue.RevOpt1Pd
      TRevOpt2Pd#(TrType, TaxYear) = TRevOpt2Pd#(TrType, TaxYear) + TaxTrans(1).Revenue.RevOpt2Pd
      TRevOpt3Pd#(TrType, TaxYear) = TRevOpt3Pd#(TrType, TaxYear) + TaxTrans(1).Revenue.RevOpt3Pd
      TPrePaidAmt#(TrType, TaxYear) = TPrePaidAmt#(TrType, TaxYear) + TaxTrans(1).Revenue.PrePaidAmt
      TPrePaidUsed#(TrType, TaxYear) = TPrePaidUsed#(TrType, TaxYear) + TaxTrans(1).Revenue.PrePaidUsed
  Next PCnt
If InterfaceMethod <> 1 Then  ' 1 is cash so skip all charges if cash
    'Now Post for bill trans and move on to next type
  'TranType1 Billing
  For T% = 1 To 51
    If TPrinciple#(1, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct)) <= 0 Then
'        GoTo NoGoodSkip1P
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TPrinciple#(1, T%)
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).TaxCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Invalid Acct"
      End If
      CreditAmt# = TPrinciple#(1, T%)
      Dsc$ = "Bills"
      GoSub PostToGeneralJournal
NoGoodSkip1P:
      Close TXGLFile
    End If
    If TRevOpt1#(1, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct)) <= 0 Then
'        GoTo NoGoodSkip2
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TRevOpt1#(1, T%)
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt1CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt1CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Invaild Acct"
      End If
      CreditAmt# = TRevOpt1#(1, T%)
      Dsc$ = "Bills"
      GoSub PostToGeneralJournal
NoGoodSkip2:
      Close TXGLFile
    End If
    If TRevOpt2#(1, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct)) <= 0 Then
'        GoTo NoGoodSkip3
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TRevOpt2#(1, T%)
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt2CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt2CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Invaild Acct"
      End If
      CreditAmt# = TRevOpt2#(1, T%)
      Dsc$ = "Bills"
      GoSub PostToGeneralJournal
NoGoodSkip3:
      Close TXGLFile
    End If
    If TRevOpt3#(1, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct)) <= 0 Then
'        GoTo NoGoodSkip4
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TRevOpt3#(1, T%)
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt3CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt3CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Invaild Acct"
      End If
      CreditAmt# = TRevOpt3#(1, T%)
      Dsc$ = "Bills"
      GoSub PostToGeneralJournal
NoGoodSkip4:
      Close TXGLFile
    End If
    If TLateList#(1, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct)) <= 0 Then
'        GoTo NoGoodSkip1L
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TLateList#(1, T%)
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).LtLstCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).LtLstCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Invaild Acct"
      End If
      CreditAmt# = TLateList#(1, T%)
      Dsc$ = "BillsLL"
      GoSub PostToGeneralJournal
NoGoodSkip1L:
      Close TXGLFile
    End If
  Next T%
  '@@@@@@@@@@@@@@@@end of billing trantype-1
  'Interest charge  tran-4
  For T% = 1 To 51
    If TInterest#(4, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)) <= 0 Then
'        GoTo NoGoodSkip5
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TInterest#(4, T%)
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).IntCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TInterest#(4, T%)
      Dsc$ = "Interest"
      GoSub PostToGeneralJournal
NoGoodSkip5:
      Close TXGLFile
    End If
  Next T%
  '@@@@@@@@@@end of tran type 4 interest charge
  ''''''''''''''''''''''''''''''''''''''''''''''
  'Ad/Collection charge tran 6
  For T% = 1 To 51
    If TCollection#(6, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).AdvDBAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).AdvDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TCollection#(6, T%)
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).AdvCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).AdvCRAcct
     If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TCollection#(6, T%)
      Dsc$ = "AdCharge"
      GoSub PostToGeneralJournal
      Close TXGLFile
    End If
  Next T%
'@@@@@@@@@@@@@@@@@@@@end of tran type 6 ad charge(collection)
''''''''''''
'Up Adj  14, 24
  For T% = 14 To 51
    If TPrinciple#(14, T%) > 0 Or TPrinciple#(15, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct)) <= 0 Then
'        GoTo NoGoodSkip100
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = Round#(TPrinciple#(14, T%) + TPrinciple#(15, T%))
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).TaxCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TPrinciple#(14, T%) + TPrinciple#(15, T%))
      Dsc$ = "UpAdjP"
      GoSub PostToGeneralJournal
NoGoodSkip100:
      Close TXGLFile
    End If
    If TRevOpt1#(14, T%) > 0 Or TRevOpt1#(15, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct)) <= 0 Then
'        GoTo NoGoodSkip200
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = Round#(TRevOpt1#(14, T%) + TRevOpt1#(15, T%))
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt1CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt1CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TRevOpt1#(14, T%) + TRevOpt1#(15, T%))
      Dsc$ = "UpAdjOR1"
      GoSub PostToGeneralJournal
NoGoodSkip200:
      Close TXGLFile
    End If
    If TRevOpt2#(14, T%) > 0 Or TRevOpt2#(15, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct)) <= 0 Then
'        GoTo NoGoodSkip300
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = Round#(TRevOpt2#(14, T%) + TRevOpt2#(15, T%))
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt2CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt2CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TRevOpt2#(14, T%) + TRevOpt2#(15, T%))
      Dsc$ = "UpAdjOR2"
      GoSub PostToGeneralJournal
NoGoodSkip300:
      Close TXGLFile
    End If
    If TRevOpt3#(14, T%) > 0 Or TRevOpt3#(15, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct)) <= 0 Then
'        GoTo NoGoodSkip400
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = Round#(TRevOpt3#(14, T%) + TRevOpt3#(15, T%))
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt3CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt3CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TRevOpt3#(14, T%) + TRevOpt3#(15, T%))
      Dsc$ = "UpAdjOR3"
      GoSub PostToGeneralJournal
NoGoodSkip400:
      Close TXGLFile
    End If
    If TLateList#(14, T%) > 0 Or TLateList#(15, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct)) <= 0 Then
'        GoTo NoGoodSkip1L00
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = Round#(TLateList#(14, T%) + TLateList#(15, T%))
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).LtLstCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).LtLstCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TLateList#(14, T%) + TLateList#(15, T%))
      Dsc$ = "UpAdjLL"
      GoSub PostToGeneralJournal
NoGoodSkip1L00:
      Close TXGLFile
    End If
    If TInterest#(14, T%) > 0 Or TInterest#(15, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)) <= 0 Then
'        GoTo NoGoodSkip500
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = Round#(TInterest#(14, T%) + TInterest#(15, T%))
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).IntCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TInterest#(14, T%) + TInterest#(15, T%))
      Dsc$ = "UpAdjInt"
      GoSub PostToGeneralJournal
NoGoodSkip500:
      Close TXGLFile
    End If
    If TCollection#(14, T%) > 0 Or TCollection#(15, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)) <= 0 Then
'        GoTo NoGoodSkip501
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = Round#(TCollection#(14, T%) + TCollection#(15, T%))
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).IntCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TCollection#(14, T%) + TCollection#(15, T%))
      Dsc$ = "UpAdjAC"
      GoSub PostToGeneralJournal
NoGoodSkip501:
      Close TXGLFile
    End If
  Next T%
'end of up adj  14, 24
'tran type 3-Release
'changed from charged to paid amounts per Bob on 7/12/06 but still used billing accts
  For T% = 1 To 51
    If TPrinciplePd#(3, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct)) <= 0 Then
'        GoTo NoGoodSkipRls1
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).TaxCRAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TPrinciplePd#(3, T%)
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TPrinciplePd#(3, T%)
      Dsc$ = "Release"
      GoSub PostToGeneralJournal
NoGoodSkipRls1:
      Close TXGLFile
    End If
    If TRevOpt1Pd#(3, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct)) <= 0 Then
'        GoTo NoGoodSkipRls2
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt1CRAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt1CRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TRevOpt1Pd#(3, T%)
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TRevOpt1Pd#(3, T%)
      Dsc$ = "Release"
      GoSub PostToGeneralJournal
NoGoodSkipRls2:
      Close TXGLFile
    End If
    If TRevOpt2Pd#(3, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct)) <= 0 Then
'        GoTo NoGoodSkipRls3
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt2CRAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt2CRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TRevOpt2Pd#(3, T%)
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TRevOpt2Pd#(3, T%)
      Dsc$ = "Release"
      GoSub PostToGeneralJournal
NoGoodSkipRls3:
      Close TXGLFile
    End If
    If TRevOpt3Pd#(3, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct)) <= 0 Then
'        GoTo NoGoodSkipRls4
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt3CRAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt3CRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TRevOpt3Pd#(3, T%)
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TRevOpt3Pd#(3, T%)
      Dsc$ = "Release"
      GoSub PostToGeneralJournal
NoGoodSkipRls4:
      Close TXGLFile
    End If
    If TLateListPd#(3, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct)) <= 0 Then
'        GoTo NoGoodSkipRls5
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).LtLstCRAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).LtLstCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TLateListPd#(3, T%)
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TLateListPd#(3, T%)
      Dsc$ = "Release"
      GoSub PostToGeneralJournal
NoGoodSkipRls5:
      Close TXGLFile
    End If
  'Interest release
    If TInterestPd#(3, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)) <= 0 Then
'        GoTo NoGoodSkipRls6
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).IntCRAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TInterestPd#(3, T%)
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TInterestPd#(3, T%)
      Dsc$ = "Release"
      GoSub PostToGeneralJournal
NoGoodSkipRls6:
      Close TXGLFile
    End If
  'Ad/Collection release
    If TCollectionPd#(3, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)) <= 0 Then
'        GoTo NoGoodSkipRls7
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).AdvCRAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).AdvCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TCollectionPd#(3, T%)
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).AdvDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).AdvDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TCollectionPd#(3, T%)
      Dsc$ = "Release"
      GoSub PostToGeneralJournal
NoGoodSkipRls7:
      Close TXGLFile
    End If
  Next T%
'tran type 13-Adj bill down
  For T% = 1 To 51
    If TPrinciple#(13, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct)) <= 0 Then
'        GoTo NoGoodSkipAdj1
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).TaxCRAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TPrinciple#(13, T%)
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TPrinciple#(13, T%)
      Dsc$ = "AdjBillDown"
      GoSub PostToGeneralJournal
NoGoodSkipAdj1:
      Close TXGLFile
    End If
    If TRevOpt1#(13, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct)) <= 0 Then
'        GoTo NoGoodSkipAdj2
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt1CRAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt1CRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TRevOpt1#(13, T%)
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TRevOpt1#(13, T%)
      Dsc$ = "AdjBillDown"
      GoSub PostToGeneralJournal
NoGoodSkipAdj2:
      Close TXGLFile
    End If
    If TRevOpt2#(13, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct)) <= 0 Then
'        GoTo NoGoodSkipAdj3
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt2CRAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt2CRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TRevOpt2#(13, T%)
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TRevOpt2#(13, T%)
      Dsc$ = "AdjBillDown"
      GoSub PostToGeneralJournal
NoGoodSkipAdj3:
      Close TXGLFile
    End If
    If TRevOpt3#(13, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct)) <= 0 Then
'        GoTo NoGoodSkipAdj4
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt3CRAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt3CRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TRevOpt3#(13, T%)
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TRevOpt3#(13, T%)
      Dsc$ = "AdjBillDown"
      GoSub PostToGeneralJournal
NoGoodSkipAdj4:
      Close TXGLFile
    End If
    If TLateList#(13, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct)) <= 0 Then
'        GoTo NoGoodSkipAdj5
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).LtLstCRAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).LtLstCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TLateList#(13, T%)
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TLateList#(13, T%)
      Dsc$ = "AdjBillDown"
      GoSub PostToGeneralJournal
NoGoodSkipAdj5:
      Close TXGLFile
    End If
  'Interest adjdownforbill
    If TInterest#(13, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)) <= 0 Then
'        GoTo NoGoodSkipAdj6
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).IntCRAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TInterest#(13, T%)
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TInterest#(13, T%)
      Dsc$ = "AdjBillDown"
      GoSub PostToGeneralJournal
NoGoodSkipAdj6:
      Close TXGLFile
    End If
  'Ad/Collection adjdownforbill
    If TCollection#(13, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)) <= 0 Then
'        GoTo NoGoodSkipAdj7
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).AdvCRAcct)
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).AdvCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TCollection#(13, T%)
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).AdvDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).AdvDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TCollection#(13, T%)
      Dsc$ = "AdjBillDown"
      GoSub PostToGeneralJournal
NoGoodSkipAdj7:
      Close TXGLFile
    End If
  Next T%

End If
'''end of tran type 13
''also end of charges ups/downs
'''''''''''''''''''''''''
  'TranType2 Payments-2,Payment w/prepay-21, Prepay only-22
  For T% = 1 To 51
    If TPrinciplePd#(2, T%) > 0 Or TPrinciplePd#(5, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXGLFile = FreeFile
        Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
        Get TXGLFile, 1, TaxGLAccts(1)
'        If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct)) <= 0 Then
'          GoTo NoGoodSkipP2
'        End If
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).TaxCRAcct)
        DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TPrinciplePd#(2, T%) + TPrinciplePd#(5, T%))
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TPrinciplePd#(2, T%) + TPrinciplePd#(5, T%))
        Dsc$ = "PaymentP"
        GoSub PostToGeneralJournal
        Close TXGLFile
      End If
      TXGLFile = FreeFile
      Open "TAXGLACT.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct)) <= 0 Then
'        GoTo NoGoodSkipP2
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxDBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(TPrinciplePd#(2, T%) + TPrinciplePd#(5, T%))
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).TaxCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TPrinciplePd#(2, T%) + TPrinciplePd#(5, T%))
      Dsc$ = "PaymentP"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
NoGoodSkipP2:
      Close TXGLFile
    End If
    If TInterestPd#(2, T%) > 0 Or TInterestPd#(5, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXGLFile = FreeFile
        Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
        Get TXGLFile, 1, TaxGLAccts(1)
'        If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)) <= 0 Then
'          GoTo NoGoodSkipP3
'        End If
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).IntCRAcct)
        DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TInterestPd#(2, T%) + TInterestPd#(5, T%))
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TInterestPd#(2, T%) + TInterestPd#(5, T%))
        Dsc$ = "PaymentI"
        GoSub PostToGeneralJournal
        Close TXGLFile
      End If
      TXGLFile = FreeFile
      Open "TAXGLACT.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)) <= 0 Then
'        GoTo NoGoodSkipP3
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntDBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(TInterestPd#(2, T%) + TInterestPd#(5, T%))
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).IntCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TInterestPd#(2, T%) + TInterestPd#(5, T%))
      Dsc$ = "PaymentI"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
NoGoodSkipP3:
      Close TXGLFile
    End If
    If TCollectionPd#(2, T%) > 0 Or TCollectionPd#(5, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXGLFile = FreeFile
        Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
        Get TXGLFile, 1, TaxGLAccts(1)
'        If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)) <= 0 Then
'          GoTo NoGoodSkipP4
'        End If
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).AdvCRAcct)
        DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).AdvCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TCollectionPd#(2, T%) + TCollectionPd#(5, T%))
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).AdvDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).AdvDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TCollectionPd#(2, T%) + TCollectionPd#(5, T%))
        Dsc$ = "PaymentA"
        GoSub PostToGeneralJournal
        Close TXGLFile
      End If
      TXGLFile = FreeFile
      Open "TAXGLACT.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).AdvDBAcct)) <= 0 Then
'        GoTo NoGoodSkipP4
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).AdvDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).AdvDBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(TCollectionPd#(2, T%) + TCollectionPd#(5, T%))
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).AdvCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).AdvCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TCollectionPd#(2, T%) + TCollectionPd#(5, T%))
      Dsc$ = "PaymentA"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
NoGoodSkipP4:
      Close TXGLFile
    End If
    If TLateListPd#(2, T%) > 0 Or TLateListPd#(5, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXGLFile = FreeFile
        Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
        Get TXGLFile, 1, TaxGLAccts(1)
'        If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct)) <= 0 Then
'          GoTo NoGoodSkipP5
'        End If
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).LtLstCRAcct)
        DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).LtLstCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TLateListPd#(2, T%) + TLateListPd#(5, T%))
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TLateListPd#(2, T%) + TLateListPd#(5, T%))
        Dsc$ = "PaymentLL"
        GoSub PostToGeneralJournal
        Close TXGLFile
      End If
      TXGLFile = FreeFile
      Open "TAXGLACT.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct)) <= 0 Then
'        GoTo NoGoodSkipP5
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(TLateListPd#(2, T%) + TLateListPd#(5, T%))
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).LtLstCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).LtLstCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TLateListPd#(2, T%) + TLateListPd#(5, T%))
      Dsc$ = "PaymentLL"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
NoGoodSkipP5:
      Close TXGLFile
    End If
    If TRevOpt1Pd#(2, T%) > 0 Or TRevOpt1Pd#(5, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXGLFile = FreeFile
        Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
        Get TXGLFile, 1, TaxGLAccts(1)
'        If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct)) <= 0 Then
'          GoTo NoGoodSkipP6
'        End If
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt1CRAcct)
        DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt1CRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TRevOpt1Pd#(2, T%) + TRevOpt1Pd#(5, T%))
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TRevOpt1Pd#(2, T%) + TRevOpt1Pd#(5, T%))
        Dsc$ = "PaymentO1"
        GoSub PostToGeneralJournal
        Close TXGLFile
      End If
      TXGLFile = FreeFile
      Open "TAXGLACT.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct)) <= 0 Then
'        GoTo NoGoodSkipP6
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(TRevOpt1Pd#(2, T%) + TRevOpt1Pd#(5, T%))
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt1CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt1CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TRevOpt1Pd#(2, T%) + TRevOpt1Pd#(5, T%))
      Dsc$ = "PaymentO1"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
NoGoodSkipP6:
      Close TXGLFile
    End If
    If TRevOpt2Pd#(2, T%) > 0 Or TRevOpt2Pd#(5, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXGLFile = FreeFile
        Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
        Get TXGLFile, 1, TaxGLAccts(1)
'        If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct)) <= 0 Then
'          GoTo NoGoodSkipP7
'        End If
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt2CRAcct)
        DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt2CRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TRevOpt2Pd#(2, T%) + TRevOpt2Pd#(5, T%))
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TRevOpt2Pd#(2, T%) + TRevOpt2Pd#(5, T%))
        Dsc$ = "PaymentO2"
        GoSub PostToGeneralJournal
        Close TXGLFile
      End If
      TXGLFile = FreeFile
      Open "TAXGLACT.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct)) <= 0 Then
'        GoTo NoGoodSkipP7
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(TRevOpt2Pd#(2, T%) + TRevOpt2Pd#(5, T%))
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt2CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt2CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TRevOpt2Pd#(2, T%) + TRevOpt2Pd#(5, T%))
      Dsc$ = "PaymentO2"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
NoGoodSkipP7:
      Close TXGLFile
    End If
    If TRevOpt3Pd#(2, T%) > 0 Or TRevOpt3Pd#(5, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXGLFile = FreeFile
        Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
        Get TXGLFile, 1, TaxGLAccts(1)
'        If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct)) <= 0 Then
'          GoTo NoGoodSkipP8
'        End If
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt3CRAcct)
        DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt3CRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TRevOpt3Pd#(2, T%) + TRevOpt3Pd#(5, T%))
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TRevOpt3Pd#(2, T%) + TRevOpt3Pd#(5, T%))
        Dsc$ = "PaymentO3"
        GoSub PostToGeneralJournal
        Close TXGLFile
      End If
      TXGLFile = FreeFile
      Open "TAXGLACT.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct)) <= 0 Then
'        GoTo NoGoodSkipP8
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(TRevOpt3Pd#(2, T%) + TRevOpt3Pd#(5, T%))
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt3CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt3CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TRevOpt3Pd#(2, T%) + TRevOpt3Pd#(5, T%))
      Dsc$ = "PaymentO3"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
NoGoodSkipP8:
      Close TXGLFile
    End If
    If TPrePaidAmt#(8, T%) > 0 Or TPrePaidAmt#(5, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLACT.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len((QPTrim$(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct))) <= 0 Then
'        GoTo NoGoodSkipP9
'      End If
'      If Len((QPTrim$(TaxSetUpRec(1).OverPayGLNum))) <= 0 Then
'        GoTo NoGoodSkipP9
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxDBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(TPrePaidAmt#(8, T%) + TPrePaidAmt#(5, T%))
      ThisAcct = AcctFind(TaxSetuprec(1).OverPayGLNum)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxSetuprec(1).OverPayGLNum
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TPrePaidAmt#(8, T%) + TPrePaidAmt#(5, T%))
      Dsc$ = "PaymentPre"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
NoGoodSkipP9:
      Close TXGLFile
    End If
  Next T%
'@@@@@@@@@@@@@@@@@@end of payment types, 2, 21, 22
''''''''''''''''''''''''''''''''''
  'Overpayment Apply during billing -9 and -24(amts applied during up adj)
  For T% = 1 To 51
    If TPrePaidUsed#(9, T%) > 0 Or TPrePaidUsed#(15, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLACT.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len((QPTrim$(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct))) <= 0 Then
'        GoTo NoGoodSkipOV1
'      End If
'      If Len((QPTrim$(TaxSetUpRec(1).OverPayGLNum))) <= 0 Then
'        GoTo NoGoodSkipOV1
'      End If
      ThisAcct = AcctFind(TaxSetuprec(1).OverPayGLNum)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxSetuprec(1).OverPayGLNum
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(TPrePaidUsed#(9, T%) + TPrePaidUsed#(15, T%))
      Dsc$ = "OverpayUsed"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = DebitAcctRecord
      GJRec1(1).AcctNum = DebitAcctNumber$
      GJRec1(1).AcctName = DebitAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).DrAmt = DebitAmt#
      GJRec1(1).EType = "D"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      Put #GJFile, , GJRec1(1)

NoGoodSkipOV1:
      Close TXGLFile
    End If
    If TPrinciplePd#(9, T%) > 0 Or TPrinciplePd#(15, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXGLFile = FreeFile
        Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
        Get TXGLFile, 1, TaxGLAccts(1)
'        If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct)) <= 0 Then
'          GoTo NoGoodSkipOV2
'        End If
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).TaxCRAcct)
        DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TPrinciplePd#(9, T%) + TPrinciplePd#(15, T%))
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TPrinciplePd#(9, T%) + TPrinciplePd#(15, T%))
        Dsc$ = "OverPayApplyP"
        GoSub PostToGeneralJournal
        Close TXGLFile
      End If
      TXGLFile = FreeFile
      Open "TAXGLACT.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).TaxCRAcct)) <= 0 Then
'        GoTo NoGoodSkipOV2
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).TaxCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TPrinciplePd#(9, T%) + TPrinciplePd#(15, T%))
      Dsc$ = "OverPayApplyP"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = CreditAcctRecord
      GJRec1(1).AcctNum = CreditAcctNumber$
      GJRec1(1).AcctName = CreditAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = CreditAmt#
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      Put #GJFile, , GJRec1(1)

NoGoodSkipOV2:
      Close TXGLFile
    End If
    If TInterestPd#(9, T%) > 0 Or TInterestPd#(15, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXGLFile = FreeFile
        Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
        Get TXGLFile, 1, TaxGLAccts(1)
'        If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)) <= 0 Then
'          GoTo NoGoodSkipOV3
'        End If
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).IntCRAcct)
        DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TInterestPd#(9, T%) + TInterestPd#(15, T%))
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TInterestPd#(9, T%) + TInterestPd#(15, T%))
        Dsc$ = "OverPayApplyI"
        GoSub PostToGeneralJournal
        Close TXGLFile
      End If
      TXGLFile = FreeFile
      Open "TAXGLACT.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).IntCRAcct)) <= 0 Then
'        GoTo NoGoodSkipOV3
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).IntCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TInterestPd#(9, T%) + TInterestPd#(15, T%))
      Dsc$ = "OverPayApplyI"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = CreditAcctRecord
      GJRec1(1).AcctNum = CreditAcctNumber$
      GJRec1(1).AcctName = CreditAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = CreditAmt#
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      Put #GJFile, , GJRec1(1)

NoGoodSkipOV3:
      Close TXGLFile
    End If
    If TCollectionPd#(9, T%) > 0 Or TCollectionPd#(15, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXGLFile = FreeFile
        Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
        Get TXGLFile, 1, TaxGLAccts(1)
'        If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)) <= 0 Then
'          GoTo NoGoodSkipOV4
'        End If
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).AdvCRAcct)
        DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).AdvCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TCollectionPd#(9, T%) + TCollectionPd#(15, T%))
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).AdvDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).AdvDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TCollectionPd#(9, T%) + TCollectionPd#(15, T%))
        Dsc$ = "OverPayApplyA"
        GoSub PostToGeneralJournal
        Close TXGLFile
      End If
      TXGLFile = FreeFile
      Open "TAXGLACT.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).AdvCRAcct)) <= 0 Then
'        GoTo NoGoodSkipOV4
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).AdvCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).AdvCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TCollectionPd#(9, T%) + TCollectionPd#(15, T%))
      Dsc$ = "OverPayApplyA"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = CreditAcctRecord
      GJRec1(1).AcctNum = CreditAcctNumber$
      GJRec1(1).AcctName = CreditAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = CreditAmt#
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      Put #GJFile, , GJRec1(1)
NoGoodSkipOV4:
      Close TXGLFile
    End If
    If TLateListPd#(9, T%) > 0 Or TLateListPd#(15, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXGLFile = FreeFile
        Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
        Get TXGLFile, 1, TaxGLAccts(1)
'        If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct)) <= 0 Then
'          GoTo NoGoodSkipOV5
'        End If
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).LtLstCRAcct)
        DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).LtLstCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(TLateListPd#(9, T%) + TLateListPd#(15, T%))
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(TLateListPd#(9, T%) + TLateListPd#(15, T%))
        Dsc$ = "OverPayApplyLL"
        GoSub PostToGeneralJournal
        Close TXGLFile
      End If
      TXGLFile = FreeFile
      Open "TAXGLACT.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).LtLstCRAcct)) <= 0 Then
'        GoTo NoGoodSkipOV5
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).LtLstCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).LtLstCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TLateListPd#(9, T%) + TLateListPd#(15, T%))
      Dsc$ = "OverPayApplyLL"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = CreditAcctRecord
      GJRec1(1).AcctNum = CreditAcctNumber$
      GJRec1(1).AcctName = CreditAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = CreditAmt#
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      Put #GJFile, , GJRec1(1)
NoGoodSkipOV5:
      Close TXGLFile
    End If
    If TRevOpt1Pd#(9, T%) > 0 Or TRevOpt1Pd#(15, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXGLFile = FreeFile
        Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
        Get TXGLFile, 1, TaxGLAccts(1)
'        If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct)) <= 0 Then
'          GoTo NoGoodSkipOV6
'        End If
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt1CRAcct)
        DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt1CRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(TRevOpt1Pd#(9, T%) + TRevOpt1Pd#(15, T%))
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(TRevOpt1Pd#(9, T%) + TRevOpt1Pd#(15, T%))
        Dsc$ = "OverPayApplyO1"
        GoSub PostToGeneralJournal
        Close TXGLFile
      End If
      TXGLFile = FreeFile
      Open "TAXGLACT.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt1CRAcct)) <= 0 Then
'        GoTo NoGoodSkipOV6
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt1CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt1CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TRevOpt1Pd#(9, T%) + TRevOpt1Pd#(15, T%))
      Dsc$ = "OverPayApplyO1"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = CreditAcctRecord
      GJRec1(1).AcctNum = CreditAcctNumber$
      GJRec1(1).AcctName = CreditAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = CreditAmt#
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      Put #GJFile, , GJRec1(1)
NoGoodSkipOV6:
      Close TXGLFile
    End If
    If TRevOpt2Pd#(9, T%) > 0 Or TRevOpt2Pd#(15, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXGLFile = FreeFile
        Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
        Get TXGLFile, 1, TaxGLAccts(1)
'        If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct)) <= 0 Then
'          GoTo NoGoodSkipOV7
'        End If
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt2CRAcct)
        DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt2CRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(TRevOpt2Pd#(9, T%) + TRevOpt2Pd#(15, T%))
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(TRevOpt2Pd#(9, T%) + TRevOpt2Pd#(15, T%))
        Dsc$ = "OverPayApplyO2"
        GoSub PostToGeneralJournal
        Close TXGLFile
      End If
      TXGLFile = FreeFile
      Open "TAXGLACT.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt2CRAcct)) <= 0 Then
'        GoTo NoGoodSkipOV7
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt2CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt2CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TRevOpt2Pd#(9, T%) + TRevOpt2Pd#(15, T%))
      Dsc$ = "OverPayApplyO2"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = CreditAcctRecord
      GJRec1(1).AcctNum = CreditAcctNumber$
      GJRec1(1).AcctName = CreditAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = CreditAmt#
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      Put #GJFile, , GJRec1(1)
NoGoodSkipOV7:
      Close TXGLFile
    End If
    If TRevOpt3Pd#(9, T%) > 0 Or TRevOpt3Pd#(15, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXGLFile = FreeFile
        Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
        Get TXGLFile, 1, TaxGLAccts(1)
'        If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct)) <= 0 Then
'          GoTo NoGoodSkipOV8
'        End If
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt3CRAcct)
        DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt3CRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(TRevOpt3Pd#(9, T%) + TRevOpt3Pd#(15, T%))
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(TRevOpt3Pd#(9, T%) + TRevOpt3Pd#(15, T%))
        Dsc$ = "OverPayApplyO3"
        GoSub PostToGeneralJournal
        Close TXGLFile
      End If
      TXGLFile = FreeFile
      Open "TAXGLACT.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt3CRAcct)) <= 0 Then
'        GoTo NoGoodSkipOV8
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt3CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt3CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TRevOpt3Pd#(9, T%) + TRevOpt3Pd#(15, T%))
      Dsc$ = "OverPayApplyO3"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = CreditAcctRecord
      GJRec1(1).AcctNum = CreditAcctNumber$
      GJRec1(1).AcctName = CreditAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = CreditAmt#
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      Put #GJFile, , GJRec1(1)
NoGoodSkipOV8:
      Close TXGLFile
    End If
  Next T%
'@@@@@@@@@@@@@@ end of tran type 9 and apply prepay amt side of 24 up adj
''''''''''''''''''''''''''''
'tran 7 adjust pay down  and  10 paydown w/prep for interface 3(Modified Accrual) do bill charge again
  For T% = 1 To 51
    If TPrinciplePd#(7, T%) > 0 Or TPrinciplePd#(10, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXGLFile = FreeFile
        Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
        Get TXGLFile, 1, TaxGLAccts(1)
'        If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct)) <= 0 Then
'          GoTo NoGoodSkipAdjP1
'        End If
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct)
        DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxDBAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(TPrinciplePd#(7, T%) + TPrinciplePd#(10, T%))
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).TaxCRAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxCRAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(TPrinciplePd#(7, T%) + TPrinciplePd#(10, T%))
        Dsc$ = "AdjPayP"
        GoSub PostToGeneralJournal
        Close TXGLFile
      End If
      TXGLFile = FreeFile
      Open "TAXGLACT.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct)) <= 0 Then
'        GoTo NoGoodSkipAdjP1
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).TaxCRAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxCRAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round#(TPrinciplePd#(7, T%) + TPrinciplePd#(10, T%))
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TPrinciplePd#(7, T%) + TPrinciplePd#(10, T%))
      Dsc$ = "AdjPayP"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If

NoGoodSkipAdjP1:
      Close TXGLFile
    End If
    If TInterestPd#(7, T%) > 0 Or TInterestPd#(10, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXGLFile = FreeFile
        Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
        Get TXGLFile, 1, TaxGLAccts(1)
'        If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)) <= 0 Then
'          GoTo NoGoodSkipAdjP2
'        End If
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)
        DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntDBAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(TInterestPd#(7, T%) + TInterestPd#(10, T%))
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).IntCRAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntCRAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(TInterestPd#(7, T%) + TInterestPd#(10, T%))
        Dsc$ = "AdjPayI"
        GoSub PostToGeneralJournal
        Close TXGLFile
      End If
      TXGLFile = FreeFile
      Open "TAXGLACT.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)) <= 0 Then
'        GoTo NoGoodSkipAdjP2
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).IntCRAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntCRAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round#(TInterestPd#(7, T%) + TInterestPd#(10, T%))
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).IntDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TInterestPd#(7, T%) + TInterestPd#(10, T%))
      Dsc$ = "AdjPayI"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
NoGoodSkipAdjP2:
      Close TXGLFile
    End If
    If TCollectionPd#(7, T%) > 0 Or TCollectionPd#(10, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXGLFile = FreeFile
        Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
        Get TXGLFile, 1, TaxGLAccts(1)
'        If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).IntDBAcct)) <= 0 Then
'          GoTo NoGoodSkipAdjP3
'        End If
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).AdvDBAcct)
        DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).AdvDBAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(TCollectionPd#(7, T%) + TCollectionPd#(10, T%))
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).AdvCRAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).AdvCRAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(TCollectionPd#(7, T%) + TCollectionPd#(10, T%))
        Dsc$ = "AdjPayA"
        GoSub PostToGeneralJournal
        Close TXGLFile
      End If
      TXGLFile = FreeFile
      Open "TAXGLACT.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).AdvDBAcct)) <= 0 Then
'        GoTo NoGoodSkipAdjP3
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).AdvCRAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).AdvCRAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round#(TCollectionPd#(7, T%) + TCollectionPd#(10, T%))
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).AdvDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).AdvDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TCollectionPd#(7, T%) + TCollectionPd#(10, T%))
      Dsc$ = "AdjPayA"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
NoGoodSkipAdjP3:
      Close TXGLFile
    End If
    If TLateListPd#(7, T%) > 0 Or TLateListPd#(10, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXGLFile = FreeFile
        Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
        Get TXGLFile, 1, TaxGLAccts(1)
'        If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct)) <= 0 Then
'          GoTo NoGoodSkipAdjP4
'        End If
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct)
        DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(TLateListPd#(7, T%) + TLateListPd#(10, T%))
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).LtLstCRAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).LtLstCRAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(TLateListPd#(7, T%) + TLateListPd#(10, T%))
        Dsc$ = "AdjPayLL"
        GoSub PostToGeneralJournal
        Close TXGLFile
      End If
      TXGLFile = FreeFile
      Open "TAXGLACT.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct)) <= 0 Then
'        GoTo NoGoodSkipAdjP4
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).LtLstCRAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).LtLstCRAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round#(TLateListPd#(7, T%) + TLateListPd#(10, T%))
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).LtLstDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TLateListPd#(7, T%) + TLateListPd#(10, T%))
      Dsc$ = "AdjPayLL"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
NoGoodSkipAdjP4:
      Close TXGLFile
    End If
    If TRevOpt1Pd#(7, T%) > 0 Or TRevOpt1Pd#(10, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXGLFile = FreeFile
        Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
        Get TXGLFile, 1, TaxGLAccts(1)
'        If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct)) <= 0 Then
'          GoTo NoGoodSkipAdjP5
'        End If
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct)
        DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(TRevOpt1Pd#(7, T%) + TRevOpt1Pd#(10, T%))
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt1CRAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt1CRAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(TRevOpt1Pd#(7, T%) + TRevOpt1Pd#(10, T%))
        Dsc$ = "AdjPayO1"
        GoSub PostToGeneralJournal
        Close TXGLFile
      End If
      TXGLFile = FreeFile
      Open "TAXGLACT.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct)) <= 0 Then
'        GoTo NoGoodSkipAdjP5
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt1CRAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt1CRAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round#(TRevOpt1Pd#(7, T%) + TRevOpt1Pd#(10, T%))
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt1DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TRevOpt1Pd#(7, T%) + TRevOpt1Pd#(10, T%))
      Dsc$ = "AdjPayO1"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
NoGoodSkipAdjP5:
      Close TXGLFile
    End If
    If TRevOpt2Pd#(7, T%) > 0 Or TRevOpt2Pd#(10, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXGLFile = FreeFile
        Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
        Get TXGLFile, 1, TaxGLAccts(1)
'        If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct)) <= 0 Then
'          GoTo NoGoodSkipAdjP6
'        End If
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct)
        DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(TRevOpt2Pd#(7, T%) + TRevOpt2Pd#(10, T%))
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt2CRAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt2CRAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(TRevOpt2Pd#(7, T%) + TRevOpt2Pd#(10, T%))
        Dsc$ = "AdjPayO2"
        GoSub PostToGeneralJournal
        Close TXGLFile
      End If
      TXGLFile = FreeFile
      Open "TAXGLACT.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct)) <= 0 Then
'        GoTo NoGoodSkipAdjP6
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt2CRAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt2CRAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round#(TRevOpt2Pd#(7, T%) + TRevOpt2Pd#(10, T%))
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt2DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TRevOpt2Pd#(7, T%) + TRevOpt2Pd#(10, T%))
      Dsc$ = "AdjPayO2"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
NoGoodSkipAdjP6:
      Close TXGLFile
    End If
    If TRevOpt3Pd#(7, T%) > 0 Or TRevOpt3Pd#(10, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXGLFile = FreeFile
        Open "TAXGLBAC.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
        Get TXGLFile, 1, TaxGLAccts(1)
'        If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct)) <= 0 Then
'          GoTo NoGoodSkipAdjP7
'        End If
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct)
        DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(TRevOpt3Pd#(7, T%) + TRevOpt3Pd#(10, T%))
        ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt3CRAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt3CRAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(TRevOpt3Pd#(7, T%) + TRevOpt3Pd#(10, T%))
        Dsc$ = "AdjPayO3"
        GoSub PostToGeneralJournal
        Close TXGLFile
      End If
      TXGLFile = FreeFile
      Open "TAXGLACT.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len(QPTrim$(TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct)) <= 0 Then
'        GoTo NoGoodSkipAdjP7
'      End If
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt3CRAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt3CRAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round#(TRevOpt3Pd#(7, T%) + TRevOpt3Pd#(10, T%))
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).Opt3DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TRevOpt3Pd#(7, T%) + TRevOpt3Pd#(10, T%))
      Dsc$ = "AdjPayO3"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
NoGoodSkipAdjP7:
      Close TXGLFile
    End If
   
  '^*&^*&^*&^* 10 and 11 prepayment adj down
    If TPrePaidUsed#(10, T%) > 0 Or TPrePaidUsed#(11, T%) > 0 Then
      TXGLFile = FreeFile
      Open "TAXGLACT.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len((QPTrim$(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct))) <= 0 Then
'        GoTo NoGoodSkipP97
'      End If
'      If Len((QPTrim$(TaxSetUpRec(1).OverPayGLNum))) <= 0 Then
'        GoTo NoGoodSkipP97
'      End If
    ThisAcct = AcctFind(TaxSetuprec(1).OverPayGLNum)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxSetuprec(1).OverPayGLNum
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(TPrePaidUsed#(10, T%) + TPrePaidUsed#(11, T%))
      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TPrePaidUsed#(10, T%) + TPrePaidUsed#(11, T%))
      Dsc$ = "AdjPre"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
NoGoodSkipP97:
      Close TXGLFile
    End If
 Next T%
'@@@@@@@@@@@@ end of tran 7 10 and 11
''Tran 12 REfund Prepayment
'    If TPrePaidUsed#(12, T%) > 0 Then
'      TXGLFile = FreeFile
'      Open "TAXGLACT.DAT" For Random Shared As TXGLFile Len = TaxAcctRecLen
'      Get TXGLFile, 1, TaxGLAccts(1)
'      If Len((QPTrim$(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct))) <= 0 Then
'        GoTo NoGoodSkipP99
'      End If
'      If Len((QPTrim$(TaxSetUpRec(1).OverPayGLNum))) <= 0 Then
'        GoTo NoGoodSkipP99
'      End If
'      ThisAcct = AcctFind(TaxSetUpRec(1).OverPayGLNum)
'      DebitAcctRecord = ThisAcct
'      DebitAcctNumber$ = TaxSetUpRec(1).OverPayGLNum
'      If ThisAcct > 0 Then
'        DebitAcctName$ = GetAcctTitle$(ThisAcct)
'      Else
'        BadAcct = BadAcct + 1
'        DebitAcctName$ = "Undefined"
'      End If
'      DebitAmt# = TPrePaidUsed#(12, T%)
'      ThisAcct = AcctFind(TaxGLAccts(1).TaxAcct(T%).TaxDBAcct)
'      CreditAcctRecord = ThisAcct
'      CreditAcctNumber$ = TaxGLAccts(1).TaxAcct(T%).TaxDBAcct
'      If ThisAcct > 0 Then
'        CreditAcctName$ = GetAcctTitle$(ThisAcct)
'      Else
'        BadAcct = BadAcct + 1
'        CreditAcctName$ = "Undefined"
'      End If
'      CreditAmt# = TPrePaidUsed#(12, T%)
'      Dsc$ = "RefundPre"
'      GoSub PostToGeneralJournal
'NoGoodSkipP99:
'      Close TXGLFile
'    End If
'
''(*&((&(*& End of 12 Refund Pre


'  TaxYear       As Integer        'protected
'  TaxDBAcct     As String * 14
'  TaxCRAcct     As String * 14
'  IntDBAcct     As String * 14
'  IntCRAcct     As String * 14
'  AdvDBAcct     As String * 14
'  AdvCRAcct     As String * 14
'  Fill1         As String * 1     'protected
'  LtLstDBAcct   As String * 14
'  LtLstCRAcct   As String * 14
'  Opt1DBAcct    As String * 14
'  Opt1CRAcct    As String * 14
'  Opt2DBAcct    As String * 14
'  Opt2CRAcct    As String * 14
'  Opt3DBAcct    As String * 14
'  Opt3CRAcct    As String * 14
''TaxAcctsType
'  TaxAcct(1 To 51) As WinTAXGLAcctRecType
'
'
''TaxGLPrePayType
'  TaxDBAcct     As String * 14
'  TaxCRAcct     As String * 14

TXBunchReturn:
  Return

PostToGeneralJournal:
  'NOTE: Journal Rec 1 is the credit, Rec 2 is the debit
  ReDim GJRec1(1 To 2) As TrEditRecType
  GJRec1(1).AcctRec = CreditAcctRecord
  GJRec1(1).AcctNum = CreditAcctNumber$
  If Len(QPTrim$(CreditAcctNumber$)) > 0 Then
    GJRec1(1).AcctName = CreditAcctName$
  Else
    GJRec1(1).AcctName = "Blank Acct"
  End If
  GJRec1(1).TRDATE = WorkDate
  GJRec1(1).Ref = Ref$
  GJRec1(1).CrAmt = CreditAmt#
  GJRec1(1).EType = "C"
  GJRec1(1).Desc = "FRMTX " + Dsc$
  GJRec1(1).LDesc = "TX Interface"
  GJRec1(1).Src = "TX"
  Put #GJFile, , GJRec1(1)

  GJRec1(2).AcctRec = DebitAcctRecord
  GJRec1(2).AcctNum = DebitAcctNumber$
  If Len(QPTrim$(DebitAcctNumber$)) > 0 Then
    GJRec1(2).AcctName = DebitAcctName$
  Else
    GJRec1(2).AcctName = "Blank Acct"
  End If
  GJRec1(2).TRDATE = WorkDate
  GJRec1(2).Ref = Ref$
  GJRec1(2).DrAmt = DebitAmt#
  GJRec1(2).EType = "D"
  GJRec1(2).Desc = "FRMTX " + Dsc$
  GJRec1(2).LDesc = "TX Interface"
  GJRec1(2).Src = "TX"
  Put #GJFile, , GJRec1(2)
Return

Clearouttots:
For ppcnt = 1 To 15
  For RevCnt = 1 To 51
    TPrinciple#(ppcnt, RevCnt) = 0
    TInterest#(ppcnt, RevCnt) = 0
    TCollection#(ppcnt, RevCnt) = 0
    TPrinciplePd#(ppcnt, RevCnt) = 0
    TInterestPd#(ppcnt, RevCnt) = 0
    TCollectionPd#(ppcnt, RevCnt) = 0
    TRevOpt1#(ppcnt, RevCnt) = 0
    TRevOpt1Pd#(ppcnt, RevCnt) = 0
    TRevOpt2#(ppcnt, RevCnt) = 0
    TRevOpt2Pd#(ppcnt, RevCnt) = 0
    TRevOpt3#(ppcnt, RevCnt) = 0
    TRevOpt3Pd#(ppcnt, RevCnt) = 0
    TLateList#(ppcnt, RevCnt) = 0
    TLateListPd#(ppcnt, RevCnt) = 0
    TPrePaidAmt#(ppcnt, RevCnt) = 0
    TPrePaidUsed#(ppcnt, RevCnt) = 0
   Next
  Next
Return

TaxEnd:
  Exit Sub
  Return
End Sub
'New VA Tax vers 2.05
Private Sub ExtractVATX(ThruDate%)
  Dim txfile As Integer, InterfaceMethod As Integer, AcctMeth As String, TranFile As Integer
  Dim TaxTranRecLen As Integer, NumOfTRecs As Long, TCnt As Long, FoundCnt As Long
  Dim NGCnt As Long, P2S As String, cnt As Long
  Dim TaxTrans(1) As TaxVATransactionType
  TaxTranRecLen = Len(TaxTrans(1))
  
  ReDim TaxSetuprec(1) As TaxVAMasterType
  ReDim TranInfo(1) As TranRecInfoType
  FrmShowPctComp.Label1 = "Verifying Interface Setup"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me

  BadTxAcct = 0
  txfile = FreeFile
  Open "TAXSETUP.DAT" For Random As #txfile Len = Len(TaxSetuprec(1))
  If LOF(txfile) > 0 Then
    Get txfile, 1, TaxSetuprec(1)
  Else
    Unload FrmShowPctComp
    MsgBox "No Tax Setup File Information.", vbOKOnly, "No Setup"
    GoTo TaxEnd
  End If
  AcctMeth$ = QPTrim$(TaxSetuprec(1).AcctgMethod)
  If (Len(AcctMeth$) = 0) Then
    Unload FrmShowPctComp
    MsgBox "The Accounting Method Is Not Setup, Please Correct Before Trying Again.", vbOKOnly, "Invalid Tax Setup"
    GoTo TaxEnd
  End If
  Select Case AcctMeth$
  Case "C"
    InterfaceMethod = 1
  Case "A"
    InterfaceMethod = 2
  Case "M"
    InterfaceMethod = 3
  Case Else
    Unload FrmShowPctComp
    MsgBox "The Accounting Method Is Invalid, Please Correct Before Trying Again.", vbOKOnly, "Invalid Tax Setup"
    GoTo TaxEnd
  End Select
  Close txfile
  FrmShowPctComp.Label1 = "Creating Interface Trans File"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  FrmShowPctComp.ShowPctComp 15, 100
  TranFile = FreeFile
  Open "TAXTRANS.DAT" For Random Shared As TranFile Len = TaxTranRecLen
  NumOfTRecs& = LOF(TranFile) \ TaxTranRecLen
  For TCnt& = NumOfTRecs& To 1 Step -1
    Get #TranFile, TCnt&, TaxTrans(1)
    If Len(QPTrim$(TaxTrans(1).Posted2GL)) = 0 Or QPTrim$(TaxTrans(1).Posted2GL) = "N" Then
      If TaxTrans(1).TransDate <= ThruDate% Then
        FoundCnt = FoundCnt + 1
        ReDim Preserve TranInfo(FoundCnt) As TranRecInfoType
        TranInfo(FoundCnt).TranDate = TaxTrans(1).TransDate
        TranInfo(FoundCnt).TranRecNo = TCnt&
      End If
    Else
      NGCnt = NGCnt + 1
    End If
    P2S$ = Str$(FoundCnt)
    If NGCnt >= 2500 Then Exit For
  Next
  'FrmShowPctComp.ShowPctComp 40, 100
  If FoundCnt = 0 Then
    Close
    Unload FrmShowPctComp
    Call MainLog("No Tx to Grab " + Str$(FoundCnt) + " for " + fpDate)
    MsgBox "No Transactions Found to Interface.", vbOKOnly, "No Trans"
    GoTo TaxEnd
  End If
  FrmShowPctComp.Label1 = "Sorting Interface Trans File"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  FrmShowPctComp.ShowPctComp 25, 100
  SortTRec TranInfo(), FoundCnt      'sort'em by date. oldest first
''QSortTRec TranInfo(), FoundCnt
  If InterfaceMethod <> 1 Then
    ExtractVATXPersBill ThruDate%, TranInfo(), FoundCnt
  End If
  If BadTxAcct = 0 Then
    ExtractVATXPersPay ThruDate%, TranInfo(), FoundCnt
  End If
  If BadTxAcct = 0 Then
    ExtractVATXReal ThruDate%, TranInfo(), FoundCnt
  End If
  If BadTxAcct > 0 Then
    Close
    Unload FrmShowPctComp
    Call MainLog("Error - TX Grab Not Created Due to invalid or missing accts.")
    MsgBox "Errors Found, Interface File Not Created. Please Review Report.", vbOKOnly, "Errors"
    frmReportOpt.Show 1
    If rptopt = 1 Then
      frmGetDistMenu.PrnEditList 2
    ElseIf rptopt = 2 Then
      frmGetDistMenu.PrnEditList2 2
    End If
    KillFileD "GLTRXED.DAT"
    Exit Sub
  End If
  'transactions as interfaced
  TranFile = FreeFile
  Open "TAXTRANS.DAT" For Random Shared As TranFile Len = TaxTranRecLen
  NumOfTRecs& = LOF(TranFile) \ TaxTranRecLen
  FrmShowPctComp.Label1 = "Updating Transaction File"

  For cnt = 1 To FoundCnt
    FrmShowPctComp.ShowPctComp cnt, FoundCnt
    Get #TranFile, TranInfo(cnt).TranRecNo, TaxTrans(1)
    TaxTrans(1).Posted2GL = "Y"
    Put #TranFile, TranInfo(cnt).TranRecNo, TaxTrans(1)
  Next cnt
  Close
  Call MainLog("TX Grab Complete for VATaxes " + fpDate)
  
  MsgBox "Transaction Grab VATax Complete.", vbOKOnly, "Complete"
  GoTo TaxEnd

TaxEnd:
  Unload FrmShowPctComp
  Close
  Exit Sub
End Sub
Private Sub ExtractVATXPersBill(ThruDate%, TranInfo() As TranRecInfoType, FoundCnt)
  Dim Ref As String, Dash80 As String, P2S As String, TXPGLFile As Integer
  Dim GJReclen As Integer, RptFile As Integer, TaxPAcctRecLen  As Integer
  Dim TaxTranRecLen As Integer, NumOfTRecs As Long, TCnt As Long
  Dim NGCnt As Integer, GJFile As Integer, CDCashAcct As String
  Dim NumEdTrans As Integer, MCFile As Integer, cnt As Integer, CDCashAcctName As String
  Dim FirstTran As Integer, ThisDate As Integer, WorkDate As Integer
  Dim DayCount As Integer, LastTran As Integer, RevCnt As Integer, FundDue As String
  Dim FindCount As Integer, FundCnt As Integer, ThisAcct As Integer
  Dim Acct As String, AcctName As String, T As Integer, BadAcct As Integer
  Dim FoundFund As Integer, PCnt As Integer, Cash As Integer, CDCashRec As Long
  Dim txfile As Integer, InterfaceMethod As Integer, AcctMeth As String, CuryrR As Integer
  Dim TaxYear As Integer, MiddleRec As Integer, TranFile As Integer, DetPad As String
  Dim DebitAcctRecord As Integer, DebitAcctNumber As String, ppcnt As Integer
  Dim DebitAcctName As String, DebitAmt As Double, CreditAcctRecord As Integer
  Dim CreditAcctName As String, CreditAmt As Double, CreditAcctNumber As String
  Dim CDDueAcct As String, CDDueRec As Long, CDDueName As String, PadChars As Integer
  Dim Dsc As String, TrType As Integer, y As Integer, CuryrP As Integer, Curyr As Integer
  Dim GJrecnum As Long
  Dim GJInfo() As GJXferRecType
  Ref$ = "TX" + Left$(Date$, 2) + Mid$(Date$, 4, 2) + Right$(Date$, 2)
 ' ReDim TranInfo(1) As TranRecInfoType
  'these are for the personal
  Dim TPrinciple1#(16, 51), TPrinciple2#(16, 51), TPrinciple3#(16, 51), TPrinciple4#(16, 51)
  Dim TPrinciple5#(16, 51), TPenalty#(16, 51), TInterest#(16, 51)
  Dim TRevOpt1#(16, 51), TRevOpt2#(16, 51), TRevOpt3#(16, 51)
  ReDim Preserve GJInfo(1 To 3) As GJXferRecType
  ReDim TaxPGLAccts(1) As TaxPVAAcctsType
  TaxPAcctRecLen = Len(TaxPGLAccts(1))
  Dim TaxTrans(1) As TaxVATransactionType
  TaxTranRecLen = Len(TaxTrans(1))
  BadAcct = 0
  ReDim TaxSetuprec(1) As TaxVAMasterType
  txfile = FreeFile
  Open "TAXSETUP.DAT" For Random As #txfile Len = Len(TaxSetuprec(1))
  If LOF(txfile) > 0 Then
    Get txfile, 1, TaxSetuprec(1)
  Else
    Unload FrmShowPctComp
    MsgBox "No Tax Setup File Information.", vbOKOnly, "No Setup"
    GoTo TaxEnd
  End If
  'If Central Depository used then will need detail for acct #
    CDActive$ = QPTrim$(TaxSetuprec(1).CntrlDepYN)
    If CDActive$ = "Y" Then
      PadChars = GLDetLen - GLFundLen
      If PadChars > 0 Then
        DetPad$ = String(PadChars, "0")
      End If
      CDCashAcct$ = TaxSetuprec(1).CDCashGL
      CDCashAcct$ = QPStrip$(CDCashAcct$)
      CDCashAcct$ = FmtAcct$(CDCashAcct$, GLFundLen%, GLAcctLen%, GLDetLen%)
      CDCashAcct$ = QPTrim$(CDCashAcct$)
      CDDueAcct$ = QPTrim$(TaxSetuprec(1).CDSubGL)
      CDDueAcct$ = QPStrip$(CDDueAcct$)
      
      CDCashRec = AcctFind(CDCashAcct$)
      If CDCashRec <= 0 Then
        Unload FrmShowPctComp
        MsgBox "The Account for Central Cash Is Not Setup, Please Correct Before Trying Again.", vbOKOnly, "Invalid Tax Setup"
        GoTo TaxEnd
      Else
        CDCashAcctName$ = GetAcctTitle$(CDCashRec)
      End If
    End If
  FrmShowPctComp.ShowPctComp 10, 100
  AcctMeth$ = QPTrim$(TaxSetuprec(1).AcctgMethod)
  CuryrP = Right$(Num2Date$(TaxSetuprec(1).PTaxYear), 4)
  If (Len(AcctMeth$) = 0) Then
    Unload FrmShowPctComp
    MsgBox "The Accounting Method Is Not Setup, Please Correct Before Trying Again.", vbOKOnly, "Invalid Tax Setup"
    GoTo TaxEnd
  End If
  Select Case AcctMeth$
  Case "C"
    InterfaceMethod = 1
  Case "A"
    InterfaceMethod = 2
  Case "M"
    InterfaceMethod = 3
  Case Else
    Unload FrmShowPctComp
    MsgBox "The Accounting Method Is Invalid, Please Correct Before Trying Again.", vbOKOnly, "Invalid Tax Setup"
    GoTo EndTax
  End Select
  Close txfile
  GJReclen = Len(GJRec(1))
'  FrmShowPctComp.Label1 = "Sorting Interface Trans File"
'  FrmShowPctComp.cmdCancel.Enabled = False
'  FrmShowPctComp.Show , Me
'  FrmShowPctComp.ShowPctComp 25, 100
  TranFile = FreeFile
  Open "TAXTRANS.DAT" For Random Shared As TranFile Len = TaxTranRecLen
  NumOfTRecs& = LOF(TranFile) \ TaxTranRecLen
'  For TCnt& = NumOfTRecs& To 1 Step -1
'    Get #TranFile, TCnt&, TaxTrans(1)
'    If Len(QPTrim$(TaxTrans(1).Posted2GL)) = 0 Or QPTrim$(TaxTrans(1).Posted2GL) = "N" Then
'      If TaxTrans(1).BillType = "P" Then
'        Select Case TaxTrans(1).TranType
'          Case 1, 3, 4, 5, 13, 14:
'            'Store trans rec numbers and dates in array
'            If TaxTrans(1).TransDate <= ThruDate% Then
'              FoundCnt = FoundCnt + 1
'              ReDim Preserve TranInfo(FoundCnt) As TranRecInfoType
'              TranInfo(FoundCnt).TranDate = TaxTrans(1).TransDate
'              TranInfo(FoundCnt).TranRecNo = TCnt&
'            End If
'          Case Else
'        End Select
'      End If
'    Else
'      NGCnt = NGCnt + 1
'    End If
'    P2S$ = Str$(FoundCnt)
'    'QPrintRC P2S$, 13, 47, Cnf.HiLite
'    'SmallPause
'    If NGCnt >= 2500 Then Exit For
'  Next
'  'FrmShowPctComp.ShowPctComp 40, 100
'  If FoundCnt = 0 Then
'    Close
'    Unload FrmShowPctComp
'    Call MainLog("No Tx to Grab " + Str$(FoundCnt) + " for " + fpDate)
'    MsgBox "No Transactions Found to Interface.", vbOKOnly, "No Trans"
'    GoTo EndTax
'  End If
'  FrmShowPctComp.Label1 = "Sorting Interface Trans File"
'  FrmShowPctComp.cmdCancel.Enabled = False
'  FrmShowPctComp.Show , Me
'  FrmShowPctComp.ShowPctComp 15, 100
'  SortTRec TranInfo(), FoundCnt      'sort'em by date. oldest first

  GJFile = FreeFile
  Open "GLTRXED.DAT" For Random Access Read Write Shared As GJFile Len = GJReclen
  GJrecnum = LOF(GJFile) \ GJReclen
  FrmShowPctComp.ShowPctComp 45, 100
  FirstTran = 1
  ThisDate = TranInfo(1).TranDate
  For cnt = 1 To FoundCnt
   ' FrmShowPctComp.ShowPctComp cnt, FoundCnt
    If ThisDate <> TranInfo(cnt).TranDate Then
      ThisDate = TranInfo(cnt).TranDate
      LastTran = cnt - 1
      GoSub ProcessThisBunchTX
      FirstTran = cnt
    End If
  Next cnt
  FrmShowPctComp.Label1 = "Tag Interface Transactions"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me

  LastTran = FoundCnt
  GoSub ProcessThisBunchTX
  BadTxAcct = BadAcct

'  If BadAcct > 0 Then
'    Close
'    Unload FrmShowPctComp
'    Call MainLog("Error - TX Grab Not Created Due to invalid or missing accts.")
'    MsgBox "Errors Found, Interface File Not Created. Please Review Report.", vbOKOnly, "Errors"
'    frmReportOpt.Show 1
'    If rptopt = 1 Then
'      frmGetDistMenu.PrnEditList 2
'    ElseIf rptopt = 2 Then
'      frmGetDistMenu.PrnEditList2 2
'    End If
'    Kill "GLTRXED.DAT"
'    Exit Sub
'  End If
'  'transactions as interfaced
'
'  For cnt = 1 To Foundcnt
'    FrmShowPctComp.ShowPctComp cnt, Foundcnt
'    Get #TranFile, TranInfo(cnt).TranRecNo, TaxTrans(1)
'    TaxTrans(1).Posted2GL = "Y"
'    Put #TranFile, TranInfo(cnt).TranRecNo, TaxTrans(1)
'  Next cnt
'  Close
'  'SLEEP 2
'  Call MainLog("TX Grab Complete for PersBill " + fpDate)
'
'  MsgBox "Transaction Grab Pers/Bill Complete.", vbOKOnly, "Complete"
'  GoTo EndTax
'
EndTax:
  Unload FrmShowPctComp
  Close
  Exit Sub

ProcessThisBunchTX:      'Initialize for This Set
GoSub Clearouttots
  For PCnt = FirstTran To LastTran
    If PCnt = FirstTran Then
      WorkDate = TranInfo(PCnt).TranDate
    End If
    Get #TranFile, TranInfo(PCnt).TranRecNo, TaxTrans(1)
    'Now Decipher by Type and Year
    Select Case TaxTrans(1).TranType
      Case 1:
        TrType = 1
      Case 2:
        TrType = 2
      Case 3:
        TrType = 3
      Case 4:
        TrType = 4
      Case 5:
        TrType = 5
      Case 6:
        TrType = 6
      Case 7:
        TrType = 7
      Case 9:
        TrType = 9
      Case 10:
        TrType = 10
      Case 11:
        TrType = 11
      Case 12:
        TrType = 12
      Case 13:
        TrType = 13
      Case 14:
        TrType = 14
      Case 21:
        TrType = 16
      Case 22:
        TrType = 8
      Case 24:
        TrType = 15
      End Select
      If TaxTrans(1).BillType = "P" Then
        TaxYear = TaxTrans(1).TaxYear
        Curyr = CuryrP
        If TaxYear < 1 Then TaxYear = Curyr
        TaxYear = TaxYear - 1979
        'Reduce Based on 1980 being = 1
        TPrinciple1#(TrType, TaxYear) = TPrinciple1#(TrType, TaxYear) + TaxTrans(1).Revenue.Principle1
        TPrinciple2#(TrType, TaxYear) = TPrinciple2#(TrType, TaxYear) + TaxTrans(1).Revenue.Principle2
        TPrinciple3#(TrType, TaxYear) = TPrinciple3#(TrType, TaxYear) + TaxTrans(1).Revenue.Principle3
        TPrinciple4#(TrType, TaxYear) = TPrinciple4#(TrType, TaxYear) + TaxTrans(1).Revenue.Principle4
        TPrinciple5#(TrType, TaxYear) = TPrinciple5#(TrType, TaxYear) + TaxTrans(1).Revenue.Principle5
        TInterest#(TrType, TaxYear) = TInterest#(TrType, TaxYear) + TaxTrans(1).Revenue.Interest
        TPenalty#(TrType, TaxYear) = TPenalty#(TrType, TaxYear) + TaxTrans(1).Revenue.Penalty
        TRevOpt1#(TrType, TaxYear) = TRevOpt1#(TrType, TaxYear) + TaxTrans(1).Revenue.RevOpt1
        TRevOpt2#(TrType, TaxYear) = TRevOpt2#(TrType, TaxYear) + TaxTrans(1).Revenue.RevOpt2
        TRevOpt3#(TrType, TaxYear) = TRevOpt3#(TrType, TaxYear) + TaxTrans(1).Revenue.RevOpt3
      End If
  Next PCnt
If InterfaceMethod <> 1 Then  ' 1 is cash so skip all charges if cash
    'Now Post for bill trans and move on to next type
  'TranType1 Billing
  For T% = 1 To 51
    'these are for the personal
    If TPrinciple1#(1, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PersDBAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PersDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TPrinciple1#(1, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PersCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PersCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Invalid Acct"
      End If
      CreditAmt# = TPrinciple1#(1, T%)
      Dsc$ = "Bills"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TPrinciple2#(1, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MTDBAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MTDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TPrinciple2#(1, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MTCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MTCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Invalid Acct"
      End If
      CreditAmt# = TPrinciple2#(1, T%)
      Dsc$ = "Bills"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TPrinciple3#(1, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MCDBAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MCDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TPrinciple3#(1, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MCCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MCCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Invalid Acct"
      End If
      CreditAmt# = TPrinciple3#(1, T%)
      Dsc$ = "Bills"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TPrinciple4#(1, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).FEDBAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).FEDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TPrinciple4#(1, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).FECRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).FECRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Invalid Acct"
      End If
      CreditAmt# = TPrinciple4#(1, T%)
      Dsc$ = "Bills"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TPrinciple5#(1, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MHDBAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MHDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TPrinciple5#(1, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MHCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MHCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Invalid Acct"
      End If
      CreditAmt# = TPrinciple5#(1, T%)
      Dsc$ = "Bills"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TRevOpt1#(1, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt1DBAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt1DBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TRevOpt1#(1, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt1CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt1CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Invaild Acct"
      End If
      CreditAmt# = TRevOpt1#(1, T%)
      Dsc$ = "Bills"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TRevOpt2#(1, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt2DBAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt2DBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TRevOpt2#(1, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt2CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt2CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Invaild Acct"
      End If
      CreditAmt# = TRevOpt2#(1, T%)
      Dsc$ = "Bills"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TRevOpt3#(1, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt3DBAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt3DBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TRevOpt3#(1, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt3CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt3CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Invaild Acct"
      End If
      CreditAmt# = TRevOpt3#(1, T%)
      Dsc$ = "Bills"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
  Next T%
  '@@@@@@@@@@@@@@@@end of billing trantype-1
  'Interest charge  tran-4
  For T% = 1 To 51
    If TInterest#(4, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).IntDBAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).IntDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TInterest#(4, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).IntCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).IntCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TInterest#(4, T%)
      Dsc$ = "Interest"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
   Next T%
  '@@@@@@@@@@end of tran type 5 Penalty charge
    'Penalty charge  tran-5
  For T% = 1 To 51
    If TPenalty#(5, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PenDBAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PenDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TPenalty#(5, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PenCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PenCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TPenalty#(5, T%)
      Dsc$ = "Penalty"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
  Next T%
  '@@@@@@@@@@end of tran type 5 Penalty charge

  ''''''''''''''''''''''''''''''''''''''''''''''
  'Ad/Collection charge tran 6
'not for pers
'@@@@@@@@@@@@@@@@@@@@end of tran type 6 ad charge(collection)
''''''''''''
'Up Adj  14, 24
  For T% = 14 To 51
    If TPrinciple1#(14, T%) > 0 Or TPrinciple1#(15, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PersDBAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PersDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = Round#(TPrinciple1#(14, T%) + TPrinciple1#(15, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PersCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PersCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TPrinciple1#(14, T%) + TPrinciple1#(15, T%))
      Dsc$ = "UpAdjP"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TPrinciple2#(14, T%) > 0 Or TPrinciple2#(15, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MTDBAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MTDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = Round#(TPrinciple2#(14, T%) + TPrinciple2#(15, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MTCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MTCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TPrinciple2#(14, T%) + TPrinciple2#(15, T%))
      Dsc$ = "UpAdjP"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TPrinciple3#(14, T%) > 0 Or TPrinciple3#(15, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MCDBAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MCDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = Round#(TPrinciple3#(14, T%) + TPrinciple3#(15, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MCCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MCCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TPrinciple3#(14, T%) + TPrinciple3#(15, T%))
      Dsc$ = "UpAdjP"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TPrinciple4#(14, T%) > 0 Or TPrinciple4#(15, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).FEDBAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).FEDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = Round#(TPrinciple4#(14, T%) + TPrinciple4#(15, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).FECRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).FECRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TPrinciple4#(14, T%) + TPrinciple4#(15, T%))
      Dsc$ = "UpAdjP"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TPrinciple5#(14, T%) > 0 Or TPrinciple5#(15, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MHDBAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MHDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = Round#(TPrinciple5#(14, T%) + TPrinciple5#(15, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MHCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MHCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TPrinciple5#(14, T%) + TPrinciple5#(15, T%))
      Dsc$ = "UpAdjP"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TRevOpt1#(14, T%) > 0 Or TRevOpt1#(15, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt1DBAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt1DBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = Round#(TRevOpt1#(14, T%) + TRevOpt1#(15, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt1CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt1CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TRevOpt1#(14, T%) + TRevOpt1#(15, T%))
      Dsc$ = "UpAdjOR1"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TRevOpt2#(14, T%) > 0 Or TRevOpt2#(15, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt2DBAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt2DBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = Round#(TRevOpt2#(14, T%) + TRevOpt2#(15, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt2CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt2CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TRevOpt2#(14, T%) + TRevOpt2#(15, T%))
      Dsc$ = "UpAdjOR2"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TRevOpt3#(14, T%) > 0 Or TRevOpt3#(15, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt3DBAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt3DBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = Round#(TRevOpt3#(14, T%) + TRevOpt3#(15, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt3CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt3CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TRevOpt3#(14, T%) + TRevOpt3#(15, T%))
      Dsc$ = "UpAdjOR3"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TInterest#(14, T%) > 0 Or TInterest#(15, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).IntDBAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).IntDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = Round#(TInterest#(14, T%) + TInterest#(15, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).IntCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).IntCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TInterest#(14, T%) + TInterest#(15, T%))
      Dsc$ = "UpAdjInt"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TPenalty#(14, T%) > 0 Or TPenalty#(15, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PenDBAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PenDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = Round#(TPenalty#(14, T%) + TPenalty#(15, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PenCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PenCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TPenalty#(14, T%) + TPenalty#(15, T%))
      Dsc$ = "UpAdjPen"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
  Next T%
'end of up adj  14, 24
'tran type 13-Adj bill down
  For T% = 1 To 51
    If TPrinciple1#(13, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PersCRAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PersCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TPrinciple1#(13, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PersDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PersDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TPrinciple1#(13, T%)
      Dsc$ = "AdjBillDown"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TPrinciple2#(13, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MTCRAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MTCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TPrinciple2#(13, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MTDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MTDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TPrinciple2#(13, T%)
      Dsc$ = "AdjBillDown"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TPrinciple3#(13, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MCCRAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MCCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TPrinciple3#(13, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MCDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MCDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TPrinciple3#(13, T%)
      Dsc$ = "AdjBillDown"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TPrinciple4#(13, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).FECRAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).FECRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TPrinciple4#(13, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).FEDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).FEDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TPrinciple4#(13, T%)
      Dsc$ = "AdjBillDown"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TPrinciple5#(13, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MHCRAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MHCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TPrinciple5#(13, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MHDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MHDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TPrinciple5#(13, T%)
      Dsc$ = "AdjBillDown"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TRevOpt1#(13, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt1CRAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt1CRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TRevOpt1#(13, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt1DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt1DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TRevOpt1#(13, T%)
      Dsc$ = "AdjBillDown"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TRevOpt2#(13, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt2CRAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt2CRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TRevOpt2#(13, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt2DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt2DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TRevOpt2#(13, T%)
      Dsc$ = "AdjBillDown"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TRevOpt3#(13, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt3CRAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt3CRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TRevOpt3#(13, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt3DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt3DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TRevOpt3#(13, T%)
      Dsc$ = "AdjBillDown"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
  'Interest adjdownforbill
    If TInterest#(13, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).IntCRAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).IntCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TInterest#(13, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).IntDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).IntDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TInterest#(13, T%)
      Dsc$ = "AdjBillDown"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TPenalty#(13, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PenCRAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PenCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TPenalty#(13, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PenDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PenDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TPenalty#(13, T%)
      Dsc$ = "AdjBillDown"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
  Next T%
End If
'''end of tran type 13
''also end of charges ups/downs
'''''''''''''''''''''''''
TXBunchReturn:
  Return

PostToGeneralJournal:
  'NOTE: Journal Rec 1 is the credit, Rec 2 is the debit
  ReDim GJRec1(1 To 2) As TrEditRecType
  GJRec1(1).AcctRec = CreditAcctRecord
  GJRec1(1).AcctNum = CreditAcctNumber$
  If Len(QPTrim$(CreditAcctNumber$)) > 0 Then
    GJRec1(1).AcctName = CreditAcctName$
  Else
    GJRec1(1).AcctName = "Blank Acct"
  End If
  GJRec1(1).TRDATE = WorkDate
  GJRec1(1).Ref = Ref$
  GJRec1(1).CrAmt = CreditAmt#
  GJRec1(1).EType = "C"
  GJRec1(1).Desc = "FRMTX " + Dsc$
  GJRec1(1).LDesc = "TX Interface"
  GJRec1(1).Src = "TX"
  GJrecnum = LOF(GJFile) \ GJReclen
  Put #GJFile, GJrecnum + 1, GJRec1(1)

  GJRec1(2).AcctRec = DebitAcctRecord
  GJRec1(2).AcctNum = DebitAcctNumber$
  If Len(QPTrim$(DebitAcctNumber$)) > 0 Then
    GJRec1(2).AcctName = DebitAcctName$
  Else
    GJRec1(2).AcctName = "Blank Acct"
  End If
  GJRec1(2).TRDATE = WorkDate
  GJRec1(2).Ref = Ref$
  GJRec1(2).DrAmt = DebitAmt#
  GJRec1(2).EType = "D"
  GJRec1(2).Desc = "FRMTX " + Dsc$
  GJRec1(2).LDesc = "TX Interface"
  GJRec1(2).Src = "TX"
  GJrecnum = LOF(GJFile) \ GJReclen
  Put #GJFile, GJrecnum + 1, GJRec1(2)
Return

Clearouttots:
For ppcnt = 1 To 16
  For RevCnt = 1 To 51
    TPrinciple1#(ppcnt, RevCnt) = 0
    TPrinciple2#(ppcnt, RevCnt) = 0
    TPrinciple3#(ppcnt, RevCnt) = 0
    TPrinciple4#(ppcnt, RevCnt) = 0
    TPrinciple5#(ppcnt, RevCnt) = 0
    TInterest#(ppcnt, RevCnt) = 0
    TPenalty#(ppcnt, RevCnt) = 0
    TRevOpt1#(ppcnt, RevCnt) = 0
    TRevOpt2#(ppcnt, RevCnt) = 0
    TRevOpt3#(ppcnt, RevCnt) = 0
   Next
  Next
Return
TaxEnd:
  Exit Sub
  Return
End Sub
Private Sub ExtractVATXReal(ThruDate%, TranInfo() As TranRecInfoType, FoundCnt)
  Dim Ref As String, Dash80 As String, P2S As String
  Dim GJReclen As Integer, RptFile As Integer
  Dim TaxTranRecLen As Integer, NumOfTRecs As Long, TCnt As Long, TaxRAcctRecLen As Integer
  Dim NGCnt As Integer, GJFile As Integer, CDCashAcct As String
  Dim NumEdTrans As Integer, MCFile As Integer, cnt As Integer, CDCashAcctName As String
  Dim FirstTran As Integer, ThisDate As Integer, WorkDate As Integer, TXRGLFile As Integer
  Dim DayCount As Integer, LastTran As Integer, RevCnt As Integer, FundDue As String
  Dim FindCount As Integer, FundCnt As Integer, ThisAcct As Integer
  Dim Acct As String, AcctName As String, T As Integer, BadAcct As Integer
  Dim FoundFund As Integer, PCnt As Integer, Cash As Integer, CDCashRec As Long
  Dim txfile As Integer, InterfaceMethod As Integer, AcctMeth As String, CuryrR As Integer
  Dim TaxYear As Integer, MiddleRec As Integer, TranFile As Integer, DetPad As String
  Dim DebitAcctRecord As Integer, DebitAcctNumber As String, ppcnt As Integer
  Dim DebitAcctName As String, DebitAmt As Double, CreditAcctRecord As Integer
  Dim CreditAcctName As String, CreditAmt As Double, CreditAcctNumber As String
  Dim CDDueAcct As String, CDDueRec As Long, CDDueName As String, PadChars As Integer
  Dim Dsc As String, TrType As Integer, y As Integer, Curyr As Integer
  Dim GJrecnum As Long
  Dim GJInfo() As GJXferRecType
  Ref$ = "TX" + Left$(Date$, 2) + Mid$(Date$, 4, 2) + Right$(Date$, 2)
 ' ReDim TranInfo(1) As TranRecInfoType
  'these are for the personal
  'these are for the real
  Dim RTPrinciple1#(16, 51), RTInterest#(16, 51), RTCollection#(16, 51)
  Dim RTPrinciple1Pd#(16, 51), RTInterestPd#(16, 51), RTCollectionPd#(16, 51)
  Dim RTPenalty#(16, 51), RTPenaltyPd#(16, 51), RTLateList#(16, 51), RTLateListPd#(16, 51)
  Dim RTRevOpt1#(16, 51), RTRevOpt1Pd#(16, 51), RTRevOpt2#(16, 51), RTRevOpt2Pd#(16, 51)
  Dim RTRevOpt3#(16, 51), RTRevOpt3Pd#(16, 51), RTPrePaidAmt#(16, 51), RTPrePaidUsed#(16, 51)
  
  ReDim Preserve GJInfo(1 To 3) As GJXferRecType
        
  ReDim TaxRGLAccts(1) As TaxRVAAcctsType
  TaxRAcctRecLen = Len(TaxRGLAccts(1))

  Dim TaxTrans(1) As TaxVATransactionType
  TaxTranRecLen = Len(TaxTrans(1))
  BadAcct = 0
  ReDim TaxSetuprec(1) As TaxVAMasterType
  txfile = FreeFile
  Open "TAXSETUP.DAT" For Random As #txfile Len = Len(TaxSetuprec(1))
  If LOF(txfile) > 0 Then
    Get txfile, 1, TaxSetuprec(1)
  Else
    Unload FrmShowPctComp
    MsgBox "No Tax Setup File Information.", vbOKOnly, "No Setup"
    GoTo TaxEnd
  End If
  'If Central Depository used then will need detail for acct #
    CDActive$ = QPTrim$(TaxSetuprec(1).CntrlDepYN)
    If CDActive$ = "Y" Then
      PadChars = GLDetLen - GLFundLen
      If PadChars > 0 Then
        DetPad$ = String(PadChars, "0")
      End If
      CDCashAcct$ = TaxSetuprec(1).CDCashGL
      CDCashAcct$ = QPStrip$(CDCashAcct$)
      CDCashAcct$ = FmtAcct$(CDCashAcct$, GLFundLen%, GLAcctLen%, GLDetLen%)
      CDCashAcct$ = QPTrim$(CDCashAcct$)
      CDDueAcct$ = QPTrim$(TaxSetuprec(1).CDSubGL)
      CDDueAcct$ = QPStrip$(CDDueAcct$)
      
      CDCashRec = AcctFind(CDCashAcct$)
      If CDCashRec <= 0 Then
        Unload FrmShowPctComp
        MsgBox "The Account for Central Cash Is Not Setup, Please Correct Before Trying Again.", vbOKOnly, "Invalid Tax Setup"
        GoTo TaxEnd
      Else
        CDCashAcctName$ = GetAcctTitle$(CDCashRec)
      End If
    End If
  FrmShowPctComp.ShowPctComp 80, 100
  AcctMeth$ = QPTrim$(TaxSetuprec(1).AcctgMethod)
  
  CuryrR = Right$(Num2Date$(TaxSetuprec(1).RTaxYear), 4)
  If (Len(AcctMeth$) = 0) Then
    Unload FrmShowPctComp
    MsgBox "The Accounting Method Is Not Setup, Please Correct Before Trying Again.", vbOKOnly, "Invalid Tax Setup"
    GoTo TaxEnd
  End If
  Select Case AcctMeth$
  Case "C"
    InterfaceMethod = 1
  Case "A"
    InterfaceMethod = 2
  Case "M"
    InterfaceMethod = 3
  Case Else
    Unload FrmShowPctComp
    MsgBox "The Accounting Method Is Invalid, Please Correct Before Trying Again.", vbOKOnly, "Invalid Tax Setup"
    GoTo EndTax
  End Select
  Close txfile
  
  GJReclen = Len(GJRec(1))
  
'  If Exist("TAXPGLACT.DAT") Then
'    TXPGLFile = FreeFile
'    Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
'    Get TXPGLFile, 1, TaxPGLAccts(1)
'  Else
'    Unload FrmShowPctComp
'    MsgBox "Tax Accounts Not Setup,Interface File Not Created.", vbOKOnly, "Tax Acct Setup Invalid"
'    GoTo EndTax
'  Close
'  End If
'  FrmShowPctComp.Label1 = "Sorting Interface Trans File"
'  FrmShowPctComp.cmdCancel.Enabled = False
'  FrmShowPctComp.Show , Me

  'Now Process the Transactions
  'ClearBox

 ' QPrintRC "Searching Cash Transactions.", 11, 26, 126
 ' QPrintRC "New Transactions:", 13, 29, Cnf.HiLite
'  FrmShowPctComp.ShowPctComp 25, 100
  TranFile = FreeFile
  Open "TAXTRANS.DAT" For Random Shared As TranFile Len = TaxTranRecLen
  NumOfTRecs& = LOF(TranFile) \ TaxTranRecLen
'  For TCnt& = NumOfTRecs& To 1 Step -1
'    Get #TranFile, TCnt&, TaxTrans(1)
'    If Len(QPTrim$(TaxTrans(1).Posted2GL)) = 0 Or QPTrim$(TaxTrans(1).Posted2GL) = "N" And TaxTrans(1).BillType = "R" Then
'      'Store trans rec numbers and dates in array
'      If TaxTrans(1).TransDate <= ThruDate% Then
'        FoundCnt = FoundCnt + 1
'        ReDim Preserve TranInfo(FoundCnt) As TranRecInfoType
'        TranInfo(FoundCnt).TranDate = TaxTrans(1).TransDate
'        TranInfo(FoundCnt).TranRecNo = TCnt&
'      End If
'    Else
'      NGCnt = NGCnt + 1
'    End If
'    P2S$ = Str$(FoundCnt)
'    'QPrintRC P2S$, 13, 47, Cnf.HiLite
'    'SmallPause
'    If NGCnt >= 2500 Then Exit For
'  Next
'  'FrmShowPctComp.ShowPctComp 40, 100
'  If FoundCnt = 0 Then
'    Close
'    Unload FrmShowPctComp
'    Call MainLog("No Tx to Grab " + Str$(FoundCnt) + " for " + fpdate)
'    MsgBox "No Transactions Found to Interface.", vbOKOnly, "No Trans"
'    GoTo EndTax
'  End If
'  FrmShowPctComp.Label1 = "Sorting Interface Trans File"
'  FrmShowPctComp.cmdCancel.Enabled = False
'  FrmShowPctComp.Show , Me
'  FrmShowPctComp.ShowPctComp 15, 100
'  SortTRec TranInfo(), FoundCnt      'sort'em by date. oldest first

  GJFile = FreeFile
  Open "GLTRXED.DAT" For Random Access Read Write Shared As GJFile Len = GJReclen
  GJrecnum = LOF(GJFile) \ GJReclen
  FrmShowPctComp.ShowPctComp 95, 100
  FirstTran = 1
  ThisDate = TranInfo(1).TranDate
  For cnt = 1 To FoundCnt
    FrmShowPctComp.ShowPctComp cnt, FoundCnt
    If ThisDate <> TranInfo(cnt).TranDate Then
      ThisDate = TranInfo(cnt).TranDate
      LastTran = cnt - 1
      GoSub ProcessThisBunchTX
      FirstTran = cnt
    End If
  Next cnt
  FrmShowPctComp.Label1 = "Tag Interface Transactions"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me

  LastTran = FoundCnt
  GoSub ProcessThisBunchTX
  BadTxAcct = BadAcct
'  If BadAcct > 0 Then
'    Close
'    Unload FrmShowPctComp
'    Call MainLog("Error - TX Grab Not Created Due to invalid or missing accts.")
'    MsgBox "Errors Found, Interface File Not Created. Please Review Report.", vbOKOnly, "Errors"
'    frmReportOpt.Show 1
'    If rptopt = 1 Then
'      frmGetDistMenu.PrnEditList 2
'    ElseIf rptopt = 2 Then
'      frmGetDistMenu.PrnEditList2 2
'    End If
'    Kill "GLTRXED.DAT"
'    Exit Sub
'  End If
'  'transactions as interfaced
'
'  For cnt = 1 To FoundCnt
'    FrmShowPctComp.ShowPctComp cnt, FoundCnt
'    Get #TranFile, TranInfo(cnt).TranRecNo, TaxTrans(1)
'    TaxTrans(1).Posted2GL = "Y"
'    Put #TranFile, TranInfo(cnt).TranRecNo, TaxTrans(1)
'  Next cnt
'  Close
'  'SLEEP 2
'  Call MainLog("TX Grab Complete for " + fpdate)
'  MsgBox "Transaction Grab Complete.", vbOKOnly, "Complete"
'  GoTo EndTax

EndTax:
  Unload FrmShowPctComp
  Close
  Exit Sub

ProcessThisBunchTX:      'Initialize for This Set
GoSub Clearouttots
  For PCnt = FirstTran To LastTran
    If PCnt = FirstTran Then
      WorkDate = TranInfo(PCnt).TranDate
    End If
    Get #TranFile, TranInfo(PCnt).TranRecNo, TaxTrans(1)
    'Now Decipher by Type and Year
    Select Case TaxTrans(1).TranType
      Case 1:
        TrType = 1
      Case 2:
        TrType = 2
      Case 3:
        TrType = 3
      Case 4:
        TrType = 4
      Case 5:
        TrType = 5
      Case 6:
        TrType = 6
      Case 7:
        TrType = 7
      Case 9:
        TrType = 9
      Case 10:
        TrType = 10
      Case 11:
        TrType = 11
      Case 12:
        TrType = 12
      Case 13:
        TrType = 13
      Case 14:
        TrType = 14
      Case 21:
        TrType = 16
      Case 22:
        TrType = 8
      Case 24:
        TrType = 15
      End Select
      If TaxTrans(1).BillType = "R" Then
        TaxYear = TaxTrans(1).TaxYear
        Curyr = CuryrR
        If TaxYear < 1 Then TaxYear = Curyr
        TaxYear = TaxYear - 1979                'Reduce Based on 1980 being = 1
        RTPrinciple1#(TrType, TaxYear) = RTPrinciple1#(TrType, TaxYear) + TaxTrans(1).Revenue.Principle1
        RTInterest#(TrType, TaxYear) = RTInterest#(TrType, TaxYear) + TaxTrans(1).Revenue.Interest
        RTCollection#(TrType, TaxYear) = RTCollection#(TrType, TaxYear) + TaxTrans(1).Revenue.Collection
        RTLateList#(TrType, TaxYear) = RTLateList#(TrType, TaxYear) + TaxTrans(1).Revenue.LateList
        RTRevOpt1#(TrType, TaxYear) = RTRevOpt1#(TrType, TaxYear) + TaxTrans(1).Revenue.RevOpt1
        RTRevOpt2#(TrType, TaxYear) = RTRevOpt2#(TrType, TaxYear) + TaxTrans(1).Revenue.RevOpt2
        RTRevOpt3#(TrType, TaxYear) = RTRevOpt3#(TrType, TaxYear) + TaxTrans(1).Revenue.RevOpt3
        RTPenalty#(TrType, TaxYear) = RTPenalty#(TrType, TaxYear) + TaxTrans(1).Revenue.Penalty
        RTPenaltyPd#(TrType, TaxYear) = RTPenaltyPd#(TrType, TaxYear) + TaxTrans(1).Revenue.PenaltyPd
        RTPrinciple1Pd#(TrType, TaxYear) = RTPrinciple1Pd#(TrType, TaxYear) + TaxTrans(1).Revenue.Principle1Pd
        RTInterestPd#(TrType, TaxYear) = RTInterestPd#(TrType, TaxYear) + TaxTrans(1).Revenue.InterestPd
        RTCollectionPd#(TrType, TaxYear) = RTCollectionPd#(TrType, TaxYear) + TaxTrans(1).Revenue.CollectionPd
        RTLateListPd#(TrType, TaxYear) = RTLateListPd#(TrType, TaxYear) + TaxTrans(1).Revenue.LateListPd
        RTRevOpt1Pd#(TrType, TaxYear) = RTRevOpt1Pd#(TrType, TaxYear) + TaxTrans(1).Revenue.RevOpt1Pd
        RTRevOpt2Pd#(TrType, TaxYear) = RTRevOpt2Pd#(TrType, TaxYear) + TaxTrans(1).Revenue.RevOpt2Pd
        RTRevOpt3Pd#(TrType, TaxYear) = RTRevOpt3Pd#(TrType, TaxYear) + TaxTrans(1).Revenue.RevOpt3Pd
        RTPrePaidAmt#(TrType, TaxYear) = RTPrePaidAmt#(TrType, TaxYear) + TaxTrans(1).Revenue.PrePaidAmt
        RTPrePaidUsed#(TrType, TaxYear) = RTPrePaidUsed#(TrType, TaxYear) + TaxTrans(1).Revenue.PrePaidUsed
      End If
  Next PCnt
If InterfaceMethod <> 1 Then  ' 1 is cash so skip all charges if cash
    'Now Post for bill trans and move on to next type
  'TranType1 Billing
  For T% = 1 To 51
       'this is for the real
    If RTPrinciple1#(1, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).TaxDBAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).TaxDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = RTPrinciple1#(1, T%)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).TaxCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).TaxCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Invalid Acct"
      End If
      CreditAmt# = RTPrinciple1#(1, T%)
      Dsc$ = "Bills"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If
    If RTRevOpt1#(1, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt1DBAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt1DBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = RTRevOpt1#(1, T%)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt1CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt1CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Invaild Acct"
      End If
      CreditAmt# = RTRevOpt1#(1, T%)
      Dsc$ = "Bills"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If
    If RTRevOpt2#(1, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt2DBAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt2DBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = RTRevOpt2#(1, T%)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt2CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt2CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Invaild Acct"
      End If
      CreditAmt# = RTRevOpt2#(1, T%)
      Dsc$ = "Bills"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If
    If RTRevOpt3#(1, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt3DBAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt3DBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = RTRevOpt3#(1, T%)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt3CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt3CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Invaild Acct"
      End If
      CreditAmt# = RTRevOpt3#(1, T%)
      Dsc$ = "Bills"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If
    If RTLateList#(1, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).LtLstDBAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).LtLstDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = RTLateList#(1, T%)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).LtLstCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).LtLstCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Invaild Acct"
      End If
      CreditAmt# = RTLateList#(1, T%)
      Dsc$ = "BillsLL"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If
  Next T%
  '@@@@@@@@@@@@@@@@end of billing trantype-1
  'Interest charge  tran-4
  For T% = 1 To 51
    If RTInterest#(4, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).IntDBAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).IntDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = RTInterest#(4, T%)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).IntCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).IntCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = RTInterest#(4, T%)
      Dsc$ = "Interest"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If
  Next T%
  '@@@@@@@@@@end of tran type 5 Penalty charge
    'Penalty charge  tran-5
  For T% = 1 To 51
    If RTPenalty#(5, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).PenDBAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).PenDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = RTPenalty#(5, T%)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).PenCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).PenCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = RTPenalty#(5, T%)
      Dsc$ = "Penalty"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If
  Next T%
  '@@@@@@@@@@end of tran type 5 Penalty charge

  ''''''''''''''''''''''''''''''''''''''''''''''
  'Ad/Collection charge tran 6
  For T% = 1 To 51   'only for real
    If RTCollection#(6, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).AdvDBAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).AdvDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = RTCollection#(6, T%)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).AdvCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).AdvCRAcct
     If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = RTCollection#(6, T%)
      Dsc$ = "AdCharge"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If
  Next T%
'@@@@@@@@@@@@@@@@@@@@end of tran type 6 ad charge(collection)
''''''''''''
'Up Adj  14, 24
  For T% = 14 To 51
    'for real
    If RTPrinciple1#(14, T%) > 0 Or RTPrinciple1#(15, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).TaxDBAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).TaxDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = Round#(RTPrinciple1#(14, T%) + RTPrinciple1#(15, T%))
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).TaxCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).TaxCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(RTPrinciple1#(14, T%) + RTPrinciple1#(15, T%))
      Dsc$ = "UpAdjP"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If
    If RTRevOpt1#(14, T%) > 0 Or RTRevOpt1#(15, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt1DBAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt1DBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = Round#(RTRevOpt1#(14, T%) + RTRevOpt1#(15, T%))
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt1CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt1CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(RTRevOpt1#(14, T%) + RTRevOpt1#(15, T%))
      Dsc$ = "UpAdjOR1"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If
    If RTRevOpt2#(14, T%) > 0 Or RTRevOpt2#(15, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt2DBAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt2DBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = Round#(RTRevOpt2#(14, T%) + RTRevOpt2#(15, T%))
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt2CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt2CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(RTRevOpt2#(14, T%) + RTRevOpt2#(15, T%))
      Dsc$ = "UpAdjOR2"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If
    If RTRevOpt3#(14, T%) > 0 Or RTRevOpt3#(15, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt3DBAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt3DBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = Round#(RTRevOpt3#(14, T%) + RTRevOpt3#(15, T%))
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt3CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt3CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(RTRevOpt3#(14, T%) + RTRevOpt3#(15, T%))
      Dsc$ = "UpAdjOR3"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If
    If RTPenalty#(14, T%) > 0 Or RTPenalty#(15, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).PenDBAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).PenDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = Round#(RTPenalty#(14, T%) + RTPenalty#(15, T%))
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).PenCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).PenCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(RTPenalty#(14, T%) + RTPenalty#(15, T%))
      Dsc$ = "UpAdjPen"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If
    If RTLateList#(14, T%) > 0 Or RTLateList#(15, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).LtLstDBAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).LtLstDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = Round#(RTLateList#(14, T%) + RTLateList#(15, T%))
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).LtLstCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).LtLstCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(RTLateList#(14, T%) + RTLateList#(15, T%))
      Dsc$ = "UpAdjLL"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If
    If RTInterest#(14, T%) > 0 Or RTInterest#(15, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).IntDBAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).IntDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = Round#(RTInterest#(14, T%) + RTInterest#(15, T%))
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).IntCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).IntCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(RTInterest#(14, T%) + RTInterest#(15, T%))
      Dsc$ = "UpAdjInt"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If
    If RTCollection#(14, T%) > 0 Or RTCollection#(15, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).AdvDBAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).AdvDBAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = Round#(RTCollection#(14, T%) + RTCollection#(15, T%))
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).AdvCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).AdvCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(RTCollection#(14, T%) + RTCollection#(15, T%))
      Dsc$ = "UpAdjAC"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If
  Next T%
'end of up adj  14, 24
'tran type 3-Release
  For T% = 1 To 51
    'for real
    If RTPrinciple1Pd#(3, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).TaxCRAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).TaxCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = RTPrinciple1Pd#(3, T%)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).TaxDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).TaxDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = RTPrinciple1Pd#(3, T%)
      Dsc$ = "Release"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If

    If RTRevOpt1Pd#(3, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt1CRAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt1CRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = RTRevOpt1Pd#(3, T%)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt1DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt1DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = RTRevOpt1Pd#(3, T%)
      Dsc$ = "Release"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If
    If RTRevOpt2Pd#(3, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt2CRAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt2CRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = RTRevOpt2Pd#(3, T%)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt2DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt2DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = RTRevOpt2Pd#(3, T%)
      Dsc$ = "Release"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If
    If RTRevOpt3Pd(3, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt3CRAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt3CRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = RTRevOpt3Pd#(3, T%)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt3DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt3DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = RTRevOpt3Pd#(3, T%)
      Dsc$ = "Release"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If

    If RTPenaltyPd#(3, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).PenCRAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).PenCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = RTPenaltyPd#(3, T%)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).PenDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).PenDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = RTPenaltyPd#(3, T%)
      Dsc$ = "Release"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If

    If RTInterestPd(3, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).IntCRAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).IntCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = RTInterestPd#(3, T%)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).IntDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).IntDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = RTInterestPd#(3, T%)
      Dsc$ = "Release"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If

    If RTLateListPd#(3, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).LtLstCRAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).LtLstCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = RTLateListPd#(3, T%)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).LtLstDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).LtLstDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = RTLateListPd#(3, T%)
      Dsc$ = "Release"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If
  'Ad/Collection release
    If RTCollectionPd#(3, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).AdvCRAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).AdvCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = RTCollectionPd#(3, T%)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).AdvDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).AdvDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = RTCollectionPd#(3, T%)
      Dsc$ = "Release"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If

  Next T%
'tran type 13-Adj bill down
  For T% = 1 To 51
  'for real
    If RTPrinciple1#(13, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).TaxCRAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).TaxCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = RTPrinciple1#(13, T%)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).TaxDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).TaxDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = RTPrinciple1#(13, T%)
      Dsc$ = "AdjBillDown"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If
    If RTRevOpt1#(13, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt1CRAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt1CRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = RTRevOpt1#(13, T%)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt1DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt1DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = RTRevOpt1#(13, T%)
      Dsc$ = "AdjBillDown"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If
    If RTRevOpt2#(13, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt2CRAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt2CRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = RTRevOpt2#(13, T%)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt2DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt2DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = RTRevOpt2#(13, T%)
      Dsc$ = "AdjBillDown"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If
    If RTRevOpt3#(13, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt3CRAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt3CRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = RTRevOpt3#(13, T%)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt3DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt3DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = RTRevOpt3#(13, T%)
      Dsc$ = "AdjBillDown"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If

    If RTLateList#(13, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).LtLstCRAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).LtLstCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = RTLateList#(13, T%)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).LtLstDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).LtLstDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = RTLateList#(13, T%)
      Dsc$ = "AdjBillDown"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If
  'Interest adjdownforbill
    If RTInterest#(13, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).IntCRAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).IntCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = RTInterest#(13, T%)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).IntDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).IntDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = RTInterest#(13, T%)
      Dsc$ = "AdjBillDown"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If
    If RTPenalty#(13, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).PenCRAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).PenCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = RTPenalty#(13, T%)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).PenDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).PenDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = RTPenalty#(13, T%)
      Dsc$ = "AdjBillDown"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If
  'Ad/Collection adjdownforbill
    If RTCollection#(13, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).AdvCRAcct)
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).AdvCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = RTCollection#(13, T%)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).AdvDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).AdvDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = RTCollection#(13, T%)
      Dsc$ = "AdjBillDown"
      GoSub PostToGeneralJournal
      Close TXRGLFile
    End If
  Next T%
End If
'''end of tran type 13
''also end of charges ups/downs
'''''''''''''''''''''''''
  'TranType2 Payments-2,Payment w/prepay-21, Prepay only-22
  For T% = 1 To 51
    'for real
    If RTPrinciple1Pd#(2, T%) > 0 Or RTPrinciple1Pd#(16, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXRGLFile = FreeFile
        Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
        Get TXRGLFile, 1, TaxRGLAccts(1)
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).TaxCRAcct)
        DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).TaxCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(RTPrinciple1Pd#(2, T%) + RTPrinciple1Pd#(16, T%))
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).TaxDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).TaxDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(RTPrinciple1Pd#(2, T%) + RTPrinciple1Pd#(16, T%))
        Dsc$ = "PaymentP"
        GoSub PostToGeneralJournal
        Close TXRGLFile
      End If
      TXRGLFile = FreeFile
      Open "TAXRGLACT.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).TaxDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).TaxDBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(RTPrinciple1Pd#(2, T%) + RTPrinciple1Pd#(16, T%))
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).TaxCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).TaxCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(RTPrinciple1Pd#(2, T%) + RTPrinciple1Pd#(16, T%))
      Dsc$ = "PaymentP"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXRGLFile
    End If
    If RTRevOpt1Pd#(2, T%) > 0 Or RTRevOpt1Pd#(16, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXRGLFile = FreeFile
        Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
        Get TXRGLFile, 1, TaxRGLAccts(1)
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt1CRAcct)
        DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt1CRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(RTRevOpt1Pd#(2, T%) + RTRevOpt1Pd#(16, T%))
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt1DBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt1DBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(RTRevOpt1Pd#(2, T%) + RTRevOpt1Pd#(16, T%))
        Dsc$ = "PaymentO1"
        GoSub PostToGeneralJournal
        Close TXRGLFile
      End If
      TXRGLFile = FreeFile
      Open "TAXRGLACT.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt1DBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt1DBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(RTRevOpt1Pd#(2, T%) + RTRevOpt1Pd#(16, T%))
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt1CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt1CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(RTRevOpt1Pd#(2, T%) + RTRevOpt1Pd#(16, T%))
      Dsc$ = "PaymentO1"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXRGLFile
    End If
    If RTRevOpt2Pd#(2, T%) > 0 Or RTRevOpt2Pd#(16, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXRGLFile = FreeFile
        Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
        Get TXRGLFile, 1, TaxRGLAccts(1)
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt2CRAcct)
        DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt2CRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(RTRevOpt2Pd#(2, T%) + RTRevOpt2Pd#(16, T%))
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt2DBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt2DBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(RTRevOpt2Pd#(2, T%) + RTRevOpt2Pd#(16, T%))
        Dsc$ = "PaymentO2"
        GoSub PostToGeneralJournal
        Close TXRGLFile
      End If
      TXRGLFile = FreeFile
      Open "TAXRGLACT.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt2DBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt2DBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(RTRevOpt2Pd#(2, T%) + RTRevOpt2Pd#(16, T%))
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt2CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt2CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(RTRevOpt2Pd#(2, T%) + RTRevOpt2Pd#(16, T%))
      Dsc$ = "PaymentO2"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXRGLFile
    End If
    If RTRevOpt3Pd#(2, T%) > 0 Or RTRevOpt3Pd#(16, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXRGLFile = FreeFile
        Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
        Get TXRGLFile, 1, TaxRGLAccts(1)
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt3CRAcct)
        DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt3CRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(RTRevOpt3Pd#(2, T%) + RTRevOpt3Pd#(16, T%))
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt3DBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt3DBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(RTRevOpt3Pd#(2, T%) + RTRevOpt3Pd#(16, T%))
        Dsc$ = "PaymentO3"
        GoSub PostToGeneralJournal
        Close TXRGLFile
      End If
      TXRGLFile = FreeFile
      Open "TAXRGLACT.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt3DBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt3DBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(RTRevOpt3Pd#(2, T%) + RTRevOpt3Pd#(16, T%))
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt3CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt3CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(RTRevOpt3Pd#(2, T%) + RTRevOpt3Pd#(16, T%))
      Dsc$ = "PaymentO3"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXRGLFile
    End If
    If RTCollectionPd#(2, T%) > 0 Or RTCollectionPd#(16, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXRGLFile = FreeFile
        Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
        Get TXRGLFile, 1, TaxRGLAccts(1)
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).AdvCRAcct)
        DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).AdvCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(RTCollectionPd#(2, T%) + RTCollectionPd#(16, T%))
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).AdvDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).AdvDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(RTCollectionPd#(2, T%) + RTCollectionPd#(16, T%))
        Dsc$ = "PaymentA"
        GoSub PostToGeneralJournal
        Close TXRGLFile
      End If
      TXRGLFile = FreeFile
      Open "TAXRGLACT.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).AdvDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).AdvDBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(RTCollectionPd#(2, T%) + RTCollectionPd#(16, T%))
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).AdvCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).AdvCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(RTCollectionPd#(2, T%) + RTCollectionPd#(16, T%))
      Dsc$ = "PaymentA"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXRGLFile
    End If
    If RTLateListPd#(2, T%) > 0 Or RTLateListPd#(16, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXRGLFile = FreeFile
        Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
        Get TXRGLFile, 1, TaxRGLAccts(1)
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).LtLstCRAcct)
        DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).LtLstCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(RTLateListPd#(2, T%) + RTLateListPd#(16, T%))
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).LtLstDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).LtLstDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(RTLateListPd#(2, T%) + RTLateListPd#(16, T%))
        Dsc$ = "PaymentLL"
        GoSub PostToGeneralJournal
        Close TXRGLFile
      End If
      TXRGLFile = FreeFile
      Open "TAXRGLACT.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).LtLstDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).LtLstDBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(RTLateListPd#(2, T%) + RTLateListPd#(16, T%))
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).LtLstCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).LtLstCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(RTLateListPd#(2, T%) + RTLateListPd#(16, T%))
      Dsc$ = "PaymentLL"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXRGLFile
    End If
    If RTInterestPd#(2, T%) > 0 Or RTInterestPd#(16, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXRGLFile = FreeFile
        Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
        Get TXRGLFile, 1, TaxRGLAccts(1)
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).IntCRAcct)
        DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).IntCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(RTInterestPd#(2, T%) + RTInterestPd#(16, T%))
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).IntDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).IntDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(RTInterestPd#(2, T%) + RTInterestPd#(16, T%))
        Dsc$ = "PaymentI"
        GoSub PostToGeneralJournal
        Close TXRGLFile
      End If
      TXRGLFile = FreeFile
      Open "TAXRGLACT.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).IntDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).IntDBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(RTInterestPd#(2, T%) + RTInterestPd#(16, T%))
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).IntCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).IntCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(RTInterestPd#(2, T%) + RTInterestPd#(16, T%))
      Dsc$ = "PaymentI"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXRGLFile
    End If
    If RTPenaltyPd#(2, T%) > 0 Or RTPenaltyPd#(16, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXRGLFile = FreeFile
        Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
        Get TXRGLFile, 1, TaxRGLAccts(1)
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).PenCRAcct)
        DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).PenCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(RTPenaltyPd#(2, T%) + RTPenaltyPd#(16, T%))
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).PenDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).PenDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(RTPenaltyPd#(2, T%) + RTPenaltyPd#(16, T%))
        Dsc$ = "PaymentPen"
        GoSub PostToGeneralJournal
        Close TXRGLFile
      End If
      TXRGLFile = FreeFile
      Open "TAXRGLACT.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).PenDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).PenDBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(RTPenaltyPd#(2, T%) + RTPenaltyPd#(16, T%))
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).PenCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).PenCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(RTPenaltyPd#(2, T%) + RTPenaltyPd#(16, T%))
      Dsc$ = "PaymentPen"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXRGLFile
    End If
'Prepayment
    If RTPrePaidAmt#(8, T%) > 0 Or RTPrePaidAmt#(16, T%) > 0 Then
      TXRGLFile = FreeFile   'use the real cash acct
      Open "TAXRGLACT.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).TaxDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).TaxDBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(RTPrePaidAmt#(8, T%) + RTPrePaidAmt#(16, T%))
      ThisAcct = AcctFind(TaxSetuprec(1).OverPayGLNum)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxSetuprec(1).OverPayGLNum
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(RTPrePaidAmt#(8, T%) + RTPrePaidAmt#(16, T%))
      Dsc$ = "PaymentPre"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXRGLFile
    End If
  Next T%
'@@@@@@@@@@@@@@@@@@end of payment types, 2, 21, 22
''''''''''''''''''''''''''''''''''
  'Overpayment Apply during billing -9 and -24(amts applied during up adj)
  For T% = 1 To 51
    If RTPrePaidUsed#(9, T%) > 0 Or RTPrePaidUsed#(15, T%) > 0 Then
      ThisAcct = AcctFind(TaxSetuprec(1).OverPayGLNum)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxSetuprec(1).OverPayGLNum
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(RTPrePaidUsed#(9, T%) + RTPrePaidUsed#(15, T%))
      Dsc$ = "OverpayUsed"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = DebitAcctRecord
      GJRec1(1).AcctNum = DebitAcctNumber$
      GJRec1(1).AcctName = DebitAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).DrAmt = DebitAmt#
      GJRec1(1).EType = "D"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      GJrecnum = LOF(GJFile) \ GJReclen
      Put #GJFile, GJrecnum + 1, GJRec1(1)
    End If
    'for real
    If RTPrinciple1Pd#(9, T%) > 0 Or RTPrinciple1Pd#(15, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXRGLFile = FreeFile
        Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
        Get TXRGLFile, 1, TaxRGLAccts(1)
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).TaxCRAcct)
        DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).TaxCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(RTPrinciple1Pd#(9, T%) + RTPrinciple1Pd#(15, T%))
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).TaxDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).TaxDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(RTPrinciple1Pd#(9, T%) + RTPrinciple1Pd#(15, T%))
        Dsc$ = "OverPayApplyP"
        GoSub PostToGeneralJournal
        Close TXRGLFile
      End If
      TXRGLFile = FreeFile
      Open "TAXRGLACT.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).TaxCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).TaxCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(RTPrinciple1Pd#(9, T%) + RTPrinciple1Pd#(15, T%))
      Dsc$ = "OverPayApplyP"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = CreditAcctRecord
      GJRec1(1).AcctNum = CreditAcctNumber$
      GJRec1(1).AcctName = CreditAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = CreditAmt#
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      GJrecnum = LOF(GJFile) \ GJReclen
      Put #GJFile, GJrecnum + 1, GJRec1(1)
      Close TXRGLFile
    End If
    If RTRevOpt1Pd#(9, T%) > 0 Or RTRevOpt1Pd#(15, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXRGLFile = FreeFile
        Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
        Get TXRGLFile, 1, TaxRGLAccts(1)
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt1CRAcct)
        DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt1CRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(RTRevOpt1Pd#(9, T%) + RTRevOpt1Pd#(15, T%))
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt1DBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt1DBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(RTRevOpt1Pd#(9, T%) + RTRevOpt1Pd#(15, T%))
        Dsc$ = "OverPayApplyO1"
        GoSub PostToGeneralJournal
        Close TXRGLFile
      End If
      TXRGLFile = FreeFile
      Open "TAXRGLACT.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt1CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt1CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(RTRevOpt1Pd#(9, T%) + RTRevOpt1Pd#(15, T%))
      Dsc$ = "OverPayApplyO1"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = CreditAcctRecord
      GJRec1(1).AcctNum = CreditAcctNumber$
      GJRec1(1).AcctName = CreditAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = CreditAmt#
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      GJrecnum = LOF(GJFile) \ GJReclen
      Put #GJFile, GJrecnum + 1, GJRec1(1)
      Close TXRGLFile
    End If
    If RTRevOpt2Pd#(9, T%) > 0 Or RTRevOpt2Pd#(15, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXRGLFile = FreeFile
        Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
        Get TXRGLFile, 1, TaxRGLAccts(1)
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt2CRAcct)
        DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt2CRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(RTRevOpt2Pd#(9, T%) + RTRevOpt2Pd#(15, T%))
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt2DBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt2DBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(RTRevOpt2Pd#(9, T%) + RTRevOpt2Pd#(15, T%))
        Dsc$ = "OverPayApplyO2"
        GoSub PostToGeneralJournal
        Close TXRGLFile
      End If
      TXRGLFile = FreeFile
      Open "TAXRGLACT.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt2CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt2CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(RTRevOpt2Pd#(9, T%) + RTRevOpt2Pd#(15, T%))
      Dsc$ = "OverPayApplyO2"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = CreditAcctRecord
      GJRec1(1).AcctNum = CreditAcctNumber$
      GJRec1(1).AcctName = CreditAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = CreditAmt#
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      GJrecnum = LOF(GJFile) \ GJReclen
      Put #GJFile, GJrecnum + 1, GJRec1(1)
      Close TXRGLFile
    End If
    If RTRevOpt3Pd#(9, T%) > 0 Or RTRevOpt3Pd#(15, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXRGLFile = FreeFile
        Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
        Get TXRGLFile, 1, TaxRGLAccts(1)
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt3CRAcct)
        DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt3CRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(RTRevOpt3Pd#(9, T%) + RTRevOpt3Pd#(15, T%))
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt3DBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt3DBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(RTRevOpt3Pd#(9, T%) + RTRevOpt3Pd#(15, T%))
        Dsc$ = "OverPayApplyO3"
        GoSub PostToGeneralJournal
        Close TXRGLFile
      End If
      TXRGLFile = FreeFile
      Open "TAXRGLACT.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt3CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt3CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(RTRevOpt3Pd#(9, T%) + RTRevOpt3Pd#(15, T%))
      Dsc$ = "OverPayApplyO3"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = CreditAcctRecord
      GJRec1(1).AcctNum = CreditAcctNumber$
      GJRec1(1).AcctName = CreditAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = CreditAmt#
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      GJrecnum = LOF(GJFile) \ GJReclen
      Put #GJFile, GJrecnum + 1, GJRec1(1)
      Close TXRGLFile
    End If
    If RTInterestPd#(9, T%) > 0 Or RTInterestPd#(15, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXRGLFile = FreeFile
        Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
        Get TXRGLFile, 1, TaxRGLAccts(1)
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).IntCRAcct)
        DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).IntCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(RTInterestPd#(9, T%) + RTInterestPd#(15, T%))
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).IntDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).IntDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(RTInterestPd#(9, T%) + RTInterestPd#(15, T%))
        Dsc$ = "OverPayApplyI"
        GoSub PostToGeneralJournal
        Close TXRGLFile
      End If
      TXRGLFile = FreeFile
      Open "TAXRGLACT.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).IntCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).IntCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(RTInterestPd#(9, T%) + RTInterestPd#(15, T%))
      Dsc$ = "OverPayApplyI"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = CreditAcctRecord
      GJRec1(1).AcctNum = CreditAcctNumber$
      GJRec1(1).AcctName = CreditAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = CreditAmt#
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      GJrecnum = LOF(GJFile) \ GJReclen
      Put #GJFile, GJrecnum + 1, GJRec1(1)
      Close TXRGLFile
    End If
    If RTPenaltyPd#(9, T%) > 0 Or RTPenaltyPd#(15, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXRGLFile = FreeFile
        Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
        Get TXRGLFile, 1, TaxRGLAccts(1)
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).PenCRAcct)
        DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).PenCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(RTPenaltyPd#(9, T%) + RTPenaltyPd#(15, T%))
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).PenDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).PenDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(RTPenaltyPd#(9, T%) + RTPenaltyPd#(15, T%))
        Dsc$ = "OverPayApplyPen"
        GoSub PostToGeneralJournal
        Close TXRGLFile
      End If
      TXRGLFile = FreeFile
      Open "TAXRGLACT.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).PenCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).PenCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(RTPenaltyPd#(9, T%) + RTPenaltyPd#(15, T%))
      Dsc$ = "OverPayApplyPen"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = CreditAcctRecord
      GJRec1(1).AcctNum = CreditAcctNumber$
      GJRec1(1).AcctName = CreditAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = CreditAmt#
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      GJrecnum = LOF(GJFile) \ GJReclen
      Put #GJFile, GJrecnum + 1, GJRec1(1)
      Close TXRGLFile
    End If
    If RTCollectionPd#(9, T%) > 0 Or RTCollectionPd#(15, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXRGLFile = FreeFile
        Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
        Get TXRGLFile, 1, TaxRGLAccts(1)
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).AdvCRAcct)
        DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).AdvCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(RTCollectionPd#(9, T%) + RTCollectionPd#(15, T%))
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).AdvDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).AdvDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(RTCollectionPd#(9, T%) + RTCollectionPd#(15, T%))
        Dsc$ = "OverPayApplyA"
        GoSub PostToGeneralJournal
        Close TXRGLFile
      End If
      TXRGLFile = FreeFile
      Open "TAXRGLACT.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).AdvCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).AdvCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(RTCollectionPd#(9, T%) + RTCollectionPd#(15, T%))
      Dsc$ = "OverPayApplyA"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = CreditAcctRecord
      GJRec1(1).AcctNum = CreditAcctNumber$
      GJRec1(1).AcctName = CreditAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = CreditAmt#
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      GJrecnum = LOF(GJFile) \ GJReclen
      Put #GJFile, GJrecnum + 1, GJRec1(1)
      Close TXRGLFile
    End If
    If RTLateListPd#(9, T%) > 0 Or RTLateListPd#(15, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXRGLFile = FreeFile
        Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
        Get TXRGLFile, 1, TaxRGLAccts(1)
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).LtLstCRAcct)
        DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).LtLstCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(RTLateListPd#(9, T%) + RTLateListPd#(15, T%))
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).LtLstDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).LtLstDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(RTLateListPd#(9, T%) + RTLateListPd#(15, T%))
        Dsc$ = "OverPayApplyLL"
        GoSub PostToGeneralJournal
        Close TXRGLFile
      End If
      TXRGLFile = FreeFile
      Open "TAXRGLACT.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).LtLstCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).LtLstCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(RTLateListPd#(9, T%) + RTLateListPd#(15, T%))
      Dsc$ = "OverPayApplyLL"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = CreditAcctRecord
      GJRec1(1).AcctNum = CreditAcctNumber$
      GJRec1(1).AcctName = CreditAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = CreditAmt#
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      GJrecnum = LOF(GJFile) \ GJReclen
      Put #GJFile, GJrecnum + 1, GJRec1(1)
      Close TXRGLFile
    End If
  Next T%
'@@@@@@@@@@@@@@ end of tran type 9 and apply prepay amt side of 24 up adj
''''''''''''''''''''''''''''
'tran 7 adjust pay down  and  10 paydown w/prep for interface 3(Modified Accrual) do bill charge again
  For T% = 1 To 51
    'for real
    If RTPrinciple1Pd#(7, T%) > 0 Or RTPrinciple1Pd#(10, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXRGLFile = FreeFile
        Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
        Get TXRGLFile, 1, TaxRGLAccts(1)
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).TaxDBAcct)
        DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).TaxDBAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(RTPrinciple1Pd#(7, T%) + RTPrinciple1Pd#(10, T%))
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).TaxCRAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).TaxCRAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(RTPrinciple1Pd#(7, T%) + RTPrinciple1Pd#(10, T%))
        Dsc$ = "AdjPayP"
        GoSub PostToGeneralJournal
        Close TXRGLFile
      End If
      TXRGLFile = FreeFile
      Open "TAXRGLACT.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).TaxCRAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).TaxCRAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round#(RTPrinciple1Pd#(7, T%) + RTPrinciple1Pd#(10, T%))
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).TaxDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).TaxDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(RTPrinciple1Pd#(7, T%) + RTPrinciple1Pd#(10, T%))
      Dsc$ = "AdjPayP"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXRGLFile
    End If
    If RTRevOpt1Pd#(7, T%) > 0 Or RTRevOpt1Pd#(10, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXRGLFile = FreeFile
        Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
        Get TXRGLFile, 1, TaxRGLAccts(1)
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt1DBAcct)
        DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt1DBAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(RTRevOpt1Pd#(7, T%) + RTRevOpt1Pd#(10, T%))
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt1CRAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt1CRAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(RTRevOpt1Pd#(7, T%) + RTRevOpt1Pd#(10, T%))
        Dsc$ = "AdjPayO1"
        GoSub PostToGeneralJournal
        Close TXRGLFile
      End If
      TXRGLFile = FreeFile
      Open "TAXRGLACT.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt1CRAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt1CRAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round#(RTRevOpt1Pd#(7, T%) + RTRevOpt1Pd#(10, T%))
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt1DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt1DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(RTRevOpt1Pd#(7, T%) + RTRevOpt1Pd#(10, T%))
      Dsc$ = "AdjPayO1"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXRGLFile
    End If
    If RTRevOpt2Pd#(7, T%) > 0 Or RTRevOpt2Pd#(10, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXRGLFile = FreeFile
        Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
        Get TXRGLFile, 1, TaxRGLAccts(1)
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt2DBAcct)
        DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt2DBAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(RTRevOpt2Pd#(7, T%) + RTRevOpt2Pd#(10, T%))
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt2CRAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt2CRAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(RTRevOpt2Pd#(7, T%) + RTRevOpt2Pd#(10, T%))
        Dsc$ = "AdjPayO2"
        GoSub PostToGeneralJournal
        Close TXRGLFile
      End If
      TXRGLFile = FreeFile
      Open "TAXRGLACT.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt2CRAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt2CRAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round#(RTRevOpt2Pd#(7, T%) + RTRevOpt2Pd#(10, T%))
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt2DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt2DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(RTRevOpt2Pd#(7, T%) + RTRevOpt2Pd#(10, T%))
      Dsc$ = "AdjPayO2"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXRGLFile
    End If
    If RTRevOpt3Pd#(7, T%) > 0 Or RTRevOpt3Pd#(10, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXRGLFile = FreeFile
        Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
        Get TXRGLFile, 1, TaxRGLAccts(1)
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt3DBAcct)
        DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt3DBAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(RTRevOpt3Pd#(7, T%) + RTRevOpt3Pd#(10, T%))
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt3CRAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt3CRAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(RTRevOpt3Pd#(7, T%) + RTRevOpt3Pd#(10, T%))
        Dsc$ = "AdjPayO3"
        GoSub PostToGeneralJournal
        Close TXRGLFile
      End If
      TXRGLFile = FreeFile
      Open "TAXRGLACT.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt3CRAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt3CRAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round#(RTRevOpt3Pd#(7, T%) + RTRevOpt3Pd#(10, T%))
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).Opt3DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).Opt3DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(RTRevOpt3Pd#(7, T%) + RTRevOpt3Pd#(10, T%))
      Dsc$ = "AdjPayO3"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXRGLFile
    End If
    If RTInterestPd#(7, T%) > 0 Or RTInterestPd#(10, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXRGLFile = FreeFile
        Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
        Get TXRGLFile, 1, TaxRGLAccts(1)
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).IntDBAcct)
        DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).IntDBAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(RTInterestPd#(7, T%) + RTInterestPd#(10, T%))
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).IntCRAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).IntCRAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(RTInterestPd#(7, T%) + RTInterestPd#(10, T%))
        Dsc$ = "AdjPayI"
        GoSub PostToGeneralJournal
        Close TXRGLFile
      End If
      TXRGLFile = FreeFile
      Open "TAXRGLACT.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).IntCRAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).IntCRAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round#(RTInterestPd#(7, T%) + RTInterestPd#(10, T%))
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).IntDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).IntDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(RTInterestPd#(7, T%) + RTInterestPd#(10, T%))
      Dsc$ = "AdjPayI"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXRGLFile
    End If
    If RTPenaltyPd#(7, T%) > 0 Or RTPenaltyPd#(10, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXRGLFile = FreeFile
        Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
        Get TXRGLFile, 1, TaxRGLAccts(1)
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).PenDBAcct)
        DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).PenDBAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(RTPenaltyPd#(7, T%) + RTPenaltyPd#(10, T%))
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).PenCRAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).PenCRAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(RTPenaltyPd#(7, T%) + RTPenaltyPd#(10, T%))
        Dsc$ = "AdjPayPen"
        GoSub PostToGeneralJournal
        Close TXRGLFile
      End If
      TXRGLFile = FreeFile
      Open "TAXRGLACT.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).PenCRAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).PenCRAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round#(RTPenaltyPd#(7, T%) + RTPenaltyPd#(10, T%))
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).PenDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).PenDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(RTPenaltyPd#(7, T%) + RTPenaltyPd#(10, T%))
      Dsc$ = "AdjPayPen"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXRGLFile
    End If
    If RTCollectionPd#(7, T%) > 0 Or RTCollectionPd#(10, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXRGLFile = FreeFile
        Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
        Get TXRGLFile, 1, TaxRGLAccts(1)
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).AdvDBAcct)
        DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).AdvDBAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(RTCollectionPd#(7, T%) + RTCollectionPd#(10, T%))
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).AdvCRAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).AdvCRAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(RTCollectionPd#(7, T%) + RTCollectionPd#(10, T%))
        Dsc$ = "AdjPayA"
        GoSub PostToGeneralJournal
        Close TXRGLFile
      End If
      TXRGLFile = FreeFile
      Open "TAXRGLACT.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).AdvCRAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).AdvCRAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round#(RTCollectionPd#(7, T%) + RTCollectionPd#(10, T%))
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).AdvCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).AdvCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(RTCollectionPd#(7, T%) + RTCollectionPd#(10, T%))
      Dsc$ = "AdjPayA"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXRGLFile
    End If
    If RTLateListPd#(7, T%) > 0 Or RTLateListPd#(10, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXRGLFile = FreeFile
        Open "TAXRGLBAC.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
        Get TXRGLFile, 1, TaxRGLAccts(1)
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).LtLstDBAcct)
        DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).LtLstDBAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(RTLateListPd#(7, T%) + RTLateListPd#(10, T%))
        ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).LtLstCRAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).LtLstCRAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(RTLateListPd#(7, T%) + RTLateListPd#(10, T%))
        Dsc$ = "AdjPayLL"
        GoSub PostToGeneralJournal
        Close TXRGLFile
      End If
      TXRGLFile = FreeFile
      Open "TAXRGLACT.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).LtLstCRAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).LtLstCRAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round#(RTLateListPd#(7, T%) + RTLateListPd#(10, T%))
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).LtLstDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).LtLstDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(RTLateListPd#(7, T%) + RTLateListPd#(10, T%))
      Dsc$ = "AdjPayLL"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXRGLFile
    End If
   
  '^*&^*&^*&^* 10 and 11 prepayment adj down
    If RTPrePaidUsed#(10, T%) > 0 Or RTPrePaidUsed#(11, T%) > 0 Then
      TXRGLFile = FreeFile
      Open "TAXRGLACt.DAT" For Random Shared As TXRGLFile Len = TaxRAcctRecLen
      Get TXRGLFile, 1, TaxRGLAccts(1)
      ThisAcct = AcctFind(TaxSetuprec(1).OverPayGLNum)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxSetuprec(1).OverPayGLNum
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(RTPrePaidUsed#(10, T%) + RTPrePaidUsed#(11, T%))
      ThisAcct = AcctFind(TaxRGLAccts(1).TaxAcct(T%).TaxDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxRGLAccts(1).TaxAcct(T%).TaxDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(RTPrePaidUsed#(10, T%) + RTPrePaidUsed#(11, T%))
      Dsc$ = "AdjPre"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXRGLFile
    End If
 Next T%
TXBunchReturn:
  Return

PostToGeneralJournal:
  'NOTE: Journal Rec 1 is the credit, Rec 2 is the debit
  ReDim GJRec1(1 To 2) As TrEditRecType
  GJRec1(1).AcctRec = CreditAcctRecord
  GJRec1(1).AcctNum = CreditAcctNumber$
  If Len(QPTrim$(CreditAcctNumber$)) > 0 Then
    GJRec1(1).AcctName = CreditAcctName$
  Else
    GJRec1(1).AcctName = "Blank Acct"
  End If
  GJRec1(1).TRDATE = WorkDate
  GJRec1(1).Ref = Ref$
  GJRec1(1).CrAmt = CreditAmt#
  GJRec1(1).EType = "C"
  GJRec1(1).Desc = "FRMTX " + Dsc$
  GJRec1(1).LDesc = "TX Interface"
  GJRec1(1).Src = "TX"
  GJrecnum = LOF(GJFile) \ GJReclen
  Put #GJFile, GJrecnum + 1, GJRec1(1)

  GJRec1(2).AcctRec = DebitAcctRecord
  GJRec1(2).AcctNum = DebitAcctNumber$
  If Len(QPTrim$(DebitAcctNumber$)) > 0 Then
    GJRec1(2).AcctName = DebitAcctName$
  Else
    GJRec1(2).AcctName = "Blank Acct"
  End If
  GJRec1(2).TRDATE = WorkDate
  GJRec1(2).Ref = Ref$
  GJRec1(2).DrAmt = DebitAmt#
  GJRec1(2).EType = "D"
  GJRec1(2).Desc = "FRMTX " + Dsc$
  GJRec1(2).LDesc = "TX Interface"
  GJRec1(2).Src = "TX"
  GJrecnum = LOF(GJFile) \ GJReclen
  Put #GJFile, GJrecnum + 1, GJRec1(2)
Return

Clearouttots:
For ppcnt = 1 To 16
  For RevCnt = 1 To 51
    RTPrinciple1#(ppcnt, RevCnt) = 0
    RTInterest#(ppcnt, RevCnt) = 0
    RTPenalty#(ppcnt, RevCnt) = 0
    RTCollection#(ppcnt, RevCnt) = 0
    RTPrinciple1Pd#(ppcnt, RevCnt) = 0
    RTInterestPd#(ppcnt, RevCnt) = 0
    RTPenaltyPd#(ppcnt, RevCnt) = 0
    RTCollectionPd#(ppcnt, RevCnt) = 0
    RTRevOpt1#(ppcnt, RevCnt) = 0
    RTRevOpt1Pd#(ppcnt, RevCnt) = 0
    RTRevOpt2#(ppcnt, RevCnt) = 0
    RTRevOpt2Pd#(ppcnt, RevCnt) = 0
    RTRevOpt3#(ppcnt, RevCnt) = 0
    RTRevOpt3Pd#(ppcnt, RevCnt) = 0
    RTLateList#(ppcnt, RevCnt) = 0
    RTLateListPd#(ppcnt, RevCnt) = 0
    RTPrePaidAmt#(ppcnt, RevCnt) = 0
    RTPrePaidUsed#(ppcnt, RevCnt) = 0

   Next
  Next
Return

TaxEnd:
  Exit Sub
  Return
End Sub

'Decals Interface
Private Sub ExtractDC(ThruDate%)
  Dim Today As String, Ref As String, Dash80 As String, P2S As String
  Dim GJReclen As Integer, RptFile As Integer, NumOFDCCatRecs As Integer
  Dim DCCatFile As Integer, NumOfTRecs As Long, TCnt As Long
  Dim FoundCnt As Integer, ccnt As Single, GJFile As Integer
  Dim NumEdTrans As Integer, MCFile As Integer, cnt As Long
  Dim FirstTran As Integer, ThisDate As Integer, WorkDate As Integer
  Dim DayCount As Integer, NGCnt As Integer
  Dim FindCount As Integer, FundCnt As Integer, Process As Integer
  Dim Acct As String, AcctName As String, AcctR As Integer
  Dim FoundFund As Integer, FCnt As Integer, Cash As Integer
  Dim DCTransRecLen As Integer, DCTransFile As Integer, AcctRec As Integer
  Dim CatCode As String, CatCodeRecord As Integer, DCCatcodereclen As Integer
  Dim SetupFile As Integer, BadAcct As Integer
  Dim PAmt As Double, VAmt As Double
  Today$ = Date$
  Ref$ = "DC" + Left$(Today$, 2) + Mid$(Today$, 4, 2) + Right$(Today$, 2)

  ReDim TranInfo(1) As TranRecInfoType
  Dim PRec#(500), PRevAmt#(500), VRec#(500), VRevAmt#(500)

  Dash80$ = String$(80, "-")
  P2S$ = Space$(4)

  GJReclen = Len(GJRec(1))
  RptFile = FreeFile
  Open "GLCMTRX.RPT" For Output As RptFile

'  ClearBox
'  QPrintRC "Searching Cash Transactions.", 11, 26, 126
'  QPrintRC "New Transactions:", 13, 29, Cnf.HiLite

  ReDim DCTransRec(1) As DCTransRecType
  DCTransRecLen = Len(DCTransRec(1))
  DCTransFile = FreeFile
  Open "DCTrans.DAT" For Random Access Read Write Shared As DCTransFile Len = DCTransRecLen
  NumOfTRecs& = LOF(DCTransFile) \ DCTransRecLen
  Lock #DCTransFile
  FrmShowPctComp.ShowPctComp 15, 100

  For TCnt& = NumOfTRecs& To 1 Step -1
    Get #DCTransFile, TCnt&, DCTransRec(1)

    If (Len(QPTrim$(DCTransRec(1).GLInterfaced)) = 0 Or QPTrim$(DCTransRec(1).GLInterfaced) = "N") And DCTransRec(1).TransType <> 1 Then
      If DCTransRec(1).TransDate <= ThruDate% Then
        'Store trans rec numbers and dates in array
        FoundCnt = FoundCnt + 1
        ReDim Preserve TranInfo(FoundCnt) As TranRecInfoType
        TranInfo(FoundCnt).TranDate = DCTransRec(1).TransDate
        TranInfo(FoundCnt).TranRecNo = TCnt&
      End If
    Else
      NGCnt = NGCnt + 1
    End If
    RSet P2S$ = Str$(FoundCnt)
    'QPrintRC P2S$, 13, 47, Cnf.HiLite
    'SmallPause
    'Allow 1500 Bad Entries Before Exiting
    If NGCnt >= 1500 Then Exit For
  Next
  FrmShowPctComp.ShowPctComp 30, 100
  If FoundCnt = 0 Then
    Unload FrmShowPctComp
    Close
    'ClearBox
    Print Chr$(7);
    Call MainLog("No DC to Grab for " + fpDate)
    MsgBox "No Transactions Found To InterFace", vbOKOnly, "No Trans"
    'SLEEP 4
    GoTo DCSendExit
  End If

  'Array (1), NumElem, Dir, StructSize, MemOff, MemSize
  SortTRec TranInfo(), FoundCnt     'sort'em by date. oldest first
  'Open GL InterFace File
  GJFile = FreeFile
  Open "GLTRXED.DAT" For Random Access Read Write Shared As GJFile Len = GJReclen
  NumEdTrans = LOF(GJFile) \ GJReclen
  Lock #GJFile

  'OPEN Decal Catagory Here
  FrmShowPctComp.ShowPctComp 40, 100
  ReDim DCCatCodeRec(1) As DCCatCodeRecType
  DCCatcodereclen = Len(DCCatCodeRec(1))
  DCCatFile = FreeFile
  Open "DCCODE.DAT" For Random Access Read Write Shared As DCCatFile Len = DCCatcodereclen
  NumOFDCCatRecs = LOF(DCCatFile) \ DCCatcodereclen


  FirstTran = 1
  ThisDate = TranInfo(1).TranDate
  WorkDate = ThisDate

  For cnt = 1 To FoundCnt
    If ThisDate <> TranInfo(cnt).TranDate Then
      ThisDate = TranInfo(cnt).TranDate
      GoSub DCProcessThisBunch
      DayCount = 0
      WorkDate = ThisDate
    End If
    Get #DCTransFile, TranInfo(cnt).TranRecNo, DCTransRec(1)
    'Find Catagory Record Number So We Can Pull G/L Revenue Account
    CatCode$ = QPTrim$(DCTransRec(1).DecalCat)
    CatCodeRecord = 0
    For ccnt! = 1 To NumOFDCCatRecs
      Get DCCatFile, ccnt!, DCCatCodeRec(1)
      If QPTrim$(DCCatCodeRec(1).CatCode) = CatCode$ Then
        CatCodeRecord = ccnt!
        Exit For
      End If
    Next ccnt!
    If CatCodeRecord > 0 Then
      If DayCount = 0 Then
          If DCTransRec(1).TransType = 2 Then
            PAmt# = DCTransRec(1).TransAmount
            PAmt# = Round#(PAmt#)
            VAmt# = 0
          ElseIf DCTransRec(1).TransType = 4 Then
            VAmt# = DCTransRec(1).TransAmount
            VAmt# = Round#(VAmt#)
            PAmt# = 0
          End If
          If PAmt# <> 0 Or VAmt# <> 0 Then
            'If There Is an Amount get catagory code record
            If CatCodeRecord >= 1 Then
              DayCount = DayCount + 1
              If PAmt# <> 0 Then
                PRec#(DayCount) = CatCodeRecord
                PRevAmt#(DayCount) = PAmt#
              ElseIf VAmt# <> 0 Then
                VRec#(DayCount) = CatCodeRecord
                VRevAmt#(DayCount) = VAmt#
              End If
            End If
          End If
        Else
          If DCTransRec(1).TransType = 2 Then
            PAmt# = DCTransRec(1).TransAmount
            PAmt# = Round#(PAmt#)
          ElseIf DCTransRec(1).TransType = 4 Then
            VAmt# = DCTransRec(1).TransAmount
            VAmt# = Round#(VAmt#)
          End If
          Do While PAmt# <> 0
            For FindCount = 1 To DayCount
              If PRec#(FindCount) = CatCodeRecord Then
                PRevAmt#(FindCount) = PRevAmt#(FindCount) + PAmt#
                PAmt# = 0
                Exit Do
              End If
            Next FindCount
            DayCount = DayCount + 1
            PRec#(DayCount) = CatCodeRecord
            PRevAmt#(DayCount) = PAmt#
            PAmt# = 0
          Loop
          Do While VAmt# <> 0
            For FindCount = 1 To DayCount
              If VRec#(FindCount) = CatCodeRecord Then
                VRevAmt#(FindCount) = VRevAmt#(FindCount) + VAmt#
                VAmt# = 0
                Exit Do
              End If
            Next FindCount
            DayCount = DayCount + 1
            VRec#(DayCount) = CatCodeRecord
            VRevAmt#(DayCount) = VAmt#
            VAmt# = 0
          Loop
        End If
    End If
  Next cnt
  FrmShowPctComp.ShowPctComp 85, 100
  GoSub DCProcessThisBunch
  If BadAcct > 0 Then
    Close
    Unload FrmShowPctComp
    Call MainLog("Error DC Grab not Created.")
    MsgBox "Errors Found, Interface File Not Created. Please Review Report.", vbOKOnly, "Errors"
    frmReportOpt.Show 1
    If rptopt = 1 Then
      frmGetDistMenu.PrnEditList 2
    ElseIf rptopt = 2 Then
      frmGetDistMenu.PrnEditList2 2
    End If

    KillFileD "GLTRXED.DAT"
    Exit Sub
  End If
  'Mark Transactions as interfaced
  For cnt = 1 To FoundCnt
    Get #DCTransFile, TranInfo(cnt).TranRecNo, DCTransRec(1)
    DCTransRec(1).GLInterfaced = "Y"
    Put #DCTransFile, TranInfo(cnt).TranRecNo, DCTransRec(1)
  Next
  FrmShowPctComp.ShowPctComp 100, 100
  Close
  Call MainLog("DC Grab Complete " + Str$(FoundCnt) + " for " + fpDate)
  MsgBox "Transaction Grab Complete.", vbOKOnly, "Complete"


DCSendExit:
  Exit Sub


DCProcessThisBunch:
  ' Must Combine By Date and Then Do Cash Debit Entry For Total by Fund
  'NOTE: Journal Rec 1 is the credit, Rec 2 is the debit
  ReDim GJRecd(1 To 2) As TrEditRecType

  If DayCount <= 0 Then Return

  FundCnt = 0   ' Set Funds Used to Zero

  For Process = 1 To DayCount
   If PRevAmt(Process) <> 0 Then
    Get #DCCatFile, PRec#(Process), DCCatCodeRec(1)
    Acct$ = DCCatCodeRec(1).REVGLNUM
    Acct$ = QPStrip$(Acct$)
    Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
    AcctRec = AcctFind(Acct$)
    If AcctRec > 0 Then
      AcctName$ = GetAcctTitle(AcctRec)
    Else
      BadAcct = BadAcct + 1
      AcctName$ = "Undefined"
    End If
    GJRecd(1).AcctRec = 0
    GJRecd(1).AcctNum = QPTrim$(Acct$)
    GJRecd(1).AcctName = AcctName$
    GJRecd(1).TRDATE = WorkDate
    GJRecd(1).Ref = Ref$
    GJRecd(1).CrAmt = PRevAmt#(Process)
    GJRecd(1).DrAmt = 0
    GJRecd(1).EType = "C"
    GJRecd(1).Desc = "FRM DC Pay"
    GJRecd(1).LDesc = "DC Interface"
    GJRecd(1).Src = "DC"
    Put #GJFile, , GJRecd(1)
    
    
    Acct$ = DCCatCodeRec(1).CashAcct
    Acct$ = QPStrip$(Acct$)
    Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
    AcctRec = AcctFind(Acct$)
    If AcctRec > 0 Then
      AcctName$ = GetAcctTitle(AcctRec)
    Else
      BadAcct = BadAcct + 1
      AcctName$ = "Undefined"
    End If
    GJRecd(2).AcctRec = 0
    GJRecd(2).AcctNum = Acct$
    GJRecd(2).AcctName = AcctName$
    GJRecd(2).TRDATE = WorkDate
    GJRecd(2).Ref = Ref$
    GJRecd(2).DrAmt = PRevAmt#(Process)
    GJRecd(2).CrAmt = 0
    GJRecd(2).EType = "D"
    GJRecd(2).Desc = "FRM DC Pay"
    GJRecd(2).LDesc = "DC Interface"
    GJRecd(2).Src = "DC"
    Put #GJFile, , GJRecd(2)
   ElseIf VRevAmt#(Process) <> 0 Then
    Get #DCCatFile, VRec#(Process), DCCatCodeRec(1)
    Acct$ = DCCatCodeRec(1).CashAcct
    Acct$ = QPStrip$(Acct$)
    Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
    AcctRec = AcctFind(Acct$)
    If AcctRec > 0 Then
      AcctName$ = GetAcctTitle(AcctRec)
    Else
      BadAcct = BadAcct + 1
      AcctName$ = "Undefined"
    End If

    GJRecd(1).AcctRec = 0
    GJRecd(1).AcctNum = QPTrim$(Acct$)
    GJRecd(1).AcctName = AcctName$
    GJRecd(1).TRDATE = WorkDate
    GJRecd(1).Ref = Ref$
    GJRecd(1).CrAmt = VRevAmt#(Process)
    GJRecd(1).DrAmt = 0
    GJRecd(1).EType = "C"
    GJRecd(1).Desc = "FRM DC VPay"
    GJRecd(1).LDesc = "DC Interface"
    GJRecd(1).Src = "DC"
    Put #GJFile, , GJRecd(1)
    
    Acct$ = DCCatCodeRec(1).REVGLNUM
    Acct$ = QPStrip$(Acct$)
    Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
    AcctRec = AcctFind(Acct$)
    If AcctRec > 0 Then
      AcctName$ = GetAcctTitle(AcctRec)
    Else
      BadAcct = BadAcct + 1
      AcctName$ = "Undefined"
    End If
    GJRecd(2).AcctRec = 0
    GJRecd(2).AcctNum = Acct$
    GJRecd(2).AcctName = AcctName$
    GJRecd(2).TRDATE = WorkDate
    GJRecd(2).Ref = Ref$
    GJRecd(2).DrAmt = VRevAmt#(Process)
    GJRecd(2).CrAmt = 0
    GJRecd(2).EType = "D"
    GJRecd(2).Desc = "FRM DC VPay"
    GJRecd(2).LDesc = "DC Interface"
    GJRecd(2).Src = "DC"
    Put #GJFile, , GJRecd(2)
   End If
  Next Process

  'Now Make Matching Debit Entries to Cash Account

'  For Cash = 1 To FundCnt
'    Acct$ = Fund$(Cash) + CashAcct$
'    Acct$ = RTrim$(Acct$)
'    Acct$ = DCCatCodeRec(1).CashAcct
'    Acct$ = QPStrip$(Acct$)
'    Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
'    AcctRec = AcctFind(Acct$)
'    If AcctRec > 0 Then
'      AcctName$ = GetAcctTitle(AcctRec)
'    Else
'      BadAcct = BadAcct + 1
'      AcctName$ = "Undefined"
'    End If
'    GJRec(1).AcctRec = 0
'    GJRec(1).AcctNum = Acct$
'    GJRec(1).AcctName = AcctName$
'    GJRec(1).TRDATE = WorkDate
'    GJRec(1).Ref = Ref$
'    GJRec(1).DrAmt = FundAmt#(Cash)
'    GJRec(1).CrAmt = 0
'    GJRec(1).EType = "D"
'    GJRec(1).Desc = "FROM DECAL I/F"
'    GJRec(1).LDesc = "DC Interface"
'    GJRec(1).Src = "DC"
'    Put #GJFile, , GJRec(1)
'  Next Cash
DCBunchReturn:
  Return

End Sub


Private Function QSortTRec(TranInfo() As TranRecInfoType, NumTrans)
  Dim TmpSort As TranRecInfoType, lngCurLow As Long, lngCurHigh As Long
  Dim OutOfOrder As Boolean, cntT As Long
  lngCurLow = LBound(TranInfo)
  lngCurHigh = UBound(TranInfo)
'      Do
'        OutOfOrder = False          'assume it's sorted
'        For cntT = 1 To NumTrans - 1
'          If TranInfo(cntT).TranDate > TranInfo(cntT + 1).TranDate Then
'            LSet TmpSort = TranInfo(cntT)
'            LSet TranInfo(cntT) = TranInfo(cntT + 1)
'            LSet TranInfo(cntT + 1) = TmpSort
'            OutOfOrder = True       'we're not done yet
'          End If
'        Next
'      Loop While OutOfOrder
  QQSortTRec TranInfo(), lngCurLow, lngCurHigh
  
End Function
Private Sub QQSortTRec(TranInfo() As TranRecInfoType, lLBound, lUBound)
Dim lngCurLow As Long, lngCurHigh As Long, lngCurMid As Long
Dim Temp As TranRecInfoType
Dim Temp2 As TranRecInfoType
lngCurLow = lLBound
lngCurHigh = lUBound
If lUBound <= lLBound Then Exit Sub
  lngCurMid = (lUBound + lLBound) \ 2
  Temp = TranInfo(lngCurMid)
  Do While (lngCurLow <= lngCurHigh)
    Do While TranInfo(lngCurLow).TranDate < Temp.TranDate
      lngCurLow = lngCurLow + 1
      If lngCurHigh = lUBound Then Exit Do
    Loop
    Do While Temp.TranDate < TranInfo(lngCurHigh).TranDate
      lngCurHigh = lngCurHigh - 1
      If lngCurHigh = lLBound Then Exit Do
    Loop
    If (lngCurLow <= lngCurHigh) Then
      Temp2 = TranInfo(lngCurLow)
      TranInfo(lngCurLow) = TranInfo(lngCurHigh)
      TranInfo(lngCurHigh) = Temp2
      lngCurLow = lngCurLow + 1
      lngCurHigh = lngCurHigh - 1
    End If
  Loop
  If lLBound < lngCurHigh Then
    QQSortTRec TranInfo(), lLBound, lngCurHigh
  End If
  If lngCurLow < lUBound Then
    QQSortTRec TranInfo(), lngCurLow, lUBound
  End If
End Sub

'''Utility Billing Transaction
''Private Sub ExtractUBold before temp detail file created 2/1/2006(ThruDate%)
''  Dim Today As String, Ref As String, Dash80 As String, P2S As String
''  Dim GJReclen As Integer, RptFile As Integer, UBTransRecLen As Integer
''  Dim UBTran As Integer, NumOfTRecs As Long, TCnt As Long, PageNo As Integer
''  Dim FoundCnt As Long, NGCnt As Integer, GJFile As Integer
''  Dim NumEdTrans As Integer, MCFile As Integer, cnt As Long
''  Dim FirstTran As Long, ThisDate As Integer, WorkDate As Integer
''  Dim DayCount As Integer, MCnt As Integer, MiscRevAmt As Double
''  Dim FindCount As Integer, FundCnt As Integer, Process As Integer
''  Dim Acct As String, AcctName As String, ThisAcct As Integer
''  Dim FoundFund As Integer, FCnt As Integer, Cash As Integer
''  Dim UBSetUpFileNum As Integer, UBSetUpLen As Integer
''  Dim AcctMeth As String, InterfaceMethod As Integer, RevCnt As Integer
''  Dim TempRev As String, NumOfRevs As Integer, BadAcct As Integer
''  Dim LastTran As Long, NumPrinted As Integer, BadCAcct As String
''  Dim ActT As Integer, ActPg As String, BadDAcct As String, PCnt As Integer
''
''  Today$ = Date$
''  Ref$ = "UB" + Left$(Today$, 2) + Mid$(Today$, 4, 2) + Right$(Today$, 2)
''  ReDim TranInfo(1) As TranRecInfoType
''  Dash80$ = String$(80, "-")
''  P2S$ = Space$(4)
''  ReDim GJRec1(1 To 2) As TrEditRecType
''  GJReclen = Len(GJRec1(1))
''  GJFile = FreeFile
''  Dim GJInfo() As GJXferRecType
''  ReDim UBSetUprec(1) As UBSetupRecType
''  LoadUBSetUpFile UBSetUpFileNum, UBSetUpLen
''  Get UBSetUpFileNum, 1, UBSetUprec(1)
''  AcctMeth$ = QPTrim$(UBSetUprec(1).MethAcct)
''  If (Len(AcctMeth$) = 0) Then
''    Unload FrmShowPctComp
''    MsgBox "The Utility Account Method Is Not Setup", vbOKOnly, "Invalid Setup Info"
''    GoTo SendExitUB
''  End If
''
''  Select Case AcctMeth$
''  Case "C"
''    InterfaceMethod = 1
''  Case "A"
''    InterfaceMethod = 2
''  Case Else
''    Unload FrmShowPctComp
''    GoTo SendExitUB
''  End Select
''
''  RptFile = FreeFile
''  Open "UBNOTFND.RPT" For Output As RptFile
''  GoSub NotFoundHeader
''
''  'ShowProcessingScrn "Verifying GL Transfer Accounts"
''  FrmShowPctComp.ShowPctComp 20, 100
''
''  For RevCnt = 1 To MaxRevsCnt
''    TempRev$ = QPTrim$(UBSetUprec(1).Revenues(RevCnt).RevName)
''    If Len(TempRev$) = 0 Then
''      NumOfRevs = RevCnt - 1
''      Exit For
''    Else
''      ReDim Preserve GJInfo(1 To RevCnt) As GJXferRecType
''      GJInfo(RevCnt).RevText = TempRev$
''      GJInfo(RevCnt).BAcctInfo.DAcctNo = UBSetUprec(1).BillAcct(RevCnt).DebitAcct
''      GJInfo(RevCnt).BAcctInfo.CAcctNo = UBSetUprec(1).BillAcct(RevCnt).CreditAcct
''      GJInfo(RevCnt).PAcctInfo.DAcctNo = UBSetUprec(1).PayAcct(RevCnt).DebitAcct
''      GJInfo(RevCnt).PAcctInfo.CAcctNo = UBSetUprec(1).PayAcct(RevCnt).CreditAcct
''      If UBSetUprec(1).Revenues(RevCnt).UseDep = "Y" Then
''        GJInfo(RevCnt).DAcctInfo.DAcctNo = UBSetUprec(1).DepAcct(RevCnt).DebitAcct
''        GJInfo(RevCnt).DAcctInfo.CAcctNo = UBSetUprec(1).DepAcct(RevCnt).CreditAcct
''      End If
''    End If
''  Next
''  FrmShowPctComp.ShowPctComp 75, 100
''
''  'check to see if they are valid GL accounts
''  GoSub ValidateGLAccounts
''  FrmShowPctComp.ShowPctComp 80, 100
''  If BadAcct Then
''    Unload FrmShowPctComp
''    GoTo SendExitUB
''  End If
''
'''  ClearBox
'''  QPrintRC "Searching Cash Transactions.", 11, 26, 126
'''  QPrintRC "New Transactions:", 13, 29, Cnf.HiLite
''
''  ReDim UBTransRec(1) As UBTransRecType
''  UBTransRecLen = Len(UBTransRec(1))
''  UBTran = FreeFile
''  Open "UBTRANS.DAT" For Random Shared As UBTran Len = UBTransRecLen
''  NumOfTRecs& = LOF(UBTran) \ UBTransRecLen
''
''  For TCnt& = NumOfTRecs& To 1 Step -1    '1 To NumOfTRecs&
''    Get #UBTran, TCnt&, UBTransRec(1)
''    If UBTransRec(1).CustAcctNo > 0 Then    'so don't get whacked trans
''    If Len(QPTrim$(UBTransRec(1).Posted2GL)) = 0 Or QPTrim$(UBTransRec(1).Posted2GL) = "N" Then
''     If UBTransRec(1).TransDate <= ThruDate% Then
''        If UBTransRec(1).TransDate <= 0 Then 'Exit For
''          FrmShowPctComp.ShowPctComp 100, 100
''          Close
''          GoSub Getout
''        End If
''      'Store trans rec numbers and dates in array
'''          UBTransRec(1).TransDate = 1
'''        End If
''        FoundCnt = FoundCnt + 1
''        ReDim Preserve TranInfo(FoundCnt) As TranRecInfoType
''        TranInfo(FoundCnt).TranDate = UBTransRec(1).TransDate
''        TranInfo(FoundCnt).TranRecNo = TCnt&
''        'If FoundCnt = 30000 Then Exit For
''
''      End If
''    Else
''      NGCnt = NGCnt + 1
''    End If
''    'RSet P2S$ = Str$(FoundCnt)
''    'QPrintRC P2S$, 13, 47, Cnf.HiLite
''    'SmallPause
''    If NGCnt >= 2500 Then
''      FrmShowPctComp.ShowPctComp 1, 1
''      Exit For
''    End If
''    End If
''    'FrmShowPctComp.ShowPctComp TCnt&, NumOfTRecs&
''  Next
''  'FrmShowPctComp.ShowPctComp 1, 1
''  FrmShowPctComp.Label1 = "Sorting Interface Trans File"
''  FrmShowPctComp.cmdCancel.Enabled = False
''  FrmShowPctComp.Show , Me
''
''  If FoundCnt = 0 Then
''    Unload FrmShowPctComp
''    Close
''    Call MainLog("No Trans UB Grab for " + fpDate)
''    MsgBox "No Transactions Found To Interface.", vbOKOnly, "No Trans"
''    GoTo SendExitUB
''  End If
''  FrmShowPctComp.ShowPctComp 25, 100
''
''  QSortTRec TranInfo(), FoundCnt
''  'sort'em by date. oldest first
''  'Array(1), NumElem, Dir, StructSize, MemOff, MemSize
''  FrmShowPctComp.ShowPctComp 50, 100
''  GJFile = FreeFile
''  Open "GLTRXED.DAT" For Random Access Read Write Shared As GJFile Len = GJReclen
''
''  FirstTran = 1
''  ThisDate = TranInfo(1).TranDate
''  For cnt = 1 To FoundCnt
''    If ThisDate <> TranInfo(cnt).TranDate Then
''      ThisDate = TranInfo(cnt).TranDate
''      LastTran = cnt - 1
''      GoSub ProcessThisBunchUB
''      FirstTran = cnt
''    End If
''    FrmShowPctComp.ShowPctComp cnt, FoundCnt
''  Next
''  FrmShowPctComp.Label1 = "Tag Interface Transactions"
''  FrmShowPctComp.cmdCancel.Enabled = False
''  FrmShowPctComp.Show , Me
''
''  LastTran = FoundCnt
''  GoSub ProcessThisBunchUB
''
''  'transactions as interfaced
''  For cnt = 1 To FoundCnt
''    FrmShowPctComp.ShowPctComp cnt, FoundCnt
''    Get #UBTran, TranInfo(cnt).TranRecNo, UBTransRec(1)
''    UBTransRec(1).Posted2GL = "Y"
''    Put #UBTran, TranInfo(cnt).TranRecNo, UBTransRec(1)
''  Next
''  Close
''  Call MainLog("Completed UB Grab " + Str$(FoundCnt) + " for " + fpDate)
''  MsgBox "Transaction Grab Complete.", vbOKOnly, "Complete"
''  'SLEEP 2
''SendExitUB:
''  Exit Sub
''NotFoundHeader:
''  PageNo = PageNo + 1
''  Print #RptFile, "Utility Billing GL Transfer Invalid Account Listing."; Tab(70); "Page:"; PageNo
''
''  Print #RptFile, QPTrim(GLUserName$)
''  Print #RptFile, "Report Date: "; Date$
''  Print #RptFile, "Revenue           Acct. Type              Debit Acct."
''  Print #RptFile, Dash80$
''  NumPrinted = 0
''  Return
''
''PrintBadAcct:
''  If Len(QPTrim$(BadCAcct$)) = 0 Then
''    BadCAcct$ = "Undefined"
''  End If
''
''  Print #RptFile, GJInfo(RevCnt).RevText;
''
''  Select Case ActT
''  Case 1
''    ActPg$ = "Billing"
''  Case 2
''    ActPg$ = "Payment"
''  Case 3
''    ActPg$ = "Deposit"
''  End Select
''  Print #RptFile, Tab(22); ActPg$;
''  Print #RptFile, Tab(43); BadDAcct$; Tab(64); BadCAcct$
''  Return
''
''ProcessThisBunchUB:
''  For RevCnt = 1 To NumOfRevs
''    GJInfo(RevCnt).BAcctInfo.CreditAmt = 0
''    GJInfo(RevCnt).BAcctInfo.DebitAmt = 0
''    GJInfo(RevCnt).PAcctInfo.CreditAmt = 0
''    GJInfo(RevCnt).PAcctInfo.DebitAmt = 0
''    GJInfo(RevCnt).DAcctInfo.CreditAmt = 0
''    GJInfo(RevCnt).DAcctInfo.DebitAmt = 0
''  Next
''
''  For PCnt = FirstTran To LastTran
''    If PCnt = FirstTran Then
''      WorkDate = TranInfo(PCnt).TranDate
''    End If
''    Get #UBTran, TranInfo(PCnt).TranRecNo, UBTransRec(1)
''
''    Select Case InterfaceMethod
''    Case 1      'Cash Central
''      Select Case UBTransRec(1).TransType
''      Case TranUtilityBill      ' 1=Utility bill
''        'no action
''      Case TranLateCharge       ' 2=late charge
''        'no action
''      Case TranReconnectFee     ' 3=reconnect fee
''        'no action
''      Case TranBillPayment      ' 4=Bill Payment
''        For RevCnt = 1 To NumOfRevs
''          GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
''
''          GJInfo(RevCnt).PAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).PAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
''
''        Next
''      Case TranAppliedDeposit   ' 5=Applied Deposit
''        'no action
''      Case TranPenaltyCharge    ' 6=Penalty Charge
''        'no action
''      Case TranDepositPayment   ' 7=Deposit Payment
''        For RevCnt = 1 To NumOfRevs
''          GJInfo(RevCnt).DAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).DAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
''
''          GJInfo(RevCnt).DAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).DAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
''
''        Next
''      Case TranDraftPayment     ' 8=Draft Payment
''        For RevCnt = 1 To NumOfRevs
''          GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
''
''          GJInfo(RevCnt).PAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).DAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
''
''        Next
''      Case TranRefundDeposit    ' 9=Refund Deposit
''        'no action
''      Case TranBeginBalance     '10=Beginning Balance
''        'no action
''      Case TranUpwardAdjustment '11=Upward Adjustments
''        'no action
''      Case TranDownwardAdjustment  '12=Downward Adjustments
''        'no action
''      Case TranOverPayAdjustment   '33=OverPayment Adjustments
''        For RevCnt = 1 To NumOfRevs
''          GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
''          GJInfo(RevCnt).PAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).PAcctInfo.DebitAmt - UBTransRec(1).RevAmt(RevCnt))
''        Next
''      Case TranDepCreditRemoval    '37=Deposit Credit Removal Not to Interface w/GL
''        'No Action !!!
''      Case TranDepPaymentVoid         ' 39=Deposit Void
''        For RevCnt = 1 To NumOfRevs
''          GJInfo(RevCnt).DAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).DAcctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
''          GJInfo(RevCnt).DAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).DAcctInfo.DebitAmt - UBTransRec(1).RevAmt(RevCnt))
''        Next
''      End Select
''
''    Case 2      'Accrual
''      Select Case UBTransRec(1).TransType
''      Case TranUtilityBill      ' 1=Utility bill
''        For RevCnt = 1 To NumOfRevs
''          GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
''          GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
''        Next
''      Case TranLateCharge       ' 2=late charge
''        For RevCnt = 1 To NumOfRevs
''          GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
''          GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
''        Next
''      Case TranReconnectFee     ' 3=reconnect fee
''        For RevCnt = 1 To NumOfRevs
''          GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
''          GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
''        Next
''      Case TranBillPayment      ' 4=Bill Payment
''        For RevCnt = 1 To NumOfRevs
''          GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
''          GJInfo(RevCnt).PAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).PAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
''        Next
''      Case TranAppliedDeposit   ' 5=Applied Deposit
''        'no action
''        'FOR RevCnt = 1 TO NumOfRevs
''        '  GJInfo(RevCnt).dacctInfo.CreditAmt = Round#(GJInfo(RevCnt).dacctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
''        '  GJInfo(RevCnt).pAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).pAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
''        'NEXT
''
''      Case TranPenaltyCharge    ' 6=Penalty Charge
''        For RevCnt = 1 To NumOfRevs
''          GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
''          GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
''        Next
''      Case TranDepositPayment   ' 7=Deposit Payment
''        For RevCnt = 1 To NumOfRevs
''          GJInfo(RevCnt).DAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).DAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
''          GJInfo(RevCnt).DAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).DAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
''        Next
''      Case TranDraftPayment     ' 8=Draft Payment
''        For RevCnt = 1 To NumOfRevs
''          GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
''          GJInfo(RevCnt).PAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).PAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
''        Next
''      Case TranRefundDeposit    ' 9=Refund Deposit
''        'no action
''        '  FOR RevCnt = 1 TO NumOfRevs
''        '    GJInfo(RevCnt).dacctInfo.CreditAmt = Round#(GJInfo(RevCnt).dacctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
''        '    GJInfo(RevCnt).dacctInfo.DebitAmt = Round#(GJInfo(RevCnt).dacctInfo.DebitAmt - UBTransRec(1).RevAmt(RevCnt))
''        '  NEXT
''      Case TranBeginBalance     '10=Beginning Balance
''        'no action
''      Case TranUpwardAdjustment '11=Upward Adjustments
''        For RevCnt = 1 To NumOfRevs
''          GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt + UBTransRec(1).RevAmt(RevCnt))
''          GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt + UBTransRec(1).RevAmt(RevCnt))
''        Next
''      Case TranDownwardAdjustment               '12=Downward Adjustments
''        For RevCnt = 1 To NumOfRevs
''          GJInfo(RevCnt).BAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).BAcctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
''          GJInfo(RevCnt).BAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).BAcctInfo.DebitAmt - UBTransRec(1).RevAmt(RevCnt))
''        Next
''      Case TranOverPayAdjustment   '33=OverPayment Adjustments
''        For RevCnt = 1 To NumOfRevs
''          GJInfo(RevCnt).PAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).PAcctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
''          GJInfo(RevCnt).PAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).PAcctInfo.DebitAmt - UBTransRec(1).RevAmt(RevCnt))
''        Next
''      Case TranDepCreditRemoval    '37=Deposit Credit Removal Not to Interface w/GL
''        'No Action !!!
''      Case TranDepPaymentVoid         ' 39=Deposit Void
''        For RevCnt = 1 To NumOfRevs
''          GJInfo(RevCnt).DAcctInfo.CreditAmt = Round#(GJInfo(RevCnt).DAcctInfo.CreditAmt - UBTransRec(1).RevAmt(RevCnt))
''          GJInfo(RevCnt).DAcctInfo.DebitAmt = Round#(GJInfo(RevCnt).DAcctInfo.DebitAmt - UBTransRec(1).RevAmt(RevCnt))
''        Next
''      End Select
''    End Select
''    'SmallPause
''
''  Next
''
''
''  'NOTE: Journal Rec 1 is the credit, Rec 2 is the debit
''  For RevCnt = 1 To NumOfRevs
''    ReDim GJRec1(1 To 2) As TrEditRecType
''    If GJInfo(RevCnt).BAcctInfo.CreditAmt <> 0 Then
''      GJRec1(1).AcctRec = GJInfo(RevCnt).BAcctInfo.CRecNo
''      GJRec1(1).AcctNum = GJInfo(RevCnt).BAcctInfo.CAcctNo
''      GJRec1(1).AcctName = GJInfo(RevCnt).BAcctInfo.CTitle
''      GJRec1(1).TRDATE = WorkDate
''      GJRec1(1).Ref = Ref$
''      GJRec1(1).CrAmt = GJInfo(RevCnt).BAcctInfo.CreditAmt
''      GJRec1(1).EType = "C"
''      GJRec1(1).Desc = "FROM UTILITIES"
''      GJRec1(1).LDesc = "UB Interface"
''      GJRec1(1).Src = "UB"
''      Put #GJFile, , GJRec1(1)
''    End If
''    If GJInfo(RevCnt).BAcctInfo.DebitAmt <> 0 Then
''      GJRec1(2).AcctRec = GJInfo(RevCnt).BAcctInfo.DRecNo
''      GJRec1(2).AcctNum = GJInfo(RevCnt).BAcctInfo.DAcctNo
''      GJRec1(2).AcctName = GJInfo(RevCnt).BAcctInfo.DTitle
''      GJRec1(2).TRDATE = WorkDate
''      GJRec1(2).Ref = Ref$
''      GJRec1(2).DrAmt = GJInfo(RevCnt).BAcctInfo.DebitAmt
''      GJRec1(2).EType = "D"
''      GJRec1(2).Desc = "FROM UTILITIES"
''      GJRec1(1).LDesc = "UB Interface"
''      GJRec1(2).Src = "UB"
''      Put #GJFile, , GJRec1(2)
''    End If
''  Next
''
''  For RevCnt = 1 To NumOfRevs
''    ReDim GJRec1(1 To 2) As TrEditRecType
''    If GJInfo(RevCnt).PAcctInfo.CreditAmt <> 0 Then
''      GJRec1(1).AcctRec = GJInfo(RevCnt).PAcctInfo.CRecNo
''      GJRec1(1).AcctNum = GJInfo(RevCnt).PAcctInfo.CAcctNo
''      GJRec1(1).AcctName = GJInfo(RevCnt).PAcctInfo.CTitle
''      GJRec1(1).TRDATE = WorkDate
''      GJRec1(1).Ref = Ref$
''      GJRec1(1).CrAmt = GJInfo(RevCnt).PAcctInfo.CreditAmt
''      GJRec1(1).EType = "C"
''      GJRec1(1).Desc = "FROM UTILITIES"
''      GJRec1(1).LDesc = "UB Interface"
''      GJRec1(1).Src = "UB"
''      Put #GJFile, , GJRec1(1)
''    End If
''    If GJInfo(RevCnt).PAcctInfo.DebitAmt <> 0 Then
''      GJRec1(2).AcctRec = GJInfo(RevCnt).PAcctInfo.DRecNo
''      GJRec1(2).AcctNum = GJInfo(RevCnt).PAcctInfo.DAcctNo
''      GJRec1(2).AcctName = GJInfo(RevCnt).PAcctInfo.DTitle
''      GJRec1(2).TRDATE = WorkDate
''      GJRec1(2).Ref = Ref$
''      GJRec1(2).DrAmt = GJInfo(RevCnt).PAcctInfo.DebitAmt
''      GJRec1(2).EType = "D"
''      GJRec1(2).Desc = "FROM UTILITIES"
''      GJRec1(1).LDesc = "UB Interface"
''      GJRec1(2).Src = "UB"
''      Put #GJFile, , GJRec1(2)
''    End If
''  Next
''
''  For RevCnt = 1 To NumOfRevs
''    ReDim GJRec1(1 To 2) As TrEditRecType
''    If GJInfo(RevCnt).DAcctInfo.CreditAmt <> 0 Then
''      GJRec1(1).AcctRec = GJInfo(RevCnt).DAcctInfo.CRecNo
''      GJRec1(1).AcctNum = GJInfo(RevCnt).DAcctInfo.CAcctNo
''      GJRec1(1).AcctName = GJInfo(RevCnt).DAcctInfo.CTitle
''      GJRec1(1).TRDATE = WorkDate
''      GJRec1(1).Ref = Ref$
''      GJRec1(1).CrAmt = GJInfo(RevCnt).DAcctInfo.CreditAmt
''      GJRec1(1).EType = "C"
''      GJRec1(1).Desc = "FROM UTILITIES"
''      GJRec1(1).LDesc = "UB Interface"
''      GJRec1(1).Src = "UB"
''      Put #GJFile, , GJRec1(1)
''    End If
''    If GJInfo(RevCnt).DAcctInfo.DebitAmt <> 0 Then
''      GJRec1(2).AcctRec = GJInfo(RevCnt).DAcctInfo.DRecNo
''      GJRec1(2).AcctNum = GJInfo(RevCnt).DAcctInfo.DAcctNo
''      GJRec1(2).AcctName = GJInfo(RevCnt).DAcctInfo.DTitle
''      GJRec1(2).TRDATE = WorkDate
''      GJRec1(2).Ref = Ref$
''      GJRec1(2).DrAmt = GJInfo(RevCnt).DAcctInfo.DebitAmt
''      GJRec1(2).EType = "D"
''      GJRec1(2).Desc = "FROM UTILITIES"
''      GJRec1(1).LDesc = "UB Interface"
''      GJRec1(2).Src = "UB"
''      Put #GJFile, , GJRec1(2)
''    End If
''  Next
''
''UBBunchReturn:
''  Return
''
''ValidateGLAccounts:
''  BadAcct = False
''  For RevCnt = 1 To NumOfRevs
''    'Billing Accounts
''    If InterfaceMethod = 2 Then
''      'NOTE: We Only check billing accounts if Accural method
''      ActT = 1
''
''      Acct$ = GJInfo(RevCnt).BAcctInfo.DAcctNo
''      Acct$ = QPStrip$(Acct$)
''      Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
''      Acct$ = QPTrim$(Acct$)
''       'AcctR = AcctFind(Acct$)
''      ThisAcct = AcctFind(Acct$) 'AcctFind(GJInfo(RevCnt).BAcctInfo.DAcctNo)
''      If ThisAcct <= 0 Then
''        BadDAcct$ = GJInfo(RevCnt).BAcctInfo.DAcctNo
''        BadAcct = True
''      Else
''        GJInfo(RevCnt).BAcctInfo.DRecNo = ThisAcct
''        GJInfo(RevCnt).BAcctInfo.DTitle = GetAcctTitle$(ThisAcct)
''        BadDAcct$ = "     OK"
''      End If
''
''      Acct$ = GJInfo(RevCnt).BAcctInfo.CAcctNo
''      Acct$ = QPStrip$(Acct$)
''      Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
''      Acct$ = QPTrim$(Acct$)
''      ThisAcct = AcctFind(Acct$)
''      'ThisAcct = AcctFind(GJInfo(RevCnt).BAcctInfo.CAcctNo)
''      If ThisAcct <= 0 Then
''        BadCAcct$ = GJInfo(RevCnt).BAcctInfo.CAcctNo
''        BadAcct = True
''      Else
''        GJInfo(RevCnt).BAcctInfo.CRecNo = ThisAcct
''        GJInfo(RevCnt).BAcctInfo.CTitle = GetAcctTitle$(ThisAcct)
''        BadCAcct$ = "     OK"
''      End If
''      GoSub PrintBadAcct
''    End If
''
''    'Payment Accounts
''    ActT = 2
''      Acct$ = GJInfo(RevCnt).PAcctInfo.DAcctNo
''      Acct$ = QPStrip$(Acct$)
''      Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
''      Acct$ = QPTrim$(Acct$)
''      ThisAcct = AcctFind(Acct$)
''
''    'ThisAcct = AcctFind(GJInfo(RevCnt).PAcctInfo.DAcctNo)
''    If ThisAcct <= 0 Then
''      BadDAcct$ = GJInfo(RevCnt).PAcctInfo.DAcctNo
''      BadAcct = True
''    Else
''      GJInfo(RevCnt).PAcctInfo.DRecNo = ThisAcct
''      GJInfo(RevCnt).PAcctInfo.DTitle = GetAcctTitle$(ThisAcct)
''      BadDAcct$ = "     OK"
''    End If
''
''      Acct$ = GJInfo(RevCnt).PAcctInfo.CAcctNo
''      Acct$ = QPStrip$(Acct$)
''      Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
''      Acct$ = QPTrim$(Acct$)
''      ThisAcct = AcctFind(Acct$)
''
''    'ThisAcct = AcctFind(GJInfo(RevCnt).PAcctInfo.CAcctNo)
''    If ThisAcct <= 0 Then
''      BadCAcct$ = GJInfo(RevCnt).PAcctInfo.CAcctNo
''      BadAcct = True
''    Else
''      GJInfo(RevCnt).PAcctInfo.CRecNo = ThisAcct
''      GJInfo(RevCnt).PAcctInfo.CTitle = GetAcctTitle$(ThisAcct)
''      BadCAcct$ = "     OK"
''    End If
''    GoSub PrintBadAcct
''
''    'Deposit Accounts
''    ActT = 3
''    If UBSetUprec(1).Revenues(RevCnt).UseDep = "Y" Then
''
''      Acct$ = GJInfo(RevCnt).DAcctInfo.DAcctNo
''      Acct$ = QPStrip$(Acct$)
''      Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
''      Acct$ = QPTrim$(Acct$)
''      ThisAcct = AcctFind(Acct$)
''
''      'ThisAcct = AcctFind(GJInfo(RevCnt).DAcctInfo.DAcctNo)
''      If ThisAcct <= 0 Then
''        BadDAcct$ = GJInfo(RevCnt).DAcctInfo.DAcctNo
''        BadAcct = True
''      Else
''        GJInfo(RevCnt).DAcctInfo.DRecNo = ThisAcct
''        GJInfo(RevCnt).DAcctInfo.DTitle = GetAcctTitle$(ThisAcct)
''        BadDAcct$ = "     OK"
''      End If
''    Else
''      BadDAcct$ = "    N/A"
''    End If
''    If UBSetUprec(1).Revenues(RevCnt).UseDep = "Y" Then
''      Acct$ = GJInfo(RevCnt).DAcctInfo.CAcctNo
''      Acct$ = QPStrip$(Acct$)
''      Acct$ = FmtAcct$(Acct$, GLFundLen%, GLAcctLen%, GLDetLen%)
''      Acct$ = QPTrim$(Acct$)
''      ThisAcct = AcctFind(Acct$)
''
''      'ThisAcct = AcctFind(GJInfo(RevCnt).DAcctInfo.CAcctNo)
''      If ThisAcct <= 0 Then
''        BadCAcct$ = GJInfo(RevCnt).DAcctInfo.CAcctNo
''        BadAcct = True
''      Else
''        GJInfo(RevCnt).DAcctInfo.CRecNo = ThisAcct
''        GJInfo(RevCnt).DAcctInfo.CTitle = GetAcctTitle$(ThisAcct)
''        BadCAcct$ = "     OK"
''      End If
''    Else
''      BadCAcct$ = "    N/A"
''    End If
''    GoSub PrintBadAcct
''  Next
''  Close RptFile
''
''  If BadAcct Then
''    Unload FrmShowPctComp
''    MsgBox "Invalid Account(s)Found, Interface File Was Not Created.", vbOKOnly, "Invalid"
''    Call MainLog("UB Grab - NOgo Invalid Accts.")
''    ViewPrint "UBNOTFND.RPT", "GL Transfer Invalid Account List."
''  End If
''  Kill "UBNOTFND.RPT"
''Return
''Getout:
''  Call MainLog("EXIT Grab via error with invalid dates in UB.")
''  'frmCitiCancel.Label1.Caption = "Invalid Dates/Utility Trans Call Software Support."
''  'frmCitiCancel.Show 1
''  MsgBox "Invalid Dates in Utility Trans, Please Call Software Support.", vbCritical, "Warning!!!!"
''  frmGetDistMenu.Show
''  Unload frmGrabTrans
''End Sub

'New VA Tax vers 2.05
Private Sub ExtractVATXPersPay(ThruDate%, TranInfo() As TranRecInfoType, FoundCnt)
  Dim Ref As String, Dash80 As String, P2S As String, TXPGLFile As Integer
  Dim GJReclen As Integer, RptFile As Integer, TaxPAcctRecLen  As Integer
  Dim TaxTranRecLen As Integer, NumOfTRecs As Long, TCnt As Long
  Dim NGCnt As Integer, GJFile As Integer, CDCashAcct As String
  Dim NumEdTrans As Integer, MCFile As Integer, cnt As Integer, CDCashAcctName As String
  Dim FirstTran As Integer, ThisDate As Integer, WorkDate As Integer
  Dim DayCount As Integer, LastTran As Integer, RevCnt As Integer, FundDue As String
  Dim FindCount As Integer, FundCnt As Integer, ThisAcct As Integer
  Dim Acct As String, AcctName As String, T As Integer, BadAcct As Integer
  Dim FoundFund As Integer, PCnt As Integer, Cash As Integer, CDCashRec As Long
  Dim txfile As Integer, InterfaceMethod As Integer, AcctMeth As String, CuryrR As Integer
  Dim TaxYear As Integer, MiddleRec As Integer, TranFile As Integer, DetPad As String
  Dim DebitAcctRecord As Integer, DebitAcctNumber As String, ppcnt As Integer
  Dim DebitAcctName As String, DebitAmt As Double, CreditAcctRecord As Integer
  Dim CreditAcctName As String, CreditAmt As Double, CreditAcctNumber As String
  Dim CDDueAcct As String, CDDueRec As Long, CDDueName As String, PadChars As Integer
  Dim Dsc As String, TrType As Integer, y As Integer, CuryrP As Integer, Curyr As Integer
  Dim GJrecnum As Long
  Dim GJInfo() As GJXferRecType
  Ref$ = "TX" + Left$(Date$, 2) + Mid$(Date$, 4, 2) + Right$(Date$, 2)
'  ReDim TranInfo(1) As TranRecInfoType
  'these are for the personal
  Dim TPrinciple1Pd#(16, 51), TPrinciple2Pd#(16, 51), TPrinciple3Pd#(16, 51), TPrinciple4Pd#(16, 51)
  Dim TPrinciple5Pd#(16, 51), TInterestPd#(16, 51), TPenaltyPd#(16, 51), TPrePaidAmt#(16, 51)
  Dim TPrePaidUsed#(16, 51), TRevOpt1Pd#(16, 51), TRevOpt2Pd#(16, 51), TRevOpt3Pd#(16, 51)
  ReDim Preserve GJInfo(1 To 3) As GJXferRecType
  ReDim TaxPGLAccts(1) As TaxPVAAcctsType
  TaxPAcctRecLen = Len(TaxPGLAccts(1))
  Dim TaxTrans(1) As TaxVATransactionType
  TaxTranRecLen = Len(TaxTrans(1))
  BadAcct = 0
  ReDim TaxSetuprec(1) As TaxVAMasterType
  txfile = FreeFile
  Open "TAXSETUP.DAT" For Random As #txfile Len = Len(TaxSetuprec(1))
  If LOF(txfile) > 0 Then
    Get txfile, 1, TaxSetuprec(1)
  Else
    Unload FrmShowPctComp
    MsgBox "No Tax Setup File Information.", vbOKOnly, "No Setup"
    GoTo TaxEnd
  End If
  'If Central Depository used then will need detail for acct #
    CDActive$ = QPTrim$(TaxSetuprec(1).CntrlDepYN)
    If CDActive$ = "Y" Then
      PadChars = GLDetLen - GLFundLen
      If PadChars > 0 Then
        DetPad$ = String(PadChars, "0")
      End If
      CDCashAcct$ = TaxSetuprec(1).CDCashGL
      CDCashAcct$ = QPStrip$(CDCashAcct$)
      CDCashAcct$ = FmtAcct$(CDCashAcct$, GLFundLen%, GLAcctLen%, GLDetLen%)
      CDCashAcct$ = QPTrim$(CDCashAcct$)
      CDDueAcct$ = QPTrim$(TaxSetuprec(1).CDSubGL)
      CDDueAcct$ = QPStrip$(CDDueAcct$)

      CDCashRec = AcctFind(CDCashAcct$)
      If CDCashRec <= 0 Then
        Unload FrmShowPctComp
        MsgBox "The Account for Central Cash Is Not Setup, Please Correct Before Trying Again.", vbOKOnly, "Invalid Tax Setup"
        GoTo TaxEnd
      Else
        CDCashAcctName$ = GetAcctTitle$(CDCashRec)
      End If
    End If
  FrmShowPctComp.ShowPctComp 10, 100
  AcctMeth$ = QPTrim$(TaxSetuprec(1).AcctgMethod)
  CuryrP = Right$(Num2Date$(TaxSetuprec(1).PTaxYear), 4)
  If (Len(AcctMeth$) = 0) Then
    Unload FrmShowPctComp
    MsgBox "The Accounting Method Is Not Setup, Please Correct Before Trying Again.", vbOKOnly, "Invalid Tax Setup"
    GoTo TaxEnd
  End If
  Select Case AcctMeth$
  Case "C"
    InterfaceMethod = 1
  Case "A"
    InterfaceMethod = 2
  Case "M"
    InterfaceMethod = 3
  Case Else
    Unload FrmShowPctComp
    MsgBox "The Accounting Method Is Invalid, Please Correct Before Trying Again.", vbOKOnly, "Invalid Tax Setup"
    GoTo EndTax
  End Select
  Close txfile

  GJReclen = Len(GJRec(1))

'  If Exist("TAXPGLACT.DAT") Then
'    TXPGLFile = FreeFile
'    Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
'    Get TXPGLFile, 1, TaxPGLAccts(1)
'  Else
'    Unload FrmShowPctComp
'    MsgBox "Tax Accounts Not Setup,Interface File Not Created.", vbOKOnly, "Tax Acct Setup Invalid"
'    GoTo EndTax
'  Close
'  End If
'  FrmShowPctComp.Label1 = "Sorting Interface Trans File"
'  FrmShowPctComp.cmdCancel.Enabled = False
'  FrmShowPctComp.Show , Me

  'Now Process the Transactions
  'ClearBox

 ' QPrintRC "Searching Cash Transactions.", 11, 26, 126
 ' QPrintRC "New Transactions:", 13, 29, Cnf.HiLite
  FrmShowPctComp.ShowPctComp 65, 100
  TranFile = FreeFile
  Open "TAXTRANS.DAT" For Random Shared As TranFile Len = TaxTranRecLen
  NumOfTRecs& = LOF(TranFile) \ TaxTranRecLen
'  For TCnt& = NumOfTRecs& To 1 Step -1
'    Get #TranFile, TCnt&, TaxTrans(1)
'    If Len(QPTrim$(TaxTrans(1).Posted2GL)) = 0 Or QPTrim$(TaxTrans(1).Posted2GL) = "N" Then
'      If TaxTrans(1).BillType = "R" Then
'        'Store trans rec numbers and dates in array
'        If TaxTrans(1).TransDate <= ThruDate% Then
'          FoundCnt = FoundCnt + 1
'          ReDim Preserve TranInfo(FoundCnt) As TranRecInfoType
'          TranInfo(FoundCnt).TranDate = TaxTrans(1).TransDate
'          TranInfo(FoundCnt).TranRecNo = TCnt&
'        End If
'      End If
'    Else
'      NGCnt = NGCnt + 1
'    End If
'    P2S$ = Str$(FoundCnt)
'    'QPrintRC P2S$, 13, 47, Cnf.HiLite
'    'SmallPause
'    If NGCnt >= 2500 Then Exit For
'  Next
'  'FrmShowPctComp.ShowPctComp 40, 100
'  If FoundCnt = 0 Then
'    Close
'    Unload FrmShowPctComp
'    Call MainLog("No Tx to Grab " + Str$(FoundCnt) + " for " + fpdate)
'    MsgBox "No Transactions Found to Interface.", vbOKOnly, "No Trans"
'    GoTo EndTax
'  End If
'  FrmShowPctComp.Label1 = "Sorting Interface Trans File"
'  FrmShowPctComp.cmdCancel.Enabled = False
'  FrmShowPctComp.Show , Me
'  FrmShowPctComp.ShowPctComp 15, 100
'  SortTRec TranInfo(), FoundCnt      'sort'em by date. oldest first

  GJFile = FreeFile
  Open "GLTRXED.DAT" For Random Access Read Write Shared As GJFile Len = GJReclen
  GJrecnum = LOF(GJFile) \ GJReclen
  FrmShowPctComp.ShowPctComp 35, 100
  FirstTran = 1
  ThisDate = TranInfo(1).TranDate
  For cnt = 1 To FoundCnt
    FrmShowPctComp.ShowPctComp cnt, FoundCnt
    If ThisDate <> TranInfo(cnt).TranDate Then
      ThisDate = TranInfo(cnt).TranDate
      LastTran = cnt - 1
      GoSub ProcessThisBunchTX
      FirstTran = cnt
    End If
  Next cnt
  FrmShowPctComp.Label1 = "Tag Interface Transactions"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me

  LastTran = FoundCnt
  GoSub ProcessThisBunchTX
  BadTxAcct = BadAcct

'  If BadAcct > 0 Then
'    Close
'    Unload FrmShowPctComp
'    Call MainLog("Error - TX Grab Not Created Due to invalid or missing accts.")
'    MsgBox "Errors Found, Interface File Not Created. Please Review Report.", vbOKOnly, "Errors"
'    frmReportOpt.Show 1
'    If rptopt = 1 Then
'      frmGetDistMenu.PrnEditList 2
'    ElseIf rptopt = 2 Then
'      frmGetDistMenu.PrnEditList2 2
'    End If
'    Kill "GLTRXED.DAT"
'    Exit Sub
'  End If
'  'transactions as interfaced
'
'  For cnt = 1 To FoundCnt
'    FrmShowPctComp.ShowPctComp cnt, FoundCnt
'    Get #TranFile, TranInfo(cnt).TranRecNo, TaxTrans(1)
'    TaxTrans(1).Posted2GL = "Y"
'    Put #TranFile, TranInfo(cnt).TranRecNo, TaxTrans(1)
'  Next cnt
'  Close
'  'SLEEP 2
'  Call MainLog("TX Grab Complete for PersPay" + fpdate)
'  MsgBox "Transaction Grab Complete Pers/Pay.", vbOKOnly, "Complete"
'  GoTo EndTax

EndTax:
  Unload FrmShowPctComp
  Close
  Exit Sub

ProcessThisBunchTX:      'Initialize for This Set
GoSub Clearouttots
  For PCnt = FirstTran To LastTran
    If PCnt = FirstTran Then
      WorkDate = TranInfo(PCnt).TranDate
    End If
    Get #TranFile, TranInfo(PCnt).TranRecNo, TaxTrans(1)
    'Now Decipher by Type and Year
    Select Case TaxTrans(1).TranType
      Case 1:
        TrType = 1
      Case 2:
        TrType = 2
      Case 3:
        TrType = 3
      Case 4:
        TrType = 4
      Case 5:
        TrType = 5
      Case 6:
        TrType = 6
      Case 7:
        TrType = 7
      Case 9:
        TrType = 9
      Case 10:
        TrType = 10
      Case 11:
        TrType = 11
      Case 12:
        TrType = 12
      Case 13:
        TrType = 13
      Case 14:
        TrType = 14
      Case 21:
        TrType = 16
      Case 22:
        TrType = 8
      Case 24:
        TrType = 15
      End Select
      If TaxTrans(1).BillType = "P" Then
        TaxYear = TaxTrans(1).TaxYear
        Curyr = CuryrP
        If TaxYear < 1 Then TaxYear = Curyr
        TaxYear = TaxYear - 1979                'Reduce Based on 1980 being = 1
        TPrinciple1Pd#(TrType, TaxYear) = TPrinciple1Pd#(TrType, TaxYear) + TaxTrans(1).Revenue.Principle1Pd
        TPrinciple2Pd#(TrType, TaxYear) = TPrinciple2Pd#(TrType, TaxYear) + TaxTrans(1).Revenue.Principle2Pd
        TPrinciple3Pd#(TrType, TaxYear) = TPrinciple3Pd#(TrType, TaxYear) + TaxTrans(1).Revenue.Principle3Pd
        TPrinciple4Pd#(TrType, TaxYear) = TPrinciple4Pd#(TrType, TaxYear) + TaxTrans(1).Revenue.Principle4Pd
        TPrinciple5Pd#(TrType, TaxYear) = TPrinciple5Pd#(TrType, TaxYear) + TaxTrans(1).Revenue.Principle5Pd
        TInterestPd#(TrType, TaxYear) = TInterestPd#(TrType, TaxYear) + TaxTrans(1).Revenue.InterestPd
        TPenaltyPd#(TrType, TaxYear) = TPenaltyPd#(TrType, TaxYear) + TaxTrans(1).Revenue.PenaltyPd
        TRevOpt1Pd#(TrType, TaxYear) = TRevOpt1Pd#(TrType, TaxYear) + TaxTrans(1).Revenue.RevOpt1Pd
        TRevOpt2Pd#(TrType, TaxYear) = TRevOpt2Pd#(TrType, TaxYear) + TaxTrans(1).Revenue.RevOpt2Pd
        TRevOpt3Pd#(TrType, TaxYear) = TRevOpt3Pd#(TrType, TaxYear) + TaxTrans(1).Revenue.RevOpt3Pd
        TPrePaidAmt#(TrType, TaxYear) = TPrePaidAmt#(TrType, TaxYear) + TaxTrans(1).Revenue.PrePaidAmt
        TPrePaidUsed#(TrType, TaxYear) = TPrePaidUsed#(TrType, TaxYear) + TaxTrans(1).Revenue.PrePaidUsed
      End If
  Next PCnt
  'TranType2 Payments-2,Payment w/prepay-21, Prepay only-22
  For T% = 1 To 51
    If TPrinciple1Pd#(2, T%) > 0 Or TPrinciple1Pd#(16, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PersCRAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PersCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TPrinciple1Pd#(2, T%) + TPrinciple1Pd#(16, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PersDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PersDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TPrinciple1Pd#(2, T%) + TPrinciple1Pd#(16, T%))
        Dsc$ = "PaymentP"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PersDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PersDBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(TPrinciple1Pd#(2, T%) + TPrinciple1Pd#(16, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PersCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PersCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TPrinciple1Pd#(2, T%) + TPrinciple1Pd#(16, T%))
      Dsc$ = "PaymentP"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXPGLFile
    End If
    If TPrinciple2Pd#(2, T%) > 0 Or TPrinciple2Pd#(16, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MTCRAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MTCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TPrinciple2Pd#(2, T%) + TPrinciple2Pd#(16, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MTDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MTDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TPrinciple2Pd#(2, T%) + TPrinciple2Pd#(16, T%))
        Dsc$ = "PaymentP"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MTDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MTDBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(TPrinciple2Pd#(2, T%) + TPrinciple2Pd#(16, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MTCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MTCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TPrinciple2Pd#(2, T%) + TPrinciple2Pd#(16, T%))
      Dsc$ = "PaymentP"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXPGLFile
    End If
    If TPrinciple3Pd#(2, T%) > 0 Or TPrinciple3Pd#(16, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MCCRAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MCCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TPrinciple3Pd#(2, T%) + TPrinciple3Pd#(16, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MCDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MCDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TPrinciple3Pd#(2, T%) + TPrinciple3Pd#(16, T%))
        Dsc$ = "PaymentP"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MCDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MCDBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(TPrinciple3Pd#(2, T%) + TPrinciple3Pd#(16, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MCCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MCCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TPrinciple3Pd#(2, T%) + TPrinciple3Pd#(16, T%))
      Dsc$ = "PaymentP"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXPGLFile
    End If
    If TPrinciple4Pd#(2, T%) > 0 Or TPrinciple4Pd#(16, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).FECRAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).FECRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TPrinciple4Pd#(2, T%) + TPrinciple4Pd#(16, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).FEDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).FEDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TPrinciple4Pd#(2, T%) + TPrinciple4Pd#(16, T%))
        Dsc$ = "PaymentP"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).FEDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).FEDBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(TPrinciple4Pd#(2, T%) + TPrinciple4Pd#(16, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).FECRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).FECRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TPrinciple4Pd#(2, T%) + TPrinciple4Pd#(16, T%))
      Dsc$ = "PaymentP"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXPGLFile
    End If
    If TPrinciple5Pd#(2, T%) > 0 Or TPrinciple5Pd#(16, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MHCRAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MHCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TPrinciple5Pd#(2, T%) + TPrinciple5Pd#(16, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MHDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MHDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TPrinciple5Pd#(2, T%) + TPrinciple5Pd#(16, T%))
        Dsc$ = "PaymentP"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MHDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MHDBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(TPrinciple5Pd#(2, T%) + TPrinciple5Pd#(16, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MHCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MHCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TPrinciple5Pd#(2, T%) + TPrinciple5Pd#(16, T%))
      Dsc$ = "PaymentP"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXPGLFile
    End If
    If TRevOpt1Pd#(2, T%) > 0 Or TRevOpt1Pd#(16, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt1CRAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt1CRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TRevOpt1Pd#(2, T%) + TRevOpt1Pd#(16, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt1DBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt1DBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TRevOpt1Pd#(2, T%) + TRevOpt1Pd#(16, T%))
        Dsc$ = "PaymentO1"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt1DBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt1DBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(TRevOpt1Pd#(2, T%) + TRevOpt1Pd#(16, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt1CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt1CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TRevOpt1Pd#(2, T%) + TRevOpt1Pd#(16, T%))
      Dsc$ = "PaymentO1"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXPGLFile
    End If
    If TRevOpt2Pd#(2, T%) > 0 Or TRevOpt2Pd#(16, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt2CRAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt2CRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TRevOpt2Pd#(2, T%) + TRevOpt2Pd#(16, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt2DBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt2DBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TRevOpt2Pd#(2, T%) + TRevOpt2Pd#(16, T%))
        Dsc$ = "PaymentO2"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt2DBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt2DBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(TRevOpt2Pd#(2, T%) + TRevOpt2Pd#(16, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt2CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt2CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TRevOpt2Pd#(2, T%) + TRevOpt2Pd#(16, T%))
      Dsc$ = "PaymentO2"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXPGLFile
    End If
    If TRevOpt3Pd#(2, T%) > 0 Or TRevOpt3Pd#(16, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt3CRAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt3CRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TRevOpt3Pd#(2, T%) + TRevOpt3Pd#(16, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt3DBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt3DBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TRevOpt3Pd#(2, T%) + TRevOpt3Pd#(16, T%))
        Dsc$ = "PaymentO3"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt3DBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt3DBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(TRevOpt3Pd#(2, T%) + TRevOpt3Pd#(16, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt3CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt3CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TRevOpt3Pd#(2, T%) + TRevOpt3Pd#(16, T%))
      Dsc$ = "PaymentO3"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXPGLFile
    End If
    If TInterestPd#(2, T%) > 0 Or TInterestPd#(16, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).IntCRAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).IntCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TInterestPd#(2, T%) + TInterestPd#(16, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).IntDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).IntDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TInterestPd#(2, T%) + TInterestPd#(16, T%))
        Dsc$ = "PaymentI"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).IntDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).IntDBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(TInterestPd#(2, T%) + TInterestPd#(16, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).IntCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).IntCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TInterestPd#(2, T%) + TInterestPd#(16, T%))
      Dsc$ = "PaymentI"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXPGLFile
    End If
    If TPenaltyPd#(2, T%) > 0 Or TPenaltyPd#(16, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PenCRAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PenCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TPenaltyPd#(2, T%) + TPenaltyPd#(16, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PenDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PenDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TPenaltyPd#(2, T%) + TPenaltyPd#(16, T%))
        Dsc$ = "PaymentPen"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PenDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PenDBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(TPenaltyPd#(2, T%) + TPenaltyPd#(16, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PenCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PenCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TPenaltyPd#(2, T%) + TPenaltyPd#(16, T%))
      Dsc$ = "PaymentPen"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXPGLFile
    End If
'Prepayment include for either type
    If TPrePaidAmt#(8, T%) > 0 Or TPrePaidAmt#(16, T%) > 0 Then
      TXPGLFile = FreeFile   'use the real cash acct
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PersDBAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PersDBAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(TPrePaidAmt#(8, T%) + TPrePaidAmt#(16, T%))
      ThisAcct = AcctFind(TaxSetuprec(1).OverPayGLNum)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxSetuprec(1).OverPayGLNum
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TPrePaidAmt#(8, T%) + TPrePaidAmt#(16, T%))
      Dsc$ = "PaymentPre"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctRecord = CDCashRec
        DebitAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          DebitAcctName$ = CDCashAcctName$
        Else
          DebitAcctName$ = "Blank Acct"
        End If
        CreditAcctNumber$ = Left$(CreditAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = CreditAcctNumber$ + DetPad$
        Else
          FundDue$ = CreditAcctNumber$
        End If
        CreditAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        CreditAcctNumber$ = FmtAcct$(CreditAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(CreditAcctNumber$)
        CreditAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXPGLFile
    End If
  Next T%
'@@@@@@@@@@@@@@@@@@end of payment types, 2, 21, 22
''''''''''''''''''''''''''''''''''
  'Overpayment Apply during billing -9 and -24(amts applied during up adj)
  For T% = 1 To 51
    If TPrePaidUsed#(9, T%) > 0 Or TPrePaidUsed#(15, T%) > 0 Then
      ThisAcct = AcctFind(TaxSetuprec(1).OverPayGLNum)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxSetuprec(1).OverPayGLNum
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(TPrePaidUsed#(9, T%) + TPrePaidUsed#(15, T%))
      Dsc$ = "OverpayUsed"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = DebitAcctRecord
      GJRec1(1).AcctNum = DebitAcctNumber$
      GJRec1(1).AcctName = DebitAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).DrAmt = DebitAmt#
      GJRec1(1).EType = "D"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      GJrecnum = LOF(GJFile) \ GJReclen
      Put #GJFile, GJrecnum + 1, GJRec1(1)
    End If
    If TPrinciple1Pd#(9, T%) > 0 Or TPrinciple1Pd#(15, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PersCRAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PersCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TPrinciple1Pd#(9, T%) + TPrinciple1Pd#(15, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PersDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PersDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TPrinciple1Pd#(9, T%) + TPrinciple1Pd#(15, T%))
        Dsc$ = "OverPayApplyP"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PersCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PersCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TPrinciple1Pd#(9, T%) + TPrinciple1Pd#(15, T%))
      Dsc$ = "OverPayApplyP"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = CreditAcctRecord
      GJRec1(1).AcctNum = CreditAcctNumber$
      GJRec1(1).AcctName = CreditAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = CreditAmt#
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      GJrecnum = LOF(GJFile) \ GJReclen
      Put #GJFile, GJrecnum + 1, GJRec1(1)
      Close TXPGLFile
    End If
    If TPrinciple2Pd#(9, T%) > 0 Or TPrinciple2Pd#(15, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MTCRAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MTCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TPrinciple2Pd#(9, T%) + TPrinciple2Pd#(15, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MTDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MTDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TPrinciple2Pd#(9, T%) + TPrinciple2Pd#(15, T%))
        Dsc$ = "OverPayApplyP"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MTCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MTCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TPrinciple2Pd#(9, T%) + TPrinciple2Pd#(15, T%))
      Dsc$ = "OverPayApplyP"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = CreditAcctRecord
      GJRec1(1).AcctNum = CreditAcctNumber$
      GJRec1(1).AcctName = CreditAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = CreditAmt#
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      GJrecnum = LOF(GJFile) \ GJReclen
      Put #GJFile, GJrecnum + 1, GJRec1(1)
      Close TXPGLFile
    End If
    If TPrinciple3Pd#(9, T%) > 0 Or TPrinciple3Pd#(15, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MCCRAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MCCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TPrinciple3Pd#(9, T%) + TPrinciple3Pd#(15, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MCDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MCDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TPrinciple3Pd#(9, T%) + TPrinciple3Pd#(15, T%))
        Dsc$ = "OverPayApplyP"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MCCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MCCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TPrinciple3Pd#(9, T%) + TPrinciple3Pd#(15, T%))
      Dsc$ = "OverPayApplyP"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = CreditAcctRecord
      GJRec1(1).AcctNum = CreditAcctNumber$
      GJRec1(1).AcctName = CreditAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = CreditAmt#
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      GJrecnum = LOF(GJFile) \ GJReclen
      Put #GJFile, GJrecnum + 1, GJRec1(1)
      Close TXPGLFile
    End If
    If TPrinciple4Pd#(9, T%) > 0 Or TPrinciple4Pd#(15, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).FECRAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).FECRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TPrinciple4Pd#(9, T%) + TPrinciple4Pd#(15, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).FEDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).FEDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TPrinciple4Pd#(9, T%) + TPrinciple4Pd#(15, T%))
        Dsc$ = "OverPayApplyP"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).FECRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).FECRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TPrinciple4Pd#(9, T%) + TPrinciple4Pd#(15, T%))
      Dsc$ = "OverPayApplyP"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = CreditAcctRecord
      GJRec1(1).AcctNum = CreditAcctNumber$
      GJRec1(1).AcctName = CreditAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = CreditAmt#
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      GJrecnum = LOF(GJFile) \ GJReclen
      Put #GJFile, GJrecnum + 1, GJRec1(1)
      Close TXPGLFile
    End If
    If TPrinciple5Pd#(9, T%) > 0 Or TPrinciple5Pd#(15, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MHCRAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MHCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TPrinciple5Pd#(9, T%) + TPrinciple5Pd#(15, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MHDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MHDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TPrinciple5Pd#(9, T%) + TPrinciple5Pd#(15, T%))
        Dsc$ = "OverPayApplyP"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MHCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MHCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TPrinciple5Pd#(9, T%) + TPrinciple5Pd#(15, T%))
      Dsc$ = "OverPayApplyP"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = CreditAcctRecord
      GJRec1(1).AcctNum = CreditAcctNumber$
      GJRec1(1).AcctName = CreditAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = CreditAmt#
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      GJrecnum = LOF(GJFile) \ GJReclen
      Put #GJFile, GJrecnum + 1, GJRec1(1)
      Close TXPGLFile
    End If
    If TRevOpt1Pd#(9, T%) > 0 Or TRevOpt1Pd#(15, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt1CRAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt1CRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(TRevOpt1Pd#(9, T%) + TRevOpt1Pd#(15, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt1DBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt1DBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(TRevOpt1Pd#(9, T%) + TRevOpt1Pd#(15, T%))
        Dsc$ = "OverPayApplyO1"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt1CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt1CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TRevOpt1Pd#(9, T%) + TRevOpt1Pd#(15, T%))
      Dsc$ = "OverPayApplyO1"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = CreditAcctRecord
      GJRec1(1).AcctNum = CreditAcctNumber$
      GJRec1(1).AcctName = CreditAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = CreditAmt#
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      GJrecnum = LOF(GJFile) \ GJReclen
      Put #GJFile, GJrecnum + 1, GJRec1(1)
      Close TXPGLFile
    End If
    If TRevOpt2Pd#(9, T%) > 0 Or TRevOpt2Pd#(15, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt2CRAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt2CRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(TRevOpt2Pd#(9, T%) + TRevOpt2Pd#(15, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt2DBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt2DBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(TRevOpt2Pd#(9, T%) + TRevOpt2Pd#(15, T%))
        Dsc$ = "OverPayApplyO2"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt2CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt2CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TRevOpt2Pd#(9, T%) + TRevOpt2Pd#(15, T%))
      Dsc$ = "OverPayApplyO2"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = CreditAcctRecord
      GJRec1(1).AcctNum = CreditAcctNumber$
      GJRec1(1).AcctName = CreditAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = CreditAmt#
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      GJrecnum = LOF(GJFile) \ GJReclen
      Put #GJFile, GJrecnum + 1, GJRec1(1)
      Close TXPGLFile
    End If
    If TRevOpt3Pd#(9, T%) > 0 Or TRevOpt3Pd#(15, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt3CRAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt3CRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(TRevOpt3Pd#(9, T%) + TRevOpt3Pd#(15, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt3DBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt3DBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(TRevOpt3Pd#(9, T%) + TRevOpt3Pd#(15, T%))
        Dsc$ = "OverPayApplyO3"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt3CRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt3CRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TRevOpt3Pd#(9, T%) + TRevOpt3Pd#(15, T%))
      Dsc$ = "OverPayApplyO3"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = CreditAcctRecord
      GJRec1(1).AcctNum = CreditAcctNumber$
      GJRec1(1).AcctName = CreditAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = CreditAmt#
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      GJrecnum = LOF(GJFile) \ GJReclen
      Put #GJFile, GJrecnum + 1, GJRec1(1)
      Close TXPGLFile
    End If
    If TInterestPd#(9, T%) > 0 Or TInterestPd#(15, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).IntCRAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).IntCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TInterestPd#(9, T%) + TInterestPd#(15, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).IntDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).IntDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TInterestPd#(9, T%) + TInterestPd#(15, T%))
        Dsc$ = "OverPayApplyI"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).IntCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).IntCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TInterestPd#(9, T%) + TInterestPd#(15, T%))
      Dsc$ = "OverPayApplyI"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = CreditAcctRecord
      GJRec1(1).AcctNum = CreditAcctNumber$
      GJRec1(1).AcctName = CreditAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = CreditAmt#
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      GJrecnum = LOF(GJFile) \ GJReclen
      Put #GJFile, GJrecnum + 1, GJRec1(1)
      Close TXPGLFile
    End If
    If TPenaltyPd#(9, T%) > 0 Or TPenaltyPd#(15, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PenCRAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PenCRAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round(TPenaltyPd#(9, T%) + TPenaltyPd#(15, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PenDBAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PenDBAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round(TPenaltyPd#(9, T%) + TPenaltyPd#(15, T%))
        Dsc$ = "OverPayApplyPen"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PenCRAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PenCRAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TPenaltyPd#(9, T%) + TPenaltyPd#(15, T%))
      Dsc$ = "OverPayApplyPen"
      ReDim GJRec1(1) As TrEditRecType
      GJRec1(1).AcctRec = CreditAcctRecord
      GJRec1(1).AcctNum = CreditAcctNumber$
      GJRec1(1).AcctName = CreditAcctName$
      GJRec1(1).TRDATE = WorkDate
      GJRec1(1).Ref = Ref$
      GJRec1(1).CrAmt = CreditAmt#
      GJRec1(1).EType = "C"
      GJRec1(1).Desc = "FRMTX " + Dsc$
      GJRec1(1).LDesc = "TX Interface"
      GJRec1(1).Src = "TX"
      GJrecnum = LOF(GJFile) \ GJReclen
      Put #GJFile, GJrecnum + 1, GJRec1(1)
      Close TXPGLFile
    End If
  Next T%
'@@@@@@@@@@@@@@ end of tran type 9 and apply prepay amt side of 24 up adj
''''''''''''''''''''''''''''
'tran 7 adjust pay down  and  10 paydown w/prep for interface 3(Modified Accrual) do bill charge again
  For T% = 1 To 51
    If TPrinciple1Pd#(7, T%) > 0 Or TPrinciple1Pd#(10, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PersDBAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PersDBAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(TPrinciple1Pd#(7, T%) + TPrinciple1Pd#(10, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PersCRAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PersCRAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(TPrinciple1Pd#(7, T%) + TPrinciple1Pd#(10, T%))
        Dsc$ = "AdjPayP"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PersCRAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PersCRAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round#(TPrinciple1Pd#(7, T%) + TPrinciple1Pd#(10, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PersDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PersDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TPrinciple1Pd#(7, T%) + TPrinciple1Pd#(10, T%))
      Dsc$ = "AdjPayP"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXPGLFile
    End If
    If TPrinciple2Pd#(7, T%) > 0 Or TPrinciple2Pd#(10, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MTDBAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MTDBAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(TPrinciple2Pd#(7, T%) + TPrinciple2Pd#(10, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MTCRAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MTCRAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(TPrinciple2Pd#(7, T%) + TPrinciple2Pd#(10, T%))
        Dsc$ = "AdjPayP"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MTCRAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MTCRAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round#(TPrinciple2Pd#(7, T%) + TPrinciple2Pd#(10, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MTDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MTDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TPrinciple2Pd#(7, T%) + TPrinciple2Pd#(10, T%))
      Dsc$ = "AdjPayP"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXPGLFile
    End If
    If TPrinciple3Pd#(7, T%) > 0 Or TPrinciple3Pd#(10, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MCDBAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MCDBAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(TPrinciple3Pd#(7, T%) + TPrinciple3Pd#(10, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MCCRAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MCCRAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(TPrinciple3Pd#(7, T%) + TPrinciple3Pd#(10, T%))
        Dsc$ = "AdjPayP"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MCCRAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MCCRAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round#(TPrinciple3Pd#(7, T%) + TPrinciple3Pd#(10, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MCDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MCDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TPrinciple3Pd#(7, T%) + TPrinciple3Pd#(10, T%))
      Dsc$ = "AdjPayP"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXPGLFile
    End If
    If TPrinciple4Pd#(7, T%) > 0 Or TPrinciple4Pd#(10, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).FEDBAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).FEDBAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(TPrinciple4Pd#(7, T%) + TPrinciple4Pd#(10, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).FECRAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).FECRAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(TPrinciple4Pd#(7, T%) + TPrinciple4Pd#(10, T%))
        Dsc$ = "AdjPayP"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).FECRAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).FECRAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round#(TPrinciple4Pd#(7, T%) + TPrinciple4Pd#(10, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).FEDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).FEDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TPrinciple4Pd#(7, T%) + TPrinciple4Pd#(10, T%))
      Dsc$ = "AdjPayP"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXPGLFile
    End If
    If TPrinciple5Pd#(7, T%) > 0 Or TPrinciple5Pd#(10, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MHDBAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MHDBAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(TPrinciple5Pd#(7, T%) + TPrinciple5Pd#(10, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MHCRAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MHCRAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(TPrinciple5Pd#(7, T%) + TPrinciple5Pd#(10, T%))
        Dsc$ = "AdjPayP"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MHCRAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MHCRAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round#(TPrinciple5Pd#(7, T%) + TPrinciple5Pd#(10, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MHDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MHDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TPrinciple5Pd#(7, T%) + TPrinciple5Pd#(10, T%))
      Dsc$ = "AdjPayP"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXPGLFile
    End If
    If TRevOpt1Pd#(7, T%) > 0 Or TRevOpt1Pd#(10, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt1DBAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt1DBAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(TRevOpt1Pd#(7, T%) + TRevOpt1Pd#(10, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt1CRAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt1CRAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(TRevOpt1Pd#(7, T%) + TRevOpt1Pd#(10, T%))
        Dsc$ = "AdjPayO1"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt1CRAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt1CRAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round#(TRevOpt1Pd#(7, T%) + TRevOpt1Pd#(10, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt1DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt1DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TRevOpt1Pd#(7, T%) + TRevOpt1Pd#(10, T%))
      Dsc$ = "AdjPayO1"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXPGLFile
    End If
    If TRevOpt2Pd#(7, T%) > 0 Or TRevOpt2Pd#(10, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt2DBAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt2DBAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(TRevOpt2Pd#(7, T%) + TRevOpt2Pd#(10, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt2CRAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt2CRAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(TRevOpt2Pd#(7, T%) + TRevOpt2Pd#(10, T%))
        Dsc$ = "AdjPayO2"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt2CRAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt2CRAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round#(TRevOpt2Pd#(7, T%) + TRevOpt2Pd#(10, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt2DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt2DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TRevOpt2Pd#(7, T%) + TRevOpt2Pd#(10, T%))
      Dsc$ = "AdjPayO2"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXPGLFile
    End If
    If TRevOpt3Pd#(7, T%) > 0 Or TRevOpt3Pd#(10, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt3DBAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt3DBAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(TRevOpt3Pd#(7, T%) + TRevOpt3Pd#(10, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt3CRAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt3CRAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(TRevOpt3Pd#(7, T%) + TRevOpt3Pd#(10, T%))
        Dsc$ = "AdjPayO3"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt3CRAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt3CRAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round#(TRevOpt3Pd#(7, T%) + TRevOpt3Pd#(10, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt3DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt3DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TRevOpt3Pd#(7, T%) + TRevOpt3Pd#(10, T%))
      Dsc$ = "AdjPayO3"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXPGLFile
    End If
    If TInterestPd#(7, T%) > 0 Or TInterestPd#(10, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).IntDBAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).IntDBAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(TInterestPd#(7, T%) + TInterestPd#(10, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).IntCRAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).IntCRAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(TInterestPd#(7, T%) + TInterestPd#(10, T%))
        Dsc$ = "AdjPayI"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).IntCRAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).IntCRAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round#(TInterestPd#(7, T%) + TInterestPd#(10, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).IntDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).IntDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TInterestPd#(7, T%) + TInterestPd#(10, T%))
      Dsc$ = "AdjPayI"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXPGLFile
    End If
    If TPenaltyPd#(7, T%) > 0 Or TPenaltyPd#(10, T%) > 0 Then
      If InterfaceMethod = 3 Then
        TXPGLFile = FreeFile
        Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
        Get TXPGLFile, 1, TaxPGLAccts(1)
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PenDBAcct)
        DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PenDBAcct
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Invaild Acct"
        End If
        DebitAmt# = Round#(TPenaltyPd#(7, T%) + TPenaltyPd#(10, T%))
        ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PenCRAcct)
        CreditAcctRecord = ThisAcct
        CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PenCRAcct
        If ThisAcct > 0 Then
          CreditAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          CreditAcctName$ = "Undefined"
        End If
        CreditAmt# = Round#(TPenaltyPd#(7, T%) + TPenaltyPd#(10, T%))
        Dsc$ = "AdjPayPen"
        GoSub PostToGeneralJournal
        Close TXPGLFile
      End If
      TXPGLFile = FreeFile
      Open "TAXPGLACT.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PenCRAcct)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PenCRAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round#(TPenaltyPd#(7, T%) + TPenaltyPd#(10, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PenDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PenDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round#(TPenaltyPd#(7, T%) + TPenaltyPd#(10, T%))
      Dsc$ = "AdjPayPen"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXPGLFile
    End If

  '^*&^*&^*&^* 10 and 11 prepayment adj down
    If TPrePaidUsed#(10, T%) > 0 Or TPrePaidUsed#(11, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLACt.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxSetuprec(1).OverPayGLNum)
      DebitAcctRecord = ThisAcct
      DebitAcctNumber$ = TaxSetuprec(1).OverPayGLNum
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Undefined"
      End If
      DebitAmt# = Round(TPrePaidUsed#(10, T%) + TPrePaidUsed#(11, T%))
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PersDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PersDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = Round(TPrePaidUsed#(10, T%) + TPrePaidUsed#(11, T%))
      Dsc$ = "AdjPre"
      GoSub PostToGeneralJournal
      'If Central Depository use Central Deposi Cash Acct
      If CDActive$ = "Y" Then
        DebitAcctNumber$ = Left$(DebitAcctNumber$, GLFundLen%)
        If PadChars > 0 Then
          FundDue$ = DebitAcctNumber$ + DetPad$
        Else
          FundDue$ = DebitAcctNumber$
        End If
        DebitAcctNumber$ = (QPTrim(CDDueAcct$) + FundDue$)
        DebitAcctNumber$ = FmtAcct$(DebitAcctNumber$, GLFundLen%, GLAcctLen%, GLDetLen%)
        ThisAcct = AcctFind(DebitAcctNumber$)
        DebitAcctRecord = ThisAcct
        If ThisAcct > 0 Then
          DebitAcctName$ = GetAcctTitle$(ThisAcct)
        Else
          BadAcct = BadAcct + 1
          DebitAcctName$ = "Undefined"
        End If
        CreditAcctRecord = CDCashRec
        CreditAcctNumber$ = CDCashAcct$
        If Len(QPTrim$(CDCashAcct$)) > 0 Then
          CreditAcctName$ = CDCashAcctName$
        Else
          CreditAcctName$ = "Blank Acct"
        End If
        GoSub PostToGeneralJournal
      End If
      Close TXPGLFile
    End If
 Next T%
 'tran type 3-Release
  For T% = 1 To 51
   If InterfaceMethod <> 1 Then
    If TPrinciple1Pd#(3, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PersCRAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PersCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TPrinciple1Pd#(3, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PersDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PersDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TPrinciple1Pd#(3, T%)
      Dsc$ = "Release"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TPrinciple2Pd#(3, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MTCRAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MTCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TPrinciple2Pd#(3, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MTDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MTDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TPrinciple2Pd#(3, T%)
      Dsc$ = "Release"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TPrinciple3Pd#(3, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MCCRAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MCCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TPrinciple3Pd#(3, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MCDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MCDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TPrinciple3Pd#(3, T%)
      Dsc$ = "Release"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TPrinciple4Pd#(3, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).FECRAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).FECRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TPrinciple4Pd#(3, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).FEDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).FEDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TPrinciple4Pd#(3, T%)
      Dsc$ = "Release"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TPrinciple5Pd#(3, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MHCRAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MHCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TPrinciple5Pd#(3, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).MHDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).MHDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TPrinciple5Pd#(3, T%)
      Dsc$ = "Release"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TRevOpt1Pd#(3, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt1CRAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt1CRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TRevOpt1Pd#(3, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt1DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt1DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TRevOpt1Pd#(3, T%)
      Dsc$ = "Release"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TRevOpt2Pd#(3, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt2CRAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt2CRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TRevOpt2Pd#(3, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt2DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt2DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TRevOpt2Pd#(3, T%)
      Dsc$ = "Release"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TRevOpt3Pd#(3, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt3CRAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt3CRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TRevOpt3Pd#(3, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).Opt3DBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).Opt3DBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TRevOpt3Pd#(3, T%)
      Dsc$ = "Release"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
  'Interest release
    If TInterestPd#(3, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).IntCRAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).IntCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TInterestPd#(3, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).IntDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).IntDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TInterestPd#(3, T%)
      Dsc$ = "Release"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
    If TPenaltyPd#(3, T%) > 0 Then
      TXPGLFile = FreeFile
      Open "TAXPGLBAC.DAT" For Random Shared As TXPGLFile Len = TaxPAcctRecLen
      Get TXPGLFile, 1, TaxPGLAccts(1)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PenCRAcct)
      DebitAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PenCRAcct
      DebitAcctRecord = ThisAcct
      If ThisAcct > 0 Then
        DebitAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        DebitAcctName$ = "Invaild Acct"
      End If
      DebitAmt# = TPenaltyPd#(3, T%)
      ThisAcct = AcctFind(TaxPGLAccts(1).TaxAcct(T%).PenDBAcct)
      CreditAcctRecord = ThisAcct
      CreditAcctNumber$ = TaxPGLAccts(1).TaxAcct(T%).PenDBAcct
      If ThisAcct > 0 Then
        CreditAcctName$ = GetAcctTitle$(ThisAcct)
      Else
        BadAcct = BadAcct + 1
        CreditAcctName$ = "Undefined"
      End If
      CreditAmt# = TPenaltyPd#(3, T%)
      Dsc$ = "Release"
      GoSub PostToGeneralJournal
      Close TXPGLFile
    End If
   End If
  Next T%

TXBunchReturn:
  Return

PostToGeneralJournal:
  'NOTE: Journal Rec 1 is the credit, Rec 2 is the debit
  ReDim GJRec1(1 To 2) As TrEditRecType
  GJRec1(1).AcctRec = CreditAcctRecord
  GJRec1(1).AcctNum = CreditAcctNumber$
  If Len(QPTrim$(CreditAcctNumber$)) > 0 Then
    GJRec1(1).AcctName = CreditAcctName$
  Else
    GJRec1(1).AcctName = "Blank Acct"
  End If
  GJRec1(1).TRDATE = WorkDate
  GJRec1(1).Ref = Ref$
  GJRec1(1).CrAmt = CreditAmt#
  GJRec1(1).EType = "C"
  GJRec1(1).Desc = "FRMTX " + Dsc$
  GJRec1(1).LDesc = "TX Interface"
  GJRec1(1).Src = "TX"
  GJrecnum = LOF(GJFile) \ GJReclen
  Put #GJFile, GJrecnum + 1, GJRec1(1)

  GJRec1(2).AcctRec = DebitAcctRecord
  GJRec1(2).AcctNum = DebitAcctNumber$
  If Len(QPTrim$(DebitAcctNumber$)) > 0 Then
    GJRec1(2).AcctName = DebitAcctName$
  Else
    GJRec1(2).AcctName = "Blank Acct"
  End If
  GJRec1(2).TRDATE = WorkDate
  GJRec1(2).Ref = Ref$
  GJRec1(2).DrAmt = DebitAmt#
  GJRec1(2).EType = "D"
  GJRec1(2).Desc = "FRMTX " + Dsc$
  GJRec1(2).LDesc = "TX Interface"
  GJRec1(2).Src = "TX"
  GJrecnum = LOF(GJFile) \ GJReclen
  Put #GJFile, GJrecnum + 1, GJRec1(2)
Return

Clearouttots:
For ppcnt = 1 To 16
  For RevCnt = 1 To 51
    TPrinciple1Pd#(ppcnt, RevCnt) = 0
    TPrinciple2Pd#(ppcnt, RevCnt) = 0
    TPrinciple3Pd#(ppcnt, RevCnt) = 0
    TPrinciple4Pd#(ppcnt, RevCnt) = 0
    TPrinciple5Pd#(ppcnt, RevCnt) = 0
    TInterestPd#(ppcnt, RevCnt) = 0
    TPenaltyPd#(ppcnt, RevCnt) = 0
    TRevOpt1Pd#(ppcnt, RevCnt) = 0
    TRevOpt2Pd#(ppcnt, RevCnt) = 0
    TRevOpt3Pd#(ppcnt, RevCnt) = 0
    TPrePaidAmt#(ppcnt, RevCnt) = 0
    TPrePaidUsed#(ppcnt, RevCnt) = 0
   Next
  Next
Return
TaxEnd:
  Exit Sub
  Return
End Sub

