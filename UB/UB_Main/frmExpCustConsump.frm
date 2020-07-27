VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmExpCustConsump 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export Customer Consumption"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   2175
   ClientWidth     =   12210
   Icon            =   "frmExpCustConsump.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      Left            =   9840
      TabIndex        =   1
      Top             =   7656
      Width           =   1332
   End
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
      Left            =   8160
      TabIndex        =   0
      Top             =   7656
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "3:59 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "5/14/2007"
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
   Begin EditLib.fpDateTime txtDate2 
      Height          =   348
      Left            =   6108
      TabIndex        =   3
      Top             =   4092
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
   Begin EditLib.fpDateTime txtDate1 
      Height          =   348
      Left            =   6108
      TabIndex        =   4
      Top             =   3528
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Export Customer Consumption"
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
      Left            =   3228
      TabIndex        =   7
      Top             =   1632
      Width           =   5772
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3216
      Top             =   1392
      Width           =   5772
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2028
      Left            =   3288
      Top             =   3048
      Width           =   5652
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Date:"
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
      Height          =   420
      Left            =   4320
      TabIndex        =   6
      Top             =   3576
      Width           =   1668
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Date:"
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
      Height          =   324
      Index           =   0
      Left            =   4416
      TabIndex        =   5
      Top             =   4140
      Width           =   1572
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3216
      Top             =   1272
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
Attribute VB_Name = "frmExpCustConsump"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Private Sub cmdExit_Click()
  frmUBExportMenu.Show
  Unload frmExpCustConsump
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via Expcustconsump by " + PWUser$
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
      cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      cmdOk_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  txtDate1.Text = Format(Now, "mm/dd/yyyy")
  txtDate2.Text = Format(Now, "mm/dd/yyyy")
  Me.HelpContextID = hlpExportConsumption
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub txtDate1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtDate2.SetFocus
  End If
End Sub

Private Sub txtDate2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    cmdOk.SetFocus
  End If
End Sub


Private Function ValidDate()
  Dim TempDate1 As Integer, TempDate2 As Integer
  If CheckValDate(txtDate1) = False And CheckValDate(txtDate2) = False Then
    MsgBox "Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
    ValidDate = False
  Else
    TempDate1 = DateDiff("d", "12/31/1979", txtDate1)
    TempDate2 = DateDiff("d", "12/31/1979", txtDate2)
    If TempDate1 > TempDate2 Then
      ValidDate = False
      MsgBox "The Starting And Ending Dates Must Be In Chronological Order Or Equal", vbOKOnly, "Invalid Date"
    Else
      ValidDate = True
    End If
  End If
End Function

Private Sub cmdOk_Click()
'
  If ValidDate Then
    DeActivateControls Me, True
'do consumpstuff here
      ExportConsumptionInformation
    'for Johnston Co only
    ''  ExportConsumptionInformationJohnst
    ActivateControls Me, True
  End If
End Sub

Private Sub ExportConsumptionInformation()
  Dim Dash80 As String, IndexName As String, IdxRecLen As Integer
  Dim UBCustRecLen As Integer, UBCust As Integer, UBTran As Integer
  Dim IdxFileSize As Long, IdxNumOfRecs As Long, cnt As Long
  Dim UBTranRecLen As Integer, NumOfRecs As Long, NumOfCust As Long
  Dim Handle As Integer, UsingBook As Boolean, NumOfPeriods As Integer
  Dim RecNo As Long, DidCnt As Long, ThisTrans As Long, FMonth As Integer
  Dim FYear As Integer, TYear As Integer, TMonth As Integer, BadCount As Integer
  Dim FMCnt As Integer, DidAMeter As Boolean, MtrCnt As Integer, MTRMulti#
  Dim MeterType As String, MeterConsp As Long, MaxMeterAmt As Long
  Dim TotalConsump As Long, QPos As Integer, LocationNumber As String
  Dim Zip As String, CCCnt As Long, MoFlag As Boolean, UBSetupreclen As Integer
'  Dim CustomerRecord  As Integer, MCnt As Integer, GTMeterConsp As Double
'  Dim Multi As Long, Cubic As Boolean, ChkMtr As Boolean
'  Dim MTRType As String, MType As String, TMeterConsp As Double
  Dim q As String, C As String, FromDate As Integer, ThruDate As Integer
  Dim UBRpt As String, zz As String, zzN As Integer, CCnt As Long
  FrmShowPctComp.Label1 = "Creating Export Files"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  q$ = Chr$(34)
  C$ = ","
  FromDate = Date2Num%(txtDate1)
  ThruDate = Date2Num%(txtDate2)
  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUp(1))
  BadCount = 0
  LoadUBSetUpFile UBSetUp(), UBSetupreclen

  If InStr(TOWNNAME$, "JOHNSTON") > 0 Then
    MoFlag = True
  Else
    MoFlag = False
  End If

  IndexName$ = BookIndexFile
  UsingBook = True

  Dash80$ = String$(80, "-")

  ReDim UBTranRec(1) As UBTransRecType
  UBTranRecLen = Len(UBTranRec(1))

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  IdxRecLen = 4               'we are using a long integer
  IdxFileSize& = FileSize(IndexName$)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
  'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
  NumOfRecs = IdxNumOfRecs

  Handle = FreeFile
  Open IndexName$ For Random Shared As Handle Len = IdxRecLen
  For cnt = 1 To IdxNumOfRecs
    Get #Handle, cnt, IdxBuff(cnt)
  Next
  Close Handle

  GoSub ExpCheckDate
  UBRpt = FreeFile
  Open UBPath$ + "UBCONSMP.TXT" For Output As UBRpt
  Print #UBRpt, q$; "ACCT"; q$; C$; q$; "LOCATION"; q$; C$; q$; "CUSTNAME";
  Print #UBRpt, q$; C$; q$; "CUSTTYPE"; q$; C$; q$; "ADDR1"; q$; C$; q$; "ADDR2"; q$; C$; q$; "CITY"; q$; C$; q$; "STATE"; q$; C$; q$; "ZIP"; q$; C$; q$; "SERVADDR";
  For zzN = NumOfPeriods To 1 Step -1
    zz$ = QPTrim$(Str$(zzN))
    Print #UBRpt, q$; C$; q$; "TRDATE"; zz$; q$; C$; q$; "CURRREAD"; zz$; q$; C$; q$; "PREVREAD"; zz$; q$; C$; q$; "CONSUMP"; zz$; q$; C$; q$; "TRANAMT"; zz$;
  Next
  Print #UBRpt, q$;
  UBTran = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTran Len = UBTranRecLen

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  NumOfCust = LOF(UBCust) \ UBCustRecLen
  NumOfPeriods = 0
  For CCnt = 1 To NumOfCust
    FrmShowPctComp.ShowPctComp CCnt, NumOfCust&
    If FrmShowPctComp.Out Then
      Close
      Unload FrmShowPctComp
      GoTo ExitConsumpHist
    End If

    RecNo& = CCnt    'IdxBuff(CCnt).RecNum
    Get #UBCust, RecNo&, UBCustRec(1)
    If UBCustRec(1).DelFlag = 0 Then
      DidCnt = 0
      ThisTrans& = UBCustRec(1).LastTrans
      Do While ThisTrans& > 0
        Get #UBTran, ThisTrans&, UBTranRec(1)
        If MoFlag Then
          If UBTranRec(1).TransDate < FromDate Then
            BadCount = BadCount + 1
            If BadCount > 3 Then
              Exit Do
            End If
          End If
        End If

        If UBTranRec(1).TransType = TranUtilityBill Then
          If UBTranRec(1).TransDate >= FromDate And UBTranRec(1).TransDate <= ThruDate Then
            If DidCnt = 0 Then
              GoSub PrintCustInfo
            End If
            GoSub PrintConsDetail
            DidCnt = DidCnt + 1
            If DidCnt = NumOfPeriods Then
              Exit Do
            End If
          End If
        End If
        ThisTrans& = UBTranRec(1).PrevTrans
      Loop
      If DidCnt > 0 Then
        If DidCnt < NumOfPeriods Then
          For zzN = DidCnt + 1 To NumOfPeriods
             Print #UBRpt, C$; q$; "01-01-1980"; q$; C$; q$; "0"; q$; C$; q$; "0"; q$; C$; q$; "0"; q$; C$; q$; "0.00"; q$;
          Next
        End If
      End If
    End If
    'ShowPctComp CCnt, NumOfRecs
    'IF CCnt > 149 THEN EXIT FOR
'    If ExitFlag Then
'      Exit For
'    End If
  Next

  Close
  
  MsgBox "File Created is UBCONSMP.TXT", vbOKOnly, "File created"

   


ExitConsumpHist:

Exit Sub

ExpCheckDate:

    FYear = Val(Right$(txtDate1, 4))
    TYear = Val(Right$(txtDate2, 4))
    FMonth = Val(Left$(txtDate1, 2))
    TMonth = Val(Left$(txtDate2, 2))
    If FYear = TYear Then
      If FMonth = TMonth Then
        NumOfPeriods = 1
        
      Else
        NumOfPeriods = (TMonth - FMonth) + 1
        
      End If
    Else
      FMCnt = (12 - FMonth) + 1
      NumOfPeriods = FMCnt + TMonth
      If TYear - FYear > 1 Then
        NumOfPeriods = NumOfPeriods + 12
      End If
      
    End If
 
ExpDateRet:
Return

PrintConsDetail:
  MTRMulti# = 1
  DidAMeter = False
  For MtrCnt = 1 To 7
    If UBTranRec(1).MtrTypes(MtrCnt) > 0 Then
      DidAMeter = True
      MTRMulti# = UBCustRec(1).LocMeters(MtrCnt).MTRMulti
      If MTRMulti# <= 0 Then
          MTRMulti# = 1
      End If

      GoSub PrintThisMeter
    End If
  Next
  If Not DidAMeter Then
    MeterType$ = "        "
    MtrCnt = 1
      MTRMulti# = UBCustRec(1).LocMeters(MtrCnt).MTRMulti
      If MTRMulti# <= 0 Then
          MTRMulti# = 1
      End If
    GoSub PrintThisMeter
  End If
Return

PrintThisMeter:
  Print #UBRpt, C$; q$; Num2Date(UBTranRec(1).TransDate); q$; C$; q$;
    If UBTranRec(1).CurRead(MtrCnt) < 0 Then
      UBTranRec(1).CurRead(MtrCnt) = 0
    End If
    If UBTranRec(1).PrevRead(MtrCnt) < 0 Then
      UBTranRec(1).PrevRead(MtrCnt) = 0
    End If
  Print #UBRpt, QPTrim$(Str$(UBTranRec(1).CurRead(MtrCnt))); q$; C$; q$;
  Print #UBRpt, QPTrim$(Str$(UBTranRec(1).PrevRead(MtrCnt))); q$; C$; q$;
  MeterConsp& = UBTranRec(1).CurRead(MtrCnt) - UBTranRec(1).PrevRead(MtrCnt)
  If MeterConsp& < 0 Then
    MaxMeterAmt& = 10& ^ (Len(Str$(UBTranRec(1).PrevRead(MtrCnt))) - 1)
    MeterConsp& = (MaxMeterAmt& - UBTranRec(1).PrevRead(MtrCnt)) + UBTranRec(1).CurRead(MtrCnt)
  End If
  MeterConsp& = MeterConsp& * MTRMulti#
  Print #UBRpt, QPTrim$(Str$(MeterConsp&)); q$; C$; q$;
  Print #UBRpt, QPTrim$(Using$("######.##", UBTranRec(1).Transamt)); q$;

  TotalConsump& = TotalConsump& + MeterConsp&

Return

PrintCustInfo:
  'IF CCCnt > 0 THEN
    Print #UBRpt,
  'END IF

  Do
    QPos = InStr(UBCustRec(1).CustName, q$)
    If QPos > 0 Then
      Mid$(UBCustRec(1).CustName, QPos, 1) = " "
    End If
  Loop While QPos > 0
'  PRINT #UBRpt, q$; QPTrim$(STR$(RecNo&)); q$; c$; q$; QPTrim$(UBCustRec(1).CustName); q$;
'  PRINT #UBRpt, q$; QPTrim$(STR$(RecNo&)); q$; c$; q$; QPTrim$(UBCustRec(1).CustName); q$;
  LocationNumber$ = QPTrim$(UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB)
  Zip$ = QPTrim$(UBCustRec(1).ZIPCODE)
  If Len(Zip$) > 5 Then
    Zip$ = Left$(Zip$, 5) + "-" + Mid$(Zip$, 6)
  End If

  Print #UBRpt, q$; QPTrim$(Str$(RecNo&)); q$; C$; q$; LocationNumber$; q$; C$;
  Print #UBRpt, q$; QPTrim$(UBCustRec(1).CustName); q$; C$; q$; QPTrim$(UBCustRec(1).CUSTTYPE); q$; C$;
  Print #UBRpt, q$; QPTrim$(UBCustRec(1).ADDR1); q$; C$; q$; QPTrim$(UBCustRec(1).ADDR2); q$; C$;
  Print #UBRpt, q$; QPTrim$(UBCustRec(1).CITY); q$; C$; q$; QPTrim$(UBCustRec(1).STATE); q$; C$;
  Print #UBRpt, q$; Zip$; q$; C$; q$; QPTrim$(UBCustRec(1).ServAddr); q$;
  'PRINT #UBRpt, q$; QPTrim$(UBCustRec(1).HPhone); q$

  'PRINT #UBRpt, UBCustRec(1).CustName
  CCCnt = CCCnt + 1
  'IF CCCnt > 99 THEN
  '  ExitFlag = True
  'END IF
Return
End Sub

Private Sub ExportConsumptionInformationJohnst()
  Dim Dash80 As String, IndexName As String, IdxRecLen As Integer
  Dim UBCustRecLen As Integer, UBCust As Integer, UBTran As Integer
  Dim IdxFileSize As Long, IdxNumOfRecs As Long, cnt As Long
  Dim UBTranRecLen As Integer, NumOfRecs As Long, NumOfCust As Long
  Dim Handle As Integer, UsingBook As Boolean, NumOfPeriods As Integer
  Dim RecNo As Long, DidCnt As Long, ThisTrans As Long, FMonth As Integer
  Dim FYear As Integer, TYear As Integer, TMonth As Integer
  Dim FMCnt As Integer, DidAMeter As Boolean, MtrCnt As Integer
  Dim MeterType As String, MeterConsp As Long, MaxMeterAmt As Long
  Dim TotalConsump As Long, QPos As Integer, LocationNumber As String
  Dim Zip As String, CCCnt As Long
'  Dim CustomerRecord  As Integer, MCnt As Integer, GTMeterConsp As Double
'  Dim Multi As Long, Cubic As Boolean, ChkMtr As Boolean
'  Dim MTRType As String, MType As String, TMeterConsp As Double
  Dim q As String, C As String, FromDate As Integer, ThruDate As Integer
  Dim UBRpt As String, zz As String, zzN As Integer, CCnt As Long
  FrmShowPctComp.Label1 = "Creating Export Files"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
'  q$ = Chr$(34)
'  C$ = ","
  'need no quotes and change delimeter to | for johnston special
  q$ = "" 'Chr$(34)
  C$ = "|"  '","
  FromDate = Date2Num%(txtDate1)
  ThruDate = Date2Num%(txtDate2)

  IndexName$ = BookIndexFile
  UsingBook = True
  Dash80$ = String$(80, "-")
  ReDim UBTranRec(1) As UBTransRecType
  UBTranRecLen = Len(UBTranRec(1))

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  IdxRecLen = 4               'we are using a long integer
  IdxFileSize& = FileSize(IndexName$)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
  'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
  NumOfRecs = IdxNumOfRecs

  Handle = FreeFile
  Open IndexName$ For Random Shared As Handle Len = IdxRecLen
  For cnt = 1 To IdxNumOfRecs
    Get #Handle, cnt, IdxBuff(cnt)
  Next
  Close Handle


  UBRpt = FreeFile
  Open UBPath$ + "UBCONSMP.TXT" For Output As UBRpt
  Print #UBRpt, q$; "ACCT"; q$; C$; q$; "LOCATION"; q$; C$; q$; "CUSTNAME";
  Print #UBRpt, q$; C$; q$; "CUSTTYPE"; q$; C$; q$; "ADDR1"; q$; C$; q$; "ADDR2"; q$; C$; q$; "CITY"; q$; C$; q$; "STATE"; q$; C$; q$; "ZIP"; q$; C$; q$; "SERVADDR";
  For zzN = NumOfPeriods To 1 Step -1
    zz$ = QPTrim$(Str$(zz))
    Print #UBRpt, q$; C$; q$; "TRDATE"; zz$; q$; C$; q$; "CURRREAD"; zz$; q$; C$; q$; "PREVREAD"; zz$; q$; C$; q$; "CONSUMP"; zz$; q$; C$; q$; "TRANAMT"; zz$;
  Next
  Print #UBRpt, q$;
  UBTran = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTran Len = UBTranRecLen

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  NumOfCust = LOF(UBCust) \ UBCustRecLen

  For CCnt = 1 To NumOfCust
    FrmShowPctComp.ShowPctComp CCnt, NumOfCust&
    If FrmShowPctComp.Out Then
      Close
      Unload FrmShowPctComp
      GoTo ExitConsumpHist
    End If

    RecNo& = CCnt    'IdxBuff(CCnt).RecNum
    Get #UBCust, RecNo&, UBCustRec(1)
    If UBCustRec(1).DelFlag = 0 Then
      DidCnt = 0
      ThisTrans& = UBCustRec(1).LastTrans
      Do While ThisTrans& > 0
        Get #UBTran, ThisTrans&, UBTranRec(1)
        If UBTranRec(1).TransType = TranUtilityBill Then
          If UBTranRec(1).TransDate >= FromDate And UBTranRec(1).TransDate <= ThruDate Then
            If DidCnt = 0 Then
              GoSub PrintCustInfo
            End If
            GoSub PrintConsDetail
            DidCnt = DidCnt + 1
            If DidCnt = NumOfPeriods Then
              Exit Do
            End If
          End If
        End If
        ThisTrans& = UBTranRec(1).PrevTrans
      Loop
      If DidCnt > 0 Then
        If DidCnt < NumOfPeriods Then
          For zzN = DidCnt + 1 To NumOfPeriods
             Print #UBRpt, C$; q$; "01-01-1980"; q$; C$; q$; "0"; q$; C$; q$; "0"; q$; C$; q$; "0"; q$; C$; q$; "0.00"; q$;
          Next
        End If
      End If
    End If
    'ShowPctComp CCnt, NumOfRecs
    'IF CCnt > 149 THEN EXIT FOR
'    If ExitFlag Then
'      Exit For
'    End If
  Next

  Close
  
  MsgBox "File Created is UBCONSMP.TXT", vbOKOnly, "File created"



ExitConsumpHist:

Exit Sub

ExpCheckDate:

    FYear = Val(Right$(FromDate, 4))
    TYear = Val(Right$(ThruDate, 4))
    FMonth = Val(Left$(FromDate, 2))
    TMonth = Val(Left$(ThruDate, 2))
    If FYear = TYear Then
      If FMonth = TMonth Then
        NumOfPeriods = 1
        
      Else
        NumOfPeriods = (TMonth - FMonth) + 1
        
      End If
    Else
      FMCnt = (12 - FMonth) + 1
      NumOfPeriods = FMCnt + TMonth
      If TYear - FYear > 1 Then
        NumOfPeriods = NumOfPeriods + 12
      End If
      
    End If
 
ExpDateRet:
Return

PrintConsDetail:
  DidAMeter = False
  For MtrCnt = 1 To 7
    If UBTranRec(1).MtrTypes(MtrCnt) > 0 Then
      DidAMeter = True
      GoSub PrintThisMeter
    End If
  Next
  If Not DidAMeter Then
    MeterType$ = "        "
    MtrCnt = 1
    GoSub PrintThisMeter
  End If
Return

PrintThisMeter:
  Print #UBRpt, C$; q$; Num2Date(UBTranRec(1).TransDate); q$; C$; q$;
    If UBTranRec(1).CurRead(MtrCnt) < 0 Then
      UBTranRec(1).CurRead(MtrCnt) = 0
    End If
    If UBTranRec(1).PrevRead(MtrCnt) < 0 Then
      UBTranRec(1).PrevRead(MtrCnt) = 0
    End If
  Print #UBRpt, QPTrim$(Str$(UBTranRec(1).CurRead(MtrCnt))); q$; C$; q$;
  Print #UBRpt, QPTrim$(Str$(UBTranRec(1).PrevRead(MtrCnt))); q$; C$; q$;
  MeterConsp& = UBTranRec(1).CurRead(MtrCnt) - UBTranRec(1).PrevRead(MtrCnt)
  If MeterConsp& < 0 Then
    MaxMeterAmt& = 10& ^ (Len(Str$(UBTranRec(1).PrevRead(MtrCnt))) - 1)
    MeterConsp& = (MaxMeterAmt& - UBTranRec(1).PrevRead(MtrCnt)) + UBTranRec(1).CurRead(MtrCnt)
  End If
  Print #UBRpt, QPTrim$(Str$(MeterConsp&)); q$; C$; q$;
  Print #UBRpt, QPTrim$(Using$("######.##", UBTranRec(1).Transamt)); q$;

  TotalConsump& = TotalConsump& + MeterConsp&

Return

PrintCustInfo:
  'IF CCCnt > 0 THEN
    Print #UBRpt,
  'END IF
  
'For Johnston Co rem out do loop
'  Do
'    QPos = InStr(UBCustRec(1).CustName, q$)
'    If QPos > 0 Then
'      Mid$(UBCustRec(1).CustName, QPos, 1) = " "
'    End If
'  Loop While QPos > 0
'
'  PRINT #UBRpt, q$; QPTrim$(STR$(RecNo&)); q$; c$; q$; QPTrim$(UBCustRec(1).CustName); q$;
'  PRINT #UBRpt, q$; QPTrim$(STR$(RecNo&)); q$; c$; q$; QPTrim$(UBCustRec(1).CustName); q$;
  LocationNumber$ = QPTrim$(UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB)
  Zip$ = QPTrim$(UBCustRec(1).ZIPCODE)
  If Len(Zip$) > 5 Then
    Zip$ = Left$(Zip$, 5) + "-" + Mid$(Zip$, 6)
  End If

  Print #UBRpt, q$; QPTrim$(Str$(RecNo&)); q$; C$; q$; LocationNumber$; q$; C$;
  Print #UBRpt, q$; QPTrim$(UBCustRec(1).CustName); q$; C$; q$; QPTrim$(UBCustRec(1).CUSTTYPE); q$; C$;
  Print #UBRpt, q$; QPTrim$(UBCustRec(1).ADDR1); q$; C$; q$; QPTrim$(UBCustRec(1).ADDR2); q$; C$;
  Print #UBRpt, q$; QPTrim$(UBCustRec(1).CITY); q$; C$; q$; QPTrim$(UBCustRec(1).STATE); q$; C$;
  Print #UBRpt, q$; Zip$; q$; C$; q$; QPTrim$(UBCustRec(1).ServAddr); q$;
  'PRINT #UBRpt, q$; QPTrim$(UBCustRec(1).HPhone); q$

  'PRINT #UBRpt, UBCustRec(1).CustName
  CCCnt = CCCnt + 1
  'IF CCCnt > 99 THEN
  '  ExitFlag = True
  'END IF
Return
End Sub

