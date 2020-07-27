VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmPRRelinkTransHist 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relink Transaction History"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmPRRelinkTransHist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   3612
      Left            =   2160
      TabIndex        =   0
      Top             =   2610
      Width           =   7356
      _Version        =   196609
      _ExtentX        =   12975
      _ExtentY        =   6371
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmPRRelinkTransHist.frx":08CA
      Begin EditLib.fpDateTime fptxtYear 
         Height          =   372
         Left            =   4272
         TabIndex        =   1
         Top             =   1584
         Width           =   1164
         _Version        =   196608
         _ExtentX        =   2053
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
         ButtonStyle     =   2
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
         Text            =   "2018"
         DateCalcMethod  =   0
         DateTimeFormat  =   5
         UserDefinedFormat=   "yyyy"
         DateMax         =   "00000000"
         DateMin         =   "00000000"
         TimeMax         =   "000000"
         TimeMin         =   "000000"
         TimeString1159  =   ""
         TimeString2359  =   ""
         DateDefault     =   "00000000"
         TimeDefault     =   "000000"
         TimeStyle       =   0
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
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
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   690
         Left            =   4176
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Press F10 to begin the procedure to recreate each employee's trasaction thread."
         Top             =   2352
         Width           =   1890
         _Version        =   131072
         _ExtentX        =   3334
         _ExtentY        =   1217
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
         ButtonDesigner  =   "frmPRRelinkTransHist.frx":08E6
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
         Height          =   690
         Left            =   1296
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Press ESC to exit this screen."
         Top             =   2352
         Width           =   1890
         _Version        =   131072
         _ExtentX        =   3334
         _ExtentY        =   1217
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
         ButtonDesigner  =   "frmPRRelinkTransHist.frx":0AC4
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "ReLink Transaction History"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   492
         Left            =   1584
         TabIndex        =   3
         Top             =   672
         Width           =   4284
      End
      Begin VB.Label Label2 
         BackColor       =   &H00D0D0D0&
         Caption         =   "Year to Relink:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   2
         Top             =   1655
         Width           =   1620
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   684
         Left            =   1488
         Top             =   528
         Width           =   4476
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   3936
      Left            =   1980
      Top             =   2466
      Width           =   7692
   End
End
Attribute VB_Name = "frmPRRelinkTransHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdEscape_Click()
   frmPayrollMainMenu.Show
   DoEvents
   Unload frmPRRelinkTransHist
End Sub

Private Sub cmdProcess_Click()
  Dim Year As Integer
  Dim LowDate As Integer
  Dim HiDate As Integer
  Dim RelinkYear$
  Dim Emp2RecLen As Integer
  Dim Emp3RecLen As Integer
  Dim TranRecLen As Integer
  Dim ENumOfRec As Integer
  Dim TNumOfRec As Long
  Dim NewRecCnt As Long
  Dim Emp1Rec As EmpData1Type
  Dim Emp2Rec As EmpData2Type
  Dim Emp3Rec As EmpData3Type
  Dim Emp3RecB As EmpData3Type
  Dim TranRec As TransRecType
  Dim cnt As Long
  Dim ECnt As Integer, E3Cnt As Integer
  Dim TotalTransRecs As Long
  Dim TCnt As Long
  Dim FirstEmpHRec As Long
  
  KillFile OldHistFileName
  RelinkYear = Mid(fptxtYear.Text, 3, 2)
  Year = Val(fptxtYear.Text)

  Select Case Year
  Case Is < 2000
    LowDate = Date2Num("01-01-19" + RelinkYear$)
    HiDate = Date2Num("12-31-19" + RelinkYear$)
  Case Else 'it greater or equal
    LowDate = Date2Num("01-01-20" + RelinkYear$)
    HiDate = Date2Num("12-31-20" + RelinkYear$)
  End Select
  
  ReDim TPntr(0 To 800) As Integer
  
  Emp2RecLen = Len(Emp2Rec)
  Emp3RecLen = Len(Emp3Rec)
  
  TranRecLen = Len(TranRec)
  
  ENumOfRec = FileSize("PRData\" + EmpData2Name) \ Emp2RecLen
  TNumOfRec = FileSize("PRData\" + TransHistFileName) \ TranRecLen
  
  NewRecCnt = 1
  
  TNumOfRec = FileSize&("PRData\" + TransHistFileName) \ TranRecLen

  Open "PRData\" + TransHistFileName For Random As #1 Len = TranRecLen
  TNumOfRec = LOF(1) \ TranRecLen

  ReDim TPins(1 To TNumOfRec) As Integer
  
  FrmShowPctComp.Label1 = "Stand By"
  FrmShowPctComp.Show , Me
  FrmShowPctComp.cmdCancel.Visible = False
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  For cnt = 1 To TNumOfRec
    Get #1, cnt, TranRec
    TPins(cnt) = TranRec.EmpPin
  Next
  Close

  Name "PRData\" + TransHistFileName As OldHistFileName
  
  Open "PRData\" + EmpData2Name For Random As #1 Len = Emp2RecLen
  Open "PRData\" + EmpData3Name For Random As #4 Len = Emp3RecLen
  Open OldHistFileName For Random As #2 Len = TranRecLen
  Open "PRData\" + TransHistFileName For Random As #3 Len = TranRecLen
  FrmShowPctComp.Label1 = "Relinking Transaction History"
  
  For ECnt = 1 To ENumOfRec
    Get #1, ECnt, Emp2Rec
    GoSub GetTransRecNums
    If TPntr(0) Then
      GoSub RebuildTransHistory
    Else
      Emp2Rec.LastTransRec = 0
    End If
    Put #1, ECnt, Emp2Rec
    Put #4, ECnt, Emp3Rec
    
    Emp3Rec = Emp3RecB
    FrmShowPctComp.ShowPctComp ECnt, ENumOfRec
  Next
  Close
  
  EnableCloseButton Me.hwnd, True
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  
  KillFile OldHistFileName
  MainLog ("Relinking for " + CStr(Year) + " completed successfully.")
  MsgBox "Relinking completed successfully."
  Call cmdEscape_Click
  Exit Sub
  
GetTransRecNums:
  ReDim TPntr(0 To 20000) As Integer
  TotalTransRecs = 0
  For TCnt = 1 To TNumOfRec
    If TPins(TCnt) = Emp2Rec.EmpPin Then
      TotalTransRecs = TotalTransRecs + 1
      TPntr(TotalTransRecs) = TCnt
    End If
    TPntr(0) = TotalTransRecs
  Next
  Return
  
RebuildTransHistory:
  FirstEmpHRec = NewRecCnt
  
  For cnt = 1 To TPntr(0)
    Get #2, TPntr(cnt), TranRec
    If cnt = 1 Then
      TranRec.PrevTransRec = 0
    Else
      TranRec.PrevTransRec = NewRecCnt - 1
    End If
    Put #3, NewRecCnt, TranRec
    NewRecCnt = NewRecCnt + 1
    Emp2Rec.LastTransRec = NewRecCnt - 1
    Select Case TranRec.CheckDate
    Case LowDate To HiDate
      GoSub SumEmpYTD
    End Select
  Next
  Return
  
SumEmpYTD:
  ''** Update employee 3 file
  ''-=-=man
  Emp3Rec.YTDGrossPay = UtilRound(Emp3Rec.YTDGrossPay + TranRec.GrossPay)
  Emp3Rec.YTDFedGrossPay = UtilRound(Emp3Rec.YTDFedGrossPay + TranRec.FedGrossPay)
  Emp3Rec.YTDStaGrossPay = UtilRound(Emp3Rec.YTDStaGrossPay + TranRec.StaGrossPay)
  Emp3Rec.YTDSocGrossPay = UtilRound(Emp3Rec.YTDSocGrossPay + TranRec.SocGrossPay)
  Emp3Rec.YTDMedGrossPay = UtilRound(Emp3Rec.YTDMedGrossPay + TranRec.MedGrossPay)
  
  Emp3Rec.YTDRegPay = UtilRound(Emp3Rec.YTDRegPay + TranRec.TotRegWage)
  Emp3Rec.YTDOTPay = UtilRound(Emp3Rec.YTDOTPay + TranRec.TotOTWage)
  
  Emp3Rec.YTDNet = UtilRound(Emp3Rec.YTDNet + TranRec.NetPay)
  
  Emp3Rec.YTDFederal = UtilRound(Emp3Rec.YTDFederal + TranRec.FedTaxAmt)
  Emp3Rec.YTDState = UtilRound(Emp3Rec.YTDState + TranRec.StaTaxAmt)
  Emp3Rec.YTDSocial = UtilRound(Emp3Rec.YTDSocial + TranRec.SocTaxAmt)
  Emp3Rec.YTDMedicare = UtilRound(Emp3Rec.YTDMedicare + TranRec.MedTaxAmt)
  Emp3Rec.YTDRetire = UtilRound(Emp3Rec.YTDRetire + TranRec.RetireAmt)
  
  'year to date totals on deductions
  For E3Cnt = 1 To 50 ' changed from 12 to 50 on 1/31/2005
    Emp3Rec.YTDDAmt(E3Cnt) = UtilRound(Emp3Rec.YTDDAmt(E3Cnt) + TranRec.DAmt(E3Cnt))
    Emp3Rec.YTDDAmtT = UtilRound(Emp3Rec.YTDDAmtT + TranRec.DAmt(E3Cnt))
  Next
  
  'year to date totals on alt earnings
  Emp3Rec.YTDEarn1 = UtilRound(Emp3Rec.YTDEarn1 + TranRec.EAmt(1))
  Emp3Rec.YTDEarn2 = UtilRound(Emp3Rec.YTDEarn2 + TranRec.EAmt(2))
  Emp3Rec.YTDEarn3 = UtilRound(Emp3Rec.YTDEarn3 + TranRec.EAmt(3))
  Emp3Rec.YTDEarnT = UtilRound(Emp3Rec.YTDEarn1 + Emp3Rec.YTDEarn2 + Emp3Rec.YTDEarn3)
  
  '    EmpRec2(1).EMPVACE = UtilRound(EmpRec2(1).EMPVBAL + EmpRec2(1).EMPVUSED)
  '    EmpRec2(1).EMPSLE = UtilRound(EmpRec2(1).EMPSLBAL + EmpRec2(1).EMPSLUSE)
  '    EmpRec2(1).EMPCTE = UtilRound(EmpRec2(1).EMPCTBAL + EmpRec2(1).EMPCTUSE)
  
  Return
  
End Sub

Function UtilRound#(DblNum#)
  UtilRound# = (Int((DblNum# * 100) + 0.5) / 100)
End Function

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
      Call cmdEscape_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%R"
      Call cmdProcess_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub LoadMe()
  fptxtYear.Text = Mid(Date, 7, 4)
'  Label3.BackStyle = 0
End Sub
