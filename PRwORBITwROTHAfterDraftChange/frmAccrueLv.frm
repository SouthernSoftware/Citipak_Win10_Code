VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmAccrueLv 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accrue Leave Benefits"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAccrueLv.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint2 
      Height          =   4620
      Left            =   2160
      TabIndex        =   2
      Top             =   2134
      Width           =   7356
      _Version        =   196609
      _ExtentX        =   12975
      _ExtentY        =   8149
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmAccrueLv.frx":08CA
      Begin EditLib.fpDateTime fptxtStart 
         Height          =   348
         Left            =   3888
         TabIndex        =   0
         Top             =   1824
         Width           =   1836
         _Version        =   196608
         _ExtentX        =   3238
         _ExtentY        =   614
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
         Text            =   "11/20/2002"
         DateCalcMethod  =   0
         DateTimeFormat  =   5
         UserDefinedFormat=   "mm/dd/yyyy"
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
      Begin EditLib.fpDateTime fptxtEnd 
         Height          =   348
         Left            =   3888
         TabIndex        =   1
         Top             =   2496
         Width           =   1836
         _Version        =   196608
         _ExtentX        =   3238
         _ExtentY        =   614
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
         Text            =   "11/20/2002"
         DateCalcMethod  =   0
         DateTimeFormat  =   5
         UserDefinedFormat=   "mm/dd/yyyy"
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
      Begin fpBtnAtlLibCtl.fpBtn cmdAccrue 
         Height          =   690
         Left            =   4080
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Press F10 to process accrual data."
         Top             =   3360
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
         ButtonDesigner  =   "frmAccrueLv.frx":08E6
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
         Height          =   690
         Left            =   1392
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Press ESC to exit this screen."
         Top             =   3360
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
         ButtonDesigner  =   "frmAccrueLv.frx":0AC4
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Accrue Leave Benefits Report"
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
         Height          =   444
         Left            =   1584
         TabIndex        =   5
         Top             =   624
         Width           =   4284
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Last Accrual Date:"
         Height          =   252
         Left            =   1536
         TabIndex        =   4
         Top             =   1968
         Width           =   2028
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Accrue Through:"
         Height          =   348
         Left            =   1536
         TabIndex        =   3
         Top             =   2544
         Width           =   2028
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   828
         Left            =   1344
         Top             =   432
         Width           =   4668
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   4896
      Left            =   1980
      Top             =   1984
      Width           =   7692
   End
End
Attribute VB_Name = "frmAccrueLv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdEscape_Click()
   frmAccrueMenu.Show
   DoEvents
   Unload frmAccrueLv
End Sub

Private Sub cmdAccrue_Click()
   Dim AccrualDate As Integer
   Dim AccrualDateString$
   Dim LastDate As Integer
   
   If CheckValDate(fptxtEnd.Text) = False Then
     MsgBox "Please enter a valid date in the Accrue Through field"
     fptxtEnd.SetFocus
     Exit Sub
   End If
   
   LastDate = Date2Num(fptxtStart.Text) 'last accrual date as shown in the start field
   AccrualDateString$ = QPTrim$(fptxtEnd.Text)
   AccrualDate = Date2Num%(AccrualDateString$) 'date to which this accrual takes place
   
   If LastDate > AccrualDate Then
     MsgBox "The Accrue Through date must be the same as or later than the Last Accrual Date"
     fptxtEnd.SetFocus
     Exit Sub
   End If
   
   If Len(QPTrim$(fptxtStart.Text)) = 0 Then 'only happens the first time accrual
   'occurs...from then on fptxtStart is automatically assigned the last accrual date saved
     MsgBox "Please enter a date from which to begin accrual."
     fptxtStart.SetFocus
     Exit Sub
   End If
   
   If Abs(LastDate - AccrualDate) >= 60 Then
     If MsgBox("The last accrual date is more than 60 days from the new accrual date. If you wish to edit this date then press Yes. Otherwise, press No to continue with the date entered.", vbYesNo) = vbYes Then
       Close
       fptxtEnd.SetFocus
       Exit Sub
     Else
       MainLog "User warned that the last accrual date " + fptxtStart.Text + " is over 60 days away from the new accrual date " + fptxtEnd.Text + " and elected to continue anyway."
     End If
   End If
   
   Call ProcessAccrual(AccrualDate)
   
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
      Call cmdEscape_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%A"
      Call cmdAccrue_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Call LoadAccrueLvScreen
End Sub

Private Sub LoadAccrueLvScreen()
   Dim Today As String * 10
   Dim AccrueHandle As Integer
   Dim AccrualDate As Integer
   Dim AccrualDateString$
   Dim AccrualRec As AccrualDates
   Dim FileHandle As Integer
   Dim One As Integer
   
   If CheckValDate(fptxtEnd.Text) = False Then
     MsgBox "Please enter a valid date in the Accrue Through field"
     fptxtEnd.SetFocus
     Exit Sub
   End If
   
   If Len(QPTrim$(fptxtStart.Text)) = 0 Then
     MsgBox "Please enter a date from which to begin accrual."
     fptxtStart.SetFocus
     Exit Sub
   End If
'   Date$ = FormatDateTime(Date, vbShortDate)
   Today = Date '$
   If Exist("PRDATA\PRACCRUE.DAT") Then 'first time through there will be
   'no prdata\praccrue.dat
     OpenAccrualDatesFile AccrueHandle
     Get AccrueHandle, 1, AccrualRec
     fptxtStart.Text = ReplaceString(MakeRegDate(AccrualRec.PreviousDate), "/", "-")
     
     If Not Exist("prdata\firstAccrual.dat") Then 'this file only exists
     'so that the user can change the start date the first time through and
     'before posting the first time...after the first posting the start date
     'is automatically assigned and prdata\firstAccrual vanishes forever
       fptxtStart.Enabled = False
     End If
     
     If Not Exist(TempAccrualName) Then 'TempAccrualName only exists after accrual
     'runs and before accrual is posted...this allows changes to be made before
     'posting
       fptxtEnd.Text = Today 'today's date is set by this default before accrual runs
     Else
       fptxtEnd.Text = ReplaceString(MakeRegDate(AccrualRec.CurrentDate), "/", "-") 'this occurs
       'if accrual has run but not posted
     End If
     Close AccrueHandle
   Else 'accrual has never been run before at this point
     fptxtStart.Text = ""
     fptxtEnd.Text = Today
     MsgBox "Please enter a date from which you wish to begin accruals."
     'when accrue is run for the first time the start date should
     'be editable until the first post is done...after that start date
     'will never be editable again
     One = 1
     FileHandle = FreeFile
     Open "prdata\firstAccrual.dat" For Output As FileHandle Len = 2 'used
     'only to allow the user to change the start date before it is cast in stone
     'as the beginning start date...from that point start date is entered automatically
     Print #FileHandle, One
     Close FileHandle
   End If
   
   AccrualDateString$ = QPTrim$(fptxtEnd.Text)
   AccrualDate = Date2Num%(AccrualDateString$)
   If AccrualRec.PreviousDate > AccrualDate Then
     MsgBox "The Accrue Through date must be the same as or later than the Last Accrual Date"
'     fptxtEnd.SetFocus
     Exit Sub
   End If
   
End Sub

Sub ProcessAccrual(AccrualDate)

  Dim VAmt#, VADJFlag As Boolean, StableEntry As Long
  Dim SAmt#, EmpName$, TotalSick#, TotalVac#
  Dim SADJFlag As Boolean
  Dim LRecLen As Long
  Dim NumLeaveRec As Long
  Dim LeaveHandle As Integer
  Dim UnitHandle As Integer, x As Long
'  Dim Image$, TImage$, Image1$, Image2$
  Dim TblPos As Integer, YrsPos As Integer, BenPos As Integer
  Dim VacPos As Integer, SickPos As Integer, LineCnt As Integer
'  Dim MaxLines As Integer
  Dim EmpRecSize As Long
  Dim NumOfRecs As Long, IdxRecLen As Integer
  Dim IdxFileSize&, INumOfRecs As Long
  Dim DHandle As Integer, RHandle As Integer
'  Dim RptTitle$, RptName$, RptFile As Integer
  Dim THandle As Integer, AccrualRptFile$
  Dim RecNo As Long
  Dim EmpTotal As Long
  Dim HireDate As Long
  Dim WhatLeaveTbl As Integer
  Dim AccrualDays As Long
  Dim YearsOfService As Integer
  Dim cnt As Long
  Dim VTableEntry As Long
  Dim EmpIdxNNameHandle As Integer
  Dim starS$, yrsEmplyd$, starV$
  Dim TempAccrualHandle As Integer '12/11/02
  Dim TempAccRec As TempAccrualType '12/11/02
  Dim AccrueHandle As Integer
  Dim AccrualRec As AccrualDates
  Dim HTableEntry As Long, HAmt#
  Dim PTableEntry As Long, PAmt#
  
  ReDim Unit(1) As UnitFileRecType
  ReDim EmpRec2(1) As EmpData2Type
'  ReDim TwoPrint(1) As String * 79
  ReDim Tot(1) As String * 8
  ReDim LeaveRec(1) As LeaveRecType
  
  If QPTrim$(fptxtStart.Text) = QPTrim$(fptxtEnd.Text) Then 'inserted to remind
  'the user that accrue has already processed today...if the user has edited
  'any accrual data then they will want to rerun this process to include the
  'change
    If MsgBox("Posting benefits has already occurred today. Do you wish to continue anyway?", vbYesNo) = vbNo Then
      Exit Sub
    End If
  End If
  
  OpenAccrualDatesFile AccrueHandle 'save current info to file
  Get AccrueHandle, 1, AccrualRec
    AccrualRec.CurrentDate = Date2Num(fptxtEnd.Text)
    AccrualRec.PreviousDate = Date2Num(fptxtStart.Text)
  Put AccrueHandle, 1, AccrualRec
  Close AccrueHandle
  
  LRecLen = Len(LeaveRec(1))
  OpenLeaveFileName LeaveHandle
  NumLeaveRec = LOF(LeaveHandle) \ Len(LeaveRec(1))
  If NumLeaveRec = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If
  ReDim LeaveRec(1 To NumLeaveRec) As LeaveRecType
  
  For x = 1 To NumLeaveRec
    Get LeaveHandle, x, LeaveRec(x)
  Next x
  Close LeaveHandle
  
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, Unit(1)
  Close UnitHandle

  EmpRecSize = Len(EmpRec2(1))

  OpenEmpIdxNNameFile EmpIdxNNameHandle
  INumOfRecs = LOF(EmpIdxNNameHandle) \ 2

  ReDim IdxBuf(1 To INumOfRecs) As Integer
  For x = 1 To INumOfRecs 'get employee records in numerical order
     Get EmpIdxNNameHandle, x, IdxBuf(x)
  Next x
  Close EmpIdxNNameHandle
  
  If INumOfRecs = 0 Then
    MsgBox "No records on file."
    Close
    Exit Sub
  End If
  
  OpenTempAccrualFile TempAccrualHandle '12/11/02
  
  OpenEmpData2File DHandle
  NumOfRecs = LOF(DHandle) \ Len(EmpRec2(1))
  
  On Error GoTo ErrorHandler
  
  FrmShowPctComp.Label1 = "Employee Leave Accrual Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdAccrue.Enabled = False '12/11/02
  
  
  For RecNo = 1 To NumOfRecs 'go thru records in numerical order
    Get DHandle, IdxBuf(RecNo), EmpRec2(1)
'    If QPTrim$(EmpRec2(1).EmpLName) = "WHITAKER" Then Stop
    If EmpRec2(1).EMPTDATE = 0 And EmpRec2(1).EMPBCODE > 0 And Not EmpRec2(1).Deleted Then
    'if employee not terminated AND they get benefits AND they weren't deleted.
      EmpTotal = EmpTotal + 1
      HireDate = EmpRec2(1).EMPHDATE
      If HireDate <= -11000 Or HireDate = 0 Then 'roughly 1950
        GoTo BadDateSkip
      End If
      WhatLeaveTbl = EmpRec2(1).LeaveTbl 'get data from
      'leave table assigned to this employee
      If WhatLeaveTbl < 1 Then
        GoSub LeaveTableIsZero
        GoTo BadDateSkip
'        WhatLeaveTbl = 1
      End If
      AccrualDays = AccrualDate - HireDate 'figure employment tenure
      If AccrualDays > 365 Then
        YearsOfService = Int(AccrualDays / 365)
      Else
        YearsOfService = 0
      End If
      For cnt = 1 To 20 '20 lines in leave table
        If YearsOfService <= LeaveRec(WhatLeaveTbl).VEntry(cnt).YEARS Then
          Exit For 'found the line we want so exit for loop
        End If
      Next
      If cnt > 20 Then cnt = 20 'been through the whole table without
      'finding the slot we need so assign be default whatever is on the last line
      If YearsOfService = LeaveRec(WhatLeaveTbl).VEntry(cnt).YEARS Then
        VTableEntry = cnt
      Else 'assign by default the last line
        VTableEntry = cnt - 1
      End If
      If VTableEntry = 0 Then VTableEntry = 1
      VAmt# = OldRound#(LeaveRec(WhatLeaveTbl).VEntry(VTableEntry).EARN * (EmpRec2(1).EMPBCODE * 0.01))
      If VAmt# > 0 Then           ' if there is amount to add
        If EmpRec2(1).EMPVBAL + VAmt# > LeaveRec(WhatLeaveTbl).VacMax Then     'if > max amt
          VAmt# = LeaveRec(WhatLeaveTbl).VacMax - EmpRec2(1).EMPVBAL   'set amt to max
'          VADJFlag = True
        End If                                             '
        EmpRec2(1).EMPVBAL = OldRound#(EmpRec2(1).EMPVBAL + VAmt#)
        TempAccRec.EMPVBAL = EmpRec2(1).EMPVBAL '12/11/02 'assigned to temp file until posting takes place
        EmpRec2(1).EMPVACE = OldRound#(EmpRec2(1).EMPVBAL + EmpRec2(1).EMPVUSED)         'set emp to amt
        TempAccRec.EMPVACE = EmpRec2(1).EMPVACE '12/11/02
      Else
        TempAccRec.EMPVBAL = EmpRec2(1).EMPVBAL '12/11/02 'if we don't assign this here then if the leave table
        'does not have a 99 somewhere to stop the search the employee with max time gets a zero which wipes
        'out the saved tenure
        TempAccRec.EMPVACE = EmpRec2(1).EMPVACE '12/11/02
      End If
      
      'sick leave is treated like vacation leave
      For cnt = 1 To 20
        If YearsOfService <= LeaveRec(WhatLeaveTbl).SEntry(cnt).YEARS Then
          Exit For
        End If
      Next
      If cnt > 20 Then cnt = 20
      If YearsOfService = LeaveRec(WhatLeaveTbl).SEntry(cnt).YEARS Then
        StableEntry = cnt
      Else
        StableEntry = cnt - 1
      End If
      If StableEntry = 0 Then StableEntry = 1
      SAmt# = OldRound#(LeaveRec(WhatLeaveTbl).SEntry(StableEntry).EARN * (EmpRec2(1).EMPBCODE * 0.01)) '8/5
      If SAmt# > 0 Then           ' if there is amount to add
        If EmpRec2(1).EMPSLBAL + SAmt# > LeaveRec(WhatLeaveTbl).SICKMAX Then   'if > max amt
'          SADJFlag = True
          SAmt# = LeaveRec(WhatLeaveTbl).SICKMAX - EmpRec2(1).EMPSLBAL
        End If
        EmpRec2(1).EMPSLBAL = OldRound#(EmpRec2(1).EMPSLBAL + SAmt#)
        TempAccRec.EMPSLBAL = EmpRec2(1).EMPSLBAL '12/11/02
        EmpRec2(1).EMPSLE = OldRound(EmpRec2(1).EMPSLBAL + EmpRec2(1).EMPSLUSE)         'set emp to amt
        TempAccRec.EMPSLE = EmpRec2(1).EMPSLE '12/11/02
      Else
        TempAccRec.EMPSLBAL = EmpRec2(1).EMPSLBAL '12/11/02
        TempAccRec.EMPSLE = EmpRec2(1).EMPSLE '12/11/02
      End If
      
      'holiday leave is treated like vacation and sick leave
      For cnt = 1 To 20
        If YearsOfService <= LeaveRec(WhatLeaveTbl).HEntry(cnt).YEARS Then
          Exit For
        End If
      Next
      If cnt > 20 Then cnt = 20
      If YearsOfService = LeaveRec(WhatLeaveTbl).HEntry(cnt).YEARS Then
        HTableEntry = cnt
      Else
        HTableEntry = cnt - 1
      End If
      If HTableEntry = 0 Then HTableEntry = 1
      HAmt# = OldRound#(LeaveRec(WhatLeaveTbl).HEntry(HTableEntry).EARN * (EmpRec2(1).EMPBCODE * 0.01)) '8/5
      If HAmt# > 0 Then           ' if there is amount to add
        If EmpRec2(1).HOLBAL + HAmt# > LeaveRec(WhatLeaveTbl).HolMax Then     'if > max amt
          HAmt# = LeaveRec(WhatLeaveTbl).HolMax - EmpRec2(1).HOLBAL
        End If
        EmpRec2(1).HOLBAL = OldRound#(EmpRec2(1).HOLBAL + HAmt#)
        TempAccRec.EMPHBAL = EmpRec2(1).HOLBAL '08/26/03
        EmpRec2(1).HOLERN = OldRound(EmpRec2(1).HOLBAL + EmpRec2(1).HolUsed)         'set emp to amt
        TempAccRec.EMPHOLE = EmpRec2(1).HOLERN '08/26/03
      Else
        TempAccRec.EMPHBAL = EmpRec2(1).HOLBAL '08/26/03
        TempAccRec.EMPHOLE = EmpRec2(1).HOLERN '08/26/03
      End If
      
      'Personal leave is treated like vacation, sick and holiday leave
      For cnt = 1 To 20
        If YearsOfService <= LeaveRec(WhatLeaveTbl).PEntry(cnt).YEARS Then
          Exit For
        End If
      Next
      If cnt > 20 Then cnt = 20
      If YearsOfService = LeaveRec(WhatLeaveTbl).PEntry(cnt).YEARS Then
        PTableEntry = cnt
      Else
        PTableEntry = cnt - 1
      End If
      If PTableEntry = 0 Then PTableEntry = 1
      PAmt# = OldRound#(LeaveRec(WhatLeaveTbl).PEntry(PTableEntry).EARN * (EmpRec2(1).EMPBCODE * 0.01)) '8/5
      If PAmt# > 0 Then           ' if there is amount to add
        If EmpRec2(1).PERBAL + PAmt# > LeaveRec(WhatLeaveTbl).PerMax Then     'if > max amt
          PAmt# = LeaveRec(WhatLeaveTbl).PerMax - EmpRec2(1).PERBAL
        End If
        EmpRec2(1).PERBAL = OldRound#(EmpRec2(1).PERBAL + PAmt#)
        TempAccRec.EMPPBAL = EmpRec2(1).PERBAL '08/26/03
        EmpRec2(1).PERERN = OldRound(EmpRec2(1).PERBAL + EmpRec2(1).PerUsed)         'set emp to amt
        TempAccRec.EMPPERE = EmpRec2(1).PERERN '08/26/03
      Else
        TempAccRec.EMPPBAL = EmpRec2(1).PERBAL '08/26/03
        TempAccRec.EMPPERE = EmpRec2(1).PERERN '08/26/03
      End If
      
      'save all to the temporary file until posting occurs
      Put TempAccrualHandle, IdxBuf(RecNo), TempAccRec '08/26/03
    End If

BadDateSkip:
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdEscape.Enabled = True
      Me.cmdAccrue.Enabled = True '12/11/02
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next
  Me.cmdEscape.Enabled = True
  Me.cmdAccrue.Enabled = True '12/11/02
  EnableCloseButton Me.hwnd, True

  Close DHandle
  Close TempAccrualHandle '12/11/02
  Close
  MsgBox "Accrual has been completed."
  frmAccrueMenu.Show
  DoEvents
  Unload frmAccrueLv
  MainLog ("Accrual processing completed but not posted for accrual date of " + fptxtEnd.Text)
  
  Exit Sub

LeaveTableIsZero:
    TempAccRec.EMPVBAL = EmpRec2(1).EMPVBAL
    TempAccRec.EMPVACE = EmpRec2(1).EMPVACE
    TempAccRec.EMPSLBAL = EmpRec2(1).EMPSLBAL
    TempAccRec.EMPSLE = EmpRec2(1).EMPSLE
    TempAccRec.EMPHBAL = EmpRec2(1).HOLBAL
    TempAccRec.EMPHOLE = EmpRec2(1).HOLERN
    TempAccRec.EMPPBAL = EmpRec2(1).PERBAL
    TempAccRec.EMPPERE = EmpRec2(1).PERERN
    Put TempAccrualHandle, IdxBuf(RecNo), TempAccRec

  Return
ErrorHandler:
  Close
  Me.cmdEscape.Enabled = True
  Me.cmdAccrue.Enabled = True '12/11/02
  EnableCloseButton Me.hwnd, True
  Unload FrmShowPctComp
  MsgBox "ERROR: If this problem persists please consult Southern Software."
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdEscape.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmAccrueLv.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub vaImprint1_GotFocus()
  fptxtEnd.SetFocus
End Sub



