VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Begin VB.Form frmTCMatchUp1 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Match Up"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
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
   Icon            =   "frmTCMatchUp1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox fptxtDelimiter 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2250
      MaxLength       =   1
      TabIndex        =   0
      Top             =   8088
      Width           =   852
   End
   Begin FPSpread.vaSpread vaSpread 
      Height          =   6852
      Left            =   276
      TabIndex        =   1
      Top             =   888
      Width           =   11052
      _Version        =   196613
      _ExtentX        =   19494
      _ExtentY        =   12086
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   4
      MaxRows         =   100
      SpreadDesigner  =   "frmTCMatchUp1.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   645
      Left            =   7170
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "Press the 'Cancel' button to exit this screen and return to the main 'Business License Reports' menu."
      Top             =   7965
      Width           =   1980
      _Version        =   131072
      _ExtentX        =   3492
      _ExtentY        =   1138
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
      ButtonDesigner  =   "frmTCMatchUp1.frx":1075
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   645
      Left            =   9330
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   $"frmTCMatchUp1.frx":1253
      Top             =   7965
      Width           =   1980
      _Version        =   131072
      _ExtentX        =   3492
      _ExtentY        =   1138
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
      ButtonDesigner  =   "frmTCMatchUp1.frx":12FE
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdLoadCoData 
      Height          =   645
      Left            =   4170
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7965
      Width           =   2580
      _Version        =   131072
      _ExtentX        =   4551
      _ExtentY        =   1138
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
      ButtonDesigner  =   "frmTCMatchUp1.frx":14DD
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Delimiter:"
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   336
      TabIndex        =   6
      Top             =   8208
      Width           =   1812
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmTCMatchUp1.frx":16C4
      ForeColor       =   &H8000000E&
      Height          =   612
      Left            =   456
      TabIndex        =   5
      Top             =   168
      Width           =   10932
   End
End
Attribute VB_Name = "frmTCMatchUp1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Public RowCnt As Integer
  Public dlm As String

Private Sub cmdExit_Click()
  frmTCMainMenu1.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdLoadCoData_Click()
  If QPTrim$(fptxtDelimiter.Text) = "" Then
    Call TCMsg(900, "Please enter a delimiter.")
    fptxtDelimiter.SetFocus
    Exit Sub
  End If
  
  Call LoadMe
  cmdProcess.Visible = True
End Sub

Private Sub cmdProcess_Click()
  Dim TempHandle As Integer
  Dim TempRec As TempConversionData
  Dim NumOfTempRecs As Long
  Dim x As Integer, y As Integer
  Dim ColCnt As Integer
  Dim ThisCol As Integer
  Dim ThisPos As Integer
  Dim Textline$
  Dim ThisFile$
  Dim LHandle As Integer
  Dim TextLen As Integer
  Dim Thisch As String
  Dim ThisWord$
  Dim FirstLine As Boolean
  Dim RecCnt As Long
  
  If QPTrim$(fptxtDelimiter.Text) = "" Then
    Call TCMsg(900, "Please enter a delimiter.")
    fptxtDelimiter.SetFocus
    Exit Sub
  End If
  
  If CheckB4Processing = False Then Exit Sub
  Call ReadFromSpreadsheet
  ReDim ColData(1 To RowCnt) As Integer
'  ReDim ColData(1 To 35) As Integer

  For x = 1 To RowCnt
    vaSpread.Col = 2
    vaSpread.Row = x
    If QPTrim$(vaSpread.Text) = "" Then
      ThisCol = -1
      GoTo NextOne
    End If
    ThisPos = InStr(vaSpread.Text, ".")
    ThisCol = CInt(Mid(vaSpread.Text, 1, ThisPos - 1))
    ColData(x) = ThisCol
NextOne:
  Next x
  
  KillFile ConversionFile
  OpenTempConvFile TempHandle, NumOfTempRecs
  FirstLine = True
  If Exist("ParcelsText.csv") Then
    LHandle = FreeFile
    ThisFile = "ParcelsText.csv"
    Open ThisFile For Input As #LHandle
      Do While Not eof(LHandle)
      GoSub ClearData
      Line Input #LHandle, Textline
      If FirstLine = True Then
        FirstLine = False
        GoTo LoopIt
      End If
      TextLen = Len(Textline)
      Textline = Textline + dlm
      For x = 1 To TextLen + 1
        Thisch = Mid(Textline, x, 1)
        If Thisch = dlm Then
          ColCnt = ColCnt + 1
          For y = 1 To RowCnt
            If ColData(y) = ColCnt Then 'match up columns
              Select Case y
                Case 1
                  TempRec.CData.CustName = ThisWord
                  Exit For
                Case 2
                  TempRec.CData.CountyAcctString = ThisWord
                  Exit For
                Case 3
                  If QPTrim$(ThisWord) = "" Then ThisWord = "0"
                  TempRec.CData.CountyAcct = CDbl(ThisWord)
                  Exit For
                Case 4
                  TempRec.CData.Addr1 = ThisWord
                  Exit For
                Case 5
                  TempRec.CData.Addr2 = ThisWord
                  Exit For
                Case 6
                  TempRec.CData.ServiceAdd = ThisWord
                  Exit For
                Case 7
                  TempRec.CData.City = ThisWord
                  Exit For
                Case 8
                  TempRec.CData.State = QPTrim$(ThisWord)
                  Exit For
                Case 9
                  TempRec.CData.Zip = ThisWord
                  Exit For
                Case 10
                  If QPTrim$(ThisWord) = "" Then ThisWord = "0"
                  TempRec.CData.CSSN = ThisWord
                  Exit For
                Case 11
                  If QPTrim$(ThisWord) = "" Then ThisWord = "0"
                  TempRec.CData.OSSN = ThisWord
                  Exit For
                Case 12
                  If QPTrim$(ThisWord) = "" Then ThisWord = "0"
                  TempRec.CData.HPHONE = ThisWord
                  Exit For
                Case 13
                  If QPTrim$(ThisWord) = "" Then ThisWord = "0"
                  TempRec.CData.WPHONE = ThisWord
                  Exit For
                Case 14
                  If QPTrim$(ThisWord) = "" Then ThisWord = "0"
                  TempRec.CData.OptSrchDesc = ThisWord
                  Exit For
                Case 15
                  If QPTrim$(ThisWord) = "" Then ThisWord = "0"
                  TempRec.CData.SName = ThisWord
                  Exit For
                Case 16
                  TempRec.CData.CTownShip = ThisWord
                  Exit For
                Case 17
                  TempRec.CData.County4BillName = ThisWord
                  Exit For
                Case 18
                  If QPTrim$(ThisWord) = "" Then ThisWord = "0"
                  TempRec.CData.County4BillNum = CLng(ThisWord)
                  Exit For
                Case 19
                  If QPTrim$(ThisWord) = "" Then ThisWord = "0"
                  TempRec.CData.Cycle = CLng(ThisWord)
                  Exit For
                Case 20
                  TempRec.CData.CycleName = ThisWord
                  Exit For
                Case 21
                  TempRec.CData.CustLateNotice = ThisWord
                  Exit For
                Case 22
                  TempRec.CData.TaxExempt = ThisWord
                  Exit For
                Case 23
                  TempRec.CData.Bankrupt = ThisWord
                  Exit For
                Case 24
                  TempRec.CData.Interest = ThisWord
                  Exit For
                Case 25
                  TempRec.CData.LateNotice = ThisWord
                  Exit For
                Case 26
                  TempRec.CData.Penalty = ThisWord
                  Exit For
                Case 27
                  TempRec.CData.PostalRt = ThisWord
                  Exit For
                Case 28
                  TempRec.CData.DeliveryPt = ThisWord
                  Exit For
                Case 29
                  TempRec.CData.DrvrsLic = ThisWord
                  Exit For
                Case 30
                  TempRec.CData.PPinNum = ThisWord
                  Exit For
                Case 31
                  If QPTrim$(ThisWord) = "" Then ThisWord = "0"
                  TempRec.CData.PersVal = CDbl(ThisWord)
                  Exit For
                Case 32
                  If QPTrim$(ThisWord) = "" Then ThisWord = "0"
                  TempRec.CData.MHValue = CDbl(ThisWord)
                  Exit For
                Case 33
                  If QPTrim$(ThisWord) = "" Then ThisWord = "0"
                  TempRec.CData.MCValue = CDbl(ThisWord)
                  Exit For
                Case 34
                  If QPTrim$(ThisWord) = "" Then ThisWord = "0"
                  TempRec.CData.CVALUE = CDbl(ThisWord)
                  Exit For
                Case 35
                  If QPTrim$(ThisWord) = "" Then ThisWord = "0"
                  TempRec.CData.MTValue = CDbl(ThisWord)
                  Exit For
                Case 36
                  If QPTrim$(ThisWord) = "" Then ThisWord = "0"
                  TempRec.CData.PEXMPSENI = CDbl(ThisWord)
                  Exit For
                Case 37
                  If QPTrim$(ThisWord) = "" Then ThisWord = "0"
                  TempRec.CData.PEXMPOTHR = CDbl(ThisWord)
                  Exit For
                Case 38
                  TempRec.CData.PDESC1 = ThisWord
                  Exit For
                Case 39
                  TempRec.CData.PDESC2 = ThisWord
                  Exit For
                Case 40
                  TempRec.CData.PDESC3 = ThisWord
                  Exit For
                Case 41
                  TempRec.CData.PDESC4 = ThisWord
                  Exit For
                Case 42
                  TempRec.CData.PDESC5 = ThisWord
                  Exit For
                Case 43
                  If QPTrim$(ThisWord) = "" Then ThisWord = "-1" '10/9/06 -1 means leave old data
                  TempRec.CData.POptRev1Chrg = CInt(ThisWord)
                  Exit For
                Case 44
                  If QPTrim$(ThisWord) = "" Then ThisWord = "-1"
                  TempRec.CData.POptRev2Chrg = CInt(ThisWord)
                  Exit For
                Case 45
                  If QPTrim$(ThisWord) = "" Then ThisWord = "-1"
                  TempRec.CData.POptRev3Chrg = CInt(ThisWord)
                  Exit For
                Case 46
                  If QPTrim$(ThisWord) = "" Then ThisWord = "0"
                  TempRec.CData.PPTRAYN = ThisWord
                  Exit For
                Case 47
                  TempRec.CData.Vin = ThisWord
                  Exit For
                Case 48
                  If QPTrim$(ThisWord) = "" Then ThisWord = "0"
                  TempRec.CData.Weight = CDbl(ThisWord)
                  Exit For
                Case 49
                  TempRec.CData.MakeMod = ThisWord
                  Exit For
                Case 50
                  If QPTrim$(ThisWord) = "" Then ThisWord = "0"
                  TempRec.CData.ModYear = CInt(ThisWord)
                  Exit For
                Case 51
                  TempRec.CData.Prorate = ThisWord
                  Exit For
                Case 52
                  If QPTrim$(ThisWord) = "" Then ThisWord = "0"
                  TempRec.CData.ProrateVal = CInt(ThisWord)
                  Exit For
                Case 53
                  TempRec.CData.RPinNum = ThisWord
                  Exit For
                Case 54
                  If QPTrim$(ThisWord) = "" Then ThisWord = "0"
                  TempRec.CData.PROPVALU = CDbl(ThisWord)
                  Exit For
                Case 55
                  If QPTrim$(ThisWord) = "" Then ThisWord = "0"
                  TempRec.CData.BLDGVAL = CDbl(ThisWord)
                  Exit For
                Case 56
                  If QPTrim$(ThisWord) = "" Then ThisWord = "0"
                  TempRec.CData.REXMPSENI = CDbl(ThisWord)
                  Exit For
                Case 57
                  If QPTrim$(ThisWord) = "" Then ThisWord = "0"
                  TempRec.CData.REXMPOTHR = CDbl(ThisWord)
                  Exit For
                Case 58
                  If QPTrim$(ThisWord) = "" Then ThisWord = "0"
                  TempRec.CData.PropSize = CDbl(ThisWord)
                  Exit For
                Case 59
                  TempRec.CData.LOTACRE = ThisWord
                  Exit For
                Case 60
                  TempRec.CData.RealAdd = ThisWord
                  Exit For
                Case 61
                  TempRec.CData.BLOCK = ThisWord
                  Exit For
                Case 62
                  TempRec.CData.Map = ThisWord
                  Exit For
                Case 63
                  TempRec.CData.RDESC1 = ThisWord
                  Exit For
                Case 64
                  TempRec.CData.RDESC2 = ThisWord
                  Exit For
                Case 65
                  TempRec.CData.RDESC3 = ThisWord
                  Exit For
                Case 66
                  If QPTrim$(ThisWord) = "" Then ThisWord = "0"
                  TempRec.CData.LOTNUMB = ThisWord
                  Exit For
                Case 67
                  If QPTrim$(ThisWord) = "" Then ThisWord = "-1"
                  TempRec.CData.ROptRev1Chrg = CInt(ThisWord)
                  Exit For
                Case 68
                  If QPTrim$(ThisWord) = "" Then ThisWord = "-1"
                  TempRec.CData.ROptRev2Chrg = CInt(ThisWord)
                  Exit For
                Case 69
                  If QPTrim$(ThisWord) = "" Then ThisWord = "-1"
                  TempRec.CData.ROptRev3Chrg = CInt(ThisWord)
                  Exit For
                Case 70
                  TempRec.CData.RealOptSearch = ThisWord
                  Exit For
                Case 71
                  TempRec.CData.RTownShip = ThisWord
                  Exit For
                Case 72
                  TempRec.CData.GISPOS = ThisWord
                  Exit For
                Case 73
                  TempRec.CData.RLateList = ThisWord
                  Exit For
                Case 74
                  TempRec.CData.MORTCODE = ThisWord
                  Exit For
                Case 75
                  TempRec.CData.ICPDesc = ThisWord
                  Exit For
                Case 76
                  TempRec.CData.LienYN = ThisWord
                  Exit For
                Case 77
                  TempRec.CData.LienDesc = ThisWord
                  Exit For
                Case 78
                  TempRec.CData.PersOptSearch = ThisWord
                  Exit For
                Case 79
                  TempRec.CData.PLateList = ThisWord
                  Exit For
                Case Else
              End Select
            End If
          Next y
          ThisWord = ""
          GoTo NewWord
        End If
        ThisWord = ThisWord + Thisch
NewWord:
      Next x
      ColCnt = 0
      RecCnt = RecCnt + 1
      Put TempHandle, RecCnt, TempRec
LoopIt:
    Loop
  End If
  
  Close
  Call Save
  Call Savemsg(800, "Spreadsheet data has been saved successfully. Data is ready for conversion")
  Exit Sub
  
ClearData:
  TempRec.CData.CustName = ""
  TempRec.CData.CountyAcctString = ""
  TempRec.CData.CountyAcct = 0
  TempRec.CData.Addr1 = ""
  TempRec.CData.Addr2 = ""
  TempRec.CData.ServiceAdd = ""
  TempRec.CData.City = ""
  TempRec.CData.State = ""
  TempRec.CData.Zip = ""
  TempRec.CData.CSSN = ""
  TempRec.CData.OSSN = ""
  TempRec.CData.HPHONE = ""
  TempRec.CData.WPHONE = ""
  TempRec.CData.OptSrchDesc = ""
  TempRec.CData.SName = ""
  TempRec.CData.CTownShip = ""
  TempRec.CData.County4BillName = ""
  TempRec.CData.County4BillNum = 0
  TempRec.CData.Cycle = 0
  TempRec.CData.CycleName = ""
  TempRec.CData.CustLateNotice = ""
  TempRec.CData.TaxExempt = ""
  TempRec.CData.Bankrupt = ""
  TempRec.CData.Interest = ""
  TempRec.CData.LateNotice = ""
  TempRec.CData.Penalty = ""
  TempRec.CData.PostalRt = ""
  TempRec.CData.DeliveryPt = ""
  TempRec.CData.DrvrsLic = ""
  TempRec.CData.PPinNum = ""
  TempRec.CData.PersVal = 0
  TempRec.CData.MHValue = 0
  TempRec.CData.MCValue = 0
  TempRec.CData.CVALUE = 0
  TempRec.CData.MTValue = 0
  TempRec.CData.PEXMPSENI = 0
  TempRec.CData.PEXMPOTHR = 0
  TempRec.CData.PDESC1 = ""
  TempRec.CData.PDESC2 = ""
  TempRec.CData.PDESC3 = ""
  TempRec.CData.PDESC4 = ""
  TempRec.CData.PDESC5 = ""
  TempRec.CData.POptRev1Chrg = 0
  TempRec.CData.POptRev2Chrg = 0
  TempRec.CData.POptRev3Chrg = 0
  TempRec.CData.PPTRAYN = ""
  TempRec.CData.Vin = ""
  TempRec.CData.Weight = 0
  TempRec.CData.MakeMod = ""
  TempRec.CData.ModYear = 0
  TempRec.CData.Prorate = ""
  TempRec.CData.ProrateVal = 0
  TempRec.CData.RPinNum = ""
  TempRec.CData.PROPVALU = 0
  TempRec.CData.BLDGVAL = 0
  TempRec.CData.REXMPSENI = 0
  TempRec.CData.REXMPOTHR = 0
  TempRec.CData.PropSize = 0
  TempRec.CData.LOTACRE = ""
  TempRec.CData.RealAdd = ""
  TempRec.CData.BLOCK = ""
  TempRec.CData.Map = ""
  TempRec.CData.RDESC1 = ""
  TempRec.CData.RDESC2 = ""
  TempRec.CData.RDESC3 = ""
  TempRec.CData.LOTNUMB = ""
  TempRec.CData.ROptRev1Chrg = 0
  TempRec.CData.ROptRev2Chrg = 0
  TempRec.CData.ROptRev3Chrg = 0
  TempRec.CData.RealOptSearch = ""
  TempRec.CData.RTownShip = ""
  TempRec.CData.GISPOS = ""
  TempRec.CData.RLateList = ""
  TempRec.CData.MORTCODE = ""
  TempRec.CData.ICPDesc = ""
  TempRec.CData.LienYN = ""
  TempRec.CData.LienDesc = ""
  TempRec.CData.PersOptSearch = ""
  TempRec.CData.PLateList = ""
 Return
 
End Sub

Private Sub Save()
  Dim SpreadRec As ConvSpreadsheet
  Dim SHandle As Integer
  Dim NumOfSRecs As Integer
  Dim x As Integer
  
  KillFile ConvSpreadFile
  OpenConvSpreadFile SHandle, NumOfSRecs
  For x = 1 To RowCnt
    vaSpread.Row = x
    vaSpread.Col = 1
    SpreadRec.Field1 = vaSpread.Text
    vaSpread.Col = 2
    SpreadRec.Field2 = vaSpread.Text
    vaSpread.Col = 3
    SpreadRec.Field3 = vaSpread.Text
    Put SHandle, x, SpreadRec
  Next x
    
  Close SHandle
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
    Case vbKeyF5:
      SendKeys "%L"
      Call cmdLoadCoData_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call FixSpread
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTCMatchUp1.")
      End
    End If
  End If

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub LoadMe()
  Dim x As Long
  Dim Textline$
  Dim ThisFile$
  Dim LHandle As Integer
  Dim WordCnt As Integer
  Dim TextLen As Integer
  Dim Thisch As String
  Dim ThisWord$
  Dim SpreadRec As ConvSpreadsheet
  Dim SHandle As Integer
  Dim NumOfSRecs As Integer
  
  RowCnt = 79
  
  Label2.Visible = False
'  Clipboard.Clear
  cmdLoadCoData.Visible = False
  cmdProcess.Visible = False
  fptxtDelimiter.Visible = True
  Label1.Visible = True
  If Exist(ConvSpreadFile) Then
    Label2.Visible = True
    cmdProcess.Visible = True
    OpenConvSpreadFile SHandle, NumOfSRecs
    For x = 1 To NumOfSRecs
      Get SHandle, x, SpreadRec
      vaSpread.Row = x
      vaSpread.Col = 1
      vaSpread.Text = SpreadRec.Field1
      vaSpread.Col = 2
      vaSpread.Text = SpreadRec.Field2
      vaSpread.Col = 3
      vaSpread.Text = SpreadRec.Field3
    Next x
    Close SHandle
    WordCnt = NumOfSRecs
  Else
    dlm = fptxtDelimiter.Text
    cmdLoadCoData.Visible = True
    vaSpread.Col = 1
    vaSpread.Row = 1
    vaSpread.Text = "Customer Name"
    vaSpread.Row = 2
    vaSpread.Text = "County Pin String"
    vaSpread.Row = 3
    vaSpread.Text = "County Pin Number"
    vaSpread.Row = 4
    vaSpread.Text = "Customer Address #1"
    vaSpread.Row = 5
    vaSpread.Text = "Customer Address #2"
    vaSpread.Row = 6
    vaSpread.Text = "Service Address"
    vaSpread.Row = 7
    vaSpread.Text = "City"
    vaSpread.Row = 8
    vaSpread.Text = "State"
    vaSpread.Row = 9
    vaSpread.Text = "Zip Code"
    vaSpread.Row = 10
    vaSpread.Text = "Customer SSN"
    vaSpread.Row = 11
    vaSpread.Text = "Other SSN"
    vaSpread.Row = 12
    vaSpread.Text = "Home Phone"
    vaSpread.Row = 13
    vaSpread.Text = "Work Phone"
    vaSpread.Row = 14
    vaSpread.Text = "Opt'l Search"
    vaSpread.Row = 15
    vaSpread.Text = "Search Name"
    vaSpread.Row = 16
    vaSpread.Text = "Customer Township"
    vaSpread.Row = 17
    vaSpread.Text = "County Name"
    vaSpread.Row = 18
    vaSpread.Text = "County Number"
    vaSpread.Row = 19
    vaSpread.Text = "Cycle Number"
    vaSpread.Row = 20
    vaSpread.Text = "Cycle Name"
    vaSpread.Row = 21
    vaSpread.Text = "Cust Late Notice"
    vaSpread.Row = 22
    vaSpread.Text = "Tax Exempt Y/N?"
    vaSpread.Row = 23
    vaSpread.Text = "Bankrupt Y/N?"
    vaSpread.Row = 24
    vaSpread.Text = "Interest Y/N?"
    vaSpread.Row = 25
    vaSpread.Text = "Late Notice Y/N?"
    vaSpread.Row = 26
    vaSpread.Text = "Penalty Y/N?"
    vaSpread.Row = 27
    vaSpread.Text = "Postal Route"
    vaSpread.Row = 28
    vaSpread.Text = "Delivery Point"
    vaSpread.Row = 29
    vaSpread.Text = "Driver's Lic #"
    vaSpread.Row = 30
    vaSpread.Text = "Pers Pin #"
    vaSpread.Row = 31
    vaSpread.Text = "Personal Value"
    vaSpread.Row = 32
    vaSpread.Text = "Mobile Home Value"
    vaSpread.Row = 33
    vaSpread.Text = "Merchant Capital Value"
    vaSpread.Row = 34
    vaSpread.Text = "Farm Equipment Value"
    vaSpread.Row = 35
    vaSpread.Text = "Machine Tools Value"
    vaSpread.Row = 36
    vaSpread.Text = "Pers Senior Exemption"
    vaSpread.Row = 37
    vaSpread.Text = "Pers Other Exemption"
    vaSpread.Row = 38
    vaSpread.Text = "Pers Description 1"
    vaSpread.Row = 39
    vaSpread.Text = "Pers Description 2"
    vaSpread.Row = 40
    vaSpread.Text = "Pers Description 3"
    vaSpread.Row = 41
    vaSpread.Text = "Pers Description 4"
    vaSpread.Row = 42
    vaSpread.Text = "Pers Description 5"
    vaSpread.Row = 43
    vaSpread.Text = "Pers Opt Rev #1"
    vaSpread.Row = 44
    vaSpread.Text = "Pers Opt Rev #2"
    vaSpread.Row = 45
    vaSpread.Text = "Pers Opt Rev #3"
    vaSpread.Row = 46
    vaSpread.Text = "PPTRA Y/N?"
    vaSpread.Row = 47
    vaSpread.Text = "Vin #"
    vaSpread.Row = 48
    vaSpread.Text = "Vehicle Weight"
    vaSpread.Row = 49
    vaSpread.Text = "Make/Model"
    vaSpread.Row = 50
    vaSpread.Text = "Model Year"
    vaSpread.Row = 51
    vaSpread.Text = "Prorate Y/N?"
    vaSpread.Row = 52
    vaSpread.Text = "Prorate Value"
    vaSpread.Row = 53
    vaSpread.Text = "Real Pin #"
    vaSpread.Row = 54
    vaSpread.Text = "Real Land Value"
    vaSpread.Row = 55
    vaSpread.Text = "Building Value"
    vaSpread.Row = 56
    vaSpread.Text = "Real Senior Exemption"
    vaSpread.Row = 57
    vaSpread.Text = "Real Other Exemption"
    vaSpread.Row = 58
    vaSpread.Text = "Parcel Size"
    vaSpread.Row = 59
    vaSpread.Text = "Lot/Acre"
    vaSpread.Row = 60
    vaSpread.Text = "Real Address"
    vaSpread.Row = 61
    vaSpread.Text = "Block"
    vaSpread.Row = 62
    vaSpread.Text = "Map"
    vaSpread.Row = 63
    vaSpread.Text = "Real Description 1"
    vaSpread.Row = 64
    vaSpread.Text = "Real Description 2"
    vaSpread.Row = 65
    vaSpread.Text = "Real Description 3"
    vaSpread.Row = 66
    vaSpread.Text = "Lot #"
    vaSpread.Row = 67
    vaSpread.Text = "Real Opt Rev #1"
    vaSpread.Row = 68
    vaSpread.Text = "Real Opt Rev #2"
    vaSpread.Row = 69
    vaSpread.Text = "Real Opt Rev #3"
    vaSpread.Row = 70
    vaSpread.Text = "Real Opt'l Search"
    vaSpread.Row = 71
    vaSpread.Text = "Real Township"
    vaSpread.Row = 72
    vaSpread.Text = "GIS Position"
    vaSpread.Row = 73
    vaSpread.Text = "Real Late List Y/N?"
    vaSpread.Row = 74
    vaSpread.Text = "Mortgage Code"
    vaSpread.Row = 75
    vaSpread.Text = "Industry, Commercial, Private"
    vaSpread.Row = 76
    vaSpread.Text = "Lien Y/N?"
    vaSpread.Row = 77
    vaSpread.Text = "Lien Description"
    vaSpread.Row = 78
    vaSpread.Text = "Personal Opt'l Search"
    vaSpread.Col = 4
    vaSpread.Row = 79
    vaSpread.Text = "Pers Late List Y/N?"
  End If
  
  If fptxtDelimiter.Text <> "" Then
    Call ReadFromSpreadsheet
  End If
  
  If WordCnt = 0 Then WordCnt = RowCnt
  
  vaSpread.Col = 4
  For x = 1 To WordCnt
    vaSpread.Row = x
    vaSpread.Text = CStr(x)
  Next x
  
  Close
  Call TCMsg(800, "Please note that you can only assign one available county data field to one assigned county data field.")
End Sub

Private Sub EditCopyProc(Text$)
   ' Copy selected text onto Clipboard.
'   Clipboard.Clear
'   Clipboard.SetText Text
  MyClip = Text
End Sub
Private Sub vaSpread_DblClick(ByVal Col As Long, ByVal Row As Long)
  Dim ThisOne$
  
  If Col = 3 Then
'    Clipboard.Clear
    vaSpread.Col = 3
    vaSpread.Row = Row
    ThisOne = vaSpread.Text
    vaSpread.Col = 4
    ThisOne = vaSpread.Text + ". " + ThisOne
    Call EditCopyProc(ThisOne$)
  ElseIf Col = 2 Then
    vaSpread.Col = 2
    vaSpread.Row = Row
'    vaSpread.Text = Clipboard.GetText
    vaSpread.Text = MyClip
  End If

End Sub

Private Function FixSpread()
  Dim COne As Integer
  Dim CTwo As Integer
  Dim CThree As Integer
  Dim CFour As Integer
  Dim CFive As Integer
  Dim CSix As Integer
  Dim coladj As Integer
  Dim x As Integer, y As Integer
  '-1 means all rows or all columns....0 means headers
    Select Case ScreenW
      Case 1280
      If Screen.TwipsPerPixelX <> 12 Then
        COne = 21
        coladj = 10
        For x = 0 To 7
          For y = 0 To 2
            vaSpread.FontName = "Tahoma"
            vaSpread.Col = y
            vaSpread.Row = x
            vaSpread.FontSize = 16
          Next y
        Next x
        vaSpread.RowHeight(-1) = 27.5
        vaSpread.RowHeight(0) = 27.5
      Else
        COne = 11.25
        coladj = 4.45
        For x = 0 To RowCnt
          For y = 0 To 4
            vaSpread.FontName = "Tahoma"
            vaSpread.Col = y
            vaSpread.Row = x
            vaSpread.FontSize = 12
          Next y
        Next x
        vaSpread.RowHeight(-1) = 15
        vaSpread.RowHeight(0) = 15
      End If
      Case 1152
      If Screen.TwipsPerPixelX <> 12 Then
        COne = 15
        coladj = 7
        For x = 0 To 7
          For y = 0 To 2
            vaSpread.FontName = "Tahoma"
            vaSpread.Col = y
            vaSpread.Row = x
            vaSpread.FontSize = 14
          Next y
        Next x
        vaSpread.RowHeight(0) = 24
        vaSpread.RowHeight(-1) = 22
      Else
        COne = 6
        coladj = 2.3
        For x = 0 To 7
          For y = 0 To 2
            vaSpread.FontName = "Tahoma"
            vaSpread.Col = y
            vaSpread.Row = x
            vaSpread.FontSize = 11
          Next y
        Next x
        vaSpread.RowHeight(0) = 19.5
        vaSpread.RowHeight(-1) = 19.5
      End If
      Case 1024
      If Screen.TwipsPerPixelX <> 12 Then
        COne = 8
        coladj = 6
        For x = 0 To 7
          For y = 0 To 2
            vaSpread.FontName = "Tahoma"
            vaSpread.Col = y
            vaSpread.Row = x
            vaSpread.FontSize = 12
          Next y
        Next x
        vaSpread.RowHeight(0) = 19.5
'        vaSpread.FontBold = True
        vaSpread.RowHeight(-1) = 19.5
      Else
        COne = 0.5
        coladj = 1.6
      End If
      Case 800
        COne = -0.6
        coladj = 1.55
        For x = 0 To 7
          For y = 0 To 2
            vaSpread.FontName = "Tahoma"
            vaSpread.Col = y
            vaSpread.Row = x
            vaSpread.FontSize = 10
          Next y
        Next x
        vaSpread.RowHeight(0) = 14.75
        vaSpread.RowHeight(-1) = 14.75
      Case Else
       
    End Select
    vaSpread.ColWidth(1) = vaSpread.ColWidth(1) + COne
    vaSpread.ColWidth(2) = vaSpread.ColWidth(2) + coladj
    vaSpread.ColWidth(3) = vaSpread.ColWidth(2) + coladj
    vaSpread.ColWidth(4) = vaSpread.ColWidth(4) + 1 '- coladj

End Function

Private Sub ReadFromSpreadsheet()
  Dim x As Long
  Dim Textline$
  Dim ThisFile$
  Dim LHandle As Integer
  Dim WordCnt As Integer
  Dim TextLen As Integer
  Dim Thisch As String
  Dim ThisWord$
  Dim SpreadRec As ConvSpreadsheet
  Dim SHandle As Integer
  Dim NumOfSRecs As Integer
  
  dlm = fptxtDelimiter.Text
  WordCnt = 0
  ReDim Words(1 To 1) As String
  
  If Exist("ParcelsText.csv") Then
    LHandle = FreeFile
    ThisFile = "ParcelsText.csv"
    Open ThisFile For Input As #LHandle
    Line Input #LHandle, Textline
    TextLen = Len(Textline)
    Textline = Textline + dlm
    For x = 1 To TextLen + 1
      Thisch = Mid(Textline, x, 1)
      If Thisch = dlm Then
        WordCnt = WordCnt + 1
        ReDim Preserve Words(1 To WordCnt) As String
        Words(WordCnt) = ThisWord
        ThisWord = ""
        GoTo NewWord
      End If
      ThisWord = ThisWord + Thisch
NewWord:
    Next x
    vaSpread.Col = 3
    For x = 1 To WordCnt
      vaSpread.Row = x
      vaSpread.Text = Words(x)
    Next x
  Else
    Call TCMsg(900, "The file 'ParcelsText.csv' cannot be found.")
    Exit Sub
  End If
  
  Close
  
'  RowCnt = WordCnt
  vaSpread.Col = 4
  For x = 1 To WordCnt
    vaSpread.Row = x
    vaSpread.Text = CStr(x)
  Next x
  
End Sub

Private Function CheckB4Processing() As Boolean
  Dim x As Integer
  
  CheckB4Processing = True
  vaSpread.Col = 2
  For x = 1 To RowCnt
    vaSpread.Row = x
    If QPTrim$(vaSpread.Text) <> "" Then
      Exit For
    End If
  Next x
  
  If x > RowCnt Then
    Call TCMsg(800, "Nothing to process. Enter data in 'Assigned County Fields' column. Processing aborted.")
    CheckB4Processing = False
  End If
  
End Function

