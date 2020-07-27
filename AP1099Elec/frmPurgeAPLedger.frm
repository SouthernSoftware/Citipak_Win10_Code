VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmPurgeAPLedger 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purge A/P Ledger"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   12195
   Icon            =   "frmPurgeAPLedger.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   765
      Left            =   2340
      TabIndex        =   7
      Top             =   6960
      Width           =   1485
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 &Go"
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
      Left            =   7584
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7584
      Width           =   1332
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
      Left            =   9312
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7584
      Width           =   1332
   End
   Begin EditLib.fpDateTime txtDate 
      Height          =   372
      Left            =   6678
      TabIndex        =   2
      Top             =   4128
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
      ButtonColor     =   14737632
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   8508
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
            TextSave        =   "3:07 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "10/2/2009"
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please Make Sure All Operators Have Exited The Program.  You Should Also Have A Backup Of The Data Before Continuing."
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
      Height          =   732
      Left            =   2418
      TabIndex        =   6
      Top             =   3024
      Width           =   7356
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Purge A/P Ledger History"
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
      Left            =   4074
      TabIndex        =   5
      Top             =   1416
      Width           =   4044
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3216
      Top             =   1176
      Width           =   5772
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter The Closing Date:"
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
      Left            =   3822
      TabIndex        =   4
      Top             =   4200
      Width           =   2676
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   2244
      Left            =   2166
      Top             =   2736
      Width           =   7860
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      Height          =   972
      Left            =   3210
      Top             =   1056
      Width           =   5772
   End
End
Attribute VB_Name = "frmPurgeAPLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim GLSetup As GLSetupRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim ApLedger As APLedger81RecType
Dim APDist As APDistRecType

Private Sub Command1_Click()
  PackLedgerFileLilesville
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  txtDate.Text = Format(Now, "mm/dd/yyyy")

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
  frmAPLdgUtilMenu.Show
  Unload frmPurgeAPLedger
End Sub

Private Sub cmdGo_Click()
  Call MainLog("Purge APLedger Started.")
  PackLedgerFile
  frmAPLdgUtilMenu.RelinkDist2Trans
  frmAPLdgUtilMenu.RelinkLedger2Vendor
  Me.cmdExit.Enabled = True
  Me.cmdGo.Enabled = True
  EnableCloseButton Me.hwnd, True
  Call MainLog("Purge APLedger Complete.")
  frmAPLdgUtilMenu.Show
  Unload frmPurgeAPLedger
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    Cancel = True
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
      SendKeys "%G"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub PackLedgerFileLilesville()
  Dim NAPDist As APDistRecType
  Dim NAPLedger As APLedger81RecType
  Dim APLedgerFile As Integer, NumTran As Long, APLRecLen As Integer
  Dim APDistFile As Integer, NumDistRecs As Long, APDRecLen As Integer
  Dim NAPLRecLen As Integer, NAPLedgerFile As Integer, NumNewTrans As Long
  Dim NAPDRecLen As Integer, NAPDistFile As Integer, NumNewDistRecs As Long
  Dim PurgeDate As Integer, cnt As Long, NextDist As Long
 'look at gosub NewLedgerRec:
  PurgeDate = DateDiff("d", "12/31/1979", "01/01/2011")
  If PurgeDate > 0 Then
    If MsgBox("Purge ALL HISTORY beyond " & Format(DateAdd("d", (PurgeDate), "12-31-1979"), "mm/dd/yyyy"), vbOKCancel, "Continue?") = vbCancel Then
       Exit Sub
    End If
    
   '--Open the A/P Ledger file
   OpenAPLedgerFile APLedgerFile, NumTran&, APLRecLen

   '--Open the Ledger Distribution file
   OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen

   '--Create a new ledger file to hold the new records
   
   NAPLRecLen = Len(ApLedger)
   NAPLedgerFile = FreeFile
   Open "APLEDGER.NEW" For Random As NAPLedgerFile Len = NAPLRecLen
   NumNewTrans& = LOF(NAPLedgerFile) \ NAPLRecLen

   '--Create a new ledger distribution file
   
   NAPDRecLen = Len(NAPDist)
   NAPDistFile = FreeFile
   Open "APDIST.NEW" For Random As NAPDistFile Len = NAPDRecLen
   NumNewDistRecs& = LOF(NAPDistFile) \ NAPDRecLen
   FrmShowPctComp.Label1 = "Purging Ledger."
   FrmShowPctComp.cmdCancel.Enabled = False
   FrmShowPctComp.Show , Me
   DoEvents
   EnableCloseButton Me.hwnd, False
   Me.cmdExit.Enabled = False
   Me.cmdGo.Enabled = False

   For cnt& = 1 To NumTran&
     FrmShowPctComp.ShowPctComp cnt&, NumTran&

''      LOCATE 3, 1
''      Print Using; "Processing Transaction: #####"; cnt&;
''      Print Using; " of ##### "; NumTran&

      Get APLedgerFile, cnt&, ApLedger

      '--Move only transactions greater that the purge date
      '--or any open invoices to new file.
       If ApLedger.TRDATE < PurgeDate Then 'OR ApLedger.PayCode < 3 THEN
      '   IF Cnt& < 626 THEN
         GoSub NewLedgerRec
         NextDist& = ApLedger.FrstDist
         If NextDist& > 0 Then
            Do
              Get APDistFile, NextDist&, APDist
              NextDist& = APDist.NextDist
              GoSub NewDistRec
            Loop Until NextDist& = 0
         End If
      End If

   Next

   Close
   '--keep the old files
   KillFileD "APLEDGER.OLD"
   KillFileD "APDIST.OLD"

   Name "APLEDGER.DAT" As "APLEDGER.OLD"
   Name "APDIST.DAT" As "APDIST.OLD"

   '--rename the new files
   Name "APLEDGER.NEW" As "APLEDGER.DAT"
   Name "APDIST.NEW" As "APDIST.DAT"

   'MsgBox "Press any key to continue with relink.", vbOKOnly, "Continue.."
  
End If
Exit Sub
NewLedgerRec:
   'STOP 'see comment below

   NumNewTrans& = NumNewTrans& + 1

   '--Version 8
   'NAPLedger.VIN = apledger.VIN
   'NAPLedger.VendorCode = apledger.VendorCode
   'NAPLedger.VRecNum = apledger.VRecNum
   'NAPLedger.TrDate = apledger.TrDate
   'NAPLedger.GLDistDate = apledger.GLDistDate
   'NAPLedger.DUEDATE = apledger.DUEDATE
   'NAPLedger.TrCode = apledger.TrCode
   'NAPLedger.DOCNum = apledger.DOCNum
   'NAPLedger.PONUM = apledger.PONUM
   'NAPLedger.PayCode = apledger.PayCode
   'NAPLedger.PrintCode = apledger.PrintCode
   'NAPLedger.PDCheckNum = apledger.PDCheckNum
   'NAPLedger.PDCheckDate = apledger.PDCheckDate
   'NAPLedger.MiscCode = apledger.MiscCode
   'NAPLedger.Amt = apledger.Amt
   ''--These fields will have to be relinked
   'NAPLedger.FrstDist = 0
   'NAPLedger.LastDist = 0
   'NAPLedger.NextTrans = 0

   '--Version 8.1
   NAPLedger.VIN = ApLedger.VIN
   NAPLedger.VendorCode = ApLedger.VendorCode
   NAPLedger.VRecNum = ApLedger.VRecNum
   NAPLedger.TRDATE = ApLedger.TRDATE
   NAPLedger.GLDistDate = ApLedger.GLDistDate
   NAPLedger.DueDate = ApLedger.DueDate
   NAPLedger.TRCode = ApLedger.TRCode
   NAPLedger.DOCNum = ApLedger.DOCNum
   NAPLedger.PONum = ApLedger.PONum
   NAPLedger.MPONum = ApLedger.MPONum
   '--Fix me according to what you're doing!
   NAPLedger.PAYCODE = ApLedger.PAYCODE
   'NAPLedger.PayCode = 1 '--Sets invoices back to open!

   NAPLedger.PrintCode = ApLedger.PrintCode
   NAPLedger.PDCheckNum = ApLedger.PDCheckNum
   NAPLedger.PDCheckDate = ApLedger.PDCheckDate
   NAPLedger.Comment = ApLedger.Comment
   NAPLedger.PSLFlag = ApLedger.PSLFlag
   NAPLedger.Get1099 = ApLedger.Get1099
   NAPLedger.Amt = ApLedger.Amt
   NAPLedger.FrstDist = 0 'apledger.FrstDist  'Relink these
   NAPLedger.LastDist = 0 'apledger.LastDist
   NAPLedger.NextTrans = 0 'apledger.NextTrans
   NAPLedger.TaxAmt = ApLedger.TaxAmt
   NAPLedger.Bankcode = ApLedger.Bankcode
   NAPLedger.ChkByte = ApLedger.ChkByte
   
   Put NAPLedgerFile, NumNewTrans&, NAPLedger

'   LOCATE 4, 1
'   Print "New Transactions: "; NumNewTrans&

Return

NewDistRec:

   NumNewDistRecs& = NumNewDistRecs& + 1

   '--Set ledger record key to new ledger record
   NAPDist.APLedgerRec = NumNewTrans&

   '--These fields stays the same
   NAPDist.DistAcctRec = APDist.DistAcctRec
   NAPDist.DistAcctNum = APDist.DistAcctNum
   NAPDist.DistAmt = APDist.DistAmt
   NAPDist.DistStat = APDist.DistStat
   NAPDist.NextDist = 0

   '--Relinking on the fly. DO NOT USE..UNTESTED
   'SELECT CASE NextDist&
   '   CASE 0
   '      '--No more distributions
   '      NAPDist.NextDist = 0
   '   CASE ELSE
   '      '--if There is another distribution
   '      '--It will be the next record number
   '      NAPDist.NextDist = NumNewDistRecs + 1

   'END SELECT

   Put NAPDistFile, NumNewDistRecs&, NAPDist
'   LOCATE 5, 1, 0
'   Print "New Distributions: "; NumNewDistRecs&

Return


End Sub

Private Sub PackLedgerFile()
  Dim NAPDist As APDistRecType
  Dim NAPLedger As APLedger81RecType
  Dim APLedgerFile As Integer, NumTran As Long, APLRecLen As Integer
  Dim APDistFile As Integer, NumDistRecs As Long, APDRecLen As Integer
  Dim NAPLRecLen As Integer, NAPLedgerFile As Integer, NumNewTrans As Long
  Dim NAPDRecLen As Integer, NAPDistFile As Integer, NumNewDistRecs As Long
  Dim PurgeDate As Integer, cnt As Long, NextDist As Long
 'look at gosub NewLedgerRec:
  PurgeDate = DateDiff("d", "12/31/1979", txtDate) '"01/01/2011")
  If PurgeDate > 0 Then
    If MsgBox("Purge ALL HISTORY through " & Format(DateAdd("d", (PurgeDate), "12-31-1979"), "mm/dd/yyyy"), vbOKCancel, "Continue?") = vbCancel Then
       Exit Sub
    End If
    
   '--Open the A/P Ledger file
   OpenAPLedgerFile APLedgerFile, NumTran&, APLRecLen

   '--Open the Ledger Distribution file
   OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen

   '--Create a new ledger file to hold the new records
   
   NAPLRecLen = Len(ApLedger)
   NAPLedgerFile = FreeFile
   Open "APLEDGER.NEW" For Random As NAPLedgerFile Len = NAPLRecLen
   NumNewTrans& = LOF(NAPLedgerFile) \ NAPLRecLen

   '--Create a new ledger distribution file
   
   NAPDRecLen = Len(NAPDist)
   NAPDistFile = FreeFile
   Open "APDIST.NEW" For Random As NAPDistFile Len = NAPDRecLen
   NumNewDistRecs& = LOF(NAPDistFile) \ NAPDRecLen
   FrmShowPctComp.Label1 = "Purging Ledger."
   FrmShowPctComp.cmdCancel.Enabled = False
   FrmShowPctComp.Show , Me
   DoEvents
   EnableCloseButton Me.hwnd, False
   Me.cmdExit.Enabled = False
   Me.cmdGo.Enabled = False

   For cnt& = 1 To NumTran&
     FrmShowPctComp.ShowPctComp cnt&, NumTran&

''      LOCATE 3, 1
''      Print Using; "Processing Transaction: #####"; cnt&;
''      Print Using; " of ##### "; NumTran&

      Get APLedgerFile, cnt&, ApLedger

      '--Move only transactions greater that the purge date
      '--or any open invoices to new file.
       If ApLedger.TRDATE > PurgeDate Then 'OR ApLedger.PayCode < 3 THEN
      '   IF Cnt& < 626 THEN
         GoSub NewLedgerRec
         NextDist& = ApLedger.FrstDist
         If NextDist& > 0 Then
            Do
              Get APDistFile, NextDist&, APDist
              NextDist& = APDist.NextDist
              GoSub NewDistRec
            Loop Until NextDist& = 0
         End If
      End If

   Next

   Close
   '--keep the old files
   KillFileD "APLEDGER.OLD"
   KillFileD "APDIST.OLD"

   Name "APLEDGER.DAT" As "APLEDGER.OLD"
   Name "APDIST.DAT" As "APDIST.OLD"

   '--rename the new files
   Name "APLEDGER.NEW" As "APLEDGER.DAT"
   Name "APDIST.NEW" As "APDIST.DAT"

   'MsgBox "Press any key to continue with relink.", vbOKOnly, "Continue.."
  
End If
Exit Sub
NewLedgerRec:
   'STOP 'see comment below

   NumNewTrans& = NumNewTrans& + 1

   '--Version 8
   'NAPLedger.VIN = apledger.VIN
   'NAPLedger.VendorCode = apledger.VendorCode
   'NAPLedger.VRecNum = apledger.VRecNum
   'NAPLedger.TrDate = apledger.TrDate
   'NAPLedger.GLDistDate = apledger.GLDistDate
   'NAPLedger.DUEDATE = apledger.DUEDATE
   'NAPLedger.TrCode = apledger.TrCode
   'NAPLedger.DOCNum = apledger.DOCNum
   'NAPLedger.PONUM = apledger.PONUM
   'NAPLedger.PayCode = apledger.PayCode
   'NAPLedger.PrintCode = apledger.PrintCode
   'NAPLedger.PDCheckNum = apledger.PDCheckNum
   'NAPLedger.PDCheckDate = apledger.PDCheckDate
   'NAPLedger.MiscCode = apledger.MiscCode
   'NAPLedger.Amt = apledger.Amt
   ''--These fields will have to be relinked
   'NAPLedger.FrstDist = 0
   'NAPLedger.LastDist = 0
   'NAPLedger.NextTrans = 0

   '--Version 8.1
   NAPLedger.VIN = ApLedger.VIN
   NAPLedger.VendorCode = ApLedger.VendorCode
   NAPLedger.VRecNum = ApLedger.VRecNum
   NAPLedger.TRDATE = ApLedger.TRDATE
   NAPLedger.GLDistDate = ApLedger.GLDistDate
   NAPLedger.DueDate = ApLedger.DueDate
   NAPLedger.TRCode = ApLedger.TRCode
   NAPLedger.DOCNum = ApLedger.DOCNum
   NAPLedger.PONum = ApLedger.PONum
   NAPLedger.MPONum = ApLedger.MPONum

   '--Fix me according to what you're doing!
   NAPLedger.PAYCODE = ApLedger.PAYCODE
   'NAPLedger.PayCode = 1 '--Sets invoices back to open!

   NAPLedger.PrintCode = ApLedger.PrintCode
   NAPLedger.PDCheckNum = ApLedger.PDCheckNum
   NAPLedger.PDCheckDate = ApLedger.PDCheckDate
   NAPLedger.Comment = ApLedger.Comment
   NAPLedger.PSLFlag = ApLedger.PSLFlag
   NAPLedger.Get1099 = ApLedger.Get1099
   NAPLedger.Amt = ApLedger.Amt
   NAPLedger.FrstDist = 0 'apledger.FrstDist  'Relink these
   NAPLedger.LastDist = 0 'apledger.LastDist
   NAPLedger.NextTrans = 0 'apledger.NextTrans
   NAPLedger.Bankcode = ApLedger.Bankcode
   NAPLedger.TaxAmt = ApLedger.TaxAmt
   NAPLedger.ChkByte = ApLedger.ChkByte

   Put NAPLedgerFile, NumNewTrans&, NAPLedger

'   LOCATE 4, 1
'   Print "New Transactions: "; NumNewTrans&

Return

NewDistRec:

   NumNewDistRecs& = NumNewDistRecs& + 1

   '--Set ledger record key to new ledger record
   NAPDist.APLedgerRec = NumNewTrans&

   '--These fields stays the same
   NAPDist.DistAcctRec = APDist.DistAcctRec
   NAPDist.DistAcctNum = APDist.DistAcctNum
   NAPDist.DistStat = APDist.DistStat
   NAPDist.DistAmt = APDist.DistAmt
   NAPDist.NextDist = 0

   '--Relinking on the fly. DO NOT USE..UNTESTED
   'SELECT CASE NextDist&
   '   CASE 0
   '      '--No more distributions
   '      NAPDist.NextDist = 0
   '   CASE ELSE
   '      '--if There is another distribution
   '      '--It will be the next record number
   '      NAPDist.NextDist = NumNewDistRecs + 1

   'END SELECT

   Put NAPDistFile, NumNewDistRecs&, NAPDist
'   LOCATE 5, 1, 0
'   Print "New Distributions: "; NumNewDistRecs&

Return


End Sub
