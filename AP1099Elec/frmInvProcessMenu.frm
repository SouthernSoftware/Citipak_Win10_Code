VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Begin VB.Form frmInvProcessMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Processing"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   12225
   Icon            =   "frmInvProcessMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdEnterEditInv 
      Height          =   480
      Left            =   4305
      TabIndex        =   0
      Top             =   3030
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmInvProcessMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintInvReg 
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   3810
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmInvProcessMenu.frx":0AB5
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintInvReg2 
      Height          =   492
      Left            =   4302
      TabIndex        =   2
      Top             =   4626
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   868
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmInvProcessMenu.frx":0CA3
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPostInv 
      Height          =   492
      Left            =   4302
      TabIndex        =   3
      Top             =   5427
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   868
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmInvProcessMenu.frx":0E98
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdVoidOpenInv 
      Height          =   492
      Left            =   4302
      TabIndex        =   4
      Top             =   6228
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   868
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmInvProcessMenu.frx":1084
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExitInvMenu 
      Height          =   492
      Left            =   4302
      TabIndex        =   5
      Top             =   7032
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   868
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmInvProcessMenu.frx":1270
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   2400
      X2              =   3360
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   4
      X1              =   8880
      X2              =   9840
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   3360
      X2              =   3360
      Y1              =   2280
      Y2              =   2412
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2400
      X2              =   2400
      Y1              =   2292
      Y2              =   2412
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2400
      X2              =   3360
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   2
      X1              =   9000
      X2              =   9000
      Y1              =   2424
      Y2              =   8280
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   5
      X1              =   8880
      X2              =   9840
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   8880
      X2              =   8880
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   9840
      X2              =   9840
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1800
      Top             =   1080
      Width           =   8652
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "INVOICE PROCESSING MENU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3456
      TabIndex        =   6
      Top             =   1464
      Width           =   5292
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   2520
      X2              =   3216
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2520
      X2              =   2520
      Y1              =   2424
      Y2              =   8280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   9000
      X2              =   9696
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00D0D0D0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1800
      Top             =   960
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00000000&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2400
      Top             =   2160
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00000000&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5952
      Index           =   0
      Left            =   2520
      Top             =   2340
      Width           =   732
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00000000&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8880
      Top             =   2160
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00000000&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5964
      Index           =   2
      Left            =   9000
      Top             =   2328
      Width           =   732
   End
End
Attribute VB_Name = "frmInvProcessMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim low As Long, High As Long
Private Sub cmdEnterEditInv_Click()
  Dim InvBusy As Boolean
  InvBusy = False
  If Exist("APIED.DAT") Then InvBusy = GetAttr("APIED.DAT") And vbReadOnly
  If Not InvBusy Then
    frmLoadingRpt.Show
    DoEvents
    Load frmInvEnterEdit
    DoEvents
    frmInvEnterEdit.Show
    Unload frmLoadingRpt
    Unload frmInvProcessMenu
    frmInvEnterEdit.FirstOpenInv
    Call MainLog("Open Invoice Enter/Edit")
  Else
    MsgBox "Posting Is In Progress, Editing Not Allowed At This Time.", vbOKOnly, "Canceled"
  End If
End Sub

Private Sub cmdPostInv_Click()
  frmPostInvoices.Show
  'do not unload menu
End Sub

Private Sub cmdPrintInvReg_Click()
    frmReportOpt.Show 1
    If rptopt = 1 Then
      PrnEditList False
    ElseIf rptopt = 2 Then
      PrnEditList2 False
    End If
End Sub

Private Sub cmdPrintInvReg2_Click()
    frmReportOpt.Show 1
    If rptopt = 1 Then
      PrnEditList True
    ElseIf rptopt = 2 Then
      PrnEditList2 True
    End If
End Sub

Private Sub cmdVoidOpenInv_Click()

  If Not Exist("APCHKINF.DAT") Then
    If Not Exist("TPAYLIST.LST") Then
      frmInvVoid.Show
      Unload frmInvProcessMenu
    Else
      MsgBox "Invoices have been selected for payment, please complete the check process before voiding invoices.", vbOKOnly, "Access Denied"
    End If
  Else
    MsgBox "Unposted checks exist, please complete the check process before voiding invoices.", vbOKOnly, "Access Denied"
  End If

End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  Me.HelpContextID = hlpInvProcess
  cmdPrintInvReg.HelpContextID = hlpInvReg
  cmdPrintInvReg2.HelpContextID = hlpInvRegVO
End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
'   Me.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExitInvMenu_Click
      KeyCode = 0
      DoEvents
    Case Else:
  End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog "Close AP"
      ClearInUse PWcnt
    End If
  End If
End Sub

Private Sub cmdExitInvMenu_Click()
  frmAPMainMenu.Show
  Unload frmInvProcessMenu
End Sub
Private Sub PrnEditList(LuneyFlag As Boolean)
  Dim CommaFmt As String, FundNum As String, PageNum As Integer
  Dim Comma2Fmt As String, DistSumLine As String, FileName As String
  Dim DebitCol As Integer, CreditCol As Integer, RegTitle As String
  Dim TransTotal As Double, TranCnt As Integer, TranCol As Integer
  Dim PrnFileNum As Integer, NumFunds As Integer, Transaction As Integer
  Dim APEditFile As Integer, NumEdTrans As Integer, Linecnt As Integer
  Dim VendorFile As Integer, NumVRecs As Integer, TempTax As Double
  Dim LedgerRecLen As Integer, ToPrint As String, TaxTotal As Double
  Dim APLedgerFile As Integer, TotTranDist As Double, HCnt As Integer
  Dim AcctDist As Integer, Found As Boolean, Fund As Integer
  Dim cnt As Integer, GrdTot As Double, NumLdgTran As Long
  Dim PrnFileNum2 As Integer, FileSub As String, ToPrintI As String
  Dim ToPrintD As String
  Dim APIED As APInv85Type
  ReDim LuneyIdx(1 To 1) As LuneySortType
  

  DebitCol = 42
  CreditCol = 58
  CommaFmt$ = "###,###,###.##"   'ten millions
  Comma2Fmt$ = "##,###.##"
  DistSumLine$ = "----------------"
  TransTotal# = 0
  TranCnt = 0

  Dim Vendor As VendorRecType

  ReDim TAPLedgerRec(1) As APLedger81RecType
  LedgerRecLen = Len(TAPLedgerRec(1))

'    If InStr(UCase$(User$), "LUNENBURG") > 0 Then
'      LuneyFlag = True
'    End If
    FileName$ = "APINVREG.PRN"
    RegTitle$ = "A/P Purchases Journal"
    TranCol = CreditCol
    'CashCol = CreditCol
  FileSub$ = "APINVSub.PRN"
  ReDim Title(1 To 7) As String

  OpenAPEditFile APEditFile, NumEdTrans
  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumLdgTran&, LedgerRecLen

  PrnFileNum = FreeFile
  Open FileName$ For Output As #PrnFileNum
  PrnFileNum2 = FreeFile
  Open FileSub$ For Output As #PrnFileNum2
  '--Get a list of active funds
  ReDim FundList(1) As String
  GetFundList FundList$(), NumFunds
  'REDIM FundSum#(1 TO NumFunds)
  ReDim FundGrdTot#(1 To NumFunds)

 ' GoSub PrintHeader
  If NumEdTrans < 1 Then
    Close
    MsgBox "No Invoices to Print.", vbOKOnly, "No Trans"
    Exit Sub
  End If
  If LuneyFlag Then
    GoSub MakeLuneyBurgSort
  End If
  For Transaction = 1 To NumEdTrans
    If LuneyFlag Then
      Get APEditFile, LuneyIdx(Transaction).TRRec, APIED
    Else
      Get APEditFile, Transaction, APIED
    End If
    Get VendorFile, APIED.VRecNum, Vendor
   
    If Not APIED.DelFlag Then
'      If APEdit.POFLAG = -32767 Then
'        Print #PRNFileNum, "************* Error: Corrupt PO Flag *************"
'        Print #PRNFileNum, "DO NOT POST!!  Contact Customer Support"
'        Exit For
'      End If

      TranCnt = TranCnt + 1
      TransTotal# = Round#(TransTotal# + APIED.InvAmt)

      '--Print 1st Line - Transaction details
      ToPrintI$ = ""

      ToPrintI$ = QPTrim(Vendor.vnum) + "~" + QPTrim(APIED.VendName)
      ToPrintI$ = ToPrintI$ + "~" + QPTrim(APIED.INVDESC)
      ToPrintI$ = ToPrintI$ + "~" + APIED.TAXYN

      If APIED.PSLFlag = "Y" Then
        ToPrintI$ = ToPrintI$ + "~" + "Y"
      ElseIf APIED.PSLFlag = "N" Then
        ToPrintI$ = ToPrintI$ + "~" + "N"
      Else
        ToPrintI$ = ToPrintI$ + "~ "
      End If

      If APIED.Get1099 = "Y" Then
        ToPrintI$ = ToPrintI$ + "~" + "Y"
      ElseIf APIED.Get1099 = "N" Then
        ToPrintI$ = ToPrintI$ + "~" + "N"
      Else
        ToPrintI$ = ToPrintI$ + "~ "
      End If

      ToPrintI$ = ToPrintI$ + "~" + QPTrim(APIED.InvNum)
      If APIED.POLINES > 0 Then
        If APIED.POUSED <> APIED.POLINES Then
          ToPrintI$ = ToPrintI$ + "~" + "Partial"
        Else
          ToPrintI$ = ToPrintI$ + "~" + "Complete"
        End If
      Else
        ToPrintI$ = ToPrintI$ + "~" + "     "
      End If
      If Len(QPTrim$(APIED.PONum)) > 0 Then
        ToPrintI$ = ToPrintI$ + "~" + QPTrim$(APIED.PONum)
      Else
        ToPrintI$ = ToPrintI$ + "~" + QPTrim$(APIED.MPONum)
      End If
      ToPrintI$ = ToPrintI$ + "~" + Format(DateAdd("d", APIED.InvDate, "12-31-1979"), "mm/dd/yy")
      ToPrintI$ = ToPrintI$ + "~" + Format(DateAdd("d", APIED.DueDate, "12-31-1979"), "mm/dd/yy")
      ToPrintI$ = ToPrintI$ + "~" + Format(DateAdd("d", APIED.DISTDATE, "12-31-1979"), "mm/dd/yy")
      ToPrintI$ = ToPrintI$ + "~" + Using$(CommaFmt$, Str$(APIED.InvAmt))
'      Print #PrnFileNum, ToPrint$
'      If InStr(APEdit.PONUM, "Multi") > 0 Then
'        For zz = 1 To 6
         ' TNum$ = APEdit.PORecs(zz)
'          If APIED.POAPLRecNum > 0 Then
'
'            Get APLedgerFile, APIED.POAPLRecNum, TAPLedgerRec(1)
'
'           ToPrint$ = Space$(80)

'            Mid$(ToPrint$, 20) = Left$(TAPLedgerRec(1).PONUM, 9)
'            Mid$(ToPrint$, 66) = Using$(CommaFmt$, Str$(TAPLedgerRec(1).Amt))
'            Mid$(ToPrint$, 50) = Format(DateAdd("d", TAPLedgerRec(1).TRDATE, "12-31-1979"), "mm/dd/yy")
'            Print #PRNFileNum, ToPrint$
'          End If
'        Next
'      End If
'      ToPrint$ = Space$(80)

      TempTax# = Round#(APIED.STAXAMT + APIED.CTAXAMT)

      If TempTax# > 0 Then
        TaxTotal# = Round#(TaxTotal# + TempTax#)
      End If
      'Mid$(ToPrint$, 62) = "Tax:"
      ToPrintI$ = ToPrintI$ + "~" + Using$(CommaFmt$, Str$(TempTax#))

'--Print Distribution Label

      '--Print Accounting Distributions
      TotTranDist# = 0

      '--Loop Thru distributions to print and summarize
      For AcctDist = 1 To 36
        '--no more distributions when we find a blank Acct Number field
        If Len(QPTrim$(APIED.Dist(AcctDist).DACN)) > 0 Then
          '--Add distribution to total
          TotTranDist# = Round#(TotTranDist# + APIED.Dist(AcctDist).DAMT)
          '--Add distribution to proper fund
          Found = False
          For Fund = 1 To NumFunds
            FundNum$ = Left$(APIED.Dist(AcctDist).DACN, GLFundLen)
            If FundNum$ = FundList$(Fund) Then
              Found = True
              FundGrdTot#(Fund) = Round#(FundGrdTot#(Fund) + APIED.Dist(AcctDist).DAMT)
              Exit For
            End If
          Next
          If Not Found Then
            MsgBox "Invalid Fund - " + FundNum$, vbOKOnly, "Error"
            Close
            Exit Sub
          End If
          '--Print this distribution
          ToPrintD$ = ""
          ToPrintD$ = QPTrim(APIED.Dist(AcctDist).DACN)
          ToPrintD$ = ToPrintD$ + "~" + QPTrim(APIED.Dist(AcctDist).DACNM)
          If APIED.Dist(AcctDist).DACODE = "T" Then
            ToPrintD$ = ToPrintD$ + "~" + "PO Dist"
          Else
            ToPrintD$ = ToPrintD$ + "~" + "       "
          End If
          ToPrintD$ = ToPrintD$ + "~" + Using$(CommaFmt$, Str$(APIED.Dist(AcctDist).DAMT))
          ToPrint$ = ToPrintI$ + "~" + ToPrintD$ + "~" + QPTrim(APIED.Vendor) + QPTrim(APIED.InvNum) + Str(TranCnt)
          Print #PrnFileNum, ToPrint$
        End If  'Active transaction test
      Next      'Distribution
      '--Summary line after last distribution
      ToPrint$ = ""

      '--Transaction Distribution Totals
      '--2 blank lines before next distribution
    End If      'Not deleted test
  Next          'Transaction

  '--Summary
  For cnt = 1 To NumFunds
    If FundGrdTot#(cnt) <> 0 Then
      ToPrint$ = ""
      ToPrint$ = FundList$(cnt)
      ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(FundGrdTot#(cnt)))
      Print #PrnFileNum2, ToPrint$
      GrdTot# = Round#(GrdTot# + FundGrdTot#(cnt))
    End If
  Next
  Close
  Load frmLoadingRpt
  ARptInvEdit.totTrans = Using$("####", Str$(TranCnt))
  ARptInvEdit.totDists = Using$(CommaFmt$, Str$(TransTotal#))
  ARptInvEdit.totTaxes = Using$(CommaFmt$, Str$(TaxTotal#))
  ARptInvEdit.totFunds = Using$(CommaFmt$, Str$(GrdTot#))
  ARptInvEdit.txtTown.Caption = GLUserName$
  ARptInvEdit.txtDate.Caption = Now
  ARptInvEdit.Label1.Caption = RegTitle$
  ARptInvEdit.GetName FileName$, FileSub$
  ARptInvEdit.startrpt

  Exit Sub
MakeLuneyBurgSort:
  ReDim LuneyIdx(1 To NumEdTrans) As LuneySortType
  For cnt = 1 To NumEdTrans
    Get APEditFile, cnt, APIED
    'If Not APIED.DelFlag Then
      LuneyIdx(cnt).Vendor = APIED.Vendor
      LuneyIdx(cnt).TRRec = cnt
    'End If
  Next
  low = LBound(LuneyIdx)
  High = UBound(LuneyIdx)
  LuneySort LuneyIdx(), low, High
Return

End Sub
Private Sub PrnEditList2(LuneyFlag As Boolean)
  Dim FF As String, MaxLines As Integer, CommaFmt As String
  Dim Comma2Fmt As String, DistSumLine As String, FileName As String
  Dim DebitCol As Integer, CreditCol As Integer, RegTitle As String
  Dim TransTotal As Double, TranCnt As Integer, TranCol As Integer
  Dim PrnFileNum As Integer, NumFunds As Integer, Transaction As Integer
  Dim APEditFile As Integer, NumEdTrans As Integer, Linecnt As Integer
  Dim VendorFile As Integer, NumVRecs As Integer, TempTax As Double
  Dim LedgerRecLen As Integer, ToPrint As String, TaxTotal As Double
  Dim APLedgerFile As Integer, NumLdgTran As Long, TotTranDist As Double
  Dim AcctDist As Integer, Found As Boolean, Fund As Integer, FundNum As String
  Dim cnt As Integer, GrdTot As Double, HCnt As Integer, PageNum As Integer
  Dim APIED As APInv85Type
  ReDim LuneyIdx(1 To 1) As LuneySortType
  
  FF$ = Chr$(12)

  MaxLines = 51

  DebitCol = 42
  CreditCol = 58
  CommaFmt$ = "###,###,###.##"   'ten millions
  Comma2Fmt$ = "##,###.##"
  DistSumLine$ = "----------------"
  TransTotal# = 0
  TranCnt = 0

  Dim Vendor As VendorRecType

  ReDim TAPLedgerRec(1) As APLedger81RecType
  LedgerRecLen = Len(TAPLedgerRec(1))

'    If InStr(UCase$(User$), "LUNENBURG") > 0 Then
'      LuneyFlag = True
'    End If
    FileName$ = "APINVREG.PRN"
    RegTitle$ = "A/P Purchases Journal"
    TranCol = CreditCol
    'CashCol = CreditCol

  ReDim Title(1 To 7) As String

  OpenAPEditFile APEditFile, NumEdTrans
  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumLdgTran&, LedgerRecLen

  PrnFileNum = FreeFile
  Open FileName$ For Output As #PrnFileNum

  '--Get a list of active funds
  ReDim FundList(1) As String
  GetFundList FundList$(), NumFunds
  'REDIM FundSum#(1 TO NumFunds)
  ReDim FundGrdTot#(1 To NumFunds)

  GoSub PrintHeader
  If NumEdTrans < 1 Then
    Close
    MsgBox "No Invoices to Print.", vbOKOnly, "No Trans"
    Exit Sub
  End If
  If LuneyFlag Then
    GoSub MakeLuneyBurgSort
  End If
  For Transaction = 1 To NumEdTrans
    If LuneyFlag Then
      Get APEditFile, LuneyIdx(Transaction).TRRec, APIED
    Else
      Get APEditFile, Transaction, APIED
    End If
    Get VendorFile, APIED.VRecNum, Vendor
   
    If Not APIED.DelFlag Then
'      If APEdit.POFLAG = -32767 Then
'        Print #PRNFileNum, "************* Error: Corrupt PO Flag *************"
'        Print #PRNFileNum, "DO NOT POST!!  Contact Customer Support"
'        Exit For
'      End If

      TranCnt = TranCnt + 1
      TransTotal# = Round#(TransTotal# + APIED.InvAmt)

      '--Print 1st Line - Transaction details
      ToPrint$ = Space$(80)

      LSet ToPrint$ = Vendor.vnum
      Mid$(ToPrint$, 12) = APIED.VendName
      Mid$(ToPrint$, 39) = APIED.INVDESC
      Mid$(ToPrint$, 68) = APIED.TAXYN

      If APIED.PSLFlag = "Y" Then
        Mid$(ToPrint$, 73) = "Y"
      ElseIf APIED.PSLFlag = "N" Then
        Mid$(ToPrint$, 73) = "N"
      End If

      If APIED.Get1099 = "Y" Then
        Mid$(ToPrint$, 78) = "Y"
      ElseIf APIED.Get1099 = "N" Then
        Mid$(ToPrint$, 78) = "N"
      End If

      Print #PrnFileNum, ToPrint$

      ToPrint$ = Space$(80)
      Mid$(ToPrint$, 1) = APIED.InvNum
      If APIED.POLINES > 0 Then
        If APIED.POUSED <> APIED.POLINES Then
          Mid$(ToPrint$, 10) = "Partial"
        Else
          Mid$(ToPrint$, 10) = "Complete"
        End If
      End If
      If Len(APIED.PONum) > 0 Then
        Mid$(ToPrint$, 20) = Left$(APIED.PONum, 9)
      Else
        Mid$(ToPrint$, 20) = Left$(APIED.MPONum, 9)
      End If
      Mid$(ToPrint$, 30) = Format(DateAdd("d", APIED.InvDate, "12-31-1979"), "mm/dd/yy")
      Mid$(ToPrint$, 40) = Format(DateAdd("d", APIED.DueDate, "12-31-1979"), "mm/dd/yy")
      Mid$(ToPrint$, 50) = Format(DateAdd("d", APIED.DISTDATE, "12-31-1979"), "mm/dd/yy")
      Mid$(ToPrint$, 66) = Using$(CommaFmt$, Str$(APIED.InvAmt))
      Print #PrnFileNum, ToPrint$
'      If InStr(APEdit.PONUM, "Multi") > 0 Then
'        For zz = 1 To 6
         ' TNum$ = APEdit.PORecs(zz)
'          If APIED.POAPLRecNum > 0 Then
'
'            Get APLedgerFile, APIED.POAPLRecNum, TAPLedgerRec(1)
'
'           ToPrint$ = Space$(80)

'            Mid$(ToPrint$, 20) = Left$(TAPLedgerRec(1).PONUM, 9)
'            Mid$(ToPrint$, 66) = Using$(CommaFmt$, Str$(TAPLedgerRec(1).Amt))
'            Mid$(ToPrint$, 50) = Format(DateAdd("d", TAPLedgerRec(1).TRDATE, "12-31-1979"), "mm/dd/yy")
'            Print #PRNFileNum, ToPrint$
'          End If
'        Next
'      End If
      ToPrint$ = Space$(80)

      TempTax# = Round#(APIED.STAXAMT + APIED.CTAXAMT)

      If TempTax# > 0 Then
        TaxTotal# = Round#(TaxTotal# + TempTax#)
      End If
      Mid$(ToPrint$, 62) = "Tax:"
      Mid$(ToPrint$, 66) = Using$(CommaFmt$, Str$(TempTax#))
      Print #PrnFileNum, ToPrint$

      '--Blank line between detail and acct'g distributions
      Print #PrnFileNum,
      Linecnt = Linecnt + 4
      
      If Linecnt >= MaxLines Then
        Print #PrnFileNum, FF$
        GoSub PrintHeader
      End If
'--Print Distribution Label
      ToPrint$ = Space$(80)
      LSet ToPrint$ = " Accounting Distribution:"
      Print #PrnFileNum, ToPrint$

      '--Print Field Titles
      ToPrint$ = Space$(80)
      Mid$(ToPrint$, 4) = "Account Number   Name"
      Print #PrnFileNum, ToPrint$
      Linecnt = Linecnt + 2

      '--Print Accounting Distributions
      TotTranDist# = 0

      '--Loop Thru distributions to print and summarize
      For AcctDist = 1 To 36
        '--no more distributions when we find a blank Acct Number field
        If Len(QPTrim$(APIED.Dist(AcctDist).DACN)) > 0 Then
          '--Add distribution to total
          TotTranDist# = Round#(TotTranDist# + APIED.Dist(AcctDist).DAMT)
          '--Add distribution to proper fund
          Found = False
          For Fund = 1 To NumFunds
            FundNum$ = Left$(APIED.Dist(AcctDist).DACN, GLFundLen)
            If FundNum$ = FundList$(Fund) Then
              Found = True
              FundGrdTot#(Fund) = Round#(FundGrdTot#(Fund) + APIED.Dist(AcctDist).DAMT)
              Exit For
            End If
          Next
          If Not Found Then
            MsgBox "Invalid Fund - " + FundNum$, vbOKOnly, "Error"
            Close
            Exit Sub
          End If
          '--Print this distribution
          ToPrint$ = Space$(80)
          Mid$(ToPrint$, 4) = APIED.Dist(AcctDist).DACN
          Mid$(ToPrint$, 21) = APIED.Dist(AcctDist).DACNM
          If APIED.Dist(AcctDist).DACODE = "T" Then
            Mid$(ToPrint$, 45) = "PO Dist"
          End If
          Mid$(ToPrint$, TranCol) = Using$(CommaFmt$, Str$(APIED.Dist(AcctDist).DAMT))
          Print #PrnFileNum, ToPrint$
          Linecnt = Linecnt + 1
          If Linecnt >= MaxLines Then
            Print #PrnFileNum, FF$
            GoSub PrintHeader
          End If
        End If  'Active transaction test
      Next      'Distribution
      '--Summary line after last distribution
      ToPrint$ = Space$(78)
      Mid$(ToPrint$, TranCol) = DistSumLine$
      Print #PrnFileNum, ToPrint$

      '--Transaction Distribution Totals
      ToPrint$ = Space$(80)
      Mid$(ToPrint$, 4) = "Total Distributed"
      Mid$(ToPrint$, TranCol) = Using$(CommaFmt$, Str$(TotTranDist#))
      Print #PrnFileNum, ToPrint$
      ToPrint$ = String$(80, "-")
      Print #PrnFileNum, ToPrint$
      Linecnt = Linecnt + 3
      If Linecnt < MaxLines Then
        Print #PrnFileNum,
        Linecnt = Linecnt + 1
      Else
        Print #PrnFileNum, FF$
        GoSub PrintHeader
      End If
      '--2 blank lines before next distribution
    End If      'Not deleted test
  Next          'Transaction

'  If Linecnt > 45 Then
'    Print #PrnFileNum, FF$
'  End If
  If Linecnt >= MaxLines Then
    Print #PrnFileNum, FF$
    GoSub PrintHeader
  End If

  '--Summary
  ToPrint$ = Space$(80)
  LSet ToPrint$ = "File Totals:"
  Print #PrnFileNum, ToPrint$

  ToPrint$ = Space$(80)
  LSet ToPrint$ = "Number of Transactions"
  Mid$(ToPrint$, 31) = Using$("####", Str$(TranCnt))
  Print #PrnFileNum, ToPrint$

  ToPrint$ = Space$(80)
  LSet ToPrint$ = "Distrubtions:"
  Mid$(ToPrint$, 25) = Using$(CommaFmt$, Str$(TransTotal#))
  Print #PrnFileNum, ToPrint$
  LSet ToPrint$ = "        Taxes:"
  Mid$(ToPrint$, 25) = Using$(CommaFmt$, Str$(TaxTotal#))
  Print #PrnFileNum, ToPrint$

  Print #PrnFileNum,

  ToPrint$ = Space$(80)
  LSet ToPrint$ = "Summary by Fund:"
  Print #PrnFileNum, ToPrint$

  For cnt = 1 To NumFunds
    If FundGrdTot#(cnt) <> 0 Then
      ToPrint$ = Space$(80)
      LSet ToPrint$ = "Fund" + " " + FundList$(cnt)
      Mid$(ToPrint$, 25) = Using$(CommaFmt$, Str$(FundGrdTot#(cnt)))
      Print #PrnFileNum, ToPrint$
      GrdTot# = Round#(GrdTot# + FundGrdTot#(cnt))
    End If
  Next

  ToPrint$ = Space$(80)
  LSet ToPrint$ = "Total All Funds"
  Mid$(ToPrint$, 25) = Using$(CommaFmt$, Str$(GrdTot#))
  Print #PrnFileNum, ToPrint$
  Print #PrnFileNum, FF$

  Close

  ViewPrint FileName$, RegTitle$
  KillFile FileName$

  Exit Sub


PrintHeader:
PageNum = PageNum + 1
  Title$(1) = GLUserName$
  Title$(2) = RegTitle$
  Title$(3) = "Run Date: " + Date$
  Title$(4) = "                                                                Page: " + Str(PageNum)
  Title$(5) = "Vendor Code & Name                    Comment"
  Title$(6) = "Invoice             PO       Date      Due Date  Post Date       Tax  PSL  1099"
  Title$(7) = String$(80, "=")

  For HCnt = 1 To 7
    Print #PrnFileNum, Title$(HCnt)
  Next
  Linecnt = 7
Return

MakeLuneyBurgSort:
  ReDim LuneyIdx(1 To NumEdTrans) As LuneySortType
  For cnt = 1 To NumEdTrans
    Get APEditFile, cnt, APIED
    'If Not APIED.DelFlag Then
      LuneyIdx(cnt).Vendor = APIED.Vendor
      LuneyIdx(cnt).TRRec = cnt
    'End If
  Next
  low = LBound(LuneyIdx)
  High = UBound(LuneyIdx)
  LuneySort LuneyIdx(), low, High
Return

End Sub

Private Sub LuneySort(Idxbuff() As LuneySortType, lLBound, lUBound)
  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
  Dim Temp As LuneySortType
  Dim Temp2 As LuneySortType
  lngCurLow = lLBound
  lngCurHigh = lUBound
  'this is to exit loop if high and low are equal
  If lUBound <= lLBound Then Exit Sub 'GoTo Proc_Exit
    'find the midpoint
    lngCurMid = (lUBound + lLBound) \ 2
    '
    Temp = Idxbuff(lngCurMid)
    Do While (lngCurLow <= lngCurHigh)
      Do While Idxbuff(lngCurLow).Vendor < Temp.Vendor
        lngCurLow = lngCurLow + 1
        If lngCurLow = lUBound Then Exit Do
      Loop
      Do While Temp.Vendor < Idxbuff(lngCurHigh).Vendor
        lngCurHigh = lngCurHigh - 1
        If lngCurHigh = lLBound Then Exit Do
      Loop
      'if low is <= high then swap
      If (lngCurLow <= lngCurHigh) Then
        Temp2 = Idxbuff(lngCurLow)
        Idxbuff(lngCurLow) = Idxbuff(lngCurHigh)
        Idxbuff(lngCurHigh) = Temp2
    '
        lngCurLow = lngCurLow + 1
        lngCurHigh = lngCurHigh - 1
      End If
    Loop
  'recurse if necessary
    If lLBound < lngCurHigh Then
      LuneySort Idxbuff(), lLBound, lngCurHigh
    End If
  'recurse if necessary
    If lngCurLow < lUBound Then
      LuneySort Idxbuff(), lngCurLow, lUBound
    End If
End Sub



