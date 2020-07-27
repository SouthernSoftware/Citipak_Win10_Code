VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmVATaxAdvColMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advertising Menu"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxAdvColMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdPost 
      Height          =   435
      Left            =   4080
      TabIndex        =   4
      Top             =   5445
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmVATaxAdvColMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintTrans 
      Height          =   450
      Left            =   4080
      TabIndex        =   2
      Top             =   4155
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   794
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
      ButtonDesigner  =   "frmVATaxAdvColMenu.frx":0ABB
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEditTrans 
      Height          =   435
      Left            =   4080
      TabIndex        =   1
      Top             =   3510
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmVATaxAdvColMenu.frx":0CAC
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCalcChrgs 
      Height          =   435
      Left            =   4080
      TabIndex        =   0
      Top             =   2880
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmVATaxAdvColMenu.frx":0E9D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   435
      Left            =   4080
      TabIndex        =   6
      Top             =   6720
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmVATaxAdvColMenu.frx":108E
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMailLbls 
      Height          =   450
      Left            =   4080
      TabIndex        =   3
      Top             =   4800
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   794
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
      ButtonDesigner  =   "frmVATaxAdvColMenu.frx":126B
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClear 
      Height          =   435
      Left            =   4080
      TabIndex        =   5
      Top             =   6075
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmVATaxAdvColMenu.frx":1459
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TAX ADVERTISING MENU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2813
      TabIndex        =   7
      Top             =   1164
      Width           =   6012
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2214
      X2              =   2214
      Y1              =   2127
      Y2              =   8015
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2199
      X2              =   2914
      Y1              =   8020
      Y2              =   8020
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8706
      X2              =   9408
      Y1              =   8020
      Y2              =   8020
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   8586
      Top             =   2027
      Width           =   971
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8706
      X2              =   8706
      Y1              =   2127
      Y2              =   8028
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   2094
      Top             =   2019
      Width           =   971
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1098
      Index           =   1
      Left            =   1493
      Top             =   813
      Width           =   8655
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1214
      Left            =   1495
      Top             =   687
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   2094
      Top             =   1886
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5910
      Index           =   0
      Left            =   2213
      Top             =   2117
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8585
      Top             =   1887
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5910
      Index           =   1
      Left            =   8706
      Top             =   2117
      Width           =   732
   End
End
Attribute VB_Name = "frmVATaxAdvColMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class

Private Sub cmdCalcChrgs_Click()
  frmVATaxCalcAdCol.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdClear_Click()
  If Not Exist(TaxAdvFile) Then
    Call TaxMsg(900, "No advertising calc files currently exist. Delete attempt aborted.")
    Exit Sub
  End If
  
  If TaxMsgWOpts(600, "WARNING: IF YOU CHOOSE TO CONTINUE THEN ALL UNPOSTED ADVERTISING CALCULATION FILES WILL BE REMOVED PERMANENTLY. IF YOU WISH TO CONTINUE THEN PRESS F10. OTHERWISE PRESS ESC TO LEAVE UNPOSTED ADVERTISING CALCULATION FILES UNCHANGED.", "F10 Delete", "ESC Abort") = "abort" Then
    Exit Sub
  Else
    KillFile TaxAdvFile
    MainLog ("User deleted unposted advertising calculations files after being warned about the consequences.")
    Call Savemsg(900, "All unposted advertising calculations files have been deleted successfully.")
  End If

End Sub

Private Sub cmdEditTrans_Click()
  Dim AdvRec As InterestRecType
  Dim NumOfARRecs As Long
  Dim ARHandle As Integer
  Dim x As Long
  
  OpenAdvColRecFile ARHandle, NumOfARRecs
  
  If NumOfARRecs = 0 Then
    Call TaxMsg(900, "There are no advertising charges transactions saved.")
    Close ARHandle
    Exit Sub
  Else
    For x = 1 To NumOfARRecs
      Get ARHandle, x, AdvRec
      If AdvRec.DelFlag = False Then
        Exit For
      End If
    Next x
  End If
  If x > NumOfARRecs Then
    Call TaxMsg(900, "There are no advertising charges transactions saved.")
    Close ARHandle
    Exit Sub
  End If
  
  Close ARHandle
  frmVATaxEditAdv.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdExit_Click()
  frmVATaxMainMenu.Show
  DoEvents
  Unload Me
End Sub


Private Sub cmdMailLbls_Click()
  Dim AdvRec As InterestRecType
  Dim NumOfARRecs As Long
  Dim ARHandle As Integer
  Dim x As Long
  OpenAdvColRecFile ARHandle, NumOfARRecs

  If NumOfARRecs = 0 Then
    Call TaxMsg(900, "There are no advertising charges transactions saved.")
    Close ARHandle
    Exit Sub
  Else
    For x = 1 To NumOfARRecs
      Get ARHandle, x, AdvRec
      If AdvRec.DelFlag = False Then
        Exit For
      End If
    Next x
  End If

  If x > NumOfARRecs Then
    Call TaxMsg(900, "There are no advertising charges transactions saved.")
    Close ARHandle
    Exit Sub
  End If

  Close ARHandle

  frmVATaxMailLblsAdv.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPost_Click()
  Dim AdvRec As InterestRecType
  Dim NumOfARRecs As Long
  Dim ARHandle As Integer
  Dim x As Long
  
  OpenAdvColRecFile ARHandle, NumOfARRecs

  If NumOfARRecs = 0 Then
    Call TaxMsg(900, "There are no advertising charges transactions saved.")
    Close ARHandle
    Exit Sub
  Else
    For x = 1 To NumOfARRecs
      Get ARHandle, x, AdvRec
      If AdvRec.DelFlag = False Then
        Exit For
      End If
    Next x
  End If
  
  If x > NumOfARRecs Then
    Call TaxMsg(900, "There are no advertising charges transactions saved.")
    Close ARHandle
    Exit Sub
  End If

  Close ARHandle
  frmVATaxAdvPost.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPrintTrans_Click()
  Dim AdvRec As InterestRecType
  Dim NumOfARRecs As Long
  Dim ARHandle As Integer
  Dim x As Long
  
  OpenAdvColRecFile ARHandle, NumOfARRecs

  If NumOfARRecs = 0 Then
    Call TaxMsg(900, "There are no advertising charges transactions saved.")
    Close ARHandle
    Exit Sub
  Else
    For x = 1 To NumOfARRecs
      Get ARHandle, x, AdvRec
      If AdvRec.DelFlag = False Then
        Exit For
      End If
    Next x
  End If
  
  If x > NumOfARRecs Then
    Call TaxMsg(900, "There are no advertising charges transactions saved.")
    Close ARHandle
    Exit Sub
  End If

  Close ARHandle
  
  frmVATaxReportOpt.Show vbModal
  If frmVATaxReportOpt.fptxtPrintType.Text = "Graphical" Then
    Unload frmVATaxReportOpt
    Call PrintGraphics
  ElseIf frmVATaxReportOpt.fptxtPrintType.Text = "Text" Then
    frmVATaxMsg.Label1.Caption = "Pitch 10 is recommended for this report."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Unload frmVATaxReportOpt
    Call PrintText
  End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpTaxAdvertising
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxAdvColMenu.")
      Call Terminate
      End
    End If
  End If

End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    'Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub PrintGraphics()
  Dim AdvTrans As InterestRecType
  Dim ATHandle As Integer
  Dim NumOfATRecs As Long
  Dim dlm$
  Dim RptFile$
  Dim RptHandle As Integer
  Dim x As Long, y As Integer
  Dim TotAmt As Double
  Dim TCnt As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim Town$
  
  On Error GoTo ERRORSTUFF
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Town$ = QPTrim$(TaxMasterRec.Name)
  dlm$ = "~"
  RptFile$ = "TAXRPTS\TAXADVCOL.RPT"
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  
  OpenAdvColRecFile ATHandle, NumOfATRecs
  For x = 1 To NumOfATRecs
    Get ATHandle, x, AdvTrans
    '                   0                 1                          2
    Print #RptHandle, Town$; dlm; AdvTrans.TaxYear; dlm; AdvTrans.CustRec; dlm;
    If AdvTrans.DelFlag = True Then
      '                            3                            4                         5
      Print #RptHandle, QPTrim$(AdvTrans.CustName); dlm; AdvTrans.InfoTxt; dlm; "Deleted"; dlm;
    Else
      TotAmt = OldRound(TotAmt + AdvTrans.Amount)
      '                            3                            4                         5
      Print #RptHandle, QPTrim$(AdvTrans.CustName); dlm; AdvTrans.InfoTxt; dlm; AdvTrans.Amount; dlm;
    End If
    '                    6             7
    Print #RptHandle, TotAmt; dlm; NumOfATRecs
  Next x
  
  Close
  
  arVATaxAdvColRpt.Show
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxAdvColMenu", "PrintGraphics", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close


End Sub

'Private Sub PrintText()
'  Dim TaxMasterRec As TaxMasterType
'  Dim TMHandle As Integer
'  Dim AdvRec As InterestRecType
'  Dim NumOfARRecs As Long
'  Dim ARHandle As Integer
'  Dim x As Long, y As Integer
'  Dim Town$
'  Dim Page As Integer
'  Dim LineCnt As Integer
'  Dim MaxLines As Integer
'  Dim RptHandle As Integer
'  Dim RptFile$, FF$
'  Dim TotAdv As Double
'  Dim ThisYear As String
'  Dim TCnt As Long
'  Dim ThisInfo As String * 30
'  Dim ThisName As String * 35
'
'  MaxLines = 56
'  FF$ = Chr(12)
'  OpenTaxSetUpFile TMHandle
'  Get TMHandle, 1, TaxMasterRec
'  Close TMHandle
'
'  Town$ = QPTrim$(TaxMasterRec.Name)
'
'  RptFile$ = "TAXRPTS\TAXADV.PRN"     'Report File Name
'  RptHandle = FreeFile
'  Open RptFile$ For Output As #RptHandle
'
'  OpenAdvColRecFile ARHandle, NumOfARRecs
'  Get ARHandle, 1, AdvRec
'  ThisYear = CStr(AdvRec.TaxYear)
'  GoSub PrintHeader
'  For x = 1 To NumOfARRecs
'    Get ARHandle, x, AdvRec
'    ThisYear = CStr(AdvRec.TaxYear)
'    ThisInfo = QPTrim$(AdvRec.InfoTxt)
'    Print #RptHandle, Using$("####0", AdvRec.CustRec); Tab(8); QPTrim$(AdvRec.CustName);
'    Print #RptHandle, Tab(44); Using$("####", AdvRec.TaxYear); Tab(50); Using$("####0", AdvRec.BillNumber);
'    Print #RptHandle, Tab(57); ThisInfo; Tab(90); Using$("$###,##0.00", AdvRec.Amount)
'    TCnt = TCnt + 1
'    LineCnt = LineCnt + 1
'    If LineCnt >= MaxLines Then
'      Print #RptHandle, FF$
'      GoSub PrintHeader
'    End If
'    TotAdv = OldRound(TotAdv + AdvRec.Amount)
'  Next x
'
'  Print #RptHandle, FF$
'  Page = Page + 1
'  Print #RptHandle, Tab(15); "Property Tax Billing: Advertising Charges Register"
'  Print #RptHandle, "Town: "; Tab(8); Town$; Tab(70); "Page #: " + CStr(Page)
'  Print #RptHandle, "Date: " + CStr(Date)
'  Print #RptHandle, "Tax Year: " + ThisYear
'  Print #RptHandle, String(100, "-")
'  Print #RptHandle, Tab(2); "Total Transactions:     "; Tab(39); Using$("#####0", TCnt)
'  Print #RptHandle, Tab(2); "Total Advertising Charges: "; Tab(30); Using$("$###,###,##0.00", TotAdv)
'
'  Print #RptHandle, FF$
'  Close
'
'  ViewPrint RptFile, "Advertising Charges", True
'
'  Exit Sub
'
'PrintHeader:
'  Page = Page + 1
'  Print #RptHandle, Tab(15); "Property Tax Billing: Advertising Charges Register"
'  Print #RptHandle, "Town: "; Tab(8); Town$; Tab(70); "Page #: " + CStr(Page)
'  Print #RptHandle, "Date: " + CStr(Date)
'  Print #RptHandle, "Current Tax Year: " + ThisYear
'  Print #RptHandle, "Acct #"; Tab(8); "Customer Name"; Tab(43); "Tax Yr"; Tab(50); "Bill #"; Tab(58); "Map\Block\Lot\Notes"; Tab(92); "Ad-Charge"
'  Print #RptHandle, String(100, "-")
'  LineCnt = 6
'  Return
'
'End Sub
Private Sub PrintText()
  Dim Page As Integer
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim RptHandle As Integer
  Dim RptFile$, FF$
  Dim AdvTrans As InterestRecType
  Dim ATHandle As Integer
  Dim NumOfATRecs As Long
  Dim TotAmt As Double
  Dim TCnt As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim Town$
  Dim TaxYear$
  Dim x As Long
  Dim ThisName As String * 35
  Dim ThisInfo As String * 30
  
  On Error GoTo ERRORSTUFF
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  MaxLines = 56
  Town$ = QPTrim$(TaxMasterRec.Name)
'  TaxYear = CStr(fptxtCurrYear.Text)
  RptFile$ = "TAXRPTS\TAXADVCOL.RPT"
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  
  OpenAdvColRecFile ATHandle, NumOfATRecs
  GoSub PrintHeader
  
  For x = 1 To NumOfATRecs
    Get ATHandle, x, AdvTrans
    TaxYear = AdvTrans.TaxYear
    ThisName = QPTrim$(AdvTrans.CustName)
    ThisInfo = QPTrim$(AdvTrans.InfoTxt)
    If AdvTrans.DelFlag = True Then
      Print #RptHandle, Using$("####0", AdvTrans.CustRec); Tab(10); ThisName; Tab(45); ThisInfo; Tab(76); "   Deleted"
    Else
      TotAmt = OldRound(TotAmt + AdvTrans.Amount)
      Print #RptHandle, Using$("####0", AdvTrans.CustRec); Tab(10); ThisName; Tab(45); ThisInfo; Tab(76); Using$("$##,##0.00", AdvTrans.Amount)
    End If
    TCnt = TCnt + 1
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
  Next x
  If LineCnt >= MaxLines - 4 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  Print #RptHandle, String$(85, "-")
  Print #RptHandle, "Transaction Count: "; Tab(26); Using$("####0", TCnt)
  Print #RptHandle, "Total Charges:     "; Tab(20); Using$("$###,##0.00", TotAmt)
  Print #RptHandle, FF$
  Close
  
  ViewPrint RptFile, "Tax Advertising Charges Report", True
  
  Exit Sub
  
PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Tax Advertising Charges Report"
  Print #RptHandle, "Town: " + Town$; Tab(75); "Page # " + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, "Tax Year: " + TaxYear
  Print #RptHandle, "Cust Num"; Tab(10); "Current Owner Name"; Tab(50); "Map\Block\Lot\Notes"; Tab(79); "Ad-Cost"
  Print #RptHandle, String(85, "-")
  LineCnt = 6
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxAdvColMenu", "PrintText", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close


End Sub
