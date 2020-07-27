VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmVATaxPayMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Payment Menu"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxPayMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdPost 
      Height          =   444
      Left            =   4080
      TabIndex        =   3
      Top             =   5016
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   783
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
      ButtonDesigner  =   "frmVATaxPayMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintTrans 
      Height          =   432
      Left            =   4080
      TabIndex        =   1
      Top             =   3912
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   762
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
      ButtonDesigner  =   "frmVATaxPayMenu.frx":0AAF
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDelete 
      Height          =   420
      Left            =   4080
      TabIndex        =   2
      Top             =   4470
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmVATaxPayMenu.frx":0C9E
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEnterPay 
      Height          =   432
      Left            =   4080
      TabIndex        =   0
      Top             =   3360
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   762
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
      ButtonDesigner  =   "frmVATaxPayMenu.frx":0E84
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   420
      Left            =   4080
      TabIndex        =   5
      Top             =   6150
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmVATaxPayMenu.frx":106B
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdRefund 
      Height          =   444
      Left            =   4080
      TabIndex        =   4
      Top             =   5580
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   783
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
      ButtonDesigner  =   "frmVATaxPayMenu.frx":1248
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   8576
      Top             =   2017
      Width           =   971
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TAX PAYMENT MENU"
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
      TabIndex        =   6
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
      Top             =   803
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
Attribute VB_Name = "frmVATaxPayMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class

Private Sub cmdDelete_Click()
  Dim Operator$
'  OPERNUM = 1
  Operator$ = CStr(OperNum)
  
  If Exist("ppayposterror" + Operator$ + ".dat") = True Then  '11/30/06
    Call TaxMsg(750, "A personal payment posting error took place the last time operator #" + CStr(OperNum) + " tried to post. No new payments can be made until this is resolved. Please call Southern Software @ 1-800-842-8190 for assistance.")
    Exit Sub
  End If
  
  If Not Exist("TAXRCPR" + Operator$ + ".DAT") And Not Exist("TAXPCPR" + Operator$ + ".DAT") Then
    frmVATaxMsg.Label1.Caption = "There are no payment records saved for operator #" + Operator + ". Delete attempt aborted."
    frmVATaxMsg.Label1.Top = 800
    frmVATaxMsg.Show vbModal
    Exit Sub
  End If
  frmVATaxDeletePayment.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdEnterPay_Click()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim ThisTYear As Integer
  
  If Exist("ppayposterror" + CStr(OperNum) + ".dat") = True Then '11/30/06
    Call TaxMsg(750, "A personal payment posting error took place the last time operator #" + CStr(OperNum) + " tried to post. No new payments can be made until this is resolved. Please call Southern Software @ 1-800-842-8190 for assistance.")
    Exit Sub
  End If
  
  If Exist("rpayposterror" + CStr(OperNum) + ".dat") = True Then '11/30/06
    Call TaxMsg(750, "A real payment posting error took place the last time operator #" + CStr(OperNum) + " tried to post. No new payments can be made until this is resolved. Please call Southern Software @ 1-800-842-8190 for assistance.")
    Exit Sub
  End If
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  ThisTYear = 0
  frmVATaxBillPostOpt.Show vbModal
  If frmVATaxBillPostOpt.fptxtPostType.Text = "Real" Then
    If TaxMasterRec.DisRPct > 0 Then
      If CheckTaxYear("R", ThisTYear) = False Then
        If TaxMsgWOpts(400, "The current real system tax year (" + CStr(TaxMasterRec.RTaxYear) + ") comes before some of the tax years for real tax bills posted (" + CStr(ThisTYear) + "). This can cause payment posting problems if discounts are allowed. If you wish to change the real system tax year then press ESC to escape and go to the System Setup screen to edit. Otherwise press F10 to continue as is.", "F10 Continue", "ESC Escape") = "abort" Then
          Unload frmVATaxMsgWOpts
          Exit Sub
        Else
          Unload frmVATaxMsgWOpts
          MainLog ("WARNING: User issued a warning that the current real system tax year (" + CStr(TaxMasterRec.RTaxYear) + ") comes before some of the tax years for tax bills posted (" + CStr(ThisTYear) + "). This can cause payment posting problems. User elected to continue anyway.")
        End If
      End If
      If RevsAndGLsOK(Me, TaxMasterRec.RTaxYear, "R") = False Then
        Exit Sub
      End If
    End If
    frmVATaxPaymentEntry.Show
    DoEvents
    Unload frmVATaxBillPostOpt
  ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Personal" Then
    If TaxMasterRec.DisPPct > 0 Then
      If CheckTaxYear("P", ThisTYear) = False Then
        If TaxMsgWOpts(400, "The current personal system tax year (" + CStr(TaxMasterRec.PTaxYear) + ") comes before some of the tax years for personal tax bills posted (" + CStr(ThisTYear) + "). This can cause payment posting problems if discounts are allowed. If you wish to change the personal system tax year then press ESC to escape and go to the System Setup screen to edit. Otherwise press F10 to continue as is.", "F10 Continue", "ESC Escape") = "abort" Then
          Unload frmVATaxMsgWOpts
          Exit Sub
        Else
          Unload frmVATaxMsgWOpts
          MainLog ("WARNING: User issued a warning that the current personal system tax year (" + CStr(TaxMasterRec.PTaxYear) + ") comes before some of the tax bills posted (" + CStr(ThisTYear) + "). This can cause payment posting problems. User elected to continue anyway.")
        End If
      End If
      If RevsAndGLsOK(Me, TaxMasterRec.PTaxYear, "P") = False Then
        Exit Sub
      End If
    End If
    frmVATaxPersPaymentEntry.Show
    DoEvents
    Unload frmVATaxBillPostOpt
  ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Exit" Then
    DoEvents
    Unload frmVATaxBillPostOpt
    Exit Sub
  End If
  
  Unload Me
End Sub

Private Sub cmdPost_Click()
  Dim Operator$
  
  If Exist("ppayposterror" + CStr(OperNum) + ".dat") = True Then '11/30/06
    Call TaxMsg(750, "A personal payment posting error took place the last time operator #" + CStr(OperNum) + " tried to post. No new payments can be made until this is resolved. Please call Southern Software @ 1-800-842-8190 for assistance.")
    Exit Sub
  End If
  
  If Exist("rpayposterror" + CStr(OperNum) + ".dat") = True Then '11/30/06
    Call TaxMsg(750, "A real payment posting error took place the last time operator #" + CStr(OperNum) + " tried to post. No new payments can be made until this is resolved. Please call Southern Software @ 1-800-842-8190 for assistance.")
    Exit Sub
  End If
  
  Operator$ = CStr(OperNum)
  If Not Exist("TAXRCPR" + Operator$ + ".DAT") And Not Exist("TAXPCPR" + Operator$ + ".DAT") Then
    frmVATaxMsg.Label1.Caption = "There are no payment records saved for operator #" + Operator + ". Posting attempt aborted."
    frmVATaxMsg.Label1.Top = 800
    frmVATaxMsg.Show vbModal
    Exit Sub
  End If
  
  If Not Exist("TAXRCPR" + Operator$ + ".DAT") And Exist("TAXPCPR" + Operator$ + ".DAT") Then
    frmVATaxPayPost.ThisBillType = "P"
    frmVATaxPayPost.Label8 = "Personal Tax Payment Post"
    GoTo JumpHere
  ElseIf Exist("TAXRCPR" + Operator$ + ".DAT") And Not Exist("TAXPCPR" + Operator$ + ".DAT") Then
    frmVATaxPayPost.ThisBillType = "R"
    frmVATaxPayPost.Label8 = "Real Tax Payment Post"
    GoTo JumpHere
  End If
  
  frmVATaxPayPostOpt.Show vbModal
  If frmVATaxPayPostOpt.fptxtPostType.Text = "Personal" Then
    frmVATaxPayPost.ThisBillType = "P"
    Unload frmVATaxPayPostOpt
  ElseIf frmVATaxPayPostOpt.fptxtPostType.Text = "Real" Then
    frmVATaxPayPost.ThisBillType = "R"
    Unload frmVATaxPayPostOpt
  ElseIf frmVATaxPayPostOpt.fptxtPostType.Text = "Both" Then
    frmVATaxPayPost.ThisBillType = "B"
    Unload frmVATaxPayPostOpt
  Else
    Unload frmVATaxPayPostOpt
    Exit Sub
  End If
  
JumpHere:
  frmVATaxPayPost.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPrintTrans_Click()
  Dim Operator$
  
  If Exist("ppayposterror" + CStr(OperNum) + ".dat") = True Then '11/30/06
    Call TaxMsg(750, "A personal payment posting error took place the last time operator #" + CStr(OperNum) + " tried to post. No new payments can be made until this is resolved. Please call Southern Software @ 1-800-842-8190 for assistance.")
    Exit Sub
  End If
  
  If Exist("rpayposterror" + CStr(OperNum) + ".dat") = True Then '11/30/06
    Call TaxMsg(750, "A real payment posting error took place the last time operator #" + CStr(OperNum) + " tried to post. No new payments can be made until this is resolved. Please call Southern Software @ 1-800-842-8190 for assistance.")
    Exit Sub
  End If
  
  Operator$ = CStr(OperNum)
  If Not Exist("TAXRCPR" + Operator$ + ".DAT") And Not Exist("TAXPCPR" + Operator$ + ".DAT") Then
    frmVATaxMsg.Label1.Caption = "There are no payment records saved for operator #" + Operator + ". Load attempt aborted."
    frmVATaxMsg.Label1.Top = 800
    frmVATaxMsg.Show vbModal
    Exit Sub
  End If
  frmVATaxPayEditList.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdRefund_Click()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim PreBal As Double
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  
  On Error GoTo ERRORSTUFF
  
  If Exist("ppayposterror" + CStr(OperNum) + ".dat") = True Then '11/30/06
    Call TaxMsg(750, "A personal payment posting error took place the last time operator #" + CStr(OperNum) + " tried to post. No new payments can be made until this is resolved. Please call Southern Software @ 1-800-842-8190 for assistance.")
    Exit Sub
  End If
  
  If Exist("rpayposterror" + CStr(OperNum) + ".dat") = True Then '11/30/06
    Call TaxMsg(750, "A real payment posting error took place the last time operator #" + CStr(OperNum) + " tried to post. No new payments can be made until this is resolved. Please call Southern Software @ 1-800-842-8190 for assistance.")
    Exit Sub
  End If
  
  RefNumCnt = 0
  RefNameCnt = 0
  frmVATaxShowPctComp.Label1 = "Searching For Prepaid Balances"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  OpenTaxCustFile TCHandle, NumOfTCRecs
  cmdDelete.Enabled = False
  cmdEnterPay.Enabled = False
  cmdExit.Enabled = False
  cmdPost.Enabled = False
  cmdPrintTrans.Enabled = False
  cmdRefund.Enabled = False
  
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TaxCust.Deleted <> 0 Then GoTo NextOne
    PreBal = GetCustBalance(x, -1)
    If PreBal < 0 Then
      RefNumCnt = RefNumCnt + 1
      ReDim Preserve AcctNumList(1 To RefNumCnt) As Long
      AcctNumList(RefNumCnt) = x
      ReDim Preserve NumPreBal(1 To RefNumCnt) As Double
      NumPreBal(RefNumCnt) = PreBal
    End If
NextOne:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdDelete.Enabled = True
      cmdEnterPay.Enabled = True
      cmdExit.Enabled = True
      cmdPost.Enabled = True
      cmdPrintTrans.Enabled = True
      cmdRefund.Enabled = True
      Exit Sub
    End If
  Next x
  frmVATaxShowPctComp.Label1 = "Searching For Prepaid Balances"
  frmVATaxShowPctComp.Show , Me
  
  OpenNameIdxFile IdxHandle, NumOfIdx
  For x = 1 To NumOfIdx
    Get IdxHandle, x, IdxRec
    Get TCHandle, IdxRec.CustRec, TaxCust
    If TaxCust.Deleted <> 0 Then GoTo NotThisOne
    PreBal = GetCustBalance(IdxRec.CustRec, -1)
    If PreBal < 0 Then
      RefNameCnt = RefNameCnt + 1
      ReDim Preserve AcctNameList(1 To RefNameCnt) As Long
      AcctNameList(RefNameCnt) = IdxRec.CustRec
      ReDim Preserve NamePreBal(1 To RefNameCnt) As Double
      NamePreBal(RefNameCnt) = PreBal
    End If
NotThisOne:
    frmVATaxShowPctComp.ShowPctComp x, NumOfIdx
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdDelete.Enabled = True
      cmdEnterPay.Enabled = True
      cmdExit.Enabled = True
      cmdPost.Enabled = True
      cmdPrintTrans.Enabled = True
      cmdRefund.Enabled = True
      Exit Sub
    End If
  Next x
  
  Close
  
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdDelete.Enabled = True
  cmdEnterPay.Enabled = True
  cmdExit.Enabled = True
  cmdPost.Enabled = True
  cmdPrintTrans.Enabled = True
  cmdRefund.Enabled = True
  
  If RefNumCnt = 0 Then
    Call TaxMsg(900, "There are no customers with prepaid balances at this time.")
    Exit Sub
  End If
  
  frmVATaxRefundOnPrepay.Show
  DoEvents
  Unload Me
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPayMenu", "cmdRefund_Click", Erl)
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
    ClearInUse PWcnt
    Terminate

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%x"
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
  Me.HelpContextID = hlpTaxPaymentMenu
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxPayMenu.")
      Call Terminate
      End
    End If
  End If

End Sub
'Private Sub Form_Resize()
'  If Me.WindowState <> vbMinimized Then
'    Me.Visible = False
'    'Temp_Class.ResizeControls Me
'    Me.Visible = True
'    Me.SetFocus
'    DoEvents
'  End If
'End Sub

Private Sub cmdExit_Click()
  frmVATaxMainMenu.Show
  DoEvents
  Unload Me
End Sub

