VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmVATaxAdvPost 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advertising Charges Post"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxAdvPost.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdPost 
      Height          =   495
      Left            =   6540
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1800
      _Version        =   131072
      _ExtentX        =   3175
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmVATaxAdvPost.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   3300
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1800
      _Version        =   131072
      _ExtentX        =   3175
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmVATaxAdvPost.frx":0AA6
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Press F10 To Post."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   2678
      TabIndex        =   2
      Top             =   4770
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Press ESC To Exit."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   5798
      TabIndex        =   1
      Top             =   4770
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   2055
      Left            =   2033
      Top             =   3450
      Width           =   7575
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Advertising Charges Post"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3826
      TabIndex        =   0
      Top             =   2145
      Width           =   4020
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   840
      Left            =   2296
      Top             =   1905
      Width           =   7020
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   990
      Left            =   2311
      Top             =   1785
      Width           =   7020
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   2055
      Left            =   2033
      TabIndex        =   3
      Top             =   3450
      Width           =   7575
   End
End
Attribute VB_Name = "frmVATaxAdvPost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  frmVATaxAdvColMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPost_Click()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxAdvRec As InterestRecType
  Dim ATHandle As Integer
  Dim NumOfATRecs As Long
  Dim NewTaxTrans As TaxTransactionType
  Dim cnt&, NextRecord&
  Dim Previous&
'  If InStr(TaxSetupRec(1).Name, "HAMLET") > 0 Then
'    HamFlag = True
'  End If
  On Error GoTo ERRORSTUFF
  
  If TaxMsgWOpts(900, "If you are sure you are ready to post then press F10 to continue. Otherwise, press ESC to abort the post attempt.", "F10 Continue", "ESC Abort") = "abort" Then
    Unload frmVATaxMsgWOpts
    Call TaxMsg(900, "Post attempt aborted.")
    Close
    Exit Sub
  Else
    Unload frmVATaxMsgWOpts
    MainLog ("Advertising charges posted.")
  End If

  OpenAdvColRecFile ATHandle, NumOfATRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  
  For cnt& = 1 To NumOfATRecs
    Get ATHandle, cnt&, TaxAdvRec
    If TaxAdvRec.DelFlag = 0 Then
      'Update the Bill transaction first
      Get TTHandle, TaxAdvRec.BillRec, TaxTrans 'get bill trans
'      If HamFlag Then
'        TaxTrans(1).Revenue.Interest = Round#(TaxTrans(1).Revenue.Interest + TaxAdvRec(1).Amount)
'      Else
        TaxTrans.Revenue.Collection = OldRound#(TaxTrans.Revenue.Collection + TaxAdvRec.Amount)
'      End If
      Put TTHandle, TaxAdvRec.BillRec, TaxTrans 'put it back
      'Now make a new clean transaction
      TaxTrans = NewTaxTrans
      TaxTrans.TransDate = Date2Num(Date$)
      TaxTrans.TaxYear = TaxAdvRec.TaxYear
      TaxTrans.TranType = 6       '4=Interest
      TaxTrans.BillType = "R"     'R=Real P=Personal Property C=Combined (NC/GA)
      TaxTrans.Amount = TaxAdvRec.Amount  'Total Transaction Amount
      TaxTrans.Description = "Col/Ad Charge Bill#" + QPTrim$(TaxAdvRec.BillNumber)
      TaxTrans.Revenue.Principle1 = 0
      TaxTrans.Revenue.Principle2 = 0
      TaxTrans.Revenue.Principle3 = 0
      TaxTrans.Revenue.Principle4 = 0
      TaxTrans.Revenue.Principle5 = 0
      TaxTrans.Revenue.LateList = 0
      TaxTrans.Revenue.RevOpt1 = 0
      TaxTrans.Revenue.RevOpt2 = 0
      TaxTrans.Revenue.RevOpt3 = 0
      TaxTrans.Revenue.PrePaidAmt = 0
      TaxTrans.Revenue.PrePaidBal = GetOverPayBalance(TaxAdvRec.CustRec, "R")
      TaxTrans.Revenue.PrePaidUsed = 0
'      If HamFlag Then
'        TaxTrans(1).Revenue.Interest = TaxAdvRec(1).Amount
'        TaxTrans(1).Revenue.Collection = 0
'      Else
        TaxTrans.Revenue.Interest = 0
        TaxTrans.Revenue.Collection = TaxAdvRec.Amount
'      End If

      TaxTrans.Revenue.Penalty = 0
      TaxTrans.Revenue.Future1 = 0
      TaxTrans.Revenue.Future2 = 0
      TaxTrans.Revenue.Principle1Pd = 0
      TaxTrans.Revenue.Principle2Pd = 0
      TaxTrans.Revenue.Principle3Pd = 0
      TaxTrans.Revenue.Principle4Pd = 0
      TaxTrans.Revenue.Principle5Pd = 0
      TaxTrans.Revenue.InterestPd = 0
      TaxTrans.Revenue.PenaltyPd = 0
      TaxTrans.Revenue.CollectionPd = 0
      TaxTrans.Revenue.Future1Pd = 0
      TaxTrans.Revenue.Future2Pd = 0
      TaxTrans.Revenue.LateListPd = 0
      TaxTrans.Revenue.RevOpt1Pd = 0
      TaxTrans.Revenue.RevOpt2Pd = 0
      TaxTrans.Revenue.RevOpt3Pd = 0
      TaxTrans.DiscAmt = 0
      TaxTrans.OperNum = OperNum
      TaxTrans.DiscXDate = 0
      TaxTrans.Posted2GL = "N"
      TaxTrans.CustomerRec = TaxAdvRec.CustRec
      TaxTrans.LastTrans = 0
      TaxTrans.BelongTo = TaxAdvRec.BillRec
      TaxTrans.RealPin = TaxAdvRec.RealPin 'added 6/6/07
      LSet TaxTrans.Padding = ""
      'Increment Transaction File Record Count
      NextRecord& = (LOF(TTHandle) / Len(TaxTrans)) + 1
      Put TTHandle, NextRecord&, TaxTrans
      'Update the Customer Pointers Now
      Get TCHandle, TaxAdvRec.CustRec, TaxCust
      
      If TaxCust.LastTrans = 0 Then
        TaxCust.LastTrans = NextRecord&
        Put TCHandle, TaxAdvRec.CustRec, TaxCust
      Else
        Previous& = TaxCust.LastTrans
        TaxCust.LastTrans = NextRecord&
        Put TCHandle, TaxAdvRec.CustRec, TaxCust
        Get TTHandle, NextRecord&, TaxTrans
        TaxTrans.LastTrans = Previous&
        Put TTHandle, NextRecord&, TaxTrans
      End If
    End If
  Next
  Close

  'Now Delete the Tax Bill File so Duplicate's Cannot Be Reproduced

  KillFile TaxAdvFile
  
  Call Savemsg(900, "The advertising charges have been posted successfully.")
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxAdvPost", "cmdPost_Click", Erl)
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

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxAdvPost.")
      Call Terminate
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdPost_Click
      KeyCode = 0
  End Select

End Sub


