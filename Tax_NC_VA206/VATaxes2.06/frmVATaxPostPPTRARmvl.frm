VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmVATaxPostPPTRARmvl 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PPTRA Removal Post"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "frmVATaxPostPPTRARmvl.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdPost 
      Height          =   492
      Left            =   6480
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6612
      Width           =   1812
      _Version        =   131072
      _ExtentX        =   3196
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmVATaxPostPPTRARmvl.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   492
      Left            =   3240
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6600
      Width           =   1812
      _Version        =   131072
      _ExtentX        =   3196
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmVATaxPostPPTRARmvl.frx":0AA6
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   840
      Left            =   2304
      Top             =   1200
      Width           =   7020
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PPTRA Removal Post"
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
      Left            =   3828
      TabIndex        =   4
      Top             =   1440
      Width           =   4020
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   2052
      Left            =   2040
      Top             =   3354
      Width           =   7572
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
      Height          =   372
      Left            =   5808
      TabIndex        =   3
      Top             =   4674
      Width           =   3132
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
      Height          =   372
      Left            =   2688
      TabIndex        =   2
      Top             =   4674
      Width           =   3132
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   996
      Left            =   2316
      Top             =   1080
      Width           =   7020
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "                                                                                Ready to Post PPTRA Removal Amounts ? "
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
      Height          =   2052
      Left            =   2040
      TabIndex        =   5
      Top             =   3354
      Width           =   7572
   End
End
Attribute VB_Name = "frmVATaxPostPPTRARmvl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  frmVATaxPPTRAMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPost_Click()
  Dim x As Long, y As Long
  Dim RmvlRec As TaxPPTRARemovalType
  Dim RHandle As Integer
  Dim NumOfRmvlRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim TaxTrans As TaxTransactionType
  Dim NewTaxTrans As TaxTransactionType
  Dim ClearTaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim SaveHere As Long
  Dim TransAmt As Double
  Dim BillNum As Long
  Dim NextTrans As Long
  Dim PostRec As TaxBillPostDateType
  Dim PostHandle As Integer
  Dim NumOfPostRecs As Long
  Dim ThisRec As Integer
  Dim ThatRec As Integer
  Dim SaveCnt As Integer
  
  ThisRec = 0
  ThatRec = 0
  SaveCnt = 0
  ReDim Save2Rec(1 To 1) As Integer
  
  OpenBillPostDateFile PostHandle, NumOfPostRecs
  OpenPPTRARmvlFile RHandle, NumOfRmvlRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  SaveHere = NumOfTTRecs + 1
  For x = 1 To NumOfRmvlRecs
    Get RHandle, x, RmvlRec
    If RmvlRec.CustAcct > 0 Then
      Get TCHandle, RmvlRec.CustAcct, TaxCust
      Get TTHandle, RmvlRec.TransRec, TaxTrans
      ThisRec = RmvlRec.TxBillPostRec
      If ThatRec = 0 Then ThatRec = ThisRec
      If ThatRec <> ThisRec Then
        GoSub Save2Posted
        ThatRec = ThisRec
      End If
      BillNum = QPTrim$(ParseBillNum(TaxTrans.Description))
      NewTaxTrans = ClearTaxTrans
      NewTaxTrans.TransDate = Date2Num%(Date$)
      NewTaxTrans.TaxYear = TaxTrans.TaxYear
      NewTaxTrans.TranType = 30
      NewTaxTrans.BillType = "P"
      NewTaxTrans.Amount = RmvlRec.PPTRADisc
      NewTaxTrans.Revenue.Principle1 = 0
      NewTaxTrans.Revenue.Principle2 = 0
      NewTaxTrans.Revenue.Principle3 = 0
      NewTaxTrans.Revenue.Principle4 = 0
      NewTaxTrans.Revenue.Principle5 = 0
      NewTaxTrans.Revenue.Interest = 0
      NewTaxTrans.Revenue.Penalty = 0
      NewTaxTrans.Revenue.Collection = 0
      NewTaxTrans.Revenue.Future1 = 0
      NewTaxTrans.Revenue.Future2 = 0
      NewTaxTrans.Revenue.Principle1Pd = 0
      NewTaxTrans.Revenue.Principle2Pd = 0
      NewTaxTrans.Revenue.Principle3Pd = 0
      NewTaxTrans.Revenue.Principle4Pd = 0
      NewTaxTrans.Revenue.Principle5Pd = 0
      NewTaxTrans.Revenue.InterestPd = 0
      NewTaxTrans.Revenue.PenaltyPd = 0
      NewTaxTrans.Revenue.CollectionPd = 0
      NewTaxTrans.Revenue.Future1Pd = 0
      NewTaxTrans.Revenue.Future2Pd = 0
      NewTaxTrans.Revenue.RevOpt1 = 0
      NewTaxTrans.Revenue.RevOpt1Pd = 0
      NewTaxTrans.Revenue.RevOpt2 = 0
      NewTaxTrans.Revenue.RevOpt2Pd = 0
      NewTaxTrans.Revenue.RevOpt3 = 0
      NewTaxTrans.Revenue.RevOpt3Pd = 0
      NewTaxTrans.Revenue.LateList = 0
      NewTaxTrans.Revenue.LateListPd = 0
      NewTaxTrans.Revenue.PrePaidAmt = 0
      NewTaxTrans.Revenue.PrePaidUsed = 0
      NewTaxTrans.Revenue.PrePaidBal = OldRound(GetOverPayBalance(RmvlRec.CustAcct, "P") + RmvlRec.PPTRADisc)
      NewTaxTrans.InternalPin = TaxTrans.InternalPin
      NewTaxTrans.Revenue.pad = ""
      NewTaxTrans.Description = "Tax Bill #" + CStr(BillNum)
      NewTaxTrans.Posted2GL = "N"
      NewTaxTrans.CustomerRec = RmvlRec.CustAcct
      NewTaxTrans.LastTrans = TaxCust.LastTrans
      NewTaxTrans.BelongTo = RmvlRec.BelongTo
      NewTaxTrans.Padding = ""
      NewTaxTrans.PersPin = TaxTrans.PersPin
      NewTaxTrans.RealPin = 0
      NewTaxTrans.CustPin = TaxTrans.CustPin
      NewTaxTrans.DiscXDate = 0
      NewTaxTrans.DiscAmt = 0
      NewTaxTrans.OperNum = OperNum
      NewTaxTrans.PersVal = 0
      NewTaxTrans.PPTRAVal = RmvlRec.PPTRAValue
      NewTaxTrans.PPTRADisc = 0
      NewTaxTrans.CntyPara = QPTrim$(TaxTrans.CntyPara)
      NewTaxTrans.CyclPara = QPTrim$(TaxTrans.CyclPara)
      NewTaxTrans.TShpPara = QPTrim$(TaxTrans.TShpPara)
      NewTaxTrans.PPTRARmvl = RmvlRec.PPTRADisc
      NewTaxTrans.PPTRARmvlDate = Date2Num%(Date$)
      Put TTHandle, SaveHere, NewTaxTrans
    
      TaxCust.LastTrans = SaveHere
      Put TCHandle, RmvlRec.CustAcct, TaxCust
      SaveHere = SaveHere + 1
      TaxTrans.Amount = TaxTrans.Amount
      TaxTrans.PPTRARmvl = RmvlRec.PPTRADisc
      TaxTrans.PPTRARmvlDate = Date2Num%(Date$)
      Put TTHandle, RmvlRec.TransRec, TaxTrans
    End If
  Next x
  GoSub Save2Posted
  Close
  
  KillFile PPTRARemovalFile
  Call Savemsg(900, "PPTRA removal amounts have been posted successfully.")
  Call MainLog("PPTRA removal amounts have been posted successfully.")
  Call cmdExit_Click
      
  Exit Sub
  
Save2Posted:
  Get PostHandle, ThatRec, PostRec
  PostRec.PPTRAPosted = "Y"
  Put PostHandle, ThatRec, PostRec

  Return

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
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxPostPPTRARmvl.")
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

