VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmVATaxManualPost 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Post Manual Tax Bill Transactions"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxManualPost.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdPost 
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6645
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
      ButtonDesigner  =   "frmVATaxManualPost.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6645
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
      ButtonDesigner  =   "frmVATaxManualPost.frx":0AA6
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "PLEASE NOTE: MANUAL TAX BILL POSTING DOES NOT INTERFACE DIRECTLY TO THE GENERAL LEDGER."
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
      Height          =   732
      Left            =   2394
      TabIndex        =   4
      Top             =   4130
      Width           =   6972
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   840
      Left            =   2297
      Top             =   1715
      Width           =   7020
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Manual Tax Bill Transactions Post"
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
      Left            =   3152
      TabIndex        =   3
      Top             =   1955
      Width           =   5325
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   3132
      Left            =   2034
      Top             =   3026
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
      Left            =   5802
      TabIndex        =   1
      Top             =   5186
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
      Left            =   2682
      TabIndex        =   0
      Top             =   5186
      Width           =   3132
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   990
      Left            =   2312
      Top             =   1595
      Width           =   7020
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "                                                                                Ready to Post Manual Tax Bill Transactions? "
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
      Height          =   3132
      Left            =   2034
      TabIndex        =   2
      Top             =   3026
      Width           =   7572
   End
End
Attribute VB_Name = "frmVATaxManualPost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class
Private Sub cmdExit_Click()
  Call TaxMsg(900, "Manual Tax Transaction Posting has been cancelled.")
  frmVATaxManualBillMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPost_Click()
  Dim TaxCustRec As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim OPTaxTrans As TaxTransactionType
  Dim NumOfOPTTRecs As Long
  Dim OPTTHandle As Integer
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim RealRec As PropertyRecType
  Dim RRHandle As Integer
  Dim NumOfRRREcs As Long
  Dim Revenue As WinRevSourceType
  Dim TaxMRec As TaxMTransactionType
  Dim TMHandle As Integer
  Dim NumOfTMRecs As Integer
  Dim PersPropRec As PersonalRecType
  Dim PRHandle As Integer
  Dim NumOfPRRecs As Long
  Dim x As Integer, y As Integer
  Dim BillTotal As Double
  Dim WhatsLeft As Double
  Dim Previous&, NextRecord&
  Dim Handle As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TSUHandle As Integer
  
  On Error GoTo ERRORSTUFF
  
  OpenTaxSetUpFile TSUHandle
  Get TSUHandle, 1, TaxMasterRec
  Close TSUHandle
  
  ReDim PayOrder(1 To 10) As Integer
  PayOrder(1) = TaxMasterRec.PersPayOrder
  PayOrder(2) = TaxMasterRec.MTPayOrder
  PayOrder(3) = TaxMasterRec.MCPayOrder
  PayOrder(4) = TaxMasterRec.FEPayOrder
  PayOrder(5) = TaxMasterRec.MHPayOrder
  PayOrder(6) = TaxMasterRec.PIntPayOrder
  PayOrder(7) = TaxMasterRec.PPenPayOrder
  PayOrder(8) = TaxMasterRec.POpt1PayOrder
  PayOrder(9) = TaxMasterRec.POpt2PayOrder
  PayOrder(10) = TaxMasterRec.POpt3PayOrder
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenRealPropFile RRHandle, NumOfRRREcs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenPersPropFile PRHandle, NumOfPRRecs
  OpenTaxManualBillFile TMHandle, NumOfTMRecs
  
  For x = 1 To NumOfTMRecs
    Get TMHandle, x, TaxMRec
    If TaxMRec.Deleted = True Then GoTo SkipIt
    If TaxMRec.BillType <> "P" Then
      BillTotal = OldRound(TaxMRec.TaxAmount + TaxMRec.IntAmount + TaxMRec.AdColAmount + TaxMRec.Penalty)
      BillTotal = OldRound(BillTotal + TaxMRec.LateList + TaxMRec.OptRev1 + TaxMRec.OptRev2 + TaxMRec.OptRev3)
    ElseIf TaxMRec.BillType = "P" Then
      BillTotal = OldRound(TaxMRec.Personal + TaxMRec.MachTools + TaxMRec.MerchCap + TaxMRec.FarmEquip + TaxMRec.MobHomes)
      BillTotal = OldRound(BillTotal + TaxMRec.IntAmount + TaxMRec.OptRev1 + TaxMRec.OptRev2 + TaxMRec.OptRev3 + TaxMRec.Penalty)
    End If
    If BillTotal = 0 Then GoTo SkipIt
    TaxTrans.TransDate = TaxMRec.TransDate
    TaxTrans.TaxYear = TaxMRec.TaxYear
    TaxTrans.BillType = TaxMRec.BillType            'R=Real P=Personal Property C=Combined (NC/GA)
    TaxTrans.TranType = 1                '1=Bill 2=Payment 3=Release 4=Interest 5=Penalty 6=Collection/Ad Cost Billing
    TaxTrans.Amount = BillTotal      'Total Transaction Amount
    If TaxMRec.BillType <> "P" Then
      TaxTrans.Revenue.Principle1 = TaxMRec.TaxAmount
      TaxTrans.Revenue.Principle2 = 0
      TaxTrans.Revenue.Principle3 = 0
      TaxTrans.Revenue.Principle4 = 0
      TaxTrans.Revenue.Principle5 = 0
    Else
      TaxTrans.Revenue.Principle1 = TaxMRec.Personal
      TaxTrans.Revenue.Principle2 = TaxMRec.MachTools
      TaxTrans.Revenue.Principle3 = TaxMRec.MerchCap
      TaxTrans.Revenue.Principle4 = TaxMRec.FarmEquip
      TaxTrans.Revenue.Principle5 = TaxMRec.MobHomes
    End If
    TaxTrans.Revenue.Interest = TaxMRec.IntAmount
    TaxTrans.Revenue.LateList = TaxMRec.LateList
    TaxTrans.Revenue.RevOpt1 = TaxMRec.OptRev1
    TaxTrans.Revenue.RevOpt2 = TaxMRec.OptRev2
    TaxTrans.Revenue.RevOpt3 = TaxMRec.OptRev3
    TaxTrans.Revenue.Penalty = TaxMRec.Penalty
    TaxTrans.Revenue.Collection = TaxMRec.AdColAmount
    TaxTrans.Revenue.Future1 = 0
    TaxTrans.Revenue.Future2 = 0
    TaxTrans.Revenue.PrePaidAmt = 0
    TaxTrans.Revenue.PrePaidUsed = 0
    If TaxMRec.BillType <> "P" Then
      TaxTrans.Revenue.PrePaidBal = GetOverPayBalance(TaxMRec.Account, "R")
    Else
      TaxTrans.Revenue.PrePaidBal = GetOverPayBalance(TaxMRec.Account, "P")
    End If
    TaxTrans.Revenue.Principle1Pd = 0
    TaxTrans.Revenue.Principle2Pd = 0
    TaxTrans.Revenue.Principle3Pd = 0
    TaxTrans.Revenue.Principle4Pd = 0
    TaxTrans.Revenue.Principle5Pd = 0
    TaxTrans.Revenue.InterestPd = 0
    TaxTrans.Revenue.PenaltyPd = 0
    TaxTrans.Revenue.CollectionPd = 0
    TaxTrans.Revenue.LateListPd = 0
    TaxTrans.Revenue.Future1Pd = 0
    TaxTrans.Revenue.Future2Pd = 0
    TaxTrans.Revenue.RevOpt1Pd = 0
    TaxTrans.Revenue.RevOpt2Pd = 0
    TaxTrans.Revenue.RevOpt3Pd = 0
    TaxTrans.Revenue.pad = ""
    TaxTrans.DiscXDate = 0
    TaxTrans.DiscAmt = 0
    TaxTrans.OperNum = OperNum
    If TaxMRec.RealRec > 0 Then
      Get RRHandle, TaxMRec.RealRec, RealRec
      TaxTrans.InternalPin = RealRec.InternalPin
      TaxTrans.RealPin = RealRec.RealPin
    Else
      TaxTrans.RealPin = 0
    End If
    If TaxMRec.PersRec > 0 Then
      Get PRHandle, TaxMRec.PersRec, PersPropRec
      TaxTrans.InternalPin = PersPropRec.InternalPin
      TaxTrans.PersPin = PersPropRec.PropPin
    Else
      TaxTrans.PersPin = 0
    End If
    If TaxMRec.PersRec = 0 And TaxMRec.RealRec = 0 Then
      TaxTrans.InternalPin = 0
      TaxTrans.RealPin = -1 'indicates MOCK
    End If
    
    TaxTrans.FromPrePay = 0
    
    TaxTrans.Description = TaxMRec.Desc
    TaxTrans.Posted2GL = "Y" ' Do Not Allow Posting 2GL of Manual Entries Probably Already Reflected in General Ledger
    TaxTrans.CustomerRec = TaxMRec.Account
    TaxTrans.LastTrans = 0
    TaxTrans.BelongTo = 0
    TaxTrans.Padding = ""
    
    'Increment Transaction File Record Count
    NextRecord& = (LOF(TTHandle) / Len(TaxTrans)) + 1
    Put TTHandle, NextRecord&, TaxTrans
    
    'Update the Customer Pointers Now
    Get TCHandle, TaxMRec.Account, TaxCustRec
    
    If TaxCustRec.LastTrans = 0 Then
      TaxCustRec.LastTrans = NextRecord&
      Put TCHandle, TaxMRec.Account, TaxCustRec
    Else
      Previous& = TaxCustRec.LastTrans
      TaxCustRec.LastTrans = NextRecord&
      Put TCHandle, TaxMRec.Account, TaxCustRec
      
      Get TTHandle, NextRecord&, TaxTrans
      TaxTrans.LastTrans = Previous&
      TaxTrans.CustPin = TaxCustRec.PIN
      Put TTHandle, NextRecord&, TaxTrans
    End If
    
    If TaxMRec.OverPayUsed <> 0 Then
      GoSub ApplyCredit
    End If
SkipIt:
  Next x
  
  Close
  KillFile TaxManualBill '5.16.07
  Call Savemsg(900, "Manual Tax Bill Posting was completed successfully.")
  MainLog ("Manual tax billing posted successfully.")
  frmVATaxManualBillMenu.Show
  DoEvents
  Unload Me
  
  Exit Sub
  
ApplyCredit:
  
  WhatsLeft = Abs(TaxMRec.OverPayUsed)
  OpenTaxTransFile OPTTHandle, NumOfOPTTRecs
  
  TaxTrans.TranType = 9
  If TaxMRec.BillType = "R" Then
    If TaxTrans.Revenue.Interest >= WhatsLeft Then
      OPTaxTrans.Revenue.InterestPd = TaxTrans.Revenue.Interest
      WhatsLeft = 0
    Else
      OPTaxTrans.Revenue.InterestPd = TaxTrans.Revenue.Interest
      WhatsLeft = OldRound(WhatsLeft - TaxTrans.Revenue.Interest)
    End If
  
    If TaxTrans.Revenue.Collection >= WhatsLeft Then
      OPTaxTrans.Revenue.CollectionPd = TaxTrans.Revenue.Collection
      WhatsLeft = 0
    Else
      OPTaxTrans.Revenue.CollectionPd = TaxTrans.Revenue.Collection
      WhatsLeft = OldRound(WhatsLeft - TaxTrans.Revenue.Collection)
    End If
  
    If TaxTrans.Revenue.LateList >= WhatsLeft Then
      OPTaxTrans.Revenue.LateListPd = TaxTrans.Revenue.LateList
      WhatsLeft = 0
    Else
      OPTaxTrans.Revenue.LateListPd = TaxTrans.Revenue.LateList
      WhatsLeft = OldRound(WhatsLeft - TaxTrans.Revenue.LateList)
    End If
  
    If TaxTrans.Revenue.Penalty >= WhatsLeft Then
      OPTaxTrans.Revenue.PenaltyPd = WhatsLeft
      WhatsLeft = 0
    Else
      OPTaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.Penalty
      WhatsLeft = OldRound(WhatsLeft - TaxTrans.Revenue.Penalty)
    End If
    
    If TaxTrans.Revenue.Principle1 >= WhatsLeft Then
      OPTaxTrans.Revenue.Principle1Pd = WhatsLeft
      WhatsLeft = 0
    Else
      OPTaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1
      WhatsLeft = OldRound(WhatsLeft - TaxTrans.Revenue.Principle1)
    End If
    
    
    If TaxTrans.Revenue.RevOpt1 >= WhatsLeft Then
      OPTaxTrans.Revenue.RevOpt1Pd = WhatsLeft
      WhatsLeft = 0
    Else
      OPTaxTrans.Revenue.RevOpt1Pd = TaxTrans.Revenue.RevOpt1
      WhatsLeft = OldRound(WhatsLeft - TaxTrans.Revenue.RevOpt1)
    End If
  
    If TaxTrans.Revenue.RevOpt2 >= WhatsLeft Then
      OPTaxTrans.Revenue.RevOpt2Pd = WhatsLeft
      WhatsLeft = 0
    Else
      OPTaxTrans.Revenue.RevOpt2Pd = TaxTrans.Revenue.RevOpt2
      WhatsLeft = OldRound(WhatsLeft - TaxTrans.Revenue.RevOpt2)
    End If
  
    If TaxTrans.Revenue.RevOpt3 >= WhatsLeft Then
      OPTaxTrans.Revenue.RevOpt3Pd = WhatsLeft
      WhatsLeft = 0
    Else
      OPTaxTrans.Revenue.RevOpt3Pd = TaxTrans.Revenue.RevOpt3
      WhatsLeft = OldRound(WhatsLeft - TaxTrans.Revenue.RevOpt3)
    End If
  ElseIf TaxMRec.BillType = "P" Then
    For y = 1 To 10
      If y = PayOrder(1) Then
        If TaxTrans.Revenue.Principle1 >= WhatsLeft Then
          OPTaxTrans.Revenue.Principle1Pd = WhatsLeft
          WhatsLeft = 0
        Else
          OPTaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1
          WhatsLeft = OldRound(WhatsLeft - TaxTrans.Revenue.Principle1)
        End If
      ElseIf y = PayOrder(2) Then
        If TaxTrans.Revenue.Principle2 >= WhatsLeft Then
          OPTaxTrans.Revenue.Principle2Pd = WhatsLeft
          WhatsLeft = 0
        Else
          OPTaxTrans.Revenue.Principle2Pd = TaxTrans.Revenue.Principle2
          WhatsLeft = OldRound(WhatsLeft - TaxTrans.Revenue.Principle2)
        End If
      ElseIf y = PayOrder(3) Then
        If TaxTrans.Revenue.Principle3 >= WhatsLeft Then
          OPTaxTrans.Revenue.Principle3Pd = WhatsLeft
          WhatsLeft = 0
        Else
          OPTaxTrans.Revenue.Principle3Pd = TaxTrans.Revenue.Principle3
          WhatsLeft = OldRound(WhatsLeft - TaxTrans.Revenue.Principle3)
        End If
      ElseIf y = PayOrder(4) Then
        If TaxTrans.Revenue.Principle4 >= WhatsLeft Then
          OPTaxTrans.Revenue.Principle4Pd = WhatsLeft
          WhatsLeft = 0
        Else
          OPTaxTrans.Revenue.Principle4Pd = TaxTrans.Revenue.Principle4
          WhatsLeft = OldRound(WhatsLeft - TaxTrans.Revenue.Principle4)
        End If
      ElseIf y = PayOrder(5) Then
        If TaxTrans.Revenue.Principle5 >= WhatsLeft Then
          OPTaxTrans.Revenue.Principle5Pd = WhatsLeft
          WhatsLeft = 0
        Else
          OPTaxTrans.Revenue.Principle5Pd = TaxTrans.Revenue.Principle5
          WhatsLeft = OldRound(WhatsLeft - TaxTrans.Revenue.Principle5)
        End If
      ElseIf y = PayOrder(6) Then
        If TaxTrans.Revenue.Interest >= WhatsLeft Then
          OPTaxTrans.Revenue.InterestPd = WhatsLeft
          WhatsLeft = 0
        Else
          OPTaxTrans.Revenue.InterestPd = TaxTrans.Revenue.Interest
          WhatsLeft = OldRound(WhatsLeft - TaxTrans.Revenue.Interest)
        End If
      ElseIf y = PayOrder(7) Then
        If TaxTrans.Revenue.Penalty >= WhatsLeft Then
          OPTaxTrans.Revenue.PenaltyPd = WhatsLeft
          WhatsLeft = 0
        Else
          OPTaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.Penalty
          WhatsLeft = OldRound(WhatsLeft - TaxTrans.Revenue.Penalty)
        End If
      ElseIf y = PayOrder(8) Then
        If TaxTrans.Revenue.RevOpt1 >= WhatsLeft Then
          OPTaxTrans.Revenue.RevOpt1Pd = WhatsLeft
          WhatsLeft = 0
        Else
          OPTaxTrans.Revenue.RevOpt1Pd = TaxTrans.Revenue.RevOpt1
          WhatsLeft = OldRound(WhatsLeft - TaxTrans.Revenue.RevOpt1)
        End If
      ElseIf y = PayOrder(9) Then
        If TaxTrans.Revenue.RevOpt2 >= WhatsLeft Then
          OPTaxTrans.Revenue.RevOpt2Pd = WhatsLeft
          WhatsLeft = 0
        Else
          OPTaxTrans.Revenue.RevOpt2Pd = TaxTrans.Revenue.RevOpt2
          WhatsLeft = OldRound(WhatsLeft - TaxTrans.Revenue.RevOpt2)
        End If
      ElseIf y = PayOrder(10) Then
        If TaxTrans.Revenue.RevOpt3 >= WhatsLeft Then
          OPTaxTrans.Revenue.RevOpt3Pd = WhatsLeft
          WhatsLeft = 0
        Else
          OPTaxTrans.Revenue.RevOpt3Pd = TaxTrans.Revenue.RevOpt3
          WhatsLeft = OldRound(WhatsLeft - TaxTrans.Revenue.RevOpt3)
        End If
      End If
    Next y
  End If
  OPTaxTrans.Revenue.Principle1 = 0
  OPTaxTrans.Revenue.Principle2 = 0
  OPTaxTrans.Revenue.Principle3 = 0
  OPTaxTrans.Revenue.Principle4 = 0
  OPTaxTrans.Revenue.Principle5 = 0
  OPTaxTrans.Revenue.Collection = 0
  OPTaxTrans.Revenue.Future1 = 0
  OPTaxTrans.Revenue.Future2 = 0
  OPTaxTrans.Revenue.Interest = 0
  OPTaxTrans.Revenue.LateList = 0
  OPTaxTrans.Revenue.Penalty = 0
'  OPTaxTrans.Revenue.PenaltyPd = 0
  OPTaxTrans.Revenue.RevOpt1 = 0
  OPTaxTrans.Revenue.RevOpt2 = 0
  OPTaxTrans.Revenue.RevOpt3 = 0
  
  OPTaxTrans.Revenue.PrePaidAmt = 0
  OPTaxTrans.Revenue.PrePaidUsed = OldRound(Abs(TaxMRec.OverPayUsed) - WhatsLeft)
  OPTaxTrans.Revenue.PrePaidBal = WhatsLeft
  OPTaxTrans.TranType = 9  '9 - Credit applied to bill 1=Bill 2=Payment 3=Release 4=Interest 5=Penalty 6=Collection/Ad Cost Billing
  OPTaxTrans.Amount = 0
  OPTaxTrans.FromPrePay = OldRound(TaxMRec.OverPayUsed - WhatsLeft)
  OPTaxTrans.InternalPin = TaxTrans.InternalPin
  OPTaxTrans.CustomerRec = GCustNum
  OPTaxTrans.CustPin = TaxTrans.CustPin
  OPTaxTrans.BelongTo = NextRecord&
  OPTaxTrans.Description = "Credit Applied to MBill# " + CStr(OPTaxTrans.BelongTo)
  OPTaxTrans.LastTrans = TaxCustRec.LastTrans
  OPTaxTrans.TaxYear = TaxTrans.TaxYear
  OPTaxTrans.DiscAmt = 0
  OPTaxTrans.OperNum = OperNum
  OPTaxTrans.DiscXDate = 0
  OPTaxTrans.PersPin = TaxTrans.PersPin
  OPTaxTrans.RealPin = TaxTrans.RealPin
  OPTaxTrans.Posted2GL = "Y"
  OPTaxTrans.TransDate = TaxTrans.TransDate
  OPTaxTrans.BillType = TaxMRec.BillType
  Get TTHandle, NextRecord, TaxTrans
  If TaxMRec.BillType = "R" Then
    TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd + OPTaxTrans.Revenue.Principle1Pd)
    TaxTrans.Revenue.Principle2Pd = 0
    TaxTrans.Revenue.Principle3Pd = 0
    TaxTrans.Revenue.Principle4Pd = 0
    TaxTrans.Revenue.Principle5Pd = 0
    TaxTrans.Revenue.InterestPd = OldRound(TaxTrans.Revenue.InterestPd + OPTaxTrans.Revenue.InterestPd)
    TaxTrans.Revenue.CollectionPd = OldRound(TaxTrans.Revenue.CollectionPd + OPTaxTrans.Revenue.CollectionPd)
    TaxTrans.Revenue.LateListPd = OldRound(TaxTrans.Revenue.LateListPd + OPTaxTrans.Revenue.LateListPd)
    TaxTrans.Revenue.PenaltyPd = OldRound(TaxTrans.Revenue.PenaltyPd + OPTaxTrans.Revenue.PenaltyPd)
    TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd + OPTaxTrans.Revenue.RevOpt1Pd)
    TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd + OPTaxTrans.Revenue.RevOpt2Pd)
    TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd + OPTaxTrans.Revenue.RevOpt3Pd)
  ElseIf TaxMRec.BillType = "P" Then
    TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd + OPTaxTrans.Revenue.Principle1Pd)
    TaxTrans.Revenue.Principle2Pd = OldRound(TaxTrans.Revenue.Principle2Pd + OPTaxTrans.Revenue.Principle2Pd)
    TaxTrans.Revenue.Principle3Pd = OldRound(TaxTrans.Revenue.Principle3Pd + OPTaxTrans.Revenue.Principle3Pd)
    TaxTrans.Revenue.Principle4Pd = OldRound(TaxTrans.Revenue.Principle4Pd + OPTaxTrans.Revenue.Principle4Pd)
    TaxTrans.Revenue.Principle5Pd = OldRound(TaxTrans.Revenue.Principle5Pd + OPTaxTrans.Revenue.Principle5Pd)
    TaxTrans.Revenue.InterestPd = OldRound(TaxTrans.Revenue.InterestPd + OPTaxTrans.Revenue.InterestPd)
    TaxTrans.Revenue.CollectionPd = 0
    TaxTrans.Revenue.LateListPd = 0
    TaxTrans.Revenue.PenaltyPd = OldRound(TaxTrans.Revenue.PenaltyPd + OPTaxTrans.Revenue.PenaltyPd)
    TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd + OPTaxTrans.Revenue.RevOpt1Pd)
    TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd + OPTaxTrans.Revenue.RevOpt2Pd)
    TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd + OPTaxTrans.Revenue.RevOpt3Pd)
  End If
  Put TTHandle, NextRecord, TaxTrans
  
  NextRecord = NextRecord + 1
  Put OPTTHandle, NextRecord, OPTaxTrans
  
  Close OPTTHandle
  Get TCHandle, TaxMRec.Account, TaxCustRec
  
  TaxCustRec.LastTrans = NextRecord
  Put TCHandle, TaxMRec.Account, TaxCustRec
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxManualPost", "cmdPost_Click", Erl)
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

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxManualPost.")
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

