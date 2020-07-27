VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmTaxInterestPost 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Interest Calculation Post"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTaxIntPost.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdPost 
      Height          =   495
      Left            =   6420
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6450
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
      ButtonDesigner  =   "frmTaxIntPost.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   3420
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6450
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
      ButtonDesigner  =   "frmTaxIntPost.frx":0AA6
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   840
      Left            =   2303
      Top             =   1905
      Width           =   7020
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Interest Calculations Post"
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
      Left            =   3833
      TabIndex        =   5
      Top             =   2145
      Width           =   4020
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   2055
      Left            =   2040
      Top             =   3450
      Width           =   7575
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
      Left            =   5805
      TabIndex        =   3
      Top             =   4770
      Width           =   3135
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
      Left            =   2685
      TabIndex        =   2
      Top             =   4770
      Width           =   3135
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   990
      Left            =   2318
      Top             =   1785
      Width           =   7020
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   $"frmTaxIntPost.frx":0C82
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
      Left            =   2040
      TabIndex        =   4
      Top             =   3450
      Width           =   7575
   End
End
Attribute VB_Name = "frmTaxInterestPost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  frmTaxInterestMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPost_Click()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim Cnt As Long, Previous&
  Dim DidSome As Long
  Dim TaxTrans As TaxTransactionType
  Dim NewTaxTrans As TaxTransactionType
  Dim ClearTaxTrans As TaxTransactionType
  Dim TaxIntRec As InterestRecType
  Dim IntTrans As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, NextRecord&
  Dim IntDateRec As TaxInterestDateType
  Dim IDHandle As Integer
  
  'on error goto ERRORSTUFF
  If TaxMsgWOpts(900, "If you are sure you are ready to post then press F10 to continue. Otherwise, press ESC to abort the post attempt.", "F10 Continue", "ESC Abort") = "abort" Then
    Unload frmTaxMsgWOpts
    Call TaxMsg(900, "Post attempt aborted.")
    Close
    Exit Sub
  Else
    Unload frmTaxMsgWOpts
    MainLog ("Interest calculations posted.")
  End If
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenInterestRecFile IRHandle, NumOfIRRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  frmTaxShowPctComp.Label1 = "Posting Interest"
  frmTaxShowPctComp.Show , Me
  frmTaxShowPctComp.cmdCancel.Visible = False
  cmdExit.Enabled = False
  cmdPost.Enabled = False
  
  For Cnt& = 1 To NumOfIRRecs
    Get IRHandle, Cnt&, TaxIntRec
    If TaxIntRec.DelFlag = 0 Then
      'Update the Bill transaction first
      'TaxIntRec(1).BillRec
      Get TTHandle, TaxIntRec.BillRec, TaxTrans 'get bill trans
      If TaxIntRec.Amount = 0 Then GoTo SkipIt 'edited to zero
      TaxTrans.Revenue.Interest = OldRound#(TaxTrans.Revenue.Interest + TaxIntRec.Amount)
      Put #TTHandle, TaxIntRec.BillRec, TaxTrans 'put it back
      'Now make a new clean transaction
      NewTaxTrans = ClearTaxTrans
      NewTaxTrans.TransDate = Date2Num%(Date$)
      NewTaxTrans.TaxYear = TaxIntRec.TaxYear
      NewTaxTrans.TranType = 4       '4=Interest
      NewTaxTrans.BillType = ""     'R=Real P=Personal Property C=Combined (NC/GA)
      NewTaxTrans.Amount = TaxIntRec.Amount  'Total Transaction Amount
      NewTaxTrans.Revenue.Interest = TaxIntRec.Amount
      NewTaxTrans.Description = "Tax Int on Bill# " + QPTrim$(TaxIntRec.BillNumber)
      NewTaxTrans.Posted2GL = "N"
      NewTaxTrans.CustomerRec = TaxIntRec.CustRec
      NewTaxTrans.CustPin = TaxIntRec.CustPin
      NewTaxTrans.RealPin = TaxIntRec.RealPin
      NewTaxTrans.PersPin = TaxIntRec.PersPin
      NewTaxTrans.LastTrans = 0
      NewTaxTrans.BelongTo = TaxIntRec.BillRec
      NewTaxTrans.Revenue.PrePaidAmt = 0
      NewTaxTrans.Revenue.PrePaidBal = OldRound(GetOverPayBalance(TaxIntRec.CustRec))
      NewTaxTrans.Revenue.PrePaidUsed = 0
      NewTaxTrans.OperNum = OperNum
      LSet NewTaxTrans.Padding = ""
      'Increment Transaction File Record Count
      NextRecord& = (LOF(TTHandle) / Len(NewTaxTrans)) + 1
      Put TTHandle, NextRecord&, NewTaxTrans
      'Update the Customer Pointers Now
      Get TCHandle, TaxIntRec.CustRec, TaxCust
      
      If TaxCust.LastTrans = 0 Then
        TaxCust.LastTrans = NextRecord&
        Put TCHandle, TaxIntRec.CustRec, TaxCust
      Else
        Previous& = TaxCust.LastTrans
        TaxCust.LastTrans = NextRecord&
        Put TCHandle, TaxIntRec.CustRec, TaxCust
        Get TTHandle, NextRecord&, NewTaxTrans
        NewTaxTrans.LastTrans = Previous&
        Put TTHandle, NextRecord&, NewTaxTrans
      End If
    End If
SkipIt:
    frmTaxShowPctComp.ShowPctComp Cnt, NumOfIRRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdPost.Enabled = True
      Exit Sub
    End If
  Next
  
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdPost.Enabled = True
  KillFile "TAXINTCK.DAT"
  OpenTxIntTickFile IDHandle
  IntDateRec.InterestDate = Date2Num%(Date$)
  Put IDHandle, 1, IntDateRec
  Close IDHandle
  Close
  
  'Now Delete the Tax Bill File so Duplicate's Cannot Be Reproduced
  KillFile "TAXINT.DAT"
  
  Call Savemsg(900, "The interest calculations have been posted successfully.")
  Call cmdExit_Click
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxInterestPost", "cmdPost_Click", Erl)
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
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxInterestPost.")
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

