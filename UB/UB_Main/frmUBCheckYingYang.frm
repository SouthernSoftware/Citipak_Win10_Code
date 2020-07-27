VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Southern Software"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkRptOnly 
      Caption         =   "Report Only"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   1200
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GO"
      Height          =   495
      Left            =   1298
      TabIndex        =   0
      Top             =   1785
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   645
      TabIndex        =   2
      Top             =   660
      Width           =   2475
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "UB Fix Ying Yang"
      Height          =   375
      Left            =   705
      TabIndex        =   1
      Top             =   270
      Width           =   2340
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefInt A-Z
Dim Ok2Exit As Boolean
Dim DoneMsgFlag As Boolean
Dim DidCnt As Long

Private Sub Command1_Click()
  If Not Ok2Exit Then
    Command1.Caption = "Wait"
    Command1.Enabled = False
    Call CheckYingYang
    Call AllDone
  Else
    End
  End If
End Sub

Private Sub Form_Load()
  Ok2Exit = False
End Sub

Private Sub CheckYingYang()
    
    Dim UBCust As NewUBCustRecType
    
    Dim CustLen As Integer, XCnt As Integer
    Dim RecCnt As Long, LopCnt As Long
    Dim DidEm As Boolean
    Dim zzCnt As Integer
    Dim dzCnt As Integer
    Dim Book As String, BookD As String
    
    Dim RevCnt As Integer
    Dim zz$, zzz$
    
    Dim TRevAmt As Double
    Dim TotNeg As Double
    
    CustLen = Len(UBCust)
    Dim UBSetupLen%
    Dim GT1 As Double, GT2 As Double
    ReDim RevList(0 To 0) As String
    ReDim CustRevs(1 To 1) As UBCustRevTotalsType
    Dim VBar$
    VBar$ = "|"
    
    Dim ChkCnt As Integer
    Dim PCnt As Integer

    ReDim UBSetUpRec(1) As UBSetupRecType
    UBSetupLen = Len(UBSetUpRec(1))
    Dim DoThisOne As Boolean
    Dim ThisAmt As Double

    DidCnt = 1
    Open "UBSETUP.DAT" For Random Shared As #1 Len = UBSetupLen    'open data file
    Get #1, 1, UBSetUpRec(1)
    Close #1

    For RevCnt = 1 To 15
        If QPTrim(UBSetUpRec(1).Revenues(RevCnt).RevName) = "" Then
            Exit For
        Else
            ReDim Preserve RevList(0 To RevCnt) As String
            RevList(RevCnt) = QPTrim(UBSetUpRec(1).Revenues(RevCnt).RevName)
        End If
    Next
    RevCnt = RevCnt - 1

    ReDim CustFixRevs(1 To 1) As UBCustRevTotalsType

    Open "ubcust.dat" For Random As #1 Len = CustLen
    Open "UBRevInfo.txt" For Output As #2

    RecCnt = LOF(1) / CustLen
    For LopCnt = 1 To RecCnt   'RecCnt
        Get #1, LopCnt, UBCust
        If UBCust.DelFlag = 0 Then     'if not deleted
            DoThisOne = False
            If DidCnt > 1 Then         'if not the first one found
                ReDim Preserve CustRevs(1 To DidCnt) As UBCustRevTotalsType
                ReDim Preserve CustFixRevs(1 To DidCnt) As UBCustRevTotalsType
            End If
            'Do they have a negative balance in a revenue?
            For zzCnt = 1 To RevCnt    'look at each revenue
                If UBCust.CurrRevAmts(zzCnt) < 0 Then  'they have a credit
                    DoThisOne = True   'got one
                    Exit For           'don't need to look further
                End If
            Next
            If DoThisOne Then          'if found one
                DoThisOne = False      'clear flag for next step
                'Do they also owe balance in a revenue?
                For zzCnt = 1 To RevCnt           'look at each revenue
                    If UBCust.CurrRevAmts(zzCnt) > 0 Then 'they also owe
                        DoThisOne = True         'Set flag for next step
                        Exit For                 'no need to look further
                    End If
                Next
            End If
            If DoThisOne Then     'they have both
                CustRevs(DidCnt).CustRec = LopCnt
                CustRevs(DidCnt).CurrBal = Round#(UBCust.CurrBalance)
                CustRevs(DidCnt).TotlBal = CustRevs(DidCnt).CurrBal
                CustFixRevs(DidCnt).TotlBal = CustRevs(DidCnt).CurrBal
                For zzCnt = 1 To RevCnt   'store revenue amounts for action below
                    CustRevs(DidCnt).CurrRev(zzCnt) = Round#(UBCust.CurrRevAmts(zzCnt))
                    CustFixRevs(DidCnt).CurrRev(zzCnt) = Round#(UBCust.CurrRevAmts(zzCnt))
                Next
                DidCnt = DidCnt + 1
            End If
        End If
    Next

    DidCnt = DidCnt - 1 'adjust count to correct amount
    'Look at each ying yang account
    For LopCnt = 1 To DidCnt
        'Sum the total negative/credit amount
        TotNeg = 0
        For zzCnt = 1 To RevCnt                           'for each revenue
            If CustRevs(LopCnt).CurrRev(zzCnt) < 0 Then   'If this one is negative
                TotNeg = TotNeg + CustRevs(LopCnt).CurrRev(zzCnt) 'add to total negative
                CustRevs(LopCnt).CurrRev(zzCnt) = 0       'set that revenue to zero
            End If
        Next
        'convert to a postive amount
        TotNeg = Abs(TotNeg)
        For zzCnt = 1 To RevCnt        'for each revenue
            If TotNeg > 0 Then         'if there is more to distribute
                TRevAmt = CustRevs(LopCnt).CurrRev(zzCnt)  'get this revenue amount
                If TRevAmt > 0 Then               'if there is an amount
                    If TRevAmt <= TotNeg Then  'if the amount is lees or equal to, adjust amount
                        CustRevs(LopCnt).CurrRev(zzCnt) = 0  'set revenue to zero
                        TotNeg = Round(TotNeg - TRevAmt)     'adjust credit remaining
                    Else                       'not enough credit. Adjust by remaining amount
                        CustRevs(LopCnt).CurrRev(zzCnt) = Round(TRevAmt - TotNeg)
                        TotNeg = 0       'no more credit left to distribute.
                        'could just exit for here
                    End If
                End If
            End If
        Next
        'if there is some credit remaining after redistribution, put in 1st revenue
        If TotNeg > 0 Then
            CustRevs(LopCnt).CurrRev(1) = -TotNeg
        End If
    Next

    For LopCnt = 1 To DidCnt
        Get #1, CustRevs(LopCnt).CustRec, UBCust
        TRevAmt = 0
        For zzCnt = 1 To RevCnt
            UBCust.CurrRevAmts(zzCnt) = CustRevs(LopCnt).CurrRev(zzCnt)
            TRevAmt = TRevAmt + UBCust.CurrRevAmts(zzCnt)
            UBCust.PrevRevAmts(zzCnt) = 0
        Next
        UBCust.CurrBalance = TRevAmt
        CustFixRevs(LopCnt).TotlBal = TRevAmt
        UBCust.PrevBalance = 0
        If chkRptOnly.Value = 0 Then
            Put #1, CustRevs(LopCnt).CustRec, UBCust
        End If
    Next

    If chkRptOnly.Value = 1 Then
        For LopCnt = 1 To DidCnt
            Print #2, CustRevs(LopCnt).CustRec, CustFixRevs(LopCnt).TotlBal
            For zzCnt = 1 To RevCnt
                Print #2, RevList(zzCnt); Tab(30); CustRevs(LopCnt).CurrRev(zzCnt); Tab(60); CustFixRevs(LopCnt).CurrRev(zzCnt)
            Next
            Print #2,
            PCnt = PCnt + 1
            If PCnt Mod 6 = 0 Then
                Print #2, Chr$(12)
            End If
        Next
        Print #2,
    End If
  Close
  DidCnt = PCnt

End Sub


Private Sub AllDone()
    If DoneMsgFlag Then
        Command1.Caption = "Already Converted."
    Else
        Label2.Caption = "Done. " + Str$(DidCnt) 'DidCnt
        Command1.Caption = "Exit"
    End If
    Command1.Enabled = True
    Command1.SetFocus
    Ok2Exit = True
End Sub



Public Function QPTrim$(Text As String)
    Dim StrLen As Long
    Dim cnt As Long
    Dim thischar As Integer
    StrLen = Len(Text)
    For cnt = 1 To StrLen
        thischar = Asc(Mid$(Text, cnt, 1))
        If thischar = 0 Then
            Mid$(Text$, cnt, 1) = " "
        End If
    Next
    QPTrim$ = Trim$(Text)
End Function

