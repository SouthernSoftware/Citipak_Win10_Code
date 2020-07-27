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
   Begin VB.CommandButton Command1 
      Caption         =   "GO"
      Height          =   495
      Left            =   1298
      TabIndex        =   0
      Top             =   1425
      Width           =   1155
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   600
      TabIndex        =   3
      Top             =   960
      Width           =   2475
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   240
      Left            =   645
      TabIndex        =   2
      Top             =   660
      Width           =   2475
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "UB Utility"
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
    Call AddZeros
    'Call CheckTransTotals
    Call AllDone
  Else
    End
  End If
End Sub

Private Sub Form_Load()
  Ok2Exit = False
End Sub

Private Sub CheckTransTotals()
    Dim UBTrans As UBTransRecType
    Dim TRLen As Integer
    Dim RCnt As Long
    Dim TRCnt As Long
    Dim PCnt As Integer
    Dim TRDate As Integer
    Dim GTotalTR As Double
    Dim GTotalRV As Double
    Dim RevCnt As Integer
    'Dim RVCnt As Integer
    
    TRLen = Len(UBTrans)
    TRDate = Date2Num("03/20/2020")
    Open "ubtrans.dat" For Random As #1 Len = TRLen
    TRCnt = LOF(1) \ TRLen
    Get #1, 2746043, UBTrans
    UBTrans.RevAmt(1) = 28.37
    Put #1, 2746043, UBTrans
    '2746043
'    For RCnt = 1 To TRCnt
'      Get #1, RCnt, UBTrans
'      If UBTrans.TransDate = TRDate Then
'         If UBTrans.TransType = 4 Then
'           GTotalTR = GTotalTR + UBTrans.Transamt
'           GTotalRV = 0
'           For RevCnt = 1 To 14
'             GTotalRV = GTotalRV + UBTrans.RevAmt(RevCnt)
'           Next
'           If GTotalRV <> UBTrans.Transamt Then Stop
'           Print UBTrans.CustAcctNo
'           PCnt = PCnt + 1
'         End If
'      End If
'    Next
    Close
    Label3.Caption = " Correction Complete." '"TR: " + Str$(GTotalTR) + " RV: " + Str$(GTotalRV)
End Sub

Private Sub AddZeros()
    
    Dim UBCust As NewUBCustRecType
    Dim CustLen As Integer, XCnt As Integer
    Dim RecCnt As Long, LopCnt As Long
    
    Dim zzCnt As Integer
    Dim BDate As Integer
    Dim GGDate As Integer
    'BDate = Date2Num("01/14/2020")
    'GGDate = Date2Num("01/07/2020")
    Dim zz As Integer

    CustLen = Len(UBCust)

    Open "ubcust.dat" For Random As #1 Len = CustLen
    'RecCnt = LOF(1) / CustLen

    'For LopCnt = 1 To RecCnt   'RecCnt
        Get #1, 5263, UBCust
        UBCust.CurrBalance = -0.01
        UBCust.CurrRevAmts(1) = -0.01
        Put #1, 5263, UBCust

    'Next
    Close
    'Label3.Caption = "Done"


''''    DidEm = False
''''    Select Case UBCust.Book
''''    Case "08"
''''        UBCust.Book = "03"
''''        DidEm = True
''''    Case "06"
''''        UBCust.Book = "04"
''''        DidEm = True
''''    End Select
''''
'''''    Select Case QPTrim(UBCust.USERCODE2)
'''''    Case "C"
'''''      UBCust.USERCODE2 = "1"
'''''      DidEm = True
'''''    Case "N"
'''''      UBCust.USERCODE2 = "2"
'''''      DidEm = True
'''''    Case "Z"
'''''      UBCust.USERCODE2 = "3"
'''''      DidEm = True
'''''    End Select
''''    If DidEm Then
''''      Put #1, LopCnt, UBCust
''''      DidCnt = DidCnt + 1
''''      DidEm = False
''''    End If
'''''    If UBCust.Status <> "A" Then
'''''      If Len(QPTrim(UBCust.Book)) = 0 Then
'''''        Get #2, LopCnt, UBCustN
'''''        If Len(QPTrim(UBCustN.Book)) > 0 Then
'''''          UBCust.Book = UBCustN.Book
'''''          UBCust.SEQNUMB = UBCustN.SEQNUMB
'''''          DidCnt = DidCnt + 1
'''''          Put #1, LopCnt, UBCust
'''''        End If
'''''      End If
'''''    End If
''''  Next
''''  Close
''''
'''''  BDate = Date2Num%("12/31/1979")
'''''  GGDate = Date2Num%("09/24/2019")
'''''  Open "ubtrans.dat" For Random As #1 Len = TRLen
'''''  RecCnt = LOF(1) / TRLen
'''''  For LopCnt = 1 To RecCnt
'''''    Get #1, LopCnt, TransRec
'''''    If TransRec.TransDate = BDate And TransRec.TransType = 1 Then '(RecCnt - LopCnt) < 5000 Then
'''''      TransRec.TransDate = GGDate
'''''      DidCnt = DidCnt + 1
'''''      'TransRec.CustAcctNo = -2
'''''      Put #1, LopCnt, TransRec
'''''    End If
'''''  Next
'''''  Close
''''
''''' End
''''
'''''    MtrNum    As String * 12
'''''    MTRMulti  As Integer
'''''    MTRType   As String * 1
'''''    MtrUnit   As String * 1
'''''    NumUser   As Integer
'''''    InsDate   As Integer
'''''    CurRead   As Long
'''''    PrevRead  As Long
'''''    CurDate   As Integer
'''''    PastDate  As Integer       'hidden & protected
'''''    ReadFlag  As String * 1    'hidden & protected
'''''    AvgUse    As Long          'hidden & protected
'''''    UseCnt    As Integer       'hidden & protected
'''''    MtrIDNO   As String * 11
'''''    MtrLat    As Double
'''''    MtrLng    As Double
''''
'''''Sets usercode
'''''  Open "ubcust.dat" For Random As #1 Len = CustLen
'''''  RecCnt = LOF(1) / CustLen
''''''
'''''  For LopCnt = 1 To RecCnt
'''''     Get #1, LopCnt, UBCust
'''''     UBCust.USERCODE2 = "5"
'''''     For zzCnt = 1 To 7
'''''       zz$ = QPTrim(UBCust.LocMeters(zzCnt).MtrNum)
'''''       If Len(zz$) > 0 Then
'''''         zzz$ = Right$(zz$, 1)
'''''         If zzz$ <> "Z" Then
'''''           UBCust.LocMeters(zzCnt).MtrNum = zz$ + "Z"
'''''         End If
'''''       End If
'''''     Next
'''''     Put #1, LopCnt, UBCust
'''''  Next
''''''''''''
''''''''''''
''''''''''''
''''' Close
'''''  Open "ubtrans.dat" For Random As #1 Len = TRLen
'''''  RecCnt = LOF(1) / TRLen
'''''  For LopCnt = 1 To RecCnt
'''''    Get #1, LopCnt, TransRec
'''''    If TransRec.CustAcctNo = 2 Then
'''''      DidCnt = DidCnt + 1
'''''      TransRec.CustAcctNo = -2
'''''      Put #1, LopCnt, TransRec
'''''    End If
'''''  Next
'''''  Close
'''''
'''''  Open "ubcust.dat" For Random As #1 Len = CustLen
''''    '  Get #1, 2, UBCust
'''''  UBCust.CurrBalance = 0
'''''  UBCust.PrevBalance = 0
'''''  UBCust.LastTrans = 0
'''''  For LopCnt = 1 To 15
'''''    UBCust.CurrRevAmts(LopCnt) = 0
'''''    UBCust.PrevRevAmts(LopCnt) = 0
'''''  Next
'''''  Put #1, 2, UBCust
'''''  Close
'''''''''''  Dim SIn$, DPos%, LPos%, StrP%, LinLop%
'''''''''''
'''''''''''  Open "SearchResults.txt" For Input As #1
'''''''''''  Open "searchout.txt" For Output As #2
'''''''''''
'''''''''''  Do
'''''''''''    Line Input #1, SIn
'''''''''''    LinLop% = LinLop% + 1
'''''''''''    StrP = 0
'''''''''''    Do
'''''''''''      DPos = InStr(StrP + 1, SIn, "\")
'''''''''''      If DPos > 0 Then
'''''''''''        StrP = DPos
'''''''''''        LPos = DPos
'''''''''''      End If
'''''''''''    Loop While DPos > 0
'''''''''''  SIn = Mid$(SIn$, LPos + 1)
'''''''''''  Print #2, SIn
'''''''''''  Loop While Not EOF(1)
'''''''''''
'''''''''''  Close
'''''''''''  Print LinLop%
''''
'''''find wacky payment
'''''  Open "UBPAY4.dat" For Random As #1 Len = PayRecLen
'''''  NumPayRec = LOF(1) / PayRecLen
'''''  For zzCnt = 1 To NumPayRec
'''''    Get #1, zzCnt, PayRec
'''''    GT1 = 0
'''''    For dzCnt = 1 To 15
'''''      'If PayRec.PaidOwed(dzCnt).AMTPD1 <> PayRec.PaidOwed(dzCnt).AMTOWE1 Then Stop
'''''      GT1 = Round#(GT1 + PayRec.PaidOwed(dzCnt).AMTPD1)
'''''    Next
'''''    GT2 = Round#(PayRec.CASHAMT + PayRec.CHKAMT - PayRec.Change)
'''''    If GT1 <> GT2 Then Stop
'''''    'Print PayRec.CustAcct
'''''    If GT1 <> PayRec.AMTPAID Then Stop
'''''  Next
'''''  Close
''''
''''  'Print NumPayRec
'''''  Open "ubtrans.dat" For Random As #1 Len = TRLen
'''''  RecCnt = LOF(1) / TRLen
'''''  For LopCnt = 1 To RecCnt
'''''    Get #1, LopCnt, TransRec
'''''    If TransRec.TransDate = BDate And TransRec.TransType = 1 Then
'''''      DidCnt = DidCnt + 1
'''''      TransRec.TransDate = GGDate
'''''      Put #1, LopCnt, TransRec
'''''    End If
'''''  Next
'''''  Close
''''
'''''  Open "ubcust.dat" For Random As #1 Len = CustLen
''''
'''''    Get #1, , UBCust
'''''    UBCust.USERCODE2 = "7"
''''    'UBCust.USERCODE1 = "10"
''''
'''''  Put #1, 873, UBCust
''''  'Print UBCust.PrevBalance
''''  'UBCust.DepositAmt = 0
''''  'UBCust.LastTrans = 0
'''''  Put #1, 4001, UBCust
''''
'''''  Close
''''
'''''  Open "ubTrans.dat" For Random As #1 Len = TRLen
'''''  RecCnt = LOF(1) / TRLen
'''''
'''''BEACON
''''
'''''  Open "ubcust.dat" For Random As #1 Len = CustLen
'''''  RecCnt = LOF(1) / CustLen
'''''
'''''  For LopCnt = 1 To RecCnt
'''''     Get #1, LopCnt, UBCust
'''''     UBCust.USERCODE2 = "Z"
'''''     UBCust.USERCODE1 = "1"
'''''     Put #1, LopCnt, UBCust
'''''   Next
''''''''''''
''''''''''''
''''''''''''
''''' Close
''''
''''
''''''    Select Case UBCust.Book
'''''
''''''    Case "01", "02", "03"
'''''
''''''      If Val(QPTrim$(UBCust.USERCODE1)) > 0 Then
'''''        'Stop
''''''      If UBCust.Status = "A" Then
'''''        'If DidCnt Mod 3 = 0 Then
'''''        '  UBCust.USERCODE2 = "J"
'''''        'Else
'''''        'End If
''''''      End If
''''''    Case Else
''''''    End Select
'''''  Next
'''''Close
'''''Label2 = Str$(dzCnt)
'''''Label2 = "Finished"
''''''    For LopCnt = 1 To RecCnt
''''''       TRate = QPTrim(UBCust.serv(zzCnt).Ratecode)
''''''       Select Case TRate
''''''       Case "NRWR", "RWR", "STCW"
''''''         UBCust.serv(10).Ratecode = "HDF"
''''''         Put #1, LopCnt, UBCust
''''''         dzCnt = dzCnt + 1
''''''         Exit For
''''''       End Select
''''''    Next
''''''  Next
''''''
''''''  'Get #1, 1265, UBCust
''''''  'UBCust.CurrBalance = UBCust.PrevBalance
''''''  'UBCust.PrevBalance = 0
''''''  'Put #1, 1265, UBCust
''''''  Close
''''  'MsgBox (Str(dzCnt) + "      " + Str(RecCnt))
''''
'''''  RecCnt = LOF(1) / CustLen
''''
'''''  Open "UB_Cycle_99.txt" For Output As #2
'''''  RecCnt = LOF(1) / CustLen
'''''  'UBCustN.DelFlag = -1
'''''  For LopCnt = 1 To RecCnt
'''''    Get #1, LopCnt, UBCust
'''''    For zzCnt = 1 To 7
'''''      UBCust.LocMeters(zzCnt).ReadFlag = "N"
'''''    Next
'''''    Put #1, LopCnt, UBCust
''''''    If UBCust.BILLCYCL = 99 Then
'''''    'If UBCust.DelFlag = 0 And UBCust.Status = "A" And UBCust.BILLCYCL = 99 Then
'''''      If UBCust.LocMeters(1).ReadFlag = "Y" Then Stop
''''''    End If
''''''    If UBCust.DelFlag = 0 Then
''''''        If Len(QPTrim(UBCust.USERCODE2)) > 0 Then
''''''            UBCust.BILLCYCL = 99
''''''            Print #2, LopCnt, UBCust.Book + "-" + UBCust.SEQNUMB, UBCust.USERCODE1, UBCust.USERCODE2,
''''''            For dzCnt = 1 To 7
''''''              Print #2, UBCust.LocMeters(dzCnt).MtrIDNO, UBCust.LocMeters(dzCnt).MtrNum,
''''''            Next
''''''            Print #2,
''''''            Put #1, LopCnt, UBCust
''''''        Else
''''''            UBCust.BILLCYCL = 0
''''''            Put #1, LopCnt, UBCust
''''''        End If
''''''    End If
''''''   Label2.Caption = "Processing: " + Str(LopCnt) + " of " + Str(RecCnt)
''''''   DoEvents
'''''
''''''SkipEm:
''''''   DidCnt = DidCnt + 1
'''''  Next
'''''  Close
'''''  Label2.Caption = "Processing Complete."
''''  'Label2.Caption = "Changed: " + Str$(DidCnt) + vbCrLf + " of " + Str(RecCnt)
''''
End Sub

Private Sub AllDone()
  If DoneMsgFlag Then
    Command1.Caption = "Already Converted."
  Else
    Label2.Caption = "Done. " ' + Str$(DidCnt) 'DidCnt
    Command1.Caption = "Exit"
  End If
  
  'Command1.Caption = "Done"
  Command1.Enabled = True
  Ok2Exit = True
End Sub

Private Sub MakeBackUp()
  Dim RecCnt As Integer, scnt As Integer
  Dim UBRateTblRecLen As Integer, wcnt As Integer
  Dim UBRateRec As UBRateTblRecType

  DoneMsgFlag = False
  
  UBRateTblRecLen = Len(UBRateRec)
  
  Open "ubrate.dat" For Random As #1 Len = UBRateTblRecLen
  RecCnt = LOF(1) / UBRateTblRecLen

  Open "OldUBRate.dat" For Random As #2 Len = UBRateTblRecLen

  For scnt = 1 To RecCnt
    Get #1, scnt, UBRateRec
    Put #2, scnt, UBRateRec
  Next
  Close
End Sub

Private Sub KeepThis()
 
'  Dim RecCnt As Long
'  Dim LopCnt As Long
'  Dim RCnt As Integer
'  Dim CVal As Double
'  Dim NVal As Double
'
'  Dim RateRecCnt As Integer, scnt As Integer
'  Dim UBRateTblRecLen As Integer, wcnt As Integer
'  Dim UBRateRec As UBRateTblRecType
'
'  DoneMsgFlag = False
'
'  UBRateTblRecLen = Len(UBRateRec)
'
'  Open "OldUBRate.dat" For Random As #1 Len = UBRateTblRecLen
'  RecCnt = LOF(1) / UBRateTblRecLen
'  Close
'  If RecCnt > 0 Then
'    DoneMsgFlag = True
'    GoTo DoneWithIt:
'  Else
'    Call MakeBackUp
'  End If
'
'Skip2Here:
'  Open "ubrate.dat" For Random As #1 Len = UBRateTblRecLen
'  RecCnt = LOF(1) / UBRateTblRecLen
'
'  For LopCnt = 1 To RecCnt
'    Get #1, LopCnt, UBRateRec
'    If InStr(UCase$(UBRateRec.RATEDESC), "SEWER") > 0 Then
'      GoSub DoSewer
'    Else
'      If InStr(UCase$(UBRateRec.RATEDESC), "FLAT") <= 0 Then
'        GoSub DoWater
'      End If
'    End If
'DoNetOne:
'    Put #1, LopCnt, UBRateRec
'  Next
'
'  Close
'  GoTo DoneWithIt
'
'DoSewer:
'    For RCnt = 1 To 10
'      CVal = UBRateRec.TblBreaks(RCnt).UNITAMT
'      If CVal > 0 Then
'        NVal = CVal * 0.03
'        NVal = NVal + CVal
'        UBRateRec.TblBreaks(RCnt).UNITAMT = NVal
'      End If
'    Next
'Return
'
'DoWater:
'    For RCnt = 1 To 10
'      CVal = UBRateRec.TblBreaks(RCnt).UNITAMT
'      If CVal > 0 Then
'        NVal = CVal * 0.06
'        NVal = NVal + CVal
'        UBRateRec.TblBreaks(RCnt).UNITAMT = NVal
'      End If
'    Next
'Return
'
'DoneWithIt:


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

