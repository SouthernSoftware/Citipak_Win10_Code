Attribute VB_Name = "ubStartUp"
Option Explicit
Global Twiddle As String
Global UBPath As String

Type SetUpAcctType
   RevName    As String * 15
   DebitAcct  As String * 14
   CreditAcct As String * 14
End Type

Type RevSetUpType
    RevName As String * 15
    UseDep   As String * 1
    USERATE  As String * 1
    TAXRATE  As Single
    UseMtr   As String * 1
    DistOr   As Integer
    ProRate  As String * 1
End Type

Type UBSetupRecType
    UTILNAME        As String * 35
    DEFCITY         As String * 18
    DEFSTATE        As String * 2
    ZIPCODE         As String * 10
    PreByBook       As String * 1
    RecpPort        As String * 1
    RECPDEFT        As String * 1
    ESTREAD         As String * 1
    BANKDFT         As String * 1
    UseSeq          As String * 1
    BILLCYCL        As String * 1
    DefLook         As String * 1
    MethAcct        As String * 1      'new 02-14-97
    SkipInactive    As String * 1
    SkipSeparator   As String * 1
    Make99File      As String * 1
    LowRead         As Integer
    HighRead        As Integer
    HHDEVICE        As String * 1    'P=PC3000 S=Sensus C=Syscom R=Radix N=None
    Revenues(1 To 15) As RevSetUpType
    BillAcct(1 To 15) As SetUpAcctType
    PayAcct(1 To 15)  As SetUpAcctType
    DepAcct(1 To 15)  As SetUpAcctType
End Type

Type UBCustIndexRecType
  RecNum As Long
End Type


Type ServicesType
    RATECODE As String * 4
    RMtrType As String * 1
End Type

Type FlatRateType
    FRDESC   As String * 18
    FRAMT    As Double
    FRFREQ   As String * 1
    REVSRC   As Integer
    NumMin   As Integer
End Type

Type RevDataType
    RevName    As String * 20
    RATECODE   As String * 4
    RevMtrType As String * 1
End Type
Type LocMeterType
    MtrNum    As String * 12
    MTRMulti  As Integer
    MtrType   As String * 1
    MtrUnit   As String * 1
    NumUser   As Integer
    InsDate   As Integer
    CurRead   As Long
    PrevRead  As Long
    CurDate   As Integer
    PastDate  As Integer       'hidden & protected
    ReadFlag  As String * 1    'hidden & protected
    AvgUse    As Long          'hidden & protected
    UseCnt    As Integer       'hidden & protected
    MtrIDNO   As String * 11
    MtrLat    As Double
    MtrLng    As Double
End Type

Type MonthlyPayType
    AMTOWED      As Double
    TotAmtPD     As Double
    PayAmt       As Double
    RevSource    As Integer
End Type


Type NewUBCustRecType
    Book          As String * 2
    SEQNUMB       As String * 6
    Status        As String * 1
    OPENDATE      As Integer
    SEARCH        As String * 10
    CustName      As String * 35
    ADDR1         As String * 35
    ADDR2         As String * 35
    SERVADDR      As String * 35
    CITY          As String * 18
    STATE         As String * 2
    ZIPCODE       As String * 10
    HPHONE        As String * 14
    WPHONE        As String * 14
    SOSEC         As String * 11
    DRVLIC        As String * 16
    CUSTTYPE      As String * 3
    Addr911       As String * 14
'051498 added bill to field. Removed 1 byte from 911 addr
    BillTo        As String * 1
'********************************************************
    BILLCOPY      As Integer
    POSTRTE       As String * 4
    BILLCYCL      As Integer
    ZONE          As String * 3
    Seq           As Long
'Page 2
    CASHONLY      As String * 1
    LATEFEE       As String * 1
    CUTOFFYN      As String * 1
    TAXEXPT       As String * 1
    SRCIT         As String * 1
    EPPFlag       As String * 1
'032299 Modified for Bank draft account type
'    EPPAMT        AS DOUBLE
'added GroupCoderec 2/1/05 for pointer to bookcode
    GroupCodeRec  As Integer
    Filler1       As String * 5
   ' Filler1       As String * 7
    USEDRAFT      As String * 1
    AcctType      As String * 1
'032299 Inserted account type
    BankName      As String * 34
    BANKLOC       As String * 30
    TRANSIT       As String * 9
    BankAcct      As String * 20
    BILLCMNT      As String * 25
    PAYCMNT       As String * 25
    PumpCode      As String * 4
    USERCODE1     As String * 4
    USERCODE2     As String * 2
    ProRatePCT    As Integer
    HHMSG1        As String * 20
    HHMSG2        As String * 20
    HHMSG3        As String * 20
'Page 3
    serv(1 To 15)      As ServicesType
    FlatRates(1 To 4)  As FlatRateType
'Page 4
    Monthly(1 To 2)    As MonthlyPayType
    MFEE1         As Double
    MFEE2         As Double
    LocMeters(1 To 7)  As LocMeterType
'END OF Quick Screen Form
    CustPIN       As Long
    LastTrans     As Long
    CurrBalance   As Double
    PrevBalance   As Double
    CurrRevAmts(1 To 15) As Double
    PrevRevAmts(1 To 15) As Double
    DepositAmt    As Double
    DelFlag       As Integer
    PreNoteFlag   As Integer
    WOLastTrans   As Long            'work order last trans pointer
    EstFlag       As String * 1
    MessageRec    As Long            ' Points to Message Record
    OldRec        As Long
    EPPLastTran   As Long
    NewNotes      As Integer
    DPCode        As String * 2
    FillPad       As String * 112
    ChkByte       As String * 1
End Type
'Type GroupCodeIndexType
'    RecordNum   As Integer
'    GroupCODE   As String * 2
'End Type
Type UBTransRecType
   TransDate              As Integer      '
   TransType              As Integer      '
   TransDesc              As String * 21  'may change
   Transamt               As Double       'total revenue amount
   RevAmt(1 To 15)        As Double       'Revenue amounts
   TaxAmt(1 To 15)        As Single       'Tax Amounts
'01-20-97 Added meter types field to hold meter type at time of transaction
   MtrTypes(1 To 7)       As Integer
'*******************
   CurRead(1 To 7)        As Long         'Last/Current meter readings
   PrevRead(1 To 7)       As Long         'Previous readings
   ESTREAD(1 To 7)        As String * 1   'Y/N Flags for meter est's
   BillNumber             As Long         'Number on the bill that Printed
   ReadDate               As Integer
   BillDate               As Integer
   PastDueDate            As Integer
   DraftDate              As Integer      '
'111398
   ProRatePCT             As Integer
   ChkByte                As String * 1   'Added check byte
   EPPFlag                As String * 1   'Equal Payment Flag
   CustStatus             As String * 1   'Customer Status at Time of Transaction
'020199
   EPPTrans               As Long         'Pointer to Equal Pay trans
   PenAtBill              As Single       'Used to flag IRR Meter (Sunset)
'****************
   PayTypeCode            As Integer      'Payment Type:  1=Cash, 2=Check, 3=Cash/Check, 4=Charge
   OperatorNumber         As Integer      '
   CustAcctNo             As Long         'Pointer to RecNo in ubcust.dat
   PrevTrans              As Long
   VoidFlag               As Integer       'Changed for wadesboro
   FromCMFlag             As Integer
   ActiveFlag             As Integer      'Valid transaction flag
   RunBalance             As Double
   CheckAmount            As Double
   CashAmount             As Double
   BillMsg                As String * 20
   ApplyDepFlag           As String * 1
   Posted2GL              As String * 1
   PrevDate               As Integer
   PenalFlag              As String * 1
   TaxExempt              As String * 1
   NONProfit              As String * 1
End Type

Type GroupCodeRecType
    Deleted       As Integer
    GroupCODE     As String * 2
    GroupCodeName As String * 30
    xtrastuff     As String * 30
End Type

Sub Main()
  Dim RetValue As Integer
  Dim RecLen As Integer
  Twiddle = "||//--\\"
    
'  App.TaskVisible = False        'don't show in task list
  UBPath$ = QPTrim$(App.Path)    'start up path
  
  If Right$(UBPath$, 1) <> "\" Then
    UBPath$ = UBPath$ + "\"
  End If
  ExpCustStuff
End Sub

Public Sub DoTheTime()
  Dim sec As Long
  sec = Timer
  Do
  Loop Until (sec + 1) < Timer
End Sub

Private Sub ExpCustStuff()
  Dim Dash80 As String, IndexName As String, IdxRecLen As Integer
  Dim UBCustRecLen As Integer, UBCust As Integer, UBTran As Integer
  Dim IdxFileSize As Long, IdxNumOfRecs As Integer, cnt As Long
  Dim UBTranRecLen As Integer, NumOfRecs As Long, NumOfCust As Long
  Dim Handle As Integer, UsingBook As Boolean, NumOfPeriods As Integer
  Dim RecNo As Long, DidCnt As Long, ThisTrans As Long, FMonth As Integer
  Dim FYear As Integer, TYear As Integer, TMonth As Integer
  Dim FMCnt As Integer, DidAMeter As Boolean, MtrCnt As Integer
  Dim MeterType As String, MeterConsp As Long, MaxMeterAmt As Long
  Dim TotalConsump As Long, QPos As Integer, LocationNumber As String
  Dim Zip As String, CCCnt As Long, NumofRevs As Integer, BuckFmt As String
  Dim Bookone As Integer, Bookto As Integer, qc As String, ThisBook As String
  Dim q As String, C As String, qcq As String, OKFlag As Boolean
  Dim UBRpt As String, zz As String, zzN As Integer, CCnt As Long
  Dim UBOwnerRecLen As Integer, UBFile As Integer, AcctNumber As Long
  Dim WhatBook As Integer, Export As Long, RCnt As Integer, FCnt As Integer
  Dim MCnt As Integer, tempTot As Double, ThisFile As Integer, GetString As String
  Dim UBTrnRec As UBTransRecType
  Dim UBTrnRecLen As Integer, UBTrnCnt As Long, PrevTranRec As Long
  Dim blnGotDate As Boolean, DueDate As String
  
  On Error GoTo endstuff
  FrmShowPctComp.Label1 = "Creating Export Files"
  FrmShowPctComp.Show
  GetString$ = ""

'  q$ = Chr$(34)
'  qc$ = q$ + ","
'  qcq$ = q$ + "," + q$
  'special for online file
  If Exist(UBPath$ + "UBOutSet.txt") Then
     ThisFile = FreeFile
     Open "UBOutSet.txt" For Input As ThisFile
        Line Input #ThisFile, GetString$
        GetString$ = RTrim$(GetString$)
     Close ThisFile
  End If
  qcq$ = "|"
  OKFlag = True
  BuckFmt$ = "########.##"
  NumofRevs = GetNumOfRevs%
  UBTrnRecLen = Len(UBTrnRec)
  
  UBTrnCnt = FileSize(UBPath$ + "UBTRANS.DAT") \ UBTrnRecLen
  
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  NumOfRecs = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  KillFile (UBPath$ + "UBOLFile.txt")
  UBRpt = FreeFile
  Open UBPath$ + "UBOLFile.txt" For Output As UBRpt
  
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As #20 Len = UBTrnRecLen
  
  For cnt = 1 To NumOfRecs
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    Get UBCust, cnt, UBCustRec(1)
    'Change for Hemingway to include inactive accounts 6/8/2015
    'If UBCustRec(1).DelFlag <> -1 And Not UBCustRec(1).Status = "I" Then
    If UBCustRec(1).DelFlag <> -1 Then
      If (UBCustRec(1).Status = "I" And (Round(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) > 0)) Or UBCustRec(1).Status <> "I" Then
        Export& = Export& + 1
        AcctNumber = cnt
        GoSub ExportThisAccount
      End If
    End If
  Next

  Close
  If Export& > 0 Then
    'MsgBox "File " & UBPath$ & "UBOnLine.txt Exported with " & Export& & " Accounts.", vbOKOnly, "Export Completed."
  Else
    GoTo endstuff
    'MsgBox "No Information Found to Export.", vbOKOnly, "Procedure Ended"
  End If
GoTo ExitMastCustListing

ExportThisAccount:

  LocationNumber$ = QPTrim$(UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB)
  Zip$ = QPTrim$(UBCustRec(1).ZIPCODE)
'  If Len(Zip$) > 5 Then
'    Zip$ = Left$(Zip$, 5) + "-" + Mid$(Zip$, 6)
'  End If
  
  PrevTranRec& = UBCustRec(1).LastTrans
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      Get 20, PrevTranRec&, UBTrnRec
      Select Case UBTrnRec.TransType
      Case 1, 101
        blnGotDate = True
        DueDate = Num2Date(UBTrnRec.PastDueDate)
        Exit Do
      End Select
      PrevTranRec& = UBTrnRec.PrevTrans
    Loop
  End If

  If Len(GetString$) > 0 Then
    Print #UBRpt, GetString$; qcq$;
  End If
  If Not blnGotDate Then
    DueDate = "??/??/????"
  End If
  
  Print #UBRpt, QPTrim$(Str$(AcctNumber));
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).SEARCH);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).CustName);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).SERVADDR);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).HPHONE);
  Print #UBRpt, qcq$; Using(BuckFmt$, UBCustRec(1).CurrBalance);
  Print #UBRpt, qcq$; Using(BuckFmt$, UBCustRec(1).PrevBalance);
  Print #UBRpt, qcq$; Using(BuckFmt$, Round(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance));

  For RCnt = 1 To NumofRevs
    Print #UBRpt, qcq$; Using(BuckFmt$, UBCustRec(1).CurrRevAmts(RCnt));
  Next
  For RCnt = (NumofRevs + 1) To 15
    Print #UBRpt, qcq$; "0";
  Next
'meters
  For MCnt = 1 To 7
    If Len(QPTrim$(UBCustRec(1).LocMeters(MCnt).MtrNum)) > 0 Then
      Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).LocMeters(MCnt).MtrNum);
    Else
      Print #UBRpt, qcq$; " ";
    End If
    
    If UBCustRec(1).LocMeters(MCnt).CurRead > 0 Then
      Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).LocMeters(MCnt).CurRead));
    Else
      Print #UBRpt, qcq$; "0";
    End If

    If UBCustRec(1).LocMeters(MCnt).PrevRead > 0 Then
      Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).LocMeters(MCnt).PrevRead));
    Else
      Print #UBRpt, qcq$; "0";
    End If

    If UBCustRec(1).LocMeters(MCnt).CurDate > 0 Then
      Print #UBRpt, qcq$; Num2Date$(UBCustRec(1).LocMeters(MCnt).CurDate);
    Else
      Print #UBRpt, qcq$; " ";
    End If
  Next
  Print #UBRpt, qcq$; DueDate;
  Print #UBRpt, qcq$; QPTrim$(Zip$)
  
'  Print #UBRpt,

Return

endstuff:
  CitiTerminate
ExitMastCustListing:

End Sub

Public Function GetNumOfRevs%()
  Dim UBSetupLen As Integer, NumofRevs As Integer, Handle As Integer
  Dim RevCnt As Integer, TempRev As String
  NumofRevs = 15
  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupLen = Len(UBSetUp(1))
  On Local Error Resume Next
'  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetupLen, 1            'load it
'  Handle = FreeFile
'  Open UBPath$ + "UBSETUP.DAT" For Random Shared As Handle Len = UBSetupLen    'open data file
'  Get #Handle, 1, UBSetUpRec(1)
    LoadUBSetUpFile UBSetUp(), UBSetupLen
'this doesn't work properly if they skip around in revenue setup list
  For RevCnt = 1 To 15
    TempRev$ = QPTrim$(UBSetUp(1).Revenues(RevCnt).RevName)
    If Len(TempRev$) = 0 Then
      NumofRevs = RevCnt - 1
      Exit For
    End If
  Next
  GetNumOfRevs = NumofRevs
  Erase UBSetUp
End Function
Public Sub LoadUBSetUpFile(UBSetUpRec() As UBSetupRecType, UBSetupLen)
  Dim Handle As Integer
  On Local Error Resume Next
  UBSetupLen = Len(UBSetUpRec(1))            'use the length as an error flag
  If Exist(UBPath$ + "UBSETUP.DAT") Then
    Handle = FreeFile
    Open UBPath$ + "UBSETUP.DAT" For Random Shared As Handle Len = UBSetupLen    'open data file
    If LOF(Handle) > 0 Then
      Get #Handle, 1, UBSetUpRec(1)
    End If
    Close Handle
  End If
End Sub

Public Function Round#(ByVal N#)
  Round# = (Int(N# * 100 + 0.5000001)) / 100
End Function
Public Static Function Using$(ByVal fmt As String, ByVal Number As Double, Optional LeadZeroFlag As Boolean)
  Dim TempNumber As String
  Dim FmtNumber As String
  Dim TempLen As Integer
  Dim BuckPos As Integer, FmtLen As Integer
  FmtLen = Len(fmt)
  BuckPos = InStr(fmt, "$")
  If BuckPos = 1 Then
    fmt = Right$(fmt, FmtLen - 1)
  ElseIf BuckPos > 1 Then
    fmt = Left$(fmt, BuckPos - 1) + Mid$(fmt, BuckPos + 1)
  End If
  FmtNumber = Space$(Len(fmt))
  TempNumber = Format(Number, fmt)
  TempLen = Len(TempNumber)
  If TempLen = 0 Then
    TempNumber = "0"
    GoTo GotZero
  End If
  If TempLen >= 2 Then
    If Mid$(TempNumber, (TempLen - 1), 1) = "." Then
      TempNumber = TempNumber + "0"
    End If
  End If
  If Right$(TempNumber, 1) = "." Then
    TempNumber = TempNumber + "00"
  End If
  If BuckPos > 0 Then
    TempNumber = "$" + TempNumber
  End If
GotZero:
  If LeadZeroFlag Then
    If TempNumber = ".00" Then
      TempNumber = "0.00"
    End If
  End If
  
  RSet FmtNumber = TempNumber
  
  Using = FmtNumber
  
'Number = 5: Fmt = "$##,##0.00": Print Right(String(Len(Fmt), " ") & Format(Number, Fmt), Len(Fmt))
End Function

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
Public Function FileSize(FileName$) As Long
  On Local Error Resume Next
  Dim FileHandle As Integer
  If Exist(FileName$) Then
    FileHandle = FreeFile
    Open FileName$ For Binary As FileHandle
    FileSize = LOF(FileHandle)
    Close FileHandle
  Else
    FileSize = 0
  End If
End Function

Public Function Exist(FileName$)
  On Local Error Resume Next
  Dim FileHandle As Integer
  Dim FileSize As Long
  FileHandle = FreeFile
  Open FileName$ For Binary As FileHandle
'  If Err Then
'    FileName$ = ""
'  End If
  FileSize = LOF(FileHandle)
  Close FileHandle
  If FileSize > 0 Then
    Exist = True
  Else
    Exist = False
    Kill FileName$
  End If
'  On Local Error GoTo 0
End Function

Public Sub KillFile(FileName$)
  If Exist(FileName) Then
    Kill FileName$
  End If
End Sub
Public Function Num2Date$(intDate%)
  On Error GoTo BadNum2Date
  If intDate% = -32767 Then
    Num2Date$ = ""
  Else
    Num2Date$ = Format(DateAdd("d", (intDate%), "12-31-1979"), "mm/dd/yyyy")
  End If
  Exit Function
BadNum2Date:
  On Error GoTo 0
  Num2Date = ""
End Function

Public Sub CitiTerminate()
  Dim UBFrmCnt As Integer
  ' Loop through the forms collection and unload each form.
  DoEvents
  For UBFrmCnt = Forms.Count - 1 To 0 Step -1
    Unload Forms(UBFrmCnt)
  Next
  DoEvents
  End
End Sub

