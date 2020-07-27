Attribute VB_Name = "zmodsavestuff"
')%*$)%(*$)%*$)%*$)(%*)$%*$)%*)$(*%)$*%
'''Private Sub PreBillReport234() 'this is graphic report
'''  Dim Temp2 As String, NumOfRevs As Integer, NumOfRates As Integer
'''  Dim UBRateTblRecLen As Integer, RateFile As Integer, cnt As Long
'''  Dim UBSetupLen As Integer, MowFlag As Boolean, TennFlag As Boolean
'''  Dim TempRev As String, DoFuelAdjFlag As Boolean, SkipInactive As Boolean
'''  Dim SkipSeparator As Boolean, ThisBook As Integer, BookNum As Integer
'''  Dim BookFlag As Boolean, ThisCycle As Integer, CycleFlag As Boolean
'''  Dim SeqFlag As String, Choice As Integer, FuelAdjAmt As Double
'''  Dim IndexName As String, UsingAcct As Boolean, IdxTypeText As String
'''  Dim AbortFlag As Boolean, TheDate As String, UBCustRecLen As Integer
'''  Dim UBBillRecLen As Integer, TBooks As Integer, NumOfRecs As Long
'''  Dim Handle As Integer, IdxRecLen As Integer, lcnt As Long
'''  Dim UBBill As Integer, UBCust As Integer, UBRpt As Integer
'''  Dim ThisCustRec As Long, BillTo As String, BadBookFlag As Boolean
'''  Dim WhatBook As Integer, FRCnt As Integer, WhatService As Integer
'''  Dim Multi As Integer, FlatAmt As Double, WhatRate As Integer
'''  Dim DoneOne As Boolean, TRevCnt As Integer, IFlag As Boolean
'''  Dim TRateCnt As Integer, MINAMT As Long, PrintedRevAmt As Boolean
'''  Dim MCCnt As Integer, CubMtr As Boolean, LocMeterType As String
'''  Dim MeterMulti As Long, MeterNum As String, ReadAmt As Long
'''  Dim MaxMeterAmt As Long, Consump As Long, ThisMeterUseCnt As Integer
'''  Dim AvgUse As Long, HiConsump As Long, LowConsump As Long
'''  Dim TTRevCnt As Integer, CurReadAmt As Long, PreReadAmt As Long
'''  Dim ConsumpFlag As Boolean, ConsumpAmt As Long, NONRateCnt As Integer
'''  Dim NONRate As Integer, CTaxAmt As Double, TXCnt As Integer
'''  Dim Bills2Print As Integer, AcctBalance As Double, WhatPump As Integer
'''  Dim TAcctBalance As Double, HasAPumpCode As Boolean, MPCnt As Integer
'''  Dim PumpMtrOK As Boolean, TotalFlatAmt As Double, TotalRevAmt As Double
'''  Dim TotalTaxAmt As Double, RaCnt As Integer, TestTot As Double
'''  Dim ZCnt As Integer, Book As String, TBookAmt As Double, TPumps As Integer
'''  Dim TBTaxAmt As Double, RCnt As Integer, TBookGTot As Double
'''  Dim TMMConsump As Double, RptText As String, TBCnt As Integer
'''  Dim CustPump As String, ThisPump As String, ToPrintM As String
'''  Dim ToPrintN As String, ToPrintR(1 To 15) As String, ToPrintT As String
'''  Dim ToPrintX As String, UBRPTG As Integer, UBRPTR As Integer
'''  Dim UBRPTB As Integer, ToPrintS As String, UBRPTS As Integer
'''  Dim UBRPTP As Integer, ToPrint As String, ThisCnt As Integer
'''  Dim PrintOK As Boolean
'''  UBLog "IN: Prebilling Report"
'''
'''  PageNo = 0
'''  Temp2$ = Space$(12)
'''  NumOfRevs = MaxRevsCnt        'assume max munber of revenue sources
'''  NumOfRates = GetNumRateRecs%
'''  ReDim UBRateTbls(1 To NumOfRates) As UBRateTblRecType
'''  UBRateTblRecLen = Len(UBRateTbls(1))
'''
'''  ReDim RateConsump(1 To NumOfRates) As Double
'''
'''  RateFile = FreeFile
'''  Open "UBRATE.DAT" For Random Shared As RateFile Len = UBRateTblRecLen
'''  For cnt = 1 To NumOfRates
'''    Get RateFile, cnt, UBRateTbls(cnt)
'''  Next
'''  Close
'''
'''  'SortT UBRateTbls(1), NumOfRates, 0, UBRateTblRecLen, 0, 4
'''  RateQSort UBRateTbls(), 1, NumOfRates
'''
'''  ReDim ProrateServ(1 To 15) As Integer
'''
'''  ReDim UBSetUpRec(1) As UBSetupRecType
'''  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
'''
'''  TownName$ = UBSetUpRec(1).UTILNAME
'''  If InStr(TownName$, "MOWAS") > 0 Then
'''    MowFlag = True
'''  End If
'''
'''  If UBSetUpRec(1).DEFSTATE = "TN" Then
'''    TennFlag = True
'''  End If
'''
'''  ReDim RevDesc(1 To MaxRevsCnt) As String * 12
'''  For cnt = 1 To MaxRevsCnt     'find last active revenue
'''    TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(cnt).REVNAME)
'''    If Len(TempRev$) = 0 Then
'''      NumOfRevs = cnt - 1       'set actual number of revenues
'''      Exit For
'''    Else        'build revenue description lines
'''      LSet RevDesc(cnt) = UCase$(TempRev$)
'''      If InStr(RevDesc(cnt), "ELECTRIC") Then
'''        DoFuelAdjFlag = True
'''      End If
'''    End If
'''  Next
'''  '111398 Prorate
'''  For cnt = 1 To MaxRevsCnt
'''    If UBSetUpRec(1).Revenues(cnt).ProRate = "Y" Then
'''      ProrateServ(cnt) = True
'''    End If
'''  Next
'''
'''  If UBSetUpRec(1).SkipInactive = "Y" Then
'''    SkipInactive = True
'''  End If
'''
'''  If UBSetUpRec(1).SkipSeparator = "Y" Then
'''    SkipSeparator = True
'''  End If
'''
'''  If UBSetUpRec(1).PreByBook = "Y" Then
'''    ThisBook = Val(fptxtRoute1)
'''    If ThisBook = 99 Then
'''      ThisBook = -1
'''    End If
'''    BookNum = ThisBook
'''    If ThisBook = -1 Then
'''      BookFlag = False
'''    ElseIf ThisBook <= 0 Then
'''      GoTo ExitPreReport
'''    Else
'''      BookFlag = True
'''    End If
'''  ElseIf UBSetUpRec(1).BILLCYCL = "Y" Then
'''    ThisCycle = Val(fptxtRoute1)
'''    If ThisCycle <= 0 Then
'''      GoTo ExitPreReport
'''    Else
'''      CycleFlag = True
'''    End If
'''  End If
'''
'''  If UBSetUpRec(1).UseSeq = "Y" Then
'''    SeqFlag$ = "Y"
'''  End If
'''  FrmShowPctComp.Label1 = "Creating PreBilling Report"
'''  FrmShowPctComp.Show , Me
'''
'''Restart:
'''  Choice = fpcboPrintOrder.ListIndex + 1
'''  'GetPreBillOrder Choice, ExitFlag, SeqFlag$
'''
''' 'If ExitFlag Then GoTo ExitPreReport
'''  If DoFuelAdjFlag Then
'''    FuelAdjAmt# = Val(fptxtAdjustment)
'''    UBLog "Fuel adjustment factor:" + Str$(FuelAdjAmt#)
'''  Else
'''    FuelAdjAmt# = 0
'''  End If
'''
'''  If FuelAdjAmt# = -10000 Then GoTo Restart
'''
'''  Select Case Choice
'''  Case 0
'''    'ExitFlag = True
'''  Case 1        'Name
'''    IndexName$ = NameIndexFile
'''    'OkFlag = True
'''  Case 2        'Acct
'''    IndexName$ = ""
'''    UsingAcct = True
'''    'OkFlag = True
'''  Case 3        'Location
'''    IndexName$ = BookIndexFile
'''    'OkFlag = True
'''  Case 4        'Postal Route
'''    IdxTypeText$ = "Postal Route"
'''    MakePostalIndex IdxTypeText$
'''    IndexName$ = TempIndexName
'''    'OkFlag = True
'''  Case 5        'ZipCode
'''    IdxTypeText$ = "Zip-Code"
'''    'this mowflag for zip index doesn't matter cause both index
'''    'routines do same thing now.
'''    If MowFlag Then
'''      MakeMowZipCodeIndex IdxTypeText$
'''    Else
'''      MakeZipCodeIndex IdxTypeText$
'''    End If
'''    IndexName$ = TempIndexName
'''    'OkFlag = True
'''  Case 6        'Sequence number
'''    IdxTypeText$ = "Sequence Number"
'''    MakeSequenceIndex IdxTypeText$, Me
'''    IndexName$ = TempIndexName
'''    'OkFlag = True
'''  End Select
'''  MakeBillFile AbortFlag, FuelAdjAmt#, ThisCycle, ThisBook
'''
'''  If AbortFlag Then GoTo ExitPreReport
'''
'''  MaxLines = 53
'''
'''  ReDim fmt$(0 To 6)
'''  fmt$(0) = String$(80, "-")
'''  fmt$(1) = "#########.##"
'''  fmt$(2) = "#########"
'''  fmt$(3) = "######.##"
'''  fmt$(4) = "###########"
'''  fmt$(5) = "$###,###,###.##"
'''  fmt$(6) = "$#,###,###.##"
'''
'''  TheDate$ = "Date: " + Date$
'''
'''  ReDim UBCustRec(1 To 2) As NewUBCustRecType
'''  UBCustRecLen = Len(UBCustRec(1))
'''
'''  ReDim UBBillRec(1) As UBTransRecType
'''  UBBillRecLen = Len(UBBillRec(1))
'''
'''  ReDim FlatTotals(1 To NumOfRevs) As Double
'''  '021998 added flat revenue totals
'''  ReDim RevTotals(1 To NumOfRevs) As Double     'Revenue total amts
'''  '052097 added tax by revenue totals
'''  ReDim TaxTotals(1 To NumOfRevs) As Double     'Tax total amts
'''  ReDim ConsumpTot(1 To NumOfRevs, 1 To 2) As Double            'Consumption total amts
'''  ReDim RateConsump(1 To NumOfRates) As Double
'''  '012698 Added bill count by rate code
'''  ReDim RateCount(1 To NumOfRates) As Long
'''  ReDim RateTotals(1 To NumOfRates) As Double   'Rates total amts
'''  '052097 added tax by rate code totals
'''  ReDim RTaxTot(1 To NumOfRates) As Double      'Rates Tax total amts
'''  '052097 added tax by book totals to type def
'''  ReDim Bookconsump(0 To 1) As BookConsumpType  'Consumption by book
'''  ReDim PumpConsump(0 To 1) As PumpConsumpType  'Consumption by pump code
'''  ReDim TaxExmp(0 To NumOfRevs) As Double
'''
'''  TBooks = 0
'''  If UsingAcct Then
'''    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
'''  Else          'load the index
'''    UBLog "Loading index file: " + IndexName$
'''    IdxRecLen = 4
'''    NumOfRecs = FileSize(IndexName$) \ 4
'''    ReDim Indexarray(1 To NumOfRecs) As UBCustIndexRecType
'''    Handle = FreeFile
'''    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
'''    For lcnt& = 1 To NumOfRecs
'''      Get #Handle, lcnt&, Indexarray(lcnt&)
'''    Next
'''    Close Handle
'''    'FGetAH IndexName$, IndexArray(1), 4, NumOfRecs
'''  End If
'''
'''  UBBill = FreeFile
'''  Open UBBillsFile For Random Shared As UBBill Len = UBBillRecLen
'''  UBCust = FreeFile
'''  Open "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
'''  UBRpt = FreeFile
'''  Open "UBPREBIL.RPT" For Output As UBRpt
'''  UBRPTG = FreeFile
'''  Open "UBPREGT.RPT" For Output As UBRPTG
'''  UBRPTR = FreeFile
'''  Open "UBPRERT.RPT" For Output As UBRPTR
'''  UBRPTB = FreeFile
'''  Open "UBPREB.RPT" For Output As UBRPTB
'''  UBRPTS = FreeFile
'''  Open "UBPRES.RPT" For Output As UBRPTS
'''  UBRPTP = FreeFile
'''  Open "UBPREP.RPT" For Output As UBRPTP
'''  'BlockClear
'''  'ShowProcessingScrn "Processing Pre-Billing Report"
'''  UBLog "Writing prebilling report to disk."
'''
'''  GoSub PrintPreHeader
'''  For cnt = 1 To NumOfRecs
'''    PrintOK = False
'''    If UsingAcct Then
'''      ThisCustRec& = cnt
'''    Else
'''      ThisCustRec& = Indexarray(cnt).RecNum
'''    End If
'''
'''    Get UBCust, ThisCustRec&, UBCustRec(1)
'''
'''    If UBCustRec(1).DelFlag Then
'''      GoTo SkipEm
'''    End If
'''
'''    If SkipInactive And UBCustRec(1).Status <> "A" Then
'''      GoTo SkipEm
'''    ElseIf UBCustRec(1).Status = "F" Then       'skip over final's
'''      GoTo SkipEm
'''    ElseIf UBCustRec(1).Status = "B" Then       'skip over B-Status
'''      GoTo SkipEm
'''    End If
'''    If BookFlag Then
'''      If Val(UBCustRec(1).Book) <> ThisBook Then
'''        GoTo SkipEm
'''      End If
'''    End If
'''
'''    If CycleFlag Then
'''      If UBCustRec(1).BILLCYCL <> ThisCycle Then
'''        GoTo SkipEm
'''      End If
'''    End If
'''
'''    Get UBBill, ThisCustRec&, UBBillRec(1)
'''
'''    If Linecnt > MaxLines Then
'''      'Print #UBRpt, FF$
'''      'GoSub PrintPreHeader
'''    End If
'''
'''    If UBBillRec(1).ActiveFlag <> 0 Then
'''      If UBCustRec(1).BillTo = "O" Then
'''        BillTo$ = " O"
'''      Else
'''        BillTo$ = " C"
'''      End If
'''      GoSub GetWhatBook
'''      If BadBookFlag Then
'''        If ErrorScrn(2, ThisCustRec&) Then
'''          AbortFlag = True
'''          Exit For
'''        End If
'''      End If
'''      Bookconsump(WhatBook).CustCnt = Bookconsump(WhatBook).CustCnt + 1
'''      ToPrintN$ = UBCustRec(1).Status + "~" + Using("  #####  ", ThisCustRec&)
'''      ToPrintN$ = ToPrintN$ + "~" + UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB + "~" + Left$(UBCustRec(1).CustName, 25) + "~" + Left$(UBCustRec(1).SERVADDR, 22)
'''      ToPrintN$ = ToPrintN$ + "~" + Using("   ###", UBBillRec(1).ProRatePCT) + "%"
'''      ToPrintN$ = ToPrintN$ + "~" + BillTo$
'''      PrintOK = True
'''      'Linecnt = Linecnt + 1
'''      For FRCnt = 1 To 4
'''        WhatService = UBCustRec(1).FlatRates(FRCnt).REVSRC
'''        If UBCustRec(1).FlatRates(FRCnt).FRAMT <> 0 And WhatService > 0 Then
'''          Multi = UBCustRec(1).FlatRates(FRCnt).NumMin
'''          If Multi < 1 Then Multi = 1
'''          FlatAmt# = Round#(UBCustRec(1).FlatRates(FRCnt).FRAMT * Multi)
'''          '021998 Added flat rate summaries
'''          FlatTotals(WhatService) = Round#(FlatTotals(WhatService) + FlatAmt#)
'''        End If
'''      Next
'''      '102798 Added to skip accts that don't have a book/seq no. "J.R."
'''    ElseIf Len(QPTrim$(UBCustRec(1).Book)) = 0 And Len(QPTrim$(UBCustRec(1).SEQNUMB)) = 0 Then
'''      GoTo SkipEm
''''      Else
''''      Stop
'''    End If
'''    WhatRate = 0
'''    DoneOne = False
'''    For TRevCnt = 1 To NumOfRevs
'''      If TRevCnt = 2 And UBBillRec(1).PenAtBill = -1 Then
'''        IFlag = True
'''      Else
'''        IFlag = False
'''      End If
'''      WhatRate = 0
'''      If UBBillRec(1).RevAmt(TRevCnt) <> 0 Then
'''        DoneOne = False
'''        ToPrintR$(TRevCnt) = RevDesc(TRevCnt)
'''        '102198 Moved out of meter loop, Stoped multi meter tax report bug
'''        If UBBillRec(1).TaxAmt(TRevCnt) > 0 Then
'''          TaxTotals(TRevCnt) = Round#(TaxTotals(TRevCnt) + UBBillRec(1).TaxAmt(TRevCnt))
'''        End If
'''        For TRateCnt = 1 To NumOfRates
'''          If UBRateTbls(TRateCnt).RATECODE = UBCustRec(1).Serv(TRevCnt).RATECODE Then
'''            MINAMT& = UBRateTbls(TRateCnt).MINUNITS
'''            WhatRate = TRateCnt
'''            '102198 Moved from meter loop, Stops multi meter tax report bug
'''            RTaxTot(WhatRate) = Round#(RTaxTot(WhatRate) + UBBillRec(1).TaxAmt(TRevCnt))
'''            Exit For
'''          End If
'''        Next
'''        If UBSetUpRec(1).Revenues(TRevCnt).UseMtr = "Y" Then
'''          RevTotals(TRevCnt) = Round#(RevTotals(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
'''          Bookconsump(WhatBook).RevAmt(TRevCnt) = Round#(Bookconsump(WhatBook).RevAmt(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
'''          Bookconsump(WhatBook).TaxAmt(TRevCnt) = Round#(Bookconsump(WhatBook).TaxAmt(TRevCnt) + UBBillRec(1).TaxAmt(TRevCnt))
'''
'''          '02-20-97 Add revenue totals by rate code
'''          If WhatRate > 0 Then
'''            RateTotals(WhatRate) = Round#(RateTotals(WhatRate) + UBBillRec(1).RevAmt(TRevCnt))
'''          End If
'''          PrintedRevAmt = False
'''          For MCCnt = 1 To 7
'''            CubMtr = False
'''            LocMeterType$ = QPTrim$(UBCustRec(1).LocMeters(MCCnt).MTRType)
'''            MeterMulti& = UBCustRec(1).LocMeters(MCCnt).MTRMulti
'''            '063098 Added adjustment for cubic meters in consumption totals
'''            If UBCustRec(1).LocMeters(MCCnt).MTRUnit = "C" Then
'''              CubMtr = True
'''            End If
'''            If MeterMulti& <= 0 Then MeterMulti& = 1
'''            If (Len(LocMeterType$) > 0) And (LocMeterType$ = UBCustRec(1).Serv(TRevCnt).RMtrType) Then
'''              DoneOne = True
'''              MeterNum$ = QPTrim$(UBCustRec(1).Serv(TRevCnt).RATECODE)
'''              'use the Meternum$ to hold the rate code temporarily
'''              If Len(MeterNum$) > 0 Then
'''                If UBCustRec(1).LocMeters(MCCnt).NumUser > 1 Then
'''                  MeterNum$ = MeterNum$ + "*" + QPTrim$(Str$(UBCustRec(1).LocMeters(MCCnt).NumUser))
'''                  'ToPrintR$ = ToPrintR$ + "~" + QPTrim$(Str$(UBCustRec(1).LocMeters(MCCnt).NumUser))
'''                End If
'''                RSet Temp2$ = MeterNum$
'''              End If
'''              ReadAmt& = UBBillRec(1).CurRead(MCCnt) - UBBillRec(1).PrevRead(MCCnt)
'''              If ReadAmt& < 0 Then              'Meter rolled over or, been misread
'''                MaxMeterAmt& = 10& ^ (Len(Str$(UBBillRec(1).PrevRead(MCCnt))) - 1)
'''                ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MCCnt)) + UBBillRec(1).CurRead(MCCnt)
'''              End If
'''              If CubMtr Then
'''                ReadAmt& = ReadAmt& * 7.481
'''              End If
'''              RateConsump(WhatRate) = RateConsump(WhatRate) + (ReadAmt& * MeterMulti&)
'''              RateCount(WhatRate) = RateCount(WhatRate) + 1
'''              Bookconsump(WhatBook).Consump(TRevCnt) = Bookconsump(WhatBook).Consump(TRevCnt) + (ReadAmt& * MeterMulti&)
'''              ConsumpTot(TRevCnt, 1) = ConsumpTot(TRevCnt, 1) + (ReadAmt& * MeterMulti&)
'''              Consump& = ReadAmt& * MeterMulti&
'''              ThisMeterUseCnt = UBCustRec(1).LocMeters(MCCnt).UseCnt
'''              If ThisMeterUseCnt <= 0 Then ThisMeterUseCnt = 1
'''              AvgUse& = UBCustRec(1).LocMeters(MCCnt).AvgUse
'''              If AvgUse& > 0 Then
'''                HiConsump& = Round#(AvgUse& * (UBSetUpRec(1).HighRead * 0.01))
'''                LowConsump& = Round#(AvgUse& * (UBSetUpRec(1).LowRead * 0.01))
'''              End If
'''              ToPrintR(MCCnt) = Temp2$ + "~" + Using(fmt$(2), UBBillRec(1).CurRead(MCCnt)) + "~" + Using(fmt$(2), UBBillRec(1).PrevRead(MCCnt)) + "~" + Using(fmt$(2), ReadAmt& * MeterMulti&)
'''              If UBCustRec(1).EstFlag = "E" Then
'''                ToPrintR$(MCCnt) = ToPrintR$(MCCnt) + "~" + " E"             'Est. Reading
'''              ElseIf Consump& < LowConsump& Then
'''                ToPrintR$(MCCnt) = ToPrintR$(MCCnt) + "~" + " L"           'Low reading
'''              ElseIf Consump& > HiConsump& Then
'''                ToPrintR$(MCCnt) = ToPrintR$(MCCnt) + "~" + " H"             'High Reading
'''              Else
'''                ToPrintR$(MCCnt) = ToPrintR$(MCCnt) + "~ "
'''              End If
'''              If Consump& < MINAMT& Then
'''                ToPrintR$(MCCnt) = ToPrintR$(MCCnt) + "~" + " M"           'Minium Usage
'''              Else
'''                ToPrintR$(MCCnt) = ToPrintR$(MCCnt) + "~ "
'''              End If
'''              If UBBillRec(1).RevAmt(TRevCnt) > 0 And PrintedRevAmt = False Then
'''                PrintedRevAmt = True
'''                ToPrintR$(MCCnt) = ToPrintR$(MCCnt) + "~" + Using(fmt$(3), UBBillRec(1).RevAmt(TRevCnt))
'''                If UBBillRec(1).ProRatePCT < 100 And ProrateServ(TRevCnt) Then
'''                  ToPrintR$(MCCnt) = ToPrintR$(MCCnt) + "~" + "*"
'''                Else
'''                  ToPrintR$(MCCnt) = ToPrintR$(MCCnt) + "~ "
'''                End If
'''                If IFlag Then
'''                  ToPrintR$(MCCnt) = ToPrintR$(MCCnt) + "~" + " IR"
'''                Else
'''                  ToPrintR$(MCCnt) = ToPrintR$(MCCnt) + "~ "
'''                End If
'''              Else
'''                ToPrintR$(MCCnt) = ToPrintR$(MCCnt) + "~ ~ ~ ~ "
'''              End If
'''             'Print #UBRpt,
'''             ' Linecnt = Linecnt + 1
'''            End If
'''          Next
'''          '071197 Added this for mccormick. Has a sewer flat rate, Sewer is set up as
'''          '      a metered service but no meter on a flat rate charge. Charge was added
'''          '      to total, but didn't show on prebilling report.
'''
'''          If Not DoneOne Then
'''            DoneOne = True
'''            ToPrintR$(TRevCnt) = ToPrintR$(TRevCnt) + " ~ ~ ~ ~ ~ ~ ~" + Using(fmt$(3), UBBillRec(1).RevAmt(TRevCnt))
'''            If UBBillRec(1).ProRatePCT < 100 And ProrateServ(TRevCnt) Then
'''              ToPrintR$(TRevCnt) = ToPrintR$(TRevCnt) + " ~ ~ *"
'''            Else
'''              ToPrintR$(TRevCnt) = ToPrintR$(TRevCnt) + " ~ ~ "
'''            End If
'''            'THIS WAS REMARKED OUT, I DON'T KNOW WHY?
'''           ' Print #UBRpt,
'''            ''''''''''''''''''''''''''''''''''''''
'''          '  Linecnt = Linecnt + 1
'''          End If
'''        Else    'it's a nonmetered service
'''          ConsumpTot(TRevCnt, 1) = ConsumpTot(TRevCnt, 1) + 1
'''          If WhatRate > 0 Then
'''            RateConsump(WhatRate) = RateConsump(WhatRate) + 1
'''            RateCount(WhatRate) = RateCount(WhatRate) + 1
'''            RateTotals(WhatRate) = Round#(RateTotals(WhatRate) + UBBillRec(1).RevAmt(TRevCnt))
'''          End If
'''          Bookconsump(WhatBook).Consump(TRevCnt) = Bookconsump(WhatBook).Consump(TRevCnt) + 1
'''          Bookconsump(WhatBook).RevAmt(TRevCnt) = Round#(Bookconsump(WhatBook).RevAmt(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
'''          Bookconsump(WhatBook).TaxAmt(TRevCnt) = Round#(Bookconsump(WhatBook).TaxAmt(TRevCnt) + UBBillRec(1).TaxAmt(TRevCnt))
'''          RevTotals(TRevCnt) = Round#(RevTotals(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
'''          ToPrintR$(TRevCnt) = ToPrintR$(TRevCnt) + " ~ ~ ~ ~ ~ ~ ~ " + Using(fmt$(3), UBBillRec(1).RevAmt(TRevCnt))
'''          If UBBillRec(1).ProRatePCT < 100 And ProrateServ(TRevCnt) Then
'''            ToPrintR$(TRevCnt) = ToPrintR$(TRevCnt) + " ~ ~ *"
'''          Else
'''            ToPrintR$(TRevCnt) = ToPrintR$(TRevCnt) + " ~ ~ "
'''          End If
'''        End If
'''        If Not DoneOne Then
'''          'Print #UBRpt,
'''          'Linecnt = Linecnt + 1
'''        End If
''''      Else
''''         ToPrintR$(TRevCnt) = ToPrintR$(TRevCnt) + " ~ ~ ~ ~ ~ ~ ~ ~ ~ "
'''      End If
'''      'If (TRevCnt = NumOfRevs) And UBBillRec(1).TransAmt = 0 Then 'changed numofrevs to 15
'''      If (TRevCnt = 15) And UBBillRec(1).TransAmt = 0 Then 'changed numofrevs to 15
'''        If UBBillRec(1).TransAmt = 0 Then       'CONSUMPTION inactive account
'''          ToPrintR$(TRevCnt) = ""
'''          For TTRevCnt = 1 To NumOfRevs
'''            For MCCnt = 1 To 7
'''              LocMeterType$ = QPTrim$(UBCustRec(1).LocMeters(MCCnt).MTRType)
'''              If (Len(LocMeterType$) > 0) And (LocMeterType$ = UBCustRec(1).Serv(TTRevCnt).RMtrType) Then
'''                If UBBillRec(1).CurRead(MCCnt) < 0 Then
'''                  UBBillRec(1).CurRead(MCCnt) = 0
'''                End If
'''                If UBBillRec(1).PrevRead(MCCnt) < 0 Then
'''                  UBBillRec(1).PrevRead(MCCnt) = 0
'''                End If
'''                CurReadAmt& = UBBillRec(1).CurRead(MCCnt)
'''                PreReadAmt& = UBBillRec(1).PrevRead(MCCnt)
'''                If CurReadAmt& <> PreReadAmt& Then
'''                  If Not ConsumpFlag Then
'''                    ToPrintN$ = UBCustRec(1).Status + "~" + Using("     #####   ", ThisCustRec&)
'''                    ToPrintN$ = ToPrintN$ + "~" + UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB + "~" + Left$(UBCustRec(1).CustName, 25) + "~" + Left$(UBCustRec(1).SERVADDR, 25)
'''                    ToPrintN$ = ToPrintN$ + "~" + Using("   ###", UBBillRec(1).ProRatePCT) + "%"
'''                    ToPrintN$ = ToPrintN$ + "~" + BillTo$
'''                    PrintOK = True
'''                    'Linecnt = Linecnt + 1
''''                  Else
''''                    ToPrintN$ = " ~ ~ ~ ~ ~ ~ ~ "
'''                  End If
'''                  ConsumpFlag = True
'''                  MeterNum$ = QPTrim$(UBCustRec(1).Serv(TTRevCnt).RATECODE)
'''                  If Len(MeterNum$) > 0 Then
'''                    If UBCustRec(1).LocMeters(MCCnt).NumUser > 1 Then
'''                      MeterNum$ = MeterNum$ + "*" + QPTrim$(Str$(UBCustRec(1).LocMeters(MCCnt).NumUser))
'''                    End If
'''                    RSet Temp2$ = MeterNum$
'''                  Else
'''                    Temp2$ = " "
'''                  End If
'''                  ConsumpAmt& = CurReadAmt& - PreReadAmt&
'''                  '103098 Added meter roll over check to inactive consumption
'''                  If ConsumpAmt& < 0 Then       'Meter rolled over or, been misread
'''                    MaxMeterAmt& = 10& ^ (Len(Str$(PreReadAmt&)) - 1)
'''                    ConsumpAmt& = (MaxMeterAmt& - PreReadAmt&) + CurReadAmt&
'''                  End If
'''                  If InStr(UBCustRec(1).CUSTTYPE, "NON") Then
'''                    'For Nonprofits include consumption as normal   'cleveland
'''                    '040998 Made changes here
'''                    For NONRateCnt = 1 To NumOfRates
'''                      If UBRateTbls(NONRateCnt).RATECODE = UBCustRec(1).Serv(TTRevCnt).RATECODE Then
'''                        NONRate = NONRateCnt
'''                        Exit For
'''                      End If
'''                    Next
'''                    If NONRate > 0 Then
'''                      RateConsump(NONRate) = RateConsump(NONRate) + ConsumpAmt&
'''                    End If
'''                    ConsumpTot(TTRevCnt, 1) = ConsumpTot(TTRevCnt, 1) + ConsumpAmt&
'''                    Bookconsump(WhatBook).Consump(TTRevCnt) = Bookconsump(WhatBook).Consump(TTRevCnt) + ConsumpAmt&
'''                    '040998 Made changes here 'cleveland
'''                  Else          'add consumption to inactives
'''                    ConsumpTot(TTRevCnt, 2) = ConsumpTot(TTRevCnt, 2) + ConsumpAmt&
'''                  End If
'''                  ToPrintR$(TTRevCnt) = RevDesc(TTRevCnt) + "~" + Temp2$ + "~" + Using(fmt$(2), CurReadAmt&) + "~" + Using(fmt$(2), PreReadAmt&) + "~" + Using(fmt$(2), ConsumpAmt&) + "~ ~ ~ ~ ~ ~"
'''                  'Linecnt = Linecnt + 1
'''                End If
'''
'''              End If
'''            Next
'''
'''          Next
'''        End If
'''        If ConsumpFlag And UBCustRec(1).Status <> "A" Then
'''          ConsumpFlag = False
'''          ToPrintM$ = "**** Consumption Noted on an Inactive Account. ****"
'''          'Linecnt = Linecnt + 1
'''          If Not SkipSeparator Then
'''          '  Print #UBRpt, fmt$(0)
'''          '  Linecnt = Linecnt + 1
'''          End If
'''        ElseIf ConsumpFlag Then
'''          'Customer Status is "A"
'''          'This happens when a cust has consumption and there rate code
'''          'has a zero calc amount. "i.e. a Church or other nonprofit"
'''          If InStr(UBCustRec(1).CUSTTYPE, "NON") Then
'''            ToPrintM$ = "*** NON-PROFIT ***"
'''            'Linecnt = Linecnt + 1
'''          End If
'''          ConsumpFlag = False
'''          If Not SkipSeparator Then
'''           ' Print #UBRpt, fmt$(0)
'''           'Linecnt = Linecnt + 1
'''          End If
'''        End If
'''      ElseIf (TRevCnt = NumOfRevs) And UBBillRec(1).TransAmt > 0 Then
'''        '102998  Moved tax printing to here "now prints one tax line per customer
'''        CTaxAmt# = 0
'''        For TXCnt = 1 To NumOfRevs
'''          If UBBillRec(1).TaxAmt(TXCnt) > 0 Then
'''            CTaxAmt# = Round#(CTaxAmt# + UBBillRec(1).TaxAmt(TXCnt))
'''          End If
'''        Next
'''        If CTaxAmt# > 0 Then
'''          ToPrintX$ = " Tax" + "~" + Using(fmt$(3), CTaxAmt#)
'''          'Linecnt = Linecnt + 1
'''        Else
'''          ToPrintX$ = " ~ "
'''        End If
'''        Bills2Print = Bills2Print + 1
'''        AcctBalance# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
'''        ToPrintT$ = "Current:" + "~" + Using(fmt$(6), UBBillRec(1).TransAmt)
'''        If AcctBalance# <> 0 Then
'''          ToPrintT$ = ToPrintT$ + "~" + "Previous:" + "~" + Using(fmt$(6), AcctBalance#)
'''          TAcctBalance# = Round#(TAcctBalance# + AcctBalance#)
'''        Else
'''          ToPrintT$ = ToPrintT$ + "~ ~ "
'''        End If
'''        ToPrintT$ = ToPrintT$ + "Total:" + "~" + Using(fmt$(6), Round#(AcctBalance# + UBBillRec(1).TransAmt))
'''        'Linecnt = Linecnt + 1
'''        If Not SkipSeparator Then
'''         ' Print #UBRpt, fmt$(0)
'''         ' Linecnt = Linecnt + 1
'''        End If
'''      End If
'''      If UBBillRec(1).TaxExempt = "Y" Then
'''        TaxExmp(TRevCnt) = Round#(TaxExmp(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
'''      End If
'''    Next
'''    '020199 Moved pump code processing to here. Stops bug in getting true
'''    '       meter consumption figures.
'''    GoSub GetWhatPump
'''    If HasAPumpCode Then
'''      PumpConsump(WhatPump).CustCnt = PumpConsump(WhatPump).CustCnt + 1
'''      For MPCnt = 1 To 7
'''        PumpMtrOK = False
'''        CubMtr = False
'''        LocMeterType$ = QPTrim$(UBCustRec(1).LocMeters(MPCnt).MTRType)
'''        Select Case LocMeterType$
'''        Case "C", "S", "W", "T"
'''          PumpMtrOK = True
'''        End Select
'''        If PumpMtrOK Then
'''          MeterMulti& = UBCustRec(1).LocMeters(MPCnt).MTRMulti
'''          If UBCustRec(1).LocMeters(MPCnt).MTRUnit = "C" Then
'''            CubMtr = True
'''          End If
'''          If MeterMulti& <= 0 Then MeterMulti& = 1
'''          ReadAmt& = UBBillRec(1).CurRead(MPCnt) - UBBillRec(1).PrevRead(MPCnt)
'''          If ReadAmt& < 0 Then  'Meter rolled over or, been misread
'''            MaxMeterAmt& = 10& ^ (Len(Str$(UBBillRec(1).PrevRead(MPCnt))) - 1)
'''            ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MPCnt)) + UBBillRec(1).CurRead(MPCnt)
'''          End If
'''          If CubMtr Then
'''            ReadAmt& = ReadAmt& * 7.481
'''          End If
'''
'''          PumpConsump(WhatPump).Consump = PumpConsump(WhatPump).Consump + (ReadAmt& * MeterMulti&)
'''        End If
'''
'''      Next  'mpcnt
'''    End If
'''
'''SkipEm:
''''    If AskAbandonPrint% Then
''''      UBLog "ABORTED: Prebilling report"
''''      UBLog "Closing files."
''''      Close
''''      AbortFlag = True
''''      Exit For
''''    End If
''''    ShowPctComp cnt, NumOfRecs
'''    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
'''    If FrmShowPctComp.Out = True Then
'''      Close
'''      FrmShowPctComp.Out = False
'''      GoTo ExitPreReport
'''    End If
'''      If PrintOK = True Then
'''      For ThisCnt = 1 To NumOfRevs
'''      ToPrint$ = ToPrintN$ + "~" + ToPrintR$(ThisCnt) + "~" + ToPrintX$ + "~" + ToPrintT$ + "~" + ToPrintM$
'''      Print #UBRpt, ToPrint$
'''      ToPrintR$(ThisCnt) = ""
'''      Next
'''      End If
'''      ToPrint$ = ""
'''      ToPrintN$ = ""
'''      ToPrintX$ = ""
'''      ToPrintT$ = ""
'''      ToPrintM$ = ""
'''
'''  Next
'''  If AbortFlag Then GoTo ExitPreReport
'''
'''  'Print #UBRpt, FF$
'''
'''  GoSub TitleLine
''''  Print #UBRpt, "Billing Grand Totals"
''''  If TennFlag Then
''''    Print #UBRpt, "                                Inactive          Taxed      NONTax     FlatRate"
''''    Print #UBRpt, "Revenue/Tax        Consump       Consump         Amount      Amount      Amount"
''''  Else
''''    Print #UBRpt, "                                 Inactive                             Flat Rate"
''''    Print #UBRpt, "Revenue/Tax    Consumption      Consumption            Amount           Amount"
''''  End If
''''  Print #UBRpt, fmt$(0)
'''
'''  TotalFlatAmt# = 0
'''  TotalRevAmt# = 0
'''  TotalTaxAmt# = 0
'''
'''  For RaCnt = 1 To NumOfRevs
'''    If TennFlag Then
'''      ToPrintS$ = UBSetUpRec(1).Revenues(RaCnt).REVNAME + "~" + Using(fmt$(4), ConsumpTot(RaCnt, 1)) + "~" + Using(fmt$(4), ConsumpTot(RaCnt, 2))
'''      If TaxTotals(RaCnt) > 0 Then
'''        ToPrintS$ = ToPrintS$ + "~" + Using(fmt$(1), Round#(RevTotals(RaCnt) - FlatTotals(RaCnt) - TaxExmp(RaCnt))) + "~" + Using(fmt$(1), TaxExmp(RaCnt)) + "~" + Using(fmt$(1), FlatTotals(RaCnt))
'''      Else
'''        ToPrintS$ = ToPrintS$ + "~" + Using(fmt$(1), Round#(RevTotals(RaCnt) - FlatTotals(RaCnt))) + "~ ~" + Using(fmt$(1), FlatTotals(RaCnt))
'''      End If
'''    Else
'''      ToPrintS$ = UBSetUpRec(1).Revenues(RaCnt).REVNAME + "~" + Using(fmt$(4), ConsumpTot(RaCnt, 1)) + "~" + Using(fmt$(4), ConsumpTot(RaCnt, 2))
'''      ToPrintS$ = ToPrintS$ + "~" + Using(fmt$(1), RevTotals(RaCnt) - FlatTotals(RaCnt)) + "~ ~" + Using(fmt$(1), FlatTotals(RaCnt))
'''    End If
'''    Print #UBRPTG, ToPrintS$
'''    ToPrintS$ = ""
'''    TotalFlatAmt# = Round#(TotalFlatAmt# + FlatTotals(RaCnt))
'''    TotalRevAmt# = Round#(TotalRevAmt# + RevTotals(RaCnt))
'''    If TaxTotals(RaCnt) > 0 Then
'''      If TennFlag Then
'''        Print #UBRPTG, " Tax" + "~ ~ ~" + Using(fmt$(1), TaxTotals(RaCnt)) + "~ "
'''      Else
'''        Print #UBRPTG, " Tax" + "~ ~ ~" + Using(fmt$(1), TaxTotals(RaCnt)) + "~ "
'''      End If
'''      TotalTaxAmt# = Round#(TotalTaxAmt# + TaxTotals(RaCnt))
'''    End If
'''
'''  Next
'''  'Print #UBRpt, fmt$(0)
'''  Print #UBRPTS, "PREVIOUS:" + "~" + Using(fmt$(6), TAcctBalance#)
'''  Print #UBRPTS, "REVENUE TOTAL:" + "~" + Using(fmt$(5), Round#(TotalRevAmt# - TotalFlatAmt#))
'''  Print #UBRPTS, "BILL COUNT:" + "~" + Using(fmt$(2), Bills2Print)
'''  Print #UBRPTS, "FLAT TOTAL:" + "~" + Using(fmt$(5), TotalFlatAmt#)
'''  Print #UBRPTS, "TAX TOTAL:" + "~" + Using(fmt$(5), TotalTaxAmt#)
'''  Print #UBRPTS, "BILLING TOTAL:" + "~" + Using(fmt$(5), Round#(TotalRevAmt# + TotalTaxAmt#))
'''  'Print #UBRpt, FF$
'''
'''  TotalRevAmt# = 0
'''
'''  GoSub RptTotRateHeader
'''
'''  For RaCnt = 1 To NumOfRates
'''    If (RateTotals(RaCnt) <> 0) Or (RateConsump(RaCnt) <> 0) Then
'''      If Len(QPTrim$(UBRateTbls(RaCnt).RATECODE)) > 0 Then
'''        ToPrintS$ = Str$(UBRateTbls(RaCnt).RATECODE) + "~" + UBRateTbls(RaCnt).RATEDESC + "~" + Using(fmt$(4), RateConsump(RaCnt))
'''        ToPrintS$ = ToPrintS$ + "~" + Using(fmt$(1), RateTotals(RaCnt))
'''        ToPrintS$ = ToPrintS$ + "~" + Using(fmt$(2), RateCount(RaCnt))
'''        'Linecnt = Linecnt + 1
'''        Print #UBRPTR, ToPrintS$
'''        ToPrintS$ = ""
'''        TotalRevAmt# = Round#(TotalRevAmt# + RateTotals(RaCnt))
'''        If RTaxTot(RaCnt) > 0 Then
'''          Print #UBRPTR, " Tax" + "~ ~ ~ " + Using(fmt$(1), RTaxTot(RaCnt)) + "~"
'''          'Linecnt = Linecnt + 1
'''        End If
'''        If Linecnt > MaxLines Then
'''          'Print #UBRpt, FF$
'''          'GoSub RptTotRateHeader
'''        End If
'''      End If
'''    End If
'''  Next
'''
''' ' Print #UBRpt, fmt$(0)
'''  Print #UBRPTS, "~ ~" + "TAX TOTAL:" + "~" + Using(fmt$(5), TotalTaxAmt#) + "~"
'''  Print #UBRPTS, "~ ~" + "TOTAL:" + "~" + Using(fmt$(5), Round#(TotalRevAmt# + TotalTaxAmt#)) + "~"
'''  'Print #UBRpt, FF$
'''  'SortT BookConsump(1), TBooks, 0, Len(BookConsump(1)), 0, -1
'''  BookCQSort Bookconsump(), 1, TBooks
'''  GoSub BookHeader
'''
'''  For cnt = 1 To TBooks
'''    TestTot# = 0
'''    For ZCnt = 1 To NumOfRevs
'''      TestTot# = Round#(TestTot# + Bookconsump(cnt).RevAmt(ZCnt))
'''    Next
'''    If TestTot# <> 0 Then
'''      If Bookconsump(cnt).Book < 10 Then
'''        Book$ = "0" + QPTrim$(Str$(Bookconsump(cnt).Book))
'''      Else
'''        Book$ = QPTrim$(Str$(Bookconsump(cnt).Book))
'''      End If
'''      ToPrintN$ = Book$ + "~" + Str(Bookconsump(cnt).CustCnt)
'''      TBookAmt# = 0
'''      TBTaxAmt# = 0
'''      For RCnt = 1 To NumOfRevs
'''        ToPrintS$ = RevDesc(RCnt) + "~" + Using(fmt$(4), Bookconsump(cnt).Consump(RCnt))
'''        ToPrintS$ = ToPrintS$ + "~" + Using("##########.##", Bookconsump(cnt).RevAmt(RCnt))
'''        TBookAmt# = Round#(TBookAmt# + Bookconsump(cnt).RevAmt(RCnt))
'''        Print #UBRPTB, ToPrintN$ + "~" + ToPrintS$
'''        ToPrintS$ = ""
'''        If Bookconsump(cnt).TaxAmt(RCnt) > 0 Then
'''          Print #UBRPTB, ToPrintN$ + "~" + " Tax" + "~ ~" + Using(fmt$(1), Bookconsump(cnt).TaxAmt(RCnt))
'''          TBTaxAmt# = Round#(TBTaxAmt# + Bookconsump(cnt).TaxAmt(RCnt))
'''          Linecnt = Linecnt + 1
'''        End If
'''        Linecnt = Linecnt + 1
'''      Next
'''      TBookGTot# = Round#(TBookGTot# + TBookAmt# + TBTaxAmt#)
'''      'Print #UBRpt, Tab(42); "Book Total:"; Tab(57); Using(fmt$(5), Round#(TBookAmt# + TBTaxAmt#))
'''      If cnt < TBooks Then
'''      '  Print #UBRpt, fmt$(0)
'''      End If
'''      'Linecnt = Linecnt + 1
'''    End If
'''    If Linecnt > MaxLines And cnt < TBooks Then
'''      'Print #UBRpt, FF$
'''      'GoSub BookHeader
'''    End If
'''  ToPrintN$ = ""
'''SkipThisBook:
'''  Next
'''
''''  Print #UBRpt, fmt$(0)
''''  Print #UBRpt, Tab(35); "Books GRAND Total:"; Tab(57); Using(fmt$(5), TBookGTot#)
''''  Print #UBRpt, FF$
'''
'''  If TPumps > 0 Then
'''    GoSub PumpHeader
'''    TMMConsump# = 0
'''    For cnt = 1 To TPumps
'''      Print #UBRPTP, PumpConsump(cnt).PumpCode + "~" + Using("###########", PumpConsump(cnt).CustCnt) + "~" + PumpConsump(cnt).Consump
'''      TMMConsump# = TMMConsump# + PumpConsump(cnt).Consump
'''    Next
'''    'Print #UBRpt, fmt$(0)
'''    Print #UBRPTP, " ~Pump Code Total:~" + Using("###########", TMMConsump#)
'''  End If
'''
'''  Close
'''
'''  UBLog "Finished writing prebilling report."
'''  Select Case Choice
'''  Case 1
'''    RptText$ = "(Customer"
'''  Case 2
'''    RptText$ = "(Account"
'''  Case 3
'''    RptText$ = "(Location"
'''  Case 4
'''    RptText$ = "(Postal RT."
'''  Case 5
'''    RptText$ = "(ZipCode"
'''  Case 6
'''    RptText$ = "(Sequence"
'''  End Select
'''  RptText$ = RptText$ + " Order)"
'''
'''  Erase UBSetUpRec, RevDesc, UBRateTbls, RateConsump
'''  Erase fmt$, UBCustRec, UBBillRec, FlatTotals
'''  Erase RevTotals, TaxTotals, ConsumpTot
'''  Erase RateTotals, RTaxTot, Bookconsump, Indexarray
'''  Erase RateCount, ProrateServ
'''  Erase PumpConsump, TaxExmp
'''
'''  If Not AbortFlag Then
''' '   ViewPrint "UBPREBIL.RPT", "Pre-Billing Report " + RptText$
'''    'PrintRptFile "Pre-Billing Report " + RptText$, "UBPREBIL.RPT", LPTPort, RetCode, EntryPoint
'''    If BookFlag Then
'''      Kill UBBillsFile
'''    End If
'''  End If
'''
'''  GoTo ExitPreReport
'''
'''PrintPreHeader:
''''  GoSub TitleLine
''''  Print #UBRpt, "Stat  Act.  Locat    Customer Name             Service Address       Prorate%"
''''  Print #UBRpt, "Revenue            R-Code     Cur Read    Pre Read     Consump        Charges"
''''  Print #UBRpt, fmt$(0)
''''  Linecnt = 5
'''Return
'''
'''GetWhatBook:
'''  BadBookFlag = False
''' WhatBook = 0
''' If Len(QPTrim$(UBCustRec(1).Book)) = 0 Then
'''   If UBCustRec(1).Status = "A" Then
'''     BadBookFlag = True
'''     'testing vvv
'''     WhatBook = 0
'''   End If
'''   GoTo ErrorBookExit
''' End If
'''
''' ThisBook = Val(UBCustRec(1).Book)
''' If TBooks > 0 Then
'''   For TBCnt = 1 To TBooks
'''     If Bookconsump(TBCnt).Book = ThisBook Then
'''       WhatBook = TBCnt
'''       Exit For
'''     End If
'''   Next
'''   If WhatBook = 0 Then
'''     TBooks = TBooks + 1
'''     ReDim Preserve Bookconsump(0 To TBooks) As BookConsumpType
'''      Bookconsump(TBooks).Book = ThisBook
'''      WhatBook = TBooks
'''    End If
'''  Else
'''    TBooks = TBooks + 1
'''    Bookconsump(TBooks).Book = ThisBook
'''    WhatBook = TBooks
'''  End If
'''
'''ErrorBookExit:
'''  Return
'''
'''GetWhatPump:
'''  HasAPumpCode = True           'assume they have a pump code
'''  WhatPump = 0
'''  If Len(QPTrim$(UBCustRec(1).PumpCode)) = 0 Then
'''    If UBCustRec(1).Status = "A" Then
'''      HasAPumpCode = False      'no pump code
'''      WhatPump = 0
'''    End If
'''    GoTo PumpCodeReturn
'''  End If
'''
'''  CustPump$ = UCase$(QPTrim$(UBCustRec(1).PumpCode))
'''
'''  'IF CustPump$ = "34" THEN STOP
'''
'''  If Len(CustPump$) > 0 Then
'''    For TBCnt = 1 To TPumps
'''      ThisPump$ = QPTrim$(PumpConsump(TBCnt).PumpCode)
'''      If ThisPump$ = CustPump$ Then
'''        WhatPump = TBCnt
'''        Exit For
'''      End If
'''    Next
'''    If WhatPump = 0 Then
'''      TPumps = TPumps + 1
'''      ReDim Preserve PumpConsump(0 To TPumps) As PumpConsumpType
'''      PumpConsump(TPumps).PumpCode = CustPump$
'''      WhatPump = TPumps
'''    End If
'''  Else
'''    TPumps = TPumps + 1
'''    PumpConsump(TPumps).PumpCode = CustPump$
'''    WhatPump = TPumps
'''  End If
'''
'''PumpCodeReturn:
'''  Return
'''
'''RptTotRateHeader:
''''  GoSub TitleLine
''''  Print #UBRpt,
''''  Print #UBRpt, "Report Totals by Rate Code"
''''  Print #UBRpt,
''''  Print #UBRpt, "Code      Rate Description            Consumption           Amount      Bills"
''''  Print #UBRpt, fmt$(0)
''''  Linecnt = 5
'''  Return
'''
'''BookHeader:
''''  GoSub TitleLine
''''  Print #UBRpt, "Report Totals by Book"
''''  Print #UBRpt,
''''  Print #UBRpt, "Book"
''''  Print #UBRpt, "Revenue                      Consumption                         Amount"
''''  Print #UBRpt, fmt$(0)
''''  Linecnt = 7
'''  Return
'''
'''PumpHeader:
''''  GoSub TitleLine
''''  Print #UBRpt, "Report Totals by Pump Code"
''''  Print #UBRpt,
''''  Print #UBRpt, "PumpCode                  Customer Count                    Consumption"
''''  Print #UBRpt, fmt$(0)
''''  Linecnt = 6
'''  Return
'''
'''TitleLine:
''''  PageNo = PageNo + 1
''''  Print #UBRpt, "Utility Pre-Billing Report.  "; TownName$; Tab(70); "Page: "; PageNo
''''  Print #UBRpt, TheDate$
''''  Return
'''ErrorAbortExit:
'''  Close
'''
'''ExitPreReport:
'''  UBLog "OUT: Prebilling Report" + CrLf$
'''
'''End Sub
'''
'''Private Sub PreBillReport211()
'''  Dim Temp2 As String, NumOfRevs As Integer, NumOfRates As Integer
'''  Dim UBRateTblRecLen As Integer, RateFile As Integer, cnt As Long
'''  Dim UBSetupLen As Integer, MowFlag As Boolean, TennFlag As Boolean
'''  Dim TempRev As String, DoFuelAdjFlag As Boolean, SkipInactive As Boolean
'''  Dim SkipSeparator As Boolean, ThisBook As Integer, BookNum As Integer
'''  Dim BookFlag As Boolean, ThisCycle As Integer, CycleFlag As Boolean
'''  Dim SeqFlag As String, Choice As Integer, FuelAdjAmt As Double
'''  Dim IndexName As String, UsingAcct As Boolean, IdxTypeText As String
'''  Dim AbortFlag As Boolean, TheDate As String, UBCustRecLen As Integer
'''  Dim UBBillRecLen As Integer, TBooks As Integer, NumOfRecs As Long
'''  Dim Handle As Integer, IdxRecLen As Integer, lcnt As Long
'''  Dim UBBill As Integer, UBCust As Integer, UBRpt As Integer
'''  Dim ThisCustRec As Long, BillTo As String, BadBookFlag As Boolean
'''  Dim WhatBook As Integer, FRCnt As Integer, WhatService As Integer
'''  Dim Multi As Integer, FlatAmt As Double, WhatRate As Integer
'''  Dim DoneOne As Boolean, TRevCnt As Integer, IFlag As Boolean
'''  Dim TRateCnt As Integer, MINAMT As Long, PrintedRevAmt As Boolean
'''  Dim MCCnt As Integer, CubMtr As Boolean, LocMeterType As String
'''  Dim MeterMulti As Long, MeterNum As String, ReadAmt As Long
'''  Dim MaxMeterAmt As Long, Consump As Long, ThisMeterUseCnt As Integer
'''  Dim AvgUse As Long, HiConsump As Long, LowConsump As Long
'''  Dim TTRevCnt As Integer, CurReadAmt As Long, PreReadAmt As Long
'''  Dim ConsumpFlag As Boolean, ConsumpAmt As Long, NONRateCnt As Integer
'''  Dim NONRate As Integer, CTaxAmt As Double, TXCnt As Integer
'''  Dim Bills2Print As Integer, AcctBalance As Double, WhatPump As Integer
'''  Dim TAcctBalance As Double, HasAPumpCode As Boolean, MPCnt As Integer
'''  Dim PumpMtrOK As Boolean, TotalFlatAmt As Double, TotalRevAmt As Double
'''  Dim TotalTaxAmt As Double, RaCnt As Integer, TestTot As Double
'''  Dim ZCnt As Integer, Book As String, TBookAmt As Double, TPumps As Integer
'''  Dim TBTaxAmt As Double, RCnt As Integer, TBookGTot As Double
'''  Dim TMMConsump As Double, RptText As String, TBCnt As Integer
'''  Dim CustPump As String, ThisPump As String, ToPrintM As String
'''  Dim ToPrintN As String, ToPrintR(1 To 15) As String, ToPrintT As String
'''  Dim ToPrintX As String, UBRPTG As Integer, UBRPTR As Integer
'''  Dim UBRPTB As Integer, ToPrintS As String, UBRPTS As Integer
'''  Dim UBRPTP As Integer, ToPrint As String, ThisCnt As Integer
'''  Dim PrintOK As Boolean, MtrPrn(1 To 7) As String
'''  UBLog "IN: Prebilling Report"
'''  PageNo = 0
'''  Temp2$ = Space$(12)
'''  NumOfRevs = MaxRevsCnt        'assume max munber of revenue sources
'''  NumOfRates = GetNumRateRecs%
'''  ReDim UBRateTbls(1 To NumOfRates) As UBRateTblRecType
'''  UBRateTblRecLen = Len(UBRateTbls(1))
'''
'''  ReDim RateConsump(1 To NumOfRates) As Double
'''
'''  RateFile = FreeFile
'''  Open "UBRATE.DAT" For Random Shared As RateFile Len = UBRateTblRecLen
'''  For cnt = 1 To NumOfRates
'''    Get RateFile, cnt, UBRateTbls(cnt)
'''  Next
'''  Close
'''
'''  'SortT UBRateTbls(1), NumOfRates, 0, UBRateTblRecLen, 0, 4
'''  RateQSort UBRateTbls(), 1, NumOfRates
'''
'''  ReDim ProrateServ(1 To 15) As Integer
'''
'''  ReDim UBSetUpRec(1) As UBSetupRecType
'''  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
'''
'''  TownName$ = UBSetUpRec(1).UTILNAME
'''  If InStr(TownName$, "MOWAS") > 0 Then
'''    MowFlag = True
'''  End If
'''
'''  If UBSetUpRec(1).DEFSTATE = "TN" Then
'''    TennFlag = True
'''  End If
'''
'''  ReDim RevDesc(1 To MaxRevsCnt) As String * 12
'''  For cnt = 1 To MaxRevsCnt     'find last active revenue
'''    TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(cnt).REVNAME)
'''    If Len(TempRev$) = 0 Then
'''      NumOfRevs = cnt - 1       'set actual number of revenues
'''      Exit For
'''    Else        'build revenue description lines
'''      LSet RevDesc(cnt) = UCase$(TempRev$)
'''      If InStr(RevDesc(cnt), "ELECTRIC") Then
'''        DoFuelAdjFlag = True
'''      End If
'''    End If
'''  Next
'''  '111398 Prorate
'''  For cnt = 1 To MaxRevsCnt
'''    If UBSetUpRec(1).Revenues(cnt).ProRate = "Y" Then
'''      ProrateServ(cnt) = True
'''    End If
'''  Next
'''
'''  If UBSetUpRec(1).SkipInactive = "Y" Then
'''    SkipInactive = True
'''  End If
'''
'''  If UBSetUpRec(1).SkipSeparator = "Y" Then
'''    SkipSeparator = True
'''  End If
'''
'''  If UBSetUpRec(1).PreByBook = "Y" Then
'''    ThisBook = Val(fptxtRoute1)
'''    If ThisBook = 99 Then
'''      ThisBook = -1
'''    End If
'''    BookNum = ThisBook
'''    If ThisBook = -1 Then
'''      BookFlag = False
'''    ElseIf ThisBook <= 0 Then
'''      GoTo ExitPreReport
'''    Else
'''      BookFlag = True
'''    End If
'''  ElseIf UBSetUpRec(1).BILLCYCL = "Y" Then
'''    ThisCycle = Val(fptxtRoute1)
'''    If ThisCycle <= 0 Then
'''      GoTo ExitPreReport
'''    Else
'''      CycleFlag = True
'''    End If
'''  End If
'''
'''  If UBSetUpRec(1).UseSeq = "Y" Then
'''    SeqFlag$ = "Y"
'''  End If
'''  FrmShowPctComp.Label1 = "Creating PreBilling Report"
'''  FrmShowPctComp.Show , Me
'''
'''Restart:
'''  Choice = fpcboPrintOrder.ListIndex + 1
'''  'GetPreBillOrder Choice, ExitFlag, SeqFlag$
'''
''' 'If ExitFlag Then GoTo ExitPreReport
'''  If DoFuelAdjFlag Then
'''    FuelAdjAmt# = Val(fptxtAdjustment)
'''    UBLog "Fuel adjustment factor:" + Str$(FuelAdjAmt#)
'''  Else
'''    FuelAdjAmt# = 0
'''  End If
'''
'''  If FuelAdjAmt# = -10000 Then GoTo Restart
'''
'''  Select Case Choice
'''  Case 0
'''    'ExitFlag = True
'''  Case 1        'Name
'''    IndexName$ = NameIndexFile
'''    'OkFlag = True
'''  Case 2        'Acct
'''    IndexName$ = ""
'''    UsingAcct = True
'''    'OkFlag = True
'''  Case 3        'Location
'''    IndexName$ = BookIndexFile
'''    'OkFlag = True
'''  Case 4        'Postal Route
'''    IdxTypeText$ = "Postal Route"
'''    MakePostalIndex IdxTypeText$
'''    IndexName$ = TempIndexName
'''    'OkFlag = True
'''  Case 5        'ZipCode
'''    IdxTypeText$ = "Zip-Code"
'''    'this mowflag for zip index doesn't matter cause both index
'''    'routines do same thing now.
'''    If MowFlag Then
'''      MakeMowZipCodeIndex IdxTypeText$
'''    Else
'''      MakeZipCodeIndex IdxTypeText$
'''    End If
'''    IndexName$ = TempIndexName
'''    'OkFlag = True
'''  Case 6        'Sequence number
'''    IdxTypeText$ = "Sequence Number"
'''    MakeSequenceIndex IdxTypeText$, Me
'''    IndexName$ = TempIndexName
'''    'OkFlag = True
'''  End Select
'''  MakeBillFile AbortFlag, FuelAdjAmt#, ThisCycle, ThisBook
'''
'''  If AbortFlag Then GoTo ExitPreReport
'''
'''  MaxLines = 53
'''
'''  ReDim fmt$(0 To 6)
'''  fmt$(0) = String$(80, "-")
'''  fmt$(1) = "#########.##"
'''  fmt$(2) = "#########"
'''  fmt$(3) = "######.##"
'''  fmt$(4) = "###########"
'''  fmt$(5) = "$###,###,###.##"
'''  fmt$(6) = "$#,###,###.##"
'''
'''  TheDate$ = "Date: " + Date$
'''
'''  ReDim UBCustRec(1 To 2) As NewUBCustRecType
'''  UBCustRecLen = Len(UBCustRec(1))
'''
'''  ReDim UBBillRec(1) As UBTransRecType
'''  UBBillRecLen = Len(UBBillRec(1))
'''
'''  ReDim FlatTotals(1 To NumOfRevs) As Double
'''  '021998 added flat revenue totals
'''  ReDim RevTotals(1 To NumOfRevs) As Double     'Revenue total amts
'''  '052097 added tax by revenue totals
'''  ReDim TaxTotals(1 To NumOfRevs) As Double     'Tax total amts
'''  ReDim ConsumpTot(1 To NumOfRevs, 1 To 2) As Double            'Consumption total amts
'''  ReDim RateConsump(1 To NumOfRates) As Double
'''  '012698 Added bill count by rate code
'''  ReDim RateCount(1 To NumOfRates) As Long
'''  ReDim RateTotals(1 To NumOfRates) As Double   'Rates total amts
'''  '052097 added tax by rate code totals
'''  ReDim RTaxTot(1 To NumOfRates) As Double      'Rates Tax total amts
'''  '052097 added tax by book totals to type def
'''  ReDim Bookconsump(0 To 1) As BookConsumpType  'Consumption by book
'''  ReDim PumpConsump(0 To 1) As PumpConsumpType  'Consumption by pump code
'''  ReDim TaxExmp(0 To NumOfRevs) As Double
'''
'''  TBooks = 0
'''  If UsingAcct Then
'''    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
'''  Else          'load the index
'''    UBLog "Loading index file: " + IndexName$
'''    IdxRecLen = 4
'''    NumOfRecs = FileSize(IndexName$) \ 4
'''    ReDim Indexarray(1 To NumOfRecs) As UBCustIndexRecType
'''    Handle = FreeFile
'''    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
'''    For lcnt& = 1 To NumOfRecs
'''      Get #Handle, lcnt&, Indexarray(lcnt&)
'''    Next
'''    Close Handle
'''    'FGetAH IndexName$, IndexArray(1), 4, NumOfRecs
'''  End If
'''
'''  UBBill = FreeFile
'''  Open UBBillsFile For Random Shared As UBBill Len = UBBillRecLen
'''  UBCust = FreeFile
'''  Open "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
'''  UBRpt = FreeFile
'''  Open "UBPREBIL.RPT" For Output As UBRpt
'''
'''  'BlockClear
'''  'ShowProcessingScrn "Processing Pre-Billing Report"
'''  UBLog "Writing prebilling report to disk."
'''
'''  GoSub PrintPreHeader
'''  For cnt = 1 To NumOfRecs
'''    PrintOK = False
'''    If UsingAcct Then
'''      ThisCustRec& = cnt
'''    Else
'''      ThisCustRec& = Indexarray(cnt).RecNum
'''    End If
'''
'''    Get UBCust, ThisCustRec&, UBCustRec(1)
'''
'''    If UBCustRec(1).DelFlag Then
'''      GoTo SkipEm
'''    End If
'''
'''    If SkipInactive And UBCustRec(1).Status <> "A" Then
'''      GoTo SkipEm
'''    ElseIf UBCustRec(1).Status = "F" Then       'skip over final's
'''      GoTo SkipEm
'''    ElseIf UBCustRec(1).Status = "B" Then       'skip over B-Status
'''      GoTo SkipEm
'''    End If
'''    If BookFlag Then
'''      If Val(UBCustRec(1).Book) <> ThisBook Then
'''        GoTo SkipEm
'''      End If
'''    End If
'''
'''    If CycleFlag Then
'''      If UBCustRec(1).BILLCYCL <> ThisCycle Then
'''        GoTo SkipEm
'''      End If
'''    End If
'''
'''    Get UBBill, ThisCustRec&, UBBillRec(1)
'''
'''    If Linecnt > MaxLines Then
'''      'Print #UBRpt, FF$
'''      GoSub PrintPreHeader
'''    End If
'''
'''    If UBBillRec(1).ActiveFlag <> 0 Then
'''      If UBCustRec(1).BillTo = "O" Then
'''        BillTo$ = " O"
'''      Else
'''        BillTo$ = " C"
'''      End If
'''      GoSub GetWhatBook
'''      If BadBookFlag Then
'''        If ErrorScrn(2, ThisCustRec&) Then
'''          AbortFlag = True
'''          Exit For
'''        End If
'''      End If
'''      Bookconsump(WhatBook).CustCnt = Bookconsump(WhatBook).CustCnt + 1
'''      'Print #UBRpt, UBCustRec(1).Status; Using("  #####  ", ThisCustRec&);
'''      'Print #UBRpt, UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; "  "; Left$(UBCustRec(1).CustName, 25); " "; Left$(UBCustRec(1).SERVADDR, 22); " ";
'''      'Print #UBRpt, Using("   ###", UBBillRec(1).ProRatePCT); "%";
'''      'Print #UBRpt, BillTo$
'''      ToPrintN$ = UBCustRec(1).Status + "~" + Using("  #####  ", ThisCustRec&)
'''      ToPrintN$ = ToPrintN$ + "~" + UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
'''      ToPrintN$ = ToPrintN$ + "~" + Left$(UBCustRec(1).CustName, 25) + "~" + Left$(UBCustRec(1).SERVADDR, 22)
'''      ToPrintN$ = ToPrintN$ + "~" + Using("   ###", UBBillRec(1).ProRatePCT) + "%"
'''      ToPrintN$ = ToPrintN$ + "~" + BillTo$
'''      PrintOK = True
'''      Linecnt = Linecnt + 1
'''      For FRCnt = 1 To 4
'''        WhatService = UBCustRec(1).FlatRates(FRCnt).REVSRC
'''        If UBCustRec(1).FlatRates(FRCnt).FRAMT <> 0 And WhatService > 0 Then
'''          Multi = UBCustRec(1).FlatRates(FRCnt).NumMin
'''          If Multi < 1 Then Multi = 1
'''          FlatAmt# = Round#(UBCustRec(1).FlatRates(FRCnt).FRAMT * Multi)
'''          '021998 Added flat rate summaries
'''          FlatTotals(WhatService) = Round#(FlatTotals(WhatService) + FlatAmt#)
'''        End If
'''      Next
'''      '102798 Added to skip accts that don't have a book/seq no. "J.R."
'''    ElseIf Len(QPTrim$(UBCustRec(1).Book)) = 0 And Len(QPTrim$(UBCustRec(1).SEQNUMB)) = 0 Then
'''      GoTo SkipEm
'''    End If
'''    WhatRate = 0
'''    DoneOne = False
'''    For TRevCnt = 1 To 15 'NumOfRevs
'''      If TRevCnt = 2 And UBBillRec(1).PenAtBill = -1 Then
'''        IFlag = True
'''      Else
'''        IFlag = False
'''      End If
'''      WhatRate = 0
'''      If UBBillRec(1).RevAmt(TRevCnt) <> 0 Then
'''        DoneOne = False
'''        'Print #UBRpt, RevDesc(TRevCnt);
'''        ToPrintR$(TRevCnt) = ToPrintR$(TRevCnt) + RevDesc(TRevCnt) + "~"
'''        '102198 Moved out of meter loop, Stoped multi meter tax report bug
'''        If UBBillRec(1).TaxAmt(TRevCnt) > 0 Then
'''          TaxTotals(TRevCnt) = Round#(TaxTotals(TRevCnt) + UBBillRec(1).TaxAmt(TRevCnt))
'''        End If
'''        For TRateCnt = 1 To NumOfRates
'''          If UBRateTbls(TRateCnt).RATECODE = UBCustRec(1).Serv(TRevCnt).RATECODE Then
'''            MINAMT& = UBRateTbls(TRateCnt).MINUNITS
'''            WhatRate = TRateCnt
'''            '102198 Moved from meter loop, Stops multi meter tax report bug
'''            RTaxTot(WhatRate) = Round#(RTaxTot(WhatRate) + UBBillRec(1).TaxAmt(TRevCnt))
'''            Exit For
'''          End If
'''        Next
'''        If UBSetUpRec(1).Revenues(TRevCnt).UseMtr = "Y" Then
'''          RevTotals(TRevCnt) = Round#(RevTotals(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
'''          Bookconsump(WhatBook).RevAmt(TRevCnt) = Round#(Bookconsump(WhatBook).RevAmt(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
'''          Bookconsump(WhatBook).TaxAmt(TRevCnt) = Round#(Bookconsump(WhatBook).TaxAmt(TRevCnt) + UBBillRec(1).TaxAmt(TRevCnt))
'''
'''          '02-20-97 Add revenue totals by rate code
'''          If WhatRate > 0 Then
'''            RateTotals(WhatRate) = Round#(RateTotals(WhatRate) + UBBillRec(1).RevAmt(TRevCnt))
'''          End If
'''          PrintedRevAmt = False
'''          For MCCnt = 1 To 7
'''            CubMtr = False
'''            LocMeterType$ = QPTrim$(UBCustRec(1).LocMeters(MCCnt).MTRType)
'''            MeterMulti& = UBCustRec(1).LocMeters(MCCnt).MTRMulti
'''            '063098 Added adjustment for cubic meters in consumption totals
'''            If UBCustRec(1).LocMeters(MCCnt).MTRUnit = "C" Then
'''              CubMtr = True
'''            End If
'''            If MeterMulti& <= 0 Then MeterMulti& = 1
'''            If (Len(LocMeterType$) > 0) And (LocMeterType$ = UBCustRec(1).Serv(TRevCnt).RMtrType) Then
'''              DoneOne = True
'''              MeterNum$ = QPTrim$(UBCustRec(1).Serv(TRevCnt).RATECODE)
'''              'use the Meternum$ to hold the rate code temporarily
'''              If Len(MeterNum$) > 0 Then
'''                If UBCustRec(1).LocMeters(MCCnt).NumUser > 1 Then
'''                  MeterNum$ = MeterNum$ + "*" + QPTrim$(Str$(UBCustRec(1).LocMeters(MCCnt).NumUser))
'''                End If
'''                RSet Temp2$ = MeterNum$
'''              End If
'''              ReadAmt& = UBBillRec(1).CurRead(MCCnt) - UBBillRec(1).PrevRead(MCCnt)
'''              If ReadAmt& < 0 Then              'Meter rolled over or, been misread
'''                MaxMeterAmt& = 10& ^ (Len(Str$(UBBillRec(1).PrevRead(MCCnt))) - 1)
'''                ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MCCnt)) + UBBillRec(1).CurRead(MCCnt)
'''              End If
'''              If CubMtr Then
'''                ReadAmt& = ReadAmt& * 7.481
'''              End If
'''              RateConsump(WhatRate) = RateConsump(WhatRate) + (ReadAmt& * MeterMulti&)
'''              RateCount(WhatRate) = RateCount(WhatRate) + 1
'''              Bookconsump(WhatBook).Consump(TRevCnt) = Bookconsump(WhatBook).Consump(TRevCnt) + (ReadAmt& * MeterMulti&)
'''              ConsumpTot(TRevCnt, 1) = ConsumpTot(TRevCnt, 1) + (ReadAmt& * MeterMulti&)
'''              Consump& = ReadAmt& * MeterMulti&
'''              ThisMeterUseCnt = UBCustRec(1).LocMeters(MCCnt).UseCnt
'''              If ThisMeterUseCnt <= 0 Then ThisMeterUseCnt = 1
'''              AvgUse& = UBCustRec(1).LocMeters(MCCnt).AvgUse
'''              If AvgUse& > 0 Then
'''                HiConsump& = Round#(AvgUse& * (UBSetUpRec(1).HighRead * 0.01))
'''                LowConsump& = Round#(AvgUse& * (UBSetUpRec(1).LowRead * 0.01))
'''              End If
'''              'Print #UBRpt, Tab(14); Temp2$; Tab(30); Using(fmt$(2), UBBillRec(1).CurRead(MCCnt)); Tab(42); UBBillRec(1).PrevRead(MCCnt); Tab(54); ReadAmt& * MeterMulti&;
'''              ToPrintR$ = ToPrintR$ + Temp2$ + "~" + Using(fmt$(2), UBBillRec(1).CurRead(MCCnt)) + "~" + Using(fmt$(2), UBBillRec(1).PrevRead(MCCnt)) + "~" + Using(fmt$(2), (ReadAmt& * MeterMulti&)) + "~"
'''
'''              If UBCustRec(1).EstFlag = "E" Then
'''                'Print #UBRpt, " E";             'Est. Reading
'''                ToPrintR$ = ToPrintR$ + "E" + "~"
'''              ElseIf Consump& < LowConsump& Then
'''                'Print #UBRpt, " L";             'Low reading
'''                ToPrintR$ = ToPrintR$ + "L" + "~"
'''              ElseIf Consump& > HiConsump& Then
'''                'Print #UBRpt, " H";             'High Reading
'''                ToPrintR$ = ToPrintR$ + "H" + "~"
'''              Else
'''                ToPrintR$ = ToPrintR$ + " ~"
'''              End If
'''              If Consump& < MINAMT& Then
'''                'Print #UBRpt, " M";             'Minium Usage
'''                ToPrintR$ = ToPrintR$ + "E" + "~"
'''              Else
'''                ToPrintR$ = ToPrintR$ + " ~"
'''              End If
'''              If UBBillRec(1).RevAmt(TRevCnt) > 0 And PrintedRevAmt = False Then
'''                PrintedRevAmt = True
'''                'Print #UBRpt, Tab(69); Using(fmt$(3), UBBillRec(1).RevAmt(TRevCnt));
'''                 ToPrintR$ = ToPrintR$ + Using(fmt$(3), UBBillRec(1).RevAmt(TRevCnt)) + "~"
'''                If UBBillRec(1).ProRatePCT < 100 And ProrateServ(TRevCnt) Then
'''                  'Print #UBRpt, "*";
'''                  ToPrintR$ = ToPrintR$ + "*~"
'''                Else
'''                  ToPrintR$ = ToPrintR$ + " ~"
'''                End If
'''                If IFlag Then
'''                  'Print #UBRpt, " IR";
'''                  ToPrintR$ = ToPrintR$ + "IR~"
'''                Else
'''                  ToPrintR$ = ToPrintR$ + " ~"
'''                End If
'''
'''              End If
'''              'Print #UBRpt,
'''              Linecnt = Linecnt + 1
'''
'''            End If
'''          Next
'''          '071197 Added this for mccormick. Has a sewer flat rate, Sewer is set up as
'''          '      a metered service but no meter on a flat rate charge. Charge was added
'''          '      to total, but didn't show on prebilling report.
'''          If Not DoneOne Then
'''            DoneOne = True
'''            'Print #UBRpt, Tab(69); Using(fmt$(3), UBBillRec(1).RevAmt(TRevCnt));
'''            ToPrintR$ = ToPrintR$ + " ~ ~ ~ ~ ~ ~ " + Using(fmt$(3), UBBillRec(1).RevAmt(TRevCnt)) + "~"
'''            If UBBillRec(1).ProRatePCT < 100 And ProrateServ(TRevCnt) Then
'''              'Print #UBRpt, "*";
'''              ToPrintR$ = ToPrintR$ + "*~"
'''            Else
'''              ToPrintR$ = ToPrintR$ + " ~"
'''            End If
'''            'THIS WAS REMARKED OUT, I DON'T KNOW WHY?
'''            'Print #UBRpt,
'''            ''''''''''''''''''''''''''''''''''''''
'''            Linecnt = Linecnt + 1
'''          End If
'''        Else    'it's a nonmetered service
'''          ConsumpTot(TRevCnt, 1) = ConsumpTot(TRevCnt, 1) + 1
'''          If WhatRate > 0 Then
'''            RateConsump(WhatRate) = RateConsump(WhatRate) + 1
'''            RateCount(WhatRate) = RateCount(WhatRate) + 1
'''            RateTotals(WhatRate) = Round#(RateTotals(WhatRate) + UBBillRec(1).RevAmt(TRevCnt))
'''          End If
'''          Bookconsump(WhatBook).Consump(TRevCnt) = Bookconsump(WhatBook).Consump(TRevCnt) + 1
'''          Bookconsump(WhatBook).RevAmt(TRevCnt) = Round#(Bookconsump(WhatBook).RevAmt(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
'''          Bookconsump(WhatBook).TaxAmt(TRevCnt) = Round#(Bookconsump(WhatBook).TaxAmt(TRevCnt) + UBBillRec(1).TaxAmt(TRevCnt))
'''          RevTotals(TRevCnt) = Round#(RevTotals(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
'''          'Print #UBRpt, Tab(69); Using(fmt$(3), UBBillRec(1).RevAmt(TRevCnt));
'''          ToPrintR$ = ToPrintR$ + "~ ~ ~ ~ ~ ~" + Using(fmt$(3), UBBillRec(1).RevAmt(TRevCnt)) + "~"
'''          If UBBillRec(1).ProRatePCT < 100 And ProrateServ(TRevCnt) Then
'''            'Print #UBRpt, "*";
'''            ToPrintR$ = ToPrintR$ + "*~ ~"
'''          Else
'''            ToPrintR$ = ToPrintR$ + " ~ ~"
'''          End If
'''        End If
'''        If Not DoneOne Then
'''          'Print #UBRpt,
'''          Linecnt = Linecnt + 1
'''        End If
'''      Else
'''       ToPrintR$ = ToPrintR$ + " ~ y~ ~ ~ ~ ~ ~ ~ ~ "
'''      End If
'''      'If (TRevCnt = NumOfRevs) And UBBillRec(1).TransAmt = 0 Then
'''      If (TRevCnt = 15) And UBBillRec(1).TransAmt = 0 Then
'''        If UBBillRec(1).TransAmt = 0 Then  'CONSUMPTION inactive account
'''          ToPrintR$ = ""
'''          For TTRevCnt = 1 To 15 'NumOfRevs
'''            For MCCnt = 1 To 7
'''              LocMeterType$ = QPTrim$(UBCustRec(1).LocMeters(MCCnt).MTRType)
'''              If (Len(LocMeterType$) > 0) And (LocMeterType$ = UBCustRec(1).Serv(TTRevCnt).RMtrType) Then
'''                If UBBillRec(1).CurRead(MCCnt) < 0 Then
'''                  UBBillRec(1).CurRead(MCCnt) = 0
'''                End If
'''                If UBBillRec(1).PrevRead(MCCnt) < 0 Then
'''                  UBBillRec(1).PrevRead(MCCnt) = 0
'''                End If
'''                CurReadAmt& = UBBillRec(1).CurRead(MCCnt)
'''                PreReadAmt& = UBBillRec(1).PrevRead(MCCnt)
'''                If CurReadAmt& <> PreReadAmt& Then
'''                  If Not ConsumpFlag Then
'''                    'Print #UBRpt, UBCustRec(1).Status; Using("     #####   ", ThisCustRec&);
'''                    'Print #UBRpt, UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; "   "; Left$(UBCustRec(1).CustName, 25); "  "; Left$(UBCustRec(1).SERVADDR, 25)
'''                    ToPrintN$ = UBCustRec(1).Status + "~" + Using("     #####   ", ThisCustRec&)
'''                    ToPrintN$ = ToPrintN$ + "~" + UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB + "~" + Left$(UBCustRec(1).CustName, 25)
'''                    ToPrintN$ = ToPrintN$ + "~" + Left$(UBCustRec(1).SERVADDR, 22) + "~ ~ "
'''                    PrintOK = True
'''                    Linecnt = Linecnt + 1
'''                  End If
'''                  ConsumpFlag = True
'''                  MeterNum$ = QPTrim$(UBCustRec(1).Serv(TTRevCnt).RATECODE)
'''                  If Len(MeterNum$) > 0 Then
'''                    If UBCustRec(1).LocMeters(MCCnt).NumUser > 1 Then
'''                      MeterNum$ = MeterNum$ + "*" + QPTrim$(Str$(UBCustRec(1).LocMeters(MCCnt).NumUser))
'''                    End If
'''                    RSet Temp2$ = MeterNum$
'''                  End If
'''                  ConsumpAmt& = CurReadAmt& - PreReadAmt&
'''                  '103098 Added meter roll over check to inactive consumption
'''                  If ConsumpAmt& < 0 Then       'Meter rolled over or, been misread
'''                    MaxMeterAmt& = 10& ^ (Len(Str$(PreReadAmt&)) - 1)
'''                    ConsumpAmt& = (MaxMeterAmt& - PreReadAmt&) + CurReadAmt&
'''                  End If
'''                  If InStr(UBCustRec(1).CUSTTYPE, "NON") Then
'''                    'For Nonprofits include consumption as normal   'cleveland
'''                    '040998 Made changes here
'''                    For NONRateCnt = 1 To NumOfRates
'''                      If UBRateTbls(NONRateCnt).RATECODE = UBCustRec(1).Serv(TTRevCnt).RATECODE Then
'''                        NONRate = NONRateCnt
'''                        Exit For
'''                      End If
'''                    Next
'''                    If NONRate > 0 Then
'''                      RateConsump(NONRate) = RateConsump(NONRate) + ConsumpAmt&
'''                    End If
'''                    ConsumpTot(TTRevCnt, 1) = ConsumpTot(TTRevCnt, 1) + ConsumpAmt&
'''                    Bookconsump(WhatBook).Consump(TTRevCnt) = Bookconsump(WhatBook).Consump(TTRevCnt) + ConsumpAmt&
'''                    '040998 Made changes here 'cleveland
'''                  Else          'add consumption to inactives
'''                    ConsumpTot(TTRevCnt, 2) = ConsumpTot(TTRevCnt, 2) + ConsumpAmt&
'''                  End If
'''                  'Print #UBRpt, RevDesc(TTRevCnt); Tab(14); Temp2$; Tab(30); Using(fmt$(2), CurReadAmt&); Tab(42); Using(fmt$(2), PreReadAmt&); Tab(54); Using(fmt$(2), ConsumpAmt&)
'''                  ToPrintR$ = ToPrintR$ + RevDesc(TTRevCnt) + "~" + Temp2$ + "~" + Using(fmt$(2), CurReadAmt&) + "~" + Using(fmt$(2), PreReadAmt&) + "~" + Using(fmt$(2), ConsumpAmt&) + "~ ~ ~ ~ ~ ~ "
'''                  Linecnt = Linecnt + 1
'''                End If
'''                End If
'''                Next ' Else
'''                 ' If (Len(LocMeterType$) > 0) And (LocMeterType$ = UBCustRec(1).Serv(TTRevCnt).RMtrType) Then
'''                  If PrintOK = True Then
'''                    ToPrintR$ = ToPrintR$ + " ~ ~ 3~ ~ ~ ~ ~ ~ ~ ~"
'''                  End If
'''                 ' End If
'''                  'ToPrintR$ = ToPrintR$ + " ~ ~ 3~ ~ ~ ~ ~ ~ ~ ~"
'''
'''
'''
'''          Next
'''
'''        End If
'''        If ConsumpFlag And UBCustRec(1).Status <> "A" Then
'''          ConsumpFlag = False
'''          'Print #UBRpt, "**** Consumption Noted on an Inactive Account. ****"
'''          ToPrintM$ = "**** Consumption Noted on an Inactive Account. ****"
'''          Linecnt = Linecnt + 1
'''          If Not SkipSeparator Then
'''            'Print #UBRpt, fmt$(0)
'''            Linecnt = Linecnt + 1
'''          End If
'''        ElseIf ConsumpFlag Then
'''          'Customer Status is "A"
'''          'This happens when a cust has consumption and there rate code
'''          'has a zero calc amount. "i.e. a Church or other nonprofit"
'''          If InStr(UBCustRec(1).CUSTTYPE, "NON") Then
'''            'Print #UBRpt, "*** NON-PROFIT ***"
'''            ToPrintM$ = "*** NON-PROFIT ***"
'''            Linecnt = Linecnt + 1
'''          End If
'''          ConsumpFlag = False
'''          If Not SkipSeparator Then
'''            'Print #UBRpt, fmt$(0)
'''            Linecnt = Linecnt + 1
'''          End If
'''        End If
'''      'ElseIf (TRevCnt = NumOfRevs) And UBBillRec(1).TransAmt > 0 Then
'''      ElseIf (TRevCnt = 15) And UBBillRec(1).TransAmt > 0 Then
'''
'''        '102998  Moved tax printing to here "now prints one tax line per customer
'''        CTaxAmt# = 0
'''        For TXCnt = 1 To 15
'''          If UBBillRec(1).TaxAmt(TXCnt) > 0 Then
'''            CTaxAmt# = Round#(CTaxAmt# + UBBillRec(1).TaxAmt(TXCnt))
'''          End If
'''        Next
'''        If CTaxAmt# > 0 Then
'''          'Print #UBRpt, " Tax"; Tab(69); Using(fmt$(3), CTaxAmt#)
'''          ToPrintX$ = "Tax" + "~" + Using(fmt$(3), CTaxAmt#)
'''          Linecnt = Linecnt + 1
'''        Else
'''          ToPrintX$ = " ~ "
'''        End If
'''        Bills2Print = Bills2Print + 1
'''        AcctBalance# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
'''        'Print #UBRpt, Tab(5); "Current:"; Using(fmt$(6), UBBillRec(1).TransAmt);
'''        ToPrintT$ = "Current:" + "~" + Using(fmt$(6), UBBillRec(1).TransAmt) + "~"
'''        If AcctBalance# <> 0 Then
'''          'Print #UBRpt, Tab(30); "Previous:"; Using(fmt$(6), AcctBalance#);
'''          ToPrintT$ = ToPrintT$ + "Previous:" + "~" + Using(fmt$(6), AcctBalance#) + "~"
'''          TAcctBalance# = Round#(TAcctBalance# + AcctBalance#)
'''        Else
'''          ToPrintT$ = ToPrintT$ + " ~ ~"
'''        End If
'''        'Print #UBRpt, Tab(55); "Total:"; Tab(66); Using(fmt$(6), Round#(AcctBalance# + UBBillRec(1).TransAmt))
'''        ToPrintT$ = ToPrintT$ + "Total:" + "~" + Using(fmt$(6), Round#(AcctBalance# + UBBillRec(1).TransAmt))
'''        Linecnt = Linecnt + 1
'''        If Not SkipSeparator Then
'''          'Print #UBRpt, fmt$(0)
'''          Linecnt = Linecnt + 1
'''        End If
'''      End If
'''      If UBBillRec(1).TaxExempt = "Y" Then
'''        TaxExmp(TRevCnt) = Round#(TaxExmp(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
'''      End If
'''
'''    Next
'''    '020199 Moved pump code processing to here. Stops bug in getting true
'''    '       meter consumption figures.
'''    GoSub GetWhatPump
'''    If HasAPumpCode Then
'''      PumpConsump(WhatPump).CustCnt = PumpConsump(WhatPump).CustCnt + 1
'''      For MPCnt = 1 To 7
'''        PumpMtrOK = False
'''        CubMtr = False
'''        LocMeterType$ = QPTrim$(UBCustRec(1).LocMeters(MPCnt).MTRType)
'''        Select Case LocMeterType$
'''        Case "C", "S", "W", "T"
'''          PumpMtrOK = True
'''        End Select
'''        If PumpMtrOK Then
'''          MeterMulti& = UBCustRec(1).LocMeters(MPCnt).MTRMulti
'''          If UBCustRec(1).LocMeters(MPCnt).MTRUnit = "C" Then
'''            CubMtr = True
'''          End If
'''          If MeterMulti& <= 0 Then MeterMulti& = 1
'''          ReadAmt& = UBBillRec(1).CurRead(MPCnt) - UBBillRec(1).PrevRead(MPCnt)
'''          If ReadAmt& < 0 Then  'Meter rolled over or, been misread
'''            MaxMeterAmt& = 10& ^ (Len(Str$(UBBillRec(1).PrevRead(MPCnt))) - 1)
'''            ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MPCnt)) + UBBillRec(1).CurRead(MPCnt)
'''          End If
'''          If CubMtr Then
'''            ReadAmt& = ReadAmt& * 7.481
'''          End If
'''
'''          PumpConsump(WhatPump).Consump = PumpConsump(WhatPump).Consump + (ReadAmt& * MeterMulti&)
'''        End If
'''
'''      Next
'''    End If
'''    If PrintOK = True Then
'''      ToPrint$ = ToPrintN$ + "~" + ToPrintR$ + "~" + ToPrintX$ + "~" + ToPrintT$ + "~" + ToPrintM$
'''      Print #UBRpt, ToPrint$
'''    End If
'''      ToPrint$ = ""
'''      ToPrintN$ = ""
'''      ToPrintR$ = ""
'''      ToPrintX$ = ""
'''      ToPrintT$ = ""
'''      ToPrintM$ = ""
'''
'''
'''SkipEm:
''''    If AskAbandonPrint% Then
''''      UBLog "ABORTED: Prebilling report"
''''      UBLog "Closing files."
''''      Close
''''      AbortFlag = True
''''      Exit For
''''    End If
''''    ShowPctComp cnt, NumOfRecs
'''    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
'''    If FrmShowPctComp.Out = True Then
'''      Close
'''      FrmShowPctComp.Out = False
'''      GoTo ExitPreReport
'''    End If
'''  Next
'''  If AbortFlag Then GoTo ExitPreReport
'''
'''  'Print #UBRpt, FF$
'''
'''  GoSub TitleLine
''''  Print #UBRpt, "Billing Grand Totals"
''''  If TennFlag Then
''''    Print #UBRpt, "                                Inactive          Taxed      NONTax     FlatRate"
''''    Print #UBRpt, "Revenue/Tax        Consump       Consump         Amount      Amount      Amount"
''''  Else
''''    Print #UBRpt, "                                 Inactive                             Flat Rate"
''''    Print #UBRpt, "Revenue/Tax    Consumption      Consumption            Amount           Amount"
''''  End If
''''  Print #UBRpt, fmt$(0)
''''
'''  TotalFlatAmt# = 0
'''  TotalRevAmt# = 0
'''  TotalTaxAmt# = 0
'''
''''  For RaCnt = 1 To NumOfRevs
''''    If TennFlag Then
''''      Print #UBRpt, UBSetUpRec(1).Revenues(RaCnt).REVNAME; Using(fmt$(4), ConsumpTot(RaCnt, 1)); Tab(30); Using(fmt$(4), ConsumpTot(RaCnt, 2));
''''      If TaxTotals(RaCnt) > 0 Then
''''        Print #UBRpt, Tab(44); Using(fmt$(1), Round#(RevTotals(RaCnt) - FlatTotals(RaCnt) - TaxExmp(RaCnt))); Tab(56); Using(fmt$(1), TaxExmp(RaCnt)); Tab(68); Using(fmt$(1), FlatTotals(RaCnt))
''''      Else
''''        Print #UBRpt, Tab(44); Using(fmt$(1), Round#(RevTotals(RaCnt) - FlatTotals(RaCnt))); Tab(68); Using(fmt$(1), FlatTotals(RaCnt))
''''      End If
''''    Else
''''      Print #UBRpt, UBSetUpRec(1).Revenues(RaCnt).REVNAME; Using(fmt$(4), ConsumpTot(RaCnt, 1)); Tab(33); Using(fmt$(4), ConsumpTot(RaCnt, 2));
''''      Print #UBRpt, Tab(50); Using(fmt$(1), RevTotals(RaCnt) - FlatTotals(RaCnt)); Tab(67); Using(fmt$(1), FlatTotals(RaCnt))
''''    End If
''''    TotalFlatAmt# = Round#(TotalFlatAmt# + FlatTotals(RaCnt))
''''    TotalRevAmt# = Round#(TotalRevAmt# + RevTotals(RaCnt))
''''    If TaxTotals(RaCnt) > 0 Then
''''      If TennFlag Then
''''        Print #UBRpt, " Tax"; Tab(44); Using(fmt$(1), TaxTotals(RaCnt))
''''      Else
''''        Print #UBRpt, " Tax"; Tab(50); Using(fmt$(1), TaxTotals(RaCnt))
''''      End If
''''      TotalTaxAmt# = Round#(TotalTaxAmt# + TaxTotals(RaCnt))
''''    End If
''''  Next
''''  Print #UBRpt, fmt$(0)
''''  Print #UBRpt, "  PREVIOUS: "; Using(fmt$(6), TAcctBalance#);
''''  Print #UBRpt, Tab(32); "REVENUE TOTAL: "; Using(fmt$(5), Round#(TotalRevAmt# - TotalFlatAmt#))
''''  Print #UBRpt, "BILL COUNT: "; Using(fmt$(2), Bills2Print);
''''  Print #UBRpt, Tab(32); "   FLAT TOTAL: "; Using(fmt$(5), TotalFlatAmt#)
''''  Print #UBRpt, Tab(32); "    TAX TOTAL: "; Using(fmt$(5), TotalTaxAmt#)
''''  Print #UBRpt, Tab(32); "BILLING TOTAL: "; Using(fmt$(5), Round#(TotalRevAmt# + TotalTaxAmt#))
''''  Print #UBRpt, FF$
''''
'''  TotalRevAmt# = 0
'''
'''  GoSub RptTotRateHeader
'''
''''  For RaCnt = 1 To NumOfRates
''''    If (RateTotals(RaCnt) <> 0) Or (RateConsump(RaCnt) <> 0) Then
''''      If Len(QPTrim$(UBRateTbls(RaCnt).RATECODE)) > 0 Then
''''        Print #UBRpt, UBRateTbls(RaCnt).RATECODE; "    "; UBRateTbls(RaCnt).RATEDESC; Tab(39); Using(fmt$(4), RateConsump(RaCnt));
''''        Print #UBRpt, Tab(55); Using(fmt$(1), RateTotals(RaCnt));
''''        Print #UBRpt, Tab(69); Using(fmt$(2), RateCount(RaCnt))
''''        Linecnt = Linecnt + 1
''''        TotalRevAmt# = Round#(TotalRevAmt# + RateTotals(RaCnt))
''''        If RTaxTot(RaCnt) > 0 Then
''''          Print #UBRpt, " Tax"; Tab(55); Using(fmt$(1), RTaxTot(RaCnt))
''''          Linecnt = Linecnt + 1
''''        End If
''''        If Linecnt > MaxLines Then
''''          Print #UBRpt, FF$
''''          GoSub RptTotRateHeader
''''        End If
''''      End If
''''    End If
''''  Next
'''
''''  Print #UBRpt, fmt$(0)
''''  Print #UBRpt, Tab(36); "TAX TOTAL:"; Tab(52); Using(fmt$(5), TotalTaxAmt#)
''''  Print #UBRpt, Tab(40); "TOTAL:"; Tab(52); Using(fmt$(5), Round#(TotalRevAmt# + TotalTaxAmt#))
''''  Print #UBRpt, FF$
''''  'SortT BookConsump(1), TBooks, 0, Len(BookConsump(1)), 0, -1
''''  BookCQSort Bookconsump(), 1, TBooks
''''  GoSub BookHeader
'''
''''  For cnt = 1 To TBooks
''''    TestTot# = 0
''''    For ZCnt = 1 To NumOfRevs
''''      TestTot# = Round#(TestTot# + Bookconsump(cnt).RevAmt(ZCnt))
''''    Next
''''    If TestTot# <> 0 Then
''''      If Bookconsump(cnt).Book < 10 Then
''''        Book$ = "0" + QPTrim$(Str$(Bookconsump(cnt).Book))
''''      Else
''''        Book$ = QPTrim$(Str$(Bookconsump(cnt).Book))
''''      End If
''''      Print #UBRpt, "Book: "; Book$; "    Customers:"; Bookconsump(cnt).CustCnt
''''      TBookAmt# = 0
''''      TBTaxAmt# = 0
''''      For RCnt = 1 To NumOfRevs
''''        Print #UBRpt, RevDesc(RCnt); Tab(30); Using(fmt$(4), Bookconsump(cnt).Consump(RCnt));
''''        Print #UBRpt, Tab(59); Using("##########.##", Bookconsump(cnt).RevAmt(RCnt))
''''        TBookAmt# = Round#(TBookAmt# + Bookconsump(cnt).RevAmt(RCnt))
''''        If Bookconsump(cnt).TaxAmt(RCnt) > 0 Then
''''          Print #UBRpt, " Tax"; Tab(60); Using(fmt$(1), Bookconsump(cnt).TaxAmt(RCnt))
''''          TBTaxAmt# = Round#(TBTaxAmt# + Bookconsump(cnt).TaxAmt(RCnt))
''''          Linecnt = Linecnt + 1
''''        End If
''''        Linecnt = Linecnt + 1
''''      Next
''''      TBookGTot# = Round#(TBookGTot# + TBookAmt# + TBTaxAmt#)
''''      Print #UBRpt, Tab(42); "Book Total:"; Tab(57); Using(fmt$(5), Round#(TBookAmt# + TBTaxAmt#))
''''      If cnt < TBooks Then
''''        Print #UBRpt, fmt$(0)
''''      End If
''''      Linecnt = Linecnt + 1
''''    End If
''''    If Linecnt > MaxLines And cnt < TBooks Then
''''      Print #UBRpt, FF$
''''      GoSub BookHeader
''''    End If
''''
''''SkipThisBook:
''''  Next
''''
''''  Print #UBRpt, fmt$(0)
''''  Print #UBRpt, Tab(35); "Books GRAND Total:"; Tab(57); Using(fmt$(5), TBookGTot#)
''''  Print #UBRpt, FF$
''''
''''  If TPumps > 0 Then
''''    GoSub PumpHeader
''''    TMMConsump# = 0
''''    For cnt = 1 To TPumps
''''      Print #UBRpt, PumpConsump(cnt).PumpCode; Tab(30); Using("###########", PumpConsump(cnt).CustCnt); Tab(60); PumpConsump(cnt).Consump
''''      TMMConsump# = TMMConsump# + PumpConsump(cnt).Consump
''''    Next
''''    Print #UBRpt, fmt$(0)
''''    Print #UBRpt, Tab(35); "Pump Code Total:"; Tab(60); Using("###########", TMMConsump#)
''''  End If
''''
'''  Close
'''
'''  UBLog "Finished writing prebilling report."
'''  Select Case Choice
'''  Case 1
'''    RptText$ = "(Customer"
'''  Case 2
'''    RptText$ = "(Account"
'''  Case 3
'''    RptText$ = "(Location"
'''  Case 4
'''    RptText$ = "(Postal RT."
'''  Case 5
'''    RptText$ = "(ZipCode"
'''  Case 6
'''    RptText$ = "(Sequence"
'''  End Select
'''  RptText$ = RptText$ + " Order)"
'''
'''  Erase UBSetUpRec, RevDesc, UBRateTbls, RateConsump
'''  Erase fmt$, UBCustRec, UBBillRec, FlatTotals
'''  Erase RevTotals, TaxTotals, ConsumpTot
'''  Erase RateTotals, RTaxTot, Bookconsump, Indexarray
'''  Erase RateCount, ProrateServ
'''  Erase PumpConsump, TaxExmp
'''
'''  If Not AbortFlag Then
'''    'ViewPrint "UBPREBIL.RPT", "Pre-Billing Report " + RptText$
'''    'PrintRptFile "Pre-Billing Report " + RptText$, "UBPREBIL.RPT", LPTPort, RetCode, EntryPoint
'''    If BookFlag Then
'''      Kill UBBillsFile
'''    End If
'''  End If
'''
'''  GoTo ExitPreReport
'''
'''PrintPreHeader:
''''  GoSub TitleLine
''''  Print #UBRpt, "Stat  Act.  Locat    Customer Name             Service Address       Prorate%"
''''  Print #UBRpt, "Revenue            R-Code     Cur Read    Pre Read     Consump        Charges"
''''  Print #UBRpt, fmt$(0)
''''  Linecnt = 5
'''Return
'''
'''GetWhatBook:
'''  BadBookFlag = False
''' WhatBook = 0
''' If Len(QPTrim$(UBCustRec(1).Book)) = 0 Then
'''   If UBCustRec(1).Status = "A" Then
'''     BadBookFlag = True
'''     'testing vvv
'''     WhatBook = 0
'''   End If
'''   GoTo ErrorBookExit
''' End If
'''
''' ThisBook = Val(UBCustRec(1).Book)
''' If TBooks > 0 Then
'''   For TBCnt = 1 To TBooks
'''     If Bookconsump(TBCnt).Book = ThisBook Then
'''       WhatBook = TBCnt
'''       Exit For
'''     End If
'''   Next
'''   If WhatBook = 0 Then
'''     TBooks = TBooks + 1
'''     ReDim Preserve Bookconsump(0 To TBooks) As BookConsumpType
'''      Bookconsump(TBooks).Book = ThisBook
'''      WhatBook = TBooks
'''    End If
'''  Else
'''    TBooks = TBooks + 1
'''    Bookconsump(TBooks).Book = ThisBook
'''    WhatBook = TBooks
'''  End If
'''
'''ErrorBookExit:
'''  Return
'''
'''GetWhatPump:
'''  HasAPumpCode = True           'assume they have a pump code
'''  WhatPump = 0
'''  If Len(QPTrim$(UBCustRec(1).PumpCode)) = 0 Then
'''    If UBCustRec(1).Status = "A" Then
'''      HasAPumpCode = False      'no pump code
'''      WhatPump = 0
'''    End If
'''    GoTo PumpCodeReturn
'''  End If
'''
'''  CustPump$ = UCase$(QPTrim$(UBCustRec(1).PumpCode))
'''
'''  'IF CustPump$ = "34" THEN STOP
'''
'''  If Len(CustPump$) > 0 Then
'''    For TBCnt = 1 To TPumps
'''      ThisPump$ = QPTrim$(PumpConsump(TBCnt).PumpCode)
'''      If ThisPump$ = CustPump$ Then
'''        WhatPump = TBCnt
'''        Exit For
'''      End If
'''    Next
'''    If WhatPump = 0 Then
'''      TPumps = TPumps + 1
'''      ReDim Preserve PumpConsump(0 To TPumps) As PumpConsumpType
'''      PumpConsump(TPumps).PumpCode = CustPump$
'''      WhatPump = TPumps
'''    End If
'''  Else
'''    TPumps = TPumps + 1
'''    PumpConsump(TPumps).PumpCode = CustPump$
'''    WhatPump = TPumps
'''  End If
'''
'''PumpCodeReturn:
'''  Return
'''
'''RptTotRateHeader:
''''  GoSub TitleLine
''''  Print #UBRpt,
''''  Print #UBRpt, "Report Totals by Rate Code"
''''  Print #UBRpt,
''''  Print #UBRpt, "Code      Rate Description            Consumption           Amount      Bills"
''''  Print #UBRpt, fmt$(0)
''''  Linecnt = 5
'''  Return
'''
'''BookHeader:
''''  GoSub TitleLine
''''  Print #UBRpt, "Report Totals by Book"
''''  Print #UBRpt,
''''  Print #UBRpt, "Book"
''''  Print #UBRpt, "Revenue                      Consumption                         Amount"
''''  Print #UBRpt, fmt$(0)
''''  Linecnt = 7
'''  Return
'''
'''PumpHeader:
''''  GoSub TitleLine
''''  Print #UBRpt, "Report Totals by Pump Code"
''''  Print #UBRpt,
''''  Print #UBRpt, "PumpCode                  Customer Count                    Consumption"
''''  Print #UBRpt, fmt$(0)
''''  Linecnt = 6
'''  Return
'''
'''TitleLine:
''''  PageNo = PageNo + 1
''''  Print #UBRpt, "Utility Pre-Billing Report.  "; TownName$; Tab(70); "Page: "; PageNo
''''  Print #UBRpt, TheDate$
'''  Return
'''ErrorAbortExit:
'''  Close
'''
'''ExitPreReport:
'''  UBLog "OUT: Prebilling Report" + CrLf$
'''
'''End Sub
'''
'''Private Sub PreBillReport2()
'''  Dim Temp2 As String, NumOfRevs As Integer, NumOfRates As Integer
'''  Dim UBRateTblRecLen As Integer, RateFile As Integer, cnt As Long
'''  Dim UBSetupLen As Integer, MowFlag As Boolean, TennFlag As Boolean
'''  Dim TempRev As String, DoFuelAdjFlag As Boolean, SkipInactive As Boolean
'''  Dim SkipSeparator As Boolean, ThisBook As Integer, BookNum As Integer
'''  Dim BookFlag As Boolean, ThisCycle As Integer, CycleFlag As Boolean
'''  Dim SeqFlag As String, Choice As Integer, FuelAdjAmt As Double
'''  Dim IndexName As String, UsingAcct As Boolean, IdxTypeText As String
'''  Dim AbortFlag As Boolean, TheDate As String, UBCustRecLen As Integer
'''  Dim UBBillRecLen As Integer, TBooks As Integer, NumOfRecs As Long
'''  Dim Handle As Integer, IdxRecLen As Integer, lcnt As Long
'''  Dim UBBill As Integer, UBCust As Integer, UBRpt As Integer
'''  Dim ThisCustRec As Long, BillTo As String, BadBookFlag As Boolean
'''  Dim WhatBook As Integer, FRCnt As Integer, WhatService As Integer
'''  Dim Multi As Integer, FlatAmt As Double, WhatRate As Integer
'''  Dim DoneOne As Boolean, TRevCnt As Integer, IFlag As Boolean
'''  Dim TRateCnt As Integer, MINAMT As Long, PrintedRevAmt As Boolean
'''  Dim MCCnt As Integer, CubMtr As Boolean, LocMeterType As String
'''  Dim MeterMulti As Long, MeterNum As String, ReadAmt As Long
'''  Dim MaxMeterAmt As Long, Consump As Long, ThisMeterUseCnt As Integer
'''  Dim AvgUse As Long, HiConsump As Long, LowConsump As Long
'''  Dim TTRevCnt As Integer, CurReadAmt As Long, PreReadAmt As Long
'''  Dim ConsumpFlag As Boolean, ConsumpAmt As Long, NONRateCnt As Integer
'''  Dim NONRate As Integer, CTaxAmt As Double, TXCnt As Integer
'''  Dim Bills2Print As Integer, AcctBalance As Double, WhatPump As Integer
'''  Dim TAcctBalance As Double, HasAPumpCode As Boolean, MPCnt As Integer
'''  Dim PumpMtrOK As Boolean, TotalFlatAmt As Double, TotalRevAmt As Double
'''  Dim TotalTaxAmt As Double, RaCnt As Integer, TestTot As Double
'''  Dim ZCnt As Integer, Book As String, TBookAmt As Double, TPumps As Integer
'''  Dim TBTaxAmt As Double, RCnt As Integer, TBookGTot As Double
'''  Dim TMMConsump As Double, RptText As String, TBCnt As Integer
'''  Dim CustPump As String, ThisPump As String
'''  UBLog "IN: Prebilling Report"
''''  If Exist("UBBILLS.DAT") And Exist("UBBILLS.PRN") Then
''''    UBLog "ERROR: UNPOSTED BILLING DETECTED!"
''''    UBLog "ASKING USER WANT TO CONTINUE?"
''''    OK = PreBillYouSure%
''''    If Not OK Then
''''      UBLog "USER ABORTED PREBILLING."
''''      AbortFlag = True
''''      GoTo ExitPreReport
''''    Else
''''      UBLog "USER WANTS TO CONTINUE!"
''''      KillFile ("UBBILLS.PRN")
''''    End If
''''  End If
'''  PageNo = 0
'''  Temp2$ = Space$(12)
'''  NumOfRevs = MaxRevsCnt        'assume max munber of revenue sources
'''  NumOfRates = GetNumRateRecs%
'''  ReDim UBRateTbls(1 To NumOfRates) As UBRateTblRecType
'''  UBRateTblRecLen = Len(UBRateTbls(1))
'''
'''  ReDim RateConsump(1 To NumOfRates) As Double
'''
'''  RateFile = FreeFile
'''  Open "UBRATE.DAT" For Random Shared As RateFile Len = UBRateTblRecLen
'''  For cnt = 1 To NumOfRates
'''    Get RateFile, cnt, UBRateTbls(cnt)
'''  Next
'''  Close
'''
'''  'SortT UBRateTbls(1), NumOfRates, 0, UBRateTblRecLen, 0, 4
'''  RateQSort UBRateTbls(), 1, NumOfRates
''''  SortT MDateIdx(1), FoundCnt, 0, 4, 0, -1
''''  'Elemen, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
''''  QPrintRC "      Writing Index Records      ", 11, 25, -1
''''''  IndexName$ = TempIndexName
''''''  KillFile IndexName$
''''''  IHandle = FreeFile
''''''    'FCreate IndexName$
''''''  Open IndexName$ For Random Shared As IHandle Len = 4
''''''  For cnt = 1 To FoundCnt
''''''    CRec =
''''''    Put IHandle, cnt, CRec
''''''    'ShowPctComp cnt, NumCustRecs                'show user percentage complete
''''''  Next
''''''  Close IHandle
''''''
''''''  Erase UBCustRec, MDateIdx
'''
'''  ReDim ProrateServ(1 To 15) As Integer
'''
'''  ReDim UBSetUpRec(1) As UBSetupRecType
'''  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
'''
'''  TownName$ = UBSetUpRec(1).UTILNAME
'''  If InStr(TownName$, "MOWAS") > 0 Then
'''    MowFlag = True
'''  End If
'''
'''  If UBSetUpRec(1).DEFSTATE = "TN" Then
'''    TennFlag = True
'''  End If
'''
'''  ReDim RevDesc(1 To MaxRevsCnt) As String * 12
'''  For cnt = 1 To MaxRevsCnt     'find last active revenue
'''    TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(cnt).REVNAME)
'''    If Len(TempRev$) = 0 Then
'''      NumOfRevs = cnt - 1       'set actual number of revenues
'''      Exit For
'''    Else        'build revenue description lines
'''      LSet RevDesc(cnt) = UCase$(TempRev$)
'''      If InStr(RevDesc(cnt), "ELECTRIC") Then
'''        DoFuelAdjFlag = True
'''      End If
'''    End If
'''  Next
'''  '111398 Prorate
'''  For cnt = 1 To MaxRevsCnt
'''    If UBSetUpRec(1).Revenues(cnt).ProRate = "Y" Then
'''      ProrateServ(cnt) = True
'''    End If
'''  Next
'''
'''  If UBSetUpRec(1).SkipInactive = "Y" Then
'''    SkipInactive = True
'''  End If
'''
'''  If UBSetUpRec(1).SkipSeparator = "Y" Then
'''    SkipSeparator = True
'''  End If
'''
'''  If UBSetUpRec(1).PreByBook = "Y" Then
'''    ThisBook = Val(fptxtRoute1)
'''    If ThisBook = 99 Then
'''      ThisBook = -1
'''    End If
'''    BookNum = ThisBook
'''    If ThisBook = -1 Then
'''      BookFlag = False
'''    ElseIf ThisBook <= 0 Then
'''      GoTo ExitPreReport
'''    Else
'''      BookFlag = True
'''    End If
'''  ElseIf UBSetUpRec(1).BILLCYCL = "Y" Then
'''    ThisCycle = Val(fptxtRoute1)
'''    If ThisCycle <= 0 Then
'''      GoTo ExitPreReport
'''    Else
'''      CycleFlag = True
'''    End If
'''  End If
'''
'''  If UBSetUpRec(1).UseSeq = "Y" Then
'''    SeqFlag$ = "Y"
'''  End If
'''  FrmShowPctComp.Label1 = "Creating PreBilling Report"
'''  FrmShowPctComp.Show , Me
'''
'''Restart:
'''  Choice = fpcboPrintOrder.ListIndex + 1
'''  'GetPreBillOrder Choice, ExitFlag, SeqFlag$
'''
''' 'If ExitFlag Then GoTo ExitPreReport
'''  If DoFuelAdjFlag Then
'''    FuelAdjAmt# = Val(fptxtAdjustment)
'''    UBLog "Fuel adjustment factor:" + Str$(FuelAdjAmt#)
'''  Else
'''    FuelAdjAmt# = 0
'''  End If
'''
'''  If FuelAdjAmt# = -10000 Then GoTo Restart
'''
'''  Select Case Choice
'''  Case 0
'''    'ExitFlag = True
'''  Case 1        'Name
'''    IndexName$ = NameIndexFile
'''    'OkFlag = True
'''  Case 2        'Acct
'''    IndexName$ = ""
'''    UsingAcct = True
'''    'OkFlag = True
'''  Case 3        'Location
'''    IndexName$ = BookIndexFile
'''    'OkFlag = True
'''  Case 4        'Postal Route
'''    IdxTypeText$ = "Postal Route"
'''    MakePostalIndex IdxTypeText$
'''    IndexName$ = TempIndexName
'''    'OkFlag = True
'''  Case 5        'ZipCode
'''    IdxTypeText$ = "Zip-Code"
'''    'this mowflag for zip index doesn't matter cause both index
'''    'routines do same thing now.
'''    If MowFlag Then
'''      MakeMowZipCodeIndex IdxTypeText$
'''    Else
'''      MakeZipCodeIndex IdxTypeText$
'''    End If
'''    IndexName$ = TempIndexName
'''    'OkFlag = True
'''  Case 6        'Sequence number
'''    IdxTypeText$ = "Sequence Number"
'''    MakeSequenceIndex IdxTypeText$, Me
'''    IndexName$ = TempIndexName
'''    'OkFlag = True
'''  End Select
'''  MakeBillFile AbortFlag, FuelAdjAmt#, ThisCycle, ThisBook
'''
'''  If AbortFlag Then GoTo ExitPreReport
'''
'''  MaxLines = 54
'''
'''  ReDim fmt$(0 To 6)
'''  fmt$(0) = String$(80, "-")
'''  fmt$(1) = "#########.##"
'''  fmt$(2) = "#########"
'''  fmt$(3) = "######.##"
'''  fmt$(4) = "###########"
'''  fmt$(5) = "$###,###,###.##"
'''  fmt$(6) = "$#,###,###.##"
'''
'''  TheDate$ = "Date: " + Date$
'''
'''  ReDim UBCustRec(1 To 2) As NewUBCustRecType
'''  UBCustRecLen = Len(UBCustRec(1))
'''
'''  ReDim UBBillRec(1) As UBTransRecType
'''  UBBillRecLen = Len(UBBillRec(1))
'''
'''  ReDim FlatTotals(1 To NumOfRevs) As Double
'''  '021998 added flat revenue totals
'''  ReDim RevTotals(1 To NumOfRevs) As Double     'Revenue total amts
'''  '052097 added tax by revenue totals
'''  ReDim TaxTotals(1 To NumOfRevs) As Double     'Tax total amts
'''  ReDim ConsumpTot(1 To NumOfRevs, 1 To 2) As Double            'Consumption total amts
'''  ReDim RateConsump(1 To NumOfRates) As Double
'''  '012698 Added bill count by rate code
'''  ReDim RateCount(1 To NumOfRates) As Long
'''  ReDim RateTotals(1 To NumOfRates) As Double   'Rates total amts
'''  '052097 added tax by rate code totals
'''  ReDim RTaxTot(1 To NumOfRates) As Double      'Rates Tax total amts
'''  '052097 added tax by book totals to type def
'''  ReDim Bookconsump(0 To 1) As BookConsumpType  'Consumption by book
'''  ReDim PumpConsump(0 To 1) As PumpConsumpType  'Consumption by pump code
'''  ReDim TaxExmp(0 To NumOfRevs) As Double
'''
'''  TBooks = 0
'''  If UsingAcct Then
'''    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
'''  Else          'load the index
'''    UBLog "Loading index file: " + IndexName$
'''    IdxRecLen = 4
'''    NumOfRecs = FileSize(IndexName$) \ 4
'''    ReDim Indexarray(1 To NumOfRecs) As UBCustIndexRecType
'''    Handle = FreeFile
'''    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
'''    For lcnt& = 1 To NumOfRecs
'''      Get #Handle, lcnt&, Indexarray(lcnt&)
'''    Next
'''    Close Handle
'''    'FGetAH IndexName$, IndexArray(1), 4, NumOfRecs
'''  End If
'''
'''  UBBill = FreeFile
'''  Open UBBillsFile For Random Shared As UBBill Len = UBBillRecLen
'''  UBCust = FreeFile
'''  Open "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
'''  UBRpt = FreeFile
'''  Open "UBPREBIL.RPT" For Output As UBRpt
'''
'''  'BlockClear
'''  'ShowProcessingScrn "Processing Pre-Billing Report"
'''  UBLog "Writing prebilling report to disk."
'''
'''  GoSub PrintPreHeader
'''  For cnt = 1 To NumOfRecs
'''    If UsingAcct Then
'''      ThisCustRec& = cnt
'''    Else
'''      ThisCustRec& = Indexarray(cnt).RecNum
'''    End If
'''
'''    Get UBCust, ThisCustRec&, UBCustRec(1)
'''
'''    If UBCustRec(1).DelFlag Then
'''      GoTo SkipEm
'''    End If
'''
'''    If SkipInactive And UBCustRec(1).Status <> "A" Then
'''      GoTo SkipEm
'''    ElseIf UBCustRec(1).Status = "F" Then       'skip over final's
'''      GoTo SkipEm
'''    ElseIf UBCustRec(1).Status = "B" Then       'skip over B-Status
'''      GoTo SkipEm
'''    End If
'''    If BookFlag Then
'''      If Val(UBCustRec(1).Book) <> ThisBook Then
'''        GoTo SkipEm
'''      End If
'''    End If
'''
'''    If CycleFlag Then
'''      If UBCustRec(1).BILLCYCL <> ThisCycle Then
'''        GoTo SkipEm
'''      End If
'''    End If
'''
'''    Get UBBill, ThisCustRec&, UBBillRec(1)
'''
'''    If Linecnt >= MaxLines Then
'''      Print #UBRpt, FF$
'''      GoSub PrintPreHeader
'''    End If
'''
'''    If UBBillRec(1).ActiveFlag <> 0 Then
'''      If UBCustRec(1).BillTo = "O" Then
'''        BillTo$ = " O"
'''      Else
'''        BillTo$ = " C"
'''      End If
'''      GoSub GetWhatBook
'''      If BadBookFlag Then
'''        If ErrorScrn(2, ThisCustRec&) Then
'''          AbortFlag = True
'''          Exit For
'''        End If
'''      End If
'''      Bookconsump(WhatBook).CustCnt = Bookconsump(WhatBook).CustCnt + 1
'''      Print #UBRpt, UBCustRec(1).Status; Using("  #####  ", ThisCustRec&);
'''      Print #UBRpt, UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; "  "; Left$(UBCustRec(1).CustName, 25); " "; Left$(UBCustRec(1).SERVADDR, 22); " ";
'''      Print #UBRpt, Using("   ###", UBBillRec(1).ProRatePCT); "%";
'''      Print #UBRpt, BillTo$
'''      Linecnt = Linecnt + 1
'''      For FRCnt = 1 To 4
'''        WhatService = UBCustRec(1).FlatRates(FRCnt).REVSRC
'''        If UBCustRec(1).FlatRates(FRCnt).FRAMT <> 0 And WhatService > 0 Then
'''          Multi = UBCustRec(1).FlatRates(FRCnt).NumMin
'''          If Multi < 1 Then Multi = 1
'''          FlatAmt# = Round#(UBCustRec(1).FlatRates(FRCnt).FRAMT * Multi)
'''          '021998 Added flat rate summaries
'''          FlatTotals(WhatService) = Round#(FlatTotals(WhatService) + FlatAmt#)
'''        End If
'''      Next
'''      '102798 Added to skip accts that don't have a book/seq no. "J.R."
'''    ElseIf Len(QPTrim$(UBCustRec(1).Book)) = 0 And Len(QPTrim$(UBCustRec(1).SEQNUMB)) = 0 Then
'''      GoTo SkipEm
'''    End If
'''    WhatRate = 0
'''    DoneOne = False
'''    For TRevCnt = 1 To NumOfRevs
'''      If TRevCnt = 2 And UBBillRec(1).PenAtBill = -1 Then
'''        IFlag = True
'''      Else
'''        IFlag = False
'''      End If
'''      WhatRate = 0
'''      If UBBillRec(1).RevAmt(TRevCnt) <> 0 Then
'''        DoneOne = False
'''        Print #UBRpt, RevDesc(TRevCnt);
'''        '102198 Moved out of meter loop, Stoped multi meter tax report bug
'''        If UBBillRec(1).TaxAmt(TRevCnt) > 0 Then
'''          TaxTotals(TRevCnt) = Round#(TaxTotals(TRevCnt) + UBBillRec(1).TaxAmt(TRevCnt))
'''        End If
'''        For TRateCnt = 1 To NumOfRates
'''          If UBRateTbls(TRateCnt).RATECODE = UBCustRec(1).Serv(TRevCnt).RATECODE Then
'''            MINAMT& = UBRateTbls(TRateCnt).MINUNITS
'''            WhatRate = TRateCnt
'''            '102198 Moved from meter loop, Stops multi meter tax report bug
'''            RTaxTot(WhatRate) = Round#(RTaxTot(WhatRate) + UBBillRec(1).TaxAmt(TRevCnt))
'''            Exit For
'''          End If
'''        Next
'''        If UBSetUpRec(1).Revenues(TRevCnt).UseMtr = "Y" Then
'''          RevTotals(TRevCnt) = Round#(RevTotals(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
'''          Bookconsump(WhatBook).RevAmt(TRevCnt) = Round#(Bookconsump(WhatBook).RevAmt(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
'''          Bookconsump(WhatBook).TaxAmt(TRevCnt) = Round#(Bookconsump(WhatBook).TaxAmt(TRevCnt) + UBBillRec(1).TaxAmt(TRevCnt))
'''
'''          '02-20-97 Add revenue totals by rate code
'''          If WhatRate > 0 Then
'''            RateTotals(WhatRate) = Round#(RateTotals(WhatRate) + UBBillRec(1).RevAmt(TRevCnt))
'''          End If
'''          PrintedRevAmt = False
'''          For MCCnt = 1 To 7
'''            CubMtr = False
'''            LocMeterType$ = QPTrim$(UBCustRec(1).LocMeters(MCCnt).MTRType)
'''            MeterMulti& = UBCustRec(1).LocMeters(MCCnt).MTRMulti
'''            '063098 Added adjustment for cubic meters in consumption totals
'''            If UBCustRec(1).LocMeters(MCCnt).MTRUnit = "C" Then
'''              CubMtr = True
'''            End If
'''            If MeterMulti& <= 0 Then MeterMulti& = 1
'''            If (Len(LocMeterType$) > 0) And (LocMeterType$ = UBCustRec(1).Serv(TRevCnt).RMtrType) Then
'''              DoneOne = True
'''              MeterNum$ = QPTrim$(UBCustRec(1).Serv(TRevCnt).RATECODE)
'''              'use the Meternum$ to hold the rate code temporarily
'''              If Len(MeterNum$) > 0 Then
'''                If UBCustRec(1).LocMeters(MCCnt).NumUser > 1 Then
'''                  MeterNum$ = MeterNum$ + "*" + QPTrim$(Str$(UBCustRec(1).LocMeters(MCCnt).NumUser))
'''                End If
'''                RSet Temp2$ = MeterNum$
'''              End If
'''              ReadAmt& = UBBillRec(1).CurRead(MCCnt) - UBBillRec(1).PrevRead(MCCnt)
'''              If ReadAmt& < 0 Then              'Meter rolled over or, been misread
'''                MaxMeterAmt& = 10& ^ (Len(Str$(UBBillRec(1).PrevRead(MCCnt))) - 1)
'''                ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MCCnt)) + UBBillRec(1).CurRead(MCCnt)
'''              End If
'''              If CubMtr Then
'''                ReadAmt& = ReadAmt& * 7.481
'''              End If
'''              RateConsump(WhatRate) = RateConsump(WhatRate) + (ReadAmt& * MeterMulti&)
'''              RateCount(WhatRate) = RateCount(WhatRate) + 1
'''              Bookconsump(WhatBook).Consump(TRevCnt) = Bookconsump(WhatBook).Consump(TRevCnt) + (ReadAmt& * MeterMulti&)
'''              ConsumpTot(TRevCnt, 1) = ConsumpTot(TRevCnt, 1) + (ReadAmt& * MeterMulti&)
'''              Consump& = ReadAmt& * MeterMulti&
'''              ThisMeterUseCnt = UBCustRec(1).LocMeters(MCCnt).UseCnt
'''              If ThisMeterUseCnt <= 0 Then ThisMeterUseCnt = 1
'''              AvgUse& = UBCustRec(1).LocMeters(MCCnt).AvgUse
'''              If AvgUse& > 0 Then
'''                HiConsump& = Round#(AvgUse& * (UBSetUpRec(1).HighRead * 0.01))
'''                LowConsump& = Round#(AvgUse& * (UBSetUpRec(1).LowRead * 0.01))
'''              End If
'''              Print #UBRpt, Tab(14); Temp2$; Tab(30); Using(fmt$(2), UBBillRec(1).CurRead(MCCnt)); Tab(42); UBBillRec(1).PrevRead(MCCnt); Tab(54); ReadAmt& * MeterMulti&;
'''              If UBCustRec(1).EstFlag = "E" Then
'''                Print #UBRpt, " E";             'Est. Reading
'''              ElseIf Consump& < LowConsump& Then
'''                Print #UBRpt, " L";             'Low reading
'''              ElseIf Consump& > HiConsump& Then
'''                Print #UBRpt, " H";             'High Reading
'''              End If
'''              If Consump& < MINAMT& Then
'''                Print #UBRpt, " M";             'Minium Usage
'''              End If
'''              If UBBillRec(1).RevAmt(TRevCnt) > 0 And PrintedRevAmt = False Then
'''                PrintedRevAmt = True
'''                Print #UBRpt, Tab(69); Using(fmt$(3), UBBillRec(1).RevAmt(TRevCnt));
'''                If UBBillRec(1).ProRatePCT < 100 And ProrateServ(TRevCnt) Then
'''                  Print #UBRpt, "*";
'''                End If
'''                If IFlag Then
'''                  Print #UBRpt, " IR";
'''                End If
'''
'''              End If
'''              Print #UBRpt,
'''              Linecnt = Linecnt + 1
'''            End If
'''          Next
'''          '071197 Added this for mccormick. Has a sewer flat rate, Sewer is set up as
'''          '      a metered service but no meter on a flat rate charge. Charge was added
'''          '      to total, but didn't show on prebilling report.
'''          If Not DoneOne Then
'''            DoneOne = True
'''            Print #UBRpt, Tab(69); Using(fmt$(3), UBBillRec(1).RevAmt(TRevCnt));
'''            If UBBillRec(1).ProRatePCT < 100 And ProrateServ(TRevCnt) Then
'''              Print #UBRpt, "*";
'''            End If
'''            'THIS WAS REMARKED OUT, I DON'T KNOW WHY?
'''            Print #UBRpt,
'''            ''''''''''''''''''''''''''''''''''''''
'''            Linecnt = Linecnt + 1
'''          End If
'''        Else    'it's a nonmetered service
'''          ConsumpTot(TRevCnt, 1) = ConsumpTot(TRevCnt, 1) + 1
'''          If WhatRate > 0 Then
'''            RateConsump(WhatRate) = RateConsump(WhatRate) + 1
'''            RateCount(WhatRate) = RateCount(WhatRate) + 1
'''            RateTotals(WhatRate) = Round#(RateTotals(WhatRate) + UBBillRec(1).RevAmt(TRevCnt))
'''          End If
'''          Bookconsump(WhatBook).Consump(TRevCnt) = Bookconsump(WhatBook).Consump(TRevCnt) + 1
'''          Bookconsump(WhatBook).RevAmt(TRevCnt) = Round#(Bookconsump(WhatBook).RevAmt(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
'''          Bookconsump(WhatBook).TaxAmt(TRevCnt) = Round#(Bookconsump(WhatBook).TaxAmt(TRevCnt) + UBBillRec(1).TaxAmt(TRevCnt))
'''          RevTotals(TRevCnt) = Round#(RevTotals(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
'''          Print #UBRpt, Tab(69); Using(fmt$(3), UBBillRec(1).RevAmt(TRevCnt));
'''          If UBBillRec(1).ProRatePCT < 100 And ProrateServ(TRevCnt) Then
'''            Print #UBRpt, "*";
'''          End If
'''        End If
'''        If Not DoneOne Then
'''          Print #UBRpt,
'''          Linecnt = Linecnt + 1
'''        End If
'''      End If
'''      If (TRevCnt = NumOfRevs) And UBBillRec(1).TransAmt = 0 Then
'''        If UBBillRec(1).TransAmt = 0 Then       'CONSUMPTION inactive account
'''          For TTRevCnt = 1 To NumOfRevs
'''            For MCCnt = 1 To 7
'''              LocMeterType$ = QPTrim$(UBCustRec(1).LocMeters(MCCnt).MTRType)
'''              If (Len(LocMeterType$) > 0) And (LocMeterType$ = UBCustRec(1).Serv(TTRevCnt).RMtrType) Then
'''                If UBBillRec(1).CurRead(MCCnt) < 0 Then
'''                  UBBillRec(1).CurRead(MCCnt) = 0
'''                End If
'''                If UBBillRec(1).PrevRead(MCCnt) < 0 Then
'''                  UBBillRec(1).PrevRead(MCCnt) = 0
'''                End If
'''                CurReadAmt& = UBBillRec(1).CurRead(MCCnt)
'''                PreReadAmt& = UBBillRec(1).PrevRead(MCCnt)
'''                If CurReadAmt& <> PreReadAmt& Then
'''                  If Not ConsumpFlag Then
'''                    Print #UBRpt, UBCustRec(1).Status; Using("     #####   ", ThisCustRec&);
'''                    Print #UBRpt, UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; "   "; Left$(UBCustRec(1).CustName, 25); "  "; Left$(UBCustRec(1).SERVADDR, 25)
'''                    Linecnt = Linecnt + 1
'''                  End If
'''                  ConsumpFlag = True
'''                  MeterNum$ = QPTrim$(UBCustRec(1).Serv(TTRevCnt).RATECODE)
'''                  If Len(MeterNum$) > 0 Then
'''                    If UBCustRec(1).LocMeters(MCCnt).NumUser > 1 Then
'''                      MeterNum$ = MeterNum$ + "*" + QPTrim$(Str$(UBCustRec(1).LocMeters(MCCnt).NumUser))
'''                    End If
'''                    RSet Temp2$ = MeterNum$
'''                  End If
'''                  ConsumpAmt& = CurReadAmt& - PreReadAmt&
'''                  '103098 Added meter roll over check to inactive consumption
'''                  If ConsumpAmt& < 0 Then       'Meter rolled over or, been misread
'''                    MaxMeterAmt& = 10& ^ (Len(Str$(PreReadAmt&)) - 1)
'''                    ConsumpAmt& = (MaxMeterAmt& - PreReadAmt&) + CurReadAmt&
'''                  End If
'''                  If InStr(UBCustRec(1).CUSTTYPE, "NON") Then
'''                    'For Nonprofits include consumption as normal   'cleveland
'''                    '040998 Made changes here
'''                    For NONRateCnt = 1 To NumOfRates
'''                      If UBRateTbls(NONRateCnt).RATECODE = UBCustRec(1).Serv(TTRevCnt).RATECODE Then
'''                        NONRate = NONRateCnt
'''                        Exit For
'''                      End If
'''                    Next
'''                    If NONRate > 0 Then
'''                      RateConsump(NONRate) = RateConsump(NONRate) + ConsumpAmt&
'''                    End If
'''                    ConsumpTot(TTRevCnt, 1) = ConsumpTot(TTRevCnt, 1) + ConsumpAmt&
'''                    Bookconsump(WhatBook).Consump(TTRevCnt) = Bookconsump(WhatBook).Consump(TTRevCnt) + ConsumpAmt&
'''                    '040998 Made changes here 'cleveland
'''                  Else          'add consumption to inactives
'''                    ConsumpTot(TTRevCnt, 2) = ConsumpTot(TTRevCnt, 2) + ConsumpAmt&
'''                  End If
'''                  Print #UBRpt, RevDesc(TTRevCnt); Tab(14); Temp2$; Tab(30); Using(fmt$(2), CurReadAmt&); Tab(42); Using(fmt$(2), PreReadAmt&); Tab(54); Using(fmt$(2), ConsumpAmt&)
'''                  Linecnt = Linecnt + 1
'''                End If
'''              End If
'''            Next
'''          Next
'''        End If
'''        If ConsumpFlag And UBCustRec(1).Status <> "A" Then
'''          ConsumpFlag = False
'''          Print #UBRpt, "**** Consumption Noted on an Inactive Account. ****"
'''          Linecnt = Linecnt + 1
'''          If Not SkipSeparator Then
'''            Print #UBRpt, fmt$(0)
'''            Linecnt = Linecnt + 1
'''          End If
'''        ElseIf ConsumpFlag Then
'''          'Customer Status is "A"
'''          'This happens when a cust has consumption and there rate code
'''          'has a zero calc amount. "i.e. a Church or other nonprofit"
'''          If InStr(UBCustRec(1).CUSTTYPE, "NON") Then
'''            Print #UBRpt, "*** NON-PROFIT ***"
'''            Linecnt = Linecnt + 1
'''          End If
'''          ConsumpFlag = False
'''          If Not SkipSeparator Then
'''            Print #UBRpt, fmt$(0)
'''            Linecnt = Linecnt + 1
'''          End If
'''        End If
'''      ElseIf (TRevCnt = NumOfRevs) And UBBillRec(1).TransAmt > 0 Then
'''        '102998  Moved tax printing to here "now prints one tax line per customer
'''        CTaxAmt# = 0
'''        For TXCnt = 1 To 15
'''          If UBBillRec(1).TaxAmt(TXCnt) > 0 Then
'''            CTaxAmt# = Round#(CTaxAmt# + UBBillRec(1).TaxAmt(TXCnt))
'''          End If
'''        Next
'''        If CTaxAmt# > 0 Then
'''          Print #UBRpt, " Tax"; Tab(69); Using(fmt$(3), CTaxAmt#)
'''          Linecnt = Linecnt + 1
'''        End If
'''        Bills2Print = Bills2Print + 1
'''        AcctBalance# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
'''        Print #UBRpt, Tab(5); "Current:"; Using(fmt$(6), UBBillRec(1).TransAmt);
'''        If AcctBalance# <> 0 Then
'''          Print #UBRpt, Tab(30); "Previous:"; Using(fmt$(6), AcctBalance#);
'''          TAcctBalance# = Round#(TAcctBalance# + AcctBalance#)
'''        End If
'''        Print #UBRpt, Tab(55); "Total:"; Tab(66); Using(fmt$(6), Round#(AcctBalance# + UBBillRec(1).TransAmt))
'''        Linecnt = Linecnt + 1
'''        If Not SkipSeparator Then
'''          Print #UBRpt, fmt$(0)
'''          Linecnt = Linecnt + 1
'''        End If
'''      End If
'''      If UBBillRec(1).TaxExempt = "Y" Then
'''        TaxExmp(TRevCnt) = Round#(TaxExmp(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
'''      End If
'''    Next
'''    '020199 Moved pump code processing to here. Stops bug in getting true
'''    '       meter consumption figures.
'''    GoSub GetWhatPump
'''    If HasAPumpCode Then
'''      PumpConsump(WhatPump).CustCnt = PumpConsump(WhatPump).CustCnt + 1
'''      For MPCnt = 1 To 7
'''        PumpMtrOK = False
'''        CubMtr = False
'''        LocMeterType$ = QPTrim$(UBCustRec(1).LocMeters(MPCnt).MTRType)
'''        Select Case LocMeterType$
'''        Case "C", "S", "W", "T"
'''          PumpMtrOK = True
'''        End Select
'''        If PumpMtrOK Then
'''          MeterMulti& = UBCustRec(1).LocMeters(MPCnt).MTRMulti
'''          If UBCustRec(1).LocMeters(MPCnt).MTRUnit = "C" Then
'''            CubMtr = True
'''          End If
'''          If MeterMulti& <= 0 Then MeterMulti& = 1
'''          ReadAmt& = UBBillRec(1).CurRead(MPCnt) - UBBillRec(1).PrevRead(MPCnt)
'''          If ReadAmt& < 0 Then  'Meter rolled over or, been misread
'''            MaxMeterAmt& = 10& ^ (Len(Str$(UBBillRec(1).PrevRead(MPCnt))) - 1)
'''            ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MPCnt)) + UBBillRec(1).CurRead(MPCnt)
'''          End If
'''          If CubMtr Then
'''            ReadAmt& = ReadAmt& * 7.481
'''          End If
'''
'''          PumpConsump(WhatPump).Consump = PumpConsump(WhatPump).Consump + (ReadAmt& * MeterMulti&)
'''        End If
'''
'''      Next
'''    End If
'''SkipEm:
''''    If AskAbandonPrint% Then
''''      UBLog "ABORTED: Prebilling report"
''''      UBLog "Closing files."
''''      Close
''''      AbortFlag = True
''''      Exit For
''''    End If
''''    ShowPctComp cnt, NumOfRecs
'''    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
'''    If FrmShowPctComp.Out = True Then
'''      Close
'''      FrmShowPctComp.Out = False
'''      GoTo ExitPreReport
'''    End If
'''
'''  Next
'''  If AbortFlag Then GoTo ExitPreReport
'''
'''  Print #UBRpt, FF$
'''
'''  GoSub TitleLine
'''  Print #UBRpt, "Billing Grand Totals"
'''  If TennFlag Then
'''    Print #UBRpt, "                                Inactive          Taxed      NONTax     FlatRate"
'''    Print #UBRpt, "Revenue/Tax        Consump       Consump         Amount      Amount      Amount"
'''  Else
'''    Print #UBRpt, "                                 Inactive                             Flat Rate"
'''    Print #UBRpt, "Revenue/Tax    Consumption      Consumption            Amount           Amount"
'''  End If
'''  Print #UBRpt, fmt$(0)
'''
'''  TotalFlatAmt# = 0
'''  TotalRevAmt# = 0
'''  TotalTaxAmt# = 0
'''
'''  For RaCnt = 1 To NumOfRevs
'''    If TennFlag Then
'''      Print #UBRpt, UBSetUpRec(1).Revenues(RaCnt).REVNAME; Using(fmt$(4), ConsumpTot(RaCnt, 1)); Tab(30); Using(fmt$(4), ConsumpTot(RaCnt, 2));
'''      If TaxTotals(RaCnt) > 0 Then
'''        Print #UBRpt, Tab(44); Using(fmt$(1), Round#(RevTotals(RaCnt) - FlatTotals(RaCnt) - TaxExmp(RaCnt))); Tab(56); Using(fmt$(1), TaxExmp(RaCnt)); Tab(68); Using(fmt$(1), FlatTotals(RaCnt))
'''      Else
'''        Print #UBRpt, Tab(44); Using(fmt$(1), Round#(RevTotals(RaCnt) - FlatTotals(RaCnt))); Tab(68); Using(fmt$(1), FlatTotals(RaCnt))
'''      End If
'''    Else
'''      Print #UBRpt, UBSetUpRec(1).Revenues(RaCnt).REVNAME; Using(fmt$(4), ConsumpTot(RaCnt, 1)); Tab(33); Using(fmt$(4), ConsumpTot(RaCnt, 2));
'''      Print #UBRpt, Tab(50); Using(fmt$(1), RevTotals(RaCnt) - FlatTotals(RaCnt)); Tab(67); Using(fmt$(1), FlatTotals(RaCnt))
'''    End If
'''    TotalFlatAmt# = Round#(TotalFlatAmt# + FlatTotals(RaCnt))
'''    TotalRevAmt# = Round#(TotalRevAmt# + RevTotals(RaCnt))
'''    If TaxTotals(RaCnt) > 0 Then
'''      If TennFlag Then
'''        Print #UBRpt, " Tax"; Tab(44); Using(fmt$(1), TaxTotals(RaCnt))
'''      Else
'''        Print #UBRpt, " Tax"; Tab(50); Using(fmt$(1), TaxTotals(RaCnt))
'''      End If
'''      TotalTaxAmt# = Round#(TotalTaxAmt# + TaxTotals(RaCnt))
'''    End If
'''  Next
'''  Print #UBRpt, fmt$(0)
'''  Print #UBRpt, "  PREVIOUS: "; Using(fmt$(6), TAcctBalance#);
'''  Print #UBRpt, Tab(32); "REVENUE TOTAL: "; Using(fmt$(5), Round#(TotalRevAmt# - TotalFlatAmt#))
'''  Print #UBRpt, "BILL COUNT: "; Using(fmt$(2), Bills2Print);
'''  Print #UBRpt, Tab(32); "   FLAT TOTAL: "; Using(fmt$(5), TotalFlatAmt#)
'''  Print #UBRpt, Tab(32); "    TAX TOTAL: "; Using(fmt$(5), TotalTaxAmt#)
'''  Print #UBRpt, Tab(32); "BILLING TOTAL: "; Using(fmt$(5), Round#(TotalRevAmt# + TotalTaxAmt#))
'''  Print #UBRpt, FF$
'''
'''  TotalRevAmt# = 0
'''
'''  GoSub RptTotRateHeader
'''
'''  For RaCnt = 1 To NumOfRates
'''    If (RateTotals(RaCnt) <> 0) Or (RateConsump(RaCnt) <> 0) Then
'''      If Len(QPTrim$(UBRateTbls(RaCnt).RATECODE)) > 0 Then
'''        Print #UBRpt, UBRateTbls(RaCnt).RATECODE; "    "; UBRateTbls(RaCnt).RATEDESC; Tab(39); Using(fmt$(4), RateConsump(RaCnt));
'''        Print #UBRpt, Tab(55); Using(fmt$(1), RateTotals(RaCnt));
'''        Print #UBRpt, Tab(69); Using(fmt$(2), RateCount(RaCnt))
'''        Linecnt = Linecnt + 1
'''        TotalRevAmt# = Round#(TotalRevAmt# + RateTotals(RaCnt))
'''        If RTaxTot(RaCnt) > 0 Then
'''          Print #UBRpt, " Tax"; Tab(55); Using(fmt$(1), RTaxTot(RaCnt))
'''          Linecnt = Linecnt + 1
'''        End If
'''        If Linecnt >= MaxLines Then
'''          Print #UBRpt, FF$
'''          GoSub RptTotRateHeader
'''        End If
'''      End If
'''    End If
'''  Next
'''
'''  Print #UBRpt, fmt$(0)
'''  Print #UBRpt, Tab(36); "TAX TOTAL:"; Tab(52); Using(fmt$(5), TotalTaxAmt#)
'''  Print #UBRpt, Tab(40); "TOTAL:"; Tab(52); Using(fmt$(5), Round#(TotalRevAmt# + TotalTaxAmt#))
'''  Print #UBRpt, FF$
'''  'SortT BookConsump(1), TBooks, 0, Len(BookConsump(1)), 0, -1
'''  BookCQSort Bookconsump(), 1, TBooks
'''  GoSub BookHeader
'''
'''  For cnt = 1 To TBooks
'''    TestTot# = 0
'''    For ZCnt = 1 To NumOfRevs
'''      TestTot# = Round#(TestTot# + Bookconsump(cnt).RevAmt(ZCnt))
'''    Next
'''    If TestTot# <> 0 Then
'''      If Bookconsump(cnt).Book < 10 Then
'''        Book$ = "0" + QPTrim$(Str$(Bookconsump(cnt).Book))
'''      Else
'''        Book$ = QPTrim$(Str$(Bookconsump(cnt).Book))
'''      End If
'''      Print #UBRpt, "Book: "; Book$; "    Customers:"; Bookconsump(cnt).CustCnt
'''      TBookAmt# = 0
'''      TBTaxAmt# = 0
'''      For RCnt = 1 To NumOfRevs
'''        Print #UBRpt, RevDesc(RCnt); Tab(30); Using(fmt$(4), Bookconsump(cnt).Consump(RCnt));
'''        Print #UBRpt, Tab(59); Using("##########.##", Bookconsump(cnt).RevAmt(RCnt))
'''        TBookAmt# = Round#(TBookAmt# + Bookconsump(cnt).RevAmt(RCnt))
'''        If Bookconsump(cnt).TaxAmt(RCnt) > 0 Then
'''          Print #UBRpt, " Tax"; Tab(60); Using(fmt$(1), Bookconsump(cnt).TaxAmt(RCnt))
'''          TBTaxAmt# = Round#(TBTaxAmt# + Bookconsump(cnt).TaxAmt(RCnt))
'''          Linecnt = Linecnt + 1
'''        End If
'''        Linecnt = Linecnt + 1
'''      Next
'''      TBookGTot# = Round#(TBookGTot# + TBookAmt# + TBTaxAmt#)
'''      Print #UBRpt, Tab(42); "Book Total:"; Tab(57); Using(fmt$(5), Round#(TBookAmt# + TBTaxAmt#))
'''      If cnt < TBooks Then
'''        Print #UBRpt, fmt$(0)
'''      End If
'''      Linecnt = Linecnt + 1
'''    End If
'''    If Linecnt >= MaxLines And cnt < TBooks Then
'''      Print #UBRpt, FF$
'''      GoSub BookHeader
'''    End If
'''
'''SkipThisBook:
'''  Next
'''
'''  Print #UBRpt, fmt$(0)
'''  Print #UBRpt, Tab(35); "Books GRAND Total:"; Tab(57); Using(fmt$(5), TBookGTot#)
'''  Print #UBRpt, FF$
'''
'''  If TPumps > 0 Then
'''    GoSub PumpHeader
'''    TMMConsump# = 0
'''    For cnt = 1 To TPumps
'''      Print #UBRpt, PumpConsump(cnt).PumpCode; Tab(30); Using("###########", PumpConsump(cnt).CustCnt); Tab(60); PumpConsump(cnt).Consump
'''      TMMConsump# = TMMConsump# + PumpConsump(cnt).Consump
'''    Next
'''    Print #UBRpt, fmt$(0)
'''    Print #UBRpt, Tab(35); "Pump Code Total:"; Tab(60); Using("###########", TMMConsump#)
'''  End If
'''
'''  Close
'''
'''  UBLog "Finished writing prebilling report."
'''  Select Case Choice
'''  Case 1
'''    RptText$ = "(Customer"
'''  Case 2
'''    RptText$ = "(Account"
'''  Case 3
'''    RptText$ = "(Location"
'''  Case 4
'''    RptText$ = "(Postal RT."
'''  Case 5
'''    RptText$ = "(ZipCode"
'''  Case 6
'''    RptText$ = "(Sequence"
'''  End Select
'''  RptText$ = RptText$ + " Order)"
'''
'''  Erase UBSetUpRec, RevDesc, UBRateTbls, RateConsump
'''  Erase fmt$, UBCustRec, UBBillRec, FlatTotals
'''  Erase RevTotals, TaxTotals, ConsumpTot
'''  Erase RateTotals, RTaxTot, Bookconsump, Indexarray
'''  Erase RateCount, ProrateServ
'''  Erase PumpConsump, TaxExmp
'''
'''  If Not AbortFlag Then
'''   ' ViewPrint "UBPREBIL.RPT", "Pre-Billing Report " + RptText$
'''    'PrintRptFile "Pre-Billing Report " + RptText$, "UBPREBIL.RPT", LPTPort, RetCode, EntryPoint
'''    Load frmLoadingRpt
'''    ARptPreBilling.GetName "UBPREBIL.RPT"
'''    ARptPreBilling.startrpt
'''
'''    If BookFlag Then
'''      Kill UBBillsFile
'''    End If
'''  End If
'''
'''  GoTo ExitPreReport
'''
'''PrintPreHeader:
'''  GoSub TitleLine
'''  Print #UBRpt, "Stat  Act.  Locat    Customer Name             Service Address       Prorate%"
'''  Print #UBRpt, "Revenue            R-Code     Cur Read    Pre Read     Consump        Charges"
'''  Print #UBRpt, fmt$(0)
'''  Linecnt = 5
'''Return
'''
'''GetWhatBook:
'''  BadBookFlag = False
''' WhatBook = 0
''' If Len(QPTrim$(UBCustRec(1).Book)) = 0 Then
'''   If UBCustRec(1).Status = "A" Then
'''     BadBookFlag = True
'''     'testing vvv
'''     WhatBook = 0
'''   End If
'''   GoTo ErrorBookExit
''' End If
'''
''' ThisBook = Val(UBCustRec(1).Book)
''' If TBooks > 0 Then
'''   For TBCnt = 1 To TBooks
'''     If Bookconsump(TBCnt).Book = ThisBook Then
'''       WhatBook = TBCnt
'''       Exit For
'''     End If
'''   Next
'''   If WhatBook = 0 Then
'''     TBooks = TBooks + 1
'''     ReDim Preserve Bookconsump(0 To TBooks) As BookConsumpType
'''      Bookconsump(TBooks).Book = ThisBook
'''      WhatBook = TBooks
'''    End If
'''  Else
'''    TBooks = TBooks + 1
'''    Bookconsump(TBooks).Book = ThisBook
'''    WhatBook = TBooks
'''  End If
'''
'''ErrorBookExit:
'''  Return
'''
'''GetWhatPump:
'''  HasAPumpCode = True           'assume they have a pump code
'''  WhatPump = 0
'''  If Len(QPTrim$(UBCustRec(1).PumpCode)) = 0 Then
'''    If UBCustRec(1).Status = "A" Then
'''      HasAPumpCode = False      'no pump code
'''      WhatPump = 0
'''    End If
'''    GoTo PumpCodeReturn
'''  End If
'''
'''  CustPump$ = UCase$(QPTrim$(UBCustRec(1).PumpCode))
'''
'''  'IF CustPump$ = "34" THEN STOP
'''
'''  If Len(CustPump$) > 0 Then
'''    For TBCnt = 1 To TPumps
'''      ThisPump$ = QPTrim$(PumpConsump(TBCnt).PumpCode)
'''      If ThisPump$ = CustPump$ Then
'''        WhatPump = TBCnt
'''        Exit For
'''      End If
'''    Next
'''    If WhatPump = 0 Then
'''      TPumps = TPumps + 1
'''      ReDim Preserve PumpConsump(0 To TPumps) As PumpConsumpType
'''      PumpConsump(TPumps).PumpCode = CustPump$
'''      WhatPump = TPumps
'''    End If
'''  Else
'''    TPumps = TPumps + 1
'''    PumpConsump(TPumps).PumpCode = CustPump$
'''    WhatPump = TPumps
'''  End If
'''
'''PumpCodeReturn:
'''  Return
'''
'''RptTotRateHeader:
'''  GoSub TitleLine
'''  Print #UBRpt,
'''  Print #UBRpt, "Report Totals by Rate Code"
'''  Print #UBRpt,
'''  Print #UBRpt, "Code      Rate Description            Consumption           Amount      Bills"
'''  Print #UBRpt, fmt$(0)
'''  Linecnt = 5
'''  Return
'''
'''BookHeader:
'''  GoSub TitleLine
'''  Print #UBRpt, "Report Totals by Book"
'''  Print #UBRpt,
'''  Print #UBRpt, "Book"
'''  Print #UBRpt, "Revenue                      Consumption                         Amount"
'''  Print #UBRpt, fmt$(0)
'''  Linecnt = 7
'''  Return
'''
'''PumpHeader:
'''  GoSub TitleLine
'''  Print #UBRpt, "Report Totals by Pump Code"
'''  Print #UBRpt,
'''  Print #UBRpt, "PumpCode                  Customer Count                    Consumption"
'''  Print #UBRpt, fmt$(0)
'''  Linecnt = 6
'''  Return
'''
'''TitleLine:
'''  PageNo = PageNo + 1
'''  Print #UBRpt, Now; Tab(27); "Utility Pre-Billing Report"; Tab(70); "Page: "; PageNo
'''  Print #UBRpt, TownName$
'''  Return
'''ErrorAbortExit:
'''  Close
'''
'''ExitPreReport:
'''  UBLog "OUT: Prebilling Report" + CrLf$
'''
'''End Sub
'''
''''


