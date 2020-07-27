Attribute VB_Name = "modUBSorts"
Option Explicit

Public Sub NameQSort(IdxBuff() As DCCustIDXRecType, lLBound, lUBound)
  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
  Dim Temp As DCCustIDXRecType
  Dim Temp2 As DCCustIDXRecType
  'temp.SearchName
  lngCurLow = lLBound
  lngCurHigh = lUBound
  'this is to exit loop if high and low are equal
  'Stop
  If lUBound <= lLBound Then Exit Sub 'GoTo Proc_Exit
    'find the midpoint
    lngCurMid = (lUBound + lLBound) \ 2
    '
    Temp = IdxBuff(lngCurMid)
    Do While (lngCurLow <= lngCurHigh)
      Do While IdxBuff(lngCurLow).IDXName < Temp.IDXName
        lngCurLow = lngCurLow + 1
        If lngCurLow = lUBound Then Exit Do
      Loop
      Do While Temp.IDXName < IdxBuff(lngCurHigh).IDXName
        lngCurHigh = lngCurHigh - 1
        If lngCurHigh = lLBound Then Exit Do
      Loop
      'if low is <= high then swap
      If (lngCurLow <= lngCurHigh) Then
        Temp2 = IdxBuff(lngCurLow)
        IdxBuff(lngCurLow) = IdxBuff(lngCurHigh)
        IdxBuff(lngCurHigh) = Temp2
        lngCurLow = lngCurLow + 1
        lngCurHigh = lngCurHigh - 1
      End If
    Loop
  'recurse if necessary
    If lLBound < lngCurHigh Then
      NameQSort IdxBuff(), lLBound, lngCurHigh
    End If
  'recurse if necessary
    If lngCurLow < lUBound Then
      NameQSort IdxBuff(), lngCurLow, lUBound
    End If
End Sub
'Public Sub AddrQSort(IdxBuff() As UBServiceAddressIndexType, lLBound, lUBound)
'  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
'  Dim Temp As UBServiceAddressIndexType
'  Dim Temp2 As UBServiceAddressIndexType
'  'temp.SearchName
'  lngCurLow = lLBound
'  lngCurHigh = lUBound
'  'this is to exit loop if high and low are equal
'  'Stop
'  If lUBound <= lLBound Then Exit Sub 'GoTo Proc_Exit
'    'find the midpoint
'    lngCurMid = (lUBound + lLBound) \ 2
'    '
'    Temp = IdxBuff(lngCurMid)
'    Do While (lngCurLow <= lngCurHigh)
'      Do While IdxBuff(lngCurLow).ServiceAddress < Temp.ServiceAddress
'        lngCurLow = lngCurLow + 1
'        If lngCurLow = lUBound Then Exit Do
'      Loop
'      Do While Temp.ServiceAddress < IdxBuff(lngCurHigh).ServiceAddress
'        lngCurHigh = lngCurHigh - 1
'        If lngCurHigh = lLBound Then Exit Do
'      Loop
'      'if low is <= high then swap
'      If (lngCurLow <= lngCurHigh) Then
'        Temp2 = IdxBuff(lngCurLow)
'        IdxBuff(lngCurLow) = IdxBuff(lngCurHigh)
'        IdxBuff(lngCurHigh) = Temp2
'        lngCurLow = lngCurLow + 1
'        lngCurHigh = lngCurHigh - 1
'      End If
'    Loop
'  'recurse if necessary
'    If lLBound < lngCurHigh Then
'      AddrQSort IdxBuff(), lLBound, lngCurHigh
'    End If
'  'recurse if necessary
'    If lngCurLow < lUBound Then
'      AddrQSort IdxBuff(), lngCurLow, lUBound
'    End If
'End Sub
'
'Public Sub SortServiceAddrs(formname As Form)
'  Dim CustRecLen As Integer, NumCustRecs As Long, IndexRecLen As Integer
'  Dim CHandle As Integer, cnt As Long, lngCurLow As Long, lngCurHigh As Long
'  Dim IHandle As Integer, IndexName As String, CRec As Long
'  'ShowProcessingScrn "Creating " + IndexText$ + " Index"
' ' QPrintRC "    Reading Customer Records     ", 11, 25, -1
'  ReDim UBCustRec(1) As NewUBCustRecType
'  CustRecLen = Len(UBCustRec(1))
'  NumCustRecs = GetNumOfCust&
'  ReDim ServIndex(1 To NumCustRecs) As UBServiceAddressIndexType
'  IndexRecLen = Len(ServIndex(1))
'  CHandle = FreeFile
'  Open UBPath$ + "UBCUST.DAT" For Random Shared As CHandle Len = CustRecLen
'  For cnt = 1 To NumCustRecs
'    Get CHandle, cnt, UBCustRec(1)
'    ServIndex(cnt).ServiceAddress = UBCustRec(1).ServAddr
'    ServIndex(cnt).RecNum = cnt
'    'ShowPctComp cnt, NumCustRecs                'show user percentage complete
'  Next
'  Close CHandle
'  'QPrintRC "         Sorting Index.        ", 11, 25, -1
'  lngCurLow = LBound(ServIndex)
'  lngCurHigh = UBound(ServIndex)
'  AddrQSort ServIndex(), lngCurLow, lngCurHigh
'  'SortT ServIndex(1), NumCustRecs, 0, 16, 0, 14
'  ' SortT (Elemen, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
' ' QPrintRC "      Writing Index Records      ", 11, 25, -1
'  IndexName$ = TempIndexName
'  KillFile IndexName$
'  IHandle = FreeFile
'  'FCreate IndexName$
'  Open IndexName$ For Random Shared As IHandle Len = 4
'  For cnt = 1 To NumCustRecs
'    CRec& = ServIndex(cnt).RecNum
'    Put IHandle, cnt, CRec&
'    'ShowPctComp cnt, NumCustRecs                'show user percentage complet
'  Next
'  Close IHandle
'
'  Erase UBCustRec, ServIndex
'End Sub


Public Sub SortDCNameIndex(formname As Form)
  Dim DCCustRecLen As Integer, NumOfDCRecs As Long, IdxRecLen As Integer
  Dim DCFile As Integer, cnt As Long, lngCurLow As Long, lngCurHigh As Long
  Dim IHandle As Integer, IndexName As String, CRec As Long, Goodcnt As Long

  ReDim DCCustRec(1) As DCCustRecType
  Dim DCCustIdxRec(1) As DCCustIDXRecType
  FrmShowPctComp.Label1 = "Reindexing Customer Information."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent

  DCCustRecLen = Len(DCCustRec(1))
  DCFile = FreeFile
  Open "DCCUST.DAT" For Random Access Read Write Shared As DCFile Len = DCCustRecLen
  NumOfDCRecs = LOF(DCFile) \ DCCustRecLen
  ReDim SName(1 To NumOfDCRecs) As DCCustIDXRecType
  Goodcnt = 0
  For cnt = 1 To NumOfDCRecs
    Get DCFile, cnt, DCCustRec(1)
    '''If cnt = 10 Then Stop
    If (DCCustRec(1).Deleted <> "Y") And (Len(QPTrim$(DCCustRec(1).CUSTNUMB)) > 0) Then
      Goodcnt = Goodcnt + 1
      'If QPTrim$(DCCustRec(1).BILLNAME) <> "" Then Stop
      SName(Goodcnt).IDXName = QPTrim$(DCCustRec(1).SORTNAME)
      SName(Goodcnt).IDXRECORD = cnt
    End If
    FrmShowPctComp.ShowPctComp cnt, NumOfDCRecs
  Next cnt
  Close DCFile
  FrmShowPctComp.Label1 = "Reindexing Customer Information."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent
 
  lngCurLow = LBound(SName)
  lngCurHigh = UBound(SName)
  NameQSort SName(), 1, Goodcnt

  'SortT Array(1), GoodRecords, Dir, SSize, MOff, MSize
  IndexName$ = "DCCUST.IDX"
  KillFile IndexName$

  IHandle = FreeFile
  Open IndexName$ For Random Shared As IHandle Len = 4
  For cnt = 1 To Goodcnt
    CRec& = SName(cnt).IDXRECORD
    Put IHandle, cnt, CRec&
    FrmShowPctComp.ShowPctComp cnt, Goodcnt               'show user percentage complet
  Next
  Close IHandle

  Erase DCCustRec, SName

End Sub

''!!! For Meter Installed Date Report
'Public Sub MtrDQSort(IdxBuff() As MtrDateSortType, lLBound, lUBound)
'  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
'  Dim Temp As MtrDateSortType
'  Dim Temp2 As MtrDateSortType
'  'temp.SearchName
'  lngCurLow = lLBound
'  lngCurHigh = lUBound
'  'this is to exit loop if high and low are equal
'  'Stop
'  If lUBound <= lLBound Then Exit Sub 'GoTo Proc_Exit
'    'find the midpoint
'    lngCurMid = (lUBound + lLBound) \ 2
'    '
'    Temp = IdxBuff(lngCurMid)
'    Do While (lngCurLow <= lngCurHigh)
'      Do While IdxBuff(lngCurLow).MtrDate < Temp.MtrDate
'        lngCurLow = lngCurLow + 1
'        If lngCurLow = lUBound Then Exit Do
'      Loop
'      Do While Temp.MtrDate < IdxBuff(lngCurHigh).MtrDate
'        lngCurHigh = lngCurHigh - 1
'        If lngCurHigh = lLBound Then Exit Do
'      Loop
'      'if low is <= high then swap
'      If (lngCurLow <= lngCurHigh) Then
'        Temp2 = IdxBuff(lngCurLow)
'        IdxBuff(lngCurLow) = IdxBuff(lngCurHigh)
'        IdxBuff(lngCurHigh) = Temp2
'        lngCurLow = lngCurLow + 1
'        lngCurHigh = lngCurHigh - 1
'      End If
'    Loop
'  'recurse if necessary
'    If lLBound < lngCurHigh Then
'      MtrDQSort IdxBuff(), lLBound, lngCurHigh
'    End If
'  'recurse if necessary
'    If lngCurLow < lUBound Then
'      MtrDQSort IdxBuff(), lngCurLow, lUBound
'    End If
'End Sub
'
'Public Sub SortMtrDateIndex(FromDate%, ThruDate%)
'  Dim CustRecLen As Integer, NumCustRecs As Long, IndexRecLen As Integer
'  Dim CHandle As Integer, cnt As Long, lngCurLow As Long, lngCurHigh As Long
'  Dim IHandle As Integer, IndexName As String, CRec As Long, MTCnt As Integer
'  Dim MtrDate As Integer, FoundCnt As Integer
''  ShowProcessingScrn "Creating Meter Date Index"
''  QPrintRC "    Reading Customer Records     ", 11, 25, -1
'
'  ReDim UBCustRec(1) As NewUBCustRecType
'  CustRecLen = Len(UBCustRec(1))
'
'  NumCustRecs = GetNumOfCust&
'
'  ReDim MDateIdx(1 To 1) As MtrDateSortType
'  IndexRecLen = Len(MDateIdx(1))
'
'  CHandle = FreeFile
'  Open UBPath$ + "UBCUST.DAT" For Random Shared As CHandle Len = CustRecLen
'  For cnt = 1 To NumCustRecs
'    Get CHandle, cnt, UBCustRec(1)
'    If UBCustRec(1).DelFlag <> -1 Then
'      For MTCnt = 1 To 7
'        MtrDate = UBCustRec(1).LocMeters(MTCnt).InsDate
'        If MtrDate >= FromDate And MtrDate <= ThruDate Then
'          FoundCnt = FoundCnt + 1
'          ReDim Preserve MDateIdx(1 To FoundCnt) As MtrDateSortType
'          MDateIdx(FoundCnt).MtrDate = UBCustRec(1).LocMeters(MTCnt).InsDate
'          MDateIdx(FoundCnt).RecNum = cnt
'          Exit For
'        End If
'      Next
'    End If
'    'ShowPctComp cnt, NumCustRecs                'show user percentage complete
'  Next
'
'  Close CHandle
'
'  'QPrintRC "         Sorting Index.        ", 11, 25, -1
'  lngCurLow = LBound(MDateIdx)
'  lngCurHigh = UBound(MDateIdx)
'  MtrDQSort MDateIdx(), lngCurLow, lngCurHigh
''  SortT MDateIdx(1), FoundCnt, 0, 4, 0, -1
''  'Elemen, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
''  QPrintRC "      Writing Index Records      ", 11, 25, -1
'  IndexName$ = TempIndexName
'  KillFile IndexName$
'  IHandle = FreeFile
'    'FCreate IndexName$
'  Open IndexName$ For Random Shared As IHandle Len = 4
'  For cnt = 1 To FoundCnt
'    CRec& = MDateIdx(cnt).RecNum
'    Put IHandle, cnt, CRec&
'    'ShowPctComp cnt, NumCustRecs                'show user percentage complete
'  Next
'  Close IHandle
'
'  Erase UBCustRec, MDateIdx
'
'End Sub
Public Sub MakeZipCodeIndex(IndexText$)
  FrmShowPctComp.Label1 = "Reading Customer Information."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent
  Dim DCCustIdxRec(1) As ZipIndexType
  ReDim DCCustRec(1) As DCCustRecType
  Dim CustRecLen As Integer, IndexRecLen As Integer, IdxRecLen As Integer
  Dim CHandle As Integer, cnt As Long, IHandle As Integer
  Dim NumCustRecs As Long, Prec As Long, NumOfBillRec As Long
  Dim Bcnt As Long, Goodcnt As Long, IndexName As String
  ReDim DCCustRec(1) As DCCustRecType
  CustRecLen = Len(DCCustRec(1))

  NumOfBillRec = FileSize("DCCUST.DAT") \ CustRecLen

  CHandle = FreeFile
  Open DCPath$ + "DCCUST.DAT" For Random Shared As CHandle Len = CustRecLen
  

  ReDim ZipIndex(1 To NumOfBillRec) As ZipIndexType
  Goodcnt = 0
  For Bcnt = 1 To NumOfBillRec
    Get CHandle, Bcnt, DCCustRec(1)
    If (DCCustRec(1).Deleted <> "Y") And (Len(QPTrim$(DCCustRec(1).CUSTNUMB)) > 0) Then
      Goodcnt = Goodcnt + 1
      ZipIndex(Goodcnt).IDXName = DCCustRec(1).ZIPCODE
      ZipIndex(Goodcnt).IDXRECORD = Bcnt
      FrmShowPctComp.ShowPctComp Bcnt, NumOfBillRec              'show user percentage complete
    End If
  Next
  Close
  Load frmInfo
  frmInfo.Label1 = "Sorting. . ."
  DoEvents
  frmInfo.Show
  DoEvents
  ZipQSort ZipIndex(), 1, Goodcnt
  Unload frmInfo
  DoEvents
  
  FrmShowPctComp.Label1 = "Writing Index Records."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show

  IndexName$ = "DCTemp.idx"
  KillFile IndexName$
  IHandle = FreeFile
  Open IndexName$ For Random Shared As IHandle Len = 4
  For cnt = 1 To Goodcnt
    Prec& = ZipIndex(cnt).IDXRECORD
    Put IHandle, cnt, Prec&
    FrmShowPctComp.ShowPctComp cnt, Goodcnt           'show user percentage complete
  Next
  Close IHandle
  Erase DCCustRec, ZipIndex

End Sub

'!!!For Labels
Public Sub ZipQSort(IdxBuff() As ZipIndexType, lLBound, lUBound)
  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
  Dim Temp As ZipIndexType
  Dim Temp2 As ZipIndexType
  'temp.SearchName
  lngCurLow = lLBound
  lngCurHigh = lUBound
  'this is to exit loop if high and low are equal
  'Stop
  If lUBound <= lLBound Then Exit Sub 'GoTo Proc_Exit
    'find the midpoint
    lngCurMid = (lUBound + lLBound) \ 2
    '
    Temp = IdxBuff(lngCurMid)
    Do While (lngCurLow <= lngCurHigh)
      Do While IdxBuff(lngCurLow).IDXName < Temp.IDXName
        lngCurLow = lngCurLow + 1
        If lngCurLow = lUBound Then Exit Do
      Loop
      Do While Temp.IDXName < IdxBuff(lngCurHigh).IDXName
        lngCurHigh = lngCurHigh - 1
        If lngCurHigh = lLBound Then Exit Do
      Loop
      'if low is <= high then swap
      If (lngCurLow <= lngCurHigh) Then
        Temp2 = IdxBuff(lngCurLow)
        IdxBuff(lngCurLow) = IdxBuff(lngCurHigh)
        IdxBuff(lngCurHigh) = Temp2
        lngCurLow = lngCurLow + 1
        lngCurHigh = lngCurHigh - 1
      End If
    Loop
  'recurse if necessary
    If lLBound < lngCurHigh Then
      ZipQSort IdxBuff(), lLBound, lngCurHigh
    End If
  'recurse if necessary
    If lngCurLow < lUBound Then
      ZipQSort IdxBuff(), lngCurLow, lUBound
    End If
End Sub
''!!! For Labels
'Public Sub PostalQSort(IdxBuff() As UBPostalIndexType, lLBound, lUBound)
'  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
'  Dim Temp As UBPostalIndexType
'  Dim Temp2 As UBPostalIndexType
'  'temp.SearchName
'  lngCurLow = lLBound
'  lngCurHigh = lUBound
'  'this is to exit loop if high and low are equal
'  'Stop
'  If lUBound <= lLBound Then Exit Sub 'GoTo Proc_Exit
'    'find the midpoint
'    lngCurMid = (lUBound + lLBound) \ 2
'    '
'    Temp = IdxBuff(lngCurMid)
'    Do While (lngCurLow <= lngCurHigh)
'      Do While Val(IdxBuff(lngCurLow).Route) < Val(Temp.Route)
'        lngCurLow = lngCurLow + 1
'        If lngCurLow = lUBound Then Exit Do
'      Loop
'      Do While Val(Temp.Route) < Val(IdxBuff(lngCurHigh).Route)
'        lngCurHigh = lngCurHigh - 1
'        If lngCurHigh = lLBound Then Exit Do
'      Loop
'      'if low is <= high then swap
'      If (lngCurLow <= lngCurHigh) Then
'        Temp2 = IdxBuff(lngCurLow)
'        IdxBuff(lngCurLow) = IdxBuff(lngCurHigh)
'        IdxBuff(lngCurHigh) = Temp2
'        lngCurLow = lngCurLow + 1
'        lngCurHigh = lngCurHigh - 1
'      End If
'    Loop
'  'recurse if necessary
'    If lLBound < lngCurHigh Then
'      PostalQSort IdxBuff(), lLBound, lngCurHigh
'    End If
'  'recurse if necessary
'    If lngCurLow < lUBound Then
'      PostalQSort IdxBuff(), lngCurLow, lUBound
'    End If
'End Sub
'Public Sub PostZipQSort(IdxBuff() As UBPostalIndexType, lLBound, lUBound)
'  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
'  Dim Temp As UBPostalIndexType
'  Dim Temp2 As UBPostalIndexType
'  'temp.SearchName
'  lngCurLow = lLBound
'  lngCurHigh = lUBound
'  'this is to exit loop if high and low are equal
'  'Stop
'  If lUBound <= lLBound Then Exit Sub 'GoTo Proc_Exit
'    'find the midpoint
'    lngCurMid = (lUBound + lLBound) \ 2
'    '
'    Temp = IdxBuff(lngCurMid)
'    Do While (lngCurLow <= lngCurHigh)
'      Do While IdxBuff(lngCurLow).ZIPCODE < Temp.ZIPCODE
'        lngCurLow = lngCurLow + 1
'        If lngCurLow = lUBound Then Exit Do
'      Loop
'      Do While Temp.ZIPCODE < IdxBuff(lngCurHigh).ZIPCODE
'        lngCurHigh = lngCurHigh - 1
'        If lngCurHigh = lLBound Then Exit Do
'      Loop
'      'if low is <= high then swap
'      If (lngCurLow <= lngCurHigh) Then
'        Temp2 = IdxBuff(lngCurLow)
'        IdxBuff(lngCurLow) = IdxBuff(lngCurHigh)
'        IdxBuff(lngCurHigh) = Temp2
'        lngCurLow = lngCurLow + 1
'        lngCurHigh = lngCurHigh - 1
'      End If
'    Loop
'  'recurse if necessary
'    If lLBound < lngCurHigh Then
'      PostZipQSort IdxBuff(), lLBound, lngCurHigh
'    End If
'  'recurse if necessary
'    If lngCurLow < lUBound Then
'      PostZipQSort IdxBuff(), lngCurLow, lUBound
'    End If
'End Sub
'Public Sub ZipLocSort(IdxBuff() As UBZipLocationIndexType, lLBound, lUBound)
'  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
'  Dim Temp As UBZipLocationIndexType
'  Dim Temp2 As UBZipLocationIndexType
'  'temp.SearchName
'  lngCurLow = lLBound
'  lngCurHigh = lUBound
'  'this is to exit loop if high and low are equal
'  'Stop
'  If lUBound <= lLBound Then Exit Sub 'GoTo Proc_Exit
'    'find the midpoint
'    lngCurMid = (lUBound + lLBound) \ 2
'    '
'    Temp = IdxBuff(lngCurMid)
'    Do While (lngCurLow <= lngCurHigh)
'      Do While IdxBuff(lngCurLow).ZIPLocat < Temp.ZIPLocat
'        lngCurLow = lngCurLow + 1
'        If lngCurLow = lUBound Then Exit Do
'      Loop
'      Do While Temp.ZIPLocat < IdxBuff(lngCurHigh).ZIPLocat
'        lngCurHigh = lngCurHigh - 1
'        If lngCurHigh = lLBound Then Exit Do
'      Loop
'      'if low is <= high then swap
'      If (lngCurLow <= lngCurHigh) Then
'        Temp2 = IdxBuff(lngCurLow)
'        IdxBuff(lngCurLow) = IdxBuff(lngCurHigh)
'        IdxBuff(lngCurHigh) = Temp2
'        lngCurLow = lngCurLow + 1
'        lngCurHigh = lngCurHigh - 1
'      End If
'    Loop
'  'recurse if necessary
'    If lLBound < lngCurHigh Then
'      ZipLocSort IdxBuff(), lLBound, lngCurHigh
'    End If
'  'recurse if necessary
'    If lngCurLow < lUBound Then
'      ZipLocSort IdxBuff(), lngCurLow, lUBound
'    End If
'End Sub
'
'Public Sub RateQSort(IdxBuff() As UBRateTblRecType, lLBound, lUBound)
'  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
'  Dim Temp As UBRateTblRecType
'  Dim Temp2 As UBRateTblRecType
'  'temp.SearchName
'  lngCurLow = lLBound
'  lngCurHigh = lUBound
'  'this is to exit loop if high and low are equal
'  'Stop
'  If lUBound <= lLBound Then Exit Sub 'GoTo Proc_Exit
'    'find the midpoint
'    lngCurMid = (lUBound + lLBound) \ 2
'    '
'    Temp = IdxBuff(lngCurMid)
'    Do While (lngCurLow <= lngCurHigh)
'      Do While IdxBuff(lngCurLow).Ratecode < Temp.Ratecode
'        lngCurLow = lngCurLow + 1
'        If lngCurLow = lUBound Then Exit Do
'      Loop
'      Do While Temp.Ratecode < IdxBuff(lngCurHigh).Ratecode
'        lngCurHigh = lngCurHigh - 1
'        If lngCurHigh = lLBound Then Exit Do
'      Loop
'      'if low is <= high then swap
'      If (lngCurLow <= lngCurHigh) Then
'        Temp2 = IdxBuff(lngCurLow)
'        IdxBuff(lngCurLow) = IdxBuff(lngCurHigh)
'        IdxBuff(lngCurHigh) = Temp2
'        lngCurLow = lngCurLow + 1
'        lngCurHigh = lngCurHigh - 1
'      End If
'    Loop
'  'recurse if necessary
'    If lLBound < lngCurHigh Then
'      RateQSort IdxBuff(), lLBound, lngCurHigh
'    End If
'  'recurse if necessary
'    If lngCurLow < lUBound Then
'      RateQSort IdxBuff(), lngCurLow, lUBound
'    End If
'End Sub
'Public Sub BookCQSort(IdxBuff() As BookConsumpType, lLBound, lUBound)
'  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
'  Dim Temp As BookConsumpType
'  Dim Temp2 As BookConsumpType
'  'temp.SearchName
'  lngCurLow = lLBound
'  lngCurHigh = lUBound
'  'this is to exit loop if high and low are equal
'  'Stop
'  If lUBound <= lLBound Then Exit Sub 'GoTo Proc_Exit
'    'find the midpoint
'    lngCurMid = (lUBound + lLBound) \ 2
'    '
'    Temp = IdxBuff(lngCurMid)
'    Do While (lngCurLow <= lngCurHigh)
'      Do While IdxBuff(lngCurLow).Book < Temp.Book
'        lngCurLow = lngCurLow + 1
'        If lngCurLow = lUBound Then Exit Do
'      Loop
'      Do While Temp.Book < IdxBuff(lngCurHigh).Book
'        lngCurHigh = lngCurHigh - 1
'        If lngCurHigh = lLBound Then Exit Do
'      Loop
'      'if low is <= high then swap
'      If (lngCurLow <= lngCurHigh) Then
'        Temp2 = IdxBuff(lngCurLow)
'        IdxBuff(lngCurLow) = IdxBuff(lngCurHigh)
'        IdxBuff(lngCurHigh) = Temp2
'        lngCurLow = lngCurLow + 1
'        lngCurHigh = lngCurHigh - 1
'      End If
'    Loop
'  'recurse if necessary
'    If lLBound < lngCurHigh Then
'      BookCQSort IdxBuff(), lLBound, lngCurHigh
'    End If
'  'recurse if necessary
'    If lngCurLow < lUBound Then
'      BookCQSort IdxBuff(), lngCurLow, lUBound
'    End If
'End Sub
'
'Public Sub BillQSort(IdxBuff() As RePrintIndexType, lLBound, lUBound)
'  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
'  Dim Temp As RePrintIndexType
'  Dim Temp2 As RePrintIndexType
'  lngCurLow = lLBound
'  lngCurHigh = lUBound
'  'this is to exit loop if high and low are equal
'  'Stop
'  If lUBound <= lLBound Then Exit Sub 'GoTo Proc_Exit
'    'find the midpoint
'    lngCurMid = (lUBound + lLBound) \ 2
'    '
'    Temp = IdxBuff(lngCurMid)
'    Do While (lngCurLow <= lngCurHigh)
'      Do While IdxBuff(lngCurLow).BillNum < Temp.BillNum
'        lngCurLow = lngCurLow + 1
'        If lngCurLow = lUBound Then Exit Do
'      Loop
'      Do While Temp.BillNum < IdxBuff(lngCurHigh).BillNum
'        lngCurHigh = lngCurHigh - 1
'        If lngCurHigh = lLBound Then Exit Do
'      Loop
'      'if low is <= high then swap
'      If (lngCurLow <= lngCurHigh) Then
'        Temp2 = IdxBuff(lngCurLow)
'        IdxBuff(lngCurLow) = IdxBuff(lngCurHigh)
'        IdxBuff(lngCurHigh) = Temp2
'        lngCurLow = lngCurLow + 1
'        lngCurHigh = lngCurHigh - 1
'      End If
'    Loop
'  'recurse if necessary
'    If lLBound < lngCurHigh Then
'      BillQSort IdxBuff(), lLBound, lngCurHigh
'    End If
'  'recurse if necessary
'    If lngCurLow < lUBound Then
'      BillQSort IdxBuff(), lngCurLow, lUBound
'    End If
'End Sub
'
'Public Sub BDSort(IdxBuff() As BDRptType, lLBound, lUBound)
'  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
'  Dim Temp As BDRptType
'  Dim Temp2 As BDRptType
'  lngCurLow = lLBound
'  lngCurHigh = lUBound
'  'this is to exit loop if high and low are equal
'  'Stop
'  If lUBound <= lLBound Then Exit Sub 'GoTo Proc_Exit
'    'find the midpoint
'    lngCurMid = (lUBound + lLBound) \ 2
'    '
'    Temp = IdxBuff(lngCurMid)
'    Do While (lngCurLow <= lngCurHigh)
'      Do While IdxBuff(lngCurLow).BankName < Temp.BankName
'        lngCurLow = lngCurLow + 1
'        If lngCurLow = lUBound Then Exit Do
'      Loop
'      Do While Temp.BankName < IdxBuff(lngCurHigh).BankName
'        lngCurHigh = lngCurHigh - 1
'        If lngCurHigh = lLBound Then Exit Do
'      Loop
'      'if low is <= high then swap
'      If (lngCurLow <= lngCurHigh) Then
'        Temp2 = IdxBuff(lngCurLow)
'        IdxBuff(lngCurLow) = IdxBuff(lngCurHigh)
'        IdxBuff(lngCurHigh) = Temp2
'        lngCurLow = lngCurLow + 1
'        lngCurHigh = lngCurHigh - 1
'      End If
'    Loop
'  'recurse if necessary
'    If lLBound < lngCurHigh Then
'      BDSort IdxBuff(), lLBound, lngCurHigh
'    End If
'  'recurse if necessary
'    If lngCurLow < lUBound Then
'      BDSort IdxBuff(), lngCurLow, lUBound
'    End If
'End Sub
'
'Public Sub MakeZipLocationIndex(IndexText$)
'  FrmShowPctComp.Label1 = "Reading Customer Information."
'  FrmShowPctComp.cmdCancel.Enabled = False
'  FrmShowPctComp.Show '1, Parent
'
'  ReDim UBCustRec(1) As NewUBCustRecType
'  Dim CustRecLen As Integer, IndexRecLen As Integer, Zp As String
'  Dim CHandle As Integer, cnt As Long, IHandle As Integer, zp4 As String
'  Dim NumCustRecs As Long, Prec As Long, NumOfRec As Long
'  Dim Bcnt As Long
'  ReDim UBCustRec(1) As NewUBCustRecType
'  CustRecLen = Len(UBCustRec(1))
'
'  NumOfRec = FileSize("UBCUST.DAT") \ CustRecLen
'
'  CHandle = FreeFile
'  Open UBPath$ + "UBCUST.DAT" For Random Shared As CHandle Len = CustRecLen
'
'  ReDim ZipIndex(1 To NumOfRec) As UBZipLocationIndexType
'  For Bcnt = 1 To NumOfRec
'    Get CHandle, Bcnt, UBCustRec(1)
'    zp4$ = Mid$(UBCustRec(1).ZIPCODE, 7, 4)
'    Zp$ = Left$(UBCustRec(1).ZIPCODE, 5)
'    If Len(QPTrim$(zp4$)) = 0 Then
'      Zp$ = Zp$ + "-0000"
'    Else
'      Zp$ = UBCustRec(1).ZIPCODE
'    End If
'    'If Len(QPTrim$(UBCustRec(1).Book)) > 0 And L(UBCustRec(1).SEQNUMB) > 0 Then
'      ZipIndex(Bcnt).ZIPLocat = Zp$ & UBCustRec(1).Book & UBCustRec(1).SEQNUMB
'    'Else
'   '   ZipIndex(Bcnt).ZIPLocat = Zp$ & "00000000"
'   ' End If
'    ZipIndex(Bcnt).RecNum = Bcnt
'    FrmShowPctComp.ShowPctComp Bcnt, NumOfRec              'show user percentage complete
'  Next
'  Close
'  Load frmInfo
'  frmInfo.Label1 = "Sorting. . ."
'  DoEvents
'  frmInfo.Show
'  DoEvents
'  ZipLocSort ZipIndex(), 1, NumOfRec
'  Unload frmInfo
'  DoEvents
'
'  FrmShowPctComp.Label1 = "Writing Index Records."
'  FrmShowPctComp.cmdCancel.Enabled = False
'  FrmShowPctComp.Show
'
' KillFile TempIndexName
'  IHandle = FreeFile
'  Open TempIndexName For Output As IHandle
'  Close IHandle
'
'  IHandle = FreeFile
'  Open TempIndexName For Random Shared As IHandle Len = 4
'  For cnt = 1 To NumOfRec
'    Prec& = ZipIndex(cnt).RecNum
'    Put IHandle, cnt, Prec&
'    FrmShowPctComp.ShowPctComp cnt, NumOfRec               'show user percentage complete
'  Next
'  Close IHandle
'
'  Erase UBCustRec, ZipIndex
'
'End Sub
'
