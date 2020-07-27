Attribute VB_Name = "modUBSorts"
Option Explicit

Public Sub NameQSort(IdxBuff() As nUBCustReIndexRecType, lLBound, lUBound)
  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
  Dim Temp As nUBCustReIndexRecType
  Dim Temp2 As nUBCustReIndexRecType
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
      Do While IdxBuff(lngCurLow).SearchName < Temp.SearchName
        lngCurLow = lngCurLow + 1
        If lngCurLow = lUBound Then Exit Do
      Loop
      Do While Temp.SearchName < IdxBuff(lngCurHigh).SearchName
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

Public Sub SeqQSort(IdxBuff() As UBSequenceIndexType, lLBound, lUBound)
  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
  Dim Temp As UBSequenceIndexType
  Dim Temp2 As UBSequenceIndexType
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
      Do While IdxBuff(lngCurLow).SeqNumber < Temp.SeqNumber
        lngCurLow = lngCurLow + 1
        If lngCurLow = lUBound Then Exit Do
      Loop
      Do While Temp.SeqNumber < IdxBuff(lngCurHigh).SeqNumber
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
      SeqQSort IdxBuff(), lLBound, lngCurHigh
    End If
  'recurse if necessary
    If lngCurLow < lUBound Then
      SeqQSort IdxBuff(), lngCurLow, lUBound
    End If
End Sub

Public Sub LocQSort(IdxBuff() As UBLocaReIndexRecTypeVB, lLBound, lUBound)
  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
  Dim Temp As UBLocaReIndexRecTypeVB
  Dim Temp2 As UBLocaReIndexRecTypeVB
  lngCurLow = lLBound
  lngCurHigh = lUBound
  'this is to exit loop if high and low are equal
  'Stop
  If lUBound <= lLBound Then Exit Sub 'GoTo Proc_Exit
    'find the midpoint
    lngCurMid = (lUBound + lLBound) \ 2
    Temp = IdxBuff(lngCurMid)
    Do While (lngCurLow <= lngCurHigh)
      Do While IdxBuff(lngCurLow).BookSEQNUMB < Temp.BookSEQNUMB
        lngCurLow = lngCurLow + 1
        If lngCurLow = lUBound Then Exit Do
      Loop
      Do While Temp.BookSEQNUMB < IdxBuff(lngCurHigh).BookSEQNUMB
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
      LocQSort IdxBuff(), lLBound, lngCurHigh
    End If
  'recurse if necessary
    If lngCurLow < lUBound Then
      LocQSort IdxBuff(), lngCurLow, lUBound
    End If

End Sub
Public Sub AddrQSort(IdxBuff() As UBServiceAddressIndexType, lLBound, lUBound)
  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
  Dim Temp As UBServiceAddressIndexType
  Dim Temp2 As UBServiceAddressIndexType
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
      Do While IdxBuff(lngCurLow).ServiceAddress < Temp.ServiceAddress
        lngCurLow = lngCurLow + 1
        If lngCurLow = lUBound Then Exit Do
      Loop
      Do While Temp.ServiceAddress < IdxBuff(lngCurHigh).ServiceAddress
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
      AddrQSort IdxBuff(), lLBound, lngCurHigh
    End If
  'recurse if necessary
    If lngCurLow < lUBound Then
      AddrQSort IdxBuff(), lngCurLow, lUBound
    End If
End Sub

Public Sub SortServiceAddrs(formname As Form)
  Dim CustRecLen As Integer, NumCustRecs As Long, IndexRecLen As Integer
  Dim CHandle As Integer, cnt As Long, lngCurLow As Long, lngCurHigh As Long
  Dim IHandle As Integer, IndexName As String, CRec As Long
  'ShowProcessingScrn "Creating " + IndexText$ + " Index"
 ' QPrintRC "    Reading Customer Records     ", 11, 25, -1

  ReDim UBCustRec(1) As NewUBCustRecType
  CustRecLen = Len(UBCustRec(1))

  NumCustRecs = GetNumOfCust&

  ReDim ServIndex(1 To NumCustRecs) As UBServiceAddressIndexType
  IndexRecLen = Len(ServIndex(1))

  CHandle = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CHandle Len = CustRecLen
  For cnt = 1 To NumCustRecs
    Get CHandle, cnt, UBCustRec(1)
    ServIndex(cnt).ServiceAddress = UBCustRec(1).SERVADDR
    ServIndex(cnt).RecNum = cnt
    'ShowPctComp cnt, NumCustRecs                'show user percentage complete
  Next

  Close CHandle

  'QPrintRC "         Sorting Index.        ", 11, 25, -1
  lngCurLow = LBound(ServIndex)
  lngCurHigh = UBound(ServIndex)
  AddrQSort ServIndex(), lngCurLow, lngCurHigh
  'SortT ServIndex(1), NumCustRecs, 0, 16, 0, 14
  ' SortT (Elemen, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
 ' QPrintRC "      Writing Index Records      ", 11, 25, -1
  IndexName$ = TempIndexName
  KillFile IndexName$
  IHandle = FreeFile
  'FCreate IndexName$
  Open IndexName$ For Random Shared As IHandle Len = 4
  For cnt = 1 To NumCustRecs
    CRec& = ServIndex(cnt).RecNum
    Put IHandle, cnt, CRec&
    'ShowPctComp cnt, NumCustRecs                'show user percentage complet
  Next
  Close IHandle

  Erase UBCustRec, ServIndex
End Sub
'!!! For Meter Installed Date Report
Public Sub MtrDQSort(IdxBuff() As MtrDateSortType, lLBound, lUBound)
  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
  Dim Temp As MtrDateSortType
  Dim Temp2 As MtrDateSortType
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
      Do While IdxBuff(lngCurLow).MtrDate < Temp.MtrDate
        lngCurLow = lngCurLow + 1
        If lngCurLow = lUBound Then Exit Do
      Loop
      Do While Temp.MtrDate < IdxBuff(lngCurHigh).MtrDate
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
      MtrDQSort IdxBuff(), lLBound, lngCurHigh
    End If
  'recurse if necessary
    If lngCurLow < lUBound Then
      MtrDQSort IdxBuff(), lngCurLow, lUBound
    End If
End Sub

Public Sub SortMtrDateIndex(FromDate%, ThruDate%)
  Dim CustRecLen As Integer, NumCustRecs As Long, IndexRecLen As Integer
  Dim CHandle As Integer, cnt As Long, lngCurLow As Long, lngCurHigh As Long
  Dim IHandle As Integer, IndexName As String, CRec As Long, MTCnt As Integer
  Dim MtrDate As Integer, FoundCnt As Integer
'  ShowProcessingScrn "Creating Meter Date Index"
'  QPrintRC "    Reading Customer Records     ", 11, 25, -1

  ReDim UBCustRec(1) As NewUBCustRecType
  CustRecLen = Len(UBCustRec(1))

  NumCustRecs = GetNumOfCust&

  ReDim MDateIdx(1 To 1) As MtrDateSortType
  IndexRecLen = Len(MDateIdx(1))

  CHandle = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CHandle Len = CustRecLen
  For cnt = 1 To NumCustRecs
    Get CHandle, cnt, UBCustRec(1)
    If UBCustRec(1).DelFlag <> -1 Then
      For MTCnt = 1 To 7
        MtrDate = UBCustRec(1).LocMeters(MTCnt).InsDate
        If MtrDate >= FromDate And MtrDate <= ThruDate Then
          FoundCnt = FoundCnt + 1
          ReDim Preserve MDateIdx(1 To FoundCnt) As MtrDateSortType
          MDateIdx(FoundCnt).MtrDate = UBCustRec(1).LocMeters(MTCnt).InsDate
          MDateIdx(FoundCnt).RecNum = cnt
          Exit For
        End If
      Next
    End If
    'ShowPctComp cnt, NumCustRecs                'show user percentage complete
  Next

  Close CHandle

  'QPrintRC "         Sorting Index.        ", 11, 25, -1
  lngCurLow = LBound(MDateIdx)
  lngCurHigh = UBound(MDateIdx)
  MtrDQSort MDateIdx(), lngCurLow, lngCurHigh
'  SortT MDateIdx(1), FoundCnt, 0, 4, 0, -1
'  'Elemen, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
'  QPrintRC "      Writing Index Records      ", 11, 25, -1
  IndexName$ = TempIndexName
  KillFile IndexName$
  IHandle = FreeFile
    'FCreate IndexName$
  Open IndexName$ For Random Shared As IHandle Len = 4
  For cnt = 1 To FoundCnt
    CRec& = MDateIdx(cnt).RecNum
    Put IHandle, cnt, CRec&
    'ShowPctComp cnt, NumCustRecs                'show user percentage complete
  Next
  Close IHandle

  Erase UBCustRec, MDateIdx

End Sub
'!!!For Labels
Public Sub ZipQSort(IdxBuff() As MOWZipIndexType, lLBound, lUBound)
  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
  Dim Temp As MOWZipIndexType
  Dim Temp2 As MOWZipIndexType
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
      Do While IdxBuff(lngCurLow).ZIPCODE < Temp.ZIPCODE
        lngCurLow = lngCurLow + 1
        If lngCurLow = lUBound Then Exit Do
      Loop
      Do While Temp.ZIPCODE < IdxBuff(lngCurHigh).ZIPCODE
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
'!!! For Labels
Public Sub PostalQSort(IdxBuff() As UBPostalIndexType, lLBound, lUBound)
  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
  Dim Temp As UBPostalIndexType
  Dim Temp2 As UBPostalIndexType
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
      Do While IdxBuff(lngCurLow).Route < Temp.Route
        lngCurLow = lngCurLow + 1
        If lngCurLow = lUBound Then Exit Do
      Loop
      Do While Temp.Route < IdxBuff(lngCurHigh).Route
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
      PostalQSort IdxBuff(), lLBound, lngCurHigh
    End If
  'recurse if necessary
    If lngCurLow < lUBound Then
      PostalQSort IdxBuff(), lngCurLow, lUBound
    End If
End Sub
Public Sub RateQSort(IdxBuff() As UBRateTblRecType, lLBound, lUBound)
  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
  Dim Temp As UBRateTblRecType
  Dim Temp2 As UBRateTblRecType
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
      Do While IdxBuff(lngCurLow).RATECODE < Temp.RATECODE
        lngCurLow = lngCurLow + 1
        If lngCurLow = lUBound Then Exit Do
      Loop
      Do While Temp.RATECODE < IdxBuff(lngCurHigh).RATECODE
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
      RateQSort IdxBuff(), lLBound, lngCurHigh
    End If
  'recurse if necessary
    If lngCurLow < lUBound Then
      RateQSort IdxBuff(), lngCurLow, lUBound
    End If
End Sub
Public Sub BookCQSort(IdxBuff() As BookConsumpType, lLBound, lUBound)
  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
  Dim Temp As BookConsumpType
  Dim Temp2 As BookConsumpType
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
      Do While IdxBuff(lngCurLow).Book < Temp.Book
        lngCurLow = lngCurLow + 1
        If lngCurLow = lUBound Then Exit Do
      Loop
      Do While Temp.Book < IdxBuff(lngCurHigh).Book
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
      BookCQSort IdxBuff(), lLBound, lngCurHigh
    End If
  'recurse if necessary
    If lngCurLow < lUBound Then
      BookCQSort IdxBuff(), lngCurLow, lUBound
    End If
End Sub

