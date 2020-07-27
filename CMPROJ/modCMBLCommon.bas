Attribute VB_Name = "modCMBLCommon"
Option Explicit
'****************BL Stuff*********************************
Public Function EmpInLicProcess(EmpNum$) As Boolean
  Dim x As Long
  Dim TempRec As TempTransPostType
  Dim TempHandle As Integer
  Dim NumOfTempRecs As Long
  
  EmpInLicProcess = False
  OpenTempPostFile TempHandle
  NumOfTempRecs = LOF(TempHandle) / Len(TempRec)
  
  If NumOfTempRecs = 0 Then
    Close TempHandle
    Exit Function
  End If
  
  For x = 1 To NumOfTempRecs
    Get TempHandle, x, TempRec
    If QPTrim$(EmpNum$) = QPTrim$(TempRec.CustomerNumber) Then
      EmpInLicProcess = True
      Exit For
    End If
  Next x
  Close TempHandle
  
End Function
Public Sub OpenTempPostFile(TempPostHandle As Integer)
  Dim TempPostRec As TempTransPostType
  Dim TempPostLen As Integer
  TempPostLen = Len(TempPostRec)
  TempPostHandle = FreeFile
  Open BLTransTempPost For Random Shared As TempPostHandle Len = TempPostLen
End Sub
Public Sub OpenBLCustFile(CustHandle As Integer)
  Dim CustRec As ARCustRecType
  Dim CustLen As Integer
  CustLen = Len(CustRec)
  CustHandle = FreeFile
  Open BLCustFileName For Random Shared As CustHandle Len = CustLen
End Sub
Public Sub OpenPenTransFile(PenTransHandle As Integer)
  Dim PenTransRec As TempPenaltyCharges
  Dim PenTransLen As Integer
  PenTransLen = Len(PenTransRec)
  PenTransHandle = FreeFile
  Open BLTempPenaltyCharges For Random Shared As PenTransHandle Len = PenTransLen
End Sub
Public Sub OpenBLTransFile(TransHandle As Integer)
  Dim TransRec As ARTransRecType
  Dim TransLen As Integer
  TransLen = Len(TransRec)
  TransHandle = FreeFile
  Open BLTransFileName For Random Shared As TransHandle Len = TransLen
End Sub

Public Function EmpInPenProcess(EmpNum$) As Boolean
  Dim PenTrans As TempPenaltyCharges
  Dim TPHandle As Integer
  Dim NumOfPen As Integer
  Dim x As Integer
  
  EmpInPenProcess = False
  OpenPenTransFile TPHandle
  NumOfPen = LOF(TPHandle) \ Len(PenTrans)
  For x = 1 To NumOfPen
    Get TPHandle, x, PenTrans
    If QPTrim$(EmpNum$) = QPTrim$(PenTrans.CustomerNumber) Then
      EmpInPenProcess = True
      Exit For
    End If
  Next x
  Close TPHandle
    
End Function
Public Function OldRound#(n As Double)
'  OldRound# = Round(n, 2)
  OldRound# = Int(n * 100 + 0.5) / 100
End Function
Public Function GetCatRecNum(BillCat$) As Integer
  Dim x As Integer
  Dim CatRec As ARNewCatCodeRecType
  Dim CHandle  As Integer
  Dim CatRecNums As Integer
  
  GetCatRecNum = 0
  OpenCatCodeFile CHandle
  CatRecNums = LOF(CHandle) / Len(CatRec)
    
  For x = 1 To CatRecNums
    Get CHandle, x, CatRec
      If QPTrim$(CatRec.CatCode) = QPTrim$(BillCat) Then
        GetCatRecNum = x
        Exit For
      End If
  Next x
  Close CHandle
  
End Function
Public Sub OpenCatCodeFile(CatCodeHandle As Integer)
  Dim CatCodeRec As ARNewCatCodeRecType
  Dim CatCodeLen As Integer
  CatCodeLen = Len(CatCodeRec)
  CatCodeHandle = FreeFile
  Open BLCatCodeName For Random Shared As CatCodeHandle Len = CatCodeLen
End Sub
Public Function GetCatDesc(CatNum$) As String
  Dim x As Integer
  Dim CatRec As ARNewCatCodeRecType
  Dim CHandle  As Integer
  Dim CatRecNums As Integer
  
  GetCatDesc = ""
  OpenCatCodeFile CHandle
  CatRecNums = LOF(CHandle) / Len(CatRec)
  If CatRecNums = 0 Then Exit Function
  For x = 1 To CatRecNums
    Get CHandle, x, CatRec
      If QPTrim$(CatRec.CatCode) = QPTrim$(CatNum$) Then
        GetCatDesc = QPTrim$(CatRec.CODEDESC)
        Exit For
      End If
  Next x
  Close CHandle
  
End Function
Public Function GetCatRec(Cat)
  Dim CatRec As ARNewCatCodeRecType
  Dim CHandle  As Integer
  Dim CatRecNums As Integer
  OpenCatCodeFile CHandle
  CatRecNums = LOF(CHandle) / Len(CatRec)
  If CatRecNums = 0 Then Exit Function
    Get CHandle, Cat, CatRec
    GetCatRec = QPTrim$(CatRec.CODEDESC)
  Close CHandle
  
End Function

Public Sub OpenTownFile(TownRecHandle As Integer)
  Dim TownRec As TownSetUpType
  Dim TownRecLen As Integer
  TownRecLen = Len(TownRec)
  TownRecHandle = FreeFile
  Open BLTownSetUpName For Random Shared As TownRecHandle Len = TownRecLen
End Sub
Public Sub OpenCustNameIdxFile(CustIdxHandle As Integer)
  Dim CustIdx As CustNameIdxType
  Dim CustIdxLen As Integer
  CustIdxLen = Len(CustIdx)
  CustIdxHandle = FreeFile
  Open CustNameIdx For Random Shared As CustIdxHandle Len = CustIdxLen
End Sub
Public Sub OpenPayFile(PayHandle As Integer, Oper As Integer)
  Dim PayRec As AREditPaymentRecType
  Dim PayRecLen As Integer
  Dim Operator$
  
  Operator$ = Str(Oper)
  PayRecLen = Len(PayRec)
  PayHandle = FreeFile
  Open BLPayFileName + Operator$ + ".DAT" For Random Shared As PayHandle Len = PayRecLen
End Sub

