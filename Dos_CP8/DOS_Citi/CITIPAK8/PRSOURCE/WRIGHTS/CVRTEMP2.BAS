DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

'$INCLUDE: 'PREmpRec.bi'
'$INCLUDE: 'PRTRANS.bi'
'$INCLUDE: 'oPRTRANS.bi'
'$INCLUDE: 'PRFiles.bi'

  CONST oEmpData2Name = "prdata\opremp2.dat"
  CONST oTransWorkFileName = "PRDATA\PRTRANST.!!!"
  CONST oTransHistFileName = "PRDATA\PRTRANSH.!!!"
  
  REDIM oEmp2(1) AS oEmpData2Type2
  REDIM nEmp2(1) AS EmpData2Type
  REDIM oTran(1) AS oTransRecType
  REDIM nTran(1) AS TransRecType

  oEmp2Len = LEN(oEmp2(1))
  nEmp2Len = LEN(nEmp2(1))
  oTranLen = LEN(oTran(1))
  nTranLen = LEN(nTran(1))

  OPEN TransWorkFileName FOR RANDOM AS #1 LEN = oTranLen
  CLOSE
  OPEN TransHistFileName FOR RANDOM AS #1 LEN = oTranLen
  CLOSE
  OPEN EmpData2Name FOR RANDOM AS #2 LEN = oEmp2Len
  CLOSE

  CLS
  PRINT "Converting Employee 2 File. . ."
  PRINT
  

  NAME EmpData2Name AS oEmpData2Name

  OPEN oEmpData2Name FOR RANDOM AS #1 LEN = oEmp2Len
  OPEN EmpData2Name FOR RANDOM AS #2 LEN = nEmp2Len

  NumOfRec = LOF(1) \ oEmp2Len

  FOR Cnt = 1 TO NumOfRec
    
    GET #1, Cnt, oEmp2(1)
      nEmp2(1).EMPNO = oEmp2(1).EMPNO
      nEmp2(1).EMPSSN = oEmp2(1).EMPSSN
      nEmp2(1).EMPLNAME = oEmp2(1).EMPLNAME
      nEmp2(1).EMPFNAME = oEmp2(1).EMPFNAME
      nEmp2(1).EMPADDR1 = oEmp2(1).EMPADDR1
      nEmp2(1).EMPADDR2 = oEmp2(1).EMPADDR2
      nEmp2(1).EMPCITY = oEmp2(1).EMPCITY
      nEmp2(1).EMPSTATE = oEmp2(1).EMPSTATE
      nEmp2(1).EMPZIP = oEmp2(1).EMPZIP
      nEmp2(1).EMPBDAY = oEmp2(1).EMPBDAY
      nEmp2(1).EMPGENDR = oEmp2(1).EMPGENDR
      nEmp2(1).EMPRACE = oEmp2(1).EMPRACE
      nEmp2(1).EMPRETNO = oEmp2(1).EMPRETNO
      nEmp2(1).EMPRETTP = oEmp2(1).EMPRETTP

      nEmp2(1).DRAFTCOD = oEmp2(1).DRAFTCOD
      nEmp2(1).EMPDDACC = oEmp2(1).EMPDDACC
      nEmp2(1).BANKNAME = oEmp2(1).BANKNAME
      nEmp2(1).BANKLOC = oEmp2(1).BANKLOC
      nEmp2(1).TRANSIT = oEmp2(1).TRANSIT

      nEmp2(1).EMPJOB = oEmp2(1).EMPJOB
      nEmp2(1).EMPWCCLS = oEmp2(1).EMPWCCLS
      nEmp2(1).EMPSTATS = oEmp2(1).EMPSTATS
      nEmp2(1).EMPBCODE = oEmp2(1).EMPBCODE
      nEmp2(1).EMPPTYPE = oEmp2(1).EMPPTYPE
      nEmp2(1).EMPPFREQ = oEmp2(1).EMPPFREQ
      nEmp2(1).EMPPRATE = oEmp2(1).EMPPRATE
      nEmp2(1).EMPORATE = oEmp2(1).EMPORATE
      nEmp2(1).EMPHDATE = oEmp2(1).EMPHDATE
      nEmp2(1).EMPRDATE = oEmp2(1).EMPRDATE
      nEmp2(1).EMPTDATE = oEmp2(1).EMPTDATE

      nEmp2(1).EMPFEDX = oEmp2(1).EMPFEDX
      nEmp2(1).EMPFEDO2 = oEmp2(1).EMPFEDO2
      nEmp2(1).EMPFEDO1 = oEmp2(1).EMPFEDO1
      nEmp2(1).EMPFEDS = oEmp2(1).EMPFEDS
      nEmp2(1).EMPFEDA = oEmp2(1).EMPFEDA
      nEmp2(1).EMPFEDAA = oEmp2(1).EMPFEDAA
      nEmp2(1).EMPSTAX = oEmp2(1).EMPSTAX
      nEmp2(1).EMPSTAO2 = oEmp2(1).EMPSTAO2
      nEmp2(1).EMPSTAO1 = oEmp2(1).EMPSTAO1
      nEmp2(1).EMPSTAS = oEmp2(1).EMPSTAS
      nEmp2(1).EMPSTAA = oEmp2(1).EMPSTAA
      nEmp2(1).EMPSTAAA = oEmp2(1).EMPSTAAA
      nEmp2(1).EMPSOCX = oEmp2(1).EMPSOCX
      nEmp2(1).EMPMEDX = oEmp2(1).EMPMEDX
      nEmp2(1).EMPEIC = oEmp2(1).EMPEIC

      FOR zz = 1 TO 12
        LSET nEmp2(1).EmpDed(zz) = oEmp2(1).EmpDed(zz)
      NEXT
      
      nEmp2(1).EMPEACT1 = oEmp2(1).EMPEACT1
      nEmp2(1).EMPEAMT1 = oEmp2(1).EMPEAMT1
      nEmp2(1).EMPEACT2 = oEmp2(1).EMPEACT2
      nEmp2(1).EMPEAMT2 = oEmp2(1).EMPEAMT2
      nEmp2(1).EMPEACT3 = oEmp2(1).EMPEACT3
      nEmp2(1).EMPEAMT3 = oEmp2(1).EMPEAMT3
      nEmp2(1).EMPHP = oEmp2(1).EMPHP

      FOR zz = 1 TO 8
        LSET nEmp2(1).EDist(zz) = oEmp2(1).EDist(zz)
      NEXT

      IF oEmp2(1).EMPVACE < -1000 THEN
        oEmp2(1).EMPVACE = 0
      END IF
      IF oEmp2(1).EMPVUSED < -1000 THEN
        oEmp2(1).EMPVUSED = 0
      END IF
      nEmp2(1).EMPVACE = oEmp2(1).EMPVACE
      nEmp2(1).EMPVUSED = oEmp2(1).EMPVUSED
      nEmp2(1).EMPVBAL = oEmp2(1).EMPVBAL

      IF oEmp2(1).EMPSLE < -1000 THEN
        oEmp2(1).EMPSLE = 0
      END IF
      IF oEmp2(1).EMPSLUSE < -1000 THEN
        oEmp2(1).EMPSLUSE = 0
      END IF
      nEmp2(1).EMPSLE = oEmp2(1).EMPSLE
      nEmp2(1).EMPSLUSE = oEmp2(1).EMPSLUSE
      nEmp2(1).EMPSLBAL = oEmp2(1).EMPSLBAL
      
      IF oEmp2(1).EMPCTE < -1000 THEN
        oEmp2(1).EMPCTE = 0
      END IF
      IF oEmp2(1).EMPCTUSE < -1000 THEN
        oEmp2(1).EMPCTUSE = 0
      END IF

      nEmp2(1).EMPCTE = oEmp2(1).EMPCTE
      nEmp2(1).EMPCTUSE = oEmp2(1).EMPCTUSE
      nEmp2(1).EMPCTBAL = oEmp2(1).EMPCTBAL

      PERERN = 0
      PERUSED = 0
      PERBAL = 0

      HOLERN = 0
      HOLUSED = 0
      HOLBAL = 0

      nEmp2(1).LeaveTbl = oEmp2(1).LeaveTbl
      nEmp2(1).UseHosp = oEmp2(1).UseHosp
      nEmp2(1).UseLife = oEmp2(1).UseLife
      nEmp2(1).LastTransRec = oEmp2(1).LastTransRec
      nEmp2(1).EmpPin = oEmp2(1).EmpPin
      nEmp2(1).Deleted = oEmp2(1).Deleted
      nEmp2(1).PreNoteFlag = oEmp2(1).PreNoteFlag
    
    PUT #2, Cnt, nEmp2(1)
  NEXT
  CLOSE


  PRINT "Finished Employee File"
  PRINT
  PRINT "Converting Transaction Work File. . ."
  PRINT

  NAME TransWorkFileName AS oTransWorkFileName
  NAME TransHistFileName AS oTransHistFileName

  OPEN oTransWorkFileName FOR RANDOM AS #1 LEN = oTranLen
  OPEN TransWorkFileName FOR RANDOM AS #2 LEN = nTranLen

  NumOfRec = LOF(1) \ oTranLen

  FOR Cnt = 1 TO NumOfRec
    GET #1, Cnt, oTran(1)
    GOSUB Old2New
    PUT #2, Cnt, nTran(1)
  NEXT

  CLOSE

  PRINT "Finished Transaction Work File"
  PRINT
  PRINT "Converting Transaction History File. . ."
  PRINT

  OPEN oTransHistFileName FOR RANDOM AS #1 LEN = oTranLen
  OPEN TransHistFileName FOR RANDOM AS #2 LEN = nTranLen

  NumOfRec = LOF(1) \ oTranLen

  FOR Cnt = 1 TO NumOfRec
    GET #1, Cnt, oTran(1)
    GOSUB Old2New
    PUT #2, Cnt, nTran(1)
  NEXT

  CLOSE

  KILL oEmpData2Name
  KILL oTransWorkFileName
  KILL oTransHistFileName

  PRINT "Finished Transaction History File"
  PRINT
  PRINT
  PRINT
  PRINT "Conversion Complete."



END


Old2New:

nTran(1).TActive = oTran(1).TActive
nTran(1).PrevTransRec = oTran(1).PrevTransRec
nTran(1).EmpPin = oTran(1).EmpPin
nTran(1).PaySFlag = oTran(1).PaySFlag
nTran(1).CHECKNUM = oTran(1).CHECKNUM
nTran(1).PayPdStart = oTran(1).PayPdStart
nTran(1).PayPdEnd = oTran(1).PayPdEnd
nTran(1).CheckDate = oTran(1).CheckDate
nTran(1).PostDate = oTran(1).PostDate
nTran(1).PayType = oTran(1).PayType
nTran(1).BaseRate = oTran(1).BaseRate
nTran(1).OTRate = oTran(1).OTRate
nTran(1).RegHrsWork = oTran(1).RegHrsWork
nTran(1).VacUsed = oTran(1).VacUsed
nTran(1).SickUsed = oTran(1).SickUsed
nTran(1).CompUsed = oTran(1).CompUsed
nTran(1).HOLHOURS = oTran(1).HOLHOURS
nTran(1).PerHours = 0
nTran(1).RegHrsPaid = oTran(1).RegHrsPaid
nTran(1).OTHours = oTran(1).OTHours
nTran(1).OTHrsPaid = oTran(1).OTHrsPaid
nTran(1).OT2Comp = oTran(1).OT2Comp

  FOR zz = 1 TO 8
    LSET nTran(1).TDist(zz) = oTran(1).TDist(zz)
  NEXT

nTran(1).TotRegWage = oTran(1).TotRegWage
nTran(1).TotOTWage = oTran(1).TotOTWage
nTran(1).GrossWage = oTran(1).GrossWage

 FOR zz = 1 TO 3
   nTran(1).EAmt(zz) = oTran(1).EAmt(zz)
 NEXT

FOR zz = 1 TO 6
  LSET nTran(1).EDist(zz) = oTran(1).EDist(zz)
NEXT

nTran(1).TotAdditEarn = oTran(1).TotAdditEarn
nTran(1).GROSSPAY = oTran(1).GROSSPAY
nTran(1).SocGrossPay = oTran(1).SocGrossPay
nTran(1).MedGrossPay = oTran(1).MedGrossPay
nTran(1).FedGrossPay = oTran(1).FedGrossPay
nTran(1).StaGrossPay = oTran(1).StaGrossPay
nTran(1).SocTaxAmt = oTran(1).SocTaxAmt
nTran(1).MedTaxAmt = oTran(1).MedTaxAmt
nTran(1).FedTaxAmt = oTran(1).FedTaxAmt
nTran(1).StaTaxAmt = oTran(1).StaTaxAmt
nTran(1).TotTaxAmt = oTran(1).TotTaxAmt
nTran(1).RetireAmt = oTran(1).RetireAmt

FOR zz = 1 TO 12
  nTran(1).DAMT(zz) = oTran(1).DAMT(zz)
NEXT

nTran(1).TotDedAmt = oTran(1).TotDedAmt
nTran(1).EICAmt = oTran(1).EICAmt
nTran(1).NETPAY = oTran(1).NETPAY
nTran(1).PeriodHistRec = oTran(1).PeriodHistRec
nTran(1).MatchRetAmt = oTran(1).MatchRetAmt
nTran(1).MatchSocAmt = oTran(1).MatchSocAmt
nTran(1).MatchMedAmt = oTran(1).MatchMedAmt
nTran(1).RetGrossPay = oTran(1).RetGrossPay
nTran(1).TaxFring = oTran(1).TaxFring

LSET nTran(1).Pad1 = ""

RETURN

