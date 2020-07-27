DEFINT A-Z

DEFSNG A-Z
DEFINT A-Z
SUB CheckCustFile

  REDIM OCust(1) AS NewUBCustRecType2
  REDIM NCust(1) AS NewUBCustRecType

  OCustLen = LEN(OCust(1))
  NCustLen = LEN(NCust(1))

  IF CnvrtCust THEN
    PRINT "Converting Customer/Owner File Structure"
    OPEN "UBCUST.DAT" FOR RANDOM SHARED AS #1 LEN = OCustLen
    OPEN "UBCUST.DAT" FOR RANDOM SHARED AS #2 LEN = NCustLen

    NumOfCust = LOF(1) \ OCustLen

    FOR Cnt = NumOfCust TO 1 STEP -1
      GET #1, Cnt, OCust(1)
      DidCnt = DidCnt + 1
      LOCATE , 1
      PRINT "Converting:"; DidCnt; " of"; NumOfCust;

      NCust(1).Book = OCust(1).Book
      NCust(1).SEQNUMB = OCust(1).SEQNUMB
      NCust(1).Status = OCust(1).Status
      NCust(1).OPENDATE = OCust(1).OPENDATE
      NCust(1).SEARCH = OCust(1).SEARCH
      NCust(1).CustName = OCust(1).CustName
      NCust(1).ADDR1 = OCust(1).ADDR1
      NCust(1).ADDR2 = OCust(1).ADDR2
      NCust(1).SERVADDR = OCust(1).SERVADDR
      NCust(1).CITY = OCust(1).CITY
      NCust(1).STATE = OCust(1).STATE
      NCust(1).ZIPCODE = OCust(1).ZIPCODE
      NCust(1).HPHONE = OCust(1).HPHONE
      NCust(1).WPHONE = OCust(1).WPHONE
      NCust(1).SOSEC = OCust(1).SOSEC
      NCust(1).DRVLIC = OCust(1).DRVLIC
      NCust(1).CUSTTYPE = OCust(1).CUSTTYPE
      NCust(1).Addr911 = OCust(1).Addr911
      NCust(1).BILLCOPY = OCust(1).BILLCOPY
      NCust(1).POSTRTE = OCust(1).POSTRTE
      NCust(1).BILLCYCL = OCust(1).BILLCYCL
      NCust(1).ZONE = OCust(1).ZONE
      NCust(1).SEQ = OCust(1).SEQ
      NCust(1).CASHONLY = OCust(1).CASHONLY
      NCust(1).LATEFEE = OCust(1).LATEFEE
      NCust(1).CUTOFFYN = OCust(1).CUTOFFYN
      NCust(1).TAXEXPT = OCust(1).TAXEXPT
      NCust(1).SRCIT = OCust(1).SRCIT
      NCust(1).EPP = OCust(1).EPP
      NCust(1).EPPAMT = OCust(1).EPPAMT
      NCust(1).USEDRAFT = OCust(1).USEDRAFT
      NCust(1).BANKNAME = OCust(1).BANKNAME
      NCust(1).BANKLOC = OCust(1).BANKLOC
      NCust(1).TRANSIT = OCust(1).TRANSIT
      NCust(1).BANKACCT = OCust(1).BANKACCT
      NCust(1).BILLCMNT = OCust(1).BILLCMNT
      NCust(1).PAYCMNT = OCust(1).PAYCMNT
      NCust(1).PUMPCODE = OCust(1).PUMPCODE
      NCust(1).USERCODE1 = OCust(1).USERCODE1
      NCust(1).USERCODE2 = OCust(1).USERCODE2
      NCust(1).HHMSG1 = OCust(1).HHMSG1
      NCust(1).HHMSG2 = OCust(1).HHMSG2
      NCust(1).HHMSG3 = OCust(1).HHMSG3

      FOR zz = 1 TO 15
        LSET NCust(1).Serv(zz) = OCust(1).Serv(zz)
        NCust(1).CurrRevAmts(zz) = OCust(1).CurrRevAmts(zz)
        NCust(1).PrevRevAmts(zz) = OCust(1).PrevRevAmts(zz)
      NEXT
      FOR zz = 1 TO 4
        LSET NCust(1).FlatRates(zz) = OCust(1).FlatRates(zz)
      NEXT
      FOR zz = 1 TO 2
        LSET NCust(1).Monthly(zz) = OCust(1).Monthly(zz)
      NEXT

      NCust(1).MFEE1 = OCust(1).MFEE1
      NCust(1).MFEE2 = OCust(1).MFEE2

      FOR zz = 1 TO 7
        LSET NCust(1).LocMeters(zz) = OCust(1).LocMeters(zz)
      NEXT

      NCust(1).CustPIN = OCust(1).CustPIN
      NCust(1).LastTrans = OCust(1).LastTrans
      NCust(1).CurrBalance = OCust(1).CurrBalance
      NCust(1).PrevBalance = OCust(1).PrevBalance
      NCust(1).DepositAmt = OCust(1).DepositAmt
      NCust(1).DelFlag = OCust(1).DelFlag
      NCust(1).PreNoteFlag = OCust(1).PreNoteFlag
      NCust(1).WOLastTrans = OCust(1).WOLastTrans
      NCust(1).EstFlag = OCust(1).EstFlag
      NCust(1).MessageRec = OCust(1).MessageRec
      NCust(1).OldRec = OCust(1).OldRec
      NCust(1).FillPad = OCust(1).FillPad
      PUT #2, Cnt, NCust(1)
    NEXT
    CLOSE
    LOCATE , 1
    PRINT "Customer/Owner File Conversion Completed OK."

  ELSE
    PRINT "Customer/Owner File Structure is OK"
  END IF

END SUB

DEFSNG A-Z
DEFINT A-Z
SUB CheckSetupFile

  OSetLen = LEN(OSetup)
  NSetLen = LEN(NSetup)

  IF FileSize("UBSETUP.DAT") <> NSetLen THEN
    PRINT "Converting Setup File Structure"

    OPEN "UBSETUP.DAT" FOR RANDOM AS #1 LEN = OSetLen
    GET #1, 1, OSetup
    CLOSE #1

    NSetup.UTILNAME = OSetup.UTILNAME
    NSetup.DEFCITY = OSetup.DEFCITY
    NSetup.DEFSTATE = OSetup.DEFSTATE
    NSetup.ZIPCODE = OSetup.ZIPCODE
    NSetup.PreByBook = OSetup.PreByBook
    NSetup.RecpPort = OSetup.RecpPort
    NSetup.RECPDEFT = OSetup.RECPDEFT
    NSetup.ESTREAD = OSetup.ESTREAD
    NSetup.BANKDFT = OSetup.BANKDFT
    NSetup.UseSeq = OSetup.UseSeq
    NSetup.BILLCYCL = OSetup.BILLCYCL
    NSetup.DefLook = OSetup.DefLook
    NSetup.MethAcct = OSetup.MethAcct
    NSetup.SkipInactive = OSetup.SkipInactive
    NSetup.SkipSeparator = OSetup.SkipSeparator
    'NSetup.PrnBStatus = OSetup.PrnBStatus
    NSetup.LowRead = OSetup.LowRead
    NSetup.HighRead = OSetup.HighRead
    NSetup.HHDEVICE = OSetup.HHDEVICE
    NSetup.FileVer = 3.01

    FOR zz = 1 TO 15
      LSET NSetup.Revenues(zz) = OSetup.Revenues(zz)
      LSET NSetup.BillAcct(zz) = OSetup.BillAcct(zz)
      LSET NSetup.PayAcct(zz) = OSetup.PayAcct(zz)
      LSET NSetup.DepAcct(zz) = OSetup.DepAcct(zz)
    NEXT
   
    OPEN "UBSETUP.DAT" FOR RANDOM AS #1 LEN = NSetLen
    PUT #1, 1, NSetup
    CLOSE #1
    PRINT "Setup File Conversion Complete OK."
    CnvrtCust = True
  ELSE
    PRINT "Setup File Structure is OK."
    CnvrtCust = False
  END IF


END SUB

