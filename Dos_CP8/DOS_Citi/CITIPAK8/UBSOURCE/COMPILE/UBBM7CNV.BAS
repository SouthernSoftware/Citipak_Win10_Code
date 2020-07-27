DEFINT A-Z

DECLARE SUB SaveNewLocaRec (Form$(), LocatRec&, CustRec&)
DECLARE SUB SaveNewCustRec (FormC$(), NewRecNum&)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ClearScrn ()
DECLARE SUB SaveOldCustRec (Form$(), RecNo&)
DECLARE SUB CustLookUp (RecNo%)
DECLARE SUB SearchGetCust (SEARCH$, RecNo&, CLSFlag%)
DECLARE SUB PrintCustList ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB InsertT (SEG StartElement AS ANY, ElSize%, NumEls%)
DECLARE SUB EditCustomer ()
DECLARE SUB AddCustomer ()
DECLARE SUB LoadMeterRec (Form$())
DECLARE SUB AddNewMeter (Cust$, EdFlag)
DECLARE SUB EditMeter ()
DECLARE SUB SaveMeterRec (Form$())
DECLARE SUB SaveScrn (Array%())
DECLARE SUB RestScrn (Array%())
DECLARE SUB AddEditCustomer (RecNo&, LocatFlag%, F5Flag%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB SaveCustRec (Form$())
DECLARE SUB AddNewCustomer ()
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE SUB SaveSetUpRec (Form$())
DECLARE SUB LoadSetUpRec (Form$())
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB EditCust ()
DECLARE SUB ShowCursor ()
DECLARE SUB HideCursor ()
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB ShowCursor ()
DECLARE SUB WaitForAction ()
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB HideCursor ()
DECLARE SUB LoadCustRec (Form$(), RecNo&, LocatRec&, TFormLen%, LockedFlag%)
DECLARE SUB BlockClear ()
DECLARE FUNCTION GetNumOfAcct% ()
DECLARE FUNCTION Date2Num (Daty$)
DECLARE FUNCTION Num2Date$ (Daty%)
DECLARE FUNCTION GetNumOfAcct% ()
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION Compare3% (BYVAL Seg1, BYVAL Adr1, BYVAL Seg2, BYVAL Adr2, NumBytes)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION QPValI% (Number$)
DECLARE FUNCTION Round# (DblNumber#)
  '$INCLUDE: 'CP72UTIL.bi'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubtrans.bi'
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  
  CLEAR
  
  DIM WBCustRec  AS CustRecType

  REDIM UBCustRec(1) AS NewUBCustRecType

  CLS
  WBCustRecLen = 512
  UBCustRecLen = LEN(UBCustRec(1))


s1:
 CLS
 COLOR 11
 PRINT "Beech Mountain Conversion Program"

 LOCATE 3, 1: PRINT "Enter Beginning Account to Convert (ESC-End) .. "; : fl = -9
 GOSUB 62000
 IF cf = 1 THEN END
 BegAcct$ = in$
 IF LEN(in$) <> 9 THEN GOTO s1
 COLOR 11
 LOCATE 4, 1: PRINT "Enter Ending Account to Convert ..... "; : fl = -9
 GOSUB 62000
 EndAcct$ = in$
 IF LEN(in$) <> 9 THEN GOTO s1
 IF EndAcct$ < BegAcct$ THEN GOTO s1
 COLOR 11
 LOCATE 7, 1: PRINT "Enter the Book to Put These Accounts Into .. "; : fl = -2
 GOSUB 62000
 Book$ = LTRIM$(in$)
 IF LEN(Book$) = 0 THEN GOTO s1
 IF LEN(Book$) = 1 THEN Book$ = "0" + Book$
LOCATE 10, 1: PRINT "Press <ESC> to Cancel or <ENTER> to Process .. "; : fl = 1
GOSUB 62000
IF cf = 1 THEN CLOSE : END




'GOTO Skip2Here
  
  CustHandle = FREEFILE
  OPEN "wbcust.dat" FOR RANDOM AS CustHandle LEN = WBCustRecLen


  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  
  NumOfRecs = LOF(CustHandle) \ WBCustRecLen
  
  
    FOR Cnt = 1 TO NumOfRecs
      REDIM UBCustRec(1) AS NewUBCustRecType
      GET #CustHandle, Cnt, WBCustRec
    
       COLOR 15
      LOCATE 12, 1: PRINT "Converting Customer Record "; Cnt; " Of "; NumOfRecs

      SEQNUMB = SEQNUMB + 10
      SEQNUMB$ = LTRIM$(STR$(SEQNUMB))
      IF LEN(SEQNUMB$) < 6 THEN SEQNUMB$ = STRING$(6 - LEN(SEQNUMB$), "0") + SEQNUMB$
      
    IF LEN(QPTrim$(WBCustRec.firstname)) = 0 THEN
       CUSTNAME$ = QPTrim$(WBCustRec.LNAME)
     ELSE
      CUSTNAME$ = QPTrim$(WBCustRec.firstname) + " " + WBCustRec.LNAME
    END IF

    CUSTNAME$ = QPTrim$(UCASE$(CUSTNAME$))

   
    IF UCASE$((QPTrim$(WBCustRec.Active))) = "Y" THEN
      UBCustRec(1).Status = "A"
    ELSE
      UBCustRec(1).Status = "I"
    END IF

    UBCustRec(1).Book = Book$
    UBCustRec(1).SEQNUMB = SEQNUMB$



    IF UBCustRec(1).Status = "I" THEN
      UBCustRec(1).Book = ""
      UBCustRec(1).SEQNUMB = ""
    END IF

    UBCustRec(1).OPENDATE = Date2Num%(WBCustRec.AcctOpen)
    UBCustRec(1).SEARCH = WBCustRec.LNAME
    UBCustRec(1).CUSTNAME = CUSTNAME$
    UBCustRec(1).ADDR1 = WBCustRec.ADD1
    UBCustRec(1).ADDR2 = WBCustRec.ADD3
    UBCustRec(1).SERVADDR = WBCustRec.ADD2
    UBCustRec(1).City = WBCustRec.City
    UBCustRec(1).STATE = WBCustRec.ST
    UBCustRec(1).ZipCode = WBCustRec.Zip
    UBCustRec(1).HPHONE = WBCustRec.PHONE
    UBCustRec(1).WPHONE = ""
    UBCustRec(1).SOSEC = ""
    UBCustRec(1).DRVLIC = ""
    UBCustRec(1).CUSTTYPE = WBCustRec.CUSTTYPE
    UBCustRec(1).Addr911 = ""
    UBCustRec(1).BILLCOPY = 1
    UBCustRec(1).POSTRTE = ""

    UBCustRec(1).BILLCYCL = 0
    UBCustRec(1).ZONE = WBCustRec.CityLimit
    UBCustRec(1).SEQ = 0
    UBCustRec(1).CASHONLY = "N"
    UBCustRec(1).LATEFEE = "Y"
    UBCustRec(1).CUTOFFYN = "Y"
    UBCustRec(1).TAXEXPT = "N"
    UBCustRec(1).SRCIT = "N"
    UBCustRec(1).EPP = "N"
    UBCustRec(1).EPPAMT = 0
    UBCustRec(1).USEDRAFT = "N"

    UBCustRec(1).BANKNAME = ""
    UBCustRec(1).BANKLOC = ""
    UBCustRec(1).TRANSIT = ""
    UBCustRec(1).BANKACCT = ""

    UBCustRec(1).BILLCMNT = ""
    UBCustRec(1).PAYCMNT = ""
    UBCustRec(1).PUMPCODE = ""
    UBCustRec(1).USERCODE1 = ""
    UBCustRec(1).USERCODE2 = ""
    UBCustRec(1).HHMSG1 = ""
    UBCustRec(1).HHMSG2 = ""
    UBCustRec(1).HHMSG3 = ""
    
    FOR LL = 1 TO 15
     UBCustRec(1).Serv(LL).RateCode = ""
     UBCustRec(1).Serv(LL).RMTRTYPE = ""
    NEXT LL

    FOR LL = 1 TO 7
     UBCustRec(1).LocMeters(LL).MTRNUM = ""
     UBCustRec(1).LocMeters(LL).MTRMulti = 1
     UBCustRec(1).LocMeters(LL).MTRType = ""
     UBCustRec(1).LocMeters(LL).MTRUnit = ""
     UBCustRec(1).LocMeters(LL).NumUser = 1
     UBCustRec(1).LocMeters(LL).InsDate = -32767
     UBCustRec(1).LocMeters(LL).CurRead = 0
     UBCustRec(1).LocMeters(LL).PrevRead = 0
     UBCustRec(1).LocMeters(LL).CurDate = -32767
     UBCustRec(1).LocMeters(LL).PastDate = -32767
     UBCustRec(1).LocMeters(LL).ReadFlag = "N"
     UBCustRec(1).LocMeters(LL).AvgUse = 0
     UBCustRec(1).LocMeters(LL).UseCnt = 1
    NEXT LL

    
    FOR LL = 1 TO 4
     UBCustRec(1).FlatRates(LL).FRDESC = ""
     UBCustRec(1).FlatRates(LL).FRAMT = 0
     UBCustRec(1).FlatRates(LL).FRFREQ = ""
     UBCustRec(1).FlatRates(LL).REVSRC = 0
     UBCustRec(1).FlatRates(LL).NumMin = 0
    NEXT LL


    'Fix Services First
     RateCode = WBCustRec.WRate
     SewerCode = VAL(WBCustRec.SewerRate)

     IF RateCode = 1 THEN
       IF WBCustRec.CityLimit = "O" THEN
         UBCustRec(1).Serv(1).RateCode = "WRO1"
         UBCustRec(1).Serv(1).RMTRTYPE = "C"
        IF SewerCode = 1 THEN
         UBCustRec(1).Serv(2).RateCode = "SRO1"
         UBCustRec(1).Serv(2).RMTRTYPE = "C"
        END IF
       END IF
      ELSE
        UBCustRec(1).Serv(1).RateCode = "WRI1"
        UBCustRec(1).Serv(1).RMTRTYPE = "C"
       IF SewerCode = 1 THEN
        UBCustRec(1).Serv(2).RateCode = "SRI1"
        UBCustRec(1).Serv(2).RMTRTYPE = "C"
       END IF
     END IF
     IF RateCode = 2 THEN
       IF WBCustRec.CityLimit = "O" THEN
         UBCustRec(1).Serv(1).RateCode = "WRO2"
         UBCustRec(1).Serv(1).RMTRTYPE = "C"
        IF SewerCode = 1 THEN
         UBCustRec(1).Serv(2).RateCode = "SRO2"
         UBCustRec(1).Serv(2).RMTRTYPE = "C"
        END IF
       END IF
      ELSE
        UBCustRec(1).Serv(1).RateCode = "WRI2"
        UBCustRec(1).Serv(1).RMTRTYPE = "C"
       IF SewerCode = 1 THEN
        UBCustRec(1).Serv(2).RateCode = "SRI2"
        UBCustRec(1).Serv(2).RMTRTYPE = "C"
       END IF
     END IF
     IF RateCode = 3 THEN
       IF WBCustRec.CityLimit = "O" THEN
         UBCustRec(1).Serv(1).RateCode = "WRO3"
         UBCustRec(1).Serv(1).RMTRTYPE = "C"
        IF SewerCode = 1 THEN
         UBCustRec(1).Serv(2).RateCode = "SRO3"
         UBCustRec(1).Serv(2).RMTRTYPE = "C"
        END IF
       END IF
      ELSE
        UBCustRec(1).Serv(1).RateCode = "WRI3"
        UBCustRec(1).Serv(1).RMTRTYPE = "C"
       IF SewerCode = 1 THEN
        UBCustRec(1).Serv(2).RateCode = "SRI3"
        UBCustRec(1).Serv(2).RMTRTYPE = "C"
       END IF
     END IF
     IF RateCode = 4 THEN
       IF WBCustRec.CityLimit = "O" THEN
         UBCustRec(1).Serv(1).RateCode = "WRO4"
         UBCustRec(1).Serv(1).RMTRTYPE = "C"
        IF SewerCode = 1 THEN
         UBCustRec(1).Serv(2).RateCode = "SRO4"
         UBCustRec(1).Serv(2).RMTRTYPE = "C"
        END IF
       END IF
      ELSE
        UBCustRec(1).Serv(1).RateCode = "WRI4"
        UBCustRec(1).Serv(1).RMTRTYPE = "C"
       IF SewerCode = 1 THEN
        UBCustRec(1).Serv(2).RateCode = "SRI4"
        UBCustRec(1).Serv(2).RMTRTYPE = "C"
       END IF
     END IF
     IF RateCode = 5 THEN
       IF WBCustRec.CityLimit = "O" THEN
         UBCustRec(1).Serv(1).RateCode = "WRO5"
         UBCustRec(1).Serv(1).RMTRTYPE = "C"
        IF SewerCode = 1 THEN
         UBCustRec(1).Serv(2).RateCode = "SRO5"
         UBCustRec(1).Serv(2).RMTRTYPE = "C"
        END IF
       END IF
      ELSE
        UBCustRec(1).Serv(1).RateCode = "WRI5"
        UBCustRec(1).Serv(1).RMTRTYPE = "C"
       IF SewerCode = 1 THEN
        UBCustRec(1).Serv(2).RateCode = "SRI5"
        UBCustRec(1).Serv(2).RMTRTYPE = "C"
       END IF
     END IF


     'Recycle Here
     Units = WBCustRec.RUnits
     IF Units > 0 THEN
        UBCustRec(1).Serv(5).RateCode = "R" + LTRIM$(STR$(Units))
     END IF



     IF WBCustRec.ramt1 > 0 THEN
      LL = 1
      UBCustRec(1).FlatRates(LL).FRDESC = WBCustRec.RDesc1
      UBCustRec(1).FlatRates(LL).FRAMT = WBCustRec.ramt1
      UBCustRec(1).FlatRates(LL).FRFREQ = "R"
      UBCustRec(1).FlatRates(LL).REVSRC = VAL(LEFT$(WBCustRec.racct1, 1))
      UBCustRec(1).FlatRates(LL).NumMin = 1
    END IF
    IF WBCustRec.ramt2 > 0 THEN
      LL = 2
    
      UBCustRec(1).FlatRates(LL).FRDESC = WBCustRec.RDesc2
      UBCustRec(1).FlatRates(LL).FRAMT = WBCustRec.ramt2
      UBCustRec(1).FlatRates(LL).FRFREQ = "R"
      UBCustRec(1).FlatRates(LL).REVSRC = VAL(LEFT$(WBCustRec.racct2, 1))
      UBCustRec(1).FlatRates(LL).NumMin = 1
    END IF
    IF WBCustRec.ramt3 > 0 THEN
      LL = 3
      UBCustRec(1).FlatRates(LL).FRDESC = WBCustRec.RDesc3
      UBCustRec(1).FlatRates(LL).FRAMT = WBCustRec.ramt3
      UBCustRec(1).FlatRates(LL).FRFREQ = "R"
      UBCustRec(1).FlatRates(LL).REVSRC = VAL(LEFT$(WBCustRec.racct3, 1))
      UBCustRec(1).FlatRates(LL).NumMin = 1
    END IF


   
    UBCustRec(1).Monthly(1).AMTOWED = 0
    UBCustRec(1).Monthly(1).TotAmtPD = 0
    UBCustRec(1).Monthly(1).PayAmt = 0
    UBCustRec(1).Monthly(1).RevSource = 0

    UBCustRec(1).Monthly(2).AMTOWED = 0
    UBCustRec(1).Monthly(2).TotAmtPD = 0
    UBCustRec(1).Monthly(2).PayAmt = 0
    UBCustRec(1).Monthly(2).RevSource = 0
    

    UBCustRec(1).MFEE1 = 0
    UBCustRec(1).MFEE2 = 0



    'Set Meter Screen
    LL = 1
    UBCustRec(1).LocMeters(LL).MTRNUM = LTRIM$(WBCustRec.Meter)
    UBCustRec(1).LocMeters(LL).MTRMulti = 100
    UBCustRec(1).LocMeters(LL).MTRType = "C"
    UBCustRec(1).LocMeters(LL).MTRUnit = "G"
    UBCustRec(1).LocMeters(LL).NumUser = WBCustRec.WMin
    UBCustRec(1).LocMeters(LL).InsDate = -32767
    UBCustRec(1).LocMeters(LL).CurRead = WBCustRec.CurRead
    UBCustRec(1).LocMeters(LL).PrevRead = WBCustRec.PastRead
    UBCustRec(1).LocMeters(LL).CurDate = -32767
    UBCustRec(1).LocMeters(LL).PastDate = -32767
    UBCustRec(1).LocMeters(LL).ReadFlag = "N"
    UBCustRec(1).LocMeters(LL).AvgUse = 0
    UBCustRec(1).LocMeters(LL).UseCnt = 1
   
'END OF Quick Screen Form
    UBCustRec(1).CustPIN = Cnt
    UBCustRec(1).LastTrans = 0
    UBCustRec(1).CurrBalance = 0
    UBCustRec(1).PrevBalance = 0
    UBCustRec(1).CurrRevAmts(1) = 0
    UBCustRec(1).CurrRevAmts(2) = 0
    UBCustRec(1).CurrRevAmts(3) = 0
    UBCustRec(1).CurrRevAmts(4) = 0
    UBCustRec(1).CurrRevAmts(5) = 0
    UBCustRec(1).CurrRevAmts(6) = 0
    UBCustRec(1).CurrRevAmts(7) = 0
    UBCustRec(1).CurrRevAmts(8) = 0

    FOR LL = 9 TO 15
     UBCustRec(1).CurrRevAmts(LL) = 0
    NEXT LL

    FOR LL = 1 TO 15
     UBCustRec(1).PrevRevAmts(LL) = 0
    NEXT LL

    UBCustRec(1).DepositAmt = 0
    IF VAL(WBCustRec.Meter) = 0 THEN
     UBCustRec(1).DelFlag = -1
     ELSE
     UBCustRec(1).DelFlag = 0
     END IF
    UBCustRec(1).PreNoteFlag = 0
    UBCustRec(1).WOLastTrans = 0
    UBCustRec(1).EstFlag = "N"
    UBCustRec(1).MessageRec = 0
    UBCustRec(1).FillPad = ""
    
    PUT UBFile, Cnt, UBCustRec(1)
    NEXT Cnt
    CLOSE
    RUN


62000 cf = 0: in$ = "": COLOR 14: RO = CSRLIN: PO = POS(x): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
62010 IF INKEY$ <> "" THEN 62010
62020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
62030 Z$ = INKEY$: IF Z$ = "" THEN 62030
62040 IF Z$ <> CHR$(13) THEN 62060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
62050 GOSUB 62200: RETURN
62060 IF Z$ = CHR$(27) OR Z$ = CHR$(21) THEN cf = 1: RETURN
      IF Z$ = CHR$(9) THEN cf = 2: RETURN
      IF LEN(Z$) <> 2 THEN 62065
      IF ASC(LEFT$(Z$, 1)) = 0 AND ASC(MID$(Z$, 2, 1)) = 59 THEN cf = 11: RETURN
      IF ASC(LEFT$(Z$, 1)) = 0 AND ASC(MID$(Z$, 2, 1)) = 60 THEN cf = 12: RETURN
62065 IF ASC(Z$) = 0 THEN 62030
62070 IF LEN(Z$) > 1 THEN cf = ASC(RIGHT$(Z$, 1)) - 57: IF cf > 0 THEN RETURN ELSE cf = cf + 27: RETURN
62080 IF Z$ = " " THEN 62130 ELSE IF Z$ = CHR$(29) THEN PRINT STRING$(ZL, 29); : GOTO 62000
62090 IF Z$ <> CHR$(8) THEN 62130 ELSE IF ZL = 0 THEN 62030 ELSE PRINT CHR$(29); : PRINT "_"; : PRINT CHR$(29); : IF fl > 0 THEN 62110 ELSE IF Z$ = "," THEN 62120
62100 IF Z$ = "." THEN ZD = 0: GOTO 62110 ELSE IF Z$ = "+" OR Z$ = "-" THEN ZS = 0
62110 in$ = LEFT$(in$, LEN(in$) - 1)
62120 ZL = ZL - 1: GOTO 62030
62130 IF ABS(fl) = ZL THEN 62030 ELSE IF fl > 0 THEN IF Z$ >= " " AND Z$ <= "z" THEN 62180
62140 IF Z$ = "." AND INSTR(in$, ".") = 0 THEN GOTO 62180
62150 IF Z$ = "," THEN PRINT ","; : ZL = ZL + 1: GOTO 62190
62160 IF (Z$ = "-" OR Z$ = "+") AND ZS = 0 AND ZL = 0 THEN ZS = 1: GOTO 62180
62170 IF Z$ < "0" OR Z$ > "9" THEN 62030
62180 PRINT Z$; : in$ = in$ + Z$: ZL = ZL + 1
62190 IF ABS(fl) = 1 THEN 62050 ELSE 62030
62200 LOCATE RO, PO: PRINT in$: RETURN
62210 GOSUB 62000: IF LEN(in$) THEN in$ = CHR$(ASC(in$) AND 223)
62220 RETURN

