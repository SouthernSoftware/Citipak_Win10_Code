10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM tbcustrep
30 REM Customer Report Menu
40 REM Last Revision Date: 8-24-87
50 REM by Nick Picerno
 
   DEFINT A-Z
   DECLARE FUNCTION QPTrim$ (A$)
   DECLARE FUNCTION num2date$ (Num%)
   DECLARE FUNCTION Date2Num% (Dat$)
   DECLARE FUNCTION WaitKey% ()
   DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
   DECLARE SUB PrintTitle (Title$)
   DECLARE SUB PrintLoading ()
   DECLARE SUB ShowCursor ()
   DECLARE SUB HideCursor ()
   
'$INCLUDE: 'DefCnf.BI'
'$INCLUDE: 'formedit.BI'
'$INCLUDE: 'fieldinf.BI'
'$INCLUDE: 'Qscr.BI'
'$INCLUDE: 'SetCnf.BI'
'$INCLUDE: 'lcTAX.BI'
   
60 CLS :
   tbpath$ = "F:\"                    ' change to "F:\" Before compiling
   

   CLOSE
   p = 0: Page = 0
   OPEN "R", 1, tbpath$ + "tbyear.dat", 4: FIELD 1, 4 AS Year$: GET 1, 1
   CurrentYear$ = Year$
   CLOSE 1
   
   'ON ERROR GOTO networkerror

     REDIM TBPTCust(1) AS TBPPCust
     REDIM TBPTCustBal(1) AS TBPPBal
     REDIM TTVeh(1) AS PPVehType
     DIM VRec&(200)
     DIM A$(50)
MainBody:
     CLS : scren$ = "PERS CERT"
     GOSUB 63500: COLOR 11
     LOCATE 3, 1: COLOR 15
     PRINT "PERSONAL PROPERTY CERTIFICATION TO TREASURER : FROM NEW .TXT FILES"
     PRINT
     PRINT "PLEASE CONTACT TREASURER AND MAKE SURE NO ONE IS IN THE TREASURERS SYSTEM"
     PRINT "BEFORE BEGINNING THIS PROCESS"
     PRINT
MB1: LOCATE 8, 1: COLOR 15
     PRINT "IS THIS A SUPPLEMENTAL CERTIFICATION IN CURRENT TAX YEAR (Y/N) .. "; : fl = 1
     GOSUB 62210
     IF NOT (in$ = "Y" OR in$ = "N") THEN GOTO MB1
     Sup$ = in$
mb21: LOCATE 9, 1: COLOR 15
     PRINT "Enter the Tax Year Being Certified (YYYY) .. "; : fl = -4
     GOSUB 62000
     IF LEN(in$) <> 4 THEN GOTO mb21
     TaxYear$ = in$

mb22: LOCATE 10, 1: COLOR 15
     PRINT "Enter the Last Account Number Being Transferred .. "; : fl = -5
     GOSUB 62000
     LastRecord& = VAL(in$)
     IF LastRecord& = 0 THEN GOTO mb22

MB2: LOCATE 10, 1: COLOR 15
     PRINT
     PRINT CHR$(7); CHR$(7);
     PRINT "ARE YOU SURE : LAST CHANCE BEFORE PROCESSING START (Y/N) .. "; : fl = 1
     GOSUB 62210
     IF NOT (in$ = "Y" OR in$ = "N") THEN GOTO MB1
     IF in$ = "N" THEN RUN "TBCMENU"


 
MB3: LOCATE 12, 1: COLOR 11
     PRINT "Now Updating the Treasurer's System : Please Wait!"

     'Open Files
     TBTFile = FREEFILE
     OPEN tbpath$ + "TBPTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBTFile LEN = LEN(TBPTCust(1))
     NumOfTRecs& = LOF(TBTFile) / LEN(TBPTCust(1))
     TTVFile = FREEFILE
     OPEN tbpath$ + "TBTVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TTVFile LEN = LEN(TTVeh(1))
     TTBFile = FREEFILE
     OPEN tbpath$ + "TBPTBAL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TTBFile LEN = LEN(TBPTBAL(1))


'First Check to See if Last Record From Commissioner is Greater than Last Record From Treasurer
    IF LastRecord& > NumOfTRecs& THEN
      FOR ULoop& = NumOfTRecs& + 1 TO LastRecord&
       GET TBTFile, ULoop&, TBPTCust(1)
        GOSUB AddNullAccount
       PUT TBTFile, ULoop&, TBPTCust(1)
       GET TTBFile, ULoop&, TBPTBAL(1)
       GOSUB PersNull
       PUT TTBFile, ULoop&, TBPTBAL(1)
      NEXT ULoop&
    END IF

'Next Initialize the Municipality Files Since They Still Contain Last Years Billing Info
'Initialize the Treasurer's File for Victoria/Kenbridge
'Not Necessary for Supplemental's
IF Sup$ <> "Y" THEN
     FOR Cnt& = 1 TO NumOfTRecs&
       GET TBTFile, Cnt&, TBPTCust(1)
       TBPTCust(1).TBCurYrBillNumber = ""
       TBPTCust(1).TBCurYrPersPropTaxAmount = 0
       TBPTCust(1).TBCurYrFarmEquipTaxAmount = 0
       TBPTCust(1).TBCurYrMachToolsTaxAmount = 0
       TBPTCust(1).TBCurYrMerchCapTaxAmount = 0
       TBPTCust(1).TBCurYrMobileHomeTaxAmount = 0
       TBPTCust(1).TBPPTRAValue = 0
       TBPTCust(1).TBPPTRADiscnt = 0
       TBPTCust(1).PersProp = 0
       TBPTCust(1).MachTools = 0
       TBPTCust(1).MobileHome = 0
       TBPTCust(1).MerchCap = 0
       TBPTCust(1).FarmEquip = 0
      PUT TBTFile, Cnt&, TBPTCust(1)
     NEXT Cnt&
END IF
REM First Do Owner Loop
  
    'OPEN "I", 10, tbpath$ + "owner.txt"
    OPEN "I", 10, "pplbline.txt"
oloop1:
    FOR Lnt& = 1 TO 23
     IF EOF(10) THEN CLOSE 10: GOTO ON1
     INPUT #10, A$(Lnt&)
    NEXT Lnt&
   ORecord& = VAL(A$(1))
   
        GET TBTFile, ORecord&, TBPTCust(1)
        TBPTCust(1).CustName = RTRIM$(A$(6))
        TBPTCust(1).Address1 = A$(7)
        TBPTCust(1).Address2 = A$(8)
        TBPTCust(1).City = A$(9)
        TBPTCust(1).State = A$(10)
        TBPTCust(1).Zip = LEFT$(A$(11), 5)
        TBPTCust(1).SocSec = A$(4)
        TBPTCust(1).SocSec2 = A$(5)
        TBPTCust(1).Short = A$(6)
        TBPTCust(1).Deleted = "N"
        PUT TBTFile, ORecord&, TBPTCust(1)
        LOCATE 18, 1: PRINT "Updating Main Acct # "; ORecord&
       GOTO oloop1

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
ON1:
   OPEN "I", 10, "PPLBLine.TXT"
ownerloop:
   FOR ont& = 1 TO 23
     IF EOF(10) THEN GOTO muni1
     INPUT #10, A$(ont&)
   NEXT ont&

   ORecord& = VAL(A$(1))
         LOCATE 18, 1: PRINT "Updating Detail Acct # "; ORecord&
       GET TBTFile, ORecord&, TBPTCust(1)
    'Get Current Balance if Sup$
     IF Sup$ = "Y" THEN
      CurBal# = TBPTCust(1).CurYrPersPropTax + TBPTCust(1).CurYrFarmEquipTax + TBPTCust(1).CurYrMachToolsTax + TBPTCust(1).CurYrMerchCapTax + TBPTCust(1).CurYrMobileHomeTax + TBPTCust(1).CurYrLateFeeTax + TBPTCust(1).CurYrIntAmount + TBPTCust(1). _
CurYrPenAmount
      CurBal# = INT((CurBal# * 100) + .5) / 100
     END IF
   

        TBPTCust(1).LBCurYrBillNumber = A$(2)
        District$ = RTRIM$(A$(3))
        IF LEN(District$) = 3 THEN
         TBPTCust(1).District = RIGHT$(A$(3), 2)
        END IF
        
        IF A$(12) = "LF" THEN
         TBPTCust(1).LateFile = "Y"
         TBPTCust(1).LBLateFeeAmount = VAL(A$(19))
        END IF

        IF A$(12) = "BP" THEN
         TBPTCust(1).LBCurYrPersPropTaxAmount = TBPTCust(1).LBCurYrPersPropTaxAmount + VAL(A$(17))
         TBPTCust(1).PersProp = TBPTCust(1).PersProp + VAL(A$(15))
        END IF

        IF A$(12) = "TP" THEN
          TBPTCust(1).LBCurYrPersPropTaxAmount = TBPTCust(1).LBCurYrPersPropTaxAmount + VAL(A$(17))
          TBPTCust(1).PersProp = TBPTCust(1).PersProp + VAL(A$(15))
         IF A$(23) = "Y" THEN
                        IF VAL(A$(15)) > 20000 THEN
                            TBPTCust(1).LBPPTRAValue = 20000
                        ELSE
                            TBPTCust(1).LBPPTRAValue = VAL(A$(15))
                        END IF
          TBPTCust(1).LBPPTRADiscnt = TBPTCust(1).LBPPTRADiscnt + VAL(A$(18))
         END IF

         'Add Detail Here
           IF Sup$ = "Y" AND CurBal# = 0 THEN
            TTVeh(1).VehSupFlag = "Y"
           ELSE
            TTVeh(1).VehSupFlag = ""
           END IF
           TTVeh(1).VehTyp = "I"
           TTVeh(1).VehVin = A$(14)
           TTVeh(1).VehMake = A$(21)
           TTVeh(1).VehYear = A$(22)
           TTVeh(1).VehModel = A$(13)
           TTVeh(1).VehWght = VAL(A$(20))
           TTVeh(1).VehValue = VAL(A$(15))
           TTVeh(1).VehQ = A$(23)
           TTVeh(1).VehLic = ""
           TTVeh(1).VehNext = 0
           TTVeh(1).VehAcct = ORecord&
           TTVeh(1).VehUpd = "Y"
           TTVeh(1).VehTaxYr = VAL(TaxYear$)
           TTVeh(1).VehPad = ""
           NextVehicleRecord& = (LOF(TTVFile) \ LEN(TTVeh(1))) + 1
           PUT TTVFile, NextVehicleRecord&, TTVeh(1)
        '' Update Accts Pointers
            IF TBPTCust(1).FirstVeh = 0 THEN
               TBPTCust(1).FirstVeh = NextVehicleRecord&
               TBPTCust(1).LastVeh = NextVehicleRecord&
               PUT TBTFile, ORecord&, TBPTCust(1)
                    ELSE
            PrevRecord& = TBPTCust(1).LastVeh
            TBPTCust(1).LastVeh = NextVehicleRecord&
            PUT TBTFile, ORecord&, TBPTCust(1)
            GET TTVFile, PrevRecord&, TTVeh(1)
            TTVeh(1).VehNext = NextVehicleRecord&
            PUT TTVFile, PrevRecord&, TTVeh(1)
         END IF
        END IF

        IF A$(12) = "MC" THEN
         TBPTCust(1).MerchCap = TBPTCust(1).MerchCap + VAL(A$(15))
         TBPTCust(1).LBCurYrMerchCapTaxAmount = TBPTCust(1).LBCurYrMerchCapTaxAmount + VAL(A$(17))
        END IF

        IF A$(12) = "MT" OR A$(12) = "RS" THEN
          TBPTCust(1).MachTools = TBPTCust(1).MachTools + VAL(A$(15))
          TBPTCust(1).LBCurYrMachToolsTaxAmount = TBPTCust(1).LBCurYrMachToolsTaxAmount + VAL(A$(17))
        END IF

        IF A$(12) = "MH" THEN
         TBPTCust(1).MobileHome = TBPTCust(1).MobileHome + VAL(A$(15))
         TBPTCust(1).LBCurYrMobileHomeTaxAmount = TBPTCust(1).LBCurYrMobileHomeTaxAmount + VAL(A$(17))
        END IF

        IF A$(12) = "FE" THEN
         TBPTCust(1).FarmEquip = TBPTCust(1).FarmEquip + VAL(A$(15))
         TBPTCust(1).LBCurYrFarmEquipTaxAmount = TBPTCust(1).LBCurYrFarmEquipTaxAmount + VAL(A$(17))
        END IF
        
        'TBPTCust(1).Sup2003 = "Y"

        PUT TBTFile, ORecord&, TBPTCust(1)
        GOTO ownerloop
      
     

PersNull:
FOR Lnt! = 1 TO 10
  TBPTCustBal(1).PerPropTaxAmt(Lnt!) = 0
  TBPTCustBal(1).FarmEquipAmt(Lnt!) = 0
  TBPTCustBal(1).MachToolsAmt(Lnt!) = 0
  TBPTCustBal(1).MerchCapAmt(Lnt!) = 0
  TBPTCustBal(1).MobileHomeAmt(Lnt!) = 0
  TBPTCustBal(1).LateFeeAmt(Lnt!) = 0
  TBPTCustBal(1).IntAmt(Lnt!) = 0
  TBPTCustBal(1).PenAmt(Lnt!) = 0
  TBPTCustBal(1).TaxNotice(Lnt!) = ""
NEXT Lnt!
RETURN

AddNullAccount:
       TBPTCust(1).CustName = ""
       TBPTCust(1).Address1 = ""
       TBPTCust(1).Address2 = ""
       TBPTCust(1).City = ""
       TBPTCust(1).State = ""
       TBPTCust(1).Zip = ""
       TBPTCust(1).SocSec = ""
       TBPTCust(1).SocSec2 = ""
       TBPTCust(1).Short = ""
       TBPTCust(1).PersProp = 0
       TBPTCust(1).MachTools = 0
       TBPTCust(1).MobileHome = 0
       TBPTCust(1).MerchCap = 0
       TBPTCust(1).FarmEquip = 0
       TBPTCust(1).Business = ""
       TBPTCust(1).LastChangeDate = -32767
       TBPTCust(1).District = ""
       TBPTCust(1).LateFile = ""
       TBPTCust(1).CurYrBillNumber = ""
       TBPTCust(1).CurYrPersPropTax = 0
       TBPTCust(1).CurYrFarmEquipTax = 0
       TBPTCust(1).CurYrMachToolsTax = 0
       TBPTCust(1).CurYrMerchCapTax = 0
       TBPTCust(1).CurYrMobileHomeTax = 0
       TBPTCust(1).CurYrLateFeeTax = 0
       TBPTCust(1).SocSec3 = ""
       TBPTCust(1).FirstTrans = 0
       TBPTCust(1).LastTrans = 0
       TBPTCust(1).Deleted = "Y"
       TBPTCust(1).CurYrIntAmount = 0
       TBPTCust(1).CurYrPenAmount = 0
       TBPTCust(1).LBCurYrBillNumber = ""
       TBPTCust(1).LBCurYrPersPropTaxAmount = 0
       TBPTCust(1).LBCurYrFarmEquipTaxAmount = 0
       TBPTCust(1).LBCurYrMachToolsTaxAmount = 0
       TBPTCust(1).LBCurYrMerchCapTaxAmount = 0
       TBPTCust(1).LBCurYrMobileHomeTaxAmount = 0
       TBPTCust(1).LBLateFeeAmount = 0
       TBPTCust(1).LName2 = A$(4)
       TBPTCust(1).FName2 = ""
       TBPTCust(1).FirstVeh = 0
       TBPTCust(1).LastVeh = 0
       TBPTCust(1).HeavyEquip = 0
       TBPTCust(1).Boats = 0
       TBPTCust(1).BusFurn = 0
       TBPTCust(1).TBCurYrBillNumber = ""
       TBPTCust(1).TBCurYrPersPropTaxAmount = 0
       TBPTCust(1).TBCurYrFarmEquipTaxAmount = 0
       TBPTCust(1).TBCurYrMachToolsTaxAmount = 0
       TBPTCust(1).TBCurYrMerchCapTaxAmount = 0
       TBPTCust(1).TBCurYrMobileHomeTaxAmount = 0
       TBPTCust(1).TBLateFeeAmount = 0
       TBPTCust(1).LASTYEARBILLED = 0
       TBPTCust(1).DMVSubmitted = "N"
       TBPTCust(1).DMVBatch = ""
       TBPTCust(1).DecalBought = "N"
       TBPTCust(1).DecalNumber = ""
       TBPTCust(1).LBPPTRAValue = 0
       TBPTCust(1).LBPPTRADiscnt = 0
       TBPTCust(1).TBPPTRAValue = 0
       TBPTCust(1).TBPPTRADiscnt = 0
       RETURN


muni1:      
   IF Sup$ = "Y" THEN
      CLOSE
      LOCATE 20, 1
      COLOR 11
      PRINT CHR$(7);
      PRINT "The Certification Has Completed For Supplemental Bills"
      PRINT "The SYSTEM is Ready For Printing the Supplemental Tax Bills."
      PRINT "Press <ENTER> to Continue "; : fl = 1
      GOSUB 62000
      RUN "TBCCM"
   END IF
   CLOSE 10: OPEN "I", 10, "LUNT09.TXT"
ml1:
   FOR ont& = 1 TO 23
     IF EOF(10) THEN GOTO muni2
     INPUT #10, A$(ont&)
   NEXT ont&
   IF VAL(A$(18)) > 0 THEN A$(23) = "Y" ELSE A$(23) = "N"

   ORecord& = VAL(A$(1))

       GET TBTFile, ORecord&, TBPTCust(1)


        TBPTCust(1).TBCurYrBillNumber = A$(2)

        IF A$(12) = "BP" THEN
         TBPTCust(1).TBCurYrPersPropTaxAmount = TBPTCust(1).TBCurYrPersPropTaxAmount + VAL(A$(17))
        END IF


        IF A$(12) = "TP" THEN
         TBPTCust(1).TBCurYrPersPropTaxAmount = TBPTCust(1).TBCurYrPersPropTaxAmount + VAL(A$(17))
         IF A$(23) = "Y" THEN
                        IF VAL(A$(15)) > 20000 THEN
                            TBPTCust(1).TBPPTRAValue = 20000
                        ELSE
                            TBPTCust(1).TBPPTRAValue = VAL(A$(15))
                        END IF
          TBPTCust(1).TBPPTRADiscnt = TBPTCust(1).TBPPTRADiscnt + VAL(A$(18))
         END IF
        END IF

        IF A$(12) = "MC" THEN
         TBPTCust(1).TBCurYrMerchCapTaxAmount = TBPTCust(1).TBCurYrMerchCapTaxAmount + VAL(A$(17))
        END IF

        IF A$(12) = "MT" OR A$(12) = "RS" THEN
          TBPTCust(1).TBCurYrMachToolsTaxAmount = TBPTCust(1).TBCurYrMachToolsTaxAmount + VAL(A$(17))
        END IF

        IF A$(12) = "MH" THEN
         TBPTCust(1).TBCurYrMobileHomeTaxAmount = TBPTCust(1).TBCurYrMobileHomeTaxAmount + VAL(A$(17))
        END IF

        IF A$(12) = "FE" THEN
         TBPTCust(1).TBCurYrFarmEquipTaxAmount = TBPTCust(1).TBCurYrFarmEquipTaxAmount + VAL(A$(17))
        END IF

        District$ = RTRIM$(A$(3))
        IF LEN(District$) = 3 THEN
         TBPTCust(1).District = RIGHT$(A$(3), 2)
        END IF

        
        PUT TBTFile, ORecord&, TBPTCust(1)
        GOTO ml1

muni2:

   CLOSE 10: OPEN "I", 10, "LUNT10.TXT"
ml2:
   FOR ont& = 1 TO 23
     IF EOF(10) THEN
      CLOSE
      LOCATE 20, 1
      COLOR 11
      PRINT CHR$(7);
      PRINT "The Certification Has Completed, Please Notify the Treasurer's Office"
      PRINT "that the SYSTEM is Ready For Printing the Bills."
      PRINT "Press <ENTER> to Continue "; : fl = 1
      GOSUB 62000
      RUN "TBCCM"
   END IF

     
     INPUT #10, A$(ont&)
   NEXT ont&
   IF VAL(A$(18)) > 0 THEN A$(23) = "Y" ELSE A$(23) = "N"

   ORecord& = VAL(A$(1))

       GET TBTFile, ORecord&, TBPTCust(1)


        TBPTCust(1).TBCurYrBillNumber = A$(2)



        IF A$(12) = "TP" OR A$(12) = "BP" THEN
         TBPTCust(1).TBCurYrPersPropTaxAmount = TBPTCust(1).TBCurYrPersPropTaxAmount + VAL(A$(17))
         IF A$(23) = "Y" THEN
                        IF VAL(A$(15)) > 20000 THEN
                            TBPTCust(1).TBPPTRAValue = 20000
                        ELSE
                            TBPTCust(1).TBPPTRAValue = VAL(A$(15))
                        END IF
          TBPTCust(1).TBPPTRADiscnt = TBPTCust(1).TBPPTRADiscnt + VAL(A$(18))
         END IF
        END IF

        IF A$(12) = "MC" THEN
         TBPTCust(1).TBCurYrMerchCapTaxAmount = TBPTCust(1).TBCurYrMerchCapTaxAmount + VAL(A$(17))
        END IF

        IF A$(12) = "MT" OR A$(12) = "RS" THEN
          TBPTCust(1).TBCurYrMachToolsTaxAmount = TBPTCust(1).TBCurYrMachToolsTaxAmount + VAL(A$(17))
        END IF

        IF A$(12) = "MH" THEN
         TBPTCust(1).TBCurYrMobileHomeTaxAmount = TBPTCust(1).TBCurYrMobileHomeTaxAmount + VAL(A$(17))
        END IF

        IF A$(12) = "FE" THEN
         TBPTCust(1).TBCurYrFarmEquipTaxAmount = TBPTCust(1).TBCurYrFarmEquipTaxAmount + VAL(A$(17))
        END IF

        District$ = RTRIM$(A$(3))
        IF LEN(District$) = 3 THEN
         TBPTCust(1).District = RIGHT$(A$(3), 2)
        END IF

        PUT TBTFile, ORecord&, TBPTCust(1)
        GOTO ml2

  
62000 COLOR 14: cf = 0: in$ = "": RO = CSRLIN: PO = POS(Xy): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
62010 IF INKEY$ <> "" THEN 62010
62020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
62030 z$ = INKEY$: IF z$ = "" THEN 62030
62040 IF z$ <> CHR$(13) THEN 62060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
62050 GOSUB 62200: RETURN
62060 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN cf = 1: RETURN
      IF z$ = CHR$(9) THEN cf = 2: RETURN
      IF LEN(z$) < 2 THEN 62065
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 59 THEN cf = 11: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 60 THEN cf = 12: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 61 THEN cf = 13: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 62 THEN cf = 14: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 63 THEN cf = 15: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 64 THEN cf = 16: RETURN

62065 IF ASC(z$) = 0 THEN 62030
62070 IF LEN(z$) > 1 THEN cf = ASC(RIGHT$(z$, 1)) - 57: IF cf > 0 THEN RETURN ELSE cf = cf + 27: RETURN
62080 IF z$ = " " THEN 62130 ELSE IF z$ = CHR$(29) THEN PRINT STRING$(ZL, 29); : GOTO 62000
62090 IF z$ <> CHR$(8) THEN 62130 ELSE IF ZL = 0 THEN 62030 ELSE PRINT CHR$(29); : PRINT "_"; : PRINT CHR$(29); : IF fl > 0 THEN 62110 ELSE IF z$ = "," THEN 62120
62100 IF z$ = "." THEN ZD = 0: GOTO 62110 ELSE IF z$ = "+" OR z$ = "-" THEN ZS = 0
62110 in$ = LEFT$(in$, LEN(in$) - 1)
62120 ZL = ZL - 1: GOTO 62030
62130 IF ABS(fl) = ZL THEN 62030 ELSE IF fl > 0 THEN IF z$ >= " " AND z$ <= "z" THEN 62180
62140 IF z$ = "." AND ZD = 0 THEN ZD = 1: GOTO 62180
62150 IF z$ = "," THEN PRINT ","; : ZL = ZL + 1: GOTO 62190
62160 IF (z$ = "-" OR z$ = "+") AND ZS = 0 AND ZL = 0 THEN ZS = 1: GOTO 62180
62170 IF z$ < "0" OR z$ > "9" THEN 62030
62180 PRINT z$; : in$ = in$ + z$: ZL = ZL + 1
62190 IF ABS(fl) = 1 THEN 62050 ELSE 62030
62200 LOCATE RO, PO: PRINT in$: RETURN
62210 GOSUB 62000: IF LEN(in$) THEN in$ = CHR$(ASC(in$) AND 223)
62220 RETURN
63500 LOCATE 1, 1: COLOR 10: PRINT "Sys. Date: "; SYSDATE$; TAB(70); scren$: PRINT STRING$(79, "_"): RETURN


networkerror:
     REM ERROR CODES
     Number = ERR
     SELECT CASE ERR
          CASE IS = 24, 25
               LOCATE 23, 1: PRINT STRING$(79, 32): LOCATE 23, 1: COLOR 28
               PRINT "Error Message:"; : COLOR 15: PRINT "Device Time-out! Make Sure Printer Is ON-LINE:Press ENTER to Resume Printing";
               fl = 1: GOSUB 62000: LOCATE 23, 1: PRINT STRING$(79, 32)
               RESUME
          CASE IS = 27
               LOCATE 23, 1: PRINT STRING$(79, 32): LOCATE 23, 1: COLOR 28
               PRINT "Error Message:"; : COLOR 15: PRINT "Printer Is OUT OF PAPER! Load Paper : Press ENTER to Resume Printing";
               fl = 1: GOSUB 62000: LOCATE 23, 1: PRINT STRING$(79, 32)
               RESUME

          CASE ELSE
               CLS
               LOCATE 15, 1: COLOR 28: PRINT "Error Message: "; : COLOR 15: PRINT "Error "; Number; " Has Occured in Line "; ERL
               PRINT "Write Down Message and Hit <ESC> Key to exit."
               fl = 1: GOSUB 62000: IF cf = 1 THEN CLOSE : GOTO 10 ELSE GOSUB 62000
            
          END SELECT

