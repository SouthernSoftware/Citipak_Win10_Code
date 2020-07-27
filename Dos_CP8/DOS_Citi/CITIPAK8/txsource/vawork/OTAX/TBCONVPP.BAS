10 DEFINT A-Z
CLEAR , , 4000

'$INCLUDE: 'TAX.BI'
  WIDTH LPRINT 136
'  ON ERROR GOTO networkerror
  DIM SHARED fields$(40)


  REDIM TBCust(1) AS TBPPCustomer
  TBfile = FREEFILE
  OPEN "TBPPCUS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBfile LEN = LEN(TBCust(1))
  OPEN "I", 10, "pp.dta"
1 IF EOF(10) THEN END
  FOR ll = 1 TO 28
  INPUT #10, a$
  fields$(ll) = a$
  NEXT ll
  PRINT fields$(1); TAB(20); fields$(21); fields$(2)
  PRINT STRING$(79, "-")
  GOSUB updaterecord
  GOTO 1
updaterecord:
  TBCust(1).Account = fields$(1)
  TBCust(1).CustName = fields$(2)
  TBCust(1).Address1 = fields$(3)
  TBCust(1).Address2 = fields$(4)
  TBCust(1).City = fields$(5)
  TBCust(1).State = fields$(6)
  TBCust(1).Zip = fields$(7)
  TBCust(1).SocSec = fields$(8)
  TBCust(1).SocSec2 = fields$(9)
  TBCust(1).TType = fields$(10)
  TBCust(1).Short = fields$(11)
  TBCust(1).PersProp = fields$(12)
  TBCust(1).MachTools = fields$(13)
  TBCust(1).MobileHome = fields$(14)
  TBCust(1).MerchCap = fields$(15)
  TBCust(1).FarmEquip = fields$(16)
  TBCust(1).Business = fields$(17)
  TBCust(1).LastChangeDate = fields$(18)
  TBCust(1).District = fields$(19)
  TBCust(1).LateFile = fields$(20)
  TBCust(1).BillNumb = fields$(21)
  TBCust(1).PersPropTax = fields$(22)
  TBCust(1).FarmEquipTax = fields$(23)
  TBCust(1).MachToolsTax = fields$(24)
  TBCust(1).MerchCapTax = fields$(25)
  TBCust(1).MobileHomeTax = fields$(26)
  TBCust(1).LateFeeTax = fields$(27)
  TBCust(1).SocSec3 = fields$(28)
  Count = Count + 1
  PUT TBfile, Count, TBCust(1)
RETURN





110 scren$ = "3.00.00": SYSDATE$ = DATE$
120 CLS : GOSUB 63500
130 LOCATE 4, 1: COLOR 11
3000 CLS : scren$ = "3.03.00": GOSUB 63500: COLOR 11
     LOCATE 1, 32: PRINT "Print Tax Bills"
     LOCATE 3, 1: COLOR 11
     PRINT "Enter the Pers. Prop. Tax Rate Per $100 (Ex. .50 for 50 cents / $100): "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     PPtaxrate = VAL(in$)
     LOCATE 4, 1: COLOR 11
     PRINT "Enter the Machinery and Tools Tax Rate Per $100 (ESC-End) : "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     MTtaxrate = VAL(in$)
     LOCATE 5, 1: COLOR 11
     PRINT "Enter the Merchants Capital Tax Rate Per $100 (ESC-End)   : "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     MCtaxrate = VAL(in$)
     LOCATE 6, 1: COLOR 11
     PRINT "Enter the Mobile Home Tax Rate Per $100 (ESC-End)         : "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     MHtaxrate = VAL(in$)
     LOCATE 7, 1: COLOR 11
     PRINT "Enter the Farm Equipment Tax Rate Per $100 (ESC-End)      : "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     FEtaxrate = VAL(in$)


3001 LOCATE 10, 1: COLOR 11
     PRINT "Print (A)ll or (I)ndividual Bills or (ESC) to End .. "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN RUN "tbbillin"
     IF NOT (in$ = "A" OR in$ = "I") THEN 3001
     IF in$ = "A" THEN
       GOSUB taxalign
3035  ' GET A SORTED INDEX FIRST
      TBfile = FREEFILE
      OPEN "TBPPCUS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBfile LEN = LEN(TBCust(1))
      OPEN "TBPBILL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #9 LEN = 24
      FIELD 9, 20 AS nothing$, 4 AS Record$
      LastRecord = LOF(9) / 24
       Count = 1
    WHILE Count <= LastRecord
     GET 9, Count
     GET TBfile, CVS(Record$), TBCust(1)
     GOSUB PrintBill
     Count = Count + 1
WEND
CLOSE
RUN "tbbillin"
ELSE
COLOR 11
LOCATE 15, 1: PRINT "Individual Printing"
GOSUB taxalign
mainbody:
 LOCATE 17, 1: PRINT "Enter Customer Acct Number to Reprint (ESC-END) .. "; : fl = -7
 GOSUB 62000
 IF cf = 1 THEN LPRINT CHR$(12); : CLOSE : RUN "tbbillin"
      AcctNumber = VAL(in$)
      TBfile = FREEFILE
      OPEN "TBRECUS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBfile LEN = LEN(TBCust(1))
      LastRecord = LOF(TBfile) / LEN(TBCust(1))
      Count = 1
WHILE Count <= LastRecord
     GET TBfile, Count, TBCust(1)
     IF VAL(TBCust(1).Account) = AcctNumber THEN GOSUB PrintBill: CLOSE TBfile: GOTO mainbody
    Count = Count + 1
WEND
    CLOSE TBfile
    GOTO mainbody

END IF





taxalign:
     LOCATE 17, 1: COLOR 11
     PRINT "Print Alignment (Y/N) or (ESC-ABORT PROCESS) .. "; : fl = 1: GOSUB 62210
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     IF NOT (in$ = "Y" OR in$ = "N") THEN GOTO taxalign
     IF in$ = "N" THEN RETURN
     LPRINT
     LPRINT TAB(18); "XXXXXXXXXXXXXX"; TAB(85); "XXXXX"
     LPRINT
     LPRINT
     LPRINT TAB(22); "X.XX"; TAB(27); "XXXXXXXX"; TAB(37); "XXXXXX.XX"
     LPRINT TAB(22); "X.XX"; TAB(27); "XXXXXXXX"; TAB(37); "XXXXXX.XX"
     LPRINT TAB(22); "X.XX"; TAB(27); "XXXXXXXX"; TAB(37); "XXXXXX.XX"
     LPRINT TAB(27); "XXXXXXXX"
     LPRINT TAB(22); "X.XX"; TAB(27); "XXXXXXXX"; TAB(37); "XXXXXX.XX"
     LPRINT TAB(22); "X.XX"; TAB(27); "XXXXXXXX"; TAB(37); "XXXXXX.XX"
     LPRINT
     LPRINT TAB(37); "XXXXXX.XX"; TAB(47); "XXXXX.XX"; TAB(56); "XXXXXX.XX"
     LPRINT
     LPRINT TAB(10); STRING$(35, "X")
     LPRINT TAB(10); STRING$(35, "X")
     LPRINT TAB(10); STRING$(35, "X")
     LPRINT TAB(10); STRING$(35, "X")
     LPRINT
     LPRINT
     LPRINT
     LPRINT
     LPRINT TAB(55); "Acct # XXXXXXX"
     GOTO taxalign

PrintBill:
     GOSUB GetDistrict
     TotalTaxAmount@ = (VAL(TBCust(1).PersPropTax) / 100) + (VAL(TBCust(1).MachToolsTax) / 100) + (VAL(TBCust(1).MerchCapTax) / 100) + (VAL(TBCust(1).MobileHomeTax) / 100) + (VAL(TBCust(1).FarmEquipTax) / 100)
     TotalTaxAmount@ = INT((TotalTaxAmount@ * 100) + .50001) / 100
     IF TBCust(1).LateFile = "Y" THEN
      TotalTaxAmount@ = TotalTaxAmount@ + (VAL(TBCust(1).LateFeeTax) / 100)
      TotalTaxAmount@ = INT((TotalTaxAmount@ * 100) + .50001) / 100
     END IF
     Penalty@ = TotalTaxAmount@ * .1
     Penalty@ = INT((Penalty@ * 100) + .50001) / 100


     ' Print the Bill Here
     LPRINT
     LPRINT TAB(18); DistrictName$; TAB(85); TBCust(1).BillNumb
     LPRINT
     LPRINT
     LPRINT TAB(22); USING "#.##"; PPtaxrate;
      LPRINT TAB(27); USING "########"; TBCust(1).PersProp;
       LPRINT TAB(37); USING "######.##"; TBCust(1).PersPropTax
     LPRINT TAB(22); USING "#.##"; MTtaxrate;
      LPRINT TAB(27); USING "########"; TBCust(1).MachTools;
       LPRINT TAB(37); USING "######.##"; TBCust(1).MachToolsTax
     LPRINT TAB(22); USING "#.##"; MCtaxrate;
      LPRINT TAB(27); USING "########"; TBCust(1).MerchCap;
       LPRINT TAB(37); USING "######.##"; TBCust(1).MerchCapTax
     IF TBCust(1).LateFile = "Y" THEN
       LPRINT TAB(37); USING "######.##"; TBCust(1).LateFeeTax
     ELSE
       LPRINT
     END IF
     LPRINT TAB(22); USING "#.##"; MHtaxrate;
      LPRINT TAB(27); USING "########"; TBCust(1).MobileHome;
       LPRINT TAB(37); USING "######.##"; TBCust(1).MobileHomeTax
     LPRINT TAB(22); USING "#.##"; FEtaxrate;
      LPRINT TAB(27); USING "########"; TBCust(1).FarmEquip;
       LPRINT TAB(37); USING "######.##"; TBCust(1).FarmEquipTax
     LPRINT
     LPRINT TAB(37); USING "######.##"; TotalTaxAmount@;
      LPRINT TAB(47); USING "#####.##"; Penalty@;
       LPRINT TAB(56); USING "######.##"; (TotalTaxAmount@ + Penalty@)
     LPRINT
     LPRINT TAB(10); TBCust(1).CustName
     LPRINT TAB(10); TBCust(1).Address1
     LPRINT TAB(10); TBCust(1).Address2
     LPRINT TAB(10); RTRIM$(TBCust(1).City); " "; TBCust(1).State; "  "; TBCust(1).Zip
     LPRINT
     LPRINT
     LPRINT
     LPRINT
     LPRINT TAB(55); "Acct #"; TBCust(1).Account
     RETURN

GetDistrict:
     ' Only Towns of Kenbridge and Victoria Need to Be Defined for Personal Property
     DistrictName$ = ""
     IF VAL(TBCust(1).District) = 9 THEN DistrictName$ = "TOWN OF KENBRIDGE"
     IF VAL(TBCust(1).District) = 10 THEN DistrictName$ = "TOWN OF VICTORIA"
     RETURN


62000 COLOR 14: cf = 0: in$ = "": RO = CSRLIN: PO = POS(x): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
62010 IF INKEY$ <> "" THEN 62010
62020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
62030 Z$ = INKEY$: IF Z$ = "" THEN 62030
62040 IF Z$ <> CHR$(13) THEN 62060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
62050 GOSUB 62200: RETURN
62060 IF Z$ = CHR$(27) OR Z$ = CHR$(21) THEN cf = 1: RETURN
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

