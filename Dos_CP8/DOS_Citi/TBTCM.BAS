  
  
DECLARE FUNCTION Num2Date$ (Num%)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PrintLoading ()
  
  
  '$INCLUDE: 'LCTax.Bi'
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  ''$INCLUDE: 'Qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  
  
  
10 CLS
  '   ON ERROR GOTO networkerror
  TBPath$ = "F:\"               ' change to "F:\" Before compiling
  
  REDIM TBRECust(1) AS TBRECust
  REDIM TBRECustBal(1) AS TBREBAL
  REDIM TBPPCust(1) AS TBPPCust
  REDIM TBPPCustBal(1) AS TBPPBal
  REDIM TBTrans(1) AS TBTrans
  REDIM TBVeh(1) AS PPVehType
  REDIM TBEX(1) AS TBExceptionType
  DIM VREC!(450)
  
  CLOSE : OPEN "R", 1, TBPath$ + "tbyear.dat", 4: FIELD 1, 4 AS year$: GET 1, 1
  CURRENTYEAR$ = year$: CLOSE
  
100 GOSUB 63000:                REM get town in town$
110 SCREN$ = "1.00.00": SYSDATE$ = DATE$
  CLOSE : OPEN "R", 1, TBPath$ + "tbyear.dat", 4: FIELD 1, 4 AS year$: GET 1, 1
  CURRENTYEAR$ = year$: CLOSE
120 CLS : GOSUB 63500
130 LOCATE 4, 1: COLOR 11
200 PRINT TAB(31); "Tax Billing System"
210 PRINT TAB(24); "Customer Maintenance - Treasurer"
220 PRINT : PRINT
230 PRINT TAB(28); "1. Change/View Customer"
240 PRINT
250 PRINT TAB(28); "2. Re-Sort the Customer Data "
260 PRINT
270 PRINT TAB(28); "3. Fix Customer History"
280 PRINT
310 LOCATE 15, 1: COLOR 11
  PRINT TAB(24); "Selection No. (ESC-Master Menu) : "; : fl = -2: GOSUB 62000
320 IF cf = 1 THEN CLOSE : RUN "tbtmenu"
325 IF LEN(in$) = 0 THEN 310
330 IF VAL(in$) < 1 OR VAL(in$) > 3 THEN 310
340 ON VAL(in$) GOTO 2000, 4000, 10000
  GOTO 10
  
  
1111 LOCATE 4, 28: fl = 3: GOSUB 62000
  IF cf = 1 THEN CLOSE : RUN
  IF LEN(in$) < 3 THEN in$ = in$ + STRING$(3 - LEN(in$), 32)
  TBRECust(1).MapNumber = LEFT$(UCASE$(in$), 3)
  LOCATE 4, 32: fl = 2: GOSUB 62000
  IF cf = 1 THEN CLOSE : RUN
  IF LEN(in$) < 2 THEN in$ = in$ + STRING$(2 - LEN(in$), 32)
  PUT TBFile, Record!, TBRECust(1)
  GET TBFile, Record!, TBRECust(1)
  TBRECust(1).MapNumber = LEFT$(TBRECust(1).MapNumber, 3) + LEFT$(UCASE$(in$), 2) + RIGHT$(TBRECust(1).MapNumber, 9)
  
  LOCATE 4, 35: fl = 2: GOSUB 62000
  IF cf = 1 THEN CLOSE : RUN
  IF LEN(in$) < 2 THEN in$ = in$ + STRING$(2 - LEN(in$), 32)
  PUT TBFile, Record!, TBRECust(1)
  GET TBFile, Record!, TBRECust(1)
  TBRECust(1).MapNumber = LEFT$(TBRECust(1).MapNumber, 5) + LEFT$(UCASE$(in$), 2) + RIGHT$(TBRECust(1).MapNumber, 7)
  
  LOCATE 4, 38: fl = 2: GOSUB 62000
  IF cf = 1 THEN CLOSE : RUN
  IF LEN(in$) < 2 THEN in$ = in$ + STRING$(2 - LEN(in$), 32)
  PUT TBFile, Record!, TBRECust(1)
  GET TBFile, Record!, TBRECust(1)
  TBRECust(1).MapNumber = LEFT$(TBRECust(1).MapNumber, 7) + LEFT$(UCASE$(in$), 2) + RIGHT$(TBRECust(1).MapNumber, 5)
  
  LOCATE 4, 41: fl = 4: GOSUB 62000
  IF cf = 1 THEN CLOSE : RUN
  IF LEN(in$) < 4 THEN in$ = in$ + STRING$(4 - LEN(in$), 32)
  PUT TBFile, Record!, TBRECust(1)
  GET TBFile, Record!, TBRECust(1)
  TBRECust(1).MapNumber = LEFT$(TBRECust(1).MapNumber, 9) + LEFT$(UCASE$(in$), 4) + RIGHT$(TBRECust(1).MapNumber, 1)
  
  LOCATE 4, 46: fl = 1: GOSUB 62000
  IF cf = 1 THEN CLOSE : RUN
  IF LEN(in$) < 1 THEN in$ = in$ + STRING$(1 - LEN(in$), 32)
  PUT TBFile, Record!, TBRECust(1)
  GET TBFile, Record!, TBRECust(1)
  TBRECust(1).MapNumber = LEFT$(TBRECust(1).MapNumber, 13) + LEFT$(UCASE$(in$), 1)
  PUT TBFile, Record!, TBRECust(1)
  RETURN
  
1112 LOCATE 5, 28: fl = 35: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBRECust(1).CustName = UCASE$(in$): RETURN
1113 LOCATE 6, 28: fl = 35: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBRECust(1).Address1 = UCASE$(in$): RETURN
1114 LOCATE 7, 28: fl = 35: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBRECust(1).Address2 = UCASE$(in$): RETURN
1115 LOCATE 8, 28: fl = 35: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBRECust(1).Address3 = UCASE$(in$): RETURN
1116 LOCATE 9, 28: fl = 5: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBRECust(1).Zipcode = UCASE$(in$): RETURN
1117 LOCATE 10, 28: fl = 25: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBRECust(1).PropDesc1 = UCASE$(in$): RETURN
1118 LOCATE 11, 28: fl = 25: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBRECust(1).PropDesc2 = UCASE$(in$): RETURN
1119 LOCATE 12, 28: fl = 25: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBRECust(1).PropDesc3 = UCASE$(in$): RETURN
1120 LOCATE 13, 28: fl = 25: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBRECust(1).PropDesc4 = UCASE$(in$): RETURN
1121 LOCATE 14, 28: fl = 2: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBRECust(1).Zone = UCASE$(in$): RETURN
1122 LOCATE 15, 28: fl = 2: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBRECust(1).ClassCode = UCASE$(in$): RETURN
1123 LOCATE 16, 28: fl = 2: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBRECust(1).District = UCASE$(in$): RETURN
1124 LOCATE 17, 28: fl = 12: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBRECust(1).Acres = in$: RETURN
1125 LOCATE 18, 28: fl = 20: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBRECust(1).Comment1 = UCASE$(in$): RETURN
1126 LOCATE 19, 28: fl = 20: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBRECust(1).Comment2 = UCASE$(in$): RETURN
1127 LOCATE 20, 28: fl = 20: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBRECust(1).Comment3 = UCASE$(in$): RETURN
  
1151 LOCATE 4, 30: fl = 25: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBPPCust(1).CustName = UCASE$(in$): RETURN
1152 LOCATE 5, 30: fl = 25: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBPPCust(1).Address1 = UCASE$(in$): RETURN
1153 LOCATE 6, 30: fl = 25: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBPPCust(1).Address2 = UCASE$(in$): RETURN
1154 LOCATE 7, 30: fl = 20: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBPPCust(1).City = UCASE$(in$): RETURN
1155 LOCATE 8, 30: fl = 2: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBPPCust(1).State = UCASE$(in$): RETURN
1156 LOCATE 9, 30: fl = 10: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBPPCust(1).Zip = UCASE$(in$): RETURN
1157 LOCATE 10, 30: fl = 12: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBPPCust(1).SocSec = UCASE$(in$): RETURN
1158 LOCATE 11, 30: fl = 12: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBPPCust(1).SocSec2 = UCASE$(in$): RETURN
1159 LOCATE 12, 30: fl = 2: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBPPCust(1).District = UCASE$(in$): RETURN
1160 LOCATE 13, 30: fl = 1: GOSUB 62210
  IF cf = 1 THEN CLOSE : RUN
  IF NOT (in$ = "Y" OR in$ = "N") THEN 1160
  TBPPCust(1).LateFile = UCASE$(in$)
  RETURN
1161 LOCATE 14, 30: fl = 1: GOSUB 62210: IF cf = 1 THEN CLOSE : RUN
  IF NOT (in$ = "Y" OR in$ = "N") THEN 1161
  TBPPCust(1).Business = UCASE$(in$)
  RETURN
1163 LOCATE 15, 30: PRINT STRING$(15, 32): LOCATE 15, 30
  fl = 10: GOSUB 62000:      IF cf = 1 THEN CLOSE : RUN
  TBPPCust(1).MachTools = VAL(in$)
  RETURN
1164 LOCATE 16, 30: PRINT STRING$(15, 32): LOCATE 16, 30
  fl = 10: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBPPCust(1).MobileHome = VAL(in$)
  RETURN
1165 LOCATE 17, 30: PRINT STRING$(15, 32): LOCATE 17, 30
  fl = 10: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBPPCust(1).MerchCap = VAL(in$)
  RETURN
1166 LOCATE 18, 30: PRINT STRING$(15, 32): LOCATE 18, 30
  fl = 10: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBPPCust(1).FarmEquip = VAL(in$)
  RETURN
1167 LOCATE 4, 30: PRINT STRING$(15, 32): LOCATE 4, 30
  fl = 10: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBPPCust(1).HeavyEquip = VAL(in$)
  RETURN
1168 LOCATE 5, 30: PRINT STRING$(15, 32): LOCATE 5, 30
  fl = 10: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBPPCust(1).Boats = VAL(in$)
  RETURN
1169 LOCATE 6, 30: PRINT STRING$(15, 32): LOCATE 6, 30
  fl = 10: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
  TBPPCust(1).BusFurn = VAL(in$)
  RETURN
1170 LOCATE 10, 30: PRINT STRING$(25, 32): LOCATE 10, 30
  fl = 22: GOSUB 62000
  TBVeh(1).VehVin = UCASE$(in$)
  RETURN
1171 LOCATE 11, 30: PRINT STRING$(25, 32): LOCATE 11, 30
  fl = 4: GOSUB 62000
  TBVeh(1).VehYear = UCASE$(in$)
  RETURN
1172 LOCATE 12, 30: PRINT STRING$(25, 32): LOCATE 12, 30
  fl = 20: GOSUB 62000
  TBVeh(1).VehMake = UCASE$(in$)
  RETURN
1173 LOCATE 13, 30: PRINT STRING$(25, 32): LOCATE 13, 30
  fl = 15: GOSUB 62000
  TBVeh(1).VehModel = UCASE$(in$)
  RETURN
1174 LOCATE 14, 30: PRINT STRING$(15, 32): LOCATE 14, 30
  fl = 15: GOSUB 62000
  TBVeh(1).VehWght = VAL(in$)
  RETURN
1175 LOCATE 15, 30: PRINT STRING$(15, 32): LOCATE 15, 30
  fl = -10: GOSUB 62000
  TBVeh(1).VehValue = VAL(in$)
  RETURN
1176 LOCATE 16, 30: PRINT STRING$(15, 32): LOCATE 16, 30
  fl = 1: GOSUB 62000
  IF NOT (in$ = "Y" OR in$ = "N") THEN 1176
  TBVeh(1).VehQ = UCASE$(in$)
  RETURN
1177 LOCATE 17, 30: PRINT STRING$(15, 32): LOCATE 17, 30
  fl = 10: GOSUB 62000
  TBVeh(1).VehLic = UCASE$(in$)
  RETURN
  
  
1211 LOCATE 6, 30: PRINT STRING$(15, 32): LOCATE 6, 30: fl = 10: GOSUB 62000: TBRECust(1).LandValue = VAL(in$): RETURN
1212 LOCATE 7, 30: PRINT STRING$(15, 32): LOCATE 7, 30: fl = 10: GOSUB 62000: TBRECust(1).BldgValue = VAL(in$): RETURN
1213 LOCATE 8, 30: PRINT STRING$(15, 32): LOCATE 8, 30: fl = 10: GOSUB 62000: TBRECust(1).LandUseValue = VAL(in$): RETURN
1214 LOCATE 9, 30: fl = 1: GOSUB 62210
  IF NOT (in$ = "Y" OR in$ = "N") THEN 1214
  TBRECust(1).NewConst = in$
  RETURN
1215 LOCATE 10, 30: fl = 1: GOSUB 62210
  IF NOT (in$ = "Y" OR in$ = "N") THEN 1215
  TBRECust(1).NewValue = in$
  RETURN
1216 LOCATE 11, 30: fl = 1: GOSUB 62210
  IF NOT (in$ = "Y" OR in$ = "N") THEN 1216
  TBRECust(1).NewAddress = in$
  RETURN
1217 LOCATE 12, 30: fl = 2: GOSUB 62000
  TBRECust(1).TransferYr = in$
  RETURN
1218 LOCATE 13, 30: fl = -10: GOSUB 62000
  TBRECust(1).NewBldgValue = VAL(in$)
  RETURN
1219 LOCATE 14, 30: fl = 25: GOSUB 62000
  TBRECust(1).NewCom1 = in$
  RETURN
1220 LOCATE 15, 30: fl = 25: GOSUB 62000
  TBRECust(1).NewCom2 = in$
  RETURN
  
  
  
2000          REM
2010 SCREN$ = "1.02.00": CLS : GOSUB 63500
2020 LOCATE 1, 32: COLOR 11: PRINT "Change/View Customer"; : LOCATE 3, 1
  PRINT "(R)eal or (P)ersonal or (S)upplemental Account (R,P,S or ESC-Menu) .. "; : fl = 1
  GOSUB 62210
  IF cf = 1 THEN 10
  IF NOT (in$ = "R" OR in$ = "P" OR in$ = "S") THEN 2020
  taxtype$ = in$
2030 LOCATE 4, 1: COLOR 11
  PRINT "Enter Customer Account Number (TAB-Last Name : ESC-Menu) .. "; : fl = -6
  GOSUB 62000
  IF cf = 1 THEN
    IF nameflag = 1 THEN
      RUN "tbnmsort"
    ELSE
      CLOSE : GOTO 10
    END IF
  END IF
  IF cf = 2 THEN
    GOSUB Namesearch
    IF ggo = 1 THEN 2075
    IF ggo = 0 THEN LOCATE 15, 1: PRINT "Account Not Found : Press Any Key to Continue "; : fl = 1: GOSUB 62000: GOTO 2000
  END IF
  Record! = VAL(in$)
2075 CLOSE : GOSUB 60000
  LOCATE 3, 1: PRINT STRING$(79, 32)
  IF taxtype$ = "R" THEN
    IF Record! < 1 OR Record! > LOF(TBFile) / LEN(TBRECust(1)) THEN 2030
    'LOCK #1, Record!
    GET 1, Record!, TBRECust(1)
    IF TBRECust(1).Deleted = "Y" THEN GOSUB Deleted: GOTO 2000
  ELSE
    IF Record! < 1 OR Record! > LOF(TBFile) / LEN(TBPPCust(1)) THEN 2030
    'LOCK #1, Record!
    GET 1, Record!, TBPPCust(1)
    IF TBPPCust(1).Deleted = "Y" THEN GOSUB Deleted: GOTO 2000
  END IF
  IF taxtype$ = "R" THEN
    Account! = Record!
    GOSUB AddReal1
    COLOR 15
    LOCATE 3, 60: PRINT USING "######"; Account!
    LOCATE 4, 28: PRINT LEFT$(TBRECust(1).MapNumber, 3)
    LOCATE 4, 32: PRINT MID$(TBRECust(1).MapNumber, 4, 2)
    LOCATE 4, 35: PRINT MID$(TBRECust(1).MapNumber, 6, 2)
    LOCATE 4, 38: PRINT MID$(TBRECust(1).MapNumber, 8, 2)
    LOCATE 4, 41: PRINT MID$(TBRECust(1).MapNumber, 10, 4)
    LOCATE 4, 46: PRINT MID$(TBRECust(1).MapNumber, 14, 1)
    LOCATE 5, 28: PRINT TBRECust(1).CustName
    LOCATE 6, 28: PRINT TBRECust(1).Address1
    LOCATE 7, 28: PRINT TBRECust(1).Address2
    LOCATE 8, 28: PRINT TBRECust(1).Address3
    LOCATE 9, 28: PRINT TBRECust(1).Zipcode
    LOCATE 10, 28: PRINT TBRECust(1).PropDesc1
    LOCATE 11, 28: PRINT TBRECust(1).PropDesc2
    LOCATE 12, 28: PRINT TBRECust(1).PropDesc3
    LOCATE 13, 28: PRINT TBRECust(1).PropDesc4
    LOCATE 14, 28: PRINT TBRECust(1).Zone
    LOCATE 15, 28: PRINT TBRECust(1).ClassCode
    LOCATE 16, 28: PRINT TBRECust(1).District
    LOCATE 17, 28: PRINT TBRECust(1).Acres
    LOCATE 18, 28: PRINT TBRECust(1).Comment1
    LOCATE 19, 28: PRINT TBRECust(1).Comment2
    LOCATE 20, 28: PRINT TBRECust(1).Comment3
    
REselection2:
    
    LOCATE 23, 1: COLOR 11
    PRINT "Select # (ESC-OK) .. "; : fl = -2
    GOSUB 62000
    IF cf = 1 THEN 2200
    IF VAL(in$) < 1 OR VAL(in$) > 17 THEN GOTO REselection2
    ON VAL(in$) GOSUB 1111, 1112, 1113, 1114, 1115, 1116, 1117, 1118, 1119, 1120, 1121, 1122, 1123, 1124, 1125, 1126, 1127
    PUT TBFile, Record!, TBRECust(1)
    GOTO REselection2
    
  ELSE
    
    REM put personal here
    Account! = Record!
    GOSUB AddPers1
    COLOR 15
    LOCATE 3, 60: PRINT USING "######"; Account!
    LOCATE 4, 30: PRINT TBPPCust(1).CustName
    LOCATE 5, 30: PRINT TBPPCust(1).Address1
    LOCATE 6, 30: PRINT TBPPCust(1).Address2
    LOCATE 7, 30: PRINT TBPPCust(1).City
    LOCATE 8, 30: PRINT TBPPCust(1).State
    LOCATE 9, 30: PRINT TBPPCust(1).Zip
    LOCATE 10, 30: PRINT TBPPCust(1).SocSec
    LOCATE 11, 30: PRINT TBPPCust(1).SocSec2
    LOCATE 12, 30: PRINT TBPPCust(1).District
    LOCATE 13, 30: PRINT TBPPCust(1).LateFile
    LOCATE 14, 30: PRINT TBPPCust(1).Business
    LOCATE 15, 30: PRINT USING "########,#"; TBPPCust(1).MachTools
    LOCATE 16, 30: PRINT USING "########,#"; TBPPCust(1).MobileHome
    LOCATE 17, 30: PRINT USING "########,#"; TBPPCust(1).MerchCap
    LOCATE 18, 30: PRINT USING "########,#"; TBPPCust(1).FarmEquip
    
PPselection2:
    
    LOCATE 23, 1: COLOR 11
    PRINT "Select # (ESC-OK) .. "; : fl = -2
    GOSUB 62000
    IF cf = 1 THEN 2300
    IF VAL(in$) < 1 OR VAL(in$) > 15 THEN GOTO PPselection2
    ON VAL(in$) GOSUB 1151, 1152, 1153, 1154, 1155, 1156, 1157, 1158, 1159, 1160, 1161, 1163, 1164, 1165, 1166
    PUT TBFile, Record!, TBPPCust(1)
    GOTO PPselection2
  END IF
  
2200 IF taxtype$ = "R" THEN
  GOSUB AddReal2
  
  COLOR 11: LOCATE 3, 1
  PRINT STRING$(79, 32): LOCATE 3, 1
  PRINT "Customer Account #"; : COLOR 15: PRINT Account!
  
  LOCATE 6, 30: PRINT USING "########,#"; TBRECust(1).LandValue
  LOCATE 7, 30: PRINT USING "########,#"; TBRECust(1).BldgValue
  LOCATE 8, 30: PRINT USING "########,#"; TBRECust(1).LandUseValue
  LOCATE 9, 30: PRINT TBRECust(1).NewConst
  LOCATE 10, 30: PRINT TBRECust(1).NewValue
  LOCATE 11, 30: PRINT TBRECust(1).NewAddress
  LOCATE 12, 30: PRINT TBRECust(1).TransferYr
  LOCATE 13, 30: PRINT USING "########,#"; TBRECust(1).NewBldgValue
  LOCATE 14, 30: PRINT TBRECust(1).NewCom1
  LOCATE 15, 30: PRINT TBRECust(1).NewCom2
2210 LOCATE 18, 1: COLOR 11
  PRINT "Selection # (ESC-OK) .. "; : fl = -2: GOSUB 62000
  IF cf = 1 THEN PUT TBFile, Record!, TBRECust(1): CLOSE : GOTO 10
  IF VAL(in$) < 1 OR VAL(in$) > 10 THEN 2210
  ON VAL(in$) GOSUB 1211, 1212, 1213, 1214, 1215, 1216, 1217, 1218, 1219, 1220
  GOTO 2210
END IF

2300 GOSUB AddPers2
  'Create VehRecord List
  VCNT = 0
  VehRecord! = TBPPCust(1).FirstVeh
  WHILE VehRecord! <> 0
    GET TBVFile, VehRecord!, TBVeh(1)
    IF TBVeh(1).VehTyp <> "D" THEN
      VCNT = VCNT + 1
      VREC!(VCNT) = VehRecord!
      VREC = 1
    END IF
    VehRecord! = TBVeh(1).VehNext
  WEND
  
  
  COLOR 15
  LOCATE 3, 60: PRINT USING "######"; Account!
  LOCATE 4, 30: PRINT USING "########,#"; TBPPCust(1).HeavyEquip
  LOCATE 5, 30: PRINT USING "########,#"; TBPPCust(1).Boats
  LOCATE 6, 30: PRINT USING "########,#"; TBPPCust(1).BusFurn
PPselection3:
  IF VCNT > 0 THEN
    GET TBVFile, VREC!(VREC), TBVeh(1)
    IF TBVeh(1).VehTaxYr = 8224 THEN
      TBVeh(1).VehTaxYr = 1998
      PUT TBVFile, VREC!(VREC), TBVeh(1)
    END IF
    COLOR 15
    LOCATE 10, 30: PRINT TBVeh(1).VehVin: COLOR 11: LOCATE 10, 50: PRINT "#"; VREC; " of "; VCNT
    COLOR 15
    LOCATE 11, 30: PRINT TBVeh(1).VehYear
    LOCATE 12, 30: PRINT TBVeh(1).VehMake
    LOCATE 13, 30: PRINT TBVeh(1).VehModel
    LOCATE 14, 30: PRINT TBVeh(1).VehWght
    LOCATE 15, 30: PRINT USING "######,#"; TBVeh(1).VehValue
    LOCATE 16, 30: PRINT TBVeh(1).VehQ
    LOCATE 17, 30: PRINT TBVeh(1).VehLic
    LOCATE 18, 30: PRINT TBVeh(1).VehTyp
    LOCATE 20, 30: PRINT TBVeh(1).VehTaxYr
  END IF
  
  'Recalculate the Personal Property Figure Here
  TotPers# = TBPPCust(1).HeavyEquip + TBPPCust(1).Boats + TBPPCust(1).BusFurn
  VehTr! = TBPPCust(1).FirstVeh
  WHILE VehTr! <> 0
    GET TBVFile, VehTr!, TBVeh(1)
    IF TBVeh(1).VehTyp <> "D" THEN
      TotPers# = TotPers# + TBVeh(1).VehValue
    END IF
    VehTr! = TBVeh(1).VehNext
  WEND
  TotPers# = INT((TotPers# * 100) + .5) / 100
  LOCATE 8, 30: PRINT USING "########,#"; TotPers#
  
  
  
  LOCATE 23, 1: COLOR 11
  PRINT "Select # (ESC-OK : <N>-Next Vehicle : <P>-Previous Vehicle) .. "; : fl = 2
  GOSUB 62000
  IF cf = 1 THEN CLOSE : GOTO 10
  IF UCASE$(in$) = "N" THEN
    'Get Next Vehicle if Any
    VREC = VREC + 1
    IF VREC > VCNT THEN VREC = VCNT
    IF VCNT > 0 THEN GET TBVFile, VREC!(VREC), TBVeh(1): GOTO PPselection3
  END IF
  IF UCASE$(in$) = "P" THEN
    'Get Prev Vehicle if Any
    VREC = VREC - 1
    IF VREC < 1 THEN VREC = 1
    IF VCNT > 0 THEN GET TBVFile, VREC!(VREC), TBVeh(1): GOTO PPselection3
  END IF
  
  IF VAL(in$) = 13 THEN
    'Add New Vehicle Here
    LOCATE 10, 1
    COLOR 11
    PRINT " 4- Vehicle VIN ............                                     "
    PRINT " 5- Vehicle Year (YYYY) ....                     "
    PRINT " 6- Vehicle Make ...........                       "
    PRINT " 7- Vehicle Model ..........                       "
    PRINT " 8- Vehicle Weight .........                          "
    PRINT " 9- Vehicle Value ..........                            "
    PRINT "10- Qualifying Vehicle Y/N .    "
    PRINT "11- Vehicle License # ......                    "
    PRINT "12- Type/Delete (I/B/D) ....  "
    PRINT "13- Add New Vehicle"
    
    GOSUB 1170
    GOSUB 1171
    GOSUB 1172
    GOSUB 1173
    GOSUB 1174
    GOSUB 1175
    GOSUB 1176
    GOSUB 1177
    TBVeh(1).VehTyp = "A": TBVeh(1).VehPad = ""
    PUT TBVFile, (LOF(TBVFile) \ LEN(TBVeh(1))) + 1, TBVeh(1)
    'Update Pointer Here
    IF TBPPCust(1).LastVeh = 0 THEN
      TBPPCust(1).FirstVeh = (LOF(TBVFile) \ LEN(TBVeh(1)))
      TBPPCust(1).LastVeh = (LOF(TBVFile) \ LEN(TBVeh(1)))
      PUT TBFile, Record!, TBPPCust(1)
    ELSE
      PrevVeh = TBPPCust(1).LastVeh
      TBPPCust(1).LastVeh = (LOF(TBVFile) \ LEN(TBVeh(1)))
      PUT TBFile, Record!, TBPPCust(1)
      GET TBVFile, PrevVeh, TBVeh(1)
      TBVeh(1).VehNext = (LOF(TBVFile) \ LEN(TBVeh(1)))
      PUT TBVFile, PrevVeh, TBVeh(1)
    END IF
  END IF
  
  
  
  IF VAL(in$) >= 1 AND VAL(in$) <= 3 THEN
    ON VAL(in$) GOSUB 1167, 1168, 1169
    PUT TBFile, Record!, TBPPCust(1)
    GOTO PPselection3
  END IF
  
  IF VAL(in$) >= 4 AND VAL(in$) <= 12 THEN
    ON (VAL(in$) - 3) GOSUB 1170, 1171, 1172, 1173, 1174, 1175, 1176, 1177
    PUT TBVFile, VREC!(VREC), TBVeh(1)
  END IF
  
  GOTO PPselection3
  
  
4000 CLS
  COLOR 11
  PRINT "Do you wish to sort now (Y/N) .. "; : fl = 1
  GOSUB 62210
  IF in$ = "Y" THEN
    RUN "TBTNMSRT"
  ELSE
    CLOSE : GOTO 10
  END IF
  
AddReal1:
  FOR zz = 4 TO 23: LOCATE zz, 1: PRINT STRING$(79, 32): NEXT zz: COLOR 11
    LOCATE 3, 1
    COLOR 10
    PRINT "Real Estate Account Data"; TAB(50); "Account # "
    COLOR 11
    PRINT " 1- Map Number ...........    -  -  -  -    -"
    PRINT " 2- Name/Address ......... "
    PRINT " 3- Address Line 1 ....... "
    PRINT " 4- Address Line 2 ....... "
    PRINT " 5- Address Line 3 ....... "
    PRINT " 6- Zip Code ............. "
    PRINT " 7- Prop Descrip Line 1 .."
    PRINT " 8- Prop Descrip Line 2 .."
    PRINT " 9- Prop Descrip Line 3 .."
    PRINT "10- Prop Descrip Line 4 .."
    PRINT "11- Zone ................."
    PRINT "12- Class Code ..........."
    PRINT "13- District ............."
    PRINT "14- Size/Acreage ........."
    PRINT "15- Comment Line 1 ......."
    PRINT "16- Comment Line 2 ......."
    PRINT "17- Comment Line 3 ......."
    RETURN
    
AddReal2:
    FOR zz = 4 TO 23: LOCATE zz, 1: PRINT STRING$(79, 32): NEXT zz: COLOR 11
      LOCATE 5, 1
      PRINT "Property Valuation Data"
      PRINT " 1- Land Value ..........."
      PRINT " 2- Bldg Value ..........."
      PRINT " 3- Land Use Exempt ......"
      PRINT " 4- New Construction Y/N ."
      PRINT " 5- New Value Y/N ........"
      PRINT " 6- New Address .........."
      PRINT " 7- Trans Year (yy) ......"
      PRINT " 8- New Bldg Value ......."
      PRINT " 9- New Bldg Comment 1 ..."
      PRINT "10- New Bldg Comment 2 ..."
      RETURN
      
AddPers1:
      FOR zz = 4 TO 23: LOCATE zz, 1: PRINT STRING$(79, 32): NEXT zz: COLOR 11
        LOCATE 3, 1
        COLOR 10
        PRINT "Personal Property Account Data"; TAB(50); "Account # "
        COLOR 11
        PRINT " 1- Customer Name .........."
        PRINT " 2- Address 1 .............."
        PRINT " 3- Address 2 .............."
        PRINT " 4- City ..................."
        PRINT " 5- State .................."
        PRINT " 6- Zip Code ..............."
        PRINT " 7- Social Sec # ..........."
        PRINT " 8- Social Sec # ..........."
        PRINT " 9- District ..............."
        PRINT "10- Late File (Y/N) ........"
        PRINT "11- Business (Y/N) ........."
        PRINT "12- Mach Tools Value ......."
        PRINT "13- Mobile Home Value ......"
        PRINT "14- Merch Cap. Value ......."
        PRINT "15- Farm Equip Value ......."
        RETURN
        
        
        
AddPers2:
        FOR zz = 4 TO 23: LOCATE zz, 1: PRINT STRING$(79, 32): NEXT zz: COLOR 11
          LOCATE 3, 1
          COLOR 11
          PRINT "Personal Property Valuation Detail Screen"
          PRINT " 1- Heavy Equipment ........"
          PRINT " 2- Boats .................."
          PRINT " 3- Bus Furniture & Equip .."
          PRINT "  - Vehicle Total .........."
          PRINT "  - Total Personal Property."
          PRINT STRING$(79, "-")
          PRINT " 4- Vehicle VIN ............"
          PRINT " 5- Vehicle Year (YYYY) ...."
          PRINT " 6- Vehicle Make ..........."
          PRINT " 7- Vehicle Model .........."
          PRINT " 8- Vehicle Weight ........."
          PRINT " 9- Vehicle Value .........."
          PRINT "10- Qualifying Vehicle Y/N ."
          PRINT "11- Vehicle License # ......"
          PRINT "12- Type/Delete (I/B/D) ...."
          PRINT "13- Add New Vehicle"
          PRINT "  - Veh Tax Year: "
          RETURN
          
          
PERSBRKDWN:
          FOR zz = 4 TO 23: LOCATE zz, 1: PRINT STRING$(79, 32): NEXT zz: COLOR 11
            LOCATE 3, 1
            
            COLOR 10
            
            RETURN
            
10000 SCREN$ = "1.10.00": CLS : GOSUB 63500
            LOCATE 1, 32: COLOR 11: PRINT "Repair Customer History"; : LOCATE 3, 1
            PRINT "(R)eal or (P)ersonal Account (R,P or ESC-Menu) .. "; : fl = 1
            GOSUB 62210
            IF cf = 1 THEN 10
            IF NOT (in$ = "R" OR in$ = "P") THEN 10000
            taxtype$ = in$
            
10030 LOCATE 4, 1: COLOR 11
            PRINT "Enter Customer Account Number (TAB-Last Name : ESC-Menu) .. "; : fl = -6
            GOSUB 62000
            IF cf = 1 THEN
              CLOSE : GOTO 10
            END IF
            IF cf = 2 THEN
              GOSUB Namesearch
              IF ggo = 1 THEN 10075
              IF ggo = 0 THEN LOCATE 10, 1: PRINT "Account Not Found : Press Any Key to Continue "; : fl = 1: GOSUB 62000: GOTO 10000
            END IF
            Record! = VAL(in$)
10075 CLOSE
              ON ERROR RESUME NEXT

            GOSUB 60000
            GOSUB 61000
            FOR Cnt! = 1 TO LOF(TBTransFile) / LEN(TBTrans(1))
              GET TBTransFile, Cnt!, TBTrans(1)
              LOCATE 15, 1: PRINT "Processing Transaction # "; Cnt!
              BigTRNum# = TBTrans(1).TrCustomerRecord + 1
              IF ERR = 6 THEN
                ERR = 0
                BadTr& = BadTr& + 1
                GOTO OhBadTR
              END IF

              
              IF TBTrans(1).TrCustomerRecord = Record! THEN
                CTR& = CTR& + 1
                'STOP
                IF taxtype$ = "R" AND TBTrans(1).TrSource = "R" THEN
                  IF Record! < 1 OR Record! > LOF(TBFile) / LEN(TBRECust(1)) THEN 10030
                  
                  GET TBFile, Record!, TBRECust(1)
                  TBRECust(1).FirstTrans = 0
                  TBRECust(1).LastTrans = 0
                  PUT TBFile, Record!, TBRECust(1)
                  IF TBRECust(1).FirstTrans = 0 THEN
                    TBRECust(1).FirstTrans = Cnt!
                    TBRECust(1).LastTrans = Cnt!
                    PUT TBFile, Record!, TBRECust(1)
                  ELSE
                    Prev! = TBRECust(1).LastTrans
                    TBRECust(1).LastTrans = Cnt!
                    PUT TBFile, Record!, TBRECust(1)
                    GET TBTranFile, Prev!, TBTrans(1)
                    TBTrans(1).TrNextTrans = Prev!
                    PUT TBTranFile, Prev!, TBTrans(1)
                  END IF
                END IF
                
              END IF
              
              IF taxtype$ = "P" AND TBTrans(1).TrSource = "P" THEN
                IF Record! < 1 OR Record! > LOF(TBFile) / LEN(TBPPCust(1)) THEN 10030
                'LOCK #TBFile, Record!
                GET TBFile, Record!, TBPPCust(1)
                TBPPCust(1).FirstTrans = 0
                TBPPCust(1).LastTrans = 0
                
                IF TBTrans(1).TrCustomerRecord = Record! THEN
                  IF TBPPCust(1).FirstTrans = 0 THEN
                    TBPPCust(1).FirstTrans = Cnt!
                    TBPPCust(1).LastTrans = Cnt!
                    PUT TBFile, Record!, TBPPCust(1)
                  ELSE
                    Prev! = TBPPCust(1).LastTrans
                    TBPPCust(1).LastTrans = Cnt!
                    PUT TBFile, Record!, TBPPCust(1)
                    GET TBTranFile, Prev!, TBTrans(1)
                    TBTrans(1).TrNextTrans = Prev!
                    PUT TBTranFile, Prev!, TBTrans(1)
                  END IF
                END IF
              END IF
              
OhBadTR:
            NEXT Cnt!
            
            CLOSE
            STOP
            GOTO 10000
            
60000  IF taxtype$ = "R" THEN
            CLOSE TBFile
            TBFile = FREEFILE
            OPEN TBPath$ + "TBRTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
          ELSEIF taxtype$ = "P" THEN
            CLOSE TBFile
            TBFile = FREEFILE
            OPEN TBPath$ + "TBPTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
            CLOSE TBVFile
            TBVFile = FREEFILE
            OPEN TBPath$ + "TBTVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBVFile LEN = LEN(TBVeh(1))
          ELSEIF taxtype$ = "S" THEN
            CLOSE TBFile
            TBFile = FREEFILE
            OPEN TBPath$ + "TBSPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
          END IF
60025 CLOSE TBBalFile
          TBBalFile = FREEFILE
          IF taxtype$ = "R" THEN
            OPEN TBPath$ + "TBRTBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBRECustBal(1))
          ELSE
            OPEN TBPath$ + "TBPTBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBPPCustBal(1))
          END IF
60030 RETURN
61000 TBTransFile = FREEFILE
          OPEN TBPath$ + "TBTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBTransFile LEN = LEN(TBTrans(1))
61020 RETURN
          
          
62000 cf = 0: COLOR 14: in$ = "": RO = CSRLIN: PO = POS(X): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
62010 IF INKEY$ <> "" THEN 62010
62020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
62030 z$ = INKEY$: IF z$ = "" THEN 62030
62040 IF z$ <> CHR$(13) THEN 62060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
62050 GOSUB 62200: RETURN
62060 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN cf = 1: RETURN
          IF z$ = CHR$(9) THEN cf = 2: RETURN
          IF LEN(z$) < 2 THEN 62065
          IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 59 THEN cf = 11: RETURN
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
63000 RETURN
63500 COLOR 10: LOCATE 1, 1: PRINT "Sys. Date: "; SYSDATE$; TAB(70); SCREN$: PRINT STRING$(79, "_"): RETURN
          
Namesearch:
          LOCATE 8, 1:
          COLOR 11
          PRINT "Enter Last Name (Partial Accepted : ESC - End) .. "; : fl = 25
          GOSUB 62000
          IF cf = 1 THEN
            ggo = 0
            FOR LLL = 8 TO 13: PRINT STRING$(79, 32): NEXT LLL
              RETURN
            END IF
            nn$ = UCASE$(in$)
            nn = LEN(nn$)
            IF nn > 20 THEN nn = 20
            nn$ = LEFT$(nn$, nn)
            
            IF taxtype$ = "R" THEN
              TBIDXFile = FREEFILE
              OPEN TBPath$ + "TBRTNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
              FIELD TBIDXFile, 20 AS lastnam$, 4 AS Mrec$
              GOSUB 60000
              FOR ll = 1 TO LOF(TBIDXFile) / 24
                GET TBIDXFile, ll
                IF LEFT$(lastnam$, nn) = nn$ THEN GOTO n1
n3:
              NEXT ll
              FOR LLL = 8 TO 13: PRINT STRING$(79, 32): NEXT LLL
                ggo = 0
                RETURN
                
n1:
                GET TBFile, CVS(Mrec$), TBRECust(1)
                LOCATE 11, 1: COLOR 11
                Nme$ = TBRECust(1).CustName
                COLOR 10
                FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
                  LOCATE 11, 1: COLOR 11
                  PRINT "Full Account Name: "; : COLOR 15: PRINT Nme$
                  COLOR 11
                  PRINT "   Account Number: "; : COLOR 15: PRINT CVS(Mrec$)
n2:
                  LOCATE 13, 1: COLOR 11
                  PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 62210
                  IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO n2
                  IF in$ = "Q" THEN
                    FOR LLL = 8 TO 13: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
                      CLOSE
                      ggo = 0
                      RETURN
                    END IF
                    IF in$ = "Y" THEN
                      ggo = 1: Record! = CVS(Mrec$)
                      FOR LLL = 8 TO 13:  LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
                        CLOSE
                        RETURN
                      END IF
                      IF in$ = "N" THEN GOTO n3
                      GOTO n2
                    END IF
                    
                    
                    TBIDXFile = FREEFILE
                    OPEN TBPath$ + "TBPTNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
                    FIELD TBIDXFile, 20 AS lastnam$, 4 AS Mrec$
                    GOSUB 60000
                    FOR ll = 1 TO LOF(TBIDXFile) / 24
                      GET TBIDXFile, ll
                      IF LEFT$(lastnam$, nn) = nn$ THEN GOTO PPn1
PPn3:
                    NEXT ll
                    LOCATE 8, 1: PRINT STRING$(79, 32)
                    ggo = 0
                    RETURN
                    
PPn1:
                    GET TBFile, CVS(Mrec$), TBPPCust(1)
                    LOCATE 11, 1: COLOR 11
                    Nme$ = TBPPCust(1).CustName
                    COLOR 10
                    FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
                      LOCATE 11, 1: COLOR 11
                      PRINT "Full Account Name: "; : COLOR 15: PRINT Nme$
                      COLOR 11
                      PRINT "   Account Number: "; : COLOR 15: PRINT CVS(Mrec$)
PPn2:
                      LOCATE 13, 1: COLOR 11
                      PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 62210
                      IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO PPn2
                      IF in$ = "Q" THEN
                        FOR LLL = 8 TO 13: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
                          CLOSE
                          ggo = 0
                          RETURN
                        END IF
                        IF in$ = "Y" THEN
                          ggo = 1: Record! = CVS(Mrec$)
                          FOR LLL = 8 TO 13: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
                            CLOSE
                            RETURN
                          END IF
                          IF in$ = "N" THEN GOTO PPn3
                          GOTO PPn2
                          
                          
                          
RealNull:
                          FOR Lnt! = 1 TO 10
                            TBRECustBal(1).TaxAmt(Lnt!) = 0
                            TBRECustBal(1).IntAmt(Lnt!) = 0
                            TBRECustBal(1).PenAmt(Lnt!) = 0
                            TBRECustBal(1).OthAmt(Lnt!) = 0
                            TBRECustBal(1).LateAmt(Lnt!) = 0
                            TBRECustBal(1).TaxNotice(Lnt!) = ""
                          NEXT Lnt!
                          RETURN
PersNull:
                          FOR Lnt! = 1 TO 10
                            TBPPCustBal(1).PerPropTaxAmt(Lnt!) = 0
                            TBPPCustBal(1).FarmEquipAmt(Lnt!) = 0
                            TBPPCustBal(1).MachToolsAmt(Lnt!) = 0
                            TBPPCustBal(1).MerchCapAmt(Lnt!) = 0
                            TBPPCustBal(1).MobileHomeAmt(Lnt!) = 0
                            TBPPCustBal(1).LateFeeAmt(Lnt!) = 0
                            TBPPCustBal(1).IntAmt(Lnt!) = 0
                            TBPPCustBal(1).PenAmt(Lnt!) = 0
                            TBPPCustBal(1).TaxNotice(Lnt!) = ""
                          NEXT Lnt!
                          RETURN
                          
                          
FigureBalance:
                          CurBal# = 0: PastDue# = 0
                          IF taxtype$ = "R" THEN
                            CurBal# = TBRECust(1).CurYrTaxAmount + TBRECust(1).CurYrIntAmount + TBRECust(1).CurYrPenAmount + TBRECust(1).CurYrOthAmount
                            CurBal# = INT((CurBal# * 100) + .5) / 100
                            FOR ll = 1 TO 10
                              PastDue# = PastDue# + TBRECustBal(1).TaxAmt(ll) + TBRECustBal(1).IntAmt(ll) + TBRECustBal(1).PenAmt(ll) + TBRECustBal(1).OthAmt(ll) + TBRECustBal(1).LateAmt(ll)
                              PastDue# = INT((PastDue# * 100) + .5) / 100
                            NEXT ll
                          ELSE
                            CurBal# = TBPPCust(1).CurYrPersPropTax + TBPPCust(1).CurYrFarmEquipTax + TBPPCust(1).CurYrMachToolsTax + TBPPCust(1).CurYrMerchCapTax + TBPPCust(1).CurYrMobileHomeTax + TBPPCust(1).CurYrLateFeeTax + TBPPCust(1). _
CurYrIntAmount + TBPPCust(1).CurYrPenAmount
                            CurBal# = INT((CurBal# * 100) + .5) / 100
                            FOR ll = 1 TO 10
                              PastDue# = PastDue# + TBPPCustBal(1).PerPropTaxAmt(ll) + TBPPCustBal(1).FarmEquipAmt(ll) + TBPPCustBal(1).MachToolsAmt(ll) + TBPPCustBal(1).MerchCapAmt(ll) + TBPPCustBal(1).MobileHomeAmt(ll) + TBPPCustBal(1).LateFeeAmt( _
ll) + TBPPCustBal(1).IntAmt(ll) + TBPPCustBal(1).PenAmt(ll)
                              PastDue# = INT((PastDue# * 100) + .5) / 100
                            NEXT ll
                            
                          END IF
                          RETURN
Deleted:
                          PRINT CHR$(7);
                          LOCATE 23, 1: COLOR 15
                          PRINT "This Account Has Been Deleted! Press <ENTER> to Continue"; :
                          fl = 1
                          GOSUB 62000
                          LOCATE 23, 1: PRINT STRING$(79, 32)
                          RETURN
HasBalance:
                          PRINT CHR$(7);
                          LOCATE 23, 1: COLOR 15
                          PRINT "This Account Has a Balance and CANNOT Be Deleted! Press <ENTER> to Continue"; :
                          fl = 1
                          GOSUB 62000
                          LOCATE 23, 1: PRINT STRING$(79, 32)
                          RETURN
                          
                          
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
                          CASE IS = 6
                            RESUME NEXT
                            
                          CASE ELSE
                            CLS
                            LOCATE 15, 1: COLOR 28: PRINT "Error Message: "; : COLOR 15: PRINT "Error "; Number; " Has Occured in Line "; ERL
                            PRINT "Write Down Message and Hit <ESC> Key to exit."
                            fl = 1: GOSUB 62000: IF cf = 1 THEN CLOSE : GOTO 10 ELSE GOSUB 62000
                            
                          END SELECT
                          
                          

