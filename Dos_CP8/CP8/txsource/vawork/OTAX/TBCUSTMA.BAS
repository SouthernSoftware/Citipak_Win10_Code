'$INCLUDE: 'LCTax.Bi'
  DECLARE FUNCTION Num2Date$ (Num%)
  DECLARE FUNCTION Date2Num% (Dat$)
10 CLS
   CLEAR
   ON ERROR GOTO networkerror
           
'  TBPath$ = "F:\"                    ' change to "F:\" Before compiling
   TBPath$ = "C:\qb45\lc\"

   REDIM TBRECust(1) AS TBRECust
   REDIM TBRECustBal(1) AS TBREBAL
   REDIM TBPPCust(1) AS TBPPCust
   REDIM TBPPCustBal(1) AS TBPPBal
   REDIM TBTrans(1) AS TBTrans

   CLOSE : OPEN "R", 1, TBPath$ + "tbyear.dat", 4: FIELD 1, 4 AS year$: GET 1, 1
    CURRENTYEAR$ = year$: CLOSE
   
100 GOSUB 63000: REM get town in town$
110 SCREN$ = "1.00.00": SYSDATE$ = DATE$
     CLOSE : OPEN "R", 1, TBPath$ + "tbyear.dat", 4: FIELD 1, 4 AS year$: GET 1, 1
     CURRENTYEAR$ = year$: CLOSE
120 CLS : GOSUB 63500
130 LOCATE 4, 1: COLOR 11
200 PRINT TAB(31); "Tax Billing System"
210 PRINT TAB(30); "Customer Maintenance "
220 PRINT
230 PRINT TAB(4); "1. Add a New Customer"; TAB(42); "6. Clear Personal Valuations"
240 PRINT
250 PRINT TAB(4); "2. Change/View Customer"; TAB(42); "7. Enter Supplemental Customers"
260 PRINT
270 PRINT TAB(4); "3. Delete a Customer"; TAB(42); "8. Enter Real Estate Assessment's"
280 PRINT
    PRINT TAB(4); "4. Re-Sort the Customer Data "; TAB(42); "9. Roll Assessments into Live Value's"
    PRINT
    PRINT TAB(4); "5. Clear Comment Lines"; TAB(41); "10. Fix Customer History"
    
310 LOCATE 20, 1: COLOR 11: PRINT TAB(26); "Selection No. (ESC-Master Menu) : "; : fl = -2: GOSUB 62000
320 IF cf = 1 THEN CLOSE : RUN "tbmenu"
325 IF LEN(in$) = 0 THEN 310
330 IF VAL(in$) < 1 OR VAL(in$) > 10 THEN 310
340 ON VAL(in$) GOTO 1000, 2000, 3000, 4000, 5000, 6000, 7000, 8000, 9000, 10000
    GOTO 10

1000 SCREN$ = "1.01.00": CLS : COLOR 10: GOSUB 63500: meter$ = ""
1020 LOCATE 1, 32: COLOR 11: PRINT "Add a New Customer"; : LOCATE 3, 1
1030 PRINT "Please Answer the Following ..."
1050 LOCATE 3, 1: COLOR 11
     PRINT "(R)eal Estate or (P)ersonal Property Customer (ESC-End) . "; : fl = 1: GOSUB 62210
     IF cf = 1 THEN GOTO 10
     IF NOT (in$ = "R" OR in$ = "P") THEN 1050
     taxtype$ = UCASE$(in$)

1195 GOSUB 60000

     IF taxtype$ = "R" THEN
      Record! = LOF(TBFile) / LEN(TBRECust(1)) + 1
1100 COLOR 11: LOCATE 3, 60: PRINT Record!
     GOSUB AddReal1
1110 GOSUB 1111
     GOSUB 1112: GOSUB 1113: GOSUB 1114: GOSUB 1115
     GOSUB 1116: GOSUB 1117: GOSUB 1118: GOSUB 1119
     GOSUB 1120: GOSUB 1121: GOSUB 1122: GOSUB 1123
     GOSUB 1124: GOSUB 1125: GOSUB 1126: GOSUB 1127

RESEL:
     LOCATE 23, 1: COLOR 11
     PRINT "Select # (ESC-OK) .. "; : fl = -2
     GOSUB 62000
     IF cf = 1 THEN
      PUT TBFile, Record!, TBRECust(1)
      GOSUB RealNull
      PUT TBBalFile, Record!, TBRECustBal(1)
      GOTO 1200
     END IF
     IF VAL(in$) < 1 OR VAL(in$) > 17 THEN GOTO RESEL
     ON VAL(in$) GOSUB 1111, 1112, 1113, 1114, 1115, 1116, 1117, 1118, 1119, 1120, 1121, 1122, 1123, 1124, 1125, 1126, 1127
     PUT TBFile, Record!, TBRECust(1)
     GOTO RESEL


     ELSE

     REM put personal property here
     Record! = LOF(TBFile) / LEN(TBPPCust(1)) + 1
     COLOR 11: LOCATE 3, 60: PRINT Record!
     GOSUB AddPers1
     GOSUB 1151
     GOSUB 1152: GOSUB 1153: GOSUB 1154: GOSUB 1155
     GOSUB 1156: GOSUB 1157: GOSUB 1158: GOSUB 1159
     GOSUB 1160: GOSUB 1161: GOSUB 1162: GOSUB 1163
     GOSUB 1164: GOSUB 1165: GOSUB 1166

REM RETURN
PPSEL:
     LOCATE 23, 1: COLOR 11
     PRINT "Select # (ESC-OK) .. "; : fl = -2
     GOSUB 62000
     IF cf = 1 THEN
      Record! = LOF(TBFile) / LEN(TBPPCust(1)) + 1
      PUT TBFile, Record!, TBPPCust(1)
      GOSUB PersNull
      PUT TBBalFile, Record!, TBPPCustBal(1)
      CLOSE : GOTO 1000
     END IF
     IF VAL(in$) < 1 OR VAL(in$) > 16 THEN GOTO PPSEL
     ON VAL(in$) GOSUB 1151, 1152, 1153, 1154, 1155, 1156, 1157, 1158, 1159, 1160, 1161, 1162, 1163, 1164, 1165, 1166
     PUT TBFile, Record!, TBPPCust(1)
     GOTO PPSEL
     END IF





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
1162 LOCATE 15, 30: PRINT STRING$(15, 32): LOCATE 15, 30
     fl = 10: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
     TBPPCust(1).PersProp = VAL(in$)
     RETURN
1163 LOCATE 16, 30: PRINT STRING$(15, 32): LOCATE 16, 30
     fl = 10: GOSUB 62000:      IF cf = 1 THEN CLOSE : RUN
     TBPPCust(1).MachTools = VAL(in$)
     RETURN
1164 LOCATE 17, 30: PRINT STRING$(15, 32): LOCATE 17, 30
     fl = 10: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
     TBPPCust(1).MobileHome = VAL(in$)
     RETURN
1165 LOCATE 18, 30: PRINT STRING$(15, 32): LOCATE 18, 30
     fl = 10: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
     TBPPCust(1).MerchCap = VAL(in$)
     RETURN
1166 LOCATE 19, 30: PRINT STRING$(15, 32): LOCATE 19, 30
     fl = 10: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
     TBPPCust(1).FarmEquip = VAL(in$)
     RETURN


1200 REM second screen here
     COLOR 11: LOCATE 3, 1
     PRINT "Customer Account #"; Account$; "  : Parcel # "; parcel$
     GOSUB AddReal2
1205 GOSUB 1211: GOSUB 1212: GOSUB 1213: GOSUB 1214
     GOSUB 1215: GOSUB 1216: GOSUB 1217: GOSUB 1218
     GOSUB 1219: GOSUB 1220
     
1210 LOCATE 22, 1: COLOR 11
     PRINT "Selection # (ESC-OK) .. "; : fl = -2: GOSUB 62000
     IF cf = 1 THEN
      PUT TBFile, Record!, TBRECust(1)
      CLOSE : GOTO 1000
     END IF
     IF VAL(in$) < 1 OR VAL(in$) > 10 THEN 1210
     ON VAL(in$) GOSUB 1211, 1212, 1213, 1214, 1215, 1216, 1217, 1218, 1219, 1220
     GOTO 1210

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

   
   
2000 REM
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
       LOCK #1, Record!
       GET 1, Record!, TBRECust(1)
       IF TBRECust(1).Deleted = "Y" THEN GOSUB Deleted: GOTO 2000
       ELSE
       IF Record! < 1 OR Record! > LOF(TBFile) / LEN(TBPPCust(1)) THEN 2030
       LOCK #1, Record!
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
      LOCATE 15, 30: PRINT USING "########,#"; TBPPCust(1).PersProp
      LOCATE 16, 30: PRINT USING "########,#"; TBPPCust(1).MachTools
      LOCATE 17, 30: PRINT USING "########,#"; TBPPCust(1).MobileHome
      LOCATE 18, 30: PRINT USING "########,#"; TBPPCust(1).MerchCap
      LOCATE 19, 30: PRINT USING "########,#"; TBPPCust(1).FarmEquip

PPselection2:

     LOCATE 23, 1: COLOR 11
     PRINT "Select # (ESC-OK) .. "; : fl = -2
     GOSUB 62000
     IF cf = 1 THEN CLOSE : GOTO 10
     IF VAL(in$) < 1 OR VAL(in$) > 16 THEN GOTO PPselection2
     ON VAL(in$) GOSUB 1151, 1152, 1153, 1154, 1155, 1156, 1157, 1158, 1159, 1160, 1161, 1162, 1163, 1164, 1165, 1166
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


3000 REM
3010 SCREN$ = "1.03.00": CLS : GOSUB 63500
3020 LOCATE 1, 32: COLOR 11: PRINT "Change/View Customer"; : LOCATE 3, 1
     PRINT "(R)eal or (P)ersonal Account (R,P or ESC-Menu) .. "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN 10
     IF NOT (in$ = "R" OR in$ = "P") THEN 3020
     taxtype$ = in$
3030 LOCATE 4, 1: COLOR 11
     PRINT "Enter Customer Account Number (TAB-Last Name : ESC-Menu) .. "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN
      CLOSE : GOTO 10
     END IF
     IF cf = 2 THEN
      GOSUB Namesearch
      IF ggo = 1 THEN 3075
      IF ggo = 0 THEN LOCATE 10, 1: PRINT "Account Not Found : Press Any Key to Continue "; : fl = 1: GOSUB 62000: GOTO 3000
     END IF
     Record! = VAL(in$)
3075 CLOSE : GOSUB 60000
     IF taxtype$ = "R" THEN
       IF Record! < 1 OR Record! > LOF(TBFile) / LEN(TBRECust(1)) THEN 3030
       LOCK #1, Record!
       GET 1, Record!, TBRECust(1)
       IF TBRECust(1).Deleted = "Y" THEN GOSUB Deleted: GOTO 3000
       ELSE
       IF Record! < 1 OR Record! > LOF(TBFile) / LEN(TBPPCust(1)) THEN 3030
       LOCK #1, Record!
       GET 1, Record!, TBPPCust(1)
       IF TBPPCust(1).Deleted = "Y" THEN GOSUB Deleted: GOTO 3000
     END IF
     GOSUB FigureBalance
     IF CurBal# <> 0 OR PastDue# <> 0 THEN
      GOSUB HasBalance: CLOSE : GOTO 3000
     END IF


     IF taxtype$ = "R" THEN
3100 CLS
     SCREN$ = "1.03.00": CLS : GOSUB 63500
      LOCATE 1, 32: COLOR 11: PRINT "Delete Real Estate Customer"; : LOCATE 3, 1

      Account! = Record!
      COLOR 15
      LOCATE 3, 5: PRINT "Account # "; USING "######"; Account!
      LOCATE 4, 5: COLOR 11: PRINT "Customer Name"
      COLOR 15
      LOCATE 5, 5: PRINT TBRECust(1).CustName
      LOCATE 6, 5: PRINT TBRECust(1).Address1
      LOCATE 7, 5: PRINT TBRECust(1).Address2
      LOCATE 8, 5: PRINT TBRECust(1).Address3
3110 LOCATE 10, 5: PRINT "Delete (Y/N) (ESC-Menu) .. "; : fl = 1
      GOSUB 62210
      IF NOT (in$ = "Y" OR in$ = "N") THEN 3110
      IF in$ = "Y" THEN
       TBRECust(1).Deleted = "Y"
       PRINT CHR$(7);
       PUT TBFile, Record!, TBRECust(1)
      END IF
      CLOSE
      GOTO 3000
      ELSE
      REM personal prop here
3200 CLS
     SCREN$ = "1.03.00": CLS : GOSUB 63500
      LOCATE 1, 32: COLOR 11: PRINT "Delete Pers Prop Customer"; : LOCATE 3, 1

      Account! = Record!
      COLOR 15
      LOCATE 3, 5: PRINT "Account # "; USING "######"; Account!
      LOCATE 4, 5: COLOR 11: PRINT "Customer Name"
      COLOR 15
      LOCATE 5, 5: PRINT TBPPCust(1).CustName
      LOCATE 6, 5: PRINT TBPPCust(1).Address1
      LOCATE 7, 5: PRINT TBPPCust(1).Address2
      LOCATE 8, 5: PRINT TBPPCust(1).City; " "; TBPPCust(1).State; " "; TBPPCust(1).Zip
3210 LOCATE 10, 5: PRINT "Delete (Y/N) (ESC-Menu) .. "; : fl = 1
      GOSUB 62210
      IF NOT (in$ = "Y" OR in$ = "N") THEN 3210
      IF in$ = "Y" THEN
       TBPPCust(1).Deleted = "Y"
       PUT TBFile, Record!, TBRECust(1)
      END IF
      CLOSE
      GOTO 3000

     END IF

4000 CLS
        COLOR 11
        PRINT "Do you wish to sort now (Y/N) .. "; : fl = 1
        GOSUB 62210
        IF in$ = "Y" THEN
        RUN "tbNMsort"
        ELSE
        CLOSE : GOTO 10
        END IF

5000 CLS : SCREN$ = "1.05.00": CLS : GOSUB 63500
     LOCATE 1, 32: COLOR 11: PRINT "Clear Comment's"; : LOCATE 3, 1
     COLOR 11
     PRINT "Do you wish to CLEAR the COMMENT LINES NOW? (Y/N) .. "; : fl = 1
     GOSUB 62210
      IF in$ = "Y" THEN
        taxtype$ = "R"
        GOSUB 60000
       FOR Cnt! = 1 TO LOF(TBFile) / LEN(TBRECust(1))
        COLOR 11
        LOCATE 5, 1: PRINT "Account # "; Cnt!; " Now Clearing"
        GET TBFile, Cnt!, TBRECust(1)
        TBRECust(1).Comment1 = ""
        TBRECust(1).Comment2 = ""
        TBRECust(1).Comment3 = ""
        PUT TBFile, Cnt!, TBRECust(1)
       NEXT Cnt!
      END IF
        CLOSE
        GOTO 10

6000 CLS : SCREN$ = "1.06.00": CLS : GOSUB 63500
     LOCATE 1, 32: COLOR 11: PRINT "Clear Personal Values"; : LOCATE 3, 1
     COLOR 11
     PRINT "Do you wish to CLEAR the Personal Values NOW? (Y/N) .. "; : fl = 1
     GOSUB 62210
     IF in$ = "Y" THEN
     PRINT "Are You Sure (Y/N) .. "; : fl = 1
     GOSUB 62210
     ELSE
     GOTO 6000
     END IF
IF in$ = "Y" THEN
 taxtype$ = "P"
 GOSUB 60000
 FOR Cnt! = 1 TO LOF(TBFile) / LEN(TBPPCust(1))
 COLOR 11
 LOCATE 5, 1: PRINT "Account # "; Cnt!; " Now Clearing"
 GET TBFile, Cnt!, TBPPCust(1)
  TBPPCust(1).PersProp = 0
  TBPPCust(1).MachTools = 0
  TBPPCust(1).MobileHome = 0
  TBPPCust(1).MerchCap = 0
  TBPPCust(1).FarmEquip = 0
  TBPPCust(1).LateFile = "N"
  TBPPCust(1).Business = "N"

 PUT TBFile, Cnt!, TBPPCust(1)
 NEXT Cnt!
 END IF
CLOSE
GOTO 10
        
7000 CLS : SCREN$ = "1.07.00": CLS : GOSUB 63500
     LOCATE 1, 32: COLOR 11: PRINT "Enter Supplemental's"; : LOCATE 3, 1

     TBFile = FREEFILE
     OPEN TBPath$ + "TBSPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
      
     Record! = LOF(TBFile) / LEN(TBPPCust(1)) + 1
     COLOR 11: LOCATE 3, 60: PRINT Record!
     GOSUB AddPers1
     GOSUB 1151
     GOSUB 1152: GOSUB 1153: GOSUB 1154: GOSUB 1155
     GOSUB 1156: GOSUB 1157: GOSUB 1158: GOSUB 1159
     GOSUB 1160: GOSUB 1161: GOSUB 1162: GOSUB 1163
     GOSUB 1164: GOSUB 1165: GOSUB 1166


PPSUPSEL:
     LOCATE 23, 1: COLOR 11
     PRINT "Select # (ESC-OK) .. "; : fl = -2
     GOSUB 62000
     IF cf = 1 THEN
      PUT TBFile, Record!, TBPPCust(1)
      CLOSE : GOTO 7000
     END IF
     IF VAL(in$) < 1 OR VAL(in$) > 16 THEN GOTO PPSUPSEL
     ON VAL(in$) GOSUB 1151, 1152, 1153, 1154, 1155, 1156, 1157, 1158, 1159, 1160, 1161, 1162, 1163, 1164, 1165, 1166
     PUT TBFile, Record!, TBPPCust(1)
     GOTO PPSUPSEL
      
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
PRINT "12- Pers Prop Value ........"
PRINT "13- Mach Tools Value ......."
PRINT "14- Mobile Home Value ......"
PRINT "15- Merch Cap. Value ......."
PRINT "16- Farm Equip Value ......."
RETURN
PERSBRKDWN:
FOR zz = 4 TO 23: LOCATE zz, 1: PRINT STRING$(79, 32): NEXT zz: COLOR 11
LOCATE 3, 1

COLOR 10

RETURN

8000 SCREN$ = "1.08.00": CLS : GOSUB 63500
8020 LOCATE 1, 32: COLOR 11: PRINT "Enter Reassessment Values"; : LOCATE 3, 1
     taxtype$ = "R"
8030 LOCATE 4, 1: COLOR 11
     PRINT "Enter Customer Account Number (TAB-Last Name : ESC-Menu) .. "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN
      CLOSE : GOTO 10
     END IF
     IF cf = 2 THEN
      GOSUB Namesearch
      IF ggo = 1 THEN 8075
      IF ggo = 0 THEN LOCATE 15, 1: PRINT "Account Not Found : Press Any Key to Continue "; : fl = 1: GOSUB 62000: GOTO 8000
     END IF
     Record! = VAL(in$)
8075 CLOSE : GOSUB 60000
     LOCATE 3, 1: PRINT STRING$(79, 32)
     IF taxtype$ = "R" THEN
       IF Record! < 1 OR Record! > LOF(TBFile) / LEN(TBRECust(1)) THEN 8030
       LOCK #1, Record!
       GET 1, Record!, TBRECust(1)
       IF TBRECust(1).Deleted = "Y" THEN GOSUB Deleted: GOTO 8000
     END IF
     Account! = Record!
8200 COLOR 11: LOCATE 3, 1
     PRINT STRING$(79, 32): LOCATE 3, 1
     PRINT "Customer Account #"; : COLOR 15: PRINT Account!; TBRECust(1).CustName
     LOCATE 4, 1: PRINT STRING$(79, 32)
     COLOR 11
     LOCATE 5, 50: PRINT "Values Now on File"
     LOCATE 6, 50: PRINT USING "$$######,#.##"; TBRECust(1).RELandValue
     LOCATE 6, 1: PRINT "1- Land Value ....... ";
     LOCATE 6, 30: PRINT STRING$(15, 32): LOCATE 6, 30: fl = 10
     GOSUB 62000: TBRECust(1).RELandValue = VAL(in$)
     IF cf = 1 THEN CLOSE : GOTO 8000
     LOCATE 6, 30: PRINT USING "########,#"; TBRECust(1).RELandValue
     COLOR 11
     LOCATE 7, 50: PRINT USING "$$######,#.##"; TBRECust(1).REBldgValue
     LOCATE 7, 1: PRINT "2- Bldg Value ....... ";
     LOCATE 7, 30: PRINT STRING$(15, 32): LOCATE 7, 30: fl = 10: GOSUB 62000: TBRECust(1).REBldgValue = VAL(in$)
     IF cf = 1 THEN CLOSE : GOTO 8000
     LOCATE 7, 30: PRINT USING "########,#"; TBRECust(1).REBldgValue

8210 LOCATE 10, 1: COLOR 11: PRINT STRING$(79, 32): LOCATE 10, 1
        LOCATE 11, 1: COLOR 11: PRINT STRING$(79, 32): LOCATE 11, 1
        LOCATE 12, 1: COLOR 11: PRINT STRING$(79, 32): LOCATE 12, 1
     PRINT "Selection # (ESC-OK) .. "; : fl = -2: GOSUB 62000
     IF cf = 1 THEN
      IF TBRECust(1).REBldgValue = 0 AND TBRECust(1).RELandValue = 0 THEN

       LOCATE 12, 1
       PRINT "Both Values are Set to Zero : Is This Correct? (Y/N) .. "; : fl = 1
       GOSUB 62210
       IF in$ <> "Y" THEN 8210
      END IF
      PUT TBFile, Record!, TBRECust(1): CLOSE : GOTO 8000
     END IF
     IF VAL(in$) < 1 OR VAL(in$) > 2 THEN 8210
     IF VAL(in$) = 1 THEN
      COLOR 11
      LOCATE 6, 1: PRINT "1- Land Value ....... ";
      LOCATE 6, 30: PRINT STRING$(15, 32): LOCATE 6, 30: fl = 10: GOSUB 62000: TBRECust(1).RELandValue = VAL(in$)
      IF cf = 1 THEN CLOSE : GOTO 8000
      LOCATE 6, 30: PRINT USING "########,#"; TBRECust(1).RELandValue
     END IF
     IF VAL(in$) = 2 THEN
      COLOR 11
      LOCATE 7, 1: PRINT "1- Bldg Value ....... ";
      LOCATE 7, 30: PRINT STRING$(15, 32): LOCATE 7, 30: fl = 10: GOSUB 62000: TBRECust(1).REBldgValue = VAL(in$)
      IF cf = 1 THEN CLOSE : GOTO 8000
      LOCATE 7, 30: PRINT USING "########,#"; TBRECust(1).REBldgValue
     END IF
     GOTO 8210
9000 SCREN$ = "1.09.00": CLS : GOSUB 63500
9020 LOCATE 1, 32: COLOR 11: PRINT "Roll ReAssessMent Values"; : LOCATE 3, 1
     taxtype$ = "R"
     LOCATE 3, 1
     PRINT "This function requires a Password to Process"
     PRINT
     PRINT "Enter the Password Now (ESC-Cancel) .. "; : fl = 6
     GOSUB 62000: IF cf = 1 THEN 10

     Password$ = UCASE$(in$)
     IF Password$ <> "SOSOFT" THEN 9000
     LOCATE 9, 1
     COLOR 11
     PRINT "Now Rolling the Assessment Values into Live Real Valuations"
     GOSUB 60000
     
     FOR Cnt! = 1 TO LOF(TBFile) / LEN(TBRECust(1))
     GET TBFile, Cnt!, TBRECust(1)

     IF TBRECust(1).Deleted <> "Y" THEN
        TBRECust(1).LandValue = TBRECust(1).RELandValue
        TBRECust(1).BldgValue = TBRECust(1).REBldgValue
        TBRECust(1).RELandValue = 0
        TBRECust(1).REBldgValue = 0
        PUT TBFile, Cnt!, TBRECust(1)
     END IF

     NEXT Cnt!


     PRINT CHR$(7);
     COLOR 11
     PRINT "All Done!"
     SLEEP 5
     CLOSE
     GOTO 10

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
     GOSUB 60000
     GOSUB 61000
   FOR Cnt! = 1 TO LOF(TBTransFile) / LEN(TBTrans(1))
       GET TBTransFile, Cnt!, TBTrans(1)
       LOCATE 15, 1: PRINT "Processing Transaction # "; Cnt!
   IF TBTrans(1).TrCustomerRecord = Record! THEN
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
       LOCK #TBFile, Record!
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


     NEXT Cnt!

     CLOSE
     GOTO 10000

60000  IF taxtype$ = "R" THEN
         TBFile = FREEFILE
         OPEN TBPath$ + "TBRECUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
        ELSEIF taxtype$ = "P" THEN
         TBFile = FREEFILE
         OPEN TBPath$ + "TBPPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
        ELSEIF taxtype$ = "S" THEN
         TBFile = FREEFILE
         OPEN TBPath$ + "TBSPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
        END IF
60025 TBBalFile = FREEFILE
      IF taxtype$ = "R" THEN
      OPEN TBPath$ + "TBREBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBRECustBal(1))
      ELSE
      OPEN TBPath$ + "TBPPBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBPPCustBal(1))
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
  OPEN TBPath$ + "TBRENAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
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
nme$ = TBRECust(1).CustName
COLOR 10
FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT nme$
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
  OPEN TBPath$ + "TBPPNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
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
nme$ = TBPPCust(1).CustName
COLOR 10
FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT nme$
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
     CurBal# = TBPPCust(1).CurYrPersPropTax + TBPPCust(1).CurYrFarmEquipTax + TBPPCust(1).CurYrMachToolsTax + TBPPCust(1).CurYrMerchCapTax + TBPPCust(1).CurYrMobileHomeTax + TBPPCust(1).CurYrLateFeeTax + TBPPCust(1).CurYrIntAmount + TBPPCust(1). _
CurYrPenAmount
     CurBal# = INT((CurBal# * 100) + .5) / 100
     FOR ll = 1 TO 10
     PastDue# = PastDue# + TBPPCustBal(1).PerPropTaxAmt(ll) + TBPPCustBal(1).FarmEquipAmt(ll) + TBPPCustBal(1).MachToolsAmt(ll) + TBPPCustBal(1).MerchCapAmt(ll) + TBPPCustBal(1).MobileHomeAmt(ll) + TBPPCustBal(1).LateFeeAmt(ll) + TBPPCustBal(1). _
IntAmt(ll) + TBPPCustBal(1).PenAmt(ll)
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

