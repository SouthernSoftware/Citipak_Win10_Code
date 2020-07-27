
  DECLARE FUNCTION num2date$ (Num%)
  DECLARE FUNCTION Date2Num% (Dat$)
  DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
  DECLARE SUB PrintTitle (Title$)
  DECLARE SUB PrintLoading ()
'$INCLUDE: 'LCTax.Bi'
'$INCLUDE: 'DefCnf.BI'
'$INCLUDE: 'formedit.BI'
'$INCLUDE: 'fieldinf.BI'
'$INCLUDE: 'SetCnf.BI'



10 CLS
   ON ERROR GOTO networkerror

    TBPath$ = "F:\"

    REDIM TBRECust(1) AS TBRECust
    REDIM TBRECustBal(1) AS TBREBAL
    REDIM TBPPCust(1) AS TBPPCust
    REDIM TBPPCustBal(1) AS TBPPBal
    REDIM TBTrans(1) AS TBTrans
    REDIM TBVeh(1) AS PPVehType
    REDIM TBEX(1) AS TBExceptionType

    REDIM TTRECust(1) AS TBRECust
    REDIM TTPPCust(1) AS TBPPCust
    REDIM TTVeh(1) AS PPVehType
    DIM VREC!(250)

100 GOSUB 63000: REM get town in town$
110 SCREN$ = "1.00.00": SYSDATE$ = DATE$
     CLOSE
     OPEN "R", 1, TBPath$ + "tbyear.dat", 4: FIELD 1, 4 AS year$: GET 1, 1
     CURRENTYEAR$ = year$: CLOSE
     TaxType$ = ""
120 CLS : GOSUB 63500
130 LOCATE 4, 1: COLOR 11
200 PRINT TAB(31); "Tax Billing System"
210 PRINT TAB(18); "Customer Maintenance - Commissioner - Supplemental"
220 PRINT
    PRINT
230 PRINT TAB(5); "1. Add a New Customer";
    PRINT TAB(42); "6. Print Real Estate Land Book"
240 PRINT
250 PRINT TAB(5); "2. Change/View Customer";
     PRINT TAB(42); "7. Print Personal Property Book"
260 PRINT
270 PRINT TAB(5); "3. Delete a Customer";
     PRINT TAB(42); "8. Certify Real Estate to Treasurer"
280 PRINT
    PRINT TAB(5); "4. Re-Sort the Customer Data";
     PRINT TAB(42); "9. Certify Pers Property to Treasurer"
    PRINT
    PRINT TAB(5); "5. Print Customer Listings";
     PRINT TAB(41); "10. N/A"
310 LOCATE 19, 1: COLOR 11: PRINT TAB(23); "Selection No. (ESC-Master Menu) : "; : fl = -2: GOSUB 62000
320 IF cf = 1 THEN CLOSE : RUN "tbcmenu"
325 IF LEN(in$) = 0 THEN 310
330 IF VAL(in$) < 1 OR VAL(in$) > 9 THEN 310
340 ON VAL(in$) GOTO 1000, 2000, 3000, 4000, 5000, 6000, 7000, 8000, 9000
    GOTO 10

1000 SCREN$ = "1.01.00": CLS : COLOR 10: GOSUB 63500: meter$ = ""
1020 LOCATE 1, 32: COLOR 11: PRINT "Add a New Customer"; : LOCATE 3, 1
1030 PRINT "Please Answer the Following ..."
1050 LOCATE 3, 1: COLOR 11
     PRINT "(R)eal Estate or (P)ersonal Property Customer (ESC-End) . "; : fl = 1: GOSUB 62210
     IF cf = 1 THEN GOTO 10
     IF NOT (in$ = "R" OR in$ = "P") THEN 1050
     TaxType$ = UCASE$(in$)

1195 GOSUB 60000

     IF TaxType$ = "R" THEN
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
     GOSUB 1160: GOSUB 1161: GOSUB 1163
     GOSUB 1164: GOSUB 1165: GOSUB 1166
      
PPSEL:
     LOCATE 23, 1: COLOR 11
     PRINT "Select # (ESC-OK) .. "; : fl = -2
     GOSUB 62000
     IF cf = 1 THEN
      PUT TBFile, Record!, TBPPCust(1)
      GOSUB PersNull
      PUT TBBalFile, Record!, TBPPCustBal(1)
      GOTO 1300
     END IF
     IF VAL(in$) < 1 OR VAL(in$) > 16 THEN GOTO PPSEL
     ON VAL(in$) GOSUB 1151, 1152, 1153, 1154, 1155, 1156, 1157, 1158, 1159, 1160, 1161, 1163, 1164, 1165, 1166
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
     LOCATE 13, 60: fl = -12: GOSUB 62000
     TBPPCust(1).TBLateFeeAmount = VAL(in$)
     LOCATE 13, 60
     PRINT USING "#####,#.##"; TBPPCust(1).TBLateFeeAmount
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
     fl = 1: GOSUB 62210
     IF NOT (in$ = "Y" OR in$ = "N") THEN 1176
     TBVeh(1).VehQ = UCASE$(in$)
     RETURN
1177 LOCATE 17, 30: PRINT STRING$(15, 32): LOCATE 17, 30
     fl = 10: GOSUB 62000
     TBVeh(1).VehLic = UCASE$(in$)
     RETURN
1178 LOCATE 18, 30: PRINT STRING$(10, 32): LOCATE 18, 30
     fl = 1: GOSUB 62210
     IF NOT (in$ = "I" OR in$ = "B" OR in$ = "D" OR in$ = "T") THEN 1178
     TBVeh(1).VehTyp = in$
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

1300 REM 2nd screen for Personal Property
     COLOR 11: LOCATE 3, 1
     GOSUB AddPers2
     GOSUB 1167: GOSUB 1168: GOSUB 1169

     
1310 LOCATE 23, 1: COLOR 11

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
     PRINT "Select # (ESC-Ok) .. "; : fl = 2
     GOSUB 62000
     IF cf = 1 THEN CLOSE : GOTO 10

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
      GOSUB 1178
      TBVeh(1).VehPad = ""
      TBVeh(1).VehAcct = Record!
      PUT TBVFile, (LOF(TBVFile) \ LEN(TBVeh(1))) + 1, TBVeh(1)
      'Update Pointer Here
      IF TBPPCust(1).FirstVeh = 0 THEN
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
      GOTO 1310
     END IF

     IF (VAL(in$) >= 4 AND VAL(in$) <= 12) AND TBPPCust(1).LastVeh > 0 THEN
      GET TBVFile, TBPPCust(1).LastVeh, TBVeh(1)
     ON (VAL(in$) - 3) GOSUB 1170, 1171, 1172, 1173, 1174, 1175, 1176, 1177, 1178
      PUT TBVFile, TBPPCust(1).LastVeh, TBVeh(1)
     END IF

     GOTO 1310






   
2000 REM
2010 SCREN$ = "1.02.00": CLS : GOSUB 63500
2020 LOCATE 1, 32: COLOR 11: PRINT "Change/View Customer"; : LOCATE 3, 1
     IF TaxType$ = "R" OR TaxType$ = "P" THEN
     PRINT "(R)eal or (P)ersonal Account (R,P or ESC-Menu) .. "; TaxType$
     ELSE
     PRINT "(R)eal or (P)ersonal Account (R,P or ESC-Menu) .. "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN 10
     IF NOT (in$ = "R" OR in$ = "P") THEN 2020
     TaxType$ = in$
     END IF

2030 LOCATE 4, 1: COLOR 11
     PRINT "Enter Acct# (TAB-Last Name : <F1>-VIN  : ESC-Menu) .. "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN
        CLOSE : GOTO 10
     END IF
     IF cf = 2 THEN
      GOSUB Namesearch
      IF ggo = 1 THEN 2075
      IF ggo = 0 THEN LOCATE 15, 1: PRINT "Account Not Found : Press Any Key to Continue "; : fl = 1: GOSUB 62000: GOTO 2000
     END IF
     IF cf = 11 THEN
      GOSUB VINsearch
      IF ggo = 1 THEN 2075
      IF ggo = 0 THEN LOCATE 15, 1: PRINT "Account Not Found : Press Any Key to Continue "; : fl = 1: GOSUB 62000: GOTO 2000
     END IF

     Record! = VAL(in$)
2075  GOSUB 60000
     LOCATE 3, 1: PRINT STRING$(79, 32)
     IF TaxType$ = "R" THEN
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
     IF TaxType$ = "R" THEN
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
      LOCATE 13, 30: PRINT TBPPCust(1).LateFile;
       COLOR 11
       LOCATE 13, 55: PRINT "Late: ";
       COLOR 15:
       PRINT USING "#####,#.##"; TBPPCust(1).TBLateFeeAmount
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

2200 IF TaxType$ = "R" THEN
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
     IF cf = 1 THEN PUT TBFile, Record!, TBRECust(1): CLOSE : GOTO 2000
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
      COLOR 15
      LOCATE 10, 30: PRINT STRING$(30, 32): LOCATE 10, 30: PRINT TBVeh(1).VehVin: COLOR 11: LOCATE 10, 50: PRINT "#"; VREC; " of "; VCNT
      COLOR 15
      LOCATE 11, 30: PRINT STRING$(30, 32): LOCATE 11, 30: PRINT TBVeh(1).VehYear
      LOCATE 12, 30: PRINT STRING$(30, 32): LOCATE 12, 30: PRINT TBVeh(1).VehMake
      LOCATE 13, 30:  PRINT STRING$(30, 32): LOCATE 13, 30: PRINT TBVeh(1).VehModel
      LOCATE 14, 30:  PRINT STRING$(30, 32): LOCATE 14, 30: PRINT TBVeh(1).VehWght
      LOCATE 15, 30: PRINT STRING$(30, 32): LOCATE 15, 30: PRINT USING "######,#"; TBVeh(1).VehValue
      LOCATE 16, 30: PRINT STRING$(30, 32): LOCATE 16, 30: PRINT TBVeh(1).VehQ
      LOCATE 17, 30: PRINT STRING$(30, 32): LOCATE 17, 30: PRINT TBVeh(1).VehLic
      LOCATE 18, 30: PRINT STRING$(30, 32): LOCATE 18, 30: PRINT TBVeh(1).VehTyp
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
     IF cf = 1 THEN CLOSE : GOTO 2000
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
      GOSUB 1178
      TBVeh(1).VehPad = ""
      PUT TBVFile, (LOF(TBVFile) \ LEN(TBVeh(1))) + 1, TBVeh(1)
      VCNT = VCNT + 1
      VREC!(VCNT) = (LOF(TBVFile) \ LEN(TBVeh(1)))
      VREC = VCNT
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
      GOTO PPselection3
     END IF



     IF VAL(in$) >= 1 AND VAL(in$) <= 3 THEN
      ON VAL(in$) GOSUB 1167, 1168, 1169
      PUT TBFile, Record!, TBPPCust(1)
      GOTO PPselection3
     END IF

     IF VAL(in$) >= 4 AND VAL(in$) <= 12 THEN
      GET TBVFile, VREC!(VREC), TBVeh(1)
     ON (VAL(in$) - 3) GOSUB 1170, 1171, 1172, 1173, 1174, 1175, 1176, 1177, 1178
      PUT TBVFile, VREC!(VREC), TBVeh(1)
     END IF

     GOTO PPselection3

3000 REM
3010 SCREN$ = "1.03.00": CLS : GOSUB 63500
3020 LOCATE 1, 32: COLOR 11: PRINT "Change/View Customer"; : LOCATE 3, 1
     PRINT "(R)eal or (P)ersonal Account (R,P or ESC-Menu) .. "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN 10
     IF NOT (in$ = "R" OR in$ = "P") THEN 3020
     TaxType$ = in$
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
     IF TaxType$ = "R" THEN
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

     IF CurBal# <> 0 OR PastDue# <> 0 THEN
      GOSUB HasBalance: CLOSE : GOTO 3000
     END IF


     IF TaxType$ = "R" THEN
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
        RUN "tbSsort"
        ELSE
        CLOSE : GOTO 10
        END IF


5000 RUN "tbslist"
6000 RUN "tbsland"
7000 RUN "tbspers"
8000 SCREN$ = "1.08.00": CLS : GOSUB 63500
     LOCATE 1, 32: COLOR 11: PRINT "Certify Real Estate - Supplemental"; : LOCATE 3, 1
     LOCATE 3, 1
     COLOR 11
     PRINT "Certify Real Estate Supplemental Entries to Treasurer (Y/N) .. "; : fl = 1
     GOSUB 62210
     IF NOT (in$ = "Y" OR in$ = "N") THEN 8000
     IF in$ = "N" THEN
      LOCATE 5, 1
      PRINT CHR$(7);
      COLOR 15
      PRINT "Certification Canceled!"
      SLEEP 3
      CLOSE
      GOTO 10
     END IF



     


ProcessRealCertification:
  TaxType$ = "R"
  GOSUB 60000
  
  LastTempRecord& = LOF(TBFile) / LEN(TBRECust(1))
  IF LastTempRecord& = 0 THEN
        CLOSE
        PRINT CHR$(7);
        COLOR 15
        PRINT "No Supplemental Records to Certify to Treasurer"
        SLEEP 6
        GOTO 10
  END IF

  

  FOR Cnt! = 1 TO LastTempRecord&
  GET TBFile, Cnt!, TBRECust(1)

  IF TBRECust(1).Deleted = "Y" THEN
   ELSE
   PUT TBTFile, Cnt!, TBRECust(1)
   PUT TTBalFile, Cnt!, TBRECustBal(1)
  END IF

  NEXT Cnt!

' All Processed So Now Set All Real to Deleted Status
  FOR Cnt! = 1 TO LastTempRecord&
   GET TBFile, Cnt!, TBRECust(1)
    TBRECust(1).Deleted = "Y"
   PUT TBFile, Cnt!, TBRECust(1)
  NEXT Cnt!
  CLOSE

        PRINT CHR$(7);
        COLOR 15
        PRINT "All Supplemental Reals Records Have Been Certified to Treasurer"
        SLEEP 6
        GOTO 10



9000 SCREN$ = "1.09.00": CLS : GOSUB 63500
     LOCATE 1, 27: COLOR 11: PRINT "Certify Personal - Supplemental";
     LOCATE 3, 1
     COLOR 11
     PRINT "Certify Personal Supplemental Entries to Treasurer (Y/N) .. "; : fl = 1
     GOSUB 62210
     IF NOT (in$ = "Y" OR in$ = "N") THEN 9000
     IF in$ = "N" THEN
      LOCATE 5, 1
      PRINT CHR$(7);
      COLOR 15
      PRINT "Certification Canceled!"
      SLEEP 3
      CLOSE
      GOTO 10
     END IF


                        
ProcessPersonalCertification:
  TaxType$ = "P"
  GOSUB 60000

  LastTempRecord& = LOF(TBFile) / LEN(TBPPCust(1))
  IF LastTempRecord& = 0 THEN
        CLOSE
        PRINT CHR$(7);
        COLOR 15
        PRINT "No Supplemental Records to Certify to Treasurer"
        SLEEP 6
        GOTO 10
  END IF
  

  FOR Cnt! = 1 TO LastTempRecord&
  GET TBFile, Cnt!, TBPPCust(1)
  GET TBBalFile, Cnt!, TBPPCustBal(1)

   IF TBPPCust(1).Deleted = "Y" THEN
    ELSE
    PUT TBTFile, Cnt!, TBPPCust(1)
    
    PUT TTBalFile, Cnt!, TBPPCustBal(1)
   END IF
  NEXT Cnt!


  'Now Do the Vehicles
  'LastPerm& will keep the last record in the already certified sup veh

   LastPerm& = LOF(TBTVFile) / LEN(TTVeh(1))
   LastPerm& = LastPerm& + 1

  'LastTemp& will keep the last record in the new temp file created by treas
  'if LastTemp& is > then LastPerm& then new records have been added and need to be transferred to the perm supp for vehicles
   LastTemp& = LOF(TBVFile) / LEN(TBVeh(1))

   IF LastTemp& >= LastPerm& THEN

   FOR Cnt! = LastPerm& TO LastTemp&
    GET TBVFile, Cnt!, TBVeh(1)
    

    PUT TBTVFile, Cnt!, TBVeh(1)
    
   NEXT Cnt!

   END IF        'End Loop for Additional Vehicles

   'Now Set All Supp's to Deleted
  FOR Cnt! = 1 TO LastTempRecord&
  GET TBFile, Cnt!, TBPPCust(1)
   TBPPCust(1).Deleted = "Y"
   PUT TBFile, Cnt!, TBPPCust(1)
  NEXT Cnt!


  CLOSE          'Close All Files
  PRINT CHR$(7);
  COLOR 15
  PRINT "All Supplemental Pers Records Have Been Certified to Treasurer"
  SLEEP 6
  GOTO 10

      
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
PRINT "12- Type/Delete (I/B/D/T) .."
PRINT "13- Add New Vehicle"
RETURN


PERSBRKDWN:
FOR zz = 4 TO 23: LOCATE zz, 1: PRINT STRING$(79, 32): NEXT zz: COLOR 11
LOCATE 3, 1
COLOR 10
RETURN

60000  IF TaxType$ = "R" THEN
         CLOSE TBFile
         TBFile = FREEFILE
         OPEN TBPath$ + "TBRSCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
        ELSEIF TaxType$ = "P" THEN
         CLOSE TBFile
         TBFile = FREEFILE
         OPEN TBPath$ + "TBPSCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
         CLOSE TBVFile
         TBVFile = FREEFILE
         OPEN TBPath$ + "TBSVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBVFile LEN = LEN(TBVeh(1))
        END IF

      ' Supplemental Permanent File Names

        IF TaxType$ = "R" THEN
         CLOSE TBTFile
         TBTFile = FREEFILE
         OPEN TBPath$ + "TTRSCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBTFile LEN = LEN(TBRECust(1))
        ELSEIF TaxType$ = "P" THEN
         CLOSE TBTFile
         TBTFile = FREEFILE
         OPEN TBPath$ + "TTPSCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBTFile LEN = LEN(TBPPCust(1))
         CLOSE TBTVFile
         TBTVFile = FREEFILE
         OPEN TBPath$ + "TTSVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBTVFile LEN = LEN(TTVeh(1))
        END IF

60025 CLOSE TBBalFile
      TBBalFile = FREEFILE
      IF TaxType$ = "R" THEN
       OPEN TBPath$ + "TBSRBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBRECustBal(1))
      ELSE
       OPEN TBPath$ + "TBSPBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBPPCustBal(1))
      END IF

      CLOSE TTBalFile
      TTBalFile = FREEFILE
      IF TaxType$ = "R" THEN
       OPEN TBPath$ + "TTSRBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TTBalFile LEN = LEN(TBRECustBal(1))
      ELSE
       OPEN TBPath$ + "TTSPBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TTBalFile LEN = LEN(TBPPCustBal(1))
      END IF

60030 RETURN
        

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

 IF TaxType$ = "R" THEN
  TBIDXFILE = 14
  OPEN TBPath$ + "TBSRNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFILE LEN = 24
  FIELD TBIDXFILE, 20 AS lastnam$, 4 AS Mrec$
  GOSUB 60000
  FOR ll = 1 TO LOF(TBIDXFILE) / 24
  GET TBIDXFILE, ll
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

  IF TaxType$ = "P" THEN
  CLOSE
  TBIDXFILE = 14
  OPEN TBPath$ + "TBSPNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFILE LEN = 24
  FIELD TBIDXFILE, 20 AS lastnam$, 4 AS Mrec$
  GOSUB 60000
  FOR ll = 1 TO LOF(TBIDXFILE) / 24
  GET TBIDXFILE, ll
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
END IF

VINsearch:
 LOCATE 8, 1:
 COLOR 11
 PRINT "Enter Vehicle VIN (Partial Accepted : ESC - End) .. "; : fl = 22
 GOSUB 62000
 IF cf = 1 THEN
  ggo = 0
  FOR LLL = 8 TO 13: PRINT STRING$(79, 32): NEXT LLL
  RETURN
 END IF
 VIN$ = UCASE$(in$)
 VinLgth = LEN(VIN$)
 GOSUB 60000
 FOR Snt& = 1 TO LOF(TBVFile) \ (LEN(TBVeh(1)))
 GET TBVFile, Snt&, TBVeh(1)
  
  IF VIN$ = RTRIM$(LEFT$(TBVeh(1).VehVin, VinLgth)) THEN
   Record! = TBVeh(1).VehAcct
   IF Record! > 0 THEN GOTO VPP1
  END IF
Vpp4:
 NEXT Snt&
 CLOSE TBVFile
 ggo = 0
 RETURN


VPP1:
GET TBFile, Record!, TBPPCust(1)
LOCATE 11, 1: COLOR 11
Nme$ = TBPPCust(1).CustName
COLOR 10
FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT Nme$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT Record!

VPP2:
LOCATE 13, 1: COLOR 11
PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 62210
IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO VPP2
IF in$ = "Q" THEN
  FOR LLL = 8 TO 13: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  CLOSE
  ggo = 0
  RETURN
END IF
IF in$ = "Y" THEN
  ggo = 1
  FOR LLL = 8 TO 13: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  CLOSE
  RETURN
END IF
IF in$ = "N" THEN GOTO Vpp4
GOTO VPP2



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
     IF TaxType$ = "R" THEN
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

