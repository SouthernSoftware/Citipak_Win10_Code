
  DECLARE FUNCTION num2date$ (Num%)
  DECLARE FUNCTION Date2Num% (Dat$)
  DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
  DECLARE SUB PrintTitle (Title$)
  DECLARE SUB PrintLoading ()
'$INCLUDE: 'LCTax.Bi'
'$INCLUDE: 'DefCnf.BI'
'$INCLUDE: 'formedit.BI'
'$INCLUDE: 'fieldinf.BI'
'$INCLUDE: 'Qscr.BI'
'$INCLUDE: 'SetCnf.BI'

10 CLS
'   ON ERROR GOTO networkerror
   TBPath$ = "F:\"                    ' change to "F:\" Before compiling
    REDIM TBRECust(1) AS TBRECust
    REDIM TBRECustBal(1) AS TBREBAL
    REDIM TBPPCust(1) AS TBPPCust
    REDIM TBPPCustBal(1) AS TBPPBal
    REDIM TBTrans(1) AS TBTrans
    REDIM TBVeh(1) AS PPVehType
    REDIM TBEX(1) AS TBExceptionType
    REDIM TBBoat(1) AS PPBoat
    REDIM TBMH(1) AS PPMHome
    DIM VREC!(250)
100 GOSUB 63000: REM get town in town$
110 SCREN$ = "1.00.00": SYSDATE$ = DATE$
     CLOSE
     OPEN "R", 1, TBPath$ + "tbyear.dat", 4: FIELD 1, 4 AS Year$: GET 1, 1
     CURRENTYEAR$ = Year$: CLOSE
     taxtype$ = ""
120 CLS : GOSUB 63500
130 LOCATE 4, 1: COLOR 11
200 PRINT TAB(31); "Tax Billing System"
210 PRINT TAB(23); "Customer Maintenance - Commissioner "
220 PRINT
    PRINT
230 PRINT TAB(4); "1. Add a New Customer"; TAB(42); "7. Duplicate VIN/Zero Value Review"
240 PRINT
250 PRINT TAB(4); "2. Change/View Customer"; TAB(42); "8. Enter Real Estate Assessment's"
260 PRINT
270 PRINT TAB(4); "3. Delete a Customer"; TAB(42); "9. Roll Assessments into Live Value's"
280 PRINT
    PRINT TAB(4); "4. Re-Sort the Customer Data "; TAB(41); "10. Update Using DMV Data"
    PRINT
    PRINT TAB(4); "5. Clear Comment/Reset High Mileage"; TAB(41); "11. Prepare NADA File"
    PRINT
    PRINT TAB(4); "6. Clear Personal Valuations"; TAB(41); "12. Update Using NADA File"
310 LOCATE 21, 1: COLOR 11: PRINT TAB(23); "Selection No. (ESC-Master Menu) : "; : FL = -2: GOSUB 62000
320 IF cf = 1 THEN CLOSE : RUN "tbcmenu"
325 IF LEN(in$) = 0 THEN 310
330 IF VAL(in$) < 1 OR VAL(in$) > 12 THEN 310
340 ON VAL(in$) GOTO 1000, 2000, 3000, 4000, 5000, 6000, 7000, 8000, 9000, 10000, 11000, 12000
    GOTO 10

1000 SCREN$ = "1.01.00": CLS : COLOR 10: GOSUB 63500: meter$ = ""
1020 LOCATE 1, 32: COLOR 11: PRINT "Add a New Customer"; : LOCATE 3, 1
1030 PRINT "Please Answer the Following ..."
1050 LOCATE 3, 1: COLOR 11
     PRINT "(R)eal Estate or (P)ersonal Property Customer (ESC-End) . "; : FL = 1: GOSUB 62210
     IF cf = 1 THEN GOTO 10
     IF NOT (in$ = "R" OR in$ = "P") THEN 1050
     taxtype$ = UCASE$(in$)

1095 GOSUB 60000

     IF taxtype$ = "R" THEN
      Record! = LOF(tbfile) / LEN(TBRECust(1)) + 1
1100 COLOR 11: LOCATE 3, 60: PRINT Record!
     GOSUB AddReal1
1110 GOSUB 1111
     GOSUB 1112: GOSUB 1113: GOSUB 1114: GOSUB 1115
     GOSUB 1116: GOSUB 1117: GOSUB 1118: GOSUB 1119
     GOSUB 1120: GOSUB 1121: GOSUB 1122: GOSUB 1123
     GOSUB 1124: GOSUB 1125: GOSUB 1126: GOSUB 1127
     GOSUB 1128


RESEL:
     LOCATE 23, 1: COLOR 11
     PRINT "Select # (ESC-OK) .. "; : FL = -2
     GOSUB 62000
     IF cf = 1 THEN
      PUT tbfile, Record!, TBRECust(1)
      GOSUB RealNull
      PUT TBBalFile, Record!, TBRECustBal(1)
      GOTO 1200
     END IF
     IF VAL(in$) < 1 OR VAL(in$) > 18 THEN GOTO RESEL
     ON VAL(in$) GOSUB 1111, 1112, 1113, 1114, 1115, 1116, 1117, 1118, 1119, 1120, 1121, 1122, 1123, 1124, 1125, 1126, 1127, 1128
     PUT tbfile, Record!, TBRECust(1)
     GOTO RESEL


     ELSE

     REM put personal property here
     Record! = LOF(tbfile) / LEN(TBPPCust(1)) + 1
     COLOR 11: LOCATE 3, 60: PRINT Record!
     GOSUB AddPers1
     GOSUB 1151
     GOSUB 1152: GOSUB 1153: GOSUB 1154: GOSUB 1155
     GOSUB 1156: GOSUB 1157: GOSUB 1158: GOSUB 1159
     GOSUB 1160: GOSUB 1161: GOSUB 1163
     'Skip 1164 old mobile home input
     GOSUB 1165: GOSUB 1166
      
PPSEL:
     LOCATE 23, 1: COLOR 11
     PRINT "Select # (ESC-OK) .. "; : FL = -2
     GOSUB 62000
     IF cf = 1 THEN
      PUT tbfile, Record!, TBPPCust(1)
      GOSUB PersNull
      PUT TBBalFile, Record!, TBPPCustBal(1)
      GOTO 1300
     END IF
     IF VAL(in$) < 1 OR VAL(in$) > 16 THEN GOTO PPSEL
     ON VAL(in$) GOSUB 1151, 1152, 1153, 1154, 1155, 1156, 1157, 1158, 1159, 1160, 1161, 1163, 1164, 1165, 1166
     PUT tbfile, Record!, TBPPCust(1)
     GOTO PPSEL
     END IF





1111 LOCATE 4, 28: FL = 3: GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN
     IF LEN(in$) < 3 THEN in$ = in$ + STRING$(3 - LEN(in$), 32)
     TBRECust(1).MapNumber = LEFT$(UCASE$(in$), 3)
       LOCATE 4, 32: FL = 2: GOSUB 62000
       IF cf = 1 THEN CLOSE : RUN
       IF LEN(in$) < 2 THEN in$ = in$ + STRING$(2 - LEN(in$), 32)
       PUT tbfile, Record!, TBRECust(1)
       GET tbfile, Record!, TBRECust(1)
       TBRECust(1).MapNumber = LEFT$(TBRECust(1).MapNumber, 3) + LEFT$(UCASE$(in$), 2) + RIGHT$(TBRECust(1).MapNumber, 9)

        LOCATE 4, 35: FL = 2: GOSUB 62000
        IF cf = 1 THEN CLOSE : RUN
         IF LEN(in$) < 2 THEN in$ = in$ + STRING$(2 - LEN(in$), 32)
         PUT tbfile, Record!, TBRECust(1)
         GET tbfile, Record!, TBRECust(1)
         TBRECust(1).MapNumber = LEFT$(TBRECust(1).MapNumber, 5) + LEFT$(UCASE$(in$), 2) + RIGHT$(TBRECust(1).MapNumber, 7)

         LOCATE 4, 38: FL = 2: GOSUB 62000
         IF cf = 1 THEN CLOSE : RUN
         IF LEN(in$) < 2 THEN in$ = in$ + STRING$(2 - LEN(in$), 32)
         PUT tbfile, Record!, TBRECust(1)
         GET tbfile, Record!, TBRECust(1)
         TBRECust(1).MapNumber = LEFT$(TBRECust(1).MapNumber, 7) + LEFT$(UCASE$(in$), 2) + RIGHT$(TBRECust(1).MapNumber, 5)

         LOCATE 4, 41: FL = 4: GOSUB 62000
         IF cf = 1 THEN CLOSE : RUN
         IF LEN(in$) < 4 THEN in$ = in$ + STRING$(4 - LEN(in$), 32)
         PUT tbfile, Record!, TBRECust(1)
         GET tbfile, Record!, TBRECust(1)
         TBRECust(1).MapNumber = LEFT$(TBRECust(1).MapNumber, 9) + LEFT$(UCASE$(in$), 4) + RIGHT$(TBRECust(1).MapNumber, 1)

          LOCATE 4, 46: FL = 1: GOSUB 62000
          IF cf = 1 THEN CLOSE : RUN
          IF LEN(in$) < 1 THEN in$ = in$ + STRING$(1 - LEN(in$), 32)
          PUT tbfile, Record!, TBRECust(1)
          GET tbfile, Record!, TBRECust(1)
          TBRECust(1).MapNumber = LEFT$(TBRECust(1).MapNumber, 13) + LEFT$(UCASE$(in$), 1)
          PUT tbfile, Record!, TBRECust(1)
          RETURN

1112 LOCATE 5, 28: FL = 35: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
        TBRECust(1).CustName = UCASE$(in$): RETURN
1113 LOCATE 6, 28: FL = 35: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
        TBRECust(1).Address1 = UCASE$(in$): RETURN
1114 LOCATE 7, 28: FL = 35: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
        TBRECust(1).Address2 = UCASE$(in$): RETURN
1115 LOCATE 8, 28: FL = 35: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
        TBRECust(1).Address3 = UCASE$(in$): RETURN
1116 LOCATE 9, 28: FL = 5: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
        TBRECust(1).Zipcode = UCASE$(in$): RETURN
1117 LOCATE 10, 28: FL = 25: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
        TBRECust(1).PropDesc1 = UCASE$(in$): RETURN
1118 LOCATE 11, 28: FL = 25: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
        TBRECust(1).PropDesc2 = UCASE$(in$): RETURN
1119 LOCATE 12, 28: FL = 25: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
        TBRECust(1).PropDesc3 = UCASE$(in$): RETURN
1120 LOCATE 13, 28: FL = 25: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
        TBRECust(1).PropDesc4 = UCASE$(in$): RETURN
1121 LOCATE 14, 28: FL = 2: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
        TBRECust(1).Zone = UCASE$(in$): RETURN
1122 LOCATE 15, 28: FL = 2: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
        TBRECust(1).ClassCode = UCASE$(in$): RETURN
1123 LOCATE 16, 28: FL = 2: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
        TBRECust(1).District = UCASE$(in$): RETURN
1124 LOCATE 17, 28: FL = 12: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
        TBRECust(1).Acres = in$: RETURN
1125 LOCATE 18, 28: FL = 20: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
        TBRECust(1).Comment1 = UCASE$(in$): RETURN
1126 LOCATE 19, 28: FL = 20: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
        TBRECust(1).Comment2 = UCASE$(in$): RETURN
1127 LOCATE 20, 28: FL = 20: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
        TBRECust(1).Comment3 = UCASE$(in$): RETURN
1128 LOCATE 21, 28: FL = 25: GOSUB 62000
        TBRECust(1).NewCom2 = in$
        RETURN


1151 LOCATE 4, 30: FL = 25: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
        TBPPCust(1).CustName = UCASE$(in$): RETURN
1152 LOCATE 5, 30: FL = 25: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
        TBPPCust(1).Address1 = UCASE$(in$): RETURN
1153 LOCATE 6, 30: FL = 25: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
        TBPPCust(1).Address2 = UCASE$(in$): RETURN
1154 LOCATE 7, 30: FL = 20: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
        TBPPCust(1).City = UCASE$(in$): RETURN
1155 LOCATE 8, 30: FL = 2: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
        TBPPCust(1).State = UCASE$(in$): RETURN
1156 LOCATE 9, 30: FL = 10: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
        TBPPCust(1).Zip = UCASE$(in$): RETURN
1157 LOCATE 10, 30: FL = 12: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
        TBPPCust(1).SocSec = UCASE$(in$): RETURN
1158 LOCATE 11, 30: FL = 12: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
        TBPPCust(1).SocSec2 = UCASE$(in$): RETURN
1159 LOCATE 12, 30: FL = 2: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
        TBPPCust(1).District = UCASE$(in$): RETURN
1160 LOCATE 13, 30: FL = 1: GOSUB 62210
        IF cf = 1 THEN CLOSE : RUN
     IF NOT (in$ = "Y" OR in$ = "N") THEN 1160
     TBPPCust(1).LateFile = UCASE$(in$)
     RETURN
1161 LOCATE 14, 30: FL = 1: GOSUB 62210: IF cf = 1 THEN CLOSE : RUN
     IF NOT (in$ = "Y" OR in$ = "N") THEN 1161
     TBPPCust(1).Business = UCASE$(in$)
     RETURN
1163 LOCATE 15, 30: PRINT STRING$(15, 32): LOCATE 15, 30
     FL = 10: GOSUB 62000:      IF cf = 1 THEN CLOSE : RUN
     TBPPCust(1).MachTools = VAL(in$)
     RETURN
1164 'LOCATE 16, 30: PRINT STRING$(15, 32): LOCATE 16, 30
     'fl = 10: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
     'TBPPCust(1).MobileHome = VAL(IN$)
     RETURN
1165 LOCATE 17, 30: PRINT STRING$(15, 32): LOCATE 17, 30
     FL = 10: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
     TBPPCust(1).MerchCap = VAL(in$)
     RETURN
1166 LOCATE 18, 30: PRINT STRING$(15, 32): LOCATE 18, 30
     FL = 10: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
     TBPPCust(1).FarmEquip = VAL(in$)
     RETURN
1167 LOCATE 4, 30: PRINT STRING$(15, 32): LOCATE 4, 30
     FL = 10: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
     TBPPCust(1).HeavyEquip = VAL(in$)
     RETURN
1168 LOCATE 5, 30: PRINT STRING$(15, 32): LOCATE 5, 30
     FL = 10: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
     TBPPCust(1).Boats = VAL(in$)
     RETURN
1169 LOCATE 6, 30: PRINT STRING$(15, 32): LOCATE 6, 30
     FL = 10: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN
     TBPPCust(1).BusFurn = VAL(in$)
     RETURN
1170 LOCATE 10, 30: PRINT STRING$(25, 32): LOCATE 10, 30
     FL = 22: GOSUB 62000
     TBVeh(1).VehVin = UCASE$(in$)
     RETURN
1171 LOCATE 11, 30: PRINT STRING$(25, 32): LOCATE 11, 30
     FL = 4: GOSUB 62000
     TBVeh(1).VehYear = UCASE$(in$)
     RETURN
1172 LOCATE 12, 30: PRINT STRING$(25, 32): LOCATE 12, 30
     FL = 20: GOSUB 62000
     TBVeh(1).VehMake = UCASE$(in$)
     RETURN
1173 LOCATE 13, 30: PRINT STRING$(25, 32): LOCATE 13, 30
     FL = 15: GOSUB 62000
     TBVeh(1).VehModel = UCASE$(in$)
     RETURN
1174 LOCATE 14, 30: PRINT STRING$(15, 32): LOCATE 14, 30
     FL = 15: GOSUB 62000
     TBVeh(1).VehWght = VAL(in$)
     RETURN
1175 LOCATE 15, 30: PRINT STRING$(15, 32): LOCATE 15, 30
     FL = -10: GOSUB 62000
     TBVeh(1).VehValue = VAL(in$)
     RETURN
1176 LOCATE 16, 30: PRINT STRING$(15, 32): LOCATE 16, 30
     FL = 1: GOSUB 62210
     IF NOT (in$ = "Y" OR in$ = "N") THEN 1176
     TBVeh(1).VehQ = UCASE$(in$)
     RETURN
1177 LOCATE 17, 30: PRINT STRING$(15, 32): LOCATE 17, 30
     FL = 10: GOSUB 62000
     TBVeh(1).VehLic = UCASE$(in$)
     RETURN
1178 LOCATE 18, 30: PRINT STRING$(10, 32): LOCATE 18, 30
     FL = 1: GOSUB 62210
     IF NOT (in$ = "I" OR in$ = "B" OR in$ = "D" OR in$ = "T" OR in$ = "M") THEN 1178
     TBVeh(1).VehTyp = in$
     RETURN
1179 'trailer code
     LOCATE 19, 30: PRINT STRING$(10, 32): LOCATE 19, 30
     FL = 1: GOSUB 62210
     IF NOT (in$ = "U" OR in$ = "B" OR in$ = "H" OR in$ = "R" OR in$ = "") THEN 1179
     TBVeh(1).VehTrail = in$
     RETURN

1180 'high mileage flag
     LOCATE 20, 30: PRINT STRING$(10, 32): LOCATE 20, 30
     FL = 1: GOSUB 62210
     IF NOT (in$ = "Y" OR in$ = "N" OR in$ = " ") THEN 1180
     TBVeh(1).VehMiles = in$
     RETURN


'START BOATS HERE'

1181 LOCATE 7, 30: PRINT STRING$(10, 32): LOCATE 7, 30
     FL = 4
     GOSUB 62000
     TBBoat(1).BoatYear = in$
     RETURN

1182 LOCATE 8, 30: PRINT STRING$(15, 32): LOCATE 8, 30
     FL = 15
     GOSUB 62000
     TBBoat(1).BoatMake = in$
     RETURN
1183 LOCATE 9, 30: PRINT STRING$(15, 32): LOCATE 9, 30
     FL = 11
     GOSUB 62000
     TBBoat(1).BoatHull = in$
     RETURN
1184 LOCATE 10, 30: PRINT STRING$(15, 32): LOCATE 10, 30
     FL = 6
     GOSUB 62000
     TBBoat(1).BoatLgth = in$
     RETURN
1185 LOCATE 11, 30: PRINT STRING$(15, 32): LOCATE 11, 30
     FL = 8
     GOSUB 62000
     TBBoat(1).BoatDate = in$
     RETURN

1186 LOCATE 12, 30: PRINT STRING$(15, 32): LOCATE 12, 30
     FL = -8
     GOSUB 62000
     TBBoat(1).BoatCost = VAL(in$)
     RETURN

1187 LOCATE 13, 30: PRINT STRING$(15, 32): LOCATE 13, 30
     FL = 15
     GOSUB 62000
     TBBoat(1).BoatMotorMake = in$
     RETURN
1188 LOCATE 14, 30: PRINT STRING$(15, 32): LOCATE 14, 30
     FL = 6
     GOSUB 62000
     TBBoat(1).BoatMotorHP = in$
     RETURN

1189 LOCATE 15, 30: PRINT STRING$(15, 32): LOCATE 15, 30
     FL = 8
     GOSUB 62000
     TBBoat(1).BoatMotorDate = in$
     RETURN
1190 LOCATE 16, 30: PRINT STRING$(15, 32): LOCATE 16, 30
     FL = -8
     GOSUB 62000
     TBBoat(1).BoatMotorCost = VAL(in$)
     RETURN
1191 LOCATE 17, 30: PRINT STRING$(15, 32): LOCATE 17, 30
     FL = 1
     GOSUB 62210
     TBBoat(1).BoatActive = in$
     RETURN


'Start M Homes HERE'

1192 LOCATE 7, 30: PRINT STRING$(30, 32): LOCATE 7, 30
     FL = 23
     GOSUB 62000
     TBMH(1).MHLandOwner = in$
     RETURN

1193 LOCATE 8, 30: PRINT STRING$(30, 32): LOCATE 8, 30
     FL = 4
     GOSUB 62000
     TBMH(1).MHYear = in$
     RETURN
1194 LOCATE 9, 30: PRINT STRING$(30, 32): LOCATE 9, 30
     FL = 15
     GOSUB 62000
     TBMH(1).MHMake = in$
     RETURN
1195 LOCATE 10, 30: PRINT STRING$(10, 32): LOCATE 10, 30
     FL = 6
     GOSUB 62000
     TBMH(1).MHLgth = in$
     RETURN
1196 LOCATE 11, 30: PRINT STRING$(10, 32): LOCATE 11, 30
     FL = 6
     GOSUB 62000
     TBMH(1).MHWdth = in$
     RETURN
1197 LOCATE 12, 30: PRINT STRING$(10, 32): LOCATE 12, 30
     FL = 8
     GOSUB 62000
     TBMH(1).MHDate = in$
     RETURN
1198 LOCATE 13, 30: PRINT STRING$(10, 32): LOCATE 10, 30
     FL = -8
     GOSUB 62000
     TBMH(1).MHCost = VAL(in$)
     RETURN
1199 LOCATE 14, 30: PRINT STRING$(10, 32): LOCATE 14, 30
     FL = 1
     GOSUB 62000
     TBMH(1).MHActive = in$
     RETURN

      


1200 REM second screen here
     COLOR 11: LOCATE 3, 1
     PRINT "Customer Account #"; Account$; "  : Parcel # "; parcel$
     GOSUB AddReal2
1205 GOSUB 1211: GOSUB 1212: GOSUB 1213: GOSUB 1214
     GOSUB 1215: GOSUB 1216: GOSUB 1217: GOSUB 1218
     GOSUB 1219: GOSUB 1220
     
1210 LOCATE 22, 1: COLOR 11
     PRINT "Selection # (ESC-OK) .. "; : FL = -2: GOSUB 62000
     IF cf = 1 THEN
      TBRECust(1).LastDateChange = Date2Num(DATE$)
      PUT tbfile, Record!, TBRECust(1)
      CLOSE : GOTO 1000
     END IF
     IF VAL(in$) < 1 OR VAL(in$) > 10 THEN 1210
     ON VAL(in$) GOSUB 1211, 1212, 1213, 1214, 1215, 1216, 1217, 1218, 1219, 1220
     GOTO 1210

1211 LOCATE 6, 30: PRINT STRING$(15, 32): LOCATE 6, 30: FL = 10: GOSUB 62000: TBRECust(1).LandValue = VAL(in$): RETURN
1212 LOCATE 7, 30: PRINT STRING$(15, 32): LOCATE 7, 30: FL = 10: GOSUB 62000: TBRECust(1).BldgValue = VAL(in$): RETURN
1213 LOCATE 8, 30: PRINT STRING$(15, 32): LOCATE 8, 30: FL = 10: GOSUB 62000: TBRECust(1).LandUseValue = VAL(in$): RETURN
1214 LOCATE 9, 30: FL = 1: GOSUB 62210
        IF NOT (in$ = "Y" OR in$ = "N") THEN 1214
        TBRECust(1).NewConst = in$
        RETURN
1215 LOCATE 10, 30: FL = 1: GOSUB 62210
        IF NOT (in$ = "Y" OR in$ = "N") THEN 1215
        TBRECust(1).NewValue = in$
        RETURN
1216 LOCATE 11, 30: FL = 1: GOSUB 62210
        IF NOT (in$ = "Y" OR in$ = "N") THEN 1216
        TBRECust(1).NewAddress = in$
        RETURN
1217 LOCATE 12, 30: FL = 2: GOSUB 62000
        TBRECust(1).TransferYr = in$
        RETURN
1218 LOCATE 13, 30: FL = -10: GOSUB 62000
        TBRECust(1).NewBldgValue = VAL(in$)
        RETURN
1219 LOCATE 14, 30: FL = 25: GOSUB 62000
        TBRECust(1).NewCom1 = in$
        RETURN
1220 LOCATE 15, 30: FL = 1: GOSUB 62210
        TBRECust(1).MortgageYN = in$
        RETURN

1300 REM 2nd screen for Personal Property
     COLOR 11: LOCATE 3, 1
     GOSUB AddPers2
     GOSUB 1167
     'GOSUB 1168  DO NOT USED SINCE BOATS NOW DETAILED OUT
     GOSUB 1169
  


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
     PRINT "Select # (ESC-Ok) .. "; : FL = 2
     GOSUB 62000
     IF cf = 1 THEN
       BoatRecord! = 0
       GOTO 1400
     END IF

     IF VAL(in$) = 15 THEN
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
        PRINT "13- Trailer Type (U,B,H,R)..  "
        PRINT "14- High Mileage Flag (Y) .. "
        PRINT "15- Add New Vehicle"

      GOSUB 1170
      GOSUB 1171
      GOSUB 1172
      GOSUB 1173
      GOSUB 1174
      GOSUB 1175
      GOSUB 1176
      GOSUB 1177
      GOSUB 1178
      GOSUB 1179
      GOSUB 1180

      TBVeh(1).VehAcct = Record!
      TBVeh(1).VehPad = ""
      PUT TBVFile, (LOF(TBVFile) \ LEN(TBVeh(1))) + 1, TBVeh(1)
      'Update Pointer Here
      IF TBPPCust(1).FirstVeh = 0 THEN
       TBPPCust(1).FirstVeh = (LOF(TBVFile) \ LEN(TBVeh(1)))
       TBPPCust(1).LastVeh = (LOF(TBVFile) \ LEN(TBVeh(1)))
       PUT tbfile, Record!, TBPPCust(1)
      ELSE
      PrevVeh = TBPPCust(1).LastVeh
       TBPPCust(1).LastVeh = (LOF(TBVFile) \ LEN(TBVeh(1)))
      PUT tbfile, Record!, TBPPCust(1)
      GET TBVFile, PrevVeh, TBVeh(1)
       TBVeh(1).VehNext = (LOF(TBVFile) \ LEN(TBVeh(1)))
      PUT TBVFile, PrevVeh, TBVeh(1)
      END IF

     END IF

     IF VAL(in$) >= 1 AND VAL(in$) <= 3 THEN
      ON VAL(in$) GOSUB 1167, 1168, 1169
      PUT tbfile, Record!, TBPPCust(1)
      GOTO 1310
     END IF

     IF (VAL(in$) >= 4 AND VAL(in$) <= 14) AND TBPPCust(1).LastVeh > 0 THEN
      GET TBVFile, TBPPCust(1).LastVeh, TBVeh(1)
     ON (VAL(in$) - 3) GOSUB 1170, 1171, 1172, 1173, 1174, 1175, 1176, 1177, 1178, 1179, 1180
      PUT TBVFile, TBPPCust(1).LastVeh, TBVeh(1)
     END IF

     GOTO 1310



1400 REM BOAT SCREEN FOR PERS PROPERTY
     COLOR 11: LOCATE 3, 1
     GOSUB AddBoat
1410 LOCATE 23, 1: COLOR 11
     
     'Recalculate the Personal Property Figure Here
     'Bring Forth Total Pers Property To Here
     TotBPers# = TotPers#
     VehTr! = TBPPCust(1).BoatFirstTrans
     BoatTotal# = 0
      WHILE VehTr! <> 0
       GET TBBTFile, VehTr!, TBBoat(1)
       IF TBBoat(1).BoatActive = "Y" THEN
        TotBPers# = TotBPers# + TBBoat(1).BoatCost + TBBoat(1).BoatMotorCost
        BoatTotal# = BoatTotal# + TBBoat(1).BoatCost + TBBoat(1).BoatMotorCost
        END IF
       VehTr! = TBBoat(1).BoatNext
      WEND
     TotBPers# = INT((TotBPers# * 100) + .5) / 100
     LOCATE 4, 30: PRINT USING "########,#"; BoatTotal#
     LOCATE 5, 30: PRINT USING "########,#"; TotBPers#
     LOCATE 23, 1: COLOR 11
     PRINT "Select # (ESC-Ok) .. "; : FL = 2
     GOSUB 62000
     IF cf = 1 THEN
     'Update Master Record with total amount for Boats Here
      BoatTotal# = 0
      WHILE VehTr! <> 0
       GET TBBTFile, VehTr!, TBBoat(1)
       IF TBBoat(1).BoatActive = "Y" THEN
        BoatTotal# = BoatTotal# + TBBoat(1).BoatCost + TBBoat(1).BoatMotorCost
        END IF
       VehTr! = TBBoat(1).BoatNext
      WEND
       TBPPCust(1).Boats = BoatTotal#
       PUT tbfile, Record!, TBPPCust(1)
       CLOSE TBBTFile 'Close Boat File
       GOTO 1500
     END IF
     IF VAL(in$) < 12 AND BoatRecord! <= 0 THEN 1400
     IF VAL(in$) = 12 THEN
     'Add New BOAT Here
       FOR zz = 7 TO 23: LOCATE zz, 1: PRINT STRING$(79, 32): NEXT zz: COLOR 11
        LOCATE 7, 1
        PRINT " 1- Boat Year ........            "
        PRINT " 2- Boat Make ........            "
        PRINT " 3- Boat Hull ........            "
        PRINT " 4- Boat Length ......                           IN FEET"
        PRINT " 5- Boat Date Acq'd...                           FORMAT MM/DD/YY"
        PRINT " 6- Boat Cost ........            "
        PRINT " 7- Motor Make .......            "
        PRINT " 8- Motor HP .........            "
        PRINT " 9- Motor Date Acq'd..                           FORMAT MM/DD/YY"
        PRINT "10- Motor Cost .......            "
        PRINT "11- Active (Y) .......            "
        PRINT "12- Add New Boat/Motor            "
        GOSUB 1181
        GOSUB 1182
        GOSUB 1183
        GOSUB 1184
        GOSUB 1185
        GOSUB 1186
        GOSUB 1187
        GOSUB 1188
        GOSUB 1189
        GOSUB 1190
        GOSUB 1191

       TBBoat(1).BoatRecord = Record!
       TBBoat(1).BoatPad = ""
       PUT TBBTFile, (LOF(TBBTFile) \ LEN(TBBoat(1))) + 1, TBBoat(1)

       BoatRecord! = LOF(TBBTFile) \ LEN(TBBoat(1))

      'Update Pointer Here
      IF TBPPCust(1).BoatFirstTrans = 0 THEN
         TBPPCust(1).BoatFirstTrans = BoatRecord!
         TBPPCust(1).BoatLastTrans = BoatRecord!
        PUT tbfile, Record!, TBPPCust(1)
      ELSE
         PrevVeh = TBPPCust(1).BoatLastTrans
         TBPPCust(1).BoatLastTrans = BoatRecord!
         PUT tbfile, Record!, TBPPCust(1)
       GET TBBTFile, PrevVeh, TBBoat(1)
         TBBoat(1).BoatNext = BoatRecord!
        PUT TBBTFile, PrevVeh, TBBoat(1)
      END IF

     END IF

     IF VAL(in$) >= 1 AND VAL(in$) <= 11 THEN
      ON VAL(in$) GOSUB 1181, 1182, 1183, 1184, 1185, 1186, 1187, 1188, 1189, 1190, 1191
      PUT TBBTFile, BoatRecord!, TBBoat(1)
      GOTO 1410
     END IF

     GOTO 1400

1500 REM MH SCREEN FOR PERS PROPERTY
     COLOR 11: LOCATE 3, 1
     GOSUB AddMHome
1510 LOCATE 23, 1: COLOR 11

     'Recalculate the Mobile Home Figure Here
     'Bring Forth Total Pers Property To Here
     VehTr! = TBPPCust(1).MHFirstTrans
     MHTotal# = 0
      WHILE VehTr! <> 0
       GET TBBTFile, VehTr!, TBMH(1)
       IF TBMH(1).MHActive = "Y" THEN
        MHTotal# = MHTotal# + TBMH(1).MHCost
        END IF
       VehTr! = TBMH(1).MHNext
      WEND
     
     LOCATE 4, 30: PRINT USING "########,#"; MHTotal#
     LOCATE 23, 1: COLOR 11
     PRINT "Select # (ESC-Ok) .. "; : FL = -1
     GOSUB 62000
     IF cf = 1 THEN CLOSE TBBTFile: GOTO 1000
     IF VAL(in$) < 9 AND MHRecord! <= 0 THEN 1500
     IF VAL(in$) = 9 THEN
     'Add New MH Here
        FOR zz = 6 TO 23: LOCATE zz, 1: PRINT STRING$(79, 32): NEXT zz: COLOR 11
        LOCATE 6, 1
        COLOR 11
        PRINT " 1- MH Land Owner ........                 "
        PRINT " 2- MH Year ..............                 "
        PRINT " 3- MH Make  .............                 "
        PRINT " 4- MH Length ............                 FORMAT IN FEET"
        PRINT " 5- MH Width .............                 FORMAT IN FEET"
        PRINT " 6- MH Date Acq'd  .......                 FORMAT IN MM/DD/YY"
        PRINT " 7- MH Cost ..............                 "
        PRINT " 8- Active (Y) ...........                 "
        PRINT " 9- Add New Mobile Home"
        

        GOSUB 1192
        GOSUB 1193
        GOSUB 1194
        GOSUB 1195
        GOSUB 1196
        GOSUB 1197
        GOSUB 1198
        GOSUB 1199


       TBMH(1).MHRecord = Record!
       TBMH(1).MHPad = ""
       PUT TBMHFile, (LOF(TBMHFile) \ LEN(TBMH(1))) + 1, TBMH(1)

       MHRecord! = LOF(TBMHFile) \ LEN(TBMH(1))

      'Update Pointer Here
      IF TBPPCust(1).MHFirstTrans = 0 THEN
         TBPPCust(1).MHFirstTrans = MHRecord!
         TBPPCust(1).MHLastTrans = MHRecord!
        PUT tbfile, Record!, TBPPCust(1)
      ELSE
         PrevVeh = TBPPCust(1).MHLastTrans
         TBPPCust(1).MHLastTrans = MHRecord!
        PUT tbfile, Record!, TBPPCust(1)
       GET TBMHFile, PrevVeh, TBMH(1)
         TBMH(1).MHNext = MHRecord!
        PUT TBMHFile, PrevVeh, TBMH(1)
      END IF

     END IF

     IF VAL(in$) >= 1 AND VAL(in$) <= 8 THEN
      ON VAL(in$) GOSUB 1192, 1193, 1194, 1195, 1196, 1197, 1198, 1199
      PUT TBMHFile, MHRecord!, TBMH(1)
      GOTO 1510
     END IF

     GOTO 1510


2000 REM
2010 SCREN$ = "1.02.00": CLS : GOSUB 63500
2020 LOCATE 1, 32: COLOR 11: PRINT "Change/View Customer"; : LOCATE 3, 1
     IF taxtype$ = "R" OR taxtype$ = "P" THEN
      PRINT "(R)eal or (P)ersonal (R,P or ESC-Menu) .. "; taxtype$
     ELSE
      PRINT "(R)eal or (P)ersonal (R,P or ESC-Menu) .. "; : FL = 1
     GOSUB 62210
     IF cf = 1 THEN 10
     IF NOT (in$ = "R" OR in$ = "P") THEN 2020
     taxtype$ = in$
     END IF

2030 IF taxtype$ = "P" THEN
     LOCATE 4, 1: COLOR 11
     PRINT "Enter Acct# (TAB-Name : <F1>-VIN : <F2>-SSN# : ESC-Menu) .. "; : FL = -6
     ELSE
     LOCATE 4, 1: COLOR 11
     PRINT "Enter Acct# (TAB-Name : <F1>-MAP : <F2>-911 Address : ESC-Menu) .. "; : FL = -6
     END IF
     GOSUB 62000
     IF cf = 1 THEN
        CLOSE : GOTO 10
     END IF
     IF cf = 2 THEN
      GOSUB Namesearch
      IF ggo = 1 THEN 2075
      IF ggo = 0 THEN LOCATE 15, 1: PRINT "Account Not Found : Press Any Key to Continue "; : FL = 1: GOSUB 62000: GOTO 2000
     END IF
     IF cf = 11 AND taxtype$ = "P" THEN
      GOSUB VINsearch
      IF ggo = 1 THEN 2075
      IF ggo = 0 THEN LOCATE 15, 1: PRINT "Account Not Found : Press Any Key to Continue "; : FL = 1: GOSUB 62000: GOTO 2000
     END IF

     IF cf = 11 AND taxtype$ = "R" THEN
      GOSUB MapSearch
      IF ggo = 1 THEN 2075
      IF ggo = 0 THEN LOCATE 15, 1: PRINT "Account Not Found : Press Any Key to Continue "; : FL = 1: GOSUB 62000: GOTO 2000
     END IF
     IF cf = 12 AND taxtype$ = "R" THEN
      GOSUB Search911
      IF ggo = 1 THEN 2075
      IF ggo = 0 THEN LOCATE 15, 1: PRINT "Account Not Found : Press Any Key to Continue "; : FL = 1: GOSUB 62000: GOTO 2000
     END IF


     IF cf = 12 AND taxtype$ = "P" THEN
      GOSUB SSsearch
      IF ggo = 1 THEN 2075
      IF ggo = 0 THEN LOCATE 15, 1: PRINT "Account Not Found : Press Any Key to Continue "; : FL = 1: GOSUB 62000: GOTO 2000
     END IF
     Record! = VAL(in$)

2075  GOSUB 60000
     LOCATE 3, 1: PRINT STRING$(79, 32)
     IF taxtype$ = "R" THEN
       IF Record! < 1 OR Record! > LOF(tbfile) / LEN(TBRECust(1)) THEN 2030
       LOCK #1, Record!
       GET 1, Record!, TBRECust(1)
       IF TBRECust(1).Deleted = "Y" THEN GOSUB Deleted: GOTO 2000
       ELSE
       IF Record! < 1 OR Record! > LOF(tbfile) / LEN(TBPPCust(1)) THEN 2030
       LOCK #1, Record!
       GET 1, Record!, TBPPCust(1)
       IF TBPPCust(1).Deleted = "Y" THEN GOSUB Deleted: GOTO 2000
     END IF
     IF taxtype$ = "R" THEN

REALEDIT:
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
      LOCATE 21, 28: PRINT TBRECust(1).NewCom2
REselection2:
     
     LOCATE 23, 1: COLOR 11
     PRINT "Select # (ESC-OK) .. "; : FL = -2
     GOSUB 62000
     IF cf = 1 THEN 2200
     IF VAL(in$) < 1 OR VAL(in$) > 18 THEN GOTO REselection2
     ON VAL(in$) GOSUB 1111, 1112, 1113, 1114, 1115, 1116, 1117, 1118, 1119, 1120, 1121, 1122, 1123, 1124, 1125, 1126, 1127, 1128
      TBRECust(1).LastDateChange = Date2Num(DATE$)
      PUT tbfile, Record!, TBRECust(1)
     GOTO REselection2
     
     ELSE

      REM put personal here
PERSEDIT:
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
      'LOCATE 16, 30: PRINT USING "########,#"; TBPPCust(1).MobileHome
      LOCATE 17, 30: PRINT USING "########,#"; TBPPCust(1).MerchCap
      LOCATE 18, 30: PRINT USING "########,#"; TBPPCust(1).FarmEquip

PPselection2:

     LOCATE 23, 1: COLOR 11
     PRINT "Select # (ESC-OK) .. "; : FL = -2
     GOSUB 62000
     IF cf = 1 THEN 2300
     IF VAL(in$) < 1 OR VAL(in$) > 15 THEN GOTO PPselection2
     ON VAL(in$) GOSUB 1151, 1152, 1153, 1154, 1155, 1156, 1157, 1158, 1159, 1160, 1161, 1163, 1164, 1165, 1166
     PUT tbfile, Record!, TBPPCust(1)
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
     LOCATE 15, 30: PRINT TBRECust(1).MortgageYN
     
2210 LOCATE 18, 1: COLOR 11
     PRINT "Selection # (ESC-OK : TAB-GOTO BACK TO PAGE 1) .. "; : FL = -2: GOSUB 62000
     IF cf = 1 THEN
      CLOSE
      GOTO 2000
     END IF
     IF cf = 2 THEN
      GOTO REALEDIT
     END IF
     
     IF VAL(in$) < 1 OR VAL(in$) > 10 THEN 2210
     ON VAL(in$) GOSUB 1211, 1212, 1213, 1214, 1215, 1216, 1217, 1218, 1219, 1220
     'Note Change Date
      TBRECust(1).LastDateChange = Date2Num(DATE$)
      PUT tbfile, Record!, TBRECust(1)
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
     ' LOCATE 5, 30: PRINT USING "########,#"; TBPPCust(1).Boats
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
      LOCATE 16, 30: PRINT STRING$(30, 32): LOCATE 16, 30: PRINT TBVeh(1).VehQ; TAB(60); "M.ACT-"; TBVeh(1).VehAcct
      LOCATE 17, 30: PRINT STRING$(30, 32): LOCATE 17, 30: PRINT TBVeh(1).VehLic
      LOCATE 18, 30: PRINT STRING$(30, 32): LOCATE 18, 30: PRINT TBVeh(1).VehTyp
      LOCATE 19, 30: PRINT STRING$(5, 32): LOCATE 19, 30: PRINT TBVeh(1).VehTrail
      LOCATE 20, 30: PRINT STRING$(5, 32): LOCATE 20, 30: PRINT TBVeh(1).VehMiles

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
     PRINT "Select # (ESC-OK : TAB-PAGE 1 : <N>-Next Veh : <P>-Prev Veh) .. "; : FL = 2
     GOSUB 62000
     IF cf = 1 THEN GOTO 2400
     IF cf = 2 THEN GOTO PERSEDIT
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
     
     IF VAL(in$) = 15 THEN
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
        PRINT "12- Type/Delete (I/B/D/T/M).  "
        PRINT "13- Trailor Type (U/B/H/R).."
        PRINT "14- High Mileage Flag (Y) .."
        PRINT "15- Add New Vehicle"

      GOSUB 1170
      GOSUB 1171
      GOSUB 1172
      GOSUB 1173
      GOSUB 1174
      GOSUB 1175
      GOSUB 1176
      GOSUB 1177
      GOSUB 1178
      GOSUB 1179
      GOSUB 1180
      TBVeh(1).VehAcct = Record!
      TBVeh(1).VehPad = ""
      PUT TBVFile, (LOF(TBVFile) \ LEN(TBVeh(1))) + 1, TBVeh(1)
      VCNT = VCNT + 1
      VREC!(VCNT) = (LOF(TBVFile) \ LEN(TBVeh(1)))
      VREC = VCNT
      'Update Pointer Here
      IF TBPPCust(1).LastVeh = 0 THEN
       TBPPCust(1).FirstVeh = (LOF(TBVFile) \ LEN(TBVeh(1)))
       TBPPCust(1).LastVeh = (LOF(TBVFile) \ LEN(TBVeh(1)))
       PUT tbfile, Record!, TBPPCust(1)
      ELSE
      PrevVeh = TBPPCust(1).LastVeh
       TBPPCust(1).LastVeh = (LOF(TBVFile) \ LEN(TBVeh(1)))
      PUT tbfile, Record!, TBPPCust(1)
      GET TBVFile, PrevVeh, TBVeh(1)
       TBVeh(1).VehNext = (LOF(TBVFile) \ LEN(TBVeh(1)))
      PUT TBVFile, PrevVeh, TBVeh(1)
      END IF
      GOTO PPselection3
     END IF



     IF VAL(in$) >= 1 AND VAL(in$) <= 3 THEN
      ON VAL(in$) GOSUB 1167, 1168, 1169
      PUT tbfile, Record!, TBPPCust(1)
      GOTO PPselection3
     END IF

     IF VAL(in$) >= 4 AND VAL(in$) <= 14 THEN
      GET TBVFile, VREC!(VREC), TBVeh(1)
     ON (VAL(in$) - 3) GOSUB 1170, 1171, 1172, 1173, 1174, 1175, 1176, 1177, 1178, 1179, 1180
      PUT TBVFile, VREC!(VREC), TBVeh(1)
     END IF

     GOTO PPselection3

2400 GOSUB AddBoat
      'Create BoatRecord List
      BCNT = 0
      BoatRecord! = TBPPCust(1).BoatFirstTrans
      WHILE BoatRecord! <> 0
        GET TBBTFile, BoatRecord!, TBBoat(1)
        IF TBBoat(1).BoatActive = "Y" THEN
         BCNT = BCNT + 1
         BREC!(BCNT) = BoatRecord!
         BREC = 1
        END IF
        BoatRecord! = TBBoat(1).BoatNext
      WEND


      COLOR 15
      LOCATE 3, 60: PRINT USING "######"; Account!
      LOCATE 4, 30: PRINT USING "########,#"; TBPPCust(1).Boats
2410 IF BCNT > 0 THEN
      GET TBBTFile, BREC!(BREC), TBBoat(1)
      COLOR 15
      LOCATE 7, 30: PRINT STRING$(30, 32): LOCATE 7, 30: PRINT TBBoat(1).BoatYear
       COLOR 11: LOCATE 10, 50: PRINT "#"; BREC; " of "; BCNT
      COLOR 15
      LOCATE 8, 30: PRINT STRING$(15, 32): LOCATE 8, 30
       PRINT TBBoat(1).BoatMake
      LOCATE 9, 30: PRINT STRING$(30, 32): LOCATE 9, 30
       PRINT TBBoat(1).BoatHull
      LOCATE 10, 30:  PRINT STRING$(20, 32): LOCATE 10, 30
       PRINT TBBoat(1).BoatLgth
      LOCATE 11, 30:  PRINT STRING$(30, 32): LOCATE 11, 30
       PRINT TBBoat(1).BoatDate
      LOCATE 12, 30: PRINT STRING$(30, 32): LOCATE 12, 30
       PRINT USING "######,#"; TBBoat(1).BoatCost
      LOCATE 13, 30: PRINT STRING$(30, 32): LOCATE 13, 30
       PRINT TBBoat(1).BoatMotorMake; TAB(60); "M.ACT-"; TBBoat(1).BoatRecord
      LOCATE 14, 30: PRINT STRING$(30, 32): LOCATE 14, 30
       PRINT TBBoat(1).BoatMotorHP
      LOCATE 15, 30: PRINT STRING$(30, 32): LOCATE 15, 30
       PRINT TBBoat(1).BoatMotorDate
      LOCATE 16, 30: PRINT STRING$(15, 32): LOCATE 16, 30
       PRINT USING "####,#"; TBBoat(1).BoatMotorCost
      LOCATE 17, 30: PRINT STRING$(5, 32): LOCATE 17, 30
       PRINT TBBoat(1).BoatActive
     END IF

     'Recalculate the Personal Property Figure Here
     TotBoat# = 0
     
     VehTr! = TBPPCust(1).BoatFirstTrans
      WHILE VehTr! <> 0
       GET TBBTFile, VehTr!, TBBoat(1)
       IF TBBoat(1).BoatActive = "Y" THEN
        TotBoat# = TotBoat# + TBBoat(1).BoatCost + TBBoat(1).BoatMotorCost
       END IF
       VehTr! = TBBoat(1).BoatNext
      WEND
     TotBoat# = INT((TotBoat# * 100) + .5) / 100
     LOCATE 4, 30: PRINT USING "########,#"; TotBoat#

     LOCATE 23, 1: COLOR 11
     PRINT "Select # (ESC-OK : <N>-Next Boat : <P>-Prev Boat) .. "; : FL = 2
     GOSUB 62000
     IF cf = 1 THEN GOTO 2500
'     IF cf = 2 THEN GOTO PERSEDIT
     IF UCASE$(in$) = "N" THEN
      'Get Next Vehicle if Any
      BREC = BREC + 1
      IF BREC > BCNT THEN BREC = BCNT
      IF BCNT > 0 THEN GET TBBTFile, BREC!(BREC), TBBoat(1): GOTO 2410
     END IF
     IF UCASE$(in$) = "P" THEN
      'Get Prev Vehicle if Any
      BREC = BREC - 1
      IF BREC < 1 THEN BREC = 1
      IF BCNT > 0 THEN GET TBBTFile, BREC!(BREC), TBBoat(1): GOTO 2410
     END IF

    '++++++++++++++++++++++++++++++++++++++++++
     IF VAL(in$) = 12 THEN
     
     'Add New BOAT Here
       FOR zz = 7 TO 23: LOCATE zz, 1: PRINT STRING$(79, 32): NEXT zz: COLOR 11
        LOCATE 7, 1
        PRINT " 1- Boat Year ........            "
        PRINT " 2- Boat Make ........            "
        PRINT " 3- Boat Hull ........            "
        PRINT " 4- Boat Length ......                           IN FEET"
        PRINT " 5- Boat Date Acq'd...                           FORMAT MM/DD/YY"
        PRINT " 6- Boat Cost ........            "
        PRINT " 7- Motor Make .......            "
        PRINT " 8- Motor HP .........            "
        PRINT " 9- Motor Date Acq'd..                           FORMAT MM/DD/YY"
        PRINT "10- Motor Cost .......            "
        PRINT "11- Active (Y) .......            "
        PRINT "12- Add New Boat/Motor            "
        GOSUB 1181
        GOSUB 1182
        GOSUB 1183
        GOSUB 1184
        GOSUB 1185
        GOSUB 1186
        GOSUB 1187
        GOSUB 1188
        GOSUB 1189
        GOSUB 1190
        GOSUB 1191

       TBBoat(1).BoatRecord = Record!
       TBBoat(1).BoatPad = ""
       PUT TBBTFile, (LOF(TBBTFile) \ LEN(TBBoat(1))) + 1, TBBoat(1)

       BoatRecord! = LOF(TBBTFile) \ LEN(TBBoat(1))

      'Update Pointer Here
      IF TBPPCust(1).BoatFirstTrans = 0 THEN
         TBPPCust(1).BoatFirstTrans = BoatRecord!
         TBPPCust(1).BoatLastTrans = BoatRecord!
        PUT tbfile, Record!, TBPPCust(1)
      ELSE
         PrevVeh = TBPPCust(1).BoatLastTrans
         TBPPCust(1).BoatLastTrans = BoatRecord!
         PUT tbfile, Record!, TBPPCust(1)
       GET TBBTFile, PrevVeh, TBBoat(1)
         TBBoat(1).BoatNext = BoatRecord!
        PUT TBBTFile, PrevVeh, TBBoat(1)
      END IF

      
      BCNT = BCNT + 1
      BREC!(BCNT) = (LOF(TBBTFile) \ LEN(TBBoat(1)))
      BREC = BCNT
      GOTO 2410
     END IF



     IF VAL(in$) >= 1 AND VAL(in$) <= 11 THEN
      GET TBBTFile, BREC!(BREC), TBBoat(1)
     ON VAL(in$) GOSUB 1181, 1182, 1183, 1184, 1185, 1186, 1187, 1188, 1189, 1190, 1191
      PUT TBBTFile, BREC!(BREC), TBBoat(1)
     END IF

     GOTO 2410

2500 ' Mobile Homes Here
      GOSUB AddMHome
      'Create BoatRecord List
      MHCNT = 0
      MHRecord! = TBPPCust(1).MHFirstTrans
      WHILE MHRecord! <> 0
        GET TBMHFile, MHRecord!, TBMH(1)
        IF TBMH(1).MHActive = "Y" THEN
         MCNT = MCNT + 1
         MREC!(MCNT) = MHRecord!
         MREC = 1
        END IF
        MHRecord! = TBMH(1).MHNext
      WEND
      
      COLOR 15
      LOCATE 3, 60: PRINT USING "######"; Account!
      LOCATE 4, 30: PRINT USING "########,#"; TBPPCust(1).MobileHome
2510 IF MCNT > 0 THEN
      GET TBMHFile, MREC!(MREC), TBMH(1)
      COLOR 15
      LOCATE 7, 30: PRINT STRING$(30, 32): LOCATE 7, 30: PRINT TBMH(1).MHLandOwner
       COLOR 11: LOCATE 10, 50: PRINT "#"; MREC; " of "; MCNT
      COLOR 15
      LOCATE 8, 30: PRINT STRING$(15, 32): LOCATE 8, 30
       PRINT TBMH(1).MHYear
      LOCATE 9, 30: PRINT STRING$(30, 32): LOCATE 9, 30
       PRINT TBMH(1).MHMake
      LOCATE 10, 30:  PRINT STRING$(20, 32): LOCATE 10, 30
       PRINT TBMH(1).MHLgth
      LOCATE 11, 30:  PRINT STRING$(30, 32): LOCATE 11, 30
       PRINT TBMH(1).MHWdth
      LOCATE 12, 30: PRINT STRING$(30, 32): LOCATE 12, 30
       PRINT TBMH(1).MHDate
      LOCATE 13, 30: PRINT STRING$(30, 32): LOCATE 13, 30
       PRINT TBMH(1).MHCost; USING "######,#"; MHCost
       PRINT TAB(60); "M.ACT-"; TBBoat(1).BoatRecord
      LOCATE 14, 30: PRINT STRING$(5, 32): LOCATE 17, 30
       PRINT TBMH(1).MHActive
     END IF

     'Recalculate the Personal Property Figure Here
     TotMH# = 0

     VehTr! = TBPPCust(1).MHFirstTrans
      WHILE VehTr! <> 0
       GET TBMHFile, VehTr!, TBMH(1)
       IF TBMH(1).MHActive = "Y" THEN
        TotMH# = TotMH# + TBMH(1).MHCost
       END IF
       VehTr! = TBMH(1).MHNext
      WEND
     TotMH# = INT((TotMH# * 100) + .5) / 100
     LOCATE 4, 30: PRINT USING "########,#"; TotMH#

     LOCATE 23, 1: COLOR 11
     PRINT "Select # (ESC-OK : <N>-Next MobHome : <P>-Prev MobHome) .. "; : FL = 2
     GOSUB 62000
     IF cf = 1 THEN CLOSE : GOTO 2000
     IF UCASE$(in$) = "N" THEN
      'Get Next Vehicle if Any
      MREC = MREC + 1
      IF MREC > MCNT THEN MREC = MCNT
      IF MCNT > 0 THEN GET TBMHFile, MREC!(MREC), TBMH(1): GOTO 2510
     END IF
     IF UCASE$(in$) = "P" THEN
      'Get Prev Vehicle if Any
      MREC = MREC - 1
      IF MREC < 1 THEN MREC = 1
      IF MCNT > 0 THEN GET TBMHFile, MREC!(MREC), TBMH(1): GOTO 2510
     END IF

    '++++++++++++++++++++++++++++++++++++++++++
     IF VAL(in$) = 9 THEN

     'Add New BOAT Here
       FOR zz = 6 TO 23: LOCATE zz, 1: PRINT STRING$(79, 32): NEXT zz: COLOR 11
        LOCATE 6, 1
        COLOR 11
        PRINT " 1- MH Land Owner ........                 "
        PRINT " 2- MH Year ..............                 "
        PRINT " 3- MH Make  .............                 "
        PRINT " 4- MH Length ............                 FORMAT IN FEET"
        PRINT " 5- MH Width .............                 FORMAT IN FEET"
        PRINT " 6- MH Date Acq'd  .......                 FORMAT IN MM/DD/YY"
        PRINT " 7- MH Cost ..............                 "
        PRINT " 8- Active (Y) ...........                 "
        PRINT " 9- Add New Mobile Home"


        GOSUB 1192
        GOSUB 1193
        GOSUB 1194
        GOSUB 1195
        GOSUB 1196
        GOSUB 1197
        GOSUB 1198
        GOSUB 1199



       TBMH(1).MHRecord = Record!
       TBMH(1).MHPad = ""
       PUT TBMHFile, (LOF(TBMHFile) \ LEN(TBMH(1))) + 1, TBMH(1)

       MHRecord! = LOF(TBBTFile) \ LEN(TBBoat(1))

      'Update Pointer Here
      IF TBPPCust(1).MHFirstTrans = 0 THEN
         TBPPCust(1).MHFirstTrans = MHRecord!
         TBPPCust(1).MHLastTrans = MHRecord!
        PUT tbfile, Record!, TBPPCust(1)
      ELSE
         PrevVeh = TBPPCust(1).MHLastTrans
         TBPPCust(1).MHLastTrans = MHRecord!
         PUT tbfile, Record!, TBPPCust(1)
       GET TBMHFile, PrevVeh, TBMH(1)
         TBMH(1).MHNext = MHRecord!
        PUT TBMHFile, PrevVeh, TBMH(1)
      END IF


      MCNT = MCNT + 1
      MREC!(MCNT) = (LOF(TBMHFile) \ LEN(TBMH(1)))
      MREC = MCNT
      GOTO 2510
     END IF



     IF VAL(in$) >= 1 AND VAL(in$) <= 8 THEN
      GET TBBTFile, BREC!(BREC), TBBoat(1)
     ON VAL(in$) GOSUB 1192, 1193, 1194, 1195, 1196, 1197, 1198, 1199
      PUT TBMHFile, MREC!(MREC), TBMH(1)
     END IF

     GOTO 2510
     
3000 REM
3010 SCREN$ = "1.03.00": CLS : GOSUB 63500
3020 LOCATE 1, 32: COLOR 11: PRINT "Delete Customer"; : LOCATE 3, 1
     PRINT "(R)eal or (P)ersonal Account (R,P or ESC-Menu) .. "; : FL = 1
     GOSUB 62210
     IF cf = 1 THEN 10
     IF NOT (in$ = "R" OR in$ = "P") THEN 3020
     taxtype$ = in$
3030 LOCATE 4, 1: COLOR 11
     PRINT "Enter Customer Account Number (TAB-Last Name : ESC-Menu) .. "; : FL = -6
     GOSUB 62000
     IF cf = 1 THEN
      CLOSE : GOTO 10
     END IF
     IF cf = 2 THEN
      GOSUB Namesearch
      IF ggo = 1 THEN 3075
      IF ggo = 0 THEN LOCATE 10, 1: PRINT "Account Not Found : Press Any Key to Continue "; : FL = 1: GOSUB 62000: GOTO 3000
     END IF
     Record! = VAL(in$)
3075 CLOSE : GOSUB 60000
     IF taxtype$ = "R" THEN
       IF Record! < 1 OR Record! > LOF(tbfile) / LEN(TBRECust(1)) THEN 3030
       LOCK #1, Record!
       GET 1, Record!, TBRECust(1)
       IF TBRECust(1).Deleted = "Y" THEN GOSUB Deleted: GOTO 3000
       ELSE
       IF Record! < 1 OR Record! > LOF(tbfile) / LEN(TBPPCust(1)) THEN 3030
       LOCK #1, Record!
       GET 1, Record!, TBPPCust(1)
       IF TBPPCust(1).Deleted = "Y" THEN GOSUB Deleted: GOTO 3000
     END IF

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
3110 LOCATE 10, 5: PRINT "Delete (Y/N) (ESC-Menu) .. "; : FL = 1
      GOSUB 62210
      IF NOT (in$ = "Y" OR in$ = "N") THEN 3110
      IF in$ = "Y" THEN
       TBRECust(1).Deleted = "Y"
       PRINT CHR$(7);
       PUT tbfile, Record!, TBRECust(1)
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
3210 LOCATE 10, 5: PRINT "Delete (Y/N) (ESC-Menu) .. "; : FL = 1
      GOSUB 62210
      IF NOT (in$ = "Y" OR in$ = "N") THEN 3210
      IF in$ = "Y" THEN
       TBPPCust(1).Deleted = "Y"
       PUT tbfile, Record!, TBPPCust(1)
      END IF
      CLOSE
      GOTO 3000

     END IF

4000 CLS
        COLOR 11
        PRINT "Do you wish to sort now (Y/N) .. "; : FL = 1
        GOSUB 62210
        IF in$ = "Y" THEN
        RUN "tbNMsort"
        ELSE
        CLOSE : GOTO 10
        END IF

5000 RUN "tbcmile"
6000 RUN "tbcpval"
7000 RUN "tbcvin"
8000 RUN "tbreassm"
9000 RUN "tbroll"
10000 RUN "tbupddmv"
11000 RUN "TBNADAPP"
12000 RUN "TBNADAUP"

60000  IF taxtype$ = "R" THEN
         CLOSE tbfile
         tbfile = FREEFILE
         OPEN TBPath$ + "TBRECUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #tbfile LEN = LEN(TBRECust(1))
        ELSEIF taxtype$ = "P" THEN
         CLOSE tbfile
         tbfile = FREEFILE
         OPEN TBPath$ + "TBPPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #tbfile LEN = LEN(TBPPCust(1))
         CLOSE TBVFile
         TBVFile = FREEFILE
         OPEN TBPath$ + "TBVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBVFile LEN = LEN(TBVeh(1))
         CLOSE TBBTFile
         TBBTFile = FREEFILE
         OPEN TBPath$ + "TBBOAT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBTFile LEN = LEN(TBBoat(1))
         CLOSE TBMHFile
         TBMHFile = FREEFILE
         OPEN TBPath$ + "TBMHOME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBMHFile LEN = LEN(TBMH(1))
       ELSEIF taxtype$ = "S" THEN
         CLOSE tbfile
         tbfile = FREEFILE
         OPEN TBPath$ + "TBSPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #tbfile LEN = LEN(TBPPCust(1))
        END IF
60025 TBBalFile = FREEFILE
      IF taxtype$ = "R" THEN
      OPEN TBPath$ + "TBREBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBRECustBal(1))
      ELSE
      OPEN TBPath$ + "TBPPBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBPPCustBal(1))
      END IF
60030 RETURN
        

62000 cf = 0: COLOR 14: in$ = "": RO = CSRLIN: PO = POS(X): ZD = 0: ZS = 0: ZL = 0: IF FL = ZD THEN FL = 1
62010 IF INKEY$ <> "" THEN 62010
62020 PRINT STRING$(ABS(FL), 95); : LOCATE RO, PO
62030 z$ = INKEY$: IF z$ = "" THEN 62030
62040 IF z$ <> CHR$(13) THEN 62060 ELSE PRINT STRING$(ABS(FL) - ZL, 32);
62050 GOSUB 62200: RETURN
62060 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN cf = 1: RETURN
      IF z$ = CHR$(9) THEN cf = 2: RETURN
      IF LEN(z$) < 2 THEN 62065
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 59 THEN cf = 11: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 60 THEN cf = 12: RETURN
62065 IF ASC(z$) = 0 THEN 62030
62070 IF LEN(z$) > 1 THEN cf = ASC(RIGHT$(z$, 1)) - 57: IF cf > 0 THEN RETURN ELSE cf = cf + 27: RETURN
62080 IF z$ = " " THEN 62130 ELSE IF z$ = CHR$(29) THEN PRINT STRING$(ZL, 29); : GOTO 62000
62090 IF z$ <> CHR$(8) THEN 62130 ELSE IF ZL = 0 THEN 62030 ELSE PRINT CHR$(29); : PRINT "_"; : PRINT CHR$(29); : IF FL > 0 THEN 62110 ELSE IF z$ = "," THEN 62120
62100 IF z$ = "." THEN ZD = 0: GOTO 62110 ELSE IF z$ = "+" OR z$ = "-" THEN ZS = 0
62110 in$ = LEFT$(in$, LEN(in$) - 1)
62120 ZL = ZL - 1: GOTO 62030
62130 IF ABS(FL) = ZL THEN 62030 ELSE IF FL > 0 THEN IF z$ >= " " AND z$ <= "z" THEN 62180
62140 IF z$ = "." AND ZD = 0 THEN ZD = 1: GOTO 62180
62150 IF z$ = "," THEN PRINT ","; : ZL = ZL + 1: GOTO 62190
62160 IF (z$ = "-" OR z$ = "+") AND ZS = 0 AND ZL = 0 THEN ZS = 1: GOTO 62180
62170 IF z$ < "0" OR z$ > "9" THEN 62030
62180 PRINT z$; : in$ = in$ + z$: ZL = ZL + 1
62190 IF ABS(FL) = 1 THEN 62050 ELSE 62030
62200 LOCATE RO, PO: PRINT in$: RETURN
62210 GOSUB 62000: IF LEN(in$) THEN in$ = CHR$(ASC(in$) AND 223)
62220 RETURN
63000 RETURN
63500 COLOR 10: LOCATE 1, 1: PRINT "Sys. Date: "; SYSDATE$; TAB(70); SCREN$: PRINT STRING$(79, "_"): RETURN

Namesearch:
 LOCATE 8, 1:
 COLOR 11
 PRINT "Enter Last Name (Partial Accepted : ESC - End) .. "; : FL = 25
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
  TBIDXFILE = 14
  OPEN TBPath$ + "TBRENAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFILE LEN = 24
  CLOSE 14
  TBIDXFILE = 14
  OPEN TBPath$ + "TBRENAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFILE LEN = 24

  FIELD TBIDXFILE, 20 AS lastnam$, 4 AS MREC$
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
  GET tbfile, CVS(MREC$), TBRECust(1)
  LOCATE 11, 1: COLOR 11
  Nme$ = TBRECust(1).CustName
  COLOR 10
  FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  LOCATE 11, 1: COLOR 11
  PRINT "Full Account Name: "; : COLOR 15: PRINT Nme$
  COLOR 11
  PRINT "   Account Number: "; : COLOR 15: PRINT CVS(MREC$)
n2:
  LOCATE 13, 1: COLOR 11
  PRINT "Correct Account (Y/N/Q) .. "; : FL = 1: GOSUB 62210
  IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO n2
  IF in$ = "Q" THEN
   FOR LLL = 8 TO 13: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
   CLOSE
   ggo = 0
   RETURN
  END IF
  IF in$ = "Y" THEN
   ggo = 1: Record! = CVS(MREC$)
   FOR LLL = 8 TO 13:  LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
   CLOSE
   RETURN
  END IF
  IF in$ = "N" THEN GOTO n3
  GOTO n2
  END IF

  IF taxtype$ = "P" THEN
  CLOSE
  TBIDXFILE = 14
  OPEN TBPath$ + "TBPPNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFILE LEN = 24
  FIELD TBIDXFILE, 20 AS lastnam$, 4 AS MREC$
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
GET tbfile, CVS(MREC$), TBPPCust(1)
LOCATE 11, 1: COLOR 11
Nme$ = TBPPCust(1).CustName
COLOR 10
FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT Nme$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT CVS(MREC$)
PPn2:
LOCATE 13, 1: COLOR 11
PRINT "Correct Account (Y/N/Q) .. "; : FL = 1: GOSUB 62210
IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO PPn2
IF in$ = "Q" THEN
  FOR LLL = 8 TO 13: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  CLOSE
  ggo = 0
  RETURN
END IF
IF in$ = "Y" THEN
  ggo = 1: Record! = CVS(MREC$)
  FOR LLL = 8 TO 13: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  CLOSE
  RETURN
END IF
IF in$ = "N" THEN GOTO PPn3
GOTO PPn2
END IF


'************************* M A P  S E A R C H  R O U T I N E ***************

MapSearch:
 LOCATE 8, 1
 COLOR 11
 PRINT "Enter the Map Number (ESC-End) .. "; : FL = 14
 GOSUB 62000
 IF cf = 1 THEN
   ggo = 0
  FOR LLL = 8 TO 13: PRINT STRING$(79, 32): NEXT LLL
  RETURN
 END IF

 MapNumber$ = UCASE$(in$)
 MapLgth = LEN(in$)
 GOSUB 60000
 FOR Snt& = 1 TO LOF(tbfile) \ (LEN(TBRECust(1)))
 GET tbfile, Snt&, TBRECust(1)

 IF LEFT$(MapNumber$, MapLgth) = LEFT$(RTRIM$(TBRECust(1).MapNumber), MapLgth) THEN
   COLOR 11
   LOCATE 14, 1: PRINT "Full Map Number is "; : COLOR 15
   PRINT LEFT$(TBRECust(1).MapNumber, 3) + "-" + MID$(TBRECust(1).MapNumber, 4, 2) + "-" + MID$(TBRECust(1).MapNumber, 6, 2) + "-" + MID$(TBRECust(1).MapNumber, 8, 2) + "-" + MID$(TBRECust(1).MapNumber, 10, 4) + "-" + MID$(TBRECust(1).MapNumber, 14 _
, 1)
   COLOR 11
   LOCATE 15, 1: PRINT "Customer Name is   "; : COLOR 15: PRINT TBRECust(1).CustName: COLOR 11
   COLOR 11
   LOCATE 16, 1: PRINT "Correct (Y/N) .. "; : FL = 1
   GOSUB 62210
   IF in$ = "Y" THEN
   Record! = Snt&
   ggo = 1
   RETURN
   END IF
 END IF
 NEXT Snt&
  ggo = 0
  CLOSE
 RETURN

'************************ E N D   M A P   S E A R C H   R O U T I N E  **********

'************************* 911  S E A R C H  R O U T I N E ***************

Search911:
 LOCATE 8, 1
 COLOR 11
 PRINT "Enter the 911 Address (or any part of) (ESC-End) .. "; : FL = 25
 GOSUB 62000
 IF cf = 1 THEN
   ggo = 0
  FOR LLL = 8 TO 13: PRINT STRING$(79, 32): NEXT LLL
  RETURN
 END IF

 MapNumber$ = UCASE$(in$)
 MapLgth = LEN(in$)
 GOSUB 60000
 FOR Snt& = 1 TO LOF(tbfile) \ (LEN(TBRECust(1)))
 GET tbfile, Snt&, TBRECust(1)

 IF LEFT$(MapNumber$, MapLgth) = LEFT$(RTRIM$(TBRECust(1).NewCom2), MapLgth) THEN
   COLOR 11
   LOCATE 14, 1: PRINT "Full Address is "; : COLOR 15
   PRINT TBRECust(1).NewCom2
   COLOR 11
   LOCATE 15, 1: PRINT "Customer Name is   "; : COLOR 15: PRINT TBRECust(1).CustName: COLOR 11
   COLOR 11
   LOCATE 16, 1: PRINT "Correct (Y/N) .. "; : FL = 1
   GOSUB 62210
   IF in$ = "Y" THEN
   Record! = Snt&
   ggo = 1
   RETURN
   END IF
 END IF
 NEXT Snt&
  ggo = 0
  CLOSE
 RETURN

'************************ E N D   9 1 1   S E A R C H   R O U T I N E  **********


        

VINsearch:
 LOCATE 8, 1:
 COLOR 11
 PRINT "Enter Vehicle VIN (Partial Accepted : ESC - End) .. "; : FL = 22
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
  
  IF VIN$ = RTRIM$(LEFT$(TBVeh(1).VehVin, VinLgth)) AND TBVeh(1).VehTyp <> "D" THEN
   Record! = TBVeh(1).VehAcct
   GET tbfile, Record!, TBPPCust(1)
   IF TBPPCust(1).Deleted <> "Y" THEN GOTO VPP1
  END IF
Vpp4:
 NEXT Snt&
 CLOSE TBVFile
 ggo = 0
 RETURN


VPP1:
GET tbfile, Record!, TBPPCust(1)
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
PRINT "Correct Account (Y/N/Q) .. "; : FL = 1: GOSUB 62210
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

SSsearch:
 LOCATE 8, 1:
 COLOR 11
 PRINT "Enter Social Security # (Partial Accepted : ESC - End) .. "; : FL = 11
 GOSUB 62000
 IF cf = 1 THEN
  ggo = 0
  FOR LLL = 8 TO 13: PRINT STRING$(79, 32): NEXT LLL
  RETURN
 END IF
 SSN$ = UCASE$(in$)
 SSLgth = LEN(SSN$)
 GOSUB 60000
 FOR Record! = 1 TO LOF(tbfile) / LEN(TBPPCust(1))
 GET tbfile, Record!, TBPPCust(1)
 IF SSN$ = LEFT$(TBPPCust(1).SocSec, SSLgth) OR SSN$ = LEFT$(TBPPCust(1).SocSec2, SSLgth) THEN
  IF TBPPCust(1).Deleted <> "Y" THEN GOTO SPP1
 END IF

Spp4:
 NEXT Record!
 CLOSE tbfile
 ggo = 0
 RETURN


SPP1:
LOCATE 11, 1: COLOR 11
Nme$ = TBPPCust(1).CustName
COLOR 10
FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT Nme$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT Record!

SPP2:
LOCATE 13, 1: COLOR 11
PRINT "Correct Account (Y/N/Q) .. "; : FL = 1: GOSUB 62210
IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO SPP2
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
IF in$ = "N" THEN GOTO Spp4
GOTO SPP2



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
    FL = 1
    GOSUB 62000
    LOCATE 23, 1: PRINT STRING$(79, 32)
    RETURN
HasBalance:
    PRINT CHR$(7);
    LOCATE 23, 1: COLOR 15
    PRINT "This Account Has a Balance and CANNOT Be Deleted! Press <ENTER> to Continue"; :
    FL = 1
    GOSUB 62000
    LOCATE 23, 1: PRINT STRING$(79, 32)
    RETURN



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
PRINT "18- 911 Address .........."
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
PRINT "10- Mortgage Customer Y .."
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
PRINT "13- Not Used  .............."
PRINT "14- Merch Cap. Value ......."
PRINT "15- Farm Equip Value ......."
RETURN



AddPers2:
FOR zz = 4 TO 23: LOCATE zz, 1: PRINT STRING$(79, 32): NEXT zz: COLOR 11
LOCATE 3, 1
COLOR 10
PRINT "Personal Property Valuation Detail Screen"
COLOR 11
PRINT " 1- Heavy Equipment ........"
PRINT " 2- Not Used ..............."
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
PRINT "12- Type/Delete (I/B/D/T/M)."
PRINT "13- Trailor Type (U/B/H/R)..  "
PRINT "14- High Mileage Flag (Y)...  "
PRINT "15- Add New Vehicle"
RETURN

AddBoat:
FOR zz = 3 TO 23: LOCATE zz, 1: PRINT STRING$(79, 32): NEXT zz: COLOR 11
LOCATE 3, 1
COLOR 10
PRINT "Personal Property Boat Valuation Detail Screen"
COLOR 11
PRINT "  - Boat Total ............."; TAB(60); "ACCT # "; Record!
PRINT "  - Total Personal Property."
PRINT STRING$(79, "-")
PRINT " 1- Boat Year ........"
PRINT " 2- Boat Make ........"
PRINT " 3- Boat Hull ........"
PRINT " 4- Boat Length ......"
PRINT " 5- Boat Date Acq'd..."
PRINT " 6- Boat Cost ........"
PRINT " 7- Motor Make ......."
PRINT " 8- Motor HP ........."
PRINT " 9- Motor Date Acq'd.."
PRINT "10- Motor Cost ......."
PRINT "11- Active (Y) ......."
PRINT "12- Add New Boat/Motor"
RETURN

AddMHome:
FOR zz = 3 TO 23: LOCATE zz, 1: PRINT STRING$(79, 32): NEXT zz: COLOR 11
LOCATE 3, 1
COLOR 10
PRINT "Personal Property Mobile Home Valuation Detail Screen"
COLOR 11
PRINT "  - MH Total ............."
PRINT STRING$(79, "-")
PRINT " 1- MH Land Owner ........"
PRINT " 2- MH Year .............."
PRINT " 3- MH Make  ............."
PRINT " 4- MH Length ............"
PRINT " 5- MH Width ............."
PRINT " 6- MH Date Acq'd  ......."
PRINT " 7- MH Cost .............."
PRINT " 8- Active (Y) ......."
PRINT " 9- Add New Mobile Home"
RETURN


networkerror:
     REM ERROR CODES
     Number = ERR
     SELECT CASE ERR
          CASE IS = 24, 25
               LOCATE 23, 1: PRINT STRING$(79, 32): LOCATE 23, 1: COLOR 28
               PRINT "Error Message:"; : COLOR 15: PRINT "Device Time-out! Make Sure Printer Is ON-LINE:Press ENTER to Resume Printing";
               FL = 1: GOSUB 62000: LOCATE 23, 1: PRINT STRING$(79, 32)
               RESUME
          CASE IS = 27
               LOCATE 23, 1: PRINT STRING$(79, 32): LOCATE 23, 1: COLOR 28
               PRINT "Error Message:"; : COLOR 15: PRINT "Printer Is OUT OF PAPER! Load Paper : Press ENTER to Resume Printing";
               FL = 1: GOSUB 62000: LOCATE 23, 1: PRINT STRING$(79, 32)
               RESUME
          CASE IS = 6
               RESUME NEXT

          CASE ELSE
               CLS
               LOCATE 15, 1: COLOR 28: PRINT "Error Message: "; : COLOR 15: PRINT "Error "; Number; " Has Occured in Line "; ERL
               PRINT "Write Down Message and Hit <ESC> Key to exit."
               FL = 1: GOSUB 62000: IF cf = 1 THEN CLOSE : GOTO 10 ELSE GOSUB 62000
            
     END SELECT

