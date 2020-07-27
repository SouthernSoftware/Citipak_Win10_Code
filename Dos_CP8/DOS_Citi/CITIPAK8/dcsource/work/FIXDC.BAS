DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DEFINT A-Z
  '$INCLUDE: 'DC.BI'


    REDIM DCCustRec(1 TO 2) AS DCCustRecType
    REDIM DCVRec(1 TO 2)  AS DCVehType
    REDIM DCTrans(1 TO 2)   AS DCTransRecType

    DCCustLen = LEN(DCCustRec(1))
    DCVehLen = LEN(DCVRec(1))
    DCTranLen = LEN(DCTrans(1))

'GOTO skip2here

    CFile = FREEFILE
    OPEN "DCCust.dat" FOR RANDOM SHARED AS CFile LEN = DCCustLen
    NumOfCust& = LOF(CFile) / DCCustLen

    VFile = FREEFILE
    OPEN "DCVEH.dat" FOR RANDOM SHARED AS VFile LEN = DCVehLen
    NumOfVeh& = LOF(VFile) / DCVehLen
    'GET VFile, Cnt&, DCVehRec(1)

    FOR Cnt& = 1 TO NumOfCust&
      GET CFile, Cnt&, DCCustRec(1)
      IF UCASE$(DCCustRec(1).Deleted) = "Y" THEN
        Del = Del + 1
        GOTO DelSkip

      END IF

      IF VAL(DCCustRec(1).CUSTNUMB) <> Cnt& THEN
        DCCustRec(1).Deleted = "Y"
        PUT CFile, Cnt&, DCCustRec(1)
        BadCnt = BadCnt + 1
        GOTO DelSkip
      END IF

      'DCCustRec(1).FirstCar = 0
      CarRecord! = DCCustRec(1).FirstCar

      WHILE (CarRecord! > 0) AND (CarRecord! <= NumOfVeh&)
        GET VFile, CarRecord!, DCVRec(1)

        IF VAL(DCCustRec(1).CUSTNUMB) <> Cnt& THEN
          BadCnt = BadCnt + 1
          GOTO DelSkip
        END IF
        IF VAL(DCCustRec(1).CUSTNUMB) <> DCVRec(1).MasterRecord THEN

        'IF Cnt& <> DCVRec(1).MasterRecord THEN 'STOP
           DCVRec(1).MasterRecord = Cnt&
           PUT VFile, CarRecord!, DCVRec(1)
           Bad = Bad + 1
         END IF
         IF DCVRec(1).NextRec <> CarRecord! THEN
           CarRecord! = DCVRec(1).NextRec
         ELSE
           CarRecord! = 0
         END IF

      WEND
DelSkip:
    NEXT
    CLOSE
    END
    'STOP

'    FOR Cnt& = 1 TO NumOfCust&
'      GET CFile, Cnt&, DCCustRec(1)
'      SELECT CASE DCCustRec(1).Deleted
'      CASE "Y", "N"
'      CASE ELSE
'        'IF DCCustRec(1).FirstCar > 0 THEN
'          STOP
'        'END IF
'
'        'NoMast = NoMast + 1
'        'DCCustRec(1).BILLNAME
'        'STOP
'      END SELECT
'        'IF DCCustRec(1).FirstCar > 0 THEN
'        '  STOP
'        'END IF
'    NEXT
'    CLOSE
'    END


ExitRelink:

