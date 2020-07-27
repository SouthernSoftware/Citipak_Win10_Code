DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'DC.BI'

    CLS
    LOCATE 2, 1
    PRINT "Correcting Decal Records. . ."

    REDIM DCCustRec(1 TO 2) AS DCCustRecType
    REDIM DCVehRec(1 TO 2)  AS DCVehType

    DCCustLen = LEN(DCCustRec(1))
    DCVehLen = LEN(DCVehRec(1))
    
    CFile = FREEFILE
    OPEN "DCCust.dat" FOR RANDOM SHARED AS CFile LEN = DCCustLen
    NumOfCust& = LOF(CFile) / DCCustLen

    VFile = FREEFILE
    OPEN "DCVEH.dat" FOR RANDOM SHARED AS VFile LEN = DCVehLen
    NumOfVeh& = LOF(VFile) / DCVehLen
    FOR Cnt& = 1 TO NumOfCust&
      LOCATE 5, 1
      PRINT "Processing:"; Cnt&; "of"; NumOfCust&;
      GET CFile, Cnt&, DCCustRec(1)
      IF LEN(QPTrim$(DCCustRec(1).CustNumb)) = 0 THEN
        DCCustRec(1).Deleted = "Y"
        Removed& = Removed& + 1
        PUT CFile, Cnt&, DCCustRec(1)
      END IF
      IF DCCustRec(1).Deleted <> "Y" THEN
        IF DCCustRec(1).FirstCar > 0 THEN
          GET VFile, DCCustRec(1).FirstCar, DCVehRec(1)
          IF DCVehRec(1).MasterRecord <> Cnt& THEN
            DCVehRec(1).MasterRecord = Cnt&
            PUT VFile, DCCustRec(1).FirstCar, DCVehRec(1)
          END IF
          NextCar& = DCVehRec(1).NextRec
          DO WHILE NextCar& > 0
            GET VFile, NextCar&, DCVehRec(1)
            IF DCVehRec(1).MasterRecord <> Cnt& THEN
              DCVehRec(1).MasterRecord = Cnt&
              PUT VFile, NextCar&, DCVehRec(1)
            END IF
            NextCar& = DCVehRec(1).NextRec
          LOOP
        END IF
      END IF
    NEXT
    CLOSE
    LOCATE 9, 1
    PRINT "Processing Complete."
    CALL KillFile("fixmastr.exe")

    END

