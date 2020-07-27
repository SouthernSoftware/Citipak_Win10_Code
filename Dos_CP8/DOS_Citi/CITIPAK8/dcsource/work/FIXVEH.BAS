DEFINT A-Z
  '$INCLUDE: 'DC.bi'                        'A/R FILE LAYOUTS

  DIM DCVRec(1)  AS DCVehType

  DCVehRecLen = LEN(DCVRec(1))
  DCVFile = FREEFILE
  OPEN "DCVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCVFile LEN = DCVehRecLen

  NumVeh& = LOF(DCVFile) / DCVehRecLen
  FOR ZZ& = 1 TO NumVeh&
    GET DCVFile, ZZ&, DCVRec(1)
    IF DCVRec(1).NextRec = ZZ& THEN
      DCVRec(1).NextRec = 0
      PUT DCVFile, ZZ&, DCVRec(1)
      Bad = Bad + 1
    END IF
  NEXT
  CLOSE

  PRINT "Bad:"; Bad; " of"; NumVeh&

