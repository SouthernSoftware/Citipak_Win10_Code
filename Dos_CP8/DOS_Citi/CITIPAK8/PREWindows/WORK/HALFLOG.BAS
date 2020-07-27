DEFINT A-Z
DIM a AS STRING * 8000
OPEN "ublog.dat" FOR BINARY AS #1
OPEN "nublog.dat" FOR BINARY AS #2
Fz& = LOF(1)
Half& = Fz& - 1000000
SEEK #1, Half&

DO
  GET #1, , a$
  PUT #2, , a$
  IF SEEK(1) >= Fz& THEN
    EXIT DO
  END IF
  PRINT "aa"
LOOP
CLOSE
END

