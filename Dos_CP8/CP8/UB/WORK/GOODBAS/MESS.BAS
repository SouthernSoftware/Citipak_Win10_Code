DEFINT A-Z

TYPE bigtype
mess1 AS STRING * 32767
'mess2 AS STRING * 32763
END TYPE

REDIM dodo(1) AS bigtype

OPEN "e:\dodo.txt" FOR RANDOM AS #1 LEN = 32767
FOR zz& = 1 TO 532000
  LOCATE 5, 1
  PRINT "Count:"; zz&;
  PUT #1, 120000 + zz&, dodo(1)
NEXT

CLOSE
END

