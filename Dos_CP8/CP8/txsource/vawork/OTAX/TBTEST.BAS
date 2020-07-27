DEFINT A-Z

DIM NM$(1050), REC(1050)
CLS

TYPE TEMPINDEX
 NAM AS STRING * 10
 REC AS SINGLE
END TYPE
DECLARE SUB SORTT2 (SEG LASTNAME, NumEls, Dir, SSize, MOff, MSize)
REDIM TEMPINDEX(1 TO 1000) AS STRUCT
GOSUB 60000
FOR X = 1 TO 1000: GET 1, X
IF VAL(ACCT$) > 0 THEN
 ARRAY(X).NAM = LEFT$(LNAME$, 10)
 TREC(X).REC = X
END IF
NEXT X
SORTT2 ARRAY(1), COUNT, 0, 8, 0, n
FOR X = 1 TO COUNT: PRINT NM$(X), REC(X), X: NEXT X


STOP


60000 CLOSE 1: OPEN "tbcust.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #1 LEN = 512
60010 FIELD 1, 12 AS ACCT$, 60 AS LNAM$, 35 AS A1$, 35 AS A2$, 20 AS cy$, 2 AS st$, 10 AS zp$, 12 AS tp$
      FIELD 1, 186 AS dummy$, 39 AS pdesc$, 22 AS pmap$, 35 AS psub$, 4 AS plot$, 35 AS pcounty$, 4 AS psize$, 8 AS pland$, 8 AS mv$
      FIELD 1, 341 AS dummy1$, 8 AS mh$, 8 AS fm$, 8 AS ls$, 4 AS op$, 4 AS me$, 6 AS homeexp$, 1 AS taxtype$, 3 AS taxtype1$
      FIELD 1, 381 AS dummy2$, 4 AS curlateamt$, 4 AS curtaxamt$, 4 AS curintamt$, 4 AS curadvamt$, 4 AS curcolamt$, 4 AS ctaxnotice$
      FIELD 1, 405 AS dummy3$, 4 AS ftr$, 4 AS ltr$, 11 AS ss$, 1 AS late$, 15 AS pin$, 68 AS blank$
    RETURN

