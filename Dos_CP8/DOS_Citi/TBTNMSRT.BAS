DEFINT A-Z

TYPE Struct
   who AS STRING * 30
   recnum AS INTEGER
END TYPE

DECLARE SUB SortT (SEG Element AS Struct, NumEls, Dir, SSize, MOff, MSize)
'$INCLUDE: 'LCTax.bi'
CLEAR , , 5000          'the sort routines need the extra stack space

REDIM TBRECust(1) AS TBRECust
taxtype$ = "R"
TBPath$ = "F:\"         'Change to F:\ when compiling"


CLS
Size = 13000
Start = 1               'start at array element 1
Dir = 0                 'sort direction - use anything else for descending
SSize = 32               'total size of each TYPE element
MOff = 0                'offset into the TYPE for the key element
MSize = 30              'size of the key element - coded as follows:
                        '   -1 = integer
                        '   -2 = long integer
                        '   -3 = single precision
                        '   -4 = double precision
                        '   +N = TYPE array/fixed-length string of length N

REDIM array(1 TO Size) AS Struct

CLS
COLOR 11
PRINT "Tax System : Last Name Sort Utility"
'PRINT 1;
GOSUB 60000

FOR x = 1 TO LOF(TBFile) / LEN(TBRECust(1))
GET 1, x, TBRECust(1)
'LOCATE 2, 1
'PRINT "Reading:"; Count;

 'IF LEFT$(TBRECust(1).CustName, 1) <= "F" AND TBRECust(1).Deleted <> "Y" THEN
 IF TBRECust(1).Deleted <> "Y" THEN
    Count = Count + 1
    array(Count).who = LEFT$(TBRECust(1).CustName, 30)
    array(Count).recnum = x
 END IF

LOCATE 2, 1
PRINT "Reading:"; Count;

NEXT x

LOCATE 3, 1
PRINT "Sorting. . ."
SortT array(Start), Count, Dir, SSize, MOff, MSize

OPEN TBPath$ + "TBRTNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #14 LEN = 24
FIELD 14, 20 AS lastnam$, 4 AS mrec$
CLOSE 14
FileNam$ = TBPath$ + "TBRTNAME.DAT"
KILL FileNam$
OPEN TBPath$ + "TBRTNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #14 LEN = 24
FIELD 14, 20 AS lastnam$, 4 AS mrec$


FOR x = 1 TO Count
  LOCATE 4, 1
  PRINT "Writing:"; x
  LSET lastnam$ = array(x).who
  LSET mrec$ = MKS$(array(x).recnum)
  PUT 14, LOF(14) / 24 + 1
NEXT
CLOSE
LOCATE 6, 1
PRINT "Sort complete."
LOCATE 9, 1
PRINT "Press any key. . . "
DO WHILE zz$ = ""
 zz$ = INKEY$
LOOP

'RUN "tbTnmsr1"

RUN "TBtcm"

60000  IF taxtype$ = "R" THEN
         TBFile = FREEFILE
         OPEN TBPath$ + "TBRTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
        ELSE
         TBFile = FREEFILE
         OPEN TBPath$ + "TBPTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
         END IF
        RETURN

