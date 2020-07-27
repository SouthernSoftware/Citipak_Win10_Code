DEFINT A-Z
TYPE Struct
   who AS STRING * 28
   recnum AS INTEGER
END TYPE

DECLARE SUB SortT (SEG Element AS Struct, NumEls, Dir, SSize, MOff, MSize)
'$INCLUDE: 'LCTax.bi'
CLEAR , , 4000          'the sort routines need the extra stack space

REDIM TBPPCust(1) AS TBPPCust
taxtype$ = "P"
TBPath$ = "F:\"         'Change to F:\ when compiling"


CLS
Size = 4100
Start = 1               'start at array element 1
Dir = 0                 'sort direction - use anything else for descending
SSize = 30               'total size of each TYPE element
MOff = 0                'offset into the TYPE for the key element
MSize = 28              'size of the key element - coded as follows:
                        '   -1 = integer
                        '   -2 = long integer
                        '   -3 = single precision
                        '   -4 = double precision
                        '   +N = TYPE array/fixed-length string of length N

REDIM array(1 TO Size) AS Struct

CLS
COLOR 11
PRINT "Tax System : Last Name Sort Utility"
GOSUB 60000
FOR x = 1 TO LOF(TBFile) / LEN(TBPPCust(1))
GET 1, x, TBPPCust(1)

  IF LEFT$(TBPPCust(1).CustName, 1) >= "E" AND LEFT$(TBPPCust(1).CustName, 1) <= "O" AND TBPPCust(1).Deleted <> "Y" THEN
    count = count + 1
    array(count).who = LEFT$(TBPPCust(1).CustName, 28)
    array(count).recnum = x
 END IF
NEXT x


SortT array(Start), count, Dir, SSize, MOff, MSize

OPEN TBPath$ + "TBPPNAME.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
FIELD 4, 20 AS lastnam$, 4 AS mrec$


FOR x = 1 TO count
'    PRINT array(x).who
    LSET lastnam$ = array(x).who
    LSET mrec$ = MKS$(array(x).recnum)
    PUT 4, LOF(4) / 24 + 1
NEXT
CLOSE
RUN "TBNMSor5"

60000  IF taxtype$ = "R" THEN
         TBFile = FREEFILE
         OPEN TBPath$ + "TBRECUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
        ELSE
         TBFile = FREEFILE
         OPEN TBPath$ + "TBPPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
         END IF
        RETURN

