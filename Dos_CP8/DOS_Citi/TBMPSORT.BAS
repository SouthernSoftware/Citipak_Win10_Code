DEFINT A-Z
TYPE Struct
   who AS STRING * 14
   recnum AS INTEGER
END TYPE

DECLARE SUB SortT (SEG Element AS Struct, NumEls, Dir, SSize, MOff, MSize)
'$INCLUDE: 'LCTax.bi'
CLEAR , , 4000          'the sort routines need the extra stack space

REDIM TBRECust(1) AS TBRECust
taxtype$ = "R"
'TBPath$ = "F:\"         'Change to F:\ when compiling"
TBPath$ = "c:\qb45\lc\"         'Change to F:\ when compiling"

CLS
Size = 8000
Start = 1               'start at array element 1
Dir = 0                 'sort direction - use anything else for descending
SSize = 16               'total size of each TYPE element
MOff = 0                'offset into the TYPE for the key element
MSize = 14              'size of the key element - coded as follows:
                        '   -1 = integer
                        '   -2 = long integer
                        '   -3 = single precision
                        '   -4 = double precision
                        '   +N = TYPE array/fixed-length string of length N

REDIM array(1 TO Size) AS Struct

CLS
COLOR 11
PRINT "Tax System : Map Sort Utility"
GOSUB 60000
FOR x = 1 TO LOF(TBFile) / LEN(TBRECust(1))
GET 1, x, TBRECust(1)

 IF TBRECust(1).Deleted <> "Y" THEN
  IF VAL(LEFT$(TBRECust(1).MapNumber, 3)) <= 40 THEN
    COUNT = COUNT + 1
    array(COUNT).who = TBRECust(1).MapNumber
    array(COUNT).recnum = x
    END IF
 END IF
NEXT x


SortT array(Start), COUNT, Dir, SSize, MOff, MSize

OPEN TBPath$ + "TBMAP.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #14 LEN = 24
FIELD 14, 20 AS lastnam$, 4 AS mrec$
CLOSE 14
FileNam$ = TBPath$ + "TBMAP.Dat"
KILL FileNam$
OPEN TBPath$ + "TBMAP.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #14 LEN = 24
FIELD 14, 20 AS lastnam$, 4 AS mrec$


FOR x = 1 TO COUNT
   LSET lastnam$ = array(x).who
    LSET mrec$ = MKS$(array(x).recnum)
    PUT 14, LOF(14) / 24 + 1
NEXT
CLOSE
RUN "tbmap1"

60000  IF taxtype$ = "R" THEN
         TBFile = FREEFILE
         OPEN TBPath$ + "TBRECUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
        ELSE
         TBFile = FREEFILE
         OPEN TBPath$ + "TBPPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
         END IF
        RETURN

