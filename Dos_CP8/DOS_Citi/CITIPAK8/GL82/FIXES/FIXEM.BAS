DEFINT A-Z

'$INCLUDE: 'ap.bi'
''$INCLUDE: 'aporg.bi'

DIM ORec AS oAPInvType
DIM NRec AS APInv81Type

OLen = LEN(ORec)
NLen = LEN(NRec)
END

NAME "APIED.DAT" AS "OAPIED.DAT"

OPEN "OAPIED.dat" FOR RANDOM AS #1 LEN = OLen
OPEN "APIED.DAT" FOR RANDOM AS #2 LEN = NLen

   FOR Cnt = 1 TO LOF(1) \ OLen
   GET #1, Cnt, ORec

   NRec.DELFLAG = ORec.DELFLAG
   NRec.VENDOR = ORec.VENDOR
   NRec.VENDNAME = ORec.VENDNAME
   'NRec.VRecNum = ORec.VRecNum
   NRec.INVNUM = ORec.PONUM
   NRec.PONUM = ORec.INVNUM
   NRec.INVAMT = ORec.INVAMT
   NRec.PAYCODE = ORec.PAYCODE
   NRec.INVDESC = ""
   NRec.INVDATE = ORec.INVDATE
   NRec.DUEDATE = ORec.DUEDATE
   NRec.DISTDATE = ORec.DISTDATE
   NRec.PSLFlag = "N"
   NRec.Get1099 = "N"
   FOR ZZ = 1 TO 24
     LSET NRec.Dist(ZZ) = ORec.Dist(ZZ)
   NEXT
   'NRec.POAPLRecNum = 0
   NRec.POFLAG = 0
   PUT #2, Cnt, NRec
 NEXT
 CLOSE

