DEFINT A-Z
 '$INCLUDE: 'btaxtran.BI'

 DIM Taxtrans AS TaxTransactionType
 DIM tLen AS INTEGER

 CLS
 PRINT
 tLen = LEN(Taxtrans)
 PRINT "Correcting account 1938. . .";
 OPEN "taxtrans.dat" FOR RANDOM AS #1 LEN = tLen
 GET #1, 850, Taxtrans
 Taxtrans.Revenue.InterestPd = Taxtrans.Revenue.Interest
 PUT #1, 850, Taxtrans
 GET #1, 852, Taxtrans
 Taxtrans.Amount = 0
 PUT #1, 852, Taxtrans
 PRINT "Done"
 PRINT

 PRINT "Correcting account 2105. . .";
 GET #1, 66131, Taxtrans
 Taxtrans.Amount = 0
 PUT #1, 66131, Taxtrans
 CLOSE
 PRINT "Done"
 PRINT
 PRINT
 PRINT "Corrections complete."
 PRINT
 KILL "fixemup.exe"

