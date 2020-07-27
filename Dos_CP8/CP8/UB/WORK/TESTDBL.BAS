DEFINT A-Z

TYPE DoubleType
 a AS SINGLE
END TYPE

DIM TestDbl AS DoubleType
TLen = LEN(TestDbl)


OPEN "tstdbl.dat" FOR RANDOM AS #1 LEN = TLen
TestDbl.a = 1001.01
PUT #1, , TestDbl
CLOSE

