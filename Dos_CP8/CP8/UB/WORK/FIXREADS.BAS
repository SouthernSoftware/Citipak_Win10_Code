DEFINT A-Z

DIM dodo AS STRING * 181



OPEN "exssi001.dat" FOR INPUT AS #1

OPEN "dodo.dat" FOR RANDOM AS #2 LEN = 181
DO

LINE INPUT #1, dodo

PUT #2, , dodo

LOOP


