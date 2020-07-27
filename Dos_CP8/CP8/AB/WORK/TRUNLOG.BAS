a$ = SPACE$(2048)
OPEN "ublog.dat" FOR BINARY AS #1
OPEN "1meg.dat" FOR BINARY AS #2
fs& = LOF(1)
SEEK #1, fs& - 1000000
DO
  GET #1, , a$
  PUT #2, , a$
LOOP UNTIL EOF(1)
CLOSE

