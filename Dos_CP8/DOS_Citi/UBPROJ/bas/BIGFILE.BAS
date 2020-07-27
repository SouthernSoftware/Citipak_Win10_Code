DEFINT A-Z

TYPE dodo
  text AS STRING * 32000
END TYPE


OPEN "c:\test.bin" FOR BINARY AS #1
a$ = "A"
PUT #1, (1024# * 1024 * 100), a$
CLOSE

