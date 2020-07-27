DEFINT A-Z

SUB UBLog (Text$) ', LogFlags%)
  LogFile = FREEFILE
  OPEN "UBLOG.DAT" FOR APPEND SHARED AS LogFile LEN = 255
  PRINT #LogFile, "UBLOG: "; DATE$; " @"; TIME$
  PRINT #LogFile, "       "; Text$
  CLOSE #LogFile
END SUB

