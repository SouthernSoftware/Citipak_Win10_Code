DEFINT A-Z

  '$INCLUDE: 'AP.bi'                        'AP File Types
  
  CLS
  REDIM APDistRec(1) AS APDistRecType
  APDistRecLen = LEN(APDistRec(1))
  
  APDistFile = FREEFILE
  OPEN "APDIST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS APDistFile LEN = APDistRecLen
  NumDistRecs& = LOF(APDistFile) \ APDistRecLen
  GET APDistFile, 22177, APDistRec(1)
  APDistRec(1).DistAmt = 0
  PUT APDistFile, 22177, APDistRec(1)
  CLOSE
  CALL KillFile("fixdist.exe")
  PRINT "Done. . ."

