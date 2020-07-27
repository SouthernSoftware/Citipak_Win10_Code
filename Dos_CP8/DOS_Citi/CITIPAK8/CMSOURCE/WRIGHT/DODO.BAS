DEFINT A-Z

DEFSNG A-Z
DEFINT A-Z
SUB DisplayCMScrn (ScrnName$)
  LibFile2Scrn "CM.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
END SUB

