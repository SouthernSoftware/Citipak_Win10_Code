DEFINT A-Z

DEFSNG A-Z
DEFINT A-Z
SUB DisplayH2Scrn (ScrnName$)
  LibFile2Scrn "UBHH", ScrnName$, MonoCode%, Attribute%, ErrorCode%
END SUB

DEFSNG A-Z
DEFINT A-Z
SUB DisplayHHScrn (ScrnName$)
  LibFile2Scrn "UBSETUP", ScrnName$, MonoCode%, Attribute%, ErrorCode%
END SUB

