DEFINT A-Z


'$INCLUDE: 'ap.bi'

  DIM VRec AS VendorRecType
  VRecLen = LEN(VRec)
  VendorFile = FREEFILE
  OPEN "apvendor.dat" FOR RANDOM ACCESS READ WRITE SHARED AS VendorFile LEN = VRecLen
  NumVRecs = LOF(VendorFile) \ VRecLen

  CLS

  FOR Cnt = 1 TO NumVRecs
    LOCATE 5, 1
    PRINT "Processing:"; Cnt; " of"; NumVRecs;
    GET #VendorFile, Cnt, VRec
    VendName$ = VRec.VName

    s = INSTR(VendName$, "&")
    DO WHILE s > 0
      VendName$ = LEFT$(VendName$, s - 1) + MID$(VendName$, s + 1)
      s = INSTR(VendName$, "&")
    LOOP
    s = INSTR(VendName$, "-")
    DO WHILE s > 0
      VendName$ = LEFT$(VendName$, s - 1) + MID$(VendName$, s + 1)
      s = INSTR(VendName$, "-")
    LOOP
    s = INSTR(VendName$, " ")
    DO WHILE s > 0
      VendName$ = LEFT$(VendName$, s - 1) + MID$(VendName$, s + 1)
      s = INSTR(VendName$, " ")
    LOOP
    s = INSTR(VendName$, ".")
    DO WHILE s > 0
      VendName$ = LEFT$(VendName$, s - 1) + MID$(VendName$, s + 1)
      s = INSTR(VendName$, ".")
    LOOP
    s = INSTR(VendName$, ",")
    DO WHILE s > 0
      VendName$ = LEFT$(VendName$, s - 1) + MID$(VendName$, s + 1)
      s = INSTR(VendName$, ",")
    LOOP
    s = INSTR(VendName$, "'")
    DO WHILE s > 0
      VendName$ = LEFT$(VendName$, s - 1) + MID$(VendName$, s + 1)
      s = INSTR(VendName$, "'")
    LOOP

    VRec.VNum = VendName$
    PUT #VendorFile, Cnt, VRec
  NEXT

  CLOSE

