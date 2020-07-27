DECLARE SUB OpenARCustFile (NumOfArRecs%, ARFile%)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DEFINT A-Z

SUB SETFEE

  LOCATE , , 0
  ShowProcessingScrn ("Processing Fees")

  REDIM ARCatCodeRec(1) AS ARNewCatCodeRecType
  ARCatCodeRecLen = LEN(ARCatCodeRec(1))
  ARCatFile = FREEFILE
  OPEN "ARCODE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARCatFile LEN = ARCatCodeRecLen
  NumOFARCatRecs = LOF(ARCatFile) \ ARCatCodeRecLen
  
  OpenARCustFile NumOfArRecs, ARFile
  REDIM ARIdxRec(1) AS ARCustIDXRecType
  IdxFile = FREEFILE
  OPEN "ARCUST.IDX" FOR RANDOM SHARED AS #IdxFile LEN = LEN(ARIdxRec(1))
  
  IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN NumOfArRecs = LOF(IdxFile) \ LEN(ARIdxRec(1))
  
  FOR Cnt! = 1 TO NumOfArRecs
    IF LEFT$(UCASE$(Order$), 4) = "NAME" THEN
      GET IdxFile, Cnt!, ARIdxRec(1)
      CustomerNumber = ARIdxRec(1).IDXRECORD
    ELSE
      CustomerNumber = Cnt!
    END IF

    GET ARFile, CustomerNumber, ARCustRec(1)
    ProRateFlag = False

    ARCustRec(1).Fee1 = 0
    ARCustRec(1).Fee2 = 0
    ARCustRec(1).Fee3 = 0
    ARCustRec(1).Fee4 = 0
    ARCustRec(1).Fee5 = 0
    PUT ARFile, CustomerNumber, ARCustRec(1)
    
    'Help$ = ARCustRec(1).BILLNAME
    'PrintHelp Help$
    'Calc Fee Amount Here

    IF ARCustRec(1).Deleted = "Y" OR ARCustRec(1).Inactive = "Y" THEN
      GOTO SkipEm
    END IF

    ProRate# = ARCustRec(1).ProRate
    IF ProRate# > 100 OR ProRate# < 0 THEN
      ProRate# = 100
    END IF
    ProAT# = ProRate#
    IF ProRate# < 100 THEN
      ProRateFlag = True
      ProRate# = Round#(ProRate# * .01)
    END IF


    CATCODE$ = QPTrim$(ARCustRec(1).BILLCAT1)
    IF LEN(CATCODE$) > 0 THEN
      FOR Snt& = 1 TO NumOFARCatRecs
        GET ARCatFile, Snt&, ARCatCodeRec(1)
        IF QPTrim$(ARCatCodeRec(1).CATCODE) = CATCODE$ THEN
          ARCustRec(1).Desc1 = ARCatCodeRec(1).CODEDESC         'Reset Code Descriptions
          IF ARCatCodeRec(1).CodeType = "F" THEN
            ARCustRec(1).Fee1 = ARCatCodeRec(1).Fee
            FeeAmt# = ARCatCodeRec(1).Fee
            GOSUB PrintLine
            GOTO C2
          END IF
          IF ARCatCodeRec(1).CodeType = "M" THEN
            Mult = ARCustRec(1).REV1
            ARCustRec(1).Fee1 = Mult * ARCatCodeRec(1).Fee
            FeeAmt# = Mult * ARCatCodeRec(1).Fee
            GOSUB PrintLine
            GOTO C2
          END IF
          IF ARCatCodeRec(1).CodeType = "S" THEN
            Revenue# = ARCustRec(1).REV1
            IF Revenue# <= ARCatCodeRec(1).Recpt1 THEN
              ARCustRec(1).Fee1 = ARCatCodeRec(1).BaseAmt1 + (ARCatCodeRec(1).Percent1 / 100) * (Revenue# - ARCatCodeRec(1).Maximum1)
              FeeAmt# = ARCatCodeRec(1).BaseAmt1 + (ARCatCodeRec(1).Percent1 / 100) * (Revenue# - ARCatCodeRec(1).Maximum1)
              GOSUB PrintLine
              GOTO C2
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt2 THEN
              ARCustRec(1).Fee1 = ARCatCodeRec(1).BaseAmt2 + (ARCatCodeRec(1).Percent2 / 100) * (Revenue# - ARCatCodeRec(1).Maximum2)
              FeeAmt# = ARCatCodeRec(1).BaseAmt2 + (ARCatCodeRec(1).Percent2 / 100) * (Revenue# - ARCatCodeRec(1).Maximum2)
              GOSUB PrintLine
              GOTO C2
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt3 THEN
              ARCustRec(1).Fee1 = ARCatCodeRec(1).BaseAmt3 + (ARCatCodeRec(1).Percent3 / 100) * (Revenue# - ARCatCodeRec(1).Maximum3)
              FeeAmt# = ARCatCodeRec(1).BaseAmt3 + (ARCatCodeRec(1).Percent3 / 100) * (Revenue# - ARCatCodeRec(1).Maximum3)
              GOSUB PrintLine
              GOTO C2
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt4 THEN
              ARCustRec(1).Fee1 = ARCatCodeRec(1).BaseAmt4 + (ARCatCodeRec(1).Percent4 / 100) * (Revenue# - ARCatCodeRec(1).Maximum4)
              FeeAmt# = ARCatCodeRec(1).BaseAmt4 + (ARCatCodeRec(1).Percent4 / 100) * (Revenue# - ARCatCodeRec(1).Maximum4)
              GOSUB PrintLine
              GOTO C2
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt5 THEN
              ARCustRec(1).Fee1 = ARCatCodeRec(1).BaseAmt5 + (ARCatCodeRec(1).Percent5 / 100) * (Revenue# - ARCatCodeRec(1).Maximum5)
              FeeAmt# = ARCatCodeRec(1).BaseAmt5 + (ARCatCodeRec(1).Percent5 / 100) * (Revenue# - ARCatCodeRec(1).Maximum5)
              GOSUB PrintLine
              GOTO C2
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt6 THEN
              ARCustRec(1).Fee1 = ARCatCodeRec(1).BaseAmt6 + (ARCatCodeRec(1).Percent6 / 100) * (Revenue# - ARCatCodeRec(1).Maximum6)
              FeeAmt# = ARCatCodeRec(1).BaseAmt6 + (ARCatCodeRec(1).Percent6 / 100) * (Revenue# - ARCatCodeRec(1).Maximum6)
              GOSUB PrintLine
              GOTO C2
            END IF
            
          END IF
        END IF  'End Test for Code
      NEXT Snt&
    ELSE
      ARCustRec(1).Fee1 = 0
    END IF      'End Test for Cat 1
    
    
C2:             'Catagory #2
    
    CATCODE$ = QPTrim$(ARCustRec(1).BILLCAT2)
    IF LEN(CATCODE$) > 0 THEN
      FOR Snt& = 1 TO NumOFARCatRecs
        GET ARCatFile, Snt&, ARCatCodeRec(1)
        IF QPTrim$(ARCatCodeRec(1).CATCODE) = CATCODE$ THEN
          IF ARCatCodeRec(1).CodeType = "F" THEN
            ARCustRec(1).Fee2 = ARCatCodeRec(1).Fee
            FeeAmt# = ARCatCodeRec(1).Fee
            GOSUB PrintLine
            GOTO C3
          END IF
          IF ARCatCodeRec(1).CodeType = "M" THEN
            Mult = ARCustRec(1).REV2
            ARCustRec(1).Fee2 = Mult * ARCatCodeRec(1).Fee
            FeeAmt# = Mult * ARCatCodeRec(1).Fee
            GOSUB PrintLine
            GOTO C3
          END IF
          IF ARCatCodeRec(1).CodeType = "S" THEN
            Revenue# = ARCustRec(1).REV2
            IF Revenue# <= ARCatCodeRec(1).Recpt1 THEN
              ARCustRec(1).Fee2 = ARCatCodeRec(1).BaseAmt1 + (ARCatCodeRec(1).Percent1 / 100) * (Revenue# - ARCatCodeRec(1).Maximum1)
              FeeAmt# = ARCatCodeRec(1).BaseAmt1 + (ARCatCodeRec(1).Percent1 / 100) * (Revenue# - ARCatCodeRec(1).Maximum1)
              GOSUB PrintLine
              GOTO C3
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt2 THEN
              ARCustRec(1).Fee2 = ARCatCodeRec(1).BaseAmt2 + (ARCatCodeRec(1).Percent2 / 100) * (Revenue# - ARCatCodeRec(1).Maximum2)
              FeeAmt# = ARCatCodeRec(1).BaseAmt2 + (ARCatCodeRec(1).Percent2 / 100) * (Revenue# - ARCatCodeRec(1).Maximum2)
              GOSUB PrintLine
              GOTO C3
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt3 THEN
              ARCustRec(1).Fee2 = ARCatCodeRec(1).BaseAmt3 + (ARCatCodeRec(1).Percent3 / 100) * (Revenue# - ARCatCodeRec(1).Maximum3)
              FeeAmt# = ARCatCodeRec(1).BaseAmt3 + (ARCatCodeRec(1).Percent3 / 100) * (Revenue# - ARCatCodeRec(1).Maximum3)
              GOSUB PrintLine
              GOTO C3
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt4 THEN
              ARCustRec(1).Fee2 = ARCatCodeRec(1).BaseAmt4 + (ARCatCodeRec(1).Percent4 / 100) * (Revenue# - ARCatCodeRec(1).Maximum4)
              FeeAmt# = ARCatCodeRec(1).BaseAmt4 + (ARCatCodeRec(1).Percent4 / 100) * (Revenue# - ARCatCodeRec(1).Maximum4)
              GOSUB PrintLine
              GOTO C3
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt5 THEN
              ARCustRec(1).Fee2 = ARCatCodeRec(1).BaseAmt5 + (ARCatCodeRec(1).Percent5 / 100) * (Revenue# - ARCatCodeRec(1).Maximum5)
              FeeAmt# = ARCatCodeRec(1).BaseAmt5 + (ARCatCodeRec(1).Percent5 / 100) * (Revenue# - ARCatCodeRec(1).Maximum5)
              GOSUB PrintLine
              GOTO C3
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt6 THEN
              ARCustRec(1).Fee2 = ARCatCodeRec(1).BaseAmt6 + (ARCatCodeRec(1).Percent6 / 100) * (Revenue# - ARCatCodeRec(1).Maximum6)
              FeeAmt# = ARCatCodeRec(1).BaseAmt6 + (ARCatCodeRec(1).Percent6 / 100) * (Revenue# - ARCatCodeRec(1).Maximum6)
              GOSUB PrintLine
              GOTO C3
            END IF
          END IF
        END IF  'End Test for Code
      NEXT Snt&
    ELSE
      ARCustRec(1).Fee2 = 0
    END IF      'End Test for Cat 1
    
    
C3:
    CATCODE$ = QPTrim$(ARCustRec(1).BILLCAT3)
    IF LEN(CATCODE$) > 0 THEN
      FOR Snt& = 1 TO NumOFARCatRecs
        GET ARCatFile, Snt&, ARCatCodeRec(1)
        IF QPTrim$(ARCatCodeRec(1).CATCODE) = CATCODE$ THEN
          IF ARCatCodeRec(1).CodeType = "F" THEN
            ARCustRec(1).Fee3 = ARCatCodeRec(1).Fee
            FeeAmt# = ARCatCodeRec(1).Fee
            GOSUB PrintLine
            GOTO c4
          END IF
          IF ARCatCodeRec(1).CodeType = "M" THEN
            Mult = ARCustRec(1).REV3
            ARCustRec(1).Fee3 = Mult * ARCatCodeRec(1).Fee
            FeeAmt# = Mult * ARCatCodeRec(1).Fee
            GOSUB PrintLine
            GOTO c4
          END IF
          IF ARCatCodeRec(1).CodeType = "S" THEN
            Revenue# = ARCustRec(1).REV3
            IF Revenue# <= ARCatCodeRec(1).Recpt1 THEN
              ARCustRec(1).Fee3 = ARCatCodeRec(1).BaseAmt1 + (ARCatCodeRec(1).Percent1 / 100) * (Revenue# - ARCatCodeRec(1).Maximum1)
              FeeAmt# = ARCatCodeRec(1).BaseAmt1 + (ARCatCodeRec(1).Percent1 / 100) * (Revenue# - ARCatCodeRec(1).Maximum1)
              GOSUB PrintLine
              GOTO c4
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt2 THEN
              ARCustRec(1).Fee3 = ARCatCodeRec(1).BaseAmt2 + (ARCatCodeRec(1).Percent2 / 100) * (Revenue# - ARCatCodeRec(1).Maximum2)
              FeeAmt# = ARCatCodeRec(1).BaseAmt2 + (ARCatCodeRec(1).Percent2 / 100) * (Revenue# - ARCatCodeRec(1).Maximum2)
              GOSUB PrintLine
              GOTO c4
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt3 THEN
              ARCustRec(1).Fee3 = ARCatCodeRec(1).BaseAmt3 + (ARCatCodeRec(1).Percent3 / 100) * (Revenue# - ARCatCodeRec(1).Maximum3)
              FeeAmt# = ARCatCodeRec(1).BaseAmt3 + (ARCatCodeRec(1).Percent3 / 100) * (Revenue# - ARCatCodeRec(1).Maximum3)
              GOSUB PrintLine
              GOTO c4
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt4 THEN
              ARCustRec(1).Fee3 = ARCatCodeRec(1).BaseAmt4 + (ARCatCodeRec(1).Percent4 / 100) * (Revenue# - ARCatCodeRec(1).Maximum4)
              FeeAmt# = ARCatCodeRec(1).BaseAmt4 + (ARCatCodeRec(1).Percent4 / 100) * (Revenue# - ARCatCodeRec(1).Maximum4)
              GOSUB PrintLine
              GOTO c4
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt5 THEN
              ARCustRec(1).Fee3 = ARCatCodeRec(1).BaseAmt5 + (ARCatCodeRec(1).Percent5 / 100) * (Revenue# - ARCatCodeRec(1).Maximum5)
              FeeAmt# = ARCatCodeRec(1).BaseAmt5 + (ARCatCodeRec(1).Percent5 / 100) * (Revenue# - ARCatCodeRec(1).Maximum5)
              GOSUB PrintLine
              GOTO c4
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt6 THEN
              ARCustRec(1).Fee3 = ARCatCodeRec(1).BaseAmt6 + (ARCatCodeRec(1).Percent6 / 100) * (Revenue# - ARCatCodeRec(1).Maximum6)
              FeeAmt# = ARCatCodeRec(1).BaseAmt6 + (ARCatCodeRec(1).Percent6 / 100) * (Revenue# - ARCatCodeRec(1).Maximum6)
              GOSUB PrintLine
              GOTO c4
            END IF
          END IF
        END IF  'End Test for Code
      NEXT Snt&
    ELSE
      ARCustRec(1).Fee3 = 0
    END IF      'End Test for Cat 3
    
c4:
    CATCODE$ = QPTrim$(ARCustRec(1).BILLCAT4)
    IF LEN(CATCODE$) > 0 THEN
      FOR Snt& = 1 TO NumOFARCatRecs
        GET ARCatFile, Snt&, ARCatCodeRec(1)
        IF QPTrim$(ARCatCodeRec(1).CATCODE) = CATCODE$ THEN
          IF ARCatCodeRec(1).CodeType = "F" THEN
            ARCustRec(1).Fee4 = ARCatCodeRec(1).Fee
            FeeAmt# = ARCatCodeRec(1).Fee
            GOSUB PrintLine
            GOTO c5
          END IF
          IF ARCatCodeRec(1).CodeType = "M" THEN
            Mult = ARCustRec(1).REV4
            ARCustRec(1).Fee4 = Mult * ARCatCodeRec(1).Fee
            FeeAmt# = Mult * ARCatCodeRec(1).Fee
            GOSUB PrintLine
            GOTO c5
          END IF
          IF ARCatCodeRec(1).CodeType = "S" THEN
            Revenue# = ARCustRec(1).REV4
            IF Revenue# <= ARCatCodeRec(1).Recpt1 THEN
              ARCustRec(1).Fee4 = ARCatCodeRec(1).BaseAmt1 + (ARCatCodeRec(1).Percent1 / 100) * (Revenue# - ARCatCodeRec(1).Maximum1)
              FeeAmt# = ARCatCodeRec(1).BaseAmt1 + (ARCatCodeRec(1).Percent1 / 100) * (Revenue# - ARCatCodeRec(1).Maximum1)
              GOSUB PrintLine
              GOTO c5
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt2 THEN
              ARCustRec(1).Fee4 = ARCatCodeRec(1).BaseAmt2 + (ARCatCodeRec(1).Percent2 / 100) * (Revenue# - ARCatCodeRec(1).Maximum2)
              FeeAmt# = ARCatCodeRec(1).BaseAmt2 + (ARCatCodeRec(1).Percent2 / 100) * (Revenue# - ARCatCodeRec(1).Maximum2)
              GOSUB PrintLine
              GOTO c5
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt3 THEN
              ARCustRec(1).Fee4 = ARCatCodeRec(1).BaseAmt3 + (ARCatCodeRec(1).Percent3 / 100) * (Revenue# - ARCatCodeRec(1).Maximum3)
              FeeAmt# = ARCatCodeRec(1).BaseAmt3 + (ARCatCodeRec(1).Percent3 / 100) * (Revenue# - ARCatCodeRec(1).Maximum3)
              GOSUB PrintLine
              GOTO c5
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt4 THEN
              ARCustRec(1).Fee4 = ARCatCodeRec(1).BaseAmt4 + (ARCatCodeRec(1).Percent4 / 100) * (Revenue# - ARCatCodeRec(1).Maximum4)
              FeeAmt# = ARCatCodeRec(1).BaseAmt4 + (ARCatCodeRec(1).Percent4 / 100) * (Revenue# - ARCatCodeRec(1).Maximum4)
              GOSUB PrintLine
              GOTO c5
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt5 THEN
              ARCustRec(1).Fee4 = ARCatCodeRec(1).BaseAmt5 + (ARCatCodeRec(1).Percent5 / 100) * (Revenue# - ARCatCodeRec(1).Maximum5)
              FeeAmt# = ARCatCodeRec(1).BaseAmt5 + (ARCatCodeRec(1).Percent5 / 100) * (Revenue# - ARCatCodeRec(1).Maximum5)
              GOSUB PrintLine
              GOTO c5
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt6 THEN
              ARCustRec(1).Fee4 = ARCatCodeRec(1).BaseAmt6 + (ARCatCodeRec(1).Percent6 / 100) * (Revenue# - ARCatCodeRec(1).Maximum6)
              FeeAmt# = ARCatCodeRec(1).BaseAmt6 + (ARCatCodeRec(1).Percent6 / 100) * (Revenue# - ARCatCodeRec(1).Maximum6)
              GOSUB PrintLine
              GOTO c5
            END IF
            
          END IF
        END IF  'End Test for Code
      NEXT Snt&
    ELSE
      ARCustRec(1).Fee4 = 0
    END IF      'End Test for Cat 1
    
c5:
    CATCODE$ = QPTrim$(ARCustRec(1).BILLCAT5)
    IF LEN(CATCODE$) > 0 THEN
      FOR Snt& = 1 TO NumOFARCatRecs
        GET ARCatFile, Snt&, ARCatCodeRec(1)
        IF QPTrim$(ARCatCodeRec(1).CATCODE) = CATCODE$ THEN
          IF ARCatCodeRec(1).CodeType = "F" THEN
            ARCustRec(1).Fee5 = ARCatCodeRec(1).Fee
            FeeAmt# = ARCatCodeRec(1).Fee
            GOSUB PrintLine
            GOTO SkipEm
          END IF
          IF ARCatCodeRec(1).CodeType = "M" THEN
            Mult = ARCustRec(1).REV5
            ARCustRec(1).Fee5 = Mult * ARCatCodeRec(1).Fee
            FeeAmt# = Mult * ARCatCodeRec(1).Fee
            GOSUB PrintLine
            GOTO SkipEm
          END IF
          IF ARCatCodeRec(1).CodeType = "S" THEN
            Revenue# = ARCustRec(1).REV5
            IF Revenue# <= ARCatCodeRec(1).Recpt1 THEN
              ARCustRec(1).Fee5 = ARCatCodeRec(1).BaseAmt1 + (ARCatCodeRec(1).Percent1 / 100) * (Revenue# - ARCatCodeRec(1).Maximum1)
              FeeAmt# = ARCatCodeRec(1).BaseAmt1 + (ARCatCodeRec(1).Percent1 / 100) * (Revenue# - ARCatCodeRec(1).Maximum1)
              GOSUB PrintLine
              GOTO SkipEm
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt2 THEN
              ARCustRec(1).Fee5 = ARCatCodeRec(1).BaseAmt2 + (ARCatCodeRec(1).Percent2 / 100) * (Revenue# - ARCatCodeRec(1).Maximum2)
              FeeAmt# = ARCatCodeRec(1).BaseAmt2 + (ARCatCodeRec(1).Percent2 / 100) * (Revenue# - ARCatCodeRec(1).Maximum2)
              GOSUB PrintLine
              GOTO SkipEm
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt3 THEN
              ARCustRec(1).Fee5 = ARCatCodeRec(1).BaseAmt3 + (ARCatCodeRec(1).Percent3 / 100) * (Revenue# - ARCatCodeRec(1).Maximum3)
              FeeAmt# = ARCatCodeRec(1).BaseAmt3 + (ARCatCodeRec(1).Percent3 / 100) * (Revenue# - ARCatCodeRec(1).Maximum3)
              GOSUB PrintLine
              GOTO SkipEm
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt4 THEN
              ARCustRec(1).Fee5 = ARCatCodeRec(1).BaseAmt4 + (ARCatCodeRec(1).Percent4 / 100) * (Revenue# - ARCatCodeRec(1).Maximum4)
              FeeAmt# = ARCatCodeRec(1).BaseAmt4 + (ARCatCodeRec(1).Percent4 / 100) * (Revenue# - ARCatCodeRec(1).Maximum4)
              GOSUB PrintLine
              GOTO SkipEm
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt5 THEN
              ARCustRec(1).Fee5 = ARCatCodeRec(1).BaseAmt5 + (ARCatCodeRec(1).Percent5 / 100) * (Revenue# - ARCatCodeRec(1).Maximum5)
              FeeAmt# = ARCatCodeRec(1).BaseAmt5 + (ARCatCodeRec(1).Percent5 / 100) * (Revenue# - ARCatCodeRec(1).Maximum5)
              GOSUB PrintLine
              GOTO SkipEm
            END IF
            IF Revenue# <= ARCatCodeRec(1).Recpt6 THEN
              ARCustRec(1).Fee5 = ARCatCodeRec(1).BaseAmt6 + (ARCatCodeRec(1).Percent6 / 100) * (Revenue# - ARCatCodeRec(1).Maximum6)
              FeeAmt# = ARCatCodeRec(1).BaseAmt6 + (ARCatCodeRec(1).Percent6 / 100) * (Revenue# - ARCatCodeRec(1).Maximum6)
              GOSUB PrintLine
              GOTO SkipEm
            END IF
            
          END IF
        END IF  'End Test for Code
      NEXT Snt&
    ELSE
      ARCustRec(1).Fee5 = 0
    END IF      'End Test for Cat 1
    
SkipEm:

    'ARCustRec(1).LicBal = Round#(ARCustRec(1).LicBal + FeeAmt#)
    'ARCustRec(1).AcctBal = Round#(ARCustRec(1).AcctBal + FeeAmt#)
    PUT ARFile, CustomerNumber, ARCustRec(1)
    ShowPctComp Cnt!, NumOfArRecs
  NEXT Cnt!
  
  GOSUB PrintSetFeeRptEnding
  CLOSE         'Close all open files now
  
  EXIT SUB
  
  
PrintLine:
RETURN
  
PrintSetFeeRptHeader:
RETURN
  
PrintSetFeeRptEnding:
RETURN

  
  
END SUB

