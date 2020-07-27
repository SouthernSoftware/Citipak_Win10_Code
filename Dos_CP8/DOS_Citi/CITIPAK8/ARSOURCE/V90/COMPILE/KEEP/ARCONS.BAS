DEFINT A-Z
DECLARE SUB ShowNoCodes ()
DECLARE SUB OpenARCustIdxFile (NumOfARIdxRecs%, ARIdxFile%)
DECLARE SUB OpenARCust (NumOfArRecs%, ARFile%)
DECLARE SUB SortARNameIndex ()
DECLARE SUB AddCustomer ()
DECLARE SUB EditCustomer ()
DECLARE SUB PrintCustomer ()
DECLARE SUB printhelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB HideCursor ()
DECLARE SUB QPrint (x$, Colr%, page%)
DECLARE SUB QPrintRC (T$, r%, c%, clr%)
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()

'$INCLUDE: 'DefCnf.BI'
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)

TYPE Struct
 who AS STRING * 14
 RecNum AS INTEGER
END TYPE



  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
'  '$INCLUDE: 'AR.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'ARV85.bi'
  '$INCLUDE: 'GL.bi'
   DIM SHARED ARCustRec(1) AS ARCustRecType
 
   DIM SHARED ARCustIdxRec(1) AS ARCustIDXRecType
   
   STACK 8000
   CONST False = 0, True = NOT False
      
     OpenARCust NumOfArRecs, ARFile
     FOR Rnt& = 1 TO (NumOfArRecs - 1)
     GET ARFile, Rnt&, ARCustRec(1)
     IF ARCustRec(1).Deleted = "Y" THEN GOTO Nexone
     FirstTime = 0
     Cnt& = Rnt&
     WHILE Cnt& <= NumOfArRecs
     GET ARFile, Cnt&, ARCustRec(1)
     IF ARCustRec(1).Deleted = "Y" THEN GOTO NextOne
     IF FirstTime = 0 THEN
      CurrentName$ = QPTrim$(ARCustRec(1).CustName)
      CurRecord& = Cnt&
      FirstTime = 1
      GOTO NextOne
     END IF
     ThisName$ = QPTrim$(ARCustRec(1).CustName)
     ThisCat$ = QPTrim$(ARCustRec(1).BillCat1)
     IF ThisName$ = CurrentName$ THEN
      ARCustRec(1).Deleted = "Y"
      PUT ARFile, Cnt&, ARCustRec(1)
      GET ARFile, CurRecord&, ARCustRec(1)
      IF LEN(QPTrim$(ARCustRec(1).BillCat1)) = 0 THEN
       ARCustRec(1).BillCat1 = ThisCat$
       GOTO EndUpdate
      END IF
      IF LEN(QPTrim$(ARCustRec(1).BillCat2)) = 0 THEN
       ARCustRec(1).BillCat2 = ThisCat$
       GOTO EndUpdate
      END IF
      IF LEN(QPTrim$(ARCustRec(1).BillCat3)) = 0 THEN
       ARCustRec(1).BillCat3 = ThisCat$
       GOTO EndUpdate
      END IF
      IF LEN(QPTrim$(ARCustRec(1).BillCat4)) = 0 THEN
       ARCustRec(1).BillCat4 = ThisCat$
       GOTO EndUpdate
      END IF
      IF LEN(QPTrim$(ARCustRec(1).BillCat5)) = 0 THEN
       ARCustRec(1).BillCat5 = ThisCat$
       GOTO EndUpdate
      END IF
EndUpdate:
      ARCustRec(1).FirstTrans = 0
      ARCustRec(1).LastTrans = 0
      ARCustRec(1).AcctBal = 0
     PUT ARFile, CurRecord&, ARCustRec(1)

    
    ELSE
      ARCustRec(1).FirstTrans = 0
      ARCustRec(1).LastTrans = 0
      ARCustRec(1).AcctBal = 0
       PUT ARFile, Cnt&, ARCustRec(1)

    END IF

NextOne:
    Cnt& = Cnt& + 1
 WEND
Nexone:
NEXT Rnt&
   
     CLOSE

     END

SUB OpenARCust (NumOfArRecs%, ARFile%)
  CLOSE ARFile

  ARCustRecLen = LEN(ARCustRec(1))
  ARFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARFile LEN = ARCustRecLen
  NumOfArRecs = LOF(ARFile) \ ARCustRecLen
  'FOR x = 1 TO NumOfArRecs
  'GET ARFile, x, ARCust(1)
  'PRINT ARCust(1).Custnumb; TAB(15); ARCust(1).FirstTrans
  'SLEEP 1
  'NEXT x
  'STOP

END SUB

SUB OpenARCustIdxFile (NumOfARIdxRecs, ARIdxFile)
  CLOSE ARIdxFile
  ARCustIdxRecLen = LEN(ARCustIdxRec(1))
  ARIdxFile = FREEFILE
  OPEN "ARCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS ARIdxFile LEN = ARCustIdxRecLen
  NumOfARIdxRecs = LOF(ARIdxFile) \ ARCustIdxRecLen
END SUB

