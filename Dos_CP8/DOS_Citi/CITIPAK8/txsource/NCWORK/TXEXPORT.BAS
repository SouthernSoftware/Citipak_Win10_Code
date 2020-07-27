DEFINT A-Z
DECLARE SUB CustHistoryRpt ()
DECLARE SUB ClearBack ()
DECLARE SUB DisplayTaxScrn (ScrnName$)
DECLARE SUB LookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, SSNFlag%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Round# (DblNumber#)
DECLARE FUNCTION ParseBillNum$ (Text$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB RestScrn (Array%())
DECLARE SUB SaveScrn (Array%())
DECLARE SUB ShowCursor ()
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE FUNCTION Monitor% ()

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TaxCONST.BI'
  '$INCLUDE: 'PROPAbst.BI'

  CONST False = 0, True = NOT False

  CrLf$ = CHR$(13) + CHR$(10)

'  TYPE HistRecInfoType
'    TranRec    AS LONG
'    TranType   AS INTEGER
'    TranDate   AS INTEGER
'    BelongTo   AS LONG
'    Printed    AS INTEGER
'  END TYPE

  REDIM TaxSetupRec(1) AS TaxMasterType
  REDIM TaxCustRec(1) AS TaxCustType
  REDIM PropertyRec(1) AS PropertyRecType
  REDIM PersRec(1) AS PersonalRecType


  'REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)

  LineCnt = 0
  CustCnt = 0

  Action = 1
'  ClearBack
'
'  ShowCursor

'  DisplayTaxScrn ScrnName$

  UsingIndex = False
  DetailFlag = True

  CLS
  PRINT
  PRINT "Exporting Tax Information. . ."
  FileName$ = "TXEXPORT.DAT"

  Oh = FREEFILE

  OPEN FileName$ FOR OUTPUT AS #Oh

  GOSUB OpenTaxFile
  GOSUB OpenPropFile
  GOSUB OpenPersFile


  FOR Cnt = 1 TO NumOfTaxRecs
    LOCATE 4, 1
    PRINT "Processing:"; Cnt; "of"; NumOfTaxRecs;
    CustRecNo = Cnt

    GET TaxFile, CustRecNo, TaxCustRec(1)

    IF TaxCustRec(1).Deleted = 0 THEN
    ELSE
      TaxCustRec(1).Deleted = -1
      PUT TaxFile, CustRecNo, TaxCustRec(1)
    END IF

    IF NOT TaxCustRec(1).Deleted THEN

        PRINT #Oh, QPTrim$(STR$(CustRecNo));
        PRINT #Oh, ",";
        PRINT #Oh, QPTrim$(TaxCustRec(1).FName); ",";
        PRINT #Oh, QPTrim$(TaxCustRec(1).LNAME); ",";
        PRINT #Oh, QPTrim$(TaxCustRec(1).Addr1); ",";
        PRINT #Oh, QPTrim$(TaxCustRec(1).Addr2); ",";
        PRINT #Oh, QPTrim$(TaxCustRec(1).City); ",";
        PRINT #Oh, QPTrim$(TaxCustRec(1).State); ",";
        PRINT #Oh, QPTrim$(TaxCustRec(1).Zip); ",";
        PRINT #Oh, QPTrim$(TaxCustRec(1).CSSN);
        IF DetailFlag = False THEN
          PRINT #Oh, ""
          GOTO SkipEm
        ELSE
          PRINT #Oh, ",";
        END IF
        'Now Show Property Records Next

        IF TaxCustRec(1).FirstPropRec > 0 THEN

          PropertyRecord! = TaxCustRec(1).FirstPropRec

          WHILE PropertyRecord! <> 0


            GET #PropTaxFile, PropertyRecord!, PropertyRec(1)

            IF PropertyRecord! <> TaxCustRec(1).FirstPropRec THEN
                PRINT #Oh, QPTrim$(STR$(CustRecNo));
                PRINT #Oh, ",";
                PRINT #Oh, QPTrim$(TaxCustRec(1).FName); ",";
                PRINT #Oh, QPTrim$(TaxCustRec(1).LNAME); ",";
                PRINT #Oh, QPTrim$(TaxCustRec(1).Addr1); ",";
                PRINT #Oh, QPTrim$(TaxCustRec(1).Addr2); ",";
                PRINT #Oh, QPTrim$(TaxCustRec(1).City); ",";
                PRINT #Oh, QPTrim$(TaxCustRec(1).State); ",";
                PRINT #Oh, QPTrim$(TaxCustRec(1).Zip); ",";
                PRINT #Oh, QPTrim$(TaxCustRec(1).CSSN); ",";
             ELSE
            END IF

            PRINT #Oh, QPTrim$(PropertyRec(1).REALPIN); ",";
            PRINT #Oh, QPTrim$(PropertyRec(1).PROPNOT1); ",";
            PRINT #Oh, QPTrim$(PropertyRec(1).PROPNOT2); ",";
            PRINT #Oh, QPTrim$(PropertyRec(1).PROPNOT3); ",";
            PRINT #Oh, QPTrim$(STR$(PropertyRec(1).PROPVALU));
             PRINT #Oh, ",";
            PRINT #Oh, QPTrim$(STR$(PropertyRec(1).EXMPSENI));
             PRINT #Oh, ",";
            PRINT #Oh, QPTrim$(PropertyRec(1).MORTCODE); ",";
            PRINT #Oh, QPTrim$(PropertyRec(1).MAP); ",";
            PRINT #Oh, QPTrim$(PropertyRec(1).BLOCK); ",";
            PRINT #Oh, QPTrim$(PropertyRec(1).LOTNUMB); ",";
            PropertyRecord! = PropertyRec(1).NextRec

           WEND

           ELSE
            PRINT #Oh, ","; ","; ","; ","; ","; ","; ","; ","; ","; ",";
           END IF


        'NOW CHECK PERSONAL PROPERTY

        IF TaxCustRec(1).FirstPersRec > 0 THEN


          PropertyRecord! = TaxCustRec(1).FirstPersRec


          WHILE PropertyRecord! <> 0

           GET #PersTaxFile, PropertyRecord!, PersRec(1)

            PRINT #Oh, QPTrim$(PersRec(1).PROPPIN); ",";
            PRINT #Oh, QPTrim$(STR$(PersRec(1).PersVal));
            PRINT #Oh, ",";
            PRINT #Oh, QPTrim$(STR$(PersRec(1).EXMPSENI));
            PRINT #Oh, ",";
            PRINT #Oh, QPTrim$(PersRec(1).DESC1); ",";
            PRINT #Oh, QPTrim$(PersRec(1).DESC2); ",";
            PRINT #Oh, QPTrim$(PersRec(1).DESC3); ",";
            PRINT #Oh, QPTrim$(PersRec(1).Desc4); ",";
            PRINT #Oh, QPTrim$(PersRec(1).Desc5); ","
            PropertyRecord! = PersRec(1).NextRec
            IF OldRecord! = PropertyRecord! THEN PropertyRecord! = 0
          WEND
          ELSE
           PRINT #Oh, ","; ","; ","; ","; ","; ","; ","; ","
        END IF
        END IF
SkipEm:
      CustCnt = CustCnt + 1

  NEXT


  CLOSE         'Close all open files now
PRINT
PRINT
PRINT
PRINT "Export file: "; FileName$
PRINT "Done."



END



OpenTaxFile:
  TaxFile = FREEFILE
  OPEN "TAXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
  RETURN
OpenPersFile:
  PersTaxFile = FREEFILE
  OPEN "TAXPERS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PersTaxFile LEN = LEN(PersRec(1))
  NumOfPersRecs = LOF(PersTaxFile) / LEN(PersRec(1))
  RETURN
OpenPropFile:
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))
  RETURN


