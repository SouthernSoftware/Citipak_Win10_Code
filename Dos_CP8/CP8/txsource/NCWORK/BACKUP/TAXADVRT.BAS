DEFINT A-Z
DECLARE SUB Abstracts ()
DECLARE SUB ShowCustHistory (CustRec&)
DECLARE SUB CustHistoryRpt ()
DECLARE SUB GetPropRecList (PropRecs() AS LONG, CustRec&)
DECLARE SUB LookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, SSNFlag%)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB AbtractListing ()
DECLARE SUB BalanceListing ()
DECLARE SUB MortgageCodeList ()
DECLARE SUB MasterValueList ()
DECLARE SUB TransactionJournal ()
DECLARE SUB LateListing ()
DECLARE SUB CustomerInquiry ()
DECLARE SUB SrCitizensList ()
DECLARE SUB Labels ()
DECLARE SUB AdListing ()
DECLARE SUB DisplayTaxScrn (ScrnName$)
DECLARE SUB CustomerListing ()
DECLARE SUB TAXCustomerMenu ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB ClearBack ()
DECLARE SUB SendDist2GL ()
DECLARE SUB UBMiscMenu ()
DECLARE SUB UBBillMenu ()
DECLARE SUB UBCustomerMenu ()
DECLARE SUB ClearScrn ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION WEnvTest% ()
DECLARE FUNCTION Round# (B#)

DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB HideCursor ()
DECLARE SUB CursorOff ()
DECLARE SUB TextCursor (MouseFg%, MouseBg%)
DECLARE SUB WaitForAction ()
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION GetCustBalance# (RecNo&)
  
  CONST False = 0, True = NOT False

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TAXCONST.BI'
  '$INCLUDE: 'PROPAbst.BI'
  '$INCLUDE: 'TAXRPTTY.BI'


  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE

  STACK 5000

  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 5)

  MChoice$(1) = "Calculate Advertising Charges"
  MChoice$(2) = "Edit Advertising Transactions"
  MChoice$(3) = "Print Advertising Edit Report "
  MChoice$(4) = "Post Advertising Transactions"
  MChoice$(5) = "Exit to OS"

  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line 18
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight

  '--Find max menu width
  FOR Cnt = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT

  '--Center Menu within Screen
  Row = ((24 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 1

  IF Exist("TAXFROMM.DAT") THEN
    'KillFile "TAXFROMM.DAT"
    MainFlag = True
  END IF

  DO

    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0

    ClearBack

    TitleBox 2, Col, MaxLen + 3, "Tax Advertising Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    ShowCursor

    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0

    SELECT CASE Choice
    CASE 1   'Calc
      AdListing
    CASE 2   'Edit
    CASE 3   'Print Report
    CASE 4   'Post
    CASE 5
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP

  IF WEnvTest THEN
    Ext$ = ".bas"
  ELSE
    Ext$ = ".exe"
  END IF

  IF INSTR(COMMAND$, "TEST") THEN
    HideCursor
    ClearScrn
    END
  END IF

  RUN "TaxMenu"

  END

SUB AdListing

  SHARED Choice$()

  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  REDIM PropDesc$(250)
  ReportFile$ = "TxAdList.PRN"  'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)

  REDIM TaxCustRec(1)  AS TaxCustType
  REDIM PropertyRec(1) AS PropertyRecType
  REDIM TaxTrans(1) AS TaxTransactionType
  REDIM AdvTrans(1) AS InterestRecType

  TAdvLen = LEN(AdvTrans(1))
  
BeginOver:
  MaxLines = 56
  LineCnt = 0
  CustCnt = 0
  
  LibName$ = "TAX"
  ScrnName$ = "ADRPT"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  REDIM Choice$(0 TO 2, 0 TO 2)
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name Order"
  Choice$(2, 0) = "Account Number"
  Choice$(0, 1) = "5"
  Choice$(1, 1) = "Screen"
  Choice$(2, 1) = "Printer"
  Form$(2, 0) = "0"
  Form$(4, 0) = DATE$
  Action = 1
  ClearBack
  
  ShowCursor
  
  DisplayTaxScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      IF LEFT$(Form$(1, 0), 1) = "N" THEN
        UsingIndex = True
      ELSE
        UsingIndex = False
      END IF
      AmountCharged# = Value#(Form$(2, 0), ecode%)
      TaxYear = Value#(Form$(3, 0), ecode%)
      PostDate$ = Form$(4, 0)
      DevSpec$ = LEFT$(Form$(5, 0), 1)
      IF AmountToCharge# <> 0 AND LEN(QPTrim$(PostDate$)) <> 10 THEN
        Action = 1
        GOTO BeginOver
      END IF
      ExitFlag = True
    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF AbortFlag THEN EXIT SUB
  
  RptHandle = FREEFILE
  
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  GOSUB PrintAdRptHeader
  
  TransFile = FREEFILE
  OPEN "TaxTrans.dat" FOR RANDOM AS TransFile LEN = LEN(TaxTrans(1))
  
  TAdvFile = FREEFILE
  OPEN TaxAdvFile FOR RANDOM SHARED AS TAdvFile LEN = TAdvLen

  TaxFile = FREEFILE
  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxCustLen

  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
  
  IF UsingIndex AND NumOfTaxRecs > 0 THEN
    GOSUB GetAdNameIndex
  END IF
  
  ClearBack
  ShowProcessingScrn "Ad/Collection Cost Listing"
  
  FOR Cnt = 1 TO NumOfTaxRecs
    IF UsingIndex THEN
      CustRecNo = Array(Cnt).RecNum
    ELSE
      CustRecNo = Cnt
    END IF
    
    GET TaxFile, CustRecNo, TaxCustRec(1)

SELECT CASE CustRecNo
CASE 1843, 1867, 1869
  STOP
END SELECT

    IF NOT TaxCustRec(1).Deleted THEN
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintAdRptHeader
      END IF
      
      TransRList = 0
      TransRecord& = TaxCustRec(1).LastTrans
      WHILE TransRecord& <> 0
        
        GET TransFile, TransRecord&, TaxTrans(1)
        CurTransRecord& = TransRecord&
        
        IF TaxTrans(1).TranType = 1 AND TaxTrans(1).TaxYear = TaxYear AND TaxCustRec(1).FirstPropRec > 0 THEN
          Balance# = TaxTrans(1).Revenue.Principle1 + TaxTrans(1).Revenue.Principle2 + TaxTrans(1).Revenue.Principle3 + TaxTrans(1).Revenue.Principle4 + TaxTrans(1).Revenue.Principle5
          Balance# = Balance# + TaxTrans(1).Revenue.INTEREST + TaxTrans(1).Revenue.PENALTY + TaxTrans(1).Revenue.Collection
          Balance# = Balance# - (TaxTrans(1).Revenue.Principle1Pd + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd)
          Balance# = Balance# - (TaxTrans(1).Revenue.InterestPd + TaxTrans(1).Revenue.PenaltyPd + TaxTrans(1).Revenue.CollectionPd)
          Balance# = Round#(Balance#)
          TransRList = TransRList + 1
          
          IF Balance# > 0 THEN
            
            'Get Prop Desc Here
            PropRec& = TaxCustRec(1).FirstPropRec
            PropTaxFile = FREEFILE
            OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
            NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))

            HowMany = 0
            z1 = 0
            WHILE PropRec& <> 0
              GET PropTaxFile, PropRec&, PropertyRec(1)
              z1 = z1 + 1
              'PropDesc$(Z1) = QPTrim$(PropertyRec(1).PROPNOT1)
              PropDesc$(z1) = QPTrim$(PropertyRec(1).MAP) + "\" + QPTrim$(PropertyRec(1).BLOCK) + "\" + QPTrim$(PropertyRec(1).LOTNUMB) + "  " + QPTrim$(PropertyRec(1).PROPNOT1)
              PropRec& = PropertyRec(1).NextRec
            WEND
            IF TransRList = 1 THEN
              PropDesc$ = PropDesc$(z1)
            END IF
            IF TransRList > 1 THEN
              z2 = (z1 - TransRList) + 1
              IF z2 < 0 THEN z2 = 1
              PropDesc$ = PropDesc$(z2)
            END IF
            
            NME$ = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
            NME$ = QPTrim$(NME$)                'this one cleans up those with only last name
            
            PRINT #RptHandle, LEFT$(NME$, 30);
            PRINT #RptHandle, TAB(32); LEFT$(PropDesc$, 30);
            PRINT #RptHandle, TAB(71); USING "######.##"; Balance#

            LineCnt = LineCnt + 2

            TotalCharged# = TotalCharged# + AmountCharged#
            
            CLOSE PropTaxFile
            
            IF AmountCharged# <> 0 THEN
              ' Now Add Amt to Bill and Put Back
              TaxTrans(1).Revenue.Collection = TaxTrans(1).Revenue.Collection + AmountCharged#
'unrem
'              PUT TransFile, TransRecord&, TaxTrans(1)

              TaxTrans(1).TransDate = Date2Num%(PostDate$)
              TaxTrans(1).TaxYear = TaxYear
              TaxTrans(1).TranType = 6          '1=Bill 2=Payment 3=Release 4=Interest 5=Penalty 6=Collection/Ad Cost Billing
              TaxTrans(1).BillType = "C"        'R=Real P=Personal Property C=Combined (NC/GA)
              TaxTrans(1).Amount = AmountCharged#               'Total Transaction Amount
              TaxTrans(1).Revenue.Principle1 = 0
              TaxTrans(1).Revenue.Principle2 = 0
              TaxTrans(1).Revenue.Principle3 = 0
              TaxTrans(1).Revenue.Principle4 = 0
              TaxTrans(1).Revenue.Principle5 = 0
              TaxTrans(1).Revenue.INTEREST = 0
              TaxTrans(1).Revenue.PENALTY = 0
              TaxTrans(1).Revenue.Collection = AmountCharged#
              TaxTrans(1).Revenue.Future1 = 0
              TaxTrans(1).Revenue.Future2 = 0
              TaxTrans(1).Revenue.Principle1Pd = 0
              TaxTrans(1).Revenue.Principle2Pd = 0
              TaxTrans(1).Revenue.Principle3Pd = 0
              TaxTrans(1).Revenue.Principle4Pd = 0
              TaxTrans(1).Revenue.Principle5Pd = 0
              TaxTrans(1).Revenue.InterestPd = 0
              TaxTrans(1).Revenue.PenaltyPd = 0
              TaxTrans(1).Revenue.CollectionPd = 0
              TaxTrans(1).Revenue.Future1Pd = 0
              TaxTrans(1).Revenue.Future2Pd = 0
              
              TaxTrans(1).Description = "Collection/Ad Charge"
              TaxTrans(1).Posted2GL = "N"
              TaxTrans(1).CustomerRec = CustRecNo
              TaxTrans(1).LastTrans = 0
              TaxTrans(1).BelongTo = CurTransRecord&
              TaxTrans(1).Padding = ""
              
              'Increment Transaction File Record Count
              NextRecord& = (LOF(TransFile) / LEN(TaxTrans(1))) + 1
'unrem
'              PUT #TransFile, NextRecord&, TaxTrans(1)
              
              'Update the Customer Pointers Now
              GET TaxFile, CustRecNo, TaxCustRec(1)
              
              IF TaxCustRec(1).LastTrans = 0 THEN
                TaxCustRec(1).LastTrans = NextRecord&
'unrem
'                PUT #TaxFile, CustRecNo, TaxCustRec(1)
              ELSE
                Previous& = TaxCustRec(1).LastTrans
                TaxCustRec(1).LastTrans = NextRecord&
'unrem
'                PUT #TaxFile, CustRecNo, TaxCustRec(1)
                
                GET #TransFile, NextRecord&, TaxTrans(1)
                TaxTrans(1).LastTrans = Previous&
'unrem
'                PUT #TransFile, NextRecord&, TaxTrans(1)
              END IF
            END IF
            'Now Grab Back the Original Transaction
            'Holding the Bill to Get the Next One
            GET TransFile, CurTransRecord&, TaxTrans(1)
            Balance# = 0
          END IF
        END IF
        TransRecord& = TaxTrans(1).LastTrans
      WEND
      CustCnt = CustCnt + 1
    END IF
    
    ShowPctComp Cnt, NumOfTaxRecs
  NEXT
  
  GOSUB PrintAdRptEnding
  
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  
  CLOSE         'Close all open files now
  
  IF DevSpec$ = "P" THEN
    EntryPoint = 4
  ELSEIF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 1
  END IF
  
  ERASE Array, Frm, Form$, Fld, TaxCustRec
  
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  'KILL ReportFile$
  
  EXIT SUB
  
PrintAdRptHeader:
  PAGE = PAGE + 1
  PRINT #RptHandle, TAB(20); "Property Tax Ad/Collection Listing"
  PRINT #RptHandle, "Report Date: "; DATE$; TAB(65); "Page #"; PAGE
  PRINT #RptHandle, "Name"; TAB(34); "MAP-BLOCK-LOT  Prop Desc1"; TAB(68); "Tax Balance"
  PRINT #RptHandle, Dash80$
  LineCnt = 3
RETURN
  
  
PrintAdRptEnding:
  PRINT #RptHandle, Dash80$
  PRINT #RptHandle, "Total Ad/Collection Amount Charged: "; USING "$$######,#.##"; TotalCharged#
  
  PRINT #RptHandle,
  PRINT #RptHandle, FF$
  RETURN
  
GetAdNameIndex:
  REDIM Array(1 TO NumOfTaxRecs) AS Struct
  FOR Cnt = 1 TO NumOfTaxRecs
    GET TaxFile, Cnt, TaxCustRec(1)
    Array(Cnt).who = UCASE$(TaxCustRec(1).SName) + " "
    Array(Cnt).RecNum = Cnt
  NEXT
  
  'Sort Them Here
  SortT Array(1), NumOfTaxRecs, 0, LEN(Array(1)), 0, 14
  RETURN
  
END SUB

