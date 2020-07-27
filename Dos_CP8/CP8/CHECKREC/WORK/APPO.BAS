DECLARE SUB POControl ()
DEFINT A-Z

'$INCLUDE: 'DefCnf.bi'                    'monitor and color info
'$INCLUDE: 'FieldInf.bi'                  'field information
'$INCLUDE: 'FormEdit.bi'                  'form editing information
'$INCLUDE: 'pageinfo.BI'                  'Form Page info
'$INCLUDE: 'QSCR.BI'                      'QuickScreen Routines
'$INCLUDE: 'GLAUX.BI'                     'Common GL Routines
'$INCLUDE: 'GL.BI'                        'GL File Types
'$INCLUDE: 'SCRAUX.bi'
'$INCLUDE: 'ap.bi'
'$INCLUDE: 'Poform.bi'

DECLARE SUB VoidPOTrans (VoidTransRecNum&)
DECLARE SUB CancelPO ()
DECLARE SUB OpenPOTransFile (TransFileNum%, NumTrans&)
DECLARE SUB Post2PO (FileName$, BadTrans%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION EDCheck4Vendor% (FORM$(), Fld() AS ANY, Terms$)
DECLARE FUNCTION EDCheck4Acct% (FORM$(), AcctFldNo%)
DECLARE FUNCTION FindVendorRec% (VendorCode$)
DECLARE SUB PostPOTrans ()
DECLARE SUB OpenAPLedgerFile (APLedgerFile%, NumTrans%, RecLen%)
DECLARE SUB OpenAPDistFile (APDistFile%, NumDistRecs&, RecLen%)
DECLARE SUB OpenAPVendorFile (APVendorFile%, NumVRecs%, VendorRecLen%)
DECLARE SUB OpenVendorFile (VendorFile%, NumVRecs%)
DECLARE SUB OpenVendorIdx (VendorIdxFile%, NumActiveVendors%)
DECLARE FUNCTION ChkForVendor% (ChkVendor$)
DECLARE SUB VertMenuT (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION FUsing$ (Number$, Mask$)
DECLARE SUB HideCursor ()
DECLARE FUNCTION Monitor% ()
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Peek1% (Segment%, Address%)
DECLARE SUB QPrint (x$, Colr%, page%)
DECLARE SUB QPrintRC (t$, r%, c%, Clr%)
DECLARE FUNCTION QPTrim$ (x$)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE FUNCTION WaitKey% ()
DECLARE SUB KillFile (FileName$)
DECLARE SUB EdEntry ()
DECLARE SUB OpenPOEditFile (POEditFile%, NumEdTrans%)
DECLARE SUB NewCJEntry ()
DECLARE FUNCTION GetNextRec% (FileNum%, NumRecs%, CurrRec%, NextRec%)
DECLARE SUB PrnEditList ()
DECLARE SUB DisplayTotals ()
DECLARE SUB ValidateForm (FORM$(), AcctFlds() AS ANY, BadFld%)
DECLARE FUNCTION ListRecs% (APType)
DECLARE SUB ResetForm (FORM$(), NumFlds%)
DECLARE SUB ShowEdMode (EDMode%)
DECLARE SUB Rec2Form (FORM$(), RecordNumber%, Fld() AS ANY, APType)
DECLARE SUB SaveTrans (FORM$(), RecordNumber%, EDMode%, APType)
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB GetFundList (FundList$(), NumFunds%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB POMenu ()
DECLARE SUB PrintForms ()
DECLARE SUB ClrBG ()
DECLARE SUB MnuTitle (Title$, x%, Y%, Wdth%, TitColor%)

CONST False = 0, True = NOT False
CONST VNameFld = 5
CONST VRecFld = 6

DIM SHARED TotLen, FundLen, AcctLen, DetLen
DIM SHARED CashAcct$, APAcct$, EncAcct$, User$ ', FBAcct$
DIM SHARED AcctIdx AS GLAcctIndexType
DIM SHARED Acct AS GLAcctRecType
DIM SHARED POEdit AS POFORMRecType2
DIM SHARED POTrans AS GLTransRecType

'$INCLUDE: 'init.bi'                      'Get GL system info

TYPE FLen
	V AS STRING * 42
END TYPE

POMenu
ClrBG
RUN "APMENU"

SUB CancelPO

	REDIM APLedgerRec(1) AS APLedger81RecType
	LdRecLen = LEN(APLedgerRec(1))
	DIM VendorIdx AS VendorIdxRecType
	DIM VENDOR AS VendorRecType

	FormName$ = "VOIDPO"
	LibName$ = "AP"
	REDIM Frm(1) AS FormInfo
	NumFlds = LibNumberOfFields(LibName$, FormName$)
	REDIM FORM$(NumFlds, 2)             'DIM the form data array
	REDIM Fld(NumFlds) AS FieldInfo     'DIM the field information array
	StartEl = 0                         'Load first form at array start
	LibGetFldDef LibName$, FormName$, StartEl, Fld(), FORM$(), ErrCode
	Frm(1).FldNo = 1                    'Start editing on field #1
	Frm(1).InsStat = False              'Set the insert state (-1 = Insert on)
	Frm(1).StartEl = 0                  'Set form starting element to 0 and
	Action = 1
	FirstTime = True

	LibFile2Scrn LibName$, FormName$, MonoCode, -1, ErrCode
	DO
		EditForm FORM$(), Fld(), Frm(1), Cnf, Action
		IF Frm(1).FldNo > 1 AND Frm(1).PrevFld = 1 THEN   'User is moving around the form
			ChkVendor$ = QPTrim$(FORM$(1, 0))
			GOSUB Check4Vendor
			IF NOT Abort THEN
				GOSUB FindPO
			ELSE
				Abort = False
			END IF
		END IF

		IF Frm(1).KeyCode = EscKey THEN
			ExitFlag = True
		ELSEIF Frm(1).KeyCode = F10Key THEN
			OkFlag = True
		END IF

	LOOP UNTIL ExitFlag OR OkFlag

	IF ExitFlag THEN EXIT SUB
	
	IF OkFlag THEN
		Temp$ = InputBox$(LibName$, "OK2VOID", Cancel%)
	END IF

	IF Cancel THEN
		EXIT SUB
	ELSE
		'VoidChkNum& = VAL(Form$(2, 0)) 'QPValL(Form$(2, 0))
		VoidPOTrans VoidTransRecNum&
	END IF

	EXIT SUB


FindPO:
	REDIM ScrBuf(2000)
	MScrnSave 1, 1, 25, 80, ScrBuf(0)             'Save the existing screen
	POCnt = 0
	List$ = SPACE$(45)

	OpenVendorFile VendorFile, NumVRecs
	OpenAPLedgerFile APLedgerFile, NumTrans, LdRecLen

	GET VendorFile, VoidVendorRecNum&, VENDOR
	NextTrans& = VENDOR.FrstTran
	DO UNTIL NextTrans& = 0
		GET APLedgerFile, NextTrans&, APLedgerRec(1)
		IF APLedgerRec(1).TrCode = 4 THEN
			 POCnt = POCnt + 1
		END IF
		NextTrans& = APLedgerRec(1).NextTrans

	LOOP

	IF POCnt = 0 THEN
		Ok = MsgBox(LibName$, "NOCHECKS")
		Frm(1).FldNo = 1
		Action = 1
		GOTO VoidPOExit
	END IF

	REDIM POList(1 TO POCnt) AS FLen       'Dim array to hold trans list

	MatchCnt = 0
	NextTrans& = VENDOR.FrstTran

	DO UNTIL NextTrans& = 0
		GET APLedgerFile, NextTrans&, APLedgerRec(1)
			IF APLedgerRec(1).TrCode = 4 THEN
				MatchCnt = MatchCnt + 1
				LSET POList(MatchCnt).V = List$
				POList(MatchCnt).V = QPTrim$(APLedgerRec(1).PONUM)
				MID$(POList(MatchCnt).V, 11) = Num2Date(APLedgerRec(1).TrDate)
				MID$(POList(MatchCnt).V, 22) = FUsing(STR$(APLedgerRec(1).Amt), "$$########,.##")
				MID$(POList(MatchCnt).V, 39) = MKL$(NextTrans&)
			END IF
			NextTrans& = APLedgerRec(1).NextTrans
		LOOP

	CLOSE

	Row = 12
	Col = 15
	LOCATE Row, Col
	Action = 0
	MaxWidth = 37
	BoxBot = 18
	ListDescTxt$ = "P O Num       Date          Amount"
	ListDesc$ = SPACE$(MaxWidth + 4)
	MID$(ListDesc$, 3) = ListDescTxt$
	QPrintRC ListDesc$, Row - 1, Col, 112       'Cnf.ActivCh
	Choice = 1
	DO
		'call vertmenu to display matching recs to choose from
		VertMenuT POList(), Choice, MaxWidth, BoxBot, Ky$, Action, Cnf
		IF Ky$ = CHR$(27) THEN
			Abort = True
			Frm(1).FldNo = 1
			Action = 1
			GOTO VoidPOExit
		END IF
		IF Choice > 0 THEN
			LSET FORM$(2, 0) = QPTrim$(LEFT$(POList(Choice).V, 10))
			SaveField 2, FORM$(), Fld(), BadFld
			VoidTransRecNum& = CVL(RIGHT$(POList(Choice).V, 4))
			Frm(1).FldNo = 3
			EXIT DO
		END IF
	LOOP

VoidPOExit:
	MScrnRest 1, 1, 25, 80, ScrBuf(0)           'Restore Screen

RETURN


Check4Vendor:
	MatchingVendors = 0
	List$ = SPACE$(65)

	VLookup$ = QPTrim$(UCASE$(ChkVendor$))
	VLookuplen = LEN(VLookup$)

	REDIM ScrBuf(2000)

	MScrnSave 1, 1, 25, 80, ScrBuf(0)             'Save the existing screen

	OpenVendorIdx VendorIdxFile, NumActiveVendors

	IF NumActiveVendors = 0 THEN
		CLOSE
		ListVendors = 0
		EXIT SUB
	END IF

	OpenVendorFile VendorFile, NumVRecs

	REDIM VendorList(1 TO NumVRecs) AS FLen       'Dim array to hold trans list

	FOR Cnt = 1 TO NumActiveVendors               'Read the file
		GET VendorIdxFile, Cnt, VendorIdx
		IF LEFT$(VendorIdx.VendorCode, VLookuplen) = VLookup$ THEN
			GET VendorFile, VendorIdx.RecNum, VENDOR
			IF VENDOR.DELFLAG = 0 THEN
				MatchingVendors = MatchingVendors + 1
				LSET VendorList(MatchingVendors).V = List$
				VendorList(MatchingVendors).V = VendorIdx.VendorCode
				MID$(VendorList(MatchingVendors).V, 14) = VENDOR.VNAME
				MID$(VendorList(MatchingVendors).V, 41) = MKI$(VendorIdx.RecNum)
			END IF
		END IF
	NEXT

	CLOSE

	IF MatchingVendors = 0 THEN
		Ok = MsgBox(LibName$, "NOVENDOR")
		VendPickList = 0
		Frm(1).FldNo = 1
		Action = 1
		GOTO VoidChkReturn
	END IF

	IF MatchingVendors = 1 THEN
		LSET FORM$(Frm(1).PrevFld, 0) = VendorList(1).V
		SaveField Frm(1).PrevFld, FORM$(), Fld(), BadFld
		Action = 1
		VoidVendorRecNum& = CVI(RIGHT$(VendorList(1).V, 2))
		GOTO VoidChkReturn
	END IF

	IF MatchingVendors > 1 THEN
		REDIM PRESERVE VendorList(1 TO MatchingVendors) AS FLen
		Row = 5: Col = 15
		LOCATE Row, Col
		Action = 0
		MaxWidth = 39
		BoxBot = 18
		ListDescTxt$ = "VendorCode  Name"
		ListDesc$ = SPACE$(MaxWidth + 4)
		MID$(ListDesc$, 3) = ListDescTxt$
		QPrintRC ListDesc$, Row - 1, Col, 112       'Cnf.ActivCh
		Choice = 1
		DO
			'call vertmenu to display matching recs to choose from
			VertMenuT VendorList(), Choice, MaxWidth, BoxBot, Ky$, Action, Cnf
			IF Ky$ = CHR$(27) THEN
				CLOSE
				Abort = True
				Frm(1).FldNo = 1
				Action = 1
				GOTO VoidChkReturn
			END IF
			IF Choice > 0 THEN
				VoidVendorRecNum& = CVI(RIGHT$(VendorList(Choice).V, 2))
				EXIT DO
			END IF
		LOOP
		'  MScrnRest 1, 1, 25, 80, ScrBuf(0)           'Restore Screen
	ELSE
	END IF

	IF Choice > 0 THEN
		LSET FORM$(Frm(1).PrevFld, 0) = VendorList(Choice).V
		SaveField Frm(1).PrevFld, FORM$(), Fld(), BadFld
		Choice = 0
	END IF

VoidChkReturn:
	MScrnRest 1, 1, 25, 80, ScrBuf(0)             'Restore Screen
	QPrintRC TName$, 3, 27, -1
	ERASE VendorList

	RETURN



END SUB

FUNCTION EDCheck4Acct% (FORM$(), AcctFld%)
 
	MatchAccts = 0
	List$ = SPACE$(65)

	Lookup$ = QPTrim$(UCASE$(FORM$(AcctFld, 0)))
	Lookuplen = LEN(Lookup$)

	REDIM ScrBuf(2000)
	MScrnSave 1, 1, 25, 80, ScrBuf(0)             'Save the existing screen
	
	OpenAcctIdx AcctIdxFileNum, NumIdxRecs
	CLOSE AcctIdxFileNum
	REDIM TempAcctIdx(1 TO NumIdxRecs) AS GLAcctIndexType

	ElSize = LEN(TempAcctIdx(1))
	FGetAH "GLACCT.IDX", TempAcctIdx(1), ElSize, NumIdxRecs

	OpenAcctFile AcctFileNum, NumAccts

	REDIM AcctList(1 TO NumIdxRecs) AS FLen       'Dim array to hold trans list
	FOR Cnt = 1 TO NumIdxRecs               'Read the file
		IF LEFT$(TempAcctIdx(Cnt).AcctNum, Lookuplen) = Lookup$ THEN
			GET AcctFileNum, TempAcctIdx(Cnt).RecNum, Acct
			IF Acct.Typ = "E" THEN
				MatchAccts = MatchAccts + 1
				LSET AcctList(MatchAccts).V = List$
				AcctList(MatchAccts).V = TempAcctIdx(Cnt).AcctNum
				MID$(AcctList(MatchAccts).V, TotLen + 4) = Acct.Title
				MID$(AcctList(MatchAccts).V, 41) = MKI$(TempAcctIdx(Cnt).RecNum)
			END IF
		END IF
	NEXT

	CLOSE AcctFileNum

	SELECT CASE MatchAccts
		CASE 0
			PLAY "L32O3EC"            'Acct NOT found
			Ok = MsgBox("GL.QSL", "NOACCT")
			EDCheck4Acct = False
		CASE 1
			AcctRecNum = CVI(RIGHT$(AcctList(1).V, 2))
			GOSUB GetAcctInfo
			EDCheck4Acct = True
		CASE ELSE
			Row = 5
			Col = 15
			LOCATE Row, Col
			Action = 0
			MaxLen = 39
			BoxBot = 18
			ListDescTxt$ = "Account No.   Description"
			ListDesc$ = SPACE$(MaxLen + 4)
			MID$(ListDesc$, 3) = ListDescTxt$
			QPrintRC ListDesc$, Row - 1, Col, 112       'Cnf.ActivCh
			DO         'call vertmenu to display matching recs to choose from
				VertMenuT AcctList(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
				IF Ky$ = CHR$(27) THEN
					EDCheck4Acct = False
					EXIT DO
				END IF
				IF Choice > 0 THEN
					AcctRecNum = CVI(RIGHT$(AcctList(Choice).V, 2))
					GOSUB GetAcctInfo
					EDCheck4Acct = True
					EXIT DO
				END IF
			LOOP
		END SELECT

	MScrnRest 1, 1, 25, 80, ScrBuf(0)             'Restore Screen

	ERASE AcctList, TempAcctIdx, ScrBuf


EXIT FUNCTION

GetAcctInfo:
	OpenAcctFile AcctFileNum, NumAccts
	GET AcctFileNum, AcctRecNum, Acct
	CLOSE AcctFileNum
	LSET FORM$(AcctFld, 0) = Acct.Num
	'LSET Form$(AcctFld + 1, 0) = Acct.Title
	LSET FORM$(AcctFld + 1, 0) = MKI$(AcctRecNum)
RETURN



END FUNCTION

FUNCTION EDCheck4Vendor% (FORM$(), Fld() AS FieldInfo, Terms$)
	 
	SHARED VendorIdx AS VendorIdxRecType
	SHARED VENDOR AS VendorRecType

	MatchingVendors = 0
	List$ = SPACE$(65)

	VLookup$ = QPTrim$(UCASE$(FORM$(VNameFld, 0)))
	VLookuplen = LEN(VLookup$)

	REDIM ScrBuf(2000)

	MScrnSave 1, 1, 25, 80, ScrBuf(0)             'Save the existing screen

	OpenVendorIdx VendorIdxFile, NumActiveVendors

	IF NumActiveVendors = 0 THEN
		CLOSE
		ListVendors = 0
		EXIT FUNCTION
	END IF

	OpenVendorFile VendorFile, NumVRecs

	REDIM VendorList(1 TO NumActiveVendors) AS FLen       'Dim array to hold trans list
	
	FOR Cnt = 1 TO NumActiveVendors               'Read the file
		GET VendorIdxFile, Cnt, VendorIdx
		IF LEFT$(VendorIdx.VendorCode, VLookuplen) = VLookup$ THEN
			GET VendorFile, VendorIdx.RecNum, VENDOR
			IF VENDOR.DELFLAG = 0 THEN
				MatchingVendors = MatchingVendors + 1
				LSET VendorList(MatchingVendors).V = List$
				VendorList(MatchingVendors).V = VendorIdx.VendorCode
				MID$(VendorList(MatchingVendors).V, 13) = VENDOR.VNAME
				MID$(VendorList(MatchingVendors).V, 41) = MKI$(VendorIdx.RecNum)
			END IF
		END IF
	NEXT

	CLOSE

	SELECT CASE MatchingVendors
	CASE 0
		Ok = MsgBox("AP", "NOVENDOR")
		VendorPickList = 0
		EDCheck4Vendor = False
	CASE 1
		EDCheck4Vendor = True
		LSET FORM$(VNameFld, 0) = VendorList(1).V
		'LSET Form$(6, 0) = STR$(CVI(MID$(VendorList(1).V, 13, 28)))
		VRecNum = CVI(MID$(VendorList(1).V, 41, 2))
		GOSUB GetVendorAddress
	CASE ELSE
		REDIM PRESERVE VendorList(1 TO MatchingVendors) AS FLen
		Row = 5: Col = 15
		LOCATE Row, Col
		Action = 0
		MaxLen = 39
		BoxBot = 18
		ListDescTxt$ = "VendorCode  Name"
		ListDesc$ = SPACE$(MaxLen + 4)
		MID$(ListDesc$, 3) = ListDescTxt$
		QPrintRC ListDesc$, Row - 1, Col, 112       'Cnf.ActivCh
		DO         'call vertmenu to display matching recs to choose from
			VertMenuT VendorList(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
			IF Ky$ = CHR$(27) THEN
				CLOSE
				EDCheck4Vendor = False
				EXIT DO
			END IF
			IF Choice > 0 THEN
				LSET FORM$(VNameFld, 0) = VendorList(Choice).V
				VRecNum = CVI(MID$(VendorList(Choice).V, 41, 2))
				GOSUB GetVendorAddress
				EDCheck4Vendor = True
				'  SaveField 2, Form$(), Fld(), BadFld
				EXIT DO
			END IF
		LOOP

	END SELECT

	MScrnRest 1, 1, 25, 80, ScrBuf(0)             'Restore Screen
	QPrintRC TName$, 3, 27, -1
	ERASE VendorList, ScrBuf

EXIT FUNCTION

GetVendorAddress:
	OpenVendorFile VendorFile, NumVRecs
	GET VendorFile, VRecNum, VENDOR

	FORM$(VRecFld, 0) = STR$(VRecNum)
	SaveField VRecFld, FORM$(), Fld(), BadFld

	FORM$(VRecFld + 1, 0) = VENDOR.VNAME
	SaveField VRecFld + 1, FORM$(), Fld(), BadFld

	FORM$(VRecFld + 2, 0) = VENDOR.Addr1
	SaveField VRecFld + 2, FORM$(), Fld(), BadFld

	FORM$(VRecFld + 3, 0) = VENDOR.Addr2
	SaveField VRecFld + 3, FORM$(), Fld(), BadFld

	FORM$(VRecFld + 4, 0) = VENDOR.City
	SaveField VRecFld + 4, FORM$(), Fld(), BadFld

	FORM$(VRecFld + 5, 0) = VENDOR.State + " " + LEFT$(VENDOR.Zip, 5) + "-" + RIGHT$(VENDOR.Zip, 4)
	SaveField VRecFld + 5, FORM$(), Fld(), BadFld

	CLOSE
	RETURN

END FUNCTION

SUB EdEntry

	SHARED VendorIdx AS VendorIdxRecType
	SHARED VENDOR AS VendorRecType

mainpobody:


	  REDIM POCont(1) AS POControlRecType
	
	LibName$ = "PO"
	FormName$ = "POFORM"
	NumOfScrns = 3
	Cnt = 0
	Saved$ = "N"

	REDIM FormName$(1 TO NumOfScrns)
	FOR Scr = 1 TO NumOfScrns
		FormName$(Scr) = FormName$ + LTRIM$(STR$(Scr))
	NEXT
	 
	'-- Get total number of fields
	NumFlds = -1
	FOR Scr = 1 TO NumOfScrns
		NumFlds = NumFlds + LibNumberOfFields(LibName$, FormName$(Scr)) + 1
	NEXT
	
	REDIM FORM$(NumFlds, 2)
	REDIM Fld(NumFlds) AS FieldInfo
	REDIM PageInfo(1 TO NumOfScrns) AS PageInfoType
	REDIM Frm(1)    AS FormInfo

	'--for each page, get first and last fields
	StartEl = 0   'Load first form at array start
	FOR Scr = 1 TO NumOfScrns     'Get field defs
		LibGetFldDef LibName$, FormName$(Scr), StartEl, Fld(), FORM$(), ErrCode
		PageInfo(Scr).PageNo = Scr
		PageInfo(Scr).FirstFld = StartEl + 1
		StartEl = StartEl + Fld(StartEl).Fields + 1 'Inc StartEl to next
		PageInfo(Scr).LastFld = StartEl - 1
	NEXT
	
	'--Clear all fields and keep a list of the acct number fields
	REDIM AcctFlds(1 TO 12)
	FOR F = 1 TO NumFlds
		LSET FORM$(F, 0) = ""
		IF QPTrim$(Fld(F).ScratchS) = "A" THEN
			Cnt = Cnt + 1
			AcctFlds(Cnt) = F
		END IF
	NEXT
	'Frm(1).PageDisable = True
	
	'--get first active record number and set Editing Mode
	OpenPOEditFile POEditFile, NumEdTrans
	CurrRec = 0: NextRec = 1
	RecNum = GetNextRec(POEditFile, NumEdTrans, CurrRec, NextRec)
	CLOSE POEditFile
	IF RecNum = 0 THEN
		EDMode = 1
	ELSE
		EDMode = 0
      END IF
	
	Scr = 1
	LibFile2Scrn LibName$, FormName$(Scr), MonoCode, -1, ErrCode
	IF EDMode = 0 THEN
		Rec2Form FORM$(), RecNum, Fld(), APType
			       
	END IF
	ShowEdMode EDMode
		IF EDMode = 1 THEN
		POFILE = FREEFILE
		OPEN "APPOCRL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #POFILE LEN = LEN(POCont(1))
		IF LOF(POFILE) > 0 THEN
			GET POFILE, 1, POCont(1)
			FORM$(2, 0) = POCont(1).PONUMBER
			FORM$(12, 0) = POCont(1).Shipto1
			FORM$(13, 0) = POCont(1).Shipto2
			FORM$(14, 0) = POCont(1).Shipto3
			FORM$(15, 0) = POCont(1).Shipto4
			FORM$(16, 0) = POCont(1).Shipto5
			FORM$(17, 0) = POCont(1).FOB
			FORM$(18, 0) = POCont(1).Shipvia
			FORM$(19, 0) = POCont(1).Terms
			FORM$(21, 0) = POCont(1).Addinst1
			FORM$(22, 0) = POCont(1).Addinst2
			FORM$(23, 0) = POCont(1).Addinst3
			CLOSE POFILE
		END IF
		END IF

	Action = 1
	
	CommaFormat$ = "#########,.##"
	GOSUB ShowDistTot

	DO
		IF FirstTime THEN
			FirstTime = 0
			Frm(1).FldNo = 1
			Scr = 1
			Action = 1
		END IF

		EditForm FORM$(), Fld(), Frm(1), Cnf, Action
		IF Scr = 1 THEN PONUMBER$ = FORM$(2, 0)

		'Alt+F3 code was here
		
		'f2 key was here
		'SELECT CASE Frm(1).KeyCode
		'END SELECT
		
		'--User is moving around the form
		IF Frm(1).PrevFld <> Frm(1).FldNo THEN

			'--Check for vendor
			IF Frm(1).PrevFld = VNameFld AND Frm(1).FldNo <> VNameFld THEN
				VendorOK = EDCheck4Vendor(FORM$(), Fld(), Terms$)
				IF VendorOK THEN
					'Frm(1).FldNo = 11
					Action = 1
				ELSE
					Frm(1).FldNo = VNameFld
					Action = 2
				END IF
			END IF
			
			'--strip out the dashes if we are in an account field
			IF QPTrim$(Fld(Frm(1).FldNo).ScratchS) = "A" THEN
				LSET FORM$(Frm(1).FldNo, 0) = StripAcct$(FORM$(Frm(1).FldNo, 0))
				Action = 1
			END IF
			
			'--check for valid acct
			IF LEN(QPTrim$(FORM$(Frm(1).PrevFld, 0))) THEN
				IF QPTrim$(Fld(Frm(1).PrevFld).ScratchS) = "A" THEN
					FAcct$ = FmtAcct$(QPTrim$(FORM$(Frm(1).PrevFld, 0)), FundLen%, AcctLen%, DetLen%)
					LSET FORM$(Frm(1).PrevFld, 0) = FAcct$
					GOSUB CheckForValidAcctNum
				END IF
			END IF

			IF QPTrim$(Fld(Frm(1).PrevFld).ScratchS) = "S" THEN
				GOSUB ShowDistTot
			END IF

		END IF      'End of test for PrevFld <> current FldNo
		
		IF SaveFlag THEN
		 GOSUB SaveRec
		 IF Saved$ = "Y" THEN GOTO mainpobody
		END IF
		
		'--Process key strokes
		SELECT CASE Frm(1).KeyCode
			CASE F2Key  'Change Editing Mode
				GOSUB CheckFormEdit
				IF CanChange THEN
					GOSUB InitForm
					SELECT CASE EDMode
						CASE 0                  'we're in Edit Mode
							EDMode = 1            'switch to add mode
							ShowEdMode EDMode
							ResetForm FORM$(), NumFlds
							Frm(1).FldNo = 1
							POFILE = FREEFILE
							OPEN "APPOCRL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #POFILE LEN = LEN(POCont(1))
							IF LOF(POFILE) > 0 THEN
								GET POFILE, 1, POCont(1)
								FORM$(2, 0) = POCont(1).PONUMBER
								FORM$(12, 0) = POCont(1).Shipto1
								FORM$(13, 0) = POCont(1).Shipto2
								FORM$(14, 0) = POCont(1).Shipto3
								FORM$(15, 0) = POCont(1).Shipto4
								FORM$(16, 0) = POCont(1).Shipto5
								FORM$(17, 0) = POCont(1).FOB
								FORM$(18, 0) = POCont(1).Shipvia
								FORM$(19, 0) = POCont(1).Terms
								FORM$(21, 0) = POCont(1).Addinst1
								FORM$(22, 0) = POCont(1).Addinst2
								FORM$(23, 0) = POCont(1).Addinst3
							END IF
			CLOSE POFILE
		




							Action = 1
							GOSUB ShowDistTot
						CASE 1  'currently adding records
							'--check to see if there are any active transactions
							OpenPOEditFile POEditFile, NumEdTrans
							CurrRec = 0: NextRec = 1
							RecNum = GetNextRec(POEditFile, NumEdTrans, CurrRec, NextRec)
							CLOSE POEditFile
							IF RecNum = 0 THEN    'No active records can't edit
								EDMode = 1
							ELSE
								EDMode = 0          'switch to edit mode
								ShowEdMode EDMode
								RecNum = ListRecs(APType)
								IF RecNum > 0 THEN
									Rec2Form FORM$(), RecNum, Fld(), APType
									Frm(1).FldNo = 1
									Action = 1
								ELSE
									CLOSE
									EXIT SUB
								END IF
							END IF
							GOSUB ShowDistTot
						CASE ELSE
					END SELECT              'Edmode
				END IF

			CASE F3Key
				IF EMode = 0 THEN
					Ok = MsgBox%("GL.QSL", "OK2DELTR")
					IF Ok = 2 THEN          'Pressed yes button #2
						FORM$(1, 0) = "-1"
						SaveField 1, FORM$(), Fld(), BadFld
						SaveTrans FORM$(), RecNum, EDMode, APType
						ResetForm FORM$(), NumFlds
						RecNum = ListRecs(APType)
						IF RecNum > 0 THEN
							Rec2Form FORM$(), RecNum, Fld(), APType
						ELSE
							EXIT SUB
						END IF
					ELSE
						Action = 2
					END IF
					Frm(1).FldNo = 1
					Action = 1
				END IF

			CASE F5KEY
				GOSUB CheckFormEdit
				IF CanChange THEN
					GOSUB InitForm
					IF EDMode = 0 THEN
						RecNum = ListRecs(APType)
						IF RecNum > 0 THEN
							Rec2Form FORM$(), RecNum, Fld(), APType
							Frm(1).FldNo = 1
							Action = 1
						END IF
					END IF
				END IF

			CASE F10Key
				'--Make sure acct fields are formatted correctly if user
				'--Preses F10 with cursor in an account field
				IF QPTrim$(Fld(Frm(1).FldNo).ScratchS) = "A" THEN
					IF LEN(QPTrim$(FORM$(Frm(1).FldNo, 0))) > 0 THEN
						FAcct$ = FmtAcct$(RTRIM$(FORM$(Frm(1).FldNo, 0)), FundLen%, AcctLen%, DetLen%)
						LSET FORM$(Frm(1).FldNo, 0) = FAcct$
						Frm(1).FldNo = 1
						Action = 1
					END IF
				END IF
				'--send user back thru loop to verify acct before saving
				Action = 1
				SaveFlag = True
			
			CASE CtrlL  'Ctrl+LeftArrow               'Get the next record
				GOSUB CheckFormEdit
				IF CanChange THEN
					IF RecNum > 0 THEN
						Frm(1).FldNo = 1
						OpenPOEditFile POEditFile, NumEdTrans
						CurrRec = RecNum: NextRec = -1
						RecNum = GetNextRec(POEditFile, NumEdTrans, CurrRec, NextRec)
						CLOSE POEditFile
						GOSUB InitForm
						Rec2Form FORM$(), RecNum, Fld(), APType
						GOSUB ShowDistTot
						ShowEdMode EDMode
						'Frm(1).PrevFld = 1
						Scr = 1
						Action = 1
					END IF
				END IF

			CASE CtrlR  'Ctrl+RightArrow             'Get the previous record
				GOSUB CheckFormEdit
				IF CanChange THEN
					IF RecNum > 0 THEN
						Frm(1).FldNo = 1
						OpenPOEditFile POEditFile, NumEdTrans
						CurrRec = RecNum: NextRec = 1
						RecNum = GetNextRec(POEditFile, NumEdTrans, CurrRec, NextRec)
						CLOSE POEditFile
						GOSUB InitForm
						Rec2Form FORM$(), RecNum, Fld(), APType
						GOSUB ShowDistTot
						ShowEdMode EDMode
						' Frm(1).PrevFld = 1
						Action = 1
						Scr = 1
					END IF
				END IF
		END SELECT
		
		'--Process screen page changes
		IF Frm(1).FldNo > PageInfo(Scr).LastFld OR Frm(1).FldNo < PageInfo(Scr).FirstFld THEN
			SELECT CASE Frm(1).FldNo
				CASE PageInfo(1).FirstFld TO PageInfo(1).LastFld
					Scr = 1
				CASE PageInfo(2).FirstFld TO PageInfo(2).LastFld
					Scr = 2
				CASE PageInfo(3).FirstFld TO PageInfo(3).LastFld
					Scr = 3
			END SELECT
			LibFile2Scrn LibName$, FormName$(Scr), MonoCode, -2, ErrCode
			GOSUB ShowDistTot
			ShowEdMode EDMode
			Action = 1
		END IF
		
		'--Buttons are not attached to the form, the following processes the mouse
		IF Frm(1).Presses THEN
			SELECT CASE Frm(1).MRow
				CASE 21
					SELECT CASE Frm(1).MCol
					CASE 34 TO 44
						PressButton F10Key, 21, 34, 44
					CASE 47 TO 60
						PressButton F3Key, 21, 47, 60
					CASE 63 TO 75
						PressButton ESC, 21, 63, 75
					END SELECT
				CASE 25
					SELECT CASE Frm(1).MCol
					CASE 40 TO 50
						PressButton F2Key, 25, 40, 50
					CASE 53 TO 63
						PressButton F5KEY, 25, 53, 63
					CASE 66 TO 70
						PressButton CtrlL, 25, 66, 70
					CASE 74 TO 78
						PressButton CtrlR, 25, 74, 78
					END SELECT
			END SELECT
		END IF
		
	LOOP UNTIL Frm(1).KeyCode = ESC

EXIT SUB
	
	
CheckForValidAcctNum:
	AcctFld = Frm(1).PrevFld
	SELECT CASE EDCheck4Acct%(FORM$(), AcctFld%)
		CASE True
			SaveField AcctFld, FORM$(), Fld(), BadFld
			Action = 2
		CASE ELSE
			Frm(1).FldNo = AcctFld
			LSET FORM$(Frm(1).FldNo, 0) = ""
			Action = 2
			SaveFlag = False
	END SELECT
	RETURN
	
	
ShowDistTot:
	TotDist# = 0
	FOR Cnt = 1 TO NumFlds        'look thru acct fld list
		IF QPTrim$(Fld(Cnt).ScratchS) = "S" THEN    '= Distribution Field
			TotDist# = Round#(TotDist# + Value(FORM$(Cnt, 0), ECode))
		END IF
	NEXT
	Dist$ = FUsing(STR$(TotDist#), CommaFormat$)
	QPrintRC Dist$, 19, 48, -1
	RETURN
	
	
SaveRec:
	GOSUB ShowDistTot
	IF TotDist# <> Value(FORM$(24, 0), ECode) THEN
		Ok = MsgBox%(LibName$, "BADPO2")
		Frm(1).FldNo = 1
		GOTO JumpOut
	END IF

	BadFld = 0
	ValidateForm FORM$(), AcctFlds(), BadFld    'Make sure we have req fields
	IF BadFld = 0 THEN
		GOSUB ShowDistTot
		FORM$(1, 0) = "0"                         'Set deleted flag to false
		SaveField 1, FORM$(), Fld(), BadFld       'Write the flag to the form$
		SaveTrans FORM$(), RecNum, EDMode, APType 'Save the transaction
		Ok = MsgBox%("GL.QSL", "SAVED")           'tell user we saved it
		IF EDMode = 1 THEN                        'if in Add Mode, clear
			ResetForm FORM$(), NumFlds              ' the form for the next entry
			POFILE = FREEFILE
			OPEN "APPOCRL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #POFILE LEN = LEN(POCont(1))
			IF LOF(POFILE) > 0 THEN
			 GET POFILE, 1, POCont(1)
			 KK = INSTR(PONUMBER$, "-")
			 IF KK > 0 THEN
			  OLDSTUFF$ = LEFT$(PONUMBER$, KK)
			  PONUMBER$ = MID$(PONUMBER$, KK + 1, LEN(PONUMBER$) - KK + 1)
			 END IF
			 PONUMBER = VAL(PONUMBER$)
			 PONUMBER = PONUMBER + 1
			 PONUMBER$ = LTRIM$(STR$(PONUMBER))
			 IF KK > 0 THEN
			 PONUMBER$ = OLDSTUFF$ + PONUMBER$
			 END IF
			 POCont(1).PONUMBER = PONUMBER$
			 PUT POFILE, 1, POCont(1)
			END IF
			CLOSE POFILE
		

		END IF
		Scr = 1
		Frm(1).FldNo = 1
		Saved$ = "Y"
	ELSE
		Ok = MsgBox%(LibName$, "BADPO1")
		BEEP
		Frm(1).FldNo = BadFld
	END IF
JumpOut:
	Action = 1
	FirstTime = -1
	SaveFlag = False
	RETURN
	
	
CheckFormEdit:
	IF Frm(1).Edited THEN
		BEEP
		CanChange = False
		Ok = MsgBox("CMNDLG", "CHANGREC")           'returns 2 for Abandon
		IF Ok = 2 THEN CanChange = True
		Action = 2
	ELSE
		CanChange = True
	END IF
	RETURN

	
InitForm:
	REDIM Frm(1)    AS FormInfo
	'Frm(1).PageDisable = True
	Action = 1
	RETURN
	
	
END SUB

FUNCTION FindVendorRec (VendorCode$)
	
	SHARED VendorIdx AS VendorIdxRecType
	
	OpenVendorIdx VendorIdxFile, NumActiveVendors
	
	IF NumActiveVendors = 0 THEN
		FindVendorRec = 0
		CLOSE VendorIdxFile
		EXIT FUNCTION
	END IF
	
	Match = False
	FirstRec = 1
	LastRec = NumActiveVendors
	Lookfor$ = QPTrim$(UCASE$(VendorCode$))
	
	DO UNTIL LastRec < FirstRec
		
		MiddleRec = (LastRec + FirstRec) \ 2
		
		GET VendorIdxFile, MiddleRec, VendorIdx
		
		IF Lookfor$ = QPTrim$(VendorIdx.VendorCode) THEN
			Match = True
			EXIT DO
		ELSEIF Lookfor$ < VendorIdx.VendorCode THEN
			LastRec = MiddleRec - 1
		ELSE
			FirstRec = MiddleRec + 1
		END IF
		
	LOOP
	
	IF Match THEN
		FindVendorRec = VendorIdx.RecNum
	ELSE
		FindVendorRec = 0
	END IF
	
	CLOSE VendorIdxFile
	
END FUNCTION

'***************************************************************************
'Finds the next undeleted record.
'Call with NextRec value of -1 for previous record, +1 for the next record.
'If a record is not found, the function returns the value of CurrRec.
'File should be opened before function is called
'***************************************************************************
FUNCTION GetNextRec (FileNum, NumRecs, CurrRec, NextRec)
	
	Found = 0
	Rec = CurrRec
	
	DO
		
		Rec = Rec + NextRec         'Set file pointer to next record
		
		IF Rec > NumRecs OR Rec <= 0 THEN           'test for beg or end of file
			Found = 0 'if no more records then get out
			EXIT DO
		END IF
		
		GET FileNum, Rec, POEdit    'Get the record
		
		IF POEdit.DELETED <> -1 THEN                'Ok if not deleted
			Found = 1
			EXIT DO   'Get out of loop when we find one
		END IF
		
	LOOP
	
	IF Found = 0 THEN
		GetNextRec = CurrRec
	ELSE
		GetNextRec = Rec
	END IF
	
END FUNCTION

FUNCTION ListRecs (APType)
	
	SHARED Cnf AS Config
	REDIM ScrBuf(2000)
	
	MScrnSave 1, 1, 25, 80, ScrBuf(0)             'Save the existing screen
	
	'QPrintRC "List ", 25, 68, -1
	
	OpenPOEditFile POEditFile, NumEdTrans
	
	REDIM TransList$(1 TO NumEdTrans)             'Dim array to hold trans list
	
	Active = 0    'Counter for Active Transactions
	
	FOR Cnt = 1 TO NumEdTrans     'Read the file
		GET POEditFile, Cnt, POEdit
		IF NOT POEdit.DELETED THEN  'if active add to list
			Active = Active + 1
			
			List$ = SPACE$(75)
			
			LSET List$ = POEdit.VNDRCODE
			MID$(List$, 12) = POEdit.PONUM
			MID$(List$, 38) = ConvDateStr1(Num2Date(POEdit.PODATE))
			MID$(List$, 53) = FUsing$(STR$(POEdit.POAMT), "#########,.##")
			MID$(List$, 73) = MKI$(Cnt)
			TransList$(Active) = List$
			
		END IF
		
	NEXT
	CLOSE
	
	IF Active = 0 THEN
		'KILL "CJED.DAT"                     '..init file to zero
		EXIT FUNCTION
	END IF
	
	REDIM PRESERVE TransList$(1 TO Active)        'Redim list to only active trans
	
	LOCATE 3, 6   'set up vertmenu
	Action = 0
	MaxLen = 65
	BoxBot = 23
	
	ListDesc$ = "   Vendor    P O Number                Date            Total Amt     "
	QPrintRC ListDesc$, 2, 6, Cnf.ActivCh
	'MPaintBox 6, 75, 7, 76, 8              'to add shadow to list
	
	DO            'call vertmenu
		VertMenu TransList$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
		IF Ky$ = CHR$(27) THEN
			CLOSE
			Abort = True
			EXIT DO
			EXIT FUNCTION
		END IF
		IF Choice > 0 THEN EXIT DO
	LOOP
	
	IF Abort THEN
		Choice = 0
	ELSE
		RecNum = CVI(MID$(TransList$(Choice), 73, 2))               'Get the record number
	END IF        'off the end of the $
	
	CLOSE POEditFile              'Clean up after ourselves
	ERASE TransList$
	MScrnRest 1, 1, 25, 80, ScrBuf(0)             'Restore Screen
	
	ListRecs = RecNum
	
	
END FUNCTION

SUB MakeTempReport

	'--Posting Report
	DIM Trans2Post AS GLTransRecType
	RecLen = LEN(Trans2Post)
	
	CJ2Post = FREEFILE
	OPEN "APINVIF.dat" FOR RANDOM AS CJ2Post LEN = RecLen
	
	NumTrans = LOF(CJ2Post) \ RecLen
	
	PrnFile = FREEFILE
	OPEN "Trans.prn" FOR OUTPUT AS #PrnFile
	
	'print field names
	ToPrint$ = SPACE$(78)
	LSET ToPrint$ = "Date"
	MID$(ToPrint$, 11) = "Acct Number"
	MID$(ToPrint$, 29) = "Description"
	MID$(ToPrint$, 46) = "Reference"
	MID$(ToPrint$, 61) = "Debit"
	MID$(ToPrint$, 74) = "Credit"
	PRINT #PrnFile, ToPrint$
	
	FOR Cnt = 1 TO NumTrans
		GET CJ2Post, Cnt, Trans2Post
		ToPrint$ = SPACE$(78)
		LSET ToPrint$ = ConvDateStr1(Num2Date$(Trans2Post.TrDate))
		MID$(ToPrint$, 11) = Trans2Post.AcctNum
		MID$(ToPrint$, 29) = LEFT$(Trans2Post.Desc, 15)
		MID$(ToPrint$, 46) = Trans2Post.Ref
		MID$(ToPrint$, 56) = FUsing(STR$(Trans2Post.DrAmt), "#######,.##")
		MID$(ToPrint$, 69) = FUsing(STR$(Trans2Post.CrAmt), "#######,.##")
		'MID$(ToPrint$, 80) = STR$(Cnt&)
		'MID$(ToPrint$, 90) = STR$(Trans2Post.NextTran)
		PRINT #PrnFile, ToPrint$
	NEXT
	
	CLOSE
	
	'Prn2Scrn "Trans.prn", 1, 25, 11, PrnFlag
	KILL "trans.prn"
	
END SUB

SUB OpenAPVendorFile (APVendorFile, NumVRecs, VendorRecLen)
	
	APVendorFile = FREEFILE
	OPEN "apvendor.dat" FOR RANDOM ACCESS READ WRITE SHARED AS APVendorFile LEN = VendorRecLen
	NumVRecs = LOF(APVendorFile) \ VendorRecLen
	
END SUB

SUB OpenPOEditFile (POEditFile, NumEdTrans)
	
	FileName$ = "APPED.DAT"
	
	EdLen = LEN(POEdit)
	POEditFile = FREEFILE
	OPEN FileName$ FOR RANDOM ACCESS READ WRITE SHARED AS POEditFile LEN = EdLen
	NumEdTrans = LOF(POEditFile) \ EdLen
	
END SUB

SUB POControl
  REDIM POCont(1) AS POControlRecType
  
  LibName$ = "PO.QSL"
  FormName$ = "POCTRL"
  REDIM FORM$(0, 0)             'Form array holds data while editing
  REDIM Fld(0) AS FieldInfo     'Field info Array
  DIM Frm AS FormInfo           'Form editing info
  LibLoadDisplayForm LibName$, FormName$, FORM$(), Fld()
  PrintTitle User$
  PrintHelp "PO Control"
  OrgVend$ = ""
  Action = 1
  GOSUB CHECKPREV

  DO
    EditForm FORM$(), Fld(), Frm, Cnf, Action

    SELECT CASE Frm.KeyCode
      CASE F10Key
	IF LEN(QPTrim$(FORM$(1, 0))) = 0 THEN
	  '--Don't allow to save with a blank vendor code
	  BEEP
	  LSET FORM$(1, 0) = ""
	  Frm.FldNo = 1
	  Action = 1
	  Done = False
	ELSE
	  GOSUB SaveControl

	  Done = True
	END IF
      CASE ESC
	Done = True
      CASE ELSE
    END SELECT

  LOOP UNTIL Done
  EXIT SUB
CHECKPREV:
  POFILE = FREEFILE
  OPEN "APPOCRL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #POFILE LEN = LEN(POCont(1))
  IF LOF(POFILE) > 0 THEN
   GET POFILE, 1, POCont(1)
    FORM$(1, 0) = POCont(1).PONUMBER
    FORM$(2, 0) = POCont(1).Shipto1
    FORM$(3, 0) = POCont(1).Shipto2
    FORM$(4, 0) = POCont(1).Shipto3
    FORM$(5, 0) = POCont(1).Shipto4
    FORM$(6, 0) = POCont(1).Shipto5
    FORM$(7, 0) = POCont(1).FOB
    FORM$(8, 0) = POCont(1).Shipvia
    FORM$(9, 0) = POCont(1).Terms
    FORM$(10, 0) = POCont(1).Addinst1
    FORM$(11, 0) = POCont(1).Addinst2
    FORM$(12, 0) = POCont(1).Addinst3
   END IF
   CLOSE POFILE
   RETURN


SaveControl:
  POFILE = FREEFILE
  OPEN "APPOCRL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #POFILE LEN = LEN(POCont(1))

    POCont(1).PONUMBER = FORM$(1, 0)
    POCont(1).Shipto1 = FORM$(2, 0)
    POCont(1).Shipto2 = FORM$(3, 0)
    POCont(1).Shipto3 = FORM$(4, 0)
    POCont(1).Shipto4 = FORM$(5, 0)
    POCont(1).Shipto5 = FORM$(6, 0)
    POCont(1).FOB = FORM$(7, 0)
    POCont(1).Shipvia = FORM$(8, 0)
    POCont(1).Terms = FORM$(9, 0)
    POCont(1).Addinst1 = FORM$(10, 0)
    POCont(1).Addinst2 = FORM$(11, 0)
    POCont(1).Addinst3 = FORM$(12, 0)
    POCont(1).Pading = ""
    PUT POFILE, 1, POCont(1)
    CLOSE POFILE
    RETURN




END SUB

SUB POMenu
	
	'$INCLUDE: 'SetCnf.bi'                    'Set monitor and color information
	
	TextCursor -2, -2             'Define the Mouse cursor
	'ShowCursor    'Show it
	
	DIM MenuText$(1 TO 7)
	MenuText$(1) = "Enter/Edit Purchase Orders"
	MenuText$(2) = "Print Purchase Order Journal"
	MenuText$(3) = "Print Purchase Order Forms"
	MenuText$(4) = "Post Entries"
	MenuText$(5) = "Cancel Open Purchase Order"
	MenuText$(6) = "Purchase Order Control"
	MenuText$(7) = "Quit"

	MaxLen = 0  'Set menu width to zero
	BoxBot = 18   'limit the box to line 20
	Action = 0    '0 means stay in the menu
	Choice = 1    'Pre-load choice to highlight
	
	FOR Cnt = 1 TO UBOUND(MenuText$)              'Find max menu width
		TLen = LEN(MenuText$(Cnt))
		IF TLen > MaxLen THEN
			MaxLen = TLen
		END IF
	NEXT
	
	'Center Menu within Screen
	Row = ((25 - (UBOUND(MenuText$))) \ 2) - 1
	Col = ((80 - MaxLen) \ 2) - 2
	
	ShowCursor
	
	DO
		LibFile2Scrn "GL.QSL", "MENUBAK", MonoCode, -1, ErrorCode
		Title$ = "Purchase Order Processing "

		MnuTitle Title$, Row, Col, MaxLen, 15
		PrintTitle User$
		PrintHelp "P.O. Processing"
		LOCATE Row, Col, 0
		VertMenu MenuText$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
		
		IF Ky$ = CHR$(27) THEN EXIT DO
		
		SELECT CASE Choice
		CASE 1
			ClrBG
			EdEntry
		CASE 2
			ClrBG
			PrnEditList
		CASE 3
			ClrBG
			PrintForms
		CASE 4
			ClrBG
			PostPOTrans
		CASE 5
			ClrBG
			CancelPO
		CASE 6
			ClrBG
			POControl
		CASE 7
			CLS : END
		CASE ELSE
		END SELECT
		
	LOOP

END SUB

SUB PostPOTrans

'  SHARED APAcct$

	SHARED VENDOR AS VendorRecType
	REDIM APDistRec(1) AS APDistRecType
	REDIM APLedgerRec(1) AS APLedger81RecType
	REDIM Tr2Post(1) AS GLTransRecType
	DIM PO AS POFORMRecType2

	LedgerRecLen = LEN(APLedgerRec(1))
	DistRecLen = LEN(APDistRec(1))
	VendorRecLen = LEN(VENDOR)

	'--Paint a background
	FOR Cnt = 2 TO 24
		QPrintRC STRING$(80, CHR$(177)), Cnt, 1, 3
	NEXT
	
	'--Verify that there are transactions
	OpenPOEditFile POEditFile, NumEdTrans
	
	'--Test for valid Vendor Numbers
	FOR Cnt = 1 TO NumEdTrans
		GET POEditFile, Cnt, PO
		IF NOT PO.DELETED THEN
			VendRecNum = FindVendorRec(PO.VNDRCODE)
			IF VendRecNum > 0 THEN
				Active = Active + 1
			ELSE
				AbortFlag = True
				COLOR 12: CLS
				BEEP
				PRINT "Error: Invalid Vendor Code"
				PRINT "Unable to locate vendor: "; PO.VNDRCODE
				PRINT "Operation Aborted."
				PRINT
				COLOR 7
				PRINT "Press any key to continue."
				k$ = INPUT$(1)
				EXIT FOR
			END IF
		END IF
	NEXT
	CLOSE
	
	IF AbortFlag THEN EXIT SUB

	'--Check for no active transactions
	IF Active = 0 THEN
		'--No active transactions - tell user and get out
		Ok = MsgBox%("GL.QSL", "NOTRANS")
		EXIT SUB
	END IF
	
	'--make sure we're ready to post
	Ok = MsgBox%("GL.QSL", "POSTGJ")
	IF Ok = 1 THEN EXIT SUB       '1=No
	
	REDIM FundList$(1)
	GetFundList FundList$(), NumFunds
	REDIM TrFundSum#(1 TO NumFunds)
	
	OpenPOEditFile POEditFile, NumEdTrans

	AP2Post = FREEFILE
	OPEN "APPOIF.DAT" FOR RANDOM AS AP2Post LEN = LEN(Tr2Post(1))
	
	RecordNum = 0 'Reset Active counter for posting
	
	FOR Cnt = 1 TO NumEdTrans     'number of invoices to process
		REDIM TrFundSum#(1 TO NumFunds)
		'Reinitialize transaction fund summary elements for next loop
		GET POEditFile, Cnt, PO
		
		IF NOT PO.DELETED THEN
			FOR AcctDist = 1 TO 12
				IF LEN(QPTrim$(PO.Items(AcctDist).ACCTNO)) THEN
					RecordNum = RecordNum + 1
					Tr2Post(1).AcctNum = PO.Items(AcctDist).ACCTNO
					Tr2Post(1).TrDate = PO.PODATE
					Tr2Post(1).Desc = PO.VNDRCODE
					Tr2Post(1).Ref = PO.PONUM
					Tr2Post(1).DrAmt = PO.Items(AcctDist).EXT
					Tr2Post(1).CrAmt = 0
					Tr2Post(1).Src = "PO" + ConvDateStr$(DATE$)
					PUT AP2Post, RecordNum, Tr2Post(1)
			
					'--Add this distribution to proper fund
					FOR Fund = 1 TO NumFunds
						FundNum$ = LEFT$(PO.Items(AcctDist).ACCTNO, FundLen)
						IF FundNum$ = FundList$(Fund) THEN
							TrFundSum#(Fund) = TrFundSum#(Fund) + Round#(PO.Items(AcctDist).EXT)
							'FundGrdTot#(Fund) = FundGrdTot#(Fund) + Round#(PO.Items(AcctDist).Ext)
							EXIT FOR
						END IF
					NEXT
			
				END IF  'test for blank distribution line
			NEXT      'Acct'g Distribution
		END IF      'not deleted
	NEXT          'transaction
	
	CLOSE
	
	'--common post & link sub in comnaux
	Post2PO "APPOIF.DAT", BadTrans%
	IF BadTrans = 0 THEN          'posted ok
		' MakeTempReport
		' KILL "APINVIF.dat"
	ELSE          '
		'--Couldn't find an account.
		'--Account was possibly deleted after entry made?

		COLOR 12: CLS
		BEEP
		PRINT "Error: PO Transaction Posting"
		PRINT "Unable to locate one or more G/L Accounts."
		PRINT "Unposted transactions listed in POPOST.LOG"
		PRINT
		COLOR 7
		PRINT "Press any key to continue."
		k$ = INPUT$(1)

	END IF

	'--Now post transaction to apledger.dat
	OpenPOEditFile POEditFile, NumEdTrans
	OpenAPLedgerFile APLedgerFile, NumLedgerRecs, LedgerRecLen
	OpenAPDistFile APDistFile, NumDistRecs&, DistRecLen
	OpenAPVendorFile APVendorFile, NumVRecs, VendorRecLen
	
	FOR Cnt = 1 TO NumEdTrans
		GET POEditFile, Cnt, PO
		IF NOT PO.DELETED THEN

			'--Update PO in vendor link to apledger
			NumLedgerRecs = NumLedgerRecs + 1
			VendRecNum = PO.VNDRREC
			OpenVendorFile VendorFile, NumVRecs
			GET VendorFile, VendRecNum, VENDOR
			
			PrevVendTrans = VENDOR.LastTran
			FrstVendTrans = VENDOR.FrstTran
			
			IF FrstVendTrans = 0 THEN
				VENDOR.LastTran = NumLedgerRecs
				VENDOR.FrstTran = NumLedgerRecs
			ELSE
				VENDOR.LastTran = NumLedgerRecs
				GET APLedgerFile, PrevVendTrans, APLedgerRec(1)
				APLedgerRec(1).NextTrans = NumLedgerRecs
				PUT APLedgerFile, PrevVendTrans, APLedgerRec(1)
			END IF

			PUT VendorFile, VendRecNum, VENDOR
			CLOSE VendorFile

			'--Post transaction to apledger
			APLedgerRec(1).VRecNum = PO.VNDRREC
			APLedgerRec(1).VendorCode = PO.VNDRCODE
			APLedgerRec(1).TrDate = PO.PODATE
			APLedgerRec(1).DOCNum = PO.PONUM
			APLedgerRec(1).PONUM = PO.PONUM
			APLedgerRec(1).TrCode = 4
			'APLedgerRec(1).PayCode = VAL(PO.PayCode)
			APLedgerRec(1).GLDistDate = PO.PODATE
			APLedgerRec(1).Amt = PO.POAMT
			APLedgerRec(1).NextTrans = 0
			APLedgerRec(1).FrstDist = NumDistRecs& + 1
			
			FOR D = 1 TO 12
				IF LEN(QPTrim$(PO.Items(D).ACCTNO)) THEN
					'--Post and link distributions
					REDIM APDistRec(1) AS APDistRecType
					NumDistRecs& = NumDistRecs& + 1
					APDistRec(1).APLedgerRec = NumLedgerRecs
					APDistRec(1).DistAcctRec = PO.Items(D).AcctRec
					APDistRec(1).DistAcctNum = PO.Items(D).ACCTNO
					APDistRec(1).DistAmt = PO.Items(D).EXT
					'APDistRec(1).DistCRAmt = 0
					APDistRec(1).NextDist = NumDistRecs& + 1
					PUT APDistFile, NumDistRecs&, APDistRec(1)
				END IF
			NEXT
			'Update the last distribution pointer
			GET APDistFile, NumDistRecs&, APDistRec(1)
			APDistRec(1).NextDist = 0
			PUT APDistFile, NumDistRecs&, APDistRec(1)
			
			APLedgerRec(1).LastDist = NumDistRecs&
			PUT APLedgerFile, NumLedgerRecs, APLedgerRec(1)

		END IF
	NEXT
	CLOSE

	'--Post ok.  Kill edit and interface files
	'IF BadTrans = 0 THEN
		KILL "APPED.DAT"
		KILL "APPOIF.DAT"
	'END IF

END SUB

SUB PrintForms

	LOCATE , , 0

	SHARED Choice$()
	REDIM Choice$(0 TO 3, 1)
	Choice$(0, 0) = "2"
	Choice$(1, 0) = "Screen"
	Choice$(2, 0) = "Prn #1"
	Choice$(3, 0) = "Prn #2"
	REDIM FORM$(0, 0)
	REDIM Fld(0) AS FieldInfo
	DIM Frm AS FormInfo
	LibName$ = "GL.QSL"
	FormName$ = "PRNGJ"
	Title$ = "PO Forms"
	FF$ = CHR$(12)

	LibLoadDisplayForm LibName$, FormName$, FORM$(), Fld()
	PrintHelp "Print PO Forms"
	Action = 1
	DO
		EditForm FORM$(), Fld(), Frm, Cnf, Action
		SELECT CASE Frm.KeyCode
			CASE F10Key 'SaveButton
				IF LEN(QPTrim$(FORM$(2, 0))) = 0 THEN
					Dev$ = "S"
					LPTNo = 1
				ELSE
					Dev$ = LEFT$(FORM$(2, 0), 1)
					LPTNo = VAL(RIGHT$(RTRIM$(FORM$(2, 0)), 1))
				END IF
			CASE ESC
				EXIT SUB
		END SELECT
	LOOP UNTIL Frm.KeyCode = F10Key
	SELECT CASE Dev$
		CASE "S"
			EntryPoint = 2
		CASE "P"
			EntryPoint = 5
	END SELECT
	
	'--Start printing forms
	DIM PO AS POFORMRecType2
	OpenPOEditFile POEditFile, NumEdTrans
	PrnFileName$ = "POFORMS.PRN"
	PrnFile = FREEFILE
	OPEN PrnFileName$ FOR OUTPUT AS #PrnFile

	FOR t = 1 TO NumEdTrans
		GET POEditFile, t, PO
		IF NOT PO.DELETED THEN
			ItemsPrinted = 0

			'PO.VndrCode
			'PO.VNDRREC
			
			'PRINT #PrnFile,
			'PRINT #PrnFile,
			
			PRINT #PrnFile, TAB(7); Num2Date(PO.PODATE);
				PRINT #PrnFile, TAB(66); PO.PONUM
			PRINT #PrnFile,
			PRINT #PrnFile,
			PRINT #PrnFile, TAB(7); PO.REQNUM
			PRINT #PrnFile,
			PRINT #PrnFile,
			PRINT #PrnFile, TAB(47); PO.SHPLINE1
			PRINT #PrnFile, TAB(47); PO.SHPLINE2
			PRINT #PrnFile, TAB(47); PO.SHPLINE3
			PRINT #PrnFile, TAB(47); PO.SHPLINE4
			PRINT #PrnFile, TAB(47); PO.SHPLINE5
			PRINT #PrnFile,
			PRINT #PrnFile,
			PRINT #PrnFile, TAB(54); PO.FOB
			PRINT #PrnFile, TAB(8); PO.VNDRINF1
			PRINT #PrnFile, TAB(8); PO.VNDRINF2;
				PRINT #PrnFile, TAB(54); PO.Shipvia
			PRINT #PrnFile, TAB(8); PO.VNDRINF3
			PRINT #PrnFile, TAB(8); PO.VNDRINF4;
				PRINT #PrnFile, TAB(54); PO.Terms
			PRINT #PrnFile, TAB(8); PO.VNDRINF5
			PRINT #PrnFile, TAB(54); PO.SHIPON
			PRINT #PrnFile,
			PRINT #PrnFile,
			PRINT #PrnFile,
			PRINT #PrnFile,
			PRINT #PrnFile,
			PRINT #PrnFile,

			'--Loop Thru Items
			DetailLines = 0
			FOR i = 1 TO 12
				IF LEN(QPTrim$(PO.Items(i).ACCTNO)) THEN
					PRINT #PrnFile, TAB(1); PO.Items(i).STKNO;
					PRINT #PrnFile, TAB(12); PO.Items(i).DESCA;
					IF PO.Items(i).QUAN > -999999.99# THEN
						PRINT #PrnFile, TAB(41); USING "####.##"; PO.Items(i).QUAN;
					END IF
					IF PO.Items(i).PRICE > -999999.99# THEN
						PRINT #PrnFile, TAB(51); USING "#####.##"; PO.Items(i).PRICE;
					END IF
					PRINT #PrnFile, TAB(62); USING "#####.##"; PO.Items(i).EXT;
					PRINT #PrnFile, TAB(72); LEFT$(StripAcct$(PO.Items(i).ACCTNO), 9)
					PRINT #PrnFile, TAB(12); PO.Items(i).DESCB
					DetailLines = DetailLines + 2
				END IF  'Active transaction test
				IF LEN(QPTrim$(PO.Items(i).ACCTNO)) = 0 AND LEN(QPTrim$(PO.Items(i).DESCA)) THEN
					PRINT #PrnFile, TAB(12); PO.Items(i).DESCA
					PRINT #PrnFile, TAB(12); PO.Items(i).DESCB
					DetailLines = DetailLines + 2
				END IF  'Active transaction test




			NEXT

			LinesToFinish = 25 - DetailLines
			FOR L = 1 TO LinesToFinish
				PRINT #PrnFile,
			NEXT

			PRINT #PrnFile,
			PRINT #PrnFile, TAB(61); USING "######.##"; PO.POAMT
			PRINT #PrnFile,
			PRINT #PrnFile, TAB(44); PO.Addinst1
			PRINT #PrnFile, TAB(44); PO.Addinst2
			PRINT #PrnFile, TAB(44); PO.Addinst3
			PRINT #PrnFile, FF$

		END IF
	NEXT

	CLOSE
	
	PrintRptFile Title$, PrnFileName$, LPTNo, RetCode%, EntryPoint
	
END SUB

SUB PrnEditList

	SHARED User$
	DebitCol = 42
	CreditCol = 58
	CommaFmt$ = "#########,.##"   'ten millions
	DistSumLine$ = "-------------"
	FF$ = CHR$(12)
	REDIM Title$(5)
	MaxLines = 56
	TransTotal# = 0
	TranCnt = 0
	ToPrint$ = SPACE$(78)

	SHARED Choice$()              'Multiple choice array
	REDIM Choice$(0 TO 3, 1)      'Choices for entry type field
	Choice$(0, 0) = "2"           'Field Number
	Choice$(1, 0) = "Screen"
	Choice$(2, 0) = "Prn #1"
	Choice$(3, 0) = "Prn #2"
	
	FileName$ = "APPOREG.PRN"
	RegTitle$ = "Purchase Orders"
	TranCol = CreditCol
	CashCol = DebitCol
	
	Title$(1) = RegTitle$
	Title$(2) = "Run Date: " + DATE$
	Title$(3) = ""
	Title$(4) = "Vendor                       Date        PO No                             Amt"
	Title$(5) = STRING$(78, "-")
	
	REDIM FORM$(0, 0)             'Holds data from form
	REDIM Fld(0) AS FieldInfo     'Field editing array
	
	'--Get user input on where to send report
	LibLoadDisplayForm "GL.QSL", "PRNGJ", FORM$(), Fld()
	PrintHelp "Print A/P Purchases Journal"
	
	DIM Frm AS FormInfo
	Action = 1
	
	'--Get the report options from user
	DO
		EditForm FORM$(), Fld(), Frm, Cnf, Action
		SELECT CASE Frm.KeyCode
		CASE F10Key 'SaveButton
			IF LEN(LTRIM$(RTRIM$(FORM$(2, 0)))) = 0 THEN
				Dev$ = "S"
				LPTNo = 1
			ELSE
				Dev$ = LEFT$(FORM$(2, 0), 1)
				LPTNo = VAL(RIGHT$(RTRIM$(FORM$(2, 0)), 1))
			END IF
		CASE ESC
			EXIT SUB
		END SELECT
	LOOP UNTIL Frm.KeyCode = F10Key
	
	'==Start processing register=>
	LOCATE 1, 1, 0

	DIM PO AS POFORMRecType2
	OpenPOEditFile POEditFile, NumEdTrans
	PrnFileNum = FREEFILE
	
	OPEN FileName$ FOR OUTPUT AS #PrnFileNum
	
	'--Get a list of active funds
	REDIM FundList$(1)
	GetFundList FundList$(), NumFunds
	'REDIM FundSum#(1 TO NumFunds)
	REDIM FundGrdTot#(1 TO NumFunds)
	
	GOSUB PrintHeader

	FOR Transaction = 1 TO NumEdTrans
		GET POEditFile, Transaction, PO
		
		IF NOT PO.DELETED THEN
			TranCnt = TranCnt + 1
			TransTotal# = TransTotal# + Round#(PO.POAMT)
			
			'--Print 1st Line - Transaction details
			LSET ToPrint$ = ""'SPACE$(78)

			LSET ToPrint$ = PO.VNDRCODE
			MID$(ToPrint$, 30) = ConvDateStr1$(Num2Date(PO.PODATE))
			MID$(ToPrint$, 42) = PO.PONUM
			MID$(ToPrint$, 66) = FUsing$(STR$(PO.POAMT), CommaFmt$)
			PRINT #PrnFileNum, ToPrint$
			
			'--Blank line between detail and acct'g distributions
			PRINT #PrnFileNum,
			
			'--Print Distribution Label
			LSET ToPrint$ = ""
			MID$(ToPrint$, 2) = "Accounting Distribution:"
			PRINT #PrnFileNum, ToPrint$
			
			'--Print Field Titles
			LSET ToPrint$ = ""
			MID$(ToPrint$, 4) = "Account Number                                         Distribution"
			PRINT #PrnFileNum, ToPrint$
			
			'--Print Accounting Distributions
			TotTranDist# = 0
			
			'--Loop Thru distributions to print and summarize
			FOR AcctDist = 1 TO 12
				IF LEN(QPTrim$(PO.Items(AcctDist).ACCTNO)) THEN

					'--Print this distribution
					LSET ToPrint$ = ""
					MID$(ToPrint$, 4) = PO.Items(AcctDist).ACCTNO
					MID$(ToPrint$, TranCol) = FUsing$(STR$(PO.Items(AcctDist).EXT), CommaFmt$)
					PRINT #PrnFileNum, ToPrint$

					TotTranDist# = TotTranDist# + PO.Items(AcctDist).EXT
					
					LineCnt = LineCnt + 1
					IF LineCnt >= MaxLines THEN
						PRINT #PrnFileNum, FF$
						GOSUB PrintHeader
					END IF

					'Sum by fund
					Found = False
					FOR Fund = 1 TO NumFunds
						FundNum$ = LEFT$(PO.Items(AcctDist).ACCTNO, FundLen)
						IF FundNum$ = FundList$(Fund) THEN
							Found = True
							FundGrdTot#(Fund) = FundGrdTot#(Fund) + Round#(PO.Items(AcctDist).EXT)
							EXIT FOR
						END IF
					NEXT
			
					IF NOT Found THEN  'Oh.Shit = True
						Ok = MsgBox("GL", "NOFUND")
					END IF
				
				END IF  'Active transaction test
				
			NEXT      'Distribution

			'--Summary line after last distribution
			LSET ToPrint$ = ""
			MID$(ToPrint$, TranCol) = DistSumLine$
			PRINT #PrnFileNum, ToPrint$

			'--Transaction Distribution Totals
			LSET ToPrint$ = ""
			MID$(ToPrint$, 4) = "Total Distributed"
			MID$(ToPrint$, TranCol) = FUsing$(STR$(TotTranDist#), CommaFmt$)
			PRINT #PrnFileNum, ToPrint$
			LineCnt = LineCnt + 2
			IF LineCnt >= MaxLines THEN
				PRINT #PrnFileNum, FF$
				GOSUB PrintHeader
			END IF
			
			'--2 blank lines before next distribution
			ToPrint$ = STRING$(78, "=")
			PRINT #PrnFileNum, ToPrint$
			PRINT #PrnFileNum,
			LineCnt = LineCnt + 2
				IF LineCnt >= MaxLines THEN
					PRINT #PrnFileNum, FF$
					GOSUB PrintHeader
				END IF
			END IF      'Not deleted test
		
	NEXT          'Transaction
	
	'PRINT #PrnFileNum, STRING$(80, "-")
	IF LineCnt > 45 THEN
		PRINT #PrnFileNum, FF$
	END IF

	'--Summary
	LSET ToPrint$ = ""
	LSET ToPrint$ = "File Totals:"
	PRINT #PrnFileNum, ToPrint$
	
	LSET ToPrint$ = ""
	LSET ToPrint$ = "Number of Transactions"
	MID$(ToPrint$, 31) = FUsing$(STR$(TranCnt), "####")
	PRINT #PrnFileNum, ToPrint$
	
	LSET ToPrint$ = ""
	LSET ToPrint$ = "Grand Totals"
	MID$(ToPrint$, 25) = FUsing$(STR$(TransTotal#), CommaFmt$)
	PRINT #PrnFileNum, ToPrint$
	
	PRINT #PrnFileNum,
	
	LSET ToPrint$ = ""
	LSET ToPrint$ = "Summary by Fund:"
	PRINT #PrnFileNum, ToPrint$
	
	FOR Cnt = 1 TO NumFunds
		IF FundGrdTot#(Cnt) > 0 THEN
			LSET ToPrint$ = ""
			LSET ToPrint$ = "Fund" + " " + FundList$(Cnt)
			MID$(ToPrint$, 25) = FUsing$(STR$(FundGrdTot#(Cnt)), CommaFmt$)
			PRINT #PrnFileNum, ToPrint$
			GrdTot# = GrdTot# + Round#(FundGrdTot#(Cnt))
		END IF
	NEXT
	
	LSET ToPrint$ = ""
	LSET ToPrint$ = "Total All Funds"
	MID$(ToPrint$, 25) = FUsing$(STR$(GrdTot#), CommaFmt$)
	PRINT #PrnFileNum, ToPrint$
	PRINT #PrnFileNum, FF$

	CLOSE

	SELECT CASE Dev$
	CASE "S"
		EntryPoint = 2
	CASE "P"
		EntryPoint = 5
	END SELECT

	PrintRptFile RegTitle$, FileName$, LPTNo, RetCode%, EntryPoint

EXIT SUB

PrintHeader:
	PRINT #PrnFileNum, User$
	FOR HCnt = 1 TO 5
		PRINT #PrnFileNum, Title$(HCnt)
	NEXT
	LineCnt = 6
RETURN

END SUB

SUB Rec2Form (FORM$(), RecordNumber%, Fld() AS FieldInfo, APType)
	
	OpenPOEditFile POEditFile, NumEdTrans
	
	GET POEditFile, RecordNumber, POEdit
	
	FORM$(0, 0) = SPACE$(LEN(POEdit))
	
	BCopy VARSEG(POEdit), VARPTR(POEdit), SSEG(FORM$(0, 0)), SADD(FORM$(0, 0)), LEN(FORM$(0, 0)), 0
	UnPackBuffer 0, 0, FORM$(), Fld()
	
	'Form$(1, 0) = STR$(POEdit.DelFlag)
	'Form$(2, 0) = Num2Date(POEdit.InvDate)
	'Form$(3, 0) = POEdit.InvNum
	'Form$(4, 0) = POEdit.Ref
	'Form$(5, 0) = STR$(POEdit.InvAMT)
	'Form$(6, 0) = POEdit.BatchNum
	'Form$(7, 0) = STR$(POEdit.TCode)
	
	CLOSE POEditFile              'close the file
	
	'Rec$ = SPACE$(5)
	'LSET Rec$ = STR$(RecordNumber%)
	'QPrintRC Rec$, 25, 67, -1
	
	
END SUB

SUB ResetForm (FORM$(), NumFlds)
	
	FOR F = 1 TO NumFlds
		FORM$(F, 0) = ""
	NEXT
	
END SUB

SUB SaveTrans (FORM$(), RecordNumber, EDMode, APType)
	
	OpenPOEditFile POEditFile, NumEdTrans
	
	IF EDMode = 1 THEN
		Rec = NumEdTrans + 1
	ELSE
		Rec = RecordNumber
	END IF
	
	BCopy SSEG(FORM$(0, 0)), SADD(FORM$(0, 0)), VARSEG(POEdit), VARPTR(POEdit), LEN(FORM$(0, 0)), 0
	
	PUT POEditFile, Rec, POEdit   'write it
	CLOSE POEditFile              'close the file

END SUB

SUB ShowEdMode (EDMode%)
	
	Status$ = SPACE$(24)
	F2Ed$ = "  F2=Edit  "
	F2New$ = "  F2=New   "
	ActiveColor = 112
	InActiveColor = 115
	
	SELECT CASE EDMode
	CASE 0
		LSET Status$ = "Edit an Existing Record"
		F2BtnCap$ = F2New$
		F5BtnColor = ActiveColor
		
	CASE 1
		LSET Status$ = "Add a New Record"
		F2BtnCap$ = F2Ed$
		F5BtnColor = InActiveColor
		
	END SELECT
	
	QPrintRC Status$, 25, 2, -1
	QPrintRC F2BtnCap$, 25, 40, -1
	MPaintBox 25, 53, 25, 63, F5BtnColor          'F5 Button enable / disable
	
END SUB

SUB ValidateForm (FORM$(), AcctFlds(), BadFld)
	'Use this sub to validate fields

	'--Check PONum
	IF LEN(QPTrim$(FORM$(2, 0))) = 0 THEN
		BadFld = 2
		EXIT SUB
	END IF

	'--Check Date
	IF Date2Num(FORM$(4, 0)) < -32000 THEN
		BadFld = 4
		EXIT SUB
	END IF
	
	'--Check Distributions
	FOR F = 1 TO 12  'UBOUND(AcctFlds, 1)
		AcctFldNum = AcctFlds(F)
		AmtFldNum = AcctFldNum - 1

		'--check for an account number without a distribution amount
		'IF LEN(Form$(AcctFldNum, 0)) > 0 THEN
		'  IF Value(Form$((AmtFldNum), 0), ECode) = 0 THEN
		'    BadFld = AcctFldNum
		'    EXIT FOR
		'  END IF
		'END IF

		'--check for a distribution without an account number
		IF Value(FORM$(AmtFldNum, 0), ECode) <> 0 THEN
			IF LEN(QPTrim$(FORM$(AcctFldNum, 0))) = 0 THEN
				BadFld = (AcctFldNum - 1)
			END IF
		END IF
	NEXT

	
END SUB

SUB VoidPOTrans (VoidTransRecNum&)

	REDIM APLedger(1) AS APLedger81RecType
	REDIM DistRec(1) AS APDistRecType
	LdRecLen = LEN(APLedger(1))
	DistRecLen = LEN(DistRec(1))

	OpenAPLedgerFile APLedgerFile, NumTrans, LdRecLen
	GET APLedgerFile, VoidTransRecNum&, APLedger(1)
	APLedger(1).TrCode = -4
	PUT APLedgerFile, VoidTransRecNum&, APLedger(1)
	CLOSE APLedgerFile

	POIFFile$ = "POVDIF.DAT"
	KillFile POIFFile$
	REDIM GLIFRec(1) AS GLTransRecType
	GLIFRecLen = LEN(GLIFRec(1))
	GLIFFile = FREEFILE
	OPEN POIFFile$ FOR RANDOM AS GLIFFile LEN = GLIFRecLen
	OpenAPDistFile APDistFile, NumDistRecs&, DistRecLen
	NextDist& = APLedger(1).FrstDist
	DistCnt = 0

	DO UNTIL NextDist& = 0
		GET APDistFile, NextDist&, DistRec(1)
		IFRec = IFRec + 1

		'--Make Debit side of entry
		GLIFRec(1).Src = "VD" + ConvDateStr$(DATE$)
		AcctNum$ = LEFT$(DistRec(1).DistAcctNum, FundLen) + EncAcct$
		GLIFRec(1).AcctNum = AcctNum$
		GLIFRec(1).TrDate = Date2Num(DATE$)
		GLIFRec(1).Desc = "CANCELLED PO"  'APLedger(1).PONum
		GLIFRec(1).Ref = APLedger(1).PONUM
		GLIFRec(1).CrAmt = 0
		GLIFRec(1).DrAmt = DistRec(1).DistAmt
		PUT GLIFFile, IFRec, GLIFRec(1)

		IFRec = IFRec + 1
		'AcctNum$ = LEFT$(DistRec(1).DistAcctNum, FundLen) + APAcct$
		'GLIFRec(1).AcctNum = AcctNum$
		GLIFRec(1).AcctNum = DistRec(1).DistAcctNum
		GLIFRec(1).CrAmt = DistRec(1).DistAmt
		GLIFRec(1).DrAmt = 0
		PUT GLIFFile, IFRec, GLIFRec(1)

		NextDist& = DistRec(1).NextDist

	LOOP

	CLOSE

	Post2PO POIFFile$, BadAcct
	KillFile POIFFile$
	
END SUB

SUB XXALTF3

		'--This is just holding this code temporarily
		'This chunk of code deletes a distribution line when Alt+F3 is pressed
		'This needs to occur before prev fld test
		'IF Frm(1).FldNo <> 7 AND Frm(1).PrevFld = 7 AND LEN(QPTrim$(Form$(7, 0))) = 0 THEN
		'  Frm(1).FldNo = 7
		'  Action = 2
		'  Frm(1).KeyCode = 0
		'END IF

		'SELECT CASE Frm(1).KeyCode '!!!
		'CASE -106   'Alt+F3
		'  IF Frm(1).FldNo > Header THEN     'if we're on an account field
		'    'for each accounting distribution
		'    FOR Cnt = 1 TO 36
		'      'if this acct field is less than the current field
		'      'set the row counter
		'      IF AcctFlds(Cnt) <= Frm(1).FldNo THEN
		'        FldRow = FldRow + 1 '+AcctFlds(Cnt)
		'      ELSE
		'        FOR Cnt3 = AcctFlds(FldRow) TO AcctFlds(FldRow) + 3
		'          LSET Form$(Cnt3, 0) = ""
		'        NEXT
		'
		'        FOR Cnt3 = FldRow TO 35
		'          LSET Form$(AcctFlds(Cnt3), 0) = Form$(AcctFlds(Cnt3 + 1), 0)
		'          LSET Form$(AcctFlds(Cnt3) + 1, 0) = Form$((AcctFlds(Cnt3 + 1) + 1), 0)
		'          LSET Form$(AcctFlds(Cnt3) + 2, 0) = Form$((AcctFlds(Cnt3 + 1) + 2), 0)
		'          LSET Form$(AcctFlds(Cnt3) + 3, 0) = Form$((AcctFlds(Cnt3 + 1) + 3), 0)
		'        NEXT
		'        'Cnt is at 36 after 35 times thru above loop
		'
		'        LSET Form$(AcctFlds(Cnt3), 0) = ""
		'        LSET Form$(AcctFlds(Cnt3) + 1, 0) = ""
		'        LSET Form$(AcctFlds(Cnt3) + 2, 0) = ""
		'        LSET Form$(AcctFlds(Cnt3) + 3, 0) = ""
		'
		'        Frm(1).FldNo = AcctFlds(FldRow)
		'        Frm(1).PrevFld = Frm(1).FldNo - 1
		'        FldRow = 0
		'        Action = 1
		'        EXIT FOR
		'      END IF
		'    NEXT
		'  END IF


END SUB

