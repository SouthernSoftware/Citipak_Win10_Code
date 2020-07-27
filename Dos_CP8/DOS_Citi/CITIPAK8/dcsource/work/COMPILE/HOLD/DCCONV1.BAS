DEFINT A-Z
DECLARE SUB OpenDCCustIdxFile (NumOfARIdxRecs%, ARIdxFile%)
DECLARE SUB OpenDCCust (NumOfDCRecs%, DCFile%)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB HideCursor ()
DECLARE SUB QPrint (x$, Colr%, page%)
DECLARE SUB QPrintRC (T$, r%, C%, clr%)
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

' GORDONSVILLE VA CONVERSION

  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'DC.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'

   DIM SHARED DCCust(1) AS DCCustRecType
   DIM SHARED DCVRec(1) AS DCVehType
   DIM SHARED OLDDCCust(1) AS OldCustType
   DIM SHARED OldTran(1) AS OLDCUSTTRANSTYPE

  'Open Vehicle File
  DCVehRecLen = LEN(DCVRec(1))
  DCVFile = FREEFILE
  OPEN "DCVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCVFile LEN = DCVehRecLen
  NumOfVrecs = LOF(DCVFile) \ DCVehRecLen

  'Open Customer File
  OpenDCCust NumOfDCRecs, DCFile
  DCTransRecLen = LEN(DCTransRec(1))

  'Open Old Customer Data File
  ODCFile = FREEFILE
  OPEN "DBCust.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ODCFile LEN = 256
  NumOfODCRecs = LOF(ODCFile) \ 256

  'Open Old Customer Trans File
  TFILE = FREEFILE
  OPEN "DBARTRAN.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TFILE LEN = 256

 CLS

  FOR Cnt! = 1 TO NumOfODCRecs
  GET ODCFile, Cnt!, OLDDCCust(1)
  LOCATE 10, 1: PRINT "Grabing Record # "; Cnt!

    'Set New Values
    Nme$ = RTRIM$(OLDDCCust(1).Firstname)
    Nme$ = Nme$ + " " + RTRIM$(OLDDCCust(1).Lname)
    Nme$ = RTRIM$(Nme$)

    'No Name No Account

  IF LEN(Nme$) > 0 THEN
    IF OLDDCCust(1).FIRSTTR > 0 THEN

    GOSUB ChkCustomer
    IF Good97 = 1 OR Good98 = 1 THEN

    NextRecord! = (LOF(DCFile) \ LEN(DCCust(1))) + 1

    DCCust(1).CUSTNUMB = LTRIM$(STR$(NextRecord!))
    DCCust(1).SORTNAME = OLDDCCust(1).Lname
    DCCust(1).BILLNAME = Nme$
    DCCust(1).ADDRESS1 = OLDDCCust(1).ADD1
    DCCust(1).ADDRESS2 = ""
    DCCust(1).CITY = OLDDCCust(1).CITY
    DCCust(1).STATE = OLDDCCust(1).ST
    DCCust(1).ZIPCODE = OLDDCCust(1).ZP
    DCCust(1).SOSEC = OLDDCCust(1).SSNUM
    DCCust(1).DRVLIC = OLDDCCust(1).statelic
    DCCust(1).DATEOPED = -32767
    DCCust(1).CASHONLY = "N"
    DCCust(1).Resident = "Y"
    DCCust(1).Owner = "Y"
    DCCust(1).HPHONE = OLDDCCust(1).PHONE
    DCCust(1).WPHONE = ""
    DCCust(1).LICENSE = ""
    DCCust(1).Valid = 1
    DCCust(1).AcctBal = 0
    DCCust(1).Deleted = "N"
    DCCust(1).FirstTrans = 0
    DCCust(1).LastTrans = 0
    DCCust(1).FirstCar = 0
    DCCust(1).LastCar = 0
    DCCust(1).RoomtoGrow = ""
    PUT DCFile, NextRecord!, DCCust(1)
    IF Good98 = 1 THEN GOSUB AddVehicle98
    IF Good98 = 0 AND Good97 = 1 THEN GOSUB AddVehicle97
  
  END IF
  END IF
  END IF

  NEXT Cnt!
  CLOSE
  END


AddVehicle97:
 Trans! = OLDDCCust(1).FIRSTTR

 WHILE Trans! > 0
    GET TFILE, Trans!, OldTran(1)

    IF RIGHT$(OldTran(1).tdate, 2) = "97" THEN
    IF OldTran(1).TType >= 1 AND OldTran(1).TType <= 6 THEN
    IF OldTran(1).TType = 1 THEN
     DCVRec(1).DecalCat = "CAR"
    END IF
    IF OldTran(1).TType = 2 THEN
     DCVRec(1).DecalCat = "MTR"
    END IF
    IF OldTran(1).TType = 3 THEN
     DCVRec(1).DecalCat = "OTH"
    END IF
    IF OldTran(1).TType = 4 THEN
     DCVRec(1).DecalCat = "TRK"
    END IF
    IF OldTran(1).TType = 5 THEN
     DCVRec(1).DecalCat = "RV "
    END IF
    IF OldTran(1).TType = 6 THEN
     DCVRec(1).DecalCat = "TRX"
    END IF

    DCVRec(1).Makemodel = OldTran(1).Tmake
    DCVRec(1).StateTag = OldTran(1).tstlic
    DCVRec(1).ExpireDate = Date2Num%("02-28-1998")
    DCVRec(1).Sticker = OldTran(1).tsticker
    DCVRec(1).Valid = "Y"
    DCVRec(1).Active = "Y"
    DCVRec(1).Notes = ""
    DCVRec(1).Desc = OldTran(1).tid
    DCVRec(1).Fee = OldTran(1).tamt
    DCVRec(1).MasterRecord = NextRecord!
    DCVRec(1).NextRec = 0
    DCVRec(1).MoreRoom = ""
    NextCar! = LOF(DCVFile) \ LEN(DCVRec(1)) + 1
    PUT DCVFile, NextCar!, DCVRec(1)

    IF DCCust(1).FirstCar = 0 THEN
     DCCust(1).FirstCar = NextCar!
     DCCust(1).LastCar = NextCar!
     PUT DCFile, NextRecord!, DCCust(1)
    ELSE
     PrevRec! = DCCust(1).LastCar
     DCCust(1).LastCar = NextCar!
     PUT DCFile, NextRecord!, DCCust(1)

     GET DCVFile, PrevRec!, DCVRec(1)
     DCVRec(1).NextRec = NextCar!
     PUT DCVFile, PrevRec!, DCVRec(1)

    END IF
    END IF
    END IF
    Trans! = OldTran(1).tnexttr
 WEND
RETURN

AddVehicle98:
 Trans! = OLDDCCust(1).FIRSTTR

 WHILE Trans! > 0
    GET TFILE, Trans!, OldTran(1)

    IF RIGHT$(OldTran(1).tdate, 2) = "98" THEN
    IF OldTran(1).TType >= 1 AND OldTran(1).TType <= 6 THEN
    IF OldTran(1).TType = 1 THEN
     DCVRec(1).DecalCat = "CAR"
    END IF
    IF OldTran(1).TType = 2 THEN
     DCVRec(1).DecalCat = "MTR"
    END IF
    IF OldTran(1).TType = 3 THEN
     DCVRec(1).DecalCat = "OTH"
    END IF
    IF OldTran(1).TType = 4 THEN
     DCVRec(1).DecalCat = "TRK"
    END IF
    IF OldTran(1).TType = 5 THEN
     DCVRec(1).DecalCat = "RV "
    END IF
    IF OldTran(1).TType = 6 THEN
     DCVRec(1).DecalCat = "TRX"
    END IF

    DCVRec(1).Makemodel = OldTran(1).Tmake
    DCVRec(1).StateTag = OldTran(1).tstlic
    DCVRec(1).ExpireDate = Date2Num%("02-28-1999")
    DCVRec(1).Sticker = OldTran(1).tsticker
    DCVRec(1).Valid = "Y"
    DCVRec(1).Active = "Y"
    DCVRec(1).Notes = ""
    DCVRec(1).Desc = OldTran(1).tid
    DCVRec(1).Fee = OldTran(1).tamt
    DCVRec(1).MasterRecord = NextRecord!
    DCVRec(1).NextRec = 0
    DCVRec(1).MoreRoom = ""
    NextCar! = LOF(DCVFile) \ LEN(DCVRec(1)) + 1
    PUT DCVFile, NextCar!, DCVRec(1)

    IF DCCust(1).FirstCar = 0 THEN
     DCCust(1).FirstCar = NextCar!
     DCCust(1).LastCar = NextCar!
     PUT DCFile, NextRecord!, DCCust(1)
    ELSE
     PrevRec! = DCCust(1).LastCar
     DCCust(1).LastCar = NextCar!
     PUT DCFile, NextRecord!, DCCust(1)

     GET DCVFile, PrevRec!, DCVRec(1)
     DCVRec(1).NextRec = NextCar!
     PUT DCVFile, PrevRec!, DCVRec(1)

    END IF
    END IF
    END IF
    Trans! = OldTran(1).tnexttr
 WEND
RETURN


ChkCustomer:
Good97 = 0
Good98 = 0

 Trans! = OLDDCCust(1).FIRSTTR

 WHILE Trans! > 0
    GET TFILE, Trans!, OldTran(1)
    IF RIGHT$(OldTran(1).tdate, 2) = "97" THEN
    Good97 = 1
    END IF
    IF RIGHT$(OldTran(1).tdate, 2) = "98" THEN
    Good98 = 1
    END IF

    Trans! = OldTran(1).tnexttr
 WEND



RETURN


END

SUB OpenDCCust (NumOfDCRecs%, DCFile%)
  CLOSE DCFile

  DCCustRecLen = LEN(DCCust(1))
  DCFile = FREEFILE
  OPEN "DCCust.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCFile LEN = DCCustRecLen
  NumOfDCRecs = LOF(DCFile) \ DCCustRecLen
  'FOR x = 1 TO NumOfDCRecs
  'GET DCFile, x, DCCust(1)
  'PRINT DCCust(1).Custnumb; TAB(15); DCCust(1).FirstTrans
  'SLEEP 1
  'NEXT x
  'STOP

END SUB

