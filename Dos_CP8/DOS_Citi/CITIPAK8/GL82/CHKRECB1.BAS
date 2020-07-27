DECLARE SUB ConvertBack ()
DEFINT A-Z

'$INCLUDE: 'oschk.bi'
'$INCLUDE: 'fileview.bi'
'$INCLUDE: 'workdec.bi'

'Subs & Functions in this module
DECLARE SUB EnterClearedChks ()
DECLARE SUB PrintClearedChks ()
DECLARE SUB PrintChkList (Title$, WhatKind%)
DECLARE SUB PrintTitleScrn (Title$)
DECLARE SUB AddChecks ()
DECLARE SUB GetFile ()
DECLARE SUB SortFile ()
DECLARE SUB util ()
DECLARE SUB PurgeClearedChks ()
DECLARE FUNCTION Round# (N#)
DECLARE SUB EnterClearedChkRange ()
DECLARE SUB util2 ()
DECLARE SUB util1 ()

'procedures in other modules
DECLARE SUB TextIn (t$, Max, NumOnly, CapsOn, ExitCode, Colr)
DECLARE FUNCTION InputBox$ (LibName$, ScrnName$, Row, Col, Max, NumOnly)
DECLARE SUB DateIn (Dt$, ExitCode, Colr)
DECLARE SUB YesNoB (YN$, ExitCode, Colr)
DECLARE SUB Prn2Prn (ReportFile$, Title$, TitLines%, MaxLines%, LPTNo%, ExitFlag%)
DECLARE SUB Prn2Scrn (FileName$, TopRow%, BotRow%, clr%, ExitFlag%)

DIM SHARED OSChk AS OSChkRecType1

'ConvertBack
'util2
'STOP

CLS
'ON ERROR GOTO GotErr

DO

   CLS
   COLOR 11
   LOCATE 1, 1
   COLOR 10
   PRINT "Sys Date: "; DATE$; ""
   
   LOCATE 1, 75
   PRINT "Menu"
   
   LOCATE 2, 1: COLOR 10
   PRINT STRING$(80, CHR$(196))

   COLOR 11
   LOCATE 4, 1
   PRINT TAB(27); " Check Reconciliation Menu"
   PRINT
   PRINT
   PRINT TAB(27); "1. Enter Cancelled Checks"
   PRINT
   PRINT TAB(27); "2. Enter a Range of Cancelled Checks"
   PRINT
   PRINT TAB(27); "3. Print Cancelled Check List"
   PRINT
   PRINT TAB(27); "4. Remove Cancelled Checks"
   PRINT
   PRINT TAB(27); "5. Print Outstanding Check List"
   PRINT
   PRINT TAB(27); "6. Add Outstanding Checks to File"
   PRINT
   PRINT TAB(27); "7. Sort Outstanding Check File"
   PRINT
   PRINT TAB(27); "8. Print Check File"
   PRINT
   PRINT TAB(27); " Selection No. (Esc End) :";
   COLOR 14: PRINT TAB(54); "_"
   
   Choice$ = INPUT$(1)

   SELECT CASE Choice$

      CASE "1"
         EnterClearedChks
      CASE "2"
         EnterClearedChkRange
      CASE "3"
         PrintTitleScrn "Cancelled Check Listing"
         PrintChkList "Cancelled Check Listing", 1
      CASE "4"
         PurgeClearedChks
      CASE "5"
         PrintTitleScrn "Outstanding Check Listing"
         PrintChkList "Outstanding Check Listing", 0
      CASE "6"
         PrintTitleScrn "Add Outstanding Checks to File"
         AddChecks
      CASE "7"
         PrintTitleScrn "Sort Outstanding Check File"
         LOCATE 4, 1: COLOR 14
         PRINT "Sorting File.  Please wait."
         SortFile
         LOCATE 4, 16
         PRINT "Sorting complete. Press any key to continue."
         K$ = INPUT$(1)
      CASE "8"
         util2
      'CASE ELSE
      
   END SELECT

LOOP UNTIL Choice$ = CHR$(27)

CLOSE
'RUN "cpmenu"

GotErr:
   ErrNum = ERR
   SELECT CASE ErrNum
      CASE 9
         CLS
         VIEW PRINT
         LOCATE 12, 1: COLOR 12
         PRINT "Not enough memory to sort checks."
         PRINT "Delete any old checks from system and try again"
         PRINT "Press any key to return to the menu."
         K$ = INPUT$(1)
         VIEW PRINT
         CLS : CLOSE : RUN "CPMENU"
      CASE 70
         CLS
         VIEW PRINT
         LOCATE 12, 1: COLOR 12
         PRINT "Access denied. Outstanding check file is in use by another user."
         PRINT "Press any key to return to the menu."
         K$ = INPUT$(1)
         VIEW PRINT
         CLS : CLOSE : RUN "CPMENU"
       CASE ELSE
         CLS
         VIEW PRINT
         LOCATE 12, 1: COLOR 12
         PRINT "An unexpected error has halted this program."
         PRINT "Module Code: CHKREC"
         PRINT USING "Error Code:  ##"; ERR
         PRINT "Press any key to exit to DOS"
         K$ = INPUT$(1)
         CLS : CLOSE : END
       END SELECT

SUB AddChecks

SHARED FileNum, RecLen, NumRecs

PrintTitleScrn "Enter Outstanding Checks"

GetFile 'Call the GetFile routine to open the crchk.dat file.
LOCK FileNum

DrawScreen:
   COLOR 11
   LOCATE 4, 1
   PRINT "Bank Code    :"
   PRINT "Check Number :"
   PRINT "Check Date   :"
   PRINT "Payee        :"
   PRINT "Check Amt    :"
   PRINT "Type (A,P)   :"

   LOCATE 11, 1
   COLOR 14
   PRINT "<Esc> to cancel."
   
   Colr = 14 'Data editing color



GetNewBankNum:
   LOCATE 4, 16
   Max = 1: NumOnly = 1
   TextIn BankNum$, Max, NumOnly, CapsOn, ExitCode, Colr
   IF ExitCode = 2 THEN GOTO Bye

   BNum = VAL(BankNum$)
   IF BNum < 0 OR BNum > 4 THEN
      PLAY "L32O3EC"
      GOTO GetNewBankNum
   END IF
   
GetNum:
   LOCATE 5, 16
   Max = 6: NumOnly = 1
   TextIn CheckNum$, Max, NumOnly, CapsOn, ExitCode, Colr
   IF ExitCode = 1 THEN GOTO GetNewBankNum
   IF ExitCode = 2 THEN GOTO Bye

   ChkNum& = VAL(CheckNum$)
   IF ChkNum& = 0 THEN
      PLAY "L32O3EC"
      GOTO GetNum
   END IF

   Found = 0
   GOSUB CheckChkNum
   IF NOT Found THEN
      GOTO GetDte
   ELSE
      GOTO GetNum
   END IF

GetDte:
   LOCATE 6, 16
   DateIn Dt$, ExitCode, Colr
   IF ExitCode = 1 THEN GOTO GetNum
   IF ExitCode = 2 THEN GOTO Bye
   IF LEN(RTRIM$(Dt$)) = 0 THEN
      PLAY "L32O3EC"
      GOTO GetDte
   END IF

GetDesc:
   LOCATE 7, 16
   Max = 20: NumOnly = 0
   TextIn Payee$, Max, NumOnly, CapsOn, ExitCode, Colr
   IF ExitCode = 1 THEN GOTO GetDte
   IF ExitCode = 2 THEN GOTO Bye

GetAmt:
   LOCATE 8, 16: ClearEOL 7
   Max = 10: NumOnly = 1
   TextIn CheckAmt$, Max, NumOnly, CapsOn, ExitCode, Colr
   ChkAmt# = Round#(VAL(CheckAmt$))
   LOCATE 7, 16: COLOR 14: PRINT USING "$#######,.##"; ChkAmt#
   IF ExitCode = 1 THEN GOTO GetDesc
   IF ExitCode = 2 THEN GOTO Bye
   IF ChkAmt# <= 0 THEN
      PLAY "L32O3EC"
      GOTO GetAmt
   END IF

GetChkType:
   LOCATE 9, 16: ClearEOL 7
   Max = 1: NumOnly = 0: CapsOn = 1
   TextIn CheckType$, Max, NumOnly, CapsOn, ExitCode, Colr
   IF NOT (CheckType$ = "A" OR CheckType$ = "P") THEN GOTO GetChkType
   SELECT CASE CheckType$
      CASE "P"
         ChkType = 1
         BNum = 0
      CASE "A"
         ChkType = 0
      CASE ELSE
   END SELECT
   IF ExitCode = 1 THEN GOTO GetAmt
   IF ExitCode = 2 THEN GOTO Bye

'Check before saving
   LOCATE 11, 1
   COLOR 10
   PRINT "Save this check? (Y/n) ";
   X$ = ""
   YesNoB X$, XCode, Colr
   LOCATE , 1
   ClearEOL 7
   IF XCode = 1 THEN GOTO GetAmt
   IF XCode = 2 THEN GOTO Bye
   IF X$ = "N" GOTO GetNum
   
'Save record & old data and jump back to start
   ChkDt$ = LEFT$(Dt$, 2) + "/" + MID$(Dt$, 3, 2) + "/" + RIGHT$(Dt$, 2)
   
   NumRecs = NumRecs + 1

   OSChk.BankCode = BNum
   OSChk.ChkNum = ChkNum&
   OSChk.ChkDate = ChkDt$
   OSChk.Desc = Payee$
   OSChk.Amt = ChkAmt#
   OSChk.Cleared = 0
   OSChk.Src = ChkType

   PUT FileNum, NumRecs, OSChk

   BankNum$ = ""
   CheckNum$ = ""          'Clear data entry fields.
   Dt$ = ""
   Payee$ = ""
   CheckAmt$ = ""
   CheckType$ = ""

   FOR LineNum = 4 TO 9
      LOCATE LineNum, 16
      ClearEOL 7
   NEXT

   GOTO GetNewBankNum

Bye:
UNLOCK FileNum
CLOSE
EXIT SUB

CheckChkNum:
   FOR Cnt = 1 TO NumRecs   'Do a sequential search for the check number.
      GET FileNum, Cnt, OSChk
      IF OSChk.ChkNum = ChkNum& THEN
         Found = -1
         LOCATE 12, 1
         COLOR 12
         PRINT USING "Check No. ###### is already in the file."; ChkNum&
         PRINT "Press any key to try again."
         K$ = INPUT$(1)
         LOCATE 12, 1: ClearEOL 7
         LOCATE 13, 1: ClearEOL 7
         LOCATE 4, 16: ClearEOL 7
         CheckNum$ = ""
         EXIT FOR
      END IF
   NEXT

RETURN

END SUB

SUB ConvertBack

SHARED FileNum, RecLen, NumRecs
PrintTitleScrn "Special for Alleghany County"
GetFile

DIM OSChkA AS OSChkRecType
OldRecLen = LEN(OSChkA)
OldFileNum = FREEFILE
OPEN "osconv.dat" FOR RANDOM ACCESS READ WRITE SHARED AS OldFileNum LEN = OldRecLen
'NumOldRecs = LOF(OldFileNum) \ OldRecLen

FOR i = 1 TO NumRecs
  GET FileNum, i, OSChk
  LOCATE 3, 1: PRINT "Converting: "; i
  OSChkA.ChkNum = OSChk.ChkNum
  OSChkA.ChkDate = OSChk.ChkDate
  OSChkA.Desc = OSChk.Desc
  OSChkA.Amt = OSChk.Amt
  OSChkA.Src = OSChk.Src
  OSChkA.Cleared = OSChk.Cleared
  OSChkA.BankCode = OSChk.BankCode
  OSChkA.Filler = OSChk.Filler

  PUT OldFileNum, i, OSChkA

NEXT
LOCATE 3, 1: PRINT "Conversion Complete.                         "
CLOSE

END SUB

SUB EnterClearedChkRange

SHARED FileNum, RecLen, NumRecs

CLS
PrintTitleScrn "Enter Cancelled Checks"
VIEW PRINT 3 TO 23

GetFile
LOCK FileNum

DO
GetCheckNumRange:
   CLS
   COLOR 11
   LOCATE 4, 1
   PRINT "Check Numbers:        thru "

   LOCATE 11, 1
   COLOR 14
   PRINT "Press Esc to End."

   LOCATE 4, 16
   Max = 6: NumOnly = 1: Colr = 14
   LCheckNum$ = ""
   TextIn LCheckNum$, Max, NumOnly, CapsOn, ExitCode, Colr
   IF ExitCode = 2 THEN
      VIEW PRINT
      EXIT DO
   END IF
   
   LOCATE 4, 28
   Max = 6: NumOnly = 1: Colr = 14
   HCheckNum$ = ""
   TextIn HCheckNum$, Max, NumOnly, CapsOn, ExitCode, Colr
   IF ExitCode = 1 GOTO GetCheckNumRange
   IF ExitCode = 2 THEN
      VIEW PRINT
      EXIT DO
   END IF

   LChkNum& = VAL(LCheckNum$)
   HChkNum& = VAL(HCheckNum$)
   Found = 0

   FOR Cnt = 1 TO NumRecs

      GET FileNum, Cnt, OSChk
      'IF OSChk.ChkNum = 30458 THEN STOP
      IF OSChk.ChkNum >= LChkNum& AND OSChk.ChkNum <= HChkNum& THEN
         OSChk.Cleared = 1
         PUT FileNum, Cnt, OSChk
      END IF

   NEXT
LOOP

GByeR:
UNLOCK FileNum
VIEW PRINT
CLOSE

'GotErr:
'   ErrorNum = ERR
'   SELECT CASE ErrorNum
'      CASE 70
'         CLS
'         LOCATE 12, 1: COLOR 14
'         PRINT "Another user has this file locked."
'         PRINT "Press any key to exit"
'         K$ = INPUT$(1)
'         CLS : CLOSE : END
'       CASE ELSE
'         CLS
'         LOCATE 12, 1: COLOR 14
'         PRINT USING "Error ## occured."; ERR
'         PRINT "Press any key to exit"
'         K$ = INPUT$(1)
'         CLS : CLOSE : END
'   END SELECT


END SUB

SUB EnterClearedChks
SHARED FileNum, RecLen, NumRecs

Colr = 14

CLS
PrintTitleScrn "Enter Cancelled Checks"

   COLOR 11
   LOCATE 4, 1
   PRINT "Bank Number : "

CnclGetBank:
   LOCATE 4, 15:
   BankNo$ = "": Max = 1: CapsOn = 1: NumOnly = 1
   TextIn BankNo$, Max, NumOnly = 1, CapsOn, ExitCode, Colr
   'IF ExitCode = 1 THEN GOTO GetType
   IF ExitCode = 2 THEN GOTO ClrGetout
   BankNum = VAL(BankNo$)
   
   
CLS
PrintTitleScrn "Enter Cancelled Checks"
VIEW PRINT 3 TO 23

GetFile
LOCK FileNum

DO
GetCheckNum:

   CLS
   QPrintRC "Bank: " + BankNo$, 3, 1, -1

   COLOR 11
   LOCATE 4, 1
   PRINT "Check Number : "

   LOCATE 11, 1
   COLOR 14
   PRINT "Press Esc to End."

   LOCATE 4, 17
   Max = 6: NumOnly = 1: Colr = 14
   CheckNum$ = ""
   TextIn CheckNum$, Max, NumOnly, CapsOn, ExitCode, Colr

   IF ExitCode = 2 THEN
      VIEW PRINT
      GOTO GBye
   END IF

   ChkNum& = VAL(CheckNum$)
   
   Found = 0

   FOR Cnt = 1 TO NumRecs   'Do a seq search for the check number.

         GET FileNum, Cnt, OSChk

         IF OSChk.ChkNum = ChkNum& AND OSChk.BankCode = BankNum THEN

            Found = -1
            
            LOCATE 5, 1
            PRINT "Check Date   : "; OSChk.ChkDate
            PRINT "Description  : "; OSChk.Desc
            PRINT "Check Amt    : "; LTRIM$(FUsing(STR$(OSChk.Amt), "########,.##"))
            PRINT "Status       : ";
            
            SELECT CASE OSChk.Cleared
               CASE 1
                  PRINT "Cleared    "
                  PRINT
                  COLOR 14: PRINT "Check is marked as cleared."
                  COLOR 12: PRINT "Change to outstanding (N/y) ";

                  X$ = ""
                  YesNoB X$, XCode, Colr
                  IF XCode = 2 THEN EXIT FOR 'Esc Pressed
               
                  IF X$ = "Y" THEN
                     OSChk.Cleared = 0
                     PUT FileNum, Cnt, OSChk
                     EXIT FOR
                  END IF
               
               CASE ELSE
                  PRINT "Outstanding"
                  PRINT
                  COLOR 10
                  PRINT "Mark this check as cleared (Y/n) ";

                  X$ = ""
                  YesNoB X$, XCode, Colr
                  IF XCode = 2 THEN EXIT FOR
                  IF X$ = "N" THEN EXIT FOR
                  
                  OSChk.Cleared = 1
                  PUT FileNum, Cnt, OSChk
                  EXIT FOR
                  
            END SELECT
            
            EXIT FOR

         END IF

   NEXT

         IF NOT Found THEN
            LOCATE 6, 1
            PRINT "Check not found.  Press any key to continue."
            wait$ = INPUT$(1)
         END IF

LOOP



GBye:
UNLOCK FileNum

ClrGetout:
VIEW PRINT
CLOSE



'GotErr:
'   ErrorNum = ERR
'   SELECT CASE ErrorNum
'      CASE 70
'         CLS
'         LOCATE 12, 1: COLOR 14
'         PRINT "Another user has this file locked."
'         PRINT "Press any key to exit"
'         K$ = INPUT$(1)
'         CLS : CLOSE : END
'       CASE ELSE
'         CLS
'         LOCATE 12, 1: COLOR 14
'         PRINT USING "Error ## occured."; ERR
'         PRINT "Press any key to exit"
'         K$ = INPUT$(1)
'         CLS : CLOSE : END
'   END SELECT

END SUB

SUB GetFile
SHARED FileNum, RecLen, NumRecs

   RecLen = LEN(OSChk)
   FileNum = FREEFILE
   OPEN "glbanktr.dat" FOR RANDOM ACCESS READ WRITE SHARED AS FileNum LEN = RecLen
   NumRecs = LOF(FileNum) \ RecLen

END SUB

SUB PrintChkList (Title$, WhatKind%)
SHARED FileNum, RecLen, NumRecs

   GetFile

   LOCK FileNum

   'Print the input screen
   COLOR 11
   LOCATE 4, 1
   PRINT "Include checks dated through (Enter=All):"
   LOCATE 6, 1
   PRINT "Include (P)ayroll, (A)ccounts Payable (Enter=All):"
   LOCATE 8, 1
   PRINT "Print Bank Number:"
   LOCATE 10, 1
   PRINT "Output to (S)creen or (P)rinter:"

   Colr = 14  'Set Data entry Color

GetDate:
   LOCATE 4, 43
   DateIn Dt$, ExitCode, Colr
   IF ExitCode = 1 THEN GOTO GetDate
   IF ExitCode = 2 THEN GOTO Getout
   Dt$ = QPTrim$(Dt$)
   IF LEN(Dt$) = 0 THEN
      SrchDate& = 999999
   ELSE
      SrchDate& = VAL(RIGHT$(Dt$, 2) + LEFT$(Dt$, 2) + MID$(Dt$, 3, 2))
   END IF

GetType:
   LOCATE 6, 52
   Type$ = "": Max = 1: CapsOn = 1: NumOnly = 0
   TextIn ChkType$, Max, NumOnly, CapsOn, ExitCode, Colr
   IF ExitCode = 1 THEN GOTO GetDate
   IF ExitCode = 2 THEN GOTO Getout
   SELECT CASE ChkType$
      CASE "P"
         ChkType = 1
         GOTO GetDev
      CASE "A"
         ChkType = 0
      CASE ELSE
         ChkType = 3
   END SELECT

GetBank:
   LOCATE 8, 20
   Type$ = "": Max = 1: CapsOn = 1: NumOnly = 1
   TextIn BankNo$, Max, NumOnly = 1, CapsOn, ExitCode, Colr
   IF ExitCode = 1 THEN GOTO GetType
   IF ExitCode = 2 THEN GOTO Getout
   BankNum = VAL(BankNo$)

GetDev:
   LOCATE 10, 34
   Max = 1: CapsOn = 1: NumOnly = 0
   TextIn DevSpec$, Max, NumOnly, CapsOn, ExitCode, Colr
   IF ExitCode = 1 THEN GOTO GetBank 'GetDate 'GetType
   IF ExitCode = 2 THEN GOTO Getout
   IF NOT (DevSpec$ = "S" OR DevSpec$ = "P") THEN
      DevSpec$ = "S"
      'PLAY "L32O3EC"
      'GOTO GetDev
   END IF

GetLPTNo:
   SELECT CASE DevSpec$
      CASE "P"
         LOCATE 10, 37
         COLOR 11
         PRINT "Printer #"

         LOCATE 10, 46
         Port$ = "1": Max = 1: NumOnly = 1
         TextIn Port$, Max, NumOnly, CapsOn, ExitCode, Colr
         IF ExitCode = 1 THEN GOTO GetDev
         IF ExitCode = 2 THEN GOTO Getout

         LPTNo = VAL(Port$)
         IF LPTNo < 1 OR LPTNo > 3 THEN
            PLAY "L32O3EC"
            GOTO GetLPTNo
         END IF

      CASE "S"
         LOCATE 10, 37
         ClearEOL 7

   END SELECT

   LOCATE 12, 1
   COLOR 10
   PRINT "Proceed with report? (Y/n) ";
   X$ = ""
   YesNoB X$, XCode, Colr
   LOCATE , 1
   ClearEOL 7
   IF XCode = 1 THEN GOTO GetDev
   IF XCode = 2 THEN GOTO Getout
   IF X$ = "N" GOTO GetDate
   'IF X$ <> "Y" GOTO GetDate

'We're ready to go

   'Report is sent to the following file which is passed to fileview for
   'screen output or printed using the BLPrint routine
   ReportFile$ = "RPT.PRN"
   
   PrnFileNum = FREEFILE
   OPEN ReportFile$ FOR OUTPUT AS #PrnFileNum

   LOCATE 12, 1
   COLOR 15
   PRINT "Processing report. Please wait."

   TotalChks# = 0  'Problem in totals using single precision
   NumChks = 0

   FOR Cnt = 1 TO NumRecs

      GET FileNum, Cnt, OSChk
      
      'SELECT CASE ChkType
      'CASE IS <= 3

        IF ChkType = 3 THEN GOTO jumpall
        IF ChkType = 1 THEN BankNum = 0

        'IF OSChk.Src = ChkType THEN
        IF OSChk.Src = ChkType AND OSChk.BankCode = BankNum THEN

jumpall:
        IF OSChk.Cleared = WhatKind THEN
            
            CheckDate& = VAL(RIGHT$(OSChk.ChkDate, 2) + LEFT$(OSChk.ChkDate, 2) + MID$(OSChk.ChkDate, 4, 2))
            IF CheckDate& <= SrchDate& THEN

               NumChks = NumChks + 1
               ChkAmt# = Round#(OSChk.Amt)
               TotalChks# = TotalChks# + ChkAmt#
               SELECT CASE OSChk.Src
                 CASE 0
                     t$ = "AP"
                 CASE 1
                     t$ = "PR"
               END SELECT

               ToPrint$ = SPACE$(80)
               LSET ToPrint$ = STR$(OSChk.ChkNum)
               MID$(ToPrint$, 10) = OSChk.ChkDate
               MID$(ToPrint$, 20) = t$
               MID$(ToPrint$, 25) = OSChk.Desc
               MID$(ToPrint$, 60) = FUsing$(STR$(OSChk.Amt), "########,.##")
               MID$(ToPrint$, 78) = STR$(OSChk.BankCode)
               PRINT #PrnFileNum, ToPrint$

            END IF

         END IF

      END IF

   NEXT

   PRINT #PrnFileNum, "" 'Add a blank line after last line

   ToPrint$ = SPACE$(80)
   LSET ToPrint$ = FUsing$(STR$(NumChks), "#####")
   MID$(ToPrint$, 8) = "Checks listed totaling: "
   MID$(ToPrint$, 60) = FUsing$(STR$(TotalChks#), "########,.##")

   PRINT #PrnFileNum, ToPrint$ + CRLF$

   UNLOCK FileNum
   CLOSE
  
   'Strings used for printing
   SELECT CASE ChkType
      CASE 0
         Kind$ = ": AP"
      CASE 1
         Kind$ = ": PR"
      CASE ELSE
         Kind$ = ": All"
   END SELECT

   CRLF$ = CHR$(13) + CHR$(10)
   FF$ = CHR$(12)
   Desc$ = "Chk No   Date      Src  Description                           Check Amt   Bank"
   DivLine$ = STRING$(79, "-")

   SELECT CASE DevSpec$
      CASE "S"
         CLS
         TitleLen = LEN(Title$ + Kind$)      'Set up screen behind fileviewer
         TitleCol = (80 - TitleLen) \ 2
         COLOR 11
         LOCATE 1, TitleCol
         PRINT Title$ + Kind$
         PRINT
         PRINT Desc$
         PRINT DivLine$
         
         LOCATE 23, 1
         COLOR 3
         PRINT "<Esc> to Exit.  F7 to Print."

         Prn2Scrn ReportFile$, 5, 21, 15, PrnFlag
         IF PrnFlag = 1 THEN GOSUB PrintIt
         
      CASE "P"
         GOSUB PrintIt:

   END SELECT


KILL ReportFile$     'Clean up after ourselves

CLOSE

EXIT SUB


PrintIt:
CLS
IF LPTNo = 0 THEN LPTNo = 1
Header$ = Title$ + Kind$ + CRLF$ + CRLF$ + Desc$ + CRLF$ + DivLine$
TitLines = 4
MaxLines = 55
Prn2Prn ReportFile$, Header$, TitLines%, MaxLines%, LPTNo%, ExitFlag%

RETURN


Getout:
UNLOCK FileNum
CLOSE
EXIT SUB

END SUB

SUB PrintTitleScrn (Title$)
   
   TitleLen = LEN(Title$)
   TitleCol = (80 - TitleLen) \ 2
   
   CLS
   LOCATE 1, TitleCol

   COLOR 11
   PRINT Title$
   
   COLOR 10
   PRINT STRING$(80, CHR$(196))

END SUB

SUB PurgeClearedChks
SHARED FileNum, RecLen, NumRecs

CLS
PrintTitleScrn "Post Cancelled Checks"

VIEW PRINT 3 TO 23

LOCATE 3, 1
COLOR 14
PRINT "This procedure removes the cleared checks from the outstanding"
PRINT "check file."
PRINT
PRINT "Press any key to continue or Esc to cancel."


K$ = INPUT$(1)

IF K$ = CHR$(27) THEN
  VIEW PRINT
  EXIT SUB
END IF

GetFile 'Get CRCHK.DAT
LOCK FileNum

'Create a new type to hold outstanding check records
DIM OSChkN AS OSChkRecType
RecLen = LEN(OSChk)
NFileNum = FREEFILE
OPEN "crchk1.dat" FOR RANDOM ACCESS READ WRITE SHARED AS NFileNum LEN = RecLen
LOCK NFileNum

   KeepRec = 0

   CLS
   FOR Cnt = 1 TO NumRecs
         
      LOCATE 4, 1
      COLOR 15
      Complete# = (Cnt / NumRecs) * 100
      PRINT USING "Processing file. ###% complete."; Complete#

      GET FileNum, Cnt, OSChk
      
      IF OSChk.Cleared = 0 THEN
      'special for algy 07/07/96     IF OSChk.ChkNum > 0 THEN

         KeepRec = KeepRec + 1

         OSChkN.ChkNum = OSChk.ChkNum
         OSChkN.ChkDate = OSChk.ChkDate
         OSChkN.Desc = OSChk.Desc
         OSChkN.Amt = OSChk.Amt
         OSChkN.Src = OSChk.Src
         OSChkN.Cleared = OSChk.Cleared
         OSChkN.BankCode = OSChk.BankCode
         OSChkN.Filler = OSChk.Filler

         PUT NFileNum, KeepRec, OSChkN

      END IF
   NEXT

UNLOCK FileNum
UNLOCK NFileNum
CLOSE

KILL "glbanktr.dat"
NAME "crchk1.dat" AS "glbanktr.dat"

VIEW PRINT

END SUB

FUNCTION Round# (N#)
    N# = INT(N# * 100 + .5) / 100
    Round# = N#
END FUNCTION

SUB SortFile
SHARED FileNum, RecLen, NumRecs
   
   GetFile
   LOCK FileNum

   REDIM FileBuff(1 TO NumRecs) AS OSChkRecType
   
   FOR Cnt = 1 TO NumRecs

      GET FileNum, Cnt, OSChk
      FileBuff(Cnt).ChkNum = OSChk.ChkNum
      FileBuff(Cnt).ChkDate = OSChk.ChkDate
      FileBuff(Cnt).Desc = OSChk.Desc
      FileBuff(Cnt).Amt = OSChk.Amt
      FileBuff(Cnt).Src = OSChk.Src
      FileBuff(Cnt).Cleared = OSChk.Cleared
      FileBuff(Cnt).BankCode = OSChk.BankCode
      FileBuff(Cnt).Filler = OSChk.Filler

   NEXT

   SortT2 FileBuff(1), NumRecs, 0, 64, 0, -2
   
   FOR Cnt = 1 TO NumRecs
      GET FileNum, Cnt, OSChk
      OSChk.ChkNum = FileBuff(Cnt).ChkNum
      OSChk.ChkDate = FileBuff(Cnt).ChkDate
      OSChk.Desc = FileBuff(Cnt).Desc
      OSChk.Amt = FileBuff(Cnt).Amt
      OSChk.Src = FileBuff(Cnt).Src
      OSChk.Cleared = FileBuff(Cnt).Cleared
      OSChk.BankCode = FileBuff(Cnt).BankCode
      OSChk.Filler = FileBuff(Cnt).Filler
      PUT FileNum, Cnt, OSChk
   NEXT
   
   UNLOCK FileNum
   CLOSE FileNum

END SUB

SUB util1

SHARED FileNum, RecLen, NumRecs

CLS
   PrintTitleScrn "Purge Checks by Date"
   VIEW PRINT 3 TO 14
   COLOR 14
   CLS 2

   PRINT "Press Esc to cancel or any other key to continue."
   K$ = INPUT$(1)

   IF K$ = CHR$(27) THEN
      VIEW PRINT
      EXIT SUB
   END IF

   'Dt$ = "03/01/96"
   PurgeDate& = 960301 'VAL(RIGHT$(Dt$, 2) + LEFT$(Dt$, 2) + MID$(Dt$, 3, 2))

   GetFile

   LOCATE 12, 1
   COLOR 11
   PRINT "Processing."

   FOR Cnt = 1 TO NumRecs

      GET FileNum, Cnt, OSChk

      CheckDate& = VAL(RIGHT$(OSChk.ChkDate, 2) + LEFT$(OSChk.ChkDate, 2) + MID$(OSChk.ChkDate, 4, 2))
      IF CheckDate& <= PurgeDate& OR OSChk.ChkNum = 0 THEN
         OSChk.Cleared = 1
         PUT FileNum, Cnt, OSChk
      END IF

'      'Changes a specific Date
'      'IF OSChk.ChkDate = Dt$ THEN
'      '   Dcnt = Dcnt + 1
'      '   OSChk.ChkDate = "12/31/79"
'      '   PUT FileNum, Cnt, OSChk
'      'END IF
'
'      IF OSChk.ChkNum < 15000 THEN
'         Ncnt = Ncnt + 1
'         OSChk.ChkNum = 1
'         PUT FileNum, Cnt, OSChk
'      END IF
'
'
'
  NEXT

'  PRINT Ncnt
'
  PRINT "Press any key to continue."
  K$ = INPUT$(1)
'
VIEW PRINT
CLOSE

'GetPDate:
'
'   CLS 2
'   COLOR 11
'   LOCATE 10, 1
'   PRINT "Enter purge date :";
'
'   LOCATE 10, 20
'   Colr = 14
'   DateIn Dt$, ExitCode, Colr
'   PurgeDate& = VAL(RIGHT$(Dt$, 2) + LEFT$(Dt$, 2) + MID$(Dt$, 3, 2))
'
'   IF ExitCode = 2 THEN
'      VIEW PRINT
'      EXIT SUB
'   END IF
'
'   LOCATE 11, 1
'   COLOR 11
'   PRINT "Ok to proceed? Last chance to cancel! (y/N) ";
'   X$ = ""
'   YesNoB X$, XCode, Colr
'   LOCATE , 1
'   ClearEOL 7
'   IF XCode = 1 THEN GOTO GetPDate
'   IF XCode = 2 THEN
'      VIEW PRINT
'      EXIT SUB
'   END IF
'   IF X$ = "N" GOTO GetPDate          'If no Start from the top
'   IF X$ <> "Y" GOTO GetPDate         'Insist on "Y" to proceed
'
'
'   SHARED OSChk AS OSChkRecType
'
'   RecLen = LEN(OSChk)
'   FileNum = FREEFILE
'   OPEN "glbanktr.dat" FOR RANDOM ACCESS READ WRITE SHARED AS FileNum LEN = RecLen
'   NumRecs = LOF(FileNum) \ RecLen
'
'   CLS
'   LOCATE 12, 1
'   COLOR 11
'   PRINT "Changing Dates."
'
'   FOR Cnt = 1 TO NumRecs
'      GET FileNum, Cnt, OSChk
'      IF OSChk.ChkDate = "02/01/95" THEN
'        OSChk.ChkDate = "02/01/96"
'        Hits = Hits + 1
'        LOCATE 13, 1
'        PRINT "Dates changed: "; Hits
'        PUT FileNum, Cnt, OSChk
'      END IF
'   NEXT
'
'  CLOSE
'
'  LOCATE 15, 1
'  PRINT "Press any key to continue."
'  K$ = INPUT$(1)

'
END SUB

SUB util2
  'STOP
SHARED FileNum, RecLen, NumRecs

   GetFile

   LOCK FileNum
   
   CLS
   Title$ = "Check Rec Utility"
   PrintTitleScrn Title$
   VIEW PRINT 3 TO 14
   COLOR 14
   Colr = 14
   CLS 2
   
UGetDev:
   LOCATE 8, 1: PRINT "Output Device:"
   LOCATE 8, 34
   Max = 1: CapsOn = 1: NumOnly = 0
   TextIn DevSpec$, Max, NumOnly, CapsOn, ExitCode, Colr
   'IF ExitCode = 1 THEN GOTO GetDate 'GetType
   IF ExitCode = 2 THEN GOTO UGetout
   IF NOT (DevSpec$ = "S" OR DevSpec$ = "P") THEN
      PLAY "L32O3EC"
      GOTO UGetDev
   END IF

UGetLPTNo:
   SELECT CASE DevSpec$
      CASE "P"
         LOCATE 8, 37
         COLOR 11
         PRINT "Printer #"

         LOCATE 8, 46
         Port$ = "1": Max = 1: NumOnly = 1
         TextIn Port$, Max, NumOnly, CapsOn, ExitCode, Colr
         IF ExitCode = 1 THEN GOTO UGetDev
         IF ExitCode = 2 THEN GOTO UGetout

         LPTNo = VAL(Port$)
         IF LPTNo < 1 OR LPTNo > 3 THEN
            PLAY "L32O3EC"
            GOTO UGetLPTNo
         END IF

      CASE "S"
         LOCATE 8, 37
         ClearEOL 7

   END SELECT

   LOCATE 10, 1
   COLOR 10
   PRINT "Proceed with report? (Y/n) ";
   X$ = ""
   YesNoB X$, XCode, Colr
   LOCATE , 1
   ClearEOL 7
   IF XCode = 1 THEN GOTO UGetDev
   IF XCode = 2 THEN GOTO UGetout
   IF X$ = "N" GOTO UGetDev
   'IF X$ <> "Y" GOTO UGetDate

'We're ready to go

   'Report is sent to the following file which is passed to fileview for
   'screen output or printed using the BLPrint routine
   ReportFile$ = "RPT.PRN"

   PrnFileNum = FREEFILE
   OPEN ReportFile$ FOR OUTPUT AS #PrnFileNum

   LOCATE 10, 1
   COLOR 15
   PRINT "Processing report. Please wait."

   TotalChks# = 0  'Problem in totals using single precision
   NumChks = 0

   FOR Cnt = 1 TO NumRecs
      GET FileNum, Cnt, OSChk

      IF LEFT$(UCASE$(OSChk.Desc), 5) = "HARLA" THEN
           OSChk.ChkNum = 1
           PUT FileNum, Cnt, OSChk
           CLOSE : END
      END IF
      'THEN STOP

      NumChks = NumChks + 1
      ChkAmt# = Round#(OSChk.Amt)
      TotalChks# = TotalChks# + ChkAmt#

      SELECT CASE OSChk.Src
        CASE 0
            t$ = "AP"
        CASE 1
            t$ = "PR"
      END SELECT

      ToPrint$ = SPACE$(80)
      LSET ToPrint$ = STR$(OSChk.ChkNum)
      MID$(ToPrint$, 10) = OSChk.ChkDate
      MID$(ToPrint$, 20) = t$
      MID$(ToPrint$, 25) = OSChk.Desc
      MID$(ToPrint$, 65) = FUsing$(STR$(OSChk.Amt), "########,.##")
      PRINT #PrnFileNum, ToPrint$

   NEXT

   PRINT #PrnFileNum, "" 'Add a blank line after last line

   ToPrint$ = SPACE$(80)
   LSET ToPrint$ = FUsing$(STR$(NumChks), "#####")
   MID$(ToPrint$, 8) = "Checks listed totaling: "
   MID$(ToPrint$, 65) = FUsing$(STR$(TotalChks#), "########,.##")

   PRINT #PrnFileNum, ToPrint$ + CRLF$

   UNLOCK FileNum
   CLOSE
   
   CRLF$ = CHR$(13) + CHR$(10)
   FF$ = CHR$(12)
   Desc$ = "Chk No   Date      Src  Description                           Check Amt Bank"
   DivLine$ = STRING$(77, "-")

   SELECT CASE DevSpec$
      CASE "S"
         VIEW PRINT
         CLS
         TitleLen = LEN(Title$ + Kind$)      'Set up screen behind fileviewer
         TitleCol = (80 - TitleLen) \ 2
         COLOR 11
         LOCATE 1, TitleCol
         PRINT Title$
         PRINT
         PRINT Desc$
         PRINT DivLine$

         LOCATE 23, 1
         COLOR 3
         PRINT "<Esc> to Exit.  F7 to Print."

         Prn2Scrn ReportFile$, 5, 21, 15, PrnFlag
         IF PrnFlag = 1 THEN GOSUB UPrintIt

      CASE "P"
         GOSUB UPrintIt:

   END SELECT


KILL ReportFile$     'Clean up after ourselves

CLOSE

EXIT SUB


UPrintIt:
CLS
IF LPTNo = 0 THEN LPTNo = 1
Header$ = Title$ + Kind$ + CRLF$ + CRLF$ + Desc$ + CRLF$ + DivLine$
TitLines = 4
MaxLines = 55
Prn2Prn ReportFile$, Header$, TitLines%, MaxLines%, LPTNo%, ExitFlag%

RETURN


UGetout:
UNLOCK FileNum
CLOSE
EXIT SUB

END SUB

SUB util3

'special for aleghany county to remove Dup pr checks from o/s check file.
'SHARED FileNum, RecLen, NumRecs
'CLS
'
'   PrintTitleScrn "Purge Checks Special"
'   VIEW PRINT 3 TO 14
'   COLOR 14
'   CLS 2
'
'   PRINT "Press Esc to cancel or any other key to continue."
'   K$ = INPUT$(1)
'  IF K$ = CHR$(27) THEN
'      VIEW PRINT
'      EXIT SUB
'   END IF
'
'   GetFile
'
'   OPEN "chklist.prn" FOR OUTPUT AS #2
'
'   LOCATE 12, 1
'   COLOR 11
'  PRINT "Processing."
'
'   FOR Cnt = 575 TO 807 STEP 2
'      GET FileNum, Cnt, OSChk
'      PRINT #2, OSChk.ChkNum; OSChk.Desc; OSChk.Amt
'      ChKAmt# = Round!(OSChk.Amt)
'      TotalChks# = TotalChks# + ChKAmt#
'
'      OSChk.ChkNum = 0
'      PUT FileNum, Cnt, OSChk
'
'   NEXT
'
'   PRINT #2, "Total Checks Removed: "; TotalChks#
'CLOSE
'
'
'VIEW PRINT
'
'   SHELL "list chklist.prn"
'


SHARED FileNum, RecLen, NumRecs
  CLS
  PRINT "Util3"
  PRINT "Press Esc to cancel or any other key to continue."
  K$ = INPUT$(1)
  IF K$ = CHR$(27) THEN
    VIEW PRINT
    EXIT SUB
  END IF

  GetFile

   FOR Cnt = 1 TO NumRecs
      GET FileNum, Cnt, OSChk

      'IF OSChk.ChkNum = 5726 THEN STOP

      SELECT CASE OSChk.ChkNum
        CASE 1300 TO 1399
          OSChk.BankCode = 3
          PUT FileNum, Cnt, OSChk
        CASE 1400 TO 1599
          OSChk.BankCode = 2
          PUT FileNum, Cnt, OSChk
        CASE IS >= 5000
          OSChk.BankCode = 1
          PUT FileNum, Cnt, OSChk

      END SELECT

   NEXT

CLOSE

END SUB

