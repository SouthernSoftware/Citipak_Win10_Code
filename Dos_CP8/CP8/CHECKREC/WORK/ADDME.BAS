DEFINT A-Z

DEFSNG A-Z
DEFINT A-Z
SUB hold2


   'RecLen = LEN(OSChk(1))
   'FileNum = FREEFILE
   'OPEN "crchk.dat" FOR RANDOM ACCESS READ WRITE SHARED AS FileNum LEN = RecLen
   'NumRecs = LOF(FileNum) \ RecLen

Colr = 14

CLS
'PrintTitleScrn "Enter Cancelled Checks"

   COLOR 11
   LOCATE 4, 1
   PRINT "Bank Number : "

CnclGetBank:
   LOCATE 4, 15:
   BankNo$ = "": Max = 2: CapsOn = 1: NumOnly = 1
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
   Max = 8: NumOnly = 1: Colr = 14
   CheckNum$ = ""
   TextIn CheckNum$, Max, NumOnly, CapsOn, ExitCode, Colr

   IF ExitCode = 2 THEN
      VIEW PRINT
      GOTO GBye
   END IF

   ChkNum! = VAL(CheckNum$)

   Found = 0

   FOR Cnt = 1 TO NumRecs   'Do a seq search for the check number.

         GET FileNum, Cnt, OSChk(1)
         IF OSChk.Cleared = WhatKind THEN
         IF OSChk(1).ChkNum = ChkNum! AND OSChk(1).BankCode = BankNum THEN

            Found = -1

            LOCATE 5, 1
            PRINT "Check Date   : "; OSChk(1).ChkDate
            PRINT "Description  : "; OSChk(1).Desc
            PRINT "Check Amt    : "; LTRIM$(FUsing(STR$(OSChk(1).Amt), "########,.##"))
            PRINT "Status       : ";

            SELECT CASE OSChk(1).Cleared
               CASE 1
                  PRINT "Cleared    "
                  PRINT
                  COLOR 14: PRINT "Check is marked as cleared."
                  COLOR 12: PRINT "Change to outstanding (N/y) ";

                  X$ = ""
                  YesNoB X$, XCode, Colr
                  IF XCode = 2 THEN EXIT FOR 'Esc Pressed

                  IF X$ = "Y" THEN
                     OSChk(1).Cleared = 0
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

                  OSChk(1).Cleared = 1
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

