DECLARE SUB OpenFundFile (FundFileNum%, NumFunds%)
DECLARE SUB OpenFundIdx (FundIdxFileNum%, NumFIdxRecs%)

   LOCATE , , 0
   REDIM Form$(0, 0)
   REDIM Fld(0) AS FieldInfo
   REDIM Choice$(0 TO 3, 1)

   DIM Frm AS FormInfo

   LibName$ = "GL.QSL"
   FormName$ = "PRNSPEC"

   'Choices for Vendor Status field
   Choice$(0, 0) = "1"
   Choice$(1, 0) = "Screen"
   Choice$(2, 0) = "Prn#1"
   Choice$(3, 0) = "Prn#2"
   'Choice$(4, 0) = "Prn#3"

   LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
   PrintHelp "Print Fund Listing"

   Action = 1

   DO                                      'Allow user to edit the form

      EditForm Form$(), Fld(), Frm, Cnf, Action

      SELECT CASE Frm.KeyCode             'Check the user's key presses

         CASE F10Key                      'SaveButton
            Lookfor$ = RTRIM$(Form$(1, 0))
            IF LEN(LTRIM$(RTRIM$(Form$(1, 0)))) = 0 THEN
               Dev$ = "S"
               LPTNo = 1
            ELSE
               Dev$ = LEFT$(Form$(1, 0), 1)
               LPTNo = VAL(RIGHT$(RTRIM$(Form$(1, 0)), 1))
            END IF

         CASE ESCKey
            EXIT DEF

      END SELECT

   LOOP UNTIL Frm.KeyCode = F10Key

   LOCATE , , 0

   'Define vars used for printing
   MaxLines = 55
   FF$ = CHR$(12)
   Header$ = "Master Fund Listing"

   LibFile2Scrn "GL.QSL", "BG", MonoCode, Attribute, ErrorCode
   PrintHelp "Processing report. Please wait."

   OpenFundIdx FundIdxFileNum, NumFIdxRecs
   OpenFundFile FundFileNum, NumFunds
   PRNFile = FREEFILE
   ReportFile$ = "FUNDLIST.PRN"
   OPEN ReportFile$ FOR OUTPUT AS #PRNFile

   GOSUB PrintFundPageHeader

   FOR Cnt = 1 TO NumFIdxRecs

      GET FundIdxFileNum, Cnt, FundIdx
      GET FundFileNum, FundIdx.RecNum, Fund

      HowMany = HowMany + 1

      ToPrint$ = SPACE$(80)
      MID$(ToPrint$, 2) = Fund.FundNum
      MID$(ToPrint$, 18) = Fund.Title
      PRINT #PRNFile, ToPrint$
      LineCnt = LineCnt + 1
      IF LineCnt > MaxLines THEN
        PRINT #PRNFile, FF$
        GOSUB PrintFundPageHeader
      END IF

   NEXT
   PRINT #PRNFile,
   PRINT #PRNFile, HowMany; "Funds listed."
   PRINT #PRNFile, FF$

   CLOSE

   SELECT CASE Dev$
      CASE "S"
         PrintRptFile Header$, ReportFile$, LPTNo, RetCode, 2
      CASE "P"
         'LibFile2Scrn "GL.QSL", "BAKGRD", MonoCode, Attribute, ErrorCode
         PrintRptFile Header$, ReportFile$, LPTNo, RetCode, 4
   END SELECT

   KILL ReportFile$


EXIT DEF


PrintFundPageHeader:
  PRINT #PRNFile, Header$
  PRINT #PRNFile,
  PRINT #PRNFile, " Fund Number     Title"
  PRINT #PRNFile, STRING$(80, "-")
  LineCnt = 4
RETURN

