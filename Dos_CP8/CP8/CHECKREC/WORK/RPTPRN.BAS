'RptPrn.bas 01/15/95 by Jim Brown
DEFINT A-Z

DECLARE SUB BLPrint (LPTNo%, ToPrint$, ErrCode%)
DECLARE SUB FileView (FileName$, ky%, Action%, FVI AS ANY, SEG Array)
DECLARE SUB QPrintRC (t$, r%, c%, Clr%)
DECLARE SUB LPrn (LPTNo%, ToPrint$, CancelFlag%)
DECLARE SUB LPrnt (ReportFile$, Title$, TitLines%, MaxLines%, LPTNo%, CancelFlag%)
DECLARE SUB Prn2Scrn (FileName$, TopRow%, BotRow%, Clr%, PrnFlag%)
DECLARE FUNCTION WaitKey ()
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)

'$INCLUDE: 'fileview.bi'

SUB LPrn (LPTNo, ToPrint$, CancelFlag%)

   REDIM ScrBuf(2000)                        'set aside screen buffer

   MScrnSave 1, 1, 25, 80, ScrBuf(0)         'save screen

   DO                                        'print the string

     BLPrint LPTNo, ToPrint$ + CHR$(13) + CHR$(10), ErrCode  'Add CrLf$

     IF ErrCode <> -1 THEN                   'anything but -1 means error
        
        PLAY "L32O3EC"
        
        MScrnSave 1, 1, 25, 80, ScrBuf(0)
        
        CLS
        QPrintRC "Printer NOT Ready!", 12, 1, 12
        QPrintRC "Press Esc to Cancel.", 13, 1, 12
        QPrintRC "Any other key to continue.", 14, 1, 12
        
        K = WaitKey                          'tell user somethings wrong
        IF K = 27 THEN                       'get out if esc key
           CancelFlag = -1                   'set cancel flag
           EXIT SUB
        END IF

        MScrnRest 1, 1, 25, 80, ScrBuf(0)    'restore screen

        'if an error, ErrCode shows how many characters were printed.
        ToPrint$ = RIGHT$(ToPrint$, LEN(ToPrint$) - ErrCode)

     END IF

   LOOP UNTIL ErrCode = -1

END SUB

SUB Prn2Prn (ReportFile$, Title$, TitLines%, MaxLines%, LPTNo%, CancelFlag%)
   
   REDIM ScrBuf(2000)                         'Set aside a screen buffer

   MScrnSave 1, 1, 25, 80, ScrBuf(0)          'Save the existing screen

   CLS
   QPrintRC "Ready the printer and press any key to continue. (Esc=Cancel)", 12, 1, 11
   
   K = WaitKey                                'Tell user to ready printer
   IF K = 27 THEN                             'wait for key press
      EXIT SUB                                'esc to cancel
   END IF

   MScrnRest 1, 1, 25, 80, ScrBuf(0)          'restore screen after message
   
   CLS
   QPrintRC "Printing... Press Esc to cancel.", 12, 1, 11
   
   ReportFileNum = FREEFILE
   OPEN ReportFile$ FOR INPUT AS #ReportFileNum    'open the print file

   LPrn LPTNo, Title$, CancelFlag             'print the header
   IF CancelFlag = -1 THEN
       CLOSE
       EXIT SUB
   END IF

   LineCnt = TitLines                          'set line counter

   DO UNTIL EOF(ReportFileNum)                 'Start reading the report file

      K$ = INKEY$                              'check for key press
      IF K$ = CHR$(27) THEN                    'user pressed Esc key
         MScrnSave 1, 1, 25, 80, ScrBuf(0)

         CLS
         QPrintRC "Are you sure you want to cancel? (Y/N)", 12, 1, 12
         
         Cancel$ = UCASE$(INPUT$(1))
         IF Cancel$ = "Y" THEN                 'cancel printing
            CLOSE                              'close file
            EXIT SUB                           'get out
         ELSE
            Cancel$ = ""                       'didn't mean to stop
            MScrnRest 1, 1, 25, 80, ScrBuf(0)  'restore screen
         END IF

      END IF

      IF LineCnt >= MaxLines THEN              'check line count

         LPrn LPTNo, CHR$(12), CancelFlag      'formfeed at max line
         IF CancelFlag = -1 THEN
             CLOSE
             EXIT SUB
         END IF

         LPrn LPTNo, Title$, CancelFlag        'print header
         IF CancelFlag = -1 THEN
             CLOSE
             EXIT SUB
         END IF

         LineCnt = TitLines                    'Increase line count by header

      END IF

      LINE INPUT #1, ToPrint$                  'Get next line from file.
      LineCnt = LineCnt + 1                    'Increase line count
      LPrn LPTNo, ToPrint$, CancelFlag         'print it
      IF CancelFlag = -1 THEN
          CLOSE
          EXIT SUB
      END IF

   LOOP

   IF CancelFlag = 0 THEN    'if all printed w/o user abort send formfeed
      LPrn LPTNo, CHR$(12), CancelFlag
   END IF

   CLOSE
   
END SUB

SUB Prn2Scrn (FileName$, TopRow, BotRow, Clr, PrnFlag)

   DIM FV AS FVInfo              'Dimension the fileview array
   DIM FVBuffer(1 TO 16384)      'Set up the fileview buffer

   FV.Colr = Clr                 'Set screen display color
   FV.TRow = TopRow             'Starting Row
   FV.LCol = 1                  'Left Col
   FV.BRow = BotRow             'Bottom Row
   FV.RCol = 80                 'Right Col
   FV.TabStop = 8

   Action = 1
   DO
      FileView FileName$, ky, Action, FV, FVBuffer(1)
      IF ky = -65 THEN           'F7 Key
         FileView FileName$, ky, 5, FV, FVBuffer(1)   'Action 5 to close.
         PrnFlag = 1             'Set flag to print to printer
         EXIT SUB
      END IF
   LOOP UNTIL ky = 27

   FileView FileName$, ky, 5, FV, FVBuffer(1)   'Action 5 to close.

END SUB

