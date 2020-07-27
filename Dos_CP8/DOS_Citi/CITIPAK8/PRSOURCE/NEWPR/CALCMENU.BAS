
  DO
    PayMenu PayProcessMenu, Choice, 8
    '-=-=-=-=-=-=-=-=
    SELECT CASE Choice
    CASE 2
      IF NOT PrdDefRec(1).MACTIVE THEN
        PCSetPeriodDefault
      ELSE
        BlockClear
        DisplayMiscScrn NoNormalNow
        WaitForAction
      END IF
    CASE 1, 3 TO 9
      IF PrdDefRec(1).PACTIVE OR Choice = 1 OR Choice = 8 THEN
        SELECT CASE Choice
        CASE 1
          'AccruLeave False       'unrem
        CASE 3  'payroll transaction entry
          EntryType = Normal
          PCGetEmpNum LastEmpNum
        CASE 4
          IF SplitFlag THEN
            'PCPrintPayRegisterS     'unrem me
          ELSE
            'PCPrintPayRegister     'unrem me
          END IF
        CASE 5
          'PRCheckMenu            'unrem
        CASE 6

        CASE 7
          EntryType = Normal
          PostTransactions
        CASE 8  'manual transaction entry
          IF PrdDefRec(1).PACTIVE THEN
            'No manual transaction durning payroll active error
            BlockClear
            DisplayMiscScrn NoManualNow
            WaitForAction
          ELSE
            ManualMenu
          END IF
        'CASE 8
        '  'VoidTransaction               'unrem

        END SELECT
      ELSE
        IF NOT PrdDefRec(1).MACTIVE THEN
          BlockClear
          DisplayMiscScrn NoPeriodDefaults
          WaitForAction
        ELSE
          BlockClear
          DisplayMiscScrn NoNormalNow
          WaitForAction
        END IF
        Choice = 1
      END IF
    END SELECT
  LOOP UNTIL Choice = EscKey

