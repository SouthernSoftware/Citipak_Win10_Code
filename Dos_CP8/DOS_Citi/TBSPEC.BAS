  DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
  DECLARE SUB PrintTitle (Title$)
  DECLARE SUB PrintLoading ()
'$INCLUDE: 'LCTax.Bi'
'$INCLUDE: 'DefCnf.BI'
'$INCLUDE: 'formedit.BI'
'$INCLUDE: 'fieldinf.BI'
''$INCLUDE: 'Qscr.BI'
'$INCLUDE: 'SetCnf.BI'





       CLOSE


      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      ReportFile$ = "TBEXC.REP"
      Header$ = "DMV Exception Listing"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      END

