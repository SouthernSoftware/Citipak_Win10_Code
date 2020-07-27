SUB ShowProcessingScrn (RptTitle$)
  
  TitleRow = 9
  TitleCol = 40 - (LEN(RptTitle$) \ 2) + 1
  
  CursorOff
  BlockClear
  DisplayUBScrn "PRORPT"
  
  HideCursor
  QPrintRC RptTitle$, TitleRow, TitleCol, 126
  QPrintRC "Processing:    % Completed.", 13, 28, Cnf.HiLite
  ShowCursor
END SUB
