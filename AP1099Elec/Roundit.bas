'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
Public Function Round#(ByVal N#)
  Round# = (Int(N# * 100 + 0.5000001)) / 100
End Function