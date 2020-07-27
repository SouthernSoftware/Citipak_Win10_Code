DEFINT A-Z
DECLARE FUNCTION uRound# (N#)

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5000001) / 100
END FUNCTION

