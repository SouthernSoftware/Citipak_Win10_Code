DEFINT A-Z
DECLARE FUNCTION Round# (N#)

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
FUNCTION Round# (N#)
  Round# = INT(N# * 100 + .5) / 100
END FUNCTION

