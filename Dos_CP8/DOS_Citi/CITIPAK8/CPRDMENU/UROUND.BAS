DEFINT A-Z
DECLARE FUNCTION Round# (N#)

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
FUNCTION Round# (N#)
  Round# = INT(N# * 100 + .50000001) / 100
END FUNCTION

