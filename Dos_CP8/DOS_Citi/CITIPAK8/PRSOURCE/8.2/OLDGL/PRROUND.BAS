DEFINT A-Z
DECLARE FUNCTION Round# (DblNum#)

FUNCTION Round# (DblNum#)
  Round# = (INT((DblNum# * 100) + .5) / 100)
END FUNCTION

