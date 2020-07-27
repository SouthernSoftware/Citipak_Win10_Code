DEFINT A-Z

  '$INCLUDE: 'Leave.Bi'
  '$INCLUDE: 'PRFILES.BI'

  DIM OLeave   AS oLeaveRecType
  DIM NLeave   AS LeaveRecType

  OLeaveLen = LEN(OLeave)
  NLeaveLen = LEN(NLeave)

  OPEN LeaveFileName FOR RANDOM AS #1 LEN = OLeaveLen
  GET #1, , OLeave
  CLOSE

  NLeave.VACMAX = OLeave.VACMAX
  NLeave.SICKMAX = OLeave.SICKMAX
  NLeave.HolMAX = 0
  NLeave.PerMAX = 0

  FOR zz = 1 TO 17
    LSET NLeave.VEntry(zz) = OLeave.VEntry(zz)
    LSET NLeave.SEntry(zz) = OLeave.SEntry(zz)
    NLeave.HEntry(zz).YEARS = 0
    NLeave.PEntry(zz).EARN = 0
  NEXT

  OPEN LeaveFileName FOR RANDOM AS #1 LEN = NLeaveLen
  PUT #1, , NLeave
  CLOSE

