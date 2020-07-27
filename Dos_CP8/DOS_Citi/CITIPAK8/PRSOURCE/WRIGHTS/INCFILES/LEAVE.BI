TYPE LeaveEntryType
    YEARS   AS INTEGER
    EARN    AS DOUBLE
END TYPE

TYPE oLeaveRecType
    VACMAX   AS DOUBLE
    VEntry(1 TO 20)  AS LeaveEntryType
    SICKMAX  AS DOUBLE
    SEntry(1 TO 20)  AS LeaveEntryType
END TYPE

TYPE LeaveRecType
    VACMAX   AS DOUBLE
    VEntry(1 TO 17)  AS LeaveEntryType
    SICKMAX  AS DOUBLE
    SEntry(1 TO 17)  AS LeaveEntryType
    HolMAX  AS DOUBLE
    HEntry(1 TO 17)  AS LeaveEntryType
    PerMAX  AS DOUBLE
    PEntry(1 TO 17)  AS LeaveEntryType
END TYPE

