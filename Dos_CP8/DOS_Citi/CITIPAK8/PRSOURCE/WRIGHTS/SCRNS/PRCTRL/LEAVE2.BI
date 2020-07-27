TYPE LeaveEntryType
    YEARS   AS INTEGER
    EARN    AS DOUBLE
END TYPE

TYPE LeaveRecType2
    VACMAX   AS DOUBLE
    VEntry(1 TO 20)  AS LeaveEntryType
    SICKMAX  AS DOUBLE
    SEntry(1 TO 20)  AS LeaveEntryType
END TYPE
