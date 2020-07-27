TYPE LeaveEntryType
    YEARS   AS INTEGER
    EARN    AS DOUBLE
END TYPE

TYPE LeaveRecType
    VACMAX   AS DOUBLE
    VEntry(1 TO 20)  AS LeaveEntryType
    SICKMAX  AS DOUBLE
    SEntry(1 TO 20)  AS LeaveEntryType
END TYPE

TYPE oLeaveEntryType
    YEARS1   AS INTEGER     '1
    YEARS2   AS INTEGER     '2
    EARN     AS DOUBLE      'earned
END TYPE

TYPE oLeaveRecType
    VACMAX   AS DOUBLE
    VEntry(1 TO 6)  AS oLeaveEntryType
    SICKMAX  AS DOUBLE
    VEntry(1 TO 6)  AS oLeaveEntryType
END TYPE

