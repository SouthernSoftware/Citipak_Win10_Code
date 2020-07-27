' RUpdate.BAS - Update files from Remote Site over Modem LAN Connection.
' Remote Update Facility
'
' Overview:
'   1-Open f:\updated.fil for input.
'   2-Read each line of this file. (Each line is a separate filename.
'   3-Copy this filename to C:\Citipak
'
'  ** This program should only be run at Betsy's office.  **

ON ERROR GOTO RUpdateError

60 CLS
100 GOSUB 63000: REM get town in town$
110 SCREN$ = "RUpdate": SYSDATE$ = DATE$
120 CLS : GOSUB 63500
130 LOCATE 5, 1: COLOR 11
200 PRINT TAB(32); "CitiPak Software"
210 PRINT TAB(29); "Remote Update Facility"
220 PRINT : PRINT

OPEN "f:\updated.fil" FOR INPUT AS #1
UpdatedFiles = 0
WHILE NOT EOF(1)
    LINE INPUT #1, a$
    filename$ = LTRIM$(RTRIM$(a$))
    IF LEN(filename$) > 0 THEN
        SHELL "xcopy " + filename$ + " c:\citipak"
        UpdatedFiles = UpdatedFiles + 1
    END IF
WEND
IF UpdatedFiles = 0 THEN
    PRINT "No files to update, you are now ready to work."
ELSE
    PRINT USING "Completed updating ### files. You are now ready to work."; UpdatedFiles
END IF
CLOSE
SHELL "copy f:\updated.fil f:\lastupdt.fil /y"
SHELL "del f:\updated.fil"
RUN "tbmenu"


END

63000 OPEN "R", 3, "wbtown.dat": FIELD 3, 40 AS TOWNNAME$: GET 3, 1: TOWN$ = TOWNNAME$: CLOSE 3: RETURN
63500 COLOR 10: LOCATE 1, 1: PRINT "Sys. Date: "; SYSDATE$; TAB(70); SCREN$: PRINT STRING$(79, "_"): RETURN


RUpdateError:
    SELECT CASE ERR
        CASE 53:    'File not found - No filelist for updates
            OPEN "f:\updated.fil" FOR OUTPUT AS #1
            CLOSE #1
            RESUME
        CASE ELSE:
            PRINT USING "Error ### has occured in line: ###,###,###"; ERR; ERL
    END SELECT
    RESUME
END

