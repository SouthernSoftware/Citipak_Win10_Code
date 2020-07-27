ECHO OFF
CLS
IF "%1"=="" GOTO SYNTAX
set %1QLB=๛
ECHOษออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออป
ECHOบ                ***** db/LIB's Quick Library Builder *****              บ
ECHOบ                                                                        บ
ECHOบ This batch file program is provided to assist you with the creation of บ
ECHOบ   new Quick Libraries for use with Microsoft BASIC PDS 7.0/7.1 or      บ
ECHOบ   Visual Basic for MS-DOS. It assumes all referenced OBJ and LIB files บ
ECHOบ   are in THIS directory.                                               บ
ECHOบ                                                                        บ
ECHOบ   Press any key to continue, or Ctrl-Break to STOP.                    บ
ECHOศออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผ
ECHO
pause
CLS          

IF "%VBDOSQLB%"=="๛" GOTO VBDOS

:BC7
IF NOT EXIST DBLIBMF.LIB goto BC7single
ECHO
TYPE DBUSERMF.RSP|FIND "+"
ECHO
ECHOษออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออป
ECHOบ           These are the Libraries and Object Files that will           บ
ECHOบ           be combined into DBUSERMF.LIB.  If you want to add           บ
ECHOบ           other libraries (for example, DBCIPHRF.LIB) then             บ
ECHOบ           add their names to the DBUSERMF.RSP file. DBLIBMF.LIB        บ
ECHOบ           will be used to create DBLIBMF.QLB and DBUSERMF.QLB.         บ
ECHOบ                                                                        บ
ECHOบ           Press any key to continue, or Ctrl-Break to STOP.            บ
ECHOศออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผ
ECHO
PAUSE
IF EXIST DBUSERMF.LIB ERASE DBUSERMF.LIB
LIB  @DBUSERMF.RSP         
IF ERRORLEVEL 1 GOTO ERROR4
ECHO
IF EXIST DBUSERMF.LIB ECHO      DBUSERMF.LIB Linkable Library built successfully.

IF EXIST DBLIBMF.QLB ERASE DBLIBMF.QLB
IF EXIST DBUSERMF.QLB ERASE DBUSERMF.QLB
IF EXIST DBUSERMF.LIB LINK /Q DBUSERMF.LIB,DBUSERMF.QLB,NUL,QBXQLB,NUL ;>NUL
IF ERRORLEVEL 1 GOTO ERROR3
IF EXIST DBLIBMF.LIB LINK /Q DBLIBMF.LIB,DBLIBMF.QLB,NUL,QBXQLB,NUL ;   >NUL
IF ERRORLEVEL 1 GOTO ERROR3
IF EXIST DBUSERMF.QLB ECHO      DBUSERMF.QLB Quick Library built successfully.
IF EXIST DBLIBMF.QLB ECHO      DBLIBMF.QLB Quick Library built successfully.
ECHO
GOTO EXIT

:BC7single
ECHO
TYPE DBUSERSF.RSP|FIND "+"
ECHO
ECHOษออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออป
ECHOบ           These are the Libraries and Object Files that will           บ
ECHOบ           be combined into DBUSERSF.LIB.  If you want to add           บ
ECHOบ           other libraries then add their names to the                  บ
ECHOบ           DBUSERSF.RSP file. DBLIBSF.LIB will be used to create        บ
ECHOบ           DBLIBSF.QLB and DBUSERSF.QLB                                 บ
ECHOบ                                                                        บ
ECHOบ           Press any key to continue, or Ctrl-Break to STOP.            บ
ECHOศออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผ
ECHO
PAUSE
IF EXIST DBUSERSF.LIB ERASE DBUSERSF.LIB
LIB @DBUSERSF.RSP
IF ERRORLEVEL 1 GOTO ERROR4
ECHO
IF EXIST DBUSERSF.LIB ECHO      DBUSERSF.LIB Linkable Library built successfully.
IF EXIST DBLIBSF.QLB ERASE DBLIBSF.QLB
IF EXIST DBUSERSF.QLB ERASE DBUSERSF.QLB
IF EXIST DBUSERSF.LIB LINK /Q DBUSERSF.LIB,DBUSERSF.QLB,NUL,QBXQLB,NUL ; >NUL
IF ERRORLEVEL 1 GOTO ERROR3
LINK /Q DBLIBSF.LIB,DBLIBSF.QLB,NUL,QBXQLB,NUL;                          >NUL
IF ERRORLEVEL 1 GOTO ERROR3
IF EXIST DBUSERSF.QLB ECHO      DBUSERSF.QLB Quick Library built successfully.
IF EXIST DBLIBSF.QLB ECHO      DBLIBSF.QLB Quick Library built successfully.
ECHO
GOTO EXIT

:VBDOS
CLS
ECHO
ECHOษออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออป
ECHOบ                 Building Quick Libraries for VBDOS                     บ
ECHOบ                                                                        บ
IF EXIST DBLIBMF.LIB ECHOบ               Using DBLIBMF.LIB to build DBLIBMF.QLB                   บ
IF EXIST DBLIBSF.LIB ECHOบ               Using DBLIBSF.LIB to build DBLIBSF.QLB                   บ
ECHOบ                                                                        บ
ECHOศออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผ
ECHO

IF EXIST DBLIBMF.LIB LINK /Q DBLIBMF.LIB,DBLIBMF.QLB,NUL,VBDOSQLB,NUL ;  >NUL
IF ERRORLEVEL 1 GOTO ERROR3
IF EXIST DBLIBSF.LIB LINK /Q DBLIBSF.LIB,DBLIBSF.QLB,NUL,VBDOSQLB,NUL;   >NUL
IF ERRORLEVEL 1 GOTO ERROR3
IF EXIST DBLIBMF.QLB ECHO          DBLIBMF.QLB Quick Library built successfully.
IF EXIST DBLIBSF.QLB ECHO          DBLIBSF.QLB Quick Library built successfully.
ECHO
GOTO EXIT

:ERROR3
ECHO
ECHO
ECHO
ECHOษออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออป
ECHOบ  *ERROR* Link for the Quick Library failed.                            บ
ECHOศออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผ
goto exit

:ERROR4
ECHO
ECHO
ECHO
ECHOษออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออป
ECHOบ  *ERROR* Link for Building LIB file failed.                            บ
ECHOศออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผ
goto exit

:SYNTAX
ECHO
ECHOษออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออป
ECHOบ  *Syntax:    BUILDLIB  VBDOS or PDS                                    บ
ECHOศออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผ
goto exit

:EXIT
set %1QLB=
