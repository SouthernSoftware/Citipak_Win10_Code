ECHO OFF
CLS
IF "%1"=="" GOTO SYNTAX
set %1QLB=๛
ECHOษออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออป
ECHOบ           ***** db/LIB's Error Diagnosis Library Builder *****         บ
ECHOบ                                                                        บ
ECHOบ This batch file program is provided to assist you with the creation of บ
ECHOบ   Error Diagnosis Libraries for use with Microsoft BASIC PDS 7.0/7.1   บ
ECHOบ   or Visual Basic for MS-DOS. It assumes all referenced OBJ and LIB    บ
ECHOบ   files are in THIS directory.                                         บ
ECHOบ                                                                        บ
ECHOบ DBLIBSF.LIB and DBLIBEDM will combine into DBLIBSFE.QLB & DBLIBSFE.LIB บ
IF EXIST DBLIBMF.LIB ECHOบ DBLIBMF.LIB and DBLIBEDM will combine into DBLIBMFE.QLB & DBLIBMFE.LIB บ
ECHOบ                                                                        บ
ECHOบ   Press any key to continue, or Ctrl-Break to STOP.                    บ
ECHOศออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผ
ECHO
PAUSE
IF EXIST DBLIBMFE.QLB ERASE DBLIBMFE.QLB
IF EXIST DBLIBSFE.QLB ERASE DBLIBSFE.QLB
IF EXIST DBLIBMFE.LIB ERASE DBLIBMFE.LIB
IF EXIST DBLIBSFE.LIB ERASE DBLIBSFE.LIB
IF "%VBDOSQLB%"=="๛" GOTO VBDOS

:BC7
IF EXIST DBLIBMF.LIB LINK /Q DBLIBMF.LIB + DBLIBEDM.LIB,DBLIBMFE.QLB,NUL,QBXQLB,NUL ;
IF ERRORLEVEL 1 GOTO ERROR3
LINK /Q DBLIBSF.LIB + DBLIBEDM.LIB,DBLIBSFE.QLB,NUL,QBXQLB,NUL ;
IF ERRORLEVEL 1 GOTO ERROR3

IF EXIST DBLIBMF.LIB LIB DBLIBMFE.LIB + DBLIBMF.LIB + DBLIBEDM.LIB;
IF ERRORLEVEL 1 GOTO ERROR4
LIB DBLIBSFE.LIB + DBLIBSF.LIB + DBLIBEDM.LIB;
IF ERRORLEVEL 1 GOTO ERROR4
GOTO LIBS

:VBDOS
IF EXIST DBLIBMF.LIB LINK /Q DBLIBMF.LIB + DBLIBEDM.OBJ,DBLIBMFE.QLB,NUL,VBDOSQLB,NUL ;
IF ERRORLEVEL 1 GOTO ERROR3
LINK /Q DBLIBSF.LIB + DBLIBEDM.OBJ,DBLIBSFE.QLB,NUL,VBDOSQLB,NUL  /Q;
IF ERRORLEVEL 1 GOTO ERROR3

IF EXIST DBLIBMF.LIB LIB DBLIBMFE.LIB + DBLIBMF.LIB + DBLIBEDM.OBJ;
IF ERRORLEVEL 1 GOTO ERROR4
LIB DBLIBSFE.LIB + DBLIBSF.LIB + DBLIBEDM.OBJ;
IF ERRORLEVEL 1 GOTO ERROR4

:LIBS
IF EXIST DBLIBMFE.QLB ECHO      DBLIBMFE.QLB Quick Library built successfully.
IF EXIST DBLIBSFE.QLB ECHO      DBLIBSFE.QLB Quick Library built successfully.
IF EXIST DBLIBMFE.LIB ECHO      DBLIBMFE.LIB Linkable Library built successfully.
IF EXIST DBLIBSFE.LIB ECHO      DBLIBSFE.LIB Linkable Library built successfully.

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
ECHOบ  Syntax:    MAKE-EDM  VBDOS or PDS                                     บ
ECHOศออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผ
goto exit

:EXIT
SET %1QLB=
