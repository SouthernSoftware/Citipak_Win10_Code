ECHO OFF
CLS
IF "%1"=="" GOTO SYNTAX
set %1QLB=�
ECHO������������������������������������������������������������������������ͻ
ECHO�                ***** db/LIB's Quick Library Builder *****              �
ECHO�                                                                        �
ECHO� This batch file program is provided to assist you with the creation of �
ECHO�   new Quick Libraries for use with Microsoft BASIC PDS 7.0/7.1 or      �
ECHO�   Visual Basic for MS-DOS. It assumes all referenced OBJ and LIB files �
ECHO�   are in THIS directory.                                               �
ECHO�                                                                        �
ECHO�   Press any key to continue, or Ctrl-Break to STOP.                    �
ECHO������������������������������������������������������������������������ͼ
ECHO
pause
CLS          

IF "%VBDOSQLB%"=="�" GOTO VBDOS

:BC7
IF NOT EXIST DBLIBMF.LIB goto BC7single
ECHO
TYPE DBUSERMF.RSP|FIND "+"
ECHO
ECHO������������������������������������������������������������������������ͻ
ECHO�           These are the Libraries and Object Files that will           �
ECHO�           be combined into DBUSERMF.LIB.  If you want to add           �
ECHO�           other libraries (for example, DBCIPHRF.LIB) then             �
ECHO�           add their names to the DBUSERMF.RSP file. DBLIBMF.LIB        �
ECHO�           will be used to create DBLIBMF.QLB and DBUSERMF.QLB.         �
ECHO�                                                                        �
ECHO�           Press any key to continue, or Ctrl-Break to STOP.            �
ECHO������������������������������������������������������������������������ͼ
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
ECHO������������������������������������������������������������������������ͻ
ECHO�           These are the Libraries and Object Files that will           �
ECHO�           be combined into DBUSERSF.LIB.  If you want to add           �
ECHO�           other libraries then add their names to the                  �
ECHO�           DBUSERSF.RSP file. DBLIBSF.LIB will be used to create        �
ECHO�           DBLIBSF.QLB and DBUSERSF.QLB                                 �
ECHO�                                                                        �
ECHO�           Press any key to continue, or Ctrl-Break to STOP.            �
ECHO������������������������������������������������������������������������ͼ
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
ECHO������������������������������������������������������������������������ͻ
ECHO�                 Building Quick Libraries for VBDOS                     �
ECHO�                                                                        �
IF EXIST DBLIBMF.LIB ECHO�               Using DBLIBMF.LIB to build DBLIBMF.QLB                   �
IF EXIST DBLIBSF.LIB ECHO�               Using DBLIBSF.LIB to build DBLIBSF.QLB                   �
ECHO�                                                                        �
ECHO������������������������������������������������������������������������ͼ
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
ECHO������������������������������������������������������������������������ͻ
ECHO�  *ERROR* Link for the Quick Library failed.                            �
ECHO������������������������������������������������������������������������ͼ
goto exit

:ERROR4
ECHO
ECHO
ECHO
ECHO������������������������������������������������������������������������ͻ
ECHO�  *ERROR* Link for Building LIB file failed.                            �
ECHO������������������������������������������������������������������������ͼ
goto exit

:SYNTAX
ECHO
ECHO������������������������������������������������������������������������ͻ
ECHO�  *Syntax:    BUILDLIB  VBDOS or PDS                                    �
ECHO������������������������������������������������������������������������ͼ
goto exit

:EXIT
set %1QLB=