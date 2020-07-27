;**********************************************************************
;*                                                                    *
;* Sample command file for automatic logon to FIDO BBS systems        *
;*                                                                    *
;**********************************************************************

EMULATE ANSI                ;* Use ANSI-BBS emulation for ANSI graphics      *
IF NOT LINKED
   DIAL "10"                ;* Call dialing directory entry number 10        *
ENDIF
TRANSMIT "!"                ;* Transmit several Carriage returns waiting     *
PAUSE "1"                   ;* 1 second in between...                        *
TRANSMIT "!"                ;*                                               *
PAUSE "1"                   ;*                                               *
TRANSMIT "!"                ;*                                               *
PAUSE "1"                   ;*                                               *
TRANSMIT "!"                ;*                                               *
PAUSE "1"                   ;*                                               *
TRANSMIT "!"                ;*                                               *
WAITFOR "Your FIRST name:"  ;* Wait for first name prompt...                 *
TRANSMIT "Albert!"          ;* Send name.                                    *
WAITFOR "Your  LAST name:"  ;* Wait for last name prompt...                  *
TRANSMIT "Einstein!"        ;* Send name.                                    *
WAITFOR "right? (y,n)"      ;* wait for name verification prompt...          *
TRANSMIT "y"                ;* send YES                                      *
WAITFOR "Password:"         ;* Wait for FIDO password prompt...              *
TRANSMIT "EMC2!"            ;* Send password                                 *
ALARM                       ;* Sound alarm to let you know logon is complete *
