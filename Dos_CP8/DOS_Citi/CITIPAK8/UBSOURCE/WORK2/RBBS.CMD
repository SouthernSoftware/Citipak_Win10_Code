;**********************************************************************
;*                                                                    *
;* Sample command file for automatic logon to RBBS systems            *
;*                                                                    *
;**********************************************************************

EMULATE ANSI                ;* Use ANSI-BBS emulation for ANSI graphics      *
IF NOT LINKED
   DIAL "15"                ;* Call dialing directory entry number 15        *
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
WAITFOR "YOUR FIRST NAME?"  ;* Wait for first name prompt...                 *
TRANSMIT "ALBERT!"          ;* Send name.                                    *
WAITFOR "YOUR LAST NAME?"   ;* Wait for last name prompt...                  *
TRANSMIT "EINSTEIN!"        ;* Send name.                                    *
WAITFOR "(dots echo)?"      ;* Wait for RBBS password prompt...              *
TRANSMIT "EMC2!"            ;* Send password                                 *
ALARM                       ;* Sound alarm to let you know logon is complete *
