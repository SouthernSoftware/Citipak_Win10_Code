;**********************************************************************
;*                                                                    *
;* Sample command file for automatic logon to the Source (r) through  *
;* Telenet.                                                           *
;*                                                                    *
;**********************************************************************

EMULATE VT100           ;* emulate VT100                                 *
IF NOT LINKED
   DIAL "27"            ;* Call dial dir entry number 27 (Telenet)       *
ENDIF
TRANSMIT "!"            ;* Transmit Carriage return <CR>                 *
TRANSMIT "!"            ;* Transmit Carriage return <CR>                 *
WAITFOR "TERMINAL=" 40  ;* Wait for "TERMINAL=" to come in COM port      *
TRANSMIT "d1!"          ;* give telenet the terminal identifier          *
WAITFOR "@" 40          ;* Wait for "@" to come in COM port              *
TRANSMIT "C 30148!"     ;* Send request to telenet for The Source        *
WAITFOR ">" 40          ;* Wait for prompt from The Source               *
TRANSMIT "ID BBX714!"   ;* Transmit user id <CR>                         *
PAUSE "20"              ;* wait... The Source takes a while at first...  *
WAITFOR "Password?" 40  ;* Wait for Source password prompt               *
TRANSMIT "BOAT77!"      ;* Transmit password <CR>                        *
ALARM                   ;* Sound alarm to let you know logon is complete *
