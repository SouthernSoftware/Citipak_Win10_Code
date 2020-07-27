;**********************************************************************
;*                                                                    *
;* Sample command file for automatic logon to CompuServe (r) through  *
;* Telenet.                                                           *
;*                                                                    *
;**********************************************************************

EMULATE VT100           ;* emulate VT100 terminal                        *
IF NOT LINKED
   DIAL "27"            ;* Call dial dir entry number 27 (Telenet)       *
ENDIF
TRANSMIT "!"            ;* Transmit Carriage return <CR>                 *
TRANSMIT "!"            ;* Transmit Carriage return <CR>                 *
WAITFOR "TERMINAL="     ;* Wait for "TERMINAL=" to come in COM port      *
TRANSMIT "d1!"          ;* give telenet the terminal identifier          *
WAITFOR "@"             ;* Wait for "@" to come in COM port              *
TRANSMIT "C 202202!"    ;* Send request to telenet for CompuServe        *
WAITFOR "User ID:"      ;* Wait for "User ID:" to come in COM port       *
TRANSMIT "70006,51!"    ;* Send user ID <CR>                             *
WAITFOR "Password:"     ;* Wait for "Password:" to come in COM port      *
TRANSMIT "BOAT*TOUCH!"  ;* Send password <CR>                            *
ALARM                   ;* Sound alarm to let you know logon is complete *
