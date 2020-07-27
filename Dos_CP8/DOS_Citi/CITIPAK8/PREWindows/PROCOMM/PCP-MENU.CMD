;*********************************************************************
;
; ProComm 2.4 Command file to dial PC Pursuit numbers.  Demonstrates
; menu building, branching on user input, etc.  Please note that only
; three cities are implemented in this demo.
;
; String variables are used extensively here.  Their uses are:
;
; S0 - holds user input for SWITCHing
; S1 - holds PCP city name
; S2 - holds PCP city area code
; S3 - holds final destination name
; S4 - holds final destination phone number
; S5 - holds user input for SWITCHing
; S6 - you must set to modem dial command and local TELENET number
; S7 - you must set to apporpriate baud rate (3 or 12)
; S8 - you must set to your PCP user id
; S9 - you must set to your PCP password
;
;*********************************************************************



; Customization area - you need to change these for your setup


   ASSIGN S9 "12"               ; Enter the baudrate to use here, 3 or 12
   ASSIGN S8 "xxxxxxx"          ; Enter the id to be used here
   ASSIGN S7 "xxxxxxxx"         ; Enter the password to be used here
   ASSIGN S6 "ATDTxxx-xxxx!"    ; Enter dial cmd and local telenet # here


   SET CR_IN CR                 ; make sure no extra LF's in menus
   CLEAR                        ; clear the screen

   LOCATE 1,0                   ; position the cursor

; Build the first menu

   MESSAGE "                       Ú´ PC PURSUIT AVAILABLE CITIES Ã¿"
   MESSAGE "                       ³                               ³"
   MESSAGE "                       ³       1- ATLANTA              ³"
   MESSAGE "                       ³       2- BOSTON               ³"
   MESSAGE "                       ³       3- DALLAS               ³"
   MESSAGE "                       ³       4- DENVER               ³"
   MESSAGE "                       ³       5- DETROIT              ³"
   MESSAGE "                       ³       6- HOUSTON              ³"
   MESSAGE "                       ³       7- LOS ANGELES          ³"
   MESSAGE "                       ³       8- NEWARK               ³"
   MESSAGE "                       ³       9- NEW YORK             ³"
   MESSAGE "                       ³      10- PHILADELPHIA         ³"
   MESSAGE "                       ³      11- SAN FRANCISCO        ³"
   MESSAGE "                       ³      12- SEATTLE              ³"
   MESSAGE "                       ³      13- WASHINGTON, DC       ³"
   MESSAGE "                       ÃÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´"
   MESSAGE "                       ³ ENTER THE # OF THE CITY OF    ³"
   MESSAGE "                       ³ YOUR CHOICE:                  ³"
   MESSAGE "                       ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ"

L1:                             ; label for use by default case

   LOCATE 18,38                 ; position cursor
   MESSAGE "  "                 ; clear prompt area
   LOCATE 18,38                 ; reposition cursor
   GET S0 2                     ; get users choice

   SWITCH S0                    ; switch on choice

      CASE "1"                  ; chose ATLANTA
         ASSIGN S1 "ATLANTA"    ; save the city name
         ASSIGN S2 "404"        ; save the area code
         GOSUB ATLANTA          ; go get number to dial
      ENDCASE

      CASE "5"                  ; chose DETROIT
         ASSIGN S1 "DETROIT"    ; save the city name
         ASSIGN S2 "313"        ; save the area code
         GOSUB DETROIT          ; go get number to dial
      ENDCASE

      CASE "13"                       ; chose WASHINGTON DC
         ASSIGN S1 "WASHINGTON, DC"   ; save the city name
         ASSIGN S2 "202"              ; save the area code
         GOSUB WASHINGTON_DC          ; go get number to dial
      ENDCASE

      DEFAULT                   ; all other cases
         LOCATE 22,24
         MESSAGE "^G"           ; error noise
         LOCATE 22,24
         MESSAGE "Not yet implemented in this demo"
         LOCATE 22,24
         PAUSE 2
         MESSAGE "                                "
         GOTO L1
      ENDCASE
   ENDSWITCH

;
; Ok, call the local TELENET number first.  This is an example of
; redialing without using the DIAL or Alt-R commands
;

   CLEAR                        ; ready to call TELENET
   SET PARITY EVEN              ; set parms for TELENET
   SET DATABITS 7
   MESSAGE "DIALING PCP"        ; let em know whats happening
   LOCATE 0,12
   MESSAGE S1                   ; include city in message

L2:                             ; label for busy TELENET

   TRANSMIT S6                  ; call local Telenet
   WAITFOR "CONNECT" 15         ; give it 15 seconds to answer

   IF NOT WAITFOR               ; it didn't connect, assume busy
      CLEAR                     ; example of redialing without Alt-R
      HANGUP
      MESSAGE "PCP TELENET IS BUSY.  REDIALING..."
      GOTO L2
   ENDIF

;
; Ok, made it to local telenet number, now call the remote PCP city
; Notice that we redial at this point as well
;

   PAUSE 2
   TRANSMIT "^m^m"              ; Get Telenet's attention
   WAITFOR "="                  ; Wait for 'TERMINAL=' prompt
   TRANSMIT "D1!"               ; Enter terminal type
   WAITFOR "@"                  ; Wait for Telenet prompt

L3:                             ; label for busy remote PCP city
   TRANSMIT "C DIAL"
   TRANSMIT S2                  ; area code
   TRANSMIT "/"
   TRANSMIT S9                  ; BAUD RATE
   TRANSMIT ","
   TRANSMIT S8                  ; ID
   TRANSMIT "!"                 ; cr

   WAITFOR "="                  ; Wait for 'PASSWORD=' prompt

   TRANSMIT S7                  ; Enter password
   TRANSMIT "!"                 ; cr

   RGET S5                      ; Get the CR echoed from our TRANSMIT
   RGET S5                      ; Get the CR that precedes PCP messages
   RGET S5                      ; Get the actual message
   PAUSE 1                      ; Display it briefly

   FIND S5 "CONNECT"            ; Look for connection

   IF NOT FOUND                 ; No connection
      CLEAR
      MESSAGE "THE REMOTE PCP IS BUSY.  REDIALING..." S1
      LOCATE 1 0
      GOTO L3
   ENDIF

;
; Ok, made it to the remote PCP city.  Display appropriate message
; then call the actual destination, again using a redial
;

   CLEAR
   MESSAGE "CONNECTED TO PCP"
   LOCATE 0,17
   MESSAGE S1
   LOCATE 1,0
   MESSAGE "DIALING"
   LOCATE 1,8
   MESSAGE S3

L4:                             ; label for busy number

   LOCATE 2 0
   TRANSMIT "ATZ!"              ; Wake up remote modem
   WAITFOR "OK"                 ; Wait for 'OK' prompt

   TRANSMIT "ATDT"
   TRANSMIT S4                  ; THE NUMBER
   TRANSMIT "!"                 ; CR

   RGET S0                      ; Get CR echoed from our TRANSMIT
   RGET S0 80 40                ; Get CR that precedes PCP messages
   RGET S0                      ; Get the real message
   PAUSE 2

   FIND S0 "BUSY"               ; was it busy?
   IF FOUND
      CLEAR
      MESSAGE "BUSY.  REDIALING..." S3
      GOTO L4
   ENDIF

   FIND S0 "CONNECT"            ; did it connect?
   IF FOUND
      SET PARITY NONE           ; Switch to N/8/1 for BBS work
      SET DATABITS 8
      CLEAR
      MESSAGE "CONNECTED TO"
      LOCATE 0,13
      MESSAGE S3
      LOCATE 1,0
      MESSAGE "LOGON PRECEEDING"
   ELSE                         ; not busy or connect, assume error
      CLEAR
      MESSAGE "NO CONNECT OR BUSY. REDIALING"
      GOTO L4
   ENDIF

   ALARM 2                      ; sound alarm for 2 seconds
   EXIT                         ; we made it!



;********************************************************************
;
; SUBROUTINES START HERE
;
; note that you could put these at the beginning of the file, and
; initially branch around them.  As each GOSUB and GOTO starts looking
; at the top of the file, it would speed up some parts of the process,
; while slowing down others.  Maybe next time we'll build a symbol table
; and eliminate this inefficiency.
;
;********************************************************************


ATLANTA:

   CLEAR
   LOCATE 5,0
   MESSAGE "                         ÚÄÄ´ ATLANTA BBS NUMBERS ÃÄÄ¿"
   MESSAGE "                         ³                           ³"
   MESSAGE "                         ³   1- ALTGATE              ³"
   MESSAGE "                         ³   2- Flagship Express     ³"
   MESSAGE "                         ³   3- BYTE Connection      ³"
   MESSAGE "                         ³   4- Microstuf BBS        ³"
   MESSAGE "                         ³   5- Software Link BBS    ³"
   MESSAGE "                         ³   6- PC-Exchange BBS      ³"
   MESSAGE "                         ³   M- Manual Dial          ³"
   MESSAGE "                         ³                           ³"
   MESSAGE "                         ÃÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´"
   MESSAGE "                         ³ ENTER THE # OF THE BBS OF ³"
   MESSAGE "                         ³ YOUR CHOICE:              ³"
   MESSAGE "                         ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ"

ATL1:

   LOCATE 17,40
   MESSAGE  " "
   LOCATE 17,40
   GET S5 1

   SWITCH S5

      CASE "1"
         ASSIGN S3 "ATLGATE"            ; save the BBS name
         ASSIGN S4 "9281876"            ; save the number
      ENDCASE

      CASE "2"
         ASSIGN S3 "Flagship Express"
         ASSIGN S4 "9344515"
      ENDCASE

      CASE "3"
         ASSIGN S3 "BYTE Connection"
         ASSIGN S4 "9410159"
      ENDCASE

      CASE "4"
         ASSIGN S3 "Microstuf_BBS"
         ASSIGN S4 "9988048"
      ENDCASE

      CASE "5"
         ASSIGN S3 "Software Link"
         ASSIGN S4 "9986737"
      ENDCASE

      CASE "6"
         ASSIGN S3 "PC-Exchange BBS"
         ASSIGN S4 "9776686"
      ENDCASE

      CASE "M"
         CLEAR
         LOCATE 5,20
         MESSAGE "NUMBER TO DIAL:"
         LOCATE 5,36
         ASSIGN S3 "MANUAL DIAL"
         GET S4
      ENDCASE

      DEFAULT
         MESSAGE "^G"
         LOCATE 17,40
         GOTO ATL1
      ENDCASE

   ENDSWITCH
RETURN

DETROIT:

   CLEAR
   LOCATE 5,0
   MESSAGE "                         ÚÄÄ´ DETROIT BBS NUMBERS ÃÄÄ¿"
   MESSAGE "                         ³                           ³"
   MESSAGE "                         ³   1- Computer Support BBS ³"
   MESSAGE "                         ³   2- The Trading Post     ³"
   MESSAGE "                         ³   3-                      ³"
   MESSAGE "                         ³   4-                      ³"
   MESSAGE "                         ³   5-                      ³"
   MESSAGE "                         ³   6-                      ³"
   MESSAGE "                         ³   M- Manual Dial          ³"
   MESSAGE "                         ³                           ³"
   MESSAGE "                         ÃÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´"
   MESSAGE "                         ³ ENTER THE # OF THE BBS OF ³"
   MESSAGE "                         ³ YOUR CHOICE:              ³"
   MESSAGE "                         ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ"

DET1:

   LOCATE 17,40
   MESSAGE  " "
   LOCATE 17,40
   GET S5 1

   SWITCH S5

      CASE "1"
         ASSIGN S3 "Computer Support BBS"
         ASSIGN S4 "9464821"
      ENDCASE

      CASE "2"
         ASSIGN S3 "The Trading Post"
         ASSIGN S4 "8827104"
      ENDCASE

      CASE "M"
         CLEAR
         LOCATE 5,20
         MESSAGE "NUMBER TO DIAL:"
         LOCATE 5,36
         ASSIGN S3 "MANUAL DIAL"
         GET S4
      ENDCASE

      DEFAULT
         MESSAGE "^G"
         LOCATE 17,40
         GOTO DET1
      ENDCASE

   ENDSWITCH
RETURN


WASHINGTON_DC:

   CLEAR
   LOCATE 5,0
   MESSAGE "                         ÚÄÄ´ WASH DC BBS NUMBERS ÃÄÄ¿"
   MESSAGE "                         ³                           ³"
   MESSAGE "                         ³   1- PCPursuit BBS        ³"
   MESSAGE "                         ³   2- In Pursuit of ...    ³"
   MESSAGE "                         ³   3- Fannie Mae           ³"
   MESSAGE "                         ³   4- Hotline BBS          ³"
   MESSAGE "                         ³   5- Arquimedes           ³"
   MESSAGE "                         ³   6- Compu Center         ³"
   MESSAGE "                         ³   M- Manual Dial          ³"
   MESSAGE "                         ³                           ³"
   MESSAGE "                         ÃÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´"
   MESSAGE "                         ³ ENTER THE # OF THE BBS OF ³"
   MESSAGE "                         ³ YOUR CHOICE:              ³"
   MESSAGE "                         ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ"

WAS1:

   LOCATE 17,40
   MESSAGE  " "
   LOCATE 17,40
   GET S5 1

   SWITCH S5

      CASE "1"
         ASSIGN S3 "PCPursuit BBS"
         ASSIGN S4 "6893561"
         RETURN
      ENDCASE

      CASE "2"
         ASSIGN S3 "In Pursuit of..."
         ASSIGN S4 "6892987"
      ENDCASE

      CASE "3"
         ASSIGN S3 "Fannie Mae"
         ASSIGN S4 "5377475"
      ENDCASE

      CASE "4"
         ASSIGN S3 "Hotline BBS"
         ASSIGN S4 "6833944"
      ENDCASE

      CASE "5"
         ASSIGN S3 "Arquimedes"
         ASSIGN S4 "6988230"
      ENDCASE

      CASE "6"
         ASSIGN S3 "Compu Center"
         ASSIGN S4 "9799250"
      ENDCASE

      CASE "M"
         CLEAR
         LOCATE 5,20
         MESSAGE "NUMBER TO DIAL:"
         LOCATE 5,36
         ASSIGN S3 "MANUAL DIAL"
         GET S4
      ENDCASE

      DEFAULT
         MESSAGE "^G"
         GOTO WAS1
      ENDCASE

   ENDSWITCH

RETURN

;************* End of Subroutines **********************************
