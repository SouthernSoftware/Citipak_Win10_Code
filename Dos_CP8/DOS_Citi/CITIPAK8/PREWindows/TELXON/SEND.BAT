TXRX SENDFILE.DAT

if errorlevel == 4 goto UserAbort
if errorlevel == 3 goto FileRefused
if errorlevel == 2 goto NoResponse
if errorlevel == 0 goto FileSent

 echo UNKNOWN ERROR > TELXON.RET
 goto ExitSend

:NoResponse
 echo NO RESPONSE  > TELXON.RET
 goto ExitSend

:FileRefused
 echo FILE REFUSED > TELXON.RET
 goto ExitSend

:UserAbort
 echo USER ABORTED > TELXON.RET
 goto ExitSend

:FileSent
 echo FILE SENT > TELXON.RET

:ExitSend
