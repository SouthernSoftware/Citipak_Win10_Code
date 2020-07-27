DEFINT A-Z
OPEN "eastmore.csv" FOR INPUT AS #1
DO WHILE NOT EOF(1)
 INPUT #1, PropertyAddress$, OwnerLName$, OwnerFName$, MailingAddress$, MailCity$, MailState$, MailZip$
 PRINT PropertyAddress$
 PRINT OwnerLName$
 PRINT OwnerFName$
 PRINT MailingAddress$
 PRINT MailCity$
 PRINT MailState$
 PRINT MailZip$
 'DO
 '  a$ = INKEY$
 'LOOP WHILE LEN(a$) = 0

LOOP
CLOSE

