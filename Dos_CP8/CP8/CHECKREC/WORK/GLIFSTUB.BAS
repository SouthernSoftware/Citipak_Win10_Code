COLOR 15
CLS
LOCATE 1, 1, 0
PRINT "The interface setup for other modules has not been released."
PRINT "Press any key to continue."
K$ = INPUT$(1)
RUN "GLMAINT"

