   IF Ky$ = CHR$(27) THEN
     REDIM ARCatCodeRec(1) AS ARCatCodeRecType
   ELSE
     GET ARCatFile, Choice, ARCatCodeRec(1)
   END IF

