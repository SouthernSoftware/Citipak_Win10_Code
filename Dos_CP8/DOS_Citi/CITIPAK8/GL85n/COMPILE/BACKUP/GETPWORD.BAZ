DECLARE FUNCTION GetPOPassWord% ()
  DEFINT A-Z
  
FUNCTION GetPOPassWord
  
  
  DO
    
    BlockClear
    WazzWind 8, 15, 16, 65, 14, 1, 1
    
    QPrintRC "[ VIEW LOG FILE PASSWORD ]", 8, 18, 14
    QPrintRC "Enter password:", 12, 22, 15
    
    IF Try > 0 THEN
      QPrintRC "       Attempt:" + STR$(Try), 14, 22, 15
    END IF
    
    PW$ = nGetPassword(12, 38, 10, Abort)
    
    IF Abort THEN
      UBLog "PASSWORD ABORTED: " + PW$
      EXIT DO
    END IF
    
    IF PW$ = "REVIEWLOG" THEN
      UBLog "PASSWORD OK!"
      OkFlag = True
    ELSE
      UBLog "INVALID PASSWORD: " + PW$
      QPrintRC "Invalid Password. Press any key to try again.", 14, 18, 31
      WaitForAction
    END IF
    
    Try = Try + 1
    
    IF Try = 4 THEN
      BlockClear
      WazzWind 8, 15, 16, 65, 15, 1, 1
      QPrintRC "Get the password from the system operator.", 11, 20, 14
      QPrintRC "Press any key to return to the Main Menu.", 13, 20, 14
      WaitForAction
      EXIT DO
    END IF
    
  LOOP UNTIL OkFlag
  
  'VGetOut:
  IF OkFlag THEN
    GetPOPassWord = True
  ELSE
    GetPOPassWord = False
  END IF
  
END FUNCTION
  
