   CLS
   OPEN "I", 10, "lurg.txt"
10 LINE INPUT #10, a$
   PRINT "SSN 1  : "; MID$(a$, 5, 12)
   PRINT "SSN 2  : "; MID$(a$, 137, 12)
   PRINT "L Name : "; MID$(a$, 17, 45)
   PRINT "F Name : "; MID$(a$, 62, 33)
   PRINT "M Name : "; MID$(a$, 95, 31)
   PRINT "Type   : "; MID$(a$, 126, 1)
   PRINT "Ownship: "; MID$(a$, 127, 10)
   PRINT "Add1   : "; MID$(a$, 259, 35)
   PRINT "Add2   : "; MID$(a$, 294, 35)
   PRINT "City   : "; MID$(a$, 329, 22)
   PRINT "State  : "; MID$(a$, 351, 2)
   PRINT "Zip    : "; MID$(a$, 353, 9)
   PRINT "VIN    : "; MID$(a$, 1246, 22)
   PRINT "VehMake: "; MID$(a$, 1268, 20)
   PRINT "VehModl: "; MID$(a$, 1682, 15)
   PRINT "Veh Yr : "; MID$(a$, 1288, 4)
   PRINT "Veh Wt : "; MID$(a$, 1350, 5)
   PRINT "Trade  : "; MID$(a$, 1758, 6)


   
   SLEEP 1
   CLS
   GOTO 10

   'PRINT a$
   STOP

