    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 22   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 5 TO 12            '--F2 HH Code list
          PressButton F2Key, 22, 5, 12
        CASE 13 TO 21           '--F3 Wrk Hist
          PressButton F3Key, 22, 13, 21
        CASE 22 TO 30           '--F4 Tran Hist
          PressButton F4KEY, 22, 22, 30
        CASE 31 TO 39           '--F6 Cons Hist
          PressButton F6KEY, 22, 31, 39
        CASE 40 TO 47           '--F7 Msgs
          IF RecNo& > 0 THEN
            PressButton F7KEY, 22, 40, 47
          END IF
        CASE 48 TO 55           '--F8 Owner
          PressButton F8KEY, 22, 48, 55
        CASE 56 TO 65           '--F10 Save Button
          PressButton F10Key, 22, 56, 65
        CASE 66 TO 75           '--ESC cancel button
          PressButton 27, 22, 66, 75
        END SELECT
      END SELECT                'row
    END IF

