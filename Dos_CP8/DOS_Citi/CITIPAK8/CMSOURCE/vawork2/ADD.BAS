
        SELECT CASE AskSavePayment(UBSetUpRec(1).RECPDEFT)
        CASE 1
          GOSUB SaveTransaction 'do the save routine

          GOSUB PrintReceipt
          GOSUB ClearForm
          Frm(1).FldNo = 1
          Action = 1

        CASE True
          ReceiptFlag = False
          GOSUB SaveTransaction 'do the save routine
          GOSUB ClearForm
          Frm(1).FldNo = 1
          Action = 1
        CASE False
          Action = 2
        END SELECT

