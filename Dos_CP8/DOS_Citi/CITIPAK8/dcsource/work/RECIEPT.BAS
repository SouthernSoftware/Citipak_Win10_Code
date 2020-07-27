DECLARE FUNCTION DCAskSavePayment% (DefaultFlag$)
DEFINT A-Z


'        SELECT CASE DCAskSavePayment(UBSetUpRec(1).RECPDEFT)
'        CASE 1 'save and print
'          SaveScrn TempScrn()
'          GOSUB SaveTransaction 'do the save routine
'          IF CaldFlag THEN
'            GOSUB CWPrintReceipt
'          ELSE
'            GOSUB PrintReceipt
'          END IF
'          RestScrn TempScrn()
'          DisplayUBScrn "PRESSKEY"
'          WaitForAction
'          RestScrn TempScrn()
'          GOSUB ClearForm
'          GOSUB ClearCity
'          Frm(1).FldNo = 1
'          Action = 1
'        CASE True
'          ReceiptFlag = False
'          SaveScrn TempScrn()
'          GOSUB SaveTransaction 'do the save routine
'          DisplayUBScrn "PRESSKEY"
'          WaitForAction
'          RestScrn TempScrn()
'          GOSUB ClearForm
'          GOSUB ClearCity
'          Frm(1).FldNo = 1
'          Action = 1
'        CASE False
'          Action = 2
'        END SELECT
'



