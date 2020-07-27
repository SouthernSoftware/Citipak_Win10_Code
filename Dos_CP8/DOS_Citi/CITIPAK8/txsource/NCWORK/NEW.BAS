DECLARE SUB DisplayAPScrn (ScrnName$)
DECLARE SUB WPickList (Items() AS ANY, Picked%(), NPicked%, DspRow%, UpDateFlag%, Cnf AS ANY)
    MPaintBox 2, 1, 24, 80, 8   'to add shadow to list
    DisplayAPScrn "PICKOPTI"
    Title$ = SPACE$(68)
    LSET Title$ = " PO Num        TR Date       Req Num            Amount"
    QPrintRC Title$, Row - 1, 7, 112
    WPickList MTChoice(), Picked(), PickedCnt, Row, UpDateFlag, Cnf

