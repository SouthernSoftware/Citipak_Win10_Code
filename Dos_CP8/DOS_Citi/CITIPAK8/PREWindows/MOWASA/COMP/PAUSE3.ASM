;Syntax: Fudge& = FudgeFactor&
;        CALL Pause3(Milliseconds%, Fudge&)
;
;Where FudgeFactor was called earlier in the program, and Milliseconds%
;ranges from 0 to 32767.


.Model Medium,BASIC
.Data
    Extrn LoWord:Word, HiWord:Word

.Code

Pause3 Proc Uses SI, MSecs:Ptr, Fudge:Ptr

    Mov  SI,MSecs         ;get the address for Milliseconds
    Mov  CX,[SI]          ;put it into CX for the Loop below
    Jcxz Exit             ;if they want zero, give 'em zero!

    Mov  SI,Fudge         ;get the address for the incoming fudge factor
    Mov  AX,[SI]          ;assign it to DX:AX
    Mov  DX,[SI+02]
    Mov  BX,44            ;divide the fudge factor to get a loop count
    Div  BX               ;now AX tells how many loops are needed for 1MS

Outer:
    Mov  LoWord,AX        ;set up for a 1 millisecond delay
    Mov  HiWord,0         ;use HiWord merely to match the code in FudgeFactor

Inner:
    Sub  LoWord,1         ;decrement HiWord:LoWord
    Sbb  HiWord,0
    Jns  Inner            ;do inner loop until Hi:Lo drops below zero
    Loop Outer            ;loop Milliseconds times

Exit:
    Ret                   ;return to BASIC

Pause3 Endp
End

