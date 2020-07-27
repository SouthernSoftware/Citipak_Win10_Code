'ASM Procedures from FORMS[7].QLB/LIB
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Direction%)
DECLARE FUNCTION Exist% (FileName$)
DECLARE SUB HideCursor ()
DECLARE FUNCTION Monitor% ()
DECLARE SUB MScrnSave (UlRow, UlCol, BRRow, BRCol, SEG Address)
DECLARE SUB MScrnRest (UlRow, UlCol, BRRow, BRCol, SEG Address)
DECLARE SUB ShowCursor ()
DECLARE SUB TextCursor (Fg%, Bg%)
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)

