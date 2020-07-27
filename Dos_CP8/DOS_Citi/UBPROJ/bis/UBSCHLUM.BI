TYPE SchlumHHType
  'ne=No Equivalent, r=Required, o=Optional
  'r=right justified l=left justified z=zero filled s=space filled
  Route        AS STRING * 10  'r       uBook-Sequence   LS
  WalkSeq      AS STRING * 4   'r
  PageNum      AS STRING * 4   'r
  ReadSeq      AS STRING * 2   'r
  HHID         AS STRING * 6   'o
  ReadDir      AS STRING * 1   'o ne
  NumDial      AS STRING * 1   'o ne
  IDExpected   AS STRING * 13  'r meter serial#
  IDCaptured   AS STRING * 13  'o
  IDOverride   AS STRING * 13  'o
  Decimals     AS STRING * 1   'o ne
  MtrRead      AS STRING * 10  'o
  ReadOVRide   AS STRING * 10  'o ne
  HighLimit    AS STRING * 10  'o
  LowLimit     AS STRING * 10  'o
  Date2Read    AS STRING * 6   'o
  Date2Exp     AS STRING * 6   'o
  NoteCodes    AS STRING * 8   'o ne
  LocatCode    AS STRING * 2   'o
  MtrRCode     AS STRING * 2   'o
  RecType      AS STRING * 2   'o   'reading status
  RecStatus    AS STRING * 1   'o ne internal debugging
  ReadDate     AS STRING * 6   'o must have  Actual read date
  ReadTime     AS STRING * 6   'o ne
  ReadType     AS STRING * 1   'o Reading class
  NetNumb      AS STRING * 2   'o ne
  ReadAtmpt    AS STRING * 1   'o ne
  UserChar     AS STRING * 7   '????? no code in schlum manual
  HHManufac    AS STRING * 1   'o ne
  ActStatus    AS STRING * 1   'o Account Status  A/I
  MtrType      AS STRING * 1   'o Type for automatic read
  ReadFailCode AS STRING * 1   'o Reading Failure code
  PrevRead     AS STRING * 10  'o Previous Mtr reading
  PrevDate     AS STRING * 6   'o Previous read date
  HHDisp1      AS STRING * 24  'r Service Addr
  HHDisp2      AS STRING * 24  'r Remote Location (Book-Sequence)
  HHDisp3      AS STRING * 24  'r More Location info
  HHDisp4      AS STRING * 24  'r Other Info (Maybe name)
  Notes1       AS STRING * 24  'o Notes Fields
  Notes2       AS STRING * 24  'o Notes Fields
  Notes3       AS STRING * 24  'o Notes Fields
  Notes4       AS STRING * 24  'o Notes Fields
  Notes5       AS STRING * 24  'o Notes Fields
  Notes6       AS STRING * 24  'o Notes Fields
  Notes7       AS STRING * 24  'o Notes Fields
  Notes8       AS STRING * 24  'o Notes Fields
  OpCode       AS STRING * 1   'r Reserved
  UBAcctNo     AS STRING * 6   'Sosoft Reserved Account Number
  MtrSlot      AS STRING * 1   'Sosoft Reserved Meter Slot
  UtilFld      AS STRING * 33  'o For us to use
  CrLf         AS STRING * 2   'r CrLf terminator
END TYPE

