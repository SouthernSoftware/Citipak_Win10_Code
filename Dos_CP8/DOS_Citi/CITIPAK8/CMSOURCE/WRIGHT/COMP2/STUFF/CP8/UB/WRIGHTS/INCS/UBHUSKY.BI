TYPE MtrNotesType
  Note    AS STRING * 20
END TYPE

TYPE UBHuskyHHRecType                'File Layout for Sending Out Records
  CustName      AS STRING * 20
  ServAddress   AS STRING * 20
  UCode1        AS STRING * 2
  UCode2        AS STRING * 2
  MeterID       AS STRING * 8
  LowRead       AS LONG
  HighRead      AS LONG
  Account       AS STRING * 8
  ReadFlag      AS STRING * 1        'Y/N
  MeterType     AS STRING * 1
  Book          AS INTEGER
  CurRead       AS LONG
  PastRead      AS LONG
  ReadDate      AS INTEGER
  'ReadTime      AS STRING * 5
  NoteStatus    AS STRING * 1    'T=Temp Note  P=Perm Note
  Notes(1 TO 3)  AS MtrNotesType
END TYPE

