TYPE UBIntermecHHRecType         '
' First two fields are required by CE File I/O dll.
  CEVariant       AS STRING * 2  'MUST BE  CHR$(8) + CHR$(0)
  CEStrLen        AS STRING * 2  'MUST BE CHR$(165) + CHR$(0)
'**************************************************************
  CustName      AS STRING * 20
  ServAddress   AS STRING * 20
  MeterID       AS STRING * 8
  LowRead       AS STRING * 9
  HighRead      AS STRING * 9
  Account       AS STRING * 8
  ReadFlag      AS STRING * 1         'Y/N
  MeterType     AS STRING * 1
  Book          AS STRING * 2
  CurRead       AS STRING * 9
  PastRead      AS STRING * 9
  ReadDate      AS STRING * 8   ' fmt mmddyyyy
  Note1         AS STRING * 20
  Note2         AS STRING * 20
  Note3         AS STRING * 20
  NoteStatus    AS STRING * 1  'T=Temp  P=Perm
END TYPE



