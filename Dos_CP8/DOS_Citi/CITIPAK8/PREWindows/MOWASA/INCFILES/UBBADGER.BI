TYPE BadgerRecType
  Fill1          AS STRING * 8   'fill spaces
  CustName       AS STRING * 20
  ServAddr       AS STRING * 20  'lj
  MtrNum1        AS STRING * 9   'Meter serial number
  Multi          AS STRING * 4   'n/a fill spaces
  Status         AS STRING * 1   'Acct Status 'A' or 'I'
  ReadCode       AS STRING * 1   'Read Code fill space
  ServFreq       AS STRING * 2   'Serv Code   'set to '1M'
  DNI            AS STRING * 2   'Dialog ID    'space fill
  MtrNum2        AS STRING * 9   'Same data as field above
  NumDials       AS STRING * 1   'N/A fill space
  HiRead         AS STRING * 9   'High reading limit
  LoRead         AS STRING * 9   'Low reading limit
  CurrRead       AS STRING * 9   'Current reading or space fill
  ReadTime       AS STRING * 8   'read time  space fill
  ReadCode2      AS STRING * 2   'space fill 'NOT same as above'
  CmntCode       AS STRING * 2   'space fill
  Fill2          AS STRING * 4   'space fill
  Account        AS STRING * 15  'location  '01-000001-01
  ReadDate       AS STRING * 8   'Read date ddmmyyyy
  DevCode        AS STRING * 1   'don't have a clue
  MMILat         AS STRING * 6   'space fill
  MMILong        AS STRING * 6   'space fill
  MMIChanl       AS STRING * 6   'space fill
  CircleCode     AS STRING * 2   'space fill
  SeqNumb        AS STRING * 6   'Read sequence number
  MfgModel       AS STRING * 20  'space fill
  UserField      AS STRING * 30  'we will use
  ReadID         AS STRING * 3   'space fill
  ReadCo1        AS STRING * 2   'space fill
  ReadCo2        AS STRING * 2   'space fill
  ReadCo3        AS STRING * 2   'space fill
  MMIReadCode    AS STRING * 6   'space fill
  Pad            AS STRING * 19  'space fill
  CrLf           AS STRING * 2   'Carriage return line feed sequence
END TYPE

