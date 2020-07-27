DEFINT A-Z
  TYPE UBIntermecHHRecType 'First two fields REQUIRED for (Pocket PC File I/O dll)
    CEVariant     AS STRING * 2  'MUST BE  CHR$(8) + CHR$(0)
    CEStrLen      AS STRING * 2  'MUST BE CHR$(165) + CHR$(0)
    '**************************************************************
    CustName      AS STRING * 20  '24
    ServAddress   AS STRING * 20  '44
    MeterID       AS STRING * 8   '52
    LowRead       AS STRING * 9   '61
    HighRead      AS STRING * 9   '70
    Account       AS STRING * 8   '78
    ReadFlag      AS STRING * 1   '79     'Y/N
  END TYPE

  DIM RecTmp AS UBIntermecHHRecType
  PRINT LEN(RecTmp)

