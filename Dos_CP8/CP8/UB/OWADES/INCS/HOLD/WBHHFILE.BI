' User-Defined Type Declaration Section
TYPE DataRecType
  WReadFlag     AS STRING * 1
  WPhysical     AS STRING * 11
  Line1         AS STRING * 20
  Line2         AS STRING * 20
  WPrevious     AS SINGLE
  WCurrent      AS SINGLE
  WAverageUse   AS SINGLE
  WCurrentUse   AS SINGLE
  WDigits       AS INTEGER
  GPrevious     AS SINGLE
  GCurrent      AS SINGLE
  GAverageUse   AS SINGLE
  GCurrentUse   AS SINGLE
  GDigits       AS INTEGER
  Note1         AS STRING * 20
  Note2         AS STRING * 20
  Note3         AS STRING * 20
  NoteChg       AS STRING * 1
  Location      AS STRING * 9
  GPhysical     AS STRING * 11
  AccountType   AS STRING * 1
  GReadFlag     AS STRING * 1
  NotUsed       AS STRING * 29
END TYPE

