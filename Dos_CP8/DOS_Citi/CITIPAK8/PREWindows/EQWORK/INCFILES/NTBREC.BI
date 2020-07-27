TYPE NTaxRecType
  Acct    AS STRING * 12
  LName   AS STRING * 60
  TheRest AS STRING * 440
END TYPE

TYPE NTaxRecType2
  Acct       AS STRING * 12
                             'We will only be sorting on the first
                             '"7" characters of the First&Last names
                             '

  FirstName  as string * 25
'      25 AS firstnam$,
  LastName   as string * 35
'  LastName  AS STRING * 60  'this speeds up the parsing of the
'  FirstFill  AS STRING * 18  'block transfer to the index.
                             'You uses a simple assignment to transfer data
                             'IndexRec(Cnt).FirstName = CustTaxRec.FirstName
                             '


  TheRest    AS STRING * 353
  Pin        AS STRING * 15
  Filler2    AS STRING * 72
END TYPE

TYPE NTax1RecType
  Stuff  AS STRING * 256
END TYPE

