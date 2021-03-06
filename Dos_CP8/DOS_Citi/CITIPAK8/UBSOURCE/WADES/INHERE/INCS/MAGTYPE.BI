TYPE Rec1A
  IDENT    AS STRING * 2
  PYEAR    AS STRING * 4
  EMPRNUM  AS STRING * 9
  BLANK1  AS STRING * 8
  FORADDR  AS STRING * 1
  EMPRNAME AS STRING * 50
  EMPRADDR AS STRING * 40
  BLANK2  AS STRING * 14
END TYPE

TYPE Rec2A
  IDENT    AS STRING * 2
  EMPRCITY    AS STRING * 25
  EMPRSTATE   AS STRING * 2
  BLANK1   AS STRING * 8
  EMPRZIPEXT   AS STRING * 5
  EMPRZIP   AS STRING * 5
  BLANK2   AS STRING * 81
END TYPE

TYPE Rec1B
  IDENT   AS STRING * 2
  PYEAR   AS STRING * 4
  EMPRNUM   AS STRING * 9
  EMPRCOMP   AS STRING * 8
  BLANK1   AS STRING * 105
END TYPE

TYPE Rec2B
  IDENT   AS STRING * 2
  BLANK1   AS STRING * 13
  FORADDR  AS STRING * 1
  ORGNAME  AS STRING * 44
  ORGADDR  AS STRING * 35
  ORGCITY  AS STRING * 20
  ORGSTATE  AS STRING * 2
  ORGZIPEXT  AS STRING * 5
  ORGZIP  AS STRING * 5
  BLANK2  AS STRING * 1
END TYPE

TYPE Rec1E
  IDENT  AS STRING * 2
  PYEAR  AS STRING * 4
  EMPRNUM  AS STRING * 9
  EMPRSTALOC  AS STRING * 9
  EMPRNAME  AS STRING * 50
  EMPRADDR  AS STRING * 40
  FORADDR  AS STRING * 1
  BLANK1  AS STRING * 13
END TYPE

TYPE Rec2E
  IDENT  AS STRING * 2
  EMPRCITY  AS STRING * 25
  EMPRSTATE  AS STRING * 10
  EMPRZIPEXT  AS STRING * 5
  EMPRZIP  AS STRING * 5
  EMPRNAMECODE  AS STRING * 1
  EMPRETYPE  AS STRING * 1
  BLANK1  AS STRING * 2
  ESTABNUM  AS STRING * 4
  LLINDICAT  AS STRING * 1
  BLANK2  AS STRING * 1
  OTHEREIN  AS STRING * 9
  BLANK3  AS STRING * 1
  TAXWITHTPP  AS STRING * 12
  BLANK4  AS STRING * 49
END TYPE

TYPE Rec1W
  IDENT  AS STRING * 2
  EMPSSN  AS STRING * 9
  EMPNAME  AS STRING * 27
  EMPADDR  AS STRING * 40
  EMPCITY  AS STRING * 25
  EMPSTATE  AS STRING * 10
  EMPZIPEXT  AS STRING * 5
  EMPZIP  AS STRING * 5
  EMPSTATU  AS STRING * 1
  BLANK1  AS STRING * 4
END TYPE
TYPE Rec2W
  IDENT  AS STRING * 2
  EMPSSW  AS STRING * 7
  BLANK1  AS STRING * 1
  EMPSST  AS STRING * 7
  BLANK2  AS STRING * 1
  EMPOTHERW  AS STRING * 9
  BLANK3  AS STRING * 1
  EMPSSTAX  AS STRING * 6
  EMPFEDTAX  AS STRING * 9
  EMPNQP  AS STRING * 9
  BLANK4  AS STRING * 1
  EMPNQPO  AS STRING * 9
  CTRLNUM  AS STRING * 7
  EMPRPGTLI AS STRING * 7
  EMPUNCTAX AS STRING * 7
  EMPEIC AS STRING * 7
  EMPATIP AS STRING * 7
  EMPFRINGE AS STRING * 9
  BLANK5 AS STRING * 1
  PENPLANI AS STRING * 1
  BLANK6 AS STRING * 1
  DEFCOMPI AS STRING * 1
  BLANK7 AS STRING * 1
  DEFCOMPC AS STRING * 9
  BLANK8 AS STRING * 1
  DEPCBEN AS STRING * 7
END TYPE
TYPE Rec3W
  IDENT AS STRING * 2
  MEDWAGES AS STRING * 9
  MEDTAX AS STRING * 7
  BLANK1 AS STRING * 110
END TYPE
TYPE Rec1I
  IDENT AS STRING * 2
  EMPSSW AS STRING * 10
  BLANK1 AS STRING * 1
  EMPSST AS STRING * 10
  BLANK2 AS STRING * 1
  EMPOTHERW AS STRING * 10
  BLANK3 AS STRING * 1
  EMPSSTAX AS STRING * 10
  BLANK4 AS STRING * 1
  EMPFEDTAX AS STRING * 10
  CRTLNUM AS STRING * 7
  EMPRPGTLI AS STRING * 10
  EMPUNCTAX AS STRING * 10
  EMPEIC AS STRING * 11
  EMPATIP AS STRING * 10
  EMPFRINGE AS STRING * 10
  BLANK5 AS STRING * 1
  DEFCOMPC AS STRING * 10
  BLANK6 AS STRING * 3
END TYPE
TYPE Rec2I
  IDENT AS STRING * 2
  DEPCBEN AS STRING * 10
  BLANK1 AS STRING * 1
  NQPLAN AS STRING * 10
  BLANK2 AS STRING * 1
  NQPLANO AS STRING * 10
  BLANK3 AS STRING * 1
  MEDWAGES AS STRING * 11
  BLANK5 AS STRING * 1
  MEDTAX AS STRING * 10
  BLANK6 AS STRING * 71
END TYPE
TYPE Rec1T
  IDENT AS STRING * 2
  NUMOFEMP AS STRING * 7
  SSWTOTAL AS STRING * 13
  BLANK1 AS STRING * 1
  SSTTOTAL AS STRING * 12
  OTHERTOTAL AS STRING * 13
  BLANK2 AS STRING * 1
  SSTAXTOTAL AS STRING * 12
  BLANK3 AS STRING * 1
  FEDTAXTOTAL AS STRING * 12
  ETLITOTAL AS STRING * 12
  UNTAXTOTAL AS STRING * 12
  EICTOTAL AS STRING * 12
  ATIPTOTAL AS STRING * 12
  BLANK4 AS STRING * 6
END TYPE
TYPE Rec2T
  IDENT AS STRING * 2
  FRINGETOTAL AS STRING * 12
  BLANK1 AS STRING * 1
  DEFCOMPCTOTAL AS STRING * 13
  BLANK2 AS STRING * 1
  DEPCBENTOTAL AS STRING * 12
  BLANK4 AS STRING * 1
  NQPTOTAL AS STRING * 13
  BLANK5 AS STRING * 1
  NQPOTOTAL AS STRING * 13
  BLANK6 AS STRING * 1
  MEDWAGETOTAL AS STRING * 13
  BLANK7 AS STRING * 1
  MEDTAXTOTAL AS STRING * 12
  BLANK8 AS STRING * 32
END TYPE

TYPE Rec1F
  IDENT AS STRING * 2
  NUMOFEMP AS STRING * 7
  SSWGTOTAL AS STRING * 16
  BLANK1 AS STRING * 1
  SSTGTOTAL AS STRING * 16
  BLANK2 AS STRING * 1
  OTHERWGTOTAL AS STRING * 16
  BLANK3 AS STRING * 1
  SSTAXGTOTAL AS STRING * 16
  BLANK4 AS STRING * 1
  FEDTAXGTOTAL AS STRING * 16
  BLANK5 AS STRING * 1
  EICGTOTAL AS STRING * 16
  BLANK6 AS STRING * 18
END TYPE

