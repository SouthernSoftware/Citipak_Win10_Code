TYPE EmployeeIndexType
    DataRecNum    AS INTEGER
END TYPE

TYPE EmpNumType
    EmpNum        AS STRING * 10
END TYPE

TYPE EmpDedType
    DPct        AS STRING * 7
    DAmt        AS DOUBLE
    DOTI        AS STRING * 1
END TYPE

TYPE EmpWageDistType
    DAcct       AS STRING * 14
    DAmt        AS DOUBLE
END TYPE

TYPE EmpData1Type
    EMPNO       AS STRING * 10     '1-10
    EMPLNAME    AS STRING * 24     '11-34
    EMPFNAME    AS STRING * 24     '35-58
    Data1RecNum AS INTEGER         '59-60
    TransRecNum AS INTEGER         '61-62
    Deleted     AS INTEGER         '63-64
    'KillFlag    AS STRING * 1
    'FillerPad   AS STRING * 1
END TYPE

TYPE EmpData2Type         'new emp 2 rec
'------ EMPMA1
    EMPNO    AS STRING * 10
    EMPSSN   AS STRING * 11
    EMPLNAME AS STRING * 24
    EMPFNAME AS STRING * 24
    EMPADDR1 AS STRING * 36
    EMPADDR2 AS STRING * 36
    EMPCITY  AS STRING * 24
    EMPSTATE AS STRING * 2
    EMPZIP   AS STRING * 10
    EMPBDAY  AS INTEGER
    EMPGENDR AS STRING * 6
    EMPRACE  AS STRING * 14
    EMPRETNO AS STRING * 16
    EMPRETTP AS STRING * 24
'new
    DRAFTCOD AS STRING * 1
    EMPDDACC AS STRING * 20
    PRENOTED AS STRING * 1
    BANKNAME AS STRING * 33
    BANKLOC  AS STRING * 30
    TRANSIT  AS STRING * 9
'new
    EMPJOB   AS STRING * 28
    EMPWCCLS AS STRING * 12
    EMPSTATS AS STRING * 10
    EMPBCODE AS DOUBLE
    EMPPTYPE AS STRING * 10
    EMPPFREQ AS STRING * 16
    EMPPRATE AS DOUBLE
    EMPORATE AS DOUBLE
    EMPHDATE AS INTEGER
    EMPRDATE AS INTEGER
    EMPTDATE AS INTEGER
'------ EMPMA2
    EMPFEDX  AS STRING * 1
    EMPFEDO2 AS STRING * 1
    EMPFEDO1 AS DOUBLE
    EMPFEDS  AS STRING * 1
    EMPFEDA  AS INTEGER       'num of allowance
    EMPFEDAA AS DOUBLE
    EMPSTAX  AS STRING * 1
    EMPSTAO2 AS STRING * 1
    EMPSTAO1 AS DOUBLE
    EMPSTAS  AS STRING * 1
    EMPSTAA  AS INTEGER       '
    EMPSTAAA AS DOUBLE
    EMPSOCX  AS STRING * 1
    EMPMEDX  AS STRING * 1
    EMPEIC   AS STRING * 1

    EmpDed(1 TO 12)  AS EmpDedType

'------ page 3
    EMPEACT1 AS STRING * 14
    EMPEAMT1 AS DOUBLE
    EMPEACT2 AS STRING * 14
    EMPEAMT2 AS DOUBLE
    EMPEACT3 AS STRING * 14
    EMPEAMT3 AS DOUBLE

    EMPHP    AS STRING * 1

    EDist(1 TO 8)   AS EmpWageDistType

    EMPVACE  AS DOUBLE
    EMPVUSED AS DOUBLE
    EMPVBAL  AS DOUBLE

    EMPSLE   AS DOUBLE
    EMPSLUSE AS DOUBLE
    EMPSLBAL AS DOUBLE

    EMPCTE   AS DOUBLE
    EMPCTUSE AS DOUBLE
    EMPCTBAL AS DOUBLE

    PERERN   AS DOUBLE
    PERUSED  AS DOUBLE
    PERBAL   AS DOUBLE

    HOLERN   AS DOUBLE
    HOLUSED  AS DOUBLE
    HOLBAL   AS DOUBLE

    LeaveTbl AS INTEGER

    ExcludeESC  AS STRING * 1
    UseLife     AS STRING * 1

'------ Misc
    LastTransRec AS INTEGER
    EmpPin       AS INTEGER
    Deleted      AS INTEGER
    'for new leave table stuff
' PreNoteFlag     AS INTEGER
    Unused       AS STRING * 51
END TYPE



TYPE EmpData3Type
    Data1RecNum     AS INTEGER
    YTDGrossPay     AS DOUBLE
    YTDSocGrossPay  AS DOUBLE
    YTDMedGrossPay  AS DOUBLE
    YTDFedGrossPay  AS DOUBLE
    YTDStaGrossPay  AS DOUBLE
    YTDOTPay        AS DOUBLE
    YTDRegPay       AS DOUBLE
    YTDNet          AS DOUBLE
    YTDSocial       AS DOUBLE
    YTDMedicare     AS DOUBLE
    YTDFederal      AS DOUBLE
    YTDState        AS DOUBLE
    YTDRetire       AS DOUBLE
    YTDDAmt(1 TO 12) AS DOUBLE
    YTDDAmtT        AS DOUBLE
    YTDEarn1        AS DOUBLE        'e
    YTDEarn2        AS DOUBLE
    YTDEarn3        AS DOUBLE
    YTDEarnT        AS DOUBLE
    YTDEIC          AS DOUBLE
    YTDOther2       AS DOUBLE
END TYPE

TYPE EmprMatDistType
    MDesc       AS STRING * 20
    MAcct       AS STRING * 14
END TYPE

TYPE jEmpData2Type         'Ashe
'------ EMPMA1
    EMPNO    AS STRING * 10                  '1
    EMPSSN   AS STRING * 11                  '2
    EMPLNAME AS STRING * 24                  '3
    EMPFNAME AS STRING * 24                  '4
    EMPADDR1 AS STRING * 36                  '5
    EMPADDR2 AS STRING * 36                  '6
    EMPCITY  AS STRING * 24                  '7
    EMPSTATE AS STRING * 2                   '8
    EMPZIP   AS STRING * 10                  '9
    EMPBDAY  AS INTEGER                      '10
    EMPGENDR AS STRING * 6                   '11
    EMPRACE  AS STRING * 14                  '12
    EMPRETNO AS STRING * 16                  '13
    EMPRETTP AS STRING * 24                  '14
'----------------------------------
'New for Ach
    ACH      AS STRING * 1                   '15
    PRENOTED AS STRING * 1                   '16
    EMPBANK  AS STRING * 25                  '17
    EMPBNKCT AS STRING * 25                  '18
    ETRANSIT AS STRING * 9                   '19
    EBNKACCT AS STRING * 13                  '20
'----------------------------------
    EMPJOB   AS STRING * 28                  '21
    EMPWCCLS AS STRING * 12                  '22
    EMPSTATS AS STRING * 10                  '23
    EMPBCODE AS DOUBLE                       '24
    EMPPTYPE AS STRING * 10                  '25
    EMPPFREQ AS STRING * 16                  '26
    EMPPRATE AS DOUBLE                       '27
    EMPORATE AS DOUBLE                       '28
    EMPHDATE AS INTEGER                      '29
    EMPRDATE AS INTEGER                      '30
    EMPTDATE AS INTEGER                      '31
'------ EMPMA2
    EMPFEDX  AS STRING * 1                   '32
    EMPFEDO2 AS STRING * 1                   '33
    EMPFEDO1 AS DOUBLE                       '34
    EMPFEDS  AS STRING * 1                   '35
    EMPFEDA  AS INTEGER  'num of allowance   '36
    EMPFEDAA AS DOUBLE                       '37
    EMPSTAX  AS STRING * 1                   '38
    EMPSTAO2 AS STRING * 1                   '39
    EMPSTAO1 AS DOUBLE                       '40
    EMPSTAS  AS STRING * 1                   '41
    EMPSTAA  AS INTEGER       '              '42
    EMPSTAAA AS DOUBLE                       '43
    EMPSOCX  AS STRING * 1                   '44
    EMPMEDX  AS STRING * 1                   '45
    EMPEIC   AS STRING * 1                   '46
    EmpDed(1 TO 12)  AS EmpDedType      '47 48 49   50 51 52   53 54 55
                                        '56 57 58   59 60 61   62 63 64
                                        '65 66 67   68 69 70   71 72 73
                                        '74 75 76   77 78 79   80 81 82
'------ page 3
    EMPEACT1 AS STRING * 14                  '83
    EMPEAMT1 AS DOUBLE                       '84
    EMPEACT2 AS STRING * 14                  '85
    EMPEAMT2 AS DOUBLE                       '86
    EMPEACT3 AS STRING * 14                  '87
    EMPEAMT3 AS DOUBLE                       '88

    EMPHP    AS STRING * 1                   '89

    EDist(1 TO 8)   AS EmpWageDistType       '90 91  92  93  94  95  96  97
                                             '98 99 100 101 102 103 104 105
    EMPVACE  AS DOUBLE                       '106
    EMPVUSED AS DOUBLE                       '107
    EMPVBAL  AS DOUBLE                       '108
    EMPSLE   AS DOUBLE                       '109
    EMPSLUSE AS DOUBLE                       '110
    EMPSLBAL AS DOUBLE                       '111
    EMPCTE   AS DOUBLE                       '112
    EMPCTUSE AS DOUBLE                       '113
    EMPCTBAL AS DOUBLE                       '114
    LeaveTbl AS INTEGER                      '115
'------ EMPMA4
    EMPNOTES AS INTEGER                      '116
'------ Misc
    LastTransRec AS INTEGER                  '117
    EmpPin   AS INTEGER                      '118
    Deleted  AS INTEGER                      '119
'02-12-97 Added employer match types
    EmprMat(1 TO 3)   AS EmprMatDistType     '120 121 122   123 124 125
                                             '126 127 128
    Unused   AS STRING * 27                  '129
END TYPE

