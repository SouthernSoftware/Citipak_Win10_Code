TYPE W2DedType
    CHKDED  AS STRING * 20
    AMTBOX  AS STRING * 3
    DEDCODE AS STRING * 4
END TYPE

TYPE W2SetUpType
                           'only two options per check box  15c 15g
    ExtrYear AS INTEGER
'    CHKDED0  AS STRING * 20
'    AMTBOX0  AS STRING * 3
'    RETCODE  AS STRING * 4

    Deds(0 TO 12) AS W2DedType

END TYPE

'Note:
'CHKDED# Fields have the following 2 choices
'Pension Plan
'Def. Compensation


'W-2 menu items
'1) Zero current year totals
'       This would set emp3 data file totals to 0           
'2) Set extract options
'       This defines which check boxes to uses for certain deducts          
'3) Extract W-2 Information
'       Gets W-2 info for employee recs and trans hist
'4) Print W-2 Edit List
'       A summary of totals extracted into W-2 edit form
'5) Print W-2 Forms
'       Print the forms

'                                       Fld    Name
TYPE W2FormType                      '  ===  ========
    FEDWAGE  AS DOUBLE               '    1  FEDWAGE
    FEDTAXWH AS DOUBLE               '    2  FEDTAXWH
    SOCWAGE  AS DOUBLE               '    3  SOCWAGE
    SOCTAXWH AS DOUBLE               '    4  SOCTAXWH
    MEDWAGES AS DOUBLE               '    5  MEDWAGES
    MEDTAXWH AS DOUBLE               '    6  MEDTAXWH
    SOCTIPS  AS DOUBLE               '    7  SOCTIPS
    ALLOCTIP AS DOUBLE               '    8  ALLOCTIP
    ADVEIC   AS DOUBLE               '    9  ADVEIC
    DEPNDCAR AS DOUBLE               '   10  DEPNDCAR
    NQPLAN   AS DOUBLE               '   11  NQPLAN
    BENFBOX1 AS DOUBLE               '   12  BENFBOX1
    BOX13AMT AS DOUBLE               '   13  BOX13AMT
    BOX13TXT AS STRING * 4           '   14  BOX13TXT
    BOX14AMT AS DOUBLE               '   15  BOX14AMT
    BOX14TXT AS STRING * 4           '   16  BOX14TXT
    BOX13AM1 AS DOUBLE               '   17  BOX13AM1
    BOX13TX1 AS STRING * 4           '   18  BOX13TX1
    BOX14AM1 AS DOUBLE               '   19  BOX14AM1
    BOX14TX1 AS STRING * 4           '   20  BOX14TX1
    BOX15A   AS STRING * 1           '   21  BOX15A
    BOX15B   AS STRING * 1           '   22  BOX15B
    BOX15C   AS STRING * 1           '   23  BOX15C
    BOX15D   AS STRING * 1           '   24  BOX15D
    BOX15E   AS STRING * 1           '   25  BOX15E
    BOX15F   AS STRING * 1           '   26  BOX15F
    BOX15G   AS STRING * 1           '   27  BOX15G
    STATE    AS STRING * 2           '   28  STATE
    STAWAGE  AS DOUBLE               '   29  STAWAGE
    STATAXWH AS DOUBLE               '   30  STATAXWH
    LOCALNAM AS STRING * 16          '   31  LOCALNAM
    LOCWAGE  AS DOUBLE               '   32  LOCWAGE
    LOCALTAX AS DOUBLE               '   33  LOCALTAX
END TYPE

