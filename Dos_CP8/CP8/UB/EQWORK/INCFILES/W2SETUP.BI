TYPE W2PRNOPT
                           'only two options per check box  15c 15g
    CHKDEDR  AS STRING * 20
    AMTBOXR  AS INTEGER
    RETCODE  AS STRING * 4

    CHKDED1  AS STRING * 20
    AMTBOX1  AS INTEGER
    DEDCODE1 AS STRING * 4

    CHKDED2  AS STRING * 20
    AMTBOX2  AS INTEGER
    DEDCODE2 AS STRING * 4

    CHKDED3  AS STRING * 20
    AMTBOX3  AS INTEGER
    DEDCODE3 AS STRING * 4

    CHKDED4  AS STRING * 20
    AMTBOX4  AS INTEGER
    DEDCODE4 AS STRING * 4

    CHKDED5  AS STRING * 20
    AMTBOX5  AS INTEGER
    DEDCODE5 AS STRING * 4

    CHKDED6  AS STRING * 20
    AMTBOX6  AS INTEGER
    DEDCODE6 AS STRING * 4

    CHKDED7  AS STRING * 20
    AMTBOX7  AS INTEGER
    DEDCODE7 AS STRING * 4

    CHKDED8  AS STRING * 20
    AMTBOX8  AS INTEGER
    DEDCODE8 AS STRING * 4

    CHKDED9  AS STRING * 20
    AMTBOX9  AS INTEGER
    DEDCODE9 AS STRING * 4

    CHKDED10 AS STRING * 20
    AMTBOX10 AS INTEGER
    DEDCOD10 AS STRING * 4

    CHKDED11 AS STRING * 20
    AMTBOX11 AS INTEGER
    DEDCOD11 AS STRING * 4

    CHKDED12 AS STRING * 20
    AMTBOX12 AS INTEGER
    DEDCOD12 AS STRING * 4
END TYPE

'Note:
'CHKDED# Fields have the following 2 choices
'Pension Plan
'Def. Compensation


'W-2 menu items
'1) Zero current year totals
'	This would set emp3 data file totals to 0		
'2) Set extract options
'	This defines which check boxes to uses for certain deducts		
'3) Extract W-2 Information
'	Gets W-2 info for employee recs and trans hist
'4) Print W-2 Edit List
'	A summary of totals extracted into W-2 edit form
'5) Print W-2 Forms
'	Print the forms
