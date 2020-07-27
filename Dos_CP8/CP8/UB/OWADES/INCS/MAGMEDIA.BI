'Diskette Data Records - General Requirements
'
'All data records must be a fixed length of 128 bytes.  Deviations
'from the prescribed record formats will prevent proper processing
'of your file by SSA.  A properly composed W-2 Copy A diskette
'file is comprised of the following records:
'
'Codes 1A & 2A  - Transmitter Records
'Codes 1B & 2B  - Basic Information Records
'Codes 1E & 2E  - Employer Records
'Codes 1W&2W&3W - Employee Wage Records
'Codes 1S & 2S  - Supplemental Records (Optional)
'Codes 1I & 2I  - Intermediate Total Records
'Codes 1T & 2T  - Total Records
'Code  1F       - Final Record
'
'NOTE:  Any data preceding the 1A identifier or which follows the
'       1F record may prevent SSA from processing your report.
'
'
'                 IV.  DISKETTE SPECIFICATIONS
'                   FOR W-2 COPY A REPORTING
'
'Diskette Specifications:  Annual W-2 Copy A Information
'Date:  July 1994
'
'Record Name:  Code 1A & 2A - Transmitter Records
'
'Code 1A - Transmitter Record
'Length=128
'
'Location   Field           Length   Description & Remarks
'
TYPE Rec1A
'1-2      Record Identifier    2     Constant "1A".
 Ident    AS STRING * 2
'
'3-6      Payment Year         4     Enter the year for which this
 PYEAR    AS STRING * 4

'          (Tax Year)                report is being prepared.
'                                    UPDATE EACH YEAR.  Example:
'                                    1994
'
 EMPRNUM  AS STRING * 9
'7-15     Transmitter's        9     Enter only NUMERIC
'         Employer                   characters.  Omit hyphens,
'         Identification             prefixes and suffixes.  This
'         Number (EIN)               number should NOT begin with
'                                    00, 07, 08, 09, 10, 17, 18,
'                                    19, 20, 26, 27, 28, 29, 30,
'                                    40, 49, 50, 60, 69, 70, 78,
'                                    79, 80, 89 or 90.  NOTE:
'                                    This EIN should match the EIN
'                                    on the file's external label.
'

 BLANK1  AS STRING * 8
'16-23    Blank                8     Enter blanks.  Reserved for
'                                    SSA use.
'

 FORADDR  AS STRING * 1
'24       Foreign Address      1     If the information shown in
'         Indicator                  positions 75-114 of the Code
'                                    1A record and positions 3-47
'                                    of the Code 2A record is for
'                                    a foreign address (i.e.,
'                                    outside of the U.S. and
'                                    U.S. territories and
'                                    possessions, and not APO or
'                                    FPO), enter the letter "X" in
'                                    this field.  Otherwise, enter
'                                    a blank.
'
   EMPRNAME AS STRING * 50
'25-74    Transmitter Name     50    Enter the name of the
'                                    organization transmitting
'                                    this file.  Left justify and
'                                    fill with blanks.
'

   EMPRADDR AS STRING * 40
'75-114   Street Address       40    Enter the street address of
'                                    the transmitter.  Left
'                                    justify and fill with blanks.
'

 BLANK2  AS STRING * 14
'115-128  Blank                14    Enter blanks.  Reserved for
'                                    SSA use.
'
'
END TYPE


'Code 2A - Transmitter Record
'Length=128
'
'Location   Field           Length   Description & Remarks
'

TYPE Rec2A

 Ident    AS STRING * 2
'1-2      Record Identifier    2     Constant "2A".
'

 EMPRCITY    AS STRING * 25
'3-27     City                 25    Left justify and fill with
'                                    blanks.  For a foreign
'                                    address, include name of
'                                    foreign "state" and
'                                    country; abbreviate city
'                                    and state as necessary;
'                                    show full country name.
'
 EMPRSTATE   AS STRING * 2
'28-29    State                2     Use a standard FIPS postal
'                                    abbreviation (Appendix B).
'                                    For a foreign address, enter
'                                    blanks.
'

 BLANK1   AS STRING * 8
'30-37    Blank                8     Enter blanks.  Reserved for
'                                    SSA use.
'

 EMPRZIPEXT   AS STRING * 5
'38-42    ZIP Code             5     Use this field as necessary
'         Extension                  for the four-digit extension
'                                    of the Zip Code, being sure
'                                    to include the hyphen in
'                                    position 38.  If this is a
'                                    foreign address, use this
'                                    field as necessary for
'                                    overflow for a Foreign
'                                    Postal Code begun in
'                                    positions 43-47; left
'                                    justify and fill with blanks.
'                                    If this field is not
'                                    applicable, enter blanks.
'

 EMPRZIP   AS STRING * 5
'43-47    ZIP Code or Foreign  5     Enter a valid ZIP Code.  For
'         Postal Code                a foreigh address, use this
'                                    field for the Foreign Postal
'                                    Code, if applicable; left
'                                    justify and fill with blanks.
'                                    If necessary, continue the
'                                    Foreign Postal Code in
'                                    positions 38-42.
'

 BLANK2   AS STRING * 81
'48-128   Blank                81    Enter blanks.  Reserved for
'                                    SSA use.
'
END TYPE

'
'Diskette Specifications:  Annual W-2 Copy A Information
'Date:  July 1994
'
'Record Name:  Codes 1B & 2B - Basic Information Records
'
'Code 1B - Basic Information Record
'Length=128
'
'Location   Field           Length   Description & Remarks
TYPE Rec1B

 Ident   AS STRING * 2
'
'1-2      Record Identifier    2     Constant "1B".
'

 PYEAR   AS STRING * 4
'3-6      Payment Year         4     Enter the year for which this
'          (Tax Year)                report is being prepared.
'                                    UPDATE EACH YEAR.  Example:
'                                    1994
'

 EMPRNUM   AS STRING * 9
'7-15     Transmitter's        9     Enter only NUMERIC
'         Employer                   characters.  Omit hyphens,
'         Identification             prefixes and suffixes.  This
'         Number (EIN)               number should NOT begin with
'                                    00, 07, 08, 09, 10, 17, 18,
'                                    19, 20, 26, 27, 28, 29, 30,
'                                    40, 49, 50, 60, 69, 70, 78,
'                                    79, 80, 89 or 90.  NOTE:
'                                    This EIN should match the
'                                    EIN on the Code 1A
'                                    (Transmitter) record and the
'                                    EIN on the file's external
'                                    label.
'

 EMPRCOMP   AS STRING * 8
'16-23    Computer             8     Enter the manufacturer's
'                                    name.  Left justify and fill
'                                    with blanks.
'

 BLANK1   AS STRING * 105
'24-128   Blank                105   Enter blanks.  Reserved for
'                                    SSA use.
'
'
END TYPE
'Code 2B - Basic Information Record
'Length=128
'
'Location   Field           Length   Description & Remarks
'
TYPE Rec2B

 Ident   AS STRING * 2
'1-2      Record Identifier    2     Constant "2B".
'

 BLANK1   AS STRING * 13
'3-15     Blank                13    Enter blanks.  Reserved for
'                                    SSA use.
'

 FORADDR  AS STRING * 1
'16       Foreign Address      1     If the information shown in
'         Indicator                  positions 61-127 is for a
'                                    foreign address (i.e.,
'                                    outside of the U.S. and U.S.
'                                    territories and possessions,
'                                    and not APO or FPO), enter
'                                    the letter "X" in this field.
'                                    Otherwise, enter a blank.
'

 ORGNAME  AS STRING * 44
'17-60    Organization         44    Enter the name of the
'         Name                       organization to which annual
'                                    filing instructions should
'                                    be sent and to which the file
'                                    should be returned if it
'                                    cannot be processed.  Left
'                                    justify and fill with blanks.
'                                    Show the mailing address of
'                                    the organization in positions
'                                    61-127.
'

 ORGADDR  AS STRING * 35
'61-95    Street               35    Enter the street address for
'         Address                    the organization shown in
'                                    positions 17-60.  Left
'                                    justify and fill with blanks.
'

 ORGCITY  AS STRING * 20
'96-115   City                 20    Left justify and fill with
'                                    blanks.  For a foreign
'                                    address, include name of
'                                    foreign "state" and country;
'                                    abbreviate city and state as
'                                    necessary; show full
'                                    country name.
'

 ORGSTATE  AS STRING * 2
'116-117  State                2     Use a standard FIPS postal
'                                    abbreviation (Appendix B).
'                                    For a foreign address, enter
'                                    blanks.
'
 ORGZIPEXT  AS STRING * 5
'118-122  ZIP Code Extension   5     Use this field as necessary
'                                    for the four-digit extension
'                                    of the Zip Code, being sure
'                                    to include the hyphen in
'                                    position 118.  If this is a
'                                    foreign address, use this
'                                    field as necessary for
'                                    overflow for a Foreign Postal
'                                    Code begun in positions
'                                    123-127; left justify and
'                                    fill with blanks.  If this
'                                    field is not applicable,
'                                    enter blanks.
'

 ORGZIP  AS STRING * 5
'123-127  ZIP Code/Foreign     5     Enter a valid ZIP Code.  For
'         Postal Code                a foreign address, use this
'                                    field for the Foreign Postal
'                                    Code, if applicable; left
'                                    justify and fill with blanks.
'                                    If necessary, continue the
'                                    Foreign Postal Code in
'                                    positions 118-122.
'

 BLANK2  AS STRING * 1
'128      Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'
END TYPE
'
'Diskette Specifications:  Annual W-2 Copy A Information
'Date:  July 1994
'
'Record Name:  Codes 1E & 2E - Employer Records
'
'Code 1E - Employer Record
'Length=128
'
'Location   Field           Length   Description & Remarks
TYPE Rec1E
 Ident  AS STRING * 2
'
'1-2      Record Identifier    2     Constant "1E".
'

 PYEAR  AS STRING * 4
'3-6      Payment Year         4     Enter the year for which this
'          (Tax Year)                report is being prepared.
'                                    Enter numeric characters
'                                    only.  UPDATE EACH YEAR.
'                                    NOTE:  All records within a
'                                    file must be for the same
'                                    payment year.
'

 EMPRNUM  AS STRING * 9
'7-15     Employer             9     Enter only NUMERIC
'         Identification             characters.  Omit all
'         Number (EIN)               hyphens, prefixes and
'                                    suffixes.  This number should
'                                    NOT begin with 00, 07, 08,
'                                    09, 10, 17, 18, 19, 20, 26,
'                                    27, 28, 29, 30, 40, 49, 50,
'                                    60, 69, 70, 78, 79, 80, 89 or
'                                    90.
'

 EMPRSTALOC  AS STRING * 9
'16-24    State/Local          9     Enter blanks if payment year
'         69 Number                  is 1987 or later.  See
'                                    Glossary for further
'                                    explanation of 69 Number.
'

 EMPRNAME  AS STRING * 50
'25-74    Employer Name        50    Left justify and fill with
'                                    blanks.
'

 EMPRADDR  AS STRING * 40
'75-114   Street Address       40    Left justify and fill with
'                                    blanks.
'

 FORADDR  AS STRING * 1
'115      Foreign Address      1     If the information shown in
'         Indicator                  positions 75-114 of the Code
'                                    1E record and in positions
'                                    3-47 of the Code 2E record is
'                                    for a foreign address (i.e.,
'                                    outside of the U.S. and U.S.
'                                    territories and possessions
'                                    and not APO or FPO), enter
'                                    the letter "X" in this field.
'                                    Otherwise, enter a blank.
'

 BLANK1  AS STRING * 13
'116-128  Blank                13    Enter blanks. Reserved for
'                                    SSA use.
'
END TYPE
'
'Code 2E - Employer Record
'Length=128
'
'Location   Field           Length   Description & Remarks
'
TYPE Rec2E
 IDENT  AS STRING * 2
'1-2      Record Identifier    2     Constant "2E".
'

 EMPRCITY  AS STRING * 25
'3-27     City                 25    Left justify and fill with
'                                    blanks.  For a foreign
'                                    address, include name of
'                                    foreign "state" and country;
'                                    abbreviate city and state as
'                                    necessary; show full country
'                                    name.
'

 EMPRSTATE  AS STRING * 10
'28-37    State                10    Use a standard FIPS postal
'                                    abbreviation (Appendix B).
'                                    For a foreign address, enter
'                                    blanks.
'

 EMPRZIPEXT  AS STRING * 5
'38-42    ZIP Code             5     Use this field as necessary
'         Extension                  for the four-digit extension
'                                    of the Zip Code, being sure
'                                    to include the hyphen in
'                                    position 38.  If this is a
'                                    foreign address, use this
'                                    field as necessary for
'                                    overflow for a Foreign Postal
'                                    Code begun in positions
'                                    43-47; left justify and fill
'                                    with blanks.  If this field
'                                    is not applicable, enter
'                                    blanks.
'

 EMPRZIP  AS STRING * 5
'43-47    ZIP Code/Foreign     5     Enter a valid ZIP Code.  For
'         Postal Code                a foreign address, use this
'                                    field for the Foreign Postal
'                                    Code, if applicable; left
'                                    justify and fill with blanks.
'                                    If necessary, continue the
'                                    Foreign Postal Code in
'                                    positions 38-42 above.
'

 EMPRNAMECODE  AS STRING * 1
'48       Name Code            1     Enter "S" if the surname
'                                    appears first in the Employee
'                                    Name field (positions 12-38)
'                                    of the following Code 1W
'                                    records.  Enter "F" if the
'                                    first name appears first in
'                                    the Employee Name field of
'                                    the following Code 1W
'                                    records.  This code may vary
'                                    with each employer record
'                                    (Code 1E/2E) as long as the
'                                    code is consistent with the
'                                    name format on the associated
'                                    Code 1W records.
'

 EMPRETYPE  AS STRING * 1
'49       Type of Employment   1     Enter the appropriate code
'                                    A - Agriculture
'                                    H - Household
'                                    M - Military
'                                    Q - Medicare Qualified
'                                        Government Employment
'                                    X - Railroad
'                                    R - Regular (All others)
'

 BLANK1  AS STRING * 2
'50-51    Blank                2     Enter blanks.  Reserved for
'                                    SSA use.
'

 ESTABNUM  AS STRING * 4
'52-55    Establishment Number 4     Enter either the
'         OR Coverage Group/         Establishment Number, the
'         Payroll Record Unit        Coverage Group/Payroll Record
'         (PRU) Number               Unit Number, or other
'                                    identifier, whichever is
'                                    applicable.  See Glossary
'                                    for explanation.  Otherwise,
'                                    enter blanks.
'

 LLINDICAT  AS STRING * 1
'56       Limitation of        1     Enter a blank if payment year
'         Liability Indicator        is 1987 or later.  See
'                                    Glossary for explanation.
'                                    If applicable, enter the
'                                    letter "L".  Otherwise, enter
'                                    a blank.  Refer to Glossary.
'

 BLANK2  AS STRING * 1
'57       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'

 OTHEREIN  AS STRING * 9
'58-66    Other EIN            9     If a form 941, 942 or 943 was
'                                    submitted to IRS for the same
'                                    tax year as this report and
'                                    the form used an EIN other
'                                    than the EIN in location
'                                    7-15 of the Code 1E record,
'                                    enter the other EIN.  Enter
'                                    blanks if no other EIN was
'                                    used.
'

 BLANK3  AS STRING * 1
'67       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'

 TAXWITHTPP  AS STRING * 12
'68-79    Income Tax Withheld  12    No negative amounts; right
'         by Third-Party Payer       justify and zero fill.
'                                    Enter the total Federal
'                                    income tax withheld by
'                                    third-parties (generally
'                                    insurance companies) from
'                                    sick or disability payments
'                                    made to the employees in
'                                    the following Code W records.
'

 BLANK4  AS STRING * 49
'80-128   Blank                49    Enter blanks.  Reserved for
'                                    SSA use.
'
'
END TYPE

'Diskette Specifications:  Annual W-2 Copy A Information
'Date:  July 1994
'
'Record Name:  Codes 1W & 2W & 3W - Employee Wage Records
'
'Code 1W - Employee Wage Record
'Length=128
'
'Location   Field           Length   Description & Remarks
TYPE Rec1W
 IDENT  AS STRING * 2
'
'1-2      Record Identifier    2     Constant "1W".
'

 EMPSSN  AS STRING * 9
'3-11     Social Security      9     Enter the employee's Social
'         Number (SSN)               Security number.
'                                    a.  Enter only NUMERIC
'                                        characters.  If the SSN
'                                        is not available, enter
'                                        the letter "I" in
'                                        position 3 and enter
'                                        blanks in positions 4-11.
'                                    b.  Omit hyphens, prefixes
'                                        and suffixes.
'                                    c.  May not be 000000000,
'                                        111111111, 333333333, or
'                                        123456789.
'                                    d.  Use the number shown on
'                                        the original/replacement
'                                        SSN card issued to the
'                                        employee by SSA.
'                                    e.  Check the three left-most
'                                        digits of the SSN.  These
'                                        digits may not be:  800
'                                        through 999, 729 through
'                                        799, or 000.
'
 EMPNAME  AS STRING * 27
'12-38    Employee Name        27    Format of employee name must
'                                    agree with the Name Code in
'                                    position 48 of the related
'                                    Code 2E record.  Left
'                                    justify and fill with blanks.
'                                    Refer to Name Formats in
'                                    Section II.
'

 EMPADDR  AS STRING * 40
'39-78    Street Address       40    Left justify and fill with
'                                    blanks.
'

 EMPCITY  AS STRING * 25
'79-103   City                 25    Left justify and fill with
'                                    blanks.  For a foreign
'                                    address, include name
'                                    of foreign "state" and
'                                    country; abbreviate
'                                    city and state as necessary;
'                                    show full country name.
'
 EMPSTATE  AS STRING * 10
'104-113  State                10    Use a standard FIPS postal
'                                    abbreviation (Appendix B).
'                                    For a foreign address, enter
'                                    blanks.
'
 EMPZIPEXT  AS STRING * 5
'114-118  ZIP Code Extension   5     Use this field as necessary
'                                    for the four-digit extension
'                                    of the Zip Code, being sure
'                                    to include the hyphen in
'                                    position 114.  If this is a
'                                    foreign address, use this
'                                    field as necessary for
'                                    overflow for a Foreign Postal
'                                    Code begun in positions
'                                    119-123; left justify and
'                                    fill with blanks.  If this
'                                    field is not applicable,
'                                    enter blanks.
'
 EMPZIP  AS STRING * 5
'119-123  ZIP Code or Foreign  5     Enter a valid ZIP Code.  For
'         Postal Code                a foreign address, use this
'                                    field for the Foreign Postal
'                                    Code, if applicable; left
'                                    justify and fill with blanks.
'                                    If necessary, continue the
'                                    Foreign Postal Code in
'                                    positions 114-118 above.
'
 EMPSTATU  AS STRING * 1
'124      Statutory Employee   1     Enter "S" for a statutory
'         Code                       employee; otherwise, enter a
'                                    blank.  Refer to Glossary.
'

 BLANK1  AS STRING * 4
'125-128  Blank                4     Enter blanks.  Reserved for
'                                    SSA use.
'
'
END TYPE
'Code 2W - Employee Wage Record
'Length=128
'
'Location   Field           Length   Description & Remarks
TYPE Rec2W
 IDENT  AS STRING * 2
'1-2      Record Identifier    2     Constant "2W".
'
 EMPSSW  AS STRING * 7
'3-9      Annual Social        7     No negative amounts; right
'         Security Wages             justify and zero fill.  The
'                                    sum of the amount in this
'                                    field and the amount in the
'                                    Annual Social Security Tips
'                                    field should not exceed the
'                                    annual maximum social
'                                    security wage base for the
'                                    tax year being reported
'                                    ($60,600.00 for 1994).
'                                    Refer to Money Amounts in
'                                    Section II.
'

 BLANK1  AS STRING * 1
'10       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'
 EMPSST  AS STRING * 7
'11-17    Annual Social        7     No negative amounts; right
'         Security Tips              justify and zero fill.  The
'                                    sum of the amount in this
'                                    field and the amount in the
'                                    Annual Social Security Wages
'                                    field should not exceed the
'                                    annual maximum social
'                                    security wage base for the
'                                    tax year being reported
'                                    ($60,600.00 for 1994).
'
 BLANK2  AS STRING * 1
'18       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'
 EMPOTHERW  AS STRING * 9
'19-27    Annual Wages, Tips   9     No negative amounts; right
'         and Other                  justify and zero fill.
'         Compensation
'
 BLANK3  AS STRING * 1
'28       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'
 EMPSSTAX  AS STRING * 6
'29-34    Social Security      6     No negative amounts; right
'         Employee Tax               justify and zero fill.
'         Withheld
'
 EMPFEDTAX  AS STRING * 9
'35-43    Federal Income Tax   9     No negative amounts; right
'         Withheld                   justify and zero fill.
'
 EMPNQP  AS STRING * 9
'44-52    Non-qualified Plan   9     Enter the amount of
'         Section 457                distributions from a Section
'         Distributions              457 plan.  No negative
'                                    amounts; right justify and
'                                    zero fill.
'
 BLANK4  AS STRING * 1
'53       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'

 EMPNQPO  AS STRING * 9
'54-62    Non-qualified Plan   9     Enter the amount of
'         Not Section 457            distributions from any
'         Distributions              non-qualified deferred
'                                    compensation plans that
'                                    are not Section 457 plans.
'                                    No negative amounts; right
'                                    justify and zero fill.
'

 CTRLNUM  AS STRING * 7
'63-69    Control Number       7     OPTIONAL; refer to Glossary.
'
 EMPRPGTLI AS STRING * 7
'70-76    Employer's Cost of   7     No negative amounts; right
'         Premiums for Group         justify and zero fill.
'         Term Life Insurance
'         Over $50,000
'
 EMPUNCTAX AS STRING * 7
'77-83    Uncollected Employee 7     Combine the uncollected
'         Tax on Tips                Social Security tax with the
'                                    uncollected Medicare tax and
'                                    report as one amount in this
'                                    field.  No negative amounts;
'                                    right justify and zero fill.
'
 EMPEIC AS STRING * 7
'84-90    Advance Earned       7     No negative amounts; right
'         Income Credit              justify and zero fill.
'
 EMPATIP AS STRING * 7
'91-97    Allocated Tips       7     No negative amounts; right
'                                    justify and zero fill.
'
 EMPFRINGE AS STRING * 9
'98-106   Fringe Benefits      9     No negative amounts; right
'                                    justify and zero fill.
'

 BLANK5 AS STRING * 1
'107      Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'
 PENPLAN AS STRING * 1
'108      Pension Plan         1     If applicable (see
'         Indicator                  Glossary--Appendix E), enter
'                                    the letter "P".  Otherwise,
'                                    enter a blank.
'
 BLANK6 AS STRING * 1
'109      Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'
 DEFCOMPI AS STRING * 1
'110      Deferred             1     Enter the letter "D" if the
'         Compensation               amount in positions 112-120
'         Indicator                  is greater than zero.
'                                    Otherwise, enter a blank.
'
 BLANK7 AS STRING * 1
'111      Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'
 DEFCOMPC AS STRING * 9
'112-120  Deferred             9     No negative amounts; right
'         Compensation               justify and zero fill.  NOTE:
'         Contributions              This amount should equal the
'                                    total of all deferred
'                                    compensation contributions to
'                                    the following plans:  401(k),
'                                    403(b), 408(k)(6), 457(b),
'                                    and 501(c)(18)(D).  DO NOT
'                                    report contributions to
'                                    Section 457(f) plans.
'
 BLANK8 AS STRING * 1
'121      Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'

 DEPCBEN AS STRING * 7
'122-128  Dependent Care       7     No negative amounts; right
'         Benefits                   justify and zero fill.
'
END TYPE
'
'Code 3W - Employee Wage Record
'Length=128
'
'Location   Field           Length   Description and Remarks
'
TYPE Rec3W
 IDENT AS STRING * 2
'1-2      Record Identifier    2     Constant "3W".
'
 MEDWAGE AS STRING * 9
'3-11     Medicare Wages       9     No negative amounts; right
'         and Tips                   justify and zero fill.  For
'                                    years prior to 1994, do not
'                                    exceed the annual maximum
'                                    Medicare wage base for the
'                                    tax year being reported.  DO
'                                    NOT USE THIS FIELD TO REPORT
'                                    DATA FOR TAX YEAR 1990 OR
'                                    EARLIER.  FOR 1991 AND LATER,
'                                    THIS FIELD MUST BE GREATER
'                                    THAN OR EQUAL TO THE SUM OF
'                                    THE FIELDS FOR ANNUAL SOCIAL
'                                    SECURITY WAGES AND ANNUAL
'                                    SOCIAL SECURITY TIPS.
'
 MEDTAX AS STRING * 7
'12-18    Medicare Tax         7     No negative amounts; right
'         Withheld                   justify and zero fill.  DO
'                                    NOT USE THIS FIELD TO REPORT
'                                    DATA FOR TAX YEAR 1990 OR
'                                    EARLIER.
'
 BLANK1 AS STRING * 110
'19-128   Blank                110   Enter blanks.  Reserved for
'                                    SSA use.
'
END TYPE
'
'Diskette Specifications:  Annual W-2 Copy A Information
'Date:  July 1994
'
'Record Name:  Codes 1I & 2I - Intermediate Total Records
'
'Code 1I - Intermediate Total Record
'Length=128
'
'Location   Field           Length   Description & Remarks
TYPE Rec1I
 IDENT AS STRING * 2
'1-2      Record Identifier    2     Constant "1I" (= numeric one
'                                    + letter I).
'

 EMPSSW AS STRING * 10
'3-12     Annual Social        10    Enter the segment total.  No
'         Security Wages             negative amounts.  Right
'                                    justify and zero fill.
'

 BLANK1 AS STRING * 1
'13       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'

 EMPSST AS STRING * 10
'14-23    Annual Social        10    Enter the segment total.  No
'         Security Tips              negative amounts.  Right
'                                    justify and zero fill.
'
 BLANK2 AS STRING * 1
'24       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.

 EMPOTHERW AS STRING * 10
'25-34    Annual Wages, Tips   10    Enter the segment total.  No
'         and Other                  negative amounts.  Right
'         Compensation               justify and zero fill.
'
 BLANK3 AS STRING * 1
'35       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'

 EMPSSTAX AS STRING * 10
'36-45    Social Security      10    Enter the segment total.  No
'         Employee Tax               negative amounts.  Right
'         Withheld                   justify and zero fill.
'

 BLANK4 AS STRING * 1
'46       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'
 EMPFEDTAX AS STRING * 10
'47-56    Federal Income Tax   10    Enter the segment total.  No
'         Withheld                   negative amounts.  Right
'                                    justify and zero fill.
'
 CRTLNUM AS STRING * 7
'57-63    Control Number       7     Use of this field is
'                                    OPTIONAL.  Refer to
'                                    Glossary.
'
 EMPRPGTLI AS STRING * 10
'64-73    Employer Cost of     10    Enter the segment total.  No
'         Premiums for Group         negative amounts.  Right
'         Term Life Insurance        justify and zero fill.
'         Over $50,000
'
 EMPUNCTAX AS STRING * 10
'74-83    Uncollected Employee 10    Enter the segment total.  No
'         Tax on Tips                negative amounts.  Right
'                                    justify and zero fill.
'
 EMPEIC AS STRING * 11
'84-94    Advance Earned       11    Enter the segment total.  No
'         Income Credit              negative amounts.  Right
'                                    justify and zero fill.
'
 EMPATIP AS STRING * 10
'95-104   Allocated Tips       10    Enter the segment total.  No
'                                    negative amounts.  Right
'                                    justify and zero fill.
'
 EMPFRINGE AS STRING * 10
'105-114  Fringe Benefits      10    Enter the segment total.  No
'                                    negative amounts.  Right
'                                    justify and zero fill.
'
 BLANK5 AS STRING * 1
'115      Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'

 DEFCOMPC AS STRING * 10
'116-125  Deferred             10    Enter the segment total.  No
'         Compensation               negative amounts.  Right
'         Contributions              justify and zero fill.
'
 BLANK6 AS STRING * 3
'126-128  Blank                3     Enter blanks.  Reserved for
'                                    SSA use.
'
END TYPE
'
'Code 2I - Intermediate Total Record
'Length=128
'
'Location   Field           Length   Description & Remarks
TYPE Rec2I
 IDENT AS STRING * 2
'1-2      Record Identifier    2     Constant "2I" (=numeric two +
'                                    letter I).
'

 DEPCBEN AS STRING * 10
'3-12     Dependent Care       10    Enter the segment total.  No
'         Benefits                   negative amounts.  Right
'                                    justify and zero fill.
'
 BLANK1 AS STRING * 1
'13       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'
 NQPLAN AS STRING * 10
'14-23    Non-qualified Plan   10    Enter the segment total.  No
'         Section 457                negative amounts.  Right
'         Distributions              justify and zero fill.
'
 BLANK2 AS STRING * 1
'24       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'
 NQPLANO AS STRING * 10
'25-34    Non-qualified Plan   10    Enter the segment total.  No
'         Not Section 457            negative amounts.  Right
'         Distributions              justify and zero fill.
'
 BLANK3 AS STRING * 1
'35       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'
 MEDWAGE AS STRING * 11
'36-46    Medicare Wages and   11    Enter the segment total.  No
'         Tips                       negative amounts.  Right
'                                    justify and zero fill.  DO
'                                    NOT USE THIS FIELD TO REPORT
'                                    DATA FOR TAX YEAR 1990 OR
'                                    EARLIER.  FOR 1991 AND LATER,
'                                    THIS FIELD MUST BE GREATER
'                                    THAN OR EQUAL TO THE SUM OF
'                                    THE FIELDS FOR ANNUAL SOCIAL
'                                    SECURITY WAGES AND ANNUAL
'                                    SOCIAL SECURITY TIPS.
'
 BLANK5 AS STRING * 1
'47       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'
 MEDTAX AS STRING * 10
'48-57    Medicare Tax         10    Enter the segment total.  No
'         Withheld                   negative amounts.  Right
'                                    justify and zero fill.  DO
'                                    NOT USE THIS FIELD TO REPORT
'                                    DATA FOR TAX YEAR 1990 OR
'                                    EARLIER.
'
 BLANK6 AS STRING * 71
'58-128   Blank                71    Enter blanks.  Reserved for
'                                    SSA use.
'
END TYPE
'
'Diskette Specifications:  Annual W-2 Copy A Information
'Date:  July 1994
'
'Record Name:  Codes 1T & 2T - Total Records
'
'Code 1T - Total Record
'Length=128
'
'Location   Field           Length   Description & Remarks
TYPE Rec1T
 IDENT AS STRING * 2
'1-2      Record Identifier    2     Constant "1T" (= numeric one
'                                    + letter T).
'
 NUMOFEMP AS STRING * 7
'3-9      Number of Employees  7     Enter the total number of
'                                    sets of Code 1W/2W/3W records
'                                    reported since the last
'                                    employer record (Code 1E/2E).
'                                    Right justify and zero fill.
'

 SSWTOTAL AS STRING * 13
'10-22    Annual Social        13    Enter the total for all
'         Security Wages             employee (Code 2W) records
'                                    reported since the last
'                                    employer record (Code 1E/2E).
'                                    Right justify and zero fill.
'

 BLANK1 AS STRING * 1
'23       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'

 SSTTOTAL AS STRING * 12
'24-35    Annual Social        12    Enter the total for all
'         Security Tips              employee (Code 2W) records
'                                    reported since the last
'                                    employer record (Code 1E/2E).
'                                    Right justify and zero fill.
'
 OTHERTOTAL AS STRING * 13
'36-48   Annual Wages, Tips    13    Enter the total for all
'        and Other                   employee (Code 2W) records
'        Compensation                reported since the last
'                                    employer record (Code 1E/2E).
'                                    Right justify and zero fill.
'
 BLANK2 AS STRING * 1
'49       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'
 SSTAXTOTAL AS STRING * 12
'50-61    Social Security      12    Enter the total for all
'         Employee Tax               employee (Code 2W) records
'         Withheld                   reported since the last
'                                    employer record (Code 1E/2E).
'                                    Right justify and zero fill.
'
 BLANK3 AS STRING * 1
'62       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'
 FEDTAXTOTAL AS STRING * 12
'63-74    Federal Income Tax   12    Enter the total for all
'         Withheld                   employee (Code 2W) records
'                                    reported since the last
'                                    employer record (Code 1E/2E).
'                                    Right justify and zero fill.
'
 ETLITOTAL AS STRING * 12
'75-86    Employer Cost of     12    Enter the total for all
'         Premiums for Group         employee (Code 2W) records
'         Term Life Insurance        reported since the last
'         Over $50,000               employer record (Code 1E/2E).
'                                    Right justify and zero fill.
'

 UNTAXTOTAL AS STRING * 12
'87-98    Uncollected Employee 12    Enter the total for all
'         Tax on Tips                employee (Code 2W) records
'                                    reported since the last
'                                    employer record (Code 1E/2E).
'                                    Right justify and zero fill.
'

 EICTOTAL AS STRING * 12
'99-110   Advance Earned       12    Enter the total for all
'         Income Credit              employee (Code 2W) records
'                                    reported since the last
'                                    employer record (Code 1E/2E).
'                                    Right justify and zero fill.
'
 ATIPTOTAL AS STRING * 12
'111-122  Allocated Tips       12    Enter the total for all
'                                    employee (Code 2W) records
'                                    reported since the last
'                                    employer record (Code 1E/2E).
'                                    Right justify and zero fill.
'
 BLANK4 AS STRING * 6
'123-128  Blank                6     Enter blanks.  Reserved for
'                                    SSA use.
'
'
END TYPE
'Code 2T - Total Record
'Length=128
'
'Location   Field           Length   Description & Remarks
TYPE Rec2T
 IDENT AS STRING * 2
'1-2      Record Identifier    2     Constant "2T" (= numeric two
'                                    + letter T).
'
 FRINGETOTAL AS STRING * 12
'3-14     Fringe Benefits      12    Enter the total for all
'                                    employee (Code 2W) records
'                                    reported since the last
'                                    employer record (Code 1E/2E).
'                                    Right justify and zero fill.
'
 BLANK1 AS STRING * 1
'15       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'
 DEFCOMPCTOTAL AS STRING * 13
'16-28    Deferred             13    Enter the total for all
'         Compensation               employee (Code 2W) records
'         Contributions              reported since the last
'                                    employer record (Code 1E/2E).
'                                    Right justify and zero fill.
'
 BLANK2 AS STRING * 1
'29       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'
 DEPCBENTOTAL AS STRING * 12
'30-41    Dependent Care       12    Enter the total for all
'         Benefits                   employee (Code 2W) records
'                                    reported since the last
'                                    employer record (Code 1E/2E).
'                                    Right justify and zero fill.
'
 BLANK4 AS STRING * 1
'42       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'
 NQPTOTAL AS STRING * 13
'43-55    Non-qualified Plan   13    Enter the total for all
'         Section 457                employee (Code 2W) records
'         Distributions              reported since the last
'                                    employer record (Code 1E/2E).
'                                    Right justify and zero fill.
'
 BLANK5 AS STRING * 1
'56       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'
 NQPOTOTAL AS STRING * 13
'57-69    Non-qualified Plan   13    Enter the total for all
'         Not Section 457            employee (Code 2W) records
'         Distributions              reported since the last
'                                    employer record (Code 1E/2E).
'                                    Right justify and zero fill.
'
 BLANK6 AS STRING * 1
'70       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'
 MEDWAGETOTAL AS STRING * 13
'71-83    Medicare Wages and   13    Enter the total for all
'         Tips                       employee (Code 3W) records
'                                    reported since the last
'                                    employer record (Code 1E/2E).
'                                    Right justify and zero fill.
'                                    DO NOT USE THIS FIELD TO
'                                    REPORT DATA FOR TAX YEAR
'                                    1990 OR EARLIER.  FOR 1991
'                                    AND LATER, THIS FIELD MUST
'                                    BE GREATER THAN OR EQUAL
'                                    TO THE SUM OF THE FIELDS
'                                    FOR ANNUAL SOCIAL SECURITY
'                                    WAGES AND ANNUAL SOCIAL
'                                    SECURITY TIPS.
'
 BLANK7 AS STRING * 1
'84       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'
 MEDTAXTOTAL AS STRING * 12
'85-96    Medicare Tax         12    Enter the total for all
'         Withheld                   employee (Code 3W) records
'                                    reported since the last
'                                    employer record (Code 1E/2E).
'                                    Right justify and zero fill.
'                                    DO NOT USE THIS FIELD TO
'                                    REPORT DATA FOR TAX YEAR
'                                    1990 OR EARLIER.
'
 BLANK8 AS STRING * 32
'97-128   Blank                32    Enter blanks.  Reserved for
'                                    SSA use.
'
END TYPE
'
'Code 1F - FINAL RECORD
'
'The Code 1F record indicates the end of the file.  It MUST be the
'last data record on the file, appearing only once, after the last
'set of Code 1T/2T records.  SSA does not process any data
'recorded after the Code 1F record.
'
'
'Diskette Specifications:  Annual W-2 Copy A Information
'Date:  July 1994
'
'Record Name:  Code 1F - Final Record
'Length=128
'
'Location   Field           Length   Description & Remarks
TYPE Rec1F
 IDENT AS STRING * 2
'1-2      Record Identifier    2     Constant "1F".
'
 NUMOFEMP AS STRING * 7
'3-9      Number of Employees  7     Enter the total number of
'                                    sets of employee (Code
'                                    1W/2W/3W) records reported on
'                                    the entire file.  Right
'                                    justify and zero fill.
'
 SSWGTOTAL AS STRING * 16
'10-25    File Total:  Annual  16    Enter the grand total of the
'         Social Security            amounts reported in positions
'         Wages                      3-9 of all Code 2W records on
'                                    this file.  Right justify and
'                                    zero fill.
'
 BLANK1 AS STRING * 1
'26       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'
 SSTGTOTAL AS STRING * 16
'27-42    File Total:  Annual  16    Enter the grand total of the
'         Social Security Tips       amounts reported in positions
'                                    11-17 of all Code 2W records
'                                    on this file.  Right justify
'                                    and zero fill.
'
 BLANK2 AS STRING * 1
'43       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'
 OTHERWGTOTAL AS STRING * 16
'44-59    File Total:  Annual  16    Enter the grand total of the
'         Wages, Tips and            amounts reported in positions
'         Other Compensation         19-27 of all Code 2W records
'                                    on this file.  Right justify
'                                    and zero fill.
'
 BLANK3 AS STRING * 1
'60       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'
 SSTAXGTOTAL AS STRING * 16
'61-76    File Total:  Social  16    Enter the grand total of the
'         Security Employee          amounts reported in positions
'         Tax Withheld               29-34 of all Code 2W records
'                                    on this file.  Right justify
'                                    and zero fill.
'
 BLANK4 AS STRING * 1
'77       Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
 FEDTAXGTOTAL AS STRING * 16
'78-93    File Total:  Federal 16    Enter the grand total of the
'         Income Tax Withheld        amounts reported in positions
'                                    35-43 of all Code 2W records
'                                    on this file.  Right justify
'                                    and zero fill.
'
 BLANK5 AS STRING * 1
'94         Blank              1     Enter a blank.  Reserved for
'                                    SSA use.
'
 EICGTOTAL AS STRING * 16
'95-110   File Total:          16    Enter the grand total of the
'         Advance Earned             amounts reported in positions
'         Income Credit              84-90 of all Code 2W records
'                                    on this file.  Right justify
'                                    and zero fill.
'
 BLANK6 AS STRING * 18
'111-128  Blank                18    Enter blanks.  Reserved for
'                                    SSA use.
'
'
'NOTE:  There is no Code 2F record.  File totals for Medicare
'       Wages and Tips and for Medicare Tax Withheld are NOT
'       reported on the Final Record.
'
END TYPE
