
'
'CODES 1S/2S - SUPPLEMENTAL RECORDS
'
'THE CODE 1S/2S RECORDS ARE NOT REQUIRED FOR AN ANNUAL W-2 COPY A
'REPORT.  SSA AND IRS DO NOT PROCESS THE INFORMATION CONTAINED ON
'THE CODE 1S/2S RECORDS.  DO NOT GENERATE CODES 1S/2S RECORDS IF
'ONLY BLANKS WOULD BE ENTERED IN POSITIONS 3-128 OF EACH RECORD.
'SEE SECTION II FOR FURTHER EXPLANATION.  CALL YOUR STATE
'DEPARTMENT OF REVENUE/TAXATION FOR STATE FILING INSTRUCTIONS.
'
'Diskette Specifications:  Annual W-2 Copy A Information
'Date:  July 1994
'
'Record Name:  Codes 1S & 2S - Supplemental Records (OPTIONAL)
'
'Code 1S - Supplemental Record
'Length=128
'
'Location   Field           Length   Description & Remarks
'
'1-2      Record Identifier    2     Constant "1S".
'
'3-128    Supplemental Data    126   To be defined by user.
'
'         OR FOR STATE/LOCAL QUARTERLY/ANNUAL REPORTING
'
'1-2      Record Identifier    2     Constant "1S".
'
'3-11     Social Security      9     Enter the employee's social
'         Number (SSN)               security number.  If not
'                                    available, enter the letter
'                                    "I" in position 3 and blanks
'                                    in positions 4-11.
'
'12-38    Employee Name        27    Enter employee's name.  Left
'                                    justify and fill with blanks.
'
'39-78    Street Address       40    Left justify and fill with
'                                    blanks.
'
'79-103   City                 25    Left justify and fill with
'                                    blanks.  For a foreign
'                                    address, include name of
'                                    foreign "state" and country;
'                                    abbreviate city and state as
'                                    necessary; show full country
'                                    name.
'
'104-113  State                10    Use a standard FIPS postal
'                                    abbreviation (Appendix B).
'                                    For a foreign address, enter
'                                    blanks.
'
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
'119-123  ZIP Code or Foreign  5     Enter a valid ZIP Code.  For
'         Postal Code                a foreign address, use this
'                                    field for the Foreign Postal
'                                    Code, if applicable; left
'                                    justify and fill with blanks.
'                                    If necessary, continue the
'                                    Foreign Postal Code in
'                                    positions 114-118 above.
'
'124      Blank                1     Enter a blank.  Reserved for
'                                    SSA use.
'
'125-126  State Code           2     Enter the appropriate FIPS
'                                    postal NUMERIC code (Appendix
'                                    B).
'
'127-128  Optional Code        2     To be defined by individual
'                                    states.
'
'
'Code 2S - Supplemental Record
'Length=128
'
'Location   Field           Length   Description & Remarks
'
'1-2      Record Identifier    2     Constant "2S".
'
'3-128    Supplemental Data    126   To be defined by user.
'
'         OR FOR STATE/LOCAL QUARTERLY/ANNUAL REPORTING
'
'1-2      Record Identifier    2     Constant "2S".
'
'3-14     State Employer       12    Left justify and fill with
'         Account Number             blanks.  See Glossary.
'
'15-18    Reporting Period     4     Enter the last month and year
'                                    for the calendar quarter for
'                                    which this report applies;
'                                    e.g., "0394" for
'                                    January-March of 1994.
'
'19-27    State Quarterly      9     Right justify and zero fill.
'         Unemployment
'         Insurance Total
'         Wages
'
'28-36    State Quarterly      9     Right justify and zero fill.
'         Unemployment
'         Insurance Total
'         Taxable Wages
'
'37-38    Number of Weeks      2     To be defined by user.
'         Worked
'
'39-42    Date First Employed  4     Enter the month and year,
'                                    e.g., "0194."
'
'43-46    Date of Separation   4     Enter the month and year,
'                                    e.g., "0194."
'
'47-51    Taxing Entity Code   5     To be defined by user.
'
'52-53    State Code           2     Enter the appropriate FIPS
'                                    postal NUMERIC code (Appendix
'                                    B).
'
'54-62    State Taxable Wages  9     Right justify and zero fill.
'
'63-70    State Income Tax     8     Right justify and zero fill.
'         Withheld
'
'71-80    Other State Data     10    To be defined by individual
'                                    taxing agencies.
'
'81       Tax Type Code        1     Enter the appropriate code
'                                    for entries in positions
'                                    87-95 and 96-102.
'                                    C - City Income Tax
'                                    D - County Income Tax
'                                    E - School District Income
'                                        Tax
'                                    F - Other Income Tax
'
'82-86    Taxing Entity Code   5     To be defined by individual
'                                    taxing agencies.
'
'87-95    Local Taxable Wages  9     To be defined by individual
'                                    taxing agencies.
'
'96-102   Local Income Tax     7     To be defined by individual
'         Withheld                   taxing agencies.
'
'103-109  State Control        7     Optional.
'         Number
'
'110-128  Blank                19    Enter blanks OR for employer
'                                    use.
'
