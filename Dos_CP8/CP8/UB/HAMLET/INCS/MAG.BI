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


'Code 2A - Transmitter Record
'Length=128
'
'Location   Field           Length   Description & Remarks
'

TYPE Rec2A

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

'Code 2B - Basic Information Record
'Length=128
'
'Location   Field           Length   Description & Remarks
'
TYPE Rec2B

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

'
'Code 2E - Employer Record
'Length=128
'
'Location   Field           Length   Description & Remarks
'
TYPE Rec2E

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

'Code 2W - Employee Wage Record
'Length=128
'
'Location   Field           Length   Description & Remarks
TYPE Rec2W

'
'Code 3W - Employee Wage Record
'Length=128
'
'Location   Field           Length   Description and Remarks
'
TYPE Rec3W

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
'
'Code 2I - Intermediate Total Record
'Length=128
'
'Location   Field           Length   Description & Remarks
TYPE Rec2I

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

'Code 2T - Total Record
'Length=128
'
'Location   Field           Length   Description & Remarks
TYPE Rec2T
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
