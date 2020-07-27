TYPE UBDraftRecType
    BANKDEST AS STRING * 9
    BANKORIG AS STRING * 9
    BANKNAME AS STRING * 23
    BANKLOC  AS STRING * 23
    FEDID    AS STRING * 9
    FEDPREFX AS STRING * 1
END TYPE

TYPE UBDraftPayRecType
    CustAcctNum   AS LONG
    DraftAmt      AS DOUBLE
END TYPE

TYPE UBDraftRecord1Type
    Field1 AS STRING * 1        ' Record Type Code Must = 1
    Field2 AS STRING * 2        ' Priority Code Must = 01
    Field3 AS STRING * 10       ' Immediate Destination Bank Transit Number (BB&T=b053101121 where b=blank space) Right Justified
    Field4 AS STRING * 10       ' Immediate Origin Bank Transit Number Must be Right Justified
    Field5 AS STRING * 6        ' Transmission File Creation Date (yymmdd)
    Field6 AS STRING * 4        ' Transmission File Creation Time (hhmm)
    Field7 AS STRING * 1        ' Field Modifier ID Must = A
    Field8 AS STRING * 3        ' Record Size  Must = 094
    Field9 AS STRING * 2        ' Blocking Factor  Must = 10
    Field10 AS STRING * 1       ' Format Code  Must = 1
    Field11 AS STRING * 23      ' Name of Destination Bank (Receiving Transmission)
    Field12 AS STRING * 23      ' Name of Originating Bank
    Field13 AS STRING * 8       ' Reserved Needs to be 8 blanks padded
END TYPE

TYPE UBDraftRecord5Type
    Field1 AS STRING * 1        ' Record Type Code Must = 5
    Field2 AS STRING * 3        ' Service Class Code Must = 200
    Field3 AS STRING * 16       ' Company Submitting Name in ALL CAPS
    Field4 AS STRING * 20       ' Discretionary Data
    Field5 AS STRING * 10       ' Company ID (Federal Tax ID Number)
    Field6 AS STRING * 3        ' Standard Entry Class (PPD for Direct Deposits and Drafts)
    Field7 AS STRING * 10       ' Company Entry Description SUCH AS UTIL BILL
    Field8 AS STRING * 6        ' Company Desc Date YYMMDD
    Field9 AS STRING * 6        ' Effective Entry Date YYMMDD
    Field10 AS STRING * 3       ' RESERVED LEAVE WITH 3 BLANKS
    Field11 AS STRING * 1       ' Must Equal 1 for Originator Status Code
    Field12 AS STRING * 8       ' Originating Fin. Inst. ID  05310112 for BB&T
    Field13 AS STRING * 7       ' Batch Number Beginning with 0000001
  END TYPE

TYPE UBDraftRecord6Type
    Field1 AS STRING * 1        ' Record Type Code  Must = 6
    Field2 AS STRING * 2        ' Transaction Code      22 Credit Checking
                                '                       27 Debit Checking
                                '                       32 Credit Savings
                                '                       37 Debit Savings
                                '                       28 PRENOTE DEBIT CHECKING  (Amt Must be all zeros
                                '                       Generally Add 1 to Get PreNote Transaction Code
    Field3 AS STRING * 8        ' Individual's Bank ID Transit Routing #
    Field4 AS STRING * 1        ' Transit Routing Check Digit
    Field5 AS STRING * 17       ' Individual's Bank Account Number
    Field6 AS STRING * 10       ' Amount (assume 2 decimal places)
    Field7 AS STRING * 15       ' Individual's ID Number to Customer (Usually Customer Utility Account Number)
    Field8 AS STRING * 22       ' Individual's Name
    Field9 AS STRING * 2        ' Set to 2 spaces (Not Used)
    Field10 AS STRING * 1       ' Set to '0' to signify no addenda records
    Field11 AS STRING * 15      ' Trace Number
                                ' Consists of Bank ID # 05310112 plus
                                ' Line Item Number starting w/ 0000001
                                ' and incrementing once for each line (Record6)
END TYPE


TYPE UBDraftRecord8Type
    Field1 AS STRING * 1        ' Record Code Must = 8
    Field2 AS STRING * 3        ' Service Class Code  Must = 200
    Field3 AS STRING * 6        ' Number of Detail (TYPE 6) Records
    Field4 AS STRING * 10       ' Hash Total
                                ' Hash#=Hash#+val(banktransit#) for Each Type 6 Record
    Field5 AS STRING * 12       ' Total Debit Amount
    Field6 AS STRING * 12       ' Total Credit Amount
    Field7 AS STRING * 10       ' Federal ID Tax Number
    Field8 AS STRING * 19       ' RESERVED KEEP BLANK
    Field9 AS STRING * 6        ' RESERVED BY FEDERAL RESERVE BANK
    Field10 AS STRING * 8      ' Originating Financial Inst. 05310112 for BB&T
    Field11 AS STRING * 7       ' Batch # Beginning with 0000001
END TYPE

TYPE UBDraftRecord9Type
    Field1 AS STRING * 1        ' Record Code Must=9
    Field2 AS STRING * 6        ' Batch Count (Sum of Batches) NORMALLY 000001
    Field3 AS STRING * 6        ' Block Count Number of Records
                                ' Found by taking Total Size of File and Dividing ty 940
    Field4 AS STRING * 8        ' Total 06 Record Type Entries
    Field5 AS STRING * 10       ' Enter Hash  See Above
    Field6 AS STRING * 12       ' Total Debit Entry Dollar Amount (Assume 2 decimal)
    Field7 AS STRING * 12       ' Total Credit Entry Dollar Amount (Assume 2 decimal)
    Field8 AS STRING * 39       ' RESERVED FOR FUTURE USE
END TYPE




