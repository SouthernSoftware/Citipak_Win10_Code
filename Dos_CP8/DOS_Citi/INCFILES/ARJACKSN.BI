   PRINT #RptHandle, ""
   PRINT #RptHandle, TAB(30); "LICENSE APPLICATION"
   PRINT #RptHandle, ""
   PRINT #RptHandle, ""
   PRINT #RptHandle, TAB(5); "Town of Jackson"; USING "              ACCOUNT NO_.####"; ARCustIdxRec(1).IDXRECORD;
   PRINT #RptHandle, TAB(58); "START DATE__________"
   PRINT #RptHandle, ""
   PRINT #RptHandle, TAB(5); "   Applicant's Name: "; ARCustRec(1).BILLNAME
   PRINT #RptHandle, TAB(5); "Applicant's Address: "; ARCustRec(1).Address1
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(5); "TAX MAP______  BLOCK______ LOT______     ZONING DISTRICT_________________"
   PRINT #RptHandle, TAB(5); "FEDERAL ID/SS NUMBER__________________   SC RETAIL TAX NO________________"
   PRINT #RptHandle, TAB(5); "TYPE OF BUSINESS:________________________________________________________"
   PRINT #RptHandle, TAB(5); "APPLICATION FOR:___ NEW___ RENEWAL___ GOING OUT OF BUSINESS(DATE)________"
   PRINT #RptHandle, TAB(5); "OWNERSHIP:___ CORPORATION___ PARTNERSHIP____ INDIVIDUAL-NO EMPLOYEES_____"
   PRINT #RptHandle, TAB(5); "NAME OF OWNER, PARTNER OR PRINCIPAL______________________________________"
   PRINT #RptHandle, TAB(5); "TELEPHONE NO. LOCAL:_____________ HOME_____________ EMERGENCY____________"
   PRINT #RptHandle, TAB(5); "FAX NO._____________  E-MAIL_____________________________________________"
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(5); "IS HAZARDOUS WASTE INVOLVED IN OPERATION? ____NO ____YES (ATTACH DETAILS)"
   PRINT #RptHandle, TAB(5); "CODE CLEARANCE: __ZONING ___INSPECTION __FIRE __HEALTH ___LAW ENFORCEMENT"
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(28); "COMPUTATION OF LICENSE TAX"
   PRINT #RptHandle, TAB(5); "COMPUTE LICENSE TAX ACCORDING TO THE FOLLOWING SCHEDULE AND MAKE CHECKS"
   PRINT #RptHandle, TAB(5); "PAYABLE TO:  'TOWN OF JACKSON'   DELIVER BY DUE DATE ____________________"
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(5); "GROSS INCOME FOR PRECEDING CALENDAR OR FISCAL YEAR     $_________________"
   PRINT #RptHandle, TAB(5); "LESS INCOME ON WHICH A LICENSE TAX WAS PAID TO ANOTHER"
   PRINT #RptHandle, TAB(5); "CITY OR COUNTY FOR OPERATIONS OUTSIDE CITY/COUNTY      $_________________"
   PRINT #RptHandle, TAB(5); "BALANCE OF GROSS INCOME SUBJECT TO LICENSE TAX         $_________________"
   PRINT #RptHandle, TAB(5); "TAX:   RATE CLASS MINIMUM ON FIRST $5,000 ____________PLUS"
   PRINT #RptHandle, TAB(5); "$_______PER $1,000 FOR INCOME OVER $5,000 ____________"
   PRINT #RptHandle, TAB(5); "[See declining rate schedule for over $1 million]       [OFFICE USE ONLY]"
   PRINT #RptHandle, TAB(5); "                           TOTAL LICENSE TAX $_________ [PAYMENT RECORD]"
   PRINT #RptHandle, TAB(5); "PENALTY AFTER DUE DATE __% PER MONTH $_________ [CHECK NO. ____________]"
   PRINT #RptHandle, TAB(5); "TOTAL LICENSE TAX AND PENALTY $_________     [DATE RECEIVED____________]"
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(35); "CERTIFICATION"
   PRINT #RptHandle, TAB(5); "I (WE) DO CERTIFY THAT THE ABOVE INFORMATION AND AMOUNT RETURNED AS GROSS"
   PRINT #RptHandle, TAB(5); "INCOME FROM MY BUSINESS IS TRUE AND CORRECT. AND I HAVE MADE NO DEDUCTIONS"
   PRINT #RptHandle, TAB(5); "EXCEPT INCOME ON WHICH I HAVE PAID BUSINESS LICENSE TAX TO ANOTHER CITY OR"
   PRINT #RptHandle, TAB(5); "COUNTY, FOR WHICH I HAVE PROOF OF PAYMENT. I AM FAMILIAR WITH THE PENALTY"
   PRINT #RptHandle, TAB(5); "PROVISIONS OF THE ORDINANCE AND GROUNDS FOR LICENSE REVOCATION, INCLUDING"
   PRINT #RptHandle, TAB(5); "MAKING FALSE OR FRAUDULENT STATEMENTS IN THIS APPLICATION. I CERTIFY THAT"
   PRINT #RptHandle, TAB(5); "ALL BUSINESS PERSONAL PROPERTY TAXES AND PAYABLES DUE TO THE CITY/COUNTY"
   PRINT #RptHandle, TAB(5); "HAVE BEEN PAID, AND THAT THE ABOVE BUSINESS NAME IS THE SAME AS REPORTED"
   PRINT #RptHandle, TAB(5); "ON DOCUMENTS FILED WITH THE STATE AND FEDERAL GOVERNMENTS. I UNDERSTAND MY"
   PRINT #RptHandle, TAB(5); "BUSINESS INCOME TAX RETURNS AND OTHER DOCUMENTS MAY BE INSPECTED TO VERIFY"
   PRINT #RptHandle, TAB(5); "GROSS INCOME OR OTHER BUSINESS DATA."
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(5); "___________________________________________________________________________"
   PRINT #RptHandle, TAB(5); "SIGNATURE                          TITLE                           DATE"
   PRINT #RptHandle, CHR$(12);
   PrnCnt& = PrnCnt& + 1

