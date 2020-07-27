
    PRINT #RptHandle, ""
    PRINT #RptHandle, TAB(26); "TOWN OF STEPHENS CITY"
    PRINT #RptHandle, TAB(24); "BUSINESS LICENSE APPLICATION"
    PRINT #RptHandle, TAB(31); "For Year: "; Form$(2, 0)
    PRINT #RptHandle, ""
    PRINT #RptHandle, ""
    PRINT #RptHandle, TAB(5); "Please print or type:"
    PRINT #RptHandle, ""
    PRINT #RptHandle, TAB(5); "Applicant Name: "; ARCustRec(1).BILLNAME; TAB(55); "Phone:"
    PRINT #RptHandle, TAB(5); "               ----------------------------------------------------------"
    PRINT #RptHandle, ""
    PRINT #RptHandle, TAB(5); "Trade Name: "; TAB(54); "FEIN or SS#"
    PRINT #RptHandle, TAB(5); "           --------------------------------------------------------------"
    PRINT #RptHandle, ""
    PRINT #RptHandle, TAB(5); "Mailing Address:                          Physical Address:"
    PRINT #RptHandle, TAB(5); "                --------------------------  -----------------------------"
    PRINT #RptHandle, TAB(5); "------------------------------------------  -----------------------------"
    PRINT #RptHandle, TAB(5); "------------------------------------------  -----------------------------"
    PRINT #RptHandle, ""
    PRINT #RptHandle, TAB(5); "Phone:                                  Phone:"
    PRINT #RptHandle, TAB(5); "      -----------------------------           ---------------------------"
    PRINT #RptHandle, ""
    PRINT #RptHandle, TAB(5); "Nature Of Business:"
    PRINT #RptHandle, TAB(5); "                   ------------------------------------------------------"
    PRINT #RptHandle, ""
    PRINT #RptHandle, ""
    PRINT #RptHandle, TAB(5); "Gross reciepts                 Estimated                         Actual"
    PRINT #RptHandle, TAB(5); "for year ending"
    PRINT #RptHandle, TAB(5); "Dec. 31, _____                -----------                     -----------"
    PRINT #RptHandle, TAB(5); "(wholesalers only- enter purchases)"
    PRINT #RptHandle, ""
    PRINT #RptHandle, ""
    PRINT #RptHandle, TAB(5); "CONTRACTORS ONLY"
    PRINT #RptHandle, TAB(5); "     Please Note: All contractors must have valid Workmans Compensation"
    PRINT #RptHandle, TAB(5); "coverage in effect for the time period covered by this license. Failure"
    PRINT #RptHandle, TAB(5); "to have proper coverage will cause your license to be revoked."
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(5); "____ I certify that I am in compliance with the provisions of the Virginia"
    PRINT #RptHandle, TAB(5); "Workmans Compensation Act, and I will notify the Town of Stephens City if"
    PRINT #RptHandle, TAB(5); "this coverage lapses during the period that this license is in effect."
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(5); "I hereby swear (or affirm) that the statements are true, full and correct"
    PRINT #RptHandle, TAB(5); "to the best of my knowledge."
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(5); "___________________________________________              ________________"
    PRINT #RptHandle, "                       signature                                       date"
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(5); "*************************************************************************"
    PRINT #RptHandle, TAB(5); "FOR OFFICE USE ONLY"
    PRINT #RptHandle, TAB(5); "Zoning classification approved for this type of business"
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(5); "Approved by ______________________________              ________________"
    PRINT #RptHandle, "                              signature                               date"

    PRINT #RptHandle, CHR$(12);
    TotalCust = TotalCust + 1

