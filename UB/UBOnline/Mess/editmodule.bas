Attribute VB_Name = "Module2"
Public Sub editadministrative(editerr, typeedit As Integer)
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwi + "incident.mdb")
'If TOBOOK = 0 And (arrestedunder18 = 1 Or arrested18andover = 1) Then
'    Set rs = db.OpenRecordset("select incidentnumber from booking where INCIDENTnumber = " + Chr$(34) + Incidentnumber + Chr$(34))
'    If rs.EOF Then
'        On Error Resume Next
'        msg = MsgBox("A valid arrest number must be entered on Booking Report.  If Booking Report has not yet been received, clear all Arrest fields until such time as paperwork is received.", 48, "Genesis Error Log")
'        GoTo exitedita
'    End If
'End If
On Error Resume Next

'RLB Code
    Dim dateBegin As Date, dateEnd As Date
    
    
    dateBegin = DateValue(incidentdate(0)) + TimeValue(TIMEOFOFFENSE(0))
    dateEnd = DateValue(incidentdate(1)) + TimeValue(TIMEOFOFFENSE(1))
    
    
    If dateBegin > dateEnd Then
        MsgBox "The date and time of the beginning of the offense, must take place before the date and time of the end of the offense.", 48, "Genesis Error Log"
        incidentdate(0).SetFocus
        GoTo exitedita
    End If
'********

'===== Data Element 5
'===== Error 153
If IsDate(EXCEPTIONALCLEARANCEDATE) And Not (offenderdeath Or noprosecution Or extraditiondenied Or victimdeclinescooperation Or juvenilenocustody) Then
    msg = MsgBox("The combination of an entered Exceptional Clearance Date and N/A is invalid.", 48, "Genesis Error Log")
    Call ShowApplicableContainers(EXCEPTIONALCLEARANCEDATE)
    EXCEPTIONALCLEARANCEDATE.SetFocus
    GoTo exitedita
End If
'===== Error 156
If Not IsDate(EXCEPTIONALCLEARANCEDATE) And (offenderdeath Or noprosecution Or extraditiondenied Or victimdeclinescooperation Or juvenilenocustody) Then
    msg = MsgBox("An Exceptional Clearance Date must be entered for this reason.", 48, "Genesis Error Log")
    Call ShowApplicableContainers(EXCEPTIONALCLEARANCEDATE)
    EXCEPTIONALCLEARANCEDATE.SetFocus
    GoTo exitedita
End If
'===== Error 155
If IsDate(EXCEPTIONALCLEARANCEDATE) Then
    If CVDate(EXCEPTIONALCLEARANCEDATE) < CVDate(incidentdate(0)) Then
        msg = MsgBox("Exceptional Clearance Date cannot be earlier than Incident Date.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(EXCEPTIONALCLEARANCEDATE)
        EXCEPTIONALCLEARANCEDATE.SetFocus
        GoTo exitedita
    End If
End If
'==== Mandatories E - 4 = GIVEN
'==== Mandatories E - 5
If (exclearunder18 = 1 Or exclearover18 = 1) And Not IsDate(EXCEPTIONALCLEARANCEDATE) Then
    msg = MsgBox("Exceptional Clearance Date is not valid.", 48, "Genesis Error Log")
    Call ShowApplicableContainers(exclearunder18)
    exclearunder18.SetFocus
    GoTo exitedita
End If
'===== Error 105
If IsDate(EXCEPTIONALCLEARANCEDATE) Then
    EXCEPTIONALCLEARANCEDATE = Format$(EXCEPTIONALCLEARANCEDATE, "mm/dd/yyyy")
End If
'===== Data Element 4
If Not (offenderdeath Or noprosecution Or extraditiondenied Or victimdeclinescooperation Or juvenilenocustody) And IsDate(EXCEPTIONALCLEARANCEDATE) Then
    msg = MsgBox("A type of Exceptional Clearance other than NA must be chosen if date entered.", 48, "Genesis Error Log")
    Call ShowApplicableContainers(offenderdeath)
    offenderdeath.SetFocus
    GoTo exitedita
End If
'===== Data Element 5
If offenderdeath Or noprosecution Or extraditiondenied Or victimdeclines Or juvenilenocustody Then
    If Not IsDate(EXCEPTIONALCLEARANCEDATE) Then
        msg = MsgBox("An exceptional clearance date must be entered if exceptionally cleared.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(EXCEPTIONALCLEARANCEDATE)
        EXCEPTIONALCLEARANCEDATE.SetFocus
        GoTo exitedita
    End If
End If
If offenderdeath Or noprosecution Or extraditiondenied Or victimdeclines Or juvenilenocustody Then
    Set rs = db.OpenRecordset("select incidentnumber from booking where INCIDENTnumber = " + Chr$(34) + incidentnumber + Chr$(34))
    If Not rs.EOF Then
        msg = MsgBox("No arrest data allowed for an exceptional clearance.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(EXCEPTIONALCLEARANCEDATE)
        EXCEPTIONALCLEARANCEDATE.SetFocus
        GoTo exitedita
    End If
End If
'===== Mandatories E - 2
If incidentnumber = "" Or Len(incidentnumber) > 12 Then
    msg = MsgBox("Incident number must be entered and be 12 or less characters long.", 48, "Genesis Error Log")
    Call ShowApplicableContainers(incidentnumber)
    incidentnumber.SetFocus
    GoTo exitedita
End If
'===== Error 172
If CVDate(incidentdate(0)) < CVDate("1/1/1991") And onpaper = 0 Then
    msg = MsgBox("Incident date is not valid.", 48, "Genesis Error Log")
    Call ShowApplicableContainers(incidentdate(0))
    incidentdate(0).SetFocus
    GoTo exitedita
End If
'===== Error 105
incidentdate(0) = Format$(incidentdate(0), "mm/dd/yyyy")
'===== Mandatories E - 3
'===== Data Element 3
'===== Error 151
If Not IsDate(TIMEOFOFFENSE(0)) Then
    msg = MsgBox("Time of Offense is not valid.", 48, "Genesis Error Log")
    Call ShowApplicableContainers(TIMEOFOFFENSE(0))
    TIMEOFOFFENSE(0).SetFocus
    GoTo exitedita
End If
'===== SC Enhancements - Administrative
If Not IsDate(TIMEOFOFFENSE(1)) Then
    msg = MsgBox("Time of Offense is not valid.", 48, "Genesis Error Log")
    Call ShowApplicableContainers(TIMEOFOFFENSE(1))
    TIMEOFOFFENSE(1).SetFocus
    GoTo exitedita
End If
If IsDate(REPORTINGOFFICERDATE(0)) Then
    REPORTINGOFFICERDATE(0) = Format$(REPORTINGOFFICERDATE(0), "mm/dd/yyyy")
End If
If reportingofficer(1) > "" Then
    If Not IsDate(REPORTINGOFFICERDATE(1)) Then
        REPORTINGOFFICERDATE(1) = REPORTINGOFFICERDATE(0)
    End If
    REPORTINGOFFICERDATE(1) = Format$(REPORTINGOFFICERDATE(1), "mm/dd/yyyy")
End If
If APPROVINGOFFICERDATE > "" And Not IsDate(APPROVINGOFFICERDATE) Then
    msg = MsgBox("Approving Date is not valid.", 48, "Genesis Error Log")
    Call ShowApplicableContainers(APPROVINGOFFICERDATE)
    APPROVINGOFFICERDATE.SetFocus
    GoTo exitedita
End If
'===== SCEnhancements - Administrative
If active = 0 And admclosed = 0 And unfounded = 0 Then
    msg = MsgBox("Either Active, Adm Closed, or Unfounded must be selected.", 48, "Genesis Error Log")
    Call ShowApplicableContainers(active)
    active.SetFocus
    GoTo exitedita
End If
'==== Mandatories E - 8A
If BIAS.ListIndex = -1 Then
    msg = MsgBox("Bias Motivation must be selected.", 48, "Genesis Error Log")
    Call ShowApplicableContainers(BIAS)
    BIAS.SetFocus
    GoTo exitedita
End If
'If typeedit = 1 Then
    Call ShowApplicableContainers(BIAS)
    BIAS.SetFocus
    GoTo goodedita
'End If
Exit Sub
exitedita:
editerr = 1
goodedita:
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub
Public Sub editevent(editerr, typeedit As Integer)
Dim testgroup, totgroup As String, temperr As Integer, tempucr, tempgroup, typeselect As String, tempvalue As Single, tempdate As String
'RLB Bandaid
On Error GoTo rlbErr
Screen.MousePointer = 11
'===== Error 101
If Not IsDate(incidentdate(0)) Then
    msg = MsgBox("Beginning incident date is not valid.", 48, "Genesis Error Log")
    Call ShowApplicableContainers(incidentdate(0))
    incidentdate(0).SetFocus
    GoTo exitedite
End If
'===== SC Enhancements - Administrative
If Not IsDate(incidentdate(1)) Then
    msg = MsgBox("Ending incident date is not valid.", 48, "Genesis Error Log")
    Call ShowApplicableContainers(incidentdate(1))
    incidentdate(1).SetFocus
    GoTo exitedite
End If
'===== Error 016, 116, 216, 316, 416, 516, 616, 716
If Len(incidentnumber) < 12 And Len(incidentnumber) > 0 Then
    incidentnumber = incidentnumber + Space$(12 - Len(incidentnumber))
End If
If age(0) > "" Then
    If Val(age(0)) = 0 And age(0) <> "00" Then
        msg = MsgBox("Age must be entered in the format of a single age: X or XX (i.e. 3 or 42) or in the format of a range of ages: XXXX (i.e. 2025, meaning 20 - 25 years old).", 48, "Genesis Error Log")
        Call ShowApplicableContainers(age(0))
        age(0).SetFocus
        GoTo exitedite
    End If
End If
For t% = 1 To Len(age(0))
    If InStr("0123456789-", Mid$(age(0), t%, 1)) = 0 Then
        msg = MsgBox("An invalid entry in age has been found.  Valid Entry Formats:  X, XX, X-Y, XX-Y, X-YY, XX-YY, XXYY", 48, "Genesis Error Log")
        t% = Len(age(0))
        Call ShowApplicableContainers(age(0))
        age(0).SetFocus
        GoTo exitedite
    End If
Next t%
If InStr(age(0), "-") > 0 Then
    ag1 = Val(Left$(age(0), InStr(age(0), "-") - 1))
    ag2 = Val(Mid$(age(0), InStr(age(0), "-") + 1))
    age(0) = Format$(ag1, "00") + Format$(ag2, "00")
End If
editeventdetail:
'===== SC Edit - Justifiable Homocide (09C) must be on separate casenumber =====
fo% = 0
fj% = 0
nolarceny = 0
tempucr = ""
ucrct% = 0
found35a = 0
ucrs$ = ""
foundweapon = 0
For t% = 0 To 9
    For tt% = 0 To ucrlist(t%).ListCount - 1
        If ucrlist(t%).Selected(tt%) = True Then
            ucrct% = ucrct% + 1
            tempucr = Mid$(ucrlist(t%).List(tt%), InStr(ucrlist(t%).List(tt%), "(") + 1, 3)
            For p% = 1 To Len(ucrs$) Step 3
                If Mid$(ucrs$, p%, 3) = tempucr Then
                    msg = MsgBox("A duplicate UCR reference for " + tempucr + " has been found.  Do not enter separate offense lines for the same UCR.", 48, "Genesis Error Log")
                    Call ShowApplicableContainers(ucrlist(t%))
                    ucrlist(t%).SetFocus
                     
                    GoTo exitedite
                End If
            Next p%
            ucrs$ = ucrs$ + tempucr
            Select Case tempucr
                Case "09A", "09B", "09C", "100", "11A", "11B", "11C", "11D", "120", "13A", "13B", "210", "520"
                    foundweapon = 1
                Case "09C"
                    fj% = 1
                    nolarceny = 1
                Case "23A", "23B", "23C", "23D", "23E", "23F", "23G", "23H"
                    fo% = 1
                Case "35A"
                    found35a = 1
                    fo% = 1
                    nolarceny = 1
                Case Else
                    nolarceny = 1
                    fo% = 1
            End Select
            tt% = ucrlist(t%).ListCount - 1
        End If
    Next tt%
Next t%
'===== Error 201
'===== SCEdit 4/21/92 P28
If tempucr = "" Then
    msg = MsgBox("A valid UCR code must be selected.", 48, "Genesis Error Log")
    Call ShowApplicableContainers(ucrlist(0))
    ucrlist(0).SetFocus
    GoTo exitedite
End If
If fo% = 1 And fj% = 1 Then
    msg = MsgBox("Justifiable Homocide must be submitted on a separate case number.", 48, "Genesis Error Log")
    Call ShowApplicableContainers(ucrlist(0))
    ucrlist(0).SetFocus
    GoTo exitedite
End If
'===== SC LEOKA
FSP% = 0
For t% = 0 To 9
    For tt% = 1 To sublist(t%).ListItems.Count
        If sublist(t%).ListItems(tt%).Selected Then
            If Left$(sublist(t%).ListItems(tt%), 1) = "P" Then
                FSP% = 1
                t% = 9
            End If
        End If
    Next tt%
Next t%
If lactivity.ListIndex = -1 Then
    If policeofficer Then
        msg = MsgBox("A LEOKA activity must be selected for Type Victim Police Officer.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(policeofficer)
        policeofficer.SetFocus
        GoTo exitedite
    End If
    If FSP% = 1 Then
        msg = MsgBox("A LEOKA activity must be selected for Type Victim Police Officer.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(sublist(0))
        sublist(0).SetFocus
        GoTo exitedite
    End If
Else
    If Not policeofficer Then
        msg = MsgBox("If an Activity of selected for LEOKA, the type victim must be police officer.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(policeofficer)
        policeofficer.SetFocus
        GoTo exitedite
    End If
End If
If FSP% = 0 And policeofficer Then
    msg = MsgBox("Type Victim is Police Officer, but Police Officer is not indicated in Subcodes.", 48, "Genesis Error Log")
    Call ShowApplicableContainers(sublist(0))
    sublist(0).SetFocus
    GoTo exitedite
End If
If FSP% = 1 And Not policeofficer Then
    msg = MsgBox("Type Victim must be Police Officer if indicated in Subcodes.", 48, "Genesis Error Log")
    Call ShowApplicableContainers(policeofficer)
    policeofficer.SetFocus
    GoTo exitedite
End If
For t% = 0 To 9
    '===== Mandatories E - 6, 24
    '===== Data Element 6
    ucrlist(t%).ListIndex = -1
    For tt% = 0 To ucrlist(t%).ListCount - 1
        If ucrlist(t%).Selected(tt%) = True Then
            ucrlist(t%).ListIndex = tt%
            tt% = ucrlist(t%).ListCount - 1
        End If
    Next tt%
    If pickoffense(t%).ListIndex <> -1 And ucrlist(t%).ListIndex = -1 Then
        msg = MsgBox("A valid UCR code must be selected.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(ucrlist(t%))
        ucrlist(t%).SetFocus
        GoTo exitedite
    End If
    If ucrlist(t%).ListIndex > -1 Then
        tempucr = Mid$(ucrlist(t%).List(ucrlist(t%).ListIndex), InStr(ucrlist(t%).List(ucrlist(t%).ListIndex), "(") + 1, 3)
        '===== SC LEOKA
        If FSP% = 1 Or policeofficer Then
            For ii% = 1 To HOMOCIDE(t%).ListItems.Count
                If HOMOCIDE(t%).ListItems(ii%).Selected Then
                    If Mid$(HOMOCIDE(t%).ListItems(ii%), InStr(HOMOCIDE(t%).ListItems(ii%), "(") + 1, 2) <> "02" Then
                        msg = MsgBox("For LEOKA, only homocide/aggravated assault circumstance of 02 is allowed.", 48, "Genesis Error Log")
                        Call ShowApplicableContainers(HOMOCIDE(t%))
                        HOMOCIDE(t%).SetFocus
                        GoTo exitedite
                    End If
                End If
            Next ii%
        End If
        '===== Error 469
        If tempucr = "11A" Or tempucr = "36B" Then
            If sex(1).ListIndex = -1 Or Left$(sex(1).List(sex(1).ListIndex), 1) <> "M" And Left$(sex(1).List(sex(1).ListIndex), 1) <> "F" Then
                msg = MsgBox("For Forcible Rape, a victim sex of M or F are allowed.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(sex(1))
                sex(1).SetFocus
                GoTo exitedite
            End If
        End If
        If tempucr = "250" Or tempucr = "280" Or tempucr = "35A" Or tempucr = "35B" Or tempucr = "39C" Or tempucr = "370" Or tempucr = "520" Or tempucr = "09A" Or tempucr = "09B" Or tempucr = "100" Or tempucr = "120" Or tempucr = "11A" Or tempucr = "11B" Or tempucr = "11C" Or tempucr = "11D" Or tempucr = "13A" Or tempucr = "13B" Or tempucr = "13C" Then
            If tempucr = "09A" Or tempucr = "09B" Or tempucr = "100" Or tempucr = "120" Or tempucr = "11A" Or tempucr = "11B" Or tempucr = "11C" Or tempucr = "11D" Or tempucr = "13A" Or tempucr = "13B" Or tempucr = "13C" Then
                foundactivity = 0
                For i% = 1 To gactivity(t%).ListItems.Count
                    If gactivity(t%).ListItems(i%).Selected = True Then
                        gactivity(t%).ListItems(i%).EnsureVisible
                        foundactivity = 1
                        i% = gactivity(t%).ListItems.Count
                    End If
                Next i%
                If foundactivity = 0 Then
                    msg = MsgBox("A valid activity type must be selected for UCR " + tempucr + ".", 48, "Genesis Error Log")
                    Call ShowApplicableContainers(gactivity(t%))
                    gactivity(t%).SetFocus
                    GoTo exitedite
                End If
            Else
                '===== Error 201
                foundactivity = 0
                For i% = 1 To activity(t%).ListItems.Count
                    If activity(t%).ListItems(i%).Selected = True Then
                        activity(t%).ListItems(i%).EnsureVisible
                        foundactivity = 1
                        i% = activity(t%).ListItems.Count
                    End If
                Next i%
                If foundactivity = 0 Then
                    msg = MsgBox("A valid activity type must be selected for UCR " + tempucr + ".", 48, "Genesis Error Log")
                    Call ShowApplicableContainers(activity(t%))
                    activity(t%).SetFocus
                    GoTo exitedite
                End If
            End If
        End If
        foundweapon = 0
        foundweapon99 = 0
        idx5 = 0
        For i% = 1 To weapontype.ListItems.Count
            If weapontype.ListItems(i%).Selected = True Then
                weapontype.ListItems(i%).EnsureVisible
                idx5 = idx5 + 1
                If idx5 > 3 Then
                    i% = weapontype.ListItems.Count
                Else
                    foundweapon = foundweapon + 1
                    If InStr(weapontype.ListItems(i%), "(99)") > 0 Then
                        foundweapon99 = 1
                    End If
                    '----edit by Ed Sloan 11/30/99 from SCLED 3/5/97 update----------------
                    '===== Error 265, 269
                    If tempucr = "13B" Then
                        If (Mid$(weapontype.ListItems(i%), InStr(weapontype.ListItems(i%), "(") + 1, 2) <> "40" And Mid$(weapontype.ListItems(i%), InStr(weapontype.ListItems(i%), "(") + 1, 2) <> "95" And Mid$(weapontype.ListItems(i%), InStr(weapontype.ListItems(i%), "(") + 1, 2) <> "99" And Mid$(weapontype.ListItems(i%), InStr(weapontype.ListItems(i%), "(") + 1, 2) <> "90") Then
                            msg = MsgBox("For simple assault, weapon codes must be 40, 90, 95 or 99", 48, "Genesis Error Log")
                            weapontype.ListItems(i%).Selected = False
                            Call ShowApplicableContainers(weapontype)
                            weapontype.SetFocus
                            GoTo exitedite
                        End If
                    End If
                    '===== Data Element 13
                    If Mid$(weapontype.ListItems(i%), InStr(weapontype.ListItems(i%), "(") + 1, 2) <> "11" And _
                        Mid$(weapontype.ListItems(i%), InStr(weapontype.ListItems(i%), "(") + 1, 2) <> "12" And _
                        Mid$(weapontype.ListItems(i%), InStr(weapontype.ListItems(i%), "(") + 1, 2) <> "13" And _
                        Mid$(weapontype.ListItems(i%), InStr(weapontype.ListItems(i%), "(") + 1, 2) <> "14" And _
                        Mid$(weapontype.ListItems(i%), InStr(weapontype.ListItems(i%), "(") + 1, 2) <> "15" Then
                            If automatic(i%) <> "N" Then
                                msg = MsgBox("Weapon Type/Automatic combination invalid.", 48, "Genesis Error Log")
                                Call ShowApplicableContainers(weapontype)
                                weapontype.SetFocus
                                GoTo exitedite
                            End If
                    End If
                End If
                End If
        Next i%
        '===== Error 201
        '===== SCEdit 8/9/95
        Select Case tempucr
            Case "09A", "09B", "09C", "100", "11A", "11B", "11C", "11D", "120", "13A", "13B", "210", "520", "980"
                If foundweapon = 0 Then
                    msg = MsgBox("A valid weapon type must be selected.", 48, "Genesis Error Log")
                    Call ShowApplicableContainers(weapontype)
                    weapontype.SetFocus
                    GoTo exitedite
                End If
        End Select
        If foundweapon > 0 Then
            Select Case tempucr
                Case "09A", "09B", "09C", "100", "11A", "11B", "11C", "11D", "120", "13A", "13B", "210", "520"
                '===== Error 219
                Case Else
                    If Left$(tempucr, 2) <> "90" Then
                        foundother% = 0
                        For q% = 0 To 9
                            If ucrlist(q%).ListIndex > -1 Then
                                Select Case Mid$(ucrlist(q%).List(ucrlist(q%).ListIndex), InStr(ucrlist(q%).List(ucrlist(q%).ListIndex), "(") + 1, 3)
                                    Case "09A", "09B", "09C", "100", "11A", "11B", "11C", "11D", "120", "13A", "13B", "210", "520"
                                        foundother% = 1
                                        q% = 9
                                End Select
                            End If
                        Next q%
                        If foundother% = 0 Then
                            msg = MsgBox("A weapon cannot be selected with this UCR: " + tempucr, 48, "Genesis Error Log")
                            Call ShowApplicableContainers(weapontype)
                            weapontype.SetFocus
                            GoTo exitedite
                        End If
                    End If
            End Select
        End If
        '===== Error 267
        If tempucr = "09A" Or tempucr = "09B" Or tempucr = "09C" Then
            If foundweapon99 > 0 Then
                msg = MsgBox("A weapon type other than 99 = None must be selected for this UCR " + tempucr + ".", 48, "Genesis Error Log")
                Call ShowApplicableContainers(weapontype)
                weapontype.SetFocus
                GoTo exitedite
            End If
        End If
        '===== Data Element 13
        '===== Error 207
        If foundweapon > 1 And foundweapon99 = 1 Then
            msg = MsgBox("If Weapon Type NONE is chosen, no other weapon types can be selected as well.", 48, "Genesis Error Log")
            Call ShowApplicableContainers(weapontype)
            weapontype.SetFocus
            GoTo exitedite
        End If
        '===== SCEdit 8/9/95
        ctpremise% = 0
        For zz% = 1 To premise(t%).ListItems.Count
            If premise(t%).ListItems(zz%).Selected Then
                ctpremise% = ctpremise% + 1
                temppremise = Mid$(premise(t%).ListItems(zz%), InStr(premise(t%).ListItems(zz%), "(") + 1, 2)
                If tempucr = "220" Then
                    If Not (temppremise = "01" Or temppremise = "02" Or temppremise = "03" Or temppremise = "04" Or temppremise = "05" Or temppremise = "06" Or temppremise = "07" Or temppremise = "08" Or temppremise = "09" Or temppremise = "11" Or temppremise = "12" Or temppremise = "14" Or temppremise = "15" Or temppremise = "17" Or temppremise = "19" Or temppremise = "20" Or temppremise = "21" Or temppremise = "22" Or temppremise = "23" Or temppremise = "24" Or temppremise = "25" Or temppremise = "26" Or temppremise = "27" Or temppremise = "28") Then
                        msg = MsgBox("For breaking and entering, premise code must be one of 01, 02, 03, 04, 05, 06, 07, 08, 09, 11, 12, 14, 15, 17, 19, 20, 21, 22, 23, 24, 25, 26, 27 or 28.", 48, "Genesis Error Log")
                        Call ShowApplicableContainers(premise(t%))
                        premise(t%).SetFocus
                        GoTo exitedite
                    End If
                End If
                If tempucr = "23C" Then
                    If Not (temppremise = "01" Or temppremise = "03" Or temppremise = "04" Or temppremise = "05" Or temppremise = "07" Or temppremise = "08" Or temppremise = "11" Or temppremise = "12" Or temppremise = "14" Or temppremise = "17" Or temppremise = "21" Or temppremise = "22" Or temppremise = "23" Or temppremise = "24" Or temppremise = "25" Or temppremise = "26" Or temppremise = "27") Then
                        msg = MsgBox("For shoplifting, premise code must be one of 01, 03, 04, 05, 07, 08, 09, 11, 12, 14, 17, 21, 22, 23, 24, 25, 26 or 27.", 48, "Genesis Error Log")
                        Call ShowApplicableContainers(premise(t%))
                        premise(t%).SetFocus
                        GoTo exitedite
                    End If
                End If
            End If
        Next zz%
        '===== SCEdit 1/31/1
        If ctpremise% = 1 Then
            If temppremise = "18" Then
                msg = MsgBox("If Premise Type 18 (Parking Lot/Parking Garage) is selected, then another premise type must also be selected to further describe it.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(premise(t%))
                premise(t%).SetFocus
                GoTo exitedite
            End If
        End If
        foundother% = 0
        If tempucr = "09C" Then
            For q% = 0 To t% - 1
                If ucrlist(q%).ListIndex > -1 Then
                    foundother% = 1
                    q% = t% - 1
                End If
            Next q%
            If foundother% = 0 Then
                For q% = t% + 1 To 9
                    If ucrlist(q%).ListIndex > -1 Then
                        foundother% = 1
                        q% = 9
                    End If
                Next q%
            End If
            If foundother% = 1 Then
                msg = MsgBox("For Justifiable Homocide there can be no other offenses.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(ucrlist(0))
                ucrlist(0).SetFocus
                GoTo exitedite
            End If
        End If
        '===== SCEdit 8/9/95
        Select Case tempucr
            Case "09C", "979", "992", "980", "978"
                If IsDate(EXCEPTIONALCLEARANCEDATE) Then
                    msg = MsgBox("Exceptional clearance is not allowed on a " + tempucr + " offense.", 48, "Genesis Error Log")
                    Call ShowApplicableContainers(ucrlist(0))
                    ucrlist(0).SetFocus
                    GoTo exitedite
                End If
        End Select
        '-------------------------------------
        '===== Error 257
        If tempucr = "220" Then
            For zz% = 1 To premise(t%).ListItems.Count
                If premise(t%).ListItems(zz%).Selected Then
                    If Mid$(premise(t%).ListItems(zz%), InStr(premise(t%).ListItems(zz%), "(") + 1, 2) = "14" Or Mid$(premise(t%).ListItems(zz%), InStr(premise(t%).ListItems(zz%), "(") + 1, 2) = "19" Then
                        If Val(entered(t%)) < 1 Or Val(entered(t%)) > 99 Then
                            msg = MsgBox("Number of premises entered must be entered (01-99).", 48, "Genesis Error Log")
                            Call ShowApplicableContainers(premise(t%))
                            premise(t%).SetFocus
                            GoTo exitedite
                        End If
                    End If
                End If
            Next zz%
        End If
        '===== Error 252
        If Val(entered(t%)) > 0 Then
            If tempucr <> "220" Then
                msg = MsgBox("If Number of Premises Entered is valued, the UCR code must be 220 (Burglary/B&E).", 48, "Genesis Error Log")
                Call ShowApplicableContainers(ucrlist(0))
                ucrlist(0).SetFocus
                GoTo exitedite
            End If
            found1419% = 0
            For zz% = 1 To premise(t%).ListItems.Count
                If premise(t%).ListItems(zz%).Selected Then
                    If Mid$(premise(t%).ListItems(zz%), InStr(premise(t%).ListItems(zz%), "(") + 1, 2) = "14" Or Mid$(premise(t%).ListItems(zz%), InStr(premise(t%).ListItems(zz%), "(") + 1, 2) = "19" Then
                        found1419% = 1
                        zz% = premise(t%).ListItems.Count
                    End If
                End If
            Next zz%
            If found1419% = 0 Then
                msg = MsgBox("If Number of Premises Entered is valued, the Premise Type must be 14 (Hotel/Motel) or 19 (Rental/Storage Facility).", 48, "Genesis Error Log")
                Call ShowApplicableContainers(premise(t%))
                premise(t%).SetFocus
                GoTo exitedite
            End If
        End If
        '===== Mandatories E - 7
        '===== Data Element 7
        '===== Error 201
        If Not completedy(t%) And Not completedn(t%) Then
            msg = MsgBox("A valid completed code of Yes or No must be selected in COMPLETED.", 48, "Genesis Error Log")
            Call ShowApplicableContainers(completedy(t%))
            completedy(t%).SetFocus
            GoTo exitedite
        End If
        '=== Additional F 2
        '===== Data Element 7
        '===== Error 256
        If tempucr = "13A" Or tempucr = "13B" Or tempucr = "13C" Then
            If Not completedy(t%) Then
                msg = MsgBox("Crimes Against Persons must show Completed.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(completedy(t%))
                completedy(t%).SetFocus
                GoTo exitedite
            End If
            If Not individual And Not policeofficer Then
                msg = MsgBox("Crimes Against Persons must be associated with an Individual or Police Officer.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(individual)
                individual.SetFocus
                GoTo exitedite
            End If
        End If
        If tempucr = "13A" Or tempucr = "13B" Then
            selw% = 0
            For iop% = 1 To weapontype.ListItems.Count
                If weapontype.ListItems(iop%).Selected Then
                    selw% = 1
                    iop% = weapontype.ListItems.Count
                End If
            Next iop%
            If selw% = 0 Then
                msg = MsgBox("A weapon type must be selected for this UCR " + tempucr + ".", 48, "Genesis Error Log")
                Call ShowApplicableContainers(weapontype)
                weapontype.SetFocus
                GoTo exitedite
            End If
            '---MOVED TO EXPORT ---
            'If injury.SelectedItem Is Nothing Then
            '    msg = MsgBox("An injury type must be selected for this UCR " + tempucr + ".", 48, "Genesis Error Log")
            '    GoTo exitedite
            'End If
        End If
        '===== Error 462
        If tempucr = "13A" Then
            selhom% = 0
            For xx% = 1 To HOMOCIDE(t%).ListItems.Count
                If HOMOCIDE(t%).ListItems(xx%).Selected Then
                    selhom% = 1
                    If Mid$(HOMOCIDE(t%).ListItems(xx%), InStr(HOMOCIDE(t%).ListItems(xx%), "(") + 1, 2) = "07" Then
                        msg = MsgBox("Reason 07 = Mercy Killing not allowed for UCR 13A.", 48, "Genesis Error Log")
                        Call ShowApplicableContainers(HOMOCIDE(t%))
                        HOMOCIDE(t%).SetFocus
                        GoTo exitedite
                    End If
                End If
            Next xx%
            If selhom% = 0 Then
                msg = MsgBox("Aggravated Assault/Homocide Circumstances must be selected for Aggravated Assault.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(HOMOCIDE(t%))
                HOMOCIDE(t%).SetFocus
                GoTo exitedite
            End If
        End If
        '===== Data Element 13
        If tempucr = "13B" Then
            For TU% = 1 To weapontype.ListItems.Count
                If weapontype.ListItems(TU%).Selected Then
                    TESTWEAPON = Mid$(weapontype.ListItems(TU%), InStr(weapontype.ListItems(TU%), "(") + 1, 2)
                    If TESTWEAPON <> "40" And TESTWEAPON <> "90" And TESTWEAPON <> "95" And TESTWEAPON <> "99" And TESTWEAPON <> "  " Then
                        msg = MsgBox("Weapon type invalid for simple assault.", 48, "Genesis Error Log")
                        Call ShowApplicableContainers(weapontype)
                        weapontype.SetFocus
                        GoTo exitedite
                    End If
                End If
            Next TU%
        End If
            
        
        '===== Additional F 4
        '===== Data Element 10
        '===== Data Element 11
        '===== Error 204,253,254
        '===== SCEdit 7.2 Amendment P7-4
        If tempucr = "220" Then
            If societypublic Then
                msg = MsgBox("Burglary/B&E cannot be Society/Public.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(societypublic)
                societypublic.SetFocus
                GoTo exitedite
            End If
        End If
        If tempucr = "220" Or tempucr = "23F" Or tempucr = "23G" Or tempucr = "240" Then
            If FORCEDENTRYY(t%) = 0 And FORCEDENTRYN(t%) = 0 Then
                msg = MsgBox("A value of Y or N must be entered for Forced Entry.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(FORCEDENTRYY(t%))
                FORCEDENTRYY(t%).SetFocus
                GoTo exitedite
            End If
        Else
            If FORCEDENTRYY(t%) = 1 Or FORCEDENTRYN(t%) = 1 Then
                msg = MsgBox("A value of Y or N CANNOT be entered for Forced Entry.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(FORCEDENTRYY(t%))
                FORCEDENTRYY(t%).SetFocus
                GoTo exitedite
            End If
        End If

        

        '===Additional F 5
        '===== Data Element 12
        If tempucr = "250" Then
            sela% = 0
            For iop% = 1 To activity(t%).ListItems.Count
                If activity(t%).ListItems(iop%).Selected Then
                    sela% = 1
                    iop% = activity(t%).ListItems.Count
                End If
            Next iop%
            If sela% = 0 Then
                msg = MsgBox("Type of Activity must be selected for Counterfeiting.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(activity(t%))
                activity(t%).SetFocus
                GoTo exitedite
            End If
        End If
            
        '===Additional F 7
        '===== Data Element 7
        '===== Data Element 12
        If tempucr = "35A" Or tempucr = "35B" Then
            sela% = 0
            For iop% = 1 To activity(t%).ListItems.Count
                If activity(t%).ListItems(iop%).Selected Then
                    sela% = 1
                    iop% = activity(t%).ListItems.Count
                End If
            Next iop%
            If sela% = 0 Then
                msg = MsgBox("Type of Activity must be selected for Drug/Narcotic Offenses.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(activity(t%))
                activity(t%).SetFocus
                GoTo exitedite
            End If
            If Not societypublic Then
                msg = MsgBox("Drug crimes victims should be Society/Public.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(societypublic)
                societypublic.SetFocus
                GoTo exitedite
            End If
        End If

        
        '===== Additional F 9
        If tempucr = "210" Then
            sela% = 0
            For iop% = 1 To weapontype.ListItems.Count
                If weapontype.ListItems(iop%).Selected Then
                    sela% = 1
                    iop% = weapontype.ListItems.Count
                End If
            Next iop%
            If sela% = 0 Then
                msg = MsgBox("A weapon type must be selected for this UCR " + tempucr + ".", 48, "Genesis Error Log")
                Call ShowApplicableContainers(weapontype)
                weapontype.SetFocus
                GoTo exitedite
            End If
            'If individual Then
            '    If injury.SelectedItem Is Nothing Then
            '        msg = MsgBox("An injury type must be selected for this UCR " + tempucr + ".", 48, "Genesis Error Log")
            '        GoTo exitedite
            '    End If
            'End If
        End If
        
        '===== Additional F 11
        '===== Data Element 7
        If tempucr = "39A" Or tempucr = "39B" Or tempucr = "39C" Or tempucr = "39D" Then
            If Not societypublic Then
                msg = MsgBox("Drug crimes victims should be Society/Public.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(societypublic)
                societypublic.SetFocus
                GoTo exitedite
            End If
            '===== Data Element 12
            If tempucr = "39C" Then
                sela% = 0
                For iop% = 1 To activity(t%).ListItems.Count
                    If activity(t%).ListItems(iop%).Selected Then
                        sela% = 1
                        iop% = activity(t%).ListItems.Count
                    End If
                Next iop%
                If sela% = 0 Then
                    msg = MsgBox("A type of activity must be entered for UCR " + tempucr + ".", 48, "Genesis Error Log")
                    Call ShowApplicableContainers(activity(t%))
                    activity(t%).SetFocus
                    GoTo exitedite
                End If
            End If
        End If
        
        '===== Additional F 12
        '===== Data Element 7, 13
        '===== Error 256
        If tempucr = "09A" Or tempucr = "09B" Or tempucr = "09C" Then
            If Not completedy(t%) Then
                msg = MsgBox("Homocide must show as completed.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(completedy(t%))
                completedy(t%).SetFocus
                GoTo exitedite
            End If
            sela% = 0
            For iop% = 1 To weapontype.ListItems.Count
                If weapontype.ListItems(iop%).Selected Then
                    sela% = 1
                    iop% = weapontype.ListItems.Count
                End If
            Next iop%
            If sela% = 0 Then
                msg = MsgBox("A weapon type must be selected for this UCR " + tempucr + ".", 48, "Genesis Error Log")
                Call ShowApplicableContainers(weapontype)
                weapontype.SetFocus
                GoTo exitedite
            End If
            If Not individual Then
                msg = MsgBox("Homocide victims should be Individual.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(individual)
                individual.SetFocus
                GoTo exitedite
            End If
            sela% = 0
            For iop% = 1 To HOMOCIDE(t%).ListItems.Count
                If HOMOCIDE(t%).ListItems(iop%).Selected Then
                    sela% = 1
                    iop% = HOMOCIDE(t%).ListItems.Count
                End If
            Next iop%
            If sela% = 0 Then
                msg = MsgBox("Aggravated Assault/Homocide Circumstances must be selected for UCR " + tempucr + ".", 48, "Genesis Error Log")
                Call ShowApplicableContainers(HOMOCIDE(t%))
                HOMOCIDE(t%).SetFocus
                GoTo exitedite
            End If
            '===== Error 463
            If tempucr = "09C" Then
                hct% = 0
                For xx% = 1 To HOMOCIDE(t%).ListItems.Count
                    If HOMOCIDE(t%).ListItems(xx%).Selected Then
                        hct% = hct% + 1
                        If Mid$(HOMOCIDE(t%).ListItems(xx%), InStr(HOMOCIDE(t%).ListItems(xx%), "(") + 1, 2) <> "20" And Mid$(HOMOCIDE(t%).ListItems(xx%), InStr(HOMOCIDE(t%).ListItems(xx%), "(") + 1, 2) <> "21" Then
                            msg = MsgBox("Aggravated Assault/Homocide Circumstances must be 20 or 21.", 48, "Genesis Error Log")
                            Call ShowApplicableContainers(HOMOCIDE(t%))
                            HOMOCIDE(t%).SetFocus
                            GoTo exitedite
                        End If
                    End If
                Next xx%
                If hct% > 2 Then
                    msg = MsgBox("A maximum of 2 Aggravated Assault/Homocide Circumstances may be selected.", 48, "Genesis Error Log")
                    Call ShowApplicableContainers(HOMOCIDE(t%))
                    HOMOCIDE(t%).SetFocus
                    GoTo exitedite
                End If
                If additional(t%).ListIndex = -1 Then
                    msg = MsgBox("Additional Circumstances must be selected.", 48, "Genesis Error Log")
                    Call ShowApplicableContainers(additional(t%))
                    additional(t%).SetFocus
                    GoTo exitedite
                End If
            End If
            
        '===== Error 456
        For xx% = 1 To HOMOCIDE(t%).ListItems.Count
            If HOMOCIDE(t%).ListItems(xx%).Selected Then
            '===== Error 480
                If Mid$(HOMOCIDE(t%).ListItems(xx%), InStr(HOMOCIDE(t%).ListItems(xx%), "(") + 1, 2) <> "20" And Mid$(HOMOCIDE(t%).ListItems(xx%), InStr(HOMOCIDE(t%).ListItems(xx%), "(") + 1, 2) = "08" And ucrct% < 2 Then
                    msg = MsgBox("If reason 08=Other Felony Involved is selected, there must be at least 2 UCR's entered.", 48, "Genesis Error Log")
                    Call ShowApplicableContainers(HOMOCIDE(t%))
                    HOMOCIDE(t%).SetFocus
                    GoTo exitedite
                End If
                If Mid$(HOMOCIDE(t%).ListItems(xx%), InStr(HOMOCIDE(t%).ListItems(xx%), "(") + 1, 2) <> "20" And Mid$(HOMOCIDE(t%).ListItems(xx%), InStr(HOMOCIDE(t%).ListItems(xx%), "(") + 1, 2) = "10" Then
                    For xxx% = 1 To xx% - 1
                        If HOMOCIDE(t%).ListItems(xxx%).Selected Then
                            msg = MsgBox("If 10=Unknown Circumstances is selected, no other circumstances can be selected.", 48, "Genesis Error Log")
                            Call ShowApplicableContainers(HOMOCIDE(t%))
                            HOMOCIDE(t%).SetFocus
                            GoTo exitedite
                        End If
                    Next xxx%
                    For xxx% = xx% + 1 To HOMOCIDE(t%).ListItems.Count
                        If HOMOCIDE(t%).ListItems(xxx%).Selected Then
                            msg = MsgBox("If 10=Unknown Circumstances is selected, no other circumstances can be selected.", 48, "Genesis Error Log")
                            Call ShowApplicableContainers(HOMOCIDE(t%))
                            HOMOCIDE(t%).SetFocus
                            GoTo exitedite
                        End If
                    Next xxx%
                End If
                '===== Error 456
                Select Case Val(Mid$(HOMOCIDE(t%).ListItems(xx%), InStr(HOMOCIDE(t%).ListItems(xx%), "(") + 1, 2))
                    Case 1 To 10
                        '===== Error 477
                        If tempucr <> "13A" And tempucr <> "09A" Then
                            msg = MsgBox("The Aggravated Assault/Homocide reason chosen is invalid for the UCR.", 48, "Genesis Error Log")
                            Call ShowApplicableContainers(HOMOCIDE(t%))
                            HOMOCIDE(t%).SetFocus
                            GoTo exitedite
                        End If
                        For xxx% = 1 To xx% - 1
                            If HOMOCIDE(t%).ListItems(xxx%).Selected Then
                                Select Case Val(Mid$(HOMOCIDE(t%).ListItems(xxx%), InStr(HOMOCIDE(t%).ListItems(xxx%), "(") + 1, 2))
                                    Case 20 To 34
                                        msg = MsgBox("Only one category may be selected within Aggravted Assault/Homocide reasons.", 48, "Genesis Error Log")
                                        Call ShowApplicableContainers(HOMOCIDE(t%))
                                        HOMOCIDE(t%).SetFocus
                                        GoTo exitedite
                                End Select
                            End If
                        Next xxx%
                        For xxx% = xx% + 1 To HOMOCIDE(t%).ListItems.Count
                            If HOMOCIDE(t%).ListItems(xxx%).Selected Then
                                Select Case Val(Mid$(HOMOCIDE(t%).ListItems(xxx%), InStr(HOMOCIDE(t%).ListItems(xxx%), "(") + 1, 2))
                                    Case 20 To 34
                                        msg = MsgBox("Only one category may be selected within Aggravted Assault/Homocide reasons.", 48, "Genesis Error Log")
                                        Call ShowApplicableContainers(HOMOCIDE(t%))
                                        HOMOCIDE(t%).SetFocus
                                        GoTo exitedite
                                End Select
                            End If
                        Next xxx%
                    Case 20 To 21
                        '===== Error 477
                        If tempucr <> "09C" Then
                            msg = MsgBox("The Aggravated Assault/Homocide reason chosen is invalid for the UCR.", 48, "Genesis Error Log")
                            Call ShowApplicableContainers(HOMOCIDE(t%))
                            HOMOCIDE(t%).SetFocus
                            GoTo exitedite
                        End If
                        For xxx% = 1 To xx% - 1
                            If HOMOCIDE(t%).ListItems(xxx%).Selected Then
                                Select Case Val(Mid$(HOMOCIDE(t%).ListItems(xxx%), InStr(HOMOCIDE(t%).ListItems(xxx%), "(") + 1, 2))
                                    Case 1 To 10
                                        msg = MsgBox("Only one category may be selected within Aggravted Assault/Homocide reasons.", 48, "Genesis Error Log")
                                        Call ShowApplicableContainers(HOMOCIDE(t%))
                                        HOMOCIDE(t%).SetFocus
                                        GoTo exitedite
                                    Case 30 To 34
                                        msg = MsgBox("Only one category may be selected within Aggravted Assault/Homocide reasons.", 48, "Genesis Error Log")
                                        Call ShowApplicableContainers(HOMOCIDE(t%))
                                        HOMOCIDE(t%).SetFocus
                                        GoTo exitedite
                                End Select
                            End If
                        Next xxx%
                        For xxx% = xx% + 1 To HOMOCIDE(t%).ListItems.Count
                            If HOMOCIDE(t%).ListItems(xxx%).Selected Then
                                Select Case Val(Mid$(HOMOCIDE(t%).ListItems(xxx%), InStr(HOMOCIDE(t%).ListItems(xxx%), "(") + 1, 2))
                                    Case 1 To 10
                                        msg = MsgBox("Only one category may be selected within Aggravted Assault/Homocide reasons.", 48, "Genesis Error Log")
                                        Call ShowApplicableContainers(HOMOCIDE(t%))
                                        HOMOCIDE(t%).SetFocus
                                        GoTo exitedite
                                    Case 30 To 34
                                        msg = MsgBox("Only one category may be selected within Aggravted Assault/Homocide reasons.", 48, "Genesis Error Log")
                                        Call ShowApplicableContainers(HOMOCIDE(t%))
                                        HOMOCIDE(t%).SetFocus
                                        GoTo exitedite
                                End Select
                            End If
                        Next xxx%
                    Case 30 To 34
                        '===== Error 477
                        If tempucr <> "09B" Then
                            msg = MsgBox("The Aggravated Assault/Homocide reason chosen is invalid for the UCR.", 48, "Genesis Error Log")
                            Call ShowApplicableContainers(HOMOCIDE(t%))
                            HOMOCIDE(t%).SetFocus
                            GoTo exitedite
                        End If
                        For xxx% = 1 To xx% - 1
                            If HOMOCIDE(t%).ListItems(xxx%).Selected Then
                                Select Case Val(Mid$(HOMOCIDE(t%).ListItems(xxx%), InStr(HOMOCIDE(t%).ListItems(xxx%), "(") + 1, 2))
                                    Case 1 To 21
                                        msg = MsgBox("Only one category may be selected within Aggravted Assault/Homocide reasons.", 48, "Genesis Error Log")
                                        Call ShowApplicableContainers(HOMOCIDE(t%))
                                        HOMOCIDE(t%).SetFocus
                                        GoTo exitedite
                                End Select
                            End If
                        Next xxx%
                        For xxx% = xx% + 1 To HOMOCIDE(t%).ListItems.Count
                            If HOMOCIDE(t%).ListItems(xxx%).Selected Then
                                Select Case Val(Mid$(HOMOCIDE(t%).ListItems(xxx%), InStr(HOMOCIDE(t%).ListItems(xxx%), "(") + 1, 2))
                                    Case 1 To 21
                                        msg = MsgBox("Only one category may be selected within Aggravted Assault/Homocide reasons.", 48, "Genesis Error Log")
                                        Call ShowApplicableContainers(HOMOCIDE(t%))
                                        HOMOCIDE(t%).SetFocus
                                        GoTo exitedite
                                End Select
                            End If
                        Next xxx%
                End Select
            End If
        Next xx%

        End If
                

        '===== Error 455
        If tempucr = "09A" Or tempucr = "09B" Or tempucr = "09C" Or tempucr = "13A" Then
            For xx% = 1 To HOMOCIDE(t%).ListItems.Count
                If HOMOCIDE(t%).ListItems(xx%).Selected Then
                    If Mid$(HOMOCIDE(t%).ListItems(xx%), InStr(HOMOCIDE(t%).ListItems(xx%), "(") + 1, 2) <> "20" And Mid$(HOMOCIDE(t%).ListItems(xx%), InStr(HOMOCIDE(t%).ListItems(xx%), "(") + 1, 2) = "20" Or _
                       Mid$(HOMOCIDE(t%).ListItems(xx%), InStr(HOMOCIDE(t%).ListItems(xx%), "(") + 1, 2) <> "20" And Mid$(HOMOCIDE(t%).ListItems(xx%), InStr(HOMOCIDE(t%).ListItems(xx%), "(") + 1, 2) = "21" Then
                        If additional(t%).ListIndex = -1 Then
                            msg = MsgBox("Additional Circumstances must be selected.", 48, "Genesis Error Log")
                            Call ShowApplicableContainers(HOMOCIDE(t%))
                            HOMOCIDE(t%).SetFocus
                            GoTo exitedite
                        End If
                    End If
                End If
            Next xx%
        End If
        
        '===== Additional F 13
        '===== Data Element 7
        If tempucr = "100" Then
            sela% = 0
            For iop% = 1 To weapontype.ListItems.Count
                If weapontype.ListItems(iop%).Selected Then
                    sela% = 1
                    iop% = weapontype.ListItems.Count
                End If
            Next iop%
            If sela% = 0 Then
                msg = MsgBox("A weapon type must be selected for this UCR " + tempucr + ".", 48, "Genesis Error Log")
                Call ShowApplicableContainers(weapontype)
                weapontype.SetFocus
                GoTo exitedite
            End If
            'If injury.SelectedItem Is Nothing Then
            '    msg = MsgBox("An injury type must be selected for this UCR " + tempucr + ".", 48, "Genesis Error Log")
            '    GoTo exitedite
            'End If
            If Not individual Then
                msg = MsgBox("Kidnaping/abduction crimes victims should be Individual.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(individual)
                individual.SetFocus
                GoTo exitedite
            End If
        End If

        '===== Additional F 16
        '===== Data Element 12
        If tempucr = "370" Then
            sela% = 0
            For iop% = 1 To activity(t%).ListItems.Count
                If activity(t%).ListItems(iop%).Selected Then
                    sela% = 1
                    iop% = activity(t%).ListItems.Count
                End If
            Next iop%
            If sela% = 0 Then
                msg = MsgBox("Type of Activity must be selected for Pornography.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(activity(t%))
                activity(t%).SetFocus
                GoTo exitedite
            End If
            If Not societypublic Then
                msg = MsgBox("Pornography crimes victims should be Society/Public.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(societypublic)
                societypublic.SetFocus
                GoTo exitedite
            End If
        End If
            
        '===== Additional F 17
        If tempucr = "40A" Then
            If Not societypublic Then
                msg = MsgBox("Prostitution crimes victims should be Society/Public.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(societypublic)
                societypublic.SetFocus
                GoTo exitedite
            End If
        End If
            
        '===== Additional F 18
        If tempucr = "120" Then
            sela% = 0
            For iop% = 1 To weapontype.ListItems.Count
                If weapontype.ListItems(iop%).Selected Then
                    sela% = 1
                    iop% = weapontype.ListItems.Count
                End If
            Next iop%
            If sela% = 0 Then
                msg = MsgBox("A weapon type must be selected for this UCR " + tempucr + ".", 48, "Genesis Error Log")
                Call ShowApplicableContainers(weapontype)
                weapontype.SetFocus
                GoTo exitedite
            End If
            'If individual Then
            '    If injury.SelectedItem Is Nothing Then
            '        msg = MsgBox("A type of injury must be selected.", 48, "Genesis Error Log")
            '        GoTo exitedite
            '    End If
            'End If
        End If
            
        '===== Additional F 19
        If tempucr = "11A" Or tempucr = "11B" Or tempucr = "11C" Or tempucr = "11D" Then
            sela% = 0
            For iop% = 1 To weapontype.ListItems.Count
                If weapontype.ListItems(iop%).Selected Then
                    sela% = 1
                    iop% = weapontype.ListItems.Count
                End If
            Next iop%
            If sela% = 0 Then
                msg = MsgBox("A weapon type must be selected for this UCR " + tempucr + ".", 48, "Genesis Error Log")
                Call ShowApplicableContainers(weapontype)
                weapontype.SetFocus
                GoTo exitedite
            End If
            If Not individual Then
                msg = MsgBox("Individual must be chosen for UCR " + tempucr + ".", 48, "Genesis Error Log")
                Call ShowApplicableContainers(individual)
                individual.SetFocus
                GoTo exitedite
            End If
        End If
        
        '===== Additional F 20
        If tempucr = "36A" Or tempucr = "36B" Then
            If Not individual Then
                msg = MsgBox("Individual must be chosen for UCR " + tempucr + ".", 48, "Genesis Error Log")
                Call ShowApplicableContainers(individual)
                individual.SetFocus
                GoTo exitedite
            End If
        End If
        
        '===== Additional F 21
        '===== Data Element 12
        If tempucr = "280" Then
            sela% = 0
            For iop% = 1 To activity(t%).ListItems.Count
                If activity(t%).ListItems(iop%).Selected Then
                    sela% = 1
                    iop% = activity(t%).ListItems.Count
                End If
            Next iop%
            If sela% = 0 Then
                msg = MsgBox("Type of Activity must be selected for Stolen Property.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(activity(t%))
                activity(t%).SetFocus
                GoTo exitedite
            End If
        End If
        
        '===== Additional F 22
        '===== Data Element 12
        If tempucr = "520" Then
            sela% = 0
            For iop% = 1 To activity(t%).ListItems.Count
                If activity(t%).ListItems(iop%).Selected Then
                    sela% = 1
                    iop% = activity(t%).ListItems.Count
                End If
            Next iop%
            If sela% = 0 Then
                msg = MsgBox("Type of Activity must be selected for Weapons Law violations.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(activity(t%))
                activity(t%).SetFocus
                GoTo exitedite
            End If
            sela% = 0
            For iop% = 1 To weapontype.ListItems.Count
                If weapontype.ListItems(iop%).Selected Then
                    sela% = 1
                    iop% = weapontype.ListItems.Count
                End If
            Next iop%
            If sela% = 0 Then
                msg = MsgBox("A weapon type must be selected for this UCR " + tempucr + ".", 48, "Genesis Error Log")
                Call ShowApplicableContainers(weapontype)
                weapontype.SetFocus
                GoTo exitedite
            End If
            If Not societypublic Then
                msg = MsgBox("Society must be chosen for UCR " + tempucr + ".", 48, "Genesis Error Log")
                Call ShowApplicableContainers(societypublic)
                societypublic.SetFocus
                GoTo exitedite
            End If
        End If
        
        '==== Mandatories E - 9
        '===== Error 201
        If ctpremise% = 0 Then
            msg = MsgBox("A valid premise type must be selected.", 48, "Genesis Error Log")
            Call ShowApplicableContainers(entered(t%))
            entered(t%).SetFocus
            GoTo exitedite
        End If
        If entered(t%) = "" Then
            entered(t%) = "0"
        End If
        If tempucr = "09A" Or tempucr = "09B" Or tempucr = "100" Or tempucr = "120" Or tempucr = "11A" Or tempucr = "11B" Or tempucr = "11C" Or tempucr = "11D" Or tempucr = "13A" Or tempucr = "13B" Or tempucr = "13C" Then
            tempactivity = ""
            foundn% = 0
            foundg% = 0
            foundj% = 0
            For q% = 1 To gactivity(t%).ListItems.Count
                If gactivity(t%).ListItems(q%).Selected Then
                    Select Case Mid$(gactivity(t%).ListItems(q%), InStr(gactivity(t%).ListItems(q%), "(") + 1, 1)
                        Case "N"
                            foundn% = 1
                        Case "G"
                            foundg% = 1
                        Case "J"
                            foundj% = 1
                    End Select
                End If
            Next q%
            If foundn% = 1 And (foundj% = 1 Or foundg% = 1) Then
                msg = MsgBox("If activity code is 'N', codes 'J' and 'G' are not permitted.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(gactivity(t%))
                gactivity(t%).SetFocus
                GoTo exitedite
            End If
        End If
    End If
Next t%
'===== Data Element 2
'===== Error 001, 101, 201, 301, 401, 501, 601, 701
If incidentnumber = "" Then
    msg = MsgBox("A valid incidentnumber must be entered.", 48, "Genesis Error Log")
    Call ShowApplicableContainers(incidentnumber)
    incidentnumber.SetFocus
    GoTo exitedite
End If
If Len(incidentnumber) > 12 Then
    msg = MsgBox("The Incident Number cannot be over 12 characters long.", 48, "Genesis Error Log")
    Call ShowApplicableContainers(incidentnumber)
    incidentnumber.SetFocus
    GoTo exitedite
End If
incidentnumber = UCase(incidentnumber)
'===== Error 017, 117, 217, 317, 417, 517, 617, 717
For t% = 1 To Len(incidentnumber)
    If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789- ", Mid$(incidentnumber, t%, 1)) = 0 Then
        msg = MsgBox("An invalid character has been found in the Incident Number field.  Valid characters are A-Z, 0-9, and Hyphen.  Do not enter any Blanks because these are computer generated.", 48, "Genesis Error Log")
        t% = Len(incidentnumber)
        Call ShowApplicableContainers(incidentnumber)
        incidentnumber.SetFocus
        GoTo exitedite
    End If
Next t%
'===== Error 015, 115, 215, 315, 415, 515, 615, 715
TEMP$ = ""
For yy% = Len(incidentnumber) To 1 Step -1
    If Mid$(incidentnumber, yy%, 1) <> " " Then
        TEMP$ = Left$(incidentnumber, yy%)
        yy% = 1
    End If
Next yy%
If InStr(TEMP$, " ") > 0 Then
    msg = MsgBox("An invalid character has been found in the Incident Number field.  Valid characters are A-Z, 0-9, and Hyphen.  Do not enter any Blanks because these are computer generated.", 48, "Genesis Error Log")
    t% = Len(incidentnumber)
    Call ShowApplicableContainers(incidentnumber)
    incidentnumber.SetFocus
    GoTo exitedite
End If
If Len(incidentnumber) < 12 And Len(incidentnumber) > 0 Then
    incidentnumber = incidentnumber + Space$(12 - Len(incidentnumber))
End If
'If Not IsDate(dispatchdate) Then
'    msg = MsgBox("Dispatch date is invalid.", 48, "Genesis Error Log")
'    GoTo exitedite
'End If
'If Not IsDate(DISPATCHTIME) Then
'    msg = MsgBox("Dispatch time is invalid.", 48, "Genesis Error Log")
'    GoTo exitedite
'End If
'If Not IsDate(TIMEARRIVED) Then
'    msg = MsgBox("Arrival time is invalid.", 48, "Genesis Error Log")
'    GoTo exitedite
'End If
'If Not IsDate(DEPARTINGTIME) Then
'    msg = MsgBox("Departure time is invalid.", 48, "Genesis Error Log")
'    GoTo exitedite
'End If
Call ShowApplicableContainers(incidentnumber)
incidentnumber.SetFocus
GoTo goodedite
exitedite:
editerr = 1

goodedite:

Exit Sub
'RLB Bandaid
rlbErr:
    If Err.Number = 5 Then Resume Next
    
End Sub

Public Sub editvictim(editerr, typeedit As Integer)
'RLB Bandaid
On Error GoTo rlbErr
If Len(incidentnumber) < 12 And Len(incidentnumber) > 0 Then
    incidentnumber = incidentnumber + Space$(12 - Len(incidentnumber))
End If
If age(1) > "" Then
    If Val(age(1)) = 0 And age(1) <> "00" Then
        msg = MsgBox("Age must be entered in the format of a single age: X or XX (i.e. 3 or 42) or in the format of a range of ages: XXXX (i.e. 2025, meaning 20 - 25 years old).", 48, "Genesis Error Log")
        Call ShowApplicableContainers(age(1))
        age(1).SetFocus
        GoTo exiteditv
    End If
End If
'===== Error 404
If age(1) <> "NN" And age(1) <> "NB" And age(1) <> "BB" Then
    For t% = 1 To Len(age(1))
        If InStr("0123456789-", Mid$(age(1), t%, 1)) = 0 Then
            msg = MsgBox("An invalid entry in age has been found.  Valid Entry Formats:  X, XX, X-Y, XX-Y, X-YY, XX-YY, XXYY", 48, "Genesis Error Log")
            t% = Len(age(1))
            Call ShowApplicableContainers(age(1))
            age(1).SetFocus
            GoTo exiteditv
        End If

    Next t%
End If
If InStr(age(1), "-") > 0 Then
    ag1 = Val(Left$(age(1), InStr(age(1), "-") - 1))
    ag2 = Val(Mid$(age(1), InStr(age(1), "-") + 1))
    age(1) = Format$(ag1, "00") + Format$(ag2, "00")
End If
'===== Error 410,422
If Len(age(1)) = 4 Then
    If Val(Left$(age(1), 2)) >= Val(Right$(age(1), 2)) Then
        msg = MsgBox("For an age range, the first age must be less than the second age.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(age(1))
        age(1).SetFocus
        GoTo exiteditv
    End If
    If Val(Left$(age(1), 2)) = 0 Then
        msg = MsgBox("The low value in an age range cannot be 0.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(age(1))
        age(1).SetFocus
        GoTo exiteditv
    End If
End If
'===== Error 450
For r% = 10 To 19
    If relationship(r%).ListIndex > -1 Then
        temprel = Mid$(relationship(r%).List(relationship(r%).ListIndex), InStr(relationship(r%).List(relationship(r%).ListIndex), "(") + 1, 2)
        If temprel = "SE" Then
            If Val(age(1)) < 10 Then
                msg = MsgBox("The relationship of victim to subject cannot be 'SE' when victim's age is less than 10.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(age(1))
                age(1).SetFocus
                GoTo exiteditv
            End If
        End If
    End If
Next r%
'===== Error 401
If Not individual And Not business And Not financialinstitution And Not government And Not religiousorganization And Not societypublic And Not other And Not unknown And Not policeofficer Then
    msg = MsgBox("A Type of Victim must be chosen.", 48, "Genesis Error Log")
    Call ShowApplicableContainers(individual)
    individual.SetFocus
    GoTo exiteditv
End If
editvictimdetail:
'===== Data Element 20, 21, 22
Dim drugs(2)
drugs(0) = ""
drugs(1) = ""
drugs(2) = ""
dct% = 0
For ii% = 0 To 2
    If drugtype(ii%).ListIndex > -1 Then
        drugs(dct%) = drugtype(ii%).List(drugtype(ii%).ListIndex)
        dct% = dct% + 1
    End If
Next ii%
If dct% > 0 Then
    For Z% = 0 To dct%
        '===== Error 306
        If Z% > 0 Then
            For zz% = 0 To Z% - 1
                If drugs(zz%) = drugs(Z%) Then
                    If InStr(drugmeasurement(zz%).List(drugmeasurement(zz%).ListIndex), "GM=") > 0 Or _
                       InStr(drugmeasurement(zz%).List(drugmeasurement(zz%).ListIndex), "KG=") > 0 Or _
                       InStr(drugmeasurement(zz%).List(drugmeasurement(zz%).ListIndex), "OZ=") > 0 Or _
                       InStr(drugmeasurement(zz%).List(drugmeasurement(zz%).ListIndex), "LB=") > 0 Then
                        If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "GM=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "KG=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "OZ=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "LB=") > 0 Then
                                msg = MsgBox("Two different measurements within the same weight category cannot be entered for the same type of drug.", 48, "Genesis Error Log")
                                Call ShowApplicableContainers(drugmeasurement(zz%))
                                drugmeasurement(zz%).SetFocus
                                GoTo exiteditv
                        End If
                    End If
                    If InStr(drugmeasurement(zz%).List(drugmeasurement(zz%).ListIndex), "ML=") > 0 Or _
                       InStr(drugmeasurement(zz%).List(drugmeasurement(zz%).ListIndex), "LT=") > 0 Or _
                       InStr(drugmeasurement(zz%).List(drugmeasurement(zz%).ListIndex), "FO=") > 0 Or _
                       InStr(drugmeasurement(zz%).List(drugmeasurement(zz%).ListIndex), "GL=") > 0 Then
                        If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "ML=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "LT=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "FO=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "GL=") > 0 Then
                                msg = MsgBox("Two different measurements within the same weight category cannot be entered for the same type of drug.", 48, "Genesis Error Log")
                                Call ShowApplicableContainers(drugmeasurement(zz%))
                                drugmeasurement(zz%).SetFocus
                                GoTo exiteditv
                        End If
                    End If
                    If InStr(drugmeasurement(zz%).List(drugmeasurement(zz%).ListIndex), "DU=") > 0 Or _
                       InStr(drugmeasurement(zz%).List(drugmeasurement(zz%).ListIndex), "NP=") > 0 Then
                        If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "DU=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "NP=") > 0 Then
                                msg = MsgBox("Two different measurements within the same weight category cannot be entered for the same type of drug.", 48, "Genesis Error Log")
                                Call ShowApplicableContainers(drugmeasurement(zz%))
                                drugmeasurement(zz%).SetFocus
                                GoTo exiteditv
                        End If
                    End If
                End If
            Next zz%
        End If
        For t% = 1 To Len(drugamt(Z%))
            If InStr("0123456789.", Mid$(drugamt(Z%), t%, 1)) = 0 Then
                msg = MsgBox("Only 0123456789. are allowed in drug amount.  Enter all fractions as decimals (i.e. 1/2 = .5).", 48, "Genesis Error Log")
                Call ShowApplicableContainers(drugamt(Z%))
                drugamt(Z%).SetFocus
                GoTo exiteditv
            End If
        Next t%
        '=====SCEdit 4/21/92 P29  Overtuend drug amt and measurement error
        'If drugamt(z%) = "" Or drugmeasurement(z%).ListIndex = -1 Then
        '    If drugs(z%) > "" And Left$(drugs(z%), 1) <> "X" And Left$(drugs(z%), 1) <> "U" Then
        '        msg = MsgBox("Drug Quantity and Measurement Type must be entered/selected.", 48, "Genesis Error Log")
        '        GoTo exiteditv
        '    End If
        'End If
        '===== Error 366
        If drugamt(Z%) > "" Then
            If drugs(Z%) = "" Or drugmeasurement(Z%).ListIndex = -1 Then
                msg = MsgBox("If a drug quantity is entered, then drug type and measurement type must also be entered.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(drugamt(Z%))
                drugamt(Z%).SetFocus
                GoTo exiteditv
            End If
        End If
        '===== Error 367
        If drugmeasurement(Z%).ListIndex > -1 Then
            If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "NP=)") > 0 Then
                If Left$(drugs(Z%), 1) <> "E" And Left$(drugs(Z%), 1) <> "G" And Left$(drugs(Z%), 1) <> "K" Then
                    msg = MsgBox("Number of Plants is only allowabled with Marijuana, Opium, and Other Hallucinogens.", 48, "Genesis Error Log")
                    Call ShowApplicableContainers(drugmeasurement(Z%))
                    drugmeasurement(Z%).SetFocus
                    GoTo exiteditv
                End If
            End If
            '===== Error 384
            If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "XX=") > 0 Then
                If Val(drugamt(Z%)) <> 1 Then
                    msg = MsgBox("If drug measurement is NOT REPORTED, drug amount must be 1.", 48, "Genesis Error Log")
                    Call ShowApplicableContainers(drugamt(Z%))
                    drugamt(Z%).SetFocus
                    GoTo exiteditv
                End If
            End If
            '===== Error 368
            If drugs(Z%) = "" Or drugamt(Z%) = "" Then
                msg = MsgBox("If a drug measurement is entered, then drug type and quantity must also be entered.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(drugamt(Z%))
                drugamt(Z%).SetFocus
                GoTo exiteditv
            End If
        End If
        '===== Error 362
        If Left$(drugs(Z%), 1) = "X" Then
            If drugtype(0).ListIndex = -1 Or drugtype(1).ListIndex = -1 Or drugtype(2).ListIndex = -1 Then
                msg = MsgBox("If X = Over 3 Drug Types is entered, then the other 2 drug types must also be entered.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(drugtype(0))
                drugtype(0).SetFocus
                GoTo exiteditv
            End If
            '===== Error 363
            If drugamt(Z%) > "" Or drugmeasurement(Z%).ListIndex > -1 Then
                msg = MsgBox("Drug Amount and Measurement are not valid entries for X=Over 3 Drug Types.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(drugamt(Z%))
                drugamt(Z%).SetFocus
                GoTo exiteditv
            End If
        End If
    Next Z%
End If
        
'==== Mandatories E - 25 = GIVEN
'==== Mandatories E - 26, 27, 28
'===== Error 453
If individual Or policeofficer Then
    If (Val(age(1)) = 0 And age(1) <> "00") Then
        msg = MsgBox("Invalid age entered.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(age(1))
        age(1).SetFocus
        GoTo exiteditv
    End If
    If sex(1).ListIndex = -1 Then
        msg = MsgBox("Invalid sex entered.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(sex(1))
        sex(1).SetFocus
        GoTo exiteditv
    End If
    If race(1).ListIndex = -1 Then
        msg = MsgBox("Invalid race entered.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(race(1))
        race(1).SetFocus
        GoTo exiteditv
    End If
    If ethnicity(1).ListIndex = -1 Then
        msg = MsgBox("Ethnicity is a required entry.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(ethnicity(1))
        ethnicity(1).SetFocus
        GoTo exiteditv
    End If
    If resident(1).ListIndex = -1 Then
        msg = MsgBox("Resident Status is a required entry.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(resident(1))
        resident(1).SetFocus
        GoTo exiteditv
    End If
Else
    '===== Error 458
    If Val(age(1)) > 0 Then
        msg = MsgBox("Age is not a valid entry for Victim if Type of Victim is not Individual.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(age(1))
        age(1).SetFocus
        GoTo exiteditv
    End If
    If sex(1).ListIndex > -1 And Left$(sex(1).List(sex(1).ListIndex), 1) <> "U" Then
        msg = MsgBox("Sex is not a valid entry for Victim if Type of Victim is not Individual.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(sex(1))
        sex(1).SetFocus
        GoTo exiteditv
    End If
    If race(1).ListIndex > -1 And Left$(race(1).List(race(1).ListIndex), 1) <> "U" Then
        msg = MsgBox("Race is not a valid entry for Victim if Type of Victim is not Individual.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(race(1))
        race(1).SetFocus
        GoTo exiteditv
    End If
    If ethnicity(1).ListIndex > -1 And Left$(ethnicity(1).List(ethnicity(1).ListIndex), 1) <> "U" Then
        msg = MsgBox("Ethnicity is not a valid entry for Victim if Type of Victim is not Individual.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(ethnicity(1))
        ethnicity(1).SetFocus
        GoTo exiteditv
    End If
    If resident(1).ListIndex > -1 And Left$(resident(1).List(resident(1).ListIndex), 1) <> "U" Then
        msg = MsgBox("Resident Status is not a valid entry for Victim if Type of Victim is not Individual.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(resident(1))
        resident(1).SetFocus
        GoTo exiteditv
    End If
End If
'===== Data Element 24
'===== Error 401
If vucrlist.ListItems.Count = 1 Then
    vucrlist.ListItems(1).Selected = True
End If
FOUNDVUCR = 0
For t% = 1 To vucrlist.ListItems.Count
    If vucrlist.ListItems(t%).Selected Then
        FOUNDVUCR = 1
        t% = vucrlist.ListItems.Count
    End If
Next t%
If FOUNDVUCR = 0 Then
    fromfind = 1
    Call Command1_Click
    fromfind = 0
    For t% = 1 To vucrlist.ListItems.Count
        If vucrlist.ListItems(t%).Selected Then
            FOUNDVUCR = 1
            t% = vucrlist.ListItems.Count
        End If
    Next t%
End If
If FOUNDVUCR = 0 Then
    msg = MsgBox("At least one UCR code must be connected to the victim.", 48, "Genesis Error Log")
    Call ShowApplicableContainers(vucrlist)
    vucrlist.SetFocus
    GoTo exiteditv
End If
'===== SCEdit 3/5/97 Attachment C
For t% = 1 To vucrlist.ListItems.Count
    If vucrlist.ListItems(t%).Selected Then
        tempvucr = Mid$(vucrlist.ListItems(t%), InStr(vucrlist.ListItems(t%), "(") + 1, 3)
        Select Case tempvucr
            Case "09A"
                For tt% = 1 To vucrlist.ListItems.Count
                    If tt% <> t% Then
                        If vucrlist.ListItems(tt%).Selected Then
                            tempvucr2 = Mid$(vucrlist.ListItems(tt%), InStr(vucrlist.ListItems(tt%), "(") + 1, 3)
                            Select Case tempvucr2
                                Case "09B", "13A", "13B", "13C"
                                    msg = MsgBox("Invalid UCR Combination for Victim", 48, "Genesis Error Log")
                                    Call ShowApplicableContainers(vucrlist)
                                    vucrlist.SetFocus
                                    GoTo exiteditv
                            End Select
                        End If
                    End If
                Next tt%
            Case "09B"
                For tt% = 1 To vucrlist.ListItems.Count
                    If tt% <> t% Then
                        If vucrlist.ListItems(tt%).Selected Then
                            tempvucr2 = Mid$(vucrlist.ListItems(tt%), InStr(vucrlist.ListItems(tt%), "(") + 1, 3)
                            Select Case tempvucr2
                                Case "09A", "13A", "13B", "13C"
                                    msg = MsgBox("Invalid UCR Combination for Victim", 48, "Genesis Error Log")
                                    Call ShowApplicableContainers(vucrlist)
                                    vucrlist.SetFocus
                                    GoTo exiteditv
                            End Select
                        End If
                    End If
                Next tt%
            Case "11A"
                For tt% = 1 To vucrlist.ListItems.Count
                    If tt% <> t% Then
                        If vucrlist.ListItems(tt%).Selected Then
                            tempvucr2 = Mid$(vucrlist.ListItems(tt%), InStr(vucrlist.ListItems(tt%), "(") + 1, 3)
                            Select Case tempvucr2
                                Case "11D", "13A", "13B", "13C", "36A", "36B"
                                    msg = MsgBox("Invalid UCR Combination for Victim", 48, "Genesis Error Log")
                                    Call ShowApplicableContainers(vucrlist)
                                    vucrlist.SetFocus
                                    GoTo exiteditv
                            End Select
                        End If
                    End If
                Next tt%
            Case "11B"
                For tt% = 1 To vucrlist.ListItems.Count
                    If tt% <> t% Then
                        If vucrlist.ListItems(tt%).Selected Then
                            tempvucr2 = Mid$(vucrlist.ListItems(tt%), InStr(vucrlist.ListItems(tt%), "(") + 1, 3)
                            Select Case tempvucr2
                                Case "11D", "13A", "13B", "13C", "36A", "36B"
                                    msg = MsgBox("Invalid UCR Combination for Victim", 48, "Genesis Error Log")
                                    Call ShowApplicableContainers(vucrlist)
                                    vucrlist.SetFocus
                                    GoTo exiteditv
                            End Select
                        End If
                    End If
                Next tt%
            Case "11C"
                For tt% = 1 To vucrlist.ListItems.Count
                    If tt% <> t% Then
                        If vucrlist.ListItems(tt%).Selected Then
                            tempvucr2 = Mid$(vucrlist.ListItems(tt%), InStr(vucrlist.ListItems(tt%), "(") + 1, 3)
                            Select Case tempvucr2
                                Case "11D", "13A", "13B", "13C", "36A", "36B"
                                    msg = MsgBox("Invalid UCR Combination for Victim", 48, "Genesis Error Log")
                                    Call ShowApplicableContainers(vucrlist)
                                    vucrlist.SetFocus
                                    GoTo exiteditv
                            End Select
                        End If
                    End If
                Next tt%
            Case "11D"
                For tt% = 1 To vucrlist.ListItems.Count
                    If tt% <> t% Then
                        If vucrlist.ListItems(tt%).Selected Then
                            tempvucr2 = Mid$(vucrlist.ListItems(tt%), InStr(vucrlist.ListItems(tt%), "(") + 1, 3)
                            Select Case tempvucr2
                                Case "11A", "11B", "11C", "13A", "13B", "13C", "36A", "36B"
                                    msg = MsgBox("Invalid UCR Combination for Victim", 48, "Genesis Error Log")
                                    Call ShowApplicableContainers(vucrlist)
                                    vucrlist.SetFocus
                                    GoTo exiteditv
                            End Select
                        End If
                    End If
                Next tt%
            Case "120"
                For tt% = 1 To vucrlist.ListItems.Count
                    If tt% <> t% Then
                        If vucrlist.ListItems(tt%).Selected Then
                            tempvucr2 = Mid$(vucrlist.ListItems(tt%), InStr(vucrlist.ListItems(tt%), "(") + 1, 3)
                            Select Case tempvucr2
                                Case "13A", "13B", "13C", "23A", "23B", "23C", "23D", "23E", "23F", "23G", "23H", "240"
                                    msg = MsgBox("Invalid UCR Combination for Victim", 48, "Genesis Error Log")
                                    Call ShowApplicableContainers(vucrlist)
                                    vucrlist.SetFocus
                                    GoTo exiteditv
                            End Select
                        End If
                    End If
                Next tt%
            Case "13A"
                For tt% = 1 To vucrlist.ListItems.Count
                    If tt% <> t% Then
                        If vucrlist.ListItems(tt%).Selected Then
                            tempvucr2 = Mid$(vucrlist.ListItems(tt%), InStr(vucrlist.ListItems(tt%), "(") + 1, 3)
                            Select Case tempvucr2
                                Case "09A", "09B", "11A", "11B", "11C", "120", "13B", "13C"
                                    msg = MsgBox("Invalid UCR Combination for Victim", 48, "Genesis Error Log")
                                    Call ShowApplicableContainers(vucrlist)
                                    vucrlist.SetFocus
                                    GoTo exiteditv
                            End Select
                        End If
                    End If
                Next tt%
            Case "13B"
                For tt% = 1 To vucrlist.ListItems.Count
                    If tt% <> t% Then
                        If vucrlist.ListItems(tt%).Selected Then
                            tempvucr2 = Mid$(vucrlist.ListItems(tt%), InStr(vucrlist.ListItems(tt%), "(") + 1, 3)
                            Select Case tempvucr2
                                Case "09A", "09B", "11A", "11B", "11C", "120", "13A", "13C", "11D"
                                    msg = MsgBox("Invalid UCR Combination for Victim", 48, "Genesis Error Log")
                                    Call ShowApplicableContainers(vucrlist)
                                    vucrlist.SetFocus
                                    GoTo exiteditv
                            End Select
                        End If
                    End If
                Next tt%
            Case "13C"
                For tt% = 1 To vucrlist.ListItems.Count
                    If tt% <> t% Then
                        If vucrlist.ListItems(tt%).Selected Then
                            tempvucr2 = Mid$(vucrlist.ListItems(tt%), InStr(vucrlist.ListItems(tt%), "(") + 1, 3)
                            Select Case tempvucr2
                                Case "09A", "09B", "11A", "11B", "11C", "120", "13A", "13B", "11D"
                                    msg = MsgBox("Invalid UCR Combination for Victim", 48, "Genesis Error Log")
                                    Call ShowApplicableContainers(vucrlist)
                                    vucrlist.SetFocus
                                    GoTo exiteditv
                            End Select
                        End If
                    End If
                Next tt%
            Case "23A", "23B", "23C", "23D", "23E", "23F", "23G", "23H", "240"
                For tt% = 1 To vucrlist.ListItems.Count
                    If tt% <> t% Then
                        If vucrlist.ListItems(tt%).Selected Then
                            tempvucr2 = Mid$(vucrlist.ListItems(tt%), InStr(vucrlist.ListItems(tt%), "(") + 1, 3)
                            Select Case tempvucr2
                                Case "120"
                                    msg = MsgBox("Invalid UCR Combination for Victim", 48, "Genesis Error Log")
                                    Call ShowApplicableContainers(vucrlist)
                                    vucrlist.SetFocus
                                    GoTo exiteditv
                            End Select
                        End If
                    End If
                Next tt%
            Case "36A", "36B"
                For tt% = 1 To vucrlist.ListItems.Count
                    If tt% <> t% Then
                        If vucrlist.ListItems(tt%).Selected Then
                            tempvucr2 = Mid$(vucrlist.ListItems(tt%), InStr(vucrlist.ListItems(tt%), "(") + 1, 3)
                            Select Case tempvucr2
                                Case "11A", "11B", "11C", "11D"
                                    msg = MsgBox("Invalid UCR Combination for Victim", 48, "Genesis Error Log")
                                    Call ShowApplicableContainers(vucrlist)
                                    vucrlist.SetFocus
                                    GoTo exiteditv
                            End Select
                        End If
                    End If
                Next tt%
        End Select
    End If
Next t%
'----edit by Ed Sloan 11/30/99 from SCLED 8/9/95 update----------------
For t% = 1 To vucrlist.ListItems.Count
    If vucrlist.ListItems(t%).Selected Then
        tempvucr = Mid$(vucrlist.ListItems(t%), InStr(vucrlist.ListItems(t%), "(") + 1, 3)
        '===== Error 464,465,467
        Select Case tempvucr
                Case "13A", "13B", "13C", "09A", "09B", "09C", "100", "11A", "11B", "11C", "11D", "36A", "36B", "36C"
                    If Not individual And Not policeofficer Then
                        msg = MsgBox("Individual must be selected for Crimes Against Person.", 48, "Genesis Error Log")
                        Call ShowApplicableContainers(individual)
                        individual.SetFocus
                        GoTo exiteditv
                    End If
                Case "90F"
                    If Not individual And Not policeofficer And Not societypublic Then
                        msg = MsgBox("Individual or Society must be selected for Family Offenses/Nonviolent.", 48, "Genesis Error Log")
                        Call ShowApplicableContainers(individual)
                        individual.SetFocus
                        GoTo exiteditv
                    End If
                Case "90Z"
                Case "90B", "90C", "90D", "90G", "90H", "90I", "90E", "90F"
                Case "90J", "36C", "980", "978", "753", "756"
                Case "35A", "35B", "39A", "39B", "39C", "39D", "370", "40A", "40B", "520"
                    If Not societypublic Then
                        msg = MsgBox("Society must be selected for Crimes Against Society.", 48, "Genesis Error Log")
                        Call ShowApplicableContainers(societypublic)
                        societypublic.SetFocus
                        GoTo exiteditv
                    End If
                Case Else
                    If societypublic Then
                        msg = MsgBox("Society cannot be selected for Crimes Against Property.", 48, "Genesis Error Log")
                        Call ShowApplicableContainers(societypublic)
                        societypublic.SetFocus
                        GoTo exiteditv
                    End If
        End Select
        '===== Error 481
        If tempvucr = "36B" And Val(age(1)) > 15 And vucrlist.ListItems(t%).Selected = True Then
            msg = MsgBox("For statutory rape, the victim must be less than or equal to 15 years of age.", 48, "Genesis Error Log")
            vucrlist.ListItems(t%).Selected = False
            Call ShowApplicableContainers(vucrlist)
            vucrlist.SetFocus
            GoTo exiteditv
        End If
        '===== SCEdit 8/9/95
        If tempvucr = "23C" Then
            If Len(age(1)) = 4 Then
                If Val(Right$(age(1), 2)) > 15 Then
                    msg = MsgBox("For Offense 23C, the victim age must be 15 years old or less.", 48, "Genesis Error Log")
                    Call ShowApplicableContainers(age(1))
                    age(1).SetFocus
                    GoTo exiteditv
                End If
            Else
            If Len(age(1)) = 2 Then
                If Val(age(1)) > 15 Then
                    msg = MsgBox("For Offense 23C, the victim age must be 15 years old or less.", 48, "Genesis Error Log")
                    Call ShowApplicableContainers(age(1))
                    age(1).SetFocus
                    GoTo exiteditv
                End If
            Else
                msg = MsgBox("For Offense 23C, the victim age must be 15 years old or less.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(age(1))
                age(1).SetFocus
                GoTo exiteditv
            End If
            End If
        End If
    End If
Next t%
'----------------------------------------------------------------------
'----edit by Ed Sloan 12/03/99 from SCLED 3/5/97 update----------------
Dim ucrexists, valinj As Boolean
ucrexists = False
valinj = False
FOUNDINJTYPE% = 0
For r% = 1 To vucrlist.ListItems.Count
    If vucrlist.ListItems(r%).Selected Then
        tempvucr = Mid$(vucrlist.ListItems(r%), InStr(vucrlist.ListItems(r%), "(") + 1, 3)
        Select Case tempvucr
            Case "100", "11A", "11B", "11C", "11D", "120", "13A", "13B", "210"
                FOUNDINJTYPE% = 1
        End Select
    End If
Next r%
For r% = 1 To vucrlist.ListItems.Count
    If vucrlist.ListItems(r%).Selected Then
        tempvucr = Mid$(vucrlist.ListItems(r%), InStr(vucrlist.ListItems(r%), "(") + 1, 3)
        ICT% = 0
        For rr% = 1 To injury.ListItems.Count
            If injury.ListItems(rr%).Selected Then
                ICT% = ICT% + 1
            End If
        Next rr%
        Select Case tempvucr
            '===== Error 479
            Case "13B"
                ucrexists = True
                For q% = 1 To injury.ListItems.Count
                    If injury.ListItems(q%).Selected Then
                        tempinj = Mid$(injury.ListItems(q%), InStr(injury.ListItems(q%), "(") + 1, 1)
                        If (tempinj = "M" Or tempinj = "N") Then
                            validinj = True
                        End If
                    End If
                Next q%
                If ucrexists And Not validinj Then
                    msg = MsgBox("For simple assault, the only injury types can be minor or none.", 48, "Genesis Error Log")
                    Call ShowApplicableContainers(vucrlist)
                    vucrlist.SetFocus
                    GoTo exiteditv
                End If
            '===== Error 401
            Case "100", "11A", "11B", "11C", "11D", "120", "13A", "13B", "210"
                If ICT% = 0 Then
                    msg = MsgBox("Type of injury must be selected for UCR " + tempvcur + ".", 48, "Genesis Error Log")
                    Call ShowApplicableContainers(vucrlist)
                    vucrlist.SetFocus
                    GoTo exiteditv
                End If
            '===== Error 419
            Case Else
                If ICT% > 0 And FOUNDINJTYPE% = 0 Then
                    foundother% = 0
                    For q% = 1 To vucrlist.ListItems.Count
                        If vucrlist.ListItems(q%).Selected Then
                            Select Case Mid$(vucrlist.ListItems(q%), InStr(vucrlist.ListItems(q%), "(") + 1, 3)
                                Case "100", "11A", "11B", "11C", "11D", "120", "13A", "13B", "210"
                                    foundother% = 1
                                    q% = vucrlist.ListItems.Count
                            End Select
                        End If
                    Next q%
                    If foundother% = 0 Then
                        msg = MsgBox("Type of injury is not applicable for UCR " + tempvcur + ".", 48, "Genesis Error Log")
                        Call ShowApplicableContainers(injury)
                        injury.SetFocus
                        GoTo exiteditv
                    End If
                End If
        End Select
    End If
Next r%

'===== Error 407
For r% = 1 To injury.ListItems.Count
    If injury.ListItems(r%).Selected Then
        If Mid$(injury.ListItems(r%), InStr(injury.ListItems(r%), "(") + 1, 1) = "N" Then
            For rr% = 1 To r% - 1
                If injury.ListItems(rr%).Selected Then
                    msg = MsgBox("When Injury Type N=None is selected, no other values may be selected.", 48, "Genesis Error Log")
                    Call ShowApplicableContainers(injury)
                    injury.SetFocus
                    GoTo exiteditv
                End If
            Next rr%
            For rr% = r% + 1 To injury.ListItems.Count
                If injury.ListItems(rr%).Selected Then
                    msg = MsgBox("When Injury Type N=None is selected, no other values may be selected.", 48, "Genesis Error Log")
                    Call ShowApplicableContainers(injury)
                    injury.SetFocus
                    GoTo exiteditv
                End If
            Next rr%
        End If
    End If
Next r%
For t% = 0 To 9
    ucrlist(t%).ListIndex = -1
    For tt% = 0 To ucrlist(t%).ListCount - 1
        If ucrlist(t%).Selected(tt%) = True Then
            ucrlist(t%).ListIndex = tt%
            tt% = ucrlist(t%).ListCount - 1
        End If
    Next tt%
    If ucrlist(t%).ListIndex > -1 Then
        tempucr = Mid$(ucrlist(t%).List(ucrlist(t%).ListIndex), InStr(ucrlist(t%).List(ucrlist(t%).ListIndex), "(") + 1, 3)
        If tempucr = "13A" Or tempucr = "13B" Or tempucr = "13C" Then
            If UCase(vsname(2)) <> "UNKNOWN" Then
                If relationship(10).ListIndex = -1 And relationship(11).ListIndex = -1 And relationship(12).ListIndex = -1 And relationship(13).ListIndex = -1 And relationship(14).ListIndex = -1 And relationship(15).ListIndex = -1 And relationship(16).ListIndex = -1 And relationship(17).ListIndex = -1 And relationship(18).ListIndex = -1 And relationship(19).ListIndex = -1 Then
                    msg = MsgBox("A relationship to subject must be selected.", 48, "Genesis Error Log")
                    Call ShowApplicableContainers(ucrlist(t%))
                    ucrlist(t%).SetFocus
                    GoTo exiteditv
                End If
            End If
        End If
        '===== Additional F 12
        '===== Data Element 7, 13
        '===== eRROR 458
        Select Case tempucr
            Case "09A", "09B", "09C", "100", "11A", "11B", "11C", "11D", "120", "13A", "13B", "13C", "36A", "36B"
                If UCase(vsname(2)) <> "UNKNOWN" Then
                    If relationship(10).ListIndex = -1 And relationship(11).ListIndex = -1 And relationship(12).ListIndex = -1 And relationship(13).ListIndex = -1 And relationship(14).ListIndex = -1 And relationship(15).ListIndex = -1 And relationship(16).ListIndex = -1 And relationship(17).ListIndex = -1 And relationship(18).ListIndex = -1 And relationship(19).ListIndex = -1 Then
                        msg = MsgBox("A relationship to subject must be selected.", 48, "Genesis Error Log")
                        Call ShowApplicableContainers(relationship(10))
                        relationship(10).SetFocus
                        GoTo exiteditv
                    End If
                End If
        End Select
        '===== Additional F 13
        '===== Data Element 7
        If tempucr = "100" Then
            If UCase(vsname(2)) <> "UNKNOWN" Then
                If relationship(10).ListIndex = -1 And relationship(11).ListIndex = -1 And relationship(12).ListIndex = -1 And relationship(13).ListIndex = -1 And relationship(14).ListIndex = -1 And relationship(15).ListIndex = -1 And relationship(16).ListIndex = -1 And relationship(17).ListIndex = -1 And relationship(18).ListIndex = -1 And relationship(19).ListIndex = -1 Then
                    msg = MsgBox("A relationship to subject must be selected.", 48, "Genesis Error Log")
                    Call ShowApplicableContainers(relationship(10))
                    relationship(10).SetFocus
                    GoTo exiteditv
                End If
            End If
        End If
        '===== Additional F 18
        If tempucr = "120" Then
            If individual Then
            If UCase(vsname(2)) <> "UNKNOWN" Then
                If relationship(10).ListIndex = -1 And relationship(11).ListIndex = -1 And relationship(12).ListIndex = -1 And relationship(13).ListIndex = -1 And relationship(14).ListIndex = -1 And relationship(15).ListIndex = -1 And relationship(16).ListIndex = -1 And relationship(17).ListIndex = -1 And relationship(18).ListIndex = -1 And relationship(19).ListIndex = -1 Then
                    msg = MsgBox("A relationship to subject must be selected.", 48, "Genesis Error Log")
                    Call ShowApplicableContainers(relationship(10))
                    relationship(10).SetFocus
                    GoTo exiteditv
                End If
            End If
            End If
        End If
        '===== Additional F 19
        If tempucr = "11A" Or tempucr = "11B" Or tempucr = "11C" Or tempucr = "11D" Then
            If UCase(vsname(2)) <> "UNKNOWN" Then
                If relationship(10).ListIndex = -1 And relationship(11).ListIndex = -1 And relationship(12).ListIndex = -1 And relationship(13).ListIndex = -1 And relationship(14).ListIndex = -1 And relationship(15).ListIndex = -1 And relationship(16).ListIndex = -1 And relationship(17).ListIndex = -1 And relationship(18).ListIndex = -1 And relationship(19).ListIndex = -1 Then
                    msg = MsgBox("A relationship to subject must be selected.", 48, "Genesis Error Log")
                    Call ShowApplicableContainers(relationship(10))
                    relationship(10).SetFocus
                    GoTo exiteditv
                End If
            End If
        End If
        '===== Additional F 20
        If tempucr = "36A" Or tempucr = "36B" Then
            If UCase(vsname(2)) <> "UNKNOWN" Then
                If relationship(10).ListIndex = -1 And relationship(11).ListIndex = -1 And relationship(12).ListIndex = -1 And relationship(13).ListIndex = -1 And relationship(14).ListIndex = -1 And relationship(15).ListIndex = -1 And relationship(16).ListIndex = -1 And relationship(17).ListIndex = -1 And relationship(18).ListIndex = -1 And relationship(19).ListIndex = -1 Then
                    msg = MsgBox("A relationship to subject must be selected.", 48, "Genesis Error Log")
                    Call ShowApplicableContainers(relationship(10))
                    relationship(10).SetFocus
                    GoTo exiteditv
                End If
            End If
        End If
    End If
Next t%

GoTo goodeditv
exiteditv:
editerr = 1
goodeditv:
Exit Sub
'RLB Bandaid
rlbErr:
    If Err.Number = 5 Then Resume Next
End Sub
Public Sub editsubject(editerr, typeedit As Integer)
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwi + "incident.mdb")
'===== SC LEOKA
If policeofficer Then
    If TWOMANVEHICLE = 0 And ONEMANVEHICLE = 0 And DETECTIVE = 0 And TODOTHER = 0 Then
        msg = MsgBox("If Type Victim is Police Officer, a selection must be made for Two Man Vehicle, One Man Vehicle, Detective/Special Assignment, or Other.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(policeofficer)
        policeofficer.SetFocus
        GoTo exitedits
    End If
    If Not TWOMANVEHICLE Then
        If ALONE = 0 And ASSISTED = 0 Then
            msg = MsgBox("If Type Victim is Police Officer and not Two Man Vehicle, a selection must be made for Alone or Assisted.", 48, "Genesis Error Log")
            Call ShowApplicableContainers(TWOMANVEHICLE)
            TWOMANVEHICLE.SetFocus
            GoTo exitedits
        End If
    End If
End If
'===== Error 761
If RUNAWAY = 1 Then
    If Len(age(2)) = 2 Then
        If Val(age(2)) > 17 Then
            msg = MsgBox("A runaway must be under the age of 18.", 48, "Genesis Error Log")
            Call ShowApplicableContainers(RUNAWAY)
            RUNAWAY.SetFocus
            GoTo exitedits
        End If
    Else
    If Len(age(2)) = 4 Then
        If Val(Right$(age(2), 2)) > 17 Then
            msg = MsgBox("A runaway must be under the age of 18.", 48, "Genesis Error Log")
            Call ShowApplicableContainers(age(2))
            age(2).SetFocus
            GoTo exitedits
        End If
    Else
        msg = MsgBox("A runaway must be under the age of 18.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(age(2))
        age(2).SetFocus
        GoTo exitedits
    End If
    End If
End If
'===== Error 665
If IsDate(DATEOFARREST) And IsDate(incidentdate(0)) Then
    If CVDate(incidentdate(0)) > CVDate(DATEOFARREST) Then
        msg = MsgBox("Date of Arrest cannot be before incident date.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(incidentdate(0))
        incidentdate(0).SetFocus
        GoTo exitedits
    End If
End If
If age(2) > "" Then
    If Val(age(2)) = 0 And age(2) <> "00" Then
        msg = MsgBox("Age must be entered in the format of a single age: X or XX (i.e. 3 or 42) or in the format of a range of ages: XXXX (i.e. 2025, meaning 20 - 25 years old).", 48, "Genesis Error Log")
        Call ShowApplicableContainers(age(2))
        age(2).SetFocus
        GoTo exitedits
    End If
Else
    '===== error 504
    msg = MsgBox("Subject age must be entered. (00 = unknown)", 48, "Genesis Error Log")
    Call ShowApplicableContainers(age(2))
    age(2).SetFocus
    GoTo exitedits
End If
'===== Error 504
If sex(2).ListIndex = -1 Then
    msg = MsgBox("A value for Sex in subject data must be entered.", 48, "Genesis Error Log")
    Call ShowApplicableContainers(sex(2))
    sex(2).SetFocus
    GoTo exitedits
End If
If race(2).ListIndex = -1 Then
    msg = MsgBox("A value for race in subject data must be entered.", 48, "Genesis Error Log")
    Call ShowApplicableContainers(race(2))
    race(2).SetFocus
    GoTo exitedits
End If
If ethnicity(2).ListIndex = -1 Then
    msg = MsgBox("A value for ethnicity in subject data must be entered.", 48, "Genesis Error Log")
    Call ShowApplicableContainers(ethnicity(2))
    ethnicity(2).SetFocus
    GoTo exitedits
End If
'===== Error 404,556
If age(2) <> "NN" And age(2) <> "NB" And age(2) <> "BB" Then
    For t% = 1 To Len(age(2))
        If InStr("0123456789-", Mid$(age(2), t%, 1)) = 0 Then
            msg = MsgBox("An invalid entry in age has been found.  Valid Entry Formats:  X, XX, X-Y, XX-Y, X-YY, XX-YY, XXYY", 48, "Genesis Error Log")
            t% = Len(age(2))
            Call ShowApplicableContainers(age(2))
            age(2).SetFocus
            GoTo exitedits
        End If
    Next t%
End If
If InStr(age(2), "-") > 0 Then
    ag1 = Val(Left$(age(2), InStr(age(2), "-") - 1))
    ag2 = Val(Mid$(age(2), InStr(age(2), "-") + 1))
    age(2) = Format$(ag1, "00") + Format$(ag2, "00")
End If
'===== Error 410,422,509,510,522
If Len(age(2)) = 4 Then
    If Val(Left$(age(2), 2)) >= Val(Right$(age(2), 2)) Then
        msg = MsgBox("For an age range, the first age must be less than the second age.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(age(2))
        GoTo exitedits
    End If
    If Val(Left$(age(2), 2)) = 0 Then
        msg = MsgBox("The low value in an age range cannot be 0.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(age(2))
        age(2).SetFocus
        GoTo exitedits
    End If
End If
editsubjectdetail:
'===== Data Element 20, 21, 22
Dim drugs(5)
drugs(0) = ""
drugs(1) = ""
drugs(2) = ""
drugs(3) = ""
drugs(4) = ""
drugs(5) = ""
dct% = 3
For ii% = 3 To 5
    If drugtype(ii%).ListIndex > -1 Then
        drugs(dct%) = drugtype(ii%).List(drugtype(ii%).ListIndex)
        dct% = dct% + 1
    End If
Next ii%
If dct% > 0 And dct% > 3 Then
    For Z% = 3 To dct%
        '===== Error 306
        If Z% > 0 Then
            For zz% = 0 To Z% - 1
                If drugs(zz%) = drugs(Z%) Then
                    If InStr(drugmeasurement(zz%).List(drugmeasurement(zz%).ListIndex), "GM=") > 0 Or _
                       InStr(drugmeasurement(zz%).List(drugmeasurement(zz%).ListIndex), "KG=") > 0 Or _
                       InStr(drugmeasurement(zz%).List(drugmeasurement(zz%).ListIndex), "OZ=") > 0 Or _
                       InStr(drugmeasurement(zz%).List(drugmeasurement(zz%).ListIndex), "LB=") > 0 Then
                        If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "GM=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "KG=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "OZ=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "LB=") > 0 Then
                                msg = MsgBox("Two different measurements within the same weight category cannot be entered for the same type of drug.", 48, "Genesis Error Log")
                                Call ShowApplicableContainers(drugmeasurement(zz%))
                                drugmeasurement(zz%).SetFocus
                                GoTo exitedits
                        End If
                    End If
                    If InStr(drugmeasurement(zz%).List(drugmeasurement(zz%).ListIndex), "ML=") > 0 Or _
                       InStr(drugmeasurement(zz%).List(drugmeasurement(zz%).ListIndex), "LT=") > 0 Or _
                       InStr(drugmeasurement(zz%).List(drugmeasurement(zz%).ListIndex), "FO=") > 0 Or _
                       InStr(drugmeasurement(zz%).List(drugmeasurement(zz%).ListIndex), "GL=") > 0 Then
                        If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "ML=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "LT=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "FO=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "GL=") > 0 Then
                                msg = MsgBox("Two different measurements within the same weight category cannot be entered for the same type of drug.", 48, "Genesis Error Log")
                                Call ShowApplicableContainers(drugmeasurement(zz%))
                                drugmeasurement(zz%).SetFocus
                                GoTo exitedits
                        End If
                    End If
                    If InStr(drugmeasurement(zz%).List(drugmeasurement(zz%).ListIndex), "DU=") > 0 Or _
                       InStr(drugmeasurement(zz%).List(drugmeasurement(zz%).ListIndex), "NP=") > 0 Then
                        If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "DU=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "NP=") > 0 Then
                                msg = MsgBox("Two different measurements within the same weight category cannot be entered for the same type of drug.", 48, "Genesis Error Log")
                                Call ShowApplicableContainers(drugmeasurement(zz%))
                                drugmeasurement(zz%).SetFocus
                                GoTo exitedits
                        End If
                    End If
                End If
            Next zz%
        End If
        For t% = 1 To Len(drugamt(Z%))
            If InStr("0123456789.", Mid$(drugamt(Z%), t%, 1)) = 0 Then
                msg = MsgBox("Only 0123456789. are allowed in drug amount.  Enter all fractions as decimals (i.e. 1/2 = .5).", 48, "Genesis Error Log")
                Call ShowApplicableContainers(drugamt(Z%))
                drugamt(Z%).SetFocus
                GoTo exitedits
            End If
        Next t%
        '=====SCEdit 4/21/92 P29  Overturned drug amt and measurement error
        'If drugamt(z%) = "" Or drugmeasurement(z%).ListIndex = -1 Then
        '    If drugs(z%) > "" And Left$(drugs(z%), 1) <> "X" And Left$(drugs(z%), 1) <> "U" Then
        '        msg = MsgBox("Drug Quantity and Measurement Type must be entered/selected.", 48, "Genesis Error Log")
        '        GoTo exitedits
        '    End If
        'End If
        '===== Error 366
        If drugamt(Z%) > "" Then
            If drugs(Z%) = "" Or drugmeasurement(Z%).ListIndex = -1 Then
                msg = MsgBox("If a drug quantity is entered, then drug type and measurement type must also be entered.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(drugamt(Z%))
                drugamt(Z%).SetFocus
                GoTo exitedits
            End If
        End If
        '===== Error 367
        If drugmeasurement(Z%).ListIndex > -1 Then
            If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "NP=") > 0 Then
                If Left$(drugs(Z%), 1) <> "E" And Left$(drugs(Z%), 1) <> "G" And Left$(drugs(Z%), 1) <> "K" Then
                    msg = MsgBox("Number of Plants is only allowabled with Marijuana, Opium, and Other Hallucinogens.", 48, "Genesis Error Log")
                    Call ShowApplicableContainers(drugmeasurement(Z%))
                    drugmeasurement(Z%).SetFocus
                    GoTo exitedits
                End If
            End If
            '===== Error 384
            If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "XX=") > 0 Then
                If Val(drugamt(Z%)) <> 1 Then
                    msg = MsgBox("If drug measurement is NOT REPORTED, drug amount must be 1.", 48, "Genesis Error Log")
                    Call ShowApplicableContainers(drugamt(Z%))
                    drugamt(Z%).SetFocus
                    GoTo exitedits
                End If
            End If
            '===== Error 368
            If drugs(Z%) = "" Or drugamt(Z%) = "" Then
                msg = MsgBox("If a drug measurement is entered, then drug type and quantity must also be entered.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(drugamt(Z%))
                drugamt(Z%).SetFocus
                GoTo exitedits
            End If
        End If
        '===== Error 362
        If Left$(drugs(Z%), 1) = "X" Then
            If drugtype(3).ListIndex = -1 Or drugtype(4).ListIndex = -1 Or drugtype(5).ListIndex = -1 Then
                msg = MsgBox("If X = Over 3 Drug Types is entered, then the other 2 drug types must also be entered.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(drugtype(3))
                drugtype(3).SetFocus
                GoTo exitedits
            End If
            '===== Error 363
            If drugamt(Z%) > "" Or drugmeasurement(Z%).ListIndex > -1 Then
                msg = MsgBox("Drug Amount and Measurement are not valid entries for X=Over 3 Drug Types.", 48, "Genesis Error Log")
                Call ShowApplicableContainers(drugamt(Z%))
                drugamt(Z%).SetFocus
                GoTo exitedits
            End If
        End If
    Next Z%
End If
For pp% = 10 To 19
    temprel = Mid$(relationship(pp%).List(relationship(pp%).ListIndex), InStr(relationship(pp%).List(relationship(pp%).ListIndex), "(") + 1, 2)
    If temprel = "CH" Or temprel = "GC" Or temprel = "SC" Then
        If Val(age(2)) < Val(age(1)) Then
            msg = MsgBox("The relationship of victim to subject cannot be 'PA', 'GP' or 'SP' when victim's age is less than subject's age.", 48, "Genesis Error Log")
            Call ShowApplicableContainers(age(2))
            age(2).SetFocus
            GoTo exitedits
        End If
    End If
    If temprel = "PA" Or temprel = "GP" Or temprel = "SP" Then
        If Val(age(1)) < Val(age(2)) Then
            msg = MsgBox("The relationship of victim to subject cannot be 'PA', 'GP' or 'SP' when victim's age is less than subject's age.", 48, "Genesis Error Log")
            Call ShowApplicableContainers(age(1))
            age(1).SetFocus
            GoTo exitedits
        End If
    End If
Next pp%
'==== Mandatories E - 37, 38, 39
'===== Error 501
'If individual And UCase(vsname(2)) <> "UNKNOWN" Then
If UCase(vsname(2)) <> "UNKNOWN" Then
    If Val(age(2)) = 0 And age(2) <> "00" Then
        msg = MsgBox("Invalid age entered.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(age(2))
        age(2).SetFocus
        GoTo exitedits
    End If
    If sex(2).ListIndex = -1 Then
        msg = MsgBox("Invalid sex entered.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(sex(2))
        sex(2).SetFocus
        GoTo exitedits
    End If
    If race(2).ListIndex = -1 Then
        msg = MsgBox("Invalid race entered.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(race(2))
        race(2).SetFocus
        GoTo exitedits
    End If
    If ethnicity(2).ListIndex = -1 Then
        msg = MsgBox("Invalid ETHNICITY entered.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(ethnicity(2))
        ethnicity(2).SetFocus
        GoTo exitedits
    End If
End If
'==== Mandatories E - 48, 49
'===== Mandatories F - 41, 42, 43, 44, 45, 46, 47, 48, 49, 52

If DATEOFARREST > "" Then
    If Not IsDate(DATEOFARREST) Then
        msg = MsgBox("Valid date of arrest must be entered.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(DATEOFARREST)
        DATEOFARREST.SetFocus
        GoTo exitedits
    End If
    '===== Error 601,701
    If Val(age(2)) = 0 And age(2) <> "00" Then
        msg = MsgBox("Invalid age entered.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(age(2))
        age(2).SetFocus
        GoTo exitedits
    End If
    If sex(2).ListIndex = -1 Then
        msg = MsgBox("Invalid sex entered.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(sex(2))
        sex(2).SetFocus
        GoTo exitedits
    End If
    If race(2).ListIndex = -1 Then
        msg = MsgBox("Invalid race entered.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(race(2))
        race(2).SetFocus
        GoTo exitedits
    End If
    '===== Error 665
    If CVDate(incidentdate(0)) > CVDate(DATEOFARREST) Then
        msg = MsgBox("Date of Arrest cannot be before incident date.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(incidentdate(0))
        incidentdate(0).SetFocus
        GoTo exitedits
    End If
End If
'===== SCEdit 4/21/92 P28
If (offenderdeath Or noprosecution Or extraditiondenied Or victimdeclinescooperation Or juvenilenocustody) Then
    If age(2) = "00" Then
        msg = MsgBox("For an exceptional clearance, the subject's age (other than 00) must be selected.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(age(2))
        age(2).SetFocus
        GoTo exitedits
    End If
    If race(2).ListIndex = -1 Or race(2).List(race(2).ListIndex) = "Unknown" Then
        msg = MsgBox("For an exceptional clearance, the subject's race (other than unknown) must be selected.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(race(2))
        race(2).SetFocus
        GoTo exitedits
    End If
    If sex(2).ListIndex = -1 Or sex(2).List(sex(2).ListIndex) = "Unknown" Then
        msg = MsgBox("For an exceptional clearance, the subject's sex (other than unknown) must be selected.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(sex(2))
        sex(2).SetFocus
        GoTo exitedits
    End If
    If ethnicity(2).ListIndex = -1 Or ethnicity(2).List(ethnicity(2).ListIndex) = "Unknown" Then
        msg = MsgBox("For an exceptional clearance, the subject's ethnicity (other than unknown) must be selected.", 48, "Genesis Error Log")
        Call ShowApplicableContainers(ethnicity(2))
        ethnicity(2).SetFocus
        GoTo exitedits
    End If
End If
GoTo goodedits
exitedits:
editerr = 1
goodedits:
On Error Resume Next
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub


