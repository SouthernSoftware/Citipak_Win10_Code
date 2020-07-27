Dim fromfind, schanged As Integer, networkpath, nwl As String, ecc As Integer
Dim holdrecv As Integer, HI As String
Dim FROMKEY As Integer, BACKTAB As Integer
Dim tempsave As Integer, TV1, TV2 As String, tempword As String
Private Sub loadcodes()
Dim db As Database, rs As Recordset, ITMX As ListItem
On Error GoTo oderror
od:
Set db = OpenDatabase(networkpath + "incident.mdb")
Set rs = db.OpenRecordset("select * from codes")
On Error Resume Next
For t% = 0 To 29
    drugmeasurement(t%).Clear
    drugtype(t%).Clear
Next t%
For t% = 0 To 5
    group(t%).Clear
Next t%
For t% = 0 To 19
    relationship(t%).Clear
Next t%
reportingofficer(0).Clear
reportingofficer(1).Clear
approvingofficer.Clear
If rs.EOF Then
    db.Close
    Exit Sub
End If
rs.MoveFirst
While Not rs.EOF
    Select Case rs("type")
        Case "group"
            For t% = 0 To 5
                group(t%).AddItem rs("code")
            Next t%
        Case "drugtype"
            For t% = 0 To 29
                drugtype(t%).AddItem rs("code")
            Next t%
        Case "measure"
            For t% = 0 To 29
                drugmeasurement(t%).AddItem rs("code")
            Next t%
    End Select
    rs.MoveNext
Wend
Set rs = db.OpenRecordset("select code from relationship order by code")
If Not rs.EOF Then
    rs.MoveFirst
End If
While Not rs.EOF
    For tt% = 0 To 19
        relationship(tt%).AddItem rs("code")
    Next tt%
    rs.MoveNext
Wend
db.Close
On Error GoTo oderror2
od2:
Set db = OpenDatabase(nwl + "lawsuite.mdb")
Set rs = db.OpenRecordset("select profname from [Deputy Query]")
If Not rs.EOF Then
    rs.MoveFirst
End If
While Not rs.EOF
    If Not IsNull(rs("profname")) Then
        reportingofficer(0).AddItem rs("profname")
        reportingofficer(1).AddItem rs("profname")
        approvingofficer.AddItem rs("profname")
    End If
    rs.MoveNext
Wend
On Error Resume Next
db.Close
Exit Sub
oderror2:
If Err > 3200 Then
    Resume od2
Else
    Resume Next
End If
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub

Private Sub SENDCHAR(CH As String)
SendKeys CH
End Sub

Private Sub sendclosepara()
SendKeys "{)}"
End Sub
Private Sub senddash()
SendKeys "-"
End Sub

Private Sub sendopenpara()
SendKeys "{(}"
End Sub
Private Sub sendslash()
SendKeys "/"
End Sub

Private Sub sendspace()
SendKeys " "
End Sub

Private Sub address_GotFocus(Index As Integer)
If Me.ActiveControl.Top + Me.ActiveControl.Height > Picture1.Height - Picture2.Top Then
    VScroll1 = Me.ActiveControl.Top - Picture1.Height + Me.ActiveControl.Height + 100
End If
If address(Index) > "" Then
    Exit Sub
End If
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
Set rs = db.OpenRecordset("SELECT * FROM PEOPLE WHERE DPnameLF = " + Chr$(34) + vsname(Index) + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    If Not IsNull(rs("DPHADDRESS")) Then
        address(Index) = rs("DPHADDRESS")
    Else
        address(Index) = ""
    End If
End If
db.Close
On Error Resume Next
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub


Private Sub age_Click(Index As Integer)

If Index > 0 Then
    
End If
age(Index).Refresh
End Sub


Private Sub age_LostFocus(Index As Integer)
If age(Index) = "" Then
    age(Index) = "00"
End If
End Sub

Private Sub alcoholno_GotFocus(Index As Integer)
If alcoholframe(Index).Top + alcoholframe(Index).Height > Picture1.Height - Picture2.Top Then
    VScroll1 = alcoholframe(Index).Top - Picture1.Height + alcoholframe(Index).Height + 100
End If

End Sub

Private Sub alcoholunknown_GotFocus(Index As Integer)
If alcoholframe(Index).Top + alcoholframe(Index).Height > Picture1.Height - Picture2.Top Then
    VScroll1 = alcoholframe(Index).Top - Picture1.Height + alcoholframe(Index).Height + 100
End If

End Sub

Private Sub alcoholyes_GotFocus(Index As Integer)
If alcoholframe(Index).Top + alcoholframe(Index).Height > Picture1.Height - Picture2.Top Then
    VScroll1 = alcoholframe(Index).Top - Picture1.Height + alcoholframe(Index).Height + 100
End If

End Sub

Private Sub ARREST_GotFocus(Index As Integer)
If Me.ActiveControl.Top + Me.ActiveControl.Height > Picture1.Height - Picture2.Top Then
    VScroll1 = Me.ActiveControl.Top - Picture1.Height + Me.ActiveControl.Height + 100
End If

End Sub

Private Sub BIRTHDATE_LostFocus(Index As Integer)
If BIRTHDATE(Index) > "" And Not IsDate(BIRTHDATE(Index)) Then
    msg = MsgBox("Date/Time entered if invalid.", 48, "Genesis Error Log")
    BIRTHDATE(Index).SetFocus
End If
BIRTHDATE(Index) = Format$(BIRTHDATE(Index), "mm/dd/yyyy")
End Sub
Private Sub city_GotFocus(Index As Integer)
If vsname(Index) = "UNKNOWN" Then
    city(Index) = ""
End If
    
End Sub
Private Sub closedrugframes_Click(Index As Integer)
sdrugframe(Index).Visible = False
If Index > 3 Then
    totalvalue(Index - 2).SetFocus
End If

End Sub
Private Sub closevucrf_Click(Index As Integer)
vucrf(Index).Visible = False
complainant(Index).SetFocus
End Sub
Private Sub birthdate_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(BIRTHDATE(Index)) = 1 Or Len(BIRTHDATE(Index)) = 4 Then
    Call sendslash
End If
End If
End Sub
Private Sub Command1_Click(Index As Integer)
If fromfind = 1 Or Val(victim(Index)) = 0 Then
    Exit Sub
End If
vucrf(Index).Left = 500
vucrf(Index).Visible = True
vucrlist(Index).SetFocus
End Sub
Private Sub Command10_Click()
relationshipframe(1).Visible = False
resident(1).SetFocus
End Sub

Private Sub Command2_Click(Index As Integer)
If fromfind = 1 Then
    Exit Sub
End If
pucrlist(Index).Top = description(Index).Top - 1000
pucrlist(Index).Left = description(Index).Left + 1000
pucrlist(Index).Visible = True
pucrlist(Index).SetFocus
End Sub

Private Sub Command7_Click(Index As Integer)
If fromfind = 1 Then
    Exit Sub
End If
sdrugframe(Index).Left = 2000
sdrugframe(Index).Top = Command7(Index).Top
sdrugframe(Index).Visible = True
'drugtype(1 + (Index * 3)).SetFocus
End Sub

Private Sub Command8_Click(Index As Integer)
relationshipframe(Index).Visible = False
resident(Index).SetFocus
End Sub

Private Sub DATERECOVERED_GotFocus(Index As Integer)
If DATERECOVERED(Index) = "" Then
    DATERECOVERED(Index) = Format$(Date$, "mm/dd/yyyy")
End If

End Sub

Private Sub daterecovered_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(DATERECOVERED(Index)) = 1 Or Len(DATERECOVERED(Index)) = 4 Then
    Call sendslash
End If
End If
End Sub

Private Sub DATERECOVERED_LostFocus(Index As Integer)
If DATERECOVERED(Index) > "" And Not IsDate(DATERECOVERED(Index)) Then
    msg = MsgBox("Date/Time entered if invalid.", 48, "Genesis Error Log")
    DATERECOVERED(Index).SetFocus
End If
DATERECOVERED(Index) = Format$(DATERECOVERED(Index), "mm/dd/yyyy")
DATERECOVERED(Index).Visible = False
If DATERECOVERED(Index) = "Date" Then
    DATERECOVERED(Index) = ""
End If
'=====Data Item 18 and 19
If BACKTAB = 0 Then
If Mid$(pucrlist(Index).List(pucrlist(Index).ListIndex), InStr(pucrlist(Index).List(pucrlist(Index).ListIndex), "(") + 1, 3) = "240" Then
    tempgroup = Val(Mid$(group(Index).List(group(Index).ListIndex), InStr(group(Index).List(group(Index).ListIndex), "(") + 1, 2))
    If (tempgroup = 3 Or tempgroup = 5 Or tempgroup = 24 Or tempgroup = 28 Or tempgroup = 37) And fromfind = 0 Then
        numvehicle((Index Mod 6) + 6).Visible = True
        numvehicle((Index Mod 6) + 6).SetFocus
    Else
        totalvalue(holdrecv + 6).SetFocus
    End If
Else
    totalvalue(holdrecv + 6).SetFocus
End If
Else
    BACKTAB = 0
End If

End Sub
Private Sub description_GotFocus(Index As Integer)
If Me.ActiveControl.Top + Me.ActiveControl.Height > Picture1.Height - Picture2.Top Then
    VScroll1 = Me.ActiveControl.Top - Picture1.Height + Me.ActiveControl.Height + 100
End If

End Sub

Private Sub drugsno_Click(Index As Integer)

If drugsno(Index) Then
    For Y% = (Index * 3) To (Index * 3) + 2
        drugtype(Y%).ListIndex = -1
        drugamt(Y%) = ""
        drugmeasurement(Y%).ListIndex = -1
    Next Y%
End If
End Sub

Private Sub drugsno_GotFocus(Index As Integer)
If Frame11(Index).Top + Frame11(Index).Height > Picture1.Height - Picture2.Top Then
    VScroll1 = Frame11(Index).Top - Picture1.Height + Frame11(Index).Height + 100
End If

End Sub

Private Sub drugsunknown_GotFocus(Index As Integer)
If Frame11(Index).Top + Frame11(Index).Height > Picture1.Height - Picture2.Top Then
    VScroll1 = Frame11(Index).Top - Picture1.Height + Frame11(Index).Height + 100
End If

End Sub

Private Sub drugsyes_GotFocus(Index As Integer)
If Frame11(Index).Top + Frame11(Index).Height > Picture1.Height - Picture2.Top Then
    VScroll1 = Frame11(Index).Top - Picture1.Height + Frame11(Index).Height + 100
End If

End Sub

Private Sub drugtype_Click(Index As Integer)

If Index > 11 Or _
    (Index > 2 And Index < 5 And Val(subject(0)) > 0) Or _
    (Index > 8 And Index < 12 And Val(subject(1)) > 0) Then
            
End If
If fromfind = 1 Then
    Exit Sub
End If
If drugamt(Index).Visible = True Then
    drugamt(Index).SetFocus
End If
End Sub

Private Sub drugtype_GotFocus(Index As Integer)
Select Case Index
    Case 0 To 2
        If sdrugframe(0).Top + sdrugframe(0).Height > Picture1.Height - Picture2.Top Then
            VScroll1 = sdrugframe(0).Top - Picture1.Height + sdrugframe(0).Height + 100
        End If
    Case 3 To 5
        If sdrugframe(1).Top + sdrugframe(1).Height > Picture1.Height - Picture2.Top Then
            VScroll1 = sdrugframe(1).Top - Picture1.Height + sdrugframe(1).Height + 100
        End If
    Case 6 To 8
        If sdrugframe(2).Top + sdrugframe(2).Height > Picture1.Height - Picture2.Top Then
            VScroll1 = sdrugframe(2).Top - Picture1.Height + sdrugframe(2).Height + 100
        End If
    Case 9 To 11
        If sdrugframe(3).Top + sdrugframe(3).Height > Picture1.Height - Picture2.Top Then
            VScroll1 = sdrugframe(3).Top - Picture1.Height + sdrugframe(3).Height + 100
        End If
    Case 12 To 14
        If sdrugframe(4).Top + sdrugframe(4).Height > Picture1.Height - Picture2.Top Then
            VScroll1 = sdrugframe(4).Top - Picture1.Height + sdrugframe(4).Height + 100
        End If
    Case 15 To 17
        If sdrugframe(5).Top + sdrugframe(5).Height > Picture1.Height - Picture2.Top Then
            VScroll1 = sdrugframe(5).Top - Picture1.Height + sdrugframe(5).Height + 100
        End If
    Case 18 To 20
        If sdrugframe(6).Top + sdrugframe(6).Height > Picture1.Height - Picture2.Top Then
            VScroll1 = sdrugframe(6).Top - Picture1.Height + sdrugframe(6).Height + 100
        End If
    Case 21 To 23
        If sdrugframe(7).Top + sdrugframe(7).Height > Picture1.Height - Picture2.Top Then
            VScroll1 = sdrugframe(7).Top - Picture1.Height + sdrugframe(7).Height + 100
        End If
    Case 24 To 26
        If sdrugframe(8).Top + sdrugframe(8).Height > Picture1.Height - Picture2.Top Then
            VScroll1 = sdrugframe(8).Top - Picture1.Height + sdrugframe(8).Height + 100
        End If
    Case 27 To 29
        If sdrugframe(9).Top + sdrugframe(9).Height > Picture1.Height - Picture2.Top Then
            VScroll1 = sdrugframe(9).Top - Picture1.Height + sdrugframe(9).Height + 100
        End If
End Select
End Sub

Private Sub Form_Resize()
Picture2.Move 0, 0
With VScroll1
    .Max = Picture2.Height - Picture1.Height
End With
VScroll1.Visible = (Picture1.Height < Picture2.Height)

End Sub

Private Sub group_Click(Index As Integer)
If FROMKEY = 1 Then
    FROMKEY = 0
    Exit Sub
End If
On Error Resume Next


group(Index).Visible = False
On Error GoTo 0
End Sub


Private Sub Form_Load()
On Error Resume Next
Kill "*.dsk"
networkpath = ""
nw$ = ""
Open "nwi.ini" For Input As #1
Line Input #1, nw$
networkpath = nw$
Close #1
nw$ = ""
nwl = ""
Open "nwL.ini" For Input As #1
Line Input #1, nw$
nwl = nw$
Close #1
Dim db As Database, rs As Recordset
On Error GoTo oderror1
od1:
Set db = OpenDatabase(networkpath + "incident.mdb")
Set rs = db.OpenRecordset("Select orinumber from system")
If rs.EOF Then
    msg = MsgBox("ORI number is missing.  No exports can be generated.  Contact technical support.", 48, "Genesis Error Log")
    db.Close
    On Error Resume Next
    Exit Sub
End If
rs.MoveFirst
orinumber = rs("orinumber")
FROMKEY = 0
On Error Resume Next


Me.Height = 7600
Me.Width = 11700
Me.Top = 0
Me.Left = 0
Screen.MousePointer = 11
Call loadcodes
On Error GoTo 0
HI = incidentnumber
Call clearroutine(0)
Call loadupkey
incidentnumber = HI
With Picture2
    .AutoSize = True
    .Move 0, 0
End With
VScroll1.Max = Picture2.Height - Picture1.Height
VScroll1.LargeChange = VScroll1.Max / 10
VScroll1.SmallChange = VScroll1.Max / 100
VScroll1.Visible = (Picture1.Height < Picture2.Height)
getoutf:
Call defaultcodes
On Error GoTo getoutf
Open "NP.TAG" For Input As #1
Line Input #1, A$
incidentnumber = A$
Line Input #1, A$
PAGE = A$
Line Input #1, A$
incidentdate = A$
Close #1
Kill "NP.TAG"
On Error GoTo oderror2
od2:
Set db = OpenDatabase(networkpath + "INCIDENT.MDB")
Set rs = db.OpenRecordset("select ucr1, ucr2, ucr3,ucr4, ucr5, ucr6,ucr7, ucr8, ucr9, ucr10 from incidentSUPPORT where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
rs.MoveFirst
On Error Resume Next
For Index = 0 To 1
    vucrlist(Index).ListItems.Clear
    For vv% = 1 To 10
        If Not IsNull(rs("ucr" + Mid$(Str$(vv%), 2))) Then
            Set rs2 = db.OpenRecordset("select code from ucr where abbrev = '" + rs("ucr" + Mid$(Str$(vv%), 2)) + "'")
            rs2.MoveFirst
            Set itmx2 = vucrlist(Index).ListItems.Add(, , rs2("code"))
        End If
    Next vv%
Next Index
For Index = 0 To 5
    For t% = 1 To vucrlist(0).ListItems.Count
        pucrlist(Index).AddItem vucrlist(0).ListItems(t%)
    Next t%
Next Index
db.Close
Call incidentnumber_Click


Exit Sub
oderror1:
If Err > 3200 Then
    Resume od1
Else
    Resume Next
End If
oderror2:
If Err > 3200 Then
    Resume od2
Else
    Resume Next
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
On Error GoTo 0
End Sub

Private Sub group_KeyPress(Index As Integer, KeyAscii As Integer)
FROMKEY = 1
End Sub

Private Sub group_LostFocus(Index As Integer)
group(Index).Visible = False
End Sub

Friend Sub incidentnumber_Click()
If incidentnumber = "" Then
    Exit Sub
End If
If Len(incidentnumber) < 12 And Len(incidentnumber) > 0 Then
    incidentnumber = incidentnumber + Space$(12 - Len(incidentnumber))
End If
HI = incidentnumber
Screen.MousePointer = 11
Call clearroutine(1)
incidentnumber = HI
Call findincident
Picture2.Refresh
Screen.MousePointer = 0
VScroll1 = 0
On Error Resume Next
ORiGINAL.SetFocus
On Error GoTo 0
End Sub

Private Sub Incidentnumber_GotFocus()
If HI > "" Then
    incidentnumber = HI
    HI = ""
End If
End Sub

Private Sub Incidentnumber_LostFocus()
If Len(incidentnumber) < 12 And Len(incidentnumber) > 0 Then
    incidentnumber = incidentnumber + Space$(12 - Len(incidentnumber))
End If
End Sub
Private Sub JAIL_GotFocus(Index As Integer)
If Me.ActiveControl.Top + Me.ActiveControl.Height > Picture1.Height - Picture2.Top Then
    VScroll1 = Me.ActiveControl.Top - Picture1.Height + Me.ActiveControl.Height + 100
End If

End Sub

Private Sub numvehicle_GotFocus(Index As Integer)
If numvehicle(Index) = "" Then
    numvehicle(Index) = "#Veh"
End If
End Sub

Private Sub numvehicle_LostFocus(Index As Integer)
numvehicle(Index).Visible = False
If numvehicle(Index) = "#Veh" Then
    numvehicle(Index) = ""
End If
If BACKTAB = 0 Then
totalvalue(holdrecv + 6).SetFocus
Else
    BACKTAB = 0
End If
End Sub

Private Sub peculiarities_GotFocus(Index As Integer)
If Me.ActiveControl.Top + Me.ActiveControl.Height > Picture1.Height - Picture2.Top Then
    VScroll1 = Me.ActiveControl.Top - Picture1.Height + Me.ActiveControl.Height + 100
End If

End Sub

Private Sub pucrlist_Click(Index As Integer)


If FROMKEY = 1 Then
    FROMKEY = 0
    Exit Sub
End If
'===== Data Element 18, 19
Dim tempgroup As Integer, ITMX As String
tempgroup = Val(Mid$(group(Index).List(group(Index).ListIndex), InStr(group(Index).List(group(Index).ListIndex), "(") + 1, 2))
pucrlist(Index).Visible = False

End Sub

Private Sub pucrlist_LostFocus(Index As Integer)
pucrlist(Index).Visible = False
If BACKTAB = 0 Then
description(Index).SetFocus
Else
    BACKTAB = 0
End If
End Sub
Private Sub relationship_LostFocus(Index As Integer)
For ty% = Index To Index
    For tv% = 0 To relationship(ty%).ListCount - 1
        If relationship(ty%).Selected(tv%) = True Then
            relationship(ty%).ListIndex = tv%
            tv% = relationship(ty%).ListCount - 1
        End If
    Next tv%
Next ty%
If relationship(Index).ListIndex = -1 Then
    If Index < 10 Then
        Call Command8_Click(0)
    Else
        Call Command8_Click(1)
    End If
End If
    
End Sub
Private Sub RUNAWAY_GotFocus(Index As Integer)
If Me.ActiveControl.Top + Me.ActiveControl.Height > Picture1.Height - Picture2.Top Then
    VScroll1 = Me.ActiveControl.Top - Picture1.Height + Me.ActiveControl.Height + 100
End If

End Sub

Private Sub SECURITIESDATE_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(SECURITIESDATE) = 1 Or Len(SECURITIESDATE) = 4 Then
    Call sendslash
End If
End If

End Sub

Private Sub setrel_Click(Index As Integer)
If Val(subject(Index)) > 0 Then
    Exit Sub
End If
If fromfind = 1 Then
    Exit Sub
End If
relationshipframe(Index).Left = 7000
relationshipframe(Index).Top = setrel(Index).Top - 1000
relationshipframe(Index).Visible = True
relationship(0 + (Index * 10)).SetFocus
End Sub



Private Sub state_GotFocus(Index As Integer)
If vsname(Index) = "UNKNOWN" Then
    state(Index) = ""
End If

End Sub

Private Sub state_LostFocus(Index As Integer)
'X% = SendMessage(state.hwnd, CB_SHOWDROPDOWN, 0, 0&)
If Len(state(Index)) > 2 Then
    msg = MsgBox("Only 2 character states are allowed in this field.", 48, "Genesis Error Log")
    state(Index) = ""
    Exit Sub
End If
End Sub

Private Sub SUMMONS_GotFocus(Index As Integer)
If Me.ActiveControl.Top + Me.ActiveControl.Height > Picture1.Height - Picture2.Top Then
    VScroll1 = Me.ActiveControl.Top - Picture1.Height + Me.ActiveControl.Height + 100
End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim editerr As Integer, begindate, enddate As String
editerr = 0
Select Case Button
    
    Case "NextPage"
        If incidentnumber = "" Then
            msg = MsgBox("A valid incidentnumber must be present.", 48, "Genesis Error Log")
            Exit Sub
        End If
        schanged = 1
        If schanged = 1 Then
            tempsave = 0
            Screen.MousePointer = 11
            editerr = 0
            For ty% = 0 To 19
                relationship(ty%).ListIndex = t - 1
                For tv% = 0 To relationship(ty%).ListCount - 1
                    If relationship(ty%).Selected(tv%) = True Then
                        relationship(ty%).ListIndex = tv%
                        tv% = relationship(ty%).ListCount - 1
                    End If
                Next tv%
            Next ty%
            Call editevent(editerr, 1)
            If editerr = 0 Then
                Call editvictim(editerr, 1)
                If editerr = 0 Then
                    Call editsubject(editerr, 1)
                    If editerr = 0 Then
                        Call editproperty(editerr, 1)
                    End If
                End If
            End If
            If editerr = 0 Then
                tempsave = 0
                Call saveincident
            Else
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If
        HI = incidentnumber
        HP = Val(PAGE)
        HP = HP + 1
        Call clearroutine(1)
        incidentnumber = ""
        sinciden.incidentnumber = HI
        PAGE = Mid$(Str$(HP), 2)
        Call sinciden.incidentnumber_Click
        
    
    Case "PriorPage"
        If incidentnumber = "" Then
            msg = MsgBox("A valid incidentnumber must be present.", 48, "Genesis Error Log")
            Exit Sub
        End If
        If vsname(0) = "" And vsname(1) = "" And description(0) = "" And totalvalue(0) = "" And VIN = "" And HULL = "" And SERIAL = "" And MODEL = "" Then
        Else
            schanged = 1
            If schanged = 1 Then
                tempsave = 0
                Screen.MousePointer = 11
                editerr = 0
                For ty% = 0 To 19
                    relationship(ty%).ListIndex = t - 1
                    For tv% = 0 To relationship(ty%).ListCount - 1
                        If relationship(ty%).Selected(tv%) = True Then
                            relationship(ty%).ListIndex = tv%
                            tv% = relationship(ty%).ListCount - 1
                        End If
                    Next tv%
                Next ty%
                Call editevent(editerr, 1)
                If editerr = 0 Then
                    Call editvictim(editerr, 1)
                    If editerr = 0 Then
                        Call editsubject(editerr, 1)
                        If editerr = 0 Then
                            Call editproperty(editerr, 1)
                        End If
                    End If
                End If
                If editerr = 0 Then
                    tempsave = 0
                    Call saveincident
                Else
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
        End If
        HI = incidentnumber
        HP = Val(PAGE)
        HP = HP - 1
        If HP < 1 Then
            Unload sinciden
            Open "pp.tag" For Output As #1
            Print #1, HI
            Close #1
            incident.Show
        Else
            Call clearroutine(1)
            sinciden.incidentnumber = HI
            PAGE = Mid$(Str$(HP), 2)
            Call sinciden.incidentnumber_Click
            
        End If
        
    Case "Save"
        tempsave = 0
        Screen.MousePointer = 11
        editerr = 0
        For ty% = 0 To 19
            relationship(ty%).ListIndex = t - 1
            For tv% = 0 To relationship(ty%).ListCount - 1
                If relationship(ty%).Selected(tv%) = True Then
                    relationship(ty%).ListIndex = tv%
                    tv% = relationship(ty%).ListCount - 1
                End If
            Next tv%
        Next ty%
        Call editevent(editerr, 1)
        If editerr = 0 Then
            Call editvictim(editerr, 1)
            If editerr = 0 Then
                Call editsubject(editerr, 1)
                If editerr = 0 Then
                    Call editproperty(editerr, 1)
                End If
            End If
        End If
        If editerr = 0 Then
            tempsave = 0
            Call saveincident
        End If
        On Error Resume Next
        Screen.MousePointer = 0
    Case "Clear"
        Screen.MousePointer = 11
        Call clearroutine(0)
        Screen.MousePointer = 0
    Case "Delete"
        Screen.MousePointer = 11
        Call deleteroutine
        Screen.MousePointer = 0
    'start print button code
    Case "Print"
        'msg = MsgBox("Function not available at this time.", 48, "Not Available")
        'Exit Sub
        If frmLogin.IPRINT <> 1 And frmLogin.ISUPERVISOR <> 1 And frmLogin.SUPERVISOR <> 1 Then
            msg = MsgBox("Insufficient authority for this operation.", 48, "Genesis Error Log")
            Exit Sub
        End If
        If incidentnumber = "" Then
            msg = MsgBox("A valid incident number must be entered.", 48, "Genesis Error Log")
            Exit Sub
        End If
        schanged = 1
        If schanged = 1 Then
            Screen.MousePointer = 11
            editerr = 0
            For ty% = 0 To 19
                relationship(ty%).ListIndex = t - 1
                For tv% = 0 To relationship(ty%).ListCount - 1
                    If relationship(ty%).Selected(tv%) = True Then
                        relationship(ty%).ListIndex = tv%
                        tv% = relationship(ty%).ListCount - 1
                    End If
                Next tv%
            Next ty%
            Call editevent(editerr, 1)
            If editerr = 0 Then
                Call editvictim(editerr, 1)
                If editerr = 0 Then
                    Call editsubject(editerr, 1)
                    If editerr = 0 Then
                            Call editproperty(editerr, 1)
                    End If
                End If
            End If
            If editerr = 0 Then
                tempsave = 0
                Call saveincident
            Else
                msg = MsgBox("An incident report cannot be printed with errors.", 48, "Genesis Error Log")
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If
        On Error GoTo bbhtest
'ian's code added(page)
        report.ReportFileName = networkpath + "sincident.rpt"
        'added (page) here
        report.SelectionFormula = "{supplemental.incidentnumber} = '" + incidentnumber + "' and {supplemental.page} = " + PAGE
        report.PrintFileType = crptCrystal
        report.Destination = crptToPrinter
        report.Action = 1
        Screen.MousePointer = 0
'end ian's code
'end print button code

    Case "TempSave"
        If incidentnumber = "" Then
            msg = MsgBox("An incident number and date and name must be entered for a TEMP SAVE.", 48, "Genesis Error Log")
            Exit Sub
        End If
        tempsave = 1
        Screen.MousePointer = 11
        Call saveincident
        Call clearroutine(0)
        Screen.MousePointer = 0
    Case "TempList"
        temprevw.Show
    Case "Defaults"
        defaults.Show
    Case "Search"
        Search.Show
    Case "Exit"
        Unload sinciden
End Select
Exit Sub
bbhtest:
Resume Next
End Sub
Private Sub loadupkey()
Dim db As Database, rs As Recordset, tabname As String
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "LAWSUITE.mdb")
Set rs = db.OpenRecordset("select DPnameLF from people")
If Not rs.EOF Then
    rs.MoveFirst
Else
    On Error Resume Next
    db.Close
    Exit Sub
End If
On Error Resume Next
vsname(0).Clear
vsname(1).Clear
While Not rs.EOF
    If Not IsNull(rs("DPnameLF")) Then
        vsname(0).AddItem rs("DPnameLF")
        vsname(1).AddItem rs("DPnameLF")
    End If
    rs.MoveNext
Wend
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub

Private Sub editevent(editerr, typeedit As Integer)
Dim testgroup, totgroup As String, temperr As Integer, tempucr, tempgroup, typeselect As String, tempvalue As Single, tempdate As String
Screen.MousePointer = 11
If Len(incidentnumber) < 12 And Len(incidentnumber) > 0 Then
    incidentnumber = incidentnumber + Space$(12 - Len(incidentnumber))
End If
If age(0) > "" Then
    If Val(age(0)) = 0 And age(0) <> "00" Then
        msg = MsgBox("Age must be entered in the format of a single age: X or XX (i.e. 3 or 42) or in the format of a range of ages: XXXX (i.e. 2025, meaning 20 - 25 years old).", 48, "Genesis Error Log")
        age(0).SetFocus
        GoTo exitedite
    End If
End If
For t% = 1 To Len(age(0))
    If InStr("0123456789-", Mid$(age(0), t%, 1)) = 0 Then
        msg = MsgBox("An invalid entry in age has been found.  Valid Entry Formats:  X, XX, X-Y, XX-Y, X-YY, XX-YY, XXYY", 48, "Genesis Error Log")
        t% = Len(age(0))
        age(0).SetFocus
        GoTo exitedite
    End If
Next t%
If InStr(age(0), "-") > 0 Then
    ag1 = Val(Left$(age(0), InStr(age(0), "-") - 1))
    ag2 = Val(Mid$(age(0), InStr(age(0), "-") + 1))
    age(0) = Format$(ag1, "00") + Format$(ag2, "00")
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
    APPROVINGOFFICERDATE.SetFocus
    GoTo exitedite
End If
editeventdetail:
'===== Data Element 2
If incidentnumber = "" Then
    msg = MsgBox("A valid incidentnumber must be entered.", 48, "Genesis Error Log")
    APPROVINGOFFICERDATE.SetFocus
    GoTo exitedite
End If
If Len(incidentnumber) > 12 Then
    msg = MsgBox("The Incident Number cannot be over 12 characters long.", 48, "Genesis Error Log")
    APPROVINGOFFICERDATE.SetFocus
    GoTo exitedite
End If
For t% = 1 To Len(incidentnumber)
    If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789- ", Mid$(incidentnumber, t%, 1)) = 0 Then
        msg = MsgBox("An invalid character has been found in the Incident Number field.  Valid characters are A-Z, 0-9, and Hyphen.  Do not enter any Blanks becuase these are computer generated.", 48, "Genesis Error Log")
        t% = Len(incidentnumber)
        APPROVINGOFFICERDATE.SetFocus
        GoTo exitedite
    End If
Next t%
If Len(incidentnumber) < 12 And Len(incidentnumber) > 0 Then
    incidentnumber = incidentnumber + Space$(12 - Len(incidentnumber))
End If
APPROVINGOFFICERDATE.SetFocus
GoTo goodedite
exitedite:
editerr = 1
goodedite:

End Sub
Private Sub editvictim(editerr, typeedit As Integer)
Dim db As Database, rs, rs2, rs3 As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(networkpath + "incident.mdb")
Set rs = db.OpenRecordset("select individual,SOCIETYPUBLIC from incidentreportC where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
Set rs2 = db.OpenRecordset("select offenderdeath, noprosecution, extraditiondenied, victimdeclinescooperation, juvenilenocustody from incidentreportO where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
Set rs3 = db.OpenRecordset("select ucr1, ucr2, ucr3, ucr4, ucr5, ucr6, ucr7, ucr8, ucr9, ucr10 from incidentsupport where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
rs.MoveFirst
rs2.MoveFirst

For t% = 0 To 1
    If Val(victim(t%)) = 0 Then
        GoTo vnextT
    End If
    If t% = 0 Then
        If TV1 = "" Then
            msg = MsgBox("A Type of Victim must be entered.", 48, "Genesis Error Log")
            victim(1).SetFocus
            GoTo exiteditv
        End If
    Else
        If TV2 = "" Then
            msg = MsgBox("A Type of Victim must be entered.", 48, "Genesis Error Log")
            victim(2).SetFocus
            GoTo exiteditv
        End If
    End If
    If age(t%) > "" Then
        If Val(age(t%)) = 0 And age(t%) <> "00" Then
            msg = MsgBox("Age must be entered in the format of a single age: X or XX (i.e. 3 or 42) or in the format of a range of ages: XXXX (i.e. 2025, meaning 20 - 25 years old).", 48, "Genesis Error Log")
        age(t%).SetFocus
        GoTo exiteditv
        End If
    End If
    '===== Error 404
    If age(t%) <> "NN" And age(t%) <> "NB" And age(t%) <> "BB" Then
        For tt% = 1 To Len(age(t%))
            If InStr("0123456789-", Mid$(age(t%), tt%, 1)) = 0 Then
                msg = MsgBox("An invalid entry in age has been found.  Valid Entry Formats:  X, XX, X-Y, XX-Y, X-YY, XX-YY, XXYY", 48, "Genesis Error Log")
                tt% = Len(age(t%))
                age(t%).SetFocus
                GoTo exiteditv
            End If
        Next tt%
    End If
    If InStr(age(t%), "-") > 0 Then
        ag1 = Val(Left$(age(t%), InStr(age(t%), "-") - 1))
        ag2 = Val(Mid$(age(t%), InStr(age(t%), "-") + 1))
        age(t%) = Format$(ag1, "00") + Format$(ag2, "00")
    End If
    '===== Error 410,422
    If Len(age(t%)) = 4 Then
        If Val(Left$(age(t%), 2)) >= Val(Right$(age(t%), 2)) Then
            msg = MsgBox("For an age range, the first age must be less than the second age.", 48, "Genesis Error Log")
            age(t%).SetFocus
            GoTo exiteditv
        End If
        If Val(Left$(age(t%), 2)) = 0 Then
            msg = MsgBox("The low value in an age range cannot be 0.", 48, "Genesis Error Log")
            age(t%).SetFocus
            GoTo exiteditv
        End If
    End If

    '===== Data Element 20, 21, 22
    '===== Data Element 20, 21, 22
    Dim drugs(2)
    drugs(0) = ""
    drugs(1) = ""
    drugs(2) = ""
    dct% = t% * 6
    For ii% = (t% * 6) To (t% * 6) + 2
        If drugtype(ii%).ListIndex > -1 Then
            drugs(dct%) = drugtype(ii%).List(drugtype(ii%).ListIndex)
            dct% = dct% + 1
        End If
    Next ii%
    If dct% > t% + 6 Then
        For Z% = (t% * 6) To dct%
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
                                drugmeasurement(Z%).SetFocus
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
                                drugmeasurement(Z%).SetFocus
                                GoTo exiteditv
                        End If
                    End If
                    If InStr(drugmeasurement(zz%).List(drugmeasurement(zz%).ListIndex), "DU=") > 0 Or _
                       InStr(drugmeasurement(zz%).List(drugmeasurement(zz%).ListIndex), "NP=") > 0 Then
                        If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "DU=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "NP=") > 0 Then
                                msg = MsgBox("Two different measurements within the same weight category cannot be entered for the same type of drug.", 48, "Genesis Error Log")
                                drugmeasurement(Z%).SetFocus
                                GoTo exiteditv
                        End If
                    End If
                End If
            Next zz%
        End If
            For TTT% = 1 To Len(drugamt(Z%))
                If InStr("0123456789.", Mid$(drugamt(Z%), TTT%, 1)) = 0 Then
                    msg = MsgBox("Only 0123456789. are allowed in drug amount.  Enter all fractions as decimals (i.e. 1/2 = .5).", 48, "Genesis Error Log")
                    drugamt(Z%).SetFocus
                    GoTo exiteditv
                End If
            Next TTT%
            If drugamt(Z%) = "" Or drugmeasurement(Z%).ListIndex = -1 Then
                If drugs(Z%) > "" And Left$(drugs(Z%), 1) <> "X" And Left$(drugs(Z%), 1) <> "U" Then
                    msg = MsgBox("Drug Quantity and Measurement Type must be entered/selected.", 48, "Genesis Error Log")
                    drugamt(Z%).SetFocus
                    GoTo exiteditv
                End If
            End If
            '===== Error 366
            If drugamt(Z%) > "" Then
                If drugs(Z%) = "" Or drugmeasurement(Z%).ListIndex = -1 Then
                    msg = MsgBox("If a drug quantity is entered, then drug type and measurement type must also be entered.", 48, "Genesis Error Log")
                    drugamt(Z%).SetFocus
                    GoTo exiteditv
                End If
            End If
            '===== Error 367
            If drugmeasurement(Z%).ListIndex > -1 Then
                If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "NP=") > 0 Then
                    If Left$(drugs(Z%), 1) <> "E" And Left$(drugs(Z%), 1) <> "G" And Left$(drugs(Z%), 1) <> "K" Then
                        msg = MsgBox("Number of Plants is only allowabled with Marijuana, Opium, and Other Hallucinogens.", 48, "Genesis Error Log")
                        drugmeasurement(Z%).SetFocus
                        GoTo exiteditv
                    End If
                End If
                '===== Error 384
                If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "XX=") > 0 Then
                    If Val(drugamt(Z%)) <> 1 Then
                        msg = MsgBox("If drug measurement is NOT REPORTED, drug amount must be 1.", 48, "Genesis Error Log")
                        drugamt(Z%).SetFocus
                        GoTo exiteditv
                    End If
                End If
                '===== Error 368
                If drugs(Z%) = "" Or drugamt(Z%) = "" Then
                    msg = MsgBox("If a drug measurement is entered, then drug type and quantity must also be entered.", 48, "Genesis Error Log")
                    drugamt(Z%).SetFocus
                    GoTo exiteditv
                End If
            End If
            '===== Error 362
            If Left$(drugs(Z%), 1) = "X" Then
                If drugtype((t% * 6)).ListIndex = -1 Or drugtype((t% * 6) + 1).ListIndex = -1 Or drugtype((t% * 6) + 2).ListIndex = -1 Then
                    msg = MsgBox("If X = Over 3 Drug Types is entered, then the other 2 drug types must also be entered.", 48, "Genesis Error Log")
                    drugtype((t% * 6)).SetFocus
                    GoTo exiteditv
                End If
                '===== Error 363
                If drugamt(Z%) > "" Or drugmeasurement(Z%).ListIndex > -1 Then
                    msg = MsgBox("Drug Amount and Measurement are not valid entries for X=Over 3 Drug Types.", 48, "Genesis Error Log")
                    drugamt(Z%).SetFocus
                    GoTo exiteditv
                End If
            End If
        Next Z%
    End If
        
    '==== Mandatories E - 25 = GIVEN
    '==== Mandatories E - 26, 27, 28
    If (t% = 0 And (TV1 = "I" Or TV1 = "P")) Or (t% = 1 And (TV2 = "I" Or TV2 = "P")) Then
        If Val(age(t%)) = 0 And age(t%) <> "00" Then
            msg = MsgBox("Invalid age entered.", 48, "Genesis Error Log")
            age(t%).SetFocus
            GoTo exiteditv
        End If
        If sex(t%).ListIndex = -1 Then
            msg = MsgBox("Invalid sex entered.", 48, "Genesis Error Log")
            sex(t%).SetFocus
            GoTo exiteditv
        End If
        If race(t%).ListIndex = -1 Then
            msg = MsgBox("Invalid race entered.", 48, "Genesis Error Log")
            race(t%).SetFocus
            GoTo exiteditv
        End If
        If ethnicity(t%).ListIndex = -1 Then
            msg = MsgBox("Ethnicity is a required entry.", 48, "Genesis Error Log")
            ethnicity(t%).SetFocus
            GoTo exiteditv
        End If
        If resident(t%).ListIndex = -1 Then
            msg = MsgBox("Resident Status is a required entry.", 48, "Genesis Error Log")
            resident(t%).SetFocus
            GoTo exiteditv
        End If
    Else
        '===== Error 458
        If age(t%) > "" Then
            msg = MsgBox("Age is not a valid entry for Victim if Type of Victim is not Individual.", 48, "Genesis Error Log")
            age(t%).SetFocus
            GoTo exiteditv
        End If
        If sex(t%).ListIndex > -1 Then
            msg = MsgBox("Sex is not a valid entry for Victim if Type of Victim is not Individual.", 48, "Genesis Error Log")
            sex(t%).SetFocus
            GoTo exiteditv
        End If
        If race(t%).ListIndex > -1 Then
            msg = MsgBox("Race is not a valid entry for Victim if Type of Victim is not Individual.", 48, "Genesis Error Log")
            race(t%).SetFocus
            GoTo exiteditv
        End If
        If ethnicity(t%).ListIndex > -1 Then
            msg = MsgBox("Ethnicity is not a valid entry for Victim if Type of Victim is not Individual.", 48, "Genesis Error Log")
            ethnicity(t%).SetFocus
            GoTo exiteditv
        End If
        If resident(t%).ListIndex > -1 Then
            msg = MsgBox("Resident Status is not a valid entry for Victim if Type of Victim is not Individual.", 48, "Genesis Error Log")
            resident(t%).SetFocus
            GoTo exiteditv
        End If
    End If

    '===== Data Element 24
    FOUNDVUCR = 0
    For tt% = 1 To vucrlist(t%).ListItems.Count
        If vucrlist(t%).ListItems(tt%).Selected Then
            FOUNDVUCR = 1
            tt% = vucrlist(t%).ListItems.Count
        End If
    Next tt%
    If FOUNDVUCR = 0 Then
        msg = MsgBox("At least one UCR code must be connected to the victim.", 48, "Genesis Error Log")
        vucrlist(t%).SetFocus
        GoTo exiteditv
    End If
    '----edit by Ed Sloan 11/30/99 from SCLED 8/9/95 update----------------
    For tt% = 1 To vucrlist(t%).ListItems.Count
        If vucrlist(t%).ListItems(tt%).Selected Then
            tempvucr = Mid$(vucrlist(t%).ListItems(tt%), InStr(vucrlist(t%).ListItems(tt%), "(") + 1, 3)
            '===== Error 464,465,467
            Select Case tempvucr
                Case "13A", "13B", "13C", "09A", "09B", "09C", "100", "11A", "11B", "11C", "11D", "36A", "36B", "36C", "90A", "90J"
                    If Not ((t% = 0 And (TV1 = "I" Or TV1 = "P")) Or (t% = 1 And (TV2 = "P" Or TV2 = "I"))) Then
                        msg = MsgBox("Individual or Police Officer must be selected for Crimes Against Person.", 48, "Genesis Error Log")
                        vucrlist(t%).SetFocus
                        GoTo exiteditv
                    End If
                Case "90Z"
                Case "90J", "36C", "980", "978", "753", "756"
                Case "35A", "35B", "39A", "39B", "39C", "39D", "370", "40A", "40B", "520", "90B", "90C", "90D", "90G", "90H", "90I", "90E", "90F"
                    If Not rs("SOCIETYPUBLIC") Then
                        msg = MsgBox("Society must be selected for Crimes Against Society.", 48, "Genesis Error Log")
                        vucrlist(t%).SetFocus
                        GoTo exiteditv
                    End If
                Case Else
                    If rs("societypublic") Then
                        msg = MsgBox("Society cannot be selected for Crimes Against Property.", 48, "Genesis Error Log")
                        vucrlist(t%).SetFocus
                        GoTo exiteditv
                    End If
            End Select
            Select Case tempvucr
                    Case "13A", "13B", "13C", "09A", "09B", "09C", "100", "11A", "11B", "11C", "11D", "36A", "36B"
                        If Not ((t% = 0 And (TV1 = "I" Or TV1 = "P")) Or (t% = 1 And (TV2 = "P" Or TV2 = "I"))) Then
                            msg = MsgBox("Individual must be selected for Crimes Against Person.", 48, "Genesis Error Log")
                            vucrlist(t%).SetFocus
                            GoTo exiteditv
                        End If
                    Case "90F"
                        If Not ((t% = 0 And (TV1 = "I" Or TV1 = "P" Or TV1 = "S")) Or (t% = 1 And (TV2 = "P" Or TV2 = "I" Or TV2 = "S"))) Then
                            msg = MsgBox("Individual or Society must be selected for Family Offenses/Nonviolent.", 48, "Genesis Error Log")
                            vucrlist(t%).SetFocus
                            GoTo exiteditv
                        End If
                    Case "90J", "36C", "980", "978", "753", "756"
                    Case "35A", "35B", "39A", "39B", "39C", "39D", "370", "40A", "40B", "520", "90B", "90C", "90D", "90G", "90H", "90I"
                        If Not ((t% = 0 And (TV1 = "S")) Or (t% = 1 And (TV2 = "S"))) Then
                            msg = MsgBox("Society must be selected for Crimes Against Society.", 48, "Genesis Error Log")
                            vucrlist(t%).SetFocus
                            GoTo exiteditv
                        End If
                    Case Else
                        If ((t% = 0 And (TV1 = "S")) Or (t% = 1 And (TV2 = "S"))) Then
                            msg = MsgBox("Society cannot be selected for Crimes Against Property.", 48, "Genesis Error Log")
                            vucrlist(t%).SetFocus
                            GoTo exiteditv
                        End If
            End Select
            '===== Error 481
            If tempvucr = "36B" And Val(age(t%)) > 15 And vucrlist(t%).ListItems(tt%).Selected = True Then
                msg = MsgBox("For statutory rape, the victim must be less than or equal to 15 years of age.", 48, "Genesis Error Log")
                age(t%).SetFocus
                GoTo exiteditv
            End If
            '===== SCEdit 8/9/95
            If tempvucr = "23C" Then
                If Len(age(t%)) = 4 Then
                    If Val(Right$(age(t%), 2)) > 15 Then
                        msg = MsgBox("For Offense 23C, the victim age must be 15 years old or less.", 48, "Genesis Error Log")
                        age(t%).SetFocus
                        GoTo exiteditv
                    End If
                Else
                If Len(age(t%)) = 2 Then
                    If Val(age(t%)) > 15 Then
                        msg = MsgBox("For Offense 23C, the victim age must be 15 years old or less.", 48, "Genesis Error Log")
                        age(t%).SetFocus
                        GoTo exiteditv
                    End If
                Else
                    msg = MsgBox("For Offense 23C, the victim age must be 15 years old or less.", 48, "Genesis Error Log")
                    age(t%).SetFocus
                    GoTo exiteditv
                End If
                End If
            End If
        End If
    Next tt%
    '----------------------------------------------------------------------
    '----edit by Ed Sloan 12/03/99 from SCLED 3/5/97 update----------------
    Dim ucrexists, valinj As Boolean
    ucrexists = False
    valinj = False
    FOUNDINJTYPE% = 0
    For r% = 1 To vucrlist(t%).ListItems.Count
        If vucrlist(t%).ListItems(r%).Selected Then
            tempvucr = Mid$(vucrlist(t%).ListItems(r%), InStr(vucrlist(t%).ListItems(r%), "(") + 1, 3)
            Select Case tempvucr
                Case "100", "11A", "11B", "11C", "11D", "120", "13A", "13B", "210"
                    FOUNDINJTYPE% = 1
            End Select
        End If
    Next r%
    For r% = 1 To vucrlist(t%).ListItems.Count
        If vucrlist(t%).ListItems(r%).Selected Then
            tempvucr = Mid$(vucrlist(t%).ListItems(r%), InStr(vucrlist(t%).ListItems(r%), "(") + 1, 3)
            ICT% = 0
            For rr% = 1 To injury(t%).ListItems.Count
                If injury(t%).ListItems(rr%).Selected Then
                    ICT% = ICT% + 1
                End If
            Next rr%
            Select Case tempvucr
                '===== Error 479
                Case "13B"
                    ucrexists = True
                    For Q% = 1 To injury(t%).ListItems.Count
                        If injury(t%).ListItems(Q%).Selected Then
                            tempinj = Mid$(injury(t%).ListItems(Q%), InStr(injury(t%).ListItems(Q%), "(") + 1, 1)
                            If (tempinj = "M" Or tempinj = "N") Then
                                validinj = True
                            End If
                        End If
                    Next Q%
                    If ucrexists And Not validinj Then
                        msg = MsgBox("For simple assault, the only injury types can be minor or none.", 48, "Genesis Error Log")
                        vucrlist(t%).SetFocus
                        GoTo exiteditv
                    End If
                '===== Error 401
                Case "100", "11A", "11B", "11C", "11D", "120", "13A", "13B", "210"
                    If ICT% = 0 Then
                        msg = MsgBox("Type of injury must be selected for UCR " + tempvcur + ".", 48, "Genesis Error Log")
                        vucrlist(t%).SetFocus
                        GoTo exiteditv
                    End If
               '===== Error 419
                Case Else
                    If ICT% > 0 And FOUNDINJTYPE% = 0 Then
                        msg = MsgBox("Type of injury is not applicable for UCR " + tempvcur + ".", 48, "Genesis Error Log")
                        vucrlist(t%).SetFocus
                        GoTo exiteditv
                    End If
            End Select
        End If
    Next r%
    '===== Error 407
    For r% = 1 To injury(t%).ListItems.Count
        If injury(t%).ListItems(r%).Selected Then
            If Mid$(injury(t%).ListItems(r%), InStr(injury(t%).ListItems(r%), "(") + 1, 1) = "N" Then
                For rr% = 1 To r% - 1
                    If injury(t%).ListItems(rr%).Selected Then
                        msg = MsgBox("When Ijury Type N=None is selected, no other values may be selected.", 48, "Genesis Error Log")
                        injury(t%).SetFocus
                        GoTo exiteditv
                    End If
                Next rr%
                For rr% = r% + 1 To injury(t%).ListItems.Count
                    If injury(t%).ListItems(rr%).Selected Then
                        msg = MsgBox("When Ijury Type N=None is selected, no other values may be selected.", 48, "Genesis Error Log")
                        injury(t%).SetFocus
                        GoTo exiteditv
                    End If
                Next rr%
            End If
        End If
    Next r%
    '===== Error 450
    For r% = t% * 10 To (t% * 10) + 9
        If relationship(r%).ListIndex > -1 Then
            temprel = Mid$(relationship(r%).List(relationship(r%).ListIndex), InStr(relationship(r%).List(relationship(r%).ListIndex), "(") + 1, 2)
            If temprel = "SE" Then
                If Val(age(t%)) < 10 Then
                    msg = MsgBox("The relationship of victim to subject cannot be 'SE' when victim's age is less than 10.", 48, "Genesis Error Log")
                    age(t%).SetFocus
                    GoTo exiteditv
                End If
            End If
        End If
    Next r%

    For tt% = 1 To 10
    
        If Not IsNull(rs3("ucr" + Mid$(Str$(tt%), 2))) Then
            tempucr = rs3("ucr" + Mid$(Str$(tt%), 2))
            If tempucr = "13A" Or tempucr = "13B" Or tempucr = "13C" Then
                If UCase(vsname(t%)) <> "UNKNOWN" Then
                    If relationship(t% * 10).ListIndex = -1 And relationship((t% * 10) + 1).ListIndex = -1 And relationship((t% * 10) + 2).ListIndex = -1 And relationship((t% * 10) + 3).ListIndex = -1 And relationship((t% * 10) + 4).ListIndex = -1 And relationship((t% * 10) + 5).ListIndex = -1 And relationship((t% * 10) + 6).ListIndex = -1 And relationship((t% * 10) + 7).ListIndex = -1 And relationship((t% * 10) + 8).ListIndex = -1 And relationship((t% * 10) + 9).ListIndex = -1 Then
                        msg = MsgBox("A relationship to subject must be selected.", 48, "Genesis Error Log")
                        relationship(t% * 10).SetFocus
                        GoTo exiteditv
                    End If
                End If
            End If
            '===== Additional F 12
            '===== Data Element 7, 13
            If tempucr = "09A" Or tempucr = "09B" Or tempucr = "09C" Then
                If UCase(vsname(t%)) <> "UNKNOWN" Then
                    If relationship(t% * 10).ListIndex = -1 And relationship((t% * 10) + 1).ListIndex = -1 And relationship((t% * 10) + 2).ListIndex = -1 And relationship((t% * 10) + 3).ListIndex = -1 And relationship((t% * 10) + 4).ListIndex = -1 And relationship((t% * 10) + 5).ListIndex = -1 And relationship((t% * 10) + 6).ListIndex = -1 And relationship((t% * 10) + 7).ListIndex = -1 And relationship((t% * 10) + 8).ListIndex = -1 And relationship((t% * 10) + 9).ListIndex = -1 Then
                        msg = MsgBox("A relationship to subject must be selected.", 48, "Genesis Error Log")
                        relationship(t% * 10).SetFocus
                        GoTo exiteditv
                    End If
                End If
            End If
            '===== Additional F 13
            '===== Data Element 7
            If tempucr = "100" Then
                If UCase(vsname(t%)) <> "UNKNOWN" Then
                    If relationship(t% * 10).ListIndex = -1 And relationship((t% * 10) + 1).ListIndex = -1 And relationship((t% * 10) + 2).ListIndex = -1 And relationship((t% * 10) + 3).ListIndex = -1 And relationship((t% * 10) + 4).ListIndex = -1 And relationship((t% * 10) + 5).ListIndex = -1 And relationship((t% * 10) + 6).ListIndex = -1 And relationship((t% * 10) + 7).ListIndex = -1 And relationship((t% * 10) + 8).ListIndex = -1 And relationship((t% * 10) + 9).ListIndex = -1 Then
                        msg = MsgBox("A relationship to subject must be selected.", 48, "Genesis Error Log")
                        relationship(t% * 10).SetFocus
                        GoTo exiteditv
                    End If
                End If
            End If
            '===== Additional F 18
            If tempucr = "120" Then
                If (t% = 0 And (TV1 = "I" Or TV1 = "P")) Or (t% = 1 And (TV2 = "I" Or TV2 = "P")) Then
                    If UCase(vsname(t%)) <> "UNKNOWN" Then
                        If relationship(t% * 10).ListIndex = -1 And relationship((t% * 10) + 1).ListIndex = -1 And relationship((t% * 10) + 2).ListIndex = -1 And relationship((t% * 10) + 3).ListIndex = -1 And relationship((t% * 10) + 4).ListIndex = -1 And relationship((t% * 10) + 5).ListIndex = -1 And relationship((t% * 10) + 6).ListIndex = -1 And relationship((t% * 10) + 7).ListIndex = -1 And relationship((t% * 10) + 8).ListIndex = -1 And relationship((t% * 10) + 9).ListIndex = -1 Then
                            msg = MsgBox("A relationship to subject must be selected.", 48, "Genesis Error Log")
                            relationship(t% * 10).SetFocus
                            GoTo exiteditv
                        End If
                    End If
                End If
            End If
            '===== Additional F 19
            If tempucr = "11A" Or tempucr = "11B" Or tempucr = "11C" Or tempucr = "11D" Then
                If UCase(vsname(t%)) <> "UNKNOWN" Then
                    If relationship(t% * 10).ListIndex = -1 And relationship((t% * 10) + 1).ListIndex = -1 And relationship((t% * 10) + 2).ListIndex = -1 And relationship((t% * 10) + 3).ListIndex = -1 And relationship((t% * 10) + 4).ListIndex = -1 And relationship((t% * 10) + 5).ListIndex = -1 And relationship((t% * 10) + 6).ListIndex = -1 And relationship((t% * 10) + 7).ListIndex = -1 And relationship((t% * 10) + 8).ListIndex = -1 And relationship((t% * 10) + 9).ListIndex = -1 Then
                        msg = MsgBox("A relationship to subject must be selected.", 48, "Genesis Error Log")
                        relationship(t% * 10).SetFocus
                        GoTo exiteditv
                    End If
                End If
            End If
            '===== Additional F 20
            If tempucr = "36A" Or tempucr = "36B" Then
                If UCase(vsname(t%)) <> "UNKNOWN" Then
                    If relationship(t% * 10).ListIndex = -1 And relationship((t% * 10) + 1).ListIndex = -1 And relationship((t% * 10) + 2).ListIndex = -1 And relationship((t% * 10) + 3).ListIndex = -1 And relationship((t% * 10) + 4).ListIndex = -1 And relationship((t% * 10) + 5).ListIndex = -1 And relationship((t% * 10) + 6).ListIndex = -1 And relationship((t% * 10) + 7).ListIndex = -1 And relationship((t% * 10) + 8).ListIndex = -1 And relationship((t% * 10) + 9).ListIndex = -1 Then
                        msg = MsgBox("A relationship to subject must be selected.", 48, "Genesis Error Log")
                        relationship(t% * 10).SetFocus
                        GoTo exiteditv
                    End If
                End If
            End If
        End If
    Next tt%

vnextT:
Next t%
GoTo goodeditv
exiteditv:
editerr = 1
goodeditv:
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
Private Sub editsubject(editerr, typeedit As Integer)

Dim db As Database, rs, rs2 As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(networkpath + "incident.mdb")
Set rs = db.OpenRecordset("select policeofficer,individual from incidentreportC where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
Set rs2 = db.OpenRecordset("select offenderdeath, noprosecution, extraditiondenied, victimdeclinescooperation, juvenilenocustody from incidentreportO where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
rs.MoveFirst

For t% = 0 To 1
    If Val(subject(t%)) = 0 Then
        subject(t%).SetFocus
        GoTo snextt
    End If
    '===== SC LEOKA
    If rs("policeofficer") Then
        If TWOMANVEHICLE(t%) = 0 And ONEMANVEHICLE(t%) = 0 And DETECTIVE(t%) = 0 And TODOTHER(t%) = 0 Then
            msg = MsgBox("If Type Victim is Police Officer, a selection must be made for Two Man Vehicle, One Man Vehicle, Detective/Special Assignment, or Other.", 48, "Genesis Error Log")
            TWOMANVEHICLE(t%).SetFocus
            GoTo exitedits
        End If
        If Not TWOMANVEHICLE(t%) Then
            If ALONE(t%) = 0 And ASSISTED(t%) = 0 Then
                msg = MsgBox("If Type Victim is Police Officer and not Two Man Vehicle, a selection must be made for Alone or Assisted.", 48, "Genesis Error Log")
                TWOMANVEHICLE(t%).SetFocus
                GoTo exitedits
            End If
        End If
    End If
    tage = age(t%)
    '===== Error 761
    If RUNAWAY(t%) = 1 Then
        If Len(tage) = 2 Then
            If Val(tage) > 17 Then
                msg = MsgBox("A runaway must be under the age of 18.", 48, "Genesis Error Log")
                RUNAWAY(t%).SetFocus
                GoTo exitedits
            End If
        Else
        If Len(tage) = 4 Then
            If Val(Right$(tage, 2)) > 17 Then
                msg = MsgBox("A runaway must be under the age of 18.", 48, "Genesis Error Log")
                RUNAWAY(t%).SetFocus
                GoTo exitedits
            End If
        Else
            msg = MsgBox("A runaway must be under the age of 18.", 48, "Genesis Error Log")
            RUNAWAY(t%).SetFocus
            GoTo exitedits
        End If
        End If
    End If
    If age(t%) > "" Then
        If Val(age(t%)) = 0 And age(t%) <> "00" Then
            msg = MsgBox("Age must be entered in the format of a single age: X or XX (i.e. 3 or 42) or in the format of a range of ages: XXXX (i.e. 2025, meaning 20 - 25 years old).", 48, "Genesis Error Log")
            age(t%).SetFocus
            GoTo exitedits
        End If
    Else
        '===== error 504
        msg = MsgBox("Subject age must be entered. (00 = unknown)", 48, "Genesis Error Log")
        age(t%).SetFocus
        GoTo exitedits
    End If
    '===== Error 504
    If sex(t%).ListIndex = -1 Then
        msg = MsgBox("A value for Sex in subject data must be entered.", 48, "Genesis Error Log")
        sex(t%).SetFocus
        GoTo exitedits
    End If
    If race(t%).ListIndex = -1 Then
        msg = MsgBox("A value for race in subject data must be entered.", 48, "Genesis Error Log")
        race(t%).SetFocus
        GoTo exitedits
    End If
    If ethnicity(t%).ListIndex = -1 Then
        msg = MsgBox("A value for ethnicity in subject data must be entered.", 48, "Genesis Error Log")
        ethnicity(t%).SetFocus
        GoTo exitedits
    End If
    '===== Error 404,556
    If age(t%) <> "NN" And age(t%) <> "NB" And age(t%) <> "BB" Then
        For tt% = 1 To Len(age(t%))
            If InStr("0123456789-", Mid$(age(t%), tt%, 1)) = 0 Then
                msg = MsgBox("An invalid entry in age has been found.  Valid Entry Formats:  X, XX, X-Y, XX-Y, X-YY, XX-YY, XXYY", 48, "Genesis Error Log")
                tt% = Len(age(t%))
                age(t%).SetFocus
                GoTo exitedits
            End If
        Next tt%
    End If
    If InStr(age(t%), "-") > 0 Then
        ag1 = Val(Left$(age(t%), InStr(age(t%), "-") - 1))
        ag2 = Val(Mid$(age(t%), InStr(age(t%), "-") + 1))
        age(t%) = Format$(ag1, "00") + Format$(ag2, "00")
    End If
    '===== Error 410,422,509,510,522
    If Len(age(t%)) = 4 Then
        If Val(Left$(age(t%), 2)) >= Val(Right$(age(t%), 2)) Then
            msg = MsgBox("For an age range, the first age must be less than the second age.", 48, "Genesis Error Log")
            age(t%).SetFocus
            GoTo exitedits
        End If
        If Val(Left$(age(t%), 2)) = 0 Then
            msg = MsgBox("The low value in an age range cannot be 0.", 48, "Genesis Error Log")
            age(t%).SetFocus
            GoTo exitedits
        End If
    End If
    '===== Data Element 20, 21, 22
    Dim drugs(2)
    drugs(0) = ""
    drugs(1) = ""
    drugs(2) = ""
    dct% = t% * 3
    For ii% = (t% * 3) To (t% * 3) + 2
        If drugtype(ii%).ListIndex > -1 Then
            drugs(dct%) = drugtype(ii%).List(drugtype(ii%).ListIndex)
            dct% = dct% + 1
        End If
    Next ii%
    If dct% > 0 And dct% > t% * 3 Then
        For Z% = (t% * 3) To dct%
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
                                drugmeasurement(Z%).SetFocus
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
                                drugmeasurement(Z%).SetFocus
                                GoTo exitedits
                        End If
                    End If
                    If InStr(drugmeasurement(zz%).List(drugmeasurement(zz%).ListIndex), "DU=") > 0 Or _
                       InStr(drugmeasurement(zz%).List(drugmeasurement(zz%).ListIndex), "NP=") > 0 Then
                        If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "DU=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "NP=") > 0 Then
                                msg = MsgBox("Two different measurements within the same weight category cannot be entered for the same type of drug.", 48, "Genesis Error Log")
                                drugmeasurement(Z%).SetFocus
                                GoTo exitedits
                        End If
                    End If
                End If
            Next zz%
        End If
            For TTT% = 1 To Len(drugamt(Z%))
                If InStr("0123456789.", Mid$(drugamt(Z%), TTT%, 1)) = 0 Then
                    msg = MsgBox("Only 0123456789. are allowed in drug amount.  Enter all fractions as decimals (i.e. 1/2 = .5).", 48, "Genesis Error Log")
                    drugamt(Z%).SetFocus
                    GoTo exitedits
                End If
            Next TTT%
            If drugamt(Z%) = "" Or drugmeasurement(Z%).ListIndex = -1 Then
                If drugs(Z%) > "" And Left$(drugs(Z%), 1) <> "X" And Left$(drugs(Z%), 1) <> "U" Then
                    msg = MsgBox("Drug Quantity and Measurement Type must be entered/selected.", 48, "Genesis Error Log")
                    drugamt(Z%).SetFocus
                    GoTo exitedits
                End If
            End If
            '===== Error 366
            If drugamt(Z%) > "" Then
                If drugs(Z%) = "" Or drugmeasurement(Z%).ListIndex = -1 Then
                    msg = MsgBox("If a drug quantity is entered, then drug type and measurement type must also be entered.", 48, "Genesis Error Log")
                    drugamt(Z%).SetFocus
                    GoTo exitedits
                End If
            End If
            '===== Error 367
            If drugmeasurement(Z%).ListIndex > -1 Then
                If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "NP=") > 0 Then
                    If Left$(drugs(Z%), 1) <> "E" And Left$(drugs(Z%), 1) <> "G" And Left$(drugs(Z%), 1) <> "K" Then
                        msg = MsgBox("Number of Plants is only allowabled with Marijuana, Opium, and Other Hallucinogens.", 48, "Genesis Error Log")
                        drugmeasurement(Z%).SetFocus
                        GoTo exitedits
                    End If
                End If
                '===== Error 384
                If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "XX=") > 0 Then
                    If Val(drugamt(Z%)) <> 1 Then
                        msg = MsgBox("If drug measurement is NOT REPORTED, drug amount must be 1.", 48, "Genesis Error Log")
                        drugmeasurement(Z%).SetFocus
                        GoTo exitedits
                    End If
                End If
                '===== Error 368
                If drugs(Z%) = "" Or drugamt(Z%) = "" Then
                    msg = MsgBox("If a drug measurement is entered, then drug type and quantity must also be entered.", 48, "Genesis Error Log")
                    drugamt(Z%).SetFocus
                    GoTo exitedits
                End If
            End If
            '===== Error 362
            If Left$(drugs(Z%), 1) = "X" Then
                If drugtype(t% * 3).ListIndex = -1 Or drugtype((t% * 3) + 1).ListIndex = -1 Or drugtype((t% * 3) + 2).ListIndex = -1 Then
                    msg = MsgBox("If X = Over 3 Drug Types is entered, then the other 2 drug types must also be entered.", 48, "Genesis Error Log")
                    drugtype(t% * 3).SetFocus
                    GoTo exitedits
                End If
                '===== Error 363
                If drugamt(Z%) > "" Or drugmeasurement(Z%).ListIndex > -1 Then
                    msg = MsgBox("Drug Amount and Measurement are not valid entries for X=Over 3 Drug Types.", 48, "Genesis Error Log")
                    drugamt(Z%).SetFocus
                    GoTo exitedits
                End If
            End If
        Next Z%
    End If
    
    '==== Mandatories E - 37, 38, 39
    '===== Error 501
'    If (t% = 0 And (TV1 = "I" Or TV1 = "P")) Or (t% = 1 And (TV2 = "I" Or TV2 = "P")) Or UCase(vsname(t%)) <> "UNKNOWN" Then
    If UCase(vsname(t%)) <> "UNKNOWN" Then
        If Val(tage) = 0 And tage <> "00" Then
            msg = MsgBox("Invalid age entered.", 48, "Genesis Error Log")
            vsname(t%).SetFocus
            GoTo exitedits
        End If
        If sex(t%).ListIndex = -1 Then
            msg = MsgBox("Invalid sex entered.", 48, "Genesis Error Log")
            sex(t%).SetFocus
            GoTo exitedits
        End If
        If race(t%).ListIndex = -1 Then
            msg = MsgBox("Invalid race entered.", 48, "Genesis Error Log")
            race(t%).SetFocus
            GoTo exitedits
        End If
        If ethnicity(t%).ListIndex = -1 Then
            msg = MsgBox("Invalid ETHNICITY entered.", 48, "Genesis Error Log")
            ethnicity(t%).SetFocus
            GoTo exitedits
        End If
    End If
    
    If (rs2("offenderdeath") Or rs2("noprosecution") Or rs2("extraditiondenied") Or rs2("victimdeclinescooperation") Or rs2("juvenilenocustody")) Then
        If age(t%) = "00" Then
            msg = MsgBox("For an exceptional clearance, the subjects age (other than 00) must be selected.", 48, "Genesis Error Log")
            age(t%).SetFocus
            GoTo exitedits
        End If
        If race(t%).ListIndex = -1 Or race(t%).List(race(t%).ListIndex) = "Unknown" Then
            msg = MsgBox("For an exceptional clearance, the subject's race (other than unknown) must be selected.", 48, "Genesis Error Log")
            race(t%).SetFocus
            GoTo exitedits
        End If
        If sex(t%).ListIndex = -1 Or sex(t%).List(sex(t%).ListIndex) = "Unknown" Then
            msg = MsgBox("For an exceptional clearance, the subject's sex (other than unknown) must be selected.", 48, "Genesis Error Log")
            sex(t%).SetFocus
            GoTo exitedits
        End If
        If ethnicity(t%).ListIndex = -1 Or ethnicity(t%).List(ethnicity(t%).ListIndex) = "Unknown" Then
            msg = MsgBox("For an exceptional clearance, the subject's ethnicity (other than unknown) must be selected.", 48, "Genesis Error Log")
            ethnicity(t%).SetFocus
            GoTo exitedits
        End If
    End If

snextt:
Next t%
GoTo goodedits
exitedits:
editerr = 1
goodedits:
db.Close
On Error Resume Next
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub

Private Sub totalvalue_GotFocus(Index As Integer)
If Me.ActiveControl.Top + Me.ActiveControl.Height > Picture1.Height - Picture2.Top Then
    VScroll1 = Me.ActiveControl.Top - Picture1.Height + Me.ActiveControl.Height + 100
End If

End Sub

Private Sub totalvalue_LostFocus(Index As Integer)
If BACKTAB = 0 Then
If Val(totalvalue(Index)) > 0 And Index >= 18 And Index <= 23 Then
    holdrecv = Index
    If fromfind = 0 Then
        DATERECOVERED(Index Mod 6).Visible = True
        DATERECOVERED(Index Mod 6).SetFocus
    End If
Else
'=====Data Item 18 and 19
If Val(totalvalue(Index)) > 0 And (Index >= 0 And Index <= 5) Then
    If Mid$(pucrlist(Index).List(pucrlist(Index).ListIndex), InStr(pucrlist(Index).List(pucrlist(Index).ListIndex), "(") + 1, 3) = "240" Then
        tempgroup = Val(Mid$(group(Index).List(group(Index).ListIndex), InStr(group(Index).List(group(Index).ListIndex), "(") + 1, 2))
        If (tempgroup = 3 Or tempgroup = 5 Or tempgroup = 24 Or tempgroup = 28 Or tempgroup = 37) And fromfind = 0 Then
            numvehicle(Index Mod 6).Visible = True
            numvehicle(Index Mod 6).SetFocus
        End If
    End If
End If
End If
If Index > 23 And Index < 30 Then
    If Mid$(group(Index Mod 6).List(group(Index Mod 6).ListIndex), InStr(group(Index Mod 6).List(group(Index Mod 6).ListIndex), "(") + 1, 2) = "10" Then
        If pucrlist(Index).ListIndex > -1 And InStr(pucrlist(Index).List(pucrlist(Index).ListIndex), "(35A)") > 0 And fromfind = 0 Then
            sdrugframe(Index + 4).Left = 2000
            sdrugframe(Index + 4).Top = description(Index).Top - sdrugframe(Index + 4).Height - 100
            sdrugframe(Index + 4).Visible = True
            drugtype(Index + 4).SetFocus
        End If
    End If
End If
Else
    BACKTAB = 0
End If
Call figure
End Sub


Private Sub TWOMANVEHICLE_GotFocus(Index As Integer)
If Me.ActiveControl.Top + Me.ActiveControl.Height > Picture1.Height - Picture2.Top Then
    VScroll1 = Me.ActiveControl.Top - Picture1.Height + Me.ActiveControl.Height + 100
End If
End Sub

Private Sub victim_LostFocus(Index As Integer)
If Val(victim(Index)) = 1 Then
    msg = MsgBox("Victim Number must be greater than 1.", 48, "Genesis Error Log")
End If
If Val(victim(Index)) > 1 Then
    If Index = 0 Then
        inp = InputBox("Enter Type of Victim (I=Individual/B=Business/F=Financial Inst./G=Government/R=Relig. Org./S=Society/O=Other/U=Unknown/P=Police Officer", "", TV1)
    Else
        inp = InputBox("Enter Type of Victim (I=Individual/B=Business/F=Financial Inst./G=Government/R=Relig. Org./S=Society/O=Other/U=Unknown/P=Police Officer", "", TV2)
    End If
    inp = UCase(inp)
    If InStr("IBFGRSOUP", inp) = 0 Then
        msg = MsgBox("Invalid Type of Victim entered.", 48, "Genesis Error Log")
    End If
    If Index = 0 Then
        TV1 = inp
    Else
        TV2 = inp
    End If
End If
End Sub

Private Sub VISIBLEINJURYNO_GotFocus(Index As Integer)
If Frame1(Index).Top + Frame1(Index).Height > Picture1.Height - Picture2.Top Then
    VScroll1 = Frame1(Index).Top - Picture1.Height + Frame1(Index).Height + 100
End If

End Sub

Private Sub VISIBLEINJURYYES_GotFocus(Index As Integer)
If Frame1(Index).Top + Frame1(Index).Height > Picture1.Height - Picture2.Top Then
    VScroll1 = Frame1(Index).Top - Picture1.Height + Frame1(Index).Height + 100
End If

End Sub

Private Sub VScroll1_Change()
Picture2.Top = -VScroll1.Value
End Sub

Private Sub vsname_Click(Index As Integer)
If vsname(Index) > "" Then
    Call FILLDATA(Index)
End If
End Sub

Private Sub vsname_GotFocus(Index As Integer)
If Me.ActiveControl.Top + Me.ActiveControl.Height > Picture1.Height - Picture2.Top Then
    VScroll1 = Me.ActiveControl.Top - Picture1.Height + Me.ActiveControl.Height + 100
End If

End Sub

Private Sub vsname_LostFocus(Index As Integer)
If Len(vsname(Index)) > 60 Then
    msg = MsgBox("A maximum of 60 characters is allowed for name.  Your entry is being truncated to 60 characters.", 48, "Genesis INformation Log")
    vsname(Index) = Left$(vsname(Index), 60)
End If
If UCase(vsname(Index)) <> "SAME AS VICTIM" And vsname(Index) > "" And InStr(vsname(Index), ",") = 0 Then
    If (Index = 0 And TV1 <> "I") Or (Index = 1 And TV2 <> "I") Then
    Else
        msg = MsgBox("All names in the Incident Report system should be entered in the format last name + comma + firstname.", 48, "Invalid Data Format")
        vsname(Index).SetFocus
    End If
End If
If vsname(Index) > "" Then
    Call FILLDATA(Index)
End If

End Sub


Private Sub WANTED_GotFocus(Index As Integer)
If Me.ActiveControl.Top + Me.ActiveControl.Height > Picture1.Height - Picture2.Top Then
    VScroll1 = Me.ActiveControl.Top - Picture1.Height + Me.ActiveControl.Height + 100
End If

End Sub

Private Sub WARRANT_GotFocus(Index As Integer)
If Me.ActiveControl.Top + Me.ActiveControl.Height > Picture1.Height - Picture2.Top Then
    VScroll1 = Me.ActiveControl.Top - Picture1.Height + Me.ActiveControl.Height + 100
End If

End Sub


Private Sub clearroutine(TP As Integer)
Dim ITMX As ListItem
TV1 = ""
TV2 = ""
For t% = 0 To 1
    alcoholyes(t%) = 0
    alcoholno(t%) = 1
    alcoholunknown(t%) = 0
    drugsyes(t%) = 0
    drugsno(t%) = 1
    drugsunknown(t%) = 0
    'vucrlist(t%).ListItems.Clear
Next t%
For t% = 0 To 5
    DATERECOVERED(t%).Visible = False
    group(t%).Visible = False
    group(t%).ListIndex = -1
    numvehicle(t%).Visible = False
    numvehicle(t% + 6).Visible = False
    pucrlist(t%).Visible = False
    pucrlist(t%).ListIndex = -1
Next t%
'ian's code
NARRATIVE.Text = ""
'end ian's code
reportingofficer(0) = ""
REPORTINGOFFICERDATE(0) = ""
reportingofficeRunit(0) = ""
reportingofficer(1) = ""
REPORTINGOFFICERDATE(1) = ""
reportingofficeRunit(1) = ""
approvingofficer = ""
APPROVINGOFFICERDATE = ""
approvingofficeRunit = ""
For t% = 0 To 41
    totalvalue(t%) = ""
Next t%
For t% = 0 To 19
    relationship(t%).ListIndex = -1
    For tt% = 0 To relationship(t%).ListCount - 1
        relationship(t%).Selected(tt%) = False
    Next tt%
Next t%
TVSTOLEN = ""
TVDAMAGED = ""
TVBURNED = ""
TVRECOVERED = ""
TVSEIZED = ""
TVCOUNTERFEIT = ""
TVUNKNOWN = ""
relationshipframe(0).Visible = False
relationshipframe(1).Visible = False
For t% = 0 To 9
    sdrugframe(t%).Visible = False
Next t%
vucrf(0).Visible = False
vucrf(1).Visible = False
LOCATIONNUMBER(0) = ""
LOCATIONNUMBER(1) = ""
For t% = 0 To 5
    group(t%).ListIndex = -1
    totalvalue(t%) = ""
    description(t%) = ""
    numvehicle(t%) = ""
    DATERECOVERED(t%) = ""
Next t%
stolen = 0
damaged = 0
burned = 0
recovered = 0
seized = 0
counterfeited = 0
typeunknown = 0
vsname(0) = ""
vsname(1) = ""
For t% = 0 To 1
    address(t%) = ""
    city(t%) = ""
    state(t%) = ""
    zipcode(t%) = ""
    race(t%).ListIndex = -1
    sex(t%).ListIndex = -1
    age(t%) = ""
    ethnicity(t%).ListIndex = -1
    BIRTHDATE(Index) = ""
    ht(t%) = ""
    weight(t%) = ""
    resident(t%).ListIndex = -1
    hair(t%) = ""
    eyes(t%) = ""
    peculiarities(t%) = ""
    VISIBLEINJURYYES(t%) = False
    VISIBLEINJURYNO(t%) = True
    NONVISIBLEINJURYYES(t%) = False
    NONVISIBLEINJURYNO(t%) = True
    TWOMANVEHICLE(t%) = 0
    ONEMANVEHICLE(t%) = 0
    DETECTIVE(t%) = 0
    TODOTHER(t%) = 0
    ALONE(t%) = 0
    ASSISTED(t%) = 0
    For tt% = 1 To injury(t%).ListItems.Count
        injury(t%).ListItems(tt%).Selected = False
    Next tt%
Next t%
WARRANT(0) = 0
WANTED(0) = 0
RUNAWAY(0) = 0
ARREST(0) = 0
JAIL(0) = 0
SUMMONS(0) = 0
WARRANT(1) = 0
WANTED(1) = 0
RUNAWAY(1) = 0
ARREST(1) = 0
JAIL(1) = 0
SUMMONS(1) = 0
For t% = 0 To 29
    drugtype(t%).ListIndex = -1
    drugmeasurement(t%).ListIndex = -1
    drugamt(t%) = ""
Next t%
SSTOLEN = 0
SRECOVERED = 0
SFOUND = 0
STOWED = 0
SSUSPECT = 0
SVICTIM = 0
SVEHICLE = 0
SGUN = 0
SBOAT = 0
sLICENSEPLATE = 0
SSECURITIES = 0
SARTICLE = 0
VIN = ""
HULL = ""
SERIAL = ""
SERIALSTATE = ""
YEARREG = ""
YEAREXP = ""
YEARN = ""
MAKE = ""
stype = ""
MODEL = ""
STYLE = ""
scolor = ""
BRANDNAME = ""
CALIBER = ""
NIC = ""
DENOMINATION = ""
ISSUER = ""
SECURITIESDATE = ""
MISCELLANEOUS = ""
For t% = 0 To 1
    HOMEDAYPHONE(t%) = ""
    HOMENIGHTPHONE(t%) = ""
    WORKDAYPHONE(t%) = ""
    WORKNIGHTPHONE(t%) = ""
    LOCATIONNUMBER(t%) = ""
Next t%
ORiGINAL = 0
MODIFIES = 0
SUPPLEMENTAL = 0
CASEst = 0
ADDITIONALV = 0
additionalo = 0
ADDITIONALS = 0
ADDITIONALR = 0
For t% = 0 To 1
    complainant(t%) = 0
    victim(t%) = ""
    subject(t%) = ""
Next t%
For t% = 0 To 3
    alcoholyes(t%) = False
    alcoholno(t%) = True
    alcoholunknown(t%) = False
    drugsyes(t%) = False
    drugsno(t%) = True
    drugsunknown(t%) = False
Next t%


VScroll1 = 0
If TP = 0 Then
    Call defaultcodes
End If
End Sub
Private Sub defaultcodes()
Dim db As Database, rs As Recordset, ITMX As ListItem
On Error GoTo oderror
od:
Set db = OpenDatabase(networkpath + "incident.mdb")
Set rs = db.OpenRecordset("select * from codes")
If rs.EOF Then
    On Error Resume Next
    db.Close
    Exit Sub
End If
rs.MoveFirst
On Error Resume Next
For t% = 0 To 1
    state(t%).Clear
    city(t%).Clear
    sex(t%).Clear
    race(t%).Clear
    ethnicity(t%).Clear
    resident(t%).Clear
    injury(t%).ListItems.Clear
Next t%
While Not rs.EOF
    Select Case rs("type")
        Case "state"
            For t% = 0 To 1
                state(t%).AddItem rs("code")
               'If UCase(rs("default")) = "Y" Then
               '     state(t%).ListIndex = state(t%).ListCount - 1
               ' End If
            Next t%
        Case "city"
            For t% = 0 To 1
                city(t%).AddItem rs("code")
                'If UCase(rs("default")) = "Y" Then
                '    city(t%).ListIndex = city(t%).ListCount - 1
                'End If
            Next t%
        Case "injury"
            For t% = 0 To 1
                Set ITMX = injury(t%).ListItems.Add(, , rs("code"))
                If UCase(rs("default")) = "Y" Then
                    injury(t%).ListItems(injury(t%).ListItems.Count).Selected = True
                    injury(t%).ListItems(injury(t%).ListItems.Count).EnsureVisible
                Else
                    injury(t%).ListItems(injury(t%).ListItems.Count).Selected = False
                End If
            Next t%
        Case "sex"
            For t% = 0 To 1
                sex(t%).AddItem rs("code")
                If UCase(rs("default")) = "Y" Then
                    sex(t%).ListIndex = sex(t%).ListCount - 1
                End If
            Next t%
        Case "race"
            For t% = 0 To 1
                race(t%).AddItem rs("code")
                If UCase(rs("default")) = "Y" Then
                    race(t%).ListIndex = race(t%).ListCount - 1
                End If
            Next t%
        Case "ethnicity"
            For t% = 0 To 1
                ethnicity(t%).AddItem rs("code")
                If UCase(rs("default")) = "Y" Then
                    ethnicity(t%).ListIndex = ethnicity(t%).ListCount - 1
                End If
            Next t%
        Case "resident"
            For t% = 0 To 1
                resident(t%).AddItem rs("code")
                If UCase(rs("default")) = "Y" Then
                    resident(t%).ListIndex = resident(t%).ListCount - 1
                End If
            Next t%
    End Select
    rs.MoveNext
Wend
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub
Private Sub editproperty(editerr, typeedit As Integer)
Dim ITMX As ListItem, db As Database, rs, rs2 As Recordset

'===== Data Element 20, 21, 22
Dim drugs(23)
drugs(6) = ""
drugs(7) = ""
drugs(8) = ""
drugs(9) = ""
drugs(10) = ""
drugs(11) = ""
drugs(12) = ""
drugs(13) = ""
drugs(14) = ""
drugs(15) = ""
drugs(16) = ""
drugs(17) = ""
drugs(18) = ""
drugs(19) = ""
drugs(20) = ""
drugs(21) = ""
drugs(22) = ""
drugs(23) = ""
For oo% = 6 To 23 Step 3
    dct% = oo%
    For ii% = oo% To oo% + 2
        If drugtype(ii%).ListIndex > -1 Then
            drugs(dct%) = drugtype(ii%).List(drugtype(ii%).ListIndex)
            dct% = dct% + 1
        End If
    Next ii%
    If dct% > 0 Then
        For Z% = oo% To dct%
            '===== Error 306
            '===== SCEdit 4/21/92 P29
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
                                        drugmeasurement(Z%).SetFocus
                                        GoTo exiteditp
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
                                        drugmeasurement(Z%).SetFocus
                                        GoTo exiteditp
                                End If
                        End If
                        If InStr(drugmeasurement(zz%).List(drugmeasurement(zz%).ListIndex), "DU=") > 0 Or _
                            InStr(drugmeasurement(zz%).List(drugmeasurement(zz%).ListIndex), "NP=") > 0 Then
                                If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "DU=") > 0 Or _
                                    InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "NP=") > 0 Then
                                        msg = MsgBox("Two different measurements within the same weight category cannot be entered for the same type of drug.", 48, "Genesis Error Log")
                                        drugmeasurement(Z%).SetFocus
                                        GoTo exiteditp
                                End If
                        End If
                    End If
                Next zz%
            End If
            For t% = 1 To Len(drugamt(Z%))
                If InStr("0123456789.", Mid$(drugamt(Z%), t%, 1)) = 0 Then
                    msg = MsgBox("Only 0123456789. are allowed in drug amount.  Enter all fractions as decimals (i.e. 1/2 = .5).", 48, "Genesis Error Log")
                    drugamt(Z%).SetFocus
                    GoTo exiteditp
                End If
            Next t%
            '===== Error 364
            If drugamt(Z%) = "" Or drugmeasurement(Z%).ListIndex = -1 Then
                If drugs(Z%) > "" And Left$(drugs(Z%), 1) <> "X" And Left$(drugs(Z%), 1) <> "U" Then
                    msg = MsgBox("Drug Quantity and Measurement Type must be entered/selected.", 48, "Genesis Error Log")
                    drugamt(Z%).SetFocus
                    GoTo exiteditp
                End If
            End If
            '===== Error 366
            If drugamt(Z%) > "" Then
                If drugs(Z%) = "" Or drugmeasurement(Z%).ListIndex = -1 Then
                    msg = MsgBox("If a drug quantity is entered, then drug type and measurement type must also be entered.", 48, "Genesis Error Log")
                    drugamt(Z%).SetFocus
                    GoTo exiteditp
                End If
            End If
            '===== Error 367
            If drugmeasurement(Z%).ListIndex > -1 Then
                If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "NP=") > 0 Then
                    If Left$(drugs(Z%), 1) <> "E" And Left$(drugs(Z%), 1) <> "G" And Left$(drugs(Z%), 1) <> "K" Then
                        msg = MsgBox("Number of Plants is only allowabled with Marijuana, Opium, and Other Hallucinogens.", 48, "Genesis Error Log")
                        drugmeasurement(Z%).SetFocus
                        GoTo exiteditp
                    End If
                End If
                '===== Error 384
                If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "XX=") > 0 Then
                    If Val(drugamt(Z%)) <> 1 Then
                        msg = MsgBox("If drug measurement is NOT REPORTED, drug amount must be 1.", 48, "Genesis Error Log")
                        drugmeasurement(Z%).SetFocus
                        GoTo exiteditp
                    End If
                End If
                '===== Error 368
                If drugs(Z%) = "" Or drugamt(Z%) = "" Then
                    msg = MsgBox("If a drug measurement is entered, then drug type and quantity must also be entered.", 48, "Genesis Error Log")
                    drugamt(Z%).SetFocus
                    GoTo exiteditp
                End If
            End If
            '===== Error 362
            If Left$(drugs(Z%), 1) = "X" Then
                If drugtype(oo%).ListIndex = -1 Or drugtype(oo% + 1).ListIndex = -1 Or drugtype(oo% + 2).ListIndex = -1 Then
                    msg = MsgBox("If X = Over 3 Drug Types is entered, then the other 2 drug types must also be entered.", 48, "Genesis Error Log")
                    drugtype(oo%).SetFocus
                    GoTo exiteditp
                End If
                '===== Error 363
                If drugamt(Z%) > "" Or drugmeasurement(Z%).ListIndex > -1 Then
                    msg = MsgBox("Drug Amount and Measurement are not valid entries for X=Over 3 Drug Types.", 48, "Genesis Error Log")
                    drugamt(Z%).SetFocus
                    GoTo exiteditp
                End If
            End If
        Next Z%
    End If
Next oo%
'GLENN
FOUNDPROP = False
For d% = 0 To 5
    If Not description(d%).Text = "" Then
        FOUNDPROP = True
        If pucrlist(d%).ListIndex = -1 Then
            msg = MsgBox("A UCR must be associated with the property described.", 48, "Genesis Error Log")
            pucrlist(d%).SetFocus
            GoTo exiteditp
        Else
            If group(d%).ListIndex = -1 Then
                msg = MsgBox("A Group must be associated with the property described.", 48, "Genesis Error Log")
                group(d%).SetFocus
                GoTo exiteditp
            End If
        End If
    End If
Next d%
'GLENN

For t% = 0 To 29
    If Val(totalvalue(t%)) > 0 Then
        '===== Error 342
        If Val(totalvalue(t%)) >= 1000000 Then
            msg = MsgBox("WARNING:  A value of $1,000,000 or greater has been entered in the property value section.  Is this correct?", 4, "Genesis Information Log")
            If msg = 7 Then
                totalvalue(t%).SetFocus
                GoTo exiteditp
            End If
        End If
        Select Case t%
            Case 0, 6, 12, 18, 24
                If pucrlist(0).ListIndex = -1 Then
                    msg = MsgBox("A UCR must be associated with the property described.", 48, "Genesis Error Log")
                    pucrlist(0).SetFocus
                    GoTo exiteditp
                End If
            Case 1, 7, 13, 19, 25
                If pucrlist(1).ListIndex = -1 Then
                    msg = MsgBox("A UCR must be associated with the property described.", 48, "Genesis Error Log")
                    pucrlist(1).SetFocus
                    GoTo exiteditp
                End If
            Case 2, 8, 14, 20, 26
                If pucrlist(2).ListIndex = -1 Then
                    msg = MsgBox("A UCR must be associated with the property described.", 48, "Genesis Error Log")
                    pucrlist(2).SetFocus
                    GoTo exiteditp
                End If
            Case 3, 9, 15, 21, 27
                If pucrlist(3).ListIndex = -1 Then
                    msg = MsgBox("A UCR must be associated with the property described.", 48, "Genesis Error Log")
                    pucrlist(3).SetFocus
                    GoTo exiteditp
                End If
            Case 4, 10, 16, 22, 28
                If pucrlist(4).ListIndex = -1 Then
                    msg = MsgBox("A UCR must be associated with the property described.", 48, "Genesis Error Log")
                    pucrlist(4).SetFocus
                    GoTo exiteditp
                End If
            Case 5, 11, 17, 23, 29
                If pucrlist(5).ListIndex = -1 Then
                    msg = MsgBox("A UCR must be associated with the property described.", 48, "Genesis Error Log")
                    pucrlist(5).SetFocus
                    GoTo exiteditp
                End If
        End Select
        
    tt% = t% Mod 6
            
    '===== Error 352
    If Val(totalvalue(tt%)) = 0 And Val(totalvalue(tt% + 6)) = 0 And Val(totalvalue(tt% + 12)) = 0 And Val(totalvalue(tt% + 18)) = 0 And Val(totalvalue(tt% + 24)) = 0 And Val(totalvalue(tt% + 30)) = 0 Then
        If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) <> "35A" Then
            If group(tt%).ListIndex > -1 Or DATERECOVERED(tt%) > "" Or numvehicle(tt%) > "" Or numvehicle(tt% + 6) > "" Or drugtype(tt% + 6).ListIndex > -1 Or drugamt(tt% + 6) > "" Or drugmeasurement(tt% + 6).ListIndex > -1 Then
                msg = MsgBox("When Type Property Loss = None, no other applicable values are allowed.", 48, "Genesis Error Log")
                totalvalue(tt%).SetFocus
                GoTo exiteditp
            End If
        Else
            If group(tt%).ListIndex > -1 Or DATERECOVERED(tt%) > "" Or numvehicle(tt%) > "" Or numvehicle(tt% + 6) > "" Then
                msg = MsgBox("When Type Property Loss = None for UCR 35A, no other applicable values are allowed, except drug-related values.", 48, "Genesis Error Log")
                group(tt%).SetFocus
                GoTo exiteditp
            End If
        End If
    End If
    If Val(totalvalue(tt% + 30)) > 0 Then
        If group(tt%).ListIndex > -1 Or DATERECOVERED(tt%) > "" Or numvehicle(tt%) > "" Or numvehicle(tt% + 6) > "" Or drugtype(tt% + 6).ListIndex > -1 Or drugamt(tt% + 6) > "" Or drugmeasurement(tt% + 6).ListIndex > -1 Then
            msg = MsgBox("When Type Property Loss = Unknown, no other applicable values are allowed.", 48, "Genesis Error Log")
            group(tt%).SetFocus
            GoTo exiteditp
        End If
    End If
        
    '===== Data Element 14, 15
    '===== Error 372,375
    If Val(totalvalue(tt%)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
        If description(tt%) = "" Or pucrlist(tt%).ListIndex = -1 Or group(tt%).ListIndex = -1 Then
            msg = MsgBox("If Burned, Counterfeited, Damaged, Recovered, Seized, or Stolen are selected, then all other PROPERTY values must be entered.", 48, "Genesis Error Log")
            totalvalue(tt%).SetFocus
            GoTo exiteditp
        End If
        If Val(totalvalue(tt% + 18)) > 0 Then
            If Not IsDate(DATERECOVERED(tt%)) Then
                msg = MsgBox("If Burned, Counterfeited, Damaged, Recovered, Seized, or Stolen are selected, the all other PROPERTY values must be entered.", 48, "Genesis Error Log")
                DATERECOVERED(tt%).SetFocus
                GoTo exiteditp
            End If
        End If
    End If
        
    For uu% = 0 To 5
        '===== Error 268
        If group(uu%).ListIndex > -1 And nolarceny = 0 Then
            Select Case Val(Mid$(group(uu%).List(group(uu%).ListIndex), InStr(group(uu%).List(group(uu%).ListIndex), "(") + 1, 2))
                Case 1, 5, 24, 28, 37
                    msg = MsgBox("Illogical property group/ucr combination.", 48, "Genesis Error Log")
                    group(uu%).SetFocus
                    GoTo exiteditp
            End Select
        End If
        If pucrlist(uu%).ListIndex > -1 Then
            '===== Error 390
            '===== SCEdit 4/21/92 P30
            Select Case Mid$(pucrlist(uu%).List(pucrlist(uu%).ListIndex), InStr(pucrlist(uu%).List(pucrlist(uu%).ListIndex), "(") + 1, 3)
                Case "240"
                    Select Case Val(Mid$(group(uu%).List(group(uu%).ListIndex), InStr(group(uu%).List(group(uu%).ListIndex), "(") + 1, 2))
                        Case 29 To 35
                            msg = MsgBox("Illogical property group/ucr combination.", 48, "Genesis Error Log")
                            pucrlist(uu%).SetFocus
                            GoTo exiteditp
                    End Select
                Case "23B"
                    Select Case Val(Mid$(group(uu%).List(group(uu%).ListIndex), InStr(group(uu%).List(group(uu%).ListIndex), "(") + 1, 2))
                        Case 1, 3, 4, 5, 12, 15, 18, 24, 28, 29, 30, 31, 32, 33, 34, 35, 37, 39
                            msg = MsgBox("Illogical property group/ucr combination.", 48, "Genesis Error Log")
                            group(uu%).SetFocus
                            GoTo exiteditp
                    End Select
                Case "23C"
                    Select Case Val(Mid$(group(uu%).List(group(uu%).ListIndex), InStr(group(uu%).List(group(uu%).ListIndex), "(") + 1, 2))
                        Case 1, 3, 5, 12, 15, 18, 24, 28, 29, 30, 31, 32, 33, 34, 35, 37, 39
                            msg = MsgBox("Illogical property group/ucr combination.", 48, "Genesis Error Log")
                            group(uu%).SetFocus
                            GoTo exiteditp
                    End Select
                Case "23F"
                    Select Case Val(Mid$(group(uu%).List(group(uu%).ListIndex), InStr(group(uu%).List(group(uu%).ListIndex), "(") + 1, 2))
                        Case 3, 5, 24, 28, 29, 30, 31, 32, 33, 34, 35, 37
                            msg = MsgBox("Illogical property group/ucr combination.", 48, "Genesis Error Log")
                            group(uu%).SetFocus
                            GoTo exiteditp
                    End Select
                Case "23G"
                    Select Case Val(Mid$(group(uu%).List(group(uu%).ListIndex), InStr(group(uu%).List(group(uu%).ListIndex), "(") + 1, 2))
                        Case 38, 88
                        Case Else
                            msg = MsgBox("Illogical property group/ucr combination.", 48, "Genesis Error Log")
                            group(uu%).SetFocus
                            GoTo exiteditp
                    End Select
                Case "23H"
                    Select Case Val(Mid$(group(uu%).List(group(uu%).ListIndex), InStr(group(uu%).List(group(uu%).ListIndex), "(") + 1, 2))
                        Case 3, 5, 24, 28, 37
                            msg = MsgBox("Illogical property group/ucr combination.", 48, "Genesis Error Log")
                            group(uu%).SetFocus
                            GoTo exiteditp
                    End Select
            End Select
            If Mid$(pucrlist(uu%).List(pucrlist(uu%).ListIndex), InStr(pucrlist(uu%).List(pucrlist(uu%).ListIndex), "(") + 1, 3) = "35A" Then
                If (Val(totalvalue(uu% + 0)) = 0 And Val(totalvalue(uu% + 6)) = 0 And Val(totalvalue(uu% + 12)) = 0 And Val(totalvalue(uu% + 18)) = 0 And Val(totalvalue(uu% + 24)) = 0 And Val(totalvalue(uu% + 30)) = 0) Then
                    If drugtype((uu% * 3) + 6).ListIndex = -1 Then
                        msg = MsgBox("A suspected drug type must be selected for this property.", 48, "Genesis Error Log")
                        pucrlist(uu%).SetFocus
                        GoTo exiteditp
                    End If
                End If
                If Val(totalvalue(uu% + 24)) > 0 Then
                    If group(uu%).ListIndex > -1 Then
                        If Mid$(group(uu%).List(group(uu%).ListIndex), InStr(group(uu%).List(group(uu%).ListIndex), "(") + 1, 2) = "10" Then
                            If drugtype((uu% * 3) + 6).ListIndex = -1 Then
                                msg = MsgBox("A suspected drug type must be selected for this property.", 48, "Genesis Error Log")
                                group(uu%).SetFocus
                                GoTo exiteditp
                            End If
                            If Val(drugamt((uu% * 3) + 6)) = 0 Or drugmeasurement((uu% * 3) + 6).ListIndex = -1 Then
                                msg = MsgBox("An amount of suspected drugs must be entered for this property.", 48, "Genesis Error Log")
                                drugamt((uu% * 3) + 6).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next uu%
        
    '===== Data Element 15
    '===== Error 353
    If InStr(group(tt%).List(group(tt%).ListIndex), "(88)") > 0 Then
        If Val(totalvalue(tt%)) <> 1 Then
            msg = MsgBox("If Type of Property = Pending Inventory(88), a value of 1 must be entered.", 48, "Genesis Error Log")
            totalvalue(tt%).SetFocus
            GoTo exiteditp
        End If
    End If
        
    '===== Data Element 16
    '===== Error 351
    If group(tt%).ListIndex > -1 Then
        tempgrp = Mid$(group(tt%).List(group(tt%).ListIndex), InStr(group(tt%).List(group(tt%).ListIndex), "(") + 1, 2)
        If (Val(totalvalue(tt%)) = 0 And Val(totalvalue(tt% + 6)) = 0 And Val(totalvalue(tt% + 12)) = 0 And Val(totalvalue(tt% + 18)) = 0 And Val(totalvalue(tt% + 24)) = 0 And Val(totalvalue(tt% + 30)) = 0) And _
            tempgrp <> "09" And tempgrp <> "22" And tempgrp <> "77" And tempgrp <> "99" Then
            msg = MsgBox("A property value of 0 is only allowed for Credit/Debit Cards, Nonnegotiable Instruments, Other, and Special Category.", 48, "Genesis Error Log")
            totalvalue(tt%).SetFocus
            GoTo exiteditp
        End If
        '---Ed Sloan added "Or Val(itmx.SubItems(2)) <> 1"-----------
        If Not (Val(totalvalue(tt%)) = 0 And Val(totalvalue(tt% + 6)) = 0 And Val(totalvalue(tt% + 12)) = 0 And Val(totalvalue(tt% + 18)) = 0 And Val(totalvalue(tt% + 24)) = 0 And Val(totalvalue(tt% + 30)) = 0) And _
            (tempgrp = "09" Or tempgrp = "22") Then
            msg = MsgBox("A property value of 0 is required for Credit/Debit Cards and Nonnegotiable Instruments.", 48, "Genesis Error Log")
            totalvalue(tt%).SetFocus
            GoTo exiteditp
        End If
        '===== Error 383
        If tempgrp = "10" And InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(35A)") > 0 And _
            (Val(totalvalue(tt%)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0) Then
            msg = MsgBox("A value is not valid for the Drugs/Narcotics and Drug/Narcotic Violations combination.", 48, "Genesis Error Log")
            totalvalue(tt%).SetFocus
            GoTo exiteditp
        End If
    Else
    '===== Error 354
        If (Val(totalvalue(tt%)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0) Then
            msg = MsgBox("If a value greater than 0 is entered, an associated property type must be selected.", 48, "Genesis Error Log")
            totalvalue(tt%).SetFocus
            GoTo exiteditp
        End If
    End If
    
    '===== Data Element 17
    '===== Error 305
    If DATERECOVERED(tt%) > "" And Not IsDate(DATERECOVERED(tt%)) Then
        msg = MsgBox("Date Recovered is not a valid date.", 48, "Genesis Error Log")
        DATERECOVERED(tt%).SetFocus
        GoTo exiteditp
    End If
    If IsDate(DATERECOVERED(tt%)) Then
        If CVDate(DATERECOVERED(tt%)) < CVDate(incidentdate) Then
            msg = MsgBox("Date Recovered cannot be earlier that Date of Offense.", 48, "Genesis Error Log")
            DATERECOVERED(tt%).SetFocus
            GoTo exiteditp
        End If
        '===== Error 356
        If Val(totalvalue(t%)) = 0 Or group(tt%).ListIndex = -1 Then
            msg = MsgBox("If Date Recovered is entered, both type and value of property must be entered.", 48, "Genesis Error Log")
            totalvalue(tt%).SetFocus
            GoTo exiteditp
        End If
        If Val(totalvalue(tt% + 18)) = 0 Then
            msg = MsgBox("If Date Recovered is entered, then Recovered must be selected.", 48, "Genesis Error Log")
            totalvalue(tt%).SetFocus
            GoTo exiteditp
        End If
    End If
    
    '===== Data Element 18
    '===== Error 357,358,359
    If tempucr = "240" Then
        If Val(totalvalue(tt%)) > 0 And Val(numvehicle(tt%)) = 0 Then
            If group(tt%).ListIndex > -1 Then
                tempgrp = Mid$(group(tt%).List(group(tt%).ListIndex), InStr(group(tt%).List(group(tt%).ListIndex), "(") + 1, 2)
                If tempgrp = "03" Or tempgrp = "05" Or tempgrp = "24" Or tempgrp = "28" Or tempgrp = "37" Then
                    msg = MsgBox("A number of stolen vehicles must be entered.", 48, "Genesis Error Log")
                    totalvalue(tt%).SetFocus
                    GoTo exiteditp
                End If
            End If
        End If
    End If
    
    '===== Data Element 19
    '===== Error 360,361,359
    '===== SCEdit 12/19/92 P35-C  allow tempgrp = 38
    If tempucr = "240" Then
        If Val(totalvalue(tt% + 18)) > 0 And Val(numvehicle(tt%)) = 0 Then
            If group(tt%).ListIndex > -1 Then
                tempgrp = Mid$(group(tt%).List(group(tt%).ListIndex), InStr(group(tt%).List(group(tt%).ListIndex), "(") + 1, 2)
                If tempgrp = "03" Or tempgrp = "05" Or tempgrp = "24" Or tempgrp = "28" Or tempgrp = "37" Or tempgrp = "38" Then
                    msg = MsgBox("A number of recovered vehicles must be entered.", 48, "Genesis Error Log")
                    totalvalue(tt%).SetFocus
                    GoTo exiteditp
                End If
            End If
        End If
    End If
    
    '===== Data Element 20
    '===== Error 362
    If Left$(drugtype(((tt% + 2) * 3)).List(drugtype(((tt% + 2) * 3)).ListIndex), 1) = "X" Or Left$(drugtype(((tt% + 2) * 3) + 1).List(drugtype(((tt% + 2) * 3) + 1).ListIndex), 1) = "X" Or Left$(drugtype(((tt% + 2) * 3) + 2).List(drugtype(((tt% + 2) * 3) + 2).ListIndex), 1) = "X" Then
        If drugtype((tt% + 2) * 3).ListIndex = -1 Or drugtype(((tt% + 2) * 3) + 1).ListIndex = -1 Or drugtype(((tt% + 2) * 3) + 2).ListIndex = -1 Then
            msg = MsgBox("If X = Over 3 Drug Types is entered, then the other 2 drug types must also be entered.", 48, "Genesis Error Log")
            drugtype((tt% + 2) * 3).SetFocus
            GoTo exiteditp
        End If
    End If
        
    End If
    
    
Next t%
On Error GoTo oderror
od:
Set db = OpenDatabase(networkpath + "INCIDENT.MDB")
Set rs = db.OpenRecordset("SELECT UCR1, UCR2, UCR3, UCR4, UCR5, UCR6, UCR7, UCR8, UCR9, UCR10 FROM INCIDENTSUPPORT WHERE INCIDENTNUMBER = '" + incidentnumber + "'")
Set rs2 = db.OpenRecordset("SELECT * FROM INCIDENTREPORTC WHERE INCIDENTNUMBER = '" + incidentnumber + "'")
If Not rs.EOF Then
    rs.MoveFirst
Else
    msg = MsgBox("Invalid incident report data.", 48, "Genesis Error Log")
    incident.SetFocus
    GoTo exiteditp
End If
If Not rs2.EOF Then
    rs2.MoveFirst
Else
    msg = MsgBox("Invalid incident report data.", 48, "Genesis Error Log")
    incident.SetFocus
    GoTo exiteditp
End If

For t% = 0 To 9
    
    If Not IsNull(rs("ucr" + Mid$(Str$(t% + 1), 2))) Then
        tempucr = rs("ucr" + Mid$(Str$(t% + 1), 2))
    
        '===Additional F 7
        '===== Data Element 7
        '===== Data Element 12
        If tempucr = "35A" Or tempucr = "35B" Then
            TP$ = "Drug Offenses"
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Not rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
                            If Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                msg = MsgBox(TP$ + " not completed must have associated value of None or Unknown on property tab.", 48, "Genesis Error Log")
                                pucrlist(tt%).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                        If rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
                            'GLENN
                            '===== Error 301
                            If (Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) = 0 Or Val(totalvalue(tt% + 30)) > 0) Then
                                If Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                    msg = MsgBox(TP$ + " completed must have associated information of Seized property entered.", 48, "Genesis Error Log")
                                    totalvalue(tt% + 0).SetFocus
                                    GoTo exiteditp
                                End If
                            End If
                        End If
                    End If
                End If
            Next tt%
        End If
        If tempucr = "35A" Then
            TP$ = "Drug Offenses"
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Val(totalvalue(tt% + 24)) > 0 Then
                            If group(tt%).ListIndex > -1 Then
                                '===== SCEdit 4/21/92 P29
                                If Mid$(group(tt%).List(group(tt%).ListIndex), InStr(group(tt%).List(group(tt%).ListIndex), "(") + 1, 2) <> "10" Then
                                    msg = MsgBox("Property Description must have a value of 10 for Seized on Drug offense.", 48, "Genesis Error Log")
                                    group(tt%).SetFocus
                                    GoTo exiteditp
                                End If
                            End If
                        End If
                    End If
                End If
            Next tt%
        End If
        If tempucr = "35B" Then
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Val(totalvalue(tt% + 24)) > 0 Then
                            If group(tt%).ListIndex > -1 Then
                                '===== SCEdit 4/21/92 P29
                                If Mid$(group(tt%).List(group(tt%).ListIndex), InStr(group(tt%).List(group(tt%).ListIndex), "(") + 1, 2) <> "11" Then
                                    msg = MsgBox("Property Description must have a value of 11 for Seized on Drug Equipment offense.", 48, "Genesis Error Log")
                                    group(tt%).SetFocus
                                    GoTo exiteditp
                                End If
                            End If
                        End If
                    End If
                End If
            Next tt%
        End If
        ' === Additional F 1
        If tempucr = "200" Then
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                        If Val(totalvalue(tt%)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                            If Not rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
                                If Val(totalvalue(tt%)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                    msg = MsgBox("Arson not completed must have associated value of None or Unknown on property tab.", 48, "Genesis Error Log")
                                    totalvalue(tt%).SetFocus
                                    GoTo exiteditp
                                End If
                            End If
                            If rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
                                If Val(totalvalue(tt% + 12)) = 0 Then
                                    msg = MsgBox("Arson completed must have associated value of Burned on property tab.", 48, "Genesis Error Log")
                                    totalvalue(tt% + 12).SetFocus
                                    GoTo exiteditp
                                End If
                                If group(tt%).ListIndex = -1 Then
                                    msg = MsgBox("Type(Group) and Total Value must be entered.", 48, "Genesis Error Log")
                                    group(tt%).SetFocus
                                    GoTo exiteditp
                                End If
                            End If
                        End If
                End If
            Next tt%
        End If
        
        ' === Additional F 3
        '===== Additional F 4
        '===== Data Element 10
        '===== Data Element 11
        If tempucr = "510" Or tempucr = "220" Then
            Select Case tempucr
                Case "510"
                    TP$ = "Bribery"
                Case "220"
                    TP$ = "Burglary"
            End Select
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Not rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
                            If Val(totalvalue(tt%)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                msg = MsgBox(TP$ + " not completed must have associated value of None or Unknown on property tab.", 48, "Genesis Error Log")
                                totalvalue(tt%).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                        If rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
                            If (Val(totalvalue(tt%)) = 0 And Val(totalvalue(tt% + 6)) = 0 And Val(totalvalue(tt% + 12)) = 0 And Val(totalvalue(tt% + 18)) = 0 And Val(totalvalue(tt% + 24)) = 0 And Val(totalvalue(tt% + 30)) = 0) Or _
                                ((Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt%)) > 0) And Val(totalvalue(tt% + 6)) = 0 And Val(totalvalue(tt% + 12)) = 0 And Val(totalvalue(tt% + 24)) = 0 And Val(totalvalue(tt% + 30)) = 0) Then
                                temperr = 0
                            Else
                                msg = MsgBox(TP$ + " completed must have associated value of None, Recovered, Stolen, or Unknown on property tab.", 48, "Genesis Error Log")
                                totalvalue(tt%).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                        If Val(totalvalue(tt%)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Then
                            If group(tt%).ListIndex = -1 Then
                                msg = MsgBox("Type(Group) and Total Value must be entered.", 48, "Genesis Error Log")
                                group(tt%).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                        If Val(totalvalue(tt% + 18)) > 0 Then
                            If Not IsDate(DATERECOVERED(tt%)) Then
                                msg = MsgBox("A valid DATE RECOVERED must be entered.", 48, "Genesis Error Log")
                                totalvalue(tt% + 18).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                    End If
                End If
            Next tt%
        End If
        
        '===Additional F 5
        '===== Data Element 12
        If tempucr = "250" Then
            TP$ = "Counterfeiting"
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Not rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
                            If Val(totalvalue(tt%)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                msg = MsgBox(TP$ + " not completed must have associated value of None or Unknown on property tab.", 48, "Genesis Error Log")
                                totalvalue(tt%).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                        If rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
                            If Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) = 0 Or Val(totalvalue(tt% + 24)) = 0 Or Val(totalvalue(tt% + 30)) = 0 Then
                                msg = MsgBox(TP$ + " completed must have associated value of Counterfeited, Recovered, or Seized on property tab.", 48, "Genesis Error Log")
                                totalvalue(tt%).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                        If Val(totalvalue(tt% + 18)) > 0 Then
                            If Not IsDate(DATERECOVERED(tt%)) Then
                                msg = MsgBox("A valid DATE RECOVERED must be entered.", 48, "Genesis Error Log")
                                totalvalue(tt% + 18).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                    End If
                End If
            Next tt%
        End If
        
        '===== Additional F 6
        '===== Data Element 7
        If tempucr = "290" Then
            TP$ = "Destruction"
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Not rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
                            If Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                msg = MsgBox(TP$ + " not completed must have associated value of None or Unknown on property tab.", 48, "Genesis Error Log")
                                totalvalue(tt% + 0).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                        If rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
                            If Val(totalvalue(tt% + 6)) = 0 Or Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                msg = MsgBox(TP$ + " completed must have associated value of Damaged on property tab.", 48, "Genesis Error Log")
                                totalvalue(tt% + 6).SetFocus
                                GoTo exiteditp
                            End If
                            If group(tt%).ListIndex = -1 Then
                                msg = MsgBox("Type(Group) and Total Value must be entered.", 48, "Genesis Error Log")
                                group(tt%).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                    End If
                Else
                'If TT% = 0 Then
                '    msg = MsgBox("Valid property must be entered.", 48, "Genesis Error Log")
                '    GoTo exiteditp
                'End If
                End If
            Next tt%
        End If
        
        '===== Additional F 8
        '===== Additional F 9
        '===== Additional F 10
        '===== Additional F 14
        '===== Additional F 18
        '===== Error 077, 078
        TP$ = ""
        Select Case tempucr
            Case "270", "200", "220", "250", "290", "210", "26A", "26B", "26C", "26D", "26E", "23A", "23B", "23C", "23D", "23E", "23F", "23G", "23H", "240", "120", "280"
                TP$ = "Crimes Against Property"
            Case "39A", "39B", "39C", "39D"
                TP$ = "Gambling"
            Case "100"
                TP$ = "Kidnaping"
            Case "35A", "35B"
                TP$ = "Drug/Narcotic Offenses"
        End Select
        '===== Error 074
'        If TP$ > "" Then
'            If FOUNDPROP = False Then
'                msg = MsgBox(TP$ + " must have property data.", 48, "Genesis Error Log")
'                GoTo exiteditp
'            End If
'        End If
        If TP$ > "" Then
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Not rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
                            If Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                msg = MsgBox(TP$ + " not completed must have associated value of None or Unknown on property tab.", 48, "Genesis Error Log")
                                totalvalue(tt% + 0).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                        If rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
                            If Val(totalvalue(tt%)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Then
                                temperr = 0
                            Else
'                            If tempucr <> "35A" Then
'                                MSG = MsgBox(TP$ + " completed must have associated value of Burned, Recovered, or Stolen on property tab.", 48, "Genesis Error Log")
'                                GoTo exiteditp
'                            End If
                            End If
                            If group(tt%).ListIndex = -1 Then
                                msg = MsgBox("Type(Group) and Total Value must be entered.", 48, "Genesis Error Log")
                                group(tt%).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                        If Val(totalvalue(tt% + 18)) > 0 Then
                            If Not IsDate(DATERECOVERED(tt%)) Then
                                msg = MsgBox("A valid DATE RECOVERED must be entered.", 48, "Genesis Error Log")
                                DATERECOVERED(tt%).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                    End If
                End If
            Next tt%
        End If
        'End If
        
        '===== Additional F 11
        '===== Data Element 7
        If tempucr = "39A" Or tempucr = "39B" Or tempucr = "39C" Or tempucr = "39D" Then
            TP$ = "Gambling Offenses"
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Not rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
                            If Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                msg = MsgBox(TP$ + " not completed must have associated value of None or Unknown on property tab.", 48, "Genesis Error Log")
                                totalvalue(tt% + 0).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                        If rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
                            If Val(totalvalue(tt% + 24)) = 0 Or Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                msg = MsgBox(TP$ + " completed must have associated value of Seized on property tab.", 48, "Genesis Error Log")
                                totalvalue(tt% + 24).SetFocus
                                GoTo exiteditp
                            End If
                            If group(tt%).ListIndex = -1 Then
                                msg = MsgBox("Type(Group) and Total Value must be entered.", 48, "Genesis Error Log")
                                group(tt%).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                    End If
                End If
            Next tt%
        End If
        
        '===== Additional F 13
        '===== Data Element 7
        If tempucr = "100" Then
            TP$ = "Kidnapping"
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Not rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
                            If Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                msg = MsgBox(TP$ + " not completed must have associated value of None or Unknown on property tab.", 48, "Genesis Error Log")
                                totalvalue(tt% + 0).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                        If rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
                            If (Val(totalvalue(tt%)) = 0 And Val(totalvalue(tt% + 6)) = 0 And Val(totalvalue(tt% + 12)) = 0 And Val(totalvalue(tt% + 18)) = 0 And Val(totalvalue(tt% + 24)) = 0 And Val(totalvalue(tt% + 30)) = 0) Or _
                                ((Val(totalvalue(tt%)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 30)) > 0) And Val(totalvalue(tt% + 6)) = 0 And Val(totalvalue(tt% + 12)) = 0 And Val(totalvalue(tt% + 24)) = 0) Then
                                temperr = 0
                            Else
                                msg = MsgBox(TP$ + " completed must have associated value of None, Recovered, Stolen, or Unknown on property tab.", 48, "Genesis Error Log")
                                totalvalue(tt%).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                        If Val(totalvalue(tt%)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Then
                            If group(tt%).ListIndex = -1 Then
                                msg = MsgBox("Type(Group) and Total Value must be entered.", 48, "Genesis Error Log")
                                totalvalue(tt%).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                        If Val(totalvalue(tt% + 18)) > 0 Then
                            If Not IsDate(DATERECOVERED(tt%)) Then
                                msg = MsgBox("A valid DATE RECOVERED must be entered.", 48, "Genesis Error Log")
                                DATERECOVERED(tt%).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                    End If
                End If
            Next tt%
        End If
        
        '===== Additional F 15
        If tempucr = "240" Then
            TP$ = "Motor Vehicle Theft"
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Not rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
                            If Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                msg = MsgBox(TP$ + " not completed must have associated value of None or Unknown on property tab.", 48, "Genesis Error Log")
                                totalvalue(tt% + 0).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                        If rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
                            If (Val(totalvalue(tt%)) > 0 Or Val(totalvalue(tt% + 18)) > 0) And Val(totalvalue(tt% + 6)) = 0 And Val(totalvalue(tt% + 12)) = 0 And Val(totalvalue(tt% + 24)) = 0 And Val(totalvalue(tt% + 30)) = 0 Then
                                temperr = 0
                            Else
                                msg = MsgBox(TP$ + " completed must have associated value of Recovered or Stolen on property tab.", 48, "Genesis Error Log")
                                totalvalue(tt%).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                        If Val(totalvalue(tt% + 18)) > 0 Then
                            If group(tt%).ListIndex = -1 Then
                                msg = MsgBox("Type(Group) and Total Value must be entered.", 48, "Genesis Error Log")
                                group(tt%).SetFocus
                                GoTo exiteditp
                            End If
                            tempgroup = Mid$(group(tt%).List(group(tt%).ListIndex), InStr(group(tt%).List(group(tt%).ListIndex), "(") + 1, 2)
                            If tempgroup = "03" Or tempgroup = "05" Or tempgroup = "24" Or tempgroup = "28" Or tempgroup = "37" Then
                                If numvehicle(tt% + 6) = 0 Then
                                    msg = MsgBox("A number of vehicles recovered must be entered.", 48, "Genesis Error Log")
                                    numvehicle(tt% + 6).SetFocus
                                    GoTo exiteditp
                                End If
                            End If
                            If Not IsDate(DATERECOVERED(tt%)) Then
                                msg = MsgBox("A valid DATE RECOVERED must be entered.", 48, "Genesis Error Log")
                                DATERECOVERED(tt%).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                        If Val(totalvalue(tt%)) > 0 Then
                            If group(tt%).ListIndex = -1 Then
                                msg = MsgBox("Type(Group) and Total Value must be entered.", 48, "Genesis Error Log")
                                group(tt%).SetFocus
                                GoTo exiteditp
                            End If
                            tempgroup = Mid$(group(tt%).List(group(tt%).ListIndex), InStr(group(tt%).List(group(tt%).ListIndex), "(") + 1, 2)
                            If tempgroup = "03" Or tempgroup = "05" Or tempgroup = "24" Or tempgroup = "28" Or tempgroup = "37" Then
                                If numvehicle(tt%) = 0 Then
                                    msg = MsgBox("A number of vehicles stolen must be entered.", 48, "Genesis Error Log")
                                    numvehicle(tt%).SetFocus
                                    GoTo exiteditp
                                End If
                            Else
                                msg = MsgBox("Invalid type (group) entered for Motor Vehicle Theft crime.", 48, "Genesis Error Log")
                                numvehicle(tt%).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                    End If
                End If
            Next tt%
        End If
        
        '===== Additional F 21
        '===== Data Element 12
        If tempucr = "280" Then
            TP$ = "Stolen Property"
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Not rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
                            If Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                msg = MsgBox(TP$ + " not completed must have associated value of None or Unknown on property tab.", 48, "Genesis Error Log")
                                totalvalue(tt% + 0).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                        If rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
                            If (Val(totalvalue(tt%)) = 0 And Val(totalvalue(tt% + 6)) = 0 And Val(totalvalue(tt% + 12)) = 0 And Val(totalvalue(tt% + 18)) = 0 And Val(totalvalue(tt% + 24)) = 0 And Val(totalvalue(tt% + 30)) = 0) Or _
                                (Val(totalvalue(tt% + 18)) > 0 And Val(totalvalue(tt%)) = 0 And Val(totalvalue(tt% + 6)) = 0 And Val(totalvalue(tt% + 12)) = 0 And Val(totalvalue(tt% + 24)) = 0 And Val(totalvalue(tt% + 30)) = 0) Then
                                temperr = 0
                            Else
                                msg = MsgBox(TP$ + " completed must have associated value of None or Recovered on property tab.", 48, "Genesis Error Log")
                                totalvalue(tt% + 18).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                        If Val(totalvalue(tt% + 18)) > 0 Then
                            If group(tt%).ListIndex = -1 Then
                                msg = MsgBox("Type(Group) and Total Value must be entered.", 48, "Genesis Error Log")
                                totalvalue(tt%).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                        If Val(totalvalue(tt% + 18)) > 0 Then
                            If Not IsDate(DATERECOVERED(tt%)) Then
                                msg = MsgBox("A valid DATE RECOVERED must be entered.", 48, "Genesis Error Log")
                                DATERECOVERED(tt%).SetFocus
                                GoTo exiteditp
                            End If
                        End If
                    End If
                End If
            Next tt%
        End If
        
        For tt% = 0 To 5
            
            '===== Data Element 14
            If tempucr <> "35A" Then
                If Val(totalvalue(tt% + 36)) > 0 Or _
                    (Val(totalvalue(tt% + 0)) = 0 And Val(totalvalue(tt% + 6)) = 0 And Val(totalvalue(tt% + 12)) = 0 And Val(totalvalue(tt% + 18)) = 0 And Val(totalvalue(tt% + 24)) = 0 And Val(totalvalue(tt% + 30)) = 0) Then
                    If group(tt%).ListIndex > -1 Or pucrlist(tt%).ListIndex > -1 Or description(tt%) > "" Then
                        msg = MsgBox("If UNKNOWN or Nothing selected in entry of PROPERTY tab, no other associated values may be selected (i.e. Type, Value, etc.).", 48, "Genesis Error Log")
                        group(tt%).SetFocus
                        GoTo exiteditp
                    End If
                End If
            End If
        
        Next tt%

        
    
        
    End If
Next t%
        
GoTo goodeditp
exiteditp:
editerr = 1
goodeditp:
On Error Resume Next
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub

Private Sub deleteroutine()
If incidentnumber = "" Then
    Exit Sub
End If
If Len(incidentnumber) < 12 And Len(incidentnumber) > 0 Then
    incidentnumber = incidentnumber + Space$(12 - Len(incidentnumber))
End If
msg = MsgBox("Are you sure you wish to delete this supplemental incident report page?", 4, "Genesis Information Log")
If msg <> 6 Then
    Exit Sub
End If
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(networkpath + "incident.mdb")
Set rs = db.OpenRecordset("select * from INCIDENTsuppORT where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
End If
On Error Resume Next
If IsDate(rs("exportdate")) Then
    msg = MsgBox("This incident report cannot be deleted because it has already been exported.", 48, "Genesis Error Log")
    db.Close
    Exit Sub
End If
Set rs = db.OpenRecordset("select * from supplemental where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34) + " and page = " + PAGE)
If Not rs.EOF Then
    rs.MoveFirst
End If
While Not rs.EOF
    rs.Delete
    rs.MoveNext
Wend
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub

Private Sub findincident()
Dim db As Database, rs, rs2 As Recordset, ecc As Integer, lu As String
On Error GoTo oderror
od:
Set db = OpenDatabase(networkpath + "incident.mdb")
Set rs = db.OpenRecordset("Select * from supplemental where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34) + " AND PAGE = " + PAGE)
If Not rs.EOF Then
    rs.MoveFirst
    On Error Resume Next
Else
    db.Close
    Exit Sub
End If
On Error Resume Next
fromfind = 1
incidentnumber = rs("incidentnumber")
LOCATIONNUMBER(0) = rs("locatioNnumber1")
LOCATIONNUMBER(1) = rs("locatioNnumber2")
vsname(0) = rs("name1")
address(0) = rs("address1")
city(0) = rs("city1")
state(0) = rs("state1")
zipcode(0) = rs("zipcode1")
computerequipment(0) = rs("computerequipment1")
computerequipment(1) = rs("computerequipment2")
'Ian's code
NARRATIVE.Text = rs("narrative")
'end Ian's code
reportingofficer(0) = rs("reportingofficer1")
If Not IsNull(rs("reportingdate1")) Then
    REPORTINGOFFICERDATE(0) = rs("reportingdate1")
End If
reportingofficeRunit(0) = rs("reportingunit1")
If Not IsNull(rs("reportingofficer2")) Then
    reportingofficer(1) = rs("reportingofficer2")
End If
If Not IsNull(rs("reportingdate2")) Then
    REPORTINGOFFICERDATE(1) = rs("reportingdate2")
End If
If Not IsNull(rs("reportingunit2")) Then
    reportingofficeRunit(1) = rs("reportingunit2")
End If
approvingofficer = rs("approvingofficer")
If Not IsNull(rs("approvingdate")) Then
    APPROVINGOFFICERDATE = rs("approvingdate")
End If
approvingofficeRunit = rs("approvingunit")
TV1 = ""
TV2 = ""
If rs("individual1") Then
    TV1 = "I"
End If
If rs("business1") Then
    TV1 = "B"
End If
If rs("financialinstitution1") Then
    TV1 = "F"
End If
If rs("government1") Then
    TV1 = "G"
End If
If rs("religiousorganization1") Then
    TV1 = "R"
End If
If rs("societypublic1") Then
    TV1 = "S"
End If
If rs("tvother1") Then
    TV1 = "O"
End If
If rs("unknown1") Then
    TV1 = "U"
End If
If rs("policeofficer1") Then
    TV1 = "P"
End If
If rs("individual2") Then
    TV2 = "I"
End If
If rs("business2") Then
    TV2 = "B"
End If
If rs("financialinstitution2") Then
    TV2 = "F"
End If
If rs("government2") Then
    TV2 = "G"
End If
If rs("religiousorganization2") Then
    TV2 = "R"
End If
If rs("societypublic2") Then
    TV2 = "S"
End If
If rs("tvother2") Then
    TV2 = "O"
End If
If rs("unknown2") Then
    TV2 = "U"
End If
If rs("policeofficer2") Then
    TV2 = "P"
End If
For t% = 0 To resident(0).ListCount - 1
    If Left$(resident(0).List(t%), 1) = rs("resident1") Then
        resident(0).ListIndex = t%
        t% = resident(0).ListCount - 1
    End If
Next t%
For t% = 0 To race(0).ListCount - 1
    If Left$(race(0).List(t%), 1) = rs("RACE1") Then
        race(0).ListIndex = t%
        t% = race(0).ListCount - 1
    End If
Next t%
For t% = 0 To sex(0).ListCount - 1
    If Left$(sex(0).List(t%), 1) = rs("SEX1") Then
        sex(0).ListIndex = t%
        t% = sex(0).ListCount - 1
    End If
Next t%
age(0) = rs("age1")
For t% = 0 To ethnicity(0).ListCount - 1
    If Left$(ethnicity(0).List(t%), 1) = rs("ETHNICITY1") Then
        ethnicity(0).ListIndex = t%
        t% = ethnicity(0).ListCount - 1
    End If
Next t%
For t% = 0 To 2
    For tt% = 0 To relationship(t%).ListCount - 1
        relationship(t%).Selected(tt%) = False
    Next tt%
    relationship(t%).ListIndex = -1
    If rs("relationship1" + Mid$(Str$(t% + 1), 2)) > "" Then
        For tt% = 0 To relationship(t%).ListCount - 1
            If rs("relationship1" + Mid$(Str$(t% + 1), 2)) = Mid$(relationship(t%).List(tt%), InStr(relationship(t%).List(tt%), "(") + 1, 2) Then
                relationship(t%).ListIndex = tt%
                relationship(t%).Selected(tt%) = True
                tt% = relationship(t%).ListCount - 1
            End If
        Next tt%
    End If
Next t%
HOMEDAYPHONE(0) = rs("homephoneDAY1")
WORKDAYPHONE(0) = rs("WORKphoneDAY1")
HOMENIGHTPHONE(0) = rs("HOMEphoneNIGHT1")
WORKNIGHTPHONE(0) = rs("WORKphoneNIGHT1")
For s% = 0 To 1
    For t% = 1 To injury(s%).ListItems.Count
        For tt% = 1 To 5
            If Not IsNull(rs("typeofinjury" + Mid$(Str$(s% + 1), 2) + Mid$(Str$(tt%), 2))) Then
                If rs("typeofinjury" + Mid$(Str$(s% + 1), 2) + Mid$(Str$(tt%), 2)) = Mid$(injury(s%).ListItems(t%), InStr(injury(s%).ListItems(t%), "(") + 1, 1) Then
                    injury(s%).ListItems(t%).Selected = True
                End If
            End If
        Next tt%
    Next t%
Next s%
If rs("visibleinjuryyes1") = "X" Then
    VISIBLEINJURYYES(0) = True
End If
If rs("visibleinjuryno1") = "X" Then
    VISIBLEINJURYNO(0) = True
End If
If rs("NONVISibleinjuryyes1") = "X" Then
    NONVISIBLEINJURYYES(0) = True
End If
If rs("NONVISibleinjuryno1") = "X" Then
    NONVISIBLEINJURYNO(0) = True
End If
ht(0) = rs("HEIGHT1")
weight(0) = rs("WEIGHT1")
hair(0) = rs("HAIR1")
eyes(0) = rs("EYES1")
peculiarities(0) = rs("PECULIARITIES1")
If rs("valcoholyes1") = "X" Then
    alcoholyes(0) = True
End If
If rs("valcoholNO1") = "X" Then
    alcoholno(0) = True
End If
If rs("valcoholUNKNOWN1") = "X" Then
    alcoholunknown(0) = True
End If
If rs("vDRUGSyes1") = "X" Then
    drugsyes(0) = True
End If
If rs("vDRUGSNO1") = "X" Then
    drugsno(0) = True
End If
If rs("vDRUGSUNKNOWN1") = "X" Then
    drugsunknown(0) = True
End If
If rs("salcoholyes1") = "X" Then
    alcoholyes(1) = True
End If
If rs("salcoholNO1") = "X" Then
    alcoholno(1) = True
End If
If rs("salcoholUNKNOWN1") = "X" Then
    alcoholunknown(1) = True
End If
If rs("sDRUGSyes1") = "X" Then
    drugsyes(1) = True
End If
If rs("sDRUGSNO1") = "X" Then
    drugsno(1) = True
End If
If rs("sDRUGSUNKNOWN1") = "X" Then
    drugsunknown(1) = True
End If


vsname(1) = rs("name2")
address(1) = rs("address2")
city(1) = rs("city2")
state(1) = rs("state2")
zipcode(1) = rs("zipcode2")
ht(1) = rs("HEIGHT2")
weight(1) = rs("WEIGHT2")
hair(1) = rs("HAIR2")
eyes(1) = rs("EYES2")
peculiarities(1) = rs("PECULIARITIES2")
For t% = 0 To resident(1).ListCount - 1
    If Left$(resident(1).List(t%), 1) = rs("resident2") Then
        resident(1).ListIndex = t%
        t% = resident(1).ListCount - 1
    End If
Next t%
For t% = 0 To race(1).ListCount - 1
    If Left$(race(1).List(t%), 1) = rs("RACE2") Then
        race(1).ListIndex = t%
        t% = race(1).ListCount - 1
    End If
Next t%
For t% = 0 To sex(1).ListCount - 1
    If Left$(sex(1).List(t%), 1) = rs("SEX2") Then
        sex(1).ListIndex = t%
        t% = sex(1).ListCount - 1
    End If
Next t%
age(1) = rs("age2")
For t% = 0 To ethnicity(1).ListCount - 1
    If Left$(ethnicity(1).List(t%), 1) = rs("ETHNICITY2") Then
        ethnicity(1).ListIndex = t%
        t% = ethnicity(1).ListCount - 1
    End If
Next t%
For t% = 10 To 12
    For tt% = 0 To relationship(t%).ListCount - 1
        relationship(t%).Selected(tt%) = False
    Next tt%
    relationship(t%).ListIndex = -1
    If rs("relationship2" + Mid$(Str$(t% - 9), 2)) > "" Then
        For tt% = 0 To relationship(t%).ListCount - 1
            If rs("relationship2" + Mid$(Str$(t% - 9), 2)) = Mid$(relationship(t%).List(tt%), InStr(relationship(t%).List(tt%), "(") + 1, 2) Then
                relationship(t%).ListIndex = tt%
                relationship(t%).Selected(tt%) = True
                tt% = relationship(t%).ListCount - 1
            End If
        Next tt%
    End If
Next t%
If Not IsNull(rs("homephoneDAY2")) Then
    HOMEDAYPHONE(1) = rs("homephoneDAY2")
End If
If Not IsNull(rs("workphoneDAY2")) Then
    WORKDAYPHONE(1) = rs("workphoneDAY2")
End If
If Not IsNull(rs("homephoneNIGHT2")) Then
    HOMENIGHTPHONE(1) = rs("homephoneNIGHT2")
End If
If Not IsNull(rs("workphoneNIGHT2")) Then
    WORKNIGHTPHONE(1) = rs("workphoneNIGHT2")
End If
If rs("visibleinjuryyes2") = "X" Then
    VISIBLEINJURYYES(1) = True
End If
If rs("visibleinjuryno2") = "X" Then
    VISIBLEINJURYNO(1) = True
End If
If rs("NONVISibleinjuryyes2") = "X" Then
    NONVISIBLEINJURYYES(1) = True
End If
If rs("NONVISibleinjuryno2") = "X" Then
    NONVISIBLEINJURYNO(1) = True
End If
If rs("valcoholyes2") = "X" Then
    alcoholyes(2) = True
End If
If rs("valcoholNO2") = "X" Then
    alcoholno(2) = True
End If
If rs("valcoholUNKNOWN2") = "X" Then
    alcoholunknown(2) = True
End If
If rs("vDRUGSyes2") = "X" Then
    drugsyes(2) = True
End If
If rs("vDRUGSNO2") = "X" Then
    drugsno(2) = True
End If
If rs("vDRUGSUNKNOWN2") = "X" Then
    drugsunknown(2) = True
End If
If rs("salcoholyes2") = "X" Then
    alcoholyes(3) = True
End If
If rs("salcoholNO2") = "X" Then
    alcoholno(3) = True
End If
If rs("salcoholUNKNOWN2") = "X" Then
    alcoholunknown(3) = True
End If
If rs("sDRUGSyes2") = "X" Then
    drugsyes(3) = True
End If
If rs("sDRUGSNO2") = "X" Then
    drugsno(3) = True
End If
If rs("sDRUGSUNKNOWN2") = "X" Then
    drugsunknown(3) = True
End If
If rs("twomanvehicle1") = "X" Then
    TWOMANVEHICLE(0) = 1
End If
If rs("onemanvehicle1") = "X" Then
    ONEMANVEHICLE(0) = 1
End If
If rs("twomanvehicle2") = "X" Then
    TWOMANVEHICLE(1) = 1
End If
If rs("onemanvehicle2") = "X" Then
    ONEMANVEHICLE(1) = 1
End If
If Not IsNull(rs("vtypeofdrug11")) Then
    For t% = 0 To drugtype(0).ListCount - 1
        If drugtype(0) = Left$(rs("vtypeofdrug11"), 1) Then
            drugtype(0).ListIndex = t%
            t% = drugtype(0).ListCount
        End If
    Next t%
End If
If Not IsNull(rs("stypeofdrug11")) Then
    For t% = 0 To drugtype(3).ListCount - 1
        If drugtype(3) = Left$(rs("stypeofdrug11"), 1) Then
            drugtype(3).ListIndex = t%
            t% = drugtype(3).ListCount
        End If
    Next t%
End If
If Not IsNull(rs("vtypeofdrug21")) Then
    For t% = 0 To drugtype(6).ListCount - 1
        If drugtype(6) = Left$(rs("vtypeofdrug21"), 1) Then
            drugtype(6).ListIndex = t%
            t% = drugtype(6).ListCount
        End If
    Next t%
End If
If Not IsNull(rs("stypeofdrug21")) Then
    For t% = 0 To drugtype(9).ListCount - 1
        If drugtype(9) = Left$(rs("stypeofdrug21"), 1) Then
            drugtype(9).ListIndex = t%
            t% = drugtype(9).ListCount
        End If
    Next t%
End If

If rs("detective1") = "X" Then
    DETECTIVE(0) = 1
End If
If rs("other1") = "X" Then
    TODOTHER(0) = 1
End If
If rs("alone1") = "X" Then
    ALONE(0) = 1
End If
If rs("assisted1") = "X" Then
    ASSISTED(0) = 1
End If
If rs("detective2") = "X" Then
    DETECTIVE(1) = 1
End If
If rs("other2") = "X" Then
    TODOTHER(1) = 1
End If
If rs("alone2") = "X" Then
    ALONE(1) = 1
End If
If rs("assisted2") = "X" Then
    ASSISTED(1) = 1
End If
If rs("runaway1") = "X" Then
    RUNAWAY(0) = 1
End If
If rs("wanted1") = "X" Then
    WANTED(0) = 1
End If
If rs("arrest1") = "X" Then
    ARREST(0) = 1
End If
If rs("warrant1") = "X" Then
    WARRANT(0) = 1
End If
If rs("jail1") = "X" Then
    JAIL(0) = 1
End If
If rs("summons1") = "X" Then
    SUMMONS(0) = 1
End If
If rs("warrant1") = "X" Then
    WARRANT(0) = 1
End If
If rs("runaway2") = "X" Then
    RUNAWAY(1) = 1
End If
If rs("wanted2") = "X" Then
    WANTED(1) = 1
End If
If rs("arrest2") = "X" Then
    ARREST(1) = 1
End If
If rs("warrant2") = "X" Then
    WARRANT(1) = 1
End If
If rs("jail2") = "X" Then
    JAIL(1) = 1
End If
If rs("summons2") = "X" Then
    SUMMONS(1) = 1
End If
If rs("warrant2") = "X" Then
    WARRANT(1) = 1
End If
For t% = 0 To 41
    Select Case t%
        Case 0 To 5
            If Not IsNull(rs("stolenvalue" + Mid$(Str$((t% Mod 6) + 1), 2))) Then
                totalvalue(t%) = rs("stolenvalue" + Mid$(Str$((t% Mod 6) + 1), 2))
            End If
        Case 6 To 11
            If Not IsNull(rs("damagedvalue" + Mid$(Str$((t% Mod 6) + 1), 2))) Then
                totalvalue(t%) = rs("damagedvalue" + Mid$(Str$((t% Mod 6) + 1), 2))
            End If
        Case 12 To 17
            If Not IsNull(rs("burnedvalue" + Mid$(Str$((t% Mod 6) + 1), 2))) Then
                totalvalue(t%) = rs("burnedvalue" + Mid$(Str$((t% Mod 6) + 1), 2))
            End If
        Case 18 To 23
            If Not IsNull(rs("recoveredvalue" + Mid$(Str$((t% Mod 6) + 1), 2))) Then
                totalvalue(t%) = rs("recoveredvalue" + Mid$(Str$((t% Mod 6) + 1), 2))
            End If
        Case 24 To 29
            If Not IsNull(rs("seizedvalue" + Mid$(Str$((t% Mod 6) + 1), 2))) Then
                totalvalue(t%) = rs("seizedvalue" + Mid$(Str$((t% Mod 6) + 1), 2))
            End If
        Case 30 To 35
            If Not IsNull(rs("COUNTERFEITvalue" + Mid$(Str$((t% Mod 6) + 1), 2))) Then
                totalvalue(t%) = rs("COUNTERFEITvalue" + Mid$(Str$((t% Mod 6) + 1), 2))
            End If
        Case 36 To 41
            If Not IsNull(rs("UNKNOWNvalue" + Mid$(Str$((t% Mod 6) + 1), 2))) Then
                totalvalue(t%) = rs("UNKNOWNvalue" + Mid$(Str$((t% Mod 6) + 1), 2))
            End If
    End Select
Next t%
For t% = 0 To 5
    If Not IsNull(rs("type" + Mid$(Str$(t% + 1), 2))) Then
        description(t%) = rs("type" + Mid$(Str$(t% + 1), 2))
    End If
Next t%
ORiGINAL = rs("original")
MODIFIES = rs("modifies")
SUPPLEMENTAL = rs("supplemental")
CASEst = rs("case")
ADDITIONALV = rs("additionalv")
additionalo = rs("additionalo")
additions = rs("additionals")
ADDITIONALR = rs("additionalr")
If rs("complainant1") = "X" Then
    complainant(0) = 1
End If
If rs("complainant2") = "X" Then
    complainant(1) = 1
End If
If Not IsNull(rs("victim1")) Then
    victim(0) = rs("victim1")
End If
If Not IsNull(rs("victim2")) Then
    victim(1) = rs("victim2")
End If
If Not IsNull(rs("subject1")) Then
    subject(0) = rs("subject1")
End If
If Not IsNull(rs("subject2")) Then
    subject(1) = rs("subject2")
End If
typeother(0) = rs("typeother1")
typeother(1) = rs("typeother2")
If Not IsNull(rs("birthdate1")) Then
    BIRTHDATE(0) = rs("birthdate1")
End If
If Not IsNull(rs("birthdate2")) Then
    BIRTHDATE(1) = rs("birthdate2")
End If

'---support
Set rs = db.OpenRecordset("Select * from supplementalsupport where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34) + " AND PAGE = " + PAGE)
If Not rs.EOF Then
    rs.MoveFirst
Else
    GoTo vehgun
End If
For r% = 0 To 1
  For t% = 1 To vucrlist(r%).ListItems.Count
    If Mid$(vucrlist(r%).ListItems(t%), InStr(vucrlist(r%).ListItems(t%), "(") + 1, 3) = rs("vucr" + Mid$(Str$(r% + 1), 2) + "1") Or _
       Mid$(vucrlist(r%).ListItems(t%), InStr(vucrlist(r%).ListItems(t%), "(") + 1, 3) = rs("vucr" + Mid$(Str$(r% + 1), 2) + "2") Or _
       Mid$(vucrlist(r%).ListItems(t%), InStr(vucrlist(r%).ListItems(t%), "(") + 1, 3) = rs("vucr" + Mid$(Str$(r% + 1), 2) + "3") Or _
       Mid$(vucrlist(r%).ListItems(t%), InStr(vucrlist(r%).ListItems(t%), "(") + 1, 3) = rs("vucr" + Mid$(Str$(r% + 1), 2) + "4") Or _
       Mid$(vucrlist(r%).ListItems(t%), InStr(vucrlist(r%).ListItems(t%), "(") + 1, 3) = rs("vucr" + Mid$(Str$(r% + 1), 2) + "5") Or _
       Mid$(vucrlist(r%).ListItems(t%), InStr(vucrlist(r%).ListItems(t%), "(") + 1, 3) = rs("vucr" + Mid$(Str$(r% + 1), 2) + "6") Or _
       Mid$(vucrlist(r%).ListItems(t%), InStr(vucrlist(r%).ListItems(t%), "(") + 1, 3) = rs("vucr" + Mid$(Str$(r% + 1), 2) + "7") Or _
       Mid$(vucrlist(r%).ListItems(t%), InStr(vucrlist(r%).ListItems(t%), "(") + 1, 3) = rs("vucr" + Mid$(Str$(r% + 1), 2) + "8") Or _
       Mid$(vucrlist(r%).ListItems(t%), InStr(vucrlist(r%).ListItems(t%), "(") + 1, 3) = rs("vucr" + Mid$(Str$(r% + 1), 2) + "9") Or _
       Mid$(vucrlist(r%).ListItems(t%), InStr(vucrlist(r%).ListItems(t%), "(") + 1, 3) = rs("vucr" + Mid$(Str$(r% + 1), 2) + "10") Then
        vucrlist(r%).ListItems(t%).Selected = True
    Else
        vucrlist(r%).ListItems(t%).Selected = fasle
    End If
  Next t%
Next r%
For r% = 0 To 5
    pucrlist(r%).ListIndex = -1
    For rr% = 0 To pucrlist(r%).ListCount - 1
        If Mid$(pucrlist(r%).List(rr%), InStr(pucrlist(r%).List(rr%), "(") + 1, 3) = rs("PUCR" + Mid$(Str$(r% + 1), 2)) Then
            pucrlist(r%).ListIndex = rr%
            rr% = pucrlist(r%).ListCount - 1
        End If
    Next rr%
Next r%
ct% = 0
For t% = 12 To 29 Step 3
    ct% = ct% + 1
    If ct% > 3 Then
        ct% = 1
    End If
    st% = (t% Mod 3) + ct%
    For tt% = 1 To 3
        If Not IsNull(rs("PTYPEOFDRUG" + Mid$(Str$(ct%), 2) + Mid$(Str$(tt%), 2))) Then
            For TTT% = 0 To drugtype(t%).ListCount - 1
                If rs("PTYPEOFDRUG" + Mid$(Str$(ct%), 2) + Mid$(Str$(tt%), 2)) = Left$(drugtype(t%).List(TTT%), 1) Then
                    drugtype(t%).ListIndex = TTT%
                    TTT% = drugtype(t%).ListCount - 1
                End If
            Next TTT%
            drugamt(t%) = rs("pdrugamt" + Mid$(Str$(ct%), 2) + Mid$(Str$(tt%), 2))
            For TTT% = 0 To drugmeasurement(t%).ListCount - 1
                If rs("pdrugmeasurement" + Mid$(Str$(ct%), 2) + Mid$(Str$(tt%), 2)) = Left$(drugmeasurement(t%).List(TTT%), 1) Then
                    drugmeasurement(t%).ListIndex = TTT%
                    TTT% = drugmeasurement(t%).ListCount - 1
                End If
            Next TTT%
        End If
    Next tt%
Next t%
For t% = 0 To 5
    If Not IsNull(rs("group" + Mid$(Str$(t% + 1), 2))) Then
        For tt% = 0 To group(t%).ListCount - 1
            If Mid$(group(t%).List(tt%), InStr(group(t%).List(tt%), "(") + 1, 2) = rs("group" + Mid$(Str$(t% + 1), 2)) Then
                group(t%).ListIndex = tt%
                tt% = group(t%).ListCount - 1
            End If
        Next tt%
    Else
        group(t%).ListIndex = -1
    End If
    group(t%).Visible = False
    If Not IsNull(rs("numvehicles" + Mid$(Str$(t% + 1), 2))) Then
        numvehicle(t%) = rs("numvehicles" + Mid$(Str$(t% + 1), 2))
    End If
        If Not IsNull(rs("daterecovered" + Mid$(Str$(t% + 1), 2))) Then
            DATERECOVERED(t%) = rs("daterecovered" + Mid$(Str$(t% + 1), 2))
        End If
Next t%
For t% = 6 To 11
    If Not IsNull(rs("numvehicles" + Mid$(Str$(t% - 5), 2))) Then
        numvehicle(t%) = rs("numvehicler" + Mid$(Str$(t% - 5), 2))
    End If
Next t%

For t% = 1 To 2
    For tt% = 3 To 9
        For TTT% = 0 To relationship(((t% - 1) * 10) + tt%).ListCount - 1
            relationship(((t% - 1) * 10) + tt%).Selected(TTT%) = False
        Next TTT%
        If Not IsNull(rs("relationship" + Mid$(Str$(t%), 2) + Mid$(Str$(tt% + 1), 2))) Then
            For TTT% = 0 To relationship(((t% - 1) * 10) + tt%).ListCount - 1
                If rs("relationship" + Mid$(Str$(t%), 2) + Mid$(Str$(tt% + 1), 2)) = Mid$(relationship(((t% - 1) * 10) + tt%).List(TTT%), InStr(relationship(((t% - 1) * 10) + tt%).List(TTT%), "(") + 1, 2) Then
                    relationship(((t% - 1) * 10) + tt%).ListIndex = TTT%
                    relationship(((t% - 1) * 10) + tt%).Selected(TTT%) = True
                    TTT% = relationship(tt%).ListCount - 1
                End If
            Next TTT%
        Else
            relationship(((t% - 1) * 10) + tt%).ListIndex = -1
        End If
    Next tt%
Next t%
If Not IsNull(rs("vtypeofdrug12")) Then
    For t% = 0 To drugtype(1).ListCount - 1
        If drugtype(1) = Left$(rs("vtypeofdrug12"), 1) Then
            drugtype(1).ListIndex = t%
            t% = drugtype(1).ListCount
        End If
    Next t%
End If
If Not IsNull(rs("vtypeofdrug13")) Then
    For t% = 0 To drugtype(2).ListCount - 1
        If drugtype(2) = Left$(rs("vtypeofdrug13"), 1) Then
            drugtype(2).ListIndex = t%
            t% = drugtype(2).ListCount
        End If
    Next t%
End If
If Not IsNull(rs("stypeofdrug12")) Then
    For t% = 0 To drugtype(4).ListCount - 1
        If drugtype(4) = Left$(rs("stypeofdrug12"), 1) Then
            drugtype(4).ListIndex = t%
            t% = drugtype(4).ListCount
        End If
    Next t%
End If
If Not IsNull(rs("stypeofdrug13")) Then
    For t% = 0 To drugtype(5).ListCount - 1
        If drugtype(5) = Left$(rs("stypeofdrug13"), 1) Then
            drugtype(5).ListIndex = t%
            t% = drugtype(5).ListCount
        End If
    Next t%
End If
If Not IsNull(rs("vtypeofdrug22")) Then
    For t% = 0 To drugtype(7).ListCount - 1
        If drugtype(7) = Left$(rs("vtypeofdrug22"), 1) Then
            drugtype(7).ListIndex = t%
            t% = drugtype(7).ListCount
        End If
    Next t%
End If
If Not IsNull(rs("vtypeofdrug23")) Then
    For t% = 0 To drugtype(8).ListCount - 1
        If drugtype(8) = Left$(rs("vtypeofdrug23"), 1) Then
            drugtype(8).ListIndex = t%
            t% = drugtype(8).ListCount
        End If
    Next t%
End If
If Not IsNull(rs("stypeofdrug22")) Then
    For t% = 0 To drugtype(10).ListCount - 1
        If drugtype(10) = Left$(rs("stypeofdrug22"), 1) Then
            drugtype(10).ListIndex = t%
            t% = drugtype(10).ListCount
        End If
    Next t%
End If
If Not IsNull(rs("stypeofdrug23")) Then
    For t% = 0 To drugtype(11).ListCount - 1
        If drugtype(11) = Left$(rs("stypeofdrug23"), 1) Then
            drugtype(11).ListIndex = t%
            t% = drugtype(11).ListCount
        End If
    Next t%
End If

vehgun:
'---vehgun
Set rs = db.OpenRecordset("Select * from vehgun where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34) + " AND PAGE = " + PAGE)
If Not rs.EOF Then
    rs.MoveFirst
Else
    db.Close
    fromfind = 0
    Exit Sub
End If
SSTOLEN = rs("stolen")
SRECOVERED = rs("recovered")
SFOUND = rs("found")
STOWED = rs("towed")
SSUSPECT = rs("suspect")
SVICTIM = rs("victim")
SVEHICLE = rs("vehicle")
SGUN = rs("gun")
SBOAT = rs("boat")
sLICENSEPLATE = rs("licenseplate")
SSECURITIES = rs("securities")
SARTICLE = rs("article")
VIN = rs("vin")
HULL = rs("hull")
SERIAL = rs("serial")
SERIALSTATE = rs("serialstate")
YEARREG = rs("yearreg")
YEAREXP = rs("yearexp")
YEARN = rs("year")
MAKE = rs("make")
stype = rs("type")
MODEL = rs("model")
STYLE = rs("style")
scolor = rs("color")
BRANDNAME = rs("brandname")
CALIBER = rs("caliber")
NIC = rs("nic")
DENOMINATION = rs("denomination")
ISSUER = rs("issuer")
If Not IsNull(rs("securitiesdate")) Then
    SECURITIESDATE = rs("securitiesdate")
Else
    SECURITIESDATE = ""
End If
MISCELLANEOUS = rs("miscellaneous")


fromfind = 0
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub

Private Sub saveincident()
Dim db As Database, rs, rs2 As Recordset, lu As String
On Error GoTo oderror1
od1:
Set db = OpenDatabase(networkpath + "incident.mdb")
Set rs = db.OpenRecordset("Select * from SUPPLEMENTAL where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34) + " and page = " + PAGE)
ecc = 0
If Not rs.EOF Then
    rs.MoveFirst
    If rs("schanged") = 1 Then
        schanged = 1
    Else
        schanged = 0
    End If
    rs.Edit
Else
    rs.AddNew
    schanged = 1
End If
schanged = 1
rs("PAGE") = Val(PAGE)
rs("incidentnumber") = incidentnumber
rs("computerequipment1") = computerequipment(0)
rs("computerequipment2") = computerequipment(1)
'ian's code
rs("narrative") = NARRATIVE.Text
'end ian's code
rs("schanged") = schanged
rs("locationnumber1") = LOCATIONNUMBER(0)
rs("locationnumber2") = LOCATIONNUMBER(1)
rs("name1") = vsname(0)
rs("address1") = address(0)
rs("city1") = city(0)
rs("state1") = state(0)
rs("zipcode1") = zipcode(0)
rs("resident1") = Left$(resident(0).List(resident(0).ListIndex), 1)
rs("race1") = Left$(race(0).List(race(0).ListIndex), 1)
rs("sex1") = Left$(sex(0).List(sex(0).ListIndex), 1)
rs("age1") = age(0)
rs("ethnicity1") = Left$(ethnicity(0).List(ethnicity(0).ListIndex), 1)
For t% = 0 To 2
    If relationship(t%).ListIndex > -1 Then
        rs("relationship1" + Mid$(Str$(t% + 1), 2)) = Mid$(relationship(t%).List(relationship(t%).ListIndex), InStr(relationship(t%).List(relationship(t%).ListIndex), "(") + 1, 2)
    Else
        rs("relationship1" + Mid$(Str$(t% + 1), 2)) = Null
    End If
Next t%
rs("reportingofficer1") = reportingofficer(0)
If IsDate(REPORTINGOFFICERDATE(0)) Then
    rs("reportingdate1") = REPORTINGOFFICERDATE(0)
End If
rs("reportingunit1") = reportingofficeRunit(0)
If reportingofficer(1) > "" Then
    rs("reportingofficer2") = reportingofficer(1)
    If IsDate(REPORTINGOFFICERDATE(1)) Then
        rs("reportingdate2") = REPORTINGOFFICERDATE(1)
    End If
    rs("reportingunit2") = reportingofficeRunit(1)
End If
rs("approvingofficer") = approvingofficer
If IsDate(APPROVINGOFFICERDATE) Then
    rs("approvingdate") = APPROVINGOFFICERDATE
End If
rs("approvingunit") = approvingofficeRunit
rs("homephoneDAY1") = HOMEDAYPHONE(0)
rs("WORKphoneDAY1") = WORKDAYPHONE(0)
rs("HOMEphoneNIGHT1") = HOMENIGHTPHONE(0)
rs("WORKphoneNIGHT1") = WORKNIGHTPHONE(0)
rs("HEIGHT1") = ht(0)
rs("WEIGHT1") = weight(0)
rs("HAIR1") = hair(0)
rs("EYES1") = eyes(0)
rs("PECULIARITIES1") = peculiarities(0)
For s% = 1 To 2
    For t% = 1 To 5
        rs("typeofinjury" + Mid$(Str$(s%), 2) + Mid$(Str$(t%), 2)) = ""
    Next t%
Next s%
For s% = 0 To 1
    IIDX% = 0
    For t% = 1 To injury(s%).ListItems.Count
        If injury(s%).ListItems(t%).Selected = True Then
            IIDX% = IIDX% + 1
            If IIDX% < 6 Then
                rs("typeofinjury" + Mid$(Str$(s% + 1), 2) + Mid$(Str$(IIDX%), 2)) = Mid$(injury(s%).ListItems(t%), InStr(injury(s%).ListItems(t%), "(") + 1, 1)
            Else
                t% = injury(s%).ListItems.Count
            End If
        End If
    Next t%
Next s%
If VISIBLEINJURYYES(0) Then
    rs("visibleinjuryyes1") = "X"
Else
    rs("visibleinjuryno1") = "X"
End If
If NONVISIBLEINJURYYES(0) Then
    rs("nonvisibleinjuryyes1") = "X"
Else
    rs("nonvisibleinjuryno1") = "X"
End If
rs("name2") = vsname(1)
rs("address2") = address(1)
rs("city2") = city(1)
rs("state2") = state(1)
rs("zipcode2") = zipcode(1)
rs("resident2") = Left$(resident(1).List(resident(1).ListIndex), 1)
rs("race2") = Left$(race(1).List(race(1).ListIndex), 1)
rs("sex2") = Left$(sex(1).List(sex(1).ListIndex), 1)
rs("age2") = age(1)
rs("ethnicity2") = Left$(ethnicity(1).List(ethnicity(1).ListIndex), 1)
For t% = 10 To 12
    If relationship(t%).ListIndex > -1 Then
        rs("relationship2" + Mid$(Str$(t% - 9), 2)) = Mid$(relationship(t%).List(relationship(t%).ListIndex), InStr(relationship(t%).List(relationship(t%).ListIndex), "(") + 1, 2)
    Else
        rs("relationship2" + Mid$(Str$(t% - 9), 2)) = Null
    End If
Next t%
rs("homephoneDAY2") = HOMEDAYPHONE(1)
rs("WORKphoneDAY2") = WORKDAYPHONE(1)
rs("HOMEphoneNIGHT2") = HOMENIGHTPHONE(1)
rs("WORKphoneNIGHT2") = WORKNIGHTPHONE(1)
rs("HEIGHT2") = ht(1)
rs("WEIGHT2") = weight(1)
rs("HAIR2") = hair(1)
rs("EYES2") = eyes(1)
rs("PECULIARITIES2") = peculiarities(1)
If VISIBLEINJURYYES(1) Then
    rs("visibleinjuryyes2") = "X"
Else
    rs("visibleinjuryno2") = "X"
End If
If NONVISIBLEINJURYYES(1) Then
    rs("nonvisibleinjuryyes2") = "X"
Else
    rs("nonvisibleinjuryno2") = "X"
End If

If alcoholyes(0) Then
    rs("valcoholyes1") = "X"
Else
If alcoholno(0) Then
    rs("valcoholno1") = "X"
Else
If alcoholunknown(0) Then
    rs("valcoholunknown1") = "X"
End If
End If
End If
If drugsyes(0) Then
    rs("vdrugsyes1") = "X"
Else
If drugsno(0) Then
    rs("vdrugsno1") = "X"
Else
If drugsunknown(0) Then
    rs("vdrugsunknown1") = "X"
End If
End If
End If
If alcoholyes(1) Then
    rs("salcoholyes1") = "X"
Else
If alcoholno(1) Then
    rs("salcoholno1") = "X"
Else
If alcoholunknown(1) Then
    rs("salcoholunknown1") = "X"
End If
End If
End If
If drugsyes(1) Then
    rs("sdrugsyes1") = "X"
Else
If drugsno(1) Then
    rs("sdrugsno1") = "X"
Else
If drugsunknown(1) Then
    rs("sdrugsunknown1") = "X"
End If
End If
End If
If alcoholyes(2) Then
    rs("valcoholyes2") = "X"
Else
If alcoholno(2) Then
    rs("valcoholno2") = "X"
Else
If alcoholunknown(2) Then
    rs("valcoholunknown2") = "X"
End If
End If
End If
If drugsyes(2) Then
    rs("vdrugsyes2") = "X"
Else
If drugsno(2) Then
    rs("vdrugsno2") = "X"
Else
If drugsunknown(2) Then
    rs("vdrugsunknown2") = "X"
End If
End If
End If
If alcoholyes(3) Then
    rs("salcoholyes2") = "X"
Else
If alcoholno(3) Then
    rs("salcoholno2") = "X"
Else
If alcoholunknown(3) Then
    rs("salcoholunknown2") = "X"
End If
End If
End If
If drugsyes(3) Then
    rs("sdrugsyes2") = "X"
Else
If drugsno(3) Then
    rs("sdrugsno2") = "X"
Else
If drugsunknown(3) Then
    rs("sdrugsunknown2") = "X"
End If
End If
End If

If TWOMANVEHICLE(0) = 1 Then
    rs("twomanvehicle1") = "X"
End If
If ONEMANVEHICLE(0) = 1 Then
    rs("onemanvehicle1") = "X"
End If
If DETECTIVE(0) = 1 Then
    rs("detective1") = "X"
End If
If TODOTHER(0) = 1 Then
    rs("other1") = "X"
End If
If ALONE(0) = 1 Then
    rs("alone1") = "X"
End If
If ASSISTED(0) = 1 Then
    rs("assisted1") = "X"
End If

If TWOMANVEHICLE(1) = 1 Then
    rs("twomanvehicle2") = "X"
End If
If ONEMANVEHICLE(1) = 1 Then
    rs("onemanvehicle2") = "X"
End If
If DETECTIVE(1) = 1 Then
    rs("detective2") = "X"
End If
If TODOTHER(1) = 1 Then
    rs("other2") = "X"
End If
If ALONE(1) = 1 Then
    rs("alone2") = "X"
End If
If ASSISTED(1) = 1 Then
    rs("assisted2") = "X"
End If

If drugtype(0).ListIndex > -1 Then
    rs("vtypeofdrug11") = Left$(drugtype(0).List(drugtype(0).ListIndex), 1)
End If
If drugtype(3).ListIndex > -1 Then
    rs("stypeofdrug11") = Left$(drugtype(3).List(drugtype(3).ListIndex), 1)
End If
If drugtype(6).ListIndex > -1 Then
    rs("vtypeofdrug21") = Left$(drugtype(6).List(drugtype(6).ListIndex), 1)
End If
If drugtype(9).ListIndex > -1 Then
    rs("stypeofdrug21") = Left$(drugtype(9).List(drugtype(9).ListIndex), 1)
End If

For yy% = 0 To 1
    rs("runaway" + Mid$(Str$(yy% + 1), 2)) = RUNAWAY(yy%)
    rs("wanted" + Mid$(Str$(yy% + 1), 2)) = WANTED(yy%)
    rs("arrest" + Mid$(Str$(yy% + 1), 2)) = ARREST(yy%)
    rs("warrant" + Mid$(Str$(yy% + 1), 2)) = WARRANT(yy%)
    rs("jail" + Mid$(Str$(yy% + 1), 2)) = JAIL(yy%)
    rs("summons" + Mid$(Str$(yy% + 1), 2)) = SUMMONS(yy%)
Next yy%

For t% = 0 To 41
    tt% = t% Mod 6
    If description(tt%) > "" Then
        rs("type" + Mid$(Str$(tt% + 1), 2)) = description(tt%)
        Select Case t%
            Case 0 To 5
                rs("stolenvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = Val(totalvalue(t%))
            Case 6 To 11
                rs("damagedvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = Val(totalvalue(t%))
            Case 12 To 17
                rs("burnedvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = Val(totalvalue(t%))
            Case 18 To 23
                rs("recoveredvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = Val(totalvalue(t%))
            Case 24 To 29
                rs("seizedvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = Val(totalvalue(t%))
            Case 30 To 35
                rs("COUNTERFEITvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = Val(totalvalue(t%))
            Case 36 To 41
                rs("UNKNOWNvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = Val(totalvalue(t%))
        End Select
    End If
Next t%
rs("original") = ORiGINAL.Value
rs("modifies") = MODIFIES.Value
rs("supplemental") = SUPPLEMENTAL.Value
rs("case") = CASEst.Value
rs("additionalv") = ADDITIONALV.Value
rs("additionalo") = additionalo.Value
rs("additionals") = ADDITIONALS.Value
rs("additionalr") = ADDITIONALR.Value

If complainant(0) = 1 Then
    rs("complainant1") = "X"
End If
If complainant(1) = 1 Then
    rs("complainant2") = "X"
End If
If victim(0) > "" Then
    rs("victim1") = Val(victim(0))
    Select Case UCase(TV1)
        Case "I"
            rs("individual1") = True
        Case "B"
            rs("business1") = True
        Case "F"
            rs("financialinstitution1") = True
        Case "G"
            rs("government1") = True
        Case "R"
            rs("religiousorganization1") = True
        Case "S"
            rs("societypublic1") = True
        Case "O"
            rs("tvother1") = True
        Case "U"
            rs("unknown1") = True
        Case "P"
            rs("policeofficer1") = True
    End Select
Else
    rs("victim1") = Null
End If
If victim(1) > "" Then
    rs("victim2") = Val(victim(1))
    Select Case UCase(TV2)
        Case "I"
            rs("individual2") = True
        Case "B"
            rs("business2") = True
        Case "F"
            rs("financialinstitution2") = True
        Case "G"
            rs("government2") = True
        Case "R"
            rs("religiousorganization2") = True
        Case "S"
            rs("societypublic2") = True
        Case "O"
            rs("tvother2") = True
        Case "U"
            rs("unknown2") = True
        Case "P"
            rs("policeofficer2") = True
    End Select
Else
    rs("victim2") = Null
End If
If subject(0) > "" Then
    rs("subject1") = Val(subject(0))
Else
    rs("subject1") = Null
End If
If subject(1) > "" Then
    rs("subject2") = Val(subject(1))
Else
    rs("subject2") = Null
End If

rs("typeother1") = typeother(0)
rs("typeother2") = typeother(1)
If IsDate(BIRTHDATE(0)) Then
    rs("birthdate1") = BIRTHDATE(0)
End If
If IsDate(BIRTHDATE(1)) Then
    rs("birthdate2") = BIRTHDATE(1)
End If
rs.Update

'---support
Set rs = db.OpenRecordset("Select * from supplementalsupport where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34) + " and page =" + PAGE)
ecc = 0
If Not rs.EOF Then
    rs.MoveFirst
    rs.Edit
Else
    rs.AddNew
End If
schanged = 1
If tempsave = 1 Then
    rs("TEMP") = "Y"
Else
    rs("TEMP") = "N"
End If
rs("INCIDENTNUMBER") = incidentnumber
rs("PAGE") = Val(PAGE)
For t% = 0 To 5
    If numvehicle(t%) > "" Then
        rs("numvehicles" + Mid$(Str$(t% + 1), 2)) = Val(numvehicle(t%))
    End If
    If IsDate(DATERECOVERED(t%)) Then
        rs("daterecovered" + Mid$(Str$(t% + 1), 2)) = DATERECOVERED(t%)
    End If
Next t%
For t% = 6 To 11
    If numvehicle(t%) > "" Then
        rs("numvehicler" + Mid$(Str$(t% - 5), 2)) = Val(numvehicle(t%))
    End If
Next t%
ct% = 0
For t% = 12 To 29 Step 3
    ct% = ct% + 1
    st% = (t% Mod 3) + ct%
    For tt% = 1 To 3
        If drugtype(t%).ListIndex > -1 Then
            rs("PTYPEOFDRUG" + Mid$(Str$(ct%), 2) + Mid$(Str$(tt%), 2)) = Left$(drugtype(t%).List(drugtype(t%).ListIndex), 1)
            rs("pdrugamt" + Mid$(Str$(ct%), 2) + Mid$(Str$(tt%), 2)) = Val(drugamt(t%))
            rs("pdrugmeasurement" + Mid$(Str$(ct%), 2) + Mid$(Str$(tt%), 2)) = Left$(drugmeasurement(t%).List(drugmeasurement(t%).ListIndex), 2)
        End If
    Next tt%
Next t%

If drugtype(1).ListIndex > -1 Then
    rs("vtypeofdrug12") = Left$(drugtype(1).List(drugtype(1).ListIndex), 1)
End If
If drugtype(2).ListIndex > -1 Then
    rs("vtypeofdrug13") = Left$(drugtype(2).List(drugtype(2).ListIndex), 1)
End If
If drugtype(4).ListIndex > -1 Then
    rs("stypeofdrug12") = Left$(drugtype(4).List(drugtype(4).ListIndex), 1)
End If
If drugtype(5).ListIndex > -1 Then
    rs("stypeofdrug13") = Left$(drugtype(5).List(drugtype(5).ListIndex), 1)
End If
If drugtype(7).ListIndex > -1 Then
    rs("vtypeofdrug22") = Left$(drugtype(7).List(drugtype(7).ListIndex), 1)
End If
If drugtype(8).ListIndex > -1 Then
    rs("vtypeofdrug23") = Left$(drugtype(8).List(drugtype(8).ListIndex), 1)
End If
If drugtype(10).ListIndex > -1 Then
    rs("stypeofdrug22") = Left$(drugtype(10).List(drugtype(10).ListIndex), 1)
End If
If drugtype(11).ListIndex > -1 Then
    rs("stypeofdrug23") = Left$(drugtype(11).List(drugtype(11).ListIndex), 1)
End If
For t% = 1 To 2
    tct% = 0
    If Not (vucrlist(t% - 1).SelectedItem Is Nothing) Then
        For tt% = 1 To vucrlist(t% - 1).ListItems.Count
            If vucrlist(t% - 1).ListItems(tt%).Selected = True Then
                tct% = tct% + 1
                rs("vucr" + Mid$(Str$(t%), 2) + Mid$(Str$(tct%), 2)) = Mid$(vucrlist(t% - 1).ListItems(tt%), InStr(vucrlist(t% - 1).ListItems(tt%), "(") + 1, 3)
            End If
        Next tt%
    End If
Next t%
For t% = 1 To 6
    If pucrlist(t% - 1).ListIndex > -1 Then
        rs("Pucr" + Mid$(Str$(t%), 2)) = Mid$(pucrlist(t% - 1).List(pucrlist(t% - 1).ListIndex), InStr(pucrlist(t% - 1).List(pucrlist(t% - 1).ListIndex), "(") + 1, 3)
    End If
Next t%
For t% = 1 To 6
    If group(t% - 1).ListIndex > -1 Then
        rs("group" + Mid$(Str$(t%), 2)) = Mid$(group(t% - 1).List(group(t% - 1).ListIndex), InStr(group(t% - 1).List(group(t% - 1).ListIndex), "(") + 1, 2)
    End If
Next t%
For t% = 1 To 2
    For tt% = 3 To 9
        If relationship(((t% - 1) * 10) + tt%).ListIndex > -1 Then
            rs("relationship" + Mid$(Str$(t%), 2) + Mid$(Str$(tt% + 1), 2)) = Mid$(relationship(((t% - 1) * 10) + tt%).List(relationship(((t% - 1) * 10) + tt%).ListIndex), InStr(relationship(((t% - 1) * 10) + tt%).List(relationship(((t% - 1) * 10) + tt%).ListIndex), "(") + 1, 2)
        Else
            rs("relationship" + Mid$(Str$(t%), 2) + Mid$(Str$(tt% + 1), 2)) = Null
        End If
    Next tt%
Next t%
rs.Update

'---vehgun
If SSTOLEN = 1 Or SRECOVERED = 1 Or SFOUND = 1 Or STOWED = 1 Or SSUSPECT = 1 Or SVICTIM = 1 Or SVEHICLE = 1 Or SGUN = 1 Or SBOAT = 1 Or sLICENSEPLATE = 1 Or SSECURITIES = 1 Or SARTICLE = 1 Or VIN > "" Or HULL > "" Or SERIAL > "" Or SERIALSTATE > "" Or YEARREG > "" Or YEAREXP > "" Or YEARN > "" Or MAKE > "" Or stype > "" Or MODEL > "" Or STYLE > "" Or scolor > "" Or BRANDNAME > "" Or CALIBER > "" Or NIC > "" Or DENOMINATION > "" Or ISSUER > "" Or IsDate(SECURITIESDATE) Or MISCELLANEOUS > "" Then
    Set rs = db.OpenRecordset("Select * from vehgun where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34) + " and page = " + PAGE)
    ecc = 0
    If Not rs.EOF Then
        rs.MoveFirst
        rs.Delete
    End If
    rs.AddNew
    rs("PAGE") = Val(PAGE)
    rs("incidentnumber") = incidentnumber
    rs("stolen") = SSTOLEN
    rs("recovered") = SRECOVERED
    rs("found") = SFOUND
    rs("towed") = STOWED
    rs("suspect") = SSUSPECT
    rs("victim") = SVICTIM
    rs("vehicle") = SVEHICLE
    rs("gun") = SGUN
    rs("boat") = SBOAT
    rs("licenseplate") = sLICENSEPLATE
    rs("securities") = SSECURITIES
    rs("article") = SARTICLE
    rs("vin") = VIN
    rs("hull") = HULL
    rs("serial") = SERIAL
    rs("serialstate") = SERIALSTATE
    rs("yearreg") = YEARREG
    rs("yearexp") = YEAREXP
    rs("year") = YEARN
    rs("make") = MAKE
    rs("type") = stype
    rs("model") = MODEL
    rs("style") = STYLE
    rs("color") = scolor
    rs("brandname") = BRANDNAME
    rs("caliber") = CALIBER
    rs("nic") = NIC
    rs("denomination") = DENOMINATION
    rs("issuer") = ISSUER
    If Not IsDate(SECURITIESDATE) Then
        rs("SECURITIESDATE") = Null
    Else
        rs("securitiesdate") = SECURITIESDATE
    End If
    rs("miscellaneous") = MISCELLANEOUS
    rs.Update
End If

For t% = 0 To 1
    Set rs = db.OpenRecordset("SELECT * FROM CODES WHERE CODE = '" + city(t%) + "' AND TYPE = 'city'")
    If rs.EOF Then
        rs.AddNew
        rs("CODE") = city(t%)
        For tt% = 0 To 1
            city(tt%).AddItem city(t%)
        Next tt%
        rs("TYPE") = "city"
        rs("DEFAULT") = "N"
        rs.Update
    End If
    Set rs = db.OpenRecordset("SELECT * FROM CODES WHERE CODE = '" + state(t%) + "' AND TYPE = 'state'")
    If rs.EOF Then
        rs.AddNew
        rs("CODE") = state(t%)
        For tt% = 0 To 1
            state(tt%).AddItem state(t%)
        Next tt%
        rs("TYPE") = "state"
        rs("DEFAULT") = "N"
        rs.Update
    End If
Next t%
On Error GoTo oderror2
od2:
Set db = OpenDatabase(nwl + "LAWSUITE.MDB")

'----- OFFICERS
If reportingofficer(0) > "" Then
    Set rs = db.OpenRecordset("select profname, type from professionals where profname =" + Chr$(34) + reportingofficer(0) + Chr$(34))
    If rs.EOF Then
        rs.AddNew
        rs("profname") = reportingofficer(0)
        reportingofficer(0).AddItem reportingofficer(0)
        reportingofficer(1).AddItem reportingofficer(0)
        approvingofficer.AddItem reportingofficer(0)
        followupofficer.AddItem reportingofficer(0)
        rs("type") = "D"
        rs.Update
    End If
End If
If reportingofficer(1) > "" Then
    Set rs = db.OpenRecordset("select profname, type from professionals where profname =" + Chr$(34) + reportingofficer(1) + Chr$(34))
    If rs.EOF Then
        rs.AddNew
        rs("profname") = reportingofficer(1)
        reportingofficer(0).AddItem reportingofficer(1)
        reportingofficer(1).AddItem reportingofficer(1)
        approvingofficer.AddItem reportingofficer(1)
        followupofficer.AddItem reportingofficer(1)
        rs("type") = "D"
        rs.Update
    End If
End If
If approvingofficer > "" Then
    Set rs = db.OpenRecordset("select profname, type from professionals where profname =" + Chr$(34) + approvingofficer + Chr$(34))
    If rs.EOF Then
        rs.AddNew
        rs("profname") = approvingofficer
        reportingofficer(0).AddItem approvingofficer
        reportingofficer(1).AddItem approvingofficer
        approvingofficer.AddItem approvingofficer
        followupofficer.AddItem approvingofficer
        rs("type") = "D"
        rs.Update
    End If
End If

'--- PEOPLE
For t% = 0 To 1
    If vsname(t%) = "UNKNOWN" Or vsname(t%) = "" Then
        GoTo nextt
    End If
    Set rs = db.OpenRecordset("select * from people where dpnamelf =" + Chr$(34) + vsname(t%) + Chr$(34))
    If rs.EOF Then
        rs.AddNew
        For tt% = 0 To 1
            vsname(tt%).AddItem vsname(t%)
        Next tt%
    Else
        rs.MoveFirst
        rs.Edit
    End If
    rs("dpnamelf") = vsname(t%)
    rs("dphaddress") = address(t%)
    rs("dphaddress2") = city(t%) + ", " + state(t%) + " " + zipcode(t%)
    rs("dpsort") = Left$(vsname(t%), 15)
    If t% < 2 Then
        If HOMEDAYPHONE(t%) > "" Then
            rs("dphphone") = HOMEDAYPHONE(t%)
        Else
        If HOMENIGHTPHONE(t%) > "" Then
            rs("dphphone") = HOMENIGHTPHONE(t%)
        End If
        End If
        If WORKDAYPHONE(t%) > "" Then
            rs("dpwphone") = HOMEDAYPHONE(t%)
        Else
        If WORKNIGHTPHONE(t%) > "" Then
            rs("dpwphone") = HOMENIGHTPHONE(t%)
        End If
        End If
        rs("HEIGHT") = ht(t%)
        rs("WEIGHT") = weight(t%)
        rs("HAIR") = hair(t%)
        rs("EYES") = eyes(t%)
        rs("PECULIARITIES") = peculiarities(t%)
        rs("resident") = Left$(resident(t%).List(resident(t%).ListIndex), 1)
        rs("race") = Left$(race(t%).List(race(t%).ListIndex), 1)
        rs("sex") = Left$(sex(t%).List(sex(t%).ListIndex), 1)
        rs("age") = age(t%)
        rs("ethnicity") = Left$(ethnicity(t%).List(ethnicity(t%).ListIndex), 1)
    End If
    hoLdname = vsname(t%)
    osort1$ = ""
    If Left$(hoLdname, 1) = " " Then
        hoLdname = Mid$(hoLdname, 2)
        osort1$ = hoLdname
    End If
    If InStr(hoLdname, " CORP") > 0 Or InStr(hoLdname, ",INC") > 0 Or InStr(hoLdname, "COMPANY") > 0 Or InStr(hoLdname, "INC.") > 0 Then
        osort1$ = hoLdname
    End If
    tso$ = hoLdname
    If InStr(tso$, " et al") > 0 Then
        tso$ = Left$(tso$, InStr(tso$, " et al") - 1)
    End If
    If InStr(tso$, " et. al.") > 0 Then
        tso$ = Left$(tso$, InStr(tso$, " et. al.") - 1)
    End If
    If InStr(tso$, ",et al") > 0 Then
        tso$ = Left$(tso$, InStr(tso$, ",et al") - 1)
    End If
    If InStr(tso$, ",et. al.") > 0 Then
        tso$ = Left$(tso$, InStr(tso$, ",et. al.") - 1)
    End If
    If Right$(tso$, 1) = "," Then
        tso$ = Left$(tso$, Len(tso$) - 1)
    End If
    If InStr(tso$, "&") > 0 Then
        tso$ = Left$(tso$, InStr(tso$, "&") - 1)
    End If
    If Right$(tso$, 1) = "," Then
        tso$ = Left$(tso$, Len(tso$) - 1)
    End If
    firstspace% = 0
    While Right$(tso$, 1) = " " And Len(tso$) > 1
        tso$ = Left$(tso$, Len(tso$) - 1)
    Wend
    For tt% = 1 To Len(tso$)
        If Mid$(tso$, tt%, 1) = " " Then
            firstspace% = tt%
            tt% = Len(tso$)
        End If
    Next tt%
    If firstspace% = 0 Then
        If osort1$ = "" Then
            osort1$ = tso$
        End If
        GoTo rsupdate
    End If
    tempsort$ = Mid$(tso$, firstspace% + 1)
    tso$ = Left$(tso$, firstspace% - 1)
    If Right$(tso$, 1) = "," Then
        tso$ = Left$(tso$, Len(tso$) - 1)
    End If
    tempsort$ = tempsort$ + " " + tso$
    If osort1$ = "" Then
        osort1$ = tempsort$
    End If
rsupdate:
    rs("dpname") = osort1$
    rs.Update
nextt:
Next t%


db.Close
On Error Resume Next
Exit Sub
oderror1:
If Err > 3200 Then
    Resume od1
Else
    Resume Next
End If
oderror2:
If Err > 3200 Then
    Resume od2
Else
    Resume Next
End If
End Sub
Private Sub sendcolon()
SendKeys ":"
End Sub
Private Sub drugmeasurement_Click(Index As Integer)

Select Case Index
    Case 3, 4, 5
        
    Case 9 To 29
        
End Select
    
End Sub
Private Sub drugamt_Change(Index As Integer)

Select Case Index
    Case 3, 4, 5
        
    Case 9 To 29
        
End Select

End Sub
Private Sub XGROUP_Click(Index As Integer)
group(Index).Top = description(Index).Top - 1000
group(Index).Left = description(Index).Left
group(Index).Visible = True
group(Index).SetFocus
End Sub

Private Sub ethnicity_Click(Index As Integer)

If Val(victim(Index)) > 0 Then
    
End If
End Sub
Private Sub resident_Click(Index As Integer)

If Val(victim(Index)) > 0 Then
    
End If
End Sub
Private Sub ORGINAL_Click()

End Sub
Private Sub OFFENDERO_Click()

End Sub
Private Sub FILLDATA(IDX As Integer)
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
Set rs = db.OpenRecordset("SELECT * FROM PEOPLE WHERE DPnameLF = " + Chr$(34) + vsname(IDX) + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    If Not IsNull(rs("DPHADDRESS")) Then
        address(IDX) = rs("DPHADDRESS")
    Else
        address(IDX) = ""
    End If
    If Not IsNull(rs("DPHADDRESS2")) Then
        If InStr(rs("DPHADDRESS2"), ",") Then
            city(IDX) = Left$(rs("DPHADDRESS2"), InStr(rs("DPHADDRESS2"), ",") - 1)
            A$ = Mid$(rs("DPHADDRESS2"), InStr(rs("DPHADDRESS2"), ",") + 1)
            For t% = 1 To Len(A$)
                If Mid$(A$, t%, 1) <> " " Then
                    A$ = Mid$(A$, t%)
                    t% = t% + Len(A$)
                End If
            Next t%
            state(IDX) = Left$(A$, 2)
            A$ = Mid$(A$, 3)
            For t% = Len(A$) To 1 Step -1
                If Mid$(A$, t%, 1) = " " Then
                    zipcode(IDX) = Mid$(A$, t% + 1)
                    t% = 1
                End If
            Next t%
        End If
    Else
        city(IDX) = ""
        state(IDX) = ""
    End If
    t% = IDX
    If t% < 2 Then
        HOMEDAYPHONE(t%) = ""
        HOMENIGHTPHONE(t%) = ""
        HOMEDAYPHONE(t%) = ""
        HOMENIGHTPHONE(t%) = ""
        ht(t%) = ""
        weight(t%) = ""
        hair(t%) = ""
        eyes(t%) = ""
        peculiarities(t%) = ""
        If Not IsNull(rs("dphphone")) Then
            HOMEDAYPHONE(t%) = rs("dphphone")
        Else
        If Not IsNull(rs("dphphone")) Then
            HOMENIGHTPHONE(t%) = rs("dphphone")
        End If
        End If
        If Not IsNull(rs("dpwphone")) Then
            HOMEDAYPHONE(t%) = rs("dpwphone")
        Else
        If Not IsNull(rs("dpwphone")) Then
            HOMENIGHTPHONE(t%) = rs("dpwphone")
        End If
        End If
        If Not IsNull(rs("HEIGHT")) Then
            ht(t%) = rs("HEIGHT")
        End If
        If Not IsNull(rs("WEIGHT")) Then
            weight(t%) = rs("WEIGHT")
        End If
        If Not IsNull(rs("HAIR")) Then
            hair(t%) = rs("HAIR")
        End If
        If Not IsNull(rs("EYES")) Then
            eyes(t%) = rs("EYES")
        End If
        If Not IsNull(rs("PECULIARITIES")) Then
            peculiarities(t%) = rs("PECULIARITIES")
        End If
        If Not IsNull(rs("resident")) Then
            For tt% = 0 To resident(t%).ListCount - 1
                If Left$(resident(t%).List(tt%), 1) = rs("resident") Then
                    resident(t%).ListIndex = tt%
                    tt% = resident(t%).ListCount - 1
                End If
            Next tt%
        Else
            resident(t%).ListIndex = -1
        End If
    End If
    If Not IsNull(rs("RACE")) Then
        For tt% = 0 To race(t%).ListCount - 1
            If Left$(race(t%).List(tt%), 1) = rs("RACE") Then
                race(t%).ListIndex = tt%
                tt% = race(t%).ListCount - 1
            End If
        Next tt%
    Else
        race(t%).ListIndex = -1
    End If
    If Not IsNull(rs("SEX")) Then
        For tt% = 0 To sex(t%).ListCount - 1
            If Left$(sex(t%).List(tt%), 1) = rs("SEX") Then
                sex(t%).ListIndex = tt%
                tt% = sex(t%).ListCount - 1
            End If
        Next tt%
    Else
        sex(t%).ListIndex = -1
    End If
    If Not IsNull(rs("ETHNICITY")) Then
        For tt% = 0 To ethnicity(t%).ListCount - 1
            If Left$(ethnicity(t%).List(tt%), 1) = rs("ETHNICITY") Then
                ethnicity(t%).ListIndex = tt%
                tt% = ethnicity(t%).ListCount - 1
            End If
        Next tt%
    Else
        ethnicity(t%).ListIndex = -1
    End If
    If Not IsNull(rs("age")) Then
        age(t%) = rs("age")
    Else
        age(t%) = ""
    End If
End If
db.Close
On Error Resume Next
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub


Private Sub figure()
For t% = 0 To 41 Step 6
    TEMP = Val(totalvalue(t%))
    For tt% = t% + 1 To t% + 5
        TEMP = TEMP + Val(totalvalue(tt%))
    Next tt%
    Select Case t%
        Case 0 To 5
            TVSTOLEN = Format$(TEMP, "########0.00")
        Case 6 To 11
            TVDAMAGED = Format$(TEMP, "########0.00")
        Case 12 To 17
            TVBURNED = Format$(TEMP, "########0.00")
        Case 18 To 23
            TVRECOVERED = Format$(TEMP, "########0.00")
        Case 24 To 29
            TVSEIZED = Format$(TEMP, "########0.00")
        Case 30 To 35
            TVCOUNTERFEIT = Format$(TEMP, "########0.00")
        Case 36 To 41
            TVUNKNOWN = Format$(TEMP, "########0.00")
    End Select
Next t%

End Sub


Private Sub reportingofficer_GotFocus(Index As Integer)
If Me.ActiveControl.Top + Me.ActiveControl.Height > Picture1.Height - Picture2.Top Then
    VScroll1 = Me.ActiveControl.Top - Picture1.Height + Me.ActiveControl.Height
End If

End Sub

Private Sub reportingofficerdate_GotFocus(Index As Integer)
If reportingofficer(Index).ListIndex > -1 And REPORTINGOFFICERDATE(Index) = "" Then
    REPORTINGOFFICERDATE(Index) = Format$(Date$, "mm/dd/yyyy")
End If
End Sub

Private Sub REPORTINGOFFICERDATE_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(REPORTINGOFFICERDATE(Index)) = 1 Or Len(REPORTINGOFFICERDATE(Index)) = 4 Then
    Call sendslash
End If
End If
End Sub

Private Sub REPORTINGOFFICERDATE_LostFocus(Index As Integer)
If REPORTINGOFFICERDATE(Index) > "" And Not IsDate(REPORTINGOFFICERDATE(Index)) Then
    msg = MsgBox("Date/Time entered is invalid.", 48, "Genesis Error Log")
    If fromfind = 0 Then
        REPORTINGOFFICERDATE(Index).SetFocus
    End If
End If
REPORTINGOFFICERDATE(Index) = Format$(REPORTINGOFFICERDATE(Index), "mm/dd/yyyy")
End Sub



Private Sub spellcheck(DONE As Boolean)
Dim wd As New Word.Application
Dim wdsp As Word.SpellingSuggestions
On Error GoTo cmdCheckErr
getout% = 0
lstframe.Visible = False
While tempword > "" And getout% = 0
    While Left$(tempword, 1) = " "
        tempword = Mid$(tempword, 2)
    Wend
    stopper% = Len(tempword) + 1
    For t% = 1 To Len(tempword)
        If InStr("!()[{]};:,./? " + Chr$(34), Mid$(tempword, t%, 1)) Then
            stopper% = t%
            t% = Len(tempword)
        End If
    Next t%
    checkword = LCase(Left$(tempword, stopper% - 1))
    
    
    If stopper% < Len(tempword) Then
        tempword = Mid$(tempword, stopper% + 1)
    Else
        tempword = ""
    End If
    
    
    
    
    If checkword > "" Then
        wd.Documents.Add
        Set wdsp = wd.GetSpellingSuggestions(checkword)
        
        If wdsp.Count > 0 Or wdsp.SpellingErrorType = wdSpellingNotInDictionary Then
            lstsuggestions.Clear
            lstframe.Visible = True
            'RLB code
            lstframe.Top = NARRATIVE.Top - lstframe.Height
            lstframe.Left = NARRATIVE.Left + CLng(lstframe.Width * 0.5)
            '***********
            getout% = 1
        Else
            lstframe.Visible = False
        End If
        
        For i% = 1 To wdsp.Count
            lstsuggestions.AddItem wdsp(i%).Name
        Next i%
        If wdsp.SpellingErrorType = wdSpellingNotInDictionary And wdsp.Count = 0 Then
            lstsuggestions.AddItem "Not found in dictionary."
        End If
        wd.Documents.Close
    Else
        
    End If
Wend
wd.Quit
Set wd = Nothing
Screen.MousePointer = 0
If tempword = "" And lstframe.Visible = False Then
    DONE = True
Else
    DONE = False
End If
Exit Sub
cmdCheckErr:
'MsgBox Err.description
Resume Next
End Sub
Private Sub Command15_Click()
Dim DONE As Boolean
tempword = NARRATIVE.Text
Call spellcheck(DONE)
If DONE Then
    msg = MsgBox("Spelling check complete.", 48, "Genesis Information Log")
End If
End Sub

Private Sub Command16_Click()
lstframe.Visible = False
End Sub

Private Sub Command17_Click()
Dim DONE As Boolean
If lstsuggestions.ListIndex = -1 Then
    Exit Sub
End If
For t% = 1 To Len(NARRATIVE.Text)
    If UCase(Mid$(NARRATIVE.Text, t%, Len(checkword))) = UCase(checkword) Then
        If UCase(Mid$(NARRATIVE.Text, t%, Len(checkword))) = Mid$(NARRATIVE.Text, t%, Len(checkword)) Then
            lstsuggestions.List(lstsuggestions.ListIndex) = UCase(lstsuggestions.List(lstsuggestions.ListIndex))
        End If
        NARRATIVE.Text = Left$(NARRATIVE.Text, t% - 1) + lstsuggestions.List(lstsuggestions.ListIndex) + Mid$(NARRATIVE.Text, t% + Len(checkword))
        t% = t% + Len(checkword)
    End If
Next t%

'RLB CODE
For t% = 1 To Len(tempword)
    If UCase(Mid$(tempword, t%, Len(checkword))) = UCase(checkword) Then
        If UCase(Mid$(tempword, t%, Len(checkword))) = Mid$(tempword, t%, Len(checkword)) Then
            lstsuggestions.List(lstsuggestions.ListIndex) = UCase(lstsuggestions.List(lstsuggestions.ListIndex))
        End If
        tempword = Left$(tempword, t% - 1) + lstsuggestions.List(lstsuggestions.ListIndex) + Mid$(tempword, t% + Len(checkword))
        t% = t% + Len(checkword)
    End If
Next t%
'*********


Call spellcheck(DONE)
If DONE Then
    msg = MsgBox("Spelling check complete.", 48, "Genesis Information Log")
End If
End Sub

Private Sub Command18_Click()
Dim DONE As Boolean

Call spellcheck(DONE)
If DONE Then
    msg = MsgBox("Spelling check complete.", 48, "Genesis Information Log")
End If
End Sub

