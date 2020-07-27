VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form iexport 
   Caption         =   "Export "
   ClientHeight    =   2595
   ClientLeft      =   1470
   ClientTop       =   1710
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2595
   ScaleWidth      =   6765
   Begin MSComctlLib.ProgressBar pb 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label statusl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   285
      TabIndex        =   2
      Top             =   1680
      Width           =   6255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EXPORT IN PROGRESS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   6135
   End
End
Attribute VB_Name = "iexport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim painted As Integer, typegroup(7) As String, incflag(999999) As Boolean, incnumber(999999) As String, incidx As Integer, INDATE1, INDATE2, BASEDATE, rbasedate, EXPORTDATE, incdate, INCSEG As String, foundproperty, PAPERON, ANYA, BOOKINGFOUND, RECOVEREDFOUND, SC, BOOKINGEXPORTED, INCCHANGED, ONCEWB, ONCEWP, ONCEW, ecc As Integer, founderrors As Integer, PROPCHECK1, PROPCHECK2 As Integer, flag As Boolean
Dim vuc(5) As String, foundinjtype As Boolean, SORDER(99999) As Long
Dim NOTONPAPER As Date, orinumber As String, incidentfound As Boolean
Dim pdt(100, 3), pdm(100, 3), propertydate(100), propertytype(100), pdq(100, 3), propertyvalue(100, 7), numvehs(100), numvehr(100) As Single, PROPERTYDRUGS(100, 7) As Boolean
Dim maxv, vidx, siDX, subjectname(100), victimdob(100) As Date, victimage(100), subjectdob(100) As Date, subjectage(100) As Integer, VICTIMTYPE(100), victimrace(100), subjectrace(100), victimsex(100), SUBJECTSEX(100) As String, victimrel(100, 10) As String, EXPYEAR, EXPMONTH As String
Private Sub Form_GotFocus()
'If pb.Value > 0 Then
'    Unload Me
'End If
End Sub
Private Sub Form_Load()
painted = 0
Me.Top = 1500
Me.Width = 6885
Me.Left = 1500
Me.Height = 3060
typegroup(1) = "stolen"
typegroup(2) = "damaged"
typegroup(3) = "burned"
typegroup(4) = "recovered"
typegroup(5) = "seized"
typegroup(6) = "counterfeit"
typegroup(7) = "unknown"
End Sub
Private Sub export()
Dim db, db2 As Database, rs, rs2, rs3, rs4, rs5 As Recordset, outrec, houtrec(500) As String, oridx As Integer, ecc As Integer, CPV As Long
Set db = OpenDatabase(nwi + "incident.mdb")
Set db2 = OpenDatabase(nwb + "booking.mdb")

oridx = 0
PCT% = 0
VCT% = 0
Open nwi + "export" + frmLogin.orinumber + Format$(INDATE1, "mm") + Format$(INDATE1, "yyyy") For Output As #1
For INC% = 1 To incidx
    oridx = 0
    Set rs = db.OpenRecordset("select * from incidentsupport where incidentnumber = " + Chr$(34) + incnumber(INC%) + Chr$(34))
    If rs.EOF Then
        GoTo WENDOUT
    End If
    'rs.MoveFirst
    rs.Edit
    If rs("oncew") = 1 Then
        wonce = 1
    End If
    ecc = rs("exclearchange")
    exdate = ""
    If IsDate(rs("exportdate")) Then
        exdate = rs("exportdate")
    End If
    INCCHANGED = 0
    If rs("schanged") = 1 Then
        INCCHANGED = 1
    End If
    PAPERON = 0
    If rs("onpaper") = 1 Then
        PAPERON = 1
    End If
    SC = 0
    ONCEWP = 0
    Set rs2 = db.OpenRecordset("select * from incidentreporto where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    If Not rs2.EOF Then
        'rs2.MoveFirst
        If rs2("ONCEW") = 1 Then
            ONCEWP = 1
        End If
        If IsDate(rs2("statuschange")) Then
            If rs2("statuschange") <= CVDate(INDATE2) And rs2("statuschange") >= CVDate(INDATE1) Then
                If rs2("statuschange") > rs("exportdate") Then
                    SC = 1
                End If
            End If
        End If
    End If
    foundpdrugs% = 0
    For t% = 1 To 6
        For tt% = 1 To 3
            If Not IsNull(rs2("PTYPEOFDRUG" + Mid$(Str$(t%), 2) + Mid$(Str$(tt%), 2))) Then
                foundpdrugs% = 1
                tt% = 3
                t% = 6
            End If
        Next tt%
    Next t%
    If foundpdrugs% = 0 Then
        Set rs3 = db.OpenRecordset("select * from SUPPLEMENTALSUPPORT where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
        If Not rs3.EOF Then
            'rs3.MoveFirst
            For t% = 1 To 6
                For tt% = 1 To 3
                    If Not IsNull(rs3("PTYPEOFDRUG" + Mid$(Str$(t%), 2) + Mid$(Str$(tt%), 2))) Then
                        foundpdrugs% = 1
                        tt% = 3
                        t% = 6
                    End If
                Next tt%
            Next t%
        End If
    End If
    RECOVEREDFOUND = 0
    For t% = 1 To 6
        If rs2("recoveredvalue" + Mid$(Str$(t%), 2)) > 0 Then
            Set rs3 = db.OpenRecordset("select * from incidentsupport where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
            If Not rs3.EOF Then
                'rs3.MoveFirst
                If rs3("daterecovered" + Mid$(Str$(t%), 2)) >= CVDate(INDATE1) And rs3("daterecovered" + Mid$(Str$(t%), 2)) <= CVDate(INDATE2) Then
                    RECOVEREDFOUND = 1
                    t% = 6
                End If
            End If
        End If
    Next t%
    Set rs3 = db.OpenRecordset("select * from supplemental where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    If Not rs3.EOF Then
        'rs3.MoveFirst
        If rs3("schanged") = 1 Then
            INCCHANGED = 1
        End If
        For t% = 1 To 6
            If rs3("recoveredvalue" + Mid$(Str$(t%), 2)) > 0 Then
                Set rs4 = db.OpenRecordset("select * from incidentsupport where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
                If Not rs4.EOF Then
                    'rs4.MoveFirst
                    If rs4("daterecovered" + Mid$(Str$(t%), 2)) >= CVDate(INDATE1) And rs4("daterecovered" + Mid$(Str$(t%), 2)) <= CVDate(INDATE2) Then
                        RECOVEREDFOUND = 1
                        t% = 6
                    End If
                End If
            End If
        Next t%
    End If
    ANYA = 0
    foundproperty = 0
    only35a = False
    For oo% = 1 To 10
        If Not IsNull(rs("UCR" + Mid$(Str$(oo%), 2))) And rs("UCR" + Mid$(Str$(oo%), 2)) > "" Then
            Select Case rs("UCR" + Mid$(Str$(oo%), 2))
                Case "90A", "90B", "90C", "90D", "90E", "90F", "90G", "90H", "90I", "90J", "90Z", "90K", "90L", "90N", "90P"
                Case Else
                    ANYA = 1
            End Select
            Select Case rs("UCR" + Mid$(Str$(oo%), 2))
                Case "100", "120", "200", "210", "220", "23A", "23B", "23C", "23D", "23E", "23F", "23G", "23H", "240", "250", "26A", "26B", "26C", "26D", "26E", "270", "280", "290", "35A", "35B", "39A", "39B", "39C", "39D", "510"
                    foundproperty = 1
            End Select
            Select Case rs("UCR" + Mid$(Str$(oo%), 2))
                Case "35A"
                    only35a = True
                Case Else
                    only35a = False
            End Select

        End If
    Next oo%
    incdate = ""
    Set rs3 = db.OpenRecordset("select * from incidentreportc where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    If Not rs3.EOF Then
        'rs3.MoveFirst
        incdate = rs3("dateofoffense1")
    End If
    BOOKINGFOUND = 0
    BOOKINGEXPORTED = 0
    ONCEWB = 0
    Set rs4 = db2.OpenRecordset("select * from booking where NUMBER < 100 and dateofarrest between #" + INDATE1 + "# and #" + INDATE2 + "# AND incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34) + " and (exportdate is null or (exportdate is not null and dateofarrest < #" + INDATE2 + "# and lastupdate between #" + INDATE1 + "# and #" + INDATE2 + "#))")
    If Not rs4.EOF Then
        'rs4.MoveFirst
        BOOKINGFOUND = 1
        If IsDate(rs4("exportdate")) Then
            BOOKINGEXPORTED = 1
        End If
        If rs4("ONCEW") = 1 Then
            ONCEWB = 1
        End If
    End If
    If ANYA = 0 And BOOKINGFOUND = 0 Then
        GoTo WENDOUT
    End If
    If Not IsNull(rs("EXPORTDATE")) Then
        EXPORTDATE = rs("EXPORTDATE")
    Else
        EXPORTDATE = ""
    End If
    rc% = 0
    flag = incflag(INC%)
    Call SETINCSEG(rc%)
    If rc% = -1 Then
        GoTo WENDOUT
    End If
    pb.Value = pb.Value + 1
    pb.Refresh
    wonce = 0
    PCT% = 0
    PAPERON = 0
    exdate = ""
    '===== Administrative Segment =====
    If (Mid$(INCSEG, 1, 1) = " " And Mid$(INCSEG, 2, 1) = " ") Or ANYA = 0 Then
        GoTo offense
    End If
    Set rs2 = db.OpenRecordset("select * from incidentreporto where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    If Not rs2.EOF Then
        'rs2.MoveFirst
    Else
        GoTo WENDOUT
    End If
    If Mid$(INCSEG, 2, 1) = "W" Then
        rs("exclearchange") = 0
        '===== Data Element 4, 5
        If Not rs2("offenderdeath") And Not rs2("noprosecution") And Not rs2("extraditiondenied") And Not rs2("victimdeclinescooperation") And Not rs2("juvenilenocustody") And RECOVEREDFOUND = 0 And (IsNull(rs2("STATUSCHANGE")) Or rs2("STATUSCHANGE") > CVDate(INDATE2) Or rs2("STATUSCHANGE") < CVDate(INDATE1)) And BOOKINGFOUND = 0 Then
            rs("tempreason") = "A Time-Window Submission cannot have a value of NA for Exceptional Clearance."
            rs("temp") = "Y"
            rs.Update
            GoTo WENDOUT
        Else
        If Not IsDate(rs2("excleardate")) And RECOVEREDFOUND = 0 And (IsNull(rs2("STATUSCHANGE")) Or rs2("STATUSCHANGE") > CVDate(INDATE2) Or rs2("STATUSCHANGE") < CVDate(INDATE1)) And BOOKINGFOUND = 0 Then
            rs("tempreason") = "A valid exceptional clearance date must be entered."
            rs("temp") = "Y"
            rs.Update
            GoTo WENDOUT
        Else
        If rs2("excleardate") < CVDate(BASEDATE) Then
            rs("tempreason") = "Exceptional Clearance Date on a Time-Window Submission cannot be prior to the system base date."
            rs("temp") = "Y"
            rs.Update
            GoTo WENDOUT
        End If
        End If
        End If
    Else
    If Mid$(INCSEG, 2, 1) = "M" Then
        rs("exclearchange") = 0
        '===== Data Element 4
        If wonce = 1 Then
            If Not rs2("offenderdeath") And Not rs2("noprosecution") And Not rs2("extraditiondenied") And Not rs2("victimdeclinescooperation") And Not rs2("juvenilenocustody") Then
                rs("tempreason") = "A Modify Submission which had a previous Time Windows submission cannot have a value of NA for Exceptional Clearance."
                rs("temp") = "Y"
                rs.Update
                GoTo WENDOUT
            End If
            If IsDate(rs2("excleardate")) Then
                If rs2("excleardate") < CVDate(BASEDATE) Then
                    rs("tempreason") = "Exceptional Clearance Date on a Time-Window Submission cannot be prior to the system base date."
                    rs("temp") = "Y"
                    rs.Update
                    GoTo WENDOUT
                End If
            Else
                rs("tempreason") = "An Exceptional Clearance Date must be entered for Modify submission."
                rs("temp") = "Y"
                rs.Update
                GoTo WENDOUT
            End If
        Else
        If rs2("offenderdeath") Or rs2("noprosecution") Or rs2("extraditiondenied") Or rs2("victimdeclinescooperation") Or rs2("juvenilenocustody") Then
            rs("tempreason") = "A Modify Submission must have a value of NA for Exceptional Clearance."
            rs("temp") = "Y"
            rs.Update
            GoTo WENDOUT
        End If
        End If
    End If
    End If
    For a% = 1 To 2
        If a% = 1 And Mid$(INCSEG, 1, 1) <> "D" Then
            GoTo NEXTA
        End If
        outrec = "    " + "1"
        If a% = 1 Then
            outrec = outrec + "D"
        Else
            outrec = outrec + Mid$(INCSEG, 2, 1)
        End If
        ' Incident Date
        ' Report Date Indicator
        ' Incident Hour
        outrec = outrec + EXPMONTH + EXPYEAR + Space(4) + orinumber + Format$(rs("incidentnumber"), "!@@@@@@@@@@@@") + Format$(rs3("dateofoffense1"), "yyyymmdd") + " " + Format$(rs3("timeofoffense1"), "hh")
        ' Cleared Exceptionally
        If rs2("offenderdeath") Then
            outrec = outrec + "A"
        Else
        If rs2("noprosecution") Then
            outrec = outrec + "B"
        Else
        If rs2("extraditiondenied") Then
            outrec = outrec + "C"
        Else
        If rs2("victimdeclinescooperation") Then
            outrec = outrec + "D"
        Else
        If rs2("juvenilenocustody") Then
            outrec = outrec + "E"
        Else
            outrec = outrec + "N"
        End If
        End If
        End If
        End If
        End If
        ' Exceptional Clearance Date
        If Not IsNull(rs2("excleardate")) Then
            outrec = outrec + Format$(rs2("excleardate"), "yyyymmdd")
        Else
            outrec = outrec + Space$(8)
        End If
        ' UCR Offense Code
        If Mid$(outrec, 6, 1) = "W" Or (Mid$(outrec, 6, 1) = "M" And wonce) Then
            For tt% = 1 To 5
                If Not IsNull(rs("ucr" + Mid$(Str$(tt%), 2))) And rs("ucr" + Mid$(Str$(tt%), 2)) > "" Then
                    outrec = outrec + rs("ucr" + Mid$(Str$(tt%), 2))
                Else
                    outrec = outrec + "   "
                End If
            Next tt%
        Else
            outrec = outrec + Space$(15)
        End If
        outrec = outrec + Space$(15)
        '===== SC Enhancements
        ' Tract
        outrec = outrec + "     "
        ' Status Ind
        If rs2("active") = "X" Then
            outrec = outrec + "A"
        Else
        If rs2("admclosed") = "X" Then
            outrec = outrec + "C"
        Else
            outrec = outrec + "U"
        End If
        End If
        ' State Change Date
        If Not IsNull(rs2("statuschange")) Then
            outrec = outrec + Format$(rs2("statuschange"), "yyyymmdd")
        Else
            outrec = outrec + "        "
        End If
        ' End Date
        outrec = outrec + Format$(rs3("dateofoffense2"), "yyyymmdd")
        ' End Time
        outrec = outrec + Format$(rs3("timeofoffense2"), "hh")
        ' Additional COunters are added as Determined
        If a% = 2 Then
            oridx = oridx + 1
        End If
        houtrec(oridx) = outrec
NEXTA:
    Next a%
    
    If Mid$(outrec, 6, 1) = "W" Then
        rs("oncew") = 1
        rs.Update
    End If
    '===== Offense Segment =====
offense:
    If Mid$(INCSEG, 3, 1) = " " Or ANYA = 0 Then
        ofct% = 0
        GoTo property
    End If
    Set rs2 = db.OpenRecordset("select * from incidentreportc where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    If Not rs2.EOF Then
        'rs2.MoveFirst
    Else
        GoTo WENDOUT
    End If
    ofct% = 0
    pr1$ = ""
    pr2$ = ""
    fe$ = ""
    cp$ = ""
    ue$ = ""
    For u% = 1 To 5
        If Not IsNull(rs("ucr" + Mid$(Str$(u%), 2))) Then
            Select Case rs("UCR" + Mid$(Str$(u%), 2))
            Case "90A", "90B", "90C", "90D", "90E", "90F", "90G", "90H", "90I", "90J", "90Z", "90K", "90L", "90N", "90P"
            Case Else
            uc$ = rs("ucr" + Mid$(Str$(u%), 2))
            If u% < 4 Then
                '===== Error 254
                '===== SCEdit 7.2 Amendments to NIBRS Volume 4 - 23F, 23G, 240 added to 220
                If uc$ = "220" Or uc$ = "23F" Or uc$ = "23G" Or uc$ = "240" Then
                    If rs2("forcedentryyes" + Mid$(Str$(u%), 2)) = 1 Then
                        fe$ = "F"
                    Else
                        fe$ = "N"
                    End If
                Else
                    fe$ = " "
                End If
                If rs2("completedyes" + Mid$(Str$(u%), 2)) Then
                    cp$ = "C"
                Else
                    cp$ = "A"
                End If
                ue$ = rs2("entered" + Mid$(Str$(u%), 2))
                pr1$ = rs2("premise" + Mid$(Str$(u%), 2))
                If Not IsNull(rs2("premise" + Mid$(Str$(u%), 2) + "a")) Then
                    pr2$ = rs2("premise" + Mid$(Str$(u%), 2) + "a")
                End If
            Else
                '===== Error 254
                '===== SCEdit 7.2 Amendments to NIBRS Volume 4 - 23F, 23G, 240 added to 220
                If uc$ = "220" Or uc$ = "23F" Or uc$ = "23G" Or uc$ = "240" Then
                    If rs("forcedentryyes" + Mid$(Str$(u%), 2)) = 1 Then
                        fe$ = "F"
                    Else
                        fe$ = "N"
                    End If
                Else
                    fe$ = " "
                End If
                If rs("completedyes" + Mid$(Str$(u%), 2)) Then
                    cp$ = "C"
                Else
                    cp$ = "A"
                End If
                pr1$ = rs("premise" + Mid$(Str$(u%), 2))
                If Not IsNull(rs("premise" + Mid$(Str$(u%), 2) + "a")) Then
                    pr2$ = rs("premise" + Mid$(Str$(u%), 2) + "a")
                End If
                ue$ = rs("entered" + Mid$(Str$(u%), 2))
            End If
            outrec = "    " + "2" + Mid$(INCSEG, 3, 1) + EXPMONTH + EXPYEAR + Space(4) + orinumber + Format$(rs2("incidentnumber"), "!@@@@@@@@@@@@")
            ' UCR Offense Code
            outrec = outrec + uc$
            ' Offense Attempted/Completed
            outrec = outrec + cp$
            ' Offender(s) Suspected of Using
            '===== Error 201, 204, 206,207
            TEMPUSING = "N  "
            Set rs3 = db.OpenRecordset("select sdrugsyes from incidentreports where sdrugsyes = 'X' and INCIDENTNUMBER = " + Chr$(34) + rs("Incidentnumber") + Chr$(34))
            Set rs4 = db.OpenRecordset("select sdrugsyes1, sdrugsyes2 from supplemental where (sdrugsyes1 = 'X' or sdrugsyes2 = 'X') and INCIDENTNUMBER = " + Chr$(34) + rs("Incidentnumber") + Chr$(34))
            If Not rs3.EOF Or Not rs4.EOF Then
                TEMPUSING = "D  "
            End If
            Set rs3 = db.OpenRecordset("select salcoholyes from incidentreports where salcoholyes = 'X' AND INCIDENTNUMBER = " + Chr$(34) + rs("Incidentnumber") + Chr$(34))
            Set rs4 = db.OpenRecordset("select salcoholyes1, salcoholyes2 from supplemental where (salcoholyes1 = 'X' or salcoholyes2 = 'X') and INCIDENTNUMBER = " + Chr$(34) + rs("Incidentnumber") + Chr$(34))
            If Not rs3.EOF Or Not rs4.EOF Then
                If Left$(TEMPUSING, 1) = "N" Then
                    Mid$(TEMPUSING, 1, 1) = "A"
                Else
                    Mid$(TEMPUSING, 2, 1) = "A"
                End If
            End If
            Set rs3 = db.OpenRecordset("select computerequipment from incidentreports where computerequipment = 1 AND INCIDENTNUMBER = " + Chr$(34) + rs("Incidentnumber") + Chr$(34))
            Set rs4 = db.OpenRecordset("select computerequipment1, computerequipment2 from supplemental where computerequipment1 = 1 or computerequipment2 = 1 and INCIDENTNUMBER = " + Chr$(34) + rs("Incidentnumber") + Chr$(34) + " AND (SUBJECT1 IS NOT NULL OR SUBJECT2 IS NOT NULL)")
            If Not rs3.EOF Then
                If Left$(TEMPUSING, 1) = "N" Then
                    Mid$(TEMPUSING, 1, 1) = "C"
                Else
                If Mid$(TEMPUSING, 2, 1) = "N" Then
                    Mid$(TEMPUSING, 2, 1) = "C"
                Else
                    Mid$(TEMPUSING, 3, 1) = "C"
                End If
                End If
            End If
            If InStr(TEMPUSING, "C") = 0 Then
                If Not rs4.EOF Then
                    'rs4.MoveFirst
                    While Not rs4.EOF
                        If Not IsNull(rs4("SUBJECT1")) And rs4("COMPUTEREQUIPMENT1") = 1 Then
                            If Left$(TEMPUSING, 1) = "N" Then
                                Mid$(TEMPUSING, 1, 1) = "C"
                            Else
                            If Mid$(TEMPUSING, 2, 1) = "N" Then
                                Mid$(TEMPUSING, 2, 1) = "C"
                            Else
                                Mid$(TEMPUSING, 3, 1) = "C"
                            End If
                            End If
                        End If
                        If Not IsNull(rs4("SUBJECT2")) And rs4("COMPUTEREQUIPMENT2") = 1 Then
                            If Left$(TEMPUSING, 1) = "N" Then
                                Mid$(TEMPUSING, 1, 1) = "C"
                            Else
                            If Mid$(TEMPUSING, 2, 1) = "N" Then
                                Mid$(TEMPUSING, 2, 1) = "C"
                            Else
                                Mid$(TEMPUSING, 3, 1) = "C"
                            End If
                            End If
                        End If
                        If InStr(TEMPUSING, "C") = 0 Then
                            rs4.MoveNext
                        Else
                            rs4.MoveLast
                        End If
                    Wend
                End If
            End If
            outrec = outrec + TEMPUSING
            ' Location Type
            outrec = outrec + pr1$
            ' Number of Premises Entered
            If uc$ = "220" And (pr$ = "14" Or pr$ = "19") And Val(ue$) > 0 Then
                outrec = outrec + Format$(Val(ue$), "00")
            Else
                outrec = outrec + Space$(2)
            End If
            ' Method of Entry
            outrec = outrec + fe$
            ' Type Criminal Activity
            If Not IsNull(rs("activity" + Mid$(Str$(u%), 2) + "1")) Then
                outrec = outrec + rs("activity" + Mid$(Str$(u%) + "1", 2))
            Else
                outrec = outrec + " "
            End If

            If Not IsNull(rs("activity" + Mid$(Str$(u%), 2) + "2")) Then
                outrec = outrec + rs("activity" + Mid$(Str$(u%) + "2", 2))
            Else
                outrec = outrec + " "
            End If
            If Not IsNull(rs("activity" + Mid$(Str$(u%), 2) + "3")) Then
                outrec = outrec + rs("activity" + Mid$(Str$(u%) + "3", 2))
            Else
                outrec = outrec + " "
            End If
            ' Type Weapon/Force Involved and Automatic Weapon Indictor
            tw$ = ""
            For yy% = 1 To 3
                If Not IsNull(rs2("weapontype" + Mid$(Str$(yy%), 2))) And rs2("weapontype" + Mid$(Str$(yy%), 2)) > "" Then
                    Select Case uc$
                        Case "09A", "09B", "09C", "100", "11A", "11B", "11C", "11D", "120", "13A", "13B", "210", "520"
                            tw$ = tw$ + rs2("weapontype" + Mid$(Str$(yy%), 2))
                            If rs2("automatic" + Mid$(Str$(yy%), 2)) <> "N" Then
                                tw$ = tw$ + rs2("automatic" + Mid$(Str$(yy%), 2))
                            Else
                                tw$ = tw$ + " "
                            End If
                        Case Else
                            tw$ = tw$ + "   "
                    End Select
                Else
                    tw$ = tw$ + "   "
                End If
            Next yy%
            outrec = outrec + tw$
            ' Victim(s) Suspected of Using
            TEMPUSING = "N  "
            Set rs3 = db.OpenRecordset("select vdrugsyes from incidentreportv where vdrugsyes = 'X' and INCIDENTNUMBER = " + Chr$(34) + rs("Incidentnumber") + Chr$(34))
            Set rs4 = db.OpenRecordset("select vdrugsyes1, vdrugsyes2 from supplemental where (vdrugsyes1 = 'X' or vdrugsyes2 = 'X') and INCIDENTNUMBER = " + Chr$(34) + rs("Incidentnumber") + Chr$(34))
            If Not rs3.EOF Or Not rs4.EOF Then
                TEMPUSING = "D  "
            End If
            Set rs3 = db.OpenRecordset("select valcoholyes from incidentreportv where valcoholyes = 'X' AND INCIDENTNUMBER = " + Chr$(34) + rs("Incidentnumber") + Chr$(34))
            Set rs4 = db.OpenRecordset("select valcoholyes1, valcoholyes2 from supplemental where (valcoholyes1 = 'X' or valcoholyes2 = 'X') and INCIDENTNUMBER = " + Chr$(34) + rs("Incidentnumber") + Chr$(34))
            If Not rs3.EOF Or Not rs4.EOF Then
                If Left$(TEMPUSING, 1) = "N" Then
                    Mid$(TEMPUSING, 1, 1) = "A"
                Else
                    Mid$(TEMPUSING, 2, 1) = "A"
                End If
            End If
            Set rs3 = db.OpenRecordset("select computerequipment from incidentreportV where computerequipment = 1 AND INCIDENTNUMBER = " + Chr$(34) + rs("Incidentnumber") + Chr$(34))
            Set rs4 = db.OpenRecordset("select computerequipment1, computerequipment2 from supplemental where computerequipment1 = 1 or computerequipment2 = 1 and INCIDENTNUMBER = " + Chr$(34) + rs("Incidentnumber") + Chr$(34) + " AND (VICTIM1 IS NOT NULL OR VICTIM2 IS NOT NULL)")
            If Not rs3.EOF Then
                If Left$(TEMPUSING, 1) = "N" Then
                    Mid$(TEMPUSING, 1, 1) = "C"
                Else
                If Mid$(TEMPUSING, 2, 1) = "N" Then
                    Mid$(TEMPUSING, 2, 1) = "C"
                Else
                    Mid$(TEMPUSING, 3, 1) = "C"
                End If
                End If
            End If
            If InStr(TEMPUSING, "C") = 0 Then
                If Not rs4.EOF Then
                    'rs4.MoveFirst
                    While Not rs4.EOF
                        If Not IsNull(rs4("VICTIM1")) And rs4("COMPUTEREQUIPMENT1") = 1 Then
                            If Left$(TEMPUSING, 1) = "N" Then
                                Mid$(TEMPUSING, 1, 1) = "C"
                            Else
                            If Mid$(TEMPUSING, 2, 1) = "N" Then
                                Mid$(TEMPUSING, 2, 1) = "C"
                            Else
                                Mid$(TEMPUSING, 3, 1) = "C"
                            End If
                            End If
                        End If
                        If Not IsNull(rs4("VICTIM2")) And rs4("COMPUTEREQUIPMENT2") = 1 Then
                            If Left$(TEMPUSING, 1) = "N" Then
                                Mid$(TEMPUSING, 1, 1) = "C"
                            Else
                            If Mid$(TEMPUSING, 2, 1) = "N" Then
                                Mid$(TEMPUSING, 2, 1) = "C"
                            Else
                                Mid$(TEMPUSING, 3, 1) = "C"
                            End If
                            End If
                        End If
                        If InStr(TEMPUSING, "C") = 0 Then
                            rs4.MoveNext
                        Else
                            rs4.MoveLast
                        End If
                    Wend
                End If
            End If
            outrec = outrec + TEMPUSING
            ' Second Premise
            outrec = outrec + pr2$ + Space$(2 - Len(pr2$))
            ' Circum
            If Not IsNull(rs("subcodes" + Mid$(Str$(u%), 2) + "1")) Then
                outrec = outrec + rs("subcodes" + Mid$(Str$(u%), 2) + "1")
            Else
                outrec = outrec + " "
            End If
            If Not IsNull(rs("subcodes" + Mid$(Str$(u%), 2) + "2")) Then
                outrec = outrec + rs("subcodes" + Mid$(Str$(u%), 2) + "2")
            Else
                outrec = outrec + " "
            End If
            If Not IsNull(rs("subcodes" + Mid$(Str$(u%), 2) + "3")) Then
                outrec = outrec + rs("subcodes" + Mid$(Str$(u%), 2) + "3")
            Else
                outrec = outrec + " "
            End If
            ' Incident Date
            outrec = outrec + Mid$(houtrec(1), 38, 8)
            ' Bias
            Set rs4 = db.OpenRecordset("select bias from incidentreporto where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
            If Not rs4.EOF Then
                'rs4.MoveFirst
                bs$ = rs4("bias")
            End If
            outrec = outrec + bs$
            ofct% = ofct% + 1
            oridx = oridx + 1
            houtrec(oridx) = outrec
            End Select
        End If
LOOPU:
    Next u%
'===== Property Segment =====
property:
    '===== Administrative Segment Counter for Offense
    If Mid$(INCSEG, 2, 1) <> " " Then
        houtrec(1) = houtrec(1) + Format$(ofct%, "00")
    End If
    If Mid$(INCSEG, 4, 1) = " " Or ANYA = 0 Then
        PCT% = 0
        GoTo victim
    End If
    PCT% = 0
    If foundproperty = 0 And foundpdrugs% = 0 And RECOVEREDFOUND = 0 Then
        GoTo victim
    End If
    allucr = ""
    pd% = 0
    Set rs2 = db.OpenRecordset("select * from incidentreporto where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    pidx% = 0
    TGIDX% = 0
    Call CLEARPROPERTY
    If Not rs2.EOF Then
        'rs2.MoveFirst
        While Not rs2.EOF
            'see saveold
            Set rs3 = db.OpenRecordset("select * from incidentsupport where incidentnumber = " + Chr$(34) + rs2("incidentnumber") + Chr$(34))
            'rs3.MoveFirst
            For eachone% = 1 To 6
                FOUNDVALUE = False
                If Not IsNull(rs3("group" + CStr(eachone%))) And Not IsNull(rs3("pucr" + CStr(eachone%))) Then
                    ' Add to all ucr list
                    fdup% = 0
                    For ii% = 1 To Len(allucr) Step 3
                        If Mid$(allucr, ii%, 3) = rs3("pucr" + CStr(eachone%)) Then
                            fdup% = 1
                            ii% = Len(allucr)
                        End If
                    Next ii%
                    If fdup% = 0 Then
                        allucr = allucr + rs3("pucr" + CStr(eachone%))
                    End If
                    FOUNDGROUP% = 0
                    For GROUPTYPE% = PROPCHECK1 To PROPCHECK2
                        ' Get information for X entries
                        If Not IsNull(rs2(typegroup(GROUPTYPE%) + "value" + CStr(eachone%))) And rs2(typegroup(GROUPTYPE%) + "value" + CStr(eachone%)) = 9999999 Then
                            ' Get property type
                            FOUNDGROUP% = 0
                            For lookup% = 1 To TGIDX%
                                If propertytype(lookup%) = rs3("group" + CStr(eachone%)) Then
                                    FOUNDGROUP% = lookup%
                                    lookup% = TGIDX%
                                End If
                            Next lookup%
                            If FOUNDGROUP% = 0 Then
                                TGIDX% = TGIDX% + 1
                                propertytype(TGIDX%) = rs3("group" + CStr(eachone%))
                                FOUNDGROUP% = TGIDX%
                            End If
                            ' Get amount
                            If propertytype(FOUNDGROUP%) = "22" Or propertytype(FOUNDGROUP%) = "99" Or propertytype(FOUNDGROUP%) = "77" Or propertytype(FOUNDGROUP%) = "09" Then
                                propertyvalue(FOUNDGROUP%, GROUPTYPE%) = 9999999
                            Else
                                propertyvalue(FOUNDGROUP%, GROUPTYPE%) = propertyvalue(FOUNDGROUP%, GROUPTYPE%) + 0
                            End If
                            PROPERTYDRUGS(FOUNDGROUP%, GROUPTYPE%) = True
                            FOUNDVALUE = True
                            ' Get drug information for seized
                            If rs3("pucr" + CStr(eachone%)) = "35A" And propertytype(FOUNDGROUP%) = "10" And GROUPTYPE% = 5 Then
                                For drugspicked% = 1 To 3
                                    If Not IsNull(rs2("ptypeofdrug" + CStr(eachone%) + CStr(drugspicked%))) Then
                                        pdt(FOUNDGROUP%, drugspicked%) = rs2("ptypeofdrug" + CStr(eachone%) + CStr(drugspicked%))
                                        pdm(FOUNDGROUP%, drugspicked%) = rs2("pdrugmeasurement" + CStr(eachone%) + CStr(drugspicked%))
                                        pdq(FOUNDGROUP%, drugspicked%) = rs2("pdrugamt" + CStr(eachone%) + CStr(drugspicked%))
                                    End If
                                Next drugspicked%
                            End If
                        Else
                        ' Get information when a value is entered
                        If Not IsNull(rs2(typegroup(GROUPTYPE%) + "value" + CStr(eachone%))) And rs2(typegroup(GROUPTYPE%) + "value" + CStr(eachone%)) > 0 Then
                            ' Get property type
                            FOUNDGROUP% = 0
                            For lookup% = 1 To TGIDX%
                                If propertytype(lookup%) = rs3("group" + CStr(eachone%)) Then
                                    FOUNDGROUP% = lookup%
                                    lookup% = TGIDX%
                                End If
                            Next lookup%
                            If FOUNDGROUP% = 0 Then
                                TGIDX% = TGIDX% + 1
                                propertytype(TGIDX%) = rs3("group" + CStr(eachone%))
                                FOUNDGROUP% = TGIDX%
                            End If
                            ' Get amount
                            propertyvalue(FOUNDGROUP%, GROUPTYPE%) = propertyvalue(FOUNDGROUP%, GROUPTYPE%) + rs2(typegroup(GROUPTYPE%) + "value" + CStr(eachone%))
                            FOUNDVALUE = True
                            If GROUPTYPE% = 4 Then
                                ' Get recovered vehicles
                                If Not IsNull(rs3("numvehicler" + CStr(eachone%))) Then
                                    numvehr(FOUNDGROUP%) = numvehr(FOUNDGROUP%) + rs3("numvehicler" + Mid$(Str$(eachone%), 2))
                                End If
                                ' Get recovered date
                                If IsDate(rs3("daterecovered" + CStr(eachone%))) Then
                                    propertydate(FOUNDGROUP%) = rs3("daterecovered" + CStr(eachone%))
                                End If
                            End If
                            If GROUPTYPE% = 1 Then
                                ' Get stolen vehicles
                                If Not IsNull(rs3("numvehicles" + CStr(eachone%))) Then
                                    numvehs(FOUNDGROUP%) = numvehs(FOUNDGROUP%) + rs3("numvehicles" + Mid$(Str$(eachone%), 2))
                                End If
                            End If
                            ' Get drug information for seized
                            If rs3("pucr" + CStr(eachone%)) = "35A" And propertytype(eachone%) = "10" And GROUPTYPE% = 5 Then
                                For drugspicked% = 1 To 3
                                    If Not IsNull(rs2("ptypeofdrug" + CStr(eachone%) + CStr(drugspicked%))) Then
                                        pdt(FOUNDGROUP%, drugspicked%) = rs2("ptypeofdrug" + CStr(eachone%) + CStr(drugspicked%))
                                        pdm(FOUNDGROUP%, drugspicked%) = rs2("pdrugmeasurement" + CStr(eachone%) + CStr(drugspicked%))
                                        pdq(FOUNDGROUP%, drugspicked%) = rs2("pdrugamt" + CStr(eachone%) + CStr(drugspicked%))
                                    End If
                                Next drugspicked%
                            End If
                        End If
                    End If
                    If FOUNDGROUP% > TGIDX% Then
                        TGIDX% = FOUNDGROUP%
                    End If
                    Next GROUPTYPE%
                End If
            Next eachone%
            rs2.MoveNext
        Wend
    End If
    Set rs2 = db.OpenRecordset("select * from SUPPLEMENTAL where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    If Not rs2.EOF And Not rs3.EOF Then
        'rs2.MoveFirst
        While Not rs2.EOF
            Set rs3 = db.OpenRecordset("select * from supplementalsupport where incidentnumber = " + Chr$(34) + rs2("incidentnumber") + Chr$(34))
            If rs3.EOF Then
                GoTo WENDOUT
            End If
            'rs3.MoveFirst
            For eachone% = 1 To 6
                FOUNDVALUE = False
                If Not IsNull(rs3("group" + CStr(eachone%))) And Not IsNull(rs3("pucr" + CStr(eachone%))) Then
                    ' Add to all ucr list
                    fdup% = 0
                    For ii% = 1 To Len(allucr) Step 3
                        If Mid$(allucr, ii%, 3) = rs3("pucr" + CStr(eachone%)) Then
                            fdup% = 1
                            ii% = Len(allucr)
                        End If
                    Next ii%
                    If fdup% = 0 Then
                        allucr = allucr + rs3("pucr" + CStr(eachone%))
                    End If
                    For GROUPTYPE% = PROPCHECK1 To PROPCHECK2
                        ' Get information for X entries
                        If Not IsNull(rs2(typegroup(GROUPTYPE%) + "value" + CStr(eachone%))) And rs2(typegroup(GROUPTYPE%) + "value" + CStr(eachone%)) = 9999999 Then
                            ' Get property type
                            FOUNDGROUP% = 0
                            For lookup% = 1 To TGIDX%
                                If propertytype(lookup%) = rs3("group" + CStr(eachone%)) Then
                                    FOUNDGROUP% = lookup%
                                    lookup% = TGIDX%
                                End If
                            Next lookup%
                            If FOUNDGROUP% = 0 Then
                                TGIDX% = TGIDX% + 1
                                propertytype(TGIDX%) = rs3("group" + CStr(eachone%))
                                FOUNDGROUP% = TGIDX%
                            End If
                            ' Get amount
                            If propertytype(FOUNDGROUP%) = "22" Or propertytype(FOUNDGROUP%) = "99" Then
                                propertyvalue(FOUNDGROUP%, GROUPTYPE%) = 9999999
                            Else
                                propertyvalue(FOUNDGROUP%, GROUPTYPE%) = propertyvalue(FOUNDGROUP%, GROUPTYPE%) + 0
                            End If
                            FOUNDVALUE = True
                            ' Get drug information for seized
                            If rs3("pucr" + CStr(eachone%)) = "35A" And propertytype(FOUNDGROUP%) = "10" And GROUPTYPE% = 5 Then
                                For drugspicked% = 1 To 3
                                    If Not IsNull(rs2("ptypeofdrug" + CStr(eachone%) + CStr(drugspicked%))) Then
                                        pdt(FOUNDGROUP%, drugspicked%) = rs2("ptypeofdrug" + CStr(eachone%) + CStr(drugspicked%))
                                        pdm(FOUNDGROUP%, drugspicked%) = rs2("pdrugmeasurement" + CStr(eachone%) + CStr(drugspicked%))
                                        pdq(FOUNDGROUP%, drugspicked%) = rs2("pdrugamt" + CStr(eachone%) + CStr(drugspicked%))
                                    End If
                                Next drugspicked%
                            End If
                        Else
                        ' Get information when a value is entered
                        If Not IsNull(rs2(typegroup(GROUPTYPE%) + "value" + CStr(eachone%))) And rs2(typegroup(GROUPTYPE%) + "value" + CStr(eachone%)) > 0 Then
                            ' Get property type
                            FOUNDGROUP% = 0
                            For lookup% = 1 To TGIDX%
                                If propertytype(lookup%) = rs3("group" + CStr(eachone%)) Then
                                    FOUNDGROUP% = lookup%
                                    lookup% = TGIDX%
                                End If
                            Next lookup%
                            If FOUNDGROUP% = 0 Then
                                TGIDX% = TGIDX% + 1
                                propertytype(TGIDX%) = rs3("group" + CStr(eachone%))
                                FOUNDGROUP% = TGIDX%
                            End If
                            ' Get amount
                            propertyvalue(FOUNDGROUP%, GROUPTYPE%) = propertyvalue(FOUNDGROUP%, GROUPTYPE%) + rs2(typegroup(GROUPTYPE%) + "value" + CStr(eachone%))
                            FOUNDVALUE = True
                            If GROUPTYPE% = 4 Then
                                ' Get recovered vehicles
                                If Not IsNull(rs3("numvehicler" + CStr(eachone%))) Then
                                    numvehr(FOUNDGROUP%) = numvehr(FOUNDGROUP%) + rs3("numvehicler" + Mid$(Str$(eachone%), 2))
                                End If
                                ' Get recovered date
                                If IsDate(rs3("daterecovered" + CStr(eachone%))) Then
                                    propertydate(FOUNDGROUP%) = rs3("daterecovered" + CStr(eachone%))
                                End If
                            End If
                            If GROUPTYPE% = 1 Then
                                ' Get stolen vehicles
                                If Not IsNull(rs3("numvehicles" + CStr(eachone%))) Then
                                    numvehs(FOUNDGROUP%) = numvehs(FOUNDGROUP%) + rs3("numvehicles" + Mid$(Str$(eachone%), 2))
                                End If
                            End If
                            ' Get drug information for seized
                            If rs3("pucr" + CStr(eachone%)) = "35A" And propertytype(eachone%) = "10" And GROUPTYPE% = 5 Then
                                For drugspicked% = 1 To 3
                                    If Not IsNull(rs2("ptypeofdrug" + CStr(eachone%) + CStr(drugspicked%))) Then
                                        pdt(FOUNDGROUP%, drugspicked%) = rs2("ptypeofdrug" + CStr(eachone%) + CStr(drugspicked%))
                                        pdm(FOUNDGROUP%, drugspicked%) = rs2("pdrugmeasurement" + CStr(eachone%) + CStr(drugspicked%))
                                        pdq(FOUNDGROUP%, drugspicked%) = rs2("pdrugamt" + CStr(eachone%) + CStr(drugspicked%))
                                    End If
                                Next drugspicked%
                            End If
                        End If
                    End If
                    If FOUNDGROUP% > TGIDX% Then
                        TGIDX% = FOUNDGROUP%
                    End If
                    Next GROUPTYPE%
                End If
            Next eachone%
            rs2.MoveNext
        Wend
    End If
    allucr = allucr + Space$(30 - Len(allucr))
    allucr = Left$(allucr, 30)
    Set rs2 = db.OpenRecordset("select * from incidentreporto where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    'rs2.MoveFirst
    rs2.Edit
    causeofloss$ = "       "
    For QQ% = 1 To TGIDX%
burned:
    If propertyvalue(QQ%, 3) = 0 Or Mid$(causeofloss$, 3, 1) = "X" Then
        GoTo counterfeited
    End If
    If propertyvalue(QQ%, 3) = 9999999 Then
        propertyvalue(QQ%, 3) = 0
    End If
    Mid$(causeofloss$, 3, 1) = "X"
    outrec = "    " + "3" + Mid$(INCSEG, 4, 1) + EXPMONTH + EXPYEAR + Space(4) + orinumber + Format$(rs2("incidentnumber"), "!@@@@@@@@@@@@")
    tv = 0
    ' Type Property/Loss
    outrec = outrec + "2"
    For yy% = 1 To 10
        If propertyvalue(yy%, 3) > 0 Or propertytype(yy%) = "09" Or propertytype(yy%) = "22" Or propertytype(yy%) = "77" Or propertytype(yy%) = "99" Then
            tv = 1
            ' Property Description
            outrec = outrec + propertytype(yy%)
            'Value of Property
            CPV = propertyvalue(yy%, 3)
            outrec = outrec + Format$(CPV, "000000000")
            'Date Recovered
            outrec = outrec + "        "
        End If
    Next yy%
    outrec = outrec + Space$(228 - Len(outrec))
    If TGIDX% > 10 Then
        tv = 0
        For tt% = 10 To TGIDX
            tv = tv + propertyvalue(tt%, 3)
        Next tt%
        outrec = outrec + "77"
        'Value of Property
        outrec = outrec + Format$(tv, "000000000")
        'Date Recovered
        outrec = outrec + "        "
    End If
    'Number of Stolen Vehicles
    outrec = outrec + "  "
    'Number of Recovered Vehicles
    outrec = outrec + "  "
    For yy% = 1 To 3
        'Suspected Drug Type
        outrec = outrec + " "
        'Estimated Drug Quantity
        outrec = outrec + "         "
        'Estimated Drug Quantity Fraction
        outrec = outrec + "   "
        'Type Drug Measurement
        outrec = outrec + "  "
    Next yy%
    'UCR Codes
    If Mid$(outrec, 6, 1) = "W" Or (Mid$(outrec, 6, 1) = "M" And wonce = 1) Then
        outrec = outrec + allucr
    Else
        outrec = outrec + Space$(30)
    End If
    '===== SC Enhancements
    'R ORI
    outrec = outrec + "    "
    If tv = 1 Then
        oridx = oridx + 1
        houtrec(oridx) = outrec
        PCT% = PCT% + 1
    End If
    If Mid$(outrec, 6, 1) = "W" Then
        rs2("oncew") = 1
        rs2.Update
    End If
counterfeited:
    If propertyvalue(QQ%, 6) = 0 Or Mid$(causeofloss$, 6, 1) = "X" Then
        GoTo damaged
    End If
    If propertyvalue(QQ%, 6) = 9999999 Then
        propertyvalue(QQ%, 6) = 0
    End If
    Mid$(causeofloss$, 6, 1) = "X"
    outrec = "    " + "3" + Mid$(INCSEG, 4, 1) + EXPMONTH + EXPYEAR + Space(4) + orinumber + Format$(rs2("incidentnumber"), "!@@@@@@@@@@@@")
    tv = 0
    ' Type Property/Loss
    outrec = outrec + "3"
    For yy% = 1 To 10
        'Property Description
        If propertyvalue(yy%, 6) > 0 Or propertytype(yy%) = "09" Or propertytype(yy%) = "22" Or propertytype(yy%) = "77" Or propertytype(yy%) = "99" Then
            tv = 1
            outrec = outrec + propertytype(yy%)
            outrec = outrec + Format$(CLng(propertyvalue(yy%, 6)), "000000000")
            outrec = outrec + "        "
        End If
    Next yy%
    outrec = outrec + Space$(228 - Len(outrec))
     If TGIDX% > 10 Then
        tv = 0
        For tt% = 10 To TGIDX
            tv = tv + propertyvalue(tt%, 6)
        Next tt%
        outrec = outrec + "77"
        'Value of Property
        outrec = outrec + Format$(tv, "000000000")
        'Date Recovered
        outrec = outrec + "        "
    End If
    'Number of Stolen Vehicles
    outrec = outrec + "  "
    'Number of Recovered Vehicles
    outrec = outrec + "  "
    For yy% = 1 To 3
        'Suspected Drug Type
        outrec = outrec + " "
        'Estimated Drug Quantity
        outrec = outrec + "         "
        'Estimated Drug Quantity Fraction
        outrec = outrec + "   "
        'Type Drug Measurement
        outrec = outrec + "  "
    Next yy%
    'UCR Codes
    If Mid$(outrec, 6, 1) = "W" Or (Mid$(outrec, 6, 1) = "M" And wonce = 1) Then
        outrec = outrec + allucr
    Else
        outrec = outrec + Space$(30)
    End If
    '===== SC Enhancements
    'R ORI
    outrec = outrec + "    "
    If tv = 1 Then
        oridx = oridx + 1
        houtrec(oridx) = outrec
        PCT% = PCT% + 1
    End If
damaged:
    If propertyvalue(QQ%, 2) = 0 Or Mid$(causeofloss$, 2, 1) = "X" Then
        GoTo recovered
    End If
    If propertyvalue(QQ%, 2) = 9999999 Then
        propertyvalue(QQ%, 2) = 0
    End If
    Mid$(causeofloss$, 2, 1) = "X"
    outrec = "    " + "3" + Mid$(INCSEG, 4, 1) + EXPMONTH + EXPYEAR + Space(4) + orinumber + Format$(rs2("incidentnumber"), "!@@@@@@@@@@@@")
    tv = 0
    ' Type Property/Loss
    outrec = outrec + "4"
    For yy% = 1 To 10
        'Property Description
        If propertyvalue(yy%, 2) > 0 Or propertytype(yy%) = "09" Or propertytype(yy%) = "22" Or propertytype(yy%) = "77" Or propertytype(yy%) = "99" Then
            tv = 1
            outrec = outrec + propertytype(yy%)
            outrec = outrec + Format$(CLng(propertyvalue(yy%, 2)), "000000000")
            outrec = outrec + "        "
        End If
    Next yy%
    outrec = outrec + Space$(228 - Len(outrec))
    If TGIDX% > 10 Then
        tv = 0
        For tt% = 10 To TGIDX
            tv = tv + propertyvalue(tt%, 2)
        Next tt%
        outrec = outrec + "77"
        'Value of Property
        outrec = outrec + Format$(tv, "000000000")
        'Date Recovered
        outrec = outrec + "        "
    End If
    'Number of Stolen Vehicles
    outrec = outrec + "  "
    'Number of Recovered Vehicles
    outrec = outrec + "  "
    For yy% = 1 To 3
        'Suspected Drug Type
        outrec = outrec + " "
        'Estimated Drug Quantity
        outrec = outrec + "         "
        'Estimated Drug Quantity Fraction
        outrec = outrec + "   "
        'Type Drug Measurement
        outrec = outrec + "  "
    Next yy%
    'UCR Codes
    If Mid$(outrec, 6, 1) = "W" Or (Mid$(outrec, 6, 1) = "M" And wonce = 1) Then
        outrec = outrec + allucr
    Else
        outrec = outrec + Space$(30)
    End If
    '===== SC Enhancements
    'R ORI
    outrec = outrec + "    "
    If tv = 1 Then
        oridx = oridx + 1
        houtrec(oridx) = outrec
        PCT% = PCT% + 1
    End If
recovered:
    If propertyvalue(QQ%, 4) = 0 Or Mid$(causeofloss$, 4, 1) = "X" Then
        GoTo seized
    End If
    If propertyvalue(QQ%, 4) = 9999999 Then
        propertyvalue(QQ%, 4) = 0
    End If
    Mid$(causeofloss$, 4, 1) = "X"
    outrec = "    " + "3" + Mid$(INCSEG, 4, 1) + EXPMONTH + EXPYEAR + Space(4) + orinumber + Format$(rs2("incidentnumber"), "!@@@@@@@@@@@@")
    tv = 0
    nvr% = 0
    ' Type Property/Loss
    outrec = outrec + "5"
    For yy% = 1 To 10
        'Property Description
        If propertyvalue(yy%, 4) > 0 Or propertytype(yy%) = "09" Or propertytype(yy%) = "22" Or propertytype(yy%) = "77" Or propertytype(yy%) = "99" Then
            tv = 1
            outrec = outrec + propertytype(yy%)
            outrec = outrec + Format$(CLng(propertyvalue(yy%, 4)), "000000000")
            nvr% = nvr% + numvehr(yy%)
            If IsDate(propertydate(yy%)) Then
                outrec = outrec + Format$(propertydate(yy%), "yyyymmdd")
            Else
                outrec = outrec + "        "
            End If
        End If
    Next yy%
    outrec = outrec + Space$(228 - Len(outrec))
     If TGIDX% > 10 Then
        tv = 0
        For tt% = 10 To TGIDX
            tv = tv + propertyvalue(tt%, 4)
            nvr% = nvr% + numvehr(yy%)
        Next tt%
        outrec = outrec + "77"
        'Value of Property
        outrec = outrec + Format$(tv, "000000000")
        'Date Recovered
        outrec = outrec + "        "
    End If
    'Number of Stolen Vehicles
    outrec = outrec + "  "
    'Number of Recovered Vehicles
    outrec = outrec + Format$(nvr%, "00")
    For yy% = 1 To 3
        'Suspected Drug Type
        outrec = outrec + " "
        'Estimated Drug Quantity
        outrec = outrec + "         "
        'Estimated Drug Quantity Fraction
        outrec = outrec + "   "
        'Type Drug Measurement
        outrec = outrec + "  "
    Next yy%
    'UCR Codes
    If Mid$(outrec, 6, 1) = "W" Or (Mid$(outrec, 6, 1) = "M" And wonce = 1) Then
        outrec = outrec + allucr
    Else
        outrec = outrec + Space$(30)
    End If
    '===== SC Enhancements
    'R ORI
    outrec = outrec + "    "
    If tv = 1 Then
        oridx = oridx + 1
        houtrec(oridx) = outrec
        PCT% = PCT% + 1
    End If
seized:
    If propertyvalue(QQ%, 5) = 0 Or Mid$(causeofloss$, 5, 1) = "X" Then
        If Not PROPERTYDRUGS(QQ%, 5) Then
            GoTo stolen
        End If
    End If
    Mid$(causeofloss$, 5, 1) = "X"
    outrec = "    " + "3" + Mid$(INCSEG, 4, 1) + EXPMONTH + EXPYEAR + Space(4) + orinumber + Format$(rs2("incidentnumber"), "!@@@@@@@@@@@@")
    tv = 0
    ' Type Property/Loss
    outrec = outrec + "6"
    For yy% = 1 To 10
        'Property Description
        If propertytype(yy%) = "10" Then
            tv = 1
            outrec = outrec + propertytype(yy%)
            outrec = outrec + "                 "
        Else
        If propertyvalue(yy%, 5) > 0 Or propertytype(yy%) = "09" Or propertytype(yy%) = "22" Or propertytype(yy%) = "77" Or propertytype(yy%) = "99" Then
            tv = 1
            outrec = outrec + propertytype(yy%)
            outrec = outrec + Format$(CLng(propertyvalue(yy%, 5)), "000000000")
            outrec = outrec + "        "
        End If
        End If
    Next yy%
    outrec = outrec + Space$(228 - Len(outrec))
     If TGIDX% > 10 Then
        tv = 0
        For tt% = 10 To TGIDX
            tv = tv + propertyvalue(tt%, 5)
        Next tt%
        outrec = outrec + "77"
        'Value of Property
        outrec = outrec + Format$(tv, "000000000")
        'Date Recovered
        outrec = outrec + "        "
    End If
    'Number of Stolen Vehicles
    outrec = outrec + "  "
    'Number of Recovered Vehicles
    outrec = outrec + "  "
    If propertytype(QQ%) = "10" Then
        For yy% = 1 To 3
            'Suspected Drug Type
            outrec = outrec + pdt(QQ%, yy%) + Space$(1 - Len(pdt(QQ%, yy%)))
            'Estimated Drug Quantity
            If pdt(QQ%, yy%) > "" Then
                outrec = outrec + Format$(Int(pdq(QQ%, yy%)), "000000000")
            Else
                outrec = outrec + Space$(9)
            End If
            'Estimated Drug Quantity Fraction
            If Int(pdq(QQ%, yy%)) <> pdq(QQ%, yy%) Then
                frac! = pdq(QQ%, yy%) - Int(pdq(QQ%, yy%))
                frac! = frac! * 1000
            Else
                frac! = 0
            End If
            If pdt(QQ%, yy%) > "" Then
                outrec = outrec + Format$(frac!, "000")
            Else
                outrec = outrec + "   "
            End If
            'Type Drug Measurement
            If pdt(QQ%, yy%) > "" Then
                outrec = outrec + pdm(QQ%, yy%) + Space$(2 - Len(pdm(QQ%, yy%)))
            Else
                outrec = outrec + "  "
            End If
        Next yy%
    Else
        outrec = outrec + Space$(45)
    End If
    'UCR Codes
    If Mid$(outrec, 6, 1) = "W" Or (Mid$(outrec, 6, 1) = "M" And wonce = 1) Then
        outrec = outrec + allucr
    Else
        outrec = outrec + Space$(30)
    End If
    '===== SC Enhancements
    'R ORI
    outrec = outrec + "    "
    If tv = 1 Then
        oridx = oridx + 1
        houtrec(oridx) = outrec
        PCT% = PCT% + 1
    End If
stolen:
    If propertyvalue(QQ%, 1) = 0 Or Mid$(causeofloss$, 1, 1) = "X" Then
        GoTo unknown
    End If
    If propertyvalue(QQ%, 1) = 9999999 Then
        propertyvalue(QQ%, 1) = 0
    End If
    Mid$(causeofloss$, 1, 1) = "X"
    outrec = "    " + "3" + Mid$(INCSEG, 4, 1) + EXPMONTH + EXPYEAR + Space(4) + orinumber + Format$(rs2("incidentnumber"), "!@@@@@@@@@@@@")
    tv = 0
    nvs% = 0
    ' Type Property/Loss
    outrec = outrec + "7"
    For yy% = 1 To 10
        'Property Description
        If propertyvalue(yy%, 1) > 0 Or propertytype(yy%) = "09" Or propertytype(yy%) = "22" Or propertytype(yy%) = "77" Or propertytype(yy%) = "99" Then
            tv = 1
            outrec = outrec + propertytype(yy%)
            If propertyvalue(yy%, 1) = 9999999 Then
                propertyvalue(yy%, 1) = 0
            End If
            outrec = outrec + Format$(CLng(propertyvalue(yy%, 1)), "000000000")
            nvs% = nvs% + numvehs(yy%)
            'If IsDate(propertydate(yy%)) Then
            '    outrec = outrec + Format$(propertydate(yy%), "yyyymmdd")
            'Else
                outrec = outrec + "        "
            'End If
        End If
    Next yy%
    outrec = outrec + Space$(228 - Len(outrec))
     If TGIDX% = 10 Then
        'Property Description
        If propertyvalue(10, 1) > 0 Then
            outrec = outrec + propertytype(10)
            nvs% = nvs% + numvehs(yy%)
            'Value of Property
            outrec = outrec + Format$(CLng(propertyvalue(10, 1)), "000000000")
            'Date Recovered
            'If IsDate(propertydate(10)) Then
            '    outrec = outrec + Format$(propertydate(10), "yyyymmdd")
            'Else
                outrec = outrec + "        "
            'End If
        Else
            outrec = outrec + "                   "
        End If
    End If
    If TGIDX% > 10 Then
        tv = 0
        For tt% = 10 To TGIDX
            tv = tv + propertyvalue(tt%, 1)
            nvr% = nvr% + numvehr(yy%)
        Next tt%
        outrec = outrec + "77"
        'Value of Property
        outrec = outrec + Format$(tv, "000000000")
        'Date Recovered
        outrec = outrec + "        "
    End If
    'Number of Stolen Vehicles
    outrec = outrec + Format$(nvs%, "00")
    'Number of Recovered Vehicles
    outrec = outrec + "  "
    For yy% = 1 To 3
        'Suspected Drug Type
        outrec = outrec + " "
        'Estimated Drug Quantity
        outrec = outrec + "         "
        'Estimated Drug Quantity Fraction
        outrec = outrec + "   "
        'Type Drug Measurement
        outrec = outrec + "  "
    Next yy%
    'UCR Codes
    If Mid$(outrec, 6, 1) = "W" Or (Mid$(outrec, 6, 1) = "M" And wonce = 1) Then
        outrec = outrec + allucr
    Else
        outrec = outrec + Space$(30)
    End If
    '===== SC Enhancements
    'R ORI
    outrec = outrec + "    "
    If tv = 1 Then
        oridx = oridx + 1
        houtrec(oridx) = outrec
        PCT% = PCT% + 1
    End If
unknown:
    If propertyvalue(QQ%, 7) = 0 Or Mid$(causeofloss$, 7, 1) = "X" Then
        GoTo pgo
    End If
    If propertyvalue(QQ%, 7) = 9999999 Then
        propertyvalue(QQ%, 7) = 0
    End If
    Mid$(causeofloss$, 7, 1) = "X"
    outrec = "    " + "3" + Mid$(INCSEG, 4, 1) + EXPMONTH + EXPYEAR + Space(4) + orinumber + Format$(rs2("incidentnumber"), "!@@@@@@@@@@@@")
    tv = 0
    ' Type Property/Loss
    outrec = outrec + "8"
    For yy% = 1 To 10
        'Property Description
        If propertyvalue(yy%, 7) > 0 Or propertytype(yy%) = "09" Or propertytype(yy%) = "22" Or propertytype(yy%) = "77" Or propertytype(yy%) = "99" Then
            tv = 1
            outrec = outrec + propertytype(yy%)
            outrec = outrec + Format$(CLng(propertyvalue(yy%, 7)), "000000000")
            outrec = outrec + "        "
        End If
    Next yy%
    outrec = outrec + Space$(228 - Len(outrec))
     If TGIDX% = 10 Then
        'Property Description
        If propertyvalue(10, 7) > 0 Then
            outrec = outrec + propertytype(10)
            'Value of Property
            outrec = outrec + Format$(CLng(propertyvalue(10, 7)), "000000000")
            outrec = outrec + "        "
        Else
            outrec = outrec + "                   "
        End If
    End If
    If TGIDX% > 10 Then
        tv = 0
        For tt% = 10 To TGIDX
            tv = tv + propertyvalue(tt%, 7)
        Next tt%
        outrec = outrec + "77"
        'Value of Property
        outrec = outrec + Format$(tv, "000000000")
        'Date Recovered
        outrec = outrec + "        "
    End If
    'Number of Stolen Vehicles
    outrec = outrec + "  "
    'Number of Recovered Vehicles
    outrec = outrec + "  "
    For yy% = 1 To 3
        'Suspected Drug Type
        outrec = outrec + " "
        'Estimated Drug Quantity
        outrec = outrec + "         "
        'Estimated Drug Quantity Fraction
        outrec = outrec + "   "
        'Type Drug Measurement
        outrec = outrec + "  "
    Next yy%
    'UCR Codes
    If Mid$(outrec, 6, 1) = "W" Or (Mid$(outrec, 6, 1) = "M" And wonce = 1) Then
        outrec = outrec + allucr
    Else
        outrec = outrec + Space$(30)
    End If
    '===== SC Enhancements
    'R ORI
    outrec = outrec + "    "
    If tv = 1 Then
        oridx = oridx + 1
        houtrec(oridx) = outrec
        PCT% = PCT% + 1
    End If
pgo:
    If Mid$(outrec, 6, 1) = "W" Then
        rs2.Edit
        rs2("ONCEW") = 1
        rs2.Update
    End If
    Next QQ%
victim:
    '==== SC Enhancements - Counter for Administrative Segment for Property
    If Mid$(INCSEG, 2, 1) <> " " Then
        houtrec(1) = houtrec(1) + Format$(PCT%, "00")
    End If

    If Mid$(INCSEG, 5, 1) = " " Or ANYA = 0 Then
        VCT% = 0
        GoTo finishv
    End If
    '===== Victim Segment =====
    VCT% = 0
    Set rs2 = db.OpenRecordset("select * from incidentreportv where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    If Not rs2.EOF Then
        'rs2.MoveFirst
    Else
        GoTo finishv
    End If
    AHERE% = 0
    For tt% = 1 To 5
        If Not IsNull(rs("vucr1" + Mid$(Str$(tt%), 2))) Then
            Set rs4 = db.OpenRecordset("SELECT abGROUP FROM UCR WHERE abGROUP = 'A' AND abbrev = '" + rs("VUCR1" + Mid$(Str$(tt%), 2)) + "'")
            If Not rs4.EOF Then
                AHERE% = 1
                tt% = 5
            End If
        End If
    Next tt%
    If AHERE% = 0 Then
        GoTo CHECKSUPPLEMENTALNEXT
    End If
    outrec = "    " + "4" + Mid$(INCSEG, 5, 1) + EXPMONTH + EXPYEAR + Space(4) + orinumber + Format$(rs2("incidentnumber"), "!@@@@@@@@@@@@")
    ' Victim Number
    outrec = outrec + "001"
    'Victim Connected to UCR Offense
    h1$ = "  "
    H2$ = "  "
    AD$ = " "
    cttt% = 0
    For poi% = 1 To 5
        vuc(poi%) = ""
    Next poi%
    For tt% = 1 To 5
        If Not IsNull(rs("vucr1" + Mid$(Str$(tt%), 2))) Then
            Set rs4 = db.OpenRecordset("SELECT abGROUP FROM UCR WHERE abGROUP = 'A' AND abbrev = '" + rs("VUCR1" + Mid$(Str$(tt%), 2)) + "'")
            If Not rs4.EOF Then
                cttt% = cttt% + 1
                outrec = outrec + rs("vucr1" + Mid$(Str$(tt%), 2))
                vuc(tt%) = rs("vucr1" + Mid$(Str$(tt%), 2))
                Select Case rs("vucr1" + Mid$(Str$(tt%), 2))
                    Case "09A", "09B", "09C", "13A"
                        For ttt% = 1 To 10
                            If rs("UCR" + Mid$(Str$(ttt%), 2)) = rs("vucr1" + Mid$(Str$(tt%), 2)) Then
                                If Not IsNull(rs("HOMOCIDE1" + Mid$(Str$(ttt%), 2))) Then
                                    h1$ = rs("HOMOCIDE1" + Mid$(Str$(ttt%), 2))
                                End If
                                If Not IsNull(rs("HOMOCIDE2" + Mid$(Str$(ttt%), 2))) Then
                                    H2$ = rs("HOMOCIDE2" + Mid$(Str$(ttt%), 2))
                                End If
                                If Not IsNull(rs("ADDITIONAL" + Mid$(Str$(ttt%), 2))) Then
                                    AD$ = rs("ADDITIONAL" + Mid$(Str$(ttt%), 2))
                                End If
                                ttt% = 10
                            End If
                        Next ttt%
                End Select
            End If
        End If
    Next tt%
    For yy% = cttt% + 1 To 5
        outrec = outrec + "   "
    Next yy%
    outrec = outrec + Space$(15)
    'Type of Victim
    Set rs3 = db.OpenRecordset("select individual, policeofficer, business, financialinstitution, other,government,unknown,religiousorganization,SOCIETYPUBLIC from incidentreportc where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    'rs3.MoveFirst
    If rs3("individual") Or rs3("policeofficer") Then
        outrec = outrec + "I"
    Else
    If rs3("business") Then
        outrec = outrec + "B"
    Else
    If rs3("financialinstitution") Then
        outrec = outrec + "F"
    Else
    If rs3("other") Then
        outrec = outrec + "O"
    Else
    If rs3("government") Then
        outrec = outrec + "G"
    Else
    If rs3("unknown") Then
        outrec = outrec + "U"
    Else
    If rs3("religiousorganization") Then
        outrec = outrec + "R"
    Else
    If rs3("societyPUBLIC") Then
        outrec = outrec + "S"
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    If Right$(outrec, 1) <> "I" Then
        outrec = outrec + "        "
    Else
        'Age of Victim
        If Len(rs2("vage")) = 2 Then
            outrec = outrec + rs2("vage") + "  "
        Else
        If Len(rs2("vage")) = 1 Then
            outrec = outrec + Format$(Val(rs2("vage")), "00") + "  "
        Else
        If Len(rs2("vage")) = 4 Then
            outrec = outrec + rs2("vage")
        Else
            outrec = outrec + "    "
        End If
        End If
        End If
        'Sex of Victim
        outrec = outrec + rs2("vsex") + Space$(1 - Len(rs2("vsex")))
        'Race of Victim
        outrec = outrec + rs2("vrace") + Space$(1 - Len(rs2("vrace")))
        'Ethnicity of Victim
        outrec = outrec + rs2("vethnicity") + Space$(1 - Len(rs2("vethnicity")))
        'Resident Status of Victim
        outrec = outrec + rs2("vresident") + Space$(1 - Len(rs2("vresident")))
    End If
    'Aggravated Assault/Homocide Circumstances 1
    outrec = outrec + h1$
    'Aggravated Assault/Homocide Circumstances 2
    outrec = outrec + H2$
    'Additional Justifiable Homocide Circumstances
    outrec = outrec + AD$
    'Type Injury 1 - 5
    For bb% = 1 To 5
        If Not IsNull(rs2("typeofinjury" + Mid$(Str$(bb%), 2))) And rs2("typeofinjury" + Mid$(Str$(bb%), 2)) > "" Then
            Call checkinjury
            If foundinjtype Then
                outrec = outrec + rs2("typeofinjury" + Mid$(Str$(bb%), 2))
            Else
                outrec = outrec + " "
            End If
        Else
            outrec = outrec + " "
        End If
    Next bb%
    'Relationship(s) of Victim to Offender(s)
    needrelationship% = 0
    For cc% = 1 To 5
        If Not IsNull(rs("vucr1" + Mid$(Str$(cc%), 2))) Then
            uc$ = rs("vucr1" + Mid$(Str$(cc%), 2))
            Select Case uc$
                Case "13A", "13B", "13C", "09A", "09B", "09C", "100", "11A", "11B", "11C", "11D", "36A", "36B", "120"
                    needrelationship% = 1
            End Select
            If needrelationship% = 1 Then
                cc% = 10
            End If
        End If
    Next cc%
    totrel% = 0
    If Not rs3("individual") And Not rs3("policeofficer") Then
        needrelationship% = 0
    End If
    If needrelationship% = 1 Then
        For cc% = 1 To 3
            If Not IsNull(rs2("vrelationship" + Mid$(Str$(cc%), 2))) Then
                If cc% = 1 Then
                    Set rs4 = db.OpenRecordset("select * from incidentreports where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
                    If Not rs4.EOF Then
                        'rs4.MoveFirst
                        If UCase(rs4("sname")) = "UNKNOWN" Then
                            If rs4("sage") = "00" And rs4("ssex") = "U" And rs4("srace") = "U" Then
                                outrec = outrec + "00  "
                            Else
                                outrec = outrec + Format$(cc%, "00")
                                outrec = outrec + rs2("vrelationship" + Mid$(Str$(cc%), 2))
                            End If
                        Else
                            outrec = outrec + Format$(cc%, "00")
                            outrec = outrec + rs2("vrelationship" + Mid$(Str$(cc%), 2))
                        End If
                    Else
                        outrec = outrec + Format$(cc%, "00")
                        outrec = outrec + rs2("vrelationship" + Mid$(Str$(cc%), 2))
                    End If
                Else
                    outrec = outrec + Format$(cc%, "00")
                    outrec = outrec + rs2("vrelationship" + Mid$(Str$(cc%), 2))
                End If
                totrel% = totrel% + 1
            Else
                outrec = outrec + "    "
            End If
        Next cc%
        For cc% = 4 To 10
            If Not IsNull(rs("vrelationship" + Mid$(Str$(cc%), 2))) Then
                If cc% = 1 Then
                    Set rs4 = db.OpenRecordset("select * from incidentreports where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
                    If Not rs4.EOF Then
                        'rs4.MoveFirst
                        If UCase(rs4("sname")) = "UNKNOWN" Then
                            If rs4("sage") = "00" And rs4("ssex") = "U" And rs4("srace") = "U" Then
                                outrec = outrec + "00  "
                            Else
                                outrec = outrec + Format$(cc%, "00")
                                outrec = outrec + rs("vrelationship" + Mid$(Str$(cc%), 2))
                            End If
                        Else
                            outrec = outrec + Format$(cc%, "00")
                            outrec = outrec + rs("vrelationship" + Mid$(Str$(cc%), 2))
                        End If
                    Else
                        outrec = outrec + Format$(cc%, "00")
                        outrec = outrec + rs("vrelationship" + Mid$(Str$(cc%), 2))
                    End If
                Else
                    outrec = outrec + Format$(cc%, "00")
                    outrec = outrec + rs("vrelationship" + Mid$(Str$(cc%), 2))
                End If
                totrel% = totrel% + 1
            Else
                outrec = outrec + "    "
            End If
        Next cc%
    Else
        outrec = outrec + Space$(40)
    End If
    If totrel% = 0 Then
        totrel% = 1
    End If
    ' SC Enhancements
    ' Leoka Activity
    If Not IsNull(rs("lactivity")) Then
        outrec = outrec + rs("lactivity")
        ' Leoka Assignment
        If rs2("vtwomanvehicle") = "X" Then
            outrec = outrec + "1"
        Else
        If rs2("vonemanvehicle") = "X" Then
            If rs2("valone") = "X" Then
                outrec = outrec + "2"
            Else
                outrec = outrec + "3"
            End If
        Else
        If rs2("vdetective") = "X" Then
            If rs2("valone") = "X" Then
                outrec = outrec + "4"
            Else
                outrec = outrec + "%"
            End If
        Else
        If rs2("vother") = "X" Then
            If rs2("valone") = "X" Then
                outrec = outrec + "6"
            Else
                outrec = outrec + "7"
            End If
        Else
            outrec = outrec + " "
        End If
        End If
        End If
        End If
    Else
        outrec = outrec + "  "
    End If
    oridx = oridx + 1
    houtrec(oridx) = outrec
    VCT% = VCT% + 1
checksupplemental:
    Set rs2 = db.OpenRecordset("select * from supplemental where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34) + " and (victim1 > 0 or victim2 > 0)")
    If Not rs2.EOF Then
        'rs2.MoveFirst
    Else
        GoTo finishv
    End If
    While Not rs2.EOF
        If Not IsNull(rs2("victim1")) And rs2("victim1") < 100 Then
            Set rs3 = db.OpenRecordset("select * from supplementalsupport where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34) + " and page = " + Str$(rs2("page")))
            'rs3.MoveFirst
            AHERE% = 0
            For tt% = 1 To 5
                If Not IsNull(rs3("vucr1" + Mid$(Str$(tt%), 2))) Then
                    Set rs4 = db.OpenRecordset("SELECT abGROUP FROM UCR WHERE abGROUP = 'A' AND abbrev = '" + rs3("VUCR1" + Mid$(Str$(tt%), 2)) + "'")
                    If Not rs4.EOF Then
                        AHERE% = 1
                        tt% = 5
                    End If
                End If
            Next tt%
            If AHERE% = 0 Then
                GoTo checksupplemental2
            End If
            ' Record Descriptor Word
            outrec = "    " + "4" + Mid$(INCSEG, 5, 1) + EXPMONTH + EXPYEAR + Space(4) + orinumber + Format$(rs2("incidentnumber"), "!@@@@@@@@@@@@")
            ' Victim Number
            outrec = outrec + Format$(rs2("victim1"), "000")
            'Victim Connected to UCR Offense
            h1$ = "  "
            H2$ = "  "
            AD$ = " "
            cttt% = 0
            For poi% = 1 To 5
                vuc(poi%) = ""
            Next poi%
            For tt% = 1 To 5
                If Not IsNull(rs3("vucr1" + Mid$(Str$(tt%), 2))) Then
                    Set rs4 = db.OpenRecordset("SELECT abGROUP FROM UCR WHERE abGROUP = 'A' AND abbrev = '" + rs3("VUCR1" + Mid$(Str$(tt%), 2)) + "'")
                    If Not rs4.EOF Then
                        cttt% = cttt% + 1
                        outrec = outrec + rs3("vucr1" + Mid$(Str$(tt%), 2))
                        vuc(tt%) = rs3("vucr1" + Mid$(Str$(tt%), 2))
                    End If
                    Select Case rs3("vucr1" + Mid$(Str$(tt%), 2))
                        Case "09A", "09B", "09C", "13A"
                            For ttt% = 1 To 10
                                If rs("UCR" + Mid$(Str$(ttt%), 2)) = rs3("vucr1" + Mid$(Str$(tt%), 2)) Then
                                    If Not IsNull(rs("HOMOCIDE1" + Mid$(Str$(ttt%), 2))) Then
                                        h1$ = rs("HOMOCIDE1" + Mid$(Str$(ttt%), 2))
                                    End If
                                    If Not IsNull(rs("HOMOCIDE2" + Mid$(Str$(ttt%), 2))) Then
                                        H2$ = rs("HOMOCIDE2" + Mid$(Str$(ttt%), 2))
                                    End If
                                    If Not IsNull(rs("ADDITIONAL" + Mid$(Str$(ttt%), 2))) Then
                                        AD$ = rs("ADDITIONAL" + Mid$(Str$(ttt%), 2))
                                    End If
                                    ttt% = 10
                                End If
                            Next ttt%
                    End Select
                End If
            Next tt%
            For yy% = cttt% + 1 To 5
                outrec = outrec + "   "
            Next yy%
            outrec = outrec + Space$(15)
            'Type of Victim
            If rs2("individual1") Or rs2("policeofficer1") Then
                outrec = outrec + "I"
            Else
            If rs2("business1") Then
                outrec = outrec + "B"
            Else
            If rs2("financialinstitution1") Then
                outrec = outrec + "F"
            Else
            If rs2("other1") Then
                outrec = outrec + "O"
            Else
            If rs2("government1") Then
                outrec = outrec + "G"
            Else
            If rs2("unknown1") Then
                outrec = outrec + "U"
            Else
            If rs2("religiousorganization1") Then
                outrec = outrec + "R"
            Else
            If rs2("society1") Then
                outrec = outrec + "S"
            End If
            End If
            End If
            End If
            End If
            End If
            End If
            End If
            If Right$(outrec, 1) <> "I" Then
                outrec = outrec + "        "
            Else
                'Age of Victim
                If Len(rs2("age1")) = 2 Then
                    outrec = outrec + rs2("age1") + "  "
                Else
                If Len(rs2("age1")) = 1 Then
                    outrec = outrec + Format$(Val(rs2("age1")), "00") + "  "
                Else
                If Len(rs2("age1")) = 4 Then
                    outrec = outrec + rs2("age1")
                Else
                    outrec = outrec + "    "
                End If
                End If
                End If
                'Sex of Victim
                outrec = outrec + rs2("sex1") + Space$(1 - Len(rs2("sex1")))
                'Race of Victim
                outrec = outrec + rs2("race1") + Space$(1 - Len(rs2("race1")))
                'Ethnicity of Victim
                outrec = outrec + rs2("ethnicity1") + Space$(1 - Len(rs2("ethnicity1")))
                'Resident Status of Victim
                outrec = outrec + rs2("resident1") + Space$(1 - Len(rs2("resident1")))
            End If
            'Aggravated Assault/Homocide Circumstances 1
            outrec = outrec + h1$
            'Aggravated Assault/Homocide Circumstances 2
            outrec = outrec + H2$
            'Additional Justifiable Homocide Circumstances
            outrec = outrec + AD$
            'Type Injury 1 - 5
            For bb% = 1 To 5
                If Not IsNull(rs2("typeofinjury1" + Mid$(Str$(bb%), 2))) And rs2("typeofinjury1" + Mid$(Str$(bb%), 2)) > "" Then
                    Call checkinjury
                    If foundinjtype Then
                        outrec = outrec + rs2("typeofinjury1" + Mid$(Str$(bb%), 2))
                    Else
                        outrec = outrec + " "
                    End If
                Else
                    outrec = outrec + " "
                End If
            Next bb%
            'Relationship(s) of Victim to Offender(s)
            needrelationship% = 0
            For cc% = 1 To 10
                If Not IsNull(rs3("vucr1" + Mid$(Str$(cc%), 2))) Then
                    uc$ = rs3("vucr1" + Mid$(Str$(cc%), 2))
                    Select Case uc$
                        Case "13A", "13B", "13C", "09A", "09B", "09C", "100", "11A", "11B", "11C", "11D", "36A", "36B", "120"
                            needrelationship% = 1
                             cc% = 10
                    End Select
                End If
            Next cc%
            totrel% = 0
            If Not rs2("individual1") And Not rs2("policeofficer1") Then
                needrelationship% = 0
            End If
            If needrelationship% = 1 Then
                For cc% = 1 To 3
                    If Not IsNull(rs2("relationship1" + Mid$(Str$(cc%), 2))) Then
                        If cc% = 1 Then
                            Set rs4 = db.OpenRecordset("select * from incidentreports where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
                            If Not rs4.EOF Then
                                'rs4.MoveFirst
                                If UCase(rs4("sname")) = "UNKNOWN" Then
                                    If rs4("sage") = "00" And rs4("ssex") = "U" And rs4("srace") = "U" Then
                                        outrec = outrec + "00  "
                                    Else
                                        outrec = outrec + Format$(cc%, "00")
                                        outrec = outrec + rs2("relationship1" + Mid$(Str$(cc%), 2))
                                    End If
                                Else
                                    outrec = outrec + Format$(cc%, "00")
                                    outrec = outrec + rs2("relationship1" + Mid$(Str$(cc%), 2))
                                End If
                            Else
                                outrec = outrec + Format$(cc%, "00")
                                outrec = outrec + rs2("relationship1" + Mid$(Str$(cc%), 2))
                            End If
                        Else
                            outrec = outrec + Format$(cc%, "00")
                            outrec = outrec + rs2("relationship1" + Mid$(Str$(cc%), 2))
                        End If
                        totrel% = totrel% + 1
                    Else
                        outrec = outrec + "    "
                    End If
                Next cc%
                For cc% = 4 To 10
                    If Not IsNull(rs3("relationship1" + Mid$(Str$(cc%), 2))) Then
                        outrec = outrec + Format$(cc%, "00")
                        outrec = outrec + rs3("relationship1" + Mid$(Str$(cc%), 2))
                        totrel% = totrel% + 1
                    Else
                        outrec = outrec + "    "
                    End If
                Next cc%
            Else
                outrec = outrec + Space$(40)
            End If
            If totrel% = 0 Then
                totrel% = 1
            End If
            ' SC Enhancements
            ' Leoka Activity
            If Not IsNull(rs("lactivity")) Then
                outrec = outrec + rs("lactivity")
            Else
                outrec = outrec + " "
            End If
            ' Leoka Assignment
            If rs2("twomanvehicle1") = "X" Then
                outrec = outrec + "1"
            Else
            If rs2("onemanvehicle1") = "X" Then
                If rs2("alone1") = "X" Then
                    outrec = outrec + "2"
                Else
                    outrec = outrec + "3"
                End If
            Else
            If rs2("detective1") = "X" Then
                If rs2("alone1") = "X" Then
                    outrec = outrec + "4"
                Else
                    outrec = outrec + "%"
                End If
            Else
            If rs2("other1") = "X" Then
                If rs2("alone1") = "X" Then
                    outrec = outrec + "6"
                Else
                    outrec = outrec + "7"
                End If
            Else
                outrec = outrec + " "
            End If
            End If
            End If
            End If
            oridx = oridx + 1
            houtrec(oridx) = outrec
            VCT% = VCT% + 1
        End If
checksupplemental2:
        If Not IsNull(rs2("victim2")) And rs2("victim2") < 100 Then
            Set rs3 = db.OpenRecordset("select * from supplementalsupport where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34) + " and page = " + Str$(rs2("page")))
            'rs3.MoveFirst
            AHERE% = 0
            For tt% = 1 To 5
                If Not IsNull(rs3("vucr2" + Mid$(Str$(tt%), 2))) Then
                    Set rs4 = db.OpenRecordset("SELECT abGROUP FROM UCR WHERE abGROUP = 'A' AND abbrev = '" + rs3("VUCR2" + Mid$(Str$(tt%), 2)) + "'")
                    If Not rs4.EOF Then
                        AHERE% = 1
                        tt% = 5
                    End If
                End If
            Next tt%
            If AHERE% = 0 Then
                GoTo CHECKSUPPLEMENTALNEXT
            End If
            ' Record Descriptor Word
            outrec = "    " + "4" + Mid$(INCSEG, 5, 1) + EXPMONTH + EXPYEAR + Space(4) + orinumber + Format$(rs2("incidentnumber"), "!@@@@@@@@@@@@")
            ' Victim Number
            outrec = outrec + Format$(rs2("victim2"), "000")
            'Victim Connected to UCR Offense
            h1$ = "  "
            H2$ = "  "
            AD$ = " "
            For poi% = 1 To 5
                vuc(poi%) = ""
            Next poi%
            cttt% = 0
            For tt% = 1 To 5
                If Not IsNull(rs3("vucr2" + Mid$(Str$(tt%), 2))) Then
                    cttt% = cttt% + 1
                    outrec = outrec + rs3("vucr2" + Mid$(Str$(tt%), 2))
                    vuc(tt%) = rs3("vucr2" + Mid$(Str$(tt%), 2))
                    Set rs4 = db.OpenRecordset("SELECT abGROUP FROM UCR WHERE abGROUP = 'A' AND abbrev = '" + rs3("VUCR2" + Mid$(Str$(tt%), 2)) + "'")
                    If Not rs4.EOF Then
                        For ttt% = 1 To 10
                            If rs("UCR" + Mid$(Str$(ttt%), 2)) = rs3("vucr2" + Mid$(Str$(tt%), 2)) Then
                                If Not IsNull(rs("HOMOCIDE1" + Mid$(Str$(ttt%), 2))) Then
                                    h1$ = rs("HOMOCIDE1" + Mid$(Str$(ttt%), 2))
                                End If
                                If Not IsNull(rs("HOMOCIDE2" + Mid$(Str$(ttt%), 2))) Then
                                    H2$ = rs("HOMOCIDE2" + Mid$(Str$(ttt%), 2))
                                End If
                                If Not IsNull(rs("ADDITIONAL" + Mid$(Str$(ttt%), 2))) Then
                                    AD$ = rs("ADDITIONAL" + Mid$(Str$(ttt%), 2))
                                End If
                                ttt% = 10
                            End If
                        Next ttt%
                    End If
                End If
            Next tt%
            For yy% = cttt% + 1 To 5
                outrec = outrec + "   "
            Next yy%
            outrec = outrec + Space$(15)
            'Type of Victim
            If rs2("individual2") Or rs2("policeofficer2") Then
                outrec = outrec + "I"
            Else
            If rs2("business2") Then
                outrec = outrec + "B"
            Else
            If rs2("financialinstitution2") Then
                outrec = outrec + "F"
            Else
            If rs2("other2") Then
                outrec = outrec + "O"
            Else
            If rs2("government2") Then
                outrec = outrec + "G"
            Else
            If rs2("unknown2") Then
                outrec = outrec + "U"
            Else
            If rs2("religiousorganization2") Then
                outrec = outrec + "R"
            Else
            If rs2("society2") Then
                outrec = outrec + "S"
            End If
            End If
            End If
            End If
            End If
            End If
            End If
            End If
            'Age of Victim
            If Len(rs2("age2")) = 2 Then
                outrec = outrec + rs2("age2") + "  "
            Else
            If Len(rs2("age2")) = 1 Then
                outrec = outrec + Format$(Val(rs2("age2")), "00") + "  "
            Else
            If Len(rs2("age2")) = 4 Then
                outrec = outrec + rs2("age")
            Else
                outrec = outrec + "    "
            End If
            End If
            End If
            'Sex of Victim
            outrec = outrec + rs2("sex2") + Space$(1 - Len(rs2("sex2")))
            'Race of Victim
            outrec = outrec + rs2("race2") + Space$(1 - Len(rs2("race2")))
            'Ethnicity of Victim
            outrec = outrec + rs2("ethnicity2") + Space$(1 - Len(rs2("ethnicity2")))
            'Resident Status of Victim
            outrec = outrec + rs2("resident2") + Space$(1 - Len(rs2("resident2")))
            'Aggravated Assault/Homocide Circumstances 1
            outrec = outrec + h1$
            'Aggravated Assault/Homocide Circumstances 2
            outrec = outrec + H2$
            'Additional Justifiable Homocide Circumstances
            outrec = outrec + AD$
            'Type Injury 1 - 5
            For bb% = 1 To 5
                If Not IsNull(rs2("typeofinjury2" + Mid$(Str$(bb%), 2))) And rs2("typeofinjury2" + Mid$(Str$(bb%), 2)) > "" Then
                    Call checkinjury
                    If foundinjtype Then
                        outrec = outrec + rs2("typeofinjury2" + Mid$(Str$(bb%), 2))
                    Else
                        outrec = outrec + " "
                    End If
                Else
                    outrec = outrec + " "
                End If
            Next bb%
            'Relationship(s) of Victim to Offender(s)
            needrelationship% = 0
            For cc% = 1 To 10
                If Not IsNull(rs3("vucr2" + Mid$(Str$(cc%), 2))) Then
                    uc$ = rs3("vucr2" + Mid$(Str$(cc%), 2))
                    Select Case uc$
                        Case "13A", "13B", "13C", "09A", "09B", "09C", "100", "11A", "11B", "11C", "11D", "36A", "36B", "120"
                            needrelationship% = 1
                            cc% = 10
                    End Select
                End If

            Next cc%
            totrel% = 0
            If needrelationship% = 1 Then
                For cc% = 1 To 3
                    If Not IsNull(rs2("relationship2" + Mid$(Str$(cc%), 2))) Then
                        If cc% = 1 Then
                            Set rs4 = db.OpenRecordset("select * from incidentreports where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
                            If Not rs4.EOF Then
                                'rs4.MoveFirst
                                If UCase(rs4("sname")) = "UNKNOWN" Then
                                    If rs4("sage") = "00" And rs4("ssex") = "U" And rs4("srace") = "U" Then
                                        outrec = outrec + "00  "
                                    Else
                                        outrec = outrec + Format$(cc%, "00")
                                        outrec = outrec + rs2("relationship2" + Mid$(Str$(cc%), 2))
                                    End If
                                Else
                                    outrec = outrec + Format$(cc%, "00")
                                    outrec = outrec + rs2("relationship2" + Mid$(Str$(cc%), 2))
                                End If
                            Else
                                outrec = outrec + Format$(cc%, "00")
                                outrec = outrec + rs2("relationship2" + Mid$(Str$(cc%), 2))
                            End If
                        Else
                            outrec = outrec + Format$(cc%, "00")
                            outrec = outrec + rs2("relationship2" + Mid$(Str$(cc%), 2))
                        End If
                        totrel% = totrel% + 1
                    Else
                        outrec = outrec + "    "
                    End If
                Next cc%
                For cc% = 4 To 10
                    If Not IsNull(rs3("relationship2" + Mid$(Str$(cc%), 2))) Then
                        outrec = outrec + Format$(cc%, "00")
                        outrec = outrec + rs3("relationship2" + Mid$(Str$(cc%), 2))
                        totrel% = totrel% + 1
                    Else
                        outrec = outrec + "    "
                    End If
                Next cc%
            Else
                outrec = outrel + Space$(40)
            End If
            If totrel% = 0 Then
                totrel% = 1
            End If
            ' SC Enhancements
            ' Leoka Activity
            If Not IsNull(rs("lactivity")) Then
                outrec = outrec + rs("lactivity")
            Else
                outrec = outrec + " "
            End If
            ' Leoka Assignment
            If rs2("twomanvehicle2") = "X" Then
                outrec = outrec + "1"
            Else
            If rs2("onemanvehicle2") = "X" Then
                If rs2("alone2") = "X" Then
                    outrec = outrec + "2"
                Else
                    outrec = outrec + "3"
                End If
            Else
            If rs2("detective2") = "X" Then
                If rs2("alone2") = "X" Then
                    outrec = outrec + "4"
                Else
                    outrec = outrec + "%"
                End If
            Else
            If rs2("other2") = "X" Then
                If rs2("alone2") = "X" Then
                    outrec = outrec + "6"
                Else
                    outrec = outrec + "7"
                End If
            Else
                outrec = outrec + " "
            End If
            End If
            End If
            End If
            oridx = oridx + 1
            houtrec(oridx) = outrec
            VCT% = VCT% + 1
        End If
CHECKSUPPLEMENTALNEXT:
        rs2.MoveNext
    Wend
finishv:
    '==== SC Enhancements - Counter for Administrative Segment
    If Mid$(INCSEG, 2, 1) <> " " Then
        houtrec(1) = houtrec(1) + Format$(VCT%, "00")
    End If
offender:
    If Mid$(INCSEG, 6, 1) = " " Or ANYA = 0 Then
        ofct% = 0
        GoTo finisho
    End If
    ofct% = 0
    '===== Offender Segment =====
    Set rs2 = db.OpenRecordset("select * from incidentreports where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    If Not rs2.EOF Then
        'rs2.MoveFirst
    Else
        GoTo finisho
    End If
    ' Record Descriptor Word
    outrec = "    " + "5" + Mid$(INCSEG, 6, 1) + EXPMONTH + EXPYEAR + Space(4) + orinumber + Format$(rs2("incidentnumber"), "!@@@@@@@@@@@@")
    ' Offender Number
    If UCase(rs2("sname")) = "UNKNOWN" And rs2("sage") = "00" And rs2("ssex") = "U" And rs2("srace") = "U" Then
        outrec = outrec + "00       "
    Else
        outrec = outrec + "01"
        'Age of Offender
        If Len(rs2("sage")) = 2 Then
            outrec = outrec + rs2("sage") + "  "
        Else
        If Len(rs2("sage")) = 1 Then
            outrec = outrec + Format$(Val(rs2("2age")), "00") + "  "
        Else
        If Len(rs2("sage")) = 4 Then
            outrec = outrec + rs2("sage")
        Else
            outrec = outrec + "    "
        End If
        End If
        End If
        'Sex of Offender
        outrec = outrec + rs2("ssex")
        'Race of Offender
        outrec = outrec + rs2("srace")
        ' SC Enhancements
        ' Ethnicity
        outrec = outrec + rs2("sethnicity")
    End If
    oridx = oridx + 1
    houtrec(oridx) = outrec
    ofct% = ofct% + 1
    
    Set rs2 = db.OpenRecordset("select * from supplemental where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34) + " and (subject1 >0 or subject2 >0)")
    If Not rs2.EOF Then
        'rs2.MoveFirst
    Else
        GoTo finisho
    End If
    While Not rs2.EOF
        If Not IsNull(rs2("subject1")) And rs2("subject1") > 0 Then
            ' Record Descriptor Word
            outrec = "    " + "5" + Mid$(INCSEG, 6, 1) + EXPMONTH + EXPYEAR + Space(4) + orinumber + Format$(rs2("incidentnumber"), "!@@@@@@@@@@@@")
            ' Offender Number
            If (UCase(rs2("name1")) = "UNKNOWN" And rs2("age1") = "00" And rs2("sex1") = "U" And rs2("race1") = "U") Then
                outrec = outrec + "00       "
            Else
                outrec = outrec + Format$(rs2("subject1"), "00")
                'Age of Offender
                If Len(rs2("age1")) = 2 Then
                    outrec = outrec + rs2("age1") + "  "
                Else
                If Len(rs2("age1")) = 1 Then
                    outrec = outrec + Format$(Val(rs2("age1")), "00") + "  "
                Else
                If Len(rs2("age1")) = 4 Then
                    outrec = outrec + rs2("age1")
                Else
                    outrec = outrec + "    "
                End If
                End If
                End If
                'Sex of Offender
                outrec = outrec + rs2("sex1")
                'Race of Offender
                outrec = outrec + rs2("race1")
                ' SC Enhancements
                ' Ethnicity
                outrec = outrec + rs2("ethnicity1")
            End If
            oridx = oridx + 1
            houtrec(oridx) = outrec
            ofct% = ofct% + 1
        End If
        
        If Not IsNull(rs2("subject2")) And rs2("subject2") > 0 Then
            ' Record Descriptor Word
            outrec = "    " + "5" + Mid$(INCSEG, 6, 1) + EXPMONTH + EXPYEAR + Space(4) + orinumber + Format$(rs2("incidentnumber"), "!@@@@@@@@@@@@")
            ' Offender Number
            If (UCase(rs2("name2")) = "UNKNOWN" And rs2("age2") = "00" And rs2("sex2") = "U" And rs2("race2") = "U") Then
                outrec = outrec + "00       "
            Else
                outrec = outrec + Format$(rs2("subject2"), "00")
                'Age of Offender
                If Len(rs2("age2")) = 2 Then
                    outrec = outrec + rs2("age2") + "  "
                Else
                If Len(rs2("age2")) = 1 Then
                    outrec = outrec + Format$(Val(rs2("age2")), "00") + "  "
                Else
                If Len(rs2("age2")) = 4 Then
                    outrec = outrec + rs2("age2")
                Else
                    outrec = outrec + "    "
                End If
                End If
                End If
                'Sex of Offender
                outrec = outrec + rs2("sex2")
                'Race of Offender
                outrec = outrec + rs2("race2")
                ' SC Enhancements
                ' Ethnicity
                outrec = outrec + rs2("ethnicity2")
            End If
            oridx = oridx + 1
            houtrec(oridx) = outrec
            ofct% = ofct% + 1
        End If
        
        rs2.MoveNext
    Wend
    
finisho:
    '==== SC Enhancements - Counter for Administrative Segment
    If Mid$(INCSEG, 2, 1) <> " " Then
        houtrec(1) = houtrec(1) + Format$(ofct%, "00")
    End If
    
arresteea:
    act% = 0
    If Mid$(INCSEG, 7, 1) = " " Or ANYA = 0 Then
        GoTo arresteeb
    End If
    Set rs2 = db2.OpenRecordset("select * from booking where number < 100 and dateofarrest between #" + INDATE1 + "# and #" + INDATE2 + "# and incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34) + " and (exportdate is null or (exportdate is not null and dateofarrest < #" + INDATE2 + "# and lastupdate between #" + INDATE1 + "# and #" + INDATE2 + "#))")
    If Not rs2.EOF Then
        'rs2.MoveFirst
    Else
        GoTo arresteeb
    End If

    While Not rs2.EOF

        rs2.Edit
        MULTIPLE% = 0
        If Not IsNull(rs2("OTHERCASEs1")) Then
            If rs2("OTHERCASEs1") > "" Then
                MULTIPLE% = 1
            End If
        End If

        ' Record Descriptor Word
        outrec = "    " + "6" + Mid$(INCSEG, 7, 1) + EXPMONTH + EXPYEAR + Space(4) + orinumber + Format$(rs2("incidentnumber"), "!@@@@@@@@@@@@")
        ' Arrestee Number
        outrec = outrec + Format$(rs2("number"), "00")
        ' Arrest Number
        outrec = outrec + Format$(rs2("INCIDENTnumber"), "!@@@@@@@@@@@@")
        'Arrest Date
        outrec = outrec + Format$(Left$(rs2("dateofarrest"), 10), "yyyymmdd")
        'Type Of Arrest
        If rs2("onview") = 1 Then
            outrec = outrec + "O"
        Else
        If rs2("summoned") = 1 Then
            outrec = outrec + "S"
        Else
            outrec = outrec + "T"
        End If
        End If
        'Multiple Arrestee Segments
        incc$ = Left$(rs("INCIDENTNUMBER"), InStr(rs("INCIDENTNUMBER"), " ") - 1)
        Set rs3 = db2.OpenRecordset("select INCIDENTNUMBER from booking where number < 100 and dateofarrest < #" + INDATE2 + "# and (OTHERCASEs1 = " + Chr$(34) + incc$ + Chr$(34) + " OR OTHERCASEs2 = " + Chr$(34) + incc$ + Chr$(34) + " OR OTHERCASEs3 = " + Chr$(34) + incc$ + Chr$(34) + " OR OTHERCASEs4 = " + Chr$(34) + incc$ + Chr$(34) + " OR OTHERCASEs5 = " + Chr$(34) + incc$ + Chr$(34) + " OR OTHERCASEs6 = " + Chr$(34) + incc$ + Chr$(34) + " OR OTHERCASEs7 = " + Chr$(34) + incc$ + Chr$(34) + " OR OTHERCASEs8 = " + Chr$(34) + incc$ + Chr$(34) + " OR OTHERCASEs9 = " + Chr$(34) + incc$ + Chr$(34) + " OR OTHERCASEs10 = " + Chr$(34) + incc$ + Chr$(34) + ")  and schanged = 1 and (exportdate is null or (exportdate is not null and lastupdate between #" + INDATE1 + "# and #" + INDATE2 + "#))")
        If rs3.EOF Then
            If MULTIPLE% = 0 Then
                outrec = outrec + "N"
            Else
                outrec = outrec + "C"
            End If
        Else
        If MULTIPLE% = 0 Then
            outrec = outrec + "M"
        End If
        End If
        'UCR Arrest Offense Code
        outrec = outrec + Mid$(rs2("CHARGEA"), InStr(rs2("CHARGEA"), "(") + 1, 3)
        'Arrestee Was Armed With #1
        outrec = outrec + rs2("armedwith1")
        'Automatic Weapon Indicator #1
        If rs2("armedwithautomatic1") = 1 Then
            outrec = outrec + "A"
        Else
        If rs2("armedwithsemiautomatic1") = 1 Then
            outrec = outrec + "S"
        Else
            outrec = outrec + " "
        End If
        End If
        'Arrestee Was Armed With #2
        'Automatic Weapon Indicator #2
        If Not IsNull(rs2("armedwith2")) And rs2("armedwith2") > "" Then
            outrec = outrec + rs2("armedwith2")
            If rs2("armedwithautomatic2") = 1 Then
                outrec = outrec + "A"
            Else
            If rs2("armedwithsemiautomatic2") = 1 Then
                outrec = outrec + "S"
            Else
                outrec = outrec + " "
            End If
            End If
        Else
            outrec = outrec + "   "
        End If
        'Age of Arrestee
        If Len(rs2("sage")) = 2 Then
            outrec = outrec + rs2("sage") + "  "
        Else
        If Len(rs2("Sage")) = 1 Then
            outrec = outrec + Format$(Val(rs2("Sage")), "00") + "  "
        Else
        If Len(rs2("sage")) = 4 Then
            outrec = outrec + rs2("sage")
        Else
            outrec = outrec + "    "
        End If
        End If
        End If
        'Sex of Arrestee
        outrec = outrec + rs2("ssex")
        'Race of Arrestee
        outrec = outrec + rs2("srace")
        'Ethnicity of Arrestee
        outrec = outrec + rs2("sethnicity")
        'Resident Status of Victim
        outrec = outrec + rs2("sresident")
        'Disposition Under 18
        If rs2("referred") = 1 Then
            outrec = outrec + "R"
        Else
        If rs2("within") Then
            outrec = outrec + "H"
        Else
            outrec = outrec + " "
        End If
        End If
        If Mid$(outrec, 6, 1) = "W" Or Mid$(outrec, 6, 1) = "M" Then
            'Clearance Indicator
            If ecc Then
                outrec = outrec + "Y"
            Else
                outrec = outrec + "N"
            End If
            'Original UCR Codes
            For tt% = 1 To 5
                If Not IsNull(rs("ucr" + Mid$(Str$(tt%), 2))) Then
                    outrec = outrec + rs("ucr" + Mid$(Str$(tt%), 2))
                Else
                    outrec = outrec + Space$(3)
                End If
            Next tt%
            outrec = outrec + Space$(15)
        Else
            outrec = outrec + Space$(31)
        End If
        '===== SC Enhancements
        ' 2nd Offense
        If Not IsNull(rs2("CHARGEB")) Then
            If rs2("CHARGEB") <> rs2("CHARGEA") Then
                outrec = outrec + Mid$(rs2("CHARGEB"), InStr(rs2("CHARGEB"), "(") + 1, 3)
            Else
                outrec = outrec + "   "
            End If
        Else
            outrec = outrec + "   "
        End If
        ' 3rd Offense
        If Not IsNull(rs2("CHARGEC")) Then
            If rs2("CHARGEC") <> rs2("CHARGEB") And rs2("CHARGEC") <> rs2("CHARGEA") Then
                outrec = outrec + Mid$(rs2("CHARGEC"), InStr(rs2("CHARGEC"), "(") + 1, 3)
            Else
                outrec = outrec + "   "
            End If
        Else
            outrec = outrec + "   "
        End If
        ' TCA
        If Not IsNull(rs2("at")) Then
            outrec = outrec + rs2("at")
        Else
            outrec = outrec + " "
        End If
        ' Drug
        If Not IsNull(rs2("dt")) Then
            outrec = outrec + rs2("dt")
        Else
            outrec = outrec + " "
        End If
        oridx = oridx + 1
        houtrec(oridx) = outrec
        act% = act% + 1

        rs2("schanged") = 0
        rs2("lastexportdate") = rs2("exportdate")
        rs2("exportdate") = Date$
        If Mid$(outrec, 6, 1) = "W" Then
            rs2("ONCEW") = 1
        End If
        rs2.Update
        rs2.MoveNext

    Wend

arresteeb:
    If Mid$(INCSEG, 8, 1) = " " Or ANYA <> 0 Then
        GoTo FINISHUP
    End If
    Set rs2 = db2.OpenRecordset("select * from booking where bgroup = 'B' and number < 100 and dateofarrest between #" + INDATE1 + "# and #" + INDATE2 + "# and incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34) + " and (exportdate is null or (exportdate is not null and dateofarrest < #" + INDATE2 + "# and lastupdate between #" + INDATE1 + "# and #" + INDATE2 + "#))")
    If Not rs2.EOF Then
        'rs2.MoveFirst
    Else
        GoTo FINISHUP
    End If

    While Not rs2.EOF

        rs2.Edit

        ' Record Descriptor Word, Arrestee Number, Arrest Date
        outrec = "    " + "7" + Mid$(INCSEG, 8, 1) + EXPMONTH + EXPYEAR + Space(4) + orinumber + Format$(rs2("INCIDENTnumber"), "!@@@@@@@@@@@@") + Format$(rs2("number"), "00") + Format$(Left$(rs2("dateofarrest"), 10), "yyyymmdd")
        'Type Of Arrest
        If rs2("onview") = 1 Then
            outrec = outrec + "O"
        Else
        If rs2("summoned") = 1 Then
            outrec = outrec + "S"
        Else
            outrec = outrec + "T"
        End If
        End If
        'UCR Arrest Offense Code
        outrec = outrec + Mid$(rs2("CHARGEA"), InStr(rs2("CHARGEA"), "(") + 1, 3)
        'Arrestee Was Armed With #1
        outrec = outrec + rs2("armedwith1")
        'Automatic Weapon Indicator #1
        If rs2("armedwithautomatic1") = 1 Then
            outrec = outrec + "A"
        Else
        If rs2("armedwithsemiautomatic1") = 1 Then
            outrec = outrec + "S"
        Else
            outrec = outrec + " "
        End If
        End If
        'Arrestee Was Armed With #2
        'Automatic Weapon Indicator #2
        If Not IsNull(rs2("armedwith2")) And rs2("armedwith2") > "" Then
            outrec = outrec + rs2("armedwith2")
            If rs2("armedwithautomatic2") = 1 Then
                outrec = outrec + "A"
            Else
            If rs2("armedwithsemiautomatic2") = 1 Then
                outrec = outrec + "S"
            Else
                outrec = outrec + " "
            End If
            End If
        Else
            outrec = outrec + "   "
        End If
        'Age of Arrestee
        If Len(rs2("sage")) = 2 Then
            outrec = outrec + rs2("sage") + "  "
        Else
        If Len(rs2("Sage")) = 1 Then
            outrec = outrec + Format$(Val(rs2("Sage")), "00") + "  "
        Else
        If Len(rs2("sage")) = 4 Then
            outrec = outrec + rs2("sage")
        Else
            outrec = outrec + "    "
        End If
        End If
        End If
        'Sex of Arrestee
        outrec = outrec + rs2("ssex")
        'Race of Arrestee
        outrec = outrec + rs2("srace")
        'Ethnicity of Arrestee
        outrec = outrec + rs2("sethnicity")
        'Resident Status of Victim
        outrec = outrec + rs2("sresident")
        'Disposition Under 18
        If rs2("referred") = 1 Then
            outrec = outrec + "R"
        Else
        If rs2("within") Then
            outrec = outrec + "H"
        Else
            outrec = outrec + " "
        End If
        End If
        '===== SC Enhancements
        '===== SC Enhancements
        ' 2nd Offense
        If Not IsNull(rs2("CHARGEB")) Then
            If rs2("CHARGEB") <> rs2("CHARGEA") Then
                outrec = outrec + Mid$(rs2("CHARGEB"), InStr(rs2("CHARGEB"), "(") + 1, 3)
            Else
                outrec = outrec + "   "
            End If
        Else
            outrec = outrec + "   "
        End If
        ' 3rd Offense
        If Not IsNull(rs2("CHARGEC")) Then
            If rs2("CHARGEC") <> rs2("CHARGEB") And rs2("CHARGEC") <> rs2("CHARGEA") Then
                outrec = outrec + Mid$(rs2("CHARGEC"), InStr(rs2("CHARGEC"), "(") + 1, 3)
            Else
                outrec = outrec + "   "
            End If
        Else
            outrec = outrec + "   "
        End If
        oridx = oridx + 1
        houtrec(oridx) = outrec
        act% = act% + 1

        rs2("schanged") = 0
        rs2("lastexportdate") = rs2("exportdate")
        rs2("exportdate") = Date$
        If Mid$(outrec, 6, 1) = "W" Then
            rs2("ONCEW") = 1
        End If
        rs2.Update
        rs2.MoveNext

    Wend
    
FINISHUP:
    '==== SC Enhancements - Counter for Administrative Segment
    If Mid$(INCSEG, 2, 1) <> " " Then
        houtrec(1) = houtrec(1) + Format$(act%, "00")
    End If
    rs.Edit
    rs("lastexportdate") = rs("exportdate")
    rs("exportdate") = Date$
    rs("schanged") = 0
    rs.Update
    Set rs2 = db.OpenRecordset("select * from supplemental where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    If Not rs2.EOF Then
        'rs2.MoveFirst
        While Not rs2.EOF
            rs2.Edit
            rs2("schanged") = 0
            rs2.Update
            rs2.MoveNext
        Wend
    End If
    For QQ% = 0 To oridx
        If houtrec(QQ%) > "" And houtrec(QQ%) <> Space(Len(houtrec(QQ%))) Then
            Print #1, houtrec(QQ%)
        End If
        houtrec(QQ%) = ""
    Next QQ%
    
WENDOUT:

Next INC%
Close

Exit Sub
End Sub

Private Sub Form_Paint()
painted = painted + 1
If painted = 1 Then
    On Error GoTo 0
    founderrors = 0
    Call GETPARAMS
    If EXPMONTH = "" Or EXPYEAR = "" Then
        Exit Sub
    End If
    Call crossref(INDATE1, INDATE2)
    If founderrors = 1 Then
        msg = MsgBox("Errors have been found in the incidents scheduled for export.  These may be accessed from the TempList screen.", 48, "Genesis Error Log")
'        db.Close
        temprevw.Show
        Screen.MousePointer = 0
        Me.Hide
        Exit Sub
    End If
    If Dir(nwi + "export" + frmLogin.orinumber + Format$(INDATE1, "mm") + Format$(INDATE1, "yyyy")) > "" Then
        msg = MsgBox("An export file for this data already exists.  Do you wish to continue with the export process?", 4, "Genesis Error Log")
        If msg <> 6 Then
            temprevw.Show
            Screen.MousePointer = 0
            Me.Hide
            Exit Sub
        End If
    End If
    If incidx > 0 Then
        pb.Max = incidx
    End If
    pb.Refresh
    pb.Value = 0
    statusl = "Exporting Data"
    statusl.Refresh
    Call export
    Me.Hide
    incident.Show
    incident.WindowState = vbMaximized
    Screen.MousePointer = 0
'    Unload Me
End If
End Sub

Private Sub crossref(INDATE1, INDATE2)
Dim db, db2 As Database, rs, rs2, rs3, rs4, rs5, rs9, rsb As Recordset, incf(999999) As Boolean, incn(999999) As String
everfound = 0
On Error GoTo oderror
od:
pb.Max = 7
pb.Refresh
statusl = "Gathering Eligible Reports"
statusl.Refresh
pb.Value = 1
pb.Refresh
Set db = OpenDatabase(nwi + "incident.mdb")
Set db2 = OpenDatabase(nwb + "booking.mdb")
'---- Incident/Supplemental Cross Reference Edits
incidx = 0

Set rs = db.OpenRecordset("select * from incidentsupport where (local = 0 and (exportdate is null or exportdate < lastupdate) and temp <> 'Y' and (lastupdate between #" + INDATE1 + "# and #" + INDATE2 + "# OR INCIDENTNUMBER IN (SELECT INCIDENTNUMBER FROM INCIDENTREPORTC WHERE DATEOFOFFENSE1 BETWEEN #" + INDATE1 + "# AND #" + INDATE2 + "#))) or exportfile = 'FLAG' order by incidentnumber")
If Not rs.EOF Then
    'rs.MoveFirst
    While Not rs.EOF
        Set rs2 = db.OpenRecordset("select incidentnumber from INCIDENTREPORTC WHERE incidentnumber = '" + rs("incidentnumber") + "' AND dateofoffense1 >= #" + rbasedate + "# and dateofoffense1 <= #" + INDATE2 + "#")
        If Not rs.EOF Then
            incidx = incidx + 1
            incn(incidx) = rs("incidentnumber") + Space$(12 - Len(rs("incidentnumber")))
            If Not IsNull(rs("exportfile")) Then
                If rs("exportfile") = "FLAG" Then
                    incf(incidx) = True
                Else
                    incf(incidx) = False
                End If
            End If
        End If
        rs.MoveNext
    Wend
End If
pb.Value = 2
pb.Refresh
On Error GoTo 0
pincidx = incidx
Set rs = db.OpenRecordset("select * from incidentsupport where local = 0 and ((daterecovered1 between #" + INDATE1 + "# and #" + INDATE2 + "# ) or (daterecovered2 between #" + INDATE1 + "# and #" + INDATE2 + "# ) or  (daterecovered3 between #" + INDATE1 + "# and #" + INDATE2 + "# ) or (daterecovered4 between #" + INDATE1 + "# and #" + INDATE2 + "# ) or (daterecovered5 between #" + INDATE1 + "# and #" + INDATE2 + "# ) or (daterecovered6 between #" + INDATE1 + "# and #" + INDATE2 + "# )) and incidentnumber in (select incidentnumber from incidentreportc where dateofoffense1 <= #" + INDATE2 + "#) order by incidentnumber")
If Not rs.EOF Then
    'rs.MoveFirst
    While Not rs.EOF
        fi% = 0
        For t% = 1 To incidx
            If incn(t%) = rs("incidentnumber") + Space$(12 - Len(rs("incidentnumber"))) Then
                fi% = 1
                t% = incidx
            End If
        Next t%
        If fi% = 0 Then
            incidx = incidx + 1
            incn(incidx) = rs("incidentnumber") + Space$(12 - Len(rs("incidentnumber")))
        End If
        rs.MoveNext
    Wend
End If
pb.Value = 3
pb.Refresh
On Error GoTo 0
pincidx = incidx
Set rs = db.OpenRecordset("select * from supplementalsupport where ((daterecovered1 between #" + INDATE1 + "# and #" + INDATE2 + "# ) or (daterecovered2 between #" + INDATE1 + "# and #" + INDATE2 + "# ) or  (daterecovered3 between #" + INDATE1 + "# and #" + INDATE2 + "# ) or (daterecovered4 between #" + INDATE1 + "# and #" + INDATE2 + "# ) or (daterecovered5 between #" + INDATE1 + "# and #" + INDATE2 + "# ) or (daterecovered6 between #" + INDATE1 + "# and #" + INDATE2 + "# )) and incidentnumber in (select incidentnumber from incidentreportc where dateofoffense1 <= #" + INDATE2 + "#) order by incidentnumber")
If Not rs.EOF Then
    'rs.MoveFirst
    While Not rs.EOF
        fi% = 0
        For t% = 1 To incidx
            If incn(t%) = rs("incidentnumber") + Space$(12 - Len(rs("incidentnumber"))) Then
                fi% = 1
                t% = incidx
            End If
        Next t%
        If fi% = 0 Then
            incidx = incidx + 1
            incn(incidx) = rs("incidentnumber") + Space$(12 - Len(rs("incidentnumber")))
        End If
        rs.MoveNext
    Wend
End If
pb.Value = 4
pb.Refresh
On Error GoTo 0
pincidx = incidx
Set rs = db.OpenRecordset("select * from incidentreporto where excleardate between #" + INDATE1 + "# and #" + INDATE2 + "# and not incidentnumber in (select incidentnumber from incidentsupport where local = 1) and incidentnumber in (select incidentnumber from incidentreportc where dateofoffense1 <= #" + INDATE2 + "#) order by incidentnumber")
If Not rs.EOF Then
    'rs.MoveFirst
    While Not rs.EOF
        fi% = 0
        For t% = 1 To incidx
            If incn(t%) = rs("incidentnumber") + Space$(12 - Len(rs("incidentnumber"))) Then
                fi% = 1
                t% = incidx
            End If
        Next t%
        If fi% = 0 Then
            incidx = incidx + 1
            incn(incidx) = rs("incidentnumber") + Space$(12 - Len(rs("incidentnumber")))
        End If
        rs.MoveNext
    Wend
End If

pb.Value = 5
pb.Refresh
t% = 1
siDX = incidx
While t% <= siDX
    Set rs2 = db.OpenRecordset("select * from incidentsupport where incidentnumber = '" + incn(t%) + "' order by incidentnumber")
    'rs2.MoveFirst
    AHERE% = 0
    If rs2("temp") <> "Y" Then
        For tt% = 1 To 5
            If Not IsNull(rs2("ucr" + Mid$(Str$(tt%), 2))) Then
                Set rs4 = db.OpenRecordset("SELECT abGROUP FROM UCR WHERE abGROUP = 'A' AND abbrev = '" + rs2("UCR" + Mid$(Str$(tt%), 2)) + "'")
                If Not rs4.EOF Then
                    AHERE% = 1
                    tt% = 5
                End If
            End If
        Next tt%
    End If
    If AHERE% = 0 Then
        For tt% = t% To siDX - 1
            incn(tt%) = incn(tt% + 1)
        Next tt%
        incn(siDX) = ""
        siDX = siDX - 1
    End If
    t% = t% + 1
Wend
incidx = siDX
        
pb.Value = 6
pb.Refresh
pincidx = incidx
Set rs = db2.OpenRecordset("select * from booking where (exportdate is null or exportdate < lastupdate) and (lastupdate between #" + INDATE1 + "# and #" + INDATE2 + "# or dateofarrest between #" + INDATE1 + "# and #" + INDATE2 + "#) and dateofarrest >= #" + rbasedate + "#  order by incidentnumber")
If Not rs.EOF Then
    'rs.MoveFirst
    While Not rs.EOF
        fi% = 0
        For t% = 1 To incidx
            If incn(t%) = rs("incidentnumber") + Space$(12 - Len(rs("incidentnumber"))) Then
                fi% = 1
                t% = incidx
            End If
        Next t%
        If fi% = 0 Then
            incidx = incidx + 1
            incn(incidx) = rs("incidentnumber") + Space$(12 - Len(rs("incidentnumber")))
        End If
        rs.MoveNext
    Wend
End If


pb.Value = 7
pb.Refresh
siDX = incidx
For t% = 1 To siDX
    NUMGR% = 1
    For tt% = 1 To siDX
        If incn(t%) > incn(tt%) Then
            NUMGR% = NUMGR% + 1
        End If
    Next tt%
    SORDER(t%) = NUMGR%
Next t%
For t% = 1 To siDX
    For tt% = 1 To siDX
        If SORDER(tt%) = t% Then
            incnumber(t%) = incn(tt%)
            incflag(t%) = incf(tt%)
            tt% = siDX
        End If
    Next tt%
Next t%

statusl = "Cross Referencing Data for Errors"
statusl.Refresh
pb.Max = incidx + 1
pb.Refresh
pb.Value = 0
pb.Refresh
For INC% = 1 To incidx
    
    
    Set rs = db.OpenRecordset("select * from incidentsupport where incidentnumber = " + Chr$(34) + incnumber(INC%) + Chr$(34))
    If Not rs.EOF Then
        'rs.MoveFirst
        rs.Edit
        pb.Value = pb.Value + 1
        pb.Refresh
        founderrors = 0
        found13ab% = 0
        found210% = 0
        found100% = 0
        found120% = 0
        FOUND11% = 0
        found09c% = 0
        found36b% = 0
        found11a% = 0
        FOUND13C% = 0
        FOUND09A% = 0
        FOUND09B% = 0
        FOUND36A% = 0
        For t% = 1 To 6
            If Not IsNull(rs("daterecovered" + CStr(t%))) Then
                If rs("daterecovered" + CStr(t%)) > CDate(INDATE2) Then
                    rs("tempreason") = "Recovery date greater than export run date."
                    founderrors = 1
                    rs("temp") = "Y"
                    rs.Update
                    GoTo nextedit
                End If
            End If
        Next t%
        Set rs3 = db.OpenRecordset("select * from incidentreporto where incidentnumber = " + Chr$(34) + incnumber(INC%) + Chr$(34))
        'rs3.MoveFirst
        If rs3("excleardate") > CDate(INDATE2) Then
            rs("tempreason") = "Exceptional clearance date greater than export run date."
            founderrors = 1
            rs("temp") = "Y"
            rs.Update
            GoTo nextedit
        End If
        '----- Error 072
        foundrecovered = False
        For p = 1 To 6
            If Not IsNull(rs3("recoveredvalue" + CStr(p))) Then
                If Val(rs3("recoveredvalue" + CStr(p))) > 0 Then
                    foundrecovered = True
                    p = 6
                End If
            End If
        Next p
        foundstolen = False
        For p = 1 To 6
            If Not IsNull(rs3("stolenvalue" + CStr(p))) Then
                If Val(rs3("stolenvalue" + CStr(p))) > 0 Then
                    foundstolen = True
                    p = 6
                End If
            End If
        Next p
        Set rs3 = db.OpenRecordset("select * from supplemental where incidentnumber = " + Chr$(34) + incnumber(INC%) + Chr$(34))
        If Not rs3.EOF Then
            'rs3.MoveFirst
            While Not rs3.EOF And Not foundrecovered
                For p = 1 To 6
                    If Not IsNull(rs3("recoveredvalue" + CStr(p))) Then
                        If Val(rs3("recoveredvalue" + CStr(p))) > 0 Then
                            foundrecovered = True
                            p = 6
                        End If
                    End If
                Next p
                rs3.MoveNext
            Wend
        End If
        Set rs3 = db.OpenRecordset("select * from supplemental where incidentnumber = " + Chr$(34) + incnumber(INC%) + Chr$(34))
        If Not rs3.EOF Then
            'rs3.MoveFirst
            While Not rs3.EOF And Not foundstolen
                For p = 1 To 6
                    If Not IsNull(rs3("stolenvalue" + CStr(p))) Then
                        If Val(rs3("stolenvalue" + CStr(p))) > 0 Then
                            foundstolen = True
                            p = 6
                        End If
                    End If
                Next p
                rs3.MoveNext
            Wend
        End If
        If foundrecovered And Not foundstolen Then
            rs("tempreason") = "Recovered property found without stolen entry data."
            founderrors = 1
            rs("temp") = "Y"
            rs.Update
            GoTo nextedit
        End If
        
        Set rs3 = db.OpenRecordset("select * from supplementalsupport where incidentnumber = " + Chr$(34) + incnumber(INC%) + Chr$(34))
        If Not rs3.EOF Then
            'rs3.MoveFirst
            For t% = 1 To 6
                If Not IsNull(rs3("daterecovered" + CStr(t%))) Then
                    If rs3("daterecovered" + CStr(t%)) > CDate(INDATE2) Then
                        rs("tempreason") = "Recovery date greater than export run date."
                        founderrors = 1
                        rs("temp") = "Y"
                        rs.Update
                        GoTo nextedit
                    End If
                End If
            Next t%
        End If
        Set rs2 = db.OpenRecordset("select * from incidentreporto where incidentnumber = " + Chr$(34) + incn(t%) + Chr$(34))
        If Not rs2.EOF Then
            'rs2.MoveFirst
            If IsDate(rs2("statuschange")) Then
                If rs2("statuschange") > CVDate(INDATE2) Then
                    rs("tempreason") = "Status change date is not within date boundaries for this export."
                    founderrors = 1
                    rs("temp") = "Y"
                    rs.Update
                    GoTo nextedit
                End If
            End If
        End If
        For t% = 1 To 10
            If Not IsNull(rs("ucr" + Mid$(Str$(t%), 2))) Then
                Select Case rs("ucr" + Mid$(Str$(t%), 2))
                    Case "13C"
                        FOUND13C% = 1
                    Case "09A"
                        FOUND09A% = 1
                    Case "09B"
                        FOUND09B% = 1
                    Case "36A"
                        FOUND36A% = 1
                    Case "13A", "13B"
                        found13ab% = 1
                    Case "09C"
                        found09c% = 1
                    Case "11A"
                        found11a% = 1
                        FOUND11% = 1
                    Case "36B"
                        found36b% = 1
                    Case "210"
                        found210% = 1
                    Case "100"
                        found100% = 1
                    Case "120"
                        found120% = 1
                    Case "11B", "11C", "11D"
                        FOUND11% = 1
                End Select
            End If
        Next t%
        FOUNDINDIV% = 0
        Set rs2 = db.OpenRecordset("select individual from incidentreportc where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
        If Not rs2.EOF Then
            'rs2.MoveFirst
            If rs2("individual") = True Then
                FOUNDINDIV% = 1
            End If
        End If
        If found13ab% = 1 Or FOUND11% = 1 Or found100% = 1 Or (found210% = 1 And FOUNDINDIV% = 1) Or (found120% = 1 And FOUNDINDIV% = 1) Then
            Set rs2 = db.OpenRecordset("select typeofinjury1 from incidentreportv where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
            If Not rs2.EOF Then
                'rs2.MoveFirst
                If IsNull(rs2("typeofinjury1")) Or rs2("typeofinjury1") = "" Then
                    foundinj% = 0
                    Set rs2 = db.OpenRecordset("select  * from supplemental where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
                    If Not rs2.EOF Then
                        'rs2.MoveFirst
                        While Not rs2.EOF And foundinj% = 0
                            If Not IsNull(rs2("victim1")) Then
                                If rs2("victim1") > 0 Then
                                    If (Not IsNull(rs2("typeofinjury11")) And rs2("typeofinjury11") > "") Or _
                                       (Not IsNull(rs2("typeofinjury12")) And rs2("typeofinjury12") > "") Or _
                                       (Not IsNull(rs2("typeofinjury13")) And rs2("typeofinjury13") > "") Or _
                                       (Not IsNull(rs2("typeofinjury14")) And rs2("typeofinjury14") > "") Or _
                                       (Not IsNull(rs2("typeofinjury15")) And rs2("typeofinjury15") > "") Then
                                        foundinj% = 1
                                    End If
                                End If
                            End If
                            If Not IsNull(rs2("victim2")) Then
                                If rs2("victim2") > 0 Then
                                    If (Not IsNull(rs2("typeofinjury21")) And rs2("typeofinjury21") > "") Or _
                                       (Not IsNull(rs2("typeofinjury22")) And rs2("typeofinjury22") > "") Or _
                                       (Not IsNull(rs2("typeofinjury23")) And rs2("typeofinjury23") > "") Or _
                                       (Not IsNull(rs2("typeofinjury24")) And rs2("typeofinjury24") > "") Or _
                                       (Not IsNull(rs2("typeofinjury25")) And rs2("typeofinjury25") > "") Then
                                        foundinj% = 1
                                    End If
                                End If
                            End If
                            rs2.MoveNext
                        Wend
                    End If
                Else
                    foundinj% = 1
                End If
            End If
            If foundinj% = 0 Then
                rs("tempreason") = "An injury type must be selected for this UCR."
                founderrors = 1
                rs("temp") = "Y"
                rs.Update
                GoTo nextedit
            End If
        End If

        '---SC Edit
        vidx = 0
        siDX = 0
        maxv = 1
        maxs = 0
        numse% = 0
        seq00% = 0
        not00% = 0
        vf% = 0
        mv% = 0
        Fs% = 0
        MS% = 0
        Set rs2 = db.OpenRecordset("select vrace,VSEX,vage, vrelationship1, vrelationship2, vrelationship3 from incidentreportv where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
        If Not rs2.EOF Then
            'rs2.MoveFirst
            victimage(1) = rs2("vage")
            victimsex(1) = rs2("VSEX")
            Select Case rs2("vsex")
                Case "F"
                    vf% = 1
                Case "M"
                    mv% = 1
            End Select
            victimrace(1) = rs2("vrace")
            If FOUNDINDIV% = 1 Then
                VICTIMTYPE(1) = "I"
            Else
                VICTIMTYPE(1) = ""
            End If
            For tt% = 1 To 3
                If Not IsNull(rs2("vrelationship" + Mid$(Str$(tt%), 2))) Then
                    victimrel(1, tt%) = rs2("vrelationship" + Mid$(Str$(tt%), 2))
                Else
                    victimrel(1, tt%) = ""
                End If
            Next tt%
        End If
        Set rs2 = db.OpenRecordset("select sbirthdate,sname,srace,SSEX,sage from incidentreports where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34) + " and (sname <> 'UNKNOWN' or (sname = 'UNKNOWN' and (ssex is not null or srace is not null or sethnicity is not null or sage is not null)))")
        If Not rs2.EOF Then
            'rs2.MoveFirst
            subjectage(1) = Val(rs2("sage"))
            SUBJECTSEX(1) = rs2("SSEX")
            subjectname(1) = rs2("sname")
            Select Case rs2("ssex")
                Case "F"
                    Fs% = 1
                Case "M"
                    MS% = 1
            End Select
            subjectrace(1) = rs2("srace")
            If UCase(rs2("sname")) = "UNKNOWN" And rs2("sage") = "00" And rs2("ssex") = "U" And rs2("srace") = "U" Then
                seq00% = 1
            Else
                not00% = 1
            End If
            maxs = 1
        End If
        Set rs2 = db.OpenRecordset("select vrelationship4, vrelationship5, vrelationship6, vrelationship7, vrelationship8, vrelationship9, vrelationship10 from incidentsupport where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
        If Not rs2.EOF Then
            'rs2.MoveFirst
            For tt% = 4 To 10
                If Not IsNull(rs2("vrelationship" + Mid$(Str$(tt%), 2))) Then
                    victimrel(1, tt%) = rs2("vrelationship" + Mid$(Str$(tt%), 2))
                Else
                    victimrel(1, tt%) = ""
                End If
            Next tt%
        End If
        Set rs2 = db.OpenRecordset("select birthdate1,birthdate2,name1, name2, INDIVIDUAL1, POLICEOFFICER1,INDIVIDUAL2,POLICEOFFICER2, name1,name2,race1,race2,SEX1,SEX2,page, subject1, subject2, victim1, victim2, age1, age2, relationship11, relationship12, relationship13, relationship21, relationship22, relationship23 from supplemental where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
        If Not rs2.EOF Then
            'rs2.MoveFirst
            Set rs3 = db.OpenRecordset("select relationship14, relationship15, relationship16, relationship17, relationship18, relationship19,relationship110,relationship24, relationship25, relationship26, relationship27, relationship28, relationship29,relationship210 from supplementalsupport where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34) + " and page = " + Str$(rs2("page")))
            While Not rs2.EOF
                If Not IsNull(rs2("victim1")) And rs2("victim1") > 0 Then
                    vidx = rs2("victim1")
                    If vidx > maxv Then
                        maxv = vidx
                    End If
                    victimage(vidx) = rs2("age1")
                    victimsex(vidx) = rs2("SEX1")
                    Select Case rs2("sex1")
                        Case "F"
                            vf% = 1
                        Case "M"
                            mv% = 1
                    End Select
                    victimrace(vidx) = rs2("race1")
                    If rs2("INDIVIDUAL1") Or rs2("POLICEOFFICER1") Then
                        VICTIMTYPE(vidx) = "I"
                    Else
                        VICTIMTYPE(vidx) = ""
                    End If
                    For tt% = 1 To 3
                        If Not IsNull(rs2("relationship1" + Mid$(Str$(tt%), 2))) Then
                            victimrel(vidx, tt%) = rs2("relationship1" + Mid$(Str$(tt%), 2))
                        Else
                            victimrel(vidx, tt%) = ""
                        End If
                    Next tt%
                    If Not rs3.EOF Then
                        For tt% = 4 To 10
                            If Not IsNull(rs3("relationship1" + Mid$(Str$(tt%), 2))) Then
                                victimrel(vidx, tt%) = rs3("relationship1" + Mid$(Str$(tt%), 2))
                            Else
                                victimrel(vidx, tt%) = ""
                            End If
                        Next tt%
                    End If
                End If
                If Not IsNull(rs2("victim2")) And rs2("victim2") > 0 Then
                    vidx = rs2("victim2")
                    If vidx > maxv Then
                        maxv = vidx
                    End If
                    victimage(vidx) = rs2("age2")
                    victimsex(vidx) = rs2("SEX2")
                    Select Case rs2("sex1")
                        Case "F"
                            vf% = 1
                        Case "M"
                            mv% = 1
                    End Select
                    victimrace(vidx) = rs2("race2")
                    If rs2("INDIVIDUAL2") Or rs2("POLICEOFFICER2") Then
                        VICTIMTYPE(vidx) = "I"
                    Else
                        VICTIMTYPE(vidx) = ""
                    End If
                    For tt% = 1 To 3
                        If Not IsNull(rs2("relationship2" + Mid$(Str$(tt%), 2))) Then
                            victimrel(vidx, tt%) = rs2("relationship2" + Mid$(Str$(tt%), 2))
                        Else
                            victimrel(vidx, tt%) = ""
                        End If
                    Next tt%
                    If Not rs3.EOF Then
                        For tt% = 4 To 10
                            If Not IsNull(rs3("relationship2" + Mid$(Str$(tt%), 2))) Then
                                victimrel(vidx, tt%) = rs3("relationship2" + Mid$(Str$(tt%), 2))
                            Else
                                victimrel(vidx, tt%) = ""
                            End If
                        Next tt%
                    End If
                End If
                If Not IsNull(rs2("subject1")) And rs2("subject1") > 0 Then
                    siDX = rs2("subject1")
                    subjectage(siDX) = Val(rs2("age1"))
                    SUBJECTSEX(siDX) = rs2("SEX1")
                    subjectname(siDX) = rs2("name1")
                    Select Case rs2("sex1")
                        Case "F"
                            Fs% = 1
                        Case "M"
                            MS% = 1
                    End Select
                    subjectrace(siDX) = rs2("race1")
                    If UCase(rs2("name1")) = "UNKNOWN" And rs2("age1") = "00" And rs2("sex1") = "U" And rs2("race1") = "U" Then
                        seq00% = 1
                    Else
                        not00% = 1
                    End If
                End If
                If Not IsNull(rs2("subject2")) And rs2("subject2") > 0 Then
                    siDX = rs2("subject2")
                    subjectage(siDX) = Val(rs2("age2"))
                    SUBJECTSEX(siDX) = rs2("SEX2")
                    subjectname(siDX) = rs2("name2")
                    Select Case rs2("sex2")
                        Case "F"
                            FF2% = 1
                        Case "M"
                            m2% = 1
                    End Select
                    subjectrace(siDX) = rs2("race2")
                    If UCase(rs2("name2")) = "UNKNOWN" And rs2("age2") = "00" And rs2("sex2") = "U" And rs2("race2") = "U" Then
                        seq00% = 1
                    Else
                        not00% = 1
                    End If
                End If
                If siDX > maxs Then
                    maxs = siDX
                End If
                rs2.MoveNext
            Wend
        End If
        '===== Error 485
        For yy% = 1 To maxv
            For tt% = 1 To 10
                If victimrel(yy%, tt%) > "" Then
                    If tt% > maxs Then
                        rs("tempreason") = "A relationship is stated for a subject not entered."
                        founderrors = 1
                        rs("temp") = "Y"
                        rs.Update
                        GoTo nextedit
                        tt% = 10
                    End If
                End If
            Next tt%
            For tt% = 1 To 5
                Select Case maxv
                    Case 1
                        If Not IsNull(rs("Vucr" + CStr(yy%) + Mid$(Str$(tt%), 2))) Then
                            Select Case rs("Vucr" + CStr(yy%) + Mid$(Str$(tt%), 2))
                                Case "13A", "13B", "13C", "09A", "09B", "09C", "100", "11A", "11B", "11C", "11D", "36A", "36B", "120"
                                For yyy% = 1 To maxs
                                    If victimrel(yy%, yyy%) = "" And VICTIMTYPE(yy%) = "I" Then
                                        rs("tempreason") = "A victim must have a relationship to all subjects for Crimes Against Person and Robbery."
                                        founderrors = 1
                                        rs("temp") = "Y"
                                        rs.Update
                                        GoTo nextedit
                                        yyy% = maxs
                                        yy% = maxv
                                    End If
                                Next yyy%
                            End Select
                        End If
                    Case Else
                        If Not IsNull(rs("Vucr1" + Mid$(Str$(tt%), 2))) Then
                            Select Case rs("Vucr1" + Mid$(Str$(tt%), 2))
                                Case "13A", "13B", "13C", "09A", "09B", "09C", "100", "11A", "11B", "11C", "11D", "36A", "36B", "120"
                                For yyy% = 1 To maxs
                                    If victimrel(yy%, yyy%) = "" And VICTIMTYPE(yy%) = "I" Then
                                        rs("tempreason") = "A victim must have a relationship to all subjects for Crimes Against Person and Robbery."
                                        founderrors = 1
                                        rs("temp") = "Y"
                                        rs.Update
                                        GoTo nextedit
                                        yyy% = maxs
                                        yy% = maxv
                                    End If
                                Next yyy%
                            End Select
                        End If
                        Set rs3 = db.OpenRecordset("select vucr11, vucr12, vucr13, vucr14, vucr15, vucr21, vucr22, vucr23, vucr24, vucr25 from supplementalsupport where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
                        If Not rs3.EOF Then
                            'rs3.MoveFirst
                            While Not rs3.EOF
                                If Not IsNull(rs3("Vucr1" + Mid$(Str$(tt%), 2))) Then
                                    Select Case rs3("Vucr1" + Mid$(Str$(tt%), 2))
                                        Case "13A", "13B", "13C", "09A", "09B", "09C", "100", "11A", "11B", "11C", "11D", "36A", "36B", "120"
                                        For yyy% = 1 To maxs
                                            If victimrel(yy%, yyy%) = "" And VICTIMTYPE(yy%) = "I" Then
                                                rs("tempreason") = "A victim must have a relationship to all subjects for Crimes Against Person and Robbery."
                                                founderrors = 1
                                                rs("temp") = "Y"
                                                rs.Update
                                                GoTo nextedit
                                                yyy% = maxs
                                                yy% = maxv
                                            End If
                                        Next yyy%
                                    End Select
                                End If
                                If Not IsNull(rs3("Vucr2" + Mid$(Str$(tt%), 2))) Then
                                    Select Case rs3("Vucr2" + Mid$(Str$(tt%), 2))
                                        Case "13A", "13B", "13C", "09A", "09B", "09C", "100", "11A", "11B", "11C", "11D", "36A", "36B", "120"
                                        For yyy% = 1 To maxs
                                            If victimrel(yy%, yyy%) = "" And VICTIMTYPE(yy%) = "I" Then
                                                rs("tempreason") = "A victim must have a relationship to all subjects for Crimes Against Person and Robbery."
                                                founderrors = 1
                                                rs("temp") = "Y"
                                                rs.Update
                                                GoTo nextedit
                                                yyy% = maxs
                                                yy% = maxv
                                            End If
                                        Next yyy%
                                    End Select
                                End If
                                rs3.MoveNext
                            Wend
                        End If
                End Select
            Next tt%
        Next yy%
        '===== Every Offense Must Have A Victim Connect To It
        For ttt% = 1 To 10
            If Not IsNull(rs("ucr" + Mid$(Str$(ttt%), 2))) Then
                fc% = 0
                For tt% = 1 To 5
                    If Not IsNull(rs("Vucr1" + Mid$(Str$(tt%), 2))) Then
                        If rs("Vucr1" + Mid$(Str$(tt%), 2)) = rs("ucr" + Mid$(Str$(ttt%), 2)) Then
                            fc% = 1
                            tt% = 5
                        End If
                    End If
                Next tt%
                If fc% = 0 Then
                    Set rs3 = db.OpenRecordset("select vucr11, vucr12, vucr13, vucr14, vucr15, vucr21, vucr22, vucr23, vucr24, vucr25 from supplementalsupport where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
                    If Not rs3.EOF Then
                        While Not rs3.EOF And fc% = 0
                            For tt% = 1 To 5
                                If Not IsNull(rs3("Vucr1" + Mid$(Str$(tt%), 2))) Then
                                    If rs3("Vucr1" + Mid$(Str$(tt%), 2)) = rs("ucr" + Mid$(Str$(ttt%), 2)) Then
                                        fc% = 1
                                        tt% = 5
                                    End If
                                End If
                                If Not IsNull(rs3("Vucr2" + Mid$(Str$(tt%), 2))) Then
                                    If rs3("Vucr2" + Mid$(Str$(tt%), 2)) = rs("ucr" + Mid$(Str$(ttt%), 2)) Then
                                        fc% = 1
                                        tt% = 5
                                    End If
                                End If
                            Next tt%
                            rs3.MoveNext
                        Wend
                    End If
                End If
                If fc% = 0 Then
                    rs("tempreason") = "Every offense must have a victim connected to it."
                    founderrors = 1
                    rs("temp") = "Y"
                    rs.Update
                    GoTo nextedit
                End If
            End If
        Next ttt%
        '===== Error 555
        If seq00% = 1 And not00% = 1 Then
            rs("tempreason") = "If an unknown offender is submitted, no known offenders may be entered."
            rs("temp") = "Y"
            founderrors = 1
            rs.Update
            GoTo nextedit
        End If
        '===== Error 560
        '===== SCEdit 12/19/92 P35C
        If found11a% = 1 Or found36b% = 1 Then
            If vf% <> MS% Then
                rs("tempreason") = "For Forcible or Statutory Rape, the sex of 1 or more subjects must differ from the victim(s)."
                rs("temp") = "Y"
                founderrors = 1
                rs.Update
                GoTo nextedit
            End If
        End If
        '===== Error 557,558
        If seq00% = 1 Then
            Set rs2 = db.OpenRecordset("select excleardate from inciDentreporto where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34) + " and excleardate is not null")
            If Not rs2.EOF Then
                rs("tempreason") = "Exceptional Clearance is not allowed for an unknown offender."
                rs("temp") = "Y"
                founderrors = 1
                rs.Update
                GoTo nextedit
            End If
        End If
        '===== Error 559
        If seq00% = 1 And not00% = 0 And found09c% = 1 Then
            rs("tempreason") = "For Justifiable Homocide, at least 1 offender must have known information."
            rs("temp") = "Y"
            founderrors = 1
            rs.Update
            GoTo nextedit
        End If
        '===== Error 460
'        If maxv > maxs Then
'            rs("tempreason") = "A relationship must be entered for each victim to each offender."
'            rs("temp") = "Y"
'            founderrors = 1
'            rs.Update
'            GoTo nextedit
'        End If
        '===== Error 476
        '=====SCEdit 12/19/92 P35C
        For tt% = 1 To maxs
            numse% = 0
            For ttt% = 1 To maxv
                If victimrel(ttt%, tt%) = "SE" Then
                    numse% = numse% + 1
                End If
            Next ttt%
            If numse% > 1 Then
                rs("tempreason") = "2 or more victims are related to the same offender as Spouse."
                rs("temp") = "Y"
                founderrors = 1
                rs.Update
                GoTo nextedit
            End If
        Next tt%
        For tt% = 1 To maxv
            numse% = 0
            For ttt% = 1 To 10
                If victimrel(tt%, ttt%) > "" Then
                    temprel = victimrel(tt%, ttt%)
                    If temprel = "SE" Then
                        numse% = numse% + 1
                    End If
                    Select Case temprel
                        '===== Error 472
                        Case "RU"
                        Case Else
                            If subjectage(ttt%) = "00" And SUBJECTSEX(ttt%) = "U" And subjectrace(ttt%) = "U" And subjectname(tt%) = "UNKNOWN" Then
                            Else
                            If subjectname(tt%) = "UNKNOWN" Then
                                rs("tempreason") = "Relationship must be RU if offender sex, age, and race are unknown."
                                rs("temp") = "Y"
                                founderrors = 1
                                rs.Update
                                GoTo nextedit
                            End If
                            End If
                    End Select
                    Select Case temprel
                        Case "SE"
                            '===== Error 450,550
                            If subjectage(ttt%) < 10 Or victimage(tt%) < 10 Then
                                rs("tempreason") = "If relationship is SE=Spouse, both victim and offender must be at least 10 years old."
                                rs("temp") = "Y"
                                founderrors = 1
                                rs.Update
                                GoTo nextedit
                            End If
                            If victimsex(tt%) = SUBJECTSEX(ttt%) Then
                                rs("tempreason") = "For relationships of BG, XS, SE, and CS, sexes have to be different."
                                rs("temp") = "Y"
                                founderrors = 1
                                rs.Update
                                GoTo nextedit
                            End If
                        Case "BG", "XS", "SE", "CS"
                            '===== Error 553
                            If victimsex(tt%) = SUBJECTSEX(ttt%) Then
                                rs("tempreason") = "For relationships of BG, XS, SE, and CS, sexes have to be different."
                                rs("temp") = "Y"
                                founderrors = 1
                                rs.Update
                                GoTo nextedit
                            End If
                        Case "HR"
                            '===== Error 553
                            If victimsex(tt%) <> SUBJECTSEX(ttt%) Then
                                rs("tempreason") = "For relationship of HR, sexes have to be the same."
                                rs("temp") = "Y"
                                founderrors = 1
                                rs.Update
                                GoTo nextedit
                            End If
                    End Select
                    '===== Error 554
                    If temprel = "CH" Or temprel = "GC" Or temprel = "SC" Then
                        If subjectage(ttt%) < victimage(tt%) Then
                            rs("tempreason") = "The relationship of victim to subject cannot be 'PA', 'GP' or 'SP' when victim's age is less than subject's age."
                            rs("temp") = "Y"
                            founderrors = 1
                            rs.Update
                            GoTo nextedit
                        End If
                    End If
                    '===== Error 554
                    If temprel = "PA" Or temprel = "GP" Or temprel = "SP" Then
                        If victimage(tt%) < subjectage(ttt%) Then
                            rs("tempreason") = "The relationship of victim to subject cannot be 'PA', 'GP' or 'SP' when victim's age is less than subject's age."
                            rs("temp") = "Y"
                            founderrors = 1
                            rs.Update
                            GoTo nextedit
                        End If
                    End If
                    If temprel = "VO" Then
                        If victimage(tt%) <> subjectage(ttt%) Then
                            rs("tempreason") = "The relationship of victim to subject cannot be 'VO' when victim's age is not the subject's age."
                            rs("temp") = "Y"
                            founderrors = 1
                            rs.Update
                            GoTo nextedit
                        End If
                    End If
                End If
            Next ttt%
        '===== Error 475
        If numse% > 1 Then
            rs("tempreason") = "Victim cannot be related to more than 1 offender as Spouse (SE)."
            rs("temp") = "Y"
            founderrors = 1
            rs.Update
            GoTo nextedit
        End If
        '=====
        Next tt%
    End If
nextedit:
    'GoTo nexti
    POPMSG$ = ""
    If founderrors = 0 Then
        incident.Picture2.Picture = LoadPicture()
        incident.incidentnumber = incnumber(INC%)
        incident.Hide
        Call incident.clearroutine(1)
        Call incident.findincident(1)
        editerr% = 0
        POPMSG$ = ""
        Call incident.editevent(editerr%, POPMSG$)
        If editerr% = 0 Then
            Call incident.editvictim(editerr%, POPMSG$)
            If editerr% = 0 Then
                Call incident.editsubject(editerr%, POPMSG$)
                If editerr% = 0 Then
                    Call incident.editadministrative(editerr%, POPMSG$)
                    If editerr% = 0 Then
                        Call incident.editproperty(editerr%, POPMSG$)
                    End If
                End If
            End If
        End If
        incidate$ = incident.incidentdate(0)
        Unload incident
        If editerr% = 0 Then
            pg% = 1
            Set rsb = db.OpenRecordset("select incidentnumber from supplemental where page = " + CStr(pg) + " and incidentnumber = '" + incnumber(INC%) + "'")
            If Not rsb.EOF Then
                incidentfound = True
            Else
                incidentfound = False
            End If
            While incidentfound
                
                Open "NP.TAG" For Output As #1
                Print #1, incnumber(INC%)
                Print #1, pg%
                Print #1, incidate$
                Close #1
                
                sinciden.Picture2.Picture = LoadPicture()
                sinciden.incidentnumber = incnumber(INC%)
                sinciden.Hide
                incidentfound = False
                sinciden.PAGE = pg%
                Call sinciden.clearroutine(1)
                Call sinciden.findincident(incidentfound)
                If incidentfound Then
                    editerr% = 0
                    POPMSG$ = ""
                    Call sinciden.editvictim(editerr%, POPMSG$)
                    If editerr% = 0 Then
                        Call sinciden.editsubject(editerr%, POPMSG$)
                        If editerr% = 0 Then
                            Call sinciden.editproperty(editerr%, POPMSG$)
                        End If
                    End If
                End If
                pg% = pg% + 1
                Set rsb = db.OpenRecordset("select incidentnumber from supplemental where page = " + CStr(pg) + " and incidentnumber = '" + incnumber(INC%) + "'")
                If Not rsb.EOF Then
                    incidentfound = True
                Else
                    incidentfound = False
                End If
            Wend
            Unload sinciden
        Else
            founderrors = 1
        End If
        
        If founderrors = 0 Then
            Set rsb = db2.OpenRecordset("select * from booking where incidentnumber = '" + incnumber(INC%) + "'")
            If Not rsb.EOF Then
                'rsb.MoveFirst
            End If
            While Not rsb.EOF
            
                booking.Picture2.Picture = LoadPicture()
                booking.WindowState = vbMinimized
                booking.incidentnumber = incn(INC%)
                booking.Hide
                Call booking.clearroutine
                booking.incidentnumber = incnumber(INC%)
                booking.subjectnumber = rsb("NUMBER")
                Call booking.findincident
                editerr% = 0
                POPMSG$ = ""
                Call booking.editroutine(editerr%, POPMSG$)
                If editerr% <> 0 Then
                    founderrors = 1
                    rsb.MoveLast
                End If
                rsb.MoveNext
                
            Wend
            Unload booking
        End If
        
        If editerr% <> 0 Then
            Set rs = db.OpenRecordset("select * from incidentsupport where incidentnumber = " + Chr$(34) + incnumber(INC%) + Chr$(34))
            If Not rs.EOF Then
                'rs.MoveFirst
                rs.Edit
                rs("tempreason") = Left$(POPMSG$, 100)
                rs("temp") = "Y"
                founderrors = 1
                everfound = 1
                rs.Update
            End If
        End If
    Else
        founderrors = 1
        everfound = 1
    End If
nexti:
Next INC%
Unload incident
founderrors = everfound
If founderrors = 0 Then
    For i% = 1 To INC%
        If incf(i%) Then
            Set rs = db.OpenRecordset("SELECT EXPORTFILE FROM INCIDENTSUPPORT WHERE INCIDENTNUMBER = '" + incn(i%) + "'")
            If Not rs.EOF Then
                rs.Edit
                rs("EXPORTFILE") = Date$ + "FLAG"
                rs.Update
            End If
        End If
    Next i%
End If
Exit Sub
oderror:
If Err > 3200 Then
    Resume
Else
    Resume
    Resume
End If

End Sub

Private Sub GETDATES()
EXPMONTH = ""
EXPYEAR = ""
If Val(Left$(Date$, 2)) = 1 Then
    inp = InputBox("Enter numeric month for export.", "Genesis Information Log", "12")
Else
    inp = InputBox("Enter numeric month for export.", "Genesis Information Log", Mid$(Str$(Val(Left$(Date$, 2)) - 1), 2))
End If
If Val(inp) < 1 Or Val(inp) > 12 Then
    msg = MsgBox("Invalid month entry.", 48, "Genesis Error Log")
    Exit Sub
End If
EXPMONTH = Format$(Val(inp), "00")
If Val(Left$(Date$, 2)) = 1 Then
    inp = InputBox("Enter numeric year for export.", "Genesis Information Log", Format$(Val(Right$(Date$, 4)) - 1, "0000"))
Else
    inp = InputBox("Enter numeric year for export.", "Genesis Information Log", Right$(Date$, 4))
End If
curyear% = Val(Right$(Date$, 4))
If Val(inp) < curyear% - 1 Or Val(inp) > curyear% Then
    msg = MsgBox("Invalid year entry.", 48, "Genesis Error Log")
    Exit Sub
End If
EXPYEAR = Format$(Val(inp), "0000")
INDATE1 = EXPMONTH + "/01/" + EXPYEAR
Select Case Val(EXPMONTH)
    Case 1, 3, 5, 7, 8, 10, 12
        INDATE2 = EXPMONTH + "/31/" + EXPYEAR
    Case 4, 6, 9, 11
        INDATE2 = EXPMONTH + "/30/" + EXPYEAR
    Case 2
        ly1% = Val(EXPYEAR) / 4
        ly2! = Val(EXPYEAR) / 4
        If ly1% = ly2! Then
            INDATE2 = EXPMONTH + "/29/" + EXPYEAR
        Else
            INDATE2 = EXPMONTH + "/28/" + EXPYEAR
        End If
End Select

End Sub
Private Sub GETPARAMS()
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("Select orinumber, notonpaper from system")
'rs.MoveFirst
orinumber = rs("orinumber")
NOTONPAPER = rs("notonpaper")
BASEDATE = "01/01/" + Mid$(Str$(Val(Right$(Date$, 4)) - 1), 2)
rbasedate = BASEDATE
If BASEDATE < NOTONPAPER Then
    BASEDATE = NOTONPAPER
End If
'===== Data Element 1
If orinumber = "" Or Left$(orinumber, 2) <> "SC" Or Right$(orinumber, 2) <> "00" Or Len(orinumber) <> 9 Then
    msg = MsgBox("Invalid ORI number in system.", 48, "Genesis Error Log")
    Unload Me
End If
db.Close
Call GETDATES
End Sub
Private Sub SETINCSEG(rc As Integer)
If flag Then
    abc123 = 1
End If
If PAPERON = 1 Then
    If ANYA = 1 Then
        If RECOVEREDFOUND = 1 Then
            If ecc = 1 Then
                INCSEG = " W W    "
            Else
                INCSEG = "   W    "
            End If
        Else
            If ecc = 1 Then
                INCSEG = " W      "
            Else
                GoTo WENDOUT
            End If
        End If
    End If
Else
If Not IsDate(EXPORTDATE) And CVDate(BASEDATE) < CVDate(incdate) Then
    If ANYA = 0 Then
        INCSEG = "I      A"
    Else
    If BOOKINGFOUND = 1 Then
        If foundproperty = 1 Then
            INCSEG = "IIIIIIIA"
        Else
            INCSEG = "III IIIA"
        End If
    Else
        If foundproperty = 1 Then
            INCSEG = "IIIIII  "
        Else
            INCSEG = "III II  "
        End If
    End If
    End If
Else
If CVDate(BASEDATE) <= CVDate(incdate) Then
    If ANYA = 0 Then
        INCSEG = "       A"
    Else
    If INCCHANGED = 1 Or RECOVEREDFOUND = 1 Then
        If BOOKINGFOUND = 0 Then
            If foundproperty = 1 Then
                INCSEG = "DIIIII  "
            Else
                INCSEG = "DII II  "
            End If
        Else
            If foundproperty = 1 Then
                INCSEG = "DIIIIIIA"
            Else
                INCSEG = "DII IIIA"
            End If
        End If
    Else
    If SC = 1 Or ecc = 1 Then
        If BOOKINGFOUND = 1 Then
            If BOOKINGEXPORTED = 1 Then
                If foundproperty = 1 Then
                    INCSEG = "DIIIIIIA"
                Else
                    INCSEG = "DII IIIA"
                End If
            Else
                INCSEG = " M    AA"
            End If
        Else
            INCSEG = " M       "
        End If
    Else
    If BOOKINGFOUND = 1 Then
        INCSEG = "      AA"
    Else
        If foundproperty = 1 Then
            INCSEG = "DIIIII  "
        Else
            INCSEG = "DII II  "
        End If
    End If
    End If
    End If
    End If
Else
If ANYA = 0 Then
    INCSEG = "       A"
Else
If SC = 1 Or ecc = 1 Then
    If RECOVEREDFOUND = 1 Then
        If BOOKINGFOUND = 1 Then
            If ONCEWB = 1 Then
                If ONCEWP = 1 Then
                    If ONCEW = 1 Then
                        INCSEG = " M M  MM"
                    Else
                        INCSEG = " W M  MM"
                    End If
                Else
                    If ONCEW = 1 Then
                        INCSEG = " M W  MM"
                    Else
                        INCSEG = " W W  MM"
                    End If
                End If
            Else
                If ONCEWP = 1 Then
                    If ONCEW = 1 Then
                        INCSEG = " M M  WW"
                    Else
                        INCSEG = " W M  WW"
                    End If
                Else
                    If ONCEW = 1 Then
                        INCSEG = " M W  WW"
                    Else
                        INCSEG = " W W  WW"
                    End If
                End If
            End If
        Else
            If ONCEWP = 1 Then
                If ONCEW = 1 Then
                    INCSEG = " M M    "
                Else
                    INCSEG = " W M    "
                End If
            Else
                If ONCEW = 1 Then
                    INCSEG = " M W    "
                Else
                    INCSEG = " W W    "
                End If
            End If
        End If
    Else
        If BOOKINGFOUND = 1 Then
            If ONCEWB = 1 Then
                If ONCEW = 1 Then
                    INCSEG = " M    MM"
                Else
                    INCSEG = " W    MM"
                End If
            Else
                If ONCEW = 1 Then
                    INCSEG = " M    WW"
                Else
                    INCSEG = " W    WW"
                End If
            End If
        Else
            If ONCEW = 1 Then
                INCSEG = " M      "
            Else
                INCSEG = " W      "
            End If
        End If
    End If
Else
    If RECOVEREDFOUND = 1 Then
        If BOOKINGFOUND = 1 Then
            If ONCEWB = 1 Then
                If ONCEWP = 1 Then
                    INCSEG = "   M  MM"
                Else
                    INCSEG = "   W  MM"
                End If
            Else
                If ONCEWP = 1 Then
                    INCSEG = "   M  WW"
                Else
                    INCSEG = "   W  WW"
                End If
            End If
        Else
            If ONCEWP = 1 Then
                INCSEG = "   M    "
            Else
                INCSEG = "   W    "
            End If
        End If
    Else
        If BOOKINGFOUND = 1 Then
            If ONCEWB = 1 Then
                INCSEG = "      MM"
            Else
                INCSEG = "      WW"
            End If
        Else
           INCSEG = "        "
        End If
    End If
End If
End If
End If
End If
End If
If flag Then
    If INCCHANGED = 1 Or RECOVEREDFOUND = 1 Then
        If BOOKINGFOUND = 0 Then
            If foundproperty = 1 Then
                INCSEG = "DIIIII  "
            Else
                INCSEG = "DII II  "
            End If
        Else
            If foundproperty = 1 Then
                INCSEG = "DIIIIIIA"
            Else
                INCSEG = "DII IIIA"
            End If
        End If
    Else
    If SC = 1 Or ecc = 1 Then
        If BOOKINGFOUND = 1 Then
            If BOOKINGEXPORTED = 1 Then
                If foundproperty = 1 Then
                    INCSEG = "DIIIIIIA"
                Else
                    INCSEG = "DII IIIA"
                End If
            Else
                INCSEG = " M    AA"
            End If
        Else
            INCSEG = " M       "
        End If
    Else
    If BOOKINGFOUND = 1 Then
        INCSEG = "      AA"
    Else
        If foundproperty = 1 Then
            INCSEG = "DIIIII  "
        Else
            INCSEG = "DII II  "
        End If
    End If
    End If
    End If
End If
If Mid$(INCSEG, 4, 1) = "W" Or Mid$(INCSEG, 4, 1) = "M" Then
    PROPCHECK1 = 4
    PROPCHECK2 = 4
Else
    PROPCHECK1 = 1
    PROPCHECK2 = 7
End If
Exit Sub
WENDOUT:
rc = -1
Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set iexport = Nothing
End Sub

Private Sub CLEARPROPERTY()
For tt% = 1 To 100
    For ttt% = 1 To 7
        propertyvalue(tt%, ttt%) = 0
        PROPERTYDRUGS(tt%, ttt%) = False
    Next ttt%
    propertytype(tt%) = "  "
    numvehs(tt%) = 0
    numvehr(tt%) = 0
    propertydate(tt%) = ""
    For ttt% = 1 To 3
        pdt(tt%, ttt%) = ""
        pdm(tt%, ttt%) = ""
        pdq(tt%, ttt%) = 0
    Next ttt%
Next tt%

End Sub
Friend Sub checkinjury()
foundinjtype = False
For hh% = 1 To 5
    Select Case vuc(hh%)
        Case "100", "11A", "11B", "11C", "11D", "120", "13A", "13B", "210"
            foundinjtype = True
    End Select
Next hh%
End Sub

