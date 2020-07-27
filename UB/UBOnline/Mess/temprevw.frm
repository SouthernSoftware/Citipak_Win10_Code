VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form temprevw 
   BackColor       =   &H00808000&
   Caption         =   "Review Temp Save Entries"
   ClientHeight    =   7590
   ClientLeft      =   90
   ClientTop       =   1350
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   11655
   WindowState     =   2  'Maximized
   Begin VB.Frame rightclickframe 
      Height          =   3255
      Left            =   2880
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
      Begin VB.Label Label8 
         Caption         =   "Mark X in Seized 1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   45
         TabIndex        =   9
         Top             =   2040
         Width           =   1605
      End
      Begin VB.Label Label7 
         Caption         =   "Change Relationship 1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   45
         TabIndex        =   8
         Top             =   1680
         Width           =   1605
      End
      Begin VB.Label Label6 
         Caption         =   "Re-Save Incident"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   45
         TabIndex        =   7
         Top             =   2760
         Width           =   1605
      End
      Begin VB.Label Label5 
         Caption         =   "Change from Temp"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   45
         TabIndex        =   6
         Top             =   2400
         Width           =   1605
      End
      Begin VB.Label Label4 
         Caption         =   "Access Incident"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   45
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Change Victim UCR"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   45
         TabIndex        =   4
         Top             =   1320
         Width           =   1605
      End
      Begin VB.Label Label3 
         Caption         =   "Change Property UCR"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   45
         TabIndex        =   3
         Top             =   960
         Width           =   1605
      End
      Begin VB.Label Label2 
         Caption         =   "Change UCR"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   45
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView TEMPLIST 
      Height          =   7350
      Left            =   75
      TabIndex        =   0
      Top             =   120
      Width           =   11550
      _ExtentX        =   20373
      _ExtentY        =   12965
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Type/Error"
         Object.Width           =   8467
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Case Number"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Incident Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Victim"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Subject"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "UCR 1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Group 1"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "temprevw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload temprevw
End Sub


Private Sub Command3_Click()
eeframe.Visible = False
End Sub

Private Sub Form_Load()
On Error Resume Next
Kill "*.dsk"
Dim db As Database, rs, rs2, rs3, rs4, rs5 As Recordset, itmx As ListItem
On Error GoTo oderror
od:
ct = 0
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("select tempreason,incidentnumber from INCIDENTSUPPORT where temp = 'Y' and incidentnumber is not null order by incidentnumber")
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        Set rs2 = db.OpenRecordset("SELECT dateofoffense1, CNAME from incidentreportc WHERE INCIDENTNUMBER = " + Chr$(34) + rs("INCIDENTNUMBER") + Chr$(34))
        Set rs3 = db.OpenRecordset("select VNAME from incidentreportv WHERE INCIDENTNUMBER = " + Chr$(34) + rs("INCIDENTNUMBER") + Chr$(34))
        Set rs4 = db.OpenRecordset("select SNAME FROM INCIDENTREPORTs WHERE INCIDENTNUMBER = " + Chr$(34) + rs("INCIDENTNUMBER") + Chr$(34))
        Set rs5 = db.OpenRecordset("select ucr1,group1 FROM INCIDENTsupport WHERE INCIDENTNUMBER = " + Chr$(34) + rs("INCIDENTNUMBER") + Chr$(34))
        If Not rs2.EOF Then
            rs2.MoveFirst
            rs3.MoveFirst
            rs4.MoveFirst
            rs5.MoveFirst
            ct = ct + 1
            If IsNull(rs("tempreason")) Then
                Set itmx = TEMPLIST.ListItems.add(, , "TEMP SAVE/SUPPLEMENTAL")
            Else
                Set itmx = TEMPLIST.ListItems.add(, , rs("tempreason"))
            End If
            itmx.SubItems(1) = rs("incidentnumber")
            If Not IsNull(rs2("dateofoffense1")) Then
                itmx.SubItems(2) = rs2("dateofoffense1")
            End If
            If Not IsNull(rs3("vname")) Then
                itmx.SubItems(3) = rs3("vname")
            End If
            If Not IsNull(rs4("sname")) Then
                itmx.SubItems(4) = rs4("sname")
            End If
            itmx.SubItems(5) = rs5("ucr1")
            If Not IsNull(rs5("group1")) Then
                itmx.SubItems(6) = rs5("group1")
            End If
        End If
        rs.MoveNext
    Wend
End If
Set rs = db.OpenRecordset("select tempreason,incidentnumber from supplementalsUPPORT where temp = 'Y' and incidentnumber in (select incidentnumber from supplemental) order by incidentnumber")
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        Set rs2 = db.OpenRecordset("SELECT dateofoffense1, CNAME from incidentreportc WHERE INCIDENTNUMBER = " + Chr$(34) + rs("INCIDENTNUMBER") + Chr$(34))
        Set rs3 = db.OpenRecordset("select VNAME from incidentreportv WHERE INCIDENTNUMBER = " + Chr$(34) + rs("INCIDENTNUMBER") + Chr$(34))
        Set rs4 = db.OpenRecordset("select SNAME FROM INCIDENTREPORTs WHERE INCIDENTNUMBER = " + Chr$(34) + rs("INCIDENTNUMBER") + Chr$(34))
        Set rs5 = db.OpenRecordset("select ucr1,group1 FROM INCIDENTsupport WHERE INCIDENTNUMBER = " + Chr$(34) + rs("INCIDENTNUMBER") + Chr$(34))
        If Not rs2.EOF Then
            rs2.MoveFirst
            rs3.MoveFirst
            rs4.MoveFirst
            ct = ct + 1
            If IsNull(rs("tempreason")) Then
                Set itmx = TEMPLIST.ListItems.add(, , "TEMP SAVE")
            Else
                Set itmx = TEMPLIST.ListItems.add(, , rs("tempreason"))
            End If
            itmx.SubItems(1) = rs("incidentnumber")
            If Not IsNull(rs2("dateofoffense1")) Then
                itmx.SubItems(2) = rs2("dateofoffense1")
            End If
            If Not IsNull(rs3("vname")) Then
                itmx.SubItems(3) = rs3("vname")
            End If
            If Not IsNull(rs4("sname")) Then
                itmx.SubItems(3) = rs4("sname")
            End If
            itmx.SubItems(5) = rs5("ucr1")
            If Not IsNull(rs5("group1")) Then
                itmx.SubItems(6) = rs5("group1")
            End If
        End If
        rs.MoveNext
    Wend
End If
db.Close
On Error Resume Next
Me.Left = 100
Me.Top = 500
Me.Height = 6645
Me.Width = 11775
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub

Private Sub Form_LostFocus()
goingelsewhere = False
End Sub

Private Sub Form_Paint()
SetAlwaysOnTop Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set temprevw = Nothing
goingelsewhere = False
End Sub

Private Sub Label6_Click()
Screen.MousePointer = 0
Me.Hide
Dim db As Database, rs As Recordset
Set db = OpenDatabase(nwi + "incident.mdb")
For t% = 1 To TEMPLIST.ListItems.Count
    If TEMPLIST.ListItems(t%).Selected Then
        Set itmx = TEMPLIST.ListItems(t%)
        incident.incidentnumber = itmx.SubItems(1)
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
        If editerr% = 0 Then
            Call incident.saveincident
        End If
        Unload incident
        TEMPLIST.ListItems(t%).Selected = False
    End If
    If TEMPLIST.SelectedItem Is Nothing Then
        t% = TEMPLIST.ListItems.Count
    End If
Next t%
GETOUT:
Screen.MousePointer = 0
Me.Show
rightclickframe.Visible = False
End Sub

Private Sub Label1_Click()
inp = InputBox("Enter Victim 1 UCR to change (1 - 5).", "Genesis Information Log", "1")
If Val(inp) < 1 Or Val(inp) > 5 Then
    GoTo GETOUT
End If
whichucr = Val(inp)
inp = InputBox("Enter new Victim 1 UCR value.", "Genesis Information Log", "")
If inp = "" Or Len(inp) <> 3 Then
    GoTo GETOUT
End If
Screen.MousePointer = 0
Dim db As Database, rs As Recordset
Set db = OpenDatabase(nwi + "incident.mdb")
For t% = 1 To TEMPLIST.ListItems.Count
    If TEMPLIST.ListItems(t%).Selected Then
        Set itmx = TEMPLIST.ListItems(t%)
        Set rs = db.OpenRecordset("select vucr1" + CStr(whichucr) + " from incidentsupport where incidentnumber = '" + itmx.SubItems(1) + "'")
        If Not rs.EOF Then
            rs.MoveFirst
            rs.Edit
            rs("vucr1" + CStr(whichucr)) = inp
            rs.Update
        End If
        TEMPLIST.ListItems(t%).Selected = False
    End If
    If TEMPLIST.SelectedItem Is Nothing Then
        t% = TEMPLIST.ListItems.Count
    End If
Next t%
GETOUT:
Screen.MousePointer = 0
rightclickframe.Visible = False
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BorderStyle = 0
Label2.BorderStyle = 0
Label3.BorderStyle = 0
Label4.BorderStyle = 0
Label5.BorderStyle = 0
Label6.BorderStyle = 0
Label1.BorderStyle = 1
Label7.BorderStyle = 0
Label8.BorderStyle = 0
End Sub

Private Sub label2_Click()
inp = InputBox("Enter UCR to change (1 - 5).", "Genesis Information Log", "1")
If Val(inp) < 1 Or Val(inp) > 5 Then
    GoTo GETOUT
End If
whichucr = Val(inp)
inp = InputBox("Enter new UCR value.", "Genesis Information Log", "")
If inp = "" Or Len(inp) <> 3 Then
    GoTo GETOUT
End If
Screen.MousePointer = 0
Dim db As Database, rs As Recordset
Set db = OpenDatabase(nwi + "incident.mdb")
For t% = 1 To TEMPLIST.ListItems.Count
    If TEMPLIST.ListItems(t%).Selected Then
        Set itmx = TEMPLIST.ListItems(t%)
        Set rs = db.OpenRecordset("select ucr" + CStr(whichucr) + " from incidentsupport where incidentnumber = '" + itmx.SubItems(1) + "'")
        If Not rs.EOF Then
            rs.MoveFirst
            rs.Edit
            rs("ucr" + CStr(whichucr)) = inp
            rs.Update
        End If
        TEMPLIST.ListItems(t%).Selected = False
    End If
    If TEMPLIST.SelectedItem Is Nothing Then
        t% = TEMPLIST.ListItems.Count
    End If
Next t%
GETOUT:
Screen.MousePointer = 0
rightclickframe.Visible = False
Exit Sub
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.BorderStyle = 0
Label1.BorderStyle = 0
Label2.BorderStyle = 0
Label3.BorderStyle = 0
Label4.BorderStyle = 0
Label5.BorderStyle = 0
Label6.BorderStyle = 0
Label2.BorderStyle = 1
Label7.BorderStyle = 0
End Sub

Private Sub Label3_Click()
inp = InputBox("Enter Property UCR to change (1 - 6).", "Genesis Information Log", "1")
If Val(inp) < 1 Or Val(inp) > 6 Then
    GoTo GETOUT
End If
whichucr = Val(inp)
inp = InputBox("Enter new Property UCR value.", "Genesis Information Log", "")
If inp = "" Or Len(inp) <> 3 Then
    GoTo GETOUT
End If
Screen.MousePointer = 0
Dim db As Database, rs As Recordset
Set db = OpenDatabase(nwi + "incident.mdb")
For t% = 1 To TEMPLIST.ListItems.Count
    If TEMPLIST.ListItems(t%).Selected Then
        Set itmx = TEMPLIST.ListItems(t%)
        Set rs = db.OpenRecordset("select pucr" + CStr(whichucr) + " from incidentsupport where incidentnumber = '" + itmx.SubItems(1) + "'")
        If Not rs.EOF Then
            rs.MoveFirst
            rs.Edit
            rs("pucr" + CStr(whichucr)) = inp
            rs.Update
        End If
        TEMPLIST.ListItems(t%).Selected = False
    End If
    If TEMPLIST.SelectedItem Is Nothing Then
        t% = TEMPLIST.ListItems.Count
    End If
Next t%
GETOUT:
Screen.MousePointer = 0
rightclickframe.Visible = False
Exit Sub

End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.BorderStyle = 0
Label1.BorderStyle = 0
Label2.BorderStyle = 0
Label3.BorderStyle = 0
Label4.BorderStyle = 0
Label5.BorderStyle = 0
Label6.BorderStyle = 0
Label3.BorderStyle = 1
Label7.BorderStyle = 0
End Sub

Private Sub Label4_Click()
IORS% = 0
For t% = 0 To Forms.Count - 1
    If UCase(Forms(t%).Name) = "INCIDENT" Then
        IORS% = 1
    End If
    If UCase(Forms(t%).Name) = "SINCIDEN" Then
        IORS% = 2
    End If
Next t%
If IORS% = 0 Then
    IORS% = 1
End If

For t% = 1 To TEMPLIST.ListItems.Count
    If TEMPLIST.ListItems(t%).Selected = True Then
        Set itmx = TEMPLIST.ListItems(t%)
        If IORS% = 1 Then
            incident.WindowState = vbMaximized
            incident.incidentnumber = itmx.SubItems(1)
            incident.optimer.Enabled = False
            Call incident.getincident
            incident.optimer.Enabled = True
        Else
            Open "NP.TAG" For Output As #1
            Print #1, itmx.SubItems(1)
            Print #1, "1"
            Print #1, itmx.SubItems(2)
            Close #1
            Unload sinciden
            sinciden.Show
        End If
        t% = TEMPLIST.ListItems.Count
    End If
Next t%
Unload temprevw
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.BorderStyle = 0
Label1.BorderStyle = 0
Label2.BorderStyle = 0
Label3.BorderStyle = 0
Label4.BorderStyle = 0
Label5.BorderStyle = 0
Label6.BorderStyle = 0
Label4.BorderStyle = 1
Label7.BorderStyle = 0
End Sub

Private Sub Label5_Click()
Screen.MousePointer = 0
Dim db As Database, rs As Recordset
Set db = OpenDatabase(nwi + "incident.mdb")
For t% = 1 To TEMPLIST.ListItems.Count
    If TEMPLIST.ListItems(t%).Selected Then
        Set itmx = TEMPLIST.ListItems(t%)
        Set rs = db.OpenRecordset("select temp from incidentsupport where incidentnumber = '" + itmx.SubItems(1) + "'")
        If Not rs.EOF Then
            rs.MoveFirst
            rs.Edit
            rs("temp") = "N"
            rs.Update
        End If
        TEMPLIST.ListItems(t%).Selected = False
    End If
    If TEMPLIST.SelectedItem Is Nothing Then
        t% = TEMPLIST.ListItems.Count
    End If
Next t%
GETOUT:
Screen.MousePointer = 0
rightclickframe.Visible = False
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.BorderStyle = 0
Label1.BorderStyle = 0
Label2.BorderStyle = 0
Label3.BorderStyle = 0
Label4.BorderStyle = 0
Label5.BorderStyle = 0
Label6.BorderStyle = 0
Label5.BorderStyle = 1
Label7.BorderStyle = 0
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.BorderStyle = 0
Label1.BorderStyle = 0
Label2.BorderStyle = 0
Label3.BorderStyle = 0
Label4.BorderStyle = 0
Label5.BorderStyle = 0
Label6.BorderStyle = 0
Label6.BorderStyle = 1
Label7.BorderStyle = 0
End Sub

Private Sub Label7_Click()
inp = InputBox("Enter new Victim 1/Subject 1 relationship value.", "Genesis Information Log", "")
If inp = "" Or Len(inp) <> 2 Then
    GoTo GETOUT
End If
Screen.MousePointer = 0
Dim db As Database, rs As Recordset
Set db = OpenDatabase(nwi + "incident.mdb")
For t% = 1 To TEMPLIST.ListItems.Count
    If TEMPLIST.ListItems(t%).Selected Then
        Set itmx = TEMPLIST.ListItems(t%)
        Set rs = db.OpenRecordset("select vrelationship1 from incidentreportv where incidentnumber = '" + itmx.SubItems(1) + "'")
        If Not rs.EOF Then
            rs.MoveFirst
            rs.Edit
            rs("vrelationship1") = inp
            rs.Update
        End If
        TEMPLIST.ListItems(t%).Selected = False
    End If
    If TEMPLIST.SelectedItem Is Nothing Then
        t% = TEMPLIST.ListItems.Count
    End If
Next t%
GETOUT:
Screen.MousePointer = 0
rightclickframe.Visible = False
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.BorderStyle = 0
Label1.BorderStyle = 0
Label2.BorderStyle = 0
Label3.BorderStyle = 0
Label4.BorderStyle = 0
Label5.BorderStyle = 0
Label6.BorderStyle = 0
Label1.BorderStyle = 0
Label7.BorderStyle = 1
End Sub

Private Sub Label8_Click()
Screen.MousePointer = 0
Dim db As Database, rs As Recordset
Set db = OpenDatabase(nwi + "incident.mdb")
For t% = 1 To TEMPLIST.ListItems.Count
    If TEMPLIST.ListItems(t%).Selected Then
        Set itmx = TEMPLIST.ListItems(t%)
        Set rs = db.OpenRecordset("select seizedvalue1 from incidentreporto where incidentnumber = '" + itmx.SubItems(1) + "'")
        If Not rs.EOF Then
            rs.MoveFirst
            rs.Edit
            rs("seizedvalue1") = 9999999
            rs.Update
        End If
        TEMPLIST.ListItems(t%).Selected = False
    End If
    If TEMPLIST.SelectedItem Is Nothing Then
        t% = TEMPLIST.ListItems.Count
    End If
Next t%
GETOUT:
Screen.MousePointer = 0
rightclickframe.Visible = False
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.BorderStyle = 1
Label1.BorderStyle = 0
Label2.BorderStyle = 0
Label3.BorderStyle = 0
Label4.BorderStyle = 0
Label5.BorderStyle = 0
Label6.BorderStyle = 0
Label1.BorderStyle = 0
Label7.BorderStyle = 0
End Sub

Private Sub TEMPLIST_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
TEMPLIST.SortKey = ColumnHeader.index - 1
If TEMPLIST.SortOrder = lvwAscending Then
    TEMPLIST.SortOrder = lvwDescending
Else
    TEMPLIST.SortOrder = lvwAscending
End If
TEMPLIST.Sorted = True

End Sub

Private Sub TEMPLIST_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    If rightclickframe.Visible = False Then
        If Y + 100 + rightclickframe.Height > TEMPLIST.Height Then
            rightclickframe.Top = TEMPLIST.Height - rightclickframe.Height
        Else
            rightclickframe.Top = Y + 100
        End If
        rightclickframe.Left = X + 100
        rightclickframe.Visible = True
    Else
        rightclickframe.Visible = False
    End If
End If
End Sub

