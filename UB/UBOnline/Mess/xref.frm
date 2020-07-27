VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form xref 
   Caption         =   "Cross Reference"
   ClientHeight    =   7455
   ClientLeft      =   90
   ClientTop       =   1545
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7455
   ScaleWidth      =   11685
   Begin VB.CommandButton Command8 
      Caption         =   "Print Cross Reference"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8895
      TabIndex        =   22
      Top             =   15
      Width           =   2775
   End
   Begin VB.CommandButton Command7 
      Caption         =   "GO TO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   8880
      TabIndex        =   19
      Top             =   5055
      Width           =   600
   End
   Begin VB.CommandButton Command6 
      Caption         =   "GO TO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   1410
      TabIndex        =   16
      Top             =   5055
      Width           =   600
   End
   Begin VB.CommandButton Command5 
      Caption         =   "GO TO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   8880
      TabIndex        =   13
      Top             =   2700
      Width           =   600
   End
   Begin VB.CommandButton Command4 
      Caption         =   "GO TO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   1410
      TabIndex        =   10
      Top             =   2700
      Width           =   600
   End
   Begin VB.CommandButton Command3 
      Caption         =   "GO TO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   8880
      TabIndex        =   7
      Top             =   390
      Width           =   600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GO TO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   1410
      TabIndex        =   5
      Top             =   390
      Width           =   600
   End
   Begin VB.OptionButton similar 
      Caption         =   "Like"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3870
      TabIndex        =   2
      Top             =   180
      Width           =   855
   End
   Begin VB.OptionButton exact 
      Caption         =   "Exact Match"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3870
      TabIndex        =   1
      Top             =   -30
      Width           =   1335
   End
   Begin VB.TextBox subject 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   3570
   End
   Begin MSComctlLib.ListView civill 
      Height          =   2100
      Left            =   15
      TabIndex        =   4
      Top             =   600
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   3704
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Service Of"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date Received"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Iter"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Plaintiff"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Defendant"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Paper Type"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Type"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Cross Reference"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5535
      TabIndex        =   3
      Top             =   15
      Width           =   3240
   End
   Begin MSComctlLib.ListView incidentl 
      Height          =   2100
      Left            =   5925
      TabIndex        =   6
      Top             =   615
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   3704
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Incident#"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Offense Date"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Offense 1"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Complainant 1"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Victim 1"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Subject 1"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "NonCriminal"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "BadCheck"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Others"
         Object.Width           =   17639
      EndProperty
   End
   Begin MSComctlLib.ListView warrantl 
      Height          =   2100
      Left            =   15
      TabIndex        =   11
      Top             =   2910
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   3704
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Warrant#"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Log Date"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Charge"
         Object.Width           =   4851
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Witness"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Issued By"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ListView ROL 
      Height          =   2100
      Left            =   5925
      TabIndex        =   14
      Top             =   2910
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   3704
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Case Number"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Plaintiff"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Defendant"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Effective Date"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Effective Time"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Expiration Date"
         Object.Width           =   2822
      EndProperty
   End
   Begin MSComctlLib.ListView bookingl 
      Height          =   2100
      Left            =   15
      TabIndex        =   17
      Top             =   5265
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   3704
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Incident Number"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Subject#"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Arrestee"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date/Time of Arrest"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Charge A"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Alias"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Next of Kin"
         Object.Width           =   4410
      EndProperty
   End
   Begin MSComctlLib.ListView servicel 
      Height          =   2100
      Left            =   5925
      TabIndex        =   20
      Top             =   5265
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   3704
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Case Number"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date Received"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Alarm"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Unlock"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Prop Chk"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Funeral"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "HouseMove"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Mental Trans"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Other Escort"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Warrant"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Unfounded"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Other"
         Object.Width           =   1323
      EndProperty
   End
   Begin VB.Label Label6 
      Caption         =   "S E R V I C E   C A L L"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5925
      TabIndex        =   21
      Top             =   5055
      Width           =   2160
   End
   Begin VB.Label Label5 
      Caption         =   "B O O K I N G"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15
      TabIndex        =   18
      Top             =   5055
      Width           =   11505
   End
   Begin VB.Label Label4 
      Caption         =   "R E S T R A I N I N G   O R D E R"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5925
      TabIndex        =   15
      Top             =   2715
      Width           =   11505
   End
   Begin VB.Label Label3 
      Caption         =   "W A R R A N T"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15
      TabIndex        =   12
      Top             =   2715
      Width           =   11505
   End
   Begin VB.Label Label2 
      Caption         =   "I N C I D E N T"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5925
      TabIndex        =   9
      Top             =   390
      Width           =   10935
   End
   Begin VB.Label Label1 
      Caption         =   "C I V I L"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15
      TabIndex        =   8
      Top             =   390
      Width           =   11505
   End
End
Attribute VB_Name = "xref"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Screen.MousePointer = 11
On Error GoTo 0
If subject = "" Then
    msg = MsgBox("You must enter all or part of the name to generate a cross reference action.", 48, "Genesis Error Log")
    Exit Sub
End If
Dim db As Database, rs, rs2, rs3, rs4 As Recordset, itmx As ListItem, tbname(4)
servicexref:
If servicel.Enabled = False Then
    GoTo bookingxref
End If
servicel.ListItems.clear
Set db = OpenDatabase(nws + "SERVICE.mdb")
If similar Then
    Set rs = db.OpenRecordset("select * from service where compsubj like '*" + subject + "*' order by received")
Else
    Set rs = db.OpenRecordset("select * from service where compsubj = " + Chr$(34) + subject + Chr$(34) + " order by received")
End If
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        Set itmx = servicel.ListItems.add(, , rs("casenumber"))
        If rs("complainant") Then
            itmx.SubItems(1) = "Compl"
        Else
            itmx.SubItems(1) = "Subject"
        End If
        itmx.SubItems(2) = rs("compsubj")
        If Not IsNull(rs("received")) Then
            itmx.SubItems(3) = rs("received")
        End If
        If rs("alarm") = 1 Then
            itmx.SubItems(4) = "X"
        End If
        If rs("unlocking") = 1 Then
            itmx.SubItems(5) = "X"
        End If
        If rs("property") = 1 Then
            itmx.SubItems(6) = "X"
        End If
        If rs("funeral") Then
            itmx.SubItems(7) = "X"
        End If
        If rs("house") Then
            itmx.SubItems(8) = "X"
        End If
        If rs("mental") Then
            itmx.SubItems(9) = "X"
        End If
        itmx.SubItems(10) = rs("escortother")
        itmx.SubItems(11) = rs("warrantnumber")
        If rs("unfounded") = 1 Then
            itmx.SubItems(12) = "X"
        End If
        itmx.SubItems(12) = rs("otherspecify")
        rs.MoveNext
    Wend
End If

bookingxref:
If bookingl.Enabled = False Then
    GoTo roxref
End If
On Error GoTo oderror1
od1:
Set db = OpenDatabase(nwb + "BOOKING.mdb")
bookingl.ListItems.clear
If similar Then
    'rlb code - removed noncrim, badcheck from following queries - fields not in booking.mdb
    Set rs = db.OpenRecordset("select number, incidentnumber, sname, alias, nextofkin, chargea, datetimeofarrest from booking where sname like '*" + subject + "*' or alias like '*" + subject + "*' or nextofkin like '*" + subject + "*'")
Else
    Set rs = db.OpenRecordset("select number, incidentnumber, sname, alias, nextofkin, chargea, datetimeofarrest from booking where sname = " + Chr$(34) + subject + Chr$(34) + " or alias = " + Chr$(34) + subject + Chr$(34) + "  or nextofkin = " + Chr$(34) + subject + Chr$(34))
End If
    '*****
On Error Resume Next
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        Set itmx = bookingl.ListItems.add(, , rs("incidentnumber"))
        itmx.SubItems(1) = rs("number")
        itmx.SubItems(2) = rs("sname")
        itmx.SubItems(3) = rs("datetimeofarrest")
        itmx.SubItems(4) = rs("chargea")
        itmx.SubItems(5) = rs("alias")
        itmx.SubItems(6) = rs("nextofkin")
        'If Not IsNull(rs("noncrim")) Then
        '    ITMX.SubItems(7) = rs("noncrim")
        'End If
        'If Not IsNull(rs("badcheck")) Then
        '    ITMX.SubItems(8) = rs("badcheck")
        'End If
        rs.MoveNext
    Wend
End If
roxref:
If ROL.Enabled = False Then
    GoTo warrantxref
End If
On Error GoTo oderror2
od2:
Set db = OpenDatabase(nwr + "ro.mdb")
ROL.ListItems.clear
If similar Then
    Set rs = db.OpenRecordset("select plaintiff, defendant, effectivedate, effectivetime, expiration, casenumber from rorder where plaintiff like '*" + subject + "*' or defendant like '*" + subject + "*' order by effectivedate")
Else
    Set rs = db.OpenRecordset("select plaintiff, defendant, effectivedate, effectivetime, expiration, casenumber from rorder where plaintiff = " + Chr$(34) + subject + Chr$(34) + " or defendant = " + Chr$(34) + subject + Chr$(34) + " order by effectivedate")
End If
On Error Resume Next
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        Set itmx = ROL.ListItems.add(, , rs("casenumber"))
        itmx.SubItems(1) = rs("plaintiff")
        itmx.SubItems(2) = rs("defendant")
        itmx.SubItems(3) = rs("effectivedate")
        itmx.SubItems(4) = rs("effectivetime")
        itmx.SubItems(5) = rs("expiration")
        rs.MoveNext
    Wend
End If
warrantxref:
If warrantl.Enabled = False Then
    GoTo civilxref
End If
On Error GoTo oderror3
od3:
Set db = OpenDatabase(nww + "warrant.mdb")
warrantl.ListItems.clear
If similar Then
    Set rs = db.OpenRecordset("select warrant, logdate, wname, witness, charge, issuedby from warrantinfo where wname like '*" + subject + "*' or witness like '*" + subject + "*' order by logdate, warrant")
Else
    Set rs = db.OpenRecordset("select warrant, logdate, wname, witness, charge, issuedby from warrantinfo where wname = " + Chr$(34) + subject + Chr$(34) + " or witness = " + Chr$(34) + subject + Chr$(34) + " order by logdate, warrant")
End If
On Error Resume Next
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        Set itmx = warrantl.ListItems.add(, , rs("warrant"))
        itmx.SubItems(1) = rs("logdate")
        itmx.SubItems(2) = rs("charge")
        itmx.SubItems(3) = rs("wname")
        itmx.SubItems(4) = rs("witness")
        itmx.SubItems(5) = rs("issuedby")
        rs.MoveNext
    Wend
End If
civilxref:
If civill.Enabled = False Then
    GoTo incidentxref
End If
tbname(1) = "Magistrate"
tbname(2) = "WritOther"
tbname(3) = "FamilyCourt"
tbname(4) = "Executions"
On Error GoTo oderror4
od4:
Set db = OpenDatabase(nwc + "civil.mdb")
civill.ListItems.clear
For t% = 1 To 4
    If similar Then
        Set rs = db.OpenRecordset("select serviceof, datereceived, iteration, plaintiff, defendant, papertype from " + tbname(t%) + " where serviceof like '*" + subject + "*' or plaintiff like '*" + subject + "*' or defendant like '*" + subject + "*' order by datereceived, iteration")
    Else
        Set rs = db.OpenRecordset("select serviceof, datereceived, iteration, plaintiff, defendant, papertype from " + tbname(t%) + " where serviceof = " + Chr$(34) + subject + Chr$(34) + " or plaintiff = " + Chr$(34) + subject + Chr$(34) + " or defendant = " + Chr$(34) + subject + Chr$(34) + " order by datereceived, iteration")
    End If
    On Error Resume Next
    If Not rs.EOF Then
        rs.MoveFirst
        While Not rs.EOF
            Set itmx = civill.ListItems.add(, , rs("serviceof"))
            itmx.SubItems(1) = rs("datereceived")
            itmx.SubItems(2) = rs("iteration")
            itmx.SubItems(3) = rs("plaintiff")
            itmx.SubItems(4) = rs("defendant")
            itmx.SubItems(5) = rs("papertype")
            itmx.SubItems(6) = tbname(t%)
            rs.MoveNext
        Wend
    End If
Next t%
incidentxref:
On Error GoTo oderror5
od5:
Set db = OpenDatabase(nwi + "incident.mdb")
incidentl.ListItems.clear
If similar Then
    Set rs = db.OpenRecordset("select incidentnumber, cname, offense1, dateofoffense1 from incidentreportc where cname like '*" + subject + "*'")
Else
    Set rs = db.OpenRecordset("select incidentnumber, cname, offense1, dateofoffense1 from incidentreportc where cname = " + Chr$(34) + subject + Chr$(34))
End If
On Error Resume Next
If Not rs.EOF Then
    rs.MoveFirst
    Set itmx = incidentl.ListItems.add(, , rs("incidentnumber"))
    itmx.SubItems(1) = rs("dateofoffense1")
    itmx.SubItems(2) = rs("offense1")
    itmx.SubItems(3) = rs("cname")
    Set rs2 = db.OpenRecordset("select vname from incidentreportv where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    rs2.MoveFirst
    itmx.SubItems(4) = rs2("vname")
    Set rs2 = db.OpenRecordset("select sname from incidentreports where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    rs2.MoveFirst
    itmx.SubItems(5) = rs2("sname")
    temp$ = ""
    Set rs2 = db.OpenRecordset("select name1, name2, complainant1, complainant2, victim1, victim2, subject1, subject2, runaway1, runaway2, wanted1, wanted2, warrant1, warrant2, arrest1, arrest2,jail1, jail2, summons1, summons2, typeother1, typeother2 from supplemental where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    If Not rs2.EOF Then
        rs2.MoveFirst
        While Not rs2.EOF
            For t% = 1 To 2
                If Not IsNull(rs2("name" + Mid$(Str$(t%), 2))) Then
                    temp$ = temp$ + rs2("name" + Mid$(Str$(t%), 2)) + "("
                    If rs2("complainant" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Complainant,"
                    End If
                    If rs2("victim" + Mid$(Str$(t%), 2)) > 0 Then
                        temp$ = temp$ + "Victim " + Mid$(Str$(rs2("victim" + Mid$(Str$(t%), 2))), 2) + ","
                    End If
                    If rs2("subject" + Mid$(Str$(t%), 2)) > 0 Then
                        temp$ = temp$ + "Subject " + Mid$(Str$(rs2("subject" + Mid$(Str$(t%), 2))), 2) + ","
                    End If
                    If rs2("Runaway" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Runaway,"
                    End If
                    If rs2("Wanted" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Wanted,"
                    End If
                    If rs2("Warrant" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Warrant,"
                    End If
                    If rs2("Arrest" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Arrest,"
                    End If
                    If rs2("Jail" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Jail,"
                    End If
                    If rs2("Summons" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Summons,"
                    End If
                    If Not IsNull(rs2("typeother" + Mid$(Str$(t%), 2))) And rs2("typeother" + Mid$(Str$(t%), 2)) > "" Then
                        temp$ = temp$ + rs2("typeother" + Mid$(Str$(t%), 2)) + ")   "
                    End If
                    If Right$(temp$, 1) = "," Then
                        Mid$(temp$, Len(temp$), 1) = ")   "
                    End If
                End If
            Next t%
            rs2.MoveNext
        Wend
    End If
    itmx.SubItems(6) = temp$
End If
If similar Then
    Set rs = db.OpenRecordset("select incidentnumber, vname from incidentreportv where vname like '*" + subject + "*'")
Else
    Set rs = db.OpenRecordset("select incidentnumber, vname from incidentreportv where vname = " + Chr$(34) + subject + Chr$(34))
End If
If Not rs.EOF Then
    rs.MoveFirst
    Set rs2 = db.OpenRecordset("select incidentnumber, cname, offense1, dateofoffense1 from incidentreportc where incidentnumber  = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    rs2.MoveFirst
    Set itmx = incidentl.ListItems.add(, , rs2("incidentnumber"))
    itmx.SubItems(1) = rs2("dateofoffense1")
    itmx.SubItems(2) = rs2("offense1")
    itmx.SubItems(3) = rs2("cname")
    itmx.SubItems(4) = rs("vname")
    Set rs2 = db.OpenRecordset("select sname from incidentreports where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    rs2.MoveFirst
    itmx.SubItems(5) = rs2("sname")
    temp$ = ""
    Set rs2 = db.OpenRecordset("select name1, name2, complainant1, complainant2, victim1, victim2, subject1, subject2, runaway1, runaway2, wanted1, wanted2, warrant1, warrant2, arrest1, arrest2,jail1, jail2, summons1, summons2, typeother1, typeother2 from supplemental where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    If Not rs2.EOF Then
        rs2.MoveFirst
        While Not rs2.EOF
            For t% = 1 To 2
                If Not IsNull(rs2("name" + Mid$(Str$(t%), 2))) Then
                    temp$ = temp$ + rs2("name" + Mid$(Str$(t%), 2)) + "("
                    If rs2("complainant" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Complainant,"
                    End If
                    If rs2("victim" + Mid$(Str$(t%), 2)) > 0 Then
                        temp$ = temp$ + "Victim " + Mid$(Str$(rs2("victim" + Mid$(Str$(t%), 2))), 2) + ","
                    End If
                    If rs2("subject" + Mid$(Str$(t%), 2)) > 0 Then
                        temp$ = temp$ + "Subject " + Mid$(Str$(rs2("subject" + Mid$(Str$(t%), 2))), 2) + ","
                    End If
                    If rs2("Runaway" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Runaway,"
                    End If
                    If rs2("Wanted" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Wanted,"
                    End If
                    If rs2("Warrant" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Warrant,"
                    End If
                    If rs2("Arrest" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Arrest,"
                    End If
                    If rs2("Jail" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Jail,"
                    End If
                    If rs2("Summons" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Summons,"
                    End If
                    If Not IsNull(rs2("typeother" + Mid$(Str$(t%), 2))) And rs2("typeother" + Mid$(Str$(t%), 2)) > "" Then
                        temp$ = temp$ + rs2("typeother" + Mid$(Str$(t%), 2)) + ")   "
                    End If
                    If Right$(temp$, 1) = "," Then
                        Mid$(temp$, Len(temp$), 1) = ")   "
                    End If
                End If
            Next t%
            rs2.MoveNext
        Wend
    End If
    itmx.SubItems(6) = temp$
End If
If similar Then
    Set rs = db.OpenRecordset("select incidentnumber, sname from incidentreports where sname like '*" + subject + "*'")
Else
    Set rs = db.OpenRecordset("select incidentnumber, sname from incidentreports where sname = " + Chr$(34) + subject + Chr$(34))
End If
If Not rs.EOF Then
    rs.MoveFirst
    Set rs2 = db.OpenRecordset("select incidentnumber, cname, offense1, dateofoffense1 from incidentreportc where incidentnumber  = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    rs2.MoveFirst
    Set itmx = incidentl.ListItems.add(, , rs2("incidentnumber"))
    itmx.SubItems(1) = rs2("dateofoffense1")
    itmx.SubItems(2) = rs2("offense1")
    itmx.SubItems(3) = rs2("cname")
    itmx.SubItems(5) = rs("sname")
    Set rs2 = db.OpenRecordset("select vname from incidentreportv where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    rs2.MoveFirst
    itmx.SubItems(4) = rs2("vname")
    temp$ = ""
    Set rs2 = db.OpenRecordset("select name1, name2, complainant1, complainant2, victim1, victim2, subject1, subject2, runaway1, runaway2, wanted1, wanted2, warrant1, warrant2, arrest1, arrest2,jail1, jail2, summons1, summons2, typeother1, typeother2 from supplemental where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    If Not rs2.EOF Then
        rs2.MoveFirst
        While Not rs2.EOF
            For t% = 1 To 2
                If Not IsNull(rs2("name" + Mid$(Str$(t%), 2))) Then
                    temp$ = temp$ + rs2("name" + Mid$(Str$(t%), 2)) + "("
                    If rs2("complainant" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Complainant,"
                    End If
                    If rs2("victim" + Mid$(Str$(t%), 2)) > 0 Then
                        temp$ = temp$ + "Victim " + Mid$(Str$(rs2("victim" + Mid$(Str$(t%), 2))), 2) + ","
                    End If
                    If rs2("subject" + Mid$(Str$(t%), 2)) > 0 Then
                        temp$ = temp$ + "Subject " + Mid$(Str$(rs2("subject" + Mid$(Str$(t%), 2))), 2) + ","
                    End If
                    If rs2("Runaway" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Runaway,"
                    End If
                    If rs2("Wanted" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Wanted,"
                    End If
                    If rs2("Warrant" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Warrant,"
                    End If
                    If rs2("Arrest" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Arrest,"
                    End If
                    If rs2("Jail" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Jail,"
                    End If
                    If rs2("Summons" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Summons,"
                    End If
                    If Not IsNull(rs2("typeother" + Mid$(Str$(t%), 2))) And rs2("typeother" + Mid$(Str$(t%), 2)) > "" Then
                        temp$ = temp$ + rs2("typeother" + Mid$(Str$(t%), 2)) + ")   "
                    End If
                    If Right$(temp$, 1) = "," Then
                        Mid$(temp$, Len(temp$), 1) = ")   "
                    End If
                End If
            Next t%
            rs2.MoveNext
        Wend
    End If
    itmx.SubItems(6) = temp$
End If
If similar Then
    Set rs = db.OpenRecordset("select incidentnumber, name1 from supplemental where name1 like '*" + subject + "*'")
Else
    Set rs = db.OpenRecordset("select incidentnumber, name1 from supplemental where name1 = " + Chr$(34) + subject + Chr$(34))
End If
If Not rs.EOF Then
    rs.MoveFirst
    Set rs2 = db.OpenRecordset("select incidentnumber, cname, offense1, dateofoffense1 from incidentreportc where incidentnumber  = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    rs2.MoveFirst
    Set itmx = incidentl.ListItems.add(, , rs2("incidentnumber"))
    itmx.SubItems(1) = rs2("dateofoffense1")
    itmx.SubItems(2) = rs2("offense1")
    itmx.SubItems(3) = rs2("cname")
    Set rs2 = db.OpenRecordset("select vname from incidentreportv where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    rs2.MoveFirst
    itmx.SubItems(4) = rs2("vname")
    Set rs2 = db.OpenRecordset("select sname from incidentreports where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    rs2.MoveFirst
    itmx.SubItems(5) = rs2("sname")
    temp$ = ""
    Set rs2 = db.OpenRecordset("select name1, name2, complainant1, complainant2, victim1, victim2, subject1, subject2, runaway1, runaway2, wanted1, wanted2, warrant1, warrant2, arrest1, arrest2,jail1, jail2, summons1, summons2, typeother1, typeother2 from supplemental where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    If Not rs2.EOF Then
        rs2.MoveFirst
        While Not rs2.EOF
            For t% = 1 To 2
                If Not IsNull(rs2("name" + Mid$(Str$(t%), 2))) Then
                    temp$ = temp$ + rs2("name" + Mid$(Str$(t%), 2)) + "("
                    If rs2("complainant" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Complainant,"
                    End If
                    If rs2("victim" + Mid$(Str$(t%), 2)) > 0 Then
                        temp$ = temp$ + "Victim " + Mid$(Str$(rs2("victim" + Mid$(Str$(t%), 2))), 2) + ","
                    End If
                    If rs2("subject" + Mid$(Str$(t%), 2)) > 0 Then
                        temp$ = temp$ + "Subject " + Mid$(Str$(rs2("subject" + Mid$(Str$(t%), 2))), 2) + ","
                    End If
                    If rs2("Runaway" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Runaway,"
                    End If
                    If rs2("Wanted" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Wanted,"
                    End If
                    If rs2("Warrant" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Warrant,"
                    End If
                    If rs2("Arrest" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Arrest,"
                    End If
                    If rs2("Jail" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Jail,"
                    End If
                    If rs2("Summons" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Summons,"
                    End If
                    If Not IsNull(rs2("typeother" + Mid$(Str$(t%), 2))) And rs2("typeother" + Mid$(Str$(t%), 2)) > "" Then
                        temp$ = temp$ + rs2("typeother" + Mid$(Str$(t%), 2)) + ")   "
                    End If
                    If Right$(temp$, 1) = "," Then
                        Mid$(temp$, Len(temp$), 1) = ")   "
                    End If
                End If
            Next t%
            rs2.MoveNext
        Wend
    End If
    itmx.SubItems(6) = temp$
End If
If similar Then
    Set rs = db.OpenRecordset("select incidentnumber, name2 from supplemental where name2 like '*" + subject + "*'")
Else
    Set rs = db.OpenRecordset("select incidentnumber, name2 from supplemental where name2 = " + Chr$(34) + subject + Chr$(34))
End If
If Not rs.EOF Then
    rs.MoveFirst
    Set rs2 = db.OpenRecordset("select incidentnumber, cname, offense1, dateofoffense1 from incidentreportc where incidentnumber  = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    rs2.MoveFirst
    Set itmx = incidentl.ListItems.add(, , rs2("incidentnumber"))
    itmx.SubItems(1) = rs2("dateofoffense1")
    itmx.SubItems(2) = rs2("offense1")
    itmx.SubItems(3) = rs2("cname")
    Set rs2 = db.OpenRecordset("select vname from incidentreportv where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    rs2.MoveFirst
    itmx.SubItems(4) = rs2("vname")
    Set rs2 = db.OpenRecordset("select sname from incidentreports where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    rs2.MoveFirst
    itmx.SubItems(5) = rs2("sname")
    temp$ = ""
    Set rs2 = db.OpenRecordset("select name1, name2, complainant1, complainant2, victim1, victim2, subject1, subject2, runaway1, runaway2, wanted1, wanted2, warrant1, warrant2, arrest1, arrest2,jail1, jail2, summons1, summons2, typeother1, typeother2 from supplemental where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    If Not rs2.EOF Then
        rs2.MoveFirst
        While Not rs2.EOF
            For t% = 1 To 2
                If Not IsNull(rs2("name" + Mid$(Str$(t%), 2))) Then
                    temp$ = temp$ + rs2("name" + Mid$(Str$(t%), 2)) + "("
                    If rs2("complainant" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Complainant,"
                    End If
                    If rs2("victim" + Mid$(Str$(t%), 2)) > 0 Then
                        temp$ = temp$ + "Victim " + Mid$(Str$(rs2("victim" + Mid$(Str$(t%), 2))), 2) + ","
                    End If
                    If rs2("subject" + Mid$(Str$(t%), 2)) > 0 Then
                        temp$ = temp$ + "Subject " + Mid$(Str$(rs2("subject" + Mid$(Str$(t%), 2))), 2) + ","
                    End If
                    If rs2("Runaway" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Runaway,"
                    End If
                    If rs2("Wanted" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Wanted,"
                    End If
                    If rs2("Warrant" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Warrant,"
                    End If
                    If rs2("Arrest" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Arrest,"
                    End If
                    If rs2("Jail" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Jail,"
                    End If
                    If rs2("Summons" + Mid$(Str$(t%), 2)) = 1 Then
                        temp$ = temp$ + "Summons,"
                    End If
                    If Not IsNull(rs2("typeother" + Mid$(Str$(t%), 2))) And rs2("typeother" + Mid$(Str$(t%), 2)) > "" Then
                        temp$ = temp$ + rs2("typeother" + Mid$(Str$(t%), 2)) + ")   "
                    End If
                    If Right$(temp$, 1) = "," Then
                        Mid$(temp$, Len(temp$), 1) = ")   "
                    End If
                End If
            Next t%
            rs2.MoveNext
        Wend
    End If
    itmx.SubItems(6) = temp$
End If
db.Close
On Error GoTo 0
Screen.MousePointer = 0
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
oderror3:
If Err > 3200 Then
    Resume od3
Else
    Resume Next
End If
oderror4:
If Err > 3200 Then
    Resume od4
Else
    Resume Next
End If
oderror5:
If Err > 3200 Then
    Resume od5
Else
    Resume Next
End If
End Sub

Private Sub Command2_Click()
If civill.SelectedItem Is Nothing Then
    msg = MsgBox("An item must be selected prior to ACCESS.", 48, "Genesis Error Log")
    Exit Sub
End If
Dim itmx As ListItem
Set itmx = civill.ListItems(civill.SelectedItem.index)
CIVIL.serviceof = itmx
CIVIL.datereceived = itmx.SubItems(1)
CIVIL.iteration = itmx.SubItems(2)
If itmx.SubItems(6) = "Magistrate" Then
    CIVIL.maintab.Tab = 0
End If
If itmx.SubItems(6) = "WritOther" Then
    CIVIL.maintab.Tab = 1
End If
If itmx.SubItems(6) = "FamilyCourt" Then
    CIVIL.maintab.Tab = 2
End If
If itmx.SubItems(6) = "Executions" Then
    CIVIL.maintab.Tab = 3
End If
Unload xref
End Sub

Private Sub Command3_Click()
If incidentl.SelectedItem Is Nothing Then
    msg = MsgBox("An item must be selected prior to ACCESS.", 48, "Genesis Error Log")
    Exit Sub
End If
Dim itmx As ListItem
Set itmx = incidentl.ListItems(incidentl.SelectedItem.index)
incident.incidentnumber = itmx
Unload xref
End Sub

Private Sub Command4_Click()
If warrantl.SelectedItem Is Nothing Then
    msg = MsgBox("An item must be selected prior to ACCESS.", 48, "Genesis Error Log")
    Exit Sub
End If
Dim itmx As ListItem
Set itmx = warrantl.ListItems(warrantl.SelectedItem.index)
warrant.warrant = itmx
Unload xref
End Sub

Private Sub Command5_Click()
If ROL.SelectedItem Is Nothing Then
    msg = MsgBox("An item must be selected prior to ACCESS.", 48, "Genesis Error Log")
    Exit Sub
End If
Dim itmx As ListItem
Set itmx = ROL.ListItems(ROL.SelectedItem.index)
ro.index = itmx.SubItems(1) + " VS. " + itmx.SubItems(2) + "  CASE#: " + itmx
Unload xref

End Sub

Private Sub Command6_Click()
If bookingl.SelectedItem Is Nothing Then
    msg = MsgBox("An item must be selected prior to ACCESS.", 48, "Genesis Error Log")
    Exit Sub
End If
Dim itmx As ListItem
Set itmx = bookingl.ListItems(bookingl.SelectedItem.index)
Load booking
booking.arrestnumber = itmx
booking.incidentnumber = itmx
booking.subjectnumber = itmx.SubItems(1)
booking.Show
Unload xref

End Sub

Private Sub Command7_Click()
If servicel.SelectedItem Is Nothing Then
    msg = MsgBox("An item must be selected prior to ACCESS.", 48, "Genesis Error Log")
    Exit Sub
End If
Dim itmx As ListItem
Set itmx = servicel.ListItems(servicel.SelectedItem.index)
service.casenumber = itmx
service.Show
Unload xref

End Sub

Private Sub Command8_Click()
Dim itmx As ListItem
a = civill.ListItems.Count + incidentl.ListItems.Count + warrantl.ListItems.Count + ROL.ListItems.Count + bookingl.ListItems.Count + servicel.ListItems.Count
If a = 0 Then
    msg = MsgBox("No Cross Reference results are present to print.", 48, "Genesis Error Log")
    Exit Sub
End If
lnct% = 0
Printer.FontName = "Times New Roman"
GoSub header
Printer.FontBold = False
Printer.FontSize = 10
Printer.Print "CIVIL RECORDS"
Printer.Print
lnct% = lnct% + 2
For t% = 1 To civill.ListItems.Count
    Set itmx = civill.ListItems(t%)
    Printer.Print Tab(5); "Service Of"; Tab(20); itmx
    Printer.Print Tab(5); "Date Received"; Tab(20); itmx.SubItems(1)
    Printer.Print Tab(5); "Iteration"; Tab(20); itmx.SubItems(2)
    Printer.Print Tab(5); "Plaintiff"; Tab(20); itmx.SubItems(3)
    Printer.Print Tab(5); "Defendant"; Tab(20); itmx.SubItems(4)
    Printer.Print Tab(5); "Paper Type"; Tab(20); itmx.SubItems(5)
    Printer.Print Tab(5); "Type"; Tab(20); itmx.SubItems(6)
    Printer.Print
    Printer.Print
    lnct% = lnct% + 10
    If lnct% > 55 Then
        GoSub header
    End If
Next t%

Printer.Print "INCIDENT REPORTS"
Printer.Print
lnct% = lnct% + 2
For t% = 1 To incidentl.ListItems.Count
    Set itmx = incidentl.ListItems(t%)
    Printer.Print Tab(5); "Incident#"; Tab(20); itmx
    Printer.Print Tab(5); "Offense Date"; Tab(20); itmx.SubItems(1)
    Printer.Print Tab(5); "Offense1"; Tab(20); itmx.SubItems(2)
    Printer.Print Tab(5); "Complainant1"; Tab(20); itmx.SubItems(3)
    Printer.Print Tab(5); "Victim1"; Tab(20); itmx.SubItems(4)
    Printer.Print Tab(5); "Subject1"; Tab(20); itmx.SubItems(5)
    Printer.Print Tab(5); "Others"; Tab(20); itmx.SubItems(6)
    Printer.Print
    Printer.Print
    lnct% = lnct% + 10
    If lnct% > 55 Then
        GoSub header
    End If
Next t%

Printer.Print "WARRANTS"
Printer.Print
lnct% = lnct% + 2
For t% = 1 To warrantl.ListItems.Count
    Set itmx = warrantl.ListItems(t%)
    Printer.Print Tab(5); "Warant #"; Tab(20); itmx
    Printer.Print Tab(5); "Log Date"; Tab(20); itmx.SubItems(1)
    Printer.Print Tab(5); "Charge"; Tab(20); itmx.SubItems(2)
    Printer.Print Tab(5); "Name"; Tab(20); itmx.SubItems(3)
    Printer.Print Tab(5); "Witness"; Tab(20); itmx.SubItems(4)
    Printer.Print Tab(5); "Issued By"; Tab(20); itmx.SubItems(5)
    Printer.Print
    Printer.Print
    lnct% = lnct% + 9
    If lnct% > 55 Then
        GoSub header
    End If
Next t%

Printer.Print "RESTRAINING ORDERS"
Printer.Print
lnct% = lnct% + 2
For t% = 1 To ROL.ListItems.Count
    Set itmx = ROL.ListItems(t%)
    Printer.Print Tab(5); "Case Number"; Tab(20); itmx
    Printer.Print Tab(5); "Plaintiff"; Tab(20); itmx.SubItems(1)
    Printer.Print Tab(5); "Defendant"; Tab(20); itmx.SubItems(2)
    Printer.Print Tab(5); "Effective Date"; Tab(20); itmx.SubItems(3)
    Printer.Print Tab(5); "Efective Time"; Tab(20); itmx.SubItems(4)
    Printer.Print Tab(5); "Expiration"; Tab(20); itmx.SubItems(5)
    Printer.Print
    Printer.Print
    lnct% = lnct% + 9
    If lnct% > 55 Then
        GoSub header
    End If
Next t%

Printer.Print "BOOKING REPORTS"
Printer.Print
lnct% = lnct% + 2
For t% = 1 To bookingl.ListItems.Count
    Set itmx = bookingl.ListItems(t%)
    Printer.Print Tab(5); "Incident#"; Tab(20); itmx
    Printer.Print Tab(5); "Subject #"; Tab(20); itmx.SubItems(1)
    Printer.Print Tab(5); "Arrestee"; Tab(20); itmx.SubItems(2)
    Printer.Print Tab(5); "Arrested On"; Tab(20); itmx.SubItems(3)
    Printer.Print Tab(5); "Charge A"; Tab(20); itmx.SubItems(4)
    Printer.Print Tab(5); "Alias"; Tab(20); itmx.SubItems(5)
    Printer.Print Tab(5); "Next of Kin"; Tab(20); itmx.SubItems(6)
    Printer.Print
    Printer.Print
    lnct% = lnct% + 10
    If lnct% > 55 Then
        GoSub header
    End If
Next t%

Printer.Print "SERVICE CALLS"
Printer.Print
lnct% = lnct% + 2
For t% = 1 To servicel.ListItems.Count
    Set itmx = servicel.ListItems(t%)
    Printer.Print Tab(5); "Case Number"; Tab(20); itmx
    Printer.Print Tab(5); "Type"; Tab(20); itmx.SubItems(1)
    Printer.Print Tab(5); "Name"; Tab(20); itmx.SubItems(2)
    Printer.Print Tab(5); "Call Received"; Tab(20); itmx.SubItems(3)
    Printer.Print Tab(5); "Alarm"; Tab(20); itmx.SubItems(4); Tab(30); "Unlocking"; Tab(45); itmx.SubItems(5); Tab(55); "Property Check"; Tab(70); itmx.SubItems(6)
    Printer.Print Tab(5); "Funeral"; Tab(20); itmx.SubItems(7); Tab(30); "House Moving"; Tab(45); itmx.SubItems(8); Tab(55); "Mental Trans"; Tab(70); itmx.SubItems(9); Tab(70); "Other Escort"; Tab(95); itmx.SubItems(10)
    Printer.Print Tab(5); "Warrant"; Tab(20); itmx.SubItems(11); Tab(30); "Unfounded"; Tab(45); itmx.SubItems(12); Tab(55); "Other"; Tab(70); itmx.SubItems(13)
    Printer.Print
    Printer.Print
    lnct% = lnct% + 8
    If lnct% > 55 Then
        GoSub header
    End If
Next t%
Printer.EndDoc
Exit Sub
header:
Printer.NewPage
Printer.FontBold = True
Printer.FontSize = 14
Printer.Print "CROSS REFERENCE REPORT"; Tab(80); Date$
Printer.Print
If similar Then
    Printer.Print "SIMILAR TO ";
End If
Printer.Print subject
Printer.Print
Printer.Print
Printer.Print
lnct% = 5
Printer.FontSize = 10
Printer.FontBold = False
Return
End Sub

Private Sub Form_Load()
xref.Left = 0
xref.Top = 0
xref.Height = 7850
xref.Width = 11800
similar = True
On Error Resume Next
a$ = ""
If frmLogin.CBROWSE(0) = 0 And frmLogin.CBROWSE(1) = 0 And frmLogin.CBROWSE(2) = 0 And frmLogin.CBROWSE(3) = 0 Then
    civill.Enabled = False
End If
If frmLogin.IBROWSE = 0 Then
    incidentl.Enabled = False
End If
If frmLogin.sbrowse = 0 Then
    servicel.Enabled = False
End If
If frmLogin.IBROWSE = 0 Then
    bookingl.Enabled = False
End If
If frmLogin.WBROWSE = 0 Then
    warrantl.Enabled = False
End If
If frmLogin.RBROWSE = 0 Then
    ROL.Enabled = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set xref = Nothing
End Sub
