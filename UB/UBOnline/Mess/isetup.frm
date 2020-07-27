VERSION 5.00
Begin VB.Form isetup 
   Caption         =   "Incident Data Setup"
   ClientHeight    =   4800
   ClientLeft      =   3405
   ClientTop       =   1800
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   6585
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5745
      Top             =   3195
   End
   Begin VB.CommandButton Command6 
      Caption         =   "C L O S E"
      Height          =   375
      Left            =   5040
      TabIndex        =   13
      Top             =   4230
      Width           =   1500
   End
   Begin VB.CommandButton Command5 
      Caption         =   "D E L E T E"
      Height          =   375
      Left            =   2655
      TabIndex        =   12
      Top             =   4245
      Width           =   1500
   End
   Begin VB.CommandButton Command4 
      Caption         =   "A D D"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   4260
      Width           =   1500
   End
   Begin VB.ListBox ucrlist 
      Height          =   1035
      Left            =   1155
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   2910
      Width           =   3525
   End
   Begin VB.ComboBox offense 
      Height          =   315
      Left            =   1155
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   2400
      Width           =   5310
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C L O S E"
      Height          =   375
      Left            =   5010
      TabIndex        =   4
      Top             =   1440
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "D E L E T E"
      Height          =   375
      Left            =   2625
      TabIndex        =   3
      Top             =   1440
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "A D D"
      Height          =   375
      Left            =   90
      TabIndex        =   2
      Top             =   1455
      Width           =   1500
   End
   Begin VB.ComboBox minor 
      Height          =   315
      Left            =   2910
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   690
      Width           =   3600
   End
   Begin VB.ComboBox major 
      Height          =   315
      Left            =   2895
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   3600
   End
   Begin VB.Label Label4 
      Caption         =   "UCR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   60
      TabIndex        =   9
      Top             =   2955
      Width           =   2805
   End
   Begin VB.Label Label3 
      Caption         =   "Offense"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   90
      TabIndex        =   7
      Top             =   2385
      Width           =   2805
   End
   Begin VB.Label Label2 
      Caption         =   "Minor Property Grouping"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   90
      TabIndex        =   6
      Top             =   645
      Width           =   2805
   End
   Begin VB.Label Label1 
      Caption         =   "Major Property Grouping"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   120
      TabIndex        =   5
      Top             =   135
      Width           =   2805
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   6600
      X2              =   15
      Y1              =   2145
      Y2              =   2145
   End
End
Attribute VB_Name = "isetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If major = "" Or minor = "" Then
    msg = MsgBox("Both Major and Minor Property Grouping must be populated.", 48, "Genesis Error Log")
    Exit Sub
End If
Dim db As Database, rs As Recordset
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("select * from pgroup where major = " + Chr$(34) + major + Chr$(34) + " and minor = " + Chr$(34) + minor + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    rs.Edit
Else
    rs.AddNew
End If
rs("major") = Left$(major, 30)
rs("minor") = Left$(minor, 30)
rs.Update
db.Close
hmajor$ = major
Call LOADMAJOR
major = hmajor$
minor.SetFocus

Call incident.LOADMAJOR
End Sub

Private Sub Command2_Click()
If major = "" Or minor = "" Then
    msg = MsgBox("Both Major and Minor Property Grouping must be populated.", 48, "Genesis Error Log")
    Exit Sub
End If
msg = MsgBox("Are you sure?", 4, "Genesis Information Log")
If msg <> 6 Then
    Exit Sub
End If
Dim db As Database, rs As Recordset
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("select * from pgroup where major = " + Chr$(34) + major + Chr$(34) + " and minor = " + Chr$(34) + minor + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    rs.Delete
End If
db.Close
Call LOADMAJOR
minor.clear
major.SetFocus

Call incident.LOADMAJOR
End Sub

Private Sub Command3_Click()
Unload Me
Call incident.LOADMAJOR
End Sub

Private Sub Command4_Click()
If offense = "" Or ucrlist.ListIndex = -1 Then
    msg = MsgBox("Both offense and ucr must be populated.", 48, "Genesis Error Log")
    Exit Sub
End If
Dim db As Database, rs As Recordset
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("select * from offense where offense = " + Chr$(34) + offense + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    rs.Edit
Else
    rs.AddNew
End If
rs("offense") = Left$(offense, 50)
rs("ucr") = ucrlist.List(ucrlist.ListIndex)
rs.Update
db.Close

Call loadoffense
ucrlist.ListIndex = -1

Call incident.loadoffense


End Sub

Private Sub Command5_Click()
If offense = "" Or ucrlist.ListIndex = -1 Then
    msg = MsgBox("Both offense and ucr must be populated.", 48, "Genesis Error Log")
    Exit Sub
End If
msg = MsgBox("Are you sure?", 4, "Genesis Information Log")
If msg <> 6 Then
    Exit Sub
End If
Dim db As Database, rs As Recordset
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("select * from offense where offense = " + Chr$(34) + offense + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    rs.Delete
End If
db.Close
Call loadoffense
ucrlist.ListIndex = -1

Call incident.loadoffense

End Sub

Private Sub Command6_Click()
Unload Me
Call incident.loadoffense
End Sub

Private Sub Form_Load()
Me.Height = 5200
Me.Width = 6700
On Error Resume Next
Kill "*.dsk"

Call LOADMAJOR
Call loadminor
Call loadoffense
Call loaducr



End Sub
Private Sub LOADMAJOR()
Dim db As Database, rs As Recordset
Set db = OpenDatabase(nwi + "incident.mdb")
major.clear
Set rs = db.OpenRecordset("select DISTINCT major from pgroup")
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        major.AddItem rs("major")
        rs.MoveNext
    Wend
End If
db.Close


End Sub
Private Sub loadoffense()
Dim db As Database, rs As Recordset
Set db = OpenDatabase(nwi + "incident.mdb")
offense.clear
Set rs = db.OpenRecordset("select offense from offense")
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        offense.AddItem rs("offense")
        rs.MoveNext
    Wend
End If
db.Close


End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    Set isetup = Nothing
End Sub

Private Sub major_Click()
If major > "" Then
    minor.SetFocus
End If
End Sub

Private Sub minor_GotFocus()
If major > "" Then
    Call loadminor
End If
End Sub
Private Sub loadminor()
Dim db As Database, rs As Recordset
Set db = OpenDatabase(nwi + "incident.mdb")
minor.clear
Set rs = db.OpenRecordset("select minor from pgroup where major = " + Chr$(34) + major + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        minor.AddItem rs("minor")
        rs.MoveNext
    Wend
End If
db.Close


End Sub

Private Sub loaducr()
Dim db As Database, rs As Recordset
Set db = OpenDatabase(nwi + "incident.mdb")
ucrlist.clear
Set rs = db.OpenRecordset("select CODE from ucr")
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        ucrlist.AddItem rs("CODE")
        rs.MoveNext
    Wend
End If
db.Close
End Sub


Private Sub offense_Click()
If offense > "" Then
    ucrlist.SetFocus
End If
End Sub

Private Sub Timer1_Timer()
    Me.Show
    Timer1.Enabled = False
End Sub

Private Sub ucrlist_GotFocus()
Dim db As Database, rs As Recordset
Set db = OpenDatabase(nwi + "incident.mdb")
finducr$ = ""
Set rs = db.OpenRecordset("select ucr from offense where offense = " + Chr$(34) + offense + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    finducr$ = rs("UCR")
End If
db.Close
If finducr$ = "" Then
    Exit Sub
End If
For t% = 0 To ucrlist.ListCount - 1
    If ucrlist.List(t%) = finducr$ Then
        ucrlist.ListIndex = t%
        t% = ucrlist.ListCount - 1
    End If
Next t%
End Sub
