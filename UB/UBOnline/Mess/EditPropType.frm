VERSION 5.00
Begin VB.Form EditPropType 
   Caption         =   "Major-Minor Propterty Type Maintenance"
   ClientHeight    =   5175
   ClientLeft      =   2265
   ClientTop       =   1590
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   5940
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.ComboBox Minor 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2880
      Width           =   4455
   End
   Begin VB.ComboBox Major 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Edit Minor Property Types"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Edit Major Property Types"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "EditPropType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
If Minor.Text = "" Then
    msg = MsgBox("A minor property type *MUST* be entered or selected from the list.", 48, "Genesis Error Log")
    Minor.SetFocus
Exit Sub
End If

If Major.Text = "" Then
    msg = MsgBox("A category in Major *MUST* be entered or selected from the list.", 48, "Genesis Error Log")
    Major.SetFocus
Exit Sub
Else
    Dim db As Database, rs As Recordset
    Set db = OpenDatabase(nwi + "incident.mdb")
    Set rs = db.OpenRecordset("select * from pgroup")
    rs.AddNew
    rs("major") = Major.Text
    rs("minor") = Minor.Text
    rs.Update
End If
Call refreshroutine
End Sub
Private Sub refreshroutine()
Major.clear
Minor.clear
Dim db As Database
Dim rs As Recordset
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("select * from pgroup order by major,minor")
rs.MoveFirst
If Not rs.EOF Then
Dim holdmajor As String
holdmajor = ""
    While Not rs.EOF
        If Not rs("Major") = holdmajor Then
            Major.AddItem (rs("Major"))
            holdmajor = rs("Major")
        End If
            Minor.AddItem (rs("Minor"))
        rs.MoveNext
    Wend
End If
Major = ""
Minor = ""
db.Close
End Sub
'add refresh after delete
Private Sub cmdDelete_Click()
If Major = "" Then
    msg = MsgBox("Please select the item you wish to delete.", 48, "Genesis Information Log")
    Major.SetFocus
    Exit Sub
End If
msg = MsgBox("Are you sure you want to delete this property type?", 4, "Genesis Information Log")
If msg = 6 Then
    Dim db As Database, rs As Recordset
    Set db = OpenDatabase(nwi + "incident.mdb")
    If Minor.Text > "" Then
        Set rs = db.OpenRecordset("select * from pgroup where major = '" + Major.Text + "' and minor = '" + Minor.Text + "'")
    Else
        Set rs = db.OpenRecordset("select * from pgroup where major = '" + Major.Text + "' and (minor is null or minor = '')")
    End If
    If Not rs.EOF Then
        rs.MoveFirst
        rs.Delete
    End If
End If
    db.Close
Call refreshroutine

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set EditPropType = Nothing
End Sub

Private Sub major_Click()
Dim db As Database, rs As Recordset
Minor.clear
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("select minor from pgroup where major = '" + Major.Text + "'")
rs.MoveFirst
If Not rs.EOF Then
    While Not rs.EOF
        Minor.AddItem (rs("minor"))
        rs.MoveNext
    Wend
End If
db.Close

End Sub

Private Sub Form_Load()
Call refreshroutine
End Sub

Private Sub cmdClear_Click()
Major = ""
Minor = ""
Major.SetFocus
End Sub

