VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Populate Rap Sheet"
   ClientHeight    =   5715
   ClientLeft      =   2700
   ClientTop       =   1545
   ClientWidth     =   3420
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   3420
   Begin VB.CommandButton Command1 
      Caption         =   "Process Rap Sheet"
      Height          =   735
      Left            =   840
      TabIndex        =   1
      Top             =   3600
      Width           =   1575
   End
   Begin MSComctlLib.ListView origination 
      Height          =   2895
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   5106
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Origination"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   4560
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nww, nwl As String
Dim dbw, dbr As Database, rsw, rsr As Recordset, itmx As ListItem
Dim orig(100) As String, oridx As Integer

Private Sub Command1_Click()
Set dbw = OpenDatabase(nww + "warrant.mdb")
Set dbr = OpenDatabase(nwl + "rapsheet.mdb")
oridx = 0
For t% = 1 To origination.ListItems.Count
    If origination.ListItems(t%).Selected = True Then
        oridx = oridx + 1
        orig(oridx) = origination.ListItems(t%)
    End If
Next t%
Set rsw = dbw.openrecordset("select * from warrant where whenarrested is not null")
rsw.movelast
maxct = rsw.recordcount
rsw.movefirst
ct = 0
While Not rsw.EOF
    ct = ct + 1
    Label1 = CStr(ct) + " of " + CStr(maxct)
    Label1.Refresh
    If Not IsNull(rsw("county")) Then
        If rsw("county") = True Then
            GoSub poprapsheet
        End If
    Else
    If Not IsNull(rsw("origination")) Then
        foundit = False
        For t% = 1 To oridx
            If rsw("origination") = orig(t%) Then
                foundit = True
                t% = oridx
            End If
        Next t%
        If foundit = True Then
            GoSub poprapsheet
        End If
    End If
    End If
    rsw.movenext
Wend
dbw.Close
dbr.Close
Exit Sub
poprapsheet:

Set rsr = dbr.openrecordset("select * from rapsheet")
rsr.AddNew
rsr("lname") = rsw("wname")
rsr("ssn") = rsw("ssn")
rsr("idnumber") = rsw("idnumber")
If Not IsNull(rsw("birthdate")) Then
    rsr("birthdate") = rsw("birthdate")
End If
rsr("arrestdate") = rsw("whenarrested")
rsr("casenumber") = rsw("casenumber")
rsr("warrantnumber") = rsw("warrant")
rsr("charge") = rsw("charge")
rs.Update
Return

End Sub

Private Sub Form_Load()
Open "c:\law enforcement suite\nww.ini" For Input As #1
Line Input #1, a$
Close #1
nww = a$
Open "c:\law enforcement suite\nwl.ini" For Input As #1
Line Input #1, a$
Close #1
nwl = a$
Set dbw = OpenDatabase(nww + "warrant.mdb")
Set rsw = dbw.openrecordset("select distinct origination from warrantinfo where origination is not null and (county is null or county = false) and whenarrested is not null")
While Not rsw.EOF
    Set itmx = origination.ListItems.Add(, , rsw("origination"))
    rsw.movenext
Wend
dbw.Close
End Sub
