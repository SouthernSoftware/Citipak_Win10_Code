Attribute VB_Name = "Module1"
Public fMainForm As frmMain, tuser, tpassword, tauthority As String


Sub Main()
    Dim fLogin As New frmLogin
    fLogin.Show vbModal
    If Not fLogin.OK Then
        'Login Failed so exit app
        End
    End If
    Unload fLogin


    frmSplash.Show
    frmSplash.Refresh
    Set fMainForm = New frmMain
    Load fMainForm
    Unload frmSplash


    fMainForm.Show
End Sub



Sub getsecurity(tuser, tpassword, tauthority As String)
Dim db As Database, rs As Recordset, t As Integer, temppass As String
Set db = OpenDatabase("enforce.mdb")
Set rs = db.OpenRecordset("select * from security where user = '" + tuser + "'")
tpassword = ""
tauthority = ""
If rs.EOF Then
    Exit Sub
End If
rs.movefirst
For t = 1 To Len(rs("password"))
    tpassword = tpassword + Chr$(Asc(Mid$(rs("password"), t, 1)) + 32)
Next t
tauthority = rs("authority")
db.Close
End Sub
