VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form service 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Genesis Service Call Report"
   ClientHeight    =   7440
   ClientLeft      =   570
   ClientTop       =   1560
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   9630
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Height          =   2925
      Left            =   30
      TabIndex        =   41
      Top             =   4500
      Width           =   9570
      Begin VB.CommandButton Spellck 
         Caption         =   "Spell Check"
         Height          =   255
         Left            =   2640
         TabIndex        =   47
         Top             =   525
         Width           =   975
      End
      Begin VB.TextBox received 
         Height          =   285
         Left            =   1140
         TabIndex        =   24
         Top             =   165
         Width           =   1935
      End
      Begin VB.TextBox completed 
         Height          =   285
         Left            =   7380
         TabIndex        =   25
         Top             =   165
         Width           =   1935
      End
      Begin VB.TextBox comments 
         Height          =   975
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   795
         Width           =   9255
      End
      Begin VB.ListBox approvingofficer 
         Height          =   840
         Left            =   4875
         TabIndex        =   29
         Top             =   1995
         Width           =   2415
      End
      Begin VB.ListBox reportingofficer 
         Height          =   840
         Left            =   60
         TabIndex        =   27
         Top             =   1995
         Width           =   2415
      End
      Begin VB.TextBox reportingofficernumber 
         Height          =   285
         Left            =   2940
         MaxLength       =   10
         TabIndex        =   28
         Top             =   1995
         Width           =   855
      End
      Begin VB.TextBox approvingofficernumber 
         Height          =   285
         Left            =   7740
         MaxLength       =   10
         TabIndex        =   30
         Top             =   1995
         Width           =   735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "TIME CALL RECEIVED"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   375
         Left            =   60
         TabIndex        =   46
         Top             =   165
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "TIME CALL COMPLETED"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   375
         Left            =   6180
         TabIndex        =   45
         Top             =   165
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "COMMENTS (If Necessary)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   60
         TabIndex        =   44
         Top             =   555
         Width           =   2775
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "APPROVING OFFICER                NUMBER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   4860
         TabIndex        =   43
         Top             =   1785
         Width           =   3855
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "REPORTING OFFICER                NUMBER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   60
         TabIndex        =   42
         Top             =   1785
         Width           =   3855
      End
   End
   Begin VB.ComboBox casenumber 
      Height          =   315
      Left            =   4395
      TabIndex        =   0
      Top             =   405
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   2190
      Left            =   45
      TabIndex        =   34
      Top             =   2325
      Width           =   9585
      Begin VB.TextBox zipcode 
         Height          =   285
         Left            =   5835
         MaxLength       =   10
         TabIndex        =   22
         Top             =   1830
         Width           =   1020
      End
      Begin VB.TextBox state 
         Height          =   285
         Left            =   5265
         MaxLength       =   2
         TabIndex        =   21
         Top             =   1845
         Width           =   480
      End
      Begin VB.TextBox currentdate 
         Height          =   285
         Left            =   2145
         MaxLength       =   10
         TabIndex        =   14
         Top             =   165
         Width           =   1455
      End
      Begin VB.TextBox incidentlocation 
         Height          =   285
         Left            =   2145
         MaxLength       =   100
         TabIndex        =   15
         Top             =   495
         Width           =   4215
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00808000&
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   45
         TabIndex        =   35
         Top             =   1005
         Width           =   2160
         Begin VB.OptionButton subject 
            BackColor       =   &H00808000&
            Caption         =   "Subject"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1230
            TabIndex        =   17
            Top             =   15
            Width           =   1695
         End
         Begin VB.OptionButton complainant 
            BackColor       =   &H00808000&
            Caption         =   "Complainant"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   15
            TabIndex        =   16
            Top             =   15
            Width           =   1215
         End
      End
      Begin VB.TextBox address1 
         Height          =   285
         Left            =   2295
         MaxLength       =   50
         TabIndex        =   19
         Top             =   1485
         Width           =   4575
      End
      Begin VB.ComboBox compsubj 
         Height          =   315
         Left            =   2295
         TabIndex        =   18
         Top             =   990
         Width           =   4095
      End
      Begin VB.TextBox address2 
         Height          =   285
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   20
         Top             =   1830
         Width           =   2910
      End
      Begin VB.TextBox phone 
         Height          =   285
         Left            =   7335
         MaxLength       =   30
         TabIndex        =   23
         Top             =   1830
         Width           =   1935
      End
      Begin VB.Image mugshot 
         BorderStyle     =   1  'Fixed Single
         Height          =   1230
         Left            =   8070
         Stretch         =   -1  'True
         Top             =   405
         Width           =   1395
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "CURRENT DATE:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   135
         TabIndex        =   39
         Top             =   165
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "INCIDENT LOCATION:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   105
         TabIndex        =   38
         Top             =   495
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   1200
         TabIndex        =   37
         Top             =   1470
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "PHONE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   7320
         TabIndex        =   36
         Top             =   1590
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "TYPE OF SERVICE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1530
      Left            =   60
      TabIndex        =   32
      Top             =   795
      Width           =   9480
      Begin VB.CheckBox alarm 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Alarm"
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   60
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox unlocking 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Unlocking Vehicle"
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   60
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
      Begin VB.CheckBox property 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Property Check"
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   75
         TabIndex        =   3
         Top             =   960
         Width           =   1815
      End
      Begin VB.CheckBox escort 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Escort (Specify)"
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   2460
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   2670
         TabIndex        =   33
         Top             =   600
         Width           =   3255
         Begin VB.OptionButton FUNERAL 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Funeral"
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   -15
            TabIndex        =   5
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton mental 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Mental Transport"
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox escortother 
            Height          =   285
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   8
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton house 
            BackColor       =   &H00C0C0C0&
            Caption         =   "House Moving"
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   1800
            TabIndex        =   6
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.CheckBox warrant 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Warrant Number"
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   6180
         TabIndex        =   9
         Top             =   240
         Width           =   1515
      End
      Begin VB.TextBox warrantnumber 
         Height          =   285
         Left            =   7710
         MaxLength       =   20
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox unfounded 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Unfounded (Specify in Comments)"
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   6180
         TabIndex        =   11
         Top             =   600
         Width           =   3255
      End
      Begin VB.CheckBox other 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Other (Specify)"
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   6195
         TabIndex        =   12
         Top             =   975
         Width           =   3255
      End
      Begin VB.TextBox otherspecify 
         Height          =   285
         Left            =   6165
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1215
         Width           =   3255
      End
      Begin Crystal.CrystalReport REPORT 
         Left            =   8310
         Top             =   -105
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         ReportFileName  =   "SERVICE"
         Destination     =   1
         PrintFileLinesPerPage=   60
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   7110
         Top             =   -105
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "service.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "service.frx":0454
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "service.frx":08A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "service.frx":0CFC
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "service.frx":1150
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "service.frx":15A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "service.frx":19F8
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "service.frx":1E4C
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "service.frx":22A0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "service.frx":26F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "service.frx":2B48
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   741
      ButtonWidth     =   1773
      ButtonHeight    =   582
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save   "
            Object.ToolTipText     =   "Save Case Number"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clear   "
            Object.ToolTipText     =   "Clear All Fields"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete   "
            Object.ToolTipText     =   "Delete Case Number"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print   "
            Object.ToolTipText     =   "Print Service Call Report"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit   "
            Object.ToolTipText     =   "Exit Service Call Repor"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CASE NUMBER:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   2775
      TabIndex        =   40
      Top             =   465
      Width           =   1455
   End
End
Attribute VB_Name = "service"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FROMXREF As Integer, nametype As Integer

Private Sub address1_GotFocus()
Dim db As Database, rs As Recordset
On Error Resume Next
If Address1 = "" And Address2 = "" And State = "" And zipcode = "" And compsubj > "" Then
    On Error GoTo oderror
od:
    Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
    Set rs = db.OpenRecordset("SELECT * FROM PEOPLE WHERE DPNAMElf = " + Chr$(34) + compsubj + Chr$(34))
    If Not rs.EOF Then
        rs.MoveFirst
        Address1 = rs("DPHADDRESS")
        Address2 = rs("DPHADDRESS2")
        If phone = "" Then
            phone = rs("DPHPHONE")
        End If
        If Not IsNull(rs("hstate")) Then
            State = rs("hstate")
        End If
        If Not IsNull(rs("hzipcode")) Then
            zipcode = rs("hzipcode")
        End If
        If Not IsNull(rs("mugshot")) Then
            mugshot.Picture = LoadPicture(rs("mugshot"))
        Else
            mugshot.Picture = LoadPicture()
        End If
    End If
End If
On Error Resume Next
db.Close
On Error GoTo 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
    
End Sub

Private Sub approvingofficernumber_GotFocus()
If approvingofficernumber > "" Or approvingofficer = "" Then
    Exit Sub
End If
Dim db As DAO.Database, rs As DAO.Recordset
Set db = DAO.OpenDatabase(nwl + "LAWSUITE.MDB")
Set rs = db.OpenRecordset("SELECT PROFIDNUM FROM PROFESSIONALS WHERE PROFNAME = '" + approvingofficer + "' AND TYPE = 'D'")
If Not rs.EOF Then
    rs.MoveFirst
    If Not IsNull(rs("PROFIDNUM")) Then
        approvingofficernumber = rs("PROFIDNUM")
    End If
End If
db.Close
End Sub

Private Sub casenumber_Change()
If FROMXREF = 1 Then
    Call findrtn
    FROMXREF = 0
End If
End Sub

Private Sub CaseNumber_Click()
Call findrtn
casenumber.Refresh
End Sub

Private Sub Comments_KeyDown(KeyCode As Integer, Shift As Integer)
If (Shift = vbCtrlMask) And (KeyCode = vbKeyF2) Then
        Call SpellCk_Click
End If
End Sub

Private Sub completed_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(completed) = 1 Then
    SendKeys ":"
End If
End If

End Sub

Private Sub completed_LostFocus()
If completed > "" And Not IsDate(completed) Then
    msg = MsgBox("Invalid time entered.", 48, "Genesis Error Log")
    completed.SetFocus
End If

End Sub

Private Sub compsubj_Click()
If compsubj = "" Then
    Exit Sub
End If
Call setpopup(compsubj, "L")
End Sub

Private Sub compsubj_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 And Shift = vbCtrlMask Then
    If nametype = 0 Then
        cleanup.fname = Me.ActiveControl.Text
        cleanup.lname = ""
    Else
        cleanup.lname = Me.ActiveControl.Text
        cleanup.fname = ""
    End If
    cleanup.Show
End If
End Sub

Private Sub compsubj_LostFocus()
If compsubj > "" And InStr(compsubj, ",") = 0 Then
    msg = MsgBox("All names in the Service Call system should be entered in the format last name + comma + firstname.", 48, "Invalid Data Format")
    compsubj.SetFocus
End If

End Sub

Private Sub currentdate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(currentdate) = 1 Or Len(currentdate) = 4 Then
    SendKeys "/"
End If
End If


End Sub

Private Sub escort_LostFocus()
If escort = 0 Then
    warrant.SetFocus
End If
End Sub

Private Sub Form_Load()
nametype = 1
For t% = 0 To Forms.Count - 1
    If Forms(t%).Name = "xref" Then
        FROMXREF = 1
        t% = Forms.Count - 1
    End If
Next t%
On Error Resume Next
Call clearrtn
Me.Top = 0
Me.Left = 0
Me.Height = 7850
Me.Width = 9750
End Sub

Private Sub loadcase()
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nws + "service.mdb")
Set rs = db.OpenRecordset("select casenumber from service order by casenumber")
If Not rs.EOF Then
    rs.MoveFirst
End If
casenumber.clear
While Not rs.EOF
    casenumber.AddItem rs("casenumber")
    rs.MoveNext
Wend
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
Private Sub clearrtn()
mugshot.Picture = LoadPicture()
casenumber = ""
alarm = 0
unlocking = 0
property = 0
escort = 0
FUNERAL = False
house = False
mental = False
escortother = ""
warrant = 0
warrantnumber = ""
unfounded = 0
Other = 0
otherspecify = ""
currentdate = Format$(Date$, "mm/dd/yyyy")
incidentlocation = ""
complainant = False
subject = False
compsubj = ""
Address1 = ""
Address2 = ""
State = ""
zipcode = ""
phone = ""
received = ""
completed = ""
comments = ""
reportingofficer.ListIndex = -1
reportingofficernumber = ""
approvingofficer.ListIndex = -1
approvingofficernumber = ""
Call loadcase
Call loadofficer
Call loadname

End Sub
Private Sub loadofficer()
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "lawsuite.mdb")
Set rs = db.OpenRecordset("select profname from professionals where type = 'D' order by profname")
If Not rs.EOF Then
    rs.MoveFirst
End If
reportingofficer.clear
approvingofficer.clear
While Not rs.EOF
    reportingofficer.AddItem rs("profname")
    approvingofficer.AddItem rs("profname")
    rs.MoveNext
Wend
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
Private Sub loadname()
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "LAWSUITE.mdb")
Set rs = db.OpenRecordset("select DPnamelf from people order by DPnamElf")
If Not rs.EOF Then
    rs.MoveFirst
End If
compsubj.clear
While Not rs.EOF
    compsubj.AddItem rs("DPnamelf")
    rs.MoveNext
Wend
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

Private Sub Form_LostFocus()
goingelsewhere = False
End Sub

Private Sub Form_Paint()
SetAlwaysOnTop Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
goingelsewhere = False
Set service = Nothing
End Sub

Private Sub received_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(received) = 1 Then
    SendKeys ":"
End If
End If

End Sub

Private Sub received_LostFocus()
If received > "" And Not IsDate(received) Then
    msg = MsgBox("Invalid time entered.", 48, "Genesis Error Log")
    received.SetFocus
End If
End Sub

Private Sub reportingofficernumber_GotFocus()
If reportingofficernumber > "" Or reportingofficer = "" Then
    Exit Sub
End If
Dim db As DAO.Database, rs As DAO.Recordset
Set db = DAO.OpenDatabase(nwl + "LAWSUITE.MDB")
Set rs = db.OpenRecordset("SELECT PROFIDNUM FROM PROFESSIONALS WHERE PROFNAME = '" + reportingofficer + "' AND TYPE = 'D'")
If Not rs.EOF Then
    rs.MoveFirst
    If Not IsNull(rs("PROFIDNUM")) Then
        reportingofficernumber = rs("PROFIDNUM")
    End If
End If
db.Close
End Sub

Private Sub SpellCk_Click()
BeginSpellCheck comments.Text, comments
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button
    Case "Save   "
        If UCase(frmLogin.txtUserName) = "DEMO" And UCase(frmLogin.txtPassword) = "DEMO" Then
            msg = MsgBox("Not available in DEMO version.", 48, "Genesis Information Log")
            Screen.MousePointer = 0
            Exit Sub
        End If
        Screen.MousePointer = 11
        Call savertn
        Screen.MousePointer = 0
    Case "Clear   "
        Screen.MousePointer = 11
        Call clearrtn
        Screen.MousePointer = 0
    Case "Delete   "
        If UCase(frmLogin.txtUserName) = "DEMO" And UCase(frmLogin.txtPassword) = "DEMO" Then
            msg = MsgBox("Not available in DEMO version.", 48, "Genesis Information Log")
            Screen.MousePointer = 0
            Exit Sub
        End If
        Screen.MousePointer = 11
        Call deletertn
        Screen.MousePointer = 0
    Case "Print   "
        Screen.MousePointer = 11
        Call printrtn
        Screen.MousePointer = 0
    Case "Exit   "
        Unload service
End Select
End Sub
Private Sub savertn()
Dim db As Database, rs As Recordset
If casenumber = "" Then
    msg = MsgBox("A case number must be entered.", 48, "Genesis Error Log")
    Exit Sub
End If
If Len(casenumber) > 20 Then
    msg = MsgBox("Case Number cannot be over 20 characters in length.", 48, "Genesis Error Log")
    Exit Sub
End If
If Len(compsubj) > 50 Then
    msg = MsgBox("Complainant/Subject Name cannot be over 50 characters in length.", 48, "Genesis Error Log")
    Exit Sub
End If
If alarm = 0 And unlocking = 0 And property = 0 And escort = 0 And warrant = 0 And unfounded = 0 And Other = 0 Then
    msg = MsgBox("A TYPE OF SERVICE must be checked.", 48, "Genesis Error Log")
    Exit Sub
End If
If escort = 1 Then
    If Not FUNERAL And Not mental And Not house And escortother = "" Then
        msg = MsgBox("If ESCORT is checked, a futher specification must be selected or entered.", 48, "Genesis Error Log")
        Exit Sub
    End If
End If
If warrant = 1 And warrantnumber = "" Then
    msg = MsgBox("If WARRANT is checked, a valid warrant number must be entered.", 48, "Genesis Error Log")
    Exit Sub
End If
If Other = 1 And otherspecify = "" Then
    msg = MsgBox("If OTHER is checked, a further specification must be entered.", 48, "Genesis Error Log")
    Exit Sub
End If
If Not IsDate(currentdate) Then
    msg = MsgBox("CURRENT DATE is not a valid date.", 48, "Genesis Error Log")
    Exit Sub
End If
If compsubj > "" And Not complainant And Not subject Then
    msg = MsgBox("Either COMPLAINANT or SUBJECT must be selected when a name is entered.", 48, "Genesis Error Log")
    Exit Sub
End If
If received > "" And Not IsDate(received) Then
    msg = MsgBox("TIME CALL RECEIVED is not a valid time.", 48, "Genesis Error Log")
    Exit Sub
End If
If completed > "" And Not IsDate(completed) Then
    msg = MsgBox("TIME CALL completed is not a valid time.", 48, "Genesis Error Log")
    Exit Sub
End If
If reportingofficer.ListIndex = -1 And approvingofficer.ListIndex = -1 Then
    msg = MsgBox("Either a reporting or approving officer must be selected.", 48, "Genesis Error Log")
    Exit Sub
End If
On Error GoTo oderror
od:
Set db = OpenDatabase(nws + "service.mdb")
Set rs = db.OpenRecordset("select * from service where casenumber =" + Chr$(34) + casenumber + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    rs.Edit
Else
    rs.AddNew
End If
rs("casenumber") = casenumber
rs("alarm") = alarm
rs("unlocking") = unlocking
rs("property") = property
rs("escort") = escort
rs("funeral") = FUNERAL
rs("house") = house
rs("mental") = mental
rs("escortother") = escortother
rs("warrant") = warrant
rs("warrantnumber") = warrantnumber
rs("unfounded") = unfounded
rs("other") = Other
rs("otherspecify") = otherspecify
rs("currentdate") = currentdate
rs("incidentlocation") = incidentlocation
rs("complainant") = complainant
rs("subject") = subject
rs("compsubj") = compsubj
rs("address1") = Address1
rs("address2") = Address2
rs("state") = State
rs("zipcode") = zipcode
rs("phone") = phone
If received = "" Then
    rs("RECEIVED") = Null
Else
    rs("received") = received
End If
If completed = "" Then
    rs("COMPLETED") = Null
Else
    rs("completed") = completed
End If
rs("comments") = comments
rs("approvingofficer") = approvingofficer.List(approvingofficer.ListIndex)
rs("reportingofficer") = reportingofficer.List(reportingofficer.ListIndex)
rs("approvingofficernumber") = approvingofficernumber
rs("reportingofficernumber") = reportingofficernumber
'CES Code
rs("userfullname") = frmLogin.userfullname
rs("userid") = frmLogin.userid
rs("ORINUMBER") = frmLogin.orinumber
rs("udate") = Format$(Now, "mm/dd/yyyy")
rs("utime") = Format$(Now, "hh:mm:ss")
'********
rs.Update
On Error Resume Next
Set db = OpenDatabase(nwl + "lawsuite.mdb")
'----- OFFICERS
If reportingofficer > "" Then
    Set rs = db.OpenRecordset("select profidnum,profname, type from professionals where profname =" + Chr$(34) + reportingofficer + Chr$(34))
    If rs.EOF Then
        rs.AddNew
    Else
        rs.MoveFirst
        rs.Edit
    End If
    rs("profname") = reportingofficer
    rs("profidnum") = reportingofficernumber
    If rs.EOF Then
        reportingofficer.AddItem reportingofficer
        approvingofficer.AddItem reportingofficer
        followupofficer.AddItem reportingofficer
    End If
    rs("type") = "D"
    rs.Update
End If
If approvingofficer > "" Then
    Set rs = db.OpenRecordset("select profidnum,profname, type from professionals where profname =" + Chr$(34) + approvingofficer + Chr$(34))
    If rs.EOF Then
        rs.AddNew
    Else
        rs.MoveFirst
        rs.Edit
    End If
    rs("profname") = approvingofficer
    rs("profidnum") = approvingofficernumber
    If rs.EOF Then
        reportingofficer.AddItem approvingofficer
        approvingofficer.AddItem approvingofficer
        followupofficer.AddItem approvingofficer
    End If
    rs("type") = "D"
    rs.Update
End If
Set rs = db.OpenRecordset("select * from people where dpnamelf = " + Chr$(34) + compsubj + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    rs.Edit
Else
    rs.AddNew
End If
rs("dpnamelf") = compsubj
rs("dphaddress") = Address1
rs("dphaddress2") = Address2
rs("dstate") = State
rs("dzipcode") = zipcode
rs("dpsort") = Left$(compsubj, 15)
If phone > "" Then
    rs("dphphone") = phone
End If
hoLdname = compsubj
osort1$ = ""
If Left$(hoLdname, 1) = " " Then
    hoLdname = Mid$(hoLdname, 2)
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
    If Mid$(tso$, tt%, 1) = "," Then
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
If Left$(tempsort$, 1) = " " Then
    tempsort$ = Mid$(tempsort$, 2)
End If
tso$ = Left$(tso$, firstspace% - 1)
If Right$(tso$, 1) = " " Then
    tso$ = Left$(tso$, Len(tso$) - 1)
End If
tempsort$ = tempsort$ + " " + tso$
If osort1$ = "" Then
    osort1$ = tempsort$
End If
If InStr(osort1$, "JR.") Then
    If Mid$(osort1$, InStr(osort1$, "JR.") + 3, 1) = " " Then
        osort1$ = Left$(osort1$, InStr(osort1$, "JR.") - 1) + Mid$(osort1$, InStr(osort1$, "JR.") + 4) + ", JR."
    Else
        osort1$ = Left$(osort1$, InStr(osort1$, "JR.") - 1) + Mid$(osort1$, InStr(osort1$, "JR.") + 3) + ", JR."
End If
End If
If InStr(osort1$, "SR.") Then
    If Mid$(osort1$, InStr(osort1$, "SR.") + 3, 1) = " " Then
        osort1$ = Left$(osort1$, InStr(osort1$, "SR.") - 1) + Mid$(osort1$, InStr(osort1$, "SR.") + 4) + ", SR."
    Else
        osort1$ = Left$(osort1$, InStr(osort1$, "SR.") - 1) + Mid$(osort1$, InStr(osort1$, "SR.") + 3) + ", SR."
End If
End If
If InStr(osort1$, "III") Then
    If Mid$(osort1$, InStr(osort1$, "III") + 3, 1) = " " Then
        osort1$ = Left$(osort1$, InStr(osort1$, "III") - 1) + Mid$(osort1$, InStr(osort1$, "III") + 4) + ", III"
    Else
        osort1$ = Left$(osort1$, InStr(osort1$, "III") - 1) + Mid$(osort1$, InStr(osort1$, "III") + 3) + ", III"
    End If
End If
If InStr(osort1$, "IV") Then
    If Mid$(osort1$, InStr(osort1$, "IV") + 2, 1) = " " Then
        osort1$ = Left$(osort1$, InStr(osort1$, "IV") - 1) + Mid$(osort1$, InStr(osort1$, "IV") + 3) + ", III"
    Else
        osort1$ = Left$(osort1$, InStr(osort1$, "IV") - 1) + Mid$(osort1$, InStr(osort1$, "IV") + 2) + ", III"
    End If
End If
If Left$(osort1$, 1) = " " Then
    osort1$ = Mid$(osort1$, 2)
End If
rsupdate:
rs("dpname") = osort1$
rs.Update


db.Close
Call clearrtn
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub
Private Sub findrtn()
If casenumber = "" Then
    msg = MsgBox("A valid case number must be entered.", 48, "Genesis Error Log")
    Exit Sub
End If
On Error Resume Next
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nws + "service.mdb")
Set rs = db.OpenRecordset("select * from service where casenumber =" + Chr$(34) + casenumber + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
Else
    On Error Resume Next
    db.Close
    msg = MsgBox("Case Number not found.", 48, "Genesis Information Log")
    Exit Sub
End If
casenumber = rs("casenumber")
alarm = rs("alarm")
unlocking = rs("unlocking")
property = rs("property")
escort = rs("escort")
FUNERAL = rs("funeral")
house = rs("house")
mental = rs("mental")
escortother = rs("escortother")
warrant = rs("warrant")
warrantnumber = rs("warrantnumber")
unfounded = rs("unfounded")
Other = rs("other")
otherspecify = rs("otherspecify")
currentdate = rs("currentdate")
incidentlocation = rs("incidentlocation")
complainant = rs("complainant")
subject = rs("subject")
compsubj = rs("compsubj")
Address1 = rs("address1")
Address2 = rs("address2")
If Not IsNull(rs("state")) Then
    State = rs("state")
End If
If Not IsNull(rs("zipcode")) Then
    zipcode = rs("zipcode")
End If
phone = rs("phone")
received = rs("received")
completed = rs("completed")
comments = rs("comments")
approvingofficer.ListIndex = -1
For t% = 0 To approvingofficer.ListCount - 1
    If rs("approvingofficer") = approvingofficer.List(t%) Then
        approvingofficer.ListIndex = t%
        t% = approvingofficer.ListCount - 1
    End If
Next t%
reportingofficer.ListIndex = -1
For t% = 0 To reportingofficer.ListCount - 1
    If rs("reportingofficer") = reportingofficer.List(t%) Then
        reportingofficer.ListIndex = t%
        t% = reportingofficer.ListCount - 1
    End If
Next t%
approvingofficernumber = rs("approvingofficernumber")
reportingofficernumber = rs("reportingofficernumber")
Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
Set rs = db.OpenRecordset("SELECT * FROM PEOPLE WHERE DPNAMElf = " + Chr$(34) + compsubj + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    If Not IsNull(rs("mugshot")) Then
        mugshot.Picture = LoadPicture(rs("mugshot"))
    Else
        mugshot.Picture = LoadPicture()
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
Private Sub deletertn()
If casenumber = "" Then
    msg = MsgBox("A valid case number must be entered.", 48, "Genesis Error Log")
    Exit Sub
End If
msg = MsgBox("Are you sure you want to delete this record?", 4, "Genesis Information Log")
If msg = 7 Then
    Exit Sub
End If
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nws + "service.mdb")
Set rs = db.OpenRecordset("select * from service where casenumber =" + Chr$(34) + casenumber + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    rs.Delete
End If
db.Close
On Error Resume Next
Call clearrtn
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub
Private Sub printrtn()
REPORT.ReportFileName = nws + "SERVICE.RPT"
REPORT.SelectionFormula = "{SERVICE.CASENUMBER} = '" + casenumber + "'"
REPORT.Action = 1
End Sub
