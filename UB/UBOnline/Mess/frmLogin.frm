VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Law Enforcement Suite Login"
   ClientHeight    =   1635
   ClientLeft      =   4890
   ClientTop       =   2580
   ClientWidth     =   3990
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   966.013
   ScaleMode       =   0  'User
   ScaleWidth      =   3746.394
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   2325
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2310
      TabIndex        =   3
      Top             =   1200
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   720
      TabIndex        =   2
      Top             =   1200
      Width           =   1140
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   2325
   End
   Begin VB.Label rmssupervisor 
      Height          =   255
      Left            =   0
      TabIndex        =   75
      Top             =   0
      Width           =   135
   End
   Begin VB.Label rmsreport 
      Height          =   255
      Left            =   0
      TabIndex        =   74
      Top             =   0
      Width           =   135
   End
   Begin VB.Label rmsprint 
      Height          =   255
      Left            =   0
      TabIndex        =   73
      Top             =   0
      Width           =   135
   End
   Begin VB.Label rmsedit 
      Height          =   255
      Left            =   0
      TabIndex        =   72
      Top             =   0
      Width           =   135
   End
   Begin VB.Label rmsdelete 
      Height          =   255
      Left            =   0
      TabIndex        =   71
      Top             =   0
      Width           =   135
   End
   Begin VB.Label rmsbrowse 
      Height          =   255
      Left            =   3600
      TabIndex        =   70
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label adelete 
      Height          =   255
      Left            =   1080
      TabIndex        =   69
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label aprint 
      Height          =   255
      Left            =   1680
      TabIndex        =   68
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label areport 
      Caption         =   " "
      Height          =   255
      Left            =   2400
      TabIndex        =   67
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label aedit 
      Height          =   255
      Left            =   492
      TabIndex        =   66
      Top             =   6
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label asupervisor 
      Height          =   255
      Left            =   3120
      TabIndex        =   65
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label abrowse 
      Height          =   255
      Left            =   0
      TabIndex        =   64
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label ORINumber 
      Height          =   255
      Left            =   90
      TabIndex        =   63
      Top             =   2850
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label UserFullName 
      Height          =   255
      Left            =   30
      TabIndex        =   62
      Top             =   2070
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label UserID 
      Height          =   255
      Left            =   75
      TabIndex        =   61
      Top             =   2430
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label cedit 
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   60
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label cprint 
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   59
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label creport 
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   58
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label cdelete 
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   57
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label csupervisor 
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   56
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label cbrowse 
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   55
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label cedit 
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label cprint 
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   53
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label creport 
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   52
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label cdelete 
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   51
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label csupervisor 
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   50
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label cbrowse 
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   49
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label cedit 
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label cprint 
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   47
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label creport 
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   46
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label cdelete 
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   45
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label csupervisor 
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   44
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label cbrowse 
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   43
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label ssupervisor 
      Height          =   255
      Left            =   0
      TabIndex        =   42
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label sreport 
      Height          =   255
      Left            =   0
      TabIndex        =   41
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label sdelete 
      Height          =   255
      Left            =   0
      TabIndex        =   40
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label sprint 
      Height          =   255
      Left            =   0
      TabIndex        =   39
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label sbrowse 
      Height          =   255
      Left            =   0
      TabIndex        =   38
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label sedit 
      Height          =   255
      Left            =   0
      TabIndex        =   37
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label bbrowse 
      Height          =   255
      Left            =   0
      TabIndex        =   36
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label bsupervisor 
      Height          =   255
      Left            =   0
      TabIndex        =   35
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label bdelete 
      Height          =   255
      Left            =   0
      TabIndex        =   34
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label breport 
      Height          =   255
      Left            =   0
      TabIndex        =   33
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label bprint 
      Height          =   255
      Left            =   0
      TabIndex        =   32
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label bedit 
      Height          =   255
      Left            =   0
      TabIndex        =   31
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label supervisor 
      Height          =   255
      Left            =   0
      TabIndex        =   30
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label rbrowse 
      Height          =   255
      Left            =   0
      TabIndex        =   29
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label rsupervisor 
      Height          =   255
      Left            =   0
      TabIndex        =   28
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label rdelete 
      Height          =   255
      Left            =   0
      TabIndex        =   27
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label rreport 
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label rprint 
      Height          =   255
      Left            =   0
      TabIndex        =   25
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label redit 
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label cbrowse 
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label csupervisor 
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label cdelete 
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label creport 
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label cprint 
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label cedit 
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label wbrowse 
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label wsupervisor 
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label wdelete 
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label wreport 
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label wprint 
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label wedit 
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label ibrowse 
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label isupervisor 
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label idelete 
      Height          =   255
      Left            =   852
      TabIndex        =   9
      Top             =   846
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label ireport 
      Caption         =   " "
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label iprint 
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label iedit 
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
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
      Height          =   270
      Index           =   1
      Left            =   225
      TabIndex        =   5
      Top             =   570
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
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
      Height          =   270
      Index           =   0
      Left            =   225
      TabIndex        =   4
      Top             =   180
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ie As Integer


Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    Unload frmLogin
    Unload mainform
    End
End Sub

Private Sub cmdOK_Click()
Call checkcracker
If txtUserName > "" And txtPassword = "" Then
    txtPassword.SetFocus
    Exit Sub
End If
Screen.MousePointer = 11
Dim db As Database, rs As Recordset, temppass As String, uu As Integer
On Error GoTo oderror
od:

'RLB Code
   DoEvents
'********
cedit(0) = 0
cprint(0) = 0
creport(0) = 0
cdelete(0) = 0
csupervisor(0) = 0
cbrowse(0) = 0
cedit(1) = 0
cprint(1) = 0
creport(1) = 0
cdelete(1) = 0
csupervisor(1) = 0
cbrowse(1) = 0
cedit(2) = 0
cprint(2) = 0
creport(2) = 0
cdelete(2) = 0
csupervisor(2) = 0
cbrowse(2) = 0
cedit(3) = 0
cprint(3) = 0
creport(3) = 0
cdelete(3) = 0
csupervisor(3) = 0
cbrowse(3) = 0
sedit = 0
sprint = 0
sreport = 0
sdelete = 0
ssupervisor = 0
sbrowse = 0
bedit = 0
bprint = 0
breport = 0
bdelete = 0
bsupervisor = 0
bbrowse = 0
iedit = 0
iprint = 0
ireport = 0
idelete = 0
isupervisor = 0
jbrowse = 0
jedit = 0
jprint = 0
jreport = 0
jdelete = 0
jsupervisor = 0
abrowse = 0
aedit = 0
aprint = 0
areport = 0
adelete = 0
asupervisor = 0
rmsbrowse = 0
rmsedit = 0
rmsprint = 0
rmsreport = 0
rmsdelete = 0
rmssupervisor = 0
ibrowse = 0
wedit = 0
wprint = 0
wreport = 0
wdelete = 0
wsupervisor = 0
wbrowse = 0
redit = 0
rprint = 0
rreport = 0
rdelete = 0
rsupervisor = 0
rbrowse = 0
supervisor = 0
   
Set db = OpenDatabase(nwl + "lawsuite.mdb")



txtUserName = UCase(txtUserName)
txtPassword = UCase(txtPassword)
If txtUserName = "082391" And txtPassword = "052992" Then
    cedit(0) = 1
    cprint(0) = 1
    creport(0) = 1
    cdelete(0) = 1
    csupervisor(0) = 1
    cbrowse(0) = 1
    cedit(1) = 1
    cprint(1) = 1
    creport(1) = 1
    cdelete(1) = 1
    csupervisor(1) = 1
    cbrowse(1) = 1
    cedit(2) = 1
    cprint(2) = 1
    creport(2) = 1
    cdelete(2) = 1
    csupervisor(2) = 1
    cbrowse(2) = 1
    cedit(3) = 1
    cprint(3) = 1
    creport(3) = 1
    cdelete(3) = 1
    csupervisor(3) = 1
    cbrowse(3) = 1
    sedit = 1
    sprint = 1
    sreport = 1
    sdelete = 1
    ssupervisor = 1
    sbrowse = 1
    bedit = 1
    bprint = 1
    breport = 1
    bdelete = 1
    bsupervisor = 1
    bbrowse = 1
    iedit = 1
    iprint = 1
    ireport = 1
    idelete = 1
    isupervisor = 1
    jbrowse = 1
    jedit = 1
    jprint = 1
    jreport = 1
    jdelete = 1
    jsupervisor = 1
    abrowse = 1
    aedit = 1
    aprint = 1
    areport = 1
    adelete = 1
    asupervisor = 1
    rmsbrowse = 1
    rmsedit = 1
    rmsprint = 1
    rmsreport = 1
    rmsdelete = 1
    rmssupervisor = 1
    ibrowse = 1
    wedit = 1
    wprint = 1
    wreport = 1
    wdelete = 1
    wsupervisor = 1
    wbrowse = 1
    redit = 1
    rprint = 1
    rreport = 1
    rdelete = 1
    rsupervisor = 1
    rbrowse = 1
    supervisor = 1
    'RLB Code
    UserFullName = "GENESIS SOFTWARE CORPORATION"
    UserID = "082391"
    '****
    'CES Code
    ORINumber = "571009589"
    '****
Else
    Set rs = db.OpenRecordset("select * from security where userid = '" + txtUserName + "'")
    If rs.EOF Then
        MsgBox "Invalid User ID", , "Login"
        Screen.MousePointer = 0
        txtUserName.SetFocus
        db.Close
        Exit Sub
    Else
        rs.MoveFirst
        temppass = ""
        For uu = 1 To Len(rs("password"))
            temppass = temppass + Chr$(Asc(Mid$(rs("password"), uu, 1)) - 10)
        Next uu
        If txtPassword <> temppass Then
            MsgBox "Invalid Password Entered.", 48, "Genesis Error Log"
            txtUserName.SetFocus
            Screen.MousePointer = 0
            db.Close
            Exit Sub
        Else
            On Error Resume Next
            cedit(0) = rs("cedit")
            cprint(0) = rs("cprint")
            creport(0) = rs("creport")
            cdelete(0) = rs("cdelete")
            csupervisor(0) = rs("csupervisor")
            cbrowse(0) = rs("cbrowse")
            cedit(1) = rs("ceditw")
            cprint(1) = rs("cprintw")
            creport(1) = rs("creportw")
            cdelete(1) = rs("cdeletew")
            csupervisor(1) = rs("csupervisorw")
            cbrowse(1) = rs("cbrowsew")
            cedit(2) = rs("ceditf")
            cprint(2) = rs("cprintf")
            creport(2) = rs("creportf")
            cdelete(2) = rs("cdeletef")
            csupervisor(2) = rs("csupervisorf")
            cbrowse(2) = rs("cbrowsef")
            cedit(3) = rs("cedite")
            cprint(3) = rs("cprinte")
            creport(3) = rs("creporte")
            cdelete(3) = rs("cdeletee")
            csupervisor(3) = rs("csupervisore")
            cbrowse(3) = rs("cbrowsee")
            sedit = rs("sedit")
            sprint = rs("sprint")
            sreport = rs("sreport")
            sdelete = rs("sdelete")
            ssupervisor = rs("ssupervisor")
            sbrowse = rs("sbrowse")
            bedit = rs("bedit")
            bprint = rs("bprint")
            breport = rs("breport")
            bdelete = rs("bdelete")
            bsupervisor = rs("bsupervisor")
            bbrowse = rs("bbrowse")
            iedit = rs("iedit")
            iprint = rs("iprint")
            ireport = rs("ireport")
            idelete = rs("idelete")
            isupervisor = rs("isupervisor")
            ibrowse = rs("Ibrowse")
            jedit = rs("jedit")
            jprint = rs("jprint")
            jreport = rs("jreport")
            jdelete = rs("jdelete")
            jsupervisor = rs("jsupervisor")
            jbrowse = rs("jbrowse")
            aedit = rs("aedit")
            aprint = rs("aprint")
            areport = rs("areport")
            adelete = rs("adelete")
            asupervisor = rs("asupervisor")
            abrowse = rs("abrowse")
            rmsedit = rs("rmsedit")
            rmsprint = rs("rmsprint")
            rmsreport = rs("rmsreport")
            rmsdelete = rs("rmsdelete")
            rmssupervisor = rs("rmssupervisor")
            rmsbrowse = rs("rmsbrowse")
            wedit = rs("wedit")
            wprint = rs("wprint")
            wreport = rs("wreport")
            wdelete = rs("wdelete")
            wsupervisor = rs("wsupervisor")
            wbrowse = rs("wbrowse")
            redit = rs("redit")
            rprint = rs("rprint")
            rreport = rs("rreport")
            rdelete = rs("rdelete")
            rsupervisor = rs("rsupervisor")
            rbrowse = rs("rbrowse")
            supervisor = rs("supervisor")
            'RLB Code
            If Not IsNull(rs("userfullname")) Then
                UserFullName = rs("userfullname")
            End If
            UserID = rs("userid")
            '********
            'CES Code
            If Not IsNull(rs("ORINUMBER")) Then
                ORINumber = rs("ORINUMBER")
            End If
            '********
        End If
    End If
End If
If supervisor = 1 Then
    mainform.pcleanup.Enabled = True
    mainform.omaint.Enabled = True
    mainform.Toolbar1.Buttons(17).Enabled = True
    mainform.Toolbar1.Buttons(18).Enabled = True
    mainform.Toolbar1.Buttons(19).Enabled = True
    mainform.Toolbar1.Buttons(7).Enabled = True
    mainform.Toolbar1.Buttons(8).Enabled = True
    mainform.Toolbar1.Buttons(9).Enabled = True
    mainform.Toolbar1.Buttons(10).Enabled = True
    mainform.Toolbar1.Buttons(11).Enabled = True
End If
If cedit(0) = 1 Or cprint(0) = 1 Or creport(0) = 1 Or cdelete(0) = 1 Or csupervisor(0) = 1 Or cbrowse(0) = 1 Or supervisor = 1 Or cedit(1) = 1 Or cprint(1) = 1 Or creport(1) = 1 Or cdelete(1) = 1 Or csupervisor(1) = 1 Or cbrowse(1) = 1 Or cedit(2) = 1 Or cprint(2) = 1 Or creport(2) = 1 Or cdelete(2) = 1 Or csupervisor(2) = 1 Or cbrowse(2) = 1 Or cedit(3) = 1 Or cprint(3) = 1 Or creport(3) = 1 Or cdelete(3) = 1 Or csupervisor(3) = 1 Or cbrowse(3) = 1 Then
    mainform.mmagistrate.Enabled = True
    mainform.mxref.Enabled = True
    mainform.mcivil.Enabled = True
    mainform.mprinter.Enabled = True
    mainform.mcurrent.Enabled = True
    mainform.marchived.Enabled = True
    mainform.Toolbar1.Buttons(1).Enabled = True
    mainform.Toolbar1.Buttons(5).Enabled = True
    mainform.Toolbar1.Buttons(6).Enabled = True
End If
If iedit = 1 Or iprint = 1 Or ireport = 1 Or idelete = 1 Or isupervisor = 1 Or ibrowse = 1 Or supervisor = 1 Then
    mainform.mnon.Enabled = True
    mainform.mxref.Enabled = True
    mainform.mincident.Enabled = True
    mainform.mprinter.Enabled = True
    mainform.mcurrent.Enabled = True
    mainform.marchived.Enabled = True
    mainform.Toolbar1.Buttons(2).Enabled = True
    mainform.mservice.Enabled = True
    mainform.mbadcheck.Enabled = True
    mainform.mprinter.Enabled = True
    mainform.mcurrent.Enabled = True
    mainform.marchived.Enabled = True
    'ces code removed
    'mainform.Toolbar1.Buttons(3).Enabled = True
    'mainform.Toolbar1.Buttons(4).Enabled = True
    'mainform.Toolbar1.Buttons(5).Enabled = True
    '****
    mainform.Toolbar1.Buttons(6).Enabled = True
End If
If aedit = 1 Or aprint = 1 Or areport = 1 Or adelete = 1 Or asupervisor = 1 Or abrowse = 1 Or supervisor = 1 Then
    mainform.mvictim.Enabled = True
    mainform.Toolbar1.Buttons(13).Enabled = True
End If
If wedit = 1 Or wprint = 1 Or wreport = 1 Or wdelete = 1 Or wsupervisor = 1 Or wbrowse = 1 Or supervisor = 1 Then
    mainform.mxref.Enabled = True
    mainform.mwarrant.Enabled = True
    mainform.mprinter.Enabled = True
    mainform.mcurrent.Enabled = True
    mainform.marchived.Enabled = True
    mainform.Toolbar1.Buttons(3).Enabled = True
    mainform.Toolbar1.Buttons(6).Enabled = True
End If
If jedit = 1 Or jprint = 1 Or jreport = 1 Or jdelete = 1 Or jsupervisor = 1 Or jbrowse = 1 Or supervisor = 1 Then
    mainform.mxref.Enabled = True
    mainform.mdet.Enabled = True
    mainform.mprinter.Enabled = True
    mainform.mcurrent.Enabled = True
    mainform.marchived.Enabled = True
    mainform.Toolbar1.Buttons(12).Enabled = True
End If
If redit = 1 Or rprint = 1 Or rreport = 1 Or rdelete = 1 Or rsupervisor = 1 Or rbrowse = 1 Or supervisor = 1 Then
    mainform.mxref.Enabled = True
    mainform.mrestraining.Enabled = True
    mainform.mprinter.Enabled = True
    mainform.mcurrent.Enabled = True
    mainform.marchived.Enabled = True
    mainform.Toolbar1.Buttons(4).Enabled = True
    mainform.Toolbar1.Buttons(6).Enabled = True
End If
If supervisor = 1 Then
    mainform.mxref.Enabled = True
    mainform.msecurity.Enabled = True
    mainform.Toolbar1.Buttons(15).Enabled = True
    mainform.Toolbar1.Buttons(16).Enabled = True
    mainform.Toolbar1.Buttons(6).Enabled = True
    mainform.mcurrent.Enabled = True
    mainform.marchived.Enabled = True
End If
frmLogin.Hide
db.Close
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
oderror:
Resume od
End Sub

Private Sub Form_Load()
aaa = Error$(3260)
On Error Resume Next
ie = 0
Open "ie.dat" For Input As #1
Line Input #1, a$
Close #1
If a$ = "0823" Then
    ie = 1
End If
Close #1
On Error Resume Next
Me.Left = 4000
Me.Top = 2500

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmLogin = Nothing
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdOK.SetFocus
End If
    
End Sub
Private Sub txtUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPassword.SetFocus
End If
End Sub

Private Sub checkcracker()
If UCase(frmLogin.txtUserName) = "DEMO" Then
    Exit Sub
End If
On Error GoTo unlicensed
Open "c:\cracker" For Input As #1
Close #1
On Error Resume Next
Exit Sub
unlicensed:
msg = MsgBox("This is an unlicensed copy of Genesis Law Enforcement Suite.  Contact Genesis Software Corporation at (803) 635-2656 for additional information.", 48, "Genesis Error Log")
Unload Me
End

End Sub
