VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.MDIForm mainform 
   BackColor       =   &H00808000&
   Caption         =   "Genesis Law Enforcement Suite version 2.0"
   ClientHeight    =   5880
   ClientLeft      =   690
   ClientTop       =   1890
   ClientWidth     =   10395
   Icon            =   "MAINFORM.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   28
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Civil Service"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Incident Report"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Warrant Manager"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Restraining Order"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Magistrate"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cross Reference"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Employee"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Registration"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Tickets"
            ImageIndex      =   24
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Investigations"
            ImageIndex      =   25
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "911-Dispatch"
            ImageIndex      =   29
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Detention Center"
            ImageIndex      =   28
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Victim Advocate"
            ImageIndex      =   35
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Printer Setup"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Security Setup"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "People Maintenance"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Professional Maintenance"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Help"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Booking"
            ImageIndex      =   33
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Property Profile"
            ImageIndex      =   30
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Classification"
            ImageIndex      =   34
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Inmate Affairs"
            ImageIndex      =   31
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Inmate Incident"
            ImageIndex      =   35
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "State Transfer"
            ImageIndex      =   32
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Detention Reports"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Detention Setup"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Classification Setup"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1095
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   35
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":0442
            Key             =   ""
            Object.Tag             =   "Magistrate Papers"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":06DA
            Key             =   ""
            Object.Tag             =   "Family Court Papers"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":096E
            Key             =   ""
            Object.Tag             =   "Writ/Other Papers"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":0C02
            Key             =   ""
            Object.Tag             =   "Executions"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":0E96
            Key             =   ""
            Object.Tag             =   "Incident Report"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":12EA
            Key             =   ""
            Object.Tag             =   "Warrant Manager"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":173E
            Key             =   ""
            Object.Tag             =   "Booking Report"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":1B92
            Key             =   ""
            Object.Tag             =   "Restraining Order"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":1EAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":2302
            Key             =   ""
            Object.Tag             =   "Exit System"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":2756
            Key             =   ""
            Object.Tag             =   "Printer Setup"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":2BAA
            Key             =   ""
            Object.Tag             =   "Security Setup"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":2FFE
            Key             =   ""
            Object.Tag             =   "System Defaults"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":3452
            Key             =   ""
            Object.Tag             =   "Help"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":38A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":3CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":3D58
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":41AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":45FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":4A52
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":4EA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":52FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":574E
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":5BA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":5FF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":6312
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":6766
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":6BBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":700E
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":7462
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":78B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":7D0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":815E
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":88BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAINFORM.frx":8D12
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5000
      Top             =   10000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.txt"
      DialogTitle     =   "Split File"
      FileName        =   "*.txt"
      Filter          =   "*.txt"
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   2640
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mpackage 
      Caption         =   "&Package"
      Begin VB.Menu mcivil 
         Caption         =   "&1)    Genesis Civil Service"
         Enabled         =   0   'False
      End
      Begin VB.Menu mincident 
         Caption         =   "&2)    Genesis Incident Report"
         Enabled         =   0   'False
      End
      Begin VB.Menu mvictim 
         Caption         =   "&3)    Genesis Victim Advocate"
         Enabled         =   0   'False
      End
      Begin VB.Menu mbadcheck 
         Caption         =   "&4)    Genesis Fraudulent Check"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnon 
         Caption         =   "&5)    Genesis Non-Criminal Police Response"
         Enabled         =   0   'False
      End
      Begin VB.Menu mservice 
         Caption         =   "&6)    Genesis Service Call Report"
         Enabled         =   0   'False
      End
      Begin VB.Menu mwarrant 
         Caption         =   "&7)    Genesis Warrant Manager"
         Enabled         =   0   'False
      End
      Begin VB.Menu mrestraining 
         Caption         =   "&8)    Genesis Restraining Order"
         Enabled         =   0   'False
      End
      Begin VB.Menu mmagistrate 
         Caption         =   "&9)    Genesis Magistrate System"
         Enabled         =   0   'False
      End
      Begin VB.Menu mxref 
         Caption         =   "&10)  Genesis Cross Reference"
      End
      Begin VB.Menu memployee 
         Caption         =   "&11)  Genesis RMS - Employee"
         Enabled         =   0   'False
      End
      Begin VB.Menu mregistation 
         Caption         =   "&12)  Genesis RMS - Registration"
         Enabled         =   0   'False
      End
      Begin VB.Menu mtickets 
         Caption         =   "&13)  Genesis RMS - Tickets"
         Enabled         =   0   'False
      End
      Begin VB.Menu minvest 
         Caption         =   "&14)  Genesis RMS - Investigations"
         Enabled         =   0   'False
      End
      Begin VB.Menu m911 
         Caption         =   "&15)  Genesis Dispatch-911"
         Enabled         =   0   'False
      End
      Begin VB.Menu mdet 
         Caption         =   "&16)  GenesisDetention Center"
         Enabled         =   0   'False
      End
      Begin VB.Menu mexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mview 
      Caption         =   "&View"
      WindowList      =   -1  'True
      Begin VB.Menu mcurrent 
         Caption         =   "&Current Data"
         Checked         =   -1  'True
         Enabled         =   0   'False
      End
      Begin VB.Menu marchived 
         Caption         =   "&Archived Data"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu msystem 
      Caption         =   "&System "
      Begin VB.Menu mprinter 
         Caption         =   "&Printer Setup"
         Enabled         =   0   'False
      End
      Begin VB.Menu msecurity 
         Caption         =   "Se&curity Setup"
         Enabled         =   0   'False
      End
      Begin VB.Menu pcleanup 
         Caption         =   "People &Cleanup"
         Enabled         =   0   'False
      End
      Begin VB.Menu omaint 
         Caption         =   "&Professional Maintenance"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mhelp 
      Caption         =   "&Help"
      Begin VB.Menu mcontents 
         Caption         =   "Help &Contents"
      End
      Begin VB.Menu mabout 
         Caption         =   "&About Genesis Law Enforcement Suite"
      End
   End
End
Attribute VB_Name = "mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub export_Click()
frmexport.Show
End Sub

Private Sub import_Click()
frmimport.Show
End Sub


Private Sub m911_Click()
If UCase(frmLogin.txtPassword) = "DEMO" And UCase(frmLogin.txtUserName) = "DEMO" Then
    MsgBox "Not available in DEMO version.", 48, "Genesis Error Log"
    Exit Sub
End If
Screen.MousePointer = 11
dispatch.Show
Screen.MousePointer = 0
End Sub

Private Sub mabout_Click()
Screen.MousePointer = 11
frmAbout.Show
Screen.MousePointer = 0
End Sub

Private Sub marchived_Click()
If marchived.checked = False Then
    mcurrent.checked = False
    marchived.checked = True
    For t% = 0 To Forms.Count - 1
        If Left$(LCase(Forms(t%).Name), 3) = "frm" Then
            If LCase(Forms(t%).Name) <> "frmlogin" Then
                Unload Forms(t%)
            End If
        End If
    Next t%
Else
    mcurrent.checked = True
    marchived.checked = False
    For t% = 0 To Forms.Count - 1
        If Left$(LCase(Forms(t%).Name), 3) = "frm" Then
            If LCase(Forms(t%).Name) <> "frmlogin" Then
                Unload Forms(t%)
            End If
        End If
    Next t%
End If

End Sub

Private Sub mbadcheck_Click()
If frmLogin.IBROWSE = 0 Then
    msg = MsgBox("Insufficient access authority.", 48, "Genesis Error Log")
    Screen.MousePointer = 0
    Exit Sub
End If
Screen.MousePointer = 11
On Error GoTo checkload
badcheck.Show
On Error GoTo 0
Screen.MousePointer = 0
Exit Sub
checkload:
If Err = 424 Then
    msg = MsgBox("Selected module is not a part of this installation.", 48, "Genesis Error Log")
Else
    msg = MsgBox(Error$(Err), 48, "Genesis Error Log")
End If
Resume Next

End Sub

Private Sub mcontents_Click()
CommonDialog2.HelpFile = "suite.HLP"
CommonDialog2.HelpCommand = cdlHelpContents
CommonDialog2.ShowHelp

End Sub
Private Sub mcurrent_Click()
If mcurrent.checked = False Then
    mcurrent.checked = True
    marchived.checked = False
    For t% = 0 To Forms.Count - 1
        If Left$(LCase(Forms(t%).Name), 3) = "frm" Then
            If LCase(Forms(t%).Name) <> "frmlogin" Then
                Unload Forms(t%)
            End If
        End If
    Next t%
Else
    mcurrent.checked = False
    marchived.checked = True
    For t% = 0 To Forms.Count - 1
        If Left$(LCase(Forms(t%).Name), 3) = "frm" Then
            If LCase(Forms(t%).Name) <> "frmlogin" Then
                Unload Forms(t%)
            End If
        End If
    Next t%
End If
    
End Sub

Private Sub mdet_Click()
Screen.MousePointer = 11
For tb = 18 To 27
    If Toolbar1.Buttons(tb).Visible = True Then
        mdet.Caption = "Detention Center"
        Toolbar1.Buttons(tb).Visible = False
    Else
        mdet.Caption = "Close Detention Center"
        Toolbar1.Buttons(tb).Visible = True
    End If
Next tb
Screen.MousePointer = 0
End Sub

Private Sub MDIForm_Load()
On Error Resume Next
NW$ = ""
Open "NW.INI" For Input As #1
Line Input #1, JK$
Line Input #1, NW$
Close #1
DOPOPUP = NW$
NW$ = ""
Open "nwc.ini" For Input As #1
Line Input #1, NW$
nwc = NW$
Line Input #1, NW$
usesf = NW$
Close #1
NW$ = ""
Open "nwl.ini" For Input As #1
Line Input #1, NW$
nwl = NW$
Close #1
NW$ = ""
Open "nww.ini" For Input As #1
Line Input #1, NW$
nww = NW$
Close #1
NW$ = ""
Open "nwrep.ini" For Input As #1
Line Input #1, NW$
If NW$ > "" Then
    nwrep = NW$
Else
    nwrep = nwl
End If
Close #1
NW$ = ""
Open "nwr.ini" For Input As #1
Line Input #1, NW$
nwr = NW$
Close #1
NW$ = ""
Open "nwj.ini" For Input As #1
Line Input #1, NW$
nwj = NW$
Close #1
NW$ = ""
Open "nwm.ini" For Input As #1
Line Input #1, NW$
nwm = NW$
Close #1
NW$ = ""
Open "nwi.ini" For Input As #1
Line Input #1, NW$
nwi = NW$
Close #1
NW$ = ""
Open "nws.ini" For Input As #1
Line Input #1, NW$
nws = NW$
Close #1
NW$ = ""
Open "nwb.ini" For Input As #1
Line Input #1, NW$
nwb = NW$
Close #1

Kill "*.dsk"
Screen.MousePointer = 11
frmLogin.Show
Screen.MousePointer = 0
Exit Sub
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next
Set mainform = Nothing
End
End Sub

Private Sub memployee_Click()
If UCase(frmLogin.txtPassword) = "DEMO" And UCase(frmLogin.txtUserName) = "DEMO" Then
    MsgBox "Not available in DEMO version.", 48, "Genesis Error Log"
    Exit Sub
End If
Screen.MousePointer = 11
Employee.Show
Screen.MousePointer = 0

End Sub

Private Sub mexit_Click()
Unload mainform
End
End Sub
Private Sub mcivil_Click()
If frmLogin.CBROWSE(0) = 0 And frmLogin.CBROWSE(1) = 0 And frmLogin.CBROWSE(2) = 0 And frmLogin.CBROWSE(3) = 0 Then
    msg = MsgBox("Insufficient access authority.", 48, "Genesis Error Log")
    Screen.MousePointer = 0
    Exit Sub
End If
Screen.MousePointer = 11
On Error GoTo checkload
CIVIL.Show
On Error GoTo 0
Screen.MousePointer = 0
Exit Sub
Exit Sub
checkload:
If Err = 424 Then
    msg = MsgBox("Selected module is not a part of this installation.", 48, "Genesis Error Log")
Else
    msg = MsgBox(Error$(Err), 48, "Genesis Error Log")
End If
Resume Next
End Sub

Private Sub mincident_Click()
If frmLogin.IBROWSE = 0 Then
    msg = MsgBox("Insufficient access authority.", 48, "Genesis Error Log")
    Screen.MousePointer = 0
    Exit Sub
End If
Screen.MousePointer = 11
On Error GoTo checkload
incident.WindowState = vbMaximized
incident.Show
On Error GoTo 0
Screen.MousePointer = 0
Exit Sub
checkload:
If Err = 424 Then
    msg = MsgBox("Selected module is not a part of this installation.", 48, "Genesis Error Log")
Else
    msg = MsgBox(Error$(Err), 48, "Genesis Error Log")
End If
Resume Next
End Sub

Private Sub minvest_Click()
If UCase(frmLogin.txtPassword) = "DEMO" And UCase(frmLogin.txtUserName) = "DEMO" Then
    MsgBox "Not available in DEMO version.", 48, "Genesis Error Log"
    Exit Sub
End If
Screen.MousePointer = 11
INVESTIGATE.Show
Screen.MousePointer = 0

End Sub

Private Sub mmagistrate_Click()
If frmLogin.CBROWSE(0) = 0 Then
    msg = MsgBox("Insufficient access authority.", 48, "Genesis Error Log")
    Screen.MousePointer = 0
    Exit Sub
End If
Screen.MousePointer = 11
On Error GoTo checkload
magistrate.Show
On Error GoTo 0
Screen.MousePointer = 0
magistrate.mpapertype.SetFocus
Exit Sub
checkload:
If Err = 424 Then
    msg = MsgBox("Selected module is not a part of this installation.", 48, "Genesis Error Log")
Else
    msg = MsgBox(Error$(Err), 48, "Genesis Error Log")
End If
Resume Next
End Sub

Private Sub mnon_Click()
Screen.MousePointer = 11
On Error GoTo checkload
noncrim.Show
On Error GoTo 0
Screen.MousePointer = 0
Exit Sub
checkload:
If Err = 424 Then
    msg = MsgBox("Selected module is not a part of this installation.", 48, "Genesis Error Log")
Else
    msg = MsgBox(Error$(Err), 48, "Genesis Error Log")
End If
Resume Next

End Sub

Private Sub mprinter_Click()
CommonDialog2.ShowPrinter
Exit Sub

End Sub

Private Sub mregistation_Click()
If UCase(frmLogin.txtPassword) = "DEMO" And UCase(frmLogin.txtUserName) = "DEMO" Then
    MsgBox "Not available in DEMO version.", 48, "Genesis Error Log"
    Exit Sub
End If
Screen.MousePointer = 11
REGISTRATION.Show
Screen.MousePointer = 0

End Sub

Private Sub mrestraining_Click()
If frmLogin.WBROWSE = 0 Then
    msg = MsgBox("Insufficient access authority.", 48, "Genesis Error Log")
    Screen.MousePointer = 0
    Exit Sub
End If
Screen.MousePointer = 11
On Error GoTo checkload
ro.Show
On Error GoTo 0
Screen.MousePointer = 0
Exit Sub
checkload:
If Err = 424 Then
    msg = MsgBox("Selected module is not a part of this installation.", 48, "Genesis Error Log")
Else
    msg = MsgBox(Error$(Err), 48, "Genesis Error Log")
End If
Resume Next
End Sub

Private Sub msecurity_Click()
If frmLogin.SUPERVISOR = 0 Then
    msg = MsgBox("Insufficient access authority.", 48, "Genesis Error Log")
    Screen.MousePointer = 0
    Exit Sub
End If
Screen.MousePointer = 11
SECURITY.Show
Screen.MousePointer = 0
End Sub
Private Sub mservice_Click()
If frmLogin.IBROWSE = 0 Then
    msg = MsgBox("Insufficient access authority.", 48, "Genesis Error Log")
    Screen.MousePointer = 0
    Exit Sub
End If
Screen.MousePointer = 11
On Error GoTo checkload
service.Show
On Error GoTo 0
Screen.MousePointer = 0
Exit Sub
checkload:
If Err = 424 Then
    msg = MsgBox("Selected module is not a part of this installation.", 48, "Genesis Error Log")
Else
    msg = MsgBox(Error$(Err), 48, "Genesis Error Log")
End If
Resume Next

End Sub

Private Sub mtickets_Click()
If UCase(frmLogin.txtPassword) = "DEMO" And UCase(frmLogin.txtUserName) = "DEMO" Then
    MsgBox "Not available in DEMO version.", 48, "Genesis Error Log"
    Exit Sub
End If
Screen.MousePointer = 11
TICKETS.Show
Screen.MousePointer = 0
End Sub

Private Sub mvictim_Click()
If frmLogin.abrowse = 0 Then
    msg = MsgBox("Insufficient access authority.", 48, "Genesis Error Log")
    Screen.MousePointer = 0
    Exit Sub
End If
Screen.MousePointer = 11
On Error GoTo checkload
advocate.WindowState = vbMaximized
advocate.Show
On Error GoTo 0
Screen.MousePointer = 0
Exit Sub
checkload:
If Err = 424 Then
    msg = MsgBox("Selected module is not a part of this installation.", 48, "Genesis Error Log")
Else
    msg = MsgBox(Error$(Err), 48, "Genesis Error Log")
End If
Resume Next

End Sub

Private Sub mwarrant_Click()
If frmLogin.WBROWSE = 0 Then
    msg = MsgBox("Insufficient access authority.", 48, "Genesis Error Log")
    Exit Sub
End If
Screen.MousePointer = 11
On Error GoTo checkload
warrant.Show
On Error GoTo 0
Screen.MousePointer = 0
Exit Sub
checkload:
If Err = 424 Then
    msg = MsgBox("Selected module is not a part of this installation.", 48, "Genesis Error Log")
Else
    msg = MsgBox(Error$(Err), 48, "Genesis Error Log")
End If
Resume Next
End Sub

Private Sub mxref_Click()
Screen.MousePointer = 11
xref.Show
Screen.MousePointer = 0

End Sub

Private Sub pcleanup_Click()
cleanup.Show
End Sub

Private Sub ptm_Click()
If UCase(frmLogin.txtPassword) = "DEMO" And UCase(frmLogin.txtUserName) = "DEMO" Then
    MsgBox "Not available in DEMO version.", 48, "Genesis Error Log"
    Exit Sub
End If
Screen.MousePointer = 11
accessories.Show
Screen.MousePointer = 0

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim foundit As Boolean
Select Case Button.index

Case 1
If frmLogin.CBROWSE(0) = 0 And frmLogin.CBROWSE(1) = 0 And frmLogin.CBROWSE(2) = 0 And frmLogin.CBROWSE(3) = 0 Then
    msg = MsgBox("Insufficient access authority.", 48, "Genesis Error Log")
    Screen.MousePointer = 0
    Exit Sub
End If
Screen.MousePointer = 11
foundit = True
On Error GoTo checkload
CIVIL.Show
On Error GoTo 0
Screen.MousePointer = 0

    
Case 2
If frmLogin.IBROWSE = 0 Then
    msg = MsgBox("Insufficient access authority.", 48, "Genesis Error Log")
    Screen.MousePointer = 0
    Exit Sub
End If
Screen.MousePointer = 11
On Error GoTo checkload
incident.WindowState = vbMaximized
incident.Show
On Error GoTo 0
Screen.MousePointer = 0

'Ces Code removed
'Case 3
'If frmLogin.IBROWSE = 0 Then
'    msg = MsgBox("Insufficient access authority.", 48, "Genesis Error Log")
'    Screen.MousePointer = 0
'    Exit Sub
'End If
'Screen.MousePointer = 11
'On Error GoTo checkload
'service.Show
'On Error GoTo 0
'Screen.MousePointer = 0
'
'Case 4
'If frmLogin.IBROWSE = 0 Then
'    msg = MsgBox("Insufficient access authority.", 48, "Genesis Error Log")
'    Screen.MousePointer = 0
'    Exit Sub
'End If
'Screen.MousePointer = 11
'On Error GoTo checkload
'badcheck.Show
'On Error GoTo 0
'Screen.MousePointer = 0
'
'Case 5
'If frmLogin.IBROWSE = 0 Then
'    msg = MsgBox("Insufficient access authority.", 48, "Genesis Error Log")
'    Exit Sub
'End If
'Screen.MousePointer = 11
'On Error GoTo checkload
'noncrim.Show
'On Error GoTo 0
'Screen.MousePointer = 0
    
Case 3
If frmLogin.WBROWSE = 0 Then
    msg = MsgBox("Insufficient access authority.", 48, "Genesis Error Log")
    Exit Sub
End If
Screen.MousePointer = 11
On Error GoTo checkload
warrant.Show
On Error GoTo 0
Screen.MousePointer = 0
    
Case 4
If frmLogin.RBROWSE = 0 Then
    msg = MsgBox("Insufficient access authority.", 48, "Genesis Error Log")
    Exit Sub
End If
Screen.MousePointer = 11
On Error GoTo checkload
ro.Show
On Error GoTo 0
Screen.MousePointer = 0
    
Case 5
If frmLogin.CBROWSE(0) = 0 Then
    msg = MsgBox("insufficient access authority.", 48, "Genesis Error Log")
    Exit Sub
End If
Screen.MousePointer = 11
On Error GoTo checkload
magistrate.Show
Screen.MousePointer = 0
On Error GoTo 0

Case 6

Screen.MousePointer = 11
xref.Show
Screen.MousePointer = 0

Case 7
If UCase(frmLogin.txtPassword) = "DEMO" And UCase(frmLogin.txtUserName) = "DEMO" Then
    MsgBox "Not available in DEMO version.", 48, "Genesis Error Log"
    Exit Sub
End If
Screen.MousePointer = 11
Employee.Show
Screen.MousePointer = 0

Case 8
If UCase(frmLogin.txtPassword) = "DEMO" And UCase(frmLogin.txtUserName) = "DEMO" Then
    MsgBox "Not available in DEMO version.", 48, "Genesis Error Log"
    Exit Sub
End If
Screen.MousePointer = 11
REGISTRATION.Show
Screen.MousePointer = 0

Case 9
If UCase(frmLogin.txtPassword) = "DEMO" And UCase(frmLogin.txtUserName) = "DEMO" Then
    MsgBox "Not available in DEMO version.", 48, "Genesis Error Log"
    Exit Sub
End If
Screen.MousePointer = 11
TICKETS.Show
Screen.MousePointer = 0

Case 10
If UCase(frmLogin.txtPassword) = "DEMO" And UCase(frmLogin.txtUserName) = "DEMO" Then
    MsgBox "Not available in DEMO version.", 48, "Genesis Error Log"
    Exit Sub
End If
Screen.MousePointer = 11
INVESTIGATE.Show
Screen.MousePointer = 0

Case 11
If UCase(frmLogin.txtPassword) = "DEMO" And UCase(frmLogin.txtUserName) = "DEMO" Then
    MsgBox "Not available in DEMO version.", 48, "Genesis Error Log"
    Exit Sub
End If
Screen.MousePointer = 11
dispatch.Show
Screen.MousePointer = 0

Case 12
Screen.MousePointer = 11
For tb = 18 To 27
    If Toolbar1.Buttons(tb).Visible = True Then
        mdet.Caption = "Detention Center"
        Toolbar1.Buttons(tb).Visible = False
    Else
        mdet.Caption = "Close Detention Center"
        Toolbar1.Buttons(tb).Visible = True
    End If
Next tb
Screen.MousePointer = 0

Case 13
Screen.MousePointer = 11
advocate.Show
Screen.MousePointer = 0

Case 14
CommonDialog2.ShowPrinter

Case 15
If frmLogin.SUPERVISOR = 0 Then
    msg = MsgBox("Insufficient access authority.", 48, "Genesis Error Log")
    Screen.MousePointer = 0
    Exit Sub
End If
Screen.MousePointer = 11
SECURITY.Show
Screen.MousePointer = 0

Case 16
If frmLogin.SUPERVISOR = 0 And frmLogin.rmsbrowse = 0 And frmLogin.rmssupervisor = 0 Then
    msg = MsgBox("Insufficient access authority.", 48, "Genesis Error Log")
    Screen.MousePointer = 0
    Exit Sub
End If
Screen.MousePointer = 11
cleanup.Show
Screen.MousePointer = 0

Case 17
If frmLogin.SUPERVISOR = 0 Then
    msg = MsgBox("Insufficient access authority.", 48, "Genesis Error Log")
    Screen.MousePointer = 0
    Exit Sub
End If
Screen.MousePointer = 11
officer.Show
Screen.MousePointer = 0
    
Case 17
CommonDialog2.HelpFile = "suite.HLP"
CommonDialog2.HelpCommand = cdlHelpContents
CommonDialog2.ShowHelp

Case 18
    Screen.MousePointer = 11
    frmBookingReport.Show
    Screen.MousePointer = 0
    
Case 23
    Screen.MousePointer = 11
    frmInmateAffairs.Show
    Screen.MousePointer = 0
    
Case 24
    Screen.MousePointer = 11
    frmInmateIncident.Show
    Screen.MousePointer = 0
    
Case 25
    Screen.MousePointer = 11
    frmStateXfer.Show
    Screen.MousePointer = 0
    
Case 26
    Screen.MousePointer = 11
    frmReports.Show
    Screen.MousePointer = 0
    
Case 21
    Screen.MousePointer = 11
    frmPropProfile.Show
    Screen.MousePointer = 0
    
Case 27
    If frmLogin.SUPERVISOR = 0 Then
        msg = MsgBox("Insufficient access authority.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        Exit Sub
    End If
    Screen.MousePointer = 11
    jsetup.Show
    Screen.MousePointer = 0

Case 22
    Screen.MousePointer = 11
    frmClassify.Show
    Screen.MousePointer = 0
    
Case 28
    If frmLogin.SUPERVISOR = 0 Then
        msg = MsgBox("Insufficient access authority.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        Exit Sub
    End If
    Screen.MousePointer = 11
    frmClassificationSetup.Show
    Screen.MousePointer = 0


End Select
Exit Sub
checkload:
If Err = 424 Then
    msg = MsgBox("Selected module is not a part of this installation.", 48, "Genesis Error Log")
Else
    msg = MsgBox(Error$(Err), 48, "Genesis Error Log")
End If
Resume Next
End Sub


