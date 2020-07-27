VERSION 5.00
Begin VB.Form frmLvOldWin2NewWin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conversion Windows to New Windows"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11652
   Icon            =   "frmLvWin2NewWin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Caption         =   "ESC E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   684
      Left            =   2784
      TabIndex        =   1
      Top             =   6792
      Width           =   1932
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "F10  &Proceed With Conversion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   684
      Left            =   5088
      TabIndex        =   0
      Top             =   6792
      Width           =   3756
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   444
      Left            =   1548
      TabIndex        =   9
      Top             =   7776
      Width           =   8556
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "AN ATTEMPT IS BEING MADE TO CONVERT WINDOWS DATA THAT HAS ALREADY BEEN CONVERTED. CONVERSION ABORTED."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1356
      Left            =   1500
      TabIndex        =   8
      Top             =   5280
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      Height          =   8508
      Left            =   192
      Top             =   192
      Width           =   11292
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "AN ATTEMPT IS BEING MADE TO CONVERT DOS DATA. THIS CONVERSION IS DESIGNED ONLY FOR WINDOWS TO WINDOWS. CONVERSION ABORTED."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1260
      Left            =   1500
      TabIndex        =   2
      Top             =   5232
      Width           =   8652
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      FillColor       =   &H000000FF&
      Height          =   1068
      Left            =   720
      Top             =   2736
      Width           =   10380
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      FillColor       =   &H000000FF&
      Height          =   1260
      Left            =   576
      Top             =   2640
      Width           =   10620
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "THE STANDARD DOS2 WIN CONVERSION HAS BEEN UPDATED TO INCLUDE THE NEW LEAVE TABLE ADDITIONS."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   828
      Left            =   912
      TabIndex        =   7
      Top             =   4176
      Width           =   9996
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CUSTOMERS BEING CONVERTED FROM DOS TO WINDOWS SHOULD NOT USE THIS PROGRAM."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   780
      Left            =   912
      TabIndex        =   6
      Top             =   2880
      Width           =   9996
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmLvWin2NewWin.frx":08CA
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1836
      Left            =   912
      TabIndex        =   5
      Top             =   480
      Width           =   9996
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "CONVERSION COMPLETED SUCCESSFULLY!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   828
      Left            =   3204
      TabIndex        =   4
      Top             =   5520
      Width           =   5244
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "THE PRSYS.DAT FILE IS CANNOT BE FOUND IN THE PRDATA FOLDER. CONVERSION IS ABORTED."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   828
      Left            =   1500
      TabIndex        =   3
      Top             =   5232
      Width           =   8652
   End
End
Attribute VB_Name = "frmLvOldWin2NewWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdConvert_Click()
  Dim OldLeaveRec As OldLeaveRecType
  Dim OldHandle As Integer
  Dim LeaveRec As LeaveRecType
  Dim NewHandle As Integer
  Dim NumOfLvRecs As Integer
  Dim x As Integer
  Dim y As Integer
  Dim UnitRec As UnitFileRecType
  Dim UHandle As Integer
  Dim THRec As TransRecType
  Dim THHandle As Integer
  Dim NumOfTHRecs As Double
  Dim TWRec As TransRecType
  Dim TWHandle As Integer
  Dim NumOfTWRecs As Double
  Dim ErnNoMatchRec As EarnNoMatchType
  Dim ErnNoHandle As Integer
  Dim OldErnHandle As Integer
  Dim OldErnRec As OldErnCodeRecType
  Dim NumOfErns As Integer
  Dim ErnHandle As Integer
  Dim ErnRec As ErnCodeRecType
  Dim ThisPct As Double
  Dim Emp2Rec As EmpData2Type
  Dim EmpHandle As Integer
  Dim NumOfEmpRecs As Integer
  
  OpenUnitFile UHandle
  Get UHandle, 1, UnitRec
  If QPTrim$(UnitRec.FileVer) = "Done" Then
    Label7.Visible = True
    Close
    Exit Sub
  End If
  
  If Exist("PRDATA\PRSYS.DAT") Then
    If FileLen("PRDATA\PRSYS.DAT") = 337 Then
      Label4.Visible = True
      Exit Sub
    End If
  Else
    Label3.Visible = True
    Exit Sub
  End If
  DoEvents
  Label8.Visible = True
  DoEvents
  UnitRec.LMT401YN = "N"
  UnitRec.FileVer = "Done"
  Put UHandle, 1, UnitRec
  Close UHandle
  
  OpenOldLeaveFileName OldHandle
  NumOfLvRecs = LOF(OldHandle) / Len(OldLeaveRec)
  
  ReDim TempVacMax(1 To NumOfLvRecs) As Double
  ReDim TempVEntry(1 To NumOfLvRecs, 1 To 20) As LeaveEntryType
  ReDim TempSICKMAX(1 To NumOfLvRecs) As Double
  ReDim TempSEntry(1 To NumOfLvRecs, 1 To 20) As LeaveEntryType
  For x = 1 To NumOfLvRecs
    Get OldHandle, x, OldLeaveRec
    TempVacMax(x) = OldLeaveRec.VacMax
    TempSICKMAX(x) = OldLeaveRec.SICKMAX
    For y = 1 To 20
      TempVEntry(x, y).EARN = OldLeaveRec.VEntry(y).EARN
      TempSEntry(x, y).EARN = OldLeaveRec.SEntry(y).EARN
      TempVEntry(x, y).YEARS = OldLeaveRec.VEntry(y).YEARS
      TempSEntry(x, y).YEARS = OldLeaveRec.SEntry(y).YEARS
    Next y
  Next x
  Close OldHandle
  DoEvents
  Label8.Caption = "Converting"
  DoEvents
  OpenLeaveFileName NewHandle
  For x = 1 To NumOfLvRecs
    LeaveRec.VacMax = TempVacMax(x)
    LeaveRec.SICKMAX = TempSICKMAX(x)
    LeaveRec.HolMax = 0
    LeaveRec.PerMax = 0
    For y = 1 To 20
      LeaveRec.VEntry(y).EARN = TempVEntry(x, y).EARN
      LeaveRec.SEntry(y).EARN = TempSEntry(x, y).EARN
      LeaveRec.HEntry(y).EARN = 0
      LeaveRec.PEntry(y).EARN = 0
      LeaveRec.VEntry(y).YEARS = TempVEntry(x, y).YEARS
      LeaveRec.SEntry(y).YEARS = TempSEntry(x, y).YEARS
      LeaveRec.HEntry(y).YEARS = 0
      LeaveRec.PEntry(y).YEARS = 0
    Next y
    Put NewHandle, x, LeaveRec
  Next x
  Close NewHandle
  
  OpenTransHistFile THHandle
  NumOfTHRecs = LOF(THHandle) / Len(THRec)
  If NumOfTHRecs = 0 Then
    Close
    GoTo NoTransHistRecs
  End If
  
  For x = 1 To NumOfTHRecs
    Get THHandle, x, THRec
      THRec.Pad1 = ""
      For y = 1 To 3
        THRec.Less401k(y) = False
      Next y
    Put THHandle, x, THRec
  Next x
  Close THHandle
  
NoTransHistRecs:

  OpenTransWorkFile TWHandle
  NumOfTWRecs = LOF(TWHandle) / Len(TWRec)
  If NumOfTWRecs = 0 Then
    Close
    GoTo NoTransWorkRecs
  End If
  
  For x = 1 To NumOfTWRecs
    Get TWHandle, x, TWRec
      TWRec.Pad1 = ""
      For y = 1 To 3
        TWRec.Less401k(y) = False
      Next y
    Put TWHandle, x, TWRec
  Next x
Close TWHandle
  
NoTransWorkRecs:

  OpenOldErnCodeFile OldErnHandle
  NumOfErns = LOF(OldErnHandle) / Len(OldErnRec)
  If NumOfErns = 0 Then
    GoTo NoErnMatchNeeded
  End If
  
  ReDim TempERNCODE1(1 To NumOfErns) As String * 10
  ReDim TempERNFWT1(1 To NumOfErns) As String * 1
  ReDim TempERNSWT1(1 To NumOfErns) As String * 1
  ReDim TempERNSOC1(1 To NumOfErns) As String * 1
  ReDim TempERNMED1(1 To NumOfErns) As String * 1
  ReDim TempERNRET1(1 To NumOfErns) As String * 1
  For x = 1 To NumOfErns
    Get OldErnHandle, x, OldErnRec
    TempERNCODE1(x) = QPTrim$(OldErnRec.ERNCODE1)
    TempERNFWT1(x) = OldErnRec.ERNFWT1
    TempERNSWT1(x) = OldErnRec.ERNSWT1
    TempERNSOC1(x) = OldErnRec.ERNSOC1
    TempERNMED1(x) = OldErnRec.ERNMED1
    TempERNRET1(x) = OldErnRec.ERNRET1
  Next x
  Close OldErnHandle
  
  
  OpenErnCodeFile ErnHandle
  For x = 1 To NumOfErns
    ErnRec.ERNCODE1 = QPTrim$(TempERNCODE1(x))
    ErnRec.ERNFWT1 = TempERNFWT1(x)
    ErnRec.ERNSWT1 = TempERNSWT1(x)
    ErnRec.ERNSOC1 = TempERNSOC1(x)
    ErnRec.ERNMED1 = TempERNMED1(x)
    ErnRec.ERNRET1 = TempERNRET1(x)
    ErnRec.EarnYN = "Y" '"Y"es include in match
    ErnRec.Pad = ""
    Put ErnHandle, x, ErnRec
  Next x
  Close ErnHandle
    
NoErnMatchNeeded:
  
  OpenEmpData2File EmpHandle
  NumOfEmpRecs = LOF(EmpHandle) / Len(Emp2Rec)
  For x = 1 To NumOfEmpRecs
    Get EmpHandle, x, Emp2Rec
    If QPTrim$(Emp2Rec.EMPRETTP) = "" Then
      Emp2Rec.YN401K = "N"
      Put EmpHandle, x, Emp2Rec
    End If
  Next x
  Close EmpHandle
  
  Label2.Visible = True
  Label8.Visible = False
End Sub

Private Sub cmdExit_Click()
  End
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Label2.Visible = False
  Label3.Visible = False
  Label4.Visible = False
  Label7.Visible = False
  Label8.Visible = False
End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%x"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      KeyCode = 0
    Case Else:
  End Select
End Sub


