VERSION 5.00
Begin VB.Form frmLvOldWin2WinConvert 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert Leave Table Old Windows to New Windows"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11652
   Icon            =   "PRLvOldWin2Win.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      Left            =   5100
      TabIndex        =   2
      Top             =   7188
      Width           =   3756
   End
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
      Left            =   2796
      TabIndex        =   1
      Top             =   7188
      Width           =   1932
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      Height          =   1116
      Left            =   768
      Top             =   2304
      Width           =   10284
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PLEASE MAKE SURE THAT ""PRSYS.DAT"" FILE EXISTS IN THE PRDATA FOLDER."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   732
      Left            =   912
      TabIndex        =   8
      Top             =   5088
      Width           =   9996
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CUSTOMERS BEING CONVERTED FROM DOS TO WINDOWS SHOULD NOT USE THIS PROGRAM. "
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
      TabIndex        =   7
      Top             =   2496
      Width           =   9996
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"PRLvOldWin2Win.frx":08CA
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1452
      Left            =   912
      TabIndex        =   6
      Top             =   3504
      Width           =   9996
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "THIS SHOULD BE A DOS TO WINDOWS CONVERSION. CONVERSION ABORTED."
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
      TabIndex        =   5
      Top             =   5940
      Width           =   8652
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
      TabIndex        =   4
      Top             =   5952
      Width           =   8652
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
      TabIndex        =   3
      Top             =   5952
      Width           =   5244
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"PRLvOldWin2Win.frx":0974
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1500
      Left            =   912
      TabIndex        =   0
      Top             =   624
      Width           =   9996
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      Height          =   8508
      Left            =   192
      Top             =   192
      Width           =   11292
   End
End
Attribute VB_Name = "frmLvOldWin2WinConvert"
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
  Dim THandle As Integer
  Dim NumOfTRecs As Double
  Dim TRec As TransRecType
  
  OpenOldLeaveFileName OldHandle
  NumOfLvRecs = LOF(OldHandle) / Len(OldLeaveRec)
  
  ReDim TempVacMax(1 To NumOfLvRecs) As Double
  ReDim TempVEntry(1 To NumOfLvRecs, 1 To 20) As LeaveEntryType
  ReDim TempSICKMAX(1 To NumOfLvRecs) As Double
  ReDim TempSEntry(1 To NumOfLvRecs, 1 To 20) As LeaveEntryType
  
  If Exist("PRDATA\PRSYS.DAT") Then
    If FileLen("PRDATA\PRSYS.DAT") = 337 Then
      Label4.Visible = True
      Exit Sub
    End If
  Else
    Label3.Visible = True
    Exit Sub
  End If

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
  
  OpenTransHistFile THandle
  NumOfTRecs = LOF(THandle) / Len(TRec)
  For x = 1 To NumOfTRecs
    Get THandle, x, TRec
    TRec.Less401k = False
    Put THandle, x, TRec
  Next x
  Close
  Label2.Visible = True
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

