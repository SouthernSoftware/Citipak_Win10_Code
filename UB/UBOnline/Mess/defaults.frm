VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form defaults 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Defaults"
   ClientHeight    =   3000
   ClientLeft      =   2520
   ClientTop       =   2625
   ClientWidth     =   7035
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   7035
   Begin VB.TextBox txtFields 
      DataField       =   "type"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   9
      Top             =   240
      Width           =   1575
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   7035
      TabIndex        =   5
      Top             =   2370
      Width           =   7035
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4575
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3000
         TabIndex        =   7
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   1440
         TabIndex        =   6
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DEFAULT"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   4
      Top             =   1425
      Width           =   375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CODE"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Top             =   855
      Width           =   5535
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   2670
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=incident.mdb;"
      OLEDBString     =   "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=incident.mdb;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select TYPE,CODE,DEFAULT from codes Order by TYPE"
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "DEFAULT:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1425
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "CODE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "TYPE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "defaults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim networkpath As String
Private Sub Command1_Click()
qa$ = ""
If txtFields(0) > "" Then
    If txtFields(1) > "" Then
        If txtFields(2) > "" Then
            qa$ = qa$ + "type = '" + txtFields(0) + "' and code = '" + txtFields(1) + "' and default = '" + txtFields(2) + "'"
        Else
            qa$ = qa$ + "type = '" + txtFields(0) + "' and code = '" + txtFields(1) + "'"
        End If
    Else
        If txtFields(2) > "" Then
            qa$ = qa$ + "type = '" + txtFields(0) + "' and default = '" + txtFields(2) + "'"
        Else
            qa$ = qa$ + "type = '" + txtFields(0) + "'"
        End If
    End If
Else
    If txtFields(1) > "" Then
        If txtFields(2) > "" Then
            qa$ = qa$ + "code = '" + txtFields(1) + "' and default = '" + txtFields(2) + "'"
        Else
            qa$ = qa$ + "code = '" + txtFields(1) + "'"
        End If
    Else
        If txtFields(2) > "" Then
            qa$ = qa$ + "default = '" + txtFields(2) + "'"
        Else
            Exit Sub
        End If
    End If
End If
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(networkpath + "incident.mdb")
Set rs = db.OpenRecordset("select * from codes where " + qa$)
On Error Resume Next
If rs.EOF Then
    Exit Sub
Else
    rs.MoveFirst
End If
txtFields(0) = rs("type")
txtFields(1) = rs("code")
txtFields(2) = rs("default")
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
networkpath = ""
nw$ = ""
Open "nwi.ini" For Input As #1
Line Input #1, nw$
networkpath = nw$
Close #1

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & description
End Sub

Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  datPrimaryRS.Caption = "Record: " & CStr(datPrimaryRS.Recordset.AbsolutePosition)
End Sub

Private Sub datPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  datPrimaryRS.Refresh
  Exit Sub
RefreshErr:
  MsgBox Err.description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  datPrimaryRS.Recordset.UpdateBatch adAffectAll
  Exit Sub
UpdateErr:
  MsgBox Err.description
End Sub

Private Sub cmdClose_Click()
  Unload defaults
End Sub

