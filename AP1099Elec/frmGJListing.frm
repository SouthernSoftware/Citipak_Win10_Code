VERSION 5.00
Begin VB.Form frmGJListing 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Journal Entries"
   ClientHeight    =   2820
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   9264
   Icon            =   "frmGJListing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   9264
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D0D0D0&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7680
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   1092
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6360
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   1092
   End
   Begin VB.ListBox lstGJEntries 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1728
      ItemData        =   "frmGJListing.frx":08CA
      Left            =   240
      List            =   "frmGJListing.frx":08CC
      TabIndex        =   0
      Top             =   480
      Width           =   8772
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To Select Item to Edit, Double-Click Item or Highlight and Click Ok."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   6132
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   9000
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Debit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   6720
      TabIndex        =   7
      Top             =   120
      Width           =   612
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Credit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   8040
      TabIndex        =   6
      Top             =   120
      Width           =   612
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   4080
      TabIndex        =   5
      Top             =   120
      Width           =   1092
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   852
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   492
   End
End
Attribute VB_Name = "frmGJListing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Dim GJEdit As TrEditRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Private Temp_Class As Resize_Class
Private Sub Form_Activate()
  lstGJEntries.ListIndex = 0
End Sub

Private Sub cmdExit_Click()
  Unload frmGJListing
End Sub

Private Sub cmdOk_Click()
  Dim TempRec As Integer
  If Not lstGJEntries = "" Then
    TempRec = Mid$(lstGJEntries, 78)
    frmGenJournalEntry.Rec2Form (TempRec)
  End If
  Unload frmGJListing
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  ListEntries
End Sub

Private Sub ListEntries()
'this fills the listbox of General Journal Entries for User to Select from or view
  Dim GJEditFile As Integer, GJEditFileNum As Integer, NumEdTrans As Integer
  Dim cnt As Integer
  Dim fmt As String
  Dim tempstr As String
  Dim strInfo As String
  Dim disdate As String
  fmt = "$########.##"
  OpenGJEditFile GJEditFileNum, NumEdTrans
  If NumEdTrans > 0 Then
    For cnt = 1 To NumEdTrans
      Get GJEditFileNum, cnt, GJEdit
      If Not GJEdit.Deleted Then
        tempstr = Space$(85)
        Mid$(tempstr, 2) = Format(DateAdd("d", (GJEdit.TRDATE), "12-31-1979"), "mm/dd/yyyy")
        Mid$(tempstr, 15) = GJEdit.AcctNum
        Mid$(tempstr, 32) = Left$(GJEdit.Desc, 15)
        Mid$(tempstr, 49) = Using(fmt, GJEdit.DrAmt)
        Mid$(tempstr, 60) = Using(fmt, GJEdit.CrAmt)
        Mid$(tempstr, 78) = Str$(cnt)
        lstGJEntries.AddItem tempstr
      End If
    Next
  Else
    MsgBox "No Entries To Display.", vbOKOnly, "No Entries"
  End If
  Close GJEditFile
End Sub

Private Sub lstGJEntries_DblClick()
  Dim TempRec As Integer
  TempRec = Mid$(lstGJEntries, 78)
  frmGenJournalEntry.Rec2Form (TempRec)
  Unload frmGJListing
End Sub
