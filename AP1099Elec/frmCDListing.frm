VERSION 5.00
Begin VB.Form frmCDListing 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Disbursement Listing"
   ClientHeight    =   3204
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   9228
   ForeColor       =   &H00000000&
   Icon            =   "frmCDListing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3204
   ScaleWidth      =   9228
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstCDEntries 
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
      Height          =   1968
      Left            =   72
      TabIndex        =   0
      Top             =   336
      Width           =   9084
   End
   Begin VB.CommandButton cmdOk 
      Appearance      =   0  'Flat
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
      Left            =   6264
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2808
      Width           =   1092
   End
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
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
      Left            =   7584
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2808
      Width           =   1092
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank"
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
      Index           =   8
      Left            =   5784
      TabIndex        =   8
      Top             =   0
      Width           =   612
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reference"
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
      Index           =   7
      Left            =   4248
      TabIndex        =   7
      Top             =   0
      Width           =   1092
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
      Left            =   504
      TabIndex        =   6
      Top             =   0
      Width           =   492
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Desc."
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
      Left            =   1896
      TabIndex        =   5
      Top             =   0
      Width           =   660
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Left            =   7896
      TabIndex        =   4
      Top             =   0
      Width           =   852
   End
   Begin VB.Line Line1 
      X1              =   24
      X2              =   8904
      Y1              =   240
      Y2              =   240
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
      Left            =   24
      TabIndex        =   3
      Top             =   2856
      Width           =   6132
   End
End
Attribute VB_Name = "frmCDListing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Dim GLCDEd(1) As CJEditRecType
Dim CJType As Integer
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Private Temp_Class As Resize_Class

Private Sub Form_Activate()
  lstCDEntries.ListIndex = 0
End Sub

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'  If ((UnloadMode = vbFormControlMenu)) Then
'    Cancel = True
'  End If
'End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyReturn:
      SendKeys "%O"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub cmdExit_Click()
  Unload frmCDListing
End Sub

Private Sub cmdOk_Click()
  Dim TempRec As Integer
  If Not lstCDEntries = "" Then
    TempRec = Mid$(lstCDEntries, 82)
    frmCashDisbEntry.Rec2Form (TempRec)
  End If
  Unload frmCDListing
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  ListEntries
End Sub

Private Sub ListEntries()
'this fills the listbox with Cash Disbursement Entries for User to Select from or view
  Dim CJEditFile As Integer, CJEditFileNum As Integer, NumEdTrans As Integer
  Dim cnt As Integer
  Dim fmt As String
  Dim tempstr As String
  Dim strInfo As String
  Dim disdate As String
  fmt = "$########.##"
  CJType = 2
  OpenCJEditFile CJEditFileNum, NumEdTrans, CJType
  If NumEdTrans > 0 Then
    For cnt = 1 To NumEdTrans
      Get CJEditFileNum, cnt, GLCDEd(1)
      If GLCDEd(1).DelFlag = 0 Then
        tempstr = Space$(90)
        Mid$(tempstr, 2) = Format(DateAdd("d", (GLCDEd(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
        Mid$(tempstr, 15) = Left$(GLCDEd(1).Desc, 20)
        Mid$(tempstr, 38) = Left$(GLCDEd(1).DOCREF, 8)
        Mid$(tempstr, 50) = Left$(GLCDEd(1).RECCODE, 2)
        Mid$(tempstr, 61) = Using(fmt, GLCDEd(1).Amt)
        Mid$(tempstr, 82) = Str$(cnt)
        lstCDEntries.AddItem tempstr
      End If
    Next
  Else
    MsgBox "No Entries To Display.", vbOKOnly, "No Entries"
  End If
  Close CJEditFileNum
End Sub

Private Sub lstCDEntries_DblClick()
  Dim TempRec As Integer
  TempRec = Mid$(lstCDEntries, 82)
  frmCashDisbEntry.Rec2Form (TempRec)
  Unload frmCDListing
End Sub


