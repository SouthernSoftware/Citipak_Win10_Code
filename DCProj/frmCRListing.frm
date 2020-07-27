VERSION 5.00
Begin VB.Form frmCRListing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Receipt Listing"
   ClientHeight    =   3372
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   9588
   Icon            =   "frmCRListing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3372
   ScaleWidth      =   9588
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
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
      Left            =   7752
      TabIndex        =   2
      Top             =   2856
      Width           =   1092
   End
   Begin VB.CommandButton cmdOk 
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
      Left            =   6432
      TabIndex        =   1
      Top             =   2856
      Width           =   1092
   End
   Begin VB.ListBox lstCREntries 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1968
      Left            =   240
      TabIndex        =   0
      Top             =   336
      Width           =   9084
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
      Left            =   192
      TabIndex        =   8
      Top             =   2904
      Width           =   6132
   End
   Begin VB.Line Line1 
      X1              =   192
      X2              =   9072
      Y1              =   288
      Y2              =   288
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
      Left            =   8064
      TabIndex        =   7
      Top             =   48
      Width           =   852
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
      Left            =   2064
      TabIndex        =   6
      Top             =   48
      Width           =   660
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
      Left            =   672
      TabIndex        =   5
      Top             =   48
      Width           =   492
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
      Left            =   4416
      TabIndex        =   4
      Top             =   48
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
      Left            =   5952
      TabIndex        =   3
      Top             =   48
      Width           =   612
   End
End
Attribute VB_Name = "frmCRListing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Dim GLCREd(1) As CJEditRecType
Dim CJType As Integer
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Private Temp_Class As Resize_Class
Private Sub Form_Activate()
  lstCREntries.ListIndex = 0
End Sub

Private Sub cmdExit_Click()
  Unload frmCRListing
End Sub
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

Private Sub cmdOk_Click()
  Dim TempRec As Integer
  If Not lstCREntries = "" Then
    TempRec = Mid$(lstCREntries, 82)
    frmCashReceiptEntry.Rec2Form (TempRec)
  End If
  Unload frmCRListing
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  ListEntries
End Sub

Private Sub ListEntries()
'this fills the listbox with Cash Disbursement Entries for User to Select from or view
  Dim CJREditFile As Integer, CJEditFileNum As Integer, NumEdTrans As Integer
  Dim cnt As Integer
  Dim fmt As String
  Dim tempstr As String
  Dim strInfo As String
  Dim disdate As String
  fmt = "$########.##"
  CJType = 1
  OpenCJEditFile CJEditFileNum, NumEdTrans, CJType
  If NumEdTrans > 0 Then
    For cnt = 1 To NumEdTrans
      Get CJEditFileNum, cnt, GLCREd(1)
      If GLCREd(1).DelFlag = 0 Then
        tempstr = Space$(90)
        Mid$(tempstr, 2) = Format(DateAdd("d", (GLCREd(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
        Mid$(tempstr, 15) = Left$(GLCREd(1).Desc, 20)
        Mid$(tempstr, 38) = Left$(GLCREd(1).DOCREF, 8)
        Mid$(tempstr, 50) = Left$(GLCREd(1).RECCODE, 2)
        Mid$(tempstr, 61) = Using(fmt, GLCREd(1).Amt)
        Mid$(tempstr, 82) = Str$(cnt)
        lstCREntries.AddItem tempstr
      End If
    Next
  Else
    MsgBox "No Entries To Display.", vbOKOnly, "No Entries"
  End If
  Close CJEditFileNum
End Sub

Private Sub lstCREntries_DblClick()
  Dim TempRec As Integer
  TempRec = Mid$(lstCREntries, 82)
  frmCashReceiptEntry.Rec2Form (TempRec)
  Unload frmCRListing
End Sub



