VERSION 5.00
Begin VB.Form frmInvListing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Edit Listing"
   ClientHeight    =   2976
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   9264
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2976
   ScaleWidth      =   9264
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
      Left            =   7866
      TabIndex        =   2
      Top             =   2304
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
      Left            =   6546
      TabIndex        =   1
      Top             =   2304
      Width           =   1092
   End
   Begin VB.ListBox lstInvEntries 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1488
      Left            =   90
      TabIndex        =   0
      Top             =   312
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
      Left            =   306
      TabIndex        =   7
      Top             =   2352
      Width           =   6132
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8880
      Y1              =   240
      Y2              =   240
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
      Left            =   7872
      TabIndex        =   6
      Top             =   0
      Width           =   852
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
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
      Left            =   1032
      TabIndex        =   5
      Top             =   0
      Width           =   660
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Date"
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
      Left            =   5232
      TabIndex        =   4
      Top             =   0
      Width           =   1308
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice "
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
      Left            =   3264
      TabIndex        =   3
      Top             =   0
      Width           =   732
   End
End
Attribute VB_Name = "frmInvListing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Dim APIED As APInv85Type
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  Unload frmInvListing
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
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%O"
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub cmdOk_Click()
  Dim TempRec As Integer
  If Not lstInvEntries = "" Then
    frmInvEnterEdit.ClearScn
    TempRec = Mid$(lstInvEntries, 82)
    frmInvEnterEdit.Rec2form (TempRec)
  End If
  Unload frmInvListing
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  ListEntries
End Sub

Private Sub ListEntries()
'this fills the listbox with iNVOICE Entries for User to Select from or view
  Dim APEditFile As Integer, NumEdTrans As Integer
  Dim cnt As Integer
  Dim fmt As String
  Dim tempstr As String
  Dim strInfo As String
  Dim disdate As String
  fmt = "$########.##"
  OpenAPEditFile APEditFile, NumEdTrans
  If NumEdTrans > 0 Then
    For cnt = 1 To NumEdTrans
      Get APEditFile, cnt, APIED
      If APIED.DELFLAG = 0 Then
        tempstr = Space$(90)
        Mid$(tempstr, 2) = APIED.VENDNAME
        Mid$(tempstr, 33) = APIED.INVNUM
        Mid$(tempstr, 45) = Format(DateAdd("d", (APIED.INVDATE), "12-31-1979"), "mm/dd/yyyy")
        Mid$(tempstr, 60) = Using(fmt, APIED.INVAMT)
        Mid$(tempstr, 82) = Str$(cnt)
        lstInvEntries.AddItem tempstr
      End If
    Next
  Else
    MsgBox "No Entries To Display.", vbOKOnly, "No Entries"
  End If
  Close APEditFile
End Sub

Private Sub lstInvEntries_DblClick()
  Dim TempRec As Integer
  frmInvEnterEdit.ClearScn
  TempRec = Mid$(lstInvEntries, 82)
  frmInvEnterEdit.Rec2form (TempRec)
  Unload frmInvListing
End Sub



