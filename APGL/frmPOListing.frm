VERSION 5.00
Begin VB.Form frmPOListing 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order  Entries Listing"
   ClientHeight    =   3504
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   9504
   Icon            =   "frmPOListing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3504
   ScaleWidth      =   9504
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
      Left            =   7746
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2904
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
      Left            =   6426
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2904
      Width           =   1092
   End
   Begin VB.ListBox lstPOEntries 
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
      Left            =   240
      TabIndex        =   0
      Top             =   432
      Width           =   9084
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "P/O No."
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
      Left            =   4872
      TabIndex        =   7
      Top             =   96
      Width           =   804
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
      Index           =   0
      Left            =   192
      TabIndex        =   6
      Top             =   2952
      Width           =   6132
   End
   Begin VB.Line Line1 
      X1              =   192
      X2              =   9072
      Y1              =   336
      Y2              =   336
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
      Left            =   8160
      TabIndex        =   5
      Top             =   96
      Width           =   852
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "- Vendor -"
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
      Left            =   1656
      TabIndex        =   4
      Top             =   96
      Width           =   996
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
      Left            =   6504
      TabIndex        =   3
      Top             =   96
      Width           =   492
   End
End
Attribute VB_Name = "frmPOListing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Dim POEdit As POFORMRecType2
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Private Temp_Class As Resize_Class
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
'    Case vbKeyDown:
'      SendKeys "{Tab}"
'      KeyCode = 0
'    Case vbKeyUp:
'      SendKeys "+{Tab}"
'      KeyCode = 0
    Case vbKeyEscape:
      cmdExit_Click
      KeyCode = 0
    Case vbKeyReturn:
      cmdOk_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub cmdExit_Click()
  Unload frmPOListing
End Sub
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'  If ((UnloadMode = vbFormControlMenu)) Then
'    Cancel = True
'  End If
'End Sub

Private Sub cmdOk_Click()
  Dim TempRec As Integer
  If Not lstPOEntries = "" Then
    frmPOEnterEdit.ClearFields
    TempRec = Mid$(lstPOEntries, 82)
    frmPOEnterEdit.Rec2Form (TempRec)
  End If
  Unload frmPOListing
End Sub

Private Sub Form_Activate()
  lstPOEntries.ListIndex = 0
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  ListEntries
End Sub

Private Sub ListEntries()
  Dim POEditFile As Integer, NumEdTrans As Integer
  Dim cnt As Integer, cntg As Integer
  Dim fmt As String
  Dim tempstr As String
  Dim strInfo As String
  Dim disdate As String
  fmt = "$########.##"
  OpenPOEditFile POEditFile, NumEdTrans
  If NumEdTrans > 0 Then
    For cnt = 1 To NumEdTrans
      Get POEditFile, cnt, POEdit
      If QPTrim(POEdit.PONum) = "N/A" Then
        If POEdit.Deleted <> True Then
          cntg = cntg + 1
        End If
      End If
    Next
  End If
  If cntg > 0 Then
    For cnt = 1 To NumEdTrans
      Get POEditFile, cnt, POEdit
      If POEdit.Deleted <> True Then
        If QPTrim(POEdit.PONum) = "N/A" Then
        tempstr = Space$(90)
        Mid$(tempstr, 2) = Left$(POEdit.VNDRCODE, 10)
        Mid$(tempstr, 13) = Left$(POEdit.VNDRINF1, 25)
        Mid$(tempstr, 41) = Left$(POEdit.PONum, 8)
        Mid$(tempstr, 50) = Format(DateAdd("d", (POEdit.PODATE), "12-31-1979"), "mm/dd/yyyy")
        Mid$(tempstr, 62) = Using(fmt, POEdit.POAmt)
        Mid$(tempstr, 82) = Str$(cnt)
        lstPOEntries.AddItem tempstr
        End If
      End If
    Next
  Else
    MsgBox "No Entries To Display.", vbOKOnly, "No Entries"
    
  End If
  Close POEditFile
End Sub

Private Sub lstPOEntries_DblClick()
  Dim TempRec As Integer
  frmPOEnterEdit.ClearFields
  TempRec = Mid$(lstPOEntries, 82)
  frmPOEnterEdit.Rec2Form (TempRec)
  Unload frmPOListing
End Sub

Private Sub lstPOEntries_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim TempRec As Integer
  If KeyCode = vbKeyReturn Then
    If Not lstPOEntries = "" Then
      frmPOEnterEdit.ClearFields
      TempRec = Mid$(lstPOEntries, 82)
      frmPOEnterEdit.Rec2Form (TempRec)
      Unload frmPOListing
    End If
  End If
End Sub
