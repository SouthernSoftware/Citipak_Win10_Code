VERSION 5.00
Begin VB.Form frmPenaltyList1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penalty Transaction List"
   ClientHeight    =   3372
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   9588
   ControlBox      =   0   'False
   Icon            =   "frmPenaltyList1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3372
   ScaleWidth      =   9588
   ShowInTaskbar   =   0   'False
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
   Begin VB.ListBox lstPenalties 
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
      TabIndex        =   7
      Top             =   2904
      Width           =   6132
   End
   Begin VB.Line Line1 
      X1              =   192
      X2              =   9336
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
      Left            =   8016
      TabIndex        =   6
      Top             =   48
      Width           =   852
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
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
      Left            =   1920
      TabIndex        =   5
      Top             =   48
      Width           =   1044
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account"
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
      Left            =   480
      TabIndex        =   4
      Top             =   48
      Width           =   852
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
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
      Left            =   3960
      TabIndex        =   3
      Top             =   48
      Width           =   1644
   End
End
Attribute VB_Name = "frmPenaltyList1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
'Private Sub Form_Activate()
'  lstPenalties.ListIndex = 0
'End Sub

Private Sub cmdExit_Click()
  Unload frmPenaltyList
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
  If Not lstPenalties = "" Then
    TempRec = Mid$(lstPenalties, 82)
    frmPenaltyEdit.PenaltyRec2Screen (TempRec)
  End If
  Unload frmPenaltyList
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  ListEntries
End Sub

Private Sub ListEntries()
'this fills the listbox with Penalty Trans Entries for User to Select from or view
  Dim cnt As Integer, PHandle As Integer, CHandle As Integer
  Dim fmt As String, PenFile As String, UBTranRecLen As Integer
  Dim tempstr As String, UBCustRecLen As Integer, NumPenRec As Long
  Dim strInfo As String, lcnt As Long
  Dim disdate As String
  ReDim UBCustRec(1) As NewUBCustRecType
  ReDim UBTranRec(1) As UBTransRecType

  UBCustRecLen = Len(UBCustRec(1))
  UBTranRecLen = Len(UBTranRec(1))
  PenFile$ = UBPath$ + "UBPENTRN.DAT"
  PHandle = FreeFile
  Open PenFile$ For Random Shared As PHandle Len = UBTranRecLen
  CHandle = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CHandle Len = UBCustRecLen
  NumPenRec& = LOF(PHandle) / UBTranRecLen


  fmt = "$########.##"
  If NumPenRec& > 0 Then
    For lcnt& = 1 To NumPenRec&
      Get PHandle, lcnt&, UBTranRec(1)
      If UBTranRec(1).CustAcctNo > 0 Then
      Get CHandle, UBTranRec(1).CustAcctNo, UBCustRec(1)
      If Not UBCustRec(1).DelFlag Then
        tempstr = Space$(90)
        Mid$(tempstr, 1) = Using$("######", UBTranRec(1).CustAcctNo)
        Mid$(tempstr, 14) = (UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB)
        Mid$(tempstr, 28) = Left$(UBCustRec(1).CustName, 30)
        If UBTranRec(1).Transamt = 0 Then
          'Mid$(tempstr, 59) = Using(fmt, "0")
          Mid$(tempstr, 59) = "DELETED"
        Else
          Mid$(tempstr, 59) = Using(fmt, UBTranRec(1).Transamt)
          'Mid$(tempstr, 61) = " "
        End If
        Mid$(tempstr, 82) = Str$(lcnt&)
        lstPenalties.AddItem tempstr
      End If
      End If
    Next
  Else
    MsgBox "No Entries To Display.", vbOKOnly, "No Entries"
  End If
  Close
End Sub

Private Sub lstPenalties_DblClick()
  Dim TempRec As Integer
  TempRec = Mid$(lstPenalties, 82)
  frmPenaltyEdit.PenaltyRec2Screen (TempRec)
  Unload frmPenaltyList
End Sub



