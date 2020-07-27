VERSION 5.00
Begin VB.Form frmFMeterReadList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Final Meter Reading List"
   ClientHeight    =   3372
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   9588
   ControlBox      =   0   'False
   Icon            =   "frmMeterReadList.frx":0000
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
   Begin VB.ListBox lstFMeterRead 
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Location"
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
      Left            =   5544
      TabIndex        =   6
      Top             =   48
      Width           =   2268
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
      TabIndex        =   5
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
      Left            =   2400
      TabIndex        =   3
      Top             =   48
      Width           =   1644
   End
End
Attribute VB_Name = "frmFMeterReadList"
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
  Unload Me
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
  Dim TempRec As Long, listnum As Integer
  If Not lstFMeterRead = "" Then
    TempRec = Mid$(lstFMeterRead, 82)
    listnum = Mid$(lstFMeterRead, 75, 5)
    frmFinalMeterReads.FMeterRec2Screen TempRec, listnum
  End If
  Unload Me
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  ListEntries
End Sub

Private Sub ListEntries()
'this fills the listbox with Meter Read Cust for User to Select from or view
  Dim cnt As Long, C2Handle As Integer
  Dim fmt As String, NumofRecs As Long
  Dim tempstr As String, UBCustRecLen As Integer
  Dim strInfo As String, lcnt As Long
  Dim disdate As String, GoodCnt As Integer
  Dim UBCustRec As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec)
  fmt = "$########.##"

  C2Handle = FreeFile
  GoodCnt = 0
  NumofRecs = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen
  Open UBPath$ + "UBCUST.DAT" For Random Shared As C2Handle Len = UBCustRecLen
  For cnt& = 1 To NumofRecs
    Get #C2Handle, cnt&, UBCustRec
      If UBCustRec.Status = "F" Then 'Val(UBCustRec.Book) = BookNumber Then
         GoodCnt = GoodCnt + 1
         'IdxBuff(GoodCnt).RecNum = cnt&
      
         tempstr = Space$(90)
        Mid$(tempstr, 1) = Using$("######", cnt&)
        Mid$(tempstr, 14) = Left$(UBCustRec.CustName, 20)
        Mid$(tempstr, 44) = Left$(UBCustRec.SERVADDR, 30)
        Mid$(tempstr, 75) = Str$(GoodCnt)
        Mid$(tempstr, 82) = Str$(cnt&)
        lstFMeterRead.AddItem tempstr

      
      End If
    Next
  Close C2Handle
  'If GoodCnt = 0 Then
    
  'Else
    
  'End If

'**(*)(*)*)(*)(*)(*)
End Sub



Private Sub lstFMeterRead_DblClick()
  Dim TempRec As Long, listnum As Integer
  If Not lstFMeterRead = "" Then
    TempRec = Mid$(lstFMeterRead, 82)
    listnum = Mid$(lstFMeterRead, 75, 5)
    frmFinalMeterReads.FMeterRec2Screen TempRec, listnum
  End If
  Unload Me
End Sub



