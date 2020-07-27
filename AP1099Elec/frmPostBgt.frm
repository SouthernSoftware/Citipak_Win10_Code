VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPostBgt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Budget Posting"
   ClientHeight    =   8640
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   12195
   Icon            =   "frmPostBgt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   375
      Left            =   2640
      Top             =   3000
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Esc E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   10080
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7560
      Width           =   1332
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 &Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   8400
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7560
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   8280
      Width           =   12192
      _ExtentX        =   21511
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7117
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "12:04 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "6/28/2008"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WARNING !"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   396
      Left            =   5190
      TabIndex        =   7
      Top             =   3024
      Width           =   1812
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3240
      Top             =   1368
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Post Budget Entries"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4080
      TabIndex        =   6
      Top             =   1608
      Width           =   4092
   End
   Begin VB.Label lblPosting 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Posting to Budget History"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   372
      Left            =   3840
      TabIndex        =   5
      Top             =   4488
      Visible         =   0   'False
      Width           =   4572
   End
   Begin VB.Label lbInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Before You Post, Make Sure You Have Printed A Budget  Edit Report.  If You Haven't, Then Exit And Do So Now."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   732
      Index           =   0
      Left            =   2760
      TabIndex        =   4
      Top             =   3456
      Width           =   6732
   End
   Begin VB.Label lbInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Ok to Begin Posting or Exit to Escape Posting Procedure. "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Index           =   1
      Left            =   2520
      TabIndex        =   3
      Top             =   4128
      Width           =   7212
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2172
      Left            =   2400
      Top             =   2808
      Width           =   7452
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      Height          =   972
      Left            =   3240
      Top             =   1248
      Width           =   5772
   End
End
Attribute VB_Name = "frmPostBgt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim BgtEdit As TrEditRecType
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim GLAcctidx As GLAcctIndexType
Dim GLAcct As GLAcctRecType
Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  KillFileD "BGTED.opn"
  Unload frmPostBgt
End Sub
Private Sub Timer1_Timer()
 ' Label2.Visible = Not Label2.Visible
  '&H0080FFFF&
  Static tog As Boolean
  tog = Not tog
  If tog Then
    Label2.ForeColor = &H80FFFF
    Shape3.BackColor = &HC0&
  Else
    Label2.ForeColor = &HFFFF&
    Shape3.BackColor = &H80&
  End If
  
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    Cancel = True
  End If
End Sub

Private Sub cmdOk_Click()
  PostBGTTrans
  KillFileD "BGTED.opn"
  Unload frmPostBgt
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"  'Arrow Down
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"   'arrow up
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"     'Esc key
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%O"     'alt O or f10
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  Me.HelpContextID = hlpPostBudgetEntries
End Sub
Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub


Private Sub PostBGTTrans()
  Dim BgtEditFileNum As Integer, NumEdTrans As Integer, NumBgtTrans As Long
  Dim BgtTransFile As Integer, BgtTransRecLen As Integer
  Dim cnt As Integer, Active As Integer, AcctRec As Integer, Prev As Long
  Dim FundCode As String, FundNum As String, ErrMsg As String
  Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer, BadTrans As Integer
  Dim AcctFile As Integer, NumAccts As Integer, CntA As Integer
  Dim TotDr As Double, TotCr As Double
  Dim BgtTrans As GLTransRecType

   '--verify that there are transactions and they are in balance.
   OpenBgtEditFile BgtEditFileNum, NumEdTrans
  If BgtEditFileNum < 0 Then
    Unload frmPostBgt
    Exit Sub
  End If
   '--summarize the file totals
   For cnt = 1 To NumEdTrans
      Get BgtEditFileNum, cnt, BgtEdit
      If Not BgtEdit.Deleted Then
         Active = Active + 1
         TotDr# = Round#(TotDr# + BgtEdit.DrAmt)
         TotCr# = Round#(TotCr# + BgtEdit.CrAmt)
      End If
   Next
   
   

   '--if no active transactions tell user and get out
   If Active = 0 Then
      MsgBox "There Are No Transactions to Post.", vbOKOnly, "Budget Posting"
      Close
      Exit Sub
   End If

   If MsgBox("Are You Sure You Wish To Post Now?", vbYesNo, "Budget Posting") = vbNo Then    'Ask user if sure ready to post
     Exit Sub
   End If
   'Transactions out of balance ask user if ok to post
   If TotDr# <> TotCr# Then
     If MsgBox("The Debits And Credits Are Out Of Balance, Do You Wish To Post Or Cancel?", vbOKCancel, "Budget Entries OutOfBalance") = vbCancel Then
       Close
       Exit Sub
     End If
   End If
   TotDr# = 0                             'init totals to zero
   TotCr# = 0
   Active = 0                             'Counter for Active Transactions

  
   OpenAcctFile AcctFile
   NumAccts = LOF(AcctFile) / Len(GLAcct)
   BgtTransFile = FreeFile
   BgtTransRecLen = Len(BgtTrans)
   Open "BgtTrans.DAT" For Random As BgtTransFile Len = BgtTransRecLen

   NumBgtTrans = LOF(BgtTransFile) \ BgtTransRecLen

   For cnt = 1 To NumEdTrans              'Assign edit file to trans format
      Get BgtEditFileNum, cnt, BgtEdit

      If Not BgtEdit.Deleted Then
         Active = Active + 1
         AcctRec = AcctFind(BgtEdit.AcctNum)
         If AcctRec > 0 Then
            Get AcctFile, AcctRec, GLAcct

            Select Case GLAcct.Typ
               Case "E"
                  GLAcct.Bgt = Round(GLAcct.Bgt + BgtEdit.DrAmt - BgtEdit.CrAmt)
               Case "R"
                  GLAcct.Bgt = Round(GLAcct.Bgt + BgtEdit.CrAmt - BgtEdit.DrAmt)

            End Select

            Put AcctFile, AcctRec, GLAcct

            BgtTrans.AcctRec = BgtEdit.AcctRec
            BgtTrans.AcctNum = BgtEdit.AcctNum
            BgtTrans.TRDATE = BgtEdit.TRDATE
            BgtTrans.Desc = BgtEdit.Desc
            BgtTrans.LDesc = BgtEdit.LDesc
            BgtTrans.Ref = BgtEdit.Ref
            BgtTrans.DrAmt = BgtEdit.DrAmt
            BgtTrans.CrAmt = BgtEdit.CrAmt
            BgtTrans.Src = "BG" + Format$(Now, "mmddyy")
            BgtTrans.NextTran = 0

            NumBgtTrans = NumBgtTrans + 1

            Put BgtTransFile, NumBgtTrans, BgtTrans

            '--------------------------------Start linking here
            If GLAcct.FrstBTran = 0 Then       'if first trans for this acct,
               GLAcct.FrstBTran = NumBgtTrans  'assign first & last pointers to
               GLAcct.LastBTran = NumBgtTrans  'this transaction
               Put AcctFile, AcctRec, GLAcct

            Else                             'otherwise
                                             'in the account file..
               Prev = GLAcct.LastBTran         'remember the prev trans pointer,
               GLAcct.LastBTran = NumBgtTrans  'reset last trans to this trans

               Put AcctFile, AcctRec, GLAcct

                                             'In the trans file...
               Get BgtTransFile, Prev, BgtTrans  'Get the last transaction
               BgtTrans.NextTran = NumBgtTrans     'reset pointer to this tran
               Put BgtTransFile, Prev, BgtTrans

           End If

         Else
            BadTrans = BadTrans + 1

         End If
      End If
   Next
   Close
   If BadTrans > 0 Then Beep
   KillFile "BGTED.dat"
   Call MainLog("Budget Post Complete.")
   MsgBox "Budget Posting Is Complete.", vbOKOnly, "Budget Posting"
End Sub


