VERSION 5.00
Begin VB.Form frmSwapRevs 
   Caption         =   "Revenue Switch"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   2880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCombine 
      Caption         =   "Combine B into A and Clear B"
      Height          =   405
      Left            =   90
      TabIndex        =   5
      Top             =   2640
      Width           =   2685
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Switch"
      Height          =   405
      Left            =   930
      TabIndex        =   4
      Top             =   2040
      Width           =   1125
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1890
      TabIndex        =   3
      Top             =   1260
      Width           =   600
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1890
      TabIndex        =   2
      Top             =   690
      Width           =   600
   End
   Begin VB.Label Label2 
      Caption         =   "Revenue B"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   450
      TabIndex        =   1
      Top             =   1260
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Revenue A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   450
      TabIndex        =   0
      Top             =   660
      Width           =   1155
   End
End
Attribute VB_Name = "frmSwapRevs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCombine_Click()
  If CInt(Text1.Text) > 0 And CInt(Text1.Text) < 16 Then
    If CInt(Text2.Text) > 0 And CInt(Text2.Text) < 16 Then
      Call combinetherevs(CInt(Text1.Text), CInt(Text2.Text))
    End If
  End If
End Sub

Private Sub Command1_Click()
  If CInt(Text1.Text) > 0 And CInt(Text1.Text) < 16 Then
    If CInt(Text2.Text) > 0 And CInt(Text2.Text) < 16 Then
      Call swaptherevs(CInt(Text1.Text), CInt(Text2.Text))
    End If
  End If
End Sub
'Switch revenues this time is 5 and 9
Private Sub swaptherevs(ByVal xx1 As Integer, ByVal xx2 As Integer)
  Dim UBTranRecLen As Integer, read As String, UBCustLen As Integer
  Dim UBFile As Integer, cntm As Integer, UBCust As Integer
  Dim TNumOfRecs As Long, cnt As Long, TrTyp As Integer
  Dim NumOfCustRecs As Long
  Dim revA As Integer
  Dim revB As Integer
  revA = xx1
  revB = xx2
  ReDim UBTranRec(1) As UBTransRecType
  UBTranRecLen = Len(UBTranRec(1))
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustLen = Len(UBCustRec(1))
  Dim temptot As Double
  Dim temptot2 As Double
  Dim temptotP As Double
  Dim temptotP2 As Double
  Dim rate1 As String * 4
  Dim rate2 As String * 4
  Dim mtr1 As String * 1
  Dim mtr2 As String * 1
  FrmShowPctComp.Label1 = "Switching customer setup and balances."
  FrmShowPctComp.Show , Me
    UBCust = FreeFile
    Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustLen
    NumOfCustRecs& = LOF(UBCust) \ UBCustLen
    For cnt& = 1 To NumOfCustRecs&
      temptot = 0
      temptot2 = 0
      temptotP = 0
      temptotP2 = 0
      rate1 = ""
      rate2 = ""
      mtr1 = ""
      mtr2 = ""
      
      FrmShowPctComp.ShowPctComp cnt&, NumOfCustRecs&
      Get UBCust, cnt&, UBCustRec(1)
      temptot = UBCustRec(1).CurrRevAmts(revA)
      temptot2 = UBCustRec(1).CurrRevAmts(revB)
      UBCustRec(1).CurrRevAmts(revA) = temptot2
      UBCustRec(1).CurrRevAmts(revB) = temptot
      temptotP = UBCustRec(1).PrevRevAmts(revA)
      temptotP2 = UBCustRec(1).PrevRevAmts(revB)
      UBCustRec(1).PrevRevAmts(revA) = temptotP2
      UBCustRec(1).PrevRevAmts(revB) = temptotP

      rate1 = UBCustRec(1).serv(revA).RATECODE
      mtr1 = UBCustRec(1).serv(revA).RMtrType
      rate2 = UBCustRec(1).serv(revB).RATECODE
      mtr2 = UBCustRec(1).serv(revB).RMtrType
      UBCustRec(1).serv(revA).RATECODE = rate2
      UBCustRec(1).serv(revA).RMtrType = mtr2
      UBCustRec(1).serv(revB).RATECODE = rate1
      UBCustRec(1).serv(revB).RMtrType = mtr1
          
      Put UBCust, cnt&, UBCustRec(1)
    Next
    Erase UBCustRec
    Close
    
    UBFile = FreeFile
    Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBFile Len = UBTranRecLen
    TNumOfRecs& = LOF(UBFile) / UBTranRecLen
    FrmShowPctComp.Label1 = "Switching revenue transaction balances."
    FrmShowPctComp.Show , Me
   
    For cnt& = 1 To TNumOfRecs&
    temptot = 0
    temptot2 = 0
      FrmShowPctComp.ShowPctComp cnt&, TNumOfRecs&
      Get UBFile, cnt&, UBTranRec(1)
      temptot = UBTranRec(1).RevAmt(revA)
      temptot2 = UBTranRec(1).RevAmt(revB)
      UBTranRec(1).RevAmt(revA) = temptot2
      UBTranRec(1).RevAmt(revB) = temptot
      Put UBFile, cnt&, UBTranRec(1)
    Next
  Erase UBTranRec
  Close
  If MsgBox("All finished", vbOKOnly, "Done") = vbOK Then
    Unload Me
  End If
End Sub
Private Sub combinetherevs(ByVal xx1 As Integer, ByVal xx2 As Integer)
  Dim UBTranRecLen As Integer, read As String, UBCustLen As Integer
  Dim UBFile As Integer, cntm As Integer, UBCust As Integer
  Dim TNumOfRecs As Long, cnt As Long, TrTyp As Integer
  Dim NumOfCustRecs As Long
  Dim revA As Integer
  Dim revB As Integer
  revA = xx1
  revB = xx2
  ReDim UBTranRec(1) As UBTransRecType
  UBTranRecLen = Len(UBTranRec(1))
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustLen = Len(UBCustRec(1))
  Dim temptot As Double
  Dim temptot2 As Double
  Dim temptotP As Double
  Dim temptotP2 As Double
  Dim rate1 As String * 4
  Dim rate2 As String * 4
  Dim mtr1 As String * 1
  Dim mtr2 As String * 1
  FrmShowPctComp.Label1 = "Combining customer setup and balances."
  FrmShowPctComp.Show , Me
    UBCust = FreeFile
    Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustLen
    NumOfCustRecs& = LOF(UBCust) \ UBCustLen
    For cnt& = 1 To NumOfCustRecs&
      temptot = 0
      temptot2 = 0
      temptotP = 0
      temptotP2 = 0
      rate1 = ""
      rate2 = ""
      mtr1 = ""
      mtr2 = ""
      
      FrmShowPctComp.ShowPctComp cnt&, NumOfCustRecs&
      Get UBCust, cnt&, UBCustRec(1)
      temptot = UBCustRec(1).CurrRevAmts(revA)
      temptot2 = UBCustRec(1).CurrRevAmts(revB)
      UBCustRec(1).CurrRevAmts(revA) = temptot2 + temptot
      UBCustRec(1).CurrRevAmts(revB) = 0
      temptotP = UBCustRec(1).PrevRevAmts(revA)
      temptotP2 = UBCustRec(1).PrevRevAmts(revB)
      UBCustRec(1).PrevRevAmts(revA) = temptotP2 + temptotP
      UBCustRec(1).PrevRevAmts(revB) = 0


'this was for a flat fee they added during a manual process not a regular billing
' so not set in cust maint.
'      rate1 = UBCustRec(1).serv(revA).RATECODE
'      mtr1 = UBCustRec(1).serv(revA).RMtrType
'      rate2 = UBCustRec(1).serv(revB).RATECODE
'      mtr2 = UBCustRec(1).serv(revB).RMtrType
'      UBCustRec(1).serv(revA).RATECODE = rate2
'      UBCustRec(1).serv(revA).RMtrType = mtr2
'      UBCustRec(1).serv(revB).RATECODE = rate1
'      UBCustRec(1).serv(revB).RMtrType = mtr1
          
      Put UBCust, cnt&, UBCustRec(1)
    Next
    Erase UBCustRec
    Close
    
    UBFile = FreeFile
    Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBFile Len = UBTranRecLen
    TNumOfRecs& = LOF(UBFile) / UBTranRecLen
    FrmShowPctComp.Label1 = "Combining revenue transaction balances."
    FrmShowPctComp.Show , Me
   
    For cnt& = 1 To TNumOfRecs&
    temptot = 0
    temptot2 = 0
      FrmShowPctComp.ShowPctComp cnt&, TNumOfRecs&
      Get UBFile, cnt&, UBTranRec(1)
      temptot = UBTranRec(1).RevAmt(revA)
      temptot2 = UBTranRec(1).RevAmt(revB)
      UBTranRec(1).RevAmt(revA) = temptot2 + temptot
      UBTranRec(1).RevAmt(revB) = 0
      Put UBFile, cnt&, UBTranRec(1)
    Next
  Erase UBTranRec
  Close
  If MsgBox("All finished", vbOKOnly, "Done") = vbOK Then
    Unload Me
  End If
End Sub
