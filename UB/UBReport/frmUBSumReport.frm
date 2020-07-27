VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Southern Software"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5475
   Icon            =   "frmUBSumReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbYears 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2910
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GO"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   2145
      Width           =   1155
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Year to export:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1230
      TabIndex        =   4
      Top             =   1470
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   840
      TabIndex        =   2
      Top             =   780
      Width           =   3795
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "UB Sales by Year Export"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   705
      TabIndex        =   1
      Top             =   270
      Width           =   4020
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefInt A-Z
Dim Ok2Exit As Boolean
Dim DoneMsgFlag As Boolean
Dim DidCnt As Long
Dim RevCnt As Integer

Private Sub Command1_Click()
  If Not Ok2Exit Then
    Command1.Caption = "Wait"
    Command1.Enabled = False
    Call AddZeros
    Call AllDone
  Else
    End
  End If
End Sub

Private Sub Form_Load()
  Ok2Exit = False
  frmSplash.Hide
  DoEvents
  Call LoadComboBox
End Sub

Private Sub LoadComboBox()
Dim fillCnt As Integer
Dim cnt As Integer

fillCnt = UBound(BillYears())
For cnt = 1 To fillCnt
  cbYears.AddItem (Str$(BillYears(cnt)))
Next
cbYears.ListIndex = 0

End Sub

Private Sub AddZeros()
    
  Dim UBCust As NewUBCustRecType
  Dim UBCustN As NewUBCustRecType
  Dim CustLen As Integer, XCnt As Integer
  Dim RecCnt As Long, LopCnt As Long
  Dim zzCnt As Integer
  Dim TransRec As UBTransRecType
  Dim TRLen As Integer
  Dim TRevAmt As Double
  TRLen = Len(TransRec)
  CustLen = Len(UBCust)
   
  Dim GTotal As Double
  Dim RTotal As Double
  Dim BillCnt As Integer
   
  'ReDim CustRevs(1 To 1) As UBCustRevTotalsType
  Dim VBar$
  VBar$ = ","
  DidCnt = 1
  Dim ChkCnt As Integer
  'Dim GotNeg As Boolean
  'Dim GotPos As Boolean
  Dim PCnt As Integer
  Dim ThisTR As Long
  Dim BFYear As Integer, BEYear As Integer
  Dim SelYear As String
  Dim FirstTR As Boolean
  ReDim CustYearTotals(1 To 1) As UBCustRevenueTotalsType
  
  Call FillRevList

  SelYear = QPTrim(cbYears.List(cbYears.ListIndex))
  
  BFYear = Date2Num("01/01/" + SelYear)
  BEYear = Date2Num("12/31/" + SelYear)
  
  Open "ubcust.dat" For Random As #1 Len = CustLen
  Open "ubtrans.dat" For Random As #3 Len = TRLen
  Open "UB_Sales_" + SelYear + ".csv" For Output As #2
  
  Label2.Caption = "Processing. . 0%"
  DoEvents
  
  DidCnt = 0
  RecCnt = LOF(1) / CustLen
  For LopCnt = 1 To RecCnt   'RecCnt
    Get #1, LopCnt, UBCust
    FirstTR = True
    BillCnt = 0
    If UBCust.DelFlag = 0 Then
      ThisTR = UBCust.LastTrans
      Do While ThisTR > 0
        Get #3, ThisTR, TransRec
        If TransRec.TransType = 1 Then
          If TransRec.TransDate >= BFYear And TransRec.TransDate <= BEYear Then
            If FirstTR Then
              FirstTR = False
              DidCnt = DidCnt + 1
              BillCnt = BillCnt + 1
              ReDim Preserve CustYearTotals(1 To DidCnt) As UBCustRevenueTotalsType
              CustYearTotals(DidCnt).CustRec = LopCnt
              CustYearTotals(DidCnt).CustName = FixCustName(UBCust.CustName)
              For zzCnt = 1 To RevCnt
                CustYearTotals(DidCnt).RevTotal(zzCnt) = TransRec.RevAmt(zzCnt)
              Next
            Else
              For zzCnt = 1 To RevCnt
                CustYearTotals(DidCnt).RevTotal(zzCnt) = CustYearTotals(DidCnt).RevTotal(zzCnt) + TransRec.RevAmt(zzCnt)
              Next
              BillCnt = BillCnt + 1
            End If
          End If
        End If
        ThisTR = TransRec.PrevTrans
      Loop
    End If
    If LopCnt Mod 10 = 0 Then
      Label2.Caption = "Processing. ." + ShowPctComp(LopCnt, RecCnt)
      DoEvents
    End If
    If BillCnt > 0 Then
      CustYearTotals(DidCnt).BilledCnt = BillCnt
    End If
  Next

  Close 1, 3

  Print #2, "ACCT"; VBar; "Exp Year"; VBar; "Cust Name"; VBar; "BilledCnt"; VBar;
  For zzCnt = 1 To RevCnt
    Print #2, RevList(zzCnt); VBar;
  Next
  Print #2, "PctOfSales"
  
  GTotal# = 0
  For LopCnt = 1 To DidCnt
    For zzCnt = 1 To RevCnt
      GTotal = Round#(GTotal + CustYearTotals(LopCnt).RevTotal(zzCnt))
    Next
  Next
  For LopCnt = 1 To DidCnt
    RTotal# = 0
    For zzCnt = 1 To RevCnt
      RTotal# = RTotal# + CustYearTotals(LopCnt).RevTotal(zzCnt)
    Next
    CustYearTotals(LopCnt).PctOfSales = (RTotal# / GTotal#) * 100
  Next

  For LopCnt = 1 To DidCnt
    Print #2, CustYearTotals(LopCnt).CustRec; VBar;
    Print #2, SelYear; VBar; QPTrim(CustYearTotals(LopCnt).CustName); VBar; CustYearTotals(LopCnt).BilledCnt; VBar;
    For zzCnt = 1 To RevCnt
       Print #2, CustYearTotals(LopCnt).RevTotal(zzCnt); VBar;
    Next
    Print #2, Format(CustYearTotals(LopCnt).PctOfSales, "#.##########")
  Next
Close
End Sub

Private Sub AllDone()
  If DoneMsgFlag Then
    Command1.Caption = "Already Converted."
  Else
    Label2.Caption = "Exported:" + Str$(DidCnt) + " Records."
    Command1.Caption = "Exit"
  End If
  Command1.Enabled = True
  Ok2Exit = True
End Sub

Private Sub FillRevList()
  ReDim UBSetUpRec(1) As UBSetupRecType
  Dim UBSetupLen%
  UBSetupLen = Len(UBSetUpRec(1))
  
  Open "UBSETUP.DAT" For Random Shared As #1 Len = UBSetupLen    'open data file
  Get #1, 1, UBSetUpRec(1)
  Close #1

  For RevCnt = 1 To 15
    If QPTrim(UBSetUpRec(1).Revenues(RevCnt).RevName) = "" Then
      Exit For
    Else
      ReDim Preserve RevList(0 To RevCnt) As String
      RevList(RevCnt) = QPTrim(UBSetUpRec(1).Revenues(RevCnt).RevName)
    End If
  Next
  RevCnt = RevCnt - 1

End Sub

Public Function FixCustName(CustNameIn As String) As String
  Dim CPos As Integer
  CPos = InStr(CustNameIn, ",")
  Do While CPos > 0
    Mid$(CustNameIn, CPos, 1) = " "
    CPos = InStr(CustNameIn, ",")
  Loop
  FixCustName = CustNameIn
  
End Function

Public Function QPTrim$(Text As String)
  Dim StrLen As Long
  Dim cnt As Long
  Dim thischar As Integer
  StrLen = Len(Text)
  For cnt = 1 To StrLen
    thischar = Asc(Mid$(Text, cnt, 1))
    If thischar = 0 Then
      Mid$(Text$, cnt, 1) = " "
    End If
  Next
  QPTrim$ = Trim$(Text)
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Close
  End
End Sub
