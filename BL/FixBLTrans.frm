VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BL Fixer"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   1515
      Width           =   1440
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "BL Customer Utility."
      Height          =   660
      Left            =   390
      TabIndex        =   0
      Top             =   420
      Width           =   4275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Dim ExitOK As Boolean
Dim DidIT As Boolean

Private Sub Command1_Click()
  If Not ExitOK Then
    Command1.Enabled = False
    Call FixBLCust
  Else
    End
  End If
End Sub

Private Sub Form_Load()
  ExitOK = False
End Sub

Private Sub FixBLCust()
  Dim BLCust As ARCustRecType
  Dim BLLen As Integer
  Dim x As Integer
  BLLen = Len(BLCust)
  
  Open "arcust.dat" For Random As #1 Len = BLLen
  Get #1, 266, BLCust
  BLCust.SortName = "FIXME"
  BLCust.Deleted = " "
  Put #1, 266, BLCust
  Close #1
  
  Command1.Enabled = True
  Command1.Caption = "EXIT"
  ExitOK = True
  Label1.Caption = "Processing Complete."

End Sub

'Private Sub FixTRDates()
'  Dim ARTran As ARTransRecType
'  Dim TRCnt As Long
'  Dim TRLen As Integer
'  Dim WTR As Long
'  Dim NumTR As Long
'  Dim BCnt As Integer
'  Dim BDate As Integer
'  Dim cAcct As String
'
'  Dim gd(1 To 10) As Integer
'
'  gd(1) = Date2Num("07/06/2007") '198
'  gd(2) = Date2Num("10/30/2008") '238
'  gd(3) = Date2Num("10/31/2008") '236
'  gd(4) = Date2Num("10/30/2008") '238
'  gd(5) = Date2Num("07/02/2009") '258
'  gd(6) = Date2Num("07/01/2009") '256
'  gd(7) = Date2Num("06/25/2009") '254
'  gd(8) = Date2Num("07/02/2009") '257
'  gd(9) = Date2Num("06/24/2009") '255
'  gd(10) = Date2Num("03/28/2014") '400
'
'  BDate = Date2Num("12/31/1979")
'
'  TRLen = Len(ARTran)
'
'  Open "artrans.dat" For Random As #1 Len = TRLen
'  TRCnt = LOF(1) / TRLen
'
'  Get #1, 1455, ARTran
'  ARTran.TransDate = gd(1)
'  Put #1, 1455, ARTran
'
'  Get #1, 1914, ARTran
'  ARTran.TransDate = gd(2)
'  Put #1, 1914, ARTran
'
'  Get #1, 1915, ARTran
'  ARTran.TransDate = gd(3)
'  Put #1, 1915, ARTran
'
'  Get #1, 1916, ARTran
'  ARTran.TransDate = gd(4)
'  Put #1, 1916, ARTran
'
'  Get #1, 2184, ARTran
'  ARTran.TransDate = gd(5)
'  Put #1, 2184, ARTran
'
'  Get #1, 2185, ARTran
'  ARTran.TransDate = gd(6)
'  Put #1, 2185, ARTran
'
'  Get #1, 2186, ARTran
'  ARTran.TransDate = gd(7)
'  Put #1, 2186, ARTran
'
'  Get #1, 2187, ARTran
'  ARTran.TransDate = gd(8)
'  Put #1, 2187, ARTran
'
'  Get #1, 2188, ARTran
'  ARTran.TransDate = gd(9)
'  Put #1, 2188, ARTran
'
'  Get #1, 3987, ARTran
'  ARTran.TransDate = gd(10)
'  Put #1, 3987, ARTran
'
'  Close
'
'  Command1.Enabled = True
'  Command1.Caption = "EXIT"
'  ExitOK = True
'  Label1.Caption = "Transaction Processing Complete."
'
'End Sub


Public Function Date2Num(TheDate$) As Integer
 'useful function throughout program...
 'takes a string date and converts into a number based on 12/31/1979
  Date2Num% = DateDiff("d", "12/31/1979", (TheDate$))
End Function

Public Function MakeRegDate(ByVal DateNumb) As String
  Dim Month As Integer, ThisDate As String
  'function does the opposite of Date2Num
  If DateNumb = -32767 Then
    MakeRegDate = "%%%%%%%%%% "
  Else
    MakeRegDate = Format(DateAdd("d", (DateNumb), "12-31-1979"), "mm/dd/yyyy")
  End If
End Function

