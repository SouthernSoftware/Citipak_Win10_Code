VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSubFundTotals 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ActiveReport1"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "arSubFundTotals.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15637
   SectionData     =   "arSubFundTotals.dsx":08CA
End
Attribute VB_Name = "arSubFundTotals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private TFile As Integer
Dim EndReport As Boolean
Dim DedCnt As Integer

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  EndReport = False
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    MsgBox "Error Number: " & Str(Number) & " " & Description, vbOKOnly, "Printer Error"
    Unload Me
  End If
  CancelDisplay = True 'removes the error message
End Sub
Private Sub ActiveReport_Initialize()
  ToolBar.Tools.Add "Exit"
  ToolBar.Font.Size = 10
  
End Sub
Private Sub ActiveReport_DataInitialize()
  TFile = FreeFile
  Open StartPath & "\PRRPTS\DISTFUNDNUMTOTALS.RPT" For Input As #TFile
  
  Fields.Add ("fldGTFundNum") '(0)
  Fields.Add ("fldGTFundFed") '(1)
  Fields.Add ("fldGTFundSta") '(2)
  Fields.Add ("fldGTFundMed") '(3)
  Fields.Add ("fldGTFundSoc") '(4)
  Fields.Add ("fldGTFundRet") '(5)
  Fields.Add ("fldGTFundDedAmt1") '(6)
  Fields.Add ("fldGTFundDedAmt2") '(7)
  Fields.Add ("fldGTFundDedAmt3") '(8)
  Fields.Add ("fldGTFundDedAmt4") '(9)
  Fields.Add ("fldGTFundDedAmt5") '(10)
  Fields.Add ("fldGTFundDedAmt6") '(11)
  Fields.Add ("fldGTFundDedAmt7") '(12)
  Fields.Add ("fldGTFundDedAmt8") '(13)
  Fields.Add ("fldGTFundDedAmt9") '(14)
  Fields.Add ("fldGTFundDedAmt10") '(15)
  Fields.Add ("fldGTFundDedAmt11") '(16)
  Fields.Add ("fldGTFundDedAmt12") '(17)
  Fields.Add ("fldGTFundDedAmt13") '(18)
  Fields.Add ("fldGTFundDedAmt14") '(19)
  Fields.Add ("fldGTFundDedAmt15") '(20)
  Fields.Add ("fldGTFundDedAmt16") '(21)
  Fields.Add ("fldGTFundDedAmt17") '(22)
  Fields.Add ("fldGTFundDedAmt18") '(23)
  Fields.Add ("fldGTFundDedAmt19") '(24)
  Fields.Add ("fldGTFundDedAmt20") '(25)
  Fields.Add ("fldGTFundDedAmt21") '(26)
  Fields.Add ("fldGTFundDedAmt22") '(27)
  Fields.Add ("fldGTFundDedAmt23") '(28)
  Fields.Add ("fldGTFundDedAmt24") '(29)
  Fields.Add ("fldGTFundDedAmt25") '(30)
  Fields.Add ("fldGTFundDedAmt26") '(31)
  Fields.Add ("fldGTFundDedAmt27") '(32)
  Fields.Add ("fldGTFundDedAmt28") '(33)
  Fields.Add ("fldGTFundDedAmt29") '(34)
  Fields.Add ("fldGTFundDedAmt30") '(35)
  Fields.Add ("fldGTFundDedAmt31") '(36)
  Fields.Add ("fldGTFundDedAmt32") '(37)
  Fields.Add ("fldGTFundDedAmt33") '(38)
  Fields.Add ("fldGTFundDedAmt34") '(39)
  Fields.Add ("fldGTFundDedAmt35") '(40)
  Fields.Add ("fldGTFundDedAmt36") '(41)
  Fields.Add ("fldGTFundDedAmt37") '(42)
  Fields.Add ("fldGTFundDedAmt38") '(43)
  Fields.Add ("fldGTFundDedAmt39") '(44)
  Fields.Add ("fldGTFundDedAmt40") '(45)
  Fields.Add ("fldGTFundDedAmt41") '(46)
  Fields.Add ("fldGTFundDedAmt42") '(47)
  Fields.Add ("fldGTFundDedAmt43") '(48)
  Fields.Add ("fldGTFundDedAmt44") '(49)
  Fields.Add ("fldGTFundDedAmt45") '(50)
  Fields.Add ("fldGTFundDedAmt46") '(51)
  Fields.Add ("fldGTFundDedAmt47") '(52)
  Fields.Add ("fldGTFundDedAmt48") '(53)
  Fields.Add ("fldGTFundDedAmt49") '(54)
  Fields.Add ("fldGTFundDedAmt50") '(55)
  
  Fields.Add ("fldEmployer") '(56)
  Fields.Add ("fldDate") '(57)
  
  End Sub
Private Sub ActiveReport_FetchData(eof As Boolean)
  Dim tLine As String
  Dim arrT() As String
  If VBA.eof(TFile) Then
    eof = True
    Exit Sub
  Else
    eof = False
  End If
  Line Input #TFile, tLine
  arrT = Split(tLine, "~")
  Fields("fldGTFundNum").Value = arrT(0)
  Fields("fldGTFundFed").Value = arrT(1)
  Fields("fldGTFundSta").Value = arrT(2)
  Fields("fldGTFundMed").Value = arrT(3)
  Fields("fldGTFundSoc").Value = arrT(4)
  Fields("fldGTFundRet").Value = arrT(5)
  Fields("fldGTFundDedAmt1").Value = arrT(6)
  Fields("fldGTFundDedAmt2").Value = arrT(7)
  Fields("fldGTFundDedAmt3").Value = arrT(8)
  Fields("fldGTFundDedAmt4").Value = arrT(9)
  Fields("fldGTFundDedAmt5").Value = arrT(10)
  Fields("fldGTFundDedAmt6").Value = arrT(11)
  Fields("fldGTFundDedAmt7").Value = arrT(12)
  Fields("fldGTFundDedAmt8").Value = arrT(13)
  Fields("fldGTFundDedAmt9").Value = arrT(14)
  Fields("fldGTFundDedAmt10").Value = arrT(15)
  Fields("fldGTFundDedAmt11").Value = arrT(16)
  Fields("fldGTFundDedAmt12").Value = arrT(17)
  Fields("fldGTFundDedAmt13").Value = arrT(18)
  Fields("fldGTFundDedAmt14").Value = arrT(19)
  Fields("fldGTFundDedAmt15").Value = arrT(20)
  Fields("fldGTFundDedAmt16").Value = arrT(21)
  Fields("fldGTFundDedAmt17").Value = arrT(22)
  Fields("fldGTFundDedAmt18").Value = arrT(23)
  Fields("fldGTFundDedAmt19").Value = arrT(24)
  Fields("fldGTFundDedAmt20").Value = arrT(25)
  Fields("fldGTFundDedAmt21").Value = arrT(26)
  Fields("fldGTFundDedAmt22").Value = arrT(27)
  Fields("fldGTFundDedAmt23").Value = arrT(28)
  Fields("fldGTFundDedAmt24").Value = arrT(29)
  Fields("fldGTFundDedAmt25").Value = arrT(30)
  Fields("fldGTFundDedAmt26").Value = arrT(31)
  Fields("fldGTFundDedAmt27").Value = arrT(32)
  Fields("fldGTFundDedAmt28").Value = arrT(33)
  Fields("fldGTFundDedAmt29").Value = arrT(34)
  Fields("fldGTFundDedAmt30").Value = arrT(35)
  Fields("fldGTFundDedAmt31").Value = arrT(36)
  Fields("fldGTFundDedAmt32").Value = arrT(37)
  Fields("fldGTFundDedAmt33").Value = arrT(38)
  Fields("fldGTFundDedAmt34").Value = arrT(39)
  Fields("fldGTFundDedAmt35").Value = arrT(40)
  Fields("fldGTFundDedAmt36").Value = arrT(41)
  Fields("fldGTFundDedAmt37").Value = arrT(42)
  Fields("fldGTFundDedAmt38").Value = arrT(43)
  Fields("fldGTFundDedAmt39").Value = arrT(44)
  Fields("fldGTFundDedAmt40").Value = arrT(45)
  Fields("fldGTFundDedAmt41").Value = arrT(46)
  Fields("fldGTFundDedAmt42").Value = arrT(47)
  Fields("fldGTFundDedAmt43").Value = arrT(48)
  Fields("fldGTFundDedAmt44").Value = arrT(49)
  Fields("fldGTFundDedAmt45").Value = arrT(50)
  Fields("fldGTFundDedAmt46").Value = arrT(51)
  Fields("fldGTFundDedAmt47").Value = arrT(52)
  Fields("fldGTFundDedAmt48").Value = arrT(53)
  Fields("fldGTFundDedAmt49").Value = arrT(54)
  Fields("fldGTFundDedAmt50").Value = arrT(55)

  Fields("fldEmployer").Value = arrT(56)
  Fields("fldDate").Value = arrT(57)

End Sub
Private Sub ActiveReport_ReportEnd()
  If TFile <> 0 Then
    Close #TFile
  End If
End Sub

Private Sub ReportFooter_Format()
End Sub

Private Sub ReportHeader_Format()
End Sub

Private Sub Detail_Format()
  Dim DedRec As DedCodeRecType
  Dim DHandle As Integer
  Dim X As Integer
  
  OpenDedCodeFile DHandle
  DedCnt = LOF(DHandle) / Len(DedRec)
  Close DHandle
  
  Select Case DedCnt
  Case 0 To 10
    Detail.Height = 700
  Case 11 To 20
    Detail.Height = 900
  Case 21 To 30
    Detail.Height = 1100
  Case 31 To 40
    Detail.Height = 1300
  Case 41 To 50
    Detail.Height = 1600
  Case Else
    Detail.Height = 1500
  End Select

End Sub
