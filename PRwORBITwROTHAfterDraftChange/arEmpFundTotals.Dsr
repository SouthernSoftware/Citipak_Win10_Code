VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arEmpFundTotals 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fund Data"
   ClientHeight    =   4410
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   8100
   Icon            =   "arEmpFundTotals.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   14288
   _ExtentY        =   7779
   SectionData     =   "arEmpFundTotals.dsx":08CA
End
Attribute VB_Name = "arEmpFundTotals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private FFile As Integer
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
    ''Me.Visible = False
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

Private Sub ActiveReport_FetchData(eof As Boolean)
  Dim fLine As String
  Dim arrF() As String
  
  If Not VBA.eof(FFile) Then
    eof = False
    Line Input #FFile, fLine
    arrF = Split(fLine, "~")
    Fields("fldEmpFundNum").Value = arrF(0)
    Fields("fldFundDedAmt1").Value = arrF(1)
    Fields("fldFundDedAmt2").Value = arrF(2)
    Fields("fldFundDedAmt3").Value = arrF(3)
    Fields("fldFundDedAmt4").Value = arrF(4)
    Fields("fldFundDedAmt5").Value = arrF(5)
    Fields("fldFundDedAmt6").Value = arrF(6)
    Fields("fldFundDedAmt7").Value = arrF(7)
    Fields("fldFundDedAmt8").Value = arrF(8)
    Fields("fldFundDedAmt9").Value = arrF(9)
    Fields("fldFundDedAmt10").Value = arrF(10)
    Fields("fldFundDedAmt11").Value = arrF(11)
    Fields("fldFundDedAmt12").Value = arrF(12)
    Fields("fldFundDedAmt13").Value = arrF(13)
    Fields("fldFundDedAmt14").Value = arrF(14)
    Fields("fldFundDedAmt15").Value = arrF(15)
    Fields("fldFundDedAmt16").Value = arrF(16)
    Fields("fldFundDedAmt17").Value = arrF(17)
    Fields("fldFundDedAmt18").Value = arrF(18)
    Fields("fldFundDedAmt19").Value = arrF(19)
    Fields("fldFundDedAmt20").Value = arrF(20)
    Fields("fldFundDedAmt21").Value = arrF(21)
    Fields("fldFundDedAmt22").Value = arrF(22)
    Fields("fldFundDedAmt23").Value = arrF(23)
    Fields("fldFundDedAmt24").Value = arrF(24)
    Fields("fldFundDedAmt25").Value = arrF(25)
    Fields("fldFundDedAmt26").Value = arrF(26)
    Fields("fldFundDedAmt27").Value = arrF(27)
    Fields("fldFundDedAmt28").Value = arrF(28)
    Fields("fldFundDedAmt29").Value = arrF(29)
    Fields("fldFundDedAmt30").Value = arrF(30)
    Fields("fldFundDedAmt31").Value = arrF(31)
    Fields("fldFundDedAmt32").Value = arrF(32)
    Fields("fldFundDedAmt33").Value = arrF(33)
    Fields("fldFundDedAmt34").Value = arrF(34)
    Fields("fldFundDedAmt35").Value = arrF(35)
    Fields("fldFundDedAmt36").Value = arrF(36)
    Fields("fldFundDedAmt37").Value = arrF(37)
    Fields("fldFundDedAmt38").Value = arrF(38)
    Fields("fldFundDedAmt39").Value = arrF(39)
    Fields("fldFundDedAmt40").Value = arrF(40)
    Fields("fldFundDedAmt41").Value = arrF(41)
    Fields("fldFundDedAmt42").Value = arrF(42)
    Fields("fldFundDedAmt43").Value = arrF(43)
    Fields("fldFundDedAmt44").Value = arrF(44)
    Fields("fldFundDedAmt45").Value = arrF(45)
    Fields("fldFundDedAmt46").Value = arrF(46)
    Fields("fldFundDedAmt47").Value = arrF(47)
    Fields("fldFundDedAmt48").Value = arrF(48)
    Fields("fldFundDedAmt49").Value = arrF(49)
    Fields("fldFundDedAmt50").Value = arrF(50)


    Fields("fldFundFed").Value = arrF(51)
    Fields("fldFundSta").Value = arrF(52)
    Fields("fldFundMed").Value = arrF(53)
    Fields("fldFundSoc").Value = arrF(54)
    Fields("fldFundRet").Value = arrF(55)
  Else
    eof = True
  End If


  If VBA.eof(FFile) Then Exit Sub

End Sub


Private Sub ActiveReport_Initialize()
  ToolBar.Tools.Add "Exit"
  ToolBar.Font.Size = 10
  
End Sub
Private Sub ActiveReport_DataInitialize()
  FFile = FreeFile
  Open StartPath & "\PRRPTS\DISTRIBUFUNDNUM.RPT" For Input As #FFile
  
  Fields.Add "fldEmpFundNum" '(0)
  Fields.Add "fldFundDedAmt1" '(1)
  Fields.Add "fldFundDedAmt2" '(2)
  Fields.Add "fldFundDedAmt3" '(3)
  Fields.Add "fldFundDedAmt4" '(4)
  Fields.Add "fldFundDedAmt5" '(5)
  Fields.Add "fldFundDedAmt6" '(6)
  Fields.Add "fldFundDedAmt7" '(7)
  Fields.Add "fldFundDedAmt8" '(8)
  Fields.Add "fldFundDedAmt9" '(9)
  Fields.Add "fldFundDedAmt10" '(10)
  Fields.Add "fldFundDedAmt11" '(11)
  Fields.Add "fldFundDedAmt12" '(12)
  Fields.Add "fldFundDedAmt13" '(13)
  Fields.Add "fldFundDedAmt14" '(14)
  Fields.Add "fldFundDedAmt15" '(15)
  Fields.Add "fldFundDedAmt16" '(16)
  Fields.Add "fldFundDedAmt17" '(17)
  Fields.Add "fldFundDedAmt18" '(18)
  Fields.Add "fldFundDedAmt19" '(19)
  Fields.Add "fldFundDedAmt20" '(20)
  Fields.Add "fldFundDedAmt21" '(21)
  Fields.Add "fldFundDedAmt22" '(22)
  Fields.Add "fldFundDedAmt23" '(23)
  Fields.Add "fldFundDedAmt24" '(24)
  Fields.Add "fldFundDedAmt25" '(25)
  Fields.Add "fldFundDedAmt26" '(26)
  Fields.Add "fldFundDedAmt27" '(27)
  Fields.Add "fldFundDedAmt28" '(28)
  Fields.Add "fldFundDedAmt29" '(29)
  Fields.Add "fldFundDedAmt30" '(30)
  Fields.Add "fldFundDedAmt31" '(31)
  Fields.Add "fldFundDedAmt32" '(32)
  Fields.Add "fldFundDedAmt33" '(33)
  Fields.Add "fldFundDedAmt34" '(34)
  Fields.Add "fldFundDedAmt35" '(35)
  Fields.Add "fldFundDedAmt36" '(36)
  Fields.Add "fldFundDedAmt37" '(37)
  Fields.Add "fldFundDedAmt38" '(38)
  Fields.Add "fldFundDedAmt39" '(39)
  Fields.Add "fldFundDedAmt40" '(40)
  Fields.Add "fldFundDedAmt41" '(41)
  Fields.Add "fldFundDedAmt42" '(42)
  Fields.Add "fldFundDedAmt43" '(43)
  Fields.Add "fldFundDedAmt44" '(44)
  Fields.Add "fldFundDedAmt45" '(45)
  Fields.Add "fldFundDedAmt46" '(46)
  Fields.Add "fldFundDedAmt47" '(47)
  Fields.Add "fldFundDedAmt48" '(48)
  Fields.Add "fldFundDedAmt49" '(49)
  Fields.Add "fldFundDedAmt50" '(50)
  
  
  Fields.Add "fldFundFed" '(51)
  Fields.Add "fldFundSta" '(52)
  Fields.Add "fldFundMed" '(53)
  Fields.Add "fldFundSoc" '(54)
  Fields.Add "fldFundRet" '(55)

End Sub
Private Sub ActiveReport_ReportEnd()
  If FFile <> 0 Then
    Close #FFile
  End If
End Sub

Private Sub ActiveReport_Terminate()
  Close
End Sub
