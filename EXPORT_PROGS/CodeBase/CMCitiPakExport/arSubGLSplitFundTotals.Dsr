VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSubGLSplitFundTotals 
   Caption         =   "ActiveReport1"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10140
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   17886
   _ExtentY        =   9102
   SectionData     =   "arSubGLSplitFundTotals.dsx":0000
End
Attribute VB_Name = "arSubGLSplitFundTotals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private TFile As Integer
Private PrintIt As Integer
Private CenImp As Boolean
Private CenImpFund$

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
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
Private Sub ActiveReport_DataInitialize()
  TFile = FreeFile
  Open StartPath & "\PRRPTS\GLFundSplTot.RPT" For Input As #TFile
  
  Fields.Add "fldFundNum" '(0)
  Fields.Add "fldTFundDebit" '(1)
  Fields.Add "fldTFundCredit" '(2)
  Fields.Add "fldFedTax" '3)
  Fields.Add "fldMedTax" '4)
  Fields.Add "fldSocTax" '5)
  Fields.Add "fldStaTax" '6)
  Fields.Add "fldRetTax" '7)
  Fields.Add "fldMedMat" '8)
  Fields.Add "fldSocMat" '9)
  Fields.Add "fldRetMat" '10)
  Fields.Add "fldSocTot"
  Fields.Add "fldMedTot"
  Fields.Add "fldRetTot"
  Fields.Add "fldFSMTot"
  Fields.Add "DedDesc1" '11)
  Fields.Add "DedAmt1" '12)
  Fields.Add "DedDesc2" '13)
  Fields.Add "DedAmt2" '14)
  Fields.Add "DedDesc3" '15)
  Fields.Add "DedAmt3" '16)
  Fields.Add "DedDesc4" '17)
  Fields.Add "DedAmt4" '18)
  Fields.Add "DedDesc5" '19)
  Fields.Add "DedAmt5" '20)
  Fields.Add "DedDesc6" '21)
  Fields.Add "DedAmt6" '22)
  Fields.Add "DedDesc7" '23)
  Fields.Add "DedAmt7" '24)
  Fields.Add "DedDesc8" '25)
  Fields.Add "DedAmt8" '26)
  Fields.Add "DedDesc9" '27)
  Fields.Add "DedAmt9" '28)
  Fields.Add "DedDesc10" '29)
  Fields.Add "DedAmt10" '30)
  Fields.Add "DedDesc11" '31)
  Fields.Add "DedAmt11" '32)
  Fields.Add "DedDesc12" '33)
  Fields.Add "DedAmt12" '34)
  Fields.Add "DedDesc13" '35)
  Fields.Add "DedAmt13" '36)
  Fields.Add "DedDesc14" '37)
  Fields.Add "DedAmt14" '38)
  Fields.Add "DedDesc15" '39)
  Fields.Add "DedAmt15" '40)
  Fields.Add "DedDesc16" '41)
  Fields.Add "DedAmt16" '42)
  Fields.Add "DedDesc17" '43)
  Fields.Add "DedAmt17" '44)
  Fields.Add "DedDesc18" '45)
  Fields.Add "DedAmt18" '46)
  Fields.Add "DedDesc19" '47)
  Fields.Add "DedAmt19" '48)
  Fields.Add "DedDesc20" '49)
  Fields.Add "DedAmt20" '50)
  Fields.Add "DedDesc21" '51)
  Fields.Add "DedAmt21" '52)
  Fields.Add "DedDesc22" '53)
  Fields.Add "DedAmt22" '54)
  Fields.Add "DedDesc23" '55)
  Fields.Add "DedAmt23" '56)
  Fields.Add "DedDesc24" '57)
  Fields.Add "DedAmt24" '58)
  Fields.Add "DedDesc25" '59)
  Fields.Add "DedAmt25" '60)
  Fields.Add "DedDesc26" '61)
  Fields.Add "DedAmt26" '62)
  Fields.Add "DedDesc27" '63)
  Fields.Add "DedAmt27" '64)
  Fields.Add "DedDesc28" '65)
  Fields.Add "DedAmt28" '66)
  Fields.Add "DedDesc29" '67)
  Fields.Add "DedAmt29" '68)
  Fields.Add "DedDesc30" '69)
  Fields.Add "DedAmt30" '70)
  Fields.Add "DedDesc31" '71)
  Fields.Add "DedAmt31" '72)
  Fields.Add "DedDesc32" '73)
  Fields.Add "DedAmt32" '74)
  Fields.Add "DedDesc33" '75)
  Fields.Add "DedAmt33" '76)
  Fields.Add "DedDesc34" '77)
  Fields.Add "DedAmt34" '78)
  Fields.Add "DedDesc35" '79)
  Fields.Add "DedAmt35" '80)
  Fields.Add "DedDesc36" '81)
  Fields.Add "DedAmt36" '82)
  Fields.Add "DedDesc37" '83)
  Fields.Add "DedAmt37" '84)
  Fields.Add "DedDesc38" '85)
  Fields.Add "DedAmt38" '86)
  Fields.Add "DedDesc39" '87)
  Fields.Add "DedAmt39" '88)
  Fields.Add "DedDesc40" '89)
  Fields.Add "DedAmt40" '90)
  Fields.Add "DedDesc41" '91)
  Fields.Add "DedAmt41" '92)
  Fields.Add "DedDesc42" '93)
  Fields.Add "DedAmt42" '94)
  Fields.Add "DedDesc43" '95)
  Fields.Add "DedAmt43" '96)
  Fields.Add "DedDesc44" '97)
  Fields.Add "DedAmt44" '98)
  Fields.Add "DedDesc45" '99)
  Fields.Add "DedAmt45" '100)
  Fields.Add "DedDesc46" '101)
  Fields.Add "DedAmt46" '102)
  Fields.Add "DedDesc47" '103)
  Fields.Add "DedAmt47" '104)
  Fields.Add "DedDesc48" '105)
  Fields.Add "DedAmt48" '106)
  Fields.Add "DedDesc49" '107)
  Fields.Add "DedAmt49" '108)
  Fields.Add "DedDesc50" '109)
  Fields.Add "DedAmt50" '110)
  Fields.Add "DedCnt" '111)
  Fields.Add "TotFedTax" '112)
  Fields.Add "TotMedTax" '113)
  Fields.Add "TotSocTax" '114)
  Fields.Add "TotStaTax" '115)
  Fields.Add "TotRetTax" '116)
  Fields.Add "TotMedMat" '117)
  Fields.Add "TotSocMat" '118)
  Fields.Add "TotRetMat" '119)
  Fields.Add "TotSoc"
  Fields.Add "TotMed"
  Fields.Add "TotRet"
  Fields.Add "TotFSM"
  Fields.Add "TDedDesc1" '120)
  Fields.Add "TDedAmt1" '121)
  Fields.Add "TDedDesc2" '122)
  Fields.Add "TDedAmt2" '123)
  Fields.Add "TDedDesc3" '124)
  Fields.Add "TDedAmt3" '125)
  Fields.Add "TDedDesc4" '126)
  Fields.Add "TDedAmt4" '127)
  Fields.Add "TDedDesc5" '128)
  Fields.Add "TDedAmt5" '129)
  Fields.Add "TDedDesc6" '130)
  Fields.Add "TDedAmt6" '131)
  Fields.Add "TDedDesc7" '132)
  Fields.Add "TDedAmt7" '133)
  Fields.Add "TDedDesc8" '134)
  Fields.Add "TDedAmt8" '135)
  Fields.Add "TDedDesc9" '136)
  Fields.Add "TDedAmt9" '137)
  Fields.Add "TDedDesc10" '138)
  Fields.Add "TDedAmt10" '139)
  Fields.Add "TDedDesc11" '140)
  Fields.Add "TDedAmt11" '141)
  Fields.Add "TDedDesc12" '142)
  Fields.Add "TDedAmt12" '143)
  Fields.Add "TDedDesc13" '144)
  Fields.Add "TDedAmt13" '145)
  Fields.Add "TDedDesc14" '146)
  Fields.Add "TDedAmt14" '147)
  Fields.Add "TDedDesc15" '148)
  Fields.Add "TDedAmt15" '149)
  Fields.Add "TDedDesc16" '150)
  Fields.Add "TDedAmt16" '151)
  Fields.Add "TDedDesc17" '152)
  Fields.Add "TDedAmt17" '153)
  Fields.Add "TDedDesc18" '154)
  Fields.Add "TDedAmt18" '155)
  Fields.Add "TDedDesc19" '156)
  Fields.Add "TDedAmt19" '157)
  Fields.Add "TDedDesc20" '158)
  Fields.Add "TDedAmt20" '159)
  Fields.Add "TDedDesc21" '160)
  Fields.Add "TDedAmt21" '161)
  Fields.Add "TDedDesc22" '162)
  Fields.Add "TDedAmt22" '163)
  Fields.Add "TDedDesc23" '164)
  Fields.Add "TDedAmt23" '165)
  Fields.Add "TDedDesc24" '166)
  Fields.Add "TDedAmt24" '167)
  Fields.Add "TDedDesc25" '168)
  Fields.Add "TDedAmt25" '169)
  Fields.Add "TDedDesc26" '170)
  Fields.Add "TDedAmt26" '171)
  Fields.Add "TDedDesc27" '172)
  Fields.Add "TDedAmt27" '173)
  Fields.Add "TDedDesc28" '174)
  Fields.Add "TDedAmt28" '175)
  Fields.Add "TDedDesc29" '176)
  Fields.Add "TDedAmt29" '177)
  Fields.Add "TDedDesc30" '178)
  Fields.Add "TDedAmt30" '179)
  Fields.Add "TDedDesc31" '180)
  Fields.Add "TDedAmt31" '181)
  Fields.Add "TDedDesc32" '182)
  Fields.Add "TDedAmt32" '183)
  Fields.Add "TDedDesc33" '184)
  Fields.Add "TDedAmt33" '185)
  Fields.Add "TDedDesc34" '186)
  Fields.Add "TDedAmt34" '187)
  Fields.Add "TDedDesc35" '188)
  Fields.Add "TDedAmt35" '189)
  Fields.Add "TDedDesc36" '190)
  Fields.Add "TDedAmt36" '191)
  Fields.Add "TDedDesc37" '192)
  Fields.Add "TDedAmt37" '193)
  Fields.Add "TDedDesc38" '194)
  Fields.Add "TDedAmt38" '195)
  Fields.Add "TDedDesc39" '196)
  Fields.Add "TDedAmt39" '197)
  Fields.Add "TDedDesc40" '198)
  Fields.Add "TDedAmt40" '199)
  Fields.Add "TDedDesc41" '200)
  Fields.Add "TDedAmt41" '201)
  Fields.Add "TDedDesc42" '202)
  Fields.Add "TDedAmt42" '203)
  Fields.Add "TDedDesc43" '204)
  Fields.Add "TDedAmt43" '205)
  Fields.Add "TDedDesc44" '206)
  Fields.Add "TDedAmt44" '207)
  Fields.Add "TDedDesc45" '208)
  Fields.Add "TDedAmt45" '209)
  Fields.Add "TDedDesc46" '210)
  Fields.Add "TDedAmt46" '211)
  Fields.Add "TDedDesc47" '212)
  Fields.Add "TDedAmt47" '213)
  Fields.Add "TDedDesc48" '214)
  Fields.Add "TDedAmt48" '215)
  Fields.Add "TDedDesc49" '216)
  Fields.Add "TDedAmt49" '217)
  Fields.Add "TDedDesc50" '218)
  Fields.Add "TDedAmt50" '219)
  Fields.Add "TotDebit" '220)
  Fields.Add "TotCredit" '221)
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
  ' Here we set the values of the fields that we defines as unbound
  ' or user defined.
  Fields("fldFundNum").Value = arrT(0)
  Fields("fldTFundDebit").Value = arrT(1)
  Fields("fldTFundCredit").Value = arrT(2)
  Fields("fldFedTax").Value = arrT(3)
  Fields("fldMedTax").Value = arrT(4)
  Fields("fldSocTax").Value = arrT(5)
  Fields("fldStaTax").Value = arrT(6)
  Fields("fldRetTax").Value = arrT(7)
  Fields("fldMedMat").Value = arrT(8)
  Fields("fldSocMat").Value = arrT(9)
  Fields("fldRetMat").Value = arrT(10)
  Fields("fldSocTot").Value = CDbl(arrT(5)) + CDbl(arrT(9))
  Fields("fldMedTot").Value = CDbl(arrT(4)) + CDbl(arrT(8))
  Fields("fldRetTot").Value = CDbl(arrT(7)) + CDbl(arrT(10))
  Fields("fldFSMTot").Value = CDbl(arrT(3)) + CDbl(arrT(5)) + CDbl(arrT(9)) + CDbl(arrT(4)) + CDbl(arrT(8))
  Fields("DedDesc1").Value = arrT(11)
  Fields("DedAmt1").Value = arrT(12)
  If arrT(13) = "" Then
    Field17.Visible = False
    Field18.Visible = False
  Else
    Field17.Visible = True
    Field18.Visible = True
  End If
  Fields("DedDesc2").Value = arrT(13)
  Fields("DedAmt2").Value = arrT(14)
  Fields("DedDesc3").Value = arrT(15)
  Fields("DedAmt3").Value = arrT(16)
  If arrT(17) = "" Then
    Field21.Visible = False
    Field22.Visible = False
  Else
    Field21.Visible = True
    Field22.Visible = True
  End If
  Fields("DedDesc4").Value = arrT(17)
  Fields("DedAmt4").Value = arrT(18)
  Fields("DedDesc5").Value = arrT(19)
  Fields("DedAmt5").Value = arrT(20)
  If arrT(21) = "" Then
    Field25.Visible = False
    Field26.Visible = False
  Else
    Field25.Visible = True
    Field26.Visible = True
  End If
  Fields("DedDesc6").Value = arrT(21)
  Fields("DedAmt6").Value = arrT(22)
  Fields("DedDesc7").Value = arrT(23)
  Fields("DedAmt7").Value = arrT(24)
  If arrT(25) = "" Then
    Field29.Visible = False
    Field30.Visible = False
  Else
    Field29.Visible = True
    Field30.Visible = True
  End If
  Fields("DedDesc8").Value = arrT(25)
  Fields("DedAmt8").Value = arrT(26)
  Fields("DedDesc9").Value = arrT(27)
  Fields("DedAmt9").Value = arrT(28)
  If arrT(29) = "" Then
    Field33.Visible = False
    Field34.Visible = False
  Else
    Field33.Visible = True
    Field34.Visible = True
  End If
  Fields("DedDesc10").Value = arrT(29)
  Fields("DedAmt10").Value = arrT(30)
  Fields("DedDesc11").Value = arrT(31)
  Fields("DedAmt11").Value = arrT(32)
  If arrT(33) = "" Then
    Field37.Visible = False
    Field38.Visible = False
  Else
    Field37.Visible = True
    Field38.Visible = True
  End If
  Fields("DedDesc12").Value = arrT(33)
  Fields("DedAmt12").Value = arrT(34)
  Fields("DedDesc13").Value = arrT(35)
  Fields("DedAmt13").Value = arrT(36)
  If QPTrim$(arrT(37)) = "" Then
    Field41.Visible = False
    Field42.Visible = False
  Else
    Field41.Visible = True
    Field42.Visible = True
  End If
  Fields("DedDesc14").Value = arrT(37)
  Fields("DedAmt14").Value = arrT(38)
  Fields("DedDesc15").Value = arrT(39)
  Fields("DedAmt15").Value = arrT(40)
  If arrT(41) = "" Then
    Field45.Visible = False
    Field46.Visible = False
  Else
    Field45.Visible = True
    Field46.Visible = True
  End If
  Fields("DedDesc16").Value = arrT(41)
  Fields("DedAmt16").Value = arrT(42)
  Fields("DedDesc17").Value = arrT(43)
  Fields("DedAmt17").Value = arrT(44)
  If arrT(45) = "" Then
    Field49.Visible = False
    Field50.Visible = False
  Else
    Field49.Visible = True
    Field50.Visible = True
  End If
  Fields("DedDesc18").Value = arrT(45)
  Fields("DedAmt18").Value = arrT(46)
  Fields("DedDesc19").Value = arrT(47)
  Fields("DedAmt19").Value = arrT(48)
  If arrT(49) = "" Then
    Field53.Visible = False
    Field54.Visible = False
  Else
    Field53.Visible = True
    Field54.Visible = True
  End If
  Fields("DedDesc20").Value = arrT(49)
  Fields("DedAmt20").Value = arrT(50)
  Fields("DedDesc21").Value = arrT(51)
  Fields("DedAmt21").Value = arrT(52)
  If arrT(53) = "" Then
    Field57.Visible = False
    Field58.Visible = False
  Else
    Field57.Visible = True
    Field58.Visible = True
  End If
  Fields("DedDesc22").Value = arrT(53)
  Fields("DedAmt22").Value = arrT(54)
  Fields("DedDesc23").Value = arrT(55)
  Fields("DedAmt23").Value = arrT(56)
  If arrT(57) = "" Then
    Field61.Visible = False
    Field62.Visible = False
  Else
    Field61.Visible = True
    Field62.Visible = True
  End If
  Fields("DedDesc24").Value = arrT(57)
  Fields("DedAmt24").Value = arrT(58)
  Fields("DedDesc25").Value = arrT(59)
  Fields("DedAmt25").Value = arrT(60)
  If arrT(61) = "" Then
    Field65.Visible = False
    Field66.Visible = False
  Else
    Field65.Visible = True
    Field66.Visible = True
  End If
  Fields("DedDesc26").Value = arrT(61)
  Fields("DedAmt26").Value = arrT(62)
  Fields("DedDesc27").Value = arrT(63)
  Fields("DedAmt27").Value = arrT(64)
  If arrT(65) = "" Then
    Field69.Visible = False
    Field70.Visible = False
  Else
    Field69.Visible = True
    Field70.Visible = True
  End If
  Fields("DedDesc28").Value = arrT(65)
  Fields("DedAmt28").Value = arrT(66)
  Fields("DedDesc29").Value = arrT(67)
  Fields("DedAmt29").Value = arrT(68)
  If arrT(69) = "" Then
    Field73.Visible = False
    Field74.Visible = False
  Else
    Field73.Visible = True
    Field74.Visible = True
  End If
  Fields("DedDesc30").Value = arrT(69)
  Fields("DedAmt30").Value = arrT(70)
  Fields("DedDesc31").Value = arrT(71)
  Fields("DedAmt31").Value = arrT(72)
  If arrT(73) = "" Then
    Field77.Visible = False
    Field78.Visible = False
  Else
    Field77.Visible = True
    Field78.Visible = True
  End If
  Fields("DedDesc32").Value = arrT(73)
  Fields("DedAmt32").Value = arrT(74)
  Fields("DedDesc33").Value = arrT(75)
  Fields("DedAmt33").Value = arrT(76)
  If arrT(77) = "" Then
    Field81.Visible = False
    Field82.Visible = False
  Else
    Field81.Visible = True
    Field82.Visible = True
  End If
  Fields("DedDesc34").Value = arrT(77)
  Fields("DedAmt34").Value = arrT(78)
  Fields("DedDesc35").Value = arrT(79)
  Fields("DedAmt35").Value = arrT(80)
  If arrT(81) = "" Then
    Field85.Visible = False
    Field86.Visible = False
  Else
    Field85.Visible = True
    Field86.Visible = True
  End If
  Fields("DedDesc36").Value = arrT(81)
  Fields("DedAmt36").Value = arrT(82)
  Fields("DedDesc37").Value = arrT(83)
  Fields("DedAmt37").Value = arrT(84)
  If arrT(85) = "" Then
    Field89.Visible = False
    Field90.Visible = False
  Else
    Field89.Visible = True
    Field90.Visible = True
  End If
  Fields("DedDesc38").Value = arrT(85)
  Fields("DedAmt38").Value = arrT(86)
  Fields("DedDesc39").Value = arrT(87)
  Fields("DedAmt39").Value = arrT(88)
  If arrT(89) = "" Then
    Field93.Visible = False
    Field94.Visible = False
  Else
    Field93.Visible = True
    Field94.Visible = True
  End If
  Fields("DedDesc40").Value = arrT(89)
  Fields("DedAmt40").Value = arrT(90)
  Fields("DedDesc41").Value = arrT(91)
  Fields("DedAmt41").Value = arrT(92)
  If arrT(93) = "" Then
    Field97.Visible = False
    Field98.Visible = False
  Else
    Field97.Visible = True
    Field98.Visible = True
  End If
  Fields("DedDesc42").Value = arrT(93)
  Fields("DedAmt42").Value = arrT(94)
  Fields("DedDesc43").Value = arrT(95)
  Fields("DedAmt43").Value = arrT(96)
  If arrT(97) = "" Then
    Field101.Visible = False
    Field102.Visible = False
  Else
    Field101.Visible = True
    Field102.Visible = True
  End If
  Fields("DedDesc44").Value = arrT(97)
  Fields("DedAmt44").Value = arrT(98)
  Fields("DedDesc45").Value = arrT(99)
  Fields("DedAmt45").Value = arrT(100)
  If arrT(101) = "" Then
    Field105.Visible = False
    Field106.Visible = False
  Else
    Field105.Visible = True
    Field106.Visible = True
  End If
  Fields("DedDesc46").Value = arrT(101)
  Fields("DedAmt46").Value = arrT(102)
  Fields("DedDesc47").Value = arrT(103)
  Fields("DedAmt47").Value = arrT(104)
  If arrT(105) = "" Then
    Field109.Visible = False
    Field110.Visible = False
  Else
    Field109.Visible = True
    Field110.Visible = True
  End If
  Fields("DedDesc48").Value = arrT(105)
  Fields("DedAmt48").Value = arrT(106)
  Fields("DedDesc49").Value = arrT(107)
  Fields("DedAmt49").Value = arrT(108)
  If arrT(109) = "" Then
    Field113.Visible = False
    Field114.Visible = False
  Else
    Field113.Visible = True
    Field114.Visible = True
  End If
  Fields("DedDesc50").Value = arrT(109)
  Fields("DedAmt50").Value = arrT(110)
  Fields("DedCnt").Value = arrT(111)
  
  Fields("TotFedTax").Value = arrT(112)
  Fields("TotMedTax").Value = arrT(113)
  Fields("TotSocTax").Value = arrT(114)
  Fields("TotStaTax").Value = arrT(115)
  Fields("TotRetTax").Value = arrT(116)
  Fields("TotMedMat").Value = arrT(117)
  Fields("TotSocMat").Value = arrT(118)
  Fields("TotRetMat").Value = arrT(119)
  Fields("TotSoc").Value = CDbl(arrT(114)) + CDbl(arrT(118))
  Fields("TotMed").Value = CDbl(arrT(113)) + CDbl(arrT(117))
  Fields("TotRet").Value = CDbl(arrT(116)) + CDbl(arrT(119))
  Fields("TotFSM").Value = CDbl(arrT(112)) + CDbl(arrT(113)) + CDbl(arrT(114)) + CDbl(arrT(117)) + CDbl(arrT(118))
  
  Fields("TDedDesc1").Value = arrT(120)
  Fields("TDedAmt1").Value = arrT(121)
  If arrT(122) = "" Then
    Field134.Visible = False
    Field135.Visible = False
  Else
    Field134.Visible = True
    Field135.Visible = True
  End If
  Fields("TDedDesc2").Value = arrT(122)
  Fields("TDedAmt2").Value = arrT(123)
  Fields("TDedDesc3").Value = arrT(124)
  Fields("TDedAmt3").Value = arrT(125)
  If arrT(126) = "" Then
    Field138.Visible = False
    Field139.Visible = False
  Else
    Field138.Visible = True
    Field139.Visible = True
  End If
  Fields("TDedDesc4").Value = arrT(126)
  Fields("TDedAmt4").Value = arrT(127)
  Fields("TDedDesc5").Value = arrT(128)
  Fields("TDedAmt5").Value = arrT(129)
  If arrT(130) = "" Then
    Field142.Visible = False
    Field143.Visible = False
  Else
    Field142.Visible = True
    Field143.Visible = True
  End If
  Fields("TDedDesc6").Value = arrT(130)
  Fields("TDedAmt6").Value = arrT(131)
  Fields("TDedDesc7").Value = arrT(132)
  Fields("TDedAmt7").Value = arrT(133)
  If arrT(134) = "" Then
    Field146.Visible = False
    Field147.Visible = False
  Else
    Field146.Visible = True
    Field147.Visible = True
  End If
  Fields("TDedDesc8").Value = arrT(134)
  Fields("TDedAmt8").Value = arrT(135)
  Fields("TDedDesc9").Value = arrT(136)
  Fields("TDedAmt9").Value = arrT(137)
  If arrT(138) = "" Then
    Field150.Visible = False
    Field151.Visible = False
  Else
    Field150.Visible = True
    Field151.Visible = True
  End If
  Fields("TDedDesc10").Value = arrT(138)
  Fields("TDedAmt10").Value = arrT(139)
  Fields("TDedDesc11").Value = arrT(140)
  Fields("TDedAmt11").Value = arrT(141)
  If arrT(142) = "" Then
    Field154.Visible = False
    Field155.Visible = False
  Else
    Field154.Visible = True
    Field155.Visible = True
  End If
  Fields("TDedDesc12").Value = arrT(142)
  Fields("TDedAmt12").Value = arrT(143)
  Fields("TDedDesc13").Value = arrT(144)
  Fields("TDedAmt13").Value = arrT(145)
  If QPTrim$(arrT(146)) = "" Then
    Field158.Visible = False
    Field159.Visible = False
  Else
    Field158.Visible = True
    Field159.Visible = True
  End If
  Fields("TDedDesc14").Value = arrT(146)
  Fields("TDedAmt14").Value = arrT(147)
  Fields("TDedDesc15").Value = arrT(148)
  Fields("TDedAmt15").Value = arrT(149)
  If arrT(150) = "" Then
    Field162.Visible = False
    Field163.Visible = False
  Else
    Field162.Visible = True
    Field163.Visible = True
  End If
  Fields("TDedDesc16").Value = arrT(150)
  Fields("TDedAmt16").Value = arrT(151)
  Fields("TDedDesc17").Value = arrT(152)
  Fields("TDedAmt17").Value = arrT(153)
  If arrT(154) = "" Then
    Field166.Visible = False
    Field167.Visible = False
  Else
    Field166.Visible = True
    Field167.Visible = True
  End If
  Fields("TDedDesc18").Value = arrT(154)
  Fields("TDedAmt18").Value = arrT(155)
  Fields("TDedDesc19").Value = arrT(156)
  Fields("TDedAmt19").Value = arrT(157)
  If arrT(158) = "" Then
    Field170.Visible = False
    Field171.Visible = False
  Else
    Field170.Visible = True
    Field171.Visible = True
  End If
  Fields("TDedDesc20").Value = arrT(158)
  Fields("TDedAmt20").Value = arrT(159)
  Fields("TDedDesc21").Value = arrT(160)
  Fields("TDedAmt21").Value = arrT(161)
  If arrT(162) = "" Then
    Field174.Visible = False
    Field175.Visible = False
  Else
    Field174.Visible = True
    Field175.Visible = True
  End If
  Fields("TDedDesc22").Value = arrT(162)
  Fields("TDedAmt22").Value = arrT(163)
  Fields("TDedDesc23").Value = arrT(164)
  Fields("TDedAmt23").Value = arrT(165)
  If arrT(166) = "" Then
    Field178.Visible = False
    Field179.Visible = False
  Else
    Field178.Visible = True
    Field179.Visible = True
  End If
  Fields("TDedDesc24").Value = arrT(166)
  Fields("TDedAmt24").Value = arrT(167)
  Fields("TDedDesc25").Value = arrT(168)
  Fields("TDedAmt25").Value = arrT(169)
  If arrT(170) = "" Then
    Field182.Visible = False
    Field183.Visible = False
  Else
    Field182.Visible = True
    Field183.Visible = True
  End If
  Fields("TDedDesc26").Value = arrT(170)
  Fields("TDedAmt26").Value = arrT(171)
  Fields("TDedDesc27").Value = arrT(172)
  Fields("TDedAmt27").Value = arrT(173)
  If arrT(174) = "" Then
    Field186.Visible = False
    Field187.Visible = False
  Else
    Field186.Visible = True
    Field187.Visible = True
  End If
  Fields("TDedDesc28").Value = arrT(174)
  Fields("TDedAmt28").Value = arrT(175)
  Fields("TDedDesc29").Value = arrT(176)
  Fields("TDedAmt29").Value = arrT(177)
  If arrT(178) = "" Then
    Field190.Visible = False
    Field191.Visible = False
  Else
    Field190.Visible = True
    Field191.Visible = True
  End If
  Fields("TDedDesc30").Value = arrT(178)
  Fields("TDedAmt30").Value = arrT(179)
  Fields("TDedDesc31").Value = arrT(180)
  Fields("TDedAmt31").Value = arrT(181)
  If arrT(182) = "" Then
    Field194.Visible = False
    Field195.Visible = False
  Else
    Field194.Visible = True
    Field195.Visible = True
  End If
  Fields("TDedDesc32").Value = arrT(182)
  Fields("TDedAmt32").Value = arrT(183)
  Fields("TDedDesc33").Value = arrT(184)
  Fields("TDedAmt33").Value = arrT(185)
  If arrT(186) = "" Then
    Field198.Visible = False
    Field199.Visible = False
  Else
    Field198.Visible = True
    Field199.Visible = True
  End If
  Fields("TDedDesc34").Value = arrT(186)
  Fields("TDedAmt34").Value = arrT(187)
  Fields("TDedDesc35").Value = arrT(188)
  Fields("TDedAmt35").Value = arrT(189)
  If arrT(190) = "" Then
    Field202.Visible = False
    Field203.Visible = False
  Else
    Field202.Visible = True
    Field203.Visible = True
  End If
  Fields("TDedDesc36").Value = arrT(190)
  Fields("TDedAmt36").Value = arrT(191)
  Fields("TDedDesc37").Value = arrT(192)
  Fields("TDedAmt37").Value = arrT(193)
  If arrT(194) = "" Then
    Field206.Visible = False
    Field207.Visible = False
  Else
    Field206.Visible = True
    Field207.Visible = True
  End If
  Fields("TDedDesc38").Value = arrT(194)
  Fields("TDedAmt38").Value = arrT(195)
  Fields("TDedDesc39").Value = arrT(196)
  Fields("TDedAmt39").Value = arrT(197)
  If arrT(198) = "" Then
    Field210.Visible = False
    Field211.Visible = False
  Else
    Field210.Visible = True
    Field211.Visible = True
  End If
  Fields("TDedDesc40").Value = arrT(198)
  Fields("TDedAmt40").Value = arrT(199)
  Fields("TDedDesc41").Value = arrT(200)
  Fields("TDedAmt41").Value = arrT(201)
  If arrT(202) = "" Then
    Field214.Visible = False
    Field215.Visible = False
  Else
    Field214.Visible = True
    Field215.Visible = True
  End If
  Fields("TDedDesc42").Value = arrT(202)
  Fields("TDedAmt42").Value = arrT(203)
  Fields("TDedDesc43").Value = arrT(204)
  Fields("TDedAmt43").Value = arrT(205)
  If arrT(206) = "" Then
    Field218.Visible = False
    Field219.Visible = False
  Else
    Field218.Visible = True
    Field219.Visible = True
  End If
  Fields("TDedDesc44").Value = arrT(206)
  Fields("TDedAmt44").Value = arrT(207)
  Fields("TDedDesc45").Value = arrT(208)
  Fields("TDedAmt45").Value = arrT(209)
  If arrT(210) = "" Then
    Field222.Visible = False
    Field223.Visible = False
  Else
    Field222.Visible = True
    Field223.Visible = True
  End If
  Fields("TDedDesc46").Value = arrT(210)
  Fields("TDedAmt46").Value = arrT(211)
  Fields("TDedDesc47").Value = arrT(212)
  Fields("TDedAmt47").Value = arrT(213)
  If arrT(214) = "" Then
    Field226.Visible = False
    Field227.Visible = False
  Else
    Field226.Visible = True
    Field227.Visible = True
  End If
  Fields("TDedDesc48").Value = arrT(214)
  Fields("TDedAmt48").Value = arrT(215)
  Fields("TDedDesc49").Value = arrT(216)
  Fields("TDedAmt49").Value = arrT(217)
  If arrT(218) = "" Then
    Field230.Visible = False
    Field231.Visible = False
  Else
    Field230.Visible = True
    Field231.Visible = True
  End If
  Fields("TDedDesc50").Value = arrT(218)
  Fields("TDedAmt50").Value = arrT(219)
  Fields("TotDebit").Value = arrT(220)
  Fields("TotCredit").Value = arrT(221)
  PrintIt = PrintIt + 1
End Sub
Private Sub ActiveReport_ReportEnd()
  If TFile <> 0 Then
    Close #TFile
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Dim SysRec As RegDSysFileRecType
  Dim SHandle As Integer
  
  OpenSysFile SHandle
  Get SHandle, 1, SysRec
  Close SHandle
  CenImpFund = Mid(SysRec.ICRACCT, 1, 2)
  If QPTrim$(SysRec.USEIMP) = "C" Or QPTrim$(SysRec.USEIMP) = "I" Then
    CenImp = True
    Line2.Visible = False
  Else
    CenImp = False
  End If

End Sub

Private Sub ReportFooter_Format()
  If Fields("DedCnt").Value <= 2 Then
    ReportFooter.Height = 3900
  ElseIf Fields("DedCnt").Value <= 4 Then
    ReportFooter.Height = 4170
  ElseIf Fields("DedCnt").Value <= 6 Then
    ReportFooter.Height = 4425
  ElseIf Fields("DedCnt").Value <= 8 Then
    ReportFooter.Height = 4710
  ElseIf Fields("DedCnt").Value <= 10 Then
    ReportFooter.Height = 4980
  ElseIf Fields("DedCnt").Value <= 12 Then
    ReportFooter.Height = 5250
  ElseIf Fields("DedCnt").Value <= 14 Then
    ReportFooter.Height = 5505
  ElseIf Fields("DedCnt").Value <= 16 Then
    ReportFooter.Height = 5790
  ElseIf Fields("DedCnt").Value <= 18 Then
    ReportFooter.Height = 6060
  ElseIf Fields("DedCnt").Value <= 20 Then
    ReportFooter.Height = 6315
  ElseIf Fields("DedCnt").Value <= 22 Then
    ReportFooter.Height = 6660
  ElseIf Fields("DedCnt").Value <= 24 Then
    ReportFooter.Height = 6870
  ElseIf Fields("DedCnt").Value <= 26 Then
    ReportFooter.Height = 7170
  ElseIf Fields("DedCnt").Value <= 28 Then
    ReportFooter.Height = 7440
  ElseIf Fields("DedCnt").Value <= 30 Then
    ReportFooter.Height = 7680
  ElseIf Fields("DedCnt").Value <= 32 Then
    ReportFooter.Height = 7950
  ElseIf Fields("DedCnt").Value <= 34 Then
    ReportFooter.Height = 8220
  ElseIf Fields("DedCnt").Value <= 36 Then
    ReportFooter.Height = 8505
  ElseIf Fields("DedCnt").Value <= 38 Then
    ReportFooter.Height = 8775
  ElseIf Fields("DedCnt").Value <= 40 Then
    ReportFooter.Height = 9030
  ElseIf Fields("DedCnt").Value <= 42 Then
    ReportFooter.Height = 9300
  ElseIf Fields("DedCnt").Value <= 44 Then
    ReportFooter.Height = 9570
  ElseIf Fields("DedCnt").Value <= 46 Then
    ReportFooter.Height = 9840
  ElseIf Fields("DedCnt").Value <= 48 Then
    ReportFooter.Height = 10100
  End If
End Sub

Private Sub GroupFooter1_Format()
'  GroupFooter1.Height = 0
End Sub

Private Sub Detail_Format()
  If CenImp = True Then 'with central and imprest the first fund is the
  'credit and debit fund only
'    If PrintIt = 1 Then
   If QPTrim$(Fields("fldFundNum").Value) = CenImpFund Then
      Detail.Visible = False
      Me.Line2.Visible = True 'made invisible at reportstart
      Exit Sub
    Else
      Detail.Visible = True
'      GoTo OK2Print
    End If
  End If
  
  If Fields("DedCnt").Value <= 2 Then
    Detail.Height = 3135
  ElseIf Fields("DedCnt").Value <= 4 Then
    Detail.Height = 3405
  ElseIf Fields("DedCnt").Value <= 6 Then
    Detail.Height = 3675
  ElseIf Fields("DedCnt").Value <= 8 Then
    Detail.Height = 3945
  ElseIf Fields("DedCnt").Value <= 10 Then
    Detail.Height = 4215
  ElseIf Fields("DedCnt").Value <= 12 Then
    Detail.Height = 4484
  ElseIf Fields("DedCnt").Value <= 14 Then
    Detail.Height = 4740
  ElseIf Fields("DedCnt").Value <= 16 Then
    Detail.Height = 4995
  ElseIf Fields("DedCnt").Value <= 18 Then
    Detail.Height = 5295
  ElseIf Fields("DedCnt").Value <= 20 Then
    Detail.Height = 5565
  ElseIf Fields("DedCnt").Value <= 22 Then
    Detail.Height = 5820
  ElseIf Fields("DedCnt").Value <= 24 Then
    Detail.Height = 6090
  ElseIf Fields("DedCnt").Value <= 26 Then
    Detail.Height = 6345
  ElseIf Fields("DedCnt").Value <= 28 Then
    Detail.Height = 6615
  ElseIf Fields("DedCnt").Value <= 30 Then
    Detail.Height = 6900
  ElseIf Fields("DedCnt").Value <= 32 Then
    Detail.Height = 7185
  ElseIf Fields("DedCnt").Value <= 34 Then
    Detail.Height = 7440
  ElseIf Fields("DedCnt").Value <= 36 Then
    Detail.Height = 7725
  ElseIf Fields("DedCnt").Value <= 38 Then
    Detail.Height = 7995
  ElseIf Fields("DedCnt").Value <= 40 Then
    Detail.Height = 8250
  ElseIf Fields("DedCnt").Value <= 42 Then
    Detail.Height = 8520
  ElseIf Fields("DedCnt").Value <= 44 Then
    Detail.Height = 8790
  ElseIf Fields("DedCnt").Value <= 46 Then
    Detail.Height = 9075
  ElseIf Fields("DedCnt").Value <= 48 Then
    Detail.Height = 9330
  End If

End Sub

