VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arBLDLQ2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Delinquent Notice #2"
   ClientHeight    =   8868
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   ControlBox      =   0   'False
   Icon            =   "arBLDLQ2.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15637
   SectionData     =   "arBLDLQ2.dsx":08CA
End
Attribute VB_Name = "arBLDLQ2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsBLTextBoxOverrider
Private Temp_Class As Resize_Class
Private hFile As Integer

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  hFile = FreeFile
  Open StartPath & "\BLRPTS\DLQ2.RPT" For Input As #hFile
  Fields.Add ("fld0") '0)
  Fields.Add ("fld1") '1)
  Fields.Add ("fld2") '2)
  Fields.Add ("fld3") '3)
  Fields.Add ("fld4") '4)
  Fields.Add ("fld5") '5)
  Fields.Add ("fld6") '6)
  Fields.Add ("fld7") '7)
  Fields.Add ("fld8") '8)
  Fields.Add ("fld9") '9)
  Fields.Add ("fld10") '10)
  Fields.Add ("fld11") '11)
  Fields.Add ("fld12") '12)
  Fields.Add ("fld13") '13)
  Fields.Add ("fld14") '14)
  Fields.Add ("fld15") '15)
  Fields.Add ("fld16") '16)
  Fields.Add ("fld17") '17)
  Fields.Add ("fld18") '18)
  Fields.Add ("fld19") '19)
  Fields.Add ("fld20") '20)
  Fields.Add ("fld21") '21)
  Fields.Add ("fld22") '22)
  Fields.Add ("fld23") '23)
  Fields.Add ("fld24") '24)
  Fields.Add ("fld25") '25)
  Fields.Add ("fld26") '26)
  Fields.Add ("fld27") '27)
  Fields.Add ("fld28") '28)
  Fields.Add ("fld29") '29)
  Fields.Add ("fld30") '30)
  Fields.Add ("fld31") '31)
  Fields.Add ("fld32") '32)
  Fields.Add ("fld33") '33)
  Fields.Add ("fld34") '34)
  Fields.Add ("fld35") '35)
  Fields.Add ("fld36") '36)
  Fields.Add ("fld37") '37)
  Fields.Add ("fld38") '38)
  Fields.Add ("fld39") '39)
  Fields.Add ("fld40") '40)
  Fields.Add ("fld41") '41)
  Fields.Add ("fld42") '42)
  Fields.Add ("fld43") '43)
  Fields.Add ("fld44") '44)
  Fields.Add ("fld45") '45)
  Fields.Add ("fld46") '46)
  Fields.Add ("fld47") '47)
  Fields.Add ("fld48") '48)
  Fields.Add ("fld49") '49)
  Fields.Add ("fld50") '50)
  Fields.Add ("fld51") '51)
  Fields.Add ("fld52") '52)
  Fields.Add ("fld53") '53)
  Fields.Add ("fld54") '54)
  Fields.Add ("fld55") '55)
  Fields.Add ("fld56") '56)
  Fields.Add ("fld57") '57)
  Fields.Add ("fld58") '58)
  Fields.Add ("fld59") '59)
  Fields.Add ("fld60") '60)
  Fields.Add ("fld61") '61)
  Fields.Add ("fld62") '62)
  Fields.Add ("fld63") '63)
  Fields.Add ("fld64") '64)
  Fields.Add ("fld65") '65)
  Fields.Add ("fld66") '66)
  Fields.Add ("fld67") '67)
  Fields.Add ("fld68") '68)
  Fields.Add ("fld69") '69)
  Fields.Add ("fld70") '70)
  Fields.Add ("fld71") '71)
  Fields.Add ("fld72") '72)
  Fields.Add ("fld73") '73)
  Fields.Add ("fld74") '74)
  Fields.Add ("fld75") '75)
  Fields.Add ("fld76") '76)
  Fields.Add ("fld77") '77)
  Fields.Add ("fld78") '78)
  Fields.Add ("fld79") '79)
  Fields.Add ("fld80") '80)
  Fields.Add ("fld81") '81)
  Fields.Add ("fld82") '82)
  Fields.Add ("fld83") '83)
  Fields.Add ("fld84") '84)
  Fields.Add ("fld85") '85)
  Fields.Add ("fld86") '86)
  Fields.Add ("fld87") '87)
  Fields.Add ("fld88") '88)
  Fields.Add ("fld89") '89)
  Fields.Add ("fld90") '90)
  Fields.Add ("fld91") '91)
  Fields.Add ("fld92") '92)
  Fields.Add ("fld93") '93)
  Fields.Add ("fld94") '94)
  Fields.Add ("fld95") '95)
  Fields.Add ("fld96") '96)
  Fields.Add ("fld97") '97)
  Fields.Add ("fld98") '98)
  Fields.Add ("fld99") '99)
  Fields.Add ("fld100") '100)
  Fields.Add ("fld101") '101)
  Fields.Add ("fld102") '102)
  Fields.Add ("fld103") '103)
  Fields.Add ("fld104") '104)
  Fields.Add ("fld105") '105)
  Fields.Add ("fld106") '106)
  Fields.Add ("fld107") '107)
  Fields.Add ("fld108") '108)
  Fields.Add ("fld109") '109)
  Fields.Add ("fld110") '110)
  Fields.Add ("fld111") '111)
  Fields.Add ("fld112") '112)
  Fields.Add ("fld113") '113)
  Fields.Add ("fld114") '114)
  Fields.Add ("fld115") '115)
  Fields.Add ("fld116") '116)
  Fields.Add ("fld117") '117)
  Fields.Add ("fld118") '118)
  Fields.Add ("fld119") '119)
  Fields.Add ("fld120") '120)
  Fields.Add ("fld121") '121)
  Fields.Add ("fld122") '122)
  Fields.Add ("fld123") '123)
  Fields.Add ("fld124") '124)
  Fields.Add ("fld125") '125)
  Fields.Add ("fld126") '126)
  Fields.Add ("fld127") '127)
  Fields.Add ("fld128") '128)
  Fields.Add ("fld129") '129)
  Fields.Add ("fld130") '130)
  Fields.Add ("fld131") '131)
  Fields.Add ("fld132") '132)
  Fields.Add ("fld133") '133)
  Fields.Add ("fld134") '134)
  Fields.Add ("fld135") '135)
  Fields.Add ("fld136") '136)
  Fields.Add ("fld137") '137)
  Fields.Add ("fld138") '138)
  Fields.Add ("fld139") '139)
  Fields.Add ("fld140") '140)
  Fields.Add ("fld141") '141)
  Fields.Add ("fld142") '142)
  Fields.Add ("fld143") '143)
  Fields.Add ("fld144") '144)
  Fields.Add ("fld145") '145)
  Fields.Add ("fld146") '146)
  Fields.Add ("fld147") '147)
  Fields.Add ("fld148") '148)
  Fields.Add ("fld149") '149)
  Fields.Add ("fld150") '150)
  Fields.Add ("fld151") '151)
  Fields.Add ("fld152") '152)
  Fields.Add ("fld153") '153)
  Fields.Add ("fld154") '154)
  Fields.Add ("fld155") '155)
  Fields.Add ("fld156") '156)
  Fields.Add ("fld157") '157)
  Fields.Add ("fld158") '158)
  Fields.Add ("fld159") '159)
  Fields.Add ("fld160") '160)
  Fields.Add ("fld161") '161)
  Fields.Add ("fld162") '162)
  Fields.Add ("fld163") '163)
  Fields.Add ("fld164") '164)
  Fields.Add ("fld165") '165)
  Fields.Add ("fld166") '166)
  Fields.Add ("fld167") '167)
  Fields.Add ("fld168") '168)
  Fields.Add ("fld169") '169)
  Fields.Add ("fld170") '170)
  Fields.Add ("fld171") '171)
  Fields.Add ("fld172") '172)
  Fields.Add ("fld173") '173)
  Fields.Add ("fld174") '174)
  Fields.Add ("fld175") '175)
  Fields.Add ("fld176") '176)
  Fields.Add ("fld177") '177)
  Fields.Add ("fld178") '178)
  Fields.Add ("fld179") '179)
  Fields.Add ("fld180") '180)
  Fields.Add ("fld181") '181)
  Fields.Add ("fld182") '182)
  Fields.Add ("fld183") '183)
  Fields.Add ("fld184") '184)
  Fields.Add ("fld185") '185)
  Fields.Add ("fld186") '186)
  Fields.Add ("fld187") '187)
  Fields.Add ("fld188") '188)
  Fields.Add ("fld189") '189)
  Fields.Add ("fld190") '190)
  Fields.Add ("fld191") '191)
  Fields.Add ("fld192") '192)
  Fields.Add ("fld193") '193)
  Fields.Add ("fld194") '194)
  Fields.Add ("fld195") '195)
  Fields.Add ("fld196") '196)
  Fields.Add ("fld197") '197)
  Fields.Add ("fld198") '198)
  Fields.Add ("fld199") '199)
  Fields.Add ("fld200") '200)
  Fields.Add ("fld201") '201)
  Fields.Add ("fld202") '202)
  Fields.Add ("fld203") '203)
  Fields.Add ("fld204") '204)
  Fields.Add ("fld205") '205)
  Fields.Add ("fld206") '206)
  Fields.Add ("fld207") '207)
  Fields.Add ("fld208") '208)
  Fields.Add ("fld209") '209)
  Fields.Add ("fld210") '210)
  Fields.Add ("fld211") '211)
  Fields.Add ("fld212") '212)
  Fields.Add ("fld213") '213)
  Fields.Add ("fld214") '214)
  Fields.Add ("fld215") '215)
  Fields.Add ("fld216") '216)
  Fields.Add ("fld217") '217)
  Fields.Add ("fld218") '218)
  Fields.Add ("fld219") '219)
  Fields.Add ("fld220") '220)
  Fields.Add ("fld221") '221)
  Fields.Add ("fld222") '222)
  Fields.Add ("fld223") '223)
  Fields.Add ("fld224") '224)
  Fields.Add ("fld225") '225)
  Fields.Add ("fld226") '226)
  Fields.Add ("fld227") '227)
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    Unload frmBLLoadReport
    frmBLMessageBoxJr.Label1.Caption = "Error Number: " & Str(Number) & " " & Description
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Unload Me
  End If
  CancelDisplay = True 'removes the error message

End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)
  Dim sLine As String
  Dim arr() As String
  Dim x As Integer
  Dim ctrl As Control
  Dim sec As Section
  Dim y As Integer

  Set sec = arBLDLQ2.Sections("Detail")

  For y = 0 To sec.Controls.Count - 1
    sec.Controls(y).Visible = True
  Next y
  
  If VBA.eof(hFile) Then
    eof = True
    Exit Sub
  Else
    eof = False
  End If
  Line Input #hFile, sLine
  arr = Split(sLine, "~")
  ' Here we set the values of the fields that we defines as unbound
  ' or user defined.
  If Len(QPTrim(arr(19))) > 0 Then
    Line7.Visible = False
    Line8.Visible = False
    Label1.Visible = False
    fldFeeM1.Visible = False
    fldBasM1.Visible = False
  End If

  If Len(QPTrim$(arr(47))) > 0 Then
    Line3.Visible = False
    Line4.Visible = False
    Line7.Visible = False
    Line8.Visible = False
    Label1.Visible = False
    fldFeeM1.Visible = False
    fldFeeS1.Visible = False
    fldBasM1.Visible = False
    fldBasS1.Visible = False
  End If

  If Len(QPTrim$(arr(49))) > 0 Then
    Line3.Visible = False
    Line4.Visible = False
    fldFeeS1.Visible = False
    fldBasS1.Visible = False
  End If

  If Len(QPTrim(arr(56))) = 0 And Len(QPTrim$(arr(84))) = 0 And Len(QPTrim$(arr(86))) = 0 Then
    Line10.Visible = False
    Line11.Visible = False
    Line12.Visible = False
    Line14.Visible = False
    Line22.Visible = False
    Label2.Visible = False
    fldFeeM2.Visible = False
    fldFeeS2.Visible = False
    fldBasM2.Visible = False
    fldBasS2.Visible = False
  End If

  If Len(QPTrim(arr(56))) > 0 Then 'step2
    Line14.Visible = False
    Line22.Visible = False
    Label2.Visible = False
    fldFeeM2.Visible = False
    fldBasM2.Visible = False
  End If

  If Len(QPTrim$(arr(84))) > 0 Then 'flat2
    Line10.Visible = False
    Line11.Visible = False
    Line14.Visible = False
    Line22.Visible = False
    Label2.Visible = False
    fldFeeM2.Visible = False
    fldFeeS2.Visible = False
    fldBasM2.Visible = False
    fldBasS2.Visible = False
  End If

  If Len(QPTrim$(arr(86))) > 0 Then 'multi2
    Line10.Visible = False
    Line11.Visible = False
    fldFeeS2.Visible = False
    fldBasS2.Visible = False
  End If

  If Len(QPTrim(arr(93))) = 0 And Len(QPTrim(arr(121))) = 0 And Len(QPTrim(arr(123))) = 0 Then
    Line16.Visible = False
    Line17.Visible = False
    Line18.Visible = False
    Line20.Visible = False
    Line21.Visible = False
    Label3.Visible = False
    fldFeeM3.Visible = False
    fldFeeS3.Visible = False
    fldBasM3.Visible = False
    fldBasS3.Visible = False
  End If

  If Len(QPTrim(arr(93))) > 0 Then 'step3
    Line20.Visible = False
    Line21.Visible = False
    Label3.Visible = False
    fldFeeM3.Visible = False
    fldBasM3.Visible = False
  End If

  If Len(QPTrim(arr(121))) > 0 Then 'flat3
    Line17.Visible = False
    Line18.Visible = False
    Line20.Visible = False
    Line21.Visible = False
    Label3.Visible = False
    fldFeeM3.Visible = False
    fldFeeS3.Visible = False
    fldBasM3.Visible = False
    fldBasS3.Visible = False
  End If

  If Len(QPTrim(arr(123))) > 0 Then 'multi3
    Line17.Visible = False
    Line18.Visible = False
    fldFeeS3.Visible = False
    fldBasS3.Visible = False
  End If

  If Len(QPTrim(arr(130))) = 0 And Len(QPTrim(arr(158))) = 0 And Len(QPTrim(arr(160))) = 0 Then
    Line24.Visible = False
    Line25.Visible = False
    Line26.Visible = False
    Line28.Visible = False
    Line29.Visible = False
    Label4.Visible = False
    fldFeeM4.Visible = False
    fldFeeS4.Visible = False
    fldBasM4.Visible = False
    fldBasS4.Visible = False
  End If

  If Len(QPTrim(arr(130))) > 0 Then 'step4
    Line28.Visible = False
    Line29.Visible = False
    Label4.Visible = False
    fldFeeM4.Visible = False
    fldBasM4.Visible = False
  End If

  If Len(QPTrim(arr(158))) > 0 Then 'flat4
    Line25.Visible = False
    Line26.Visible = False
    Line28.Visible = False
    Line29.Visible = False
    Label4.Visible = False
    fldFeeM4.Visible = False
    fldFeeS4.Visible = False
    fldBasM4.Visible = False
    fldBasS4.Visible = False
  End If

  If Len(QPTrim(arr(160))) > 0 Then 'multi4
    Line25.Visible = False
    Line26.Visible = False
    fldFeeS4.Visible = False
    fldBasS4.Visible = False
  End If

  If Len(QPTrim(arr(167))) = 0 And Len(QPTrim(arr(195))) = 0 And Len(QPTrim(arr(197))) = 0 Then
    Line32.Visible = False
    Line33.Visible = False
    Line34.Visible = False
    Line36.Visible = False
    Line37.Visible = False
    Label5.Visible = False
    fldFeeM5.Visible = False
    fldFeeS5.Visible = False
    fldBasM5.Visible = False
    fldBasS5.Visible = False
  End If

  If Len(QPTrim(arr(167))) > 0 Then 'step5
    Line36.Visible = False
    Line37.Visible = False
    Label5.Visible = False
    fldFeeM5.Visible = False
    fldBasM5.Visible = False
  End If

  If Len(QPTrim(arr(195))) > 0 Then 'flat5
    Line33.Visible = False
    Line34.Visible = False
    Line36.Visible = False
    Line37.Visible = False
    Label5.Visible = False
    fldFeeM5.Visible = False
    fldFeeS5.Visible = False
    fldBasM5.Visible = False
    fldBasS5.Visible = False
  End If

  If Len(QPTrim(arr(197))) > 0 Then 'multi5
    Line33.Visible = False
    Line34.Visible = False
    fldFeeS5.Visible = False
    fldBasS5.Visible = False
  End If
  
  If arr(227) = "" Then
    Line41.Visible = False
  End If
  
  Fields("fld0").Value = arr(0)
  Fields("fld1").Value = arr(1)
  Fields("fld2").Value = arr(2)
  Fields("fld3").Value = arr(3)
  Fields("fld4").Value = arr(4)
  Fields("fld5").Value = arr(5)
  Fields("fld6").Value = arr(6)
  Fields("fld7").Value = arr(7)
  Fields("fld8").Value = arr(8)
  Fields("fld9").Value = arr(9)
  Fields("fld10").Value = arr(10)
  Fields("fld11").Value = arr(11)
  Fields("fld12").Value = arr(12)
  Fields("fld13").Value = arr(13)
  Fields("fld14").Value = arr(14)
  Fields("fld15").Value = arr(15)
  Fields("fld16").Value = arr(16)
  Fields("fld17").Value = arr(17)
  Fields("fld18").Value = arr(18)
  Fields("fld19").Value = arr(19)
  Fields("fld20").Value = arr(20)
  Fields("fld21").Value = arr(21)
  Fields("fld22").Value = arr(22)
  Fields("fld23").Value = arr(23)
  Fields("fld24").Value = arr(24)
  Fields("fld25").Value = arr(25)
  Fields("fld26").Value = arr(26)
  Fields("fld27").Value = arr(27)
  Fields("fld28").Value = arr(28)
  Fields("fld29").Value = arr(29)
  Fields("fld30").Value = arr(30)
  Fields("fld31").Value = arr(31)
  Fields("fld32").Value = arr(32)
  Fields("fld33").Value = arr(33)
  Fields("fld34").Value = arr(34)
  Fields("fld35").Value = arr(35)
  Fields("fld36").Value = arr(36)
  Fields("fld37").Value = arr(37)
  Fields("fld38").Value = arr(38)
  Fields("fld39").Value = arr(39)
  Fields("fld40").Value = arr(40)
  Fields("fld41").Value = arr(41)
  Fields("fld42").Value = arr(42)
  Fields("fld43").Value = arr(43)
  Fields("fld44").Value = arr(44)
  Fields("fld45").Value = arr(45)
  Fields("fld46").Value = arr(46)
  Fields("fld47").Value = arr(47)
  Fields("fld48").Value = arr(48)
  Fields("fld49").Value = arr(49)
  Fields("fld50").Value = arr(50)
  Fields("fld51").Value = arr(51)
  Fields("fld52").Value = arr(52)
  Fields("fld53").Value = arr(53)
  Fields("fld54").Value = arr(54)
  Fields("fld55").Value = arr(55)
  Fields("fld56").Value = arr(56)
  Fields("fld57").Value = arr(57)
  Fields("fld58").Value = arr(58)
  Fields("fld59").Value = arr(59)
  Fields("fld60").Value = arr(60)
  Fields("fld61").Value = arr(61)
  Fields("fld62").Value = arr(62)
  Fields("fld63").Value = arr(63)
  Fields("fld64").Value = arr(64)
  Fields("fld65").Value = arr(65)
  Fields("fld66").Value = arr(66)
  Fields("fld67").Value = arr(67)
  Fields("fld68").Value = arr(68)
  Fields("fld69").Value = arr(69)
  Fields("fld70").Value = arr(70)
  Fields("fld71").Value = arr(71)
  Fields("fld72").Value = arr(72)
  Fields("fld73").Value = arr(73)
  Fields("fld74").Value = arr(74)
  Fields("fld75").Value = arr(75)
  Fields("fld76").Value = arr(76)
  Fields("fld77").Value = arr(77)
  Fields("fld78").Value = arr(78)
  Fields("fld79").Value = arr(79)
  Fields("fld80").Value = arr(80)
  Fields("fld81").Value = arr(81)
  Fields("fld82").Value = arr(82)
  Fields("fld83").Value = arr(83)
  Fields("fld84").Value = arr(84)
  Fields("fld85").Value = arr(85)
  Fields("fld86").Value = arr(86)
  Fields("fld87").Value = arr(87)
  Fields("fld88").Value = arr(88)
  Fields("fld89").Value = arr(89)
  Fields("fld90").Value = arr(90)
  Fields("fld91").Value = arr(91)
  Fields("fld92").Value = arr(92)
  Fields("fld93").Value = arr(93)
  Fields("fld94").Value = arr(94)
  Fields("fld95").Value = arr(95)
  Fields("fld96").Value = arr(96)
  Fields("fld97").Value = arr(97)
  Fields("fld98").Value = arr(98)
  Fields("fld99").Value = arr(99)
  Fields("fld100").Value = arr(100)
  Fields("fld101").Value = arr(101)
  Fields("fld102").Value = arr(102)
  Fields("fld103").Value = arr(103)
  Fields("fld104").Value = arr(104)
  Fields("fld105").Value = arr(105)
  Fields("fld106").Value = arr(106)
  Fields("fld107").Value = arr(107)
  Fields("fld108").Value = arr(108)
  Fields("fld109").Value = arr(109)
  Fields("fld110").Value = arr(110)
  Fields("fld111").Value = arr(111)
  Fields("fld112").Value = arr(112)
  Fields("fld113").Value = arr(113)
  Fields("fld114").Value = arr(114)
  Fields("fld115").Value = arr(115)
  Fields("fld116").Value = arr(116)
  Fields("fld117").Value = arr(117)
  Fields("fld118").Value = arr(118)
  Fields("fld119").Value = arr(119)
  Fields("fld120").Value = arr(120)
  Fields("fld121").Value = arr(121)
  Fields("fld122").Value = arr(122)
  Fields("fld123").Value = arr(123)
  Fields("fld124").Value = arr(124)
  Fields("fld125").Value = arr(125)
  Fields("fld126").Value = arr(126)
  Fields("fld127").Value = arr(127)
  Fields("fld128").Value = arr(128)
  Fields("fld129").Value = arr(129)
  Fields("fld130").Value = arr(130)
  Fields("fld131").Value = arr(131)
  Fields("fld132").Value = arr(132)
  Fields("fld133").Value = arr(133)
  Fields("fld134").Value = arr(134)
  Fields("fld135").Value = arr(135)
  Fields("fld136").Value = arr(136)
  Fields("fld137").Value = arr(137)
  Fields("fld138").Value = arr(138)
  Fields("fld139").Value = arr(139)
  Fields("fld140").Value = arr(140)
  Fields("fld141").Value = arr(141)
  Fields("fld142").Value = arr(142)
  Fields("fld143").Value = arr(143)
  Fields("fld144").Value = arr(144)
  Fields("fld145").Value = arr(145)
  Fields("fld146").Value = arr(146)
  Fields("fld147").Value = arr(147)
  Fields("fld148").Value = arr(148)
  Fields("fld149").Value = arr(149)
  Fields("fld150").Value = arr(150)
  Fields("fld151").Value = arr(151)
  Fields("fld152").Value = arr(152)
  Fields("fld153").Value = arr(153)
  Fields("fld154").Value = arr(154)
  Fields("fld155").Value = arr(155)
  Fields("fld156").Value = arr(156)
  Fields("fld157").Value = arr(157)
  Fields("fld158").Value = arr(158)
  Fields("fld159").Value = arr(159)
  Fields("fld160").Value = arr(160)
  Fields("fld161").Value = arr(161)
  Fields("fld162").Value = arr(162)
  Fields("fld163").Value = arr(163)
  Fields("fld164").Value = arr(164)
  Fields("fld165").Value = arr(165)
  Fields("fld166").Value = arr(166)
  Fields("fld167").Value = arr(167)
  Fields("fld168").Value = arr(168)
  Fields("fld169").Value = arr(169)
  Fields("fld170").Value = arr(170)
  Fields("fld171").Value = arr(171)
  Fields("fld172").Value = arr(172)
  Fields("fld173").Value = arr(173)
  Fields("fld174").Value = arr(174)
  Fields("fld175").Value = arr(175)
  Fields("fld176").Value = arr(176)
  Fields("fld177").Value = arr(177)
  Fields("fld178").Value = arr(178)
  Fields("fld179").Value = arr(179)
  Fields("fld180").Value = arr(180)
  Fields("fld181").Value = arr(181)
  Fields("fld182").Value = arr(182)
  Fields("fld183").Value = arr(183)
  Fields("fld184").Value = arr(184)
  Fields("fld185").Value = arr(185)
  Fields("fld186").Value = arr(186)
  Fields("fld187").Value = arr(187)
  Fields("fld188").Value = arr(188)
  Fields("fld189").Value = arr(189)
  Fields("fld190").Value = arr(190)
  Fields("fld191").Value = arr(191)
  Fields("fld192").Value = arr(192)
  Fields("fld193").Value = arr(193)
  Fields("fld194").Value = arr(194)
  Fields("fld195").Value = arr(195)
  Fields("fld196").Value = arr(196)
  Fields("fld197").Value = arr(197)
  Fields("fld198").Value = arr(198)
  Fields("fld199").Value = arr(199)
  Fields("fld200").Value = arr(200)
  Fields("fld201").Value = arr(201)
  Fields("fld202").Value = arr(202)
  Fields("fld203").Value = arr(203)
  Fields("fld204").Value = arr(204)
  Fields("fld205").Value = arr(205)
  Fields("fld206").Value = arr(206)
  Fields("fld207").Value = arr(207)
  Fields("fld208").Value = arr(208)
  Fields("fld209").Value = arr(209)
  Fields("fld210").Value = arr(210)
  Fields("fld211").Value = arr(211)
  Fields("fld212").Value = arr(212)
  Fields("fld213").Value = arr(213)
  Fields("fld214").Value = arr(214)
  Fields("fld215").Value = arr(215)
  Fields("fld216").Value = arr(216)
  Fields("fld217").Value = arr(217)
  Fields("fld218").Value = arr(218)
  Fields("fld219").Value = arr(219)
  Fields("fld220").Value = arr(220)
  Fields("fld221").Value = arr(221)
  Fields("fld222").Value = arr(222)
  Fields("fld223").Value = arr(223)
  Fields("fld224").Value = arr(224)
  Fields("fld225").Value = arr(225)
  Fields("fld226").Value = arr(226)
  Fields("fld227").Value = arr(227)
End Sub

Private Sub ActiveReport_Initialize()
  Me.ToolBar.Tools.Add "&Close"
  Me.ToolBar.Tools.Add "Save/&Excel"
  Me.ToolBar.Tools.Add "&Text"
End Sub

Private Sub ActiveReport_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
  If KeyCode = vbKeyEscape Then
    Unload Me
    KeyCode = 0
  End If
  If Shift = 4 Then
    If KeyCode = vbKeyC Then
      Unload Me
      KeyCode = 0
    ElseIf KeyCode = vbKeyE Then
      Screen.MousePointer = vbHourglass
      ExportReport 1
      Screen.MousePointer = vbDefault
      DoEvents
      frmBLMessageBoxJr.Label1.Caption = "File - BLApp8.xls, created in the Citipak Directory."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmBLMessageBoxJr.Label1.Caption = "File - BLApp8.txt, created in the Citipak Directory."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      KeyCode = 0
    End If
  End If
End Sub

Private Sub ActiveReport_Terminate()
  Close
End Sub

Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
  If Tool = "&Close" Then
    Close
    Unload Me
  End If
  If Tool = "Save/&Excel" Then
    Screen.MousePointer = vbHourglass
    ExportReport 1
    Screen.MousePointer = vbDefault
    DoEvents
    frmBLMessageBoxJr.Label1.Caption = "File - BLApp8.xls, created in the Citipak Directory."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmBLMessageBoxJr.Label1.Caption = "File - BLApp8.txt, created in the Citipak Directory."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  End If
End Sub
Private Sub ExportReport(x As Integer)
  Dim oEXL As ActiveReportsExcelExport.ARExportExcel
  Dim oTXT As ActiveReportsTextExport.ARExportText
  Dim outfile As String
  If Right$(StartPath, 1) = ":" Then
    outfile = StartPath
  Else
    outfile = StartPath & "\"
  End If
  Select Case x
    Case 1   '"Excel"
        Set oEXL = New ActiveReportsExcelExport.ARExportExcel
        oEXL.FileName = outfile & "BLApp8.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "BLApp8.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_ReportEnd()
  Unload frmBLLoadReport
  If hFile <> 0 Then
    Close #hFile
  End If
End Sub

Private Sub Form_Load()
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Unload frmBLLoadReport
  Me.Zoom = -1
End Sub
