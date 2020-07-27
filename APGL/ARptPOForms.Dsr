VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARptPOForms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order Forms"
   ClientHeight    =   7170
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   8970
   Icon            =   "ARptPOForms.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   15822
   _ExtentY        =   12647
   SectionData     =   "ARptPOForms.dsx":08CA
End
Attribute VB_Name = "ARptPOForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportFile As String
Private hFile As Integer
Dim rpt As ActiveReport
'Dim headers(1 To 239) As String
Dim cnt As Integer

Public Sub GetName(RName As String)
  ReportFile$ = RName$
End Sub

Private Sub ActiveReport_DataInitialize()
Dim cntB As Integer
   hFile = FreeFile
   Open ReportFile$ For Input As #hFile

'    headers(1) = "Header1"
'    headers(2) = "Header2"
'    headers(3) = "Header3"
'    headers(4) = "Header4"
'    headers(5) = "PODate"
'    headers(6) = "PONum"
'    headers(7) = "Vend1"
'    headers(8) = "Ship1"
'    headers(9) = "Vend2"
'    headers(10) = "Ship2"
'    headers(11) = "Vend3"
'    headers(12) = "Ship3"
'    headers(13) = "Vend4"
'    headers(14) = "Ship4"
'    headers(15) = "ShipVia"
'    headers(16) = "FOB"
'    headers(17) = "Dept"
'    headers(18) = "ShipOn"
'    headers(19) = "Terms"
'    For cntB = 1 To 216
'      headers(19 + cntB) = ("Body" & cntB)
'    Next
    
'    headers(236) = "POAmt"
'    headers(237) = "Add1"
'    headers(238) = "Add2"
'    headers(239) = "Add3"

'
'    ' This sets up the fields used in data binding
'    'Fields.Add "ProductID"
'    For cnt = 1 To 239
'      Fields.Add headers(cnt)
'    Next
  Fields.Add "Header1"
  Fields.Add "Header2"
  Fields.Add "Header3"
  Fields.Add "Header4"
  Fields.Add "PODate"
  Fields.Add "PONum"
  Fields.Add "Vend1"
  Fields.Add "Ship1"
  Fields.Add "Vend2"
  Fields.Add "Ship2"
  Fields.Add "Vend3"
  Fields.Add "Ship3"
  Fields.Add "Vend4"
  Fields.Add "Ship4"
  Fields.Add "Ship5"
  Fields.Add "ShipVia"
  Fields.Add "FOB"
  Fields.Add "Dept"
  Fields.Add "ShipOn"
  Fields.Add "Terms"
  Fields.Add "Body1"
  Fields.Add "Body2"
  Fields.Add "Body3"
  Fields.Add "Body4"
  Fields.Add "Body5"
  Fields.Add "Body6"
  Fields.Add "Body7"
  Fields.Add "Body8"
  Fields.Add "Body9"
  Fields.Add "Body10"
  Fields.Add "Body11"
  Fields.Add "Body12"
  Fields.Add "Body13"
  Fields.Add "Body14"
  Fields.Add "Body15"
  Fields.Add "Body16"
  Fields.Add "Body17"
  Fields.Add "Body18"
  Fields.Add "Body19"
  Fields.Add "Body20"
  Fields.Add "Body21"
  Fields.Add "Body22"
  Fields.Add "Body23"
  Fields.Add "Body24"
  Fields.Add "Body25"
  Fields.Add "Body26"
  Fields.Add "Body27"
  Fields.Add "Body28"
  Fields.Add "Body29"
  Fields.Add "Body30"
  Fields.Add "Body31"
  Fields.Add "Body32"
  Fields.Add "Body33"
  Fields.Add "Body34"
  Fields.Add "Body35"
  Fields.Add "Body36"
  Fields.Add "Body37"
  Fields.Add "Body38"
  Fields.Add "Body39"
  Fields.Add "Body40"
  Fields.Add "Body41"
  Fields.Add "Body42"
  Fields.Add "Body43"
  Fields.Add "Body44"
  Fields.Add "Body45"
  Fields.Add "Body46"
  Fields.Add "Body47"
  Fields.Add "Body48"
  Fields.Add "Body49"
  Fields.Add "Body50"
  Fields.Add "Body51"
  Fields.Add "Body52"
  Fields.Add "Body53"
  Fields.Add "Body54"
  Fields.Add "Body55"
  Fields.Add "Body56"
  Fields.Add "Body57"
  Fields.Add "Body58"
  Fields.Add "Body59"
  Fields.Add "Body60"
  Fields.Add "Body61"
  Fields.Add "Body62"
  Fields.Add "Body63"
  Fields.Add "Body64"
  Fields.Add "Body65"
  Fields.Add "Body66"
  Fields.Add "Body67"
  Fields.Add "Body68"
  Fields.Add "Body69"
  Fields.Add "Body70"
  Fields.Add "Body71"
  Fields.Add "Body72"
  Fields.Add "Body73"
  Fields.Add "Body74"
  Fields.Add "Body75"
  Fields.Add "Body76"
  Fields.Add "Body77"
  Fields.Add "Body78"
  Fields.Add "Body79"
  Fields.Add "Body80"
  Fields.Add "Body81"
  Fields.Add "Body82"
  Fields.Add "Body83"
  Fields.Add "Body84"
  Fields.Add "Body85"
  Fields.Add "Body86"
  Fields.Add "Body87"
  Fields.Add "Body88"
  Fields.Add "Body89"
  Fields.Add "Body90"
  Fields.Add "Body91"
  Fields.Add "Body92"
  Fields.Add "Body93"
  Fields.Add "Body94"
  Fields.Add "Body95"
  Fields.Add "Body96"
  Fields.Add "Body97"
  Fields.Add "Body98"
  Fields.Add "Body99"
  Fields.Add "Body100"
  Fields.Add "Body101"
  Fields.Add "Body102"
  Fields.Add "Body103"
  Fields.Add "Body104"
  Fields.Add "Body105"
  Fields.Add "Body106"
  Fields.Add "Body107"
  Fields.Add "Body108"
  Fields.Add "Body109"
  Fields.Add "Body110"
  Fields.Add "Body111"
  Fields.Add "Body112"
  Fields.Add "Body113"
  Fields.Add "Body114"
  Fields.Add "Body115"
  Fields.Add "Body116"
  Fields.Add "Body117"
  Fields.Add "Body118"
  Fields.Add "Body119"
  Fields.Add "Body120"
  Fields.Add "Body121"
  Fields.Add "Body122"
  Fields.Add "Body123"
  Fields.Add "Body124"
  Fields.Add "Body125"
  Fields.Add "Body126"
  Fields.Add "Body127"
  Fields.Add "Body128"
  Fields.Add "Body129"
  Fields.Add "Body130"
  Fields.Add "Body131"
  Fields.Add "Body132"
  Fields.Add "Body133"
  Fields.Add "Body134"
  Fields.Add "Body135"
  Fields.Add "Body136"
  Fields.Add "Body137"
  Fields.Add "Body138"
  Fields.Add "Body139"
  Fields.Add "Body140"
  Fields.Add "Body141"
  Fields.Add "Body142"
  Fields.Add "Body143"
  Fields.Add "Body144"
  Fields.Add "Body145"
  Fields.Add "Body146"
  Fields.Add "Body147"
  Fields.Add "Body148"
  Fields.Add "Body149"
  Fields.Add "Body150"
  Fields.Add "Body151"
  Fields.Add "Body152"
  Fields.Add "Body153"
  Fields.Add "Body154"
  Fields.Add "Body155"
  Fields.Add "Body156"
  Fields.Add "Body157"
  Fields.Add "Body158"
  Fields.Add "Body159"
  Fields.Add "Body160"
  Fields.Add "Body161"
  Fields.Add "Body162"
  Fields.Add "Body163"
  Fields.Add "Body164"
  Fields.Add "Body165"
  Fields.Add "Body166"
  Fields.Add "Body167"
  Fields.Add "Body168"
  Fields.Add "Body169"
  Fields.Add "Body170"
  Fields.Add "Body171"
  Fields.Add "Body172"
  Fields.Add "Body173"
  Fields.Add "Body174"
  Fields.Add "Body175"
  Fields.Add "Body176"
  Fields.Add "Body177"
  Fields.Add "Body178"
  Fields.Add "Body179"
  Fields.Add "Body180"
  Fields.Add "Body181"
  Fields.Add "Body182"
  Fields.Add "Body183"
  Fields.Add "Body184"
  Fields.Add "Body185"
  Fields.Add "Body186"
  Fields.Add "Body187"
  Fields.Add "Body188"
  Fields.Add "Body189"
  Fields.Add "Body190"
  Fields.Add "Body191"
  Fields.Add "Body192"
  Fields.Add "Body193"
  Fields.Add "Body194"
  Fields.Add "Body195"
  Fields.Add "Body196"
  Fields.Add "Body197"
  Fields.Add "Body198"
  Fields.Add "Body199"
  Fields.Add "Body200"
  Fields.Add "Body201"
  Fields.Add "Body202"
  Fields.Add "Body203"
  Fields.Add "Body204"
  Fields.Add "Body205"
  Fields.Add "Body206"
  Fields.Add "Body207"
  Fields.Add "Body208"
  Fields.Add "Body209"
  Fields.Add "Body210"
  Fields.Add "Body211"
  Fields.Add "Body212"
  Fields.Add "Body213"
  Fields.Add "Body214"
  Fields.Add "Body215"
  Fields.Add "Body216"
  Fields.Add "POAmt"
  Fields.Add "Add1"
  Fields.Add "Add2"
  Fields.Add "Add3"


End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    Unload frmLoadingRpt
    MsgBox "Error Number: " & Str(Number) & " " & Description, vbOKOnly, "Printer Error"
    Unload Me
  End If
  CancelDisplay = True
End Sub

'
Private Sub ActiveReport_FetchData(reof As Boolean)

Dim sLine As String
Dim arr() As String
'
'    ' We reached the end of the file we exit leaving the
'    ' eof parameter as True (default except on first call) that will
'    ' tell AR that we are done feeding data
'    ' otherwise we have to set the eof parameter to False so that
'    ' AR continues fetching data, until we're done
'    ' if the report had a data control, the value of the parameter
'    ' will be ignored, AR will always follow the data control's recordset
'    ' EOF property
On Error GoTo ERRORSTUFF
    If VBA.eof(hFile) Then
        reof = True
        Exit Sub
    Else
        reof = False
    End If

    Line Input #hFile, sLine
    arr = Split(sLine, "~")
    Fields("Header1").Value = arr(0)
    Fields("Header2").Value = arr(1)
    Fields("Header3").Value = arr(2)
    Fields("Header4").Value = arr(3)
    Fields("PODate").Value = arr(4)
    Fields("PONum").Value = arr(5)
    Fields("Vend1").Value = arr(6)
    Fields("Ship1").Value = arr(7)
    Fields("Vend2").Value = arr(8)
    Fields("Ship2").Value = arr(9)
    Fields("Vend3").Value = arr(10)
    Fields("Ship3").Value = arr(11)
    Fields("Vend4").Value = arr(12)
    Fields("Ship4").Value = arr(13)
    Fields("Ship5").Value = arr(14)
    Fields("ShipVia").Value = arr(15)
    Fields("FOB").Value = arr(16)
    Fields("Dept").Value = arr(17)
    Fields("ShipOn").Value = arr(18)
    Fields("Terms").Value = arr(19)
    Fields("Body1").Value = arr(20)
    Fields("Body2").Value = arr(21)
    Fields("Body3").Value = arr(22)
    Fields("Body4").Value = arr(23)
    Fields("Body5").Value = arr(24)
    Fields("Body6").Value = arr(25)
    Fields("Body7").Value = arr(26)
    Fields("Body8").Value = arr(27)
    Fields("Body9").Value = arr(28)
    Fields("Body10").Value = arr(29)
    Fields("Body11").Value = arr(30)
    Fields("Body12").Value = arr(31)
    Fields("Body13").Value = arr(32)
    Fields("Body14").Value = arr(33)
    Fields("Body15").Value = arr(34)
    Fields("Body16").Value = arr(35)
    Fields("Body17").Value = arr(36)
    Fields("Body18").Value = arr(37)
    Fields("Body19").Value = arr(38)
    Fields("Body20").Value = arr(39)
    Fields("Body21").Value = arr(40)
    Fields("Body22").Value = arr(41)
    Fields("Body23").Value = arr(42)
    Fields("Body24").Value = arr(43)
    Fields("Body25").Value = arr(44)
    Fields("Body26").Value = arr(45)
    Fields("Body27").Value = arr(46)
    Fields("Body28").Value = arr(47)
    Fields("Body29").Value = arr(48)
    Fields("Body30").Value = arr(49)
    Fields("Body31").Value = arr(50)
    Fields("Body32").Value = arr(51)
    Fields("Body33").Value = arr(52)
    Fields("Body34").Value = arr(53)
    Fields("Body35").Value = arr(54)
    Fields("Body36").Value = arr(55)
    Fields("Body37").Value = arr(56)
    Fields("Body38").Value = arr(57)
    Fields("Body39").Value = arr(58)
    Fields("Body40").Value = arr(59)
    Fields("Body41").Value = arr(60)
    Fields("Body42").Value = arr(61)
    Fields("Body43").Value = arr(62)
    Fields("Body44").Value = arr(63)
    Fields("Body45").Value = arr(64)
    Fields("Body46").Value = arr(65)
    Fields("Body47").Value = arr(66)
    Fields("Body48").Value = arr(67)
    Fields("Body49").Value = arr(68)
    Fields("Body50").Value = arr(69)
    Fields("Body51").Value = arr(70)
    Fields("Body52").Value = arr(71)
    Fields("Body53").Value = arr(72)
    Fields("Body54").Value = arr(73)
    Fields("Body55").Value = arr(74)
    Fields("Body56").Value = arr(75)
    Fields("Body57").Value = arr(76)
    Fields("Body58").Value = arr(77)
    Fields("Body59").Value = arr(78)
    Fields("Body60").Value = arr(79)
    Fields("Body61").Value = arr(80)
    Fields("Body62").Value = arr(81)
    Fields("Body63").Value = arr(82)
    Fields("Body64").Value = arr(83)
    Fields("Body65").Value = arr(84)
    Fields("Body66").Value = arr(85)
    Fields("Body67").Value = arr(86)
    Fields("Body68").Value = arr(87)
    Fields("Body69").Value = arr(88)
    Fields("Body70").Value = arr(89)
    Fields("Body71").Value = arr(90)
    Fields("Body72").Value = arr(91)
    Fields("Body73").Value = arr(92)
    Fields("Body74").Value = arr(93)
    Fields("Body75").Value = arr(94)
    Fields("Body76").Value = arr(95)
    Fields("Body77").Value = arr(96)
    Fields("Body78").Value = arr(97)
    Fields("Body79").Value = arr(98)
    Fields("Body80").Value = arr(99)
    Fields("Body81").Value = arr(100)
    Fields("Body82").Value = arr(101)
    Fields("Body83").Value = arr(102)
    Fields("Body84").Value = arr(103)
    Fields("Body85").Value = arr(104)
    Fields("Body86").Value = arr(105)
    Fields("Body87").Value = arr(106)
    Fields("Body88").Value = arr(107)
    Fields("Body89").Value = arr(108)
    Fields("Body90").Value = arr(109)
    Fields("Body91").Value = arr(110)
    Fields("Body92").Value = arr(111)
    Fields("Body93").Value = arr(112)
    Fields("Body94").Value = arr(113)
    Fields("Body95").Value = arr(114)
    Fields("Body96").Value = arr(115)
    Fields("Body97").Value = arr(116)
    Fields("Body98").Value = arr(117)
    Fields("Body99").Value = arr(118)
    Fields("Body100").Value = arr(119)
    Fields("Body101").Value = arr(120)
    Fields("Body102").Value = arr(121)
    Fields("Body103").Value = arr(122)
    Fields("Body104").Value = arr(123)
    Fields("Body105").Value = arr(124)
    Fields("Body106").Value = arr(125)
    Fields("Body107").Value = arr(126)
    Fields("Body108").Value = arr(127)
    Fields("Body109").Value = arr(128)
    Fields("Body110").Value = arr(129)
    Fields("Body111").Value = arr(130)
    Fields("Body112").Value = arr(131)
    Fields("Body113").Value = arr(132)
    Fields("Body114").Value = arr(133)
    Fields("Body115").Value = arr(134)
    Fields("Body116").Value = arr(135)
    Fields("Body117").Value = arr(136)
    Fields("Body118").Value = arr(137)
    Fields("Body119").Value = arr(138)
    Fields("Body120").Value = arr(139)
    Fields("Body121").Value = arr(140)
    Fields("Body122").Value = arr(141)
    Fields("Body123").Value = arr(142)
    Fields("Body124").Value = arr(143)
    Fields("Body125").Value = arr(144)
    Fields("Body126").Value = arr(145)
    Fields("Body127").Value = arr(146)
    Fields("Body128").Value = arr(147)
    Fields("Body129").Value = arr(148)
    Fields("Body130").Value = arr(149)
    Fields("Body131").Value = arr(150)
    Fields("Body132").Value = arr(151)
    Fields("Body133").Value = arr(152)
    Fields("Body134").Value = arr(153)
    Fields("Body135").Value = arr(154)
    Fields("Body136").Value = arr(155)
    Fields("Body137").Value = arr(156)
    Fields("Body138").Value = arr(157)
    Fields("Body139").Value = arr(158)
    Fields("Body140").Value = arr(159)
    Fields("Body141").Value = arr(160)
    Fields("Body142").Value = arr(161)
    Fields("Body143").Value = arr(162)
    Fields("Body144").Value = arr(163)
    Fields("Body145").Value = arr(164)
    Fields("Body146").Value = arr(165)
    Fields("Body147").Value = arr(166)
    Fields("Body148").Value = arr(167)
    Fields("Body149").Value = arr(168)
    Fields("Body150").Value = arr(169)
    Fields("Body151").Value = arr(170)
    Fields("Body152").Value = arr(171)
    Fields("Body153").Value = arr(172)
    Fields("Body154").Value = arr(173)
    Fields("Body155").Value = arr(174)
    Fields("Body156").Value = arr(175)
    Fields("Body157").Value = arr(176)
    Fields("Body158").Value = arr(177)
    Fields("Body159").Value = arr(178)
    Fields("Body160").Value = arr(179)
    Fields("Body161").Value = arr(180)
    Fields("Body162").Value = arr(181)
    Fields("Body163").Value = arr(182)
    Fields("Body164").Value = arr(183)
    Fields("Body165").Value = arr(184)
    Fields("Body166").Value = arr(185)
    Fields("Body167").Value = arr(186)
    Fields("Body168").Value = arr(187)
    Fields("Body169").Value = arr(188)
    Fields("Body170").Value = arr(189)
    Fields("Body171").Value = arr(190)
    Fields("Body172").Value = arr(191)
    Fields("Body173").Value = arr(192)
    Fields("Body174").Value = arr(193)
    Fields("Body175").Value = arr(194)
    Fields("Body176").Value = arr(195)
    Fields("Body177").Value = arr(196)
    Fields("Body178").Value = arr(197)
    Fields("Body179").Value = arr(198)
    Fields("Body180").Value = arr(199)
    Fields("Body181").Value = arr(200)
    Fields("Body182").Value = arr(201)
    Fields("Body183").Value = arr(202)
    Fields("Body184").Value = arr(203)
    Fields("Body185").Value = arr(204)
    Fields("Body186").Value = arr(205)
    Fields("Body187").Value = arr(206)
    Fields("Body188").Value = arr(207)
    Fields("Body189").Value = arr(208)
    Fields("Body190").Value = arr(209)
    Fields("Body191").Value = arr(210)
    Fields("Body192").Value = arr(211)
    Fields("Body193").Value = arr(212)
    Fields("Body194").Value = arr(213)
    Fields("Body195").Value = arr(214)
    Fields("Body196").Value = arr(215)
    Fields("Body197").Value = arr(216)
    Fields("Body198").Value = arr(217)
    Fields("Body199").Value = arr(218)
    Fields("Body200").Value = arr(219)
    Fields("Body201").Value = arr(220)
    Fields("Body202").Value = arr(221)
    Fields("Body203").Value = arr(222)
    Fields("Body204").Value = arr(223)
    Fields("Body205").Value = arr(224)
    Fields("Body206").Value = arr(225)
    Fields("Body207").Value = arr(226)
    Fields("Body208").Value = arr(227)
    Fields("Body209").Value = arr(228)
    Fields("Body210").Value = arr(229)
    Fields("Body211").Value = arr(230)
    Fields("Body212").Value = arr(231)
    Fields("Body213").Value = arr(232)
    Fields("Body214").Value = arr(233)
    Fields("Body215").Value = arr(234)
    Fields("Body216").Value = arr(235)
    Fields("POAmt").Value = arr(236)
    Fields("Add1").Value = arr(237)
    Fields("Add2").Value = arr(238)
    Fields("Add3").Value = arr(239)


'    ' Here we set the values of the fields that we defines as unbound
'    ' or user defined.
'    For cnt = 1 To 238
'       Fields(headers(cnt)) = arr(cnt - 1)
'
'    Next
  Exit Sub
'    ("Fund").Value = arr(0)
ERRORSTUFF:

      Unload frmLoadingRpt
  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "ARptPOForms", "Fetch Data", Erl)
    Case emrExitProc:
      Resume Proc_Exit
    Case emrResume:
      Resume
    Case emrResumeNext:
      Resume Next
    Case Else
      '--- Technically, this should never happen.
      Resume Proc_Exit
  End Select

Proc_Exit:
  '--- Cleanup code goes here...
    Close
    Unload Me
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
    End If
    If KeyCode = vbKeyE Then
      Screen.MousePointer = vbHourglass
      ExportReport 1
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - POForm.xls, created in the Citipak Directory.", vbOKOnly
    End If
    If KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - POForm.txt, created in the Citipak Directory.", vbOKOnly
    End If
  End If
End Sub

Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
'  KillFile ReportFile$
End Sub

Private Sub ActiveReport_ReportEnd()
Dim STUF As Integer
  If hFile <> 0 Then
    Close #hFile
  End If
  Unload frmLoadingRpt
  DoEvents
  STUF = Me.Pages.Count
  Me.Show 1
End Sub


Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
  If Tool = "&Close" Then
    Unload Me
  End If
  If Tool = "Save/&Excel" Then
    Screen.MousePointer = vbHourglass
    ExportReport 1
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - POForm.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "/&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - POForm.txt, created in the Citipak Directory.", vbOKOnly
  End If
End Sub

Public Sub startrpt()
  Me.Run
End Sub
Private Sub ActiveReport_Initialize()
  Me.Toolbar.Tools.Add "&Close"
  Me.Toolbar.Tools.Add "Save/&Excel"
  Me.Toolbar.Tools.Add "/&Text"
  
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
        oEXL.FileName = outfile & "POForm.xls"
        oEXL.Export Me.Pages
        
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "POForm.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
''
''Me.Pages.Save "check.rdf"
End Sub
