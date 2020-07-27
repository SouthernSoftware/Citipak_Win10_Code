VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arGrossWage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gross Wage Report"
   ClientHeight    =   8850
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11640
   Icon            =   "arGrossWage.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20532
   _ExtentY        =   15610
   SectionData     =   "arGrossWage.dsx":08CA
End
Attribute VB_Name = "arGrossWage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private HFile As Integer
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
      MsgBox "File - GrossWageRpt.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - GrossWageRpt.txt, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    End If
  End If
End Sub

Private Sub ActiveReport_Terminate()
  Close
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
    MsgBox "File - GrossWageRpt.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - GrossWageRpt.txt, created in the Citipak Directory.", vbOKOnly
  End If
End Sub
Private Sub ExportReport(X As Integer)
  Dim oEXL As ActiveReportsExcelExport.ARExportExcel
  Dim oTXT As ActiveReportsTextExport.ARExportText
  Dim outfile As String
  If Right$(StartPath, 1) = ":" Then
    outfile = StartPath
  Else
    outfile = StartPath & "\"
  End If
  Select Case X
    Case 1   '"Excel"
        Set oEXL = New ActiveReportsExcelExport.ARExportExcel
        oEXL.FileName = outfile & "GrossWageRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "GrossWageRpt.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_DataInitialize()
  HFile = FreeFile
  Open StartPath & "\PRRPTS\GROSWAGEG.RPT" For Input As #HFile
  Fields.Add "fldEmployerph"
  Fields.Add "fldStartph"
  Fields.Add "fldEndph"
  Fields.Add "fldEmpNumdt"
  Fields.Add "fldEmpNamedt"
  Fields.Add "fldGrsPaydt"
  Fields.Add "fldFedGrsdt"
  Fields.Add "fldFedTaxdt"
  Fields.Add "fldSocGrsdt"
  Fields.Add "fldSocTaxdt"
  Fields.Add "fldMedGrsdt"
  Fields.Add "fldMedTaxdt"
  Fields.Add "fldEICdt"
  Fields.Add "fldStaTaxdt"
  Fields.Add "fldGrsPayrf"
  Fields.Add "fldFedGrsrf"
  Fields.Add "fldFedTaxrf"
  Fields.Add "fldSocGrsrf"
  Fields.Add "fldSocTaxrf"
  Fields.Add "fldMedGrsrf"
  Fields.Add "fldMedTaxrf"
  Fields.Add "fldEICrf"
  Fields.Add "fldStaTaxrf"
  Fields.Add "fldEndItgh"
  End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    Unload frmLoadingRpt
    MsgBox "Error Number: " & Str(Number) & " " & Description, vbOKOnly, "Printer Error"
    Unload Me
  End If
  CancelDisplay = True 'removes the error message
End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)
  Dim sLine As String
  Dim arr() As String
  
  If VBA.eof(HFile) Then
    eof = True
    Exit Sub
  Else
    eof = False
  End If
  Line Input #HFile, sLine
  arr = Split(sLine, "~")
  ' Here we set the values of the fields that we defines as unbound
  ' or user defined.
  Fields("fldEmployerph").Value = arr(0)
  Fields("fldStartph").Value = arr(1)
  Fields("fldEndph").Value = arr(2)
  Fields("fldEmpNumdt").Value = arr(3)
  Fields("fldEmpNamedt").Value = arr(4)
  Fields("fldGrsPaydt").Value = arr(5)
  Fields("fldFedGrsdt").Value = arr(6)
  Fields("fldFedTaxdt").Value = arr(7)
  Fields("fldSocGrsdt").Value = arr(8)
  Fields("fldSocTaxdt").Value = arr(9)
  Fields("fldMedGrsdt").Value = arr(10)
  Fields("fldMedTaxdt").Value = arr(11)
  Fields("fldEICdt").Value = arr(12)
  Fields("fldStaTaxdt").Value = arr(13)
  Fields("fldGrsPayrf").Value = arr(14)
  Fields("fldFedGrsrf").Value = arr(15)
  Fields("fldFedTaxrf").Value = arr(16)
  Fields("fldSocGrsrf").Value = arr(17)
  Fields("fldSocTaxrf").Value = arr(18)
  Fields("fldMedGrsrf").Value = arr(19)
  Fields("fldMedTaxrf").Value = arr(20)
  Fields("fldEICrf").Value = arr(21)
  Fields("fldStaTaxrf").Value = arr(22)
  Fields("fldEndItgh").Value = arr(23)
End Sub
Private Sub ActiveReport_ReportEnd()
  Unload frmLoadingRpt
  If HFile <> 0 Then
    Close #HFile
  End If
End Sub

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

Private Sub ActiveReport_ReportStart()
  Me.fldTimeDate.Text = Now
  Me.Zoom = -1
  Label24.Visible = False
End Sub

Private Sub Detail_Format()
  GroupHeader1.GroupValue = Fields("fldEndItgh").Value
    
End Sub

Private Sub PageHeader_Format()
  If GroupHeader1.GroupValue = "END" Then PageHeader.Height = 1164

End Sub

Private Sub ReportFooter_Format()
  Label24.Visible = True
  Label2.Visible = False
  Label3.Visible = False
  Label4.Visible = False
  Label5.Visible = False
  Label6.Visible = False
  Label7.Visible = False
  Label8.Visible = False
  Label9.Visible = False
  Label10.Visible = False
  Label11.Visible = False
  Label12.Visible = False
End Sub
