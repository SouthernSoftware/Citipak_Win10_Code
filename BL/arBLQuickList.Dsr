VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arBLQuickList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Customer Quick Listing"
   ClientHeight    =   8868
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   ControlBox      =   0   'False
   Icon            =   "arBLQuickList.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15637
   SectionData     =   "arBLQuickList.dsx":08CA
End
Attribute VB_Name = "arBLQuickList"
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
  Open StartPath & "\BLRPTS\ARQKCUST.RPT" For Input As #hFile
  Fields.Add ("fldTown") '0)
  Fields.Add ("fldCustNum") '1)
  Fields.Add ("fldBillName") '2)
  Fields.Add ("fldProrate") '3)
  Fields.Add ("fldCustName") '4)
  Fields.Add ("fldLicense") '5)
  Fields.Add ("fldVDate") '6)
  Fields.Add ("fldBCat1") '7)
  Fields.Add ("fldBCat2") '8)
  Fields.Add ("fldBCat3") '9)
  Fields.Add ("fldBCat4") '10)
  Fields.Add ("fldBCat5") '11)
  Fields.Add ("fldBCat")
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
  Fields("fldTown").Value = arr(0)
  Fields("fldCustNum").Value = arr(1)
  Fields("fldBillName").Value = arr(2)
  Fields("fldProrate").Value = arr(3)
  Fields("fldCustName").Value = arr(4)
  Fields("fldLicense").Value = arr(5)
  Fields("fldVDate").Value = arr(6)
  Fields("fldBCat1").Value = arr(7)
  Fields("fldBCat2").Value = arr(8)
  Fields("fldBCat3").Value = arr(9)
  Fields("fldBCat4").Value = arr(10)
  Fields("fldBCat5").Value = arr(11)
  If Len(QPTrim$(arr(11))) > 0 Then
    Fields("fldBCat").Value = arr(7) + "/" + arr(8) + "/" + arr(9) + "/" + arr(10) + "/" + arr(11)
  ElseIf Len(QPTrim$(arr(10))) > 0 Then
    Fields("fldBCat").Value = arr(7) + "/" + arr(8) + "/" + arr(9) + "/" + arr(10)
  ElseIf Len(QPTrim$(arr(9))) > 0 Then
    Fields("fldBCat").Value = arr(7) + "/" + arr(8) + "/" + arr(9)
  ElseIf Len(QPTrim$(arr(8))) > 0 Then
    Fields("fldBCat").Value = arr(7) + "/" + arr(8)
  ElseIf Len(QPTrim$(arr(7))) > 0 Then
    Fields("fldBCat").Value = arr(7)
  Else
    Fields("fldBCat").Value = ""
  End If
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
      frmBLMessageBoxJr.Label1.Caption = "File - BLQCustList.xls, created in the Citipak Directory."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmBLMessageBoxJr.Label1.Caption = "File - BLQCustList.txt, created in the Citipak Directory."
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
    frmBLMessageBoxJr.Label1.Caption = "File - BLQCustList.xls, created in the Citipak Directory."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmBLMessageBoxJr.Label1.Caption = "File - BLQCustList.txt, created in the Citipak Directory."
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
        oEXL.FileName = outfile & "BLQCustList.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "BLQCustList.txt"
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
  Me.fldTimeDate.Text = Date
  Me.Zoom = -1
End Sub
Private Sub ReportFooter_Format()
  Line2.LineStyle = ddLSSolid
End Sub