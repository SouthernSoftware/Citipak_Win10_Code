VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Begin VB.Form frmARViewer 
   BackColor       =   &H008F8265&
   BorderStyle     =   0  'None
   Caption         =   "Reprint"
   ClientHeight    =   9210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11730
   Icon            =   "frmARViewer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9210
   ScaleWidth      =   11730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARViewer21 
      Height          =   9180
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11676
      _ExtentX        =   20585
      _ExtentY        =   16193
      SectionData     =   "frmARViewer.frx":08CA
   End
End
Attribute VB_Name = "frmARViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim DedRec(1 To 50) As DedCodeRecType
  Dim DHandle As Integer
  Dim OutFileErr As Boolean

Private Sub ARViewer21_KeyUp(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = vbKeyEscape Then
    Unload frmARViewer
    DoEvents
    KeyCode = 0
    Exit Sub
  End If
  
  Select Case KeyCode
    Case vbKeyC:
      Unload frmARViewer
      DoEvents
      KeyCode = 0
    Case vbKeyE:
      Screen.MousePointer = vbHourglass
      ExportReport 1
      Screen.MousePointer = vbDefault
      If OutFileErr = True Then
        OutFileErr = False
        Exit Sub
      End If
      DoEvents
      MsgBox "File - " & ThisRpt & ".xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    Case vbKeyT
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      If OutFileErr = True Then
        OutFileErr = False
        Exit Sub
      End If
      DoEvents
      MsgBox "File - " & ThisRpt & ".txt, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub ARViewer21_ToolbarClick(ByVal Tool As DDActiveReportsViewer2Ctl.IDDTool)
  
  If Tool = "&Close" Then
    Unload Me
  End If
  If Tool = "Save/&Excel" Then
    Screen.MousePointer = vbHourglass
    ExportReport 1
    Screen.MousePointer = vbDefault
    If OutFileErr = True Then
      OutFileErr = False
      Exit Sub
    End If
    DoEvents
    MsgBox "File - " & ThisRpt & ".xls, created in the Citipak Directory.", vbOKOnly
  End If
  
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    If OutFileErr = True Then
      OutFileErr = False
      Exit Sub
    End If
    DoEvents
    MsgBox "File - " & ThisRpt & ".txt, created in the Citipak Directory.", vbOKOnly
  End If

End Sub

Private Sub LoadRpt()
  Dim x As Integer
  Unload frmLoadingReprint
  Select Case ThisRpt$
    Case "  GL Register"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\PRGLIFG.RDF"
    Case "  GL Register Non-Split"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\PRGLIFNSG.RDF"
    Case "  Earnings Register"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\REGISTERG.RDF"
    Case "  Earnings Register Non-Split"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\REGISTERNSG.RDF"
    Case "  YTD Wage Report"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\YTDWAGEG.RDF"
    Case "  Terminated Employee Report"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\EMPRINTTERMEMPLISTG.RDF"
    Case "  Supplemental Retirement Report"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\401KG.RDF"
    Case "  SEPP Contribution Report"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\SeppContG.RDF"
    Case "  NC State Retirement Report"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\RETIREG.RDF"
    Case "  SC State Retirement Report"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\SCRETIREG.RDF"
    Case "  VA State Retirement Report"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\VARETIREG.RDF"
    Case "  Employee List Report"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\EMPRINTEMPLISTG.RDF"
    Case "  Benefit Accrual Report"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\BENEACCRG.RDF"
    Case "  Gross Wage Report"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\GROSWAGEG.RDF"
    Case "  ESC 1st Quarter Report"
      GlblQtr = 1
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\ESCQTR1.RDF"
    Case "  ESC 2nd Quarter Report"
      GlblQtr = 2
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\ESCQTR2.RDF"
    Case "  ESC 3rd Quarter Report"
      GlblQtr = 3
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\ESCQTR3.RDF"
    Case "  ESC 4th Quarter Report"
      GlblQtr = 4
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\ESCQTR4.RDF"
    Case "  Employee Data Report"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\EMPDATAG.RDF"
    Case "  Checks Issued by Employee"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\CHKISSUEG.RDF"
    Case "  Employee Earnings History"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\EMPHISTG.RDF"
    Case "  Employee Earnings History Summary"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\EMPHISTSUMG.RDF"
    Case "  Earnings Distribution Register"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\DISTRIBUACCTNUMG.RDF"
    Case "  Fund Number Register"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\DISTRIBUFUNDNUM.RDF"
    Case "  Earnings Distribution Register non-Split"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\DISTRIBUNS.RDF"
    Case "  Worker's Comp Report"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\COMPWAGEG.RDF"
    Case "  Checks in Numerical Order Report"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\CHKSBYRANGEG.RDF"
    
    Case "  Employees to Draft Report"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\PPDFG.RDF"
    Case "  Employee Draft List"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\EMPDFLSTG.RDF"
    Case "  Accrual Report"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\ACCRUALG.RDF"
    Case "  Check Register"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\CHECKREGG.RDF"
    Case "  W2 Report"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\W2REPORTG.RDF"
    Case "  Manual Transaction Register"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\MANREGISG.RDF"
    Case "  Employee Emergency Information"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\EMERGENCYG.RDF"
    Case "  Employee Pay Rate Report"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRPTS\PAYRATEG.RDF"
    Case "  941 Assistance Report"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\941FORMSG.RDF"
    Case "  Void Check Review"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\VOIDPRNG.RDF"
    Case "  Tax Fringe Report"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\TAXFRING.RDF"
    Case "  Deduction Report ALL"
      ARViewer21.Zoom = -1
      ARViewer21.Pages.Load StartPath & "\PRRDF\DEDUCALL.RDF"
    Case Else
      If InStr(1, ThisRpt, "  Deduction") = 1 Then
        For x = 1 To 50
          If ThisRpt$ = "  Deduction " & DedRec(x).DCDESC1 & " Report" Then
            DeductionSelNum = x
            ARViewer21.Zoom = -1
            ARViewer21.Pages.Load StartPath & "\PRRDF\DEDUCTG" & x & ".RDF"
            Exit For
          End If
        Next x
      Else
        GoTo Done
      End If
Done:
  End Select

End Sub

Private Sub cmdExit_Click()
  Unload frmARViewer
  DoEvents
End Sub

Private Sub Form_Load()
  Dim x As Integer
  
  OutFileErr = False
  
  OpenDedCodeFile DHandle
  For x = 1 To 50
    Get DHandle, x, DedRec(x)
  Next x
  Close DHandle
  
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  
  ARViewer21.ToolBar.Tools.Add "&Close"
  ARViewer21.ToolBar.Tools.Add "Save/&Excel"
  ARViewer21.ToolBar.Tools.Add "&Text"
  
  Call SaveRpt

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub
Private Sub ExportReport(x As Integer)
  Dim oEXL As ActiveReportsExcelExport.ARExportExcel
  Dim oTXT As ActiveReportsTextExport.ARExportText
  Dim outfile As String
  Dim newtxtFile$, cnt As Integer
  Dim newxlsFile$
  Dim Report As ActiveReport
  
  Select Case ThisRpt$
    Case "  GL Register"
      newxlsFile = "GLRegister.xls"
      newtxtFile = "GLRegister.txt"
      Set Report = arGLRegister
    Case "  GL Register Non-Split"
      newxlsFile = "GLRegisterNS.xls"
      newtxtFile = "GLRegisterNS.txt"
      Set Report = arGLRegisterNS
    Case "  Earnings Register"
      newxlsFile = "EarningsRegister.xls"
      newtxtFile = "EarningsRegister.txt"
      Set Report = arPRRegister
    Case "  Earnings Register Non-Split"
      newxlsFile = "EarningsRegisterNS.xls"
      newtxtFile = "EarningsRegisterNS.txt"
      Set Report = arPayRollRegisterNS
    Case "  YTD Wage Report"
      newxlsFile = "YTDWage.xls"
      newtxtFile = "YTDWage.txt"
      Set Report = arYTDWageRpt
    Case "  Terminated Employee Report"
      newxlsFile = "TermEmp.xls"
      newtxtFile = "TermEmp.txt"
      Set Report = arTermEmpRpt
    Case "  Supplemental Retirement Report"
      newxlsFile = "SuppRet.xls"
      newtxtFile = "SuppRet.txt"
      Set Report = arSuppRet
    Case "  SEPP Contribution Report"
      newxlsFile = "SEPPCont.xls"
      newtxtFile = "SEPPCont.txt"
      Set Report = arSeppCon
    Case "  NC State Retirement Report"
      newxlsFile = "NCStateRet.xls"
      newtxtFile = "NCStateRet.txt"
      Set Report = arRetRpt
    Case "  SC State Retirement Report"
      newxlsFile = "SCStateRet.xls"
      newtxtFile = "SCStateRet.txt"
      Set Report = arRetRptSC
    Case "  VA State Retirement Report"
      newxlsFile = "VAStateRet.xls"
      newtxtFile = "VAStateRet.txt"
      Set Report = arRetRptVA
    
    Case "  Employee List Report"
      newxlsFile = "EMPLIST.xls"
      newtxtFile = "EMPLIST.txt"
      Set Report = arPrintAlphaNum
    
    Case "  Benefit Accrual Report"
      newxlsFile = "BeneAccr.xls"
      newtxtFile = "BeneAccr.txt"
      Set Report = arLvBnftRpt
    Case "  Gross Wage Report"
      newxlsFile = "GrosWage.xls"
      newtxtFile = "GrosWage.txt"
      Set Report = arGrossWage
    Case "  ESC 1st Quarter Report"
      newxlsFile = "ESC1.xls"
      newtxtFile = "ESC1.txt"
      Set Report = arESCRpt
    Case "  ESC 2nd Quarter Report"
      newxlsFile = "ESC2.xls"
      newtxtFile = "ESC2.txt"
      Set Report = arESCRpt
    Case "  ESC 3rd Quarter Report"
      newxlsFile = "ESC3.xls"
      newtxtFile = "ESC3.txt"
      Set Report = arESCRpt
    Case "  ESC 4th Quarter Report"
      newxlsFile = "ESC4.xls"
      newtxtFile = "ESC4.txt"
      Set Report = arESCRpt
    Case "  Employee Data Report"
      newxlsFile = "EmpData.xls"
      newtxtFile = "EmpData.txt"
      Set Report = arEmpDataRpt
    Case "  Checks Issued by Employee"
      newxlsFile = "EmpChksIss.xls"
      newtxtFile = "EmpChksIss.txt"
      Set Report = arEmpChksIssued
    Case "  Employee Earnings History"
      newxlsFile = "EmpEarnHist.xls"
      newtxtFile = "EmpEarnHist.txt"
      Set Report = arEarningsHistory
    Case "  Employee Earnings History Summary"
      newxlsFile = "EmpEarnHistSum.xls"
      newtxtFile = "EmpEarnHistSum.txt"
      Set Report = arEarnHistSumOnly
    Case "  Earnings Distribution Register"
      newxlsFile = "AcctNumReg.xls"
      newtxtFile = "AcctNumReg.txt"
      Set Report = arDistRegAcctNum
'    Case "  Fund Number Register"
'      newxlsFile = "FundNumReg.xls"
'      newtxtFile = "FundNumReg.txt"
'      Set Report = arDistRegFundNum
    Case "  Earnings Distribution Register non-Split"
      newxlsFile = "EarnDistRegisterNS.xls"
      newtxtFile = "EarnDistRegisterNS.txt"
      Set Report = arEarnDistRegNS
    Case "  Worker's Comp Report"
      newxlsFile = "CompWage.xls"
      newtxtFile = "CompWage.txt"
      Set Report = arCompWageRpt
    Case "  Checks in Numerical Order Report"
      newxlsFile = "ChksByNum.xls"
      newtxtFile = "ChksByNum.txt"
      Set Report = arChksByNumRpt
    Case "  Employees to Draft Report"
      newxlsFile = "Emp2Draft.xls"
      newtxtFile = "Emp2Draft.txt"
      Set Report = arEmp2Draft
    Case "  Employee Draft List"
      newxlsFile = "EmpDraftList.xls"
      newtxtFile = "EmpDraftList.txt"
      Set Report = arEmpDraftList
    Case "  Accrual Report"
      newxlsFile = "AccrualRpt.xls"
      newtxtFile = "AccrualRpt.txt"
      Set Report = arLvBnfts
    Case "  Check Register"
      newxlsFile = "CheckReg.xls"
      newtxtFile = "CheckReg.txt"
      Set Report = arCheckRegister
    Case "  W2 Report"
      newxlsFile = "W2Report.xls"
      newtxtFile = "W2Report.txt"
      Set Report = arW2Report
    Case "  Manual Transaction Register"
      newxlsFile = "ManTranReg.xls"
      newtxtFile = "ManTranReg.txt"
      Set Report = arManTranEntry
    Case "  Employee Emergency Information"
      newxlsFile = "EmergencyRpt.xls"
      newtxtFile = "EmergencyRpt.txt"
      Set Report = arEmergency
    Case "  Employee Pay Rate Report"
      newxlsFile = "PayRateRpt.xls"
      newtxtFile = "PayRateRpt.txt"
      Set Report = arPayRate
    Case "  941 Assistance Report"
      newxlsFile = "941Rpt.xls"
      newtxtFile = "941Rpt.txt"
      Set Report = ar941
    Case "  Void Check Review"
      newxlsFile = "VoidChkRvw.xls"
      newtxtFile = "VoidChkRvw.txt"
      Set Report = arPRVoidChkPrint
    Case "  Tax Fringe Report"
      newxlsFile = "TaxFringeRpt.xls"
      newtxtFile = "TaxFringeRpt.txt"
      Set Report = arPRVoidChkPrint
    Case "  Deduction Report ALL" 'added 9/4/03
      newxlsFile = "DeducAllRpt.xls"
      newtxtFile = "DeducAllRpt.txt"
      Set Report = arTaxFringRpt
    Case Else
     If InStr(1, ThisRpt, "  Deduction") = 1 Then
        For cnt = 1 To 50
          If ThisRpt$ = "  Deduction " & DedRec(cnt).DCDESC1 & " Report" Then
            newxlsFile = QPTrim$(DedRec(cnt).DCDESC1) & ".xls"
            newtxtFile = QPTrim$(DedRec(cnt).DCDESC1) & ".txt"
            Set Report = arPRDeductionRpt
            Exit For
          End If
        Next cnt
      Else
        MsgBox "ERROR in creating outfile"
        OutFileErr = True
        Exit Sub
      End If
    End Select
    
  If Right$(StartPath, 1) = ":" Then
    outfile = StartPath
  Else
    outfile = StartPath & "\"
  End If
  
  Select Case x
    Case 1   '"Excel"
        Set oEXL = New ActiveReportsExcelExport.ARExportExcel
        oEXL.FileName = outfile & newxlsFile
        Call SaveRpt
        oEXL.Export Report.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & newtxtFile
        Call SaveRpt
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Report.Pages
  End Select
End Sub

Private Sub SaveRpt()
  Dim x As Integer
  Select Case ThisRpt$
    Case "  GL Register"
      arGLRegister.Run
      arGLRegister.Pages.save StartPath & "\PRRDF\PRGLIFG.RDF"
      Unload arGLRegister
    Case "  GL Register Non-Split"
      arGLRegisterNS.Run
      arGLRegisterNS.Pages.save StartPath & "\PRRDF\PRGLIFNSG.RDF"
      Unload arGLRegisterNS
    Case "  Earnings Register"
      arPRRegister.Run
      arPRRegister.Pages.save StartPath & "\PRRDF\REGISTERG.RDF"
      Unload arPRRegister
    Case "  Earnings Register Non-Split"
      arPayRollRegisterNS.Run
      arPayRollRegisterNS.Pages.save StartPath & "\PRRDF\REGISTERNSG.RDF"
      Unload arPayRollRegisterNS
    Case "  YTD Wage Report"
      arYTDWageRpt.Run
      arYTDWageRpt.Pages.save StartPath & "\PRRDF\YTDWAGEG.RDF"
      Unload arYTDWageRpt
    Case "  Terminated Employee Report"
      arTermEmpRpt.Run
      arTermEmpRpt.Pages.save StartPath & "\PRRDF\EMPRINTTERMEMPLISTG.RDF"
      Unload arTermEmpRpt
    Case "  Supplemental Retirement Report"
      arSuppRet.Run
      arSuppRet.Pages.save StartPath & "\PRRDF\401KG.RDF"
      Unload arSuppRet
    Case "  SEPP Contribution Report"
      arSeppCon.Run
      arSeppCon.Pages.save StartPath & "\PRRDF\SeppContG.RDF"
      Unload arSeppCon
    Case "  NC State Retirement Report"
      arRetRpt.Run
      arRetRpt.Pages.save StartPath & "\PRRDF\RETIREG.RDF"
      Unload arRetRpt
    Case "  SC State Retirement Report"
      arRetRptSC.Run
      arRetRptSC.Pages.save StartPath & "\PRRDF\SCRETIREG.RDF"
      Unload arRetRptSC
    Case "  VA State Retirement Report"
      arRetRptVA.Run
      arRetRptVA.Pages.save StartPath & "\PRRDF\VARETIREG.RDF"
      Unload arRetRptVA
    Case "  Employee List Report"
      arPrintAlphaNum.Run
      arPrintAlphaNum.Pages.save StartPath & "\PRRDF\EMPRINTEMPLISTG.RDF"
      Unload arPrintAlphaNum
    Case "  Benefit Accrual Report"
      arLvBnftRpt.Run
      arLvBnftRpt.Pages.save StartPath & "\PRRDF\BENEACCRG.RDF"
      Unload arLvBnftRpt
    Case "  Gross Wage Report"
      arGrossWage.Run
      arGrossWage.Pages.save StartPath & "\PRRDF\GROSWAGEG.RDF"
      Unload arGrossWage
    Case "  ESC 1st Quarter Report"
      GlblQtr = 1
      arESCRpt.Run
      arESCRpt.Pages.save StartPath & "\PRRDF\ESCQTR1.RDF"
      Unload arESCRpt
    Case "  ESC 2nd Quarter Report"
      GlblQtr = 2
      arESCRpt.Run
      arESCRpt.Pages.save StartPath & "\PRRDF\ESCQTR2.RDF"
      Unload arESCRpt
    Case "  ESC 3rd Quarter Report"
      GlblQtr = 3
      arESCRpt.Run
      arESCRpt.Pages.save StartPath & "\PRRDF\ESCQTR3.RDF"
      Unload arESCRpt
    Case "  ESC 4th Quarter Report"
      GlblQtr = 4
      arESCRpt.Run
      arESCRpt.Pages.save StartPath & "\PRRDF\ESCQTR4.RDF"
      Unload arESCRpt
    Case "  Employee Data Report"
      arEmpDataRpt.Run
      arEmpDataRpt.Pages.save StartPath & "\PRRDF\EMPDATAG.RDF"
      Unload arEmpDataRpt
    Case "  Checks Issued by Employee"
      arEmpChksIssued.Run
      arEmpChksIssued.Pages.save StartPath & "\PRRDF\CHKISSUEG.RDF"
      Unload arEmpChksIssued
    Case "  Employee Earnings History"
      arEarningsHistory.Run
      arEarningsHistory.Pages.save StartPath & "\PRRDF\EMPHISTG.RDF"
      Unload arEarningsHistory
    Case "  Employee Earnings History Summary"
      arEarnHistSumOnly.Run
      arEarnHistSumOnly.Pages.save StartPath & "\PRRDF\EMPHISTSUMG.RDF"
      Unload arEarnHistSumOnly
    Case "  Earnings Distribution Register"
      arDistRegAcctNum.Run
      arDistRegAcctNum.Pages.save StartPath & "\PRRDF\DISTRIBUACCTNUMG.RDF"
      Unload arDistRegAcctNum
'    Case "  Fund Number Register"
'      arDistRegFundNum.Run
'      arDistRegFundNum.Pages.save StartPath & "\PRRDF\DISTRIBUFUNDNUM.RDF"
'      Unload arDistRegFundNum
    Case "  Earnings Distribution Register non-Split"
      arEarnDistRegNS.Run
      arEarnDistRegNS.Pages.save StartPath & "\PRRDF\DISTRIBUNS.RDF"
      Unload arEarnDistRegNS
    Case "  Worker's Comp Report"
      arCompWageRpt.Run
      arCompWageRpt.Pages.save StartPath & "\PRRDF\COMPWAGEG.RDF"
      Unload arCompWageRpt
    Case "  Checks in Numerical Order Report"
      arChksByNumRpt.Run
      arChksByNumRpt.Pages.save StartPath & "\PRRDF\CHKSBYRANGEG.RDF"
      Unload arChksByNumRpt
    Case "  Employees to Draft Report"
      arEmp2Draft.Run
      arEmp2Draft.Pages.save StartPath & "\PRRDF\PPDFG.RDF"
      Unload arEmp2Draft
    Case "  Employee Draft List"
      arEmpDraftList.Run
      arEmpDraftList.Pages.save StartPath & "\PRRDF\EMPDFLSTG.RDF"
      Unload arEmpDraftList
    Case "  Accrual Report"
      arLvBnfts.Run
      arLvBnfts.Pages.save StartPath & "\PRRDF\ACCRUALG.RDF"
      Unload arLvBnfts
    Case "  Check Register"
      arCheckRegister.Run
      arCheckRegister.Pages.save StartPath & "\PRRDF\CHECKREGG.RDF"
      Unload arCheckRegister
    Case "  W2 Report"
      arW2Report.Run
      arW2Report.Pages.save StartPath & "\PRRDF\W2REPORTG.RDF"
      Unload arW2Report
    Case "  Manual Transaction Register"
      arManTranEntry.Run
      arManTranEntry.Pages.save StartPath & "\PRRDF\MANREGISG.RDF"
      Unload arManTranEntry
    Case "  Employee Emergency Information"
      arEmergency.Run
      arEmergency.Pages.save StartPath & "\PRRDF\EMERGENCYG.RDF"
      Unload arEmergency
    Case "  Employee Pay Rate Report"
      arPayRate.Run
      arPayRate.Pages.save StartPath & "\PRRDF\PAYRATEG.RDF"
      Unload arPayRate
    Case "  941 Assistance Report"
      ar941.Run
      ar941.Pages.save StartPath & "\PRRDF\941FORMSG.RDF"
      Unload ar941
    Case "  Void Check Review"
      arPRVoidChkPrint.Run
      arPRVoidChkPrint.Pages.save StartPath & "\PRRDF\VOIDPRNG.RDF"
      Unload ar941
    Case "  Tax Fringe Report"
      arTaxFringRpt.Run
      arTaxFringRpt.Pages.save StartPath & "\PRRDF\TAXFRING.RDF"
      Unload ar941
    Case "  Deduction Report ALL"
      arPRDeducAll.Run
      arPRDeducAll.Pages.save StartPath & "\PRRDF\DEDUCALL.RDF"
      Unload arPRDeducAll
    Case Else
      If InStr(1, ThisRpt, "  Deduction") = 1 Then
        For x = 1 To 50
          If ThisRpt$ = "  Deduction " & DedRec(x).DCDESC1 & " Report" Then
            DeductionSelNum = x
            arPRDeductionRpt.Run
            arPRDeductionRpt.Pages.save StartPath & "\PRRDF\DEDUCTG" & x & ".RDF"
            Unload arPRDeductionRpt
           Exit For
          End If
        Next x
      Else
        MsgBox "ERROR In Selection"
      End If
  End Select
  Call LoadRpt
SelErr:
End Sub

