VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARSubBookSum 
   BorderStyle     =   0  'None
   ClientHeight    =   4788
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9240
   Icon            =   "ARSubBookSum.dsx":0000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   16298
   _ExtentY        =   8446
   SectionData     =   "ARSubBookSum.dsx":08CA
End
Attribute VB_Name = "ARSubBookSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportFile As String
Private hFile As Integer
Dim rpt As ActiveReport
Dim tempTot As Double, TotrptVend As Double
Public distDet As Boolean

'Public Sub GetName(RName As String)
'  ReportFile$ = RName$
'
'End Sub
Private Sub ActiveReport_DataInitialize()
    ReportFile$ = Me.ParentReport.SubFile2
    hFile = FreeFile
    Open ReportFile$ For Input As #hFile
'
'    ' This sets up the fields used in data binding
'    'Fields.Add "ProductID"
    Fields.Add "Book"
    Fields.Add "Count"
    Fields.Add "Cash"
    Fields.Add "Check"
    Fields.Add "Charge"
    Fields.Add "Total"
    Fields.Add "Change"
End Sub
Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    Unload frmLoadingRpt
    MsgBox "Error Number: " & Str(Number) & " " & Description, vbOKOnly, "Printer Error"
    Unload Me
  End If
  CancelDisplay = True
End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)

Dim sLine As String
Dim arr() As String
'On Error GoTo ERRORSTUFF
'    ' We reached the end of the file we exit leaving the
'    ' eof parameter as True (default except on first call) that will
'    ' tell AR that we are done feeding data
'    ' otherwise we have to set the eof parameter to False so that
'    ' AR continues fetching data, until we're done
'    ' if the report had a data control, the value of the parameter
'    ' will be ignored, AR will always follow the data control's recordset
'    ' EOF property
    If VBA.eof(hFile) Then
        eof = True
        Exit Sub
    Else
        eof = False
    End If

    Line Input #hFile, sLine
    arr = Split(sLine, "~")
'    ' Here we set the values of the fields that we defines as unbound
'    ' or user defined.
    Fields("Book").Value = arr(0)
    Fields("Count").Value = arr(1)
    Fields("Cash").Value = arr(2)
    Fields("Check").Value = arr(3)
    Fields("Charge").Value = arr(4)
    Fields("Total").Value = arr(5)
    Fields("Change").Value = arr(6)
'If something wrong in file give message instead of crashing
Exit Sub
ERRORSTUFF:
'      Unload frmLoadingRpt
'  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "ARptVendHist", "Fetch Data", Erl)
'    Case emrExitProc:
'      Resume Proc_Exit
'    Case emrResume:
'      Resume
'    Case emrResumeNext:
'      Resume Next
'    Case Else
'      '--- Technically, this should never happen.
'      Resume Proc_Exit
'  End Select
   MsgBox "Err.Number, Err.Description, Err.Source", vbOKOnly, "Error"
   GoSub Proc_Exit
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    Unload Me
End Sub
Private Sub ActiveReport_ReportEnd()
    If hFile <> 0 Then
        Close #hFile
    End If
End Sub
