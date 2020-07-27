VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARptTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ActiveReport1"
   ClientHeight    =   4380
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   12024
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   21209
   _ExtentY        =   7726
   SectionData     =   "ARptTest.dsx":0000
End
Attribute VB_Name = "ARptTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ReportFile As String
Private hFile As Integer
Public Sub GetName(RName As String)
  ReportFile$ = RName$
End Sub
Private Sub ActiveReport_DataInitialize()
   
    hFile = FreeFile
    Open ReportFile$ For Input As #hFile
'
'    ' This sets up the fields used in data binding
'    'Fields.Add "ProductID"
    Fields.Add "LineTyp"
    Fields.Add "Date"
    Fields.Add "Desc"
    Fields.Add "Ref"
    Fields.Add "Debit"
    Fields.Add "Credit"
    Fields.Add "BalSrc"
    
End Sub
'
Private Sub ActiveReport_FetchData(eof As Boolean)

Dim sLine As String
Dim arr() As String
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
    arr = Split(sLine, ",")
'    ' Here we set the values of the fields that we defines as unbound
'    ' or user defined.
    Fields("LineTyp").Value = arr(0)
    Fields("Date").Value = arr(1)
    Fields("Desc").Value = arr(2)
    Fields("Ref").Value = arr(3)
    Fields("Debit").Value = arr(4)
    Fields("Credit").Value = arr(5)
    Fields("BalSrc").Value = arr(6)

End Sub

Public Sub startrpt()
  Me.Run
End Sub


Private Sub ActiveReport_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
  If KeyCode = vbKeyEscape Then
    Unload Me
    KeyCode = 0
  End If
End Sub
'

Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
  KillFile ReportFile$
End Sub
Private Sub ActiveReport_ReportEnd()
    If hFile <> 0 Then
        Close #hFile
    End If
  Unload frmLoadingRpt
  Me.Show 1
End Sub
Private Sub ActiveReport_ToolbarClick(ByVal Tool As ddactivereports2.DDTool)
  If Tool = "&Close" Then
    Unload Me
  End If
  If Tool = "&Save Report" Then
  End If

End Sub


