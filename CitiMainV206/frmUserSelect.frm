VERSION 5.00
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserSelect 
   BackColor       =   &H008A775B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Maintenance"
   ClientHeight    =   8880
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   12195
   Icon            =   "frmUserSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboUsers 
      Height          =   384
      Left            =   3624
      TabIndex        =   0
      Top             =   3672
      Width           =   4944
      _Version        =   196608
      _ExtentX        =   8721
      _ExtentY        =   677
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BackColor       =   16777215
      ForeColor       =   0
      Text            =   ""
      Columns         =   3
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   3
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   2
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   8421504
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   0
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   8421504
      BorderDropShadowWidth=   3
      ScrollHScale    =   2
      ScrollHInc      =   0
      ColsFrozen      =   0
      ScrollBarV      =   1
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   0
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   2
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   3
      DataFieldList   =   ""
      ColumnEdit      =   -1
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   12632256
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   12632256
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   12632256
      ExtendRow       =   0
      ListPosition    =   0
      ButtonThreeDAppearance=   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      AutoSearchFill  =   -1  'True
      AutoSearchFillDelay=   100
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmUserSelect.frx":08CA
   End
   Begin VB.CommandButton cmdPrintPassList 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F5 &List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      HelpContextID   =   55
      Left            =   6226
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   1332
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F4 &Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   4634
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   1332
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F2 &New"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   3042
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   1332
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Esc E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   7818
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5040
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   8625
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7117
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "8:04 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "5/15/2018"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select F2 to Add New User."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   2496
      TabIndex        =   9
      Top             =   3096
      Width           =   7188
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "No Users - Add New or Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4518
      TabIndex        =   8
      Top             =   1992
      Visible         =   0   'False
      Width           =   3156
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select the User in the list below and F4 to Edit or"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   2508
      TabIndex        =   7
      Top             =   2712
      Width           =   7188
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      Height          =   3252
      Left            =   2406
      Top             =   2496
      Width           =   7380
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   780
      Left            =   2580
      Top             =   816
      Width           =   7020
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Password Maintenance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   3312
      TabIndex        =   5
      Top             =   1008
      Width           =   5580
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   900
      Left            =   2592
      Top             =   696
      Width           =   7020
   End
End
Attribute VB_Name = "frmUserSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim Citipass As CitiPassType
Dim TempRec As Integer

Private Sub cmdEdit_Click()
  If fpcboUsers.ListIndex <> -1 Then
    fpcboUsers.col = 0
    TempRec = QPTrim(fpcboUsers.ColText)
    frmEnterEditPass.Rec2Form (TempRec)
    frmEnterEditPass.Show
  Else
    MsgBox "You Must First Select A User To Edit.", vbOKOnly, "Invalid Selection"
  End If
End Sub
Private Sub cmdPrintPassList_Click()
  frmReportOpt.Show 1
  If rptopt = 1 Then
    PrintPassListReport True
  ElseIf rptopt = 2 Then
    PrintPassListReport False
  End If
End Sub

Private Sub cmdExit_Click()
  'ClearInUse (PWcnt)
'  LevelPass = 0
'  PWcnt = 0
'  PWUser = ""
  frmPassLogin.Show
  Unload frmUserSelect
End Sub

Private Sub cmdNew_Click()
  Dim NumPassRecs As Integer, cnt As Integer, PassRecLen As Integer
  'OpenCitiPassFile CitiPassFile, NumPassRecs
  PassRecLen = Len(Citipass)
  NumPassRecs = LOF(CPAdminhand) \ PassRecLen

  fpcboUsers.Clear
  fpcboUsers.ListIndex = -1
  DoEvents
  frmEnterEditPass.Recnum = 0
  frmEnterEditPass.fpControlNum = 0
  frmEnterEditPass.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExit_Click
      KeyCode = 0
    Case vbKeyF2:
      cmdNew_Click
      KeyCode = 0
    Case vbKeyF4:
      cmdEdit_Click
      KeyCode = 0
    Case vbKeyF5:
      cmdPrintPassList_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Dim NumPassRecs As Integer, PassRecLen As Integer
  PassRecLen = Len(Citipass)
  NumPassRecs = LOF(CPAdminhand) \ PassRecLen
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me

If NumPassRecs > 0 Then
  FillUsers fpcboUsers
Else
  fpcboUsers.Enabled = False
  cmdEdit.Enabled = False
  Label3.Visible = True
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  MainLog "Out PW Entry/Edit"
  LevelPass = 0
  PWcnt = 0
  PWUser = ""
  Close CPAdminhand
  frmPassLogin.Show
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub fpcboUsers_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboUsers.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboUsers.ListIndex = -1
    fpcboUsers.Action = ActionClearSearchBuffer
  End If
  If fpcboUsers.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      cmdEdit.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        cmdExit.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub
Public Sub FillUsers(txt As fpCombo)
  Dim NumPassRecs As Integer, cnt As Integer, PassRecLen As Integer
  'OpenCitiPassFile CitiPassFile, NumPassRecs
  PassRecLen = Len(Citipass)
  NumPassRecs = LOF(CPAdminhand) \ PassRecLen
  txt.Row = -1
  For cnt = 1 To NumPassRecs
    Get CPAdminhand, cnt, Citipass
    If Not Citipass.DelFlag Then
      txt.InsertRow = Str$(cnt) & Chr$(9) & Str$(Citipass.PassNum) & Chr$(9) & QPTrim(Citipass.username)
    End If
  Next
  'Close CitiPassFile
End Sub
'Private Sub Cleanup()
'  Dim NumPassRecs As Integer, cnt As Integer, CitiPassFile As Integer
'  Dim NPRecs As Integer, NCitiPassFile As Integer, NRecLen As Integer
'  OpenCitiPassFile CitiPassFile, NumPassRecs
'  NRecLen = Len(CitiPass)
'  If NumPassRecs = 0 Then
'    Close
'    Exit Sub
'  End If
'
'  NCitiPassFile = FreeFile
'  Open "NCitipass.dat" For Output As #NCitiPassFile
'  Close NCitiPassFile
'  Open "NCitiPass.dat" For Random Shared As NCitiPassFile Len = NRecLen
'  For cnt = 1 To NumPassRecs
'    Get CitiPassFile, cnt, CitiPass
'    If Not CitiPass.DelFlag Then
'      Put NCitiPassFile, , CitiPass
'    End If
'  Next
'  Close CitiPassFile
'  Close NCitiPassFile
'  Kill "CitiPass.dat"
'  Name "NCitiPass.dat" As "CitiPass.dat"
'End Sub
Private Sub PrintPassListReport(grpt As Boolean)
  Dim NumPassRecs As Integer, cnt As Integer, PassRecLen As Integer
  Dim MaxLines As Integer, cnt2 As Integer
  Dim Linecnt As Integer, md As String, cnt3 As Integer
  Dim PRNFile As Integer, Howmany As Integer
  Dim ReportFile As String, ToPrint As String
  Dim FF As String, Header As String, mdtprint As String
  
 '  Stop
   'Define vars used for printing
   MaxLines = 55
   FF$ = Chr$(12)
   Header$ = "User PassCode Listing"

'   LibFile2Scrn "GL.QSL", "BG", MonoCode, Attribute, ErrorCode
'   PrintHelp "Processing report. Please wait."
   PRNFile = FreeFile
   ReportFile$ = "PCLst.rpt"
   Open ReportFile$ For Output As #PRNFile
GoSub PrintPageHeader
  'OpenCitiPassFile CitiPassFile, NumPassRecs
  PassRecLen = Len(Citipass)
  NumPassRecs = LOF(CPAdminhand) \ PassRecLen
  For cnt = 1 To NumPassRecs
    Get CPAdminhand, cnt, Citipass
    If Not Citipass.DelFlag Then
      'txt.InsertRow = Str$(cnt) & Chr$(9) & Str$(Citipass.PassNum) & Chr$(9) & QPTrim(Citipass.username)
    
        Howmany = Howmany + 1
  
        ToPrint$ = Space$(35)
        md$ = ""
        Mid$(ToPrint$, 4) = Str$(Citipass.PassNum)
        Mid$(ToPrint$, 14) = QPTrim(Citipass.username)
        'Mid$(ToPrint$, 50) = GLBank.GLAcct
        If Citipass.Administ Then
          md$ = "ADMIN** FULL ACCESS"
          mdtprint$ = mdtprint$ + md$
        Else
          For cnt2 = 1 To 10
            cnt3 = 0
            md$ = ""
            Select Case cnt2
              Case 1:
                md$ = md$ + " BL-"
              Case 2:
                md$ = md$ + " AP-"
              Case 3:
                md$ = md$ + " GL-"
              Case 4:
                md$ = md$ + " PR-"
              Case 5:
                md$ = md$ + " FA-"
'              Case 6:
'                 GoTo SKIPEM
              Case 7:
                md$ = md$ + " TX-"
              Case 8:
                md$ = md$ + " CM-"
              Case 9:
                md$ = md$ + " UB-"
              Case 10:
                md$ = md$ + " DC-"
              Case Else
                md$ = md$ + ""
            End Select
            If Citipass.Module(cnt2).FullAccess Then
              md$ = md$ + "F"
              cnt3 = cnt3 + 1
              'Mid$(ToPrint$, 50) = "Full"
            End If
            If Citipass.Module(cnt2).ReportsOnly Then
              md$ = md$ + "R"
              cnt3 = cnt3 + 1
              'Mid$(ToPrint$, 55) = "Rpt"
            End If
            If Citipass.Module(cnt2).PaymentAccess Then
              If cnt2 = 3 Then
                md$ = md$ + "C"
                cnt3 = cnt3 + 1
                'Mid$(ToPrint$, 60) = "Close"
              ElseIf cnt2 = 2 Then
                md$ = md$ + "O"
                cnt3 = cnt3 + 1
                'Mid$(ToPrint$, 60) = "PO"
              Else
                md$ = md$ + "P"
                cnt3 = cnt3 + 1
                'Mid$(ToPrint$, 60) = "Pmt"
              End If
            End If
            If Citipass.Module(cnt2).Adjustments Then
              If cnt2 = 9 Then
                md$ = md$ + "A"
                cnt3 = cnt3 + 1
              End If
            End If
            md$ = md$ + ","
           If Not cnt3 > 0 Then
            md$ = ""
           End If
           
           mdtprint$ = mdtprint$ + md$
          Next
        End If
      
        
        Print #PRNFile, ToPrint$ + mdtprint$
        mdtprint$ = ""
        Linecnt = Linecnt + 1
        If Linecnt > MaxLines Then
          Print #PRNFile, FF$
          GoSub PrintPageHeader
        End If
      End If
    
    
  Next

   Print #PRNFile,
   Print #PRNFile, Howmany; "User Codes listed."
   If grpt = False Then
    Print #PRNFile, FF$
   End If

   Close PRNFile
   If grpt = True Then
     'Load frmLoadingRpt
     ARptLineRpt.GetName ReportFile$
     ARptLineRpt.startrpt
   Else
    ViewPrint ReportFile$, "User PassCode Printout"
    Kill ReportFile$
   End If
Exit Sub

PrintPageHeader:
  Print #PRNFile,
  Print #PRNFile,
  Print #PRNFile, Header$
  Print #PRNFile,
  Print #PRNFile, "BL=Business License"
  Print #PRNFile, "AP=Accounts Payable"
  Print #PRNFile, "GL=General Ledger"
  Print #PRNFile, "PR=Payroll"
  Print #PRNFile, "FA=Fixed Assets"
  Print #PRNFile, "TX=Tax Billing"
  Print #PRNFile, "CM=Cash Management"
  Print #PRNFile, "UB=Utility Billing"
  Print #PRNFile, "DC=Vehicle Decals"
  Print #PRNFile,
  Print #PRNFile, "Codes F=Full Access, R=Reports, P=Payments, C=CloseYear, O=PO's, A-Adjustments"
  Print #PRNFile,
  Print #PRNFile, "User Number        Name                Access Rights   "
  Print #PRNFile, String$(80, "-")
  Linecnt = 9
Return

End Sub


