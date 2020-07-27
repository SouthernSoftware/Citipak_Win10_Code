VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmRptRateList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rate Table List Report"
   ClientHeight    =   8640
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmRptRateList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   348
      Left            =   6360
      TabIndex        =   0
      Top             =   3888
      Width           =   1908
      _Version        =   196608
      _ExtentX        =   3365
      _ExtentY        =   614
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   1
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   1
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483627
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
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
      ColumnEdit      =   0
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   3504
      EditHeight      =   -1
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   -2147483637
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   -2147483637
      ExtendRow       =   0
      ListPosition    =   0
      ButtonThreeDAppearance=   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      AutoSearchFill  =   0   'False
      AutoSearchFillDelay=   500
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmRptRateList.frx":08CA
   End
   Begin VB.CommandButton cmdExit 
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
      Height          =   492
      Left            =   10080
      TabIndex        =   2
      Top             =   7296
      Width           =   1332
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F10 &Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   8400
      TabIndex        =   1
      Top             =   7296
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   8280
      Width           =   12192
      _ExtentX        =   21505
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7133
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "9:36 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "2/17/2004"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Report Type: "
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
      Left            =   3792
      TabIndex        =   5
      Top             =   3936
      Width           =   2388
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   3540
      Left            =   2790
      Top             =   2400
      Width           =   6612
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3210
      Top             =   984
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Rate Table List Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3618
      TabIndex        =   4
      Top             =   1224
      Width           =   5004
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3210
      Top             =   864
      Width           =   5772
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuPrnScn 
         Caption         =   "Prin&t Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmRptRateList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Private Sub cmdExit_Click()
  Load frmUBSetupMenu
  DoEvents
  frmUBSetupMenu.Show
  Unload frmRptBillPayTax
  DoEvents
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
Private Sub fpcboRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptType.ListDown = True
  End If
  If fpcboRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      cmdPrint.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        'fpcboDetail.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via RptRateList by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub

Private Sub cmdPrint_Click()
  
  DeActivateControls Me, True
  Select Case fpcboRptType.ListIndex
  Case 0:
    PrintRateListing True   'Graphics Version
  Case 1:
    PrintRateListing False  'Plain old text version
    ActivateControls Me, True
  Case Else:
    ActivateControls Me, True
  End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
'    Case vbKeyDown, vbKeyReturn:
'      SendKeys "{Tab}"
'      KeyCode = 0
'    Case vbKeyUp:
'      SendKeys "+{Tab}"
'      KeyCode = 0
    Case vbKeyEscape:
      KeyCode = 0
      DoEvents
      Call cmdExit_Click
    Case vbKeyF10:
      KeyCode = 0
      DoEvents
      Call cmdPrint_Click
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
End Sub

Private Sub Form_Resize()
  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
  End If
End Sub

Private Sub PrintRateListing(graphicflag As Boolean)
  
  
  Dim UBRateTblRecLen As Integer, NumRateRecs As Integer
  Dim NumPrinted As Integer
  Dim RCnt As Integer, cnt As Integer
  Dim UBFile As Integer, RPTFile As Integer
  Dim ReportFile As String
  ReDim UBRateTblRec(1) As UBRateTblRecType
  ReDim StepText(1 To 10) As String * 40
  Dim Dash80 As String * 78
    
  UBRateTblRecLen = Len(UBRateTblRec(1))
  NumRateRecs = FileSize(UBPath + "UBRATE.DAT") \ UBRateTblRecLen

  If NumRateRecs = 0 Then
    GoTo ExitRateListing
  End If
  
  If graphicflag Then
    GoTo GraphicRateList
  End If
  
  Dash80$ = String$(78, "-")
  
  NumPrinted = 0

  FrmShowPctComp.Label1 = "Creating Bill/Payment Tax Report."
  FrmShowPctComp.Show , Me

  ReportFile$ = UBPath + "RATELIST.RPT"
  
  UBFile = FreeFile
  Open UBPath + "UBRATE.DAT" For Random Shared As UBFile Len = UBRateTblRecLen
  
  RPTFile = FreeFile
  Open ReportFile$ For Output As RPTFile
  
  GoSub PrintRateHeader
  For cnt = 1 To NumRateRecs
    Get UBFile, cnt, UBRateTblRec(1)
    If NumPrinted = 3 Then
      Print #RPTFile, Dash80$
      Print #RPTFile, Chr$(12)
      GoSub PrintRateHeader
    End If
    Print #RPTFile, "       Rate Code:  "; UBRateTblRec(1).RATECODE
    Print #RPTFile, "     Description:  "; UBRateTblRec(1).RATEDESC
    Print #RPTFile, "  Minimum Charge:"; Using$("#######.##", Str$(UBRateTblRec(1).MINAMT))
    Print #RPTFile, "   Minimum Units:"; Using$("##########", Str$(UBRateTblRec(1).MINUNITS))
    Print #RPTFile, "      Max Amount:"; Using$("######.##", Str$(UBRateTblRec(1).MaxAmt))
    Print #RPTFile, "      [ Step ]        [ Beg Unit ]     [ Amount/Unit ]"
    For RCnt = 1 To 10
      LSet StepText$(RCnt) = ""
      If UBRateTblRec(1).TblBreaks(RCnt).UNITS >= 0 Then
        Mid$(StepText$(RCnt), 8) = Using$("########", Str$(UBRateTblRec(1).TblBreaks(RCnt).UNITS))
      End If
      If UBRateTblRec(1).TblBreaks(RCnt).UNITAMT >= 0 Then
        Mid$(StepText$(RCnt), 25) = Using$("####.######", Str$(UBRateTblRec(1).TblBreaks(RCnt).UNITAMT))
      End If
    Next
    Print #RPTFile, "     First Break:"; StepText$(1)
    Print #RPTFile, "    Second Break:"; StepText$(2)
    Print #RPTFile, "     Third Break:"; StepText$(3)
    Print #RPTFile, "    Fourth Break:"; StepText$(4)
    Print #RPTFile, "     Fifth Break:"; StepText$(5)
    Print #RPTFile, "     Sixth Break:"; StepText$(6)
    Print #RPTFile, "   Seventh Break:"; StepText$(7)
    Print #RPTFile, "    Eighth Break:"; StepText$(8)
    Print #RPTFile, "     Ninth Break:"; StepText$(9)
    Print #RPTFile, "        All Over:"; StepText$(10)
    Print #RPTFile,
    NumPrinted = NumPrinted + 1
    FrmShowPctComp.ShowPctComp cnt, NumRateRecs

  Next
  Print #RPTFile, Dash80$
  Print #RPTFile, Chr$(12)
  Close

  Erase UBRateTblRec, StepText
  DoEvents
  ViewPrint ReportFile$, "Rate Table List Report"
'  PrintRptFile "Rate Code Listing Report.", "RATELIST.RPT", 1, RetCode%, 1
  KillFile "RATELIST.RPT"
  
  GoTo ExitRateListing

PrintRateHeader:
  PageNo = PageNo + 1
  Print #RPTFile, "Utility Billing Rate Table Listing."
  Print #RPTFile, TOWNNAME$; Tab(70); "Page:"; PageNo
  Print #RPTFile, "Report Date: "; Date$
  Print #RPTFile, Dash80$
  NumPrinted = 0
Return

GraphicRateList:

ExitRateListing:

End Sub


