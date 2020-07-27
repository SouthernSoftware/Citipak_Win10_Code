VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLCheck4BlankCatDesc 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Conversion: Check For Blank Category Descriptions By Customer"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmBLCheck4BlankCatDesc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpList 
      Height          =   3300
      Left            =   1665
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3585
      Width           =   8325
      _Version        =   196608
      _ExtentX        =   14684
      _ExtentY        =   5821
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      Columns         =   3
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   0
      WrapList        =   0   'False
      WrapWidth       =   0
      SelMax          =   -1
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
      ReadOnly        =   -1  'True
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   2
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   1
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
      ColumnHeaderShow=   -1  'True
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   -2147483637
      ExtendRow       =   0
      DataField       =   ""
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      ColDesigner     =   "frmBLCheck4BlankCatDesc.frx":08CA
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "ESC EXIT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1056
      TabIndex        =   4
      Top             =   7284
      Width           =   1644
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "F10 CONTINUE/DON'T FIX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4896
      TabIndex        =   3
      Top             =   7284
      Width           =   3228
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5 PRINT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2976
      TabIndex        =   2
      Top             =   7284
      Width           =   1644
   End
   Begin VB.CommandButton cmdFix 
      Caption         =   "F6 CONTINUE/FIX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8352
      TabIndex        =   1
      Top             =   7248
      Width           =   2316
   End
   Begin EditLib.fpText fptxtChoice 
      Height          =   300
      Left            =   8964
      TabIndex        =   5
      Top             =   516
      Width           =   972
      _Version        =   196608
      _ExtentX        =   1714
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "fpText1"
      CharValidationText=   ""
      MaxLength       =   255
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   444
      Left            =   2256
      TabIndex        =   8
      Top             =   7920
      Width           =   6924
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmBLCheck4BlankCatDesc.frx":0C79
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1740
      Left            =   1104
      TabIndex        =   7
      Top             =   1728
      Width           =   9372
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NOTICE: BLANK CATEGORY DESCRIPTIONS BY CUSTOMER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   684
      Left            =   3552
      TabIndex        =   6
      Top             =   768
      Width           =   4572
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   8124
      Left            =   636
      Top             =   372
      Width           =   10380
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   924
      Left            =   3444
      Top             =   648
      Width           =   4764
   End
End
Attribute VB_Name = "frmBLCheck4BlankCatDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdContinue_Click()
  fptxtChoice.Text = "continue"
  Me.Hide
End Sub

Private Sub cmdExit_Click()
  fptxtChoice.Text = "exit"
  Me.Hide
End Sub

Private Sub PrintText()
  Dim x As Integer
  Dim RptFile$
  Dim RptHandle As Integer
  Dim DosCustRec As DosARCustRecType
  Dim DosCustHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim FF$
  Dim PntCnt As Integer
  Dim y As Integer
  Dim PrintLine$
  
  On Error Resume Next
  FF$ = Chr$(12)
  MaxLines = 57
  LineCnt = 0
  RptFile = "BlankCat.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  GoSub PrintHeader
  
  OpenDosCustFile DosCustHandle
  
  For x = 1 To BlankCatDescCnt
    PrintLine = ""
    Get DosCustHandle, BlankCatDesc(x), DosCustRec
    If LineCnt > MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
    Print #RptHandle, "Name: "; QPTrim$(DosCustRec.CustName); Tab(50); "Number: "; QPTrim$(DosCustRec.CUSTNUMB)
    Print #RptHandle, String$(80, "=")
    LineCnt = LineCnt + 2
    If QPTrim$(DosCustRec.BILLCAT1) <> "" Then
      If QPTrim$(DosCustRec.DESC1) = "" Then
        PrintLine = QPTrim$(DosCustRec.BILLCAT1)
      End If
    End If
    If QPTrim$(DosCustRec.BILLCAT2) <> "" Then
      If QPTrim$(DosCustRec.DESC2) = "" Then
        PrintLine = PrintLine + "         " + QPTrim$(DosCustRec.BILLCAT2)
      End If
    End If
    If QPTrim$(DosCustRec.BILLCAT3) <> "" Then
      If QPTrim$(DosCustRec.DESC3) = "" Then
        PrintLine = PrintLine + "           " + QPTrim$(DosCustRec.BILLCAT3)
      End If
    End If
    If QPTrim$(DosCustRec.BILLCAT4) <> "" Then
      If QPTrim$(DosCustRec.DESC4) = "" Then
        PrintLine = PrintLine + "           " + QPTrim$(DosCustRec.BILLCAT4)
      End If
    End If
    If QPTrim$(DosCustRec.BILLCAT5) <> "" Then
      If QPTrim$(DosCustRec.DESC5) = "" Then
        PrintLine = PrintLine + "           " + QPTrim$(DosCustRec.BILLCAT5)
      End If
    End If
    Print #RptHandle, PrintLine
    Print #RptHandle, String$(80, "-")
    Print #RptHandle,
    Print #RptHandle,
    PntCnt = PntCnt + 1
    LineCnt = LineCnt + 3
SkipIt:
  Next x
  Print #RptHandle, FF$
  Close
  
  If PntCnt = 0 Then
    Exit Sub
  End If
  
  ViewPrint RptFile, "Blank Category Descriptions", True
  KillFile "BlankCat.PRN"
  Exit Sub
  
PrintHeader:
  Print #RptHandle, "Blank Category Descriptions"
  Print #RptHandle, Date
  Print #RptHandle, "Business License Conversion: List of Customers with Missing Descriptions"
  Print #RptHandle, "Missing Codes"
  Print #RptHandle, " #1"; Tab(16); "#2"; Tab(32); "#3"; Tab(48); "#4"; Tab(64); "#5"
  Print #RptHandle, String$(80, "=")
  LineCnt = 5
Return
End Sub

Private Sub cmdFix_Click()
  Dim DosCustRec As DosARCustRecType
  Dim DosCustHandle As Integer
  Dim x As Integer, y As Integer
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeHandle As Integer
  Dim NumOfCodeRecs As Integer
  Dim ThisDesc$
  Dim ThisCat$
  Dim ThisCnt As Integer
  Dim FixNum As Integer
  Dim ThisPct As Integer
  
  OpenDosCustFile DosCustHandle
  OpenCatCodeFile CodeHandle
  NumOfCodeRecs = LOF(CodeHandle) / Len(CodeRec)
  
  For x = 1 To BlankCatDescCnt
    FixNum = 0
    Get DosCustHandle, BlankCatDesc(x), DosCustRec
    If QPTrim$(DosCustRec.BILLCAT1) <> "" Then
      If QPTrim$(DosCustRec.DESC1) = "" Then
        ThisCat = QPTrim$(DosCustRec.BILLCAT1)
        GoSub FixIt
        DosCustRec.DESC1 = ThisDesc
        FixNum = FixNum + 1
      End If
    End If
    If QPTrim$(DosCustRec.BILLCAT2) <> "" Then
      If QPTrim$(DosCustRec.DESC2) = "" Then
        ThisCat = QPTrim$(DosCustRec.BILLCAT2)
        GoSub FixIt
        DosCustRec.DESC2 = ThisDesc
        FixNum = FixNum + 1
      End If
    End If
    If QPTrim$(DosCustRec.BILLCAT3) <> "" Then
      If QPTrim$(DosCustRec.DESC3) = "" Then
        ThisCat = QPTrim$(DosCustRec.BILLCAT3)
        GoSub FixIt
        DosCustRec.DESC3 = ThisDesc
        FixNum = FixNum + 1
      End If
    End If
    If QPTrim$(DosCustRec.BILLCAT4) <> "" Then
      If QPTrim$(DosCustRec.DESC4) = "" Then
        ThisCat = QPTrim$(DosCustRec.BILLCAT4)
        GoSub FixIt
        DosCustRec.DESC4 = ThisDesc
        FixNum = FixNum + 1
      End If
    End If
    If QPTrim$(DosCustRec.BILLCAT5) <> "" Then
      If QPTrim$(DosCustRec.DESC5) = "" Then
        ThisCat = QPTrim$(DosCustRec.BILLCAT5)
        GoSub FixIt
        DosCustRec.DESC5 = ThisDesc
        FixNum = FixNum + 1
      End If
    End If
    
    If FixNum > 0 Then
      Put DosCustHandle, BlankCatDesc(x), DosCustRec
    End If
    ThisPct = OldRound(x / BlankCatDescCnt) * 100
    DoEvents
    Label3.Caption = "Fixing blank category descriptions is " + CStr(ThisPct) + "% completed."
    DoEvents
  Next x
  
  Close DosCustHandle
  Close CodeHandle
  
  frmBLMessageBoxJr.Label1.Caption = "A total of " + CStr(ThisCnt) + " blank category descriptions have been automatically assigned the appropriate category description."
  frmBLMessageBoxJr.Label1.Top = 700
  frmBLMessageBoxJr.Show vbModal
  fptxtChoice.Text = "continue"
  
  Me.Hide
  Exit Sub
  
FixIt:
  ThisDesc = ""
  For y = 1 To NumOfCodeRecs
    Get CodeHandle, y, CodeRec
    If QPTrim(CodeRec.CATCODE) = ThisCat Then
      ThisDesc = QPTrim$(CodeRec.CODEDESC)
      Exit For
    End If
  Next y
  ThisCnt = ThisCnt + 1
  Return

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      Call cmdContinue_Click
      KeyCode = 0
    Case vbKeyF5:
      Call cmdPrint_Click
      KeyCode = 0
    Case vbKeyF6:
      Call cmdFix_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Call LoadMe
End Sub

Private Sub LoadMe()
  Dim x As Integer
  Dim DosNumOfCustRecs As Integer
  Dim DosCustRec As DosARCustRecType
  Dim DosCustHandle As Integer
  Dim DosCustRec2 As DOSARCustRecType2
  Dim DosCustHandle2 As Integer
  
  DoEvents
  fptxtChoice.Visible = False
  
  If Version = 1 Then
    OpenDosCustFile DosCustHandle
  End If
  
  For x = 1 To BlankCatDescCnt
    Get DosCustHandle, BlankCatDesc(x), DosCustRec
    If QPTrim$(DosCustRec.BILLCAT1) <> "" Then
      If QPTrim$(DosCustRec.DESC1) = "" Then
        fpList.AddItem DosCustRec.BILLCAT1 + Chr(9) + QPTrim$(DosCustRec.CustName) + Chr(9) + QPTrim$(DosCustRec.CUSTNUMB)
      End If
    End If
    If QPTrim$(DosCustRec.BILLCAT2) <> "" Then
      If QPTrim$(DosCustRec.DESC2) = "" Then
        fpList.AddItem DosCustRec.BILLCAT2 + Chr(9) + QPTrim$(DosCustRec.CustName) + Chr(9) + QPTrim$(DosCustRec.CUSTNUMB)
      End If
    End If
    If QPTrim$(DosCustRec.BILLCAT3) <> "" Then
      If QPTrim$(DosCustRec.DESC3) = "" Then
        fpList.AddItem DosCustRec.BILLCAT3 + Chr(9) + QPTrim$(DosCustRec.CustName) + Chr(9) + QPTrim$(DosCustRec.CUSTNUMB)
      End If
    End If
    If QPTrim$(DosCustRec.BILLCAT4) <> "" Then
      If QPTrim$(DosCustRec.DESC4) = "" Then
        fpList.AddItem DosCustRec.BILLCAT4 + Chr(9) + QPTrim$(DosCustRec.CustName) + Chr(9) + QPTrim$(DosCustRec.CUSTNUMB)
      End If
    End If
    If QPTrim$(DosCustRec.BILLCAT5) <> "" Then
      If QPTrim$(DosCustRec.DESC5) = "" Then
        fpList.AddItem DosCustRec.BILLCAT5 + Chr(9) + QPTrim$(DosCustRec.CustName) + Chr(9) + QPTrim$(DosCustRec.CUSTNUMB)
      End If
    End If
  Next x

  If Version = 1 Then
    Close DosCustHandle
  End If
  
End Sub

Private Sub PrintGraphics()
  Dim x As Integer
  Dim RptFile$
  Dim RptHandle As Integer
  Dim DosCustRec As DosARCustRecType
  Dim DosCustHandle As Integer
  Dim DosCustRec2 As DOSARCustRecType2
  Dim DosCustHandle2 As Integer
  Dim dlm$
  
  On Error Resume Next
  dlm$ = "~"
  RptFile = "BLRPTS\BLNKLIC.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  If Version = 1 Then
    OpenDosCustFile DosCustHandle
  ElseIf Version = 2 Then
    OpenDosCustFile2 DosCustHandle2
  End If
  
  For x = 1 To NextBlank
    If Version = 1 Then
      Get DosCustHandle, DupBlanks(x), DosCustRec
      If QPTrim$(DosCustRec.LICENSE) = "" Then DosCustRec.LICENSE = "EMPTY"
      Print #RptHandle, DosCustRec.LICENSE; dlm; QPTrim$(DosCustRec.CustName); dlm; QPTrim$(DosCustRec.CUSTNUMB)
    ElseIf Version = 2 Then
      Get DosCustHandle2, DupBlanks(x), DosCustRec2
      If QPTrim$(DosCustRec2.LICENSE) = "" Then DosCustRec2.LICENSE = "EMPTY"
      Print #RptHandle, DosCustRec2.LICENSE; dlm; QPTrim$(DosCustRec2.CustName); dlm; QPTrim$(DosCustRec2.CUSTNUMB)
    End If
  Next x
  Close
  
  arBLCnvtChk4BlankLic.Show vbModal

End Sub

Private Sub cmdPrint_Click()
  MsgBox "Pitch 10 is recommended for this report."
  Call PrintText
End Sub




