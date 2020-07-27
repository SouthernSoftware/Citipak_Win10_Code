VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLCheck4DupCats 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Conversion: Duplicate Category Codes"
   ClientHeight    =   7935
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   8820
   Icon            =   "frmBLCheck4DupCats.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   8820
   StartUpPosition =   1  'CenterOwner
   Begin LpLib.fpList fpList 
      Height          =   2625
      Left            =   390
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3840
      Width           =   8040
      _Version        =   196608
      _ExtentX        =   14182
      _ExtentY        =   4630
      TextAlias       =   ""
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
      ColDesigner     =   "frmBLCheck4DupCats.frx":08CA
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "F10 CONTINUE "
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
      Left            =   726
      TabIndex        =   6
      Top             =   6912
      Width           =   2028
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "ESC ABORT"
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
      Left            =   6054
      TabIndex        =   2
      Top             =   6912
      Width           =   2028
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
      Left            =   3414
      TabIndex        =   1
      Top             =   6912
      Width           =   2028
   End
   Begin EditLib.fpText fptxtChoice 
      Height          =   300
      Left            =   384
      TabIndex        =   5
      Top             =   384
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmBLCheck4DupCats.frx":0C81
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   2172
      Left            =   432
      TabIndex        =   4
      Top             =   1536
      Width           =   7980
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NOTICE: DUPLICATE CATEGORY CODES DETECTED"
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
      Height          =   732
      Left            =   2124
      TabIndex        =   3
      Top             =   504
      Width           =   4428
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   7596
      Left            =   156
      Top             =   168
      Width           =   8508
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   972
      Left            =   1908
      Top             =   348
      Width           =   4764
   End
End
Attribute VB_Name = "frmBLCheck4DupCats"
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
  fptxtChoice.Text = "abort"
  Me.Hide
End Sub

Private Sub PrintText()
  Dim x As Integer
  Dim RptFile$
  Dim RptHandle As Integer
  Dim DosCodeRec As DosARNewCatCodeRecType
  Dim DosCodeHandle As Integer
  Dim DosCodeRec2 As DosARNewCatCodeRecType2
  Dim DosCodeHandle2 As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim FF$, ThisCnt$
  Dim ThisRec As Integer
  On Error Resume Next
  FF$ = Chr$(12)
  MaxLines = 57
  LineCnt = 0
  RptFile = "BLRPTS\DUPCAT.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  GoSub PrintHeader
  
  If CatVersion = 1 Then
    OpenDosCatFile DosCodeHandle
  ElseIf CatVersion = 2 Then
    OpenDosCatFile2 DosCodeHandle2
  End If
  
  For x = 1 To DupCatCnt
    If DupCats(x) < 0 Then
      ThisRec = -DupCats(x)
    Else
      ThisRec = DupCats(x)
    End If
    If CatVersion = 1 Then
      Get DosCodeHandle, ThisRec, DosCodeRec
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintHeader
      End If
      If QPTrim$(DosCodeRec.CATCODE) = "" Then DosCodeRec.CATCODE = "EMPTY"
      If DupCats(x) < 0 Then
        Print #RptHandle, Tab(8); QPTrim$(DosCodeRec.CATCODE); Tab(27); QPTrim$(DosCodeRec.CODEDESC) + " ACTIVE"; Tab(65); CStr(ThisRec)
      ElseIf DupCats(x) > 0 Then
        Print #RptHandle, Tab(8); QPTrim$(DosCodeRec.CATCODE); Tab(27); QPTrim$(DosCodeRec.CODEDESC) + " NOW INACTIVE"; Tab(65); CStr(ThisRec)
      End If
      LineCnt = LineCnt + 1
    ElseIf CatVersion = 2 Then
      Get DosCodeHandle2, ThisRec, DosCodeRec2
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintHeader
      End If
      If QPTrim$(DosCodeRec2.CATCODE) = "" Then DosCodeRec2.CATCODE = "EMPTY"
      If DupCats(x) < 0 Then
        Print #RptHandle, Tab(8); QPTrim$(DosCodeRec.CATCODE); Tab(27); QPTrim$(DosCodeRec.CODEDESC) + " ACTIVE"; Tab(65); CStr(ThisRec)
      ElseIf DupCats(x) > 0 Then
        Print #RptHandle, Tab(8); QPTrim$(DosCodeRec.CATCODE); Tab(27); QPTrim$(DosCodeRec.CODEDESC) + " NOW INACTIVE"; Tab(65); CStr(ThisRec)
      End If
      LineCnt = LineCnt + 1
    End If
SkipIt:
  Next x
  Print #RptHandle, FF$
  Close
  
  ViewPrint RptFile, "Duplicate Category Codes", True
  KillFile "DupCat.PRN"
  Exit Sub
  
PrintHeader:
  Print #RptHandle, "Report Date: "; Date$
  Print #RptHandle, "Duplicate Category Codes Report"
  Print #RptHandle, "All customers using INACTIVE codes are now switched to the ACTIVE codes"
  Print #RptHandle,
  Print #RptHandle, Tab(3); "Category Codes"; Tab(27); "Description"; Tab(62); "Rec Num"
  Print #RptHandle, String$(80, "=")
  LineCnt = 6
  
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
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Call LoadMe
End Sub

Private Sub LoadMe()
  Dim x As Integer
  Dim DosCodeRec As DosARNewCatCodeRecType
  Dim DosCodeHandle As Integer
  Dim DosCodeRec2 As DosARNewCatCodeRecType2
  Dim DosCodeHandle2 As Integer
  Dim ThisRec As Integer
  
  fptxtChoice.Visible = False
  If CatVersion = 1 Then
    OpenDosCatFile DosCodeHandle
  ElseIf CatVersion = 2 Then
    OpenDosCatFile2 DosCodeHandle2
  End If
  
  For x = 1 To DupCatCnt
    If DupCats(x) < 0 Then
      ThisRec = -DupCats(x)
    Else
      ThisRec = DupCats(x)
    End If
    If CatVersion = 1 Then
      Get DosCodeHandle, ThisRec, DosCodeRec
      fpList.AddItem QPTrim$(DosCodeRec.CATCODE) + Chr(9) + QPTrim$(DosCodeRec.CODEDESC) + Chr(9) + CStr(ThisRec)
    ElseIf CatVersion = 2 Then
      Get DosCodeHandle2, ThisRec, DosCodeRec2
      fpList.AddItem QPTrim$(DosCodeRec2.CATCODE) + Chr(9) + QPTrim$(DosCodeRec2.CODEDESC) + Chr(9) + CStr(ThisRec)
    End If
  Next x
  
  Close
End Sub

Private Sub PrintGraphics()
  Dim x As Integer
  Dim RptFile$
  Dim RptHandle As Integer
  Dim DosCodeRec As DosARNewCatCodeRecType
  Dim DosCodeHandle As Integer
  Dim DosCodeRec2 As DosARNewCatCodeRecType2
  Dim DosCodeHandle2 As Integer
  Dim dlm$
  Dim ThisRec As Integer
  
  On Error Resume Next
  dlm$ = "~"
  RptFile = "BLRPTS\DUPCAT.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  If CatVersion = 1 Then
    OpenDosCatFile DosCodeHandle
  ElseIf CatVersion = 2 Then
    OpenDosCatFile2 DosCodeHandle2
  End If
  
  For x = 1 To DupCatCnt
    If CatVersion = 1 Then
      If DupCats(x) < 0 Then
        ThisRec = -DupCats(x)
        Get DosCodeHandle, ThisRec, DosCodeRec
        Print #RptHandle, QPTrim$(DosCodeRec.CATCODE); dlm; QPTrim$(DosCodeRec.CODEDESC) + " ACTIVE"; dlm; CStr(ThisRec)
      ElseIf DupCats(x) > 0 Then
        ThisRec = DupCats(x)
        Get DosCodeHandle, ThisRec, DosCodeRec
        Print #RptHandle, QPTrim$(DosCodeRec.CATCODE); dlm; QPTrim$(DosCodeRec.CODEDESC) + " NOW INACTIVE"; dlm; CStr(ThisRec)
      End If
    ElseIf CatVersion = 2 Then
      If DupCats(x) < 0 Then
        ThisRec = -DupCats(x)
        Get DosCodeHandle2, ThisRec, DosCodeRec2
        Print #RptHandle, QPTrim$(DosCodeRec2.CATCODE); dlm; QPTrim$(DosCodeRec2.CODEDESC) + " ACTIVE"; dlm; CStr(ThisRec)
      ElseIf DupCats(x) > 0 Then
        Get DosCodeHandle2, ThisRec, DosCodeRec2
        Print #RptHandle, QPTrim$(DosCodeRec2.CATCODE); dlm; QPTrim$(DosCodeRec2.CODEDESC) + " NOW INACTIVE"; dlm; CStr(ThisRec)
      End If
    End If
  Next x
  Close
  
  arBLCnvtDupCatNums.Show vbModal

End Sub

Private Sub cmdPrint_Click()
  Dim PrintType$
  
  frmBLReportOpt.Show vbModal 'opens small screen from which the
  'user selects the printing method
  PrintType$ = frmBLReportOpt.fptxtPrintType
  Select Case PrintType$
    Case "Graphical"
      Call PrintGraphics
    Case "Text"
      MsgBox "Pitch 10 is recommended for this report."
      Call PrintText
    Case "Exit"
  End Select
  Unload frmBLReportOpt

End Sub


