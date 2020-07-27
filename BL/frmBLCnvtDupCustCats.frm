VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLCnvtDupCustCats 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Conversion: Customers with Duplicate Category Numbers"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmBLCnvtDupCustCats.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpList 
      Height          =   3030
      Left            =   1650
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4155
      Width           =   8325
      _Version        =   196608
      _ExtentX        =   14684
      _ExtentY        =   5345
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
      ColDesigner     =   "frmBLCnvtDupCustCats.frx":08CA
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
      Left            =   4571
      TabIndex        =   3
      Top             =   7647
      Width           =   2505
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "F10 CONTINUE/FIX"
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
      Left            =   7181
      TabIndex        =   2
      Top             =   7647
      Width           =   2505
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
      Left            =   1955
      TabIndex        =   1
      Top             =   7647
      Width           =   2505
   End
   Begin EditLib.fpText fptxtChoice 
      Height          =   300
      Left            =   8958
      TabIndex        =   4
      Top             =   447
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
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   1212
      Left            =   3318
      Top             =   579
      Width           =   4764
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   8124
      Left            =   1566
      Top             =   303
      Width           =   8508
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NOTICE: CUSTOMERS WITH DUPLICATE CATEGORY CODES"
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
      Left            =   3534
      TabIndex        =   6
      Top             =   831
      Width           =   4428
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmBLCnvtDupCustCats.frx":0C0E
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
      Height          =   2100
      Left            =   2880
      TabIndex        =   5
      Top             =   1950
      Width           =   5865
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmBLCnvtDupCustCats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdContinue_Click()
  Dim x As Integer, y As Integer
  Dim DosCustRec As DosARCustRecType
  Dim DosCustHandle As Integer
  Dim Nextx As Integer
  Dim DupCnt As Integer
  Dim HoldCode As String
  
  OpenDosCustFile DosCustHandle
  ReDim ThisCode(1 To 1) As String

  For x = 1 To DCCCnt 'DCCnt = # of customers with duplicate
  'category codes
    DupCnt = 0
    Get DosCustHandle, DupCustCats(x), DosCustRec 'pull up each affected customer
      If QPTrim$(DosCustRec.BILLCAT1) <> "" Then 'assign array with non-empty codes
        DupCnt = DupCnt + 1
        ReDim Preserve ThisCode(1 To DupCnt) As String
        ThisCode(DupCnt) = QPTrim$(DosCustRec.BILLCAT1)
      End If
      If QPTrim$(DosCustRec.BILLCAT2) <> "" Then
        DupCnt = DupCnt + 1
        ReDim Preserve ThisCode(1 To DupCnt) As String
        ThisCode(DupCnt) = QPTrim$(DosCustRec.BILLCAT2)
      End If
      If QPTrim$(DosCustRec.BILLCAT3) <> "" Then
        DupCnt = DupCnt + 1
        ReDim Preserve ThisCode(1 To DupCnt) As String
        ThisCode(DupCnt) = QPTrim$(DosCustRec.BILLCAT3)
      End If
      If QPTrim$(DosCustRec.BILLCAT4) <> "" Then
        DupCnt = DupCnt + 1
        ReDim Preserve ThisCode(1 To DupCnt) As String
        ThisCode(DupCnt) = QPTrim$(DosCustRec.BILLCAT4)
      End If
      If QPTrim$(DosCustRec.BILLCAT5) <> "" Then
        DupCnt = DupCnt + 1
        ReDim Preserve ThisCode(1 To DupCnt) As String
        ThisCode(DupCnt) = QPTrim$(DosCustRec.BILLCAT5)
      End If
      Nextx = 1
      Do 'loop through codes for this customer one at a time
      'comparing each against the rest looking for the duplicates
        If Nextx > DupCnt Then Exit Do
        If ThisCode(Nextx) = "" Then GoTo EmptyString
        HoldCode = ThisCode(Nextx)
        For y = 1 To DupCnt
          If y = Nextx Then GoTo SkipThisOne
          If ThisCode(y) = HoldCode Then 'found one
            ThisCode(y) = ""
            Select Case y 'don't care about replacing the first code
              Case 2: 'look for the number of this duplicate and delete
                DosCustRec.BILLCAT2 = ""
                DosCustRec.DESC2 = ""
                Put DosCustHandle, DupCustCats(x), DosCustRec
              Case 3:
                DosCustRec.BILLCAT3 = ""
                DosCustRec.DESC3 = ""
                Put DosCustHandle, DupCustCats(x), DosCustRec
              Case 4:
                DosCustRec.BILLCAT4 = ""
                DosCustRec.DESC4 = ""
                Put DosCustHandle, DupCustCats(x), DosCustRec
              Case 5:
                DosCustRec.BILLCAT5 = ""
                DosCustRec.DESC5 = ""
                Put DosCustHandle, DupCustCats(x), DosCustRec
             End Select
          End If
SkipThisOne:
        Next y 'now compare the next number with the control code (HoldCode)
EmptyString:
        Nextx = Nextx + 1 'been through all codes so update the control code
      Loop
  Next x 'move to next duplicate customer
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
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim FF$
  Dim PntCnt As Integer
  
  On Error Resume Next
  FF$ = Chr$(12)
  MaxLines = 57
  LineCnt = 0
  RptFile = "DupCustCats.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  GoSub PrintHeader
  
  OpenDosCustFile DosCustHandle
  
  For x = 1 To DCCCnt
    Get DosCustHandle, DupCustCats(x), DosCustRec
    If LineCnt > MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
    Print #RptHandle, Tab(8); QPTrim$(DosCustRec.CUSTNUMB); Tab(27); QPTrim$(DosCustRec.CustName); Tab(68); QPTrim$(DosCustRec.BILLCAT1)
    Print #RptHandle, Tab(8); QPTrim$(DosCustRec.CUSTNUMB); Tab(27); QPTrim$(DosCustRec.CustName); Tab(68); QPTrim$(DosCustRec.BILLCAT2)
    Print #RptHandle, Tab(8); QPTrim$(DosCustRec.CUSTNUMB); Tab(27); QPTrim$(DosCustRec.CustName); Tab(68); QPTrim$(DosCustRec.BILLCAT3)
    Print #RptHandle, Tab(8); QPTrim$(DosCustRec.CUSTNUMB); Tab(27); QPTrim$(DosCustRec.CustName); Tab(68); QPTrim$(DosCustRec.BILLCAT4)
    Print #RptHandle, Tab(8); QPTrim$(DosCustRec.CUSTNUMB); Tab(27); QPTrim$(DosCustRec.CustName); Tab(68); QPTrim$(DosCustRec.BILLCAT5)
    PntCnt = PntCnt + 1
    LineCnt = LineCnt + 5
SkipIt:
  Next x
  Print #RptHandle, FF$
  Close
  
  If PntCnt = 0 Then
    Exit Sub
  End If
  
  ViewPrint RptFile, "Duplicate Customer Category Codes", True
  KillFile "DupCustCats.PRN"
  Exit Sub
  
PrintHeader:
  Print #RptHandle, "Duplicate Customer Category Codes Report"
  Print #RptHandle, Date
  Print #RptHandle,
  Print #RptHandle, Tab(3); "Cust Number"; Tab(27); "Business Name"; Tab(70); "Cust Codes"
  Print #RptHandle, String$(80, "=")
  LineCnt = 5
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
  Dim DosNumOfCustRecs As Integer
  Dim DosCustRec As DosARCustRecType
  Dim DosCustHandle As Integer
  
  DoEvents
  fptxtChoice.Visible = False
  
  OpenDosCustFile DosCustHandle
  For x = 1 To DCCCnt
    Get DosCustHandle, DupCustCats(x), DosCustRec
    fpList.AddItem QPTrim$(DosCustRec.CUSTNUMB) + Chr(9) + QPTrim$(DosCustRec.CustName) + Chr(9) + QPTrim$(DosCustRec.BILLCAT1)
    fpList.AddItem QPTrim$(DosCustRec.CUSTNUMB) + Chr(9) + QPTrim$(DosCustRec.CustName) + Chr(9) + QPTrim$(DosCustRec.BILLCAT2)
    fpList.AddItem QPTrim$(DosCustRec.CUSTNUMB) + Chr(9) + QPTrim$(DosCustRec.CustName) + Chr(9) + QPTrim$(DosCustRec.BILLCAT3)
    fpList.AddItem QPTrim$(DosCustRec.CUSTNUMB) + Chr(9) + QPTrim$(DosCustRec.CustName) + Chr(9) + QPTrim$(DosCustRec.BILLCAT4)
    fpList.AddItem QPTrim$(DosCustRec.CUSTNUMB) + Chr(9) + QPTrim$(DosCustRec.CustName) + Chr(9) + QPTrim$(DosCustRec.BILLCAT5)
SkipIt:
  Next x

  Close DosCustHandle

End Sub

Private Sub PrintGraphics()
  Dim x As Integer
  Dim RptFile$
  Dim RptHandle As Integer
  Dim DosCustRec As DosARCustRecType
  Dim DosCustHandle As Integer
  Dim dlm$
  
  On Error Resume Next
  dlm$ = "~"
  RptFile = "BLRPTS\DupCustCats.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  OpenDosCustFile DosCustHandle
  For x = 1 To DCCCnt
    Get DosCustHandle, DupCustCats(x), DosCustRec
    Print #RptHandle, QPTrim$(DosCustRec.CUSTNUMB); dlm; QPTrim$(DosCustRec.CustName); dlm; QPTrim$(DosCustRec.BILLCAT1)
    Print #RptHandle, QPTrim$(DosCustRec.CUSTNUMB); dlm; QPTrim$(DosCustRec.CustName); dlm; QPTrim$(DosCustRec.BILLCAT2)
    Print #RptHandle, QPTrim$(DosCustRec.CUSTNUMB); dlm; QPTrim$(DosCustRec.CustName); dlm; QPTrim$(DosCustRec.BILLCAT3)
    Print #RptHandle, QPTrim$(DosCustRec.CUSTNUMB); dlm; QPTrim$(DosCustRec.CustName); dlm; QPTrim$(DosCustRec.BILLCAT4)
    Print #RptHandle, QPTrim$(DosCustRec.CUSTNUMB); dlm; QPTrim$(DosCustRec.CustName); dlm; QPTrim$(DosCustRec.BILLCAT5)
  Next x
  Close
  
  arBLCnvtDupCustCats.Show vbModal

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



