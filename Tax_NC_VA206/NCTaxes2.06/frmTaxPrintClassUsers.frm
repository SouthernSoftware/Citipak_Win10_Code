VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmTaxPrintClassUsers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Real Classifications Currently In Use"
   ClientHeight    =   6405
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   10605
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   10605
   StartUpPosition =   1  'CenterOwner
   Begin LpLib.fpList fpList1 
      Height          =   2352
      Left            =   600
      TabIndex        =   0
      Top             =   2772
      Width           =   9480
      _Version        =   196608
      _ExtentX        =   16722
      _ExtentY        =   4149
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
      ReadOnly        =   0   'False
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
      ColumnHeaderShow=   0   'False
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
      ColDesigner     =   "frmTaxPrintClassUsers.frx":0000
   End
   Begin EditLib.fpText fptxtAnswer 
      Height          =   132
      Left            =   204
      TabIndex        =   1
      Top             =   6000
      Visible         =   0   'False
      Width           =   372
      _Version        =   196608
      _ExtentX        =   661
      _ExtentY        =   238
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
      Text            =   ""
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
   Begin fpBtnAtlLibCtl.fpBtn cmdConvert 
      Height          =   420
      Left            =   4530
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1575
      _Version        =   131072
      _ExtentX        =   2778
      _ExtentY        =   741
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmTaxPrintClassUsers.frx":0384
   End
   Begin EditLib.fpText fptxtRealClass 
      Height          =   372
      Left            =   4290
      TabIndex        =   3
      Top             =   120
      Width           =   3612
      _Version        =   196608
      _ExtentX        =   6371
      _ExtentY        =   656
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
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
      AlignTextH      =   1
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
      Text            =   ""
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
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   420
      Left            =   2460
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1572
      _Version        =   131072
      _ExtentX        =   2773
      _ExtentY        =   741
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmTaxPrintClassUsers.frx":0562
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrint 
      Height          =   420
      Left            =   6555
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5865
      Width           =   1575
      _Version        =   131072
      _ExtentX        =   2778
      _ExtentY        =   741
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmTaxPrintClassUsers.frx":0740
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Real Pin#"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   7560
      TabIndex        =   11
      Top             =   2460
      Width           =   2052
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "If you wish to convert this list of customers to the new classification then press F5."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1170
      TabIndex        =   10
      Top             =   1920
      Width           =   8292
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   3372
      Left            =   324
      Top             =   2400
      Width           =   9972
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Real Classification"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1890
      TabIndex        =   9
      Top             =   240
      Width           =   2292
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmTaxPrintClassUsers.frx":091D
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   1134
      TabIndex        =   8
      Top             =   720
      Width           =   8292
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3804
      TabIndex        =   7
      Top             =   2460
      Width           =   2052
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Acct Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   684
      TabIndex        =   6
      Top             =   2460
      Width           =   1572
   End
End
Attribute VB_Name = "frmTaxPrintClassUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim ThisClassCnt As Long

Private Sub cmdConvert_Click()
  Dim TaxCustRec As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim RealRec As PropertyRecType
  Dim RRHandle As Integer
  Dim NumOfRRREcs As Long
  Dim NextRec As Long
  Dim CnvtCnt As Long
  Dim ThisRec As Long
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxPropFile RRHandle, NumOfRRREcs
  For x = 0 To fpList1.ListCount - 1
    fpList1.Col = 0
    fpList1.Row = x
    fpList1.ListIndex = x
    ThisRec = CLng(fpList1.ColText)
    If ThisRec > 0 Then
      Get TCHandle, ThisRec, TaxCustRec
        NextRec = TaxCustRec.FirstPropRec
        Do While NextRec > 0
          Get RRHandle, NextRec, RealRec
            If QPTrim$(RealRec.ICPDesc) = QPTrim$(fptxtRealClass.Text) Then
              RealRec.ICPDesc = QPTrim$(frmTaxRealClassSetup.NewDesc)
              CnvtCnt = CnvtCnt + 1
              Put RRHandle, NextRec, RealRec
            End If
            NextRec = RealRec.NextRec
        Loop
    End If
  Next x
  Close RRHandle
  Close TCHandle
  
  If CnvtCnt > 0 Then
    Call TaxMsg(900, "A total of " + CStr(CnvtCnt) + " customers have been updated to the new real classification description.")
    fptxtAnswer = "convert"
  Else
    Call TaxMsg(900, "ERROR: No real properties were converted. Please call Southern Software @ 1-800-842-8190 for assitance.")
  End If
  
  Me.Hide
End Sub

Private Sub cmdExit_Click()
  Me.Hide
End Sub

Private Sub cmdPrint_Click()
  frmTaxReportOpt.Show vbModal
  If frmTaxReportOpt.fptxtPrintType.Text = "Graphical" Then
    Unload frmTaxReportOpt
    Call PrintGraphics
  ElseIf frmTaxReportOpt.fptxtPrintType.Text = "Text" Then
    frmTaxMsg.Label1.Caption = "Pitch 10 is recommended for this report."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Unload frmTaxReportOpt
    Call PrintText
  End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdPrint_Click
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%v"
      Call cmdConvert_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Dim x As Long
  ThisClassCnt = frmTaxRealClassSetup.ThisClassCnt
  fptxtAnswer.Text = "none"
  fptxtRealClass.Text = QPTrim$(frmTaxRealClassSetup.ThisDesc)
  For x = 1 To ThisClassCnt
    fpList1.InsertRow = CStr(ClassUsersAcct(x)) + Chr(9) + QPTrim$(ClassUsersName(x)) + Chr(9) + QPTrim$(ClassRealPin(x))
  Next x

End Sub

Private Sub PrintGraphics()
  Dim x As Long
  Dim TaxCustRec As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim dlm$
  Dim RptFile$
  Dim RptHandle As Integer
  Dim Desc$
  Dim Town$
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim RealRec As PropertyRecType
  Dim RRHandle As Integer
  Dim NumOfRRREcs As Long
  Dim NextRec As Long
  Dim ThisRec As Long
  
'  'on error goto ERRORSTUFF
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Town = QPTrim$(TaxMasterRec.City)
  Desc$ = QPTrim$(fptxtRealClass.Text)
  dlm = "~"
  RptFile$ = "TAXRPTS\TXClass.RPT"
  
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxPropFile RRHandle, NumOfRRREcs
  For x = 0 To fpList1.ListCount - 1
    fpList1.Col = 0
    fpList1.Row = x
    fpList1.ListIndex = x
    ThisRec = fpList1.ColText
    If ThisRec > 0 Then
      Get TCHandle, ThisRec, TaxCustRec
        NextRec = TaxCustRec.FirstPropRec
        If NextRec > 0 Then
          Do While NextRec > 0
            Get RRHandle, NextRec, RealRec
            If QPTrim$(RealRec.ICPDesc) = Desc$ Then
              '                   0          1                  2                      3                         4
              Print #RptHandle, Town; dlm; Desc$; dlm; TaxCustRec.Acct; dlm; TaxCustRec.CustName; dlm; QPTrim$(RealRec.RealPin)
            End If
            NextRec = RealRec.NextRec
          Loop
        End If
    End If
  Next x
  Close TCHandle
  Close RRHandle
  Close RptHandle
  arTaxClass.Show vbModal
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPrintClassUsers", "PrintGraphics", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate

End Sub

Private Sub PrintText()
  Dim x As Long
  Dim TaxCustRec As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim Desc$
  Dim Town$
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim FF$, Page As Integer
  Dim Line$
  Dim RealRec As PropertyRecType
  Dim RRHandle As Integer
  Dim NumOfRRREcs As Long
  Dim ThisRec As Long
  Dim NextRec As Long
  
'  'on error goto ERRORSTUFF
  
  Line$ = String(80, "-")
  FF$ = Chr$(12)
  MaxLines = 58
  LineCnt = 0
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Town = QPTrim$(TaxMasterRec.City)
  Desc$ = QPTrim$(fptxtRealClass.Text)
  RptFile$ = "TAXRPTS\TXCounty.PRN"
  
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  GoSub PrintHeader
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxPropFile RRHandle, NumOfRRREcs
  For x = 0 To fpList1.ListCount - 1
    fpList1.Col = 0
    fpList1.Row = x
    fpList1.ListIndex = x
    ThisRec = CLng(fpList1.ColText)
    If ThisRec > 0 Then
      Get TCHandle, ThisRec, TaxCustRec
      NextRec = TaxCustRec.FirstPropRec
      Do While NextRec > 0
        Get RRHandle, NextRec, RealRec
        If QPTrim$(RealRec.ICPDesc) = Desc$ Then
          Print #RptHandle, Tab(5); CStr(TaxCustRec.Acct); Tab(20); QPTrim$(TaxCustRec.CustName); Tab(70); QPTrim$(RealRec.RealPin)
          LineCnt = LineCnt + 1
          If LineCnt >= MaxLines Then
            Print #RptHandle, FF$
            GoSub PrintHeader
          End If
        End If
        NextRec = RealRec.NextRec
      Loop
    End If
  Next x
  Print #RptHandle, FF$
  
  Close TCHandle
  Close RptHandle

  ViewPrint RptFile$, "Current Tax Real Classification Customers", True
  
  KillFile RptFile$
  
  Exit Sub
  
PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(22); "Current Tax Real Classification Customers"
  Print #RptHandle, "For Classification: " + QPTrim(fptxtRealClass.Text); Tab(65); "Page #" + CStr(Page)
  Print #RptHandle, "For: " + Town
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, Tab(2); "Account #"; Tab(20); "Customer Name"; Tab(70); "Real Pin#"
  Print #RptHandle, Line
  LineCnt = 6
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPrintClassUsers", "PrintText", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate

End Sub


