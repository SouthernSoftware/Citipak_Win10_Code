VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmReprint 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RDF"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmRDFs.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpListRpts 
      Height          =   3150
      Left            =   2070
      TabIndex        =   0
      Top             =   2730
      Width           =   7590
      _Version        =   196608
      _ExtentX        =   13388
      _ExtentY        =   5556
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
      Columns         =   2
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
      ColDesigner     =   "frmRDFs.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   585
      Left            =   2835
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Press ESC to exit this screen."
      Top             =   6435
      Width           =   2790
      _Version        =   131072
      _ExtentX        =   4921
      _ExtentY        =   1032
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
      ButtonDesigner  =   "frmRDFs.frx":0BAE
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrint 
      Height          =   585
      Left            =   6192
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Press ESC to exit this screen."
      Top             =   6432
      Width           =   2790
      _Version        =   131072
      _ExtentX        =   4921
      _ExtentY        =   1032
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
      ButtonDesigner  =   "frmRDFs.frx":0D8C
   End
   Begin EditLib.fpText fpText1 
      Height          =   612
      Left            =   720
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   -1320
      Width           =   4212
      _Version        =   196608
      _ExtentX        =   7429
      _ExtentY        =   1080
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      ForeColor       =   8454143
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483637
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
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
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   1
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
      Text            =   "Employee Listing Report"
      CharValidationText=   ""
      MaxLength       =   255
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fpText2 
      Height          =   732
      Left            =   1824
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   576
      Width           =   5412
      _Version        =   196608
      _ExtentX        =   9546
      _ExtentY        =   1291
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   13684944
      ForeColor       =   65535
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483637
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   1
      BorderColor     =   -2147483643
      BorderWidth     =   3
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
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   1
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
      Text            =   " Reports Available for Reprint"
      CharValidationText=   ""
      MaxLength       =   255
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483643
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   6972
      Left            =   1164
      Top             =   948
      Width           =   9324
   End
End
Attribute VB_Name = "frmReprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim DedRec(1 To 50) As DedCodeRecType
  Dim DHandle As Integer

Private Sub cmdExit_Click()
  frmReportsProcessing.Show
  DoEvents
  Unload frmReprint
End Sub

Private Sub cmdPrint_Click()
  Call fpListRpts_DblClick

End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  LoadThisForm
  Me.HelpContextID = hlpReprintReports
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub LoadThisForm()
  Dim x As Integer
  Dim FoundOne As Boolean
  
  FoundOne = False
  OpenDedCodeFile DHandle
  For x = 1 To 50
    Get DHandle, x, DedRec(x)
  Next x
  Close DHandle
  
  'All FoundOne code inserted on 5/27/04
  If Exist(StartPath & "\PRRPTS\PRGLIFG.RPT") Then
    fpListRpts.AddItem "  GL Register" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\PRGLIFG.RPT")
    FoundOne = True
  End If
    
  If Exist(StartPath & "\PRRPTS\PRGLIFNSG.RPT") Then
    fpListRpts.AddItem "  GL Register Non-Split" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\PRGLIFNSG.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\DISTRIBUACCTNUMG.RPT") Then
    fpListRpts.AddItem "  Earnings Distribution Register" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\DISTRIBUACCTNUMG.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\DISTRIBUNSG.RPT") Then
    fpListRpts.AddItem "  Earnings Distribution Register non-Split" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\DISTRIBUNSG.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\REGISTERG.RPT") Then
    fpListRpts.AddItem "  Earnings Register" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\REGISTERG.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\REGISTERNSG.RPT") Then
    fpListRpts.AddItem "  Earnings Register Non-Split" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\REGISTERNSG.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\YTDWAGEG.RPT") Then
    fpListRpts.AddItem "  YTD Wage Report" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\YTDWAGEG.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\EMPrintTermEmpListG.RPT") Then
    fpListRpts.AddItem "  Terminated Employee Report" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\EMPrintTermEmpListG.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\401KG.RPT") Then
    fpListRpts.AddItem "  Supplemental Retirement Report" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\401KG.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\SeppContG.RPT") Then
    fpListRpts.AddItem "  SEPP Contribution Report" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\SeppContG.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\RETIREG.RPT") Then
    fpListRpts.AddItem "  NC State Retirement Report" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\RETIREG.RPT") & Chr(9) & FileDateTime(StartPath & "\PRRPTS\RETIREG.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\SCRETIREG.RPT") Then
    fpListRpts.AddItem "  SC State Retirement Report" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\SCRETIREG.RPT") & Chr(9) & FileDateTime(StartPath & "\PRRPTS\SCRETIREG.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\VARETIREG.RPT") Then
    fpListRpts.AddItem "  VA State Retirement Report" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\VARETIREG.RPT") & Chr(9) & FileDateTime(StartPath & "\PRRPTS\VARETIREG.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\EMPRINTEMPLISTG.RPT") Then
    fpListRpts.AddItem "  Employee List Report" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\EMPRINTEMPLISTG.RPT")
    FoundOne = True
  End If
  
  For x = 1 To 50
    If Exist(StartPath & "\PRRPTS\DEDUCTG" & x & ".RPT") Then
      fpListRpts.AddItem "  Deduction " & DedRec(x).DCDESC1 & " Report" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\DEDUCTG" & x & ".RPT")
      FoundOne = True
    End If
  Next x
  
  If Exist(StartPath & "\PRRPTS\DEDUCALL.RPT") Then 'added 9/4/03
    fpListRpts.AddItem "  Deduction Report ALL" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\DEDUCALL.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\BENEACCRG.RPT") Then
    fpListRpts.AddItem "  Benefit Accrual Report" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\BENEACCRG.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\GROSWAGEG.RPT") Then
    fpListRpts.AddItem "  Gross Wage Report" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\GROSWAGEG.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\ESCQTR1.RPT") Then
    fpListRpts.AddItem "  ESC 1st Quarter Report" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\ESCQTR1.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\ESCQTR2.RPT") Then
    fpListRpts.AddItem "  ESC 2nd Quarter Report" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\ESCQTR2.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\ESCQTR3.RPT") Then
    fpListRpts.AddItem "  ESC 3rd Quarter Report" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\ESCQTR3.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\ESCQTR4.RPT") Then
    fpListRpts.AddItem "  ESC 4th Quarter Report" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\ESCQTR4.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\EMPDATAG.RPT") Then
    fpListRpts.AddItem "  Employee Data Report" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\EMPDATAG.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\CHKISSUEG.RPT") Then
    fpListRpts.AddItem "  Checks Issued by Employee" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\CHKISSUEG.RPT")
    FoundOne = True
  End If
   
  If Exist(StartPath & "\PRRPTS\EMPHISTG.RPT") Then
    fpListRpts.AddItem "  Employee Earnings History" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\EMPHISTG.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\EMPHISTSUMG.RPT") Then
    fpListRpts.AddItem "  Employee Earnings History Summary" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\EMPHISTSUMG.RPT")
    FoundOne = True
  End If
  
'  If Exist(StartPath & "\PRRPTS\DISTRIBUACCTNUMG.RPT") Then
'    fpListRpts.AddItem "  Earnings Distribution" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\DISTRIBUACCTNUMG.RPT")
'  End If
  
  If Exist(StartPath & "\PRRPTS\DISTRIBUFUNDNUM.RPT") Then
    fpListRpts.AddItem "  Fund Number Register" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\DISTRIBUFUNDNUM.RPT")
    FoundOne = True
  End If
  
'  If Exist(StartPath & "\PRRPTS\DISTRIBUNSG.RPT") Then
'    fpListRpts.AddItem "  Earnings Distribution Register non-Split" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\DISTRIBUNSG.RPT")
'  End If
  
  If Exist(StartPath & "\PRRPTS\COMPWAGEG.RPT") Then
    fpListRpts.AddItem "  Worker's Comp Report" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\COMPWAGEG.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\CHKSBYRANGEG.RPT") Then
    fpListRpts.AddItem "  Checks in Numerical Order Report" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\CHKSBYRANGEG.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\PPDFG.RPT") Then
    fpListRpts.AddItem "  Employees to Draft Report" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\PPDFG.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\EMPDFLSTG.RPT") Then
    fpListRpts.AddItem "  Employee Draft List" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\EMPDFLSTG.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\ACCRUALG.RPT") Then
    fpListRpts.AddItem "  Accrual Report" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\ACCRUALG.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\CHECKREGG.RPT") Then
    fpListRpts.AddItem "  Check Register" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\CHECKREGG.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\W2REPORTG.RPT") Then
    fpListRpts.AddItem "  W2 Report" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\W2REPORTG.RPT")
    FoundOne = True
  End If

  If Exist(StartPath & "\PRRPTS\MANREGISG.RPT") Then
    fpListRpts.AddItem "  Manual Transaction Register" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\MANREGISG.RPT")
    FoundOne = True
  End If

  If Exist(StartPath & "\PRRPTS\EMERGENCYG.RPT") Then
    fpListRpts.AddItem "  Employee Emergency Information" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\EMERGENCYG.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\PayRate.RPT") Then
    fpListRpts.AddItem "  Employee Pay Rate Report" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\PayRate.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\941FORMS.RPT") Then
    fpListRpts.AddItem "  941 Assistance Report" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\941FORMS.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\VOIDPRN.RPT") Then
    fpListRpts.AddItem "  Void Check Review" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\VOIDPRN.RPT")
    FoundOne = True
  End If
  
  If Exist(StartPath & "\PRRPTS\TAXFRING.RPT") Then
    fpListRpts.AddItem "  Tax Fringe Report" & Chr(9) & FileDateTime(StartPath & "\PRRPTS\TAXFRING.RPT")
    FoundOne = True
  End If
  
  If FoundOne = True Then
    fpListRpts.ListIndex = 0
  End If
End Sub

Private Sub fpListRpts_DblClick()
  frmLoadingReprint.Show
  DoEvents
  fpListRpts.Col = 0
  ThisRpt$ = fpListRpts.ColText
  If QPTrim$(ThisRpt$) = "" Then
    MsgBox "Please make a selection from the list provided."
    Unload frmLoadingReprint
    Exit Sub
  End If
  frmARViewer.Show
  DoEvents

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      Call Terminate
      End
    End If
  End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyReturn:
      Call cmdPrint_Click
      KeyCode = 0
    Case vbKeyDown:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdPrint_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

