VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmEmpQuickMaintAltEarn 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll: Quick Employee Maintenance"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmEmpQuickMaintAltEarn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbEarn 
      Height          =   405
      Left            =   4699
      TabIndex        =   1
      Top             =   2820
      Width           =   4095
      _Version        =   196608
      _ExtentX        =   7223
      _ExtentY        =   714
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
      Object.TabStop         =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   0
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
      ScrollBarH      =   1
      DataFieldList   =   ""
      ColumnEdit      =   -1
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
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
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmEmpQuickMaintAltEarn.frx":08CA
   End
   Begin LpLib.fpCombo fpcmbParameters 
      Height          =   405
      Left            =   5993
      TabIndex        =   0
      Top             =   2200
      Width           =   2295
      _Version        =   196608
      _ExtentX        =   4048
      _ExtentY        =   714
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
      Object.TabStop         =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   0
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
      ScrollBarH      =   1
      DataFieldList   =   ""
      ColumnEdit      =   -1
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
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
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmEmpQuickMaintAltEarn.frx":0B89
   End
   Begin VB.CheckBox chkTerm 
      BackColor       =   &H008F8265&
      Caption         =   "Include Terminated Employees"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4133
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1740
      Width           =   3375
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   690
      Left            =   8412
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Press F10 to save all the changes made on this spreadsheet."
      Top             =   7620
      Width           =   1890
      _Version        =   131072
      _ExtentX        =   3334
      _ExtentY        =   1217
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
      ButtonDesigner  =   "frmEmpQuickMaintAltEarn.frx":0E48
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
      Height          =   690
      Left            =   6053
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Press ESC to exit this screen after each cell is examined for unsaved changes."
      Top             =   7620
      Width           =   1890
      _Version        =   131072
      _ExtentX        =   3334
      _ExtentY        =   1217
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
      ButtonDesigner  =   "frmEmpQuickMaintAltEarn.frx":1024
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExitNow 
      Height          =   690
      Left            =   3720
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Press to exit this screen without testing each cell for unsaved changes."
      Top             =   7620
      Width           =   1890
      _Version        =   131072
      _ExtentX        =   3334
      _ExtentY        =   1217
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
      ButtonDesigner  =   "frmEmpQuickMaintAltEarn.frx":1200
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdGL 
      Height          =   690
      Left            =   1339
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Press to bring up a complete list of all General Ledger numbers."
      Top             =   7620
      Width           =   1890
      _Version        =   131072
      _ExtentX        =   3334
      _ExtentY        =   1217
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
      ButtonDesigner  =   "frmEmpQuickMaintAltEarn.frx":13E0
   End
   Begin FPSpread.vaSpread vaSpread 
      Height          =   3615
      Left            =   720
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3600
      Width           =   10215
      _Version        =   196613
      _ExtentX        =   18018
      _ExtentY        =   6376
      _StockProps     =   64
      ColsFrozen      =   3
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   12648447
      MaxCols         =   6
      MaxRows         =   1000000
      OperationMode   =   2
      ProcessTab      =   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   12648447
      SpreadDesigner  =   "frmEmpQuickMaintAltEarn.frx":15BF
      VisibleCols     =   6
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClearClip 
      Height          =   405
      Left            =   9240
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Press to delete any GL number left on the clipboard that is no longer needed."
      Top             =   3000
      Width           =   1890
      _Version        =   131072
      _ExtentX        =   3334
      _ExtentY        =   714
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
      ButtonDesigner  =   "frmEmpQuickMaintAltEarn.frx":1C8AC
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Parameters:"
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
      Height          =   300
      Left            =   3353
      TabIndex        =   9
      Top             =   2320
      Width           =   2430
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   4035
      Left            =   540
      Top             =   3420
      Width           =   10575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Quick Maintenance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3323
      TabIndex        =   8
      Top             =   570
      Width           =   4995
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   3293
      Top             =   420
      Width           =   5055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Alternate Earnings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   930
      Width           =   3795
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Earning:"
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
      Height          =   300
      Left            =   2846
      TabIndex        =   6
      Top             =   2955
      Width           =   1650
   End
End
Attribute VB_Name = "frmEmpQuickMaintAltEarn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim EmployeeCount As Integer
  Dim ThisEarnNum As Integer
  Dim NumOfEarnRecs As Integer
  Dim EarnList() As ErnCodeRecType
  Dim ThisLoadSpread As Integer
  Dim ChangeSpot() As Integer
  Dim ThisChange As Integer
  Dim DontExit As Boolean
  Dim ThisParameter$
  Dim ThisEarn$
  Dim ThisTerm As Integer
  Dim BooBoo As Boolean
  
Private Sub chkTerm_Click()
  If chkTerm.Value = ThisTerm Then Exit Sub
  
  If ThisChange > 0 Then
    DontExit = True
    Call cmdEscape_Click
    If BooBoo = True Then
      BooBoo = False
      chkTerm.Value = ThisTerm
      Exit Sub
    End If
  End If
  
  ThisTerm = chkTerm.Value
  Call LoadSpread(ThisLoadSpread)
End Sub

Private Sub cmdClearClip_Click()
  Clipboard.Clear
End Sub

Private Sub cmdEscape_Click()
  Dim EmpRec As EmpData2Type
  Dim EHandle As Integer
  Dim NumOfRows As Integer
  Dim x As Integer
  Dim ThisRec As Integer
  Dim ThisAmt$
  Dim ThisAcct$
  Dim ThisGL$
  
  If ThisChange = 0 Then GoTo NoChanges
  On Error GoTo ERRORSTUFF
  NumOfRows = vaSpread.MaxRows
  OpenEmpData2File EHandle

  For x = 1 To ThisChange
    vaSpread.Row = ChangeSpot(x)
    vaSpread.Col = 6
    ThisRec = CInt(vaSpread.Text)
    Get EHandle, ThisRec, EmpRec
    vaSpread.Col = 4
    ThisAcct = QPTrim$(vaSpread.Text)
    If ThisEarnNum = 1 Then
      If QPTrim$(vaSpread.Text) <> QPTrim$(EmpRec.EMPEACT1) Then
         vaSpread.SetFocus
         vaSpread.OperationMode = OperationModeNormal
         vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
         If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPEACT1) <> "" Then
           frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Earnings #1 Account Number' field on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPEACT1) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
         ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EMPEACT1) <> "" Then
           frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Earnings #1 Account Number' field on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPEACT1) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
         ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPEACT1) = "" Then
           frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Earnings #1 Account Number' field on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
         End If
         frmMessageW3Opts.Label1.Top = 750
         frmMessageW3Opts.cmdCont.Text = "F10 Save"
         frmMessageW3Opts.cmdOption.Text = "F5 Review"
         frmMessageW3Opts.cmdExit.Text = "ESC Exit"
         frmMessageW3Opts.Show vbModal
         If frmMessageW3Opts.fptxtChoice.Text = "option" Then
           If QPTrim$(fpcmbParameters.Text) <> ThisParameter Then
             DontExit = False
             BooBoo = True
             fpcmbParameters.Text = ThisParameter
           End If
           If chkTerm.Value <> ThisTerm Then
             DontExit = False
             BooBoo = True
             chkTerm.Value = ThisTerm
           End If
           If QPTrim$(fpcmbEarn.Text) <> ThisEarn Then
             DontExit = False
             BooBoo = True
             fpcmbEarn.Text = ThisEarn
           End If
           Unload frmMessageWOpts
           Close
           Exit Sub
         ElseIf frmMessageW3Opts.fptxtChoice.Text = "continue" Then
           Unload frmMessageWOpts
           ThisGL$ = QPTrim$(vaSpread.Text)
           If CheckGLNum(ThisGL$) = False Then
             frmMessage.Label1.Caption = "On row #" + CStr(vaSpread.Row) + " the General Ledger number is not valid. Please enter a valid General Ledger number."
             frmMessage.Label1.Top = 700
             frmMessage.Show vbModal
             vaSpread.SetFocus
             vaSpread.OperationMode = OperationModeNormal
             vaSpread.SetActiveCell 4, vaSpread.Row
             DontExit = False
             BooBoo = True
             If QPTrim$(fpcmbParameters.Text) <> ThisParameter Then
               fpcmbParameters.Text = ThisParameter
             End If
             If chkTerm.Value <> ThisTerm Then
               chkTerm.Value = ThisTerm
             End If
             If QPTrim$(fpcmbEarn.Text) <> ThisEarn Then
               fpcmbEarn.Text = ThisEarn
             End If
             Close
             Exit Sub
           End If
           EmpRec.EMPEACT1 = QPTrim$(vaSpread.Text)
           Put EHandle, ThisRec, EmpRec
           MsgBox "Your change has been saved successfully"
         Else
           Unload frmMessageWOpts
           MainLog ("User warned that a change was made in the Earnings #1 Account Number field for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EMPEACT1) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
         End If
         ThisAcct = QPTrim$(vaSpread.Text)
      End If
      
      vaSpread.Col = 5
      ThisAmt = ReplaceString(vaSpread.Text, ",", "")
      ThisAmt = QPTrim$(ReplaceString(ThisAmt, "$", ""))
      If Val(ThisAmt) <> EmpRec.EMPEAMT1 Then
         vaSpread.SetFocus
         vaSpread.OperationMode = OperationModeNormal
         vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
         If Val(ThisAmt) <> 0 And EmpRec.EMPEAMT1 <> 0 Then
           frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Earnings Amount #1' on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("#,##0.00", EmpRec.EMPEAMT1)) + " to " + QPTrim$(Using("#,##0.00", Val(ThisAmt))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
         ElseIf Val(ThisAmt) = 0 And EmpRec.EMPEAMT1 <> 0 Then
           frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Earnings Amount #1' on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("#,##0.00", EmpRec.EMPEAMT1)) + " to '0'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
         ElseIf Val(ThisAmt) <> 0 And EmpRec.EMPEAMT1 = 0 Then
           frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Earnings Amount #1' on row #" + CStr(vaSpread.Row) + " from '0' to " + QPTrim$(Using("#,##0.00", Val(ThisAmt))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
         End If
         frmMessageW3Opts.Label1.Top = 750
         frmMessageW3Opts.cmdCont.Text = "F10 Save"
         frmMessageW3Opts.cmdOption.Text = "F5 Review"
         frmMessageW3Opts.cmdExit.Text = "ESC Exit"
         frmMessageW3Opts.Show vbModal
         If frmMessageW3Opts.fptxtChoice.Text = "option" Then
           If QPTrim$(fpcmbParameters.Text) <> ThisParameter Then
             DontExit = False
             BooBoo = True
             fpcmbParameters.Text = ThisParameter
           End If
           If chkTerm.Value <> ThisTerm Then
             DontExit = False
             BooBoo = True
             chkTerm.Value = ThisTerm
           End If
           If QPTrim$(fpcmbEarn.Text) <> ThisEarn Then
             DontExit = False
             BooBoo = True
             fpcmbEarn.Text = ThisEarn
           End If
           Unload frmMessageWOpts
           Close
           Exit Sub
         ElseIf frmMessageW3Opts.fptxtChoice.Text = "continue" Then
           Unload frmMessageWOpts
           If QPTrim$(ThisAcct) = "" And CDbl(ThisAmt) > 0 Then
             frmMessage.Label1.Caption = "On row #" + CStr(vaSpread.Row) + " you are attempting to save an earnings amount but no earnings account number has been entered. Please enter an earnings account number."
             frmMessage.Label1.Top = 700
             frmMessage.Show vbModal
             vaSpread.SetFocus
             vaSpread.OperationMode = OperationModeNormal
             vaSpread.SetActiveCell 4, vaSpread.Row
             DontExit = False
             BooBoo = True
             If QPTrim$(fpcmbParameters.Text) <> ThisParameter Then
               fpcmbParameters.Text = ThisParameter
             End If
             If chkTerm.Value <> ThisTerm Then
               chkTerm.Value = ThisTerm
             End If
             If QPTrim$(fpcmbEarn.Text) <> ThisEarn Then
               fpcmbEarn.Text = ThisEarn
             End If
             Close
             Exit Sub
           End If
           EmpRec.EMPEAMT1 = CDbl(ThisAmt)
           Put EHandle, ThisRec, EmpRec
           MsgBox "Your change has been saved successfully"
         Else
           Unload frmMessageWOpts
           MainLog ("User warned that a change was made in the Earnings Amount #1 field for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(Using("#,##0.00", EmpRec.EMPEAMT1)) + " to " + QPTrim$(Using("#,##0.00", Val(ThisAmt))) + " but declined to save it.")
         End If
      End If
    ElseIf ThisEarnNum = 2 Then
      If QPTrim$(UCase(vaSpread.Text)) <> QPTrim$(EmpRec.EMPEACT2) Then
         vaSpread.SetFocus
         vaSpread.OperationMode = OperationModeNormal
         vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
         If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPEACT2) <> "" Then
           frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Earnings #2 Account Number' field on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPEACT2) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
         ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EMPEACT2) <> "" Then
           frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Earnings #2 Account Number' field on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPEACT2) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
         ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPEACT2) = "" Then
           frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Earnings #2 Account Number' field on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
         End If
         frmMessageW3Opts.Label1.Top = 750
         frmMessageW3Opts.cmdCont.Text = "F10 Save"
         frmMessageW3Opts.cmdOption.Text = "F5 Review"
         frmMessageW3Opts.cmdExit.Text = "ESC Exit"
         frmMessageW3Opts.Show vbModal
         If frmMessageW3Opts.fptxtChoice.Text = "option" Then
           If QPTrim$(fpcmbParameters.Text) <> ThisParameter Then
             DontExit = False
             BooBoo = True
             fpcmbParameters.Text = ThisParameter
           End If
           If chkTerm.Value <> ThisTerm Then
             DontExit = False
             BooBoo = True
             chkTerm.Value = ThisTerm
           End If
           If QPTrim$(fpcmbEarn.Text) <> ThisEarn Then
             DontExit = False
             BooBoo = True
             fpcmbEarn.Text = ThisEarn
           End If
           Unload frmMessageWOpts
           Close
           Exit Sub
         ElseIf frmMessageW3Opts.fptxtChoice.Text = "continue" Then
           Unload frmMessageWOpts
           ThisGL$ = QPTrim$(vaSpread.Text)
           If CheckGLNum(ThisGL$) = False Then
             frmMessage.Label1.Caption = "On row #" + CStr(vaSpread.Row) + " the General Ledger number is not valid. Please enter a valid General Ledger number."
             frmMessage.Label1.Top = 700
             frmMessage.Show vbModal
             vaSpread.SetFocus
             vaSpread.OperationMode = OperationModeNormal
             vaSpread.SetActiveCell 4, vaSpread.Row
             DontExit = False
             BooBoo = True
             If QPTrim$(fpcmbParameters.Text) <> ThisParameter Then
               fpcmbParameters.Text = ThisParameter
             End If
             If chkTerm.Value <> ThisTerm Then
               chkTerm.Value = ThisTerm
             End If
             If QPTrim$(fpcmbEarn.Text) <> ThisEarn Then
               fpcmbEarn.Text = ThisEarn
             End If
             Close
             Exit Sub
           End If
           EmpRec.EMPEACT2 = QPTrim$(vaSpread.Text)
           Put EHandle, ThisRec, EmpRec
           MsgBox "Your change has been saved successfully"
         Else
           Unload frmMessageWOpts
           MainLog ("User warned that a change was made in the Earnings #2 Account Number field for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EMPEACT2) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
         End If
         ThisAcct = QPTrim$(vaSpread.Text)
      End If
      vaSpread.Col = 5
      ThisAmt = ReplaceString(vaSpread.Text, ",", "")
      ThisAmt = QPTrim$(ReplaceString(ThisAmt, "$", ""))
      If Val(ThisAmt) <> EmpRec.EMPEAMT2 Then
        vaSpread.SetFocus
        vaSpread.OperationMode = OperationModeNormal
        vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
        If Val(ThisAmt) <> 0 And EmpRec.EMPEAMT2 <> 0 Then
          frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Earnings Amount #2' field on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("#,##0.00", EmpRec.EMPEAMT2)) + " to " + QPTrim$(Using("#,##0.00", Val(ThisAmt))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
        ElseIf Val(ThisAmt) = 0 And EmpRec.EMPEAMT2 <> 0 Then
          frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Earnings Amount #2' field on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("#,##0.00", EmpRec.EMPEAMT2)) + " to '0'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
        ElseIf Val(ThisAmt) <> 0 And EmpRec.EMPEAMT2 = 0 Then
          frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Earnings Amount #2' field on row #" + CStr(vaSpread.Row) + " from '0' to " + QPTrim$(Using("#,##0.00", Val(ThisAmt))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
        End If
        frmMessageW3Opts.Label1.Top = 750
        frmMessageW3Opts.cmdCont.Text = "F10 Save"
        frmMessageW3Opts.cmdOption.Text = "F5 Review"
        frmMessageW3Opts.cmdExit.Text = "ESC Exit"
        frmMessageW3Opts.Show vbModal
        If frmMessageW3Opts.fptxtChoice.Text = "option" Then
          If QPTrim$(fpcmbParameters.Text) <> ThisParameter Then
            DontExit = False
            BooBoo = True
            fpcmbParameters.Text = ThisParameter
          End If
          If chkTerm.Value <> ThisTerm Then
            DontExit = False
            BooBoo = True
            chkTerm.Value = ThisTerm
          End If
          If QPTrim$(fpcmbEarn.Text) <> ThisEarn Then
            DontExit = False
            BooBoo = True
            fpcmbEarn.Text = ThisEarn
          End If
          Unload frmMessageWOpts
          Close
          Exit Sub
        ElseIf frmMessageW3Opts.fptxtChoice.Text = "continue" Then
          Unload frmMessageWOpts
          If QPTrim$(ThisAcct) = "" And CDbl(ThisAmt) > 0 Then
            frmMessage.Label1.Caption = "On row #" + CStr(vaSpread.Row) + " you are attempting to save an earnings amount but no earnings account number has been entered. Please enter an earnings account number."
            frmMessage.Label1.Top = 700
            frmMessage.Show vbModal
            vaSpread.SetFocus
            vaSpread.OperationMode = OperationModeNormal
            vaSpread.SetActiveCell 4, vaSpread.Row
            DontExit = False
            BooBoo = True
            If QPTrim$(fpcmbParameters.Text) <> ThisParameter Then
              fpcmbParameters.Text = ThisParameter
            End If
            If chkTerm.Value <> ThisTerm Then
              chkTerm.Value = ThisTerm
            End If
            If QPTrim$(fpcmbEarn.Text) <> ThisEarn Then
              fpcmbEarn.Text = ThisEarn
            End If
            Close
            Exit Sub
          End If
          EmpRec.EMPEAMT2 = CDbl(ThisAmt)
          Put EHandle, ThisRec, EmpRec
          MsgBox "Your change has been saved successfully"
        Else
          Unload frmMessageWOpts
          MainLog ("User warned that a change was made in the Earnings Amount #2 field for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(Using("#,##0.00", EmpRec.EMPEAMT2)) + " to " + QPTrim$(Using("#,##0.00", Val(ThisAmt))) + " but declined to save it.")
        End If
      End If
    ElseIf ThisEarnNum = 3 Then
      If QPTrim$(UCase(vaSpread.Text)) <> QPTrim$(EmpRec.EMPEACT3) Then
        vaSpread.SetFocus
        vaSpread.OperationMode = OperationModeNormal
        vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
        If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPEACT3) <> "" Then
          frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Earnings #3 Account Number' field on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPEACT3) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
        ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EMPEACT3) <> "" Then
          frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Earnings #3 Account Number' field on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPEACT3) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
        ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPEACT3) = "" Then
          frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Earnings #3 Account Number' field on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
        End If
        frmMessageW3Opts.Label1.Top = 750
        frmMessageW3Opts.cmdCont.Text = "F10 Save"
        frmMessageW3Opts.cmdOption.Text = "F5 Review"
        frmMessageW3Opts.cmdExit.Text = "ESC Exit"
        frmMessageW3Opts.Show vbModal
        If frmMessageW3Opts.fptxtChoice.Text = "option" Then
          If QPTrim$(fpcmbParameters.Text) <> ThisParameter Then
            DontExit = False
            BooBoo = True
            fpcmbParameters.Text = ThisParameter
          End If
          If chkTerm.Value <> ThisTerm Then
            DontExit = False
            BooBoo = True
            chkTerm.Value = ThisTerm
          End If
          If QPTrim$(fpcmbEarn.Text) <> ThisEarn Then
            DontExit = False
            BooBoo = True
            fpcmbEarn.Text = ThisEarn
          End If
          Unload frmMessageWOpts
          Close
          Exit Sub
        ElseIf frmMessageW3Opts.fptxtChoice.Text = "continue" Then
          Unload frmMessageWOpts
          ThisGL$ = QPTrim$(vaSpread.Text)
          If CheckGLNum(ThisGL$) = False Then
            frmMessage.Label1.Caption = "On row #" + CStr(vaSpread.Row) + " the General Ledger number is not valid. Please enter a valid General Ledger number."
            frmMessage.Label1.Top = 700
            frmMessage.Show vbModal
            vaSpread.SetFocus
            vaSpread.OperationMode = OperationModeNormal
            vaSpread.SetActiveCell 4, vaSpread.Row
            DontExit = False
            BooBoo = True
            If QPTrim$(fpcmbParameters.Text) <> ThisParameter Then
              fpcmbParameters.Text = ThisParameter
            End If
            If chkTerm.Value <> ThisTerm Then
              chkTerm.Value = ThisTerm
            End If
            If QPTrim$(fpcmbEarn.Text) <> ThisEarn Then
              fpcmbEarn.Text = ThisEarn
            End If
            Close
            Exit Sub
          End If
          EmpRec.EMPEACT3 = QPTrim$(vaSpread.Text)
          Put EHandle, ThisRec, EmpRec
          MsgBox "Your change has been saved successfully"
        Else
          Unload frmMessageWOpts
          MainLog ("User warned that a change was made in the Earnings #3 Account Number field for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EMPEACT3) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
        End If
        ThisAcct = QPTrim$(vaSpread.Text)
      End If
      vaSpread.Col = 5
      ThisAmt = ReplaceString(vaSpread.Text, ",", "")
      ThisAmt = QPTrim$(ReplaceString(ThisAmt, "$", ""))
      If Val(ThisAmt) <> EmpRec.EMPEAMT3 Then
        vaSpread.SetFocus
        vaSpread.OperationMode = OperationModeNormal
        vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
        If Val(ThisAmt) <> 0 And EmpRec.EMPEAMT3 <> 0 Then
          frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Earnings Amount #3' on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("#,##0.00", EmpRec.EMPEAMT3)) + " to " + QPTrim$(Using("#,##0.00", Val(ThisAmt))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
        ElseIf Val(ThisAmt) = 0 And EmpRec.EMPEAMT3 <> 0 Then
          frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Earnings Amount #3' on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("#,##0.00", EmpRec.EMPEAMT3)) + " to '0'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
        ElseIf Val(ThisAmt) <> 0 And EmpRec.EMPEAMT3 = 0 Then
          frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Earnings Amount #3' on row #" + CStr(vaSpread.Row) + " from '0' to " + QPTrim$(Using("#,##0.00", Val(ThisAmt))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
        End If
        frmMessageW3Opts.Label1.Top = 750
        frmMessageW3Opts.cmdCont.Text = "F10 Save"
        frmMessageW3Opts.cmdOption.Text = "F5 Review"
        frmMessageW3Opts.cmdExit.Text = "ESC Exit"
        frmMessageW3Opts.Show vbModal
        If frmMessageW3Opts.fptxtChoice.Text = "option" Then
          If QPTrim$(fpcmbParameters.Text) <> ThisParameter Then
            DontExit = False
            BooBoo = True
            fpcmbParameters.Text = ThisParameter
          End If
          If chkTerm.Value <> ThisTerm Then
            DontExit = False
            BooBoo = True
            chkTerm.Value = ThisTerm
          End If
          If QPTrim$(fpcmbEarn.Text) <> ThisEarn Then
            DontExit = False
            BooBoo = True
            fpcmbEarn.Text = ThisEarn
          End If
          Unload frmMessageWOpts
          Close
          Exit Sub
        ElseIf frmMessageW3Opts.fptxtChoice.Text = "continue" Then
          Unload frmMessageWOpts
          If QPTrim$(ThisAcct) = "" And CDbl(ThisAmt) > 0 Then
            frmMessage.Label1.Caption = "On row #" + CStr(vaSpread.Row) + " you are attempting to save an earnings amount but no earnings account number has been entered. Please enter an earnings account number."
            frmMessage.Label1.Top = 700
            frmMessage.Show vbModal
            vaSpread.SetFocus
            vaSpread.OperationMode = OperationModeNormal
            vaSpread.SetActiveCell 4, vaSpread.Row
            DontExit = False
            BooBoo = True
            If QPTrim$(fpcmbParameters.Text) <> ThisParameter Then
              fpcmbParameters.Text = ThisParameter
            End If
            If chkTerm.Value <> ThisTerm Then
              chkTerm.Value = ThisTerm
            End If
            If QPTrim$(fpcmbEarn.Text) <> ThisEarn Then
              fpcmbEarn.Text = ThisEarn
            End If
            Close
            Exit Sub
          End If
          EmpRec.EMPEAMT3 = CDbl(ThisAmt)
          Put EHandle, ThisRec, EmpRec
          MsgBox "Your change has been saved successfully"
        Else
          Unload frmMessageWOpts
          MainLog ("User warned that a change was made in the earnings #3 amount for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(Using("#,##0.00", EmpRec.EMPEAMT3)) + " to " + QPTrim$(Using("#,##0.00", Val(ThisAmt))) + " but declined to save it.")
        End If
      End If
    End If
  Next x

  Close
  
NoChanges:
  If DontExit = True Then
    DontExit = False
    Exit Sub
  End If
  
  frmEmpQuickMaintMenu.Show
  DoEvents
  Unload Me
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmEmpQuickMaintAltEarn", "cmdEscape_Click", Erl)
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

  
End Sub

Private Sub cmdExitNow_Click()
  frmEmpQuickMaintMenu.Show
  DoEvents
  Unload Me
End Sub


Private Sub cmdGL_Click()
  frmGLPickList.Show vbModal, Me
End Sub

Private Sub cmdSave_Click()
  Dim EmpRec As EmpData2Type
  Dim EHandle As Integer
  Dim NumOfRows As Integer
  Dim x As Integer
  Dim ThisRec As Integer
  Dim ThisAmt$
  Dim ThisAcct$
  Dim ThisGL$
  
  If ThisChange = 0 Then
    frmMessage.Label1.Caption = "No changes made. Save aborted."
    frmMessage.Label1.Top = 900
    frmMessage.Show vbModal
    GoTo NoChanges
  End If
  
  On Error GoTo ERRORSTUFF
  
  If CheckAcctAmt = False Then
    Exit Sub
  End If
  
  frmLoadingRpt.Label1.Caption = "Saving......"
  frmLoadingRpt.Show
  DoEvents
  
  NumOfRows = vaSpread.MaxRows
  OpenEmpData2File EHandle
  
  For x = 1 To ThisChange
    vaSpread.Row = ChangeSpot(x)
    vaSpread.Col = 6
    ThisRec = vaSpread.Value
    Get EHandle, ThisRec, EmpRec
    vaSpread.Col = 4
    ThisGL$ = QPTrim(vaSpread.Text)
    If QPTrim$(ThisGL) <> "" And CheckGLNum(ThisGL$) = False Then
      frmMessage.Label1.Caption = "On row #" + CStr(vaSpread.Row) + " the General Ledger number is not valid. Please enter a valid General Ledger number."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell 4, vaSpread.Row
      DontExit = False
      BooBoo = True
      If QPTrim$(fpcmbParameters.Text) <> ThisParameter Then
        fpcmbParameters.Text = ThisParameter
      End If
      If chkTerm.Value <> ThisTerm Then
        chkTerm.Value = ThisTerm
      End If
      If QPTrim$(fpcmbEarn.Text) <> ThisEarn Then
        fpcmbEarn.Text = ThisEarn
      End If
      Close
      Exit Sub
    End If
    
    If ThisEarnNum = 1 Then
      EmpRec.EMPEACT1 = QPTrim(vaSpread.Text)
      vaSpread.Col = 5
      ThisAmt = ReplaceString$(vaSpread.Text, "$", "")
      ThisAmt = QPTrim$(ReplaceString$(ThisAmt, ",", ""))
      EmpRec.EMPEAMT1 = Val(ThisAmt)
    ElseIf ThisEarnNum = 2 Then
      EmpRec.EMPEACT2 = QPTrim(vaSpread.Text)
      vaSpread.Col = 5
      ThisAmt = ReplaceString$(vaSpread.Text, "$", "")
      ThisAmt = QPTrim$(ReplaceString$(ThisAmt, ",", ""))
      EmpRec.EMPEAMT2 = Val(ThisAmt)
    ElseIf ThisEarnNum = 3 Then
      EmpRec.EMPEACT3 = QPTrim(vaSpread.Text)
      vaSpread.Col = 5
      ThisAmt = ReplaceString$(vaSpread.Text, "$", "")
      ThisAmt = QPTrim$(ReplaceString$(ThisAmt, ",", ""))
      EmpRec.EMPEAMT3 = Val(ThisAmt)
    End If
    Put EHandle, ThisRec, EmpRec
  Next x
  
  Close
  Unload frmLoadingRpt
  MsgBox "Employee data has been saved successfully."
  ThisChange = 0
NoChanges:
  frmEmpQuickMaintMenu.Show
  DoEvents
  Unload Me
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmEmpQuickMaintAltEarn", "cmdSave_Click", Erl)
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

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdEscape_Click
      KeyCode = 0
    Case vbKeyF2:
      SendKeys "%C"
      Call cmdClearClip_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      Call cmdSave_Click
      KeyCode = 0
    Case vbKeyF6:
      SendKeys "%F"
      Call cmdExitNow_Click
      KeyCode = 0
    Case vbKeyF12:
      SendKeys "%G"
      Call cmdGL_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub LoadMe()
  Dim IdxRec As NameSortIdxType
  Dim XHandle As Integer
  Dim EarnRec As ErnCodeRecType
  Dim EHandle As Integer
  Dim NumOfEarnRecs As Integer
  Dim x As Integer
  
  OpenErnCodeFile EHandle
  NumOfEarnRecs = LOF(EHandle) / Len(EarnRec)
  ReDim EarnList(1 To NumOfEarnRecs) As ErnCodeRecType
  For x = 1 To NumOfEarnRecs
    Get EHandle, x, EarnRec
    EarnList(x) = EarnRec
    fpcmbEarn.AddItem QPTrim$(EarnRec.ERNCODE1)
  Next x
  Close EHandle
  
  fpcmbEarn.Text = QPTrim$(EarnList(1).ERNCODE1)
  
  OpenEmpIdxLNameFile XHandle
  EmployeeCount = LOF(XHandle) / 2
  Close
  
  ThisEarnNum = 1
  DontExit = False
  ThisChange = 0
  ThisParameter = "All Employees"
  ThisTerm = chkTerm.Value
  ThisEarn = QPTrim$(EarnList(1).ERNCODE1)
  BooBoo = False
  
  fpcmbParameters.Text = "All Employees"
  fpcmbParameters.AddItem "All Employees"
  fpcmbParameters.AddItem "Full-Time"
  fpcmbParameters.AddItem "Part-Time"
  fpcmbParameters.AddItem "Seasonal"
  fpcmbParameters.AddItem "Temporary"
  
  ThisLoadSpread = 1
  Call LoadSpread(ThisLoadSpread)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdEscape.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmEmpQuickMaintAltEarn.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadSpread(SpreadType As Integer)
  Dim x As Integer, y As Integer
  Dim EmpRec As EmpData2Type
  Dim EHandle As Integer
  Dim NumOfEmpRecs As Integer
  Dim IdxRec As NameSortIdxType
  Dim XHandle As Integer
  Dim RowMax As Integer
  Dim EmpType$
  
  On Error GoTo ERRORSTUFF
  
  vaSpread.MaxRows = EmployeeCount
  
  Call ClearChanges
  ReDim ChangeSpot(1 To vaSpread.MaxRows) As Integer
  
  OpenEmpIdxLNameFile XHandle
  NumOfEmpRecs = LOF(XHandle) / 2
  If NumOfEmpRecs = 0 Then
    MsgBox "No employee records have been saved."
    Close
    Exit Sub
  End If
  
  ReDim ThisIdx(1 To NumOfEmpRecs) As Integer
  For x = 1 To NumOfEmpRecs
    Get XHandle, x, IdxRec.DataRecNum
    ThisIdx(x) = IdxRec.DataRecNum
  Next x
  Close XHandle
  
  vaSpread.ClearRange -1, -1, -1, -1, True

  OpenEmpData2File EHandle
  For x = 1 To NumOfEmpRecs
    Get EHandle, ThisIdx(x), EmpRec
    If EmpRec.Deleted = -1 Then GoTo SkipEmp
    EmpType = QPTrim$(EmpRec.EMPSTATS)
    Select Case SpreadType
      Case 1
      Case 2
        If EmpType <> "Full-Time" Then GoTo SkipEmp
      Case 3
        If EmpType <> "Part-Time" Then GoTo SkipEmp
      Case 4
        If EmpType <> "Seasonal" Then GoTo SkipEmp
      Case 5
        If EmpType <> "Temporary" Then GoTo SkipEmp
      Case Else
    End Select
    If chkTerm.Value = 0 Then
      If EmpRec.EMPTDATE > 0 Then GoTo SkipEmp
    End If
    RowMax = RowMax + 1
    vaSpread.Col = 1
    vaSpread.Row = RowMax
    vaSpread.Text = QPTrim$(EmpRec.EmpNo)
    vaSpread.Col = 2
    vaSpread.Row = RowMax
    vaSpread.Text = QPTrim$(EmpRec.EmpLName)
    vaSpread.Col = 3
    vaSpread.Row = RowMax
    vaSpread.Text = QPTrim$(EmpRec.EmpFName)
    vaSpread.Col = 4
    vaSpread.Row = RowMax
    If ThisEarnNum = 1 Then
      vaSpread.Text = QPTrim$(EmpRec.EMPEACT1)
    ElseIf ThisEarnNum = 2 Then
      vaSpread.Text = QPTrim$(EmpRec.EMPEACT2)
    ElseIf ThisEarnNum = 3 Then
      vaSpread.Text = QPTrim$(EmpRec.EMPEACT3)
    End If
    vaSpread.Col = 5
    vaSpread.Row = RowMax
    If ThisEarnNum = 1 Then
      vaSpread.Text = EmpRec.EMPEAMT1
    ElseIf ThisEarnNum = 2 Then
      vaSpread.Text = EmpRec.EMPEAMT2
    ElseIf ThisEarnNum = 3 Then
      vaSpread.Text = EmpRec.EMPEAMT3
    End If
    vaSpread.Col = 6
    vaSpread.Row = RowMax
    vaSpread.Text = ThisIdx(x)
SkipEmp:
  Next x
  
  Close EHandle
  
  If RowMax = 0 Then
    MsgBox "There are no employees that fit the parameters entered."
    vaSpread.MaxRows = EmployeeCount
    Close
    Exit Sub
  End If
  
  vaSpread.MaxRows = RowMax
  vaSpread.OperationMode = OperationModeNormal
  vaSpread.SetActiveCell 4, 1
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmEmpQuickMaintAltEarn", "LoadSpread", Erl)
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
  
End Sub

Private Sub fpcmbEarn_Click()
  Dim x As Integer
  
  If ThisEarn = QPTrim$(fpcmbEarn.Text) Then Exit Sub
  If ThisChange > 0 Then
    DontExit = True
    Call cmdEscape_Click
    If BooBoo = True Then
      BooBoo = False
      Exit Sub
    End If
  End If
  
  ThisEarn = QPTrim$(fpcmbEarn.Text)
  ThisEarnNum = fpcmbEarn.ListIndex + 1
  Call LoadSpread(ThisLoadSpread)
End Sub

Private Sub fpcmbEarn_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbEarn.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbEarn.ListIndex = -1
  End If
  If fpcmbEarn.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcmbParameters.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbParameters_Click()
  If ThisParameter = QPTrim$(fpcmbParameters.Text) Then Exit Sub
  If ThisChange > 0 Then
    DontExit = True
    Call cmdEscape_Click
    If BooBoo = True Then
      BooBoo = False
      fpcmbParameters.Text = ThisParameter
      GoTo BooBooFound
    End If
  End If
  
  ThisParameter = QPTrim$(fpcmbParameters.Text)
  
  Select Case QPTrim$(fpcmbParameters.Text)
    Case "All Employees"
      ThisLoadSpread = 1
    Case "Full-Time"
      ThisLoadSpread = 2
    Case "Part-Time"
      ThisLoadSpread = 3
    Case "Seasonal"
      ThisLoadSpread = 4
    Case "Temporary"
      ThisLoadSpread = 5
    Case Else
  End Select
  ThisParameter = QPTrim$(fpcmbParameters.Text)
  
  Call LoadSpread(ThisLoadSpread)
  
BooBooFound:
End Sub

Private Sub fpcmbParameters_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbParameters.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbParameters.ListIndex = -1
  End If
  If fpcmbParameters.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcmbEarn.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Function CheckAcctAmt() As Boolean
  Dim x As Integer
  Dim ThisAcct$
  Dim ThisAmt$

  CheckAcctAmt = True

  For x = 1 To ThisChange
    vaSpread.Row = ChangeSpot(x)
    vaSpread.Col = 4
    ThisAcct = QPTrim(vaSpread.Text)
    vaSpread.Col = 5
    ThisAmt = ReplaceString(vaSpread.Text, "$", "")
    ThisAmt = QPTrim$(ReplaceString(ThisAmt, ",", ""))
    If Val(ThisAmt) > 0 And QPTrim$(ThisAcct) = "" Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell 4, vaSpread.Row
      frmMessage.Label1.Caption = "On row " + CStr(vaSpread.Row) + " you are attempting to save a value in the 'Earnings Amount' cell but there is nothing saved for 'Earnings Account Number'. Please enter an earnings account number if you wish to save a value in the 'Earnings Amount' cell."
      frmMessage.Label1.Top = 650
      frmMessage.Show vbModal
      CheckAcctAmt = False
      Exit Function
    End If
  Next x
  
End Function

Private Sub vaSpread_Change(ByVal Col As Long, ByVal Row As Long)
  Dim x As Integer
  
  For x = 1 To ThisChange
    If ChangeSpot(x) = Row Then
      Exit For
    End If
  Next x
  
  If x > ThisChange Then
    ThisChange = ThisChange + 1
    ChangeSpot(ThisChange) = Row
  End If
  
End Sub

Private Sub ClearChanges()
  ThisChange = 0
  ReDim ChangeSpot(0 To 0) As Integer
End Sub

Private Sub vaSpread_Click(ByVal Col As Long, ByVal Row As Long)
  vaSpread.OperationMode = OperationModeRow
End Sub

Private Sub vaSpread_DblClick(ByVal Col As Long, ByVal Row As Long)
  Dim ThisText$
  'in order to edit account numbers you have to double click
  'the cell...double clicking also inserts whatever is on the
  'clipboard into the cell so a provision had to be made to
  'give the user the option to copy the clipboard contents to
  'the cell or edit the cell
  
  If Col < 4 Then
    MsgBox "This column is read only"
  End If
  
  If Col = 4 Then
    ThisText = Clipboard.GetText
    If QPTrim$(ThisText) = "" Then
      Exit Sub
    End If
    vaSpread.Col = 4
    vaSpread.Row = Row
    If MsgBox("Do you wish to copy the contents of the clipboard (" + ThisText + ") into this cell?", vbYesNo) = vbYes Then
      vaSpread.Text = Clipboard.GetText
      ThisChange = ThisChange + 1
      ChangeSpot(ThisChange) = Row
    Else
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell Col, Row
    End If
  End If
End Sub

Private Sub vaSpread_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
  vaSpread.Row = Row
  vaSpread.Col = 1
  vaSpread.BackColor = &HC0FFFF
  vaSpread.Col = 2
  vaSpread.BackColor = &HC0FFFF
  vaSpread.Col = 3
  vaSpread.BackColor = &HC0FFFF
  vaSpread.Col = 4
  vaSpread.BackColor = &HC0FFFF
  vaSpread.Col = 5
  vaSpread.BackColor = &HC0FFFF

End Sub

Private Sub vaSpread_KeyPress(KeyAscii As Integer)
  vaSpread.OperationMode = OperationModeRow

End Sub

Private Sub vaSpread_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
  vaSpread.BackColorStyle = BackColorStyleUnderGrid
  vaSpread.Row = Row
  vaSpread.Col = 1
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 2
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 3
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 4
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 5
  vaSpread.BackColor = &H80000005

End Sub

Private Function CheckGLNum(ThisGL$) As Boolean
   Dim SHandle As Integer
   Dim SysRec As RegDSysFileRecType
   Dim x As Integer
   Dim JGLIdxRec(1) As JGLAcctIdxType
   Dim GLIdxNum$
   Dim GLDHandle As Integer
   Dim GLIdxRecLen As Integer
   Dim GLDescRecLen As Integer
   Dim TotalAccts As Integer
   Dim GLIDATDesc$
   Dim GLDesc(1) As GLAcctRecType
   Dim GLIdxHandle As Integer
   Dim FundLength As Integer
   Dim AcctLength As Integer
   Dim DetLength As Integer
   
   On Error GoTo ERRORSTUFF
   
   CheckGLNum = True
   
   OpenSysFile SHandle
   Get SHandle, 1, SysRec
   Close SHandle
   
   If QPTrim$(SysRec.GLCheckYN) = "N" Then Exit Function
   
   Call GetAcctStruct(CurrCitiPath, FundLength, AcctLength, DetLength)
   If FundLength = 0 And AcctLength = 0 And DetLength = 0 Then
     Exit Function
   End If
   
   If Exist(QPTrim$(CurrCitiPath) + "GLACCT.IDX") Then
     GLIdxNum$ = QPTrim$(CurrCitiPath) + "GLACCT.IDX"
   ElseIf Exist(QPTrim$(CurrCitiPath) + "\GLACCT.IDX") Then
     GLIdxNum$ = QPTrim$(CurrCitiPath) + "\GLACCT.IDX"
   Else
     MsgBox "No G/L account number validation possible...GLACCT.IDX could not be found."
     Exit Function
   End If

   If Exist(QPTrim$(CurrCitiPath) + "GLACCT.DAT") Then
     GLIDATDesc$ = QPTrim$(CurrCitiPath) + "GLACCT.DAT"
   ElseIf Exist(QPTrim$(CurrCitiPath) + "\GLACCT.DAT") Then
     GLIDATDesc$ = QPTrim$(CurrCitiPath) + "\GLACCT.DAT"
   Else
     MsgBox "No G/L account number validation possible...GLACCT.DAT could not be found."
     Exit Function
   End If
   
   GLIdxRecLen = Len(JGLIdxRec(1))
   GLDescRecLen = Len(GLDesc(1))
   TotalAccts = FileSize(GLIDATDesc$) \ GLDescRecLen
   
   If TotalAccts = 0 Then Exit Function
   ReDim DescBuff(1 To TotalAccts)
   GLIdxHandle = FreeFile
   Open GLIdxNum$ For Random As GLIdxHandle Len = GLIdxRecLen
   For x = 1 To TotalAccts
     Get GLIdxHandle, x, JGLIdxRec(1)
     DescBuff(x) = JGLIdxRec(1).RecNo
   Next x
   Close GLIdxHandle
   
   GLDHandle = FreeFile
   ThisGL$ = ReplaceString(ThisGL, "-", "")
   Open GLIDATDesc$ For Random As GLDHandle Len = GLDescRecLen
   For x = 1 To TotalAccts
   If DescBuff(x) = 0 Then GoTo DescBuffIsZero
     Get GLDHandle, DescBuff(x), GLDesc(1)
       If ThisGL$ = QPTrim$(ReplaceString(GLDesc(1).Num, "-", "")) Then
         Exit For
       End If
DescBuffIsZero:
   Next x
   Close GLDHandle
   
   If x > TotalAccts Then CheckGLNum = False
   
   Exit Function
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmEmpQuickMaintAltEarn", "CheckGLNum", Erl)
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
   
End Function

