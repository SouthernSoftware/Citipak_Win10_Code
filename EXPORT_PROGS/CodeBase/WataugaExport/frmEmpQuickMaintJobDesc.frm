VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmEmpQuickMaintJobDesc 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll: Quick Employee Maintenance"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmEmpQuickMaintJobDesc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbParameters 
      Height          =   405
      Left            =   6000
      TabIndex        =   1
      Top             =   2280
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
      ColDesigner     =   "frmEmpQuickMaintJobDesc.frx":08CA
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
      Left            =   4200
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1740
      Width           =   3495
   End
   Begin FPSpread.vaSpread vaSpread 
      Height          =   4095
      Left            =   660
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3060
      Width           =   10335
      _Version        =   196613
      _ExtentX        =   18230
      _ExtentY        =   7223
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
      MaxCols         =   16
      MaxRows         =   1000000
      OperationMode   =   2
      ProcessTab      =   -1  'True
      RestrictCols    =   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   12648447
      SpreadDesigner  =   "frmEmpQuickMaintJobDesc.frx":0C6D
      VisibleCols     =   13
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   690
      Left            =   7673
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
      ButtonDesigner  =   "frmEmpQuickMaintJobDesc.frx":1C398
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
      Height          =   690
      Left            =   4920
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
      ButtonDesigner  =   "frmEmpQuickMaintJobDesc.frx":1C5AC
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExitNow 
      Height          =   690
      Left            =   2078
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
      ButtonDesigner  =   "frmEmpQuickMaintJobDesc.frx":1C7C0
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "* = Required Field"
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
      Left            =   600
      TabIndex        =   9
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Job Description"
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
      Left            =   4133
      TabIndex        =   8
      Top             =   930
      Width           =   3315
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
      TabIndex        =   7
      Top             =   570
      Width           =   4995
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   540
      Top             =   2940
      Width           =   10575
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
      Left            =   3360
      TabIndex        =   6
      Top             =   2400
      Width           =   2415
   End
End
Attribute VB_Name = "frmEmpQuickMaintJobDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim EmployeeCount As Integer
  Dim GThisRow1 As Integer
  Dim ThisLoadSpread As Integer
  Dim ChangeSpot() As Integer
  Dim ThisChange As Integer
  Dim ThisParameter$
  Dim ThisTerm As Integer
  Dim DontExit As Boolean
  Dim BooBoo As Boolean

Private Sub chkTerm_Click()
  If chkTerm.Value = ThisTerm Then Exit Sub
  If ThisChange > 0 Then
    DontExit = True
    Call cmdEscape_Click
    If BooBoo = True Then
      chkTerm.Value = ThisTerm
      BooBoo = False
      Exit Sub
    End If
  End If
  ThisTerm = chkTerm.Value
  Call LoadSpread(ThisLoadSpread)
End Sub

Private Sub cmdEscape_Click()
  Dim EmpRec As EmpData2Type
  Dim EHandle As Integer
  Dim NumOfRows As Integer
  Dim x As Integer
  Dim ThisRec As Integer
  Dim ThisPRate$
  Dim ThisOTRate$
  Dim ThisDate$
  Dim ThisAmt$
  Dim OldPRate As Double
  Dim OldORate As Double
  Dim OldPFreq$
  Dim OldPType$
  Dim NewPRate As Double
  Dim NewORate As Double
  Dim NewPFreq$
  Dim NewPType$
  Dim UpdatePay As Boolean
  
  If ThisChange = 0 Then GoTo NoChanges
  On Error GoTo ERRORSTUFF
  
  UpdatePay = False
  NumOfRows = vaSpread.MaxRows
  OpenEmpData2File EHandle
  
  For x = 1 To ThisChange
    OldPRate = 0
    OldORate = 0
    OldPFreq = ""
    OldPType = ""
    NewPRate = 0
    NewORate = 0
    NewPFreq = ""
    NewPType = ""
    UpdatePay = False
    vaSpread.Row = ChangeSpot(x)
    vaSpread.Col = 16
    ThisRec = CInt(vaSpread.Text)
    Get EHandle, ThisRec, EmpRec
    vaSpread.Col = 4
    If QPTrim$(vaSpread.Text) <> QPTrim$(UCase(EmpRec.EMPJOB)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPJOB) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the job title on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPJOB) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EMPJOB) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the job title on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPJOB) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPJOB) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the job title on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        Unload frmMessageWOpts
        Close
        Exit Sub
      ElseIf frmMessageW3Opts.fptxtChoice.Text = "continue" Then
        Unload frmMessageWOpts
        EmpRec.EMPJOB = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the job title for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EMPJOB) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    vaSpread.Col = 5
    If QPTrim$(vaSpread.Text) <> QPTrim$(UCase(EmpRec.EMPWCCLS)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPWCCLS) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the W/C Code on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPWCCLS) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EMPWCCLS) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the W/C Code on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPWCCLS) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPWCCLS) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the W/C Code on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        Unload frmMessageWOpts
        Close
        Exit Sub
      ElseIf frmMessageW3Opts.fptxtChoice.Text = "continue" Then
        Unload frmMessageWOpts
        If QPTrim$(vaSpread.Text) = "" Then
          frmMessage.Label1.Caption = "This is a required field and cannot be left blank. Save aborted."
          frmMessage.Label1.Top = 800
          frmMessage.Show vbModal
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
          vaSpread.SetFocus
          vaSpread.OperationMode = OperationModeNormal
          vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
          DontExit = False
          BooBoo = True
          Close
          Exit Sub
        End If
        EmpRec.EMPWCCLS = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the W/C Code for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EMPWCCLS) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    vaSpread.Col = 6
    If QPTrim$(UCase(vaSpread.Text)) <> QPTrim$(UCase(EmpRec.EMPSTATS)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPSTATS) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the status on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPSTATS) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EMPSTATS) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the status on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPSTATS) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPSTATS) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the status on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        Unload frmMessageWOpts
        Close
        Exit Sub
      ElseIf frmMessageW3Opts.fptxtChoice.Text = "continue" Then
        Unload frmMessageWOpts
        If QPTrim$(vaSpread.Text) = "" Then
          frmMessage.Label1.Caption = "This is a required field and cannot be left blank. Save aborted."
          frmMessage.Label1.Top = 800
          frmMessage.Show vbModal
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
          DontExit = False
          BooBoo = True
          vaSpread.SetFocus
          vaSpread.OperationMode = OperationModeNormal
          vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
          Close
          Exit Sub
        End If
        EmpRec.EMPSTATS = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the status for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EMPSTATS) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    vaSpread.Col = 7
    OldPType = QPTrim$(EmpRec.EMPPTYPE)
    NewPType = QPTrim$(UCase(vaSpread.Text))
    If QPTrim$(UCase(vaSpread.Text)) <> QPTrim$(UCase(EmpRec.EMPPTYPE)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPPTYPE) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the pay type on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPPTYPE) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EMPPTYPE) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the pay type on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPPTYPE) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPPTYPE) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the pay type on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        Unload frmMessageWOpts
        Close
        Exit Sub
      ElseIf frmMessageW3Opts.fptxtChoice.Text = "continue" Then
        Unload frmMessageWOpts
        If QPTrim$(vaSpread.Text) = "" Then
          frmMessage.Label1.Caption = "This is a required field and cannot be left blank. Save aborted."
          frmMessage.Label1.Top = 800
          frmMessage.Show vbModal
          vaSpread.SetFocus
          vaSpread.OperationMode = OperationModeNormal
          vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
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
          DontExit = False
          BooBoo = True
          Close
          Exit Sub
        End If
        EmpRec.EMPPTYPE = QPTrim$(vaSpread.Text)
        UpdatePay = True
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        NewPType = QPTrim$(EmpRec.EMPPTYPE)
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the pay type for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EMPPTYPE) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    vaSpread.Col = 8
    OldPFreq = QPTrim$(UCase(EmpRec.EMPPFREQ))
    NewPFreq = QPTrim$(UCase(vaSpread.Text))
    If QPTrim$(UCase(vaSpread.Text)) <> QPTrim$(UCase(EmpRec.EMPPFREQ)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPPFREQ) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the frequency on row #" + CStr(vaSpread.Row) + " from " + UCase(QPTrim$(EmpRec.EMPPFREQ)) + " to " + UCase(QPTrim$(vaSpread.Text)) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EMPPFREQ) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the frequency on row #" + CStr(vaSpread.Row) + " from " + UCase(QPTrim$(EmpRec.EMPPFREQ)) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPPFREQ) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the frequency on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + UCase(QPTrim$(vaSpread.Text)) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        Unload frmMessageWOpts
        Close
        Exit Sub
      ElseIf frmMessageW3Opts.fptxtChoice.Text = "continue" Then
        Unload frmMessageWOpts
        If QPTrim$(vaSpread.Text) = "" Then
          frmMessage.Label1.Caption = "This is a required field and cannot be left blank. Save aborted."
          frmMessage.Label1.Top = 800
          frmMessage.Show vbModal
          vaSpread.SetFocus
          vaSpread.OperationMode = OperationModeNormal
          vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
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
          DontExit = False
          BooBoo = True
          Close
          Exit Sub
        End If
        EmpRec.EMPPFREQ = QPTrim$(vaSpread.Text)
        UpdatePay = True
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        NewPFreq = QPTrim$(EmpRec.EMPPFREQ)
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the frequency for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EMPPFREQ) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    vaSpread.Col = 9
    If Val(vaSpread.Text) <> CStr(EmpRec.EMPBCODE) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(ReplaceString(vaSpread.Text, "%", "")) <> "" And EmpRec.EMPBCODE <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the benefit percentage on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using$("##0.00", EmpRec.EMPBCODE)) + "% to " + QPTrim$(vaSpread.Text) + "%. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(ReplaceString(vaSpread.Text, "%", "")) = "" And EmpRec.EMPBCODE <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the benefit percentage on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using$("##0.00", EmpRec.EMPBCODE)) + "% to '0.00%'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(ReplaceString(vaSpread.Text, "%", "")) <> "" And EmpRec.EMPBCODE = 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the benefit percentage on row #" + CStr(vaSpread.Row) + " from '0.00%' to " + QPTrim$(vaSpread.Text) + "%. To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        Unload frmMessageWOpts
        Close
        Exit Sub
      ElseIf frmMessageW3Opts.fptxtChoice.Text = "continue" Then
        Unload frmMessageWOpts
        If QPTrim$(ReplaceString(vaSpread.Text, "%", "")) = "" Then
          frmMessage.Label1.Caption = "This is a required field and cannot be left blank. Save aborted."
          frmMessage.Label1.Top = 800
          frmMessage.Show vbModal
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
          DontExit = False
          BooBoo = True
          vaSpread.SetFocus
          vaSpread.OperationMode = OperationModeNormal
          vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
          Close
          Exit Sub
        End If
        EmpRec.EMPBCODE = CDbl(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the benefit percent for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(Using$("##0.00", EmpRec.EMPBCODE)) + "%" + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    vaSpread.Col = 10
    ThisPRate = ReplaceString(vaSpread.Text, "$", "")
    ThisPRate = ReplaceString(ThisPRate, ",", "")
    NewPRate = Val(ThisPRate)
    If Val(ThisPRate) <> EmpRec.EMPPRATE Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(ReplaceString(vaSpread.Text, "$", "")) <> "" And EmpRec.EMPPRATE <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the pay rate on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using$("$#,##0.00", EmpRec.EMPPRATE)) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(ReplaceString(vaSpread.Text, "$", "")) = "" And EmpRec.EMPPRATE <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the pay rate on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using$("$#,##0.00", EmpRec.EMPPRATE)) + " to '$0.00'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(ReplaceString(vaSpread.Text, "$", "")) <> "" And EmpRec.EMPPRATE = 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the pay rate on row #" + CStr(vaSpread.Row) + " from '$0.00' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        Unload frmMessageWOpts
        Close
        Exit Sub
      ElseIf frmMessageW3Opts.fptxtChoice.Text = "continue" Then
        Unload frmMessageWOpts
        If CDbl(ReplaceString(vaSpread.Text, "$", "")) = 0 Then
          frmMessage.Label1.Caption = "This is a required field and cannot be left blank. Save aborted."
          frmMessage.Label1.Top = 800
          frmMessage.Show vbModal
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
          DontExit = False
          BooBoo = True
          vaSpread.SetFocus
          vaSpread.OperationMode = OperationModeNormal
          vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
          Close
          Exit Sub
        End If
        ThisAmt = ReplaceString$(vaSpread.Text, "$", "")
        ThisAmt = ReplaceString$(ThisAmt, ",", "")
        EmpRec.EMPPRATE = Val(ThisAmt)
        UpdatePay = True
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        NewPRate = EmpRec.EMPPRATE
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the pay rate for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(Using$("$#,##0.00", EmpRec.EMPPRATE)) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    vaSpread.Col = 11
    ThisOTRate = ReplaceString(vaSpread.Text, "$", "")
    ThisOTRate = ReplaceString(ThisOTRate, ",", "")
    OldORate = EmpRec.EMPORATE
    NewORate = Val(ThisOTRate)
    If Val(ThisOTRate) <> EmpRec.EMPORATE Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(ReplaceString(vaSpread.Text, "$", "")) <> "" And EmpRec.EMPORATE <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the OT rate on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using$("$#,##0.00", EmpRec.EMPORATE)) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(ReplaceString(vaSpread.Text, "$", "")) = "" And EmpRec.EMPORATE <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the OT rate on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using$("$#,##0.00", EmpRec.EMPORATE)) + " to '$0.00'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(ReplaceString(vaSpread.Text, "$", "")) <> "" And EmpRec.EMPORATE = 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the OT rate on row #" + CStr(vaSpread.Row) + " from '$0.00' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        Unload frmMessageWOpts
        Close
        Exit Sub
      ElseIf frmMessageW3Opts.fptxtChoice.Text = "continue" Then
        Unload frmMessageWOpts
        ThisAmt = ReplaceString$(vaSpread.Text, "$", "")
        ThisAmt = ReplaceString$(ThisAmt, ",", "")
        EmpRec.EMPORATE = Val(ThisAmt)
        UpdatePay = True
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        NewORate = EmpRec.EMPORATE
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the OT rate for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(Using$("$#,##0.00", EmpRec.EMPORATE)) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    vaSpread.Col = 12
    If EmpRec.EMPHDATE = 0 Then
      ThisDate = "BLANK"
    Else
      ThisDate = MakeRegDate(EmpRec.EMPHDATE)
    End If
    If QPTrim$(vaSpread.Text) <> MakeRegDate(EmpRec.EMPHDATE) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) = "" And EmpRec.EMPHDATE = 0 Then GoTo NoHireDate
      If QPTrim$(vaSpread.Text) <> "" And EmpRec.EMPHDATE <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the hire date on row #" + CStr(vaSpread.Row) + " from " + ThisDate + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And EmpRec.EMPHDATE <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the hire date on row #" + CStr(vaSpread.Row) + " from " + ThisDate + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And EmpRec.EMPHDATE = 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the hire date on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        Unload frmMessageWOpts
        Close
        Exit Sub
      ElseIf frmMessageW3Opts.fptxtChoice.Text = "continue" Then
        Unload frmMessageWOpts
        If QPTrim$(ReplaceString(vaSpread.Text, "/", "")) = "" Then
           EmpRec.EMPHDATE = 0
        Else
          EmpRec.EMPHDATE = Date2Num(vaSpread.Text)
        End If
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the hire date for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + ThisDate + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
NoHireDate:
    vaSpread.Col = 13
    If EmpRec.EMPRDATE = 0 Then
      ThisDate = "BLANK"
    Else
      ThisDate = MakeRegDate(EmpRec.EMPRDATE)
    End If
    If QPTrim$(vaSpread.Text) <> MakeRegDate(EmpRec.EMPRDATE) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) = "" And EmpRec.EMPRDATE = 0 Then GoTo NoReviewDate
      If QPTrim$(vaSpread.Text) <> "" And EmpRec.EMPRDATE <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the review date on row #" + CStr(vaSpread.Row) + " from " + ThisDate + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And EmpRec.EMPRDATE <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the review date on row #" + CStr(vaSpread.Row) + " from " + ThisDate + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And EmpRec.EMPRDATE = 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the review date on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        Unload frmMessageWOpts
        Close
        Exit Sub
      ElseIf frmMessageW3Opts.fptxtChoice.Text = "continue" Then
        Unload frmMessageWOpts
        If QPTrim$(ReplaceString(vaSpread.Text, "/", "")) = "" Then
          EmpRec.EMPRDATE = 0
        Else
          EmpRec.EMPRDATE = Date2Num(vaSpread.Text)
        End If
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the review date for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + MakeRegDate(EmpRec.EMPRDATE) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
NoReviewDate:
    vaSpread.Col = 14
    If EmpRec.EMPTDATE = 0 Then
      ThisDate = "BLANK"
    Else
      ThisDate = MakeRegDate(EmpRec.EMPTDATE)
    End If
      If QPTrim$(vaSpread.Text) = "" And EmpRec.EMPRDATE = 0 Then GoTo NoTermDate
    If QPTrim$(vaSpread.Text) <> MakeRegDate(EmpRec.EMPTDATE) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) = "" And EmpRec.EMPTDATE = 0 Then GoTo NoTermDate
      If QPTrim$(vaSpread.Text) <> "" And EmpRec.EMPTDATE <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the termination date on row #" + CStr(vaSpread.Row) + " from " + ThisDate + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And EmpRec.EMPTDATE <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the termination date on row #" + CStr(vaSpread.Row) + " from " + ThisDate + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And EmpRec.EMPTDATE = 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the termination date on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        Unload frmMessageWOpts
        Close
        Exit Sub
      ElseIf frmMessageW3Opts.fptxtChoice.Text = "continue" Then
        Unload frmMessageWOpts
        If QPTrim$(ReplaceString(vaSpread.Text, "/", "")) = "" Then
          EmpRec.EMPTDATE = 0
        Else
          EmpRec.EMPTDATE = Date2Num(vaSpread.Text)
        End If
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the termination date for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + ThisDate + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
NoTermDate:
   vaSpread.Col = 15
    If QPTrim$(UCase(vaSpread.Text)) <> QPTrim$(UCase(EmpRec.Comment)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.Comment) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the comment on row #" + QPTrim$(vaSpread.Row) + " from " + QPTrim$(EmpRec.Comment) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.Comment) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the comment on row #" + QPTrim$(vaSpread.Row) + " from " + QPTrim$(EmpRec.Comment) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.Comment) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the comment on row #" + QPTrim$(vaSpread.Row) + " from  'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        Unload frmMessageWOpts
        Close
        Exit Sub
      ElseIf frmMessageW3Opts.fptxtChoice.Text = "continue" Then
        Unload frmMessageWOpts
        EmpRec.Comment = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the comment for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.Comment) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    If UpdatePay = True Then
      UpdatePay = False
      frmLoadingRpt.Label1.Caption = "Updating Pay Records..."
      frmLoadingRpt.Show
      Call UpdatePayRateEscapeVrs(QPTrim$(EmpRec.EMPJOB), NewPType, NewORate, NewPRate, NewPFreq, OldPType, OldORate, OldPRate, OldPFreq, ThisRec)
      Unload frmLoadingRpt
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmEmpQuickMaintJobDesc", "cmdEscape_Click", Erl)
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

Private Sub cmdSave_Click()
  Dim EmpRec As EmpData2Type
  Dim EHandle As Integer
  Dim NumOfRows As Integer
  Dim x As Integer
  Dim ThisRec As Integer
  Dim ThisPAmt$
  Dim ThisOAmt$
  Dim ThisFreq$
  Dim ThisType$
  Dim UpDate As Boolean
  
  If ThisChange = 0 Then
    frmMessage.Label1.Caption = "No changes made. Save aborted."
    frmMessage.Label1.Top = 900
    frmMessage.Show vbModal
    GoTo NoChanges
  End If

  On Error GoTo ERRORSTUFF
  
  If RequiredFieldsOK = False Then
    vaSpread.OperationMode = OperationModeRow
    Unload frmLoadingRpt
    Exit Sub
  End If
  
  NumOfRows = vaSpread.MaxRows
  OpenEmpData2File EHandle
  
  For x = 1 To ThisChange 'NumOfRows
    UpDate = False
    vaSpread.Row = ChangeSpot(x)
    vaSpread.Col = 16
    ThisRec = vaSpread.Value
    Get EHandle, ThisRec, EmpRec
    vaSpread.Col = 4
    EmpRec.EMPJOB = QPTrim$(vaSpread.Text)
    vaSpread.Col = 5
    EmpRec.EMPWCCLS = QPTrim$(vaSpread.Text)
    vaSpread.Col = 6
    EmpRec.EMPSTATS = QPTrim$(vaSpread.Text)
    vaSpread.Col = 7
    EmpRec.EMPPTYPE = QPTrim$(vaSpread.Text)
    ThisType = QPTrim$(vaSpread.Text)
    If UCase(ThisType) <> UCase(EmpRec.EMPPTYPE) Then UpDate = True
    vaSpread.Col = 8
    If QPTrim$(UCase(EmpRec.EMPPFREQ)) <> QPTrim$(UCase(vaSpread.Text)) Then UpDate = True
    ThisFreq = QPTrim$(vaSpread.Text)
    EmpRec.EMPPFREQ = QPTrim$(vaSpread.Text)
    If UCase(ThisFreq) <> UCase(EmpRec.EMPPFREQ) Then UpDate = True
    vaSpread.Col = 9
    EmpRec.EMPBCODE = Val(ReplaceString(vaSpread.Text, "%", ""))
    vaSpread.Col = 10
    ThisPAmt = ReplaceString(vaSpread.Text, "$", "")
    ThisPAmt = ReplaceString(ThisPAmt, ",", "")
    If EmpRec.EMPPRATE <> Val(ThisPAmt) Then UpDate = True
    EmpRec.EMPPRATE = Val(ThisPAmt)
    vaSpread.Col = 11
    ThisOAmt = ReplaceString(vaSpread.Text, "$", "")
    ThisOAmt = ReplaceString(ThisOAmt, ",", "")
    If EmpRec.EMPORATE <> Val(ThisOAmt) Then UpDate = True
    EmpRec.EMPORATE = Val(ThisOAmt)
    vaSpread.Col = 12
    If ReplaceString(vaSpread.Text, "/", "") <> "" Then
      EmpRec.EMPHDATE = Date2Num(vaSpread.Text)
    Else
      EmpRec.EMPHDATE = 0
    End If
    vaSpread.Col = 13
    If ReplaceString(vaSpread.Text, "/", "") <> "" Then
      EmpRec.EMPRDATE = Date2Num(vaSpread.Text)
    Else
      EmpRec.EMPRDATE = 0
    End If
    vaSpread.Col = 14
    If ReplaceString(vaSpread.Text, "/", "") <> "" Then
      EmpRec.EMPTDATE = Date2Num(vaSpread.Text)
    Else
      EmpRec.EMPTDATE = 0
    End If
    vaSpread.Col = 15
    EmpRec.Comment = QPTrim$(vaSpread.Text)
    If UpDate = True Then
      frmLoadingRpt.Label1.Caption = "Updating Pay Rate..."
      frmLoadingRpt.Show
      DoEvents
      Call UpdatePayRate(QPTrim$(EmpRec.EMPJOB), ThisType, Val(ThisOAmt), Val(ThisPAmt), ThisFreq, ThisRec, False)
      Unload frmLoadingRpt
      DoEvents
      UpDate = False
      ThisType = ""
      ThisOAmt = 0
      ThisPAmt = 0
      ThisFreq = ""
    End If
    Put EHandle, ThisRec, EmpRec
  Next x
  
  Close
  
  MsgBox "Employee data has been saved successfully."
  
  ThisChange = 0
  
NoChanges:
  frmEmpQuickMaintMenu.Show
  DoEvents
  Unload Me
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmEmpQuickMaintJobDesc", "cmdSave_Click", Erl)
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
    Case vbKeyF10:
      SendKeys "%S"
      Call cmdSave_Click
      KeyCode = 0
    Case vbKeyF6:
      SendKeys "%F"
      Call cmdExitNow_Click
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
  
  GThisRow1 = 0
  ThisParameter = "All Employees"
  ThisTerm = chkTerm.Value
  BooBoo = False
  DontExit = False
  
  OpenEmpIdxLNameFile XHandle
  EmployeeCount = LOF(XHandle) / 2
  Close
  
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
      MainLog ("Payroll.exe terminated via menu bar on frmEmpQuickMaintJobDesc.")
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
  Dim RetRec As RetireRecType
  Dim RHandle As Integer
  Dim NumOfRetRecs As Integer
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
    vaSpread.Text = QPTrim$(EmpRec.EMPJOB)
    vaSpread.Col = 5
    vaSpread.Row = RowMax
    vaSpread.Text = QPTrim$(EmpRec.EMPWCCLS)
    vaSpread.Col = 6
    vaSpread.Row = RowMax
    vaSpread.Text = QPTrim$(EmpRec.EMPSTATS)
    vaSpread.Col = 7
    vaSpread.Row = RowMax
    vaSpread.Text = QPTrim$(EmpRec.EMPPTYPE)
    vaSpread.Col = 8
    vaSpread.Row = RowMax
    vaSpread.Text = QPTrim$(UCase(EmpRec.EMPPFREQ))
    vaSpread.Col = 9
    vaSpread.Row = RowMax
    vaSpread.Text = Using$("##0.00", EmpRec.EMPBCODE)
    vaSpread.Col = 10
    vaSpread.Row = RowMax
    vaSpread.Text = EmpRec.EMPPRATE
    vaSpread.Col = 11
    vaSpread.Row = RowMax
    vaSpread.Text = EmpRec.EMPORATE
    vaSpread.Col = 12
    vaSpread.Row = RowMax
    If EmpRec.EMPHDATE = 0 Then
      vaSpread.Text = ""
    Else
      vaSpread.Text = MakeRegDate(EmpRec.EMPHDATE)
    End If
'    vaSpread.TypePicMask = "99//99//9999"
    vaSpread.Col = 13
    vaSpread.Row = RowMax
    If EmpRec.EMPRDATE = 0 Then
      vaSpread.Text = ""
    Else
      vaSpread.Text = MakeRegDate(EmpRec.EMPRDATE)
    End If
'    vaSpread.TypePicMask = "99//99//9999"
    vaSpread.Col = 14
    vaSpread.Row = RowMax
    If EmpRec.EMPTDATE = 0 Then
      vaSpread.Text = ""
    Else
      vaSpread.Text = MakeRegDate(EmpRec.EMPTDATE)
    End If
'    vaSpread.TypePicMask = "99//99//9999"
    vaSpread.Col = 15
    vaSpread.Row = RowMax
    vaSpread.Text = QPTrim$(EmpRec.Comment)
    vaSpread.Col = 16
    vaSpread.Row = RowMax
    vaSpread.Text = CStr(ThisIdx(x))
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmEmpQuickMaintJobDesc", "LoadSpread", Erl)
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
  Call LoadSpread(ThisLoadSpread)

BooBooFound:

End Sub
Private Sub vaSpread_Click(ByVal Col As Long, ByVal Row As Long)
  vaSpread.OperationMode = OperationModeRow
End Sub

Private Sub vaSpread_DblClick(ByVal Col As Long, ByVal Row As Long)
  If Col < 4 Then
    MsgBox "This column is read only"
  End If
End Sub


Private Sub vaSpread_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
  'make the active row's back color yellow
  GThisRow1 = Row
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
  vaSpread.Col = 6
  vaSpread.BackColor = &HC0FFFF
  vaSpread.Col = 7
  vaSpread.BackColor = &HC0FFFF
  vaSpread.Col = 8
  vaSpread.BackColor = &HC0FFFF
  vaSpread.Col = 9
  vaSpread.BackColor = &HC0FFFF
  vaSpread.Col = 10
  vaSpread.BackColor = &HC0FFFF
  vaSpread.Col = 11
  vaSpread.BackColor = &HC0FFFF
  vaSpread.Col = 12
  vaSpread.BackColor = &HC0FFFF
  vaSpread.Col = 13
  vaSpread.BackColor = &HC0FFFF
  vaSpread.Col = 14
  vaSpread.BackColor = &HC0FFFF
End Sub

Private Sub vaSpread_KeyPress(KeyAscii As Integer)
  vaSpread.OperationMode = OperationModeRow
End Sub

Private Sub vaSpread_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'  If Col = 6 Then
'    Call SetRowColor(Row)
'  End If
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
  vaSpread.Col = 6
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 7
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 8
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 9
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 10
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 11
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 12
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 13
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 14
  vaSpread.BackColor = &H80000005

End Sub

Private Sub MakeRowWhite(RowNum As Integer)
  vaSpread.BackColorStyle = BackColorStyleUnderGrid
  vaSpread.Row = RowNum
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
  vaSpread.Col = 6
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 7
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 8
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 9
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 10
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 11
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 12
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 13
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 14
  vaSpread.BackColor = &H80000005

End Sub

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
  
  If Col = 6 Then
    vaSpread.Row = Row
    vaSpread.Col = 6
    If QPTrim$(UCase$(vaSpread.Text)) = UCase("Full-Time") Then
      vaSpread.Col = 9
      vaSpread.Text = "100.00"
    Else
      vaSpread.Col = 9
      vaSpread.Text = "0.00"
    End If
  End If
  
End Sub

Private Sub ClearChanges()
  ThisChange = 0
  ReDim ChangeSpot(0 To 0) As Integer
End Sub

Private Function RequiredFieldsOK() As Boolean
  Dim Emp2Rec As EmpData2Type
  Dim EmpCnt As Integer
  Dim EmpHandle As Integer
  Dim x As Integer
  Dim ThisRec As Integer
  
  RequiredFieldsOK = True
  OpenEmpData2File EmpHandle
  
  For x = 1 To ThisChange '  vaSpread.MaxRows
    vaSpread.Col = 16
    vaSpread.Row = ChangeSpot(x)
    ThisRec = vaSpread.Value
    Get EmpHandle, ThisRec, Emp2Rec
    vaSpread.Col = 5
    If QPTrim$(vaSpread.Text) = "" Then
      frmMessage.Label1.Caption = "W/C Code is a required field. Please supply this data on row # " + CStr(vaSpread.Row) + "."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      Close
      RequiredFieldsOK = False
      Exit Function
    End If
    vaSpread.Col = 6
    If QPTrim$(vaSpread.Text) = "" Then
      frmMessage.Label1.Caption = "Status is a required field. Please supply this data on row # " + CStr(vaSpread.Row) + "."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      Close
      RequiredFieldsOK = False
      Exit Function
    End If
    vaSpread.Col = 7
    If QPTrim$(vaSpread.Text) = "" Then
      frmMessage.Label1.Caption = "PayType is a required field. Please supply this data on row # " + CStr(vaSpread.Row) + "."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      Close
      RequiredFieldsOK = False
      Exit Function
    End If
    vaSpread.Col = 8
    If QPTrim$(ReplaceString(vaSpread.Text, "-", "")) = "" Then
      frmMessage.Label1.Caption = "Frequency is a required field. Please supply this data on row # " + CStr(vaSpread.Row) + "."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      Close
      RequiredFieldsOK = False
      Exit Function
    End If
    vaSpread.Col = 9
    If QPTrim$(ReplaceString(vaSpread.Text, "-", "")) = "" Then
      frmMessage.Label1.Caption = "Benefit Pct is a required field. Please supply this data on row # " + CStr(vaSpread.Row) + "."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      Close
      RequiredFieldsOK = False
      Exit Function
    End If
    vaSpread.Col = 10
    If QPTrim$(vaSpread.Text) = "" Then
      frmMessage.Label1.Caption = "Pay Rate is a required field. Please supply this data on row # " + CStr(vaSpread.Row) + "."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      Close
      RequiredFieldsOK = False
      Exit Function
    End If
  Next x
  
  Close EmpHandle
  
End Function

'Public Sub SetRowColor(ThisRow As Long)
'  vaSpread.BackColorStyle = BackColorStyleUnderGrid
'  vaSpread.Row = ThisRow
'  vaSpread.Col = 1
'  vaSpread.BackColor = &H80000005
'  vaSpread.Col = 2
'  vaSpread.BackColor = &H80000005
'  vaSpread.Col = 3
'  vaSpread.BackColor = &H80000005
'  vaSpread.Col = 4
'  vaSpread.BackColor = &H80000005
'  vaSpread.Col = 5
'  vaSpread.BackColor = &H80000005
'  vaSpread.Col = 6
'  vaSpread.BackColor = &H80000005
'  vaSpread.Col = 7
'  vaSpread.BackColor = &H80000005
'  vaSpread.Col = 8
'  vaSpread.BackColor = &H80000005
'  vaSpread.Col = 9
'  vaSpread.BackColor = &H80000005
'  vaSpread.Col = 10
'  vaSpread.BackColor = &H80000005
'  vaSpread.Col = 11
'  vaSpread.BackColor = &H80000005
'  vaSpread.Col = 12
'  vaSpread.BackColor = &H80000005
'  vaSpread.Col = 13
'  vaSpread.BackColor = &H80000005
'  vaSpread.Col = 14
'  vaSpread.BackColor = &H80000005
'
'
'End Sub
