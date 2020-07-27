VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmEmpQuickMaintPers 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll: Quick Employee Maintenance"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   ForeColor       =   &H00000000&
   Icon            =   "frmEmpQuickMaint.frx":0000
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
      Left            =   6120
      TabIndex        =   1
      Top             =   2040
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
      ColDesigner     =   "frmEmpQuickMaint.frx":08CA
   End
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   360
   End
   Begin FPSpread.vaSpread vaSpread 
      Height          =   4095
      Left            =   660
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2880
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
      MaxCols         =   20
      MaxRows         =   1000000
      OperationMode   =   2
      ProcessTab      =   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   12648447
      SpreadDesigner  =   "frmEmpQuickMaint.frx":0B89
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
      Left            =   4073
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1560
      Width           =   3495
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   690
      Left            =   7673
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Press F10 to save all the changes made on this spreadsheet."
      Top             =   7440
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
      ButtonDesigner  =   "frmEmpQuickMaint.frx":1C351
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
      Height          =   690
      Left            =   4875
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Press ESC to exit this screen after each cell is examined for unsaved changes."
      Top             =   7440
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
      ButtonDesigner  =   "frmEmpQuickMaint.frx":1C52D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExitNow 
      Height          =   690
      Left            =   2078
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Press to exit this screen without testing each cell for unsaved changes."
      Top             =   7440
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
      ButtonDesigner  =   "frmEmpQuickMaint.frx":1C709
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
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   540
      Top             =   2760
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
      Left            =   3293
      TabIndex        =   7
      Top             =   2160
      Width           =   2670
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Data"
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
      Left            =   4163
      TabIndex        =   6
      Top             =   720
      Width           =   3315
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   3293
      Top             =   210
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
      TabIndex        =   3
      Top             =   360
      Width           =   4995
   End
End
Attribute VB_Name = "frmEmpQuickMaintPers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim EmployeeCount As Integer
  Dim GThisRow1 As Integer
  Dim BigEmpNum As String * 10
  Dim ThisLoadSpread As Integer
  Dim ChangeSpot() As Integer
  Dim ThisChange As Integer
  Dim ThisParameter$
  Dim ThisTerm As Integer
  Dim BooBoo As Boolean
  Dim DontExit As Boolean
  Dim LastCell As Integer

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

Private Sub cmdEscape_Click()
  Dim EmpRec As EmpData2Type
  Dim EHandle As Integer
  Dim NumOfRows As Integer
  Dim x As Integer
  Dim ThisRec As Integer
  Dim ThisSavedPhone$
  Dim ThisSpreadPhone$
  Dim ThisRNum$
  Dim ThisRType$
  Dim ChangeMade As Boolean
  
  If ThisChange = 0 Then GoTo NoChanges
  On Error GoTo ERRORSTUFF
  
  NumOfRows = vaSpread.MaxRows
  OpenEmpData2File EHandle
  
  For x = 1 To ThisChange
    ChangeMade = False
    vaSpread.Row = ChangeSpot(x)
    vaSpread.Col = 20
    ThisRec = CInt(vaSpread.Text)
    Get EHandle, ThisRec, EmpRec
    vaSpread.Col = 2
    If QPTrim$(vaSpread.Text) <> QPTrim$(EmpRec.EmpNo) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EmpNo) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the employee number on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EmpNo) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EmpNo) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the employee number on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EmpNo) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EmpNo) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the employee number on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
          Unload frmMessageWOpts
          DontExit = False
          Close
          Exit Sub
        End If
        If CheckThisEmpNum(QPTrim$(vaSpread.Text), vaSpread.Row) = False Then
          DontExit = False
          BooBoo = True
          If QPTrim$(fpcmbParameters.Text) <> ThisParameter Then
            fpcmbParameters.Text = ThisParameter
          End If
          If chkTerm.Value <> ThisTerm Then
            chkTerm.Value = ThisTerm
          End If
          Close
          Exit Sub
        End If
        EmpRec.EmpNo = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        ChangeMade = True
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the employee number for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EmpNo) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 3
    If QPTrim$(vaSpread.Text) <> QPTrim$(UCase(EmpRec.EmpLName)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EmpLName) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the employee last name on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EmpLName) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EmpLName) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the employee last name on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EmpLName) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EmpLName) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the employee last name on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
          Unload frmMessageWOpts
          DontExit = False
          Close
          Exit Sub
        End If
        EmpRec.EmpLName = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        ChangeMade = True
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the employee last name for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EmpLName) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 4
    If QPTrim$(vaSpread.Text) <> QPTrim$(UCase(EmpRec.EmpFName)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EmpFName) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the employee first name on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EmpFName) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EmpFName) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the employee first name on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EmpFName) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EmpFName) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the employee first name on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
          Close
          Exit Sub
        End If
        EmpRec.EmpFName = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        ChangeMade = True
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the employee first name for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EmpFName) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 5
    If QPTrim$(vaSpread.Text) <> QPTrim$(UCase(EmpRec.EmpAddr1)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EmpAddr1) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the address on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EmpAddr1) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EmpAddr1) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the address on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EmpAddr1) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EmpAddr1) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the address on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
          Unload frmMessageWOpts
          DontExit = False
          Close
          Exit Sub
        End If
        EmpRec.EmpAddr1 = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in address #1 for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EmpAddr1) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 6
    If QPTrim$(vaSpread.Text) <> QPTrim$(UCase(EmpRec.EMPADDR2)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPADDR2) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the address on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPADDR2) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EMPADDR2) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the address on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPADDR2) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPADDR2) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the address on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        EmpRec.EMPADDR2 = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in address #2 for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EMPADDR2) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 7
    If QPTrim$(vaSpread.Text) <> QPTrim$(UCase(EmpRec.EmpCity)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EmpCity) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the city on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EmpCity) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EmpCity) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the city on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EmpCity) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EmpCity) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the city on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
          Close
          Exit Sub
        End If
        EmpRec.EmpCity = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the city for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EmpCity) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 8
    If QPTrim$(vaSpread.Text) <> QPTrim$(UCase(EmpRec.EmpState)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EmpState) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the state on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EmpState) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EmpState) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the state on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EmpState) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EmpState) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the state on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
          Close
          Exit Sub
        End If
        EmpRec.EmpState = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the state for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EmpState) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 9
    If QPTrim$(ReplaceString(vaSpread.Text, "-", "")) <> QPTrim$(ReplaceString(EmpRec.EmpZip, "-", "")) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(ReplaceString(vaSpread.Text, "-", "")) <> "" And QPTrim$(ReplaceString(EmpRec.EmpZip, "-", "")) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the zip code on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EmpZip) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(ReplaceString(vaSpread.Text, "-", "")) = "" And QPTrim$(ReplaceString(EmpRec.EmpZip, "-", "")) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the zip code on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EmpZip) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(ReplaceString(vaSpread.Text, "-", "")) <> "" And QPTrim$(ReplaceString(EmpRec.EmpZip, "-", "")) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the zip code on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        If QPTrim$(ReplaceString(vaSpread.Text, "-", "")) = "" Then
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
          Close
          Exit Sub
        End If
        EmpRec.EmpZip = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the zip code for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EmpZip) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 10
    If QPTrim$(vaSpread.Text) <> QPTrim$(AddDashToSSN(EmpRec.EmpSSN)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(ReplaceString(vaSpread.Text, "-", "")) <> "" And QPTrim$(AddDashToSSN(EmpRec.EmpSSN)) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the social security number on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(AddDashToSSN(EmpRec.EmpSSN)) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(ReplaceString(vaSpread.Text, "-", "")) = "" And QPTrim$(AddDashToSSN(EmpRec.EmpSSN)) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the social security number on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(AddDashToSSN(EmpRec.EmpSSN)) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(ReplaceString(vaSpread.Text, "-", "")) <> "" And QPTrim$(AddDashToSSN(EmpRec.EmpSSN)) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the social security number on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        If QPTrim$(ReplaceString(vaSpread.Text, "-", "")) = "" Then
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
          Close
          Exit Sub
        End If
        EmpRec.EmpSSN = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the social security number for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(AddDashToSSN(EmpRec.EmpSSN)) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 11
    If Date2Num(vaSpread.Text) <> EmpRec.EMPBDAY Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And EmpRec.EMPBDAY <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the birth date on row #" + CStr(vaSpread.Row) + " from " + MakeRegDate(EmpRec.EMPBDAY) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And EmpRec.EMPBDAY <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the birth date on row #" + CStr(vaSpread.Row) + " from " + MakeRegDate(EmpRec.EMPBDAY) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And EmpRec.EMPBDAY = 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the birth date on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        EmpRec.EMPBDAY = Date2Num(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the birth date for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + MakeRegDate(EmpRec.EMPBDAY) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 12
    If QPTrim$(UCase(vaSpread.Text)) <> QPTrim$(UCase(EmpRec.EMPGENDR)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPGENDR) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the gender on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPGENDR) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EMPGENDR) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the gender on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPGENDR) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPGENDR) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the gender on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
          Close
          Exit Sub
        End If
        EmpRec.EMPGENDR = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the gender for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EMPGENDR) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 13
    If QPTrim$(vaSpread.Text) <> QPTrim$(UCase(EmpRec.EMPRACE)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPRACE) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the race on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPRACE) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EMPRACE) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the race on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPRACE) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPRACE) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the race on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        EmpRec.EMPRACE = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the race for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EMPRACE) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 14
    If QPTrim$(vaSpread.Text) <> QPTrim$(EmpRec.EMPRETNO) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPRETNO) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the retirement number on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPRETNO) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EMPRETNO) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the retirement number on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPRETNO) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPRETNO) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the retirement number on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        If Mid(vaSpread.Text, 1, 1) = "T" Then
          frmMessageWOpts.Label1.Caption = "Placing a 'T' in front of the retirement number indicates this employee is temporarily suspended from participating in any state retirement. Do you wish to continue saving anyway?"
          frmMessageWOpts.Label1.Top = 700
          frmMessageWOpts.cmdCont.Text = "F10 Save Anyway"
          frmMessageWOpts.cmdExit.Text = "ESC Don't Save"
          frmMessageWOpts.Show vbModal
          If frmMessageWOpts.fptxtChoice.Text = "abort" Then
            Unload frmMessageWOpts
            GoTo DontSaveRet
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
          Else
            Unload frmMessageWOpts
          End If
        End If
        ThisRNum = QPTrim$(vaSpread.Text)
        vaSpread.Col = 15
        If ThisRNum <> "" And QPTrim$(vaSpread.Text) = "" Then
          frmMessage.Label1.Caption = "The retirement number and the retirement type must both have values before either can be saved. Please select a retirement type."
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
          Close
          Exit Sub
        ElseIf ThisRNum = "" And QPTrim$(vaSpread.Text) <> "" Then
          frmMessage.Label1.Caption = "The retirement number and the retirement type must both have values before either can be saved. Please select a retirement type."
          frmMessage.Label1.Top = 800
          frmMessage.Show vbModal
          vaSpread.SetFocus
          vaSpread.OperationMode = OperationModeNormal
          vaSpread.SetActiveCell vaSpread.Col - 1, vaSpread.Row
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
          Close
          Exit Sub
        End If
        
        vaSpread.Col = 16
        EmpRec.EMPRETNO = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the retirement number for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EMPRETNO) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
DontSaveRet:
    vaSpread.Col = 14
    ThisRNum = QPTrim$(vaSpread.Text)
    vaSpread.Col = 15
    If QPTrim$(UCase(vaSpread.Text)) <> QPTrim$(UCase(EmpRec.EMPRETTP)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPRETTP) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the retirement type on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPRETTP) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EMPRETTP) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the retirement type on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPRETTP) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPRETTP) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the retirement type on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        If ThisRNum <> "" And QPTrim$(vaSpread.Text) = "" Then
          frmMessage.Label1.Caption = "The retirement number and the retirement type must both have values before either can be saved. Please enter a retirement number."
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
          Close
          Exit Sub
        ElseIf ThisRNum = "" And QPTrim$(vaSpread.Text) <> "" Then
          frmMessage.Label1.Caption = "The retirement number and the retirement type must both have values before either can be saved. Please enter a retirement number."
          frmMessage.Label1.Top = 800
          frmMessage.Show vbModal
          vaSpread.SetFocus
          vaSpread.OperationMode = OperationModeNormal
          vaSpread.SetActiveCell vaSpread.Col - 1, vaSpread.Row
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
          Close
          Exit Sub
        End If
        EmpRec.EMPRETTP = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the retirement type for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EMPRETTP) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 16
    If QPTrim$(vaSpread.Text) <> QPTrim$(UCase(EmpRec.EmrgncyCntctName)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EmrgncyCntctName) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the emergency contact name on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EmrgncyCntctName) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EmrgncyCntctName) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the emergency contact name on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EmrgncyCntctName) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EmrgncyCntctName) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the emergency contact name on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        EmpRec.EmrgncyCntctName = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the emergency contact name for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EmrgncyCntctName) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 17
    If QPTrim$(vaSpread.Text) <> QPTrim$(UCase(EmpRec.EmrgncyCntctRelation)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EmrgncyCntctRelation) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the emergency contact relation on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EmrgncyCntctRelation) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EmrgncyCntctRelation) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the emergency contact relation on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EmrgncyCntctRelation) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EmrgncyCntctRelation) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the emergency contact relation on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        EmpRec.EmrgncyCntctRelation = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the emergency contact relation for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EmrgncyCntctRelation) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 18
    ThisSavedPhone$ = QPTrim$(ReplaceString$(EmpRec.EmrgncyCntctPhnNum, "-", ""))
    ThisSavedPhone$ = QPTrim$(ReplaceString$(ThisSavedPhone, "(", ""))
    ThisSavedPhone$ = QPTrim$(ReplaceString$(ThisSavedPhone, ")", ""))
    If Val(ThisSavedPhone) = 0 Then ThisSavedPhone = ""
    ThisSpreadPhone$ = QPTrim$(ReplaceString$(vaSpread.Text, "-", ""))
    ThisSpreadPhone$ = QPTrim$(ReplaceString$(ThisSpreadPhone, "(", ""))
    ThisSpreadPhone$ = QPTrim$(ReplaceString$(ThisSpreadPhone, ")", ""))
    If Val(ThisSpreadPhone) = 0 Then ThisSpreadPhone = ""
    If ThisSpreadPhone <> ThisSavedPhone Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If ThisSpreadPhone <> "" And ThisSavedPhone <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the emergency contact phone number on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EmrgncyCntctPhnNum) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf ThisSpreadPhone = "" And ThisSavedPhone <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the emergency contact phone number on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EmrgncyCntctPhnNum) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf ThisSpreadPhone <> "" And ThisSavedPhone = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the emergency contact phone number on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + " To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        EmpRec.EmrgncyCntctPhnNum = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the emergency contact phone number for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EmrgncyCntctPhnNum) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 19
    ThisSavedPhone$ = QPTrim$(ReplaceString$(EmpRec.HomePhone, "-", ""))
    ThisSavedPhone$ = QPTrim$(ReplaceString$(ThisSavedPhone, "(", ""))
    ThisSavedPhone$ = QPTrim$(ReplaceString$(ThisSavedPhone, ")", ""))
    If Val(ThisSavedPhone) = 0 Then ThisSavedPhone = ""
    ThisSpreadPhone$ = QPTrim$(ReplaceString$(vaSpread.Text, "-", ""))
    ThisSpreadPhone$ = QPTrim$(ReplaceString$(ThisSpreadPhone, "(", ""))
    ThisSpreadPhone$ = QPTrim$(ReplaceString$(ThisSpreadPhone, ")", ""))
    If Val(ThisSpreadPhone) = 0 Then ThisSpreadPhone = ""
    If ThisSpreadPhone$ <> ThisSavedPhone$ Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If ThisSpreadPhone$ <> "" And ThisSavedPhone$ <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the home phone number on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.HomePhone) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf ThisSpreadPhone$ = "" And ThisSavedPhone$ <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the home phone number on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.HomePhone) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf ThisSpreadPhone$ <> "" And ThisSavedPhone$ = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the home phone number on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        EmpRec.HomePhone = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the home phone number for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EmrgncyCntctPhnNum) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    If ChangeMade = True Then
      Call MakeEmpIndexs
      ChangeMade = False
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmEmpQuickMaintPers", "cmdEscape_Click", Erl)
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
  Dim EmpRec1 As EmpData1Type '5/13/05
  Dim E1Handle As Integer '5/13/05
  Dim NumOfRows As Integer
  Dim x As Integer
  Dim ThisRec As Integer
  
  If ThisChange = 0 Then
    frmMessage.Label1.Caption = "No changes made. Save aborted."
    frmMessage.Label1.Top = 900
    frmMessage.Show vbModal
    GoTo NoChanges
  End If

  On Error GoTo ERRORSTUFF
  
  frmLoadingRpt.Label1.Caption = "Verifying Data..."
  frmLoadingRpt.Show
  DoEvents
  
  vaSpread.Col = 2
  For x = 1 To vaSpread.MaxRows
    vaSpread.Row = x
    If vaSpread.BackColor = &H8080FF Then vaSpread.BackColor = &H80000005
  Next x
  
  If RequiredFieldsOK = False Then
    vaSpread.OperationMode = OperationModeRow
    Unload frmLoadingRpt
    Exit Sub
  End If
  
  If CheckEmpNum = False Then
    Unload frmLoadingRpt
    Exit Sub
  End If
  
  If TestForRetNumT = False Then
    Unload frmLoadingRpt
    Exit Sub
  End If
  
  Unload frmLoadingRpt
  
  NumOfRows = vaSpread.MaxRows
  OpenEmpData2File EHandle
  OpenEmpData1File E1Handle
  
  FrmShowPctComp.Show , Me
  FrmShowPctComp.cmdCancel.Visible = False
  FrmShowPctComp.Label1.Caption = "Saving..."
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdExitNow.Enabled = False
  Me.cmdSave.Enabled = False
  
  For x = 1 To ThisChange 'NumOfRows
    vaSpread.Row = ChangeSpot(x) 'x
    vaSpread.Col = 20
    ThisRec = vaSpread.Value
    Get EHandle, ThisRec, EmpRec
    Get E1Handle, ThisRec, EmpRec1 '5/13/05
    vaSpread.Col = 2
    EmpRec.EmpNo = QPTrim$(vaSpread.Text)
    EmpRec1.EmpNo = QPTrim$(vaSpread.Text) '5/13/05
    vaSpread.Col = 3
    EmpRec.EmpLName = QPTrim$(vaSpread.Text)
    EmpRec1.EmpLName = QPTrim$(vaSpread.Text) '5/13/05
    vaSpread.Col = 4
    EmpRec.EmpFName = QPTrim$(vaSpread.Text)
    EmpRec1.EmpFName = QPTrim$(vaSpread.Text) '5/13/05
    vaSpread.Col = 5
    EmpRec.EmpAddr1 = QPTrim$(vaSpread.Text)
    vaSpread.Col = 6
    EmpRec.EMPADDR2 = QPTrim$(vaSpread.Text)
    vaSpread.Col = 7
    EmpRec.EmpCity = QPTrim$(vaSpread.Text)
    vaSpread.Col = 8
    EmpRec.EmpState = QPTrim$(vaSpread.Text)
    vaSpread.Col = 9
    EmpRec.EmpZip = QPTrim$(ReplaceString(vaSpread.Text, "-", ""))
    vaSpread.Col = 10
    EmpRec.EmpSSN = Mid(vaSpread.Text, 1, 3) + Mid(vaSpread.Text, 5, 2) + Mid(vaSpread.Text, 8, 4)
    vaSpread.Col = 11
    If QPTrim$(ReplaceString(vaSpread.Text, "/", "")) = "" Then
      EmpRec.EMPBDAY = 0
    Else
      EmpRec.EMPBDAY = Date2Num(QPTrim$(vaSpread.Text))
    End If
    vaSpread.Col = 12
    EmpRec.EMPGENDR = QPTrim$(vaSpread.Text)
    vaSpread.Col = 13
    EmpRec.EMPRACE = QPTrim$(vaSpread.Text)
    vaSpread.Col = 14
    EmpRec.EMPRETNO = QPTrim$(vaSpread.Text)
    vaSpread.Col = 15
    EmpRec.EMPRETTP = QPTrim$(vaSpread.Text)
    If CheckRet(EmpRec.EMPRETNO, EmpRec.EMPRETTP, vaSpread.Row) = False Then
      Close
      If QPTrim$(EmpRec.EMPRETNO) = "" Then
        vaSpread.OperationMode = OperationModeNormal
        vaSpread.SetActiveCell 14, vaSpread.Row
      ElseIf QPTrim$(EmpRec.EMPRETTP) = "" Then
        vaSpread.OperationMode = OperationModeNormal
        vaSpread.SetActiveCell 15, vaSpread.Row
      End If
      Exit Sub
    End If
    vaSpread.Col = 16
    EmpRec.EmrgncyCntctName = QPTrim$(vaSpread.Text)
    vaSpread.Col = 17
    EmpRec.EmrgncyCntctRelation = QPTrim$(vaSpread.Text)
    vaSpread.Col = 18
    EmpRec.EmrgncyCntctPhnNum = QPTrim$(vaSpread.Text)
    vaSpread.Col = 19
    EmpRec.HomePhone = QPTrim$(vaSpread.Text)
    Put EHandle, ThisRec, EmpRec
    Put E1Handle, ThisRec, EmpRec1 '5/13/05
    FrmShowPctComp.ShowPctComp x, ThisChange 'NumOfRows
  Next x
  
  Close
  Me.cmdExitNow.Enabled = True
  Me.cmdSave.Enabled = False
  Me.cmdEscape.Enabled = True
  Unload FrmShowPctComp
  
  MsgBox "Employee data has been saved successfully."
  
  ThisChange = 0
  
NoChanges:

  frmEmpQuickMaintMenu.Show
  DoEvents
  Unload Me
  
  Call MakeEmpIndexs
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmEmpQuickMaintPers", "cmdSave_Click", Erl)
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
  OpenEmpIdxLNameFile XHandle
  EmployeeCount = LOF(XHandle) / 2
  Close
  
  ThisLoadSpread = 1
  ThisParameter = "All Employees"
  ThisTerm = chkTerm.Value
  BooBoo = False
  DontExit = False
  
  fpcmbParameters.Text = "All Employees"
  fpcmbParameters.AddItem "All Employees"
  fpcmbParameters.AddItem "Full-Time"
  fpcmbParameters.AddItem "Part-Time"
  fpcmbParameters.AddItem "Seasonal"
  fpcmbParameters.AddItem "Temporary"
  
  ThisChange = 0
  
  Call LoadSpread(ThisLoadSpread)
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdEscape.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmEmpQuickMaintPers.")
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
  
  GoSub LoadRet
  
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
    vaSpread.Col = 2
    vaSpread.Row = RowMax
    vaSpread.Text = QPTrim$(EmpRec.EmpNo)
    vaSpread.Col = 3
    vaSpread.Row = RowMax
    vaSpread.Text = QPTrim$(EmpRec.EmpLName)
    vaSpread.Col = 4
    vaSpread.Row = RowMax
    vaSpread.Text = QPTrim$(EmpRec.EmpFName)
    vaSpread.Col = 5
    vaSpread.Row = RowMax
    vaSpread.Text = QPTrim$(EmpRec.EmpAddr1)
    vaSpread.Col = 6
    vaSpread.Row = RowMax
    vaSpread.Text = QPTrim$(EmpRec.EMPADDR2)
    vaSpread.Col = 7
    vaSpread.Row = RowMax
    vaSpread.Text = QPTrim$(EmpRec.EmpCity)
    vaSpread.Col = 8
    vaSpread.Row = RowMax
    vaSpread.Text = QPTrim$(EmpRec.EmpState)
    vaSpread.Col = 9
    vaSpread.Row = RowMax
    EmpRec.EmpZip = ReplaceString(EmpRec.EmpZip, "-", "")
    vaSpread.Text = Mid(QPTrim$(EmpRec.EmpZip), 1, 5) + "-" + Mid(QPTrim$(EmpRec.EmpZip), 6, 10)
'    vaSpread.TypePicMask = "99999-9999"
    vaSpread.Col = 10
    vaSpread.Row = RowMax
    vaSpread.Text = AddDashToSSN(QPTrim$(EmpRec.EmpSSN))
    vaSpread.Col = 11
    vaSpread.Row = RowMax
    vaSpread.Text = MakeRegDate(EmpRec.EMPBDAY)
'    vaSpread.TypePicMask = "99//99//9999"
    vaSpread.Col = 12
    vaSpread.Row = RowMax
    vaSpread.Text = QPTrim$(EmpRec.EMPGENDR)
    vaSpread.Col = 13
    vaSpread.Row = RowMax
    vaSpread.Text = QPTrim$(EmpRec.EMPRACE)
    vaSpread.Col = 14
    vaSpread.Row = RowMax
    vaSpread.Text = QPTrim$(EmpRec.EMPRETNO)
    vaSpread.Col = 15
    vaSpread.Row = RowMax
    vaSpread.Text = QPTrim$(EmpRec.EMPRETTP)
    vaSpread.Col = 16
    vaSpread.Row = RowMax
    vaSpread.Text = QPTrim$(EmpRec.EmrgncyCntctName)
    vaSpread.Col = 17
    vaSpread.Row = RowMax
    vaSpread.Text = QPTrim$(EmpRec.EmrgncyCntctRelation)
    vaSpread.Col = 18
    vaSpread.Row = RowMax
    vaSpread.TypePicMask = "(999)-999-9999"
    vaSpread.Text = QPTrim$(EmpRec.EmrgncyCntctPhnNum)
    vaSpread.Col = 19
    vaSpread.Row = RowMax
    vaSpread.TypePicMask = "(999)-999-9999"
    vaSpread.Text = QPTrim$(EmpRec.HomePhone)
    vaSpread.Col = 20
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
  
  Exit Sub
  
LoadRet:
  OpenRetFile RHandle
  NumOfRetRecs = LOF(RHandle) / Len(RetRec)
  If NumOfRetRecs = 0 Then
    MsgBox "No retirement records have been saved."
    Close RHandle
    GoTo NoRetRecs
  End If
  
  vaSpread.Col = 15
  For y = 1 To NumOfEmpRecs
    vaSpread.Row = y
    vaSpread.TypeComboBoxClear vaSpread.Col, vaSpread.Row
    For x = 1 To NumOfRetRecs
      Get RHandle, x, RetRec
        If QPTrim$(RetRec.TYPEDES1) = "" Then GoTo NoRec
        vaSpread.TypeComboBoxIndex = -1
        vaSpread.TypeComboBoxString = QPTrim$(RetRec.TYPEDES1)
NoRec:
    Next x
    vaSpread.TypeComboBoxIndex = -1
    vaSpread.TypeComboBoxString = " "
  Next y
  Close RHandle
  
NoRetRecs:

  vaSpread.OperationMode = OperationModeNormal
  vaSpread.SetActiveCell 2, 1

  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmEmpQuickMaintPers", "LoadSpread", Erl)
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
    DontExit = True 'prevents the program from exiting after
    'using the exit check when switching data (ex. Fulltime to Parttime)
    Call cmdEscape_Click
    If BooBoo = True Then 'a change was found during exit check so
    'return to screen because user wanted to review
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

Private Sub vaSpread_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
  'make the active row's back color yellow
  GThisRow1 = Row
  vaSpread.Row = Row
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
  vaSpread.Col = 15
  vaSpread.BackColor = &HC0FFFF
  vaSpread.Col = 16
  vaSpread.BackColor = &HC0FFFF
  vaSpread.Col = 17
  vaSpread.BackColor = &HC0FFFF
  vaSpread.Col = 18
  vaSpread.BackColor = &HC0FFFF
  vaSpread.Col = 19
  vaSpread.BackColor = &HC0FFFF
  vaSpread.Col = 20
  vaSpread.BackColor = &HC0FFFF
End Sub


Private Sub vaSpread_KeyPress(KeyAscii As Integer)
  vaSpread.OperationMode = OperationModeRow
End Sub

Private Sub vaSpread_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'  LastCell = Col
End Sub

Private Sub vaSpread_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
  vaSpread.BackColorStyle = BackColorStyleUnderGrid
  vaSpread.Row = Row
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
  vaSpread.Col = 15
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 16
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 17
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 18
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 19
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 20
  vaSpread.BackColor = &H80000005

End Sub

Private Function CheckThisEmpNum(thisNum$, ThisRow) As Boolean
  Dim Emp2Rec As EmpData2Type
  Dim EmpCnt As Integer
  Dim EmpHandle As Integer
  Dim x As Integer
  Dim Nextx As Integer
  Dim ThisRec As Integer
  Dim NextChange As Integer
  
  CheckThisEmpNum = True
  OpenEmpData2File EmpHandle
  vaSpread.Row = ThisRow
  vaSpread.Col = 20
  ThisRec = vaSpread.Value
  vaSpread.Col = 2
  thisNum = QPTrim$(vaSpread.Text)
  For x = 1 To EmployeeCount
    Get EmpHandle, x, Emp2Rec
    If Emp2Rec.Deleted = -1 Then GoTo SkipEmp
    If ThisRec <> x Then
      If QPTrim$(Emp2Rec.EmpNo) = thisNum Then
        MakeRowWhite (GThisRow1)
        vaSpread.SetFocus
        vaSpread.OperationMode = OperationModeNormal
        vaSpread.SetActiveCell 2, ThisRow
        CheckThisEmpNum = False
        MsgBox "This employee number is already in use. Please select another."
        Close EmpHandle
        Exit Function
      End If
    End If
SkipEmp:
  Next x
  
  Close EmpHandle
  
End Function

Private Function TestForRetNumT() As Boolean
  Dim x As Integer
  Dim ThisTCnt As Integer
  Dim FirstOne As Integer
  
  TestForRetNumT = True
  vaSpread.Col = 14
  ThisTCnt = 0
  FirstOne = 0
  ReDim TCnt(1 To 1) As Integer
  
  For x = 1 To ThisChange
    vaSpread.Row = ChangeSpot(x)
    If Mid(vaSpread.Text, 1, 1) = "T" Then
      ThisTCnt = ThisTCnt + 1
      If FirstOne = 0 Then
        FirstOne = ChangeSpot(x)
      End If
      ReDim Preserve TCnt(1 To ThisTCnt) As Integer
      TCnt(ThisTCnt) = x
    End If
  Next x
  
  If ThisTCnt > 0 Then
    vaSpread.Row = FirstOne
    vaSpread.SetFocus
    vaSpread.OperationMode = OperationModeNormal
    vaSpread.SetActiveCell 14, FirstOne
    If ThisTCnt > 1 Then
      frmMessageWOpts.Label1.Caption = "There are " + CStr(ThisTCnt) + " retirement numbers that begin with a 'T'. This 'T' (temporary) flag prevents the employee from participating in any state retirement programs. Please make sure these numbers are correct before continuing to save."
    ElseIf ThisTCnt = 1 Then
      frmMessageWOpts.Label1.Caption = "There is " + CStr(ThisTCnt) + " retirement number that begins with a 'T'. This 'T' (temporary) flag prevents the employee from participating in any state retirement programs. Please make sure this number is correct before continuing to save."
    End If
    frmMessageWOpts.Label1.Top = 700
    frmMessageWOpts.cmdCont.Text = "F10 Continue"
    frmMessageWOpts.cmdExit.Width = 2500
    frmMessageWOpts.cmdExit.Text = "ESC Stop and Review"
    frmMessageWOpts.Show vbModal
    If frmMessageWOpts.fptxtChoice.Text = "abort" Then
      Unload frmMessageWOpts
      TestForRetNumT = False
    Else
      Unload frmMessageWOpts
    End If
  End If
End Function

Private Function RequiredFieldsOK() As Boolean
  Dim Emp2Rec As EmpData2Type
  Dim EmpCnt As Integer
  Dim EmpHandle As Integer
  Dim x As Integer
  Dim ThisRec As Integer
  
  RequiredFieldsOK = True
  OpenEmpData2File EmpHandle
  
  For x = 1 To ThisChange
    vaSpread.Col = 20
    vaSpread.Row = ChangeSpot(x)
    ThisRec = vaSpread.Value
    Get EmpHandle, ThisRec, Emp2Rec
    vaSpread.Col = 2
    If QPTrim$(vaSpread.Text) = "" Then
      frmMessage.Label1.Caption = "An employee number is a required field. Please supply this data on row # " + CStr(vaSpread.Row(x)) + "."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      Close
      RequiredFieldsOK = False
      Exit Function
    End If
    vaSpread.Col = 3
    If QPTrim$(vaSpread.Text) = "" Then
      frmMessage.Label1.Caption = "An employee last name is a required field. Please supply this data on row # " + CStr(vaSpread.Row) + "."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      Close
      RequiredFieldsOK = False
      Exit Function
    End If
    vaSpread.Col = 4
    If QPTrim$(vaSpread.Text) = "" Then
      frmMessage.Label1.Caption = "An employee first name is a required field. Please supply this data on row # " + CStr(vaSpread.Row) + "."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      Close
      RequiredFieldsOK = False
      Exit Function
    End If
    vaSpread.Col = 5
    If QPTrim$(vaSpread.Text) = "" Then
      frmMessage.Label1.Caption = "An employee address is a required field. Please supply this data on row # " + CStr(vaSpread.Row) + "."
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
      frmMessage.Label1.Caption = "An employee city is a required field. Please supply this data on row # " + CStr(vaSpread.Row) + "."
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
    If QPTrim$(vaSpread.Text) = "" Then
      frmMessage.Label1.Caption = "An employee state is a required field. Please supply this data on row # " + CStr(vaSpread.Row) + "."
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
      frmMessage.Label1.Caption = "An employee zip code is a required field. Please supply this data on row # " + CStr(vaSpread.Row) + "."
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
    If QPTrim$(ReplaceString(vaSpread.Text, "-", "")) = "" Then
      frmMessage.Label1.Caption = "An employee social security number is a required field. Please supply this data on row # " + CStr(vaSpread.Row) + "."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      Close
      RequiredFieldsOK = False
      Exit Function
    End If
    vaSpread.Col = 12
    If QPTrim$(vaSpread.Text) = "" Then
      frmMessage.Label1.Caption = "An employee gender is a required field. Please supply this data on row # " + CStr(vaSpread.Row) + "."
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

Private Sub MakeRowWhite(RowNum As Integer)
  vaSpread.BackColorStyle = BackColorStyleUnderGrid
  vaSpread.Row = RowNum
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
  vaSpread.Col = 15
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 16
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 17
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 18
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 19
  vaSpread.BackColor = &H80000005
  vaSpread.Col = 20
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
  
End Sub

Private Sub ClearChanges()
  ThisChange = 0
  ReDim ChangeSpot(0 To 0) As Integer
End Sub

Private Function CheckEmpNum() As Boolean
  Dim Emp2Rec As EmpData2Type
  Dim EmpCnt As Integer
  Dim EmpHandle As Integer
  Dim x As Integer
  Dim thisNum$
  Dim Nextx As Integer
  Dim ThisRec As Integer
  Dim NextChange As Integer
  
  CheckEmpNum = True
  vaSpread.Col = 2
  Nextx = ChangeSpot(1)
  NextChange = 1
  OpenEmpData2File EmpHandle
  Do
    vaSpread.Row = Nextx
    vaSpread.Col = 20
    ThisRec = vaSpread.Value
    vaSpread.Col = 2
    thisNum = QPTrim$(vaSpread.Text)
    For x = 1 To EmployeeCount
      Get EmpHandle, x, Emp2Rec
      If Emp2Rec.Deleted = -1 Then GoTo SkipEmp
      If ThisRec <> x Then
        If Val(Mid(thisNum, 1, 1)) = 0 Then
          vaSpread.SetFocus
          vaSpread.OperationMode = OperationModeNormal
          vaSpread.SetActiveCell 2, Nextx
          CheckEmpNum = False
          MsgBox "Please enter another employee number that does not begin with a zero."
          Exit Function
        End If

        If QPTrim$(Emp2Rec.EmpNo) = thisNum Then
          MakeRowWhite (GThisRow1)
          vaSpread.OperationMode = OperationModeNormal
          vaSpread.SetFocus
          vaSpread.SetActiveCell 2, Nextx
          CheckEmpNum = False
          MsgBox "This employee number is already in use. Please select another."
          Close EmpHandle
          Exit Function
        End If
      End If
SkipEmp:
    Next x
    If NextChange = ThisChange Then Exit Do 'vaSpread.MaxRows
    NextChange = NextChange + 1
    Nextx = ChangeSpot(NextChange)
  Loop
  
  Close EmpHandle
  
End Function

Private Function CheckRet(RetNum$, RetType$, ThisRow As Integer) As Boolean
  
  CheckRet = True
  If QPTrim$(RetNum$) = "" And QPTrim$(RetType$) <> "" Then
    Unload FrmShowPctComp
    EnableCloseButton Me.hwnd, True
    Me.cmdEscape.Enabled = True
    Me.cmdExitNow.Enabled = True
    Me.cmdSave.Enabled = True
    frmMessage.Label1.Caption = "Both the retirement number and the retirement type must have values if either one has a value. Please enter a retirement number in the retirement number cell on row #" + CStr(ThisRow) + "."
    frmMessage.Label1.Top = 700
    frmMessage.Show vbModal
'    vaSpread.Col = 13
    CheckRet = False
    Exit Function
  ElseIf QPTrim$(RetNum$) <> "" And QPTrim$(RetType$) = "" Then
    Unload FrmShowPctComp
    EnableCloseButton Me.hwnd, True
    Me.cmdEscape.Enabled = True
    Me.cmdExitNow.Enabled = True
    Me.cmdSave.Enabled = True
    frmMessage.Label1.Caption = "Both the retirement number and the retirement type must have values if either one has a value. Please select a retirement type in the retirement type cell on row #" + CStr(ThisRow) + "."
    frmMessage.Label1.Top = 700
    frmMessage.Show vbModal
'    vaSpread.Col = 14
    CheckRet = False
    Exit Function
  End If
  
End Function

