VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmEmpQuickMaintDirDep 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll: Quick Employee Maintenance"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmEmpQuickMaintDirDep.frx":0000
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
      ColDesigner     =   "frmEmpQuickMaintDirDep.frx":08CA
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1680
      Width           =   3495
   End
   Begin FPSpread.vaSpread vaSpread 
      Height          =   3975
      Left            =   660
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3120
      Width           =   10335
      _Version        =   196613
      _ExtentX        =   18230
      _ExtentY        =   7011
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
      MaxCols         =   10
      MaxRows         =   1000000
      OperationMode   =   2
      ProcessTab      =   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   12648447
      SpreadDesigner  =   "frmEmpQuickMaintDirDep.frx":0B89
      VisibleCols     =   9
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   690
      Left            =   7673
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Press F10 to save all the changes made on this spreadsheet."
      Top             =   7560
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
      ButtonDesigner  =   "frmEmpQuickMaintDirDep.frx":1BFDE
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
      Height          =   690
      Left            =   4883
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Press ESC to exit this screen with each cell examined for unsaved changes."
      Top             =   7560
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
      ButtonDesigner  =   "frmEmpQuickMaintDirDep.frx":1C1BA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExitNow 
      Height          =   690
      Left            =   2078
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Press ESC to exit this screen after each cell is examined for unsaved changes."
      Top             =   7560
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
      ButtonDesigner  =   "frmEmpQuickMaintDirDep.frx":1C396
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
      Left            =   3180
      TabIndex        =   8
      Top             =   2400
      Width           =   2670
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   540
      Top             =   3000
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
      TabIndex        =   2
      Top             =   510
      Width           =   4995
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   3293
      Top             =   360
      Width           =   5055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Direct Deposit Info"
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
      TabIndex        =   0
      Top             =   870
      Width           =   3315
   End
End
Attribute VB_Name = "frmEmpQuickMaintDirDep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim EmployeeCount As Integer
  Dim ThisLoadSpread As Integer
  Dim ChangeSpot() As Integer
  Dim ThisChange As Integer
  Dim ThisParameter$
  Dim ThisTerm As Integer
  Dim BooBoo As Boolean
  Dim DontExit As Boolean
  Dim GlobalCode$
  
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
  Dim x As Integer
  Dim ThisRec As Integer
  Dim ThisText$
  Dim BDCode$
  
  If ThisChange = 0 Then GoTo NoChanges
  
  On Error GoTo ERRORSTUFF
  
  OpenEmpData2File EHandle
  
  For x = 1 To ThisChange
    vaSpread.Row = ChangeSpot(x)
    vaSpread.Col = 10
    ThisRec = CInt(vaSpread.Text)
    Get EHandle, ThisRec, EmpRec
    vaSpread.Col = 4
    BDCode = QPTrim$(vaSpread.Text)
    ThisText = QPTrim$(EmpRec.DRAFTCOD)
    If QPTrim$(UCase(vaSpread.Text)) <> QPTrim$(UCase(EmpRec.DRAFTCOD)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.DRAFTCOD) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the bank draft code on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.DRAFTCOD) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.DRAFTCOD) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the bank draft code on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.DRAFTCOD) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.DRAFTCOD) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the bank draft code on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        Unload frmMessageW3Opts
        Close
        Exit Sub
      ElseIf frmMessageW3Opts.fptxtChoice.Text = "continue" Then
        Unload frmMessageW3Opts
        vaSpread.Col = 5
        If QPTrim$(vaSpread.Text) = "" And BDCode <> "" Then
          frmMessage.Label1.Caption = "Since there is a value for 'Bank Draft Code' a 'Bank Account Number' is required."
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
        vaSpread.Col = 6
        If QPTrim$(vaSpread.Text) = "" And BDCode <> "" Then
          frmMessage.Label1.Caption = "Since there is a value for 'Bank Draft Code' a 'Prenoted' value is required."
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
        vaSpread.Col = 7
        If QPTrim$(vaSpread.Text) = "" And BDCode <> "" Then
          frmMessage.Label1.Caption = "Since there is a value for 'Bank Draft Code' a 'Bank Name' is required."
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
        vaSpread.Col = 8
        If QPTrim$(vaSpread.Text) = "" And BDCode <> "" Then
          frmMessage.Label1.Caption = "Since there is a value for 'Bank Draft Code' a 'Bank Location' is required."
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
        vaSpread.Col = 9
        If QPTrim$(vaSpread.Text) = "" And BDCode <> "" Then
          frmMessage.Label1.Caption = "Since there is a value for 'Bank Draft Code' a 'Bank Transit Number' is required."
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
        vaSpread.Col = 4
        EmpRec.DRAFTCOD = QPTrim$(vaSpread.Text)
        ThisText = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the bank draft code for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.DRAFTCOD) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    vaSpread.Col = 5
    If QPTrim$(UCase(vaSpread.Text)) <> QPTrim$(UCase(EmpRec.EMPDDACC)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPDDACC) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the bank account number on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPDDACC) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EMPDDACC) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the bank account number on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPDDACC) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPDDACC) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the bank account number on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        Unload frmMessageW3Opts
        Close
        Exit Sub
      ElseIf frmMessageW3Opts.fptxtChoice.Text = "continue" Then
        Unload frmMessageWOpts
        EmpRec.EMPDDACC = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made the bank account number for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EMPDDACC) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    vaSpread.Col = 6
    If QPTrim$(UCase(vaSpread.Text)) <> QPTrim$(UCase(EmpRec.PRENOTED)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.PRENOTED) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the prenoted field on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.PRENOTED) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.PRENOTED) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the prenoted field on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.PRENOTED) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.PRENOTED) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the prenoted field on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        Unload frmMessageW3Opts
        Close
        Exit Sub
      ElseIf frmMessageW3Opts.fptxtChoice.Text = "continue" Then 'save it
        Unload frmMessageW3Opts
        EmpRec.PRENOTED = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageW3Opts
        MainLog ("User warned that a change was made the prenoted field for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.PRENOTED) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    vaSpread.Col = 7
    If QPTrim$(UCase(vaSpread.Text)) <> QPTrim$(UCase(EmpRec.BankName)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.BankName) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the bank name on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.BankName) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.BankName) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the bank name on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.BankName) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.BankName) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the bank name on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        Unload frmMessageW3Opts
        Close
        Exit Sub
      ElseIf frmMessageW3Opts.fptxtChoice.Text = "continue" Then
        Unload frmMessageW3Opts
        EmpRec.BankName = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageW3Opts
        MainLog ("User warned that a change was made the bank name for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.BankName) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    vaSpread.Col = 8
    If QPTrim$(UCase(vaSpread.Text)) <> QPTrim$(UCase(EmpRec.BANKLOC)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.BANKLOC) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the bank location on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.BANKLOC) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.BANKLOC) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the bank location on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.BANKLOC) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.BANKLOC) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the bank location on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        Unload frmMessageW3Opts
        Close
        Exit Sub
      ElseIf frmMessageW3Opts.fptxtChoice.Text = "continue" Then
        Unload frmMessageW3Opts
        EmpRec.BANKLOC = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageW3Opts
        MainLog ("User warned that a change was made the bank location for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.BANKLOC) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    vaSpread.Col = 9
    If QPTrim$(UCase(vaSpread.Text)) <> QPTrim$(UCase(EmpRec.TRANSIT)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.TRANSIT) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the bank transit number on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.TRANSIT) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.TRANSIT) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the bank transit number on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.TRANSIT) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.TRANSIT) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the bank transit number on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        Unload frmMessageW3Opts
        Close
        Exit Sub
      ElseIf frmMessageW3Opts.fptxtChoice.Text = "continue" Then
        Unload frmMessageW3Opts
        EmpRec.TRANSIT = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageW3Opts
        MainLog ("User warned that a change was made the bank transit number for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.TRANSIT) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmEmpQuickMaintDirDep", "cmdEscape_Click", Erl)
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
  
  If ThisChange = 0 Then
    frmMessage.Label1.Caption = "No changes made. Save aborted."
    frmMessage.Label1.Top = 900
    frmMessage.Show vbModal
    GoTo NoChanges
  End If

  On Error GoTo ERRORSTUFF
  
  frmLoadingRpt.Label1.Caption = "Saving......"
  frmLoadingRpt.Show
  DoEvents
  
  If CheckFields = False Then
    Unload frmLoadingRpt
    Exit Sub
  End If
  
  NumOfRows = vaSpread.MaxRows
  OpenEmpData2File EHandle
  
  For x = 1 To NumOfRows
    vaSpread.Row = x
    vaSpread.Col = 10
    ThisRec = vaSpread.Value
    Get EHandle, ThisRec, EmpRec
    vaSpread.Col = 4
    EmpRec.DRAFTCOD = QPTrim$(vaSpread.Text)
    vaSpread.Col = 5
    EmpRec.EMPDDACC = QPTrim$(vaSpread.Text)
    vaSpread.Col = 6
    EmpRec.PRENOTED = QPTrim$(vaSpread.Text)
    vaSpread.Col = 7
    EmpRec.BankName = QPTrim$(vaSpread.Text)
    vaSpread.Col = 8
    EmpRec.BANKLOC = QPTrim$(vaSpread.Text)
    vaSpread.Col = 9
    EmpRec.TRANSIT = QPTrim$(vaSpread.Text)
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmEmpQuickMaintDirDep", "cmdSave_Click", Erl)
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
      SendKeys "%C"
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
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub LoadMe()
  Dim IdxRec As NameSortIdxType
  Dim XHandle As Integer
  
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
      MainLog ("Payroll.exe terminated via menu bar on frmEmpQuickMaintDirDep.")
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
  Dim noCode As Boolean
  On Error GoTo ERRORSTUFF
  
  noCode = True
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
    noCode = True
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
    vaSpread.Text = QPTrim$(EmpRec.DRAFTCOD)
    If QPTrim$(EmpRec.DRAFTCOD) = "" Then
      noCode = True
    Else
      noCode = False
    End If
    vaSpread.Col = 5
    vaSpread.Row = RowMax
    If noCode = True Then
      vaSpread.Lock = True
    End If
    vaSpread.Text = QPTrim$(EmpRec.EMPDDACC)
    vaSpread.Col = 6
    vaSpread.Row = RowMax
    If noCode = True Then
      vaSpread.Lock = True
    End If
    vaSpread.Text = QPTrim$(EmpRec.PRENOTED)
    vaSpread.Col = 7
    vaSpread.Row = RowMax
    If noCode = True Then
      vaSpread.Lock = True
    End If
    vaSpread.Text = QPTrim$(EmpRec.BankName)
    vaSpread.Col = 8
    vaSpread.Row = RowMax
    If noCode = True Then
      vaSpread.Lock = True
    End If
    vaSpread.Text = QPTrim$(EmpRec.BANKLOC)
    vaSpread.Col = 9
    vaSpread.Row = RowMax
    If noCode = True Then
      vaSpread.Lock = True
    End If
    vaSpread.Text = QPTrim$(EmpRec.TRANSIT)
    vaSpread.Col = 10
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmEmpQuickMaintDirDep", "LoadSpread", Erl)
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
  
  If Col > 4 Then
    vaSpread.Col = 4
    vaSpread.Row = Row
    If QPTrim$(vaSpread.Text) = "" Then
      MsgBox "No access to this cell until 'Bank Draft Code' is assigned a value."
    End If
  End If
End Sub

Private Sub vaSpread_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
  'make the active row's back color yellow
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
End Sub

Private Sub vaSpread_EnterRow(ByVal Row As Long, ByVal RowIsLast As Long)
  vaSpread.Col = 4
  vaSpread.Row = Row
  GlobalCode = QPTrim$(vaSpread.Text)
End Sub

Private Sub vaSpread_KeyPress(KeyAscii As Integer)
  vaSpread.OperationMode = OperationModeRow
End Sub

Private Sub vaSpread_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
  Dim IsCode As Boolean
  Dim T5$, T6$, T7$, T8$, T9$
  
  vaSpread.Col = 4
  vaSpread.Row = Row
  If QPTrim$(vaSpread.Text) <> "" Then
    IsCode = True
  Else
    IsCode = False
  End If
    
'  If IsCode = True Then
'    vaSpread.Col = 5
'    T5 = QPTrim$(vaSpread.Text)
'    vaSpread.Col = 6
'    T6 = QPTrim$(vaSpread.Text)
'    vaSpread.Col = 7
'    T7 = QPTrim$(vaSpread.Text)
'    vaSpread.Col = 8
'    T8 = QPTrim$(vaSpread.Text)
'    vaSpread.Col = 9
'    T9 = QPTrim$(vaSpread.Text)
'    If T5 = "" Or T6 = "" Or T7 = "" Or T8 = "" Or T9 = "" Then
'      frmMessage.Label1.Caption = "If a value for 'Bank Draft Code' exists then all bank draft fields for that row must also be assigned values. Please make sure all fields for this employee have values."
'      frmMessage.Label1.Top = 750
'      frmMessage.Show vbModal
'      Unload frmMessage
'      If T5 = "" Then
'        vaSpread.Col = 5
'        vaSpread.SetFocus
'        vaSpread.OperationMode = OperationModeNormal
'        vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
'      ElseIf T6 = "" Then
'        vaSpread.Col = 6
'        vaSpread.SetFocus
'        vaSpread.OperationMode = OperationModeNormal
'        vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
'      ElseIf T7 = "" Then
'        vaSpread.Col = 7
'        vaSpread.SetFocus
'        vaSpread.OperationMode = OperationModeNormal
'        vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
'      ElseIf T8 = "" Then
'        vaSpread.Col = 8
'        vaSpread.SetFocus
'        vaSpread.OperationMode = OperationModeNormal
'        vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
'      ElseIf T9 = "" Then
'        vaSpread.Col = 9
'        vaSpread.SetFocus
'        vaSpread.OperationMode = OperationModeNormal
'        vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
'      End If
'    End If
'  End If
    
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

End Sub

Private Sub vaSpread_Change(ByVal Col As Long, ByVal Row As Long)
  Dim x As Integer
  Dim Col4Text$
  Dim ThisText$
  Dim OtherVals As Boolean
  
  OtherVals = False
  For x = 1 To ThisChange
    If ChangeSpot(x) = Row Then 'this row already has a change on it
      Exit For
    End If
  Next x
  
  If x > ThisChange Then
    ThisChange = ThisChange + 1
    ChangeSpot(ThisChange) = Row
  End If
  
  If Col = 6 Then
    vaSpread.Col = 6
    vaSpread.Row = Row
    Select Case QPTrim$(vaSpread.Text)
      Case "", "Y", "N"
      Case "y"
        vaSpread.Text = "Y"
      Case "n"
        vaSpread.Text = "N"
      Case Else
        vaSpread.Text = ""
    End Select
  End If
  
  If Col = 4 Then
    vaSpread.Col = 4
    vaSpread.Row = Row
    Select Case QPTrim$(vaSpread.Text)
      Case "", "C", "S"
      Case "c"
        vaSpread.Text = "C"
      Case "s"
        vaSpread.Text = "S"
      Case Else
        vaSpread.Text = ""
    End Select
    If QPTrim$(vaSpread.Text) = "" Then
      vaSpread.Col = 5
      If QPTrim$(vaSpread.Text) <> "" Then
         OtherVals = True
         GoTo AskOK
      End If
      vaSpread.Col = 6
      If QPTrim$(vaSpread.Text) <> "" Then
         OtherVals = True
         GoTo AskOK
      End If
      vaSpread.Col = 7
      If QPTrim$(vaSpread.Text) <> "" Then
         OtherVals = True
         GoTo AskOK
      End If
      vaSpread.Col = 8
      If QPTrim$(vaSpread.Text) <> "" Then
         OtherVals = True
         GoTo AskOK
      End If
      vaSpread.Col = 9
      If QPTrim$(vaSpread.Text) <> "" Then
         OtherVals = True
         GoTo AskOK
      End If
AskOK:
      If OtherVals = False Then
        vaSpread.Col = 5
        vaSpread.Lock = True
        vaSpread.Col = 6
        vaSpread.Lock = True
        vaSpread.Col = 7
        vaSpread.Lock = True
        vaSpread.Col = 8
        vaSpread.Lock = True
        vaSpread.Col = 9
        vaSpread.Lock = True
        GoTo DoneHere
      End If
      frmMessageWOpts.Label1.Caption = "Since there is no value for 'Bank Draft Code' there is no need for any other bank draft values for this employee. OK to delete unnecessary values?"
      frmMessageWOpts.Label1.Top = 800
      frmMessageWOpts.cmdCont.Text = "F10 Delete"
      frmMessageWOpts.cmdExit.Text = "ESC Exit"
      frmMessageWOpts.Show vbModal
      If frmMessageWOpts.fptxtChoice.Text = "continue" Then
        Unload frmMessageWOpts
        vaSpread.Col = 5
        vaSpread.Text = ""
        vaSpread.Lock = True
        vaSpread.Col = 6
        vaSpread.Text = ""
        vaSpread.Lock = True
        vaSpread.Col = 7
        vaSpread.Text = ""
        vaSpread.Lock = True
        vaSpread.Col = 8
        vaSpread.Text = ""
        vaSpread.Lock = True
        vaSpread.Col = 9
        vaSpread.Text = ""
        vaSpread.Lock = True
      Else
        vaSpread.Col = 4
        vaSpread.SetFocus
        vaSpread.OperationMode = OperationModeNormal
        vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
        vaSpread.Text = GlobalCode
      End If
    Else
      vaSpread.Col = 5
      vaSpread.Lock = False
      vaSpread.Col = 6
      vaSpread.Lock = False
      vaSpread.Col = 7
      vaSpread.Lock = False
      vaSpread.Col = 8
      vaSpread.Lock = False
      vaSpread.Col = 9
      vaSpread.Lock = False
    End If
  End If
  
DoneHere:
  vaSpread.Col = 4
  GlobalCode = QPTrim$(vaSpread.Text)
End Sub

Private Sub ClearChanges()
  ThisChange = 0
  ReDim ChangeSpot(0 To 0) As Integer
End Sub

Private Function CheckFields() As Boolean
  Dim ThisText$
  Dim x As Integer
  
  CheckFields = True
  
  For x = 1 To ThisChange
    vaSpread.Col = 4
    vaSpread.Row = ChangeSpot(x)
    ThisText$ = QPTrim$(vaSpread.Text)
    vaSpread.Col = 5
    If ThisText = "" And QPTrim$(vaSpread.Text) <> "" Then
      frmMessageWOpts.Label1.Caption = "Since there is no value for 'Bank Draft Code' there is no need for a value for 'Bank Account Number'. OK to delete?"
      frmMessageWOpts.Label1.Top = 900
      frmMessageWOpts.cmdCont.Text = "F10 Delete"
      frmMessageWOpts.cmdExit.Text = "ESC Exit"
      frmMessageWOpts.Show vbModal
      If frmMessageWOpts.fptxtChoice.Text = "continue" Then
        Unload frmMessageWOpts
        vaSpread.Text = ""
      Else
        Unload frmMessageWOpts
        vaSpread.OperationMode = OperationModeNormal
        vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
        DontExit = False
        CheckFields = False
        Exit Function
      End If
    End If
    vaSpread.Col = 6
    If ThisText = "" And QPTrim$(vaSpread.Text) <> "" Then
      frmMessageWOpts.Label1.Caption = "Since there is no value for 'Bank Draft Code' there is no need for a value for 'Prenoted'. OK to delete?"
      frmMessageWOpts.Label1.Top = 900
      frmMessageWOpts.cmdCont.Text = "F10 Delete"
      frmMessageWOpts.cmdExit.Text = "ESC Exit"
      frmMessageWOpts.Show vbModal
      If frmMessageWOpts.fptxtChoice.Text = "continue" Then
        Unload frmMessageWOpts
        vaSpread.Text = ""
      Else
        Unload frmMessageWOpts
        vaSpread.OperationMode = OperationModeNormal
        vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
        DontExit = False
        CheckFields = False
        Exit Function
      End If
    End If
    vaSpread.Col = 7
    If ThisText = "" And QPTrim$(vaSpread.Text) <> "" Then
      frmMessageWOpts.Label1.Caption = "Since there is no value for 'Bank Draft Code' there is no need for a value for 'Bank Name'. OK to delete?"
      frmMessageWOpts.Label1.Top = 900
      frmMessageWOpts.cmdCont.Text = "F10 Delete"
      frmMessageWOpts.cmdExit.Text = "ESC Exit"
      frmMessageWOpts.Show vbModal
      If frmMessageWOpts.fptxtChoice.Text = "continue" Then
        Unload frmMessageWOpts
        vaSpread.Text = ""
      Else
        Unload frmMessageWOpts
        vaSpread.OperationMode = OperationModeNormal
        vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
        DontExit = False
        CheckFields = False
        Exit Function
      End If
    End If
    vaSpread.Col = 8
    If ThisText = "" And QPTrim$(vaSpread.Text) <> "" Then
      frmMessageWOpts.Label1.Caption = "Since there is no value for 'Bank Draft Code' there is no need for a value for 'Bank Location'. OK to delete?"
      frmMessageWOpts.Label1.Top = 900
      frmMessageWOpts.cmdCont.Text = "F10 Delete"
      frmMessageWOpts.cmdExit.Text = "ESC Exit"
      frmMessageWOpts.Show vbModal
      If frmMessageWOpts.fptxtChoice.Text = "continue" Then
        Unload frmMessageWOpts
        vaSpread.Text = ""
      Else
        Unload frmMessageWOpts
        vaSpread.OperationMode = OperationModeNormal
        vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
        DontExit = False
        CheckFields = False
        Exit Function
      End If
    End If
    vaSpread.Col = 9
    If ThisText = "" And QPTrim$(vaSpread.Text) <> "" Then
      frmMessageWOpts.Label1.Caption = "Since there is no value for 'Bank Draft Code' there is no need for a value for 'Bank Transit Number'. OK to delete?"
      frmMessageWOpts.Label1.Top = 900
      frmMessageWOpts.cmdCont.Text = "F10 Delete"
      frmMessageWOpts.cmdExit.Text = "ESC Exit"
      frmMessageWOpts.Show vbModal
      If frmMessageWOpts.fptxtChoice.Text = "continue" Then
        Unload frmMessageWOpts
        vaSpread.Text = ""
      Else
        Unload frmMessageWOpts
        vaSpread.OperationMode = OperationModeNormal
        vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
        DontExit = False
        CheckFields = False
        Exit Function
      End If
    End If
  Next x
        
  For x = 1 To ThisChange
    vaSpread.Col = 4
    vaSpread.Row = ChangeSpot(x)
    ThisText = QPTrim(vaSpread.Text)
    If ThisText = "" Then GoTo SkipIt
    vaSpread.Col = 5
    If QPTrim$(vaSpread.Text) = "" Then
      frmMessage.Label1.Caption = "Since there is a value for 'Bank Draft Code' a 'Bank Account Number' is required."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      DontExit = False
      CheckFields = False
      Exit Function
    End If
    vaSpread.Col = 6
    If QPTrim$(vaSpread.Text) = "" Then
      frmMessage.Label1.Caption = "Since there is a value for 'Bank Draft Code' a value for 'Prenoted' is required."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      DontExit = False
      CheckFields = False
      Exit Function
    End If
    vaSpread.Col = 7
    If QPTrim$(vaSpread.Text) = "" Then
      frmMessage.Label1.Caption = "Since there is a value for 'Bank Draft Code' a 'Bank Name' is required."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      DontExit = False
      CheckFields = False
      Exit Function
    End If
    vaSpread.Col = 8
    If QPTrim$(vaSpread.Text) = "" Then
      frmMessage.Label1.Caption = "Since there is a value for 'Bank Draft Code' then a 'Bank Location' is required."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      DontExit = False
      CheckFields = False
      Exit Function
    End If
    vaSpread.Col = 9
    If QPTrim$(vaSpread.Text) = "" Then
      frmMessage.Label1.Caption = "Since there is a value for 'Bank Draft Code' then a 'Bank Transit Number' is required."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      DontExit = False
      CheckFields = False
      Exit Function
    End If
SkipIt:
  Next x
  
End Function
