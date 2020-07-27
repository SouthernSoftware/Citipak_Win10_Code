VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmEmpQuickMaintBenefits 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PayRoll: Quick Employee Maintenance"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmEmpQuickMaintBenefits.frx":0000
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
      Left            =   5986
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
      ColDesigner     =   "frmEmpQuickMaintBenefits.frx":08CA
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
      Left            =   4066
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1740
      Width           =   3495
   End
   Begin FPSpread.vaSpread vaSpread 
      Height          =   3975
      Left            =   720
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3120
      Width           =   10215
      _Version        =   196613
      _ExtentX        =   18018
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
      MaxCols         =   22
      MaxRows         =   1000000
      OperationMode   =   2
      ProcessTab      =   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   12648447
      SpreadDesigner  =   "frmEmpQuickMaintBenefits.frx":0B89
      VisibleCols     =   6
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
      ButtonDesigner  =   "frmEmpQuickMaintBenefits.frx":1C689
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
      Height          =   690
      Left            =   4883
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
      ButtonDesigner  =   "frmEmpQuickMaintBenefits.frx":1C865
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
      ButtonDesigner  =   "frmEmpQuickMaintBenefits.frx":1CA41
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
      Left            =   3346
      TabIndex        =   8
      Top             =   2320
      Width           =   2430
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   533
      Top             =   2940
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
      Left            =   3316
      TabIndex        =   7
      Top             =   570
      Width           =   4995
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   3286
      Top             =   420
      Width           =   5055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Benefit Schedule"
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
      Left            =   3953
      TabIndex        =   6
      Top             =   930
      Width           =   3795
   End
End
Attribute VB_Name = "frmEmpQuickMaintBenefits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim EmployeeCount As Integer
  Dim NumOfEarnRecs As Integer
  Dim ThisLoadSpread As Integer
  Dim ChangeSpot() As Integer
  Dim ThisChange As Integer
  Dim DontExit As Boolean
  Dim ThisParameter$
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
'  On Error GoTo ERRORSTUFF
  NumOfRows = vaSpread.MaxRows
  OpenEmpData2File EHandle

  For x = 1 To ThisChange
    vaSpread.Row = ChangeSpot(x)
    vaSpread.Col = 22
    ThisRec = CInt(vaSpread.Text)
    Get EHandle, ThisRec, EmpRec
    vaSpread.Col = 4
    If Val(vaSpread.Text) <> EmpRec.EMPVACE Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If Val(vaSpread.Text) <> 0 And EmpRec.EMPVACE <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in 'Vac Earned' field on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using$("###0.00", EmpRec.EMPVACE)) + " to " + QPTrim$(Using$("###0.00", Val(vaSpread.Text))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf Val(vaSpread.Text) = 0 And EmpRec.EMPVACE <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Vac Earned' field on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using$("###0.00", EmpRec.EMPVACE)) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf Val(vaSpread.Text) <> 0 And EmpRec.EMPVACE = 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Vac Earned' field on row #" + CStr(vaSpread.Row) + " from '0' to " + QPTrim$(Using$("###0.00", Val(vaSpread.Text))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        EmpRec.EMPVACE = Val(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your change has been saved successfully"
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the 'Vac Earned' field for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(Using$("###0.00", EmpRec.EMPVACE)) + " to " + QPTrim$(Using$("###0.00", vaSpread.Text)) + " but declined to save it.")
      End If
      ThisAcct = QPTrim$(vaSpread.Text)
    End If
    vaSpread.Col = 5
    If Val(vaSpread.Text) <> EmpRec.EMPVUSED Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If Val(vaSpread.Text) <> 0 And EmpRec.EMPVUSED <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Vac Used' on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("###0.00", EmpRec.EMPVUSED)) + " to " + QPTrim$(Using("###0.00", Val(vaSpread.Text))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf Val(vaSpread.Text) = 0 And EmpRec.EMPVUSED <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Vac Used' on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("###0.00", EmpRec.EMPVUSED)) + " to '0'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf Val(vaSpread.Text) <> 0 And EmpRec.EMPVUSED = 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Vac Used' on row #" + CStr(vaSpread.Row) + " from '0' to " + QPTrim$(Using("###0.00", Val(vaSpread.Text))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        EmpRec.EMPVUSED = Val(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your change has been saved successfully"
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the 'Vac Used' amount for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(Using("###0.00", EmpRec.EMPVUSED)) + " to " + QPTrim$(Using("###0.00", Val(vaSpread.Text))) + " but declined to save it.")
      End If
    End If
  
    vaSpread.Col = 7
    If Val(vaSpread.Text) <> EmpRec.EMPSLE Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If Val(vaSpread.Text) <> 0 And EmpRec.EMPSLE <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in 'Sick Earned' field on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using$("###0.00", EmpRec.EMPSLE)) + " to " + QPTrim$(Using$("###0.00", Val(vaSpread.Text))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf Val(vaSpread.Text) = 0 And EmpRec.EMPSLE <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Sick Earned' field on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using$("###0.00", EmpRec.EMPSLE)) + " to '0'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf Val(vaSpread.Text) <> 0 And EmpRec.EMPSLE = 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Sick Earned' field on row #" + CStr(vaSpread.Row) + " from '0' to " + QPTrim$(Using$("###0.00", Val(vaSpread.Text))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        EmpRec.EMPSLE = Val(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your change has been saved successfully"
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the 'Sick Earned' field for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(Using$("###0.00", EmpRec.EMPSLE)) + " to " + QPTrim$(Using$("###0.00", vaSpread.Text)) + " but declined to save it.")
      End If
      ThisAcct = QPTrim$(vaSpread.Text)
    End If
    vaSpread.Col = 8
    If Val(vaSpread.Text) <> EmpRec.EMPSLUSE Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If Val(vaSpread.Text) <> 0 And EmpRec.EMPSLUSE <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Sick Used' on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("###0.00", EmpRec.EMPSLUSE)) + " to " + QPTrim$(Using("###0.00", Val(vaSpread.Text))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf Val(vaSpread.Text) = 0 And EmpRec.EMPSLUSE <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Sick Used' on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("###0.00", EmpRec.EMPSLUSE)) + " to '0'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf Val(vaSpread.Text) <> 0 And EmpRec.EMPSLUSE = 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Sick Used' on row #" + CStr(vaSpread.Row) + " from '0' to " + QPTrim$(Using("###0.00", Val(vaSpread.Text))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        EmpRec.EMPSLUSE = Val(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your change has been saved successfully"
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the 'Sick Used' amount for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(Using("###0.00", EmpRec.EMPSLUSE)) + " to " + QPTrim$(Using("###0.00", Val(vaSpread.Text))) + " but declined to save it.")
      End If
    End If
  
    vaSpread.Col = 10
    If Val(vaSpread.Text) <> EmpRec.EMPCTE Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If Val(vaSpread.Text) <> 0 And EmpRec.EMPCTE <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in 'Comp Earned' field on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using$("###0.00", EmpRec.EMPCTE)) + " to " + QPTrim$(Using$("###0.00", Val(vaSpread.Text))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf Val(vaSpread.Text) = 0 And EmpRec.EMPCTE <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Comp Earned' field on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using$("###0.00", EmpRec.EMPCTE)) + " to '0'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf Val(vaSpread.Text) <> 0 And EmpRec.EMPCTE = 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Comp Earned' field on row #" + CStr(vaSpread.Row) + " from '0' to " + QPTrim$(Using$("###0.00", Val(vaSpread.Text))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        EmpRec.EMPCTE = Val(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your change has been saved successfully"
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the 'Comp Earned' field for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(Using$("###0.00", EmpRec.EMPCTE)) + " to " + QPTrim$(Using$("###0.00", vaSpread.Text)) + " but declined to save it.")
      End If
      ThisAcct = QPTrim$(vaSpread.Text)
    End If
    vaSpread.Col = 11
    If Val(vaSpread.Text) <> EmpRec.EMPCTUSE Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If Val(vaSpread.Text) <> 0 And EmpRec.EMPCTUSE <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Comp Used' on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("###0.00", EmpRec.EMPCTUSE)) + " to " + QPTrim$(Using("###0.00", Val(vaSpread.Text))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf Val(vaSpread.Text) = 0 And EmpRec.EMPCTUSE <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Comp Used' on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("###0.00", EmpRec.EMPCTUSE)) + " to '0'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf Val(vaSpread.Text) <> 0 And EmpRec.EMPCTUSE = 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Comp Used' on row #" + CStr(vaSpread.Row) + " from '0' to " + QPTrim$(Using("###0.00", Val(vaSpread.Text))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        EmpRec.EMPCTUSE = Val(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your change has been saved successfully"
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the 'Comp Used' amount for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(Using("###0.00", EmpRec.EMPCTUSE)) + " to " + QPTrim$(Using("###0.00", Val(vaSpread.Text))) + " but declined to save it.")
      End If
    End If
  
    vaSpread.Col = 13
    If Val(vaSpread.Text) <> EmpRec.PERERN Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If Val(vaSpread.Text) <> 0 And EmpRec.PERERN <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in 'Pers Earned' field on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using$("###0.00", EmpRec.PERERN)) + " to " + QPTrim$(Using$("###0.00", Val(vaSpread.Text))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf Val(vaSpread.Text) = 0 And EmpRec.PERERN <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Pers Earned' field on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using$("###0.00", EmpRec.PERERN)) + " to '0'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf Val(vaSpread.Text) <> 0 And EmpRec.PERERN = 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Pers Earned' field on row #" + CStr(vaSpread.Row) + " from '0' to " + QPTrim$(Using$("###0.00", Val(vaSpread.Text))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        EmpRec.PERERN = Val(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your change has been saved successfully"
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the 'Pers Earned' field for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(Using$("###0.00", EmpRec.PERERN)) + " to " + QPTrim$(Using$("###0.00", vaSpread.Text)) + " but declined to save it.")
      End If
      ThisAcct = QPTrim$(vaSpread.Text)
    End If
    vaSpread.Col = 14
    If Val(vaSpread.Text) <> EmpRec.PerUsed Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If Val(vaSpread.Text) <> 0 And EmpRec.PerUsed <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Pers Used' on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("###0.00", EmpRec.PerUsed)) + " to " + QPTrim$(Using("###0.00", Val(vaSpread.Text))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf Val(vaSpread.Text) = 0 And EmpRec.PerUsed <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Pers Used' on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("###0.00", EmpRec.PerUsed)) + " to '0'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf Val(vaSpread.Text) <> 0 And EmpRec.PerUsed = 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Pers Used' on row #" + CStr(vaSpread.Row) + " from '0' to " + QPTrim$(Using("###0.00", Val(vaSpread.Text))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        EmpRec.PerUsed = Val(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your change has been saved successfully"
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the 'Per Used' amount for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(Using("###0.00", EmpRec.PerUsed)) + " to " + QPTrim$(Using("###0.00", Val(vaSpread.Text))) + " but declined to save it.")
      End If
    End If
  
    vaSpread.Col = 16
    If Val(vaSpread.Text) <> EmpRec.HOLERN Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If Val(vaSpread.Text) <> 0 And EmpRec.HOLERN <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in 'Hol Earned' field on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using$("###0.00", EmpRec.HOLERN)) + " to " + QPTrim$(Using$("###0.00", Val(vaSpread.Text))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf Val(vaSpread.Text) = 0 And EmpRec.HOLERN <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Hol Earned' field on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using$("###0.00", EmpRec.HOLERN)) + " to '0'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf Val(vaSpread.Text) <> 0 And EmpRec.HOLERN = 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Hol Earned' field on row #" + CStr(vaSpread.Row) + " from '0' to " + QPTrim$(Using$("###0.00", Val(vaSpread.Text))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        EmpRec.HOLERN = Val(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your change has been saved successfully"
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the 'Hol Earned' field for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(Using$("###0.00", EmpRec.HOLERN)) + " to " + QPTrim$(Using$("###0.00", vaSpread.Text)) + " but declined to save it.")
      End If
      ThisAcct = QPTrim$(vaSpread.Text)
    End If
    vaSpread.Col = 17
    If Val(vaSpread.Text) <> EmpRec.HolUsed Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If Val(vaSpread.Text) <> 0 And EmpRec.HolUsed <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Hol Used' on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("###0.00", EmpRec.HolUsed)) + " to " + QPTrim$(Using("###0.00", Val(vaSpread.Text))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf Val(vaSpread.Text) = 0 And EmpRec.HolUsed <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Hol Used' on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("###0.00", EmpRec.HolUsed)) + " to '0'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf Val(vaSpread.Text) <> 0 And EmpRec.HolUsed = 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Hol Used' on row #" + CStr(vaSpread.Row) + " from '0' to " + QPTrim$(Using("###0.00", Val(vaSpread.Text))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        EmpRec.HolUsed = Val(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your change has been saved successfully"
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the 'Hol Used' amount for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(Using("###0.00", EmpRec.HolUsed)) + " to " + QPTrim$(Using("###0.00", Val(vaSpread.Text))) + " but declined to save it.")
      End If
    End If
  
    vaSpread.Col = 19
    If Val(vaSpread.Text) <> EmpRec.LeaveTbl Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If Val(vaSpread.Text) <> 0 And EmpRec.LeaveTbl <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Leave Table' field on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using$("###0.00", EmpRec.LeaveTbl)) + " to " + QPTrim$(Using$("###0.00", Val(vaSpread.Text))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf Val(vaSpread.Text) = 0 And EmpRec.LeaveTbl <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Leave Table' field on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using$("###0.00", EmpRec.LeaveTbl)) + " to '0'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf Val(vaSpread.Text) <> 0 And EmpRec.LeaveTbl = 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the 'Leave Table' field on row #" + CStr(vaSpread.Row) + " from '0' to " + QPTrim$(Using$("###0.00", Val(vaSpread.Text))) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        EmpRec.LeaveTbl = Val(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your change has been saved successfully"
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the 'Leave Table' field for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(Using$("###0.00", EmpRec.LeaveTbl)) + " to " + QPTrim$(Using$("###0.00", vaSpread.Text)) + " but declined to save it.")
      End If
      ThisAcct = QPTrim$(vaSpread.Text)
    End If
    vaSpread.Col = 20
    If QPTrim$(vaSpread.Text) <> QPTrim$(EmpRec.YN401K) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.YN401K) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the '401K Matching?' field on row #" + CStr(vaSpread.Row) + " from " + EmpRec.YN401K + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.YN401K) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the '401K Matching?' field on row #" + CStr(vaSpread.Row) + " from " + EmpRec.YN401K + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.YN401K) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the '401K Matching?' field on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        EmpRec.YN401K = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your change has been saved successfully"
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the '401K Match?' field for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + EmpRec.YN401K + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    vaSpread.Col = 21
    If QPTrim$(vaSpread.Text) <> QPTrim$(EmpRec.ExcludeESC) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.ExcludeESC) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in 'Exclude On ESC' field on row #" + CStr(vaSpread.Row) + " from " + EmpRec.ExcludeESC + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.ExcludeESC) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in 'Exclude On ESC' field on row #" + CStr(vaSpread.Row) + " from " + EmpRec.ExcludeESC + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.ExcludeESC) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in 'Exclude On ESC' field on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        EmpRec.ExcludeESC = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your change has been saved successfully"
      Else
        Unload frmMessageWOpts
        MainLog ("User warned that a change was made in the 'Exclude on ESC' field for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + EmpRec.ExcludeESC + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmEmpQuickMaintBenefits", "cmdEscape_Click", Erl)
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
  
  frmLoadingRpt.Label1.Caption = "Saving......"
  frmLoadingRpt.Show
  DoEvents
  
  NumOfRows = vaSpread.MaxRows
  OpenEmpData2File EHandle
  
  For x = 1 To ThisChange
    vaSpread.Row = ChangeSpot(x)
    vaSpread.Col = 22
    ThisRec = vaSpread.Value
    Get EHandle, ThisRec, EmpRec
    vaSpread.Col = 4
    EmpRec.EMPVACE = Val(vaSpread.Text)
    vaSpread.Col = 5
    EmpRec.EMPVUSED = Val(vaSpread.Text)
    vaSpread.Col = 6
    EmpRec.EMPVBAL = Val(vaSpread.Text)
    vaSpread.Col = 7
    EmpRec.EMPSLE = Val(vaSpread.Text)
    vaSpread.Col = 8
    EmpRec.EMPSLUSE = Val(vaSpread.Text)
    vaSpread.Col = 9
    EmpRec.EMPSLBAL = Val(vaSpread.Text)
    vaSpread.Col = 10
    EmpRec.EMPCTE = Val(vaSpread.Text)
    vaSpread.Col = 11
    EmpRec.EMPCTUSE = Val(vaSpread.Text)
    vaSpread.Col = 12
    EmpRec.EMPCTBAL = Val(vaSpread.Text)
    vaSpread.Col = 13
    EmpRec.PERERN = Val(vaSpread.Text)
    vaSpread.Col = 14
    EmpRec.PerUsed = Val(vaSpread.Text)
    vaSpread.Col = 15
    EmpRec.PERBAL = Val(vaSpread.Text)
    vaSpread.Col = 16
    EmpRec.HOLERN = Val(vaSpread.Text)
    vaSpread.Col = 17
    EmpRec.HolUsed = Val(vaSpread.Text)
    vaSpread.Col = 18
    EmpRec.HOLBAL = Val(vaSpread.Text)
    vaSpread.Col = 19
    EmpRec.LeaveTbl = Val(vaSpread.Text)
    vaSpread.Col = 20
    EmpRec.YN401K = QPTrim$(vaSpread.Text)
    vaSpread.Col = 21
    EmpRec.ExcludeESC = QPTrim$(vaSpread.Text)
    
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
   Unload frmLoadingRpt
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmEmpQuickMaintBenefits", "cmdSave_Click", Erl)
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
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub LoadMe()
  Dim IdxRec As NameSortIdxType
  Dim XHandle As Integer
  Dim x As Integer
  
  OpenEmpIdxLNameFile XHandle
  EmployeeCount = LOF(XHandle) / 2
  Close
  
  DontExit = False
  ThisChange = 0
  ThisParameter = "All Employees"
  ThisTerm = chkTerm.Value
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
      MainLog ("Payroll.exe terminated via menu bar on frmEmpQuickMaintWageDist.")
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
  Dim ThisCol As Integer
  Dim LHandle As Integer
  Dim LeaveRec As LeaveRecType
  Dim NumOfRecs As Integer
  
  On Error GoTo ERRORSTUFF
  
  OpenLeaveFileName LHandle
  NumOfRecs = LOF(LHandle) \ Len(LeaveRec)
  Close LHandle
  
  ReDim LeaveTbl(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    LeaveTbl(x) = x
  Next x
    
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
    vaSpread.Row = RowMax
    vaSpread.Col = 1
    vaSpread.Text = QPTrim$(EmpRec.EmpNo)
    vaSpread.Col = 2
    vaSpread.Text = QPTrim$(EmpRec.EmpLName)
    vaSpread.Col = 3
    vaSpread.Text = QPTrim$(EmpRec.EmpFName)
    vaSpread.Col = 4
    vaSpread.Text = EmpRec.EMPVACE
    vaSpread.Col = 5
    vaSpread.Text = EmpRec.EMPVUSED
    vaSpread.Col = 6
    vaSpread.Text = EmpRec.EMPVBAL
    vaSpread.Col = 7
    vaSpread.Text = EmpRec.EMPSLE
    vaSpread.Col = 8
    vaSpread.Text = EmpRec.EMPSLUSE
    vaSpread.Col = 9
    vaSpread.Text = EmpRec.EMPSLBAL
    vaSpread.Col = 10
    vaSpread.Text = EmpRec.EMPCTE
    vaSpread.Col = 11
    vaSpread.Text = EmpRec.EMPCTUSE
    vaSpread.Col = 12
    vaSpread.Text = EmpRec.EMPCTBAL
    vaSpread.Col = 13
    vaSpread.Text = EmpRec.PERERN
    vaSpread.Col = 14
    vaSpread.Text = EmpRec.PerUsed
    vaSpread.Col = 15
    vaSpread.Text = EmpRec.PERBAL
    vaSpread.Col = 16
    vaSpread.Text = EmpRec.HOLERN
    vaSpread.Col = 17
    vaSpread.Text = EmpRec.HolUsed
    vaSpread.Col = 18
    vaSpread.Text = EmpRec.HOLBAL
    vaSpread.Col = 19
    vaSpread.TypeComboBoxIndex = -1
    vaSpread.TypeComboBoxString = 0
    For y = 1 To NumOfRecs
      vaSpread.TypeComboBoxIndex = -1
      vaSpread.TypeComboBoxString = LeaveTbl(y)
    Next y
    vaSpread.Text = EmpRec.LeaveTbl
    vaSpread.Col = 20
    vaSpread.Text = EmpRec.YN401K
    vaSpread.Col = 21
    vaSpread.Text = EmpRec.ExcludeESC
    vaSpread.Col = 22
    vaSpread.Value = ThisIdx(x)
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmEmpQuickMaintBenefits", "LoadSpread", Erl)
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
  ThisParameter = QPTrim$(fpcmbParameters.Text)
  
  Call LoadSpread(ThisLoadSpread)
  
BooBooFound:
End Sub


Private Sub vaSpread_Change(ByVal Col As Long, ByVal Row As Long)
  Dim ThisEarned As Double
  Dim ThisUsed As Double
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
  
  vaSpread.Row = Row
  Select Case Col
    Case 4, 5
      vaSpread.Col = 4
        ThisEarned = Val(vaSpread.Text)
      vaSpread.Col = 5
        ThisUsed = Val(vaSpread.Text)
      vaSpread.Col = 6
        vaSpread.Text = CStr(ThisEarned - ThisUsed)
    Case 7, 8
      vaSpread.Col = 7
        ThisEarned = Val(vaSpread.Text)
      vaSpread.Col = 8
        ThisUsed = Val(vaSpread.Text)
      vaSpread.Col = 9
        vaSpread.Text = CStr(ThisEarned - ThisUsed)
    Case 10, 11
      vaSpread.Col = 10
        ThisEarned = Val(vaSpread.Text)
      vaSpread.Col = 11
        ThisUsed = Val(vaSpread.Text)
      vaSpread.Col = 12
        vaSpread.Text = CStr(ThisEarned - ThisUsed)
    Case 13, 14
      vaSpread.Col = 13
        ThisEarned = Val(vaSpread.Text)
      vaSpread.Col = 14
        ThisUsed = Val(vaSpread.Text)
      vaSpread.Col = 15
        vaSpread.Text = CStr(ThisEarned - ThisUsed)
    Case 16, 17
      vaSpread.Col = 16
        ThisEarned = Val(vaSpread.Text)
      vaSpread.Col = 17
        ThisUsed = Val(vaSpread.Text)
      vaSpread.Col = 18
        vaSpread.Text = CStr(ThisEarned - ThisUsed)
    Case Else
  End Select
End Sub

Private Sub ClearChanges()
  ThisChange = 0
  ReDim ChangeSpot(0 To 0) As Integer
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
  vaSpread.Col = 21
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
  vaSpread.Col = 21
  vaSpread.BackColor = &H80000005

End Sub

