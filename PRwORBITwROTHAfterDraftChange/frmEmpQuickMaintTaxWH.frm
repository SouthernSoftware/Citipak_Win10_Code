VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmEmpQuickMaintTaxWH 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll: Quick Employee Maintenance"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmEmpQuickMaintTaxWH.frx":0000
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
      TabIndex        =   9
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
      ColDesigner     =   "frmEmpQuickMaintTaxWH.frx":08CA
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
      Left            =   4219
      TabIndex        =   1
      Top             =   1755
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Left            =   386
      Top             =   555
   End
   Begin FPSpread.vaSpread vaSpread 
      Height          =   3975
      Left            =   840
      TabIndex        =   0
      Top             =   3195
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
      MaxCols         =   19
      MaxRows         =   1000000
      OperationMode   =   2
      ProcessTab      =   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   12648447
      SpreadDesigner  =   "frmEmpQuickMaintTaxWH.frx":0B89
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   690
      Left            =   7819
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Press F10 to save all the changes made on this spreadsheet."
      Top             =   7635
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
      ButtonDesigner  =   "frmEmpQuickMaintTaxWH.frx":1C47D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
      Height          =   690
      Left            =   5021
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Press ESC to exit this screen after each cell is examined for unsaved changes."
      Top             =   7635
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
      ButtonDesigner  =   "frmEmpQuickMaintTaxWH.frx":1C659
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExitNow 
      Height          =   690
      Left            =   2224
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Press to exit this screen without testing each cell for unsaved changes."
      Top             =   7635
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
      ButtonDesigner  =   "frmEmpQuickMaintTaxWH.frx":1C835
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
      Left            =   3469
      TabIndex        =   8
      Top             =   555
      Width           =   4995
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   3439
      Top             =   405
      Width           =   5055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Withholdings"
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
      Left            =   4309
      TabIndex        =   7
      Top             =   915
      Width           =   3315
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
      Left            =   3315
      TabIndex        =   6
      Top             =   2355
      Width           =   2670
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   675
      Top             =   3075
      Width           =   10575
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
      Left            =   750
      TabIndex        =   5
      Top             =   2715
      Width           =   2175
   End
End
Attribute VB_Name = "frmEmpQuickMaintTaxWH"
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
  Dim BooBoo As Boolean
  Dim DontExit As Boolean
  Dim BigEmpNum As String * 10

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
  Dim ThisAmtPct$
  
  If ThisChange = 0 Then GoTo NoChanges
  On Error GoTo ERRORSTUFF
  
  NumOfRows = vaSpread.MaxRows
  OpenEmpData2File EHandle
  
  For x = 1 To ThisChange 'NumOfRows
    vaSpread.Row = ChangeSpot(x)
    vaSpread.Col = 19
    ThisRec = CInt(vaSpread.Text)
    Get EHandle, ThisRec, EmpRec
    vaSpread.Col = 4
    If QPTrim$(UCase(vaSpread.Text)) <> QPTrim$(UCase(EmpRec.EMPFEDX)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPFEDX) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the federal exemption status on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPFEDX) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EMPFEDX) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the federal exemption status on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPFEDX) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPFEDX) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the federal exemption status on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
          Exit Sub
        End If
        EmpRec.EMPFEDX = QPTrim$(UCase(vaSpread.Text))
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageW3Opts
        MainLog ("User warned that a change was made in the federal exemption status for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EMPFEDX) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 5
    If QPTrim$(Mid(vaSpread.Text, 1, 1)) <> QPTrim$(UCase(EmpRec.EMPFEDO2)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPFEDO2) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the federal amount/percent on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPFEDO2) + " to " + Mid(vaSpread.Text, 1, 1) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EMPFEDO2) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the federal amount/percent on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPFEDO2) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPFEDO2) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the federal amount/percent on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + Mid(vaSpread.Text, 1, 1) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        If CheckThisAmtPctAndFigure(vaSpread.Row) = False Then
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
        EmpRec.EMPFEDO2 = Mid(vaSpread.Text, 1, 1)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageW3Opts
        MainLog ("User warned that a change was made in the federal amount/percent for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EMPFEDO2) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    ThisAmtPct = Mid(vaSpread.Text, 1, 1)
    
    vaSpread.Col = 6
    ThisAmt = QPTrim$(vaSpread.Text)
    ThisAmt = ReplaceString(ThisAmt, ",", "")
    If Val(ThisAmt) <> EmpRec.EMPFEDO1 Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If Val(ThisAmt) <> 0 And EmpRec.EMPFEDO1 <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the federal figure amount on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("$#,##0.00", EmpRec.EMPFEDO1)) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf Val(ThisAmt) = 0 And EmpRec.EMPFEDO1 <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the federal figure amount on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("$#,##0.00", EmpRec.EMPFEDO1)) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf Val(ThisAmt) <> 0 And EmpRec.EMPFEDO1 = 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the federal figure amount on row #" + CStr(vaSpread.Row) + " from '0.00' to " + QPTrim(vaSpread.Text) + " . To review this change press F5. To save this change press F10. To abandon this change press ESC."
      End If
      '9/9/04
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
        If CheckThisAmtPctAndFigure(vaSpread.Row) = False Then
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
        EmpRec.EMPFEDO1 = Val(ThisAmt)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageW3Opts
        MainLog ("User warned that a change was made in the federal figure amount for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(Using("$#,##0.00", EmpRec.EMPFEDO1)) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 7
    If Mid(vaSpread.Text, 1, 1) <> QPTrim$(UCase(EmpRec.EMPFEDS)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the federal status on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPFEDS) + " to " + Mid(vaSpread.Text, 1, 1) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the federal status on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPFEDS) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf EmpRec.EMPFEDS = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the federal status on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + Mid(vaSpread.Text, 1, 1) + " . To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
          Exit Sub
        End If
        EmpRec.EMPFEDS = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageW3Opts
        MainLog ("User warned that a change was made in the federal status for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EMPFEDS) + " to " + Mid(vaSpread.Text, 1, 1) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 8
    If Val(QPTrim(vaSpread.Text)) <> EmpRec.EMPFEDA Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And EmpRec.EMPFEDA <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the federal allowances on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("#0", EmpRec.EMPFEDA)) + " to " + QPTrim(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And EmpRec.EMPFEDA <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the federal allowances on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("#0", EmpRec.EMPFEDA)) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And EmpRec.EMPFEDA = 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the federal allowances on row #" + CStr(vaSpread.Row) + " from '0' to " + QPTrim(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
          BooBoo = True
          Exit Sub
        End If
        EmpRec.EMPFEDA = Val(QPTrim$(vaSpread.Text))
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageW3Opts
        MainLog ("User warned that a change was made in the federal allowances for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(Using("#0", EmpRec.EMPFEDA)) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 9
    ThisAmt = ReplaceString(vaSpread.Text, "$", "")
    ThisAmt = ReplaceString(ThisAmt, ",", "")
    If Val(ThisAmt) <> EmpRec.EMPFEDAA Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(ThisAmt) <> "" And EmpRec.EMPFEDAA <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the federal withholdng allowances on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("#0", EmpRec.EMPFEDAA)) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(ThisAmt) = "" And EmpRec.EMPFEDAA <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the federal withholdng allowances on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("#0", EmpRec.EMPFEDAA)) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(ThisAmt) <> "" And EmpRec.EMPFEDAA = 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the federal withholdng allowances on row #" + CStr(vaSpread.Row) + " from '0' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        EmpRec.EMPFEDAA = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageW3Opts
        MainLog ("User warned that a change was made in the federal withholdng allowances for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(Using("$#0", EmpRec.EMPFEDAA)) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 10
    If QPTrim$(vaSpread.Text) <> QPTrim$(UCase(EmpRec.EMPSTAX)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPSTAX) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the state exemption status on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPSTAX) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EMPSTAX) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the state exemption status on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPSTAX) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPSTAX) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the state exemption status on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
          Exit Sub
        End If
        EmpRec.EMPSTAX = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageW3Opts
        MainLog ("User warned that a change was made in the state exemption status for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EMPSTAX) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 11
    If QPTrim$(Mid(vaSpread.Text, 1, 1)) <> QPTrim$(UCase(EmpRec.EMPSTAO2)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPSTAO2) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the state amount/percent on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPSTAO2) + " to " + Mid(vaSpread.Text, 1, 1) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EMPSTAO2) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the state amount/percent on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPSTAO2) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPSTAO2) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the state amount/percent on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + Mid(vaSpread.Text, 1, 1) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        If CheckThisAmtPctAndFigure(vaSpread.Row) = False Then
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
        EmpRec.EMPSTAO2 = Mid(vaSpread.Text, 1, 1)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageW3Opts
        MainLog ("User warned that a change was made in the state amount/percent for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EMPSTAO2) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    ThisAmtPct = Mid(vaSpread.Text, 1, 1)
    
    vaSpread.Col = 12
    ThisAmt = QPTrim$(vaSpread.Text)
    ThisAmt = ReplaceString(ThisAmt, ",", "")
    If Val(ThisAmt) <> EmpRec.EMPSTAO1 Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If Val(ThisAmt) <> 0 And EmpRec.EMPSTAO1 <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the state figure amount on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("$#,##0.00", EmpRec.EMPSTAO1)) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf Val(ThisAmt) = 0 And EmpRec.EMPSTAO1 <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the state figure amount on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("$#,##0.00", EmpRec.EMPSTAO1)) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf Val(ThisAmt) <> 0 And EmpRec.EMPSTAO1 = 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the state figure amount on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("$#,##0.00", EmpRec.EMPSTAO1)) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        If CheckThisAmtPctAndFigure(vaSpread.Row) = False Then
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
        EmpRec.EMPSTAO1 = Val(ThisAmt)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageW3Opts
        MainLog ("User warned that a change was made in the state figure amount for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(Using("$#,##0.00", EmpRec.EMPSTAO1)) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 13
    If Mid(vaSpread.Text, 1, 1) <> QPTrim$(UCase(EmpRec.EMPSTAS)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And EmpRec.EMPSTAS <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the state status on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPSTAS) + " to " + Mid(vaSpread.Text, 1, 1) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And EmpRec.EMPSTAS <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the state status on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPSTAS) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And EmpRec.EMPSTAS = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the state status on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
          Exit Sub
        End If
        EmpRec.EMPSTAS = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageW3Opts
        MainLog ("User warned that a change was made in the state status for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EMPSTAS) + " to " + Mid(vaSpread.Text, 1, 1) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 14
    If Val(QPTrim$(vaSpread.Text)) <> EmpRec.EMPSTAA Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And EmpRec.EMPSTAA <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the state allowances on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("#0", EmpRec.EMPSTAA)) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And EmpRec.EMPSTAA <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the state allowances on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("#0", EmpRec.EMPSTAA)) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And EmpRec.EMPSTAA = 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the state allowances on row #" + CStr(vaSpread.Row) + " from '0' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
          BooBoo = True
          Exit Sub
        End If
        EmpRec.EMPSTAA = Val(QPTrim$(vaSpread.Text))
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageW3Opts
        MainLog ("User warned that a change was made in the state allowances for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(Using("#0", EmpRec.EMPSTAA)) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 15
    ThisAmt = ReplaceString(vaSpread.Text, "$", "")
    ThisAmt = ReplaceString(ThisAmt, ",", "")
    If Val(ThisAmt) <> EmpRec.EMPSTAAA Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And EmpRec.EMPSTAAA <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the additional state withholding amount on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("$#,##0.00", EmpRec.EMPSTAAA)) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And EmpRec.EMPSTAAA <> 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the additional state withholding amount on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(Using("$#,##0.00", EmpRec.EMPSTAAA)) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And EmpRec.EMPSTAAA = 0 Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the additional state withholding amount on row #" + CStr(vaSpread.Row) + " from '0' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        EmpRec.EMPSTAAA = Val(ThisAmt)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageW3Opts
        MainLog ("User warned that a change was made in the additional state withholding amount for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(Using("$#,##0.00", EmpRec.EMPSTAAA)) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 16
    If QPTrim$(vaSpread.Text) <> QPTrim$(UCase(EmpRec.EMPSOCX)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPSOCX) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the social security exemption on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPSOCX) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EMPSOCX) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the social security exemption on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPSOCX) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPSOCX) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the social security exemption on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        EmpRec.EMPSOCX = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageW3Opts
        MainLog ("User warned that a change was made in the social security exemption for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EMPSOCX) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 17
    If QPTrim$(vaSpread.Text) <> QPTrim$(UCase(EmpRec.EMPMEDX)) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPMEDX) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the medicare exemption on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPMEDX) + " to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EMPMEDX) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the medicare exemption on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPMEDX) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPMEDX) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the medicare exemption on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + QPTrim$(vaSpread.Text) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        EmpRec.EMPMEDX = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageW3Opts
        MainLog ("User warned that a change was made in the social security exemption for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EMPSOCX) + " to " + QPTrim$(vaSpread.Text) + " but declined to save it.")
      End If
    End If
    
    vaSpread.Col = 18
    If Mid(vaSpread.Text, 1, 1) <> QPTrim$(EmpRec.EMPEIC) Then
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      If QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPEIC) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the EIC code on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPEIC) + " to " + Mid(vaSpread.Text, 1, 1) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) = "" And QPTrim$(EmpRec.EMPEIC) <> "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the EIC code on row #" + CStr(vaSpread.Row) + " from " + QPTrim$(EmpRec.EMPEIC) + " to 'BLANK'. To review this change press F5. To save this change press F10. To abandon this change press ESC."
      ElseIf QPTrim$(vaSpread.Text) <> "" And QPTrim$(EmpRec.EMPEIC) = "" Then
        frmMessageW3Opts.Label1.Caption = "A change has been made in the EIC code on row #" + CStr(vaSpread.Row) + " from 'BLANK' to " + Mid(vaSpread.Text, 1, 1) + ". To review this change press F5. To save this change press F10. To abandon this change press ESC."
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
        EmpRec.EMPEIC = QPTrim$(vaSpread.Text)
        Put EHandle, ThisRec, EmpRec
        MsgBox "Your data has been saved successfully."
      Else
        Unload frmMessageW3Opts
        MainLog ("User warned that a change was made in the EIC code for " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + " from " + QPTrim$(EmpRec.EMPEIC) + " to " + Mid(vaSpread.Text, 1, 1) + " but declined to save it.")
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmEmpQuickMaintTaxWH", "cmdEscape_Click", Erl)
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
  Dim ThisOne$
  
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
  
  vaSpread.Col = 1
  For x = 1 To vaSpread.MaxRows
    vaSpread.Row = x
    If vaSpread.BackColor = &H8080FF Then vaSpread.BackColor = &H80000005
  Next x
  
  If CheckAmtPctAndFigure = False Then
    Unload frmLoadingRpt
    Close
    Exit Sub
  End If
  
  If RequiredFieldsOK = False Then
    Unload frmLoadingRpt
    Close
    Exit Sub
  End If
  
  Unload frmLoadingRpt
  
  NumOfRows = vaSpread.MaxRows
  OpenEmpData2File EHandle
  
  FrmShowPctComp.Show , Me
  FrmShowPctComp.cmdCancel.Visible = False
  FrmShowPctComp.Label1.Caption = "Saving..."
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdExitNow.Enabled = False
  Me.cmdSave.Enabled = False
  
  For x = 1 To ThisChange
    ThisOne = ""
    vaSpread.Row = ChangeSpot(x)
    vaSpread.Col = 19
    ThisRec = vaSpread.Value
    Get EHandle, ThisRec, EmpRec
    vaSpread.Col = 4
    EmpRec.EMPFEDX = UCase(Mid(vaSpread.Text, 1, 1))
    vaSpread.Col = 5
    EmpRec.EMPFEDO2 = Mid(vaSpread.Text, 1, 1)
    vaSpread.Col = 6
    ThisOne = ReplaceString(vaSpread.Text, "$", "")
    ThisOne = ReplaceString(ThisOne, ",", "")
    EmpRec.EMPFEDO1 = Val(ThisOne)
    vaSpread.Col = 7
    EmpRec.EMPFEDS = UCase(Mid(vaSpread.Text, 1, 1))
    vaSpread.Col = 8
    EmpRec.EMPFEDA = Val(QPTrim$(vaSpread.Text))
    vaSpread.Col = 9
    ThisOne = ReplaceString(vaSpread.Text, "$", "")
    ThisOne = ReplaceString(ThisOne, ",", "")
    EmpRec.EMPFEDAA = Val(ThisOne)
    vaSpread.Col = 10
    EmpRec.EMPSTAX = UCase(Mid(vaSpread.Text, 1, 1))
    vaSpread.Col = 11
    EmpRec.EMPSTAO2 = Mid(vaSpread.Text, 1, 1)
    vaSpread.Col = 12
    ThisOne = ReplaceString(vaSpread.Text, "$", "")
    ThisOne = ReplaceString(ThisOne, ",", "")
    EmpRec.EMPSTAO1 = Val(ThisOne)
    vaSpread.Col = 13
    EmpRec.EMPSTAS = UCase(Mid(vaSpread.Text, 1, 1))
    vaSpread.Col = 14
    EmpRec.EMPSTAA = Val(QPTrim$(vaSpread.Text))
    vaSpread.Col = 15
    ThisOne = ReplaceString(vaSpread.Text, "$", "")
    ThisOne = ReplaceString(ThisOne, ",", "")
    EmpRec.EMPSTAAA = Val(ThisOne)
    vaSpread.Col = 16
    EmpRec.EMPSOCX = UCase(Mid(vaSpread.Text, 1, 1))
    vaSpread.Col = 17
    EmpRec.EMPMEDX = UCase(Mid(vaSpread.Text, 1, 1))
    vaSpread.Col = 18
    EmpRec.EMPEIC = Mid(vaSpread.Text, 1, 1)
    
    Put EHandle, ThisRec, EmpRec
    
    FrmShowPctComp.ShowPctComp x, ThisChange
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
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmEmpQuickMaintTaxWH", "cmdSave_Click", Erl)
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
  Dim UnitRec As UnitFileRecType
  Dim UHandle As Integer
  Dim ThisState$
  
  OpenUnitFile UHandle
  Get UHandle, 1, UnitRec
  Close UHandle
  ThisState = UnitRec.UFSTATE
  
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
  
  vaSpread.Col = 13
  Select Case ThisState
    Case "GA":
      vaSpread.TypeComboBoxList = "H-GASingle" + Chr(9) + "G-GAMarried" + Chr(9) + "F-GAHead of HouseHold"
    Case "SC":
      vaSpread.TypeComboBoxList = "S-SC Tax Table"
    Case "OK":
      vaSpread.TypeComboBoxList = "S-Single" + Chr(9) + "M-Married, Head of Household" + Chr(9) + "D-Dual Income Married"
    Case "AR":
      vaSpread.TypeComboBoxList = "S-Single" + Chr(9) + "M-Married (1 Exempt'n)" + Chr(9) + "H-Married/Head Fam(2 Exempt'ns"
   Case Else:
      vaSpread.TypeComboBoxList = "S-Single" + Chr(9) + "M-Married" + Chr(9) + "H-Head of HouseHold"
  End Select
  
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
      MainLog ("Payroll.exe terminated via menu bar on frmEmpQuickMaintTaxWH.")
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
  Dim UnitRec As UnitFileRecType
  Dim UHandle As Integer
  Dim State$
  
  On Error GoTo ERRORSTUFF
  
  Call ClearChanges
  ReDim ChangeSpot(1 To vaSpread.MaxRows) As Integer
  
  OpenUnitFile UHandle
  Get UHandle, 1, UnitRec
  Close UHandle
  State = QPTrim$(UnitRec.UFSTATE)
  vaSpread.Col = 13
  Select Case State
    Case "GA"
      vaSpread.TypeComboBoxList = "GASingle" + Chr$(9) + "GAMarried" + Chr$(9) + "GAHead of Household"
    Case "OK"
      vaSpread.TypeComboBoxList = "Single" + Chr$(9) + "Married, Head Of Household" + Chr$(9) + "Dual Income Married"
    Case "SC"
      vaSpread.TypeComboBoxList = "Single"
    Case Else
      vaSpread.TypeComboBoxList = "S-Single" + Chr$(9) + "M-Married" + Chr$(9) + "H-Head of Household"
  End Select
  
  vaSpread.MaxRows = EmployeeCount
  
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
    vaSpread.Text = EmpRec.EMPFEDX
    vaSpread.Col = 5
    vaSpread.Row = RowMax
    If EmpRec.EMPFEDO2 = "A" Then
      vaSpread.Text = "Amount"
    ElseIf EmpRec.EMPFEDO2 = "P" Then
      vaSpread.Text = "Percent"
    Else
      vaSpread.Text = ""
    End If
    vaSpread.Col = 6
    vaSpread.Row = RowMax
    vaSpread.Text = EmpRec.EMPFEDO1
    vaSpread.Col = 7
    vaSpread.Row = RowMax
    If EmpRec.EMPFEDS = "S" Then
      vaSpread.Text = "S-Single"
    ElseIf EmpRec.EMPFEDS = "M" Then
      vaSpread.Text = "M-Married"
    End If
    vaSpread.Col = 8
    vaSpread.Row = RowMax
    vaSpread.Text = EmpRec.EMPFEDA
    vaSpread.Col = 9
    vaSpread.Row = RowMax
    vaSpread.Text = EmpRec.EMPFEDAA
    vaSpread.Col = 10
    vaSpread.Row = RowMax
    vaSpread.Text = EmpRec.EMPSTAX
    vaSpread.Col = 11
    vaSpread.Row = RowMax
    If EmpRec.EMPSTAO2 = "A" Then
      vaSpread.Text = "Amount"
    ElseIf EmpRec.EMPSTAO2 = "P" Then
      vaSpread.Text = "Percent"
    Else
      vaSpread.Text = ""
    End If
    vaSpread.Col = 12
    vaSpread.Row = RowMax
    vaSpread.Text = EmpRec.EMPSTAO1
    vaSpread.Col = 13
    vaSpread.Row = RowMax
    Select Case State
      Case "SC"
        If EmpRec.EMPSTAS = "S" Then
          vaSpread.Text = "S-Single"
        Else
          vaSpread.Text = ""
        End If
      Case "GA"
        If EmpRec.EMPSTAS = "H" Then
          vaSpread.Text = "GASingle"
        ElseIf EmpRec.EMPSTAS = "G" Then
          vaSpread.Text = "GAMarried"
        ElseIf EmpRec.EMPSTAS = "F" Then
          vaSpread.Text = "GAHead Of Household"
        Else
          vaSpread.Text = ""
        End If
      Case "OK"
        If EmpRec.EMPSTAS = "S" Then
          vaSpread.Text = "Single"
        ElseIf EmpRec.EMPSTAS = "M" Then
          vaSpread.Text = "Married, Head of Household"
        ElseIf EmpRec.EMPSTAS = "D" Then
          vaSpread.Text = "Dual Income Married"
        Else
          vaSpread.Text = ""
        End If
      Case Else
        If EmpRec.EMPSTAS = "S" Then
          vaSpread.Text = "S-Single"
        ElseIf EmpRec.EMPSTAS = "M" Then
          vaSpread.Text = "M-Married"
        ElseIf EmpRec.EMPSTAS = "H" Then
          vaSpread.Text = "H-Head of Household"
        Else
          vaSpread.Text = ""
        End If
    End Select
    vaSpread.Col = 14
    vaSpread.Row = RowMax
    vaSpread.Text = EmpRec.EMPSTAA
    vaSpread.Col = 15
    vaSpread.Row = RowMax
    vaSpread.Text = EmpRec.EMPSTAAA
    vaSpread.Col = 16
    vaSpread.Row = RowMax
    vaSpread.Text = EmpRec.EMPSOCX
    vaSpread.Col = 17
    vaSpread.Row = RowMax
    vaSpread.Text = EmpRec.EMPMEDX
    vaSpread.Col = 18
    vaSpread.Row = RowMax
    If EmpRec.EMPEIC = "0" Then
      vaSpread.Text = "0-No Certificate"
    ElseIf EmpRec.EMPEIC = "1" Then
      vaSpread.Text = "1-Employee Only"
    ElseIf EmpRec.EMPEIC = "2" Then
      vaSpread.Text = "2-Employee & Spouse"
    End If
    vaSpread.Col = 19
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmEmpQuickMaintTaxWH", "LoadSpread", Erl)
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
  vaSpread.Col = 15
  vaSpread.BackColor = &HC0FFFF
  vaSpread.Col = 16
  vaSpread.BackColor = &HC0FFFF
  vaSpread.Col = 17
  vaSpread.BackColor = &HC0FFFF
  vaSpread.Col = 18
  vaSpread.BackColor = &HC0FFFF
End Sub

Private Sub vaSpread_KeyPress(KeyAscii As Integer)
  vaSpread.OperationMode = OperationModeRow
End Sub

Private Sub vaSpread_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
  vaSpread.OperationMode = OperationModeRow
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

End Sub

Private Function TestForRetNumT() As Boolean
  Dim x As Integer
  Dim ThisTCnt As Integer
  Dim FirstOne As Integer
  
  TestForRetNumT = True
  vaSpread.Col = 13
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
      TCnt(ThisTCnt) = ChangeSpot(x)
    End If
  Next x
  
  If ThisTCnt > 0 Then
    vaSpread.Row = FirstOne
    vaSpread.SetFocus
    vaSpread.OperationMode = OperationModeNormal
    vaSpread.SetActiveCell 13, FirstOne
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
    vaSpread.Col = 19
    vaSpread.Row = ChangeSpot(x)
    ThisRec = vaSpread.Value
    Get EmpHandle, ThisRec, Emp2Rec
    vaSpread.Col = 4
    If QPTrim$(vaSpread.Text) = "" Then
      frmMessage.Label1.Caption = "Federal exempt value is a required field. Please supply this data on row # " + CStr(vaSpread.Row) + "."
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
      frmMessage.Label1.Caption = "Federal status name is a required field. Please supply this data on row # " + CStr(vaSpread.Row) + "."
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
      frmMessage.Label1.Caption = "Federal allowances value is a required field. Please supply this data on row # " + CStr(vaSpread.Row) + "."
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
      frmMessage.Label1.Caption = "State exempt is a required field. Please supply this data on row # " + CStr(vaSpread.Row) + "."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      Close
      RequiredFieldsOK = False
      Exit Function
    End If
    vaSpread.Col = 13
    If QPTrim$(vaSpread.Text) = "" Then
      frmMessage.Label1.Caption = "State status is a required field. Please supply this data on row # " + CStr(vaSpread.Row) + "."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      Close
      RequiredFieldsOK = False
      Exit Function
    End If
    vaSpread.Col = 14
    If QPTrim$(vaSpread.Text) = "" Then
      frmMessage.Label1.Caption = "State allowances is a required field. Please supply this data on row # " + CStr(vaSpread.Row) + "."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      Close
      RequiredFieldsOK = False
      Exit Function
    End If
    vaSpread.Col = 16
    If QPTrim$(vaSpread.Text) = "" Then
      frmMessage.Label1.Caption = "Social security exempt is a required field. Please supply this data on row # " + CStr(vaSpread.Row) + "."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      Close
      RequiredFieldsOK = False
      Exit Function
    End If
    vaSpread.Col = 17
    If QPTrim$(vaSpread.Text) = "" Then
      frmMessage.Label1.Caption = "Medicare exempt is a required field. Please supply this data on row # " + CStr(vaSpread.Row) + "."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell vaSpread.Col, vaSpread.Row
      Close
      RequiredFieldsOK = False
      Exit Function
    End If
    vaSpread.Col = 18
    If QPTrim$(vaSpread.Text) = "" Then
      frmMessage.Label1.Caption = "EIC code is a required field. Please supply this data on row # " + CStr(vaSpread.Row) + "."
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

End Sub

Private Function CheckAmtPctAndFigure() As Boolean
  Dim x As Integer
  Dim ThisAmtPct$
  Dim ThisAmt$
  
  CheckAmtPctAndFigure = True
  For x = 1 To ThisChange
    vaSpread.Row = ChangeSpot(x)
    vaSpread.Col = 5
    ThisAmtPct = QPTrim$(vaSpread.Text)
    vaSpread.Col = 6
    ThisAmt = ReplaceString(vaSpread.Text, ",", "")
    If Val(ThisAmt) > 0 And ThisAmtPct = "" Then
      frmMessage.Label1.Caption = "A value has been entered for 'Fixed Federal Figure' on row " + CStr(vaSpread.Row) + " but there is no federal amount or percent saved. Please enter either amount or percent in column #5 before continuing."
      frmMessage.Label1.Top = 700
      frmMessage.Show vbModal
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell 5, vaSpread.Row
      CheckAmtPctAndFigure = False
      Exit Function
    End If
    If Val(ThisAmt) > 100 And Mid(ThisAmtPct, 1, 1) = "P" Then
      frmMessage.Label1.Caption = "Percent has been entered for 'Federal Amt/Pct' on row " + CStr(vaSpread.Row) + " but the amount for 'Fixed Federal Figure' is more than 100%. This percent figure is invalid."
      frmMessage.Label1.Top = 700
      frmMessage.Show vbModal
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell 6, vaSpread.Row
      CheckAmtPctAndFigure = False
      Exit Function
    End If
    vaSpread.Row = ChangeSpot(x)
    vaSpread.Col = 11
    ThisAmtPct = QPTrim$(vaSpread.Text)
    vaSpread.Col = 12
    ThisAmt = ReplaceString(vaSpread.Text, ",", "")
    If Val(ThisAmt) > 0 And ThisAmtPct = "" Then
      frmMessage.Label1.Caption = "A value has been entered for 'Fixed State Figure' on row " + CStr(vaSpread.Row) + " but there is no state amount or percent saved. Please enter either amount or percent in column #11 before continuing."
      frmMessage.Label1.Top = 700
      frmMessage.Show vbModal
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell 11, vaSpread.Row
      CheckAmtPctAndFigure = False
      Exit Function
    End If
    If Val(ThisAmt) > 100 And Mid(ThisAmtPct, 1, 1) = "P" Then
      frmMessage.Label1.Caption = "Percent has been entered for 'State Amt/Pct' on row " + CStr(vaSpread.Row) + " but the amount for 'Fixed State Figure' is more than 100%. This percent figure is invalid."
      frmMessage.Label1.Top = 700
      frmMessage.Show vbModal
      vaSpread.SetFocus
      vaSpread.OperationMode = OperationModeNormal
      vaSpread.SetActiveCell 12, vaSpread.Row
      CheckAmtPctAndFigure = False
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

'Private Function CheckValAmtPct(WhichOnes As Integer) As Boolean
'  Dim x As Integer
'  Dim ThisAmtPct$
'  Dim ThisAmt$
'  Dim ThisMany As Integer
'
'  CheckValAmtPct = True
'  If WhichOnes = -1 Then
'    ThisMany = vaSpread.MaxRows
'  Else
'    ThisMany = 1
'  End If
'
'  For x = 1 To ThisMany
'    If ThisMany > 1 Then
'      vaSpread.Row = x
'    Else
'      vaSpread.Row = WhichOnes
'    End If
'    vaSpread.Col = 5
'    ThisAmtPct = Mid(vaSpread.Text, 1, 1)
'    vaSpread.Col = 6
'    ThisAmt = ReplaceString(vaSpread.Text, ",", "")
'    If Val(ThisAmt) > 0 And QPTrim$(ThisAmtPct) = "" Then
'      vaSpread.SetFocus
'      vaSpread.OperationMode = OperationModeNormal
'      vaSpread.SetActiveCell 5, vaSpread.Row
'      frmMessage.Label1.Caption = "On row " + CStr(vaSpread.Row) + " you are attempting to save a value in the 'Withholding' cell but there is nothing saved for 'Deduction Amt/Pct'. Please enter either 'Amount' or 'Percent' if you wish to save a value in the 'Withholding' cell."
'      frmMessage.Label1.Top = 650
'      frmMessage.Show vbModal
'      CheckValAmtPct = False
'      Exit Function
'    End If
'    If QPTrim$(ThisAmtPct) = "P" And Val(ThisAmt) > 100 Then
'      vaSpread.SetFocus
'      vaSpread.OperationMode = OperationModeNormal
'      vaSpread.SetActiveCell 6, vaSpread.Row
'      frmMessage.Label1.Caption = "On row " + CStr(vaSpread.Row) + " you are attempting to save a 'Percent' for 'Deductions Amt/Pct' but the withholding amount is greater than 100. Percentages are limited to 100%. Please reduce the 'Withholding' value."
'      frmMessage.Label1.Top = 700
'      frmMessage.Show vbModal
'      CheckValAmtPct = False
'      Exit Function
'    End If
'  Next x
'
'End Function

Private Function CheckThisAmtPctAndFigure(ThisRow As Integer) As Boolean
  Dim x As Integer
  Dim ThisAmtPct$
  Dim ThisAmt$
  
  CheckThisAmtPctAndFigure = True
  vaSpread.Row = ThisRow
  vaSpread.Col = 5
  ThisAmtPct = QPTrim$(vaSpread.Text)
  vaSpread.Col = 6
  ThisAmt = ReplaceString(vaSpread.Text, ",", "")
  If Val(ThisAmt) > 0 And ThisAmtPct = "" Then
    frmMessage.Label1.Caption = "A value has been entered for 'Fixed Federal Figure' on row " + CStr(vaSpread.Row) + " but there is no federal amount or percent saved. Please enter either amount or percent in column #5 before continuing."
    frmMessage.Label1.Top = 700
    frmMessage.Show vbModal
    vaSpread.SetFocus
    vaSpread.OperationMode = OperationModeNormal
    vaSpread.SetActiveCell 5, vaSpread.Row
    CheckThisAmtPctAndFigure = False
    Exit Function
  End If
  If Val(ThisAmt) > 100 And Mid(ThisAmtPct, 1, 1) = "P" Then
    frmMessage.Label1.Caption = "Percent has been entered for 'Federal Amt/Pct' on row " + CStr(vaSpread.Row) + " but the amount for 'Fixed Federal Figure' is more than 100%. This percent figure is invalid."
    frmMessage.Label1.Top = 700
    frmMessage.Show vbModal
    vaSpread.SetFocus
    vaSpread.OperationMode = OperationModeNormal
    vaSpread.SetActiveCell 6, vaSpread.Row
    CheckThisAmtPctAndFigure = False
    Exit Function
  End If
  vaSpread.Row = ThisRow
  vaSpread.Col = 11
  ThisAmtPct = QPTrim$(vaSpread.Text)
  vaSpread.Col = 12
  ThisAmt = ReplaceString(vaSpread.Text, ",", "")
  If Val(ThisAmt) > 0 And ThisAmtPct = "" Then
    frmMessage.Label1.Caption = "A value has been entered for 'Fixed State Figure' on row " + CStr(vaSpread.Row) + " but there is no state amount or percent saved. Please enter either amount or percent in column #11 before continuing."
    frmMessage.Label1.Top = 700
    frmMessage.Show vbModal
    vaSpread.SetFocus
    vaSpread.OperationMode = OperationModeNormal
    vaSpread.SetActiveCell 11, vaSpread.Row
    CheckThisAmtPctAndFigure = False
    Exit Function
  End If
  If Val(ThisAmt) > 100 And Mid(ThisAmtPct, 1, 1) = "P" Then
    frmMessage.Label1.Caption = "Percent has been entered for 'State Amt/Pct' on row " + CStr(vaSpread.Row) + " but the amount for 'Fixed State Figure' is more than 100%. This percent figure is invalid."
    frmMessage.Label1.Top = 700
    frmMessage.Show vbModal
    vaSpread.SetFocus
    vaSpread.OperationMode = OperationModeNormal
    vaSpread.SetActiveCell 12, vaSpread.Row
    CheckThisAmtPctAndFigure = False
    Exit Function
  End If
  
End Function
