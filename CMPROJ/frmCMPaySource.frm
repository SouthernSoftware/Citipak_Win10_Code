VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmCMPaySource 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Payment Source"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   1890
   ClientWidth     =   12210
   Icon            =   "frmCMPaySource.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboPaySource 
      Height          =   345
      Left            =   5205
      TabIndex        =   0
      Top             =   3315
      Width           =   3825
      _Version        =   196608
      _ExtentX        =   6747
      _ExtentY        =   609
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BackColor       =   16777215
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   1
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   2
      SearchMethod    =   2
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   0
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
      Appearance      =   0
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
      AutoSearchFill  =   -1  'True
      AutoSearchFillDelay=   100
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmCMPaySource.frx":08CA
   End
   Begin VB.CommandButton cmdTaxAdjust 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Tax Adjustments"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3790
      TabIndex        =   15
      Top             =   6240
      Width           =   2196
   End
   Begin VB.CommandButton cmdVoidPayments 
      Caption         =   "&Void Payments"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1398
      TabIndex        =   4
      Top             =   6240
      Width           =   2172
   End
   Begin VB.CommandButton cmdUtilAdj 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Utility Adjustments"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   6206
      TabIndex        =   5
      Top             =   6240
      Width           =   2196
   End
   Begin VB.CommandButton cmdBLAdj 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&BL Adjustments"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   8622
      TabIndex        =   6
      Top             =   6240
      Width           =   2196
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   336
      Left            =   0
      TabIndex        =   8
      Top             =   8532
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "7:47 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "2/4/2020"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdExit 
      Height          =   480
      Left            =   8040
      TabIndex        =   3
      Top             =   4656
      Width           =   1332
      _Version        =   131072
      _ExtentX        =   2350
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmCMPaySource.frx":0BF1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdOK 
      Height          =   480
      Left            =   6216
      TabIndex        =   2
      Top             =   4656
      Width           =   1332
      _Version        =   131072
      _ExtentX        =   2350
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmCMPaySource.frx":0DCD
   End
   Begin EditLib.fpDateTime txtDate1 
      Height          =   348
      Left            =   5184
      TabIndex        =   1
      Top             =   3744
      Width           =   1884
      _Version        =   196608
      _ExtentX        =   3323
      _ExtentY        =   614
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
      ThreeDInsideHighlightColor=   -2147483637
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
      ButtonStyle     =   2
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
      InvalidColor    =   -2147483643
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483643
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "10/03/2001"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19800101"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "20010101"
      TimeDefault     =   "000000"
      TimeStyle       =   0
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      PopUpType       =   1
      DateCalcY2KSplit=   60
      CaretPosition   =   0
      IncYear         =   1
      IncMonth        =   1
      IncDay          =   1
      IncHour         =   1
      IncMinute       =   1
      IncSecond       =   1
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   948
      Left            =   1230
      Top             =   6000
      Width           =   9756
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Date:"
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
      Height          =   396
      Index           =   1
      Left            =   2976
      TabIndex        =   14
      Top             =   3816
      Width           =   2088
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H8000000E&
      BorderWidth     =   3
      Height          =   1356
      Left            =   2448
      Top             =   3024
      Width           =   7332
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Source:"
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
      Left            =   2976
      TabIndex        =   13
      Top             =   3336
      Width           =   2088
   End
   Begin VB.Label lblOperator 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   4368
      TabIndex        =   12
      Top             =   2688
      Width           =   852
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operator Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   312
      Index           =   0
      Left            =   1944
      TabIndex        =   11
      Top             =   2736
      Width           =   2304
   End
   Begin VB.Label lblOperName 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   7344
      TabIndex        =   10
      Top             =   2688
      Width           =   2436
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operator Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   312
      Left            =   5424
      TabIndex        =   9
      Top             =   2736
      Width           =   1824
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Management Payment Source"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3288
      TabIndex        =   7
      Top             =   1632
      Width           =   5652
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3228
      Top             =   1392
      Width           =   5772
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3228
      Top             =   1272
      Width           =   5772
   End
End
Attribute VB_Name = "frmCMPaySource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim Oper As String
Dim BadDate As Boolean
Private Sub CheckPayDate()
Dim payDate As String
  payDate$ = txtDate1.Text
  If Val(Left$(payDate$, 2)) < 1 Or Val(Left$(payDate$, 2)) > 12 Then
    If Val(Mid$(payDate$, 4, 2)) < 1 Or Val(Mid$(payDate$, 4, 2)) > 31 Then
      BadDate = True
    Else
      BadDate = False
    End If
  Else
    BadDate = False
  End If
End Sub

Private Sub cmdBLAdj_Click()
  Dim CMSetuplen As Integer
  ReDim CMSetUpRec(1) As CMSetupType
  CMSetuplen = Len(CMSetUpRec(1))
  LoadCMSetUpFile CMSetUpRec(), CMSetuplen
  If QPTrim(CMSetUpRec(1).Pass4Adj) = "Y" Then
  'do the password screen
    frmPassWord.Callingfrm = 3
    frmPassWord.Show 1
  ElseIf QPTrim(CMSetUpRec(1).Pass4Adj) = "F" Then
  'get opernum properties if full access then goahead on
    If LevelPass = 1 Then
      Load frmBLAdjustBal
      frmBLAdjustBal.Show
      Unload Me
    Else
      MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
    End If
  Else 'nobody cares
    Load frmBLAdjustBal
    frmBLAdjustBal.Show
    Unload Me
  End If
  Erase CMSetUpRec
End Sub

Private Sub cmdTaxAdjust_Click()
  Dim CMSetuplen As Integer, Oktogoon As Boolean
'  Dim intHasTaxes As Integer
  ReDim CMSetUpRec(1) As CMSetupType
  Oktogoon = False
  CMSetuplen = Len(CMSetUpRec(1))
  LoadCMSetUpFile CMSetUpRec(), CMSetuplen
  
  Select Case intHasTaxes
  Case 1 'NC Taxes
    Dim TaxMasterRec As TaxMasterType
    Dim TMHandle As Integer
    OpenTaxSetUpFile TMHandle
    Get TMHandle, 1, TaxMasterRec
    Close TMHandle

    If RevsAndGLsOK(Me, TaxMasterRec.TaxYear) = False Then
        Exit Sub
    End If
    If QPTrim(CMSetUpRec(1).Pass4Adj) = "Y" Then
        'do the password screen
        frmPassWord.Callingfrm = 4
        frmPassWord.Show 1
    ElseIf QPTrim(CMSetUpRec(1).Pass4Adj) = "F" Then
        'get opernum properties if full access then goahead on
        If LevelPass = 1 Then
            Load frmTaxAdjustments
            frmTaxAdjustments.Show
            Unload Me
        Else
            MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
        End If
    Else 'nobody cares
        Load frmTaxAdjustments
        frmTaxAdjustments.Show
        Unload Me
    End If
  
  Case 2 'VA Taxes
    frmVATaxBillPostOpt.Show vbModal
    If frmVATaxBillPostOpt.fptxtPostType.Text = "Real" Then
        If QPTrim(CMSetUpRec(1).Pass4Adj) = "Y" Then
            'do the password screen
            frmPassWord.Callingfrm = 5
            frmPassWord.Show 1
        ElseIf QPTrim(CMSetUpRec(1).Pass4Adj) = "F" Then
        'get opernum properties if full access then goahead on
            If LevelPass = 1 Then
                Load frmVATaxAdjustments
                frmVATaxAdjustments.Show
                Unload Me
            Else
                MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
            End If
        Else 'nobody cares
            Load frmVATaxAdjustments
            frmVATaxAdjustments.Show
            Unload Me
        End If
    ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Personal" Then
        If QPTrim(CMSetUpRec(1).Pass4Adj) = "Y" Then
            'do the password screen
            frmPassWord.Callingfrm = 6
            frmPassWord.Show 1
        ElseIf QPTrim(CMSetUpRec(1).Pass4Adj) = "F" Then
            'get opernum properties if full access then goahead on
            If LevelPass = 1 Then
                Load frmVATaxPAdjustments
                frmVATaxPAdjustments.Show
                Unload Me
            Else
                MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
            End If
        Else 'nobody cares
            Load frmVATaxPAdjustments
            frmVATaxPAdjustments.Show
            Unload Me
        End If
    ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Exit" Then
        DoEvents
        Unload frmVATaxBillPostOpt
        Exit Sub
    End If
  
  Case Else ' Doesn't have taxes
  
  End Select
  
  
'  If Exist(UBPath$ + "CitiTaxes.EXE") Then
'    Dim TaxMasterRec As TaxMasterType
'    Dim TMHandle As Integer
'    OpenTaxSetUpFile TMHandle
'    Get TMHandle, 1, TaxMasterRec
'    Close TMHandle
'
'    If RevsAndGLsOK(Me, TaxMasterRec.TaxYear) = False Then
'      Exit Sub
'    End If
'    If QPTrim(CMSetUpRec(1).Pass4Adj) = "Y" Then
'    'do the password screen
'      frmPassWord.Callingfrm = 4
'      frmPassWord.Show 1
'    ElseIf QPTrim(CMSetUpRec(1).Pass4Adj) = "F" Then
'    'get opernum properties if full access then goahead on
'      If LevelPass = 1 Then
'        Load frmTaxAdjustments
'        frmTaxAdjustments.Show
'        Unload Me
'      Else
'        MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
'      End If
'    Else 'nobody cares
'      Load frmTaxAdjustments
'      frmTaxAdjustments.Show
'      Unload Me
'    End If
'
'  ElseIf Exist(UBPath$ + "VACitiTax.EXE") Then
'    frmVATaxBillPostOpt.Show vbModal
'    If frmVATaxBillPostOpt.fptxtPostType.Text = "Real" Then
'      If QPTrim(CMSetUpRec(1).Pass4Adj) = "Y" Then
'      'do the password screen
'        frmPassWord.Callingfrm = 5
'        frmPassWord.Show 1
'      ElseIf QPTrim(CMSetUpRec(1).Pass4Adj) = "F" Then
'      'get opernum properties if full access then goahead on
'        If LevelPass = 1 Then
'          Load frmVATaxAdjustments
'          frmVATaxAdjustments.Show
'          Unload Me
'        Else
'          MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
'        End If
'      Else 'nobody cares
'        Load frmVATaxAdjustments
'        frmVATaxAdjustments.Show
'        Unload Me
'      End If
'    ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Personal" Then
'      If QPTrim(CMSetUpRec(1).Pass4Adj) = "Y" Then
'      'do the password screen
'        frmPassWord.Callingfrm = 6
'        frmPassWord.Show 1
'      ElseIf QPTrim(CMSetUpRec(1).Pass4Adj) = "F" Then
'      'get opernum properties if full access then goahead on
'        If LevelPass = 1 Then
'          Load frmVATaxPAdjustments
'          frmVATaxPAdjustments.Show
'          Unload Me
'        Else
'          MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
'        End If
'      Else 'nobody cares
'        Load frmVATaxPAdjustments
'        frmVATaxPAdjustments.Show
'        Unload Me
'      End If
'    ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Exit" Then
'      DoEvents
'      Unload frmVATaxBillPostOpt
'      Exit Sub
'    End If
'  End If
  Erase CMSetUpRec
End Sub

Private Sub cmdUtilAdj_Click()
  Dim CMSetuplen As Integer
  ReDim CMSetUpRec(1) As CMSetupType
  CMSetuplen = Len(CMSetUpRec(1))
  LoadCMSetUpFile CMSetUpRec(), CMSetuplen
  If QPTrim(CMSetUpRec(1).Pass4Adj) = "Y" Then
  'do the password screen
    frmPassWord.Callingfrm = 2
    frmPassWord.Show 1
  ElseIf QPTrim(CMSetUpRec(1).Pass4Adj) = "F" Then
  'get opernum properties if full access then goahead on
    If LevelPass = 1 Then
      Load frmUBAdjustmentEntry
      frmUBAdjustmentEntry.Show
      Unload Me
    Else
      MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
    End If
  Else 'nobody cares
    Load frmUBAdjustmentEntry
    frmUBAdjustmentEntry.Show
    Unload Me
  End If
  Erase CMSetUpRec
End Sub

Private Sub cmdVoidPayments_Click()
  Dim CMSetuplen As Integer
  ReDim CMSetUpRec(1) As CMSetupType
  CMSetuplen = Len(CMSetUpRec(1))
  LoadCMSetUpFile CMSetUpRec(), CMSetuplen
  If QPTrim(CMSetUpRec(1).Pass4Voids) = "Y" Then
  'do the password screen
    frmPassWord.Callingfrm = 1
    frmPassWord.Show 1
  ElseIf QPTrim(CMSetUpRec(1).Pass4Voids) = "F" Then
  'get opernum properties if full access then goahead on
    If LevelPass = 1 Then
      Load frmVoidSearch
      frmVoidSearch.Show
      Unload Me
    Else
      MsgBox "Your password does not allow access to this option.", vbOKOnly, "Access Denied"
    End If
  Else 'nobody cares
    Load frmVoidSearch
    frmVoidSearch.Show
    Unload Me
  End If
  Erase CMSetUpRec
End Sub

Private Sub cmdOk_Click()
  Dim Today As String, chkthedate As Integer, entdate As Integer
  Dim FntSize As Integer, DCSetuplen As Integer
  Dim TMHandle As Integer
  Dim ThisTYear As Integer

  CheckPayDate
  If BadDate = False Then
    'do stuff
    Today = Format(Now, "mm/dd/yyyy")
    chkthedate = Date2Num(Today)
    entdate = Date2Num(txtDate1)
    If entdate > (chkthedate + 30) Or entdate < (chkthedate - 30) Then
      
      CMLog "Date outOrange entered payentry, give opt to cancel- OPER:" + Str$(OperNum)
      ReDim MsgText(0 To 5) As String
      FntSize = frmMsgDialog.Label(1).FontSize
      frmMsgDialog.Label(1).FontSize = (FntSize + 2)
      frmMsgDialog.Label(3).FontSize = (FntSize + 2)
      frmMsgDialog.Label(4).FontSize = (FntSize + 2)
      frmMsgDialog.Label(2).FontSize = (FntSize + 2)
      MsgText(0) = "WARNING:"
      MsgText(1) = ""
      MsgText(2) = "DATE Entered is NOT Within"
      MsgText(3) = "Monthly Date Range."
      MsgText(4) = "Select OK to continue"
      MsgText(5) = "or Cancel to Change."
      If GetOKorNot(MsgText()) Then
        CMLog "Continue CMpay entry with out of range date-" + txtDate1.Text
      Else
        CMLog "Cancel CMpay entry so can check date."
        Exit Sub
      End If
    End If

    savePDate
    Select Case fpcboPaySource.Text

      Case "Utility Billing Payment":
        frmPayUtilEntry.Wheretogo frmCMMainMenu, frmCMMainMenu, , txtDate1
        frmPayUtilEntry.Show
        DoEvents
        Unload Me
      Case "Utility Deposit Entry":
        frmPayDepEntry.Wheretogo frmCMMainMenu, frmCMMainMenu, , txtDate1
        frmPayDepEntry.Show
        DoEvents
        Unload Me
      Case "Business License Payment":
        frmPayBLEntry.Wheretogo frmCMMainMenu, frmCMMainMenu, , txtDate1
        frmPayBLEntry.Show
        DoEvents
        Unload Me
      Case "Miscellaneous Payment":
        frmPayMiscEntry.Wheretogo frmCMMainMenu, frmCMMainMenu, , txtDate1
        frmPayMiscEntry.Show
        DoEvents
        Unload Me
      Case "Tax Billing Payment":
'        Dim TaxMasterRec As TaxMasterType
'        OpenTaxSetUpFile TMHandle
'        Get TMHandle, 1, TaxMasterRec
'        Close TMHandle
'        ThisTYear = 0
'         If CheckTaxYear(ThisTYear) = False Then
'           If TaxMsgWOpts(400, "The current system tax year (" + CStr(TaxMasterRec.TaxYear) + ") comes before some of the tax years for tax bills posted (" + CStr(ThisTYear) + "). This can cause payment posting problems if discounts are allowed. If you wish to change the system tax year then press ESC to escape and go to the System Setup screen to edit. Otherwise press F10 to continue as is.", "F10 Continue", "ESC Escape") = "abort" Then
'             Unload frmTaxMsgWOpts
'             Exit Sub
'           Else
'             Unload frmTaxMsgWOpts
'             TXLog ("WARNING: User issued a warning that the system tax year (" + CStr(TaxMasterRec.TaxYear) + ") comes before some of the tax years for tax bills posted (" + CStr(ThisTYear) + "). This can cause payment posting problems. User elected to continue anyway.")
'           End If
'         End If
'
'        If RevsAndGLsOK(Me, TaxMasterRec.TaxYear) = False Then
'          Exit Sub
'        End If
        frmTaxPaymentEntry.Wheretogo frmCMMainMenu, frmCMMainMenu, , txtDate1
        frmTaxPaymentEntry.Show
        DoEvents
        Unload Me
      Case "VA Tax Billing Payment":
'        Dim TaxMaster As VATaxMasterType
'        OpenVATaxSetUpFile TMHandle
'        Get TMHandle, 1, TaxMaster
'        Close TMHandle
'        ThisTYear = 0
        frmVATaxBillPostOpt.Show vbModal
        If frmVATaxBillPostOpt.fptxtPostType.Text = "Real" Then
'          If VACheckTaxYear("R", ThisTYear) = False Then
'            If VATaxMsgWOpts(400, "The current real system tax year (" + CStr(TaxMaster.RTaxYear) + ") comes before some of the tax years for real tax bills posted (" + CStr(ThisTYear) + "). This can cause payment posting problems if discounts are allowed. If you wish to change the real system tax year then press ESC to escape and go to the System Setup screen to edit. Otherwise press F10 to continue as is.", "F10 Continue", "ESC Escape") = "abort" Then
'              Unload frmVATaxMsgWOpts
'              Exit Sub
'            Else
'              Unload frmVATaxMsgWOpts
'              TXLog ("WARNING: User issued a warning that the current real system tax year (" + CStr(TaxMaster.RTaxYear) + ") comes before some of the tax years for tax bills posted (" + CStr(ThisTYear) + "). This can cause payment posting problems. User elected to continue anyway.")
'            End If
'          End If
'          If VARevsAndGLsOK(Me, TaxMaster.RTaxYear, "R") = False Then
'            Exit Sub
'          End If
          frmVATaxPaymentEntry.Wheretogo frmCMMainMenu, frmCMMainMenu, , txtDate1
          frmVATaxPaymentEntry.Show
          DoEvents
          Unload frmVATaxBillPostOpt
        ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Personal" Then
'          If VACheckTaxYear("P", ThisTYear) = False Then
'            If VATaxMsgWOpts(400, "The current personal system tax year (" + CStr(TaxMaster.PTaxYear) + ") comes before some of the tax years for personal tax bills posted (" + CStr(ThisTYear) + "). This can cause payment posting problems if discounts are allowed. If you wish to change the personal system tax year then press ESC to escape and go to the System Setup screen to edit. Otherwise press F10 to continue as is.", "F10 Continue", "ESC Escape") = "abort" Then
'              Unload frmVATaxMsgWOpts
'              Exit Sub
'            Else
'              Unload frmVATaxMsgWOpts
'              TXLog ("WARNING: User issued a warning that the current personal system tax year (" + CStr(TaxMaster.PTaxYear) + ") comes before some of the tax bills posted (" + CStr(ThisTYear) + "). This can cause payment posting problems. User elected to continue anyway.")
'            End If
'          End If
'          If VARevsAndGLsOKP(Me, TaxMaster.PTaxYear, "P") = False Then
'            Exit Sub
'          End If
          frmVATaxPersPaymentEntry.Wheretogo frmCMMainMenu, frmCMMainMenu, , txtDate1
          frmVATaxPersPaymentEntry.Show
          DoEvents
          Unload frmVATaxBillPostOpt
        ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Exit" Then
          DoEvents
          Unload frmVATaxBillPostOpt
          Exit Sub
        End If
        Unload Me
      Case "Vehicle Decal Purchase":
        'if decal files the allow payment
        If Exist(UBPath$ + "DCSetup.DAT") Then
          ReDim DCSetup(1) As DCSetupType
          LoadDCSetUpFile DCSetup(), DCSetuplen
          If DCSetup(1).DCVers = "205" Then
            frmPayDecalEntry.Wheretogo frmCMMainMenu, frmCMMainMenu, , txtDate1
            frmPayDecalEntry.Show
            DoEvents
            Unload Me
          End If
        End If
      Case Else:
        MsgBox "Invalid Selection", vbOKOnly, "Invalid Source"
    End Select
  Else
    MsgBox "Invalid Date", vbOKOnly, "Invalid Entry"
  End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If fpCmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        CMLog "Closed via SelectPaySource by " + PWUser$ + " operator-" + Oper$
        CitiTerminate
      End If
    End If
  End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  Select Case KeyCode
'    Case vbKeyDown, vbKeyReturn:
'      SendKeys "{Tab}"
'      KeyCode = 0
'    Case vbKeyUp:
'      SendKeys "+{Tab}"
'      KeyCode = 0
    Case vbKeyEscape:
      KeyCode = 0
      DoEvents
      Call fpCmdExit_Click
    Case vbKeyF10:
      KeyCode = 0
      DoEvents
      Call cmdOk_Click
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TownName$
  txtDate1.Text = Format(Now, "mm/dd/yyyy")
  fpcboPaySource.AddItem "Utility Billing Payment"
  fpcboPaySource.AddItem "Utility Deposit Entry"
  fpcboPaySource.AddItem "Business License Payment"
  fpcboPaySource.AddItem "Miscellaneous Payment"

    Select Case intHasTaxes
    Case 1 'NC Taxes
        fpcboPaySource.AddItem "Tax Billing Payment"
    Case 2 'VA Taxes
        fpcboPaySource.AddItem "VA Tax Billing Payment"
    Case Else
        cmdTaxAdjust.Enabled = False
    End Select
  
'  If Exist(UBPath$ + "CitiTaxes.EXE") Then
'    fpcboPaySource.AddItem "Tax Billing Payment"
'  ElseIf Exist(UBPath$ + "VACitiTax.EXE") Then
'    fpcboPaySource.AddItem "VA Tax Billing Payment"
'  Else
'    cmdTaxAdjust.Enabled = False
'  End If
  
  If Exist(UBPath$ + "DCCust.DAT") Then
'    If Exist(UBPath$ + "DC.EXE") Then
      fpcboPaySource.AddItem "Vehicle Decal Purchase"
'    End If
  End If
  
  lblOperator = OperNum
  lblOperName.Caption = PWUser
  Oper$ = QPTrim(lblOperator.Caption)
  GetPayDate
  CMLog " IN Oper " + Oper$ + "CMPaySource"
End Sub
Public Sub GetPayDate()
  Dim lenRP As Integer, RP1 As Integer, gpay As Integer
  Dim RcptPrnFile As ReceiptPRNType
  RP1 = FreeFile
  lenRP = Len(RcptPrnFile)
  If Exist(RcptFileName$) Then
    Open RcptFileName$ For Random Shared As RP1 Len = lenRP
    Get RP1, 1, RcptPrnFile
      gpay = RcptPrnFile.PaymDate
    Close
  End If
  If gpay > Date2Num(txtDate1) Then
    txtDate1.Text = Num2Date(gpay)
  End If

End Sub
Private Sub savePDate()
  Dim RP1 As Integer, lenRP As Integer
  Dim RcptPrnFile As ReceiptPRNType
  RP1 = FreeFile
  lenRP = Len(RcptPrnFile)
  Open RcptFileName$ For Random Shared As RP1 Len = lenRP
    Get #RP1, 1, RcptPrnFile
    RcptPrnFile.PaymDate = Date2Num(txtDate1)
    Put #RP1, 1, RcptPrnFile
  Close
  CMLog PWUser + " " + Str(OperNum) + " Log in payments with date - " + txtDate1
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ''' Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
'
'  If Me.Visible Then
'    Temp_Class.ResizeControls Me
'    DoEvents
'  End If
End Sub

Private Sub fpcboPaySource_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboPaySource.ListDown = True
  End If
  If fpcboPaySource.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      txtDate1.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpCmdExit.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub


Private Sub fpCmdExit_Click()
  frmCMMainMenu.Show
  Unload Me
End Sub

Private Sub txtDate1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
    cmdOk.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    fpcboPaySource.SetFocus
  End If
End Sub
