VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmTaxAbsList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Abstract List"
   ClientHeight    =   8736
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   11640
   Icon            =   "frmTaxAbsList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8736
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpListReal 
      Height          =   3396
      Left            =   1500
      TabIndex        =   3
      Top             =   2400
      Width           =   8652
      _Version        =   196608
      _ExtentX        =   15261
      _ExtentY        =   5990
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
      Object.TabStop         =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Columns         =   6
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
      ColDesigner     =   "frmTaxAbsList.frx":08CA
   End
   Begin LpLib.fpList fpListPers 
      Height          =   3240
      Left            =   1560
      TabIndex        =   6
      Top             =   2400
      Width           =   8412
      _Version        =   196608
      _ExtentX        =   14838
      _ExtentY        =   5715
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Columns         =   6
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
      ColDesigner     =   "frmTaxAbsList.frx":0E57
   End
   Begin EditLib.fpText fptxtMatched 
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6600
      Width           =   1215
      _Version        =   196608
      _ExtentX        =   2143
      _ExtentY        =   661
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
      Height          =   540
      Left            =   3480
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7320
      Width           =   2055
      _Version        =   131072
      _ExtentX        =   3625
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmTaxAbsList.frx":13EE
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDelete 
      Height          =   540
      Left            =   6120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7320
      Width           =   2055
      _Version        =   131072
      _ExtentX        =   3625
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmTaxAbsList.frx":1604
   End
   Begin EditLib.fpText fptxtCust 
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   1320
      Width           =   4695
      _Version        =   196608
      _ExtentX        =   8281
      _ExtentY        =   661
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Matched:"
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
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   3975
      Left            =   1260
      Top             =   2160
      Width           =   9135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Abstract List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3143
      TabIndex        =   2
      Top             =   510
      Width           =   5295
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1493
      Top             =   360
      Width           =   8655
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1493
      Top             =   240
      Width           =   8655
   End
End
Attribute VB_Name = "frmTaxAbsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
Private Sub cmdExit_Click()
  Unload Me
  DoEvents
End Sub

Private Sub cmdDelete_Click()
  Dim ThisPin$
  Dim ThisCust$
  Dim PersVal$
  Dim ThisMap$
  Dim ThisBlock$
  Dim ThisLot$
  Dim MobVal$
  Dim MerchVal$
  Dim FarmVal$
  Dim MachVal$
  Dim OK2Exit As Boolean
  
  OK2Exit = False
  ThisCust = QPTrim$(fptxtCust.Text)
  If frmTaxAbsMaint.fptxtChoice.Text = "pers" Then
    If fpListPers.ListIndex = -1 Then
      frmTaxMsg.Label1.Caption = "Please select one of the records in the list."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show vbModal
      fpListPers.ListIndex = 0
      Exit Sub
    End If
    frmTaxMsgWOpts.Label1.Caption = "Are you sure you want to delete this property?"
    frmTaxMsgWOpts.Label1.Top = 900
    frmTaxMsgWOpts.cmdCont.Text = "F10 OK To Delete"
    frmTaxMsgWOpts.cmdExit.Text = "ESC Don't Delete"
    frmTaxMsgWOpts.Show vbModal
    If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
      Unload frmTaxMsgWOpts
      fpListPers.ListIndex = 0
      Exit Sub
    End If
    fpListPers.Row = fpListPers.ListIndex
    fpListPers.col = 0
    ThisPin = QPTrim$(fpListPers.ColText)
    fpListPers.col = 1
    PersVal = QPTrim$(fpListPers.ColText)
    fpListPers.col = 2
    MobVal = QPTrim$(fpListPers.ColText)
    fpListPers.col = 3
    MerchVal = QPTrim$(fpListPers.ColText)
    fpListPers.col = 4
    FarmVal = QPTrim$(fpListPers.ColText)
    fpListPers.col = 5
    MachVal = QPTrim$(fpListPers.ColText)
    ReDim PersRecs(0 To 0) As Long
    Call GetPersRecList(PersRecs(), GCustNum, ThisCust)
    Call DelPersAbstract(PersRecs(), fpListPers.ListIndex + 1, GCustNum)
    MainLog ("PERSONAL PROPERTY DELETION: User deleted the following personal property for : " + ThisCust + " - Pin # " + ThisPin + " - Personal Value: " + PersVal + " - Mobile Value: " + MobVal + " - Merchant Value: " + MerchVal + " - Farm Value: " + FarmVal + " - Machine Value: " + MachVal + ".")
    If PersRecs(0) = 0 Then
      OK2Exit = True
    Else
      fpListPers.Clear
    End If
  ElseIf frmTaxAbsMaint.fptxtChoice.Text = "real" Then
    If fpListReal.ListIndex = -1 Then
      frmTaxMsg.Label1.Caption = "Please select one of the records in the list."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show vbModal
      fpListReal.ListIndex = 0
      Exit Sub
    End If
    frmTaxMsgWOpts.Label1.Caption = "Are you sure you want to delete this property?"
    frmTaxMsgWOpts.Label1.Top = 900
    frmTaxMsgWOpts.cmdCont.Text = "F10 OK To Delete"
    frmTaxMsgWOpts.cmdExit.Text = "ESC Don't Delete"
    frmTaxMsgWOpts.Show vbModal
    If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
      Unload frmTaxMsgWOpts
      fpListReal.ListIndex = 0
      Exit Sub
    End If
    fpListReal.Row = fpListReal.ListIndex
    fpListReal.col = 0
    ThisPin = QPTrim$(fpListReal.ColText)
    fpListReal.col = 1
    ThisMap = QPTrim$(fpListReal.ColText)
    fpListReal.col = 2
    ThisBlock = QPTrim$(fpListReal.ColText)
    fpListReal.col = 3
    ThisLot = QPTrim$(fpListReal.ColText)
    fpListReal.col = 4
    PersVal = QPTrim$(fpListReal.ColText)
    ReDim RealRecs(0 To 0) As Long
    Call GetRealRecList(RealRecs(), GCustNum, ThisCust)
    Call DelRealAbstract(RealRecs(), fpListPers.ListIndex + 1, GCustNum)
    MainLog ("REAL PROPERTY DELETION: User deleted the following real property for : " + ThisCust + " - Pin Number: " + ThisPin + " - Map: " + ThisMap + " - Block: " + ThisBlock + " - Lot: " + ThisLot + " - Value: " + PersVal + ".")
    If RealRecs(0) = 0 Then
      OK2Exit = True
    Else
      fpListReal.Clear
    End If
  End If
  
  frmTaxMsg.Label1.Caption = "The property was deleted successfully."
  frmTaxMsg.Label1.Top = 900
  frmTaxMsg.Show vbModal
  If OK2Exit = False Then
    Call Loadme
  Else
    Unload Me
    DoEvents
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%D"
      Call cmdDelete_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call Loadme
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxAbsList.")
      Call Terminate
      End
    End If
  End If

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If

End Sub

Private Sub Loadme()
  Dim TaxRec As TaxCustType
  Dim THandle As Integer
  Dim NumOfTaxRecs As Long
  Dim PersRec As PersonalRecType
  Dim PersHandle As Integer
  Dim NumOfPersRecs As Long
  Dim RealRec As PropertyRecType
  Dim RealHandle As Integer
  Dim NumOfRealRecs As Long
  Dim x As Long
  Dim ThisPropRec As Long
  Dim ThisPersRec As Long
  Dim ThisCnt As Integer
  Dim One As String * 6
  Dim Two As String * 6
  Dim Three As String * 6
  
  OpenTaxCustFile THandle, NumOfTaxRecs
  Get THandle, GCustNum, TaxRec
  Close THandle
  fptxtCust.Text = QPTrim$(TaxRec.CustName)
  ThisPropRec = TaxRec.FirstPropRec
  ThisPersRec = TaxRec.FirstPersRec
  ThisCnt = 0
  If frmTaxAbsMaint.fptxtChoice.Text = "real" Then
    fpListReal.Visible = True
    fpListPers.Visible = False
    OpenTaxPropFile RealHandle, NumOfRealRecs
    Do While ThisPropRec > 0
      Get RealHandle, ThisPropRec, RealRec
      ThisCnt = ThisCnt + 1
      RSet One = QPTrim$(RealRec.Map)
      RSet Two = QPTrim$(RealRec.BLOCK)
      RSet Three = QPTrim$(RealRec.LOTNUMB)
      
      fpListReal.InsertRow = RealRec.RealPin + Chr(9) + One + Chr(9) + Two + Chr(9) + Three + Chr(9) + Using$("$#,###,##0.00", RealRec.PROPVALU) + Chr(9) + Using$("#########0", RealRec.CustPin)
      ThisPropRec = RealRec.NextRec
    Loop
    Close RealHandle
    fpListReal.ListIndex = 0
  ElseIf frmTaxAbsMaint.fptxtChoice.Text = "pers" Then
    fpListReal.Visible = False
    fpListPers.Visible = True
    OpenTaxPersFile PersHandle, NumOfPersRecs
    Do While ThisPersRec > 0
      Get PersHandle, ThisPersRec, PersRec
      ThisCnt = ThisCnt + 1
      fpListPers.InsertRow = PersRec.PropPin + Chr(9) + Using$("$#,###,##0.00", PersRec.PersVal) + Chr(9) + Using$("$#,###,##0.00", PersRec.MHVALUE) + Chr(9) + Using$("$#,###,##0.00", PersRec.MCVALUE) + Chr(9) + Using$("$#,###,##0.00", PersRec.CVALUE) + Chr(9) + Using$("$#,###,##0.00", PersRec.MTVALUE)
      ThisPersRec = PersRec.NextRec
    Loop
    Close PersHandle
    fpListPers.ListIndex = 0
  End If
  fptxtMatched.Text = CStr(ThisCnt)
  
End Sub

Private Sub fpListPers_DblClick()
  Call cmdDelete_Click
End Sub

Private Sub fpListReal_DblClick()
  Call cmdDelete_Click

End Sub
