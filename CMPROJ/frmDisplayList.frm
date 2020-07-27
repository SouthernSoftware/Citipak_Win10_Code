VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmDisplayList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caption"
   ClientHeight    =   4875
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   9300
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin LpLib.fpList fpList1 
      Height          =   3210
      Left            =   90
      TabIndex        =   0
      Top             =   435
      Width           =   9090
      _Version        =   196608
      _ExtentX        =   16034
      _ExtentY        =   5662
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
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
      SelMax          =   1
      AutoSearch      =   1
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   0
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
      ScrollBarV      =   0
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   2
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   0   'False
      DataAutoSizeCols=   0
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
      ColDesigner     =   "frmDisplayList.frx":0000
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdOK 
      Height          =   480
      Left            =   6408
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4248
      Width           =   1212
      _Version        =   131072
      _ExtentX        =   2138
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmDisplayList.frx":0358
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdExit 
      Height          =   480
      Left            =   7776
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4248
      Width           =   1212
      _Version        =   131072
      _ExtentX        =   2138
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmDisplayList.frx":0531
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7344
      TabIndex        =   4
      Top             =   72
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3072
      TabIndex        =   3
      Top             =   72
      Width           =   2148
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer/Owner"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   456
      TabIndex        =   2
      Top             =   72
      Width           =   1908
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "To Select Double-Click Item or Highlight and Click Ok."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   1
      Top             =   4344
      Width           =   5604
   End
End
Attribute VB_Name = "frmDisplayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RecNo As Long, AcctNum As Long
Dim fromform As Form, toform As Form, codeopt As Integer, SrchDel As Integer
Public Sub Wheretogo(xfrm As Form, tfrm As Form, Optional opt As Integer, Optional SDel As Integer)
  
  Set fromform = xfrm
  Set toform = tfrm
  If opt <> 0 Then
    codeopt = opt
  Else
    codeopt = 0
  End If
  If SDel <> 0 Then
    SrchDel = SDel
  Else
    SrchDel = 0
  End If
    
   Me.fpList1.ListIndex = 0
End Sub

Private Sub fpCmdExit_Click()
  SearchRec = 0
  codeopt = 0
  DoEvents
  ActivateControls frmCustEditLookUP
  Unload frmDisplayList
End Sub

Private Sub Form_Activate()
  SearchRec& = 0
'  Me.fpList1.ListIndex = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      KeyCode = 0
      Call fpCmdExit_Click
    Case vbKeyF10, vbKeyReturn
      KeyCode = 0
      Call fpCmdOk_Click
    Case Else:
  End Select
End Sub

Private Sub fpCmdOk_Click()
  If fpList1.SelCount > 0 Then
    Call fpList1_DblClick
  End If
End Sub

Private Sub fpList1_DblClick()
  Dim xx As Integer
  fpList1.col = 1                       'switch to the hidden RecNo. column
  SearchRec& = Val(fpList1.ColText) 'get customer recno
  'If codeopt <> 0 Then
  If SearchRec& > 0 Then   'if user selected an account
    If SrchDel = 1 Then
      If OKDeleteCust(SearchRec&) Then
        DeActivateControls Me
        frmInfo.Label1 = "Loading. . ."
        frmInfo.Show
        DoEvents
      'here
        toform.fpCustRecNo = QPTrim$(Str$(SearchRec&))   'set hidden recno field on edit form
        toform.Wheretogo fromform, toform, 2 'send code 1 for search screen
        Load toform
        toform.Show
        DoEvents
        Unload frmInfo
      '  Unload frmCustEditLookUP
      End If
    ElseIf SrchDel = 2 Then
      If OKFinalCust(SearchRec&) Then
        DeActivateControls Me
        frmInfo.Label1 = "Loading. . ."
        frmInfo.Show
        DoEvents
      'here
        toform.fpCustRecNo = QPTrim$(Str$(SearchRec&))   'set hidden recno field on edit form
        toform.Wheretogo fromform, toform, 2 'send code 1 for search screen
        Load toform
        toform.Show
        DoEvents
        Unload frmInfo
      '  Unload frmCustEditLookUP
      End If
    ElseIf SrchDel = 3 Then
'      If OKApplyDep(SearchRec&) Then
'        DeActivateControls Me
'        frmInfo.Label1 = "Loading. . ."
'        frmInfo.Show
'        DoEvents
'      'here
'        toform.fpCustRecNo = QPTrim$(Str$(SearchRec&))   'set hidden recno field on edit form
'        toform.Wheretogo fromform, toform, 2 'send code 2 for list
'        Load toform
'        toform.Show
'        DoEvents
'        Unload frmInfo
'      '  Unload frmCustEditLookUP
'      End If
    ElseIf SrchDel = 4 Then
'      If OKDepCreditAdj(SearchRec&) Then
'        DeActivateControls Me
'        frmInfo.Label1 = "Loading. . ."
'        frmInfo.Show
'        DoEvents
'      'here
'        toform.fpCustRecNo = QPTrim$(Str$(SearchRec&))   'set hidden recno field on edit form
'        toform.Wheretogo fromform, toform, 2 'send code 2 for list
'        Load toform
'        toform.Show
'        DoEvents
'        Unload frmInfo
'      '  Unload frmCustEditLookUP
'      End If
    ElseIf SrchDel = 5 Then
'      If OKDepRefund(SearchRec&) Then
'        DeActivateControls Me
'        frmInfo.Label1 = "Loading. . ."
'        frmInfo.Show
'        DoEvents
'      'here
'        toform.fpCustRecNo = QPTrim$(Str$(SearchRec&))   'set hidden recno field on edit form
'        toform.Wheretogo fromform, toform, 2 'send code 2 for list
'        Load toform
'        toform.Show
'        DoEvents
'        Unload frmInfo
'      '  Unload frmCustEditLookUP
'      End If
    Else
      DeActivateControls Me
      frmInfo.Label1 = "Loading. . ."
      frmInfo.Show
      DoEvents
  ''here
      toform.fpCustRecNo = QPTrim$(Str$(SearchRec&))   'set hidden recno field on edit form
      toform.Wheretogo fromform, toform, 2
      Load toform
      toform.Show
      DoEvents
      Unload frmInfo
  '  '  Unload frmCustEditLookUP
    End If
  Else
    
    'frmCustEditLookUP.fpSearchText.SetFocus
    SearchRec = 0
    codeopt = 0
    MsgBox "Invalid Record.", vbOKOnly, "Invalid Record."
    ActivateControls frmCustEditLookUP
    Unload frmDisplayList

  End If
'  Else
'
'    Unload frmDisplayList
'  End If
End Sub
