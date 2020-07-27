VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmRateDelete 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Rate to Delete"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   1890
   ClientWidth     =   12210
   Icon            =   "frmRateDelete.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpList1 
      Height          =   3648
      Left            =   1560
      TabIndex        =   3
      Top             =   2544
      Width           =   9084
      _Version        =   196608
      _ExtentX        =   16023
      _ExtentY        =   6435
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
      Sorted          =   1
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
      ScrollBarH      =   3
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
      ColDesigner     =   "frmRateDelete.frx":08CA
   End
   Begin EditLib.fpLongInteger fpRateDelRec 
      Height          =   252
      Left            =   1488
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1104
      Visible         =   0   'False
      Width           =   684
      _Version        =   196608
      _ExtentX        =   1206
      _ExtentY        =   444
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   1
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   -1  'True
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   0
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "0"
      MaxValue        =   "2147483647"
      MinValue        =   "-2147483648"
      NegFormat       =   1
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   635
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
            TextSave        =   "10:26 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "6/28/2008"
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
   Begin fpBtnAtlLibCtl.fpBtn fpCmdOK 
      Height          =   480
      Left            =   7872
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6360
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
      ButtonDesigner  =   "frmRateDelete.frx":0BAE
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdExit 
      Height          =   480
      Left            =   9240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6360
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
      ButtonDesigner  =   "frmRateDelete.frx":0D87
   End
   Begin EditLib.fpLongInteger fpRateEntryFlag 
      Height          =   252
      Left            =   9984
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2184
      Visible         =   0   'False
      Width           =   684
      _Version        =   196608
      _ExtentX        =   1206
      _ExtentY        =   444
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   1
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   -1  'True
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   0
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "0"
      MaxValue        =   "2147483647"
      MinValue        =   "-2147483648"
      NegFormat       =   1
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
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
      Left            =   1824
      TabIndex        =   10
      Top             =   6456
      Width           =   5604
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "            Rate"
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
      Left            =   1584
      TabIndex        =   9
      Top             =   2184
      Width           =   1932
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description       "
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
      Left            =   4536
      TabIndex        =   8
      Top             =   2184
      Width           =   2148
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Base Amount"
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
      Left            =   8160
      TabIndex        =   7
      Top             =   2184
      Width           =   1380
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rate Table Delete"
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
      Left            =   3798
      TabIndex        =   2
      Top             =   1128
      Width           =   4668
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3222
      Top             =   888
      Width           =   5772
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3222
      Top             =   768
      Width           =   5772
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   4956
      Left            =   1338
      Top             =   2016
      Width           =   9540
   End
End
Attribute VB_Name = "frmRateDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BeenDone As Boolean
Dim WhatRate As Integer
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim RateRec As Integer
Dim Build As String * 80
Dim RateFile As Integer, RateRecNo As Integer
''Dim Changed As Boolean
Dim RateRecCnt As Integer, dcnt As Integer
Dim UBRateTblRecLen As Integer, cnt As Integer
Dim UBRateRec As UBRateTblRecType
'
'Private Sub fpCmdExit_Click()
'  fpRateEntryFlag.Value = False
'  BeenDone = False
'  RateRec = 0
'  Unload Me
'End Sub
'
'Private Sub Form_Activate()
'  If Not BeenDone Then
'    BeenDone = True
Private Sub LoadRatesList()
    RateRecCnt = GetNumRateRecs
    UBRateTblRecLen = Len(UBRateRec)
    RateFile = FreeFile
    Open UBPath + "UBRATE.DAT" For Random Shared As RateFile Len = UBRateTblRecLen
    For cnt = 1 To RateRecCnt
      Get RateFile, cnt, UBRateRec
      LSet Build$ = QPTrim$(UBRateRec.Ratecode)
      Mid$(Build$, 20) = QPTrim$(UBRateRec.RATEDESC)
      Mid$(Build$, 55) = Using("$######.##", UBRateRec.MINAMT)
      Mid$(Build$, 75) = Chr9$ + Str$(cnt)
      fpList1.AddItem Build$
      dcnt = dcnt + 1
    Next
    Close RateFile

End Sub

Private Sub fpCmdExit_Click()
  Call RateDeleteExit
End Sub

'    RateRecCnt = GetNumRateRecs
'    UBRateTblRecLen = Len(UBRateRec)
'    RateFile = FreeFile
'    Open UBPath + "UBRATE.DAT" For Random Shared As RateFile Len = UBRateTblRecLen
'    For cnt = 1 To RateRecCnt
'      Get RateFile, cnt, UBRateRec
'      LSet Build$ = QPTrim$(UBRateRec.RATECODE)
'      Mid$(Build$, 20) = QPTrim$(UBRateRec.RATEDESC)
'      Mid$(Build$, 55) = Using("$######.##", UBRateRec.MINAMT)
'      Mid$(Build$, 75) = Chr9$ + Str$(cnt)
'      frmRateDisplayList.fpList1.AddItem Build$
'      dcnt = dcnt + 1
'    Next
'    Close RateFile
'  End If
'
'  If tmpLastRate > 0 Then
'    Me.fpList1.ListIndex = tmpLastRate
'  Else
'    Me.fpList1.ListIndex = 0
'  End If
'
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'  Select Case KeyCode
'    Case vbKeyEscape:
'      KeyCode = 0
'      Call fpCmdExit_Click
'    Case vbKeyF10, vbKeyReturn
'      KeyCode = 0
'      Call fpCmdOk_Click
'    Case Else:
'  End Select
'End Sub
'
Private Sub fpCmdOk_Click()
  If fpList1.SelCount > 0 Then
    Call fpList1_DblClick
  End If
End Sub
'
Private Sub fpList1_DblClick()
  'Dim xx As Integer
  fpList1.col = 1                       'switch to the hidden RecNo. column
  RateRec = Val(fpList1.ColText) 'get customer recno
  If RateRec > 0 Then
    frmRateDelete.fpRateDelRec.Text = RateRec
    tmpLastRate = Me.fpList1.ListIndex
    Call CheckDeleteRate
    DoEvents

  Else
    frmRateDelete.fpRateDelRec = 0
    tmpLastRate = 0
    Call RateDeleteExit
  End If

 ' Call fpCmdExit_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      KeyCode = 0
      Call RateDeleteExit
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  LoadRatesList
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ' Me.Visible = False
    Temp_Class.ResizeControls Me
   ' Me.Visible = True
   ' Me.SetFocus
  End If
  DoEvents
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      UBLog "Closed via RateDelete by " + PWUser$
      CitiTerminate
    End If
  End If
End Sub

'Private Sub Form_Activate()
''  Stop
'  'If frmRateDisplayList.fpRateEntryFlag = True Then
'  frmRateDisplayList.fpRateEntryFlag = True
'  If Val(fpRateDelRec) = -1 And Not BeenDone Then
'    BeenDone = True
''    Load frmRateDisplayList
''    DoEvents
''    frmRateDisplayList.Show vbModal
'    If Val(fpRateDelRec) > 0 Then
'      Call CheckDeleteRate
'      DoEvents
''      Stop
'    Else
'      fpRateDelRec = 0
'      Call RateDeleteExit
'    End If
'    DoEvents
'  Else
'    BeenDone = True
'  End If
'End Sub
Private Sub CheckDeleteRate()
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer
  Dim UBFile As Integer, RateRecNo As Integer
  Dim UBFile1 As Integer, UBFile2 As Integer
  Dim UBRateTblRecLen As Integer, UBCustRecLen As Integer
  Dim cnt As Integer, SCnt As Integer
  Dim CustRate As String, DelRate As String
  Dim NoDelFlag As Boolean, BlankFlag As Boolean
  Dim NumOfCRecs As Long, CCnt As Long
  Dim NumOfRate As Integer
  Dim UBRateRec As UBRateTblRecType
  Dim UBCustRec As NewUBCustRecType
  
  UBCustRecLen = Len(UBCustRec)      'Length of Cust Record Structure
  UBRateTblRecLen = Len(UBRateRec)
  WhatRate = Val(fpRateDelRec)
  DeActivateControls Me
  
  FrmShowPctComp.Label1 = "Scanning Customer Information."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent

  If WhatRate > 0 Then
    UBFile = FreeFile      'open file and get the code for the selected rate
    Open UBPath + "UBRATE.DAT" For Random Shared As UBFile Len = UBRateTblRecLen
    Get UBFile, WhatRate, UBRateRec
    Close UBFile
  
    DelRate$ = QPTrim$(UBRateRec.Ratecode)
    If Len(DelRate$) = 0 Then      'if it was a blank rate table, set blank
      BlankFlag = True
      Unload FrmShowPctComp         'And jump over cust search
      GoTo BlankEntry
    End If
    If Exist(UBPath$ + "UBCUST.DAT") Then
      UBFile = FreeFile              'open cust file & prepare to search
      Open UBPath + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
      NumOfCRecs& = LOF(UBFile) \ UBCustRecLen
      For CCnt = 1 To NumOfCRecs&          'look at all customers
        Get #UBFile, CCnt, UBCustRec
        If Not UBCustRec.DelFlag Then   'not deleted ones
          For SCnt = 1 To 15               'look at all 15 possiable revenues
            CustRate$ = QPTrim$(UBCustRec.serv(SCnt).Ratecode)
            If Len(CustRate$) > 0 Then     'if there is a rate code
              If DelRate$ = CustRate$ Then 'if they are the same then they
                NoDelFlag = True           'can't delete this rate code
                Exit For                   'no need to continue looking
              End If                       'at this customer
            End If                         '
          Next                             '
          If NoDelFlag Then                'done with the search
            FrmShowPctComp.ShowPctComp 0, 0
            Exit For
          End If
        End If
        FrmShowPctComp.ShowPctComp CCnt, NumOfCRecs&
      Next
      Close UBFile
    Else
      'since no cust go ahead and delete
    End If
  Else
    GoTo ExitRateDelete
  End If

BlankEntry:
  If NoDelFlag Then                    'if they can't delete
    UBLog "ERROR: CAN'T DELETE RATE: " + DelRate$  'log it
    GoSub NODeleteErr                              'show em the error
  Else                                 'they can delete this rate
    MsgText(0) = "Warning!"
    MsgText(1) = "THIS IS THE LAST CHANCE!"
    MsgText(2) = ""
    MsgText(3) = "RATE CODE:  " + DelRate$
    MsgText(4) = "DELETE THIS RATE?"
    MsgText(5) = "ARE YOU SURE!"
    frmMsgDialog.Label(0).FontSize = frmMsgDialog.Label(0).FontSize + 5
    frmMsgDialog.Label(4).FontSize = frmMsgDialog.Label(4).FontSize + 5
    If GetOKorNot(MsgText(), False) Then
      GoSub DeleteRateRecord         'so go do it!
    Else
      'maybe show aborted
    End If
  End If '
  ActivateControls Me

ExitRateDelete:
  UBLog "OUT: DELETE Rate Code" + CrLf$           'log function exit
  Call RateDeleteExit                       'exit delete function
  
Exit Sub

NODeleteErr:           'Can't delete error display
    MsgText(0) = "ERROR"
    MsgText(1) = "ERROR ERROR ERROR!"
    MsgText(2) = ""
    MsgText(3) = "CAN NOT DELETE RATE CODE: " + DelRate$
    MsgText(4) = ""
    MsgText(5) = "THERE ARE CUSTOMERS USING THAT RATE"
    GetOKorNot MsgText(), True
    ActivateControls Me
    

Return


DeleteRateRecord:
  UBLog "DELETED RATE: " + DelRate$ + " REC:" + Str$(WhatRate) 'log deleted rate
  Call KillFileD(UBPath + "TRATETBL.TMP")              'make sure the temp files not there
  Name UBPath + "UBRATE.DAT" As UBPath + "TRATETBL.TMP" 'rename the rate file to the temp name
'
  UBFile1 = FreeFile                   'open old rate file
  Open UBPath + "TRATETBL.TMP" For Random Shared As UBFile1 Len = UBRateTblRecLen

  UBFile2 = FreeFile                   'open new rate file
  Open UBPath + "UBRATE.DAT" For Random Shared As UBFile2 Len = UBRateTblRecLen

  NumOfRate = LOF(UBFile1) \ UBRateTblRecLen
  For cnt = 1 To NumOfRate                'step thru the file
    If cnt <> WhatRate Then                  'if this isn't rate to delete
      Get UBFile1, cnt, UBRateRec         'get the rec from old file
      Put UBFile2, , UBRateRec           'then write this rec to new file
    End If                                '
  Next                                    'go till all are processed
  Close                                   'close up
  Call KillFileD(UBPath$ + "TRATETBL.TMP")                  'kill old rate file
  DoEvents
  frmDataUpdated.Show vbModal
Return

End Sub

Private Sub RateDeleteExit()
 'On Local Error Resume Next
  BeenDone = False
  WhatRate = 0
  Load frmUBRateMenu
  DoEvents
  frmUBRateMenu.Show
  DoEvents
  Unload frmRateDelete
End Sub
