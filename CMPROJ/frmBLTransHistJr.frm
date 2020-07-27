VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLTransHistJr 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Transaction History"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmBLTransHistJr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpList 
      Height          =   3708
      Left            =   1836
      TabIndex        =   6
      Tag             =   $"frmBLTransHistJr.frx":08CA
      Top             =   2616
      Width           =   8064
      _Version        =   196608
      _ExtentX        =   14224
      _ExtentY        =   6540
      TextAlias       =   ""
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Columns         =   5
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
      ColDesigner     =   "frmBLTransHistJr.frx":09CF
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   540
      Left            =   2112
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   $"frmBLTransHistJr.frx":0D5D
      Top             =   7212
      Width           =   2100
      _Version        =   131072
      _ExtentX        =   3704
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
      ButtonDesigner  =   "frmBLTransHistJr.frx":0E2D
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   348
      Left            =   1200
      TabIndex        =   8
      Top             =   7116
      Width           =   540
      _Version        =   131072
      _ExtentX        =   952
      _ExtentY        =   614
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   9405029
      ForeColor       =   8454143
      Text            =   ""
      Shape           =   0
      ShapeRoundWidth =   195
      ShapeRoundHeight=   195
      BorderWidth     =   -1
      BorderColor     =   -2147483630
      ThreeDWidth     =   -1
      ThreeDShadowColor=   -2147483632
      ThreeDHighlightColor=   16777215
      ThreeDText      =   0
      ThreeDTextHighlightColor=   16777215
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignPictureH   =   0
      AlignPictureV   =   0
      PictureStyle    =   0
      WordWrap        =   -1  'True
      ScaleMode       =   1
      ThreeDStyle     =   2
      Position        =   0
      PosBaseX        =   0
      PosBaseY        =   0
      PosOffsetX      =   -100
      PosOffsetY      =   300
      MaxWidth        =   6000
      CloudInset      =   100
      CloudMinWidth   =   600
      TailShape       =   2
      TailType        =   2
      TailBaseOffsetOutside=   300
      TailBaseOffsetInside=   100
      TailBaseAxisOutside=   0
      TailBaseAxisInside=   0
      TailBubbleCount =   3
      AlignTextH      =   1
      AlignTextV      =   1
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      AutoScan        =   0
      ScanAllDescendants=   -1  'True
      Interval        =   500
      IntervalNext    =   200
      AutoSize        =   -1  'True
      UseTagProp      =   -1  'True
      HideOnInactiveApp=   0   'False
      HideOnMouseDown =   2
      HideOnKeyDown   =   2
      HideOnFocus     =   0   'False
      ScanDisabledControls=   -1  'True
      ThreeDAppearance=   0
      FollowFocus     =   0   'False
      TemplateName    =   ""
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   540
      Left            =   7584
      TabIndex        =   10
      TabStop         =   0   'False
      Tag             =   $"frmBLTransHistJr.frx":1010
      Top             =   7212
      Width           =   1956
      _Version        =   131072
      _ExtentX        =   3450
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
      ButtonDesigner  =   "frmBLTransHistJr.frx":1098
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDetail 
      Height          =   540
      Left            =   4890
      TabIndex        =   11
      TabStop         =   0   'False
      Tag             =   $"frmBLTransHistJr.frx":1276
      Top             =   7215
      Width           =   1965
      _Version        =   131072
      _ExtentX        =   3466
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
      ButtonDesigner  =   "frmBLTransHistJr.frx":1323
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080FFFF&
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   1212
      Left            =   2988
      Top             =   336
      Width           =   5676
   End
   Begin VB.Label lblBalloon 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "HELP BALLOONS ON"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   2112
      TabIndex        =   9
      Top             =   7788
      Width           =   2100
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   5004
      Left            =   1188
      Top             =   2028
      Width           =   9276
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Balance"
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
      Height          =   252
      Left            =   8676
      TabIndex        =   5
      Top             =   2220
      Width           =   1260
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Trans Amt"
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
      Height          =   252
      Left            =   6708
      TabIndex        =   4
      Top             =   2220
      Width           =   1596
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Height          =   252
      Left            =   4224
      TabIndex        =   3
      Top             =   2220
      Width           =   1596
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Trans Date"
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
      Height          =   252
      Left            =   1860
      TabIndex        =   2
      Top             =   2220
      Width           =   1356
   End
   Begin VB.Label lblBal 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   348
      Left            =   4188
      TabIndex        =   1
      Tag             =   "This field shows the current outstanding balance for this customer."
      Top             =   1020
      Width           =   3276
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   348
      Left            =   3372
      TabIndex        =   0
      Top             =   492
      UseMnemonic     =   0   'False
      Width           =   4908
   End
End
Attribute VB_Name = "frmBLTransHistJr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Public ThisRec$
  Public ThisName$
  Private Temp_Class As Resize_Class

Private Sub cmdDetail_Click()

  If fpList.ListIndex = -1 Then
    frmBLMessageBoxJr.Label1.Caption = "Please make a selection from the list."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  fpList.col = 4
  fpList.Row = fpList.ListIndex
  ThisRec = fpList.ColText
  
  frmBLTransDetail.Show vbModal
End Sub

Private Sub cmdExit_Click()
  KillFile "transhistjr.dat"
  If Exist("adjustbalance.dat") Then
    frmBLAdjustBal.Show
'  ElseIf Exist("customeredit.dat") Then
'    frmBLCustEdit.Show
'  Else
'    frmBLCustInfoTrans.Show
  End If
  DoEvents
  Unload Me
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    lblBalloon.Visible = True
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
  End If
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      ClearInUse PWcnt
      BLLog ("terminated via menu bar on frmBLTransHistJr.")
      CMLog ("terminated via menu bar on frmBLTransHistJr.")
      CitiTerminate
      End
    End If
  End If

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ''' Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      Call cmdExit_Click
      SendKeys "%C"
      KeyCode = 0
    Case vbKeyF4:
      Call cmdDetail_Click
      SendKeys "%D"
      KeyCode = 0
    Case vbKeyF1:
      Call cmdHelp_Click
      SendKeys "%H"
      KeyCode = 0
    Case Else:
  End Select
  
End Sub
Public Sub LoadMe()
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim x As Integer
  Dim TransRec As ARTransRecType
  Dim TransHandle As Integer
  Dim NumOfTransRecs As Double
  Dim TransCnt As Integer
  Dim TransRecd&
  Dim TransDesc$
  Dim One As Integer
  Dim DHandle As Integer
  Dim RecNum As Integer
  Dim UpDown$
  
  RecNum = ThisCustXNum
  
  lblBalloon.Visible = False
  
  OpenBLCustFile CustHandle
  Get CustHandle, RecNum, CustRec
  Close CustHandle
  
  ThisName$ = QPTrim$(CustRec.BILLNAME)
  
  lblName.Caption = QPTrim$(CustRec.CustName)
  
  If CustRec.AcctBal < 0 Then
    lblBal.ForeColor = &H8000&
  Else
    lblBal.ForeColor = &HFF&
  End If
  
  lblBal.Caption = "CURRENT BALANCE: " + QPTrim$(Using("$#,###,##0.00", CustRec.AcctBal))
  
  OpenBLTransFile TransHandle
  NumOfTransRecs = LOF(TransHandle) / Len(TransRec)
  
  TransRecd& = CustRec.FirstTrans
  TransCnt = 0
  ReDim TransArray(1 To 1) As Long
  Do While TransRecd& > 0
    Get TransHandle, TransRecd&, TransRec
      TransCnt = TransCnt + 1
      ReDim Preserve TransArray(1 To TransCnt) As Long
      TransArray(TransCnt) = TransRecd&
      TransRecd& = TransRec.NextTrans
  Loop
  
  TransDesc = ""
  For x = TransCnt To 1 Step -1
    Get TransHandle, TransArray(x), TransRec
      If TransRec.DetailTransType > 0 Then
        Select Case TransRec.DetailTransType
          Case 101
            TransDesc = "Penalty Charge"
            UpDown = "Up"
          Case 110
            TransDesc = "License Charge"
            UpDown = "Up"
          Case 201
            TransDesc = "Penalty Payment"
            UpDown = "Down"
          Case 210
            TransDesc = "License Payment"
            UpDown = "Down"
          Case 211
            TransDesc = "License & Penalty Payment"
            UpDown = "Down"
          Case 301
            TransDesc = "Penalty Adjustment Down"
            UpDown = "Down"
          Case 310
            TransDesc = "License Adjustment Down"
            UpDown = "Down"
          Case 311
            TransDesc = "License & Penalty Adjustment Down"
            UpDown = "Down"
          Case 401
           If TransRec.TransType = 13 Then
              TransDesc = "Down Pay Adjustment"
            Else
              TransDesc = "Penalty Adjustment Up"
            End If
            UpDown = "Up"
          Case 410
            If TransRec.TransType = 13 Then
              TransDesc = "Down Pay Adjustment"
            Else
              TransDesc = "License Adjustment Up"
            End If
            UpDown = "Up"
          Case 411
            If TransRec.TransType = 13 Then
              TransDesc = "Down Pay Adjustment"
            Else
              TransDesc = "License & Penalty Adjustment Up"
            End If
            UpDown = "Up"
          Case Else
            TransDesc = "Unknown"
        End Select
      Else
        Select Case TransRec.TransType
          Case 1
            TransDesc = "License Charge"
            UpDown = "Up"
          Case 2
            TransDesc = "Payment"
            UpDown = "Down"
          Case 6
            TransDesc = "Penalty Charge"
            UpDown = "Up"
          Case 13
            TransDesc = "Payment Adjustment Down"
            UpDown = "Down"
          Case 23
            TransDesc = "Billing Adjustment Down"
            UpDown = "Down"
          Case 24
            TransDesc = "Billing Adjustment Up"
            UpDown = "Up"
          Case Else
            TransDesc = "Unknown"
        End Select
      End If
      If UpDown = "Down" Then
        fpList.AddItem "  " + Num2Date(TransRec.TransDate) + Chr(9) + TransDesc + Chr(9) + Using("$#,###,##0.00", -TransRec.TransAmount) + Chr(9) + Using("$#,###,##0.00", TransRec.BalanceAfterTrans) + Chr(9) + CStr(TransArray(x))
      Else
        fpList.AddItem "  " + Num2Date(TransRec.TransDate) + Chr(9) + TransDesc + Chr(9) + Using("$#,###,##0.00", TransRec.TransAmount) + Chr(9) + Using("$#,###,##0.00", TransRec.BalanceAfterTrans) + Chr(9) + CStr(TransArray(x))
      End If
  Next x
  
  fpList.ListIndex = 0
  
  Close TransHandle
  Close CustHandle
  
  One = 1
  DHandle = FreeFile
  Open "transhistjr.dat" For Output As DHandle Len = 2
  Print #DHandle, One
  Close DHandle
  
  Call FixFonts
  
End Sub

Private Sub fpList_DblClick()
  Call cmdDetail_Click
End Sub

Private Sub FixFonts()
  Dim x As Integer
  
  On Error Resume Next
  Select Case ScreenW
    Case 1280
      fpList.col = 0
      fpList.ColWidth = 12
      fpList.col = 1
      fpList.ColWidth = 29
      fpList.col = 2
      fpList.ColWidth = 19
      fpList.col = 3
      fpList.ColWidth = 18
    Case 800
    Case Else
  End Select

End Sub

