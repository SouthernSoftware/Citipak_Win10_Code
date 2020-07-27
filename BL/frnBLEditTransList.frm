VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmBLEditTransList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Transaction Listing"
   ClientHeight    =   9216
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11724
   Icon            =   "frnBLEditTransList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9216
   ScaleMode       =   0  'User
   ScaleWidth      =   11724
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpList1 
      Height          =   3912
      Left            =   936
      TabIndex        =   0
      Tag             =   $"frnBLEditTransList.frx":08CA
      Top             =   2472
      Width           =   9852
      _Version        =   196608
      _ExtentX        =   17378
      _ExtentY        =   6900
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.4
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
      Columns         =   3
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   0
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   0
      WrapList        =   0   'False
      WrapWidth       =   0
      SelMax          =   -1
      AutoSearch      =   2
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
      ColDesigner     =   "frnBLEditTransList.frx":0A6C
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClose 
      Height          =   492
      Left            =   4800
      TabIndex        =   1
      Tag             =   "Press 'ESC Cancel' to exit this screen and return to the 'Payment Entry' menu."
      Top             =   7050
      Width           =   2124
      _Version        =   131072
      _ExtentX        =   3746
      _ExtentY        =   868
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
      ButtonDesigner  =   "frnBLEditTransList.frx":0DAF
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   492
      Left            =   7247
      TabIndex        =   3
      Tag             =   $"frnBLEditTransList.frx":0F8D
      Top             =   7050
      Width           =   2124
      _Version        =   131072
      _ExtentX        =   3746
      _ExtentY        =   868
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
      ButtonDesigner  =   "frnBLEditTransList.frx":101F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   492
      Left            =   2354
      TabIndex        =   4
      Tag             =   $"frnBLEditTransList.frx":11FE
      Top             =   7050
      Width           =   2124
      _Version        =   131072
      _ExtentX        =   3746
      _ExtentY        =   868
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
      ButtonDesigner  =   "frnBLEditTransList.frx":128F
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   444
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   780
      _Version        =   131072
      _ExtentX        =   1376
      _ExtentY        =   783
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
      ShapeRoundWidth =   180
      ShapeRoundHeight=   180
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
      MaxWidth        =   3000
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
   Begin VB.Label lblBalloon 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "HELP BALLOONS ON"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   2354
      TabIndex        =   6
      Top             =   7580
      Width           =   2100
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   6684
      Left            =   552
      Top             =   1248
      Width           =   10620
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction List"
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
      Height          =   444
      Left            =   3912
      TabIndex        =   2
      Top             =   1734
      Width           =   3900
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   636
      Left            =   3840
      Top             =   1578
      Width           =   4044
   End
End
Attribute VB_Name = "frmBLEditTransList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class
  Dim TCnt As Integer

Private Sub cmdClose_Click()
  frmBLEnterPayments.Show
  DoEvents
  Unload frmBLEditTransList2
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "F1 Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    lblBalloon.Visible = True
    cmdHelp.ToolTipText = ""
    cmdClose.ToolTipText = ""
    cmdProcess.ToolTipText = ""
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
  End If
End Sub

Private Sub cmdProcess_Click()
  Call fpList1_DblClick
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call LoadMe
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyReturn
      Call fpList1_DblClick
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      Call cmdClose_Click
      SendKeys "%C"
      KeyCode = 0
    Case vbKeyF10:
      Call cmdProcess_Click
      SendKeys "%P"
      KeyCode = 0
    Case vbKeyF1:
      Call cmdHelp_Click
      SendKeys "%T"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub LoadMe()
  Dim PayRec As AREditPaymentRecType
  Dim PayHandle As Integer
  Dim NextRec As Integer
  Dim NumOfTransRecs As Integer
  Dim x As Integer
  'this list loads all operator's unposted transactions...
  'however, only those listed under the current operator can
  'be edited by the current operator
  On Error GoTo ERRORSTUFF

  lblBalloon.Visible = False
'  cmdHelp.ToolTipText = "Press the 'F1 Turn Help On' button to activate the informational balloons. Press 'F1 Turn Help Off' to deactivate the informational balloons."
'  cmdClose.ToolTipText = "Press to exit this screen."
'  cmdProcess.ToolTipText = "After highlighting an entry in the list above press this button to bring up the transaction edit screen for that transaction."

  TCnt = 0

  'first open this operator's file and get his transactions
  OpenPayFile PayHandle, OPERNUM
  NumOfTransRecs = LOF(PayHandle) / Len(PayRec)
  If NumOfTransRecs = 0 Then
    'OK...there are no transactions saved for this operator
    'so we write out a line stating that fact and then
    'we go on to look for other operator's transactions
    fpList1.InsertRow = ""
    fpList1.InsertRow = "" + Chr$(9) + "NO TRANSACTION RECORDS ON FILE FOR OPERATOR # " + CStr(OPERNUM)
    TCnt = TCnt + 2
    Close PayHandle
    GoTo CheckOtherOps 'looks at other operator's transactions
  End If
  For x = 1 To NumOfTransRecs 'find 'em and print 'em
    Get PayHandle, x, PayRec
    If QPTrim$(PayRec.CustNumber) = "" Then GoTo SkipDel
    TCnt = TCnt + 1
    fpList1.InsertRow = QPTrim$(PayRec.CustNumber) & " " & Chr$(9) & QPTrim$(PayRec.CustName) & Chr$(9) & Str(x)
SkipDel:
  Next x

CheckOtherOps:
  Close PayHandle
  Call OtherEditChk
  fpList1.Row = -1

  Exit Sub

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLEditTransList2", "LoadMe", Erl)
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
    ClearInUse PWcnt
    Terminate


End Sub

Private Sub fpList1_DblClick()
  Dim PayRec As AREditPaymentRecType
  Dim PayHandle As Integer
  Dim NextRec As Integer
  Dim NumOfTransRecs As Integer
  Dim x As Integer
  Dim Number As Integer
  Dim Found As Boolean
  Dim Nextx As Integer
  Dim CustNum$
  Dim ThisCol$

  On Error GoTo ERRORSTUFF
  fpList1.Row = -1

  If fpList1.ListIndex <= TCnt Then
    fpList1.Col = 0
    If fpList1.ColText = "" Then
      frmBLMessageBoxJr.Label1.Caption = "Please select a row with a customer number listed."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Exit Sub
    End If
  End If

  If fpList1.ListIndex > TCnt Then
    frmBLMessageBoxJr.Label1.Caption = "Unable to bring up transaction data for another operator."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If

  fpList1.Col = 2 'assign variables from the user selected row
  If QPTrim$(fpList1.ColText) <> "" Then
    Number = CInt(fpList1.ColText)
  Else
    frmBLMessageBoxJr.Label1.Caption = "Please select a row with a customer number listed."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If

  OpenPayFile PayHandle, OPERNUM
  NumOfTransRecs = LOF(PayHandle) / Len(PayRec)

  If NumOfTransRecs = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "Error: No transaction records on file for current operator."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If

  Nextx = 1
  For x = 1 To NumOfTransRecs
  Get PayHandle, x, PayRec
    If Number = Nextx Then
      If QPTrim$(PayRec.CustName) <> "" Then
        'row with the right code
        Found = True
        GPayNum = Nextx 'now you can assign the correct global
        CustNum$ = QPTrim$(PayRec.CustNumber)
        Exit For
      End If
    Else
      Found = False
      GoTo NotAMatch
    End If

NotAMatch:
  Nextx = Nextx + 1
  Next x
  Close PayHandle

  If Found = True Then
    EditFlag = True
    Close
    frmBLTransEntry.Show
    DoEvents
    Unload frmBLEditTransList2
  Else
    frmBLMessageBoxJr.Label1.Caption = "No match found for selected transaction. Please make sure a transaction is highlighted and try again."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If

  Exit Sub

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLEditTransList2", "fpList1_DblClick", Erl)
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
    ClearInUse PWcnt
    Terminate


End Sub
Private Sub OtherEditChk()
  Dim x As Integer
  Dim OHandle As Integer
  Dim OperRec As CitiPassType 'CMOperRecType
  Dim NumOperRecs As Integer
  Dim Operator$
  Dim Y As Integer
  Dim PayHandle As Integer
  Dim EditPayRec As AREditPaymentRecType
  Dim NumOfPayRecs As Integer
  Dim OCnt As Integer

  On Error GoTo ERRORSTUFF

  'go to the password file and get the
  'operator numbers
  OpenCitiPassFile OHandle, NumOperRecs
  If NumOperRecs = 0 Then
    Close
    Exit Sub
  End If

  ReDim OpIdx(1 To NumOperRecs) As Integer
  For x = 1 To NumOperRecs
    Get OHandle, x, OperRec
      'load an array with the operator numbers
      OpIdx(x) = OperRec.PassNum
  Next x
  Close OHandle
  OCnt = 0
  For x = 1 To NumOperRecs
    Operator = Str(OpIdx(x))
    'if the operator number for the current operator
    'comes up then skip over it
    If Operator = Str(OPERNUM) Then GoTo SkipIt
    If Exist(BLPayFileName + Operator$ + ".DAT") Then
      'if the file above exists then this operator has
      'saved at least one transaction
      fpList1.InsertRow = " " + Chr(9) + "Operator #" + Str(OpIdx(x))
      OpenPayFile PayHandle, OpIdx(x) 'look thru all operator files
      NumOfPayRecs = LOF(PayHandle) / Len(EditPayRec)
      For Y = 1 To NumOfPayRecs 'if you find this customer already
      'has
        Get PayHandle, Y, EditPayRec
        If QPTrim$(EditPayRec.CustNumber) = "" Then GoTo Deleted
        OCnt = OCnt + 1
        fpList1.InsertRow = QPTrim$(EditPayRec.CustNumber) + Chr(9) + QPTrim$(EditPayRec.CustName) + Chr(9) + Str(Y)
Deleted:
      Next Y
      If OCnt = 0 Then
        fpList1.InsertRow = " " + Chr(9) + "All transactions have been deleted."
      Else
        OCnt = 0
      End If

    End If
  fpList1.InsertRow = " "
SkipIt:
  Next x

  Close

  Exit Sub

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLEditTransList2", "PrintText", Erl)
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
    ClearInUse PWcnt
    Terminate


End Sub

