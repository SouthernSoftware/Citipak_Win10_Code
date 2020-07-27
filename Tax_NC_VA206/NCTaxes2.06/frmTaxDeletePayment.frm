VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmTaxDeletePayment 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Payment Deletion"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTaxDeletePayment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpListPay 
      Height          =   3912
      Left            =   780
      TabIndex        =   0
      Top             =   2400
      Width           =   10092
      _Version        =   196608
      _ExtentX        =   17801
      _ExtentY        =   6900
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      Columns         =   5
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   1
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
      ColDesigner     =   "frmTaxDeletePayment.frx":08CA
   End
   Begin EditLib.fpText fptxtOperator 
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2895
      _Version        =   196608
      _ExtentX        =   5106
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
      ControlType     =   1
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
      Left            =   822
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7032
      Width           =   2052
      _Version        =   131072
      _ExtentX        =   3619
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
      ButtonDesigner  =   "frmTaxDeletePayment.frx":0D4E
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDelete 
      Height          =   540
      Left            =   6224
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7035
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
      ButtonDesigner  =   "frmTaxDeletePayment.frx":0F2C
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClearAll 
      Height          =   540
      Left            =   3360
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7035
      Width           =   2415
      _Version        =   131072
      _ExtentX        =   4260
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
      ButtonDesigner  =   "frmTaxDeletePayment.frx":110A
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSelectAll 
      Height          =   540
      Left            =   8760
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7035
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
      ButtonDesigner  =   "frmTaxDeletePayment.frx":12F1
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1486
      Top             =   675
      Width           =   8655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Payment List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3136
      TabIndex        =   3
      Top             =   825
      Width           =   5295
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   4455
      Left            =   593
      Top             =   2235
      Width           =   10455
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1486
      Top             =   555
      Width           =   8655
   End
End
Attribute VB_Name = "frmTaxDeletePayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim OpPayRecs() As Integer
  Dim DeleteThese() As Integer
  Dim DeleteAccts() As Long
  Dim dcnt As Integer

Private Sub cmdClearAll_Click()
  fpListPay.Action = ActionDeselectAll
End Sub

Private Sub cmdDelete_Click()
  'OpenTempPayFile is the same as open TaxCPRFileName
  'OpenPayListFile is the same as open TaxLOPFileName
  Dim PayRec As TaxPaymentRecType
  Dim PayRecNew() As TaxPaymentRecType
  Dim PayHandle As Integer
  Dim PayHandleNew As Integer
  Dim NumOfPRecs As Integer
  Dim NumOfPRecsNew As Integer
  Dim x As Integer, y As Integer
  Dim ListRec As PayListType
  Dim NumOfListRecs As Integer
  Dim ListRecNew() As PayListType
  Dim ListHandle As Integer
  Dim ListHandleNew As Integer
  Dim NewCnt As Integer
  Dim Operator$
  Dim Oper$
  Dim ThisLastRec As Long
  Dim MatchCnt As Integer
  
  'on error goto ERRORSTUFF
  
  Operator = CStr(OperNum)
  dcnt = 0
  ReDim DeleteThese(1 To 1) As Integer
  ReDim DeleteAccts(1 To 1) As Long
  fpListPay.Col = 0
  For x = 0 To fpListPay.ListCount - 1 'changed from 1 to fpListPay.ListCount on 6/29/06
'  For x = 1 To fpListPay.ListCount
    fpListPay.Row = x
      fpListPay.Col = 0
      fpListPay.ListIndex = fpListPay.Row
    If fpListPay.Selected = True Then
      dcnt = dcnt + 1
      ReDim Preserve DeleteThese(1 To dcnt) As Integer
      ReDim Preserve DeleteAccts(1 To dcnt) As Long
      fpListPay.Col = 4
      DeleteThese(dcnt) = CInt(fpListPay.ColText) 'x
      fpListPay.Col = 0
      fpListPay.ListIndex = fpListPay.Row
      DeleteAccts(dcnt) = CLng(fpListPay.ColText)
    End If
  Next x
  
  If dcnt = 0 Then
    frmTaxMsg.Label1.Caption = "No payments have been selected. Deletion attempt aborted."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Close
    Exit Sub
  End If
  
  OpenPayListFile ListHandle, OperNum
  NumOfListRecs = LOF(ListHandle) / Len(ListRec)
  If NumOfListRecs = 0 Then
    frmTaxMsg.Label1.Caption = "No bills have been tagged for payment. Delete attempt aborted."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Close
    Exit Sub
  End If
  
  If TaxMsgWOpts(900, "Are you sure you want to delete this transaction? Press F10 to delete. Otherwise, press ESC to abort.", "F10 Delete", "ESC Abort") = "abort" Then
    Close
    Exit Sub
  End If
  
  NewCnt = 0
  For x = 1 To NumOfListRecs
    Get ListHandle, x, ListRec
    ListRec.PrevListRec = ListRec.PrevListRec
    If ListRec.PrevListRec = -1 Then GoTo SkipIt
    For y = 1 To dcnt
      If ListRec.CustRec = DeleteAccts(y) Then
        ListRec.PrevListRec = -1
        Put ListHandle, x, ListRec
      End If
    Next y
SkipIt:
  Next x
  Close ListHandle
  
  OpenTempPayFile PayHandle, OperNum
  NumOfPRecs = LOF(PayHandle) / Len(PayRec)
  ReDim PayRecNew(1 To 1) As TaxPaymentRecType
  NewCnt = 0
  For x = 1 To NumOfPRecs
    Get PayHandle, x, PayRec
    For y = 1 To dcnt
      If DeleteThese(y) = x Then
        PayRec.LastPayRec = 0
        Put PayHandle, x, PayRec
        'added Mainlog on 6/29/06
        MainLog ("Payment for acct # " + CStr(PayRec.CustAcct) + " for " + QPTrim$(Using$("$###,##0.00", PayRec.TotPaid)) + " was deleted successfully.")
      End If
    Next y
  Next x
  Close PayHandle
  Close 'added 7/17/07
  
  If NumOfPRecs = 1 Then
    KillFile "TAXLOP" + Operator$ + ".DAT"
    KillFile "TAXCPR" + Operator$ + ".DAT"
  Else
    OpenTempPayFile PayHandle, OperNum
    NumOfPRecs = LOF(PayHandle) / Len(PayRec)
    For x = 1 To NumOfPRecs
      Get PayHandle, x, PayRec
      If PayRec.LastPayRec = 0 Then
        MatchCnt = MatchCnt + 1
      End If
    Next x
    Close
    If MatchCnt = NumOfPRecs Then
      KillFile "TAXLOP" + Operator$ + ".DAT"
      KillFile "TAXCPR" + Operator$ + ".DAT"
    End If
  End If
  
NoMore2:
  If dcnt > 1 Then
    frmTaxMsg.Label1.Caption = CStr(dcnt) + " payments have been deleted successfully."
  Else
    frmTaxMsg.Label1.Caption = CStr(dcnt) + " payment has been deleted successfully."
  End If
  
  frmTaxMsg.Label1.Top = 900
  frmTaxMsg.Show vbModal
  
  fpListPay.Action = ActionClear
  Call ReloadMe
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxDeletePayment", "cmdDelete_Click", Erl)
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

Private Sub cmdExit_Click()
  Unload Me
  DoEvents
  frmTaxPayMenu.Show
End Sub

Private Sub cmdSelectAll_Click()
  fpListPay.Action = ActionSelectAll
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
    Case vbKeyF3:
      SendKeys "%C"
      Call cmdClearAll_Click
      KeyCode = 0
    Case vbKeyF7:
      SendKeys "%S"
      Call cmdSelectAll_Click
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
  Me.HelpContextID = hlpDeleteTax
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxDeletePayment.")
      Call Terminate
      End
    End If
  End If

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub LoadMe()
  'OpenTempPayFile is the same as open TaxCPRFileName
  'OpenPayListFile is the same as open TaxLOPFileName
'  OpenPayListFile ListHandle, OPERNUM
  Dim PayRec As TaxPaymentRecType
  Dim PayHandle As Integer
  Dim NumOfPRecs As Integer
  Dim x As Integer
  Dim ListRec As PayListType
  Dim ListHandle As Integer
  Dim OHandle As Integer
  Dim NumOfOperRecs As Integer
  Dim OperRec As CitiPassType
  Dim NumOfOperPayRecs As Integer
  
'  OPERNUM = 1
  fptxtOperator.Text = "Operator # " + CStr(OperNum)
  OpenTempPayFile PayHandle, OperNum
  NumOfPRecs = LOF(PayHandle) / Len(PayRec)
  If NumOfPRecs = 0 Then
    frmTaxMsg.Label1.Caption = "There are no payment records saved for operator #" + CStr(OperNum) + "."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Close
    Exit Sub
  End If
  
  For x = 1 To NumOfPRecs
  Get PayHandle, x, PayRec
    If PayRec.LastPayRec <> 0 Then
      fpListPay.InsertRow = CStr(PayRec.CustAcct) + Chr(9) + QPTrim$(PayRec.CustName) + Chr(9) + QPTrim$(Using$("$###,##0.00", PayRec.AmtPaid)) + Chr(9) + QPTrim$(Using$("$###,##0.00", PayRec.AmtOwed)) + Chr(9) + CStr(x)
      DoEvents
    End If
  Next x
  Close PayHandle
  
'  fpListPay.ListIndex = -1
  fpListPay.Action = ActionDeselectAll
  
End Sub

Private Sub ReloadMe()
  Dim PayRec As TaxPaymentRecType
  Dim PayHandle As Integer
  Dim NumOfPRecs As Integer
  Dim x As Integer
  
'  OPERNUM = 1
  fptxtOperator.Text = "Operator # " + CStr(OperNum)
  If Exist("TAXCPR" + CStr(OperNum) + ".DAT") Then
    OpenTempPayFile PayHandle, OperNum
    NumOfPRecs = LOF(PayHandle) / Len(PayRec)
    For x = 1 To NumOfPRecs
    Get PayHandle, x, PayRec
      If PayRec.LastPayRec <> 0 Then
        fpListPay.InsertRow = CStr(PayRec.CustAcct) + Chr(9) + QPTrim$(PayRec.CustName) + Chr(9) + QPTrim$(Using$("$###,##0.00", PayRec.AmtPaid)) + Chr(9) + QPTrim$(Using$("$###,##0.00", PayRec.AmtOwed)) + Chr(9) + CStr(x)
      End If
    Next x
    Close PayHandle
  
    fpListPay.Refresh
    fpListPay.ListIndex = -1
  End If
End Sub

