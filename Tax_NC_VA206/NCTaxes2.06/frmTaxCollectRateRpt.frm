VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmTaxCollectRateRpt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Collection Rate Report"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTaxCollectRateRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6270
      Left            =   1920
      TabIndex        =   2
      Top             =   1230
      Width           =   7785
      _Version        =   196609
      _ExtentX        =   13732
      _ExtentY        =   11060
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmTaxCollectRateRpt.frx":08CA
      Begin LpLib.fpList fpList 
         Height          =   1776
         Left            =   3240
         TabIndex        =   0
         Top             =   1920
         Width           =   1572
         _Version        =   196608
         _ExtentX        =   2773
         _ExtentY        =   3133
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
         Object.TabStop         =   0   'False
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Columns         =   0
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
         ColDesigner     =   "frmTaxCollectRateRpt.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   384
         Left            =   2928
         TabIndex        =   1
         Top             =   4368
         Width           =   3564
         _Version        =   196608
         _ExtentX        =   6286
         _ExtentY        =   677
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
         Object.TabStop         =   -1  'True
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
         AutoSearchFillDelay=   200
         EditMarginLeft  =   1
         EditMarginTop   =   1
         EditMarginRight =   0
         EditMarginBottom=   3
         ResizeRowToFont =   0   'False
         TextTipMultiLine=   0
         AutoMenu        =   -1  'True
         EditAlignH      =   1
         EditAlignV      =   0
         ColDesigner     =   "frmTaxCollectRateRpt.frx":0C1E
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   645
         Left            =   870
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   5250
         Width           =   1620
         _Version        =   131072
         _ExtentX        =   2857
         _ExtentY        =   1138
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
         ButtonDesigner  =   "frmTaxCollectRateRpt.frx":0FC1
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   645
         Left            =   5235
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   $"frmTaxCollectRateRpt.frx":119F
         Top             =   5250
         Width           =   1740
         _Version        =   131072
         _ExtentX        =   3069
         _ExtentY        =   1138
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
         ButtonDesigner  =   "frmTaxCollectRateRpt.frx":124A
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdClear 
         Height          =   645
         Left            =   2760
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   5250
         Width           =   2220
         _Version        =   131072
         _ExtentX        =   3916
         _ExtentY        =   1138
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
         ButtonDesigner  =   "frmTaxCollectRateRpt.frx":1429
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Year:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3480
         TabIndex        =   7
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   3660
         Left            =   888
         Top             =   1368
         Width           =   5976
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Collection Rate Report"
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
         Height          =   396
         Left            =   1680
         TabIndex        =   6
         Top             =   456
         Width           =   4332
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   696
         Left            =   1416
         Top             =   312
         Width           =   4908
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Report Type:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1275
         TabIndex        =   5
         Top             =   4485
         Width           =   1500
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   6540
      Left            =   1800
      Top             =   1095
      Width           =   8055
   End
End
Attribute VB_Name = "frmTaxCollectRateRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim AllFlag As Boolean
  Dim AllYears() As Integer

Private Sub cmdClear_Click()
  fpList.Action = ActionDeselectAll
End Sub

Private Sub cmdExit_Click()
  frmTaxReportsMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdProcess_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Integer
  Dim NextRec As Long
  Dim ThisYear As Integer
  Dim YrCnt As Integer
  Dim ThisPaid As Double
  Dim ThisChrg As Double
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim PctPaid As Double
  
'  'on error goto ERRORSTUFF
  
  If fpList.ListIndex = -1 Then
    Call TaxMsg(900, "Please select a year or years from the list.")
    Exit Sub
  End If
  
  fpList.Row = 0
  If fpList.Selected = True Then
    AllFlag = True
    YrCnt = fpList.ListCount - 1
    ReDim TaxYear(1 To YrCnt) As Integer
    For x = 1 To YrCnt
'      fpList.Row = x
'      fpList.ListIndex = x
      TaxYear(x) = AllYears(x) ' CInt(fpList.Text)
    Next x
  Else
    AllFlag = False
    ReDim TaxYear(1 To 1) As Integer
    For x = 1 To fpList.ListCount - 1
      fpList.Row = x
      If fpList.Selected = True Then
        fpList.ListIndex = x
        YrCnt = YrCnt + 1
        ReDim Preserve TaxYear(1 To YrCnt) As Integer
        TaxYear(YrCnt) = CInt(fpList.Text)
      End If
    Next x
  End If
  
  ReDim YrTaxChrgs(1 To YrCnt) As Double
  ReDim YrTaxPaid(1 To YrCnt) As Double
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  frmTaxShowPctComp.Label1 = "Creating Collection Rate Report"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdClear.Enabled = False
  cmdProcess.Enabled = False
  
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    For y = 1 To YrCnt
      If TaxTrans.TaxYear = TaxYear(y) Then
        Exit For
      End If
    Next y
    If y > YrCnt Then GoTo GetNext
    If TaxTrans.TranType <> 1 Then GoTo GetNext
    If TaxTrans.CustomerRec < 0 Then GoTo GetNext
    Get TCHandle, TaxTrans.CustomerRec, TaxCust
    If TaxCust.Deleted <> 0 Then GoTo GetNext
    ThisPaid = 0
    ThisChrg = 0
    
    ThisChrg = OldRound(ThisChrg + TaxTrans.Revenue.Collection + TaxTrans.Revenue.Interest)
    ThisChrg = OldRound(ThisChrg + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Penalty)
    ThisChrg = OldRound(ThisChrg + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2)
    ThisChrg = OldRound(ThisChrg + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4)
    ThisChrg = OldRound(ThisChrg + TaxTrans.Revenue.Principle5 + TaxTrans.Revenue.RevOpt1)
    ThisChrg = OldRound(ThisChrg + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
    YrTaxChrgs(y) = OldRound(YrTaxChrgs(y) + ThisChrg)
    
    ThisPaid = OldRound(TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.InterestPd)
    ThisPaid = OldRound(ThisPaid + TaxTrans.Revenue.LateListPd + TaxTrans.Revenue.PenaltyPd)
    ThisPaid = OldRound(ThisPaid + TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd)
    ThisPaid = OldRound(ThisPaid + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd)
    ThisPaid = OldRound(ThisPaid + TaxTrans.Revenue.Principle5Pd + TaxTrans.Revenue.RevOpt1Pd)
    ThisPaid = OldRound(ThisPaid + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt) 'added .DiscAmt on 6/15/07
    YrTaxPaid(y) = OldRound(YrTaxPaid(y) + ThisPaid)
GetNext:
   frmTaxShowPctComp.ShowPctComp x, NumOfTTRecs
   If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdClear.Enabled = True
      cmdProcess.Enabled = True
      Exit Sub
    End If
  Next x
  
  Close
  
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdClear.Enabled = True
  cmdProcess.Enabled = True
  
  If fpcmbPrintOpt.Text = "Graphical" Then
    Call PrintGraphics(YrCnt, YrTaxChrgs(), YrTaxPaid(), TaxYear())
  Else
    frmTaxMsg.Label1.Caption = "Pitch 10 is recommended for this printout."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Call PrintText(YrCnt, YrTaxChrgs(), YrTaxPaid(), TaxYear())
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxCollectRateRpt", "cmdProcess_Click", Erl)
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
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
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
  Me.HelpContextID = hlpCollectionRate
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxCollectRateRpt.")
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
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Integer
  Dim YrCnt As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim BigYr As Integer
  Dim ThisBigYr As Integer
  Dim HoldYr As Integer
  Dim Nextx As Integer
  Dim Thisx As Integer
  
  'on error goto ERRORSTUFF
  
  frmTaxLoadReport.Label1.Caption = "Loading Years"
  frmTaxLoadReport.Show
  DoEvents
  ReDim Years(1 To 1) As Integer
  YrCnt = 0
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    If YrCnt = 0 Then
      If TaxTrans.TaxYear > 0 Then
        YrCnt = YrCnt + 1
        ReDim Preserve Years(1 To YrCnt) As Integer
        Years(YrCnt) = TaxTrans.TaxYear
      End If
    Else
      For y = 1 To YrCnt
        If TaxTrans.TaxYear = Years(y) Then
          Exit For
        End If
      Next y
      If y > YrCnt Then
        YrCnt = YrCnt + 1
        ReDim Preserve Years(1 To YrCnt) As Integer
        Years(YrCnt) = TaxTrans.TaxYear
      End If
    End If
  Next x
  Close TTHandle
  
  BigYr = 0
  For x = 1 To YrCnt
    If Years(x) > BigYr Then
      BigYr = Years(x)
    End If
  Next x
  
  Nextx = 1
  ThisBigYr = BigYr + 1
  Do While Nextx <= YrCnt
    For x = Nextx To YrCnt
      If Years(x) < ThisBigYr Then
        ThisBigYr = Years(x)
        Thisx = x
      End If
    Next x
    HoldYr = Years(Nextx)
    Years(Nextx) = Years(Thisx)
    Years(Thisx) = HoldYr
    Nextx = Nextx + 1
    ThisBigYr = BigYr + 1
  Loop
    
  fpList.AddItem "All"
  For x = YrCnt To 1 Step -1
    fpList.AddItem CStr(Years(x))
  Next x
  
  ReDim AllYears(1 To YrCnt) As Integer
  Nextx = YrCnt
  For x = 1 To YrCnt
    AllYears(Nextx) = Years(x)
    Nextx = Nextx - 1
  Next x
    
  DoEvents
  
  fpList.MultiSelect = MultiSelectNone
  fpList.ListIndex = 0
  fpList.MultiSelect = MultiSelectSimple
  
  Unload frmTaxLoadReport
  fpcmbPrintOpt.Text = "Graphical"
  fpcmbPrintOpt.AddItem "Graphical"
  fpcmbPrintOpt.AddItem "Text"
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxCollectRateRpt", "LoadMe", Erl)
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


Private Sub fpList_Click()
  Dim x As Integer
  
  If fpList.ListIndex = 0 Then
    For x = 1 To fpList.ListCount
      fpList.Row = x
      fpList.Selected = False
    Next x
  ElseIf fpList.ListIndex <> 0 Then
    fpList.Row = 0
    fpList.Selected = False
  End If
  
End Sub

Private Sub PrintGraphics(YrCnt As Integer, YrTaxChrgs() As Double, YrTaxPaid() As Double, TaxYear() As Integer)
  Dim x As Integer
  Dim RptFile As String
  Dim RptHandle As Integer
  Dim dlm$
  Dim PctPaid As Double
  Dim TownName$
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim GTotChrgs As Double
  Dim GTotPaid As Double
  Dim GTotPct As Double
  
  'on error goto ERRORSTUFF
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  TownName = QPTrim$(TaxMasterRec.Name)
  dlm$ = "~"
  
  RptFile$ = "TAXRPTS\COLLECTRT.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  For x = 1 To YrCnt
    If YrTaxChrgs(x) <= 0 Then GoTo NextOne
    PctPaid = YrTaxPaid(x) / YrTaxChrgs(x)
    GTotChrgs = OldRound(GTotChrgs + YrTaxChrgs(x))
    GTotPaid = OldRound(GTotPaid + YrTaxPaid(x))
    If GTotPaid > 0 Then
      GTotPct = GTotPaid / GTotChrgs
    End If
    '
    Print #RptHandle, TownName; dlm; TaxYear(x); dlm; YrTaxChrgs(x); dlm; YrTaxPaid(x); dlm; PctPaid; dlm;
    '                     5               6             7
    Print #RptHandle, GTotChrgs; dlm; GTotPaid; dlm; GTotPct
NextOne:
  Next x
  
  Close
  
  arTaxCollRateRpt.Show
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxCollectRateRpt", "PrintGraphics", Erl)
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

Private Sub PrintText(YrCnt As Integer, YrTaxChrgs() As Double, YrTaxPaid() As Double, TaxYear() As Integer)
  Dim x As Integer
  Dim RptFile As String
  Dim RptHandle As Integer
  Dim PctPaid As Double
  Dim TownName$
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim GTotChrgs As Double
  Dim GTotPaid As Double
  Dim GTotPct As Double
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim FF$
  Dim Page As Integer
  
  'on error goto ERRORSTUFF
  
  FF$ = Chr(12)
  MaxLines = 58
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  TownName = QPTrim$(TaxMasterRec.Name)
  
  RptFile$ = "TAXRPTS\COLLECTRT.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  GoSub PrintHeader
  For x = 1 To YrCnt
    If YrTaxChrgs(x) <= 0 Then GoTo NextOne
    PctPaid = (YrTaxPaid(x) / YrTaxChrgs(x)) * 100
    GTotChrgs = OldRound(GTotChrgs + YrTaxChrgs(x))
    GTotPaid = OldRound(GTotPaid + YrTaxPaid(x))
    If GTotPaid > 0 Then
      GTotPct = (GTotPaid / GTotChrgs) * 100
    End If
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
    Print #RptHandle, Tab(4); Using$("###0", TaxYear(x)); Tab(17); Using$("$###,###,##0.00", YrTaxChrgs(x)); Tab(42); Using$("$###,###,##0.00", YrTaxPaid(x)); Tab(67); Using$("##0.00", PctPaid) + "%"
    LineCnt = LineCnt + 1
NextOne:
  Next x
  If LineCnt >= MaxLines - 3 Then
    Page = Page + 1
    Print #RptHandle, Tab(20); "Tax Collection Rate Report"
    Print #RptHandle, TownName; Tab(65); "Page #: " + CStr(Page)
    Print #RptHandle, "Report Date: " + CStr(Date)
    Print #RptHandle,
    Print #RptHandle, String(80, "-")
  End If
  
  Print #RptHandle,
  Print #RptHandle, String(80, "-")
  Print #RptHandle, Tab(2); "Totals: "; Tab(17); Using$("$###,###,##0.00", GTotChrgs); Tab(42); Using$("$###,###,##0.00", GTotPaid); Tab(67); Using$("##0.00", GTotPct) + "%"
  Print #RptHandle, FF$
  Close
  
  ViewPrint RptFile, "Tax Collections Rate Report, True"
  
  Exit Sub
  
PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Tax Collection Rate Report"
  Print #RptHandle, TownName; Tab(65); "Page #: " + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle,
  Print #RptHandle, Tab(2); "Tax Year"; Tab(17); "Charges For Year"; Tab(37); "Collections For Year"; Tab(60); "Collection Percentage"
  Print #RptHandle, String(80, "-")
  LineCnt = 6
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxCollectRateRpt", "PrintText", Erl)
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
