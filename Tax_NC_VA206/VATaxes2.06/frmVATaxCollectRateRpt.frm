VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmVATaxCollectRateRpt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Collection Rate Report"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxCollectRateRpt.frx":0000
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
      Picture         =   "frmVATaxCollectRateRpt.frx":08CA
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   384
         Left            =   2928
         TabIndex        =   1
         Top             =   3888
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
         ColDesigner     =   "frmVATaxCollectRateRpt.frx":08E6
      End
      Begin LpLib.fpList fpList 
         Height          =   1776
         Left            =   1560
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
         ColDesigner     =   "frmVATaxCollectRateRpt.frx":0BDD
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00D0D0D0&
         Height          =   1932
         Left            =   3720
         TabIndex        =   9
         Top             =   1800
         Width           =   2772
         Begin VB.OptionButton OptBoth 
            BackColor       =   &H00D0D0D0&
            Caption         =   "Use Both Summary"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   12
            Top             =   396
            Width           =   2172
         End
         Begin VB.OptionButton OptReal 
            BackColor       =   &H00D0D0D0&
            Caption         =   "Real Only Summary"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   11
            Top             =   864
            Width           =   2172
         End
         Begin VB.OptionButton OptPers 
            BackColor       =   &H00D0D0D0&
            Caption         =   "Personal Only Summary"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   10
            Top             =   1332
            Width           =   2412
         End
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   636
         Left            =   960
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   5160
         Width           =   1620
         _Version        =   131072
         _ExtentX        =   2857
         _ExtentY        =   1122
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
         ButtonDesigner  =   "frmVATaxCollectRateRpt.frx":0E69
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   636
         Left            =   5232
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   $"frmVATaxCollectRateRpt.frx":1047
         Top             =   5160
         Width           =   1740
         _Version        =   131072
         _ExtentX        =   3069
         _ExtentY        =   1122
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
         ButtonDesigner  =   "frmVATaxCollectRateRpt.frx":10F2
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdClear 
         Height          =   636
         Left            =   2796
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   5160
         Width           =   2220
         _Version        =   131072
         _ExtentX        =   3916
         _ExtentY        =   1122
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
         ButtonDesigner  =   "frmVATaxCollectRateRpt.frx":12D1
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
         Height          =   252
         Left            =   1800
         TabIndex        =   5
         Top             =   1560
         Width           =   1092
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   3300
         Left            =   1008
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
         Height          =   390
         Left            =   1800
         TabIndex        =   4
         Top             =   450
         Width           =   4335
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   1530
         Top             =   315
         Width           =   4905
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
         Height          =   348
         Left            =   1272
         TabIndex        =   3
         Top             =   4008
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
Attribute VB_Name = "frmVATaxCollectRateRpt"
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
  frmVATaxReportsMenu.Show
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
  
  On Error GoTo ERRORSTUFF
'  If OptRealD.Value = True Or OptPersD.Value = True Then
'    Call ProcessDet
'    Exit Sub
'  End If
  
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
      TaxYear(x) = AllYears(x)
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
  frmVATaxShowPctComp.Label1 = "Creating Collection Rate Report"
  frmVATaxShowPctComp.Show , Me
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
    Get TCHandle, TaxTrans.CustomerRec, TaxCust
    If OptReal.Value = True And TaxTrans.BillType = "P" Then GoTo GetNext
    If OptPers.Value = True And TaxTrans.BillType = "R" Then GoTo GetNext
    
    If TaxCust.Deleted <> 0 Then GoTo GetNext
    ThisPaid = 0
    ThisChrg = 0

    ThisChrg = OldRound(ThisChrg + TaxTrans.Revenue.Collection + TaxTrans.Revenue.Interest)
    ThisChrg = OldRound(ThisChrg + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Penalty)
    ThisChrg = OldRound(ThisChrg + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2)
    ThisChrg = OldRound(ThisChrg + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4)
    ThisChrg = OldRound(ThisChrg + TaxTrans.Revenue.Principle5 + TaxTrans.Revenue.RevOpt1)
    ThisChrg = OldRound(ThisChrg + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3 + TaxTrans.PPTRARmvl)
    YrTaxChrgs(y) = OldRound(YrTaxChrgs(y) + ThisChrg)
    ThisPaid = OldRound(TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.InterestPd)
    ThisPaid = OldRound(ThisPaid + TaxTrans.Revenue.LateListPd + TaxTrans.Revenue.PenaltyPd)
    ThisPaid = OldRound(ThisPaid + TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd)
    ThisPaid = OldRound(ThisPaid + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd)
    ThisPaid = OldRound(ThisPaid + TaxTrans.Revenue.Principle5Pd + TaxTrans.Revenue.RevOpt1Pd)
    ThisPaid = OldRound(ThisPaid + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.PPTRADisc)
    
    YrTaxPaid(y) = OldRound(YrTaxPaid(y) + ThisPaid)

GetNext:
   frmVATaxShowPctComp.ShowPctComp x, NumOfTTRecs
   If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdClear.Enabled = True
      cmdProcess.Enabled = True
      Exit Sub
    End If
  Next x
  
  Close
  
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdClear.Enabled = True
  cmdProcess.Enabled = True
  
  If fpcmbPrintOpt.Text = "Graphical" Then
    Call PrintGraphics(YrCnt, YrTaxChrgs(), YrTaxPaid(), TaxYear())
  Else
    frmVATaxMsg.Label1.Caption = "Pitch 10 is recommended for this printout."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Call PrintText(YrCnt, YrTaxChrgs(), YrTaxPaid(), TaxYear())
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxCollectRateRpt", "cmdProcess_Click", Erl)
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
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxCollectRateRpt.")
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
  
  On Error GoTo ERRORSTUFF
  
  OptBoth.Value = True
  frmVATaxLoadReport.Label1.Caption = "Loading Years"
  frmVATaxLoadReport.Show
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
  
  Unload frmVATaxLoadReport
  fpcmbPrintOpt.Text = "Graphical"
  fpcmbPrintOpt.AddItem "Graphical"
  fpcmbPrintOpt.AddItem "Text"
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxCollectRateRpt", "LoadMe", Erl)
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
  Dim BillType$
  
  On Error GoTo ERRORSTUFF
  
  If OptBoth.Value = True Then
    BillType$ = "Real and Personal"
  ElseIf OptReal.Value = True Then
    BillType$ = "Real Only"
  ElseIf OptPers.Value = True Then
    BillType$ = "Personal Only"
  End If
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
    '                     5               6             7             8
    Print #RptHandle, GTotChrgs; dlm; GTotPaid; dlm; GTotPct; dlm; BillType$
NextOne:
  Next x
  
  Close
  
  arVATaxCollRateRpt.Show
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxCollectRateRpt", "PrintGraphics", Erl)
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
  Dim BillType$
  
  On Error GoTo ERRORSTUFF
  
  If OptBoth.Value = True Then
    BillType$ = "Real and Personal"
  ElseIf OptReal.Value = True Then
    BillType$ = "Real Only"
  ElseIf OptPers.Value = True Then
    BillType$ = "Personal Only"
  End If
  
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
  Print #RptHandle, "BillType: " + BillType$
  Print #RptHandle,
  Print #RptHandle, Tab(2); "Tax Year"; Tab(17); "Charges For Year"; Tab(37); "Collections For Year"; Tab(60); "Collection Percentage"
  Print #RptHandle, String(80, "-")
  LineCnt = 7
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxCollectRateRpt", "PrintText", Erl)
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

Private Sub PrintGraphicsDet(YrCnt As Integer, YrTaxChrgs() As Double, YrTaxPaid() As Double, TaxYear() As Integer)
'  Dim x As Integer
'  Dim RptFile As String
'  Dim RptHandle As Integer
'  Dim dlm$
'  Dim PctPaid As Double
'  Dim TownName$
'  Dim TaxMasterRec As TaxMasterType
'  Dim TMHandle As Integer
'  Dim GTotChrgs As Double
'  Dim GTotPaid As Double
'  Dim GTotPct As Double
'  Dim BillType$
'
'  On Error GoTo ERRORSTUFF
''  AllYears (x)
'  If OptRealD.Value = True Then
'    BillType$ = "Real Only"
'  ElseIf OptPersD.Value = True Then
'    BillType$ = "Personal Only"
'  End If
'
'  OpenTaxSetUpFile TMHandle
'  Get TMHandle, 1, TaxMasterRec
'  Close TMHandle
'
'  TownName = QPTrim$(TaxMasterRec.Name)
'  dlm$ = "~"
'
'  RptFile$ = "TAXRPTS\COLLECTRT.RPT"
'  RptHandle = FreeFile
'  Open RptFile For Output As #RptHandle
'
'  For x = 1 To YrCnt
'    If YrTaxChrgs(x) <= 0 Then GoTo NextOne
'    PctPaid = YrTaxPaid(x) / YrTaxChrgs(x)
'    GTotChrgs = OldRound(GTotChrgs + YrTaxChrgs(x))
'    GTotPaid = OldRound(GTotPaid + YrTaxPaid(x))
'    If GTotPaid > 0 Then
'      GTotPct = GTotPaid / GTotChrgs
'    End If
'    '
'    Print #RptHandle, TownName; dlm; TaxYear(x); dlm; YrTaxChrgs(x); dlm; YrTaxPaid(x); dlm; PctPaid; dlm;
'    '                     5               6             7             8
'    Print #RptHandle, GTotChrgs; dlm; GTotPaid; dlm; GTotPct; dlm; BillType$
'NextOne:
'  Next x
'
'  Close
'
'  arVATaxCollRateRpt.Show
'
'  Exit Sub
'
'ERRORSTUFF:
'   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxCollectRateRpt", "PrintGraphics", Erl)
'     Case emrExitProc:
'       Resume Proc_Exit
'     Case emrResume:
'       Resume
'     Case emrResumeNext:
'       Resume Next
'     Case Else
'      '--- Technically, this should never happen.
'       Resume Proc_Exit
'   End Select
'
'Proc_Exit:
'  '--- Cleanup code goes here...
'    Close
End Sub

Private Sub ProcessDet()
'  Dim TaxTrans As TaxTransactionType
'  Dim TTHandle As Integer
'  Dim NumOfTTRecs As Long
'  Dim x As Long, y As Integer
'  Dim NextRec As Long
'  Dim ThisYear As Integer
'  Dim YrCnt As Integer
'  Dim ThisPaid As Double
'  Dim ThisChrg As Double
'  Dim TaxCust As TaxCustType
'  Dim TCHandle As Integer
'  Dim NumOfTCRecs As Long
'  Dim PctPaid As Double
'  Dim PctPrincPaid As Double
'  Dim PctRIntPaid As Double
'  Dim PctAdvPaid As Double
'  Dim PctLateListPaid As Double
'  Dim PctRPenPaid As Double
'  Dim PctROpt1Paid As Double
'  Dim PctROpt2Paid As Double
'  Dim PctROpt3Paid As Double
'  Dim PctPersPaid As Double
'  Dim PctMTPaid As Double
'  Dim PctMCPaid As Double
'  Dim PctFEPaid As Double
'  Dim PctMHPaid As Double
'  Dim PctPIntPaid As Double
'  Dim PctPPenPaid As Double
'  Dim PctPOpt1Paid As Double
'  Dim PctPOpt2Paid As Double
'  Dim PctPOpt3Paid As Double
'  Dim RptFile As String
'  Dim RptHandle As Integer
'  Dim dlm$
'  Dim TownName$
'  Dim TaxMasterRec As TaxMasterType
'  Dim TMHandle As Integer
'  Dim GTotChrgs As Double
'  Dim GTotPrincChrgs As Double
'  Dim GTotRIntChrgs As Double
'  Dim GTotAdvChrgs As Double
'  Dim GTotLateListChrgs As Double
'  Dim GTotRPenChrgs As Double
'  Dim GTotROpt1Chrgs As Double
'  Dim GTotROpt2Chrgs As Double
'  Dim GTotROpt3Chrgs As Double
'  Dim GTotPersChrgs As Double
'  Dim GTotMTChrgs As Double
'  Dim GTotMCChrgs As Double
'  Dim GTotFEChrgs As Double
'  Dim GTotMHChrgs As Double
'  Dim GTotPIntChrgs As Double
'  Dim GTotPPenChrgs As Double
'  Dim GTotPOpt1Chrgs As Double
'  Dim GTotPOpt2Chrgs As Double
'  Dim GTotPOpt3Chrgs As Double
'  Dim GTotPaid As Double
'  Dim GTotPrincPaid As Double
'  Dim GTotRIntPaid As Double
'  Dim GTotAdvPaid As Double
'  Dim GTotLateListPaid As Double
'  Dim GTotRPenPaid As Double
'  Dim GTotROpt1Paid As Double
'  Dim GTotROpt2Paid As Double
'  Dim GTotROpt3Paid As Double
'  Dim GTotPersPaid As Double
'  Dim GTotMTPaid As Double
'  Dim GTotMCPaid As Double
'  Dim GTotFEPaid As Double
'  Dim GTotMHPaid As Double
'  Dim GTotPIntPaid As Double
'  Dim GTotPPenPaid As Double
'  Dim GTotPOpt1Paid As Double
'  Dim GTotPOpt2Paid As Double
'  Dim GTotPOpt3Paid As Double
'  Dim GTotPct As Double
'  Dim GTotPrincPct As Double
'  Dim GTotRIntPct As Double
'  Dim GTotAdvPct As Double
'  Dim GTotLateListPct As Double
'  Dim GTotRPenPct As Double
'  Dim GTotROpt1Pct As Double
'  Dim GTotROpt2Pct As Double
'  Dim GTotROpt3Pct As Double
'  Dim GTotPersPct As Double
'  Dim GTotMTPct As Double
'  Dim GTotMCPct As Double
'  Dim GTotFEPct As Double
'  Dim GTotMHPct As Double
'  Dim GTotPIntPct As Double
'  Dim GTotPPenPct As Double
'  Dim GTotPOpt1Pct As Double
'  Dim GTotPOpt2Pct As Double
'  Dim GTotPOpt3Pct As Double
'  Dim LineCnt As Integer
'  Dim MaxLines As Integer
'  Dim FF$
'  Dim Page As Integer
'  Dim BillType$
'
'  On Error GoTo ERRORSTUFF
'
'  If fpList.ListIndex = -1 Then
'    Call TaxMsg(900, "Please select a year or years from the list.")
'    Exit Sub
'  End If
'
'  OpenTaxSetUpFile TMHandle
'  Get TMHandle, 1, TaxMasterRec
'  Close TMHandle
'
'  TownName = QPTrim$(TaxMasterRec.Name)
'
'  If OptReal.Value = True Then
'    BillType$ = "Real Only"
'  ElseIf OptPers.Value = True Then
'    BillType$ = "Personal Only"
'  End If
'
'  fpList.Row = 0
'  If fpList.Selected = True Then
'    AllFlag = True
'    YrCnt = fpList.ListCount - 1
'    ReDim TaxYear(1 To YrCnt) As Integer
'    For x = 1 To YrCnt
'      TaxYear(x) = AllYears(x)
'    Next x
'  Else
'    AllFlag = False
'    ReDim TaxYear(1 To 1) As Integer
'    For x = 1 To fpList.ListCount - 1
'      fpList.Row = x
'      If fpList.Selected = True Then
'        fpList.ListIndex = x
'        YrCnt = YrCnt + 1
'        ReDim Preserve TaxYear(1 To YrCnt) As Integer
'        TaxYear(YrCnt) = CInt(fpList.Text)
'      End If
'    Next x
'  End If
'
'  ReDim YrTaxChrgs(1 To YrCnt) As Double
'
'  ReDim YrPrincChrgs(1 To YrCnt) As Double
'  ReDim YrRIntChrgs(1 To YrCnt) As Double
'  ReDim YrAdvChrgs(1 To YrCnt) As Double
'  ReDim YrLateListChrgs(1 To YrCnt) As Double
'  ReDim YrRPenChrgs(1 To YrCnt) As Double
'  ReDim YrROpt1Chrgs(1 To YrCnt) As Double
'  ReDim YrROpt2Chrgs(1 To YrCnt) As Double
'  ReDim YrROPt3Chrgs(1 To YrCnt) As Double
'
'  ReDim YrPersChrgs(1 To YrCnt) As Double
'  ReDim YrMTChrgs(1 To YrCnt) As Double
'  ReDim YrMCChrgs(1 To YrCnt) As Double
'  ReDim YrFEChrgs(1 To YrCnt) As Double
'  ReDim YrMHChrgs(1 To YrCnt) As Double
'  ReDim YrPIntChrgs(1 To YrCnt) As Double
'  ReDim YrPPenChrgs(1 To YrCnt) As Double
'  ReDim YrPOpt1Chrgs(1 To YrCnt) As Double
'  ReDim YrPOpt2Chrgs(1 To YrCnt) As Double
'  ReDim YrPOpt3Chrgs(1 To YrCnt) As Double
'
'  ReDim YrTaxPaid(1 To YrCnt) As Double
'
'  ReDim YrPrincPaid(1 To YrCnt) As Double
'  ReDim YrRIntPaid(1 To YrCnt) As Double
'  ReDim YrAdvPaid(1 To YrCnt) As Double
'  ReDim YrLateListPaid(1 To YrCnt) As Double
'  ReDim YrRPenPaid(1 To YrCnt) As Double
'  ReDim YrROpt1Paid(1 To YrCnt) As Double
'  ReDim YrROpt2Paid(1 To YrCnt) As Double
'  ReDim YrROpt3Paid(1 To YrCnt) As Double
'
'  ReDim YrPersPaid(1 To YrCnt) As Double
'  ReDim YrMTPaid(1 To YrCnt) As Double
'  ReDim YrMCPaid(1 To YrCnt) As Double
'  ReDim YrFEPaid(1 To YrCnt) As Double
'  ReDim YrMHPaid(1 To YrCnt) As Double
'  ReDim YrPIntPaid(1 To YrCnt) As Double
'  ReDim YrPPenPaid(1 To YrCnt) As Double
'  ReDim YrPOpt1Paid(1 To YrCnt) As Double
'  ReDim YrPOpt2Paid(1 To YrCnt) As Double
'  ReDim YrPOpt3Paid(1 To YrCnt) As Double
'
'  OpenTaxTransFile TTHandle, NumOfTTRecs
'  OpenTaxCustFile TCHandle, NumOfTCRecs
'  frmVATaxShowPctComp.Label1 = "Creating Collection Rate Report"
'  frmVATaxShowPctComp.Show , Me
'  EnableCloseButton Me.hwnd, False
'  cmdExit.Enabled = False
'  cmdClear.Enabled = False
'  cmdProcess.Enabled = False
'
'  For x = 1 To NumOfTTRecs
'    Get TTHandle, x, TaxTrans
'    For y = 1 To YrCnt
'      If TaxTrans.TaxYear = TaxYear(y) Then
'        Exit For
'      End If
'    Next y
'    If y > YrCnt Then GoTo GetNext
'    If TaxTrans.TranType <> 1 Then GoTo GetNext
'    Get TCHandle, TaxTrans.CustomerRec, TaxCust
'    If OptRealD.Value = True And TaxTrans.BillType = "P" Then GoTo GetNext
'    If OptPersD.Value = True And TaxTrans.BillType = "R" Then GoTo GetNext
'
'    If TaxCust.Deleted <> 0 Then GoTo GetNext
'    ThisPaid = 0
'    ThisChrg = 0
'
'    ThisChrg = OldRound(ThisChrg + TaxTrans.Revenue.Collection + TaxTrans.Revenue.Interest)
'    ThisChrg = OldRound(ThisChrg + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Penalty)
'    ThisChrg = OldRound(ThisChrg + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2)
'    ThisChrg = OldRound(ThisChrg + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4)
'    ThisChrg = OldRound(ThisChrg + TaxTrans.Revenue.Principle5 + TaxTrans.Revenue.RevOpt1)
'    ThisChrg = OldRound(ThisChrg + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3 + TaxTrans.PPTRARmvl)
'    YrTaxChrgs(y) = OldRound(YrTaxChrgs(y) + ThisChrg)
'    If OptRealD.Value = True Then
'      YrPrincChrgs(y) = OldRound(YrPrincChrgs(y) + TaxTrans.Revenue.Principle1)
'      YrRIntChrgs(y) = OldRound(YrRIntChrgs(y) + TaxTrans.Revenue.Interest)
'      YrAdvChrgs(y) = OldRound(YrAdvChrgs(y) + TaxTrans.Revenue.Collection)
'      YrLateListChrgs(y) = OldRound(YrLateListChrgs(y) + TaxTrans.Revenue.LateList)
'      YrRPenChrgs(y) = OldRound(YrRPenChrgs(y) + TaxTrans.Revenue.Penalty)
'      YrROpt1Chrgs(y) = OldRound(YrROpt1Chrgs(y) + TaxTrans.Revenue.RevOpt1)
'      YrROpt2Chrgs(y) = OldRound(YrROpt2Chrgs(y) + TaxTrans.Revenue.RevOpt2)
'      YrROPt3Chrgs(y) = OldRound(YrROPt3Chrgs(y) + TaxTrans.Revenue.RevOpt3)
'    ElseIf OptPersD.Value = True Then
'      YrPersChrgs(y) = OldRound(YrPersChrgs(y) + TaxTrans.Revenue.Principle1 + TaxTrans.PPTRARmvl)
'      YrPIntChrgs(y) = OldRound(YrPIntChrgs(y) + TaxTrans.Revenue.Interest)
'      YrMTChrgs(y) = OldRound(YrMTChrgs(y) + TaxTrans.Revenue.Principle2)
'      YrMCChrgs(y) = OldRound(YrMCChrgs(y) + TaxTrans.Revenue.Principle3)
'      YrFEChrgs(y) = OldRound(YrFEChrgs(y) + TaxTrans.Revenue.Principle4)
'      YrMHChrgs(y) = OldRound(YrMHChrgs(y) + TaxTrans.Revenue.Principle5)
'      YrPPenChrgs(y) = OldRound(YrPPenChrgs(y) + TaxTrans.Revenue.Penalty)
'      YrPOpt1Chrgs(y) = OldRound(YrPOpt1Chrgs(y) + TaxTrans.Revenue.RevOpt1)
'      YrPOpt2Chrgs(y) = OldRound(YrPOpt2Chrgs(y) + TaxTrans.Revenue.RevOpt2)
'      YrPOpt3Chrgs(y) = OldRound(YrPOpt3Chrgs(y) + TaxTrans.Revenue.RevOpt3)
'    End If
'
'    ThisPaid = OldRound(TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.InterestPd)
'    ThisPaid = OldRound(ThisPaid + TaxTrans.Revenue.LateListPd + TaxTrans.Revenue.PenaltyPd)
'    ThisPaid = OldRound(ThisPaid + TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd)
'    ThisPaid = OldRound(ThisPaid + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd)
'    ThisPaid = OldRound(ThisPaid + TaxTrans.Revenue.Principle5Pd + TaxTrans.Revenue.RevOpt1Pd)
'    ThisPaid = OldRound(ThisPaid + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.PPTRADisc)
'
'    YrTaxPaid(y) = OldRound(YrTaxPaid(y) + ThisPaid)
'
'    If OptRealD.Value = True Then
'      YrPrincPaid(y) = OldRound(YrPrincPaid(y) + TaxTrans.Revenue.Principle1Pd)
'      YrRIntPaid(y) = OldRound(YrRIntPaid(y) + TaxTrans.Revenue.InterestPd)
'      YrAdvPaid(y) = OldRound(YrAdvPaid(y) + TaxTrans.Revenue.CollectionPd)
'      YrLateListPaid(y) = OldRound(YrLateListPaid(y) + TaxTrans.Revenue.LateListPd)
'      YrRPenPaid(y) = OldRound(YrRPenPaid(y) + TaxTrans.Revenue.PenaltyPd)
'      YrROpt1Paid(y) = OldRound(YrROpt1Paid(y) + TaxTrans.Revenue.RevOpt1Pd)
'      YrROpt2Paid(y) = OldRound(YrROpt2Paid(y) + TaxTrans.Revenue.RevOpt2Pd)
'      YrROpt3Paid(y) = OldRound(YrROpt3Paid(y) + TaxTrans.Revenue.RevOpt3Pd)
'    ElseIf OptPersD.Value = True Then
'      YrPersPaid(y) = OldRound(YrPersPaid(y) + TaxTrans.Revenue.Principle1Pd + TaxTrans.PPTRADisc)
'      YrPIntPaid(y) = OldRound(YrPIntPaid(y) + TaxTrans.Revenue.InterestPd)
'      YrMTPaid(y) = OldRound(YrMTPaid(y) + TaxTrans.Revenue.Principle2Pd)
'      YrMCPaid(y) = OldRound(YrMCPaid(y) + TaxTrans.Revenue.Principle3Pd)
'      YrFEPaid(y) = OldRound(YrFEPaid(y) + TaxTrans.Revenue.Principle4Pd)
'      YrMHPaid(y) = OldRound(YrMHPaid(y) + TaxTrans.Revenue.Principle5Pd)
'      YrPPenPaid(y) = OldRound(YrPPenPaid(y) + TaxTrans.Revenue.PenaltyPd)
'      YrPOpt1Paid(y) = OldRound(YrPOpt1Paid(y) + TaxTrans.Revenue.RevOpt1Pd)
'      YrPOpt2Paid(y) = OldRound(YrPOpt2Paid(y) + TaxTrans.Revenue.RevOpt2Pd)
'      YrPOpt3Paid(y) = OldRound(YrPOpt3Paid(y) + TaxTrans.Revenue.RevOpt3Pd)
'    End If
'GetNext:
'   frmVATaxShowPctComp.ShowPctComp x, NumOfTTRecs
'   If frmVATaxShowPctComp.Out = True Then
'      Close
'      frmVATaxShowPctComp.Out = False
'      Unload frmVATaxShowPctComp
'      EnableCloseButton Me.hwnd, True
'      cmdExit.Enabled = True
'      cmdClear.Enabled = True
'      cmdProcess.Enabled = True
'      Exit Sub
'    End If
'  Next x
'
'  Close
'
'  Unload frmVATaxShowPctComp
'  EnableCloseButton Me.hwnd, True
'  cmdExit.Enabled = True
'  cmdClear.Enabled = True
'  cmdProcess.Enabled = True
'
'  If fpcmbPrintOpt.Text = "Graphical" Then
'    GoSub PrintGraphics
'  Else
'    frmVATaxMsg.Label1.Caption = "Pitch 10 is recommended for this printout."
'    frmVATaxMsg.Label1.Top = 900
'    frmVATaxMsg.Show vbModal
'    GoSub PrintText
'  End If
'
'  Exit Sub
'
'PrintGraphics:
'  dlm$ = "~"
'
'  RptFile$ = "TAXRPTS\COLLECTRTD.RPT"
'  RptHandle = FreeFile
'  Open RptFile For Output As #RptHandle
'
'  For x = 1 To YrCnt
'    If YrTaxChrgs(x) <= 0 Then GoTo NextOne
'    PctPaid = YrTaxPaid(x) / YrTaxChrgs(x)
'    GTotChrgs = OldRound(GTotChrgs + YrTaxChrgs(x))
'    GTotPaid = OldRound(GTotPaid + YrTaxPaid(x))
'    If GTotPaid > 0 Then
'      GTotPct = GTotPaid / GTotChrgs
'    End If
'    If BillType = "Real Only" Then
'      GTotPrincChrgs = OldRound(GTotPrincChrgs + YrPrincChrgs(x))
'      GTotPrincPaid = OldRound(GTotPrincPaid + YrPrincPaid(x))
'      If GTotPrincPaid > 0 Then
'        GTotPrincPct = GTotPrincPaid / GTotPrincChrgs
'      End If
'      If YrPrincChrgs(x) > 0 Then
'        PctPrincPaid = YrPrincPaid(x) / YrPrincChrgs(x)
'      Else
'        PctPrincPaid = 0
'      End If
'      GTotRIntChrgs = OldRound(GTotRIntChrgs + YrRIntChrgs(x))
'      GTotRIntPaid = OldRound(GTotRIntPaid + YrRIntPaid(x))
'      If GTotRIntPaid > 0 Then
'        GTotRIntPct = GTotRIntPaid / GTotRIntChrgs
'      End If
'      If YrRIntChrgs(x) > 0 Then
'        PctRIntPaid = YrRIntPaid(x) / YrRIntChrgs(x)
'      Else
'        PctRIntPaid = 0
'      End If
'      GTotAdvChrgs = OldRound(GTotAdvChrgs + YrAdvChrgs(x))
'      GTotAdvPaid = OldRound(GTotAdvPaid + YrAdvPaid(x))
'      If GTotAdvPaid > 0 Then
'        GTotAdvPct = GTotAdvPaid / GTotAdvChrgs
'      End If
'      If YrAdvChrgs(x) > 0 Then
'        PctAdvPaid = YrAdvPaid(x) / YrAdvChrgs(x)
'      Else
'        PctAdvPaid = 0
'      End If
'      GTotLateListChrgs = OldRound(GTotLateListChrgs + YrLateListChrgs(x))
'      GTotLateListPaid = OldRound(GTotLateListPaid + YrLateListPaid(x))
'      If GTotLateListPaid > 0 Then
'        GTotLateListPct = GTotLateListPaid / GTotLateListChrgs
'      End If
'      If YrLateListChrgs(x) > 0 Then
'        PctLateListPaid = YrLateListPaid(x) / YrLateListChrgs(x)
'      Else
'        PctLateListPaid = 0
'      End If
'      GTotRPenChrgs = OldRound(GTotRPenChrgs + YrRPenChrgs(x))
'      GTotRPenPaid = OldRound(GTotRPenPaid + YrRPenPaid(x))
'      If GTotRPenPaid > 0 Then
'        GTotRPenPct = GTotRPenPaid / GTotRPenChrgs
'      End If
'      If YrRPenChrgs(x) > 0 Then
'        PctRPenPaid = YrRPenPaid(x) / YrRPenChrgs(x)
'      Else
'        PctRPenPaid = 0
'      End If
'      GTotROpt1Chrgs = OldRound(GTotROpt1Chrgs + YrROpt1Chrgs(x))
'      GTotROpt1Paid = OldRound(GTotROpt1Paid + YrROpt1Paid(x))
'      If GTotROpt1Paid > 0 Then
'        GTotROpt1Pct = GTotROpt1Paid / GTotROpt1Chrgs
'      End If
'      If YrROpt1Chrgs(x) > 0 Then
'        PctROpt1Paid = YrROpt1Paid(x) / YrROpt1Chrgs(x)
'      Else
'        PctROpt1Paid = 0
'      End If
'      GTotROpt2Chrgs = OldRound(GTotROpt2Chrgs + YrROpt2Chrgs(x))
'      GTotROpt2Paid = OldRound(GTotROpt2Paid + YrROpt2Paid(x))
'      If GTotROpt2Paid > 0 Then
'        GTotROpt2Pct = GTotROpt2Paid / GTotROpt2Chrgs
'      End If
'      If YrROpt2Chrgs(x) > 0 Then
'        PctROpt2Paid = YrROpt2Paid(x) / YrROpt2Chrgs(x)
'      Else
'        PctROpt2Paid = 0
'      End If
'      GTotROpt3Chrgs = OldRound(GTotROpt3Chrgs + YrROPt3Chrgs(x))
'      GTotROpt3Paid = OldRound(GTotROpt3Paid + YrROpt3Paid(x))
'      If GTotROpt3Paid > 0 Then
'        GTotROpt3Pct = GTotROpt3Paid / GTotROpt3Chrgs
'      End If
'      If YrROPt3Chrgs(x) > 0 Then
'        PctROpt3Paid = YrROpt3Paid(x) / YrROPt3Chrgs(x)
'      Else
'        PctROpt3Paid = 0
'      End If
'      PctPersPaid = 0
'      PctMTPaid = 0
'      PctMCPaid = 0
'      PctFEPaid = 0
'      PctMHPaid = 0
'      PctPIntPaid = 0
'      PctPPenPaid = 0
'      PctPOpt1Paid = 0
'      PctPOpt2Paid = 0
'      PctPOpt3Paid = 0
'    ElseIf BillType = "Personal Only" Then
'      GTotPersChrgs = OldRound(GTotPersChrgs + YrPersChrgs(x))
'      GTotPersPaid = OldRound(GTotPersPaid + YrPersPaid(x))
'      If GTotPersPaid > 0 Then
'        GTotPersPct = GTotPersPaid / GTotPersChrgs
'      End If
'      If YrPersChrgs(x) > 0 Then
'        PctPersPaid = YrPersPaid(x) / YrPersChrgs(x)
'      Else
'        PctPersPaid = 0
'      End If
'      GTotPIntChrgs = OldRound(GTotPIntChrgs + YrPIntChrgs(x))
'      GTotPIntPaid = OldRound(GTotPIntPaid + YrPIntPaid(x))
'      If GTotPIntPaid > 0 Then
'        GTotPIntPct = GTotPIntPaid / GTotPIntChrgs
'      End If
'      If YrPIntChrgs(x) > 0 Then
'        PctPIntPaid = YrPIntPaid(x) / YrPIntChrgs(x)
'      Else
'        PctPIntPaid = 0
'      End If
'      GTotMTChrgs = OldRound(GTotMTChrgs + YrMTChrgs(x))
'      GTotMTPaid = OldRound(GTotMTPaid + YrMTPaid(x))
'      If GTotMTPaid > 0 Then
'        GTotMTPct = GTotMTPaid / GTotMTChrgs
'      End If
'      If YrMTChrgs(x) > 0 Then
'        PctMTPaid = YrMTPaid(x) / YrMTChrgs(x)
'      Else
'        PctMTPaid = 0
'      End If
'      GTotMCChrgs = OldRound(GTotMCChrgs + YrMCChrgs(x))
'      GTotMCPaid = OldRound(GTotMCPaid + YrMCPaid(x))
'      If GTotMCPaid > 0 Then
'        GTotMCPct = GTotMCPaid / GTotMCChrgs
'      End If
'      If YrMCChrgs(x) > 0 Then
'        PctMCPaid = YrMCPaid(x) / YrMCChrgs(x)
'      Else
'        PctMCPaid = 0
'      End If
'      GTotFEPaid = OldRound(GTotFEPaid + YrFEPaid(x))
'      GTotFEChrgs = OldRound(GTotFEChrgs + YrFEChrgs(x))
'      If GTotFEPaid > 0 Then
'        GTotFEPct = GTotFEPaid / GTotFEChrgs
'      End If
'      If YrFEChrgs(x) > 0 Then
'        PctFEPaid = YrFEPaid(x) / YrFEChrgs(x)
'      Else
'        PctFEPaid = 0
'      End If
'      GTotMHChrgs = OldRound(GTotMHChrgs + YrMHChrgs(x))
'      GTotMHPaid = OldRound(GTotMHPaid + YrMHPaid(x))
'      If GTotMHPaid > 0 Then
'        GTotMHPct = GTotMHPaid / GTotMHChrgs
'      End If
'      If YrMHChrgs(x) > 0 Then
'        PctMHPaid = YrMHPaid(x) / YrMHChrgs(x)
'      Else
'        PctMHPaid = 0
'      End If
'      GTotPPenChrgs = OldRound(GTotPPenChrgs + YrPPenChrgs(x))
'      GTotPPenPaid = OldRound(GTotPPenPaid + YrPPenPaid(x))
'      If GTotPPenPaid > 0 Then
'        GTotPPenPct = GTotPPenPaid / GTotPPenChrgs
'      End If
'      If YrPPenChrgs(x) > 0 Then
'        PctPPenPaid = YrPPenPaid(x) / YrPPenChrgs(x)
'      Else
'        PctPPenPaid = 0
'      End If
'      GTotPOpt1Chrgs = OldRound(GTotPOpt1Chrgs + YrPOpt1Chrgs(x))
'      GTotPOpt1Paid = OldRound(GTotPOpt1Paid + YrPOpt1Paid(x))
'      If GTotPOpt1Paid > 0 Then
'        GTotPOpt1Pct = GTotPOpt1Paid / GTotPOpt1Chrgs
'      End If
'      If YrPOpt1Chrgs(x) > 0 Then
'        PctPOpt1Paid = YrPOpt1Paid(x) / YrPOpt1Chrgs(x)
'      Else
'        PctPOpt1Paid = 0
'      End If
'      GTotPOpt2Chrgs = OldRound(GTotPOpt2Chrgs + YrPOpt2Chrgs(x))
'      GTotPOpt2Paid = OldRound(GTotPOpt2Paid + YrPOpt2Paid(x))
'      If GTotPOpt2Paid > 0 Then
'        GTotPOpt2Pct = GTotPOpt2Paid / GTotPOpt2Chrgs
'      End If
'      If YrPOpt2Chrgs(x) > 0 Then
'        PctPOpt2Paid = YrPOpt2Paid(x) / YrPOpt2Chrgs(x)
'      Else
'        PctPOpt2Paid = 0
'      End If
'      GTotPOpt3Chrgs = OldRound(GTotPOpt3Chrgs + YrPOpt3Chrgs(x))
'      GTotPOpt3Paid = OldRound(GTotPOpt3Paid + YrPOpt3Paid(x))
'      If GTotPOpt3Paid > 0 Then
'        GTotPOpt3Pct = GTotPOpt3Paid / GTotPOpt3Chrgs
'      End If
'      If YrPOpt3Chrgs(x) > 0 Then
'        PctPOpt3Paid = YrPOpt3Paid(x) / YrPOpt3Chrgs(x)
'      Else
'        PctPOpt3Paid = 0
'      End If
'      PctPrincPaid = 0
'      PctAdvPaid = 0
'      PctLateListPaid = 0
'      PctRIntPaid = 0
'      PctRPenPaid = 0
'      PctROpt1Paid = 0
'      PctROpt2Paid = 0
'      PctROpt3Paid = 0
'    End If
'    '
'    Print #RptHandle, TownName; dlm; TaxYear(x); dlm; YrTaxChrgs(x); dlm; YrTaxPaid(x); dlm; PctPaid; dlm;
'    '                     5               6             7             8                   9
'    Print #RptHandle, GTotChrgs; dlm; GTotPaid; dlm; GTotPct; dlm; BillType$; dlm; YrPrincChrgs(x); dlm;
'    '                       10                     11                  12                     13
'    Print #RptHandle, YrPrincPaid(x); dlm; PctPrincPaid; dlm; YrRIntChrgs(x); dlm; YrRIntPaid(x); dlm;
'    '                      14                 15                    16                   17
'    Print #RptHandle, PctRIntPaid; dlm; YrAdvChrgs(x); dlm; YrAdvPaid(x); dlm; PctAdvPaid; dlm;
'    '                         18                         19                        20                     21
'    Print #RptHandle, YrLateListChrgs(x); dlm; YrLateListPaid(x); dlm; PctLateListPaid; dlm; YrRPenChrgs(x); dlm;
'    '                        22               23                 24                    25
'    Print #RptHandle, YrRPenPaid(x); dlm; PctRPenPaid; dlm; YrROpt1Chrgs(x); dlm; YrROpt1Paid(x); dlm;
'    '                       26                  27                  28                   29
'    Print #RptHandle, PctROpt1Paid; dlm; YrROpt2Chrgs(x); dlm; YrROpt2Paid(x); dlm; PctROpt2Paid; dlm;
'    '                       30                  31                      32                     33
'    Print #RptHandle, YrROPt3Chrgs(x); dlm; YrROpt3Paid(x); dlm; PctROpt3Paid; dlm; YrPersChrgs(x); dlm;
'    '                        34                   35                 36                   37                  38
'    Print #RptHandle, YrPersChrgs(x); dlm; PctPersPaid; dlm; YrMTChrgs(x); dlm; YrMTPaid(x); dlm; PctMTPaid; dlm;
'    '                      39               40                41               42               43
'    Print #RptHandle, YrMCChrgs(x); dlm; YrMCPaid(x); dlm; PctMCPaid; dlm; YrFEChrgs(x); dlm; YrFEPaid(x); dlm;
'    '                    44              45                 46                47               48
'    Print #RptHandle, PctFEPaid; dlm; YrMHChrgs(x); dlm; YrMHPaid(x); dlm; PctMHPaid; dlm; YrPIntChrgs(x); dlm;
'    '                       49                     50            51                  52                  53
'    Print #RptHandle, YrPIntPaid(x); dlm; PctPIntPaid; dlm; YrPPenChrgs(x); dlm; YrPPenPaid(x); dlm; PctPPenPaid; dlm;
'    '                      54                   55                  56                  57                    58                   59
'    Print #RptHandle, YrPOpt1Chrgs(x); dlm; YrPOpt1Paid(x); dlm; PctPOpt1Paid; dlm; YrPOpt2Chrgs(x); dlm; YrPOpt2Paid(x); dlm; PctPOpt2Paid; dlm;
'    '                       60                     61                  62           63                    64                 65
'    Print #RptHandle, YrPOpt3Chrgs(x); dlm; YrPOpt3Paid(x); dlm; PctPOpt3Paid; GTotPrincChrgs; dlm; GTotPrincPaid; dlm; GTotPrincPct; dlm;
'    '                      66                 67                 68                69                 70                71
'    Print #RptHandle, GTotRIntChrgs; dlm; GTotRIntPaid; dlm; GTotRIntPct; dlm; GTotAdvChrgs; dlm; GTotAdvPaid; dlm; GTotAdvPct; dlm;
'    '                      72                        73                   74                    75                 76                 77
'    Print #RptHandle, GTotLateListChrgs; dlm; GTotLateListPaid; dlm; GTotLateListPct; dlm; GTotRPenChrgs; dlm; GTotRPenPaid; dlm; GTotRPenPct; dlm;
'    '                      78                    79                 80                 81                   82                  83
'    Print #RptHandle, GTotROpt1Chrgs; dlm; GTotROpt1Paid; dlm; GTotROpt1Pct; dlm; GTotROpt2Chrgs; dlm; GTotROpt2Paid; dlm; GTotROpt2Pct; dlm;
'    '                      84                   85                  86                87                  88                  89
'    Print #RptHandle, GTotROpt3Chrgs; dlm; GTotROpt3Paid; dlm; GTotROpt3Pct; dlm; GTotPersChrgs; dlm; GTotPersPaid; dlm; GTotPersPct; dlm;
'    '                     90                91              92               93                94              95               96
'    Print #RptHandle, GTotMTChrgs; dlm; GTotMTPaid; dlm; GTotMTPct; dlm; GTotMCChrgs; dlm; GTotMCPaid; dlm; GTotMCPct; dlm; GTotFEChrgs; dlm;
'    '                     97               98              99               100               101              102                 103
'    Print #RptHandle, GTotFEPaid; dlm; GTotFEPct; dlm; GTotMHChrgs; dlm; GTotMHPaid; dlm; GTotMHPct; dlm; GTotPIntChrgs; dlm; GTotPIntPaid; dlm;
'    '                     104               105                 106               107                108                  109
'    Print #RptHandle, GTotPIntPct; dlm; GTotPPenChrgs; dlm; GTotPPenPaid; dlm; GTotPPenPct; dlm; GTotPOpt1Chrgs; dlm; GTotPOpt1Paid; dlm;
'    '                     110                  111                112                 113                 114
'    Print #RptHandle, GTotPOpt1Pct; dlm; GTotPOpt2Chrgs; dlm; GTotPOpt2Paid; dlm; GTotPOpt2Pct; dlm; GTotPOpt3Chrgs; dlm;
'    '                     115                 116
'    Print #RptHandle, GTotPOpt3Paid; dlm; GTotPOpt3Pct
'
'
'NextOne:
'  Next x
'
'  Close
'
'  arVATaxCollRateRptDet.Show
'
'  Return
'
'PrintText:
'
'
'  Return
'
'
'ERRORSTUFF:
'   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxCollectRateRpt", "cmdProcess_Click", Erl)
'     Case emrExitProc:
'       Resume Proc_Exit
'     Case emrResume:
'       Resume
'     Case emrResumeNext:
'       Resume Next
'     Case Else
'      '--- Technically, this should never happen.
'       Resume Proc_Exit
'   End Select
'
'Proc_Exit:
'  '--- Cleanup code goes here...
'    Close
  
End Sub

