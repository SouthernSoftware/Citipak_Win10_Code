VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxDeleteRateCode 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Optional Revenue Rate Code"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxDeleteRateCode.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpListR 
      Height          =   1728
      Left            =   2460
      TabIndex        =   7
      Top             =   2280
      Width           =   6732
      _Version        =   196608
      _ExtentX        =   11874
      _ExtentY        =   3048
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
      Columns         =   4
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
      ColDesigner     =   "frmVATaxDeleteRateCode.frx":08CA
   End
   Begin LpLib.fpList fpListP 
      Height          =   1728
      Left            =   2460
      TabIndex        =   8
      Top             =   4680
      Width           =   6732
      _Version        =   196608
      _ExtentX        =   11874
      _ExtentY        =   3048
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
      Columns         =   4
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
      ColDesigner     =   "frmVATaxDeleteRateCode.frx":0C06
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   624
      Left            =   3060
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "Press this button to exit back to the main business license menu."
      Top             =   7440
      Width           =   2376
      _Version        =   131072
      _ExtentX        =   4191
      _ExtentY        =   1101
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
      ButtonDesigner  =   "frmVATaxDeleteRateCode.frx":0F42
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDelete 
      Height          =   624
      Left            =   6144
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "Press this button to exit back to the main business license menu."
      Top             =   7440
      Width           =   2388
      _Version        =   131072
      _ExtentX        =   4212
      _ExtentY        =   1101
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
      ButtonDesigner  =   "frmVATaxDeleteRateCode.frx":1121
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Revenue"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   336
      Left            =   2520
      TabIndex        =   10
      Top             =   4440
      Width           =   1932
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Real Revenue"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   336
      Left            =   2520
      TabIndex        =   9
      Top             =   2040
      Width           =   1932
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "To Select Double-Click Item or Highlight and Click Delete."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   2280
      TabIndex        =   4
      Top             =   6960
      Width           =   7080
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Method"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   216
      Left            =   7080
      TabIndex        =   3
      Top             =   1716
      Width           =   1932
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   336
      Left            =   4680
      TabIndex        =   2
      Top             =   1716
      Width           =   2412
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Opt'l Revenue #"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   336
      Left            =   2640
      TabIndex        =   1
      Top             =   1716
      Width           =   1932
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   5256
      Left            =   2100
      Top             =   1476
      Width           =   7452
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1500
      Top             =   492
      Width           =   8652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Delete Optional Revenue Rate Codes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   3156
      TabIndex        =   0
      Top             =   660
      Width           =   5292
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1500
      Top             =   396
      Width           =   8652
   End
End
Attribute VB_Name = "frmVATaxDeleteRateCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class

Private Sub cmdDelete_Click()
  If fpListR.SelCount > 0 Then
    Call fpListR_DblClick
  ElseIf fpListP.SelCount > 0 Then
    Call fpListP_DblClick
  End If
End Sub

Private Sub cmdExit_Click()
  frmVATaxRateMenu.Show
  DoEvents
  Unload Me
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
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpDeleteAnExisting
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxDeleteRateCode.")
      Call Terminate
      End
    End If
  End If

End Sub
'Private Sub Form_Resize()
'  If Me.WindowState <> vbMinimized Then
'    Me.Visible = False
'    'Temp_Class.ResizeControls Me
'    Me.Visible = True
'    Me.SetFocus
'    DoEvents
'  End If
'End Sub

Private Sub LoadMe()
  Dim TblRec As OptRevRateTablesType
  Dim TRHandle As Integer
  Dim NumOfTRRecs As Integer
  Dim x As Integer
  Dim Method$
  
  OpenTaxRateTables TRHandle, NumOfTRRecs
  For x = 1 To NumOfTRRecs
    Get TRHandle, x, TblRec
    If TblRec.Deleted = True Then GoTo Deleted
    If TblRec.Type = "F" Then
      Method$ = "Flat Rate"
    ElseIf TblRec.Type = "S" Then
      Method$ = "Step Flat"
    ElseIf TblRec.Type = "P" Then
      Method$ = "Step Pct"
    End If
    If TblRec.RevType = "R" Then
      fpListR.InsertRow = CStr(TblRec.OptRevNum) + Chr(9) + QPTrim$(TblRec.Desc) + Chr(9) + Method + Chr(9) + CStr(x)
    ElseIf TblRec.RevType = "P" Then
      fpListP.InsertRow = CStr(TblRec.OptRevNum) + Chr(9) + QPTrim$(TblRec.Desc) + Chr(9) + Method + Chr(9) + CStr(x)
    End If
Deleted:
  Next x
  Close TRHandle

End Sub

Private Sub fpListP_Click()
  fpListR.Action = ActionDeselectAll
End Sub

Private Sub fpListR_Click()
  fpListP.Action = ActionDeselectAll
End Sub

Private Sub fpListR_DblClick()
  Dim ThisRec As Integer
  Dim TblRec As OptRevRateTablesType
  Dim TRHandle As Integer
  Dim NumOfTRRecs As Integer
  Dim RealRec As PropertyRecType
  Dim RRHandle As Integer
  Dim NumOfRRREcs As Long
  Dim x As Long
  Dim ThisRateDesc$
  
  On Error GoTo ERRORSTUFF
  
  If fpListR.ListIndex = -1 Then
    Call TaxMsg(900, "Please make a selection from the real revenue list.")
    Exit Sub
  End If
  
  If TaxMsgWOpts(900, "Are you sure you wish to delete this real rate code?", "F10 Delete", "ESC Don't Delete") = "abort" Then
    Unload frmVATaxMsgWOpts
    Close
    Exit Sub
  End If
  
  fpListR.Col = 3
  fpListR.Row = fpListR.ListIndex
  ThisRec = CInt(fpListR.ColText)
  
  OpenTaxRateTables TRHandle, NumOfTRRecs
  Get TRHandle, ThisRec, TblRec
  
  ReDim RealProp(1 To 1) As Long
  ReDim RealRev(1 To 1) As Integer
  RealCnt = 0
  OpenRealPropFile RRHandle, NumOfRRREcs
  For x = 1 To NumOfRRREcs
    Get RRHandle, x, RealRec
    If RealRec.Deleted = True Then GoTo Deleted
    If RealRec.OptRev1Chrg = ThisRec Then
      RealCnt = RealCnt + 1
      ReDim Preserve RealProp(1 To RealCnt) As Long
      ReDim Preserve RealRev(1 To RealCnt) As Integer
      RealProp(RealCnt) = x
      RealRev(RealCnt) = 1
    End If
    If RealRec.OptRev2Chrg = ThisRec Then
      RealCnt = RealCnt + 1
      ReDim Preserve RealProp(1 To RealCnt) As Long
      ReDim Preserve RealRev(1 To RealCnt) As Integer
      RealProp(RealCnt) = x
      RealRev(RealCnt) = 2
    End If
    If RealRec.OptRev3Chrg = ThisRec Then
      RealCnt = RealCnt + 1
      ReDim Preserve RealProp(1 To RealCnt) As Long
      ReDim Preserve RealRev(1 To RealCnt) As Integer
      RealProp(RealCnt) = x
      RealRev(RealCnt) = 3
    End If
Deleted:
  Next x
  Close RRHandle
  
  If RealCnt = 0 Then
    TblRec.Deleted = True
    Put TRHandle, ThisRec, TblRec
    Close TRHandle
    Call TaxMsg(900, "The real rate code has been deleted successfully.")
    Call Reload
  Else
    Close TRHandle '
    frmVATaxDeletedInstances.Show
  End If
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxDeleteRateCode", "fpListR_DblClick", Erl)
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
Private Sub fpListP_DblClick()
  Dim ThisRec As Integer
  Dim TblRec As OptRevRateTablesType
  Dim TRHandle As Integer
  Dim NumOfTRRecs As Integer
  Dim PersRec As PersonalRecType
  Dim PRHandle As Integer
  Dim NumOfPRRecs As Long
  Dim x As Long
  Dim ThisRateDesc$
  
  On Error GoTo ERRORSTUFF
  
  If fpListP.ListIndex = -1 Then
    Call TaxMsg(900, "Please make a selection from the personal revenue list.")
    Exit Sub
  End If
  
  If TaxMsgWOpts(900, "Are you sure you wish to delete this personal rate code?", "F10 Delete", "ESC Don't Delete") = "abort" Then
    Unload frmVATaxMsgWOpts
    Close
    Exit Sub
  End If
  
  fpListP.Col = 3
  fpListP.Row = fpListP.ListIndex
  ThisRec = CInt(fpListP.ColText)
  
  OpenTaxRateTables TRHandle, NumOfTRRecs
  Get TRHandle, ThisRec, TblRec
  
  ReDim PersProp(1 To 1) As Long
  ReDim PersRev(1 To 1) As Integer
  RealCnt = 0
  OpenPersPropFile PRHandle, NumOfPRRecs
  For x = 1 To NumOfPRRecs
    Get PRHandle, x, PersRec
    If PersRec.Deleted = True Then GoTo Deleted
    If PersRec.OptRev1Chrg = ThisRec Then
      PersCnt = PersCnt + 1
      ReDim Preserve PersProp(1 To PersCnt) As Long
      ReDim Preserve PersRev(1 To PersCnt) As Integer
      PersProp(PersCnt) = x
      PersRev(PersCnt) = 1
    End If
    If PersRec.OptRev2Chrg = ThisRec Then
      PersCnt = PersCnt + 1
      ReDim Preserve PersProp(1 To PersCnt) As Long
      ReDim Preserve PersRev(1 To PersCnt) As Integer
      PersProp(PersCnt) = x
      PersRev(PersCnt) = 2
    End If
    If PersRec.OptRev3Chrg = ThisRec Then
      PersCnt = PersCnt + 1
      ReDim Preserve PersProp(1 To PersCnt) As Long
      ReDim Preserve PersRev(1 To PersCnt) As Integer
      PersProp(PersCnt) = x
      PersRev(PersCnt) = 3
    End If
Deleted:
  Next x
  Close PRHandle
  
  If PersCnt = 0 Then
    TblRec.Deleted = True
    Put TRHandle, ThisRec, TblRec
    Close TRHandle
    Call TaxMsg(900, "The personal rate code has been deleted successfully.")
    Call Reload
  Else
    Close TRHandle '
    frmVATaxDeletedInstances.Show
  End If
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxDeleteRateCode", "fpListR_DblClick", Erl)
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

Private Sub Reload()
  Dim TblRec As OptRevRateTablesType
  Dim TRHandle As Integer
  Dim NumOfTRRecs As Integer
  Dim x As Integer
  Dim Method$
  
  fpListR.Clear
  fpListP.Clear
  OpenTaxRateTables TRHandle, NumOfTRRecs
  For x = 1 To NumOfTRRecs
    Get TRHandle, x, TblRec
    If TblRec.Deleted = True Then GoTo Deleted
    If TblRec.Type = "F" Then
      Method$ = "Flat Rate"
    ElseIf TblRec.Type = "S" Then
      Method$ = "Step Flat"
    ElseIf TblRec.Type = "P" Then
      Method$ = "Step Pct"
    End If
    If TblRec.RevType = "R" Then
      fpListR.InsertRow = CStr(TblRec.OptRevNum) + Chr(9) + QPTrim$(TblRec.Desc) + Chr(9) + Method + Chr(9) + CStr(x)
    ElseIf TblRec.RevType = "P" Then
      fpListP.InsertRow = CStr(TblRec.OptRevNum) + Chr(9) + QPTrim$(TblRec.Desc) + Chr(9) + Method + Chr(9) + CStr(x)
    End If
Deleted:
  Next x
  Close TRHandle

End Sub
