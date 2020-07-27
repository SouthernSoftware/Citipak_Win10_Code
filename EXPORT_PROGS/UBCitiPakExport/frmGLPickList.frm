VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmGLPickList 
   BackColor       =   &H008F8265&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "General Ledger Accounts"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   6960
   Icon            =   "frmGLPickList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin LpLib.fpList fpList1 
      Height          =   3765
      Left            =   270
      TabIndex        =   0
      Top             =   1200
      Width           =   6435
      _Version        =   196608
      _ExtentX        =   11351
      _ExtentY        =   6641
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
      Columns         =   2
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   1
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
      ColDesigner     =   "frmGLPickList.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClose 
      Height          =   495
      Left            =   3768
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5808
      Width           =   1350
      _Version        =   131072
      _ExtentX        =   2381
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmGLPickList.frx":0BAE
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   495
      Left            =   2016
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5808
      Width           =   1350
      _Version        =   131072
      _ExtentX        =   2381
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmGLPickList.frx":0DC3
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The General Ledger numbers are displayed without the fund numbers because the split accounting method is being used."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   5160
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   6588
      Left            =   48
      Top             =   48
      Width           =   6876
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "General Ledger Accounts Lookup"
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
      Left            =   1752
      TabIndex        =   1
      Top             =   480
      Width           =   3516
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   636
      Left            =   1488
      Top             =   336
      Width           =   4044
   End
End
Attribute VB_Name = "frmGLPickList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdClose_Click()
   Unload frmGLPickList
   DoEvents
End Sub

Private Sub cmdHelp_Click()
  MsgBox "You can cut and paste the correct G/L number by highlighting the desired number in the list and then double clicking on it. Next double click the field where the number should go and the number will appear there."
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdClose_Click
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%H"
      Call cmdHelp_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
   Dim JGLIdxRec(1) As JGLAcctIdxType
   Dim GLIdxNum$
   Dim GLDHandle As Integer
   Dim GLIdxRecLen As Integer
   Dim GLDescRecLen As Integer
   Dim TotalAccts As Integer
   Dim Nextx As Integer, x As Integer
   Dim GLIDATDesc$
   Dim GLDesc(1) As GLAcctRecType
   Dim GLIdxHandle As Integer
   Dim SysHandle As Integer
   Dim SysRec As RegDSysFileRecType
   Dim SysRecCnt As Integer
   Dim GLFundLen%, GLAcctLen%, GLDetLen%
   Dim DedsOpen As Boolean

   'the g/l list cannot be accessed until the system
   'record has been saved so look for where the gl .dat
   'files should be located
   'CurrCitiPath is a global that is defined when the progam first opens
   'and also redefined if a change is made while the System Interface
   'screen is open and that is why we use CurrCitiPath here because
   'this screen can be accessed while the Interface is running and before
   'a new path is saved
   On Error GoTo ERRORSTUFF

   DedsOpen = False '01/18/05
   If Exist("prdeductions.dat") Then '01/18/05
      DedsOpen = True '01/18/05
   End If '01/18/05
   Call GetAcctStruct(CurrCitiPath, GLFundLen%, GLAcctLen%, GLDetLen%) '01/18/05
'   DoEvents
   OpenSysFile SysHandle
   Get SysHandle, 1, SysRec
   SysRecCnt = LOF(SysHandle) / Len(SysRec) '7/25
   If SysRecCnt <> 0 Then
     If SysRec.SplitFlag = "Y" And DedsOpen = True Then '01/18/05
       Label2.Visible = True '01/18/05
     End If '01/18/05
   End If
'   DoEvents
'   OpenSysFile SysHandle
'   SysRecCnt = LOF(SysHandle) / Len(SysRec) '7/25
'   If SysRecCnt <> 0 Then
'     Get SysHandle, 1, SysRec
'     If CheckCitiDir(CurrCitiPath) = 0 Then
'       frmWarnOpenWNoDir.Show vbModal, Me
'       Close SysHandle
'       Exit Sub
'     End If
'   Else
'     MsgBox "System control file has not been saved"
'   End If
'   Close SysHandle
   
   If Exist(CurrCitiPath + "GLACCT.IDX") Then
     GLIdxNum$ = CurrCitiPath + "GLACCT.IDX"
   ElseIf Exist(CurrCitiPath + "\GLACCT.IDX") Then
     GLIdxNum$ = CurrCitiPath + "\GLACCT.IDX"
   Else
     MsgBox "Path to GLACCT.IDX could not be found"
     Exit Sub
   End If
   
   If Exist(CurrCitiPath + "GLACCT.DAT") Then
     GLIDATDesc$ = CurrCitiPath + "GLACCT.DAT"
   ElseIf Exist(CurrCitiPath + "\GLACCT.DAT") Then
     GLIDATDesc$ = CurrCitiPath + "\GLACCT.DAT"
   Else
     MsgBox "Path to GLACCT.DAT could not be found"
     Exit Sub
   End If
NoFileYet:
   GLIdxRecLen = Len(JGLIdxRec(1))
   GLDescRecLen = Len(GLDesc(1))
   TotalAccts = FileSize(GLIDATDesc$) \ GLDescRecLen
   
   If TotalAccts = 0 Then Exit Sub
   
   ReDim DescBuff(1 To TotalAccts)
   GLIdxHandle = FreeFile
   Open GLIdxNum$ For Random As GLIdxHandle Len = GLIdxRecLen
   For x = 1 To TotalAccts
     Get GLIdxHandle, x, JGLIdxRec(1)
     DescBuff(x) = JGLIdxRec(1).RecNo
   Next x
   Close GLIdxHandle
   GLDHandle = FreeFile
   Open GLIDATDesc$ For Random As GLDHandle Len = GLDescRecLen
   For x = 1 To TotalAccts
     If DescBuff(x) <> 0 And Not GLDesc(1).Deleted Then 'added Not GLDesc(1).Deleted on 11/14/2002
       Get GLDHandle, DescBuff(x), GLDesc(1)
       If DedsOpen = True Then
         If SysRec.SplitFlag = "Y" Then '01/18/05
           fpList1.InsertRow = QPTrim$(GLDesc(1).Title) & " " & Chr$(9) & Mid(GLDesc(1).Num, (GLFundLen + 2), Len(GLDesc(1).Num)) '01/18/05
         Else
           fpList1.InsertRow = QPTrim$(GLDesc(1).Title) & " " & Chr$(9) & QPTrim$(GLDesc(1).Num) '01/18/05
         End If
       Else
         fpList1.InsertRow = QPTrim$(GLDesc(1).Title) & " " & Chr$(9) & QPTrim$(GLDesc(1).Num) '01/18/05
       End If
     End If
   Next x
   Close GLDHandle
   fpList1.Row = 0
   fpList1.Selected = True
   
   Exit Sub
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmGLPickList", "Form Load", Erl)
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
    Unload Me
End Sub

Private Sub EditCopyProc(Text$)
   ' Copy selected text onto Clipboard.
   Clipboard.Clear
   Clipboard.SetText Text
End Sub

Private Sub fpList1_DblClick()
  Dim ThisOne$
  Clipboard.Clear

  fpList1.Row = -1
  fpList1.Col = 1
  ThisOne = fpList1.ColText
  Call EditCopyProc(ThisOne$)
  Unload frmGLPickList
End Sub



