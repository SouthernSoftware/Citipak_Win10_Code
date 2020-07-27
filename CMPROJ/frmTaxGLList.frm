VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmTaxGLList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Ledger Account List"
   ClientHeight    =   2820
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   8364
   Icon            =   "frmTaxGLList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   8364
   Begin LpLib.fpList fpList1 
      Height          =   1440
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   6075
      _Version        =   196608
      _ExtentX        =   10716
      _ExtentY        =   2540
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
      ColDesigner     =   "frmTaxGLList.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClose 
      Height          =   495
      Left            =   6720
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1110
      _Version        =   131072
      _ExtentX        =   1958
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
      ButtonDesigner  =   "frmTaxGLList.frx":0D02
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   495
      Left            =   6720
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1110
      _Version        =   131072
      _ExtentX        =   1958
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
      ButtonDesigner  =   "frmTaxGLList.frx":0EDF
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   510
      Left            =   1380
      Top             =   285
      Width           =   4050
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
      Height          =   330
      Left            =   1635
      TabIndex        =   3
      Top             =   360
      Width           =   3510
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   2505
      Left            =   120
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "frmTaxGLList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdClose_Click()
   Unload frmTaxGLList
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
   Dim IdxRec As JGLAcctIdxType
   Dim GLIdxNum$
   Dim IdxHandle As Integer
   Dim IdxCnt As Integer
   Dim x As Integer
   Dim GLRec As GLAcctRecType
   Dim GLHandle As Integer
   Dim GLCnt As Integer
   
   On Error GoTo ERRORSTUFF
   
   OpenGLIdxFile IdxHandle, IdxCnt
   
   If IdxCnt = 0 Then
     MsgBox "ERROR: No General Ledger index file could be found. General Ledger list cannot be displayed."
     Close IdxHandle
     Exit Sub
   End If
   ReDim IdxRecs(1 To IdxCnt) As Integer
   For x = 1 To IdxCnt
     Get IdxHandle, x, IdxRec
     IdxRecs(x) = IdxRec.RecNo
   Next x
   Close IdxHandle
   
   OpenGLAcctFile GLHandle, GLCnt
   If GLCnt = 0 Then
     frmTaxMsg.Label1.Caption = "ERROR: No General Ledger file could be found. The General Ledger list cannot be loaded."
     frmTaxMsg.Label1.Top = 900
     frmTaxMsg.Show vbModal
     Close GLHandle
     Exit Sub
   End If
   
   If GLCnt < IdxCnt Then
     frmTaxMsg.Label1.Caption = "ERROR: The GL index count is greater than the GL file count."
     frmTaxMsg.Label1.Top = 900
     frmTaxMsg.Show vbModal
   End If
   
   For x = 1 To IdxCnt
     If IdxRecs(x) <> 0 Then
       Get GLHandle, IdxRecs(x), GLRec
       If Not GLRec.Deleted Then
         fpList1.InsertRow = QPTrim$(GLRec.Title) & " " & Chr$(9) & QPTrim$(GLRec.Num)
       End If
     End If
   Next x
   Close GLHandle
   fpList1.Row = 0
   fpList1.Selected = True
   
   Exit Sub
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxGLList", "Form Load", Erl)
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
  fpList1.col = 1
  ThisOne = fpList1.ColText
  Call EditCopyProc(ThisOne$)
'  If frmTaxSystemSetup.Visible = True Then
'    Unload Me
'  End If
End Sub




