VERSION 5.00
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmFATagList 
   BackColor       =   &H008F8265&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Tag List"
   ClientHeight    =   6672
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   6960
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6672
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin LpLib.fpList fpList1 
      Height          =   3504
      Left            =   1464
      TabIndex        =   0
      Top             =   1224
      Width           =   4044
      _Version        =   196608
      _ExtentX        =   7133
      _ExtentY        =   6181
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
      Columns         =   1
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
      ColDesigner     =   "frmFATagList.frx":0000
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "F10 &Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3762
      TabIndex        =   2
      Top             =   5802
      Width           =   1356
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "F5 &Help"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2010
      TabIndex        =   1
      Top             =   5802
      Width           =   1356
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   636
      Left            =   1482
      Top             =   330
      Width           =   4044
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Tag Numbers Lookup"
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
      Left            =   1602
      TabIndex        =   3
      Top             =   474
      Width           =   3900
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   6588
      Left            =   42
      Top             =   42
      Width           =   6876
   End
End
Attribute VB_Name = "frmFATagList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Over As clsFATextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdClose_Click()
   Unload frmFATagList
   DoEvents
End Sub

Private Sub cmdHelp_Click()
  MsgBox "You can cut and paste a Tag Number number by highlighting the desired number in the list and then double clicking on it. Next double click the field where the number should go and the number will appear there. You will need to modify this number to a new one."
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%C"
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%H"
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
   Dim TagRec As FAItemRecType
   Dim THandle As Integer
   Dim TotalAccts As Integer
   Dim X As Integer
   Dim DoWhatFlag As BadFACodeNumOption
   Dim n As Integer
   Dim Nextx As Integer
   Dim Y As Integer
   Dim ThisText$
   Dim BigNum As Double
   Dim SmallNum As Double
   Dim TempNum As String
   Dim ThisX As Integer
   Dim ThisNum As Double
   Dim ValidRecs As Long
   Dim ThisTag As String
'   On Error GoTo ERRORSTUFF
   
   If Not Exist("FAITEMS.DAT") Then
     MsgBox "Path to FAITEMS.DAT could not be found"
     Exit Sub
   End If

   OpenFAItemFile THandle
   TotalAccts = LOF(THandle) \ Len(TagRec)
   If TotalAccts = 0 Then Exit Sub
   
   ReDim TagIdx(1 To TotalAccts) As Double
   BigNum = 0
   For X = 1 To TotalAccts
     Get THandle, X, TagRec
     ThisTag = QPTrim$(TagRec.ITEMTAG)
     If ThisTag <> "" Then
       TagIdx(X) = CDbl(ThisTag)
       ValidRecs = ValidRecs + 1
       If CDbl(TagIdx(X)) > BigNum Then
         BigNum = TagIdx(X)
       End If
     End If
   Next X
   
   BigNum = BigNum + 1
   SmallNum = BigNum
   Nextx = 1
   Do
     For X = Nextx To ValidRecs
       ThisNum = TagIdx(X)
       If ThisNum <= 0 Then Stop
       If ThisNum < SmallNum Then
         SmallNum = ThisNum
         ThisX = X
       End If
NoNumber:
     Next X
     TagIdx(ThisX) = TagIdx(Nextx)
     TagIdx(Nextx) = SmallNum
     Nextx = Nextx + 1
     If Nextx = ValidRecs Then Exit Do
     SmallNum = BigNum
   Loop
   Close THandle
   
   For X = 1 To ValidRecs
     fpList1.InsertRow = TagIdx(X)
   Next X
   Close THandle
   fpList1.Row = 0
   fpList1.Selected = True
ZeroText:
   Exit Sub
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFATagList", "Form Load", Erl)
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
  fpList1.Col = 0
  ThisOne = QPTrim$(fpList1.ColText)
  Call EditCopyProc(ThisOne$)
  Unload frmFATagList
End Sub






