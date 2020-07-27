VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLCategoryList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Category List"
   ClientHeight    =   6990
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   8640
   Icon            =   "frmBLCategoryList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin LpLib.fpList fpList1 
      Height          =   3765
      Left            =   825
      TabIndex        =   0
      Top             =   1650
      Width           =   6930
      _Version        =   196608
      _ExtentX        =   12224
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
      ColDesigner     =   "frmBLCategoryList.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   492
      Left            =   1386
      TabIndex        =   2
      Top             =   5808
      Width           =   2460
      _Version        =   131072
      _ExtentX        =   4339
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
      ButtonDesigner  =   "frmBLCategoryList.frx":0C92
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClose 
      Height          =   492
      Left            =   4794
      TabIndex        =   3
      Top             =   5808
      Width           =   2460
      _Version        =   131072
      _ExtentX        =   4339
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
      ButtonDesigner  =   "frmBLCategoryList.frx":0EAD
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmBLCategoryList.frx":10C3
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Left            =   624
      TabIndex        =   4
      Top             =   960
      Width           =   7356
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   720
      X2              =   1296
      Y1              =   1488
      Y2              =   2688
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   6588
      Left            =   144
      Top             =   204
      Width           =   8364
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Category List"
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
      Left            =   2424
      TabIndex        =   1
      Top             =   642
      Width           =   3900
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   636
      Left            =   2298
      Top             =   492
      Width           =   4044
   End
End
Attribute VB_Name = "frmBLCategoryList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class
Private Sub cmdClose_Click()
  
  Unload frmBLCategoryList
  If Exist("advanceltrprint.dat") Then
    frmBLPrintAdvanceLetter.fptxtCatCode.SetFocus
  End If
  If Exist("custappIssue.dat") Then
    frmBLAppListIssue.fptxtCatCode.SetFocus
  End If
  If Exist("custappsRenews.dat") Then
    frmBLPrintAppsRenwls.fptxtCatCode.SetFocus
  End If
  If Exist("custquickList.dat") Then
    frmBLQuickList.fptxtCatCode.SetFocus
  End If
  If Exist("custappList.dat") Then
    frmBLAppListing.fptxtCatCode.SetFocus
  End If
  If Exist("custXlicList.dat") Then
    frmBLXLicList.fptxtCatCode.SetFocus
  End If
  If Exist("custlicList.dat") Then
    frmBLLicListRpt.fptxtCatCode.SetFocus
  End If
  If Exist("custlistRpt.dat") Then
    frmBLCustListRpt.fptxtCatCode.SetFocus
  End If
  If Exist("custbalList.dat") Then
    frmBLCustBalListing.fptxtCatCode.SetFocus
  End If
  If Exist("categoryedit.dat") Then
    If frmBLCatEdit.fptxtCatCode.Enabled = True Then
      frmBLCatEdit.fptxtCatCode.SetFocus
    End If
  End If
  If Exist("inoutrpt.dat") Then
    frmBLInOutRpt.fptxtCatCode.SetFocus
  End If
  If Exist("custByCat.dat") Then
    frmBLCustByCat.fptxtCatCode.SetFocus
  End If
  
  KillFile "catlistopen.dat"
  
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "F1 &Turn Help Off"
    Label2.Visible = True
    Line1.Visible = True
  Else
    cmdHelp.Text = "F1 &Turn Help On"
    Label2.Visible = False
    Line1.Visible = False
  End If
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
      SendKeys "%C"
      Call cmdClose_Click
      KeyCode = 0
    Case vbKeyF1:
      SendKeys "%T"
      Call cmdHelp_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub LoadMe()
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeIdxRec As CatCodeIdxType
  Dim CodeIdxHandle As Integer
  Dim CodeIdxRecNum As Integer
  Dim CHandle As Integer
  Dim TotalAccts As Integer
  Dim x As Integer
  Dim CatCodeCnt As Integer
  Dim Nextx As Integer
  
  On Error GoTo ERRORSTUFF
  
  Label2.Visible = False
  Line1.Visible = False
 
  If Not Exist("arcatcodeidx.dat") Then 'no file there
    frmBLMessageBoxJr.Label1.Caption = "No Category Code Index has been saved."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  OpenCatCodeIdxFile CodeIdxHandle
  CodeIdxRecNum = LOF(CodeIdxHandle) \ Len(CodeIdxRec)
  If CodeIdxRecNum = 0 Then 'file is there but there is nothing in it
    frmBLMessageBoxJr.Label1.Caption = "No Category Codes in index."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  ReDim CodeIdx(1 To CodeIdxRecNum) As Integer
  For x = 1 To CodeIdxRecNum
    Get CodeIdxHandle, x, CodeIdxRec
    CodeIdx(x) = CodeIdxRec.CatCodeRec 'load array with record pointers
  Next x
  Close CodeIdxHandle
  
  OpenCatCodeFile CHandle
  CatCodeCnt = LOF(CHandle) / Len(CodeRec)
  
  If CatCodeCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No Category Codes on file."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  If Exist("custbalList.dat") Then
    fpList1.InsertRow = "ALL" & " " & Chr$(9) & "INCLUDE ALL AMOUNTS"
    fpList1.InsertRow = "PENALTIES" & " " & Chr$(9) & "PENALTY AMOUNTS ONLY"
    fpList1.InsertRow = "ISSUANCE" & " " & Chr$(9) & "ISSUANCE FEE AMOUNTS ONLY"
   End If
  
  If Exist("custlistRpt.dat") Or Exist("custlicList.dat") Or Exist("custXlicList.dat") Or Exist("custappList.dat") _
  Or Exist("custquickList.dat") Or Exist("custappsRenews.dat") Or Exist("custappIssue.dat") Or Exist("advanceltrprint.dat") _
  Or Exist("inoutrpt.dat") Or Exist("custByCat.dat") Then
    fpList1.InsertRow = "ALL" & " " & Chr$(9) & "INCLUDE ALL CATEGORY CODES"
  End If
  
  For x = 1 To CodeIdxRecNum
    Get CHandle, CodeIdx(x), CodeRec
    If Len(QPTrim(CodeRec.CatCode)) = 0 Then GoTo BadCode
    fpList1.InsertRow = QPTrim$(CodeRec.CatCode) & " " & Chr$(9) & QPTrim$(CodeRec.CODEDESC)
BadCode:
  Nextx = Nextx + 1
  Next x
  Close CHandle
  fpList1.Row = 0
  fpList1.Selected = True 'set focus to first line
ZeroText:
  Exit Sub
   

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCategoryList", "LoadMe", Erl)
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
    Unload Me
End Sub

Private Sub fpList1_DblClick()
  Dim CodeRec As ARNewCatCodeRecType
  Dim CHandle As Integer
  Dim TotalAccts As Integer
  Dim x As Integer
  Dim DESC$
  Dim Code$
  Dim Found As Boolean
  Dim One As Integer
  Dim DHandle As Integer
  Dim Thisx As Integer
  
  On Error GoTo ERRORSTUFF
  
  One = 1
  DHandle = FreeFile
  Open "catlistopen.dat" For Output As DHandle Len = 2
  Print #DHandle, One
  Close DHandle
  
  frmBLCategoryList.Hide
   
  fpList1.Col = 0 'assign variables from the user selected row
  fpList1.Row = -1
  Code$ = QPTrim$(fpList1.ColText)
  
  If Exist("categoryedit.dat") Then
    Call frmBLCatEdit.cmdExit_Click
    If ItemChangeFlag = True Then
      ItemChangeFlag = False
      Unload frmBLCategoryList
      Exit Sub
    End If
  End If
GCatNumIsZero:

'  fpList1.Col = 0 'assign variables from the user selected row
'  fpList1.Row = -1
'  Code$ = QPTrim$(fpList1.ColText)
  
  If Exist("custappList.dat") Or Exist("custappIssue.dat") Or Exist("custlistRpt.dat") Or Exist("custbalList.dat") Or Exist("custlicList.dat") Or Exist("custXlicList.dat") _
  Or Exist("custquickList.dat") Or Exist("custappsRenews.dat") Or Exist("advanceltrprint.dat") Or Exist("inoutrpt.dat") Or Exist("custByCat.dat") Then
    Found = True
    GoTo UserWantsAll
  End If
  
  fpList1.Col = 1
  DESC$ = QPTrim$(fpList1.ColText)
  
  OpenCatCodeFile CHandle
  TotalAccts = LOF(CHandle) \ Len(CodeRec)
  
  If TotalAccts = 0 Then Exit Sub
  
  For x = 1 To TotalAccts
    Get CHandle, x, CodeRec
    If Code$ = QPTrim$(CodeRec.CatCode) And DESC$ = QPTrim$(CodeRec.CODEDESC) Then 'match the selected
    'row with the right code
      Found = True
      fpList1.Row = -1
      GCatNum = x 'now you can assign the correct global
      Exit For
    Else
      Found = False
      GoTo NotAMatch
    End If
      
NotAMatch:
   Next x
  Close CHandle
  
UserWantsAll:

  If Exist("advanceltrprint.dat") Then
    If Found = True Then
      frmBLPrintAdvanceLetter.fptxtCatCode.Text = Code
      Unload frmBLCategoryList
      Exit Sub
    Else
      frmBLMessageBoxJr.Label1.Caption = "No match found."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Exit Sub
    End If
  End If
  
  If Exist("custappIssue.dat") Then
    If Found = True Then
      frmBLAppListIssue.fptxtCatCode.Text = Code
      Unload frmBLCategoryList
      Exit Sub
    Else
      frmBLMessageBoxJr.Label1.Caption = "No match found."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Exit Sub
    End If
  End If
  
  If Exist("custappsRenews.dat") Then
    If Found = True Then
      frmBLPrintAppsRenwls.fptxtCatCode.Text = Code
      Unload frmBLCategoryList
      Exit Sub
    Else
      frmBLMessageBoxJr.Label1.Caption = "No match found."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Exit Sub
    End If
  End If
  
  If Exist("custquickList.dat") Then
    If Found = True Then
      frmBLQuickList.fptxtCatCode.Text = Code
      Unload frmBLCategoryList
      Exit Sub
    Else
      frmBLMessageBoxJr.Label1.Caption = "No match found."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Exit Sub
    End If
  End If

  If Exist("custappList.dat") Then
    If Found = True Then
      frmBLAppListing.fptxtCatCode.Text = Code
      Unload frmBLCategoryList
      Exit Sub
    Else
      frmBLMessageBoxJr.Label1.Caption = "No match found."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Exit Sub
    End If
  End If
  
  If Exist("custXlicList.dat") Then
    If Found = True Then
      frmBLXLicList.fptxtCatCode.Text = Code
      Unload frmBLCategoryList
      Exit Sub
    Else
      frmBLMessageBoxJr.Label1.Caption = "No match found."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Exit Sub
    End If
  End If
  
  If Exist("custlicList.dat") Then
    If Found = True Then
      frmBLLicListRpt.fptxtCatCode.Text = Code
      Unload frmBLCategoryList
      Exit Sub
    Else
      frmBLMessageBoxJr.Label1.Caption = "No match found."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Exit Sub
    End If
  End If
  
  If Exist("custlistRpt.dat") Then
    If Found = True Then
      frmBLCustListRpt.fptxtCatCode.Text = Code
      Unload frmBLCategoryList
      Exit Sub
    Else
      frmBLMessageBoxJr.Label1.Caption = "No match found."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Exit Sub
    End If
  End If
  
  If Exist("custbalList.dat") Then
    If Found = True Then
      frmBLCustBalListing.fptxtCatCode.Text = Code
      Unload frmBLCategoryList
      Exit Sub
    Else
      frmBLMessageBoxJr.Label1.Caption = "No match found."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Exit Sub
    End If
  End If
  
  If Exist("inoutrpt.dat") Then
    If Found = True Then
      frmBLInOutRpt.fptxtCatCode.Text = Code
      Unload frmBLCategoryList
      Exit Sub
    Else
      frmBLMessageBoxJr.Label1.Caption = "No match found."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Exit Sub
    End If
  End If
  
  If Exist("custByCat.dat") Then
    If Found = True Then
      frmBLCustByCat.fptxtCatCode.Text = Code
      Unload frmBLCategoryList
      Exit Sub
    Else
      frmBLMessageBoxJr.Label1.Caption = "No match found."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Exit Sub
    End If
  End If
  
  If Exist("categoryedit.dat") Then
    If Found = True Then
      Call frmBLCatEdit.LoadMe
      Unload frmBLCategoryList
      KillFile "catlistopen.dat"
      Exit Sub
    Else
      frmBLMessageBoxJr.Label1.Caption = "No match found."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      KillFile "catlistopen.dat"
      Exit Sub
    End If
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCategoryList", "fpList1_DblClick", Erl)
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
  

