VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmTaxManualBillEdit 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manual Bill Edit List"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTaxManualBillEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpListInEdit 
      Height          =   3012
      Left            =   1260
      TabIndex        =   2
      Top             =   3048
      Width           =   9012
      _Version        =   196608
      _ExtentX        =   15896
      _ExtentY        =   5313
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      Columns         =   6
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
      ColDesigner     =   "frmTaxManualBillEdit.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   492
      Left            =   3828
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7116
      Width           =   1584
      _Version        =   131072
      _ExtentX        =   2794
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmTaxManualBillEdit.frx":0C26
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   492
      Left            =   6228
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7116
      Width           =   1584
      _Version        =   131072
      _ExtentX        =   2794
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmTaxManualBillEdit.frx":0E02
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Year"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   8400
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Property Class"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bill #"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   3975
      Left            =   960
      Top             =   2445
      Width           =   9615
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1500
      Top             =   1223
      Width           =   8655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Manual Tax Bill Edit List"
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
      Left            =   3150
      TabIndex        =   1
      Top             =   1388
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1500
      Top             =   1118
      Width           =   8655
   End
End
Attribute VB_Name = "frmTaxManualBillEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  
Private Sub cmdExit_Click()
  KillFile "C:\CPWork\manualedit.dat"
  frmTaxManualBillMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdProcess_Click()
  Dim One As Integer
  Dim AHandle As Integer
  
  If fpListInEdit.ListIndex = -1 Then
    Call TaxMsg(900, "Please make a selection from the list.")
    Exit Sub
  End If
  One = 1
  AHandle = FreeFile
  Open "C:\CPWork\manualedit.dat" For Output As AHandle
  Print #AHandle, One
  Close AHandle
  fpListInEdit.Row = fpListInEdit.ListIndex
  fpListInEdit.Col = 4
  ThisMRec = CInt(fpListInEdit.ColText)
  fpListInEdit.Col = 5
  GCustNum = CLng(fpListInEdit.ColText)
  frmTaxManualBillEntry.Show
  frmTaxManualBillEntry.PostSaveLoad = True
  Call frmTaxManualBillEntry.EnterEditCheck
'  frmTaxManualBillEntry.fpListInEdit.ListIndex = ThisMRec - 1
  DoEvents
  frmTaxManualBillEntry.PostSaveLoad = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdExit_Click
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
  Me.HelpContextID = hlpEditTransaction
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxManualBillEdit.")
      KillFile "C:\CPWork\manualedit.dat"
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
  Dim TaxMRec As TaxMTransactionType
  Dim TMHandle As Integer
  Dim NumOfTMRecs As Integer
  Dim x As Integer
  Dim ThisClass$
  
  OpenTaxManualBillFile TMHandle, NumOfTMRecs
  For x = 1 To NumOfTMRecs
    Get TMHandle, x, TaxMRec
    If TaxMRec.Deleted = True Then GoTo Deleted
    If TaxMRec.Class = "P" Then
      ThisClass = "PERSONAL"
    ElseIf TaxMRec.Class = "R" Then
      ThisClass = "REAL"
    ElseIf TaxMRec.Class = "M" Then
      ThisClass = "MOCK"
    Else
      ThisClass = "UNKNOWN"
    End If
    fpListInEdit.AddItem CStr(TaxMRec.BillNum) + Chr(9) + ThisClass + Chr(9) + QPTrim$(TaxMRec.TName) + Chr(9) + CStr(TaxMRec.TaxYear) + Chr(9) + CStr(x) + Chr(9) + CStr(TaxMRec.Account)
Deleted:
  Next x
  
  Close TMHandle
  
  fpListInEdit.ListIndex = 0
End Sub

Private Sub fpListInEdit_DblClick()
  Call cmdProcess_Click
End Sub

Public Sub ClearAndUpdateList()
  Dim TaxMRec As TaxMTransactionType
  Dim TMHandle As Integer
  Dim NumOfTMRecs As Integer
  Dim x As Integer
  Dim ThisClass$
  Dim ThisIdx As Integer
  
  ThisIdx = fpListInEdit.ListIndex
  fpListInEdit.Clear
  
  OpenTaxManualBillFile TMHandle, NumOfTMRecs
  For x = 1 To NumOfTMRecs
    Get TMHandle, x, TaxMRec
    If TaxMRec.Deleted = True Then GoTo Deleted
    If TaxMRec.Class = "P" Then
      ThisClass = "PERSONAL"
    ElseIf TaxMRec.Class = "R" Then
      ThisClass = "REAL"
    ElseIf TaxMRec.Class = "M" Then
      ThisClass = "MOCK"
    Else
      ThisClass = "UNKNOWN"
    End If
    fpListInEdit.AddItem CStr(TaxMRec.BillNum) + Chr(9) + ThisClass + Chr(9) + QPTrim$(TaxMRec.TName) + Chr(9) + CStr(TaxMRec.TaxYear) + Chr(9) + CStr(x) + Chr(9) + CStr(TaxMRec.Account)
Deleted:
  Next x
  
  Close TMHandle
  
  fpListInEdit.ListIndex = ThisIdx
  
End Sub
  

