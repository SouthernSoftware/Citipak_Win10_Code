VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmCMDispList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   1890
   ClientWidth     =   12210
   Icon            =   "frmCMDispList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpCMList 
      Height          =   3660
      Left            =   1575
      TabIndex        =   0
      Top             =   2520
      Width           =   9075
      _Version        =   196608
      _ExtentX        =   16007
      _ExtentY        =   6456
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
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
      Columns         =   2
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
      SelMax          =   1
      AutoSearch      =   1
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   0
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
      ScrollBarV      =   0
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   2
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   0   'False
      DataAutoSizeCols=   0
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   3
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
      ColDesigner     =   "frmCMDispList.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdOK 
      Height          =   480
      Left            =   7830
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6720
      Width           =   1290
      _Version        =   131072
      _ExtentX        =   2275
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmCMDispList.frx":0BAE
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdExit 
      Height          =   480
      Left            =   9195
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6720
      Width           =   1275
      _Version        =   131072
      _ExtentX        =   2249
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmCMDispList.frx":0D87
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "7:57 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "2/4/2020"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H8000000E&
      BorderWidth     =   3
      Height          =   5484
      Left            =   1260
      Top             =   1920
      Width           =   9708
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3228
      Top             =   744
      Width           =   5772
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Management Payment List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3288
      TabIndex        =   9
      Top             =   984
      Width           =   5652
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Source"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   9480
      TabIndex        =   8
      Top             =   2208
      Width           =   1044
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   8136
      TabIndex        =   7
      Top             =   2208
      Width           =   1044
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "TR Amt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   6600
      TabIndex        =   6
      Top             =   2208
      Width           =   1044
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   4056
      TabIndex        =   5
      Top             =   2208
      Width           =   1332
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Item or Highlight and Double-Click for Details."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   1728
      TabIndex        =   4
      Top             =   6840
      Width           =   5604
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Trans Date "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   1776
      TabIndex        =   3
      Top             =   2208
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3228
      Top             =   624
      Width           =   5772
   End
End
Attribute VB_Name = "frmCMDispList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim BeenDone As Boolean
Dim RCnt As Integer, NumofRevs As Integer
Dim RevText$(1 To MaxRevsCnt)
Dim Metered(1 To MaxRevsCnt) As Boolean
Dim fromform As Form, toform As Form, codeopt As Integer, WhattoDo As Integer
Public Sub Wheretogo(xfrm As Form, tfrm As Form, Optional opt As Integer, Optional DelOpt As Integer)
  
  Set fromform = xfrm
  Set toform = tfrm
  If opt <> 0 Then
    codeopt = opt
  Else
    codeopt = 0
  End If
  If DelOpt <> 0 Then
    WhattoDo = DelOpt
  Else
    WhattoDo = 0
  End If
   Me.fpCMList.ListIndex = 0
End Sub

Private Sub fpCmdExit_Click()
  SearchRec = 0
  BeenDone = False
  Me.fpCMList.Clear
  Unload frmCMDispList
End Sub
Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TownName$
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub Form_Activate()
  SearchRec& = 0
  If Not BeenDone Then
    BeenDone = True
    Me.fpCMList.ListIndex = 0
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      KeyCode = 0
      Call fpCmdExit_Click
    Case vbKeyReturn
      KeyCode = 0
      DoEvents
      Call fpCMList_DblClick  'fpCmdOK_Click
    Case Else:
  End Select
End Sub

Private Sub fpCmdOk_Click()
  If fpCMList.SelCount > 0 Then
    Call fpCMList_DblClick
  End If
End Sub
Private Sub fpCMList_DblClick()
  Dim WhatRec As Long
  fpCMList.col = 1                       'switch to the hidden RecNo. column
  WhatRec = Val(fpCMList.ColText)     'get customer recno
  If WhatRec > 0 Then
    If WhattoDo = 0 Then
      frmReportOpt.Show 1
      DeActivateControls Me
      If rptopt = 1 Then
        'do the graphics
       PrintJournal WhatRec, 1
      ElseIf rptopt = 2 Then
        'do the text
       PrintJournal WhatRec, 2
       ActivateControls Me
      Else
        ActivateControls Me
      End If
    ElseIf WhattoDo = 1 Then
      'do the void
        DeActivateControls Me
        frmInfo.Label1 = "Loading. . ."
        frmInfo.Show
        DoEvents
      'here
        toform.fpTransRecNo = QPTrim$(Str$(WhatRec))   'set hidden recno field on edit form
        toform.Wheretogo fromform, toform, 2 'send code 1 for search screen
        Load toform
        toform.Show
        DoEvents
        Unload frmInfo
        Unload Me
    Else
      'not anything yet
    End If
  Else
  '  msgstuff
  End If
  DoEvents
  'preload stuff here
  'frmTRDetail.Show vbModal

'  Unload frmTRDispList
End Sub

Private Sub fpCMList_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      KeyCode = 0
      Call fpCmdExit_Click
    Case vbKeyReturn
      KeyCode = 0
      DoEvents
      Call fpCMList_DblClick  'fpCmdOK_Click
    Case vbKeyTab
      KeyCode = 0
      DoEvents
      Call fpCmdExit_Click
    Case Else:
  End Select
End Sub

Private Sub PrintJournal(CMRecNo As Long, rptopt As Integer)
  Dim BegDate As Integer, EndDate As Integer, FromDate As String
  Dim ThruDate As String, RecSource As String, OperatorNumber As Integer
  Dim ReportFile As String, Fmt1 As String
  Dim Fmt3 As String, Fmt4 As String
  Dim CMTrRecLen As Integer, TRHandle As Integer, TrNumRecs As Long
  Dim Max As Long, Size As Long, Start As Integer, sDir As Integer
  Dim SSize As Integer, MOff As Integer, MSize As Integer, RptHandle As Integer
  Dim NumOfMiscRecs As Long, cnt As Long, RptType As Integer
  Dim TRType As String, TPaid As Double
  Dim TxRev As Double, TRev As Integer
  Dim TotalAmount As Double, Change As Double
  Dim PrintMiscFlag As Integer, MCnt As Integer
  Dim MiscRevAmt As Double, NumofRevs As Integer, RCnt As Integer
  Dim PrintUtilFlag As Integer, PrintTaxFlag As Integer, Header As String
  Dim Page As Integer, BegRecNo As Long, TransDate As Integer
  Dim UBSetupLen As Integer, TxPOpt1 As String, TxPOpt2 As String, TxPOpt3 As String
  Dim TxOpt1 As String, TxOpt2 As String, TxOpt3 As String, TMHandle As Integer
  Dim RevCnt As Integer, OutOfOrder As Boolean, x As Integer
  Dim Temp2 As Integer, uCnt As Integer, dcnt As Integer, VCCode As String, VCDesc As String
  Dim TCnt As Long, PrnOpr As String, Vlook As Long, VCat As Integer
'  ReDim RevName$(10), TotalMiscRec$(200), TotalMiscDesc$(200), TotalMiscAmt#(200), MiscCodeGL$(200)
'  ReDim TotalUtilRevAmt#(15)
'  ReDim TotalDepRevAmt#(15)
  'ReDim RevText$(15)
  Dim MCFile As Integer
  ReDim UBSetUpRec(1) As UBSetupRecType
'  ReDim DistArray(1 To 1) As DistArrayType

  ReportFile$ = UBPath$ + "CMJOURNL.PRN"  'Report File Name
  Fmt1$ = "###,###.##"
  Fmt3$ = "$#,###,###.##"
  Fmt4$ = "$###,###,###.##"

  FF$ = Chr$(12)
 ' If RptType = 0 Then
    MaxLines = 53
 ' Else
 '   MaxLines = 48
 ' End If
  LineCnt = 0
  ReDim CMTrRec(1) As CMTransRecType            ' open transaction file
  CMTrRecLen = Len(CMTrRec(1))
  TRHandle = FreeFile
  Open UBPath$ + "CMTRANS.DAT" For Random Access Read Write Shared As TRHandle Len = CMTrRecLen
  TrNumRecs& = LOF(TRHandle) \ CMTrRecLen

  Max& = TrNumRecs& '(FRE(-1) - 16000) \ 16
  Size = Max&

  Start = 1     'start at array element 1
  sDir = 0       'sort direction - use anything else for descending
  SSize = 16    'total size of each TYPE element
  MOff = 0      'offset into the TYPE for the key element
  MSize = 16    'size of the key element - coded as follows:

  '   -1 = integer
  '   -2 = long integer
  '   -3 = single precision
  '   -4 = double precision
  '   +N = TYPE array/fixed-length string of length N
    Select Case intHasTaxes
    Case 1 'NC Taxes
        ReDim TaxMasterRec(1) As TaxMasterType
        OpenTaxSetUpFile TMHandle
        Get TMHandle, 1, TaxMasterRec(1)
        Close TMHandle
        TxOpt1 = QPTrim$(TaxMasterRec(1).OptRev1)
        TxOpt2 = QPTrim$(TaxMasterRec(1).OptRev2)
        TxOpt3 = QPTrim$(TaxMasterRec(1).OptRev3)
    Case 2 'VA Taxes
        ReDim TaxMaster(1) As VATaxMasterType
        OpenVATaxSetUpFile TMHandle
        Get TMHandle, 1, TaxMaster(1)
        Close TMHandle
        TxOpt1 = QPTrim$(TaxMaster(1).OptRev1)
        TxOpt2 = QPTrim$(TaxMaster(1).OptRev2)
        TxOpt3 = QPTrim$(TaxMaster(1).OptRev3)
        TxPOpt1 = QPTrim$(TaxMaster(1).POptRev1)
        TxPOpt2 = QPTrim$(TaxMaster(1).POptRev2)
        TxPOpt3 = QPTrim$(TaxMaster(1).POptRev3)
    Case Else
    
    End Select
    
'  If Exist("TAXSETUP.DAT") Then
'    If Exist(UBPath$ + "VACitiTax.EXE") Then
'      ReDim TaxMaster(1) As VATaxMasterType
'      OpenVATaxSetUpFile TMHandle
'      Get TMHandle, 1, TaxMaster(1)
'      Close TMHandle
'      TxOpt1 = QPTrim$(TaxMaster(1).OptRev1)
'      TxOpt2 = QPTrim$(TaxMaster(1).OptRev2)
'      TxOpt3 = QPTrim$(TaxMaster(1).OptRev3)
'      TxPOpt1 = QPTrim$(TaxMaster(1).POptRev1)
'      TxPOpt2 = QPTrim$(TaxMaster(1).POptRev2)
'      TxPOpt3 = QPTrim$(TaxMaster(1).POptRev3)
'    Else
'      ReDim TaxMasterRec(1) As TaxMasterType
'      OpenTaxSetUpFile TMHandle
'      Get TMHandle, 1, TaxMasterRec(1)
'      Close TMHandle
'      TxOpt1 = QPTrim$(TaxMasterRec(1).OptRev1)
'      TxOpt2 = QPTrim$(TaxMasterRec(1).OptRev2)
'      TxOpt3 = QPTrim$(TaxMasterRec(1).OptRev3)
'    End If
'  End If
  
  ReDim Array1(1 To Size) As struct
  RptType = rptopt
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  MCFile = FreeFile
  OpenMiscCodeFile NumOfMiscRecs     ' opens misc code file
  ReDim MiscCodeRec(1) As MiscCodeRecType
  If RptType = 2 Then
    Print #RptHandle, Chr$(27); Chr$(58);         ' oki 320 12 cpi
  End If
  GoSub PrintRptHeader

    Get TRHandle, CMRecNo, CMTrRec(1)
        TRType$ = ""
        Select Case CMTrRec(1).TransSource
        Case 1
          TRType$ = "Miscellaneous"
        Case 201
          TRType$ = "VOID Miscellaneous"
        Case 27
          TRType$ = "Utility Deposit"
        Case 24
          TRType$ = "Utility Billing"
        Case 227
          TRType$ = "VOID Utility Deposit"
        Case 224
          TRType$ = "VOID Utility Billing"
        Case 30 To 39, 131
          TRType$ = "Tax Billing"
        Case 161
          TRType$ = "Real Tax Billing"
        Case 171
          TRType$ = "Pers Tax Billing"
        Case 261
          TRType$ = "VOID Real Tax Billing"
        Case 271
          TRType$ = "VOID Pers Tax Billing"
        Case 40 To 49, 141
          TRType$ = "Business License"
        Case 231
          TRType$ = "VOID Tax Billing"
        Case 241
          TRType$ = "VOID Business License"
        Case 50 To 59, 151
          TRType$ = "Decal Paymemt"
        Case 251
          TRType$ = "VOID Decal"
        End Select
          Page = Page + 1
          
     '####################################
      Print #RptHandle, " "
      Print #RptHandle, " "
      Print #RptHandle, Tab(27); "Cash Receipts Detail Report"
      Print #RptHandle, Tab(7); "Date: "; Date$; Tab(70); "Page: "; 1
      Print #RptHandle, Tab(7); " Current Operator: "; OperNum
      Print #RptHandle, String$(80, "-")
      Print #RptHandle, Tab(20); "  Transaction Date: "; Num2Date(CMTrRec(1).TransDate)
      Print #RptHandle, Tab(20); "     Customer Name: "; Left$(CMTrRec(1).TransName, 18)
      Print #RptHandle, Tab(20); "    Account Number: "; CMTrRec(1).TransAcctNum
      Print #RptHandle, Tab(20); "  Transaction Type: "; TRType$
      Print #RptHandle, Tab(20); "    Receipt Number: "; Str$(CMRecNo)
      If CMTrRec(1).TransVoidNum > 0 And CMTrRec(1).TransSource > 200 Then
        Print #RptHandle, Tab(20); " Void For Receipt#: "; CMTrRec(1).TransVoidNum
      ElseIf CMTrRec(1).TransVoidNum > 0 And CMTrRec(1).TransSource < 200 Then
        Print #RptHandle, Tab(20); "Voided By Receipt#: "; CMTrRec(1).TransVoidNum
      End If
      Print #RptHandle, Tab(20); "  Payment Operator: "; CMTrRec(1).TransOperNum
      Print #RptHandle, Tab(20); "       Description: "; QPTrim$(CMTrRec(1).TransDesc)
      Print #RptHandle, Tab(20); "Transaction Amount: "; Using(Fmt1$, CMTrRec(1).TransAmount)
      Print #RptHandle, " "
      Print #RptHandle, Tab(10); "-------------------Payment Amounts-------------------- "
       
          Print #RptHandle, Tab(15); "    Cash: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransCash)
          If CMTrRec(1).TransTender = 4 Then
            Print #RptHandle, Tab(15); "   Check: "; Tab(40); Using(Fmt1$, 0#)
            Print #RptHandle, Tab(15); "  Charge: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransCheck)
          Else
            Print #RptHandle, Tab(15); "   Check: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransCheck)
            Print #RptHandle, Tab(15); "  Charge: "; Tab(40); Using(Fmt1$, 0#)
         End If
          Print #RptHandle, Tab(15); "Amt Paid: "; Tab(40); Using(Fmt1$, Round(CMTrRec(1).TransCheck + CMTrRec(1).TransCash))
          Print #RptHandle, Tab(15); "Amt Owed: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransAmtOwed)
'?????????HEY LOOK HERE MADE THIS CHANGE WHY WOULDN'T WORK??????????????
        'Did each change on indiv case types
        Select Case CMTrRec(1).TransSource
          Case 1, 201
          Change# = Round#((CMTrRec(1).TransCash + CMTrRec(1).TransCheck) - CMTrRec(1).TransAmount)
        'If CHANGE# < 0 Then CHANGE# = 0
          Print #RptHandle, Tab(15); "  Change: "; Tab(40); Using(Fmt1$, Change#)
          ' Misc Code Breakdown Dist.****************
          PrintMiscFlag = 0
          Print #RptHandle, " "
          Print #RptHandle, Tab(10); "----------------Miscellaneous Code BrkDwn----------------"
          For MCnt = 1 To 5
            MiscRevAmt# = (CMTrRec(1).TransRevAmt(MCnt))
            MiscRevAmt# = Round#(MiscRevAmt#)
            If MiscRevAmt# <> 0 Then
              ' If There Is an Amount in Misc Rev 1-5 then get code record number from 6-10
              If CMTrRec(1).TransRevAmt(MCnt + 5) >= 1 Then
                Get MCFile, CMTrRec(1).TransRevAmt(MCnt + 5), MiscCodeRec(1)
                  Print #RptHandle, Tab(15); MiscCodeRec(1).MiscCode;
                  Print #RptHandle, Tab(25); QPTrim$(MiscCodeRec(1).Description);
                  Print #RptHandle, Tab(50); Using(Fmt1$, MiscRevAmt#)
                  PrintMiscFlag = 1
              End If
            End If
          Next MCnt
          '  If PrintMiscFlag = 1 Then Print #RptHandle, String$(80, "-"): Linecnt = Linecnt + 1
          'End Misc Code Print ********************************
        'End If
        Case 20 To 29, 224, 227
          TPaid# = 0
          For TRev = 1 To 15
            TPaid# = Round#(TPaid# + CMTrRec(1).TransRevAmt(TRev))
          Next
          Change# = Round#((CMTrRec(1).TransCheck + CMTrRec(1).TransCash) - TPaid#)
          'If CHANGE# < 0 Then CHANGE# = 0
          Print #RptHandle, Tab(15); "  Change: "; Tab(40); Using(Fmt1$, Change#)
   'If CMTrRec(1).TransSource >= 20 And CMTrRec(1).TransSource <= 29 Then
          If CMTrRec(1).TransSource <> 27 Or CMTrRec(1).TransSource <> 227 Then
            'Utility Breakdown Dist. *****************
            GoSub GetRevenueSources
            If NumofRevs > 0 Then
                Print #RptHandle, " "
                Print #RptHandle, Tab(10); "------------------Utility Revenue BrkDwn-----------------"
              For RCnt = 1 To NumofRevs Step 2
                  Print #RptHandle, Tab(15); RevText$(RCnt);
                  Print #RptHandle, Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(RCnt))
                  Print #RptHandle, Tab(15); RevText$(RCnt + 1);
                  Print #RptHandle, Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(RCnt + 1))
                  PrintUtilFlag = 1
              Next RCnt
            End If
          End If
        'End If
        Case 30 To 39, 131, 231
        If CMTrRec(1).TransSource >= 30 And CMTrRec(1).TransSource <= 39 Then
          TxRev# = 0
          For TRev = 1 To 9
            TxRev# = Round#(TxRev# + CMTrRec(1).TransRevAmt(TRev))
          Next
          TotalAmount# = Round#(TotalAmount# + TxRev#)
          Change# = Round#((CMTrRec(1).TransCash + CMTrRec(1).TransCheck) - TxRev#)
        Else '        Case 131, 231
          TxRev# = 0
          For TRev = 1 To 7
            TxRev# = Round#(TxRev# + CMTrRec(1).TransRevAmt(TRev))
          Next
          TxRev# = Round#(TxRev# + CMTrRec(1).TransRevAmt(9))
         ' TxRev# = Round#(TxRev# - CMTrRec(1).TransRevAmt(8))
          TotalAmount# = Round#(TotalAmount# + TxRev#)
          Change# = Round#((CMTrRec(1).TransCash + CMTrRec(1).TransCheck) - TxRev#)
        End If
        'Change# = Round#((CMTrRec(1).TransCash + CMTrRec(1).TransCheck) - CMTrRec(1).TransAmount)
        'If CHANGE# < 0 Then CHANGE# = 0
        Print #RptHandle, Tab(15); "  Change: "; Tab(40); Using(Fmt1$, Change#)
        
        If CMTrRec(1).TransSource >= 30 And CMTrRec(1).TransSource <= 39 Then
          'Tax Breakdown Dist.     *****************
          Print #RptHandle, " "
          Print #RptHandle, Tab(10); "-----------------------Tax BrkDwn---------------------"
          Print #RptHandle, Tab(15); "     Tax: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(1))
          Print #RptHandle, Tab(15); "Interest: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(2))
          Print #RptHandle, Tab(15); " Penalty: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(3))
          Print #RptHandle, Tab(15); "   Storm: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(4))
          Print #RptHandle, Tab(15); "Past Tax: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(6))
          Print #RptHandle, Tab(15); "Interest: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(7))
          Print #RptHandle, Tab(15); " Penalty: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(8))
          Print #RptHandle, Tab(15); "   Storm: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(9))
        Else
          Print #RptHandle, " "
          Print #RptHandle, Tab(10); "-----------------------Tax BrkDwn---------------------"
          Print #RptHandle, Tab(15); "PrePayment: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(9));
          Print #RptHandle, Tab(15); "# of Bills: "; Tab(40); CInt(CMTrRec(1).TransRevAmt(10))
          Print #RptHandle, Tab(15); " Principle:"; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(1));
          Print #RptHandle, Tab(15); "  Interest: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(2));
          Print #RptHandle, Tab(15); "Collection: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(3));
          Print #RptHandle, Tab(15); " Late List: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(4))
          Print #RptHandle, Tab(15); TxOpt1$; ": "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(5));
          Print #RptHandle, Tab(15); TxOpt2$; ": "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(6));
          Print #RptHandle, Tab(15); TxOpt3$; ": "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(7));
          Print #RptHandle, Tab(15); "  Discount: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(8))
        End If
        Case 161, 261
          TxRev# = 0
          For TRev = 1 To 8
            TxRev# = Round#(TxRev# + CMTrRec(1).TransRevAmt(TRev))
          Next
          TxRev# = Round#(TxRev# + CMTrRec(1).TransRevAmt(10))
          'TxRev# = Round#(TxRev# - CMTrRec(1).TransRevAmt(9))
          TotalAmount# = Round#(TotalAmount# + TxRev#)
          Change# = Round#((CMTrRec(1).TransCash + CMTrRec(1).TransCheck) - TxRev#)

          'Change# = Round#((CMTrRec(1).TransCash + CMTrRec(1).TransCheck) - CMTrRec(1).TransAmount)
          'If CHANGE# < 0 Then CHANGE# = 0
          Print #RptHandle, Tab(15); "  Change: "; Tab(40); Using(Fmt1$, Change#)
          Print #RptHandle, " "
          Print #RptHandle, Tab(10); "-----------------------Tax BrkDwn---------------------"
          Print #RptHandle, Tab(15); "PrePayment: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(10));
          Print #RptHandle, Tab(15); "# of Bills: "; Tab(40); CInt(CMTrRec(1).TransRevAmt(11))
          Print #RptHandle, Tab(15); " Principle:"; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(1));
          Print #RptHandle, Tab(15); "  Interest: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(2));
          Print #RptHandle, Tab(15); "Collection: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(3));
          Print #RptHandle, Tab(15); " Late List: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(4))
          Print #RptHandle, Tab(15); "  Penaltly: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(5));
          Print #RptHandle, Tab(15); TxOpt1$; ": "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(6));
          Print #RptHandle, Tab(15); TxOpt2$; ": "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(7));
          Print #RptHandle, Tab(15); TxOpt3$; ": "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(8));
          Print #RptHandle, Tab(15); "  Discount: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(9))
        Case 171, 271
          TxRev# = 0
          For TRev = 1 To 10
            TxRev# = Round#(TxRev# + CMTrRec(1).TransRevAmt(TRev))
          Next
          TxRev# = Round#(TxRev# + CMTrRec(1).TransRevAmt(12))
          'TxRev# = Round#(TxRev# - CMTrRec(1).TransRevAmt(11))
          TotalAmount# = Round#(TotalAmount# + TxRev#)
          Change# = Round#((CMTrRec(1).TransCash + CMTrRec(1).TransCheck) - TxRev#)
          'Change# = Round#((CMTrRec(1).TransCash + CMTrRec(1).TransCheck) - CMTrRec(1).TransAmount)
          'If CHANGE# < 0 Then CHANGE# = 0
          Print #RptHandle, Tab(15); "  Change: "; Tab(40); Using(Fmt1$, Change#)
          Print #RptHandle, " "
          Print #RptHandle, Tab(10); "-----------------------Tax BrkDwn---------------------"
          Print #RptHandle, Tab(15); "PrePayment: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(12));
          Print #RptHandle, Tab(15); "# of Bills: "; Tab(40); CInt(CMTrRec(1).TransRevAmt(13))
          Print #RptHandle, Tab(15); "Principle1: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(1));
          Print #RptHandle, Tab(15); "Principle2: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(2));
          Print #RptHandle, Tab(15); "Principle3: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(3));
          Print #RptHandle, Tab(15); "Principle4: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(4));
          Print #RptHandle, Tab(15); "Principle5: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(5));
          Print #RptHandle, Tab(15); "  Interest: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(6));
          Print #RptHandle, Tab(15); "  Penaltly: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(7));
          Print #RptHandle, Tab(15); TxPOpt1$; ": "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(8));
          Print #RptHandle, Tab(15); TxPOpt2$; ": "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(9));
          Print #RptHandle, Tab(15); TxPOpt3$; ": "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(10));
          Print #RptHandle, Tab(15); "  Discount: "; Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(11))
        Case 151, 251
          Change# = Round#((CMTrRec(1).TransCash + CMTrRec(1).TransCheck) - CMTrRec(1).TransAmount)
           'If CHANGE# < 0 Then CHANGE# = 0
          Print #RptHandle, Tab(15); "  Change: "; Tab(40); Using(Fmt1$, Change#)
          Print #RptHandle, " "
          GoSub GetVehInfo
           
        Case Else
           Change# = Round#((CMTrRec(1).TransCash + CMTrRec(1).TransCheck) - CMTrRec(1).TransAmount)
           'If CHANGE# < 0 Then CHANGE# = 0
           Print #RptHandle, Tab(15); "  Change: "; Tab(40); Using(Fmt1$, Change#)

        End Select

  Print #RptHandle, String$(80, "-")
  If RptType = 2 Then
    Print #RptHandle, Chr$(18);   ' oki 320 12 cpi
  End If
  Close         'Close all open files now


  'Erase RevName$, TotalMiscRec$, TotalMiscDesc$, TotalMiscAmt#
  'Erase TotalUtilRevAmt#, MiscCodeGL$
  'Erase Array1, CMTRRec, RevText$, MiscCodeRec, UBSetUpRec
  'Erase DistArray

  If RptType = 2 Then
    ViewPrint ReportFile$, Header$
  Else
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmCMDispList

      ARptLineRpt.GetName ReportFile$
      ARptLineRpt.startrpt

  End If
  Exit Sub

PrintRptHeader:
Return

GetRevenueSources:

  NumofRevs = MaxRevsCnt
  ReDim UBSetUpRec(1) As UBSetupRecType
  ReDim DistArray(1 To MaxRevsCnt) As DistArrayType
  ReDim UBSetUpRec(1) As UBSetupRecType
  UBSetupLen = Len(UBSetUpRec(1))
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen

  For RevCnt = 1 To MaxRevsCnt
    RevText$(RevCnt) = Left$(QPTrim$(UBSetUpRec(1).Revenues(RevCnt).RevName), 14)
    DistArray(RevCnt).DistOrder = UBSetUpRec(1).Revenues(RevCnt).DistOr
    DistArray(RevCnt).DistCnt = RevCnt
    If Len(RevText$(RevCnt)) = 0 Then
      NumofRevs = RevCnt - 1
      Exit For
    End If
  Next

  ReDim Preserve DistArray(1 To NumofRevs) As DistArrayType

  Do
    OutOfOrder = False          'assume it's sorted
    For x = 1 To NumofRevs - 1
      If DistArray(x).DistOrder > DistArray(x + 1).DistOrder Then
        Temp2 = DistArray(x).DistOrder
        DistArray(x).DistOrder = DistArray(x + 1).DistOrder
        DistArray(x + 1).DistOrder = Temp2
        'SWAP DistArray(x), DistArray(x + 1)     'if we had to swap
        OutOfOrder = True       'we're not done yet
      End If
    Next
  Loop While OutOfOrder

  
Return

GetVehInfo:
  VCCode$ = ""
  VCDesc$ = ""
  VCat = CInt(CMTrRec(1).TransRevAmt(2))
  Vlook = CLng(CMTrRec(1).TransRevAmt(1))
  Dim DCCatCodeRec As DCCatCodeRecType
  Dim DCCatCodeRecLen As Integer, ghandle As Integer
  DCCatCodeRecLen = Len(DCCatCodeRec)
  ghandle = FreeFile
  Open "DCCODE.DAT" For Random Access Read Write Shared As ghandle Len = DCCatCodeRecLen
  Get #ghandle, VCat, DCCatCodeRec
  If DCCatCodeRec.InactiveFlag <> "Y" Then
    VCCode$ = QPTrim$(DCCatCodeRec.CATCODE)
    VCDesc$ = QPTrim$(DCCatCodeRec.CODEDESC)
  End If
  Close ghandle
  ReDim DCVRec(1) As DCVehType
  Dim DCVehReclen  As Integer, DCvFile As Integer
  DCVehReclen = Len(DCVRec(1))
  DCvFile = FreeFile
  Open "DCVEH.DAT" For Random Access Read Write Shared As DCvFile Len = DCVehReclen
  Get DCvFile, Vlook, DCVRec(1)
          Print #RptHandle, Tab(10); "---------------------Vehicle Information---------------------"
          Print #RptHandle, Tab(15); "Decal Code: "; Tab(40); VCCode$ & " " & VCDesc$
          Print #RptHandle, Tab(15); "   Vehicle: "; Tab(40); QPTrim$(DCVRec(1).makemodel)
          Print #RptHandle, Tab(15); " State Lic: "; Tab(40); QPTrim$(DCVRec(1).StateTag)
          Print #RptHandle, Tab(15); "  Veh Desc: "; Tab(40); QPTrim$(DCVRec(1).Desc)
          Print #RptHandle, Tab(15); "     Notes: "; Tab(40); QPTrim$(DCVRec(1).Notes)
          Print #RptHandle, Tab(15); "       Fee: "; Tab(40); Using(Fmt1$, DCVRec(1).Fee)
  Close DCvFile

Return
End Sub

