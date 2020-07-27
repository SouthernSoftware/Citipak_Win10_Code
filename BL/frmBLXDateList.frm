VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLXDateList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Expiration Date List"
   ClientHeight    =   6825
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   7845
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   7845
   StartUpPosition =   1  'CenterOwner
   Begin LpLib.fpList fpList1 
      Height          =   3885
      Left            =   570
      TabIndex        =   0
      Top             =   1290
      Width           =   6690
      _Version        =   196608
      _ExtentX        =   11800
      _ExtentY        =   6853
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      ColDesigner     =   "frmBLXDateList.frx":0000
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClose 
      Height          =   492
      Left            =   4176
      TabIndex        =   1
      Top             =   5532
      Width           =   2220
      _Version        =   131072
      _ExtentX        =   3916
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
      ButtonDesigner  =   "frmBLXDateList.frx":0390
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   492
      Left            =   1452
      TabIndex        =   3
      Top             =   5532
      Width           =   2220
      _Version        =   131072
      _ExtentX        =   3916
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
      ButtonDesigner  =   "frmBLXDateList.frx":05A6
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmBLXDateList.frx":07C2
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   288
      TabIndex        =   4
      Top             =   6048
      Width           =   7404
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1392
      X2              =   1200
      Y1              =   5136
      Y2              =   6096
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   6684
      Left            =   48
      Top             =   48
      Width           =   7740
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   636
      Left            =   1206
      Top             =   384
      Width           =   5424
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Business License Expiration Dates"
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
      Height          =   300
      Left            =   1398
      TabIndex        =   2
      Top             =   540
      Width           =   5052
   End
End
Attribute VB_Name = "frmBLXDateList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class
Private Sub cmdClose_Click()
  Unload frmBLXDateList
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
    Case vbKeyReturn:
      Call fpList1_DblClick
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdClose_Click
      KeyCode = 0
    Case vbKeyF3:
      SendKeys "%H"
      Call cmdHelp_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub LoadMe()
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim x As Integer
  Dim NoInActive As Boolean
  
  On Error Resume Next
  
  Label2.Visible = False
  Line1.Visible = False

  NoInActive = False
  If Exist("XlistInactiveY.dat") Then 'this .dat file is created only when
  'the inactive option is yes on the Expired License Report screen
    NoInActive = True
  End If
  OpenCustNameIdxFile IdxHandle
  NumOfCustRecs = LOF(IdxHandle) / Len(IdxRec)
  
  If NumOfCustRecs = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No valid license numbers are currently saved"
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  OpenCustFile CustHandle
  For x = 1 To NumOfCustRecs
    Get IdxHandle, x, IdxRec
    Get CustHandle, IdxRec.CustRec, CustRec
    If QPTrim$(CustRec.SortName) = "DELETED" Or QPTrim$(CustRec.Deleted) = "Y" Then GoTo NotThisOne
    If NoInActive = False Then
      If CustRec.Inactive = "Y" Then GoTo NotThisOne
    End If
    fpList1.InsertRow = QPTrim$(CustRec.BillName) & Chr$(9) & MakeRegDate(CustRec.VALID)
NotThisOne:
  Next x
  
  fpList1.ListIndex = 0
  
  Close CustHandle
  Close IdxHandle
 
End Sub

Private Sub fpList1_DblClick()
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim x As Integer
  Dim XDate$
  
  On Error GoTo ERRORSTUFF
  
  frmBLXDateList.Hide
   
  fpList1.Col = 1
  XDate$ = QPTrim$(fpList1.ColText)
  
  If Exist("custappsRenews.dat") Then
    frmBLPrintAppsRenwls.fptxtNewXDate.Text = XDate$
    frmBLPrintAppsRenwls.fptxtNewXDate.SetFocus
    Unload frmBLXDateList
    Exit Sub
  End If
  
  If Exist("pencalcscr.dat") Then
    frmBLPenCalc.fptxtXDate.Text = XDate$
    frmBLPenCalc.fptxtXDate.SetFocus
    Unload frmBLXDateList
    Exit Sub
  End If
  
  If Exist("dlnqnotice.dat") Then
    frmBLDelinquentNotices.fptxtXDate.Text = XDate$
    frmBLDelinquentNotices.fptxtXDate.SetFocus
    Unload frmBLXDateList
    Exit Sub
  End If
  
  If Exist("dlnqmllbls.dat") Then
    frmBLDlqntMailLbls.fptxtXDate.Text = XDate$
    frmBLDlqntMailLbls.fptxtXDate.SetFocus
    Unload frmBLXDateList
    Exit Sub
  End If
  
  If Exist("advanceltrprint.dat") Then
    frmBLPrintAdvanceLetter.fptxtNewXDate.Text = XDate$
    frmBLPrintAdvanceLetter.fptxtNewXDate.SetFocus
    Unload frmBLXDateList
    Exit Sub
  End If

  If Exist("mllbls.dat") Then
    frmBLMailLbls.fptxtXDate.Text = XDate$
    frmBLMailLbls.fptxtXDate.SetFocus
    Unload frmBLXDateList
    Exit Sub
  End If
  
  If Exist("custXlicList.dat") Then
    frmBLXLicList.fptxtXDate.Text = XDate$
    frmBLXLicList.fptxtXDate.SetFocus
    Unload frmBLXDateList
    Exit Sub
  End If
  
  If Exist("setstatus.dat") Then
    frmBLChangeLicPrintStatus.fpDateXDate.Text = XDate$
    frmBLChangeLicPrintStatus.fpDateXDate.SetFocus
    Unload frmBLXDateList
    Exit Sub
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLXDateList", "fpList1_DblClick", Erl)
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
