VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmBLCustomerList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Customer List"
   ClientHeight    =   6996
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   7920
   Icon            =   "frmBLCustomerList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6996
   ScaleWidth      =   7920
   StartUpPosition =   1  'CenterOwner
   Begin LpLib.fpList fpList1 
      Height          =   4080
      Left            =   600
      TabIndex        =   0
      ToolTipText     =   "Customer listing."
      Top             =   1296
      Width           =   6780
      _Version        =   196608
      _ExtentX        =   11959
      _ExtentY        =   7197
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
      ColDesigner     =   "frmBLCustomerList.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClose 
      Height          =   492
      Left            =   4140
      TabIndex        =   3
      Top             =   5568
      Width           =   2364
      _Version        =   131072
      _ExtentX        =   4170
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
      ButtonDesigner  =   "frmBLCustomerList.frx":0C22
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   492
      Left            =   1416
      TabIndex        =   2
      Top             =   5568
      Width           =   2364
      _Version        =   131072
      _ExtentX        =   4170
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
      ButtonDesigner  =   "frmBLCustomerList.frx":0E00
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmBLCustomerList.frx":0FE3
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Left            =   480
      TabIndex        =   4
      Top             =   6144
      Width           =   6924
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   768
      X2              =   1344
      Y1              =   6192
      Y2              =   4992
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   636
      Left            =   1938
      Top             =   492
      Width           =   4044
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer List"
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
      Left            =   2064
      TabIndex        =   1
      Top             =   642
      Width           =   3900
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   6684
      Left            =   144
      Top             =   156
      Width           =   7644
   End
End
Attribute VB_Name = "frmBLCustomerList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim fromform As Form, toform As Form, codeopt As Integer
Public Sub Wheretogo(xfrm As Form, tfrm As Form, Optional opt As Integer)
  Set fromform = xfrm
  Set toform = tfrm
  If opt <> 0 Then
    codeopt = opt
  Else
    codeopt = 0
  End If
End Sub

Private Sub cmdClose_Click()
  fpList1.Clear
  Unload frmBLCustomerList
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
  Set Over = New clsTextBoxOverRider
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
      Call cmdClose_Click
      SendKeys "%C"
      KeyCode = 0
    Case vbKeyF1:
      Call cmdHelp_Click
      SendKeys "%T"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub LoadMe()
  Dim CustRec As ARCustRecType
  Dim CustIdxRec As CustNameIdxType
  Dim CustIdxHandle As Integer
  Dim CustIdxRecNum As Integer
  Dim CHandle As Integer
  Dim TotalAccts As Integer
  Dim x As Integer
  Dim CustCnt As Integer
   
  On Error GoTo ERRORSTUFF
  
  Label2.Visible = False
  Line1.Visible = False

  If Not Exist("arcustnameidx.dat") Then 'no file there
    frmBLMessageBoxJr.Label1.Caption = "No Customer Name Index has been saved."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  OpenCustNameIdxFile CustIdxHandle
  CustIdxRecNum = LOF(CustIdxHandle) \ Len(CustIdxRec)
  If CustIdxRecNum = 0 Then 'file is there but there is nothing in it
    frmBLMessageBoxJr.Label1.Caption = "No Customers in index."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
   
  ReDim CustIdx(1 To CustIdxRecNum) As Integer
  For x = 1 To CustIdxRecNum
    Get CustIdxHandle, x, CustIdxRec
    CustIdx(x) = CustIdxRec.CustRec 'load array with record pointers
  Next x
  Close CustIdxHandle
  
  If Not Exist("ARCUST.DAT") Then
    frmBLMessageBoxJr.Label1.Caption = "Path to ARCUST.DAT could not be found."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
   
  OpenBLCustFile CHandle
  
  If CustIdxRecNum = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No Customers on file."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  For x = 1 To CustIdxRecNum 'CustCnt
    Get CHandle, CustIdx(x), CustRec
    If Len(QPTrim(CustRec.BILLNAME)) = 0 Or QPTrim(CustRec.SORTNAME) = "DELETED" Then GoTo BadCode
    fpList1.InsertRow = QPTrim$(CustRec.CustName) & " " & Chr$(9) & QPTrim$(CustRec.CUSTNUMB)
BadCode:
  Next x
  Close CHandle
  fpList1.Row = 0
  fpList1.Selected = True 'set focus to first line
ZeroText:
  Exit Sub
   

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustomerList", "LoadMe", Erl)
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
 '   ClearInUse PWcnt
 '   CitiTerminate
    Unload Me
End Sub

Private Sub fpList1_DblClick()
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim TotalAccts As Integer
  Dim x As Integer
  Dim Number$
  Dim Name$
  Dim Found As Boolean
  Dim One As Integer
  Dim DHandle As Integer
  
  On Error GoTo ERRORSTUFF
  One = 1
  DHandle = FreeFile
  Open "custlistopen.dat" For Output As DHandle Len = 2
  Print #DHandle, One
  Close DHandle
   
  'if the customer edit screen is open and the user is
  'changing from one customer to another then this next
  'code checks to see if any changes were made first and if so
  'gives the user the opportunity to save them
'NOT IN CM
'  If Exist("customeredit.dat") Then
'    frmBLCustomerList.Hide
'    Call frmBLCustEdit.cmdExit_Click
'    If ItemChangeFlag = True Then
'      ItemChangeFlag = False
'      Unload frmBLCustomerList
'      Exit Sub
'    End If
'  End If
GCustNumIsZero:
  fpList1.col = 0 'assign variables from the user selected row
  Name$ = QPTrim$(fpList1.ColText)
  fpList1.col = 1
  Number = QPTrim$(fpList1.ColText)
  
  OpenBLCustFile CHandle
  TotalAccts = LOF(CHandle) \ Len(CustRec)
  
  If TotalAccts = 0 Then Exit Sub
  
  For x = 1 To TotalAccts
    Get CHandle, x, CustRec
    If Name$ = QPTrim$(CustRec.CustName) And Number = QPTrim$(CustRec.CUSTNUMB) Then 'match the selected
    'row with the right code
      Found = True
      fpList1.Row = -1
      GCustNum = x 'now you can assign the correct global
      Exit For
    Else
      Found = False
      GoTo NotAMatch
    End If
      
NotAMatch:
   Next x
  Close CHandle
  
  If Found = True Then
    If codeopt = 2 Then
'    If Exist("adjustbalance.dat") Then
'      KillFile "custlistopen.dat"
'      Call frmBLAdjustBal.ClearScreen
      Call frmBLAdjustBal.LoadMe
      Unload frmBLCustomerList
'      Exit Sub
'    ElseIf Exist("transentry.dat") Then
'      KillFile "custlistopen.dat"
    ElseIf codeopt = 1 Then
      Call frmPayBLEntry.EnterEditChk
      Unload frmBLCustomerList
      
'      Exit Sub
    End If
'    'Call frmBLCustEdit.LoadMe
'    Unload frmBLCustomerList
'    Exit Sub
  Else
    frmBLMessageBoxJr.Label1.Caption = "No match found."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  Exit Sub

ERRORSTUFF:
   Unload FrmShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustomerList", "fpList1_DblClick", Erl)
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
  '  ClearInUse PWcnt
  '  CitiTerminate
  '

End Sub
  


