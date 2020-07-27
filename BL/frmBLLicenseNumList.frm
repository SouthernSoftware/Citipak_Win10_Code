VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLLicenseNumList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License License List"
   ClientHeight    =   6315
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   5580
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   5580
   StartUpPosition =   2  'CenterScreen
   Begin LpLib.fpList fpList1 
      Height          =   3480
      Left            =   1275
      TabIndex        =   0
      Top             =   1260
      Width           =   3000
      _Version        =   196608
      _ExtentX        =   5292
      _ExtentY        =   6138
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
      Columns         =   1
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
      ColDesigner     =   "frmBLLicenseNumList.frx":0000
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   492
      Left            =   690
      TabIndex        =   1
      ToolTipText     =   "Press to bring up a brief help screen."
      Top             =   5268
      Width           =   1980
      _Version        =   131072
      _ExtentX        =   3492
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
      ButtonDesigner  =   "frmBLLicenseNumList.frx":039C
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClose 
      Height          =   492
      Left            =   2898
      TabIndex        =   2
      ToolTipText     =   "Press to exit this screen."
      Top             =   5268
      Width           =   1980
      _Version        =   131072
      _ExtentX        =   3492
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
      ButtonDesigner  =   "frmBLLicenseNumList.frx":057F
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmBLLicenseNumList.frx":075D
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4332
      Left            =   192
      TabIndex        =   4
      Top             =   192
      Width           =   1020
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1104
      X2              =   1536
      Y1              =   624
      Y2              =   1440
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   636
      Left            =   1122
      Top             =   390
      Width           =   3312
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   6108
      Left            =   96
      Top             =   96
      Width           =   5388
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Current License Numbers"
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
      Left            =   1224
      TabIndex        =   3
      Top             =   540
      Width           =   3036
   End
End
Attribute VB_Name = "frmBLLicenseNumList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class
Private Sub cmdClose_Click()
  Unload frmBLLicenseNumList
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
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim x As Integer
  Dim BigNum As Double
  Dim ThisNum As Double
  Dim ThisLic$
  Dim SmallNum As Double
  Dim Nextx As Integer
  Dim LicCnt As Integer
  Dim ThisRec As Integer
  Dim HoldDbl As Double
  
  On Error Resume Next
  
  Label2.Visible = False
  Line1.Visible = False
  Nextx = 1
  ThisLic = FirstLicenseNum
  BigNum = CDbl(ThisLic)
  OpenCustFile CustHandle
  NumOfCustRecs = LOF(CustHandle) / Len(CustRec)
  ReDim LicArray(1 To 1) As Double
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
    ThisLic = QPTrim$(CustRec.LICENSE)
    If Not IsNumeric(ThisLic) Then GoTo Invalid
    LicArray(Nextx) = CDbl(ThisLic)
    Nextx = Nextx + 1
    ReDim Preserve LicArray(1 To Nextx)
    LicCnt = LicCnt + 1
Invalid:
  Next x
  Close CustHandle
  
  If LicCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No valid license numbers are currently saved"
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  SmallNum = BigNum + 1
  Nextx = 1
  
  Do
    For x = Nextx To LicCnt
      If LicArray(x) < SmallNum Then
        SmallNum = LicArray(x)
        ThisRec = x
      End If
    Next x
    HoldDbl = LicArray(Nextx)
    LicArray(Nextx) = LicArray(ThisRec)
    LicArray(ThisRec) = HoldDbl
    If Nextx = LicCnt Then Exit Do
    Nextx = Nextx + 1
    SmallNum = BigNum + 1
  Loop
  
  For x = LicCnt To 1 Step -1
'  For x = 1 To LicCnt
    fpList1.InsertRow = LicArray(x)
  Next x
  
  fpList1.ListIndex = 0
End Sub

Private Sub fpList1_DblClick()
  Dim ThisNum As String
  
  On Error Resume Next
  
  fpList1.Col = 0
  
  ThisNum = fpList1.ColText
  If QPTrim$(ThisNum$) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Nothing selected."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  If Exist("customeredit.dat") Then
    frmBLCustEdit.fptxtLicNum.Text = ThisNum
  Else
    frmBLPrintLic.fptxtBegNum.Text = ThisNum
  End If
  Call cmdClose_Click
  
End Sub
