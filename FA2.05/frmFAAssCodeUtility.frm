VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Begin VB.Form frmFAAssCodeUtility 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fixed Assets Asset Code Utility"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmFAAssCodeUtility.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3165
      Left            =   2310
      TabIndex        =   3
      Top             =   1890
      Width           =   7035
      _Version        =   196613
      _ExtentX        =   12409
      _ExtentY        =   5583
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   3
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmFAAssCodeUtility.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   675
      Left            =   2437
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5535
      Width           =   1875
      _Version        =   131072
      _ExtentX        =   3307
      _ExtentY        =   1191
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
      ButtonDesigner  =   "frmFAAssCodeUtility.frx":217E
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCont 
      Height          =   675
      Left            =   4717
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5535
      Width           =   4500
      _Version        =   131072
      _ExtentX        =   7937
      _ExtentY        =   1191
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
      ButtonDesigner  =   "frmFAAssCodeUtility.frx":2392
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Asset Code Conversion Utility"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   390
      Left            =   3832
      TabIndex        =   0
      Top             =   765
      Width           =   3945
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   690
      Left            =   3652
      Top             =   630
      Width           =   4350
   End
End
Attribute VB_Name = "frmFAAssCodeUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdCont_Click()
  Dim CodeHandle As Integer
  Dim CODEREC As FAAssetCodeRecType
  Dim x As Integer
  Dim SaveThis$
  Dim ItemHandle As Integer
  Dim NumOfItems As Long
  Dim Y As Long
  Dim ItemRec As DosFAItemRecType
  Dim ItemRecV1 As DosFAItemRecTypeV1
  Dim NumOfCodes As Integer
  Dim ItemsUpdated As Integer
  
  frmFAMsgBox.Label1.Caption = "Please Note: This utility feature automatically updates all fixed assets " _
  + "that use the invalid asset codes to the new edited asset codes."
  frmFAMsgBox.Label1.Top = 900
  frmFAMsgBox.Show vbModal
  
  OpenFACodeNameFile CodeHandle
  NumOfCodes = LOF(CodeHandle) / Len(CODEREC)
  vaSpread1.Col = 3
  
  ItemsUpdated = 0
  For x = 1 To NumOfBad
    vaSpread1.Row = x
    If Len(vaSpread1.Text) <> 4 Then
      If QPTrim$(vaSpread1.Text) = "" Then
        frmFAMsgBox.Label1.Caption = "Please enter a valid number on row " + CStr(x) + "."
      Else
        frmFAMsgBox.Label1.Caption = "Please edit " + QPTrim$(vaSpread1.Text) + " on row " + CStr(x) + " and make it a four digit number."
      End If
      frmFAMsgBox.Label1.Top = 1000
      frmFAMsgBox.Show vbModal
      vaSpread1.SetActiveCell 3, x
      Exit Sub
    End If
    For Y = 1 To NumOfCodes
      Get CodeHandle, Y, CODEREC
        If QPTrim(vaSpread1.Text) = QPTrim$(CODEREC.ASSETCODE) Then
          frmFAMsgBox.Label1.Caption = "The entry on row " + CStr(x) + " is already being used. Please enter a number that is not already being used."
          frmFAMsgBox.Label1.Top = 1000
          frmFAMsgBox.Show vbModal
          vaSpread1.SetActiveCell 3, x
          Close
          Exit Sub
        End If
    Next Y
  Next x
  
  If ThisVersion = 0 Then
    OpenFAItemFile ItemHandle
  Else
    OpenFAItemFileV1 ItemHandle
  End If
  
  If ThisVersion = 0 Then
    NumOfItems = LOF(ItemHandle) / Len(ItemRec)
    vaSpread1.Col = 3
    If ThisVersion = 0 Then
      For x = 1 To NumOfBad
        vaSpread1.Row = x
        Get CodeHandle, NNRecNum(x), CODEREC
          SaveThis$ = QPTrim$(vaSpread1.Text)
          CODEREC.ASSETCODE = SaveThis$
        Put CodeHandle, NNRecNum(x), CODEREC
        For Y = 1 To NumOfItems
          Get ItemHandle, Y, ItemRec
          If QPTrim$(ItemRec.ASSETCODE) = NotNumber(x) Then
            ItemRec.ASSETCODE = SaveThis
            ItemsUpdated = ItemsUpdated + 1
            Put ItemHandle, Y, ItemRec
          End If
        Next Y
        frmFAMsgBox.Label1.Caption = "A total of " + CStr(ItemsUpdated) + " items have had their assets codes " _
          + " updated from " + NotNumber(x) + " to " + SaveThis + "."
        frmFAMsgBox.Label1.Top = 900
        frmFAMsgBox.Show vbModal
        ItemsUpdated = 0
      Next x
    End If
  Else
    NumOfItems = LOF(ItemHandle) / Len(ItemRecV1)
    vaSpread1.Col = 3
    If ThisVersion = 0 Then
      For x = 1 To NumOfBad
        vaSpread1.Row = x
        Get CodeHandle, NNRecNum(x), CODEREC
          SaveThis$ = QPTrim$(vaSpread1.Text)
          CODEREC.ASSETCODE = SaveThis$
        Put CodeHandle, NNRecNum(x), CODEREC
        For Y = 1 To NumOfItems
          Get ItemHandle, Y, ItemRecV1
          If QPTrim$(ItemRec.ASSETCODE) = NotNumber(x) Then
            ItemRec.ASSETCODE = SaveThis
            ItemsUpdated = ItemsUpdated + 1
            Put ItemHandle, Y, ItemRecV1
          End If
        Next Y
        frmFAMsgBox.Label1.Caption = "A total of " + CStr(ItemsUpdated) + " items have had their assets codes " _
          + " updated from " + NotNumber(x) + " to " + SaveThis + "."
        frmFAMsgBox.Label1.Top = 900
        frmFAMsgBox.Show vbModal
        ItemsUpdated = 0
      Next x
    End If
  End If
   
  Close
  cmdCont.Enabled = False
  frmFAMsgBox.Label1.Caption = "The Asset Codes have been updated. Press ESC to return to the opening conversion screen and restart the conversion."
  frmFAMsgBox.Label1.Top = 900
  frmFAMsgBox.Show vbModal

End Sub

Private Sub cmdExit_Click()
  frmFixedAssetsConversion.Show
  DoEvents
  Unload Me
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      Call cmdExit_Click
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyF10:
      Call cmdCont_Click
      SendKeys "%S"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call LoadMe
End Sub

Private Sub LoadMe()
  Dim x As Integer
  
  Call FixSpread
  If NumOfBad = 0 Then Call cmdExit_Click
  For x = 1 To NumOfBad
    vaSpread1.Col = 1
    vaSpread1.Row = x
    vaSpread1.Text = NotNumber(x)
    vaSpread1.Col = 2
    vaSpread1.Text = NNDesc(x)
  Next x
  
End Sub

Private Sub FixSpread()
  Dim COne As Integer
  Dim CTwo As Integer
  Dim CThree As Integer
  Dim CFour As Integer
  Dim CFive As Integer
  Dim CSix As Integer
  Dim cnt As Integer
  '-1 means all rows or all columns....0 means headers
'    GoTo SkipAdjust
    Select Case ScreenW
      Case 1280
        If Screen.TwipsPerPixelX <> 12 Then
          COne = 5
          coladj = 10
          vaSpread1.FontSize = 18
          vaSpread1.RowHeight(-1) = 22
          vaSpread1.RowHeight(0) = 22
        Else
          COne = 13
          coladj = 4.5
          vaSpread1.RowHeight(-1) = 18
          vaSpread1.RowHeight(0) = 18
        End If
      Case 1152
        If Screen.TwipsPerPixelX <> 12 Then
          COne = 14
          coladj = 7
          vaSpread1.FontSize = 14
          vaSpread1.RowHeight(0) = 18.5
          vaSpread1.RowHeight(-1) = 18.5
        Else
          COne = 6.65
          coladj = 2.25
          vaSpread1.RowHeight(0) = 16
          vaSpread1.RowHeight(-1) = 17
        End If
      Case 1024
        If Screen.TwipsPerPixelX <> 12 Then
          COne = 13.49
          coladj = 5.65
          vaSpread1.RowHeight(0) = 14
          vaSpread1.RowHeight(-1) = 14
        Else
          COne = 1.2
          coladj = 0 '.35
        End If
      Case 800
        COne = 0
        coladj = -0.5
        vaSpread1.Font.Size = 12
        vaSpread1.RowHeight(-1) = 14
      Case Else
    End Select
SkipAdjust:
    vaSpread1.ColWidth(1) = vaSpread1.ColWidth(1)
    vaSpread1.ColWidth(2) = vaSpread1.ColWidth(2) + COne
    vaSpread1.ColWidth(3) = vaSpread1.ColWidth(3) + coladj

End Sub

