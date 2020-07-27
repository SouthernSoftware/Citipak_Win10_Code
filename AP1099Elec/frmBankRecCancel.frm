VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmChkRecCancel 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Reconciliation Check List"
   ClientHeight    =   8895
   ClientLeft      =   30
   ClientTop       =   495
   ClientWidth     =   12195
   Icon            =   "frmBankRecCancel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fplstChks 
      Height          =   1860
      Left            =   1350
      TabIndex        =   4
      Top             =   3000
      Width           =   9540
      _Version        =   196608
      _ExtentX        =   16828
      _ExtentY        =   3281
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Columns         =   7
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   1
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
      ColDesigner     =   "frmBankRecCancel.frx":08CA
   End
   Begin VB.CommandButton cmddatesrt 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Sort List by Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   1008
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7992
      Width           =   2844
   End
   Begin VB.CommandButton cmdSource 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Sort List by Source"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   1008
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7608
      Width           =   2844
   End
   Begin VB.CommandButton cmdchk 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Sort List by Check Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   1008
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7224
      Width           =   2844
   End
   Begin EditLib.fpText fptxtTotChks 
      Height          =   324
      Left            =   4206
      TabIndex        =   14
      Top             =   6648
      Width           =   1140
      _Version        =   196608
      _ExtentX        =   2011
      _ExtentY        =   572
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   1
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Alt-C &Clear All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   9072
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6096
      Width           =   1620
   End
   Begin VB.CommandButton cmdMark 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Alt-M &Mark All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   7128
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6096
      Width           =   1620
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Esc E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   9720
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7512
      Width           =   1236
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 &Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   7860
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7512
      Width           =   1236
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   8532
      Width           =   12192
      _ExtentX        =   21511
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7117
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "12:04 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "6/28/2008"
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
   Begin EditLib.fpText fptxtTag 
      Height          =   324
      Left            =   6798
      TabIndex        =   16
      Top             =   6648
      Width           =   1140
      _Version        =   196608
      _ExtentX        =   2011
      _ExtentY        =   572
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   1
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tagged"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   10152
      TabIndex        =   17
      Top             =   2640
      Width           =   660
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tagged:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   5934
      TabIndex        =   15
      Top             =   6720
      Width           =   828
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Checks:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   2886
      TabIndex        =   13
      Top             =   6720
      Width           =   1308
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Spacebar or Click to Toggle, F10 to Continue. "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   324
      Left            =   1794
      TabIndex        =   12
      Top             =   6144
      Width           =   4140
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Chk No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   1536
      TabIndex        =   11
      Top             =   2640
      Width           =   876
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   7
      Left            =   4680
      TabIndex        =   8
      Top             =   2640
      Width           =   1092
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Source"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   7560
      TabIndex        =   7
      Top             =   2640
      Width           =   684
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   3
      Left            =   8880
      TabIndex        =   6
      Top             =   2640
      Width           =   852
   End
   Begin VB.Label Label51 
      BackStyle       =   0  'Transparent
      Caption         =   "Chk Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   2712
      TabIndex        =   5
      Top             =   2640
      Width           =   996
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   852
      Left            =   2580
      Top             =   960
      Width           =   7020
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Checks To Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3684
      TabIndex        =   3
      Top             =   1200
      Width           =   4836
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   996
      Left            =   2592
      Top             =   816
      Width           =   7020
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   3156
      Left            =   1194
      Top             =   2856
      Width           =   9828
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      Height          =   4620
      Left            =   990
      Top             =   2448
      Width           =   10212
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuPrnScn 
         Caption         =   "Prin&t Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmChkRecCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim PCnt As Long, DCnt As Long, Edit As Boolean

Private Sub cmdchk_Click()
  fplstChks.Col = 1
  fplstChks.ColSortSeq = -1
  fplstChks.Col = 3
  fplstChks.ColSortSeq = -1
  fplstChks.Col = 0
  fplstChks.ColSortSeq = 0
  fplstChks.ColSorted = SortedDescending
End Sub

Private Sub cmddatesrt_Click()
  fplstChks.Col = 0
  fplstChks.ColSortSeq = -1
  fplstChks.Col = 3
  fplstChks.ColSortSeq = -1
  fplstChks.Col = 1
  fplstChks.ColSortSeq = 0
  fplstChks.ColSorted = SortedAscending
End Sub
Private Sub cmdSource_Click()
  fplstChks.Col = 1
  fplstChks.ColSortSeq = -1
  fplstChks.Col = 0
  fplstChks.ColSortSeq = -1
  fplstChks.Col = 3
  fplstChks.ColSortSeq = 0
  fplstChks.ColSorted = SortedDescending
End Sub

Private Sub cmdExit_Click()
  If Edit = True Then
    If MsgBox("Abandon Changes?", vbYesNo, "Abandon?") = vbNo Then
      savechks
    End If
  End If
  KillFileD "crchek.opn"
  frmBankReconMenu.Show
  Unload frmChkRecCancel
End Sub

Private Sub cmdOk_Click()
  If Not Exist("APCHKINF.DAT") Then
    savechks
    KillFileD "crchek.opn"
    frmBankReconMenu.Show
    Unload frmChkRecCancel
  Else
    MsgBox "UnPosted Check File Exists - Editing Check Reconciliation NOT Allowed.", vbOKOnly, "Canceled"
    KillFileD "crchek.opn"
    frmBankReconMenu.Show
    Unload frmChkRecCancel
  End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      If Edit = True Then
        If MsgBox("Abandon Changes and Close?", vbYesNo, "Abandon?") = vbNo Then
          Cancel = True
        End If
      End If
      KillFileD "crchek.opn"
      ClearInUse PWcnt
    End If
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%O"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  Me.HelpContextID = hlpSelectChecksTo
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Loadchks
  Edit = False
End Sub
Private Sub cmdClear_Click()
  Dim ccnt As Integer
  fplstChks.Action = ActionDeselectAll
  For ccnt = 0 To fplstChks.ListCount - 1
    fplstChks.ListIndex = ccnt
    fplstChks.Col = 5
    fplstChks.ColText = ""
  Next
  PCnt = 0
  fptxtTag = PCnt
End Sub

Private Sub cmdMark_Click()
  Dim ccnt As Integer
  fplstChks.Action = ActionSelectAll
  For ccnt = 0 To fplstChks.ListCount - 1
    fplstChks.ListIndex = ccnt
    fplstChks.Col = 5
    fplstChks.ColText = "*"
  Next
 PCnt = DCnt
 fptxtTag = PCnt
End Sub
Private Sub fplstChks_SelChange(ItemIndex As Long)
'When item in list is selected need to update the tagged column with *
  If fplstChks.Selected(ItemIndex) = True Then
    fplstChks.Col = 5
    fplstChks.ColText = "*"
  Else
    fplstChks.ColText = ""
  End If
  fptxtTag = fplstChks.SelCount
  Edit = True
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub Loadchks()
  Dim ccnt As Integer, OSChekFile As Integer, OSChkDate As String
  Dim NumOSChks As Integer, T As String, Tag As String
  Dim tempstr As String, TempAmt As String, ncnt As Integer
  Dim OSChek As OSChekRecType
  OpenOSChekFile OSChekFile, NumOSChks
  If OSChekFile < 0 Then
    Exit Sub
  End If
  If NumOSChks = 0 Then
    MsgBox "No Checks To Cancel", vbOKOnly, "NO Checks"
    Close
    GoTo AbortExit
  End If
  DCnt = 0
  PCnt = 0
  fplstChks.Clear
  For ccnt = 1 To NumOSChks
    Get OSChekFile, ccnt, OSChek
    DCnt = DCnt + 1
    fptxtTotChks = DCnt

    Select Case OSChek.Src
    Case 0
      T$ = "A"
    Case 1
      T$ = "P"
    End Select
    If OSChek.Cleared = 0 Then
      Tag$ = ""
    Else
      Tag$ = "*"
    End If
    OSChkDate$ = Format(DateAdd("d", OSChek.chkdate, "12-31-1979"), "mm/dd/yy")
    tempstr = Space$(30)
    TempAmt$ = Using("#######.##", Str$(OSChek.Amt))
    fplstChks.AddItem Using("########", Str$(OSChek.ChkNum)) & Chr$(9) & OSChkDate$ & Chr$(9) & QPTrim$(OSChek.Desc) & Chr$(9) & T$ & Chr$(9) & TempAmt$ & Chr$(9) & Tag$ & Chr$(9) & ccnt
    fplstChks.ListApplyTo = ListApplyToIndividual
    fplstChks.Col = 4
    fplstChks.AlignH = AlignHRight
    fplstChks.Col = 5
    fplstChks.AlignH = AlignHRight
    If Tag$ = "*" Then
      fplstChks.Selected(ccnt - 1) = True
    End If
    Next
  Close
  fptxtTag = fplstChks.SelCount
AbortExit:

End Sub
Private Sub savechks()
  Dim ccnt As Integer, OSChekFile As Integer, OSChkDate As String
  Dim NumOSChks As Integer, Num As Long
  Dim OSChek As OSChekRecType
  OpenOSChekFile OSChekFile, NumOSChks
  For ccnt = 0 To fplstChks.ListCount - 1
    fplstChks.ListIndex = ccnt
    fplstChks.Col = 6
    Num = QPTrim(fplstChks.ColText)
    Get OSChekFile, Num, OSChek
    fplstChks.Col = 5
    If fplstChks.ColText = "*" Then
      OSChek.Cleared = 1
    Else
      OSChek.Cleared = 0
    End If
    Put OSChekFile, Num, OSChek
  Next
  Close
  MsgBox "Data Saved.", vbOKOnly, "Saved"
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

