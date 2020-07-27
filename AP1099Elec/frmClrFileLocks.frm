VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmClrFileLocks 
   BackColor       =   &H008A775B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clear File Locks"
   ClientHeight    =   8892
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   12192
   Icon            =   "frmClrFileLocks.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8892
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboModule 
      Height          =   384
      Left            =   6084
      TabIndex        =   0
      Top             =   3720
      Width           =   2028
      _Version        =   196608
      _ExtentX        =   3577
      _ExtentY        =   677
      Text            =   ""
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
      Object.TabStop         =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   0
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   2
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   2
      SearchMethod    =   2
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
      DataFieldList   =   ""
      ColumnEdit      =   -1
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
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
      ListPosition    =   0
      ButtonThreeDAppearance=   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      AutoSearchFill  =   -1  'True
      AutoSearchFillDelay=   100
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmClrFileLocks.frx":08CA
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
      Height          =   492
      Left            =   9456
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7008
      Width           =   1524
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
      Height          =   492
      Left            =   7488
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7008
      Width           =   1572
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   8532
      Width           =   12192
      _ExtentX        =   21505
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7133
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "12:30 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "10/5/2004"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Clear Locks On:"
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
      Height          =   324
      Index           =   0
      Left            =   4092
      TabIndex        =   6
      Top             =   3768
      Width           =   1764
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select A File To Clear Locks On:"
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
      Height          =   492
      Index           =   0
      Left            =   3336
      TabIndex        =   5
      Top             =   2808
      Width           =   5532
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3240
      Top             =   1008
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Clear Locks "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3792
      TabIndex        =   4
      Top             =   1248
      Width           =   4620
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2580
      Left            =   2688
      Top             =   2232
      Width           =   6972
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      Height          =   972
      Left            =   3240
      Top             =   888
      Width           =   5772
   End
End
Attribute VB_Name = "frmClrFileLocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim GLCREd As CJEditRecType
Dim GLCDEd As CJEditRecType
Dim POEdit As POFORMRecType2
Dim APIED As APInv85Type
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        ClearInUse PWcnt
      End If
    End If
  End If
End Sub


Private Sub cmdExit_Click()
  frmGLUtilMenu.Show
  Unload frmClrFileLocks
End Sub

Private Sub cmdOk_Click()
  FrmShowPctComp.Label1 = "Clearing File Locks"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmClrFileLocks
  Select Case fpcboModule.ListIndex
    Case 0:
      ClrCDLocks
    Case 1:
      ClrCRLocks
    Case 2:
      ClrInvLocks
    Case 3:
      ClrPOLocks
    Case Else:
      MsgBox "You Must Make A Selection.", vbOKOnly, "Selection Invalid"
  End Select
  ActivateControls frmClrFileLocks
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
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  fpcboModule.AddItem "Cash Disbursements" 'listindex 0
  fpcboModule.AddItem "Cash Receipts"      'listindex 1
  fpcboModule.AddItem "A/P Invoices"       'listindex 2
  fpcboModule.AddItem "Purchase Orders"    'listindex 3
End Sub

Public Sub ClrCDLocks()
  Dim CJEditFile As Integer, CJEditFileNum As Integer, NumEdTrans As Integer
  Dim cnt As Integer, CJType As Integer
  CJType = 2
  If Exist("GLCDEd.dat") Then
    OpenCJEditFile CJEditFileNum, NumEdTrans, CJType
    If NumEdTrans > 0 Then
      For cnt = 1 To NumEdTrans
        FrmShowPctComp.ShowPctComp cnt, NumEdTrans
        Get CJEditFileNum, cnt, GLCDEd
        GLCDEd.LOCKED = False
        Put CJEditFileNum, cnt, GLCDEd
      Next
    End If
  Close CJEditFileNum
  Call MainLog("CD File Locks Cleared.")
  MsgBox "Clear Cash Disbursements File Locks Complete.", vbOKOnly, "Complete"
  End If
End Sub
Public Sub ClrCRLocks()
  Dim CJEditFile As Integer, CJEditFileNum As Integer, NumEdTrans As Integer
  Dim cnt As Integer, CJType As Integer
  CJType = 1
  If Exist("GLCREd.dat") Then
    OpenCJEditFile CJEditFileNum, NumEdTrans, CJType
    If NumEdTrans > 0 Then
      For cnt = 1 To NumEdTrans
        FrmShowPctComp.ShowPctComp cnt, NumEdTrans
        Get CJEditFileNum, cnt, GLCREd
        GLCREd.LOCKED = False
        Put CJEditFileNum, cnt, GLCREd
      Next
    End If
  Close CJEditFileNum
  Call MainLog("CR File Locks Cleared.")
  MsgBox "Clear Cash Receipts File Locks Complete.", vbOKOnly, "Complete"
  End If
End Sub
Public Sub ClrInvLocks()
  Dim APEditFile As Integer, NumEdTrans As Integer
  Dim cnt As Integer
  If Exist("APIED.DAT") Then
    OpenAPEditFile APEditFile, NumEdTrans
    If NumEdTrans > 0 Then
      For cnt = 1 To NumEdTrans
        FrmShowPctComp.ShowPctComp cnt, NumEdTrans
        Get APEditFile, cnt, APIED
        APIED.LOCKED = False
        Put APEditFile, cnt, APIED
      Next
    Else
      FrmShowPctComp.ShowPctComp 1, 1
    End If
    Close APEditFile
    Call MainLog("Invoice File Locks Cleared.")
    MsgBox "Clear A/P Invoice File Locks Complete.", vbOKOnly, "Complete"
  End If
End Sub
Public Sub ClrPOLocks()
  Dim POEditFile As Integer, NumEdTrans As Integer
  Dim cnt As Integer
  If Exist("APPED.dat") Then
    OpenPOEditFile POEditFile, NumEdTrans
    If NumEdTrans > 0 Then
      For cnt = 1 To NumEdTrans
        FrmShowPctComp.ShowPctComp cnt, NumEdTrans
        Get POEditFile, cnt, POEdit
        POEdit.LOCKED = False
        Put POEditFile, cnt, POEdit
      Next
    Else
      FrmShowPctComp.ShowPctComp 100, 100
    End If
    Close POEditFile
    Call MainLog("PO File Locks Cleared.")
    MsgBox "Clear Purchase Order File Locks Complete.", vbOKOnly, "Complete"
  Else
    FrmShowPctComp.ShowPctComp 1, 1
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
