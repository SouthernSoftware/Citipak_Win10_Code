VERSION 5.00
Begin VB.Form frmGLUtilMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Ledger Utilities Menu"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   12195
   Icon            =   "frmGLUtilMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdEditPOTrans 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Edit PO Trans"
      Height          =   372
      Left            =   6288
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7416
      Width           =   2412
   End
   Begin VB.CommandButton cmdPurgePOs 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Purge PO Trans "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6288
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5496
      Width           =   2412
   End
   Begin VB.CommandButton cmdViewMainLog 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Vie&w G/L Log"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6288
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6960
      Width           =   2412
   End
   Begin VB.CommandButton cmdClrLocksMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Clear Locks "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6288
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2160
      Width           =   2412
   End
   Begin VB.CommandButton cmdListPo 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "List &PO Trans"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6288
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6468
      Width           =   2412
   End
   Begin VB.CommandButton cmdRelinkPO 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Re-Lin&k PO Trans "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6288
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5988
      Width           =   2412
   End
   Begin VB.CommandButton cmdRelinkAP 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "A&P Ledger Utilities"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6288
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5016
      Width           =   2412
   End
   Begin VB.CommandButton cmdViewGLLog 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "&View G/L Util Log"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3504
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7416
      Width           =   2412
   End
   Begin VB.CommandButton cmdFixBgtTrans 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "F&ix Bgt Transaction"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6288
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4536
      Width           =   2412
   End
   Begin VB.CommandButton cmdFixBgtEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Fix Bgt Edit File "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6288
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4056
      Width           =   2412
   End
   Begin VB.CommandButton cmdEditAcct 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Edit Acc&ount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6288
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3576
      Width           =   2412
   End
   Begin VB.CommandButton cmdSearchDup 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Searc&h Duplicate Accts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6288
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3096
      Width           =   2412
   End
   Begin VB.CommandButton cmdSearchRep 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Search/&Replace Trans"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6288
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2616
      Width           =   2412
   End
   Begin VB.CommandButton cmdRelinkBgt 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Re-Link &Bgt Trans "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3504
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6960
      Width           =   2412
   End
   Begin VB.CommandButton cmdPrnTransFile 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Print &Transaction File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3528
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   2412
   End
   Begin VB.CommandButton cmdScanRange 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Scan For Out of Range"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3528
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2628
      Width           =   2412
   End
   Begin VB.CommandButton cmdUpdateQuery 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Update Query"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3504
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3588
      Width           =   2412
   End
   Begin VB.CommandButton cmdListTag 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "&List Tagged Trans"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3528
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4068
      Width           =   2412
   End
   Begin VB.CommandButton cmdTagRec 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Tag Re&cords"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3528
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4548
      Width           =   2412
   End
   Begin VB.CommandButton cmdEditTrans 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Edit Transactions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3528
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3108
      Width           =   2412
   End
   Begin VB.CommandButton cmdClearTag 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Clear &All Tagged Records"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3528
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5496
      Width           =   2412
   End
   Begin VB.CommandButton cmdUnTag 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "U&n-Tag Records"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3528
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5028
      Width           =   2412
   End
   Begin VB.CommandButton cmdExitGLUtilMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "E&xit Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4944
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7872
      Width           =   2412
   End
   Begin VB.CommandButton cmdDeleteTag 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Delete Tagged Records"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3528
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5988
      Width           =   2412
   End
   Begin VB.CommandButton cmdRelinkGL 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Re-Link &G/L Trans "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3528
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6468
      Width           =   2412
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   8850
      X2              =   9810
      Y1              =   2196
      Y2              =   2196
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2370
      X2              =   3330
      Y1              =   2196
      Y2              =   2196
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GENERAL LEDGER UTILITIES"
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
      Left            =   3750
      TabIndex        =   25
      Top             =   1236
      Width           =   4692
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   8970
      X2              =   8970
      Y1              =   2196
      Y2              =   8076
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2496
      X2              =   2496
      Y1              =   2208
      Y2              =   8088
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8970
      Top             =   2196
      Width           =   732
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2496
      Top             =   2196
      Width           =   732
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   8970
      X2              =   9690
      Y1              =   8076
      Y2              =   8076
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2490
      X2              =   3210
      Y1              =   8076
      Y2              =   8076
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   8850
      X2              =   8850
      Y1              =   2076
      Y2              =   2196
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   3
      X1              =   8850
      X2              =   9810
      Y1              =   2076
      Y2              =   2076
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   9810
      X2              =   9810
      Y1              =   2100
      Y2              =   2196
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   2370
      X2              =   3330
      Y1              =   2076
      Y2              =   2076
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   3330
      X2              =   3330
      Y1              =   2076
      Y2              =   2196
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2370
      X2              =   2370
      Y1              =   2076
      Y2              =   2196
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1770
      Top             =   876
      Width           =   8652
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00D0D0D0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1770
      Top             =   756
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2370
      Top             =   1956
      Width           =   972
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   8850
      Top             =   1956
      Width           =   972
   End
End
Attribute VB_Name = "frmGLUtilMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim GLSetup As GLSetupRecType
Dim GLTrans   As GLTransRecType
Dim POTrans As GLTransRecType

Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim FY1BegDate As Integer, FY1EndDate As Integer, FY2BegDate As Integer, FY2EndDate As Integer

Private Sub cmdClearTag_Click()
  frmClearTags.Show
  Unload frmGLUtilMenu
End Sub

Private Sub cmdClrLocksMenu_Click()
  frmClrFileLocks.Show
  Unload frmGLUtilMenu
End Sub

Private Sub cmdDeleteTag_Click()
  frmDeleteTagTrans.Show
  Unload frmGLUtilMenu
End Sub

Private Sub cmdEditAcct_Click()
  frmEditAccount.Show
  Unload frmGLUtilMenu
End Sub

Private Sub cmdEditPOTrans_Click()
  frmEditPOTrans.Show
  Unload frmGLUtilMenu
End Sub

Private Sub cmdEditTrans_Click()
  frmEditTransaction.Show
  Unload frmGLUtilMenu
End Sub

Private Sub cmdFixBgtEdit_Click()
  frmFixBgtEdit.Show
  Unload frmGLUtilMenu
End Sub

Private Sub cmdFixBgtTrans_Click()
  frmFixBgtTransDate.Show
  Unload frmGLUtilMenu
End Sub

Private Sub cmdListPo_Click()
  frmReportOpt.Show 1
  If rptopt = 1 Then
    PrnPOTransFile
  ElseIf rptopt = 2 Then
    PrnPOTransFile2
  End If
End Sub

Private Sub cmdListTag_Click()
  frmReportOpt.Show 1
  If rptopt = 1 Then
    ListMTrans
  ElseIf rptopt = 2 Then
    ListMTrans2
  End If
End Sub

Private Sub cmdPrnTransFile_Click()
  frmReportOpt.Show 1
  If rptopt = 1 Then
    PrnTransFile
  ElseIf rptopt = 2 Then
    PrnTransFile2
  End If
End Sub

Private Sub cmdPurgePOs_Click()
  frmPurgePOs.Show
End Sub

Private Sub cmdRelinkAP_Click()
  frmAPLdgUtilMenu.Show
  Unload frmGLUtilMenu
End Sub

Private Sub cmdRelinkBgt_Click()
  frmRelinkBgtTrans.Show
  Unload frmGLUtilMenu
End Sub

Private Sub cmdRelinkGL_Click()
  frmRelinkGLTrans.Show
  Unload frmGLUtilMenu
End Sub

Private Sub cmdRelinkPO_Click()
  
  ReLinkPOTrans frmGLUtilMenu
End Sub

Private Sub cmdScanRange_Click()
  frmReportOpt.Show 1
  If rptopt = 1 Then
    DateScan
  ElseIf rptopt = 2 Then
    DateScan2
  End If
End Sub

Private Sub cmdSearchDup_Click()
  frmSearch4Dup.Show
  Unload frmGLUtilMenu
End Sub

Private Sub cmdSearchRep_Click()
  frmSRTransDate.Show
  Unload frmGLUtilMenu
End Sub

Private Sub cmdTagRec_Click()
  frmTagTrans.Show
  Unload frmGLUtilMenu
End Sub

Private Sub cmdUnTag_Click()
  frmUnTagTrans.Show
  Unload frmGLUtilMenu
End Sub

Private Sub cmdUpdateQuery_Click()
  frmUpdateQuery.Show
  Unload frmGLUtilMenu
End Sub

Private Sub cmdViewGLLog_Click()
  ViewLog
End Sub

Private Sub cmdViewMainLog_Click()
  ViewMainLog
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  Me.HelpContextID = hlpGLUtil
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
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitGLUtilMenu.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        Call MainLog("Close via GL Util.")
        ClearInUse PWcnt
      End If
    End If
  End If
End Sub

Private Sub cmdExitGLUtilMenu_Click()
  Call MainLog("Exit GL Util.")
  frmGLConfigUtilMenu.Show
  Unload frmGLUtilMenu
End Sub
Private Sub DateScan()
  Dim CommaFmt As String, TotalFmt As String, RptTitle As String
  Dim BadDate As Integer, cnt As Integer
  Dim totDebits As Double, totCredits As Double
  Dim LogFile As Integer, LogFileName As String
  Dim FundIdxFileNum As Integer, NumFunds As Integer
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, TRRec As Long
  Dim GLTransFile As Integer, NumTrans As Long
  CommaFmt$ = "#########.##"  'format takes 14 chars
  TotalFmt$ = "#,###,###,###.##" 'format takes 16 chars
  FrmShowPctComp.Label1 = "Printing Dates Out Of Range Report"
  FrmShowPctComp.Show , Me
  DeActivateControls frmGLUtilMenu
  DoEvents
   GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate

   OpenTransFile GLTransFile, NumTrans&

   LogFile = FreeFile
   LogFileName$ = "DATESCAN.LOG"
   Open LogFileName$ For Append As #LogFile

   Print #LogFile, "Date Scan started @ " + Date$ + " @ " + Time$
   Print #LogFile, "Scanning for dates outside of " + Format(DateAdd("d", FY1BegDate, "12-31-1979"), "mm/dd/yy") + " and " + Format(DateAdd("d", FY2EndDate, "12-31-1979"), "mm/dd/yy")
   For TRRec& = 1 To NumTrans&

      Get GLTransFile, TRRec&, GLTrans

      If GLTrans.TRDATE < FY1BegDate Or GLTrans.TRDATE > FY2EndDate Then
         BadDate = BadDate + 1
         GoSub ShowTrans
      End If
    FrmShowPctComp.ShowPctComp TRRec&, NumTrans&
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmGLUtilMenu
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
  Next          'Process next transaction
 
   'LOCATE 4, 1
   'PRINT STRING$(80, " ")
   'LOCATE 4, 1
   '
   ''IF GoodCnt& = NumTrans& THEN
   ''  PRINT "No date problems in transaction file."
   ''  PRINT #LogFile, "No date problems in transaction file."
   ''END IF
   'IF BadDate > 0 THEN
   '   PRINT BadDate; "date(s) are out of range."
   '   PRINT #LogFile, BadDate; "date(s) are out of range."
   '   PRINT #LogFile, USING "Total Debits  : ########,.##"; TotDebits#
   '   PRINT #LogFile, USING "Total Credits : ########,.##"; TotCredits#
   'ELSE
   '   PRINT "No date problems in transaction file."
   '   PRINT #LogFile, "No date problems in transaction file."
   'END IF

  ' PRINT
  ' PRINT "Press any key to continue."
  ' K$ = INPUT$(1)

   Close

   RptTitle$ = "Date Scan"
   ActivateControls frmGLUtilMenu
   ARptErrorLog.Caption = RptTitle$
   ARptErrorLog.GetName LogFileName$
   ARptErrorLog.startrpt
   KillFile LogFileName$
   
   Exit Sub
ShowTrans:
   Print #LogFile, "Record: " + Str(TRRec&)
   Print #LogFile, "Date:   " + Format(DateAdd("d", GLTrans.TRDATE, "12-31-1979"), "mm/dd/yy")
   Print #LogFile, "Desc:   " + GLTrans.Desc
   Print #LogFile, "Dr Amt: " + Str(GLTrans.DrAmt)
   Print #LogFile, "Cr Amt: " + Str(GLTrans.CrAmt)
   Print #LogFile, "Src:    " + GLTrans.Src
   totDebits# = totDebits# + GLTrans.DrAmt
   totCredits# = totCredits# + GLTrans.CrAmt
Return
CancelExit:
  Exit Sub


End Sub

Private Sub DateScan2()
  Dim CommaFmt As String, TotalFmt As String, RptTitle As String
  Dim BadDate As Integer, cnt As Integer
  Dim totDebits As Double, totCredits As Double
  Dim LogFile As Integer, LogFileName As String
  Dim FundIdxFileNum As Integer, NumFunds As Integer
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, TRRec As Long
  Dim GLTransFile As Integer, NumTrans As Long
  CommaFmt$ = "#########.##"  'format takes 14 chars
  TotalFmt$ = "#,###,###,###.##" 'format takes 16 chars
  FrmShowPctComp.Label1 = "Printing Dates Out Of Range Report"
  FrmShowPctComp.Show , Me
  DeActivateControls frmGLUtilMenu
  DoEvents
   GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate

   OpenTransFile GLTransFile, NumTrans&

   LogFile = FreeFile
   LogFileName$ = "DATESCAN.LOG"
   Open LogFileName$ For Append As #LogFile

   Print #LogFile,
   Print #LogFile, "Date Scan started @ " + Date$ + " @ "; Time$
   Print #LogFile, "Scanning for dates outside of " + Format(DateAdd("d", FY1BegDate, "12-31-1979"), "mm/dd/yy") + " and " + Format(DateAdd("d", FY2EndDate, "12-31-1979"), "mm/dd/yy")
   For TRRec& = 1 To NumTrans&

      Get GLTransFile, TRRec&, GLTrans

      If GLTrans.TRDATE < FY1BegDate Or GLTrans.TRDATE > FY2EndDate Then
         BadDate = BadDate + 1
         GoSub ShowTrans
      End If
    FrmShowPctComp.ShowPctComp TRRec&, NumTrans&
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmGLUtilMenu
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
  Next          'Process next transaction

  

 
   'LOCATE 4, 1
   'PRINT STRING$(80, " ")
   'LOCATE 4, 1
   '
   ''IF GoodCnt& = NumTrans& THEN
   ''  PRINT "No date problems in transaction file."
   ''  PRINT #LogFile, "No date problems in transaction file."
   ''END IF
   'IF BadDate > 0 THEN
   '   PRINT BadDate; "date(s) are out of range."
   '   PRINT #LogFile, BadDate; "date(s) are out of range."
   '   PRINT #LogFile, USING "Total Debits  : ########,.##"; TotDebits#
   '   PRINT #LogFile, USING "Total Credits : ########,.##"; TotCredits#
   'ELSE
   '   PRINT "No date problems in transaction file."
   '   PRINT #LogFile, "No date problems in transaction file."
   'END IF

  ' PRINT
  ' PRINT "Press any key to continue."
  ' K$ = INPUT$(1)

   Close

   RptTitle$ = "Date Scan"
   
   ViewPrint LogFileName$, RptTitle$
   KillFile LogFileName$
   ActivateControls frmGLUtilMenu
   Exit Sub
ShowTrans:
   Print #LogFile,
   Print #LogFile, "Record: "; TRRec&
   Print #LogFile, "Date:   "; Format(DateAdd("d", GLTrans.TRDATE, "12-31-1979"), "mm/dd/yy")
   Print #LogFile, "Desc:   "; GLTrans.Desc
   Print #LogFile, "Dr Amt: "; GLTrans.DrAmt
   Print #LogFile, "Cr Amt: "; GLTrans.CrAmt
   Print #LogFile, "Src:    "; GLTrans.Src
   Print #LogFile,
   totDebits# = totDebits# + GLTrans.DrAmt
   totCredits# = totCredits# + GLTrans.CrAmt
Return
CancelExit:
  Exit Sub


End Sub
Private Sub ViewMainLog()
Dim RptTitle As String, RptFileName As String
  frmReportOpt.Show 1
  If Exist("AcctLog.DAT") Then
     RptTitle$ = "G/L LogFile"
     RptFileName$ = "AcctLog.dat"
     If rptopt = 1 Then
       ARptErrorLog.Caption = RptTitle$
       ARptErrorLog.GetName RptFileName$
       ARptErrorLog.startrpt
     ElseIf rptopt = 2 Then
       ViewPrint RptFileName$, RptTitle$
     End If
  Else
    MsgBox "No Log file.  Press any key to continue.", vbOKOnly, "AcctLog.DAT"
  End If

End Sub

Private Sub ViewLog()
Dim RptTitle As String, RptFileName As String
  frmReportOpt.Show 1
  If Exist("GLUTIL.LOG") Then
     RptTitle$ = "G/L Utility Log"
     RptFileName$ = "GLUTIL.LOG"
     If rptopt = 1 Then
        ARptErrorLog.Caption = RptTitle$
        ARptErrorLog.GetName RptFileName$
        ARptErrorLog.startrpt
     ElseIf rptopt = 2 Then
        ViewPrint RptFileName$, RptTitle$
     End If
  Else
    MsgBox "No Log file.  Press any key to continue.", vbOKOnly, "GLUtil.log"
  End If

End Sub
Public Sub PurgePOs(xme As Form, pdate As Integer)
  Dim PO As GLTransRecType
  Dim PY As GLTransRecType
  Dim TotLen As Integer, Yr As String
  Dim TransRecLen As Integer, POTransFile As Integer, NumTrans As Long
  Dim PYTransFile As Integer, PORec As Long
  Dim FundCode As String, ClosingThisFund As Boolean, F As Integer
  Dim PYRec As Long, HistDir As String, tstdir As String
  'QPrintRC "Updating Transaction Files.", 5, 10, 15
  'PrintHelp "Parsing Histories..."
  TotLen% = GLFundLen + GLAcctLen + GLDetLen
  Yr$ = Format(DateAdd("d", (pdate), "12-31-1979"), "mm/dd/yy")
  'Clean up old closing files
  KillFile "POTrans.PY"
  'KillFile "potrans.oyr"
  TransRecLen = Len(PO)
  POTransFile = FreeFile
  Open "POTRANS.DAT" For Random Access Read Write Shared As POTransFile Len = TransRecLen
  NumTrans& = LOF(POTransFile) \ TransRecLen

  PYTransFile = FreeFile
  Open "POTRANS.PY" For Random As PYTransFile Len = TransRecLen
  Call MainLog("Purge Pos via Menu Prior to  - " + Yr$)
  FrmShowPctComp.Label1 = "Updating PO Transaction Files."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , xme
  DoEvents
  For PORec& = 1 To NumTrans&
    FrmShowPctComp.ShowPctComp PORec&, NumTrans&
    Get POTransFile, PORec&, PO
    If PO.TRDATE >= pdate Then
'only copy records for current years
      PYRec& = PYRec& + 1
      LSet PY = PO
      Put PYTransFile, PYRec&, PY
    End If

  Next

  Close
 
  '--save the original transaction file
  KillFileD "POTRANS.OLD"
  Name "potrans.dat" As "potrans.OLD"

  '--rename the potrans.py file to .dat then relink
  Name "potrans.py" As "potrans.dat"
  Call MainLog("Purge completed")
  
End Sub
Private Sub PrnPOTransFile()
  Dim CommaFmt As String, TotalFmt As String, RptTitle As String
  Dim RptFile As Integer
  Dim RptFileName As String, BgtFmt As String, Newrp As String
  Dim POTransFile As Integer, TransRecLen As Integer, NumTrans As Long
  Dim ToPrint As String, Linecnt As Integer, TRRec As Long
  Dim TCnt As Long, Diff As Double, Debits As Double, Credits As Double
  CommaFmt$ = "#########.##"  'format takes 14 chars
  TotalFmt$ = "#,###,###,###.##" 'format takes 16 chars
  BgtFmt$ = "###,###,###"         'format takes 11 chars
  TransRecLen = Len(POTrans)
  POTransFile = FreeFile
  Open "POTRANS.DAT" For Random Shared As POTransFile Len = TransRecLen
  NumTrans& = LOF(POTransFile) \ TransRecLen

  'OPEN "POTRANS.NEw" FOR RANDOM SHARED AS #10 LEN = TransRecLen
  '--open a report file to print to
  RptFile = FreeFile
  Newrp = "POTR"
  GetRPTName Newrp
  RptFileName$ = Newrp
  Open RptFileName$ For Output As RptFile
  FrmShowPctComp.Label1 = "Printing PO Transaction File"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmGLUtilMenu
  For TRRec& = 1 To NumTrans&
    FrmShowPctComp.ShowPctComp TRRec&, NumTrans&
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmGLUtilMenu
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

     Get POTransFile, TRRec&, POTrans
     'IF POTrans.TrDate >= TDate THEN
     '  PUT #10, , POTrans
       ToPrint$ = ""
       ToPrint$ = POTrans.AcctNum
       ToPrint$ = ToPrint$ + "~" + Format(DateAdd("d", (POTrans.TRDATE), "12-31-1979"), "mm/dd/yy")
       ToPrint$ = ToPrint$ + "~" + Left$(POTrans.Desc, 20)
       ToPrint$ = ToPrint$ + "~" + POTrans.Ref
  '     IF INSTR(POTrans.Ref, "3410") > 0 THEN
  '       STOP
  '     END IF
       ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(POTrans.DrAmt))
       ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(POTrans.CrAmt))
       ToPrint$ = ToPrint$ + "~" + POTrans.Src
       ToPrint$ = ToPrint$ + "~" + "Tr#:" + Str$(TRRec&)

       'MID$(ToPrint$, 96) = "Nx:" + STR$(POTrans.NextTran)
       Print #RptFile, ToPrint$

       TCnt& = TCnt& + 1
       Debits# = Round#(Debits# + POTrans.DrAmt)
       Credits# = Round#(Credits# + POTrans.CrAmt)
   ' END IF

   Next          'Process next transaction

  
  Close
  Diff# = Round#(Debits# - Credits#)

  RptTitle$ = "PO Transaction Records"
  ActivateControls frmGLUtilMenu
  Load frmLoadingRpt
  ARptTransQuery.Title.Caption = RptTitle$
  ARptTransQuery.totDeb = Using$(TotalFmt$, Str$(Debits#))
  ARptTransQuery.totCred = Using$(TotalFmt$, Str$(Credits#))
  ARptTransQuery.totBal = Using$(TotalFmt$, Str$(Diff#))
  ARptTransQuery.totRecs = Using$(BgtFmt$, Str$(TCnt&))
  ARptTransQuery.txtDate = Now
  ARptTransQuery.txtTown = GLUserName$
  ARptTransQuery.GetName RptFileName$
  ARptTransQuery.startrpt
  
CancelExit:
  Exit Sub
End Sub


Private Sub PrnPOTransFile2()
  Dim CommaFmt As String, TotalFmt As String, RptTitle As String
  Dim MaxLines As Integer, RptFile As Integer, Pitch12 As String
  Dim RptFileName As String, BgtFmt As String, Newrp As String
  Dim POTransFile As Integer, TransRecLen As Integer, NumTrans As Long
  Dim ToPrint As String, Linecnt As Integer, TRRec As Long
  Dim TCnt As Long, Diff As Double, Debits As Double, Credits As Double
  CommaFmt$ = "#########.##"  'format takes 14 chars
  TotalFmt$ = "#,###,###,###.##" 'format takes 16 chars
  BgtFmt$ = "###,###,###"         'format takes 11 chars
  TransRecLen = Len(POTrans)
  POTransFile = FreeFile
  Open "POTRANS.DAT" For Random Shared As POTransFile Len = TransRecLen
  NumTrans& = LOF(POTransFile) \ TransRecLen

  'OPEN "POTRANS.NEw" FOR RANDOM SHARED AS #10 LEN = TransRecLen
  '--open a report file to print to
  RptFile = FreeFile
  Newrp = "POTR"
  GetRPTName Newrp
  RptFileName$ = Newrp
  Open RptFileName$ For Output As RptFile
  FrmShowPctComp.Label1 = "Printing PO Transaction File"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmGLUtilMenu
  For TRRec& = 1 To NumTrans&
    FrmShowPctComp.ShowPctComp TRRec&, NumTrans&
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmGLUtilMenu
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

     Get POTransFile, TRRec&, POTrans
     'IF POTrans.TrDate >= TDate THEN
     '  PUT #10, , POTrans
       ToPrint$ = Space$(132)
       LSet ToPrint$ = POTrans.AcctNum
       Mid$(ToPrint$, 16) = Format(DateAdd("d", (POTrans.TRDATE), "12-31-1979"), "mm/dd/yy")
       Mid$(ToPrint$, 26) = Left$(POTrans.Desc, 20)
       Mid$(ToPrint$, 47) = POTrans.Ref
  '     IF INSTR(POTrans.Ref, "3410") > 0 THEN
  '       STOP
  '     END IF
       Mid$(ToPrint$, 57) = Using$(CommaFmt$, Str$(POTrans.DrAmt))
       Mid$(ToPrint$, 72) = Using$(CommaFmt$, Str$(POTrans.CrAmt))
       Mid$(ToPrint$, 87) = POTrans.Src
       Mid$(ToPrint$, 96) = "Tr#:" + Str$(TRRec&)

       'MID$(ToPrint$, 96) = "Nx:" + STR$(POTrans.NextTran)
       Print #RptFile, ToPrint$

       TCnt& = TCnt& + 1
       Debits# = Round#(Debits# + POTrans.DrAmt)
       Credits# = Round#(Credits# + POTrans.CrAmt)
   ' END IF

   Next          'Process next transaction

  ActivateControls frmGLUtilMenu

  Diff# = Round#(Debits# - Credits#)
  Print #RptFile,
  Print #RptFile, "File Totals"
  Print #RptFile, "---------------"
  Print #RptFile, "Total Records  : "; Using$(BgtFmt$, Str$(TCnt&))
  Print #RptFile, "Debit Total    : "; Using$(TotalFmt$, Str$(Debits#))
  Print #RptFile, "Credit Total   : "; Using$(TotalFmt$, Str$(Credits#))
  Print #RptFile, "Balance        : "; Using$(TotalFmt$, Str$(Diff#))

  Close


  RptTitle$ = "PO Transaction Records"
  ViewPrint RptFileName$, RptTitle$, True
  Kill RptFileName$
CancelExit:
  Exit Sub
End Sub
Private Sub ListMTrans()
  Dim CommaFmt As String, TotalFmt As String, RptTitle As String
  Dim SumLine As String, PRNFile As Integer, Match As Integer
  Dim TotDr As Double, TotCr As Double, RecNum As String
  Dim RptFile As Integer, Pitch12 As String, AType As String
  Dim RptFileName As String, BgtFmt As String
  Dim FundIdxFileNum As Integer, NumFunds As Integer, Acct As Integer
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, TRRec As Long
  Dim GLTransFile As Integer, NumTrans As Long, Marked As Long
  Dim Pct As String, ToPrint As String, Linecnt As Integer
  Dim TCnt As Long, Diff As Double, Debits As Double, Credits As Double
  CommaFmt$ = "#########.##"  'format takes 14 chars
  TotalFmt$ = "#,###,###,###.##" 'format takes 16 chars

  BgtFmt$ = "###,###,###"         'format takes 11 chars

  OpenTransFile GLTransFile, NumTrans&

  '--open a report file to print to
  RptFile = FreeFile
  RptFileName$ = "MTRLIST.PRN"
  Open RptFileName$ For Output As RptFile
  FrmShowPctComp.Label1 = "Printing Transaction File"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmGLUtilMenu
  For TRRec& = 1 To NumTrans&
''''''     Complete! = (TrRec& / NumTrans&) * 100
''''''     Pct$ = FUsing(Str$(Complete!), "###")
''''''     QPrintRC Pct$, 25, 14, -1
     FrmShowPctComp.ShowPctComp TRRec&, NumTrans&
     Get GLTransFile, TRRec&, GLTrans
     If GLTrans.Marked = True Then
        ToPrint$ = ""
        ToPrint$ = QPTrim(GLTrans.AcctNum)
        ToPrint$ = ToPrint$ + "~" + Format(DateAdd("d", GLTrans.TRDATE, "12-31-1979"), "mm/dd/yy")
        ToPrint$ = ToPrint$ + "~" + QPTrim(Left$(GLTrans.Desc, 20))
        ToPrint$ = ToPrint$ + "~" + QPTrim(GLTrans.Ref)
        ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(GLTrans.DrAmt))
        ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(GLTrans.CrAmt))
        ToPrint$ = ToPrint$ + "~" + QPTrim(GLTrans.Src)
        ToPrint$ = ToPrint$ + "~" + "Tr#:" + Str$(TRRec&)
        ToPrint$ = ToPrint$ + "~" + "Nx:" + Str$(GLTrans.NextTran)
        Print #RptFile, ToPrint$

        Marked& = Marked& + 1

        Debits# = Round#(Debits# + GLTrans.DrAmt)
        Credits# = Round#(Credits# + GLTrans.CrAmt)
        
        If FrmShowPctComp.Out = True Then
          Close
          FrmShowPctComp.Out = False
          ActivateControls frmGLUtilMenu
          Unload FrmShowPctComp
          GoTo CancelExit
        End If
      End If
  Next          'Process next transaction

  
  Diff# = Round#(Debits# - Credits#)

  Close

  ActivateControls frmGLUtilMenu
    Load frmLoadingRpt
  ARptListMTrans.totDeb = Using$(TotalFmt$, Str$(Debits#))
  ARptListMTrans.totCred = Using$(TotalFmt$, Str$(Credits#))
  ARptListMTrans.totBal = Using$(TotalFmt$, Str$(Diff#))
  ARptListMTrans.totRecs = Using$(BgtFmt$, Str$(Marked&))
  ARptListMTrans.txtDate = Now
  ARptListMTrans.txtTown = GLUserName$
  ARptListMTrans.GetName RptFileName$
  ARptListMTrans.startrpt


CancelExit:
  Exit Sub
End Sub
Private Sub ListMTrans2()
  Dim CommaFmt As String, TotalFmt As String, RptTitle As String
  Dim SumLine As String, FF As String, PRNFile As Integer, Match As Integer
  Dim MaxLines As Integer, TotDr As Double, TotCr As Double, RecNum As String
  Dim RptFile As Integer, Pitch12 As String, AType As String
  Dim RptFileName As String, BgtFmt As String
  Dim FundIdxFileNum As Integer, NumFunds As Integer, Acct As Integer
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, TRRec As Long
  Dim GLTransFile As Integer, NumTrans As Long, Marked As Long
  Dim Pct As String, ToPrint As String, Linecnt As Integer
  Dim TCnt As Long, Diff As Double, Debits As Double, Credits As Double
  CommaFmt$ = "#########.##"  'format takes 14 chars
  TotalFmt$ = "#,###,###,###.##" 'format takes 16 chars

  BgtFmt$ = "###,###,###"         'format takes 11 chars


  OpenTransFile GLTransFile, NumTrans&

  '--open a report file to print to
  RptFile = FreeFile
  RptFileName$ = "MTRLIST.PRN"
  Open RptFileName$ For Output As RptFile
  FrmShowPctComp.Label1 = "Printing Transaction File"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmGLUtilMenu
  For TRRec& = 1 To NumTrans&
''''''     Complete! = (TrRec& / NumTrans&) * 100
''''''     Pct$ = FUsing(Str$(Complete!), "###")
''''''     QPrintRC Pct$, 25, 14, -1
     FrmShowPctComp.ShowPctComp TRRec&, NumTrans&
     Get GLTransFile, TRRec&, GLTrans
     If GLTrans.Marked = True Then
        ToPrint$ = Space$(132)
        LSet ToPrint$ = QPTrim(GLTrans.AcctNum)
        Mid$(ToPrint$, 16) = Format(DateAdd("d", GLTrans.TRDATE, "12-31-1979"), "mm/dd/yy")
        Mid$(ToPrint$, 26) = QPTrim(Left$(GLTrans.Desc, 20))
        Mid$(ToPrint$, 47) = QPTrim(GLTrans.Ref)
        Mid$(ToPrint$, 57) = Using$(CommaFmt$, Str$(GLTrans.DrAmt))
        Mid$(ToPrint$, 72) = Using$(CommaFmt$, Str$(GLTrans.CrAmt))
        Mid$(ToPrint$, 87) = QPTrim(GLTrans.Src)
        Mid$(ToPrint$, 96) = "Tr#:" + Str$(TRRec&)
        Mid$(ToPrint$, 110) = "Nx:" + Str$(GLTrans.NextTran)
        Print #RptFile, ToPrint$

        Marked& = Marked& + 1

        Debits# = Round#(Debits# + GLTrans.DrAmt)
        Credits# = Round#(Credits# + GLTrans.CrAmt)
        
        If FrmShowPctComp.Out = True Then
          Close
          FrmShowPctComp.Out = False
          ActivateControls frmGLUtilMenu
          Unload FrmShowPctComp
          GoTo CancelExit
        End If
      End If
  Next          'Process next transaction

  ActivateControls frmGLUtilMenu
  Diff# = Round#(Debits# - Credits#)
  Print #RptFile,
  Print #RptFile, "File Totals"
  Print #RptFile, "---------------"
  Print #RptFile, "Marked Records  : "; Using$(BgtFmt$, Str$(Marked&))
  Print #RptFile, "Debit Total    : "; Using$(TotalFmt$, Str$(Debits#))
  Print #RptFile, "Credit Total   : "; Using$(TotalFmt$, Str$(Credits#))
  Print #RptFile, "Balance        : "; Using$(TotalFmt$, Str$(Diff#))

  Close



  RptTitle$ = "List Marked Transactions"
  'EntryPoint = 2
  ViewPrint RptFileName$, RptTitle$, True
CancelExit:
  Exit Sub
End Sub
Private Sub PrnTransFile()
  Dim CommaFmt As String, TotalFmt As String, RptTitle As String
  Dim SumLine As String, FF As String, PRNFile As Integer, Match As Integer
  Dim MaxLines As Integer, TotDr As Double, TotCr As Double, RecNum As String
  Dim RptFile As Integer, Pitch12 As String, AType As String
  Dim RptFileName As String, BgtFmt As String
  Dim FundIdxFileNum As Integer, NumFunds As Integer, Acct As Integer
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, TRRec As Long
  Dim GLTransFile As Integer, NumTrans As Long
  Dim Pct As String, ToPrint As String, Linecnt As Integer
  Dim TCnt As Long, Diff As Double, Debits As Double, Credits As Double
  CommaFmt$ = "###########.##"  'format takes 14 chars
  TotalFmt$ = "###,###,###,###.##" 'format takes 16 chars

  BgtFmt$ = "###,###,###,###"         'format takes 11 chars

'  If Mid$(fpcboNewExist.Text, 1, 1) = "E" Then
'    If Exist("trlist.prn") Then
'      RptFileName$ = "trlist.prn"
'      GoTo PrintRpt
'    End If
'  End If

  OpenTransFile GLTransFile, NumTrans&

  '--open a report file to print to
  RptFile = FreeFile
  RptFileName$ = "TRLIST.PRN"
  Open RptFileName$ For Output As RptFile
  FrmShowPctComp.Label1 = "Printing Transaction File"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmGLUtilMenu
  For TRRec& = 1 To NumTrans&
''''''     Complete! = (TrRec& / NumTrans&) * 100
''''''     Pct$ = FUsing(Str$(Complete!), "###")
''''''     QPrintRC Pct$, 25, 14, -1

     Get GLTransFile, TRRec&, GLTrans
     
     ToPrint$ = ""
     ToPrint$ = QPTrim(GLTrans.AcctNum) + "~"
     ToPrint$ = ToPrint$ + Format(DateAdd("d", GLTrans.TRDATE, "12-31-1979"), "mm/dd/yy")
     ToPrint$ = ToPrint$ + "~" + QPTrim(Left$(GLTrans.Desc, 20))
     '''ToPrint$ = ToPrint$ + "~" + QPTrim(Left$(GLTrans.LDesc, 20))
     ToPrint$ = ToPrint$ + "~" + QPTrim(Left$(GLTrans.Ref, 9))
     ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, QPTrim(Str$(GLTrans.DrAmt)))
     ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, QPTrim(Str$(GLTrans.CrAmt)))
     ToPrint$ = ToPrint$ + "~" + QPTrim(GLTrans.Src)
     ToPrint$ = ToPrint$ + "~" + "Tr#:" + Str$(TRRec&)
     Print #RptFile, ToPrint$

     TCnt& = TCnt& + 1
    Debits# = Round#(Debits# + GLTrans.DrAmt)
     Credits# = Round#(Credits# + GLTrans.CrAmt)
    FrmShowPctComp.ShowPctComp TRRec&, NumTrans&
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmGLUtilMenu
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
  Next          'Process next transaction



  Diff# = Round#(Debits# - Credits#)

  Close
  ActivateControls frmGLUtilMenu
  Load frmLoadingRpt
  ARptTransQuery.totDeb = Using$(TotalFmt$, Str$(Debits#))
  ARptTransQuery.totCred = Using$(TotalFmt$, Str$(Credits#))
  ARptTransQuery.totBal = Using$(TotalFmt$, Str$(Diff#))
  ARptTransQuery.totRecs = Using$(BgtFmt$, Str$(TCnt&))
  ARptTransQuery.txtDate = Now
  ARptTransQuery.txtTown = GLUserName$
  ARptTransQuery.GetName RptFileName$
  ARptTransQuery.startrpt


CancelExit:
  Exit Sub
End Sub
Private Sub PrnTransFile2()
  Dim CommaFmt As String, TotalFmt As String, RptTitle As String
  Dim SumLine As String, FF As String, PRNFile As Integer, Match As Integer
  Dim MaxLines As Integer, TotDr As Double, TotCr As Double, RecNum As String
  Dim RptFile As Integer, Pitch12 As String, AType As String
  Dim RptFileName As String, BgtFmt As String
  Dim FundIdxFileNum As Integer, NumFunds As Integer, Acct As Integer
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, TRRec As Long
  Dim GLTransFile As Integer, NumTrans As Long
  Dim Pct As String, ToPrint As String, Linecnt As Integer
  Dim TCnt As Long, Diff As Double, Debits As Double, Credits As Double
  CommaFmt$ = "###########.##"  'format takes 14 chars
  TotalFmt$ = "###,###,###,###.##" 'format takes 16 chars

  BgtFmt$ = "###,###,###,###"         'format takes 11 chars

'  If Mid$(fpcboNewExist.Text, 1, 1) = "E" Then
'    If Exist("trlist.prn") Then
'      RptFileName$ = "trlist.prn"
'      GoTo PrintRpt
'    End If
'  End If

  OpenTransFile GLTransFile, NumTrans&

  '--open a report file to print to
  RptFile = FreeFile
  RptFileName$ = "TRLIST.PRN"
  Open RptFileName$ For Output As RptFile
  FrmShowPctComp.Label1 = "Printing Transaction File"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmGLUtilMenu
  For TRRec& = 1 To NumTrans&
''''''     Complete! = (TrRec& / NumTrans&) * 100
''''''     Pct$ = FUsing(Str$(Complete!), "###")
''''''     QPrintRC Pct$, 25, 14, -1

     Get GLTransFile, TRRec&, GLTrans
     
     ToPrint$ = Space$(132)
     LSet ToPrint$ = QPTrim(GLTrans.AcctNum)
     Mid$(ToPrint$, 16) = Format(DateAdd("d", GLTrans.TRDATE, "12-31-1979"), "mm/dd/yy")
     Mid$(ToPrint$, 26) = QPTrim(Left$(GLTrans.Desc, 20))
     Mid$(ToPrint$, 47) = QPTrim(Left$(GLTrans.Ref, 9))
     Mid$(ToPrint$, 57) = Using$(CommaFmt$, QPTrim(Str$(GLTrans.DrAmt)))
     Mid$(ToPrint$, 72) = Using$(CommaFmt$, QPTrim(Str$(GLTrans.CrAmt)))
     Mid$(ToPrint$, 87) = QPTrim(GLTrans.Src)
     Mid$(ToPrint$, 96) = "Tr#:" + Str$(TRRec&)
     'MID$(ToPrint$, 96) = "Nx:" + STR$(Trans.NextTran)
     Print #RptFile, ToPrint$

     TCnt& = TCnt& + 1
     Debits# = Round#(Debits# + GLTrans.DrAmt)
     Credits# = Round#(Credits# + GLTrans.CrAmt)
    FrmShowPctComp.ShowPctComp TRRec&, NumTrans&
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmGLUtilMenu
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
  Next          'Process next transaction


  Diff# = Round#(Debits# - Credits#)
  Print #RptFile,
  Print #RptFile, "File Totals"
  Print #RptFile, "---------------"
  Print #RptFile, "Total Records  : "; Using$(BgtFmt$, Str$(TCnt&))
  Print #RptFile, "Debit Total    : "; Using$(TotalFmt$, Str$(Debits#))
  Print #RptFile, "Credit Total   : "; Using$(TotalFmt$, Str$(Credits#))
  Print #RptFile, "Balance        : "; Using$(TotalFmt$, Str$(Diff#))

  Close
ActivateControls frmGLUtilMenu



  RptTitle$ = "List Transaction Records"
  'EntryPoint = 2
  ViewPrint RptFileName$, RptTitle$, True
CancelExit:
  Exit Sub
End Sub
