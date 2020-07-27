VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.OCX"
Begin VB.Form frmBudPrepMaint 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Budget Preparation Maintenance"
   ClientHeight    =   8640
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12216
   Icon            =   "frmBud.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn fpBtnCopy4 
      Height          =   600
      Left            =   4656
      TabIndex        =   15
      Top             =   7296
      Width           =   636
      _Version        =   131072
      _ExtentX        =   1122
      _ExtentY        =   1058
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
      BackStyle       =   0
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
      ButtonDesigner  =   "frmBud.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn fpBtnCopy3 
      Height          =   600
      Left            =   3600
      TabIndex        =   14
      Top             =   7290
      Width           =   630
      _Version        =   131072
      _ExtentX        =   1111
      _ExtentY        =   1058
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
      BackStyle       =   0
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
      ButtonDesigner  =   "frmBud.frx":1786
   End
   Begin fpBtnAtlLibCtl.fpBtn fpBtnCopy2 
      Height          =   600
      Left            =   2496
      TabIndex        =   10
      Top             =   7296
      Width           =   636
      _Version        =   131072
      _ExtentX        =   1122
      _ExtentY        =   1058
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
      BackStyle       =   0
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
      ButtonDesigner  =   "frmBud.frx":2642
   End
   Begin fpBtnAtlLibCtl.fpBtn fpBtnCopy1 
      Height          =   600
      Left            =   1440
      TabIndex        =   13
      Top             =   7290
      Width           =   630
      _Version        =   131072
      _ExtentX        =   1111
      _ExtentY        =   1058
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
      BackStyle       =   0
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
      ButtonDesigner  =   "frmBud.frx":34FE
   End
   Begin VB.CheckBox chkSelection 
      Caption         =   "Check1"
      Height          =   192
      Left            =   5904
      TabIndex        =   18
      Top             =   7536
      Width           =   204
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "Check2"
      Height          =   204
      Left            =   5904
      TabIndex        =   17
      Top             =   7872
      Width           =   204
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "F3 &Clear Worksheet"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   684
      Left            =   10224
      TabIndex        =   7
      ToolTipText     =   "This Option Will Clear Editable Columns."
      Top             =   576
      Width           =   1788
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   " F8 &Print"
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
      Left            =   7392
      TabIndex        =   6
      Top             =   7632
      Width           =   1356
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export to XLS"
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
      Left            =   336
      TabIndex        =   5
      Top             =   720
      Width           =   1788
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "F10 &Save"
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
      Left            =   9024
      TabIndex        =   4
      Top             =   7632
      Width           =   1356
   End
   Begin VB.CommandButton cmdExit 
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
      Left            =   10656
      TabIndex        =   2
      Top             =   7632
      Width           =   1356
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   8340
      Width           =   12216
      _ExtentX        =   21548
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7154
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "9:44 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "3/25/03"
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
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   5484
      Left            =   288
      TabIndex        =   0
      Top             =   1584
      Width           =   11652
      _Version        =   196613
      _ExtentX        =   20553
      _ExtentY        =   9673
      _StockProps     =   64
      ArrowsExitEditMode=   -1  'True
      BackColorStyle  =   1
      ButtonDrawMode  =   1
      ColsFrozen      =   3
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColor       =   8421504
      MaxCols         =   11
      ProcessTab      =   -1  'True
      SpreadDesigner  =   "frmBud.frx":43BA
      VisibleCols     =   9
      ClipboardOptions=   4
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Print Range"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   444
      Left            =   5808
      TabIndex        =   21
      Top             =   7200
      Width           =   1404
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   348
      Left            =   6240
      TabIndex        =   20
      Top             =   7824
      Width           =   444
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Selection"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   348
      Left            =   6192
      TabIndex        =   19
      Top             =   7488
      Width           =   1020
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   732
      Left            =   5808
      Top             =   7440
      Width           =   1260
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copy Shortcuts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   540
      Left            =   240
      TabIndex        =   16
      Top             =   7344
      Width           =   924
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Recommended    to Approved    "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   444
      Left            =   4464
      TabIndex        =   12
      Top             =   7824
      Width           =   1164
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Estimated to Requested"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   444
      Left            =   2256
      TabIndex        =   11
      Top             =   7824
      Width           =   1164
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Current to Estimated"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   444
      Left            =   1152
      TabIndex        =   9
      Top             =   7824
      Width           =   1212
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Requested to Recommended"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   444
      Left            =   3360
      TabIndex        =   8
      Top             =   7824
      Width           =   1116
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H80000013&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      Height          =   5652
      Left            =   174
      Top             =   1488
      Width           =   11868
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   852
      Left            =   3120
      Top             =   384
      Width           =   6492
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Budget Preparation Worksheet"
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
      Left            =   3984
      TabIndex        =   3
      Top             =   624
      Width           =   4812
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   972
      Left            =   3120
      Top             =   288
      Width           =   6492
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
Attribute VB_Name = "frmBudPrepMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim GLAcctidx As GLAcctIndexType
Dim GLAcct As GLAcctRecType
Dim tempstr As String, lblNum As Integer, TempPrint As String
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim Fund As String, SetFund1 As String, SetFund2 As String, Acctcd As String, Detcd As String
Dim Revrange As Integer, Exprange As Integer
'Added column 11 to spreadsheet for gltype R-revenue, E-expense
'but for now not using - need to fix for way to calculate total of
'each and difference for total , for now column is just hidden.3-25-03 PKS
Private Sub cmdClear_Click()
  If MsgBox("Are You Sure You Wish to Clear Worksheet?", vbYesNo, "Budget Preparation") = vbYes Then
    vaSpread1.ClearRange 7, 1, 10, (vaSpread1.DataRowCnt - 1), True
    Call MainLog("BudPrep Cleared.")
  End If
End Sub
'This is to fix spreadsheet for various resolutions
Public Function Fixspread()
    Select Case screenW
      Case 1280
      If Screen.TwipsPerPixelX <> 12 Then
        coladj = 10
        vaSpread1.RowHeight(-1) = 19
      Else
        coladj = 6.1
        vaSpread1.RowHeight(-1) = 18.5
      End If
      Case 1152
      If Screen.TwipsPerPixelX <> 12 Then
        coladj = 7
        vaSpread1.RowHeight(-1) = 17
      Else
        coladj = 4.6
        vaSpread1.RowHeight(-1) = 15.5
      End If
      Case 1024
      If Screen.TwipsPerPixelX <> 12 Then
        coladj = 5
        vaSpread1.RowHeight(-1) = 18
      Else
        coladj = 3.1
      End If
      Case 800
        coladj = 2.9
        'vaSpread1.Font.Size = 8
        vaSpread1.RowHeight(-1) = 14
      Case Else
        'don't worry be happpy
    End Select
    vaSpread1.ColWidth(-1) = vaSpread1.ColWidth(-1) + coladj
End Function

Private Sub cmdExport_Click()
'****This is where will need to fix Path or shortcut for excel
 ' this is only temp to see if works
  If MsgBox("Would You Like To Save Your Work Before Exporting?", vbYesNo, "Budget Preparation") = vbYes Then
    SaveBudPrep
  End If
  vaSpread1.Protect = False
  
  vaSpread1.ExportToExcel "BudPrep", "BudPrep", "BUDlOG.TXT"
  vaSpread1.Protect = True
  Call MainLog("BudPrep ExporttoExcel.")
  MsgBox "The File 'BudPrep.xls' Has Been Created In Your Citipak Folder", vbOKOnly, "Export Complete"
'  vaSpread1.ExportToExcel "C:\BudPrep", "BudPrep", ""
'  frmBudPrepMaint.WindowState = vbMinimized
'  Shell "c:\program files\microsoft office\office\Excel C:\BudPrep"
End Sub
Private Sub chkAll_Click()
  If chkAll.Value = vbChecked Then
    chkSelection.Value = vbUnchecked
  Else
    chkSelection.Value = vbChecked
  End If
End Sub


Private Sub chkSelection_Click()
  If chkSelection.Value = vbChecked Then
    chkAll.Value = vbUnchecked
  Else
    chkAll.Value = vbChecked
  End If
End Sub

Public Sub SetPrinter(DefPrinter)
Dim P As Printer
  For Each P In Printers
    If P.Port = DefPrinter Then
      ' Set printer as system default.
      Set Printer = P
      ' Stop looking for a printer.
      Exit For
    End If
  Next
End Sub
Private Sub cmdPrint_Click()
  If MsgBox("Would You Like To Save Your Work Before Printing?", vbYesNo, "Budget Preparation") = vbYes Then
    SaveBudPrep
  End If
  '***Tried to set default printer but no work, system default and vaspread
  '****default 2 diff things.
  '****TempPrint was to save default before changing. Might use elsewhere.
  'TempPrint = Printer.Port
  'frmPrint2.Show 1
  'SetPrinter TempPrint
    PrintLandscp
  
End Sub
Public Sub PrintLandscp()
  Dim x As Long, c As Variant, r As Variant, c2 As Variant, r2 As Variant
  Dim PRange As Integer
  On Error GoTo Errorhand
  If chkSelection.Value <> 0 Then
    PRange = 1
  Else
    PRange = 0
  End If
  'For CopyLoop = 1 To Copies
  If PRange = 0 Then
    vaSpread1.PrintType = PrintTypeAll
    'vaSpread1.PrintOrientation = POrient
  Else
    'vaSpread1.PrintOrientation = POrient
    vaSpread1.PrintType = PrintTypeCellRange
    vaSpread1.GetSelection x, c, r, c2, r2
    vaSpread1.Row = r
    vaSpread1.Col = c
    vaSpread1.Row2 = r2
    vaSpread1.Col2 = c2
  End If
  vaSpread1.PrintColHeaders = True
  
  vaSpread1.PrintHeader = Now & "/c Budget Prep Worksheet  Page /p of /pc"
  
  'vaSpread1.PrintHeader = True
  'If CopyLoop = 1 Then
   ' vaSpread1.PrintSheet 1
  ' vaSpread1.PrintSmartPrint = True
   frmViewBud.Show
  'Else
    'vaSpread1.PrintSheet
  'End If
  'Next CopyLoop
  'Printer.EndDoc
  Exit Sub
Errorhand:
  MsgBox "There was an error during the printing operation.", vbOKOnly, "Printing Error"
  Exit Sub
End Sub

Private Sub cmdSave_Click()
  If MsgBox("Are You Sure You Are Ready To Save The Budget?", vbOKCancel, "Save Budget") = vbOK Then
    SaveBudPrep
    MsgBox "Budget Preparation Saved.", vbOKOnly, "Budget Saved"
    frmBudgetMaintMenu.Show
    Unload frmBudPrepMaint
  End If
End Sub


Private Sub Form_Load()
  Dim cnt As Integer
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  Fixspread
  BudPrepFillForm
  chkAll.Value = 1
  chkSelection.Value = 0
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub cmdExit_Click()
  If MsgBox("Have You Made Changes You Wish To Save?", vbYesNo, "Budget Preparation") = vbYes Then
    SaveBudPrep
    MsgBox "Budget Preparation Saved.", vbOKOnly, "Budget Saved"
  End If
  frmBudgetMaintMenu.Show
  Unload frmBudPrepMaint
End Sub

Public Function BudPrepFillForm()
  Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer, MaxNum As Integer
  Dim AcctFile As Integer, NumAccts As Integer, CntA As Integer, Num As Integer
  Dim AcctNumber As String, Dept As String, Det As String
  OpenAcctFile AcctFile
  OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
  NumAccts = LOF(AcctFile) / Len(GLAcct)
  For CntA = 1 To NumAIdxRecs
    Get AcctIdxFileNum, CntA, GLAcctidx
    AcctNumber$ = QPTrim$(GLAcctidx.AcctNum)
    Fund$ = Left$(AcctNumber$, GLFundLen)
    Dept$ = Mid$(AcctNumber$, GLFundLen + 2, GLAcctLen)
    Det$ = Right$(AcctNumber$, GLDetLen)

    Get AcctFile, GLAcctidx.RecNum, GLAcct
    If GLAcct.Deleted = 0 Then
      If GLAcct.Typ = "E" Or GLAcct.Typ = "R" Then
        If Fund$ >= SetFund1 And Fund$ <= SetFund2 Then
          If InStr(Dept$, Acctcd$) And InStr(Det$, Detcd$) Then
            MaxNum = MaxNum + 1
          End If
        End If
      End If
    End If
  Next
  vaSpread1.MaxRows = MaxNum + 3
  MaxNum = 0


  For CntA = 1 To NumAIdxRecs
    Get AcctIdxFileNum, CntA, GLAcctidx
    AcctNumber$ = QPTrim$(GLAcctidx.AcctNum)
    Fund$ = Left$(AcctNumber$, GLFundLen)
    Dept$ = Mid$(AcctNumber$, GLFundLen + 2, GLAcctLen)
    Det$ = Right$(AcctNumber$, GLDetLen)

    Get AcctFile, GLAcctidx.RecNum, GLAcct
    If GLAcct.Deleted = 0 Then
      If GLAcct.Typ = "E" Or GLAcct.Typ = "R" Then
        If Fund$ >= SetFund1 And Fund$ <= SetFund2 Then
          If InStr(Dept$, Acctcd$) And InStr(Det$, Detcd$) Then

          vaSpread1.Row = vaSpread1.DataRowCnt + 1
          vaSpread1.Col = 1
          vaSpread1.Text = GLAcctidx.RecNum
          vaSpread1.Col = 2
          vaSpread1.Text = QPTrim(GLAcct.Num)
          vaSpread1.Col = 3
          vaSpread1.Text = GLAcct.Title
          vaSpread1.Col = 4
          vaSpread1.Text = GLAcct.Bgt
          vaSpread1.Col = 5
          vaSpread1.Text = GLAcct.YTD
          vaSpread1.Col = 6
          vaSpread1.Text = GLAcct.PYAct
          vaSpread1.Col = 7
          vaSpread1.Text = GLAcct.NYEst
          vaSpread1.Col = 8
          vaSpread1.Text = GLAcct.NYReq
          vaSpread1.Col = 9
          vaSpread1.Text = GLAcct.NYRec
          vaSpread1.Col = 10
          vaSpread1.Text = GLAcct.NYApp
'see top for notes on column 11
          vaSpread1.Col = 11
          vaSpread1.Text = GLAcct.Typ
          If GLAcct.Typ = "R" Then
            Revrange = vaSpread1.Row
          Else
            Exprange = vaSpread1.Row
          End If
'''        ElseIf Mid$(GLAcct.Num, 1, GLFundLen) = SetFund Then
'''          vaSpread1.Row = vaSpread1.DataRowCnt + 1
'''          vaSpread1.Col = 1
'''          vaSpread1.Text = GLAcctidx.RecNum
'''          vaSpread1.Col = 2
'''          vaSpread1.Text = QPTrim(GLAcct.Num)
'''          vaSpread1.Col = 3
'''          vaSpread1.Text = GLAcct.title
'''          vaSpread1.Col = 4
'''          vaSpread1.Text = GLAcct.Bgt
'''          vaSpread1.Col = 5
'''          vaSpread1.Text = GLAcct.YTD
'''          vaSpread1.Col = 6
'''          vaSpread1.Text = GLAcct.PYAct
'''          vaSpread1.Col = 7
'''          vaSpread1.Text = GLAcct.NYEst
'''          vaSpread1.Col = 8
'''          vaSpread1.Text = GLAcct.NYReq
'''          vaSpread1.Col = 9
'''          vaSpread1.Text = GLAcct.NYRec
'''          vaSpread1.Col = 10
'''          vaSpread1.Text = GLAcct.NYApp
          End If
        End If
      End If
    End If
  Next
  Close AcctFile
  Close AcctIdxFileNum
  MaxNum = vaSpread1.DataRowCnt
  Num = 1
  vaSpread1.Row = MaxNum + 1
  vaSpread1.Col = 3
  vaSpread1.Text = "Totals:"
  vaSpread1.Col = 4
  vaSpread1.Formula = "Sum(R1C:R[-1]C)"
  vaSpread1.Col = 5
  vaSpread1.Formula = "Sum(R1C:R[-1]C)"
  vaSpread1.Col = 6
  vaSpread1.Formula = "Sum(R1C:R[-1]C)"
  vaSpread1.Col = 7
  vaSpread1.Formula = "Sum(R1C:R[-1]C)"
  vaSpread1.Col = 8
  vaSpread1.Formula = "Sum(R1C:R[-1]C)"
  vaSpread1.Col = 9
  vaSpread1.Formula = "Sum(R1C:R[-1]C)"
  vaSpread1.Col = 10
  vaSpread1.Formula = "Sum(R1C:R[-1]C)"
  
  vaSpread1.Row = vaSpread1.DataRowCnt
  vaSpread1.Row2 = vaSpread1.DataRowCnt
  vaSpread1.Col = 1
  vaSpread1.Col2 = 11
  vaSpread1.BlockMode = True
  'Marked the columns to lock in design then this actually locks them
  vaSpread1.Lock = True
  vaSpread1.Protect = True
  vaSpread1.BlockMode = False
'  totRevs
End Function
'see top for notes on future use for column 11 to total revs and expenses
'to get total of difference
'Private Function totRevs()
'  Dim MaxNum As Integer, Num As Integer
'  Dim revb As Integer, reve As Integer, expb As Integer, expe As Integer
'  If Revrange < Exprange Then
'    revb = 1
'    reve = Revrange
'    expb = Revrange + 1
'    expe = Exprange
'  Else
'    expb = 1
'    expe = Exprange
'    revb = Exprange + 1
'    reve = Revrange
'  End If
'  MaxNum = vaSpread1.DataRowCnt
'  Num = 1
'  vaSpread1.Row = MaxNum + 2
'  vaSpread1.Col = 3
'  vaSpread1.Text = "Total Revenues:"
'  vaSpread1.Col = 4
'  vaSpread1.Formula = "Sum(Rrevb:Rreve)"
'  vaSpread1.Col = 5
'  vaSpread1.Formula = "Sum(RrevbC:RreveC)"
'  vaSpread1.Col = 6
'  vaSpread1.Formula = "Sum(R1C:R[-1]C)"
'  vaSpread1.Col = 7
'  vaSpread1.Formula = "Sum(R1C:R[-1]C)"
'  vaSpread1.Col = 8
'  vaSpread1.Formula = "Sum(R1C:R[-1]C)"
'  vaSpread1.Col = 9
'  vaSpread1.Formula = "Sum(R1C:R[-1]C)"
'  vaSpread1.Col = 10
'  vaSpread1.Formula = "Sum(R1C:R[-1]C)"
'  vaSpread1.Row = vaSpread1.DataRowCnt
'  vaSpread1.Row2 = vaSpread1.DataRowCnt
'  vaSpread1.Col = 1
'  vaSpread1.Col2 = 11
'  vaSpread1.BlockMode = True
'  'Marked the columns to lock in design then this actually locks them
'  vaSpread1.Lock = True
'  vaSpread1.Protect = True
'  vaSpread1.BlockMode = False
'
'End Function
Public Function SetOptions(Fund1 As String, Fund2 As String, AcctCode As String, Detcode As String)
  SetFund1 = QPTrim(Fund1)
  SetFund2 = QPTrim(Fund2)
  Acctcd = AcctCode
  Detcd = Detcode
End Function
Private Sub SaveBudPrep()
  Dim cnt As Integer, AcctFileNum As Integer, NumAccts As Integer, Rec As Integer
  OpenAcctFile AcctFileNum
  NumAccts = LOF(AcctFileNum) / Len(GLAcct)
  For cnt = 1 To vaSpread1.MaxRows
    vaSpread1.Row = cnt
    vaSpread1.Col = 1
    If vaSpread1.Text <> "" Then
      Rec = vaSpread1.Text
      Get AcctFileNum, Rec, GLAcct
      If GLAcct.Deleted = 0 Then
        vaSpread1.Col = 7
        If vaSpread1.Text <> "" Then
          GLAcct.NYEst = vaSpread1.Text
        Else
          GLAcct.NYEst = 0
        End If
        vaSpread1.Col = 8
        If vaSpread1.Text <> "" Then
          GLAcct.NYReq = vaSpread1.Text
        Else
          GLAcct.NYReq = 0
        End If
        vaSpread1.Col = 9
        If vaSpread1.Text <> "" Then
          GLAcct.NYRec = vaSpread1.Text
        Else
          GLAcct.NYRec = 0
        End If
        vaSpread1.Col = 10
        If vaSpread1.Text <> "" Then
          GLAcct.NYApp = vaSpread1.Text
        Else
          GLAcct.NYApp = 0
        End If
        Put AcctFileNum, Rec, GLAcct
      End If
    End If
  Next
  'mainlog updates log file if used save command or saved on exit
  Call MainLog("BudPrep Saved.")
  Close AcctFileNum
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
      SendKeys "%S"
      KeyCode = 0
    Case vbKeyF8:
      SendKeys "%P"
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      If MsgBox("Save Information?", vbYesNo, "Save?") = vbYes Then
        SaveBudPrep
        MsgBox "Budget Preparation Saved.", vbOKOnly, "Budget Saved"
      End If
      ClearInUse PWcnt
    End If
  End If
End Sub

Private Sub fpBtnCopy1_Click()
  If MsgBox("This Action Will Change The Amounts In The Estimated Column, Continue -'OK' or Cancel", vbOKCancel, "Budget Preparation") = vbOK Then
    If MsgBox("Do You Wish Increase The Figures That Are Being Copied By A Percentage?", vbYesNo, "Budget Preparation") = vbYes Then
      frmBudIncCopy.Show 1
      lblNum = 1
      frmBudIncCopy.lbl1.Visible = True
    Else
      vaSpread1.CopyRange 4, 1, 4, vaSpread1.DataRowCnt - 1, 7, 1
      Call MainLog("BudPrep Copied To Estimated.")
      fpBtnCopy1.BackColor = &H80000003
    End If
  End If
End Sub
Private Sub fpBtnCopy2_Click()
  If MsgBox("This Action Will Change The Amounts In The Requested Column, Continue -'OK' or Cancel", vbOKCancel, "Budget Preparation") = vbOK Then
    If MsgBox("Do You Wish Increase The Figures That Are Being Copied By A Percentage?", vbYesNo, "Budget Preparation") = vbYes Then
      frmBudIncCopy.Show 1
      lblNum = 2
      frmBudIncCopy.lbl2.Visible = True
    Else
      vaSpread1.CopyRange 7, 1, 7, vaSpread1.DataRowCnt - 1, 8, 1
      Call MainLog("BudPrep Copied To Requested.")
      fpBtnCopy2.BackColor = &H80000003
    End If
  End If
End Sub
Private Sub fpBtnCopy3_Click()
  If MsgBox("This Action Will Change The Amounts In The Recommended Column, Continue -'OK' or Cancel", vbOKCancel, "Budget Preparation") = vbOK Then
    If MsgBox("Do You Wish Increase The Figures That Are Being Copied By A Percentage?", vbYesNo, "Budget Preparation") = vbYes Then
      frmBudIncCopy.Show 1
      lblNum = 3
      frmBudIncCopy.lbl3.Visible = True
    Else
      vaSpread1.CopyRange 8, 1, 8, vaSpread1.DataRowCnt - 1, 9, 1
      Call MainLog("BudPrep Copied To Recommended.")
      fpBtnCopy3.BackColor = &H80000003
    End If
  End If
End Sub
Private Sub fpBtnCopy4_Click()
  If MsgBox("This Action Will Change The Amounts In The Approved Column, Continue -'OK' or Cancel", vbOKCancel, "Budget Preparation") = vbOK Then
    If MsgBox("Do You Wish Increase The Figures That Are Being Copied By A Percentage?", vbYesNo, "Budget Preparation") = vbYes Then
      frmBudIncCopy.Show 1
      lblNum = 4
      frmBudIncCopy.lbl4.Visible = True
    Else
      vaSpread1.CopyRange 9, 1, 9, vaSpread1.DataRowCnt - 1, 10, 1
      Call MainLog("BudPrep Copied To Approved.")
      fpBtnCopy4.BackColor = &H80000003
    End If
  End If
End Sub
Public Function CalcPer(AmtPer As Double, fpcboIncDec As String)
Dim cnt As Integer, NewAmt As Double
If lblNum = 1 Then
  fpBtnCopy1.BackColor = &H80000003
  For cnt = 1 To vaSpread1.DataRowCnt - 1
    vaSpread1.Col = 4
    vaSpread1.Row = cnt
    NewAmt = Round(vaSpread1.Text * AmtPer)
    If fpcboIncDec = "Increase" Then
      NewAmt = Round(NewAmt + vaSpread1.Text)
    Else
      NewAmt = Round(vaSpread1.Text - NewAmt)
    End If
    vaSpread1.Col = 7
    vaSpread1.Text = NewAmt
  Next
ElseIf lblNum = 2 Then
  fpBtnCopy2.BackColor = &H80000003
  For cnt = 1 To vaSpread1.DataRowCnt - 1
    vaSpread1.Col = 7
    vaSpread1.Row = cnt
    NewAmt = Round(vaSpread1.Text * AmtPer)
    If fpcboIncDec = "Increase" Then
      NewAmt = Round(NewAmt + vaSpread1.Text)
    Else
      NewAmt = Round(vaSpread1.Text - NewAmt)
    End If
    vaSpread1.Col = 8
    vaSpread1.Text = NewAmt
  Next
ElseIf lblNum = 3 Then
  fpBtnCopy3.BackColor = &H80000003
  For cnt = 1 To vaSpread1.DataRowCnt - 1
    vaSpread1.Col = 8
    vaSpread1.Row = cnt
    NewAmt = Round(vaSpread1.Text * AmtPer)
    If fpcboIncDec = "Increase" Then
      NewAmt = Round(NewAmt + vaSpread1.Text)
    Else
      NewAmt = Round(vaSpread1.Text - NewAmt)
    End If
    vaSpread1.Col = 9
    vaSpread1.Text = NewAmt
  Next
ElseIf lblNum = 4 Then
  fpBtnCopy4.BackColor = &H80000003
  For cnt = 1 To vaSpread1.DataRowCnt - 1
    vaSpread1.Col = 9
    vaSpread1.Row = cnt
    NewAmt = Round(vaSpread1.Text * AmtPer)
    If fpcboIncDec = "Increase" Then
      NewAmt = Round(NewAmt + vaSpread1.Text)
    Else
      NewAmt = Round(vaSpread1.Text - NewAmt)
    End If
    vaSpread1.Col = 10
    vaSpread1.Text = NewAmt
  Next
End If
End Function

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

'  'CommonDialog1.CancelError = True
'  On Error GoTo Cancel
'  CommonDialog1.PrinterDefault = True
'  CommonDialog1.Flags = &H100000 Or &H4 Or &H40000 Or &H200
'  CommonDialog1.ShowPrinter
'  Printer.Orientation = CommonDialog1.Orientation
  'DefPrinter = Printer.Port
'    LPTHandle = FreeFile
' ' For CopyLoop = 1 To CommonDialog1.Copies
'    Open DefPrinter For Output As LPTHandle
'    RptHandle = FreeFile
'    Open strReportFile For Input As RptHandle
'    Do
'      Line Input #RptHandle, ToPrint$
'      Print #LPTHandle, ToPrint$
'    Loop Until EOF(RptHandle)
'    Close LPTHandle, RptHandle
'  'Next CopyLoop
'  'Printer.EndDoc
'Cancel:

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
