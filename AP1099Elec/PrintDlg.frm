VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PrintDlg 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print"
   ClientHeight    =   2352
   ClientLeft      =   2160
   ClientTop       =   2028
   ClientWidth     =   6528
   ForeColor       =   &H00000000&
   Icon            =   "PrintDlg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2352
   ScaleWidth      =   6528
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Apply"
      Height          =   375
      Left            =   3488
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1848
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Print"
      Height          =   375
      Left            =   816
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1848
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Printer &Setup"
      Height          =   375
      Left            =   2152
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1848
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00D0D0D0&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4824
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1848
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Print Options"
      ForeColor       =   &H00000000&
      Height          =   1515
      Left            =   2820
      TabIndex        =   10
      Top             =   180
      Width           =   3615
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Shadows"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   1920
         TabIndex        =   17
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Color"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   16
         Top             =   900
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Data Cells Only"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   15
         Top             =   600
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Border"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   14
         Top             =   300
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Grid Lines"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   900
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Row Headers"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Value           =   1  'Checked
         Width           =   1635
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Column Headers"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   300
         Value           =   1  'Checked
         Width           =   1635
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Page Range"
      Height          =   1515
      Left            =   72
      TabIndex        =   1
      Top             =   180
      Width           =   2715
      Begin VB.CommandButton Command1 
         BackColor       =   &H00D0D0D0&
         Caption         =   "Setup"
         Height          =   315
         Left            =   1752
         MaskColor       =   &H00D0D0D0&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   132
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1620
         TabIndex        =   8
         Text            =   "1"
         Top             =   1116
         Width           =   315
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   6
         Text            =   "1"
         Top             =   1116
         Width           =   315
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pages"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   168
         TabIndex        =   5
         Top             =   1140
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Current Page"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   168
         TabIndex        =   4
         Top             =   876
         Width           =   1395
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Selected Cells"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   168
         TabIndex        =   3
         Top             =   612
         Width           =   1395
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Full Spreadsheet"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   168
         TabIndex        =   2
         Top             =   348
         Value           =   -1  'True
         Width           =   1608
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "to"
         ForeColor       =   &H00000000&
         Height          =   252
         Left            =   1380
         TabIndex        =   7
         Top             =   1176
         Width           =   432
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   48
      Top             =   1392
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "PrintDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This is sample from fp
Private Sub Command1_Click()
    pagesetup.Show
End Sub

Private Sub Command2_Click()
  PrintSpread
End Sub
Private Sub Command6_Click()
    PrintSpread
    Screen.MousePointer = 11
    frmBudPrepMaint.vaSpread1.PrintSheet
    Screen.MousePointer = 0
    Unload frmViewBud
    Unload Me
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub
Sub PrintSpread()
'Set printing options for spreadsheet
    frmBudPrepMaint.vaSpread1.PrintColHeaders = Check1(0).Value
    frmBudPrepMaint.vaSpread1.PrintRowHeaders = Check1(1).Value
    frmBudPrepMaint.vaSpread1.PrintBorder = Check1(3).Value
    frmBudPrepMaint.vaSpread1.PrintColor = Check1(5).Value
    frmBudPrepMaint.vaSpread1.PrintGrid = Check1(2).Value
    frmBudPrepMaint.vaSpread1.PrintShadows = Check1(6).Value
    frmBudPrepMaint.vaSpread1.PrintUseDataMax = Check1(4).Value
'    frmBudPrepMaint.vaSpread1.Font.Name = CommonDialog1.fontname
'    frmBudPrepMaint.vaSpread1.Font.Size = CommonDialog1.FontSize
'    frmBudPrepMaint.vaSpread1.Font.Bold = CommonDialog1.FontBold
'    frmBudPrepMaint.vaSpread1.Font.Italic = CommonDialog1.FontItalic
'    frmBudPrepMaint.vaSpread1.Font.Underline = CommonDialog1.FontUnderline
'    frmBudPrepMaint.vaSpread1.FontStrikethru = CommonDialog1.FontStrikethru

'Page Range
    'All
    If Option1(0).Value = True Then
        frmBudPrepMaint.vaSpread1.PrintType = PrintTypeAll
        
    'Selected cells
    ElseIf Option1(1).Value = True Then
        frmBudPrepMaint.vaSpread1.Col = frmBudPrepMaint.vaSpread1.SelBlockCol
        frmBudPrepMaint.vaSpread1.Col2 = frmBudPrepMaint.vaSpread1.SelBlockCol2
        frmBudPrepMaint.vaSpread1.Row = frmBudPrepMaint.vaSpread1.SelBlockRow
        frmBudPrepMaint.vaSpread1.Row2 = frmBudPrepMaint.vaSpread1.SelBlockRow2
        frmBudPrepMaint.vaSpread1.PrintType = PrintTypeCellRange
    'Current Page
    ElseIf Option1(2).Value = True Then
        frmBudPrepMaint.vaSpread1.PrintType = PrintTypeCurrentPage
        
    'Pages
    Else
        frmBudPrepMaint.vaSpread1.PrintPageStart = CInt(Text1(0).Text)
        frmBudPrepMaint.vaSpread1.PrintPageEnd = CInt(Text1(1).Text)
        frmBudPrepMaint.vaSpread1.PrintType = PrintTypePageRange
    End If
    frmViewBud.vaSpreadPreview1.hWndSpread = frmBudPrepMaint.vaSpread1.hwnd
    
    'Print control
    
End Sub

Private Sub Command4_Click()
    CommonDialog1.ShowPrinter
End Sub

Private Sub Form_Load()
  Me.Option1(1).Value = True
End Sub

'Private Sub Command5_Click()
'  CommonDialog1.Flags = cdlCFPrinterFonts Or cdlCFFixedPitchOnly
'  CommonDialog1.ShowFont
'
'End Sub


Private Sub Option1_Click(Index As Integer)
    If Index = 3 Then
        Text1(0).Enabled = True
        Text1(1).Enabled = True
        Text1(0).SetFocus
    Else
        Text1(0).Enabled = False
        Text1(1).Enabled = False
    End If
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Verify if a numeric number
    
    If Not IsNumeric(Text1(Index)) Then
        Text1(Index).Text = "1"
    End If
End Sub
