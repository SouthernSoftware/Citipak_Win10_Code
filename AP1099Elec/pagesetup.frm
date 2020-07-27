VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form pagesetup 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Setup"
   ClientHeight    =   4800
   ClientLeft      =   4272
   ClientTop       =   2232
   ClientWidth     =   4464
   ForeColor       =   &H00000000&
   Icon            =   "pagesetup.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4800
   ScaleWidth      =   4464
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Header / Footer Display"
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   180
      TabIndex        =   17
      Top             =   2280
      Width           =   4095
      Begin VB.CommandButton Command1 
         BackColor       =   &H00D0D0D0&
         Caption         =   "Change Header/Footer Attributes"
         Height          =   375
         Left            =   600
         MaskColor       =   &H00D0D0D0&
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   2955
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00D0D0D0&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   3000
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00D0D0D0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1560
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Page Margins (inch)"
      ForeColor       =   &H00000000&
      Height          =   975
      Index           =   1
      Left            =   180
      TabIndex        =   3
      Top             =   1200
      Width           =   4035
      Begin MSMask.MaskEdBox pagemargin 
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   11
         Top             =   240
         Width           =   495
         _ExtentX        =   868
         _ExtentY        =   445
         _Version        =   393216
         MaxLength       =   4
         Mask            =   "#.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox pagemargin 
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   12
         Top             =   540
         Width           =   495
         _ExtentX        =   868
         _ExtentY        =   445
         _Version        =   393216
         MaxLength       =   4
         Mask            =   "#.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox pagemargin 
         Height          =   255
         Index           =   2
         Left            =   2700
         TabIndex        =   13
         Top             =   240
         Width           =   495
         _ExtentX        =   868
         _ExtentY        =   445
         _Version        =   393216
         MaxLength       =   4
         Mask            =   "#.##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox pagemargin 
         Height          =   255
         Index           =   3
         Left            =   2700
         TabIndex        =   14
         Top             =   540
         Width           =   495
         _ExtentX        =   868
         _ExtentY        =   445
         _Version        =   393216
         MaxLength       =   4
         Mask            =   "#.##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bottom:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   2100
         TabIndex        =   7
         Top             =   540
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Top:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   2100
         TabIndex        =   6
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Right:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   540
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Left:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Preview Zoom"
      ForeColor       =   &H00000000&
      Height          =   855
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   3300
      Width           =   4095
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   2115
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Zoom:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   420
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Page Orientation"
      Height          =   975
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   4035
      Begin VB.OptionButton porientation 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Landscape"
         Height          =   195
         Index           =   1
         Left            =   2700
         TabIndex        =   16
         Top             =   480
         Width           =   1155
      End
      Begin VB.OptionButton porientation 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Portrait"
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   15
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   312
         Index           =   1
         Left            =   2100
         Picture         =   "pagesetup.frx":08CA
         Top             =   360
         Width           =   396
      End
      Begin VB.Image Image1 
         Height          =   384
         Index           =   0
         Left            =   480
         Picture         =   "pagesetup.frx":1334
         Top             =   360
         Width           =   324
      End
   End
End
Attribute VB_Name = "pagesetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'got this from sample from fp
Private Sub Command1_Click()
    headerfooter.Show 1
End Sub

Private Sub Command2_Click(Index As Integer)
    'OK button
    If Index = 0 Then
        frmViewBud.GetZoom Combo1.ListIndex
        'Update margins
        frmBudPrepMaint.vaSpread1.PrintMarginTop = CDbl(pagemargin(2).Text) * 1440
        frmBudPrepMaint.vaSpread1.PrintMarginBottom = CDbl(pagemargin(3).Text) * 1440
        frmBudPrepMaint.vaSpread1.PrintMarginLeft = CDbl(pagemargin(0).Text) * 1440
        frmBudPrepMaint.vaSpread1.PrintMarginRight = CDbl(pagemargin(1).Text) * 1440
        
        'Change the page orientation
        'Portrait
        If porientation(0).Value = True Then
            frmBudPrepMaint.vaSpread1.PrintOrientation = PrintOrientationPortrait
        'Landscape
        Else
            frmBudPrepMaint.vaSpread1.PrintOrientation = PrintOrientationLandscape
        End If
       
        'set zoom attributes
        zoomindex = Combo1.ListIndex
    End If
    
    Unload Me
End Sub


Private Sub Form_Load()
    
    'Get page margins (convert to inches) and format
    pagemargin(0).Text = Format(frmBudPrepMaint.vaSpread1.PrintMarginLeft / 1440, "0.00")
    pagemargin(1).Text = Format(frmBudPrepMaint.vaSpread1.PrintMarginRight / 1440, "0.00")
    pagemargin(2).Text = Format(frmBudPrepMaint.vaSpread1.PrintMarginTop / 1440, "0.00")
    pagemargin(3).Text = Format(frmBudPrepMaint.vaSpread1.PrintMarginBottom / 1440, "0.00")
    
    'Get page orientation
    If frmBudPrepMaint.vaSpread1.PrintOrientation = PrintOrientationLandscape Then
        porientation(1) = True
    Else
        porientation(0) = True
    End If
      
    'Populate Zooming combobox
    Combo1.AddItem "200%"
    Combo1.AddItem "150%"
    Combo1.AddItem "100%"
    Combo1.AddItem "75%"
    Combo1.AddItem "50%"
    Combo1.AddItem "25%"
    Combo1.AddItem "10%"
    Combo1.AddItem "Page Width"
    Combo1.AddItem "Page Height"
    Combo1.AddItem "Whole Page"
    Combo1.AddItem "Two Pages"
    Combo1.AddItem "Three Pages"
    Combo1.AddItem "Four Pages"
    Combo1.AddItem "Six Pages"
    
    'Get the zoom display
    Combo1.ListIndex = zoomindex
    
End Sub

