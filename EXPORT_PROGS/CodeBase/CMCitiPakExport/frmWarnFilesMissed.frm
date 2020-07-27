VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmWarnFilesMissed 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10575
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   5160
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10476
      _Version        =   196609
      _ExtentX        =   18478
      _ExtentY        =   9102
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   192
      Caption         =   ""
      FrameColor      =   192
      FrameThreeDHighlightColor=   8454143
      FrameThreeDShadowColor=   8454143
      FrameThreeDStyle=   2
      FrameThreeDWidth=   4
      FrameWidth      =   8
      Picture         =   "frmWarnFilesMissed.frx":0000
      Begin LpLib.fpList fplistFileNames 
         Height          =   1485
         Left            =   810
         TabIndex        =   1
         Top             =   2445
         Width           =   8850
         _Version        =   196608
         _ExtentX        =   15610
         _ExtentY        =   2619
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
         Columns         =   0
         Sorted          =   0
         LineWidth       =   1
         SelDrawFocusRect=   -1  'True
         ColumnSeparatorChar=   9
         ColumnSearch    =   -1
         ColumnWidthScale=   2
         RowHeight       =   -1
         MultiSelect     =   0
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
         ReadOnly        =   -1  'True
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
         ColDesigner     =   "frmWarnFilesMissed.frx":001C
      End
      Begin VB.Timer Timer1 
         Interval        =   355
         Left            =   8448
         Top             =   384
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdReturn 
         Height          =   540
         Left            =   4272
         TabIndex        =   4
         Top             =   4176
         Width           =   2052
         _Version        =   131072
         _ExtentX        =   3619
         _ExtentY        =   952
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
         ButtonDesigner  =   "frmWarnFilesMissed.frx":02E0
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWarnFilesMissed.frx":04F2
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1308
         Left            =   2496
         TabIndex        =   3
         Top             =   864
         Width           =   5676
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ERROR!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   300
         Left            =   4386
         TabIndex        =   2
         Top             =   480
         Width           =   1980
      End
   End
End
Attribute VB_Name = "frmWarnFilesMissed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdReturn_Click()
  Unload frmWarnFilesMissed
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyF10:
      Call cmdReturn_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Dim x As Integer
  For x = 1 To 21
    If QPTrim$(UCase(OutFileNames(x))) = "" Then Exit For
    Select Case QPTrim$(UCase(OutFileNames(x)))
      Case UCase("PRData\P9013-39MSK.txt"):
        fplistFileNames.AddItem "A copy of P9013-39MSK.txt needs to be placed in the PRDATA folder."
        fplistFileNames.AddItem "  "
      Case UCase("PRData\P9013-42MSK.txt"):
        fplistFileNames.AddItem "A copy of P9013-42MSK.txt needs to be placed in the PRDATA folder."
        fplistFileNames.AddItem "  "
      Case UCase("PRData\P9028MSK.txt"):
        fplistFileNames.AddItem "A copy of P9028MSK.txt needs to be placed in the PRDATA folder."
        fplistFileNames.AddItem "  "
      Case UCase("PRData\P9007MSK.txt"):
        fplistFileNames.AddItem "A copy of P9007MSK.txt needs to be placed in the PRDATA folder."
        fplistFileNames.AddItem "  "
      Case UCase("PRData\Laser1MSK.txt"):
        fplistFileNames.AddItem "A copy of Laser1MSK.txt needs to be placed in the PRDATA folder."
        fplistFileNames.AddItem "  "
      Case UCase("PRData\Laser2Msk.txt"):
        fplistFileNames.AddItem "A copy of Laser2Msk.txt needs to be placed in the PRDATA folder."
        fplistFileNames.AddItem "  "
        
      Case UCase("PRData\PRUNIT.DAT"): '7/20 from here to end select added
        fplistFileNames.AddItem "Go to the Control Menu and save data in the Employer Information folder."
        fplistFileNames.AddItem "  "
      Case UCase("PRData\PREMPN.IDX"):
        fplistFileNames.AddItem "A copy of PREMPN.IDX needs to be placed in the PRData folder."
        fplistFileNames.AddItem "  "
      Case UCase("PRData\PRCHECKS.DAT"):
        fplistFileNames.AddItem "A copy of PRCHECKS.DAT needs to be placed in the PRData folder."
        fplistFileNames.AddItem "  "
      Case UCase("PRData\PRSYS.DAT"):
        fplistFileNames.AddItem "Go to the Control Menu and save data in the System Interface folder."
        fplistFileNames.AddItem "  "
      Case UCase("PRData\PRDEDCOD.DAT"):
        fplistFileNames.AddItem "Go to the Control Menu and save data in the Deductions folder."
        fplistFileNames.AddItem "  "
      Case UCase("PRData\PRERNCOD.DAT"):
        fplistFileNames.AddItem "Go to the Control Menu and save data in the Earnings Code folder."
        fplistFileNames.AddItem "  "
      Case UCase("PRData\PRTRANST.DAT"):
        fplistFileNames.AddItem "A copy of PRTRANST.DAT needs to be placed in the PRData folder."
        fplistFileNames.AddItem "  "
      Case UCase("PRData\PREMP2.DAT"):
        fplistFileNames.AddItem "No employee (PREMP2.DAT) data file can be found."
        fplistFileNames.AddItem "  "
      Case UCase("PRData\PREMP3.DAT"):
        fplistFileNames.AddItem "No employee (PREMP3.DAT) data file can be found."
        fplistFileNames.AddItem "  "
      Case UCase("PRData\PRPRNSET.DAT"):
        fplistFileNames.AddItem "Go to the Control Menu and save data in the Printer Setup folder."
        fplistFileNames.AddItem "  "
      Case UCase("PRData\PRPRNDF.DAT"):
        fplistFileNames.AddItem "A copy of PRPRNDF.DAT needs to be added to the PRDATA folder."
        fplistFileNames.AddItem "  "
      Case "":
        GoTo EmptyString
      Case Else:
        fplistFileNames.AddItem "Error: A file is missing. Please call Southern Software for assistance."
        fplistFileNames.AddItem "  "
    End Select
EmptyString:
  Next x
  MainLog ("Files Missing warning issued.")
End Sub

Private Sub Timer1_Timer()
  Static tog As Boolean
  tog = Not tog
  If tog Then
    vaImprint1.BackColor = 210
  Else
    vaImprint1.BackColor = 192
  End If
End Sub


