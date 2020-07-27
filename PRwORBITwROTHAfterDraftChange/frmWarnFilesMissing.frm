VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmWarnFilesMissing 
   BackColor       =   &H000000C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10530
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   5160
      Left            =   48
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
      FrameThreeDWidth=   4
      FrameWidth      =   8
      Picture         =   "frmWarnFilesMissing.frx":0000
      Begin LpLib.fpList fplistFileNames 
         Height          =   1485
         Left            =   840
         TabIndex        =   3
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
         ColDesigner     =   "frmWarnFilesMissing.frx":001C
      End
      Begin VB.Timer Timer1 
         Interval        =   355
         Left            =   10080
         Top             =   96
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdReturn 
         Height          =   540
         Left            =   4026
         TabIndex        =   4
         Top             =   4224
         Width           =   2484
         _Version        =   131072
         _ExtentX        =   4382
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
         ButtonDesigner  =   "frmWarnFilesMissing.frx":03E9
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
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   4278
         TabIndex        =   2
         Top             =   480
         Width           =   1980
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWarnFilesMissing.frx":05FB
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
         Height          =   1260
         Left            =   2430
         TabIndex        =   1
         Top             =   960
         Width           =   5676
      End
   End
End
Attribute VB_Name = "frmWarnFilesMissing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdReturn_Click()
  Unload frmWarnFilesMissing
  DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF10 Then
    KeyCode = 0
    Unload frmWarnFilesMissing
  End If

End Sub

Private Sub Form_Load()
  Dim x As Integer
  For x = 1 To 21
    If QPTrim$(UCase(OutFileNames(x))) = "" Then Exit For
    Select Case QPTrim$(UCase(OutFileNames(x)))
      Case "PRDATA\PRSTADEF.DAT":
        fplistFileNames.AddItem "A copy of PRSTADEF.DAT needs to be placed in the PRDATA folder."
        fplistFileNames.AddItem "  "
      Case "PRDATA\PRERNCOD.DAT":
        fplistFileNames.AddItem "Go to Control Maintenance Menu, then save at least one entry in Earnings Codes."
        fplistFileNames.AddItem "  "
      Case "PRDATA\PRLEAVE.DAT":
        fplistFileNames.AddItem "Go to Control Maintenance Menu, open Leave Benefit Table then save."
        fplistFileNames.AddItem "  "
      Case "PRDATA\PRSTATAX.DAT":
        fplistFileNames.AddItem "Go to Control Maintenance Menu, open State Tax Table then save."
        fplistFileNames.AddItem "  "
      Case "PRDATA\PRFEDTAX.DAT":
        fplistFileNames.AddItem "Go to Control Maintenance Menu, open Federal Tax Table then save data."
        fplistFileNames.AddItem "  "
      Case "PRDATA\PREICTBL.DAT":
        fplistFileNames.AddItem "Go to Control Maintenance Menu, open EIC Table then save data."
        fplistFileNames.AddItem "  "
      Case "PRDATA\PRUNIT.DAT":
        fplistFileNames.AddItem "Go to Control Maintenance Menu, open Employer File then save data."
        fplistFileNames.AddItem "  "
      Case "PRDATA\PREMP1.DAT":
        fplistFileNames.AddItem "Go to Employee Maintenance Menu and save data for at least one employee."
        fplistFileNames.AddItem "  "
      Case "PRDATA\PREMP2.DAT":
        fplistFileNames.AddItem "Go to Employee Maintenance Menu and save data for at least one new employee."
        fplistFileNames.AddItem "  "
      Case "PRDATA\PREMP3.DAT":
        fplistFileNames.AddItem "Go to Employee Maintenance Menu and save data for at least one employee."
        fplistFileNames.AddItem "  "
      Case "PRDATA\PRSYS.DAT":
        fplistFileNames.AddItem "Go to Control Maintenance Menu, open System File then save data."
        fplistFileNames.AddItem "  "
      Case "PRDATA\PRCHECKS.DAT":
        fplistFileNames.AddItem "Go to Payroll Processing Menu: Print Payroll Checks: Print Check Register and print data."
        fplistFileNames.AddItem "  "
      Case "PRDATA\PRPPDEF.DAT":
        fplistFileNames.AddItem "Go to Payroll Processing Menu, open Set Pay Period Defaults then save data."
        fplistFileNames.AddItem "  "
      Case "PRDATA\PRRETIRE.DAT":
        fplistFileNames.AddItem "Go to Control Maintenance Menu, open Retirement File then save data."
        fplistFileNames.AddItem "  "
      Case "PRDATA\PRDEDCOD.DAT":
        fplistFileNames.AddItem "Go to Control Maintenance Menu, open Deduction Code then save at least one deduction."
        fplistFileNames.AddItem "  "
      Case "PRDATA\PRDRAFTI.DAT":
        fplistFileNames.AddItem "Go to Control Maintenance Menu, open ACH Bank Draft Information then save data."
        fplistFileNames.AddItem "  "
      Case "PRDATA\PRPRNDF.DAT":
        fplistFileNames.AddItem "The file PRPRNDF.DAT needs to be copied into the PRDATA folder."
        fplistFileNames.AddItem "  "
      Case "PRDATA\crchek.dat":
        fplistFileNames.AddItem "Go to Payroll Processing Menu, open Post Transactions then post."
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

