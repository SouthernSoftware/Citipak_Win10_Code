VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmWarnInUse 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   8790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8925
   Icon            =   "frmWarnInUse.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   355
      Left            =   4080
      Top             =   96
   End
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6396
      Left            =   1200
      TabIndex        =   0
      Top             =   864
      Width           =   6540
      _Version        =   196609
      _ExtentX        =   11536
      _ExtentY        =   11282
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.26
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   192
      Caption         =   ""
      Picture         =   "frmWarnInUse.frx":08CA
      Begin LpLib.fpList fpList1 
         Height          =   1200
         Left            =   1245
         TabIndex        =   3
         Top             =   4755
         Width           =   4005
         _Version        =   196608
         _ExtentX        =   7064
         _ExtentY        =   2117
         TextAlias       =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
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
         ColDesigner     =   "frmWarnInUse.frx":08E6
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWarnInUse.frx":0CB3
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   3564
         Left            =   720
         TabIndex        =   2
         Top             =   1056
         Width           =   5052
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Warning!   Warning!   Warning!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   16.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   492
         Left            =   528
         TabIndex        =   1
         Top             =   384
         Width           =   5436
         WordWrap        =   -1  'True
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdContinue 
      Height          =   675
      Left            =   1770
      TabIndex        =   4
      Top             =   7830
      Width           =   1935
      _Version        =   131072
      _ExtentX        =   3413
      _ExtentY        =   1191
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
      ButtonDesigner  =   "frmWarnInUse.frx":0E0A
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   690
      Left            =   5160
      TabIndex        =   5
      Top             =   7824
      Width           =   1935
      _Version        =   131072
      _ExtentX        =   3413
      _ExtentY        =   1217
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
      ButtonDesigner  =   "frmWarnInUse.frx":1022
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Height          =   6876
      Left            =   960
      Top             =   648
      Width           =   7020
   End
End
Attribute VB_Name = "frmWarnInUse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum PRInUse
'begin by creating a new property = SaveChangeOptions1
  priuInvalidOption = 0 'enum member
  priuContinue 'enum member
  priuExit 'enum member
End Enum

Private m_priuOption As PRInUse
Property Get Selection() As PRInUse
  'create a new property called Selection
  Selection = m_priuOption
End Property

Private Sub cmdExit_Click()
  'if review is chosen then the selection is scoReviewChanges
  m_priuOption = priuExit
  Unload frmWarnInUse
  MainLog ("Payroll in use warning issued...exit option selected.")
End Sub

Private Sub cmdContinue_Click()
  'if save is chosen then the selection is scoSave
  m_priuOption = priuContinue
  Unload frmWarnInUse
  MainLog ("Payroll in use warning issued...continue option selected.")
End Sub

Private Sub Timer1_Timer()
'the timer is set to 355 which means that everytime
'355 is reached this sub starts over...since tog
'is static it is remembered even though the sub closes
  Static tog As Boolean
  tog = Not tog
  If tog Then
    vaImprint1.BackColor = 210
  Else
    vaImprint1.BackColor = 192
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      SendKeys "%X"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%C"
      Call cmdContinue_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Dim CitiPassFile As Integer
  Dim CitiPass As CitiPassType
  Dim NumPassRecs As Integer
  Dim x As Integer
  
  OpenCitiPassFile CitiPassFile, NumPassRecs 'reassign all globals
  For x = 1 To NumPassRecs
    Get CitiPassFile, x, CitiPass
    If CitiPass.FlagMod = 4 And CitiPass.Module(4).FullAccess = True Then
      fpList1.InsertRow = CitiPass.UserName & "on   " & CitiPass.CompName
    End If
  Next x
  Close CitiPassFile

End Sub
