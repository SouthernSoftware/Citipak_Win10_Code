VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmFAWarnOverWriteCode 
   BackColor       =   &H000000C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8190
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   355
      Left            =   7536
      Top             =   4464
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   540
      Left            =   5328
      TabIndex        =   0
      ToolTipText     =   "Click this button to save the data entered. Data existing before changes will be overwritten."
      Top             =   4848
      Width           =   1860
      _Version        =   131072
      _ExtentX        =   3281
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
      ButtonDesigner  =   "frmFAWarnOverWriteCode.frx":0000
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdReview 
      Height          =   540
      Left            =   3072
      TabIndex        =   1
      ToolTipText     =   "Click this button to return to the screen to examine the changes made."
      Top             =   4848
      Width           =   1476
      _Version        =   131072
      _ExtentX        =   2603
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
      ButtonDesigner  =   "frmFAWarnOverWriteCode.frx":01E0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   540
      Left            =   1050
      TabIndex        =   2
      ToolTipText     =   "Click this button to return to the previous screen without saving any changes made."
      Top             =   4845
      Width           =   1485
      _Version        =   131072
      _ExtentX        =   2619
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
      ButtonDesigner  =   "frmFAWarnOverWriteCode.frx":03BD
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdGo2Add 
      Height          =   540
      Left            =   2352
      TabIndex        =   16
      ToolTipText     =   "Click this button to add this new data to the current list retaining the data that existed before the changes were made."
      Top             =   5712
      Width           =   3492
      _Version        =   131072
      _ExtentX        =   6159
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
      ButtonDesigner  =   "frmFAWarnOverWriteCode.frx":0599
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "To ADD a new number press F12"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   516
      Left            =   936
      TabIndex        =   17
      Top             =   2112
      Width           =   6312
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "to Go To Add New Asset Code."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   396
      Left            =   3312
      TabIndex        =   15
      Top             =   4128
      Width           =   3720
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Press"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   1776
      TabIndex        =   14
      Top             =   4128
      Width           =   720
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(F12)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   372
      Left            =   2520
      TabIndex        =   13
      Top             =   4128
      Width           =   696
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(ESC)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   372
      Left            =   2508
      TabIndex        =   12
      Top             =   3732
      Width           =   696
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Press"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   1764
      TabIndex        =   11
      Top             =   3732
      Width           =   720
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "to REVIEW Changes."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   396
      Left            =   3300
      TabIndex        =   10
      Top             =   3732
      Width           =   3336
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(X)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   372
      Left            =   2532
      TabIndex        =   9
      Top             =   3324
      Width           =   528
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "to ABANDON Changes."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   3144
      TabIndex        =   8
      Top             =   3312
      Width           =   3312
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "to SAVE Changes."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   396
      Left            =   3348
      TabIndex        =   7
      Top             =   2916
      Width           =   2184
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(F10)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   372
      Left            =   2508
      TabIndex        =   6
      Top             =   2916
      Width           =   696
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmFAWarnOverWriteCode.frx":0782
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1812
      Left            =   432
      TabIndex        =   5
      Top             =   252
      Width           =   7368
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Press"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   1740
      TabIndex        =   4
      Top             =   2916
      Width           =   696
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Press"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   1764
      TabIndex        =   3
      Top             =   3324
      Width           =   720
   End
End
Attribute VB_Name = "frmFAWarnOverWriteCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum WarnOption
  wInvalidOption = 0
  wExit
  wReturn
  wSave
  wGo2Add
End Enum

Private m_wOption As WarnOption

'// Create a property to get the Selection value.
'   NOTE: A Read-Only property has a Property Get but
'         no Property Let or Property Set
Property Get Selection() As WarnOption
  Selection = m_wOption
End Property

Private Sub cmdExit_Click()
'  On Error Resume Next
  m_wOption = wExit
  Unload frmFAWarnOverWriteCode
  MainLog ("Exit option activated on frmFAWarnOverWriteCode.")

End Sub

Private Sub cmdReturn_Click()

End Sub

Private Sub cmdReview_Click()
'  On Error Resume Next
  m_wOption = wReturn
  Unload frmFAWarnOverWriteCode
  MainLog ("Return option activated on frmFAWarnOverWriteCode.")

End Sub

Private Sub cmdSave_Click()
'  On Error Resume Next
  m_wOption = wSave
  Unload frmFAWarnOverWriteCode
  MainLog ("Save option activated on frmFAWarnOverWriteCode.")
  

End Sub

Private Sub cmdGo2Add_Click()
'  On Error Resume Next
  m_wOption = wGo2Add
  Unload frmFAWarnOverWriteCode
  MainLog ("Go2Add option activated on frmFAWarnOverWriteCode.")
  

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyX:
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyEscape:
      Call cmdReturn_Click
      KeyCode = 0
    Case vbKeyF10:
      Call cmdSave_Click
      KeyCode = 0
    Case vbKeyF12:
      Call cmdGo2Add_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Timer1_Timer()
  Static tog As Boolean
  tog = Not tog
  If tog Then
    Me.BackColor = 210
  Else
    Me.BackColor = 192
  End If
End Sub



