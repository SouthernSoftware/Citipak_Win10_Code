VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "IMP32X30.OCX"
Begin VB.Form frmFABadTagNum 
   BackColor       =   &H000000C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3744
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6816
   LinkTopic       =   "Form1"
   ScaleHeight     =   3744
   ScaleWidth      =   6816
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   3744
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6816
      _Version        =   196609
      _ExtentX        =   12023
      _ExtentY        =   6604
      _StockProps     =   70
      BackColor       =   192
      Caption         =   ""
      FrameColor      =   192
      FrameThreeDHighlightColor=   8454143
      FrameThreeDShadowColor=   8454143
      FrameThreeDWidth=   4
      FrameWidth      =   8
      Picture         =   "frmFABadTagNum.frx":0000
      Begin VB.Timer Timer1 
         Interval        =   355
         Left            =   6336
         Top             =   96
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   540
         Left            =   378
         TabIndex        =   1
         Top             =   2304
         Width           =   2748
         _Version        =   131072
         _ExtentX        =   4847
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
         ButtonDesigner  =   "frmFABadTagNum.frx":001C
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdReturn 
         Height          =   540
         Left            =   3690
         TabIndex        =   2
         Top             =   2304
         Width           =   2748
         _Version        =   131072
         _ExtentX        =   4847
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
         ButtonDesigner  =   "frmFABadTagNum.frx":0207
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdSave 
         Height          =   540
         Left            =   2034
         TabIndex        =   3
         Top             =   2976
         Width           =   2748
         _Version        =   131072
         _ExtentX        =   4847
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
         ButtonDesigner  =   "frmFABadTagNum.frx":03EE
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ERROR!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   2418
         TabIndex        =   5
         Top             =   192
         Width           =   1980
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmFABadTagNum.frx":05D1
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1596
         Left            =   570
         TabIndex        =   4
         Top             =   624
         Width           =   5676
      End
   End
End
Attribute VB_Name = "frmFABadTagNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum BadFATagNumOption
  badfatInvalidOption = 0
  badfatExit
  badfatReturn
  badfatSave
End Enum

Private m_badfatOption As BadFATagNumOption

'// Create a property to get the Selection value.
'   NOTE: A Read-Only property has a Property Get but
'         no Property Let or Property Set
Property Get Selection() As BadFATagNumOption
  Selection = m_badfatOption
End Property

Private Sub cmdExit_Click()
'  On Error Resume Next
  m_badfatOption = badfatExit
  Unload frmFABadTagNum
  MainLog ("Exit option activated on frmFABadTagNum.")

End Sub

Private Sub cmdReturn_Click()
'  On Error Resume Next
  m_badfatOption = badfatReturn
  Unload frmFABadTagNum
  MainLog ("Return option activated on frmFABadTagNum.")

End Sub

Private Sub cmdSave_Click()
'  On Error Resume Next
  m_badfatOption = badfatSave
  Unload frmFABadTagNum
  MainLog ("Save option activated on frmFABadTagNum.")
  

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyF11:
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      Call cmdReturn_Click
      KeyCode = 0
    Case vbKeyF9:
      Call cmdSave_Click
      KeyCode = 0
    Case Else:
  End Select

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




