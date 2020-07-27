VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmWarnYrNotSubsequent 
   BackColor       =   &H000000C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6780
   Icon            =   "frmWarnYrNotSubsequent.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   4224
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6816
      _Version        =   196609
      _ExtentX        =   12023
      _ExtentY        =   7451
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.24
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
      Picture         =   "frmWarnYrNotSubsequent.frx":08CA
      Begin VB.Timer Timer1 
         Interval        =   355
         Left            =   6336
         Top             =   96
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   540
         Left            =   960
         TabIndex        =   1
         ToolTipText     =   "Click this button to stop the depreciation process and close this message."
         Top             =   3312
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
         ButtonDesigner  =   "frmWarnYrNotSubsequent.frx":08E6
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   540
         Left            =   3840
         TabIndex        =   5
         ToolTipText     =   "Click this button to override this warning and continue with the depreciation process."
         Top             =   3312
         Width           =   2052
         _Version        =   131072
         _ExtentX        =   3619
         _ExtentY        =   952
         Enabled         =   -1  'True
         MouseIcon       =   "frmWarnYrNotSubsequent.frx":0AFA
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
         ButtonDesigner  =   "frmWarnYrNotSubsequent.frx":13D4
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWarnYrNotSubsequent.frx":26DC
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
         Height          =   1785
         Left            =   570
         TabIndex        =   4
         Top             =   840
         Width           =   5670
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Press F10 to Exit."
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
         Height          =   396
         Left            =   480
         TabIndex        =   3
         Top             =   2688
         Width           =   5868
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "WARNING!"
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
         Left            =   2400
         TabIndex        =   2
         Top             =   336
         Width           =   1980
      End
   End
End
Attribute VB_Name = "frmWarnYrNotSubsequent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum NotSubsequentOption
  nsInvalidOption = 0
  nsExit
  nsProcess
End Enum

Private m_nsOption As NotSubsequentOption

'// Create a property to get the Selection value.
'   NOTE: A Read-Only property has a Property Get but
'         no Property Let or Property Set
Property Get Selection() As NotSubsequentOption
  Selection = m_nsOption
End Property

Private Sub cmdExit_Click()
'  On Error Resume Next
  m_nsOption = nsExit
  Unload frmWarnYrNotSubsequent
  MainLog ("Exit option activated on frmWarnYrNotSubsequent.")

End Sub

Private Sub cmdProcess_Click()
'  On Error Resume Next
  m_nsOption = nsProcess
  Unload frmWarnYrNotSubsequent
  MainLog ("Process option activated on frmWarnYrNotSubsequent.")

End Sub

Private Sub cmdOk_Click()
  Unload frmWarnYrNotSubsequent
  MainLog ("Warning for year to depreciate is not subsequent to the last depreciated year issued.")
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyF10:
    Call cmdExit_Click
    KeyCode = 0
  Case vbKeyF3:
    Call cmdProcess_Click
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


