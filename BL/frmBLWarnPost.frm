VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmBLWarnPost 
   BackColor       =   &H000000C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7155
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   355
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox fptxtChoice 
      Height          =   288
      Left            =   96
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   3984
      Width           =   492
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   540
      Left            =   1152
      TabIndex        =   1
      Top             =   3648
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
      ButtonDesigner  =   "frmBLWarnPost.frx":0000
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPost 
      Height          =   540
      Left            =   3936
      TabIndex        =   2
      Top             =   3648
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
      ButtonDesigner  =   "frmBLWarnPost.frx":0216
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ONCE POSTED,  BUSINESS LICENSES ASSOCIATED WITH THIS POST CANNOT BE REPRINTED."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   876
      Left            =   624
      TabIndex        =   5
      Top             =   2448
      Width           =   5820
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmBLWarnPost.frx":042A
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1740
      Left            =   672
      TabIndex        =   4
      Top             =   768
      Width           =   5820
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WARNING!!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   396
      Left            =   1608
      TabIndex        =   3
      Top             =   240
      Width           =   3852
   End
End
Attribute VB_Name = "frmBLWarnPost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
  frmBLWarnPost.Hide
  fptxtChoice = "exit"
End Sub

Private Sub cmdPost_Click()
  frmBLWarnPost.Hide
  fptxtChoice = "post"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdPost_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  fptxtChoice.Visible = False
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


