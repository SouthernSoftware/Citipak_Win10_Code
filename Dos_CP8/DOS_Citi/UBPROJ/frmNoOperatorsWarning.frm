VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Begin VB.Form frmNoOperatorsWarning 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WARNING !!!"
   ClientHeight    =   4020
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   7428
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   10.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7428
   StartUpPosition =   2  'CenterScreen
   Begin fpBtnAtlLibCtl.fpBtn cmdOK 
      Height          =   468
      Left            =   1992
      TabIndex        =   0
      Top             =   3216
      Width           =   1284
      _Version        =   131072
      _ExtentX        =   2265
      _ExtentY        =   825
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
      ButtonDesigner  =   "frmNoOperatorsWarning.frx":0000
   End
   Begin VB.Timer frmNoOperTimer 
      Interval        =   355
      Left            =   168
      Top             =   192
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCancel 
      Height          =   468
      Left            =   4032
      TabIndex        =   5
      Top             =   3216
      Width           =   1284
      _Version        =   131072
      _ExtentX        =   2265
      _ExtentY        =   825
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
      ButtonDesigner  =   "frmNoOperatorsWarning.frx":01D5
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HAS COMPLETED!!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   324
      Index           =   2
      Left            =   168
      TabIndex        =   4
      Top             =   1560
      Width           =   7080
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CONTINUE WITH REINDEX/RELINK??"
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
      Height          =   348
      Index           =   5
      Left            =   180
      TabIndex        =   3
      Top             =   2352
      Width           =   7080
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ALL UTILITY BILLING OPERATIONS MUST EXIT THE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   300
      Index           =   0
      Left            =   168
      TabIndex        =   2
      Top             =   792
      Width           =   7080
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UTILITY BILLING PROGRAM UNTIL THIS PROCEDURE "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   348
      Index           =   1
      Left            =   168
      TabIndex        =   1
      Top             =   1176
      Width           =   7080
   End
End
Attribute VB_Name = "frmNoOperatorsWarning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  DoItFlag = False
  Call UnLoadWarning
End Sub

Private Sub cmdOK_Click()
  DoItFlag = True
  Call UnLoadWarning
End Sub


Private Sub Form_Paint()
  If Me.Visible Then
    Me.cmdOK.SetFocus
  End If
End Sub

Private Sub frmNoOperTimer_Timer()
  Dim BkColor As Integer
  Static tog As Boolean
  tog = Not tog
  If tog Then
    Me.BackColor = 230
  Else
    Me.BackColor = 192
  End If
End Sub

Private Sub Form_Load()
  Dim RetVal As Long, winhand As Long
  MakeWindowTopMost Me.hwnd, True
  DoEvents
End Sub

Private Sub UnLoadWarning()
  MakeWindowTopMost Me.hwnd, False
  DoEvents
  Unload frmNoOperatorsWarning
End Sub

