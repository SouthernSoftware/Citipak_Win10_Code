VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmWarningACH 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   4596
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6972
      _Version        =   196609
      _ExtentX        =   12298
      _ExtentY        =   8107
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
      FrameColor      =   192
      FrameThreeDHighlightColor=   8454143
      FrameThreeDShadowColor=   8454143
      FrameThreeDWidth=   4
      FrameWidth      =   8
      Picture         =   "frmWarningACH.frx":0000
      Begin VB.Timer Timer1 
         Interval        =   355
         Left            =   6528
         Top             =   96
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdACH 
         Height          =   540
         Left            =   960
         TabIndex        =   7
         Top             =   3024
         Width           =   5028
         _Version        =   131072
         _ExtentX        =   8869
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
         ButtonDesigner  =   "frmWarningACH.frx":001C
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdEmpInfo 
         Height          =   540
         Left            =   1824
         TabIndex        =   8
         Top             =   3744
         Width           =   3300
         _Version        =   131072
         _ExtentX        =   5821
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
         ButtonDesigner  =   "frmWarningACH.frx":024C
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "option and enter ""N"" in the ""Allow Bank Drafts"" field."
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
         Height          =   588
         Left            =   276
         TabIndex        =   6
         Top             =   2256
         Width           =   6396
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "complete this data or go to the EMPLOYER FILE"
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
         Height          =   396
         Left            =   276
         TabIndex        =   5
         Top             =   1884
         Width           =   6396
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DRAFT INFORMATION"" control option. Please"
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
         Height          =   396
         Left            =   276
         TabIndex        =   4
         Top             =   1500
         Width           =   6396
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "You have elected to allow ACH Bank Drafts"
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
         Height          =   348
         Left            =   804
         TabIndex        =   3
         Top             =   732
         Width           =   5340
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "but nothing has been saved in the ACH BANK"
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
         Height          =   396
         Left            =   276
         TabIndex        =   2
         Top             =   1116
         Width           =   6396
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
         Left            =   2484
         TabIndex        =   1
         Top             =   300
         Width           =   1980
      End
   End
End
Attribute VB_Name = "frmWarningACH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdACH_Click()
  Unload frmWarningACH
  DoEvents
  frmACHDraftInfo.Show
  DoEvents
  Unload frmPayrollMainMenu
  MainLog ("ACH warning issued...go to ACH screen option chosen.")
End Sub

Private Sub cmdEmpInfo_Click()
  Unload frmWarningACH
  DoEvents
  frmEmployerInfoFile.Show
  frmEmployerInfoFile.fpcomboBankDraftYN.SetFocus
  Unload frmPayrollMainMenu
  MainLog ("ACH warning issued...go to Employer Info screen option chosen.")
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyF10:
      Call cmdACH_Click
      KeyCode = 0
    Case vbKeyF11:
      Call cmdEmpInfo_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

