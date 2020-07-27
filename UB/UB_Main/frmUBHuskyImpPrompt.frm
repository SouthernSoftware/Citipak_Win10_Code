VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Begin VB.Form frmUBHuskyImpPrompt 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Husky Reading Import"
   ClientHeight    =   3684
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   6060
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   10.2
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
   ScaleHeight     =   3684
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin fpBtnAtlLibCtl.fpBtn cmdOK 
      Height          =   468
      Left            =   1380
      TabIndex        =   4
      Top             =   2760
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
      ButtonDesigner  =   "frmUBHuskyImpPrompt.frx":0000
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCancel 
      Height          =   468
      Left            =   3300
      TabIndex        =   5
      Top             =   2760
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
      ButtonDesigner  =   "frmUBHuskyImpPrompt.frx":01D6
   End
   Begin VB.Label lblHHImpTxt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   372
      Index           =   3
      Left            =   432
      TabIndex        =   3
      Top             =   2040
      Width           =   4932
   End
   Begin VB.Label lblHHImpTxt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CONTINUE. CLICK 'CANCEL' TO ABORT."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   372
      Index           =   2
      Left            =   504
      TabIndex        =   2
      Top             =   1488
      Width           =   4932
   End
   Begin VB.Label lblHHImpTxt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HAS FINISHED, THEN CLICK 'YES' TO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   372
      Index           =   1
      Left            =   504
      TabIndex        =   1
      Top             =   984
      Width           =   4932
   End
   Begin VB.Label lblHHImpTxt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WAIT UNTIL THE HUSKY FILE TRANSFER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   372
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   432
      Width           =   4932
   End
End
Attribute VB_Name = "frmUBHuskyImpPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

