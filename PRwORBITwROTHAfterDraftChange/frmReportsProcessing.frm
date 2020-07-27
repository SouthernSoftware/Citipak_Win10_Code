VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmReportsProcessing 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reports Processing"
   ClientHeight    =   8880
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11655
   Icon            =   "frmReportsProcessing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn GrossWageReportCmmd 
      Height          =   375
      Left            =   3060
      TabIndex        =   5
      Top             =   4608
      Width           =   2760
      _Version        =   131072
      _ExtentX        =   4868
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmReportsProcessing.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn TerminatedEmployeesCmmd 
      Height          =   375
      Left            =   3060
      TabIndex        =   3
      Top             =   3468
      Width           =   2760
      _Version        =   131072
      _ExtentX        =   4868
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmReportsProcessing.frx":0AA9
   End
   Begin fpBtnAtlLibCtl.fpBtn PrintEmployeeDataFileCmmd 
      Height          =   375
      Left            =   3060
      TabIndex        =   1
      Top             =   2328
      Width           =   2760
      _Version        =   131072
      _ExtentX        =   4868
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmReportsProcessing.frx":0C95
   End
   Begin fpBtnAtlLibCtl.fpBtn ActiveEmployeeListCmmd 
      Height          =   375
      Left            =   3072
      TabIndex        =   2
      Top             =   2904
      Width           =   2760
      _Version        =   131072
      _ExtentX        =   4868
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmReportsProcessing.frx":0E81
   End
   Begin fpBtnAtlLibCtl.fpBtn EmployeeEarningsHistCmmd 
      Height          =   375
      Left            =   3060
      TabIndex        =   4
      Top             =   4044
      Width           =   2760
      _Version        =   131072
      _ExtentX        =   4868
      _ExtentY        =   661
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
      DrawFocusRect   =   3
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
      ButtonDesigner  =   "frmReportsProcessing.frx":1069
   End
   Begin fpBtnAtlLibCtl.fpBtn PayrollDeductionsTakenCmmd 
      Height          =   375
      Left            =   3060
      TabIndex        =   6
      Top             =   5175
      Width           =   2760
      _Version        =   131072
      _ExtentX        =   4868
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmReportsProcessing.frx":124D
   End
   Begin fpBtnAtlLibCtl.fpBtn ESCReportCmmd 
      Height          =   375
      Left            =   3060
      TabIndex        =   7
      Top             =   5748
      Width           =   2760
      _Version        =   131072
      _ExtentX        =   4868
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmReportsProcessing.frx":1439
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdReprint 
      Height          =   375
      Left            =   3060
      TabIndex        =   8
      Top             =   6336
      Width           =   2760
      _Version        =   131072
      _ExtentX        =   4868
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmReportsProcessing.frx":1611
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEmergency 
      Height          =   375
      Left            =   3060
      TabIndex        =   9
      Top             =   6888
      Width           =   2760
      _Version        =   131072
      _ExtentX        =   4868
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmReportsProcessing.frx":17F4
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPayRate 
      Height          =   375
      Left            =   5916
      TabIndex        =   19
      Top             =   6885
      Width           =   2670
      _Version        =   131072
      _ExtentX        =   4710
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmReportsProcessing.frx":19DB
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSupRetRpts 
      Height          =   375
      Left            =   5916
      TabIndex        =   16
      Top             =   5184
      Width           =   2670
      _Version        =   131072
      _ExtentX        =   4710
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmReportsProcessing.frx":1BBF
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAnnWrkrsComp 
      Height          =   375
      Left            =   5916
      TabIndex        =   17
      Top             =   5748
      Width           =   2670
      _Version        =   131072
      _ExtentX        =   4710
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmReportsProcessing.frx":1DA7
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSEPPCon 
      Height          =   375
      Left            =   5916
      TabIndex        =   18
      Top             =   6324
      Width           =   2670
      _Version        =   131072
      _ExtentX        =   4710
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmReportsProcessing.frx":1F8F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   375
      Left            =   5916
      TabIndex        =   20
      ToolTipText     =   "Press to access a detailed employee report."
      Top             =   7440
      Width           =   2670
      _Version        =   131072
      _ExtentX        =   4710
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmReportsProcessing.frx":2175
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdRetireRpts 
      Height          =   375
      Left            =   5916
      TabIndex        =   15
      Top             =   4608
      Width           =   2670
      _Version        =   131072
      _ExtentX        =   4710
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmReportsProcessing.frx":235A
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdChecksbyNum 
      Height          =   375
      Left            =   5904
      TabIndex        =   14
      Top             =   4044
      Width           =   2670
      _Version        =   131072
      _ExtentX        =   4710
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmReportsProcessing.frx":2539
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdChecksIss 
      Height          =   375
      Left            =   5904
      TabIndex        =   13
      Top             =   3465
      Width           =   2670
      _Version        =   131072
      _ExtentX        =   4710
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmReportsProcessing.frx":271D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdYTDWageDis 
      Height          =   375
      Left            =   5916
      TabIndex        =   12
      Top             =   2904
      Width           =   2670
      _Version        =   131072
      _ExtentX        =   4710
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmReportsProcessing.frx":28FF
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdLeaveBen 
      Height          =   375
      Left            =   5916
      TabIndex        =   11
      Top             =   2328
      Width           =   2670
      _Version        =   131072
      _ExtentX        =   4710
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmReportsProcessing.frx":2AE9
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdTaxFringe 
      Height          =   375
      Left            =   3060
      TabIndex        =   10
      ToolTipText     =   "Press to access a detailed employee report."
      Top             =   7440
      Width           =   2760
      _Version        =   131072
      _ExtentX        =   4868
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmReportsProcessing.frx":2CCB
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   2101
      Top             =   2103
      Width           =   972
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   8593
      Top             =   2103
      Width           =   971
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "REPORTS MENU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2820
      TabIndex        =   0
      Top             =   1250
      Width           =   6012
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Height          =   1097
      Left            =   1500
      Top             =   897
      Width           =   8655
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   8710.757
      X2              =   9412.576
      Y1              =   7884
      Y2              =   7884
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   2205.432
      X2              =   2919.248
      Y1              =   7884
      Y2              =   7884
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   8710.757
      X2              =   8710.757
      Y1              =   2151.243
      Y2              =   7892.757
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   2220.428
      X2              =   2220.428
      Y1              =   2151.243
      Y2              =   7879.135
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5910
      Index           =   0
      Left            =   2220
      Top             =   2201
      Width           =   732
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5910
      Index           =   1
      Left            =   8712
      Top             =   2201
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8592
      Top             =   1971
      Width           =   972
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2100
      Top             =   1971
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1214
      Left            =   1500
      Top             =   770
      Width           =   8652
   End
End
Attribute VB_Name = "frmReportsProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Public Enum ReportOpt
  roInvalidOption = 0
  roOn
  roOff
End Enum
Private m_roOption As ReportOpt
Property Get Selection() As ReportOpt
  Selection = m_roOption
End Property

Private Sub ActiveEmployeeListCmmd_Click()
   m_roOption = roOn
   frmPrintAlphaNum.Show
   DoEvents
   Unload frmReportsProcessing
End Sub

Private Sub cmdAnnWrkrsComp_Click()
  InFileNames(1) = "PRDATA\PRUNIT.DAT"
  InFileNames(2) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 2) = False Then
    Close
    Exit Sub
  End If
  frmCompWageRpt.Show
  DoEvents
  Unload frmReportsProcessing
End Sub


Private Sub cmdChecksbyNum_Click()
  InFileNames(1) = "PRDATA\PRUNIT.DAT"
  InFileNames(2) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 2) = False Then
    Close
    Exit Sub
  End If
  frmChecksbyNumber.Show
  DoEvents
  Unload frmReportsProcessing
End Sub

Private Sub cmdChecksIss_Click()
  InFileNames(1) = "PRDATA\PRUNIT.DAT"
  InFileNames(2) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 2) = False Then
    Close
    Exit Sub
  End If
  frmEmpChksIssued.Show
  DoEvents
  Unload frmReportsProcessing
End Sub

Private Sub cmdEmergency_Click()
  InFileNames(1) = "PRDATA\PRUNIT.DAT"
  InFileNames(2) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 2) = False Then
    Close
    Exit Sub
  End If
  m_roOption = roOn
  frmEmergency.Show
  DoEvents
  Unload frmEmployeeMaintMenu

End Sub

Private Sub cmdLeaveBen_Click()
  InFileNames(1) = "PRDATA\PRUNIT.DAT"
  InFileNames(2) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 2) = False Then
    Close
    Exit Sub
  End If
  frmLeaveBenefit.Show
  DoEvents
  Unload frmReportsProcessing
End Sub

Private Sub cmdPayRate_Click()
  frmPayRateRpt.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdReprint_Click()
  frmReprint.Show
  DoEvents
  Unload frmReportsProcessing
End Sub

Private Sub cmdRetireRpts_Click()
  Dim UnitHandle As Integer
  Dim UnitFileRec As UnitFileRecType
  Dim State As String
   
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, UnitFileRec
  Close UnitHandle
  State = QPTrim$(UnitFileRec.UFSTATE)
  InFileNames(1) = "PRDATA\PRUNIT.DAT"
  InFileNames(2) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 2) = False Then
    Close
    Exit Sub
  End If
  If State <> "NC" Then 'new for ORBIT
    frmRetRpt.Show
  ElseIf State = "NC" Then
    If Date2Num(Date) >= Date2Num("05/01/2007") Then
      If Not Exist(OrbitHeader) Then
        MsgBox ("Before continuing please save the NC ORBIT data located on the Employer File screen on the Control Maintenance Menu.")
        Exit Sub
      ElseIf Not Exist(OrbitEmpData) Then
        MsgBox ("No employees have been set up for the NC ORBIT program. Employees are set up for the NC ORBIT program on the Employee Maintenance screen.")
        Exit Sub
      Else
        frmORBITMenu.Show
      End If
    Else
      frmRetRpt.Show
    End If
  End If
  DoEvents
  Unload frmReportsProcessing
End Sub

Private Sub cmdSEPPCon_Click()
  InFileNames(1) = "PRDATA\PRUNIT.DAT"
  InFileNames(2) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 2) = False Then
    Close
    Exit Sub
  End If
  frmSEPPCon.Show
  DoEvents
  Unload frmReportsProcessing
End Sub

Private Sub cmdSupRetRpts_Click()
  InFileNames(1) = "PRDATA\PRUNIT.DAT"
  InFileNames(2) = "PRDATA\PREMP2.DAT"
  InFileNames(3) = "PRDATA\PRDEDCOD.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 3) = False Then
    Close
    Exit Sub
  End If
  frmSupRetReport.Show
  DoEvents
  Unload frmReportsProcessing
End Sub

Private Sub cmdTaxFringe_Click()
  frmTaxFringeRpt.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdYTDWageDis_Click()
  InFileNames(1) = "PRDATA\PRUNIT.DAT"
  InFileNames(2) = "PRDATA\PREMP2.DAT"
  InFileNames(3) = "PRDATA\PRSYS.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 3) = False Then
    Close
    Exit Sub
  End If
  frmYTDWageDist.Show
  DoEvents
  Unload frmReportsProcessing
End Sub

Private Sub cmdExit_Click()
 
  m_roOption = roOff
  frmPayrollMainMenu.Show
  DoEvents
  Unload frmReportsProcessing
   
End Sub

Private Sub EmployeeEarningsHistCmmd_Click()
  InFileNames(1) = "PRDATA\PRUNIT.DAT"
  InFileNames(2) = "PRDATA\PREMP2.DAT"
  InFileNames(3) = "PRDATA\PRDEDCOD.DAT"
  InFileNames(4) = "PRDATA\PRERNCOD.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 4) = False Then
    Close
    Exit Sub
  End If
  frmEmpHistRptSplash.Show
  DoEvents
  Unload frmReportsProcessing
End Sub

Private Sub ESCReportCmmd_Click()
  InFileNames(1) = "PRDATA\PRUNIT.DAT"
  InFileNames(2) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 2) = False Then
    Close
    Exit Sub
  End If
  frmESCQrtRpt.Show
  DoEvents
  Unload frmReportsProcessing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"
      Call cmdExit_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Me.HelpContextID = hlpPayrollReports
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  m_roOption = roOff '
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
'  Command1.Visible = False
'  Command2.Visible = False
'  Command3.Visible = False
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub fpBtn1_Click()

End Sub

Private Sub GrossWageReportCmmd_Click()
  InFileNames(1) = "PRDATA\PRUNIT.DAT"
  InFileNames(2) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 2) = False Then
    Close
    Exit Sub
  End If
  frmGrossWageReport.Show
  DoEvents
  Unload frmReportsProcessing
End Sub

Private Sub PayrollDeductionsTakenCmmd_Click()
  InFileNames(1) = "PRDATA\PRUNIT.DAT"
  InFileNames(2) = "PRDATA\PREMP2.DAT"
  InFileNames(3) = "PRDATA\PRDEDCOD.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 3) = False Then
    Close
    Exit Sub
  End If
  frmPRDeduction.Show
  DoEvents
  Unload frmReportsProcessing
End Sub

Private Sub PrintEmployeeDataFileCmmd_Click()
  m_roOption = roOn
  frmEmpDataPrint.Show
  DoEvents
  Unload frmReportsProcessing
End Sub

Private Sub TerminatedEmployeesCmmd_Click()
  m_roOption = roOn
  Call frmEmployeeMaintMenu.PrintTerminatedEmplListCmmd_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmReportsProcessing.")
      Call Terminate
      End
    End If
  End If
End Sub

