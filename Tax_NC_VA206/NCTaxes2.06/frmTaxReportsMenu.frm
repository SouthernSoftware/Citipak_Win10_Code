VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmTaxReportsMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Billing Reports Menu"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTaxReportsMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdExpReal 
      Height          =   420
      Left            =   5880
      TabIndex        =   15
      Top             =   6375
      Width           =   2655
      _Version        =   131072
      _ExtentX        =   4683
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmTaxReportsMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAdvRpt 
      Height          =   456
      Left            =   3120
      TabIndex        =   12
      Top             =   5820
      Width           =   2652
      _Version        =   131072
      _ExtentX        =   4678
      _ExtentY        =   804
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
      ButtonDesigner  =   "frmTaxReportsMenu.frx":0AB3
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMastCustList 
      Height          =   444
      Left            =   3120
      TabIndex        =   2
      Top             =   2962
      Width           =   2652
      _Version        =   131072
      _ExtentX        =   4678
      _ExtentY        =   783
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
      ButtonDesigner  =   "frmTaxReportsMenu.frx":0C9F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMailLbls 
      Height          =   420
      Left            =   5880
      TabIndex        =   5
      Top             =   3555
      Width           =   2655
      _Version        =   131072
      _ExtentX        =   4683
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmTaxReportsMenu.frx":0E8A
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMasterVal 
      Height          =   420
      Left            =   3120
      TabIndex        =   4
      Top             =   3555
      Width           =   2655
      _Version        =   131072
      _ExtentX        =   4683
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmTaxReportsMenu.frx":106C
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdTransJrnl 
      Height          =   432
      Left            =   3120
      TabIndex        =   6
      Top             =   4110
      Width           =   2652
      _Version        =   131072
      _ExtentX        =   4678
      _ExtentY        =   762
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
      ButtonDesigner  =   "frmTaxReportsMenu.frx":1258
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMastBalList 
      Height          =   465
      Left            =   5880
      TabIndex        =   3
      Top             =   2955
      Width           =   2655
      _Version        =   131072
      _ExtentX        =   4683
      _ExtentY        =   820
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
      ButtonDesigner  =   "frmTaxReportsMenu.frx":143F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdLateList 
      Height          =   444
      Left            =   3120
      TabIndex        =   8
      Top             =   4672
      Width           =   2652
      _Version        =   131072
      _ExtentX        =   4678
      _ExtentY        =   783
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
      ButtonDesigner  =   "frmTaxReportsMenu.frx":1629
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMortCodeRpt 
      Height          =   456
      Left            =   5880
      TabIndex        =   7
      Top             =   4110
      Width           =   2652
      _Version        =   131072
      _ExtentX        =   4678
      _ExtentY        =   804
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
      ButtonDesigner  =   "frmTaxReportsMenu.frx":1809
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCustInq 
      Height          =   432
      Left            =   3120
      TabIndex        =   0
      Top             =   2400
      Width           =   2652
      _Version        =   131072
      _ExtentX        =   4678
      _ExtentY        =   762
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
      ButtonDesigner  =   "frmTaxReportsMenu.frx":19F1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCustTransHist 
      Height          =   432
      Left            =   5880
      TabIndex        =   1
      Top             =   2400
      Width           =   2652
      _Version        =   131072
      _ExtentX        =   4678
      _ExtentY        =   762
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
      ButtonDesigner  =   "frmTaxReportsMenu.frx":1BD5
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSenCtzns 
      Height          =   420
      Left            =   3120
      TabIndex        =   10
      Top             =   5265
      Width           =   2655
      _Version        =   131072
      _ExtentX        =   4683
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmTaxReportsMenu.frx":1DBF
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintAbs 
      Height          =   465
      Left            =   5880
      TabIndex        =   9
      Top             =   4665
      Width           =   2655
      _Version        =   131072
      _ExtentX        =   4683
      _ExtentY        =   820
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
      ButtonDesigner  =   "frmTaxReportsMenu.frx":1FA6
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdRealHist 
      Height          =   420
      Left            =   5880
      TabIndex        =   11
      Top             =   5265
      Width           =   2655
      _Version        =   131072
      _ExtentX        =   4683
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmTaxReportsMenu.frx":2189
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCollRate 
      Height          =   432
      Left            =   5880
      TabIndex        =   13
      Top             =   5820
      Width           =   2652
      _Version        =   131072
      _ExtentX        =   4678
      _ExtentY        =   762
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
      ButtonDesigner  =   "frmTaxReportsMenu.frx":2372
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCustExp 
      Height          =   420
      Left            =   3120
      TabIndex        =   14
      Top             =   6375
      Width           =   2655
      _Version        =   131072
      _ExtentX        =   4683
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmTaxReportsMenu.frx":2556
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   420
      Left            =   4560
      TabIndex        =   18
      Top             =   7470
      Width           =   2655
      _Version        =   131072
      _ExtentX        =   4683
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmTaxReportsMenu.frx":273E
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDiscsGiven 
      Height          =   432
      Left            =   3120
      TabIndex        =   16
      Top             =   6920
      Width           =   2652
      _Version        =   131072
      _ExtentX        =   4678
      _ExtentY        =   762
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
      ButtonDesigner  =   "frmTaxReportsMenu.frx":291B
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdIndComPrvRpt 
      Height          =   432
      Left            =   5880
      TabIndex        =   17
      Top             =   6920
      Width           =   2652
      _Version        =   131072
      _ExtentX        =   4678
      _ExtentY        =   762
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
      ButtonDesigner  =   "frmTaxReportsMenu.frx":2B00
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1098
      Index           =   1
      Left            =   1493
      Top             =   813
      Width           =   8655
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   2094
      Top             =   2019
      Width           =   971
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8706
      X2              =   8706
      Y1              =   2127
      Y2              =   8028
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   8586
      Top             =   2019
      Width           =   971
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8706
      X2              =   9408
      Y1              =   8020
      Y2              =   8020
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2199
      X2              =   2914
      Y1              =   8020
      Y2              =   8020
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2214
      X2              =   2214
      Y1              =   2127
      Y2              =   8015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TAX REPORTS MENU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2813
      TabIndex        =   19
      Top             =   1164
      Width           =   6012
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1214
      Left            =   1495
      Top             =   687
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   2094
      Top             =   1886
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5910
      Index           =   0
      Left            =   2213
      Top             =   2117
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8585
      Top             =   1887
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5910
      Index           =   1
      Left            =   8706
      Top             =   2117
      Width           =   732
   End
End
Attribute VB_Name = "frmTaxReportsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim Break As Integer

Private Sub cmdAdvRpt_Click()
  frmTaxAdColRpt.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdCollRate_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Close TTHandle
  
  If NumOfTTRecs = 0 Then
    Call TaxMsg(900, "No tax transactions have been saved.")
    Exit Sub
  End If
  
  frmTaxCollectRateRpt.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdCustExp_Click()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  If NumOfTCRecs = 0 Then
    Call TaxMsg(900, "There are no tax customer records saved.")
    Close TCHandle
    Exit Sub
  End If
  
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TaxCust.Deleted = 0 Then
      Exit For
    End If
  Next x
  Close TCHandle
  
  If x > NumOfTCRecs Then
    Call TaxMsg(900, "There are no valid tax customer records saved.")
    Exit Sub
  End If
    
  frmTaxExpCustInfo.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdCustInq_Click()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  If NumOfTCRecs = 0 Then
    Call TaxMsg(900, "There are no tax customer records saved.")
    Close TCHandle
    Exit Sub
  End If
  
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TaxCust.Deleted = 0 Then
      Exit For
    End If
  Next x
  Close TCHandle
  
  If x > NumOfTCRecs Then
    Call TaxMsg(900, "There are no valid tax customer records saved.")
    Exit Sub
  End If
    
  frmTaxCustInq.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdCustTransHist_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  If NumOfTCRecs = 0 Then
    Call TaxMsg(900, "There are no tax customer records saved.")
    Close TCHandle
    Exit Sub
  End If
  
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TaxCust.Deleted = 0 Then
      Exit For
    End If
  Next x
  Close TCHandle
  
  If x > NumOfTCRecs Then
    Call TaxMsg(900, "There are no valid tax customer records saved.")
    Exit Sub
  End If
    
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Close TTHandle
  If NumOfTTRecs = 0 Then
    Call TaxMsg(900, "There are no transactions saved.")
    Exit Sub
  End If
  
  frmTaxCustTHistRpt.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdDiscsGiven_Click()
  frmTaxDiscountsGiven.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdExit_Click()
  frmTaxMainMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdExpReal_Click()
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim x As Long
  
  OpenRealPropFile RHandle, NumOfRealRecs
  If NumOfRealRecs = 0 Then
    Call TaxMsg(900, "There are no real property records saved.")
    Close RHandle
    Exit Sub
  End If
  For x = 1 To NumOfRealRecs
    Get RHandle, x, RealRec
    If RealRec.Deleted = 0 Then
      Exit For
    End If
  Next x
  Close RHandle
  
  If x > NumOfRealRecs Then
    Call TaxMsg(900, "There are no valid real property records saved.")
    Exit Sub
  End If
    
  frmTaxExpRealInfo.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdIndComPrvRpt_Click()
  frmTaxRealClassRpt.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdLateList_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  If NumOfTCRecs = 0 Then
    Call TaxMsg(900, "There are no tax customer records saved.")
    Close TCHandle
    Exit Sub
  End If
  
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TaxCust.Deleted = 0 Then
      Exit For
    End If
  Next x
  Close TCHandle
  
  If x > NumOfTCRecs Then
    Call TaxMsg(900, "There are no valid tax customer records saved.")
    Exit Sub
  End If
    
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Close TTHandle
  If NumOfTTRecs = 0 Then
    Call TaxMsg(900, "There are no transactions saved.")
    Exit Sub
  End If
  
  frmTaxLateListRpt.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdMailLbls_Click()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  If NumOfTCRecs = 0 Then
    Call TaxMsg(900, "There are no tax customer records saved.")
    Close TCHandle
    Exit Sub
  End If
  
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TaxCust.Deleted = 0 Then
      Exit For
    End If
  Next x
  Close TCHandle
  
  If x > NumOfTCRecs Then
    Call TaxMsg(900, "There are no valid tax customer records saved.")
    Exit Sub
  End If
    
  frmTaxMailingLblsGeneral.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdMastBalList_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Close TTHandle
  
  If NumOfTTRecs = 0 Then
    Call TaxMsg(900, "There are no transaction records saved.")
    Exit Sub
  End If
  
  frmTaxMasterBalList.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdMastCustList_Click()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  If NumOfTCRecs = 0 Then
    Call TaxMsg(900, "There are no tax customer records saved.")
    Close TCHandle
    Exit Sub
  End If
  
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TaxCust.Deleted = 0 Then
      Exit For
    End If
  Next x
  Close TCHandle
  
  If x > NumOfTCRecs Then
    Call TaxMsg(900, "There are no valid tax customer records saved.")
    Exit Sub
  End If
    
  frmTaxCustListRpt.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdMasterVal_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  If NumOfTCRecs = 0 Then
    Call TaxMsg(900, "There are no tax customer records saved.")
    Close TCHandle
    Exit Sub
  End If
  
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TaxCust.Deleted = 0 Then
      Exit For
    End If
  Next x
  Close TCHandle
  
  If x > NumOfTCRecs Then
    Call TaxMsg(900, "There are no valid tax customer records saved.")
    Exit Sub
  End If
    
'  OpenTaxTransFile TTHandle, NumOfTTRecs
'  Close TTHandle
'  If NumOfTTRecs = 0 Then
'    Call TaxMsg(900, "There are no transactions saved.")
'    Exit Sub
'  End If
  
  frmTaxValuationListing.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdMortCodeRpt_Click()
  Dim MortRec As MortCodeRecType
  Dim MHandle As Integer
  Dim NumOfMCodes As Integer
  Dim x As Integer
  Dim MortCnt As Integer
  
  OpenMortCodeFile MHandle, NumOfMCodes
  
  If NumOfMCodes = 0 Then
    Call TaxMsg(900, "There are no mortgage codes saved.")
    Close MHandle
    Exit Sub
  End If
  
  For x = 1 To NumOfMCodes
    Get MHandle, x, MortRec
    If MortRec.Deleted = 0 Then
      Exit For
    End If
  Next x
  
  If x > NumOfMCodes Then
    Call TaxMsg(900, "There are no valid mortgage codes saved.")
    Close MHandle
    Exit Sub
  End If
  Close MHandle
  
  frmTaxMortCodeRpt.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPrintAbs_Click()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  If NumOfTCRecs = 0 Then
    Call TaxMsg(900, "There are no tax customer records saved.")
    Close TCHandle
    Exit Sub
  End If
  
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TaxCust.Deleted = 0 Then
      Exit For
    End If
  Next x
  Close TCHandle
  
  If x > NumOfTCRecs Then
    Call TaxMsg(900, "There are no valid tax customer records saved.")
    Exit Sub
  End If
    
  frmTaxAbstractRpt.Show
  
'  frmTaxReportMoreOpts.Show vbModal
'  If frmTaxReportMoreOpts.fptxtPrintType.Text = "Graphical" Then
'    If frmTaxReportMoreOpts.fptxtBreak.Text = "1" Then Break = 1 Else Break = 2
'    Unload frmTaxReportOpt
'    Call PrintGraphicsAbRpt
'  ElseIf frmTaxReportMoreOpts.fptxtPrintType.Text = "Text" Then
'    If frmTaxReportMoreOpts.fptxtBreak.Text = "1" Then Break = 1 Else Break = 2
'    frmTaxMsg.Label1.Caption = "Pitch 12 is recommended for this report."
'    frmTaxMsg.Label1.Top = 900
'    frmTaxMsg.Show vbModal
'    Unload frmTaxReportMoreOpts
'    Call PrintTextAbRpt
'  End If
  
End Sub

Private Sub cmdRealHist_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  If NumOfTCRecs = 0 Then
    Call TaxMsg(900, "There are no tax customer records saved.")
    Close TCHandle
    Exit Sub
  End If
  
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TaxCust.Deleted = 0 Then
      Exit For
    End If
  Next x
  Close TCHandle
  
  If x > NumOfTCRecs Then
    Call TaxMsg(900, "There are no valid tax customer records saved.")
    Exit Sub
  End If
    
  OpenRealPropFile RHandle, NumOfRRecs
  If NumOfRRecs = 0 Then
    Call TaxMsg(900, "There are no real property records saved.")
    Close RHandle
    Exit Sub
  End If
  
  For x = 1 To NumOfRRecs
    Get RHandle, x, RealRec
    If RealRec.Deleted = 0 Then
      Exit For
    End If
  Next x
  Close RHandle
  
  If x > NumOfRRecs Then
    Call TaxMsg(900, "There are no valid real property records saved.")
    Exit Sub
  End If
    
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Close TTHandle
  If NumOfTTRecs = 0 Then
    Call TaxMsg(900, "There are no transactions saved.")
    Exit Sub
  End If
  
  frmTaxRealPropHist.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdSenCtzns_Click()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  If NumOfTCRecs = 0 Then
    Call TaxMsg(900, "There are no tax customer records saved.")
    Close TCHandle
    Exit Sub
  End If
  
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TaxCust.Deleted = 0 Then
      Exit For
    End If
  Next x
  Close TCHandle
  
  If x > NumOfTCRecs Then
    Call TaxMsg(900, "There are no valid tax customer records saved.")
    Exit Sub
  End If
    
  frmTaxSeniorDscRpt.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdTransJrnl_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  If NumOfTCRecs = 0 Then
    Call TaxMsg(900, "There are no tax customer records saved.")
    Close TCHandle
    Exit Sub
  End If
  
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TaxCust.Deleted = 0 Then
      Exit For
    End If
  Next x
  Close TCHandle
  
  If x > NumOfTCRecs Then
    Call TaxMsg(900, "There are no valid tax customer records saved.")
    Exit Sub
  End If
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Close TTHandle
  If NumOfTTRecs = 0 Then
    Call TaxMsg(900, "There are no transactions saved.")
    Exit Sub
  End If
  
  frmTaxTransJournal.Show
  DoEvents
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%x"
      Call cmdExit_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpTaxReportsMenu
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxReportsMenu.")
      Call Terminate
      End
    End If
  End If

End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub PrintGraphicsAbRpt()
  Dim x As Long
  Dim y As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim PropRec As PropertyRecType
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim NextRec As Long
  Dim RealVal As Double
  Dim PersVal As Double
  Dim TotVal As Double
  Dim PrintHeader As Boolean
  Dim RealCnt As Integer
  Dim PersCnt As Integer
  Dim Town$
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim PrintDesc As Boolean
  Dim dlm$
  Dim PCnt As Long
  
  'on error goto ERRORSTUFF
  
  dlm$ = "~"
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  Town$ = QPTrim$(TaxMasterRec.Name)
  
  RptFile$ = "TAXRPTS\ABSTLIST.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle

  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenRealPropFile RHandle, NumOfRealRecs
  OpenPersPropFile PHandle, NumOfPersRecs
  
  frmTaxShowPctComp.Label1 = "Gathering Property Data"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
    RealVal = 0
    PersVal = 0
    TotVal = 0
    RealCnt = 0
    PersCnt = 0
    'look for valid property for this customer
    If TaxCust.FirstPropRec > 0 Then
      NextRec = TaxCust.FirstPropRec
      Do While NextRec > 0
        Get RHandle, NextRec, RealRec
        If RealRec.Deleted = -1 Then GoTo NotThisOne
        RealCnt = RealCnt + 1
NotThisOne:
        NextRec = RealRec.NextRec
      Loop
    End If
    If TaxCust.FirstPersRec > 0 Then
      NextRec = TaxCust.FirstPersRec
      Do While NextRec > 0
        Get PHandle, NextRec, PersRec
        If PersRec.Deleted = -1 Then GoTo NotThisPers
        PersCnt = PersCnt + 1
NotThisPers:
        NextRec = PersRec.NextRec
      Loop
    End If
    
    If RealCnt = 0 Then GoTo NoReal
    If TaxCust.FirstPropRec > 0 Then
      NextRec = TaxCust.FirstPropRec
      Do While NextRec > 0
        Get RHandle, NextRec, RealRec
        If RealRec.Deleted = -1 Then GoTo RealDeleted
        GoSub PrintReal
RealDeleted:
        NextRec = RealRec.NextRec
      Loop
    End If
NoReal:
    If PersCnt > 0 Then
      NextRec = TaxCust.FirstPersRec
      Do While NextRec > 0
        Get PHandle, NextRec, PersRec
        If PersRec.Deleted = -1 Then GoTo PersDeleted
        GoSub PrintPers
PersDeleted:
        NextRec = PersRec.NextRec
      Loop
    End If
SkipIt:
    frmTaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  Close
  
  If PCnt = 0 Then
    Call TaxMsg(900, "There are no properties to report.")
    Exit Sub
  End If
  
  arTaxAbstractRpt.Show
  DoEvents
  
  Exit Sub
  
PrintReal:
  '                   0                     1                   2
  Print #RptHandle, Town$; dlm; QPTrim$(TaxCust.CustName); dlm; x; dlm;
  '                             3                           4
  Print #RptHandle, QPTrim$(TaxCust.Addr1); dlm; QPTrim$(TaxCust.Addr2); dlm;
  '                                                 5
  Print #RptHandle, QPTrim$(TaxCust.City) + ", " + QPTrim$(TaxCust.State) + "  " + QPTrim$(TaxCust.Zip); dlm;
  '                   6                      7                         8
  Print #RptHandle, "REAL"; dlm; QPTrim$(RealRec.RealPin); dlm; RealRec.PROPVALU; dlm;
  '                             9                             10
  Print #RptHandle, QPTrim$(RealRec.PropAddr); dlm; QPTrim$(RealRec.Map) + "/" + QPTrim$(RealRec.BLOCK) + "/" + QPTrim$(RealRec.LOTNUMB); dlm;
  '                        11                      12                   13
  Print #RptHandle, RealRec.PROPNOT1; dlm; RealRec.PROPNOT2; dlm; RealRec.PROPNOT3; dlm;
  '                 14       15       16
  Print #RptHandle, ""; dlm; ""; dlm; ""; dlm;
  '                 17       18
  Print #RptHandle, ""; dlm; ""; dlm;
  '                 19       20
  Print #RptHandle, ""; dlm; ""; dlm;
  '                 21       22
  Print #RptHandle, ""; dlm; ""; dlm;
  '                 23       24        25
  Print #RptHandle, ""; dlm; ""; dlm; Break
  
  PCnt = PCnt + 1
  
  Return
  
PrintPers:
  '                   0                     1                   2
  Print #RptHandle, Town$; dlm; QPTrim$(TaxCust.CustName); dlm; x; dlm;
  '                             3                           4
  Print #RptHandle, QPTrim$(TaxCust.Addr1); dlm; QPTrim$(TaxCust.Addr2); dlm;
  '                                                 5
  Print #RptHandle, QPTrim$(TaxCust.City) + ", " + QPTrim$(TaxCust.State) + "  " + QPTrim$(TaxCust.Zip); dlm;
  '                     6                       7                   8
  Print #RptHandle, "PERSONAL"; dlm; QPTrim$(PersRec.PropPin); dlm; ""; dlm;
  '                 9        10
  Print #RptHandle, ""; dlm; ""; dlm;
  '                 11       12       13
  Print #RptHandle, ""; dlm; ""; dlm; ""; dlm;
  '                       14                    15
  Print #RptHandle, PersRec.PersVal; dlm; PersRec.CVALUE; dlm;
  '                        16                   17
  Print #RptHandle, PersRec.MHVALUE; dlm; PersRec.MTVALUE; dlm;
  '                       18                   19
  Print #RptHandle, PersRec.MCVALUE; dlm; PersRec.DESC1; dlm;
  '                      20                   21
  Print #RptHandle, PersRec.DESC2; dlm; PersRec.DESC3; dlm;
  '                             22                  23               24         25
  Print #RptHandle, QPTrim$(PersRec.Desc4); dlm; PersRec.Desc5; dlm; ""; dlm; Break
  
  PCnt = PCnt + 1
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxReportsMenu", "PrintGraphicsAbRpt", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate
End Sub

Private Sub PrintTextAbRpt()
  Dim x As Long
  Dim y As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim PropRec As PropertyRecType
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim FF$
  Dim Page As Integer
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim RptFile$
  Dim RptHandle As Integer
  Dim NextRec As Long
  Dim RealVal As Double
  Dim PersVal As Double
  Dim TotVal As Double
  Dim PrintHeader As Boolean
  Dim RealCnt As Integer
  Dim PersCnt As Integer
  Dim Town$
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim PrintDesc As Boolean
  Dim PCnt As Long
  Dim NewOne As Boolean
  
  'on error goto ERRORSTUFF
  
  NewOne = False
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  Town$ = QPTrim$(TaxMasterRec.Name)
  
  FF$ = Chr(12)
  MaxLines = 58
  
  RptFile$ = "TAXRPTS\ABSTLIST.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle

  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenRealPropFile RHandle, NumOfRealRecs
  OpenPersPropFile PHandle, NumOfPersRecs
  
  frmTaxShowPctComp.Label1 = "Gathering Property Data"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
    RealVal = 0
    PersVal = 0
    TotVal = 0
    RealCnt = 0
    PersCnt = 0
    NewOne = True
    'look for valid property for this customer
    If TaxCust.FirstPropRec > 0 Then
      NextRec = TaxCust.FirstPropRec
      Do While NextRec > 0
        Get RHandle, NextRec, RealRec
        If RealRec.Deleted = -1 Then GoTo NotThisOne
        RealCnt = RealCnt + 1
NotThisOne:
        NextRec = RealRec.NextRec
      Loop
    End If
    If TaxCust.FirstPersRec > 0 Then
      NextRec = TaxCust.FirstPersRec
      Do While NextRec > 0
        Get PHandle, NextRec, PersRec
        If PersRec.Deleted = -1 Then GoTo NotThisPers
        PersCnt = PersCnt + 1
NotThisPers:
        NextRec = PersRec.NextRec
      Loop
    End If
    If RealCnt > 0 Or PersCnt > 0 Then
      If LineCnt <> 7 Then
        Print #RptHandle, FF$
      End If
      GoSub PrintCustHeader
    End If
    
    If RealCnt = 0 Then GoTo NoReal
    If TaxCust.FirstPropRec > 0 Then
      NextRec = TaxCust.FirstPropRec
      Do While NextRec > 0
        If NewOne = False And Break = 2 Then
          Print #RptHandle, FF$
          GoSub PrintCustHeader
        End If
        Get RHandle, NextRec, RealRec
        If RealRec.Deleted = -1 Then GoTo RealDeleted
        GoSub PrintReal
RealDeleted:
        NewOne = False
        NextRec = RealRec.NextRec
      Loop
    End If
NoReal:
     If PersCnt > 0 Then
       NextRec = TaxCust.FirstPersRec
       Do While NextRec > 0
         If NewOne = False And Break = 2 Then
           Print #RptHandle, FF$
           GoSub PrintCustHeader
         End If
         Get PHandle, NextRec, PersRec
         If PersRec.Deleted = -1 Then GoTo PersDeleted
         GoSub PrintPers
PersDeleted:
         NextRec = PersRec.NextRec
         NewOne = False
       Loop
     End If
     If PersCnt > 0 Or RealCnt > 0 Then
       Print #RptHandle, String$(84, "=")
       LineCnt = LineCnt + 1
       If LineCnt > MaxLines Then
         Print #RptHandle, FF$
         LineCnt = 0
       End If
     End If
SkipIt:
    frmTaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  Print #RptHandle, FF$
  
  Close
  
  If PCnt = 0 Then
    Call TaxMsg(900, "There are no properties to report.")
    Exit Sub
  End If
  
  ViewPrint RptFile, "Property Listing", True
  
  Exit Sub
  
PrintCustHeader:
  Print #RptHandle, "Abstract Listing: " + Town$
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, "Account of: "; Tab(15); QPTrim$(TaxCust.CustName); Tab(67); "Acct #: " + Using$("####0", x)
  Print #RptHandle, Tab(15); QPTrim$(TaxCust.Addr1)
  Print #RptHandle, Tab(15); QPTrim$(TaxCust.Addr2)
  Print #RptHandle, Tab(15); QPTrim$(TaxCust.City) + ", " + QPTrim$(TaxCust.State) + "  " + QPTrim$(TaxCust.Zip)
  Print #RptHandle, Tab(5); String(79, "-")
  LineCnt = 7
  
  Return
  
PrintReal:
  PrintDesc = False
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintCustHeader
    LineCnt = 0
  End If
  Print #RptHandle, Tab(5); "*** REAL PROPERTY ***"
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintCustHeader
    LineCnt = 0
  End If
  Print #RptHandle, Tab(10); "Address:"; Tab(22); QPTrim$(RealRec.PropAddr)
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintCustHeader
    LineCnt = 0
  End If
  
  Print #RptHandle, Tab(10); "PIN #"; Tab(22); QPTrim$(RealRec.RealPin); Tab(45); "MAP/BLOCK/LOT: "; Tab(62); QPTrim$(RealRec.Map) + "/" + QPTrim$(RealRec.BLOCK) + "/" + QPTrim$(RealRec.LOTNUMB)
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintCustHeader
    LineCnt = 0
  End If
  Print #RptHandle, Tab(10); "REAL VALUE:"; Tab(22); Using$("$###,###,##0.00", RealRec.PROPVALU)
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintCustHeader
    LineCnt = 0
  End If
  If QPTrim$(RealRec.PROPNOT1) = "" And QPTrim$(RealRec.PROPNOT2) = "" And QPTrim$(RealRec.PROPNOT3) = "" Then
    Print #RptHandle, Tab(10); "DESC:"; Tab(22); "NO DESCRIPTION AVAILABLE"
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintCustHeader
      LineCnt = 0
    End If
    GoTo NoRealDesc
  End If
  If QPTrim$(RealRec.PROPNOT1) <> "" Then
    Print #RptHandle, Tab(10); "DESC:"; Tab(22); RealRec.PROPNOT1
    PrintDesc = True
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintCustHeader
      LineCnt = 0
    End If
  End If
    
  If QPTrim$(RealRec.PROPNOT2) <> "" Then
    If PrintDesc = False Then
      Print #RptHandle, Tab(10); "DESC:"; Tab(22); RealRec.PROPNOT2
      PrintDesc = True
    Else
      Print #RptHandle, Tab(22); RealRec.PROPNOT2
    End If
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintCustHeader
      LineCnt = 0
    End If
  End If
  
  If QPTrim$(RealRec.PROPNOT3) <> "" Then
    If PrintDesc = False Then
      Print #RptHandle, Tab(10); "DESC:"; Tab(22); RealRec.PROPNOT3
      PrintDesc = True
    Else
      Print #RptHandle, Tab(22); RealRec.PROPNOT3
    End If
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintCustHeader
      LineCnt = 0
    End If
  End If
  
  Print #RptHandle,
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintCustHeader
    LineCnt = 0
  End If
NoRealDesc:

  PCnt = PCnt + 1
  
  Return
  
PrintPers:
  PrintDesc = False
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintCustHeader
    LineCnt = 0
  End If
  Print #RptHandle, Tab(5); "*** PERSONAL PROPERTY ***"
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintCustHeader
    LineCnt = 0
  End If
  Print #RptHandle, Tab(5); "PIN #"; Tab(22); QPTrim$(PersRec.PropPin)
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintCustHeader
    LineCnt = 0
  End If
  Print #RptHandle, Tab(5); "VALUE AMOUNTS"
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintCustHeader
    LineCnt = 0
  End If
  Print #RptHandle, Tab(10); "PERSONAL:"; Tab(30); Using$("$###,###,##0.00", PersRec.PersVal)
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintCustHeader
    LineCnt = 0
  End If
  Print #RptHandle, Tab(10); "FARM EQUIPMENT:"; Tab(30); Using$("$###,###,##0.00", PersRec.CVALUE)
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintCustHeader
    LineCnt = 0
  End If
  Print #RptHandle, Tab(10); "MOBILE HOMES:"; Tab(30); Using$("$###,###,##0.00", PersRec.MHVALUE)
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintCustHeader
    LineCnt = 0
  End If
  Print #RptHandle, Tab(10); "MACHINE/TOOLS:"; Tab(30); Using$("$###,###,##0.00", PersRec.MTVALUE)
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintCustHeader
    LineCnt = 0
  End If
  Print #RptHandle, Tab(10); "MERCHANT CAPITAL:"; Tab(30); Using$("$###,###,##0.00", PersRec.MCVALUE)
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintCustHeader
    LineCnt = 0
  End If
  
  If QPTrim$(PersRec.DESC1) = "" And QPTrim$(PersRec.DESC2) = "" And QPTrim$(PersRec.DESC3) = "" And QPTrim$(PersRec.Desc4) = "" And QPTrim$(PersRec.Desc5) = "" Then
    Print #RptHandle, Tab(10); "DESCRIPTION:"; Tab(30); "NO DESCRIPTION AVAILABLE"
    GoTo NoPersDesc
  End If
  
  If QPTrim$(PersRec.DESC1) <> "" Then
    PrintDesc = True
    Print #RptHandle, Tab(10); "DESCRIPTION:"; Tab(30); PersRec.DESC1
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintCustHeader
      LineCnt = 0
    End If
  End If
    
  If QPTrim$(PersRec.DESC2) <> "" Then
    If PrintDesc = False Then
      Print #RptHandle, Tab(10); "DESCRIPTION:"; Tab(30); PersRec.DESC2
      PrintDesc = True
    Else
      Print #RptHandle, Tab(30); PersRec.DESC2
    End If
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintCustHeader
      LineCnt = 0
    End If
  End If
  
  If QPTrim$(PersRec.DESC3) <> "" Then
    If PrintDesc = False Then
      Print #RptHandle, Tab(10); "DESCRIPTION:"; Tab(30); PersRec.DESC3
      PrintDesc = True
    Else
      Print #RptHandle, Tab(30); PersRec.DESC3
    End If
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintCustHeader
      LineCnt = 0
    End If
  End If
  
  If QPTrim$(PersRec.Desc4) <> "" Then
    If PrintDesc = False Then
      Print #RptHandle, Tab(10); "DESCRIPTION:"; Tab(30); PersRec.Desc4
      PrintDesc = True
    Else
      Print #RptHandle, Tab(30); PersRec.Desc4
    End If
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintCustHeader
      LineCnt = 0
    End If
  End If
  
  If QPTrim$(PersRec.Desc5) <> "" Then
    If PrintDesc = False Then
      Print #RptHandle, Tab(10); "DESCRIPTION:"; Tab(30); PersRec.Desc5
      PrintDesc = True
    Else
      Print #RptHandle, Tab(30); PersRec.Desc5
    End If
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintCustHeader
      LineCnt = 0
    End If
  End If
  
  Print #RptHandle,
  LineCnt = LineCnt + 1
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintCustHeader
    LineCnt = 0
  End If
NoPersDesc:

  PCnt = PCnt + 1

  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxReportsMenu", "PrintTextAbRpt", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate
  
End Sub

