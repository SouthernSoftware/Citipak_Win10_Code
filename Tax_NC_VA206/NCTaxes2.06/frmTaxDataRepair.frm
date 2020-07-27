VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmTaxDataRepair 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Repair DOS Data"
   ClientHeight    =   9360
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "frmTaxDataRepair.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox tbxTransNum 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   9960
      TabIndex        =   84
      Top             =   7080
      Width           =   1455
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdUpdateAdd1andAdd2Long 
      Height          =   360
      Left            =   6840
      TabIndex        =   82
      Top             =   8280
      Width           =   2895
      _Version        =   131072
      _ExtentX        =   5106
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixRealTransIEAds 
      Height          =   375
      Left            =   6840
      TabIndex        =   77
      Top             =   6840
      Width           =   2895
      _Version        =   131072
      _ExtentX        =   5106
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":0AB4
      Begin fpBtnAtlLibCtl.fpBtn fpBtn24 
         Height          =   375
         Left            =   0
         TabIndex        =   78
         Top             =   480
         Width           =   2895
         _Version        =   131072
         _ExtentX        =   5106
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
         ButtonDesigner  =   "frmTaxDataRepair.frx":0C9C
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdStringOrNumber 
      Height          =   375
      Left            =   3840
      TabIndex        =   69
      Top             =   8280
      Width           =   2895
      _Version        =   131072
      _ExtentX        =   5106
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":0E77
      Begin fpBtnAtlLibCtl.fpBtn fpBtn20 
         Height          =   375
         Left            =   0
         TabIndex        =   70
         Top             =   5040
         Width           =   2895
         _Version        =   131072
         _ExtentX        =   5106
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
         ButtonDesigner  =   "frmTaxDataRepair.frx":105D
      End
      Begin fpBtnAtlLibCtl.fpBtn fpBtn21 
         Height          =   375
         Left            =   0
         TabIndex        =   71
         Top             =   480
         Width           =   2895
         _Version        =   131072
         _ExtentX        =   5106
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
         ButtonDesigner  =   "frmTaxDataRepair.frx":1238
         Begin fpBtnAtlLibCtl.fpBtn fpBtn22 
            Height          =   375
            Left            =   0
            TabIndex        =   72
            Top             =   5040
            Width           =   2895
            _Version        =   131072
            _ExtentX        =   5106
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
            ButtonDesigner  =   "frmTaxDataRepair.frx":141A
         End
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixHildebran 
      Height          =   375
      Left            =   3840
      TabIndex        =   53
      Top             =   7320
      Width           =   2895
      _Version        =   131072
      _ExtentX        =   5106
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":15F5
      Begin fpBtnAtlLibCtl.fpBtn fpBtn6 
         Height          =   375
         Left            =   0
         TabIndex        =   54
         Top             =   5040
         Width           =   2895
         _Version        =   131072
         _ExtentX        =   5106
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
         ButtonDesigner  =   "frmTaxDataRepair.frx":17D5
      End
      Begin fpBtnAtlLibCtl.fpBtn fpBtn7 
         Height          =   375
         Left            =   0
         TabIndex        =   55
         Top             =   480
         Width           =   2895
         _Version        =   131072
         _ExtentX        =   5106
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
         ButtonDesigner  =   "frmTaxDataRepair.frx":19B0
         Begin fpBtnAtlLibCtl.fpBtn fpBtn8 
            Height          =   375
            Left            =   0
            TabIndex        =   56
            Top             =   5040
            Width           =   2895
            _Version        =   131072
            _ExtentX        =   5106
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
            ButtonDesigner  =   "frmTaxDataRepair.frx":1B92
         End
      End
      Begin fpBtnAtlLibCtl.fpBtn fpBtn5 
         Height          =   375
         Left            =   0
         TabIndex        =   57
         Top             =   0
         Width           =   2895
         _Version        =   131072
         _ExtentX        =   5106
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
         ButtonDesigner  =   "frmTaxDataRepair.frx":1D6D
         Begin fpBtnAtlLibCtl.fpBtn fpBtn9 
            Height          =   375
            Left            =   0
            TabIndex        =   58
            Top             =   5040
            Width           =   2895
            _Version        =   131072
            _ExtentX        =   5106
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
            ButtonDesigner  =   "frmTaxDataRepair.frx":1F4D
         End
         Begin fpBtnAtlLibCtl.fpBtn fpBtn10 
            Height          =   375
            Left            =   0
            TabIndex        =   59
            Top             =   480
            Width           =   2895
            _Version        =   131072
            _ExtentX        =   5106
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
            ButtonDesigner  =   "frmTaxDataRepair.frx":2128
            Begin fpBtnAtlLibCtl.fpBtn fpBtn11 
               Height          =   375
               Left            =   0
               TabIndex        =   60
               Top             =   5040
               Width           =   2895
               _Version        =   131072
               _ExtentX        =   5106
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
               ButtonDesigner  =   "frmTaxDataRepair.frx":230A
            End
         End
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMakeBillTypesC 
      Height          =   375
      Left            =   1080
      TabIndex        =   30
      Tag             =   "This makes all bill types equal ""C"" instead of ""R"" or ""P"""
      Top             =   1680
      Width           =   1095
      _Version        =   131072
      _ExtentX        =   1931
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":24E5
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdRepairDatesOnFixedTrans 
      Height          =   375
      Left            =   1080
      TabIndex        =   28
      Tag             =   $"frmTaxDataRepair.frx":26BF
      Top             =   4440
      Width           =   1095
      _Version        =   131072
      _ExtentX        =   1931
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":2756
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess7 
      Height          =   375
      Left            =   3000
      TabIndex        =   20
      Tag             =   $"frmTaxDataRepair.frx":2930
      Top             =   4440
      Width           =   1095
      _Version        =   131072
      _ExtentX        =   1931
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":2A37
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess4 
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Tag             =   "If customer pin numbers are not the same as the customer records then run this procedure. It will match them back up."
      Top             =   4350
      Width           =   1335
      _Version        =   131072
      _ExtentX        =   2355
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":2C14
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess1 
      Height          =   375
      Left            =   8640
      TabIndex        =   3
      Tag             =   $"frmTaxDataRepair.frx":2DF2
      Top             =   4920
      Width           =   1575
      _Version        =   131072
      _ExtentX        =   2778
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":3010
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   630
      Left            =   10080
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1335
      _Version        =   131072
      _ExtentX        =   2355
      _ExtentY        =   1111
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":31EF
   End
   Begin EditLib.fpDateTime fptxtBegDate 
      Height          =   375
      Left            =   9270
      TabIndex        =   1
      Top             =   4020
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   2
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   12648447
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "02/24/2005"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "00000000"
      DateMin         =   "00000000"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "00000000"
      TimeDefault     =   "000000"
      TimeStyle       =   0
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      PopUpType       =   1
      DateCalcY2KSplit=   60
      CaretPosition   =   0
      IncYear         =   1
      IncMonth        =   1
      IncDay          =   1
      IncHour         =   1
      IncMinute       =   1
      IncSecond       =   1
      ButtonColor     =   13684944
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpDateTime fptxtEndDate 
      Height          =   375
      Left            =   9270
      TabIndex        =   2
      Top             =   4500
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   2
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   12648447
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "02/24/2005"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "00000000"
      DateMin         =   "00000000"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "00000000"
      TimeDefault     =   "000000"
      TimeStyle       =   0
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      PopUpType       =   1
      DateCalcY2KSplit=   60
      CaretPosition   =   0
      IncYear         =   1
      IncMonth        =   1
      IncDay          =   1
      IncHour         =   1
      IncMinute       =   1
      IncSecond       =   1
      ButtonColor     =   13684944
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess2 
      Height          =   360
      Left            =   3000
      TabIndex        =   4
      Tag             =   $"frmTaxDataRepair.frx":33CB
      Top             =   1920
      Width           =   1095
      _Version        =   131072
      _ExtentX        =   1931
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":34AC
   End
   Begin EditLib.fpDateTime fptxtFiscalBeg 
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   1920
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   2
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   12648447
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "02/24"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd"
      DateMax         =   "00000000"
      DateMin         =   "00000000"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "00000000"
      TimeDefault     =   "000000"
      TimeStyle       =   0
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      PopUpType       =   1
      DateCalcY2KSplit=   60
      CaretPosition   =   0
      IncYear         =   1
      IncMonth        =   1
      IncDay          =   1
      IncHour         =   1
      IncMinute       =   1
      IncSecond       =   1
      ButtonColor     =   13684944
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpDateTime fptxtFiscalEnd 
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   2640
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   2
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   12648447
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "02/24"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd"
      DateMax         =   "00000000"
      DateMin         =   "00000000"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "00000000"
      TimeDefault     =   "000000"
      TimeStyle       =   0
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      PopUpType       =   1
      DateCalcY2KSplit=   60
      CaretPosition   =   0
      IncYear         =   1
      IncMonth        =   1
      IncDay          =   1
      IncHour         =   1
      IncMinute       =   1
      IncSecond       =   1
      ButtonColor     =   13684944
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess3 
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      ToolTipText     =   $"frmTaxDataRepair.frx":368A
      Top             =   3120
      Width           =   1335
      _Version        =   131072
      _ExtentX        =   2355
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":3778
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess6 
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Tag             =   $"frmTaxDataRepair.frx":3956
      Top             =   3120
      Width           =   1095
      _Version        =   131072
      _ExtentX        =   1931
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":3A4D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess8 
      Height          =   375
      Left            =   9480
      TabIndex        =   22
      Tag             =   $"frmTaxDataRepair.frx":3C2B
      Top             =   1560
      Width           =   1095
      _Version        =   131072
      _ExtentX        =   1931
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":3D32
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess5 
      Height          =   375
      Left            =   1080
      TabIndex        =   24
      Tag             =   $"frmTaxDataRepair.frx":3F0F
      Top             =   3120
      Width           =   1095
      _Version        =   131072
      _ExtentX        =   1931
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":409B
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCnvrtPstdBills 
      Height          =   495
      Left            =   9360
      TabIndex        =   26
      Top             =   2640
      Width           =   1335
      _Version        =   131072
      _ExtentX        =   2355
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":4278
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixRealPropCustPin 
      Height          =   375
      Left            =   7680
      TabIndex        =   32
      Top             =   1680
      Width           =   1095
      _Version        =   131072
      _ExtentX        =   1931
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":445F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixHBInsertFireTax 
      Height          =   375
      Left            =   840
      TabIndex        =   34
      Top             =   7800
      Width           =   2895
      _Version        =   131072
      _ExtentX        =   5106
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":4639
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixHBLateList 
      Height          =   375
      Left            =   840
      TabIndex        =   35
      Top             =   8760
      Width           =   2895
      _Version        =   131072
      _ExtentX        =   5106
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":4826
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixHarrisburg 
      Height          =   375
      Left            =   840
      TabIndex        =   36
      Top             =   5880
      Width           =   2895
      _Version        =   131072
      _ExtentX        =   5106
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":4A11
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixBSL 
      Height          =   375
      Left            =   840
      TabIndex        =   37
      Top             =   5400
      Width           =   2895
      _Version        =   131072
      _ExtentX        =   5106
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":4BF2
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixAdTrans 
      Height          =   375
      Left            =   7680
      TabIndex        =   38
      ToolTipText     =   "Inserts real pin numbers where applicable"
      Top             =   2880
      Width           =   1095
      _Version        =   131072
      _ExtentX        =   1931
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":4DD7
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixAdvChrgAndReal 
      Height          =   375
      Left            =   840
      TabIndex        =   40
      Top             =   6840
      Width           =   2895
      _Version        =   131072
      _ExtentX        =   5106
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":4FB1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixAvery 
      Height          =   375
      Left            =   840
      TabIndex        =   41
      Top             =   6360
      Width           =   2895
      _Version        =   131072
      _ExtentX        =   5106
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":519B
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixIndianTrails 
      Height          =   375
      Left            =   840
      TabIndex        =   42
      Top             =   7320
      Width           =   2895
      _Version        =   131072
      _ExtentX        =   5106
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":537E
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdStripBoiling 
      Height          =   375
      Left            =   840
      TabIndex        =   43
      Top             =   8280
      Width           =   2895
      _Version        =   131072
      _ExtentX        =   5106
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":5562
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixFaison 
      Height          =   375
      Left            =   3840
      TabIndex        =   44
      Top             =   5400
      Width           =   2895
      _Version        =   131072
      _ExtentX        =   5106
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":5744
      Begin fpBtnAtlLibCtl.fpBtn cmd 
         Height          =   375
         Left            =   0
         TabIndex        =   45
         Top             =   5040
         Width           =   2895
         _Version        =   131072
         _ExtentX        =   5106
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
         ButtonDesigner  =   "frmTaxDataRepair.frx":5921
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixAddresses 
      Height          =   375
      Left            =   3840
      TabIndex        =   46
      ToolTipText     =   $"frmTaxDataRepair.frx":5AFC
      Top             =   5880
      Width           =   2895
      _Version        =   131072
      _ExtentX        =   5106
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":5B92
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixBeechMtn 
      Height          =   375
      Left            =   3840
      TabIndex        =   47
      Top             =   6360
      Width           =   2895
      _Version        =   131072
      _ExtentX        =   5106
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":5D78
      Begin fpBtnAtlLibCtl.fpBtn fpBtn2 
         Height          =   375
         Left            =   0
         TabIndex        =   48
         Top             =   5040
         Width           =   2895
         _Version        =   131072
         _ExtentX        =   5106
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
         ButtonDesigner  =   "frmTaxDataRepair.frx":5F59
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixBeachMntWatauga 
      Height          =   375
      Left            =   3840
      TabIndex        =   49
      Top             =   6840
      Width           =   2895
      _Version        =   131072
      _ExtentX        =   5106
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":6134
      Begin fpBtnAtlLibCtl.fpBtn fpBtn3 
         Height          =   375
         Left            =   0
         TabIndex        =   50
         Top             =   5040
         Width           =   2895
         _Version        =   131072
         _ExtentX        =   5106
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
         ButtonDesigner  =   "frmTaxDataRepair.frx":631E
      End
      Begin fpBtnAtlLibCtl.fpBtn fpBtn1 
         Height          =   375
         Left            =   0
         TabIndex        =   51
         Top             =   480
         Width           =   2895
         _Version        =   131072
         _ExtentX        =   5106
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
         ButtonDesigner  =   "frmTaxDataRepair.frx":64F9
         Begin fpBtnAtlLibCtl.fpBtn fpBtn4 
            Height          =   375
            Left            =   0
            TabIndex        =   52
            Top             =   5040
            Width           =   2895
            _Version        =   131072
            _ExtentX        =   5106
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
            ButtonDesigner  =   "frmTaxDataRepair.frx":66DB
         End
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixMagnolia 
      Height          =   375
      Left            =   3840
      TabIndex        =   61
      Top             =   7800
      Width           =   2895
      _Version        =   131072
      _ExtentX        =   5106
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":68B6
      Begin fpBtnAtlLibCtl.fpBtn fpBtn13 
         Height          =   375
         Left            =   0
         TabIndex        =   62
         Top             =   5040
         Width           =   2895
         _Version        =   131072
         _ExtentX        =   5106
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
         ButtonDesigner  =   "frmTaxDataRepair.frx":6A95
      End
      Begin fpBtnAtlLibCtl.fpBtn fpBtn14 
         Height          =   375
         Left            =   0
         TabIndex        =   63
         Top             =   480
         Width           =   2895
         _Version        =   131072
         _ExtentX        =   5106
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
         ButtonDesigner  =   "frmTaxDataRepair.frx":6C70
         Begin fpBtnAtlLibCtl.fpBtn fpBtn15 
            Height          =   375
            Left            =   0
            TabIndex        =   64
            Top             =   5040
            Width           =   2895
            _Version        =   131072
            _ExtentX        =   5106
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
            ButtonDesigner  =   "frmTaxDataRepair.frx":6E52
         End
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdFixMaxton 
         Height          =   375
         Left            =   0
         TabIndex        =   65
         Top             =   0
         Width           =   2895
         _Version        =   131072
         _ExtentX        =   5106
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
         ButtonDesigner  =   "frmTaxDataRepair.frx":702D
         Begin fpBtnAtlLibCtl.fpBtn fpBtn16 
            Height          =   375
            Left            =   0
            TabIndex        =   66
            Top             =   5040
            Width           =   2895
            _Version        =   131072
            _ExtentX        =   5106
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
            ButtonDesigner  =   "frmTaxDataRepair.frx":720A
         End
         Begin fpBtnAtlLibCtl.fpBtn fpBtn17 
            Height          =   375
            Left            =   0
            TabIndex        =   67
            Top             =   480
            Width           =   2895
            _Version        =   131072
            _ExtentX        =   5106
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
            ButtonDesigner  =   "frmTaxDataRepair.frx":73E5
            Begin fpBtnAtlLibCtl.fpBtn fpBtn18 
               Height          =   375
               Left            =   0
               TabIndex        =   68
               Top             =   5040
               Width           =   2895
               _Version        =   131072
               _ExtentX        =   5106
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
               ButtonDesigner  =   "frmTaxDataRepair.frx":75C7
            End
         End
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdUpdateAdd1AndAdd2Short 
      Height          =   360
      Left            =   3840
      TabIndex        =   73
      Top             =   8760
      Width           =   2895
      _Version        =   131072
      _ExtentX        =   5106
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":77A2
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixWhiteLake 
      Height          =   360
      Left            =   6840
      TabIndex        =   74
      Top             =   5880
      Width           =   2895
      _Version        =   131072
      _ExtentX        =   5106
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":798D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixMaggieValley 
      Height          =   375
      Left            =   6840
      TabIndex        =   75
      Top             =   6360
      Width           =   2895
      _Version        =   131072
      _ExtentX        =   5106
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":7B6E
      Begin fpBtnAtlLibCtl.fpBtn fpBtn19 
         Height          =   375
         Left            =   0
         TabIndex        =   76
         Top             =   480
         Width           =   2895
         _Version        =   131072
         _ExtentX        =   5106
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
         ButtonDesigner  =   "frmTaxDataRepair.frx":7D52
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdInsertPrePay 
      Height          =   360
      Left            =   6840
      TabIndex        =   79
      Top             =   7320
      Width           =   2895
      _Version        =   131072
      _ExtentX        =   5106
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":7F2D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClearPayments 
      Height          =   375
      Left            =   6840
      TabIndex        =   80
      Top             =   7800
      Width           =   2895
      _Version        =   131072
      _ExtentX        =   5106
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":8111
      Begin fpBtnAtlLibCtl.fpBtn fpBtn26 
         Height          =   375
         Left            =   0
         TabIndex        =   81
         Top             =   480
         Width           =   2895
         _Version        =   131072
         _ExtentX        =   5106
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
         ButtonDesigner  =   "frmTaxDataRepair.frx":82FB
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExportFiles 
      Height          =   360
      Left            =   6840
      TabIndex        =   83
      ToolTipText     =   "Files exported are: Cust Pin, County #, Personal Pin #, Personal Value, Property Type"
      Top             =   8760
      Width           =   2895
      _Version        =   131072
      _ExtentX        =   5106
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":84D6
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdRemoveTrans 
      Height          =   375
      Left            =   9960
      TabIndex        =   85
      Top             =   7440
      Width           =   1455
      _Version        =   131072
      _ExtentX        =   2566
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
      ButtonDesigner  =   "frmTaxDataRepair.frx":86B5
   End
   Begin VB.Shape Shape16 
      Height          =   975
      Left            =   9840
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H0080FFFF&
      Height          =   2055
      Left            =   7920
      Top             =   3420
      Width           =   3135
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H0080FFFF&
      Height          =   1095
      Left            =   9120
      Top             =   2220
      Width           =   1935
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Insert Real Pin To Adv Trans"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7440
      TabIndex        =   39
      ToolTipText     =   "Use this utility only if the transaction journal release report is not displaying revenues correctly."
      Top             =   2325
      Width           =   1575
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Fix Real Prop Cust Pin"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7440
      TabIndex        =   33
      ToolTipText     =   "Use this utility only if the transaction journal release report is not displaying revenues correctly."
      Top             =   1140
      Width           =   1575
   End
   Begin VB.Shape Shape13 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H0080FFFF&
      Height          =   975
      Left            =   840
      Top             =   1150
      Width           =   1575
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Make All Bill Types Equal ""C"""
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   840
      TabIndex        =   31
      Top             =   1150
      Width           =   1575
   End
   Begin VB.Shape Shape12 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H0080FFFF&
      Height          =   1215
      Left            =   840
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Fix Dates On Repaired Transactions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   840
      TabIndex        =   29
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Convert Posted Fields"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9120
      TabIndex        =   27
      ToolTipText     =   "Use this utility only if the transaction journal release report is not displaying revenues correctly."
      Top             =   2220
      Width           =   1935
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Reconstruct history by eliminating faulty negative balances"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   840
      TabIndex        =   25
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H0080FFFF&
      Height          =   1335
      Left            =   840
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Relink Posting Errors"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9120
      TabIndex        =   23
      Top             =   1140
      Width           =   1935
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H0080FFFF&
      Height          =   855
      Left            =   9120
      Top             =   1140
      Width           =   1935
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H0080FFFF&
      Height          =   1215
      Left            =   2640
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Fix Accumulated BelongTo Trans More Than Bill Itself"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   2640
      TabIndex        =   21
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Bill Trans Only: Make Paid Equal Charged If Paid Is More"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   2640
      TabIndex        =   19
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H0080FFFF&
      Height          =   1215
      Left            =   2640
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H0080FFFF&
      Height          =   1095
      Left            =   4680
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Make Pin Numbers and Acct Numbers The Same"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4680
      TabIndex        =   17
      Top             =   3820
      Width           =   2535
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Fiscal Year Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   4680
      TabIndex        =   16
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Beginning Fiscal Year Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   4560
      TabIndex        =   15
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Make Zero Value Tax Years Correspond to Its Trans Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4680
      TabIndex        =   14
      Top             =   1140
      Width           =   2535
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H0080FFFF&
      Height          =   2535
      Left            =   4680
      Top             =   1140
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Make Zero Value Years Equal To Bill Trans Tax Years"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   2640
      TabIndex        =   13
      Top             =   1140
      Width           =   1935
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H0080FFFF&
      Height          =   1215
      Left            =   2640
      Top             =   1140
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Repair Negative Values and Future1 and Future2 Values"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7920
      TabIndex        =   12
      Top             =   3420
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Begin Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   8070
      TabIndex        =   11
      Top             =   4110
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   7950
      TabIndex        =   10
      Top             =   4590
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1500
      Top             =   285
      Width           =   8655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax DOS Data Repair"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3120
      TabIndex        =   0
      Top             =   450
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1500
      Top             =   180
      Width           =   8655
   End
   Begin VB.Shape Shape14 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H0080FFFF&
      Height          =   975
      Left            =   7440
      Top             =   1140
      Width           =   1575
   End
   Begin VB.Shape Shape15 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H0080FFFF&
      Height          =   975
      Left            =   7440
      Top             =   2340
      Width           =   1575
   End
End
Attribute VB_Name = "frmTaxDataRepair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
Private Sub UpdateTransWithNewPins()
  '2/6/09
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Integer
  Dim arrFrom() As Variant
  arrFrom() = Array("(A)", "`", "LOT 19-3-0 MAP 11", "1 TRACT 6 861/603", "MAP 6 LOT 13", "FR MARLIE KING", _
  "MAP 3 LOT 106R TIMBE", "FR CLEARWATER", "SPLIT PER SURVEY 3", "SPLIT PER SURVEY 6", "SPLIT PER SURVEY 7", _
  "SPLIT PER SURVEY 9", "LOT 95 TURTLE COVE", "MAP 1 PARCEL 66", "4014843", "4014767", "4014844", _
  "LOT 90 TURTLE COVE", "4014760", "LOT 124 TIMBERLODGE", "MAP 1 LOT 54-1")
  Dim arrTo() As Variant
  arrTo() = Array("135200284227", "135210458691", "135205281597", "135217007497", _
  "135218420076", "134220823225", "135217104746", "135217104530", "134200912399", "134200914308", _
  "134200913374", "134200913331", "134207683675", "4014792", "134220815725", "135210369666&461847", _
  "135218415805", "134207686620", "135217010757", "135217105612", "134220913899")
  OpenTaxTransFile TTHandle, NumOfTTRecs
  y = 0
  frmTaxShowPctComp.Label1 = "Fixing White Lake"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  Do While y < 21
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    If QPTrim$(TaxTrans.RealPin) = arrFrom(y) Then
      TaxTrans.RealPin = arrTo(y)
      Put TTHandle, x, TaxTrans
    End If
  Next x
  y = y + 1
  frmTaxShowPctComp.ShowPctComp y, 21
  Loop
  Unload frmTaxShowPctComp
  Close
  
End Sub

Private Sub cmdDelete_Click()
  Dim TaxCust As TaxCustType
  Dim THandle As Integer
  Dim NumOfCust As Long
  Dim x As Long
  
  OpenTaxCustFile THandle, NumOfCust
  For x = 1 To NumOfCust
    Get THandle, x, TaxCust
    TaxCust.Deleted = 0
    Put THandle, x, TaxCust
  Next x
  
  Close
  MsgBox ("Finished.")
End Sub

Private Sub cmdEmpty_Click()

End Sub

Private Sub cmdClearPayments_Click()
  Call RemovePayments
End Sub

Private Sub cmdExit_Click()
'DALE1
  'frmTaxMainMenu.Show
  DoEvents
  Unload Me
End Sub
Private Sub Look4OrphanPayTrans()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Long
  Dim OCnt As Long
  Dim BadCnt As Long
  ReDim BadCNum(1 To 1) As Long
  ReDim BadTrans(1 To 1) As Long
  ReDim OTrans(1 To 1) As Long
  ReDim OTAmt(1 To 1) As Double
  Dim CustRecPay As Long
  Dim CustRecChg As Long
  Dim BT As Long
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
'    If x = 78736 Then Stop
'    If x = 77463 Then Stop
'    TaxTrans.CustomerRec = TaxTrans.CustomerRec
    If TaxTrans.TranType = 22 Then
'      CustRecPay = TaxTrans.CustomerRec
'      BT = TaxTrans.BelongTo
'      Get TTHandle, TaxTrans.BelongTo, TaxTrans
      Debug.Print CStr(TaxTrans.CustomerRec) & "  " & Using$("$##,###.##", TaxTrans.Revenue.PrePaidAmt)
      BadCnt = BadCnt + 1
'      If TaxTrans.CustomerRec <> CustRecPay Then
'        Stop
'      End If
'      OCnt = OCnt + 1
'      ReDim Preserve OTrans(1 To OCnt) As Long
'      OTrans(OCnt) = TaxTrans.BelongTo
'      ReDim Preserve OTAmt(1 To OCnt) As Double
'      OTAmt(OCnt) = TaxTrans.Revenue.Principle1
    End If
  Next x
  
'  For x = 1 To NumOfTTRecs
'    Get TTHandle, x, TaxTrans
'    If TaxTrans.TranType = 1 Then
'      For y = 1 To OCnt
'        If x = OTrans(y) Then
''          If TaxTrans.Revenue.Principle1Pd = 0 Or TaxTrans.Revenue.InterestPd = 0 Then
'          If TaxTrans.Revenue.Principle1Pd <> OTAmt(y) Then
'            BadCnt = BadCnt + 1
'            ReDim Preserve BadTrans(1 To BadCnt) As Long
'            BadTrans(BadCnt) = x
'            ReDim Preserve BadCNum(1 To BadCnt) As Long
'            BadCNum(BadCnt) = TaxTrans.CustomerRec
'          End If
'        End If
'      Next y
'    End If
'  Next x
'
'  Debug.Print "Cust #    Trans #"
'  For x = 1 To BadCnt
'    Debug.Print CStr(BadCNum(x)) & "   " & CStr(BadTrans(x))
'  Next x
    
  Close
  
  MsgBox ("Finished.")
  
End Sub

Private Sub StripOutTrans()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim FileName As String
  Dim ThisFile As Integer
  Dim NextRec As Long
  Dim EmptyTaxTrans As TaxTransactionType
  Dim Cnt As Long
  Dim PriorRec As Long
  
  FileName = "BoilingErrors.txt"
  ThisFile = FreeFile
  Open FileName For Output As ThisFile
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For y = 1 To NumOfTCRecs
    Get TCHandle, y, TaxCust
    If TaxCust.Acct > 10499 Then
      NextRec = TaxCust.LastTrans
      If NextRec > 0 Then
        Get TTHandle, NextRec, TaxTrans
        If TaxTrans.TransDate < Date2Num("11/07/2007") Then
          TaxCust.LastTrans = 0
          Put TCHandle, y, TaxCust
        End If
      End If
      Do While NextRec > 0
        Get TTHandle, NextRec, TaxTrans
        If TaxTrans.TransDate < Date2Num("11/07/2007") Then
          TaxTrans = EmptyTaxTrans
          Put TTHandle, NextRec, TaxTrans
          Cnt = Cnt + 1
          Print #ThisFile, CStr(TaxCust.Acct) + "~" + QPTrim$(TaxCust.CustName)
          If PriorRec > 0 Then
            Get TTHandle, PriorRec, TaxTrans
            TaxTrans.LastTrans = 0
            Put TTHandle, PriorRec, TaxTrans
          End If
        End If
        PriorRec = NextRec
        NextRec = TaxTrans.LastTrans
      Loop
    End If
    frmTaxShowPctComp.ShowPctComp y, NumOfTCRecs
  Next y
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
    
  Close ThisFile
  
  MsgBox ("A total of " + CStr(Cnt) + " transactions were stripped. Look for 'BoilingErrors.txt' to see all transactions stripped.")
  
End Sub

Private Sub cmdExportFiles_Click()
  Call ExportFiles
End Sub

Private Sub cmdFixAddresses_Click()
  Call UpdateAddressFields
End Sub
Private Sub ExportFiles()
  Dim AHandle As Integer
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Long
  Dim TaxCust As TaxCustType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim PersType As String
  Dim PersAmt As Double
  Dim x As Long
  On Error Resume Next
  OpenPersPropFile PHandle, NumOfPRecs
  OpenTaxCustFile THandle, NumOfTRecs
  AHandle = FreeFile
  Open "taxexportfiles.txt" For Output As AHandle
  Print #AHandle, "Citipak Acct Num~String County Acct Num~Numeric County Acct Num~Personal Pin Num~Personal Value~Prop Type"
  For x = 1 To NumOfPRecs
    Get PHandle, x, PersRec
    If PersRec.CustPin > 0 Then
      Get THandle, PersRec.CustPin, TaxCust
      If PersRec.CVALUE > 0 Then
        PersType = "Farm Eq"
        PersAmt = PersRec.CVALUE
      ElseIf PersRec.MCVALUE > 0 Then
        PersType = "Merch Cap"
        PersAmt = PersRec.MCVALUE
      ElseIf PersRec.MHVALUE > 0 Then
        PersType = "Mobile Home"
        PersAmt = PersRec.MHVALUE
      ElseIf PersRec.MTVALUE > 0 Then
        PersType = "Machine Tools"
        PersAmt = PersRec.MTVALUE
      ElseIf PersRec.PersVal > 0 Then
        PersType = "Personal"
        PersAmt = PersRec.PersVal
      End If
    End If
    Print #AHandle, CStr(PersRec.CustPin) & "~" & QPTrim$(TaxCust.CountyAcctString) & "~" & CStr(TaxCust.CountyAcct) & "~" & QPTrim$(PersRec.PropPin) & "~" & Using("$#,###,###.##", PersAmt) & "~" & PersType
  Next x
  Close
  
  MsgBox ("Look for taxexportfiles.txt in the Citipak directory.")
End Sub

Private Sub cmdFixAdTrans_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  Dim RealPin As String
  Dim Cnt As Long
  
  frmTaxShowPctComp.Label1 = "Stripping Transactions"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    If TaxTrans.TranType = 6 Then
      If QPTrim$(TaxTrans.RealPin) = "" Then
        If TaxTrans.BelongTo > 0 Then
          Get TTHandle, TaxTrans.BelongTo, TaxTrans
          RealPin = QPTrim$(TaxTrans.RealPin)
          Get TTHandle, x, TaxTrans
          TaxTrans.RealPin = RealPin
          Put TTHandle, x, TaxTrans
          Cnt = Cnt + 1
        End If
      End If
    End If
  Next x
  Close
  Call Savemsg(900, "A total of " + CStr(Cnt) + " transactions were modified successfully.")

End Sub

Private Sub cmdFixCalabash_Click()
' Call FixCalabash
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  Dim Date1 As Integer
  Dim Date2 As Integer
  Dim Cnt As Integer
  Date1 = Date2Num("12/16/2009")
  Date2 = Date2Num("12/17/2009")
  '12/17/2009
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    If TaxTrans.TranType = 1 Then
      If TaxTrans.TransDate = Date1 Or TaxTrans.TransDate = Date2 Then
        ClearTrans (x)
        Cnt = Cnt + 1
      End If
    End If
  Next x
  
  Close
  MsgBox ("A total of " + CStr(Cnt) + " bills were zeroed out.")

End Sub

Private Sub cmdFixCarShores_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim Cnt As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  OpenRealPropFile RHandle, NumOfRealRecs
  For x = 1 To NumOfRealRecs
    Get RHandle, x, RealRec
    If RealRec.PROPDATE = Date2Num("07/17/2007") Then
      Debug.Print CStr(RealRec.CustPin)
    
'    Get TCHandle, TaxTrans.CustomerRec, TaxCust
'    Get TCHandle, 1965, TaxCust
'    If TaxCust.FirstPropRec = 0 Then
'      TaxTrans.RealPin = ""
'      Put TTHandle, x, TaxTrans
'      Cnt = Cnt + 1
    End If
  Next x
  
  Close
  
  MsgBox ("Done.")
   
End Sub

Private Sub FixHamlet3_18_09()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim TaxCust As TaxCustType
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  
  'fix for 1565
  Get TCHandle, 1565, TaxCust
  TaxCust.LastTrans = 213601
  Put TCHandle, 1565, TaxCust
  Get TTHandle, 213601, TaxTrans
  TaxTrans.LastTrans = 212412
  TaxTrans.CustomerRec = 1565
  TaxTrans.CustPin = 1565
  Put TTHandle, 213601, TaxTrans
  
  'fix for 6073
  Get TCHandle, 6073, TaxCust
  TaxCust.LastTrans = 0
  Put TCHandle, 6073, TaxCust
  
  'fix for 1779
  Get TCHandle, 1779, TaxCust
  TaxCust.LastTrans = 213732
  Put TCHandle, 1779, TaxCust
  Get TTHandle, 213732, TaxTrans
  TaxTrans.LastTrans = 212860
  TaxTrans.CustomerRec = 1779
  TaxTrans.CustPin = 1779
  Put TTHandle, 213732, TaxTrans
  
  'fix for 6114
  Get TCHandle, 6114, TaxCust
  TaxCust.LastTrans = 0
  Put TCHandle, 6114, TaxCust

  'fix for 1934
  Get TCHandle, 1934, TaxCust
  TaxCust.LastTrans = 213669
  Put TCHandle, 1934, TaxCust
  Get TTHandle, 213669, TaxTrans
  TaxTrans.LastTrans = 213200
  TaxTrans.CustomerRec = 1934
  TaxTrans.CustPin = 1934
  Put TTHandle, 213669, TaxTrans
  
  'fix for 6092
  Get TCHandle, 6092, TaxCust
  TaxCust.LastTrans = 0
  Put TCHandle, 6092, TaxCust

  'fix for 2017
  Get TCHandle, 2017, TaxCust
  TaxCust.LastTrans = 213706
  Put TCHandle, 2017, TaxCust
  Get TTHandle, 213706, TaxTrans
  TaxTrans.LastTrans = 212752
  TaxTrans.CustomerRec = 2017
  TaxTrans.CustPin = 2017
  Put TTHandle, 213706, TaxTrans
  
  'fix for 6104
  Get TCHandle, 6104, TaxCust
  TaxCust.LastTrans = 0
  Put TCHandle, 6104, TaxCust

  'fix for 2126
  Get TCHandle, 2126, TaxCust
  TaxCust.LastTrans = 213723
  Put TCHandle, 2126, TaxCust
  Get TTHandle, 213723, TaxTrans
  TaxTrans.LastTrans = 212834
  TaxTrans.CustomerRec = 2126
  TaxTrans.CustPin = 2126
  Put TTHandle, 213723, TaxTrans
  
  'fix for 6111
  Get TCHandle, 6111, TaxCust
  TaxCust.LastTrans = 0
  Put TCHandle, 6111, TaxCust

  'fix for 2852
  Get TCHandle, 2852, TaxCust
  TaxCust.LastTrans = 213392
  Put TCHandle, 2852, TaxCust
  Get TTHandle, 213392, TaxTrans
  TaxTrans.LastTrans = 211663
  TaxTrans.CustomerRec = 2852
  TaxTrans.CustPin = 2852
  Put TTHandle, 213392, TaxTrans
  
  'fix for 6002
  Get TCHandle, 6002, TaxCust
  TaxCust.LastTrans = 0
  Put TCHandle, 6002, TaxCust

  'fix for 3228
  Get TCHandle, 3228, TaxCust
  TaxCust.LastTrans = 213740
  Put TCHandle, 3228, TaxCust
  Get TTHandle, 213740, TaxTrans
  TaxTrans.LastTrans = 212876
  TaxTrans.CustomerRec = 3228
  TaxTrans.CustPin = 3228
  Put TTHandle, 213740, TaxTrans
  
  'fix for 6115
  Get TCHandle, 6115, TaxCust
  TaxCust.LastTrans = 0
  Put TCHandle, 6115, TaxCust

  'fix for 5270
  Get TCHandle, 5270, TaxCust
  TaxCust.LastTrans = 213589
  Put TCHandle, 5270, TaxCust
  Get TTHandle, 213589, TaxTrans
  TaxTrans.LastTrans = 212350
  TaxTrans.CustomerRec = 5270
  TaxTrans.CustPin = 5270
  Put TTHandle, 213589, TaxTrans
  
  'fix for 6071
  Get TCHandle, 6071, TaxCust
  TaxCust.LastTrans = 0
  Put TCHandle, 6071, TaxCust

  'fix for 5477
  Get TCHandle, 5477, TaxCust
  TaxCust.LastTrans = 213698
  Put TCHandle, 5477, TaxCust
  Get TTHandle, 213698, TaxTrans
  TaxTrans.LastTrans = 212740
  TaxTrans.CustomerRec = 5477
  TaxTrans.CustPin = 5477
  Put TTHandle, 213698, TaxTrans
  
  'fix for 6101
  Get TCHandle, 6101, TaxCust
  TaxCust.LastTrans = 0
  Put TCHandle, 6101, TaxCust
  
  'fix for 5781
  Get TCHandle, 5781, TaxCust
  TaxCust.LastTrans = 213711
  Put TCHandle, 5781, TaxCust
  Get TTHandle, 213711, TaxTrans
  TaxTrans.LastTrans = 212782
  TaxTrans.CustomerRec = 5781
  TaxTrans.CustPin = 5781
  Put TTHandle, 213711, TaxTrans
  
  'fix for 6105
  Get TCHandle, 6105, TaxCust
  TaxCust.LastTrans = 0
  Put TCHandle, 6105, TaxCust
  
  'fix for 5867
  Get TCHandle, 5867, TaxCust
  TaxCust.LastTrans = 213401
  Put TCHandle, 5867, TaxCust
  Get TTHandle, 213401, TaxTrans
  TaxTrans.LastTrans = 212740
  TaxTrans.CustomerRec = 5867
  TaxTrans.CustPin = 5867
  Put TTHandle, 213401, TaxTrans
  
  'fix for 6004
  Get TCHandle, 6004, TaxCust
  TaxCust.LastTrans = 0
  Put TCHandle, 6004, TaxCust

  'fix for 5961
  Get TCHandle, 5961, TaxCust
  TaxCust.LastTrans = 213705
  Put TCHandle, 5961, TaxCust
  Get TTHandle, 213705, TaxTrans
  TaxTrans.LastTrans = 212749
  TaxTrans.CustomerRec = 5961
  TaxTrans.CustPin = 5961
  Put TTHandle, 213705, TaxTrans
  
  'fix for 6103
  Get TCHandle, 6103, TaxCust
  TaxCust.LastTrans = 0
  Put TCHandle, 6103, TaxCust

  Close
  MsgBox ("Finished.")
  
End Sub

Private Sub cmdFixAvery_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Integer
  Dim BelongTo As Long
  Dim Collection As Double
  Dim CollectionPd As Double
  Dim Interest As Double
  Dim InterestPd As Double
  Dim LateList As Double
  Dim LateListPd As Double
  Dim Penalty As Double
  Dim PenaltyPd As Double
  Dim PrePaidUsed As Double
  Dim Principle1 As Double
  Dim Principle1Pd As Double
  Dim Principle2 As Double
  Dim Principle2Pd As Double
  Dim Principle3 As Double
  Dim Principle3Pd As Double
  Dim Principle4 As Double
  Dim Principle4Pd As Double
  Dim Principle5 As Double
  Dim Principle5Pd As Double
  Dim RevOpt1 As Double
  Dim RevOpt1Pd As Double
  Dim RevOpt2 As Double
  Dim RevOpt2Pd As Double
  Dim RevOpt3 As Double
  Dim RevOpt3Pd As Double
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For x = 6100 To NumOfTTRecs
   Get TTHandle, x, TaxTrans
   If x >= 6100 And x <= 6109 Then GoSub FixIt
   If x >= 6369 And x <= 6378 Then GoSub FixIt
  Next x
  
  Close
  MsgBox ("Completed.")
  Exit Sub
  
FixIt:
    If TaxTrans.BelongTo = 0 Then
      GoTo GoHere
    End If
    Collection = TaxTrans.Revenue.Collection
    CollectionPd = TaxTrans.Revenue.CollectionPd
    Interest = TaxTrans.Revenue.Interest
    InterestPd = TaxTrans.Revenue.InterestPd
    LateList = TaxTrans.Revenue.LateList
    LateListPd = TaxTrans.Revenue.LateListPd
    Penalty = TaxTrans.Revenue.Penalty
    PenaltyPd = TaxTrans.Revenue.PenaltyPd
    PrePaidUsed = TaxTrans.Revenue.PrePaidUsed
    Principle1 = TaxTrans.Revenue.Principle1
    Principle1Pd = TaxTrans.Revenue.Principle1Pd
    Principle2 = TaxTrans.Revenue.Principle2
    Principle2Pd = TaxTrans.Revenue.Principle2Pd
    Principle3 = TaxTrans.Revenue.Principle3
    Principle3Pd = TaxTrans.Revenue.Principle3Pd
    Principle4 = TaxTrans.Revenue.Principle4
    Principle4Pd = TaxTrans.Revenue.Principle4Pd
    Principle5 = TaxTrans.Revenue.Principle5
    Principle5Pd = TaxTrans.Revenue.Principle5Pd
    RevOpt1 = TaxTrans.Revenue.RevOpt1
    RevOpt1Pd = TaxTrans.Revenue.RevOpt1Pd
    RevOpt2 = TaxTrans.Revenue.RevOpt2
    RevOpt2Pd = TaxTrans.Revenue.RevOpt2Pd
    RevOpt3 = TaxTrans.Revenue.RevOpt3
    RevOpt3Pd = TaxTrans.Revenue.RevOpt3Pd
    
    BelongTo = TaxTrans.BelongTo
    
    Get TTHandle, BelongTo, TaxTrans
    TaxTrans.Revenue.Collection = TaxTrans.Revenue.Collection - Collection
    TaxTrans.Revenue.CollectionPd = TaxTrans.Revenue.CollectionPd - CollectionPd
    TaxTrans.Revenue.Interest = TaxTrans.Revenue.Interest - Interest
    TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd - InterestPd
    TaxTrans.Revenue.LateList = TaxTrans.Revenue.LateList - LateList
    TaxTrans.Revenue.LateListPd = TaxTrans.Revenue.LateListPd - LateListPd
    TaxTrans.Revenue.Penalty = TaxTrans.Revenue.Penalty - Penalty
    TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - PenaltyPd
    TaxTrans.Revenue.Principle1 = TaxTrans.Revenue.Principle1 - Principle1
    TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - Principle1Pd
    TaxTrans.Revenue.Principle2 = TaxTrans.Revenue.Principle2 - Principle2
    TaxTrans.Revenue.Principle2Pd = TaxTrans.Revenue.Principle2Pd - Principle2Pd
    TaxTrans.Revenue.Principle3 = TaxTrans.Revenue.Principle3 - Principle3
    TaxTrans.Revenue.Principle3Pd = TaxTrans.Revenue.Principle3Pd - Principle3Pd
    TaxTrans.Revenue.Principle4 = TaxTrans.Revenue.Principle4 - Principle4
    TaxTrans.Revenue.Principle4Pd = TaxTrans.Revenue.Principle4Pd - Principle4Pd
    TaxTrans.Revenue.Principle5 = TaxTrans.Revenue.Principle5 - Principle5
    TaxTrans.Revenue.Principle5Pd = TaxTrans.Revenue.Principle5Pd - Principle5Pd
    TaxTrans.Revenue.RevOpt1 = TaxTrans.Revenue.RevOpt1 - RevOpt1
    TaxTrans.Revenue.RevOpt1Pd = TaxTrans.Revenue.RevOpt1Pd - RevOpt1Pd
    TaxTrans.Revenue.RevOpt2 = TaxTrans.Revenue.RevOpt2 - RevOpt2
    TaxTrans.Revenue.RevOpt2Pd = TaxTrans.Revenue.RevOpt2Pd - RevOpt2Pd
    TaxTrans.Revenue.RevOpt3 = TaxTrans.Revenue.RevOpt3 - RevOpt3
    TaxTrans.Revenue.RevOpt3Pd = TaxTrans.Revenue.RevOpt3Pd - RevOpt3Pd
    Put TTHandle, BelongTo, TaxTrans
    Get TTHandle, x, TaxTrans
GoHere:
    
    TaxTrans.Amount = 0
    TaxTrans.Revenue.Collection = 0
    TaxTrans.Revenue.CollectionPd = 0
    TaxTrans.Revenue.Interest = 0
    TaxTrans.Revenue.InterestPd = 0
    TaxTrans.Revenue.LateList = 0
    TaxTrans.Revenue.LateListPd = 0
    TaxTrans.Revenue.Penalty = 0
    TaxTrans.Revenue.PenaltyPd = 0
    TaxTrans.Revenue.PrePaidUsed = 0
    TaxTrans.Revenue.Principle1 = 0
    TaxTrans.Revenue.Principle1Pd = 0
    TaxTrans.Revenue.Principle2 = 0
    TaxTrans.Revenue.Principle2Pd = 0
    TaxTrans.Revenue.Principle3 = 0
    TaxTrans.Revenue.Principle3Pd = 0
    TaxTrans.Revenue.Principle4 = 0
    TaxTrans.Revenue.Principle4Pd = 0
    TaxTrans.Revenue.Principle5 = 0
    TaxTrans.Revenue.Principle5Pd = 0
    TaxTrans.Revenue.RevOpt1 = 0
    TaxTrans.Revenue.RevOpt1Pd = 0
    TaxTrans.Revenue.RevOpt2 = 0
    TaxTrans.Revenue.RevOpt2Pd = 0
    TaxTrans.Revenue.RevOpt3 = 0
    TaxTrans.Revenue.RevOpt3Pd = 0
    Put TTHandle, x, TaxTrans
    Return

End Sub

Private Sub cmdFixClyde_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Integer
  Dim NextRec As Long
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
    Get TCHandle, 857, TaxCust
    TaxCust.Acct = 857
    Put TCHandle, 857, TaxCust
    NextRec = TaxCust.LastTrans
    Do While NextRec > 0
      Get TTHandle, NextRec, TaxTrans
      TaxTrans.CustomerRec = 857
      TaxTrans.CustPin = 857
      Put TTHandle, NextRec, TaxTrans
      NextRec = TaxTrans.LastTrans
    Loop
   Close
   MsgBox ("Done.")

End Sub

Private Sub cmdFixHamlet_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Integer
  Dim IntTot As Double
  Dim NextRec As Long
  
'  Call FixHamlet3_18_09
'  Exit Sub
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  'fix for 5607 from #6215
  Get TCHandle, 6215, TaxCust '5/10/2010
  TaxCust.LastTrans = 0
  Put TCHandle, 6215, TaxCust
  
  Get TTHandle, 240030, TaxTrans '5/10/2010
  TaxTrans.LastTrans = 238971
  Put TTHandle, 240030, TaxTrans
  
  Get TTHandle, 238971, TaxTrans
  TaxTrans.LastTrans = 237549
  TaxTrans.CustomerRec = 5607
  TaxTrans.CustPin = 5607
  Put TTHandle, 238971, TaxTrans
  
  'fix for 361 from #6272
  Get TCHandle, 6272, TaxCust '5/10/2010
  TaxCust.LastTrans = 0
  Put TCHandle, 6272, TaxCust
  
  Get TTHandle, 239356, TaxTrans '5/10/2010
  TaxTrans.LastTrans = 239251#
  Put TTHandle, 239356, TaxTrans
  
  Get TTHandle, 239251, TaxTrans
  TaxTrans.LastTrans = 238352
  TaxTrans.CustomerRec = 361
  TaxTrans.CustPin = 361
  Put TTHandle, 239251, TaxTrans
  
  'fix for 132 from #6186
  Get TCHandle, 6186, TaxCust '5/10/2010
  TaxCust.LastTrans = 0
  Put TCHandle, 6186, TaxCust
  
  Get TTHandle, 239515, TaxTrans '5/10/2010
  TaxTrans.LastTrans = 238815
  Put TTHandle, 239515, TaxTrans
  
  Get TTHandle, 238815, TaxTrans
  TaxTrans.LastTrans = 236964
  TaxTrans.CustomerRec = 132
  TaxTrans.CustPin = 132
  Put TTHandle, 238815, TaxTrans
  
  'fix for 5185 from #6245
  Get TCHandle, 6245, TaxCust '5/10/2010
  TaxCust.LastTrans = 0
  Put TCHandle, 6245, TaxCust
  
  Get TTHandle, 240574, TaxTrans '5/10/2010
  TaxTrans.LastTrans = 239144
  Put TTHandle, 240574, TaxTrans
  
  Get TTHandle, 239144, TaxTrans
  TaxTrans.LastTrans = 238171
  TaxTrans.CustomerRec = 5184
  TaxTrans.CustPin = 5184
  Put TTHandle, 239144, TaxTrans
  
  'fix for 741 from #6211
  Get TCHandle, 6211, TaxCust '5/10/2010
  TaxCust.LastTrans = 0
  Put TCHandle, 6211, TaxCust
  
  Get TTHandle, 239945, TaxTrans '5/10/2010
  TaxTrans.LastTrans = 238950
  Put TTHandle, 239945, TaxTrans
  
  Get TTHandle, 238950, TaxTrans
  TaxTrans.LastTrans = 237452
  TaxTrans.CustomerRec = 741
  TaxTrans.CustPin = 741
  Put TTHandle, 238950, TaxTrans
 
  
  Get TCHandle, 6217, TaxCust '4/30/2010
  TaxCust.LastTrans = 0
  Put TCHandle, 6217, TaxCust
  Get TTHandle, 240043, TaxTrans
  TaxTrans.LastTrans = 238977
  Put TTHandle, 240043, TaxTrans
  Get TTHandle, 238977, TaxTrans
  TaxTrans.LastTrans = 237564
  TaxTrans.CustomerRec = 895
  TaxTrans.CustPin = 895
  Put TTHandle, 238977, TaxTrans
  
  
'  Call FixHamlet3_18_09
'  Exit Sub
 
'  OpenTaxTransFile TTHandle, NumOfTTRecs
'  'fix for 3398 6/9/09
'  Get TTHandle, 84405, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Interest = 0
'  TaxTrans.Revenue.InterestPd = 0
'  Put TTHandle, 84405, TaxTrans
'
  'fix for 5867 5/6/09
'  Get TTHandle, 215533, TaxTrans
'  TaxTrans.BelongTo = 197024
'  TaxTrans.Revenue.Principle1Pd = 10.4
'  TaxTrans.Revenue.PrePaidAmt = 2.36
'  TaxTrans.Revenue.PrePaidBal = 2.36
'  TaxTrans.Description = "1016"
'  TaxTrans.CustomerRec = 5867
'  TaxTrans.CustPin = 5867
'  TaxTrans.RealPin = "749109262014"
'  TaxTrans.TranType = 21
'  Put TTHandle, 215533, TaxTrans
'
'  Get TTHandle, 215525, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 293.47
'  TaxTrans.Revenue.InterestPd = 2#
'  TaxTrans.BelongTo = 197024
'  TaxTrans.Description = "1016"
'  TaxTrans.RealPin = "749109262014"
'  TaxTrans.CustomerRec = 5867
'  TaxTrans.CustPin = 5867
'  Put TTHandle, 215525, TaxTrans
'
'  Get TTHandle, 215364, TaxTrans
'  TaxTrans.BelongTo = 197024
'  TaxTrans.Description = "1016"
'  TaxTrans.RealPin = "749109262014"
'  TaxTrans.CustomerRec = 5867
'  TaxTrans.CustPin = 5867
'  Put TTHandle, 215364, TaxTrans
'
'  Get TTHandle, 197024, TaxTrans
'  TaxTrans.Revenue.CollectionPd = 6
'  TaxTrans.Revenue.Principle1Pd = 293.4
'  TaxTrans.Revenue.Interest = 12.47
'  TaxTrans.Revenue.InterestPd = 12.47
'  Put TTHandle, 197024, TaxTrans
'
'  'fix for 5477 5/6/09
'  Get TTHandle, 198960, TaxTrans
'  TaxTrans.Revenue.CollectionPd = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Interest = 12.47
'  TaxTrans.Revenue.InterestPd = 0
'  Put TTHandle, 198960, TaxTrans
  
  'fix for #734 on 4/8/09
'  Get TTHandle, 167545, TaxTrans
'  TaxTrans.Revenue.CollectionPd = 6
'  TaxTrans.Revenue.PrePaidAmt = 6
'  TaxTrans.TranType = 21
'  Put TTHandle, 167545, TaxTrans
'
'   'fix for #4735 on 4/8/09
'  Get TTHandle, 167465, TaxTrans
'  TaxTrans.Revenue.CollectionPd = 6
'  TaxTrans.Revenue.PrePaidAmt = 6
'  TaxTrans.TranType = 21
'  Put TTHandle, 167465, TaxTrans
 
  'fix for 1731
'  Get TTHandle, 170279, TaxTrans
'  TaxTrans.Revenue.PrePaidUsed = 0
'  Put TTHandle, 170279, TaxTrans
'
'  Get TTHandle, 149846, TaxTrans
'  TaxTrans.Revenue.Collection = 6
'  TaxTrans.Revenue.CollectionPd = 6
'  Put TTHandle, 149846, TaxTrans
'

  'fix for 1340
'  Get TTHandle, 1590, TaxTrans
'  TaxTrans.Revenue.InterestPd = 4.28
'  Put TTHandle, 1590, TaxTrans
'
'  Get TTHandle, 7564, TaxTrans
'  TaxTrans.Revenue.InterestPd = 36.11
'  Put TTHandle, 7564, TaxTrans
'
'  Get TTHandle, 7565, TaxTrans
'  TaxTrans.Revenue.InterestPd = 44.01
'  Put TTHandle, 7565, TaxTrans
'
'  Get TTHandle, 9317, TaxTrans
'  TaxTrans.BelongTo = TaxTrans.BelongTo
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.InterestPd = 0
'  Put TTHandle, 9317, TaxTrans
'
'  Get TTHandle, 9318, TaxTrans
'  TaxTrans.BelongTo = TaxTrans.BelongTo
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.InterestPd = 0
'  Put TTHandle, 9318, TaxTrans
'
'  Get TTHandle, 9319, TaxTrans
'  TaxTrans.BelongTo = TaxTrans.BelongTo
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.InterestPd = 0
'  Put TTHandle, 9319, TaxTrans
'
'  Get TTHandle, 9320, TaxTrans
'  TaxTrans.BelongTo = TaxTrans.BelongTo
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.InterestPd = 0
'  Put TTHandle, 9320, TaxTrans
'
'  Get TTHandle, 8624, TaxTrans
'  TaxTrans.BelongTo = TaxTrans.BelongTo
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Interest = 0
'  Put TTHandle, 8624, TaxTrans
'
'  Get TTHandle, 170321, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.InterestPd = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  Put TTHandle, 170321, TaxTrans
  Close
  MsgBox ("Finished.")
  
End Sub

Private Sub cmdFixClarkton_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  'fix for 661 6/12/2009
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Principle1Pd = 140.79
  TaxTrans.Revenue.PrePaidUsed = 140.79
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.TranType = 9
  TaxTrans.Description = "Credit Applied To Bill #52932"
  TaxTrans.TransDate = Date2Num("08/13/2008")
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.CollectionPd = 0
  TaxTrans.Revenue.LateListPd = 0
  TaxTrans.Revenue.RevOpt1Pd = 0
  TaxTrans.Revenue.RevOpt2Pd = 0
  TaxTrans.Revenue.RevOpt3Pd = 0
  TaxTrans.CustPin = 661
  TaxTrans.DiscXDate = Date2Num("12/31/1979")
  TaxTrans.RealPin = "25-05048"
  TaxTrans.PersPin = " "
  TaxTrans.Posted2GL = "N"
  TaxTrans.TaxYear = 2008
  TaxTrans.DiscAmt = 0
  TaxTrans.OperNum = 0
  TaxTrans.Amount = 0
  TaxTrans.CustomerRec = 661
  TaxTrans.LastTrans = 30502
  TaxTrans.BelongTo = 30502
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  Put TTHandle, NumOfTTRecs + 1, TaxTrans
  
  Get TTHandle, 32981, TaxTrans
  TaxTrans.LastTrans = NumOfTTRecs + 1
  Put TTHandle, 32981, TaxTrans
  
  Get TTHandle, 30502, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 140.79
  Put TTHandle, 30502, TaxTrans
  
  'fix for 334
'  Get TTHandle, 26031, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  Put TTHandle, 26031, TaxTrans
'
'  Get TTHandle, 26030, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 0
'  Put TTHandle, 26030, TaxTrans
'
'  Get TTHandle, 14038, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  Put TTHandle, 14038, TaxTrans

  'more on 510 2/28/08
'  Get TTHandle, 25125, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.InterestPd = 0
'  TaxTrans.Revenue.CollectionPd = 0
'  TaxTrans.TranType = 2
'  TaxTrans.Description = "Error Fix: SS"
'  Put TTHandle, 25125, TaxTrans
'
'  Get TTHandle, 26131, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.Interest = 0
'  TaxTrans.Revenue.Collection = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.InterestPd = 0
'  TaxTrans.Revenue.CollectionPd = 0
'  Put TTHandle, 26131, TaxTrans
  
'  'fix for #510
'  Get TTHandle, 25516, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.InterestPd = 0
'  Put TTHandle, 25516, TaxTrans
'
'  Get TTHandle, 25515, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.InterestPd = 0
'  Put TTHandle, 25515, TaxTrans
'
'  Get TTHandle, 25539, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.InterestPd = 0
'  TaxTrans.Revenue.Interest = 0
'  Put TTHandle, 25539, TaxTrans
  
  
  Close
  MsgBox ("Finished.")
End Sub

Private Sub cmdFixAdvChrgAndReal_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  Dim RealRec As PropertyRecType
  Dim NumOfRRecs As Long
  Dim RHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTaxCusts As Long
  Dim RCnt As Integer
  Dim RealIdx As Long
  Dim FileName$
  Dim ThisFile As Integer
  Dim ThisCnt As Integer
  Dim ChangeCnt As Integer
  
  FileName = "advchrgaddedtoreal.txt"
  ThisFile = FreeFile
  Open FileName For Output As ThisFile
  
  OpenRealPropFile RHandle, NumOfRRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxCustFile TCHandle, NumOfTaxCusts
  RCnt = 0
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    ThisCnt = 0
    If TaxTrans.TranType = 6 Then
      If TaxTrans.Amount = 0 Then GoTo Skip
      If QPTrim$(TaxTrans.RealPin) = "" Then
        Get TCHandle, TaxTrans.CustomerRec, TaxCust
        RealIdx = TaxCust.FirstPropRec
        Do While RealIdx > 0
          Get RHandle, RealIdx, RealRec
          RCnt = RCnt + 1
          ThisCnt = ThisCnt + 1
          If ThisCnt = 1 Then
            Print #ThisFile, "*" + QPTrim$(TaxCust.CustName) + "~" + CStr(TaxCust.Acct) + "~" + QPTrim$(RealRec.RealPin) + "~" + FormatCurrency(TaxTrans.Amount, 2)
          Else
            Print #ThisFile, QPTrim$(TaxCust.CustName) + "~" + CStr(TaxCust.Acct) + "~" + QPTrim$(RealRec.RealPin) + "~" + FormatCurrency(TaxTrans.Amount, 2)
          End If
          RealIdx = RealRec.NextRec
        Loop
        If ThisCnt = 1 Then
          TaxTrans.RealPin = QPTrim$(RealRec.RealPin)
          ChangeCnt = ChangeCnt + 1
          Put TTHandle, x, TaxTrans
        End If
       End If
     End If
Skip:
  Next x
  Print #ThisFile, "A total of " + CStr(ChangeCnt) + " changes were made. A total of " + CStr(RCnt) + " errors were found. "
  
  Close
  MsgBox ("Process completed with " + CStr(RCnt) + " errors found and " + CStr(ChangeCnt) + " errors fixed.")
End Sub
Private Sub ClearTrans(ByVal RecNum As Long)
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  OpenTaxTransFile THandle, NumOfTRecs
  '#72 3/24/09
  Get THandle, RecNum, TaxTrans
    If TaxTrans.TranType = 2 Then
      Call ClearPayOnBill(TaxTrans.BelongTo, TaxTrans.Amount, TaxTrans.Revenue.Principle1Pd, TaxTrans.Revenue.Principle2Pd, TaxTrans.Revenue.Principle3Pd, TaxTrans.Revenue.Principle4Pd, TaxTrans.Revenue.Principle5Pd, TaxTrans.Revenue.InterestPd, TaxTrans.Revenue.CollectionPd)
    End If
    TaxTrans.Amount = 0
    TaxTrans.Revenue.Collection = 0
    TaxTrans.Revenue.CollectionPd = 0
    TaxTrans.Revenue.Interest = 0
    TaxTrans.Revenue.InterestPd = 0
    TaxTrans.Revenue.LateList = 0
    TaxTrans.Revenue.LateListPd = 0
    TaxTrans.Revenue.Penalty = 0
    TaxTrans.Revenue.PenaltyPd = 0
    TaxTrans.Revenue.PrePaidUsed = 0
    TaxTrans.Revenue.Principle1 = 0
    TaxTrans.Revenue.Principle1Pd = 0
    TaxTrans.Revenue.Principle2 = 0
    TaxTrans.Revenue.Principle2Pd = 0
    TaxTrans.Revenue.Principle3 = 0
    TaxTrans.Revenue.Principle3Pd = 0
    TaxTrans.Revenue.Principle4 = 0
    TaxTrans.Revenue.Principle4Pd = 0
    TaxTrans.Revenue.Principle5 = 0
    TaxTrans.Revenue.Principle5Pd = 0
    TaxTrans.Revenue.RevOpt1 = 0
    TaxTrans.Revenue.RevOpt1Pd = 0
    TaxTrans.Revenue.RevOpt2 = 0
    TaxTrans.Revenue.RevOpt2Pd = 0
    TaxTrans.Revenue.RevOpt3 = 0
    TaxTrans.Revenue.RevOpt3Pd = 0
    TaxTrans.Revenue.PrePaidBal = TaxTrans.Revenue.PrePaidBal - TaxTrans.Revenue.PrePaidAmt
    TaxTrans.Revenue.PrePaidAmt = 0
    TaxTrans.Revenue.PrePaidUsed = 0
    Put THandle, RecNum, TaxTrans
  Close THandle
End Sub
Private Sub ClearPayOnBill(ByVal Rec As Long, ByVal Amt As Double, ByVal Amt1 As Double, ByVal Amt2 As Double, ByVal Amt3 As Double, ByVal Amt4 As Double, ByVal Amt5 As Double, ByVal Amt6 As Double, ByVal Amt7 As Double)
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  OpenTaxTransFile THandle, NumOfTRecs
  Get THandle, Rec, TaxTrans
   ' TaxTrans.Amount = TaxTrans.Amount - Amt
    TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - Amt1
    TaxTrans.Revenue.Principle2Pd = TaxTrans.Revenue.Principle2Pd - Amt2
    TaxTrans.Revenue.Principle3Pd = TaxTrans.Revenue.Principle3Pd - Amt3
    TaxTrans.Revenue.Principle4Pd = TaxTrans.Revenue.Principle4Pd - Amt4
    TaxTrans.Revenue.Principle5Pd = TaxTrans.Revenue.Principle5Pd - Amt5
    TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd - Amt6
    TaxTrans.Revenue.CollectionPd = TaxTrans.Revenue.CollectionPd - Amt7
    Put THandle, Rec, TaxTrans
  Close THandle
End Sub

Private Sub cmdFixBeachMntWatauga_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTaxCusts As Long
  Dim NextRec As Long
  Dim ArrString As String
  Dim Cnt As Integer
  
'  ArrString = "23784, 23785, 23786, 23787, 23788, 23789, 23790, 23791, 23792, 23793, "
'  ArrString = ArrString + "23794, 23795, 23796, 23797, 23798, 23799, 23800, 23801, "
'  ArrString = ArrString + "23802, 23803, 23804, 23805, 23806, 23807, 23808, 23809, "
'  ArrString = ArrString + "23810, 23811, 23812, 23813, 23814, 23815, 23816, 23817, "
'  ArrString = ArrString + "23818, 23819, 23820, 23821, 23822, 23823, 23824, 23825, "
'  ArrString = ArrString + "23826, 23827, 23828, 23829, 23830, 23831, 23832, 23833, "
'  ArrString = ArrString + "23834, 23835, 23836, 23837, 23838, 23839, 23840, 23841, "
'  ArrString = ArrString + "23842, 23843, 23844, 23845, 23846, 23847, 23848, "
'  ArrString = ArrString + "23850, 23851, 23853, 23854, 23855, 23856, 23857, 23858, "
'  ArrString = ArrString + "23859, 23860, 23861, 23862, 23863, 23864, 23865, 23866, "
'  ArrString = ArrString + "23867, 24985, 24986, 24987, 24988, 24989, 24990, 24991, "
'  ArrString = ArrString + "24992, 24993, 24994, 24995, 24997, 24998, 24999, 25000, "
'  ArrString = ArrString + "25001, 25002, 25003, 25004, 25005, 25006, 25007, 25008, "
'  ArrString = ArrString + "25009, 25010, 25011, 25012, 25013, 24014, 25015, 25016, "
'  ArrString = ArrString + "25017, 25018, 25019, 25020, 25021, 25022, 25023, 25024, "
'  ArrString = ArrString + "32959, 32960, 32962, 32963, 32966, 32986, 32989, 32990, "
'  ArrString = ArrString + "33017, 33069, 23664, "
  ArrString = "44492, 45337, " '1/20/2010

  Dim CArr() As Long
  Call BuildArray(ArrString, CArr(), Cnt)
  
'  OpenTaxTransFile TTHandle, NumOfTTRec
'  OpenTaxCustFile TCHandle, NumOfTaxCusts
  'many fixes on 11/19/2009
  For x = 1 To Cnt
    ClearTrans (CArr(x))
  Next x
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Get TTHandle, 42137, TaxTrans '1/20/2010
  TaxTrans.Revenue.Principle1Pd = 2215.04
  Put TTHandle, 42137, TaxTrans
  
  Close
  MsgBox "Completed successfully."

End Sub
Private Sub BuildArray(ByVal ArrString As String, ByRef CArr() As Long, ByRef Cnt As Integer)
  Dim x As Integer, y As Integer
  Dim ch As String
  Dim NewWord As String
  Dim Lgt As Integer
  ReDim arr(1 To 1) As Long
  Lgt = Len(ArrString)
  
  For x = 1 To Lgt
    ch = Mid(ArrString, x, 1)
    If QPTrim$(ch) = "" Then GoTo NextOne
    If ch <> "," Then
      NewWord = NewWord + ch
    Else
      Cnt = Cnt + 1
      ReDim Preserve arr(1 To Cnt) As Long
      arr(Cnt) = CLng(NewWord)
      NewWord = ""
    End If
NextOne:
  Next x
  CArr() = arr()
End Sub

Private Sub cmdFixBeechMtn_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTaxCusts As Long
  Dim NextRec As Long
  Dim Cnt As Integer
  Dim ArrString As String
  'many fixes on 11/19/2009
'  ArrString = "6193, 6194, 6195, 6196, 6197, 6198, 6199, 6200, 6201, 6202, 6203, "
'  ArrString = ArrString + "6391, 6392, 6393, 6394, 6395, 6396, 6397, 6398, "
'  ArrString = ArrString + "6399, 6400, 6401, 6402, 6403, 6076, 6088, 6218, 8712, "
  'fix for 3868 6/18/2010
  ArrString = "55516, 55515, 55514, 51287, "
  Dim CArr() As Long
  Call BuildArray(ArrString, CArr(), Cnt)
  
  For x = 1 To Cnt
    ClearTrans (CArr(x))
  Next x
  'fix for 1383 6/18/2010
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Get TTHandle, 55727, TaxTrans
      TaxTrans.RealPin = ""
  Put TTHandle, 55727, TaxTrans
  
  
  'fix for #3868
'  Get TTHandle, 5055, TaxTrans
'   TaxTrans.Amount = 279.68
'   TaxTrans.Revenue.Principle1 = 279.68
'   TaxTrans.Revenue.Principle1Pd = 279.68
'   Put TTHandle, 5055, TaxTrans
 
  
'  OpenTaxCustFile TCHandle, NumOfTaxCusts
  
  'fix for #596 on 10/1/09
'  Get TTHandle, 20419, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Interest = 7.36
'  Put TTHandle, 20419, TaxTrans
  
   'fix for 596 on 9/29/09
'  ClearTrans (32957)
'  ClearTrans (32956)
'  ClearTrans (32955)
'  ClearTrans (32954)
 
'  Get TCHandle, 3724, TaxCust
'  NextRec = TaxCust.LastTrans
'  Do While NextRec > 0
'    Get TTHandle, NextRec, TaxTrans
'    If NextRec >= 25420 Then
'      ClearTrans (NextRec)
'    End If
'    NextRec = TaxTrans.LastTrans
'  Loop
  
  'fix for #1627 9/11/2009
'  ClearTrans (32938)
'  ClearTrans (25726)
'  ClearTrans (25727)
'  Get TTHandle, 17014, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 1.84
'  Put TTHandle, 17014, TaxTrans
'
'  Get TTHandle, 17013, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 13.07
'  Put TTHandle, 17013, TaxTrans
'
'  'fix for #3621 9/11/2009
'  Get TTHandle, 32937, TaxTrans
'  TaxTrans.RealPin = ""
'  Put TTHandle, 32937, TaxTrans
  
  'fix for #3621 on 6/12/09
'  Get TTHandle, 32025, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  Put TTHandle, 32025, TaxTrans
  
  
  'fix for #1187 on 3/10/09
'  For x = 23617 To NumOfTTRecs
'    Get TTHandle, x, TaxTrans
'    If TaxTrans.BelongTo = 19268 Then
'      TaxTrans.Amount = 0
'      TaxTrans.Revenue.Collection = 0
'      TaxTrans.Revenue.CollectionPd = 0
'      TaxTrans.Revenue.Interest = 0
'      TaxTrans.Revenue.InterestPd = 0
'      TaxTrans.Revenue.LateList = 0
'      TaxTrans.Revenue.LateListPd = 0
'      TaxTrans.Revenue.Penalty = 0
'      TaxTrans.Revenue.PenaltyPd = 0
'      TaxTrans.Revenue.PrePaidUsed = 0
'      TaxTrans.Revenue.Principle1 = 0
'      TaxTrans.Revenue.Principle1Pd = 0
'      TaxTrans.Revenue.Principle2 = 0
'      TaxTrans.Revenue.Principle2Pd = 0
'      TaxTrans.Revenue.Principle3 = 0
'      TaxTrans.Revenue.Principle3Pd = 0
'      TaxTrans.Revenue.Principle4 = 0
'      TaxTrans.Revenue.Principle4Pd = 0
'      TaxTrans.Revenue.Principle5 = 0
'      TaxTrans.Revenue.Principle5Pd = 0
'      TaxTrans.Revenue.RevOpt1 = 0
'      TaxTrans.Revenue.RevOpt1Pd = 0
'      TaxTrans.Revenue.RevOpt2 = 0
'      TaxTrans.Revenue.RevOpt2Pd = 0
'      TaxTrans.Revenue.RevOpt3 = 0
'      TaxTrans.Revenue.RevOpt3Pd = 0
'      Put TTHandle, x, TaxTrans
'    End If
'  Next x
'
'  Get TTHandle, 19268, TaxTrans
'  TaxTrans.Revenue.Principle1 = 51.2
'  Put TTHandle, 19268, TaxTrans
  
  Close
  MsgBox "Completed successfully."

End Sub



Private Sub cmdFixBSL_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxCustRec As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTaxCusts As Long
  Dim x As Long
  OpenTaxTransFile TTHandle, NumOfTTRecs
  
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    If TaxTrans.TranType = 6 And TaxTrans.BelongTo = 23861 Then
      TaxTrans.RealPin = "156NB00609"
      Put TTHandle, x, TaxTrans
    End If
    If TaxTrans.TranType = 6 And TaxTrans.BelongTo = 23640 Then
      TaxTrans.RealPin = "157OG020"
      Put TTHandle, x, TaxTrans
    End If
    If TaxTrans.TranType = 6 And TaxTrans.BelongTo = 23639 Then
      TaxTrans.RealPin = "157OG021"
      Put TTHandle, x, TaxTrans
    End If
    If TaxTrans.TranType = 6 And TaxTrans.BelongTo = 23601 Then
      TaxTrans.RealPin = "157AB046"
      Put TTHandle, x, TaxTrans
    End If
    If TaxTrans.TranType = 6 And TaxTrans.BelongTo = 23958 Then
      TaxTrans.RealPin = "173HD00706"
      Put TTHandle, x, TaxTrans
    End If
    If TaxTrans.TranType = 6 And TaxTrans.BelongTo = 23069 Then
      TaxTrans.RealPin = "142JB010"
      Put TTHandle, x, TaxTrans
    End If
    If TaxTrans.TranType = 6 And TaxTrans.BelongTo = 23070 Then
      TaxTrans.RealPin = "142JB011"
      Put TTHandle, x, TaxTrans
    End If
    If TaxTrans.TranType = 6 And TaxTrans.BelongTo = 23618 Then
      TaxTrans.RealPin = "14200009"
      Put TTHandle, x, TaxTrans
    End If
    If TaxTrans.TranType = 6 And TaxTrans.BelongTo = 23000 Then
      TaxTrans.RealPin = "142OA007"
      Put TTHandle, x, TaxTrans
    End If
    If TaxTrans.TranType = 6 And TaxTrans.BelongTo = 23770 Then
      TaxTrans.RealPin = "156MC00204"
      Put TTHandle, x, TaxTrans
    End If
    If TaxTrans.TranType = 6 And TaxTrans.BelongTo = 18555 Then
      TaxTrans.RealPin = "156NF012"
      Put TTHandle, x, TaxTrans
    End If
    If TaxTrans.TranType = 6 And TaxTrans.BelongTo = 18703 Then
      TaxTrans.RealPin = "142GH002"
      Put TTHandle, x, TaxTrans
    End If
    If TaxTrans.TranType = 6 And TaxTrans.BelongTo = 19513 Then
      TaxTrans.RealPin = "142GD01201"
      Put TTHandle, x, TaxTrans
    End If
    If TaxTrans.TranType = 6 And TaxTrans.BelongTo = 20306 Then
      TaxTrans.RealPin = "156NF017"
      Put TTHandle, x, TaxTrans
    End If
    If TaxTrans.TranType = 6 And TaxTrans.BelongTo = 20307 Then
      TaxTrans.RealPin = "156NF019"
      Put TTHandle, x, TaxTrans
    End If
    If TaxTrans.TranType = 6 And TaxTrans.BelongTo = 20308 Then
      TaxTrans.RealPin = "156NF018"
      Put TTHandle, x, TaxTrans
    End If
    If TaxTrans.TranType = 6 And TaxTrans.BelongTo = 20776 Then
      TaxTrans.RealPin = "157OE010"
      Put TTHandle, x, TaxTrans
    End If
    If TaxTrans.TranType = 6 And TaxTrans.BelongTo = 20971 Then
      TaxTrans.RealPin = "156JA003"
      Put TTHandle, x, TaxTrans
    End If
    If TaxTrans.TranType = 6 And TaxTrans.BelongTo = 23709 Then
      TaxTrans.RealPin = "142OA008"
      Put TTHandle, x, TaxTrans
    End If
    If TaxTrans.TranType = 6 And TaxTrans.BelongTo = 24123 Then
      TaxTrans.RealPin = "157GF134"
      Put TTHandle, x, TaxTrans
    End If
  Next

  Get TTHandle, 71955, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  Put TTHandle, 71955, TaxTrans
  
  
'  'fix for 572
'  Get TTHandle, 198, TaxTrans
'  TaxTrans.Revenue.Interest = 255.98
'  TaxTrans.CustomerRec = 572
'  TaxTrans.CustPin = 572
'  Put TTHandle, 198, TaxTrans
'
'  'fix for 10402
'  Get TTHandle, 17416, TaxTrans
'  TaxTrans.Revenue.Interest = 0
'  TaxTrans.Amount = 0
'  TaxTrans.BelongTo = 198
'  Put TTHandle, 17416, TaxTrans
'
'  Get TTHandle, 28947, TaxTrans
'  TaxTrans.Revenue.Interest = 0
'  TaxTrans.Amount = 0
'  TaxTrans.BelongTo = 198
'  Put TTHandle, 28947, TaxTrans
'
'  Get TTHandle, 38384, TaxTrans
'  TaxTrans.Revenue.Interest = 0
'  TaxTrans.Amount = 0
'  TaxTrans.BelongTo = 198
'  Put TTHandle, 38384, TaxTrans
'
'  Get TTHandle, 42266, TaxTrans
'  TaxTrans.Revenue.Interest = 0
'  TaxTrans.Amount = 0
'  TaxTrans.BelongTo = 198
'  Put TTHandle, 42266, TaxTrans
'
'  'fix for 10754
'  Get TTHandle, 71164, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 8.45
'  Put TTHandle, 71164, TaxTrans
'
'  Get TTHandle, 78178, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 8.45
'  TaxTrans.Amount = 8.45
'  Put TTHandle, 78178, TaxTrans
'
'  Get TTHandle, 78179, TaxTrans
'  TaxTrans.Amount = 17.97
'  TaxTrans.Revenue.PrePaidAmt = 17.97
'  TaxTrans.Revenue.PrePaidBal = 17.97
'  TaxTrans.Revenue.PrePaidUsed = 0
'  Put TTHandle, 78179, TaxTrans
'
'  'fix for 10755
'  Get TTHandle, 71165, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 2.03
'  Put TTHandle, 71165, TaxTrans
'
'  Get TTHandle, 78180, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 2.03
'  TaxTrans.Amount = 2.03
'  Put TTHandle, 78180, TaxTrans
'
'  Get TTHandle, 78181, TaxTrans
'  TaxTrans.Amount = 17.97
'  TaxTrans.Revenue.PrePaidAmt = 17.97
'  TaxTrans.Revenue.PrePaidBal = 17.97
'  TaxTrans.Revenue.PrePaidUsed = 0
'  Put TTHandle, 78181, TaxTrans
'
'  'fix for 3795
'  Get TTHandle, 20776, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Interest = 1.59
'  TaxTrans.Revenue.InterestPd = 0
'  TaxTrans.Revenue.CollectionPd = 0
'  Put TTHandle, 20776, TaxTrans
'
'  'fix for 1490
'  Get TTHandle, 18703, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 51.48
'  TaxTrans.Revenue.Interest = 2.9
'  TaxTrans.Revenue.InterestPd = 2.74
'  TaxTrans.Revenue.CollectionPd = 1.5
'  Put TTHandle, 18703, TaxTrans
'
'  'fix for 11365
'  Get TTHandle, 71775, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 48#
'  Put TTHandle, 71775, TaxTrans
'
'  Get TTHandle, 83926, TaxTrans
'  TaxTrans.Amount = 48.96
'  TaxTrans.Revenue.Principle1Pd = 48#
'  Put TTHandle, 83926, TaxTrans
'
'  Get TTHandle, 83927, TaxTrans
'  TaxTrans.Amount = 3.31
'  TaxTrans.Revenue.PrePaidAmt = 3.31
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.PrePaidBal = 3.31
'  Put TTHandle, 83927, TaxTrans
'
'  'fix for 11366
'  Get TTHandle, 71776, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 48#
'  Put TTHandle, 71776, TaxTrans
'
'  Get TTHandle, 83928, TaxTrans
'  TaxTrans.Amount = 48.96
'  TaxTrans.Revenue.Principle1Pd = 48#
'  Put TTHandle, 83928, TaxTrans
'
'  Get TTHandle, 83929, TaxTrans
'  TaxTrans.Amount = 3.31
'  TaxTrans.Revenue.PrePaidAmt = 3.31
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.PrePaidBal = 3.31
'  Put TTHandle, 83929, TaxTrans
'
'  'fix for 11367
'  Get TTHandle, 71777, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 84#
'  Put TTHandle, 71777, TaxTrans
'
'  Get TTHandle, 83930, TaxTrans
'  TaxTrans.Amount = 85.68
'  TaxTrans.Revenue.Principle1Pd = 84#
'  Put TTHandle, 83930, TaxTrans
'
'  Get TTHandle, 83931, TaxTrans
'  TaxTrans.Amount = 3.31
'  TaxTrans.Revenue.PrePaidAmt = 3.31
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.PrePaidBal = 3.31
'  Put TTHandle, 83931, TaxTrans
'
'  'fix for 1029
'  Get TTHandle, 18290, TaxTrans
'  TaxTrans.LastTrans = 10328
'  Put TTHandle, 18290, TaxTrans
'
'  Get TTHandle, 10328, TaxTrans
'  TaxTrans.LastTrans = 10330
'  TaxTrans.TranType = 2
'  TaxTrans.Revenue.Principle1Pd = 9.18
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Description = "Pay on 10330"
'  TaxTrans.BelongTo = 10330
'  Put TTHandle, 10328, TaxTrans
'
'  Get TTHandle, 10330, TaxTrans
'  TaxTrans.LastTrans = 0
'  Put TTHandle, 10330, TaxTrans
  
'  Get TTHandle, 49346, TaxTrans
'  TaxTrans.RealPin = "173GC00202"
'  Put TTHandle, 49346, TaxTrans
'
'  OpenTaxCustFile TCHandle, NumOfTaxCusts
'  Get TCHandle, 6256, TaxCustRec
'  TaxCustRec.LastTrans = 51957
'  Put TCHandle, 6256, TaxCustRec
'
'  Get TTHandle, 71955, TaxTrans
'  TaxTrans.CustomerRec = 11265
'  TaxTrans.CustPin = 11265
'  Put TTHandle, 71955, TaxTrans
'
'  Get TTHandle, 48807, TaxTrans
'  TaxTrans.RealPin = "127JA00105"
'  Put TTHandle, 48807, TaxTrans
'
'  Call FixPrePayOnlyRealPins
  Close
  MsgBox "Completed successfully."
End Sub

Private Sub FixPrePayOnlyRealPins()
  Dim TaxTrans As TaxTransactionType
  Dim TT2Handle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  
  OpenTaxTransFile TT2Handle, NumOfTTRecs
  For x = 1 To NumOfTTRecs
    Get TT2Handle, x, TaxTrans
    If TaxTrans.TranType = 22 Then
      TaxTrans.RealPin = ""
      Put TT2Handle, x, TaxTrans
    End If
  Next x
  Close TT2Handle
End Sub

Private Sub cmdFixBSLAgain_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim FileName As String
  Dim ThisFile As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTaxCusts As Long
  
  FileName = "BoilingErrorsAgain.txt"
  ThisFile = FreeFile
  Open FileName For Output As ThisFile
  Print #ThisFile, "Cust Acct # ~ Cust Name ~ Action Taken ~ Amount ~ Trans Rec #"
  OpenTaxCustFile TCHandle, NumOfTaxCusts
  OpenTaxTransFile TTHandle, NumOfTTRecs
  
  'fix for 10754
  Get TTHandle, 76130, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Interest = 0
  TaxTrans.Description = "ErrorFix"
  Get TCHandle, TaxTrans.CustomerRec, TaxCust
  Print #ThisFile, CStr(TaxCust.Acct) + "~" + QPTrim$(TaxCust.CustName) + "~ Interest Removed" + "~" + Using$("##.##", CStr(TaxTrans.Amount)) + "~ 76130"
  Put TTHandle, 76130, TaxTrans
  
  Get TTHandle, 78179, TaxTrans
  TaxTrans.Amount = 18.81
  TaxTrans.TaxYear = 2007
  TaxTrans.BelongTo = 0
  TaxTrans.Revenue.PrePaidAmt = 18.81
  TaxTrans.TranType = 22
  TaxTrans.Revenue.Principle1Pd = 18.81
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.CollectionPd = 0
  TaxTrans.RealPin = ""
  TaxTrans.PersPin = ""
  TaxTrans.Description = "Prepay"
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.PrePaidBal = OldRound(GetOverPayBalance(TaxTrans.CustomerRec) + 18.81)
  Get TCHandle, TaxTrans.CustomerRec, TaxCust
  Print #ThisFile, CStr(TaxCust.Acct) + "~" + QPTrim$(TaxCust.CustName) + "~ Changed Trans From Pay To Prepay" + "~" + Using$("##.##", CStr(TaxTrans.Amount)) + "~ 78179"
  Put TTHandle, 78179, TaxTrans
  
  'fix for 10753
  Get TTHandle, 76129, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Interest = 0
  TaxTrans.Description = "ErrorFix"
  Get TCHandle, TaxTrans.CustomerRec, TaxCust
  Print #ThisFile, CStr(TaxCust.Acct) + "~" + QPTrim$(TaxCust.CustName) + "~ Interest Removed" + "~" + Using$("##.##", CStr(TaxTrans.Amount)) + "~ 76129"
  Put TTHandle, 76129, TaxTrans
  
  'fix for 10755
  Get TTHandle, 76131, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Interest = 0
  TaxTrans.Description = "ErrorFix"
  Get TCHandle, TaxTrans.CustomerRec, TaxCust
  Print #ThisFile, CStr(TaxCust.Acct) + "~" + QPTrim$(TaxCust.CustName) + "~ Interest Removed" + "~" + Using$("##.##", CStr(TaxTrans.Amount)) + "~ 76131"
  Put TTHandle, 76131, TaxTrans
  
  Get TTHandle, 78181, TaxTrans
  TaxTrans.Amount = 18.81
  TaxTrans.TaxYear = 2007
  TaxTrans.BelongTo = 0
  TaxTrans.Revenue.PrePaidAmt = 18.81
  TaxTrans.TranType = 22
  TaxTrans.Revenue.Principle1Pd = 18.81
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.CollectionPd = 0
  TaxTrans.RealPin = ""
  TaxTrans.PersPin = ""
  TaxTrans.Description = "Prepay"
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.PrePaidBal = OldRound(GetOverPayBalance(TaxTrans.CustomerRec) + 18.81)
  Get TCHandle, TaxTrans.CustomerRec, TaxCust
  Print #ThisFile, CStr(TaxCust.Acct) + "~" + QPTrim$(TaxCust.CustName) + "~ Changed Trans From Pay To Prepay" + "~" + Using$("##.##", CStr(TaxTrans.Amount)) + "~ 78181"
  Put TTHandle, 78181, TaxTrans
  
  'fix for 11365
  Get TTHandle, 76132, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Interest = 0
  TaxTrans.Description = "ErrorFix"
  Get TCHandle, TaxTrans.CustomerRec, TaxCust
  Print #ThisFile, CStr(TaxCust.Acct) + "~" + QPTrim$(TaxCust.CustName) + "~ Interest Removed" + "~" + Using$("##.##", CStr(TaxTrans.Amount)) + "~ 76132"
  Put TTHandle, 76132, TaxTrans
  
  Get TTHandle, 83744, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Interest = 0
  TaxTrans.Description = "ErrorFix"
  Get TCHandle, TaxTrans.CustomerRec, TaxCust
  Print #ThisFile, CStr(TaxCust.Acct) + "~" + QPTrim$(TaxCust.CustName) + "~ Interest Removed" + "~" + Using$("##.##", CStr(TaxTrans.Amount)) + "~ 83744"
  Put TTHandle, 83744, TaxTrans
  
  Get TTHandle, 83927, TaxTrans
  TaxTrans.Amount = 3.49
  TaxTrans.TaxYear = 2007
  TaxTrans.BelongTo = 0
  TaxTrans.Revenue.PrePaidAmt = 3.49
  TaxTrans.TranType = 22
  TaxTrans.Revenue.Principle1Pd = 3.49
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.CollectionPd = 0
  TaxTrans.RealPin = ""
  TaxTrans.PersPin = ""
  TaxTrans.Description = "Prepay"
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.PrePaidBal = OldRound(GetOverPayBalance(TaxTrans.CustomerRec) + 3.49)
  Get TCHandle, TaxTrans.CustomerRec, TaxCust
  Print #ThisFile, CStr(TaxCust.Acct) + "~" + QPTrim$(TaxCust.CustName) + "~ Changed Trans From Pay To Prepay" + "~" + Using$("##.##", CStr(TaxTrans.Amount)) + "~ 83927"
  Put TTHandle, 83927, TaxTrans
  
  'fix for 11366
  Get TTHandle, 76133, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Interest = 0
  TaxTrans.Description = "ErrorFix"
  Get TCHandle, TaxTrans.CustomerRec, TaxCust
  Print #ThisFile, CStr(TaxCust.Acct) + "~" + QPTrim$(TaxCust.CustName) + "~ Interest Removed" + "~" + Using$("##.##", CStr(TaxTrans.Amount)) + "~ 76133"
  Put TTHandle, 76133, TaxTrans
  
  Get TTHandle, 83746, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Interest = 0
  TaxTrans.Description = "ErrorFix"
  Get TCHandle, TaxTrans.CustomerRec, TaxCust
  Print #ThisFile, CStr(TaxCust.Acct) + "~" + QPTrim$(TaxCust.CustName) + "~ Interest Removed" + "~" + Using$("##.##", CStr(TaxTrans.Amount)) + "~ 83746"
  Put TTHandle, 83746, TaxTrans
  
  Get TTHandle, 83929, TaxTrans
  TaxTrans.Amount = 3.49
  TaxTrans.TaxYear = 2007
  TaxTrans.BelongTo = 0
  TaxTrans.Revenue.PrePaidAmt = 3.49
  TaxTrans.TranType = 22
  TaxTrans.Revenue.Principle1Pd = 3.49
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.CollectionPd = 0
  TaxTrans.RealPin = ""
  TaxTrans.PersPin = ""
  TaxTrans.Description = "Prepay"
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.PrePaidBal = OldRound(GetOverPayBalance(TaxTrans.CustomerRec) + 3.49)
  Get TCHandle, TaxTrans.CustomerRec, TaxCust
  Print #ThisFile, CStr(TaxCust.Acct) + "~" + QPTrim$(TaxCust.CustName) + "~ Changed Trans From Pay To Prepay" + "~" + Using$("##.##", CStr(TaxTrans.Amount)) + "~ 83929"
  Put TTHandle, 83929, TaxTrans
  
  'fix for 11367
  Get TTHandle, 76134, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Interest = 0
  TaxTrans.Description = "ErrorFix"
  Get TCHandle, TaxTrans.CustomerRec, TaxCust
  Print #ThisFile, CStr(TaxCust.Acct) + "~" + QPTrim$(TaxCust.CustName) + "~ Interest Removed" + "~" + Using$("##.##", CStr(TaxTrans.Amount)) + "~ 76134"
  Put TTHandle, 76134, TaxTrans
  
  Get TTHandle, 83748, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Interest = 0
  TaxTrans.Description = "ErrorFix"
  Get TCHandle, TaxTrans.CustomerRec, TaxCust
  Print #ThisFile, CStr(TaxCust.Acct) + "~" + QPTrim$(TaxCust.CustName) + "~ Interest Removed" + "~" + Using$("##.##", CStr(TaxTrans.Amount)) + "~ 83748"
  Put TTHandle, 83748, TaxTrans
  
  Get TTHandle, 83931, TaxTrans
  TaxTrans.Amount = 3.49
  TaxTrans.TaxYear = 2007
  TaxTrans.BelongTo = 0
  TaxTrans.Revenue.PrePaidAmt = 3.49
  TaxTrans.TranType = 22
  TaxTrans.Revenue.Principle1Pd = 3.49
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.CollectionPd = 0
  TaxTrans.RealPin = ""
  TaxTrans.PersPin = ""
  TaxTrans.Description = "Prepay"
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.PrePaidBal = OldRound(GetOverPayBalance(TaxTrans.CustomerRec) + 3.49)
  Get TCHandle, TaxTrans.CustomerRec, TaxCust
  Print #ThisFile, CStr(TaxCust.Acct) + "~" + QPTrim$(TaxCust.CustName) + "~ Changed Trans From Pay To Prepay" + "~" + Using$("##.##", CStr(TaxTrans.Amount)) + "~ 83931"
  Put TTHandle, 83931, TaxTrans
  
  Close
  MsgBox ("Finished. Look for BoilingErrorsAgain.txt.")
End Sub

Private Sub cmdFixBSLMay_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Get TTHandle, 87496, TaxTrans
  TaxTrans.RealPin = "142JC004"
  Put TTHandle, 87496, TaxTrans
  
  Get TTHandle, 87494, TaxTrans
  TaxTrans.RealPin = "142JC004"
  Put TTHandle, 87494, TaxTrans
  
  Close
  MsgBox ("Finished.")
End Sub

Private Sub FixFairmontCust1()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  'fix for cust #1 on 2/23/09
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    If TaxTrans.BelongTo = 8933 Or TaxTrans.BelongTo = 8934 Then
      TaxTrans.Amount = 0
      TaxTrans.DiscAmt = 0
      TaxTrans.Revenue.Principle1 = 0
      TaxTrans.Revenue.Principle1Pd = 0
      TaxTrans.Revenue.Interest = 0
      TaxTrans.Revenue.InterestPd = 0
      TaxTrans.Revenue.PrePaidAmt = 0
      TaxTrans.Revenue.PrePaidBal = 0
      TaxTrans.Revenue.PrePaidUsed = 0
      Put TTHandle, x, TaxTrans
    End If
  Next x
  
  Get TTHandle, 8933, TaxTrans
  TaxTrans.DiscAmt = 0
  Put TTHandle, 8933, TaxTrans
  
  Get TTHandle, 8934, TaxTrans
  TaxTrans.DiscAmt = 0
  Put TTHandle, 8934, TaxTrans
  
  Get TTHandle, 9497, TaxTrans
  TaxTrans.Amount = 33.83
  TaxTrans.Revenue.Principle1Pd = 32.43
  TaxTrans.Revenue.InterestPd = 1.4
  Put TTHandle, 9497, TaxTrans
  
  Get TTHandle, 9498, TaxTrans
  TaxTrans.Amount = 586.5
  TaxTrans.Revenue.Principle1Pd = 586.5
  Put TTHandle, 9498, TaxTrans
  
  Get TTHandle, 13626, TaxTrans
  TaxTrans.Amount = 1.4
  TaxTrans.Revenue.Interest = 1.4
  Put TTHandle, 13626, TaxTrans
  
  Get TTHandle, 42070, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  Put TTHandle, 42070, TaxTrans
  
  Get TTHandle, 44064, TaxTrans
  TaxTrans.Amount = 586.5
  TaxTrans.Revenue.Principle1Pd = 586.5
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  Put TTHandle, 44064, TaxTrans
  
  Close
  
End Sub

Private Sub cmdFixFairmont_Click()
   Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim NextRec As Long
  Dim ODate As Integer
  Call FixFairmontCust1
  
  ODate = Date2Num("1/1/2007")
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  'fix for cust #2059 on 2/23/09
  Get TTHandle, 9609, TaxTrans
  TaxTrans.Amount = 33.81
  TaxTrans.Revenue.Principle1Pd = 33.81
  Put TTHandle, 9609, TaxTrans

  Get TTHandle, 9205, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 33.81
  Put TTHandle, 9205, TaxTrans
  
  Get TTHandle, 9608, TaxTrans
  TaxTrans.Amount = 519.32
  TaxTrans.Revenue.Principle1Pd = 519.32
  Put TTHandle, 9608, TaxTrans

  Get TTHandle, 9206, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 519.32
  Put TTHandle, 9206, TaxTrans
  

  'fix for cust #1722 on 2/23/09
  
  Get TTHandle, 7540, TaxTrans
  TaxTrans.Revenue.Principle1 = 802.47
  TaxTrans.Revenue.Principle1Pd = 786.42
  Put TTHandle, 7540, TaxTrans

  Get TTHandle, 7541, TaxTrans
  TaxTrans.Revenue.Principle1 = 57.27
  TaxTrans.Revenue.Principle1Pd = 57.27 '41.22
  TaxTrans.DiscAmt = 0
  TaxTrans.Revenue.Interest = 0.32
  TaxTrans.Revenue.InterestPd = 0.32
  Put TTHandle, 7541, TaxTrans

  Get TTHandle, 9623, TaxTrans
  TaxTrans.Amount = 41.22
  TaxTrans.Revenue.Principle1Pd = 41.22
  TaxTrans.DiscAmt = 0
  Put TTHandle, 9623, TaxTrans
  
  Get TTHandle, 9624, TaxTrans
  TaxTrans.Amount = 786.42
  TaxTrans.Revenue.Principle1Pd = 786.42
  Put TTHandle, 9624, TaxTrans
  
  Get TTHandle, 12848, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.Principle1 = 0
  Put TTHandle, 12848, TaxTrans
  
'  Get TTHandle, 12847, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.Interest = 0
'  Put TTHandle, 12847, TaxTrans
'
  'fix for cust #1026 on 2/23/09
  Get TTHandle, 9469, TaxTrans
  TaxTrans.Revenue.Principle1 = 0
  TaxTrans.DiscAmt = 0
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Principle1Pd = 0
  Put TTHandle, 9469, TaxTrans
 
  Get TTHandle, 8769, TaxTrans
  TaxTrans.Amount = 734.16
  TaxTrans.Revenue.Principle1 = 734.16
  TaxTrans.Revenue.Principle1Pd = 719.48
  TaxTrans.DiscAmt = 14.68
  Put TTHandle, 8769, TaxTrans
  
  'fix for cust #170 on 2/23/09
  Get TTHandle, 8846, TaxTrans
  TaxTrans.Amount = 527.16
  TaxTrans.Revenue.Principle1Pd = 516.62
  TaxTrans.DiscAmt = 10.54
  Put TTHandle, 8846, TaxTrans
  
  Get TTHandle, 9309, TaxTrans
  TaxTrans.Revenue.Principle1 = 0
  TaxTrans.DiscAmt = 0
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Principle1Pd = 0
  Put TTHandle, 9309, TaxTrans

  Get TTHandle, 9310, TaxTrans
  TaxTrans.Revenue.Principle1 = 0
  TaxTrans.DiscAmt = 0
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Principle1Pd = 0
  Put TTHandle, 9310, TaxTrans
  
  Get TTHandle, 9311, TaxTrans
  TaxTrans.Revenue.Principle1 = 0
  TaxTrans.DiscAmt = 0
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Principle1Pd = 0
  Put TTHandle, 9311, TaxTrans
  
  
  'fix for cust #443 on 2/23/09
  Get TTHandle, 8421, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 1477.5
  Put TTHandle, 8421, TaxTrans
  
  Get TTHandle, 9344, TaxTrans
  TaxTrans.Amount = 1477.5
  TaxTrans.Revenue.Principle1Pd = 1477.5
  Put TTHandle, 9344, TaxTrans
  
  Get TTHandle, 8423, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 137.27
  Put TTHandle, 8423, TaxTrans

  Get TTHandle, 9343, TaxTrans
  TaxTrans.Amount = 137.27
  TaxTrans.Revenue.Principle1Pd = 137.27
  Put TTHandle, 9343, TaxTrans

  
  Get TTHandle, 8422, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 1323.32
  Put TTHandle, 8422, TaxTrans

  Get TTHandle, 9342, TaxTrans
  TaxTrans.Amount = 1323.32
  TaxTrans.Revenue.Principle1Pd = 1323.32
  Put TTHandle, 9342, TaxTrans
  
  
  'fix for cust #10 on 2/17/09
  Get TTHandle, 9261, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 47.33
  Put TTHandle, 9261, TaxTrans
  
  Get TTHandle, 9408, TaxTrans
  TaxTrans.Amount = 47.33
  TaxTrans.Revenue.Principle1Pd = 47.33
  Put TTHandle, 9408, TaxTrans
  
'  Get TCHandle, 1751, TaxCust
'  NextRec = TaxCust.LastTrans
'  Do While NextRec > 0
'    Get TTHandle, NextRec, TaxTrans
'    If TaxTrans.TransDate >= ODate Then
'      If TaxTrans.BelongTo = 7884 Then
'        GoSub Zero
'      End If
'    End If
'    NextRec = TaxTrans.LastTrans
'  Loop
'
'  Get TTHandle, 9371, TaxTrans
'  TaxTrans.Amount = 1987.59
'  TaxTrans.Revenue.Principle1Pd = 1987.59
'  Put TTHandle, 9371, TaxTrans
'
'  Get TTHandle, 7884, TaxTrans
'  TaxTrans.Amount = 2027.35
'  TaxTrans.Revenue.Principle1 = 2027.35
'  TaxTrans.Revenue.Principle1Pd = 1987.59
'  Put TTHandle, 7884, TaxTrans
'
'  Get TTHandle, 9372, TaxTrans
'  TaxTrans.Amount = 16069.41
'  TaxTrans.Revenue.Principle1Pd = 16069.41
'  Put TTHandle, 9372, TaxTrans
'
'  Get TTHandle, 7883, TaxTrans
'  TaxTrans.Amount = 16390.8
'  TaxTrans.Revenue.Principle1 = 16390.8
'  TaxTrans.Revenue.Principle1Pd = 16069.41
'  Put TTHandle, 7883, TaxTrans
  
  Close
  MsgBox ("Finished.")
  
  Exit Sub
  
Zero:
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.Principle1 = 0
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.Interest = 0
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.Collection = 0
  TaxTrans.Revenue.CollectionPd = 0
  TaxTrans.Revenue.LateList = 0
  TaxTrans.Revenue.LateListPd = 0
  TaxTrans.Revenue.RevOpt1 = 0
  TaxTrans.Revenue.RevOpt1Pd = 0
  TaxTrans.Revenue.RevOpt2 = 0
  TaxTrans.Revenue.RevOpt2Pd = 0
  TaxTrans.Revenue.RevOpt3 = 0
  TaxTrans.Revenue.RevOpt3Pd = 0
  TaxTrans.Amount = 0
  Put TTHandle, NextRec, TaxTrans
 Return
  
End Sub

Private Sub cmdFixCanton_Click()
  Call ClearTrans(113165)
  Call ClearTrans(113175)
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Get TTHandle, 106283, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 504.95
  TaxTrans.Revenue.InterestPd = 0
  Put TTHandle, 106283, TaxTrans
  Close
 
 MsgBox ("Done.")

End Sub

Private Sub cmdFixFaison_Click()
  Dim RealRec As PropertyRecType
  Dim NumOfRealRecs As Long
  Dim RRHandle As Integer

  OpenRealPropFile RRHandle, NumOfRealRecs
  Get RRHandle, 416, RealRec
  RealRec.RealPin = ""
  Put RRHandle, 416, RealRec
  Close
  MsgBox ("Done.")

End Sub

Private Sub cmdFixHarrisburg_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Get TTHandle, 25942, TaxTrans
  TaxTrans.RealPin = "01-010-B-0295.000000"
  Put TTHandle, 25942, TaxTrans
  Close
  MsgBox ("Done.")

End Sub

Private Sub cmdFixHBInsertFireTax_Click()
  Call AddFireTaxToHarrisburgPersPropBills
End Sub

Private Sub cmdFixHBRealOpt_Click()
  Call AddLateListTaxToHarrisburg
End Sub

Private Sub cmdFixHBLateList_Click()
  Call AddLateListTaxToHarrisburg
End Sub

Private Sub cmdFixKenansville_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  Dim ThisDate As Integer
  Dim Cnt As Integer
  
  ThisDate = Date2Num("10/27/2009")
  OpenTaxTransFile TTHandle, NumOfTTRecs
'  Cnt = 0
'  For x = 1 To NumOfTTRecs
'    Get TTHandle, x, TaxTrans
'    If TaxTrans.TransDate >= ThisDate And TaxTrans.TranType = 1 Then
'      Call ClearTrans(x)
'      Cnt = Cnt + 1
'    End If
'  Next x
'
'  Close
'  MsgBox ("A total of " + CStr(Cnt) + " billing transactions have been zeroed out.")
  Get TTHandle, 1612, TaxTrans
  TaxTrans.Amount = 389.84
  TaxTrans.Revenue.Principle1Pd = 389.84
  Put TTHandle, 1612, TaxTrans
  
  Get TTHandle, 1030, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 389.84
  Put TTHandle, 1030, TaxTrans
  
  Get TTHandle, 1611, TaxTrans
  TaxTrans.Amount = 98.23
  TaxTrans.Revenue.Principle1Pd = 98.23
  Put TTHandle, 1611, TaxTrans
  
  Get TTHandle, 1029, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 98.23
  Put TTHandle, 1029, TaxTrans
  
  Get TTHandle, 1610, TaxTrans
  TaxTrans.Amount = 74.15
  TaxTrans.Revenue.Principle1Pd = 74.15
  Put TTHandle, 1610, TaxTrans
  
  Get TTHandle, 1028, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 74.15
  Put TTHandle, 1028, TaxTrans
  
  Get TTHandle, 1614, TaxTrans
  TaxTrans.Amount = 875.24
  TaxTrans.Revenue.Principle1Pd = 875.24
  Put TTHandle, 1614, TaxTrans
  
  Get TTHandle, 1026, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 875.24
  Put TTHandle, 1026, TaxTrans
  
  Close
  MsgBox ("Completed.")

End Sub

Private Sub cmdFixMaggieValley_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim ArrString As String
  Dim Cnt As Integer, x As Integer
  Dim CArr() As Long
  ArrString = "102336, 102337, 102338, 102339, 102340, " '5/10/2010
  ArrString = ArrString + "102341, 102342, 102343, 102344, "
  ArrString = ArrString + "102345, 102346, 102347, 102348,"
  Call BuildArray(ArrString, CArr(), Cnt)
  
  For x = 1 To Cnt
    ClearTrans (CArr(x))
  Next x
  ArrString = "91271, 91270, 91269, 91268, 91267, 91266, 91265, "
  ArrString = ArrString + "91264, 91263, 91262, 91261, 91260, 91259, 91258"
  
  ReDim CArr(1 To 1) As Long
  Cnt = 0
  Call BuildArray(ArrString, CArr(), Cnt)
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For x = 1 To Cnt
    Get TTHandle, CArr(x), TaxTrans
    TaxTrans.Revenue.CollectionPd = TaxTrans.Revenue.Collection
    Put TTHandle, CArr(x), TaxTrans
  Next x
  Close
  MsgBox ("Done.")

End Sub

Private Sub cmdFixMagnolia_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxCustRec As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTaxCusts As Long
  Dim NextRec As Long
  Dim BelongTo As Long
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Get TTHandle, 31626, TaxTrans
  TaxTrans.RealPin = "5720274"
  Put TTHandle, 31626, TaxTrans
  
'  OpenTaxCustFile TCHandle, NumOfTaxCusts
'
'  Get TCHandle, 479, TaxCustRec
'  NextRec = TaxCustRec.LastTrans
'  BelongTo = 1173
'  Get TTHandle, BelongTo, TaxTrans
'  TaxTrans.Description = "1"
'  Put TTHandle, BelongTo, TaxTrans
'  Do While NextRec > 0
'    Get TTHandle, NextRec, TaxTrans
'    If TaxTrans.BelongTo = BelongTo Then
'      TaxTrans.Description = "1"
'      Put TTHandle, NextRec, TaxTrans
'    End If
'    NextRec = TaxTrans.LastTrans
'  Loop
'
'  NextRec = TaxCustRec.LastTrans
'  BelongTo = 1176
'  Get TTHandle, BelongTo, TaxTrans
'  TaxTrans.Description = "2"
'  Put TTHandle, BelongTo, TaxTrans
'  Do While NextRec > 0
'    Get TTHandle, NextRec, TaxTrans
'    If TaxTrans.BelongTo = BelongTo Then
'      TaxTrans.Description = "2"
'      Put TTHandle, NextRec, TaxTrans
'    End If
'    NextRec = TaxTrans.LastTrans
'  Loop
'
'  Get TCHandle, 482, TaxCustRec
'  NextRec = TaxCustRec.LastTrans
'  BelongTo = 1192
'  Get TTHandle, BelongTo, TaxTrans
'  TaxTrans.Description = "1"
'  Put TTHandle, BelongTo, TaxTrans
'  Do While NextRec > 0
'    Get TTHandle, NextRec, TaxTrans
'    If TaxTrans.BelongTo = BelongTo Then
'      TaxTrans.Description = "1"
'      Put TTHandle, NextRec, TaxTrans
'    End If
'    NextRec = TaxTrans.LastTrans
'  Loop
'
'  NextRec = TaxCustRec.LastTrans
'  BelongTo = 1195
'  Get TTHandle, BelongTo, TaxTrans
'  TaxTrans.Description = "2"
'  Put TTHandle, BelongTo, TaxTrans
'  Do While NextRec > 0
'    Get TTHandle, NextRec, TaxTrans
'    If TaxTrans.BelongTo = BelongTo Then
'      TaxTrans.Description = "2"
'      Put TTHandle, NextRec, TaxTrans
'    End If
'    NextRec = TaxTrans.LastTrans
'  Loop

  Close
  MsgBox ("Finished.")
  
  
End Sub

Private Sub cmdFixHildebran_Click()
  Dim TaxCust As TaxCustType
  Dim x As Integer
  Dim NumOfTCRecs As Long
  Dim THandle As Integer
  Dim Cnt As Integer
  OpenTaxCustFile THandle, NumOfTCRecs
  Dim Add1 As String
  Dim Add2 As String
'  For x = 1 To NumOfTCRecs
'    Get THandle, x, TaxCust
'    If QPTrim$(TaxCust.OptSrchDesc) <> "" Then
'      TaxCust.CountyAcctString = QPTrim$(TaxCust.OptSrchDesc)
'      Put THandle, x, TaxCust
'      cnt = cnt + 1
'    End If
'  Next x
  For x = 1 To NumOfTCRecs
    Get THandle, x, TaxCust
    Add1 = QPTrim$(TaxCust.Addr1)
    Add2 = QPTrim$(TaxCust.Addr2)
    If Add1 = Add2 And Add1 <> "" Then
      TaxCust.Addr2 = ""
      Put THandle, x, TaxCust
      Cnt = Cnt + 1
    End If
  Next
  
  Close
  MsgBox ("Updated " & CStr(Cnt) & " address 2s.")
End Sub

Private Sub cmdFixMaxton_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  '4/27/09
  'fix for 630
  Get TTHandle, 79629, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.CollectionPd = 0
  Put TTHandle, 79629, TaxTrans
  
  Get TTHandle, 79628, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.CollectionPd = 0
  Put TTHandle, 79628, TaxTrans
 
  Get TTHandle, 79622, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.CollectionPd = 0
  Put TTHandle, 79622, TaxTrans
 
  Get TTHandle, 79621, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.CollectionPd = 0
  Put TTHandle, 79621, TaxTrans
  
   Get TTHandle, 460, TaxTrans
  TaxTrans.Revenue.Principle1 = 437.68
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.Interest = 187.94
  TaxTrans.Revenue.InterestPd = 25.15
  TaxTrans.Revenue.Collection = 2.5
  TaxTrans.Revenue.CollectionPd = 2.5
  Put TTHandle, 460, TaxTrans
  
  Get TTHandle, 461, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 329.33
  TaxTrans.Revenue.Interest = 82.42
  TaxTrans.Revenue.InterestPd = 51.14
  TaxTrans.Revenue.Collection = 0
  Put TTHandle, 461, TaxTrans
 
  
'  'fix for 1160
'  Get TTHandle, 52285, TaxTrans
'  TaxTrans.LastTrans = 51748
'  Put TTHandle, 52285, TaxTrans
'
'  Get TTHandle, 52080, TaxTrans
'  TaxTrans.BelongTo = 38748
'  TaxTrans.CustomerRec = 64
'  TaxTrans.CustPin = 64
'  TaxTrans.LastTrans = 49070
'  Put TTHandle, 52080, TaxTrans
'
'  Get TTHandle, 52660, TaxTrans
'  TaxTrans.LastTrans = 52080
'  Put TTHandle, 52660, TaxTrans
  
  'fix for #837
'  Get TTHandle, 10838, TaxTrans
'  TaxTrans.BelongTo = 0
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  Put TTHandle, 10838, TaxTrans
'
'  Get TTHandle, 34861, TaxTrans
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.Principle1Pd = 56
'  TaxTrans.BelongTo = 31679
'  TaxTrans.TranType = 2
'  Put TTHandle, 34861, TaxTrans
'
'  Get TTHandle, 31679, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 56#
'  Put TTHandle, 31679, TaxTrans
  
  
'  Get TTHandle, 32591, TaxTrans
'  TaxTrans.Altered = 0
'  TaxTrans.Description = "SS Removed"
'  TaxTrans.FromPrePay = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  Put TTHandle, 32591, TaxTrans
'
'  Get TTHandle, 10967, TaxTrans
'  TaxTrans.Altered = 0
'  TaxTrans.FromPrePay = 0
'  TaxTrans.Description = "Bill #781"
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.TranType = 2
'  TaxTrans.TaxYear = 2006
'  TaxTrans.BelongTo = 10162
'  Put TTHandle, 10967, TaxTrans
'
'  Get TTHandle, 32590, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 0
'  Put TTHandle, 32590, TaxTrans
  Close
  
  MsgBox ("Done.")
     
End Sub

Private Sub cmdFixIndianTrails_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim NextRec As Long
  Dim StopDate As Integer
  Dim BelongTo As Long
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
   'fix #314 on 6/2/2010
  Get TTHandle, 91367, TaxTrans
  TaxTrans.LastTrans = 82051
  Put TTHandle, 91367, TaxTrans
  
  Get TTHandle, 82051, TaxTrans
  TaxTrans.LastTrans = 83748
  TaxTrans.BelongTo = 83748
  TaxTrans.Description = "Bill Num: 14277"
  Put TTHandle, 82051, TaxTrans
  
  Get TTHandle, 83748, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 3.4
  TaxTrans.Revenue.Interest = 0
  TaxTrans.LastTrans = 75078
  Put TTHandle, 83748, TaxTrans
  
  Get TTHandle, 57449, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 3.4
  Put TTHandle, 57449, TaxTrans
  
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    If TaxTrans.TranType = 4 And TaxTrans.BelongTo = 83748 Then
    Call ClearTrans(x)
    End If
  Next x
 
  
  'fix for 2276, 3129, 3578, 6251, 10689, 13890, 15656, 15658, 16229
'  StopDate = Date2Num("9/27/2006")
'  For x = 1 To NumOfTCRecs
'    Get TCHandle, x, TaxCust
'    Select Case x
'     Case 2276, 3129, 3578, 6251, 10689, 13890, 15656, 15658, 16229
'      GoSub FixIndianTrails
'    End Select
'  Next x
  
  Close
  MsgBox ("Done.")
  Exit Sub
  
FixIndianTrails:
  NextRec = TaxCust.LastTrans
  Do While NextRec > 0
    Get TTHandle, NextRec, TaxTrans
    If TaxTrans.TranType = 4 And TaxTrans.TransDate > StopDate Then
      TaxTrans.Revenue.Interest = 0
      TaxTrans.Amount = 0
      BelongTo = TaxTrans.BelongTo
      Put TTHandle, NextRec, TaxTrans
    ElseIf TaxTrans.TranType = 6 And TaxTrans.TransDate > StopDate Then
      TaxTrans.Revenue.Collection = 0
      TaxTrans.Amount = 0
      Put TTHandle, NextRec, TaxTrans
    End If
    NextRec = TaxTrans.LastTrans
  Loop
  Get TTHandle, BelongTo, TaxTrans
  TaxTrans.Revenue.Interest = 0
  TaxTrans.Revenue.Collection = 0
  Put TTHandle, BelongTo, TaxTrans
    
  Return
  
  'fix for custs# 443 & 3129
'  Get TTHandle, 269931, TaxTrans
'  TaxTrans.BelongTo = 252455
'  TaxTrans.Description = "252455"
'  TaxTrans.RealPin = "7084037"
'  Put TTHandle, 269931, TaxTrans
'
'  Get TTHandle, 249968, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 131.06
'  Put TTHandle, 249968, TaxTrans
'
'  Get TTHandle, 252455, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 4.25
'  Put TTHandle, 252455, TaxTrans
'
'  Get TTHandle, 298042, TaxTrans
'  TaxTrans.RealPin = "7084037"
'  Put TTHandle, 298042, TaxTrans
'
'  'fix for custs #1234 & 3578
'  Get TTHandle, 270233, TaxTrans
'  TaxTrans.BelongTo = 252879
'  TaxTrans.RealPin = "7066077"
'  TaxTrans.Description = "3803"
'  Put TTHandle, 270233, TaxTrans
'
'  Get TTHandle, 250688, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 105.08
'  Put TTHandle, 250688, TaxTrans
'
'  Get TTHandle, 252879, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 97.09
'  Put TTHandle, 252879, TaxTrans
'
'  Get TTHandle, 298053, TaxTrans
'  TaxTrans.RealPin = "7066077"
'  Put TTHandle, 298053, TaxTrans
'
'  'fix for 2293 & 10689
'  Get TTHandle, 270186, TaxTrans
'  TaxTrans.BelongTo = 258071
'  TaxTrans.RealPin = "7045208"
'  TaxTrans.Description = "8500"
'  Put TTHandle, 270186, TaxTrans
'
'  Get TTHandle, 258071, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 124.03
'  Put TTHandle, 258071, TaxTrans
'
'  Get TTHandle, 251670, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 149.63
'  Put TTHandle, 251670, TaxTrans
'
'  Get TTHandle, 298126, TaxTrans
'  TaxTrans.RealPin = "7045208"
'  Put TTHandle, 298126, TaxTrans
'
'  'fix for 3084 & 16229
'  Get TTHandle, 269830, TaxTrans
'  TaxTrans.BelongTo = 262456
'  TaxTrans.RealPin = "7096776"
'  TaxTrans.Description = "12885"
'  Put TTHandle, 269830, TaxTrans
'
'  Get TTHandle, 252408, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 267.66
'  Put TTHandle, 252408, TaxTrans
'
'  Get TTHandle, 262456, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 32.4
'  Put TTHandle, 262456, TaxTrans
'
'  Get TTHandle, 298188, TaxTrans
'  TaxTrans.RealPin = "7096776"
'  Put TTHandle, 298188, TaxTrans
'
'  'fix for 4822 & 15656
'  Get TTHandle, 269827, TaxTrans
'  TaxTrans.BelongTo = 261883
'  TaxTrans.RealPin = "7021475"
'  TaxTrans.Description = "12312"
'  Put TTHandle, 269827, TaxTrans
'
'  Get TTHandle, 254053, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 48.44
'  Put TTHandle, 254053, TaxTrans
'
'  Get TTHandle, 261883, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 39
'  Put TTHandle, 261883, TaxTrans
'
'  Get TTHandle, 298183, TaxTrans
'  TaxTrans.RealPin = "7021475"
'  Put TTHandle, 298183, TaxTrans
'
'  'fix for 11573 & 2276
'  Get TTHandle, 270274, TaxTrans
'  TaxTrans.BelongTo = 251653
'  TaxTrans.RealPin = "7066478"
'  TaxTrans.Description = "2082"
'  Put TTHandle, 270274, TaxTrans
'
'  Get TTHandle, 258494, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 205.12
'  Put TTHandle, 258494, TaxTrans
'
'  Get TTHandle, 251653, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 105.16
'  Put TTHandle, 251653, TaxTrans
'
'  Get TTHandle, 298028, TaxTrans
'  TaxTrans.RealPin = "7066478"
'  Put TTHandle, 298028, TaxTrans
'
'  'fix for 11747 & 15658
'  Get TTHandle, 269829, TaxTrans
'  TaxTrans.BelongTo = 261885
'  TaxTrans.RealPin = "7021477"
'  TaxTrans.Description = "12314"
'  Put TTHandle, 269829, TaxTrans
'
'  Get TTHandle, 258657, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 211.3
'  Put TTHandle, 258657, TaxTrans
'
'  Get TTHandle, 261885, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 39#
'  Put TTHandle, 261885, TaxTrans
'
'  Get TTHandle, 298184, TaxTrans
'  TaxTrans.RealPin = "7021477"
'  Put TTHandle, 298184, TaxTrans
'
'  'fix for 12230 & 13890
'  Get TTHandle, 271310, TaxTrans
'  TaxTrans.BelongTo = 260376
'  TaxTrans.RealPin = "7042089"
'  TaxTrans.Description = "10805"
'  Put TTHandle, 271310, TaxTrans
'
'  Get TTHandle, 259103, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 163.8
'  Put TTHandle, 259103, TaxTrans
'
'  Get TTHandle, 260376, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 124.05
'  Put TTHandle, 260376, TaxTrans
'
'  Get TTHandle, 298169, TaxTrans
'  TaxTrans.RealPin = "7042089"
'  Put TTHandle, 298169, TaxTrans
'
'  'fix for 16589 & 6251
'  Get TTHandle, 269796, TaxTrans
'  TaxTrans.BelongTo = 255405
'  TaxTrans.RealPin = "7057162"
'  TaxTrans.Description = "5834"
'  Put TTHandle, 269796, TaxTrans
'
'  Get TTHandle, 262816, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 140.5
'  Put TTHandle, 262816, TaxTrans
'
'  Get TTHandle, 255405, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 83.96
'  Put TTHandle, 255405, TaxTrans
'
'  Get TTHandle, 298091, TaxTrans
'  TaxTrans.RealPin = "7057162"
'  Put TTHandle, 298091, TaxTrans
'
'  'fix for 6520 & 6596
'  Get TCHandle, 6520, TaxCust
'  TaxCust.LastTrans = 291841
'  Put TCHandle, 6520, TaxCust
'
'  Get TTHandle, 291841, TaxTrans
'  TaxTrans.CustomerRec = 6520
'  TaxTrans.CustPin = 6520
'  TaxTrans.LastTrans = 289203
'  Put TTHandle, 291841, TaxTrans
'
'  Get TTHandle, 292532, TaxTrans
'  TaxTrans.LastTrans = 289206
'  Put TTHandle, 292532, TaxTrans
  
'  Close
'  MsgBox ("Finished.")

End Sub

Private Sub cmdFixMaxtonAdTrans_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  Dim ThisDate As Integer
  Dim BelongTo As Long
  Dim AdAmt As Double
  Dim FileName$
  Dim ThisFile As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim Cnt As Long
  
'  OpenTaxCustFile TCHandle, NumOfTCRecs
'  FileName = "maxtonAdfix.txt"
'  ThisFile = FreeFile
'  Open FileName For Output As ThisFile
'  Print #ThisFile, "Customer Name ~ Customer Number ~ Ad Amount ~ Bill Rec Num"
'  ThisDate = Date2Num("4/25/2008")
  OpenTaxTransFile TTHandle, NumOfTTRecs
  
  Get TTHandle, 32342, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 387.06
  TaxTrans.Revenue.InterestPd = 20.18
  Put TTHandle, 32342, TaxTrans
'  cnt = 0
'  For x = 1 To NumOfTTRecs
'    Get TTHandle, x, TaxTrans
'    If TaxTrans.TranType = 6 Then
'      If TaxTrans.TransDate = ThisDate Then
'        TaxTrans.Amount = 0
'        AdAmt = TaxTrans.Revenue.Collection
'        TaxTrans.Revenue.Collection = 0
'        Put TTHandle, x, TaxTrans
'        Get TCHandle, TaxTrans.CustomerRec, TaxCust
'        Print #ThisFile, QPTrim$(TaxCust.CustName) + "~" + CStr(TaxCust.Acct) + "~" + FormatCurrency(AdAmt, 2);
'        BelongTo = TaxTrans.BelongTo
'        Get TTHandle, BelongTo, TaxTrans
'        If TaxCust.Acct = 2157 Then
'         TaxTrans.Revenue.Collection = 0
'        Else
'         TaxTrans.Revenue.Collection = OldRound(TaxTrans.Revenue.Collection - AdAmt)
'        End If
'        Print #ThisFile, "~" + CStr(BelongTo)
'        Put TTHandle, BelongTo, TaxTrans
'        cnt = cnt + 1
'      End If
'    End If
'  Next x
  
  Close
  
'  MsgBox ("A total of " + CStr(cnt) + " ad charges were removed. The details are found in 'maxtonAdfix'.txt")
     
End Sub

Private Sub cmdFixNorthwest_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Get TTHandle, 30, TaxTrans
  TaxTrans.Description = "Bill # 9999"
  Put TTHandle, 30, TaxTrans
  
  Close
  MsgBox ("Done.")
End Sub

Private Sub cmdFixRealPropCustPin_Click()
  Dim RealPropRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim x As Long
  Dim TaxCustRec As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTaxCusts As Long
  Dim ThisRec As Long
  Dim Cnt As Long
  
  frmTaxShowPctComp.Label1 = "Fixing Real Prop Cust Pins"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  OpenRealPropFile RHandle, NumOfRealRecs
  OpenTaxCustFile TCHandle, NumOfTaxCusts
  For x = 1 To NumOfTaxCusts
    Get TCHandle, x, TaxCustRec
    If TaxCustRec.FirstPropRec > 0 Then ThisRec = TaxCustRec.FirstPropRec
    Do While ThisRec > 0
      Get RHandle, ThisRec, RealPropRec
      If RealPropRec.CustPin <> TaxCustRec.Acct Then
        RealPropRec.CustPin = TaxCustRec.Acct
        Put RHandle, ThisRec, RealPropRec
        Cnt = Cnt + 1
      End If
      ThisRec = RealPropRec.NextRec
    Loop
    frmTaxShowPctComp.ShowPctComp x, NumOfTaxCusts
  Next x
  Close RHandle
  Close TCHandle
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  
  Call Savemsg(900, "A total of " + CStr(Cnt) + " real property records were corrected successfully.")

End Sub

Private Sub cmdFixSparta_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  Dim Cnt As Integer
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    If TaxTrans.TransDate = Date2Num("07/17/2007") And TaxTrans.TranType = 1 Then
      TaxTrans.Amount = 0
      TaxTrans.Revenue.Principle1 = 0
      TaxTrans.Revenue.LateList = 0
      TaxTrans.Revenue.Collection = 0
      Put TTHandle, x, TaxTrans
      Cnt = Cnt + 1
    End If
  Next x
  
  Close
  MsgBox ("Done. Cnt = " & CStr(Cnt) & ".")
End Sub

Private Sub cmdFixSunset_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  'fix for #4177
  Get TTHandle, 8325, TaxTrans
  TaxTrans.Revenue.Interest = TaxTrans.Revenue.Interest + 1.39
  Put TTHandle, 8325, TaxTrans
  
  'fix for #4183
  Get TTHandle, 11767, TaxTrans
  TaxTrans.TransDate = Date2Num("04/13/1999")
  Put TTHandle, 11767, TaxTrans
  
  Get TTHandle, 11696, TaxTrans
  TaxTrans.TransDate = Date2Num("04/13/1999")
  Put TTHandle, 11696, TaxTrans
  
  Get TTHandle, 20491, TaxTrans
  TaxTrans.Amount = TaxTrans.Amount + 1538.94
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd + 1538.94
  TaxTrans.Description = "SS added 1538.94 so bal = 0"
  Put TTHandle, 20491, TaxTrans
 
  Close
  
  MsgBox ("Done.")
  
End Sub

Private Sub cmdFixRealTransIEAds_Click()
  Call FixRealTransHistIEAds
End Sub
Private Sub FixRealTransHistIEAds()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim Cnt As Integer
  Dim NextRec As Long
  Dim ThisDate As Integer
  Dim PropCnt As Integer
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  Dim RealCnt As Integer
  Dim vCnt As Integer
  Dim x As Long
  Dim AHandle As Integer
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenRealPropFile RHandle, NumOfRRecs
  ThisDate = Date2Num("04/05/2007")
  
  AHandle = FreeFile
  Open "adtransrealupdate.txt" For Output As AHandle
  Print #AHandle, "Cust Name" + "~" + "Cust Pin" + "~" + "Real Pin" + "~" + "Amount" + "~" + "Updated?"
  
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    If TaxTrans.TransDate = ThisDate And TaxTrans.TranType = 6 Then 'ad trans on specified date
      Get TCHandle, TaxTrans.CustomerRec, TaxCust 'pull affected customer
      NextRec = TaxCust.FirstPropRec
      Cnt = 0
      Do While NextRec > 0 'if they have real property (which they should unless they
      'have since sold it)
        Get RHandle, NextRec, RealRec 'if they have just one prop then we are reasonable
        'certain this is the one we want...more than one property is problematic
        Cnt = Cnt + 1
        NextRec = RealRec.NextRec
      Loop
      If Cnt = 1 Then 'makes the next code as valid as possible
        Get RHandle, TaxCust.FirstPropRec, RealRec
        TaxTrans.RealPin = RealRec.RealPin
        Put TTHandle, x, TaxTrans
        Print #AHandle, QPTrim$(TaxCust.CustName) + "~" + CStr(TaxCust.Acct) + "~" + QPTrim$(RealRec.RealPin) + "~" + Using$("###.##", TaxTrans.Amount) + "~" + "Yes"
        vCnt = vCnt + 1
      Else
        RealCnt = RealCnt + 1
        Print #AHandle, QPTrim$(TaxCust.CustName) + "~" + CStr(TaxCust.Acct) + "~" + QPTrim$(RealRec.RealPin) + "~" + Using$("###.##", TaxTrans.Amount) + "~" + "No"
      End If
    End If
  Next x
  
  Close
  MsgBox ("For advertising transactions on 4/5/07 a total of " + CStr(vCnt) + " real properties were updated. A total of " + CStr(RealCnt) + " that could be affected were not changed.")
  MsgBox ("Look for a file in this directory named 'adtransrealupdate.txt' for details delimited by a '~'.")

End Sub
Private Sub cmdFixSprucePine_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTaxCusts As Long
  
'  Call PatchSprucePine
'  OpenTaxCustFile TCHandle, NumOfTaxCusts
  OpenTaxTransFile TTHandle, NumOfTTRecs
'fix for #79 on 2/6/09
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Get TTHandle, 6438, TaxTrans
  TaxTrans.BelongTo = 0
  TaxTrans.Amount = 4.3
  TaxTrans.BillType = "C"
  TaxTrans.TransDate = Date2Num("09/05/2005")
  TaxTrans.CustomerRec = 79
  TaxTrans.TranType = 1
  TaxTrans.Revenue.Principle1 = 4.3
  TaxTrans.Revenue.Interest = 0.99
  TaxTrans.TaxYear = 2005
  TaxTrans.LastTrans = 6197
  TaxTrans.CustPin = 79
  TaxTrans.Description = "Tax Bill #82"
  TaxTrans.DiscAmt = 0
  TaxTrans.DiscXDate = Date2Num("12/31/1979")
  TaxTrans.FromPrePay = 0
  TaxTrans.InternalPin = 79
  TaxTrans.OperNum = 0
  TaxTrans.PersPin = ""
  TaxTrans.Posted2GL = "N"
  TaxTrans.RealPin = ""
  TaxTrans.TShpPara = ""
  TaxTrans.Revenue.Collection = 0
  TaxTrans.Revenue.CollectionPd = 0
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.LateList = 0
  TaxTrans.Revenue.LateListPd = 0
  TaxTrans.Revenue.Penalty = 0
  TaxTrans.Revenue.PenaltyPd = 0
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.Principle2 = 0
  TaxTrans.Revenue.Principle2Pd = 0
  TaxTrans.Revenue.Principle3 = 0
  TaxTrans.Revenue.Principle3Pd = 0
  TaxTrans.Revenue.Principle4 = 0
  TaxTrans.Revenue.Principle4Pd = 0
  TaxTrans.Revenue.Principle5 = 0
  TaxTrans.Revenue.Principle5Pd = 0
  TaxTrans.Revenue.RevOpt1 = 0
  TaxTrans.Revenue.RevOpt1Pd = 0
  TaxTrans.Revenue.RevOpt2 = 0
  TaxTrans.Revenue.RevOpt2Pd = 0
  TaxTrans.Revenue.RevOpt3 = 0
  TaxTrans.Revenue.RevOpt3Pd = 0
  Put TTHandle, 6438, TaxTrans
  
  Get TTHandle, 8163, TaxTrans
  TaxTrans.LastTrans = 6438
  Put TTHandle, 8163, TaxTrans
  
  'fix for 1474 and 22 on 2/4/09
'  Get TTHandle, 26286, TaxTrans
'  TaxTrans.TransDate = Date2Num("01/07/2009")
'  Put TTHandle, 26286, TaxTrans
'
'  Get TTHandle, 26285, TaxTrans
'  TaxTrans.TransDate = Date2Num("01/07/2009")
'  Put TTHandle, 26285, TaxTrans
  
  'cust #74
'  Get TTHandle, 24432, TaxTrans
'  TaxTrans.BelongTo = 22846
'  TaxTrans.TaxYear = 2008
'  TaxTrans.Description = "Bill #20"
'  Put TTHandle, 24432, TaxTrans
'
'  Get TTHandle, 22846, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 183.29
'  Put TTHandle, 22846, TaxTrans
'
'  Get TTHandle, 6435, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 388.66
'  Put TTHandle, 6435, TaxTrans
'
'  'cust #72
'  Get TTHandle, 6433, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 315.84
'  Put TTHandle, 6433, TaxTrans
'  'cust #73
'  Get TTHandle, 6434, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 27.37
'  Put TTHandle, 6434, TaxTrans
'
'  'cust #75
'  Get TTHandle, 6436, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 9164.26
'  Put TTHandle, 6436, TaxTrans
'  'cust #76
'  Get TTHandle, 6437, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 262.74
'  Put TTHandle, 6437, TaxTrans
'
  
'  'fix for cust #71
'  Get TTHandle, 6432, TaxTrans
'  TaxTrans.CustomerRec = 71
'  TaxTrans.CustPin = 71
'  TaxTrans.BelongTo = 0
'  TaxTrans.Description = "Bill #76"
'  TaxTrans.InternalPin = 0
'  TaxTrans.LastTrans = 3224
'  TaxTrans.RealPin = ""
'  TaxTrans.PersPin = ""
'  TaxTrans.DiscAmt = 0
'  TaxTrans.TaxYear = 2005
'  Put TTHandle, 6432, TaxTrans
'  Get TTHandle, 9253, TaxTrans
'  TaxTrans.LastTrans = 6432
'  TaxTrans.TaxYear = 2005
'  TaxTrans.BelongTo = 6432
'  Put TTHandle, 9253, TaxTrans
'  Get TTHandle, 3224, TaxTrans
'  TaxTrans.TaxYear = 2004
'  Put TTHandle, 3224, TaxTrans
  
  'fix for 928
'  Get TTHandle, 21461, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.LastTrans = 14359
'  Put TTHandle, 21461, TaxTrans
'
'  Get TTHandle, 16161, TaxTrans
'  TaxTrans.BelongTo = 13145
'  TaxTrans.CustomerRec = 1118
'  TaxTrans.CustPin = 1118
'  TaxTrans.LastTrans = 15982
'  Put TTHandle, 16161, TaxTrans
'
'  'fix for 1118
'  Get TTHandle, 18230, TaxTrans
'  TaxTrans.LastTrans = 16161
'  Put TTHandle, 18230, TaxTrans
'
'  Get TTHandle, 13145, TaxTrans
'  TaxTrans.Revenue.Principle1 = 57.62
'  TaxTrans.Revenue.Principle1Pd = 57.62
'  TaxTrans.Revenue.Interest = 1.58
'  TaxTrans.Revenue.InterestPd = 1.58
'  Put TTHandle, 13145, TaxTrans
'
'  Get TTHandle, 21463, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Interest = 0
'  TaxTrans.Revenue.Principle1 = 0
'  Put TTHandle, 21463, TaxTrans
'
'  Get TTHandle, 21462, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.InterestPd = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  Put TTHandle, 21462, TaxTrans
  Close

  MsgBox ("Finished.")

End Sub

Private Sub cmdFixSugarMtn_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  'fix for 82
  Get TTHandle, 51242, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.Principle1Pd = 0
  Put TTHandle, 51242, TaxTrans
  
  Get TTHandle, 57152, TaxTrans
  TaxTrans.Amount = 274.47
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.Principle1Pd = 274.47
  Put TTHandle, 57152, TaxTrans
  
  'fix for 665
  Get TTHandle, 51823, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.Principle1Pd = 0
  Put TTHandle, 51823, TaxTrans
  
  Get TTHandle, 53322, TaxTrans
  TaxTrans.Amount = 115.5
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.Principle1Pd = 115.5
  Put TTHandle, 53322, TaxTrans
  
  'fix for 673
  Get TTHandle, 51832, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.Principle1Pd = 0
  Put TTHandle, 51832, TaxTrans
  
  Get TTHandle, 54640, TaxTrans
  TaxTrans.Amount = 458.37
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.Principle1Pd = 458.37
  Put TTHandle, 54640, TaxTrans
  
  Get TTHandle, 14922, TaxTrans
  TaxTrans.Amount = 2.42
  TaxTrans.Revenue.InterestPd = 0.38
  Put TTHandle, 14922, TaxTrans
  
  'fix for 1376
  Get TTHandle, 52532, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.Principle1Pd = 0
  Put TTHandle, 52532, TaxTrans
  
  Get TTHandle, 53469, TaxTrans
  TaxTrans.Amount = 541.2
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.Principle1Pd = 541.2
  Put TTHandle, 53469, TaxTrans
  
 
'
'  Get TTHandle, 5262, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Interest = 0
'  Put TTHandle, 5262, TaxTrans
'
'  Get TTHandle, 4012, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Interest = 0
'  Put TTHandle, 4012, TaxTrans
'
'  Get TTHandle, 3702, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Interest = 0
'  Put TTHandle, 3702, TaxTrans
  
'  'fix for #774
'  OpenTaxTransFile TTHandle, NumOfTTRecs
'  Get TTHandle, 58486, TaxTrans
'  TaxTrans.Revenue.Principle1 = 554.84
'  Put TTHandle, 58486, TaxTrans
'
'  'fix for #1379
'  OpenTaxTransFile TTHandle, NumOfTTRecs
'  Get TTHandle, 58419, TaxTrans
'  TaxTrans.Revenue.Principle1 = 538.69
'  Put TTHandle, 58419, TaxTrans
  
  Close
  MsgBox ("Finished.")

End Sub

Private Sub cmdFixSevenDevils_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  OpenTaxTransFile TTHandle, NumOfTTRecs
  
  'fix for 133
  Get TTHandle, 22930, TaxTrans
  TaxTrans.Amount = 85.8
  TaxTrans.Revenue.Principle1Pd = 85.8
  Put TTHandle, 22930, TaxTrans
  
  Get TTHandle, 19661, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  Put TTHandle, 19661, TaxTrans
  
  'fix for 474
'  Get TTHandle, 20022, TaxTrans
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  Put TTHandle, 20022, TaxTrans
'
'  Get TTHandle, 20906, TaxTrans
'  TaxTrans.Revenue.PrePaidUsed = 6.44
'  TaxTrans.Amount = 946.56
'  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd + 6.44
'  TaxTrans.BelongTo = 20021
'  Put TTHandle, 20906, TaxTrans
  
'  Get TTHandle, 20378, TaxTrans
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  Put TTHandle, 20378, TaxTrans
'
'  Get TTHandle, 20377, TaxTrans
''  TaxTrans.Revenue.PrePaidUsed = 0
''  TaxTrans.Amount = 1386.69
'  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd ' - 0.98
'  Put TTHandle, 20377, TaxTrans
'
'  Get TTHandle, 20964, TaxTrans
'  TaxTrans.Revenue.PrePaidUsed = 0.98
'  TaxTrans.Amount = 1386.69
'  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd + 0.98
'  TaxTrans.BelongTo = 20377
'  Put TTHandle, 20964, TaxTrans
  
  Close
  MsgBox ("Finished.")
End Sub

Private Sub cmdFixSunsetBeach_Click()
 Dim TaxCust As TaxCustType
 Dim TCHandle As Integer
 Dim x As Long
 Dim NumOfTCRecs As Long
 
 OpenTaxCustFile TCHandle, NumOfTCRecs
 For x = 1 To NumOfTCRecs
   Get TCHandle, x, TaxCust
   If TaxCust.Cycle <> 1 Then
     TaxCust.Cycle = 2
     TaxCust.CycleName = "SUNSET"
     Put TCHandle, x, TaxCust
   End If
 Next x
 
 Close
 MsgBox ("Finished.")
End Sub

Private Sub cmdFixTroy_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  
  'fix for 2180 -> 872
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Get TTHandle, 88774, TaxTrans
  TaxTrans.CustomerRec = 872
  TaxTrans.CustPin = 872
  TaxTrans.LastTrans = 88773
  Put TTHandle, 88774, TaxTrans
  
  Get TTHandle, 88773, TaxTrans
  TaxTrans.CustomerRec = 872
  TaxTrans.CustPin = 872
  TaxTrans.LastTrans = 88437
  Put TTHandle, 88773, TaxTrans
  
  Get TTHandle, 94988, TaxTrans
  TaxTrans.CustomerRec = 872
  TaxTrans.CustPin = 872
  TaxTrans.LastTrans = 88774
  Put TTHandle, 94988, TaxTrans
  
  Get TTHandle, 89402, TaxTrans
  TaxTrans.LastTrans = 88436
  Put TTHandle, 89402, TaxTrans
  
  Close
  MsgBox ("Completed successfully.")
  
End Sub

Private Sub cmdFixWarsawNC_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  
  'fix for 2116's property history
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Get TTHandle, 56282, TaxTrans
  TaxTrans.BelongTo = TaxTrans.BelongTo
  TaxTrans.CustomerRec = TaxTrans.CustomerRec
  TaxTrans.PersPin = "47"
  TaxTrans.RealPin = ""
  Put TTHandle, 56282, TaxTrans
  
  Get TTHandle, 56283, TaxTrans
  TaxTrans.BelongTo = TaxTrans.BelongTo
  TaxTrans.CustomerRec = TaxTrans.CustomerRec
  TaxTrans.PersPin = ""
  TaxTrans.RealPin = "1831"
  Put TTHandle, 56283, TaxTrans

  Get TTHandle, 56284, TaxTrans
  TaxTrans.BelongTo = TaxTrans.BelongTo
  TaxTrans.CustomerRec = TaxTrans.CustomerRec
  TaxTrans.PersPin = "151"
  TaxTrans.RealPin = ""
  Put TTHandle, 56284, TaxTrans

  Get TTHandle, 56285, TaxTrans
  TaxTrans.BelongTo = TaxTrans.BelongTo
  TaxTrans.CustomerRec = TaxTrans.CustomerRec
  TaxTrans.PersPin = ""
  TaxTrans.RealPin = "36"
  Put TTHandle, 56285, TaxTrans

  Get TTHandle, 56286, TaxTrans
  TaxTrans.BelongTo = TaxTrans.BelongTo
  TaxTrans.CustomerRec = TaxTrans.CustomerRec
  TaxTrans.PersPin = ""
  TaxTrans.RealPin = "2305-1"
  Put TTHandle, 56286, TaxTrans

  Get TTHandle, 56287, TaxTrans
  TaxTrans.BelongTo = TaxTrans.BelongTo
  TaxTrans.CustomerRec = TaxTrans.CustomerRec
  TaxTrans.PersPin = ""
  TaxTrans.RealPin = "41"
  Put TTHandle, 56287, TaxTrans

  Get TTHandle, 56288, TaxTrans
  TaxTrans.BelongTo = TaxTrans.BelongTo
  TaxTrans.CustomerRec = TaxTrans.CustomerRec
  TaxTrans.PersPin = ""
  TaxTrans.RealPin = ""
  Put TTHandle, 56288, TaxTrans

  Get TTHandle, 56289, TaxTrans
  TaxTrans.BelongTo = TaxTrans.BelongTo
  TaxTrans.CustomerRec = TaxTrans.CustomerRec
  TaxTrans.PersPin = ""
  TaxTrans.RealPin = ""
  Put TTHandle, 56289, TaxTrans

  Get TTHandle, 71617, TaxTrans
  TaxTrans.BelongTo = TaxTrans.BelongTo
  TaxTrans.CustomerRec = TaxTrans.CustomerRec
  TaxTrans.PersPin = ""
  TaxTrans.RealPin = ""
  Put TTHandle, 71617, TaxTrans
  
  'fix for 2220's property hisroty
  Get TTHandle, 64879, TaxTrans
  TaxTrans.BelongTo = TaxTrans.BelongTo
  TaxTrans.CustomerRec = TaxTrans.CustomerRec
  TaxTrans.PersPin = ""
  TaxTrans.RealPin = "2042"
  Put TTHandle, 64879, TaxTrans

  Close
  MsgBox ("All done.")


End Sub

Private Sub cmdFixWhiteLakeOld_Click()
  Dim OldTaxTrans As TaxTransactionType
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  Dim TaxCust As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim NextRec As Long
  
'  Call UpdateTransWithNewPins
  OpenTaxTransFile TTHandle, NumOfTTRecs
  'fix for 2376
  ClearTrans (92117)
 
  '8/25/09 fix for #618
'  Get TTHandle, 129424, TaxTrans
'  TaxTrans.Revenue.PrePaidUsed = 289.46
'  Put TTHandle, 129424, TaxTrans
'  'fix for #5329
'  Get TTHandle, 129796, TaxTrans
'  TaxTrans.Revenue.PrePaidUsed = 18.52
'  Put TTHandle, 129796, TaxTrans
'  'fix for #5829
'  Get TTHandle, 129260, TaxTrans
'  TaxTrans.Revenue.PrePaidUsed = 9.59
'  Put TTHandle, 129260, TaxTrans
  
  
  '7/13/09 #691
'  Get TTHandle, 105002, TaxTrans
'    TaxTrans.Amount = 0
'    TaxTrans.Revenue.Collection = 0
'    TaxTrans.Revenue.CollectionPd = 0
'    TaxTrans.Revenue.Interest = 0
'    TaxTrans.Revenue.InterestPd = 0
'    TaxTrans.Revenue.LateList = 0
'    TaxTrans.Revenue.LateListPd = 0
'    TaxTrans.Revenue.Penalty = 0
'    TaxTrans.Revenue.PenaltyPd = 0
'    TaxTrans.Revenue.PrePaidUsed = 0
'    TaxTrans.Revenue.Principle1 = 0
'    TaxTrans.Revenue.Principle1Pd = 0
'    TaxTrans.Revenue.Principle2 = 0
'    TaxTrans.Revenue.Principle2Pd = 0
'    TaxTrans.Revenue.Principle3 = 0
'    TaxTrans.Revenue.Principle3Pd = 0
'    TaxTrans.Revenue.Principle4 = 0
'    TaxTrans.Revenue.Principle4Pd = 0
'    TaxTrans.Revenue.Principle5 = 0
'    TaxTrans.Revenue.Principle5Pd = 0
'    TaxTrans.Revenue.RevOpt1 = 0
'    TaxTrans.Revenue.RevOpt1Pd = 0
'    TaxTrans.Revenue.RevOpt2 = 0
'    TaxTrans.Revenue.RevOpt2Pd = 0
'    TaxTrans.Revenue.RevOpt3 = 0
'    TaxTrans.Revenue.RevOpt3Pd = 0
'  Put TTHandle, 105002, TaxTrans
'
'  Get TTHandle, 115861, TaxTrans
'    TaxTrans.Amount = 0
'    TaxTrans.Revenue.Collection = 0
'    TaxTrans.Revenue.CollectionPd = 0
'    TaxTrans.Revenue.Interest = 0
'    TaxTrans.Revenue.InterestPd = 0
'    TaxTrans.Revenue.LateList = 0
'    TaxTrans.Revenue.LateListPd = 0
'    TaxTrans.Revenue.Penalty = 0
'    TaxTrans.Revenue.PenaltyPd = 0
'    TaxTrans.Revenue.PrePaidUsed = 0
'    TaxTrans.Revenue.Principle1 = 0
'    TaxTrans.Revenue.Principle1Pd = 0
'    TaxTrans.Revenue.Principle2 = 0
'    TaxTrans.Revenue.Principle2Pd = 0
'    TaxTrans.Revenue.Principle3 = 0
'    TaxTrans.Revenue.Principle3Pd = 0
'    TaxTrans.Revenue.Principle4 = 0
'    TaxTrans.Revenue.Principle4Pd = 0
'    TaxTrans.Revenue.Principle5 = 0
'    TaxTrans.Revenue.Principle5Pd = 0
'    TaxTrans.Revenue.RevOpt1 = 0
'    TaxTrans.Revenue.RevOpt1Pd = 0
'    TaxTrans.Revenue.RevOpt2 = 0
'    TaxTrans.Revenue.RevOpt2Pd = 0
'    TaxTrans.Revenue.RevOpt3 = 0
'    TaxTrans.Revenue.RevOpt3Pd = 0
'  Put TTHandle, 115861, TaxTrans
  
'  OpenTaxCustFile CHandle, NumOfCRecs
'  'fix for cust #382
'  Get CHandle, 382, TaxCust
'  TaxCust.LastTrans = 103888
'  Put CHandle, 382, TaxCust
'
'  'fix for cust #4408
'  Get TTHandle, 105115, TaxTrans
'  TaxTrans.LastTrans = 104951
'  Put TTHandle, 105115, TaxTrans
'
'  Get TTHandle, 104951, TaxTrans
'  TaxTrans.CustomerRec = 4408
'  TaxTrans.LastTrans = 104015
'  Put TTHandle, 104951, TaxTrans
  
'  'fix for 1398
'  Get TTHandle, 99120, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  Put TTHandle, 99120, TaxTrans
'
'  Get TTHandle, 99102, TaxTrans
'  TaxTrans.Amount = 7.82
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  Put TTHandle, 99102, TaxTrans
'
'  Get TTHandle, 71590, TaxTrans
'  TaxTrans.Amount = 7.48
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.FromPrePay = 0
'  Put TTHandle, 71590, TaxTrans
 
'  For x = 1 To NumOfTTRecs
'    Get TTHandle, x, TaxTrans
'    If TaxTrans.TransDate = Date2Num("10/11/2007") Then
'      If TaxTrans.TranType = 3 Then
'        TaxTrans.Posted2GL = "N"
'        Put TTHandle, x, TaxTrans
'      End If
'    End If
'  Next x
  
'  'fix for 978
'  Get TTHandle, 98125, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  Put TTHandle, 98125, TaxTrans
'
'  Get TTHandle, 98128, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  Put TTHandle, 98128, TaxTrans
'
'  Get TTHandle, 97038, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.LateListPd = 0
'  TaxTrans.DiscAmt = 0
'  Put TTHandle, 97038, TaxTrans
'
'  Get TTHandle, 94716, TaxTrans
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.LateList = 0
'  TaxTrans.Revenue.LateListPd = 0
'  TaxTrans.Amount = 0
'  TaxTrans.DiscAmt = 0
'  Put TTHandle, 94716, TaxTrans
'

'  Get TTHandle, 94716, TaxTrans
'  TaxTrans.DiscAmt = 0.67
'  TaxTrans.Revenue.Principle1 = TaxTrans.Revenue.Principle1
'  TaxTrans.Revenue.Principle1Pd = 29.67
'  TaxTrans.Revenue.LateList = TaxTrans.Revenue.LateList
'  TaxTrans.Amount = 33.4
'  TaxTrans.DiscAmt = 0.67
'  Put TTHandle, 94716, TaxTrans
  
  
  
'  'fix for 3461
'  Get TTHandle, 75446, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 14.03
'  TaxTrans.Revenue.Interest = 0.83
'  Put TTHandle, 75446, TaxTrans
'
'  Get TTHandle, 92608, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Interest = 0#
'  Put TTHandle, 92608, TaxTrans
'
'  Get TTHandle, 91839, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Interest = 0#
'  Put TTHandle, 91839, TaxTrans
'
'  Get TTHandle, 81341, TaxTrans
'  TaxTrans.Revenue.Interest = 0#
'  TaxTrans.Amount = 1.02
'  TaxTrans.TransDate = Date2Num%("08/04/2006")
'  TaxTrans.TranType = 9
'  TaxTrans.Revenue.Principle1Pd = 1.02
'  TaxTrans.Revenue.InterestPd = 0
'  TaxTrans.Revenue.CollectionPd = 0
'  TaxTrans.Revenue.LateListPd = 0
'  TaxTrans.Revenue.RevOpt1Pd = 0
'  TaxTrans.Revenue.RevOpt2Pd = 0
'  TaxTrans.Revenue.RevOpt3Pd = 0
'  TaxTrans.CustPin = 3461
'  TaxTrans.DiscXDate = 0
'  TaxTrans.RealPin = ""
'  TaxTrans.PersPin = ""
'  TaxTrans.Posted2GL = "N"
'  TaxTrans.TaxYear = 2006
'  TaxTrans.DiscAmt = 0
'  TaxTrans.OperNum = OperNum
'  TaxTrans.Amount = 0
'  TaxTrans.FromPrePay = 1.02
'  TaxTrans.Description = "Credit Applied to Bill# " + Str$(383)
'  TaxTrans.CustomerRec = 3461
'  TaxTrans.BelongTo = 75446
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidUsed = 1.02
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.InternalPin = 3461
'  TaxTrans.CntyPara = ""
'  TaxTrans.CyclPara = ""
'  TaxTrans.TShpPara = ""
'  Put TTHandle, 81341, TaxTrans
  
  
  
  'fix for #2376
'  Get TTHandle, 76839, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 14.18
'  TaxTrans.Revenue.Interest = 0
'  Put TTHandle, 76839, TaxTrans
'
'  Get TTHandle, 78736, TaxTrans
'  TaxTrans.BelongTo = 76839
'  TaxTrans.Description = "1556"
'  TaxTrans.CustomerRec = 2376
'  Put TTHandle, 78736, TaxTrans
'
'  Get TTHandle, 86060, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Interest = 0
'  Put TTHandle, 86060, TaxTrans
'
'  Get TTHandle, 87075, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Interest = 0
'  Put TTHandle, 87075, TaxTrans
'
'  Get TTHandle, 88300, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Interest = 0
'  Put TTHandle, 88300, TaxTrans
'
'  Get TTHandle, 89442, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Interest = 0
'  Put TTHandle, 89442, TaxTrans
'
'  Get TTHandle, 90313, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Interest = 0
'  Put TTHandle, 90313, TaxTrans
'
'  Get TTHandle, 91252, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Interest = 0
'  Put TTHandle, 91252, TaxTrans
'  Close
'
'  'fix for #3252
'  OpenTaxCustFile CHandle, NumOfCRecs
'  Get CHandle, 3252, TaxCust
'  NextRec = TaxCust.LastTrans
'  TaxTrans.TranType = 9
'  TaxTrans.TransDate = Date2Num("8/04/2006")
'  TaxTrans.Revenue.Principle1Pd = 1.35
'  TaxTrans.Amount = 0
'  TaxTrans.CustPin = TaxCust.PIN
'  TaxTrans.BelongTo = 78263
'  TaxTrans.CustomerRec = 3252
'  TaxTrans.TaxYear = 2006
'  TaxTrans.Description = "Credit Applied to Bill #2761"
'  TaxTrans.FromPrePay = 1.35
'  TaxTrans.LastTrans = NumOfTTRecs + 1
'  TaxTrans.LastTrans = NextRec
'  TaxCust.LastTrans = NumOfTTRecs + 1
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidUsed = 1.35
'  TaxTrans.Revenue.PrePaidBal = 0
'  OpenTaxTransFile TTHandle, NumOfTTRecs
'  Get TTHandle, 78263, OldTaxTrans
'  OldTaxTrans.Revenue.Principle1Pd = OldTaxTrans.Revenue.Principle1Pd + 1.35
'  Put TTHandle, 78263, OldTaxTrans
'  Close TTHandle
'  TaxTrans.Posted2GL = OldTaxTrans.Posted2GL
'  TaxTrans.DiscXDate = OldTaxTrans.DiscXDate
'  TaxTrans.InternalPin = OldTaxTrans.InternalPin
'  TaxTrans.OperNum = OldTaxTrans.OperNum
'  TaxTrans.RealPin = QPTrim$(OldTaxTrans.RealPin)
'  TaxTrans.PersPin = QPTrim$(OldTaxTrans.PersPin)
'  Put CHandle, 3252, TaxCust
'  OpenTaxTransFile TTHandle, NumOfTTRecs
'  Put TTHandle, NumOfTTRecs + 1, TaxTrans
  
  Close
  MsgBox ("Finished.")
  
'  Call FixWhiteLake
'  Call FixWhiteLakeOverPay
End Sub

Private Sub cmdFixWJefferson_Click()
  Call FixWestJeffersonPers
End Sub

Private Sub cmdFixTrentwood_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxCust As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim x As Long
  Dim LastTrans As Long
  
  
  OpenTaxCustFile CHandle, NumOfCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  'fix for 1960 7/27/09
  Get CHandle, 1960, TaxCust
  TaxCust.LastTrans = 36430
  Put CHandle, 1960, TaxCust
  
  Get TTHandle, 36430, TaxTrans
  TaxTrans.CustomerRec = 1960
  TaxTrans.CustPin = 1960
  TaxTrans.LastTrans = 35047
  Put TTHandle, 36430, TaxTrans
  
  'fix for #2516 7/27/09
  Get CHandle, 2516, TaxCust
  TaxCust.LastTrans = 9823
  Put CHandle, 2516, TaxCust
  
  
  
'  Get TTHandle, 23306, TaxTrans
'   TaxTrans.Amount = 0
'   TaxTrans.Revenue.Principle1 = 0
'   TaxTrans.Revenue.Principle1Pd = 0
'   TaxTrans.BelongTo = 0
'   TaxTrans.CustomerRec = 363
'   TaxTrans.CustPin = 363
'  Put TTHandle, 23306, TaxTrans
  
  'fix 606
'  Get CHandle, 606, TaxCust
'    LastTrans = TaxCust.LastTrans
'    TaxTrans.Amount = 261.06
'    TaxTrans.Revenue.Principle1Pd = 261.06
'    TaxTrans.BelongTo = 23034
'    TaxTrans.CustomerRec = 606
'    TaxTrans.CustPin = 606
'    TaxTrans.Description = "Bill #584"
'    TaxTrans.DiscAmt = 0
'    TaxTrans.FromPrePay = 0
'    TaxTrans.LastTrans = LastTrans
'    TaxTrans.TaxYear = 2006
'    TaxTrans.TransDate = Date2Num("12/14/2006")
'    TaxTrans.TranType = 2
'    Put TTHandle, NumOfTTRecs + 1, TaxTrans
'    Get TTHandle, 23034, TaxTrans
'    TaxTrans.Revenue.Principle1Pd = 261.06
'    Put TTHandle, 23034, TaxTrans
'    TaxCust.LastTrans = NumOfTTRecs + 1
'  Put CHandle, 606, TaxCust
'
'  'fix 1070
'   Get TTHandle, 26639, TaxTrans
'    TaxTrans.Amount = 0
'    TaxTrans.Revenue.Principle1Pd = 0
'    TaxTrans.BelongTo = 0
'    TaxTrans.Revenue.PrePaidAmt = 0
'    TaxTrans.Revenue.PrePaidBal = 0
'    TaxTrans.Revenue.PrePaidUsed = 0
'    TaxTrans.Description = "Payment fixed."
'   Put TTHandle, 26639, TaxTrans
'
'  'fix 971
'  Get CHandle, 971, TaxCust
'    LastTrans = TaxCust.LastTrans
'    TaxTrans.Amount = 896.86
'    TaxTrans.Revenue.Principle1Pd = 896.06
'    TaxTrans.BelongTo = 23198
'    TaxTrans.CustomerRec = 971
'    TaxTrans.CustPin = 971
'    TaxTrans.Description = "Bill #748"
'    TaxTrans.DiscAmt = 0
'    TaxTrans.FromPrePay = 0
'    TaxTrans.LastTrans = LastTrans
'    TaxTrans.TaxYear = 2006
'    TaxTrans.TransDate = Date2Num("12/27/2006")
'    TaxTrans.TranType = 2
'    Put TTHandle, NumOfTTRecs + 2, TaxTrans
'    Get TTHandle, 23198, TaxTrans
'    TaxTrans.Revenue.Principle1Pd = 896.06
'    Put TTHandle, 23198, TaxTrans
'    TaxCust.LastTrans = NumOfTTRecs + 2
'  Put CHandle, 971, TaxCust
'
'  'fix 701
'   Get TTHandle, 26947, TaxTrans
'    TaxTrans.Amount = 0
'    TaxTrans.Revenue.Principle1Pd = 0
'    TaxTrans.BelongTo = 0
'    TaxTrans.Revenue.PrePaidAmt = 0
'    TaxTrans.Revenue.PrePaidBal = 0
'    TaxTrans.Revenue.PrePaidUsed = 0
'    TaxTrans.Description = "Payment fixed."
'   Put TTHandle, 26947, TaxTrans
'
'  'fix 2134
'  Get CHandle, 2134, TaxCust
'    LastTrans = TaxCust.LastTrans
'    TaxTrans.Amount = 460.86
'    TaxTrans.Revenue.Principle1Pd = 460.86
'    TaxTrans.BelongTo = 24237
'    TaxTrans.CustomerRec = 2134
'    TaxTrans.CustPin = 2134
'    TaxTrans.Description = "Bill #1787"
'    TaxTrans.DiscAmt = 0
'    TaxTrans.FromPrePay = 0
'    TaxTrans.LastTrans = LastTrans
'    TaxTrans.TaxYear = 2006
'    TaxTrans.TransDate = Date2Num("12/27/2006")
'    TaxTrans.TranType = 2
'    Put TTHandle, NumOfTTRecs + 3, TaxTrans
'    Get TTHandle, 24237, TaxTrans
'    TaxTrans.Revenue.Principle1Pd = 460.86
'    Put TTHandle, 24237, TaxTrans
'    TaxCust.LastTrans = NumOfTTRecs + 3
'  Put CHandle, 2134, TaxCust
'
'  'fix 2341
'   Get TTHandle, 26966, TaxTrans
'    TaxTrans.Amount = 0
'    TaxTrans.Revenue.Principle1Pd = 0
'    TaxTrans.BelongTo = 0
'    TaxTrans.Revenue.PrePaidAmt = 0
'    TaxTrans.Revenue.PrePaidBal = 0
'    TaxTrans.Revenue.PrePaidUsed = 0
'    TaxTrans.Description = "Payment fixed."
'   Put TTHandle, 26966, TaxTrans
'
'  'fix 2108
'  Get CHandle, 2108, TaxCust
'    LastTrans = TaxCust.LastTrans
'    TaxTrans.Amount = 518.32
'    TaxTrans.Revenue.Principle1Pd = 518.32
'    TaxTrans.BelongTo = 24272
'    TaxTrans.CustomerRec = 2108
'    TaxTrans.CustPin = 2108
'    TaxTrans.Description = "Bill #1822"
'    TaxTrans.DiscAmt = 0
'    TaxTrans.FromPrePay = 0
'    TaxTrans.LastTrans = LastTrans
'    TaxTrans.TaxYear = 2006
'    TaxTrans.TransDate = Date2Num("12/18/2006")
'    TaxTrans.TranType = 2
'    Put TTHandle, NumOfTTRecs + 4, TaxTrans
'    Get TTHandle, 24272, TaxTrans
'    TaxTrans.Revenue.Principle1Pd = 518.32
'    Put TTHandle, 24272, TaxTrans
'    TaxCust.LastTrans = NumOfTTRecs + 4
'  Put CHandle, 2108, TaxCust
'
'  'fix 2341
'   Get TTHandle, 26753, TaxTrans
'    TaxTrans.Amount = 0
'    TaxTrans.Revenue.Principle1Pd = 0
'    TaxTrans.BelongTo = 0
'    TaxTrans.Revenue.PrePaidAmt = 0
'    TaxTrans.Revenue.PrePaidBal = 0
'    TaxTrans.Revenue.PrePaidUsed = 0
'    TaxTrans.Description = "Payment fixed."
'   Put TTHandle, 26753, TaxTrans
   
  Close
  MsgBox ("Finished.")
End Sub

Private Sub cmdFixYadkin_Click()
  Dim TransRec As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim x As Long
  
  OpenTaxTransFile THandle, NumOfTRecs
  'fix #1276
  Get THandle, 21121, TransRec
  TransRec.LastTrans = 20964
  Put THandle, 21121, TransRec
  
  'fix #2423
  Get THandle, 21340, TransRec
  TransRec.LastTrans = 21119
  Put THandle, 21340, TransRec
  
  Get THandle, 21119, TransRec
  TransRec.LastTrans = 20961
  TransRec.CustPin = 2423
  TransRec.CustomerRec = 2423
  TransRec.BelongTo = 15485
  TransRec.Description = "Bill #1022"
  Put THandle, 21119, TransRec
  
  'fix #1646
  
  Get THandle, 21341, TransRec
  TransRec.LastTrans = 21120
  Put THandle, 21341, TransRec
  
  Get THandle, 21120, TransRec
  TransRec.LastTrans = 20963
  TransRec.CustPin = 1646
  TransRec.CustomerRec = 1646
  TransRec.BelongTo = 15483
  TransRec.Description = "Bill #1020"
  Put THandle, 21120, TransRec
  
  Close
  MsgBox ("Done.")
End Sub

Private Sub cmdFixWatauga_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Integer
  Dim BelongTo As Long
  Dim Collection As Double
  Dim CollectionPd As Double
  Dim Interest As Double
  Dim InterestPd As Double
  Dim LateList As Double
  Dim LateListPd As Double
  Dim Penalty As Double
  Dim PenaltyPd As Double
  Dim PrePaidUsed As Double
  Dim Principle1 As Double
  Dim Principle1Pd As Double
  Dim Principle2 As Double
  Dim Principle2Pd As Double
  Dim Principle3 As Double
  Dim Principle3Pd As Double
  Dim Principle4 As Double
  Dim Principle4Pd As Double
  Dim Principle5 As Double
  Dim Principle5Pd As Double
  Dim RevOpt1 As Double
  Dim RevOpt1Pd As Double
  Dim RevOpt2 As Double
  Dim RevOpt2Pd As Double
  Dim RevOpt3 As Double
  Dim RevOpt3Pd As Double
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For x = 23664 To NumOfTTRecs
   Get TTHandle, x, TaxTrans
   If x = 23664 Then GoSub FixIt
   If x >= 23685 And x <= 23715 Then GoSub FixIt
   If x >= 23717 And x <= 23757 Then GoSub FixIt
   If x >= 23759 And x <= 23763 Then GoSub FixIt
   If x >= 23765 And x <= 23783 Then GoSub FixIt
   If x >= 24839 And x <= 24880 Then GoSub FixIt
   If x >= 24882 And x <= 24904 Then GoSub FixIt
   If x >= 24906 And x <= 24917 Then GoSub FixIt
   Select Case x
     Case 31140, 30052, 28857, 27619, 26311
       GoSub FixIt
   End Select
   
  Next x
  
  Get TTHandle, 20202, TaxTrans
    TaxTrans.Revenue.Principle1Pd = 1321.6
  Put TTHandle, 20202, TaxTrans
  Close
  MsgBox ("Completed.")
  Exit Sub
  
FixIt:
    If x = 23775 Then
      GoTo GoHere
    End If
    If x = 24908 Then
      GoTo GoHere
    End If
    If TaxTrans.BelongTo = 0 Then
      GoTo GoHere
    End If
    Collection = TaxTrans.Revenue.Collection
    CollectionPd = TaxTrans.Revenue.CollectionPd
    Interest = TaxTrans.Revenue.Interest
    InterestPd = TaxTrans.Revenue.InterestPd
    LateList = TaxTrans.Revenue.LateList
    LateListPd = TaxTrans.Revenue.LateListPd
    Penalty = TaxTrans.Revenue.Penalty
    PenaltyPd = TaxTrans.Revenue.PenaltyPd
    PrePaidUsed = TaxTrans.Revenue.PrePaidUsed
    Principle1 = TaxTrans.Revenue.Principle1
    Principle1Pd = TaxTrans.Revenue.Principle1Pd
    Principle2 = TaxTrans.Revenue.Principle2
    Principle2Pd = TaxTrans.Revenue.Principle2Pd
    Principle3 = TaxTrans.Revenue.Principle3
    Principle3Pd = TaxTrans.Revenue.Principle3Pd
    Principle4 = TaxTrans.Revenue.Principle4
    Principle4Pd = TaxTrans.Revenue.Principle4Pd
    Principle5 = TaxTrans.Revenue.Principle5
    Principle5Pd = TaxTrans.Revenue.Principle5Pd
    RevOpt1 = TaxTrans.Revenue.RevOpt1
    RevOpt1Pd = TaxTrans.Revenue.RevOpt1Pd
    RevOpt2 = TaxTrans.Revenue.RevOpt2
    RevOpt2Pd = TaxTrans.Revenue.RevOpt2Pd
    RevOpt3 = TaxTrans.Revenue.RevOpt3
    RevOpt3Pd = TaxTrans.Revenue.RevOpt3Pd
    
    BelongTo = TaxTrans.BelongTo
    
    Get TTHandle, BelongTo, TaxTrans
    TaxTrans.Revenue.Collection = TaxTrans.Revenue.Collection - Collection
    TaxTrans.Revenue.CollectionPd = TaxTrans.Revenue.CollectionPd - CollectionPd
    TaxTrans.Revenue.Interest = TaxTrans.Revenue.Interest - Interest
    TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd - InterestPd
    TaxTrans.Revenue.LateList = TaxTrans.Revenue.LateList - LateList
    TaxTrans.Revenue.LateListPd = TaxTrans.Revenue.LateListPd - LateListPd
    TaxTrans.Revenue.Penalty = TaxTrans.Revenue.Penalty - Penalty
    TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - PenaltyPd
    TaxTrans.Revenue.Principle1 = TaxTrans.Revenue.Principle1 - Principle1
    TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - Principle1Pd
    TaxTrans.Revenue.Principle2 = TaxTrans.Revenue.Principle2 - Principle2
    TaxTrans.Revenue.Principle2Pd = TaxTrans.Revenue.Principle2Pd - Principle2Pd
    TaxTrans.Revenue.Principle3 = TaxTrans.Revenue.Principle3 - Principle3
    TaxTrans.Revenue.Principle3Pd = TaxTrans.Revenue.Principle3Pd - Principle3Pd
    TaxTrans.Revenue.Principle4 = TaxTrans.Revenue.Principle4 - Principle4
    TaxTrans.Revenue.Principle4Pd = TaxTrans.Revenue.Principle4Pd - Principle4Pd
    TaxTrans.Revenue.Principle5 = TaxTrans.Revenue.Principle5 - Principle5
    TaxTrans.Revenue.Principle5Pd = TaxTrans.Revenue.Principle5Pd - Principle5Pd
    TaxTrans.Revenue.RevOpt1 = TaxTrans.Revenue.RevOpt1 - RevOpt1
    TaxTrans.Revenue.RevOpt1Pd = TaxTrans.Revenue.RevOpt1Pd - RevOpt1Pd
    TaxTrans.Revenue.RevOpt2 = TaxTrans.Revenue.RevOpt2 - RevOpt2
    TaxTrans.Revenue.RevOpt2Pd = TaxTrans.Revenue.RevOpt2Pd - RevOpt2Pd
    TaxTrans.Revenue.RevOpt3 = TaxTrans.Revenue.RevOpt3 - RevOpt3
    TaxTrans.Revenue.RevOpt3Pd = TaxTrans.Revenue.RevOpt3Pd - RevOpt3Pd
    Put TTHandle, BelongTo, TaxTrans
    Get TTHandle, x, TaxTrans
GoHere:
    
    TaxTrans.Amount = 0
    TaxTrans.Revenue.Collection = 0
    TaxTrans.Revenue.CollectionPd = 0
    TaxTrans.Revenue.Interest = 0
    TaxTrans.Revenue.InterestPd = 0
    TaxTrans.Revenue.LateList = 0
    TaxTrans.Revenue.LateListPd = 0
    TaxTrans.Revenue.Penalty = 0
    TaxTrans.Revenue.PenaltyPd = 0
    TaxTrans.Revenue.PrePaidUsed = 0
    TaxTrans.Revenue.Principle1 = 0
    TaxTrans.Revenue.Principle1Pd = 0
    TaxTrans.Revenue.Principle2 = 0
    TaxTrans.Revenue.Principle2Pd = 0
    TaxTrans.Revenue.Principle3 = 0
    TaxTrans.Revenue.Principle3Pd = 0
    TaxTrans.Revenue.Principle4 = 0
    TaxTrans.Revenue.Principle4Pd = 0
    TaxTrans.Revenue.Principle5 = 0
    TaxTrans.Revenue.Principle5Pd = 0
    TaxTrans.Revenue.RevOpt1 = 0
    TaxTrans.Revenue.RevOpt1Pd = 0
    TaxTrans.Revenue.RevOpt2 = 0
    TaxTrans.Revenue.RevOpt2Pd = 0
    TaxTrans.Revenue.RevOpt3 = 0
    TaxTrans.Revenue.RevOpt3Pd = 0
    Put TTHandle, x, TaxTrans
    Return

End Sub

Private Sub cmdFixYadkinville_Click()
'  Dim TaxTrans As TaxTransactionType
'  Dim TTHandle As Integer
'  Dim NumOfTTRecs As Long
'  Dim TaxCust As TaxCustType
'  Dim CHandle As Integer
'  Dim NumOfCRecs As Long
'  Dim NextRec As Long
'  Dim x As Long
'  Dim Acct(1 To 60) As Long
'  Dim Amt(1 To 60) As Double
'  Dim wrkbk As Workbook
'  Dim Path As String
'
''  Path = App.Path
''  Set wrkbk = GetObject(Path & "\Yadkinville.xls")
''  For x = 1 To 5
''    Acct(x) = wrkbk.Worksheets(1).Range("A" & x).Value
''    amt(x) = wrkbk.Worksheets(1).Range("D" & x).Value
''  Next x
'
''  OpenTaxCustFile CHandle, NumOfCRecs
'  OpenTaxTransFile TTHandle, NumOfTTRecs
'   'fix for 1489 1/9/09
'   Get TTHandle, 13846, TaxTrans
'   TaxTrans.Revenue.CollectionPd = 0
'   TaxTrans.Revenue.Collection = 0
'   TaxTrans.Amount = 0
'   Put TTHandle, 13846, TaxTrans
'
'
'  'fix for #411 on 1/5/09
''  Get TTHandle, 21109, TaxTrans
''  TaxTrans.Revenue.CollectionPd = 0
''  TaxTrans.Revenue.Collection = 0
''  TaxTrans.Amount = 0
''  Put TTHandle, 21109, TaxTrans
'
'  'fix 1489
''  Get TTHandle, 15314, TaxTrans
''  TaxTrans.Revenue.CollectionPd = 0
''  TaxTrans.Revenue.Collection = 0
''  Put TTHandle, 15314, TaxTrans
''
''  Get TTHandle, 25417, TaxTrans
''  TaxTrans.Revenue.CollectionPd = 0
''  TaxTrans.Revenue.Collection = 0
''  TaxTrans.Amount = TaxTrans.Amount - 1
''  Put TTHandle, 25417, TaxTrans
''
''  'fix for #575 12/12/08
''  Get TTHandle, 15664, TaxTrans
''  TaxTrans.Revenue.CollectionPd = 0
''  TaxTrans.Revenue.Collection = 0
''  Put TTHandle, 15664, TaxTrans
''
''  Get TTHandle, 24841, TaxTrans
''  TaxTrans.Revenue.CollectionPd = 0
''  TaxTrans.Revenue.Collection = 0
''  TaxTrans.Amount = TaxTrans.Amount - 1
''  Put TTHandle, 24841, TaxTrans
''
''  'fix for #427 12/12/08
''  Get TTHandle, 15355, TaxTrans
''  TaxTrans.Revenue.CollectionPd = 0
''  TaxTrans.Revenue.Collection = 0
''  Put TTHandle, 15355, TaxTrans
''
''  Get TTHandle, 21166, TaxTrans
''  TaxTrans.Revenue.CollectionPd = 0
''  TaxTrans.Revenue.Collection = 0
''  TaxTrans.Amount = TaxTrans.Amount - 1
''  Put TTHandle, 21166, TaxTrans
'
'
'
'  'fix #553
''  Get TTHandle, 22905, TaxTrans
''  TaxTrans.LastTrans = 17043
''  Put TTHandle, 22905, TaxTrans
''
''  'fix #554
''  Get TTHandle, 21353, TaxTrans
''  TaxTrans.LastTrans = 21128
''  Put TTHandle, 21353, TaxTrans
''
''  Get TTHandle, 21128, TaxTrans
''  TaxTrans.LastTrans = 20977
''  TaxTrans.CustomerRec = 554
''  TaxTrans.CustPin = 554
''  Put TTHandle, 21128, TaxTrans
'
''  'fix cust #827
''  Get TTHandle, 21103, TaxTrans
''  TaxTrans.LastTrans = 20905
''  Put TTHandle, 21103, TaxTrans
''
''  Get TTHandle, 21101, TaxTrans
''  TaxTrans.CustomerRec = 829
''  TaxTrans.CustPin = 829
''  TaxTrans.BelongTo = 15133
''  TaxTrans.LastTrans = 20903
''  Put TTHandle, 21101, TaxTrans
''
''  'fix cust 829
''  Get TTHandle, 21288, TaxTrans
''  TaxTrans.LastTrans = 21101
''  Put TTHandle, 21288, TaxTrans
'
''  For x = 1 To 5
''    Get CHandle, Acct(x), TaxCust
''    NextRec = TaxCust.LastTrans
''    Do While NextRec > 0
''      Get TTHandle, NextRec, TaxTrans
''      If TaxTrans.TransDate = Date2Num("07/05/07") And TaxTrans.TranType = 1 Then
''        TaxTrans.Revenue.Principle1 = amt(x)
''        TaxTrans.Amount = amt(x)
''        Put TTHandle, NextRec, TaxTrans
''        Exit Do
''      End If
''    Loop
''  Next x
'
'  Close TTHandle
'
'  MsgBox ("Completed successfully.")
  
End Sub

Private Sub cmdFixSylva_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim Cnt As Integer
  Dim NextRec As Long
  Dim ThisDate As Integer
  Dim PropCnt As Integer
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  Dim RealCnt As Integer
  Dim vCnt As Integer
  Dim x As Long
  Dim AHandle As Integer
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenRealPropFile RHandle, NumOfRRecs
  ThisDate = Date2Num("04/02/2007")
  
  AHandle = FreeFile
  Open "adtransrealupdate.txt" For Output As AHandle
  Print #AHandle, "Cust Name" + "~" + "Cust Pin" + "~" + "Real Pin" + "~" + "Amount" + "~" + "Updated?"
  
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    If TaxTrans.TransDate = ThisDate And TaxTrans.TranType = 6 Then 'ad trans on specified date
      Get TCHandle, TaxTrans.CustomerRec, TaxCust 'pull affected customer
      NextRec = TaxCust.FirstPropRec
      Cnt = 0
      Do While NextRec > 0 'if they have real property (which they should unless they
      'have since sold it)
        Get RHandle, NextRec, RealRec 'if they have just one prop then we are reasonable
        'certain this is the one we want...more than one property is problematic
        Cnt = Cnt + 1
        NextRec = RealRec.NextRec
      Loop
      If Cnt = 1 Then 'makes the next code as valid as possible
        Get RHandle, TaxCust.FirstPropRec, RealRec
        TaxTrans.RealPin = RealRec.RealPin
        Put TTHandle, x, TaxTrans
        Print #AHandle, QPTrim$(TaxCust.CustName) + "~" + CStr(TaxCust.Acct) + "~" + QPTrim$(RealRec.RealPin) + "~" + Using$("###.##", TaxTrans.Amount) + "~" + "Yes"
        vCnt = vCnt + 1
      Else
        RealCnt = RealCnt + 1
        Print #AHandle, QPTrim$(TaxCust.CustName) + "~" + CStr(TaxCust.Acct) + "~" + QPTrim$(RealRec.RealPin) + "~" + Using$("###.##", TaxTrans.Amount) + "~" + "No"
      End If
    End If
  Next x
  
  'fix for 1895
'  Get TCHandle, 1895, TaxCust
'  TaxCust.LastTrans = 11055
'  Put TCHandle, 1895, TaxCust
'
'  Get TTHandle, 50473, TaxTrans
'  TaxTrans.LastTrans = 49911
'  TaxTrans.CustomerRec = 127
'  Put TTHandle, 50473, TaxTrans
'
'  Get TCHandle, 127, TaxCust
'  TaxCust.LastTrans = 50473
'  Put TCHandle, 127, TaxCust
  
'  'fix for #1343
'  Get TTHandle, 38465, TaxTrans
'  TaxTrans.Amount = 528.78
'  TaxTrans.Revenue.Principle1Pd = 528.78
'  Put TTHandle, 38465, TaxTrans
  
  Close
  MsgBox ("For advertising transactions on 4/2/07 a total of " + CStr(vCnt) + " real properties were updated. A total of " + CStr(RealCnt) + " that could be affected were not changed.")
  MsgBox ("Look for a file in this directory named 'adtransrealupdate.txt' for details delimited by a '~'.")


End Sub

Private Sub cmdFixTrentwoods_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  'fix for #73
  Get TCHandle, 73, TaxCust
  TaxCust.LastTrans = 26962
  Put TCHandle, 73, TaxCust
  
  Get TTHandle, 26962, TaxTrans
  TaxTrans.LastTrans = 25683
  TaxTrans.BelongTo = 25683
  TaxTrans.CustomerRec = 73
  TaxTrans.CustPin = 73
  TaxTrans.Description = "16811"
  Put TTHandle, 26962, TaxTrans

  'fix for 2530
  Get TCHandle, 2530, TaxCust
  TaxCust.LastTrans = 8510
  Put TCHandle, 2530, TaxCust
  
  'fix for #373
  Get TCHandle, 373, TaxCust
  TaxCust.LastTrans = 32307
  Put TCHandle, 373, TaxCust
  
  Get TTHandle, 32307, TaxTrans
  TaxTrans.LastTrans = 30635
  TaxTrans.BelongTo = 30635
  TaxTrans.CustomerRec = 373
  TaxTrans.CustPin = 373
  TaxTrans.Description = "958"
  Put TTHandle, 32307, TaxTrans

  'fix for 170
  Get TCHandle, 170, TaxCust
  TaxCust.LastTrans = 19794
  Put TCHandle, 170, TaxCust
  
  'fix for #806
  Get TCHandle, 806, TaxCust
  TaxCust.LastTrans = 31353
  Put TCHandle, 806, TaxCust
  
  Get TTHandle, 31353, TaxTrans
  TaxTrans.LastTrans = 28813
  TaxTrans.BelongTo = 28813
  TaxTrans.CustomerRec = 806
  TaxTrans.CustPin = 806
  TaxTrans.Description = "693"
  Put TTHandle, 31353, TaxTrans

  'fix for 2559
  Get TCHandle, 2559, TaxCust
  TaxCust.LastTrans = 0
  Put TCHandle, 2559, TaxCust
  
  'fix for #1661
  Get TCHandle, 1661, TaxCust
  TaxCust.LastTrans = 31299
  Put TCHandle, 1661, TaxCust
  
  Get TTHandle, 31299, TaxTrans
  TaxTrans.LastTrans = 28882
  TaxTrans.BelongTo = 28882
  TaxTrans.CustomerRec = 1661
  TaxTrans.CustPin = 1661
  TaxTrans.Description = "762"
  Put TTHandle, 31299, TaxTrans

  'fix for 211
  Get TCHandle, 211, TaxCust
  TaxCust.LastTrans = 30989
  Put TCHandle, 211, TaxCust
  
  'fix for #1747
  Get TCHandle, 1747, TaxCust
  TaxCust.LastTrans = 31343
  Put TCHandle, 1747, TaxCust
  
  Get TTHandle, 31343, TaxTrans
  TaxTrans.LastTrans = 30249
  TaxTrans.BelongTo = 30249
  TaxTrans.CustomerRec = 1747
  TaxTrans.CustPin = 1747
  TaxTrans.Description = "2129"
  Put TTHandle, 31343, TaxTrans

  'fix for 241
  Get TCHandle, 241, TaxCust
  TaxCust.LastTrans = 26373
  Put TCHandle, 241, TaxCust
  
  
  
  
  Close
  MsgBox ("Finished.")

End Sub

Private Sub cmdFixSilva_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    If QPTrim$(TaxTrans.RealPin) = "7641075830" And TaxTrans.CustomerRec = 1690 Then
      TaxTrans.PersPin = "0"
      Put TTHandle, x, TaxTrans
    End If
  Next x
  Close
  MsgBox ("Done")
End Sub

Private Sub cmdFixWhiteLake_Click()
  Dim OldTaxTrans As TaxTransactionType
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  Dim TaxCust As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim NextRec As Long
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Get TTHandle, 115860, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 33.24
  TaxTrans.Revenue.InterestPd = 0
  Put TTHandle, 115860, TaxTrans
  Close
  MsgBox ("Finished.")
End Sub

Private Sub cmdInsertPrePay_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Get TTHandle, 36318, TaxTrans 'remove $1.82 from this trans
  TaxTrans.Amount = 250.52
  TaxTrans.Revenue.Principle1 = 233.35
  Put TTHandle, 36318, TaxTrans
  
  'add prepay trans
  TaxTrans.TransDate = Date2Num("07/30/2008")
  TaxTrans.TranType = 22 'overpay only
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.CollectionPd = 0
  TaxTrans.Revenue.LateListPd = 0
  TaxTrans.Revenue.RevOpt1Pd = 0
  TaxTrans.Revenue.RevOpt2Pd = 0
  TaxTrans.Revenue.RevOpt3Pd = 0
  TaxTrans.CustPin = 629
  TaxTrans.DiscXDate = Date2Num("01/07/2008")
  TaxTrans.RealPin = " "
  TaxTrans.PersPin = " "
  TaxTrans.Posted2GL = "Y"
  TaxTrans.TaxYear = 2008
  TaxTrans.DiscAmt = 0
  TaxTrans.OperNum = 0
  TaxTrans.Amount = 1.82
  TaxTrans.Description = "Prepay"
  TaxTrans.CustomerRec = 629
  TaxTrans.LastTrans = 36318
  TaxTrans.BelongTo = 0
  TaxTrans.Revenue.PrePaidAmt = 1.82
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.PrePaidBal = 1.82
  Put TTHandle, NumOfTTRecs + 1, TaxTrans
  
  Get TTHandle, 36397, TaxTrans
  TaxTrans.LastTrans = NumOfTTRecs + 1
  Put TTHandle, 36397, TaxTrans
  
  Close
  MsgBox ("Done.")

End Sub

Private Sub cmdMakeBillTypesC_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  Dim Cnt As Long
  
  frmTaxShowPctComp.Label1 = "Fixing Bill Types"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    If TaxTrans.BillType <> "C" Then
      TaxTrans.BillType = "C"
      Put TTHandle, x, TaxTrans
      Cnt = Cnt + 1
    End If
    frmTaxShowPctComp.ShowPctComp x, NumOfTTRecs
  Next x
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  Close
  Call Savemsg(900, "A total of " + CStr(Cnt) + " transactions were modified successfully.")
End Sub

Private Sub cmdOrphan_Click()
  Call Look4OrphanPayTrans
End Sub

Private Sub cmdProcess1_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Long
  Dim Balance#
  Dim Principle1 As Double
  Dim Principle2 As Double
  Dim Principle3 As Double
  Dim Principle4 As Double
  Dim Principle5 As Double
  Dim Interest As Double
  Dim Future1 As Double
  Dim Future2 As Double
  Dim Collection As Double
  Dim LateList As Double
  Dim RevOpt1 As Double
  Dim RevOpt2 As Double
  Dim RevOpt3 As Double
  Dim Principle1Pd As Double
  Dim Principle2Pd As Double
  Dim Principle3Pd As Double
  Dim Principle4Pd As Double
  Dim Principle5Pd As Double
  Dim LateListPd As Double
  Dim Future1Pd As Double
  Dim Future2Pd As Double
  Dim CollectionPd As Double
  Dim InterestPd As Double
  Dim RevOpt1Pd As Double
  Dim RevOpt2Pd As Double
  Dim RevOpt3Pd As Double
  Dim Principle1Bill As Double
  Dim Principle2Bill As Double
  Dim Principle3Bill As Double
  Dim Principle4Bill As Double
  Dim Principle5Bill As Double
  Dim InterestBill As Double
  Dim Future1Bill As Double
  Dim Future2Bill As Double
  Dim CollectionBill As Double
  Dim LateListBill As Double
  Dim RevOpt1Bill As Double
  Dim RevOpt2Bill As Double
  Dim RevOpt3Bill As Double
  Dim NegCnt As Long
  Dim TaxCust As TaxCustType
  Dim NumOfTCRecs As Long
  Dim TCHandle As Integer
  Dim NextRec As Long
  Dim SaveHere As Long
  Dim SaveCnt As Long
  Dim StartDate As Integer
  Dim EndDate As Integer
  Dim TotRev As Double
'  Dim TotRevP As Double
  
  StartDate = Date2Num(fptxtBegDate.Text)
  EndDate = Date2Num(fptxtEndDate.Text)
  ReDim RevPay(1 To 9) As Double
  ReDim RevChg(1 To 9) As Double
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
     Select Case x
       Case 76303
         TaxTrans.Amount = TaxTrans.Amount
         TaxTrans.Revenue.Interest = 0 'TaxTrans.Revenue.Interest
         TaxTrans.Revenue.Principle1 = 0
         TaxTrans.Amount = 0
         Put TTHandle, x, TaxTrans
       Case 8436
         TaxTrans.Amount = TaxTrans.Amount
         TaxTrans.Revenue.Principle1 = 21.01 'TaxTrans.Revenue.Principle1
         TaxTrans.Revenue.Principle1Pd = 21.01 'TaxTrans.Revenue.Principle1Pd
         Put TTHandle, x, TaxTrans
       Case 11494
         TaxTrans.Amount = TaxTrans.Amount
         TaxTrans.Amount = 0
         TaxTrans.Revenue.Interest = 0 'TaxTrans.Revenue.Interest
         Put TTHandle, x, TaxTrans
       Case 11346
         TaxTrans.Amount = TaxTrans.Amount
         TaxTrans.Amount = 0
         TaxTrans.Revenue.Interest = 0 'TaxTrans.Revenue.Interest
         Put TTHandle, x, TaxTrans
       Case 11195
         TaxTrans.Amount = TaxTrans.Amount
         TaxTrans.Amount = 0
         TaxTrans.Revenue.Interest = 0 'TaxTrans.Revenue.Interest
         Put TTHandle, x, TaxTrans
       Case 11026
         TaxTrans.Amount = TaxTrans.Amount
         TaxTrans.Amount = 0
         TaxTrans.Revenue.Interest = 0 'TaxTrans.Revenue.Interest
         Put TTHandle, x, TaxTrans
       Case 10505
         TaxTrans.Amount = TaxTrans.Amount
         TaxTrans.Amount = 0
         TaxTrans.Revenue.Interest = 0 'TaxTrans.Revenue.Interest
         Put TTHandle, x, TaxTrans
       Case 8993
         TaxTrans.Amount = TaxTrans.Amount
          TaxTrans.Amount = 0
         TaxTrans.Revenue.Interest = 0 'TaxTrans.Revenue.Interest
         Put TTHandle, x, TaxTrans
       Case Else
     End Select
''    If TaxTrans.CustomerRec = 7052 Then '
'      If TaxTrans.TranType = 1 Then
''        Select Case x
'        Select Case TaxTrans.CustomerRec
'          Case 1800
'            TaxTrans.Description = "M Tax Bill 100"
''            RevChg(1) = TaxTrans.Revenue.Principle1
''            TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd
'          Case 1801
'            TaxTrans.Description = "M Tax Bill 200"
''            RevChg(2) = TaxTrans.Revenue.Principle1
'          Case 1802
'            TaxTrans.Description = "M Tax Bill 300"
''            RevChg(3) = TaxTrans.Revenue.Principle1
'          Case 1803
'            TaxTrans.Description = "M Tax Bill 400"
''            RevChg(4) = TaxTrans.Revenue.Principle1
'          Case 1804
'            TaxTrans.Description = "M Tax Bill 500"
''            RevChg(5) = TaxTrans.Revenue.Principle1
'          Case 1805
'            TaxTrans.Description = "M Tax Bill 600"
''            RevChg(6) = TaxTrans.Revenue.Principle1
'          Case 1806
'            TaxTrans.Description = "M Tax Bill 700"
''            RevChg(7) = TaxTrans.Revenue.Principle1
'          Case Else
''            RevChg(8) = TaxTrans.Revenue.Principle1
'        End Select
'        Put TTHandle, x, TaxTrans
''      Else
''        Select Case TaxTrans.BelongTo
''          Case 16892
''            RevPay(1) = OldRound(RevPay(1) + TaxTrans.Revenue.Principle1Pd)
'            TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd
'            TaxTrans.Revenue.Interest = TaxTrans.Revenue.Interest
'            TaxTrans.Revenue.CollectionPd = TaxTrans.Revenue.CollectionPd
'            TaxTrans.Revenue.LateListPd = TaxTrans.Revenue.LateListPd
'            TaxTrans.Revenue.Future1Pd = TaxTrans.Revenue.Future1Pd
'            TaxTrans.Revenue.Future2Pd = TaxTrans.Revenue.Future2Pd
'            TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd
'            TaxTrans.Revenue.Principle2Pd = TaxTrans.Revenue.Principle2Pd
'            TaxTrans.Revenue.Principle3Pd = TaxTrans.Revenue.Principle3Pd
'            TaxTrans.Revenue.Principle4Pd = TaxTrans.Revenue.Principle4Pd
'            TaxTrans.Revenue.Principle5Pd = TaxTrans.Revenue.Principle5Pd
'            TaxTrans.Revenue.Collection = TaxTrans.Revenue.Collection
'            TaxTrans.Revenue.LateList = TaxTrans.Revenue.LateList
'            TaxTrans.Revenue.Future1 = TaxTrans.Revenue.Future1
'            TaxTrans.Revenue.Future2 = TaxTrans.Revenue.Future2
'            TaxTrans.Revenue.Principle1 = TaxTrans.Revenue.Principle1
'            TaxTrans.Revenue.Principle2 = TaxTrans.Revenue.Principle2
'            TaxTrans.Revenue.Principle3 = TaxTrans.Revenue.Principle3
'            TaxTrans.Revenue.Principle4 = TaxTrans.Revenue.Principle4
'            TaxTrans.Revenue.Principle5 = TaxTrans.Revenue.Principle5
'            TaxTrans.Amount = TaxTrans.Amount
''          Case 62842
''            RevPay(2) = OldRound(RevPay(2) + TaxTrans.Revenue.Principle1Pd)
''          Case 122662
''            RevPay(3) = OldRound(RevPay(3) + TaxTrans.Revenue.Principle1Pd)
''          Case 181625
''            RevPay(4) = OldRound(RevPay(4) + TaxTrans.Revenue.Principle1Pd)
''          Case 245680
''            RevPay(5) = OldRound(RevPay(5) + TaxTrans.Revenue.Principle1Pd)
''          Case 299621
''            RevPay(6) = OldRound(RevPay(6) + TaxTrans.Revenue.Principle1Pd)
''          Case 348941
''            RevPay(7) = OldRound(RevPay(7) + TaxTrans.Revenue.Principle1Pd)
''          Case Is > 0
''            RevPay(8) = OldRound(RevPay(8) + TaxTrans.Revenue.Principle1Pd)
''          Case Else
''            RevPay(9) = OldRound(RevPay(9) + TaxTrans.Revenue.Principle1Pd)
''       End Select
'     End If
'''      TotRevP = TotRev + TaxTrans.Revenue.Principle1Pd
'''      TotRev = TotRev + TaxTrans.Revenue.Principle1
'''      If TaxTrans.BelongTo > 0 Then
'''        Debug.Print CStr(TaxTrans.BelongTo)
'''      End If
''
'''    End If
  Next x
  Exit Sub
'
'  For x = 1 To 9
'    Debug.Print "Charges/Payments for " + CStr(x) + ":" + Using$("$##,##0.00", RevChg(x)) + " / " + Using$("$##,##0.00", RevPay(x))
'  Next x
  frmTaxShowPctComp.Label1 = "Fixing Negative Balances"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  ReDim NegCust(1 To 1) As Long
  ReDim NegTrans(1 To 1) As Long
  
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    TaxTrans.CustomerRec = TaxTrans.CustomerRec
    If TaxTrans.TransDate < StartDate Or TaxTrans.TransDate > EndDate Then GoTo OD
    If TaxTrans.TranType = 1 Then
      TotRev = OldRound#(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TotRev = OldRound#(TotRev + TaxTrans.Revenue.LateList)
'      If OldRound(TaxTrans.Amount) > TotRev Then Stop
      Balance# = OldRound#(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      Balance# = OldRound#(Balance# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.LateList)
      Balance# = OldRound#(Balance# + TaxTrans.Revenue.Future1 + TaxTrans.Revenue.Future2 + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      Balance# = OldRound#(Balance# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      Balance# = OldRound#(Balance# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      Balance# = OldRound#(Balance# - (TaxTrans.Revenue.Future1Pd + TaxTrans.Revenue.Future2Pd + TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd))
      If Balance# < 0 Then
        NegCnt = NegCnt + 1
        ReDim Preserve NegTrans(1 To NegCnt) As Long
        NegTrans(NegCnt) = x
        ReDim Preserve NegCust(1 To NegCnt) As Long
        NegCust(NegCnt) = TaxTrans.CustomerRec
'        If NegCust(NegCnt) = 3722 Then Stop
      End If
    End If
    frmTaxShowPctComp.ShowPctComp x, NumOfTTRecs
OD:
  Next x
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  frmTaxShowPctComp.Label1 = "Fixing Negative Balances"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
    
  For x = 1 To NegCnt
    If NegCust(x) <= 0 Then GoTo SkipIt
    Get TCHandle, NegCust(x), TaxCust
'    If TaxCust.Acct = 3722 Then Stop '4/17/06
      SaveCnt = 0
      Principle1Pd = 0
      Future1Pd = 0
      Future2Pd = 0
      InterestPd = 0
      CollectionPd = 0
      LateListPd = 0
      Interest = 0
      For y = 1 To NumOfTTRecs
        Get TTHandle, y, TaxTrans
'        If y = 8537 And TaxCust.Acct = 3722 Then Stop
        If TaxTrans.TranType = 1 Then 'check all billing trans and strip out negatives
          If TaxTrans.Revenue.Principle1 < 0 Then
            TaxTrans.Revenue.Principle1 = 0
            TaxTrans.Amount = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.LateList)
            Put TTHandle, y, TaxTrans
          End If
          If TaxTrans.Revenue.Future1 < 0 Then
            TaxTrans.Revenue.Future1 = 0
            Put TTHandle, y, TaxTrans
          End If
          If TaxTrans.Revenue.Future2 < 0 Then
            TaxTrans.Revenue.Future2 = 0
            Put TTHandle, y, TaxTrans
          End If
          If TaxTrans.Revenue.Interest < 0 Then
            TaxTrans.Revenue.Interest = 0
            Put TTHandle, y, TaxTrans
          End If
          If TaxTrans.Revenue.Collection < 0 Then
            TaxTrans.Revenue.Collection = 0
            Put TTHandle, y, TaxTrans
          End If
          If TaxTrans.Revenue.LateList < 0 Then
            TaxTrans.Revenue.LateList = 0
            TaxTrans.Amount = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.LateList)
            Put TTHandle, y, TaxTrans
          End If
          GoTo NextTrans
        End If
        If TaxTrans.BelongTo = NegTrans(x) Then
'          If TaxTrans.BelongTo = 16892 Then Stop
          Get TTHandle, TaxTrans.BelongTo, TaxTrans
          Principle1Bill = TaxTrans.Revenue.Principle1
          InterestBill = TaxTrans.Revenue.Interest
          Future1Bill = TaxTrans.Revenue.Future1
          Future2Bill = TaxTrans.Revenue.Future2
          CollectionBill = TaxTrans.Revenue.Collection
          LateListBill = TaxTrans.Revenue.LateList
          Get TTHandle, y, TaxTrans
'          If TaxTrans.TranType <> 2 Then
'            Debug.Print CStr(TaxTrans.TranType)
'          End If
'          If TaxTrans.TranType = 8 Then Stop
          Select Case TaxTrans.TranType
            Case 2, 7 'payment, adjustment
                Principle1Pd = OldRound#(Principle1Pd + TaxTrans.Revenue.Principle1Pd) 'collect all revenues and
                'wait for a problem
                If Principle1Pd > Principle1Bill Then 'OK...we have a problem
                  If TaxTrans.Revenue.Principle1 < 0 Then TaxTrans.Revenue.Principle1 = 0
                  TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd - (Principle1Pd - Principle1Bill))
                  Principle1Pd = OldRound(Principle1Pd - (Principle1Pd - Principle1Bill))
                  TaxTrans.Amount = OldRound(TaxTrans.Amount - Principle1Pd) 'reduce the existing total amount for
                  'this transaction by the amount of the overpay
                  Put TTHandle, y, TaxTrans 'save it to this transaction
                  SaveHere = TaxTrans.BelongTo 'establish the bill record
                  Get TTHandle, TaxTrans.BelongTo, TaxTrans 'pull the bill record
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.Principle1Pd = Principle1Pd 'correct the bill record for this resource
                  'TaxTrans.Amount is not affected because this value only reflects the original charges
                  'and not any updates to the bill
                  Put TTHandle, SaveHere, TaxTrans 'save it to the bill record
                  Get TTHandle, y, TaxTrans 'go back to original record
                End If
                Future1Pd = OldRound(Future1Pd + TaxTrans.Revenue.Future1Pd)
                If Future1Pd > Future1Bill Then
                  If TaxTrans.Revenue.Future1 < 0 Then TaxTrans.Revenue.Future1 = 0
                  TaxTrans.Revenue.Future1Pd = OldRound(TaxTrans.Revenue.Future1Pd - (Future1Pd - Future1Bill))
                  Future1Pd = OldRound(Future1Pd - (Future1Pd - Future1Bill))
                  TaxTrans.Amount = OldRound(TaxTrans.Amount - Future1Pd)
                  Put TTHandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get TTHandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.Future1Pd = OldRound(TaxTrans.Revenue.Future1Pd - (Future1Pd - Future1Bill))
                  Put TTHandle, SaveHere, TaxTrans
                  Get TTHandle, y, TaxTrans
                End If
                Future2Pd = OldRound(Future2Pd + TaxTrans.Revenue.Future2Pd)
                If Future2Pd > Future2Bill Then
                  If TaxTrans.Revenue.Future2 < 0 Then TaxTrans.Revenue.Future2 = 0
                  TaxTrans.Revenue.Future2Pd = OldRound(TaxTrans.Revenue.Future2Pd - (Future2Pd - Future2Bill))
                  Future2Pd = OldRound(Future2Pd - (Future2Pd - Future2Bill))
                  TaxTrans.Amount = OldRound(TaxTrans.Amount - Future2Pd)
                  Put TTHandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get TTHandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.Future2Pd = OldRound(TaxTrans.Revenue.Future2Pd - (Future2Pd - Future2Bill))
                  Put TTHandle, SaveHere, TaxTrans
                  Get TTHandle, y, TaxTrans
                End If
                InterestPd = OldRound(InterestPd + TaxTrans.Revenue.InterestPd)
                If InterestPd > InterestBill Then
                  If TaxTrans.Revenue.Interest < 0 Then TaxTrans.Revenue.Interest = 0
                  TaxTrans.Revenue.InterestPd = OldRound(TaxTrans.Revenue.InterestPd - (InterestPd - InterestBill))
                  InterestPd = OldRound(InterestPd - (InterestPd - InterestBill))
                  TaxTrans.Amount = TaxTrans.Amount - InterestPd
                  Put TTHandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get TTHandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.InterestPd = InterestPd
                  Put TTHandle, SaveHere, TaxTrans
                  Get TTHandle, y, TaxTrans
                End If
                CollectionPd = OldRound(CollectionPd + TaxTrans.Revenue.CollectionPd)
                If CollectionPd > CollectionBill Then
                  If TaxTrans.Revenue.Collection < 0 Then TaxTrans.Revenue.Collection = 0
                  TaxTrans.Revenue.CollectionPd = OldRound(TaxTrans.Revenue.CollectionPd - (CollectionPd - CollectionBill))
                  CollectionPd = OldRound(CollectionPd - (CollectionPd - CollectionBill))
                  TaxTrans.Amount = OldRound(TaxTrans.Amount - CollectionPd)
                  Put TTHandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get TTHandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.CollectionPd = OldRound(TaxTrans.Revenue.CollectionPd - (CollectionPd - CollectionBill))
                  Put TTHandle, SaveHere, TaxTrans
                  Get TTHandle, y, TaxTrans
                End If
                LateListPd = OldRound(LateListPd + TaxTrans.Revenue.LateListPd)
                If LateListPd > LateListBill Then
                  If TaxTrans.Revenue.LateList < 0 Then TaxTrans.Revenue.LateList = 0
                  TaxTrans.Revenue.LateListPd = OldRound(TaxTrans.Revenue.LateListPd - (LateListPd - LateListBill))
                  LateListPd = OldRound(LateListPd - (LateListPd - LateListBill))
                  TaxTrans.Amount = OldRound(TaxTrans.Amount - LateListPd)
                  Put TTHandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get TTHandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.LateListPd = OldRound(TaxTrans.Revenue.LateListPd - (LateListPd - LateListBill))
                  Put TTHandle, SaveHere, TaxTrans
                  Get TTHandle, y, TaxTrans
                End If
              Case 4:
                Interest = OldRound(Interest + TaxTrans.Revenue.Interest)
                If Interest > InterestBill Then
                  If TaxTrans.Revenue.Interest < 0 Then TaxTrans.Revenue.Interest = 0
                  TaxTrans.Revenue.Interest = OldRound(TaxTrans.Revenue.Interest - (Interest - InterestBill))
                  Interest = OldRound(Interest - (Interest - InterestBill))
                  TaxTrans.Amount = TaxTrans.Revenue.Interest ' OldRound(TaxTrans.Amount - Interest)
                  Put TTHandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get TTHandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.InterestPd = InterestBill ' OldRound(TaxTrans.Revenue.InterestPd - (Interest - InterestBill))
                  Put TTHandle, SaveHere, TaxTrans
                  Get TTHandle, y, TaxTrans
                End If
              Case 6, 8:
                Collection = OldRound(Collection + TaxTrans.Revenue.Collection)
                If Collection > CollectionBill Then
                  If TaxTrans.Revenue.Collection < 0 Then TaxTrans.Revenue.Collection = 0
                  TaxTrans.Revenue.Collection = OldRound(TaxTrans.Revenue.Collection - (Collection - CollectionBill))
                  Collection = OldRound(Collection - (Collection - CollectionBill))
                  TaxTrans.Amount = TaxTrans.Revenue.Collection 'OldRound(TaxTrans.Amount - Collection)
                  Put TTHandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get TTHandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.CollectionPd = CollectionBill 'OldRound(TaxTrans.Revenue.CollectionPd - (Collection - CollectionBill))
                  Put TTHandle, SaveHere, TaxTrans
                  Get TTHandle, y, TaxTrans
                End If
            End Select
         End If
NextTrans:
      Next y
      If SaveCnt = 0 Then 'been through all transactions for this negative balance and
      'couldn't find a transaction that belonged to it. This means something is negative that
      'shouldn't be
        Get TTHandle, NegTrans(x), TaxTrans 'always trans type #1
        If TaxTrans.Revenue.Principle1 < 0 Then TaxTrans.Revenue.Principle1 = 0
        TaxTrans.Amount = TaxTrans.Revenue.Principle1
        If TaxTrans.Revenue.Interest < 0 Then TaxTrans.Revenue.Interest = 0
        TaxTrans.Amount = OldRound(TaxTrans.Amount + TaxTrans.Revenue.Interest)
        If TaxTrans.Revenue.Collection < 0 Then TaxTrans.Revenue.Collection = 0
        TaxTrans.Amount = OldRound(TaxTrans.Amount + TaxTrans.Revenue.Collection)
        If TaxTrans.Revenue.LateList < 0 Then TaxTrans.Revenue.LateList = 0
        TaxTrans.Amount = OldRound(TaxTrans.Amount + TaxTrans.Revenue.LateList)
        If TaxTrans.Revenue.Future1 < 0 Then TaxTrans.Revenue.Future1 = 0
        TaxTrans.Amount = OldRound(TaxTrans.Amount + TaxTrans.Revenue.Future1)
        If TaxTrans.Revenue.Future2 < 0 Then TaxTrans.Revenue.Future2 = 0
        TaxTrans.Amount = OldRound(TaxTrans.Amount + TaxTrans.Revenue.Future2)
        Put TTHandle, NegTrans(x), TaxTrans
      End If
    frmTaxShowPctComp.ShowPctComp x, NegCnt
SkipIt:
  Next x
  
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
           
  frmTaxShowPctComp.Label1 = "Deleting All Future Field Values"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  For y = 1 To NumOfTTRecs 'strip out all future1 and future2 fields
    Get TTHandle, y, TaxTrans
    If TaxTrans.Revenue.Future1Pd <> 0 Then
      TaxTrans.Revenue.Future1Pd = 0
    End If
    If TaxTrans.Revenue.Future1 <> 0 Then
      TaxTrans.Revenue.Future1 = 0
    End If
    If TaxTrans.Revenue.Future2Pd <> 0 Then
      TaxTrans.Revenue.Future2Pd = 0
    End If
    If TaxTrans.Revenue.Future2 <> 0 Then
      TaxTrans.Revenue.Future2 = 0
    End If
    Put TTHandle, y, TaxTrans
    frmTaxShowPctComp.ShowPctComp y, NumOfTTRecs
  Next y
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
          
  Close
  Call Savemsg(900, "The negative and future vales repairs have completed successfully.")
End Sub


Private Sub cmdProcess9_Click()
  'Yadkinville repair job-------------------------------
'  Dim TransRec As TaxTransactionType
'  Dim THandle As Integer
'  Dim NumOfTRecs As Long
'  Dim x As Long
'  Dim TransYear As Integer
'  Dim TaxYearS As String
'  Dim TransYearS As String
'
'  OpenTaxTransFile THandle, NumOfTRecs
'  For x = 1 To NumOfTRecs
'    Get THandle, x, TransRec
'    If TransRec.TranType = 1 Then GoTo SkipIt
'    TransYearS = MakeRegDate(TransRec.TransDate)
''    TransYear = 2006 ' CInt(Mid(TransYearS, 7, 4))
'    TransYearS = Mid(TransYearS, 1, 6)
'    TransYearS = TransYearS + "2006"
'    TransRec.TransDate = Date2Num(TransYearS)
'    Put THandle, x, TransRec
'SkipIt:
'  Next x
'  Close
'  Call TaxMsg(900, "Yadkinville non-bill transactions have been amended successfully.")
  'Yadkinville repair job^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'  Call MakeAmtEqualRevs

  Dim TaxTrans As TaxTransactionType
  Dim NumOfTrans As Long
  Dim THandle As Integer
  Dim x As Long
  Dim TCnt As Long
  Dim CountIt As Boolean
  Dim ThisCollection As Double
  Dim ThisPenalty As Double
  Dim ThisInterest As Double
  Dim ThisLateList As Double
  Dim ThisPrinciple1 As Double
  Dim ThisPrinciple2 As Double
  Dim ThisPrinciple3 As Double
  Dim ThisPrinciple4 As Double
  Dim ThisPrinciple5 As Double
  Dim ThisRevOpt1 As Double
  Dim ThisRevOpt2 As Double
  Dim ThisRevOPt3 As Double
  Dim ThisBelongTo As Long
  
  frmTaxShowPctComp.Label1 = "Updating Release Revenues"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  
  OpenTaxTransFile THandle, NumOfTrans
  For x = 1 To NumOfTrans
    CountIt = False
    Get THandle, x, TaxTrans
    If TaxTrans.TranType = 3 Then
      ThisCollection = 0
      ThisPenalty = 0
      ThisInterest = 0
      ThisLateList = 0
      ThisPrinciple1 = 0
      ThisPrinciple2 = 0
      ThisPrinciple3 = 0
      ThisPrinciple4 = 0
      ThisPrinciple5 = 0
      ThisRevOpt1 = 0
      ThisRevOpt2 = 0
      ThisRevOPt3 = 0
      If TaxTrans.Revenue.Penalty > 0 And TaxTrans.Revenue.PenaltyPd = 0 Then
        TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.Penalty
        TaxTrans.Revenue.Penalty = 0
        ThisPenalty = TaxTrans.Revenue.PenaltyPd
        CountIt = True
      End If
      If TaxTrans.Revenue.Collection > 0 And TaxTrans.Revenue.CollectionPd = 0 Then
        TaxTrans.Revenue.CollectionPd = TaxTrans.Revenue.Collection
        TaxTrans.Revenue.Collection = 0
        ThisCollection = TaxTrans.Revenue.CollectionPd
        CountIt = True
      End If
      If TaxTrans.Revenue.Interest > 0 And TaxTrans.Revenue.InterestPd = 0 Then
        TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.Interest
        TaxTrans.Revenue.Interest = 0
        ThisInterest = TaxTrans.Revenue.InterestPd
        CountIt = True
      End If
      If TaxTrans.Revenue.LateList > 0 And TaxTrans.Revenue.LateListPd = 0 Then
        TaxTrans.Revenue.LateListPd = TaxTrans.Revenue.LateList
        TaxTrans.Revenue.LateList = 0
        ThisLateList = TaxTrans.Revenue.LateListPd
        CountIt = True
      End If
      If TaxTrans.Revenue.Principle1 > 0 And TaxTrans.Revenue.Principle1Pd = 0 Then
        TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1
        TaxTrans.Revenue.Principle1 = 0
        ThisPrinciple1 = TaxTrans.Revenue.Principle1Pd
        CountIt = True
      End If
      If TaxTrans.Revenue.Principle2 > 0 And TaxTrans.Revenue.Principle2Pd = 0 Then
        TaxTrans.Revenue.Principle2Pd = TaxTrans.Revenue.Principle2
        TaxTrans.Revenue.Principle2 = 0
        ThisPrinciple2 = TaxTrans.Revenue.Principle2Pd
        CountIt = True
      End If
      If TaxTrans.Revenue.Principle3 > 0 And TaxTrans.Revenue.Principle3Pd = 0 Then
        TaxTrans.Revenue.Principle3Pd = TaxTrans.Revenue.Principle3
        TaxTrans.Revenue.Principle3 = 0
        ThisPrinciple3 = TaxTrans.Revenue.Principle3Pd
        CountIt = True
      End If
      If TaxTrans.Revenue.Principle4 > 0 And TaxTrans.Revenue.Principle4Pd = 0 Then
        TaxTrans.Revenue.Principle4Pd = TaxTrans.Revenue.Principle4
        TaxTrans.Revenue.Principle4 = 0
        ThisPrinciple4 = TaxTrans.Revenue.Principle4Pd
        CountIt = True
      End If
      If TaxTrans.Revenue.Principle5 > 0 And TaxTrans.Revenue.Principle5Pd = 0 Then
        TaxTrans.Revenue.Principle5Pd = TaxTrans.Revenue.Principle5
        TaxTrans.Revenue.Principle5 = 0
        ThisPrinciple5 = TaxTrans.Revenue.Principle5Pd
        CountIt = True
      End If
      If TaxTrans.Revenue.RevOpt1 > 0 And TaxTrans.Revenue.RevOpt1Pd = 0 Then
        TaxTrans.Revenue.RevOpt1Pd = TaxTrans.Revenue.RevOpt1
        TaxTrans.Revenue.RevOpt1 = 0
        ThisRevOpt1 = TaxTrans.Revenue.RevOpt1Pd
        CountIt = True
      End If
      If TaxTrans.Revenue.RevOpt2 > 0 And TaxTrans.Revenue.RevOpt2Pd = 0 Then
        TaxTrans.Revenue.RevOpt2Pd = TaxTrans.Revenue.RevOpt2
        TaxTrans.Revenue.RevOpt2 = 0
        ThisRevOpt2 = TaxTrans.Revenue.RevOpt2Pd
        CountIt = True
      End If
      If TaxTrans.Revenue.RevOpt3 > 0 And TaxTrans.Revenue.RevOpt3Pd = 0 Then
        TaxTrans.Revenue.RevOpt3Pd = TaxTrans.Revenue.RevOpt3
        TaxTrans.Revenue.RevOpt3 = 0
        ThisRevOPt3 = TaxTrans.Revenue.RevOpt3Pd
        CountIt = True
      End If
      Put THandle, x, TaxTrans
   End If
   If CountIt = True Then
     ThisBelongTo = TaxTrans.BelongTo
     Get THandle, ThisBelongTo, TaxTrans
       If ThisCollection > 0 Then
         TaxTrans.Revenue.Collection = OldRound(TaxTrans.Revenue.Collection + ThisCollection)
         TaxTrans.Revenue.CollectionPd = OldRound(TaxTrans.Revenue.CollectionPd + ThisCollection)
       End If
       If ThisPenalty > 0 Then
         TaxTrans.Revenue.Penalty = OldRound(TaxTrans.Revenue.Penalty + ThisPenalty)
         TaxTrans.Revenue.PenaltyPd = OldRound(TaxTrans.Revenue.PenaltyPd + ThisPenalty)
       End If
       If ThisInterest > 0 Then
         TaxTrans.Revenue.Interest = OldRound(TaxTrans.Revenue.Interest + ThisInterest)
         TaxTrans.Revenue.InterestPd = OldRound(TaxTrans.Revenue.InterestPd + ThisInterest)
       End If
       If ThisLateList > 0 Then
         TaxTrans.Revenue.LateList = OldRound(TaxTrans.Revenue.LateList + ThisLateList)
         TaxTrans.Revenue.LateListPd = OldRound(TaxTrans.Revenue.LateListPd + ThisLateList)
       End If
       If ThisPrinciple1 > 0 Then
         TaxTrans.Revenue.Principle1 = OldRound(TaxTrans.Revenue.Principle1 + ThisPrinciple1)
         TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd + ThisPrinciple1)
       End If
       If ThisPrinciple2 > 0 Then
         TaxTrans.Revenue.Principle2 = OldRound(TaxTrans.Revenue.Principle2 + ThisPrinciple2)
         TaxTrans.Revenue.Principle2Pd = OldRound(TaxTrans.Revenue.Principle2Pd + ThisPrinciple2)
       End If
       If ThisPrinciple3 > 0 Then
         TaxTrans.Revenue.Principle3 = OldRound(TaxTrans.Revenue.Principle3 + ThisPrinciple3)
         TaxTrans.Revenue.Principle3Pd = OldRound(TaxTrans.Revenue.Principle3Pd + ThisPrinciple3)
       End If
       If ThisPrinciple4 > 0 Then
         TaxTrans.Revenue.Principle4 = OldRound(TaxTrans.Revenue.Principle4 + ThisPrinciple4)
         TaxTrans.Revenue.Principle4Pd = OldRound(TaxTrans.Revenue.Principle4Pd + ThisPrinciple4)
       End If
       If ThisPrinciple5 > 0 Then
         TaxTrans.Revenue.Principle5 = OldRound(TaxTrans.Revenue.Principle5 + ThisPrinciple5)
         TaxTrans.Revenue.Principle5Pd = OldRound(TaxTrans.Revenue.Principle5Pd + ThisPrinciple5)
       End If
       If ThisRevOpt1 > 0 Then
         TaxTrans.Revenue.RevOpt1 = OldRound(TaxTrans.Revenue.RevOpt1 + ThisRevOpt1)
         TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd + ThisRevOpt1)
       End If
       If ThisRevOpt2 > 0 Then
         TaxTrans.Revenue.RevOpt2 = OldRound(TaxTrans.Revenue.RevOpt2 + ThisRevOpt2)
         TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd + ThisRevOpt2)
       End If
       If ThisRevOPt3 > 0 Then
         TaxTrans.Revenue.RevOpt3 = OldRound(TaxTrans.Revenue.RevOpt3 + ThisRevOPt3)
         TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd + ThisRevOPt3)
       End If
       Put THandle, ThisBelongTo, TaxTrans
     TCnt = TCnt + 1
   End If
   frmTaxShowPctComp.ShowPctComp x, NumOfTrans
  Next x
  
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
    
  Close THandle
  
  Call Savemsg(900, "A total of " + CStr(TCnt) + " release transactions were updated successfully.")
  
End Sub

Private Sub cmdRelinkBSL_Click()
  Dim TransRec As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  Dim CHandle As Integer
  Dim CustRec As TaxCustType
  Dim NumOfCRecs As Long
  Dim x As Long
  Dim NextRec As Long
  Dim y As Long
  Dim ThisPin As String
  Dim Cnt As Long
  Dim AHandle As Integer
  
  AHandle = FreeFile
  Open "bslrelink.dat" For Output As AHandle
  Print #AHandle, "CustName" + "~" + "Old Cust Pin" + "~" + "New Cust Pin" + "~" + "Real Pin"
  OpenTaxTransFile THandle, NumOfTRecs
  OpenTaxCustFile CHandle, NumOfCRecs
  OpenRealPropFile RHandle, NumOfRRecs
   
  For x = 1 To NumOfCRecs
    Get CHandle, x, CustRec
    NextRec = CustRec.FirstPropRec
    Do While NextRec > 0
      Get RHandle, NextRec, RealRec
      ThisPin = QPTrim$(RealRec.RealPin)
      For y = 1 To NumOfTRecs
        Get THandle, y, TransRec
        If QPTrim$(TransRec.RealPin) = ThisPin Then
          If TransRec.CustomerRec <> CustRec.PIN Then
            Print #AHandle, QPTrim$(CustRec.CustName) + "~" + CStr(TransRec.CustPin) + "~" + CStr(CustRec.PIN) + "~" + QPTrim$(RealRec.RealPin)
            TransRec.CustPin = CustRec.PIN
            Put THandle, y, TransRec
            Cnt = Cnt + 1
          End If
        End If
      Next y
      NextRec = RealRec.NextRec
    Loop
  Next x
  
  
End Sub

Private Sub cmdRemoveTrans_Click()
ClearTrans (CLng(tbxTransNum.Text))
MsgBox (tbxTransNum.Text + " removed successfully.")
End Sub

Private Sub cmdRepairDatesOnFixedTrans_Click()
  Dim TransRec As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim x As Long
  Dim TransYear As Integer
  Dim TaxYearS As String
  Dim TransYearS As String
  
  OpenTaxTransFile THandle, NumOfTRecs
  For x = 1 To NumOfTRecs
    Get THandle, x, TransRec
    If TransRec.TranType <> 1 Then GoTo SkipIt 'added 6/30/06 to change only type 1 transactions
    TransYearS = MakeRegDate(TransRec.TransDate)
    TransYear = CInt(Mid(TransYearS, 7, 4))
    If TransYear <> TransRec.TaxYear Then
      TaxYearS = CStr(TransRec.TaxYear)
      TransYearS = Mid(TransYearS, 1, 6)
      TransYearS = TransYearS + CStr(TaxYearS)
      TransRec.TransDate = Date2Num(TransYearS)
      Put THandle, x, TransRec
    End If
SkipIt:
  Next x
  Close
    
End Sub

Private Sub cmdRunningBal_Click()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim Balance As Double
  Dim AHandle As Integer
  
  AHandle = FreeFile
  Open "runningbal.dat" For Output As AHandle
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    Balance = GetCustBalance(TaxCust.Acct, -1)
    Print #AHandle, CStr(TaxCust.Acct) & "~" & QPTrim$(TaxCust.CustName) & "~" & Using("$##,###.##", Balance)
  Next x
  
  Close
  MsgBox ("Completed successfully.")
End Sub

Private Sub cmdStringOrNumber_Click()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim Scnt As Long
  Dim NCnt As Long
  Dim XCnt As Long
  Dim AHandle As Integer
  Dim CAStr As String
  AHandle = FreeFile
  Open "countyacctTest" For Output As AHandle
  
  Open "countyacctTest.txt" For Output As AHandle
  OpenTaxCustFile TCHandle, NumOfTCRecs
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TaxCust.CountyAcct > 0 Then
      NCnt = NCnt + 1
      Print #AHandle, "Number" & "~" & CStr(TaxCust.Acct) & "~" & QPTrim$(TaxCust.CustName) & "~" & CStr(TaxCust.CountyAcct)
    End If
    CAStr = QPTrim(TaxCust.CountyAcctString)
    If CAStr <> "" Then
      Scnt = Scnt + 1
      Print #AHandle, "String" & "~" & CStr(TaxCust.Acct) & "~" & QPTrim$(TaxCust.CustName) & "~" & CAStr
    End If
    If TaxCust.CountyAcct = 0 And CAStr = "" Then
      XCnt = XCnt + 1
      Print #AHandle, "Neither" & "~" & CStr(TaxCust.Acct) & "~" & QPTrim$(TaxCust.CustName) & "~" & CAStr
    End If
  Next x
  
  Close
  MsgBox ("County numbers = " & CStr(NCnt) & ". County strings = " & CStr(Scnt) & ". Neither = " & CStr(XCnt) & ".")

End Sub

Private Sub cmdStripBoiling_Click()
 Call StripOutTrans
End Sub

Private Sub cmdUpdateAdd1andAdd2Long_Click()
  Dim x As Long, y As Long
  Dim TextLine$
  Dim ThisFile$
  Dim Handle As Integer
  Dim WordCnt As Integer
  Dim TextLen As Integer
  Dim Thisch As String
  Dim ThisWord$
  Dim CntyNum As String
  Dim Add1 As String
  Dim Add2 As String
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim dlm As String
  Dim Cnt As Integer
  Dim track As Integer
  
  If MsgBox("Did you make the last line = 'End~~'?", vbYesNo) = vbNo Then
    Exit Sub
  End If
  dlm = "~"
  track = 0
  WordCnt = 0
  ReDim Words(1 To 1) As String
  frmTaxShowPctComp.Label1 = "Addresses Update"
  frmTaxShowPctComp.Show , Me

  If Exist("addresses.csv") Then
    Handle = FreeFile
    ThisFile = "addresses.csv"
    Open ThisFile For Input As #Handle
    Do While ThisWord <> "End"
      Line Input #Handle, TextLine
      If InStr(TextLine, "End~~") Then Exit Do
      track = track + 1
    Loop
    Close
    OpenTaxCustFile TCHandle, NumOfTCRecs
    Handle = FreeFile
    Open ThisFile For Input As #Handle
    Do While ThisWord <> "End"
      Line Input #Handle, TextLine
      TextLen = Len(TextLine)
      TextLine = TextLine + dlm
      For x = 1 To TextLen + 1
        Thisch = Mid(TextLine, x, 1)
        If Thisch = dlm Then
          WordCnt = WordCnt + 1
          ReDim Preserve Words(1 To WordCnt) As String
          If WordCnt = 1 Then
            CntyNum = ThisWord
            ThisWord = ""
            GoTo NewWord
          ElseIf WordCnt = 2 Then
            Add1 = ThisWord
            ThisWord = ""
            GoTo NewWord
          ElseIf WordCnt = 3 Then
            Add2 = ThisWord
            GoSub SaveAdd
            Add1 = ""
            Add2 = ""
            CntyNum = ""
            ThisWord = ""
            WordCnt = 0
            GoTo NewLoop
          End If
        End If
        ThisWord = ThisWord + Thisch
        If ThisWord = "End" Then Exit Do

NewWord:
      Next x
NewLoop:
    frmTaxShowPctComp.ShowPctComp Cnt, track
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
    End If
    Loop
  Else
    MsgBox ("The file 'addresses.csv' cannot be found.")
    Exit Sub
  End If
  Unload frmTaxShowPctComp
  
  MsgBox (CStr(Cnt) & " addresses were updated successfully.")
  Close
  Exit Sub
  
SaveAdd:
  Cnt = Cnt + 1
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If QPTrim$(TaxCust.CountyAcctString) = "" Then
      TaxCust.CountyAcctString = CStr(TaxCust.CountyAcct)
    End If
    If QPTrim$(TaxCust.CountyAcctString) = CntyNum Then
      TaxCust.Addr1 = Add1
      TaxCust.Addr2 = Add2
      Put TCHandle, x, TaxCust
    End If
Skip:
  Next x
  
  Return
End Sub

Private Sub cmdUpdateAdd1AndAdd2Short_Click()
  Dim x As Long, y As Long
  Dim TextLine$
  Dim ThisFile$
  Dim Handle As Integer
  Dim WordCnt As Integer
  Dim TextLen As Integer
  Dim Thisch As String
  Dim ThisWord$
  Dim CntyNum As String
  Dim Add1 As String
  Dim Add2 As String
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim dlm As String
  Dim Cnt As Integer
  Dim track As Integer
  
  If MsgBox("Did you make the last line = 'End~~'?", vbYesNo) = vbNo Then
    Exit Sub
  End If
  dlm = "~"
  track = 0
  WordCnt = 0
  ReDim Words(1 To 1) As String
  frmTaxShowPctComp.Label1 = "Addresses Update"
  frmTaxShowPctComp.Show , Me

  If Exist("addresses.csv") Then
    Handle = FreeFile
    ThisFile = "addresses.csv"
    Open ThisFile For Input As #Handle
    Do While ThisWord <> "End"
      Line Input #Handle, TextLine
      If InStr(TextLine, "End~~") Then Exit Do
      track = track + 1
    Loop
    Close
    OpenTaxCustFile TCHandle, NumOfTCRecs
    Handle = FreeFile
    Open ThisFile For Input As #Handle
    Do While ThisWord <> "End"
      Line Input #Handle, TextLine
      TextLen = Len(TextLine)
      TextLine = TextLine + dlm
      For x = 1 To TextLen + 1
        Thisch = Mid(TextLine, x, 1)
        If Thisch = dlm Then
          WordCnt = WordCnt + 1
          ReDim Preserve Words(1 To WordCnt) As String
          If WordCnt = 1 Then
            CntyNum = ThisWord
            ThisWord = ""
            GoTo NewWord
          ElseIf WordCnt = 2 Then
            Add1 = ThisWord
            ThisWord = ""
            GoTo NewWord
          ElseIf WordCnt = 3 Then
            Add2 = ThisWord
            GoSub SaveAdd
            Add1 = ""
            Add2 = ""
            CntyNum = ""
            ThisWord = ""
            WordCnt = 0
            GoTo NewLoop
          End If
        End If
        ThisWord = ThisWord + Thisch
        If ThisWord = "End" Then Exit Do

NewWord:
      Next x
NewLoop:
    frmTaxShowPctComp.ShowPctComp Cnt, track
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
    End If
    Loop
  Else
    MsgBox ("The file 'addresses.csv' cannot be found.")
    Exit Sub
  End If
  Unload frmTaxShowPctComp
  
  MsgBox (CStr(Cnt) & " addresses were updated successfully.")
  Close
  Exit Sub
  
SaveAdd:
  Cnt = Cnt + 1
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If QPTrim$(TaxCust.CountyAcctString) = "" Then
      TaxCust.CountyAcctString = CStr(TaxCust.CountyAcct)
    End If
    If QPTrim$(TaxCust.CountyAcctString) = CntyNum Then
      TaxCust.Addr1 = Add1
      TaxCust.Addr2 = Add2
      Put TCHandle, x, TaxCust
      Exit For
    End If
Skip:
  Next x
  
  Return
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
'    Case vbKeyF10:
'      SendKeys "%P"
'      Call cmdProcess1_Click
'      KeyCode = 0
    Case vbKeyF9:
      SendKeys "%r"
      Call cmdProcess2_Click
      KeyCode = 0
    Case vbKeyF8:
      SendKeys "%0"
      Call cmdProcess3_Click
      KeyCode = 0
    Case vbKeyF7:
      SendKeys "%c"
      Call cmdProcess4_Click
      KeyCode = 0
    Case vbKeyF6:
      SendKeys "%e"
      Call cmdProcess5_Click
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%s"
      Call cmdProcess6_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      ClearInUse PWcnt
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxDataRepair.")
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

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  MainLog ("User opened frmTaxDataRepair.")
  Call LoadMe
End Sub

Private Sub LoadMe()
  Dim BHandle As Integer
  Dim CDateStr$
  Dim TaxTrans As TaxTransactionType
  Dim NumOfTRecs As Long
  Dim x As Long
  Dim THandle As Integer
  
'  Label14.Visible = False
'  cmdProcess9.Visible = False
'  Shape11.Visible = False
'  Call FixMaggieValley
'  Call FixSunsetBeach
'  Call FixCalabash
'  Call FindOddRealOwner
  
  Label3.Visible = False
  Shape3.Visible = False
  Label1.Visible = False
  Label4.Visible = False
  
  fptxtBegDate.Visible = False
  fptxtEndDate.Visible = False
  cmdProcess1.Visible = False
  Call FixSpecificData
  OpenTaxTransFile THandle, NumOfTRecs
  If NumOfTRecs = 0 Then
    Call TaxMsg(900, "No transactions stored.")
    Close
    Exit Sub
  End If
  
  For x = 1 To NumOfTRecs
    Get THandle, x, TaxTrans
    If TaxTrans.TransDate > 0 Then
      fptxtBegDate.Text = MakeRegDate(x)
      Close
      Exit For
    End If
  Next x
    
'  lblBalloon.Visible = False
  If Exist("cnvtdate.dat") Then
    BHandle = FreeFile
    Open "cnvtdate.dat" For Input As BHandle
    Input #BHandle, CDateStr$
    Close BHandle
    If QPTrim(CDateStr$) = "" Then
      fptxtEndDate.Text = Date
      Exit Sub
    End If
    fptxtEndDate.Text = MakeRegDate(CInt(CDateStr$))
'    lblMessage.Visible = True
'    lblMessage.Caption = "The date in the 'Ending Date' field is the conversion date for DOS to Windows,  " + fptxtEndDate.Text + "."
  Else
'    lblMessage.Visible = False
'    fptxtEndDate.Text = Date
  End If
  
  fptxtFiscalBeg.Text = "07/01"
  fptxtFiscalEnd.Text = "06/30"
  
End Sub

Private Sub cmdProcess2_Click()
  Dim x As Long
  Dim TransRec As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxYear As Integer
  Dim YrCnt As Long
  
  frmTaxShowPctComp.Label1 = "Repairing Tax Years"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  OpenTaxTransFile THandle, NumOfTRecs
  
  If NumOfTRecs = 0 Then
    Call TaxMsg(900, "There are no transactions saved.")
    Exit Sub
  End If
  
  For x = 1 To NumOfTRecs
    Get THandle, x, TransRec
    If TransRec.TranType <> 1 Then
      If TransRec.TaxYear = 0 Then
        If TransRec.BelongTo > 0 Then
          Get THandle, TransRec.BelongTo, TransRec
          TaxYear = TransRec.TaxYear
          Get THandle, x, TransRec
          TransRec.TaxYear = TaxYear
          Put THandle, x, TransRec
          YrCnt = YrCnt + 1
        End If
      End If
    End If
    frmTaxShowPctComp.ShowPctComp x, NumOfTRecs
  Next x
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  Close
  Call Savemsg(900, "A total of " + CStr(YrCnt) + " errant tax years were corrected successfully.")
  
End Sub

Private Sub ResequencePins()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  
  frmTaxShowPctComp.Label1 = "Resequencing Customer Pin Numbers"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  OpenTaxCustFile TCHandle, NumOfTCRecs
  If NumOfTCRecs = 0 Then
    Close
    Call TaxMsg(900, "No tax customers saved. Re-sequencing aborted.")
    Exit Sub
  End If
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    TaxCust.Acct = x
    TaxCust.PIN = x
    Put TCHandle, x, TaxCust
    frmTaxShowPctComp.ShowPctComp x, NumOfTCRecs
  Next x
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  Close
  Call Savemsg(900, "Re-sequencing of pin numbers completed successfully.")
  
End Sub

Private Sub cmdProcess3_Click()
  Dim x As Long
  Dim TransRec As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxYear As Integer
  Dim YrCnt As Long
  Dim UseThisYear As Integer
  Dim FiscBeg As Integer
  Dim TestYear$
  Dim EndOfYear As Integer
  Dim BegOfYear As Integer
  Dim JulBeg As Integer
  Dim TestJulBeg As Integer
  Dim TestJulEnd As Integer
  
  frmTaxShowPctComp.Label1 = "Repairing Tax Years"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  OpenTaxTransFile THandle, NumOfTRecs
  
  UseThisYear = Year(fptxtFiscalBeg.Text)
  TestYear = "12/31/" + CStr(UseThisYear)
  EndOfYear = Date2Num(TestYear)
  FiscBeg = Date2Num(fptxtFiscalBeg.Text)
  JulBeg = EndOfYear - FiscBeg 'julian start date
  If NumOfTRecs = 0 Then
    Call TaxMsg(900, "There are no transactions saved.")
    Close
    Exit Sub
  End If
  
  For x = 1 To NumOfTRecs
    Get THandle, x, TransRec
    If TransRec.TaxYear = 0 Then
      If TransRec.TranType <> 1 And TransRec.TranType <> 22 Then GoTo SkipIt
      If TransRec.TransDate < 0 Then
        TransRec.TransDate = 0
      End If
      TestJulBeg = TransRec.TransDate
      TestYear = Mid(MakeRegDate(TransRec.TransDate), 7, 4)
      EndOfYear = Date2Num("12/31/" + TestYear)
      BegOfYear = Date2Num("01/01/" + TestYear)
      If (EndOfYear - TestJulBeg) >= (EndOfYear - (JulBeg + BegOfYear)) Then
        TransRec.TaxYear = CInt(TestYear)
      Else
        TransRec.TaxYear = CInt(TestYear) + 1
      End If
      YrCnt = YrCnt + 1
      Put THandle, x, TransRec
    End If
    frmTaxShowPctComp.ShowPctComp x, NumOfTRecs
SkipIt:
  Next x
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  Call cmdProcess2_Click
  Close
  Call Savemsg(900, "A total of " + CStr(YrCnt) + " errant tax years were corrected successfully.")
  
End Sub
Private Sub cmdFixMaxton0Years_Click()
  Dim x As Long
  Dim TransRec As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxYear As Integer
  Dim YrCnt As Long
  Dim UseThisYear As Integer
  Dim FiscBeg As Integer
  Dim TestYear$
  Dim EndOfYear As Integer
  Dim BegOfYear As Integer
  Dim JulBeg As Integer
  Dim TestJulBeg As Integer
  Dim TestJulEnd As Integer
  
  frmTaxShowPctComp.Label1 = "Repairing Tax Years"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  OpenTaxTransFile THandle, NumOfTRecs
  
  UseThisYear = Year(fptxtFiscalBeg.Text)
  TestYear = "12/31/" + CStr(UseThisYear)
  EndOfYear = Date2Num(TestYear)
  FiscBeg = Date2Num(fptxtFiscalBeg.Text)
  JulBeg = EndOfYear - FiscBeg 'julian start date
  If NumOfTRecs = 0 Then
    Call TaxMsg(900, "There are no transactions saved.")
    Close
    Exit Sub
  End If
  
  For x = 1 To NumOfTRecs
    Get THandle, x, TransRec
    If TransRec.TaxYear = 0 Then
      If TransRec.TranType <> 22 Then GoTo SkipIt
      If TransRec.TransDate < 0 Then
        TransRec.TransDate = 0
      End If
      TestJulBeg = TransRec.TransDate
      TestYear = Mid(MakeRegDate(TransRec.TransDate), 7, 4)
      EndOfYear = Date2Num("12/31/" + TestYear)
      BegOfYear = Date2Num("01/01/" + TestYear)
      If (EndOfYear - TestJulBeg) >= (EndOfYear - (JulBeg + BegOfYear)) Then
        TransRec.TaxYear = CInt(TestYear)
      Else
        TransRec.TaxYear = CInt(TestYear) + 1
      End If
      YrCnt = YrCnt + 1
      Put THandle, x, TransRec
    End If
    frmTaxShowPctComp.ShowPctComp x, NumOfTRecs
SkipIt:
  Next x
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
'  Call cmdProcess2_Click
  Close
  Call Savemsg(900, "A total of " + CStr(YrCnt) + " errant tax years were corrected successfully.")
  
End Sub

Private Sub cmdProcess4_Click()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim Cnt As Long
  
  frmTaxShowPctComp.Label1 = "Resequencing Customer Pin Numbers"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  OpenTaxCustFile TCHandle, NumOfTCRecs
  If NumOfTCRecs = 0 Then
    Close
    Call TaxMsg(900, "No tax customers saved. Re-sequencing aborted.")
    Exit Sub
  End If
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TaxCust.Acct <> x Then Cnt = Cnt + 1
    TaxCust.Acct = x
    TaxCust.PIN = x
    Put TCHandle, x, TaxCust
    frmTaxShowPctComp.ShowPctComp x, NumOfTCRecs
  Next x
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  Close
  Call Savemsg(900, CStr(Cnt) + " Re-sequencing of pin numbers completed successfully.")
  
End Sub

Private Sub fpcmdHelp_Click()
'  If InStr(fpcmdHelp.Text, "On") Then
'    fpcmdHelp.Text = "F1 &Turn Help Off"
'    btnHelp.AutoScan = fpAutoScanPopupOnly
'    lblBalloon.Visible = True
'  ElseIf InStr(fpcmdHelp.Text, "Off") Then
'    fpcmdHelp.Text = "F1 &Turn Help On"
'    btnHelp.AutoScan = fpAutoScanOff
'    lblBalloon.Visible = False
'  End If
End Sub

'Private Sub cmdProcess5_Click()
'  Dim TaxTrans As TaxTransactionType
'  Dim TTHandle As Integer
'  Dim NumOfTTRecs As Long
'  Dim x As Long
'  Dim ThisAmtP As Double
'  Dim ThisAmtB As Double
'  Dim RepairCnt As Long
'
'  OpenTaxTransFile TTHandle, NumOfTTRecs
'  frmTaxShowPctComp.Label1 = "Making Transaction Totals Equal Revenues"
'  frmTaxShowPctComp.CmdCancel.Visible = False
'  frmTaxShowPctComp.Show , Me
'  DoEvents
'  EnableCloseButton Me.hwnd, False
'  DoEvents
'  For x = 1 To NumOfTTRecs
'  Get TTHandle, x, TaxTrans
''    If x = 26519 Then Stop
'    ThisAmtP = OldRound(TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.Future1Pd + TaxTrans.Revenue.Future2Pd)
'    ThisAmtP = OldRound(ThisAmtP + TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.LateListPd + TaxTrans.Revenue.PenaltyPd)
'    ThisAmtP = OldRound(ThisAmtP + TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd)
'    ThisAmtP = OldRound(ThisAmtP + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd)
'    ThisAmtP = OldRound(ThisAmtP + TaxTrans.Revenue.Principle5Pd + TaxTrans.Revenue.RevOpt1Pd)
'    ThisAmtP = OldRound(ThisAmtP + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd)
''    If TaxTrans.TranType <> 1 Then
'      ThisAmtB = OldRound(TaxTrans.Revenue.Collection + TaxTrans.Revenue.Future1 + TaxTrans.Revenue.Future2)
'      ThisAmtB = OldRound(ThisAmtB + TaxTrans.Revenue.Interest + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Penalty)
''    Else
''      ThisAmtB = OldRound(TaxTrans.Revenue.Future1 + TaxTrans.Revenue.Future2)
''      ThisAmtB = OldRound(ThisAmtB + TaxTrans.Revenue.LateList) ' + TaxTrans.Revenue.Penalty)
''    End If
'    ThisAmtB = OldRound(ThisAmtB + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2)
'    ThisAmtB = OldRound(ThisAmtB + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4)
'    ThisAmtB = OldRound(ThisAmtB + TaxTrans.Revenue.Principle5 + TaxTrans.Revenue.RevOpt1)
'    ThisAmtB = OldRound(ThisAmtB + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
'
'    Select Case TaxTrans.TranType
'      Case 1:
'        If TaxTrans.Amount <> ThisAmtB Then
'          TaxTrans.Amount = ThisAmtB
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 2:
'        If TaxTrans.Amount <> ThisAmtP Then
'          TaxTrans.Amount = ThisAmtP
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 3:
'        If TaxTrans.Amount <> ThisAmtB Then
'          TaxTrans.Amount = ThisAmtB
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 4:
'        If TaxTrans.Amount <> ThisAmtB Then
'          TaxTrans.Amount = ThisAmtB
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 6, 8:
'        If TaxTrans.Amount <> ThisAmtB Then
'          TaxTrans.Amount = ThisAmtB
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 7:
'        If TaxTrans.Amount <> ThisAmtP Then
'          TaxTrans.Amount = ThisAmtP
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 8:
'        If TaxTrans.Amount <> ThisAmtB Then
'          TaxTrans.Amount = ThisAmtB
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 9:
'        If TaxTrans.Amount <> ThisAmtP Then
'          TaxTrans.Amount = ThisAmtP
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 13:
'        If TaxTrans.Amount <> ThisAmtB Then
'          TaxTrans.Amount = ThisAmtB
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 14:
'         If TaxTrans.Amount <> ThisAmtB Then
'          TaxTrans.Amount = ThisAmtB
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 21:
'        If TaxTrans.Amount <> ThisAmtP Then
'          TaxTrans.Amount = ThisAmtP
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 22:
'        If TaxTrans.Amount <> ThisAmtP Then
'          TaxTrans.Amount = ThisAmtP
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 10:
'        If TaxTrans.Amount <> ThisAmtP Then
'          TaxTrans.Amount = ThisAmtP
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 11:
'        If TaxTrans.Amount <> ThisAmtP Then
'          TaxTrans.Amount = ThisAmtP
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 12:
'        If TaxTrans.Amount <> ThisAmtP Then
'          TaxTrans.Amount = ThisAmtP
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 24:
'        If TaxTrans.Amount <> ThisAmtB Then
'          TaxTrans.Amount = ThisAmtB
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'    End Select
'
'    frmTaxShowPctComp.ShowPctComp x, NumOfTTRecs
'  Next x
'  Unload frmTaxShowPctComp
'  EnableCloseButton Me.hwnd, True
'
'  Close
'
'  Call Savemsg(900, CStr(RepairCnt) + " transaction amounts have been updated successfully")
'
'End Sub

Private Sub cmdProcess6_Click()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim x As Long, y As Long
  Dim TotRev As Double
  Dim BelongToCnt  As Integer
  Dim TotPrincPd As Double
  Dim TotLateListPd As Double
  Dim TotCollectionPd As Double
  Dim TotInterestPd As Double
  Dim TotPenaltyPd As Double
  Dim PrincBilledPaid As Double
  Dim LateListBilledPaid As Double
  Dim CollectionBilledPaid As Double
  Dim InterestBilledPaid As Double
  Dim PenaltyBilledPaid As Double
  Dim ErrCnt As Long
  
  If TaxMsgWOpts(900, "Run this only after you have already repaired negative values. THIS PROCESS IS VERY LENGTHY.", "F10 Continue", "ESC Quit Now") = "abort" Then
    Close
    Exit Sub
  End If
  OpenTaxTransFile THandle, NumOfTRecs
  ReDim ErrTrans(1 To 1) As Long
  For x = 4998 To NumOfTRecs
    Get THandle, x, TaxTrans
    If TaxTrans.TranType = 1 Then
      If OldRound(TaxTrans.Revenue.Principle1Pd) > OldRound(TaxTrans.Revenue.Principle1) Or OldRound(TaxTrans.Revenue.CollectionPd) > OldRound(TaxTrans.Revenue.Collection) _
      Or OldRound(TaxTrans.Revenue.LateListPd) > OldRound(TaxTrans.Revenue.LateList) Or OldRound(TaxTrans.Revenue.InterestPd) > OldRound(TaxTrans.Revenue.Interest) Or _
      TaxTrans.Revenue.PenaltyPd > TaxTrans.Revenue.Penalty Then
        ErrCnt = ErrCnt + 1
        ReDim Preserve ErrTrans(1 To ErrCnt) As Long
        ErrTrans(ErrCnt) = x
      End If
    End If
  Next x
  
  If ErrCnt = 0 Then
    Close
    Call TaxMsg(900, "There are no billing transactions affected by this error.")
    Exit Sub
  End If
  
  frmTaxShowPctComp.Label1 = "Making Billing Paid Values Equal Belong To Values"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  For x = 1 To ErrCnt
    Get THandle, ErrTrans(x), TaxTrans
      PrincBilledPaid = TaxTrans.Revenue.Principle1Pd
      LateListBilledPaid = TaxTrans.Revenue.LateListPd
      CollectionBilledPaid = TaxTrans.Revenue.CollectionPd
      InterestBilledPaid = TaxTrans.Revenue.InterestPd
      PenaltyBilledPaid = TaxTrans.Revenue.PenaltyPd
      TotPrincPd = 0
      TotLateListPd = 0
      TotCollectionPd = 0
      TotInterestPd = 0
      TotPenaltyPd = 0
      For y = ErrTrans(x) + 1 To NumOfTRecs 'start at ErrTrans(x) because all trans related to bill
        'will come after the bill trans
        Get THandle, y, TaxTrans
        If TaxTrans.BelongTo = ErrTrans(x) Then
          TotPrincPd = OldRound#(TotPrincPd + TaxTrans.Revenue.Principle1Pd)
          TotLateListPd = OldRound#(TotLateListPd + TaxTrans.Revenue.LateListPd)
          TotCollectionPd = OldRound#(TotCollectionPd + TaxTrans.Revenue.CollectionPd)
          TotInterestPd = OldRound#(TotInterestPd + TaxTrans.Revenue.InterestPd)
          TotPenaltyPd = OldRound#(TotPenaltyPd + TaxTrans.Revenue.PenaltyPd)
        End If
'        frmTaxShowPctComp.ShowPctComp2 y, NumOfTRecs
      Next y
'      frmTaxShowPctComp.Label1 = "Making Billing Paid Values Equal Belong To Values"
'      frmTaxShowPctComp.CmdCancel.Visible = False
'      frmTaxShowPctComp.Show , Me
'      DoEvents
'      EnableCloseButton Me.hwnd, False
      If TotPrincPd < PrincBilledPaid Then
        Get THandle, ErrTrans(x), TaxTrans
        BelongToCnt = BelongToCnt + 1
        TaxTrans.Revenue.Principle1Pd = TotPrincPd
        Put THandle, ErrTrans(x), TaxTrans
        End If
      If TotLateListPd < LateListBilledPaid Then
        Get THandle, ErrTrans(x), TaxTrans
        BelongToCnt = BelongToCnt + 1
        TaxTrans.Revenue.LateListPd = TotLateListPd
        Put THandle, ErrTrans(x), TaxTrans
      End If
      If TotCollectionPd < CollectionBilledPaid Then
        Get THandle, ErrTrans(x), TaxTrans
        BelongToCnt = BelongToCnt + 1
        TaxTrans.Revenue.CollectionPd = TotCollectionPd
        Put THandle, ErrTrans(x), TaxTrans
      End If
      If TotInterestPd < InterestBilledPaid Then
        Get THandle, ErrTrans(x), TaxTrans
        BelongToCnt = BelongToCnt + 1
        TaxTrans.Revenue.InterestPd = TotInterestPd
        Put THandle, ErrTrans(x), TaxTrans
      End If
      If TotPenaltyPd < PenaltyBilledPaid Then
        Get THandle, ErrTrans(x), TaxTrans
        BelongToCnt = BelongToCnt + 1
        TaxTrans.Revenue.PenaltyPd = TotPenaltyPd
        Put THandle, ErrTrans(x), TaxTrans
      End If
    frmTaxShowPctComp.ShowPctComp x, ErrCnt
  Next x
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  Close
  
  Call Savemsg(900, CStr(BelongToCnt) + " billing transaction amounts have been updated successfully")

End Sub

Private Sub cmdProcess7_Click()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxCust As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim x As Long, y As Long, z As Integer
  Dim AmtBalance As Double
  Dim RevBalance As Double
  Dim NextRec As Long
  Dim ErrCnt As Long
  Dim BTCnt As Integer
  Dim TotPrincPd As Double
  Dim TotLateListPd As Double
  Dim TotCollectionPd As Double
  Dim TotInterestPd As Double
  Dim TotPenaltyPd As Double
  Dim PrincBilledPaid As Double
  Dim LateListBilledPaid As Double
  Dim CollectionBilledPaid As Double
  Dim InterestBilledPaid As Double
  Dim PenaltyBilledPaid As Double
  Dim RevOpt1 As Double
  Dim RevOpt2 As Double
  Dim RevOpt3 As Double
  Dim Principle1Pd As Double
  Dim Principle2Pd As Double
  Dim Principle3Pd As Double
  Dim Principle4Pd As Double
  Dim Principle5Pd As Double
  Dim LateListPd As Double
  Dim Future1Pd As Double
  Dim Future2Pd As Double
  Dim CollectionPd As Double
  Dim Collection As Double
  Dim InterestPd As Double
  Dim Interest As Double
  Dim RevOpt1Pd As Double
  Dim RevOpt2Pd As Double
  Dim RevOpt3Pd As Double
  Dim Principle1BillPd As Double
  Dim Principle1Bill As Double
  Dim Principle2Bill As Double
  Dim Principle3Bill As Double
  Dim Principle4Bill As Double
  Dim Principle5Bill As Double
  Dim InterestBill As Double
  Dim InterestBillPd As Double
  Dim Future1Bill As Double
  Dim Future2Bill As Double
  Dim Future1BillPd As Double
  Dim Future2BillPd As Double
  Dim CollectionBill As Double
  Dim CollectionBillPd As Double
  Dim LateListBill As Double
  Dim LateListBillPd As Double
  Dim RevOpt1Bill As Double
  Dim RevOpt2Bill As Double
  Dim RevOpt3Bill As Double
  Dim SaveHere As Long
  Dim SaveCnt As Long
  Dim ThisPinNum As Long
  Dim FixNextRec As Long
  Dim SaveFixRec As Long
  Dim WrongCustRec As Long
  Dim NewRec As Long
  
  ReDim ErrCust(1 To 1) As Long
  OpenTaxCustFile CHandle, NumOfCRecs
  OpenTaxTransFile THandle, NumOfTRecs
'  For x = 1 To NumOfTRecs
'    Get THandle, x, TaxTrans
'    If x = 4998 Then Stop
'    If TaxTrans.BelongTo = 4998 Then Stop
'    If TaxTrans.Amount = 13.65 And TaxTrans.BelongTo = 198 Then
'      Stop
'      TaxTrans.CustomerRec = TaxTrans.CustomerRec
'      SaveCnt = SaveCnt + 1
'    End If
'  Next x
  SaveCnt = 0
  frmTaxShowPctComp.Label1 = "Finding Errant Balances"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  For x = 1 To 1 'NumOfCRecs
    x = 1356
    Get CHandle, x, TaxCust
    If TaxCust.Deleted <> 0 Then GoTo SkipMe
    RevBalance = 0
'    If x = 66 Then Stop
    AmtBalance = GetCustBalance(x, -1)
    NextRec = TaxCust.LastTrans
    Do While NextRec > 0
      Get THandle, NextRec, TaxTrans
'      If NextRec = 4998 Then Stop
      If TaxTrans.TranType <> 1 Then GoTo NotThisOne
      RevBalance# = OldRound#(RevBalance# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      RevBalance# = OldRound#(RevBalance# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      RevBalance# = OldRound#(RevBalance# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      RevBalance# = OldRound#(RevBalance# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      RevBalance# = OldRound#(RevBalance# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt))
NotThisOne:
      If TaxTrans.Revenue.CollectionPd < 0 Then TaxTrans.Revenue.CollectionPd = 0
      Put THandle, NextRec, TaxTrans
      NextRec = TaxTrans.LastTrans
    Loop
    If OldRound(RevBalance) <> OldRound(AmtBalance) Then
      ErrCnt = ErrCnt + 1
      ReDim Preserve ErrCust(1 To ErrCnt) As Long
      ErrCust(ErrCnt) = x
'      AmtBalance = GetCustBalance(x, -1)
    End If
SkipMe:
    frmTaxShowPctComp.ShowPctComp x, NumOfCRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      Exit Sub
    End If
  Next x
  
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
    
  frmTaxShowPctComp.Label1 = "Examining Transactions and Fixing Problems"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  
  For x = 1 To ErrCnt
    ReDim BelongTo(1 To 1) As Long
    BTCnt = 0
    Get CHandle, ErrCust(x), TaxCust
    
    NextRec = TaxCust.LastTrans
    Do While NextRec > 0
      Get THandle, NextRec, TaxTrans
'      If NextRec = 4998 Then Stop
      If TaxTrans.TranType = 1 Then
        BTCnt = BTCnt + 1
        ReDim Preserve BelongTo(1 To BTCnt) As Long
        BelongTo(BTCnt) = NextRec
      End If
      NextRec = TaxTrans.LastTrans
    Loop
    For z = 1 To BTCnt
      Get THandle, BelongTo(z), TaxTrans
      Principle1Bill = OldRound(TaxTrans.Revenue.Principle1)
      InterestBill = OldRound(TaxTrans.Revenue.Interest)
      Future1Bill = OldRound(TaxTrans.Revenue.Future1)
      Future2Bill = OldRound(TaxTrans.Revenue.Future2)
      CollectionBill = OldRound(TaxTrans.Revenue.Collection)
      LateListBill = OldRound(TaxTrans.Revenue.LateList)
      Principle1BillPd = OldRound(TaxTrans.Revenue.Principle1Pd)
      InterestBillPd = OldRound(TaxTrans.Revenue.InterestPd)
      Future1BillPd = OldRound(TaxTrans.Revenue.Future1Pd)
      Future2BillPd = OldRound(TaxTrans.Revenue.Future2Pd)
      CollectionBillPd = OldRound(TaxTrans.Revenue.CollectionPd)
      LateListBillPd = OldRound(TaxTrans.Revenue.LateListPd)
      Principle1Pd = 0
      Future1Pd = 0
      Future2Pd = 0
      InterestPd = 0
      CollectionPd = 0
      LateListPd = 0
      Interest = 0
      Collection = 0
      For y = BelongTo(z) To NumOfTRecs
        Get THandle, y, TaxTrans
'        If y = 25359 Then Stop
        If TaxTrans.BelongTo = BelongTo(z) Then
          Select Case TaxTrans.TranType
            Case 2, 7 'payment, adjustment
                Principle1Pd = OldRound#(Principle1Pd + TaxTrans.Revenue.Principle1Pd) 'collect all revenues and
                'wait for a problem
                If Principle1Pd > Principle1Bill Then 'OK...we have a problem
                  If TaxTrans.Revenue.Principle1 < 0 Then TaxTrans.Revenue.Principle1 = 0
                  TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd - (Principle1Pd - Principle1Bill))
                  Principle1Pd = OldRound(Principle1Pd - (Principle1Pd - Principle1Bill))
                  TaxTrans.Amount = OldRound(TaxTrans.Amount - Principle1Pd) 'reduce the existing total amount for
                  'this transaction by the amount of the overpay
                  Put THandle, y, TaxTrans 'save it to this transaction
                  SaveHere = TaxTrans.BelongTo 'establish the bill record
                  Get THandle, TaxTrans.BelongTo, TaxTrans 'pull the bill record
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.Principle1Pd = Principle1Pd 'correct the bill record for this resource
                  'TaxTrans.Amount is not affected because this value only reflects the original charges
                  'and not any updates to the bill
                  Put THandle, SaveHere, TaxTrans 'save it to the bill record
                  Get THandle, y, TaxTrans 'go back to original record
                End If
                Future1Pd = OldRound(Future1Pd + TaxTrans.Revenue.Future1Pd)
                If Future1Pd > Future1Bill Then
                  If TaxTrans.Revenue.Future1 < 0 Then TaxTrans.Revenue.Future1 = 0
                  TaxTrans.Revenue.Future1Pd = OldRound(TaxTrans.Revenue.Future1Pd - (Future1Pd - Future1Bill))
                  Future1Pd = OldRound(Future1Pd - (Future1Pd - Future1Bill))
                  TaxTrans.Amount = OldRound(TaxTrans.Amount - Future1Pd)
                  Put THandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get THandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.Future1Pd = OldRound(TaxTrans.Revenue.Future1Pd - (Future1Pd - Future1Bill))
                  Put THandle, SaveHere, TaxTrans
                  Get THandle, y, TaxTrans
                End If
                Future2Pd = OldRound(Future2Pd + TaxTrans.Revenue.Future2Pd)
                If Future2Pd > Future2Bill Then
                  If TaxTrans.Revenue.Future2 < 0 Then TaxTrans.Revenue.Future2 = 0
                  TaxTrans.Revenue.Future2Pd = OldRound(TaxTrans.Revenue.Future2Pd - (Future2Pd - Future2Bill))
                  Future2Pd = OldRound(Future2Pd - (Future2Pd - Future2Bill))
                  TaxTrans.Amount = OldRound(TaxTrans.Amount - Future2Pd)
                  Put THandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get THandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.Future2Pd = OldRound(TaxTrans.Revenue.Future2Pd - (Future2Pd - Future2Bill))
                  Put THandle, SaveHere, TaxTrans
                  Get THandle, y, TaxTrans
                End If
                InterestPd = OldRound(InterestPd + TaxTrans.Revenue.InterestPd)
                If InterestPd > InterestBill Then
                  If TaxTrans.Revenue.Interest < 0 Then TaxTrans.Revenue.Interest = 0
                  TaxTrans.Revenue.InterestPd = OldRound(TaxTrans.Revenue.InterestPd - (InterestPd - InterestBill))
                  InterestPd = OldRound(InterestPd - (InterestPd - InterestBill))
                  TaxTrans.Amount = TaxTrans.Amount - InterestPd
                  Put THandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get THandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.InterestPd = InterestPd
                  Put THandle, SaveHere, TaxTrans
                  Get THandle, y, TaxTrans
                End If
                CollectionPd = OldRound(CollectionPd + TaxTrans.Revenue.CollectionPd)
                If CollectionPd > CollectionBill Then
                  If TaxTrans.Revenue.Collection < 0 Then TaxTrans.Revenue.Collection = 0
                  TaxTrans.Revenue.CollectionPd = OldRound(TaxTrans.Revenue.CollectionPd - (CollectionPd - CollectionBill))
                  CollectionPd = OldRound(CollectionPd - (CollectionPd - CollectionBill))
                  TaxTrans.Amount = OldRound(TaxTrans.Amount - CollectionPd)
                  Put THandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get THandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.CollectionPd = OldRound(TaxTrans.Revenue.CollectionPd - (CollectionPd - CollectionBill))
                  Put THandle, SaveHere, TaxTrans
                  Get THandle, y, TaxTrans
                End If
                LateListPd = OldRound(LateListPd + TaxTrans.Revenue.LateListPd)
                If LateListPd > LateListBill Then
                  If TaxTrans.Revenue.LateList < 0 Then TaxTrans.Revenue.LateList = 0
                  TaxTrans.Revenue.LateListPd = OldRound(TaxTrans.Revenue.LateListPd - (LateListPd - LateListBill))
                  LateListPd = OldRound(LateListPd - (LateListPd - LateListBill))
                  TaxTrans.Amount = OldRound(TaxTrans.Amount - LateListPd)
                  Put THandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get THandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.LateListPd = OldRound(TaxTrans.Revenue.LateListPd - (LateListPd - LateListBill))
                  Put THandle, SaveHere, TaxTrans
                  Get THandle, y, TaxTrans
                End If
              Case 4:
                Interest = OldRound(Interest + TaxTrans.Revenue.Interest)
                If Interest > InterestBill Then
                  If TaxTrans.Revenue.Interest < 0 Then TaxTrans.Revenue.Interest = 0
                  TaxTrans.Revenue.Interest = OldRound(TaxTrans.Revenue.Interest - (Interest - InterestBill))
                  Interest = OldRound(Interest - (Interest - InterestBill))
                  TaxTrans.Amount = TaxTrans.Revenue.Interest ' OldRound(TaxTrans.Amount - Interest)
                  Put THandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get THandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.InterestPd = InterestBill ' OldRound(TaxTrans.Revenue.InterestPd - (Interest - InterestBill))
                  Put THandle, SaveHere, TaxTrans
                  Get THandle, y, TaxTrans
                End If
              Case 6, 8:
                Collection = OldRound(Collection + TaxTrans.Revenue.Collection)
                If Collection > CollectionBill Then
                  If TaxTrans.Revenue.Collection < 0 Then TaxTrans.Revenue.Collection = 0
                  TaxTrans.Revenue.Collection = OldRound(TaxTrans.Revenue.Collection - (Collection - CollectionBill))
                  Collection = OldRound(Collection - (Collection - CollectionBill))
                  TaxTrans.Amount = TaxTrans.Revenue.Collection 'OldRound(TaxTrans.Amount - Collection)
                  Put THandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get THandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.CollectionPd = CollectionBill 'OldRound(TaxTrans.Revenue.CollectionPd - (Collection - CollectionBill))
                  Put THandle, SaveHere, TaxTrans
                  Get THandle, y, TaxTrans
                End If
            End Select
         End If
NextTrans:
      Next y
      If Principle1BillPd <> Principle1Pd Then
        Get THandle, BelongTo(z), TaxTrans
        TaxTrans.Revenue.Principle1 = Principle1Pd
        Put THandle, BelongTo(z), TaxTrans
        SaveCnt = SaveCnt + 1
      End If
      If LateListBillPd <> LateListPd Then
        Get THandle, BelongTo(z), TaxTrans
        TaxTrans.Revenue.LateList = LateListPd
        Put THandle, BelongTo(z), TaxTrans
        SaveCnt = SaveCnt + 1
      End If
      If CollectionBillPd <> Collection Then
        Get THandle, BelongTo(z), TaxTrans
        TaxTrans.Revenue.Collection = Collection
        Put THandle, BelongTo(z), TaxTrans
        SaveCnt = SaveCnt + 1
      End If
      If InterestBillPd <> InterestBill Then
        Get THandle, BelongTo(z), TaxTrans
        TaxTrans.Revenue.InterestPd = InterestPd
        Put THandle, BelongTo(z), TaxTrans
        SaveCnt = SaveCnt + 1
      End If
    Next z
    frmTaxShowPctComp.ShowPctComp x, ErrCnt
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      Exit Sub
    End If
  Next x
  
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  Call Savemsg(900, "Procedure completed successfully.")
End Sub

Private Sub cmdProcess8_Click()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxCust As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim x As Long, y As Long, z As Long, q As Integer
  Dim ErrCnt As Long
  Dim CustCnt As Long
  Dim NextRec As Long
  Dim BillCnt As Long
  Dim SaveHere As Long
  Dim SaveCnt As Long
  Dim ThisPinNum As Long
  Dim FixNextRec As Long
  Dim SaveFixRec As Long
  Dim WrongCustRec As Long
  Dim NewRec As Long
  Dim RightRec As Long
  Dim RightCnt As Integer
  Dim MaxBillCnt As Integer
  Dim UseThisRec As Long
  Dim FixCnt As Long
  
  OpenTaxCustFile CHandle, NumOfCRecs
  OpenTaxTransFile THandle, NumOfTRecs
  
  ReDim CustList(1 To 1) As Long
  For x = 1 To NumOfCRecs
    Get CHandle, x, TaxCust
    If TaxCust.Deleted <> 0 Then GoTo Deleted
    CustCnt = CustCnt + 1
    ReDim Preserve CustList(1 To CustCnt) As Long
    CustList(CustCnt) = x
Deleted:
  Next x
  
'  ReDim CustList(1 To 2) As Long
'  CustList(1) = 66
'  CustList(2) = 5899
'  CustCnt = 2

  frmTaxShowPctComp.Label1 = "Building Arrays Of Customer Transactions"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  
  ReDim CustBillCnt(1 To CustCnt) As Integer
  MaxBillCnt = 0
  For x = 1 To CustCnt
    Get CHandle, CustList(x), TaxCust
    BillCnt = 0
    NextRec = TaxCust.LastTrans
    Do While NextRec > 0
      Get THandle, NextRec, TaxTrans
      If TaxTrans.TranType = 1 Then
        BillCnt = BillCnt + 1
      End If
      NextRec = TaxTrans.LastTrans
    Loop
    If BillCnt > MaxBillCnt Then
      MaxBillCnt = BillCnt
    End If
    frmTaxShowPctComp.ShowPctComp x, CustCnt
  Next x
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  ReDim CustBillQ(1 To CustCnt, 1 To MaxBillCnt) As Long
 
  frmTaxShowPctComp.Label1 = "Building Arrays Of Customer Transactions"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  For x = 1 To CustCnt
    Get CHandle, CustList(x), TaxCust
    BillCnt = 0
    NextRec = TaxCust.LastTrans
    Do While NextRec > 0
      Get THandle, NextRec, TaxTrans
      If TaxTrans.TranType = 1 Then
        BillCnt = BillCnt + 1
        CustBillQ(x, BillCnt) = NextRec
      End If
      NextRec = TaxTrans.LastTrans
    Loop
    If BillCnt = 0 Then
      CustBillQ(x, 1) = -1
      BillCnt = 1
    End If
    CustBillCnt(x) = BillCnt
    frmTaxShowPctComp.ShowPctComp x, CustCnt
  Next x
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  frmTaxShowPctComp.Label1 = "Repairing Orphan Transactions (Final Procedure)"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  For x = 1 To CustCnt 'looking for transactions that do not belong in this Queue
    Get CHandle, CustList(x), TaxCust
    NextRec = TaxCust.LastTrans
    Do While NextRec > 0
      Get THandle, NextRec, TaxTrans
      UseThisRec = TaxTrans.LastTrans
      If TaxTrans.TranType <> 1 Then 'type 1 trans are in the CustBillQ array
        For y = 1 To CustBillCnt(x)
          If CustBillQ(x, y) < 0 Then GoTo NewLoop
          If TaxTrans.BelongTo = CustBillQ(x, y) Then Exit For 'match em up by
          'making sure that each non 1 trans matches up with a 1 trans for this customer
NewLoop:
        Next y
        If y > CustBillCnt(x) Then 'NextRec is an orphan
          GoSub GetRightRec 'go find who it belongs to
          If RightCnt <> 1 Then
            GoTo NoRightRec
          End If
          FixCnt = FixCnt + 1
          Get CHandle, CustList(x), TaxCust 'get the original cust
          FixNextRec = NextRec 'remove this trans from wrong cust Q
          Get THandle, FixNextRec, TaxTrans
          If FixNextRec = TaxCust.LastTrans Then
            TaxCust.LastTrans = TaxTrans.LastTrans 'links past the bad trans
            Put CHandle, CustList(x), TaxCust
          Else
            Get THandle, NextRec, TaxTrans
            SaveHere = TaxTrans.LastTrans
            Get THandle, SaveFixRec, TaxTrans
            TaxTrans.LastTrans = SaveHere 'otherwise it links up with next trans
            Put THandle, SaveFixRec, TaxTrans
          End If
          Get CHandle, RightRec, TaxCust 'now go to correct cust and add the
          'trans posted in error to this customer's Q
          NewRec = TaxCust.LastTrans 'save this to add to the changing trans
          TaxCust.LastTrans = NextRec 'add the new link to the top of the Q
          Put CHandle, RightRec, TaxCust 'save it
          Get THandle, NextRec, TaxTrans 'go back to the trans that is changing Q's
          TaxTrans.LastTrans = NewRec 'insert new link
          Put THandle, NextRec, TaxTrans 'save
          Get CHandle, CustList(x), TaxCust 'return to original customer
        End If
      End If
NoRightRec:
      SaveFixRec = NextRec 'save this rec for linking later on if necessary
      NextRec = UseThisRec
    Loop
    frmTaxShowPctComp.ShowPctComp x, CustCnt
  Next x
  
  Close
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  Call Savemsg(800, "A total of " + CStr(FixCnt) + " orphan transactions were fixed successfully.")
  
  Exit Sub
  
GetRightRec:
  RightCnt = 0
  For z = 1 To CustCnt
    Get CHandle, CustList(z), TaxCust
    For q = 1 To CustBillCnt(z)
      If CustBillQ(z, q) < 0 Then GoTo AnotherLoop
      If CustBillQ(z, q) = TaxTrans.BelongTo Then
        RightRec = CustList(z)
        RightCnt = RightCnt + 1
      End If
AnotherLoop:
    Next q
  Next z
  
  Return
  
  End Sub
  
Private Sub FixSpecificData()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxCust As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim x As Long, NextRec As Long
  Dim ThisAmt As Double
  Dim Princ As Double
  Dim Adv As Double
  Dim LateList As Double
  Dim Interest As Double
  Dim Opt1 As Double
  Dim Opt2 As Double
  Dim Opt3 As Double
  Dim GetBill As Long
  Dim ThisDate$
  Dim SaveAmt As Double
  
  
'  OpenTaxCustFile CHandle, NumOfCRecs
'  Get CHandle, 1, TaxCust
  
'  For x = 1 To NumOfCRecs 'for northwest
'    Get CHandle, x, TaxCust
'    If InStr(TaxCust.CSSN, "N") Then
'      TaxCust.CSSN = ""
'      Put CHandle, x, TaxCust
'    End If
'  Next x
'  Close CHandle
'  TaxTrans.Revenue.Principle1 = TaxTrans.Revenue.Principle1
'  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd
'  TaxTrans.Revenue.Principle2 = TaxTrans.Revenue.Principle2
'  TaxTrans.Revenue.Principle2Pd = TaxTrans.Revenue.Principle2Pd
'  TaxTrans.Revenue.Principle3 = TaxTrans.Revenue.Principle3
'  TaxTrans.Revenue.Principle3Pd = TaxTrans.Revenue.Principle3Pd
'  TaxTrans.Revenue.Principle4 = TaxTrans.Revenue.Principle4
'  TaxTrans.Revenue.Principle4Pd = TaxTrans.Revenue.Principle4Pd
'  TaxTrans.Revenue.Principle5 = TaxTrans.Revenue.Principle5
'  TaxTrans.Revenue.Principle5Pd = TaxTrans.Revenue.Principle5Pd
'  TaxTrans.Revenue.Collection = TaxTrans.Revenue.Collection
'  TaxTrans.Revenue.CollectionPd = TaxTrans.Revenue.CollectionPd
'  TaxTrans.Revenue.Interest = TaxTrans.Revenue.Interest
'  TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd
'  TaxTrans.Revenue.LateList = TaxTrans.Revenue.LateList
'  TaxTrans.Revenue.LateListPd = TaxTrans.Revenue.LateListPd
'  TaxTrans.Revenue.Penalty = TaxTrans.Revenue.Penalty
'  TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd
'  TaxTrans.Revenue.RevOpt1 = TaxTrans.Revenue.RevOpt1
'  TaxTrans.Revenue.RevOpt1Pd = TaxTrans.Revenue.RevOpt1Pd
'  TaxTrans.Revenue.RevOpt2 = TaxTrans.Revenue.RevOpt2
'  TaxTrans.Revenue.RevOpt2Pd = TaxTrans.Revenue.RevOpt2Pd
'  TaxTrans.Revenue.RevOpt3 = TaxTrans.Revenue.RevOpt3
'  TaxTrans.Revenue.RevOpt3Pd = TaxTrans.Revenue.RevOpt3Pd
'  TaxTrans.Amount = TaxTrans.Amount
'
'  OpenTaxTransFile THandle, NumOfTRecs
'  For x = 1 To NumOfTRecs
'    Get THandle, x, TaxTrans
'      If x = 812 Then Stop
'      TaxTrans.CustPin = TaxTrans.CustPin
     'HarrisBurg Fix 6/5/06
'     Get THandle, 3531, TaxTrans
'     TaxTrans.Revenue.Principle1Pd = 231.42
'     Put THandle, 3531, TaxTrans
'
'     Get THandle, 3190, TaxTrans
'     TaxTrans.Revenue.Principle1Pd = 153.86
'     Put THandle, 3190, TaxTrans
'
'     Get THandle, 7429, TaxTrans
'     TaxTrans.Amount = 0
'     TaxTrans.Revenue.Principle1Pd = 0
'     Put THandle, 7429, TaxTrans
'
'     Get THandle, 2022, TaxTrans
'     TaxTrans.Revenue.Principle1Pd = 215.99
'     Put THandle, 7429, TaxTrans
'
'     Get THandle, 3744, TaxTrans
'     TaxTrans.Revenue.Principle1Pd = 288.64
'     Put THandle, 3744, TaxTrans
'
'     Get THandle, 3355, TaxTrans
'     TaxTrans.Revenue.Principle1Pd = 281.02
'     Put THandle, 3355, TaxTrans
'
'     Get THandle, 587, TaxTrans
'     TaxTrans.Revenue.Principle1Pd = 199#
'     Put THandle, 587, TaxTrans
'
'     Close THandle
     'fix for Harrisburg on 6/5/06^^^^^^^^^^
     
'    Get THandle, 20379, TaxTrans
'    TaxTrans.Revenue.Principle1Pd = 76.34
'    Put THandle, 20379, TaxTrans
'
'    Get THandle, 37734, TaxTrans
'    TaxTrans.Revenue.Principle1Pd = 0
'    TaxTrans.Amount = 0
'    Put THandle, 37734, TaxTrans
'
'    Get THandle, 32345, TaxTrans
'    TaxTrans.Revenue.Principle1Pd = 39.6
'    Put THandle, 32345, TaxTrans
'
'    Get THandle, 37438, TaxTrans
'    TaxTrans.Revenue.Principle1Pd = 0
'    TaxTrans.Amount = 0
'    Put THandle, 37438, TaxTrans
'
'    Get THandle, 32365, TaxTrans
'    TaxTrans.Revenue.Principle1Pd = 39.6
'    Put THandle, 32365, TaxTrans
'
'    Get THandle, 52171, TaxTrans
'    TaxTrans.Revenue.Principle1Pd = 0
'    TaxTrans.Amount = 0
'    Put THandle, 52171, TaxTrans
'
'    Get THandle, 47748, TaxTrans
'    TaxTrans.Revenue.Principle1Pd = 59.4
'    Put THandle, 47748, TaxTrans
'
'    Get THandle, 54199, TaxTrans
'    TaxTrans.Revenue.Principle1Pd = 0
'    TaxTrans.Amount = 0
'    Put THandle, 54199, TaxTrans
'
'    Get THandle, 46768, TaxTrans
'    TaxTrans.Revenue.Principle1Pd = 1170.13
'    Put THandle, 46768, TaxTrans
'
'    Get THandle, 13114, TaxTrans
'    TaxTrans.Revenue.Principle1Pd = 0
'    TaxTrans.Amount = 0
'    Put THandle, 13114, TaxTrans
'
'    Get THandle, 10276, TaxTrans
'    TaxTrans.Revenue.Principle1Pd = 0.04
'    Put THandle, 10276, TaxTrans
'
'    Get THandle, 38944, TaxTrans
'    TaxTrans.Revenue.Principle1Pd = 0
'    TaxTrans.Amount = 0
'    Put THandle, 38944, TaxTrans
'
'    Get THandle, 35154, TaxTrans
'    TaxTrans.Revenue.Principle1Pd = 36#
'    Put THandle, 35154, TaxTrans
    
'    If TaxTrans.BelongTo = 16892 And TaxTrans.TranType = 2 Then Stop
'    TaxTrans.Amount = TaxTrans.Amount
'    TaxTrans.TranType = TaxTrans.TranType
'    TaxTrans.CustomerRec = TaxTrans.CustomerRec
'    ThisDate$ = MakeRegDate(TaxTrans.TransDate)
'   -----------Fix for Harrisburg's 1/5/06 errors---------
'    If x = 8760 Then
'      SaveAmt = TaxTrans.Amount
'      TaxTrans.BelongTo = 2264
'      TaxTrans.Description = "2264"
'      Put THandle, x, TaxTrans
'      Get THandle, 3190, TaxTrans
'      TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd - SaveAmt)
'      Put THandle, 3190, TaxTrans
'      Get THandle, 2264, TaxTrans
'      TaxTrans.Revenue.Principle1Pd = SaveAmt
'      Put THandle, 2264, TaxTrans
'    End If
'
'    If x = 8752 Then
'      SaveAmt = TaxTrans.Amount
'      TaxTrans.BelongTo = 2303
'      TaxTrans.Description = "2303"
'      Put THandle, x, TaxTrans
'      Get THandle, 3744, TaxTrans
'      TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd - SaveAmt)
'      Put THandle, 3744, TaxTrans
'      Get THandle, 2303, TaxTrans
'      TaxTrans.Revenue.Principle1Pd = SaveAmt
'      Put THandle, 2303, TaxTrans
'    End If
'
'    If x = 8744 Then
'      SaveAmt = TaxTrans.Amount
'      TaxTrans.BelongTo = 3925
'      TaxTrans.Description = "3925"
'      Put THandle, x, TaxTrans
'      Get THandle, 3531, TaxTrans
'      TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd - SaveAmt)
'      Put THandle, 3531, TaxTrans
'      Get THandle, 3925, TaxTrans
'      TaxTrans.Revenue.Principle1Pd = SaveAmt
'      Put THandle, 3925, TaxTrans
'    End If
'
'    If x = 8712 Then
'      SaveAmt = TaxTrans.Amount
'      TaxTrans.BelongTo = 457
'      TaxTrans.Description = "457"
'      Put THandle, x, TaxTrans
'      Get THandle, 587, TaxTrans
'      TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd - SaveAmt)
'      Put THandle, 587, TaxTrans
'      Get THandle, 457, TaxTrans
'      TaxTrans.Revenue.Principle1Pd = SaveAmt
'      Put THandle, 457, TaxTrans
'    End If
'
'    If x = 8716 Then
'      SaveAmt = TaxTrans.Amount
'      TaxTrans.BelongTo = 3895
'      TaxTrans.Description = "3895"
'      Put THandle, x, TaxTrans
'      Get THandle, 3355, TaxTrans
'      TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd - SaveAmt)
'      Put THandle, 3355, TaxTrans
'      Get THandle, 3895, TaxTrans
'      TaxTrans.Revenue.Principle1Pd = SaveAmt
'      Put THandle, 3895, TaxTrans
'    End If
'
'    Get CHandle, 3578, TaxCust
'    NextRec = TaxCust.LastTrans
'    Do While NextRec > 0
'      Get THandle, NextRec, TaxTrans
'      If NextRec = 8883 Then
'        GetBill = TaxTrans.BelongTo
'        TaxTrans.BelongTo = 1841 'change from 453 to 1881
'        ThisAmt = TaxTrans.Amount
'        Princ = TaxTrans.Revenue.Principle1Pd
'        Adv = TaxTrans.Revenue.CollectionPd
'        LateList = TaxTrans.Revenue.LateListPd
'        Interest = TaxTrans.Revenue.InterestPd
'        Opt1 = TaxTrans.Revenue.RevOpt1Pd
'        Opt2 = TaxTrans.Revenue.RevOpt2Pd
'        Opt3 = TaxTrans.Revenue.RevOpt3Pd
'        TaxTrans.Description = "1049 1841"
'        Put THandle, 8883, TaxTrans
'        Get THandle, GetBill, TaxTrans
'        TaxTrans.Revenue.Principle1Pd = OldRound#(TaxTrans.Revenue.Principle1Pd - Princ)
'        TaxTrans.Revenue.CollectionPd = OldRound#(TaxTrans.Revenue.CollectionPd - Adv)
'        TaxTrans.Revenue.LateListPd = OldRound#(TaxTrans.Revenue.LateListPd - LateList)
'        TaxTrans.Revenue.InterestPd = OldRound#(TaxTrans.Revenue.InterestPd - Interest)
'        TaxTrans.Revenue.RevOpt1Pd = OldRound#(TaxTrans.Revenue.RevOpt1Pd - Opt1)
'        TaxTrans.Revenue.RevOpt2Pd = OldRound#(TaxTrans.Revenue.RevOpt2Pd - Opt2)
'        TaxTrans.Revenue.RevOpt3Pd = OldRound#(TaxTrans.Revenue.RevOpt3Pd - Opt3)
'        Put THandle, GetBill, TaxTrans
'        Get THandle, 1841, TaxTrans
'        TaxTrans.Revenue.Principle1Pd = Princ
'        TaxTrans.Revenue.CollectionPd = Adv
'        TaxTrans.Revenue.LateListPd = LateList
'        TaxTrans.Revenue.InterestPd = Interest
'        TaxTrans.Revenue.RevOpt1Pd = Opt1
'        TaxTrans.Revenue.RevOpt2Pd = Opt2
'        TaxTrans.Revenue.RevOpt3Pd = Opt3
'        Put THandle, 1841, TaxTrans
'        Exit Do
'      End If
'      NextRec = TaxTrans.LastTrans
'    Loop
'   ^^^^^^^^^^^Fix for Harrisburg's 1/5/06 errors^^^^^^^^^^^^

'    If TaxTrans.BelongTo = 3190 Then Stop
'    TaxTrans.TranType = TaxTrans.TranType
'    TaxTrans.Amount = TaxTrans.Amount
'    TaxTrans.CustomerRec = TaxTrans.CustomerRec
'    fix for sunset beach
'    If x = 11485 Or x = 11337 Or x = 11186 Or x = 11017 Or x = 10496 Or x = 8944 Then
'      SaveAmt = TaxTrans.Amount
'      Get THandle, 8272, TaxTrans
'      TaxTrans.Revenue.Interest = OldRound(TaxTrans.Revenue.Interest + SaveAmt)
'      Put THandle, 8272, TaxTrans
'    End If
'    If x = 11484 Or x = 11336 Or x = 11185 Or x = 11016 Or x = 10495 Or x = 8938 Then
'      SaveAmt = TaxTrans.Amount
'      Get THandle, 8269, TaxTrans
'      TaxTrans.Revenue.Interest = OldRound(TaxTrans.Revenue.Interest + SaveAmt)
'      Put THandle, 8269, TaxTrans
'    End If
'    If x = 11491 Or x = 11343 Or x = 11192 Or x = 11023 Or x = 10502 Or x = 8966 Then
'      SaveAmt = TaxTrans.Amount
'      Get THandle, 8285, TaxTrans
'      TaxTrans.Revenue.Interest = OldRound(TaxTrans.Revenue.Interest + SaveAmt)
'      Put THandle, 8285, TaxTrans
'    End If
'  Next x

'  Close CHandle
'  Close THandle

End Sub

Private Sub cmdProcess5_Click()
  Dim TaxCust As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim x As Long
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NewTaxTrans As TaxTransactionType
  Dim NewTHandle As Integer
  Dim NumOfTRecs As Long
  Dim NextRec As Long
  Dim Balance As Double
  Dim TransCnt As Long
  Dim FirstTrans As Boolean
  Dim TotCustBal As Double
  Dim BHandle As Integer
  Dim CnvtDate As Integer
  Dim CDateStr$
  Dim LastTrans As Long
  Dim TotAmt As Double
  
  If Exist("cnvtdate.dat") Then
    BHandle = FreeFile
    Open "cnvtdate.dat" For Input As BHandle
    Input #BHandle, CDateStr$
    Close BHandle
    CDateStr$ = MakeRegDate(CInt(CDateStr$))
    CnvtDate = Date2Num(CDateStr$)
  Else
    CnvtDate = 0
  End If

  OpenTaxCustFile CHandle, NumOfCRecs
  OpenTaxTransFile THandle, NumOfTRecs
  OpenNewTaxTransFile NewTHandle
  frmTaxShowPctComp.Label1 = "Clearing Negative Customer Balances"
  frmTaxShowPctComp.Show
  frmTaxMainMenu.cmdExit.Enabled = False
  For x = 1 To NumOfCRecs
    Get CHandle, x, TaxCust
    FirstTrans = True
    TotCustBal = 0
    TotAmt = 0
    If TaxCust.Deleted <> 0 Then
      TaxCust.LastTrans = 0
      Put CHandle, x, TaxCust
      GoTo Skip
    End If
    NextRec = TaxCust.LastTrans
    Do While NextRec > 0
      Get THandle, NextRec, TaxTrans
'      If TaxTrans.CustomerRec = 16 And TaxTrans.Revenue.Principle1 = 20.81 And TaxTrans.Revenue.Interest = 2.62 Then Stop
'      If NextRec = 43348 Then Stop
'      If TaxTrans.TransDate > CnvtDate Then GoTo KeepInQ
      If TaxTrans.TranType = 22 Then 'prepaid only is treated as a stand alone transaction
        Balance# = TaxTrans.Revenue.PrePaidAmt
        GoTo KeepInQ
      End If
      If TaxTrans.TranType = 21 Then 'prepaid with billpay is changed and saved to 22 and then treated
        'as a stand alone transaction
        TaxTrans.TranType = 22
        Balance# = TaxTrans.Revenue.PrePaidAmt
        GoTo KeepInQ
      End If
      If TaxTrans.TranType = 1 Then
        Balance# = OldRound#(TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
        Balance# = OldRound#(Balance# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
        Balance# = OldRound#(Balance# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
        Balance# = OldRound#(Balance# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
        Balance# = OldRound#(Balance# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt))
        If Balance <= 0 Then GoTo LoopAgain
'        TotCustBal = OldRound#(TotCustBal + Balance#)
KeepInQ:
        TotAmt = OldRound(TotAmt + TaxTrans.Amount)
        NewTaxTrans.TransDate = TaxTrans.TransDate 'Date2Num(Date)
        NewTaxTrans.TaxYear = TaxTrans.TaxYear
        NewTaxTrans.BillType = "C" 'TaxTrans.BillType           'R=Real P=Personal Property C=Combined (NC/GA)
        NewTaxTrans.TranType = TaxTrans.TranType                '1=Bill 2=Payment 3=Release 4=Interest 5=Penalty 6=Collection/Ad Cost Billing
'        If TaxTrans.TransDate <= CnvtDate Then
          NewTaxTrans.Amount = Balance# ' OldRound#(TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1)
'        Else
'          If TaxTrans.TranType = 22 Then
'            NewTaxTrans.Amount = Balance
'          Else
'            NewTaxTrans.Amount = TaxTrans.Amount
'          End If
'        End If
        NewTaxTrans.Revenue.Principle1 = TaxTrans.Revenue.Principle1
        NewTaxTrans.Revenue.Principle2 = TaxTrans.Revenue.Principle2
        NewTaxTrans.Revenue.Principle3 = TaxTrans.Revenue.Principle3
        NewTaxTrans.Revenue.Principle4 = TaxTrans.Revenue.Principle4
        NewTaxTrans.Revenue.Principle5 = TaxTrans.Revenue.Principle5
        NewTaxTrans.Revenue.Interest = TaxTrans.Revenue.Interest
        NewTaxTrans.Revenue.LateList = TaxTrans.Revenue.LateList
        NewTaxTrans.Revenue.RevOpt1 = TaxTrans.Revenue.RevOpt1
        NewTaxTrans.Revenue.RevOpt2 = TaxTrans.Revenue.RevOpt2
        NewTaxTrans.Revenue.RevOpt3 = TaxTrans.Revenue.RevOpt3
        NewTaxTrans.Revenue.Penalty = TaxTrans.Revenue.Penalty
        NewTaxTrans.Revenue.Collection = TaxTrans.Revenue.Collection
        NewTaxTrans.Revenue.Future1 = 0
        NewTaxTrans.Revenue.Future2 = 0
        NewTaxTrans.Revenue.PrePaidAmt = TaxTrans.Revenue.PrePaidAmt
        NewTaxTrans.Revenue.PrePaidUsed = TaxTrans.Revenue.PrePaidUsed
        NewTaxTrans.Revenue.PrePaidBal = TaxTrans.Revenue.PrePaidBal
        If TaxTrans.TranType = 22 Then
          NewTaxTrans.Revenue.Principle1Pd = 0
        Else
          NewTaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd
        End If
        NewTaxTrans.Revenue.Principle2Pd = TaxTrans.Revenue.Principle2Pd
        NewTaxTrans.Revenue.Principle3Pd = TaxTrans.Revenue.Principle3Pd
        NewTaxTrans.Revenue.Principle4Pd = TaxTrans.Revenue.Principle4Pd
        NewTaxTrans.Revenue.Principle5Pd = TaxTrans.Revenue.Principle5Pd
        NewTaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd
        NewTaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd
        NewTaxTrans.Revenue.CollectionPd = TaxTrans.Revenue.CollectionPd
        NewTaxTrans.Revenue.LateListPd = TaxTrans.Revenue.LateListPd
        NewTaxTrans.Revenue.Future1Pd = 0
        NewTaxTrans.Revenue.Future2Pd = 0
        NewTaxTrans.Revenue.RevOpt1Pd = TaxTrans.Revenue.RevOpt1Pd
        NewTaxTrans.Revenue.RevOpt2Pd = TaxTrans.Revenue.RevOpt2Pd
        NewTaxTrans.Revenue.RevOpt3Pd = TaxTrans.Revenue.RevOpt3Pd
        NewTaxTrans.Revenue.pad = ""
        NewTaxTrans.DiscXDate = TaxTrans.DiscXDate
        NewTaxTrans.DiscAmt = TaxTrans.DiscAmt
        NewTaxTrans.OperNum = TaxTrans.OperNum
        NewTaxTrans.InternalPin = TaxTrans.InternalPin
        NewTaxTrans.RealPin = TaxTrans.RealPin
        NewTaxTrans.InternalPin = TaxTrans.InternalPin
        NewTaxTrans.PersPin = TaxTrans.PersPin
        NewTaxTrans.FromPrePay = QPTrim$(TaxTrans.FromPrePay)
        If TaxTrans.TransDate <= CnvtDate Then
          NewTaxTrans.Description = "Initialize " + CStr(TransCnt + 1)
        Else
          NewTaxTrans.Description = TaxTrans.Description
        End If
        NewTaxTrans.Altered = TaxTrans.Altered
        NewTaxTrans.Posted2GL = TaxTrans.Posted2GL
        NewTaxTrans.CustomerRec = x
        NewTaxTrans.BelongTo = TaxTrans.BelongTo
        NewTaxTrans.Padding = ""
        NewTaxTrans.CntyPara = TaxTrans.CntyPara
        NewTaxTrans.CustPin = x
        NewTaxTrans.DMVSubmitted = TaxTrans.DMVSubmitted
        NewTaxTrans.DMVBatch = TaxTrans.DMVBatch
        NewTaxTrans.TShpPara = TaxTrans.TShpPara
        If FirstTrans = True Then
          FirstTrans = False
          NewTaxTrans.LastTrans = 0
        Else
          NewTaxTrans.LastTrans = TransCnt
        End If
        TransCnt = TransCnt + 1
        Put NewTHandle, TransCnt, NewTaxTrans
      End If
LoopAgain:
      NextRec = TaxTrans.LastTrans
    Loop
    If TotAmt = 0 Then
      TaxCust.LastTrans = 0
      Put CHandle, x, TaxCust
    Else
      TaxCust.LastTrans = TransCnt
      Put CHandle, x, TaxCust
    End If
Skip:
    frmTaxShowPctComp.ShowPctComp x, NumOfCRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      frmTaxMainMenu.cmdExit.Enabled = True
      Exit Sub
    End If
  Next x
  Unload frmTaxShowPctComp
  frmTaxMainMenu.cmdExit.Enabled = True
  Close
  
  KillFile App.Path + "\OLDWINTAXTRANS.DAT"
  Name App.Path + "\TAXTRANS.DAT" As App.Path + "\OLDWINTAXTRANS.DAT"
  Name App.Path + "\NEWTAXTRANS.DAT" As App.Path + "\TAXTRANS.DAT"
'  Call MakeAmtEqualRevsAfterClearingNegs
  Call TaxMsg(800, "Balance update has completed successfully. The old transaction file is now saved as 'OLDWINTAXTRANS.DAT'.")
  
  Exit Sub
  
CheckBillBal:

  Balance# = OldRound#(TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
  Balance# = OldRound#(Balance# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
  Balance# = OldRound#(Balance# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
  Balance# = OldRound#(Balance# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
  Balance# = OldRound#(Balance# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt))

  Return

End Sub

Private Sub MakeAmtEqualRevs()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim x As Long
  Dim TotRev As Double
  
  If Exist("negscleared.dat") Then
    Call TaxMsg(900, "This procedure has already been successfully executed.")
    Exit Sub
  End If
  
  ReDim TypeCnt(1 To 17) As Long
  
  frmTaxShowPctComp.Label1 = "Making Amounts Equal Revenues"
  frmTaxShowPctComp.Show
  frmTaxMainMenu.cmdExit.Enabled = False
  OpenTaxTransFile THandle, NumOfTRecs
  For x = 1 To NumOfTRecs
    Get THandle, x, TaxTrans
    Select Case TaxTrans.TranType
      Case 1 'billing
        TotRev = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.LateList) '+ TaxTrans.Revenue.Interest + TaxTrans.Revenue.Collection)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(1) = TypeCnt(1) + 1
        End If
        TotRev = 0
      Case 2 'payment
        TotRev = OldRound(TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.LateListPd + TaxTrans.Revenue.CollectionPd)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd + TaxTrans.Revenue.InterestPd)
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(2) = TypeCnt(2) + 1
        End If
        TotRev = 0
      Case 3 'release
        TotRev = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Collection)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5 + TaxTrans.Revenue.Interest)
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(3) = TypeCnt(3) + 1
        End If
        TotRev = 0
      Case 4 'Interest
        TotRev = TaxTrans.Revenue.Interest
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(4) = TypeCnt(4) + 1
        End If
        TotRev = 0
      Case 5 'Penalty
        TotRev = TaxTrans.Revenue.Penalty
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(5) = TypeCnt(5) + 1
        End If
        TotRev = 0
      Case 6 'Advertising
        TotRev = TaxTrans.Revenue.Collection
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(6) = TypeCnt(6) + 1
        End If
        TotRev = 0
      Case 7 'adjust pay down
        TotRev = OldRound(TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.LateListPd + TaxTrans.Revenue.CollectionPd)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd + TaxTrans.Revenue.InterestPd)
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(7) = TypeCnt(7) + 1
        End If
        TotRev = 0
      Case 9 'credit applied at billing
        TotRev = TaxTrans.Revenue.PrePaidUsed
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(8) = TypeCnt(8) + 1
        End If
        TotRev = 0
      Case 13 'adjust bill down
        TotRev = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Collection)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5 + TaxTrans.Revenue.Interest)
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(9) = TypeCnt(9) + 1
        End If
        TotRev = 0
      Case 14 'adjust bill up
        TotRev = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Collection)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5 + TaxTrans.Revenue.Interest)
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(10) = TypeCnt(10) + 1
        End If
        TotRev = 0
      Case 21 'bill pay/overpay
        TotRev = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Collection)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5 + TaxTrans.Revenue.Interest)
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(11) = TypeCnt(11) + 1
        End If
        TotRev = 0
      Case 22
        'not needed
      Case 24 'adjust bill up affecting credit
        TotRev = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Collection)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5 + TaxTrans.Revenue.Interest)
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(12) = TypeCnt(12) + 1
        End If
        TotRev = 0
      Case 10 'adjust bill down affecting credit
        TotRev = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Collection)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5 + TaxTrans.Revenue.Interest)
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(13) = TypeCnt(13) + 1
        End If
        TotRev = 0
      Case 11 'adjust prepay down
        'not needed
      Case 12 'refund prepay
        'not needed
      Case Else
    End Select
    frmTaxShowPctComp.ShowPctComp x, NumOfTRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      frmTaxMainMenu.cmdExit.Enabled = True
      Exit Sub
    End If
  Next x
      
  Unload frmTaxShowPctComp
  frmTaxMainMenu.cmdExit.Enabled = True
  Close
  
  For x = 1 To 17
    If TypeCnt(x) > 0 Then
      Exit For
    End If
  Next x
  
  If x <= 17 Then
    frmTaxAmtToRevsList.Show vbModal
  Else
    Call TaxMsg(900, "There were no transactions that needed this repair.")
  End If
  
End Sub

Private Sub MakeAmtEqualRevsAfterClearingNegs()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim x As Long
  Dim TotRev As Double
  Dim One As Integer
  Dim ThisFile As Integer
  Dim FileName$
  
  ReDim TypeCnt(1 To 17) As Long
  
  frmTaxShowPctComp.Label1 = "Making Amounts Equal Revenues"
  frmTaxShowPctComp.Show
  frmTaxMainMenu.cmdExit.Enabled = False
  OpenTaxTransFile THandle, NumOfTRecs
  For x = 1 To NumOfTRecs
    Get THandle, x, TaxTrans
    If TaxTrans.TranType <> 1 Then
      Call TaxMsg(900, "ERROR: There is a transaction type that is not a billing transaction. Please call Southern Software at 1-800-842-8190 for assistance.")
      Exit For
    End If
  Next x
  If x <= NumOfTRecs Then
    Close
    Exit Sub
  End If
  
  For x = 1 To NumOfTRecs
    Get THandle, x, TaxTrans
    TaxTrans.CustomerRec = TaxTrans.CustomerRec
    Select Case TaxTrans.TranType
      Case 1 'billing
        TotRev = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Collection)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
        TotRev = OldRound(TotRev - OldRound(TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.LateListPd + TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.CollectionPd))
        TotRev = OldRound(TotRev - OldRound(TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd))
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(1) = TypeCnt(1) + 1
        End If
        TotRev = 0
      Case Else
    End Select
    frmTaxShowPctComp.ShowPctComp x, NumOfTRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      frmTaxMainMenu.cmdExit.Enabled = True
      Exit Sub
    End If
  Next x
      
  Unload frmTaxShowPctComp
  frmTaxMainMenu.cmdExit.Enabled = True
  Close
  For x = 1 To 17
    If TypeCnt(x) > 0 Then
      Exit For
    End If
  Next x
  If x <= 17 Then
    frmTaxAmtToRevsList.Show vbModal
  Else
    Call TaxMsg(900, "There were no transactions that needed this repair.")
  End If
  
  FileName = "negscleared.dat"
  ThisFile = FreeFile
  Open FileName For Output As ThisFile
  One = 1
  Print #ThisFile, One
  Close ThisFile
  
End Sub

Private Sub FixMaggieValley()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim x As Long
  
  OpenTaxTransFile THandle, NumOfTRecs
  'for customer #368
  Get THandle, 7622, TaxTrans
  TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd - 0.42)
  Put THandle, 7622, TaxTrans
  
  'for customer #2093
  Get THandle, 26347, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 206.22
  TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd - 7.22
  Put THandle, 26347, TaxTrans
  
  'for customer #317
  Get THandle, 41779, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 127.01
  Put THandle, 41779, TaxTrans
  
  'for customer #396
  Get THandle, 861, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 13.89
  Put THandle, 861, TaxTrans
  
  'for customer #276
  Get THandle, 965, TaxTrans
  TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd - 0.03)
  TaxTrans.Amount = TaxTrans.Amount
  Put THandle, 965, TaxTrans
  
  Close THandle
  
End Sub

Private Sub FixSunsetBeach()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim x As Long
  
  OpenTaxTransFile THandle, NumOfTRecs
  'for customer #4128
  Get THandle, 11486, TaxTrans
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Amount = 0
  Put THandle, 11486, TaxTrans

  Get THandle, 11338, TaxTrans
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Amount = 0
  Put THandle, 11338, TaxTrans

  Get THandle, 11187, TaxTrans
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Amount = 0
  Put THandle, 11187, TaxTrans

  Get THandle, 11018, TaxTrans
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Amount = 0
  Put THandle, 11018, TaxTrans

  Get THandle, 10497, TaxTrans
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Amount = 0
  Put THandle, 10497, TaxTrans

  Get THandle, 8946, TaxTrans
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Amount = 0
  Put THandle, 8946, TaxTrans

  'for customer #4149
  Get THandle, 11492, TaxTrans
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Amount = 0
  Put THandle, 11492, TaxTrans

  Get THandle, 11344, TaxTrans
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Amount = 0
  Put THandle, 11344, TaxTrans

  Get THandle, 11193, TaxTrans
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Amount = 0
  Put THandle, 11193, TaxTrans

  Get THandle, 11024, TaxTrans
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Amount = 0
  Put THandle, 11024, TaxTrans

  Get THandle, 10503, TaxTrans
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Amount = 0
  Put THandle, 10503, TaxTrans

  Get THandle, 8967, TaxTrans
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Amount = 0
  Put THandle, 8967, TaxTrans
  
  Close THandle
End Sub

Private Sub FixCalabash()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim x As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim RealPropRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim NextRec As Long
  
  OpenRealPropFile RHandle, NumOfRealRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  Get TCHandle, 1, TaxCust
  NextRec = TaxCust.FirstPropRec
  Get RHandle, NextRec, RealPropRec
  RealPropRec.NextRec = 0
  RealPropRec.CustPin = 0
  Put RHandle, NextRec, RealPropRec
  TaxCust.FirstPropRec = 1717
  Put TCHandle, x, TaxCust
  Close
  MsgBox ("Finished.")
  Exit Sub
  
'  OpenTaxTransFile THandle, NumOfTRecs
  
  
  'for customer# 1517
'  Get THandle, 23119, TaxTrans
'  TaxTrans.TaxYear = 2005
'  TaxTrans.BelongTo = 0
'  TaxTrans.CustPin = 0
'  TaxTrans.CustomerRec = 1517
'  Put THandle, 23119, TaxTrans
  
  'for cust# 4
'  Get THandle, 25840, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.InterestPd = 0
'  TaxTrans.BelongTo = 0
'  TaxTrans.TaxYear = 0
'  TaxTrans.Description = "Deleted"
'  Put THandle, 25840, TaxTrans

  'for cust# 1536
'  Get THandle, 25448, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.InterestPd = 0
'  TaxTrans.BelongTo = 0
'  TaxTrans.TaxYear = 0
'  TaxTrans.Description = "Deleted"
'  Put THandle, 25448, TaxTrans
'  Close
  Get THandle, 23132, TaxTrans
  TaxTrans.LastTrans = NumOfTRecs + 1
  Put THandle, 23132, TaxTrans
  
  TaxTrans.TransDate = Date2Num("02/25/2005")
  TaxTrans.TranType = 2
  TaxTrans.Revenue.Principle1Pd = 19.45
  TaxTrans.Revenue.InterestPd = 0.54
  TaxTrans.Revenue.CollectionPd = 0
  TaxTrans.Revenue.LateListPd = 0
  TaxTrans.Revenue.RevOpt1Pd = 0
  TaxTrans.Revenue.RevOpt2Pd = 0
  TaxTrans.Revenue.RevOpt3Pd = 0
  TaxTrans.CustPin = 0
  TaxTrans.DiscXDate = 0
  TaxTrans.RealPin = 0
  TaxTrans.PersPin = 0
  TaxTrans.Posted2GL = "N"
  TaxTrans.TaxYear = 2004
  TaxTrans.DiscAmt = 0
  TaxTrans.OperNum = 0
  TaxTrans.Amount = 19.99
  TaxTrans.Description = "Payment Inserted(SS) Bill #40"
  TaxTrans.CustomerRec = 1313
  TaxTrans.BelongTo = 17132
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.InternalPin = 0
  TaxTrans.LastTrans = NumOfTRecs + 2
  Put THandle, NumOfTRecs + 1, TaxTrans
  
  TaxTrans.TransDate = Date2Num("02/25/2005")
  TaxTrans.TranType = 2
  TaxTrans.Revenue.Principle1Pd = 19.45
  TaxTrans.Revenue.InterestPd = 2.34
  TaxTrans.Revenue.CollectionPd = 3#
  TaxTrans.Revenue.LateListPd = 0
  TaxTrans.Revenue.RevOpt1Pd = 0
  TaxTrans.Revenue.RevOpt2Pd = 0
  TaxTrans.Revenue.RevOpt3Pd = 0
  TaxTrans.CustPin = 0
  TaxTrans.DiscXDate = 0
  TaxTrans.RealPin = 0
  TaxTrans.PersPin = 0
  TaxTrans.Posted2GL = "N"
  TaxTrans.TaxYear = 2003
  TaxTrans.DiscAmt = 0
  TaxTrans.OperNum = 0
  TaxTrans.Amount = 24.79
  TaxTrans.Description = "Payment Inserted(SS) Bill #35"
  TaxTrans.CustomerRec = 1313
  TaxTrans.BelongTo = 10604
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.InternalPin = 0
  TaxTrans.LastTrans = NumOfTRecs + 3
  Put THandle, NumOfTRecs + 2, TaxTrans
  
  TaxTrans.TransDate = Date2Num("02/25/2005")
  TaxTrans.TranType = 2
  TaxTrans.Revenue.Principle1Pd = 11.78
  TaxTrans.Revenue.InterestPd = 2.41
  TaxTrans.Revenue.CollectionPd = 0
  TaxTrans.Revenue.LateListPd = 0
  TaxTrans.Revenue.RevOpt1Pd = 0
  TaxTrans.Revenue.RevOpt2Pd = 0
  TaxTrans.Revenue.RevOpt3Pd = 0
  TaxTrans.CustPin = 0
  TaxTrans.DiscXDate = 0
  TaxTrans.RealPin = 0
  TaxTrans.PersPin = 0
  TaxTrans.Posted2GL = "N"
  TaxTrans.TaxYear = 2002
  TaxTrans.DiscAmt = 0
  TaxTrans.OperNum = 0
  TaxTrans.Amount = 14.19
  TaxTrans.Description = "Payment Inserted(SS) Bill #33"
  TaxTrans.CustomerRec = 1313
  TaxTrans.BelongTo = 5238
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.InternalPin = 0
  TaxTrans.LastTrans = 21079
  Put THandle, NumOfTRecs + 3, TaxTrans
  
  MsgBox ("Fix completed.")
  
End Sub

Private Sub FindOddRealOwner()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  Dim x As Long, NextRec As Long, y As Integer
  Dim RealRec As PropertyRecType
  Dim RealCnt As Integer
  Dim ThisPin As String
  Dim FoundCnt As Integer
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxPropFile RHandle, NumOfRRecs
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TaxCust.FirstPropRec = 0 Then GoTo SkipIt
    ReDim ThesePins(1 To 1) As String
    RealCnt = 0
    NextRec = TaxCust.FirstPropRec
    Do While NextRec > 0
      Get RHandle, NextRec, RealRec
      RealCnt = RealCnt + 1
      ReDim Preserve ThesePins(1 To RealCnt) As String
      ThesePins(RealCnt) = QPTrim$(RealRec.RealPin)
      ThisPin = ThesePins(RealCnt)
      FoundCnt = 0
      For y = 1 To RealCnt
        If ThisPin = ThesePins(y) Then
          FoundCnt = FoundCnt + 1
'          If FoundCnt > 1 Then Stop
        End If
      Next y
      NextRec = RealRec.NextRec
    Loop
'
SkipIt:
  Next x
  
'  For x = 1 To NumOfRRecs
'    Get RHandle, x, RealRec
'    Get TCHandle, RealRec.CustPin, TaxCust
'    NextRec = TaxCust.FirstPropRec
'    Do While NextRec > 0
'      Get RHandle, NextRec, RealRec
'      If RealRec.CustPin <> TaxCust.Acct Then Stop
'      NextRec = RealRec.NextRec
'    Loop
'  Next x
   
  Close
  Call TaxMsg(900, "Real owner search completed.")
End Sub

Private Sub AddFireTaxToHarrisburgPersPropBills()
  Dim TransRec As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TAHandle As Integer
  Dim NumOfTARecs As Long
  Dim TaxAdjTrans As TaxTransactionType
  Dim x As Long
  Dim NextRec As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim BillNum$, AdjAmt#
  Dim NextTransRec&
  Dim CreditAmt As Double
  Dim CreditBalance As Double
  Dim ThisAmt As Double
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Long
  Dim CompPin As String
  Dim ThisVal As Double
  Dim y As Long
  Dim ThisDate As Integer
  Dim TransPin$
  Dim PersPin$, RealPin$
  Dim NewBalThisBill#
  Dim PersCnt As Long
  Dim RealCnt As Long
  Dim ThisYear As Integer

  ThisYear = Val(Right$(Date$, 4))
  Print ThisYear
  
'  Exit Sub
'  Dim ErrorCnt As Integer
'  ReDim Err(1 To 1) As Long

  OpenRealPropFile RHandle, NumOfRRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    NextRec = TaxCust.FirstPropRec
    Do While NextRec > 0
      Get RHandle, NextRec, RealRec
      If RealRec.OptRev1Chrg <> 1 Then
        RealRec.OptRev1Chrg = 1
        Put RHandle, NextRec, RealRec
      End If
      NextRec = RealRec.NextRec
    Loop
  Next x
  
  OpenPersPropFile PHandle, NumOfPRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxTransFile TAHandle, NumOfTARecs
  frmTaxShowPctComp.Label1 = "Fixing Harrisburg's Fire Tax"
  frmTaxShowPctComp.Show
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
'    If TaxCust.FirstPropRec > 0 Then GoTo SkipIt
    NextRec = TaxCust.FirstPersRec
    Do While NextRec > 0
      Get PHandle, NextRec, PersRec
      ThisVal = OldRound(PersRec.MCVALUE + PersRec.CVALUE + PersRec.MHVALUE + PersRec.MTVALUE + PersRec.PersVal - (PersRec.EXMPOTHR + PersRec.EXMPSENI))
      CompPin = QPTrim$(PersRec.PropPin)
      For y = 1 To NumOfTTRecs
        Get TTHandle, y, TransRec
'        If TransRec.TaxYear < 2006 Then GoTo SkipIt
'        If TransRec.TaxYear < 2008 Then GoTo SkipIt
        'If TransRec.TaxYear < 2009 Then GoTo SkipIt
        If TransRec.TranType <> 1 Then GoTo SkipIt
        If TransRec.TaxYear < 2013 Then GoTo SkipIt
        TransPin = QPTrim$(TransRec.PersPin)
        If CompPin = TransPin Then
          AdjAmt = OldRound(ThisVal * 0.001115)
          PersPin = QPTrim$(PersRec.PropPin)
          GoSub AdjustBillUpPers
          PersCnt = PersCnt + 1
          Exit For
        End If
SkipIt:
      Next y
      NextRec = PersRec.NextRec
    Loop
    NextRec = TaxCust.FirstPropRec
    Do While NextRec > 0
      Get RHandle, NextRec, RealRec
      ThisVal = OldRound(RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI))
      CompPin = QPTrim$(RealRec.RealPin)
      For y = 1 To NumOfTTRecs
        Get TTHandle, y, TransRec
'        If TransRec.TaxYear < 2006 Then GoTo SkipItR
'        If TransRec.TaxYear < 2008 Then GoTo SkipItR
        If TransRec.TranType <> 1 Then GoTo SkipItR
        If TransRec.TaxYear < ThisYear Then GoTo SkipItR
        TransPin = TransRec.RealPin
        If TransRec.Revenue.RevOpt1 > 0 Then GoTo SkipItR
        If CompPin = TransPin Then
          AdjAmt = OldRound(ThisVal * 0.00075)
          RealPin = QPTrim$(RealRec.RealPin)
          GoSub AdjustBillUpReal
          RealCnt = RealCnt + 1
          Exit For
        End If
SkipItR:
      Next y
      NextRec = RealRec.NextRec
    Loop
NextCust:
    
    frmTaxShowPctComp.Label1 = "Customer: " + CStr(x)
    DoEvents
    If x Mod 10 = 0 Then
      frmTaxShowPctComp.ShowPctComp x, NumOfTCRecs
    End If
    
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      Exit Sub
    End If
  Next x
  Unload frmTaxShowPctComp
'  frmTaxMainMenu.cmdExit.Enabled = True
  
  Close
  Call TaxMsg(800, "A total of " + CStr(PersCnt) + " personal bills and " + CStr(RealCnt) + " real bills have been adjusted successfully. ")
  Exit Sub
  
AdjustBillUpPers:
  TaxAdjTrans.TransDate = Date2Num(Date)
  CreditAmt = 0 'CDbl(fpCurrPrepayBal.Value)
  If CreditAmt <= 0 Then
    TaxAdjTrans.TranType = 14 'adjust bill up with no affect on credit balance
    TaxAdjTrans.Revenue.RevOpt1 = AdjAmt#
    TaxAdjTrans.Amount = AdjAmt#
    TaxAdjTrans.CustomerRec = x
    TaxAdjTrans.LastTrans = TaxCust.LastTrans
    TaxAdjTrans.BelongTo = y
    Get #TTHandle, y, TransRec

    BillNum$ = CompPin
    TaxAdjTrans.Description = "Tax Adj Bill Up #" + BillNum$
    TransRec.Revenue.RevOpt1 = OldRound(TransRec.Revenue.RevOpt1 + AdjAmt#)
    NewBalThisBill = 0
    NewBalThisBill = OldRound(TransRec.Revenue.Collection + TransRec.Revenue.Future1 + TransRec.Revenue.Future2)
    NewBalThisBill = OldRound(NewBalThisBill + TransRec.Revenue.Interest + TransRec.Revenue.LateList + TransRec.Revenue.Penalty)
    NewBalThisBill = OldRound(NewBalThisBill + TransRec.Revenue.Principle1 + TransRec.Revenue.Principle2 + TransRec.Revenue.Principle3)
    NewBalThisBill = OldRound(NewBalThisBill + TransRec.Revenue.Principle4 + TransRec.Revenue.Principle5 + TransRec.Revenue.RevOpt1)
    NewBalThisBill = OldRound(NewBalThisBill + TransRec.Revenue.RevOpt2 + TransRec.Revenue.RevOpt3)
    NewBalThisBill = OldRound(NewBalThisBill - (TransRec.Revenue.CollectionPd + TransRec.Revenue.Future1Pd + TransRec.Revenue.Future2Pd))
    NewBalThisBill = OldRound(NewBalThisBill - (TransRec.Revenue.Principle1Pd + TransRec.Revenue.Principle2Pd + TransRec.Revenue.Principle3Pd))
    NewBalThisBill = OldRound(NewBalThisBill - (TransRec.Revenue.InterestPd + TransRec.Revenue.LateListPd + TransRec.Revenue.PenaltyPd))
    NewBalThisBill = OldRound(NewBalThisBill - (TransRec.Revenue.Principle4Pd + TransRec.Revenue.Principle5Pd + TransRec.Revenue.RevOpt1Pd))
    NewBalThisBill = OldRound(NewBalThisBill - (TransRec.Revenue.RevOpt2Pd + TransRec.Revenue.RevOpt3Pd))
  
    TaxAdjTrans.TaxYear = TransRec.TaxYear
    TaxAdjTrans.RealPin = "0"
    TaxAdjTrans.PersPin = PersPin
    TaxAdjTrans.CustPin = TaxCust.PIN
    TaxAdjTrans.OperNum = 0
    Put #TTHandle, y, TransRec
PrePayOnly:
    NextTransRec& = (LOF(TTHandle) / Len(TransRec)) + 1

    TaxCust.LastTrans = NextTransRec&

    Put #TAHandle, NextTransRec&, TaxAdjTrans
    Put #TCHandle, x, TaxCust
  
  End If
  
  Return

AdjustBillUpReal:
  TaxAdjTrans.TransDate = Date2Num(Date)
  CreditAmt = 0 'CDbl(fpCurrPrepayBal.Value)
  If CreditAmt <= 0 Then
    TaxAdjTrans.TranType = 14 'adjust bill up with no affect on credit balance
    TaxAdjTrans.Revenue.RevOpt1 = AdjAmt#
    TaxAdjTrans.Amount = AdjAmt#
    TaxAdjTrans.CustomerRec = x
    TaxAdjTrans.LastTrans = TaxCust.LastTrans
    TaxAdjTrans.BelongTo = y
    Get #TTHandle, y, TransRec

    BillNum$ = ParseBillNum(TransRec.Description) '  CompPin
    TaxAdjTrans.Description = "Tax Adj Bill Up #" + BillNum$
    TransRec.Revenue.RevOpt1 = OldRound(TransRec.Revenue.RevOpt1 + AdjAmt#)
    NewBalThisBill = 0
    NewBalThisBill = OldRound(TransRec.Revenue.Collection + TransRec.Revenue.Future1 + TransRec.Revenue.Future2)
    NewBalThisBill = OldRound(NewBalThisBill + TransRec.Revenue.Interest + TransRec.Revenue.LateList + TransRec.Revenue.Penalty)
    NewBalThisBill = OldRound(NewBalThisBill + TransRec.Revenue.Principle1 + TransRec.Revenue.Principle2 + TransRec.Revenue.Principle3)
    NewBalThisBill = OldRound(NewBalThisBill + TransRec.Revenue.Principle4 + TransRec.Revenue.Principle5 + TransRec.Revenue.RevOpt1)
    NewBalThisBill = OldRound(NewBalThisBill + TransRec.Revenue.RevOpt2 + TransRec.Revenue.RevOpt3)
    NewBalThisBill = OldRound(NewBalThisBill - (TransRec.Revenue.CollectionPd + TransRec.Revenue.Future1Pd + TransRec.Revenue.Future2Pd))
    NewBalThisBill = OldRound(NewBalThisBill - (TransRec.Revenue.Principle1Pd + TransRec.Revenue.Principle2Pd + TransRec.Revenue.Principle3Pd))
    NewBalThisBill = OldRound(NewBalThisBill - (TransRec.Revenue.InterestPd + TransRec.Revenue.LateListPd + TransRec.Revenue.PenaltyPd))
    NewBalThisBill = OldRound(NewBalThisBill - (TransRec.Revenue.Principle4Pd + TransRec.Revenue.Principle5Pd + TransRec.Revenue.RevOpt1Pd))
    NewBalThisBill = OldRound(NewBalThisBill - (TransRec.Revenue.RevOpt2Pd + TransRec.Revenue.RevOpt3Pd))
    TaxAdjTrans.TaxYear = TransRec.TaxYear
    TaxAdjTrans.PersPin = "0"
    TaxAdjTrans.RealPin = RealPin
    TaxAdjTrans.CustPin = TaxCust.PIN
    TaxAdjTrans.OperNum = 0
    Put #TTHandle, y, TransRec
PrePayOnlyR:
    NextTransRec& = (LOF(TTHandle) / Len(TransRec)) + 1

    TaxCust.LastTrans = NextTransRec&

    Put #TAHandle, NextTransRec&, TaxAdjTrans
    Put #TCHandle, x, TaxCust
  
  End If
  
  Return

End Sub

Private Sub AddLateListTaxToHarrisburg()
  Dim ColCnt As Integer
  Dim ThisCol As Integer
  Dim ThisPos As Integer
  Dim TextLine$
  Dim ThisFile$
  Dim LHandle As Integer
  Dim TextLen As Integer
  Dim Thisch As String
  Dim ThisWord$
  Dim FirstLine As Boolean
  Dim RecCnt As Long
  Dim x As Long, y As Long, z As Long, q As Long
  Dim WordCnt As Integer
  Dim dlm$
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  Dim RCnt As Long
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Long
  Dim PCnt As Long
  Dim LineCnt As Long
  Dim CountyNum As Long
  Dim LLAmt#
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim ThisPin$
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim NextRec As Long
  Dim TF462 As Boolean
  Dim TF3273 As Boolean
  Dim TF671 As Boolean
'  Dim FoundIt As Boolean
  
  Dim ThisYear As Integer

  ThisYear = Val(Right$(Date$, 4))
  Print ThisYear
  
  TF462 = False
  TF3273 = False
  TF671 = False
  
  dlm = "~"
  
'  If TaxMsgWOpts(700, "Be sure to put the pin number in the first column and the optional search data in the second column. Name the .csv file 'HBRealOpt.csv'.", "Continue", "Abort") = "abort" Then
'    Exit Sub
'  End If
  If Exist("HBLateList.csv") Then
    OpenRealPropFile RHandle, NumOfRRecs
    OpenPersPropFile PHandle, NumOfPRecs
    OpenTaxTransFile TTHandle, NumOfTTRecs
    OpenTaxCustFile TCHandle, NumOfTCRecs
    LHandle = FreeFile
    ThisFile = "HBLateList.csv"
    Open ThisFile For Input As #LHandle
    Do
      Line Input #LHandle, TextLine
      LineCnt = LineCnt + 1
      If eof(LHandle) Then Exit Do
    Loop
    Close LHandle
    LHandle = FreeFile
    ThisFile = "HBLateList.csv"
    Open ThisFile For Input As #LHandle
    frmTaxShowPctComp.Label1 = "Adding Late List Tax To Harrisburg"
    frmTaxShowPctComp.Show
    DoEvents
    'Word(1) = County Nbr, Word(2) = Late List Val
    For z = 1 To LineCnt
      Line Input #LHandle, TextLine
      TextLen = Len(TextLine)
      TextLine = TextLine + dlm
      WordCnt = 0
      ReDim Words(1 To 2) As String
      For x = 1 To TextLen + 1
        Thisch = Mid(TextLine, x, 1)
        If Thisch = dlm Then
          WordCnt = WordCnt + 1
          ReDim Preserve Words(1 To WordCnt) As String
          Words(WordCnt) = ThisWord
          ThisWord = ""
          GoTo NewWord
        End If
        ThisWord = ThisWord + Thisch
NewWord:
      Next x
      LLAmt# = CDbl(Words(2))
      CountyNum = QPTrim$(Words(1))
      For x = 1 To NumOfTCRecs
        Get TCHandle, x, TaxCust
'        If x = 3313 Then Stop
'        FoundIt = False
'        TaxCust.CountyAcctString = TaxCust.CountyAcctString
        If TaxCust.CountyAcct = CLng(Words(1)) Or QPTrim$(TaxCust.CountyAcctString) = Words(1) Then
'        If TaxCust.CountyAcct = CLng(Words(1)) Then
        NextRec = NumOfTTRecs + 1
          For y = 1 To NumOfTTRecs
            Get TTHandle, y, TaxTrans
            If TaxCust.FirstPersRec > 0 Or TaxCust.FirstPropRec > 0 Then
              If TaxTrans.CustomerRec = TaxCust.Acct And TaxTrans.TaxYear = ThisYear And TaxTrans.TranType = 1 Then
'              If TaxTrans.CustomerRec = TaxCust.Acct And TaxTrans.TaxYear = 2008 And TaxTrans.TranType = 1 Then
'              If TaxTrans.CustomerRec = TaxCust.Acct And TaxTrans.TaxYear = 2007 And TaxTrans.TranType = 1 Then
'                If TaxTrans.CustomerRec = 3015 Then
'                  TaxTrans.Revenue.LateList = 162.2
'                  TaxTrans.Amount = OldRound(TaxTrans.Amount + 162.2)
'                  Put TTHandle, y, TaxTrans
'                  RCnt = RCnt + 1
'                ElseIf TaxTrans.CustomerRec = 671 Then
'                  If TF671 = False Then
'                    TaxTrans.Revenue.LateList = CDbl(Words(2))
'                    TaxTrans.Amount = OldRound(TaxTrans.Amount + CDbl(Words(2)))
'                    Put TTHandle, y, TaxTrans
'                    RCnt = RCnt + 1
'                    TF671 = True
'                  End If
'                ElseIf TaxTrans.CustomerRec = 3273 Then
'                  If TF3273 = False Then
'                    TaxTrans.Revenue.LateList = CDbl(Words(2))
'                    TaxTrans.Amount = OldRound(TaxTrans.Amount + CDbl(Words(2)))
'                    Put TTHandle, y, TaxTrans
'                    RCnt = RCnt + 1
'                    TF3273 = True
'                  End If
'                ElseIf TaxTrans.CustomerRec = 462 Then
'                  If TF462 = False Then
'                    TaxTrans.Revenue.LateList = CDbl(Words(2))
'                    TaxTrans.Amount = OldRound(TaxTrans.Amount + CDbl(Words(2)))
'                    Put TTHandle, y, TaxTrans
'                    RCnt = RCnt + 1
'                    TF462 = True
'                  End If
'                Else
                  TaxTrans.Revenue.LateList = CDbl(Words(2))
                  TaxTrans.Amount = OldRound(TaxTrans.Amount + CDbl(Words(2)))
                  Put TTHandle, y, TaxTrans
                  RCnt = RCnt + 1
'                  FoundIt = True
                End If
              End If
'            End If
           Next y
'           If FoundIt = False Then
'             Debug.Print CStr(x) & "~" & CStr(CountyNum) & "~" & QPTrim$(TaxCust.CustName)
'           End If
        End If
KeepGoing:
      Next x
      frmTaxShowPctComp.Label1 = "Customer: " + CStr(z)
      frmTaxShowPctComp.ShowPctComp z, LineCnt
      If frmTaxShowPctComp.Out = True Then
        Close
        frmTaxShowPctComp.Out = False
        Unload frmTaxShowPctComp
        Exit Sub
      End If
    Next z
  Else
    Call TaxMsg(900, "The file 'HBLateList.csv' cannot be found.")
    Exit Sub
  End If
  
  Close
  Call TaxMsg(800, CStr(RCnt) + " late listing taxes were updated successfully.")
End Sub

Private Sub FixWhiteLake()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim x As Long
  Dim PostDate As Integer
  Dim DiscDate As Integer
  Dim ThisDate$
  Dim TaxMaster As TaxMasterType
  Dim TMHandle As Integer
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMaster
  TaxMaster.DiscXDate = Date2Num("08/31/2006")
  Put TMHandle, 1, TaxMaster
  Close TMHandle
  
  PostDate = Date2Num("08/04/2006")
  DiscDate = Date2Num("08/31/2006")
  OpenTaxTransFile THandle, NumOfTRecs
  For x = 1 To NumOfTRecs
    Get THandle, x, TaxTrans
    If TaxTrans.TranType = 1 Then
      If TaxTrans.TransDate = PostDate Then
        TaxTrans.DiscXDate = DiscDate
        Put THandle, x, TaxTrans
      End If
    End If
  Next x
  
  Close THandle

End Sub

Private Sub FixWestJeffersonPers()
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim TaxCust As TaxCustType
  Dim THandle As Integer
  Dim NumOfTaxCustRecs As Long
  Dim WhatPers&, x As Long, y As Long
  Dim NextRec As Long
  Dim PCnt As Integer
  Dim PValue As Double
  Dim ThisRec As Long
  
  OpenTaxCustFile THandle, NumOfTaxCustRecs
  OpenPersPropFile PHandle, NumOfPersRecs
  For x = 1 To NumOfTaxCustRecs
    Get THandle, x, TaxCust
    NextRec = TaxCust.FirstPersRec
    PValue = 0
    Do While NextRec > 0
      Get PHandle, NextRec, PersRec
      PValue = OldRound(PValue + PersRec.CVALUE + PersRec.MCVALUE + PersRec.MHVALUE + PersRec.MTVALUE + PersRec.PersVal - PersRec.EXMPOTHR - PersRec.EXMPSENI)
      NextRec = PersRec.NextRec
    Loop
    If PValue = 0 Then
      TaxCust.FirstPersRec = 0
      Put THandle, x, TaxCust
      PCnt = PCnt + 1
    End If
  Next x
  Call TaxMsg(800, "A total of " + CStr(PCnt) + " personal properties were removed.")

  PCnt = 0
  For x = 1000 To NumOfTaxCustRecs
    Get THandle, x, TaxCust
    NextRec = TaxCust.FirstPersRec
    If NextRec > 0 Then
      For y = 1 To 999
        If y = x Then GoTo Skip
        Get THandle, y, TaxCust
        If TaxCust.FirstPersRec = NextRec Then
          TaxCust.FirstPersRec = 0
          Put THandle, y, TaxCust
          PCnt = PCnt + 1
        End If
Skip:
      Next y
    End If
  Next x
  
  For x = 1 To NumOfTaxCustRecs
    Get THandle, x, TaxCust
    NextRec = TaxCust.FirstPersRec
    If NextRec > 0 Then
      For y = 1 To NumOfTaxCustRecs
        If y = x Then GoTo Skip1
        Get THandle, y, TaxCust
        If TaxCust.FirstPersRec = NextRec Then
          Stop
          PCnt = PCnt + 1
        End If
Skip1:
      Next y
    End If
  Next x
    
  Close
  Call TaxMsg(800, "A total of " + CStr(PCnt) + " personal properties were removed.")
  
End Sub

Private Sub FixWhiteLakeOverPay()
  Dim OverPayAmt As Double
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim BadBillTaxTrans As TaxTransactionType
  Dim PayTranRecAdd As TaxTransactionType
  Dim PayTranRecRemove As TaxTransactionType
  Dim GoodTaxCust As TaxCustType
  Dim BadTaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, y As Long
  Dim GoodTaxCustRec As Long
  Dim BadTaxCustRec As Long
  Dim GoodBillRec&
  Dim BillNum$
  Dim BadOPRec As Long
  Dim BadBillRec As Long
  Dim TotalPaid#
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  
  'fix 2869 to 3763
  GoodTaxCustRec = 3763
  BadTaxCustRec = 2869
  OverPayAmt = 10.1
  BadBillRec = 77378
  GoodBillRec = 77816
  BadOPRec = 77379
  GoSub FixIt
  
  'fix 4637 to 4689
  GoodTaxCustRec = 4689
  BadTaxCustRec = 4637
  OverPayAmt = 1.06
  BadBillRec = 76015
  GoodBillRec = 78109
  BadOPRec = 76016
  GoSub FixIt
  
  'fix 756 to 4650
  GoodTaxCustRec = 4650
  BadTaxCustRec = 756
  OverPayAmt = 2.32
  BadBillRec = 76305
  GoodBillRec = 76542
  BadOPRec = 76306
  GoSub FixIt
  
  'fix 4323 to 2380
  GoodTaxCustRec = 2380
  BadTaxCustRec = 4323
  OverPayAmt = 0.31
  BadBillRec = 76559
  GoodBillRec = 76843
  BadOPRec = 76560
  GoSub FixIt
  
  'fix 1472 to 3470
  GoodTaxCustRec = 3470
  BadTaxCustRec = 1472
  OverPayAmt = 1.34
  BadBillRec = 77602
  GoodBillRec = 78077
  BadOPRec = 77603
  GoSub FixIt
  
  'fix 1472 to 3470
  GoodTaxCustRec = 2401
  BadTaxCustRec = 2792
  OverPayAmt = 15.09
  BadBillRec = 76630
  GoodBillRec = 76925
  BadOPRec = 76631
  GoSub FixIt
  
  'fix 3370 to 3312
  GoodTaxCustRec = 3312
  BadTaxCustRec = 3370
  OverPayAmt = 2.52
  BadBillRec = 75782
  GoodBillRec = 75925
  BadOPRec = 75783
  GoSub FixIt
  
  'fix 3641 to 1347
  GoodTaxCustRec = 1347
  BadTaxCustRec = 3641
  OverPayAmt = 10.8
  BadBillRec = 77010
  GoodBillRec = 77378
  BadOPRec = 77011
  GoSub FixIt
  
  'fix 3100 to 2332
  GoodTaxCustRec = 2332
  BadTaxCustRec = 3100
  OverPayAmt = 835.43
  BadBillRec = 76233
  GoodBillRec = 76461
  BadOPRec = 76234
  GoSub FixIt
  
  'fix 4144 to 4708
  GoodTaxCustRec = 4708
  BadTaxCustRec = 4144
  OverPayAmt = 211.5
  BadBillRec = 75908
  GoodBillRec = 76073
  BadOPRec = 75909
  GoSub FixIt
  
  'fix 567 to 4641
  GoodTaxCustRec = 4641
  BadTaxCustRec = 567
  OverPayAmt = 0.78
  BadBillRec = 75971
  GoodBillRec = 76146
  BadOPRec = 75972
  GoSub FixIt
  
  'fix 4543 to 4566
  GoodTaxCustRec = 4566
  BadTaxCustRec = 4543
  OverPayAmt = 54.98
  BadBillRec = 77448
  GoodBillRec = 77900
  BadOPRec = 77449
  GoSub FixIt
  
  Close
  Call TaxMsg(900, "Finished.")
  Exit Sub
  
FixIt:
  For x = 1 To 1
    Get TCHandle, GoodTaxCustRec, GoodTaxCust
    Get TCHandle, BadTaxCustRec, BadTaxCust
    Get TTHandle, GoodBillRec, TaxTrans
    BillNum = ParseBillNum(TaxTrans.Description)
    Get TTHandle, BadOPRec, PayTranRecRemove
    GoSub OverPayAdd
    GoSub OverPayRemove
  Next x
  
  Return
  
OverPayAdd:
  TotalPaid# = OverPayAmt
  PayTranRecAdd.TransDate = Date2Num%(Date$)
  PayTranRecAdd.TranType = 9
  PayTranRecAdd.Revenue.Principle1Pd = PayTranRecRemove.Revenue.Principle1Pd
  PayTranRecAdd.Revenue.InterestPd = 0
  PayTranRecAdd.Revenue.CollectionPd = 0
  PayTranRecAdd.Revenue.LateListPd = PayTranRecRemove.Revenue.LateListPd
  PayTranRecAdd.Revenue.RevOpt1Pd = PayTranRecRemove.Revenue.RevOpt1Pd
  PayTranRecAdd.Revenue.RevOpt2Pd = PayTranRecRemove.Revenue.RevOpt2Pd
  PayTranRecAdd.Revenue.RevOpt3Pd = PayTranRecRemove.Revenue.RevOpt3Pd
  PayTranRecAdd.CustPin = GoodTaxCust.PIN
  PayTranRecAdd.DiscXDate = TaxTrans.DiscXDate
  PayTranRecAdd.RealPin = QPTrim$(TaxTrans.RealPin)
  PayTranRecAdd.PersPin = QPTrim$(TaxTrans.PersPin)
  PayTranRecAdd.Posted2GL = "N"
  PayTranRecAdd.TaxYear = TaxTrans.TaxYear
  PayTranRecAdd.DiscAmt = 0
  PayTranRecAdd.OperNum = PayTranRecRemove.OperNum
  PayTranRecAdd.Amount = 0
  PayTranRecAdd.FromPrePay = TotalPaid#
  PayTranRecAdd.Description = "Credit Applied to Bill# " + BillNum
  PayTranRecAdd.CustomerRec = TaxTrans.CustomerRec
  PayTranRecAdd.LastTrans = GoodTaxCust.LastTrans
  PayTranRecAdd.BelongTo = GoodBillRec&
  PayTranRecAdd.Revenue.PrePaidAmt = 0
  PayTranRecAdd.Revenue.PrePaidUsed = OverPayAmt
  PayTranRecAdd.Revenue.PrePaidBal = OldRound(GetOverPayBalance(GoodTaxCust.Acct) - OverPayAmt)
  PayTranRecAdd.InternalPin = TaxTrans.InternalPin
  PayTranRecAdd.CntyPara = ""
  PayTranRecAdd.CyclPara = ""
  PayTranRecAdd.TShpPara = ""
  Get TTHandle, GoodBillRec&, TaxTrans
    TaxTrans.Revenue.Principle1Pd = OldRound#(TaxTrans.Revenue.Principle1Pd + PayTranRecAdd.Revenue.Principle1Pd) '4/22/05
    TaxTrans.Revenue.InterestPd = OldRound#(TaxTrans.Revenue.InterestPd + PayTranRecAdd.Revenue.Interest)
    TaxTrans.Revenue.CollectionPd = OldRound#(TaxTrans.Revenue.CollectionPd + PayTranRecAdd.Revenue.Collection)
    TaxTrans.Revenue.LateListPd = OldRound#(TaxTrans.Revenue.LateListPd + PayTranRecAdd.Revenue.LateListPd)
    TaxTrans.Revenue.RevOpt1Pd = OldRound#(TaxTrans.Revenue.RevOpt1Pd + PayTranRecAdd.Revenue.RevOpt1Pd)
    TaxTrans.Revenue.RevOpt2Pd = OldRound#(TaxTrans.Revenue.RevOpt2Pd + PayTranRecAdd.Revenue.RevOpt2Pd)
    TaxTrans.Revenue.RevOpt3Pd = OldRound#(TaxTrans.Revenue.RevOpt3Pd + PayTranRecAdd.Revenue.RevOpt3Pd)
      
  Put TTHandle, GoodBillRec&, TaxTrans
  
  GoodBillRec& = GoodBillRec& + 1

  Put TTHandle, GoodBillRec&, PayTranRecAdd
  
  GoodTaxCust.LastTrans = GoodBillRec&
  Put TCHandle, GoodTaxCust.Acct, GoodTaxCust
  
  Return
  
OverPayRemove:
  TotalPaid# = 0
  PayTranRecRemove.Revenue.Principle1Pd = 0
  PayTranRecRemove.Revenue.InterestPd = 0
  PayTranRecRemove.Revenue.CollectionPd = 0
  PayTranRecRemove.Revenue.LateListPd = 0
  PayTranRecRemove.Revenue.RevOpt1Pd = 0
  PayTranRecRemove.Revenue.RevOpt2Pd = 0
  PayTranRecRemove.Revenue.RevOpt3Pd = 0
  PayTranRecRemove.CustPin = 0
  PayTranRecRemove.DiscXDate = 0
  PayTranRecRemove.RealPin = ""
  PayTranRecRemove.PersPin = ""
  PayTranRecRemove.Posted2GL = "N"
  PayTranRecRemove.TaxYear = TaxTrans.TaxYear
  PayTranRecRemove.DiscAmt = 0
  PayTranRecRemove.OperNum = OperNum
  PayTranRecRemove.Amount = 0
  PayTranRecRemove.FromPrePay = 0
  PayTranRecRemove.Description = ""
  PayTranRecRemove.CustomerRec = 0
  PayTranRecRemove.LastTrans = 0
  PayTranRecRemove.BelongTo = 0
  PayTranRecRemove.Revenue.PrePaidAmt = 0
  PayTranRecRemove.Revenue.PrePaidUsed = 0
  PayTranRecRemove.Revenue.PrePaidBal = 0
  PayTranRecRemove.InternalPin = 0
  PayTranRecRemove.CntyPara = ""
  PayTranRecRemove.CyclPara = ""
  PayTranRecRemove.TShpPara = ""
  Get TTHandle, BadBillRec, BadBillTaxTrans
    BadBillTaxTrans.Revenue.Principle1Pd = OldRound#(BadBillTaxTrans.Revenue.Principle1Pd - PayTranRecAdd.Revenue.Principle1Pd) '4/22/05
    BadBillTaxTrans.Revenue.InterestPd = OldRound#(BadBillTaxTrans.Revenue.InterestPd - PayTranRecAdd.Revenue.Interest)
    BadBillTaxTrans.Revenue.CollectionPd = OldRound#(BadBillTaxTrans.Revenue.CollectionPd - PayTranRecAdd.Revenue.Collection)
    BadBillTaxTrans.Revenue.LateListPd = OldRound#(BadBillTaxTrans.Revenue.LateListPd - PayTranRecAdd.Revenue.LateListPd)
    BadBillTaxTrans.Revenue.RevOpt1Pd = OldRound#(BadBillTaxTrans.Revenue.RevOpt1Pd - PayTranRecAdd.Revenue.RevOpt1Pd)
    BadBillTaxTrans.Revenue.RevOpt2Pd = OldRound#(BadBillTaxTrans.Revenue.RevOpt2Pd - PayTranRecAdd.Revenue.RevOpt2Pd)
    BadBillTaxTrans.Revenue.RevOpt3Pd = OldRound#(BadBillTaxTrans.Revenue.RevOpt3Pd - PayTranRecAdd.Revenue.RevOpt3Pd)
  Put TTHandle, BadBillRec, BadBillTaxTrans
  
'  GoodBillRec& = GoodBillRec& + 1

  Put TTHandle, BadOPRec, PayTranRecRemove
  
  BadTaxCust.LastTrans = BadBillRec
  Put TCHandle, BadTaxCust.Acct, BadTaxCust
  
  Return
  
End Sub

Private Sub cmdCnvrtPstdBills_Click()
  Dim ThisFile$
  Dim THandle As Integer
  Dim TaxBill As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long, y As Integer
  Dim THandleOld As Integer
  Dim TaxBillOld As TaxBillTypeOld
  Dim TBHandleOld As Integer
  Dim NumOfTBRecsOld As Long
  Dim DirContents() As String
  Dim DirCnt As Integer
  Dim MyPath$
  Dim MyName$
  Dim BigCnt As Long
  
  DirCnt = 0
  MyPath = StartPath + "\TAXBILLBU\"
  MyName$ = Dir(MyPath, vbDirectory)
  Do While MyName <> ""
    MyName = Dir
    If Len(MyName) > 4 Then
      DirCnt = DirCnt + 1
      ReDim Preserve DirContents(DirCnt) As String
      DirContents(DirCnt) = MyPath + MyName
    End If
  Loop
  If DirCnt = 0 Then
    Call TaxMsg(900, "There are no files to convert.")
    Close
    Exit Sub
  End If
  
  For y = 1 To DirCnt
    ThisFile = DirContents(y)
    OpenOldPostedReprintFile THandleOld, NumOfTBRecsOld, ThisFile
    If NumOfTBRecsOld > BigCnt Then BigCnt = NumOfTBRecsOld
    Close THandleOld
  Next y
  
  For y = 1 To DirCnt
    ThisFile = DirContents(y)
    OpenPostedReprintFile THandle, NumOfTBRecs, ThisFile
    If NumOfTBRecs > BigCnt Then BigCnt = NumOfTBRecs
    Close THandle
  Next y
  
  ReDim CustRec(1 To DirCnt, 1 To BigCnt) As Long
  ReDim CustName(1 To DirCnt, 1 To BigCnt) As String
  ReDim CustAdd1(1 To DirCnt, 1 To BigCnt) As String
  ReDim CustAdd2(1 To DirCnt, 1 To BigCnt) As String
  ReDim CustAdd3(1 To DirCnt, 1 To BigCnt) As String
  ReDim CustZip(1 To DirCnt, 1 To BigCnt) As String
  ReDim RDesc1(1 To DirCnt, 1 To BigCnt) As String
  ReDim RDesc2(1 To DirCnt, 1 To BigCnt) As String
  ReDim RealPin(1 To DirCnt, 1 To BigCnt) As String
  ReDim PersPin(1 To DirCnt, 1 To BigCnt) As String
  ReDim RealValue(1 To DirCnt, 1 To BigCnt) As Double
  ReDim PersValue(1 To DirCnt, 1 To BigCnt) As Double
  ReDim ExptValue(1 To DirCnt, 1 To BigCnt) As Double
  ReDim RealTaxDue(1 To DirCnt, 1 To BigCnt) As Double
  ReDim PersTaxDue(1 To DirCnt, 1 To BigCnt) As Double
  ReDim LateTaxDue(1 To DirCnt, 1 To BigCnt) As Double
  ReDim TotalBillDue(1 To DirCnt, 1 To BigCnt) As Double
  ReDim BillNumber(1 To DirCnt, 1 To BigCnt) As Long
  ReDim TaxYear(1 To DirCnt, 1 To BigCnt) As Integer
  ReDim BillPrinted(1 To DirCnt, 1 To BigCnt) As Integer
  ReDim RealPropRecord(1 To DirCnt, 1 To BigCnt) As Long
  ReDim PersPropRecord(1 To DirCnt, 1 To BigCnt) As Long
  ReDim PriorYrBalance(1 To DirCnt, 1 To BigCnt) As Double
  ReDim RealTaxRate(1 To DirCnt, 1 To BigCnt) As Double
  ReDim PersTaxRate(1 To DirCnt, 1 To BigCnt) As Double
  ReDim CustPin(1 To DirCnt, 1 To BigCnt) As Long
  ReDim TownShip(1 To DirCnt, 1 To BigCnt) As String
  ReDim MORTCODE(1 To DirCnt, 1 To BigCnt) As String
  ReDim LotOrAcre(1 To DirCnt, 1 To BigCnt) As String
  ReDim LASize(1 To DirCnt, 1 To BigCnt) As String
  ReDim MortRec(1 To DirCnt, 1 To BigCnt) As Integer
  ReDim CarShore(1 To DirCnt, 1 To BigCnt) As Double
  ReDim RDesc3(1 To DirCnt, 1 To BigCnt) As String
  ReDim InternalPin(1 To DirCnt, 1 To BigCnt) As Long
  ReDim OptRevTax1(1 To DirCnt, 1 To BigCnt) As Double
  ReDim OptRevTax2(1 To DirCnt, 1 To BigCnt) As Double
  ReDim OptRevTax3(1 To DirCnt, 1 To BigCnt) As Double
  ReDim OverPayAmt(1 To DirCnt, 1 To BigCnt) As Double
  ReDim Padding(1 To DirCnt, 1 To BigCnt) As String
  
  For y = 1 To DirCnt
    ThisFile = DirContents(y)
    OpenOldPostedReprintFile THandleOld, NumOfTBRecsOld, ThisFile
    For x = 1 To NumOfTBRecsOld
      Get THandleOld, x, TaxBillOld
      CustRec(y, x) = TaxBillOld.CustRec
      CustName(y, x) = TaxBillOld.CustName
      CustAdd1(y, x) = TaxBillOld.CustAdd1
      CustAdd2(y, x) = TaxBillOld.CustAdd2
      CustAdd3(y, x) = TaxBillOld.CustAdd3
      CustZip(y, x) = TaxBillOld.CustZip
      RDesc1(y, x) = TaxBillOld.RDesc1
      RDesc2(y, x) = TaxBillOld.RDesc2
      RealPin(y, x) = TaxBillOld.RealPin
      PersPin(y, x) = TaxBillOld.PersPin
      RealValue(y, x) = TaxBillOld.RealValue
      PersValue(y, x) = TaxBillOld.PersValue
      ExptValue(y, x) = TaxBillOld.ExptValue
      RealTaxDue(y, x) = TaxBillOld.RealTaxDue
      PersTaxDue(y, x) = TaxBillOld.PersTaxDue
      LateTaxDue(y, x) = TaxBillOld.LateTaxDue
      TotalBillDue(y, x) = TaxBillOld.TotalBillDue
      BillNumber(y, x) = TaxBillOld.BillNumber
      TaxYear(y, x) = TaxBillOld.TaxYear
      BillPrinted(y, x) = TaxBillOld.BillPrinted
      RealPropRecord(y, x) = TaxBillOld.RealPropRecord
      PersPropRecord(y, x) = TaxBillOld.PersPropRecord
      PriorYrBalance(y, x) = TaxBillOld.PriorYrBalance
      RealTaxRate(y, x) = TaxBillOld.RealTaxRate
      PersTaxRate(y, x) = TaxBillOld.PersTaxRate
      CustPin(y, x) = TaxBillOld.CustPin
      TownShip(y, x) = TaxBillOld.TownShip
      MORTCODE(y, x) = TaxBillOld.MORTCODE
      LotOrAcre(y, x) = TaxBillOld.LotOrAcre
      LASize(y, x) = TaxBillOld.LASize
      MortRec(y, x) = TaxBillOld.MortRec
      CarShore(y, x) = TaxBillOld.CarShore
      RDesc3(y, x) = TaxBillOld.RDesc3
      InternalPin(y, x) = TaxBillOld.InternalPin
      OptRevTax1(y, x) = TaxBillOld.OptRevTax1
      OptRevTax2(y, x) = TaxBillOld.OptRevTax2
      OptRevTax3(y, x) = TaxBillOld.OptRevTax3
      OverPayAmt(y, x) = TaxBillOld.OverPayAmt
      Padding(y, x) = "Converted"
    Next x
    Close THandleOld
  Next y
    
  Close
  
  For y = 1 To DirCnt
    ThisFile = DirContents(y)
    OpenPostedReprintFile THandle, NumOfTBRecs, ThisFile
    For x = 1 To NumOfTBRecs
      Get THandle, x, TaxBill
      TaxBill.CustRec = CustRec(y, x)
      TaxBill.CustName = CustName(y, x)
      TaxBill.CustAdd1 = CustAdd1(y, x)
      TaxBill.CustAdd2 = CustAdd2(y, x)
      TaxBill.CustAdd3 = CustAdd3(y, x)
      TaxBill.CustZip = CustZip(y, x)
      TaxBill.RDesc1 = RDesc1(y, x)
      TaxBill.RDesc2 = RDesc2(y, x)
      TaxBill.RealPin = RealPin(y, x)
      TaxBill.PersPin = PersPin(y, x)
      TaxBill.RealValue = RealValue(y, x)
      TaxBill.PersValue = PersValue(y, x)
      TaxBill.ExptValue = ExptValue(y, x)
      TaxBill.RealTaxDue = RealTaxDue(y, x)
      TaxBill.PersTaxDue = PersTaxDue(y, x)
      TaxBill.LateTaxDue = LateTaxDue(y, x)
      TaxBill.TotalBillDue = TotalBillDue(y, x)
      TaxBill.BillNumber = BillNumber(y, x)
      TaxBill.TaxYear = TaxYear(y, x)
      TaxBill.BillPrinted = BillPrinted(y, x)
      TaxBill.RealPropRecord = RealPropRecord(y, x)
      TaxBill.PersPropRecord = PersPropRecord(y, x)
      TaxBill.PriorYrBalance = PriorYrBalance(y, x)
      TaxBill.RealTaxRate = RealTaxRate(y, x)
      TaxBill.PersTaxRate = PersTaxRate(y, x)
      TaxBill.CustPin = CustPin(y, x)
      TaxBill.TownShip = TownShip(y, x)
      TaxBill.MORTCODE = MORTCODE(y, x)
      TaxBill.LotOrAcre = LotOrAcre(y, x)
      TaxBill.LASize = LASize(y, x)
      TaxBill.MortRec = MortRec(y, x)
      TaxBill.CarShore = CarShore(y, x)
      TaxBill.RDesc3 = RDesc3(y, x)
      TaxBill.InternalPin = InternalPin(y, x)
      TaxBill.OptRevTax1 = OptRevTax1(y, x)
      TaxBill.OptRevTax2 = OptRevTax2(y, x)
      TaxBill.OptRevTax3 = OptRevTax3(y, x)
      TaxBill.OverPayAmt = OverPayAmt(y, x)
      TaxBill.Padding = Padding(y, x)
      TaxBill.SetDscvry2No = "N"
      Put THandle, x, TaxBill
    Next x
    Close THandle
  Next y
  Close
  
  Call Savemsg(900, "Conversion completed successfully.")
    
End Sub

Private Sub cmdFixBSLAdvChrgAndReal_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  Dim RealRec As PropertyRecType
  Dim NumOfRRecs As Long
  Dim RHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTaxCusts As Long
  Dim RCnt As Integer
  Dim RealIdx As Long
  Dim ThisCnt As Integer
  Dim ChangeCnt As Integer
  Dim CustRec As Long
  Dim FileName$
  Dim ThisFile As Integer
  
  FileName = "bsladvchrgaddedtoreal.txt"
  ThisFile = FreeFile
  Open FileName For Output As ThisFile
  
  OpenRealPropFile RHandle, NumOfRRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxCustFile TCHandle, NumOfTaxCusts
  RCnt = 0
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    ThisCnt = 0
    If TaxTrans.TranType = 6 Then
      GoSub GetCustRec
      If CustRec = 0 Then GoTo Skip
      Get TCHandle, CustRec, TaxCust
      RealIdx = TaxCust.FirstPropRec
      Do While RealIdx > 0
        Get RHandle, RealIdx, RealRec
        RCnt = RCnt + 1
        ThisCnt = ThisCnt + 1
        If ThisCnt = 1 Then
          Print #ThisFile, "*" + QPTrim$(TaxCust.CustName) + "~" + CStr(TaxCust.Acct) + "~" + QPTrim$(RealRec.RealPin) + "~" + FormatCurrency(TaxTrans.Amount, 2)
        Else
          Print #ThisFile, QPTrim$(TaxCust.CustName) + "~" + CStr(TaxCust.Acct) + "~" + QPTrim$(RealRec.RealPin) + "~" + FormatCurrency(TaxTrans.Amount, 2)
        End If
        RealIdx = RealRec.NextRec
      Loop
      If ThisCnt = 1 Then
        TaxTrans.RealPin = QPTrim$(RealRec.RealPin)
        ChangeCnt = ChangeCnt + 1
        Put TTHandle, x, TaxTrans
      End If
    End If
Skip:
  Next x
  Print #ThisFile, "A total of " + CStr(ChangeCnt) + " changes were made. A total of " + CStr(RCnt) + " errors were found. "
  Close
  MsgBox ("Process completed with " + CStr(RCnt) + " errors found and " + CStr(ChangeCnt) + " errors fixed.")

Exit Sub

GetCustRec:
  Select Case TaxTrans.BelongTo
    Case 20367
      CustRec = 10522
    Case 19982
      CustRec = 10526
    Case 22838
      CustRec = 10527
    Case 19655
      CustRec = 10578
    Case 19686
      CustRec = 10989
    Case 18227
      CustRec = 11013
    Case 17614
      CustRec = 11014
    Case 18892
      CustRec = 11041
    Case 20965
      CustRec = 11136
    Case Else
      CustRec = 0
  End Select
Return

End Sub

Private Sub UpdateAddressFields()
  Dim x As Long, y As Long
  Dim TextLine$
  Dim ThisFile$
  Dim AHandle As Integer
  Dim WordCnt As Integer
  Dim TextLen As Integer
  Dim Thisch As String
  Dim ThisWord$
  Dim CntyNum As String
  Dim Add1 As String
  Dim Add2 As String
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim dlm As String
  Dim Cnt As Integer
  Dim track As Integer
  
  If MsgBox("Did you make the last line(s) = 'End~~'?", vbYesNo) = vbNo Then
    Exit Sub
  End If
  dlm = "~"
  track = 0
  WordCnt = 0
  ReDim Words(1 To 1) As String
  frmTaxShowPctComp.Label1 = "Addresses Update"
  frmTaxShowPctComp.Show , Me

  If Exist("addresses.csv") Then
    AHandle = FreeFile
    ThisFile = "addresses.csv"
    Open ThisFile For Input As #AHandle
    Do While ThisWord <> "End"
      Line Input #AHandle, TextLine
      If InStr(TextLine, "End~~") Then Exit Do
      track = track + 1
    Loop
    Close
    OpenTaxCustFile TCHandle, NumOfTCRecs
    AHandle = FreeFile
    Open ThisFile For Input As #AHandle
    Do While ThisWord <> "End"
      Line Input #AHandle, TextLine
      TextLen = Len(TextLine)
      TextLine = TextLine + dlm
      For x = 1 To TextLen + 1
        Thisch = Mid(TextLine, x, 1)
        If Thisch = dlm Then
          WordCnt = WordCnt + 1
          ReDim Preserve Words(1 To WordCnt) As String
          If WordCnt = 1 Then
            CntyNum = ThisWord
            ThisWord = ""
            GoTo NewWord
          ElseIf WordCnt = 2 Then
            Add1 = ThisWord
            ThisWord = ""
            GoTo NewWord
          ElseIf WordCnt = 3 Then
            Add2 = ThisWord
            GoSub SaveAdd
            Add1 = ""
            Add2 = ""
            CntyNum = ""
            ThisWord = ""
            WordCnt = 0
            GoTo NewLoop
          End If
        End If
        ThisWord = ThisWord + Thisch
        If ThisWord = "End" Then Exit Do

NewWord:
      Next x
NewLoop:
    frmTaxShowPctComp.ShowPctComp Cnt, track
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
    End If
    Loop
  Else
    MsgBox ("The file 'addresses.csv' cannot be found.")
    Exit Sub
  End If
  Unload frmTaxShowPctComp
  MsgBox (CStr(Cnt) & " addresses were updated successfully.")
  Close
  DoEvents
  
  Exit Sub
  
SaveAdd:
   Cnt = Cnt + 1
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If QPTrim$(TaxCust.CountyAcctString) = "" Then
      TaxCust.CountyAcctString = CStr(TaxCust.CountyAcct)
    End If
    If QPTrim$(TaxCust.CountyAcctString) = CntyNum Then
      TaxCust.Addr1 = Add1
      TaxCust.Addr2 = Add2
      Put TCHandle, x, TaxCust
      Exit For
    End If
Skip:
  Next x
  
  Return

End Sub

Private Sub PatchSprucePine()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  Dim NextTrans As Long
  Dim Belong As Long
  Dim Amount As Double
  Dim CustRec As Long
  Dim BillType As String
  
  ReDim arr(1 To 5) As Long
  arr(1) = 9710
  arr(2) = 8070
  arr(3) = 8704
  arr(4) = 9379
  arr(5) = 9622
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For x = 1 To 5
    Get TTHandle, arr(x), TaxTrans
    Belong = TaxTrans.BelongTo
    NextTrans = TaxTrans.LastTrans
    Amount = TaxTrans.Amount
    CustRec = TaxTrans.CustomerRec
    GoSub Bill
    Get TTHandle, arr(x), TaxTrans
    TaxTrans.LastTrans = Belong
    Put TTHandle, arr(x), TaxTrans
  Next x
  
  Close
  MsgBox ("Done.")
  Exit Sub
  
Bill:
  TaxTrans.TransDate = Date2Num%("09/09/2005")
  TaxTrans.TaxYear = 2005
  TaxTrans.TranType = 1
  If x = 4 Then
    TaxTrans.BillType = "C"
  Else
    TaxTrans.BillType = "P"
  End If
  TaxTrans.Amount = Amount

  TaxTrans.Revenue.Principle1 = Amount
  TaxTrans.Revenue.Principle2 = 0
  TaxTrans.Revenue.Principle3 = 0
  TaxTrans.Revenue.Principle4 = 0
  TaxTrans.Revenue.Principle5 = 0
  TaxTrans.Revenue.Interest = 0
  TaxTrans.Revenue.Penalty = 0
  TaxTrans.Revenue.Collection = 0
  TaxTrans.Revenue.Future1 = 0
  TaxTrans.Revenue.Future2 = 0
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.Principle2Pd = 0
  TaxTrans.Revenue.Principle3Pd = 0
  TaxTrans.Revenue.Principle4Pd = 0
  TaxTrans.Revenue.Principle5Pd = 0
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.PenaltyPd = 0
  TaxTrans.Revenue.CollectionPd = 0
  TaxTrans.Revenue.Future1Pd = 0
  TaxTrans.Revenue.Future2Pd = 0
  TaxTrans.Revenue.RevOpt1 = 0
  TaxTrans.Revenue.RevOpt1Pd = 0
  TaxTrans.Revenue.RevOpt2 = 0
  TaxTrans.Revenue.RevOpt2Pd = 0
  TaxTrans.Revenue.RevOpt3 = 0
  TaxTrans.Revenue.RevOpt3Pd = 0
  TaxTrans.Revenue.LateList = 0
  TaxTrans.Revenue.LateListPd = 0
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.InternalPin = CustRec
  TaxTrans.Revenue.pad = ""

  TaxTrans.Description = "Tax Bill #" + CStr(Belong)
  TaxTrans.Posted2GL = "Y"
  TaxTrans.CustomerRec = CustRec
  TaxTrans.LastTrans = 0
  TaxTrans.BelongTo = 0
  TaxTrans.Padding = ""
  TaxTrans.PersPin = ""
  TaxTrans.RealPin = ""
  TaxTrans.CustPin = CustRec
  TaxTrans.DiscXDate = Date2Num("09/09/2005")
  TaxTrans.DiscAmt = 0
  TaxTrans.OperNum = 0
  TaxTrans.CntyPara = ""
  TaxTrans.CyclPara = ""
  TaxTrans.TShpPara = ""
  TaxTrans.LastTrans = NextTrans
  Put TTHandle, Belong, TaxTrans

Return

End Sub

Private Sub RemovePayments()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  Dim PayDate As Integer
  Dim BelongTo As Integer
  Dim Amount As Double
  Dim DiscAmt As Double
  Dim FromPrePay As Double
  Dim CollectionPd As Double
  Dim Future1Pd As Double
  Dim Future2Pd As Double
  Dim InterestPd As Double
  Dim LateListPd As Double
  Dim PenaltyPd As Double
  Dim PrePaidAmt As Double
  Dim PrePaidBal As Double
  Dim PrePaidUsed As Double
  Dim Principle1Pd As Double
  Dim Principle2Pd As Double
  Dim Principle3Pd As Double
  Dim Principle4Pd As Double
  Dim Principle5Pd As Double
  Dim RevOpt1Pd As Double
  Dim RevOpt2Pd As Double
  Dim RevOpt3Pd As Double
  Dim Cnt As Integer
  Dim BillDate As Integer
  Dim XDate As Integer
  
  PayDate = Date2Num("09/02/2008")
  BillDate = Date2Num("08/05/2008")
  XDate = Date2Num("09/05/2008")
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    If TaxTrans.TransDate = BillDate And TaxTrans.TranType = 1 Then 'needed this for customer
    'but normally would not be included
      TaxTrans.DiscXDate = XDate
      Put TTHandle, x, TaxTrans
    End If
    BelongTo = TaxTrans.BelongTo
    If TaxTrans.TransDate = PayDate And TaxTrans.TranType = 2 Then
      Cnt = Cnt + 1
      CollectionPd = TaxTrans.Revenue.CollectionPd
      Future1Pd = TaxTrans.Revenue.Future1Pd
      Future2Pd = TaxTrans.Revenue.Future2Pd
      InterestPd = TaxTrans.Revenue.InterestPd
      LateListPd = TaxTrans.Revenue.LateListPd
      PenaltyPd = TaxTrans.Revenue.PenaltyPd
      PrePaidAmt = TaxTrans.Revenue.PrePaidAmt
      PrePaidBal = TaxTrans.Revenue.PrePaidBal
      PrePaidUsed = TaxTrans.Revenue.PrePaidUsed
      Principle1Pd = TaxTrans.Revenue.Principle1Pd
      Principle2Pd = TaxTrans.Revenue.Principle2Pd
      Principle3Pd = TaxTrans.Revenue.Principle3Pd
      Principle4Pd = TaxTrans.Revenue.Principle4Pd
      Principle5Pd = TaxTrans.Revenue.Principle5Pd
      RevOpt1Pd = TaxTrans.Revenue.RevOpt1Pd
      RevOpt2Pd = TaxTrans.Revenue.RevOpt2Pd
      RevOpt3Pd = TaxTrans.Revenue.RevOpt3Pd
      Get TTHandle, BelongTo, TaxTrans
      TaxTrans.Revenue.CollectionPd = OldRound(TaxTrans.Revenue.CollectionPd - CollectionPd)
      TaxTrans.Revenue.Future1Pd = OldRound(TaxTrans.Revenue.Future1Pd - Future1Pd)
      TaxTrans.Revenue.Future2Pd = OldRound(TaxTrans.Revenue.Future2Pd - Future2Pd)
      TaxTrans.Revenue.InterestPd = OldRound(TaxTrans.Revenue.InterestPd - InterestPd)
      TaxTrans.Revenue.LateListPd = OldRound(TaxTrans.Revenue.LateListPd - LateListPd)
      TaxTrans.Revenue.PenaltyPd = OldRound(TaxTrans.Revenue.PenaltyPd - PenaltyPd)
      TaxTrans.Revenue.PrePaidAmt = OldRound(TaxTrans.Revenue.PrePaidAmt - PrePaidAmt)
      TaxTrans.Revenue.PrePaidBal = OldRound(TaxTrans.Revenue.PrePaidBal - PrePaidBal)
      TaxTrans.Revenue.PrePaidUsed = OldRound(TaxTrans.Revenue.PrePaidUsed - PrePaidUsed)
      TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd - Principle1Pd)
      TaxTrans.Revenue.Principle2Pd = OldRound(TaxTrans.Revenue.Principle2Pd - Principle2Pd)
      TaxTrans.Revenue.Principle3Pd = OldRound(TaxTrans.Revenue.Principle3Pd - Principle3Pd)
      TaxTrans.Revenue.Principle4Pd = OldRound(TaxTrans.Revenue.Principle4Pd - Principle4Pd)
      TaxTrans.Revenue.Principle5Pd = OldRound(TaxTrans.Revenue.Principle5Pd - Principle5Pd)
      TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd - RevOpt1Pd)
      TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd - RevOpt2Pd)
      TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd - RevOpt3Pd)
      Put TTHandle, BelongTo, TaxTrans
      Get TTHandle, x, TaxTrans
      TaxTrans.Amount = 0
      TaxTrans.DiscAmt = 0
      TaxTrans.FromPrePay = 0
      TaxTrans.Revenue.CollectionPd = 0
      TaxTrans.Revenue.Future1Pd = 0
      TaxTrans.Revenue.Future2Pd = 0
      TaxTrans.Revenue.InterestPd = 0
      TaxTrans.Revenue.LateListPd = 0
      TaxTrans.Revenue.PenaltyPd = 0
      TaxTrans.Revenue.PrePaidAmt = 0
      TaxTrans.Revenue.PrePaidBal = 0
      TaxTrans.Revenue.PrePaidUsed = 0
      TaxTrans.Revenue.Principle1Pd = 0
      TaxTrans.Revenue.Principle2Pd = 0
      TaxTrans.Revenue.Principle3Pd = 0
      TaxTrans.Revenue.Principle4Pd = 0
      TaxTrans.Revenue.Principle5Pd = 0
      TaxTrans.Revenue.RevOpt1Pd = 0
      TaxTrans.Revenue.RevOpt2Pd = 0
      TaxTrans.Revenue.RevOpt3Pd = 0
      Put TTHandle, x, TaxTrans
    End If
  Next x
  
  Close
  MsgBox ("A total of " & CStr(Cnt) & " payment transactions for 9/2/2008 were zeroed out successfully.")
End Sub

