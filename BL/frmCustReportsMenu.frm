VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmBLCustReportsMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Customer Reports Menu"
   ClientHeight    =   8730
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11565
   Icon            =   "frmCustReportsMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11565
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdCustList 
      Height          =   356
      Left            =   3960
      TabIndex        =   6
      Tag             =   "Press to bring up a report screen for tabulating customers in detail."
      Top             =   3682
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   628
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
      ButtonDesigner  =   "frmCustReportsMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCustTransHist 
      Height          =   345
      Left            =   3960
      TabIndex        =   4
      Tag             =   "Press to bring up a report screen for license transaction activity by customer."
      Top             =   2850
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   609
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
      ButtonDesigner  =   "frmCustReportsMenu.frx":0AB7
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   300
      Left            =   3120
      TabIndex        =   1
      Top             =   6000
      Width           =   690
      _Version        =   131072
      _ExtentX        =   1217
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   9405029
      ForeColor       =   8454143
      Text            =   ""
      Shape           =   0
      ShapeRoundWidth =   180
      ShapeRoundHeight=   180
      BorderWidth     =   -1
      BorderColor     =   -2147483630
      ThreeDWidth     =   -1
      ThreeDShadowColor=   -2147483632
      ThreeDHighlightColor=   16777215
      ThreeDText      =   0
      ThreeDTextHighlightColor=   16777215
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignPictureH   =   0
      AlignPictureV   =   0
      PictureStyle    =   0
      WordWrap        =   -1  'True
      ScaleMode       =   1
      ThreeDStyle     =   2
      Position        =   0
      PosBaseX        =   0
      PosBaseY        =   0
      PosOffsetX      =   -100
      PosOffsetY      =   300
      MaxWidth        =   5000
      CloudInset      =   100
      CloudMinWidth   =   600
      TailShape       =   2
      TailType        =   2
      TailBaseOffsetOutside=   300
      TailBaseOffsetInside=   100
      TailBaseAxisOutside=   0
      TailBaseAxisInside=   0
      TailBubbleCount =   3
      AlignTextH      =   1
      AlignTextV      =   1
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      AutoScan        =   0
      ScanAllDescendants=   -1  'True
      Interval        =   500
      IntervalNext    =   200
      AutoSize        =   -1  'True
      UseTagProp      =   -1  'True
      HideOnInactiveApp=   0   'False
      HideOnMouseDown =   2
      HideOnKeyDown   =   2
      HideOnFocus     =   0   'False
      ScanDisabledControls=   -1  'True
      ThreeDAppearance=   0
      FollowFocus     =   0   'False
      TemplateName    =   ""
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdBalList 
      Height          =   356
      Left            =   3960
      TabIndex        =   2
      Tag             =   "Press to bring up a report screen for customer balance listings."
      Top             =   2010
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   628
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
      ButtonDesigner  =   "frmCustReportsMenu.frx":0CA7
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdTransJrnl 
      Height          =   356
      Left            =   3960
      TabIndex        =   3
      Tag             =   "Press to bring up a report screen for license transaction types (charge, adjustment, etc.)"
      Top             =   2428
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   628
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
      ButtonDesigner  =   "frmCustReportsMenu.frx":0E93
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCatList 
      Height          =   345
      Left            =   3960
      TabIndex        =   5
      Tag             =   "Press to bring up a report screen for tabulating categories."
      Top             =   3270
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   609
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
      ButtonDesigner  =   "frmCustReportsMenu.frx":107B
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdLicList 
      Height          =   356
      Left            =   3960
      TabIndex        =   7
      Tag             =   "Press to bring up a report screen for listing customers by business license number including expiration dates."
      Top             =   4100
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   628
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
      ButtonDesigner  =   "frmCustReportsMenu.frx":1268
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExprLicList 
      Height          =   356
      Left            =   3960
      TabIndex        =   8
      Tag             =   "Press to bring up a report showing all businesses with expired licenses."
      Top             =   4518
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   628
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
      ButtonDesigner  =   "frmCustReportsMenu.frx":144B
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAppList 
      Height          =   356
      Left            =   3960
      TabIndex        =   9
      Tag             =   "Press to bring up a report screen for a brief list of all active businesses."
      Top             =   4936
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   628
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
      ButtonDesigner  =   "frmCustReportsMenu.frx":1636
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdInOut 
      Height          =   356
      Left            =   3960
      TabIndex        =   10
      Tag             =   "Click this button to bring up a report comparing businesses inside the city limits with those outside the city limits."
      Top             =   5354
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   628
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
      ButtonDesigner  =   "frmCustReportsMenu.frx":181D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCatRevRpt 
      Height          =   356
      Left            =   3960
      TabIndex        =   11
      Tag             =   "Press to access a revenue profile for each business license category."
      Top             =   5772
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   628
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
      ButtonDesigner  =   "frmCustReportsMenu.frx":1A0B
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCustByCat 
      Height          =   345
      Left            =   3960
      TabIndex        =   12
      Tag             =   "Press to bring up a report listing all categories and the customers assigned to them."
      Top             =   6195
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   609
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
      ButtonDesigner  =   "frmCustReportsMenu.frx":1BF6
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   360
      Left            =   3960
      TabIndex        =   14
      Top             =   7030
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
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
      ButtonDesigner  =   "frmCustReportsMenu.frx":1DDF
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   356
      Left            =   3960
      TabIndex        =   15
      Tag             =   "Click this button to return to the main Business License menu."
      Top             =   7455
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   628
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
      ButtonDesigner  =   "frmCustReportsMenu.frx":1FC4
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMailLbls 
      Height          =   360
      Left            =   3960
      TabIndex        =   13
      Top             =   6608
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
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
      ButtonDesigner  =   "frmCustReportsMenu.frx":21A9
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   155
      Index           =   3
      Left            =   8550
      Top             =   1995
      Width           =   985
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8666
      X2              =   8666
      Y1              =   2136
      Y2              =   8008
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   150
      Index           =   4
      Left            =   1970
      Top             =   2000
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BUSINESS LICENSE REPORTS"
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
      Left            =   2775
      TabIndex        =   0
      Top             =   1170
      Width           =   6012
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2086
      X2              =   2086
      Y1              =   2133
      Y2              =   8005
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2086
      X2              =   2795
      Y1              =   8025
      Y2              =   8025
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8655
      X2              =   9355
      Y1              =   8020
      Y2              =   8020
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1092
      Index           =   1
      Left            =   1455
      Top             =   820
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1455
      Top             =   690
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   1966
      Top             =   1890
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2086
      Top             =   2133
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8550
      Top             =   1890
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8655
      Top             =   2131
      Width           =   732
   End
End
Attribute VB_Name = "frmBLCustReportsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdCatRevRpt_Click()
  If Not Exist("arcatcodeidx.dat") Then 'no file there
    frmBLMessageBoxJr.Label1.Caption = "No category codes have been saved. Please save data for at least one category code. Loading aborted."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  If Not Exist("artrans.dat") Then 'no file there
    frmBLMessageBoxJr.Label1.Caption = "No transactions have taken place. Loading aborted."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  frmBLCatRevRpt.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdCustByCat_Click()
  If Not Exist("arcust.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Error: No customers have been saved. Form load aborted."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  If Not Exist("arcatcodeidx.dat") Then 'no file there
    frmBLMessageBoxJr.Label1.Caption = "No category codes have been saved. Please save data for at least one category code. Loading aborted."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  frmBLCustByCat.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "Turn Menu Hel&p Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "Turn Menu Hel&p On"
    btnHelp.AutoScan = fpAutoScanOff
  End If
End Sub

Private Sub cmdAppList_Click()
  If Not Exist("arcust.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Error: No customers have been saved. Form load aborted."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  frmBLAppListing.Show
  DoEvents
  Unload frmBLCustReportsMenu
End Sub

Private Sub cmdBalList_Click()
  If Not Exist("arcust.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Error: No customers have been saved. Form load aborted."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  frmBLCustBalListing.Show
  DoEvents
  Unload frmBLCustReportsMenu
End Sub

Private Sub cmdCatList_Click()
  Dim PrintType$
  Dim ReportFile$
  Dim FF$
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim CodeRec As ARNewCatCodeRecType
  Dim TrHandle As Integer
  Dim RptHandle As Integer
  Dim TRNumRecs As Integer
  Dim Count As Double
  Dim NumOfIdxRecs As Integer
  Dim IdxHandle As Integer
  Dim CodeIdx As CatCodeIdxType
  Dim x As Integer
  Dim cnt As Integer
  Dim Page As Integer
  Dim TotalCodes As Integer
  Dim dlm$
  Dim TownName$
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  
'  On Error GoTo ERRORSTUFF
  
  If Not Exist("arcatcodeidx.dat") Then 'no file there
    frmBLMessageBoxJr.Label1.Caption = "No category codes have been saved. Please save data for at least one category code. Loading aborted."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  dlm$ = "~"
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  TownName$ = QPTrim$(TownRec.TownName)
  frmBLReportOpt.Show vbModal 'opens small screen from which the
  'user selects the printing method
  PrintType$ = frmBLReportOpt.fptxtPrintType
  Unload frmBLReportOpt
  
  Select Case PrintType$
    Case "Graphical"
      GoSub PrintGraphics
    Case "Text"
      frmBLMessageBoxJr.Label1.Caption = "Pitch 10 is recommended for this report."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      GoSub PrintText
    Case "Exit"
  End Select
  
  Close
  cmdHelp.Text = "Turn Menu &Help On"
  btnHelp.AutoScan = fpAutoScanOff
  
  Exit Sub
  
PrintText:
  
  ReportFile$ = "ARCODLST.PRN"  'Report File Name
  FF$ = Chr$(12)
  MaxLines = 53
  LineCnt = 0
  
  OpenCatCodeIdxFile IdxHandle
  NumOfIdxRecs = LOF(IdxHandle) / Len(CodeIdx)
  OpenCatCodeFile TrHandle
  TRNumRecs = LOF(TrHandle) / Len(CodeRec)
  
  If TRNumRecs <> NumOfIdxRecs Then
    frmBLMessageBoxJr.Label1.Caption = "Error: The number of categories saved and the number of categories indexed are not the same. Re-index category codes or call Southern Software at 1-800-842-8190 for assistance."
    frmBLMessageBoxJr.Label1.Top = 600
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  ReDim IdxRecs(1 To NumOfIdxRecs) As Integer
  For x = 1 To NumOfIdxRecs
    Get IdxHandle, x, CodeIdx
    IdxRecs(x) = CodeIdx.CatCodeRec
  Next x
  Close IdxHandle
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  Print #RptHandle, Chr$(27); Chr$(58);         ' oki 320 12 cpi
  GoSub PrintRptHeader

  For cnt = 1 To TRNumRecs 'Count
    Get TrHandle, IdxRecs(cnt), CodeRec
    If Left$(CodeRec.CatCode, 1) <> " " Then
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintRptHeader
      End If
      Print #RptHandle, CodeRec.CatCode; Tab(8); Left$(CodeRec.CODEDESC, 30);
      Print #RptHandle, Tab(40); GetGLNum(CodeRec.REVGLNUM);
      Print #RptHandle, Tab(55); GetGLNum(CodeRec.ARGLACCT);
      Print #RptHandle, Tab(70); GetGLNum(CodeRec.CASHACCT)
      TotalCodes = TotalCodes + 1
      LineCnt = LineCnt + 1
    End If
  Next cnt
  GoSub PrintRptEnding
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  Close         'Close all open files now

  ViewPrint ReportFile, "Category Code Listing", True
  Kill ReportFile$
  Return

PrintRptHeader:
  Page = Page + 1
  Print #RptHandle, Tab(20); "Business License  : Category Code Listing "
  Print #RptHandle, "Report Date: "; Date$; Tab(65); "Page #"; Page
  Print #RptHandle, "Code "; Tab(8); "Description"; Tab(40); "Rev GL #"; Tab(55); "A/R GL #"; Tab(70); "Cash GL #"
  Print #RptHandle, String$(85, "=")
  LineCnt = 5
  Return

PrintRptEnding:
  Print #RptHandle, String$(80, "-")
  Print #RptHandle, "Number of Codes .. "; Using("###,##0", TotalCodes)
  Print #RptHandle, FF$
  Return

PrintGraphics:
  ReportFile$ = "BLRPTS\ARCODLST.RPT"  'Report File Name
  
  OpenCatCodeIdxFile IdxHandle
  NumOfIdxRecs = LOF(IdxHandle) / Len(CodeIdx)
  OpenCatCodeFile TrHandle
  TRNumRecs = LOF(TrHandle) / Len(CodeRec)
  
  If TRNumRecs <> NumOfIdxRecs Then
    frmBLMessageBoxJr.Label1.Caption = "Error: The number of categories saved and the number of categories indexed are not the same. Re-index category codes or call Southern Software at 1-800-842-8190 for assistance."
    frmBLMessageBoxJr.Label1.Top = 600
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  ReDim IdxRecs(1 To NumOfIdxRecs) As Integer
  For x = 1 To NumOfIdxRecs
    Get IdxHandle, x, CodeIdx
    IdxRecs(x) = CodeIdx.CatCodeRec
  Next x
  Close IdxHandle
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  For cnt = 1 To TRNumRecs 'Count
    Get TrHandle, IdxRecs(cnt), CodeRec
    If Left$(CodeRec.CatCode, 1) <> " " Then
      Print #RptHandle, TownName$; dlm; QPTrim$(CodeRec.CatCode); dlm;
      Print #RptHandle, QPTrim$(CodeRec.CODEDESC); dlm; GetGLNum(CodeRec.REVGLNUM); dlm;
      Print #RptHandle, GetGLNum(CodeRec.ARGLACCT); dlm; GetGLNum(CodeRec.CASHACCT)
      TotalCodes = TotalCodes + 1
    End If
  Next cnt
  Close         'Close all open files now

  arBLCatListRpt.Show
  frmBLLoadReport.Show
  
  Return

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustReportsMenu", "cmdCatList_Click", Erl)
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

Private Sub cmdCustList_Click()
  If Not Exist("arcust.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Error: No customers have been saved. Form load aborted."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  frmBLCustListRpt.Show
  DoEvents
  Unload frmBLCustReportsMenu
End Sub

Private Sub cmdCustTransHist_Click()
  If Not Exist("arcust.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Error: No customers have been saved. Form load aborted."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  frmBLCustTransHist.Show
  DoEvents
  Unload frmBLCustReportsMenu
End Sub

Private Sub cmdExit_Click()
  KillFile "custrptsmenu.dat"
  frmBLMainMenu.Show
  DoEvents
  Unload frmBLCustReportsMenu
End Sub

Private Sub cmdExprLicList_Click()
  If Not Exist("arcust.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Error: No customers have been saved. Form load aborted."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  frmBLXLicList.Show
  DoEvents
  Unload frmBLCustReportsMenu
End Sub

Private Sub cmdInOut_Click()
  If Not Exist("arcust.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Error: No customers have been saved. Form load aborted."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  frmBLInOutRpt.Show
  DoEvents
  Unload frmBLCustReportsMenu
End Sub

Private Sub cmdLicList_Click()
  If Not Exist("arcust.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Error: No customers have been saved. Form load aborted."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  frmBLLicListRpt.Show
  DoEvents
  Unload frmBLCustReportsMenu
End Sub

Private Sub cmdQuick_Click()
  If Not Exist("arcust.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Error: No customers have been saved. Form load aborted."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  frmBLQuickList.Show
  DoEvents
  Unload frmBLCustReportsMenu
End Sub

Private Sub cmdMailLbls_Click()
  If Not Exist("arcust.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Error: No customers have been saved. Form load aborted."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  frmBLMailLbls.Show
  DoEvents
  Unload frmBLCustMaintMenu

End Sub

Private Sub cmdTransJrnl_Click()
  If Not Exist("arcust.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Error: No customers have been saved. Form load aborted."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  frmBLTransJournal.Show
  DoEvents
  Unload frmBLCustReportsMenu
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
  Dim One As Integer
  Dim DHandle As Integer
  
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
'  If Exist("inquiry.dat") Then KillFile "inquiry.dat" 'this .dat
  'file is created when the customer inquiry button is pressed...
  'when the customer lookup form closes and it returns to this form
  'this form will be reloaded and this file will be deleted
  One = 1
  DHandle = FreeFile
  Open "custrptsmenu.dat" For Output As DHandle Len = 2
  Print #DHandle, One
  Close DHandle


End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      ClearInUse PWcnt
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLCustReportsMenu.")
      Call Terminate
      End
    End If
  End If
End Sub

