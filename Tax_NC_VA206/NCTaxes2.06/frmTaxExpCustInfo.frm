VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmTaxExpCustInfo 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Customer Information Export"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "frmTaxExpCustInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkCoHardNum 
      BackColor       =   &H008F8265&
      Caption         =   "Co Hard Num"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   1200
      TabIndex        =   14
      Top             =   7560
      Width           =   1995
   End
   Begin VB.CheckBox chkPersCnty 
      BackColor       =   &H008F8265&
      Caption         =   "Cust County"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   8640
      TabIndex        =   45
      Top             =   7200
      Width           =   1908
   End
   Begin VB.CheckBox chkCustName 
      BackColor       =   &H008F8265&
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   1200
      TabIndex        =   8
      Top             =   5400
      Width           =   1908
   End
   Begin VB.CheckBox chkPersPin 
      BackColor       =   &H008F8265&
      Caption         =   "Pers Prop Pin"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   8640
      TabIndex        =   44
      Top             =   6840
      Width           =   1908
   End
   Begin VB.CheckBox chkRealPin 
      BackColor       =   &H008F8265&
      Caption         =   "Real Prop Pin"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   8640
      TabIndex        =   43
      Top             =   6480
      Width           =   1908
   End
   Begin VB.CheckBox chkBalance 
      BackColor       =   &H008F8265&
      Caption         =   "Customer Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   8640
      TabIndex        =   42
      Top             =   6120
      Width           =   2025
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdUntagAll 
      Height          =   450
      Left            =   7680
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1470
      _Version        =   131072
      _ExtentX        =   2593
      _ExtentY        =   794
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
      ButtonDesigner  =   "frmTaxExpCustInfo.frx":08CA
   End
   Begin VB.CheckBox chkEmployer 
      BackColor       =   &H008F8265&
      Caption         =   "Employer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   8640
      TabIndex        =   31
      Top             =   5760
      Width           =   1908
   End
   Begin VB.CheckBox chkTownship 
      BackColor       =   &H008F8265&
      Caption         =   "Cust Township"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   8640
      TabIndex        =   30
      Top             =   5400
      Width           =   1908
   End
   Begin VB.CheckBox chkOptSrch 
      BackColor       =   &H008F8265&
      Caption         =   "Opt Search Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   8640
      TabIndex        =   29
      Top             =   5040
      Width           =   2025
   End
   Begin VB.CheckBox chkCycle 
      BackColor       =   &H008F8265&
      Caption         =   "Billing Cycle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   6120
      TabIndex        =   28
      Top             =   7200
      Width           =   1908
   End
   Begin VB.CheckBox chkBankrupt 
      BackColor       =   &H008F8265&
      Caption         =   "Bankrupt Y/N?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   6120
      TabIndex        =   27
      Top             =   6840
      Width           =   1908
   End
   Begin VB.CheckBox chkLateNotice 
      BackColor       =   &H008F8265&
      Caption         =   "Late Notice Y/N?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   6120
      TabIndex        =   26
      Top             =   6480
      Width           =   1908
   End
   Begin VB.CheckBox chkChrgInt 
      BackColor       =   &H008F8265&
      Caption         =   "Chrg Interest Y/N?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   6120
      TabIndex        =   25
      Top             =   6120
      Width           =   2148
   End
   Begin VB.CheckBox chkTaxExempt 
      BackColor       =   &H008F8265&
      Caption         =   "Tax Exempt Y/N?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   6120
      TabIndex        =   24
      Top             =   5760
      Width           =   2025
   End
   Begin VB.CheckBox chkPostRt 
      BackColor       =   &H008F8265&
      Caption         =   "Postal Route"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   6120
      TabIndex        =   23
      Top             =   5400
      Width           =   1908
   End
   Begin VB.CheckBox chkDelPnt 
      BackColor       =   &H008F8265&
      Caption         =   "Delivery Point"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   6120
      TabIndex        =   22
      Top             =   5040
      Width           =   1908
   End
   Begin VB.CheckBox chkOtherSSN 
      BackColor       =   &H008F8265&
      Caption         =   "Other Soc Sec #"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   3600
      TabIndex        =   21
      Top             =   7200
      Width           =   1908
   End
   Begin VB.CheckBox chkSSN 
      BackColor       =   &H008F8265&
      Caption         =   "Social Security #"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   3600
      TabIndex        =   20
      Top             =   6840
      Width           =   1908
   End
   Begin VB.CheckBox chkDriversLic 
      BackColor       =   &H008F8265&
      Caption         =   "Driver's License #"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   3600
      TabIndex        =   19
      Top             =   6480
      Width           =   2025
   End
   Begin VB.CheckBox chkZip 
      BackColor       =   &H008F8265&
      Caption         =   "Zip Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   3600
      TabIndex        =   18
      Top             =   6120
      Width           =   1908
   End
   Begin VB.CheckBox chkState 
      BackColor       =   &H008F8265&
      Caption         =   "State"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   3600
      TabIndex        =   17
      Top             =   5760
      Width           =   1908
   End
   Begin VB.CheckBox chkCity 
      BackColor       =   &H008F8265&
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   3600
      TabIndex        =   16
      Top             =   5400
      Width           =   1908
   End
   Begin VB.CheckBox chkServiceAddress 
      BackColor       =   &H008F8265&
      Caption         =   "Service Address"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   3600
      TabIndex        =   15
      Top             =   5040
      Width           =   1908
   End
   Begin VB.CheckBox chkAddress2 
      BackColor       =   &H008F8265&
      Caption         =   "Address #2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   1200
      TabIndex        =   13
      Top             =   7200
      Width           =   1908
   End
   Begin VB.CheckBox chkAddress1 
      BackColor       =   &H008F8265&
      Caption         =   "Address #1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   1200
      TabIndex        =   12
      Top             =   6840
      Width           =   1908
   End
   Begin VB.CheckBox chkWorkPhone 
      BackColor       =   &H008F8265&
      Caption         =   "Work Phone"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   1200
      TabIndex        =   11
      Top             =   6480
      Width           =   1908
   End
   Begin VB.CheckBox chkHomePhone 
      BackColor       =   &H008F8265&
      Caption         =   "Home Phone"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   1200
      TabIndex        =   10
      Top             =   6120
      Width           =   1908
   End
   Begin VB.CheckBox chkSearchName 
      BackColor       =   &H008F8265&
      Caption         =   "Search Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   1200
      TabIndex        =   9
      Top             =   5760
      Width           =   1908
   End
   Begin VB.CheckBox chkAcctNum 
      BackColor       =   &H008F8265&
      Caption         =   "Pin/Acct Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   1200
      TabIndex        =   7
      Top             =   5040
      Width           =   1908
   End
   Begin VB.CheckBox chkFileUnique 
      BackColor       =   &H008F8265&
      Caption         =   "Unique File Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   324
      Left            =   3300
      TabIndex        =   1
      Top             =   2520
      Width           =   2412
   End
   Begin VB.CheckBox chkActive 
      BackColor       =   &H008F8265&
      Caption         =   "Active Customers Only"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   324
      Left            =   3120
      TabIndex        =   0
      Top             =   1560
      Width           =   2412
   End
   Begin VB.CheckBox chkQuotes 
      BackColor       =   &H008F8265&
      Caption         =   "Double Quotes  """
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   6900
      TabIndex        =   6
      Top             =   3720
      Width           =   2052
   End
   Begin VB.OptionButton OptDelimiter1 
      BackColor       =   &H008F8265&
      Caption         =   "Comma  ,"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   6948
      TabIndex        =   3
      Top             =   1920
      Width           =   1692
   End
   Begin VB.OptionButton OptDelimiter2 
      BackColor       =   &H008F8265&
      Caption         =   "Pipe Symbol  |"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   6948
      TabIndex        =   4
      Top             =   2292
      Width           =   1740
   End
   Begin VB.OptionButton OptDelimiter3 
      BackColor       =   &H008F8265&
      Caption         =   "Tab"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   6948
      TabIndex        =   5
      Top             =   2640
      Width           =   1692
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   492
      Left            =   3840
      TabIndex        =   32
      TabStop         =   0   'False
      Tag             =   "Press the 'Cancel' button to exit this screen and return to the main 'Business License Reports' menu."
      Top             =   7920
      Width           =   1740
      _Version        =   131072
      _ExtentX        =   3069
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmTaxExpCustInfo.frx":0AA7
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   492
      Left            =   6072
      TabIndex        =   33
      TabStop         =   0   'False
      Tag             =   $"frmTaxExpCustInfo.frx":0C85
      Top             =   7920
      Width           =   1740
      _Version        =   131072
      _ExtentX        =   3069
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmTaxExpCustInfo.frx":0D30
   End
   Begin EditLib.fpText fptxtFileName 
      Height          =   396
      Left            =   2820
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "The program will generate a unique file name and display the name in this text box."
      Top             =   2880
      Width           =   2892
      _Version        =   196608
      _ExtentX        =   5101
      _ExtentY        =   698
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   1
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
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
      ControlType     =   1
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   50
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdTagAll 
      Height          =   450
      Left            =   6120
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1470
      _Version        =   131072
      _ExtentX        =   2593
      _ExtentY        =   794
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
      ButtonDesigner  =   "frmTaxExpCustInfo.frx":0F0F
   End
   Begin VB.CheckBox chkCoStrNum 
      BackColor       =   &H008F8265&
      Caption         =   "Co String Num"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   8640
      TabIndex        =   46
      Top             =   7560
      Width           =   1995
   End
   Begin VB.Line Line13 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   8280
      X2              =   10680
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Line Line12 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   1080
      X2              =   3240
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Line Line11 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   3240
      X2              =   8280
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Line Line10 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   1080
      X2              =   1080
      Y1              =   4920
      Y2              =   7920
   End
   Begin VB.Line Line9 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   10680
      X2              =   10680
      Y1              =   4920
      Y2              =   7920
   End
   Begin VB.Line Line8 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   1080
      X2              =   10680
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line7 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   8280
      X2              =   8280
      Y1              =   4920
      Y2              =   7920
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   5760
      X2              =   5760
      Y1              =   4920
      Y2              =   7560
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   3240
      X2              =   3240
      Y1              =   4920
      Y2              =   7920
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Fields to Include in Export:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   2400
      TabIndex        =   39
      Top             =   4464
      Width           =   3732
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   9242
      X2              =   6312
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Delimiter:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   6660
      TabIndex        =   35
      Top             =   1560
      Width           =   2268
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   6300
      X2              =   2340
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   6300
      X2              =   2340
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   """TaxCustEx.ASC"" will be used if option above is not selected."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   552
      Left            =   2820
      TabIndex        =   38
      Top             =   3600
      Width           =   3012
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Select for Unique File Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   2820
      TabIndex        =   37
      Top             =   2160
      Width           =   3132
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   6300
      X2              =   6300
      Y1              =   1440
      Y2              =   4200
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2772
      Left            =   2340
      Top             =   1440
      Width           =   6912
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Text Qualifier:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   6660
      TabIndex        =   36
      Top             =   3360
      Width           =   2004
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   708
      Left            =   2940
      Top             =   480
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Export Customer Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3348
      TabIndex        =   34
      Top             =   648
      Width           =   5004
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   828
      Left            =   2940
      Top             =   360
      Width           =   5772
   End
End
Attribute VB_Name = "frmTaxExpCustInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub chkFileUnique_Click()
  Dim ThisFile$
  Dim Ext$
  Dim Cnt As Integer
  Dim chkCust$
  
  If chkFileUnique.Value = 1 Then
    Ext$ = ".ASC"
    ThisFile$ = "TXX"
    For Cnt = 1 To 5
      GetRPTName ThisFile$
      chkCust$ = ThisFile$ + Ext$
      If Exist(chkCust$) = False Then
        ThisFile$ = chkCust$
        Exit For
      End If
    Next Cnt
    fptxtFileName.Text = ThisFile$
  Else
    fptxtFileName.Text = ""
  End If
  
End Sub

Private Sub cmdProcess_Click()
  Dim q$
  Dim qc$
  Dim qcq$
  Dim Ext$, x As Long
  Dim ThisFile$, chkCust$
  Dim TaxRpt As Integer
  Dim Cnt As Integer
  Dim ThisRec As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim XCnt As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim PrintType$
  Dim ThisBal As Double
  
  If QPTrim$(fptxtFileName.Text) <> "" Then
    ThisFile$ = fptxtFileName.Text
  Else
    ThisFile$ = "TaxCustEx.ASC"
    KillFile ThisFile$
  End If
  
  TaxRpt = FreeFile
  Open ThisFile$ For Output As TaxRpt
  
  q$ = ""
  qc$ = ""
  qcq$ = ""
  If OptDelimiter1.Value = True And chkQuotes.Value <> 1 Then
    qcq$ = ","
  ElseIf OptDelimiter2.Value = True And chkQuotes.Value <> 1 Then
    qcq$ = "|"
  ElseIf OptDelimiter3.Value = True And chkQuotes.Value <> 1 Then
    qcq$ = Chr$(9)  'this is tab
  ElseIf chkQuotes.Value = 1 Then
    q$ = Chr$(34) 'this is one quote (")
    If OptDelimiter1.Value = True Then
      qc$ = q$ + ","
      qcq$ = q$ + "," + q$
    ElseIf OptDelimiter2.Value = True Then
      qc$ = q$ + "|"
      qcq$ = q$ + "|" + q$
    ElseIf OptDelimiter3.Value = True Then
      qc$ = q$ + Chr$(9)
      qcq$ = q$ + Chr$(9) + q$ 'this is tab
    End If
  End If
  GoSub DoHeaders
  
  OpenRealPropFile RHandle, NumOfRealRecs
  OpenPersPropFile PHandle, NumOfPersRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TaxCust.Deleted <> 0 Then GoTo NotThisOne
    If chkActive.Value = 1 Then
      If TaxCust.Active <> "Y" Then GoTo NotThisOne
    End If
    GoSub ExportThisOne
NotThisOne:
  Next x
  
  Close
  
  If XCnt > 0 Then
    Call Savemsg(800, "The customer export file has been saved as " + ThisFile$ + " in the Citipak folder.")
  Else
    Call TaxMsg(900, "There are no customers that fit the parameters entered.")
  End If
  
  Exit Sub
  
DoHeaders:
  If chkAcctNum.Value = 1 Then Print #TaxRpt, q$; "Account #";
  If chkCustName.Value = 1 Then Print #TaxRpt, qcq$; "Name";
  If chkSearchName.Value = 1 Then Print #TaxRpt, qcq$; "Search Name";
  If chkHomePhone.Value = 1 Then Print #TaxRpt, qcq$; "Home Phone";
  If chkWorkPhone.Value = 1 Then Print #TaxRpt, qcq$; "Work Phone";
  If chkAddress1.Value = 1 Then Print #TaxRpt, qcq$; "Address 1";
  If chkAddress2.Value = 1 Then Print #TaxRpt, qcq$; "Address 2";
  If chkCoHardNum.Value = 1 Then Print #TaxRpt, qcq$; "Co Hard Num";
  If chkServiceAddress.Value = 1 Then Print #TaxRpt, qcq$; "Service Address";
  If chkCity.Value = 1 Then Print #TaxRpt, qcq$; "City";
  If chkState.Value = 1 Then Print #TaxRpt, qcq$; "State";
  If chkZip.Value = 1 Then Print #TaxRpt, qcq$; "Zip Code";
  If chkDriversLic.Value = 1 Then Print #TaxRpt, qcq$; "Drivers License #";
  If chkSSN.Value = 1 Then Print #TaxRpt, qcq$; "Social Security #";
  If chkOtherSSN.Value = 1 Then Print #TaxRpt, qcq$; "Other Soc Sec #";
  If chkDelPnt.Value = 1 Then Print #TaxRpt, qcq$; "Delivery Point";
  If chkPostRt.Value = 1 Then Print #TaxRpt, qcq$; "Postal Route";
  If chkTaxExempt.Value = 1 Then Print #TaxRpt, qcq$; "Tax Exempt Y/N?";
  If chkChrgInt.Value = 1 Then Print #TaxRpt, qcq$; "Charge Interest Y/N?";
  If chkLateNotice.Value = 1 Then Print #TaxRpt, qcq$; "Late Notice Y/N?";
  If chkBankrupt.Value = 1 Then Print #TaxRpt, qcq$; "Bankrupt Y/N?";
  If chkCycle.Value = 1 Then Print #TaxRpt, qcq$; "Billing Cycle";
  If chkOptSrch.Value = 1 Then Print #TaxRpt, qcq$; "Optional Search Name";
  If chkTownship.Value = 1 Then Print #TaxRpt, qcq$; "Customer Township";
  If chkEmployer.Value = 1 Then Print #TaxRpt, qcq$; "Employer";
  If chkBalance.Value = 1 Then Print #TaxRpt, qcq$; "Customer Balance";
  If chkRealPin.Value = 1 Then Print #TaxRpt, qcq$; "Real Prop Pin #";
  If chkPersPin.Value = 1 Then Print #TaxRpt, qcq$; "Pers Prop Pin #";
  If chkPersCnty.Value = 1 Then Print #TaxRpt, qcq$; "Customer County";
  If chkCoStrNum.Value = 1 Then Print #TaxRpt, qcq$; "Co String Num";
  Print #TaxRpt, q$
  
  Return
  
ExportThisOne:
  If chkRealPin.Value = 1 And chkPersPin.Value = 1 Then
    PrintType = "B"
  ElseIf chkRealPin.Value = 1 And chkPersPin.Value = 0 Then
    PrintType = "R"
  ElseIf chkRealPin.Value = 0 And chkPersPin.Value = 1 Then
    PrintType = "P"
  Else
    PrintType = "N"
  End If
  
  If PrintType = "R" Or PrintType = "B" Then
    ThisRec = TaxCust.FirstPropRec
    If ThisRec > 0 Then
      Do While ThisRec > 0
        Get RHandle, ThisRec, RealRec
        XCnt = XCnt + 1
        If chkAcctNum.Value = 1 Then Print #TaxRpt, q$; QPTrim$(Str$(TaxCust.PIN));
        If chkCustName.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.CustName);
        If chkSearchName.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.SName);
        If chkHomePhone.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.HPHONE);
        If chkWorkPhone.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.WPHONE);
        If chkAddress1.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.Addr1);
        If chkAddress2.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.Addr2);
        If chkCoHardNum.Value = 1 Then Print #TaxRpt, qcq$; CStr(TaxCust.CountyAcct);
        If chkServiceAddress.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.ServiceAdd);
        If chkCity.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.City);
        If chkState.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.State);
        If chkZip.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.Zip);
        If chkDriversLic.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.DrvrsLic);
        If chkSSN.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.CSSN);
        If chkOtherSSN.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.OSSN);
        If chkDelPnt.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.DeliveryPt);
        If chkPostRt.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.PostalRt);
        If chkTaxExempt.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.TaxExempt);
        If chkChrgInt.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.Interest);
        If chkLateNotice.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.LateNotice);
        If chkBankrupt.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.Bankrupt);
        If chkCycle.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.CycleName);
        If chkOptSrch.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.OptSrchDesc);
        If chkTownship.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.TownShip);
        If chkEmployer.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.Employer);
        If chkPersCnty.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.County4BillName);
        If TaxCust.CountyAcct > 0 Then
          If chkCoStrNum.Value = 1 Then Print #TaxRpt, qcq$; CStr(TaxCust.CountyAcct);
        ElseIf QPTrim$(TaxCust.CountyAcctString) <> "" Then
          If chkCoStrNum.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.CountyAcctString);
        End If
        If chkBalance.Value = 1 Then
          ThisBal = GetCustBalance(x, -1)
          Print #TaxRpt, qcq$; Using$("$###,###,##0.00", ThisBal);
        End If
        If chkRealPin.Value = 1 Then
          Get RHandle, ThisRec, RealRec
          Print #TaxRpt, qcq$; QPTrim$(RealRec.RealPin);
        End If
        If chkPersPin.Value = 1 Then Print #TaxRpt, qcq$; "NA";
        Print #TaxRpt, q$
        ThisRec = RealRec.NextRec
      Loop
    End If
  End If
  
  If PrintType = "P" Or PrintType = "B" Then
    ThisRec = TaxCust.FirstPersRec
    If ThisRec > 0 Then
      Do While ThisRec > 0
        XCnt = XCnt + 1
        If chkAcctNum.Value = 1 Then Print #TaxRpt, q$; QPTrim$(Str$(TaxCust.PIN));
        If chkCustName.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.CustName);
        If chkSearchName.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.SName);
        If chkHomePhone.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.HPHONE);
        If chkWorkPhone.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.WPHONE);
        If chkAddress1.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.Addr1);
        If chkAddress2.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.Addr2);
        If chkCoHardNum.Value = 1 Then Print #TaxRpt, qcq$; CStr(TaxCust.CountyAcct);
        If chkServiceAddress.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.ServiceAdd);
        If chkCity.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.City);
        If chkState.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.State);
        If chkZip.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.Zip);
        If chkDriversLic.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.DrvrsLic);
        If chkSSN.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.CSSN);
        If chkOtherSSN.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.OSSN);
        If chkDelPnt.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.DeliveryPt);
        If chkPostRt.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.PostalRt);
        If chkTaxExempt.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.TaxExempt);
        If chkChrgInt.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.Interest);
        If chkLateNotice.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.LateNotice);
        If chkBankrupt.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.Bankrupt);
        If chkCycle.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.CycleName);
        If chkOptSrch.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.OptSrchDesc);
        If chkTownship.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.TownShip);
        If chkEmployer.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.Employer);
        If chkPersCnty.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.County4BillName);
'        If chkCoStrNum.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.CountyAcctString);
        If TaxCust.CountyAcct > 0 Then
          If chkCoStrNum.Value = 1 Then Print #TaxRpt, qcq$; CStr(TaxCust.CountyAcct);
        ElseIf QPTrim$(TaxCust.CountyAcctString) <> "" Then
          If chkCoStrNum.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.CountyAcctString);
        End If
        If chkBalance.Value = 1 Then
          ThisBal = GetCustBalance(x, -1)
          Print #TaxRpt, qcq$; Using$("$###,###,##0.00", ThisBal);
        End If
        If chkRealPin.Value = 1 Then Print #TaxRpt, qcq$; "NA";
        If chkPersPin.Value = 1 Then
          Get PHandle, ThisRec, PersRec
          Print #TaxRpt, qcq$; QPTrim$(PersRec.PropPin);
        End If
        Print #TaxRpt, q$
        ThisRec = PersRec.NextRec
      Loop
    End If
  End If
  
  If PrintType = "N" Then
    XCnt = XCnt + 1
    If chkAcctNum.Value = 1 Then Print #TaxRpt, q$; QPTrim$(Str$(TaxCust.PIN));
    If chkCustName.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.CustName);
    If chkSearchName.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.SName);
    If chkHomePhone.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.HPHONE);
    If chkWorkPhone.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.WPHONE);
    If chkAddress1.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.Addr1);
    If chkAddress2.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.Addr2);
    If chkCoHardNum.Value = 1 Then Print #TaxRpt, qcq$; CStr(TaxCust.CountyAcct);
    If chkServiceAddress.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.ServiceAdd);
    If chkCity.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.City);
    If chkState.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.State);
    If chkZip.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.Zip);
    If chkDriversLic.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.DrvrsLic);
    If chkSSN.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.CSSN);
    If chkOtherSSN.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.OSSN);
    If chkDelPnt.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.DeliveryPt);
    If chkPostRt.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.PostalRt);
    If chkTaxExempt.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.TaxExempt);
    If chkChrgInt.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.Interest);
    If chkLateNotice.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.LateNotice);
    If chkBankrupt.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.Bankrupt);
    If chkCycle.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.CycleName);
    If chkOptSrch.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.OptSrchDesc);
    If chkTownship.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.TownShip);
    If chkEmployer.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.Employer);
    If chkPersCnty.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.County4BillName);
'    If chkCoStrNum.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.CountyAcctString);
    If TaxCust.CountyAcct > 0 Then
      If chkCoStrNum.Value = 1 Then Print #TaxRpt, qcq$; CStr(TaxCust.CountyAcct);
    ElseIf QPTrim$(TaxCust.CountyAcctString) <> "" Then
      If chkCoStrNum.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(TaxCust.CountyAcctString);
    End If
    If chkBalance.Value = 1 Then
      ThisBal = GetCustBalance(x, -1)
      Print #TaxRpt, qcq$; Using$("$###,###,##0.00", ThisBal);
    End If
    If chkRealPin.Value = 1 Then Print #TaxRpt, qcq$; "NA";
    If chkPersPin.Value = 1 Then Print #TaxRpt, qcq$; "NA";
    Print #TaxRpt, q$
  End If
  
  Return
  
End Sub

Private Sub cmdTagAll_Click()
   chkAcctNum.Value = 1
   chkCustName.Value = 1
   chkSearchName.Value = 1
   chkHomePhone.Value = 1
   chkWorkPhone.Value = 1
   chkAddress1.Value = 1
   chkAddress2.Value = 1
   chkCoHardNum.Value = 1
   chkServiceAddress.Value = 1
   chkCity.Value = 1
   chkState.Value = 1
   chkZip.Value = 1
   chkDriversLic.Value = 1
   chkSSN.Value = 1
   chkOtherSSN.Value = 1
   chkDelPnt.Value = 1
   chkPostRt.Value = 1
   chkTaxExempt.Value = 1
   chkChrgInt.Value = 1
   chkLateNotice.Value = 1
   chkBankrupt.Value = 1
   chkCycle.Value = 1
   chkOptSrch.Value = 1
   chkTownship.Value = 1
   chkEmployer.Value = 1
   chkBalance.Value = 1
   chkRealPin.Value = 1
   chkPersPin.Value = 1
   chkPersCnty.Value = 1
   chkCoStrNum.Value = 1
End Sub

Private Sub cmdUntagAll_Click()
  chkAcctNum.Value = 0
  chkCustName.Value = 0
  chkSearchName.Value = 0
  chkHomePhone.Value = 0
  chkWorkPhone.Value = 0
  chkAddress1.Value = 0
  chkAddress2.Value = 0
  chkCoHardNum = 0
  chkServiceAddress.Value = 0
  chkCity.Value = 0
  chkState.Value = 0
  chkZip.Value = 0
  chkDriversLic.Value = 0
  chkSSN.Value = 0
  chkOtherSSN.Value = 0
  chkDelPnt.Value = 0
  chkPostRt.Value = 0
  chkTaxExempt.Value = 0
  chkChrgInt.Value = 0
  chkLateNotice.Value = 0
  chkBankrupt.Value = 0
  chkCycle.Value = 0
  chkOptSrch.Value = 0
  chkTownship.Value = 0
  chkEmployer.Value = 0
  chkBalance.Value = 0
  chkRealPin.Value = 0
  chkPersPin.Value = 0
  chkPersCnty.Value = 0
  chkCoStrNum.Value = 0
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpCustomerInfo
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxExpCustInfo.")
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub cmdExit_Click()
  frmTaxReportsMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub LoadMe()
  OptDelimiter1.Value = True
  chkActive.Value = 1
End Sub
