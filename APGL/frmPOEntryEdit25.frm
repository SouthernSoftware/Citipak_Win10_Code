VERSION 5.00
Object = "{48932A52-981F-101B-A7FB-4A79242FD97B}#3.1#0"; "TAB32X30.OCX"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "IMP32X30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPOEntryEdit25 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order Entry/Edit"
   ClientHeight    =   8844
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   12192
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8844
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin TabproLib.vaTabPro vaTabPro1 
      Height          =   6060
      Left            =   918
      TabIndex        =   0
      Top             =   960
      Width           =   10356
      _Version        =   196609
      _ExtentX        =   18267
      _ExtentY        =   10689
      _StockProps     =   100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   9405029
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabsPerRow      =   1
      TabCount        =   1
      ThreeD          =   0   'False
      ShowFocusRect   =   0   'False
      ActiveTabBold   =   0   'False
      OffsetFromClientTop=   -1  'True
      BookRingShowHole=   -1  'True
      PageMax         =   2
      DataFormat      =   ""
      PageEarMarkAlignNext=   1
      BookCornerGuardWidth=   108
      BookCornerGuardLength=   396
      ThreeDInnerWidthActive=   0
      DrawFocusRect   =   1
      DataField       =   ""
      TabCaption      =   "frmPOEntryEdit25.frx":0000
      PageEarMarkPictureNext=   "frmPOEntryEdit25.frx":0154
      PageEarMarkPicturePrev=   "frmPOEntryEdit25.frx":0170
      EarMarkPictureNext=   "frmPOEntryEdit25.frx":018C
      EarMarkPicturePrev=   "frmPOEntryEdit25.frx":01A8
      Begin ImpproLib.vaImprint vaImprint1 
         Height          =   5844
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   10356
         _Version        =   196609
         _ExtentX        =   18267
         _ExtentY        =   10308
         _StockProps     =   70
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         AutoSize        =   1
         Picture         =   "frmPOEntryEdit25.frx":01C4
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "F10 &Save"
      Height          =   492
      Left            =   6720
      TabIndex        =   3
      Top             =   7344
      Width           =   1332
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Esc E&xit"
      Height          =   492
      Left            =   10110
      TabIndex        =   2
      Top             =   7344
      Width           =   1332
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "F3 &Delete"
      Enabled         =   0   'False
      Height          =   492
      Left            =   8415
      TabIndex        =   1
      Top             =   7344
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   8484
      Width           =   12192
      _ExtentX        =   21505
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7133
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "1:38 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "2/15/02"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Enter/Edit Purchase Orders"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4092
      TabIndex        =   5
      Top             =   312
      Width           =   4020
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   636
      Left            =   2580
      Top             =   216
      Width           =   7020
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   2592
      Top             =   72
      Width           =   7020
   End
End
Attribute VB_Name = "frmPOEntryEdit25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

