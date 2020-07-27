VERSION 5.00
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "IMP32X30.OCX"
Object = "{48932A52-981F-101B-A7FB-4A79242FD97B}#3.1#0"; "TAB32X30.OCX"
Begin VB.Form frmCustAddEdit 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   Icon            =   "frmCustAddEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin TabproLib.vaTabPro vaTabPro1 
      Height          =   7140
      Left            =   504
      TabIndex        =   0
      Top             =   576
      Width           =   11268
      _Version        =   196609
      _ExtentX        =   19876
      _ExtentY        =   12594
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabHeight       =   0
      Tab             =   3
      TabShape        =   3
      ApplyTo         =   2
      OffsetFromClientTop=   -1  'True
      ChamferedWidth  =   0
      ChamferedHeight =   0
      DataFormat      =   ""
      BookCornerGuardWidth=   108
      BookCornerGuardLength=   396
      ThreeDOuterWidth=   0
      ThreeDOuterWidthActive=   0
      ThreeDInnerWidth=   0
      ThreeDInnerWidthActive=   0
      DataField       =   ""
      TabCaption      =   "frmCustAddEdit.frx":030A
      PageEarMarkPictureNext=   "frmCustAddEdit.frx":059A
      PageEarMarkPicturePrev=   "frmCustAddEdit.frx":05B6
      EarMarkPictureNext=   "frmCustAddEdit.frx":05D2
      EarMarkPicturePrev=   "frmCustAddEdit.frx":05EE
      Begin ImpproLib.vaImprint vaImprint4 
         Height          =   7008
         Left            =   36
         TabIndex        =   4
         Top             =   48
         Width           =   11196
         _Version        =   196609
         _ExtentX        =   19748
         _ExtentY        =   12361
         _StockProps     =   70
         Caption         =   "vaImprint4"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmCustAddEdit.frx":060A
      End
      Begin ImpproLib.vaImprint vaImprint3 
         Height          =   7008
         Left            =   -23232
         TabIndex        =   3
         Top             =   -19056
         Width           =   11196
         _Version        =   196609
         _ExtentX        =   19748
         _ExtentY        =   12361
         _StockProps     =   70
         Caption         =   "vaImprint3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Picture         =   "frmCustAddEdit.frx":0626
      End
      Begin ImpproLib.vaImprint vaImprint2 
         Height          =   7008
         Left            =   -23232
         TabIndex        =   2
         Top             =   -19056
         Width           =   11196
         _Version        =   196609
         _ExtentX        =   19748
         _ExtentY        =   12361
         _StockProps     =   70
         Caption         =   "vaImprint2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Picture         =   "frmCustAddEdit.frx":0642
      End
      Begin ImpproLib.vaImprint vaImprint1 
         Height          =   7008
         Left            =   -23232
         TabIndex        =   1
         Top             =   -19056
         Width           =   11196
         _Version        =   196609
         _ExtentX        =   19748
         _ExtentY        =   12361
         _StockProps     =   70
         Caption         =   "vaImprint1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Picture         =   "frmCustAddEdit.frx":065E
      End
   End
End
Attribute VB_Name = "frmCustAddEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

