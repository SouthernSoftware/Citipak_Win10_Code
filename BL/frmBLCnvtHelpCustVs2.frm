VERSION 5.00
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "IMP32X30.OCX"
Begin VB.Form frmBLCnvtHelpCustVs2 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Conversion: Help For Customer Conversion Version #2"
   ClientHeight    =   7656
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   10932
   Icon            =   "frmBLCnvtHelpCustVs2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7656
   ScaleWidth      =   10932
   StartUpPosition =   2  'CenterScreen
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   7308
      Left            =   228
      TabIndex        =   0
      Top             =   168
      Width           =   10476
      _Version        =   196609
      _ExtentX        =   18478
      _ExtentY        =   12890
      _StockProps     =   70
      BackColor       =   12648447
      Caption         =   ""
      Picture         =   "frmBLCnvtHelpCustVs2.frx":08CA
      Begin VB.CommandButton cmdExit 
         Caption         =   "ESC Exit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   4176
         TabIndex        =   1
         Top             =   6336
         Width           =   2364
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "If category codes are not being converted then all category data for customers will be deleted."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   588
         Left            =   144
         TabIndex        =   10
         Top             =   1728
         Width           =   9660
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "4. Version #2 customer code conversion makes all customers active."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   348
         Left            =   336
         TabIndex        =   9
         Top             =   4416
         Width           =   10044
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "VERSION #2 CUSTOMER CODE CONVERSION"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   288
         TabIndex        =   8
         Top             =   480
         Width           =   9804
      End
      Begin VB.Label Label9 
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmBLCnvtHelpCustVs2.frx":08E6
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   828
         Left            =   336
         TabIndex        =   7
         Top             =   816
         Width           =   9708
      End
      Begin VB.Label Label11 
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "1. Version #2 customer code conversion looks for and reports blank license numbers."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   336
         TabIndex        =   6
         Top             =   2544
         Width           =   9804
      End
      Begin VB.Label Label10 
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmBLCnvtHelpCustVs2.frx":0983
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   540
         Left            =   336
         TabIndex        =   5
         Top             =   2976
         Width           =   10044
      End
      Begin VB.Label Label7 
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmBLCnvtHelpCustVs2.frx":0A0A
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   540
         Left            =   336
         TabIndex        =   4
         Top             =   3648
         Width           =   10044
      End
      Begin VB.Label Label4 
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "5. Version #2 customer code conversion indexes customer billing names, license numbers, sort names, customer numbers."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   540
         Left            =   336
         TabIndex        =   3
         Top             =   4896
         Width           =   10044
      End
      Begin VB.Label Label5 
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "6. Version #2 customer code conversion zeroes out all customer balances."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   336
         TabIndex        =   2
         Top             =   5568
         Width           =   10044
      End
   End
End
Attribute VB_Name = "frmBLCnvtHelpCustVs2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      Call cmdExit_Click
      KeyCode = 0
  End Select

End Sub


