VERSION 5.00
Begin VB.Form frmPWold 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GLUtility LogIn"
   ClientHeight    =   2508
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   5592
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2508
   ScaleWidth      =   5592
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Esc &Cancel"
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
      Left            =   3024
      TabIndex        =   2
      Top             =   1488
      Width           =   1308
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "F10 &Enter"
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
      Left            =   1200
      TabIndex        =   1
      Top             =   1488
      Width           =   1308
   End
   Begin VB.TextBox txtGLUtilPassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      IMEMode         =   3  'DISABLE
      Left            =   2976
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   720
      Width           =   1404
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   1104
      TabIndex        =   3
      Top             =   816
      Width           =   1692
   End
End
Attribute VB_Name = "frmPWold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
