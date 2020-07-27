VERSION 5.00
Begin VB.Form frmCitiCancel 
   BackColor       =   &H000000C0&
   BorderStyle     =   0  'None
   Caption         =   "CitiPak"
   ClientHeight    =   2460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5268
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5268
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   1980
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   1308
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Citipak Has Been Terminated.  Contact Software Support."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   828
      Left            =   552
      TabIndex        =   0
      Top             =   672
      Width           =   4164
   End
End
Attribute VB_Name = "frmCitiCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOk_Click()
  ClearInUse PWcnt
  Unload frmCitiCancel
  End
End Sub
