VERSION 5.00
Begin VB.Form frmLoading 
   BackColor       =   &H80000018&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1644
   ClientLeft      =   36
   ClientTop       =   36
   ClientWidth     =   3744
   ControlBox      =   0   'False
   FillColor       =   &H8000000A&
   Icon            =   "frmLoading.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1644
   ScaleWidth      =   3744
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading. . ."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1254
      TabIndex        =   0
      Top             =   636
      Width           =   1236
   End
End
Attribute VB_Name = "frmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

