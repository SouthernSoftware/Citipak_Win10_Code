VERSION 5.00
Begin VB.Form frmSpell 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SpellCheck"
   ClientHeight    =   2685
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   Begin VB.Frame lstframe 
      BackColor       =   &H00800000&
      Caption         =   "Spelling Suggestions"
      ForeColor       =   &H0000FFFF&
      Height          =   2685
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton cmdAdd 
         Caption         =   "^ Add"
         Height          =   285
         Left            =   3075
         TabIndex        =   8
         Top             =   2325
         Width           =   570
      End
      Begin VB.TextBox txtUserSug 
         Height          =   285
         Left            =   135
         TabIndex        =   7
         Top             =   2325
         Width           =   2925
      End
      Begin VB.ListBox lstsuggestions 
         Height          =   1815
         Left            =   135
         TabIndex        =   4
         Top             =   480
         Width           =   3495
      End
      Begin VB.CommandButton CmdClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   3765
         TabIndex        =   3
         Top             =   2235
         Width           =   1095
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "Change"
         Height          =   375
         Left            =   3765
         TabIndex        =   2
         Top             =   465
         Width           =   1095
      End
      Begin VB.CommandButton cmdSkip 
         Caption         =   "Skip"
         Height          =   375
         Left            =   3765
         TabIndex        =   1
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label checkword 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   15
      Top             =   2745
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmSpell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub cmdAdd_Click()
    lstsuggestions.AddItem txtUserSug, lstsuggestions.ListCount
    lstsuggestions.ListIndex = lstsuggestions.ListCount - 1
End Sub

Private Sub cmdChange_Click()
SpellChange
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSkip_Click()
SkipWord
End Sub

Private Sub Form_Load()
blnFrmSpellShowing = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
blnFrmSpellShowing = False
Set frmSpell = Nothing
End Sub

Private Sub Timer1_Timer()
SetAlwaysOnTop Me
Timer1.Enabled = False
End Sub
