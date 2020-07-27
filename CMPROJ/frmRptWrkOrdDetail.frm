VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmRptWrkOrdDetail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Work Order Detail"
   ClientHeight    =   6336
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   8988
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6336
   ScaleWidth      =   8988
   StartUpPosition =   2  'CenterScreen
   Begin fpBtnAtlLibCtl.fpBtn CmdOk 
      Height          =   480
      Left            =   3852
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5712
      Width           =   1284
      _Version        =   131072
      _ExtentX        =   2265
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmRptWrkOrdDetail.frx":0000
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "6)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   72
      TabIndex        =   34
      Top             =   5184
      Width           =   276
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "5)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   72
      TabIndex        =   33
      Top             =   4860
      Width           =   276
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "4)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   72
      TabIndex        =   32
      Top             =   4524
      Width           =   276
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "3)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   72
      TabIndex        =   31
      Top             =   4188
      Width           =   276
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "2)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   72
      TabIndex        =   30
      Top             =   3864
      Width           =   276
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "1)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   72
      TabIndex        =   29
      Top             =   3528
      Width           =   276
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "6)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   72
      TabIndex        =   28
      Top             =   2928
      Width           =   276
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "5)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   72
      TabIndex        =   27
      Top             =   2600
      Width           =   276
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "4)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   72
      TabIndex        =   26
      Top             =   2268
      Width           =   276
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "3)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   72
      TabIndex        =   25
      Top             =   1936
      Width           =   276
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "2)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   72
      TabIndex        =   24
      Top             =   1604
      Width           =   276
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "1)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   72
      TabIndex        =   23
      Top             =   1272
      Width           =   276
   End
   Begin VB.Label LabelR6 
      Caption         =   "Rem 6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   276
      Left            =   336
      TabIndex        =   22
      Top             =   5160
      Width           =   8556
   End
   Begin VB.Label LabelR5 
      Caption         =   "Rem 5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   276
      Left            =   336
      TabIndex        =   21
      Top             =   4836
      Width           =   8556
   End
   Begin VB.Label LabelR4 
      Caption         =   "Rem 4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   276
      Left            =   336
      TabIndex        =   20
      Top             =   4524
      Width           =   8556
   End
   Begin VB.Label LabelR3 
      Caption         =   "Rem 3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   276
      Left            =   336
      TabIndex        =   19
      Top             =   4212
      Width           =   8556
   End
   Begin VB.Label LabelR2 
      Caption         =   "Rem 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   276
      Left            =   336
      TabIndex        =   18
      Top             =   3888
      Width           =   8556
   End
   Begin VB.Label LabelR1 
      Caption         =   "Rem 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   276
      Left            =   336
      TabIndex        =   17
      Top             =   3552
      Width           =   8556
   End
   Begin VB.Label LabelI6 
      Caption         =   "Inf 6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   276
      Left            =   336
      TabIndex        =   16
      Top             =   2928
      Width           =   8556
   End
   Begin VB.Label LabelI5 
      Caption         =   "Inf 5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   276
      Left            =   336
      TabIndex        =   15
      Top             =   2592
      Width           =   8556
   End
   Begin VB.Label LabelI4 
      Caption         =   "Inf 4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   276
      Left            =   336
      TabIndex        =   14
      Top             =   2268
      Width           =   8556
   End
   Begin VB.Label LabelI3 
      Caption         =   "Inf 3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   300
      Left            =   336
      TabIndex        =   13
      Top             =   1932
      Width           =   8556
   End
   Begin VB.Label LabelI2 
      Caption         =   "Inf 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   276
      Left            =   336
      TabIndex        =   12
      Top             =   1608
      Width           =   8556
   End
   Begin VB.Label LabelI1 
      Caption         =   "Inf 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   276
      Left            =   336
      TabIndex        =   11
      Top             =   1272
      Width           =   8556
   End
   Begin VB.Label LabelCompBy 
      Caption         =   "Complete By"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   276
      Left            =   6720
      TabIndex        =   10
      Top             =   672
      Width           =   1716
   End
   Begin VB.Label LabelCompDate 
      Caption         =   "Completed Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   276
      Left            =   6720
      TabIndex        =   9
      Top             =   360
      Width           =   1716
   End
   Begin VB.Label LabelEntryDate 
      Caption         =   "Entry Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   276
      Left            =   2496
      TabIndex        =   8
      Top             =   648
      Width           =   1740
   End
   Begin VB.Label LabelWONum 
      Caption         =   "wono"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   276
      Left            =   2496
      TabIndex        =   7
      Top             =   336
      Width           =   1164
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   8880
      Y1              =   5544
      Y2              =   5544
   End
   Begin VB.Line Line2 
      X1              =   1200
      X2              =   8808
      Y1              =   3432
      Y2              =   3432
   End
   Begin VB.Line Line1 
      X1              =   2760
      X2              =   8808
      Y1              =   1152
      Y2              =   1152
   End
   Begin VB.Label Labe54 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Order Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   300
      Left            =   168
      TabIndex        =   6
      Top             =   1008
      Width           =   2652
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Order No:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   252
      Left            =   744
      TabIndex        =   5
      Top             =   360
      Width           =   1812
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Entry Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   300
      Left            =   1176
      TabIndex        =   4
      Top             =   672
      Width           =   1284
   End
   Begin VB.Label Labe56 
      BackStyle       =   0  'Transparent
      Caption         =   "Completed Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   300
      Left            =   4872
      TabIndex        =   3
      Top             =   360
      Width           =   1764
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   300
      Left            =   168
      TabIndex        =   2
      Top             =   3288
      Width           =   1212
   End
   Begin VB.Label Label42 
      BackStyle       =   0  'Transparent
      Caption         =   "Complete By Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   300
      Left            =   4680
      TabIndex        =   1
      Top             =   672
      Width           =   1956
   End
End
Attribute VB_Name = "frmRptWrkOrdDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then
    KeyCode = 0
    DoEvents
    Call cmdOk_Click
  Else
    KeyCode = 0
    DoEvents
    Call cmdOk_Click
  End If
End Sub

Private Sub cmdOk_Click()
  DoEvents
  Unload frmRptWrkOrdDetail
End Sub

Private Sub fpCmdOK_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Button = 0
  Call cmdOk_Click
End Sub
