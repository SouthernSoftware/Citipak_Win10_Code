VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Begin VB.Form frmTRDetailDC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Detail"
   ClientHeight    =   5472
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   8280
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   10.2
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5472
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   Begin fpBtnAtlLibCtl.fpBtn fpCmdOK 
      Height          =   408
      Left            =   4380
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4728
      Width           =   1668
      _Version        =   131072
      _ExtentX        =   2942
      _ExtentY        =   720
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
      DrawFocusRect   =   1
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
      ButtonDesigner  =   "frmTRDetailDC.frx":0000
   End
   Begin fpBtnAtlLibCtl.fpBtn fpcmdPrint 
      Height          =   408
      Left            =   2388
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4728
      Width           =   1668
      _Version        =   131072
      _ExtentX        =   2942
      _ExtentY        =   720
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
      DrawFocusRect   =   1
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
      ButtonDesigner  =   "frmTRDetailDC.frx":01D7
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Balance After Trans"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   13
      Left            =   3936
      TabIndex        =   30
      Top             =   1512
      Width           =   2268
   End
   Begin VB.Label Bal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   6312
      TabIndex        =   29
      Top             =   1512
      Width           =   1236
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Charge"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   12
      Left            =   696
      TabIndex        =   28
      Top             =   2088
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Check"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   11
      Left            =   696
      TabIndex        =   27
      Top             =   1800
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Decal Category"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   10
      Left            =   168
      TabIndex        =   26
      Top             =   2688
      Width           =   1908
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Operator"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   9
      Left            =   528
      TabIndex        =   25
      Top             =   912
      Width           =   1548
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Cash"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   8
      Left            =   696
      TabIndex        =   24
      Top             =   1512
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Sticker"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   7
      Left            =   528
      TabIndex        =   23
      Top             =   4032
      Width           =   1548
   End
   Begin VB.Label Label6 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5448
      TabIndex        =   21
      Top             =   576
      Width           =   2652
   End
   Begin VB.Label Oper 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   2184
      TabIndex        =   20
      Top             =   912
      Width           =   300
   End
   Begin VB.Label Charge 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   2184
      TabIndex        =   19
      Top             =   2088
      Width           =   1236
   End
   Begin VB.Label Chk 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   2184
      TabIndex        =   18
      Top             =   1800
      Width           =   1236
   End
   Begin VB.Label Cash 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   2184
      TabIndex        =   17
      Top             =   1512
      Width           =   1236
   End
   Begin VB.Label Label6b 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5448
      TabIndex        =   16
      Top             =   912
      Width           =   2652
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5448
      TabIndex        =   15
      Top             =   240
      Width           =   2652
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2184
      TabIndex        =   14
      Top             =   576
      Width           =   1260
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2184
      TabIndex        =   13
      Top             =   240
      Width           =   1260
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Type:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   6
      Left            =   3528
      TabIndex        =   12
      Top             =   576
      Width           =   1812
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "VIN#/Desc:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   3528
      TabIndex        =   11
      Top             =   240
      Width           =   1812
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Amount:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   264
      TabIndex        =   10
      Top             =   576
      Width           =   1812
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Transaction Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   264
      TabIndex        =   9
      Top             =   240
      Width           =   1812
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Expire Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   528
      TabIndex        =   8
      Top             =   3696
      Width           =   1548
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "State Tag"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Top             =   3360
      Width           =   1596
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Make/Model"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   48
      TabIndex        =   6
      Top             =   3024
      Width           =   2028
   End
   Begin VB.Label Sticker 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   2184
      TabIndex        =   4
      Top             =   4032
      Width           =   2028
   End
   Begin VB.Label Expire 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   2184
      TabIndex        =   3
      Top             =   3684
      Width           =   2028
   End
   Begin VB.Label State 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   2184
      TabIndex        =   2
      Top             =   3348
      Width           =   2532
   End
   Begin VB.Label Make 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   2184
      TabIndex        =   1
      Top             =   3024
      Width           =   3732
   End
   Begin VB.Label Cat 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   2184
      TabIndex        =   0
      Top             =   2688
      Width           =   1932
   End
End
Attribute VB_Name = "frmTRDetailDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then
  KeyCode = 0
  Call fpCmdOk_Click
  End If
End Sub

Private Sub fpCmdOk_Click()
  DoEvents
  Unload frmTRDetailDC
End Sub

Private Sub fpCmdOK_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Button = 0
  Call fpCmdOk_Click
End Sub

Private Sub fpCmdPrint_Click()
  Dim ReportFile As String, DCRpt As Integer, cnt As Integer, go2line As Integer
  Dim gofrom As Integer
  ReportFile$ = UBPath$ + "DCTRDetl.RPT"
  DCRpt = FreeFile
  Open ReportFile$ For Output As DCRpt
  Print #DCRpt, ""
  Print #DCRpt, Now
  Print #DCRpt, Tab(2); QPTrim$(frmTRDetail.Caption)
  Print #DCRpt, "-------------------------------------------------------------------------"
  Print #DCRpt, Tab(2); "Transaction Date: "; Label3.Caption; Tab(44); "Vin#/Desc: "; Label5.Caption
  Print #DCRpt, Tab(2); "    Total Amount: "; Label4.Caption; Tab(44); "Type: "; Label6.Caption
  Print #DCRpt, Tab(2); "        Operator: "; Oper.Caption; Tab(51); Label6b.Caption
  Print #DCRpt,
  Print #DCRpt, Tab(2); "            Cash: "; Cash.Caption; Tab(40); "Balance After Trans: "; Bal.Caption
  Print #DCRpt, Tab(2); "           Check: "; Chk.Caption
  Print #DCRpt, Tab(2); "          Charge: "; Charge.Caption
  Print #DCRpt,
  Print #DCRpt, Tab(2); "  Decal Category: "; Cat.Caption
  Print #DCRpt, Tab(2); "      Make/Model: "; Make.Caption
  Print #DCRpt, Tab(2); "       State Tag: "; State.Caption
  Print #DCRpt, Tab(2); "     Expire Date: "; Expire.Caption
  Print #DCRpt, Tab(2); "         Sticker: "; Sticker.Caption
  Close #DCRpt
  PrintTRDetlScreenDC

End Sub
