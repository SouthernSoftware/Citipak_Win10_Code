VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUBCustMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Maint Menu"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   ClipControls    =   0   'False
   Icon            =   "FrmCustMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCustInfo 
      Caption         =   "Customer &Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   3828
      TabIndex        =   5
      Top             =   4740
      Width           =   4524
   End
   Begin VB.CommandButton cmdMeterCoords 
      Caption         =   "&Edit Meter Coordinates"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   3828
      TabIndex        =   2
      Top             =   3072
      Width           =   4524
   End
   Begin VB.CommandButton cmdCustQuickRate 
      BackColor       =   &H008F8265&
      Caption         =   "Quick Customer Listing by Rate"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   3828
      MaskColor       =   &H8000000F&
      TabIndex        =   9
      Top             =   6960
      Width           =   4524
   End
   Begin VB.CommandButton cmdCustConsumpRpt 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Customer &Consumption History"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   3828
      TabIndex        =   6
      Top             =   5292
      Width           =   4524
   End
   Begin VB.CommandButton cmdAddCustomer 
      Caption         =   "&Add a New Customer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   3828
      TabIndex        =   0
      Top             =   1968
      Width           =   4524
   End
   Begin VB.CommandButton cmdEditCustomer 
      Caption         =   "&Edit an Existing Customer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   3828
      TabIndex        =   1
      Top             =   2520
      Width           =   4524
   End
   Begin VB.CommandButton cmdSetCustFinal 
      Caption         =   "&Set a Customer to Final"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   3828
      TabIndex        =   3
      Top             =   3636
      Width           =   4524
   End
   Begin VB.CommandButton cmdCustQuickRpt 
      Caption         =   "Quick Customer Listing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   3828
      TabIndex        =   8
      Top             =   6396
      Width           =   4524
   End
   Begin VB.CommandButton cmdDeleteCustomer 
      Caption         =   "&Delete a Customer Account"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   3828
      TabIndex        =   4
      Top             =   4188
      Width           =   4524
   End
   Begin VB.CommandButton cmdCustTransRpt 
      Caption         =   "Customer Transaction History"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   3828
      TabIndex        =   7
      Top             =   5844
      Width           =   4524
   End
   Begin VB.CommandButton cmdExitCustomerMenu 
      Caption         =   "E&xit to Previous Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   3828
      TabIndex        =   10
      Top             =   7512
      Width           =   4524
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21548
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7154
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "9:29 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "12/20/2004"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "CUSTOMER MAINT MENU"
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
      Left            =   3348
      TabIndex        =   11
      Top             =   1176
      Width           =   5292
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1776
      Top             =   768
      Width           =   8652
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1776
      Top             =   648
      Width           =   8652
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   9828
      X2              =   9828
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   8868
      X2              =   8868
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   3
      X1              =   8868
      X2              =   9828
      Y1              =   1944
      Y2              =   1944
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   8868
      X2              =   9828
      Y1              =   2064
      Y2              =   2064
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   8868
      Top             =   1824
      Width           =   972
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   8988
      X2              =   9708
      Y1              =   7944
      Y2              =   7944
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   8988
      X2              =   8988
      Y1              =   2064
      Y2              =   7944
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8988
      Top             =   2064
      Width           =   732
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   3348
      X2              =   3348
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2388
      X2              =   2388
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   2388
      X2              =   3348
      Y1              =   1944
      Y2              =   1944
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2388
      X2              =   3348
      Y1              =   2064
      Y2              =   2064
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2388
      Top             =   1824
      Width           =   972
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2508
      X2              =   3228
      Y1              =   7944
      Y2              =   7944
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2508
      X2              =   2508
      Y1              =   2064
      Y2              =   7944
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2508
      Top             =   2064
      Width           =   732
   End
End
Attribute VB_Name = "frmUBCustMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim fromform As Form, toform As Form, codeopt As Integer
Public Sub Wheretogo(xfrm As Form, tfrm As Form, Optional opt As Integer)
  Set fromform = xfrm
  Set toform = tfrm
  If opt <> 0 Then
    codeopt = opt
  Else
    codeopt = 0
  End If
End Sub

Private Sub cmdAddCustomer_Click()
  DeActivateControls Me
  frmInfo.Label1 = "Loading. . ."
  frmInfo.Show
  DoEvents
  frmCustAddEdit.LabelAcctNo = " ???"
  frmCustAddEdit.Wheretogo frmUBCustMenu, frmUBCustMenu, 0
  frmCustAddEdit.Show
  Unload frmInfo
  DoEvents
  ActivateControls Me
  Unload frmUBCustMenu
End Sub

Private Sub cmdCustInfo_Click()
  frmCustEditLookUP.Caption = "Customer Information"
  frmCustEditLookUP.Label1.Caption = "Customer Information"
  frmCustEditLookUP.Wheretogo frmUBCustMenu, frmUBCustMenu, , 7
  DoEvents
  frmCustEditLookUP.Show
  'Unload frmUBCustMenu
End Sub

Private Sub cmdCustQuickRate_Click()
  Load frmRptCustbyRate
  DoEvents
  frmRptCustbyRate.Show
  Unload frmUBCustMenu
End Sub

Private Sub cmdCustQuickRpt_Click()
  Load frmRptCustList
  DoEvents
  frmRptCustList.Show
  Unload frmUBCustMenu
End Sub

Private Sub cmdCustConsumpRpt_Click()
  frmCustEditLookUP.Caption = "Customer Consumption History"
  frmCustEditLookUP.Label1.Caption = "Customer Consumption History"
  frmCustEditLookUP.Wheretogo frmUBCustMenu, frmRptCustConsHist
  DoEvents
  frmCustEditLookUP.Show
  Unload frmUBCustMenu
'  Load frmRptCustHistory
'  frmRptCustHistory.RptType = True
'  DoEvents
'  frmRptCustHistory.Caption = "Customer Consumption History"
'  frmRptCustHistory.Label1 = frmRptCustHistory.Caption
'  frmRptCustHistory.fpDetailFlag.Visible = False
'  frmRptCustHistory.DetailLabel.Visible = False
'  frmRptCustHistory.Wheretogo frmUBCustMenu, frmRptCustHistory
'  frmRptCustHistory.Show
'  Unload frmUBCustMenu
End Sub

Private Sub cmdCustTransRpt_Click()
  frmCustEditLookUP.Caption = "Customer Transaction History"
  frmCustEditLookUP.Label1.Caption = "Customer Transaction History"
  frmCustEditLookUP.Wheretogo frmUBCustMenu, frmRptCustTranHist
  DoEvents
  frmCustEditLookUP.Show
  Unload frmUBCustMenu

'  Load frmRptCustHistory
'  frmRptCustHistory.RptType = False
'  DoEvents
'  frmRptCustHistory.Caption = "Customer Transaction History"
'  frmRptCustHistory.Label1 = frmRptCustHistory.Caption
'  frmRptCustHistory.Wheretogo frmUBCustMenu, frmRptCustHistory
'  frmRptCustHistory.Show
'  Unload frmUBCustMenu
End Sub

Private Sub cmdDeleteCustomer_Click()
  frmCustEditLookUP.Caption = "Customer Delete Search"
  frmCustEditLookUP.Label1.Caption = "Customer Delete Search"
  frmCustEditLookUP.Wheretogo frmUBCustMenu, frmCustDelete, , 1
  'Load frmCustEditLookUP
  DoEvents
  frmCustEditLookUP.Show
  Unload frmUBCustMenu

End Sub

Private Sub cmdEditCustomer_Click()
  frmCustEditLookUP.Wheretogo frmUBCustMenu, frmCustAddEdit
'  Load frmCustEditLookUP
  DoEvents
  frmCustEditLookUP.Show
  Unload frmUBCustMenu
End Sub

Private Sub cmdExitCustomerMenu_Click()
  Load frmUBMainMenu
  DoEvents
  frmUBMainMenu.Show
  Unload frmUBCustMenu
End Sub

Private Sub cmdMeterCoords_Click()
  frmCustEditLookUP.Wheretogo frmUBCustMenu, frmMtrCoorEdit
  DoEvents
  frmCustEditLookUP.Show
  Unload frmUBCustMenu
End Sub

Private Sub cmdSetCustFinal_Click()
  Dim FntSize As Integer
  ReDim MsgText(0 To 5) As String
  If Not Exist(UBPath$ + "UBCust.dat") Then
    frmMsgDialog.RetLabel = "-2"
    UBLog "ERROR: NO Cust Info"
    FntSize = frmMsgDialog.Label(3).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = "NO CUSTOMER INFORMATION!"
    MsgText(3) = ""
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
    Exit Sub
  End If

 'Need to do warning here if have not posted reg bills!!!!
   If Exist(UBPath$ + "UBBILLS.DAT") And Exist(UBPath$ + "UBBILLS.PRN") Then
     UBLog "ERROR: UNPOSTED BILLING DETECTED!"
     UBLog "ASKING USER WANT TO CONTINUE?"
     FntSize = frmMsgDialog.Label(3).FontSize
     frmMsgDialog.Label(1).FontSize = (FntSize + 2)
     frmMsgDialog.Label(3).FontSize = (FntSize + 2)
     MsgText(0) = "ERROR:"
     MsgText(1) = ""
     MsgText(2) = "UNPOSTED BILLING DETECTED!"
     MsgText(3) = "Files will Be Deleted."
     MsgText(4) = "Are You Sure You Want To Continue?"
     MsgText(5) = ""
     If GetOKorNot(MsgText()) Then
       UBLog "USER WANTS TO CONTINUE!"
       KillFile (UBPath$ + "UBBILLS.PRN")
       KillFile (UBPath$ + "UBBILLS.Dat")
       UBLog "From SetFinal USER Deleted PREBILLING and BillFile."
     Else
       UBLog "SetFinal Warn of Prebill/Bills User Cancels so won't delete files"
       Exit Sub
    End If
  End If
  'This is if have printed regular prebilling warn and delete if continue
   If Exist(UBPath$ + "UBBILLS.DAT") Then
     UBLog "ERROR: REGULAR PREBILLING HAS BEEN PRINTED!"
     UBLog "ASKING USER WANT TO CONTINUE?"
     FntSize = frmMsgDialog.Label(3).FontSize
     frmMsgDialog.Label(1).FontSize = (FntSize + 2)
     frmMsgDialog.Label(3).FontSize = (FntSize + 2)
     MsgText(0) = "ERROR:"
     MsgText(1) = ""
     MsgText(2) = "REGULAR PREBILLING DETECTED!"
     MsgText(3) = "File will be DELETED."
     MsgText(4) = "Are You Sure You Want To Continue?"
     MsgText(5) = ""
     If GetOKorNot(MsgText()) Then
       UBLog "From SetFinal USER Deleted PREBILLING."
       KillFile (UBPath$ + "UBBILLS.Dat")
     Else
       Exit Sub
    End If
  End If

  frmCustEditLookUP.Caption = "Set Customer to Final Search"
  frmCustEditLookUP.Label1.Caption = "Customer Set Final Search"
  frmCustEditLookUP.Wheretogo frmUBCustMenu, frmCustFinal, , 2
  'Load frmCustEditLookUP
  DoEvents
  frmCustEditLookUP.Show
  Unload frmUBCustMenu

End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  'screenW = (Screen.Width / Screen.TwipsPerPixelX)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitCustomerMenu.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via CustomerMenu by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ' Me.Visible = False
    Temp_Class.ResizeControls Me
  '  Me.Visible = True
   ' Me.SetFocus
  End If
  DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyHome
      cmdAddCustomer.SetFocus
    Case vbKeyEnd
      cmdExitCustomerMenu.SetFocus
    Case Else:
  End Select
End Sub



