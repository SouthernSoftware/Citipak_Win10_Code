VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDCPurchaseMenu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter, Edit Decal Purchase"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   1890
   ClientWidth     =   12210
   Icon            =   "frmDCPurcahseMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "7:40 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "8/2/2018"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPurchase 
      Height          =   495
      Left            =   3855
      TabIndex        =   0
      Top             =   2280
      Width           =   4515
      _Version        =   131072
      _ExtentX        =   7964
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
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
      ButtonDesigner  =   "frmDCPurcahseMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDelPay 
      Height          =   495
      Left            =   3855
      TabIndex        =   1
      Top             =   3045
      Width           =   4515
      _Version        =   131072
      _ExtentX        =   7964
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
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
      ButtonDesigner  =   "frmDCPurcahseMenu.frx":0ABA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintJournalE 
      Height          =   495
      Left            =   3855
      TabIndex        =   2
      Top             =   3825
      Width           =   4515
      _Version        =   131072
      _ExtentX        =   7964
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
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
      ButtonDesigner  =   "frmDCPurcahseMenu.frx":0CB3
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintJournalN 
      Height          =   495
      Left            =   3855
      TabIndex        =   3
      Top             =   4590
      Width           =   4515
      _Version        =   131072
      _ExtentX        =   7964
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
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
      ButtonDesigner  =   "frmDCPurcahseMenu.frx":0EAA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPostPayments 
      Height          =   480
      Left            =   3855
      TabIndex        =   4
      Top             =   5370
      Width           =   4515
      _Version        =   131072
      _ExtentX        =   7964
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
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
      ButtonDesigner  =   "frmDCPurcahseMenu.frx":10A0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExitMenu 
      Height          =   495
      Left            =   3855
      TabIndex        =   6
      Top             =   6915
      Width           =   4515
      _Version        =   131072
      _ExtentX        =   7964
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
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
      ButtonDesigner  =   "frmDCPurcahseMenu.frx":1292
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDefExpire 
      Height          =   480
      Left            =   3855
      TabIndex        =   5
      Top             =   6150
      Width           =   4515
      _Version        =   131072
      _ExtentX        =   7964
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
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
      ButtonDesigner  =   "frmDCPurcahseMenu.frx":1473
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1788
      Top             =   744
      Width           =   8652
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
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2508
      X2              =   3228
      Y1              =   7944
      Y2              =   7944
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2388
      X2              =   3348
      Y1              =   2064
      Y2              =   2064
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   2400
      X2              =   3360
      Y1              =   1944
      Y2              =   1944
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2388
      X2              =   2388
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   3348
      X2              =   3348
      Y1              =   1944
      Y2              =   2064
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
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   8988
      X2              =   9708
      Y1              =   7944
      Y2              =   7944
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   8868
      X2              =   9828
      Y1              =   2064
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
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   8868
      X2              =   8868
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   9828
      X2              =   9828
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase, Edit Decal Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3540
      TabIndex        =   8
      Top             =   1104
      Width           =   5148
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2388
      Top             =   1824
      Width           =   972
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   8868
      Top             =   1824
      Width           =   972
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1788
      Top             =   624
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2508
      Top             =   2064
      Width           =   732
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8988
      Top             =   2064
      Width           =   732
   End
End
Attribute VB_Name = "frmDCPurchaseMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim DefPayDate As String
Public BoxFileName As String, InvalidDate As Boolean
Public Sub setstuff(dt As String)
DefPayDate = dt
End Sub

Private Sub cmdDefExpire_Click()
  Load frmExpirationDefault
  DoEvents
  frmExpirationDefault.Show
End Sub

Private Sub cmdDelPay_Click()
    Load frmPaymentDelete
    DoEvents
    frmPaymentDelete.Show
End Sub

Private Sub cmdExitMenu_Click()
  Load frmDCMainMenu
  DoEvents
  frmDCMainMenu.Show
  Unload Me
End Sub
Private Sub cmdPurchase_Click()
  Dim FntSize As Integer, RecpPort As String
  Dim RP As Integer, lenRP As Integer
  Dim RcptPrnFile As ReceiptPRNType
  'If Not OPERNUM = 98 Then
  frmInfo.Label1 = "Verifying Receipt Printer..."
  frmInfo.Show
  DoEvents
    If Not Exist(RcptFileName$) Then
      Unload frmInfo
      ReDim MsgText(0 To 5) As String
      FntSize = frmMsgDialog.Label(1).FontSize
      frmMsgDialog.Label(1).FontSize = (FntSize + 2)
      MsgText(0) = "WARNING:"
      MsgText(1) = ""
      MsgText(2) = "RECEIPT SETUP FILE NOT FOUND!"
      MsgText(3) = "If you continue receipt printing"
      MsgText(4) = "will be disabled."
      MsgText(5) = "Receipt setup option is on CitiPak Main Menu."
      If GetOKorNot(MsgText()) Then
        DCLog "USER WANTS TO CONTINUE!"
      Else
        DCLog "USER ABORTED."
        Exit Sub
      End If
    Else
      RP = FreeFile
      lenRP = Len(RcptPrnFile)
      Open RcptFileName$ For Random Shared As RP Len = lenRP
      Get RP, 1, RcptPrnFile
      RecpPort = QPTrim(RcptPrnFile.RcpPort)
      Close
      If RcptPrnFile.PrnDefYN = 1 Then
        On Error GoTo noprnfound
       ' Printer.NewPage
        Open RecpPort For Output As RP
        Close RP
       End If
    End If
    frmPayPurchaseEntry.Wheretogo frmDCPurchaseMenu, frmDCPurchaseMenu, , DefPayDate
    DoEvents
    frmPayPurchaseEntry.Show
    Unload frmInfo
    'Unload frmDCPurchaseMenu
 'End If
Exit Sub
noprnfound:
        Unload frmInfo
        ReDim MsgText(0 To 5) As String
        FntSize = frmMsgDialog.Label(1).FontSize
        frmMsgDialog.Label(1).FontSize = (FntSize + 2)
        MsgText(0) = "WARNING:"
        MsgText(1) = ""
        MsgText(2) = "RECEIPT PRINTER NOT FOUND!"
        MsgText(3) = "If you continue receipt printing"
        MsgText(4) = "will be disabled."
        MsgText(5) = "Receipt setup option is on CitiPak Main Menu."
        If GetOKorNot(MsgText()) Then
          DCLog "USER WANTS TO CONTINUE!"
          frmPayPurchaseEntry.Wheretogo frmDCPurchaseMenu, frmDCPurchaseMenu, , DefPayDate
          DoEvents
          frmPayPurchaseEntry.Show
        Else
          DCLog "USER ABORTED."
          Exit Sub
        End If
End Sub

Private Sub cmdPostPayments_Click()
  Dim FntSize As Integer, PayBillName As String
  'Dim OPERNUM As Integer
  On Error GoTo ERRORSTUFF
  PayBillName$ = DCPath$ + "DCPAY" + QPTrim$(Str$(OperNum)) + ".DAT"

  If FileSize&(PayBillName$) <= 0 Then
  ReDim MsgText(0 To 5) As String

     frmMsgDialog.RetLabel = "-2"
     FntSize = frmMsgDialog.Label(2).FontSize

     frmMsgDialog.Label(2).FontSize = (FntSize + 2)
     MsgText(0) = "ERROR:"
     MsgText(1) = ""
     MsgText(2) = ""
     MsgText(3) = "NO UNPOSTED PAYMENT TRANSACTIONS!"
     MsgText(4) = ""
     MsgText(5) = ""
     GetOKorNot MsgText(), True
    GoTo Exitthis
  End If
  DCLog " IN: DC POST PAYMENTS,  OPER:" + Str$(OperNum)
  ChkTranDate
  If InvalidDate = True Then
    DCLog "Invalid Date found dcpaypost, give opt to cancel- OPER:" + Str$(OperNum)
    ReDim MsgText(0 To 5) As String
    FntSize = frmMsgDialog.Label(1).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    frmMsgDialog.Label(4).FontSize = (FntSize + 2)
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    MsgText(0) = "WARNING:"
    MsgText(1) = ""
    MsgText(2) = "Date of one or more payments"
    MsgText(3) = "is NOT within monthly date range."
    MsgText(4) = ""
    MsgText(5) = "OK to continue, or Cancel."
    If GetOKorNot(MsgText()) Then
      DCLog "Continue pay post with out of range dates."
    Else
      DCLog "Cancel pay post so can check dates."
      GoTo Exitthis
    End If
  End If

  DoItFlag = False
    frmNoOperatorsWarning.Label(5).Caption = "Post Payment Transactions"
    Load frmNoOperatorsWarning
    frmNoOperatorsWarning.Show vbModal
    If Not DoItFlag Then
      GoTo Exitthis
    End If
  DeActivateControls Me
  PostPayments
  ActivateControls Me
  MsgBox "Posting Payments Completed.", vbOKOnly, "Procedure Complete"
Exitthis:
Exit Sub
ERRORSTUFF:
  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "CMPayUtilEntry", "cmdSave", Erl)
    Case emrExitProc:
      Resume Proc_Exit
    Case emrResume:
      Resume
    Case emrResumeNext:
      Resume Next
    Case Else
      '--- Technically, this should never happen.
      Resume Proc_Exit
  End Select
Proc_Exit:
  '--- Cleanup code goes here...
    Close
   'Unload Me
    ActivateControls Me
End Sub

Private Sub cmdPrintJournalE_Click()
Dim FntSize As Integer, PayFileName As String
  Dim Tot As Integer
  Tot = 0

  PayFileName$ = DCPath$ + "DCPAY" + QPTrim$(Str$(OperNum)) + ".DAT"
  If Exist(PayFileName$) Then Tot = Tot + 1
  ReDim MsgText(0 To 5) As String
   If Tot < 1 Then
     frmMsgDialog.RetLabel = "-2"
     FntSize = frmMsgDialog.Label(2).FontSize

     frmMsgDialog.Label(2).FontSize = (FntSize + 2)
     MsgText(0) = "ERROR:"
     MsgText(1) = ""
     MsgText(2) = ""
     MsgText(3) = "NO UNPOSTED PAYMENT TRANSACTIONS!"
     MsgText(4) = ""
     MsgText(5) = ""
     GetOKorNot MsgText(), True
  Else
    frmReportOpt.Show 1
    DeActivateControls Me
    If rptopt > 0 Then
     PrintEditList rptopt, 0
    Else
      ActivateControls Me
    End If

  End If
 ' PrintEditList 2, 0
End Sub

Private Sub cmdPrintJournalN_Click()
Dim FntSize As Integer, PayFileName As String
  Dim Tot As Integer
  Tot = 0

  PayFileName$ = DCPath$ + "DCPAY" + QPTrim$(Str$(OperNum)) + ".DAT"
  If Exist(PayFileName$) Then Tot = Tot + 1
  ReDim MsgText(0 To 5) As String
   If Tot < 1 Then
     frmMsgDialog.RetLabel = "-2"
     FntSize = frmMsgDialog.Label(2).FontSize

     frmMsgDialog.Label(2).FontSize = (FntSize + 2)
     MsgText(0) = "ERROR:"
     MsgText(1) = ""
     MsgText(2) = ""
     MsgText(3) = "NO UNPOSTED PAYMENT TRANSACTIONS!"
     MsgText(4) = ""
     MsgText(5) = ""
     GetOKorNot MsgText(), True
  Else
    frmReportOpt.Show 1
    DeActivateControls Me
    If rptopt > 0 Then
     PrintEditList rptopt, 1
    Else
      ActivateControls Me
    End If

  End If
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  'screenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpPurchaseDecals
  Refresh
  DoEvents
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitMenu.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        DCLog "Closed via PaymentMenu by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ' Me.Visible = False
    Temp_Class.ResizeControls Me
   ' Me.Visible = True
   ' Me.SetFocus
  End If
  DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExitMenu_Click
      KeyCode = 0
    Case vbKeyHome
      cmdPurchase.SetFocus
    Case vbKeyEnd
      cmdExitMenu.SetFocus
    Case Else:
  End Select
End Sub

Private Sub PostPayments()
  Dim PayFileName As String, NumOfDCRecs As Long, DCFile As Integer
  Dim DCEditRecLen As Integer, DCEdFile As Integer, DCTransRecLen As Integer
  Dim DCTransFile As Integer, NumOfTransRecs As Long, NextTransRec As Long
  Dim cnt As Long, Prev As Long, DCVehReclen As Integer, DCvFile As Integer
  Dim NumOfVRecs As Long, VehRecord As Long
  PayFileName$ = DCPath$ + "DCPAY" + QPTrim$(Str$(OperNum)) + ".DAT"
  ReDim EDitPaymentRec(1) As DCEditPaymentRecType
  ReDim DCCustRec(1) As DCCustRecType
  OpenDCCustFile NumOfDCRecs, DCFile

  DCEditRecLen = Len(EDitPaymentRec(1))
  DCEdFile = FreeFile
  Open PayFileName$ For Random Access Read Write Shared As DCEdFile Len = DCEditRecLen
  NumOfDCRecs = LOF(DCEdFile) \ DCEditRecLen
  
  ReDim DCVRec(1) As DCVehType
  
  ReDim DCTransRec(1) As DCTransRecType
  DCTransRecLen = Len(DCTransRec(1))
  DCTransFile = FreeFile
  Open "DCTrans.DAT" For Random Access Read Write Shared As DCTransFile Len = DCTransRecLen
  NumOfTransRecs = LOF(DCTransFile) \ DCTransRecLen
  NextTransRec = NumOfTransRecs + 1

  Do
    cnt = cnt + 1
    Get DCEdFile, cnt, EDitPaymentRec(1)

    If EDitPaymentRec(1).Amount >= 0 And Val(EDitPaymentRec(1).CustNumber) > 0 Then

      'GoSub OldVehPost
      GoSub UpdateVehRecord
      ''If EDitPaymentRec(1).NewVeh = "Y" Then GoSub UpdateVendorPointer
      
      Get DCFile, Val(EDitPaymentRec(1).CustNumber), DCCustRec(1)
      ' Post Charge First to Offset Payment of Decal
      DCTransRec(1).CustomerNumber = EDitPaymentRec(1).CustNumber
      DCTransRec(1).TransDate = EDitPaymentRec(1).TranDate
      DCTransRec(1).TransAmount = EDitPaymentRec(1).Amount
      DCTransRec(1).TransType = 1               ' Type 1 = Charge
      DCTransRec(1).TRVinDesc = EDitPaymentRec(1).VinDesc
      DCTransRec(1).TransTender = EDitPaymentRec(1).TransTender
      DCTransRec(1).CashAmount = EDitPaymentRec(1).CashAmt
      DCTransRec(1).ChkAmount = EDitPaymentRec(1).CheckAmt
      DCTransRec(1).BalanceAfterTrans = DCCustRec(1).AcctBal + EDitPaymentRec(1).Amount
      DCTransRec(1).makemodel = EDitPaymentRec(1).makemodel
      DCTransRec(1).StateTag = EDitPaymentRec(1).StateTag
      DCTransRec(1).Sticker = EDitPaymentRec(1).Sticker
      DCTransRec(1).ExpireDate = EDitPaymentRec(1).ExpDate
      DCTransRec(1).OperNum = EDitPaymentRec(1).OperNum
      If Len(QPTrim$(EDitPaymentRec(1).PayDesc)) > 0 Then
        DCTransRec(1).ExtraDesc = "DC-" + EDitPaymentRec(1).PayDesc
      Else
        DCTransRec(1).ExtraDesc = "DC-Purchase"
      End If
      DCTransRec(1).ExtraRoom = ""
      DCTransRec(1).NextTrans = 0
      DCTransRec(1).GLInterfaced = "Y"
      DCTransRec(1).DecalCat = EDitPaymentRec(1).DecalCat       'Dale Need This in his stuff
      DCTransRec(1).VehRecord = EDitPaymentRec(1).VehRecord
      DCTransRec(1).ChkByte = Chr$(1)
      DCTransRec(1).VoidFlag = "N"
      Put DCTransFile, NextTransRec, DCTransRec(1)

      Get DCFile, Val(EDitPaymentRec(1).CustNumber), DCCustRec(1)
      DCCustRec(1).AcctBal = DCCustRec(1).AcctBal + EDitPaymentRec(1).Amount
      Put DCFile, Val(EDitPaymentRec(1).CustNumber), DCCustRec(1)
      If DCCustRec(1).FirstTrans = 0 Then
        DCCustRec(1).FirstTrans = NextTransRec
        DCCustRec(1).LastTrans = NextTransRec
        Put DCFile, Val(EDitPaymentRec(1).CustNumber), DCCustRec(1)
      Else
        Prev = DCCustRec(1).LastTrans
        DCCustRec(1).LastTrans = NextTransRec
        Put DCFile, Val(EDitPaymentRec(1).CustNumber), DCCustRec(1)
        Get DCTransFile, Prev, DCTransRec(1)
        DCTransRec(1).NextTrans = NextTransRec
        Put DCTransFile, Prev, DCTransRec(1)
      End If
      NextTransRec = NextTransRec + 1

      ' Post Transaction Record First
      DCTransRec(1).CustomerNumber = EDitPaymentRec(1).CustNumber
      DCTransRec(1).TransDate = EDitPaymentRec(1).TranDate
      DCTransRec(1).TransAmount = EDitPaymentRec(1).Amount
      DCTransRec(1).TransType = 2               ' Type 2 = Payment
      DCTransRec(1).TRVinDesc = EDitPaymentRec(1).VinDesc
      DCTransRec(1).TransTender = EDitPaymentRec(1).TransTender
      DCTransRec(1).CashAmount = EDitPaymentRec(1).CashAmt
      DCTransRec(1).ChkAmount = EDitPaymentRec(1).CheckAmt
      DCTransRec(1).BalanceAfterTrans = DCCustRec(1).AcctBal - EDitPaymentRec(1).Amount
      DCTransRec(1).OperNum = EDitPaymentRec(1).OperNum
      If Len(QPTrim$(EDitPaymentRec(1).PayDesc)) > 0 Then
        DCTransRec(1).ExtraDesc = "DC-" + EDitPaymentRec(1).PayDesc
      Else
        DCTransRec(1).ExtraDesc = "DC-Payment"
      End If
      DCTransRec(1).ExtraRoom = ""
      DCTransRec(1).NextTrans = 0
      DCTransRec(1).GLInterfaced = "N"
      DCTransRec(1).DecalCat = EDitPaymentRec(1).DecalCat
      DCTransRec(1).ChkByte = Chr$(1)
      DCTransRec(1).VoidFlag = "N"
      DCTransRec(1).VehRecord = EDitPaymentRec(1).VehRecord
      Put DCTransFile, NextTransRec, DCTransRec(1)

      Get DCFile, Val(EDitPaymentRec(1).CustNumber), DCCustRec(1)
      DCCustRec(1).AcctBal = DCCustRec(1).AcctBal - EDitPaymentRec(1).Amount
      DCCustRec(1).LICENSE = EDitPaymentRec(1).Sticker
      Put DCFile, Val(EDitPaymentRec(1).CustNumber), DCCustRec(1)

      If DCCustRec(1).FirstTrans = 0 Then
        DCCustRec(1).FirstTrans = NextTransRec
        DCCustRec(1).LastTrans = NextTransRec
        Put DCFile, Val(EDitPaymentRec(1).CustNumber), DCCustRec(1)
      Else
        Prev = DCCustRec(1).LastTrans
        DCCustRec(1).LastTrans = NextTransRec
        Put DCFile, Val(EDitPaymentRec(1).CustNumber), DCCustRec(1)
        Get DCTransFile, Prev, DCTransRec(1)
        DCTransRec(1).NextTrans = NextTransRec
        Put DCTransFile, Prev, DCTransRec(1)
      End If
      NextTransRec = NextTransRec + 1
    End If

  Loop Until cnt > NumOfDCRecs
  Close
  Kill PayFileName$
  ' Show All Posted
  DCLog "POSTED:" + Str$(NumOfDCRecs)
 ' MsgBox "Posting Complete", vbOKOnly, "Complete"
  Close
  Exit Sub
'OldVehPost:
'  If EditPaymentRec(1).Owner = "Y" Then Return
'  ReDim DCOVRec(1) As DCOldVehType
'  DCOVreclen = Len(DCOVRec(1))
'  DCOVFile = FreeFile
'  Open "DCOLDVEH.DAT" For Random Access Read Write Shared As DCOVFile Len = DCOVreclen
'  NumOfOVRecs! = LOF(DCOVFile) \ DCOVreclen
'  NextOVRec! = NumOfOVRecs! + 1
'  DCOVRec(1).Make = LTrim$(EditPaymentRec(1).OldMake)
'  DCOVRec(1).year = LTrim$(EditPaymentRec(1).OldDesc)
'  DCOVRec(1).CustRec = Val(EditPaymentRec(1).CustNumber)
'  'DCOVRec(1).
'  DCOVRec(1).MoreRoom = ""
'  Put DCOVFile, NextOVRec!, DCOVRec(1)
'  Close DCOVFile
'Return

UpdateVehRecord:
  DCVehReclen = Len(DCVRec(1))
  DCvFile = FreeFile
  Open "DCVEH.DAT" For Random Access Read Write Shared As DCvFile Len = DCVehReclen
  NumOfVRecs = LOF(DCvFile) \ DCVehReclen
  VehRecord = EDitPaymentRec(1).VehRecord
  If VehRecord < 0 Or VehRecord > NumOfVRecs Then Close DCvFile: Return
  Get DCvFile, VehRecord, DCVRec(1)
  DCVRec(1).ExpireDate = EDitPaymentRec(1).ExpDate
  DCVRec(1).Sticker = LTrim$(EDitPaymentRec(1).Sticker)
  DCVRec(1).Valid = "Y"
  DCVRec(1).Fee = EDitPaymentRec(1).Amount
  DCVRec(1).DecalCat = EDitPaymentRec(1).DecalCat
  DCVRec(1).makemodel = EDitPaymentRec(1).makemodel
  DCVRec(1).StateTag = EDitPaymentRec(1).StateTag
  DCVRec(1).Active = "Y"
  DCVRec(1).Desc = EDitPaymentRec(1).VinDesc
  DCVRec(1).Notes = EDitPaymentRec(1).Notes
  DCVRec(1).MoreRoom = ""
  DCVRec(1).PBFlag = EDitPaymentRec(1).PersBuss
  DCVRec(1).MasterRecord = Val(EDitPaymentRec(1).CustNumber)
  Put DCvFile, VehRecord, DCVRec(1)
  Close DCvFile
Return

'UpdateVendorPointer:
'  ReDim DCCustREc(1) As DCCustRecType
'  If fpCustRecNo > 0 Then
'    OpenDCCustFile NumOfDCRecs, DCFile
'    Get DCFile, fpCustRecNo, DCCustREc(1)
'    If DCCustREc(1).FirstCar = 0 Then
'      DCCustREc(1).FirstCar = VehRecord
'      DCCustREc(1).LastCar = VehRecord
'      Put DCFile, fpCustRecNo, DCCustREc(1)
'    Else
'      PrevRec = DCCustREc(1).LastCar
'      DCCustREc(1).LastCar = VehRecord
'      Put DCFile, fpCustRecNo, DCCustREc(1)
'
'      Get DCvFile, PrevRec, DCVRec(1)
'      DCVRec(1).NextRec = VehRecord
'      Put DCvFile, PrevRec, DCVRec(1)
'    End If
'    Close DCvFile
'  End If
'
'Return
 

End Sub
'' OPERNUM , PostDate$
'  Dim PayBillName As String, PayDepoName As String
'  Dim UBCustRecLen As Integer, UBPayRecLen As Integer, UBTransRecLen As Integer
'  Dim TranFile As Integer, CHandle As Integer, thandle As Integer
'  Dim PHandle As Integer, NumPayRecs As Long, cnt As Long
'  Dim RevAmts As Integer, NextTransRec As Long, TotalCustBalance As Double
'  Dim CustChCnt As Integer
'
'  PayBillName$ = UBPath$ + "UBPAY" + QPTrim$(Str$(OperNum)) + ".DAT"
'  PayDepoName$ = UBPath$ + "UBDEP" + QPTrim$(Str$(OperNum)) + ".DAT"
'  FrmShowPctComp.Label1 = "Posting Deposit Transactions"
'  FrmShowPctComp.Show
'
'  UBLog "POSTING TRANSACTIONS START"
'
'  ReDim UBTransRec(1) As UBTransRecType
'  ReDim TUBTransRec(1) As UBTransRecType
'  ReDim UBCustRec(1) As NewUBCustRecType
'  ReDim UBPaymentRec(1) As UBPaymentRecType
'
'  UBCustRecLen = Len(UBCustRec(1))
'  UBPayRecLen = Len(UBPaymentRec(1))
'  UBTransRecLen = Len(UBTransRec(1))
'
'  TranFile = FreeFile
'  Open UBPath$ + "UBTRANS.DAT" For Random Shared As TranFile Len = UBTransRecLen
'  Close TranFile
'
'
'  UBLog "POSTING: DEPOSITS"
'  If FileSize&(PayDepoName$) > 0 Then
'    CHandle = FreeFile
'    Open UBPath$ + "UBCUST.DAT" For Random Shared As CHandle Len = UBCustRecLen
'    thandle = FreeFile
'    Open UBPath$ + "UBTRANS.DAT" For Random Shared As thandle Len = UBTransRecLen
'    PHandle = FreeFile
'    Open PayDepoName$ For Random Shared As PHandle Len = UBPayRecLen
'
'    NumPayRecs& = LOF(PHandle) \ UBPayRecLen
'
'    'ShowProcessingScrn "Posting Deposit Transactions"
'    For cnt& = 1 To NumPayRecs&
'      FrmShowPctComp.ShowPctComp cnt&, NumPayRecs&
'      If FrmShowPctComp.Out = True Then
'        Close
'        FrmShowPctComp.Out = False
'        Close
'        Exit Sub
'      End If
'      LSet UBTransRec(1) = TUBTransRec(1)
'      Get PHandle, cnt&, UBPaymentRec(1) ',  UBPayRecLen
'      Get CHandle, UBPaymentRec(1).CustAcct, UBCustRec(1) ',  UBCustRecLen
'      UBTransRec(1).TransDate = UBPaymentRec(1).PAYDATE
'      UBTransRec(1).TransType = TranDepositPayment
'
'      '022098 Added
'      If Len(QPTrim$(UBPaymentRec(1).Desc)) = 0 Then
'        UBTransRec(1).TransDesc = "DEPOSIT PAYMENT"
'      Else
'        UBTransRec(1).TransDesc = UBPaymentRec(1).Desc
'        UBTransRec(1).BillMsg = "DEPOSIT PAYMENT"
'      End If
'      '^^This holds the Payment Description
'
'      'UBTransRec(1)CustLocation = UBPaymentRec(1).CUSTACCT
'      UBTransRec(1).OperatorNumber = OperNum
'      UBTransRec(1).CustAcctNo = UBPaymentRec(1).CustAcct
'      UBTransRec(1).CustStatus = UBCustRec(1).Status
'      UBTransRec(1).Transamt = UBPaymentRec(1).AMTPAID
'      UBTransRec(1).CheckAmount = UBPaymentRec(1).CHKAMT
'      UBTransRec(1).CashAmount = UBPaymentRec(1).CashAmt
'
'      If UBTransRec(1).CheckAmount > 0 And UBTransRec(1).CashAmount > 0 Then
'        UBTransRec(1).PayTypeCode = 3
'      ElseIf UBTransRec(1).CashAmount > 0 Then
'        UBTransRec(1).PayTypeCode = 1
'      ElseIf UBTransRec(1).CheckAmount > 0 Then
'        If QPTrim(UBPaymentRec(1).TENDERTY) = "Charge" Then
'          UBTransRec(1).PayTypeCode = 4
'        Else
'          UBTransRec(1).PayTypeCode = 2
'        End If
'      End If
'
'      For RevAmts = 1 To MaxRevsCnt
'        UBTransRec(1).RevAmt(RevAmts) = UBPaymentRec(1).PaidOwed(RevAmts).AMTPD1
'      Next
'   '05-05-97 FIX added run balance to deposit trans
'   UBTransRec(1).RunBalance = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
'
'   UBTransRec(1).PrevTrans = UBCustRec(1).LastTrans
'   NextTransRec& = (LOF(thandle) \ UBTransRecLen) + 1
'
'   If NextTransRec& <= 0 Then
'     NextTransRec& = 1
'   End If
'
'   Put thandle, NextTransRec&, UBTransRec(1) ',  UBTransRecLen
'
'   'UBCustRec(1).DepositAmt = UBTransRec(1).TransAmt
'
'   '04-14-98 Testing
'   UBCustRec(1).DepositAmt = Round#(UBCustRec(1).DepositAmt + UBTransRec(1).Transamt)
'
'   UBCustRec(1).LastTrans = NextTransRec&
'   Put CHandle, UBPaymentRec(1).CustAcct, UBCustRec(1) ',  UBCustRecLen
'   'ShowPctComp cnt&, NumPayRecs&
' Next
'    Close CHandle
'    Close thandle
'    Close PHandle
'
'    KillFile PayDepoName$
'  End If
'  UBLog "POSTED:" + Str$(NumPayRecs&)
'  '**********
'  UBLog "POSTING: PAYMENTS"
'  If FileSize&(PayBillName$) > 0 Then
'    CHandle = FreeFile
'    Open UBPath$ + "UBCUST.DAT" For Random Shared As CHandle Len = UBCustRecLen
'    thandle = FreeFile
'    Open UBPath$ + "UBTRANS.DAT" For Random Shared As thandle Len = UBTransRecLen
'    PHandle = FreeFile
'    Open PayBillName$ For Random Shared As PHandle Len = UBPayRecLen
'
'    NumPayRecs& = LOF(PHandle) \ UBPayRecLen
'    FrmShowPctComp.Label1 = "Posting Payment Transactions"
'    FrmShowPctComp.Show
'
'    'ShowProcessingScrn "Posting Payment Transactions"
'    For cnt& = 1 To NumPayRecs&
'      FrmShowPctComp.ShowPctComp cnt&, NumPayRecs&
'      If FrmShowPctComp.Out = True Then
'        Close
'        FrmShowPctComp.Out = False
'        Close
'        Exit Sub
'      End If
'
'      LSet UBTransRec(1) = TUBTransRec(1)
'      Get PHandle, cnt&, UBPaymentRec(1) ',  UBPayRecLen
'      Get CHandle, UBPaymentRec(1).CustAcct, UBCustRec(1) ',  UBCustRecLen
'      UBTransRec(1).TransDate = UBPaymentRec(1).PAYDATE
'      UBTransRec(1).TransType = TranBillPayment
'
'      '052698 Added tax exempt flag to trans rec. For payment summary report
'      UBTransRec(1).TaxExempt = UBPaymentRec(1).TaxExempt
'
'      '022098 Added
'      If Len(QPTrim$(UBPaymentRec(1).Desc)) = 0 Then
'        UBTransRec(1).TransDesc = "BILLING-PAYMENT"
'            Else
'        UBTransRec(1).TransDesc = UBPaymentRec(1).Desc
'        UBTransRec(1).BillMsg = "BILLING-PAYMENT"
'      End If
'      '^^This holds the Payment Description
'      'UBTransRec(1)CustLocation = UBPaymentRec(1).CUSTACCT
'      UBTransRec(1).OperatorNumber = OperNum
'      UBTransRec(1).CustAcctNo = UBPaymentRec(1).CustAcct
'      UBTransRec(1).CustStatus = UBCustRec(1).Status
'      UBTransRec(1).Transamt = UBPaymentRec(1).AMTPAID
'      UBTransRec(1).CheckAmount = UBPaymentRec(1).CHKAMT
'      UBTransRec(1).CashAmount = UBPaymentRec(1).CashAmt
'
'      If UBTransRec(1).CheckAmount > 0 And UBTransRec(1).CashAmount > 0 Then
'        UBTransRec(1).PayTypeCode = 3
'      ElseIf UBTransRec(1).CashAmount > 0 Then
'        UBTransRec(1).PayTypeCode = 1
'      ElseIf UBTransRec(1).CheckAmount > 0 Then
'        If QPTrim(UBPaymentRec(1).TENDERTY) = "Charge" Then
'          UBTransRec(1).PayTypeCode = 4
'        Else
'          UBTransRec(1).PayTypeCode = 2
'        End If
'      End If
'
'      'IF UBCustRec(1).PrevBalance > 0 THEN
'      '050597 changed to zero if <> zero
'      If UBCustRec(1).PrevBalance <> 0 Then
'        If UBTransRec(1).Transamt >= UBCustRec(1).PrevBalance Then
'          UBCustRec(1).PrevBalance = 0
'        ElseIf UBTransRec(1).Transamt < UBCustRec(1).PrevBalance Then
'          UBCustRec(1).PrevBalance = Round#(UBCustRec(1).PrevBalance - UBTransRec(1).Transamt)
'        End If
'      End If
'
'      For RevAmts = 1 To MaxRevsCnt
'        UBTransRec(1).RevAmt(RevAmts) = UBPaymentRec(1).PaidOwed(RevAmts).AMTPD1
'        UBCustRec(1).CurrRevAmts(RevAmts) = Round#(UBCustRec(1).CurrRevAmts(RevAmts) - UBTransRec(1).RevAmt(RevAmts))
'        'This is for previous bill distribution
'        'UBCustRec(1).PrevRevAmts(RevAmts) = Round#(UBCustRec(1).PrevRevAmts(RevAmts) - UBTransRec(1).RevAmt(RevAmts))
'        'IF UBCustRec(1).PrevRevAmts(RevAmts) < 0 THEN
'        '  UBCustRec(1).PrevRevAmts(RevAmts) = 0
'        'END IF
'      Next
'      TotalCustBalance# = 0
'      For RevAmts = 1 To MaxRevsCnt
'        TotalCustBalance# = Round#(TotalCustBalance# + UBCustRec(1).CurrRevAmts(RevAmts))
'      Next
'      UBCustRec(1).CurrBalance = Round#(TotalCustBalance# - UBCustRec(1).PrevBalance)
'      '02-26-97 Was not adding prev bal
'      UBTransRec(1).RunBalance = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
'      UBTransRec(1).PrevTrans = UBCustRec(1).LastTrans
'      UBTransRec(1).VoidFlag = False
'      UBTransRec(1).FromCMFlag = False
'      NextTransRec& = (LOF(thandle) \ UBTransRecLen) + 1
'      If NextTransRec& <= 0 Then
'        NextTransRec& = 1
'      End If
'      Put thandle, NextTransRec&, UBTransRec(1) ',  UBTransRecLen
'      UBCustRec(1).LastTrans = NextTransRec&
'      If Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) = 0 Then
'        If UBCustRec(1).Status = "B" Then
'          UBCustRec(1).Status = "I"
'          CustChCnt = CustChCnt + 1
'          UBLog "POSTING: SET CUST STATUS to I. Acct:" + Str$(UBPaymentRec(1).CustAcct)
'        End If
'      End If
'      Put CHandle, UBPaymentRec(1).CustAcct, UBCustRec(1) ',  UBCustRecLen
'      'ShowPctComp cnt&, NumPayRecs&
'    Next
'    Close CHandle
'    Close thandle
'    Close PHandle
'    KillFile PayBillName$
'  End If
'  UBLog "POSTED:" + Str$(NumPayRecs&)
'  If CustChCnt > 0 Then
'    UBLog "POSTING: CUST STATUS CHANGED:" + Str$(CustChCnt)
'  End If
'  UBLog "POSTING TRANSACTIONS FINISH"
'
''  BlockClear
''  DisplayUBScrn "UPDATEOK"
''  WaitForAction
'
'  Erase UBTransRec, TUBTransRec, UBCustRec, UBPaymentRec
'
'ExitPayPost:
'  UBLog "OUT: UB POST PAYMENTS,  OPER:" + Str$(OperNum)
Private Sub ChkTranDate()
  Dim PayBillName As String, Today As String
  Dim DCPayRecLen As Integer
  Dim CHandle As Integer, THandle As Integer, chkthedate As Integer
  Dim PHandle As Integer, NumPayRecs As Long, cnt As Long
  InvalidDate = False
  PayBillName$ = DCPath$ + "DCPAY" + QPTrim$(Str$(OperNum)) + ".DAT"
  DCLog "Check Payment Date BP"
  FrmShowPctComp.Label1 = "Checking Transaction Dates"
  FrmShowPctComp.Show
  Today = Format(Now, "mm/dd/yyyy")
  Dim DCPaymentRec As DCEditPaymentRecType
  chkthedate = Date2Num(Today)
  DCPayRecLen = Len(DCPaymentRec)
  If FileSize&(PayBillName$) > 0 Then
    PHandle = FreeFile
    Open PayBillName$ For Random Shared As PHandle Len = DCPayRecLen
    NumPayRecs& = LOF(PHandle) \ DCPayRecLen
    For cnt& = 1 To NumPayRecs&
      FrmShowPctComp.ShowPctComp cnt&, NumPayRecs&
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        Close
        Exit Sub
      End If
      Get PHandle, cnt&, DCPaymentRec ',  UBPayRecLen
      If DCPaymentRec.TranDate > (chkthedate + 30) Or DCPaymentRec.TranDate < (chkthedate - 30) Then
        InvalidDate = True
        Unload FrmShowPctComp
        Close
        Exit Sub
      End If
    Next
    Close
  End If
End Sub
Private Sub PrintEditList(rptopt As Integer, Order As Integer) '(OPERNUM, PostDate$)
  Dim cnt As Long, x As Integer, Dash1 As String, CatCnt As Long
  Dim IndexRecLen As Integer, Lp As Long
  Dim Operator As String, Page As Integer, NumOfDCRecs As Long
  Dim PayFileName As String, PayJourName As String, CatLoop As Long
  Dim Header As String, CMOperRecLen As Integer, DCPayRecLen As Integer
  Dim CMFile As Integer, NumRecs As Long, PayOKFlag As Boolean
  Dim TotalRecs As Long, TotalValue As Double, NumOfRecs As Long
  Dim RptHandle As Integer, PHandle As Integer, TrNumRecs  As Long
  Dim DoneCnt As Integer, Pmnt As String, TrHandle As Integer
  Dim TotalCash As Double, TotalCheck As Double, TotalCust As Long
  Dim TotalAmount As Double, TotalChange As Double, TotalReceipts As Integer
  Dim RCnt As Integer, Diff As Double, ReportFile As String
  Dim GTotal As Double, TTax As Double, PostDate As String, ToPrint As String
  Dim Graph As Boolean, ReportSum1 As String, ReportSum2 As String
  Dim SumRpt1 As Integer, SumRpt2 As Integer, TotalChrge As Double
  Dim tmp As DistArrayType, SumPrnt As String, TotalChks As Integer
  Dim lngCurLow As Long, lngCurHigh As Long, AcctNo As Long, CHandle As Integer
  Dim IndexName As String, IHandle As Integer, dcnt As Long, CRec As Long
  Dim DHandle As Integer, Oper As String, DCFile As Integer
  ReDim DCPaymentRec(1) As DCEditPaymentRecType
  Operator$ = QPTrim$(Str$(OperNum))
  Oper$ = Operator$
  PayFileName$ = DCPath$ + "DCPAY" + Oper$ + ".DAT"
  DCPayRecLen = Len(DCPaymentRec(1))
  

  Dim Cat$(250), CatAmt#(250)   'Set Maximum Catagories at 250
  ReportFile$ = "DCPAYED.PRN"   'Report File Name
  FF$ = Chr$(12)
  MaxLines = 53
  Linecnt = 0

  ToPrint$ = ""
  If rptopt = 1 Then
    Graph = True
  Else
    Graph = False
  End If
  FrmShowPctComp.Label1 = "Creating Decal Payment Edit Report"
  FrmShowPctComp.Show

  PostDate$ = Format(Now, "mm/dd/yyyy")
  If Order = 1 Then
    CHandle = FreeFile
    Open PayFileName$ For Random Shared As CHandle Len = DCPayRecLen
    NumOfRecs& = LOF(CHandle) \ DCPayRecLen

    ReDim ServIndex(1 To NumOfRecs) As DCCustIDXRecType
    IndexRecLen = Len(ServIndex(1))
    For cnt& = 1 To NumOfRecs&
      Get CHandle, cnt&, DCPaymentRec(1)
      ServIndex(cnt).IDXName = QPStripLast$(DCPaymentRec(1).CustName)
      ServIndex(cnt).IDXRECORD = cnt
    Next
    Close CHandle

    lngCurLow = LBound(ServIndex)
    lngCurHigh = UBound(ServIndex)
    NameQSort ServIndex(), lngCurLow, lngCurHigh
    IndexName$ = "DCPTemp.IDX"
    'KillFile IndexName$
    IHandle = FreeFile
    Open IndexName$ For Random Shared As IHandle Len = 4
    For cnt = 1 To lngCurHigh
      CRec& = ServIndex(cnt).IDXRECORD
      Put IHandle, cnt, CRec&
    Next
    Close IHandle
  End If
 
  DCFile = FreeFile
  Open PayFileName$ For Random Access Read Write Shared As DCFile Len = DCPayRecLen
  NumOfDCRecs = LOF(DCFile) \ DCPayRecLen

  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  GoSub PrintRptHeader

  For cnt = 1 To NumOfDCRecs
      FrmShowPctComp.ShowPctComp cnt&, NumOfDCRecs
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        Close
        Exit Sub
      End If

      If Order = 1 Then
        AcctNo& = ServIndex(cnt).IDXRECORD
      Else
        AcctNo& = cnt&
      End If
    Get DCFile, AcctNo&, DCPaymentRec(1)

      If Linecnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintRptHeader
      End If
      If DCPaymentRec(1).Amount >= 0 Then
        Print #RptHandle, QPTrim$(DCPaymentRec(1).CustNumber); Tab(10); Mid$(QPTrim$(DCPaymentRec(1).CustName), 1, 20);
       ' Tab(55); Using$("$###,###.##", DCPaymentRec(1).Amount);
        Print #RptHandle, Tab(32); Using("######.##", DCPaymentRec(1).CashAmt);
        If DCPaymentRec(1).TransTender = 4 Then
          Print #RptHandle, Tab(42); Using("######.##", 0);
          Print #RptHandle, Tab(52); Using("######.##", DCPaymentRec(1).CheckAmt);
          TotalChrge# = Round#(TotalChrge# + DCPaymentRec(1).CheckAmt)
        Else
          Print #RptHandle, Tab(42); Using("######.##", DCPaymentRec(1).CheckAmt);
          Print #RptHandle, Tab(52); Using("######.##", 0);
          TotalCheck# = Round#(TotalCheck# + DCPaymentRec(1).CheckAmt)
          If DCPaymentRec(1).CheckAmt > 0 Then
            TotalChks = TotalChks + 1
          End If
        End If
        
        TotalCash# = Round#(TotalCash# + DCPaymentRec(1).CashAmt)
        TotalAmount# = Round#(TotalAmount# + DCPaymentRec(1).Amount)
        TotalChange# = Round#(TotalChange# + DCPaymentRec(1).Change)
        TotalReceipts = TotalReceipts + 1

        Print #RptHandle, Tab(62); Using("######.##", Round#(Round#(DCPaymentRec(1).CheckAmt + DCPaymentRec(1).CashAmt) - DCPaymentRec(1).Change));
        Print #RptHandle, Tab(72); Using("######.##", DCPaymentRec(1).Change)
        
        Print #RptHandle, Tab(10); QPTrim$(DCPaymentRec(1).CustAddr); Tab(50); "    Category: "; DCPaymentRec(1).DecalCat
        Print #RptHandle, "     VIN/Desc: "; QPTrim$(DCPaymentRec(1).VinDesc); Tab(50); "    Sticker#: "; DCPaymentRec(1).Sticker
        Print #RptHandle, "   Make/Model: "; QPTrim$(DCPaymentRec(1).makemodel); Tab(50); "  State Tag#: "; RTrim$(DCPaymentRec(1).StateTag)
        Print #RptHandle, " Payment Date: "; Num2Date(DCPaymentRec(1).TranDate); Tab(50); " Expire Date: "; Num2Date(DCPaymentRec(1).ExpDate)
        Print #RptHandle, "     Resident: "; DCPaymentRec(1).resident; Tab(28); "Vehicle Owned: "; DCPaymentRec(1).Owner;
        Print #RptHandle, Tab(50); "Prs/Business: "; DCPaymentRec(1).PersBuss
        Print #RptHandle, String$(80, "-")
        TotalCust = TotalCust + 1
        TotalValue# = TotalValue# + DCPaymentRec(1).Amount
        TotalValue# = Int((TotalValue# * 100) + 0.5) / 100
        GoSub CatagoryTotal
        Linecnt = Linecnt + 10
      End If
   ' End If
    
  Next cnt
  GoSub PrintRptEnding
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  Close         'Close all open files now
  Header$ = "Payment Edit Listing"
If Not Graph Then
  ViewPrint ReportFile$, Header$
  'KILL ReportFile$
Else
  Load frmLoadingRpt
  ARptLineRpt.GetName ReportFile$
  ARptLineRpt.startrpt

End If
DoneHere1:
  ActivateControls Me
  Exit Sub

CatagoryTotal:
  If CatCnt = 0 Then
    CatCnt = 1
    Cat$(1) = LTrim$(DCPaymentRec(1).DecalCat)
    CatAmt#(1) = DCPaymentRec(1).Amount
    Return
  End If
  For CatLoop = 1 To CatCnt
    If Cat$(CatLoop) = LTrim$(DCPaymentRec(1).DecalCat) Then
      CatAmt#(CatLoop) = CatAmt#(CatLoop) + DCPaymentRec(1).Amount
      Return
    End If
  Next CatLoop
  CatCnt = CatCnt + 1
  Cat$(CatCnt) = LTrim$(DCPaymentRec(1).DecalCat)
  CatAmt#(CatCnt) = DCPaymentRec(1).Amount

  Return

PrintRptHeader:
  Page = Page + 1
  Print #RptHandle, ""
  Print #RptHandle, ""
  Print #RptHandle, Tab(20); "Vehicle Decals Payment EDIT Listing"
  Print #RptHandle, "Posting Date: "; PostDate$
  Print #RptHandle, "    Operator: "; PWUser; ","; Operator$; Tab(72); "Page #"; Page
  Print #RptHandle, ""
  Print #RptHandle, "Cust #"; Tab(10); "Billing Name"; Tab(35); "Cash"; Tab(45); "Check";
  Print #RptHandle, Tab(55); "Charge"; Tab(65); "Amount"; Tab(74); "Change"
  Print #RptHandle, Tab(65); "Applied"; Tab(74); "Given"
  Print #RptHandle, String$(80, "=")
  Linecnt = 5
  Return
'PrintRptHeader:
'  If Not Graph Then
'  Page = Page + 1
'  Print #RptHandle, "Decal Payment Edit Journal"
'  Print #RptHandle, "Posting Date: "; PostDate$
'  Print #RptHandle, "    Operator: "; PWUser; Tab(89); "Page #"; Page
'  Print #RptHandle, ""
'  Print #RptHandle, "       "; Tab(12); "        "; Tab(44); "           "; Tab(61); "            "; Tab(98); "Amount Paid"
'  Print #RptHandle, " Date"; Tab(11); "Acct No      Customer"; Tab(60); "Cash"; Tab(74); "Check"; Tab(87); "Charge"; Tab(98); " on Account"; Tab(115); "Change"
'  Print #RptHandle, Dash1$
'  Linecnt = 6
' End If
' Return
'
PrintRptEnding:
  Print #RptHandle, "Number of Entries .. "; Using$("##,###", TotalCust);
  Print #RptHandle, Tab(55); Using$("$##,###,#.##", TotalValue#)
  Print #RptHandle, FF$
  Page = Page + 1
  Print #RptHandle, ""
  Print #RptHandle, ""
  Print #RptHandle, Tab(20); "Vehicle Decals Payment EDIT Listing"
  Print #RptHandle, "Posting Date: "; PostDate$
  Print #RptHandle, "    Operator: "; PWUser; ","; Operator$; Tab(72); "Page #"; Page
  Print #RptHandle, ""
  Print #RptHandle, String$(80, "=")
  Print #RptHandle, ""
  Print #RptHandle, "Catagory Totals"
  Print #RptHandle, "Catagory"; Tab(20); "       Amount"
  For Lp = 1 To CatCnt
    Print #RptHandle, Cat$(Lp); Tab(20); Using$("$#,###,###.##", CatAmt#(Lp))
  Next Lp
 ' If Not Graph Then
    Print #RptHandle, Dash1$
    Print #RptHandle, " "
    Print #RptHandle, Tab(10); "     Total Cash: "; Using("###,###.##", TotalCash#)
    Print #RptHandle, Tab(10); "    Total Check: "; Using("###,###.##", TotalCheck#)
    Print #RptHandle, Tab(10); "   Total Charge: "; Using("###,###.##", TotalChrge#)
    Print #RptHandle, Tab(10); "                 "; "---------------"
    Print #RptHandle, Tab(10); " Total Received: "; Using("#######.##", Round#(TotalCash# + TotalCheck# + TotalChrge#))
    Print #RptHandle, " "
    Print #RptHandle, Tab(10); "   Total Change: "; Using("###,###.##", TotalChange#)
    Print #RptHandle, " "
    Print #RptHandle, Tab(10); "  Total Applied: "; Using("###,###.##", TotalAmount#)
    Print #RptHandle, " "
    Print #RptHandle, "Total Number of Receipts: "; Using("##,###", TotalReceipts)
    Print #RptHandle, "  Total Number of Checks: "; Using("##,###", TotalChks)
    If Not Graph Then Print #RptHandle, FF$
    GTotal# = 0
 ' Else
 '   GTotal# = 0
    SumPrnt$ = ""
'  End If
  Return

End Sub

