VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmVATaxPayOperEntry 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment Operator Entry"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxPayOperEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin EditLib.fpDateTime txtDate1 
      Height          =   348
      Left            =   5730
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "The date you enter here will be the date that appears on the 'Payment Entry' screen. The date on that screen is not editable."
      Top             =   4770
      Width           =   1740
      _Version        =   196608
      _ExtentX        =   3069
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483637
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   2
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
      InvalidColor    =   -2147483643
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483643
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "05/13/2005"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19800101"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "20010101"
      TimeDefault     =   "000000"
      TimeStyle       =   0
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      PopUpType       =   1
      DateCalcY2KSplit=   60
      CaretPosition   =   0
      IncYear         =   1
      IncMonth        =   1
      IncDay          =   1
      IncHour         =   1
      IncMinute       =   1
      IncSecond       =   1
      ButtonColor     =   13684944
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn CmdExit 
      Height          =   630
      Left            =   3120
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   "Press this button to return to the business license main menu."
      Top             =   6000
      Width           =   2160
      _Version        =   131072
      _ExtentX        =   3810
      _ExtentY        =   1111
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
      ButtonDesigner  =   "frmVATaxPayOperEntry.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdOk 
      Height          =   630
      Left            =   6360
      TabIndex        =   10
      TabStop         =   0   'False
      Tag             =   $"frmVATaxPayOperEntry.frx":0AA6
      Top             =   6000
      Width           =   2160
      _Version        =   131072
      _ExtentX        =   3810
      _ExtentY        =   1111
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
      ButtonDesigner  =   "frmVATaxPayOperEntry.frx":0B5B
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operator Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   336
      Left            =   3234
      TabIndex        =   8
      Top             =   3498
      Width           =   2328
   End
   Begin VB.Label lblOperName 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   5730
      TabIndex        =   7
      Tag             =   "This field indicates the current user's operator number. This field is not editable."
      Top             =   3450
      Width           =   2184
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operator Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   336
      Left            =   3234
      TabIndex        =   6
      Top             =   3018
      Width           =   2328
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   396
      Index           =   1
      Left            =   3474
      TabIndex        =   5
      Top             =   4842
      Width           =   2088
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Source:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   336
      Left            =   3402
      TabIndex        =   4
      Top             =   4026
      Width           =   2160
   End
   Begin VB.Label lblOperator 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   5730
      TabIndex        =   3
      Tag             =   "This field indicates the current operator number. This field is not editable."
      Top             =   2970
      Width           =   1068
   End
   Begin VB.Label lblSource 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   5730
      TabIndex        =   2
      Tag             =   "This field refers to the business license program currently operating."
      Top             =   3930
      Width           =   2184
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Operator Entry"
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
      Left            =   3816
      TabIndex        =   1
      Top             =   1776
      Width           =   4020
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   720
      Left            =   2304
      Top             =   1596
      Width           =   7020
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   2652
      Left            =   2760
      Top             =   2772
      Width           =   6204
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   2802
      X2              =   8922
      Y1              =   4530
      Y2              =   4530
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   864
      Left            =   2316
      Top             =   1476
      Width           =   7020
   End
End
Attribute VB_Name = "frmVATaxPayOperEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim BadDate As Boolean
Private Sub cmdExit_Click()
  frmVATaxMainMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub fpCmdOk_Click()
  frmVATaxLoadingRpt.Label1 = "Searching For Receipt Printer"
  frmVATaxLoadingRpt.Show
  DoEvents
  If RcpCheck = False Then
    frmVATaxMsgWOpts.Label1.Caption = "ATTENTION: RECEIPT PRINTING WILL NOT BE POSSIBLE BECAUSE THE CURRENT PATH TO THE RECEIPT PRINTER, " + RecpPort + ", CANNOT BE FOUND. The receipt printer path is administered from the Citipak main menu under the 'Receipt Printer Setup' button. Please refer to that link to make receipt printer path corrections. Do you wish to continue to the 'Payment Entry' menu anyway?"
    frmVATaxMsgWOpts.Label1.Top = 300
    frmVATaxMsgWOpts.Label1.Height = 1850
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.cmdExit.Text = "ESC Abort Load"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      RecpDef = 98
      MainLog ("Upon loading the payment entry menu the user was warned that receipt printing would not be possible because the receipt printer path was not found. The user elected to continue loading the payment entry screen.")
    Else
      Unload frmVATaxMsgWOpts
      Unload frmVATaxLoadingRpt
      Exit Sub
    End If
  End If
  
  CheckPayDate
  If BadDate = False Then
    'do stuff
    PayDate = txtDate1.Text
    Call savePDate
    frmVATaxPayMenu.Show
    DoEvents
    Unload frmVATaxLoadingRpt
    Unload Me
    
  Else
    MsgBox "Invalid Date", vbOKOnly, "Invalid Entry"
  End If
  Unload frmVATaxLoadingRpt
End Sub

Private Sub txtDate1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpCmdOk.SetFocus
  End If
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        MainLog ("Closed via frmVATaxPayOperEntry by " + PWUser$)
        Call Terminate
      End If
    End If
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      KeyCode = 0
      DoEvents
      Call cmdExit_Click
    Case vbKeyF10:
      KeyCode = 0
      DoEvents
      Call fpCmdOk_Click
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Me.HelpContextID = hlpPaymentOperator
  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.Visible Then
    'Temp_Class.ResizeControls Me
    DoEvents
  End If
End Sub
Private Sub CheckPayDate()
'Dim PayDate As String
  PayDate$ = txtDate1.Text
  If Val(Left$(PayDate$, 2)) < 1 Or Val(Left$(PayDate$, 2)) > 12 Then
    If Val(Mid$(PayDate$, 4, 2)) < 1 Or Val(Mid$(PayDate$, 4, 2)) > 31 Then
      BadDate = True
    Else
      BadDate = False
    End If
  Else
    BadDate = False
  End If
End Sub

Private Sub LoadMe()
  txtDate1.Text = Format(Now, "mm/dd/yyyy")
  lblOperator = OperNum
  lblOperName.Caption = 1 'PWUser
  lblSource.Caption = "Tax Billing"
  GCustNum = 0
  Call GetRcpInfo
  Call GetPayDate
End Sub

Private Function RcpCheck() As Boolean
  Dim PHandle As Integer
  
  On Local Error GoTo NoPathError
  RcpCheck = True
  PHandle = FreeFile
'  RecpPort = "Dell Laser Printer S2500" 'use only for testing
'  MsgBox "RecpPort = " + RecpPort
  Open RecpPort For Output As PHandle
  Close PHandle
  
  Exit Function
  
NoPathError:
  RcpCheck = False
  Close PHandle
  
End Function
Public Sub GetPayDate()
  Dim lenRP As Integer, RP1 As Integer, gpay As Integer
  Dim RcptPrnFile As ReceiptPRNType
  
  On Error GoTo ERRORSTUFF
  
  RP1 = FreeFile
  lenRP = Len(RcptPrnFile)
'  If Exist("C:\RcptPrn.dat") Then
'    Open "c:\RcptPrn.dat" For Random Shared As RP1 Len = lenRP
  If Exist(RcptFileName$) Then '2/14/08
    Open RcptFileName$ For Random Shared As RP1 Len = lenRP '2/14/08
    Get RP1, 1, RcptPrnFile
      gpay = RcptPrnFile.PaymDate
    Close RP1
  End If
  If gpay > Date2Num(txtDate1.Text) Then
    txtDate1.Text = MakeRegDate(gpay)
  Else
    txtDate1.Text = CStr(Date)
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPayOperEntry", "GetPayDate", Erl)
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
    ClearInUse PWcnt
    Terminate

End Sub

Private Sub savePDate()
  Dim RP1 As Integer, lenRP As Integer
  Dim RcptPrnFile As ReceiptPRNType
  
  On Error GoTo ERRORSTUFF
  
  RP1 = FreeFile
  lenRP = Len(RcptPrnFile)
'  Open "c:\RcptPrn.dat" For Random Shared As RP1 Len = lenRP
  Open RcptFileName$ For Random Shared As RP1 Len = lenRP '2/14/08
    Get #RP1, 1, RcptPrnFile
    RcptPrnFile.PaymDate = Date2Num(txtDate1.Text)
    Put #RP1, 1, RcptPrnFile
  Close
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPayOperEntry", "savePDate", Erl)
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
    ClearInUse PWcnt
    Terminate

End Sub
