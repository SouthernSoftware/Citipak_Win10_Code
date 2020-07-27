VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmPaymentDelete 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Payment "
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   Icon            =   "frmPaymentDelete.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Caption         =   "Esc E&xit"
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
      Left            =   8310
      TabIndex        =   7
      Top             =   5496
      Width           =   1596
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "F10 &Delete"
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
      Left            =   6582
      TabIndex        =   6
      Top             =   5496
      Width           =   1596
   End
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   4020
      Left            =   2166
      TabIndex        =   0
      Top             =   2088
      Width           =   7884
      _Version        =   196609
      _ExtentX        =   13906
      _ExtentY        =   7091
      _StockProps     =   70
      Caption         =   ""
      Picture         =   "frmPaymentDelete.frx":08CA
      Begin LpLib.fpList fplstPayments 
         Height          =   2388
         Left            =   192
         TabIndex        =   1
         Top             =   360
         Width           =   7452
         _Version        =   196608
         _ExtentX        =   13144
         _ExtentY        =   4212
         TextAlias       =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Columns         =   5
         Sorted          =   0
         LineWidth       =   1
         SelDrawFocusRect=   -1  'True
         ColumnSeparatorChar=   9
         ColumnSearch    =   -1
         ColumnWidthScale=   2
         RowHeight       =   -1
         MultiSelect     =   1
         WrapList        =   0   'False
         WrapWidth       =   0
         SelMax          =   -1
         AutoSearch      =   1
         SearchMethod    =   0
         VirtualMode     =   0   'False
         VRowCount       =   0
         DataSync        =   3
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483627
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ScrollHScale    =   2
         ScrollHInc      =   0
         ColsFrozen      =   0
         ScrollBarV      =   1
         NoIntegralHeight=   0   'False
         HighestPrecedence=   0
         AllowColResize  =   0
         AllowColDragDrop=   0
         ReadOnly        =   0   'False
         VScrollSpecial  =   0   'False
         VScrollSpecialType=   0
         EnableKeyEvents =   -1  'True
         EnableTopChangeEvent=   -1  'True
         DataAutoHeadings=   -1  'True
         DataAutoSizeCols=   2
         SearchIgnoreCase=   -1  'True
         ScrollBarH      =   1
         VirtualPageSize =   0
         VirtualPagesAhead=   0
         ExtendCol       =   0
         ColumnLevels    =   1
         ListGrayAreaColor=   -2147483637
         GroupHeaderHeight=   -1
         GroupHeaderShow =   0   'False
         AllowGrpResize  =   0
         AllowGrpDragDrop=   0
         MergeAdjustView =   0   'False
         ColumnHeaderShow=   0   'False
         ColumnHeaderHeight=   -1
         GrpsFrozen      =   0
         BorderGrayAreaColor=   -2147483637
         ExtendRow       =   0
         DataField       =   ""
         OLEDragMode     =   0
         OLEDropMode     =   0
         Redraw          =   -1  'True
         ResizeRowToFont =   0   'False
         TextTipMultiLine=   0
         ColDesigner     =   "frmPaymentDelete.frx":08E6
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Spacebar or Click to Toggle, F10 to Continue. "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   516
         Left            =   144
         TabIndex        =   5
         Top             =   3432
         Width           =   4140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   7
         Left            =   3600
         TabIndex        =   4
         Top             =   24
         Width           =   1668
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   1
         Left            =   6360
         TabIndex        =   3
         Top             =   24
         Width           =   1044
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   2940
         Left            =   96
         Top             =   240
         Width           =   7668
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Left            =   648
         TabIndex        =   2
         Top             =   24
         Width           =   924
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   8
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
            TextSave        =   "3:04 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "2/15/2006"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LabelHead 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Payment to Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3612
      TabIndex        =   10
      Top             =   1104
      Width           =   5004
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   852
      Left            =   2580
      Top             =   840
      Width           =   7020
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Invoices For Payment"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12
      Left            =   3684
      TabIndex        =   9
      Top             =   1080
      Width           =   4836
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   996
      Left            =   2592
      Top             =   720
      Width           =   7020
   End
End
Attribute VB_Name = "frmPaymentDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim CashFlag As Boolean, uselook As Boolean, CustAcct As Long
Dim EditFlag As Boolean, TempAmtRecv As Double, Answer As Integer
Dim ChkOKFlag As Boolean, BeenDone As Boolean, PayListCnt As Long
Dim DistArray() As DistArrayType
Dim PayList() As PayListType
Dim codeopt As Integer, noreset As Boolean
Dim Oper As String, PayListRec As Long, RecpPort As Integer
Dim RevText$(1 To MaxRevsCnt)
'opt 1 means from payment delete, 2 is for deposit delete
Public Sub Wherefrom(opt As Integer)
  If opt <> 0 Then
    codeopt = opt
  Else
    codeopt = 0
  End If
End Sub

Private Sub cmdExit_Click()
'  noreset = True
'  Chk4Change
'  If Answer = 1 Then
'    Exit Sub
'  ElseIf Answer = 2 Then
'    fpCmdSave_Click
'  End If
'  CustAcct = 0
'  fpCustRecNo = 0
'  BeenDone = False
'  If codeopt = 1 Then
'    ActivateControls frmCustEditLookUP
'  ElseIf codeopt = 2 Then
'    ActivateControls frmDisplayList
'  End If
'  If codeopt = 0 Then
'    Load frmUBPaymentMenu
'    DoEvents
'    frmUBPaymentMenu.Show
'  End If
  If codeopt = 1 Then
    UBLog "OUT: UTIL Delete Payment" + " Oper:" + Oper$
  ElseIf codeopt = 2 Then
    UBLog "OUT: UTIL Delete Deposit" + " Oper:" + Oper$
  End If
  Unload Me
  DoEvents
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
        UBLog "Closed via PaymentDelete by " + PWUser$ + " operator-" + Oper$
        CitiTerminate
      End If
    End If
  End If

'  If ((UnloadMode = vbFormControlMenu)) Then
'    If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
'      Cancel = True
'    Else
'      If codeopt = 1 Then
'        UBLog "OUT: UTIL Delete Payment" + " Oper:" + Oper$
'      ElseIf codeopt = 2 Then
'        UBLog "OUT: UTIL Delete Deposit" + " Oper:" + Oper$
'      End If
'    End If
'  End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  Select Case KeyCode
'    Case vbKeyDown, vbKeyReturn:
'      SendKeys "{Tab}"
'      KeyCode = 0
'    Case vbKeyUp:
'      SendKeys "+{Tab}"
'      KeyCode = 0
    Case vbKeyEscape:
      KeyCode = 0
      DoEvents
      Call cmdExit_Click
    Case vbKeyF10:
      KeyCode = 0
      DoEvents
      Call cmdDelete_Click
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  
  If codeopt = 1 Then
    Me.HelpContextID = hlpDeleteAPayment
    UBLog " IN Oper " + Oper$ + ": UTIL Delete Payment"
    LoadPayList
  ElseIf codeopt = 2 Then
    Me.HelpContextID = hlpDeleteADeposit
    UBLog " IN Oper " + Oper$ + ": UTIL Delete Deposit"
    LabelHead.Caption = "Select Deposit to Delete"
    frmPaymentDelete.Caption = "Delete Deposit"
    LoadDepList
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
Private Sub LoadPayList()
  Dim cnt As Long, PHandle As Integer, ListFile As Integer
  Dim PayFileName As String, UBPayRecLen As Integer, PayListRec As Long
  Dim PayRecpName As String, NumOfRecs As Long, fmt As String
  Dim PCustAcct As Long, TPayFileName As String, TotalRecs As Long
  Dim tmp As String * 80
  ReDim UBPaymentRec(1) As UBPaymentRecType
  fmt$ = "#######.##"

  UBPayRecLen = Len(UBPaymentRec(1))
  
  Oper$ = Str$(OPERNUM)
  PayFileName$ = UBPath$ + "UBPAY" + QPTrim(Oper$) + ".DAT"
  TPayFileName$ = UBPath$ + "UBPAY" + QPTrim(Oper$) + ".$$$"

  'ReDim TransList(1 To TotalRecs&) As FLen2
 ' ReDim Picked(1 To TotalRecs&) As Integer

  'WhatCnt& = 1
  PHandle = FreeFile
  
  Open PayFileName$ For Random Shared As PHandle Len = UBPayRecLen
  NumOfRecs& = LOF(PHandle) \ UBPayRecLen
  'TotalRecs& = NumOfRecs
  If NumOfRecs& > 0 Then
    For cnt& = 1 To NumOfRecs&
      Get PHandle, cnt&, UBPaymentRec(1)
      tmp$ = QPTrim(UBPaymentRec(1).CustName) & Chr$(9) & "Payment" & Chr$(9) & Using$(fmt$, Str$(UBPaymentRec(1).AMTPAID)) & Chr$(9) & " " & Chr$(9) & cnt&
      
      fplstPayments.InsertRow = tmp$
      tmp$ = ""
    Next
  End If
  Close PHandle
  PayListCnt& = NumOfRecs&

End Sub
Private Sub LoadDepList()
  Dim cnt As Long, PHandle As Integer, ListFile As Integer
  Dim PayFileName As String, UBPayRecLen As Integer, PayListRec As Long
  Dim PayRecpName As String, NumOfRecs As Long, fmt As String
  Dim PCustAcct As Long, TPayFileName As String, TotalRecs As Long
  Dim tmp As String * 80
  ReDim UBPaymentRec(1) As UBPaymentRecType
  fmt$ = "#######.##"

  UBPayRecLen = Len(UBPaymentRec(1))
  
  Oper$ = Str$(OPERNUM)
  PayFileName$ = UBPath$ + "UBDEP" + QPTrim(Oper$) + ".DAT"
  TPayFileName$ = UBPath$ + "UBDEP" + QPTrim(Oper$) + ".$$$"

  'ReDim TransList(1 To TotalRecs&) As FLen2
 ' ReDim Picked(1 To TotalRecs&) As Integer

  'WhatCnt& = 1
  PHandle = FreeFile
  
  Open PayFileName$ For Random Shared As PHandle Len = UBPayRecLen
  NumOfRecs& = LOF(PHandle) \ UBPayRecLen
  'TotalRecs& = NumOfRecs
  If NumOfRecs& > 0 Then
    For cnt& = 1 To NumOfRecs&
      Get PHandle, cnt&, UBPaymentRec(1)
      tmp$ = QPTrim(UBPaymentRec(1).CustName) & Chr$(9) & "Deposit" & Chr$(9) & Using$(fmt$, Str$(UBPaymentRec(1).AMTPAID)) & Chr$(9) & " " & Chr$(9) & cnt&
      
      fplstPayments.InsertRow = tmp$
      tmp$ = ""
    Next
  End If
  Close PHandle
  PayListCnt& = NumOfRecs&

End Sub

Private Sub cmdDelete_Click()
  Dim FntSize As Integer
  Dim PCnt As Integer, NumPicked As Integer, PFile As Integer
  Dim PayFileName As String, UBPayRecLen As Integer, TFile As Integer
  Dim TPayFileName As String, cnt As Integer
  ReDim UBPaymentRec(1) As UBPaymentRecType
  If codeopt = 1 Then
    PayFileName$ = UBPath$ + "UBPAY" + QPTrim$(Str$(OPERNUM)) + ".DAT"
    TPayFileName$ = UBPath$ + "UBPAY" + QPTrim$(Str$(OPERNUM)) + ".$$$"
  ElseIf codeopt = 2 Then
    PayFileName$ = UBPath$ + "UBDEP" + QPTrim$(Str$(OPERNUM)) + ".DAT"
    TPayFileName$ = UBPath$ + "UBDEP" + QPTrim$(Str$(OPERNUM)) + ".$$$"
  End If
  UBPayRecLen = Len(UBPaymentRec(1))
  ReDim MsgText(0 To 5) As String
  For PCnt = 0 To fplstPayments.ListCount - 1
    If fplstPayments.Selected(PCnt) Then
      NumPicked = NumPicked + 1
    End If
  Next
  If Not NumPicked > 0 Then
    frmMsgDialog.RetLabel = "-2"
    FntSize = frmMsgDialog.Label(2).FontSize
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = ""
    MsgText(3) = "NO TRANSACTIONS SELECTED!"
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
    Exit Sub
  ElseIf NumPicked > 0 Then
    If MsgBox("Are You Sure You Wish to Continue With Deletion?", vbYesNo, "Delete") = vbYes Then
      PFile = FreeFile
      Open PayFileName$ For Random Shared As PFile Len = UBPayRecLen
      TFile = FreeFile
      Open TPayFileName$ For Random Shared As TFile Len = UBPayRecLen
      For PCnt = 0 To fplstPayments.ListCount - 1
        If Not fplstPayments.Selected(PCnt) Then
          fplstPayments.col = 4
          fplstPayments.Row = PCnt
          cnt = QPTrim(fplstPayments.ColList)
          Get PFile, cnt, UBPaymentRec(1)
          Put TFile, , UBPaymentRec(1)
        End If
      Next
      Close
      KillFile PayFileName$
      Name TPayFileName$ As PayFileName$
      If codeopt = 1 Then
        UBLog "Delete UTIL " + Str(NumPicked) + " Payment(s)" + " Oper:" + Oper$
      ElseIf codeopt = 2 Then
        UBLog "Delete UTIL " + Str(NumPicked) + " Deposit(s)" + " Oper:" + Oper$
      End If
    Else
      Exit Sub
    End If
  End If
  If FileSize(PayFileName$) = 0 Then
    KillFile PayFileName$
  End If
  MsgBox "Deletion Complete", vbOKOnly, "Complete"
  cmdExit_Click
End Sub
