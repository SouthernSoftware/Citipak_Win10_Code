VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Begin VB.Form frmUBControlMaint 
   AutoRedraw      =   -1  'True
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Maintenance"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1896
   ClientWidth     =   12216
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   10.2
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "6frmUBControlMaint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn fpCmdSave 
      Height          =   384
      Left            =   8160
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7920
      Width           =   1236
      _Version        =   131072
      _ExtentX        =   2180
      _ExtentY        =   677
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
      ButtonDesigner  =   "6frmUBControlMaint.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdExit 
      Height          =   384
      Left            =   9456
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7920
      Width           =   1236
      _Version        =   131072
      _ExtentX        =   2180
      _ExtentY        =   677
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
      ButtonDesigner  =   "6frmUBControlMaint.frx":0AA5
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   336
      Left            =   0
      TabIndex        =   0
      Top             =   8532
      Width           =   12216
      _ExtentX        =   21548
      _ExtentY        =   593
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
            TextSave        =   "3:02 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "5/23/2005"
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
End
Attribute VB_Name = "frmUBControlMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim BeenDone As Boolean

Private Sub btnPgUp_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 1 Then
    DoEvents
    SendKeys "{PgUp}", True
  End If
End Sub

Private Sub btnPgDn_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 1 Then
    DoEvents
    SendKeys "{PgDn}", True
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  BeenDone = False   'clear variable for revenue waring
  If ((UnloadMode = vbFormControlMenu)) Then
    If fpCmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via control maint by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub
Private Sub fpCmdExit_Click()
  Select Case CheckSaveControlFile
  Case True:  '-1 save chenges
    Call fpCmdSave_Click
  Case False:  '0= exit
    Call ExitControlMaint
  Case Else     '1 is review
    'continue editing
  End Select
End Sub

Private Sub fpCmdSave_Click()
  Call SaveControlFile
  Call ExitControlMaint
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  fpcboLockbox.InsertRow = "6" & Chr$(9) & "6 Digit Account"
  fpcboLockbox.InsertRow = "8" & Chr$(9) & "8 Digit Account"
  fpcboLookUp.InsertRow = "1" & Chr$(9) & "Account Number"
  fpcboLookUp.InsertRow = "2" & Chr$(9) & "Search Name"
  fpcboLookUp.InsertRow = "3" & Chr$(9) & "Meter Number"
  fpcboLookUp.InsertRow = "4" & Chr$(9) & "Service Address"
  fpcboLookUp.InsertRow = "5" & Chr$(9) & "Location Number"
  fpcboLookUp.InsertRow = "6" & Chr$(9) & "911 Address"
  LoadUpdateForm
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    'DoEvents
    Temp_Class.ResizeControls Me
    'Me.Visible = True
    'DoEvents
  End If
  DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      KeyCode = 0
      Call fpCmdExit_Click
    Case vbKeyPageDown
      If Not fpHHDevice.ListDown Then
        If vaTabPro1.ActiveTab < 4 Then
          vaTabPro1.ActiveTab = vaTabPro1.ActiveTab + 1
        Else
          vaTabPro1.ActiveTab = 0
        End If
      End If
    Case vbKeyPageUp
      If Not fpHHDevice.ListDown Then
        If vaTabPro1.ActiveTab > 0 Then
          vaTabPro1.ActiveTab = vaTabPro1.ActiveTab - 1
        Else
          vaTabPro1.ActiveTab = 4
        End If
      End If
    Case vbKeyF10
      Call SaveControlFile
      Call UPDateOK

      Call ExitControlMaint
    Case Else:
  End Select
End Sub

Private Sub LoadUpdateForm()
  Dim UBSetUpRec(1) As UBSetupRecType
  Dim UBSetupLen As Integer, cnt As Integer
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  Me.fpUtilName = QPTrim$(UBSetUpRec(1).UTILNAME)
  Me.fpDefCity = QPTrim$(UBSetUpRec(1).DEFCITY)
  Me.fpDefState = QPTrim$(UBSetUpRec(1).DEFSTATE)
  Me.fpZipCode = QPTrim$(UBSetUpRec(1).ZIPCODE)
  Me.fpPreByBook = QPTrim$(UBSetUpRec(1).PreByBook)
  If Val(UBSetUpRec(1).LockBoxDef) = 6 Then
    Me.fpcboLockbox.ListIndex = 0
  Else
    Me.fpcboLockbox.ListIndex = 1
  End If
  Me.fpRecpDeft = QPTrim$(UBSetUpRec(1).RECPDEFT)
  Me.fpEstRead = QPTrim$(UBSetUpRec(1).ESTREAD)
  Me.fpBankDft = QPTrim$(UBSetUpRec(1).BANKDFT)
  Me.fpUseSeq = QPTrim$(UBSetUpRec(1).UseSeq)
  Me.fpBillCycl = QPTrim$(UBSetUpRec(1).BILLCYCL)
  Me.fpMethAcct = QPTrim$(UBSetUpRec(1).MethAcct)
  Me.fpSkipInact = QPTrim$(UBSetUpRec(1).SkipInactive)
  Me.fpSkipSeprat = QPTrim$(UBSetUpRec(1).SkipSeparator)
  Me.fpMake99File = QPTrim$(UBSetUpRec(1).Make99File)
  Me.fpLowRead = QPTrim$(Str$(UBSetUpRec(1).LowRead))
  Me.fpHighRead = QPTrim$(Str$(UBSetUpRec(1).HighRead))
  If Val(UBSetUpRec(1).DefLook) > 0 Then
    fpcboLookUp.ListIndex = Val(UBSetUpRec(1).DefLook) - 1
  Else
    fpcboLookUp.ListIndex = 0
  End If
  'Me.fpDefLook = QPTrim$(UBSetUpRec(1).DefLook)
  For cnt = 1 To 15
    Me.fpRevSource(cnt - 1).Text = QPTrim$(UBSetUpRec(1).Revenues(cnt).RevName)
    Me.PG3RevLBL(cnt - 1).Caption = QPTrim$(UBSetUpRec(1).Revenues(cnt).RevName)
    Me.PG4RevLBL(cnt - 1).Caption = QPTrim$(UBSetUpRec(1).Revenues(cnt).RevName)
    Me.PG5RevLBL(cnt - 1).Caption = QPTrim$(UBSetUpRec(1).Revenues(cnt).RevName)
    Me.fpBilDebAct(cnt - 1).Text = QPTrim$(UBSetUpRec(1).BillAcct(cnt).DebitAcct)
    Me.fpBilCrdAct(cnt - 1).Text = QPTrim$(UBSetUpRec(1).BillAcct(cnt).CreditAcct)
    Me.fpPayDebAct(cnt - 1).Text = QPTrim$(UBSetUpRec(1).PayAcct(cnt).DebitAcct)
    Me.fpPayCrdAct(cnt - 1).Text = QPTrim$(UBSetUpRec(1).PayAcct(cnt).CreditAcct)
    Me.fpDepDebAct(cnt - 1).Text = QPTrim$(UBSetUpRec(1).DepAcct(cnt).DebitAcct)
    Me.fpDepCrdAct(cnt - 1).Text = QPTrim$(UBSetUpRec(1).DepAcct(cnt).CreditAcct)
    Me.fpTextUseDep(cnt - 1).Text = QPTrim$(UBSetUpRec(1).Revenues(cnt).UseDep)
    Me.fpUseRate(cnt - 1) = QPTrim$(UBSetUpRec(1).Revenues(cnt).USERATE)
    Me.fpTaxRate(cnt - 1) = UBSetUpRec(1).Revenues(cnt).TAXRATE
    Me.fpMetered(cnt - 1).Text = QPTrim$(UBSetUpRec(1).Revenues(cnt).UseMtr)
    Me.fpDefDist(cnt - 1) = UBSetUpRec(1).Revenues(cnt).DistOr
    Me.fpProrate(cnt - 1).Text = QPTrim$(UBSetUpRec(1).Revenues(cnt).ProRate)
  Next
  
  fpHHDevice.SearchText = QPTrim$(UBSetUpRec(1).HHDEVICE)
  fpHHDevice.ColumnSearch = 0
  fpHHDevice.Action = ActionSearch
  If fpHHDevice.SearchIndex <> -1 Then
    fpHHDevice.ListIndex = fpHHDevice.SearchIndex
  End If
  fpHHDevice.ColumnSearch = 0

End Sub


Private Sub vaTabPro1_TabPageShown(ActiveTab As Integer, ActivePage As Integer)
  Dim cnt As Integer
  Select Case ActiveTab
  Case 0
    Me.fpUtilName.SetFocus
  Case 1
    Me.fpRevSource(0).SetFocus
  Case 2
    Me.fpBilDebAct(0).SetFocus
  Case 3
    Me.fpPayDebAct(0).SetFocus
  Case 4
    Me.fpDepDebAct(0).SetFocus
  End Select
  For cnt = 1 To 15
    Me.PG3RevLBL(cnt - 1).Caption = Me.fpRevSource(cnt - 1).Text
    Me.PG4RevLBL(cnt - 1).Caption = Me.fpRevSource(cnt - 1).Text
    Me.PG5RevLBL(cnt - 1).Caption = Me.fpRevSource(cnt - 1).Text
  Next
  
  If ActiveTab = 1 And Not BeenDone Then
    BeenDone = True
    Load frmRevWarning
    frmRevWarning.Show vbModal, Me
  End If
  
End Sub

Private Function CheckSaveControlFile%()
  Dim UBSetUpRec(1) As UBSetupRecType
  Dim Temp As String
  Dim UBSetupLen As Integer, Changed As Boolean, cnt As Integer
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  
  If QPTrim$(Me.fpUtilName) <> QPTrim$(UBSetUpRec(1).UTILNAME) Then
    Changed = True
    GoTo ExitCheck
  End If
  If QPTrim$(Me.fpDefCity) <> QPTrim$(UBSetUpRec(1).DEFCITY) Then
    Changed = True
    GoTo ExitCheck
  End If
  If QPTrim$(Me.fpDefState) <> QPTrim$(UBSetUpRec(1).DEFSTATE) Then
    Changed = True
    GoTo ExitCheck
  End If
  
  Temp$ = QPTrim$(Me.fpZipCode)
  If Right$(Temp$, 1) = "-" Then
    Temp$ = Left$(Temp$, Len(Temp$) - 1)
  End If
  If Temp$ <> QPTrim$(UBSetUpRec(1).ZIPCODE) Then
    Changed = True
    GoTo ExitCheck
  End If
  
  If QPTrim$(Me.fpPreByBook) <> QPTrim$(UBSetUpRec(1).PreByBook) Then
    Changed = True
    GoTo ExitCheck
  End If
'  If QPTrim$(Me.fpRecpPort) <> QPTrim$(UBSetUpRec(1).RecpPort) Then
'    Changed = True
'    GoTo ExitCheck
'  End If
  fpcboLockbox.col = 0
  If QPTrim$(fpcboLockbox.ColText) <> QPTrim$(UBSetUpRec(1).LockBoxDef) Then
    Changed = True
    GoTo ExitCheck
  End If

  If QPTrim$(Me.fpRecpDeft) <> QPTrim$(UBSetUpRec(1).RECPDEFT) Then
    Changed = True
    GoTo ExitCheck
  End If
  If QPTrim$(Me.fpEstRead) <> QPTrim$(UBSetUpRec(1).ESTREAD) Then
    Changed = True
    GoTo ExitCheck
  End If
  If QPTrim$(Me.fpBankDft) <> QPTrim$(UBSetUpRec(1).BANKDFT) Then
    Changed = True
    GoTo ExitCheck
  End If
  If QPTrim$(Me.fpUseSeq) <> QPTrim$(UBSetUpRec(1).UseSeq) Then
    Changed = True
    GoTo ExitCheck
  End If
  If QPTrim$(Me.fpBillCycl) <> QPTrim$(UBSetUpRec(1).BILLCYCL) Then
    Changed = True
    GoTo ExitCheck
  End If
  fpcboLookUp.col = 0
  If QPTrim$(fpcboLookUp.ColText) <> QPTrim$(UBSetUpRec(1).DefLook) Then
    Changed = True
    GoTo ExitCheck
  End If
  If QPTrim$(Me.fpMethAcct) <> QPTrim$(UBSetUpRec(1).MethAcct) Then
    Changed = True
    GoTo ExitCheck
  End If
  If QPTrim$(Me.fpSkipInact) <> QPTrim$(UBSetUpRec(1).SkipInactive) Then
    Changed = True
    GoTo ExitCheck
  End If
  If QPTrim$(Me.fpSkipSeprat) <> QPTrim$(UBSetUpRec(1).SkipSeparator) Then
    Changed = True
    GoTo ExitCheck
  End If
  If QPTrim$(Me.fpMake99File) <> QPTrim$(UBSetUpRec(1).Make99File) Then
    Changed = True
    GoTo ExitCheck
  End If
  If QPTrim$(Me.fpLowRead) <> QPTrim$(Str$(UBSetUpRec(1).LowRead)) Then
    Changed = True
    GoTo ExitCheck
  End If
  If QPTrim$(Me.fpHighRead) <> QPTrim$(Str$(UBSetUpRec(1).HighRead)) Then
    Changed = True
    GoTo ExitCheck
  End If
  
  For cnt = 1 To 15
    If QPTrim$(Me.fpRevSource(cnt - 1).Text) <> QPTrim$(UBSetUpRec(1).Revenues(cnt).RevName) Then
      Changed = True
      GoTo ExitCheck
    End If
    If QPTrim$(Me.fpBilDebAct(cnt - 1).Text) <> QPTrim$(UBSetUpRec(1).BillAcct(cnt).DebitAcct) Then
      Changed = True
      GoTo ExitCheck
    End If
    If QPTrim$(Me.fpBilCrdAct(cnt - 1).Text) <> QPTrim$(UBSetUpRec(1).BillAcct(cnt).CreditAcct) Then
      Changed = True
      GoTo ExitCheck
    End If
    If QPTrim$(Me.fpPayDebAct(cnt - 1).Text) <> QPTrim$(UBSetUpRec(1).PayAcct(cnt).DebitAcct) Then
      Changed = True
      GoTo ExitCheck
    End If
    If QPTrim$(Me.fpPayCrdAct(cnt - 1).Text) <> QPTrim$(UBSetUpRec(1).PayAcct(cnt).CreditAcct) Then
      Changed = True
      GoTo ExitCheck
    End If
    If QPTrim$(Me.fpDepDebAct(cnt - 1).Text) <> QPTrim$(UBSetUpRec(1).DepAcct(cnt).DebitAcct) Then
      Changed = True
      GoTo ExitCheck
    End If
    If QPTrim$(Me.fpDepCrdAct(cnt - 1).Text) <> QPTrim$(UBSetUpRec(1).DepAcct(cnt).CreditAcct) Then
      Changed = True
      GoTo ExitCheck
    End If
    If QPTrim$(Me.fpTextUseDep(cnt - 1).Text) <> QPTrim$(UBSetUpRec(1).Revenues(cnt).UseDep) Then
      Changed = True
      GoTo ExitCheck
    End If
    If QPTrim$(Me.fpUseRate(cnt - 1)) <> QPTrim$(UBSetUpRec(1).Revenues(cnt).USERATE) Then
      Changed = True
      GoTo ExitCheck
    End If
    If QPTrim$(Me.fpTaxRate(cnt - 1)) <> UBSetUpRec(1).Revenues(cnt).TAXRATE Then
      Changed = True
      GoTo ExitCheck
    End If
    If QPTrim$(Me.fpMetered(cnt - 1).Text) <> QPTrim$(UBSetUpRec(1).Revenues(cnt).UseMtr) Then
      Changed = True
      GoTo ExitCheck
    End If
    If Val(QPTrim$(Me.fpDefDist(cnt - 1))) <> UBSetUpRec(1).Revenues(cnt).DistOr Then
      Changed = True
      GoTo ExitCheck
    End If
    If QPTrim$(Me.fpProrate(cnt - 1).Text) <> QPTrim$(UBSetUpRec(1).Revenues(cnt).ProRate) Then
      Changed = True
      GoTo ExitCheck
    End If
  Next
  Temp$ = Left$(QPTrim$(fpHHDevice.Text), 1)
  If Temp$ <> QPTrim$(UBSetUpRec(1).HHDEVICE) Then
    Changed = True
  End If
  
ExitCheck:
  If Changed Then
    Load frmChangedWarning
    frmChangedWarning.Show vbModal, Me
    Select Case SaveFlag
    Case False
      CheckSaveControlFile = False
    Case True
      CheckSaveControlFile = True
    Case 1
      CheckSaveControlFile = 1
    End Select
  Else
    CheckSaveControlFile = False
  End If
End Function

Private Sub SaveControlFile()
  Dim UBSetUpRec(1) As UBSetupRecType
  Dim fpZip As String
  Dim Handle As Integer
  Dim UBSetupLen As Integer, cnt As Integer
  
  LSet UBSetUpRec(1).UTILNAME = QPTrim$(Me.fpUtilName)
  LSet UBSetUpRec(1).DEFCITY = QPTrim$(Me.fpDefCity)
  LSet UBSetUpRec(1).DEFSTATE = QPTrim$(Me.fpDefState)
  
  fpZip = QPTrim$(Me.fpZipCode)
  If Right$(fpZip$, 1) = "-" Then
    fpZip$ = Left$(fpZip$, Len(fpZip$) - 1)
  End If
  LSet UBSetUpRec(1).ZIPCODE = fpZip$
  LSet UBSetUpRec(1).PreByBook = QPTrim$(Me.fpPreByBook)
  'LSet UBSetUpRec(1).RecpPort = QPTrim$(Me.fpRecpPort)
  fpcboLockbox.col = 0
  LSet UBSetUpRec(1).LockBoxDef = QPTrim$(fpcboLockbox.ColText)
  LSet UBSetUpRec(1).RECPDEFT = QPTrim$(Me.fpRecpDeft)
  LSet UBSetUpRec(1).ESTREAD = QPTrim$(Me.fpEstRead)
  LSet UBSetUpRec(1).BANKDFT = QPTrim$(Me.fpBankDft)
  LSet UBSetUpRec(1).UseSeq = QPTrim$(Me.fpUseSeq)
  LSet UBSetUpRec(1).BILLCYCL = QPTrim$(Me.fpBillCycl)
  fpcboLookUp.col = 0
  LSet UBSetUpRec(1).DefLook = QPTrim$(fpcboLookUp.ColText)
  LSet UBSetUpRec(1).MethAcct = QPTrim$(Me.fpMethAcct)
  LSet UBSetUpRec(1).SkipInactive = QPTrim$(Me.fpSkipInact)
  LSet UBSetUpRec(1).SkipSeparator = QPTrim$(Me.fpSkipSeprat)
  LSet UBSetUpRec(1).Make99File = QPTrim$(Me.fpMake99File)
  UBSetUpRec(1).LowRead = Val(QPTrim$(Me.fpLowRead))
  UBSetUpRec(1).HighRead = Val(QPTrim$(Me.fpHighRead))
  
  For cnt = 1 To 15
    LSet UBSetUpRec(1).Revenues(cnt).RevName = QPTrim$(Me.fpRevSource(cnt - 1).Text)
    LSet UBSetUpRec(1).BillAcct(cnt).DebitAcct = QPTrim$(Me.fpBilDebAct(cnt - 1).Text)
    LSet UBSetUpRec(1).BillAcct(cnt).CreditAcct = QPTrim$(Me.fpBilCrdAct(cnt - 1).Text)
    LSet UBSetUpRec(1).PayAcct(cnt).DebitAcct = QPTrim$(Me.fpPayDebAct(cnt - 1).Text)
    LSet UBSetUpRec(1).PayAcct(cnt).CreditAcct = QPTrim$(Me.fpPayCrdAct(cnt - 1).Text)
    LSet UBSetUpRec(1).DepAcct(cnt).DebitAcct = QPTrim$(Me.fpDepDebAct(cnt - 1).Text)
    LSet UBSetUpRec(1).DepAcct(cnt).CreditAcct = QPTrim$(Me.fpDepCrdAct(cnt - 1).Text)
    LSet UBSetUpRec(1).Revenues(cnt).UseDep = QPTrim$(Me.fpTextUseDep(cnt - 1).Text)
    LSet UBSetUpRec(1).Revenues(cnt).USERATE = QPTrim$(Me.fpUseRate(cnt - 1))
    UBSetUpRec(1).Revenues(cnt).TAXRATE = Me.fpTaxRate(cnt - 1)
    LSet UBSetUpRec(1).Revenues(cnt).UseMtr = QPTrim$(Me.fpMetered(cnt - 1).Text)
    UBSetUpRec(1).Revenues(cnt).DistOr = Me.fpDefDist(cnt - 1)
    LSet UBSetUpRec(1).Revenues(cnt).ProRate = QPTrim$(Me.fpProrate(cnt - 1).Text)
  Next
  
  LSet UBSetUpRec(1).HHDEVICE = QPTrim$(fpHHDevice.Text)
  
  UBSetupLen = Len(UBSetUpRec(1))
  Handle = FreeFile
  Open UBPath$ + "UBSETUP.DAT" For Random Shared As Handle Len = UBSetupLen    'open data file
  Put #Handle, 1, UBSetUpRec(1)
  Close Handle
  TOWNNAME$ = QPTrim$(UBSetUpRec(1).UTILNAME)
  
End Sub

Private Sub ExitControlMaint()
  Load frmUBSetupMenu
  DoEvents
  frmUBSetupMenu.Show
  Unload frmUBControlMaint
End Sub
'~~~~~~~~~~~~~~~~~~First page keydowns~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub fpUtilName_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyDown, vbKeyReturn
    KeyCode = 0
    Me.fpDefCity.SetFocus
  Case vbKeyUp
    KeyCode = 0
    SendKeys "{pgup}"
  End Select
End Sub
Private Sub fpDefCity_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyDown, vbKeyReturn
    KeyCode = 0
    Me.fpDefState.SetFocus
  Case vbKeyUp
    KeyCode = 0
    fpUtilName.SetFocus
  End Select

End Sub
Private Sub fpDefState_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyDown, vbKeyReturn
    KeyCode = 0
    fpZipCode.SetFocus
  Case vbKeyUp
    KeyCode = 0
    fpDefCity.SetFocus
  End Select
End Sub
Private Sub fpZipCode_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyDown, vbKeyReturn
    KeyCode = 0
    fpPreByBook.SetFocus
  Case vbKeyUp
    KeyCode = 0
    fpDefState.SetFocus
  End Select
End Sub
Private Sub fpPreByBook_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyDown, vbKeyReturn
    KeyCode = 0
    fpcboLockbox.SetFocus
  Case vbKeyUp
    KeyCode = 0
    fpZipCode.SetFocus
  End Select
End Sub
Private Sub fpcboLockbox_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboLockbox.ListDown = True
  End If
  If fpcboLockbox.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Then
     KeyCode = 0
     fpRecpDeft.SetFocus
    Else
      If KeyCode = vbKeyUp Or KeyCode = vbKeyLeft Then
        KeyCode = 0
        fpPreByBook.SetFocus
      End If
      If KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
        KeyCode = 0
      End If
    End If
  End If

  
End Sub
Private Sub fpRecpDeft_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyDown, vbKeyReturn
    KeyCode = 0
    fpEstRead.SetFocus
  Case vbKeyUp
    KeyCode = 0
    fpcboLockbox.SetFocus
  End Select
End Sub
Private Sub fpEstRead_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyDown, vbKeyReturn
    KeyCode = 0
    fpBankDft.SetFocus
  Case vbKeyUp
    KeyCode = 0
    fpRecpDeft.SetFocus
  End Select
End Sub
Private Sub fpBankDft_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyDown, vbKeyReturn
    KeyCode = 0
    fpUseSeq.SetFocus
  Case vbKeyUp
    KeyCode = 0
    fpEstRead.SetFocus
  End Select
End Sub
Private Sub fpUseSeq_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyDown, vbKeyReturn
    KeyCode = 0
    fpBillCycl.SetFocus
  Case vbKeyUp
    KeyCode = 0
    fpBankDft.SetFocus
  End Select
End Sub
Private Sub fpBillCycl_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyDown, vbKeyReturn
    KeyCode = 0
    fpMethAcct.SetFocus
  Case vbKeyUp
    KeyCode = 0
    fpUseSeq.SetFocus
  End Select
End Sub
Private Sub fpMethAcct_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyDown, vbKeyReturn
    KeyCode = 0
    fpSkipInact.SetFocus
  Case vbKeyUp
    KeyCode = 0
    fpBillCycl.SetFocus
  End Select
End Sub
Private Sub fpSkipInact_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyDown, vbKeyReturn
    KeyCode = 0
    fpSkipSeprat.SetFocus
  Case vbKeyUp
    KeyCode = 0
    fpMethAcct.SetFocus
  End Select
End Sub
Private Sub fpSkipSeprat_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyDown, vbKeyReturn
    KeyCode = 0
    fpMake99File.SetFocus
  Case vbKeyUp
    KeyCode = 0
    fpSkipInact.SetFocus
  End Select
End Sub
Private Sub fpMake99File_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyDown, vbKeyReturn
    KeyCode = 0
    fpcboLookUp.SetFocus
  Case vbKeyUp
    KeyCode = 0
    fpSkipSeprat.SetFocus
  End Select
End Sub
Private Sub fpcboLookUp_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboLookUp.ListDown = True
  End If
  If fpcboLookUp.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Then
      fpHHDevice.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Or KeyCode = vbKeyLeft Then
        fpMake99File.SetFocus
        KeyCode = 0
      End If
      If KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
        KeyCode = 0
      End If
    End If
  End If

End Sub
Private Sub fpHHDevice_KeyDown(KeyCode As Integer, Shift As Integer)
  If Not fpHHDevice.ListDown Then
    Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      Me.fpLowRead.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpcboLookUp.SetFocus
    Case vbKeyPageUp, vbKeyPageDown
      KeyCode = 0
    Case vbKeySpace
      fpHHDevice.ListDown = True
    End Select
'  Else
'    Select Case KeyCode
'    Case vbKeyPageUp, vbKeyPageDown
'      KeyCode = 0
'    End Select
  End If
End Sub
Private Sub fpLowRead_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyDown, vbKeyReturn
    KeyCode = 0
    fpHighRead.SetFocus
  Case vbKeyUp
    KeyCode = 0
    fpHHDevice.SetFocus
  End Select
End Sub
Private Sub fpHighRead_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyDown, vbKeyReturn
    KeyCode = 0
    SendKeys "{pgdn}"
  Case vbKeyUp
    KeyCode = 0
    Me.fpLowRead.SetFocus
  End Select
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~End of Page 1 keydowns
'
'~~~~~~~~~~~~~~~~~~~~~~Page 2 keydowns
Private Sub fpRevSource_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
    fpTextUseDep(Index).SetFocus
  ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyLeft Then
    If Index > 0 Then
      fpProrate(Index - 1).SetFocus
    Else
      SendKeys "{pgup}"
    End If
  End If
'  If Index = 0 Then
'    Select Case KeyCode
'    Case vbKeyUp:
'      SendKeys "{pgup}"
'    Case vbKeyDown:
'      Me.fpRevSource(1).SetFocus
'    Case Else:
'    End Select
'  End If
End Sub
Private Sub fpTextUseDep_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
    fpUseRate(Index).SetFocus
  ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Then
    fpRevSource(Index).SetFocus
  End If
End Sub
Private Sub fpUseRate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
    fpTaxRate(Index).SetFocus
  ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Then
    fpTextUseDep(Index).SetFocus
  End If
End Sub
Private Sub fpTaxRate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
    fpMetered(Index).SetFocus
  ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Then
    fpUseRate(Index).SetFocus
  End If
End Sub
Private Sub fpMetered_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
    fpDefDist(Index).SetFocus
  ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Then
    fpTaxRate(Index).SetFocus
  End If
End Sub
Private Sub fpDefDist_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
    fpProrate(Index).SetFocus
  ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Then
    fpMetered(Index).SetFocus
  End If
End Sub
Private Sub fpProrate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
    If Index = 14 Then
      SendKeys "{pgdn}"
    Else
      fpRevSource(Index + 1).SetFocus
    End If
  ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Then
    fpDefDist(Index).SetFocus
  End If
'  If Index = 14 Then
'    Select Case KeyCode
'    Case vbKeyUp:
'      Me.fpProrate(13).SetFocus
'    Case vbKeyDown:
'      SendKeys "{pgdn}"
'    Case Else:
'    End Select
'  End If
End Sub
'~~~~~~~~~~~~~~~~~~~~~~End Page 2~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~Page 3 keydowns~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub fpBilDebAct_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
    fpBilCrdAct(Index).SetFocus
  ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Then
    If Index > 0 Then
      fpBilCrdAct(Index - 1).SetFocus
    Else
      SendKeys "{pgup}"
    End If
  End If
'  If Index = 0 Then
'    Select Case KeyCode
'    Case vbKeyUp:
'      SendKeys "{pgup}"
'    Case vbKeyDown, vbKeyReturn:
'      Me.fpBilCrdAct(Index).SetFocus
'    Case Else:
'    End Select
'  End If
End Sub
Private Sub fpBilCrdAct_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
    If Index = 14 Then
      SendKeys "{pgdn}"
    Else
      fpBilDebAct(Index + 1).SetFocus
    End If
  ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Then
    fpBilDebAct(Index).SetFocus
  End If
  
'  If Index = 14 Then
'    Select Case KeyCode
'    Case vbKeyUp:
'      Me.fpBilCrdAct(13).SetFocus
'    Case vbKeyDown, vbKeyReturn:
'      SendKeys "{pgdn}"
'    Case Else:
'    End Select
'  End If
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~Page 4 keydowns~~~~~~~~~~~~~~~~~~
Private Sub fpPayDebAct_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
    fpPayCrdAct(Index).SetFocus
  ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Then
    If Index > 0 Then
      fpPayCrdAct(Index - 1).SetFocus
    Else
      SendKeys "{pgup}"
    End If
  End If
  
'  If Index = 0 Then
'    Select Case KeyCode
'    Case vbKeyUp:
'      SendKeys "{pgup}"
'    Case vbKeyDown:
'      Me.fpPayDebAct(1).SetFocus
'    Case Else:
'    End Select
'  End If
End Sub
Private Sub fpPayCrdAct_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
    If Index = 14 Then
      SendKeys "{pgdn}"
    Else
      fpPayDebAct(Index + 1).SetFocus
    End If
  ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Then
    fpPayDebAct(Index).SetFocus
  End If
'  If Index = 14 Then
'    Select Case KeyCode
'    Case vbKeyUp:
'      Me.fpPayCrdAct(13).SetFocus
'    Case vbKeyDown:
'      SendKeys "{pgdn}"
'    Case Else:
'    End Select
'  End If
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~Page 5 keydowns~~~~~~~~~~~~~~~~
Private Sub fpDepDebAct_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
    fpDepCrdAct(Index).SetFocus
  ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Then
    If Index > 0 Then
      fpDepCrdAct(Index - 1).SetFocus
    Else
      SendKeys "{pgup}"
    End If
  End If
  
'  If Index = 0 Then
'    Select Case KeyCode
'    Case vbKeyUp:
'      SendKeys "{pgup}"
'    Case vbKeyDown:
'      Me.fpDepDebAct(1).SetFocus
'    Case Else:
'    End Select
'  End If
End Sub
Private Sub fpDepCrdAct_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
    If Index = 14 Then
      SendKeys "{pgdn}"
    Else
      fpDepDebAct(Index + 1).SetFocus
    End If
  ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Then
    fpDepDebAct(Index).SetFocus
  End If
 
'  If Index = 14 Then
'    Select Case KeyCode
'    Case vbKeyUp:
'      Me.fpDepCrdAct(13).SetFocus
'    Case vbKeyDown:
'      SendKeys "{pgdn}"
'    Case Else:
'    End Select
'  End If
End Sub
'~~~~~~~~~~~~~~~~~~~End of Pages 3,4,5 keydowns~~~~~~~~~~~~~~~~~~~~


