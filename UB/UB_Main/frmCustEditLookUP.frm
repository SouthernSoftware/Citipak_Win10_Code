VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmCustEditLookUP 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Existing Customer"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   1890
   ClientWidth     =   12180
   ClipControls    =   0   'False
   Icon            =   "frmCustEditLookUP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   8535
      Width           =   12180
      _ExtentX        =   21484
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7117
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "3:03 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "2/2/2011"
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
   Begin EditLib.fpText fpSearchText 
      Height          =   348
      Left            =   5160
      TabIndex        =   2
      Top             =   4608
      Width           =   3996
      _Version        =   196608
      _ExtentX        =   7048
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   0   'False
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   1
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   0   'False
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   35
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdSearch 
      Height          =   480
      Left            =   3810
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6690
      Width           =   1320
      _Version        =   131072
      _ExtentX        =   2328
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
      ButtonDesigner  =   "frmCustEditLookUP.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdExit 
      Height          =   480
      Left            =   7056
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6696
      Width           =   1332
      _Version        =   131072
      _ExtentX        =   2350
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
      ButtonDesigner  =   "frmCustEditLookUP.frx":0AA7
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdChoice 
      Height          =   480
      Left            =   5436
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6696
      Width           =   1332
      _Version        =   131072
      _ExtentX        =   2350
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
      ButtonDesigner  =   "frmCustEditLookUP.frx":0C83
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Look-Up"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   444
      Left            =   4602
      TabIndex        =   4
      Top             =   3672
      Width           =   2988
   End
   Begin VB.Label PromptLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   324
      Left            =   2820
      TabIndex        =   3
      Top             =   4632
      Width           =   2196
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2316
      Left            =   2574
      Top             =   3312
      Width           =   7044
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3210
      Top             =   1368
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edit an Existing Customer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3270
      TabIndex        =   1
      Top             =   1608
      Width           =   5652
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3210
      Top             =   1248
      Width           =   5772
   End
End
Attribute VB_Name = "frmCustEditLookUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Public DefLookUp As Integer
Dim RecNo As Long, AcctNum As Long
'Dim Multimedia As New Mmedia
Dim fromform As Form, toform As Form, codeopt As Integer, SrchDel As Integer
'codeopt used to determine if need to go back to list or search screen
'when move from said screens
'use NotDel for Delete lookup to inidicate if acct is deletable, 1 means from Delete lookup
'use srchdel- 1 for delete, 2 for final, 3 for applydeposit, 4 is for DepositCreditRemoval
'5 is for deposit refund, 6 is for deposit reversal, 7 is for CustInfo
Public Sub Wheretogo(xfrm As Form, tfrm As Form, Optional opt As Integer, Optional SDel As Integer)
  Set fromform = xfrm
  Set toform = tfrm
  If opt <> 0 Then
    codeopt = opt
  Else
    codeopt = 0
  End If
  If SDel <> 0 Then
    SrchDel = SDel
  Else
    SrchDel = 0
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If fpCmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via CustLookup by " + PWUser$
        CitiTerminate
      End If
    End If
  End If

End Sub

Private Sub fpSearchText_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpCmdChoice_Click()
  DefLookUp = DefLookUp + 1
  Call SetPromptLabel
End Sub

Private Sub fpCmdExit_Click()
  Load fromform
  DoEvents
  fromform.Show
  Unload frmCustEditLookUP
  DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      KeyCode = 0
      Call fpCmdExit_Click
    Case vbKeyF10, vbKeyReturn
      KeyCode = 0
      Call fpCmdSearch_Click
    Case vbKeyF7:
      KeyCode = 0
      Call fpCmdChoice_Click
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  DefLookUp = GetDefaultLookUP    'get the user default lookup
  Call SetPromptLabel             'set lookup prompt
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    'Me.Visible = True
    'Me.SetFocus
    'Rem this out for error in XP When relog Welcome Screensave comes up
  End If
  DoEvents
End Sub

Private Sub SetPromptLabel()

  If DefLookUp > 6 Or DefLookUp < 1 Then
    DefLookUp = 1
  End If
  Select Case DefLookUp
  Case 1:
    Me.PromptLabel = "Account Number:"
  Case 2:
    Me.PromptLabel = "Search Name:"
  Case 3:
    Me.PromptLabel = "Meter Number:"
  Case 4:
    Me.PromptLabel = "Service Address:"
  Case 5:
    Me.PromptLabel = "Location Number:"
  Case 6:
    Me.PromptLabel = "911/Other:"
  End Select

End Sub

Private Sub fpCmdSearch_Click()
  Dim LookFor As String
  LookFor$ = QPTrim$(Me.fpSearchText)
  Call LookUpList(LookFor$, DefLookUp, False, False, Me)
  'RecNo& = LookUp(LookFor$, DefLookUp, False, False, Me)
'  If RecNo& > 0 Then  'if user selected an account
'    DeActivateControls Me
'    frmInfo.Label1 = "Loading. . ."
'    frmInfo.Show
'    DoEvents
''here
'    toform.fpCustRecNo = QPTrim$(Str$(RecNo&))   'set hidden recno field on edit form
'    Load toform
'    toform.Show
'    DoEvents
'    Unload frmInfo
'  '  Unload frmCustEditLookUP
'  Else
'    Me.fpSearchText.SetFocus
'  End If
End Sub
Public Sub LookUpList(LookFor$, FindType%, ClearScrn%, ActiveOnly%, ParentForm As Form)
  
  Dim AcctNum As Long, TCnt As Integer
  ReDim UBCustRec(1) As NewUBCustRecType
  ReDim UBCustSN(1) As nUBCustReIndexRecType
  Dim UBCustRecLen As Integer, UBCustSNLen As Integer
  Dim C1Handle As Integer, R1Handle As Integer, dcnt As Long
  Dim SearchLen As Integer, AbortFlag As Integer
  Dim NumOfCust As Long, CCnt As Long, NumOfRecs As Long
  Dim UBSearchN As String, Build As String * 80
  Dim TCustName As String, IndexName As String
  Dim OK2Search As Integer, DashPos As Integer
  Dim LNum As String, Book As String, SeqN As String
  Dim SAddrFlag As Integer, AddrOKFlag As Integer
  Dim IdxRecLen As Integer, Handle As Integer
  Dim IdxFileSize As Long, IdxNumOfRecs As Long, MidRec As Long
  Dim FirstRec As Long, LastRec As Long, LastSRec As Long
  Dim BotOffSet As Long, TopOffSet As Long, FirstMatchRec As Long

  UBCustRecLen = Len(UBCustRec(1))
  
  Select Case FindType
  Case 2, 3, 4, 6:   'all but account and location lookups
    Load frmDisplayList
  Case Else:
  End Select
  
  LookFor$ = UCase$(LookFor$)
    
  Select Case FindType
  Case 1    'account lookup
    AcctNum& = Val(LookFor$)
    If AcctNum& < 1 Or AcctNum& > GetNumOfCust Then
      Load frmLookupError
      frmLookupError.Label = "Invalid Account Number!"
      frmLookupError.Show vbModal
      RecNo& = 0
    Else
      If IsDeleted(AcctNum&) Then
        Load frmLookupError
        frmLookupError.Label = "Deleted Account!"
        frmLookupError.Show vbModal
        RecNo& = 0
      Else
        RecNo& = AcctNum&
        GoTo NoNeedForList
      End If
    End If
  Case 2    'Name lookup
    If Len(LookFor$) = 0 Then
      LookFor$ = Space$(10)
    End If
    GoSub Search4Cust
    If AbortFlag Then
      GoTo ExitLookUp
    End If
    If dcnt = 0 Then
      frmLookupError.Label = "No Matching Account Found"
      frmLookupError.Show vbModal
      RecNo& = 0
    Else
      frmDisplayList.Caption = "Matching Accounts"
      frmDisplayList.Label2 = "Service Address"
      GoTo NeedList
      'RecNo& = SearchRec
    End If
  Case 3    'meter number
    If Len(LookFor$) = 0 Then
      frmLookupError.Label = "Invalid Search!"
      frmLookupError.Show vbModal
      Unload frmLookupError
      GoTo ExitLookUp
    End If
    GoSub Search4Meter
    If AbortFlag Then
      GoTo ExitLookUp
    End If

    If dcnt = 0 Then
      frmLookupError.Label = "No Matching Account Found"
      frmLookupError.Show vbModal
      Unload frmLookupError
      RecNo& = 0
    Else
      frmDisplayList.Label2 = "Meter No."
      GoTo NeedList
    End If
  Case 4    'service address
    If Len(LookFor$) = 0 Then
      frmLookupError.Label = "Invalid Search!"
      frmLookupError.Show vbModal
      Unload frmLookupError
      GoTo ExitLookUp
    End If
    
    SAddrFlag = True
    
    GoSub Search4SAddr
    If AbortFlag Then
      GoTo ExitLookUp
    End If

    If dcnt = 0 Then
      frmLookupError.Label = "No Matching Account Found"
      frmLookupError.Show vbModal
      Unload frmLookupError
      RecNo& = 0
    Else
      frmDisplayList.Label2 = "Service Address"
      GoTo NeedList
    End If
  Case 5    'Location lookup
    If AcctNum& > 0 Then
      RecNo& = AcctNum&
    End If
    OK2Search = False
    LNum$ = LookFor$
    DashPos = InStr(LNum$, "-")

    If Len(LNum$) < 2 Then  'OR DashPos <= 0 THEN
      frmLookupError.Label = "Invalid Search!"
      frmLookupError.Show vbModal
      Unload frmLookupError
      GoTo ExitLookUp
    ElseIf DashPos > 1 Then
      Book$ = FmtBook$(Left$(LNum$, DashPos - 1))
      SeqN$ = FmtSeqN$(Mid$(LNum$, DashPos + 1))
      LNum$ = Book$ + "-" + SeqN$
      OK2Search = True
    Else
      Book$ = FmtBook$(Left$(LNum$, 2))
      SeqN$ = FmtSeqN$(Mid$(LNum$, 3))
      LNum$ = Book$ + "-" + SeqN$
      OK2Search = True
    End If
    If OK2Search Then
      ParentForm.fpSearchText = LNum$
      GoSub Search4LNum
      If AcctNum& > 0 Then
        RecNo& = AcctNum&
        GoTo NoNeedForList
      ElseIf AcctNum& = 0 Then
        RecNo& = 0
        frmLookupError.Label = "No Matching Location Found"
        frmLookupError.Show vbModal
        Unload frmLookupError
      End If
    End If
  Case 6   '911 Address
    If Len(LookFor$) = 0 Then
      frmLookupError.Label = "Invalid Search!"
      frmLookupError.Show vbModal
      Unload frmLookupError
      GoTo ExitLookUp
    End If
    SAddrFlag = False
    GoSub Search4SAddr
    If AbortFlag Then
      GoTo ExitLookUp
    End If

    If dcnt = 0 Then
      frmLookupError.Label = "No Matching Account Found"
      frmLookupError.Show vbModal
      Unload frmLookupError
      RecNo& = 0
    Else
      frmDisplayList.Label2 = "911 Address"
      'frmDisplayList.Show vbModal, ParentForm
      'RecNo& = SearchRec
      GoTo NeedList
    End If
  End Select
  GoTo ExitLookUp
NoNeedForList:
      If RecNo& > 0 Then  'if user selected an account
        If SrchDel = 1 Then
          If OKDeleteCust(RecNo&) Then
            DeActivateControls Me, , True
            frmInfo.Label1 = "Loading. . ."
            frmInfo.Show
            DoEvents
          'here
            toform.fpCustRecNo = QPTrim$(Str$(RecNo&))   'set hidden recno field on edit form
            toform.Wheretogo fromform, toform, 1 'send code 1 for search screen
            Load toform
            toform.Show
            DoEvents
            Unload frmInfo
          '  Unload frmCustEditLookUP
          Else
            Me.fpSearchText.SetFocus
          End If
        ElseIf SrchDel = 2 Then
          If OKFinalCust(RecNo&) Then
            DeActivateControls Me, , True
            frmInfo.Label1 = "Loading. . ."
            frmInfo.Show
            DoEvents
          'here
            toform.fpCustRecNo = QPTrim$(Str$(RecNo&))   'set hidden recno field on edit form
            toform.Wheretogo fromform, toform, 1 'send code 1 for search screen
            Load toform
            toform.Show
            DoEvents
            Unload frmInfo
          '  Unload frmCustEditLookUP
          Else
            Me.fpSearchText.SetFocus
          End If
        ElseIf SrchDel = 3 Then
          If OKApplyDep(RecNo&) Then
            DeActivateControls Me, , True
            frmInfo.Label1 = "Loading. . ."
            frmInfo.Show
            DoEvents
          'here
            toform.fpCustRecNo = QPTrim$(Str$(RecNo&))   'set hidden recno field on edit form
            toform.Wheretogo fromform, toform, 1 'send code 1 for search screen
            Load toform
            toform.Show
            DoEvents
            Unload frmInfo
          '  Unload frmCustEditLookUP
          Else
            Me.fpSearchText.SetFocus
          End If
        ElseIf SrchDel = 4 Then
          If OKDepCreditAdj(RecNo&) Then
            DeActivateControls Me, , True
            frmInfo.Label1 = "Loading. . ."
            frmInfo.Show
            DoEvents
          'here
            toform.fpCustRecNo = QPTrim$(Str$(RecNo&))   'set hidden recno field on edit form
            toform.Wheretogo fromform, toform, 1 'send code 1 for search screen
            Load toform
            toform.Show
            DoEvents
            Unload frmInfo
          '  Unload frmCustEditLookUP
          Else
            Me.fpSearchText.SetFocus
          End If
        ElseIf SrchDel = 5 Then
          If OKDepRefund(RecNo&) Then
            DeActivateControls Me, , True
            frmInfo.Label1 = "Loading. . ."
            frmInfo.Show
            DoEvents
          'here
            toform.fpCustRecNo = QPTrim$(Str$(RecNo&))   'set hidden recno field on edit form
            toform.Wheretogo fromform, toform, 1 'send code 1 for search screen
            Load toform
            toform.Show
            DoEvents
            Unload frmInfo
          '  Unload frmCustEditLookUP
          Else
            Me.fpSearchText.SetFocus
          End If
        ElseIf SrchDel = 6 Then
          If OKDepReverse(RecNo&) Then
            DeActivateControls Me, , True
            frmInfo.Label1 = "Loading. . ."
            frmInfo.Show
            DoEvents
          'here
            toform.fpCustRecNo = QPTrim$(Str$(RecNo&))   'set hidden recno field on edit form
            toform.Wheretogo fromform, toform, 1 'send code 1 for search screen
            Load toform
            toform.Show
            DoEvents
            Unload frmInfo
          '  Unload frmCustEditLookUP
          Else
            Me.fpSearchText.SetFocus
          End If
        ElseIf SrchDel = 7 Then  'THIS IS FOR CUST INFO PRINT
          'DeActivateControls Me, , True
          frmInfo.Label1 = "Loading. . ."
          frmInfo.Show
          DoEvents
          frmReportOpt.Show 1
          If rptopt = 1 Then
         'do the graphics
            PrintCustInfo RecNo&, 1
          ElseIf rptopt = 2 Then
          'do the text
            PrintCustInfo RecNo&, 2
          End If
           Unload frmInfo
          
        Else
          DeActivateControls Me, , True
          frmInfo.Label1 = "Loading. . ."
          frmInfo.Show
          DoEvents
          'here
          toform.fpCustRecNo = QPTrim$(Str$(RecNo&))   'set hidden recno field on edit form
          toform.Wheretogo fromform, toform, 1 'send code 1 for search screen
          Load toform
          toform.Show
          DoEvents
          Unload frmInfo
        End If

      Else
        Me.fpSearchText.SetFocus
      End If
GoTo ExitLookUp
'************************************************************
NeedList:
   DeActivateControls Me, , True
   frmDisplayList.Wheretogo fromform, toform, 2, SrchDel 'code 2 for list
   frmDisplayList.Show

GoTo ExitLookUp
Search4LNum:
  
  IdxRecLen = 4 'we are using a integer
  IdxFileSize& = FileSize(UBPath$ + "UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
  
  FrmShowPctComp.Label1 = "Searching for Location"
  FrmShowPctComp.Show
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.ShowPctComp 1, 10
  
  C1Handle = FreeFile
  Open UBPath$ + "UBCUSTBK.IDX" For Random Shared As C1Handle Len = IdxRecLen
  For CCnt = 1 To IdxNumOfRecs
    Get C1Handle, CCnt, IdxBuff(CCnt)
    FrmShowPctComp.ShowPctComp CCnt, IdxNumOfRecs
  Next
  Close C1Handle
  
  SearchLen = Len(LookFor$)
  
  FirstRec = 1
  LastRec = IdxNumOfRecs
  
  BotOffSet = 0
  TopOffSet = IdxNumOfRecs
  
  C1Handle = FreeFile
  Open UBCustFile For Random Shared As C1Handle Len = UBCustRecLen
  UBLog PWUser + (" Open custfile in lookup")
  MidRec = (LastRec + FirstRec) \ 2
  
  Do
    If LastSRec = MidRec Then
      Exit Do
    End If
    LastSRec = MidRec
    Get C1Handle, IdxBuff(MidRec).RecNum, UBCustRec(1)
    UBSearchN$ = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
    If (LNum$ = UBSearchN$) And (UBCustRec(1).DelFlag = 0) Then
      If MidRec - BotOffSet > 1 Then
        MidRec = MidRec - 1
      Else
        FirstMatchRec = MidRec
      End If
    ElseIf LNum$ < UBSearchN$ Then             'lower
      TopOffSet = MidRec
      MidRec = TopOffSet - ((TopOffSet - BotOffSet) \ 2)
    Else        'higher
      BotOffSet = MidRec
      MidRec = BotOffSet + ((TopOffSet - BotOffSet) \ 2) + 1
      If MidRec = IdxNumOfRecs + 1 Then
        Exit Do
      End If
    End If
    If TopOffSet = BotOffSet Then
      Exit Do
    End If
  Loop Until FirstMatchRec
  Close C1Handle
  UBLog PWUser + (" closed custfile in lookup")
  If FirstMatchRec = 0 Then
    AcctNum& = 0
  Else
    AcctNum& = IdxBuff(FirstMatchRec).RecNum
  End If
  
  If ActiveOnly And UBCustRec(1).Status <> "A" Then
    AcctNum& = 0
  ElseIf (ActiveOnly = 1) And (UBCustRec(1).Status <> "I") Then
    AcctNum& = 0
  End If

ExitLSearch:
  Erase UBCustRec, IdxBuff
Return

'************************************************************
Search4SAddr:
  IndexName$ = TempIndexName
  BldSvcAddrsIndex LookFor$, ActiveOnly%, SAddrFlag

  ''  SortServiceAddrs frmCustEditLookUP
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize&(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
    For CCnt& = 1 To IdxNumOfRecs
      Get #Handle, CCnt&, IdxBuff(CCnt&)
    Next
    Close Handle

  UBCustRecLen = Len(UBCustRec(1))
  'NumOfCust& = GetNumOfCust&
  If SAddrFlag Then
    FrmShowPctComp.Label1 = "Searching for Service Address"
  Else
    FrmShowPctComp.Label1 = "Searching for 911 Address"
  End If
  FrmShowPctComp.Show

  C1Handle = FreeFile
  Open UBCustFile For Random Shared As C1Handle Len = UBCustRecLen
  UBLog PWUser + (" Open custfile in lookup search4addr")
  dcnt = 0
  For CCnt& = 1 To IdxNumOfRecs
    If IdxBuff(CCnt&).RecNum > 0 Then
    Get C1Handle, IdxBuff(CCnt&).RecNum, UBCustRec(1)
'    If Not UBCustRec(1).DelFlag Then
'      If (ActiveOnly = 0) Or ((ActiveOnly = True) And ((UBCustRec(1).Status = "A"))) Then
'        GoSub CheckLoadEM2
'      ElseIf (ActiveOnly = 1) And (UBCustRec(1).Status = "I") Then
        GoSub CheckLoadEM2
'      End If
'    End If
    End If
    FrmShowPctComp.ShowPctComp CCnt&, IdxNumOfRecs
    If FrmShowPctComp.Out Then
      Unload FrmShowPctComp
      AbortFlag = True
      Exit For
    End If
  Next
  Close C1Handle
  UBLog PWUser + (" Closed custfile in lookup search4addr")

Return

CheckLoadEM2:
  AddrOKFlag = False
  If SAddrFlag Then
    If InStr(UBCustRec(1).ServAddr, LookFor$) > 0 Then
      AddrOKFlag = True
    End If
  Else
    If InStr(UBCustRec(1).Addr911, LookFor$) > 0 Then
      AddrOKFlag = True
    End If
  End If
  If AddrOKFlag Then
    LSet Build$ = Left$(QPTrim$(UBCustRec(1).CustName), 30)
    If SAddrFlag Then
      Mid$(Build$, 32, 25) = Left$(QPTrim$(UBCustRec(1).ServAddr), 25)
    Else
      Mid$(Build$, 32, 25) = QPTrim$(UBCustRec(1).Addr911)
    End If
    Mid$(Build$, 60) = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
    Mid$(Build$, 74) = Chr9$ + Str$(IdxBuff(CCnt&).RecNum)
    Mid$(Build$, 71) = QPTrim$(UBCustRec(1).Status)
    frmDisplayList.fpList1.AddItem Build$
    dcnt = dcnt + 1
  End If
Return

'*************************************************************

Search4Meter:
  
  UBCustRecLen = Len(UBCustRec(1))
  NumOfCust& = GetNumOfCust&
  
  FrmShowPctComp.Label1 = "Searching for Meter Number"
  FrmShowPctComp.Show

  C1Handle = FreeFile
  Open UBCustFile For Random Shared As C1Handle Len = UBCustRecLen
  UBLog PWUser + (" Open custfile in lookup search4meter")
  dcnt = 0
  For CCnt& = 1 To NumOfCust&
    Get C1Handle, CCnt&, UBCustRec(1)
    If Not UBCustRec(1).DelFlag Then
      'IF NOT ActiveOnly OR (ActiveOnly AND (UBCustRec(1).Status = "A")) THEN
      If (ActiveOnly = 0) Or ((ActiveOnly = True) And ((UBCustRec(1).Status = "A"))) Then
        GoSub CheckEM2
      ElseIf (ActiveOnly = 1) And (UBCustRec(1).Status = "I") Then
        GoSub CheckEM2
      End If
    End If
    FrmShowPctComp.ShowPctComp CCnt&, NumOfCust&
    If FrmShowPctComp.Out Then
      Unload FrmShowPctComp
      AbortFlag = True
      Exit For
    End If
  Next
  Close C1Handle
  UBLog PWUser + (" Closed custfile in lookup search4meter")

Return
  
CheckEM2:
  For TCnt = 1 To 7
    If InStr(UBCustRec(1).LocMeters(TCnt).MtrNum, LookFor$) > 0 Then
      LSet Build$ = Left$(QPTrim$(UBCustRec(1).CustName), 30)
      Mid$(Build$, 32, 12) = QPTrim$(UBCustRec(1).LocMeters(TCnt).MtrNum)
      Mid$(Build$, 60) = QPTrim(UBCustRec(1).Book) + "-" + QPTrim(UBCustRec(1).SEQNUMB)
      Mid$(Build$, 74) = Chr9$ + Str$(CCnt&)
      Mid$(Build$, 71) = QPTrim$(UBCustRec(1).Status)
      frmDisplayList.fpList1.AddItem Build$
      dcnt = dcnt + 1
    End If
  Next
Return

'************************************************************
Search4Cust:
  UBCustSNLen = Len(UBCustSN(1))
  
  FrmShowPctComp.Label1 = "Searching Customers"
  FrmShowPctComp.Show
  
  SearchLen = Len(LookFor$)
  
  C1Handle = FreeFile
  Open UBPath$ + "UBCUSTSN.DAT" For Random Shared As C1Handle Len = UBCustSNLen
  'open short name data file
  R1Handle = FreeFile
  Open UBCustFile For Random Shared As R1Handle Len = UBCustRecLen
  'open customer data file
  UBLog PWUser + (" Open custfile in lookup search4cust")

  NumOfCust& = LOF(C1Handle) / UBCustSNLen
  
  For CCnt& = 1 To NumOfCust&
    Get C1Handle, CCnt&, UBCustSN(1)
      UBSearchN$ = Left$(UBCustSN(1).SearchName, SearchLen)
      If (LookFor$ = UBSearchN$) Then
        If Len(QPTrim$(UBCustSN(1).DelFlag)) Then GoTo DelSkip2
        If (ActiveOnly = 0) Or ((ActiveOnly = True) And ((UBCustSN(1).Status = "A"))) Then
          GoSub CustLoadEM2
        ElseIf (ActiveOnly = 1) And (UBCustSN(1).Status = "I") Then
          GoSub CustLoadEM2
        End If
      End If
DelSkip2:
    'Next
    FrmShowPctComp.ShowPctComp CCnt&, NumOfCust&
    If FrmShowPctComp.Out Then
      Unload FrmShowPctComp
      AbortFlag = True
      Exit For
    End If
    'ShowPctCompL CCnt&, NumChunks&
    'ShowSearchWheel 12, 44
  Next
  
  Close C1Handle               'close files
  Close R1Handle
  UBLog PWUser + (" Closed custfile in lookup search4cust")

Return
  
CustLoadEM2:
  
  Get R1Handle, UBCustSN(1).RecNum, UBCustRec(1)
  
  dcnt = dcnt + 1
  LSet Build$ = Left$(QPTrim$(UBCustRec(1).CustName), 26)
  Mid$(Build$, 28) = Left$(QPTrim$(UBCustRec(1).ServAddr), 30)
  Mid$(Build$, 60) = QPTrim(UBCustRec(1).Book) + "-" + QPTrim(UBCustRec(1).SEQNUMB)
  Mid$(Build$, 71) = QPTrim$(UBCustRec(1).Status)
  Mid$(Build$, 74) = Chr9$ + Str$(UBCustSN(1).RecNum)
  frmDisplayList.fpList1.AddItem Build$
  
Return
'************************************************************
'OKDeleteCust:
'  ReDim MsgText(0 To 5) As String
'  Dim FntSize As Integer, TotalBalance As Double
'  Dim M1 As String, M2 As String
'  Dim UBCustF As Integer
'  If Recno& > 0 Then
'  ReDim UBCustRec(1 To 2) As NewUBCustRecType
'  UBCustRecLen = Len(UBCustRec(1))
'  UBCustF = FreeFile
'  Open "UBCUST.DAT" For Random Shared As UBCustF Len = UBCustRecLen
'  Get UBCustF, Recno&, UBCustRec(1)
'  Close UBCustF
'
'  TotalBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
'  If TotalBalance# <> 0 Then
'    UBLog "NODELETE:" + Str$(Recno&) + " BAL:" + Str$(TotalBalance#)
'    M1$ = "This account HAS A BALANCE"
'    M2$ = "CAN NOT DELETE THIS ACCOUNT!"
'    Custok = False
'  ElseIf UBCustRec(1).DepositAmt <> 0 Then
'    UBLog "NODELETE:" + Str$(Recno&) + " DEP:" + Str$(UBCustRec(1).DepositAmt)
'    M1$ = "This account HAS A DEPOSIT"
'    M2$ = "CAN NOT DELETE THIS ACCOUNT!"
'    Custok = False
'  ElseIf UBCustRec(1).Status <> "I" Then
'    UBLog "NODELETE:" + Str$(Recno&) + " NOT INACTIVE"
'    M1$ = "This account IS NOT INACTIVE"
'    M2$ = "CAN NOT DELETE THIS ACCOUNT!"
'    Custok = False
'  Else
'    Custok = True
'  End If
'  If Custok = False Then
'    frmMsgDialog.RetLabel = "-2"
'    FntSize = frmMsgDialog.Label(2).FontSize
'    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
'    MsgText(0) = ""
'    MsgText(1) = ""
'    MsgText(2) = M1$
'    MsgText(3) = ""
'    MsgText(4) = M2$
'    MsgText(5) = ""
'    GetOKorNot MsgText(), True
'
'  End If
'End If
'Return

ExitLookUp:
End Sub
Public Sub BldSvcAddrsIndex(LookFor$, ActiveOnly%, SAddrFlag)
  Dim CustRecLen As Integer, NumOfCust As Long, IndexRecLen As Integer
  Dim CHandle As Integer, cnt As Long, lngCurLow As Long, lngCurHigh As Long
  Dim IHandle As Integer, IndexName As String, CRec As Long, dcnt As Long
  Dim CCnt As Long, AbortFlag As Boolean
  'ShowProcessingScrn "Creating " + IndexText$ + " Index"
 ' QPrintRC "    Reading Customer Records     ", 11, 25, -1
    
  FrmShowPctComp.Label1 = "Searching......."
  
  FrmShowPctComp.Show

  ReDim UBCustRec(1) As NewUBCustRecType
  CustRecLen = Len(UBCustRec(1))

  NumOfCust = GetNumOfCust&

  ReDim ServIndex(1 To 1) As UBServiceAddressIndexType
  IndexRecLen = Len(ServIndex(1))

  CHandle = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CHandle Len = CustRecLen
  UBLog PWUser + (" Open custfile in lookup on bldsvcaddrsindex")
  dcnt = 0
  For CCnt& = 1 To NumOfCust&
    Get CHandle, CCnt&, UBCustRec(1)
    If Not UBCustRec(1).DelFlag Then
      If (ActiveOnly = 0) Or ((ActiveOnly = True) And ((UBCustRec(1).Status = "A"))) Then
        If SAddrFlag Then
          If InStr(UBCustRec(1).ServAddr, LookFor$) > 0 Then
            dcnt = dcnt + 1
            GoSub getloaded
          End If
        Else
          If InStr(UBCustRec(1).Addr911, LookFor$) > 0 Then
            dcnt = dcnt + 1
            GoSub getloaded
          End If
        End If
      ElseIf (ActiveOnly = 1) And (UBCustRec(1).Status = "I") Then
        If SAddrFlag Then
          If InStr(UBCustRec(1).ServAddr, LookFor$) > 0 Then
            dcnt = dcnt + 1
            GoSub getloaded
          End If
        Else
          If InStr(UBCustRec(1).Addr911, LookFor$) > 0 Then
            dcnt = dcnt + 1
            GoSub getloaded
          End If
        End If
      End If
    End If
    FrmShowPctComp.ShowPctComp CCnt&, NumOfCust&
    If FrmShowPctComp.Out Then
      Unload FrmShowPctComp
      AbortFlag = True
      Close
      Exit Sub
    End If
  Next
Close CHandle
UBLog PWUser + (" Closed custfile in lookup on bldsvcaddrsindex")

  'QPrintRC "         Sorting Index.        ", 11, 25, -1
  lngCurLow = LBound(ServIndex)
  lngCurHigh = UBound(ServIndex)
  AddrQSort ServIndex(), lngCurLow, lngCurHigh
  'SortT ServIndex(1), NumCustRecs, 0, 16, 0, 14
  ' SortT (Elemen, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
 ' QPrintRC "      Writing Index Records      ", 11, 25, -1
  IndexName$ = TempIndexName '"UBTemp" + QPTrim(Str(OPERNUM)) + ".idx"
  KillFile IndexName$
  IHandle = FreeFile
  'FCreate IndexName$
  Open IndexName$ For Random Shared As IHandle Len = 4
  For dcnt = 1 To lngCurHigh
    CRec& = ServIndex(dcnt).RecNum
    Put IHandle, dcnt, CRec&
    'ShowPctComp dcnt, lngCurHigh                'show user percentage complet
  Next
  Close IHandle

  Erase UBCustRec, ServIndex
Exit Sub
getloaded:
    ReDim Preserve ServIndex(1 To dcnt) As UBServiceAddressIndexType
    ServIndex(dcnt).ServiceAddress = UBCustRec(1).ServAddr
    ServIndex(dcnt).RecNum = CCnt
    'ShowPctComp CCnt, NumCustRecs                'show user percentage complete
Return
End Sub


