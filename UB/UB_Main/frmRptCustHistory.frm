VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmRptCustHistory 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12192
   ClipControls    =   0   'False
   Icon            =   "frmRptCustHistory.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12192
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   336
      Left            =   0
      TabIndex        =   0
      Top             =   8532
      Width           =   12192
      _ExtentX        =   21505
      _ExtentY        =   593
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7133
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "4:39 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "5/22/2003"
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
         Size            =   10.8
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
      CaretInsert     =   0
      CaretOverWrite  =   3
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
      Left            =   6798
      TabIndex        =   5
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
      ButtonDesigner  =   "frmRptCustHistory.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdExit 
      Height          =   480
      Left            =   4056
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
      ButtonDesigner  =   "frmRptCustHistory.frx":0AA7
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdChoice 
      Height          =   480
      Left            =   5430
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
      ButtonDesigner  =   "frmRptCustHistory.frx":0C83
   End
   Begin EditLib.fpBoolean fpDetailFlag 
      Height          =   348
      Left            =   5160
      TabIndex        =   8
      Top             =   5112
      Width           =   324
      _Version        =   196608
      _ExtentX        =   572
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "N"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   0
      ToggleTrue      =   "Yy"
      TextTrue        =   "Y"
      Value           =   1
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "N"
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin VB.Label DetailLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Detail:"
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
      Height          =   324
      Left            =   4260
      TabIndex        =   9
      Top             =   5136
      Width           =   756
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Look-Up"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
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
         Size            =   10.8
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
      Left            =   2592
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
      Left            =   3870
      TabIndex        =   1
      Top             =   1608
      Width           =   4452
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
Attribute VB_Name = "frmRptCustHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim DefLookUp As Integer
Dim RecNo As Long, AcctNum As Long
Dim ConsumpFlag As String
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

Property Get RptType() As Integer
  RptType = ConsumpFlag
End Property

Property Let RptType(ByVal NewRptType As Integer)
  ConsumpFlag = NewRptType
End Property

Private Sub fpCmdChoice_Click()
  DefLookUp = DefLookUp + 1
  Call SetPromptLabel
End Sub

Private Sub fpCmdExit_Click()
  Load frmUBCustMenu
  DoEvents
  frmUBCustMenu.Show
  Unload frmRptCustHistory
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
  StatusBar1.Panels.Item(1).Text = TownName$
  DefLookUp = GetDefaultLookUP    'get the user default lookup
  Call SetPromptLabel             'set lookup prompt
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
  
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
  'DeActivateControls Me
  RecNo& = LookUp(LookFor$, DefLookUp, False, False, Me)
'  ActivateControls Me
  If RecNo& > 0 Then
    frmLoadingRpt.Show
    DoEvents
    If ConsumpFlag Then
      Call CustConsumpHistRpt
    Else
      Call CustTransHistoryRpt
    End If
    Unload frmLoadingRpt
  Else
    Me.fpSearchText.SetFocus
  End If
 
End Sub
Private Sub LookUpList(LookFor$, FindType%, ClearScrn%, ActiveOnly%, ParentForm As Form)
  
  Dim AcctNum As Long, TCnt As Integer
  ReDim UBCustRec(1) As NewUBCustRecType
  ReDim UBCustSN(1) As nUBCustReIndexRecType
  Dim UBCustRecLen As Integer, UBCustSNLen As Integer
  Dim C1Handle As Integer, R1Handle As Integer, DCnt As Integer
  Dim SearchLen As Integer, AbortFlag As Integer
  Dim NumOfCust As Long, CCnt As Long
  Dim UBSearchN As String, Build As String * 80
  Dim TCustName As String
  Dim OK2Search As Integer, DashPos As Integer
  Dim LNum As String, Book As String, SeqN As String
  Dim SAddrFlag As Integer, AddrOKFlag As Integer
  Dim IdxRecLen As Integer
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
    If DCnt = 0 Then
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

    If DCnt = 0 Then
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

    If DCnt = 0 Then
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

    If DCnt = 0 Then
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
          DeActivateControls Me
          frmInfo.Label1 = "Loading. . ."
          frmInfo.Show
          DoEvents
        'here
          toform.fpCustRecNo = QPTrim$(Str$(RecNo&))   'set hidden recno field on edit form
          toform.Wheretogo fromform, toform, 1
          Load toform
          toform.Show
          DoEvents
          Unload frmInfo
        '  Unload frmCustEditLookUP
        Else
          Me.fpSearchText.SetFocus
        End If
GoTo ExitLookUp
'************************************************************
NeedList:
   DeActivateControls Me
   frmDisplayList.Wheretogo fromform, toform, 0
   frmDisplayList.Show

GoTo ExitLookUp
Search4LNum:
  
  IdxRecLen = 4 'we are using a integer
  IdxFileSize& = FileSize("UBCUSTBK.IDX")
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
  UBCustRecLen = Len(UBCustRec(1))
  NumOfCust& = GetNumOfCust&
  If SAddrFlag Then
    FrmShowPctComp.Label1 = "Searching for Service Address"
  Else
    FrmShowPctComp.Label1 = "Searching for 911 Address"
  End If
  FrmShowPctComp.Show

  C1Handle = FreeFile
  Open UBCustFile For Random Shared As C1Handle Len = UBCustRecLen
  
  DCnt = 0
  For CCnt& = 1 To NumOfCust&
    Get C1Handle, CCnt&, UBCustRec(1)
    If Not UBCustRec(1).DelFlag Then
      If (ActiveOnly = 0) Or ((ActiveOnly = True) And ((UBCustRec(1).Status = "A"))) Then
        GoSub CheckLoadEM2
      ElseIf (ActiveOnly = 1) And (UBCustRec(1).Status = "I") Then
        GoSub CheckLoadEM2
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

Return

CheckLoadEM2:
  AddrOKFlag = False
  If SAddrFlag Then
    If InStr(UBCustRec(1).SERVADDR, LookFor$) > 0 Then
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
      Mid$(Build$, 32, 25) = Left$(QPTrim$(UBCustRec(1).SERVADDR), 25)
    Else
      Mid$(Build$, 32, 25) = QPTrim$(UBCustRec(1).Addr911)
    End If
    Mid$(Build$, 60) = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
    Mid$(Build$, 75) = Chr9$ + Str$(CCnt&)
    Mid$(Build$, 71) = QPTrim$(UBCustRec(1).Status)
    frmDisplayList.fpList1.AddItem Build$
    DCnt = DCnt + 1
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
  
  DCnt = 0
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
  
Return
  
CheckEM2:
  For TCnt = 1 To 7
    If InStr(UBCustRec(1).LocMeters(TCnt).MtrNum, LookFor$) > 0 Then
      LSet Build$ = Left$(QPTrim$(UBCustRec(1).CustName), 30)
      Mid$(Build$, 32, 12) = QPTrim$(UBCustRec(1).LocMeters(TCnt).MtrNum)
      Mid$(Build$, 60) = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
      Mid$(Build$, 75) = Chr9$ + Str$(CCnt&)
      Mid$(Build$, 71) = QPTrim$(UBCustRec(1).Status)
      frmDisplayList.fpList1.AddItem Build$
      DCnt = DCnt + 1
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
  
Return
  
CustLoadEM2:
  
  Get R1Handle, UBCustSN(1).RecNum, UBCustRec(1)
  
  DCnt = DCnt + 1
  LSet Build$ = Left$(QPTrim$(UBCustRec(1).CustName), 26)
  Mid$(Build$, 28) = Left$(QPTrim$(UBCustRec(1).SERVADDR), 30)
  Mid$(Build$, 60) = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
  Mid$(Build$, 71) = QPTrim$(UBCustRec(1).Status)
  Mid$(Build$, 75) = Chr9$ + Str$(UBCustSN(1).RecNum)
  frmDisplayList.fpList1.AddItem Build$
  
Return
'************************************************************

ExitLookUp:
End Sub

'***************************************
Private Sub CustConsumpHistRpt()
  Dim Dash80 As String, F As String
  ReDim UBCustRec(1) As NewUBCustRecType
  ReDim UBTranRec(1) As UBTransRecType
  ReDim UBSetUpRec(1) As UBSetupRecType
  Dim DidCnt As Integer, CCnt As Integer
  Dim UBCustRecLen As Integer, UBTranRecLen As Integer
  Dim UBSetupLen As Integer, RevCnt As Integer, Rev2Flag As Integer
  Dim UBCust As Integer, UBRpt As Integer, UBTran As Integer
  Dim TroyFlag As Integer, AbortFlag As Integer
    
  Dim MCnt As Integer
  Dim ThisTrans As Long, MaxMeterAmt As Long
  Dim MeterType As String, ReportFile As String
  Dim EstCnt As Integer, CubMeter As Integer
  Dim MTRMulti As Double, MeterConsp As Double, TotalConsp As Double
    
  Dim DidAMeter As Integer, EstFlag As Integer
  Dim MtrCnt As Integer

  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  
  ReportFile$ = UBPath$ + "UBCONSMP.RPT"
  
  If InStr(UBSetUpRec(1).UTILNAME, "TROY") > 0 Then
    TroyFlag = True
  End If

  If RecNo& = 0 Then
    GoTo ExitConsumpHist
  End If

  Dash80$ = String$(80, "-")
    
  UBTranRecLen = Len(UBTranRec(1))
  UBCustRecLen = Len(UBCustRec(1))
  
  UBCust = FreeFile
  Open UBCustFile For Random Shared As UBCust Len = UBCustRecLen
  Get #UBCust, RecNo&, UBCustRec(1)
  Close UBCust
  
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt

  UBTran = FreeFile
  Open "UBTRANS.DAT" For Random Shared As UBTran Len = UBTranRecLen

  GoSub DoConsRptHeader

  ThisTrans& = UBCustRec(1).LastTrans

  Do While ThisTrans& > 0
    Get #UBTran, ThisTrans&, UBTranRec(1)
    If UBTranRec(1).TransType = TranUtilityBill Then
      GoSub PrintConsDetail
      DidCnt = DidCnt + 1
      If DidCnt = 12 Then
        Exit Do
      End If
    End If
    ThisTrans& = UBTranRec(1).PrevTrans
  Loop
  GoSub DoConsFooter

  Close

  If Not AbortFlag Then
    ViewPrint ReportFile$, "Customer Consumption Report."
    'PrintRptFile "Customer Consumption Report.", "UBCONSMP.RPT", 1, RetCode, EntryPoint
  End If


ExitConsumpHist:
Exit Sub

PrintConsDetail:
  
  DidAMeter = False
  EstFlag = False
  For EstCnt = 1 To 7
    If UBTranRec(1).ESTREAD(EstCnt) = "Y" Then
      EstFlag = True
      Exit For
    End If
  Next
  For MtrCnt = 1 To 7
    MTRMulti# = 0
    For MCnt = 1 To 7
      If UBTranRec(1).MtrTypes(MtrCnt) = GetCustMeterType%(UBCustRec(), MCnt) Then
        MTRMulti# = UBCustRec(1).LocMeters(MCnt).MTRMulti
        Exit For
      End If
    Next
    If MTRMulti# = 0 Then
      If TroyFlag Then
        MTRMulti# = 100
      Else
        MTRMulti# = 1
      End If
    End If

    If UBTranRec(1).MtrTypes(MtrCnt) <> 0 Then
      DidAMeter = True
      Select Case UBTranRec(1).MtrTypes(MtrCnt)
      Case MtrWaterOnly
        MeterType$ = "      Water"
        F$ = "W"
      Case MtrSewerOnly
        MeterType$ = "      Sewer"
        F$ = "S"
      Case MtrCombined
        MeterType$ = "Water/Sewer"
        F$ = "C"
      Case MtrElectric
        MeterType$ = "   Electric"
        F$ = "E"
      Case MtrDemand
        MeterType$ = " D Electric"
        F$ = "D"
      Case MtrGas
        MeterType$ = "  Gas Meter"
        F$ = "G"
      Case MtrTouchRead
        MeterType$ = " Touch Read"
        F$ = "T"
      Case MtrLightsService
        MeterType$ = "  L Service"
      Case Else
        MeterType$ = "  ?????????"
      End Select
      For CCnt = 1 To 7
        If UBCustRec(1).LocMeters(CCnt).MTRType = F$ Then
          If UBCustRec(1).LocMeters(CCnt).MTRUnit = "C" Then
            CubMeter = True
          Else
            CubMeter = False
          End If
          Exit For
        End If
      Next
      GoSub PrintThisMeter
    End If
  Next
  If Not DidAMeter Then
    MeterType$ = "        "
    MtrCnt = 1
    GoSub PrintThisMeter
  End If

Return

PrintThisMeter:

  Print #UBRpt, Num2Date(UBTranRec(1).TransDate);
  If EstFlag Then
    Print #UBRpt, "*E";
  End If
  Print #UBRpt, Tab(19); MeterType$;
  Print #UBRpt, Tab(34); Using$("##########", UBTranRec(1).CurRead(MtrCnt));
  Print #UBRpt, Tab(46); Using$("##########", UBTranRec(1).PrevRead(MtrCnt));
  MeterConsp# = UBTranRec(1).CurRead(MtrCnt) - UBTranRec(1).PrevRead(MtrCnt)
  If MeterConsp# < 0 Then
    MaxMeterAmt& = 10& ^ (Len(Str$(UBTranRec(1).PrevRead(MtrCnt))) - 1)
    MeterConsp# = (MaxMeterAmt& - UBTranRec(1).PrevRead(MtrCnt)) + UBTranRec(1).CurRead(MtrCnt)
  End If
  MeterConsp# = MeterConsp# * MTRMulti#
  If CubMeter Then
    MeterConsp# = MeterConsp# * 7.481
  End If
  Print #UBRpt, Tab(56); Using$("##########", MeterConsp#);
  If UBTranRec(1).ReadDate <= 0 Then
    Print #UBRpt, "     ??-??-????"
  Else
    Print #UBRpt, "     "; Num2Date$(UBTranRec(1).ReadDate) '; "!"; UBTranRec(1).EstRead(MtrCnt); "!"
  End If

  TotalConsp# = TotalConsp# + MeterConsp#

Return

DoConsRptHeader:
  Print #UBRpt, Tab(28); "Consumption History Report. "
  Print #UBRpt,
  Print #UBRpt, "Customer: "; UBCustRec(1).CustName; Tab(57); "Report Date: "; Date$
  Print #UBRpt,
  Print #UBRpt, "Transaction                         Current   Previous"
  Print #UBRpt, "   Date            Meter Type       Reading    Reading       Usage    ReadDate"
  Print #UBRpt, Dash80$
Return

DoConsFooter:
  If DidCnt > 0 Then
    Print #UBRpt, Dash80$
    Print #UBRpt, "Average Consumption: "; Using$("#########", TotalConsp# / DidCnt)
  Else
    Print #UBRpt, "NO TRANSACTIONS!!!"
    Print #UBRpt, Dash80$
  End If
Return
End Sub

'**************************************

Private Sub CustTransHistoryRpt()
  Dim t As String, Dash80 As String
  ReDim UBCustRec(1) As NewUBCustRecType
  ReDim UBTranRec(1) As UBTransRecType
  ReDim UBSetUpRec(1) As UBSetupRecType
  
  ReDim TotalConsump(1 To 7) As Long
  ReDim DidCnt(1 To 7) As Integer
    
  Dim UBCustRecLen As Integer, UBTranRecLen As Integer
  Dim UBSetupLen As Integer, RevCnt As Integer, Rev2Flag As Integer
  Dim TempRev As String, ReportFile As String
  Dim RevText1 As String, RevText2 As String
  Dim UBCust As Integer, UBRpt As Integer, UBTran As Integer
  Dim Cubic As Integer, MTCnt As Integer, NumOfRevs As Integer
  Dim ThisTrans As Long, MeterConsp As Long, MaxMeterAmt As Long
  Dim FirstTrans As Integer, TYear As Integer
  Dim LastDate As String, MeterType As String
  Dim PDate As Integer, AbortFlag As Integer
  Dim DidEst As Integer, EstCnt As Integer
  Dim DetailFlag As Integer, DidAMeter As Integer
  Dim MtrCnt As Integer, WhatMtrCNT As Integer
  Dim PrintedOne As Integer, TabStop As Integer
  Dim RevOffset As Integer
  
  ReportFile$ = UBPath$ + "UBTRAHIS.RPT"
  DetailFlag = frmRptCustHistory.fpDetailFlag.Text = "Y"
  
  t$ = Space$(10)
  MaxLines = 40

  Dash80$ = String$(80, "-")
  UBCustRecLen = Len(UBCustRec(1))
  UBTranRecLen = Len(UBTranRec(1))

  LoadUBSetUpFile UBSetUpRec(), UBSetupLen

  For RevCnt = 1 To MaxRevsCnt
    TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(RevCnt).REVNAME)
    If Len(TempRev$) = 0 Then
      NumOfRevs = RevCnt - 1
      Exit For
    Else
      RSet t$ = QPTrim$(Left$(TempRev$, 8))
      If RevCnt <= 8 Then
        RevText1$ = RevText1$ + t$
      Else
        RevText2$ = RevText2$ + t$
      End If
    End If
  Next

  If Len(QPTrim$(RevText2$)) > 0 Then
    Rev2Flag = True
  End If
  
  UBCust = FreeFile
  Open UBCustFile For Random Shared As UBCust Len = UBCustRecLen
  Get #UBCust, RecNo&, UBCustRec(1)
  Close UBCust

  For MTCnt = 1 To 7
    If UBCustRec(1).LocMeters(MTCnt).MTRUnit = "C" Then
      Cubic = True
      Exit For
    End If
  Next

  UBRpt = FreeFile
  Open ReportFile For Output As UBRpt

  UBTran = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTran Len = UBTranRecLen

  GoSub DOTranHistHeader

  ThisTrans& = UBCustRec(1).LastTrans
  
  FirstTrans = True

  Do While ThisTrans& > 0
    Get #UBTran, ThisTrans&, UBTranRec(1)
      If FirstTrans Then
        LastDate$ = Num2Date$(UBTranRec(1).TransDate)
        TYear = Val(Right$(LastDate$, 4))
        PDate = Date2Num(Left$(LastDate$, 3) + "01-" + QPTrim$(Str$(TYear - 1)))
        FirstTrans = False
      End If
      
      GoSub DOTransDetail
      Print #UBRpt, Dash80$
      Linecnt = Linecnt + 1
      If Linecnt > MaxLines Then
        Print #UBRpt, FF$
        GoSub DOTranHistHeader
      End If
    ThisTrans& = UBTranRec(1).PrevTrans
  Loop
  GoSub DOTranHistFooter

  Close
  If Not AbortFlag Then
    ViewPrint ReportFile$, "Customer Transaction Report."
  End If

ExitTransHist:
Exit Sub

DOTransDetail:
  Print #UBRpt, Num2Date(UBTranRec(1).TransDate);

  Select Case UBTranRec(1).TransType
    Case TranUtilityBill, TranUtilityBill + 100
      DidEst = False
      For EstCnt = 1 To 7
        If UBTranRec(1).ESTREAD(EstCnt) = "Y" Then
          DidEst = True
          Exit For
        End If
      Next

      Print #UBRpt, Tab(16); "Utility Bill";
      If DidEst Then
        Print #UBRpt, "*e";
      End If
      If DetailFlag Then
        Print #UBRpt, Tab(31); Num2Date$(UBTranRec(1).ReadDate); Tab(43); Num2Date$(UBTranRec(1).PrevDate);
      End If
      Print #UBRpt, Tab(57); Using("$#####.##", UBTranRec(1).TransAmt); Tab(71); Using("$#####.##", UBTranRec(1).RunBalance);
      If DetailFlag Then
        GoSub DoMtrDetail
        GoSub PrintRevDetail
      Else
        Print #UBRpt,
      End If
    Case TranLateCharge, TranLateCharge + 100
      Print #UBRpt, Tab(16); "Late Charge"; Tab(57); Using$("$#####.##", UBTranRec(1).TransAmt); Tab(71); Using("$#####.##", UBTranRec(1).RunBalance)
      If DetailFlag Then
        GoSub PrintRevDetail
      End If
    Case TranReconnectFee, TranReconnectFee + 100
      Print #UBRpt, Tab(16); "Reconnect Fee";
      If DetailFlag Then
        GoSub PrintRevDetail
      End If
    Case TranBillPayment, TranBillPayment + 100
      Print #UBRpt, Tab(16); "Bill Payment"; Tab(57); Using$("$#####.##", UBTranRec(1).TransAmt); Tab(71); Using("$#####.##", UBTranRec(1).RunBalance)
      If DetailFlag Then
        GoSub PrintRevDetail
      End If
    Case TranAppliedDeposit, TranAppliedDeposit + 100
      Print #UBRpt, Tab(16); "Applied Deposit"; Tab(57); Using$("$#####.##", UBTranRec(1).TransAmt); Tab(67); Using("$#####.##", UBTranRec(1).RunBalance)
      If DetailFlag Then
        GoSub PrintRevDetail
      End If
    Case TranPenaltyCharge, TranPenaltyCharge + 100
      Print #UBRpt, Tab(16); "Penalty Charge"; Tab(57); Using$("$#####.##", UBTranRec(1).TransAmt); Tab(71); Using("$#####.##", UBTranRec(1).RunBalance)
      If DetailFlag Then
        GoSub PrintRevDetail
      End If
    Case TranDepositPayment, TranDepositPayment + 100
      Print #UBRpt, Tab(16); "Deposit Payment"; Tab(57); Using$("$#####.##", UBTranRec(1).TransAmt); Tab(71); Using("$#####.##", UBTranRec(1).RunBalance)
      If DetailFlag Then
        GoSub PrintRevDetail
      End If
    Case TranDraftPayment, TranDraftPayment + 100
      Print #UBRpt, Tab(16); "Draft Payment";
      If DetailFlag Then
        GoSub PrintRevDetail
      End If
    Case TranRefundDeposit, TranRefundDeposit + 100
      Print #UBRpt, Tab(16); "Refund Deposit"; Tab(57); Using$("$#####.##", UBTranRec(1).TransAmt); Tab(71); Using("$#####.##", UBTranRec(1).RunBalance)
    Case TranBeginBalance, TranBeginBalance + 100
      Print #UBRpt, Tab(16); "Beginning Balance";
    Case TranUpwardAdjustment, TranUpwardAdjustment + 100
      Print #UBRpt, Tab(16); "UP Adjustment  " + Left$(UBTranRec(1).BillMsg, 25); Tab(57); Using$("$#####.##", UBTranRec(1).TransAmt); Tab(71); Using("$#####.##", UBTranRec(1).RunBalance)
      If DetailFlag Then
        GoSub PrintRevDetail
      End If
    Case TranDownwardAdjustment, TranDownwardAdjustment + 100
      Print #UBRpt, Tab(16); "DN Adjustment  " + Left$(UBTranRec(1).BillMsg, 25); Tab(57); Using$("$#####.##", UBTranRec(1).TransAmt); Tab(71); Using("$#####.##", UBTranRec(1).RunBalance)
      If DetailFlag Then
        GoSub PrintRevDetail
      End If
    Case TranMiscPayment, TranMiscPayment + 100
      Print #UBRpt, Tab(16); "Misc Payment"
      If DetailFlag Then
        GoSub PrintRevDetail
      End If
  End Select
skipit:
Return

DoMtrDetail:
  DidAMeter = False
  For MtrCnt = 1 To 7
    If UBTranRec(1).MtrTypes(MtrCnt) > 0 Then
      DidAMeter = True
      Select Case UBTranRec(1).MtrTypes(MtrCnt)
      Case MtrWaterOnly
        MeterType$ = "      Water"
      Case MtrSewerOnly
        MeterType$ = "      Sewer"
      Case MtrCombined
        MeterType$ = "Water/Sewer"
      Case MtrElectric
        MeterType$ = "   Electric"
      Case MtrDemand
        MeterType$ = " D Electric"
      Case MtrGas
        MeterType$ = "  Gas Meter"
      Case MtrTouchRead
        MeterType$ = " Touch Read"
      Case MtrLightsService
        MeterType$ = "  L Service"
      End Select
      WhatMtrCNT = UBTranRec(1).MtrTypes(MtrCnt)
      If WhatMtrCNT = 0 Then
        WhatMtrCNT = 1
      End If
      GoSub PrintMtrDetail
    End If
  Next
  If Not DidAMeter Then
    MeterType$ = "        "
    MtrCnt = 1
    GoSub PrintMtrDetail
  End If
Return

PrintMtrDetail:
  Print #UBRpt, Tab(16); MeterType$;
  Print #UBRpt, Tab(31); Using("##########", UBTranRec(1).CurRead(MtrCnt));
  Print #UBRpt, Tab(43); Using("##########", UBTranRec(1).PrevRead(MtrCnt));
  MeterConsp& = UBTranRec(1).CurRead(MtrCnt) - UBTranRec(1).PrevRead(MtrCnt)
  If MeterConsp& < 0 Then
    MaxMeterAmt& = 10& ^ (Len(Str$(UBTranRec(1).PrevRead(MtrCnt))) - 1)
    MeterConsp& = (MaxMeterAmt& - UBTranRec(1).PrevRead(MtrCnt)) + UBTranRec(1).CurRead(MtrCnt)
  End If
  If Cubic Then
    MeterConsp& = MeterConsp& * 7.481
  End If
  Print #UBRpt, Tab(57); Using$("##########", MeterConsp&)
  If DidAMeter Then
    TotalConsump(WhatMtrCNT) = TotalConsump(WhatMtrCNT) + MeterConsp&
    DidCnt(WhatMtrCNT) = DidCnt(WhatMtrCNT) + 1
  End If
  Linecnt = Linecnt + 1
Return

PrintRevDetail:
    PrintedOne = False
    For RevCnt = 0 To 7
      If UBTranRec(1).RevAmt(RevCnt + 1) <> 0 Then
        PrintedOne = True
        TabStop = (RevCnt * 10) + 1
        Print #UBRpt, Tab(TabStop); Using$("#######.##", UBTranRec(1).RevAmt(RevCnt + 1));
      End If
    Next
    If PrintedOne Then
      Print #UBRpt,
      Linecnt = Linecnt + 1
    End If
    RevOffset = 7
    PrintedOne = False
    For RevCnt = 0 To 6
      If UBTranRec(1).RevAmt(RevCnt + 1 + RevOffset) <> 0 Then
        PrintedOne = True
        TabStop = (RevCnt * 10) + 1
        Print #UBRpt, Tab(TabStop); Using$("#######.##", UBTranRec(1).RevAmt(RevCnt + 1 + RevOffset));
      End If
    Next
    If PrintedOne Then
      Print #UBRpt,
      Linecnt = Linecnt + 1
    End If

Return

DOTranHistHeader:
  Linecnt = 7
  Print #UBRpt, Tab(28); "Transaction History Report. "
  Print #UBRpt, "Customer: "; UBCustRec(1).CustName; Tab(57); "Report Date: "; Date$
  If DetailFlag Then
    Print #UBRpt, " Account:"; RecNo&
    Print #UBRpt, "Ser Addr: "; UBCustRec(1).SERVADDR
    Print #UBRpt, "Location: "; QPTrim$(UBCustRec(1).Book); "-"; QPTrim$(UBCustRec(1).SEQNUMB)
    Linecnt = Linecnt + 2
    For MtrCnt = 1 To 7
      If Len(QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)) > 0 Then
        Print #UBRpt, Tab(6); "Mtr# "; QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
        Linecnt = Linecnt + 1
      End If
    Next
  End If
  Print #UBRpt,
  If DetailFlag Then
    Print #UBRpt, "Trans Date     Trans Type     Cur.Date     Pre.Date      TR Amount      Balance"
  Else
    Print #UBRpt, "Trans Date     Trans Type                                TR Amount      Balance"
  End If
  If DetailFlag Then
    Print #UBRpt, "               Meter Type     Cur.Read     Pre.Read       Usage"
  End If
  If DetailFlag Then
    Print #UBRpt, RevText1$
    If Rev2Flag Then
      Print #UBRpt, RevText2$
      Linecnt = 8
    End If
  Else
    Linecnt = 5
  End If
  Print #UBRpt, Dash80$
Return

DOTranHistFooter:
  If FirstTrans Then
    Print #UBRpt, "NO TRANSACTIONS!!!"
    Print #UBRpt, Dash80$
  End If
  For MtrCnt = 1 To 7
    If DidCnt(MtrCnt) > 0 Then
      Print #UBRpt, "Average Consumption: "; Using$("#########", TotalConsump(MtrCnt) / DidCnt(MtrCnt))
    End If
  Next
  Print #UBRpt, FF$
Return

End Sub


