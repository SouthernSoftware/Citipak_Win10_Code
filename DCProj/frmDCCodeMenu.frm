VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDCCodeMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Decal Code Maintenance "
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   Icon            =   "frmDCCodeMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
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
            TextSave        =   "5:07 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "11/14/2005"
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
   Begin fpBtnAtlLibCtl.fpBtn cmdAddCode 
      Height          =   492
      Left            =   3984
      TabIndex        =   0
      Top             =   3024
      Width           =   4260
      _Version        =   131072
      _ExtentX        =   7514
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmDCCodeMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEditCode 
      Height          =   492
      Left            =   3984
      TabIndex        =   1
      Top             =   3828
      Width           =   4260
      _Version        =   131072
      _ExtentX        =   7514
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmDCCodeMenu.frx":0AB4
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdListDecals 
      Height          =   492
      Left            =   3984
      TabIndex        =   2
      Top             =   4620
      Width           =   4260
      _Version        =   131072
      _ExtentX        =   7514
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmDCCodeMenu.frx":0CA4
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExitDC 
      Height          =   492
      Left            =   3984
      TabIndex        =   3
      Top             =   5424
      Width           =   4260
      _Version        =   131072
      _ExtentX        =   7514
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmDCCodeMenu.frx":0E94
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   3348
      X2              =   3348
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DECAL CODE MAINTENANCE"
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
      TabIndex        =   4
      Top             =   1176
      Width           =   5292
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1788
      Top             =   744
      Width           =   8652
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
      FillColor       =   &H00D0D0D0&
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
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8988
      Top             =   2064
      Width           =   732
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
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2376
      Top             =   1824
      Width           =   996
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
      BackColor       =   &H00E0E0E0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2508
      Top             =   2064
      Width           =   732
   End
End
Attribute VB_Name = "frmDCCodeMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
'Dim FormOver As clsFormOverRider
Private Temp_Class As Resize_Class

Private Sub cmdAddCode_Click()
  frmCodeAddEdit.Wheretogo frmDCCodeMenu, frmDCCodeMenu, 0
  'frmCodeAddEdit.SetScreen
  frmCodeAddEdit.Show
  DoEvents
  Unload Me
'
'  Load frmCodeAddEdit
'  DoEvents
'  frmCodeAddEdit.fpCodeRecNo = 0
'  frmCodeAddEdit.Show
'  Unload Me
End Sub

Private Sub cmdEditCode_Click()
  Dim FntSize As Integer
  ReDim MsgText(0 To 5) As String
  If Not Exist(DCPath$ + "DCCODE.DAT") Then
    frmMsgDialog.RetLabel = "-2"
    DCLog "ERROR: NO DecalCode FILE"
    FntSize = frmMsgDialog.Label(3).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = "NO DECAL CODE FILE!"
    MsgText(3) = ""
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
    Exit Sub
  Else
'    Load frmCodeAddEdit
'    DoEvents
'    frmCodeAddEdit.fpCodeRecNo = -1
'    frmCodeAddEdit.Show
'    Unload Me
  frmCodeAddEdit.Wheretogo frmDCCodeMenu, frmDCCodeMenu, 1
  
  DoEvents
  frmCodeAddEdit.Show
  'frmCodeAddEdit.SetScreen
  Unload Me
 End If

End Sub

Private Sub cmdListDecals_Click()
    frmReportOpt.Show 1
    DeActivateControls Me
    If rptopt = 1 Then
     PrintCodeListing rptopt, frmDCCodeMenu
    ElseIf rptopt = 2 Then
     PrintCodeListing rptopt, frmDCCodeMenu
     ActivateControls Me
    Else
      ActivateControls Me
    End If
End Sub

'LevelPass 1 is Full Access, 2 is Payments, 3 is Reports Only
Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  Me.HelpContextID = hlpDecalCategory
'  Refresh
'  DoEvents
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    If cmdExitDC.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        Call DCLog("Close via DC Code Menu" + PWUser$)
        CitiTerminate
      End If
    End If
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    'Me.Visible = True
    'Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub cmdExitDC_Click()
  Load frmDCMainMenu
  DoEvents
  frmDCMainMenu.Show
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExitDC_Click
      KeyCode = 0
    Case vbKeyHome
      cmdAddCode.SetFocus
    Case vbKeyEnd
      cmdExitDC.SetFocus
    Case Else:
  End Select
End Sub

Public Sub PrintCodeListing(rptopt, Optional xx As Form)
  Dim DCCodeRecLen As Integer, NumCodeRecs As Integer
  Dim NumPrinted As Integer, graphicflag As Boolean
  Dim RCnt As Integer, cnt As Integer
  Dim DCFile As Integer, RPTFile As Integer
  Dim ReportFile As String, ToPrint As String
  ReDim DCCodeRec(1) As DCCatCodeRecType
  Dim Dash80 As String * 78
  DCCodeRecLen = Len(DCCodeRec(1))
  NumCodeRecs = FileSize(DCPath + "DCCODE.DAT") \ DCCodeRecLen
  If rptopt = 1 Then
    graphicflag = True
  Else
    graphicflag = False
  End If
  If NumCodeRecs = 0 Then
    GoTo ExitCodeListing
  End If
  
  Dash80$ = String$(78, "-")
  
  NumPrinted = 0

  FrmShowPctComp.Label1 = "Creating Decal Code Listing."
  FrmShowPctComp.Show , Me

  ReportFile$ = DCPath + "CodeLIST.RPT"
  
  DCFile = FreeFile
  Open DCPath + "DCCODE.DAT" For Random Shared As DCFile Len = DCCodeRecLen
  
  RPTFile = FreeFile
  Open ReportFile$ For Output As RPTFile
  If graphicflag Then
    GoSub GraphicCodeList
  Else
    GoSub PrintCodeHeader
    For cnt = 1 To NumCodeRecs
      Get DCFile, cnt, DCCodeRec(1)
      If NumPrinted = 50 Then
        Print #RPTFile, Dash80$
        Print #RPTFile, Chr$(12)
        GoSub PrintCodeHeader
      End If
      Print #RPTFile, QPTrim$(DCCodeRec(1).CATCODE);
      Print #RPTFile, Tab(5); Mid$(DCCodeRec(1).CODEDESC, 1, 15);
      Print #RPTFile, Tab(25); Using$("#######.##", Str$(DCCodeRec(1).Fee));
      Print #RPTFile, Tab(38); QPTrim$(DCCodeRec(1).CASHACCT);
      Print #RPTFile, Tab(55); QPTrim$(DCCodeRec(1).REVGLNUM);
      If QPTrim$(DCCodeRec(1).InactiveFlag) = "Y" Then
        Print #RPTFile, Tab(70); QPTrim$(DCCodeRec(1).InactiveFlag)
      Else
        Print #RPTFile, Tab(70); "N"
      End If
      NumPrinted = NumPrinted + 1
      FrmShowPctComp.ShowPctComp cnt, NumCodeRecs
  
    Next
    Print #RPTFile, Dash80$
    Print #RPTFile, Tab(2); "Total Codes "; Str(NumPrinted)
    Print #RPTFile, Chr$(12)
    Close
  End If
  Erase DCCodeRec
  DoEvents
  If graphicflag Then
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom xx
    ARptCodeList.txtDate = Now
    ARptCodeList.txtTown = TOWNNAME$
    ARptCodeList.Title = "Decal Code List Report"
    ARptCodeList.totCode = Str(NumPrinted)
    ARptCodeList.GetName ReportFile$
    ARptCodeList.startrpt
  Else
    ViewPrint ReportFile$, "Decal Code List Report"
  '  PrintRptFile "Rate Code Listing Report.", "RATELIST.RPT", 1, RetCode%, 1
    KillFile "CodeLIST.RPT"
  End If
  GoTo ExitCodeListing

PrintCodeHeader:
  PageNo = PageNo + 1
  Print #RPTFile, "Vehicle Decal Code Listing."
  Print #RPTFile, TOWNNAME$; Tab(70); "Page:"; PageNo
  Print #RPTFile, "Report Date: "; Date$
  Print #RPTFile, Dash80$
  Print #RPTFile, "Code";
  Print #RPTFile, Tab(7); "Description";
  Print #RPTFile, Tab(28); "Fee";
  Print #RPTFile, Tab(38); "Cash GL(dr)";
  Print #RPTFile, Tab(55); "Rev GL(cr)";
  Print #RPTFile, Tab(67); "Inactive"
  Print #RPTFile, Dash80$
  NumPrinted = 0
Return

GraphicCodeList:
  For cnt = 1 To NumCodeRecs
    Get DCFile, cnt, DCCodeRec(1)
    ToPrint$ = Str(cnt) + "~" + QPTrim$(DCCodeRec(1).CATCODE)
    ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCodeRec(1).CODEDESC)
    ToPrint$ = ToPrint$ + "~" + Using$("#######.##", Str$(DCCodeRec(1).Fee))
    ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCodeRec(1).CASHACCT)
    ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCodeRec(1).REVGLNUM)
    If QPTrim$(DCCodeRec(1).InactiveFlag) = "Y" Then
      ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCodeRec(1).InactiveFlag)
    Else
      ToPrint$ = ToPrint$ + "~" + "N"
    End If
    Print #RPTFile, ToPrint$
    ToPrint$ = ""
    NumPrinted = NumPrinted + 1
    FrmShowPctComp.ShowPctComp cnt, NumCodeRecs

  Next
  Close
Return
ExitCodeListing:

End Sub



