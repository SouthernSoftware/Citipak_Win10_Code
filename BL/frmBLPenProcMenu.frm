VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmBLPenProcMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Penalty Processing"
   ClientHeight    =   8730
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11565
   Icon            =   "frmBLPenProcMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11565
   StartUpPosition =   2  'CenterScreen
   Tag             =   "7"
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   300
      Left            =   5475
      TabIndex        =   1
      Top             =   7320
      Width           =   690
      _Version        =   131072
      _ExtentX        =   1206
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   9405029
      ForeColor       =   8454143
      Text            =   ""
      Shape           =   0
      ShapeRoundWidth =   180
      ShapeRoundHeight=   180
      BorderWidth     =   -1
      BorderColor     =   -2147483630
      ThreeDWidth     =   -1
      ThreeDShadowColor=   -2147483632
      ThreeDHighlightColor=   16777215
      ThreeDText      =   0
      ThreeDTextHighlightColor=   16777215
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignPictureH   =   0
      AlignPictureV   =   0
      PictureStyle    =   0
      WordWrap        =   -1  'True
      ScaleMode       =   1
      ThreeDStyle     =   2
      Position        =   0
      PosBaseX        =   0
      PosBaseY        =   0
      PosOffsetX      =   -100
      PosOffsetY      =   300
      MaxWidth        =   5000
      CloudInset      =   100
      CloudMinWidth   =   600
      TailShape       =   2
      TailType        =   2
      TailBaseOffsetOutside=   300
      TailBaseOffsetInside=   100
      TailBaseAxisOutside=   0
      TailBaseAxisInside=   0
      TailBubbleCount =   3
      AlignTextH      =   1
      AlignTextV      =   1
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      AutoScan        =   0
      ScanAllDescendants=   -1  'True
      Interval        =   500
      IntervalNext    =   200
      AutoSize        =   -1  'True
      UseTagProp      =   -1  'True
      HideOnInactiveApp=   0   'False
      HideOnMouseDown =   2
      HideOnKeyDown   =   2
      HideOnFocus     =   0   'False
      ScanDisabledControls=   -1  'True
      ThreeDAppearance=   0
      FollowFocus     =   0   'False
      TemplateName    =   ""
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPenProc 
      Height          =   492
      Left            =   3960
      TabIndex        =   2
      Tag             =   "Press to bring up a screen from which to calculate business license penalty fees."
      Top             =   2525
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
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
      ButtonDesigner  =   "frmBLPenProcMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPenRpt 
      Height          =   480
      Left            =   3960
      TabIndex        =   3
      Tag             =   "Press to bring up a screen from which to print a detailed report of penalty assessments."
      Top             =   3135
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
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
      ButtonDesigner  =   "frmBLPenProcMenu.frx":0AB1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPost 
      Height          =   492
      Left            =   3960
      TabIndex        =   4
      Tag             =   "Press to bring up a screen from which to post pending penalty assessments."
      Top             =   3731
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
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
      ButtonDesigner  =   "frmBLPenProcMenu.frx":0C99
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintNotices 
      Height          =   492
      Left            =   3960
      TabIndex        =   5
      Tag             =   "Press to begin the process of printing delinquent notices."
      Top             =   4334
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
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
      ButtonDesigner  =   "frmBLPenProcMenu.frx":0E86
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdLabels 
      Height          =   492
      Left            =   3960
      TabIndex        =   6
      Tag             =   "Press to bring up a screen from which you can print mailing labels for delinquent notices."
      Top             =   4937
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
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
      ButtonDesigner  =   "frmBLPenProcMenu.frx":1072
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClearPen 
      Height          =   492
      Left            =   3960
      TabIndex        =   7
      Tag             =   "Press to clear the current unposted penalty fee file."
      Top             =   5540
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
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
      ButtonDesigner  =   "frmBLPenProcMenu.frx":125A
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   480
      Left            =   3960
      TabIndex        =   8
      Top             =   6150
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
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
      ButtonDesigner  =   "frmBLPenProcMenu.frx":1449
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   492
      Left            =   3960
      TabIndex        =   9
      Tag             =   "Click this button to return to the main Business License main menu."
      Top             =   6746
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
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
      ButtonDesigner  =   "frmBLPenProcMenu.frx":162E
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8666
      X2              =   8666
      Y1              =   2136
      Y2              =   8008
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   150
      Index           =   3
      Left            =   1970
      Top             =   2000
      Width           =   990
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   155
      Index           =   4
      Left            =   8550
      Top             =   1995
      Width           =   990
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1092
      Index           =   1
      Left            =   1455
      Top             =   820
      Width           =   8652
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8655
      X2              =   9359
      Y1              =   8010
      Y2              =   8010
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2086
      X2              =   2795
      Y1              =   8025
      Y2              =   8025
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2086
      X2              =   2086
      Y1              =   2136
      Y2              =   8008
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PENALTY PROCESSING"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2775
      TabIndex        =   0
      Top             =   1170
      Width           =   6012
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1455
      Top             =   690
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   1966
      Top             =   1890
      Width           =   972
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8550
      Top             =   1890
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8655
      Top             =   2130
      Width           =   732
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5895
      Index           =   0
      Left            =   2085
      Top             =   2130
      Width           =   735
   End
End
Attribute VB_Name = "frmBLPenProcMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdClearPen_Click()
  If Not Exist("artmppen.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "No unposted penalty fees exist. Clear penalty fees aborted."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  frmBLMessageBoxJrWOpts.Label1.Caption = "If you wish to delete the current unposted penalty fee files then press F10. Otherwise, press ESC to abort the process."
  frmBLMessageBoxJrWOpts.Label1.Top = 800
  frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 &Continue"
  frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Abort"
  frmBLMessageBoxJrWOpts.Show vbModal
  If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
    Unload frmBLMessageBoxJrWOpts
    Close
    Exit Sub
  Else
    Unload frmBLMessageBoxJrWOpts
  End If
  
  KillFile "artmppen.dat"
  MainLog ("User warned that continuing would delete the unposted penalty fee file. User elected to continue and delete that file (attmppen.dat).")
  
  cmdHelp.Text = "Turn Menu &Help On"
  btnHelp.AutoScan = fpAutoScanOff
  frmBLMessageBoxJr.Label1.Caption = "The unposted penalty fee file has been deleted successfully."
  frmBLMessageBoxJr.Label1.Top = 900
  frmBLMessageBoxJr.Show vbModal
  
End Sub

Private Sub cmdExit_Click()
  KillFile "pencalc.dat"
  frmBLMainMenu.Show
  DoEvents
  Unload frmBLPenProcMenu
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "Turn Menu &Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "Turn Menu &Help On"
    btnHelp.AutoScan = fpAutoScanOff
  End If
End Sub

Private Sub cmdLabels_Click()
  frmBLDlqntMailLbls.Show
  DoEvents
  Unload frmBLPenProcMenu
End Sub

Private Sub cmdPenProc_Click()
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim Towncnt As Integer
  
  On Error Resume Next
  
  OpenTownFile TownHandle
  Towncnt = LOF(TownHandle) / Len(TownRec)
  If Towncnt = 0 Then
    frmBLMessageBoxJrWOpts.Label1.Caption = "No Town Setup records have been saved. Penalty charges are determined either by a percentage or by a fixed amount. These values are saved on the Town Setup screen. Would you like to jump to that screen now?"
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Jump"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC No"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
      Unload frmBLMessageBoxJrWOpts
      Close TownHandle
      Close TownHandle
      frmBLTownSetup.Show
      DoEvents
      frmBLPenCalc.Hide
      Exit Sub
    Else
      Unload frmBLMessageBoxJrWOpts
      Close
      Exit Sub
    End If
  Else
    Get TownHandle, 1, TownRec
  End If
  Close TownHandle
  frmBLPenCalc.Show
  DoEvents
  Unload frmBLPenProcMenu
End Sub

Private Sub cmdPenRpt_Click()
  Dim PrintType$
  
  If Not Exist("artmppen.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "No penalty transactions have been saved."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  frmBLReportOpt.Show vbModal 'opens small screen from which the
  'user selects the printing method
  PrintType$ = frmBLReportOpt.fptxtPrintType
  Select Case PrintType$
    Case "Graphical"
      Call PrintGraphics
    Case "Text"
      frmBLMessageBoxJr.Label1.Caption = "Pitch 10 is recommended for this report."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Call PrintText
    Case "Exit"
  End Select
  Unload frmBLReportOpt
End Sub

Private Sub PrintText()
  Dim ReportFile$
  Dim FF$
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim CustCnt As Integer
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim TransRec As TempPenaltyCharges
  Dim THandle As Integer
  Dim TRNumRecs As Double
  Dim cnt As Double
  Dim TempNum$
  Dim CustNum As Long
  Dim TotalBal#
  Dim Page As Integer
  Dim RptHandle As Integer
  
  On Error GoTo ERRORSTUFF
  ReportFile$ = "ARPENLST.PRN"  'Report File Name
  FF$ = Chr$(12)
  MaxLines = 58
  LineCnt = 0
  CustCnt = 0

  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  GoSub PrintCustBalRptHeader
  OpenPenTransFile THandle
  TRNumRecs = LOF(THandle) / Len(TransRec)
  OpenCustFile CHandle
  frmBLShowPctComp.Label1 = "Calculating Penalty Amounts"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdPenProc.Enabled = False
  cmdPost.Enabled = False
  cmdPenRpt.Enabled = False
  
  For cnt = 1 To TRNumRecs
    Get THandle, cnt, TransRec
    TempNum$ = QPTrim$(TransRec.CustomerNumber)
    If Len(TempNum$) = 0 Then
      GoTo NoPenalty
    End If
    CustNum = Val(TempNum$)
    Get #CHandle, CustNum, CustRec

    Print #RptHandle, Using("####0", CustRec.CustNumb);
    Print #RptHandle, Tab(10); CustRec.BillName;
    Print #RptHandle, Tab(67); Using("$##,##0.00", TransRec.TransAmount)
    CustCnt = CustCnt + 1
    TotalBal# = OldRound(TotalBal# + TransRec.TransAmount)
    LineCnt = LineCnt + 1
    If LineCnt > MaxLines Then
      LineCnt = 1
      Print #RptHandle, Chr$(12)
      GoSub PrintCustBalRptHeader
    End If

NoPenalty:
    frmBLShowPctComp.ShowPctComp cnt, TRNumRecs
    If frmBLShowPctComp.Out = True Then
      Close
      frmBLShowPctComp.Out = False
      Unload frmBLShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdPenProc.Enabled = True
      cmdPost.Enabled = True
      cmdPenRpt.Enabled = True
      Exit Sub
    End If
  Next
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdPenProc.Enabled = True
  cmdPost.Enabled = True
  cmdPenRpt.Enabled = True
  GoSub PrintCustBalRptEnding
  Close         'Close all open files now

  If CustCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No penalty transactions have been saved."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Kill ReportFile$
    Exit Sub
  End If
  
  ViewPrint ReportFile, "Penalty Report", True

  Exit Sub

PrintCustBalRptHeader:
  Page = Page + 1
  Print #RptHandle, Tab(20); "Business License's : Penalty Transaction Listing"
  Print #RptHandle, "Report Date: "; Date$; Tab(65); "Page #"; Page
  Print #RptHandle, "Cust #"; Tab(10); "Customer Name"; Tab(67); "Penalty Amount"
  Print #RptHandle, String$(80, "=")
  LineCnt = 5
Return

PrintCustBalRptEnding:
  Print #RptHandle, String$(80, "-")
  Print #RptHandle, "Total Customers Printed: "; Using("####0", CustCnt); Tab(51); "Total Penalty: "; Using("$###,##0.00", TotalBal#)
  Print #RptHandle, FF$
Return

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustMaintMenu", "PrintText", Erl)
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
Private Sub PrintGraphics()
  Dim ReportFile$
  Dim CustCnt As Integer
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim TransRec As TempPenaltyCharges
  Dim THandle As Integer
  Dim TRNumRecs As Double
  Dim cnt As Double
  Dim TempNum$
  Dim CustNum As Long
  Dim TotalBal#
  Dim RptHandle As Integer
  Dim dlm$
  Dim TownHandle As Integer
  Dim TownRec As TownSetUpType
  Dim TownName$
  
  On Error GoTo ERRORSTUFF
  dlm$ = "~"
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  TownName = QPTrim$(TownRec.TownName)
  
  ReportFile$ = "BLRPTS\ARPENLST.RPT"  'Report File Name
  CustCnt = 0

  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  OpenPenTransFile THandle
  TRNumRecs = LOF(THandle) / Len(TransRec)
  
  OpenCustFile CHandle
  frmBLShowPctComp.Label1 = "Calculating Penalty Amounts"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdPenProc.Enabled = False
  cmdPost.Enabled = False
  cmdPenRpt.Enabled = False

  For cnt = 1 To TRNumRecs
    Get THandle, cnt, TransRec
    TempNum$ = QPTrim$(TransRec.CustomerNumber)
    If Len(TempNum$) = 0 Then
      GoTo NoPenalty
    End If
    CustNum = Val(TempNum$)
    Get #CHandle, CustNum, CustRec
    Print #RptHandle, TownName$; dlm;
    Print #RptHandle, QPTrim$(CustRec.CustNumb); dlm;
    Print #RptHandle, QPTrim$(CustRec.BillName); dlm;
    Print #RptHandle, TransRec.TransAmount
    CustCnt = CustCnt + 1
    TotalBal# = OldRound(TotalBal# + TransRec.TransAmount)

NoPenalty:
    frmBLShowPctComp.ShowPctComp cnt, TRNumRecs
    If frmBLShowPctComp.Out = True Then
      Close
      frmBLShowPctComp.Out = False
      Unload frmBLShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdPenProc.Enabled = True
      cmdPost.Enabled = True
      cmdPenRpt.Enabled = True
      Exit Sub
    End If
  Next
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdPenProc.Enabled = True
  cmdPost.Enabled = True
  cmdPenRpt.Enabled = True
  Close         'Close all open files now

  arBLPenaltyList.Show
  frmBLLoadReport.Show
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustMaintMenu", "PrintText", Erl)
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
Public Sub cmdPost_Click()
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim THandle As Integer
  Dim TransRec As ARTransRecType
  Dim NumOfTransRecs As Double
  Dim NextTransRec As Double
  Dim NumOfARRecs As Integer
  Dim PenTrans As TempPenaltyCharges
  Dim PCnt As Double
  Dim NumOfPen As Double
  Dim TPHandle As Integer
  Dim CustRecNum$
  Dim CustomerNumber As Integer
  Dim Prev As Long
  
  On Error GoTo ERRORSTUFF
  frmBLPostPenalty.Show vbModal
  If QPTrim$(frmBLPostPenalty.fptxtChoice.Text) = "exit" Then
    Close
    If Exist("dlnqnotice.dat") Then
      'if the user is printing delinquent notices and in so doing
      'reacts to a pop-up on that screen that tells him
      'that a penalty file is unposted and would he like to
      'post it now before printing delinquent notices by choosing
      'yes then he is brought here...if at that point he chooses
      'to abort the post then by deleting the file the delinquent print
      'screen opens when it loads (dlnqnotice.dat) will signal the
      'delinquent screen that the user aborted the post so return
      'the user to the delinquent print screen without printing
      KillFile "dlnqnotice.dat"
    End If
    Exit Sub
  End If
  Unload frmBLPostPenalty
  
  OpenCustFile CHandle
  NumOfARRecs = LOF(CHandle) \ Len(CustRec)
  
  OpenTransFile THandle
  NumOfTransRecs = LOF(THandle) \ Len(TransRec)
  NextTransRec = NumOfTransRecs + 1
  
  OpenPenTransFile TPHandle
  NumOfPen = LOF(TPHandle) \ Len(PenTrans)
  If NumOfPen = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no penalty files to post."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  frmBLShowPctComp.Label1 = "Posting Penalty Amounts"
  frmBLShowPctComp.Show
  frmBLShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdPenProc.Enabled = False
  cmdPost.Enabled = False
  cmdPenRpt.Enabled = False
  For PCnt = 1 To NumOfPen
    Get TPHandle, PCnt, PenTrans
    CustRecNum$ = QPTrim$(PenTrans.CustomerNumber)
    If Len(CustRecNum$) > 0 Then
      CustomerNumber = Val(CustRecNum$)
      Get #CHandle, CustomerNumber, CustRec
      CustRec.PenBal = OldRound(CustRec.PenBal + PenTrans.TransAmount)
      CustRec.AcctBal = OldRound(CustRec.AcctBal + PenTrans.TransAmount)
      
      TransRec.BalanceAfterTrans = CustRec.AcctBal
      TransRec.DetailTransType = 101
      TransRec.TransType = 6
      TransRec.TransDesc = "PENALTY"
      TransRec.CashAmount = 0
      TransRec.ChkAmount = 0
      TransRec.CatCodeRec1 = GetCatRecNum(CustRec.BILLCAT1)
      TransRec.CatCodeRec2 = GetCatRecNum(CustRec.BILLCAT2)
      TransRec.CatCodeRec3 = GetCatRecNum(CustRec.BILLCAT3)
      TransRec.CatCodeRec4 = GetCatRecNum(CustRec.BILLCAT4)
      TransRec.CatCodeRec5 = GetCatRecNum(CustRec.BILLCAT5)
      TransRec.CatLicBal1 = CustRec.FeeLicBal1
      TransRec.CatLicBal2 = CustRec.FeeLicBal2
      TransRec.CatLicBal3 = CustRec.FeeLicBal3
      TransRec.CatLicBal4 = CustRec.FeeLicBal4
      TransRec.CatLicBal5 = CustRec.FeeLicBal5
      TransRec.CatLicAmt1 = 0
      TransRec.CatLicAmt2 = 0
      TransRec.CatLicAmt3 = 0
      TransRec.CatLicAmt4 = 0
      TransRec.CatLicAmt5 = 0
      TransRec.CustomerNumber = QPTrim$(PenTrans.CustomerNumber)
      TransRec.ExtraRoom = ""
      TransRec.FeeAmt = 0
      TransRec.LicAmt = 0
      TransRec.IssAmt = 0
      TransRec.LicBal = 0
      TransRec.PenBal = CustRec.PenBal
      TransRec.IssBal = CustRec.IssuanceBal
      TransRec.PenAmt = PenTrans.PenAmt
      TransRec.TransDate = PenTrans.TransDate
      TransRec.Posted2GL = "N"
      TransRec.TransAmount = PenTrans.TransAmount
      TransRec.NextTrans = 0
      
      Put THandle, NextTransRec, TransRec
      'unrem
      If CustRec.FirstTrans = 0 Then
        CustRec.FirstTrans = NextTransRec
        CustRec.LastTrans = NextTransRec
        Put CHandle, CustomerNumber, CustRec
        'unrem
      Else
        Prev = CustRec.LastTrans
        CustRec.LastTrans = NextTransRec
        Put CHandle, CustomerNumber, CustRec
        'unrem
        Get THandle, Prev, TransRec
        TransRec.NextTrans = NextTransRec
        Put THandle, Prev, TransRec
        'unrem
      End If
      NextTransRec = NextTransRec + 1
    End If
    frmBLShowPctComp.ShowPctComp PCnt, NumOfPen
  Next
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdPenProc.Enabled = True
  cmdPost.Enabled = True
  cmdPenRpt.Enabled = True
  
  Close
  Call KillFile("artmppen.dat")
  frmBLSucSave.Label1.Caption = "Penalty records have been successfully posted."
  frmBLSucSave.Label1.Top = 700
  frmBLSucSave.Show vbModal
  MainLog ("Penalty fees posted successfully.")

  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustMaintMenu", "PrintText", Erl)
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

Private Sub cmdPrintNotices_Click()
  Dim TownRec As TownSetUpType
  Dim THandle As Integer
  
  On Error Resume Next
  OpenTownFile THandle
  Get THandle, 1, TownRec
  Close THandle
  If Exist("artownsu.dat") Then
    If TownRec.DLQNotice = 5 Then
      frmBLMessageBoxJrWOpts.Label1.Caption = "No delinquent notice form has been selected. Would you like to jump to the Town Setup screen to select a delinquent notice form?"
      frmBLMessageBoxJrWOpts.Label1.Top = 700
      frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Jump"
      frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC No"
      frmBLMessageBoxJrWOpts.Show vbModal
      If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
        Unload frmBLMessageBoxJrWOpts
        Close
        frmBLTownSetup.Show
        frmBLTownSetup.fpcmbDLQNotice.SetFocus
        DoEvents
        Unload Me
        Exit Sub
      Else
        Unload frmBLMessageBoxJrWOpts
        Close
        Exit Sub
      End If
    End If
  Else
    frmBLMessageBoxJrWOpts.Label1.Caption = "Town setup records have not been saved. Would you like to jump to the Town Setup screen now?"
    frmBLMessageBoxJrWOpts.Label1.Top = 800
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Jump"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC No"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
      Unload frmBLMessageBoxJrWOpts
      Close
      frmBLTownSetup.Show
      frmBLTownSetup.fptxtTownName.SetFocus
      DoEvents
      Unload Me
      Exit Sub
    Else
      Unload frmBLMessageBoxJrWOpts
      Close
      Exit Sub
    End If
  End If
  
  frmBLDelinquentNotices.Show
  DoEvents
  Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
  Dim FileHandle As Integer
  Dim One As Integer
  
  FileHandle = FreeFile
  One = 1
  Open "pencalc.dat" For Output As FileHandle Len = 2
  Print #FileHandle, One
  Close FileHandle
  
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"
      Call cmdExit_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      KillFile "pencalc.dat"
      ClearInUse PWcnt
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLPenProcMenu.")
      Call Terminate
      End
    End If
  End If
End Sub

