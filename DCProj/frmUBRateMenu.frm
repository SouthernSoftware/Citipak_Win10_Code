VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUBRateMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rate Table Menu"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   Icon            =   "frmUBRateMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExitRateMenu 
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
      Height          =   540
      Left            =   3912
      TabIndex        =   6
      Top             =   6120
      Width           =   4356
   End
   Begin VB.CommandButton cmdPrintRateList 
      Caption         =   "Print Rate Table Listing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3912
      TabIndex        =   5
      Top             =   5364
      Width           =   4356
   End
   Begin VB.CommandButton cmdDeleteRateCode 
      Caption         =   "Delete an Existing Rate Table"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3900
      TabIndex        =   4
      Top             =   4596
      Width           =   4356
   End
   Begin VB.CommandButton cmdEditRateCode 
      Caption         =   "Edit an Existing Rate Table"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3900
      TabIndex        =   3
      Top             =   3840
      Width           =   4356
   End
   Begin VB.CommandButton cmdAddNewRate 
      Caption         =   "Add a New Rate Table"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3912
      TabIndex        =   2
      Top             =   3072
      Width           =   4356
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
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
            TextSave        =   "11:57 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "9/30/2004"
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
      Caption         =   "Rate Table Menu"
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
      TabIndex        =   0
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
      FillColor       =   &H8000000B&
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
Attribute VB_Name = "frmUBRateMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim wf As Integer  'use this to indicate what form the call came from

Private Sub cmdAddNewRate_Click()
  Load frmEditAddRateTable
  DoEvents
  frmEditAddRateTable.fpRateRecNo = 0
  frmEditAddRateTable.Show
  Unload frmUBRateMenu
End Sub

Private Sub cmdDeleteRateCode_Click()
  Dim FntSize As Integer
  ReDim MsgText(0 To 5) As String
  If Not Exist(UBPath$ + "UBRate.dat") Then
    frmMsgDialog.RetLabel = "-2"
    UBLog "ERROR: NO RateCode FILE"
    FntSize = frmMsgDialog.Label(3).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = "NO RATE CODE FILE!"
    MsgText(3) = ""
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
    Exit Sub
  Else
    Load frmRateDelete
    frmRateDelete.fpRateDelRec = -1
    DoEvents
    frmRateDelete.Show
    Unload frmUBRateMenu
  End If
End Sub

Private Sub cmdEditRateCode_Click()
  Dim FntSize As Integer
  ReDim MsgText(0 To 5) As String
  If Not Exist(UBPath$ + "UBRate.dat") Then
    frmMsgDialog.RetLabel = "-2"
    UBLog "ERROR: NO RateCode FILE"
    FntSize = frmMsgDialog.Label(3).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = "NO RATE CODE FILE!"
    MsgText(3) = ""
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
    Exit Sub
  Else
    Load frmEditAddRateTable
    DoEvents
    frmEditAddRateTable.fpRateRecNo = -1
    frmEditAddRateTable.Show
    Unload frmUBRateMenu
  End If
End Sub

Private Sub cmdExitRateMenu_Click()
  Load frmUBSetupMenu
  DoEvents
  frmUBSetupMenu.Show
  Unload frmUBRateMenu
End Sub

Private Sub cmdPrintRateList_Click()
  Dim FntSize As Integer
  ReDim MsgText(0 To 5) As String
  If Not Exist(UBPath$ + "UBRate.dat") Then
    frmMsgDialog.RetLabel = "-2"
    UBLog "ERROR: NO RateCode FILE"
    FntSize = frmMsgDialog.Label(3).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = "NO RATE CODE FILE!"
    MsgText(3) = ""
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
    Exit Sub
  Else
    frmReportOpt.Show 1
    DeActivateControls Me
    If rptopt = 1 Then
      'do the graphics
      wf = 1
      PrintRateListing True
    ElseIf rptopt = 2 Then
      'do the text
      PrintRateListing False
      ActivateControls Me
      Me.cmdPrintRateList.SetFocus
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
  wf = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitRateMenu.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via RateMenu by " + PWUser$
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
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyHome
      cmdAddNewRate.SetFocus
    Case vbKeyEnd
      cmdExitRateMenu.SetFocus
    Case Else:
  End Select
End Sub
Public Sub PrintRateListing(graphicflag As Boolean, Optional xx As Form)
  Dim UBRateTblRecLen As Integer, NumRateRecs As Integer
  Dim NumPrinted As Integer
  Dim RCnt As Integer, cnt As Integer
  Dim UBFile As Integer, RPTFile As Integer
  Dim ReportFile As String, ToPrint As String
  ReDim UBRateTblRec(1) As UBRateTblRecType
  ReDim StepText(1 To 10) As String * 40
  Dim Dash80 As String * 78
  UBRateTblRecLen = Len(UBRateTblRec(1))
  NumRateRecs = FileSize(UBPath + "UBRATE.DAT") \ UBRateTblRecLen
  
  If NumRateRecs = 0 Then
    GoTo ExitRateListing
  End If
  
  Dash80$ = String$(78, "-")
  
  NumPrinted = 0

  FrmShowPctComp.Label1 = "Creating Bill/Payment Tax Report."
  FrmShowPctComp.Show , Me

  ReportFile$ = UBPath + "RATELIST.RPT"
  
  UBFile = FreeFile
  Open UBPath + "UBRATE.DAT" For Random Shared As UBFile Len = UBRateTblRecLen
  
  RPTFile = FreeFile
  Open ReportFile$ For Output As RPTFile
  If graphicflag Then
    GoSub GraphicRateList
  Else
    GoSub PrintRateHeader
    For cnt = 1 To NumRateRecs
      Get UBFile, cnt, UBRateTblRec(1)
      If NumPrinted = 3 Then
        Print #RPTFile, Dash80$
        Print #RPTFile, Chr$(12)
        GoSub PrintRateHeader
      End If
      Print #RPTFile, "       Rate Code:  "; UBRateTblRec(1).RATECODE
      Print #RPTFile, "     Description:  "; UBRateTblRec(1).RATEDESC
      Print #RPTFile, "  Minimum Charge:"; Using$("#######.##", Str$(UBRateTblRec(1).MINAMT))
      Print #RPTFile, "   Minimum Units:"; Using$("##########", Str$(UBRateTblRec(1).MINUNITS))
      Print #RPTFile, "      Max Amount:"; Using$("######.##", Str$(UBRateTblRec(1).MaxAmt))
      Print #RPTFile, "      [ Step ]        [ Beg Unit ]     [ Amount/Unit ]"
      For RCnt = 1 To 10
        LSet StepText$(RCnt) = ""
        If UBRateTblRec(1).TblBreaks(RCnt).UNITS >= 0 Then
          Mid$(StepText$(RCnt), 8) = Using$("########", Str$(UBRateTblRec(1).TblBreaks(RCnt).UNITS))
        End If
        If UBRateTblRec(1).TblBreaks(RCnt).UNITAMT >= 0 Then
          Mid$(StepText$(RCnt), 25) = Using$("####.######", Str$(UBRateTblRec(1).TblBreaks(RCnt).UNITAMT))
        End If
      Next
      Print #RPTFile, "     First Break:"; StepText$(1)
      Print #RPTFile, "    Second Break:"; StepText$(2)
      Print #RPTFile, "     Third Break:"; StepText$(3)
      Print #RPTFile, "    Fourth Break:"; StepText$(4)
      Print #RPTFile, "     Fifth Break:"; StepText$(5)
      Print #RPTFile, "     Sixth Break:"; StepText$(6)
      Print #RPTFile, "   Seventh Break:"; StepText$(7)
      Print #RPTFile, "    Eighth Break:"; StepText$(8)
      Print #RPTFile, "     Ninth Break:"; StepText$(9)
      Print #RPTFile, "        All Over:"; StepText$(10)
      Print #RPTFile,
      NumPrinted = NumPrinted + 1
      FrmShowPctComp.ShowPctComp cnt, NumRateRecs
  
    Next
    Print #RPTFile, Dash80$
    Print #RPTFile, Chr$(12)
    Close
  End If
  Erase UBRateTblRec, StepText
  DoEvents
  If graphicflag Then
    Load frmLoadingRpt
    If wf = 1 Then
      frmLoadingRpt.setwherefrom frmUBRateMenu
    Else
      frmLoadingRpt.setwherefrom xx
    End If
    ARptRateList.txtDate = Now
    ARptRateList.txtTown = TOWNNAME$
    ARptRateList.Title = "Rate Table List Report"
    ARptRateList.GetName ReportFile$
    ARptRateList.startrpt
  Else
    ViewPrint ReportFile$, "Rate Table List Report"
  '  PrintRptFile "Rate Code Listing Report.", "RATELIST.RPT", 1, RetCode%, 1
    KillFile "RATELIST.RPT"
  End If
  GoTo ExitRateListing

PrintRateHeader:
  PageNo = PageNo + 1
  Print #RPTFile, "Utility Billing Rate Table Listing."
  Print #RPTFile, TOWNNAME$; Tab(70); "Page:"; PageNo
  Print #RPTFile, "Report Date: "; Date$
  Print #RPTFile, Dash80$
  NumPrinted = 0
Return

GraphicRateList:
  For cnt = 1 To NumRateRecs
    Get UBFile, cnt, UBRateTblRec(1)
    ToPrint$ = UBRateTblRec(1).RATECODE
    ToPrint$ = ToPrint$ + "~" + UBRateTblRec(1).RATEDESC
    ToPrint$ = ToPrint$ + "~" + Using$("#######.##", Str$(UBRateTblRec(1).MINAMT))
    ToPrint$ = ToPrint$ + "~" + Using$("##########", Str$(UBRateTblRec(1).MINUNITS))
    ToPrint$ = ToPrint$ + "~" + Using$("######.##", Str$(UBRateTblRec(1).MaxAmt))
    For RCnt = 1 To 10
      LSet StepText$(RCnt) = ""
      If UBRateTblRec(1).TblBreaks(RCnt).UNITS >= 0 Then
        ToPrint$ = ToPrint$ + "~" + Using$("########", Str$(UBRateTblRec(1).TblBreaks(RCnt).UNITS))
      Else
        ToPrint$ = ToPrint$ + "~ "
      End If
      If UBRateTblRec(1).TblBreaks(RCnt).UNITAMT >= 0 Then
        ToPrint$ = ToPrint$ + "~" + Using$("####.######", Str$(UBRateTblRec(1).TblBreaks(RCnt).UNITAMT))
      Else
        ToPrint$ = ToPrint$ + "~ "
      End If
    Next
    Print #RPTFile, ToPrint$
    ToPrint$ = ""
    NumPrinted = NumPrinted + 1
    FrmShowPctComp.ShowPctComp cnt, NumRateRecs

  Next
  Close
Return
ExitRateListing:

End Sub



