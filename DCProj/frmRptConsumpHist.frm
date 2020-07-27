VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmRptConsumpHist 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Consumption History"
   ClientHeight    =   5664
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   9948
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5664
   ScaleWidth      =   9948
   StartUpPosition =   2  'CenterScreen
   Begin LpLib.fpList fpConsumpList 
      Height          =   3936
      Left            =   216
      TabIndex        =   0
      Top             =   864
      Width           =   9540
      _Version        =   196608
      _ExtentX        =   16828
      _ExtentY        =   6943
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   10.8
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Columns         =   2
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   0
      WrapList        =   0   'False
      WrapWidth       =   0
      SelMax          =   1
      AutoSearch      =   1
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   0
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
      ScrollBarV      =   0
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   2
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   0   'False
      DataAutoSizeCols=   0
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   3
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
      ColDesigner     =   "frmRptConsumpHist.frx":0000
   End
   Begin fpBtnAtlLibCtl.fpBtn CmdOk 
      Height          =   480
      Left            =   5196
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4992
      Width           =   1284
      _Version        =   131072
      _ExtentX        =   2265
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
      ButtonDesigner  =   "frmRptConsumpHist.frx":031C
   End
   Begin fpBtnAtlLibCtl.fpBtn fpcmdPrint 
      Height          =   480
      Left            =   3468
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4992
      Width           =   1284
      _Version        =   131072
      _ExtentX        =   2265
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
      ButtonDesigner  =   "frmRptConsumpHist.frx":04F2
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Name/Acct"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   288
      TabIndex        =   9
      Top             =   192
      Width           =   5892
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Read Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   2088
      TabIndex        =   7
      Top             =   600
      Width           =   1380
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Consumption"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   8136
      TabIndex        =   6
      Top             =   600
      Width           =   1380
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Trans Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   384
      TabIndex        =   4
      Top             =   600
      Width           =   1380
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Meter Type"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   3744
      TabIndex        =   3
      Top             =   600
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Current"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   5448
      TabIndex        =   2
      Top             =   600
      Width           =   1092
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   6744
      TabIndex        =   1
      Top             =   600
      Width           =   1092
   End
End
Attribute VB_Name = "frmRptConsumpHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Or KeyCode = vbKeyReturn Then
    KeyCode = 0
    Call cmdOk_Click
  End If
End Sub

Private Sub cmdOk_Click()
  DoEvents
  Unload frmRptConsumpHist
End Sub

Private Sub fpCmdOK_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Button = 0
  Call cmdOk_Click
End Sub
Public Sub ShowCustConsHist(CustRec&)
  Dim Build As String * 80
  Dim UBSetupLen As Integer, NumofRevs As Integer, RevCnt As Integer
  Dim RLen As Integer, UBCustRecLen As Integer, UBTranRecLen As Integer
  Dim UBFile As Integer, CurBal As Double, PreBal As Double
  Dim UBTran As Integer, PrevTranRec As Long, MtrCnt As Integer
  Dim dcnt As Integer, MeterType As String, MeterConsp As Long
  Dim MaxMeterAmt As Long, MTRMulti As Double, MCnt As Integer
  ReDim Metered(1 To 15)
  ReDim UBSetUpRec(1) As UBSetupRecType
  UBSetupLen = Len(UBSetUpRec(1))
'  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetupLen, 1
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen

'  If InStr(UBSetUpRec(1).UTILNAME, "TROY") > 0 Then
'    TroyFlag = True
'  End If
'  If InStr(UBSetUpRec(1).UTILNAME, "HAMLET") > 0 Then
'    HamFlag = True
'  End If

  NumofRevs = MaxRevsCnt
  For RevCnt = 1 To 15
    RLen = Len(QPTrim$(Left$(UBSetUpRec(1).Revenues(RevCnt).RevName, 14)))
    If RLen >= 0 Then
      NumofRevs = RevCnt - 1
      Exit For
    End If
    If UBSetUpRec(1).Revenues(RevCnt).UseMtr = "Y" Then
      Metered(RevCnt) = True
    End If
  Next

  ReDim UBTranRec(1) As UBTransRecType
  ReDim UBCustRec(1) As NewUBCustRecType

  UBCustRecLen = Len(UBCustRec(1))
  UBTranRecLen = Len(UBTranRec(1))

  UBFile = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
  Get UBFile, CustRec&, UBCustRec(1)
  Close UBFile
  Label7.Caption = QPTrim$(UBCustRec(1).CustName) & "  Acct: " & Str(CustRec&)
  CurBal# = UBCustRec(1).CurrBalance
  PreBal# = UBCustRec(1).PrevBalance

  UBTran = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTran Len = UBTranRecLen

  PrevTranRec& = UBCustRec(1).LastTrans

  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      Get UBTran, PrevTranRec&, UBTranRec(1)
      If UBTranRec(1).TransType = TranUtilityBill Or UBTranRec(1).TransType = TranUtilityBill + 100 Then
        For MtrCnt = 1 To 7
          If UBTranRec(1).MtrTypes(MtrCnt) <> 0 Then
            dcnt = dcnt + 1
           ' ReDim Preserve MChoice(1 To DCnt) As FLen2
           ' If HamFlag Then
             Build$ = ""
             LSet Build$ = " " + Num2Date(UBTranRec(1).TransDate)
           
           ' Else
            Mid$(Build$, 15) = Num2Date(UBTranRec(1).ReadDate)
          ' End If
            Select Case UBTranRec(1).MtrTypes(MtrCnt)
            Case MtrWaterOnly
              MeterType$ = "Water"
            Case MtrSewerOnly
              MeterType$ = "Sewer"
            Case MtrCombined
              MeterType$ = "Combined"
            Case MtrElectric
              MeterType$ = "Electric"
            Case MtrDemand
              MeterType$ = "D Electric"
            Case MtrGas
              MeterType$ = "Gas Meter"
            Case MtrTouchRead
              MeterType$ = "Touch Read"
            Case MtrLightsService
              MeterType$ = "L Service"
            Case -1
              MeterType$ = "L Service"
            End Select

            Mid$(Build$, 28) = MeterType$
            Mid$(Build$, 40) = Using$("##########", Str$(UBTranRec(1).CurRead(MtrCnt)))
            Mid$(Build$, 53) = Using$("##########", Str$(UBTranRec(1).PrevRead(MtrCnt)))
            MeterConsp& = UBTranRec(1).CurRead(MtrCnt) - UBTranRec(1).PrevRead(MtrCnt)
            If MeterConsp& < 0 Then
              MaxMeterAmt& = 10& ^ (Len(Str$(UBTranRec(1).PrevRead(MtrCnt))) - 1)
              MeterConsp& = (MaxMeterAmt& - UBTranRec(1).PrevRead(MtrCnt)) + UBTranRec(1).CurRead(MtrCnt)
            End If
'working here
'            MTRMulti# = 0
'            For MCnt = 1 To 7
'              If UBTranRec(1).MtrTypes(MtrCnt) = GetCustMeterType%(UBCustRec(), MCnt) Then
                MTRMulti# = UBCustRec(1).LocMeters(MtrCnt).MTRMulti
'                If UBCustRec(1).LocMeters(MCnt).MTRUnit = "C" Then
'                  MeterConsp& = MeterConsp& * 7.481
'        'What to do here ??????? if diff multi used on combined
'                  Exit For
'                End If
'              End If
'            Next

            If MTRMulti# = 0 Then
              'If TroyFlag Then
               ' MTRMulti# = 100
              'Else
                MTRMulti# = 1
              'End If
            End If
            MeterConsp& = MeterConsp& * MTRMulti#
            If UBCustRec(1).LocMeters(MtrCnt).MtrUnit = "C" Then
              MeterConsp& = MeterConsp& * 7.481
            End If

            Mid$(Build$, 67) = Using$("##########", Str$(MeterConsp&))
          End If
        If Len(QPTrim(Build$)) <> 0 Then
          frmRptConsumpHist.fpConsumpList.AddItem Build$
          Build$ = ""
        End If

        Next
      End If
'      If Len(QPTrim(Build$)) <> 0 Then
'        frmRptConsumpHist.fpConsumpList.AddItem Build$
'      End If
      PrevTranRec& = UBTranRec(1).PrevTrans
    Loop

    Close UBTran
    If dcnt > 0 Then
      frmRptConsumpHist.Show 1
    Else
      MsgBox "No Transactions to Display.", vbOKOnly, "No Transactions"
    End If
  Else
    Close UBTran
    MsgBox "No Transactions to Display.", vbOKOnly, "No Transactions"
    
  End If

  Erase Metered, UBSetUpRec
  Erase UBTranRec, UBCustRec
Exit Sub


End Sub

Private Sub fpcmdPrint_Click()
  Dim ReportFile As String, UBRpt As Integer, cnt As Integer, go2line As Integer
  Dim gofrom As Integer
  ReportFile$ = UBPath$ + "UBCnHist.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
  Print #UBRpt, ""
  Print #UBRpt, Now
  Print #UBRpt, Tab(2); "Customer Consumption History List"
  Print #UBRpt, Tab(2); QPTrim$(Label7.Caption)
  Print #UBRpt, "----------------------------------------------------------------------------"
  Print #UBRpt, Tab(2); "TransDate"; Tab(16); "ReadDate"; Tab(27); "Meter Type"; Tab(43); "Current"; Tab(56); "Previous"; Tab(67); "Consumption"
  Print #UBRpt, Tab(2); "---------"; Tab(16); "--------"; Tab(27); "----------"; Tab(43); "-------"; Tab(56); "--------"; Tab(67); "-----------"
  If fpConsumpList.ListCount >= fpConsumpList.ListIndex + 18 Then
    go2line = fpConsumpList.ListIndex + 18
  Else
    go2line = fpConsumpList.ListCount - 1
  End If
  gofrom = fpConsumpList.ListIndex
  For cnt = gofrom To go2line
    fpConsumpList.ListIndex = cnt
    fpConsumpList.col = 0
    Print #UBRpt, fpConsumpList.ColText
  Next
  Close #UBRpt
  PrintConsmpScreen

End Sub
