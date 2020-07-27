VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUBSetupMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Maintenance Menu"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   1890
   ClientWidth     =   12210
   ClipControls    =   0   'False
   Icon            =   "FrmSetupMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdGroupCodeMaint 
      Caption         =   "&Group Code Maintenance"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3384
      TabIndex        =   6
      Top             =   6864
      Width           =   2652
   End
   Begin fpBtnAtlLibCtl.fpBtn fpSPTOption 
      Height          =   372
      Left            =   6216
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2016
      Visible         =   0   'False
      Width           =   2580
      _Version        =   131072
      _ExtentX        =   4551
      _ExtentY        =   656
      Enabled         =   0   'False
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
      ButtonDesigner  =   "FrmSetupMenu.frx":08CA
   End
   Begin VB.CommandButton cndBillSetup 
      Caption         =   "Bill Information Setup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   516
      Left            =   3372
      TabIndex        =   4
      Top             =   5448
      Width           =   2652
   End
   Begin VB.CommandButton cmdUBViewSysLog 
      BackColor       =   &H008F8265&
      Caption         =   "View System Log File"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   6216
      MaskColor       =   &H8000000F&
      TabIndex        =   8
      Top             =   3228
      Width           =   2652
   End
   Begin VB.CommandButton cmdSetupMaint 
      Caption         =   "Billing Configuration"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3372
      TabIndex        =   0
      Top             =   2496
      Width           =   2652
   End
   Begin VB.CommandButton cmdRateCodeMenu 
      Caption         =   "Rate Table Maintenance"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3372
      TabIndex        =   1
      Top             =   3228
      Width           =   2652
   End
   Begin VB.CommandButton cmdUBSysDraft 
      Caption         =   "Bank Draft Setup"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3372
      TabIndex        =   2
      Top             =   3972
      Width           =   2652
   End
   Begin VB.CommandButton cmdResetProrate 
      Caption         =   "Reset Prorate Percentages"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   6216
      TabIndex        =   7
      Top             =   2496
      Width           =   2652
   End
   Begin VB.CommandButton cmdRelinkHistory 
      Caption         =   "Relink Utility Files"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3372
      TabIndex        =   3
      Top             =   4704
      Width           =   2652
   End
   Begin VB.CommandButton cmd9 
      Caption         =   "Export Information"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   6216
      TabIndex        =   9
      Top             =   3972
      Width           =   2652
   End
   Begin VB.CommandButton cmdRecalcConsumption 
      BackColor       =   &H008F8265&
      Caption         =   "Recalc Consumption"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   6216
      MaskColor       =   &H8000000F&
      TabIndex        =   10
      Top             =   4704
      Width           =   2652
   End
   Begin VB.CommandButton cndWorkOrderSetup 
      Caption         =   "&Work Order Defaults"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   516
      Left            =   6216
      TabIndex        =   11
      Top             =   5448
      Width           =   2652
   End
   Begin VB.CommandButton cmdEditLateLetter 
      Caption         =   "&Edit Late Notice"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   516
      Left            =   6216
      TabIndex        =   12
      Top             =   6152
      Width           =   2652
   End
   Begin VB.CommandButton cmdExitSetupMenu 
      Caption         =   "E&xit to Previous Menu"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   6228
      TabIndex        =   13
      Top             =   6864
      Width           =   2652
   End
   Begin VB.CommandButton cmdReindex 
      Caption         =   "Reindex Utility Files"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3372
      TabIndex        =   5
      Top             =   6144
      Width           =   2652
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   15
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
            TextSave        =   "3:13 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "7/6/2018"
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "UTILITY SYSTEM SETUP MENU"
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
      Left            =   3348
      TabIndex        =   14
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
Attribute VB_Name = "frmUBSetupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmd9_Click()
  Load frmUBExportMenu
  DoEvents
  frmUBExportMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdGroupCodeMaint_Click()
  frmGroupCodeEntryEdit.Show
End Sub

Private Sub cmdRecalcConsumption_Click()
  Dim FntSize As Integer
  If Not Exist(UBPath$ + "UBCust.dat") Or Not Exist(UBPath$ + "UBTrans.dat") Then
    ReDim MsgText(0 To 5) As String
    frmMsgDialog.RetLabel = "-2"
    FntSize = frmMsgDialog.Label(2).FontSize
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = ""
    MsgText(3) = "NO CUSTOMER INFORMATION!"
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
    Exit Sub
  End If
  Load frmUBRecalcConsumption
  DoEvents
  frmUBRecalcConsumption.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdUBSysDraft_Click()
  Load frmUBSysDraftEdit
  DoEvents
  frmUBSysDraftEdit.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdResetProrate_Click()
  Dim FntSize As Integer
  If Not Exist(UBPath$ + "UBCust.dat") Or Not Exist(UBPath$ + "UBTrans.dat") Then
    ReDim MsgText(0 To 5) As String
    frmMsgDialog.RetLabel = "-2"
    FntSize = frmMsgDialog.Label(2).FontSize
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = ""
    MsgText(3) = "NO CUSTOMER INFORMATION!"
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
    Exit Sub
  End If

  Call ResetProRates
End Sub

Private Sub cmdReindex_Click()
  Dim FntSize As Integer
  If Not Exist(UBPath$ + "UBCust.dat") Then 'Or Not Exist(UBPath$ + "UBTrans.dat") Then
    ReDim MsgText(0 To 5) As String
    frmMsgDialog.RetLabel = "-2"
    FntSize = frmMsgDialog.Label(2).FontSize
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = ""
    MsgText(3) = "NO CUSTOMER INFORMATION!"
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
    Exit Sub
  End If

  ReIndexSystem True
End Sub

Private Sub cmdRateCodeMenu_Click()
  Load frmUBRateMenu
  DoEvents
  frmUBRateMenu.Show
  DoEvents
  Unload Me
End Sub
Private Sub cmdEditLateLetter_Click()
  Load frmUBLateNotices
  DoEvents
  frmUBLateNotices.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdRelinkHistory_Click()
  Dim FntSize As Integer
  If Not Exist(UBPath$ + "UBCust.dat") Or Not Exist(UBPath$ + "UBTrans.dat") Then
    ReDim MsgText(0 To 5) As String
    frmMsgDialog.RetLabel = "-2"
    FntSize = frmMsgDialog.Label(2).FontSize
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = ""
    MsgText(3) = "NO CUSTOMER INFORMATION!"
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
    Exit Sub
  End If

  Load frmRelinkDialog
  DoEvents
  frmRelinkDialog.Show
  Unload Me
End Sub

Private Sub cmdSetupMaint_Click()
  DeActivateControls Me
  DoEvents
  frmInfo.Label1 = "Loading. . ."
  frmInfo.Show
  DoEvents
  Load frmUBControlMaint
  DoEvents
  frmUBControlMaint.Show
  Unload frmInfo
  Unload frmUBSetupMenu
End Sub

Private Sub cmdExitSetupMenu_Click()
  Load frmUBMainMenu
  DoEvents
  frmUBMainMenu.Show
  Unload Me
End Sub

Private Sub cmdUBViewSysLog_Click()
  frmReportOpt.Show 1
  DeActivateControls Me
  If rptopt = 1 Then
    'do the graphics
   Call ViewUBLogFile(True)
  ElseIf rptopt = 2 Then
    'do the text
   Call ViewUBLogFile(False)
   DoEvents
   ActivateControls Me
  Else
    DoEvents
    ActivateControls Me
  End If
  'need to add the access validation here
  'Call ViewUBLogFile  'need password stuff here
End Sub

Private Sub cndBillSetup_Click()
  Load frmBillInfoSetup
  DoEvents
  frmBillInfoSetup.Show
  DoEvents
  Unload Me
End Sub


Private Sub cndWorkOrderSetup_Click()
  Load frmWODefaults
  DoEvents
  frmWODefaults.Show
  DoEvents
  Unload Me
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  'screenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpUtilitySystemSetup
  If InStr(UCase(TOWNNAME$), "SPRUCE") Or InStr(UCase(TOWNNAME$), "FAIRMONT") Or InStr(UCase(TOWNNAME$), "DUMFRIES") Then
    fpSPTOption.Enabled = True
    fpSPTOption.Visible = True
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitSetupMenu.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via SetUpMenu by " + PWUser$
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
  End If
  DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      KeyCode = 0
      SendKeys "%X"
    Case vbKeyHome
      cmdSetupMaint.SetFocus
    Case vbKeyEnd
      cmdExitSetupMenu.SetFocus
    Case Else:
  End Select
End Sub

Private Sub ViewUBLogFile(Grpt As Boolean)
  
  Dim LogFile As String
  Dim LogSize As Long, MaxSize As Long
  Dim CopyPos As Long
  Dim BackName As String, TempDate As String
  Dim cnt As Integer, UBBFile As Integer, UBLFile As Integer
  ReDim CBuff(1) As String
  Dim Nul As String
  On Error GoTo resetstuff
  LogFile = UBPath + "UBLOG.DAT"
  TempDate = Date$
  Nul = Chr$(0)
  
  LogSize = FileSize(LogFile)
  MaxSize = 2048000
  
  If LogSize > MaxSize Then
    GoSub MakeLogFileBackUpName
    Kill BackName
    If SH_Rename(LogFile, BackName) = False Then
    'Name LogFile As BackName
    Exit Sub
    End If
    CBuff(1) = Space$(2048)

    UBBFile = FreeFile
    Open BackName For Binary Shared As #UBBFile
    UBLFile = FreeFile
    Open LogFile For Binary Shared As #UBLFile

    CopyPos = LogSize - (MaxSize / 8)

    Seek #UBBFile, CopyPos

    Do
      Get #UBBFile, , CBuff(1)
      Do
        cnt = InStr(CBuff(1), Nul)
        If cnt > 0 Then
          Mid$(CBuff(1), cnt, 1) = " "
        End If
      Loop While cnt > 0
      Put #UBLFile, , CBuff(1)
    Loop Until eof(UBBFile)

    Close
    UBLog "UB: Resized LOG File."
    End If
 
  Erase CBuff
  
  UBLog "IN: View Utiility Billing Log File."
  
  DoEvents
  If Grpt = False Then
    ViewPrint LogFile, "Utility Billing Log File"
  Else
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmUBSetupMenu
    ARptLineRpt.GetName LogFile$
    ARptLineRpt.startrpt
  End If
  GoTo ExitReviewLog

MakeLogFileBackUpName:
  Do
    cnt = InStr(TempDate, "-")
    If cnt > 0 Then
      TempDate = Left$(TempDate, cnt - 1) + Mid$(TempDate, cnt + 1)
    End If
  Loop While cnt > 0
  TempDate = Left$(TempDate, 4) + Right$(TempDate, 2)
  For cnt = 1 To 9
    BackName = UBPath + "UB" + TempDate + ".DA" + QPTrim$(Str$(cnt))
    If Not Exist(BackName) Then
      Exit For
    End If
  Next
  If cnt > 9 Then    'if cnt is >9 then use this
    BackName = UBPath + "BIGLOG.DAT"
    Call KillFile(BackName)
  End If
Return
resetstuff:
  Close
  ActivateControls Me
  
ExitReviewLog:
 
End Sub

Private Sub fpSPTOption_Click() 'Laser Tax Bill
  If InStr(UCase(TOWNNAME$), "SPRUCE") Or InStr(UCase(TOWNNAME$), "FAIRMONT") Or InStr(UCase(TOWNNAME$), "DUMFRIES") Then
    Load frmTempTaxBillPrint
    DoEvents
    frmTempTaxBillPrint.Show
    DoEvents
    Unload Me
  End If
End Sub
