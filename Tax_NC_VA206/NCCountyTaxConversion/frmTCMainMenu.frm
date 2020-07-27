VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmTCMainMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu vs 2.05"
   ClientHeight    =   8760
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   11652
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   10.2
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTCMainMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   492
      Left            =   4080
      TabIndex        =   4
      Top             =   6960
      Width           =   3624
      _Version        =   131072
      _ExtentX        =   6392
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
      ButtonDesigner  =   "frmTCMainMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMatch 
      Height          =   516
      Left            =   4080
      TabIndex        =   2
      Top             =   4080
      Width           =   3624
      _Version        =   131072
      _ExtentX        =   6392
      _ExtentY        =   910
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
      ButtonDesigner  =   "frmTCMainMenu.frx":0AE9
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdNew 
      Height          =   516
      Left            =   4080
      TabIndex        =   1
      Top             =   3360
      Width           =   3624
      _Version        =   131072
      _ExtentX        =   6392
      _ExtentY        =   910
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
      ButtonDesigner  =   "frmTCMainMenu.frx":0D09
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdConvert 
      Height          =   516
      Left            =   4080
      TabIndex        =   3
      Top             =   5520
      Width           =   3624
      _Version        =   131072
      _ExtentX        =   6392
      _ExtentY        =   910
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
      ButtonDesigner  =   "frmTCMainMenu.frx":0F2A
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrint 
      Height          =   516
      Left            =   4080
      TabIndex        =   7
      Top             =   4800
      Width           =   3624
      _Version        =   131072
      _ExtentX        =   6392
      _ExtentY        =   910
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
      ButtonDesigner  =   "frmTCMainMenu.frx":1148
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdResults 
      Height          =   516
      Left            =   4080
      TabIndex        =   5
      Top             =   6240
      Width           =   3624
      _Version        =   131072
      _ExtentX        =   6392
      _ExtentY        =   910
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
      ButtonDesigner  =   "frmTCMainMenu.frx":1366
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdInstructions 
      Height          =   516
      Left            =   4080
      TabIndex        =   0
      Top             =   2640
      Width           =   3624
      _Version        =   131072
      _ExtentX        =   6392
      _ExtentY        =   910
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
      ButtonDesigner  =   "frmTCMainMenu.frx":1580
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8640
      X2              =   8640
      Y1              =   2148
      Y2              =   8036
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   132
      Left            =   8520
      Top             =   2040
      Width           =   972
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2280
      X2              =   2280
      Y1              =   2148
      Y2              =   8036
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   132
      Left            =   2160
      Top             =   2040
      Width           =   972
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NC TAX CONVERSION MAIN MENU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2820
      TabIndex        =   6
      Top             =   1188
      Width           =   6012
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1104
      Index           =   1
      Left            =   1500
      Top             =   830
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1500
      Top             =   708
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   3
      Left            =   2160
      Top             =   1908
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5916
      Index           =   0
      Left            =   2280
      Top             =   2136
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   8520
      Top             =   1908
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5916
      Index           =   1
      Left            =   8640
      Top             =   2136
      Width           =   732
   End
End
Attribute VB_Name = "frmTCMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdConvert_Click()
  If Not Exist("TAXCUST.DAT") Then
    Call TCMsg(900, "The file 'TAXCUST.DAT' could not be found. Process aborted.")
    Exit Sub
  Else
    If TCMsgWOpts(900, "Please be sure the 'TAXCUST.DAT' has not already been converted.", "F10 Continue", "ESC Exit") = "abort" Then
      Exit Sub
    End If
  End If
  
  If Not Exist("TAXPERS.DAT") Then
    Call TCMsg(900, "The file 'TAXPERS.DAT' could not be found. Process aborted.")
    Exit Sub
  End If
  
  If Not Exist("TAXPROP.DAT") Then
    Call TCMsg(900, "The file 'TAXPROP.DAT' could not be found. Process aborted.")
    Exit Sub
  End If
  
  frmTCConvert.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdExit_Click()
  End
End Sub

Private Sub cmdInstructions_Click()
  frmTCInstructions.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdMatch_Click()
  
  If Not Exist(App.Path + "\ParcelsText.csv") Then
    Call TCMsg(900, "The required file 'ParcelsText.csv' cannot be located. Matching up data is aborted.")
    Exit Sub
  End If
  
  frmTCMatchUp.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdNew_Click()
  If Not Exist("ParcelsText.csv") Then
    Call TCMsg(900, "The required file 'ParcelsText.csv' cannot be located. Clearing existing data is aborted.")
    Exit Sub
  End If
  
  If Exist(ConvSpreadFile) And Exist(ConversionFile) Then
    If TCMsgWOpts(900, "Are you sure you want to clear existing data? Press F10 to clear. Otherwise, press ESC to exit.", "F10 Clear", "ESC Exit") = "abort" Then
      Exit Sub
    End If
  End If
  
  KillFile ConvSpreadFile
  KillFile ConversionFile
  KillFile ConvResults
  KillFile ConvErrors
  
  Call Savemsg(900, "Clearing existing data has been completed successfully.")
  
End Sub

Private Sub cmdPrint_Click()
  Dim MaxLines As Integer
  Dim FF$
  Dim Page As Integer
  Dim x As Long
  Dim RptFile As String
  Dim RptHandle As Integer
  Dim TempHandle As Integer
  Dim TempRec As TempConversionData
  Dim NumOfTempRecs As Long
  Dim ThisCity As String * 20
  Dim CntyNum As String
  Dim AddDesc As String
  Dim CustCnt As Long
  Dim TRealVal As Double
  Dim TPersVal As Double
  
  If Not Exist(ConversionFile) Then
    Call TCMsg(900, "Please process the county data first. Load attempt aborted.")
    Exit Sub
  End If
  
  FF$ = Chr(12)
  MaxLines = 58
  RptFile$ = "TCRPTS\MATCH.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  OpenTempConvFile TempHandle, NumOfTempRecs
  Print #RptHandle, "Customer Name"; Tab(51); "Property Pin#"; Tab(71); "County Number"; Tab(91); "City"; Tab(111); "Real Prop Value"; Tab(131); "Pers Prop Value"; Tab(151); "Address/Desc"
  For x = 1 To NumOfTempRecs
    Get TempHandle, x, TempRec
      ThisCity = QPTrim$(TempRec.CData.City)
'    If QPTrim$(TempRec.CData.PinNum) = "029KD031" Then Stop
'    If QPTrim$(TempRec.CData.CountyAcctString) = "10317250" Then Stop
      TRealVal = OldRound(TRealVal + TempRec.CData.PROPVALU)
      TPersVal = OldRound(TPersVal + TempRec.CData.PersVal)
      If TempRec.CData.CountyAcct > 0 Then
        CntyNum = CStr(TempRec.CData.CountyAcct)
      ElseIf TempRec.CData.CountyAcctString <> "" Then
        CntyNum = TempRec.CData.CountyAcctString
      Else
        CntyNum = "Unknown"
      End If
      If TempRec.CData.RealAdd <> "" Then
        AddDesc = TempRec.CData.RealAdd
      ElseIf TempRec.CData.RDESC1 <> "" Then
        AddDesc = TempRec.CData.RDESC1
      ElseIf TempRec.CData.PDESC1 <> "" Then
        AddDesc = TempRec.CData.PDESC1
      End If
      Print #RptHandle, QPTrim$(TempRec.CData.CustName); Tab(51); TempRec.CData.PinNum; Tab(71); CntyNum; Tab(91); ThisCity; Tab(111); Using$("$###,###,##0.00", TempRec.CData.PROPVALU); Tab(131); Using$("$###,###,##0.00", TempRec.CData.PersVal); Tab(151); AddDesc
      CustCnt = CustCnt + 1
  Next x
  Print #RptHandle,
  Print #RptHandle, "Number of Entries:  " + Using$("###,##0", CustCnt)
  Print #RptHandle, "Total Real Value:     " + Using$("$#,###,###,##0.00", TRealVal)
  Print #RptHandle, "Total Personal Value: " + Using$("$#,###,###,##0.00", TPersVal)
  Close
  ViewPrint RptFile, "Match Up Data", True
End Sub

Private Sub cmdResults_Click()
  Dim ConvRec As ConvResultsType
  Dim CRHandle As Integer
  Dim NumOfCRRecs As Long
  Dim x As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim ERptFile$
  Dim ERptHandle As Integer
  Dim dlm$
  Dim TRealVal As Double
  Dim TRealOXVal As Double
  Dim TRealSXVal As Double
  Dim TPersVal As Double
  Dim TMCVal As Double
  Dim TMHVal As Double
  Dim TMTVal As Double
  Dim TCVal As Double
  Dim TPersOXVal As Double
  Dim TPersSXVal As Double
  Dim GTPersVal As Double
  Dim ThisPersVal As Double
  Dim ErrorRec As ConvErrorType
  Dim EHandle As Integer
  Dim NumOfERecs As Long
  Dim TrunName As String * 20
  Dim ThisError$
  
  If Not Exist(ConvResults) Then
    Call TCMsg(900, "Please convert the county data first.")
    Exit Sub
  End If
  
  dlm = "~"
  
  RptFile$ = "TCRPTS\CSTLSTSM.RPT"
  
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  frmTCShowPctComp.Label1 = "Generating Report"
  frmTCShowPctComp.Show , Me
  
  OpenConvResultsFile CRHandle, NumOfCRRecs
  For x = 1 To NumOfCRRecs
    Get CRHandle, x, ConvRec
'    If x = 2987 Then Stop
'    If QPTrim$(ConvRec.CountyAcctString) = "10317250" Then Stop
'    If ConvRec.CountyAcct <> 0 Then Stop
'    If QPTrim$(ConvRec.PinNum) = "0880-20-91-5404" Then Stop
    TRealVal = OldRound(TRealVal + ConvRec.PROPVALU)
    TRealOXVal = OldRound(TRealOXVal + ConvRec.REXMPOTHR)
    TRealSXVal = OldRound(TRealSXVal + ConvRec.REXMPSENI)
    TPersVal = OldRound(TPersVal + ConvRec.PersVal)
    TMCVal = OldRound(TMCVal + ConvRec.MCVALUE)
    TMHVal = OldRound(TMHVal + ConvRec.MHVALUE)
    TMTVal = OldRound(TMTVal + ConvRec.MTVALUE)
    TCVal = OldRound(TCVal + ConvRec.CVALUE)
    TPersOXVal = OldRound(TPersOXVal + ConvRec.PEXMPOTHR)
    TPersSXVal = OldRound(TPersSXVal + ConvRec.PEXMPSENI)
    ThisPersVal = OldRound(ConvRec.PersVal + ConvRec.MCVALUE + ConvRec.MHVALUE + ConvRec.MTVALUE + ConvRec.CVALUE)
    GTPersVal = OldRound(GTPersVal + ThisPersVal)
    TrunName = QPTrim$(ConvRec.CustName)
    '                             0                             1                                 2
    Print #RptHandle, QPTrim$(ConvRec.PinNum); dlm; TrunName; dlm; QPTrim$(ConvRec.CountyAcctString); dlm;
    '                         3                         4                     5                       6
    Print #RptHandle, ConvRec.CountyAcct; dlm; ConvRec.PROPVALU; dlm; ConvRec.REXMPOTHR; dlm; ConvRec.REXMPSENI; dlm;
    '                         7                     8                    9                    10                    11
    Print #RptHandle, ConvRec.PersVal; dlm; ConvRec.MTVALUE; dlm; ConvRec.MCVALUE; dlm; ConvRec.CVALUE; dlm; ConvRec.MHVALUE; dlm;
    '                        12                      13                  14              15               16
    Print #RptHandle, ConvRec.PEXMPOTHR; dlm; ConvRec.PEXMPSENI; dlm; TRealVal; dlm; TRealOXVal; dlm; TRealSXVal; dlm;
    '                    17            18           19           20           21           22                23
    Print #RptHandle, TPersVal; dlm; TMCVal; dlm; TMHVal; dlm; TMTVal; dlm; TCVal; dlm; TPersOXVal; dlm; TPersSXVal; dlm;
    '                     24              25
    Print #RptHandle, GTPersVal; dlm; NumOfCRRecs
    
    frmTCShowPctComp.ShowPctComp x, NumOfCRRecs
    If frmTCShowPctComp.Out = True Then
      Close
      frmTCShowPctComp.Out = False
      Unload frmTCShowPctComp
      Exit Sub
    End If
  Next x
  
  Unload frmTCShowPctComp
  Close
  
  ERptFile$ = "TCRPTS\TCERRORS.RPT"
  ERptHandle = FreeFile
  Open ERptFile For Output As #ERptHandle
  
  OpenConvErrorsFile EHandle, NumOfERecs
  For x = 1 To NumOfERecs
    Get EHandle, x, ErrorRec
    If ErrorRec.ErrorType = 1 Then
      ThisError = "Both real and personal values are greater than zero and using the same pin number."
    ElseIf ErrorRec.ErrorType = 2 Then
      ThisError = "Both real or personal values equal zero."
    End If
    TrunName = QPTrim$(ErrorRec.CustName)
    '                            0                                1                                   2
    Print #ERptHandle, ErrorRec.CountyAcct; dlm; QPTrim$(ErrorRec.CountyAcctString); dlm; TrunName; dlm;
    '                      3                  4                      5                      6
    Print #ERptHandle, ErrorRec.ErrorType; dlm; ErrorRec.PersTot; dlm; ErrorRec.PersXTot; dlm; ErrorRec.PinNum; dlm;
    '                         7                      8
    Print #ERptHandle, ErrorRec.RealTot; dlm; ErrorRec.RealXTot
  Next x
  
  Close
    
  arTCResultsRpt.Show
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Dim FileLen As Integer
  
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  FileLen = Len(Me.Caption)
  FileVers = Mid(Me.Caption, FileLen - 3, 4)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTCMainMenu.")
      End
    End If
  End If

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If

End Sub


