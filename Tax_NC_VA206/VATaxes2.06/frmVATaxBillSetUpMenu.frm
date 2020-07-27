VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmVATaxBillSetUpMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Billing Setup Menu"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxBillSetUpMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdGLSUBilling 
      Height          =   435
      Left            =   3960
      TabIndex        =   3
      Top             =   3675
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmVATaxBillSetUpMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdTaxGLSUPayments 
      Height          =   432
      Left            =   3960
      TabIndex        =   2
      Top             =   3120
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   762
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
      ButtonDesigner  =   "frmVATaxBillSetUpMenu.frx":0ABB
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMortCodeMaint 
      Height          =   432
      Left            =   3960
      TabIndex        =   1
      Top             =   2580
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   762
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
      ButtonDesigner  =   "frmVATaxBillSetUpMenu.frx":0CAD
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDefaultSettings 
      Height          =   432
      Left            =   3960
      TabIndex        =   0
      Top             =   2040
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   762
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
      ButtonDesigner  =   "frmVATaxBillSetUpMenu.frx":0E9A
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdRelinkAbs 
      Height          =   435
      Left            =   3960
      TabIndex        =   5
      Top             =   4755
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmVATaxBillSetUpMenu.frx":1089
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   432
      Left            =   3960
      TabIndex        =   10
      Top             =   7476
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   762
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
      ButtonDesigner  =   "frmVATaxBillSetUpMenu.frx":126D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdRelinkTrans 
      Height          =   435
      Left            =   3960
      TabIndex        =   6
      Top             =   5295
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmVATaxBillSetUpMenu.frx":144A
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdRateTables 
      Height          =   435
      Left            =   3960
      TabIndex        =   4
      Top             =   4215
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmVATaxBillSetUpMenu.frx":1631
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdReindex 
      Height          =   435
      Left            =   3960
      TabIndex        =   7
      Top             =   5835
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmVATaxBillSetUpMenu.frx":1825
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExport 
      Height          =   432
      Left            =   3960
      TabIndex        =   8
      Top             =   6396
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   762
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
      ButtonDesigner  =   "frmVATaxBillSetUpMenu.frx":1A00
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdImport 
      Height          =   432
      Left            =   3960
      TabIndex        =   9
      Top             =   6936
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   762
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
      ButtonDesigner  =   "frmVATaxBillSetUpMenu.frx":1BEE
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1098
      Index           =   1
      Left            =   1493
      Top             =   803
      Width           =   8655
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   2094
      Top             =   2019
      Width           =   971
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8706
      X2              =   8706
      Y1              =   2127
      Y2              =   8028
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   8586
      Top             =   2017
      Width           =   971
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8706
      X2              =   9408
      Y1              =   8020
      Y2              =   8020
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2199
      X2              =   2914
      Y1              =   8020
      Y2              =   8020
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2214
      X2              =   2214
      Y1              =   2127
      Y2              =   8015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TAX SETUP MENU"
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
      Left            =   2813
      TabIndex        =   11
      Top             =   1164
      Width           =   6012
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1214
      Left            =   1495
      Top             =   687
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   2094
      Top             =   1886
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5910
      Index           =   0
      Left            =   2213
      Top             =   2117
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8585
      Top             =   1887
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5910
      Index           =   1
      Left            =   8706
      Top             =   2117
      Width           =   732
   End
End
Attribute VB_Name = "frmVATaxBillSetUpMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Public IExFN As String
  Private Temp_Class As Resize_Class
Private Sub cmdDefaultSettings_Click()
  frmVATaxSystemSetup.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdExit_Click()
  If Exist("C:\CPWork\exitsetup.dat") Then
    Unload frmVATaxSystemSetup
    Unload frmVATaxRevSpreadsheets
    Unload frmVATaxRealPctSetup
    Unload frmVATaxPersPctSetup
    KillFile "C:\CPWork\exitsetup.dat"
  End If
  
  frmVATaxMainMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdExport_Click()
  If Not Exist(TaxCustFile) Then
    Call TaxMsg(900, "No customer files have been saved. Export attempt aborted.")
    Exit Sub
  End If
  
  DeActivateControls Me
  ExpPostalCass
  ActivateControls Me
End Sub

Private Sub cmdGLSUBilling_Click()
  frmVATaxBillPostOpt.Show vbModal
  If frmVATaxBillPostOpt.fptxtPostType.Text = "Real" Then
    'do the graphics
    frmVATaxBillGLSetUp.Show
  ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Personal" Then
    'do the text
    frmVATaxPBillGLSetUp.Show
  ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Exit" Then
    Exit Sub
  End If
  
  DoEvents
  Unload Me
End Sub

Private Sub cmdImport_Click()
  Dim FntSize As Integer, msgtogo As String
  
  If Not Exist(TaxCustFile) Then
    Call TaxMsg(900, "No customer files have been saved. Export attempt aborted.")
    Exit Sub
  End If
  
  If TaxMsgWOpts(700, "Please note that if you are not CASS Certified this feature is unusable. Call Southern Software @ 1-800-842-8190 for assistance.", "F10 Continue", "ESC Exit") = "abort" Then
    Unload frmVATaxMsgWOpts
    Exit Sub
  End If
  MainLog ("WARNING: User warned when attempting to import CASS text file that if they were not CASS Certified then this feature is unusable. User continued anyway.")
  
  IExFN = ""
  msgtogo = "TxPostal.cmv" '"Enter the Name of the Import File."
  'this hardcodes the file name
  'on the frmimpexpmsg the field is set to readonly so cant change
  frmVATaxImpExpMsg.txtFileName.Text = "TxPostal.cmv" ' "Enter Import File Name Here"
  frmVATaxImpExpMsg.txtFileName.Visible = True
  frmVATaxImpExpMsg.Label1 = msgtogo
  frmVATaxImpExpMsg.Show 1, Me
  If frmVATaxImpExpMsg.Exout <> 1 Then
    DeActivateControls Me
    ImpPostalCass
    ActivateControls Me
  Else
    Exit Sub
  End If

End Sub

Private Sub cmdMortCodeMaint_Click()
  frmVATaxMortSetup.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdRateTables_Click()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim One As Integer
  Dim AHandle As Integer
  
  If Exist(TaxSetupName) Then
    OpenTaxSetUpFile TMHandle
    Get TMHandle, 1, TaxMasterRec
    Close TMHandle
    If QPTrim$(TaxMasterRec.OptRev1) = "" And QPTrim$(TaxMasterRec.POptRev1) = "" Then
      If QPTrim$(TaxMasterRec.OptRev2) = "" And QPTrim$(TaxMasterRec.POptRev2) = "" Then
        If QPTrim$(TaxMasterRec.OptRev3) = "" And QPTrim$(TaxMasterRec.POptRev3) = "" Then
'          If TaxMsgWOpts(600, "No optional revenue files have been saved. To set up an optional revenue please refer to Tab 2 on the 'Tax System Default Settings' screen accessed from this menu. If you would like to jump there now press F10. Otherwise, press ESC to return to the menu.", "F10 Jump", "ESC Exit") = "abort" Then
'            Unload frmVATaxMsgWOpts
'            Exit Sub
'          Else
'            Unload frmVATaxMsgWOpts
            One = 1
            AHandle = FreeFile
            Open "C:\CPWork\ratetbls.dat" For Output As AHandle
            Print #AHandle, One
            Close AHandle
            Call TaxMsg(800, "There are no real or personal optional revenues saved. Please refer to the second tab of the System Setup screen to add optional revenues.")
'            frmVATaxSystemSetup.Show
'            frmVATaxRevSpreadsheets.Show
'            frmVATaxSystemSetup.vaTabPro1.ActiveTab = 1
'            frmVATaxRevSpreadsheets.vaSpread1.SetActiveCell 1, 5
            DoEvents
'            Unload Me
'            Exit Sub
'          End If
        End If
      End If
    End If
  Else
    If TaxMsgWOpts(800, "Please set up the tax system defaults before continuing. If you would like to jump to that screen now press F10. Otherwise, press ESC to return to the menu.", "F10 Jump", "ESC Escape") = "abort" Then
      Unload frmVATaxMsgWOpts
      Exit Sub
    Else
      One = 1
      AHandle = FreeFile
      Open "C:\CPWork\ratetbls.dat" For Output As AHandle
      Print #AHandle, One
      Close AHandle
      frmVATaxSystemSetup.Show
      DoEvents
      Unload Me
      Exit Sub
    End If
  End If

  frmVATaxRateMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdReindex_Click()
  frmVATaxSaveAnimation.Label1.Caption = "Reindexing"
  frmVATaxSaveAnimation.Show
  Call CreateCustNameIdx
  Call CreateSrchNameIdx
  Call CreateOptCustIdx
  Call CreateSSIdx
  Unload frmVATaxSaveAnimation
  Call Savemsg(900, "Reindexing has completed successfully.")
End Sub

Private Sub cmdRelinkAbs_Click()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  Dim x As Long, ThisPin&
  Dim PHandle As Integer
  Dim NumOfPRecs As Long
  Dim FoundIt As Integer
  Dim ThisPersRec&
  Dim ThisRealRec&
  Dim RptFile$
  Dim RptHandle As Integer
  Dim FF$
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim Page As Integer
  Dim dlm$
  Dim PersCnt&
  Dim CCnt&
  Dim BadPersCnt As Long
  Dim BadRealCnt As Long
  Dim RealCnt&
  Dim PersDesc$
  
  On Error GoTo ERRORSTUFF
  
  ReDim RealRec(1 To 2) As PropertyRecType
  OpenTaxPropFile RHandle, NumOfRRecs
  ReDim PersRec(1 To 2) As PersonalRecType
  OpenTaxPersFile PHandle, NumOfPRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs

  ReDim BadReal(1 To 1) As Long
  BadRealCnt = 0
  ReDim BadPers(1 To 1) As Long
  BadPersCnt = 0
  
  frmVATaxShowPctComp.Label1 = "Relinking"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  For x = 1 To NumOfPRecs&  'clear pers pointers
    Get #PHandle, x, PersRec(1)
    PersRec(1).NextRec = 0
    Put #PHandle, x, PersRec(1)
    frmVATaxShowPctComp.ShowPctComp x, NumOfPRecs
    DoEvents
  Next x
  
  frmVATaxShowPctComp.Label1 = "Relinking"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  For x = 1 To NumOfRRecs&  'clear real pointers
    Get #RHandle, x, RealRec(1)
    RealRec(1).NextRec = 0
    Put #RHandle, x, RealRec(1)
    frmVATaxShowPctComp.ShowPctComp x, NumOfRRecs
    DoEvents
  Next x

  frmVATaxShowPctComp.Label1 = "Relinking"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  ReDim CPins(1 To NumOfTCRecs&) As Long
  For x = 1 To NumOfTCRecs&
    Get TCHandle, x, TaxCust            'read the customers PINS
    CPins(x) = x
    TaxCust.FirstPropRec = 0 'zero old pointers
    TaxCust.FirstPersRec = 0 'zero old pointers
    Put TCHandle, x, TaxCust              'update cust file
    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
    DoEvents
  Next x
  
  frmVATaxShowPctComp.Label1 = "Relinking"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  'Process the Personal abstracts
  For PersCnt& = 1 To NumOfPRecs&
    Get PHandle, PersCnt&, PersRec(1)
    ThisPin& = PersRec(1).CustPin&
    FoundIt = 0
    For CCnt& = 1 To NumOfTCRecs&                'Look for pin in customer
      If ThisPin& = CPins(CCnt&) Then           'found matching pin
        Get TCHandle, CCnt&, TaxCust       'get the cust rec
        If TaxCust.FirstPersRec > 0 Then     'if there are others
          ThisPersRec& = TaxCust.FirstPersRec
          Do
            Get PHandle, ThisPersRec&, PersRec(2)
            If PersRec(2).NextRec = 0 Then
              PersRec(2).NextRec = PersCnt&     'point to this pers rec
              Put PHandle, ThisPersRec&, PersRec(2)
              Exit Do
            End If
            ThisPersRec& = PersRec(2).NextRec
          Loop
        Else    'no first personal rec
          TaxCust.FirstPersRec = PersCnt&    'point cust to this pers rec
          Put #TCHandle, CCnt&, TaxCust       'update cust file
          PersRec(1).NextRec = 0                'set pers next pointer to 0
          Put #PHandle, PersCnt&, PersRec(1)   'update pers file
        End If
        FoundIt = -1
        Exit For                'done with this pers rec
      End If

    Next CCnt
    If FoundIt = 0 Then
      BadPersCnt = BadPersCnt + 1
      ReDim Preserve BadPers(1 To BadPersCnt) As Long
      BadPers(BadPersCnt) = PersCnt&
    End If
    frmVATaxShowPctComp.ShowPctComp PersCnt, NumOfPRecs
    DoEvents
  Next PersCnt
  
  frmVATaxShowPctComp.Label1 = "Relinking"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  For RealCnt& = 1 To NumOfRRecs&
    Get #RHandle, RealCnt&, RealRec(1)
    ThisPin& = RealRec(1).CustPin&
    FoundIt = 0
    For CCnt& = 1 To NumOfTCRecs&                'Look for pin in customer
      If ThisPin& = CPins(CCnt&) Then           'found matching pin
        Get TCHandle, CCnt&, TaxCust         'get the cust rec
        If TaxCust.FirstPropRec > 0 Then     'if there are others
          ThisRealRec& = TaxCust.FirstPropRec
          Do
            Get #RHandle, ThisRealRec&, RealRec(2)
            If RealRec(2).NextRec = 0 Then
              RealRec(2).NextRec = RealCnt&     'point to this real rec
              Put RHandle, ThisRealRec&, RealRec(2)
              Exit Do
            End If
            ThisRealRec& = RealRec(2).NextRec
          Loop
        Else
          TaxCust.FirstPropRec = RealCnt&    'point cust to this pers rec
          Put TCHandle, CCnt&, TaxCust       'update cust file
          RealRec(1).NextRec = 0                'set real next pointer to 0
          Put RHandle, RealCnt&, RealRec(1)   'update pers file
        End If
        FoundIt = -1
        Exit For                'done with this pers rec
      End If
    Next CCnt

    If FoundIt = 0 Then
      BadRealCnt = BadRealCnt + 1
      ReDim Preserve BadReal(1 To BadRealCnt) As Long
      BadReal(BadRealCnt) = RealCnt&
    End If
    frmVATaxShowPctComp.ShowPctComp RealCnt, NumOfRRecs
    DoEvents
  Next RealCnt
  
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
 
  Call Savemsg(900, "Abstract relinking has completed successfully.")
  
  If BadRealCnt > 0 And BadPersCnt = 0 Then
    If TaxMsgWOpts(750, "There were " + CStr(BadRealCnt) + " real estate properties that could not be relinked. Press F10 for a print out of these properties.", "F10 Print", "ESC Don't Print") = "abort" Then
      Unload frmVATaxMsgWOpts
      Exit Sub
    Else
      GoSub PrintIt
    End If
  ElseIf BadRealCnt = 0 And BadPersCnt > 0 Then
    If TaxMsgWOpts(750, "There were " + CStr(BadPersCnt) + " personal properties that could not be relinked. Press F10 for a print out of these properties.", "F10 Print", "ESC Don't Print") = "abort" Then
      Unload frmVATaxMsgWOpts
      Exit Sub
    Else
      GoSub PrintIt
    End If
  ElseIf BadRealCnt > 0 And BadPersCnt > 0 Then
    If TaxMsgWOpts(700, "There were " + CStr(BadPersCnt) + " personal properties that could not be relinked and " + CStr(BadRealCnt) + " real properties that could not be relinked. Press F10 for a print out of these properties.", "F10 Print", "ESC Don't Print") = "abort" Then
      Unload frmVATaxMsgWOpts
      Exit Sub
    Else
      GoSub PrintIt
    End If
  End If
  
  Close
  Exit Sub

PrintIt:
  frmVATaxReportOpt.Show vbModal
  If frmVATaxReportOpt.fptxtPrintType.Text = "Graphical" Then
    Unload frmVATaxReportOpt
    GoSub PrintInGraphics
  ElseIf frmVATaxReportOpt.fptxtPrintType.Text = "Text" Then
    Unload frmVATaxReportOpt
    GoSub PrintInText
  End If
   
  Return

PrintInGraphics:
  dlm$ = "~"
  RptFile$ = "TAXRPTS\ABSRELINKERR.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  If BadRealCnt > 0 Then
    For x = 1 To BadRealCnt
      Get RHandle, BadReal(x), RealRec(1)
      '                    0                       1
      Print #RptHandle, "REAL"; dlm; QPTrim$(RealRec(1).RealPin); dlm;
      If QPTrim$(RealRec(1).PropAddr) <> "" Then
        '                            2
        Print #RptHandle, QPTrim(RealRec(1).PropAddr); dlm;
      ElseIf QPTrim$(RealRec(1).PROPNOT1) <> "" Then
        '                          2
        Print #RptHandle, RealRec(1).PROPNOT1; dlm;
      Else
        '                                      2
        Print #RptHandle, "Map/Block/Lot: " + QPTrim$(RealRec(1).Map) + "/" + QPTrim$(RealRec(1).BLOCK) + "/" + QPTrim$(RealRec(1).LOTNUMB); dlm;
      End If
      '                        3
      Print #RptHandle, RealRec(1).CustPin
    Next x
  End If
  
  If BadPersCnt > 0 Then
    For x = 1 To BadPersCnt
      Get PHandle, BadPers(x), PersRec(1)
      '
      Print #RptHandle, "PERSONAL"; dlm; QPTrim$(PersRec(1).PropPin); dlm;
      If PersRec(1).PersVal > 0 Then
        PersDesc = "Personal"
      End If
      If PersRec(1).CVALUE > 0 Then
        If Len(PersDesc) = 0 Then
          PersDesc = "Farm Eq"
        Else
          PersDesc = PersDesc + "/Farm Eq"
        End If
      End If
      If PersRec(1).MHValue > 0 Then
        If Len(PersDesc) = 0 Then
          PersDesc = "Mobile Home"
        Else
          PersDesc = PersDesc + "/Mobile Home"
        End If
      End If
      If PersRec(1).MTValue > 0 Then
        If Len(PersDesc) = 0 Then
          PersDesc = "Mach Tools"
        Else
          PersDesc = PersDesc + "/Mach Tools"
        End If
      End If
      If PersRec(1).MCValue > 0 Then
        If Len(PersDesc) = 0 Then
          PersDesc = "Merch Cap"
        Else
          PersDesc = PersDesc + "/Merch Cap"
        End If
      End If
      If Len(PersDesc) > 0 Then
        Print #RptHandle, PersDesc; dlm;
      Else
        Print #RptHandle, "NO DESCRIPTION"; dlm;
      End If
    Next x
    Print #RptHandle, PersRec(1).CustPin
  End If
  
  Close
  arVATaxRelinkErr.Show
  
  Return
  
PrintInText:
  FF$ = Chr(12)
  MaxLines = 58
  RptFile$ = "TAXRPTS\ABSRELINKERR.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  If BadRealCnt > 0 Then
    Print #RptHandle, "REAL PROPERTY LISTING: NO CUSTOMER/OWNER FOUND"
    Print #RptHandle, "Cust Pin #"; Tab(12); "Prop Pin #"; Tab(37); "Description/Address"
    Print #RptHandle, String(75, "-")
    LineCnt = 3
    For x = 1 To BadRealCnt
      Get RHandle, BadReal(x), RealRec(1)
      Print #RptHandle, Using$("#####0", RealRec(1).CustPin); Tab(12); QPTrim$(RealRec(1).RealPin); Tab(37);
      If QPTrim$(RealRec(1).PropAddr) <> "" Then
        Print #RptHandle, QPTrim(RealRec(1).PropAddr)
      ElseIf QPTrim$(RealRec(1).PROPNOT1) <> "" Then
        Print #RptHandle, RealRec(1).PROPNOT1
      Else
        Print #RptHandle, "Map/Block/Lot: " + QPTrim$(RealRec(1).Map) + "/" + QPTrim$(RealRec(1).BLOCK) + "/" + QPTrim$(RealRec(1).LOTNUMB)
      End If
      LineCnt = LineCnt + 1
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        Print #RptHandle, "REAL PROPERTY LISTING: NO CUSTOMER/OWNER FOUND"
        Print #RptHandle, "Cust Pin #"; Tab(12); "Prop Pin #"; Tab(37); "Description/Address"
        Print #RptHandle, String(75, "-")
        LineCnt = 3
      End If
    Next x
  End If
  
  If BadPersCnt > 0 Then
    If BadRealCnt > 0 Then
      Print #RptHandle, FF$
    End If
    Print #RptHandle, "PERSONAL PROPERTY LISTING: NO CUSTOMER/OWNER FOUND"
    Print #RptHandle, "Cust Pin #"; Tab(12); "Prop Pin #"; Tab(37); "Description/Address"
    Print #RptHandle, String(75, "-")
    LineCnt = 3
    For x = 1 To BadPersCnt
      Get PHandle, BadPers(x), PersRec(1)
      Print #RptHandle, Using$("#####0", PersRec(1).CustPin); Tab(12); QPTrim$(PersRec(1).PropPin); Tab(37);
      If QPTrim$(PersRec(1).DESC1) <> "" Then
        Print #RptHandle, QPTrim(PersRec(1).DESC1)
      Else
        If PersRec(1).PersVal > 0 Then
          PersDesc = "Personal"
        End If
        If PersRec(1).CVALUE > 0 Then
          If Len(PersDesc) = 0 Then
            PersDesc = "Farm Eq"
          Else
            PersDesc = PersDesc + "/Farm Eq"
          End If
        End If
        If PersRec(1).MHValue > 0 Then
          If Len(PersDesc) = 0 Then
            PersDesc = "Mobile Home"
          Else
            PersDesc = PersDesc + "/Mobile Home"
          End If
        End If
        If PersRec(1).MTValue > 0 Then
          If Len(PersDesc) = 0 Then
            PersDesc = "Mach Tools"
          Else
            PersDesc = PersDesc + "/Mach Tools"
          End If
        End If
        If PersRec(1).MCValue > 0 Then
          If Len(PersDesc) = 0 Then
            PersDesc = "Merch Cap"
          Else
            PersDesc = PersDesc + "/Merch Cap"
          End If
        End If
        If Len(PersDesc) > 0 Then
          Print #RptHandle, PersDesc
        End If
      End If
      LineCnt = LineCnt + 1
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        Print #RptHandle, "PERSONAL PROPERTY LISTING: NO CUSTOMER/OWNER FOUND"
        Print #RptHandle, "Cust Pin #"; Tab(12); "Prop Pin #"; Tab(37); "Description/Address"
        Print #RptHandle, String(75, "-")
        LineCnt = 3
      End If
    Next x
  End If
  Print #RptHandle, FF$
  
  Close
  
  ViewPrint RptFile, "Relink Property Errors", True
    
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillSetUpMenu", "cmdReLinkAbs_Click", Erl)
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
 
End Sub

Private Sub cmdRelinkTrans_Click()
  Dim TaxCust As TaxCustType
  Dim NumOfTCRecs As Long
  Dim TaxTran As TaxTransactionType
  Dim TCHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TTHandle As Integer
  Dim cnt As Long
  Dim BadCnt As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim FF$
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim Page As Integer
  Dim dlm$
  Dim ThisTransType$
  
  On Error GoTo ERRORSTUFF
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  
  For cnt& = 1 To NumOfTCRecs&
    Get TCHandle, cnt&, TaxCust
    TaxCust.LastTrans = 0
    Put TCHandle, cnt&, TaxCust
  Next
  frmVATaxShowPctComp.Label1 = "Relinking"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  
  BadCnt = 0
  ReDim BadTran(1 To 1) As Long
  
  For cnt& = 1 To NumOfTTRecs&
    Get TTHandle, cnt&, TaxTran
    If TaxTran.CustomerRec > 0 And TaxTran.CustomerRec <= NumOfTCRecs& Then
      Get TCHandle, TaxTran.CustomerRec, TaxCust
      TaxTran.LastTrans = TaxCust.LastTrans
      TaxCust.LastTrans = cnt&
      Put TCHandle, TaxTran.CustomerRec, TaxCust
      Put TTHandle, cnt&, TaxTran
    Else
      BadCnt = BadCnt + 1
      ReDim Preserve BadTran(1 To BadCnt) As Long
      BadTran(BadCnt) = cnt&
    End If
    frmVATaxShowPctComp.ShowPctComp cnt, NumOfTTRecs
    DoEvents
  Next cnt
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True

  Call Savemsg(900, "Relinking has completed successfully.")
  If BadCnt > 0 Then
    If TaxMsgWOpts(800, "There are " + CStr(BadCnt) + " transaction(s) that could not be relinked. Press F10 to print a list of these transactions.", "F10 Print", "ESC Don't Print") = "abort" Then
      Unload frmVATaxMsgWOpts
    Else
      frmVATaxReportOpt.Show vbModal
      If frmVATaxReportOpt.fptxtPrintType.Text = "Graphical" Then
        Unload frmVATaxReportOpt
        GoSub PrintInGraphics
      ElseIf frmVATaxReportOpt.fptxtPrintType.Text = "Text" Then
        Unload frmVATaxReportOpt
        GoSub PrintInText
      End If
    End If
  End If
  Close
     
  Exit Sub
  
PrintInGraphics:
  dlm$ = "~"
  RptFile$ = "TAXRPTS\RELINKERR.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  For cnt = 1 To BadCnt
    Get TTHandle, BadTran(cnt), TaxTran
    GoSub GetType
    '                          0                     1                      2                                  3                              4
    Print #RptHandle, TaxTran.CustomerRec; dlm; TaxTran.Amount; dlm; MakeRegDate(TaxTran.TransDate); dlm; ThisTransType
  Next cnt
  
  arVATaxRelinkErrTrans.Show
  
  Return
  
PrintInText:
  FF$ = Chr(12)
  MaxLines = 58
  RptFile$ = "TAXRPTS\ABSRELINKERR.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  Print #RptHandle, "Relink Error Report for Transactions"
  Print #RptHandle, "Report Date: " + CStr(Now)
  Print #RptHandle, "Cust Rec#"; Tab(17); "Trans Amt"; Tab(28); "Trans Date"; Tab(42); "Trans Type"
  Print #RptHandle, String(75, "-")
  LineCnt = 4
  For cnt = 1 To BadCnt
    Get TTHandle, BadTran(cnt), TaxTran
    GoSub GetType
    Print #RptHandle, Using$("#####0", TaxTran.CustomerRec); Tab(12); Using$("$##,###,##0.00", TaxTran.Amount); Tab(28); MakeRegDate(TaxTran.TransDate); Tab(42); ThisTransType
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      Print #RptHandle, "Relink Error Report for Transactions"
      Print #RptHandle, "Report Date: " + CStr(Now)
      Print #RptHandle, "Cust Rec#"; Tab(17); "Trans Amt"; Tab(28); "Trans Date"; Tab(42); "Trans Type"
      Print #RptHandle, String(75, "-")
      LineCnt = 4
    End If
  Next cnt
  
  Close
  ViewPrint RptFile, "Relink Error Report", True
  
  Return

GetType:
   Select Case TaxTran.TranType
     Case 1
       ThisTransType = "Billing"
     Case 2
       ThisTransType = "Payment"
     Case 3
       ThisTransType = "Release"
     Case 4
       ThisTransType = "Interest"
     Case 5
       ThisTransType = "Penalty"
     Case 6
       ThisTransType = "Advertising Charge"
     Case 7
       If TaxTran.CustPin = 0 Then
         ThisTransType = "Adjust 'DOS'"
       Else
         ThisTransType = "Adjust Pay Down"
       End If
     Case 9
       ThisTransType = "Credit Applied at Billing"
     Case 13
       ThisTransType = "Adjust Bill Down"
     Case 14
       ThisTransType = "Adjust Bill Up"
     Case 21
       ThisTransType = "Billpay/Overpay"
     Case 22
       ThisTransType = "Overpayment"
     Case 10
       ThisTransType = "Adjust Pay Down Affecting Credit Balance"
     Case 11
       ThisTransType = "Adjust Prepay Down"
     Case 12
       ThisTransType = "Refund Prepay"
     Case 24
       ThisTransType = "Adjust Bill Up Affecting Credit Balance"
     Case Else
       ThisTransType = "Unknown"
    End Select
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillSetUpMenu", "cmdRelinkTrans", Erl)
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
 
End Sub

Private Sub cmdTaxGLSUPayments_Click()
  frmVATaxBillPostOpt.Show vbModal
  If frmVATaxBillPostOpt.fptxtPostType.Text = "Real" Then
    'do the graphics
    frmVATaxPayGLSetup.Show
  ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Personal" Then
    'do the text
    frmVATaxPPayGLSetUp.Show
  ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Exit" Then
    Exit Sub
  End If
  
  DoEvents
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%M"
      Call cmdExit_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpTaxSystemSetupAnd
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxBillSetupMenu.")
      Call Terminate
      End
    End If
  End If

End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub ExpPostalCass()
  Dim q$, Zip$
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim TaxRpt As Integer
  Dim cnt As Long
  Dim Export&, Cty$
  Dim Address1$, Address2$
  
  On Error GoTo ERRORSTUFF
  
  frmVATaxShowPctComp.Label1 = "Creating Export File"
  frmVATaxShowPctComp.cmdCancel.Enabled = False
  frmVATaxShowPctComp.Show , Me
  q$ = ","
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  KillFile "TaxPostal.txt"
  TaxRpt = FreeFile
  Open "TaxPostal.txt" For Output As TaxRpt

  For cnt = 1 To NumOfTCRecs
    Get TCHandle, cnt, TaxCust
    '*************************************
    '   Main body of Printing goes here
    If TaxCust.Deleted = 0 Then
      Export& = Export& + 1
      GoSub ExportThisAccount
    End If
    frmVATaxShowPctComp.ShowPctComp cnt, NumOfTCRecs
    If frmVATaxShowPctComp.Out Then
      Close
      Unload frmVATaxShowPctComp
      GoTo ExitHere
    End If
  Next cnt

  Close
  
  frmVATaxShowPctComp.ShowPctComp 1, 1
  If Export& > 0 Then
    Call Savemsg(800, "File " & "TaxPostal.txt Exported with " & CStr(Export&) & " Accounts.")
  Else
    Call TaxMsg(900, "No Information Found to Export.")
  End If
  
  GoTo ExitHere

ExportThisAccount:

  Zip$ = QPTrim$(TaxCust.Zip)
  Address1$ = QPStripCom$(TaxCust.Addr1)
  Address2$ = QPStripCom$(TaxCust.Addr2)
  Cty$ = QPStripCom$(TaxCust.City)
  Print #TaxRpt, QPTrim$(Str$(cnt));
  Print #TaxRpt, q$; Address1$;
  Print #TaxRpt, q$; Address2$;
  Print #TaxRpt, q$; Cty$;
  Print #TaxRpt, q$; QPTrim$(TaxCust.State);
  Print #TaxRpt, q$; Zip$

Return

ExitHere:

  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillSetUpMenu", "ExpPostalCass", Erl)
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
 
End Sub

Private Sub ImpPostalCass()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim TaxRpt As Integer
  Dim Acct$, Import&
  Dim Add1$
  Dim Add2$
  Dim Cty$
  Dim ST$
  Dim Zp$
  Dim Dp$
  Dim X1$
  Dim RR$
  Dim Lot$
  Dim X2$, cnt As Long
  Dim Zip$, ErrCnt As Long
  Dim Address1$
  Dim Address2$
  
  On Error GoTo ERRORSTUFF
  
  frmVATaxShowPctComp.Label1 = "Creating Export File"
  frmVATaxShowPctComp.cmdCancel.Enabled = False
  frmVATaxShowPctComp.Show , Me
  On Local Error GoTo impend
  If Len(IExFN) > 0 Then
    If Not Exist(IExFN) Then
      frmVATaxShowPctComp.ShowPctComp 1, 1
      Call TaxMsg(800, "The Import File " + IExFN + " Does Not Exist In The Citipak Folder.")
      Exit Sub
    End If
  Else
    frmVATaxShowPctComp.ShowPctComp 1, 1
    Call TaxMsg(900, "The Import File Name Entered Is Not Valid.")
    Exit Sub
  End If
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
'  ReDim UBCustRec(1) As NewUBCustRecType
'  UBCustRecLen = Len(UBCustRec(1))

'  ReDim UBOwnerRec(1) As UBOwnerRecType
'  UBOwnerRecLen = Len(UBOwnerRec(1))

'  IdxRecLen = 4               'we are using a long integer
'  IdxFileSize& = FileSize(IndexName$)
'  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
'  ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
  'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
'  NumOfRecs = IdxNumOfRecs

'  Handle = FreeFile
'  Open IndexName$ For Random Shared As Handle Len = IdxRecLen
'  For cnt = 1 To IdxNumOfRecs
'    Get #Handle, cnt, IdxBuff(cnt)
'  Next
'  Close Handle

'  UBFile = FreeFile
'  Open UBPath$ + "UBOWNER.DAT" For Random Shared As UBFile Len = UBOwnerRecLen
'  NumOfRecs& = GetNumOfCust&
'  UBCust = FreeFile
'  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  
  'KillFile (UBPath$ + "UBPostal.txt")
  TaxRpt = FreeFile
  Open IExFN For Input As TaxRpt
  'Input header record
  Input #TaxRpt, Acct$, Add1$, Add2$, Cty$, ST$, Zp$, Dp$, X1$, RR$, Lot$, X2$
  Do While Not eof(TaxRpt)
    Acct = ""
    Add1 = ""
    Add2 = ""
    Cty = ""
    ST = ""
    Zp = ""
    Dp = ""
    X1 = ""
    RR = ""
    Lot = ""
    X2 = ""
    Input #TaxRpt, Acct$, Add1$, Add2$, Cty$, ST$, Zp$, Dp$, X1$, RR$, Lot$, X2$
    If Val(Acct$) > NumOfTCRecs Then
      ErrCnt = ErrCnt + 1
    End If
  Loop
  Close TaxRpt
  
  If ErrCnt > 0 Then
    frmVATaxShowPctComp.ShowPctComp 1, 1
    Call TaxMsg(800, "ERROR: " + CStr(ErrCnt) + " file(s) were found to be in error. Please contact the originator of the file for assistance. Import aborted.")
    Close
    Exit Sub
  End If
  
  TaxRpt = FreeFile
  Open IExFN For Input As TaxRpt
  Input #TaxRpt, Acct$, Add1$, Add2$, Cty$, ST$, Zp$, Dp$, X1$, RR$, Lot$, X2$
  Do While Not eof(TaxRpt)
    Acct = ""
    Add1 = ""
    Add2 = ""
    Cty = ""
    ST = ""
    Zp = ""
    Dp = ""
    X1 = ""
    RR = ""
    Lot = ""
    X2 = ""
    Input #TaxRpt, Acct$, Add1$, Add2$, Cty$, ST$, Zp$, Dp$, X1$, RR$, Lot$, X2$
    cnt = Val(Acct)
  'If cnt is greater than the total num of customers then bad data
    If cnt <= NumOfTCRecs& Then
      Get TCHandle, cnt, TaxCust
      If TaxCust.Deleted = 0 Then
        Import& = Import& + 1
        GoSub ImportThisAccount
      End If
    Else
      frmVATaxShowPctComp.ShowPctComp 1, 1
      Call TaxMsg(800, "No Information Found to Import. Procedure Ended.")
      GoTo impend
    End If
  Loop
  
  Close
  
  frmVATaxShowPctComp.ShowPctComp 1, 1
  
  If Import& > 0 Then
    Call Savemsg(800, "File " & IExFN & " Successfully Imported with " & CStr(Import&) & " Accounts.")
  Else
    Call TaxMsg(900, "No Information Found to Import. Procedure Ended.")
  End If
  
  GoTo ExitHere

ImportThisAccount:
  Zip$ = QPTrim$(Zp)
  Address1$ = QPTrim$(Add1)
  Address2$ = QPTrim$(Add2)
  TaxCust.Addr1 = Address1$
  TaxCust.Addr2 = Address2$
  TaxCust.City = QPTrim$(Cty)
  TaxCust.State = QPTrim$(ST)
  TaxCust.Zip = QPTrim$(Zip$)
  TaxCust.DeliveryPt = QPTrim$(Dp$)
  TaxCust.PostalRt = QPTrim$(RR$)
  Put TCHandle, cnt, TaxCust
 
  Return

impend:

  Close

ExitHere:

  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillSetUpMenu", "ImpPostalCass", Erl)
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
 
End Sub

