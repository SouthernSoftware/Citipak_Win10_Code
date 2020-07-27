VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmVATaxBillingMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Billing Menu"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxBillingMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdInterest 
      Height          =   435
      Left            =   4080
      TabIndex        =   5
      Top             =   4980
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
      ButtonDesigner  =   "frmVATaxBillingMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPost 
      Height          =   435
      Left            =   4080
      TabIndex        =   3
      Top             =   3855
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
      ButtonDesigner  =   "frmVATaxBillingMenu.frx":0AAC
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintReprint 
      Height          =   435
      Left            =   4080
      TabIndex        =   1
      Tag             =   "0"
      Top             =   2730
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
      ButtonDesigner  =   "frmVATaxBillingMenu.frx":0C8E
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrebill 
      Height          =   435
      Left            =   4080
      TabIndex        =   0
      Top             =   2160
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
      ButtonDesigner  =   "frmVATaxBillingMenu.frx":0E79
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdLateNotice 
      Height          =   435
      Left            =   4080
      TabIndex        =   6
      Top             =   5550
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
      ButtonDesigner  =   "frmVATaxBillingMenu.frx":106B
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   435
      Left            =   4080
      TabIndex        =   9
      Top             =   7260
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
      ButtonDesigner  =   "frmVATaxBillingMenu.frx":1251
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMortExport 
      Height          =   435
      Left            =   4080
      TabIndex        =   2
      Top             =   3300
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
      ButtonDesigner  =   "frmVATaxBillingMenu.frx":142E
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdReprintPosted 
      Height          =   435
      Left            =   4080
      TabIndex        =   7
      Tag             =   "0"
      Top             =   6120
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
      ButtonDesigner  =   "frmVATaxBillingMenu.frx":1622
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPenalty 
      Height          =   420
      Left            =   4080
      TabIndex        =   4
      Top             =   4425
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmVATaxBillingMenu.frx":180E
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClear 
      Height          =   435
      Left            =   4080
      TabIndex        =   8
      Top             =   6690
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
      ButtonDesigner  =   "frmVATaxBillingMenu.frx":19EF
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TAX BILLING MENU"
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
      TabIndex        =   10
      Top             =   1164
      Width           =   6012
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2214
      X2              =   2214
      Y1              =   2127
      Y2              =   8015
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2199
      X2              =   2914
      Y1              =   8020
      Y2              =   8020
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8706
      X2              =   9408
      Y1              =   8020
      Y2              =   8020
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   8586
      Top             =   2017
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
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   2094
      Top             =   2019
      Width           =   971
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
Attribute VB_Name = "frmVATaxBillingMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim Allow As Boolean
  Public MortOK As Boolean
  Public Real As Boolean
Private Sub cmdClear_Click()
  Dim RBillInfo As VARETaxBillInfoType
  Dim PBillInfo As VAPPTaxBillInfoType
  Dim BIHandle As Integer
  Dim ThisYear$
  
  If Exist(RealTaxBillInfoFile) And Exist(PersTaxBillInfoFile) Then
    frmVATaxBillPostOpt.Show vbModal
    If frmVATaxBillPostOpt.fptxtPostType.Text = "Personal" Then
      GoSub DeletePersonal
    ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Real" Then
      GoSub DeleteReal
    ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Exit" Then
      DoEvents
      Unload frmVATaxBillPostOpt
      Exit Sub
    End If
  ElseIf Exist(PersTaxBillInfoFile) And Not Exist(RealTaxBillInfoFile) Then
    GoSub DeletePersonal
  ElseIf Not Exist(PersTaxBillInfoFile) And Exist(RealTaxBillInfoFile) Then
    GoSub DeleteReal
  ElseIf Not Exist(PersTaxBillInfoFile) And Not Exist(RealTaxBillInfoFile) Then
    Call TaxMsg(800, "No unposted billing files currently exist. Delete attempt aborted.")
  End If
  
  Exit Sub
  
DeletePersonal:
  If TaxMsgWOpts(700, "WARNING: IF YOU CHOOSE TO CONTINUE THEN ALL UNPOSTED PERSONAL BILLING FILES WILL BE REMOVED PERMANENTLY. IF YOU WISH TO CONTINUE THEN PRESS F10. OTHERWISE PRESS ESC TO LEAVE UNPOSTED PERSONAL BILLING FILES UNCHANGED.", "F10 Delete", "ESC Abort") = "abort" Then
    Exit Sub
  Else
    KillFile "txpblsprn.dat"
    KillFile PersTaxBillOPFile
    KillFile PersTaxBillInfoFile
    KillFile PersTaxBillFile '5.16.07
    MainLog ("User deleted unposted personal billing files after being warned about the consequences.")
    Call TaxMsg(800, "All unposted personal billing files have been deleted successfully.")
  End If
  
  Return
  
DeleteReal:
  If TaxMsgWOpts(700, "WARNING: IF YOU CHOOSE TO CONTINUE THEN ALL UNPOSTED REAL BILLING FILES WILL BE REMOVED PERMANENTLY. IF YOU WISH TO CONTINUE THEN PRESS F10. OTHERWISE PRESS ESC TO LEAVE UNPOSTED REAL BILLING FILES UNCHANGED.", "F10 Delete", "ESC Abort") = "abort" Then
    Exit Sub
  Else
    OpenRealBillInfoFile BIHandle
    Get BIHandle, 1, RBillInfo
    Close BIHandle
  
    ThisYear = CStr(RBillInfo.TaxYear)
    KillFile "mortx" + ThisYear + ".dat"
    KillFile "txrblsprn.dat"
    KillFile RealTaxBillOPFile
    KillFile RealTaxBillInfoFile
    KillFile RealTaxBillFile '5.16.07
    MainLog ("User deleted unposted real billing files after being warned about the consequences.")
    Call Savemsg(800, "All unposted real billing files have been deleted successfully.")
  End If
  
  Return
End Sub

Private Sub cmdExit_Click()
  If Exist("C:\CPWork\exitsetup.dat") Then
    Unload frmVATaxSystemSetup
    Unload frmVATaxRevSpreadsheets
    Unload frmVATaxRealPctSetup
    Unload frmVATaxPersPctSetup
    KillFile "C:\CPWork\exitsetup.dat"
  End If
  
  KillFile "C:\CPWork\lateltr.dat"
  frmVATaxMainMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdInterest_Click()
'  Dim MoveOn As Boolean
'
'  MoveOn = True
'  If Check4PayBatch("P") = True Then
'    frmVATaxUnpostedPaylist.BillType = "P"
'    frmVATaxUnpostedPaylist.Label1.Caption = "An unposted personal payment file is ready for posting. Interest calculations cannot be conducted until these personal payments are posted. Operators involved are shown in the list below."
'    frmVATaxUnpostedPaylist.Show vbModal
'    MoveOn = False
'  End If
'
'  If Check4PayBatch("R") = True Then
'    frmVATaxUnpostedPaylist.BillType = "R"
'    frmVATaxUnpostedPaylist.Label1.Caption = "An unposted real payment file is ready for posting. Interest calculations cannot be conducted until these real payments are posted. Operators involved are shown in the list below."
'    frmVATaxUnpostedPaylist.Show vbModal
'    Exit Sub
'  End If
'
'  If MoveOn = False Then Exit Sub
  
  If Allow = False Then
    Call TaxMsg(900, "No tax customers have been saved. Access denied.")
    Exit Sub
  End If
  
  frmVATaxInterestMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdLateNotice_Click()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim One As Integer
  Dim AHandle As Integer
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  If TaxMasterRec.LateForm = 0 Then
    If TaxMsgWOpts(700, "Late notice letters cannot be processed because there have not been any late letters saved. If you would like to jump to the System SetUp screen to set up a late letter than press F10. Otherwise, press ESC to return to the menu.", "F10 Jump", "ESC Exit") = "abort" Then
      Unload frmVATaxMsgWOpts
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
      One = 1
      AHandle = FreeFile
      Open "C:\CPWork\lateltr.dat" For Output As AHandle
      Print #AHandle, One
      Close AHandle
      frmVATaxSystemSetup.Show
      DoEvents
      Me.Hide
      Exit Sub
    End If
  End If
    
  If Allow = False Then
    Call TaxMsg(900, "No tax customers have been saved. Access denied.")
    Exit Sub
  End If
  
  frmVATaxLateNoticeMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdMortExport_Click()
  Dim MortRec As MortCodeRecType
  Dim MHandle As Integer
  Dim NumOfMCodes As Integer
  Dim x As Integer
  Dim MortCnt As Integer
  
  OpenMortCodeFile MHandle, NumOfMCodes
  
  If NumOfMCodes = 0 Then
    Call TaxMsg(900, "There are no mortgage codes saved.")
    Close MHandle
    Exit Sub
  End If
  
  For x = 1 To NumOfMCodes
    Get MHandle, x, MortRec
    If MortRec.Deleted = 0 Then
      Exit For
    End If
  Next x
  
  If x > NumOfMCodes Then
    Call TaxMsg(900, "There are no valid mortgage codes saved.")
    Close MHandle
    Exit Sub
  End If
  Close MHandle
  
  If Not Exist("txrblsprn.dat") Then
    Call TaxMsg(800, "Please print real tax bills before creating mortgage company export files.")
    Close
    Exit Sub
  End If
  
  frmVATaxMortgageExport.Show
  DoEvents
  Unload Me
    
End Sub

Private Sub cmdPenalty_Click()
  Dim MoveOn As Boolean
  
  MoveOn = True
  If Check4PayBatch("P") = True Then
    frmVATaxUnpostedPaylist.BillType = "P"
    frmVATaxUnpostedPaylist.Label1.Caption = "An unposted personal payment file is ready for posting. Penalty calculations cannot be conducted until these personal payments are posted. Operators involved are shown in the list below."
    frmVATaxUnpostedPaylist.Show vbModal
    MoveOn = False
  End If
  
  If Check4PayBatch("R") = True Then
    frmVATaxUnpostedPaylist.BillType = "R"
    frmVATaxUnpostedPaylist.Label1.Caption = "An unposted real payment file is ready for posting. Penalty calculations cannot be conducted until these real payments are posted. Operators involved are shown in the list below."
    frmVATaxUnpostedPaylist.Show vbModal
    Exit Sub
  End If
  
  If MoveOn = False Then Exit Sub
  frmVATaxPenaltyMenu.Show
  DoEvents
  Unload Me
End Sub

Public Sub cmdPost_Click()
  Dim TBPostRec As TaxBillPostDateType
  Dim PostHandle As Integer
  Dim NumOfPostRecs As Long
  Dim RTaxBill As VARETaxBillType
  Dim PTaxBill As VAPPTaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim OPBillRec As TaxTransactionType
  Dim OPHandle As Integer
  Dim NumOfOPRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim RealRec As PropertyRecType
  Dim RRHandle As Integer
  Dim NumOfRRREcs As Long
  Dim PersRec As PersonalRecType
  Dim PRHandle As Integer
  Dim NumOfPRRecs As Long
  Dim RBillInfo As VARETaxBillInfoType
  Dim PBillInfo As VAPPTaxBillInfoType
  Dim BIHandle As Integer
  Dim x As Long, NextRecord&
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim Previous&
  Dim POverPayYN As Boolean, y As Long
  Dim ROverPayYN As Boolean
  Dim EmptyPay As TaxTransactionType
  Dim OverPayAmt As Double
  Dim PayTranRec As TaxTransactionType
  Dim TotalPaid#
  Dim MortRec As MortCodeRecType
  Dim MCHandle As Integer
  Dim NumOfMCRecs As Integer
  Dim ThisYear$
  Dim One As Integer
  Dim AHandle As Integer
  Dim FileName$
  Dim DupCnt As Integer
  Dim FirstTrans As Long
  Dim LastTrans As Long
  Dim NextRec As Long
  
  On Error GoTo ERRORSTUFF
  
  If Not Exist("txrblsprn.dat") And Not Exist("txpblsprn.dat") Then
    Call TaxMsg(900, "Please print either real or personal tax bills before posting.")
    Exit Sub
  End If
  
  If Exist(RealTaxBillInfoFile) And Exist(PersTaxBillInfoFile) Then
    frmVATaxBillPostOpt.Show vbModal
    If frmVATaxBillPostOpt.fptxtPostType.Text = "Personal" Then
      GoSub PostPersonal
    ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Exit" Then
      DoEvents
      Unload frmVATaxBillPostOpt
      Exit Sub
    End If
  ElseIf Exist(PersTaxBillInfoFile) Then
    GoSub PostPersonal
  End If
  
  If Not Exist("txrblsprn.dat") Then
    Call TaxMsg(900, "ERROR: Please print real tax bills before attempting to post.")
    Exit Sub
  End If
    
  ROverPayYN = False
  If Exist(RealTaxBillOPFile) Then
    ROverPayYN = True
  End If
  
  OpenMortCodeFile MCHandle, NumOfMCRecs
  Close MCHandle
  
  OpenRealBillInfoFile BIHandle
  Get BIHandle, 1, RBillInfo
  Close BIHandle
  
  ThisYear = CStr(RBillInfo.TaxYear)
  DupCnt = CountReprintFiles("POSTR" + Mid(RBillInfo.CountyPara, 1, 3) + Mid(RBillInfo.CyclePara, 1, 3) + Mid(RBillInfo.TwnShpPara, 1, 3) + ThisYear)
  FileName = "TAXBILLBU\POSTR" + Mid(RBillInfo.CountyPara, 1, 3) + Mid(RBillInfo.CyclePara, 1, 3) + Mid(RBillInfo.TwnShpPara, 1, 3) + ThisYear + CStr(DupCnt) + ".DAT"
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  TaxMasterRec.DiscRXDate = RBillInfo.XDate '9/20/05
  TaxMasterRec.RTaxYear = RBillInfo.TaxYear '9/20/05
  Put TMHandle, 1, TaxMasterRec '9/20/05
  Close TMHandle
  
  If MortOK = True Then
    Unload frmVATaxMortgageExport
    GoTo SkipMort
  End If
  If NumOfMCRecs > 0 And Not Exist("mortx" + ThisYear + ".dat") Then
    If TaxMsgWOpts(600, "WARNING: If you wish to create mortgage company tax bill export files then please do so before continuing with this post procedure. If you wish to create the mortgage company tax bill export files then press F10. Otherwise, press ESC to continue posting.", "F10 Make Files", "ESC Continue Post") = "abort" Then
      MainLog ("WARNING: User warned that mortgage company export files should be created before posting tax bills but they elected to post without creating export files.")
      Unload frmVATaxMsgWOpts
    Else
      Unload frmVATaxMsgWOpts
      One = 1
      AHandle = FreeFile
      Open "frombillpost.dat" For Output As AHandle
      Print #AHandle, One
      Close AHandle
      frmVATaxMortgageExport.Show
      MortOK = True
      DoEvents
      Exit Sub
    End If
  End If
  
SkipMort:
  MortOK = False
  If TaxMsgWOpts(900, "READY TO POST REAL TAX BILLS? Press F10 to continue. Otherwise, press ESC to Exit.", "F10 POST", "ESC EXIT") = "abort" Then
    Close
    Exit Sub
  End If
  
  OpenRealTaxBillFile TBHandle, NumOfTBRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxPropFile RRHandle, NumOfRRREcs
  FirstTrans = 0
  LastTrans = 0
  frmVATaxShowPctComp.Label1 = "Posting Real Tax Bills"
  frmVATaxShowPctComp.Show , Me
  frmVATaxShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  If ROverPayYN = True Then
    OpenRealTaxBillOverPayFile OPHandle, NumOfOPRecs
  End If
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, RTaxBill
      If RTaxBill.SetDscvry2No = "Y" Then
        Get RRHandle, RTaxBill.RealPropRecord, RealRec
        RealRec.PROPDISC = "N"
        RealRec.LateList = "N"
        Put RRHandle, RTaxBill.RealPropRecord, RealRec
      End If
      'Update the Transaction File First
      If RTaxBill.BillPrinted = False Then GoTo NotARealBill
      TaxTrans.TransDate = Date2Num%(Date$)
      TaxTrans.TaxYear = RTaxBill.TaxYear
      TaxTrans.TranType = 1  '1=Bill 2=Payment 3=Release 4=Interest 5=Penalty 6=Collection/Ad Cost Billing 7=Adjustment 9=Credit Applied at Billing
      TaxTrans.BillType = "R"
      TaxTrans.Amount = RTaxBill.TotalBillDue      'Total Transaction Amount
      TaxTrans.Revenue.Principle1 = OldRound(RTaxBill.TotalBillDue - (RTaxBill.OptRevTax1 + RTaxBill.OptRevTax2 + RTaxBill.OptRevTax3 + RTaxBill.LateTaxDue))
      TaxTrans.Revenue.Principle2 = 0
      TaxTrans.Revenue.Principle3 = 0
      TaxTrans.Revenue.Principle4 = 0
      TaxTrans.Revenue.Principle5 = 0
      TaxTrans.Revenue.Interest = 0
      TaxTrans.Revenue.Penalty = 0
      TaxTrans.Revenue.Collection = 0
      TaxTrans.Revenue.Future1 = 0
      TaxTrans.Revenue.Future2 = 0
      TaxTrans.Revenue.Principle1Pd = 0
      TaxTrans.Revenue.Principle2Pd = 0
      TaxTrans.Revenue.Principle3Pd = 0
      TaxTrans.Revenue.Principle4Pd = 0
      TaxTrans.Revenue.Principle5Pd = 0
      TaxTrans.Revenue.InterestPd = 0
      TaxTrans.Revenue.PenaltyPd = 0
      TaxTrans.Revenue.CollectionPd = 0
      TaxTrans.Revenue.Future1Pd = 0
      TaxTrans.Revenue.Future2Pd = 0
      TaxTrans.Revenue.RevOpt1 = RTaxBill.OptRevTax1
      TaxTrans.Revenue.RevOpt1Pd = 0
      TaxTrans.Revenue.RevOpt2 = RTaxBill.OptRevTax2
      TaxTrans.Revenue.RevOpt2Pd = 0
      TaxTrans.Revenue.RevOpt3 = RTaxBill.OptRevTax3
      TaxTrans.Revenue.RevOpt3Pd = 0
      TaxTrans.Revenue.LateList = RTaxBill.LateTaxDue
      TaxTrans.Revenue.LateListPd = 0
      TaxTrans.Revenue.PrePaidAmt = 0
      TaxTrans.Revenue.PrePaidUsed = 0
      TaxTrans.Revenue.PrePaidBal = GetOverPayBalance(RTaxBill.CustRec, "R")
      TaxTrans.InternalPin = RTaxBill.InternalPin
      TaxTrans.Revenue.pad = ""
      
      TaxTrans.Description = "Tax Bill #" + Str$(RTaxBill.BillNumber)
      TaxTrans.Posted2GL = "N"
      TaxTrans.CustomerRec = RTaxBill.CustRec
      TaxTrans.LastTrans = 0
      TaxTrans.BelongTo = 0
      TaxTrans.Padding = ""
      TaxTrans.PersPin = 0
      TaxTrans.RealPin = QPTrim$(RTaxBill.RealPin)
      TaxTrans.CustPin = RTaxBill.CustPin
      TaxTrans.DiscXDate = TaxMasterRec.DiscRXDate
      TaxTrans.DiscAmt = 0
      TaxTrans.OperNum = OperNum
      TaxTrans.PersVal = 0
      TaxTrans.PPTRAVal = 0
      TaxTrans.PPTRADisc = 0
      TaxTrans.CntyPara = QPTrim$(RBillInfo.CountyPara)
      TaxTrans.CyclPara = QPTrim$(RBillInfo.CyclePara)
      TaxTrans.TShpPara = QPTrim$(RBillInfo.TwnShpPara)
      TaxTrans.PPTRARmvl = 0
      TaxTrans.PPTRARmvlDate = 0
    'Increment Transaction File Record Count
      NextRecord& = (LOF(TTHandle) / Len(TaxTrans)) + 1
      RTaxBill.PostDate = TaxTrans.TransDate
      RTaxBill.TransRec = NextRecord&
      Put TBHandle, x, RTaxBill
      If FirstTrans = 0 Then FirstTrans = NextRecord&
      LastTrans = NextRecord&
      Put TTHandle, NextRecord&, TaxTrans
      TaxTrans.Amount = TaxTrans.Amount
      'Update the Customer Pointers Now
      Get TCHandle, RTaxBill.CustRec, TaxCust
      If TaxCust.LastTrans = 0 Then
        TaxCust.LastTrans = NextRecord&
        Put TCHandle, RTaxBill.CustRec, TaxCust
      Else
        Previous& = TaxCust.LastTrans
        TaxCust.LastTrans = NextRecord&
        Put TCHandle, RTaxBill.CustRec, TaxCust
        
        Get TTHandle, NextRecord&, TaxTrans
        TaxTrans.LastTrans = Previous&
        Put TTHandle, NextRecord&, TaxTrans
      End If
      'Now Update the Property Records with the Tax Year to prevent duplicate billing per year
      If RTaxBill.RealPropRecord > 0 Then
        Get RRHandle, RTaxBill.RealPropRecord, RealRec
        RealRec.LastYrPrinted = RTaxBill.TaxYear
        RealRec.PROPDISC = "N"
        RealRec.LateList = "N"
        Put RRHandle, RTaxBill.RealPropRecord, RealRec
      End If
      
      If ROverPayYN = True Then
        For y = 1 To NumOfOPRecs
          Get OPHandle, y, OPBillRec
          If OPBillRec.BelongTo = x Then
            GoSub RealOverPay
            Exit For
          End If
        Next y
      End If
NotARealBill:
      frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
      If frmVATaxShowPctComp.Out = True Then
        Close
        frmVATaxShowPctComp.Out = False
        Unload frmVATaxShowPctComp
        EnableCloseButton Me.hwnd, True
        Exit Sub
      End If
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  Close
  
  Call Savemsg(900, "Posting Real Tax Bills has completed successfully.")
  
  KillFile "mortx" + ThisYear + ".dat"
  KillFile "frombillpost.dat"
  KillFile "txrblsprn.dat"
  KillFile RealTaxBillOPFile
  KillFile RealTaxBillInfoFile
  KillFile "RZIPIDX.DAT"
  KillFile "MORTIDX.DAT"
  
  If Exist(FileName) Then
    If TaxMsgWOpts(800, "The backup file " + FileName + " already exists. Press F10 to overwrite. Otherwise, press ESC to keep the old file.", "F10 Overwrite", "ESC Keep Old") = "abort" Then
      Unload frmVATaxMsgWOpts
      MainLog ("WARNING: User warned that the file " + FileName + " already exists and they elected to not overwrite that file with a new post.")
    Else
      Unload frmVATaxMsgWOpts
      KillFile FileName
      Name RealTaxBillFile As FileName '5.16.07 changed to make local
      Call TaxMsg(900, "The backup file " + FileName + " has been saved in the Citipak directory.")
    End If
  Else
    Name RealTaxBillFile As FileName '5.16.07 changed to make local
    Call TaxMsg(900, "The backup file " + FileName + " has been saved in the Citipak directory.")
  End If
    
  GoSub SaveToTaxBillPostFile
  Exit Sub
  
RealOverPay:
  TotalPaid# = 0
  OverPayAmt = OPBillRec.Amount
'      PayTranRec = EmptyPay
'make a new clean payment trans
  TotalPaid# = OldRound#(OPBillRec.Revenue.Principle1Pd + OPBillRec.Revenue.LateListPd)  '4/22/05
  TotalPaid# = OldRound(TotalPaid# + OPBillRec.Revenue.RevOpt1Pd + OPBillRec.Revenue.RevOpt2Pd + OPBillRec.Revenue.RevOpt3Pd)
  If TotalPaid# = 0 Then
    GoTo SkipThisRealRec
  End If
      'PayTranRec = the new record for tax transaction records
  PayTranRec.TransDate = Date2Num%(Date$)
  PayTranRec.TranType = 9
'      PayTranRec.Revenue.Principle1Pd = OldRound(PayListRec.Principle1 + PayListRec.DiscAmt)
  PayTranRec.Revenue.Principle1Pd = OPBillRec.Revenue.Principle1Pd
  PayTranRec.Revenue.InterestPd = 0
  PayTranRec.Revenue.CollectionPd = 0
  PayTranRec.Revenue.LateListPd = OPBillRec.Revenue.LateListPd
  PayTranRec.Revenue.RevOpt1Pd = OPBillRec.Revenue.RevOpt1Pd
  PayTranRec.Revenue.RevOpt2Pd = OPBillRec.Revenue.RevOpt2Pd
  PayTranRec.Revenue.RevOpt3Pd = OPBillRec.Revenue.RevOpt3Pd
  PayTranRec.CustPin = TaxCust.PIN
  PayTranRec.DiscXDate = TaxTrans.DiscXDate
  PayTranRec.RealPin = QPTrim$(TaxTrans.RealPin)
  PayTranRec.PersPin = ""
  PayTranRec.BillType = "R"
  PayTranRec.Posted2GL = "N"
  PayTranRec.TaxYear = TaxTrans.TaxYear
  PayTranRec.DiscAmt = 0
  PayTranRec.OperNum = OperNum
  PayTranRec.Amount = 0  'OverPayAmt'TotalPaid#'9/9/05
  PayTranRec.FromPrePay = TotalPaid#
  PayTranRec.Description = "Credit Applied to Bill# " + Str$(RTaxBill.BillNumber)
  PayTranRec.CustomerRec = TaxTrans.CustomerRec
  PayTranRec.LastTrans = TaxCust.LastTrans
  PayTranRec.BelongTo = NextRecord& 'OPBillRec.BelongTo
  PayTranRec.Revenue.PrePaidAmt = 0
  PayTranRec.Revenue.PrePaidUsed = OverPayAmt
  PayTranRec.Revenue.PrePaidBal = OldRound(GetOverPayBalance(RTaxBill.CustRec, "R") - OverPayAmt)
  PayTranRec.InternalPin = TaxTrans.InternalPin
  PayTranRec.CntyPara = ""
  PayTranRec.CyclPara = ""
  PayTranRec.TShpPara = ""
  PayTranRec.PPTRADisc = 0
  PayTranRec.PPTRAVal = 0
  PayTranRec.PPTRARmvl = 0
  PayTranRec.PPTRARmvlDate = 0
  Get TTHandle, NextRecord&, TaxTrans
    TaxTrans.Revenue.Principle1Pd = OldRound#(TaxTrans.Revenue.Principle1Pd + PayTranRec.Revenue.Principle1Pd) '4/22/05
    TaxTrans.Revenue.InterestPd = OldRound#(TaxTrans.Revenue.InterestPd + PayTranRec.Revenue.Interest)
    TaxTrans.Revenue.CollectionPd = OldRound#(TaxTrans.Revenue.CollectionPd + PayTranRec.Revenue.Collection)
    TaxTrans.Revenue.LateListPd = OldRound#(TaxTrans.Revenue.LateListPd + PayTranRec.Revenue.LateListPd)
    TaxTrans.Revenue.RevOpt1Pd = OldRound#(TaxTrans.Revenue.RevOpt1Pd + PayTranRec.Revenue.RevOpt1Pd)
    TaxTrans.Revenue.RevOpt2Pd = OldRound#(TaxTrans.Revenue.RevOpt2Pd + PayTranRec.Revenue.RevOpt2Pd)
    TaxTrans.Revenue.RevOpt3Pd = OldRound#(TaxTrans.Revenue.RevOpt3Pd + PayTranRec.Revenue.RevOpt3Pd)
  Put TTHandle, NextRecord&, TaxTrans
  
  NextRecord& = NextRecord& + 1

  Put TTHandle, NextRecord&, PayTranRec
  
  TaxCust.LastTrans = NextRecord&
  Put TCHandle, RTaxBill.CustRec, TaxCust
SkipThisRealRec:
  Return

PostPersonal:
  If TaxMsgWOpts(900, "READY TO POST PERSONAL TAX BILLS? Press F10 to continue. Otherwise, press ESC to Exit.", "F10 POST", "ESC EXIT") = "abort" Then
    Close
    Exit Sub
  End If
  FirstTrans = 0
  LastTrans = 0
  
  OpenPersBillInfoFile BIHandle '8/9/2006 moved to here from just above ThisYear =
  Get BIHandle, 1, PBillInfo
  Close BIHandle
  OpenPersTaxBillFile TBHandle, NumOfTBRecs
  OpenTaxPersFile PRHandle, NumOfPRRecs
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  
  TaxMasterRec.DiscPXDate = PBillInfo.XDate '9/20/05
  TaxMasterRec.PTaxYear = PBillInfo.TaxYear '9/20/05
  Put TMHandle, 1, TaxMasterRec '9/20/05
  Close TMHandle
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  If Not Exist("txpblsprn.dat") Then
    Call TaxMsg(900, "ERROR: Please print personal tax bills before attempting to post.")
    Exit Sub
  End If
    
  POverPayYN = False
  If Exist(PersTaxBillOPFile) Then
    POverPayYN = True
  End If
  
  ThisYear = CStr(PBillInfo.TaxYear)
  DupCnt = CountReprintFiles("POSTP" + Mid(PBillInfo.CountyPara, 1, 3) + Mid(PBillInfo.CyclePara, 1, 3) + Mid(PBillInfo.TwnShpPara, 1, 3) + ThisYear)
  FileName = "TAXBILLBU\POSTP" + Mid(PBillInfo.CountyPara, 1, 3) + Mid(PBillInfo.CyclePara, 1, 3) + Mid(PBillInfo.TwnShpPara, 1, 3) + ThisYear + CStr(DupCnt) + ".DAT"
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  TaxMasterRec.DiscPXDate = PBillInfo.XDate '9/20/05
  TaxMasterRec.RTaxYear = PBillInfo.TaxYear '9/20/05
  Put TMHandle, 1, TaxMasterRec '9/20/05
  Close TMHandle

  frmVATaxShowPctComp.Label1 = "Posting Personal Tax Bills"
  frmVATaxShowPctComp.Show , Me
  frmVATaxShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  If POverPayYN = True Then
    OpenPersTaxBillOverPayFile OPHandle, NumOfOPRecs
  End If
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, PTaxBill
      If PTaxBill.SetDscvry2No = "Y" Then 'added loop 12/4/2006
        Get TCHandle, PTaxBill.CustRec, TaxCust
        NextRec = TaxCust.FirstPersRec
        Do While NextRec > 0
          Get PRHandle, NextRec, PersRec
          PersRec.DISCOV = "N"
          Put PRHandle, NextRec, PersRec
          NextRec = PersRec.NextRec
        Loop
      End If
      If PTaxBill.BillPrinted = False Then GoTo NotAPersBill
      Get TCHandle, PTaxBill.CustRec, TaxCust
      'Update the Transaction File First
      NextRec = TaxCust.FirstPersRec 'added loop 12/4/2006
      Do While NextRec > 0
        Get PRHandle, NextRec, PersRec
        PersRec.DISCOV = "N"
        Put PRHandle, NextRec, PersRec
        NextRec = PersRec.NextRec
      Loop
      TaxTrans.TransDate = Date2Num%(Date$)
      TaxTrans.TaxYear = PTaxBill.TaxYear
      TaxTrans.TranType = 1  '1=Bill 2=Payment 3=Release 4=Interest 5=Penalty 6=Collection/Ad Cost Billing 7=Adjustment 9=Credit Applied at Billing
      TaxTrans.BillType = "P"
      TaxTrans.Amount = PTaxBill.TotalBillDue      'Total Transaction Amount
      TaxTrans.Revenue.Principle1 = PTaxBill.PersTaxDue
      TaxTrans.Revenue.Principle2 = PTaxBill.MTTaxDue
      TaxTrans.Revenue.Principle3 = PTaxBill.MCTaxDue
      TaxTrans.Revenue.Principle4 = PTaxBill.FETaxDue
      TaxTrans.Revenue.Principle5 = PTaxBill.MHTaxDue
      TaxTrans.Revenue.Interest = 0
      TaxTrans.Revenue.Penalty = 0
      TaxTrans.Revenue.Collection = 0
      TaxTrans.Revenue.Future1 = 0
      TaxTrans.Revenue.Future2 = 0
      TaxTrans.Revenue.Principle1Pd = 0
      TaxTrans.Revenue.Principle2Pd = 0
      TaxTrans.Revenue.Principle3Pd = 0
      TaxTrans.Revenue.Principle4Pd = 0
      TaxTrans.Revenue.Principle5Pd = 0
      TaxTrans.Revenue.InterestPd = 0
      TaxTrans.Revenue.PenaltyPd = 0
      TaxTrans.Revenue.CollectionPd = 0
      TaxTrans.Revenue.Future1Pd = 0
      TaxTrans.Revenue.Future2Pd = 0
      TaxTrans.Revenue.RevOpt1 = PTaxBill.OptRevTax1
      TaxTrans.Revenue.RevOpt1Pd = 0
      TaxTrans.Revenue.RevOpt2 = PTaxBill.OptRevTax2
      TaxTrans.Revenue.RevOpt2Pd = 0
      TaxTrans.Revenue.RevOpt3 = PTaxBill.OptRevTax3
      TaxTrans.Revenue.RevOpt3Pd = 0
      TaxTrans.Revenue.LateList = PTaxBill.LateTaxDue
      TaxTrans.Revenue.LateListPd = 0
      TaxTrans.Revenue.PrePaidAmt = 0
      TaxTrans.Revenue.PrePaidUsed = 0
      TaxTrans.Revenue.PrePaidBal = GetOverPayBalance(PTaxBill.CustRec, "P")
      TaxTrans.InternalPin = PTaxBill.InternalPin
      TaxTrans.Revenue.pad = ""
    
      TaxTrans.Description = "Tax Bill #" + Str$(PTaxBill.BillNumber)
      TaxTrans.Posted2GL = "N"
      TaxTrans.CustomerRec = PTaxBill.CustRec
      TaxTrans.LastTrans = 0
      TaxTrans.BelongTo = 0
      TaxTrans.Padding = ""
      TaxTrans.PersPin = QPTrim$(PTaxBill.PersPin)
      TaxTrans.RealPin = 0
      TaxTrans.CustPin = PTaxBill.CustPin
      TaxTrans.DiscXDate = TaxMasterRec.DiscPXDate
      TaxTrans.DiscAmt = 0
      TaxTrans.OperNum = OperNum
      TaxTrans.PersVal = PTaxBill.PersValue
      TaxTrans.PPTRAVal = PTaxBill.PPTRAValue
      TaxTrans.PPTRADisc = PTaxBill.PPTRADiscnt
      TaxTrans.CntyPara = QPTrim$(PBillInfo.CountyPara)
      TaxTrans.CyclPara = QPTrim$(PBillInfo.CyclePara)
      TaxTrans.TShpPara = QPTrim$(PBillInfo.TwnShpPara)
      TaxTrans.PPTRARmvl = 0
      TaxTrans.PPTRARmvlDate = 0
    'Increment Transaction File Record Count
      NextRecord& = (LOF(TTHandle) / Len(TaxTrans)) + 1
      PTaxBill.PostDate = TaxTrans.TransDate
      PTaxBill.TransRec = NextRecord&
      Put TBHandle, x, PTaxBill
      If FirstTrans = 0 Then FirstTrans = NextRecord&
      LastTrans = NextRecord&
      
      Put TTHandle, NextRecord&, TaxTrans
      TaxTrans.Amount = TaxTrans.Amount
      'Update the Customer Pointers Now
      Get TCHandle, PTaxBill.CustRec, TaxCust
      If TaxCust.LastTrans = 0 Then
        TaxCust.LastTrans = NextRecord&
        Put TCHandle, PTaxBill.CustRec, TaxCust
      Else
        Previous& = TaxCust.LastTrans
        TaxCust.LastTrans = NextRecord&
        Put TCHandle, PTaxBill.CustRec, TaxCust
        
        Get TTHandle, NextRecord&, TaxTrans
        TaxTrans.LastTrans = Previous&
        Put TTHandle, NextRecord&, TaxTrans
      End If
      'Now Update the Property Records with the Tax Year to prevent duplicate billing per year
      
      If PTaxBill.PersPropRecord > 0 Then
        Get PRHandle, PTaxBill.PersPropRecord, PersRec
        PersRec.LastYrPrinted = PTaxBill.TaxYear
        PersRec.DISCOV = "N"
        Put PRHandle, PTaxBill.PersPropRecord, PersRec
      End If
     
      If POverPayYN = True Then
        For y = 1 To NumOfOPRecs
          Get OPHandle, y, OPBillRec
          If OPBillRec.BelongTo = x Then
            GoSub PersOverPay
            Exit For
          End If
        Next y
      End If
NotAPersBill:
      frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
      If frmVATaxShowPctComp.Out = True Then
        Close
        frmVATaxShowPctComp.Out = False
        Unload frmVATaxShowPctComp
        EnableCloseButton Me.hwnd, True
        Exit Sub
      End If
  Next x
  
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  Close
  
  Call Savemsg(900, "Posting Personal Tax Bills has completed successfully.")
  
'  KillFile "mortx" + ThisYear + ".dat"
  KillFile "frombillpost.dat"
  KillFile "txpblsprn.dat"
  KillFile PersTaxBillOPFile
  KillFile PersTaxBillInfoFile
  KillFile "PZIPIDX.DAT"
  
  If Exist(FileName) Then
    If TaxMsgWOpts(800, "The backup file " + FileName + " already exists. Press F10 to overwrite. Otherwise, press ESC to keep the old file.", "F10 Overwrite", "ESC Keep Old") = "abort" Then
      Unload frmVATaxMsgWOpts
      MainLog ("WARNING: User warned that the file " + FileName + " already exists and they elected to not overwrite that file with a new post.")
    Else
      Unload frmVATaxMsgWOpts
      KillFile FileName
      Name PersTaxBillFile As FileName '5.16.07 changed to make local
      Call TaxMsg(900, "The backup file " + FileName + " has been saved in the Citipak directory.")
    End If
  Else
    Name PersTaxBillFile As FileName '5.16.07 changed to make local
    Call TaxMsg(900, "The backup file " + FileName + " has been saved in the Citipak directory.")
  End If
    
  GoSub SaveToTaxBillPostFile

  Exit Sub
  
PersOverPay:
  TotalPaid# = 0
  OverPayAmt = OPBillRec.Amount
'      PayTranRec = EmptyPay
'make a new clean payment trans
  TotalPaid# = OldRound#(OPBillRec.Revenue.Principle1Pd + OPBillRec.Revenue.Principle2Pd + OPBillRec.Revenue.Principle3Pd)
  TotalPaid = OldRound(TotalPaid# + OPBillRec.Revenue.Principle4Pd + OPBillRec.Revenue.Principle5Pd + OPBillRec.Revenue.LateListPd)
  TotalPaid = OldRound(TotalPaid + OPBillRec.Revenue.RevOpt1Pd + OPBillRec.Revenue.RevOpt2Pd + OPBillRec.Revenue.RevOpt3Pd) 'added this line on 10/17/08
  If TotalPaid# = 0 Then
    GoTo SkipThisPersRec
  End If
      'PayTranRec = the new record for tax transaction records
  PayTranRec.TransDate = Date2Num%(Date$)
  PayTranRec.TranType = 9
'      PayTranRec.Revenue.Principle1Pd = OldRound(PayListRec.Principle1 + PayListRec.DiscAmt)
  PayTranRec.Revenue.Principle1Pd = OPBillRec.Revenue.Principle1Pd
  PayTranRec.Revenue.Principle2Pd = OPBillRec.Revenue.Principle2Pd
  PayTranRec.Revenue.Principle3Pd = OPBillRec.Revenue.Principle3Pd
  PayTranRec.Revenue.Principle4Pd = OPBillRec.Revenue.Principle4Pd
  PayTranRec.Revenue.Principle5Pd = OPBillRec.Revenue.Principle5Pd
  PayTranRec.Revenue.InterestPd = 0
  PayTranRec.Revenue.CollectionPd = 0
  PayTranRec.Revenue.LateListPd = OPBillRec.Revenue.LateListPd
  PayTranRec.Revenue.RevOpt1Pd = OPBillRec.Revenue.RevOpt1Pd
  PayTranRec.Revenue.RevOpt2Pd = OPBillRec.Revenue.RevOpt2Pd
  PayTranRec.Revenue.RevOpt3Pd = OPBillRec.Revenue.RevOpt3Pd
  PayTranRec.CustPin = TaxCust.PIN
  PayTranRec.DiscXDate = TaxTrans.DiscXDate
  PayTranRec.RealPin = ""
  PayTranRec.PersPin = QPTrim$(TaxTrans.PersPin)
  PayTranRec.Posted2GL = "N"
  PayTranRec.TaxYear = TaxTrans.TaxYear
  PayTranRec.DiscAmt = 0
  PayTranRec.OperNum = OperNum
  PayTranRec.BillType = "P"
  PayTranRec.Amount = 0  'OverPayAmt'TotalPaid#'9/9/05
  PayTranRec.FromPrePay = TotalPaid#
  PayTranRec.Description = "Credit Applied to Bill# " + Str$(PTaxBill.BillNumber)
  PayTranRec.CustomerRec = TaxTrans.CustomerRec
  PayTranRec.LastTrans = TaxCust.LastTrans
  PayTranRec.BelongTo = NextRecord& 'OPBillRec.BelongTo
  PayTranRec.Revenue.PrePaidAmt = 0
  PayTranRec.Revenue.PrePaidUsed = OverPayAmt
  PayTranRec.Revenue.PrePaidBal = OldRound(GetOverPayBalance(PTaxBill.CustRec, "P") - OverPayAmt)
  PayTranRec.InternalPin = TaxTrans.InternalPin
  PayTranRec.CntyPara = ""
  PayTranRec.CyclPara = ""
  PayTranRec.TShpPara = ""
  PayTranRec.PPTRADisc = 0
  PayTranRec.PPTRAVal = 0
  PayTranRec.PPTRARmvl = 0
  PayTranRec.PPTRARmvlDate = 0
  Get TTHandle, NextRecord&, TaxTrans
    TaxTrans.Revenue.Principle1Pd = OldRound#(TaxTrans.Revenue.Principle1Pd + PayTranRec.Revenue.Principle1Pd)
    TaxTrans.Revenue.Principle2Pd = OldRound#(TaxTrans.Revenue.Principle2Pd + PayTranRec.Revenue.Principle2Pd)
    TaxTrans.Revenue.Principle3Pd = OldRound#(TaxTrans.Revenue.Principle3Pd + PayTranRec.Revenue.Principle3Pd)
    TaxTrans.Revenue.Principle4Pd = OldRound#(TaxTrans.Revenue.Principle4Pd + PayTranRec.Revenue.Principle4Pd)
    TaxTrans.Revenue.Principle5Pd = OldRound#(TaxTrans.Revenue.Principle5Pd + PayTranRec.Revenue.Principle5Pd)
    TaxTrans.Revenue.InterestPd = OldRound#(TaxTrans.Revenue.InterestPd + PayTranRec.Revenue.Interest)
    TaxTrans.Revenue.CollectionPd = OldRound#(TaxTrans.Revenue.CollectionPd + PayTranRec.Revenue.Collection)
    TaxTrans.Revenue.LateListPd = OldRound#(TaxTrans.Revenue.LateListPd + PayTranRec.Revenue.LateListPd)
    TaxTrans.Revenue.RevOpt1Pd = OldRound#(TaxTrans.Revenue.RevOpt1Pd + PayTranRec.Revenue.RevOpt1Pd)
    TaxTrans.Revenue.RevOpt2Pd = OldRound#(TaxTrans.Revenue.RevOpt2Pd + PayTranRec.Revenue.RevOpt2Pd)
    TaxTrans.Revenue.RevOpt3Pd = OldRound#(TaxTrans.Revenue.RevOpt3Pd + PayTranRec.Revenue.RevOpt3Pd)
  Put TTHandle, NextRecord&, TaxTrans
  
  NextRecord& = NextRecord& + 1

  Put TTHandle, NextRecord&, PayTranRec
  
  TaxCust.LastTrans = NextRecord&
  Put TCHandle, PTaxBill.CustRec, TaxCust
  
SkipThisPersRec:

  Return
  
SaveToTaxBillPostFile:
  OpenBillPostDateFile PostHandle, NumOfPostRecs
  TBPostRec.PostDate = Date2Num%(Date$)
  TBPostRec.PostYear = TaxTrans.TaxYear
  TBPostRec.BillType = TaxTrans.BillType
  TBPostRec.BackUpName = FileName
  TBPostRec.FirstTrans = FirstTrans
  TBPostRec.LastTrans = LastTrans
  TBPostRec.PPTRAPosted = "N"
  TBPostRec.pad = ""
  Put PostHandle, NumOfPostRecs + 1, TBPostRec
  Close TBHandle
  
  Return
  

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillingMenu", "cmdPost_Click", Erl)
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

Private Sub cmdPrebill_Click()
  
  If Allow = False Then
    Call TaxMsg(900, "No tax customers have been saved. Access denied.")
    Exit Sub
  End If
  
  frmVATaxPrebilling.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPrintReprint_Click()
'  If Check4PayBatch("P") = True Then
'    frmVATaxUnpostedPaylist.BillType = "P"
'    frmVATaxUnpostedPaylist.Show vbModal
'    Call TaxMsg(800, "An unposted personal payment file is ready for posting. Bill printing cannot be conducted until these personal payments are posted.")
'    Exit Sub
'  End If
  
'  If Check4PayBatch("R") = True Then
'    frmVATaxUnpostedPaylist.BillType = "R"
'    frmVATaxUnpostedPaylist.Show vbModal
'    Call TaxMsg(800, "An unposted real payment file is ready for posting. Bill printing cannot be conducted until these real payments are posted.")
'    Exit Sub
'  End If
  
  If Not Exist(RealTaxBillFile) And Not Exist(PersTaxBillFile) Then
    Call TaxMsg(900, "Please process pre-billing before printing bills.")
    Exit Sub
  End If
  
  frmVATaxBillPrintMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%x"
      Call cmdExit_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  
  Allow = True
  OpenTaxCustFile TCHandle, NumOfTCRecs
  Close TCHandle
  If NumOfTCRecs = 0 Then
    Allow = False
  End If
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpTaxBillingMenu
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxBillingMenu.")
      KillFile "C:\CPWork\lateltr.dat"
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

Private Sub cmdReprintPosted_Click()
  Dim x As Integer
  Dim ThisYear As Integer
  Dim ThisFile$
  Dim GotIt As Boolean
  Dim MyFile$, MyPath$, MyName$
  
  GotIt = False
  ThisYear = 1979
  MyPath = StartPath + "\TAXBILLBU\"
  MyName$ = Dir(MyPath, vbDirectory)
  frmVATaxBillPostOpt.Show vbModal
  If frmVATaxBillPostOpt.fptxtPostType.Text = "Real" Then
    Do While MyName <> ""
      MyName = Dir
      If Len(MyName) > 4 Then
        If Mid(MyName, 5, 1) = "R" Then
          GotIt = True
          Exit Do
        End If
      End If
    Loop
    If GotIt = True Then
      Real = True
      frmVATaxReprintPosted.Show
      DoEvents
      Unload Me
    Else
      Call TaxMsg(900, "There are no posted real tax bill files saved at this time.")
      Exit Sub
    End If
  ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Exit" Then
    DoEvents
    Unload frmVATaxBillPostOpt
    Exit Sub
  Else
    Do While MyName <> ""
      MyName = Dir
      If Len(MyName) > 4 Then
        If Mid(MyName, 5, 1) = "P" Then
          GotIt = True
          Exit Do
        End If
      End If
    Loop
    If GotIt = True Then
      Real = False
      frmVATaxReprintPosted.Show
      DoEvents
      Unload Me
    Else
      Call TaxMsg(900, "There are no posted personal tax bill files saved at this time.")
      Exit Sub
    End If
  End If

End Sub

