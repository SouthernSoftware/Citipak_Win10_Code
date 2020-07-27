VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmTaxBillingMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Billing Menu"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTaxBillingMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdInterest 
      Height          =   435
      Left            =   4005
      TabIndex        =   4
      Top             =   4845
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
      ButtonDesigner  =   "frmTaxBillingMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPost 
      Height          =   435
      Left            =   4005
      TabIndex        =   3
      Top             =   4290
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
      ButtonDesigner  =   "frmTaxBillingMenu.frx":0AAC
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintReprint 
      Height          =   435
      Left            =   4005
      TabIndex        =   1
      Tag             =   "0"
      Top             =   3180
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
      ButtonDesigner  =   "frmTaxBillingMenu.frx":0C8E
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrebill 
      Height          =   435
      Left            =   4005
      TabIndex        =   0
      Top             =   2610
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
      ButtonDesigner  =   "frmTaxBillingMenu.frx":0E79
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdLateNotice 
      Height          =   435
      Left            =   4005
      TabIndex        =   5
      Top             =   5400
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
      ButtonDesigner  =   "frmTaxBillingMenu.frx":106B
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   435
      Left            =   4005
      TabIndex        =   7
      Top             =   7110
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
      ButtonDesigner  =   "frmTaxBillingMenu.frx":1251
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMortExport 
      Height          =   435
      Left            =   4005
      TabIndex        =   2
      Top             =   3735
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
      ButtonDesigner  =   "frmTaxBillingMenu.frx":142E
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdReprintPosted 
      Height          =   450
      Left            =   4005
      TabIndex        =   6
      Tag             =   "0"
      Top             =   5970
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   794
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
      ButtonDesigner  =   "frmTaxBillingMenu.frx":1622
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClear 
      Height          =   450
      Left            =   4005
      TabIndex        =   9
      Tag             =   "0"
      Top             =   6540
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   794
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
      ButtonDesigner  =   "frmTaxBillingMenu.frx":180E
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
      TabIndex        =   8
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
      Top             =   2027
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
      Top             =   813
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
Attribute VB_Name = "frmTaxBillingMenu"
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

Private Sub cmdClear_Click()
  Dim BillInfo As TaxBillInfoType
  Dim BIHandle As Integer
  Dim ThisYear$
  
  If Not Exist(TaxBillFile) Then
    Call TaxMsg(900, "No billing files currently exist. Delete attempt aborted.")
    Exit Sub
  End If
  
  OpenBillInfoFile BIHandle
  Get BIHandle, 1, BillInfo
  Close BIHandle
  
  ThisYear = CStr(BillInfo.TaxYear)
  
  If TaxMsgWOpts(700, "WARNING: IF YOU CHOOSE TO CONTINUE THEN ALL UNPOSTED BILLING FILES WILL BE REMOVED PERMANENTLY. IF YOU WISH TO CONTINUE THEN PRESS F10. OTHERWISE PRESS ESC TO LEAVE UNPOSTED BILLING FILES UNCHANGED.", "F10 Delete", "ESC Abort") = "abort" Then
    Exit Sub
  Else
    KillFile "mortx" + ThisYear + ".dat"
    KillFile "txblsprn.dat"
    KillFile TaxBillOPFile
    KillFile TaxBillFile
    MainLog ("User deleted unposted billing files after being warned about the consequences.")
    Call Savemsg(900, "All unposted billing files have been deleted successfully.")
  End If
End Sub

Private Sub cmdExit_Click()
  KillFile "C:\CPWork\lateltr.dat"
  frmTaxMainMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdInterest_Click()
  If Check4PayBatch = True Then
    frmTaxUnpostedPayList.Show vbModal
'    Call TaxMsg(800, "An unposted payment file is ready for posting. Interest calculations cannot be conducted until these payments are posted.")
    Exit Sub
  End If
 
  If Allow = False Then
    Call TaxMsg(900, "No tax customers have been saved. Access denied.")
    Exit Sub
  End If
  frmTaxInterestMenu.Show
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
      Unload frmTaxMsgWOpts
      Exit Sub
    Else
      Unload frmTaxMsgWOpts
      One = 1
      AHandle = FreeFile
      Open "C:\CPWork\lateltr.dat" For Output As AHandle
      Print #AHandle, One
      Close AHandle
      frmTaxSystemSetup.Show
      DoEvents
      Me.Hide
      Exit Sub
    End If
  End If
    
  If Allow = False Then
    Call TaxMsg(900, "No tax customers have been saved. Access denied.")
    Exit Sub
  End If
  
  frmTaxLateNoticeMenu.Show
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
  
  If Not Exist("txblsprn.dat") Then
    Call TaxMsg(800, "Please print tax bills before creating mortgage company export files.")
    Close
    Exit Sub
  End If
  
  frmTaxMortgageExport.Show
  DoEvents
  Unload Me
    
End Sub

Public Sub cmdPost_Click()
  Dim TaxBill As TaxBillType
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
  Dim BillInfo As TaxBillInfoType
  Dim BIHandle As Integer
  Dim x As Long, NextRecord&
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim Previous&
  Dim OverPayYN As Boolean, y As Long
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
  Dim DupCnt As Integer 'added 9/15/06
  Dim NextRec As Long 'added 12/5/06
  
'  If Check4PayBatch = True Then
'    frmTaxUnpostedPayList.Show vbModal
'    Call TaxMsg(800, "An unposted payment file is ready for posting. Bill posting cannot be conducted until these payments are posted.")
'    Exit Sub
'  End If
 
  'on error goto ERRORSTUFF
  
  OverPayYN = False
  If Exist(TaxBillOPFile) Then
    OverPayYN = True
  End If
  
  If Not Exist("txblsprn.dat") Then
    Call TaxMsg(900, "ERROR: Please print tax bills before attempting to post.")
    Exit Sub
  End If
  
  OpenMortCodeFile MCHandle, NumOfMCRecs
  Close MCHandle
  
  OpenBillInfoFile BIHandle
  Get BIHandle, 1, BillInfo
  Close BIHandle
  
  ThisYear = CStr(BillInfo.TaxYear)
'  FileName = "TAXBILLBU\POSTT" + Mid(BillInfo.CountyPara, 1, 3) + Mid(BillInfo.CyclePara, 1, 3) + Mid(BillInfo.TwnShpPara, 1, 3) + ThisYear + ".DAT"
  DupCnt = CountReprintFiles("POSTT" + Mid(BillInfo.CountyPara, 1, 3) + Mid(BillInfo.CyclePara, 1, 3) + Mid(BillInfo.TwnShpPara, 1, 3) + ThisYear) 'added 9/15/06
  FileName = "TAXBILLBU\POSTT" + Mid(BillInfo.CountyPara, 1, 3) + Mid(BillInfo.CyclePara, 1, 3) + Mid(BillInfo.TwnShpPara, 1, 3) + ThisYear + CStr(DupCnt) + ".DAT"
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  TaxMasterRec.DiscXDate = BillInfo.XDate '9/20/05
  TaxMasterRec.TaxYear = BillInfo.TaxYear '9/20/05
  Put TMHandle, 1, TaxMasterRec '9/20/05
  Close TMHandle
  
  If MortOK = True Then
    Unload frmTaxMortgageExport
    GoTo SkipMort
  End If
  If NumOfMCRecs > 0 And Not Exist("mortx" + ThisYear + ".dat") Then
    If TaxMsgWOpts(600, "WARNING: If you wish to create mortgage company tax bill export files then please do so before continuing with this post procedure. If you wish to create the mortgage company tax bill export files then press F10. Otherwise, press ESC to continue posting.", "F10 Make Files", "ESC Continue Post") = "abort" Then
      MainLog ("WARNING: User warned that mortgage company export files should be created before posting tax bills but they elected to post without creating export files.")
      Unload frmTaxMsgWOpts
    Else
      Unload frmTaxMsgWOpts
      One = 1
      AHandle = FreeFile
      Open "frombillpost.dat" For Output As AHandle
      Print #AHandle, One
      Close AHandle
      frmTaxMortgageExport.Show
      MortOK = True
      DoEvents
      Exit Sub
    End If
  End If
  
SkipMort:
  MortOK = False
  If TaxMsgWOpts(900, "READY TO POST? Press F10 to continue. Otherwise, press ESC to Exit.", "F10 POST", "ESC EXIT") = "abort" Then
    Close
    Exit Sub
  End If
  
  OpenTaxBillFile TBHandle, NumOfTBRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxPropFile RRHandle, NumOfRRREcs
  OpenTaxPersFile PRHandle, NumOfPRRecs
'  OpenBillInfoFile BIHandle'moved to above on 9/20/05
'  Get BIHandle, 1, BillInfo
'  Close BIHandle
'  TaxMasterRec.DiscXDate = BillInfo.XDate '9/20/05
'  TaxMasterRec.TaxYear = BillInfo.TaxYear '9/20/05
'  Put TMHandle, 1, TaxMasterRec '9/20/05
'  Close TMHandle
  
  frmTaxShowPctComp.Label1 = "Posting Tax Bills"
  frmTaxShowPctComp.Show , Me
  frmTaxShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  If OverPayYN = True Then
    OpenTaxBillOverPayFile OPHandle, NumOfOPRecs
  End If
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, TaxBill
         If TaxBill.SetDscvry2No = "Y" Then 'added if on 12/5/06
          If TaxBill.RealPropRecord > 0 Then
            Get RRHandle, TaxBill.RealPropRecord, RealRec
            RealRec.PROPDISC = "N"
            RealRec.LateList = "N"
            Put RRHandle, TaxBill.RealPropRecord, RealRec
          End If
          
          Get TCHandle, TaxBill.CustRec, TaxCust
          NextRec = TaxCust.FirstPersRec
          Do While NextRec > 0
            Get PRHandle, NextRec, PersRec
            PersRec.DISCOV = "N"
            PersRec.LateList = "N"
            Put PRHandle, NextRec, PersRec
            NextRec = PersRec.NextRec
          Loop
        End If
       If TaxBill.BillPrinted = False Then GoTo Skip
       Get TCHandle, TaxBill.CustRec, TaxCust
       NextRec = TaxCust.FirstPersRec 'added loop 12/5/2006
       Do While NextRec > 0
         Get PRHandle, NextRec, PersRec
         PersRec.DISCOV = "N"
         PersRec.LateList = "N"
         Put PRHandle, NextRec, PersRec
         NextRec = PersRec.NextRec
       Loop
      'Update the Transaction File First
        TaxTrans.TransDate = Date2Num%(Date$)
        TaxTrans.TaxYear = TaxBill.TaxYear
        TaxTrans.TranType = 1  '1=Bill 2=Payment 3=Release 4=Interest 5=Penalty 6=Collection/Ad Cost Billing 7=Adjustment 9=Credit Applied at Billing
        If InStr(BillInfo.SplitPara, "REAL") Then
          TaxTrans.BillType = "R"
        ElseIf InStr(BillInfo.SplitPara, "PERSONAL") Then
          TaxTrans.BillType = "P"
        Else
          TaxTrans.BillType = "C"         'R=Real P=Personal Property C=Combined (NC/GA)
        End If
        TaxTrans.Amount = TaxBill.TotalBillDue      'Total Transaction Amount
      
        TaxTrans.Revenue.Principle1 = OldRound(TaxBill.TotalBillDue - (TaxBill.OptRevTax1 + TaxBill.OptRevTax2 + TaxBill.OptRevTax3 + TaxBill.LateTaxDue))
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
        TaxTrans.Revenue.RevOpt1 = TaxBill.OptRevTax1
        TaxTrans.Revenue.RevOpt1Pd = 0
        TaxTrans.Revenue.RevOpt2 = TaxBill.OptRevTax2
        TaxTrans.Revenue.RevOpt2Pd = 0
        TaxTrans.Revenue.RevOpt3 = TaxBill.OptRevTax3
        TaxTrans.Revenue.RevOpt3Pd = 0
        TaxTrans.Revenue.LateList = TaxBill.LateTaxDue
        TaxTrans.Revenue.LateListPd = 0
        TaxTrans.Revenue.PrePaidAmt = 0
        TaxTrans.Revenue.PrePaidUsed = 0
        TaxTrans.Revenue.PrePaidBal = GetOverPayBalance(TaxBill.CustRec)
        TaxTrans.InternalPin = TaxBill.InternalPin
        TaxTrans.Revenue.pad = ""
    
        TaxTrans.Description = "Tax Bill #" + Str$(TaxBill.BillNumber)
        TaxTrans.Posted2GL = "N"
        TaxTrans.CustomerRec = TaxBill.CustRec
        TaxTrans.LastTrans = 0
        TaxTrans.BelongTo = 0
        TaxTrans.Padding = ""
        TaxTrans.PersPin = QPTrim$(TaxBill.PersPin)
        TaxTrans.RealPin = QPTrim$(TaxBill.RealPin)
        TaxTrans.CustPin = TaxBill.CustPin
        TaxTrans.DiscXDate = TaxMasterRec.DiscXDate
        TaxTrans.DiscAmt = 0
        TaxTrans.OperNum = OperNum
        TaxTrans.CntyPara = QPTrim$(BillInfo.CountyPara)
        TaxTrans.CyclPara = QPTrim$(BillInfo.CyclePara)
        TaxTrans.TShpPara = QPTrim$(BillInfo.TwnShpPara)
      'Increment Transaction File Record Count
        NextRecord& = (LOF(TTHandle) / Len(TaxTrans)) + 1
      
        Put TTHandle, NextRecord&, TaxTrans
        TaxTrans.Amount = TaxTrans.Amount
      'Update the Customer Pointers Now
        Get TCHandle, TaxBill.CustRec, TaxCust
        If TaxCust.LastTrans = 0 Then
          TaxCust.LastTrans = NextRecord&
          Put TCHandle, TaxBill.CustRec, TaxCust
        Else
          Previous& = TaxCust.LastTrans
          TaxCust.LastTrans = NextRecord&
          Put TCHandle, TaxBill.CustRec, TaxCust
        
          Get TTHandle, NextRecord&, TaxTrans
          TaxTrans.LastTrans = Previous&
          Put TTHandle, NextRecord&, TaxTrans
        End If
      'Now Update the Property Records with the Tax Year to prevent duplicate billing per year
        If TaxBill.RealPropRecord > 0 Then
          Get RRHandle, TaxBill.RealPropRecord, RealRec
          RealRec.LastYrPrinted = TaxBill.TaxYear
          RealRec.PROPDISC = "N"
          RealRec.LateList = "N"
          Put RRHandle, TaxBill.RealPropRecord, RealRec
        End If
      
        If TaxBill.PersPropRecord > 0 Then
          Get PRHandle, TaxBill.PersPropRecord, PersRec
          PersRec.LastYrPrinted = TaxBill.TaxYear
          PersRec.DISCOV = "N"
          PersRec.LateList = "N"
          Put PRHandle, TaxBill.PersPropRecord, PersRec
        End If
      
        If OverPayYN = True Then
          For y = 1 To NumOfOPRecs
            Get OPHandle, y, OPBillRec
            If OPBillRec.BelongTo = x Then
'            If OPBillRec.BelongTo = TaxBill.BillNumber Then
              GoSub OverPay
              Exit For
            End If
          Next y
        End If
Skip:
    frmTaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  Close
  
  Call Savemsg(900, "Posting Tax Bills has completed successfully.")
  
  KillFile "mortx" + ThisYear + ".dat"
  KillFile "frombillpost.dat"
  KillFile "txblsprn.dat"
  KillFile TaxBillOPFile
  KillFile "ZIPIDX.DAT"
  KillFile "MORTIDX.DAT"
  
  If Exist(FileName) Then
    If TaxMsgWOpts(800, "The backup file " + FileName + " already exists. Press F10 to overwrite. Otherwise, press ESC to keep the old file.", "F10 Overwrite", "ESC Keep Old") = "abort" Then
      Unload frmTaxMsgWOpts
      MainLog ("WARNING: User warned that the file " + FileName + " already exists and they elected to not overwrite that file with a new post.")
    Else
      KillFile FileName
      Unload frmTaxMsgWOpts
      Name "TAXTBILL.DAT" As FileName 'added 8/17/06
      Call TaxMsg(800, "The backup file " + FileName + " has been saved in the Citipak directory.") 'added 8/17/06
    End If
  Else 'added this else on 8/17/06
    Name "TAXTBILL.DAT" As FileName
    Call TaxMsg(800, "The backup file " + FileName + " has been saved in the Citipak directory.")
  End If
    
  Exit Sub
  
OverPay:
  TotalPaid# = 0
  OverPayAmt = OPBillRec.Amount
'      PayTranRec = EmptyPay
'make a new clean payment trans
  TotalPaid# = OldRound#(OPBillRec.Revenue.Principle1Pd + OPBillRec.Revenue.LateListPd)  '4/22/05
  TotalPaid# = OldRound(TotalPaid# + OPBillRec.Revenue.RevOpt1Pd + OPBillRec.Revenue.RevOpt2Pd + OPBillRec.Revenue.RevOpt3Pd)
  If TotalPaid# = 0 Then
    GoTo SkipThisRec
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
  PayTranRec.PersPin = QPTrim$(TaxTrans.PersPin)
  PayTranRec.Posted2GL = "N"
  PayTranRec.TaxYear = TaxTrans.TaxYear
  PayTranRec.DiscAmt = 0
  PayTranRec.OperNum = OperNum
  PayTranRec.Amount = 0  'OverPayAmt'TotalPaid#'9/9/05
  PayTranRec.FromPrePay = TotalPaid#
  PayTranRec.Description = "Credit Applied to Bill# " + Str$(TaxBill.BillNumber)
  PayTranRec.CustomerRec = TaxTrans.CustomerRec
  PayTranRec.LastTrans = TaxCust.LastTrans
  PayTranRec.BelongTo = NextRecord& 'OPBillRec.BelongTo
  PayTranRec.Revenue.PrePaidAmt = 0
  PayTranRec.Revenue.PrePaidUsed = OverPayAmt
  PayTranRec.Revenue.PrePaidBal = OldRound(GetOverPayBalance(TaxBill.CustRec) - OverPayAmt)
  PayTranRec.InternalPin = TaxTrans.InternalPin
  PayTranRec.CntyPara = ""
  PayTranRec.CyclPara = ""
  PayTranRec.TShpPara = ""
  Get TTHandle, NextRecord&, TaxTrans
    TaxTrans.Revenue.Principle1Pd = OldRound#(TaxTrans.Revenue.Principle1Pd + PayTranRec.Revenue.Principle1Pd) '4/22/05
    TaxTrans.Revenue.InterestPd = OldRound#(TaxTrans.Revenue.InterestPd + PayTranRec.Revenue.Interest)
    TaxTrans.Revenue.CollectionPd = OldRound#(TaxTrans.Revenue.CollectionPd + PayTranRec.Revenue.Collection)
    TaxTrans.Revenue.LateListPd = OldRound#(TaxTrans.Revenue.LateListPd + PayTranRec.Revenue.LateListPd)
    TaxTrans.Revenue.RevOpt1Pd = OldRound#(TaxTrans.Revenue.RevOpt1Pd + PayTranRec.Revenue.RevOpt1Pd)
    TaxTrans.Revenue.RevOpt2Pd = OldRound#(TaxTrans.Revenue.RevOpt2Pd + PayTranRec.Revenue.RevOpt2Pd)
    TaxTrans.Revenue.RevOpt3Pd = OldRound#(TaxTrans.Revenue.RevOpt3Pd + PayTranRec.Revenue.RevOpt3Pd)
'      TaxTranRec.Revenue.Future1Pd = OldRound#(TaxTranRec.DiscAmt + PayListRec.DiscAmt)
'      TaxTranRec.DiscAmt = OldRound#(TaxTranRec.DiscAmt + PayListRec.DiscAmt)
      
  Put TTHandle, NextRecord&, TaxTrans
  
  NextRecord& = NextRecord& + 1

  Put TTHandle, NextRecord&, PayTranRec
  
  TaxCust.LastTrans = NextRecord&
  Put TCHandle, TaxBill.CustRec, TaxCust
SkipThisRec:
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxBillingMenu", "cmdPost_Click", Erl)
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
  If Check4PayBatch = True Then
    frmTaxUnpostedPayList.Show vbModal
'    Call TaxMsg(800, "An unposted payment file is ready for posting. Prebilling cannot be conducted until these payments are posted.")
    Exit Sub
  End If
  If Allow = False Then
    Call TaxMsg(900, "No tax customers have been saved. Access denied.")
    Exit Sub
  End If
  frmTaxPrebilling.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPrintReprint_Click()
'  If Check4PayBatch = True Then
'    frmTaxUnpostedPayList.Show vbModal
'    Call TaxMsg(800, "An unposted payment file is ready for posting. Bill printing cannot be conducted until these payments are posted.")
'    Exit Sub
'  End If
  
  If Not Exist(TaxBillFile) Then
    Call TaxMsg(900, "Please process pre-billing before printing bills.")
    Exit Sub
  End If
  
  frmTaxBillPrintMenu.Show
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
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxBillingMenu.")
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
  Dim MyPath$
  Dim MyName$
  
  GotIt = False
  MyPath = StartPath + "\TAXBILLBU\"
  MyName$ = Dir(MyPath, vbDirectory)
  Do While MyName <> ""
    MyName = Dir
    If Len(MyName) > 4 Then
      GotIt = True
      Exit Do
    End If
  Loop
  
'  ThisYear = 1979
'  For x = 1 To 51
'    ThisYear = ThisYear + 1
'    ThisFile = StartPath + "\POSTT" + CStr(ThisYear) + ".DAT"
'    If Exist(ThisFile) Then
'      GotIt = True
'      Exit For
'    End If
'  Next x
  
  If GotIt = True Then
    frmTaxReprintPosted.Show
    DoEvents
    Unload Me
  Else
    Call TaxMsg(900, "There are no posted tax bill files saved at this time.")
    Exit Sub
  End If
End Sub

