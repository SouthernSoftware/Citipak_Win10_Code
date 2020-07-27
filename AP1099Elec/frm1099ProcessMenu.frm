VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Begin VB.Form frm1099ProcessMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Federal 1099 Processing"
   ClientHeight    =   8880
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   12195
   Icon            =   "frm1099ProcessMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdSetupPayer 
      Height          =   492
      Left            =   4296
      TabIndex        =   0
      Top             =   2736
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
      ButtonDesigner  =   "frm1099ProcessMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExtract1099 
      Height          =   495
      Left            =   4260
      TabIndex        =   1
      Top             =   3480
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frm1099ProcessMenu.frx":0AB9
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAddNew1099 
      Height          =   492
      Left            =   4296
      TabIndex        =   2
      Top             =   4236
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
      ButtonDesigner  =   "frm1099ProcessMenu.frx":0CA9
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEditView1099 
      Height          =   492
      Left            =   4296
      TabIndex        =   3
      Top             =   4992
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
      ButtonDesigner  =   "frm1099ProcessMenu.frx":0E96
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrint1099Rpt 
      Height          =   492
      Left            =   4296
      TabIndex        =   4
      Top             =   5748
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
      ButtonDesigner  =   "frm1099ProcessMenu.frx":1085
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrint1099Forms 
      Height          =   492
      Left            =   4296
      TabIndex        =   5
      Top             =   6492
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
      ButtonDesigner  =   "frm1099ProcessMenu.frx":126E
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit1099Menu 
      Height          =   492
      Left            =   4296
      TabIndex        =   6
      Top             =   7248
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
      ButtonDesigner  =   "frm1099ProcessMenu.frx":1456
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00FFFFFF&
      Height          =   156
      Left            =   8856
      Top             =   2256
      Width           =   972
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      Height          =   156
      Left            =   2376
      Top             =   2256
      Width           =   972
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   8976
      X2              =   9672
      Y1              =   8304
      Y2              =   8304
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2496
      X2              =   2496
      Y1              =   2424
      Y2              =   8280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   2496
      X2              =   3192
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1099 PROCESSING MENU"
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
      Left            =   3516
      TabIndex        =   7
      Top             =   1440
      Width           =   5292
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1800
      Top             =   1080
      Width           =   8628
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   2
      X1              =   8976
      X2              =   8976
      Y1              =   2424
      Y2              =   8280
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00D0D0D0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1236
      Left            =   1800
      Top             =   936
      Width           =   8628
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00000000&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   2
      Left            =   8856
      Top             =   2112
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2496
      Top             =   2400
      Width           =   732
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5916
      Index           =   2
      Left            =   8976
      Top             =   2400
      Width           =   732
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00D0D0D0&
      Height          =   300
      Left            =   2376
      Top             =   2112
      Width           =   972
   End
End
Attribute VB_Name = "frm1099ProcessMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim VendorIdx As VendorIdxRecType
Dim Vendor As VendorRecType
Dim AP1099 As AP1099RecType
Private Temp_Class As Resize_Class
Dim grpt As Boolean

Private Sub cmdAddNew1099_Click()
  New1099 = True
  frmAddEdit1099.Show
  Unload frm1099ProcessMenu
End Sub

Private Sub cmdEditView1099_Click()
  New1099 = False
  If Exist("ap1099.dat") Then
    frmAddEdit1099.Show
    Unload frm1099ProcessMenu
    frm1099List.Show 1, frmAddEdit1099
  Else
    MsgBox "No Entries To Edit.", vbOKOnly, "No Entries"
  End If
End Sub

Private Sub cmdExtract1099_Click()
  Dim Year As String
  Year$ = QPTrim$(Str$(Val(Right$(Date$, 4)) - 1))
  frmWarning.Label1 = "This operation will extract 1099 information."
  frmWarning.Label6 = "from the Vendor Transaction History."
  frmWarning.Label5.Visible = False
  frmWarning.Label2 = "This operation will overwrite any previous 1099 files!!!"
  frmWarning.Show 1, Me
  Select Case frmWarning.nogo
  Case True  'ok=1 then no don't continue
    Exit Sub
  Case False
    Extract1099Info (Year$)
    MsgBox "Extract 1099 Procedure Complete.", vbOKOnly, "Complete"
  End Select

End Sub

Private Sub cmdPrint1099Forms_Click()
  If Exist("AP1099.DAT") Then
    frmReportOpt.Show 1
    If rptopt = 1 Then
      grpt = True
      Print1099Forms 0
    ElseIf rptopt = 2 Then
      grpt = False
      Print1099Forms 0
    End If
  Else
    MsgBox "1099 File Information Does NOT Exist", vbOKOnly, "No 1099's"
  End If
End Sub

Private Sub cmdPrint1099Rpt_Click()
  If Exist("AP1099.DAT") Then
    frmReportOpt.Show 1
    If rptopt = 1 Then
      Print1099Report
    ElseIf rptopt = 2 Then
      Print1099Report2
    End If
  Else
    MsgBox "1099 File Information Does NOT Exist", vbOKOnly, "No 1099's"
  End If
End Sub

Private Sub cmdSetupPayer_Click()
  frmSetupPayer.Show
  Unload frm1099ProcessMenu
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName$, GLFundLen, GLAcctLen, GLDetLen
  Me.HelpContextID = hlp1099Proc
  cmdExtract1099.HelpContextID = hlpExtract1099
  cmdPrint1099Rpt.HelpContextID = hlpPrint1099Rep
  cmdPrint1099Forms.HelpContextID = hlpPrint1099Forms
End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      cmdExit1099Menu_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
  If cmdExit1099Menu.Enabled = False Then
    Cancel = True
  Else
    If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      ClearInUse PWcnt
      MainLog "Close via A/P 1099"
    End If
  End If
  End If
End Sub

Private Sub cmdExit1099Menu_Click()
  frmAPReportsMenu.Show
  Unload frm1099ProcessMenu
End Sub
Private Sub Extract1099Info(Year$)
  Dim LDate As Integer, HDate As Integer, cnt As Integer
  Dim VendorIdxFile As Integer, NumActiveVendors As Integer
  Dim VendorFile As Integer, NumVRecs As Integer, NextTran As Long
  Dim APLedgerFile As Integer, NumTrans As Long, LdRecLen As Integer
  Dim FRecLen As Integer, Num1099Recs As Integer, Fed1099File As Integer
  Dim V1099Amt As Double, BaseExtAmt As Double
  FrmShowPctComp.Label1 = "Extracting For Year: " + Year$
  FrmShowPctComp.CmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frm1099ProcessMenu
  Dim PRecLen As Integer, Payfile As Integer
  Dim payer As AP1099PayerRecType
  Payfile = FreeFile
  PRecLen = Len(payer)
  Open "APPAYER.DAT" For Random As Payfile Len = PRecLen
  If LOF(1) > 0 Then
    Get #1, 1, payer
    BaseExtAmt = payer.BaseAmt
  Else
    BaseExtAmt = 0
  End If
  Close Payfile
  LDate = DateDiff("d", "12/31/1979", "01/01/" + Year$)
  HDate = DateDiff("d", "12/31/1979", "12/31/" + Year$)

  Dim VendorIdx As VendorIdxRecType
  OpenVendorIdx VendorIdxFile, NumActiveVendors

  Dim Vendor As VendorRecType
  OpenVendorFile VendorFile, NumVRecs
  Dim ApLedger As APLedger81RecType
  LdRecLen = Len(ApLedger)
  OpenAPLedgerFile APLedgerFile, NumTrans&, LdRecLen

  Dim AP1099 As AP1099RecType
  FRecLen = Len(AP1099)
  Open1099File FRecLen, Num1099Recs, Fed1099File

  If Num1099Recs > 0 Then
    Close Fed1099File
    KillFile "AP1099.DAT"
    Open1099File FRecLen, Num1099Recs, Fed1099File
  End If

  For cnt = 1 To NumActiveVendors
    FrmShowPctComp.ShowPctComp cnt, NumActiveVendors
    Get VendorIdxFile, cnt, VendorIdx
    Get VendorFile, VendorIdx.RecNum, Vendor
    'QPrintRC "Searching: " + Vendor.VNAME, 25, 2, -1
    If Vendor.Get1099 = "Y" Then
      'IF INSTR(Vendor.VName, "ROY") > 0 THEN STOP
      NextTran& = Vendor.FrstTran
      If NextTran& > 0 Then
        V1099Amt# = 0
        Do
          Get APLedgerFile, NextTran&, ApLedger
          If ApLedger.TRCode = 1 Then
            If ApLedger.Get1099 = "Y" Then
            If ApLedger.TRDATE >= LDate And ApLedger.TRDATE <= HDate Then
              V1099Amt# = V1099Amt# + ApLedger.Amt
            End If
            End If
          End If
          NextTran& = ApLedger.NextTrans
        Loop Until NextTran& = 0
      '''''' ' If V1099Amt# > 20000 Then
        If V1099Amt# >= BaseExtAmt# Then
          GoSub Send2File
        End If
      End If
    End If
  Next

  Close
  ActivateControls frm1099ProcessMenu
Exit Sub

Send2File:
  Num1099Recs = Num1099Recs + 1

  AP1099.Deleted = 0
  AP1099.RecID = Vendor.Fedid
  AP1099.RecName = Vendor.VNAME
  'add the dba 12-2003
  AP1099.DBA = Vendor.DBA
  'AP1099.RecADDR = ""
'Just to test address line 2 for extract
'  If Len(QPTrim(Vendor.Addr2)) > 0 Then
    AP1099.RecADDR = Vendor.Addr1
    AP1099.RecADDR2 = Vendor.Addr2
'  Else
'    AP1099.RecADDR = ""
'    AP1099.RecADDR2 = Vendor.Addr1
'  End If
  AP1099.RecCSZ = QPTrim$(Vendor.City) + " " + Vendor.State + " " + Vendor.Zip
  AP1099.RecACCT = ""
  AP1099.NOTICE = ""
  AP1099.BOX1 = 0
  AP1099.BOX2 = 0
  AP1099.BOX3 = 0
  AP1099.BOX4 = 0
  AP1099.BOX5 = 0
  AP1099.BOX6 = 0
  AP1099.BOX7 = V1099Amt#
  AP1099.BOX8 = 0
  AP1099.BOX9 = ""
  AP1099.BOX10 = 0
  AP1099.BOX13 = 0
  AP1099.BOX14 = 0
  AP1099.BOX15 = ""
  AP1099.BOX16 = 0
  AP1099.BOX17 = ""
  AP1099.BOX18 = 0
  AP1099.Void = 0
  AP1099.Corrected = 0
  Put Fed1099File, Num1099Recs, AP1099
  Return
End Sub

Private Sub Print1099Report()
  Dim FRecLen As Integer, Num1099Recs As Integer, Fed1099File As Integer
  Dim PRNFile As Integer, ReportFile As String, fmt As String
  Dim Header As String, Page As Integer, r As Integer, lc As Integer
  Dim TotAmtRep As Double, TaxesWH As Double, GrandTotAmtRep As Double
  Dim totforms As Integer, ToPrint As String, VFlag As String
  FrmShowPctComp.Label1 = "Creating 1099 Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frm1099ProcessMenu

 '--Open the 1099 data file
  FRecLen = Len(AP1099)
  Open1099File FRecLen, Num1099Recs, Fed1099File

  '--Open the print file
  PRNFile = FreeFile
  ReportFile$ = "AP1099.PRN"
  Open ReportFile$ For Output As #PRNFile
  fmt$ = "$##,###,###.##"
  Header$ = "1099 Misc Report Totals"

  '--Print the Forms
  Page = 0
  
  For r = 1 To Num1099Recs
    FrmShowPctComp.ShowPctComp r, Num1099Recs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frm1099ProcessMenu
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

    Get Fed1099File, r, AP1099
   If AP1099.Deleted = 0 Then
      If AP1099.Void <> 1 Then
        TotAmtRep# = Round#(TotAmtRep# + AP1099.BOX1 + AP1099.BOX2)
        TotAmtRep# = Round#(TotAmtRep# + AP1099.BOX3 + AP1099.BOX5)
        TotAmtRep# = Round#(TotAmtRep# + AP1099.BOX6 + AP1099.BOX7)
        TotAmtRep# = Round#(TotAmtRep# + AP1099.BOX8 + AP1099.BOX10)
        TotAmtRep# = Round#(TotAmtRep# + AP1099.BOX14)
        TaxesWH# = Round#(TaxesWH# + AP1099.BOX4)
        GrandTotAmtRep# = GrandTotAmtRep# + TotAmtRep#
        VFlag = ""
      Else
        VFlag = "Voided"
      End If
      ToPrint$ = Space(80)
      ToPrint$ = AP1099.RecID + VFlag + "~"
      ToPrint$ = ToPrint$ + QPTrim(AP1099.RecName) + "~" + Using("$#,###,###.##", TotAmtRep#)
      ToPrint$ = ToPrint$ + "~" + QPTrim$(AP1099.RecADDR) + "~" + QPTrim$(AP1099.RecADDR2)
      ToPrint$ = ToPrint$ + "~" + QPTrim$(AP1099.RecCSZ) + "~" + QPTrim(AP1099.DBA)
      Print #PRNFile, ToPrint$
      
      totforms = totforms + 1
      TotAmtRep# = 0
    End If
  Next
  Close
  Load frmLoadingRpt
  ActivateControls frm1099ProcessMenu
  ARpt1099List.txtTown = QPTrim(GLUserName$)
  ARpt1099List.Label1.Caption = Header$
  ARpt1099List.txtDate.Caption = Now
  ARpt1099List.totforms = Using("###", totforms)
  ARpt1099List.totFed = Using(fmt$, TaxesWH#)
  ARpt1099List.totAmt = Using(fmt$, GrandTotAmtRep#)
  ARpt1099List.GetName ReportFile$
  ARpt1099List.startrpt

  Exit Sub

CancelExit:
  Exit Sub
End Sub
Private Sub Print1099Report2()
  Dim FRecLen As Integer, Num1099Recs As Integer, Fed1099File As Integer
  Dim PRNFile As Integer, ReportFile As String, fmt As String
  Dim Header As String, Page As Integer, r As Integer, lc As Integer
  Dim TotAmtRep As Double, TaxesWH As Double, GrandTotAmtRep As Double
  Dim totforms As Integer, VFlag As String
  FrmShowPctComp.Label1 = "Creating 1099 Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frm1099ProcessMenu

 '--Open the 1099 data file
  FRecLen = Len(AP1099)
  Open1099File FRecLen, Num1099Recs, Fed1099File

  '--Open the print file
  PRNFile = FreeFile
  ReportFile$ = "AP1099.PRN"
  Open ReportFile$ For Output As #PRNFile
  fmt$ = "$##,###,###.##"
  Header$ = "1099 Misc Report Totals"

  '--Print the Forms
  Page = 0
  GoSub Heading
  For r = 1 To Num1099Recs
    FrmShowPctComp.ShowPctComp r, Num1099Recs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frm1099ProcessMenu
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

    Get Fed1099File, r, AP1099
    If lc >= 60 Then
     Print #PRNFile, Chr$(12);
     GoSub Heading
    End If
   If AP1099.Deleted = 0 Then
      If AP1099.Void <> 1 Then
        TotAmtRep# = Round#(TotAmtRep# + AP1099.BOX1 + AP1099.BOX2)
        TotAmtRep# = Round#(TotAmtRep# + AP1099.BOX3 + AP1099.BOX5)
        TotAmtRep# = Round#(TotAmtRep# + AP1099.BOX6 + AP1099.BOX7)
        TotAmtRep# = Round#(TotAmtRep# + AP1099.BOX8 + AP1099.BOX10)
        TaxesWH# = Round#(TaxesWH# + AP1099.BOX4)
        GrandTotAmtRep# = GrandTotAmtRep# + TotAmtRep#
        VFlag = ""
      Else
        VFlag = "Voided"
      End If
      Print #PRNFile, AP1099.RecID; VFlag;
      Print #PRNFile, Tab(20); RTrim$(AP1099.RecName); Tab(65); Using("$#,###,###.##", TotAmtRep#)
      Print #PRNFile, Tab(20); QPTrim(AP1099.DBA)
      Print #PRNFile, RTrim$(AP1099.RecADDR) + " " + RTrim$(AP1099.RecADDR2);
      Print #PRNFile, Tab(50); RTrim$(AP1099.RecCSZ)
      Print #PRNFile, String$(79, "-")
      lc = lc + 4
      totforms = totforms + 1
      TotAmtRep# = 0
    End If
  Next

  Print #PRNFile, "1099 Misc Report Totals"
  Print #PRNFile, "---------------------------------"
  Print #PRNFile, "Total Forms:"; Tab(24); Using("###", totforms)
  Print #PRNFile, "Federal Tax W/H:"; Tab(20); Using(fmt$, TaxesWH#)
  Print #PRNFile, "Total Amt Reported:"; Tab(20); Using(fmt$, GrandTotAmtRep#)
  Print #PRNFile, Chr$(12)

  Close
  ActivateControls frm1099ProcessMenu
  ViewPrint ReportFile$, Header$
  KillFile ReportFile$
  

  Exit Sub

Heading:
  Page = Page + 1
  Print #PRNFile, Tab(30); "1099 Misc Report"
  Print #PRNFile,
  Print #PRNFile, "Report Date: "; Date$; Tab(65); "Page #"; Page
  Print #PRNFile, "Acct #"; Tab(20); "Acct Name"; Tab(65); "Amount"
  Print #PRNFile, "Address"
  Print #PRNFile, String$(79, "="): lc = 7
  Return
CancelExit:
  Exit Sub
End Sub

Public Sub Print1099Forms(Individual As Integer)
  Dim PRecLen As Integer, PayerFile As Integer, FRecLen As Integer
  Dim Num1099Recs As Integer, Fed1099File As Integer, PRNFile As Integer
  Dim ReportFile As String, fmt As String, r As Integer, Header As String
  Dim ToPrint As String, ToPrint2 As String, BPrnCnt As Integer, endit As Boolean
  '--Get the Payer info
  ReDim payer(1) As AP1099PayerRecType
  PRecLen = Len(payer(1))
  PayerFile = FreeFile
  Open "APPAYER.DAT" For Random As PayerFile Len = PRecLen
  Get PayerFile, 1, payer(1)
  Close PayerFile
  '--Open the 1099 data file
  FRecLen = Len(AP1099)
  Open1099File FRecLen, Num1099Recs, Fed1099File
  ToPrint$ = ""
  ToPrint2$ = " ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ "

  '--Open the print file
  PRNFile = FreeFile
  ReportFile$ = "AP1099.PRN"
  Open ReportFile$ For Output As #PRNFile
'[][][][]
'**** took comma out!!!!!
'[][]][]
  fmt$ = ("#######.##")
  If Individual <> 0 Then
  '--Print specific form
    r = Individual
    Get Fed1099File, r, AP1099
    If Not AP1099.Deleted Then
      GoSub PrintForm
    End If
  Else
  '--Print all the Forms
    For r = 1 To Num1099Recs
      Get Fed1099File, r, AP1099
   If AP1099.Deleted = 0 Then
        GoSub PrintForm
  '      didcnt = didcnt + 1
  '      IF didcnt > 0 THEN
  '        EXIT FOR
  '      END IF
      End If
    Next
  End If
  If Not grpt Then
    Print #PRNFile, Chr$(12)
  Else
    endit = True
    GoSub Dblcheck
  End If
  Close

  If Not grpt Then
    ViewPrint ReportFile$, Header$, , , True, "AP1099.MSK"
  'KILL ReportFile$
  Else
    Load frmLoadingRpt
    ActivateControls frm1099ProcessMenu
    ARpt1099Form.GetName ReportFile$
    ARpt1099Form.startrpt
  End If
Exit Sub


PrintForm:
If Not grpt Then
'line 1
  Print #PRNFile, "~"; Tab(70); "~"
'line 2
  Print #PRNFile, " "
'line 3
  If AP1099.Void = 1 Then
    Print #PRNFile, Tab(24); "X";
  End If
  If AP1099.Corrected = 1 Then
    Print #PRNFile, Tab(32); "X"
  Else
    Print #PRNFile, " "
  End If
'line 4
  Print #PRNFile, " "
'line 5
  Print #PRNFile, Tab(6); payer(1).Name; Tab(42); Using0(fmt$, AP1099.BOX1)
'line 6
  Print #PRNFile, Tab(6); payer(1).ADDR
'line 7
  Print #PRNFile, Tab(6); payer(1).Addr2
'line 8
  Print #PRNFile, Tab(6); payer(1).CSZ; Tab(42); Using0(fmt$, AP1099.BOX2)
'line 9
  Print #PRNFile, " "
'line 10
  Print #PRNFile, " " 'TAB(6); LEFT$(Payer(1).FedID, 16);
'line 11
  Print #PRNFile, Tab(42); Using0(fmt$, AP1099.BOX3); Tab(57); Using0(fmt$, AP1099.BOX4)
'line 12
  Print #PRNFile, " "
'line 13
  Print #PRNFile, " "
'line 14
  Print #PRNFile, " "
'line 15
  Print #PRNFile, Tab(6); Left$(payer(1).Fedid, 16); Tab(24); QPTrim$(AP1099.RecID); Tab(42); Using0(fmt$, AP1099.BOX5); Tab(57); Using0(fmt$, AP1099.BOX6)
'line 16
  Print #PRNFile, " "
'line 17
  Print #PRNFile, Tab(6); AP1099.RecName
'line 18
  If Len(QPTrim(AP1099.DBA)) <> 0 Then
    Print #PRNFile, Tab(6); "D/B/A: " + AP1099.DBA
  Else
    Print #PRNFile, " "
  End If
'line 19
  Print #PRNFile, Tab(42); Using0(fmt$, AP1099.BOX7); Tab(57); Using0(fmt$, AP1099.BOX8)
'line 20
  Print #PRNFile, " "
'line 21
  Print #PRNFile, Tab(6); AP1099.RecADDR
'line 22
  Print #PRNFile, Tab(6); AP1099.RecADDR2; Tab(52); QPTrim$(AP1099.BOX9); Tab(57); Using0(fmt$, AP1099.BOX10)
'line 23
  Print #PRNFile, " "
'line 24
  Print #PRNFile, Tab(6); AP1099.RecCSZ
'line 25
  Print #PRNFile, " "
'line 26
  Print #PRNFile, " "
'line 27
  Print #PRNFile, Tab(6); AP1099.RecACCT; Tab(36); AP1099.NOTICE;
  Print #PRNFile, Tab(42); Using0(fmt$, AP1099.BOX13); Tab(57); Using0(fmt$, AP1099.BOX14)
'line 28
  Print #PRNFile, " "
'line 29
  Print #PRNFile, Tab(6); AP1099.BOX15;
  Print #PRNFile, Tab(42); Using0(fmt$, AP1099.BOX16); Tab(55); QPTrim(AP1099.BOX17); Tab(69); Using0(fmt$, AP1099.BOX18)
'line 30
  Print #PRNFile, " "
'line 31
  Print #PRNFile, " "
'line 32
  Print #PRNFile, " "
'line 33
  Print #PRNFile, "~"; Tab(70); "~"
  
  Else
  
  ToPrint$ = ToPrint$ + QPTrim(payer(1).Name) + "~" + QPTrim(payer(1).ADDR) + "~" + QPTrim(payer(1).Addr2) + "~"
  ToPrint$ = ToPrint$ + QPTrim(payer(1).CSZ) + "~" + QPTrim(payer(1).Fedid) + "~" + QPTrim$(AP1099.RecID) + "~"
  ToPrint$ = ToPrint$ + QPTrim(AP1099.RecName) + "~"
  If Len(QPTrim(AP1099.DBA)) <> 0 Then
   ToPrint$ = ToPrint$ + "D/B/A: " + QPTrim(AP1099.DBA) + "~"
  Else
   ToPrint$ = ToPrint$ + QPTrim(AP1099.DBA) + "~"
  End If
  ToPrint$ = ToPrint$ + QPTrim(AP1099.RecADDR) + "~" + QPTrim(AP1099.RecADDR2) + "~"
  ToPrint$ = ToPrint$ + QPTrim(AP1099.RecCSZ) + "~" + QPTrim(AP1099.RecACCT) + "~" + QPTrim(AP1099.NOTICE) + "~"
  ToPrint$ = ToPrint$ + Using0(fmt$, AP1099.BOX1) + "~"
  ToPrint$ = ToPrint$ + Using0(fmt$, AP1099.BOX2) + "~"
  ToPrint$ = ToPrint$ + Using0(fmt$, AP1099.BOX3) + "~"
  ToPrint$ = ToPrint$ + Using0(fmt$, AP1099.BOX4) + "~"
  ToPrint$ = ToPrint$ + Using0(fmt$, AP1099.BOX5) + "~"
  ToPrint$ = ToPrint$ + Using0(fmt$, AP1099.BOX6) + "~"
  ToPrint$ = ToPrint$ + Using0(fmt$, AP1099.BOX7) + "~"
  ToPrint$ = ToPrint$ + Using0(fmt$, AP1099.BOX8) + "~"
  ToPrint$ = ToPrint$ + QPTrim$(AP1099.BOX9) + "~"
  ToPrint$ = ToPrint$ + Using0(fmt$, AP1099.BOX10) + "~"
  ToPrint$ = ToPrint$ + Using0(fmt$, AP1099.BOX13) + "~"
  ToPrint$ = ToPrint$ + Using0(fmt$, AP1099.BOX14) + "~"
  ToPrint$ = ToPrint$ + QPTrim$(AP1099.BOX15) + "~"
  ToPrint$ = ToPrint$ + Using0(fmt$, AP1099.BOX16) + "~"
  ToPrint$ = ToPrint$ + QPTrim$(AP1099.BOX17) + "~"
  ToPrint$ = ToPrint$ + Using0(fmt$, AP1099.BOX18) + "~"
  If AP1099.Void = 1 Then
    ToPrint$ = ToPrint$ + "X" + "~"
  Else
    ToPrint$ = ToPrint$ + " " + "~"
  End If
  If AP1099.Corrected = 1 Then
    ToPrint$ = ToPrint$ + "X" + "~"
  Else
    ToPrint$ = ToPrint$ + " " + "~"
  End If
'  Print #PRNFile, ToPrint$
'  ToPrint$ = ""
  BPrnCnt = BPrnCnt + 1
  End If
  
'  Return
Dblcheck:
    If BPrnCnt = 2 Then
      Print #PRNFile, ToPrint$
      ToPrint$ = ""
      BPrnCnt = 0
    ElseIf BPrnCnt = 1 And endit = True Then
      ToPrint$ = ToPrint$ + ToPrint2$
      Print #PRNFile, ToPrint$
      ToPrint$ = ""
      BPrnCnt = 0
'    ElseIf BPrnCnt = 2 And endit = True Then
'      ToPrint$ = ToPrint$ + ToPrint2$
'      Print #PRNFile, ToPrint$
'      ToPrint$ = ""
'      BPrnCnt = 0
    End If

Return


End Sub



