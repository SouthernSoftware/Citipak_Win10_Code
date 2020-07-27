VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUBExportMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export Utility Billing Information Menu"
   ClientHeight    =   8868
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   12216
   Icon            =   "frmUBExportMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdPostalImport 
      BackColor       =   &H008F8265&
      Caption         =   "Import P&ostal Address Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   516
      Left            =   3840
      MaskColor       =   &H8000000F&
      TabIndex        =   7
      Top             =   5628
      Width           =   4524
   End
   Begin VB.CommandButton cmdExpPostal 
      BackColor       =   &H008F8265&
      Caption         =   "Export &Postal Address Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   516
      Left            =   3840
      MaskColor       =   &H8000000F&
      TabIndex        =   6
      Top             =   4896
      Width           =   4524
   End
   Begin VB.CommandButton cmdExportSReads 
      BackColor       =   &H008F8265&
      Caption         =   "Export &S Reading Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   516
      Left            =   3840
      MaskColor       =   &H8000000F&
      TabIndex        =   2
      Top             =   4152
      Width           =   4524
   End
   Begin VB.CommandButton cmdExitUBExportMenu 
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
      Height          =   516
      Left            =   3840
      TabIndex        =   3
      Top             =   6360
      Width           =   4524
   End
   Begin VB.CommandButton cmdExportConsumption 
      Caption         =   "Export &Consumption Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   516
      Left            =   3840
      TabIndex        =   1
      Top             =   3420
      Width           =   4524
   End
   Begin VB.CommandButton cmdExportCustomer 
      BackColor       =   &H008F8265&
      Caption         =   "&Export Customer Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   516
      Left            =   3840
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   2688
      Width           =   4524
   End
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
            TextSave        =   "4:05 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "5/12/2005"
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
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   8976
      X2              =   8976
      Y1              =   2064
      Y2              =   7944
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   8976
      X2              =   9696
      Y1              =   7944
      Y2              =   7944
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   3
      X1              =   8856
      X2              =   9816
      Y1              =   2064
      Y2              =   2064
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   8856
      X2              =   9816
      Y1              =   1944
      Y2              =   1944
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   8856
      X2              =   8856
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   9816
      X2              =   9816
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2520
      X2              =   2520
      Y1              =   2064
      Y2              =   7944
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2520
      X2              =   3240
      Y1              =   7944
      Y2              =   7944
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   2400
      X2              =   3360
      Y1              =   2064
      Y2              =   2064
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2400
      X2              =   3360
      Y1              =   1944
      Y2              =   1944
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2400
      X2              =   2400
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   3360
      X2              =   3360
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1776
      Top             =   744
      Width           =   8652
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UTILITY BILLING EXPORT MENU"
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
      Left            =   3012
      TabIndex        =   4
      Top             =   1104
      Width           =   6156
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
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2520
      Top             =   2064
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2400
      Top             =   1824
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8976
      Top             =   2064
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   8856
      Top             =   1824
      Width           =   972
   End
End
Attribute VB_Name = "frmUBExportMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Public IExFN As String
Private Sub cmdExportConsumption_Click()
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

  frmExpCustConsump.Show
  Unload frmUBExportMenu
End Sub

Private Sub cmdExportCustomer_Click()
  Dim FntSize As Integer
  If Not Exist(UBPath$ + "UBCust.dat") Then
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
  'For Johnston
  ''frmExpCustomerInfo2.Show
  frmExpCustomerInfo.Show
  Unload frmUBExportMenu
End Sub

Private Sub cmdExportSReads_Click()
  Dim FntSize As Integer
  If Not Exist(UBPath$ + "UBCust.dat") Then
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

  DeActivateControls Me
'do consumpstuff here
    BuckSportReadingExport
  ActivateControls Me
End Sub

Private Sub cmdExpPostal_Click()
  Dim FntSize As Integer
  If Not Exist(UBPath$ + "UBCust.dat") Then
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
  DeActivateControls Me
  ExpPostalCass
  ActivateControls Me
End Sub

Private Sub CmdPostalImport_Click()
  Dim FntSize As Integer, msgtogo As String
  If Not Exist(UBPath$ + "UBCust.dat") Then
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
  IExFN = ""
  msgtogo = "The Name of the Import File."
  'this hardcodes the file name
  'on the frmimpexpmsg the field is set to readonly so cant change
  frmImpExpMsg.txtFileName.Text = "UBPostIn.CSV"
  frmImpExpMsg.txtFileName.Visible = True
  frmImpExpMsg.Label1 = msgtogo
  frmImpExpMsg.Show 1, Me
  If frmImpExpMsg.Exout <> 1 Then
    DeActivateControls Me
    IExFN = UBPath$ + IExFN
    ImpPostalCass
    ActivateControls Me
  Else
    Exit Sub
  End If
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitUBExportMenu.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via UBExportMenu by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      KeyCode = 0
      Call cmdExitUBExportMenu_Click
    Case Else:
  End Select
End Sub

Private Sub cmdExitUBExportMenu_Click()
  Load frmUBSetupMenu
  DoEvents
  frmUBSetupMenu.Show
  DoEvents
  Unload Me
End Sub
Private Sub BuckSportReadingExport()
  Dim Dash80 As String, IndexName As String, IdxRecLen As Integer
  Dim UBCustRecLen As Integer, UBCust As Integer, UBTran As Integer
  Dim IdxFileSize As Long, IdxNumOfRecs As Long, cnt As Long
  Dim SRecLen As Integer, NumOfRecs As Long, NumOfCust As Long
  Dim Handle As Integer, UsingBook As Boolean, NumOfPeriods As Integer
  Dim UBRpt As String, UCode As String
  FrmShowPctComp.Label1 = "'S' Reading Export."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me

  IndexName$ = BookIndexFile
  UsingBook = True
  CrLf$ = Chr$(13) + Chr$(10)

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  ReDim SExport(1) As SReadType
  SRecLen = Len(SExport(1))

  IdxRecLen = 4               'we are using a long integer
  IdxFileSize& = FileSize(IndexName$)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
'  FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
  NumOfRecs = IdxNumOfRecs

  Handle = FreeFile
  Open IndexName$ For Random Shared As Handle Len = IdxRecLen
  For cnt = 1 To IdxNumOfRecs
    Get #Handle, cnt, IdxBuff(cnt)
  Next
  Close Handle
  KillFile UBPath$ + "UBSREAD.TXT"
  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen

  UBRpt = FreeFile
  Open UBPath$ + "UBSREAD.TXT" For Random As UBRpt Len = SRecLen
  For cnt = 1 To NumOfRecs
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs&
    If FrmShowPctComp.Out Then
      Close
      Unload FrmShowPctComp
      Exit Sub
    End If

    Get UBCust, IdxBuff(cnt).RecNum, UBCustRec(1)
    If UBCustRec(1).DelFlag = 0 And UBCustRec(1).Status = "A" Then
      UCode$ = UCase$(QPTrim$(UBCustRec(1).USERCODE1))
      If UCode$ = "S" Then
        ReDim SExport(1) As SReadType
        SExport(1).CrLf$ = CrLf$
        SExport(1).Book = QPTrim(UBCustRec(1).Book)
        SExport(1).Seq = QPTrim(UBCustRec(1).SEQNUMB)
        LSet SExport(1).CustName = QPTrim$(UBCustRec(1).CustName)
        LSet SExport(1).ServAddr = Left$(QPTrim$(UBCustRec(1).ServAddr), 19)
        LSet SExport(1).CurrRead = QPTrim$(Str$(UBCustRec(1).LocMeters(1).CurRead))
        SExport(1).ReadDate = Num2Date$(Trim(UBCustRec(1).LocMeters(1).CurDate))
        Put #UBRpt, , SExport(1)
      End If
    End If

  Next

  Close UBCust, UBRpt

  Erase IdxBuff, UBCustRec

   MsgBox "File Created is UBSREAD.TXT", vbOKOnly, "File created"
End Sub

Private Sub ExpPostalCass()
  Dim Dash80 As String, IndexName As String, IdxRecLen As Integer
  Dim UBCustRecLen As Integer, UBCust As Integer, UBTran As Integer
  Dim IdxFileSize As Long, IdxNumOfRecs As Long, cnt As Long
  Dim UBTranRecLen As Integer, NumOfRecs As Long, NumOfCust As Long
  Dim Handle As Integer, UsingBook As Boolean, NumOfPeriods As Integer
  Dim RecNo As Long, DidCnt As Long, ThisTrans As Long, FMonth As Integer
  Dim FYear As Integer, TYear As Integer, TMonth As Integer
  Dim FMCnt As Integer, DidAMeter As Boolean, MtrCnt As Integer
  Dim MeterType As String, MeterConsp As Long, MaxMeterAmt As Long
  Dim TotalConsump As Long, QPos As Integer, LocationNumber As String
  Dim Zip As String, CCCnt As Long, NumofRevs As Integer, BuckFmt As String
  Dim Bookone As Integer, Bookto As Integer, qc As String, ThisBook As String
  Dim q As String, C As String, qcq As String, OKFlag As Boolean
  Dim UBRpt As String, zz As String, zzN As Integer, CCnt As Long
  Dim UBOwnerRecLen As Integer, UBFile As Integer, AcctNumber As Long
  Dim WhatBook As Integer, Export As Long, RCnt As Integer, FCnt As Integer
  Dim MCnt As Integer, Address1 As String, Address2 As String, Cty As String
  FrmShowPctComp.Label1 = "Creating Export File"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  q$ = "," 'Chr$(34)
  'qc$ = q$ + ","
  'qcq$ = q$ + "," + q$
  'Bookone = Val(QPTrim(fptxtRoute1))
  'Bookto = Val(QPTrim(fptxtRoute2))
  'IndexName$ = BookIndexFile
  'UsingBook = True
  OKFlag = True
  'BuckFmt$ = "######.##"
  'NumofRevs = GetNumOfRevs%
  
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

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

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  NumOfRecs = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen

  
  KillFile (UBPath$ + "UBPostal.txt")
  UBRpt = FreeFile
  Open UBPath$ + "UBPostal.txt" For Output As UBRpt

  For cnt = 1 To NumOfRecs
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out Then
      Close
      Unload FrmShowPctComp
      GoTo ExitHere
    End If

    
    Get UBCust, cnt, UBCustRec(1)
    'Get UBFile, cnt, UBOwnerRec(1)

    '*************************************
    '   Main body of Printing goes here
    If UBCustRec(1).DelFlag <> -1 And QPTrim$(UBCustRec(1).Status) = "A" Then
'      ThisBook$ = UBCustRec(1).Book
'      If Left$(ThisBook$, 1) = "0" Then
'        WhatBook = Val(Right$(ThisBook$, 1))
'      Else
'        WhatBook = Val(ThisBook$)
'      End If
'      If WhatBook <= Bookto And WhatBook >= Bookone Then
        Export& = Export& + 1
        GoSub ExportThisAccount
'     End If
    End If
  Next

  Close
  FrmShowPctComp.ShowPctComp 1, 1
  If Export& > 0 Then
    MsgBox "File " & UBPath$ & "UBPostal.txt Exported with " & Export& & " Accounts.", vbOKOnly, "Export Completed."
  Else
    MsgBox "No Information Found to Export.", vbOKOnly, "Procedure Ended"
  End If
GoTo ExitHere

ExportThisAccount:

  LocationNumber$ = QPTrim$(UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB)
  Zip$ = QPTrim$(UBCustRec(1).ZIPCODE)
'  If Len(Zip$) > 5 Then
'    Zip$ = Left$(Zip$, 5) + "-" + Mid$(Zip$, 6)
'  End If
  Address1$ = QPStripCom$(UBCustRec(1).ADDR1)
  Address2$ = QPStripCom$(UBCustRec(1).ADDR2)
  Cty$ = QPStripCom$(UBCustRec(1).CITY)
  Print #UBRpt, QPTrim$(Str$(cnt)); 'q$;
  'Print #UBRpt, qcq$; LocationNumber$;
  'Print #UBRpt, qcq$; UBCustRec(1).Status;
  'Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).CustName);
  Print #UBRpt, q$; Address1$;           'QPTrim$(UBCustRec(1).ADDR1);
  Print #UBRpt, q$; Address2$;          'QPTrim$(UBCustRec(1).ADDR2);
  'Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).ServAddr);
  Print #UBRpt, q$; Cty$;
  Print #UBRpt, q$; QPTrim$(UBCustRec(1).STATE);
  Print #UBRpt, q$; Zip$
'  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).POSTRTE);
'
'This is if want to export owner also????????????????
'  Print #UBRpt, q$; QPTrim$(UBOwnerRec(1).ADDR1);
'  Print #UBRpt, q$; QPTrim$(UBOwnerRec(1).ADDR2);
'  Print #UBRpt, q$; QPTrim$(UBOwnerRec(1).CITY);
'  Print #UBRpt, q$; QPTrim$(UBOwnerRec(1).STATE);
'  Print #UBRpt, q$; QPTrim$(UBOwnerRec(1).ZIPCODE)

Return
ExitHere:

End Sub

Private Sub ImpPostalCass()
  Dim Dash80 As String, IndexName As String, IdxRecLen As Integer
  Dim UBCustRecLen As Integer, UBCust As Integer, UBTran As Integer
  Dim IdxFileSize As Long, IdxNumOfRecs As Long, cnt As Long
  Dim UBTranRecLen As Integer, NumOfRecs As Long, NumOfCust As Long
  Dim Handle As Integer, UsingBook As Boolean, NumOfPeriods As Integer
  Dim RecNo As Long, DidCnt As Long, ThisTrans As Long, FMonth As Integer
  Dim FYear As Integer, TYear As Integer, TMonth As Integer
  Dim FMCnt As Integer, DidAMeter As Boolean, MtrCnt As Integer
  Dim MeterType As String, MeterConsp As Long, MaxMeterAmt As Long
  Dim TotalConsump As Long, QPos As Integer, LocationNumber As String
  Dim Zip As String, CCCnt As Long, NumofRevs As Integer, BuckFmt As String
  Dim Bookone As Integer, Bookto As Integer, qc As String, ThisBook As String
  Dim q As String, C As String, qcq As String, OKFlag As Boolean
  Dim UBRpt As String, zz As String, zzN As Integer, CCnt As Long
  Dim UBOwnerRecLen As Integer, UBFile As Integer, AcctNumber As Long
  Dim WhatBook As Integer, Import As Long, RCnt As Integer, FCnt As Integer
  Dim MCnt As Integer, Address1 As String, Address2 As String, Dp As String
  Dim Acct As String, Add1 As String, Add2 As String, ST As String, Cty As String, Zp As String
  Dim X1 As String, RR As String, Lot As String, X2 As String
  FrmShowPctComp.Label1 = "Creating Export File"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  On Local Error GoTo impend
  If Len(IExFN) > 0 Then
    If Not Exist(IExFN) Then
      FrmShowPctComp.ShowPctComp 1, 1
      MsgBox "Import File Does Not Exist", vbOKOnly, "Procedure Cancelled"
      Exit Sub
    End If
  Else
    FrmShowPctComp.ShowPctComp 1, 1
    MsgBox "Invalid File Name", vbOKOnly, "Procedure Cancelled"
    Exit Sub
  End If
  OKFlag = True
  'BuckFmt$ = "######.##"
  'NumofRevs = GetNumOfRevs%
  
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

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
  NumOfRecs& = GetNumOfCust&
  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  
  'KillFile (UBPath$ + "UBPostal.txt")
  UBRpt = FreeFile
  Open IExFN For Input As UBRpt
  'Input header record
  Input #UBRpt, Acct$, Add1$, Add2$, Cty$, ST$, Zp$, Dp$, X1$, RR$, Lot$, X2$
  Do While Not eof(UBRpt)
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
    Input #UBRpt, Acct$, Add1$, Add2$, Cty$, ST$, Zp$, Dp$, X1$, RR$, Lot$, X2$
    cnt = Val(Acct)
  'If cnt is greater than the total num of customers then bad data
    If cnt <= NumOfRecs& Then
      Get UBCust, cnt, UBCustRec(1)
      If UBCustRec(1).DelFlag <> -1 And QPTrim$(UBCustRec(1).Status) = "A" Then
          Import& = Import& + 1
          GoSub ImportThisAccount
      End If
    Else
      FrmShowPctComp.ShowPctComp 1, 1
      MsgBox "No Information Found to Import.", vbOKOnly, "Procedure Ended"
      GoTo impend
    End If

  Loop
  
  Close
  FrmShowPctComp.ShowPctComp 1, 1
  If Import& > 0 Then
    MsgBox "File " & IExFN & " Imported with " & Import& & " Accounts.", vbOKOnly, "Import Completed."
  Else
    MsgBox "No Information Found to Import.", vbOKOnly, "Procedure Ended"
  End If
GoTo ExitHere

ImportThisAccount:

  Zip$ = QPTrim$(Zp)
  Address1$ = QPTrim$(Add1)
  Address2$ = QPTrim$(Add2)
  UBCustRec(1).ADDR1 = Address1$
  UBCustRec(1).ADDR2 = Address2$
  UBCustRec(1).CITY = QPTrim$(Cty)
  UBCustRec(1).STATE = QPTrim$(ST)
  UBCustRec(1).ZIPCODE = QPTrim$(Zip$)
  UBCustRec(1).DPCode = QPTrim$(Dp$)
  UBCustRec(1).POSTRTE = QPTrim$(RR$)
  Put UBCust, cnt, UBCustRec(1)
Return
impend:
Close
ExitHere:

End Sub

