VERSION 5.00
Begin VB.Form frmVATaxStandardBill 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Standard Tax Bill Format"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "frmVATaxStandardBill.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdRealTest 
      Caption         =   "F6 &Real Test Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4080
      TabIndex        =   3
      Top             =   4920
      Width           =   3336
   End
   Begin VB.CommandButton cmdPersTest 
      Caption         =   "F5 &Personal Test Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4080
      TabIndex        =   2
      Top             =   2640
      Width           =   3336
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Esc E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4680
      TabIndex        =   1
      Top             =   7080
      Width           =   2292
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1692
      Left            =   3360
      Top             =   4320
      Width           =   4812
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1692
      Left            =   3360
      Top             =   2040
      Width           =   4812
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   4692
      Left            =   1080
      Top             =   1680
      Width           =   9492
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   516
      Left            =   2940
      Top             =   360
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Bill Standard"
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
      Left            =   3372
      TabIndex        =   0
      Top             =   456
      Width           =   5016
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   636
      Left            =   2940
      Top             =   240
      Width           =   5772
   End
End
Attribute VB_Name = "frmVATaxStandardBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  'Private Temp_Class As Resize_Class
  Dim Over As clsTextBoxOverRider
Private Sub cmdExit_Click()
  Unload Me
  DoEvents
End Sub

Private Sub cmdPersTest_Click()
  Call PrintPersVAStandard
End Sub

Private Sub cmdRealTest_Click()
  Call PrintRealVAStandard
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxStandardBill.")
      Call Terminate
      End
    End If
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
      KeyCode = 0
      DoEvents
      Call cmdExit_Click
    Case vbKeyF5:
      KeyCode = 0
      DoEvents
      Call cmdPersTest_Click
    Case vbKeyF6:
      KeyCode = 0
      DoEvents
      Call cmdRealTest_Click
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
End Sub

Private Sub Form_Resize()
  If Me.Visible Then
    'Temp_Class.ResizeControls Me
    DoEvents
  End If
End Sub
Private Sub PrintPersVAStandard()
  'checked OK against mask (taxppmsk.dat) on 10/21/2005
  Dim x As Long, PYearStr$
  Dim File$, LC As Integer, CustName$
  Dim WhatYear As Integer, WhatPers&
  Dim PPTRAVal#, RptHandle As Integer
  Dim NumOfPRecs As Long, PPTRADiscount#
  Dim VehDesc$, PERC!
  Dim TaxAmt#, LCnt As Integer
  Dim PYear As Integer
  Dim MinVehVal As Double
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim MultiYear As Integer
  Dim TownName$, City$, State$, Zip$
  Dim Add1$, Add2$, Add3$
  Dim ThisForm As Form
  Dim MaxVehTaxVal#, MinVehTaxVal#
  Dim GPPTRADisc#
  Dim RptFile$
  
  RptHandle = FreeFile
  RptFile$ = "STANDPTST.PRN"
  Call TaxMsg(900, "Pitch 10 is recommended for this bill.")
  
  Open RptFile For Output As RptHandle
  Set ThisForm = New frmVATaxSystemSetup
  If Exist("TAXSETUP.DAT") Then
    OpenTaxSetUpFile TMHandle
    Get TMHandle, 1, TaxMasterRec
    Close TMHandle
    MinVehTaxVal = TaxMasterRec.MinVehTaxVal
    MaxVehTaxVal = TaxMasterRec.MaxVehTaxVal
    GPPTRADisc = TaxMasterRec.PPTRADisc
    MultiYear = TaxMasterRec.MultiYear
    TownName = QPTrim$(TaxMasterRec.Name)
    Add1 = QPTrim$(TaxMasterRec.Add1)
    Add2 = QPTrim$(TaxMasterRec.Add2)
    Add3 = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip)
  Else
    If QPTrim$(ThisForm.fpcmbMultiYear.Text) <> "" Then
      MultiYear = CInt(ThisForm.fpcmbMultiYear.Text)
    Else
      MultiYear = 1
    End If
    If CDbl(ThisForm.fpCurrMaxVehAmt.Value) > 0 Then
      MaxVehTaxVal = CDbl(ThisForm.fpCurrMaxVehAmt.Value)
    Else
      MaxVehTaxVal = 20000
    End If
    If CDbl(ThisForm.fpCurrMinVehAmt.Value) > 0 Then
      MinVehTaxVal = CDbl(ThisForm.fpCurrMinVehAmt.Value)
    Else
      MinVehTaxVal = 20000
    End If
    If CDbl(ThisForm.fpDSPPTRADisc.Value) > 0 Then
      GPPTRADisc = CDbl(ThisForm.fpDSPPTRADisc.Value)
    Else
      GPPTRADisc = 70
    End If
    If QPTrim$(ThisForm.fptxtNameOfTaxAuth.Text) <> "" Then
      TownName = QPTrim$(ThisForm.fptxtNameOfTaxAuth.Text)
    Else
      TownName = "Town Of Your Town"
    End If
    If QPTrim$(ThisForm.fptxtAdd1.Text) <> "" Then
      Add1 = QPTrim$(ThisForm.fptxtAdd1.Text)
    Else
      Add1 = "120 Main St"
    End If
    If QPTrim$(ThisForm.fptxtAdd2.Text) <> "" Then
      Add2 = QPTrim$(ThisForm.fptxtAdd2.Text)
    Else
      Add2 = "PO Box 1190"
    End If
    If QPTrim$(ThisForm.fptxtCity.Text) <> "" Then
      City$ = QPTrim$(ThisForm.fptxtCity.Text)
    Else
      City = "Your Town"
    End If
    If QPTrim$(ThisForm.fptxtState.Text) <> "" Then
      State$ = QPTrim$(ThisForm.fptxtState.Text)
    Else
      State = "XX"
    End If
    If QPTrim$(ReplaceString(ThisForm.fptxtZip.Text, "-", "")) <> "" Then
      Zip$ = QPTrim$(ThisForm.fptxtZip.Text)
    Else
      Zip$ = "55555-5555"
    End If
    Add3 = City$ + ", " + State + " " + Zip
  End If
  
  WhatYear = CInt(Mid(Date, 7, 4))
  
  If WhatYear = 1999 Then PERC! = 27.5
  If WhatYear = 2000 Then PERC! = 47.5
  
  If WhatYear >= 2001 Then PERC! = GPPTRADisc
  CustName$ = "John Q Public"
  Print #RptHandle, "~"
  Print #RptHandle, Tab(63); "TAX YEAR: "; WhatYear
  Print #RptHandle, Tab(75); Using$("#####", 1)
  Print #RptHandle, " "
  Print #RptHandle, " "
  If InStr(TaxMasterRec.Name, "HALIFAX") Then
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, " "
  Else
    Print #RptHandle, " "
    Print #RptHandle, Tab(5); TownName$
    Print #RptHandle, Tab(5); Add1$
    Print #RptHandle, Tab(5); Add2$
    Print #RptHandle, Tab(5); Add3$
  End If
  Print #RptHandle, " "
  Print #RptHandle, " "
  Print #RptHandle, " "
  Print #RptHandle, Tab(5); "Acct # "; Using$("#####0", 100)
  Print #RptHandle, Tab(5); CustName$
  Print #RptHandle, Tab(5); "700 Elm Street"
  Print #RptHandle, Tab(5); "PO Box 567"
  Print #RptHandle, Tab(5); "Your Town, XX 55555-5555"
'  Print #RptHandle, " " 'added line
  For LC = 18 To 21 'made 18 = 19
   Print #RptHandle, " "
  Next LC
  Print #RptHandle, Tab(37); "PROP VALUE"; Tab(50); "TAX AMOUNT"; Tab(62); "PPTRA DIS"; Tab(72); "TOTAL DUE"
  Print #RptHandle, " "
 'Line 24 Starts Here
  Print #RptHandle, "Personal Property"; Tab(32); Using$("#.00", 0.25);
   Print #RptHandle, Tab(37); Using$("#####0.00", 12500);
   Print #RptHandle, Tab(51); Using$("#####0.00", 31.25);
   Print #RptHandle, Tab(63); Using$("####0.00", 21.88);
   Print #RptHandle, Tab(72); Using$("#####0.00", 9.37)
  Print #RptHandle, "Machinery/Tools"; Tab(32); Using$("#.00", 0.25);
   Print #RptHandle, Tab(37); Using$("#####0.00", 0);
   Print #RptHandle, Tab(51); Using$("#####0.00", 0);
   Print #RptHandle, Tab(72); Using$("#####0.00", 0)
  Print #RptHandle, "Farm Equipment";
   Print #RptHandle, Tab(32); Using("#.00", 0.25);
   Print #RptHandle, Tab(37); Using$("######.##", 0);
   Print #RptHandle, Tab(51); Using$("######.##", 0);
   Print #RptHandle, Tab(72); Using$("######.##", 0)
  Print #RptHandle, "Mobile Homes";
   Print #RptHandle, Tab(32); Using$("#.00", 0.25);
   Print #RptHandle, Tab(37); Using$("#####0.00", 0);
   Print #RptHandle, Tab(51); Using$("#####0.00", 0);
   Print #RptHandle, Tab(72); Using$("#####0.00", 0)
  Print #RptHandle, "Merchant Capital";
   Print #RptHandle, Tab(32); Using$("#.00", 0.25);
   Print #RptHandle, Tab(37); Using$("#####0.00", 0);
   Print #RptHandle, Tab(51); Using$("#####0.00", 0);
   Print #RptHandle, Tab(72); Using$("#####0.00", 0)
  Print #RptHandle, " PPTRA Vehicle Information"
 'Line 30 to 35 here to print vehicles
  Print #RptHandle, "*" + "VIN # 7878787FRT87877";
  Print #RptHandle, Tab(37); Using$("#####0.00", 12500);
  Print #RptHandle, Tab(51); Using$("#####0.00", 31.25);
  Print #RptHandle, Tab(63); Using$("#####.00", 21.88)
  For LCnt = 1 To 6: Print #RptHandle, "": Next LCnt
  
'  Print #RptHandle, Tab(48); "Total Tax Due ";
'  Print #RptHandle, Using$("$#######0.00", 9.37)
'  Print #RptHandle, Tab(48); "Tax Due Date: " + CStr(Date)
  If InStr(TaxMasterRec.Name, "HALIFAX") = 0 Then
'    Print #RptHandle,
'    Print #RptHandle,
'    Print #RptHandle,
'    Print #RptHandle,
'    Print #RptHandle,
  End If
  Print #RptHandle, Tab(48); "Total Tax Due ";
  Print #RptHandle, Using$("$#######0.00", 9.37)
  Print #RptHandle, Tab(48); "Tax Due Date: " + CStr(Date)
  If InStr(TaxMasterRec.Name, "HALIFAX") Then
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle,
  Else
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle,
  End If
  Print #RptHandle, "BN"; Using$("####0", 1)
  Print #RptHandle, "~"
  
  Close
  ViewPrint RptFile$, "Standard Bill Test Print", True
End Sub

Private Sub PrintRealVAStandard()
 'checked OK against mask (TAXREMSK.DAT) on 10/21/2005
 'STANDARD REAL ESTATE BILL FORMAT AS SOLD BY SOUTHERN SOFTWARE
 'TAXRESTD.BI
  Dim LC As Long, RealTaxRate#
  Dim CustName As String * 45, WhatYear As Integer
  Dim TaxAmt#, LCnt As Integer
  Dim MultiYear As Integer
  Dim TownName$, City$, State$, Zip$
  Dim Add1$, Add2$, Add3$
  Dim ThisForm As Form
  Dim RptFile$, RptHandle As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  
  Call TaxMsg(900, "Pitch 10 is recommended for this bill.")
  RptHandle = FreeFile
  RptFile$ = "STANDRTST.PRN"
  
  Open RptFile For Output As RptHandle
  Set ThisForm = New frmVATaxSystemSetup
  If Exist("TAXSETUP.DAT") Then
    OpenTaxSetUpFile TMHandle
    Get TMHandle, 1, TaxMasterRec
    Close TMHandle
    MultiYear = TaxMasterRec.MultiYear
    TownName = QPTrim$(TaxMasterRec.Name)
    Add1 = QPTrim$(TaxMasterRec.Add1)
    Add2 = QPTrim$(TaxMasterRec.Add2)
    Add3 = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip)
  Else
    If QPTrim$(ThisForm.fpcmbMultiYear.Text) <> "" Then
      MultiYear = CInt(ThisForm.fpcmbMultiYear.Text)
    Else
      MultiYear = 1
    End If
    If QPTrim$(ThisForm.fptxtNameOfTaxAuth.Text) <> "" Then
      TownName = QPTrim$(ThisForm.fptxtNameOfTaxAuth.Text)
    Else
      TownName = "Town Of Your Town"
    End If
    If QPTrim$(ThisForm.fptxtAdd1.Text) <> "" Then
      Add1 = QPTrim$(ThisForm.fptxtAdd1.Text)
    Else
      Add1 = "120 Main St"
    End If
    If QPTrim$(ThisForm.fptxtAdd2.Text) <> "" Then
      Add2 = QPTrim$(ThisForm.fptxtAdd2.Text)
    Else
      Add2 = "PO Box 1190"
    End If
    If QPTrim$(ThisForm.fptxtCity.Text) <> "" Then
      City$ = QPTrim$(ThisForm.fptxtCity.Text)
    Else
      City = "Your Town"
    End If
    If QPTrim$(ThisForm.fptxtState.Text) <> "" Then
      State$ = QPTrim$(ThisForm.fptxtState.Text)
    Else
      State = "XX"
    End If
    If QPTrim$(ReplaceString(ThisForm.fptxtZip.Text, "-", "")) <> "" Then
      Zip$ = QPTrim$(ThisForm.fptxtZip.Text)
    Else
      Zip$ = "55555-5555"
    End If
    Add3 = City$ + ", " + State + " " + Zip
  End If
  
  WhatYear = CInt(Mid(Date, 7, 4))
  
  RealTaxRate# = 0.25
  WhatYear = CInt(Mid(Date, 7, 4))

  CustName$ = "John Q Public"
  Print #RptHandle, "~"
  Print #RptHandle, Tab(64); "TAX YEAR: "; WhatYear
  Print #RptHandle, Tab(75); Using$("#####", 100)
  Print #RptHandle, " "
  Print #RptHandle, " "
  Print #RptHandle, " " 'added
  If InStr(TaxMasterRec.Name, "HALIFAX") Then
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, " "
  Else
    Print #RptHandle, Tab(5); TownName$
    Print #RptHandle, Tab(5); Add1$
    Print #RptHandle, Tab(5); Add2$
    Print #RptHandle, Tab(5); Add3$
  End If
  Print #RptHandle, " "
  Print #RptHandle, " "
  Print #RptHandle, " " 'added
  Print #RptHandle, " " 'added
  Print #RptHandle, Tab(5); "PIN:  " + "1234567"
  Print #RptHandle, Tab(5); "ACCT: " + Using$("#####", 150)
  Print #RptHandle, Tab(5); CustName$
  Print #RptHandle, Tab(5); "120 Main Street"
  Print #RptHandle, Tab(5); "PO Box 1190"
  Print #RptHandle, Tab(5); "Your Town, XX 55555-5555"

  For LC = 19 To 20 'made 18 = 19
    Print #RptHandle, " "
  Next LC
  Print #RptHandle, Tab(41); "LAND"; Tab(51); "BUILDING"; Tab(64); "TOTAL"; Tab(72); "TOTAL DUE"
  Print #RptHandle, " "
 'Line 23 Starts Here
  Print #RptHandle, "3 Acre Parcel";
  Print #RptHandle, Tab(30); Using("#0.00", 0.25);
  Print #RptHandle, Tab(37); Using("######0.00", 150000);
  Print #RptHandle, Tab(50); Using("#####0.00", 0);
  Print #RptHandle, Tab(61); Using("#####0.00", 150000);
  Print #RptHandle, Tab(71); Using("######0.00", 375)
  Print #RptHandle, "Riverside Township"

 'Lines 25 to 36 are blank
  For LCnt = 25 To 36: Print #RptHandle, "": Next LCnt
'Line 37 for Totals
  Print #RptHandle, Tab(48); "Total Tax Due ... ";
  Print #RptHandle, Using$("$######0.00", 375)
  Print #RptHandle, Tab(48); "Tax Due Date: " + CStr(Date)
  Print #RptHandle, ""
  Print #RptHandle,
  
  Print #RptHandle, "BN"; Using$("#####", 1)
  Print #RptHandle, "~"
  
  Close
  ViewPrint RptFile$, "Real Tax Bill Print Test", True
End Sub

