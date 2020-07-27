VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmExpCustomerInfo2 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export Customer Information"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   2172
   ClientWidth     =   12216
   Icon            =   "frmExpCustInfo2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CheckBox ChkBank 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Include Bank Draft Info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   4824
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   4560
      Width           =   2772
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "F10 &Ok"
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
      Left            =   8424
      TabIndex        =   3
      Top             =   6744
      Width           =   1332
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
      Left            =   10128
      TabIndex        =   4
      Top             =   6744
      Width           =   1332
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
            TextSave        =   "10:14 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "5/17/2005"
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
   Begin EditLib.fpText fptxtRoute2 
      Height          =   348
      Left            =   6852
      TabIndex        =   1
      Top             =   3804
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   2
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "0123456789"
      MaxLength       =   2
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtRoute1 
      Height          =   348
      Left            =   5052
      TabIndex        =   0
      Top             =   3804
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   2
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "0123456789"
      MaxLength       =   2
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label LabelB1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Book Range for Export: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   3720
      TabIndex        =   9
      Top             =   3240
      Width           =   4716
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Export Customer Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3624
      TabIndex        =   8
      Top             =   1296
      Width           =   5004
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3216
      Top             =   1056
      Width           =   5772
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   324
      Index           =   0
      Left            =   4056
      TabIndex        =   7
      Top             =   3804
      Width           =   828
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   324
      Index           =   2
      Left            =   6144
      TabIndex        =   6
      Top             =   3804
      Width           =   540
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      Height          =   2220
      Left            =   2880
      Top             =   2976
      Width           =   6468
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3216
      Top             =   936
      Width           =   5772
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuPrnScn 
         Caption         =   "Prin&t Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmExpCustomerInfo2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via Expcustinfo by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
'    Case vbKeyDown, vbKeyReturn:
'      SendKeys "{Tab}"
'      KeyCode = 0
'    Case vbKeyUp:
'      SendKeys "+{Tab}"
'      KeyCode = 0
    Case vbKeyEscape:
      cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      cmdOk_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub fptxtRoute1_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtRoute1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtRoute2.SetFocus
  End If
End Sub
Private Sub fptxtRoute2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    CmdOk.SetFocus
  End If
End Sub

Private Sub fptxtRoute2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Function ValidRoutes()
  If fptxtRoute1 <> "" And fptxtRoute2 <> "" Then
    If fptxtRoute1 > fptxtRoute2 Then
      MsgBox "Invalid Selection, The Beginning Book Should Be Less or Equal to Ending Book.", vbOKOnly, "Invalid Selection"
      ValidRoutes = False
    Else
      ValidRoutes = True
    End If
  Else
    MsgBox "Book May Not Be Left Blank.", vbOKOnly, "Invalid Selection"
  End If
End Function

Private Sub cmdExit_Click()
  frmUBExportMenu.Show
  Unload frmExpCustomerInfo
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$

End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
Private Sub cmdOk_Click()
'
  If ValidRoutes Then
    DeActivateControls Me, True
'do consumpstuff here
    ' ExpCustStuff
   'For Johnston
     ExpCustStuffJohnst
    ActivateControls Me, True
  End If
End Sub

Private Sub ExpCustStuff()
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
  Dim MCnt As Integer, tempTot As Double
  FrmShowPctComp.Label1 = "Creating Export Files"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  q$ = Chr$(34)
  qc$ = q$ + ","
  qcq$ = q$ + "," + q$
  'special for johnston co
'  qcq$ = "|"
  Bookone = Val(QPTrim(fptxtRoute1))
  Bookto = Val(QPTrim(fptxtRoute2))
  IndexName$ = BookIndexFile
  UsingBook = True
  OKFlag = True
  BuckFmt$ = "########.##"
  NumofRevs = GetNumOfRevs%
  
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  ReDim UBOwnerRec(1) As UBOwnerRecType
  UBOwnerRecLen = Len(UBOwnerRec(1))

  IdxRecLen = 4               'we are using a long integer
  IdxFileSize& = FileSize(IndexName$)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
  'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
  NumOfRecs = IdxNumOfRecs

  Handle = FreeFile
  Open IndexName$ For Random Shared As Handle Len = IdxRecLen
  For cnt = 1 To IdxNumOfRecs
    Get #Handle, cnt, IdxBuff(cnt)
  Next
  Close Handle

  UBFile = FreeFile
  Open UBPath$ + "UBOWNER.DAT" For Random Shared As UBFile Len = UBOwnerRecLen

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  KillFile (UBPath$ + "UBCustEx.ASC")
  UBRpt = FreeFile
  Open UBPath$ + "UBCustEx.ASC" For Output As UBRpt
  GoSub DoHeaders
  For cnt = 1 To NumOfRecs
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out Then
      Close
      Unload FrmShowPctComp
      GoTo ExitMastCustListing
    End If

    AcctNumber = IdxBuff(cnt).RecNum
    Get UBCust, AcctNumber, UBCustRec(1)
    Get UBFile, AcctNumber, UBOwnerRec(1)

    '*************************************
    '   Main body of Printing goes here
    
'' 'This is just to temp export for Johnston Co.
''    If UBCustRec(1).DelFlag <> -1 Then
''        GoSub ExportThisAccount
''    End If
''
    If UBCustRec(1).DelFlag <> -1 And UBCustRec(1).Status = "A" Then
      ThisBook$ = UBCustRec(1).Book
      If Left$(ThisBook$, 1) = "0" Then
        WhatBook = Val(Right$(ThisBook$, 1))
      Else
        WhatBook = Val(ThisBook$)
      End If
      If WhatBook <= Bookto And WhatBook >= Bookone Then
        Export& = Export& + 1
        GoSub ExportThisAccount
      End If
    End If
  Next

  Close
  If Export& > 0 Then
    MsgBox "File " & UBPath$ & "UBCustEx.ASC Exported with " & Export& & " Accounts.", vbOKOnly, "Export Completed."
  Else
    MsgBox "No Information Found to Export.", vbOKOnly, "Procedure Ended"
  End If
GoTo ExitMastCustListing
  
ExportThisAccount:

  LocationNumber$ = QPTrim$(UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB)
  Zip$ = QPTrim$(UBCustRec(1).ZIPCODE)
  If Len(Zip$) > 5 Then
    Zip$ = Left$(Zip$, 5) + "-" + Mid$(Zip$, 6)
  End If

  Print #UBRpt, q$; QPTrim$(Str$(AcctNumber));
  Print #UBRpt, qcq$; LocationNumber$;
  Print #UBRpt, qcq$; UBCustRec(1).Status;
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).CustName);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).ADDR1);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).ADDR2);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).ServAddr);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).CITY);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).STATE);
  Print #UBRpt, qcq$; Zip$;
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).HPHONE);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).WPHONE);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).SOSEC);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).DRVLIC);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).CUSTTYPE);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).Addr911);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).BillTo);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).POSTRTE);
  'BILLCYCL      AS INTEGER
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).ZONE);
  'SEQ           AS LONG
  'Page 2
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).CASHONLY);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).LATEFEE);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).CUTOFFYN);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).TAXEXPT);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).SRCIT);
  'EPPFlag
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).BILLCMNT);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).PAYCMNT);

  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).PumpCode);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).USERCODE1);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).USERCODE2);
  If UBCustRec(1).ProRatePCT > 0 Then
    Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).ProRatePCT));
  Else
    Print #UBRpt, qcq$; "100";
  End If

  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).HHMSG1);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).HHMSG2);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).HHMSG3);

  For RCnt = 1 To NumofRevs
    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).serv(RCnt).Ratecode);
    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).serv(RCnt).RMtrType);
  Next

'flatrates
  For FCnt = 1 To 4
    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).FlatRates(FCnt).FRDESC);
    If UBCustRec(1).FlatRates(FCnt).FRAMT > 0 Then
      Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).FlatRates(FCnt).FRAMT));
    Else
      Print #UBRpt, qcq$; "0.00";
    End If
    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).FlatRates(FCnt).FRFREQ);

    If UBCustRec(1).FlatRates(FCnt).REVSRC > 0 Then
      Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).FlatRates(FCnt).REVSRC));
    Else
      Print #UBRpt, qcq$; "0";
    End If
    If UBCustRec(1).FlatRates(FCnt).NumMin > 0 Then
      Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).FlatRates(FCnt).NumMin));
    Else
      Print #UBRpt, qcq$; "0";
    End If
  Next

'meters
  For MCnt = 1 To 7
    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).LocMeters(MCnt).MtrNum);
    If UBCustRec(1).LocMeters(MCnt).MTRMulti > 0 Then
      Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).LocMeters(MCnt).MTRMulti));
    Else
      Print #UBRpt, qcq$; "0";
    End If
    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).LocMeters(MCnt).MtrType);
    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).LocMeters(MCnt).MtrUnit);

    If UBCustRec(1).LocMeters(MCnt).NumUser > 0 Then
      Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).LocMeters(MCnt).NumUser));
    Else
      Print #UBRpt, qcq$; "0";
    End If
    If UBCustRec(1).LocMeters(MCnt).InsDate > 0 Then
      Print #UBRpt, qcq$; Num2Date$(UBCustRec(1).LocMeters(MCnt).InsDate);
    Else
      Print #UBRpt, qcq$; "??/??/????";
    End If
    If UBCustRec(1).LocMeters(MCnt).CurRead > 0 Then
      Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).LocMeters(MCnt).CurRead));
    Else
      Print #UBRpt, qcq$; "0";
    End If

    If UBCustRec(1).LocMeters(MCnt).PrevRead > 0 Then
      Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).LocMeters(MCnt).PrevRead));
    Else
      Print #UBRpt, qcq$; "0";
    End If

    If UBCustRec(1).LocMeters(MCnt).CurDate > 0 Then
      Print #UBRpt, qcq$; Num2Date$(UBCustRec(1).LocMeters(MCnt).CurDate);
    Else
      Print #UBRpt, qcq$; "??/??/????";
    End If
    If UBCustRec(1).LocMeters(MCnt).PastDate > 0 Then
      Print #UBRpt, qcq$; Num2Date$(UBCustRec(1).LocMeters(MCnt).PastDate);
    Else
      Print #UBRpt, qcq$; "??/??/????";
    End If
'    ReadFlag  AS STRING * 1    'hidden & protected
'    AvgUse    AS LONG          'hidden & protected
'    UseCnt    AS INTEGER       'hidden & protected
  Next

  Print #UBRpt, qcq$; Using(BuckFmt$, UBCustRec(1).CurrBalance);
  Print #UBRpt, qcq$; Using(BuckFmt$, UBCustRec(1).PrevBalance);

  For RCnt = 1 To NumofRevs
    Print #UBRpt, qcq$; Using(BuckFmt$, UBCustRec(1).CurrRevAmts(RCnt));
  Next

  Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).OwnLName);
  Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).OwnFName);
  Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).ADDR1);
  Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).ADDR2);
  Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).CITY);
  Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).STATE);
  Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).ZIPCODE);
  Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).HPHONE);
  Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).WPHONE);
  'don't want the bankdraft info
  If ChkBank.Value = 1 Then
    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).USEDRAFT);
    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).AcctType);
    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).BankName);
    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).BANKLOC);
    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).TRANSIT);
    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).BankAcct);
  End If
  Print #UBRpt, qcq$; Using(BuckFmt$, UBCustRec(1).DepositAmt);
  Print #UBRpt, qcq$; Num2Date$(UBCustRec(1).OPENDATE);
  'Print #UBRpt,
  Print #UBRpt, q$

Return

DoHeaders:
  Print #UBRpt, q$; "Account";
  Print #UBRpt, qcq$; "Location";
  Print #UBRpt, qcq$; "Status";
  Print #UBRpt, qcq$; "CustName";
  Print #UBRpt, qcq$; "ADDR1";
  Print #UBRpt, qcq$; "ADDR2";
  Print #UBRpt, qcq$; "ServAddr";
  Print #UBRpt, qcq$; "CITY";
  Print #UBRpt, qcq$; "STATE";
  Print #UBRpt, qcq$; "Zip";
  Print #UBRpt, qcq$; "HPHONE";
  Print #UBRpt, qcq$; "WPHONE";
  Print #UBRpt, qcq$; "SOSEC";
  Print #UBRpt, qcq$; "DRVLIC";
  Print #UBRpt, qcq$; "CUSTTYPE";
  Print #UBRpt, qcq$; "Addr911";
  Print #UBRpt, qcq$; "BillTo";
  Print #UBRpt, qcq$; "POSTRTE";
  'BILLCYCL      AS INTEGER
  Print #UBRpt, qcq$; "ZONE";
  'SEQ           AS LONG
  'Page 2
  Print #UBRpt, qcq$; "CASHONLY";
  Print #UBRpt, qcq$; "LATEFEE";
  Print #UBRpt, qcq$; "CUTOFFYN";
  Print #UBRpt, qcq$; "TAXEXPT";
  Print #UBRpt, qcq$; "SRCIT";
  'EPPFlag
  Print #UBRpt, qcq$; "BILLCMNT";
  Print #UBRpt, qcq$; "PAYCMNT";

  Print #UBRpt, qcq$; "PumpCode";
  Print #UBRpt, qcq$; "USERCODE1";
  Print #UBRpt, qcq$; "USERCODE2";
'  If UBCustRec(1).ProRatePCT > 0 Then
'    Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).ProRatePCT));
'  Else
'    Print #UBRpt, qcq$; "100";
'  End If
  Print #UBRpt, qcq$; "ProRate";
  Print #UBRpt, qcq$; "HHMSG1";
  Print #UBRpt, qcq$; "HHMSG2";
  Print #UBRpt, qcq$; "HHMSG3";

  For RCnt = 1 To NumofRevs
    Print #UBRpt, qcq$; Str(RCnt) + "RATECODE";
    Print #UBRpt, qcq$; "RMtrType";
  Next

'flatrates
  For FCnt = 1 To 4
    Print #UBRpt, qcq$; Str(FCnt) + "FRDESC";
   ' If UBCustRec(1).FlatRates(FCnt).FRAMT > 0 Then
      Print #UBRpt, qcq$; "FRAMT";
    'Else
   '   Print #UBRpt, qcq$; "0.00";
    'End If
    Print #UBRpt, qcq$; "FRFREQ";

    'If UBCustRec(1).FlatRates(FCnt).REVSRC > 0 Then
      Print #UBRpt, qcq$; "REVSRC";
   ' Else
   '   Print #UBRpt, qcq$; "0";
   ' End If
   ' If UBCustRec(1).FlatRates(FCnt).NumMin > 0 Then
      Print #UBRpt, qcq$; "NumMin";
  '  Else
  '    Print #UBRpt, qcq$; "0";
  '  End If
  Next

'meters
  For MCnt = 1 To 7
    Print #UBRpt, qcq$; Str(MCnt) + "MtrNum";
    'If UBCustRec(1).LocMeters(MCnt).MTRMulti > 0 Then
      Print #UBRpt, qcq$; "MTRMulti";
   ' Else
   '   Print #UBRpt, qcq$; "0";
   ' End If
    Print #UBRpt, qcq$; "MtrType";
    Print #UBRpt, qcq$; "MTRUnit";

   ' If UBCustRec(1).LocMeters(MCnt).NumUser > 0 Then
      Print #UBRpt, qcq$; "NumUser";
   ' Else
   '   Print #UBRpt, qcq$; "0";
   ' End If
  '  If UBCustRec(1).LocMeters(MCnt).InsDate > 0 Then
      Print #UBRpt, qcq$; "InsDate";
  '  Else
  '    Print #UBRpt, qcq$; "??/??/????";
  '  End If
  '  If UBCustRec(1).LocMeters(MCnt).CurRead > 0 Then
      Print #UBRpt, qcq$; "CurRead";
  '  Else
  '    Print #UBRpt, qcq$; "0";
  '  End If

  '  If UBCustRec(1).LocMeters(MCnt).PrevRead > 0 Then
      Print #UBRpt, qcq$; "PrevRead";
  '  Else
  '    Print #UBRpt, qcq$; "0";
  '  End If

   ' If UBCustRec(1).LocMeters(MCnt).CurDate > 0 Then
      Print #UBRpt, qcq$; "CurDate";
  '  Else
   '   Print #UBRpt, qcq$; "??/??/????";
  '  End If
'    If UBCustRec(1).LocMeters(MCnt).PastDate > 0 Then
      Print #UBRpt, qcq$; "PastDate";
 '   Else
 '     Print #UBRpt, qcq$; "??/??/????";
 '   End If
'    ReadFlag  AS STRING * 1    'hidden & protected
'    AvgUse    AS LONG          'hidden & protected
'    UseCnt    AS INTEGER       'hidden & protected
  Next

  Print #UBRpt, qcq$; "CurrBalance";
  Print #UBRpt, qcq$; "PrevBalance";

  For RCnt = 1 To NumofRevs
    Print #UBRpt, qcq$; "CurrRevAmts";
  Next

  Print #UBRpt, qcq$; "OwnLName";
  Print #UBRpt, qcq$; "OwnFName";
  Print #UBRpt, qcq$; "ADDR1";
  Print #UBRpt, qcq$; "ADDR2";
  Print #UBRpt, qcq$; "CITY";
  Print #UBRpt, qcq$; "STATE";
  Print #UBRpt, qcq$; "ZIPCODE";
  Print #UBRpt, qcq$; "HPHONE";
  Print #UBRpt, qcq$; "WPHONE";
  'don't want the bankdraft info
  If ChkBank.Value = 1 Then
    Print #UBRpt, qcq$; "USEDRAFT";
    Print #UBRpt, qcq$; "AcctType";
    Print #UBRpt, qcq$; "BankName";
    Print #UBRpt, qcq$; "BANKLOC";
    Print #UBRpt, qcq$; "TRANSIT";
    Print #UBRpt, qcq$; "BankAcct";
  End If
  Print #UBRpt, qcq$; "Deposit";
  Print #UBRpt, qcq$; "OpenDate";
  Print #UBRpt, q$
Return
ExitMastCustListing:

End Sub
Private Sub ExpCustStuffJohnst()
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
  Dim MCnt As Integer, tempTot As Double
  FrmShowPctComp.Label1 = "Creating Export Files"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
'   'special for johnston co
  q$ = ""
  qcq$ = "|"
  
'  q$ = Chr$(34)
'  qc$ = q$ + ","
'  qcq$ = q$ + "," + q$
  Bookone = Val(QPTrim(fptxtRoute1))
  Bookto = Val(QPTrim(fptxtRoute2))
  IndexName$ = BookIndexFile
  UsingBook = True
  OKFlag = True
  BuckFmt$ = "########.##"
  NumofRevs = GetNumOfRevs%
  
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  ReDim UBOwnerRec(1) As UBOwnerRecType
  UBOwnerRecLen = Len(UBOwnerRec(1))

  IdxRecLen = 4               'we are using a long integer
  IdxFileSize& = FileSize(IndexName$)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
  'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
  NumOfRecs = IdxNumOfRecs

  Handle = FreeFile
  Open IndexName$ For Random Shared As Handle Len = IdxRecLen
  For cnt = 1 To IdxNumOfRecs
    Get #Handle, cnt, IdxBuff(cnt)
  Next
  Close Handle

  UBFile = FreeFile
  Open UBPath$ + "UBOWNER.DAT" For Random Shared As UBFile Len = UBOwnerRecLen

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  
  KillFile (UBPath$ + "UBCustEx.ASC")
  UBRpt = FreeFile
  Open UBPath$ + "UBCustEx.ASC" For Output As UBRpt
  GoSub DoHeaders
  For cnt = 1 To NumOfRecs
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out Then
      Close
      Unload FrmShowPctComp
      GoTo ExitMastCustListing
    End If

    AcctNumber = IdxBuff(cnt).RecNum
    Get UBCust, AcctNumber, UBCustRec(1)
    Get UBFile, AcctNumber, UBOwnerRec(1)

    '*************************************
    '   Main body of Printing goes here
    
'''This is just to temp export for Johnston Co.
    If UBCustRec(1).DelFlag <> -1 Then
        GoSub ExportThisAccount
    End If
      
''''''''''''''''''''''
      
'    If UBCustRec(1).DelFlag <> -1 And UBCustRec(1).Status = "A" Then
'      ThisBook$ = UBCustRec(1).Book
'      If Left$(ThisBook$, 1) = "0" Then
'        WhatBook = Val(Right$(ThisBook$, 1))
'      Else
'        WhatBook = Val(ThisBook$)
'      End If
'      If WhatBook <= Bookto And WhatBook >= Bookone Then
'        Export& = Export& + 1
'        GoSub ExportThisAccount
'      End If
'    End If
  Next

  Close
  If Export& > 0 Then
    MsgBox "File " & UBPath$ & "UBCustEx.ASC Exported with " & Export& & " Accounts.", vbOKOnly, "Export Completed."
  Else
    MsgBox "No Information Found to Export.", vbOKOnly, "Procedure Ended"
  End If
GoTo ExitMastCustListing
  
ExportThisAccount:

  LocationNumber$ = QPTrim$(UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB)
  Zip$ = QPTrim$(UBCustRec(1).ZIPCODE)
  If Len(Zip$) > 5 Then
    Zip$ = Left$(Zip$, 5) + "-" + Mid$(Zip$, 6)
  End If

  Print #UBRpt, q$; QPTrim$(Str$(AcctNumber));
  Print #UBRpt, qcq$; LocationNumber$;
  Print #UBRpt, qcq$; UBCustRec(1).Status;
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).CustName);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).ADDR1);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).ADDR2);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).ServAddr);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).CITY);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).STATE);
  Print #UBRpt, qcq$; Zip$;
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).HPHONE);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).WPHONE);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).SOSEC);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).DRVLIC);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).CUSTTYPE);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).Addr911);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).BillTo);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).POSTRTE);
  'BILLCYCL      AS INTEGER
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).ZONE);
  'SEQ           AS LONG
  'Page 2
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).CASHONLY);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).LATEFEE);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).CUTOFFYN);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).TAXEXPT);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).SRCIT);
  'EPPFlag
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).BILLCMNT);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).PAYCMNT);

  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).PumpCode);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).USERCODE1);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).USERCODE2);
  If UBCustRec(1).ProRatePCT > 0 Then
    Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).ProRatePCT));
  Else
    Print #UBRpt, qcq$; "100";
  End If

  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).HHMSG1);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).HHMSG2);
  Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).HHMSG3);

  For RCnt = 1 To NumofRevs
    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).serv(RCnt).Ratecode);
    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).serv(RCnt).RMtrType);
  Next

'flatrates
  For FCnt = 1 To 4
    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).FlatRates(FCnt).FRDESC);
    If UBCustRec(1).FlatRates(FCnt).FRAMT > 0 Then
      Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).FlatRates(FCnt).FRAMT));
    Else
      Print #UBRpt, qcq$; "0.00";
    End If
    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).FlatRates(FCnt).FRFREQ);

    If UBCustRec(1).FlatRates(FCnt).REVSRC > 0 Then
      Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).FlatRates(FCnt).REVSRC));
    Else
      Print #UBRpt, qcq$; "0";
    End If
    If UBCustRec(1).FlatRates(FCnt).NumMin > 0 Then
      Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).FlatRates(FCnt).NumMin));
    Else
      Print #UBRpt, qcq$; "0";
    End If
  Next

'meters
  For MCnt = 1 To 7
    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).LocMeters(MCnt).MtrNum);
    If UBCustRec(1).LocMeters(MCnt).MTRMulti > 0 Then
      Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).LocMeters(MCnt).MTRMulti));
    Else
      Print #UBRpt, qcq$; "0";
    End If
    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).LocMeters(MCnt).MtrType);
    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).LocMeters(MCnt).MtrUnit);

    If UBCustRec(1).LocMeters(MCnt).NumUser > 0 Then
      Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).LocMeters(MCnt).NumUser));
    Else
      Print #UBRpt, qcq$; "0";
    End If
    If UBCustRec(1).LocMeters(MCnt).InsDate > 0 Then
      Print #UBRpt, qcq$; Num2Date$(UBCustRec(1).LocMeters(MCnt).InsDate);
    Else
      Print #UBRpt, qcq$; "??/??/????";
    End If
    If UBCustRec(1).LocMeters(MCnt).CurRead > 0 Then
      Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).LocMeters(MCnt).CurRead));
    Else
      Print #UBRpt, qcq$; "0";
    End If

    If UBCustRec(1).LocMeters(MCnt).PrevRead > 0 Then
      Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).LocMeters(MCnt).PrevRead));
    Else
      Print #UBRpt, qcq$; "0";
    End If

    If UBCustRec(1).LocMeters(MCnt).CurDate > 0 Then
      Print #UBRpt, qcq$; Num2Date$(UBCustRec(1).LocMeters(MCnt).CurDate);
    Else
      Print #UBRpt, qcq$; "??/??/????";
    End If
    If UBCustRec(1).LocMeters(MCnt).PastDate > 0 Then
      Print #UBRpt, qcq$; Num2Date$(UBCustRec(1).LocMeters(MCnt).PastDate);
    Else
      Print #UBRpt, qcq$; "??/??/????";
    End If
'    ReadFlag  AS STRING * 1    'hidden & protected
'    AvgUse    AS LONG          'hidden & protected
'    UseCnt    AS INTEGER       'hidden & protected
  Next

  Print #UBRpt, qcq$; Using(BuckFmt$, UBCustRec(1).CurrBalance);
  Print #UBRpt, qcq$; Using(BuckFmt$, UBCustRec(1).PrevBalance);

  For RCnt = 1 To NumofRevs
    Print #UBRpt, qcq$; Using(BuckFmt$, UBCustRec(1).CurrRevAmts(RCnt));
  Next

  Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).OwnLName);
  Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).OwnFName);
  Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).ADDR1);
  Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).ADDR2);
  Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).CITY);
  Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).STATE);
  Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).ZIPCODE);
  Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).HPHONE);
  Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).WPHONE);
  'don't want the bankdraft info
  If ChkBank.Value = 1 Then
    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).USEDRAFT);
    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).AcctType);
    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).BankName);
    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).BANKLOC);
    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).TRANSIT);
    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).BankAcct);
  End If
  Print #UBRpt, qcq$; Using(BuckFmt$, UBCustRec(1).DepositAmt);
  Print #UBRpt, qcq$; Num2Date$(UBCustRec(1).OPENDATE);
  Print #UBRpt, qcq$; UBCustRec(1).BILLCYCL;
  Print #UBRpt, q$

Return

DoHeaders:
  Print #UBRpt, q$; "Account";
  Print #UBRpt, qcq$; "Location";
  Print #UBRpt, qcq$; "Status";
  Print #UBRpt, qcq$; "CustName";
  Print #UBRpt, qcq$; "ADDR1";
  Print #UBRpt, qcq$; "ADDR2";
  Print #UBRpt, qcq$; "ServAddr";
  Print #UBRpt, qcq$; "CITY";
  Print #UBRpt, qcq$; "STATE";
  Print #UBRpt, qcq$; "Zip";
  Print #UBRpt, qcq$; "HPHONE";
  Print #UBRpt, qcq$; "WPHONE";
  Print #UBRpt, qcq$; "SOSEC";
  Print #UBRpt, qcq$; "DRVLIC";
  Print #UBRpt, qcq$; "CUSTTYPE";
  Print #UBRpt, qcq$; "Addr911";
  Print #UBRpt, qcq$; "BillTo";
  Print #UBRpt, qcq$; "POSTRTE";
  'BILLCYCL      AS INTEGER
  Print #UBRpt, qcq$; "ZONE";
  'SEQ           AS LONG
  'Page 2
  Print #UBRpt, qcq$; "CASHONLY";
  Print #UBRpt, qcq$; "LATEFEE";
  Print #UBRpt, qcq$; "CUTOFFYN";
  Print #UBRpt, qcq$; "TAXEXPT";
  Print #UBRpt, qcq$; "SRCIT";
  'EPPFlag
  Print #UBRpt, qcq$; "BILLCMNT";
  Print #UBRpt, qcq$; "PAYCMNT";

  Print #UBRpt, qcq$; "PumpCode";
  Print #UBRpt, qcq$; "USERCODE1";
  Print #UBRpt, qcq$; "USERCODE2";
'  If UBCustRec(1).ProRatePCT > 0 Then
'    Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).ProRatePCT));
'  Else
'    Print #UBRpt, qcq$; "100";
'  End If
  Print #UBRpt, qcq$; "ProRate";
  Print #UBRpt, qcq$; "HHMSG1";
  Print #UBRpt, qcq$; "HHMSG2";
  Print #UBRpt, qcq$; "HHMSG3";

  For RCnt = 1 To NumofRevs
    Print #UBRpt, qcq$; Str(RCnt) + "RATECODE";
    Print #UBRpt, qcq$; "RMtrType";
  Next

'flatrates
  For FCnt = 1 To 4
    Print #UBRpt, qcq$; Str(FCnt) + "FRDESC";
   ' If UBCustRec(1).FlatRates(FCnt).FRAMT > 0 Then
      Print #UBRpt, qcq$; "FRAMT";
    'Else
   '   Print #UBRpt, qcq$; "0.00";
    'End If
    Print #UBRpt, qcq$; "FRFREQ";

    'If UBCustRec(1).FlatRates(FCnt).REVSRC > 0 Then
      Print #UBRpt, qcq$; "REVSRC";
   ' Else
   '   Print #UBRpt, qcq$; "0";
   ' End If
   ' If UBCustRec(1).FlatRates(FCnt).NumMin > 0 Then
      Print #UBRpt, qcq$; "NumMin";
  '  Else
  '    Print #UBRpt, qcq$; "0";
  '  End If
  Next

'meters
  For MCnt = 1 To 7
    Print #UBRpt, qcq$; Str(MCnt) + "MtrNum";
    'If UBCustRec(1).LocMeters(MCnt).MTRMulti > 0 Then
      Print #UBRpt, qcq$; "MTRMulti";
   ' Else
   '   Print #UBRpt, qcq$; "0";
   ' End If
    Print #UBRpt, qcq$; "MtrType";
    Print #UBRpt, qcq$; "MTRUnit";

   ' If UBCustRec(1).LocMeters(MCnt).NumUser > 0 Then
      Print #UBRpt, qcq$; "NumUser";
   ' Else
   '   Print #UBRpt, qcq$; "0";
   ' End If
  '  If UBCustRec(1).LocMeters(MCnt).InsDate > 0 Then
      Print #UBRpt, qcq$; "InsDate";
  '  Else
  '    Print #UBRpt, qcq$; "??/??/????";
  '  End If
  '  If UBCustRec(1).LocMeters(MCnt).CurRead > 0 Then
      Print #UBRpt, qcq$; "CurRead";
  '  Else
  '    Print #UBRpt, qcq$; "0";
  '  End If

  '  If UBCustRec(1).LocMeters(MCnt).PrevRead > 0 Then
      Print #UBRpt, qcq$; "PrevRead";
  '  Else
  '    Print #UBRpt, qcq$; "0";
  '  End If

   ' If UBCustRec(1).LocMeters(MCnt).CurDate > 0 Then
      Print #UBRpt, qcq$; "CurDate";
  '  Else
   '   Print #UBRpt, qcq$; "??/??/????";
  '  End If
'    If UBCustRec(1).LocMeters(MCnt).PastDate > 0 Then
      Print #UBRpt, qcq$; "PastDate";
 '   Else
 '     Print #UBRpt, qcq$; "??/??/????";
 '   End If
'    ReadFlag  AS STRING * 1    'hidden & protected
'    AvgUse    AS LONG          'hidden & protected
'    UseCnt    AS INTEGER       'hidden & protected
  Next

  Print #UBRpt, qcq$; "CurrBalance";
  Print #UBRpt, qcq$; "PrevBalance";

  For RCnt = 1 To NumofRevs
    Print #UBRpt, qcq$; "CurrRevAmts";
  Next

  Print #UBRpt, qcq$; "OwnLName";
  Print #UBRpt, qcq$; "OwnFName";
  Print #UBRpt, qcq$; "ADDR1";
  Print #UBRpt, qcq$; "ADDR2";
  Print #UBRpt, qcq$; "CITY";
  Print #UBRpt, qcq$; "STATE";
  Print #UBRpt, qcq$; "ZIPCODE";
  Print #UBRpt, qcq$; "HPHONE";
  Print #UBRpt, qcq$; "WPHONE";
  'don't want the bankdraft info
  If ChkBank.Value = 1 Then
    Print #UBRpt, qcq$; "USEDRAFT";
    Print #UBRpt, qcq$; "AcctType";
    Print #UBRpt, qcq$; "BankName";
    Print #UBRpt, qcq$; "BANKLOC";
    Print #UBRpt, qcq$; "TRANSIT";
    Print #UBRpt, qcq$; "BankAcct";
  End If
  Print #UBRpt, qcq$; "Deposit";
  Print #UBRpt, qcq$; "OpenDate";
  Print #UBRpt, qcq$; "Bill Cycle";
  Print #UBRpt, q$
Return
ExitMastCustListing:

End Sub


