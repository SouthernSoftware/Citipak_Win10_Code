VERSION 5.00
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmFixedAssetsConversion 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fixed Assets Conversion"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmFixedAssetsConversion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpList1 
      Height          =   2355
      Left            =   3120
      TabIndex        =   5
      Top             =   5955
      Width           =   5295
      _Version        =   196608
      _ExtentX        =   9340
      _ExtentY        =   4154
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Columns         =   0
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
      SelMax          =   -1
      AutoSearch      =   1
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
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
      ScrollBarV      =   1
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   0
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   2
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   1
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
      ColDesigner     =   "frmFixedAssetsConversion.frx":08CA
   End
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   4044
      Left            =   2652
      TabIndex        =   0
      Top             =   1044
      Width           =   6156
      _Version        =   196609
      _ExtentX        =   10858
      _ExtentY        =   7133
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.27
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12648447
      Caption         =   ""
      Picture         =   "frmFixedAssetsConversion.frx":0B56
      Begin VB.CommandButton cmdHelp 
         Caption         =   "F1 &Turn Help On"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   636
         Left            =   384
         TabIndex        =   8
         Top             =   2016
         Width           =   2604
      End
      Begin VB.CommandButton cmdVersion 
         Caption         =   "F5 Check &Version"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   636
         Left            =   3216
         TabIndex        =   6
         Top             =   1104
         Width           =   2604
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "ESC E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   636
         Left            =   3216
         TabIndex        =   3
         Top             =   2016
         Width           =   2604
      End
      Begin VB.CommandButton cmdConvertNow 
         Caption         =   "F10 &Convert Now"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   636
         Left            =   384
         TabIndex        =   2
         Top             =   1104
         Width           =   2604
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Fixed Asset Conversion has completed."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   684
         Left            =   1344
         TabIndex        =   4
         Top             =   2976
         Width           =   3612
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Fixed Assets Conversion"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   396
         Left            =   1056
         TabIndex        =   1
         Top             =   336
         Width           =   4284
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   $"frmFixedAssetsConversion.frx":0B72
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3420
      Left            =   9120
      TabIndex        =   9
      Top             =   3504
      Width           =   2268
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   8736
      X2              =   10128
      Y1              =   2592
      Y2              =   4896
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   $"frmFixedAssetsConversion.frx":0C6E
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2892
      Left            =   144
      TabIndex        =   10
      Top             =   3504
      Width           =   2268
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   1104
      X2              =   2976
      Y1              =   4656
      Y2              =   2496
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   $"frmFixedAssetsConversion.frx":0D29
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   924
      Left            =   576
      TabIndex        =   11
      Top             =   48
      Width           =   10668
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "List Of Vendor Names"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   348
      Left            =   3024
      TabIndex        =   7
      Top             =   5376
      Width           =   5484
   End
End
Attribute VB_Name = "frmFixedAssetsConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class

Public Sub Begin()
  Dim DOSFAItemRec As DosFAItemRecType
  Dim DOSFAItemRecLen As Integer
  Dim DOSFAHandle As Integer
  Dim DOSFAItemRecV1 As DosFAItemRecTypeV1
  Dim DOSFAItemRecLenV1 As Integer
  Dim DOSFAHandleV1 As Integer
  Dim FAItemRec As FAItemRecType
  Dim FAItemRecLen As Integer
  Dim FAHandle As Integer
  Dim NumOfFARecs As Long
  Dim DOSCodeRec As DosFAAssetCodeRecType
  Dim DOSACHandle As Integer
  Dim DOSCodeRecLen As Integer
  Dim CODEREC As FAAssetCodeRecType
  Dim ACHandle As Integer
  Dim CodeRecLen As Integer
  Dim FASetup As FASetupRecType
  Dim SetUpHandle As Integer
  Dim SetUpRecLen As Integer
  Dim DOSFASetup As DosFASetupRecType
  Dim DOSSetUpHandle As Integer
  Dim DOSSetUpRecLen As Integer
  Dim x As Long
  Dim TEMPTownName As String * 25
  Dim TEMPPct1St     As Integer
  Dim TEMPPRate1St As String * 1
  Dim TEMPFiller1    As String * 100
  Dim BigDept As Integer
  Dim NrmlDpr As Double
  Dim CurrentVal As Double
  Dim LifeLeft As Double
  Dim Remain As Double
  Dim ConvertAssCodes As Integer
  
  ConvertAssCodes = CheckAssCodes
  If ConvertAssCodes = 2 Or ConvertAssCodes = 4 Then
    Exit Sub
  End If
  
  If ThisVersion = 1 Then
    frmFAMsgWOpts.Label1.Caption = "This data being converted does not include: 1) End Of Life date and 2) Option to include/exclude for depreciation. Both of these fields exist in the update. The data being converted does include three fixed asset item descriptions where the update only has two. If you continue converting this" _
      & " data then all fixed assets will be earmarked for depreciation and the End Of Life date will be calculated based on each fixed asset's life expectancy and when each was purchased. The last of the three descriptions will not be included in this conversion. If you" _
      & " wish to continue with the conversion press F10. Otherwise press ESC to abort the conversion."
    frmFAMsgWOpts.cmdCont.Text = "F10 Continue"
    frmFAMsgWOpts.cmdExit.Text = "ESC Abort"
    frmFAMsgWOpts.Show vbModal
    If frmFAMsgWOpts.fptxtChoice.Text = "abort" Then
      Unload frmFAMsgWOpts
      Exit Sub
    End If
  End If
  
  BigDept = 0
  Label2.Visible = True
  DoEvents
  Label2.Caption = "Extracting Departments"
  DoEvents
  
  If ThisVersion = 0 Then
    Call ExtractDepts(BigDept)
  ElseIf ThisVersion = 1 Then
    Call ExtractDeptsV1(BigDept)
  End If
  
  If ThisVersion = 0 Then
    DOSFAItemRecLen = Len(DOSFAItemRec)
    DOSFAHandle = FreeFile
    Open "FAITEMS.DAT" For Random Shared As DOSFAHandle Len = DOSFAItemRecLen
    
    NumOfFARecs = LOF(DOSFAHandle) / Len(DOSFAItemRec)
    
    ReDim TEMPITEMTAG(1 To NumOfFARecs) As String * 20
    ReDim TEMPISTATUS(1 To NumOfFARecs) As String * 1
    ReDim TEMPDEPYN(1 To NumOfFARecs) As String * 1
    ReDim TEMPAQURDATE(1 To NumOfFARecs) As Integer
    ReDim TEMPIDESC1(1 To NumOfFARecs) As String * 30
    ReDim TEMPIDESC2(1 To NumOfFARecs) As String * 30
    ReDim TEMPGLACCT(1 To NumOfFARecs) As String * 14
    ReDim TEMPIDEPT(1 To NumOfFARecs) As String * 4
    ReDim TEMPASSETCODE(1 To NumOfFARecs) As String * 4
    ReDim TEMPILIFE(1 To NumOfFARecs) As Double
    ReDim TEMPORGCOST(1 To NumOfFARecs) As Double
    ReDim TEMPDEP2DATE(1 To NumOfFARecs) As Double
    ReDim TEMPCURRVAL(1 To NumOfFARecs) As Double
    ReDim TEMPCDEPDATE(1 To NumOfFARecs) As Integer
    ReDim TEMPDISPDATE(1 To NumOfFARecs) As Integer
    ReDim TEMPVENDOR(1 To NumOfFARecs) As String * 30
    ReDim TEMPSERIALNO(1 To NumOfFARecs) As String * 30
    ReDim TEMPITEMMFG(1 To NumOfFARecs) As String * 30
    ReDim TEMPCONTACT(1 To NumOfFARecs) As String * 30
    ReDim TEMPITEMLOC(1 To NumOfFARecs) As String * 30
    ReDim TEMPEOLDATE(1 To NumOfFARecs) As Integer
    ReDim TEMPFill1(1 To NumOfFARecs) As String * 86
    ReDim TEMPFileVer(1 To NumOfFARecs) As Integer
    
    Label2.Caption = "Converting Item Data"
    DoEvents
  
    For x = 1 To NumOfFARecs
      Get DOSFAHandle, x, DOSFAItemRec
      TEMPITEMTAG(x) = QPTrim$(DOSFAItemRec.ITEMTAG)
      TEMPISTATUS(x) = QPTrim$(DOSFAItemRec.ISTATUS)
      TEMPDEPYN(x) = QPTrim$(DOSFAItemRec.DEPYN)
      TEMPAQURDATE(x) = DOSFAItemRec.AQURDATE
      TEMPIDESC1(x) = QPTrim$(DOSFAItemRec.IDESC1)
      TEMPIDESC2(x) = QPTrim$(DOSFAItemRec.IDESC2)
      TEMPGLACCT(x) = QPTrim$(DOSFAItemRec.GLACCT)
      If Len(QPTrim$(DOSFAItemRec.IDEPT)) = 0 Then
        TEMPIDEPT(x) = BigDept + 1
      Else
        TEMPIDEPT(x) = QPTrim$(DOSFAItemRec.IDEPT)
      End If
      TEMPASSETCODE(x) = QPTrim$(DOSFAItemRec.ASSETCODE)
      TEMPILIFE(x) = DOSFAItemRec.ILIFE
      TEMPORGCOST(x) = DOSFAItemRec.ORGCOST
      TEMPDEP2DATE(x) = DOSFAItemRec.DEP2DATE
      TEMPCURRVAL(x) = DOSFAItemRec.CURRVAL
      TEMPCDEPDATE(x) = DOSFAItemRec.CDEPDATE
      If CheckValDate(MakeRegDate(DOSFAItemRec.DISPDATE)) = False Then
        TEMPDISPDATE(x) = 0
      Else
        TEMPDISPDATE(x) = DOSFAItemRec.DISPDATE
      End If
      TEMPVENDOR(x) = QPTrim$(DOSFAItemRec.VENDOR)
      TEMPSERIALNO(x) = QPTrim$(DOSFAItemRec.SERIALNO)
      TEMPITEMMFG(x) = QPTrim$(DOSFAItemRec.ITEMMFG)
      TEMPCONTACT(x) = QPTrim$(DOSFAItemRec.CONTACT)
      TEMPITEMLOC(x) = QPTrim$(DOSFAItemRec.ITEMLOC)
      TEMPEOLDATE(x) = DOSFAItemRec.EOLDate
      TEMPFill1(x) = QPTrim$(DOSFAItemRec.Fill1)
    Next x
    
    Close DOSFAHandle
    
    FAItemRecLen = Len(FAItemRec)
    FAHandle = FreeFile
    Open "FAITEMS.DAT" For Random Shared As FAHandle Len = FAItemRecLen
    For x = 1 To NumOfFARecs
      FAItemRec.ITEMTAG = QPTrim$(TEMPITEMTAG(x))
      FAItemRec.ISTATUS = QPTrim$(TEMPISTATUS(x))
      FAItemRec.DEPYN = QPTrim$(TEMPDEPYN(x))
      FAItemRec.AQURDATE = TEMPAQURDATE(x)
      FAItemRec.IDESC1 = QPTrim$(TEMPIDESC1(x))
      FAItemRec.IDESC2 = QPTrim$(TEMPIDESC2(x))
      FAItemRec.GLACCT = QPTrim$(TEMPGLACCT(x))
      FAItemRec.IDEPT = Val(TEMPIDEPT(x))
      FAItemRec.ASSETCODE = QPTrim$(TEMPASSETCODE(x))
      FAItemRec.ILIFE = TEMPILIFE(x)
      FAItemRec.ORGCOST = TEMPORGCOST(x)
      FAItemRec.DEP2DATE = TEMPDEP2DATE(x)
      FAItemRec.CDEPDATE = TEMPCDEPDATE(x)
      FAItemRec.DISPDATE = TEMPDISPDATE(x)
      If TEMPDISPDATE(x) > 0 Then
        FAItemRec.DsplFlag = 2
        FAItemRec.CURRVAL = 0
      Else
        FAItemRec.DsplFlag = 0
        FAItemRec.CURRVAL = TEMPORGCOST(x) - TEMPDEP2DATE(x)
      End If
      FAItemRec.VENDOR = QPTrim$(TEMPVENDOR(x))
      FAItemRec.SERIALNO = QPTrim$(TEMPSERIALNO(x))
      FAItemRec.ITEMMFG = QPTrim$(TEMPITEMMFG(x))
      FAItemRec.CONTACT = QPTrim$(TEMPCONTACT(x))
      FAItemRec.ITEMLOC = QPTrim$(TEMPITEMLOC(x))
      FAItemRec.EOLDate = TEMPEOLDATE(x)
      FAItemRec.VHCLMAKE = ""
      FAItemRec.VHCLMODL = ""
      FAItemRec.VHCLVIN = ""
      FAItemRec.VHCLTAG = ""
      FAItemRec.VHCLCOLR = ""
      FAItemRec.WARRXDAT = 0
      FAItemRec.Fill1 = QPTrim$(TEMPFill1(x))
      FAItemRec.FundNum = Val(Mid(FAItemRec.GLACCT, 1, 2))
      FAItemRec.DisposAmt = 0
      FAItemRec.LastDprRec = 0
      If TEMPDISPDATE(x) > 0 Then
        FAItemRec.LifeLeft = 0
      ElseIf TEMPDEP2DATE(x) = 0 Then
        FAItemRec.LifeLeft = TEMPILIFE(x)
      ElseIf TEMPDEP2DATE(x) = TEMPORGCOST(x) Then
        FAItemRec.LifeLeft = 0
      ElseIf TEMPILIFE(x) = 0 Then
        FAItemRec.LifeLeft = 0
      Else
        NrmlDpr = OldRound(TEMPORGCOST(x) / TEMPILIFE(x))
        CurrentVal = TEMPORGCOST(x) - TEMPDEP2DATE(x)
        If NrmlDpr > CurrentVal Then
          FAItemRec.LifeLeft = 1
        Else
          LifeLeft = CurrentVal / NrmlDpr
          LifeLeft = OldRound(LifeLeft)
          LifeLeft = OldRound(LifeLeft * 100)
          Remain = Right(CStr(LifeLeft), 2)
          If Val(Remain) = 0 Or Val(Remain) = 98 Or Val(Remain) = 99 Then
            LifeLeft = CurrentVal / NrmlDpr
            LifeLeft = CInt(LifeLeft)
            FAItemRec.LifeLeft = LifeLeft
          ElseIf Val(Remain) = 50 Then
            LifeLeft = OldRound(CurrentVal / NrmlDpr)
            LifeLeft = OldRound(LifeLeft + 0.5)
            FAItemRec.LifeLeft = LifeLeft
          Else
            LifeLeft = OldRound(LifeLeft / 100)
            Remain = Right(CStr(LifeLeft), 2)
            If InStr(1, Remain, ".") Then Remain = Remain * 100
            If Val(Remain) < 50 Then
              LifeLeft = CurrentVal / NrmlDpr
              LifeLeft = CInt(LifeLeft) + 1
              FAItemRec.LifeLeft = LifeLeft
            Else
              LifeLeft = CurrentVal / NrmlDpr
              FAItemRec.LifeLeft = CInt(LifeLeft)
              FAItemRec.LifeLeft = LifeLeft
            End If
          End If
        End If
      End If
      FAItemRec.PONum = "NA"
      FAItemRec.CheckNum = "NA"
      FAItemRec.DsplMethod = "NA"
      Put FAHandle, x, FAItemRec
    Next x
    
    Close FAHandle
    
  ElseIf ThisVersion = 1 Then
  
    DOSFAItemRecLenV1 = Len(DOSFAItemRecV1)
    DOSFAHandleV1 = FreeFile
    Open "FAITEMS.DAT" For Random Shared As DOSFAHandleV1 Len = DOSFAItemRecLenV1
    
    NumOfFARecs = LOF(DOSFAHandleV1) / Len(DOSFAItemRecV1)
    
    ReDim TEMPITEMTAG(1 To NumOfFARecs) As String * 20
    ReDim TEMPISTATUS(1 To NumOfFARecs) As String * 1
    ReDim TEMPDEPYN(1 To NumOfFARecs) As String * 1
    ReDim TEMPAQURDATE(1 To NumOfFARecs) As Integer
    ReDim TEMPIDESC1(1 To NumOfFARecs) As String * 30
    ReDim TEMPIDESC2(1 To NumOfFARecs) As String * 30
    ReDim TEMPGLACCT(1 To NumOfFARecs) As String * 14
    ReDim TEMPIDEPT(1 To NumOfFARecs) As String * 4
    ReDim TEMPASSETCODE(1 To NumOfFARecs) As String * 4
    ReDim TEMPILIFE(1 To NumOfFARecs) As Double
    ReDim TEMPORGCOST(1 To NumOfFARecs) As Double
    ReDim TEMPDEP2DATE(1 To NumOfFARecs) As Double
    ReDim TEMPCURRVAL(1 To NumOfFARecs) As Double
    ReDim TEMPCDEPDATE(1 To NumOfFARecs) As Integer
    ReDim TEMPDISPDATE(1 To NumOfFARecs) As Integer
    ReDim TEMPVENDOR(1 To NumOfFARecs) As String * 30
    ReDim TEMPSERIALNO(1 To NumOfFARecs) As String * 30
    ReDim TEMPITEMMFG(1 To NumOfFARecs) As String * 30
    ReDim TEMPCONTACT(1 To NumOfFARecs) As String * 30
    ReDim TEMPITEMLOC(1 To NumOfFARecs) As String * 30
    ReDim TEMPEOLDATE(1 To NumOfFARecs) As Integer
    ReDim TEMPFill1(1 To NumOfFARecs) As String * 86
    
    Label2.Caption = "Converting Item Data"
    DoEvents
  
    For x = 1 To NumOfFARecs
      Get DOSFAHandleV1, x, DOSFAItemRecV1
      TEMPITEMTAG(x) = QPTrim$(DOSFAItemRecV1.ITEMTAG)
      TEMPISTATUS(x) = QPTrim$(DOSFAItemRecV1.ISTATUS)
      TEMPDEPYN(x) = "Y"
      TEMPAQURDATE(x) = DOSFAItemRecV1.AQURDATE
      TEMPIDESC1(x) = QPTrim$(DOSFAItemRecV1.IDESC1)
      TEMPIDESC2(x) = QPTrim$(DOSFAItemRecV1.IDESC2)
      TEMPGLACCT(x) = QPTrim$(DOSFAItemRecV1.GLACCT)
      If Len(QPTrim$(DOSFAItemRecV1.IDEPT)) = 0 Then
        TEMPIDEPT(x) = BigDept + 1
      Else
        TEMPIDEPT(x) = QPTrim$(DOSFAItemRecV1.IDEPT)
      End If
      TEMPASSETCODE(x) = QPTrim$(DOSFAItemRecV1.ASSETCODE)
      TEMPILIFE(x) = DOSFAItemRecV1.ILIFE
      TEMPORGCOST(x) = DOSFAItemRecV1.ORGCOST
      TEMPDEP2DATE(x) = DOSFAItemRecV1.DEP2DATE
      TEMPCURRVAL(x) = 0
      TEMPCDEPDATE(x) = DOSFAItemRecV1.CDEPDATE
      If CheckValDate(MakeRegDate(DOSFAItemRecV1.DISPDATE)) = False Then
        TEMPDISPDATE(x) = 0
      Else
        TEMPDISPDATE(x) = DOSFAItemRecV1.DISPDATE
      End If
      TEMPVENDOR(x) = QPTrim$(DOSFAItemRecV1.VENDOR)
      TEMPSERIALNO(x) = QPTrim$(DOSFAItemRecV1.SERIALNO)
      TEMPITEMMFG(x) = QPTrim$(DOSFAItemRecV1.ITEMMFG)
      TEMPCONTACT(x) = QPTrim$(DOSFAItemRecV1.CONTACT)
      TEMPITEMLOC(x) = "Not Saved"
      TEMPEOLDATE(x) = 0
      TEMPFill1(x) = QPTrim$(DOSFAItemRecV1.Fill1)
    Next x
    
    Close DOSFAHandleV1
  
    FAItemRecLen = Len(FAItemRec)
    FAHandle = FreeFile
    Open "FAITEMS.DAT" For Random Shared As FAHandle Len = FAItemRecLen
    For x = 1 To NumOfFARecs
      FAItemRec.ITEMTAG = QPTrim$(TEMPITEMTAG(x))
      FAItemRec.ISTATUS = QPTrim$(TEMPISTATUS(x))
      FAItemRec.DEPYN = QPTrim$(TEMPDEPYN(x))
      FAItemRec.AQURDATE = TEMPAQURDATE(x)
      FAItemRec.IDESC1 = QPTrim$(TEMPIDESC1(x))
      FAItemRec.IDESC2 = QPTrim$(TEMPIDESC2(x))
      FAItemRec.GLACCT = QPTrim$(TEMPGLACCT(x))
      FAItemRec.IDEPT = Val(TEMPIDEPT(x))
      FAItemRec.ASSETCODE = QPTrim$(TEMPASSETCODE(x))
      FAItemRec.ILIFE = TEMPILIFE(x)
      FAItemRec.ORGCOST = TEMPORGCOST(x)
      FAItemRec.DEP2DATE = TEMPDEP2DATE(x)
      FAItemRec.CDEPDATE = TEMPCDEPDATE(x)
      FAItemRec.DISPDATE = TEMPDISPDATE(x)
      If TEMPDISPDATE(x) > 0 Then
        FAItemRec.DsplFlag = 2
        FAItemRec.CURRVAL = 0
      Else
        FAItemRec.DsplFlag = 0
        FAItemRec.CURRVAL = TEMPORGCOST(x) - TEMPDEP2DATE(x)
      End If
      FAItemRec.VENDOR = QPTrim$(TEMPVENDOR(x))
      FAItemRec.SERIALNO = QPTrim$(TEMPSERIALNO(x))
      FAItemRec.ITEMMFG = QPTrim$(TEMPITEMMFG(x))
      FAItemRec.CONTACT = QPTrim$(TEMPCONTACT(x))
      FAItemRec.ITEMLOC = QPTrim$(TEMPITEMLOC(x))
      FAItemRec.VHCLMAKE = ""
      FAItemRec.VHCLMODL = ""
      FAItemRec.VHCLVIN = ""
      FAItemRec.VHCLTAG = ""
      FAItemRec.VHCLCOLR = ""
      FAItemRec.WARRXDAT = 0
      FAItemRec.Fill1 = QPTrim$(TEMPFill1(x))
      FAItemRec.FundNum = Val(Mid(FAItemRec.GLACCT, 1, 2))
      FAItemRec.DisposAmt = 0
      FAItemRec.LastDprRec = 0
      If TEMPDISPDATE(x) > 0 Then
        FAItemRec.LifeLeft = 0
      ElseIf TEMPDEP2DATE(x) = 0 Then
        FAItemRec.LifeLeft = TEMPILIFE(x)
      ElseIf TEMPDEP2DATE(x) = TEMPORGCOST(x) Then
        FAItemRec.LifeLeft = 0
      ElseIf TEMPILIFE(x) = 0 Then
        FAItemRec.LifeLeft = 0
      Else
        NrmlDpr = OldRound(TEMPORGCOST(x) / TEMPILIFE(x))
        CurrentVal = TEMPORGCOST(x) - TEMPDEP2DATE(x)
        If NrmlDpr > CurrentVal Then
          FAItemRec.LifeLeft = 1
        Else
          LifeLeft = CurrentVal / NrmlDpr
          LifeLeft = OldRound(LifeLeft)
          LifeLeft = OldRound(LifeLeft * 100)
          Remain = Right(CStr(LifeLeft), 2)
          If Val(Remain) = 0 Or Val(Remain) = 98 Or Val(Remain) = 99 Then
            LifeLeft = CurrentVal / NrmlDpr
            LifeLeft = CInt(LifeLeft)
            FAItemRec.LifeLeft = LifeLeft
          ElseIf Val(Remain) = 50 Then
            LifeLeft = OldRound(CurrentVal / NrmlDpr)
            LifeLeft = OldRound(LifeLeft + 0.5)
            FAItemRec.LifeLeft = LifeLeft
          Else
            LifeLeft = OldRound(LifeLeft / 100)
            Remain = Right(CStr(LifeLeft), 2)
            If InStr(1, Remain, ".") Then Remain = Remain * 100
            If Val(Remain) < 50 Then
              LifeLeft = CurrentVal / NrmlDpr
              LifeLeft = CInt(LifeLeft) + 1
              FAItemRec.LifeLeft = LifeLeft
            Else
              LifeLeft = CurrentVal / NrmlDpr
              FAItemRec.LifeLeft = CInt(LifeLeft)
              FAItemRec.LifeLeft = LifeLeft
            End If
          End If
        End If
      End If
      FAItemRec.EOLDate = EOLDate(FAItemRec.AQURDATE, FAItemRec.ILIFE)
      FAItemRec.PONum = "NA"
      FAItemRec.CheckNum = "NA"
      FAItemRec.DsplMethod = "NA"
      Put FAHandle, x, FAItemRec
    Next x
  
    Close FAHandle
  End If
  
  If ConvertAssCodes = 3 Then GoTo DontConvertAssCode
  
  Label2.Caption = "Converting Asset Codes"
  DoEvents

  DOSCodeRecLen = Len(DOSCodeRec)
  DOSACHandle = FreeFile
  Open "FACODES.DAT" For Random Shared As DOSACHandle Len = DOSCodeRecLen
  Dim NumOfCodes As Integer
  NumOfCodes = LOF(DOSACHandle) \ Len(DOSCodeRec)
  ReDim TEMPASSETCODE(1 To NumOfCodes) As String * 4
  ReDim TEMPAssetStatus(1 To NumOfCodes) As String * 10
  ReDim TEMPAssetDesc(1 To NumOfCodes) As String * 20
  For x = 1 To NumOfCodes
    Get DOSACHandle, x, DOSCodeRec
    If QPTrim$(DOSCodeRec.ASSETCODE) = "A" Then
      TEMPASSETCODE(x) = "0000"
    Else
      TEMPASSETCODE(x) = QPTrim$(DOSCodeRec.ASSETCODE)
    End If
    TEMPAssetStatus(x) = QPTrim$(DOSCodeRec.AssetStatus)
    TEMPAssetDesc(x) = QPTrim$(DOSCodeRec.AssetDesc)
  Next x
  Close DOSACHandle

  CodeRecLen = Len(CODEREC)
  ACHandle = FreeFile
  Open "FACODES.DAT" For Random Shared As ACHandle Len = CodeRecLen
  For x = 1 To NumOfCodes
    Get ACHandle, x, CODEREC
    CODEREC.ASSETCODE = TEMPASSETCODE(x)
    CODEREC.AssetDesc = TEMPAssetDesc(x)
    CODEREC.AssetStatus = TEMPAssetStatus(x)
    Put ACHandle, x, CODEREC
  Next x
  Close ACHandle
  
DontConvertAssCode:

  DOSSetUpRecLen = Len(DOSFASetup)
  DOSSetUpHandle = FreeFile
  Open "FASETUP.DAT" For Random Shared As DOSSetUpHandle Len = DOSSetUpRecLen
  Get DOSSetUpHandle, 1, DOSFASetup
    TEMPTownName = QPTrim$(DOSFASetup.TownName)
    TEMPPct1St = DOSFASetup.Pct1St
    TEMPPRate1St = QPTrim$(DOSFASetup.PRate1St)
    TEMPFiller1 = QPTrim$(DOSFASetup.Filler1)
  Close DOSSetUpHandle

  SetUpRecLen = Len(FASetup)
  SetUpHandle = FreeFile
  Open "FASETUP.DAT" For Random Shared As SetUpHandle Len = SetUpRecLen
  FASetup.TownName = QPTrim$(TEMPTownName)
  FASetup.Pct1St = CDbl(TEMPPct1St)
  FASetup.PRate1St = QPTrim$(TEMPPRate1St)
  FASetup.Filler1 = QPTrim$(TEMPFiller1)
  FASetup.DeprType = "NOT SAVED"
  Put SetUpHandle, 1, FASetup
  Close SetUpHandle
  Label2.Visible = True
  
  If ConvertAssCodes <> 3 Then
    Label2.Caption = "Creating Asset Code Index"
    DoEvents
    Call CreateAssetIdx
  End If
  
  Label2.Caption = "Creating Tag Item Index...could take a while"
  DoEvents
  Call CreateTagIdx
  Label2.Caption = "Conversion completed successfully!"
  DoEvents
End Sub

Private Sub cmdConvertNow_Click()
  Call Begin
End Sub

Private Sub cmdExit_Click()
  End
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Caption, "On") Then
    cmdHelp.Caption = "F1 &Turn Help Off"
    Label4.Visible = True
    Label6.Visible = True
    Label7.Visible = True
    Line1.Visible = True
    Line2.Visible = True
  Else
    cmdHelp.Caption = "F1 &Turn Help On"
    Label4.Visible = False
    Label6.Visible = False
    Label7.Visible = False
    Line1.Visible = False
    Line2.Visible = False
  End If
End Sub

Private Sub cmdVersion_Click()
  If InStr(cmdVersion.Caption, "2") Then
    cmdVersion.Caption = "F5 Check Version 1"
    ThisVersion = 1
    Call CheckVersion
    cmdConvertNow.Caption = "F10 Convert Vers #2"
  Else
    cmdVersion.Caption = "F5 Check Version 2"
    ThisVersion = 0
    Call CheckVersion
    cmdConvertNow.Caption = "F10 Convert Vers #1"
  End If
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Label2.Visible = False
  ThisVersion = 0
  cmdVersion.Caption = "F5 Check Version 1"
  cmdConvertNow.Enabled = False
  Label4.Visible = False
  Label6.Visible = False
  Label7.Visible = False
  Line1.Visible = False
  Line2.Visible = False
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
    Case vbKeyEscape:
      Call cmdExit_Click
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyF10:
      Call Begin
      SendKeys "%C"
      KeyCode = 0
    Case vbKeyF5:
      Call cmdVersion_Click
      SendKeys "%V"
      KeyCode = 0
    Case vbKeyF1:
      Call cmdHelp_Click
      SendKeys "%T"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Public Sub ExtractDepts(ByRef BigDept As Integer)
  Dim FAItemRec As DosFAItemRecType
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim x As Integer
  Dim Nextx As Integer
  Dim ThisDept As Integer
  Dim Y As Integer
  Dim DeptHandle As Integer
  Dim DeptRec As FADeptCodeType
  Dim EmptyStringFlag As Boolean
  
  EmptyStringFlag = False
  OpenFAItemFile FAHandle
  NumOfFARecs = LOF(FAHandle) / Len(FAItemRec)
  For x = 1 To NumOfFARecs
    Get FAHandle, x, FAItemRec
    If Val(QPTrim$(FAItemRec.IDEPT)) > 0 Then
      ThisDept = CInt(FAItemRec.IDEPT)
      Exit For
    End If
  Next x
  ReDim Depts(1 To NumOfFARecs) As Integer

  Nextx = 1
  Depts(Nextx) = ThisDept
  NumOfDepts = NumOfDepts + 1
  Y = 1
  Do
    For x = 1 To NumOfFARecs
      Get FAHandle, x, FAItemRec
        If Len(QPTrim$(FAItemRec.IDEPT)) = 0 Then
          EmptyStringFlag = True
          GoTo EmptyString
        End If
        For Y = 1 To NumOfFARecs
          If Depts(Y) = 0 Then GoTo EmptyY
          If Val(FAItemRec.IDEPT) = Depts(Y) Then
            GoTo FoundIt
          End If
EmptyY:
        Next Y
FoundIt:
        If Y = NumOfFARecs + 1 Then
          If IsNumeric(FAItemRec.IDEPT) = False Then GoTo EmptyString
          NumOfDepts = NumOfDepts + 1
          Depts(NumOfDepts) = Val(QPTrim(FAItemRec.IDEPT))
        End If
      If x = NumOfFARecs Then Exit Do
EmptyString:
    Next x
  Loop
  Close FAHandle
  
  OpenFADeptCodeFile DeptHandle
  
  For x = 1 To NumOfDepts
    If Depts(x) > BigDept Then
      BigDept = Depts(x)
    End If
    DeptRec.DeptDesc = "UNKNOWN"
    DeptRec.DeptNum = Depts(x)
    Put DeptHandle, x, DeptRec
    Debug.Print Depts(x)
  Next x
  
  If EmptyStringFlag = True Then
    DeptRec.DeptNum = BigDept + 1
    DeptRec.DeptDesc = "UNKNOWN"
    Put DeptHandle, NumOfDepts + 1, DeptRec
  End If
  
  Close DeptHandle
  Call CreateDeptIdx
End Sub

Public Sub CreateDeptIdx()
  Dim BigNum As Integer
  Dim ThisNum As Integer
  Dim ThisX As Integer
  Dim SmallNum As Integer
  Dim x As Integer
  Dim Nextx As Integer
  Dim DeptHandle As Integer
  Dim DeptRec As FADeptCodeType
  Dim DeptItemRecLen As Integer
  Dim NumOfDeptRecs As Integer
  Dim DeptIdx As DeptNumbSortIdxType
  Dim DeptIdxHandle As Integer
  Dim DeptIdxRecNum As Integer
  Dim RecNum As Integer
  Dim HoldThis As DeptNumbSortIdxType
  
  OpenFADeptCodeFile DeptHandle
  
  NumOfDeptRecs = LOF(DeptHandle) \ Len(DeptRec)
  ReDim TempDeptIdx(1 To NumOfDeptRecs) As DeptNumbSortIdxType
  
  BigNum = 0
  For x = 1 To NumOfDeptRecs
    Get DeptHandle, x, DeptRec
    TempDeptIdx(x).DeptRecNum = x
    TempDeptIdx(x).DeptNumb = DeptRec.DeptNum
    TempDeptIdx(x).DeptIdxDesc = DeptRec.DeptDesc
    ThisNum = DeptRec.DeptNum
    If ThisNum > BigNum Then
      BigNum = ThisNum
    End If
BadNum:
  Next x
  Close DeptHandle
  
  BigNum = BigNum + 1
  SmallNum = BigNum
  Nextx = 1
  
  Do
    For x = Nextx To NumOfDeptRecs
      ThisNum = TempDeptIdx(x).DeptNumb
      If ThisNum < SmallNum Then
        SmallNum = ThisNum
        ThisX = x
      End If
    Next x
    HoldThis = TempDeptIdx(Nextx)
    TempDeptIdx(Nextx) = TempDeptIdx(ThisX)
    TempDeptIdx(ThisX) = HoldThis
    If Nextx = NumOfDeptRecs Then Exit Do
    Nextx = Nextx + 1
    SmallNum = BigNum
  Loop
  
  OpenDeptIdxFile DeptIdxHandle
  For x = 1 To NumOfDeptRecs
    DeptIdx = TempDeptIdx(x)
'    Debug.Print DeptIdx.DeptNumb
    Put DeptIdxHandle, x, DeptIdx
  Next x
  
  Close

End Sub

Public Sub ExtractDeptsV1(ByRef BigDept As Integer)
  Dim FAItemRec As DosFAItemRecTypeV1
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim x As Integer
  Dim Nextx As Integer
  Dim ThisDept As Integer
  Dim Y As Integer
  Dim DeptHandle As Integer
  Dim DeptRec As FADeptCodeType
  Dim EmptyStringFlag As Boolean
  
  EmptyStringFlag = False
  OpenFAItemFileV1 FAHandle
  NumOfFARecs = LOF(FAHandle) / Len(FAItemRec)
  For x = 1 To NumOfFARecs
    Get FAHandle, x, FAItemRec
    If Val(QPTrim$(FAItemRec.IDEPT)) > 0 Then
      ThisDept = CInt(FAItemRec.IDEPT)
      Exit For
    End If
  Next x
  ReDim Depts(1 To NumOfFARecs) As Integer

  Nextx = 1
  Depts(Nextx) = ThisDept
  NumOfDepts = NumOfDepts + 1
  Y = 1
  Do
    For x = 1 To NumOfFARecs
      Get FAHandle, x, FAItemRec
        If Len(QPTrim$(FAItemRec.IDEPT)) = 0 Then
          EmptyStringFlag = True
          GoTo EmptyString
        End If
        For Y = 1 To NumOfFARecs
          If Depts(Y) = 0 Then GoTo EmptyY
          If Val(FAItemRec.IDEPT) = Depts(Y) Then
            GoTo FoundIt
          End If
EmptyY:
        Next Y
FoundIt:
        If Y = NumOfFARecs + 1 Then
          If IsNumeric(FAItemRec.IDEPT) = False Then GoTo EmptyString
          NumOfDepts = NumOfDepts + 1
          Depts(NumOfDepts) = Val(QPTrim(FAItemRec.IDEPT))
        End If
      If x = NumOfFARecs Then Exit Do
EmptyString:
    Next x
  Loop
  Close FAHandle
  
  OpenFADeptCodeFile DeptHandle
  
  For x = 1 To NumOfDepts
    If Depts(x) > BigDept Then
      BigDept = Depts(x)
    End If
    DeptRec.DeptDesc = "UNKNOWN"
    DeptRec.DeptNum = Depts(x)
    Put DeptHandle, x, DeptRec
    Debug.Print Depts(x)
  Next x
  
  If EmptyStringFlag = True Then
    DeptRec.DeptNum = BigDept + 1
    DeptRec.DeptDesc = "UNKNOWN"
    Put DeptHandle, NumOfDepts + 1, DeptRec
  End If
  
  Close DeptHandle
  Call CreateDeptIdx
End Sub

Private Function EOLDate(Start As Integer, ILIFE As Double) As Integer
  Dim ThisYear$
  Dim StartYear$
  Dim EndYear$
  Dim YearDif As Double
  Dim ConvertYear$
  
  ThisYear$ = MakeRegDate(Start)
  StartYear$ = Mid(ThisYear$, 7, 4)
  YearDif = CDbl(StartYear) + ILIFE
  EndYear$ = Mid(ThisYear$, 1, 6) + CStr(YearDif)
  EOLDate = Date2Num(EndYear$)
End Function

Private Function CheckAssCodes() As Integer
  Dim CodeHandle As Integer
  Dim CODEREC As FAAssetCodeRecType
  Dim NumOfRecs As Integer
  Dim x As Integer
  
  CheckAssCodes = 1
  OpenFACodeNameFile CodeHandle
  NumOfRecs = LOF(CodeHandle) / Len(CODEREC)
  If NumOfRecs = 0 Then
    Close
    Exit Function
  End If
  
  ReDim NotNumber(1 To 1) As String
  ReDim NNDesc(1 To 1) As String
  ReDim NNRecNum(1 To 1) As Integer
  
  NumOfBad = 0
  For x = 1 To NumOfRecs
    Get CodeHandle, x, CODEREC
    If Not IsNumeric(CODEREC.ASSETCODE) Then
      NumOfBad = NumOfBad + 1
      ReDim Preserve NotNumber(1 To NumOfBad) As String
      ReDim Preserve NNDesc(1 To NumOfBad) As String
      ReDim Preserve NNRecNum(1 To NumOfBad) As Integer
      NotNumber(NumOfBad) = QPTrim$(CODEREC.ASSETCODE)
      NNDesc(NumOfBad) = QPTrim$(CODEREC.AssetDesc)
      NNRecNum(NumOfBad) = x
    End If
  Next x
  Close
  
  If NumOfBad > 0 Then
    frmFABadCodeWarn.Show vbModal
    If frmFABadCodeWarn.fptxtChoice.Text = "abort" Then
      Unload frmFABadCodeWarn
      CheckAssCodes = 2
      Exit Function
    ElseIf frmFABadCodeWarn.fptxtChoice.Text = "jump" Then
      Unload frmFABadCodeWarn
      CheckAssCodes = 4
      DoEvents
      frmFAAssCodeUtility.Show
      DoEvents
      Unload Me
      Exit Function
    Else
      Unload frmFABadCodeWarn
      CheckAssCodes = 3
    End If
  End If
End Function

