VERSION 5.00
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmPOs 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PO List"
   ClientHeight    =   2160
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   6060
   Icon            =   "frmPOs.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin LpLib.fpList fplstPOs 
      Height          =   1035
      Left            =   570
      TabIndex        =   0
      Top             =   285
      Width           =   4890
      _Version        =   196608
      _ExtentX        =   8625
      _ExtentY        =   1826
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
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
      BorderStyle     =   2
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   2
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
      ScrollBarH      =   3
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
      ColDesigner     =   "frmPOs.frx":08CA
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3144
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1470
      Width           =   1092
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D0D0D0&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4536
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1488
      Width           =   1092
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Double-Click Item or Highlight and Click Ok to Select."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   564
      Index           =   0
      Left            =   24
      TabIndex        =   3
      Top             =   1488
      Width           =   2940
   End
End
Attribute VB_Name = "frmPOs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Dim Vendor As VendorRecType
Private Temp_Class As Resize_Class
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'  If ((UnloadMode = vbFormControlMenu)) Then
'    Cancel = True
'  End If
'End Sub

Private Sub cmdExit_Click()
  Unload frmPOs
End Sub
Private Sub cmdOk_Click()
  Dim TempDist As Long, TempLeg As Long
  If fplstPOs.ListCount > 0 Then
  If fplstPOs.Selected Then
    TempDist = Val(Mid$(fplstPOs.Text, 50))
    If TempDist > 0 Then
      frmInvEnterEdit.fptxtPo.Text = ""
      TempLeg = Val(Mid$(fplstPOs.Text, 40, 10))
      frmInvEnterEdit.GetPOInfo TempDist, TempLeg
    Else
      MsgBox "You Must Select A PO", vbOKOnly, "No Selection"
      Exit Sub
    End If
  End If
  End If
  Unload frmPOs
  frmInvEnterEdit.fpInvAmt.SetFocus
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
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%O"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  
End Sub


Private Sub fplstPOs_DblClick()
  Dim TempDist As Long, TempLeg As Long
  frmInvEnterEdit.fptxtPo.Text = ""
  TempLeg = Val(Mid$(fplstPOs.Text, 40, 10))
  TempDist = Val(Mid$(fplstPOs.Text, 50))
  If TempDist > 0 Then
    frmInvEnterEdit.GetPOInfo TempDist, TempLeg
    Unload frmPOs
    frmInvEnterEdit.fpInvAmt.SetFocus
  Else
    MsgBox "You Must Select A PO", vbOKOnly, "No Selection"
  End If
End Sub

Private Sub fplstPOs_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim TempDist As Long, TempLeg As Long
  If KeyCode = vbKeyReturn Then
    If fplstPOs.Selected Then
      frmInvEnterEdit.fptxtPo.Text = ""
      TempLeg = Mid$(fplstPOs.Text, 40, 10)
      TempDist = Mid$(fplstPOs.Text, 50)
      If TempDist > 0 Then
        frmInvEnterEdit.GetPOInfo TempDist, TempLeg
        Unload frmPOs
        frmInvEnterEdit.fpInvAmt.SetFocus
      Else
        MsgBox "You Must Select A PO", vbOKOnly, "No Selection"
      End If
    End If
  End If
End Sub
Public Sub Loadpos()
  Dim VendorFile As Integer, NumVRecs As Integer, VRecNum As Integer
  Dim Last As Integer, cnt As Integer, Dcnt As Integer, TmpAcct As Integer
  frmInvEnterEdit.fpcboVendName.col = 2
  VRecNum = QPTrim(frmInvEnterEdit.fpcboVendName.ColText)
  If VRecNum > 0 Then
    OpenVendorFile VendorFile, NumVRecs
    Get VendorFile, VRecNum, Vendor
    FindPO VRecNum
  End If
 Close
End Sub
Public Sub FindPO(vrec As Integer)
  Dim POCnt As Integer, NextTrans As Long, fmt As String
  Dim VendorFile As Integer, NumVRecs As Integer, tempstr As String
  Dim APLedgerFile As Integer, NumTrans As Long, LdRecLen As Integer
  Dim APEditFile As Integer, NumEdTrans As Integer, CntInv, oktolist As Boolean
  Dim APIED As APInv85Type
  Dim APLedgerRec(1) As APLedger81RecType
  LdRecLen = Len(APLedgerRec(1))
  OpenAPEditFile APEditFile, NumEdTrans
  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTrans, LdRecLen
  fmt = "$ ###,###,###.##"
  Get VendorFile, vrec, Vendor
  NextTrans& = Vendor.FrstTran
  Do Until NextTrans& = 0
    Get APLedgerFile, NextTrans&, APLedgerRec(1)
    If APLedgerRec(1).TRCode = 4 Then
      POCnt = POCnt + 1
      'For CntInv = 1 To NumEdTrans
        'Get APEditFile, CntInv, APIED
        'If Not APIED.DELFLAG Then
        'If APIED.POAPLRecNum = NextTrans& Then  'Used on another invoice this edit
         'But if on this invoice bring up
         'If APIED.POAPLRecNum <> frmInvEnterEdit.fpAPLegNum Then
         ' If APIED.POAPLRecNum > 0 Then
         ' POCnt = POCnt - 1
          'End If
         'End If
        'End If
       ' End If
      'Next
    End If
    NextTrans& = APLedgerRec(1).NextTrans

  Loop
  If POCnt <> 0 Then

  NextTrans& = Vendor.FrstTran
  fplstPOs.Clear
  Do Until NextTrans& = 0
    Get APLedgerFile, NextTrans&, APLedgerRec(1)
    If APLedgerRec(1).TRCode = 4 Then
'      If NumEdTrans > 0 Then
'      For CntInv = 1 To NumEdTrans
'        Get APEditFile, CntInv, APIED
'        If Not APIED.DELFLAG Then
'        If APIED.POAPLRecNum <> NextTrans& Or APIED.POAPLRecNum = frmInvEnterEdit.fpAPLegNum Then
          oktolist = True
        Else
          oktolist = False
          'Exit For
        'End If
        'End If
      'Next
      'Else
        'oktolist = True
      End If
      If oktolist = True Then
          tempstr = Space$(60)
          Mid$(tempstr, 1) = QPTrim$(APLedgerRec(1).PONum)
          Mid$(tempstr, 10) = Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
          Mid$(tempstr, 20) = Using(fmt, Str$(APLedgerRec(1).Amt))
          Mid$(tempstr, 40) = NextTrans&
          Mid$(tempstr, 50) = APLedgerRec(1).FrstDist
          fplstPOs.AddItem tempstr
      End If
    
    'End If
    NextTrans& = APLedgerRec(1).NextTrans
  Loop
  
  Else
    
    MsgBox "No Purchase Orders For This Vendor", vbOKOnly, "No PO's"
    Unload frmPOs
  End If
  Close
End Sub

