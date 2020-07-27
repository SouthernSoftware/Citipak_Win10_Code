VERSION 5.00
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPOApproval 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Approve Purchase Orders"
   ClientHeight    =   8850
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   Icon            =   "frmPOApproval.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fplstPOs 
      Height          =   2310
      Left            =   2820
      TabIndex        =   0
      Top             =   2835
      Width           =   6600
      _Version        =   196608
      _ExtentX        =   11642
      _ExtentY        =   4075
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
      MultiSelect     =   1
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
      ColDesigner     =   "frmPOApproval.frx":08CA
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Escape E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   8757
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7272
      Width           =   1740
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 &Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   6405
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7272
      Width           =   1740
   End
   Begin VB.CommandButton cmdMark 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Alt-M &Mark All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   1695
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7272
      Width           =   1740
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Alt-C &Clear All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   4053
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7272
      Width           =   1740
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   5
      Top             =   8604
      Width           =   12192
      _ExtentX        =   21511
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7117
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "5:01 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "8/29/2008"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Purchase Orders For Approval"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3498
      TabIndex        =   8
      Top             =   1008
      Width           =   5196
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3036
      Top             =   768
      Width           =   6132
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "* Press SpaceBar or Mouse to Toggle"
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
      Height          =   348
      Left            =   2256
      TabIndex        =   7
      Top             =   6648
      Width           =   4332
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Purchase Orders From List:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   492
      Left            =   2496
      TabIndex        =   6
      Top             =   2208
      Width           =   3900
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   2820
      Left            =   2634
      Top             =   2640
      Width           =   6924
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      Height          =   972
      Left            =   3036
      Top             =   648
      Width           =   6132
   End
   Begin VB.Label lbInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Once Approved, Purchase Orders May NOT Be Edited!!! "
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
      Height          =   420
      Index           =   2
      Left            =   3216
      TabIndex        =   9
      Top             =   5904
      Width           =   5964
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   612
      Left            =   2856
      Top             =   5736
      Width           =   6540
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
Attribute VB_Name = "frmPOApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim POControl As POControlRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim Dept As String
Private Sub cmdClear_Click()
  fplstPOs.Action = ActionDeselectAll
End Sub

Private Sub cmdExit_Click()
  frmPOProcessMenu.Show
  Unload frmPOApproval
End Sub

Private Sub cmdMark_Click()
  fplstPOs.Action = ActionSelectAll
End Sub

Private Sub GetPONums()
  Dim POFile As Integer, POFileLen As Integer, NumRecs As Integer
  Dim POEditFile As Integer, NumEdTrans As Integer, cnt As Integer
  Dim PONumber As String, KK As Integer, OLDSTUFF As String
  Dim Pcnt As Integer, Rec As Integer, PONumb As Long
  Dim PO As POFORMRecType2
  ReDim POCont(1) As POControlRecType
  OpenPOFile POFile, NumRecs
  If LOF(POFile) > 0 Then
    Get POFile, 1, POCont(1)

    OpenPOEditFile POEditFile, NumEdTrans

    For Pcnt = 0 To fplstPOs.ListCount - 1
    If fplstPOs.Selected(Pcnt) Then
      cnt = cnt + 1
      fplstPOs.Row = Pcnt
      Rec = QPTrim(Mid$(fplstPOs.ColList, 55, 5))
      Get POEditFile, Rec, PO
      If PO.LOCKED = False Then
        PO.PONum = POCont(1).PONumber
        PO.Deleted = 1
        Put POEditFile, Rec, PO
        PONumber$ = POCont(1).PONumber
        KK = InStr(PONumber$, "-")
        If KK > 0 Then
          OLDSTUFF$ = Left$(PONumber$, KK)
          PONumber$ = Mid$(PONumber$, KK + 1, Len(PONumber$) - KK + 1)
        End If
        PONumb& = Val(PONumber$)
        PONumb& = PONumb& + 1
        PONumber$ = LTrim$(Str$(PONumb&))
        If KK > 0 Then
          PONumber$ = OLDSTUFF$ + PONumber$
        End If
        'Close POFile
        'OpenPOFile POFile, NumRecs
        POCont(1).PONumber = PONumber$
        Put POFile, 1, POCont(1)
      Else
        Close
        MsgBox "Purchase Order Being Edited-Make Sure All Operators Exit POEntry/Edit Before Trying Approval Again.", vbOKOnly, "Procedure Canceled"
        Exit Sub
      End If
    End If
  Next Pcnt
  End If
  Close

  Exit Sub

End Sub

Private Sub cmdSave_Click()
  If fplstPOs.SelCount > 0 Then
    GetPONums
  Else
    If MsgBox("No PO's Were Selected. Do You Wish to Retry or Exit?", vbRetryCancel, "Retry?") = vbRetry Then
      Exit Sub
    End If
  End If
  frmPOProcessMenu.Show
  Unload frmPOApproval
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog "Close AP"
      ClearInUse PWcnt
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
      cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      cmdSave_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  Me.HelpContextID = hlpAppPO
  POsList fplstPOs
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Public Sub SetDept(DeptPick As String)
  If QPTrim(DeptPick) = "ALL" Then
    Dept = 0
  Else
    Dept = QPTrim(DeptPick)
  End If
End Sub
  Private Function POsList(x As fpList)
  Dim POEdit As POFORMRecType2
  Dim POEditFile As Integer, NumEdTrans As Integer, cnt As Integer
  Dim Tmpstr As String, fmt As String
  OpenPOEditFile POEditFile, NumEdTrans
  fmt$ = "$########.##"
  For cnt = 1 To NumEdTrans
    Get POEditFile, cnt, POEdit
    If POEdit.Deleted <> True And Left$(POEdit.PONum, 3) = "N/A" Then
      If Val(Dept) <> 0 Then
        If QPTrim(POEdit.REQNUM) = Dept Then
          Tmpstr = Space$(60)
          Mid$(Tmpstr, 1, 8) = Left$(POEdit.VNDRCODE, 8)
          Mid$(Tmpstr, 15, 20) = Left$(POEdit.VNDRINF1, 20)
          Mid$(Tmpstr, 38, 15) = Using(fmt$, POEdit.POAmt)
          Mid$(Tmpstr, 55, 5) = Str$(cnt)
          x.AddItem Tmpstr
        End If
      Else
        Tmpstr = Space$(60)
        Mid$(Tmpstr, 1, 8) = Left$(POEdit.VNDRCODE, 8)
        Mid$(Tmpstr, 15, 20) = Left$(POEdit.VNDRINF1, 20)
        Mid$(Tmpstr, 38, 15) = Using(fmt$, POEdit.POAmt)
        Mid$(Tmpstr, 55, 5) = Str$(cnt)
        x.AddItem Tmpstr
      End If
    End If
  Next
  Close POEditFile
End Function

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
