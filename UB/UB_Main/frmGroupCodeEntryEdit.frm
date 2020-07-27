VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Begin VB.Form frmGroupCodeEntryEdit 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Group Code Entry/Edit"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12225
   ClipControls    =   0   'False
   Icon            =   "frmGroupCodeEntryEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPrintCodes 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F5 &Print Codes"
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
      Left            =   5688
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7536
      Width           =   1836
   End
   Begin VB.CommandButton cmdAdd2List 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "F9 &Add To List"
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
      Left            =   8214
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   1908
   End
   Begin VB.CommandButton cmdAddNew 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F2 &New"
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
      Left            =   1386
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7536
      Width           =   1356
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F3 &Edit"
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
      Left            =   3042
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7536
      Width           =   1356
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D0D0D0&
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
      Left            =   9474
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7536
      Width           =   1356
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 &Save"
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
      Left            =   7818
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7536
      Width           =   1356
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   8385
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "5/15/2018"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "9:00 AM"
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
   Begin EditLib.fpText txtGroupNum 
      Height          =   372
      Left            =   3072
      TabIndex        =   0
      Top             =   2640
      Width           =   1092
      _Version        =   196608
      _ExtentX        =   1926
      _ExtentY        =   656
      Enabled         =   0   'False
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      ForeColor       =   0
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   8421504
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   8421504
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
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
      ThreeDTextHighlightColor=   14737632
      ThreeDTextShadowColor=   8421504
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
      InvalidColor    =   16777215
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   16777215
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   2
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   12632256
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   0
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   8421504
      BorderDropShadowWidth=   3
      ButtonColor     =   12632256
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText txtGroupDesc 
      Height          =   372
      Left            =   4494
      TabIndex        =   1
      Top             =   2640
      Width           =   3372
      _Version        =   196608
      _ExtentX        =   5948
      _ExtentY        =   656
      Enabled         =   0   'False
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      ForeColor       =   0
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   8421504
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
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
      ThreeDTextHighlightColor=   14737632
      ThreeDTextShadowColor=   8421504
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   1
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   16777215
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   16777215
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   30
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   12632256
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   4210752
      BorderDropShadowWidth=   3
      ButtonColor     =   12632256
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3060
      Left            =   2682
      TabIndex        =   3
      Top             =   3960
      Width           =   6852
      _Version        =   196613
      _ExtentX        =   12086
      _ExtentY        =   5397
      _StockProps     =   64
      ArrowsExitEditMode=   -1  'True
      BackColorStyle  =   1
      ButtonDrawMode  =   1
      ColsFrozen      =   3
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   13684944
      GridColor       =   8421504
      MaxCols         =   4
      ProcessTab      =   -1  'True
      ScrollBars      =   2
      ShadowColor     =   13684944
      ShadowDark      =   8421504
      ShadowText      =   0
      SpreadDesigner  =   "frmGroupCodeEntryEdit.frx":08CA
      VisibleCols     =   3
      ClipboardOptions=   4
   End
   Begin VB.Label lblEdit 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1488
      TabIndex        =   15
      Top             =   3360
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label lblNew 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1488
      TabIndex        =   14
      Top             =   1848
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Line Line1 
      X1              =   1854
      X2              =   10350
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   5580
      Left            =   1854
      Top             =   1680
      Width           =   8508
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      FillColor       =   &H00D0D0D0&
      Height          =   3204
      Left            =   2610
      Top             =   3888
      Width           =   6996
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Group Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   2502
      TabIndex        =   13
      Top             =   2256
      Width           =   1812
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Group Code Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   4374
      TabIndex        =   12
      Top             =   2280
      Width           =   3156
   End
   Begin VB.Label lblNewCode 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "New Group Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   1974
      TabIndex        =   11
      Top             =   1848
      Width           =   2340
   End
   Begin VB.Label lblEditCode 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Edit Existing Group Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   1974
      TabIndex        =   10
      Top             =   3384
      Width           =   3300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Group Code Entry/Edit"
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
      Left            =   4320
      TabIndex        =   8
      Top             =   744
      Width           =   3612
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   732
      Left            =   3240
      Top             =   552
      Width           =   5772
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      Height          =   852
      Left            =   3240
      Top             =   432
      Width           =   5772
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuPrnScrn 
         Caption         =   "&Print Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmGroupCodeEntryEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
'Dim GLSetup As GLSetupRecType
'Dim GLFund As GLFundRecType
'Dim GLAcct As GLAcctRecType
'Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Public RecordNum As Integer
Private Temp_Class As Resize_Class



Private Sub cmdAddNew_Click()
  Dim MaxNum As Integer
  Dim GroupCde As GroupCodeRecType
  Dim GrpCodeRecLen As Integer, ghandle As Integer, cnt As Integer
  Dim NumofGrps As Integer
  GrpCodeRecLen = Len(GroupCde)
  ghandle = FreeFile
  Open UBPath$ + "UBGrpCde.DAT" For Random Shared As ghandle Len = GrpCodeRecLen
  NumofGrps = LOF(ghandle) \ GrpCodeRecLen
  RecordNum = NumofGrps + 1
  lblNew.Visible = True
  lblEdit.Visible = False
  txtGroupNum.Enabled = True
  txtGroupDesc.Enabled = True
  cmdAdd2List.Enabled = True
  vaSpread1.Lock = True
  txtGroupNum.SetFocus
End Sub

Private Sub cmdEdit_Click()
  lblNew.Visible = False
  lblEdit.Visible = True
  txtGroupNum.Enabled = False
  txtGroupDesc.Enabled = False
  cmdAdd2List.Enabled = False
  vaSpread1.Lock = False
  vaSpread1.SetFocus
End Sub

Private Sub cmdPrintCodes_Click()
  Dim FntSize As Integer
  ReDim MsgText(0 To 5) As String
  If Not Exist(UBPath$ + "UBGrpCde.dat") Then
    frmMsgDialog.RetLabel = "-2"
    UBLog "ERROR: NO Group Code FILE"
    FntSize = frmMsgDialog.Label(3).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = "NO Group CODE FILE!"
    MsgText(3) = ""
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
    Exit Sub
  Else
    frmReportOpt.Show 1
    DeActivateControls Me
    If rptopt = 1 Then
      'do the graphics
      PrintCodeListing True
    ElseIf rptopt = 2 Then
      'do the text
      PrintCodeListing False
      ActivateControls Me
    Else
      ActivateControls Me
    End If
    Unload Me
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      ClearInUse PWcnt
    End If
  End If
End Sub
'This is to fix spreadsheet for various resolutions
Public Function Fixspread()
    Select Case screenW
      Case 1280
      If Screen.TwipsPerPixelX <> 12 Then
        coladj = 10
        vaSpread1.RowHeight(-1) = 19
      Else
        coladj = 6.1
        vaSpread1.RowHeight(-1) = 18.5
      End If
      Case 1152
      If Screen.TwipsPerPixelX <> 12 Then
        coladj = 7
        vaSpread1.RowHeight(-1) = 17
      Else
        coladj = 4.6
        vaSpread1.RowHeight(-1) = 15.5
      End If
      Case 1024
      If Screen.TwipsPerPixelX <> 12 Then
        coladj = 5
        vaSpread1.RowHeight(-1) = 18
      Else
        coladj = 3.1
      End If
      Case 800
        coladj = 2.9
        'vaSpread1.Font.Size = 8
        vaSpread1.RowHeight(-1) = 14
      Case Else
        'don't worry be happpy
    End Select
    'vaSpread1.ColWidth(-1) = vaSpread1.ColWidth(-1) + coladj
End Function
Private Sub cmdSave_Click()
  Dim goahead As Integer
  goahead = 0
  If Len(txtGroupNum) <> 0 And Len(txtGroupDesc) <> 0 Then
    If MsgBox("Do You wish to abandon the new code?", vbOKCancel, "Abandon New") = vbOK Then
      goahead = 1
    End If
  Else
    goahead = 1
  End If
  If goahead > 0 Then
    If MsgBox("Are You Sure You Are Ready To Save The Group Codes?", vbOKCancel, "Save Codes") = vbOK Then
      SaveCodes
      MsgBox "Group Codes Saved.", vbOKOnly, "Procedure Complete"
      frmUBSetupMenu.Show
      Unload Me
    End If
  End If
End Sub


Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
'Private Sub cmdDelete_Click()
''Only allow delete if not in use by accounts or saved previously
'  If RecordNum > 0 Then
'    If FindAcct(txtFundNum, GLFundLen) > 0 Then
'      MsgBox "This Fund May Not Be Deleted", vbOKOnly, "Deletion Denied"
'      Exit Sub
'    Else
'      If MsgBox("Are You Sure You Wish to Delete This Fund, OK to Delete, Cancel to Abort Deletion.", vbOKCancel, "Delete Fund") = vbOK Then
'        DeleteFund
'        txtFundNum = ""
'        txtTitle = ""
'        txtFundNum.SetFocus
'        lblEditFund.Visible = False
'      Else
'        Exit Sub
'      End If
'    End If
'  Else
'    MsgBox "This Fund Has Not Been Saved and Does Not Need To Be Deleted", vbOKOnly, "Deletion Denied"
'  End If
'End Sub
'Private Sub cmdSave_Click()
''Do not save if blank fields
'  If txtFundNum = "" Or txtTitle = "" Then
'    MsgBox "A Blank Field May Not Be Saved.", vbOKOnly
'  Else
'    Call SaveFund
'    lblNewFund.Visible = False
'  End If
''set focus back to fund number after save or after message
'  txtFundNum.SetFocus
'End Sub
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
    Case vbKeyF2:
      SendKeys "%N"
      KeyCode = 0
    Case vbKeyF3:
      SendKeys "%E"
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%P"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      KeyCode = 0
    Case vbKeyF9:
      SendKeys "%A"
      KeyCode = 0

    Case Else:
  End Select
End Sub
Private Function CodeSearch()
  Dim FoundCode As Boolean
  FoundCode = False 'assume we can't find it
  RecordNum = FindCode(txtGroupNum)
    If RecordNum > 0 Then
      MsgBox "Duplicate NOT Allowed, Try a different code or use original", vbOKOnly, "Invalid Entry"
      txtGroupNum = ""
      txtGroupNum.SetFocus
    Else
      'oh go on
  End If
'  FundSearch = FoundFund
End Function

Private Sub Form_Load()
  Dim cnt As Integer
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  'StatusBar1.Panels.Item(1).Text = GLUserName
  Me.HelpContextID = hlpGroupCode
  Fixspread
  CodeFillForm
  If vaSpread1.DataRowCnt > 0 Then
    lblNew.Visible = False
    lblEdit.Visible = True
    txtGroupNum.Enabled = False
    txtGroupDesc.Enabled = False
    vaSpread1.Lock = False
    cmdAdd2List.Enabled = False
    'vaSpread1.SetFocus
  Else
    lblNew.Visible = True
    lblEdit.Visible = False
    txtGroupNum.Enabled = True
    txtGroupDesc.Enabled = True
    vaSpread1.Lock = True
    cmdAdd2List.Enabled = True
    'txtGroupNum.SetFocus
  End If
End Sub
Public Function FindCode(CodeNum$)
  Dim MaxNum As Integer
  Dim GroupCde As GroupCodeRecType
  Dim GrpCodeRecLen As Integer, ghandle As Integer, cnt As Integer
  Dim NumofGrps As Integer
  Dim Match As Boolean, LookFor As String
  GrpCodeRecLen = Len(GroupCde)
  CodeNum$ = LTrim$(CodeNum$)
  ghandle = FreeFile
  Open UBPath$ + "UBGrpCde.DAT" For Random Shared As ghandle Len = GrpCodeRecLen
  NumofGrps = LOF(ghandle) \ GrpCodeRecLen
  For cnt = 1 To NumofGrps
    Get #ghandle, cnt, GroupCde
    LookFor$ = Trim$(GroupCde.GroupCode)
    If CodeNum$ = LookFor$ Then
      'If GroupCde.Deleted = 0 Then
        Match = True
        Close #ghandle
        Exit For
      'End If
    End If
  Next
  If Match Then
    FindCode = cnt
  Else
    FindCode = 0
    Close #ghandle
  End If
End Function

'  Set Over = New clsTextBoxOverRider
'  Over.OverRide Me
'  Set Temp_Class = New Resize_Class
'  Temp_Class.InitResizeClass Me
'  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
'  StatusBar1.Panels.Item(1).Text = GLUserName
'  Me.HelpContextID = hlpAddChangeDelete
'End Sub
Private Sub cmdExit_Click()
  Dim goahead As Integer
'  goahead = 0
'  If Len(txtGroupNum) <> 0 And Len(txtGroupDesc) <> 0 Then
'      goahead = 1
'    End If
'  Else
'    goahead = 1
'  End If
  Unload Me
End Sub
Private Sub cmdAdd2List_Click()
  Dim MaxNum As Integer
  Dim GroupCde As GroupCodeRecType
  Dim GrpCodeRecLen As Integer, ghandle As Integer, cnt As Integer
  Dim NumofGrps As Integer
  GrpCodeRecLen = Len(GroupCde)
  If Len(txtGroupNum) > 0 And Len(txtGroupDesc) > 0 Then
    ghandle = FreeFile
    Open UBPath$ + "UBGrpCde.DAT" For Random Shared As ghandle Len = GrpCodeRecLen
    NumofGrps = LOF(ghandle) \ GrpCodeRecLen
    GroupCde.Deleted = 0
    GroupCde.GroupCode = QPTrim$(txtGroupNum)
    GroupCde.GroupCodeName = QPTrim$(txtGroupDesc)
    GroupCde.xtrastuff = ""
    Put #ghandle, NumofGrps + 1, GroupCde
    Close
    vaSpread1.ClearRange 1, 1, 4, vaSpread1.DataRowCnt, True
    CodeFillForm
  Else
    MsgBox "Not enough information for Group code, Please complete code and desc fields.", vbOKOnly, "Invalid Entry"
  End If
End Sub

Public Function CodeFillForm()
  Dim MaxNum As Integer
  Dim GroupCde As GroupCodeRecType
  Dim GrpCodeRecLen As Integer, ghandle As Integer, cnt As Integer
  Dim NumofGrps As Integer
  GrpCodeRecLen = Len(GroupCde)
  
  ghandle = FreeFile
  Open UBPath$ + "UBGrpCde.DAT" For Random Shared As ghandle Len = GrpCodeRecLen
  NumofGrps = LOF(ghandle) \ GrpCodeRecLen
  For cnt = 1 To NumofGrps
    Get #ghandle, cnt, GroupCde
        vaSpread1.Row = cnt 'vaSpread1.DataRowCnt + 1
        vaSpread1.col = 1
        vaSpread1.Text = cnt
        vaSpread1.col = 2
        vaSpread1.Text = QPTrim(GroupCde.GroupCode)
        vaSpread1.col = 3
        vaSpread1.Text = QPTrim(GroupCde.GroupCodeName)
        vaSpread1.col = 4
      If GroupCde.Deleted = 0 Then
        vaSpread1.Text = False
      Else
        vaSpread1.Text = True
      End If
  Next
  Close #ghandle
  vaSpread1.MaxRows = vaSpread1.DataRowCnt

End Function
Private Sub SaveCodes()
  Dim MaxNum As Integer, Rec As Integer
  Dim GroupCde As GroupCodeRecType
  Dim GrpCodeRecLen As Integer, ghandle As Integer, cnt As Integer
  Dim NumofGrps As Integer
  GrpCodeRecLen = Len(GroupCde)
  ghandle = FreeFile
    Open UBPath$ + "UBGrpCde.DAT" For Random Shared As ghandle Len = GrpCodeRecLen
    NumofGrps = LOF(ghandle) \ GrpCodeRecLen
  For cnt = 1 To vaSpread1.DataRowCnt
    vaSpread1.Row = cnt
    vaSpread1.col = 1
    If vaSpread1.Text <> "" Then
      Rec = vaSpread1.Text
      Get #ghandle, cnt, GroupCde

      vaSpread1.col = 2
      GroupCde.GroupCode = QPTrim$(vaSpread1.Text)
      vaSpread1.col = 3
      GroupCde.GroupCodeName = QPTrim$(vaSpread1.Text)
      vaSpread1.col = 4
      If vaSpread1.Text = False Then
        GroupCde.Deleted = 0
      Else
        GroupCde.Deleted = 1
      End If
    
      GroupCde.xtrastuff = ""
   
      Put #ghandle, Rec, GroupCde
    End If
  Next
  Close
'  'mainlog updates log file if used save command or saved on exit
  UBLog "Group codes Saved."
'  Close AcctFileNum
End Sub
Private Sub ChkforChange()

  Dim MaxNum As Integer, Rec As Integer
  Dim GroupCde As GroupCodeRecType
  Dim GrpCodeRecLen As Integer, ghandle As Integer, cnt As Integer
  Dim NumofGrps As Integer
  GrpCodeRecLen = Len(GroupCde)
  ghandle = FreeFile
    Open UBPath$ + "UBGrpCde.DAT" For Random Shared As ghandle Len = GrpCodeRecLen
    NumofGrps = LOF(ghandle) \ GrpCodeRecLen
  For cnt = 1 To vaSpread1.DataRowCnt
    vaSpread1.Row = cnt
    vaSpread1.col = 1
    If vaSpread1.Text <> "" Then
      Rec = vaSpread1.Text
      Get #ghandle, cnt, GroupCde

      vaSpread1.col = 2
      GroupCde.GroupCode = QPTrim$(vaSpread1.Text)
      vaSpread1.col = 3
      GroupCde.GroupCodeName = QPTrim$(vaSpread1.Text)
      vaSpread1.col = 4
      If vaSpread1.Text = False Then
        GroupCde.Deleted = 0
      Else
        GroupCde.Deleted = 1
      End If
    
      GroupCde.xtrastuff = ""
   
      Put #ghandle, Rec, GroupCde
    End If
  Next
  Close

End Sub

Private Sub txtGroupNum_LostFocus()
  CodeSearch
End Sub
Private Sub vaSpread1_Change(ByVal col As Long, ByVal Row As Long)
  Dim FoundCode As Boolean, txtcode As String
  FoundCode = False 'assume we can't find it
  If col = 2 Then
  vaSpread1.col = col
  vaSpread1.Row = Row
  txtcode = vaSpread1.Text
  RecordNum = FindCode(txtcode)
    If RecordNum > 0 Then
      MsgBox "Duplicate NOT Allowed, Try a different code or use original", vbOKOnly, "Invalid Entry"
      CodeFillForm
    Else
      'oh go on
  End If
 End If
End Sub
Private Sub PrintCodeListing(graphicflag As Boolean)
  Dim GroupCde As GroupCodeRecType
  Dim GrpCodeRecLen As Integer, ghandle As Integer, cnt As Integer
  Dim NumofGrps As Integer, RPTFile As Integer
  Dim ReportFile As String, ToPrint As String
  Dim Dash80 As String * 78
  GrpCodeRecLen = Len(GroupCde)
  ghandle = FreeFile
    Open UBPath$ + "UBGrpCde.DAT" For Random Shared As ghandle Len = GrpCodeRecLen
    NumofGrps = LOF(ghandle) \ GrpCodeRecLen
  
  Dash80$ = String$(78, "-")

  FrmShowPctComp.Label1 = "Creating Group Codes Report."
  FrmShowPctComp.Show , Me

  ReportFile$ = UBPath + "CodeLIST.RPT"
  
  RPTFile = FreeFile
  Open ReportFile$ For Output As RPTFile
    GoSub PrintCodeHeader
  For cnt = 1 To NumofGrps
    Get #ghandle, cnt, GroupCde
    FrmShowPctComp.ShowPctComp cnt, NumofGrps
      If GroupCde.Deleted <> 0 Then
        Print #RPTFile, Tab(10); "Inactive";
      End If

      Print #RPTFile, Tab(30); QPTrim$(GroupCde.GroupCode);
      Print #RPTFile, Tab(50); QPTrim$(GroupCde.GroupCodeName)
    Next
    Print #RPTFile, Dash80$
    Print #RPTFile, Chr$(12)
    Close
  DoEvents
  If graphicflag Then
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmGroupCodeEntryEdit
    ARptLineRpt.GetName ReportFile$
    ARptLineRpt.startrpt
  Else
    ViewPrint ReportFile$, "Group Code List Report"
    KillFile "CodeLIST.RPT"
  End If
  GoTo ExitListing

PrintCodeHeader:
  PageNo = PageNo + 1
  Print #RPTFile, " "
  Print #RPTFile, " "
  Print #RPTFile, "Utility Billing Group Code Listing."
  Print #RPTFile, TOWNNAME$; Tab(70); "Page:"; PageNo
  Print #RPTFile, "Report Date: "; Date$
  Print #RPTFile, Dash80$
  Print #RPTFile, Tab(30); "Group Code";
  Print #RPTFile, Tab(50); "Description"
Return
ExitListing:

End Sub




