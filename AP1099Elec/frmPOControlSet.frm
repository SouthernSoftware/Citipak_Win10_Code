VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPOControlSet 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order Default Information"
   ClientHeight    =   8844
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmPOControlSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8844
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtAddinst3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   6939
      MaxLength       =   20
      TabIndex        =   15
      Top             =   6456
      Width           =   2844
   End
   Begin VB.TextBox txtAddinst2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   6939
      MaxLength       =   20
      TabIndex        =   14
      Top             =   5976
      Width           =   2844
   End
   Begin VB.TextBox txtAddinst1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   6939
      MaxLength       =   20
      TabIndex        =   13
      Top             =   5496
      Width           =   2844
   End
   Begin VB.TextBox txtTerms 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   3219
      MaxLength       =   20
      TabIndex        =   12
      Top             =   6360
      Width           =   2844
   End
   Begin VB.TextBox txtShipVia 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   3219
      MaxLength       =   20
      TabIndex        =   11
      Top             =   5700
      Width           =   2844
   End
   Begin VB.TextBox txtFOB 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   3219
      MaxLength       =   20
      TabIndex        =   10
      Top             =   5064
      Width           =   2844
   End
   Begin VB.TextBox txtShipto5 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   6915
      MaxLength       =   30
      TabIndex        =   9
      Top             =   4464
      Width           =   3780
   End
   Begin VB.TextBox txtShipto4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   6915
      MaxLength       =   30
      TabIndex        =   8
      Top             =   4008
      Width           =   3780
   End
   Begin VB.TextBox txtShipto3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   6915
      MaxLength       =   30
      TabIndex        =   7
      Top             =   3564
      Width           =   3780
   End
   Begin VB.TextBox txtShipto2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   6915
      MaxLength       =   30
      TabIndex        =   6
      Top             =   3108
      Width           =   3780
   End
   Begin VB.TextBox txtShipto1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   6915
      MaxLength       =   30
      TabIndex        =   5
      Top             =   2664
      Width           =   3780
   End
   Begin VB.TextBox txtHeader4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   1755
      MaxLength       =   35
      TabIndex        =   4
      Top             =   4080
      Width           =   4308
   End
   Begin VB.TextBox txtHeader3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   1755
      MaxLength       =   35
      TabIndex        =   3
      Top             =   3624
      Width           =   4308
   End
   Begin VB.TextBox txtHeader2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   1755
      MaxLength       =   35
      TabIndex        =   2
      Top             =   3168
      Width           =   4308
   End
   Begin VB.TextBox txtHeader1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   1755
      MaxLength       =   35
      TabIndex        =   1
      Top             =   2712
      Width           =   4308
   End
   Begin VB.TextBox txtPONumber 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   408
      Left            =   4647
      MaxLength       =   8
      TabIndex        =   0
      Top             =   1740
      Width           =   1236
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
      Left            =   9747
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7488
      Width           =   1332
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
      Left            =   8007
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7488
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   18
      Top             =   8484
      Width           =   12192
      _ExtentX        =   21505
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7133
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "2:07 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "3/14/2005"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      X1              =   2448
      X2              =   9720
      Y1              =   2208
      Y2              =   2208
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   5892
      Left            =   1116
      Top             =   1416
      Width           =   9948
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(Next PO Number to Use!)"
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
      Left            =   6036
      TabIndex        =   27
      Top             =   1728
      Width           =   3084
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Instructions:"
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
      Left            =   6528
      TabIndex        =   26
      Top             =   5136
      Width           =   2628
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Terms"
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
      Left            =   1920
      TabIndex        =   25
      Top             =   6384
      Width           =   924
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ship Via"
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
      Left            =   1800
      TabIndex        =   24
      Top             =   5736
      Width           =   1044
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "FOB Point"
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
      Left            =   1704
      TabIndex        =   23
      Top             =   5112
      Width           =   1140
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ship To Information:"
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
      Left            =   6600
      TabIndex        =   22
      Top             =   2328
      Width           =   2292
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PO Form Heading Information:"
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
      Left            =   1344
      TabIndex        =   21
      Top             =   2376
      Width           =   3396
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PO Number"
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
      Left            =   3012
      TabIndex        =   20
      Top             =   1752
      Width           =   1428
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order Control"
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
      Left            =   4236
      TabIndex        =   19
      Top             =   576
      Width           =   3708
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   852
      Left            =   2580
      Top             =   336
      Width           =   7020
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00D0D0D0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   996
      Left            =   2592
      Top             =   216
      Width           =   7020
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
Attribute VB_Name = "frmPOControlSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim POControl As POControlRecType
Private Temp_Class As Resize_Class
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

Private Sub cmdSave_Click()
  If txtPONumber <> "" Then
    SaveControl
    cmdExit_Click
  Else
    MsgBox "You Must Enter A Valid PO Number.", vbOKOnly, "Retry"
  End If
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Me.HelpContextID = hlpPOControl
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  Loadtoform
End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
 '   Me.SetFocus
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

Private Sub cmdExit_Click()
  frmPOProcessMenu.Show
  Unload frmPOControlSet
End Sub

Private Sub Loadtoform()
  Dim POFile As Integer, POFileLen As Integer, NumRecs As Integer
  ReDim POCont(1) As POControlRecType
  OpenPOFile POFile, NumRecs
  If LOF(POFile) > 0 Then
    Get POFile, 1, POCont(1)
    txtPONumber = QPTrim(POCont(1).PONumber)
    txtHeader1 = QPTrim(POCont(1).Header1)
    txtHeader2 = QPTrim(POCont(1).Header2)
    txtHeader3 = QPTrim(POCont(1).Header3)
    txtHeader4 = QPTrim(POCont(1).Header4)
    txtShipTo1 = QPTrim(POCont(1).Shipto1)
    txtShipTo2 = QPTrim(POCont(1).Shipto2)
    txtShipTo3 = QPTrim(POCont(1).Shipto3)
    txtShipTo4 = QPTrim(POCont(1).Shipto4)
    txtShipTo5 = QPTrim(POCont(1).Shipto5)
    txtFOB = QPTrim(POCont(1).FOB)
    txtShipVia = QPTrim(POCont(1).Shipvia)
    txtTerms = QPTrim(POCont(1).Terms)
    txtAddinst1 = QPTrim(POCont(1).Addinst1)
    txtAddinst2 = QPTrim(POCont(1).Addinst2)
    txtAddinst3 = QPTrim(POCont(1).Addinst3)
  End If
  Close POFile
End Sub

Private Sub SaveControl()
  Dim POFile As Integer, POFileLen As Integer, NumRecs As Integer
  ReDim POCont(1) As POControlRecType
  OpenPOFile POFile, NumRecs
  POCont(1).PONumber = QPTrim(txtPONumber)
  POCont(1).Header1 = QPTrim(txtHeader1)
  POCont(1).Header2 = QPTrim(txtHeader2)
  POCont(1).Header3 = QPTrim(txtHeader3)
  POCont(1).Header4 = QPTrim(txtHeader4)
  POCont(1).Shipto1 = QPTrim(txtShipTo1)
  POCont(1).Shipto2 = QPTrim(txtShipTo2)
  POCont(1).Shipto3 = QPTrim(txtShipTo3)
  POCont(1).Shipto4 = QPTrim(txtShipTo4)
  POCont(1).Shipto5 = QPTrim(txtShipTo5)
  POCont(1).FOB = QPTrim(txtFOB)
  POCont(1).Shipvia = QPTrim(txtShipVia)
  POCont(1).Terms = QPTrim(txtTerms)
  POCont(1).Addinst1 = QPTrim(txtAddinst1)
  POCont(1).Addinst2 = QPTrim(txtAddinst2)
  POCont(1).Addinst3 = QPTrim(txtAddinst3)
  POCont(1).Pading = ""
  Put POFile, 1, POCont(1)
  Close POFile
  
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
