VERSION 5.00
Begin VB.Form frm10EarningsCodeMaint 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Additional Earning Codes"
   ClientHeight    =   9996
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   13728
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9996
   ScaleWidth      =   13728
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSave 
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
      Height          =   528
      Left            =   5904
      TabIndex        =   20
      Top             =   6912
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
      Height          =   528
      Left            =   9312
      TabIndex        =   23
      Top             =   6912
      Width           =   1332
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "F3 &Delete"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   528
      Left            =   7632
      TabIndex        =   22
      Top             =   6912
      Width           =   1332
   End
   Begin VB.ComboBox comboRET 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Index           =   3
      Left            =   10080
      TabIndex        =   19
      ToolTipText     =   "Are earnings exempt from Retirement?."
      Top             =   5664
      Width           =   732
   End
   Begin VB.ComboBox comboRET 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Index           =   2
      Left            =   10080
      TabIndex        =   12
      ToolTipText     =   "Are earnings exempt from Retirement?."
      Top             =   4704
      Width           =   732
   End
   Begin VB.ComboBox comboRET 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Index           =   1
      Left            =   10080
      TabIndex        =   6
      ToolTipText     =   "Are earnings exempt from Retirement?."
      Top             =   3744
      Width           =   732
   End
   Begin VB.ComboBox comboMED 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Index           =   3
      Left            =   9000
      TabIndex        =   18
      ToolTipText     =   "Are earnings Medicare Tax Exempt?"
      Top             =   5664
      Width           =   732
   End
   Begin VB.ComboBox comboMED 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Index           =   2
      Left            =   9000
      TabIndex        =   11
      ToolTipText     =   "Are earnings Medicare Tax Exempt?"
      Top             =   4704
      Width           =   732
   End
   Begin VB.ComboBox comboMED 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Index           =   1
      Left            =   9000
      TabIndex        =   5
      ToolTipText     =   "Are earnings Medicare Tax Exempt?"
      Top             =   3744
      Width           =   732
   End
   Begin VB.ComboBox comboSOC 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Index           =   3
      Left            =   7800
      TabIndex        =   17
      ToolTipText     =   "Are earnings Social Security Tax Exempt?"
      Top             =   5664
      Width           =   732
   End
   Begin VB.ComboBox comboSOC 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Index           =   2
      Left            =   7800
      TabIndex        =   10
      ToolTipText     =   "Are earnings Social Security Tax Exempt?"
      Top             =   4704
      Width           =   732
   End
   Begin VB.ComboBox comboSOC 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Index           =   1
      Left            =   7800
      TabIndex        =   4
      ToolTipText     =   "Are earnings Social Security Tax Exempt?"
      Top             =   3744
      Width           =   732
   End
   Begin VB.ComboBox comboSWT 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Index           =   3
      Left            =   6600
      TabIndex        =   16
      ToolTipText     =   "Are earnings State Tax Exempt?"
      Top             =   5664
      Width           =   732
   End
   Begin VB.ComboBox comboSWT 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Index           =   2
      Left            =   6600
      TabIndex        =   9
      ToolTipText     =   "Are earnings State Tax Exempt?"
      Top             =   4704
      Width           =   732
   End
   Begin VB.ComboBox comboSWT 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Index           =   1
      Left            =   6600
      TabIndex        =   3
      ToolTipText     =   "Are earnings State Tax Exempt?"
      Top             =   3744
      Width           =   732
   End
   Begin VB.ComboBox comboFWT 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Index           =   3
      Left            =   5400
      TabIndex        =   15
      ToolTipText     =   "Are earnings Federal Tax Exempt?"
      Top             =   5664
      Width           =   732
   End
   Begin VB.ComboBox comboFWT 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Index           =   2
      Left            =   5400
      TabIndex        =   8
      ToolTipText     =   "Are earnings Federal Tax Exempt?"
      Top             =   4704
      Width           =   732
   End
   Begin VB.ComboBox comboFWT 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Index           =   1
      Left            =   5400
      TabIndex        =   2
      ToolTipText     =   "Are earnings Federal Tax Exempt?"
      Top             =   3744
      Width           =   732
   End
   Begin VB.TextBox txtDescrp1 
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
      Index           =   3
      Left            =   1680
      TabIndex        =   14
      ToolTipText     =   "Enter the Earnings Code Description."
      Top             =   5664
      Width           =   3012
   End
   Begin VB.TextBox txtDescrp1 
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
      Index           =   2
      Left            =   1680
      TabIndex        =   7
      ToolTipText     =   "Enter the Earnings Code Description."
      Top             =   4704
      Width           =   3012
   End
   Begin VB.TextBox txtDescrp1 
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
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      ToolTipText     =   "Enter the Earnings Code Description."
      Top             =   3744
      Width           =   3012
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "RET"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   9960
      TabIndex        =   28
      Top             =   3144
      Width           =   972
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "MED"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   8796
      TabIndex        =   27
      Top             =   3144
      Width           =   972
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "SOC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   7620
      TabIndex        =   26
      Top             =   3144
      Width           =   972
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "SWT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   6456
      TabIndex        =   25
      Top             =   3144
      Width           =   972
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "FWT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   5280
      TabIndex        =   24
      Top             =   3144
      Width           =   972
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "Withholding on Earnings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   5760
      TabIndex        =   21
      Top             =   2424
      Width           =   4692
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   2016
      TabIndex        =   13
      Top             =   2424
      Width           =   2292
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   4284
      Left            =   912
      Top             =   2136
      Width           =   10572
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   1092
      Index           =   1
      Left            =   1776
      Top             =   720
      Width           =   8652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Additional Earning Codes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3000
      TabIndex        =   0
      Top             =   1080
      Width           =   6012
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1776
      Top             =   600
      Width           =   8652
   End
End
Attribute VB_Name = "frm10EarningsCodeMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim changeFlag As Integer

Private Sub cmdSave_Click()
   changeFlag = 0
   Dim ErnCodeHandle As Integer, x As Integer
   Dim ErnCodeFileRec(1 To 3) As ErnCodeRecType
   
   For x = 1 To 3
     ErnCodeFileRec(x).ERNCODE1 = QPTrim$(txtDescrp1(x).Text)
     ErnCodeFileRec(x).ERNFWT1 = QPTrim$(comboFWT(x).Text)
     ErnCodeFileRec(x).ERNSWT1 = QPTrim$(comboSWT(x).Text)
     ErnCodeFileRec(x).ERNSOC1 = QPTrim$(comboSOC(x).Text)
     ErnCodeFileRec(x).ERNMED1 = QPTrim$(comboMED(x).Text)
     ErnCodeFileRec(x).ERNRET1 = QPTrim$(comboRET(x).Text)
   Next
      
   OpenErnCodeFile ErnCodeHandle
   For x = 1 To 3
     Put ErnCodeHandle, x, ErnCodeFileRec(x)
   Next
   Close ErnCodeHandle

   MsgBox "Your Information has been saved.", vbOKOnly
   frmControlFileMaint.Show
   Unload frmRetireFileMaint
NoRetFileYet:

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  LoadUnitFile
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub LoadUnitFile()
'   On Error Resume Next
   changeFlag = 0
   Dim ErnCodeHandle As Integer, x As Integer
   Dim ErnCodeFileRec(1 To 3) As ErnCodeRecType
   Dim FileSize As Long
   OpenErnCodeFile ErnCodeHandle
   FileSize = LOF(ErnCodeHandle)
   If FileSize = 0 Then
      For x = 1 To 3
        comboFWT(x).AddItem ("Y")
        comboFWT(x).AddItem ("N")
        comboSWT(x).AddItem ("Y")
        comboSWT(x).AddItem ("N")
        comboSOC(x).AddItem ("Y")
        comboSOC(x).AddItem ("N")
        comboMED(x).AddItem ("Y")
        comboMED(x).AddItem ("N")
        comboRET(x).AddItem ("Y")
        comboRET(x).AddItem ("N")
      Next
      'file is zero bytes
      GoTo NoUnitFileYet
   Else
     For x = 1 To 3
        Get ErnCodeHandle, x, ErnCodeFileRec(x)
     Next
   End If
   Close ErnCodeHandle
   'load form info
   For x = 1 To 3
     txtDescrp1(x).Text = ErnCodeFileRec(x).ERNCODE1
     comboFWT(x).AddItem ("Y")
     comboFWT(x).AddItem ("N")
     comboSWT(x).AddItem ("Y")
     comboSWT(x).AddItem ("N")
     comboSOC(x).AddItem ("Y")
     comboSOC(x).AddItem ("N")
     comboMED(x).AddItem ("Y")
     comboMED(x).AddItem ("N")
     comboRET(x).AddItem ("Y")
     comboRET(x).AddItem ("N")
     comboFWT(x).Text = ErnCodeFileRec(x).ERNFWT1
     comboSWT(x).Text = ErnCodeFileRec(x).ERNSWT1
     comboSOC(x).Text = ErnCodeFileRec(x).ERNSOC1
     comboMED(x).Text = ErnCodeFileRec(x).ERNMED1
     comboRET(x).Text = ErnCodeFileRec(x).ERNRET1
   Next

NoUnitFileYet:
End Sub
Private Sub cmdExit_Click()

   changeFlag = 0
   Dim DoWhatFlag As SaveChangeOptions1, x As Integer
'   Dim save As Integer, review As Integer, abandon As Integer
   Dim ErnCodeHandle As Integer
   Dim ErnCodeFileRec(1 To 3) As ErnCodeRecType
   OpenErnCodeFile ErnCodeHandle
   For x = 1 To 3
      Get ErnCodeHandle, x, ErnCodeFileRec(x)
   Next
   Close ErnCodeHandle
   For x = 1 To 3
      If QPTrim$(ErnCodeFileRec(x).ERNCODE1) <> QPTrim$(txtDescrp1(x).Text) Then
      'if it has changed then set the changeFlag to 1 and reset focus to
      'where the change was made
        changeFlag = 1
        txtDescrp1(x).SetFocus
        Exit For
      End If
      If QPTrim$(ErnCodeFileRec(x).ERNFWT1) <> QPTrim$(comboFWT(x).Text) Then
        changeFlag = 1
        comboFWT(x).SetFocus
        Exit For
      End If
      If QPTrim$(ErnCodeFileRec(x).ERNSWT1) <> QPTrim$(comboSWT(x).Text) Then
        changeFlag = 1
        comboSWT(x).SetFocus
        Exit For
      End If
      If QPTrim$(ErnCodeFileRec(x).ERNSOC1) <> QPTrim$(comboSOC(x).Text) Then
        changeFlag = 1
        comboSOC(x).SetFocus
        Exit For
      End If
      If QPTrim$(ErnCodeFileRec(x).ERNMED1) <> QPTrim$(comboMED(x).Text) Then
        changeFlag = 1
        comboMED(x).SetFocus
        Exit For
      End If
      If QPTrim$(ErnCodeFileRec(x).ERNRET1) <> QPTrim$(comboRET(x).Text) Then
        changeFlag = 1
        comboRET(x).SetFocus
        Exit For
      End If
   Next x
   'check each textbox to see if a change has been made
   
   'if no changes were made then move back to control menu
   If changeFlag = 0 Then 'no changes detected
      frmControlFileMaint.Show
      Unload frmRetireFileMaint
      GoTo endClick
   'if a change was made then bring up a warning window that forces
   'the user to decide whether to save, review or abandon changes
   Else
      DoWhatFlag = PromptSaveChanges(Me)
      Select Case DoWhatFlag
      Case SaveChangeOptions1.scoSaveChanges 'save changes
        Call cmdSave_Click
      Case SaveChangeOptions1.scoReviewChanges 'review is just bringing back the current form
      Case SaveChangeOptions1.scoAbandonChanges 'abandon
        frmControlFileMaint.Show
        Unload frmEarningsCodeMaint
      Case Else:
        'Do nothing because we don't know about any options except
        'save, review or abandon...used as a placeholder for adding
        'other options at a later date
      End Select
      
   End If

endClick:
End Sub
