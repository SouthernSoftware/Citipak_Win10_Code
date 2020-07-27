VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmACHDraftInfo 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ACH Draft Information"
   ClientHeight    =   8565
   ClientLeft      =   30
   ClientTop       =   585
   ClientWidth     =   11655
   Icon            =   "frmACHDraftInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8540.847
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtImmDestNum 
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
      Left            =   6204
      MaxLength       =   9
      TabIndex        =   1
      ToolTipText     =   "Enter the Bank's  Transit/Routing ABA Number here of where you will be sending the ACH transaction to."
      Top             =   2964
      Width           =   3852
   End
   Begin VB.TextBox txtImmOriginNum 
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
      Left            =   6204
      MaxLength       =   9
      TabIndex        =   2
      ToolTipText     =   "Enter the distributing Bank's ABA Number"
      Top             =   3399
      Width           =   3852
   End
   Begin VB.TextBox txtDestBankNum 
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
      Left            =   6204
      MaxLength       =   23
      TabIndex        =   3
      ToolTipText     =   "Enter the Destinations Bank Name Here."
      Top             =   3834
      Width           =   3852
   End
   Begin VB.TextBox txtOriBankName 
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
      Left            =   6204
      MaxLength       =   23
      TabIndex        =   4
      ToolTipText     =   "Enter the Originating Banks Name Here."
      Top             =   4269
      Width           =   3852
   End
   Begin VB.TextBox txtFedIDPrefix 
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
      Left            =   6204
      MaxLength       =   1
      TabIndex        =   5
      ToolTipText     =   "No help for this field."
      Top             =   4704
      Width           =   3852
   End
   Begin VB.TextBox txtComFedID 
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
      Left            =   6204
      MaxLength       =   9
      TabIndex        =   6
      ToolTipText     =   "Enter Your Federal ID Number Without the '-'."
      Top             =   5136
      Width           =   3852
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   645
      Left            =   8880
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Press F10 to commit the above data to memory."
      Top             =   6960
      Width           =   1470
      _Version        =   131072
      _ExtentX        =   2593
      _ExtentY        =   1138
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmACHDraftInfo.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   645
      Left            =   8880
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "Press ESC to exit this screen and return to the 'ACH Bank Draft Maintenance' menu."
      Top             =   6120
      Width           =   1470
      _Version        =   131072
      _ExtentX        =   2593
      _ExtentY        =   1138
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmACHDraftInfo.frx":0AA6
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   9357.591
      X2              =   10677.25
      Y1              =   2871.878
      Y2              =   2871.878
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   1439.629
      X2              =   2999.228
      Y1              =   2871.878
      Y2              =   2871.878
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   10677.25
      X2              =   10677.25
      Y1              =   5614.123
      Y2              =   2871.878
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   1439.629
      X2              =   1439.629
      Y1              =   5614.123
      Y2              =   2861.906
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   1439.629
      X2              =   10677.25
      Y1              =   5609.137
      Y2              =   5609.137
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "401K CENTER"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6240
      TabIndex        =   24
      Top             =   7305
      UseMnemonic     =   0   'False
      Width           =   1905
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "401K CENTER"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6240
      TabIndex        =   23
      Top             =   6930
      UseMnemonic     =   0   'False
      Width           =   1905
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "053101121"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6255
      TabIndex        =   22
      Top             =   6540
      Width           =   1905
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "053101121"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   6255
      TabIndex        =   21
      Top             =   6150
      Width           =   1905
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Originating Bank Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   345
      Left            =   1170
      TabIndex        =   20
      Top             =   7305
      Width           =   4785
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Destination Bank Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   345
      Left            =   1170
      TabIndex        =   19
      Top             =   6930
      Width           =   4785
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Immediate Origin Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   345
      Left            =   1170
      TabIndex        =   18
      Top             =   6540
      Width           =   4785
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Immediate Destination Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   390
      Left            =   1170
      TabIndex        =   17
      Top             =   6105
      Width           =   4785
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Example:"
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
      Height          =   375
      Left            =   4425
      TabIndex        =   16
      Top             =   5775
      Width           =   1305
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Company Federal ID Number:"
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
      Height          =   348
      Left            =   1740
      TabIndex        =   15
      Top             =   5196
      Width           =   4185
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Federal ID Prefix Number:"
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
      Height          =   300
      Left            =   1740
      TabIndex        =   14
      Top             =   4758
      Width           =   4185
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Originating Bank Name:"
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
      Height          =   348
      Left            =   1740
      TabIndex        =   13
      Top             =   4278
      Width           =   4185
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Destination Bank Name:"
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
      Height          =   348
      Left            =   1740
      TabIndex        =   12
      Top             =   3846
      Width           =   4185
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Immediate Origin Number:"
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
      Height          =   345
      Left            =   1740
      TabIndex        =   11
      Top             =   3420
      Width           =   4185
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Immediate Destination Number:"
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
      Height          =   330
      Left            =   1740
      TabIndex        =   10
      Top             =   3000
      Width           =   4185
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "ALL FIELDS ARE REQUIRED"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   4665
      TabIndex        =   9
      Top             =   2460
      Width           =   4425
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "NOTE!:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Left            =   3225
      TabIndex        =   8
      Top             =   2450
      Width           =   1425
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   465
      Left            =   2895
      Top             =   2427
      Width           =   6465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "Company Information Required for  ACH Draft Transmission"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   492
      Left            =   1740
      TabIndex        =   7
      Top             =   1686
      Width           =   8748
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   588
      Left            =   1452
      Top             =   1614
      Width           =   9228
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   1092
      Index           =   1
      Left            =   1716
      Top             =   294
      Width           =   8652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ACH Draft Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2940
      TabIndex        =   0
      Top             =   654
      Width           =   6012
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1716
      Top             =   174
      Width           =   8652
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
Attribute VB_Name = "frmACHDraftInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim changeFlag As Boolean

Private Sub cmdExit_Click()
   changeFlag = 0
   Dim DoWhatFlag As SaveChangeOptions1
   Dim save As Integer, review As Integer, abandon As Integer
   Dim PRDraftFileHandle As Integer
   Dim PRDraftFileRec As DraftInfoFileName
   
   OpenPRDraftFile PRDraftFileHandle
   Get PRDraftFileHandle, 1, PRDraftFileRec
   Close PRDraftFileHandle
   'check each textbox to see if a change has been made
   If QPTrim$(PRDraftFileRec.BANKDEST) <> QPTrim$(txtImmDestNum.Text) Then
   'if it has then set the changeFlag to 1 and reset focus to
   'where the change was made
      changeFlag = True
      txtImmDestNum.SetFocus
   End If
   
   If QPTrim$(PRDraftFileRec.BANKORIG) <> QPTrim$(txtImmOriginNum.Text) Then
   'if it has then set the changeFlag to 1 and reset focus to
   'where the change was made
      changeFlag = True
      txtImmOriginNum.SetFocus
   End If
   
   If QPTrim$(PRDraftFileRec.BankName) <> QPTrim$(txtDestBankNum.Text) Then
   'if it has then set the changeFlag to 1 and reset focus to
   'where the change was made
      changeFlag = True
      txtDestBankNum.SetFocus
   End If
   
   If QPTrim$(PRDraftFileRec.BANKLOC) <> QPTrim$(txtOriBankName.Text) Then
   'if it has then set the changeFlag to 1 and reset focus to
   'where the change was made
      changeFlag = True
      txtOriBankName.SetFocus
   End If
   
    If QPTrim$(PRDraftFileRec.FEDPREFX) <> QPTrim$(txtFedIDPrefix.Text) Then
   'if it has then set the changeFlag to 1 and reset focus to
   'where the change was made
      changeFlag = True
      txtFedIDPrefix.SetFocus
   End If
   
    If QPTrim$(PRDraftFileRec.FEDID) <> QPTrim$(txtComFedID.Text) Then
   'if it has then set the changeFlag to 1 and reset focus to
   'where the change was made
      changeFlag = True
      txtComFedID.SetFocus
   End If
   
   'if no changes were made then move back to control menu
   If changeFlag = False Then 'no changes detected
      If Not Exist("quickmaintdd.dat") Then
        frmACHControlMenu.Show
        DoEvents
        Unload frmACHDraftInfo
      Else
        frmEmpQuickMaintDirDep.Show
        DoEvents
        Unload Me
        KillFile "quickmaintdd.dat"
      End If
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
        If Not Exist("quickmaintdd.dat") Then
          frmACHControlMenu.Show
          DoEvents
          Unload frmACHDraftInfo
        Else
          frmEmpQuickMaintDirDep.Show
          DoEvents
          Unload Me
        End If
      Case Else:
        'Do nothing because we don't know about any options except
        'save, review or abandon...used as a placeholder for adding
        'other options at a later date
      End Select
      
   End If

endClick:

End Sub

Private Sub cmdSave_Click()
   Dim tempImmDestNum As String, tempOriNum As String, tempDestBankName As String
   Dim tempOriBankName As String, tempFedIDPre As String, tempComFedID As String
   Dim PRDraftFileHandle As Integer
   Dim PRDraftFileRec As DraftInfoFileName
   
   tempImmDestNum = QPTrim$(txtImmDestNum.Text)
   If tempImmDestNum = "" Then
      MsgBox "Please enter an Immediate Destination Number"
      txtImmDestNum.SetFocus
      GoTo BadUnitData
   End If
   
   tempOriNum = QPTrim$(txtImmOriginNum.Text)
   If tempOriNum = "" Then
      MsgBox "Please enter an Immediate Origin Number"
      txtImmOriginNum.SetFocus
      GoTo BadUnitData
   End If
   
   tempDestBankName = QPTrim$(txtDestBankNum.Text)
   If tempDestBankName = "" Then
      MsgBox "Please enter a Destination Bank Number"
      txtDestBankNum.SetFocus
      GoTo BadUnitData
   End If
   
   tempOriBankName = QPTrim$(txtOriBankName.Text)
   If tempOriBankName = "" Then
      MsgBox "Please enter an Originating Bank Number"
      txtOriBankName.SetFocus
      GoTo BadUnitData
   End If
   
  tempFedIDPre = QPTrim$(txtFedIDPrefix.Text)
   If tempFedIDPre = "" Then
      MsgBox "Please enter a Federal ID Prefix Number"
      txtFedIDPrefix.SetFocus
      GoTo BadUnitData
   End If
   
   tempComFedID = QPTrim$(txtComFedID.Text)
   If tempComFedID = "" Then
      MsgBox "Please enter a Company Federal ID Number"
      txtComFedID.SetFocus
      GoTo BadUnitData
   End If
   
   OpenPRDraftFile PRDraftFileHandle
   
   PRDraftFileRec.BANKDEST = tempImmDestNum
   PRDraftFileRec.BANKORIG = tempOriNum
   PRDraftFileRec.BankName = tempDestBankName
   PRDraftFileRec.BANKLOC = tempOriBankName
   PRDraftFileRec.FEDPREFX = tempFedIDPre
   PRDraftFileRec.FEDID = tempComFedID
   
   Put PRDraftFileHandle, 1, PRDraftFileRec
   Close PRDraftFileHandle

   MsgBox "Your Information has been saved.", vbOKOnly
   If Not Exist("quickmaintdd.dat") Then
     frmACHControlMenu.Show
   Else
     frmEmpQuickMaintDirDep.Show
   End If
   DoEvents
   Unload frmACHDraftInfo
   MainLog ("ACH Draft data saved.")
BadUnitData:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      Call cmdSave_Click
      KeyCode = 0
  End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  LoadUnitFile
  Me.HelpContextID = hlpACHBankDraft
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub LoadUnitFile()
'   On Error Resume Next
   changeFlag = 0
   Dim PRDraftFileHandle As Integer
   Dim PRDraftFileRec As DraftInfoFileName
   Dim FileSize As Long
   
   OpenPRDraftFile PRDraftFileHandle
   FileSize = LOF(PRDraftFileHandle)
   If FileSize = 0 Then
      'file is zero bytes
     Close
     GoTo NoPRDraftFileYet
   Else
     Get PRDraftFileHandle, 1, PRDraftFileRec
   End If
   Close PRDraftFileHandle
   'load form info
   txtImmDestNum.Text = PRDraftFileRec.BANKDEST
   txtImmOriginNum.Text = PRDraftFileRec.BANKORIG
   txtDestBankNum.Text = PRDraftFileRec.BankName
   txtOriBankName.Text = PRDraftFileRec.BANKLOC
   txtFedIDPrefix.Text = PRDraftFileRec.FEDPREFX
   txtComFedID.Text = PRDraftFileRec.FEDID

NoPRDraftFileYet:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmACHDraftInfo.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub mnuExit_Click()
  Call cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
  MainLog ("ACH Draft control screen printed.")
End Sub
