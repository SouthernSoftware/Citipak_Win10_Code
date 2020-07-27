VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmWarningPostPayroll 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Posting Warning"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmWarningPostPayroll.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   3840
      Left            =   2424
      TabIndex        =   0
      Top             =   2592
      Width           =   6816
      _Version        =   196609
      _ExtentX        =   12023
      _ExtentY        =   6779
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.26
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   192
      Caption         =   ""
      FrameColor      =   12632256
      FrameThreeDHighlightColor=   8454143
      FrameThreeDShadowColor=   8454143
      FrameThreeDWidth=   4
      FrameWidth      =   8
      Picture         =   "frmWarningPostPayroll.frx":08CA
      Begin VB.Timer Timer1 
         Interval        =   355
         Left            =   6384
         Top             =   96
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
         Height          =   540
         Left            =   1350
         TabIndex        =   3
         Top             =   2550
         Width           =   1560
         _Version        =   131072
         _ExtentX        =   2752
         _ExtentY        =   952
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
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
         ButtonDesigner  =   "frmWarningPostPayroll.frx":08E6
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdPost 
         Height          =   540
         Left            =   3888
         TabIndex        =   4
         Top             =   2544
         Width           =   1572
         _Version        =   131072
         _ExtentX        =   2773
         _ExtentY        =   952
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
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
         ButtonDesigner  =   "frmWarningPostPayroll.frx":0AFC
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "READY TO POST PAYROLL?"
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
         Height          =   348
         Left            =   1248
         TabIndex        =   5
         Top             =   672
         Width           =   4284
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " Press ""F10"" to POST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1680
         TabIndex        =   2
         Top             =   1824
         Width           =   3468
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Press ""ESC"" to CANCEL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1536
         TabIndex        =   1
         Top             =   1392
         Width           =   3756
      End
   End
End
Attribute VB_Name = "frmWarningPostPayroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdEscape_Click()
  Unload frmWarningPostPayroll
End Sub

Private Sub cmdPost_Click()
    frmPostInProg.Show
    DoEvents
    Call PostTransactions    'from PRCalcMenu
    Call PostVoidCheckData 'added 6/17/04
'   -----------test-------------
'  Dim NumAccts As Integer
'  Dim VoidRec As VoidCheckType
'  Dim TVHandle As Integer
'  Dim X As Double, Y As Integer
'
'  OpenVoidChkPostFile TVHandle
'  NumAccts = LOF(TVHandle) / Len(VoidRec)
'  For X = 1 To NumAccts
'    Get TVHandle, X, VoidRec
'      Debug.Print VoidRec.PRNetGL + " PRNET            " + CStr(VoidRec.PRNet)
'      Debug.Print VoidRec.SOCWHGL + " SOC Withholdings " + CStr(VoidRec.SOCWHAmt)
'      Debug.Print VoidRec.MEDWHGL + " MED Withholdings " + CStr(VoidRec.MEDWHAmt)
'      Debug.Print VoidRec.SOCMATCRGL + " SOC Match Liab   " + CStr(VoidRec.SOCMATCRAmt)
'      Debug.Print VoidRec.MEDMATCRGL + " MED Match Liab   " + CStr(VoidRec.MEDMATCRAmt)
'      Debug.Print VoidRec.FEDWHGL + " FED Withholdings " + CStr(VoidRec.FEDWHAmt)
'      Debug.Print VoidRec.STAWHGL + " STA Withholdings " + CStr(VoidRec.STAWHAmt)
'      Debug.Print VoidRec.RETWHGL + " RET Withholdings " + CStr(VoidRec.RETWHAmt)
'      Debug.Print VoidRec.RETMATCRGL + " RET Match Liab   " + CStr(VoidRec.RETMATCRAmt)
'      For Y = 1 To 50
'        If VoidRec.DedData(Y).DAmt > 0 Then
'          Debug.Print VoidRec.DedData(Y).DedGLNum + " Deduction        " + CStr(VoidRec.DedData(Y).DAmt)
'        End If
'      Next Y
'      Debug.Print VoidRec.WagesGL + "  Wages           " + CStr(VoidRec.WagesAmt)
'      Debug.Print VoidRec.SOCMATDBGL + " SOC Match        " + CStr(VoidRec.SOCMATDBAmt)
'      Debug.Print VoidRec.MEDMATDBGL + " MED Match        " + CStr(VoidRec.MEDMATDBAmt)
'      Debug.Print VoidRec.RETMATDBGL + " RET Match        " + CStr(VoidRec.RETMATDBAmt)
'      Debug.Print VoidRec.CheckAmt
'      Debug.Print VoidRec.CheckDate
'      Debug.Print VoidRec.CheckNum
'      Debug.Print VoidRec.TransRec
'      Debug.Print VoidRec.VoidFlag
'  Next X
'  Close TVHandle

'   -----------test-------------
    
'    frmPayrollProcessingMenu.Show
    DoEvents
    Unload frmPostInProg
    Call DeActivateControls
    DoEvents
    Unload frmWarningPostPayroll
'    Call ActivateControls '8/6 not needed
    MainLog ("Payroll was posted.")
    KillFile ("prdata\ChecksPrinted.opn") '8/8

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
      Call cmdEscape_Click
      KeyCode = 0
    Case vbKeyF10:
      Call cmdPost_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
'  Me.HelpContextID = hlpPostPayroll
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdEscape.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      Call Terminate
      MainLog ("Payroll.exe terminated via menu bar on frmWarningPostPayroll.")
      End
    End If
  End If
End Sub
Private Sub DeActivateControls()
  Dim cnt As Integer
  Dim x As Control
  Dim cmdButton As CommandButton

  cmdEscape.Enabled = False
  cmdPost.Enabled = False
  
  For cnt = 0 To Me.Count - 1
    Set x = Me.Controls.Item(cnt)
      If TypeOf x Is CommandButton Then
        x.Enabled = False
      End If
  Next cnt
  
  EnableCloseButton Me.hwnd, False
     
End Sub

Private Sub ActivateControls()
  Dim cmdButton As CommandButton
  Dim x As Control
  Dim cnt As Integer
  
  cmdEscape.Enabled = True
  cmdPost.Enabled = True
  
  For cnt = 0 To Me.Count - 1
    Set x = Me.Controls.Item(cnt)
      If TypeOf x Is CommandButton Then
        x.Enabled = True
      End If
  Next cnt
  EnableCloseButton Me.hwnd, True
     
End Sub

Private Sub Timer1_Timer()
  Static tog As Boolean
  tog = Not tog
  If tog Then
    vaImprint1.BackColor = 210
  Else
    vaImprint1.BackColor = 192
  End If
End Sub

