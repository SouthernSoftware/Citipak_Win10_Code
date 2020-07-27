VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmWarningYearEnd 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Year End Warning"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmWarningYearEnd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdNo 
      Height          =   540
      Left            =   6240
      TabIndex        =   10
      Top             =   7488
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
      ButtonDesigner  =   "frmWarningYearEnd.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdYes 
      Height          =   540
      Left            =   3750
      TabIndex        =   11
      Top             =   7485
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
      ButtonDesigner  =   "frmWarningYearEnd.frx":0AA4
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000000C0&
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Height          =   7932
      Left            =   1860
      Top             =   468
      Width           =   7932
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      X1              =   2916.249
      X2              =   8795.735
      Y1              =   4629.442
      Y2              =   4629.442
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      Height          =   612
      Left            =   2916
      Top             =   6312
      Width           =   5892
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      Height          =   2292
      Left            =   2916
      Top             =   3552
      Width           =   5892
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      Height          =   1572
      Left            =   2916
      Top             =   1392
      Width           =   5892
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Are you ready to begin a new YEAR?"
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
      Height          =   492
      Left            =   2556
      TabIndex        =   9
      Top             =   6432
      Width           =   6612
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "the Payroll Data files."
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
      Height          =   492
      Left            =   2556
      TabIndex        =   8
      Top             =   5352
      Width           =   6612
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please make sure you have BACKED UP"
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
      Height          =   492
      Left            =   2556
      TabIndex        =   7
      Top             =   4992
      Width           =   6612
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WILL NOT BE AFFECTED!"
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
      Height          =   492
      Left            =   2556
      TabIndex        =   6
      Top             =   4152
      Width           =   6612
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "W-2 processing for the prior year"
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
      Height          =   492
      Left            =   2556
      TabIndex        =   5
      Top             =   3672
      Width           =   6612
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " HOWEVER!"
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
      Height          =   372
      Left            =   4716
      TabIndex        =   4
      Top             =   3072
      Width           =   2532
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " will ZERO the Employee's YTD Totals."
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
      Height          =   492
      Left            =   2556
      TabIndex        =   3
      Top             =   2472
      Width           =   6612
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " payroll processing. This procedure"
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
      Height          =   492
      Left            =   2556
      TabIndex        =   2
      Top             =   1992
      Width           =   6612
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You are about to start a NEW YEAR of"
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
      Height          =   492
      Left            =   2556
      TabIndex        =   1
      Top             =   1512
      Width           =   6612
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WARNING!!!"
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
      Height          =   372
      Left            =   4716
      TabIndex        =   0
      Top             =   912
      Width           =   2532
   End
End
Attribute VB_Name = "frmWarningYearEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      Call cmdNO_Click
      SendKeys "%N"
      KeyCode = 0
    Case vbKeyF10:
      Call cmdYes_Click
      SendKeys "%P"
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Me.HelpContextID = hlpInitializeNewYear
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub cmdNO_Click()
  frmControlFileMaint.Show
  DoEvents
  Unload frmWarningYearEnd
  MainLog ("Year End Warning issued...NO option chosen.")
  
End Sub

Private Sub cmdYes_Click()
  Call ClearTotals
  MainLog ("Year End Warning issued...YES option chosen.")
End Sub

Sub ClearTotals()
  Dim EmpData3FileHandle As Integer, x As Long
  Dim EmpData3FileRec As EmpData3Type
  Dim Emp3RecNum As Long, y As Integer
  OpenEmpData3File EmpData3FileHandle
  Emp3RecNum = LOF(EmpData3FileHandle) \ Len(EmpData3FileRec)
  
  If Emp3RecNum = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If
  For x = 1 To Emp3RecNum
  Get EmpData3FileHandle, x, EmpData3FileRec
    EmpData3FileRec.Data1RecNum = 0
    EmpData3FileRec.YTDGrossPay = 0
    EmpData3FileRec.YTDSocGrossPay = 0
    EmpData3FileRec.YTDMedGrossPay = 0
    EmpData3FileRec.YTDFedGrossPay = 0
    EmpData3FileRec.YTDStaGrossPay = 0
    EmpData3FileRec.YTDOTPay = 0
    EmpData3FileRec.YTDRegPay = 0
    EmpData3FileRec.YTDNet = 0
    EmpData3FileRec.YTDSocial = 0
    EmpData3FileRec.YTDMedicare = 0
    EmpData3FileRec.YTDFederal = 0
    EmpData3FileRec.YTDState = 0
    EmpData3FileRec.YTDRetire = 0
    For y = 1 To 50 ' changed from 12 to 50 on 1/31/2005
      EmpData3FileRec.YTDDAmt(y) = 0
    Next y
    EmpData3FileRec.YTDDAmtT = 0
    EmpData3FileRec.YTDEarn1 = 0
    EmpData3FileRec.YTDEarn2 = 0
    EmpData3FileRec.YTDEarn3 = 0
    EmpData3FileRec.YTDEarnT = 0
    EmpData3FileRec.YTDEIC = 0
    EmpData3FileRec.YTDOther2 = 0
  Put EmpData3FileHandle, x, EmpData3FileRec
  Next x
  Close EmpData3FileHandle
  MsgBox "Data files have been updated", vbOKOnly
  frmControlFileMaint.Show
  DoEvents
  Unload frmWarningYearEnd
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdNo.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmWarningYearEnd.")
      Call Terminate
      End
    End If
  End If
End Sub

