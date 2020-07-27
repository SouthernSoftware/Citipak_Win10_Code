VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUBMeterMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Meter Reading Menu"
   ClientHeight    =   8868
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   12216
   Icon            =   "frmUBMeterMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdManualReadings 
      BackColor       =   &H008F8265&
      Caption         =   "&Manual Meter Reading Entry"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   3846
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   1992
      Width           =   4524
   End
   Begin VB.CommandButton cmdHHReadEntry 
      BackColor       =   &H008F8265&
      Caption         =   "&Hand-Held Meter Reading Entry"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   3846
      MaskColor       =   &H8000000F&
      TabIndex        =   1
      Top             =   2544
      Width           =   4524
   End
   Begin VB.CommandButton cmdEstimateReads 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Estimated Meter Reading Entry"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   3864
      TabIndex        =   2
      Top             =   3084
      Width           =   4524
   End
   Begin VB.CommandButton cmdReadingReport 
      Caption         =   "Print Reading &Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   3846
      TabIndex        =   4
      Top             =   4176
      Width           =   4524
   End
   Begin VB.CommandButton cmdHiLowReport 
      Caption         =   "Print &High/Low Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   3846
      TabIndex        =   5
      Top             =   4728
      Width           =   4524
   End
   Begin VB.CommandButton cmdInactiveConsumption 
      Caption         =   "Inactive &Consumption Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   3846
      TabIndex        =   6
      Top             =   5280
      Width           =   4524
   End
   Begin VB.CommandButton cmdReadingList 
      Caption         =   "Print Meter Reading &List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   3864
      TabIndex        =   3
      Top             =   3636
      Width           =   4524
   End
   Begin VB.CommandButton cmdReadingSheets 
      Caption         =   "Print Meter Reading &Sheets"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   3846
      TabIndex        =   8
      Top             =   6372
      Width           =   4524
   End
   Begin VB.CommandButton cmdReadingNotes 
      Caption         =   "Print Meter Reading &Notes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   3846
      TabIndex        =   7
      Top             =   5820
      Width           =   4524
   End
   Begin VB.CommandButton cmdExitUBMeterMenu 
      Caption         =   "E&xit to Previous Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   3846
      TabIndex        =   10
      Top             =   7464
      Width           =   4524
   End
   Begin VB.CommandButton cmdMetersNoSerialNumbers 
      Caption         =   "Meters &NO Serial Numbers"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   3846
      TabIndex        =   9
      Top             =   6912
      Width           =   4524
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   11
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
            TextSave        =   "11:06 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "4/1/2004"
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Meter Reading Menu"
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
      Left            =   3540
      TabIndex        =   12
      Top             =   1092
      Width           =   5148
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   9828
      X2              =   9828
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   8868
      X2              =   8868
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   3
      X1              =   8868
      X2              =   9828
      Y1              =   1944
      Y2              =   1944
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   8868
      X2              =   9828
      Y1              =   2064
      Y2              =   2064
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   8988
      X2              =   9708
      Y1              =   7944
      Y2              =   7944
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   8988
      X2              =   8988
      Y1              =   2064
      Y2              =   7944
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   3348
      X2              =   3348
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2388
      X2              =   2388
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   2388
      X2              =   3348
      Y1              =   1944
      Y2              =   1944
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2388
      X2              =   3348
      Y1              =   2064
      Y2              =   2064
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2508
      X2              =   3228
      Y1              =   7944
      Y2              =   7944
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2508
      X2              =   2508
      Y1              =   2064
      Y2              =   7944
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1788
      Top             =   744
      Width           =   8652
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1788
      Top             =   624
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2388
      Top             =   1824
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2508
      Top             =   2064
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   8868
      Top             =   1824
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8988
      Top             =   2064
      Width           =   732
   End
End
Attribute VB_Name = "frmUBMeterMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdEstimateReads_Click()
  frmEstimateMtrReads.Show
  Unload frmUBMeterMenu
End Sub

Private Sub cmdExitUBMeterMenu_Click()
  frmUBBillingMenu.Show
  Unload frmUBMeterMenu
End Sub

Private Sub cmdHHReadEntry_Click()

  frmUBHHMenu.Show
  Unload frmUBMeterMenu
End Sub

Private Sub cmdHiLowReport_Click()

  frmRptHighLowReads.Show
  Unload frmUBMeterMenu
End Sub

Private Sub cmdInactiveConsumption_Click()
  frmRptInactiveConsump.Show
  Unload frmUBMeterMenu
End Sub


Private Sub cmdManualReadings_Click()
  frmManualMeterOpt.Show
  Unload frmUBMeterMenu
End Sub

Private Sub cmdReadingList_Click()
'send value of 1 for list
  frmRptMeterReadList.GetLorR 1
  frmRptMeterReadList.Show
  Unload frmUBMeterMenu
End Sub

Private Sub cmdReadingNotes_Click()
  Load frmRptMeterNotes
  frmRptMeterNotes.Exit2Flag = "1"
  DoEvents
  frmRptMeterNotes.Show
  'frmRptMeterNotes.Show
  'Unload frmUBMeterMenu
End Sub

Private Sub cmdReadingReport_Click()
'send value of 2 for report
  frmRptMeterReadList.GetLorR 2
  frmRptMeterReadList.Label1.Caption = "Print Meter Reading Report"
  frmRptMeterReadList.Caption = "Meter Reading Report"
  frmRptMeterReadList.Show
  Unload frmUBMeterMenu
End Sub

Private Sub cmdReadingSheets_Click()
  frmRptMeterSheets.Show
  Unload frmUBMeterMenu
End Sub
Private Sub cmdMetersNoSerialNumbers_Click()
  frmRptMtrNoSerial.Show
  Unload frmUBMeterMenu
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
   ' Me.Visible = True
   ' Me.SetFocus
  End If
  DoEvents
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitUBMeterMenu.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via MeterMenu by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExitUBMeterMenu_Click
      KeyCode = 0
    Case vbKeyHome
      KeyCode = 0
      cmdManualReadings.SetFocus
    Case vbKeyEnd
      KeyCode = 0
      cmdExitUBMeterMenu.SetFocus
    Case Else:
  End Select
End Sub

