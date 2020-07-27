VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmUBRecalcConsumption 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recalculate Average Consumption"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   ControlBox      =   0   'False
   Icon            =   "frmUBRecalcConsumption.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H000000C0&
      Height          =   4284
      Left            =   3054
      TabIndex        =   4
      Top             =   1656
      Width           =   6108
      Begin VB.Timer frmBlinkTimer 
         Interval        =   333
         Left            =   2904
         Top             =   3408
      End
      Begin EditLib.fpLongInteger fpNumMonths 
         Height          =   300
         Left            =   3480
         TabIndex        =   1
         Top             =   2472
         Width           =   396
         _Version        =   196608
         _ExtentX        =   698
         _ExtentY        =   529
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
         BorderColor     =   -2147483642
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
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   0
         MarginRight     =   3
         MarginBottom    =   0
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "12"
         MaxValue        =   "36"
         MinValue        =   "1"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdCancel 
         Height          =   372
         Left            =   3276
         TabIndex        =   3
         Top             =   3384
         Width           =   1188
         _Version        =   131072
         _ExtentX        =   2096
         _ExtentY        =   656
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
         DrawFocusRect   =   3
         DrawFocusRectCell=   -1
         GrayAreaPictureStyle=   0
         Static          =   0   'False
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   1
         DropShadowOffsetY=   1
         DropShadowType  =   0
         DropShadowColor =   0
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmUBRecalcConsumption.frx":030A
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdOK 
         Height          =   372
         Left            =   1692
         TabIndex        =   2
         Top             =   3408
         Width           =   1188
         _Version        =   131072
         _ExtentX        =   2096
         _ExtentY        =   656
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
         DrawFocusRect   =   3
         DrawFocusRectCell=   -1
         GrayAreaPictureStyle=   0
         Static          =   0   'False
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   1
         DropShadowOffsetY=   1
         DropShadowType  =   0
         DropShadowColor =   0
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmUBRecalcConsumption.frx":04E3
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "WITH THIS PROCEDURE!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   348
         Index           =   3
         Left            =   408
         TabIndex        =   9
         Top             =   1632
         Width           =   5496
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "UTILITY BILLING PROGRAM BEFORE CONTINUING"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   348
         Index           =   2
         Left            =   408
         TabIndex        =   8
         Top             =   1272
         Width           =   5496
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "ALL UTILITY BILLING OPERATORS MUST EXIT THE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   348
         Index           =   1
         Left            =   408
         TabIndex        =   7
         Top             =   912
         Width           =   5496
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "WARNING   WARNING  WARNING"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   372
         Index           =   0
         Left            =   768
         TabIndex        =   6
         Top             =   384
         Width           =   4632
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Months:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   300
         Index           =   4
         Left            =   960
         TabIndex        =   5
         Top             =   2496
         Width           =   2376
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
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
            TextSave        =   "9:01 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "2/18/2005"
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
End
Attribute VB_Name = "frmUBRecalcConsumption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim FntSize As Double, AvgCnt As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      KeyCode = 0
      Call cmdCancel_Click
    Case Else:
  End Select
End Sub

Private Sub cmdOk_GotFocus()
  If FntSize <= 0 Then
    FntSize = cmdOk.FontSize
  End If
  cmdOk.FontBold = True
  cmdCancel.FontBold = False
  cmdOk.FontSize = FntSize + 1
  cmdCancel.FontSize = FntSize - 1
End Sub

Private Sub cmdCancel_GotFocus()
  If FntSize <= 0 Then
    FntSize = cmdCancel.FontSize
  End If
  cmdOk.FontBold = False
  cmdCancel.FontBold = True
  cmdCancel.FontSize = FntSize + 1
  cmdOk.FontSize = FntSize - 1
End Sub
  
Private Sub cmdCancel_Click()
  FntSize = 0
  Load frmUBSetupMenu
  DoEvents
  frmUBSetupMenu.Show
  'frmUBSetupMenu.cmdRelinkHistory.SetFocus
  DoEvents
  Unload Me
End Sub

Private Sub cmdOk_Click()
  AvgCnt = fpNumMonths.Value
  If AvgCnt < 2 Or AvgCnt > 36 Then
    MsgBox "         Invalid Number of Months!         " & Chr$(13) & "         Range:  2 - 36", vbOKOnly
    fpNumMonths.Value = 2
  Else
    DeActivateControls Me
    DoEvents
    Call RecalcConsumption
    DoEvents
    ActivateControls Me
    Call cmdCancel_Click
  End If

End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  frmBlinkTimer.Enabled = False
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ' Me.Visible = False
    Temp_Class.ResizeControls Me
   ' Me.Visible = True
   ' Me.SetFocus
  End If
  DoEvents
End Sub

Private Sub RecalcConsumption()
  frmBlinkTimer.Enabled = True
  
  DoEvents
  UBLog " IN: Recalc Average Consumption"
  
  Dim UBCustRecLen As Integer
  Dim UBTranRecLen As Integer
  Dim UBFile As Integer
  Dim UBTran As Integer
  Dim NumOfCRecs As Long
  Dim CCnt As Long
  Dim LastTran As Long
  Dim DidCnt As Integer
  Dim MCnt As Integer
  
  ReDim UBCustRec(1) As NewUBCustRecType
  ReDim UBTranRec(1) As UBTransRecType
  
  UBCustRecLen = Len(UBCustRec(1))              'Length of Cust Record Structure
  UBTranRecLen = Len(UBTranRec(1))             'Length of Tran Record Structure
  
  frmUBRecalcConsumption.Label(1) = "NO UTILITY BILLING OPERATIONS UNTIL"
  frmUBRecalcConsumption.Label(1).Alignment = 2
  frmUBRecalcConsumption.Label(2) = "RECALCULATE CONSUMPTION COMPLETES!"
  frmUBRecalcConsumption.Label(2).Alignment = 2
  frmUBRecalcConsumption.Label(3) = ""
  
  DoEvents
  
  UBLog "AVG: Recalc Using" + Str$(AvgCnt) + " Months"
  FrmShowPctComp.Label1 = "Recalculating Consumption."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show
  
  DoEvents
  
  UBFile = FreeFile
  Open UBPath + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
  
  UBTran = FreeFile
  Open UBPath + "UBTRANS.DAT" For Random Shared As UBTran Len = UBTranRecLen

  NumOfCRecs = LOF(UBFile) \ UBCustRecLen
  
  For CCnt = 1 To NumOfCRecs
    FrmShowPctComp.ShowPctComp CCnt, NumOfCRecs
    Get UBFile, CCnt, UBCustRec(1)
    ReDim TotalUse(1 To 7) As Long
    ReDim UseCnt(1 To 7) As Integer
    DidCnt = 0
    LastTran& = UBCustRec(1).LastTrans
    Do While LastTran& > 0
      Get UBTran, LastTran&, UBTranRec(1)
      If UBTranRec(1).TransType = TranUtilityBill Then
        For MCnt = 1 To 7
          If UBTranRec(1).CurRead(MCnt) > 0 Then
            TotalUse(MCnt) = TotalUse(MCnt) + (UBTranRec(1).CurRead(MCnt) - UBTranRec(1).PrevRead(MCnt))
            UseCnt(MCnt) = UseCnt(MCnt) + 1
          End If
        Next
        DidCnt = DidCnt + 1
        If DidCnt >= AvgCnt Then
          Exit Do
        End If
      End If
      LastTran& = UBTranRec(1).PrevTrans
    Loop

    For MCnt = 1 To 7
      If TotalUse(MCnt) > 0 Then
        UBCustRec(1).LocMeters(MCnt).AvgUse = TotalUse(MCnt) / UseCnt(MCnt)
        UBCustRec(1).LocMeters(MCnt).UseCnt = UseCnt(MCnt)
      End If
    Next
    Put UBFile, CCnt, UBCustRec(1)
  Next
  Close
  Erase UBCustRec, UBTranRec
  UBLog "AVG: Recalc Average Use Complete."
  DoEvents
  Call UPDateOK
  DoEvents
  UBLog "OUT: Recalc Average Consumption" + CrLf$
 ' Call cmdCancel_Click
End Sub

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'  Cancel = True
'  DoEvents
'End Sub

Private Sub frmBlinkTimer_Timer()
  Dim BkColor As Integer
  Static tog As Boolean
  tog = Not tog
  If tog Then
    Me.Frame1.BackColor = 210
  Else
    Me.Frame1.BackColor = &HC0&
  End If
  DoEvents
End Sub
