VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRelinkDialog 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relink Transaction History"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   ControlBox      =   0   'False
   Icon            =   "frmRelinkDialog.frx":0000
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
      Height          =   4860
      Left            =   2994
      TabIndex        =   1
      Top             =   1368
      Width           =   6228
      Begin VB.Timer frmBlinkTimer 
         Interval        =   333
         Left            =   5736
         Top             =   4320
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdCancel 
         Height          =   372
         Left            =   3324
         TabIndex        =   2
         Top             =   4176
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
         ButtonDesigner  =   "frmRelinkDialog.frx":030A
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdOK 
         Height          =   372
         Left            =   1740
         TabIndex        =   3
         Top             =   4176
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
         ButtonDesigner  =   "frmRelinkDialog.frx":04E3
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
         Left            =   840
         TabIndex        =   10
         Top             =   456
         Width           =   4632
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
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
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   6000
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
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
         Left            =   120
         TabIndex        =   8
         Top             =   1416
         Width           =   6000
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "WITH THIS PROCEDURE!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   372
         Index           =   3
         Left            =   1344
         TabIndex        =   7
         Top             =   1968
         Width           =   3552
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CONTINUE WITH RELINK?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   348
         Index           =   4
         Left            =   684
         TabIndex        =   6
         Top             =   3432
         Width           =   4872
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "FAILURE TO STOP ALL OPERATIONS COULD RESULT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   324
         Index           =   5
         Left            =   120
         TabIndex        =   5
         Top             =   2496
         Width           =   6000
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "IN YOUR UTILITY BILLING DATA BEING DESTROYED!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   324
         Index           =   6
         Left            =   120
         TabIndex        =   4
         Top             =   2808
         Width           =   6000
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
End
Attribute VB_Name = "frmRelinkDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim FntSize As Double
Dim UBRTCustRec   As NewUBCustRecType
Dim UBRTTransRec  As UBTransRecType
Dim WorkOrderRec As WorkOrderRecType

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
  frmUBSetupMenu.cmdRelinkHistory.SetFocus
  DoEvents
  Unload frmRelinkDialog
End Sub

Private Sub cmdOk_Click()
  frmRelinkDialog.frmBlinkTimer.Enabled = True
  DeActivateControls Me
  DoEvents
  Call UBRelinkTransactions
  DoEvents
  ActivateControls Me
  Call cmdCancel_Click
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
   ' Me.Visible = False
    Temp_Class.ResizeControls Me
   ' Me.Visible = True
   ' Me.SetFocus
  End If
  DoEvents
  frmRelinkDialog.frmBlinkTimer.Enabled = False
End Sub

Private Sub UBRelinkTransactions()
  
  DoEvents
  UBLog " IN: Relink Utility Files"
  
  Dim UBCustRecLen As Integer, UBTranRecLen As Integer, WorkOrderRecLen As Integer
  Dim UBFile As Integer, UBTran As Integer, UBWrkOrd As Integer
  Dim NumOfCRecs As Long, NumOfTRecs As Long, NumOfWORecs As Long
  Dim OddRecs As Integer, RecCnt As Long
  Dim TRRecs As Long, PutRec As Long
  Dim CCnt As Long, ChkCnt As Long
  Dim BlockSize As Long
  Dim NumChunks As Long
  Dim MaxBlockCnt As Integer
  
  UBCustRecLen = Len(UBRTCustRec)              'Length of Cust Record Structure
  UBTranRecLen = Len(UBRTTransRec)             'Length of Tran Record Structure
  WorkOrderRecLen = Len(WorkOrderRec)
  
  frmRelinkDialog.Label(1).FontSize = frmRelinkDialog.Label(1).FontSize + 2
  frmRelinkDialog.Label(2).FontSize = frmRelinkDialog.Label(1).FontSize
  
  frmRelinkDialog.Label(1) = "NO UTILITY BILLING OPERATIONS UNTIL"
  frmRelinkDialog.Label(2) = "RELINK FILES COMPLETES!"
  
  DoEvents
  
  FrmShowPctComp.Label1 = "Checking Customers."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show
  
  DoEvents
  
  UBLog "BEGIN: Pass 1 of 3"
  
  UBTran = FreeFile
  Open UBPath + "UBTRANS.DAT" For Random Shared As UBTran Len = UBTranRecLen
  NumOfTRecs = LOF(UBTran) \ UBTranRecLen

  UBFile = FreeFile
  Open UBPath + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
  NumOfCRecs = LOF(UBFile) \ UBCustRecLen
    
  For CCnt = 1 To NumOfCRecs
    Get UBFile, CCnt, UBRTCustRec
    UBRTCustRec.LastTrans = 0
    UBRTCustRec.WOLastTrans = 0
    Put UBFile, CCnt, UBRTCustRec
    ChkCnt = ChkCnt + 1
    If ChkCnt >= 100 Then
      FrmShowPctComp.ShowPctComp CCnt, NumOfCRecs
      ChkCnt = 0
    End If
  Next
  
  Unload FrmShowPctComp
  
  DoEvents
  FrmShowPctComp.Label1 = "Relinking Transactions."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show
  DoEvents
    
  UBLog "       Pass 2 of 3"
  MaxBlockCnt = 1024
  
  ReDim TransBuff(1 To MaxBlockCnt) As UBTransRecType
''************************************
  NumChunks& = NumOfTRecs \ MaxBlockCnt
''****DO NOT CHANGE THE DIVISION HERE!
  OddRecs = UBMod(NumOfTRecs, MaxBlockCnt)

  If NumChunks& = 0 Then        'if the actual cust count is less than
    MaxBlockCnt = OddRecs       'the work buffer
    NumChunks& = 1
    OddRecs = 0
  End If
  
  For CCnt& = 1 To NumChunks&
    For RecCnt = 1 To MaxBlockCnt
      TRRecs = TRRecs + 1
      Get UBTran, TRRecs, TransBuff(RecCnt)
      TransBuff(RecCnt).PenAtBill = TRRecs
    Next
    For RecCnt = 1 To MaxBlockCnt
      If (TransBuff(RecCnt).CustAcctNo > 0) And (TransBuff(RecCnt).CustAcctNo <= NumOfCRecs) Then
        Get UBFile, TransBuff(RecCnt).CustAcctNo, UBRTCustRec
        TransBuff(RecCnt).PrevTrans = UBRTCustRec.LastTrans
        PutRec = TransBuff(RecCnt).PenAtBill
        UBRTCustRec.LastTrans = PutRec
        Put UBFile, TransBuff(RecCnt).CustAcctNo, UBRTCustRec
        Put UBTran, PutRec, TransBuff(RecCnt)
      End If
    Next
    FrmShowPctComp.ShowPctComp TRRecs, NumOfTRecs
  Next
  
  If OddRecs Then
    For CCnt = TRRecs + 1 To NumOfTRecs
      Get UBTran, CCnt, TransBuff(1)
      TransBuff(1).PenAtBill = CCnt
      If (TransBuff(1).CustAcctNo > 0) And (TransBuff(1).CustAcctNo <= NumOfCRecs) Then
        Get UBFile, TransBuff(1).CustAcctNo, UBRTCustRec
        TransBuff(1).PrevTrans = UBRTCustRec.LastTrans
        PutRec = TransBuff(1).PenAtBill
        UBRTCustRec.LastTrans = PutRec
        Put UBFile, TransBuff(1).CustAcctNo, UBRTCustRec
        Put UBTran, PutRec, TransBuff(1)
      End If
      FrmShowPctComp.ShowPctComp CCnt, NumOfTRecs
    Next
  End If
  Close UBTran
  Unload FrmShowPctComp
  DoEvents
  Erase TransBuff
  
  UBLog "       Pass 3 of 3"
  
  FrmShowPctComp.Label1 = "Relinking WorkOrders."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show
  DoEvents
    
  UBWrkOrd = FreeFile
  Open UBPath + "UBWRKORD.DAT" For Random Shared As UBWrkOrd Len = WorkOrderRecLen
  NumOfWORecs& = LOF(UBWrkOrd) \ WorkOrderRecLen
 
  For CCnt = 1 To NumOfWORecs
    Get UBWrkOrd, CCnt, WorkOrderRec
    If (WorkOrderRec.CustRec > 0) And (WorkOrderRec.CustRec <= NumOfCRecs) Then
      Get UBFile, WorkOrderRec.CustRec, UBRTCustRec
      WorkOrderRec.PrevTransRec = UBRTCustRec.WOLastTrans
      UBRTCustRec.WOLastTrans = CCnt
      Put UBFile, WorkOrderRec.CustRec, UBRTCustRec
      Put UBWrkOrd, CCnt&, WorkOrderRec
    End If
    FrmShowPctComp.ShowPctComp CCnt, NumOfWORecs
  Next
  Unload FrmShowPctComp
  
  Close
  DoEvents
  UBLog "RELINK: Utility Files Completed."
  ReIndexSystem False
  Call UPDateOK
  
  'Unload frmDataUpdated
  
  DoEvents

ExitRelink:
  UBLog "OUT: Relink Transaction History" + CrLf$
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
    Me.Frame1.BackColor = 220
    'Me.Label(1).BackColor = 230
    'Me.Label(2).BackColor = 230
  Else
    Me.Frame1.BackColor = &HC0&
    'Me.Label(1).BackColor = 192
    'Me.Label(2).BackColor = 192
  End If
  DoEvents
End Sub
