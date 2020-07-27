VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmResetPrices 
   AutoRedraw      =   -1  'True
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reset Decal Prices"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmResetPrices.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   336
      Left            =   0
      TabIndex        =   0
      Top             =   8532
      Width           =   12216
      _ExtentX        =   21548
      _ExtentY        =   593
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
            TextSave        =   "9:39 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "3/1/2006"
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
   Begin fpBtnAtlLibCtl.fpBtn fpCmdOK 
      Height          =   480
      Left            =   4086
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5256
      Width           =   1740
      _Version        =   131072
      _ExtentX        =   3069
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmResetPrices.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdExit 
      Height          =   480
      Left            =   6390
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5256
      Width           =   1740
      _Version        =   131072
      _ExtentX        =   3069
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmResetPrices.frx":0AA8
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Press F10 to Process or ESC to Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Left            =   3318
      TabIndex        =   5
      Top             =   3960
      Width           =   5604
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This procedure will Reset decal prices on each vehicle back to the original category price."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1284
      Left            =   3480
      TabIndex        =   4
      Top             =   2592
      Width           =   5292
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   2652
      Left            =   3126
      Top             =   2208
      Width           =   5964
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reset Decal Prices"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   3330
      TabIndex        =   1
      Top             =   1008
      Width           =   5652
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000005&
      Height          =   612
      Left            =   3222
      Top             =   888
      Width           =   5772
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   732
      Left            =   3228
      Top             =   768
      Width           =   5772
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   2724
      Left            =   3120
      Top             =   2208
      Width           =   6036
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   3108
      Left            =   2952
      Top             =   2016
      Width           =   6348
   End
End
Attribute VB_Name = "frmResetPrices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    'DoEvents
    Temp_Class.ResizeControls Me
   ' DoEvents
   ' Me.Visible = True
   ' Me.AutoRedraw = False
   ' DoEvents
  End If
  DoEvents
End Sub

Private Sub fpCmdExit_Click()
  frmDCCustomerMenu.Show
  Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If fpCmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        DCLog "Closed via ClearDecals by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub
Private Sub Form_Load()
  Dim dcSetUpRec(1) As DCSetupType
  Dim RecLen As Integer
  'BlockInput True
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      KeyCode = 0
      Call fpCmdExit_Click
    Case vbKeyF10, vbKeyReturn
      KeyCode = 0
      Call fpCmdOk_Click
    Case Else:
  End Select
End Sub

Private Sub fpCmdOk_Click()
  ResetCustomer
End Sub
Private Sub ResetCustomer()
  Dim NumOfVRecs As Long, DCvFile As Integer, DCVehReclen As Integer
  Dim TrHandle As Integer, IdxCustRecLen As Integer, IdxTrHandle As Integer
  ReDim DCVRec(1) As DCVehType
  Dim TrNumRecs  As Long, CustRecLen As Integer, IdxTrNumRecs As Long
  Dim cnt As Long, CarRecord As Long, DCOVreclen As Integer, CodeRecCnt As Long
  Dim cntC As Long, DCCodeRecLen As Integer, CodeFile As Integer
  ReDim DCCustRec(1) As DCCustRecType
  Dim DCCodeRec As DCCatCodeRecType
  FrmShowPctComp.Label1 = "Reset Decal Prices"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , frmResetPrices
 
  
  DCCodeRecLen = Len(DCCodeRec)
  CodeFile = FreeFile
  Open DCPath + "DCCODE.DAT" For Random Shared As CodeFile Len = DCCodeRecLen
  CodeRecCnt = LOF(CodeFile) \ DCCodeRecLen

  'Open Vehicle File
  DCVehReclen = Len(DCVRec(1))
  DCvFile = FreeFile
  Open "DCVEH.DAT" For Random Access Read Write Shared As DCvFile Len = DCVehReclen
  NumOfVRecs = LOF(DCvFile) \ DCVehReclen

   ' open customer file
  CustRecLen = Len(DCCustRec(1))
  TrHandle = FreeFile
  Open "DCCUST.DAT" For Random Access Read Write Shared As TrHandle Len = CustRecLen
  TrNumRecs = LOF(TrHandle) \ CustRecLen
   For cnt = 1 To TrNumRecs
    FrmShowPctComp.ShowPctComp cnt, TrNumRecs
    Get TrHandle, cnt, DCCustRec(1)
    If DCCustRec(1).Deleted <> "Y" Then
      CarRecord = DCCustRec(1).FirstCar
      While CarRecord > 0
        Get DCvFile, CarRecord, DCVRec(1)
        If DCVRec(1).Active = "Y" Then
          For cntC = 1 To CodeRecCnt
            Get CodeFile, cntC, DCCodeRec
            If QPTrim$(DCVRec(1).DecalCat) = QPTrim$(DCCodeRec.CATCODE) Then
              DCVRec(1).Fee = DCCodeRec.Fee
              Put DCvFile, CarRecord, DCVRec(1)
              Exit For
            End If
          Next
         End If
        CarRecord = DCVRec(1).NextRec
      Wend
     End If
   Next cnt
  Close         'Close all open files now
  DCLog PWUser$ + " Reset Prices"
MsgBox "Reset Decal Prices Complete.", vbOKOnly, "Complete"

End Sub
