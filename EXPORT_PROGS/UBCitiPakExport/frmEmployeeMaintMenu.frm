VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmEmployeeMaintMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Maintenance Menu"
   ClientHeight    =   8880
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11655
   Icon            =   "frmEmployeeMaintMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8654.619
   ScaleMode       =   0  'User
   ScaleWidth      =   11667.02
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn PrintEmplDataFileCmmd 
      Height          =   495
      Left            =   4005
      TabIndex        =   2
      Top             =   3690
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmEmployeeMaintMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn EditViewEmplRecordCmmd 
      Height          =   495
      Left            =   4005
      TabIndex        =   1
      Top             =   3075
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmEmployeeMaintMenu.frx":0AEE
   End
   Begin fpBtnAtlLibCtl.fpBtn AddNewEmplyeeCmmd 
      Height          =   495
      Left            =   4005
      TabIndex        =   0
      Top             =   2475
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmEmployeeMaintMenu.frx":0D13
   End
   Begin fpBtnAtlLibCtl.fpBtn PrintEmplListCmmd 
      Height          =   495
      Left            =   4005
      TabIndex        =   3
      Top             =   4290
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmEmployeeMaintMenu.frx":0F38
   End
   Begin fpBtnAtlLibCtl.fpBtn PrintTerminatedEmplListCmmd 
      Height          =   495
      Left            =   4005
      TabIndex        =   4
      Top             =   4890
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmEmployeeMaintMenu.frx":1157
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEmergency 
      Height          =   495
      Left            =   4005
      TabIndex        =   5
      Top             =   5505
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmEmployeeMaintMenu.frx":1381
   End
   Begin fpBtnAtlLibCtl.fpBtn exitCmd 
      Height          =   495
      Left            =   4005
      TabIndex        =   8
      Top             =   7320
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmEmployeeMaintMenu.frx":15AA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdQuickMaint 
      Height          =   495
      Left            =   4005
      TabIndex        =   6
      Top             =   6105
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmEmployeeMaintMenu.frx":17D4
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMailLabels 
      Height          =   495
      Left            =   4005
      TabIndex        =   7
      Top             =   6705
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmEmployeeMaintMenu.frx":19F6
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   2094
      Top             =   2102
      Width           =   971
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   8593
      Top             =   2103
      Width           =   971
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2223.291
      X2              =   2223.291
      Y1              =   2150.985
      Y2              =   7888.569
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EMPLOYEE  MAINTENANCE  MENU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2820
      TabIndex        =   9
      Top             =   1250
      Width           =   6012
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      Height          =   1097
      Left            =   1500
      Top             =   897
      Width           =   8655
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   9424.709
      X2              =   8721.985
      Y1              =   7894.417
      Y2              =   7894.417
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8721.985
      X2              =   8721.985
      Y1              =   2153.909
      Y2              =   7888.569
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2208.275
      X2              =   2923.011
      Y1              =   7894.417
      Y2              =   7894.417
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1214
      Left            =   1500
      Top             =   770
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2101
      Top             =   1971
      Width           =   972
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8592
      Top             =   1971
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5900
      Index           =   1
      Left            =   8711
      Top             =   2201
      Width           =   730
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5900
      Index           =   0
      Left            =   2220
      Top             =   2197
      Width           =   732
   End
End
Attribute VB_Name = "frmEmployeeMaintMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Public Enum NewEmpOpt
  neoInvalidOption = 0
  neoOn
  neoOff
End Enum
Private m_neoOption As NewEmpOpt
'// Create a property to get the Selection value.
'   NOTE: A Read-Only property has a Property Get but
'         no Property Let or Property Set
Property Get Selection() As NewEmpOpt
  Selection = m_neoOption
End Property

Private Sub AddNewEmplyeeCmmd_Click()
  InFileNames(1) = "PRDATA\PRDEDCOD.DAT"
  InFileNames(2) = "PRDATA\PRERNCOD.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 2) = False Then
    Close
    Exit Sub
  End If
  m_neoOption = neoOn
  
  frmBackGround.Show
  DoEvents
  frmLoadingEmpEdit.Show
  DoEvents
  frmEditEmpData.Show
  Unload frmLoadingEmpEdit
  Unload frmBackGround
  Unload frmEmployeeMaintMenu
  MainLog ("New Employee Data screen accessed.")
End Sub

Private Sub cmdEmergency_Click()
  frmEmergency.Show
  DoEvents
  Unload frmEmployeeMaintMenu
End Sub

Private Sub cmdMailLabels_Click()
  frmEmpMailLabels.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdQuickMaint_Click()
  frmEmpQuickMaintMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub EditViewEmplRecordCmmd_Click()
  InFileNames(1) = "PRDATA\PRDEDCOD.DAT"
  InFileNames(2) = "PRDATA\PRERNCOD.DAT"
  InFileNames(3) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 3) = False Then
    Close
    Exit Sub
  End If
  m_neoOption = neoOff
  frmEmployeeLookUp.Show
  DoEvents
  Unload frmEmployeeMaintMenu
End Sub

Private Sub exitCmd_Click()
   frmPayrollMainMenu.Show
   DoEvents
   Unload frmEmployeeMaintMenu
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%M"
      Call exitCmd_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpEmployeeFile
 End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Public Sub PrintEmplDataFileCmmd_Click()
  frmEmpDataPrint.Show
  DoEvents
  Unload frmEmployeeMaintMenu
  Exit Sub

End Sub

Private Sub PrintEmplListCmmd_Click()
  InFileNames(1) = "PRDATA\PREMP2.DAT"
  InFileNames(2) = "PRDATA\PRUNIT.DAT"
  InFileNames(3) = "PRDATA\PRPRNDF.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 3) = False Then
    Close
    Exit Sub
  End If
  frmPrintAlphaNum.Show
  DoEvents
  Unload frmEmployeeMaintMenu
End Sub

Private Sub PrintGraphics()
  Dim RptName As String, EmpIdxLNameHandle As Integer
  Dim ThisSort() As Integer
  Dim IdxNumOfRecs As Integer, UnitHandle As Integer
  Dim Emp2RecLen As Integer, UnitFileRec As UnitFileRecType
  Dim cnt As Integer
  Dim RptHandle As Integer, TermCnt As Integer
  Dim RptTitle As String, JobDesc As String * 20 '8/7
  Dim EmpIdxLNameRec As NameSortIdxType
  Dim x As Integer, City As String
  Dim EmpData2FileHandle As Integer
  Dim EmpData2FileRec As EmpData2Type
  Dim EmpIdxLNameCnt As Integer, NumOfTerms As Integer
  Dim FandLName As String * 20
  Dim dlm$
  
  dlm$ = "~"
  InFileNames(1) = "PRDATA\PREMP2.DAT"
  InFileNames(2) = "PRDATA\PRPRNDF.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 2) = False Then
    Close
    Exit Sub
  End If
  
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, UnitFileRec
  City = QPTrim$(UnitFileRec.UFEMPR)
  Close UnitHandle
  
  RptName$ = "PRRPTS\EMPrintTermEmpListG.RPT"
 
  RptTitle$ = "Terminated Employee Listing"
  
  OpenEmpIdxLNameFile EmpIdxLNameHandle
  EmpIdxLNameCnt = LOF(EmpIdxLNameHandle) / 2
  
  If EmpIdxLNameCnt = 0 Then
    MsgBox "No files on record."
    Close EmpIdxLNameHandle
    Exit Sub
  End If
  
  FrmShowPctComp.Label1 = "Terminated Employees Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  
  ReDim ThisSort(EmpIdxLNameCnt)
  For x = 1 To EmpIdxLNameCnt
     Get EmpIdxLNameHandle, x, ThisSort(x)
  Next x
  Close EmpIdxLNameHandle
  IdxNumOfRecs = EmpIdxLNameCnt

  If IdxNumOfRecs = 0 Then
     MsgBox "No employee entries found"
     GoTo EndTrans
  End If
  
  RptHandle = FreeFile
  On Error GoTo ErrorHandler
  Open RptName$ For Output As RptHandle
  OpenEmpData2File EmpData2FileHandle
  NumOfTerms = 0
  For cnt = 1 To IdxNumOfRecs
    If ThisSort(cnt) <> 0 Then
      Get EmpData2FileHandle, ThisSort(cnt), EmpData2FileRec
      If EmpData2FileRec.EMPTDATE = 0 Then GoTo NotTerm
      If Not EmpData2FileRec.Deleted Then
        FandLName = QPTrim$(EmpData2FileRec.EmpLName) + ", " + QPTrim$(EmpData2FileRec.EmpFName)
        JobDesc = QPTrim$(EmpData2FileRec.EMPJOB)
        NumOfTerms = NumOfTerms + 1
        Print #RptHandle, City; dlm; QPTrim$(EmpData2FileRec.EmpNo); dlm; FandLName; dlm;
        Print #RptHandle, JobDesc; dlm; MakeRegDate(EmpData2FileRec.EMPTDATE); dlm; Using$("######", NumOfTerms)
      End If
    End If
NotTerm:
    FrmShowPctComp.ShowPctComp cnt, IdxNumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next
  
  Close EmpData2FileHandle
  Close RptHandle
  
  If NumOfTerms = 0 Then
    MsgBox "There are no employees designated as terminated at this time."
    EnableCloseButton Me.hwnd, True
    Exit Sub
  End If
  
  arTermEmpRpt.Show
  frmLoadingRpt.Show
  EnableCloseButton Me.hwnd, True
  MainLog ("Terminated Employee Report processed.")
  
Exit Sub

ErrorHandler:
  Close
  Unload FrmShowPctComp
  MsgBox "ERROR: If this problem persists please consult Southern Software."

EndTrans:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If exitCmd.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmEmployeeMaintMenu.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub PrintText()
  Dim RptName As String, EmpIdxLNameHandle As Integer
  Dim ThisSort() As Integer
  Dim IdxNumOfRecs As Integer, UnitHandle As Integer
  Dim Emp2RecLen As Integer, UnitFileRec As UnitFileRecType
  Dim cnt As Integer, NumOfPages As Integer
  Dim RptHandle As Integer, TermCnt As Integer
  Dim RptTitle As String, JobDesc As String * 20 '8/7
  Dim EmpIdxLNameRec As NameSortIdxType
  Dim FF As String, x As Integer, City As String
  Dim EmpData2FileHandle As Integer
  Dim EmpData2FileRec As EmpData2Type
  Dim EmpIdxLNameCnt As Integer, NumOfTerms As Integer
  Dim LineCnt As Integer, MaxLines As Integer
  Dim FandLName As String * 20
  Dim Today As String * 11
  
  InFileNames(1) = "PRDATA\PREMP2.DAT"
  InFileNames(2) = "PRDATA\PRPRNDF.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 2) = False Then
    Close
    Exit Sub
  End If
  
'  Date$ = FormatDateTime(Date, vbShortDate)
  Today = Date '$
  MaxLines = 55
  NumOfPages = 1
  
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, UnitFileRec
  City = QPTrim$(UnitFileRec.UFEMPR)
  Close UnitHandle
  
  FF$ = Chr$(12)
  RptName$ = "EMPrintTermEmpList.RPT"
 
  RptTitle$ = "Terminated Employee Listing"
  
  OpenEmpIdxLNameFile EmpIdxLNameHandle
  EmpIdxLNameCnt = LOF(EmpIdxLNameHandle) / 2
  
  If EmpIdxLNameCnt = 0 Then
    MsgBox "No files on record."
    Close EmpIdxLNameHandle
    Exit Sub
  End If
  
  FrmShowPctComp.Label1 = "Terminated Employees Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  
  ReDim ThisSort(EmpIdxLNameCnt)
  For x = 1 To EmpIdxLNameCnt
     Get EmpIdxLNameHandle, x, ThisSort(x)
  Next x
  Close EmpIdxLNameHandle
  IdxNumOfRecs = EmpIdxLNameCnt

  If IdxNumOfRecs = 0 Then
     MsgBox "No employee entries found"
     GoTo EndTrans
  End If
  
  RptHandle = FreeFile
  
  Open RptName$ For Output As RptHandle
  RPTSetupPRN 3, RptHandle
  OpenEmpData2File EmpData2FileHandle
  GoSub PrintEmpListHeader
  NumOfTerms = 0
  For cnt = 1 To IdxNumOfRecs
    If ThisSort(cnt) <> 0 Then
      Get EmpData2FileHandle, ThisSort(cnt), EmpData2FileRec
      If EmpData2FileRec.EMPTDATE = 0 Then GoTo NotTerm
      If Not EmpData2FileRec.Deleted Then
        FandLName = QPTrim$(EmpData2FileRec.EmpLName) + ", " + QPTrim$(EmpData2FileRec.EmpFName)
        JobDesc = QPTrim$(EmpData2FileRec.EMPJOB)
        NumOfTerms = NumOfTerms + 1
        Print #RptHandle, Tab(2); QPTrim$(EmpData2FileRec.EmpNo);
        Print #RptHandle, Tab(12); FandLName;
        Print #RptHandle, Tab(37); JobDesc;
        Print #RptHandle, Tab(63); MakeRegDate(EmpData2FileRec.EMPTDATE)
        LineCnt = LineCnt + 1
        If LineCnt > MaxLines Then
          NumOfPages = NumOfPages + 1
          Print #RptHandle, FF$
          GoSub PrintEmpListHeader
        End If
      End If
    End If
NotTerm:
    FrmShowPctComp.ShowPctComp cnt, IdxNumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If

  Next
  Print #RptHandle, "-------------------------------------------------------------------------------"
  Print #RptHandle, "Total Terminated Employees: "; Using$("######", NumOfTerms)
  Print #RptHandle, FF$
  
  RPTSetupPRN 123, RptHandle '7/24
  Close EmpData2FileHandle
  Close RptHandle
  
  If NumOfTerms = 0 Then
    MsgBox "There are no employees designated as terminated at this time."
    EnableCloseButton Me.hwnd, True
    Exit Sub
  End If
  
  ViewPrint RptName$, RptTitle$
  EnableCloseButton Me.hwnd, True
  MainLog ("Terminated Employee Report processed.")
  
  If frmReportsProcessing.Selection = roOn Then 'added this
  'if statement on 8/24 so if this report is called
  'from the Reports Processing menu then when this sub
  'is done it exits back to that menu
    frmReportsProcessing.Show
    DoEvents
    Unload frmEmployeeMaintMenu
  End If
Exit Sub

PrintEmpListHeader:
  Print #RptHandle, City
  Print #RptHandle, "Report Date"; Tab(14); Today; Tab(62); "Page "; NumOfPages
  Print #RptHandle, ""
  Print #RptHandle, "Number            Name              Job Description        Termination Date"
  Print #RptHandle, "-------------------------------------------------------------------------------"
  LineCnt = 5
Return
EndTrans:
End Sub

Public Sub PrintTerminatedEmplListCmmd_Click()
  frmReportOpt.Show vbModal
  If RptOpt = 2 Then
    Call PrintText
    Exit Sub
  ElseIf RptOpt = 1 Then
    Call PrintGraphics
  Else
    Exit Sub
  End If

End Sub
