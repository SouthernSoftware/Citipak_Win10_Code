VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmFAInternalOnly 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Southern Software Only!"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmFAInternalOnly.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   4092
      Left            =   1956
      TabIndex        =   1
      Top             =   2376
      Width           =   7740
      _Version        =   196609
      _ExtentX        =   13652
      _ExtentY        =   7218
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmFAInternalOnly.frx":08CA
      Begin VB.TextBox fptxtYear 
         Alignment       =   2  'Center
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
         Left            =   4272
         MaxLength       =   4
         TabIndex        =   0
         ToolTipText     =   "Enter the year for which a depreciation reversal is necessary."
         Top             =   1872
         Width           =   1116
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   675
         Left            =   1590
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2835
         Width           =   1875
         _Version        =   131072
         _ExtentX        =   3307
         _ExtentY        =   1191
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
         GrayAreaColor   =   13684944
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
         ButtonDesigner  =   "frmFAInternalOnly.frx":08E6
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   690
         Left            =   4464
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2832
         Width           =   1890
         _Version        =   131072
         _ExtentX        =   3334
         _ExtentY        =   1217
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
         GrayAreaColor   =   13684944
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
         ButtonDesigner  =   "frmFAInternalOnly.frx":0AC2
      End
      Begin VB.Label Label2 
         BackColor       =   &H00D0D0D0&
         Caption         =   "Year To Delete:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   2400
         TabIndex        =   3
         Top             =   1968
         Width           =   1740
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   684
         Left            =   1488
         Top             =   576
         Width           =   4908
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Depreciation Deletion"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   1590
         TabIndex        =   2
         Top             =   720
         Width           =   4815
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   4284
      Left            =   1860
      Top             =   2292
      Width           =   7932
   End
End
Attribute VB_Name = "frmFAInternalOnly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  frmFAMainMenu.Show
  Close
  DoEvents
  Unload frmFAInternalOnly
End Sub

Private Sub cmdProcess_Click()
  Dim x As Long
  Dim DprRec As DprHistType
  Dim DPRHandle As Integer
  Dim DprCnt As Long
  Dim FAItemRec As FAItemRecType
  Dim FAHandle As Integer
  Dim NumOfFARecs As Long
  Dim ThisYear$
  Dim DelDprRecCnt As Long
  Dim ThisTag$
  Dim YearCnt As Long
  Dim Y As Long
  Dim BigYear As Integer
  Dim LifeLeft As Double
  Dim NmlDprAmt As Double
  Dim CurrValue As Double
  Dim Dif As Double
  Dim Remain As String * 2
  Dim ThisTagNum$
  
  On Error GoTo ERRORSTUFF
  
  DelDprRecCnt = 1
  ThisYear = QPTrim$(fptxtYear.Text)
  OpenDprHistFile DPRHandle
  DprCnt = LOF(DPRHandle) / Len(DprRec)
  If DprCnt = 0 Then
    Close
    MsgBox "There are no depreciation records saved."
    Exit Sub
  End If
  
  ReDim DelDprRec(1 To DelDprRecCnt) As Long
  
  For x = 1 To DprCnt
    Get DPRHandle, x, DprRec
    If DprRec.SoSoftFlag = True Then 'there is already a sosoft reversal
    'has not been re-depreciated...this reversal needs to stop here
      If MsgBox("There is a depreciation file that has been reversed and has not been posted. This could cause unexpected depreciation results and is not recommended. Do you want to continue anyway?", vbYesNo) = vbNo Then
        Close
        Exit Sub 'no don't continue
      Else
        Exit For 'go ahead anyway
      End If
    End If
  Next x
  
  BigYear = 0
  'we only want to be able to reverse a depreciation for the most current year
  For x = 1 To DprCnt
    Get DPRHandle, x, DprRec
    If Val(DprRec.DprYear) > BigYear Then BigYear = CInt(DprRec.DprYear) 'find most recent year
    If QPTrim$(DprRec.DprYear) = ThisYear$ Then 'go ahead and save records for
    'the year entered
      DelDprRec(DelDprRecCnt) = x 'save this record
      DelDprRecCnt = DelDprRecCnt + 1 'count the valid records
      YearCnt = YearCnt + 1 'look thru records to validate the year entered on screen
      ReDim Preserve DelDprRec(1 To DelDprRecCnt) As Long 'update array count
    End If
  Next x
  
  If YearCnt = 0 Then 'no dates found in records that match date on screen
    MsgBox "No depreciation amounts to delete for the year entered."
    Close
    Exit Sub
  End If
  
  If CStr(BigYear) <> ThisYear Then 'the year on the screen is not the most
  'current depreciation year
    MsgBox "Only the most current depreciation can be reversed. The most current year is " + CStr(BigYear) + "."
    fptxtYear.SetFocus
    Close
    Exit Sub
  End If
  
  'at this point this reversal has been cleared
  OpenFAItemFile FAHandle
  NumOfFARecs = LOF(FAHandle) / Len(FAItemRec)
  frmFAShowPctComp.Label1 = "Deleting Depreciation"
  frmFAShowPctComp.cmdCancel.Visible = False 'don't want this process canceled
  frmFAShowPctComp.Show
  DoEvents
  
  For x = 1 To DelDprRecCnt - 1 '-1 because DelDprRecCnt gets one extra
  'loop while the count takes place
    Get DPRHandle, DelDprRec(x), DprRec 'retrieve a depreciation record
      ThisTagNum = QPTrim$(DprRec.ItemTag)
      For Y = 1 To NumOfFARecs 'check out each fixed asset item record
        Get FAHandle, Y, FAItemRec
        If QPTrim$(FAItemRec.ItemTag) = ThisTagNum Then 'This item matches the depreciation record
          FAItemRec.CURRVAL = FAItemRec.CURRVAL + DprRec.DprAmt 'reassign the current value by adding back the
          'amount it was depreciated
          FAItemRec.DEP2DATE = FAItemRec.DEP2DATE - DprRec.DprAmt 'reassign the depreciation to date amount
          'by subtracting the depreciation amount for this depreciation process
          If Not FAItemRec.LifeLeft + 1 > FAItemRec.ILIFE Then 'if an item has
          'a current depreciation amount less than the normal full year's depreciation
          'amount because this is the first time it's being depreciated under a
          'percentage then the life left will not be reduced...so if you add a one
          'to it it will actually end up having a life left more than the original life
            FAItemRec.LifeLeft = FAItemRec.LifeLeft + 1
          End If
          Put FAHandle, Y, FAItemRec 'save it
          Exit For
        End If
      Next Y
      If Y > NumOfFARecs Then 'this should never happen but if a tag number
      'showed up in the depreciation records saved for this year can't be
      'found in the FArecords then this msgbox pops up
        MsgBox "Item Number " + ThisTagNum + " was not updated."
      End If
      DprRec.SoSoftFlag = True 'important to set this flag so we know that
      'this process has taken place...this flag will be set to false when
      'this reversal is re-depreciated and posted
      DprRec.DprToDate = DprRec.DprToDate - DprRec.DprAmt 'reset depreciation record
      DprRec.BookTotal = DprRec.BookTotal + DprRec.DprAmt 'reset depreciation record
      DprRec.DprAmt = 0 'delete depreciation amount for this record
'      DprRec.DprYear = DprRec.DprYear - 1
      If Not DprRec.LifeLeft + 1 > DprRec.Life Then
        DprRec.LifeLeft = DprRec.LifeLeft + 1 'reset life left
      End If
      Put DPRHandle, DelDprRec(x), DprRec 'save to depreciation record
      frmFAShowPctComp.ShowPctComp x, DelDprRecCnt
      If frmFAShowPctComp.Out = True Then
        Close
        frmFAShowPctComp.Out = False
        Unload frmFAShowPctComp
        EnableCloseButton Me.hwnd, True
        Me.cmdExit.Enabled = True
        Me.cmdProcess.Enabled = True
        Exit Sub
      End If
  Next x
  
  
  Close DPRHandle
  Close FAHandle
  
  Unload frmFAShowPctComp
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Me.cmdProcess.Enabled = True
  MsgBox "Depreciation deletion is complete for year " + ThisYear + "."
  Call cmdExit_Click
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAInternalOnly", "cmdProcess_Click", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    ClearInUse (PWcnt)
    Terminate
    Unload Me
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
'    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
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
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      ClearInUse PWcnt
      MainLog ("FixedAsset.exe terminated via menu bar on frmFAInternalOnly.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim x As Long
  Dim DprRec As DprHistType
  Dim DPRHandle As Integer
  Dim DprCnt As Long
  Dim BigYear As Integer
  
  OpenDprHistFile DPRHandle
  DprCnt = LOF(DPRHandle) / Len(DprRec)
  If DprCnt = 0 Then
    Close
    MsgBox "There are no depreciation records saved."
    fptxtYear.Text = "NONE"
    Exit Sub
  End If
  
  BigYear = 0
  For x = 1 To DprCnt
    Get DPRHandle, x, DprRec
    If Val(DprRec.DprYear) > BigYear Then BigYear = CInt(DprRec.DprYear) 'find most recent year
  Next x
  
  Close DPRHandle
  fptxtYear.Text = CStr(BigYear)
End Sub
