VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmWOPrintBook 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Work Orders By Book"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   Icon            =   "frmWOPrintBook.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   348
      Left            =   6276
      TabIndex        =   3
      Top             =   4656
      Width           =   2100
      _Version        =   196608
      _ExtentX        =   3704
      _ExtentY        =   614
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   1
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   1
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483627
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ScrollHScale    =   2
      ScrollHInc      =   0
      ColsFrozen      =   0
      ScrollBarV      =   1
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   0
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   2
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   3
      DataFieldList   =   ""
      ColumnEdit      =   0
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   3504
      EditHeight      =   -1
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   -2147483637
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   -2147483637
      ExtendRow       =   0
      ListPosition    =   0
      ButtonThreeDAppearance=   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      AutoSearchFill  =   0   'False
      AutoSearchFillDelay=   500
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmWOPrintBook.frx":08CA
   End
   Begin EditLib.fpText fptxtCopies 
      Height          =   348
      Left            =   6276
      TabIndex        =   2
      Top             =   4140
      Width           =   660
      _Version        =   196608
      _ExtentX        =   1164
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
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
      ThreeDOutsideStyle=   2
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
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "0123456789"
      MaxLength       =   3
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
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
      Height          =   492
      Left            =   10080
      TabIndex        =   5
      Top             =   7464
      Width           =   1332
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F10 &Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   8400
      TabIndex        =   4
      Top             =   7464
      Width           =   1332
   End
   Begin EditLib.fpText fptxtRoute2 
      Height          =   348
      Left            =   6276
      TabIndex        =   1
      Top             =   3612
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
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
      ThreeDOutsideStyle=   2
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
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "0123456789"
      MaxLength       =   2
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
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
   Begin EditLib.fpText fptxtRoute1 
      Height          =   348
      Left            =   6276
      TabIndex        =   0
      Top             =   3096
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
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
      ThreeDOutsideStyle=   2
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
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "0123456789"
      MaxLength       =   2
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   9
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
            TextSave        =   "12:33 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "6/23/2005"
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
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Report Type:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   3756
      TabIndex        =   11
      Top             =   4704
      Width           =   2388
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Copies:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Index           =   0
      Left            =   4572
      TabIndex        =   10
      Top             =   4188
      Width           =   1572
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2724
      Left            =   2700
      Top             =   2688
      Width           =   6780
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   708
      Left            =   3222
      Top             =   1200
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Work Orders By Book"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   0
      Left            =   3654
      TabIndex        =   8
      Top             =   1368
      Width           =   5004
   End
   Begin VB.Label LabelB1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From Book:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   4668
      TabIndex        =   7
      Top             =   3156
      Width           =   1476
   End
   Begin VB.Label LabelB2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To Book:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   4764
      TabIndex        =   6
      Top             =   3672
      Width           =   1380
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   828
      Left            =   3222
      Top             =   1080
      Width           =   5772
   End
End
Attribute VB_Name = "frmWOPrintBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider

Private Sub cmdPrint_Click()
  PrintWOsBook
End Sub

Private Sub fpcboRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptType.ListDown = True
  End If
  If fpcboRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      cmdPrint.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fptxtCopies.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fptxtCopies_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtRoute1_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtRoute1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtRoute2.SetFocus
  End If
End Sub

Private Sub fptxtRoute2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtRoute2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtCopies.SetFocus
  End If
End Sub
Private Sub fptxtCopies_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboRptType.SetFocus
  End If
End Sub
Private Sub cmdExit_Click()
  frmUBWorkOrderMenu.Show
  Unload frmWOPrintBook
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via WOPrintBook by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
'    Case vbKeyDown, vbKeyReturn:
'      SendKeys "{Tab}"
'      KeyCode = 0
'    Case vbKeyUp:
'      SendKeys "+{Tab}"
'      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  fptxtRoute1 = 0
  fptxtRoute2 = 99
  fptxtCopies = 1
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
  Me.HelpContextID = hlpPrintWorkOrdersBy
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub

Private Sub PrintWOsBook()
  Dim UBCustRecLen As Integer, WorkOrderRecLen As Integer
  Dim Dash As String, PrintSingleFlag As Boolean, Copies As Integer
  Dim ReportFile As String, RptHandle As Integer, IdxName As String
  Dim IdxRecLen As Integer, IdxFileSize As Long, IdxNumOfRecs As Long
  Dim NumOfRecs As Long, Handle As Integer, UBCustF As Integer
  Dim UBWOFile As Integer, lcnt As Long, Book As Integer, cnt As Long
  Dim BegRoute As Integer, EndRoute As Integer, Acct As Long
  Dim Header As String, CopyCnt As Integer, MtrCnt As Integer
  Dim Rem1 As String, Rem2 As String, Rem3 As String, Rem4 As String
  Dim Rem5 As String, Rem6 As String, ToPrint As String
  Dim graphicflag As Boolean
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  
  ReDim WorkOrderRec(1) As WorkOrderRecType
  WorkOrderRecLen = Len(WorkOrderRec(1))
  Rem1$ = ""
  Rem2$ = ""
  Rem3$ = ""
  Rem4$ = ""
  Rem5$ = ""
  Rem6$ = ""
  If fpcboRptType.ListIndex = 0 Then
    graphicflag = True
  Else
    graphicflag = False
  End If
  DeActivateControls Me
  FrmShowPctComp.Label1 = "Creating Work Order"
  FrmShowPctComp.Show , Me

  If graphicflag = True Then
    Dash$ = String$(83, "_")
  Else
    Dash$ = String$(79, "_")
  End If
  ToPrint$ = ""
  FF$ = Chr$(12)
  BegRoute = Val(fptxtRoute1)
  EndRoute = Val(fptxtRoute2)

skipthis:
Copies = Val(fptxtCopies)
  If Copies < 1 Then
    Copies = 1
  End If

  'Open Report File
  ReportFile$ = UBPath$ + "WORKORDR.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  ' Location Order ********************************************************
  
  IdxName$ = UBPath$ + "UBCUSTBK.IDX"
  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = FileSize&(IdxName$)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen

  ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
  'FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
  NumOfRecs = IdxNumOfRecs
  Handle = FreeFile
  Open IdxName$ For Random Shared As Handle Len = IdxRecLen
  For cnt& = 1 To IdxNumOfRecs
    Get #Handle, cnt&, IdxBuff(cnt&)
  Next
  Close Handle

  UBCustF = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCustF Len = UBCustRecLen

  UBWOFile = FreeFile
  Open UBPath$ + "UBWRKORD.DAT" For Random Shared As UBWOFile Len = WorkOrderRecLen

    cnt& = 1
    'ShowProcessingScrn "Processing Work Orders"
    For lcnt& = 1 To IdxNumOfRecs
      FrmShowPctComp.ShowPctComp lcnt, IdxNumOfRecs
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        ActivateControls Me
        GoTo ExitHere
      End If

      Get #UBCustF, IdxBuff(lcnt&).RecNum, UBCustRec(1)
      If UBCustRec(1).DelFlag <> 0 Then
        GoTo DelSkip
      End If
      Book = Val(UBCustRec(1).Book)
      If Book >= BegRoute And Book <= EndRoute Then
        If UBCustRec(1).WOLastTrans > 0 Then
          Get #UBWOFile, UBCustRec(1).WOLastTrans, WorkOrderRec(1)
          If WorkOrderRec(1).CompletedDate <= 0 Then
            Acct& = IdxBuff(lcnt&).RecNum
            GoSub PrintThemOne
          End If
        End If
      End If
      'ShowPctComp lcnt&, IdxNumOfRecs
DelSkip:
    Next

  'PRINT #RptHandle, FF$

  Close
  Erase UBCustRec, WorkOrderRec, IdxBuff

  Header$ = "Customer Work Orders "
  'PrintRptFile Header$, ReportFile$, LPTPort, RetCode, EntryPoint
  If graphicflag = False Then
    ViewPrint ReportFile$, Header$
    ActivateControls Me
  Else
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmWOPrintBook
    ARptWorkOrder.GetName ReportFile$
    ARptWorkOrder.startrpt
  End If
ExitHere:
  Exit Sub

PrintThemOne:
  If Len(QPTrim(WorkOrderRec(1).RepliesText.Text(1))) > 0 Then
    Rem1$ = QPTrim(WorkOrderRec(1).RepliesText.Text(1))
  Else
    Rem1$ = Dash$
  End If
  If Len(QPTrim(WorkOrderRec(1).RepliesText.Text(2))) > 0 Then
    Rem2$ = QPTrim(WorkOrderRec(1).RepliesText.Text(2))
  Else
    Rem2$ = Dash$
  End If
  If Len(QPTrim(WorkOrderRec(1).RepliesText.Text(3))) > 0 Then
    Rem3$ = QPTrim(WorkOrderRec(1).RepliesText.Text(3))
  Else
    Rem3$ = Dash$
  End If
  If Len(QPTrim(WorkOrderRec(1).RepliesText.Text(4))) > 0 Then
    Rem4$ = QPTrim(WorkOrderRec(1).RepliesText.Text(4))
  Else
    Rem4$ = Dash$
  End If
  If Len(QPTrim(WorkOrderRec(1).RepliesText.Text(5))) > 0 Then
    Rem5$ = QPTrim(WorkOrderRec(1).RepliesText.Text(5))
  Else
    Rem5$ = Dash$
  End If

  If Len(QPTrim(WorkOrderRec(1).RepliesText.Text(6))) > 0 Then
    Rem6$ = QPTrim(WorkOrderRec(1).RepliesText.Text(6))
  Else
    Rem6$ = "BY: ______________________________   DATE: ____________________"
  End If
 
  For CopyCnt = 1 To Copies
  If graphicflag = False Then
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, "Printed:"; Now
    Print #RptHandle, " "
    Print #RptHandle, Tab(14); "W O R K   O R D E R   :   U T I L I T Y   D E P T ."
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, "    Work Order#: "; Using("######", UBCustRec(1).WOLastTrans); Tab(30); "Date Issued: "; Num2Date$(WorkOrderRec(1).ENTRYDATE)
    Print #RptHandle, "      Location#: "; UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; Tab(30); "Complete By: "; Num2Date$(WorkOrderRec(1).CompleteByDate)
    Print #RptHandle, "       Account#: "; Acct&; Tab(30); "  Completed: "; Num2Date$(WorkOrderRec(1).CompletedDate)
    Print #RptHandle, "  Customer Name: "; UBCustRec(1).CustName
    Print #RptHandle, "Service Address: "; UBCustRec(1).ServAddr
    Print #RptHandle, String$(79, "-")
    Print #RptHandle, "Instruction or Description of Work Needed"
    Print #RptHandle, String$(79, "-")
    Print #RptHandle, WorkOrderRec(1).OrdersText.Text(1)
    Print #RptHandle, WorkOrderRec(1).OrdersText.Text(2)
    Print #RptHandle, WorkOrderRec(1).OrdersText.Text(3)
    Print #RptHandle, WorkOrderRec(1).OrdersText.Text(4)
    Print #RptHandle, WorkOrderRec(1).OrdersText.Text(5)
    Print #RptHandle, WorkOrderRec(1).OrdersText.Text(6)
    Print #RptHandle, " "
    Print #RptHandle, String$(79, "-")
    Print #RptHandle, "Remarks Noted by Worker"
    Print #RptHandle, String$(79, "-")
    Print #RptHandle, Rem1$
    Print #RptHandle, " "
    Print #RptHandle, Rem2$
    Print #RptHandle, " "
    Print #RptHandle, Rem3$
    Print #RptHandle, " "
    Print #RptHandle, Rem4$
    Print #RptHandle, " "
    Print #RptHandle, Rem5$
    Print #RptHandle, " "
    Print #RptHandle, Rem6$
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, "Meter Numbers:"

    For MtrCnt = 1 To 7
      If Len(QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)) > 0 Then
        Print #RptHandle, QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
      End If
    Next
    Print #RptHandle, FF$;
  Else
    ToPrint$ = Num2Date$(WorkOrderRec(1).ENTRYDATE) + "~"
    ToPrint$ = ToPrint$ + Using("######", UBCustRec(1).WOLastTrans) + "~"
    ToPrint$ = ToPrint$ + UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB + "~"
    ToPrint$ = ToPrint$ + Str(Acct&) + "~"
    ToPrint$ = ToPrint$ + UBCustRec(1).CustName + "~"
    ToPrint$ = ToPrint$ + UBCustRec(1).ServAddr + "~"
    ToPrint$ = ToPrint$ + WorkOrderRec(1).OrdersText.Text(1) + "~"
    ToPrint$ = ToPrint$ + WorkOrderRec(1).OrdersText.Text(2) + "~"
    ToPrint$ = ToPrint$ + WorkOrderRec(1).OrdersText.Text(3) + "~"
    ToPrint$ = ToPrint$ + WorkOrderRec(1).OrdersText.Text(4) + "~"
    ToPrint$ = ToPrint$ + WorkOrderRec(1).OrdersText.Text(5) + "~"
    ToPrint$ = ToPrint$ + WorkOrderRec(1).OrdersText.Text(6) + "~"
    ToPrint$ = ToPrint$ + Rem1$ + "~"
    ToPrint$ = ToPrint$ + Rem2$ + "~"
    ToPrint$ = ToPrint$ + Rem3$ + "~"
    ToPrint$ = ToPrint$ + Rem4$ + "~"
    ToPrint$ = ToPrint$ + Rem5$ + "~"
    ToPrint$ = ToPrint$ + Rem6$

    For MtrCnt = 1 To 7
      If Len(QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)) > 0 Then
        ToPrint$ = ToPrint$ + "~" + QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
      Else
        ToPrint$ = ToPrint$ + "~ "
      End If
    Next
    ToPrint$ = ToPrint$ + "~" + Num2Date$(WorkOrderRec(1).CompleteByDate) + "~"
    ToPrint$ = ToPrint$ + Num2Date$(WorkOrderRec(1).CompletedDate)

    Print #RptHandle, ToPrint$
    ToPrint$ = ""
  End If
  Next CopyCnt
  Return
End Sub

