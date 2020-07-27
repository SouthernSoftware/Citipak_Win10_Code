VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmEmergency 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emergency Contact Information"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   FillColor       =   &H00C0C0C0&
   Icon            =   "frmEmergency.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   5340
      Left            =   2208
      TabIndex        =   5
      Top             =   1726
      Width           =   7212
      _Version        =   196609
      _ExtentX        =   12721
      _ExtentY        =   9419
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
      FrameColor      =   -2147483627
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmEmergency.frx":08CA
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3510
         TabIndex        =   4
         Top             =   3450
         Width           =   2355
         _Version        =   196608
         _ExtentX        =   4154
         _ExtentY        =   714
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Text            =   ""
         Columns         =   0
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
         ScrollBarH      =   1
         DataFieldList   =   ""
         ColumnEdit      =   -1
         ColumnBound     =   -1
         Style           =   2
         MaxDrop         =   8
         ListWidth       =   -1
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
         EditAlignH      =   1
         EditAlignV      =   0
         ColDesigner     =   "frmEmergency.frx":08E6
      End
      Begin VB.CheckBox CheckNumber 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Sort by Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   396
         Left            =   3870
         TabIndex        =   3
         Top             =   2736
         Width           =   2340
      End
      Begin VB.CheckBox CheckName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Sort by Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   396
         Left            =   1125
         TabIndex        =   2
         Top             =   2736
         Value           =   1  'Checked
         Width           =   1980
      End
      Begin EditLib.fpText fptxtFirstEmpNo 
         Height          =   396
         Left            =   4080
         TabIndex        =   0
         Top             =   1392
         Width           =   1308
         _Version        =   196608
         _ExtentX        =   2307
         _ExtentY        =   698
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
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
         AlignTextH      =   1
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   0
         CaretOverWrite  =   3
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
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         CharValidationText=   "1 2 3 4 5 6 7 8 9 0 ,"
         MaxLength       =   255
         MultiLine       =   0   'False
         PasswordChar    =   ""
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText fptxtLastEmpNo 
         Height          =   396
         Left            =   4080
         TabIndex        =   1
         Top             =   2016
         Width           =   1308
         _Version        =   196608
         _ExtentX        =   2307
         _ExtentY        =   698
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
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
         AlignTextH      =   1
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   0
         CaretOverWrite  =   3
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
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         CharValidationText=   "1 2 3 4 5 6 7 8 9 0 ,"
         MaxLength       =   255
         MultiLine       =   0   'False
         PasswordChar    =   ""
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   690
         Left            =   4176
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Press F10 to generate an employee emergency information report."
         Top             =   4224
         Width           =   1890
         _Version        =   131072
         _ExtentX        =   3334
         _ExtentY        =   1217
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
         ButtonDesigner  =   "frmEmergency.frx":0BDD
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
         Height          =   690
         Left            =   1296
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Press ESC to exit this screen."
         Top             =   4224
         Width           =   1890
         _Version        =   131072
         _ExtentX        =   3334
         _ExtentY        =   1217
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
         ButtonDesigner  =   "frmEmergency.frx":0DF4
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   588
         Left            =   768
         Top             =   2640
         Width           =   5820
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Print Option:"
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
         Left            =   1392
         TabIndex        =   9
         Top             =   3540
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Emergency Contact Information"
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
         Height          =   492
         Left            =   1392
         TabIndex        =   8
         Top             =   528
         Width           =   4524
      End
      Begin VB.Label Label2 
         BackColor       =   &H00D0D0D0&
         Caption         =   "First Employee No:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1776
         TabIndex        =   7
         Top             =   1536
         Width           =   2124
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Last Employee No:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1536
         TabIndex        =   6
         Top             =   2160
         Width           =   2268
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   684
         Left            =   1248
         Top             =   384
         Width           =   4812
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   5628
      Left            =   2100
      Top             =   1618
      Width           =   7452
   End
End
Attribute VB_Name = "frmEmergency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub CheckName_Click()
  If CheckName.Value = 1 Then CheckNumber.Value = 0
End Sub

Private Sub CheckNumber_Click()
  If CheckNumber.Value = 1 Then CheckName.Value = 0
End Sub

Private Sub cmdEscape_Click()
'this report can come from either the Employee Maintenance
'Menu or the Reports Processing Menu...the "roOn" tells
'the program which menu to return to upon exit
  If frmReportsProcessing.Selection = roOn Then
    frmReportsProcessing.Show
    DoEvents
    Unload frmEmergency
  Else
    frmEmployeeMaintMenu.Show
    DoEvents
    Unload frmEmergency
  End If
End Sub

Private Sub PrintGraphics()
  
  Dim RptName As String, EmpIdxLNameHandle As Integer
  Dim EmpIdxNNameHandle As Integer, ThisSortName() As Integer
  Dim FldFlag As String, DescFlag As String
  Dim IdxNumOfRecs As Integer, UnitHandle As Integer
  Dim Emp2RecLen As Integer, UnitFileRec As UnitFileRecType
  Dim DataFileSize As Long, cnt As Integer
  Dim RptHandle As Integer, D2Handle As Integer
  Dim RptTitle As String, D2Name As Long
  Dim EmpIdxLNameRec As NameSortIdxType
  Dim EmpIdxNNameRec As NumbSortIdxType
  Dim HDate$, ThisSortNumber() As Integer
  Dim FF As String, x As Integer, City As String
  Dim EmpData2FileHandle As Integer
  Dim EmpData2FileRec As EmpData2Type
  Dim EmpIdxLNameCnt As Integer
  Dim EmpIdxNNameCnt As Integer
  Dim LineCnt As Integer, MaxLines As Integer
  ReDim Emp2Data(1) As EmpData2Type
  Dim FandLName As String * 20
  Dim Today As String * 11
  Dim ValidEmpCnt As Integer
  Dim BDate$
  Dim dlm$
  
  dlm$ = "~"
'  Date$ = FormatDateTime(Date, vbShortDate)
  Today = Date '$
  
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, UnitFileRec
  City = QPTrim$(UnitFileRec.UFEMPR)
  Close UnitHandle
  RptName$ = "PRRPTS\EMERGENCYG.RPT"
 
'  RptTitle$ = "Employee Information Listing in Alphabetic Order"
  OpenEmpIdxLNameFile EmpIdxLNameHandle
  EmpIdxLNameCnt = LOF(EmpIdxLNameHandle) / 2
  If EmpIdxLNameCnt = 0 Then
    MsgBox "No records on file"
    Close
    MainLog ("Emergency Report exited with No records on file.")
    Exit Sub
  End If
  FrmShowPctComp.Label1 = "Employee Emergency Information Report"
  ReDim ThisSortName(EmpIdxLNameCnt)
  For x = 1 To EmpIdxLNameCnt 'load array with employee data
  'sorted by last name
     Get EmpIdxLNameHandle, x, ThisSortName(x)
  Next x
  Close EmpIdxLNameHandle
  IdxNumOfRecs = EmpIdxLNameCnt
  
  OpenEmpIdxNNameFile EmpIdxNNameHandle
  EmpIdxNNameCnt = LOF(EmpIdxNNameHandle) / 2
  If EmpIdxNNameCnt = 0 Then
    MsgBox "No records on file"
    Close
    MainLog ("Emergency Report exited with No records on file.")
    Exit Sub
  End If
  ReDim ThisSortNumber(EmpIdxNNameCnt)

  For x = 1 To EmpIdxNNameCnt 'load array with employee data
  'sorted by last name
     Get EmpIdxNNameHandle, x, ThisSortNumber(x)
  Next x
  Close EmpIdxNNameHandle
  
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  
  RptHandle = FreeFile
  On Error GoTo ErrorHandler
  Open RptName$ For Output As RptHandle
  
  OpenEmpData2File EmpData2FileHandle
  For cnt = 1 To IdxNumOfRecs
    If CheckNumber.Value = 1 Then
      Get EmpData2FileHandle, ThisSortNumber(cnt), EmpData2FileRec
    Else
      Get EmpData2FileHandle, ThisSortName(cnt), EmpData2FileRec
    End If
    If Val(EmpData2FileRec.EmpNo) < Val(fptxtFirstEmpNo.Text) Or Val(EmpData2FileRec.EmpNo) > Val(fptxtLastEmpNo.Text) Then GoTo SkipEm
    FandLName = QPTrim$(EmpData2FileRec.EmpLName) + "  " + QPTrim$(EmpData2FileRec.EmpFName)
    If Not EmpData2FileRec.Deleted Then 'And EmpData2FileRec.EMPTDATE = 0 Then
      ValidEmpCnt = ValidEmpCnt + 1
      If Len(QPTrim$(FandLName)) = 0 Then GoTo SkipEm
      '                   0                  1                              2                       3                                    4                                                5
      Print #RptHandle, City; dlm; QPTrim$(EmpData2FileRec.EmpNo); dlm; FandLName; dlm; QPTrim$(EmpData2FileRec.EmpAddr1); dlm; QPTrim$(EmpData2FileRec.EMPADDR2); dlm; QPTrim$(EmpData2FileRec.EmpCity); dlm;
      '                                 6                                     7                                        8                                          9
      Print #RptHandle, QPTrim$(EmpData2FileRec.EmpState); dlm; QPTrim$(EmpData2FileRec.EmpZip); dlm; QPTrim$(EmpData2FileRec.HomePhone); dlm; QPTrim$(EmpData2FileRec.EmrgncyCntctName); dlm;
      '                                  10                                                 11
      Print #RptHandle, QPTrim$(EmpData2FileRec.EmrgncyCntctPhnNum); dlm; QPTrim$(EmpData2FileRec.EmrgncyCntctRelation)
SkipEm:
    End If
    FrmShowPctComp.ShowPctComp cnt, IdxNumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Me.cmdEscape.Enabled = True
      Unload FrmShowPctComp
      Exit Sub
    End If
  
  Next
    
  Close EmpData2FileHandle
  Close RptHandle
  arEmergency.Show
  
  frmLoadingRpt.Show
  EnableCloseButton Me.hwnd, True
  Me.cmdEscape.Enabled = True
  MainLog ("Employee List report processed.")
Exit Sub

EndTrans:
ErrorHandler:
  Close
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload FrmShowPctComp
  MsgBox "ERROR: If this problem persists please consult Southern Software."

End Sub

Private Sub cmdProcess_Click()
  If fpcomboPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
    Exit Sub
  ElseIf fpcomboPrintOpt.Text = "Text" Then
    Call PrintText
  Else
    Exit Sub
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyReturn:
      Call cmdProcess_Click
      KeyCode = 0
    Case vbKeyDown:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdEscape_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Dim EmpData1Handle As Integer, EmpIdxLNameHandle As Integer
  Dim IdxRecPointer As Integer, NumOfRecs As Integer
  Dim EmpData1Rec As EmpData1Type

  OpenEmpData1File EmpData1Handle
  OpenEmpIdxNNameFile EmpIdxLNameHandle
  NumOfRecs = LOF(EmpIdxLNameHandle) / 2
  If NumOfRecs = 0 Then
    MsgBox "No records on file."
    Close
    Exit Sub
  End If
  Get #EmpIdxLNameHandle, 1, IdxRecPointer
  Get #EmpData1Handle, IdxRecPointer, EmpData1Rec
  fptxtFirstEmpNo.Text = Val(EmpData1Rec.EmpNo)
   
  Get #EmpIdxLNameHandle, NumOfRecs, IdxRecPointer
  Get #EmpData1Handle, IdxRecPointer, EmpData1Rec
  fptxtLastEmpNo.Text = Val(EmpData1Rec.EmpNo)
  
  Close EmpIdxLNameHandle, EmpData1Handle
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  fpcomboPrintOpt.AddItem "Graphical"
  fpcomboPrintOpt.AddItem "Text"
  fpcomboPrintOpt.Text = "Graphical"
  Me.HelpContextID = hlpPrintEmpEmergency
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
      MainLog ("Payroll.exe terminated via Menu Bar on frmEmergency.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub fpcomboPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboPrintOpt.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboPrintOpt.ListIndex = -1
  End If
  If fpcomboPrintOpt.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      cmdEscape.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub PrintText()

  Dim MaxLines As Integer, LineCnt As Integer
  Dim EmpRecSize As Long
  Dim IdxRecLen As Integer, IdxFileSize&
  Dim NumOfRecs As Long, cnt As Integer
  Dim UnitFileName As Integer, RptTitle$, UTemp$
  Dim RptName$, RptHandle As Integer
  Dim DHandle As Integer, Today$
  Dim RecNo As Long, Page As Integer, CrLf$
  Dim Dash As String * 80 ', OKFlag As Boolean
  Dim City$, UnitHandle As Integer, x As Integer
  Dim UnitFileRec As UnitFileRecType
  Dim EmpIdxLNameHandle As Integer, EmpHistoryRpt$
  Dim EmpIdxNNameHandle As Integer
  Dim Emp2Rec As EmpData2Type, FF$
  Dim TotNumOfChks As Integer
  Dim TotAmtOfChks As Double, FandLName$
  ReDim TransHRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim Pg(1) As String * 3
  
  FF$ = Chr(12)

'  Date$ = FormatDateTime(Date, vbShortDate)
  Today = Date '$
  
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, UnitFileRec
  City = QPTrim$(UnitFileRec.UFEMPR)
  Close UnitHandle

  MaxLines = 52
  LineCnt = 0
  Dash = String$(78, "-") + CrLf$

'  EmpRecSize = Len(Emp2Rec)
  IdxRecLen = 2
  IdxFileSize& = FileSize(PRData + EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen

  OpenEmpIdxLNameFile EmpIdxLNameHandle
  
  NumOfRecs = LOF(EmpIdxLNameHandle) / 2
  If NumOfRecs = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If
  
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  FrmShowPctComp.Label1 = "Employee Checks Issued Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  If CheckNumber.Value = 1 Then
    OpenEmpIdxNNameFile EmpIdxNNameHandle
    For x = 1 To NumOfRecs
      Get #EmpIdxNNameHandle, x, IdxBuff(x)
    Next x
    Close EmpIdxNNameHandle
  Else
    For x = 1 To NumOfRecs
      Get #EmpIdxLNameHandle, x, IdxBuff(x)
    Next x
    Close EmpIdxLNameHandle
  End If
  
  RptTitle$ = "Employee Emergency Report"
  RptName$ = "PRRPTS\EMERGENCY.RPT"
  RptHandle = FreeFile
  Open RptName$ For Output As RptHandle
  
  OpenEmpData2File DHandle
  GoSub PrintHeader
  For cnt = 1 To NumOfRecs
    Get DHandle, IdxBuff(cnt), Emp2Rec
    If Val(Emp2Rec.EmpNo) < Val(fptxtFirstEmpNo.Text) Or Val(Emp2Rec.EmpNo) > Val(fptxtLastEmpNo.Text) Then GoTo SkipEm
    FandLName = QPTrim$(Emp2Rec.EmpLName) + ",  " + QPTrim$(Emp2Rec.EmpFName)
    If Not Emp2Rec.Deleted Then 'And EmpData2FileRec.EMPTDATE = 0 Then
      If Len(QPTrim$(FandLName)) = 0 Then GoTo SkipEm
      Print #RptHandle, Tab(2); QPTrim$(Emp2Rec.EmpNo); Tab(14); FandLName
      Print #RptHandle, Tab(14); QPTrim$(Emp2Rec.EmpAddr1)
      Print #RptHandle, Tab(14); QPTrim$(Emp2Rec.EMPADDR2)
      Print #RptHandle, Tab(14); QPTrim$(Emp2Rec.EmpCity); Tab(41); QPTrim$(Emp2Rec.EmpState); Tab(45); QPTrim$(Emp2Rec.EmpZip)
      Print #RptHandle,
      Print #RptHandle, Tab(14); "Employee Home Phone: "; Tab(47); QPTrim$(Emp2Rec.HomePhone)
      Print #RptHandle, Tab(14); "Emergency Contact Name: "; Tab(47); QPTrim$(Emp2Rec.EmrgncyCntctName)
      Print #RptHandle, Tab(14); "Contact Relationship: "; Tab(47); QPTrim$(Emp2Rec.EmrgncyCntctRelation)
      Print #RptHandle, Tab(14); "Emergency Phone Number: "; Tab(47); QPTrim$(Emp2Rec.EmrgncyCntctPhnNum)
      LineCnt = LineCnt + 9
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintHeader
      Else
        Print #RptHandle, Dash$
        LineCnt = LineCnt + 1
      End If
SkipEm:
    End If
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Me.cmdEscape.Enabled = True
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next
  
  Print #RptHandle, FF$
'  RPTSetupPRN 123, RHandle '7/24 revised 8/15/02
  Close DHandle
  Close RptHandle
  ViewPrint RptName$, RptTitle$
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  MainLog ("Checks Issued Report processed.")
  Exit Sub

PrintHeader:
  Page = Page + 1
  RSet Pg(1) = Page
  UTemp$ = Space$(80)
  LSet UTemp$ = City
  Mid$(UTemp$, 71) = "Page:" + Pg(1) + CrLf$
  Print #RptHandle, UTemp$
  Print #RptHandle, "Emergency Employee Data" + CrLf$
  Print #RptHandle, Today
  Print #RptHandle,
  Print #RptHandle, "Employee"; Tab(12); "Employee Name and Address"
  Print #RptHandle, " Number"; Tab(12); "Emergency Information"
  Print #RptHandle, Dash
  LineCnt = 6
  Return


End Sub

