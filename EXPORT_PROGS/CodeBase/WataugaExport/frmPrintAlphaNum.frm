VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmPrintAlphaNum 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Listing Report"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmPrintAlphaNum.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   4956
      Left            =   2244
      TabIndex        =   5
      Top             =   1942
      Width           =   7164
      _Version        =   196609
      _ExtentX        =   12636
      _ExtentY        =   8742
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.24
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
      Picture         =   "frmPrintAlphaNum.frx":08CA
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3360
         TabIndex        =   4
         Top             =   3165
         Width           =   2370
         _Version        =   196608
         _ExtentX        =   4180
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
         ColDesigner     =   "frmPrintAlphaNum.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbEmpType 
         Height          =   405
         Left            =   3360
         TabIndex        =   3
         Top             =   2640
         Width           =   2370
         _Version        =   196608
         _ExtentX        =   4180
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
         ColDesigner     =   "frmPrintAlphaNum.frx":0BDD
      End
      Begin LpLib.fpCombo fpcmbEmpStatus 
         Height          =   405
         Left            =   3360
         TabIndex        =   2
         Top             =   2115
         Width           =   2370
         _Version        =   196608
         _ExtentX        =   4180
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
         ColDesigner     =   "frmPrintAlphaNum.frx":0ED4
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
         Left            =   3915
         TabIndex        =   1
         Top             =   1392
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
         Left            =   1200
         TabIndex        =   0
         Top             =   1392
         Value           =   1  'Checked
         Width           =   1980
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   690
         Left            =   4176
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Press F10 to generate the employee list desired."
         Top             =   3888
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
         ButtonDesigner  =   "frmPrintAlphaNum.frx":11CB
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
         Height          =   690
         Left            =   1296
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Press ESC to exit this screen."
         Top             =   3888
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
         ButtonDesigner  =   "frmPrintAlphaNum.frx":13AA
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Employee Status:"
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
         Left            =   1248
         TabIndex        =   9
         Top             =   2208
         Width           =   1884
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Employee Type:"
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
         Left            =   1296
         TabIndex        =   8
         Top             =   2736
         Width           =   1884
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   588
         Left            =   720
         Top             =   1296
         Width           =   6012
      End
      Begin VB.Label Label4 
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
         Left            =   1680
         TabIndex        =   7
         Top             =   3264
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   " Employee Listing Report"
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
         Left            =   1632
         TabIndex        =   6
         Top             =   528
         Width           =   4044
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   684
         Left            =   1440
         Top             =   384
         Width           =   4428
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   5244
      Left            =   2112
      Top             =   1810
      Width           =   7452
   End
End
Attribute VB_Name = "frmPrintAlphaNum"
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

Private Sub cmdProcess_Click()
  InFileNames(1) = "PRDATA\PRUNIT.DAT"
  InFileNames(2) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 2) = False Then
    Close
    Exit Sub
  End If

  If fpcomboPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
    Exit Sub
  ElseIf fpcomboPrintOpt.Text = "Text" Then
    Call PrintText
  Else
    Exit Sub
  End If
End Sub

Private Sub cmdEscape_Click()
'this report can come from either the Employee Maintenance
'Menu or the Reports Processing Menu...the "roOn" tells
'the program which menu to return to upon exit
  If frmReportsProcessing.Selection = roOn Then
    frmReportsProcessing.Show
    DoEvents
    Unload frmPrintAlphaNum
  Else
    frmEmployeeMaintMenu.Show
    DoEvents
    Unload frmPrintAlphaNum
  End If
End Sub

Private Sub PrintGraphics()
  
  Dim RptName As String, EmpIdxLNameHandle As Integer
  Dim EmpIdxNNameHandle As Integer, ThisSort() As Integer
  Dim FldFlag As String, DescFlag As String
  Dim IdxNumOfRecs As Integer, UnitHandle As Integer
  Dim Emp2RecLen As Integer, UnitFileRec As UnitFileRecType
  Dim DataFileSize As Long, cnt As Integer
  Dim RptHandle As Integer, D2Handle As Integer
  Dim RptTitle As String, D2Name As Long
  Dim EmpIdxLNameRec As NameSortIdxType
  Dim EmpIdxNNameRec As NumbSortIdxType
  Dim HDate$
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
  Dim ThisCnt As Integer
  
  On Error GoTo ErrorHandler
  
  dlm$ = "~"
  Today = Date '$
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, UnitFileRec
  City = QPTrim$(UnitFileRec.UFEMPR)
  Close UnitHandle
  RptName$ = "PRRPTS\EMPrintEmpListG.RPT"
  If CheckName.Value = 1 Then
    RptTitle$ = "Employee Information Listing in Alphabetic Order"
    OpenEmpIdxLNameFile EmpIdxLNameHandle
    EmpIdxLNameCnt = LOF(EmpIdxLNameHandle) / 2
    If EmpIdxLNameCnt = 0 Then
      MsgBox "No records on file"
      Close
      MainLog ("Employee List Report exited with No records on file.")
      Exit Sub
    End If
    FrmShowPctComp.Label1 = "Alphabetic List of Employees Report"
    ReDim ThisSort(EmpIdxLNameCnt)
    For x = 1 To EmpIdxLNameCnt 'load array with employee data
    'sorted by last name
       Get EmpIdxLNameHandle, x, ThisSort(x)
    Next x
    Close EmpIdxLNameHandle
    IdxNumOfRecs = EmpIdxLNameCnt
  End If
  If CheckNumber.Value = 1 Then
    RptTitle$ = "Employee Information Listing in Numeric Order"
    OpenEmpIdxNNameFile EmpIdxNNameHandle
    EmpIdxNNameCnt = LOF(EmpIdxNNameHandle) / 2 'Len(EmpIdxNNameRec)
    If EmpIdxNNameCnt = 0 Then
      MsgBox "No records on file."
      Close
      Exit Sub
    End If
    FrmShowPctComp.Label1 = "List of Employees by Employee Number Report"
    ReDim ThisSort(EmpIdxNNameCnt)
    For x = 1 To EmpIdxNNameCnt 'load array sorted by employee
    'number
       Get EmpIdxNNameHandle, x, ThisSort(x)
    Next x
    Close EmpIdxNNameHandle
    IdxNumOfRecs = EmpIdxNNameCnt
  End If

  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  
  RptHandle = FreeFile
  Open RptName$ For Output As RptHandle
  
  OpenEmpData2File EmpData2FileHandle
  
  For cnt = 1 To IdxNumOfRecs
    Get EmpData2FileHandle, ThisSort(cnt), EmpData2FileRec
    BDate = MakeRegDate(EmpData2FileRec.EMPBDAY)
    If BDate = "12/31/1979" Then BDate = "NO RECORD"
    HDate = MakeRegDate(EmpData2FileRec.EMPHDATE)
    If HDate = "12/31/1979" Then HDate = "NO RECORD"
    FandLName = QPTrim$(EmpData2FileRec.EmpLName) + "  " + QPTrim$(EmpData2FileRec.EmpFName)
    If EmpData2FileRec.EMPPRATE < 0 Then
       EmpData2FileRec.EMPPRATE = 0
    End If
    'Filter
    
    If fpcmbEmpType.Text = "ALL" Then
      GoTo AllsGood
    ElseIf fpcmbEmpType.Text = "Full-Time" Then
      If QPTrim$(EmpData2FileRec.EMPSTATS) <> "Full-Time" Then
        GoTo SkipEm
      End If
    ElseIf fpcmbEmpType.Text = "Part-Time" Then
      If QPTrim$(EmpData2FileRec.EMPSTATS) <> "Part-Time" Then
        GoTo SkipEm
      End If
    ElseIf fpcmbEmpType.Text = "Seasonal" Then
      If QPTrim$(EmpData2FileRec.EMPSTATS) <> "Seasonal" Then
        GoTo SkipEm
      End If
    ElseIf fpcmbEmpType.Text = "Temporary" Then
      If QPTrim$(EmpData2FileRec.EMPSTATS) <> "Temporary" Then
        GoTo SkipEm
      End If
    End If
    
AllsGood:
    If fpcmbEmpStatus.Text = "ALL" Then
      GoTo AllsGoodAgain
    ElseIf fpcmbEmpStatus.Text = "Active" Then
      If EmpData2FileRec.EMPTDATE <> 0 Then
        GoTo SkipEm
      End If
    ElseIf fpcmbEmpStatus.Text = "Terminated" Then
      If EmpData2FileRec.EMPTDATE = 0 Then
        GoTo SkipEm
      End If
    End If
    
AllsGoodAgain:
    If Not EmpData2FileRec.Deleted Then 'And EmpData2FileRec.EMPTDATE = 0 Then 'CheckValDate(Format(DateAdd("d", (EmpData2FileRec.EMPTDATE), "12-31-1979"), "mm/dd/yyyy")) = False Then
      ValidEmpCnt = ValidEmpCnt + 1
      If Len(QPTrim$(FandLName)) = 0 Then GoTo SkipEm
      ThisCnt = ThisCnt + 1
      Print #RptHandle, QPTrim$(EmpData2FileRec.EmpNo); dlm; FandLName; dlm; AddDashToSSN(EmpData2FileRec.EmpSSN); dlm; Using$("$###0.00", EmpData2FileRec.EMPPRATE); dlm;
      Print #RptHandle, Left$(QPTrim$(EmpData2FileRec.EMPPTYPE), 1); dlm; BDate$; dlm; HDate$; dlm; QPTrim$(EmpData2FileRec.EMPJOB); dlm; fpcmbEmpStatus.Text; dlm; fpcmbEmpType.Text
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
  EnableCloseButton Me.hwnd, True
  Me.cmdEscape.Enabled = True
  MainLog ("Employee List report processed.")
  If ThisCnt = 0 Then
    MsgBox "There are no employees listed for this criteria."
    fpcmbEmpStatus.SetFocus
    Exit Sub
  End If
  
  If CheckName.Value = 1 Then
    arPrintAlphaNum.lblRptName = "Alphabetic Employee List"
  Else
    arPrintAlphaNum.lblRptName = "Numeric Employee List"
  End If
  arPrintAlphaNum.lblCity = City
  arPrintAlphaNum.lblTotals = "Total Employees: " & Using$("######", ValidEmpCnt)
  arPrintAlphaNum.Show
  frmLoadingRpt.Show
Exit Sub

Return

ErrorHandler:
  Close
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload FrmShowPctComp
  MsgBox "ERROR: If this problem persists please consult Southern Software."

EndTrans:
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
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Call LoadMe
  Me.HelpContextID = hlpPrintEmployeeList
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
      MainLog ("Payroll.exe terminated via Menu Bar on frmPrintAlphaNum.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub PrintText()
  
  Dim RptName As String, EmpIdxLNameHandle As Integer
  Dim EmpIdxNNameHandle As Integer, ThisSort() As Integer
  Dim FldFlag As String, DescFlag As String
  Dim IdxNumOfRecs As Integer, UnitHandle As Integer
  Dim Emp2RecLen As Integer, UnitFileRec As UnitFileRecType
  Dim DataFileSize As Long, cnt As Integer
  Dim RptHandle As Integer, D2Handle As Integer
  Dim RptTitle As String, D2Name As Long
  Dim EmpIdxLNameRec As NameSortIdxType
  Dim EmpIdxNNameRec As NumbSortIdxType
  Dim HDate$
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
  Dim ThisCnt As Integer
  
  Today = Date '$
  MaxLines = 55
  
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, UnitFileRec
  City = QPTrim$(UnitFileRec.UFEMPR)
  Close UnitHandle
  FF$ = Chr$(12)
  RptName$ = "EMPrintEmpList.RPT"
 
  If CheckName.Value = 1 Then
    RptTitle$ = "Employee Information Listing in Alphabetic Order"
    OpenEmpIdxLNameFile EmpIdxLNameHandle
    EmpIdxLNameCnt = LOF(EmpIdxLNameHandle) / 2
    If EmpIdxLNameCnt = 0 Then
      MsgBox "No records on file"
      Close
      MainLog ("Employee List Report exited with No records on file.")
      Exit Sub
    End If
    FrmShowPctComp.Label1 = "Alphabetic List of Employees Report"
    ReDim ThisSort(EmpIdxLNameCnt)
    For x = 1 To EmpIdxLNameCnt 'load array with employee data
    'sorted by last name
       Get EmpIdxLNameHandle, x, ThisSort(x)
    Next x
    Close EmpIdxLNameHandle
    IdxNumOfRecs = EmpIdxLNameCnt
  End If
  
  If CheckNumber.Value = 1 Then
    RptTitle$ = "Employee Information Listing in Numeric Order"
    OpenEmpIdxNNameFile EmpIdxNNameHandle
    EmpIdxNNameCnt = LOF(EmpIdxNNameHandle) / 2 'Len(EmpIdxNNameRec)
    If EmpIdxNNameCnt = 0 Then
      MsgBox "No records on file."
      Close
      Exit Sub
    End If
    FrmShowPctComp.Label1 = "List of Employees by Employee Number Report"
    ReDim ThisSort(EmpIdxNNameCnt)
    For x = 1 To EmpIdxNNameCnt 'load array sorted by employee
    'number
       Get EmpIdxNNameHandle, x, ThisSort(x)
    Next x
    Close EmpIdxNNameHandle
    IdxNumOfRecs = EmpIdxNNameCnt
  End If

  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  
  RptHandle = FreeFile
  
  Open RptName$ For Output As RptHandle
  RPTSetupPRN 2, RptHandle
  
  OpenEmpData2File EmpData2FileHandle
  GoSub PrintEmpListHeader
  For cnt = 1 To IdxNumOfRecs
    Get EmpData2FileHandle, ThisSort(cnt), EmpData2FileRec
    BDate = MakeRegDate(EmpData2FileRec.EMPBDAY)
    If BDate = "12/31/1979" Then BDate = "%%%%%%%%%% "
    HDate = MakeRegDate(EmpData2FileRec.EMPHDATE)
    If HDate = "12/31/1979" Then HDate = "%%%%%%%%%% "
    FandLName = QPTrim$(EmpData2FileRec.EmpLName) + ", " + QPTrim$(EmpData2FileRec.EmpFName)
    If EmpData2FileRec.EMPPRATE < 0 Then
       EmpData2FileRec.EMPPRATE = 0
    End If
    'Filter
    
    If fpcmbEmpType.Text = "ALL" Then
      GoTo AllsGood
    ElseIf fpcmbEmpType.Text = "Full-Time" Then
      If QPTrim$(EmpData2FileRec.EMPSTATS) <> "Full-Time" Then
        GoTo NotNow
      End If
    ElseIf fpcmbEmpType.Text = "Part-Time" Then
      If QPTrim$(EmpData2FileRec.EMPSTATS) <> "Part-Time" Then
        GoTo NotNow
      End If
    ElseIf fpcmbEmpType.Text = "Seasonal" Then
      If QPTrim$(EmpData2FileRec.EMPSTATS) <> "Seasonal" Then
        GoTo NotNow
      End If
    ElseIf fpcmbEmpType.Text = "Temporary" Then
      If QPTrim$(EmpData2FileRec.EMPSTATS) <> "Temporary" Then
        GoTo NotNow
      End If
    End If
    
AllsGood:
    If fpcmbEmpStatus.Text = "ALL" Then
      GoTo AllsGoodAgain
    ElseIf fpcmbEmpStatus.Text = "Active" Then
      If EmpData2FileRec.EMPTDATE <> 0 Then
        GoTo NotNow
      End If
    ElseIf fpcmbEmpStatus.Text = "Terminated" Then
      If EmpData2FileRec.EMPTDATE = 0 Then
        GoTo NotNow
      End If
    End If
    
AllsGoodAgain:
    If Not EmpData2FileRec.Deleted Then 'And EmpData2FileRec.EMPTDATE = 0 Then 'CheckValDate(Format(DateAdd("d", (EmpData2FileRec.EMPTDATE), "12-31-1979"), "mm/dd/yyyy")) = False Then
      ThisCnt = ThisCnt + 1
      ValidEmpCnt = ValidEmpCnt + 1
      Print #RptHandle, QPTrim$(EmpData2FileRec.EmpNo);
      Print #RptHandle, Tab(12); FandLName;
      Print #RptHandle, Tab(34); AddDashToSSN(EmpData2FileRec.EmpSSN);
      If EmpData2FileRec.EMPPRATE < 0 Then
        EmpData2FileRec.EMPPRATE = 0
      End If
      Print #RptHandle, Tab(48); Using$("$###0.00", EmpData2FileRec.EMPPRATE);
      Print #RptHandle, Tab(58); Left$(QPTrim$(EmpData2FileRec.EMPPTYPE), 1);
      Print #RptHandle, Tab(62); BDate$;
      Print #RptHandle, Tab(75); HDate$; Tab(90); QPTrim$(EmpData2FileRec.EMPJOB)
      LineCnt = LineCnt + 1
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintEmpListHeader
      End If
    End If
NotNow:
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
  Print #RptHandle, "-------------------------------------------------------------------------------------------------------------------"
  Print #RptHandle, "Total Employees: "; Using$("######", ValidEmpCnt)
  Print #RptHandle, FF$
  
  RPTSetupPRN 123, RptHandle '7/24
    
  Close EmpData2FileHandle
  Close RptHandle
  
  EnableCloseButton Me.hwnd, True
  Me.cmdEscape.Enabled = True
  MainLog ("Employee List report processed.")
  
  If ThisCnt = 0 Then
    MsgBox "There are no employees listed for this criteria."
    fpcmbEmpStatus.SetFocus
    Exit Sub
  End If
  
  ViewPrint RptName$, RptTitle$, True, , False
Exit Sub

PrintEmpListHeader:
  Print #RptHandle, City
  Print #RptHandle, "Report Date"; Tab(14); Today
  Print #RptHandle, "Employee Status: "; Tab(20); fpcmbEmpStatus.Text
  Print #RptHandle, "Employee Type: "; Tab(20); fpcmbEmpType.Text
  Print #RptHandle, ""
  Print #RptHandle, "Number         Name                 SSN          Pay    Pay  BirthDate    HireDate              Job Title"
  Print #RptHandle, "                                                 Rate   Type                      "
  Print #RptHandle, "-------------------------------------------------------------------------------------------------------------------"
  LineCnt = 8
Return
EndTrans:
End Sub

Private Sub fpcmbEmpStatus_Click()
  If fpcmbEmpStatus.Text = "" Then
    fpcmbEmpStatus.Text = "ALL"
  End If
End Sub

Private Sub fpcmbEmpStatus_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbEmpStatus.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbEmpStatus.ListIndex = -1
  End If
  If fpcmbEmpStatus.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbEmpType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbEmpType_Click()
  If fpcmbEmpType.Text = "" Then
    fpcmbEmpType.Text = "ALL"
  End If
End Sub

Private Sub fpcmbEmpType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbEmpType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbEmpType.ListIndex = -1
  End If
  If fpcmbEmpType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcomboPrintOpt.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
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

Private Sub LoadMe()
  fpcmbEmpStatus.Text = "ALL"
  fpcmbEmpStatus.AddItem "ALL"
  fpcmbEmpStatus.AddItem "Active"
  fpcmbEmpStatus.AddItem "Terminated"
  
  fpcmbEmpType.Text = "ALL"
  fpcmbEmpType.AddItem "ALL"
  fpcmbEmpType.AddItem "Full-Time"
  fpcmbEmpType.AddItem "Part-Time"
  fpcmbEmpType.AddItem "Seasonal"
  fpcmbEmpType.AddItem "Temporary"
  
  fpcomboPrintOpt.AddItem "Graphical"
  fpcomboPrintOpt.AddItem "Text"
  fpcomboPrintOpt.Text = "Graphical"
  

End Sub

Private Sub fpcomboPrintOpt_LostFocus()
  If QPTrim$(fpcomboPrintOpt.Text) = "" Then
    fpcomboPrintOpt.Text = "Graphical"
  End If
End Sub
