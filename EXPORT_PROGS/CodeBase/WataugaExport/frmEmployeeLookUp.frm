VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmEmployeeLookUp 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Look-Up"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmployeeLookUp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8840
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpListSearch 
      Height          =   2400
      Left            =   1695
      TabIndex        =   4
      Top             =   5955
      Width           =   8175
      _Version        =   196608
      _ExtentX        =   14420
      _ExtentY        =   4233
      TextAlias       =   ""
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
      Object.TabStop         =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Columns         =   4
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   0
      WrapList        =   0   'False
      WrapWidth       =   0
      SelMax          =   -1
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
      BorderStyle     =   1
      BorderColor     =   8454143
      BorderWidth     =   2
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ScrollHScale    =   2
      ScrollHInc      =   0
      ColsFrozen      =   0
      ScrollBarV      =   1
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   2
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   3
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   1
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
      ColumnHeaderShow=   -1  'True
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   -2147483637
      ExtendRow       =   0
      DataField       =   ""
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      ColDesigner     =   "frmEmployeeLookUp.frx":08CA
   End
   Begin EditLib.fpText fptxtLastName 
      Height          =   492
      Left            =   4368
      TabIndex        =   1
      ToolTipText     =   "Enter a Complete or Partial Last Name here. Entering ""Mc"" will find ""McCoy, McDonald"". Press (F5) to do Look-Up, (ESC) to Cancel. "
      Top             =   1896
      Width           =   3900
      _Version        =   196608
      _ExtentX        =   6879
      _ExtentY        =   868
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
      ThreeDInsideHighlightColor=   -2147483637
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
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
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
      CharValidationText=   ""
      MaxLength       =   24
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtFirstName 
      Height          =   492
      Left            =   4356
      TabIndex        =   2
      ToolTipText     =   "Enter a Complete or Partial First Name here. Entering ""Je"" will find ""Jeff, Jerry"". Press (F5) to do Look-Up, (ESC) to Cancel."
      Top             =   2616
      Width           =   3900
      _Version        =   196608
      _ExtentX        =   6879
      _ExtentY        =   868
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
      ThreeDInsideHighlightColor=   -2147483637
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
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
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
      CharValidationText=   ""
      MaxLength       =   24
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtNumber 
      Height          =   492
      Left            =   4716
      TabIndex        =   3
      ToolTipText     =   $"frmEmployeeLookUp.frx":0C8E
      Top             =   4032
      Width           =   3924
      _Version        =   196608
      _ExtentX        =   6921
      _ExtentY        =   868
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
      ThreeDInsideHighlightColor=   -2147483637
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
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
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
      CharValidationText=   ""
      MaxLength       =   10
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   615
      Left            =   6600
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Press ESC to exit this screen."
      Top             =   4848
      Width           =   1575
      _Version        =   131072
      _ExtentX        =   2778
      _ExtentY        =   1085
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
      ButtonDesigner  =   "frmEmployeeLookUp.frx":0D1F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSearch 
      Height          =   615
      Left            =   3510
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Press F5 to activate a search for a specific employee or group of employees."
      Top             =   4845
      Width           =   1560
      _Version        =   131072
      _ExtentX        =   2752
      _ExtentY        =   1085
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
      ButtonDesigner  =   "frmEmployeeLookUp.frx":0EFB
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
      ForeColor       =   &H8000000E&
      Height          =   348
      Left            =   2436
      TabIndex        =   7
      Top             =   4272
      Width           =   2052
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2052
      Left            =   1716
      Top             =   3672
      Width           =   8172
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      ForeColor       =   &H8000000E&
      Height          =   492
      Left            =   2736
      TabIndex        =   6
      Top             =   2784
      Width           =   1452
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      ForeColor       =   &H8000000E&
      Height          =   492
      Left            =   2724
      TabIndex        =   5
      Top             =   2064
      Width           =   1404
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1932
      Left            =   1728
      Top             =   1536
      Width           =   8172
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   972
      Index           =   1
      Left            =   1500
      Top             =   330
      Width           =   8652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Look-up"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2796
      TabIndex        =   0
      Top             =   570
      Width           =   6012
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1092
      Left            =   1500
      Top             =   210
      Width           =   8652
   End
End
Attribute VB_Name = "frmEmployeeLookUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim recordNum As Integer
Dim foundIt As Integer

Private Sub cmdExit_Click()
   frmEmployeeMaintMenu.Show
   DoEvents
   Unload frmEmployeeLookUp
End Sub

Private Sub cmdSearch_Click()
   Dim EmpData2FileHandle As Integer, rowCnt As Integer, EmpData1FileHandle As Integer
   Dim EmpData2FileRec As EmpData2Type, EmpData1FileRec As EmpData1Type
   Dim EmpRecordNum As Integer, x As Integer, Found As Boolean
   Dim EmployeeLastName As String, EmployeeFirstName As String
   Dim EmployeeNumber As String, NumEmpRec As Integer
   Dim MatchCnt As Integer, recordNum As Integer
   Dim ELNFlag As Integer, EFNFlag As Integer, ENFlag As Integer
   Dim IdxRecLen, IdxFileSize&
   Dim NumOfRecs As Integer, FoundCnt As Integer
   Dim IdxBuff() As Integer, NHandle As Integer
   Dim OnlyOneEmpNum As Long
   
   On Error GoTo ERRORSTUFF
   fpListSearch.Clear
   
   ELNFlag = 0
   EFNFlag = 0
   ENFlag = 0
   rowCnt = 1
   
   If Len(QPTrim$(fptxtLastName.Text)) > 0 Then
      EmployeeLastName = UCase$(QPTrim$(fptxtLastName.Text))
      ELNFlag = True
   End If
   
   If Len(QPTrim$(fptxtFirstName.Text)) > 0 Then
      EmployeeFirstName = UCase$(QPTrim$(fptxtFirstName.Text))
      EFNFlag = True
   End If
   
   If Len(QPTrim$(fptxtNumber.Text)) > 0 Then
      EmployeeNumber = QPTrim$(fptxtNumber.Text)
      ENFlag = True
   End If
   
  IdxRecLen = 2
  IdxFileSize& = FileSize(PRData + EmpIdxNName)
  NumOfRecs = IdxFileSize& \ IdxRecLen
  If NumOfRecs = 0 Then
    MsgBox "There are no employee records on file"
    Close
    Exit Sub
  End If
  OpenEmpIdxNNameFile NHandle
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    Get NHandle, x, IdxBuff(x)
  Next x
  Close NHandle

  OpenEmpData2File EmpData2FileHandle
'  OpenEmpData1File EmpData1FileHandle '8/19 commented out
   
  NumEmpRec = LOF(EmpData2FileHandle) / Len(EmpData2FileRec)
  For x = 1 To NumOfRecs
    Get EmpData2FileHandle, IdxBuff(x), EmpData2FileRec
'      If QPTrim$(EmpData2FileRec.EmpLName) = "ROBBINS" Then Stop
'    Get EmpData1FileHandle, IdxBuff(x), EmpData1FileRec '8/19 commented out
      If EmpData2FileRec.Deleted = -1 Then GoTo NotAMatch
      If Len(QPTrim$(EmpData2FileRec.EmpNo)) = 0 Then GoTo NotAMatch
      Found = True
      If ELNFlag Then
        If InStr(UCase$(EmpData2FileRec.EmpLName), EmployeeLastName) > 0 Then
          Found = True
        Else
          Found = False
          GoTo NotAMatch
        End If
      End If
      If EFNFlag Then
        If InStr(UCase$(EmpData2FileRec.EmpFName), EmployeeFirstName) > 0 Then
          Found = True
        Else
          Found = False
          GoTo NotAMatch
        End If
      End If
      If ENFlag Then
        If InStr(EmpData2FileRec.EmpNo, EmployeeNumber) > 0 Then
          Found = True
        Else
          Found = False
          GoTo NotAMatch
        End If
      End If
      If Found Then
        FoundCnt = FoundCnt + 1
        fpListSearch.Row = -1
        MatchCnt = MatchCnt + 1
        RecNum = x
        fpListSearch.InsertRow = EmpData1FileRec.TransRecNum & Chr$(9) & "    " & QPTrim$(EmpData2FileRec.EmpNo) & Chr$(9) & "  " & QPTrim$(EmpData2FileRec.EmpLName) & Chr$(9) & "   " & QPTrim$(EmpData2FileRec.EmpFName)
        DoEvents
        'only used if no more than one found
        OnlyOneEmpNum = QPTrim$(EmpData2FileRec.EmpNo)
      End If
NotAMatch:
      Next x
  
  If MatchCnt <= 0 Then
    MsgBox "No match found"
    Close
  End If
 'If FoundCnt is more than one then we can't load the
 'frmEditEmpData yet because we don't know which of the
 'list is to be selected
  If FoundCnt = 1 Then
    frmLoadingEmpEdit.Show
    DoEvents
    For x = 1 To NumEmpRec
      Get EmpData2FileHandle, x, EmpData2FileRec
        If EmpData2FileRec.Deleted = -1 Then GoTo NotThisTime 'added 9/3/04
        'to keep deleted employees matching emp numbers from pre-empting
        'the non-deleted employees
        If OnlyOneEmpNum = QPTrim$(EmpData2FileRec.EmpNo) Then
          RecNum = x
          Exit For
        Else
          Found = False
          GoTo NotThisTime
        End If
NotThisTime:
    Next x
    
    fptxtLastName.Text = ""
    fptxtFirstName.Text = ""
    fptxtNumber.Text = ""
    fpListSearch.Clear
    frmEditEmpData.Show
    DoEvents
'    Unload frmEmployeeLookUp'7/25/03
    DoEvents
    Unload frmLoadingEmpEdit
    FoundCnt = 0
  ElseIf FoundCnt > 1 Then
    fpListSearch.ListIndex = 0
  End If
  Close EmpData2FileHandle
'  Close EmpData1FileHandle '8/19 commented out
  Close
EndTrans:

  Exit Sub
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmEmployeeLookUp", "cmdSearch_Click", Erl)
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
    Close
End Sub

Private Sub cmdSearch_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
     Call fpListSearch_DblClick
     KeyCode = 0
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = vbKeyReturn Then
  'the next line was included to allow the user to have data
  'in the data fields and a selection in the list, use the
  'enter key as a way to process the selection...previously
  'if there was any data inserted in any field and a selection
  'was made in the list and enter was depressed then nothing
  'would happen
    If fpListSearch.ListIndex <> -1 Then GoTo EmpAlreadySelected '8/6
    If Len(fptxtNumber.Text) > 0 Or Len(fptxtLastName.Text) > 0 Or Len(fptxtFirstName.Text) > 0 Then
      Call cmdSearch_Click
      KeyCode = 0
      Exit Sub
    End If
EmpAlreadySelected: '8/6
    fpListSearch.Col = 1
    If QPTrim$(fpListSearch.ColText) = "" Then
      MsgBox "No employee has been selected"
      Exit Sub
    Else
      Call fpListSearch_DblClick
      KeyCode = 0
      Exit Sub
    End If
  End If
  Select Case KeyCode
    Case vbKeyDown:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%S"
      Call cmdSearch_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub fpListSearch_DblClick()
  Dim EmpData2FileHandle As Integer
  Dim EmpData2FileRec As EmpData2Type
  Dim EmpData1FileHandle As Integer
  Dim EmpData1FileRec As EmpData1Type
  Dim NumEmpRec As Integer, x As Integer
  Dim EmployeeLastName As String
  Dim EmployeeFirstName As String
  Dim EmployeeNumber As String, Found As Boolean
  
  Call DeActivateControls  '8/6 swapped with next line
  frmLoadingEmpEdit.Show '8/6 swapped with above line
  DoEvents
  fpListSearch.Col = 2
  'trap for double clicking on nothing
  If QPTrim$(fpListSearch.ColText) = "" Then
    MsgBox "No employee has been selected"
    Unload frmLoadingEmpEdit
    Exit Sub
  End If
  EmployeeLastName = QPTrim$(fpListSearch.ColText)
  fpListSearch.Col = 3
  EmployeeFirstName = QPTrim$(fpListSearch.ColText)
  fpListSearch.Col = 1
  EmployeeNumber = QPTrim$(fpListSearch.ColText)
  OpenEmpData2File EmpData2FileHandle
'  OpenEmpData1File EmpData1FileHandle '8/19 commented out
  
  NumEmpRec = LOF(EmpData2FileHandle) / Len(EmpData2FileRec)
  For x = 1 To NumEmpRec
     Get EmpData2FileHandle, x, EmpData2FileRec
'     Get EmpData1FileHandle, x, EmpData1FileRec '8/19 commented out
       If InStr(UCase$(EmpData2FileRec.EmpLName), EmployeeLastName) > 0 And InStr(UCase$(EmpData2FileRec.EmpFName), EmployeeFirstName) > 0 And InStr(EmpData2FileRec.EmpNo, EmployeeNumber) > 0 _
       And Len(QPTrim$(EmpData2FileRec.EmpNo)) = Len(QPTrim$(EmployeeNumber)) Then '8/7 added Len = Len because
       'if two people had the same name and the emp number of one had a number that
       'included the other's (ie. 123 vs 1234) then then smaller number would not be accessed ever
         Found = True
         fpListSearch.Row = -1
         RecNum = x
         Exit For
       Else
         Found = False
         GoTo NotAMatch
       End If
      
NotAMatch:
  Next x
'  fpListSearch.Clear '8/6 commented out (not needed since form unloads)
  Close EmpData2FileHandle
'  Close EmpData1FileHandle '8/19 commented out
  Close

  Load frmEditEmpData
  DoEvents
  frmEditEmpData.Show
  DoEvents
'  fptxtLastName.Text = "" '8/6 commented out (fixed bug that caused a crash
'  when payroll terminated)
'  fptxtFirstName.Text = "" '8/6 commented out
'  fptxtNumber.Text = "" '8/6 commented out
'  Unload frmEmployeeLookUp
'  frmEmployeeLookUp.Hide
  Unload frmLoadingEmpEdit
'  Call ActivateControls '8/6 not needed
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmEmployeeLookUp.")
      Call Terminate
      End
    End If
  End If
End Sub
Private Sub DeActivateControls()
  Dim cnt As Integer
  Dim x As Control
  Dim cmdButton As CommandButton

  cmdSearch.Enabled = False
  cmdExit.Enabled = False
  
  For cnt = 0 To Me.Count - 1
    Set x = Me.Controls.Item(cnt)
      If TypeOf x Is CommandButton Then
        x.Enabled = False
      End If
  Next cnt
    EnableCloseButton Me.hwnd, False
     
End Sub
Public Sub ActivateControls()
  Dim cmdButton As CommandButton
  Dim x As Control
  Dim cnt As Integer
  
  cmdSearch.Enabled = True
  cmdExit.Enabled = True
  
  For cnt = 0 To Me.Count - 1
    Set x = Me.Controls.Item(cnt)
      If TypeOf x Is CommandButton Then
        x.Enabled = True
      End If
  Next cnt
  EnableCloseButton Me.hwnd, True
     
End Sub

