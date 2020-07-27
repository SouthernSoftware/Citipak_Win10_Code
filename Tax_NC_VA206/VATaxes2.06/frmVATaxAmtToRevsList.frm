VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxAmtToRevsList 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Message"
   ClientHeight    =   4308
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   6900
   Icon            =   "frmVATaxAmtToRevsList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4308
   ScaleWidth      =   6900
   StartUpPosition =   1  'CenterOwner
   Begin LpLib.fpList fpList 
      Height          =   2352
      Left            =   924
      TabIndex        =   0
      Top             =   1056
      Width           =   5052
      _Version        =   196608
      _ExtentX        =   8911
      _ExtentY        =   4149
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
      Columns         =   0
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
      DataField       =   ""
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      ColDesigner     =   "frmVATaxAmtToRevsList.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   672
      Left            =   2508
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3567
      Width           =   1872
      _Version        =   131072
      _ExtentX        =   3302
      _ExtentY        =   1185
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
      ButtonDesigner  =   "frmVATaxAmtToRevsList.frx":0B1E
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The following list indicates the number of changes made for each transaction type."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   648
      Left            =   360
      TabIndex        =   2
      Top             =   213
      UseMnemonic     =   0   'False
      Width           =   6156
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   756
      Left            =   132
      Top             =   69
      Width           =   6636
   End
End
Attribute VB_Name = "frmVATaxAmtToRevsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdExit_Click()
  Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyEscape:
    Call cmdExit_Click
    KeyCode = 0
  Case Else:
  End Select
End Sub
Private Sub Form_Load()
  Dim x As Integer
  Dim ThisCnt As Long
  
  For x = 1 To 17
    ThisCnt = TypeCnt(x)
    If ThisCnt > 0 Then
      Select Case x
        Case 1
          fpList.AddItem "Billing " + CStr(ThisCnt)
        Case 2
          fpList.AddItem "Payments " + CStr(ThisCnt)
        Case 3
          fpList.AddItem "Release " + CStr(ThisCnt)
        Case 4
          fpList.AddItem "Interest " + CStr(ThisCnt)
        Case 5
          fpList.AddItem "Penalty " + CStr(ThisCnt)
        Case 6
          fpList.AddItem "Advertising " + CStr(ThisCnt)
        Case 7
          fpList.AddItem "Adjust Pay Down " + CStr(ThisCnt)
        Case 8
          fpList.AddItem "Credit At Billing " + CStr(ThisCnt)
        Case 9
          fpList.AddItem "Adjust Bill Down " + CStr(ThisCnt)
        Case 10
          fpList.AddItem "Adjust Bill Up " + CStr(ThisCnt)
        Case 11
          fpList.AddItem "Bill Pay/Overpay " + CStr(ThisCnt)
        Case 12
          fpList.AddItem "Adjust Bill Up/Affect Credit " + CStr(ThisCnt)
        Case 13
          fpList.AddItem "Adjust Bill Dn/Affect Credit " + CStr(ThisCnt)
      End Select
    End If
  Next x
          
End Sub


