VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmChkReprintInfo 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check Reprinting Information"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11652
   Icon            =   "frmChkReprintInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   5484
      Left            =   1848
      TabIndex        =   4
      Top             =   1698
      Width           =   7932
      _Version        =   196609
      _ExtentX        =   13991
      _ExtentY        =   9673
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDShadowColor=   -2147483633
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmChkReprintInfo.frx":08CA
      Begin EditLib.fpDateTime fpDTDateOfChks 
         Height          =   396
         Left            =   3984
         TabIndex        =   3
         Top             =   3600
         Width           =   1836
         _Version        =   196608
         _ExtentX        =   3238
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
         ButtonStyle     =   2
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
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "11/20/2002"
         DateCalcMethod  =   0
         DateTimeFormat  =   0
         UserDefinedFormat=   "mm/dd/yyyy"
         DateMax         =   "00000000"
         DateMin         =   "00000000"
         TimeMax         =   "000000"
         TimeMin         =   "000000"
         TimeString1159  =   ""
         TimeString2359  =   ""
         DateDefault     =   "00000000"
         TimeDefault     =   "000000"
         TimeStyle       =   0
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         PopUpType       =   1
         DateCalcY2KSplit=   60
         CaretPosition   =   0
         IncYear         =   1
         IncMonth        =   1
         IncDay          =   1
         IncHour         =   1
         IncMinute       =   1
         IncSecond       =   1
         ButtonColor     =   13684944
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText fptxtfirstBadChk 
         Height          =   396
         Left            =   4608
         TabIndex        =   0
         Top             =   1584
         Width           =   1788
         _Version        =   196608
         _ExtentX        =   3154
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
      Begin EditLib.fpText fptxtRestart 
         Height          =   396
         Left            =   4608
         TabIndex        =   2
         Top             =   2832
         Width           =   1788
         _Version        =   196608
         _ExtentX        =   3154
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
      Begin EditLib.fpText fptxtLastBadChk 
         Height          =   396
         Left            =   4608
         TabIndex        =   1
         Top             =   2208
         Width           =   1788
         _Version        =   196608
         _ExtentX        =   3154
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
      Begin fpBtnAtlLibCtl.fpBtn cmdAlignTest 
         Height          =   690
         Left            =   3315
         TabIndex        =   10
         ToolTipText     =   "Press to print a check mask used to align your printer."
         Top             =   4440
         Width           =   1470
         _Version        =   131072
         _ExtentX        =   2593
         _ExtentY        =   1217
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
         DrawFocusRect   =   4
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
         ButtonDesigner  =   "frmChkReprintInfo.frx":08E6
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
         Height          =   690
         Left            =   1080
         TabIndex        =   11
         ToolTipText     =   "Press to exit this screen."
         Top             =   4440
         Width           =   2070
         _Version        =   131072
         _ExtentX        =   3651
         _ExtentY        =   1217
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
         DrawFocusRect   =   4
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
         ButtonDesigner  =   "frmChkReprintInfo.frx":0AC2
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdPrintChecks 
         Height          =   696
         Left            =   4968
         TabIndex        =   12
         ToolTipText     =   "Press to begin printing payroll checks starting with the number entered above."
         Top             =   4440
         Width           =   2064
         _Version        =   131072
         _ExtentX        =   3641
         _ExtentY        =   1228
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
         DrawFocusRect   =   4
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
         ButtonDesigner  =   "frmChkReprintInfo.frx":0CA0
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Restarting Check Number:"
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
         Left            =   1296
         TabIndex        =   9
         Top             =   2928
         Width           =   3036
      End
      Begin VB.Label Label4 
         BackColor       =   &H00D0D0D0&
         Caption         =   "Last Bad Check Number:"
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
         Left            =   1680
         TabIndex        =   8
         Top             =   2304
         Width           =   2796
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Reprinting Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   396
         Left            =   1776
         TabIndex        =   7
         Top             =   624
         Width           =   4572
      End
      Begin VB.Label Label2 
         BackColor       =   &H00D0D0D0&
         Caption         =   "First Bad Check Number:"
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
         Left            =   1680
         TabIndex        =   6
         Top             =   1680
         Width           =   2796
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Date of Checks:"
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
         Left            =   1728
         TabIndex        =   5
         Top             =   3696
         Width           =   2028
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   828
         Left            =   1536
         Top             =   384
         Width           =   4956
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   5760
      Left            =   1716
      Top             =   1554
      Width           =   8220
   End
End
Attribute VB_Name = "frmChkReprintInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdAlignTest_Click()
  Dim Handle As Integer
  Dim TempHandle As Integer
  Dim SysRec As RegDSysFileRecType
  Dim SHandle As Integer
  Dim TextLine$
  Dim cnt As Integer
  
  InFileNames(1) = "PRDATA\PRSYS.DAT" '7/20 added
  InFileNames(2) = "PRDATA\PRPRNDF.DAT" '7/20 added
  
  If FilesROK(Me, InFileNames(), OutFileNames(), 2) = False Then '7/20 added
    Close '7/20
    Exit Sub '7/20
  End If '7/20
  
  BadMaskFlag = False
  OpenSysFile SHandle
  Get SHandle, 1, SysRec
  Close SHandle
  
  Select Case SysRec.CheckStyle
  Case 1:
    alnRpt = "PRData\P9013-39MSK.txt"
    InFileNames(1) = "PRData\P9013-39MSK.txt"
  Case 2:
    alnRpt = "PRData\P9013-42MSK.txt"
    InFileNames(1) = "PRData\P9013-42MSK.txt"
  Case 3:
    alnRpt = "PRData\P9028MSK.txt"
    InFileNames(1) = "PRData\P9028MSK.txt"
  Case 4:
    alnRpt = "PRData\P9007MSK.txt"
    InFileNames(1) = "PRData\P9007MSK.txt"
  Case 5:
    alnRpt = "PRData\Laser1MSK.txt"
    InFileNames(1) = "PRData\Laser1MSK.txt"
  Case 6:
    alnRpt = "PRData\Laser2Msk.txt"
    InFileNames(1) = "PRData\Laser2Msk.txt"
  Case 7:
    alnRpt = "PRData\P42CUSTMSK.txt"
    InFileNames(1) = "PRData\P42CUSTMSK.txt"
  Case Else:
    alnRpt = ""
  End Select
  
   If FilesROK(Me, InFileNames(), OutFileNames(), 1) = False Then
    Close
    Exit Sub
  End If
  
  Handle = FreeFile
  Open alnRpt For Input As #Handle
  
  TempHandle = FreeFile
  Open "PRDATA\TALIGN.MSK" For Output As #TempHandle
  RPTSetupPRN 15, TempHandle
  Do While Not eof(Handle)
    Line Input #Handle, TextLine   ' Read line into variable.
    Print #TempHandle, TextLine
  Loop
  RPTSetupPRN 123, TempHandle
  Close
  alnRpt = "PRDATA\TALIGN.MSK"
  
  doAlign = True
  frmPrintChks.Show 1
  alnRpt = ""
  
  If BadMaskFlag = False Then 'If statement as of 7/23
    fptxtRestart.Text = Val(fptxtRestart.Text) + NumOfAligns
  Else
    BadMaskFlag = False
  End If
  
End Sub

Private Sub cmdEscape_Click()
  ReprintChkOn = False
  frmChkPrintingMenu.Show
  DoEvents
  Unload frmChkReprintInfo
End Sub

Private Sub cmdPrintChecks_Click()
  Dim FChkNum&
  Dim LChkNum&
  Dim StartEmp As Integer
  Dim Num2Print As Long
  Dim CheckDate As Long
  Dim CheckNum As Long
  Dim SysHandle As Integer
  Dim SysRec As RegDSysFileRecType
  
  OpenSysFile SysHandle
  Get SysHandle, 1, SysRec
  Close SysHandle
  
  If SysRec.CheckStyle <= 4 Then
    InFileNames(1) = "PRDATA\PRSYS.DAT" '7/20 added
    InFileNames(2) = "PRDATA\PRUNIT.DAT" '7/20 added
    InFileNames(3) = "PRDATA\PRDEDCOD.DAT" '7/20 added
    InFileNames(4) = "PRDATA\PRERNCOD.DAT" '7/20 added
    InFileNames(5) = "PRDATA\PREMPN.IDX" '7/20 added
    InFileNames(6) = "PRDATA\PREMP2.DAT" '7/20 added
    InFileNames(7) = "PRDATA\PREMP3.DAT" '7/20 added
    InFileNames(8) = "PRDATA\PRPRNSET.DAT" '7/20 added
    InFileNames(9) = "PRDATA\PRTRANST.DAT" '7/20 added
    InFileNames(10) = "PRDATA\PRCHECKS.DAT" '7/20 added
    
    If FilesROK(Me, InFileNames(), OutFileNames(), 10) = False Then '7/20 retooled
      Close
      Exit Sub
    End If
  Else 'no need to check for prprnset if checks are laser
    InFileNames(1) = "PRDATA\PRSYS.DAT" '7/20 added
    InFileNames(2) = "PRDATA\PRUNIT.DAT" '7/20 added
    InFileNames(3) = "PRDATA\PRDEDCOD.DAT" '7/20 added
    InFileNames(4) = "PRDATA\PRERNCOD.DAT" '7/20 added
    InFileNames(5) = "PRDATA\PREMPN.IDX" '7/20 added
    InFileNames(6) = "PRDATA\PREMP2.DAT" '7/20 added
    InFileNames(7) = "PRDATA\PREMP3.DAT" '7/20 added
    InFileNames(8) = "PRDATA\PRTRANST.DAT" '7/20 added
    InFileNames(9) = "PRDATA\PRCHECKS.DAT" '7/20 added
    If FilesROK(Me, InFileNames(), OutFileNames(), 9) = False Then '7/20 retooled
      Close
      Exit Sub
    End If
  End If
  
  FChkNum = Val(fptxtfirstBadChk.Text)
  LChkNum = Val(fptxtLastBadChk.Text)
  StartEmp = GetStartEmp(FChkNum&)
  If (FChkNum& > LChkNum&) Or StartEmp = 0 Then
    frmWarnBadCheckNum.Show vbModal, Me
    Exit Sub
  Else
    Num2Print = (LChkNum& - FChkNum&) + 1
  End If
  CheckDate = Date2Num(fpDTDateOfChks.Text)
  CheckNum& = Val(fptxtRestart.Text)
  If CheckNum& <= 0 Then
     frmWarnBadCheckNum.Show vbModal, Me
    Exit Sub
  End If
  doAlign = True
  Call PrintChecks(StartEmp, CheckNum&, Num2Print, 1, CheckNum, CheckDate, True)
  Call GetVoidChkData
  
  MainLog ("Check reprint completed.")
'  frmChkPrintingMenu.Show
'  DoEvents
'  Unload frmChkReprintInfo

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%A"
      Call cmdAlignTest_Click
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdEscape_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdPrintChecks_Click
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
  Me.HelpContextID = hlpReprintSelected
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdEscape.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      Call UnloadAllFormsAndOpn(RegExit)
      MainLog ("Payroll.exe terminated via menu bar on frmChkReprintInfo.")
      End
    End If
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub
Private Sub LoadMe()
  Dim Today As String * 10
  Dim SysRec As RegDSysFileRecType
  Dim SHandle As Integer
  
  OpenSysFile SHandle
  Get SHandle, 1, SysRec
  Close SHandle
  
  If SysRec.CheckStyle = 5 Then cmdAlignTest.Visible = False
  ReprintChkOn = True
  Today = Date '$
  fpDTDateOfChks.Text = Today
End Sub


