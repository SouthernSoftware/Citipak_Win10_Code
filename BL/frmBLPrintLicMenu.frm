VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmBLPrintLicMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Business Licenses"
   ClientHeight    =   8730
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11565
   Icon            =   "frmBLPrintLicMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11565
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   300
      Left            =   5472
      TabIndex        =   1
      Top             =   8064
      Width           =   684
      _Version        =   131072
      _ExtentX        =   1206
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   9405029
      ForeColor       =   8454143
      Text            =   ""
      Shape           =   0
      ShapeRoundWidth =   180
      ShapeRoundHeight=   180
      BorderWidth     =   -1
      BorderColor     =   -2147483630
      ThreeDWidth     =   -1
      ThreeDShadowColor=   -2147483632
      ThreeDHighlightColor=   16777215
      ThreeDText      =   0
      ThreeDTextHighlightColor=   16777215
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignPictureH   =   0
      AlignPictureV   =   0
      PictureStyle    =   0
      WordWrap        =   -1  'True
      ScaleMode       =   1
      ThreeDStyle     =   2
      Position        =   0
      PosBaseX        =   0
      PosBaseY        =   0
      PosOffsetX      =   -100
      PosOffsetY      =   300
      MaxWidth        =   5000
      CloudInset      =   100
      CloudMinWidth   =   600
      TailShape       =   2
      TailType        =   2
      TailBaseOffsetOutside=   300
      TailBaseOffsetInside=   100
      TailBaseAxisOutside=   0
      TailBaseAxisInside=   0
      TailBubbleCount =   3
      AlignTextH      =   1
      AlignTextV      =   1
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      AutoScan        =   0
      ScanAllDescendants=   -1  'True
      Interval        =   500
      IntervalNext    =   200
      AutoSize        =   -1  'True
      UseTagProp      =   -1  'True
      HideOnInactiveApp=   0   'False
      HideOnMouseDown =   2
      HideOnKeyDown   =   2
      HideOnFocus     =   0   'False
      ScanDisabledControls=   -1  'True
      ThreeDAppearance=   0
      FollowFocus     =   0   'False
      TemplateName    =   ""
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSetCustLic 
      Height          =   444
      Left            =   3960
      TabIndex        =   2
      Tag             =   "Press to bring up a screen from which to set customer business licenses to print."
      Top             =   2370
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   783
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
      ButtonDesigner  =   "frmBLPrintLicMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdLicsReg 
      Height          =   435
      Left            =   3960
      TabIndex        =   3
      Tag             =   "Press to bring up a screen from which to process a business license register."
      Top             =   2895
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmBLPrintLicMenu.frx":0ABC
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintLaser 
      Height          =   435
      Left            =   3960
      TabIndex        =   4
      Tag             =   "Press to bring up a screen from which to print business licenses to a laser printer."
      Top             =   3420
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmBLPrintLicMenu.frx":0CAA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintForms 
      Height          =   435
      Left            =   3960
      TabIndex        =   5
      Tag             =   "Press to bring up a screen from which to print business licenses to a tractor fed printer."
      Top             =   3945
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmBLPrintLicMenu.frx":0E97
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdReprint 
      Height          =   444
      Left            =   3960
      TabIndex        =   6
      Tag             =   $"frmBLPrintLicMenu.frx":1087
      Top             =   4462
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   783
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
      ButtonDesigner  =   "frmBLPrintLicMenu.frx":114A
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPost 
      Height          =   444
      Left            =   3960
      TabIndex        =   7
      Tag             =   "Press to bring up a screen from which to post business licenses."
      Top             =   4985
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   783
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
      ButtonDesigner  =   "frmBLPrintLicMenu.frx":133C
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintNoPost 
      Height          =   444
      Left            =   3960
      TabIndex        =   8
      Tag             =   "Use this feature when a license needs to be printed without charging the customer a fee."
      Top             =   5508
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   783
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
      ButtonDesigner  =   "frmBLPrintLicMenu.frx":1521
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClearFlags 
      Height          =   444
      Left            =   3960
      TabIndex        =   9
      Tag             =   $"frmBLPrintLicMenu.frx":170F
      Top             =   6031
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   783
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
      ButtonDesigner  =   "frmBLPrintLicMenu.frx":17FA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   444
      Left            =   3960
      TabIndex        =   10
      Top             =   6554
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   783
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
      ButtonDesigner  =   "frmBLPrintLicMenu.frx":19EB
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   444
      Left            =   3960
      TabIndex        =   11
      Tag             =   "Click this button to return to the main Business License menu."
      Top             =   7080
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   783
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
      ButtonDesigner  =   "frmBLPrintLicMenu.frx":1BD0
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   155
      Index           =   3
      Left            =   8550
      Top             =   1995
      Width           =   990
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   150
      Index           =   4
      Left            =   1970
      Top             =   2000
      Width           =   990
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8666
      X2              =   8666
      Y1              =   2136
      Y2              =   8008
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1092
      Index           =   1
      Left            =   1455
      Top             =   820
      Width           =   8652
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8655
      X2              =   9369
      Y1              =   8010
      Y2              =   8010
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2086
      X2              =   2795
      Y1              =   8025
      Y2              =   8025
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2086
      X2              =   2086
      Y1              =   2133
      Y2              =   8005
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LICENSE PROCESSING "
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
      Left            =   2775
      TabIndex        =   0
      Top             =   1170
      Width           =   6012
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1455
      Top             =   690
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   1966
      Top             =   1890
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2086
      Top             =   2130
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8550
      Top             =   1890
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8655
      Top             =   2130
      Width           =   732
   End
End
Attribute VB_Name = "frmBLPrintLicMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdClearFlags_Click()
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim x As Integer
  Dim DelCnt As Integer
  Dim DelFile As Integer
  
  frmBLMessageBoxJrWOpts.Label1.Caption = "This action will delete all files associated with any current business license fee processing. All 'Set Renewal Flag (Y/N)?' customer flags set to 'Y', making them eligible for a license fee renewal, will be reset to 'N'. Any business license forms already printed will need to be evaluated for suitability and disposed of as necessary. Do you wish to continue anyway?"
  frmBLMessageBoxJrWOpts.Label1.Top = 350
  frmBLMessageBoxJrWOpts.Label1.Height = 1500
  frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
  frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Abort"
  frmBLMessageBoxJrWOpts.Show vbModal
  If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
    Unload frmBLMessageBoxJrWOpts
    Close
    Exit Sub
  Else
    Unload frmBLMessageBoxJrWOpts
  End If
  
  OpenCustFile CustHandle
  NumOfCustRecs = LOF(CustHandle) / Len(CustRec)
  If NumOfCustRecs = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no customers on file at this time."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  DelCnt = 0
  
  frmBLShowPctComp.Label1 = "Clearing Business License Fee Files"
  frmBLShowPctComp.cmdCancel.Visible = False
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
      If CustRec.IssueLicense = "Y" Then
        CustRec.IssueLicense = "N"
        DelCnt = DelCnt + 1
      End If
    Put CustHandle, x, CustRec
    frmBLShowPctComp.ShowPctComp x, NumOfCustRecs
    If frmBLShowPctComp.Out = True Then
      Close
      frmBLShowPctComp.Out = False
      Unload frmBLShowPctComp
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  
  Close
  DelFile = 0
  If Exist("artmppst.dat") Then
    DelFile = 1
    KillFile "artmppst.dat"
  End If
  
  If DelFile = 1 Then
    If Exist("artmplic.dat") And Exist("licprnOK.dat") Then
        DelFile = 4
        KillFile "licprnOK.dat"
        KillFile "artmplic.dat"
    ElseIf Exist("licprnOK.dat") Then
        DelFile = 3
        KillFile "licprnOK.dat"
    ElseIf Exist("artmplic.dat") Then
        DelFile = 2
        KillFile "artmplic.dat"
    End If
  End If
  
  If DelCnt > 0 Then
    frmBLMessageBoxJr.Label1.Caption = "A total of " + CStr(DelCnt) + " customer flags have been reset to 'N' and all business license fee processing files have been removed successfully. Business license fee processing may be restarted at any time."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
  Else
    frmBLMessageBoxJr.Label1.Caption = "There were no customer flags discovered that needed resetting. Business license fee processing may be restarted at any time."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
  End If
  
  If DelCnt > 0 Then
    If DelFile = 1 Then
      MainLog ("Business license fee processing cleared. A total of " + CStr(DelCnt) + " customer flags were reset to 'N'. The file artmppst.dat was deleted.")
    ElseIf DelFile = 2 Then
      MainLog ("Business license fee processing cleared. A total of " + CStr(DelCnt) + " customer flags were reset to 'N'. The files artmppst.dat & artmplic.dat was deleted.")
    ElseIf DelFile = 3 Then
      MainLog ("Business license fee processing cleared. A total of " + CStr(DelCnt) + " customer flags were reset to 'N'. The files artmppst.dat and licprnOK.dat were deleted.")
    ElseIf DelFile = 4 Then
      MainLog ("Business license fee processing cleared. A total of " + CStr(DelCnt) + " customer flags were reset to 'N'. The files artmppst.dat, artmplic.dat & licprnOK.dat was deleted.")
    End If
  Else
    MainLog ("Business license fee processing file clearing was attempted but no customer flags were found that needed resetting.")
  End If
  
End Sub

Private Sub cmdExit_Click()
  frmBLMainMenu.Show
  DoEvents
  Unload frmBLPrintLicMenu
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "&Turn Menu Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "&Turn Menu Help On"
    btnHelp.AutoScan = fpAutoScanOff
  End If
End Sub

Private Sub cmdPost_Click()
  Dim TempRec As TempTransPostType
  Dim TempHandle As Integer
  Dim TransRec As ARTransRecType
  Dim TransHandle As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfTempRecs As Long
  Dim x As Long
  Dim NumOfTransRecs As Long
  Dim NextTransRec As Long
  Dim LicCount As Integer
  Dim LaserRec1 As LaserLetterType1
  Dim LaserRec2 As LaserLetterType2
  Dim LHandle As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  
  On Error GoTo ERRORSTUFF
  
  If Not Exist("artmppst.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Posting cannot take place before business license registers are processed. Please run business license registers before continuing."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  Else
    If Not Exist("licprnOK.dat") And Not Exist("artmplic.dat") Then
      frmBLMessageBoxJr.Label1.Caption = "Printing business licenses has not yet taken place. This step is required before posting because important business license dates are established at that time. Please print business licenses before continuing."
      frmBLMessageBoxJr.Label1.Top = 600
      frmBLMessageBoxJr.Show vbModal
      Close
      Exit Sub
    End If
    
    OpenTownFile TownHandle
    Get TownHandle, 1, TownRec
    Close TownHandle
    
    OpenTempPostFile TempHandle
    NumOfTempRecs = LOF(TempHandle) / Len(TempRec)
    
    If NumOfTempRecs = 0 Then
      frmBLMessageBoxJr.Label1.Caption = "No business licenses to post. Posting aborted."
      frmBLMessageBoxJr.Label1.Top = 800
      frmBLMessageBoxJr.Show vbModal
      KillFile "licprnOK.dat"
      KillFile "artmppst.dat"
      KillFile "artmplic.dat"
      Close
      Exit Sub
    End If
  End If
  
  frmBLWarnPost.Show vbModal
  If frmBLWarnPost.fptxtChoice = "exit" Then
    Unload frmBLWarnPost
    Close
    Exit Sub
  End If
  
  Unload frmBLWarnPost
  OpenCustFile CustHandle
'  GoSub UniversalSave
'  If TempRec.ChargeAccount = True Then
    OpenTransFile TransHandle
    NumOfTransRecs = LOF(TransHandle) / Len(TransRec)
    NextTransRec = NumOfTransRecs + 1
    GoSub UniversalSave
'    GoSub ChargeYes
'  End If
  
  Close
  KillFile "artmppst.dat"
  KillFile "artmplic.dat"
  KillFile "licprnOK.dat"
  
  Call CreateLicNumIdx
  
  frmBLSucSave.Label1.Caption = "License data has been successfully posted."
  frmBLSucSave.Label1.Top = 700
  frmBLSucSave.Show vbModal
  MainLog ("Business licenses posted.")
  
  Close
  Exit Sub
  
UniversalSave:
  For x = 1 To NumOfTempRecs
    Get TempHandle, x, TempRec
    Get CustHandle, TempRec.CustomerNumber, CustRec
    CustRec.LICENSE = TempRec.LICENSE
    CustRec.VALID = TempRec.VALID
    CustRec.Prorate = 100
    CustRec.IssueLicense$ = "N"
    CustRec.AcctBal = TempRec.AcctBal
    CustRec.IssuanceFee = TempRec.IssFee
    CustRec.IssuanceBal = TempRec.IssFeeBal
    CustRec.Fee1 = TempRec.CatFee1
    CustRec.Fee2 = TempRec.CatFee2
    CustRec.Fee3 = TempRec.CatFee3
    CustRec.Fee4 = TempRec.CatFee4
    CustRec.Fee5 = TempRec.CatFee5
    CustRec.FeeAmt = CustRec.Fee1 + CustRec.Fee2 + CustRec.Fee3 + CustRec.Fee4 + CustRec.Fee5
    CustRec.PenBal = TempRec.PenBal
    CustRec.FeeLicBal1 = TempRec.CatFeeBal1
    CustRec.FeeLicBal2 = TempRec.CatFeeBal2
    CustRec.FeeLicBal3 = TempRec.CatFeeBal3
    CustRec.FeeLicBal4 = TempRec.CatFeeBal4
    CustRec.FeeLicBal5 = TempRec.CatFeeBal5
    CustRec.FeeBal = CustRec.FeeLicBal1 + CustRec.FeeLicBal2 + CustRec.FeeLicBal3 + CustRec.FeeLicBal4 + CustRec.FeeLicBal5 'CustRec.FeeBal + CustRec.FeeAmt
    CustRec.LicBal = TempRec.LicBal
    TransRec.LicBal = CustRec.LicBal
    TransRec.LicAmt = TempRec.CatFee1 + TempRec.CatFee2 + TempRec.CatFee3 + TempRec.CatFee4 + TempRec.CatFee5
    TransRec.FeeAmt = TempRec.CatFee1 + TempRec.CatFee2 + TempRec.CatFee3 + TempRec.CatFee4 + TempRec.CatFee5
    TransRec.PenAmt = 0
    TransRec.PenBal = TempRec.PenBal
    TransRec.IssBal = CustRec.IssuanceBal
    TransRec.CustomerNumber = TempRec.CustomerNumber
    TransRec.TransDate = TempRec.TransDate
    TransRec.TransAmount = TempRec.TransAmount
    TransRec.TransType = TempRec.TransType
    TransRec.TransDesc = TempRec.TransDesc
    TransRec.CashAmount = 0
    TransRec.ChkAmount = 0
    TransRec.BalanceAfterTrans = TempRec.BalanceAfterTrans
    TransRec.ExtraRoom = ""
    TransRec.Posted2GL = TempRec.Posted2GL
    TransRec.NextTrans = 0
    TransRec.CatCodeRec1 = TempRec.CatCodeRec1
    TransRec.CatCodeRec2 = TempRec.CatCodeRec2
    TransRec.CatCodeRec3 = TempRec.CatCodeRec3
    TransRec.CatCodeRec4 = TempRec.CatCodeRec4
    TransRec.CatCodeRec5 = TempRec.CatCodeRec5
    
    TransRec.CatLicAmt1 = TempRec.CatFee1
    TransRec.DetailTransType = 0
    If TempRec.CatFee1 > 0 Then
      TransRec.DetailTransType = 110
    End If
    
'    TransRec.IssAmt = TempRec.IssFee
'    If TempRec.IssFee > 0 Then
    TransRec.IssAmt = TownRec.IssFee
    If TownRec.IssFee > 0 Then
      TransRec.DetailTransType = 110
    End If
    
    TransRec.CatLicAmt2 = TempRec.CatFee2
    If TempRec.CatFee2 > 0 Then
      TransRec.DetailTransType = 110
    End If
    
    TransRec.CatLicAmt3 = TempRec.CatFee3
    If TempRec.CatFee3 > 0 Then
      TransRec.DetailTransType = 110
    End If
    
    TransRec.CatLicAmt4 = TempRec.CatFee4
    If TempRec.CatFee4 > 0 Then
      TransRec.DetailTransType = 110
    End If
    
    TransRec.CatLicAmt5 = TempRec.CatFee5
    If TempRec.CatFee5 > 0 Then
      TransRec.DetailTransType = 110
    End If
    
    TransRec.CatLicBal1 = CustRec.FeeLicBal1
    TransRec.CatLicBal2 = CustRec.FeeLicBal2
    TransRec.CatLicBal3 = CustRec.FeeLicBal3
    TransRec.CatLicBal4 = CustRec.FeeLicBal4
    TransRec.CatLicBal5 = CustRec.FeeLicBal5
    'save this new post in the next transaction record
    Put TransHandle, NextTransRec, TransRec
    If CustRec.FirstTrans = 0 Then
      CustRec.FirstTrans = NextTransRec
      CustRec.LastTrans = NextTransRec
      Put CustHandle, TempRec.CustomerNumber, CustRec
    Else
      CustRec.LastTrans = NextTransRec
      Put CustHandle, TempRec.CustomerNumber, CustRec
      'get the transaction just prior to this one
      'for this customer and update the .NextTrans field
      'with the new post record number...this continues the
      'linkage among transactions for this customer
      Get TransHandle, TempRec.Prev, TransRec
      TransRec.NextTrans = NextTransRec
      Put TransHandle, TempRec.Prev, TransRec
    End If
    NextTransRec = NextTransRec + 1
  Next x
  
  Return
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLPrintLicMenu", "cmdPost_Click", Erl)
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
    ClearInUse PWcnt
    Terminate
  
  
End Sub

Private Sub cmdPrintForms_Click()
  frmBLPrintLic.Show
  DoEvents
  Unload frmBLPrintLicMenu
End Sub

Private Sub cmdPrintLaser_Click()
  frmBLLicFormLaser.Show
  DoEvents
  Unload frmBLPrintLicMenu
End Sub

Private Sub cmdPrintNoPost_Click()
  frmBLMessageBoxJr.Label1.Caption = "NOTICE: All business licenses printed from the 'Print Business Licenses: No Posting' screen make use of the CURRENT data saved for each customer. This data is collected and saved on the 'Customer Maintenance' screen."
  frmBLMessageBoxJr.Label1.Top = 550
  frmBLMessageBoxJr.Show vbModal
  
  frmBLPrintLicNoPost.Show
  DoEvents
  Unload frmBLPrintLicMenu
End Sub

Private Sub cmdReprint_Click()
  On Error Resume Next
  If Not Exist("artmplic.dat") Then
    frmBLMessageBoxJrWOpts.Label1.Caption = "Reprints are not possible until Form Fed Business Licenses are printed first. If you wish to print Form Fed Business Licenses then press F10 to jump to that screen."
    frmBLMessageBoxJrWOpts.Label1.Top = 700
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Jump"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
      Unload frmBLMessageBoxJrWOpts
      frmBLPrintLic.Show
      DoEvents
      Unload Me
      Exit Sub
    Else
      Unload frmBLMessageBoxJrWOpts
      Exit Sub
    End If
  End If
  
  frmBLReprintLic.Show
  DoEvents
  Unload frmBLPrintLicMenu
End Sub

Private Sub Form_Load()
  'License processing begins with setting the customers
  'who have licenses set to expire on the date entered
  'by the user which sets .IssueLicense to "Y"
  
  'then registers are run which generates the license
  'fees and sets these numbers in a temporary file for
  'posting later
  
  'then the actual business license forms are run which
  'generates the new expiration date, the license number
  'and the transaction linking value...unless the user
  'opts to not charge the customers in which case all fee
  'related temp fields are zeroed out but the license number
  'and new expiration date are still saved for posting...
  'this data is saved in the same temporary file as that
  'generated in the registers above
  
  'then the final posting takes place in which all the data
  'is committed to memory in permanent files: CustRec and
  'TransRec
  
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"
      Call cmdExit_Click
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
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLPrintLicMenu.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub cmdLicsReg_Click()
  Dim Towncnt As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  
  OpenTownFile TownHandle
  Towncnt = LOF(TownHandle) / Len(TownRec)
  Close TownHandle
  If Towncnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No town setup records have been saved. Please save the town setup records before continuing."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  frmBLLicRegister.Show
  DoEvents
  Unload frmBLPrintLicMenu
End Sub

Private Sub cmdSetCustLic_Click()
  frmBLChangeLicPrintStatus.Show
  DoEvents
  Unload frmBLPrintLicMenu
End Sub

