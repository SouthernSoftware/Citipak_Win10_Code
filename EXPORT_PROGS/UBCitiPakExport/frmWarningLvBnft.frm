VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmWarningLvBnft 
   BackColor       =   &H000000C0&
   BorderStyle     =   0  'None
   Caption         =   "Add New Form Warning"
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6810
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   355
      Left            =   6108
      Top             =   240
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
      Height          =   540
      Left            =   1452
      TabIndex        =   3
      Top             =   2688
      Width           =   1572
      _Version        =   131072
      _ExtentX        =   2773
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmWarningLvBnft.frx":0000
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAddNew 
      Height          =   540
      Left            =   3900
      TabIndex        =   4
      Top             =   2688
      Width           =   1572
      _Version        =   131072
      _ExtentX        =   2773
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmWarningLvBnft.frx":0216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      Caption         =   "Add a New Leave Table, Are you Sure?"
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
      Height          =   780
      Left            =   1242
      TabIndex        =   2
      Top             =   468
      Width           =   4332
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Press ""ESC"" to CANCEL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1794
      TabIndex        =   1
      Top             =   1572
      Width           =   3228
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      BackStyle       =   0  'Transparent
      Caption         =   " Press ""F8"" to Add New"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1794
      TabIndex        =   0
      Top             =   2052
      Width           =   3228
   End
End
Attribute VB_Name = "frmWarningLvBnft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Public Enum SaveChangeLBOptions1
  sclboInvalidOption = 0
  sclboAddNew
  sclboEscape
End Enum

Private m_sclboOption As SaveChangeLBOptions1

'// Create a property to get the Selection value.
'   NOTE: A Read-Only property has a Property Get but
'         no Property Let or Property Set
Property Get Selection() As SaveChangeLBOptions1
  Selection = m_sclboOption
End Property

Private Sub cmdEscape_Click()
'  On Error Resume Next
  m_sclboOption = sclboEscape
  Unload frmWarningLvBnft
  MainLog ("Add another leave table warning issued...escape option chosen.")

End Sub

Private Sub cmdAddNew_Click()
'  On Error Resume Next
  m_sclboOption = sclboAddNew
  Unload frmWarningLvBnft
  MainLog ("Add another leave table warning issued...add new option chosen.")

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyF8:
      Call cmdAddNew_Click
      KeyCode = 0
    Case vbKeyEscape:
      Call cmdEscape_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Timer1_Timer()
  Static tog As Boolean
  tog = Not tog
  If tog Then
    Me.BackColor = 210
    Label1.BackColor = 210
    Label2.BackColor = 210
    Label3.BackColor = 210
  Else
    Me.BackColor = 192
    Label1.BackColor = 192
    Label2.BackColor = 192
    Label3.BackColor = 192
  End If
End Sub

