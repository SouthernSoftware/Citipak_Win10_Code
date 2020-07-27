VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmBLTransDetail 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Transaction Detail"
   ClientHeight    =   4875
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   8430
   Icon            =   "frmBLTransDetail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   8430
   StartUpPosition =   1  'CenterOwner
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   540
      Left            =   4815
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "Press 'Escape' to close this screen."
      Top             =   3885
      Width           =   1965
      _Version        =   131072
      _ExtentX        =   3466
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmBLTransDetail.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   540
      Left            =   1650
      TabIndex        =   33
      TabStop         =   0   'False
      Tag             =   $"frmBLTransDetail.frx":0AA8
      Top             =   3888
      Width           =   2100
      _Version        =   131072
      _ExtentX        =   3704
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmBLTransDetail.frx":0B78
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   348
      Left            =   240
      TabIndex        =   34
      Top             =   3888
      Width           =   540
      _Version        =   131072
      _ExtentX        =   952
      _ExtentY        =   614
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
      ShapeRoundWidth =   195
      ShapeRoundHeight=   195
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
      MaxWidth        =   6000
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
   Begin VB.Label lblBalloon 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "HELP BALLOONS ON"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   1650
      TabIndex        =   35
      Top             =   4464
      Width           =   2100
   End
   Begin VB.Label lblCatBal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Index           =   6
      Left            =   6432
      TabIndex        =   32
      Tag             =   "The amount shown here is the balance for this revenue source after this transaction was posted."
      Top             =   3360
      Width           =   1212
   End
   Begin VB.Label lblCatBal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Index           =   5
      Left            =   6432
      TabIndex        =   31
      Tag             =   "The amount shown here is the balance for this revenue source after this transaction was posted."
      Top             =   3072
      Width           =   1212
   End
   Begin VB.Label lblCatBal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Index           =   4
      Left            =   6432
      TabIndex        =   30
      Tag             =   "The amount shown here is the balance for this revenue source after this transaction was posted."
      Top             =   2784
      Width           =   1212
   End
   Begin VB.Label lblCatBal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Index           =   3
      Left            =   6432
      TabIndex        =   29
      Tag             =   "The amount shown here is the balance for this revenue source after this transaction was posted."
      Top             =   2496
      Width           =   1212
   End
   Begin VB.Label lblCatBal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Index           =   2
      Left            =   6432
      TabIndex        =   28
      Tag             =   "The amount shown here is the balance for this revenue source after this transaction was posted."
      Top             =   2208
      Width           =   1212
   End
   Begin VB.Label lblCatBal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Index           =   1
      Left            =   6432
      TabIndex        =   27
      Tag             =   "The amount shown here is the balance for this revenue source after this transaction was posted."
      Top             =   1920
      Width           =   1212
   End
   Begin VB.Label lblCatBal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Index           =   0
      Left            =   6432
      TabIndex        =   26
      Tag             =   "The amount shown here is the balance for this revenue source after this transaction was posted."
      Top             =   1632
      Width           =   1212
   End
   Begin VB.Label lblCatAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Index           =   6
      Left            =   4944
      TabIndex        =   25
      Tag             =   "Transaction amount for this revenue source."
      Top             =   3360
      Width           =   1212
   End
   Begin VB.Label lblCatAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Index           =   5
      Left            =   4944
      TabIndex        =   24
      Tag             =   "Transaction amount for this revenue source."
      Top             =   3072
      Width           =   1212
   End
   Begin VB.Label lblCatAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Index           =   4
      Left            =   4944
      TabIndex        =   23
      Tag             =   "Transaction amount for this revenue source."
      Top             =   2784
      Width           =   1212
   End
   Begin VB.Label lblCatAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Index           =   3
      Left            =   4944
      TabIndex        =   22
      Tag             =   "Transaction amount for this revenue source."
      Top             =   2496
      Width           =   1212
   End
   Begin VB.Label lblCatAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Index           =   2
      Left            =   4944
      TabIndex        =   21
      Tag             =   "Transaction amount for this revenue source."
      Top             =   2208
      Width           =   1212
   End
   Begin VB.Label lblCatAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Index           =   1
      Left            =   4944
      TabIndex        =   20
      Tag             =   "Transaction amount for this revenue source."
      Top             =   1920
      Width           =   1212
   End
   Begin VB.Label lblCatAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Index           =   0
      Left            =   4944
      TabIndex        =   19
      Tag             =   "Transaction amount for this revenue source."
      Top             =   1632
      Width           =   1212
   End
   Begin VB.Label lblRevDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev Desc"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Index           =   6
      Left            =   480
      TabIndex        =   18
      Top             =   3360
      Width           =   4332
   End
   Begin VB.Label lblRevDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev Desc"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Index           =   5
      Left            =   480
      TabIndex        =   17
      Top             =   3072
      Width           =   4332
   End
   Begin VB.Label lblRevDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev Desc"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Index           =   4
      Left            =   480
      TabIndex        =   16
      Top             =   2784
      Width           =   4332
   End
   Begin VB.Label lblRevDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev Desc"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Index           =   3
      Left            =   480
      TabIndex        =   15
      Tag             =   $"frmBLTransDetail.frx":0D5B
      Top             =   2496
      Width           =   4332
   End
   Begin VB.Label lblRevDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev Desc"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Index           =   2
      Left            =   480
      TabIndex        =   14
      Top             =   2208
      Width           =   4332
   End
   Begin VB.Label lblRevDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev Desc"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Index           =   1
      Left            =   480
      TabIndex        =   13
      Top             =   1920
      Width           =   4332
   End
   Begin VB.Label lblRevDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev Desc"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Index           =   0
      Left            =   480
      TabIndex        =   12
      Top             =   1632
      Width           =   4332
   End
   Begin VB.Label lblPrintType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   4515
      TabIndex        =   11
      Tag             =   "The transaction type provides details as to the nature of this transaction."
      Top             =   339
      Width           =   3705
   End
   Begin VB.Label lblPrintDesc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Note"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   5085
      TabIndex        =   10
      Tag             =   "This field contains any user supplied description or note for this transaction."
      Top             =   771
      Width           =   3090
   End
   Begin VB.Label lblPrintAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Left            =   2256
      TabIndex        =   9
      Tag             =   "The amount shown here is the total amount of money handled for this transaction."
      Top             =   768
      Width           =   1356
   End
   Begin VB.Label lblPrintDate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Left            =   1920
      TabIndex        =   8
      Tag             =   "This is the date on which this transaction took place."
      Top             =   336
      Width           =   1212
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   300
      Left            =   6384
      TabIndex        =   7
      Top             =   1248
      Width           =   1260
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   "Type:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   345
      Left            =   3930
      TabIndex        =   6
      Top             =   339
      Width           =   585
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Note:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   345
      Left            =   3930
      TabIndex        =   5
      Top             =   771
      Width           =   1170
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Revenue"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   300
      Left            =   4896
      TabIndex        =   4
      Top             =   1248
      Width           =   1260
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Source"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   300
      Left            =   1920
      TabIndex        =   3
      Top             =   1248
      Width           =   876
   End
   Begin VB.Label lblAmt 
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Amount:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   348
      Left            =   240
      TabIndex        =   2
      Top             =   768
      Width           =   1980
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   348
      Left            =   240
      TabIndex        =   1
      Top             =   336
      Width           =   1692
   End
End
Attribute VB_Name = "frmBLTransDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call LoadMe
End Sub


Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
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
      SendKeys "%C"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF1:
      SendKeys "%H"
      Call cmdHelp_Click
      KeyCode = 0
    Case Else:
  End Select
  
End Sub
Public Sub LoadMe()
  Dim TransRec As ARTransRecType
  Dim TransHandle As Integer
  Dim ThisRec As Double
  Dim ThisTransDesc$
  Dim CatRec As ARNewCatCodeRecType
  Dim CHandle  As Integer
  Dim CatRecNums As Integer
  Dim NextEntry As Integer
  Dim x As Integer
  Dim UpDown$
  
  lblBalloon.Visible = False
  ThisRec = CDbl(frmBLTransHistJr.ThisRec)
  OpenTransFile TransHandle
  Get TransHandle, ThisRec, TransRec
  Close TransHandle
  
  NextEntry = 0
  Me.Caption = "Transaction Detail - " + QPTrim$(frmBLTransHistJr.ThisName)
  ThisTransDesc$ = ""
  
  lblPrintDate.Caption = MakeRegDate(TransRec.TransDate)
  Select Case TransRec.TransType
    Case 1
      ThisTransDesc$ = "License Charge"
      UpDown = "Up"
    Case 2
      ThisTransDesc$ = "Payment"
      UpDown = "Down"
    Case 6
      ThisTransDesc$ = "Penalty Charge"
      UpDown = "Up"
    Case 13
      ThisTransDesc$ = "Payment Adjustment Down"
      UpDown = "Down"
    Case 23
      ThisTransDesc$ = "Billing Adjustment Down"
      UpDown = "Down"
    Case 24
      ThisTransDesc$ = "Billing Adjustment Up"
      UpDown = "Up"
    Case Else
      ThisTransDesc$ = "Unknown"
  End Select
  
  If TransRec.DetailTransType > 0 Then
    Select Case TransRec.DetailTransType
      Case 101
        ThisTransDesc$ = "Penalty Charge"
        UpDown = "Up"
      Case 110
        ThisTransDesc$ = "License Charge"
        UpDown = "Up"
      Case 201
        ThisTransDesc$ = "Penalty Payment"
        UpDown = "Down"
      Case 210
        ThisTransDesc$ = "License Payment"
        UpDown = "Down"
      Case 211
        ThisTransDesc$ = "License and Penalty Payment"
        UpDown = "Down"
      Case 301
        ThisTransDesc$ = "Penalty Adjustment Down"
        UpDown = "Down"
      Case 310
        ThisTransDesc$ = "License Adjustment Down"
        UpDown = "Down"
      Case 311
        ThisTransDesc$ = "License and Penalty Adjustment Down"
        UpDown = "Down"
      Case 401
        If TransRec.TransType = 13 Then
          ThisTransDesc = "Down Pay Adjustment"
        Else
          ThisTransDesc = "Penalty Adjustment Up"
        End If
        UpDown = "Up"
      Case 410
        If TransRec.TransType = 13 Then
          ThisTransDesc = "Down Pay Adjustment"
        Else
          ThisTransDesc = "License Adjustment Up"
        End If
        UpDown = "Up"
      Case 411
        If TransRec.TransType = 13 Then
          ThisTransDesc = "Down Pay Adjustment"
        Else
          ThisTransDesc$ = "License and Penalty Adjustment Up"
        End If
        UpDown = "Up"
      Case Else
        ThisTransDesc$ = "Unknown"
    End Select
  Else
    Select Case TransRec.TransType
      Case 1
        ThisTransDesc$ = "License Charge"
        UpDown = "Up"
      Case 2
        ThisTransDesc$ = "Payment"
        UpDown = "Down"
      Case 6
        ThisTransDesc$ = "Penalty Charge"
        UpDown = "Up"
      Case 13
        ThisTransDesc$ = "Adjustment Down"
        UpDown = "Down"
      Case 23
        ThisTransDesc$ = "Adjustment Down"
        UpDown = "Down"
      Case 24
        ThisTransDesc$ = "Adjustment Up"
        UpDown = "Up"
      Case Else
        ThisTransDesc$ = "Unknown"
    End Select
  End If
  lblPrintType.Caption = ThisTransDesc$
  
  For x = 0 To 6
    lblRevDesc(x).Caption = ""
    lblCatAmt(x).Caption = ""
    lblCatBal(x).Caption = ""
  Next x
  
  OpenCatCodeFile CHandle
  CatRecNums = LOF(CHandle) / Len(CatRec)
  If CatRecNums = 0 Then Exit Sub
  
  If TransRec.CatCodeRec1 > 0 Then
    Get CHandle, TransRec.CatCodeRec1, CatRec
    lblRevDesc(NextEntry).Caption = QPTrim$(CatRec.CODEDESC)
    lblRevDesc(NextEntry).Tag = "Description for license category #1."
    If UpDown = "Down" And TransRec.CatLicAmt1 <> 0 Then
      lblCatAmt(NextEntry).Caption = Using("$#,###,##0.00", -TransRec.CatLicAmt1)
    Else
      lblCatAmt(NextEntry).Caption = Using("$#,###,##0.00", TransRec.CatLicAmt1)
    End If
    lblCatBal(NextEntry).Caption = Using("$#,###,##0.00", TransRec.CatLicBal1)
    NextEntry = NextEntry + 1
  End If
  If TransRec.CatCodeRec2 > 0 Then
    Get CHandle, TransRec.CatCodeRec2, CatRec
    lblRevDesc(NextEntry).Caption = QPTrim$(CatRec.CODEDESC)
    lblRevDesc(NextEntry).Tag = "Description for license category #2."
    If UpDown = "Down" And TransRec.CatLicAmt2 <> 0 Then
      lblCatAmt(NextEntry).Caption = Using("$#,###,##0.00", -TransRec.CatLicAmt2)
    Else
      lblCatAmt(NextEntry).Caption = Using("$#,###,##0.00", TransRec.CatLicAmt2)
    End If
    lblCatBal(NextEntry).Caption = Using("$#,###,##0.00", TransRec.CatLicBal2)
    NextEntry = NextEntry + 1
  End If
  If TransRec.CatCodeRec3 > 0 Then
    Get CHandle, TransRec.CatCodeRec3, CatRec
    lblRevDesc(NextEntry).Caption = QPTrim$(CatRec.CODEDESC)
    lblRevDesc(NextEntry).Tag = "Description for license category #3."
    If UpDown = "Down" And TransRec.CatLicAmt3 <> 0 Then
      lblCatAmt(NextEntry).Caption = Using("$#,###,##0.00", -TransRec.CatLicAmt3)
    Else
      lblCatAmt(NextEntry).Caption = Using("$#,###,##0.00", TransRec.CatLicAmt3)
    End If
    lblCatBal(NextEntry).Caption = Using("$#,###,##0.00", TransRec.CatLicBal3)
    NextEntry = NextEntry + 1
  End If
  If TransRec.CatCodeRec4 > 0 Then
    Get CHandle, TransRec.CatCodeRec4, CatRec
    lblRevDesc(NextEntry).Caption = QPTrim$(CatRec.CODEDESC)
    lblRevDesc(NextEntry).Tag = "Description for license category #4."
    If UpDown = "Down" And TransRec.CatLicAmt4 <> 0 Then
      lblCatAmt(NextEntry).Caption = Using("$#,###,##0.00", -TransRec.CatLicAmt4)
    Else
      lblCatAmt(NextEntry).Caption = Using("$#,###,##0.00", TransRec.CatLicAmt4)
    End If
    lblCatBal(NextEntry).Caption = Using("$#,###,##0.00", TransRec.CatLicBal4)
    NextEntry = NextEntry + 1
  End If
  If TransRec.CatCodeRec5 > 0 Then
    Get CHandle, TransRec.CatCodeRec5, CatRec
    lblRevDesc(NextEntry).Caption = QPTrim$(CatRec.CODEDESC)
    lblRevDesc(NextEntry).Tag = "Description for license category #5."
    If UpDown = "Down" And TransRec.CatLicAmt5 <> 0 Then
      lblCatAmt(NextEntry).Caption = Using("$#,###,##0.00", -TransRec.CatLicAmt5)
    Else
      lblCatAmt(NextEntry).Caption = Using("$#,###,##0.00", TransRec.CatLicAmt5)
    End If
    lblCatBal(NextEntry).Caption = Using("$#,###,##0.00", TransRec.CatLicBal5)
    NextEntry = NextEntry + 1
  End If
  
  Close CHandle
  
  lblRevDesc(NextEntry).Caption = "PENALTY"
  lblRevDesc(NextEntry).Tag = "Penalty fee activity."
  If UpDown = "Down" And TransRec.PenAmt <> 0 Then
    lblCatAmt(NextEntry).Caption = Using("$#,###,##0.00", -TransRec.PenAmt)
  Else
    lblCatAmt(NextEntry).Caption = Using("$#,###,##0.00", TransRec.PenAmt)
  End If
  lblCatBal(NextEntry).Caption = Using("$#,###,##0.00", TransRec.PenBal)
  NextEntry = NextEntry + 1
  
  lblRevDesc(NextEntry).Caption = "ISSUANCE FEE"
  lblRevDesc(NextEntry).Tag = "Issuance fee activity."
  If UpDown = "Down" And TransRec.IssAmt <> 0 Then
    lblCatAmt(NextEntry).Caption = Using("$#,###,##0.00", -TransRec.IssAmt)
  Else
    lblCatAmt(NextEntry).Caption = Using("$#,###,##0.00", TransRec.IssAmt)
  End If
  lblCatBal(NextEntry).Caption = Using("$#,###,##0.00", TransRec.IssBal)
  For x = 0 To 6
    If Len(lblRevDesc(x).Caption) = 0 Then
      lblRevDesc(x).Tag = ""
      lblCatAmt(x).Tag = ""
      lblCatBal(x).Tag = ""
    End If
  Next x
  
  lblPrintDesc.Caption = QPTrim$(TransRec.TransDesc)
  
  If UpDown = "Down" Then
    lblPrintAmt.Caption = QPTrim$(Using("$#,###,##0.00", -TransRec.TransAmount))
  Else
    lblPrintAmt.Caption = QPTrim$(Using("$#,###,##0.00", TransRec.TransAmount))
  End If
  
End Sub
Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    lblBalloon.Visible = True
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
  End If
End Sub

