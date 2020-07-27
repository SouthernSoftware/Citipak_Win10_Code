VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Begin VB.Form frmViewBud 
   BackColor       =   &H008A775B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Preview Budget Prep"
   ClientHeight    =   8916
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   12216
   Icon            =   "frmViewBud.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8916
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   288
      Top             =   96
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   396
      Left            =   168
      TabIndex        =   1
      Top             =   48
      Width           =   11940
      _Version        =   196613
      _ExtentX        =   21061
      _ExtentY        =   698
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowColor     =   12632256
      ShadowDark      =   8421504
      SpreadDesigner  =   "frmViewBud.frx":08CA
   End
   Begin FPSpread.vaSpreadPreview vaSpreadPreview1 
      Height          =   8004
      Left            =   168
      TabIndex        =   0
      Top             =   408
      Width           =   11940
      _Version        =   196613
      _ExtentX        =   21061
      _ExtentY        =   14118
      _StockProps     =   96
      BorderStyle     =   1
      AllowUserZoom   =   -1  'True
      GrayAreaColor   =   8421504
      GrayAreaMarginH =   720
      GrayAreaMarginType=   0
      GrayAreaMarginV =   720
      PageBorderColor =   8388608
      PageBorderWidth =   2
      PageShadowColor =   0
      PageShadowWidth =   2
      PageViewPercentage=   100
      PageViewType    =   0
      ScrollBarH      =   1
      ScrollBarV      =   1
      ScrollIncH      =   360
      ScrollIncV      =   360
      PageMultiCntH   =   1
      PageMultiCntV   =   1
      PageGutterH     =   -1
      PageGutterV     =   -1
      ScriptEnhanced  =   0   'False
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   2
      Top             =   8664
      Width           =   12216
      _ExtentX        =   21548
      _ExtentY        =   445
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
            TextSave        =   "12:30 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "10/15/2004"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmViewBud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      
      Unload Me
      KeyCode = 0
    Case Else:
  End Select
End Sub


Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub Form_Activate()
 
    'Attach preview control to Spread
   
    Me.vaSpreadPreview1.hWndSpread = frmBudPrepMaint.vaSpread1.hwnd
    
    'Update page count listing
    UpdatePageCount
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName

  
    SetupToolbar
    
    'Disable Previous button
    DisableButton 6, "LEFT"
        
    'Get the zoom display
    zoomindex = 3
    GetZoom zoomindex
    
    'Set up page numbering
    If frmBudPrepMaint.vaSpread1.PrintPageCount = 1 Then
        'Disable Next button if only one page
        DisableButton 4, "LEFT"
    End If
           
End Sub
Sub SetupToolbar()
Dim i As Integer

    'Specify whether Edit Mode is to remain on when switching between cells
    vaSpread1.EditModePermanent = True

    vaSpread1.Col = -1
    vaSpread1.Row = -1
    vaSpread1.Lock = True
    
    'Set the number of rows in the spreadsheet
    vaSpread1.MaxRows = 1
 
    'Set the height of a selected row
    vaSpread1.RowHeight(1) = 15
   
    'Set the number of columns in the spreadsheet
    vaSpread1.MaxCols = 17
 
    'Set the column widths
    For i = 1 To vaSpread1.MaxCols Step 2
        vaSpread1.ColWidth(i) = 0.3
    Next i
   
    'Resize wide column
    vaSpread1.ColWidth(14) = 15
    
    'Show or hide the column headers
    vaSpread1.DisplayColHeaders = False
    vaSpread1.DisplayRowHeaders = False
    
    'Turn off scroll bars
    vaSpread1.ScrollBars = ScrollBarsNone
    
    'Turn off border
    vaSpread1.BorderStyle = BorderStyleNone
      
    'Select row(s)
    vaSpread1.Row = 1
    vaSpread1.Col = -1

    'Determine the color of background, foreground and border color
    vaSpread1.ForeColor = RGB(0, 0, 0)
    vaSpread1.BackColor = RGB(192, 192, 192)
    vaSpread1.fontname = "MS Sans Serif"
    vaSpread1.FontSize = 8
    vaSpread1.FontBold = False
    
    'Select a single cell
    vaSpread1.Col = 2
    vaSpread1.Row = 1

    'Define cells as type BUTTON
    vaSpread1.CellType = CellTypeButton
    vaSpread1.Lock = False
    vaSpread1.TypeButtonText = "Print"
    'Set vaSpread1.TypeButtonPicture = LoadPicture(StartPath & "..\..\..\..\..\files\RIGHT.BMP")
    vaSpread1.TypeButtonAlign = TypeButtonAlignLeft
    
    'Select a single cell
    vaSpread1.Col = 4
    vaSpread1.Row = 1

    'Define cells as type BUTTON
    vaSpread1.CellType = CellTypeButton
    vaSpread1.Lock = False
    vaSpread1.TypeButtonText = "Next"
    'Set vaSpread1.TypeButtonPicture = LoadPicture(StartPath & "..\..\..\..\..\files\LEFT.BMP")
    vaSpread1.TypeButtonAlign = TypeButtonAlignRight
    
    'Select a single cell
    vaSpread1.Col = 6
    vaSpread1.Row = 1

    'Define cells as type BUTTON
    vaSpread1.CellType = CellTypeButton
    vaSpread1.Lock = False
    vaSpread1.TypeButtonText = "Previous"
    'Set vaSpread1.TypeButtonPicture = LoadPicture(App.Path & "..\..\..\..\..\files\ZOOM.BMP")
    vaSpread1.TypeButtonAlign = TypeButtonAlignRight
    
    'Select a single cell
    vaSpread1.Col = 8
    vaSpread1.Row = 1

    'Define cells as type BUTTON
    vaSpread1.CellType = CellTypeButton
    vaSpread1.Lock = False
    vaSpread1.TypeButtonText = "Zoom"
    'Set vaSpread1.TypeButtonPicture = LoadPicture(App.Path & "..\..\..\..\..\files\PRINT.BMP")
    vaSpread1.TypeButtonAlign = TypeButtonAlignRight
    
    'Select a single cell
    vaSpread1.Col = 10
    vaSpread1.Row = 1

    'Define cells as type BUTTON
    vaSpread1.CellType = CellTypeButton
    vaSpread1.Lock = False
    vaSpread1.TypeButtonText = "Setup"
    'Set vaSpread1.TypeButtonPicture = LoadPicture(App.Path & "..\..\..\..\..\files\SETUP.BMP")
    vaSpread1.TypeButtonAlign = TypeButtonAlignRight
    
    
    'Select a single cell
    vaSpread1.Col = 16
    vaSpread1.Row = 1

    'Define cells as type BUTTON
    vaSpread1.CellType = CellTypeButton
    vaSpread1.Lock = False
    vaSpread1.TypeButtonText = "Close"
    'Set vaSpread1.TypeButtonPicture = LoadPicture(App.Path & "..\..\..\..\..\files\CLOSE.BMP")
    vaSpread1.TypeButtonAlign = TypeButtonAlignRight
    vaSpread1.TextTip = TextTipFloating
    Dim bRet As Boolean
    bRet = vaSpread1.SetTextTipAppearance("MS Sans Serif", 8, 0, 0, &HC0FFFF, &H0)
    vaSpread1.CursorType = CursorTypeLockedCell
    vaSpread1.CursorStyle = CursorStyleArrow
    vaSpread1.NoBeep = True
    
End Sub
Sub DisableButton(Col As Long, bitmapdirection As String)
'Disable specified button
    vaSpread1.Redraw = False
    
    vaSpread1.Row = 1
    vaSpread1.Col = Col
    
    vaSpread1.Lock = True
    vaSpread1.TypeButtonTextColor = RGB(128, 128, 128)
    vaSpread1.Protect = True
    'Set vaSpread1.TypeButtonPicture = LoadPicture(App.Path & "..\..\..\..\..\files\" & bitmapdirection & "DIS.BMP")
    
    vaSpread1.Redraw = True
End Sub
Sub EnableButton(Col As Long, bitmapdirection As String)
'Enable specified button
    vaSpread1.Redraw = False
    
    vaSpread1.Row = 1
    vaSpread1.Col = Col
    
    vaSpread1.Lock = False
    vaSpread1.TypeButtonTextColor = RGB(0, 0, 0)
    vaSpread1.Protect = False
    'Set vaSpread1.TypeButtonPicture = LoadPicture(App.Path & "..\..\..\..\..\files\" & bitmapdirection & ".BMP")
    
    vaSpread1.Redraw = True
End Sub

'Private Sub Form_Resize()
'    vaSpread1.Move 0, 0, ScaleWidth, vaSpread1.Height
'    fpSpreadPreview1.Move 0, vaSpread1.Height, ScaleWidth, ScaleHeight - vaSpread1.Height
'End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

    vaSpread1.Col = Col
    vaSpread1.Row = Row
    
    If vaSpread1.CellType = CellTypeButton Then
        Select Case Col
            Case 2  'Print
                PrintDlg.Show
                 'CommonDialog1.ShowPrinter
            Case 4  'Next
                If vaSpreadPreview1.PageCurrent < frmBudPrepMaint.vaSpread1.PrintPageCount Then
                    vaSpreadPreview1.PageCurrent = vaSpreadPreview1.PageCurrent + vaSpreadPreview1.PagesPerScreen
                    EnableButton Col, "RIGHT"
                    'Enable Previous button
                    EnableButton 6, "LEFT"
                   'Update page count listing
                    UpdatePageCount
                End If
                
                 'If at last page, disable button
                    If vaSpreadPreview1.PageCurrent >= frmBudPrepMaint.vaSpread1.PrintPageCount Then
                        DisableButton Col, "RIGHT"
                    End If
                
            Case 6 'Previous
                If vaSpreadPreview1.PageCurrent > 1 Then
                    vaSpreadPreview1.PageCurrent = vaSpreadPreview1.PageCurrent - vaSpreadPreview1.PagesPerScreen
                    EnableButton Col, "LEFT"
                    EnableButton 4, "RIGHT"
                    'Update page count listing
                    UpdatePageCount
                End If
                
                'If at first page, disable button
                If vaSpreadPreview1.PageCurrent = 1 Then
                    DisableButton Col, "LEFT"
                End If
                
            Case 8 'Zoom
                vaSpreadPreview1.ZoomState = 3
            Case 10 'Setup
                pagesetup.Show 1
             
            Case 16 'Close
                Unload Me
        End Select
    End If
End Sub
Sub UpdatePageCount()
 'Page Count
    vaSpread1.Row = 1
    vaSpread1.Col = 14
    vaSpread1.Text = "Page " & vaSpreadPreview1.PageCurrent & " of " & frmBudPrepMaint.vaSpread1.PrintPageCount
    
End Sub

Private Sub vaSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    With vaSpread1
        .Col = Col
        .Row = Row
        If .CellType = CellTypeButton And Not .Lock Then
            ShowTip = True
            TipText = .TypeButtonText
        ElseIf .CellType = CellTypeEdit And .Text <> "" Then
            ShowTip = True
            TipText = .Text
        End If
    End With
End Sub

Private Sub vaSpreadPreview1_PageChange(ByVal Page As Long)
    UpdatePageCount
End Sub


Public Sub GetZoom(zoomlabel As Integer)
'Set up the print previews zoom

        Select Case zoomlabel
            Case 0
                vaSpreadPreview1.PageViewType = 2
                vaSpreadPreview1.PageViewPercentage = 200
            
            Case 1
                vaSpreadPreview1.PageViewType = 2
                vaSpreadPreview1.PageViewPercentage = 150

            Case 2
                vaSpreadPreview1.PageViewType = 2
                vaSpreadPreview1.PageViewPercentage = 100

            Case 3
                vaSpreadPreview1.PageViewType = 2
                vaSpreadPreview1.PageViewPercentage = 75

            Case 4
                vaSpreadPreview1.PageViewType = 2
                vaSpreadPreview1.PageViewPercentage = 50

            Case 5
                vaSpreadPreview1.PageViewType = 2
                vaSpreadPreview1.PageViewPercentage = 25

            Case 6
                vaSpreadPreview1.PageViewType = 2
                vaSpreadPreview1.PageViewPercentage = 10

            Case 7
                vaSpreadPreview1.PageViewType = 3
                
            Case 8
                vaSpreadPreview1.PageViewType = 4
                
            Case 9
                vaSpreadPreview1.PageViewType = 0
                
            Case 10
                vaSpreadPreview1.PageViewType = 5
                vaSpreadPreview1.PageMultiCntH = 2
                vaSpreadPreview1.PageMultiCntV = 1
                
            Case 11
                vaSpreadPreview1.PageViewType = 5
                vaSpreadPreview1.PageMultiCntH = 3
                vaSpreadPreview1.PageMultiCntV = 1
                
            Case 12
                vaSpreadPreview1.PageViewType = 5
                vaSpreadPreview1.PageMultiCntH = 2
                vaSpreadPreview1.PageMultiCntV = 2
                
            Case 13
                vaSpreadPreview1.PageViewType = 5
                vaSpreadPreview1.PageMultiCntH = 3
                vaSpreadPreview1.PageMultiCntV = 2

        End Select
      
End Sub

