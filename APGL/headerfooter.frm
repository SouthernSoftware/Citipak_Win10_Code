VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Begin VB.Form headerfooter 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Change Header/Footer Attributes"
   ClientHeight    =   5868
   ClientLeft      =   -96
   ClientTop       =   540
   ClientWidth     =   9672
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "headerfooter.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5868
   ScaleWidth      =   9672
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   2244
      Left            =   288
      TabIndex        =   15
      Top             =   2736
      Width           =   9228
      _Version        =   196613
      _ExtentX        =   16277
      _ExtentY        =   3958
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
      ShadowText      =   0
      SpreadDesigner  =   "headerfooter.frx":08CA
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00D0D0D0&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8100
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00D0D0D0&
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Apply to"
      Height          =   975
      Left            =   300
      TabIndex        =   0
      Top             =   1680
      Width           =   9195
      Begin VB.CommandButton Command4 
         BackColor       =   &H00D0D0D0&
         Caption         =   "Reset Text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7680
         MaskColor       =   &H00D0D0D0&
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   420
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00D0D0D0&
         Caption         =   "Save Header Text"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         MaskColor       =   &H00D0D0D0&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   420
         Width           =   1695
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "FooterText"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   660
         TabIndex        =   4
         Top             =   600
         Width           =   1275
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Header Text"
         Height          =   255
         Index           =   0
         Left            =   660
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Press the Save button before changing the 'Apply to' to avoid losing your changes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   4080
         TabIndex        =   11
         Top             =   360
         Width           =   3075
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5520
      Top             =   5145
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Right click the mouse to insert/remove page numbering to the highlighted cell"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   2520
      TabIndex        =   14
      Top             =   5100
      Width           =   2475
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "To Insert Page Numbers:"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   13
      Top             =   5100
      Width           =   2235
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "- You can change the font characteristics by pressing the button next to the selected cell"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   900
      TabIndex        =   9
      Top             =   1320
      Width           =   8235
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "- Text in the first column will be left justified, center column will be centered, and the right column will be right justified"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   900
      TabIndex        =   8
      Top             =   1020
      Width           =   8235
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "2. Enter your text below and press the OK button to save your changes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   7
      Top             =   720
      Width           =   5235
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1. Choose to apply text to the Header or Footer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   6
      Top             =   420
      Width           =   5235
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  To customize header or footer text:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "headerfooter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fpfontbold As Integer
Dim fpfontunderline As Integer
Dim fpfontstrikethru As Integer
Dim fpfontname As String
Dim fpfontsize As Integer
Dim fpfontitalic As Integer

Sub SaveConfiguration()
'Save the current font/text configuration
Dim i As Integer, j As Integer
Dim headerstring As String

    headerstring = ""
    vaSpread1.EditMode = False
    DoEvents
    
    'Build font configurations
    'Loop through all rows
    For i = 1 To vaSpread1.DataRowCnt
        vaSpread1.Row = i
        'Loop through the three data columns
        For j = 1 To 6 Step 2
            vaSpread1.Col = j
            'if contains header text
            If vaSpread1.Text <> "" Then
                'Get cell characteristics
                headerstring = headerstring & GetCellData(j, i)
            End If
        Next j
        
        headerstring = headerstring & "/n"
        
    Next i
    
    'Header string
    If Option3(0).Value = True Then
        frmBudPrepMaint.vaSpread1.PrintHeader = headerstring
    'Footer string
    Else
        frmBudPrepMaint.vaSpread1.PrintFooter = headerstring
    End If
    
End Sub
Private Sub Command1_Click()
'OK
    SaveConfiguration
    Unload Me
End Sub

Private Sub Command2_Click()
'Cancel
    Unload Me
End Sub

Sub InitSpread()
'Set up spread
Dim i As Integer
    vaSpread1.FontBold = False
    
    'Specify whether Edit Mode is to remain on when switching between cells
    'Remove text in headers
    vaSpread1.ColHeaderDisplay = DispBlank
    
    'Sets Border Appearance
    vaSpread1.Appearance = Appearance3D
 
    'Set the number of columns in the spreadsheet
    vaSpread1.MaxCols = 6
    
    'Set the width of a selected column
    vaSpread1.ColWidth(0) = 5#
    
    'Set the number of rows in the spreadsheet
    vaSpread1.MaxRows = 9
    
    'Define grid type and style for the spreadsheet
    vaSpread1.GridShowHoriz = False
    vaSpread1.GridShowVert = False
    vaSpread1.GridSolid = False

    'Show or hide the column headers
    vaSpread1.DisplayColHeaders = True
    vaSpread1.Row = 0
    vaSpread1.Col = 1
  
    'Show or hide the row headers
    vaSpread1.DisplayRowHeaders = True

    'Determine if the horz and/or the vert scroll bars are to be displayed
    vaSpread1.ScrollBars = ScrollBarsVertical
      
    'Make button cell types
    For i = 2 To 6 Step 2
        vaSpread1.Col = i
        vaSpread1.Row = -1
        vaSpread1.FontBold = True
        vaSpread1.CellType = CellTypeButton
        vaSpread1.TypeButtonText = "..."
    Next i
    'only show buttons for active row
    vaSpread1.ButtonDrawMode = 4
    
    'Data cells
    For i = 1 To 6 Step 2
        vaSpread1.ColWidth(i) = 20#
    Next i
    
    'Change row header text
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Col = 0
        vaSpread1.Row = i
        vaSpread1.Text = "Line " & i
    Next i
    
    'Set Col Header Text
    vaSpread1.Row = 0
    vaSpread1.Col = 1
    vaSpread1.Text = "Left Justify"
    vaSpread1.Col = 3
    vaSpread1.Text = "Center"
    vaSpread1.Col = 5
    vaSpread1.Text = "Right Justify"
    
    vaSpread1.ProcessTab = True
    
    'Set Borders
    For i = 2 To 6 Step 2
        Call vaSpread1.SetCellBorder(i, 1, i, vaSpread1.MaxRows, 2, RGB(192, 192, 192), CellBorderStyleDot)
        vaSpread1.ColWidth(i) = 3#
    Next i
    
End Sub

Function GetCellData(Col As Integer, Row As Integer) As String
'Checks the font char. for the cell and builds cell header string
Dim celldata As String

    celldata = ""
    
    vaSpread1.Col = Col
    vaSpread1.Row = Row
    
    'font name
    celldata = "/fn" & Chr$(34) & vaSpread1.fontname & Chr$(34)
    
    'font size
    celldata = celldata & "/fz" & Chr$(34) & vaSpread1.FontSize & Chr$(34)
    
    'font bold
    If vaSpread1.FontBold = True Then
        celldata = celldata & "/fb1"
    Else
        celldata = celldata & "/fb0"
    End If
    
    'font italic
    If vaSpread1.FontItalic = True Then
        celldata = celldata & "/fi1"
    Else
        celldata = celldata & "/fi0"
    End If
    
    'font underline
    If vaSpread1.FontUnderline = True Then
        celldata = celldata & "/fu1"
    Else
        celldata = celldata & "/fu0"
    End If
    
    'font strikethru
    If vaSpread1.FontStrikethru = True Then
        celldata = celldata & "/fk1"
    Else
        celldata = celldata & "/fk0"
    End If
    
    'add justify info
    If Col = 1 Then
        'left justify
        celldata = celldata & "/l"
    ElseIf Col = 3 Then
        'center
        celldata = celldata & "/c"
    ElseIf Col = 5 Then
        'right justify
        celldata = celldata & "/r"
    End If
    
    'add text, if not page number
    If vaSpread1.Text <> "<Page Number>" Then
        celldata = celldata & vaSpread1.Text
    Else    'page numbering
        celldata = celldata & "/p"
    End If
      
    'Send back
    GetCellData = celldata
    
End Function

Private Sub Command3_Click()
    SaveConfiguration
    Command3.Enabled = False
End Sub

Private Sub Command4_Click()
Dim choice As Integer
    
    choice = MsgBox("Are you sure you want to clear the text and formatting for the selected header?", 4, "Clear Text")
    
    If choice = 6 Then  'Yes
        If Option3(0).Value = True Then
            frmViewBud.vaSpread1.PrintHeader = ""
            'Reset the display
            GetHeaderDetail frmViewBud.vaSpread1.PrintHeader
        Else
            frmViewBud.vaSpread1.PrintFooter = ""
            'Reset the display
            GetHeaderDetail frmViewBud.vaSpread1.PrintFooter
        End If
    End If
    
End Sub

Private Sub Form_Load()
  
    'Initialize Spread for entering headers
    InitSpread

    'get the header and/or footer text

    GetHeaderDetail frmViewBud.vaSpread1.PrintHeader
    
End Sub

Sub GetHeaderDetail(detailtext As String)
'Reads the current header/footer and displays it in Spread
Dim startpoint As Integer
Dim fontstart As Integer
Dim fontname As String

    'Clear spread text
    Call vaSpread1.ClearRange(1, 1, vaSpread1.MaxCols, vaSpread1.MaxRows, False)
    
    'if no header text, exit
    If detailtext = "" Then
        Exit Sub
    End If
    
    vaSpread1.Row = 1
    vaSpread1.Col = 1
'/fn"MS Sans Serif"/fz"12"/fb1/fi0/fu0/fk0/ltretre/fn"MS Sans Serif"/fz"8.25"/fb0/fi0/fu0/fk0/ctretre/fn"MS Sans Serif"/fz"8.25"/fb0/fi0/fu0/fk0/rtreter/n/fn"MS Sans Serif"/fz"8.25"/fb0/fi0/fu0/fk0/ctretre/n/fn"MS Sans Serif"/fz"8.25"/fb0/fi0/fu0/fk0/ctretre/n
'/fn,/fz,/fbx,/fix,/fux,/fkx,/n,/<justify>
 
    startpoint = 1
    fontstart = 1
    'loop through string
    Do
        If Mid$(detailtext, startpoint, 1) = "/" Then
            startpoint = startpoint + 1
            Select Case Mid$(detailtext, startpoint, 1)
                Case "n"    'new line
                    vaSpread1.Row = vaSpread1.Row + 1
                Case "l"    'left justify
                    AddCellData 1
                Case "r"    'right justify
                    AddCellData 5
                Case "c"    'center
                    AddCellData 3
                Case "p"    'page numbering
                    'Make static cell
                    vaSpread1.CellType = CellTypeStaticText
                    vaSpread1.ForeColor = &H808080     'gray
                    vaSpread1.Text = "<Page Number>"
                Case "f"    'font
                    startpoint = startpoint + 1
                    Select Case Mid$(detailtext, startpoint, 1)
                    '/fn,/fz,/fbx,/fix,/fux,/fkx,/n,/<justify>
                        Case "n"    'font name: /fn"MS Sans Serif"
                            startpoint = startpoint + 2 'beginning of font name
                            fontstart = startpoint
                            'repeat until found end quote
                            While Mid$(detailtext, startpoint, 1) <> """"
                                startpoint = startpoint + 1
                                'error check
                                If startpoint - fontstart = 100 Then
                                    MsgBox "Error parsing FONT NAME", 0, "GetHeaderDetail"
                                    Exit Sub
                                End If
                            Wend
'                            fpSpread1.fontname = Mid$(detailtext, fontstart, startpoint - fontstart - 1)
                            fpfontname = Mid$(detailtext, fontstart, startpoint - fontstart)
                            
                        Case "z"    'font size: /fz"12"
                            startpoint = startpoint + 2 'beginning of font size
                            fontstart = startpoint
                            'repeat until found end quote
                            While Mid$(detailtext, startpoint, 1) <> """"
                                startpoint = startpoint + 1
                                'error check
                                If startpoint - fontstart = 100 Then
                                    MsgBox "Error parsing FONT SIZE", 0, "GetHeaderDetail"
                                    Exit Sub
                                End If
                            Wend
                            'fpSpread1.FontSize = Mid$(detailtext, fontstart, startpoint - fontstart - 1)
                            fpfontsize = Mid(detailtext, fontstart, startpoint - fontstart)
                            
                        Case "b"    'bold
                            startpoint = startpoint + 1
                            If Mid$(detailtext, startpoint, 1) = "0" Then
                                'fpSpread1.fontbold = False
                                fpfontbold = False
                            Else
                                'fpSpread1.fontbold = True
                                fpfontbold = True
                            End If
                        Case "i"    'italic
                            startpoint = startpoint + 1
                            If Mid$(detailtext, startpoint, 1) = "0" Then
                                'fpSpread1.FontItalic = False
                                fpfontitalic = False
                            Else
                                'fpSpread1.FontItalic = True
                                fpfontitalic = True
                            End If
                        Case "u"    'underline
                            startpoint = startpoint + 1
                            If Mid$(detailtext, startpoint, 1) = "0" Then
                                'fpSpread1.FontUnderline = False
                                fpfontunderline = False
                            Else
                                'fpSpread1.FontUnderline = True
                                fpfontunderline = True
                            End If
                        Case "k"    'strikethru
                            startpoint = startpoint + 1
                            If Mid$(detailtext, startpoint, 1) = "0" Then
                                'fpSpread1.FontStrikethru = False
                                fpfontstrikethru = False
                            Else
                                'fpSpread1.FontStrikethru = True
                                fpfontstrikethru = True
                            End If
                    End Select  'font
            End Select  'mainFrm
        
        'found text
        Else
            fontstart = startpoint
            'loop until found "/"
            While Mid$(detailtext, startpoint, 1) <> "/"
                startpoint = startpoint + 1
                'error check
                If startpoint - fontstart = 100 Then
                    MsgBox "Error parsing TEXT", 0, "GetHeaderDetail"
                    Exit Sub
                End If
            Wend
            vaSpread1.Text = Mid$(detailtext, fontstart, startpoint - fontstart)
            'decrement startpoint for next read
            startpoint = startpoint - 1
        
        End If  'finding "/"
        
        'increment counter
        startpoint = startpoint + 1
        
    Loop Until startpoint >= Len(detailtext)
    
End Sub
Sub AddCellData(Col As Long)
'Adds the cell font info. to the specified col
'Needed to wait to receive the column number info.

    vaSpread1.Col = Col
    
    vaSpread1.fontname = fpfontname
    vaSpread1.FontSize = fpfontsize
    vaSpread1.FontBold = fpfontbold
    vaSpread1.FontUnderline = fpfontunderline
    vaSpread1.FontItalic = fpfontitalic
    vaSpread1.FontStrikethru = fpfontstrikethru

End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
      CommonDialog1.Flags = cdlCFBoth Or cdlCFEffects
      
      vaSpread1.Col = Col - 1
      vaSpread1.Row = Row
      
      CommonDialog1.fontname = vaSpread1.Font.Name
      CommonDialog1.FontSize = vaSpread1.Font.Size
      CommonDialog1.FontBold = vaSpread1.Font.Bold
      CommonDialog1.FontItalic = vaSpread1.Font.Italic
      CommonDialog1.FontUnderline = vaSpread1.Font.Underline
      CommonDialog1.FontStrikethru = vaSpread1.Font.Strikethrough
      CommonDialog1.Color = vaSpread1.ForeColor
      
      CommonDialog1.ShowFont   ' Display Font common dialog box.
      
      
      ' Set Cancel to True.   CommonDialog1.CancelError = True
      On Error GoTo ErrHandler   ' Set the Flags property.
      vaSpread1.Font.Name = CommonDialog1.fontname
      vaSpread1.Font.Size = CommonDialog1.FontSize
      vaSpread1.Font.Bold = CommonDialog1.FontBold
      vaSpread1.Font.Italic = CommonDialog1.FontItalic
      vaSpread1.Font.Underline = CommonDialog1.FontUnderline
      vaSpread1.FontStrikethru = CommonDialog1.FontStrikethru
ErrHandler:
End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    If Mode = 1 Then
        If Command3.Enabled = False Then
            Command3.Enabled = True
        End If
    End If
End Sub

Private Sub vaSpread1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'Insert/Remove page numbering
Dim choice As Integer

    If Button = 2 Then  'right mouse button
        'Make sure valid column
        Select Case vaSpread1.ActiveCol
            Case 1, 3, 5
                vaSpread1.Row = vaSpread1.ActiveRow
                vaSpread1.Col = vaSpread1.ActiveCol
                Select Case vaSpread1.Text
                    Case ""     'Insert text
                        'Make static cell
                        vaSpread1.CellType = CellTypeStaticText
                        vaSpread1.ForeColor = &H808080     'gray
                        vaSpread1.Text = "<Page Number>"
                    Case "<Page Number>"    'Remove numbering
                        'Make edit cell
                        vaSpread1.CellType = CellTypeEdit
                        vaSpread1.ForeColor = RGB(0, 0, 0)
                        vaSpread1.Text = ""
                    Case Else
                        choice = MsgBox("Inserting page numbering in this cell will overwrite your exiting text.  Do you want to replace the existing text with page numbers?", 36, "Replace Existing Text")
                        If choice = 6 Then  'yes
                            'Make static cell
                            vaSpread1.CellType = CellTypeStaticText
                            vaSpread1.ForeColor = &H808080     'gray
                            vaSpread1.Text = "<Page Number>"
                        End If
                End Select
        End Select
    End If
End Sub

Private Sub Option3_Click(Index As Integer)
Dim choice As Integer

    If Index = 0 Then   'Header
        Option3(0).FontBold = True
        Option3(1).FontBold = False
        GetHeaderDetail frmViewBud.vaSpread1.PrintHeader
        Command3.Caption = "Save Header Text"
                
    Else    'Footer
        Option3(0).FontBold = False
        Option3(1).FontBold = True
        GetHeaderDetail frmViewBud.vaSpread1.PrintFooter
        Command3.Caption = "Save Footer Text"
    End If
End Sub


