VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmTCMainMenu1 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu vs 2.05"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
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
   Icon            =   "frmTCMainMenu1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   450
      Left            =   4080
      TabIndex        =   0
      Top             =   7095
      Width           =   3630
      _Version        =   131072
      _ExtentX        =   6403
      _ExtentY        =   794
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
      ButtonDesigner  =   "frmTCMainMenu1.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMatch 
      Height          =   450
      Left            =   4080
      TabIndex        =   1
      Top             =   5085
      Width           =   3630
      _Version        =   131072
      _ExtentX        =   6403
      _ExtentY        =   794
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
      ButtonDesigner  =   "frmTCMainMenu1.frx":0AE9
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdNew 
      Height          =   450
      Left            =   4080
      TabIndex        =   2
      Top             =   3045
      Width           =   3630
      _Version        =   131072
      _ExtentX        =   6403
      _ExtentY        =   794
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
      ButtonDesigner  =   "frmTCMainMenu1.frx":0D09
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdConvert 
      Height          =   450
      Left            =   4080
      TabIndex        =   3
      Top             =   6090
      Width           =   3630
      _Version        =   131072
      _ExtentX        =   6403
      _ExtentY        =   794
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
      ButtonDesigner  =   "frmTCMainMenu1.frx":0F2A
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrint 
      Height          =   450
      Left            =   4080
      TabIndex        =   4
      Top             =   5580
      Width           =   3630
      _Version        =   131072
      _ExtentX        =   6403
      _ExtentY        =   794
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
      ButtonDesigner  =   "frmTCMainMenu1.frx":1148
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdResults 
      Height          =   450
      Left            =   4080
      TabIndex        =   5
      Top             =   6585
      Width           =   3630
      _Version        =   131072
      _ExtentX        =   6403
      _ExtentY        =   794
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
      ButtonDesigner  =   "frmTCMainMenu1.frx":1366
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdInstructions 
      Height          =   450
      Left            =   4080
      TabIndex        =   6
      Top             =   2535
      Width           =   3630
      _Version        =   131072
      _ExtentX        =   6403
      _ExtentY        =   794
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
      ButtonDesigner  =   "frmTCMainMenu1.frx":1580
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdChangeCoNums 
      Height          =   450
      Left            =   4080
      TabIndex        =   8
      ToolTipText     =   "If a customer has no real and no personal property then add ""Old"" to the county number."
      Top             =   3540
      Width           =   3630
      _Version        =   131072
      _ExtentX        =   6403
      _ExtentY        =   794
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
      ButtonDesigner  =   "frmTCMainMenu1.frx":1798
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAppendwithPP 
      Height          =   450
      Left            =   4080
      TabIndex        =   9
      ToolTipText     =   "Works in VA only if the customer has personal property, no real property and the county number is a string."
      Top             =   4050
      Width           =   3630
      _Version        =   131072
      _ExtentX        =   6403
      _ExtentY        =   794
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
      ButtonDesigner  =   "frmTCMainMenu1.frx":1985
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClearCntyNums 
      Height          =   450
      Left            =   4080
      TabIndex        =   10
      Top             =   4575
      Width           =   3630
      _Version        =   131072
      _ExtentX        =   6403
      _ExtentY        =   794
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
      ButtonDesigner  =   "frmTCMainMenu1.frx":1B70
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1104
      Index           =   1
      Left            =   1500
      Top             =   840
      Width           =   8652
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VA TAX CONVERSION MAIN MENU"
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
      Left            =   2820
      TabIndex        =   7
      Top             =   1200
      Width           =   6012
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   132
      Left            =   2160
      Top             =   2052
      Width           =   972
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2280
      X2              =   2280
      Y1              =   2160
      Y2              =   8048
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   132
      Left            =   8520
      Top             =   2052
      Width           =   972
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8640
      X2              =   8640
      Y1              =   2160
      Y2              =   8048
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1500
      Top             =   720
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   3
      Left            =   2160
      Top             =   1920
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5916
      Index           =   0
      Left            =   2280
      Top             =   2148
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   8520
      Top             =   1920
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5916
      Index           =   1
      Left            =   8640
      Top             =   2148
      Width           =   732
   End
End
Attribute VB_Name = "frmTCMainMenu1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdAppendwithPP_Click()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim cnt As Long
  Dim CntIt As Boolean
  Dim BadCnt As Long
  
  If Exist("TaxCust.dat") Then
    OpenTaxCustFile TCHandle, NumOfTCRecs
    For x = 1 To NumOfTCRecs
      Get TCHandle, x, TaxCust
      CntIt = False
      If TaxCust.FirstPersRec > 0 And TaxCust.FirstPropRec = 0 Then
        If Len(QPTrim$(TaxCust.CountyAcctString)) > 16 Then
          BadCnt = BadCnt + 1
          GoTo NextOne
        End If
        If QPTrim$(TaxCust.CountyAcctString) <> "" Or TaxCust.CountyAcct <> 0 Then 'added CountyAcct on 5/15/07
          If QPTrim$(TaxCust.CountyAcctString) <> "" Then
            TaxCust.CountyAcctString = QPTrim$(TaxCust.CountyAcctString) + "PP"
            CntIt = True
          ElseIf TaxCust.CountyAcct <> 0 Then
            TaxCust.CountyAcctString = CStr(TaxCust.CountyAcct) & "PP"
            TaxCust.CountyAcct = 0
          End If
          Put TCHandle, x, TaxCust
        End If
        If CntIt = True Then cnt = cnt + 1
      End If
NextOne:
    Next x
  Else
    Call TCMsg(800, "The file 'TAXCUST.DAT' could not be found.")
    Exit Sub
  End If
  
  If BadCnt = 0 Then
    Call TCMsg(750, "A total of " + CStr(cnt) + " county account numbers were modified successfully.")
  Else
    Call TCMsg(800, "A total of " + CStr(cnt) + " county account numbers were modified successfully and a total of " + CStr(BadCnt) + " were too long to append.")
  End If
  
End Sub

Private Sub cmdChangeCoNums_Click()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim cnt As Long
  Dim CntIt As Boolean
  If Exist("TaxCust.dat") Then
    OpenTaxCustFile TCHandle, NumOfTCRecs
    For x = 1 To NumOfTCRecs
      Get TCHandle, x, TaxCust
      CntIt = False
      If TaxCust.FirstPersRec = 0 And TaxCust.FirstPropRec = 0 Then
        If TaxCust.CountyAcct > 0 Then
          TaxCust.CountyAcct = -TaxCust.CountyAcct
          CntIt = True
        End If
        If QPTrim$(TaxCust.CountyAcctString) <> "" Then
          TaxCust.CountyAcctString = "old" + TaxCust.CountyAcctString
          CntIt = True
        End If
        If CntIt = True Then cnt = cnt + 1
      End If
    Next x
  Else
    Call TCMsg(800, "The file 'TAXCUST.DAT' could not be found.")
    Exit Sub
  End If
    
  Call TCMsg(800, "A total of " + CStr(cnt) + " county account numbers were modified successfully.")
End Sub

Private Sub cmdClearPCntyNums_Click()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim cnt As Long
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TaxCust.FirstPersRec > 0 Then
      TaxCust.CountyAcct = 0
      Put TCHandle, x, TaxCust
      cnt = cnt + 1
    End If
  Next x
  
  Close TCHandle
  Call TCMsg(800, "A total of " + CStr(cnt) + " county numbers have been removed successfully.")

End Sub

Private Sub cmdClearCntyNums_Click()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim cnt As Long
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    TaxCust.CountyAcct = 0
    Put TCHandle, x, TaxCust
      cnt = cnt + 1
  Next x
  
  Close TCHandle
  Call TCMsg(800, "A total of " + CStr(cnt) + " county numbers have been removed successfully.")

End Sub


Private Sub cmdConvert_Click()
  WhichOne = "B"
  If Not Exist("TAXCUST.DAT") Then
    Call TCMsg(900, "The file 'TAXCUST.DAT' could not be found. Process aborted.")
    Exit Sub
  Else
    If TCMsgWOpts(900, "Please be sure the 'TAXCUST.DAT' has not already been converted.", "F10 Continue", "ESC Exit") = "abort" Then
      Exit Sub
    End If
  End If
  
  If Not Exist("TAXPERS.DAT") And Not Exist("TAXPROP.DAT") Then
    Call TCMsg(800, "Neither the file 'TAXPERS.DAT' nor the file 'TAXPROP.DAT' could be found. Process aborted.")
    Exit Sub
  End If
  
  If Not Exist("TAXPERS.DAT") Then
    Call TCMsg(800, "The file 'TAXPERS.DAT' could not be found. Process will default to 'Real Only'.")
    WhichOne = "R"
  End If
  
  If Not Exist("TAXPROP.DAT") Then
    Call TCMsg(800, "The file 'TAXPROP.DAT' could not be found. Process will default to 'Personal Only'.")
    WhichOne = "P"
  End If
  
  frmTCConvert1.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdExit_Click()
  End
End Sub

Private Sub cmdInstructions_Click()
  frmTCInstructions.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdMatch_Click()
  
  If Not Exist(App.Path + "\ParcelsText.csv") Then
    Call TCMsg(900, "The required file 'ParcelsText.csv' cannot be located. Matching up data is aborted.")
    Exit Sub
  End If
  
  frmTCMatchUp1.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdNew_Click()
  If Not Exist("ParcelsText.csv") Then
    Call TCMsg(900, "The required file 'ParcelsText.csv' cannot be located. Clearing existing data is aborted.")
    Exit Sub
  End If
  
  If Exist(ConvSpreadFile) And Exist(ConversionFile) Then
    If TCMsgWOpts(900, "Are you sure you want to clear existing data? Press F10 to clear. Otherwise, press ESC to exit.", "F10 Clear", "ESC Exit") = "abort" Then
      Exit Sub
    End If
  End If
  
  KillFile ConvSpreadFile
  KillFile ConversionFile
  KillFile ConvResults
  KillFile ConvErrors
  
  Call Savemsg(900, "Clearing existing data has been completed successfully.")
  
End Sub

Private Sub cmdPrint_Click()
  Dim MaxLines As Integer
  Dim FF$
  Dim Page As Integer
  Dim x As Long
  Dim RptFile As String
  Dim RptHandle As Integer
  Dim TempHandle As Integer
  Dim TempRec As TempConversionData
  Dim NumOfTempRecs As Long
  Dim ThisCity As String * 20
  Dim CntyNum As String
  Dim AddDesc As String
  Dim CustCnt As Long
  Dim TRealVal As Double
  Dim TPersVal As Double
  
  If Not Exist(ConversionFile) Then
    Call TCMsg(900, "Please process the county data first. Load attempt aborted.")
    Exit Sub
  End If
  
  FF$ = Chr(12)
  MaxLines = 58
  RptFile$ = "TCRPTS\MATCH.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  OpenTempConvFile TempHandle, NumOfTempRecs
'  Print #RptHandle, "Customer Name"; Tab(51); "Property Pin#"; Tab(71); "County Number"; Tab(91); "City"; Tab(111); "Real Prop Value"; Tab(131); "Pers Prop Value"; Tab(151); "Address/Desc"; Tab(185); "Building Value"; Tab(200); "Search Name"; Tab(220); "PPTRA Y/N?"; Tab(232); "LOT NUMB"
  Print #RptHandle, "Customer Name"; Tab(51); "Real Prop Pin#"; Tab(75); "Pers Prop Pin"; Tab(95); "County Number"; Tab(117); "City"; Tab(137); "Real Prop Value"; Tab(157); "Pers Prop Value"; Tab(187); "Address/Desc" '; Tab(207); "Search Name"
  For x = 1 To NumOfTempRecs
    Get TempHandle, x, TempRec
      ThisCity = QPTrim$(TempRec.CData.City)
      TRealVal = OldRound(TRealVal + TempRec.CData.PROPVALU)
      TPersVal = OldRound(TPersVal + TempRec.CData.PersVal)
      If TempRec.CData.CountyAcct > 0 Then
        CntyNum = CStr(TempRec.CData.CountyAcct)
      ElseIf TempRec.CData.CountyAcctString <> "" Then
        CntyNum = TempRec.CData.CountyAcctString
      Else
        CntyNum = "Unknown"
      End If
      If TempRec.CData.RealAdd <> "" Then
        AddDesc = TempRec.CData.RealAdd
      ElseIf TempRec.CData.RDESC1 <> "" Then
        AddDesc = TempRec.CData.RDESC1
      ElseIf TempRec.CData.PDESC1 <> "" Then
        AddDesc = TempRec.CData.PDESC1
      End If
      TempRec.CData.State = TempRec.CData.State
      Print #RptHandle, QPTrim$(TempRec.CData.CustName); Tab(51); QPTrim$(TempRec.CData.RPinNum); Tab(75); QPTrim$(TempRec.CData.PPinNum); Tab(95); CntyNum; Tab(117); ThisCity; Tab(137); Using$("$###,###,##0.00", TempRec.CData.PROPVALU); Tab(157); Using$("$###,###,##0.00", TempRec.CData.PersVal); Tab(187); AddDesc '; Tab(207); " " + QPTrim$(TempRec.CData.SName)
      CustCnt = CustCnt + 1
  Next x
  Print #RptHandle,
  Print #RptHandle, "Number of Entries:  " + Using$("###,##0", CustCnt)
  Print #RptHandle, "Total Real Value:     " + Using$("$#,###,###,##0.00", TRealVal)
  Print #RptHandle, "Total Personal Value: " + Using$("$#,###,###,##0.00", TPersVal)
  Close
  ViewPrint RptFile, "Match Up Data", True
End Sub

Private Sub cmdResults_Click()
  Dim ConvRec As ConvResultsType
  Dim CRHandle As Integer
  Dim NumOfCRRecs As Long
  Dim x As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim ERptFile$
  Dim ERptHandle As Integer
  Dim dlm$
  Dim TRealVal As Double
  Dim TBldgVal As Double
  Dim TRealOXVal As Double
  Dim TRealSXVal As Double
  Dim TPersVal As Double
  Dim TMCVal As Double
  Dim TMHVal As Double
  Dim TMTVal As Double
  Dim TCVal As Double
  Dim TPersOXVal As Double
  Dim TPersSXVal As Double
  Dim GTPersVal As Double
  Dim ThisPersVal As Double
  Dim ErrorRec As ConvErrorType
  Dim EHandle As Integer
  Dim NumOfERecs As Long
  Dim TrunName As String * 20
  Dim ThisError$
  
  If Not Exist(ConvResults) Then
    Call TCMsg(900, "Please convert the county data first.")
    Exit Sub
  End If
  
  dlm = "~"
  
  RptFile$ = "TCRPTS\CSTLSTSM.RPT"
  
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  frmTCShowPctComp.Label1 = "Generating Report"
  frmTCShowPctComp.Show , Me
  
  OpenConvResultsFile CRHandle, NumOfCRRecs
  For x = 1 To NumOfCRRecs
    Get CRHandle, x, ConvRec
    TRealVal = OldRound(TRealVal + ConvRec.PROPVALU)
    TBldgVal = OldRound(TBldgVal + ConvRec.BLDGVAL)
    TRealOXVal = OldRound(TRealOXVal + ConvRec.REXMPOTHR)
    TRealSXVal = OldRound(TRealSXVal + ConvRec.REXMPSENI)
    TPersVal = OldRound(TPersVal + ConvRec.PersVal)
    TMCVal = OldRound(TMCVal + ConvRec.MCValue)
    TMHVal = OldRound(TMHVal + ConvRec.MHValue)
    TMTVal = OldRound(TMTVal + ConvRec.MTValue)
    TCVal = OldRound(TCVal + ConvRec.CVALUE)
    TPersOXVal = OldRound(TPersOXVal + ConvRec.PEXMPOTHR)
    TPersSXVal = OldRound(TPersSXVal + ConvRec.PEXMPSENI)
    ThisPersVal = OldRound(ConvRec.PersVal + ConvRec.MCValue + ConvRec.MHValue + ConvRec.MTValue + ConvRec.CVALUE)
    GTPersVal = OldRound(GTPersVal + ThisPersVal)
    TrunName = QPTrim$(ConvRec.CustName)
    '                             0                    1                                 2
    Print #RptHandle, QPTrim$(ConvRec.RPinNum); dlm; TrunName; dlm; QPTrim$(ConvRec.CountyAcctString); dlm;
    '                         3                         4                     5                       6
    Print #RptHandle, ConvRec.CountyAcct; dlm; ConvRec.PROPVALU; dlm; ConvRec.REXMPOTHR; dlm; ConvRec.REXMPSENI; dlm;
    '                         7                     8                    9                    10                    11
    Print #RptHandle, ConvRec.PersVal; dlm; ConvRec.MTValue; dlm; ConvRec.MCValue; dlm; ConvRec.CVALUE; dlm; ConvRec.MHValue; dlm;
    '                        12                      13                  14              15               16
    Print #RptHandle, ConvRec.PEXMPOTHR; dlm; ConvRec.PEXMPSENI; dlm; TRealVal; dlm; TRealOXVal; dlm; TRealSXVal; dlm;
    '                    17            18           19           20           21           22                23
    Print #RptHandle, TPersVal; dlm; TMCVal; dlm; TMHVal; dlm; TMTVal; dlm; TCVal; dlm; TPersOXVal; dlm; TPersSXVal; dlm;
    '                     24              25               26                    27                         28
    Print #RptHandle, GTPersVal; dlm; NumOfCRRecs; dlm; TBldgVal; dlm; QPTrim$(ConvRec.PPinNum); dlm; ConvRec.BLDGVAL
    
    frmTCShowPctComp.ShowPctComp x, NumOfCRRecs
    If frmTCShowPctComp.Out = True Then
      Close
      frmTCShowPctComp.Out = False
      Unload frmTCShowPctComp
      Exit Sub
    End If
  Next x
  
  Unload frmTCShowPctComp
  Close
  
  ERptFile$ = "TCRPTS\TCERRORS.RPT"
  ERptHandle = FreeFile
  Open ERptFile For Output As #ERptHandle
  
  OpenConvErrorsFile EHandle, NumOfERecs
  For x = 1 To NumOfERecs
    Get EHandle, x, ErrorRec
    If ErrorRec.ErrorType = 1 Then
      ThisError = "Both real and personal values are greater than zero and using the same pin number."
    ElseIf ErrorRec.ErrorType = 2 Then
      ThisError = "Both real or personal values equal zero."
    End If
    TrunName = QPTrim$(ErrorRec.CustName)
    '                            0                                1                                   2
    Print #ERptHandle, ErrorRec.CountyAcct; dlm; QPTrim$(ErrorRec.CountyAcctString); dlm; TrunName; dlm;
    '                      3                              4                      5                      6
    Print #ERptHandle, ErrorRec.ErrorType; dlm; ErrorRec.PersTot; dlm; ErrorRec.PersXTot; dlm; ErrorRec.RPinNum; dlm;
    '                         7                      8                      9
    Print #ERptHandle, ErrorRec.RealTot; dlm; ErrorRec.RealXTot; dlm; ErrorRec.PPinNum
  Next x
  
  Close
    
  arTCResultsRpt.Show
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Dim FileLen As Integer
  
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  FileLen = Len(Me.Caption)
  FileVers = Mid(Me.Caption, FileLen - 3, 4)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTCMainMenu1.")
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
    DoEvents
  End If

End Sub



Private Sub fpBtn1_Click()

End Sub
