VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUBEditMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UB V205 Edit Menu"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   1890
   ClientWidth     =   12210
   Icon            =   "frmUBEditMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      Caption         =   "MBR(Inv Cred)/(REV <> Tot)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6300
      TabIndex        =   30
      Top             =   7380
      Width           =   3255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Remove customers"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2760
      TabIndex        =   29
      Top             =   7860
      Width           =   2325
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Strip Trans"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2790
      TabIndex        =   28
      Top             =   7410
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   8340
      TabIndex        =   27
      Top             =   7860
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Revenue Switch"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6300
      TabIndex        =   26
      Top             =   6930
      Width           =   3255
   End
   Begin VB.CommandButton cmdPrintSumTrans 
      Caption         =   "Print Transaction Summary"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6300
      TabIndex        =   25
      Top             =   6510
      Width           =   3252
   End
   Begin VB.CommandButton cmdrestorereads 
      Caption         =   "Restore Last Reads From Last Bill"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6300
      TabIndex        =   20
      Top             =   5610
      Width           =   3252
   End
   Begin VB.CommandButton cmdblankowners 
      Caption         =   "Blank Owners"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6300
      TabIndex        =   19
      Top             =   5190
      Width           =   3252
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Deposit Balances"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6300
      TabIndex        =   17
      Top             =   4290
      Width           =   3252
   End
   Begin VB.CommandButton cmdFixCMTaxTrans 
      Caption         =   "Fix CM Tax Trans"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6300
      TabIndex        =   21
      Top             =   6060
      Width           =   3252
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change User Codes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6300
      TabIndex        =   16
      Top             =   3840
      Width           =   3252
   End
   Begin VB.CommandButton cmdBillcopy 
      Caption         =   "Set Bill Copies"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6300
      TabIndex        =   15
      Top             =   3396
      Width           =   3252
   End
   Begin VB.CommandButton cmdrecalcbal 
      Caption         =   "Recalc UB Cust Balances"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6300
      TabIndex        =   18
      Top             =   4740
      Width           =   3252
   End
   Begin VB.CommandButton cmdSetAllowLFCO 
      Caption         =   "Set Allow Late Fee/Cutoff"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6300
      TabIndex        =   14
      Top             =   2952
      Width           =   3252
   End
   Begin VB.CommandButton cmdClrCustBalances 
      Caption         =   "Clear Cust Balances"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6300
      TabIndex        =   13
      Top             =   2508
      Width           =   3252
   End
   Begin VB.CommandButton cmdClearMonthAmts 
      Caption         =   "Clear Month Charge Amts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6300
      TabIndex        =   12
      Top             =   2064
      Width           =   3252
   End
   Begin VB.CommandButton cmdChangeMulti 
      Caption         =   "Change &Multipliers"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2832
      TabIndex        =   11
      Top             =   7005
      Width           =   3204
   End
   Begin VB.CommandButton cmdPrintJournal 
      Caption         =   "Print Transactions(Del Cust)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2832
      TabIndex        =   10
      Top             =   6555
      Width           =   3204
   End
   Begin VB.CommandButton cmdAssignRates 
      Caption         =   "Assign Ra&te Codes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2832
      TabIndex        =   9
      Top             =   6120
      Width           =   3204
   End
   Begin VB.CommandButton cmdSetCycle 
      Caption         =   "Assign C&ycle Numbers"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2850
      TabIndex        =   8
      Top             =   5670
      Width           =   3204
   End
   Begin VB.CommandButton cmdUnDeleteCust 
      BackColor       =   &H008F8265&
      Caption         =   "Un&Delete Customer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2832
      MaskColor       =   &H8000000F&
      TabIndex        =   1
      Top             =   2508
      Width           =   3204
   End
   Begin VB.CommandButton cmdSequenceLoc 
      Caption         =   "Re-Sequence &Location #'s"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2832
      TabIndex        =   7
      Top             =   5220
      Width           =   3204
   End
   Begin VB.CommandButton cmdEditTransDates 
      BackColor       =   &H008F8265&
      Caption         =   "&Edit Transaction Dates"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2820
      MaskColor       =   &H8000000F&
      TabIndex        =   6
      Top             =   4740
      Width           =   3204
   End
   Begin VB.CommandButton cmdEditCMTRans 
      BackColor       =   &H008F8265&
      Caption         =   "&CM Edit Transaction"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2832
      MaskColor       =   &H8000000F&
      TabIndex        =   5
      Top             =   4284
      Width           =   3204
   End
   Begin VB.CommandButton cmdEditTrans 
      BackColor       =   &H008F8265&
      Caption         =   "&Edit Customer Transaction"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2832
      MaskColor       =   &H8000000F&
      TabIndex        =   4
      Top             =   3840
      Width           =   3204
   End
   Begin VB.CommandButton cmdExitMenu 
      Caption         =   "E&xit "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5298
      TabIndex        =   22
      Top             =   7920
      Width           =   1692
   End
   Begin VB.CommandButton cmdSetFlagUnread 
      Caption         =   "Set Reading Flags &Unread"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2832
      TabIndex        =   3
      Top             =   3396
      Width           =   3204
   End
   Begin VB.CommandButton cmdSetFlagRead 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Set Reading Flags &Read"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2832
      TabIndex        =   2
      Top             =   2952
      Width           =   3204
   End
   Begin VB.CommandButton cmdAdjBalances 
      BackColor       =   &H008F8265&
      Caption         =   "&Adjust Customer Balances"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2832
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   2064
      Width           =   3204
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   23
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "12:15 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "5/7/2014"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   1788
      X2              =   1788
      Y1              =   2064
      Y2              =   7944
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   1788
      X2              =   2508
      Y1              =   7944
      Y2              =   7944
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   1668
      X2              =   2628
      Y1              =   2064
      Y2              =   2064
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   1680
      X2              =   2640
      Y1              =   1944
      Y2              =   1944
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   1668
      X2              =   1668
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2628
      X2              =   2628
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   9804
      X2              =   9804
      Y1              =   2064
      Y2              =   7944
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   9804
      X2              =   10524
      Y1              =   7944
      Y2              =   7944
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   9684
      X2              =   10644
      Y1              =   2064
      Y2              =   2064
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   3
      X1              =   9684
      X2              =   10644
      Y1              =   1944
      Y2              =   1944
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   9684
      X2              =   9684
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   10644
      X2              =   10644
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UB Util Edit Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3540
      TabIndex        =   24
      Top             =   1104
      Width           =   5148
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   1788
      Top             =   2064
      Width           =   732
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   9804
      Top             =   2064
      Width           =   732
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1302
      Top             =   744
      Width           =   9612
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1302
      Top             =   624
      Width           =   9612
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   1668
      Top             =   1824
      Width           =   972
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   9684
      Top             =   1824
      Width           =   972
   End
End
Attribute VB_Name = "frmUBEditMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim ReadFlag As Boolean
Private Sub cmdAdjBalances_Click()
  Load frmBalAdjEntry
  frmBalAdjEntry.Show
  Unload Me
End Sub
Private Sub cmdAssignRates_Click()
  Load frmAssignRate
  frmAssignRate.Show
  Unload Me
End Sub

Private Sub cmdBillcopy_Click()
  Load frmSetBillCopies
  frmSetBillCopies.Show
  Unload Me
End Sub

Private Sub cmdblankowners_Click()
  Setownerblanks
End Sub

Private Sub cmdChangeMulti_Click()
  Load frmChangeMult
  frmChangeMult.Show
  Unload Me
End Sub

Private Sub cmdClearMonthAmts_Click()
  If MsgBox("This procedure sets all monthly charge fields to 0.  Are you sure you wish to continue?", vbOKCancel, "Clear Month Amts") = vbOK Then
    ClearMonthAmts
  End If
End Sub
Private Sub Setownerblanks()
  Dim UBOwnerRec As UBOwnerRecType
  Dim UBCustRecLen As Integer
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
'  FrmShowPctComp.Label1 = "Updating Owners"
'  FrmShowPctComp.Show
  Dim UBFile2 As Integer, UBFile As Integer, OwnerRecLen As Integer
  Dim NumOfRecs&, cnt&
  OwnerRecLen = Len(UBOwnerRec)
 
  UBFile = FreeFile
  Open UBPath$ + "UBCUST.dat" For Random Shared As UBFile Len = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  UBFile2 = FreeFile
  Open UBOwnerFile For Random Shared As UBFile2 Len = OwnerRecLen
 For cnt& = 1 To NumOfRecs&
 
  UBOwnerRec.OwnFName = String(20, " ") 'new owner info until user
  UBOwnerRec.OwnLName = String(15, " ")
  UBOwnerRec.ADDR1 = String(35, " ")
  UBOwnerRec.ADDR2 = String(35, " ")
  UBOwnerRec.CITY = String(18, " ")
  UBOwnerRec.STATE = String(2, " ")
  UBOwnerRec.ZIPCODE = String(10, " ")
  UBOwnerRec.HPHONE = String(14, " ")
  UBOwnerRec.WPHONE = String(14, " ")
  UBOwnerRec.ChkByte = Chr$(1)
  Put UBFile2, cnt&, UBOwnerRec


   
 Next
  Close
  Close UBFile2
  MsgBox "Procedure Complete", vbOKOnly, "Complete"

End Sub

Private Sub cmdClrCustBalances_Click()
  If MsgBox("This procedure sets all Customer balances to 0.  Are you sure you wish to continue?", vbOKCancel, "Clear Balances") = vbOK Then
    ClearBalances
  End If
End Sub

Private Sub cmdEditCMTRans_Click()
  Load frmCMSearch
  Unload Me
  DoEvents
  frmCMSearch.Show
End Sub

Private Sub cmdEditTrans_Click()
  frmCustEditLookUP.Caption = "Edit Customer Trans Find"
  frmCustEditLookUP.Label1.Caption = "Edit Customer Trans Find"
  frmCustEditLookUP.Wheretogo frmTRDispList2, frmTRDispList2
  Unload Me
  DoEvents
  frmCustEditLookUP.Show
  DoEvents
End Sub

Private Sub cmdEditTransDates_Click()
  Load frmUtilDateEdit
  Unload Me
  DoEvents
  frmUtilDateEdit.Show
End Sub


Private Sub cmdFixCMTaxTrans_Click()
FixTaxTrans
End Sub

Private Sub cmdPrintJournal_Click()
  Load frmRptTransJournal
  Unload Me
  frmRptTransJournal.Show
End Sub


Private Sub cmdPrintSumTrans_Click()
  Load frmRptTransSummary
  Unload Me
  frmRptTransSummary.Show
End Sub

Private Sub cmdrecalcbal_Click()
  If MsgBox("This procedure recalculates all Customer balances.  Depending on the amount of transactions and customers, this could take some time.  Are you sure you wish to continue?", vbOKCancel, "Recalc Bal") = vbOK Then
    RecalcUBCustBalances
  End If
End Sub

Private Sub cmdrestorereads_Click()
  If MsgBox("This procedure restores readings and dates from last billing into cust info. Are you sure you wish to continue?", vbOKCancel, "Restore reads and dates") = vbOK Then
    Fixlastreadsndate
  End If
End Sub

Private Sub cmdSequenceLoc_Click()
  Load frmReSequenceLoc
  Unload Me
  frmReSequenceLoc.Show
End Sub

Private Sub cmdSetAllowLFCO_Click()
  If MsgBox("This procedure sets all Customer- Allow Late Fee and Allow CutOff values to 'Y'.  Are you sure you wish to continue?", vbOKCancel, "Clear Month Amts") = vbOK Then
    SetAllowPenaltyY
  End If
End Sub

Private Sub cmdSetCycle_Click()
  Load frmSetCycle
  Unload Me
  frmSetCycle.Show
End Sub

Private Sub cmdSetFlagRead_Click()
  ReadFlag = True
  SetReadFlag (ReadFlag)
End Sub
Private Sub cmdSetFlagUnread_Click()
  ReadFlag = False
  SetReadFlag (ReadFlag)
End Sub
Private Sub SetReadFlag(ReadFlag)
  ReDim UBCustRec(1) As NewUBCustRecType
  Dim UBCustRecLen As Integer, UBFile As Integer, NumOfRecs As Long
  Dim cnt As Long, Mcnt As Integer
  UBCustRecLen = Len(UBCustRec(1))

  Select Case ReadFlag
  Case True
    If MsgBox("Set Meters as Read?", vbYesNo, "Continue?") = vbYes Then
      GoSub GOSetReadFlags
    End If
  Case Else
    If MsgBox("Set Meters as UNRead?", vbYesNo, "Continue?") = vbYes Then
      GoSub GOSetReadFlags
    End If
  End Select

Exit Sub

GOSetReadFlags:
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  FrmShowPctComp.Label1 = "Updating Read Flags...."
  FrmShowPctComp.Show

  UBFile = FreeFile
  Open UBPath$ + "UBCUST.dat" For Random Shared As UBFile Len = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  For cnt& = 1 To NumOfRecs&
    FrmShowPctComp.ShowPctComp cnt&, NumOfRecs&
    Get UBFile, cnt&, UBCustRec(1)
    For Mcnt = 1 To 7
      If ReadFlag Then
        UBCustRec(1).LocMeters(Mcnt).ReadFlag = "Y"
      Else
        UBCustRec(1).LocMeters(Mcnt).ReadFlag = ""
      End If
    Next
    Put UBFile, cnt&, UBCustRec(1)
  Next
  Close
  MsgBox "Procedure Complete", vbOKOnly, "Complete"
Return
End Sub


Private Sub cmdExitMenu_Click()
  DoEvents
  
  Unload Me
End Sub


'Private Sub cmdStrip_Click()
'  Load frmUtilStripTrans
'  frmUtilStripTrans.Show
'  Unload Me
'End Sub

Private Sub cmdUnDeleteCust_Click()
  Load frmCustUnDelete
  frmCustUnDelete.Show
  Unload Me
End Sub



'Private Sub Command1_Click()
''
''Fix3badcust 'FixBrokenMsgNum
'
' 'SetAvgusetoONE ''this sets all ub cust avguse and usecnt to 1
'End Sub

Private Sub Command2_Click()
  Load frmChangeUserCodes
  frmChangeUserCodes.Show
  Unload Me
End Sub

Private Sub Command3_Click()
  Load frmRptMastDep
  frmRptMastDep.Show
  Unload Me
End Sub

Private Sub Command1_Click()
  Load frmSwapRevs
  frmSwapRevs.Show , Me
  


End Sub

Private Sub Command4_Click()
  Load frmUtilStripTrans
  frmUtilStripTrans.Show
  Unload Me

  'FixBrokenMsgFile
'  If MsgBox("This will add fix invalid WO.  Do you wish to continue?", vbYesNo, "ATTENTION!!!") = vbYes Then
'    Call fixwos
'    'Fixnewcust4Harrisburg
'  End If
  
  
End Sub

Private Sub Command5_Click()

  Dim xxnum As Long
  If Len(QPTrim$(Me.Text1.Text)) > 0 Then
    xxnum = CInt(Me.Text1.Text)
    If MsgBox("This will eliminate any customer accounts greater than " + CStr(xxnum) + "  Are you sure you wish to continue?", vbYesNo) = vbYes Then
        FixBrokenCustFile (xxnum)
    End If
  End If


End Sub

Private Sub Command6_Click()
  Load frmRptMastBal
  Unload Me
  frmRptMastBal.Show
End Sub

'Private Sub Command4_Click()
'  DeleteBlankCusts
'End Sub
''Call fixaworkorder
''Private Sub Command1_Click()
''  FixCurrPrevMult
''  FixReadsonTrans
''End Sub

''
''Private Sub Command1_Click()
''FixBrokenAverage
''End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
     

  'screenW = (Screen.Width / Screen.TwipsPerPixelX)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitMenu.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed ubutil"
        'CitiTerminate
      End If
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
  DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyHome
      cmdAdjBalances.SetFocus
    Case vbKeyEnd
      cmdExitMenu.SetFocus
    Case Else:
  End Select
End Sub

Private Sub FixTaxTrans() 'recalcs revs on tax trans in cm and if trans tot diff then changes transtot to match rev tots
  Dim CMTrRecLen As Integer, TrHandle As Integer, TrNumRecs As Long, cnt As Long, TrType As String
  Dim TxRev As Double, TRev As Integer, FDate As String
  Dim TotalAmount As Double, CHANGE As Double
  Dim keepup As Long

  ReDim CMTrRec(1) As CMTransRecType            ' open transaction file
  CMTrRecLen = Len(CMTrRec(1))
  TrHandle = FreeFile
  Open UBPath$ + "CMTRANS.DAT" For Random Access Read Write Shared As TrHandle Len = CMTrRecLen
  TrNumRecs& = LOF(TrHandle) \ CMTrRecLen
  For cnt = 1 To TrNumRecs&
      Get TrHandle, cnt, CMTrRec(1)
        Select Case CMTrRec(1).TransSource
'        Case 30 To 39, 131, 231, 161, 261, 171, 271
'          TrType$ = "Tax Billing"
'          lblTaxInfo.Visible = True
'        Case Else
'          NoDoModTrans = True
'        End Select
       Case 131, 231
          TxRev# = 0
          For TRev = 1 To 7
            TxRev# = Round#(TxRev# + CMTrRec(1).TransRevAmt(TRev))
          Next
          TxRev# = Round#(TxRev# + CMTrRec(1).TransRevAmt(9))
        Case 161, 261
          TxRev# = 0
          For TRev = 1 To 8
            TxRev# = Round#(TxRev# + CMTrRec(1).TransRevAmt(TRev))
          Next
          TxRev# = Round#(TxRev# + CMTrRec(1).TransRevAmt(10))
        Case 171, 271
          TxRev# = 0
          For TRev = 1 To 10
            TxRev# = Round#(TxRev# + CMTrRec(1).TransRevAmt(TRev))
          Next
          TxRev# = Round#(TxRev# + CMTrRec(1).TransRevAmt(12))
        Case Else
         ' Next cnt
        End Select
        If TxRev# <> 0 Then
          If TxRev# <> CMTrRec(1).TransAmount Then
            CMTrRec(1).TransAmount = Round(TxRev#)
            keepup = keepup + 1
           Put TrHandle, cnt, CMTrRec(1)
          End If
          TxRev# = 0
        End If
  Next
  MsgBox "Edited " & keepup & " records"
  Close
  End Sub
Private Sub fixaworkorder()
  Dim WorkOrderRecLen As Integer
  Dim UBWrkOrd As Integer

  ReDim WorkOrderRec(1) As WorkOrderRecType
  WorkOrderRecLen = Len(WorkOrderRec(1))


    UBWrkOrd = FreeFile
    Open UBPath$ + "UBWRKORD.DAT" For Random Shared As UBWrkOrd Len = WorkOrderRecLen
    Get UBWrkOrd, 4636, WorkOrderRec(1)
    If WorkOrderRec(1).CustRec = 2870 Then
    WorkOrderRec(1).CompletedDate = WorkOrderRec(1).ENTRYDATE
    WorkOrderRec(1).CompleteByDate = WorkOrderRec(1).ENTRYDATE
    Put UBWrkOrd, 4636, WorkOrderRec(1)
    End If
    
    Get UBWrkOrd, 4538, WorkOrderRec(1)
    If WorkOrderRec(1).CustRec = 3813 Then
    WorkOrderRec(1).CompletedDate = WorkOrderRec(1).ENTRYDATE
    WorkOrderRec(1).CompleteByDate = WorkOrderRec(1).ENTRYDATE
    Put UBWrkOrd, 4538, WorkOrderRec(1)
    End If
   Close
    MsgBox "Workorder is completed"
   
End Sub
Private Sub GetDepTots()
  Dim UBCustRecLen As Integer, UBSetupreclen As Integer
  Dim UsingBook As Boolean, UsingName As Boolean
  Dim AcctNumber As Long, UBCust As Integer, UsingAcct As Boolean
  Dim IndexName As String, UBRpt As Integer, SEQNUMB As String
  Dim IdxRecLen As Integer, IdxFileSize As Long, OKToSkip As Boolean
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, Handle As Integer
  Dim cnt As Long, TDeposit As Double, ToPrint As String
  Dim Book As String, CustCnt As Long, ReportFile As String
  ReDim RevAmts(1 To 15) As Double
  Dim UBTransRecLen As Integer, NextTranRecs As Long
  Dim TransDate As Integer, Transamt As Double
  Dim RevCnt As Integer
  Dim UBTran As Integer, NumOfTranRecs As Long, PrevLastTrans As Long
  Dim TotalDepAmt As Double, LastTran As Long
  
 
  Dim UBCustRec As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec)
  Dim UBTransRec As UBTransRecType
  UBTransRecLen = Len(UBTransRec)
 
    UBTran = FreeFile
    Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTran Len = UBTransRecLen

  
  ToPrint$ = ""

    NumOfRecs = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  ReportFile$ = UBPath$ + "Deposit.txt"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
  
  For cnt = 1 To NumOfRecs
   
      AcctNumber = cnt

    Get UBCust, AcctNumber, UBCustRec
      If UBCustRec.DelFlag = 0 Then
        'If Round#(UBCustRec.DepositAmt) <> 0 Then
          LastTran& = UBCustRec.LastTrans
  
            TotalDepAmt# = 0
            ReDim RevAmts(1 To 15) As Double
            If LastTran& > 0 Then
              Do
                Get #UBTran, LastTran&, UBTransRec
                If UBTransRec.TransType = TranDepositPayment Then
                  For RevCnt = 1 To 15
                    If UBTransRec.RevAmt(RevCnt) > 0 Then
                      RevAmts(RevCnt) = Round#(RevAmts(RevCnt) + UBTransRec.RevAmt(RevCnt))
                      TotalDepAmt# = Round#(TotalDepAmt# + UBTransRec.RevAmt(RevCnt))
                    End If
                  Next
                ElseIf (UBTransRec.TransType = TranAppliedDeposit) Or (UBTransRec.TransType = TranRefundDeposit) Or (UBTransRec.TransType = TranDepPaymentVoid) Then
                  RevAmts(RevCnt) = Round#(RevAmts(RevCnt) - UBTransRec.RevAmt(RevCnt))
                      TotalDepAmt# = Round#(TotalDepAmt# - UBTransRec.RevAmt(RevCnt))
                End If
                LastTran& = UBTransRec.PrevTrans
              Loop While LastTran& > 0
            End If
        If TotalDepAmt# > 0 Or Round#(UBCustRec.DepositAmt) <> 0 Then
        '_________
          ToPrint$ = Str$(AcctNumber)
          ToPrint$ = ToPrint$ + "|" + Str$(UBCustRec.DepositAmt)
          For RevCnt = 1 To 15
            ToPrint$ = ToPrint$ + "|" + Str$(RevAmts(RevCnt))
          Next
          Print #UBRpt, ToPrint$
          ToPrint$ = ""
        End If
        'End If
    End If

  Next
  
  Close UBCust, UBRpt
  Close UBTran
MsgBox ("OK")
ExitDepositListing:
  Exit Sub

End Sub
Private Sub fixwos()
  Dim WorkOrderRecLen As Integer, whattrans As Long
  Dim UBWrkOrd As Integer

  ReDim WorkOrderRec(1) As WorkOrderRecType
  WorkOrderRecLen = Len(WorkOrderRec(1))

  UBWrkOrd = FreeFile
  Open UBPath$ + "UBWRKORD.DAT" For Random Shared As UBWrkOrd Len = WorkOrderRecLen
  whattrans = 2233
    Get UBWrkOrd, whattrans, WorkOrderRec(1)
    WorkOrderRec(1).CompletedDate = Date2Num("09/16/2010")


  Put UBWrkOrd, whattrans, WorkOrderRec(1)
  MsgBox "Edited "
End Sub

