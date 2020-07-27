VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmExpCustomerInfo 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export Customer Information"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   2175
   ClientWidth     =   12210
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExpCustInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CheckBox optgroup 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Select Group Range for Export:"
      Height          =   372
      Left            =   624
      TabIndex        =   6
      Top             =   2784
      Width           =   3612
   End
   Begin VB.CheckBox optCycle 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Select Cycle Range for Export:"
      Height          =   372
      Left            =   624
      TabIndex        =   3
      Top             =   2232
      Width           =   3612
   End
   Begin VB.CheckBox optbook 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Select Book Range for Export: "
      Height          =   372
      Left            =   624
      TabIndex        =   0
      Top             =   1704
      Width           =   3612
   End
   Begin VB.CommandButton cmdUntagAll 
      Caption         =   "&UnTag All"
      Height          =   324
      Left            =   6432
      TabIndex        =   16
      Top             =   3336
      Width           =   1476
   End
   Begin VB.CommandButton cmdTagAll 
      Caption         =   "Tag &All"
      Height          =   324
      Left            =   4776
      TabIndex        =   15
      Top             =   3336
      Width           =   1476
   End
   Begin VB.CheckBox chkDPCode 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DP Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   576
      TabIndex        =   27
      Top             =   6000
      Width           =   2508
   End
   Begin VB.CheckBox chkGroup 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Group Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   3288
      TabIndex        =   36
      Top             =   4176
      Width           =   2508
   End
   Begin VB.CheckBox chkSearchName 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Search Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   576
      TabIndex        =   19
      Top             =   4176
      Width           =   2508
   End
   Begin VB.CheckBox chkFileUnique 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Unique File Name"
      Height          =   324
      Left            =   9240
      TabIndex        =   10
      Top             =   2808
      Width           =   2412
   End
   Begin VB.CheckBox chkActive 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Active Customers Only"
      Height          =   324
      Left            =   9096
      TabIndex        =   9
      Top             =   1848
      Width           =   2652
   End
   Begin VB.CheckBox chkSeq 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Read Sequence"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   6000
      TabIndex        =   57
      Top             =   5088
      Width           =   2508
   End
   Begin VB.CheckBox chkQuotes 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Double Quotes  """
      Height          =   372
      Left            =   9672
      TabIndex        =   14
      Top             =   6072
      Width           =   2052
   End
   Begin VB.OptionButton OptDelimiter3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tab"
      Height          =   372
      Left            =   9672
      TabIndex        =   13
      Top             =   5232
      Width           =   2052
   End
   Begin VB.OptionButton OptDelimiter2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pipe Symbol  |"
      Height          =   372
      Left            =   9672
      TabIndex        =   12
      Top             =   4860
      Width           =   2052
   End
   Begin VB.OptionButton OptDelimiter1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Comma  ,"
      Height          =   372
      Left            =   9672
      TabIndex        =   11
      Top             =   4488
      Width           =   2052
   End
   Begin VB.CheckBox chkMembFees 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Membership Fees"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   6000
      TabIndex        =   64
      Top             =   6684
      Width           =   2508
   End
   Begin VB.CheckBox chkMonthly 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Monthly Payment Info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   6000
      TabIndex        =   63
      Top             =   6456
      Width           =   2508
   End
   Begin VB.CheckBox chkCustMsgs 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Customer Messages"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   6000
      TabIndex        =   67
      Top             =   7368
      Width           =   2508
   End
   Begin VB.CheckBox chkBillCycl 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Billing Cycle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   6000
      TabIndex        =   60
      Top             =   5772
      Width           =   2508
   End
   Begin VB.CheckBox chkDeposit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Deposit Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   3288
      TabIndex        =   50
      Top             =   7368
      Width           =   2508
   End
   Begin VB.CheckBox chkOwner 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Owner Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   3288
      TabIndex        =   45
      Top             =   6228
      Width           =   2508
   End
   Begin VB.CheckBox chkRevBal 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Revenue Balances"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   3288
      TabIndex        =   49
      Top             =   7140
      Width           =   2508
   End
   Begin VB.CheckBox chkPrevBal 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Previous Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   3288
      TabIndex        =   48
      Top             =   6912
      Width           =   2508
   End
   Begin VB.CheckBox chkCurrBal 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Current Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   3288
      TabIndex        =   47
      Top             =   6684
      Width           =   2508
   End
   Begin VB.CheckBox chkMeterInfo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Meter Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   6000
      TabIndex        =   66
      Top             =   7140
      Width           =   2508
   End
   Begin VB.CheckBox chkFlatInfo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Flat Rate Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   6000
      TabIndex        =   65
      Top             =   6912
      Width           =   2508
   End
   Begin VB.CheckBox chkRevInfo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Revenue Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   6000
      TabIndex        =   62
      Top             =   6228
      Width           =   2508
   End
   Begin VB.CheckBox chkHHMsg3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hand Held Message 3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   6000
      TabIndex        =   56
      Top             =   4860
      Width           =   2508
   End
   Begin VB.CheckBox chkHHMsg2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hand Held Message 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   6000
      TabIndex        =   55
      Top             =   4632
      Width           =   2508
   End
   Begin VB.CheckBox chkHHMsg1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hand Held Message 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   6000
      TabIndex        =   54
      Top             =   4404
      Width           =   2508
   End
   Begin VB.CheckBox chkProRatePCT 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ProRate Percent"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   6000
      TabIndex        =   53
      Top             =   4176
      Width           =   2508
   End
   Begin VB.CheckBox chkUsercode2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "User Code 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   6000
      TabIndex        =   52
      Top             =   3948
      Width           =   2508
   End
   Begin VB.CheckBox chkUsercode1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "User Code 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   6000
      TabIndex        =   51
      Top             =   3720
      Width           =   2508
   End
   Begin VB.CheckBox chkPumpcode 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pump Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   3288
      TabIndex        =   37
      Top             =   4404
      Width           =   2508
   End
   Begin VB.CheckBox chkPayCmnt 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Payment Comment"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   3288
      TabIndex        =   38
      Top             =   4632
      Width           =   2508
   End
   Begin VB.CheckBox chkBillCmnt 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Billing Comment"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   3288
      TabIndex        =   39
      Top             =   4860
      Width           =   2508
   End
   Begin VB.CheckBox chkSrCit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Senior Citizen Flag"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   3288
      TabIndex        =   44
      Top             =   6000
      Width           =   2508
   End
   Begin VB.CheckBox chkTaxExpt 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tax Exempt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   3288
      TabIndex        =   43
      Top             =   5772
      Width           =   2508
   End
   Begin VB.CheckBox chkCutOff 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Allow Cut Off Flag"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   3288
      TabIndex        =   42
      Top             =   5544
      Width           =   2508
   End
   Begin VB.CheckBox chkLateFee 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Allow Late Fee Flag"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   3288
      TabIndex        =   41
      Top             =   5316
      Width           =   2508
   End
   Begin VB.CheckBox chkCashOnly 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cash Only Flag"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   3288
      TabIndex        =   40
      Top             =   5088
      Width           =   2508
   End
   Begin VB.CheckBox chkZone 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Zone"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   6000
      TabIndex        =   59
      Top             =   5544
      Width           =   2508
   End
   Begin VB.CheckBox chkCustType 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Customer Type"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   576
      TabIndex        =   32
      Top             =   7140
      Width           =   2508
   End
   Begin VB.CheckBox chkDrLic 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Driver License #"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   576
      TabIndex        =   31
      Top             =   6912
      Width           =   2508
   End
   Begin VB.CheckBox chkSoSec 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Social Security #"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   576
      TabIndex        =   30
      Top             =   6684
      Width           =   2508
   End
   Begin VB.CheckBox chkWPhone 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Work Phone #"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   576
      TabIndex        =   29
      Top             =   6456
      Width           =   2508
   End
   Begin VB.CheckBox chkPostrte 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Postal Route"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   6000
      TabIndex        =   58
      Top             =   5316
      Width           =   2508
   End
   Begin VB.CheckBox chkBillTo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bill To"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   6000
      TabIndex        =   61
      Top             =   6000
      Width           =   2508
   End
   Begin VB.CheckBox chkAddr911 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Address 911"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   576
      TabIndex        =   33
      Top             =   7368
      Width           =   2508
   End
   Begin VB.CheckBox chkHPhone 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Home Phone #"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   576
      TabIndex        =   28
      Top             =   6228
      Width           =   2508
   End
   Begin VB.CheckBox chkZip 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Zip Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   576
      TabIndex        =   26
      Top             =   5772
      Width           =   2508
   End
   Begin VB.CheckBox chkState 
      BackColor       =   &H00C0C0C0&
      Caption         =   "State"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   576
      TabIndex        =   25
      Top             =   5544
      Width           =   2508
   End
   Begin VB.CheckBox chkCity 
      BackColor       =   &H00C0C0C0&
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   576
      TabIndex        =   24
      Top             =   5316
      Width           =   2508
   End
   Begin VB.CheckBox chkServAddr 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Service Address"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   576
      TabIndex        =   23
      Top             =   5088
      Width           =   2508
   End
   Begin VB.CheckBox chkAddr2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Address Line 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   576
      TabIndex        =   22
      Top             =   4860
      Width           =   2508
   End
   Begin VB.CheckBox chkAddr1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Address Line 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   576
      TabIndex        =   21
      Top             =   4632
      Width           =   2508
   End
   Begin VB.CheckBox chkCustName 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   576
      TabIndex        =   20
      Top             =   4404
      Width           =   2508
   End
   Begin VB.CheckBox chkStatus 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   3288
      TabIndex        =   34
      Top             =   3720
      Width           =   2508
   End
   Begin VB.CheckBox chkLocation 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Location Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   576
      TabIndex        =   18
      Top             =   3948
      Width           =   2508
   End
   Begin VB.CheckBox chkAcctNum 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Account Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   576
      TabIndex        =   17
      Top             =   3720
      Width           =   2508
   End
   Begin VB.CheckBox chkopendate 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Open Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   3288
      TabIndex        =   35
      Top             =   3948
      Width           =   2508
   End
   Begin VB.CheckBox ChkBank 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Include Bank Draft Info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   3288
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   46
      Top             =   6456
      Width           =   2508
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "F10 &Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   8928
      TabIndex        =   68
      Top             =   7248
      Width           =   1332
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Esc E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   10536
      TabIndex        =   69
      Top             =   7248
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   70
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
            TextSave        =   "2:00 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "6/13/2006"
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
   Begin EditLib.fpText fptxtRoute2 
      Height          =   348
      Left            =   6816
      TabIndex        =   2
      Top             =   1728
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      ThreeDOutsideStyle=   2
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "0123456789"
      MaxLength       =   2
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtRoute1 
      Height          =   348
      Left            =   5280
      TabIndex        =   1
      Top             =   1728
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      ThreeDOutsideStyle=   2
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "0123456789"
      MaxLength       =   2
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtCycle2 
      Height          =   348
      Left            =   6816
      TabIndex        =   5
      Top             =   2256
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      ThreeDOutsideStyle=   2
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "0123456789"
      MaxLength       =   2
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtCycle1 
      Height          =   348
      Left            =   5280
      TabIndex        =   4
      Top             =   2256
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      ThreeDOutsideStyle=   2
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "0123456789"
      MaxLength       =   2
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtGroup2 
      Height          =   348
      Left            =   6816
      TabIndex        =   8
      Top             =   2784
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      ThreeDOutsideStyle=   2
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "0123456789"
      MaxLength       =   2
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtGroup1 
      Height          =   348
      Left            =   5280
      TabIndex        =   7
      Top             =   2784
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      ThreeDOutsideStyle=   2
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "0123456789"
      MaxLength       =   2
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   324
      Index           =   5
      Left            =   6144
      TabIndex        =   82
      Top             =   2784
      Width           =   540
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   324
      Index           =   4
      Left            =   4368
      TabIndex        =   81
      Top             =   2784
      Width           =   828
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   324
      Index           =   3
      Left            =   6144
      TabIndex        =   80
      Top             =   2256
      Width           =   540
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   324
      Index           =   1
      Left            =   4368
      TabIndex        =   79
      Top             =   2256
      Width           =   828
   End
   Begin VB.Line Line4 
      X1              =   384
      X2              =   8616
      Y1              =   2688
      Y2              =   2688
   End
   Begin VB.Line Line3 
      X1              =   384
      X2              =   8616
      Y1              =   2136
      Y2              =   2136
   End
   Begin VB.Line Line2 
      X1              =   8688
      X2              =   11928
      Y1              =   2328
      Y2              =   2328
   End
   Begin VB.Line Line1 
      X1              =   8688
      X2              =   11928
      Y1              =   4032
      Y2              =   4032
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   """UBCustEx.ASC"" will be used if Option above is not selected."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   672
      Left            =   9240
      TabIndex        =   78
      Top             =   3144
      Width           =   2424
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      Height          =   5340
      Left            =   8664
      Top             =   1584
      Width           =   3276
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      Height          =   4452
      Left            =   288
      Top             =   3264
      Width           =   8388
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Select for Unique File Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   8760
      TabIndex        =   77
      Top             =   2448
      Width           =   3132
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Delimiter:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   8880
      TabIndex        =   76
      Top             =   4128
      Width           =   2268
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Text Qualifier:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   8904
      TabIndex        =   75
      Top             =   5736
      Width           =   2004
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Fields to Include in Export:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   360
      TabIndex        =   74
      Top             =   3360
      Width           =   3972
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Export Information"
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
      Left            =   3624
      TabIndex        =   73
      Top             =   600
      Width           =   5004
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   708
      Left            =   3216
      Top             =   432
      Width           =   5772
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   324
      Index           =   0
      Left            =   4368
      TabIndex        =   72
      Top             =   1728
      Width           =   828
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   324
      Index           =   2
      Left            =   6144
      TabIndex        =   71
      Top             =   1728
      Width           =   540
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      Height          =   1692
      Left            =   288
      Top             =   1584
      Width           =   8388
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   828
      Left            =   3216
      Top             =   312
      Width           =   5772
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuPrnScn 
         Caption         =   "Prin&t Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmExpCustomerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider

Private Sub cmdTagAll_Click()
'  chkAcctNum.Value = 1
'  chkLocation.Value = 1
'  chkSearchName.Value = 1
'  chkCustName.Value = 1
'  chkAddr1.Value = 1
'  chkAddr2.Value = 1
'  chkServAddr.Value = 1
'  chkCity.Value = 1
'  chkState.Value = 1
'  chkZip.Value = 1
'  chkDPCode.Value = 1
'  chkHPhone.Value = 1
'  chkWPhone.Value = 1
'  chkSoSec.Value = 1
'  chkDrLic.Value = 1
'  chkCustType.Value = 1
'  chkAddr911.Value = 1
'  chkStatus.Value = 1
'  chkopendate.Value = 1
'  chkGroup.Value = 1
'  chkPumpcode.Value = 1
'  chkPayCmnt.Value = 1
'  chkBillCmnt.Value = 1
'  chkCashOnly.Value = 1
'  chkLateFee.Value = 1
'  chkCutOff.Value = 1
'  chkTaxExpt.Value = 1
'  chkSrCit.Value = 1
'  chkOwner.Value = 1
'  ChkBank.Value = 1
'  chkCurrBal.Value = 1
'  chkPrevBal.Value = 1
'  chkRevBal.Value = 1
'  chkDeposit.Value = 1
'  chkUsercode1.Value = 1
'  chkUsercode2.Value = 1
'  chkProRatePCT.Value = 1
'  chkHHMsg1.Value = 1
'  chkHHMsg2.Value = 1
'  chkHHMsg3.Value = 1
'  chkSeq.Value = 1
'  chkPostrte.Value = 1
'  chkZone.Value = 1
'  chkBillCycl.Value = 1
'  chkBillTo.Value = 1
'  chkRevInfo.Value = 1
'  chkMonthly.Value = 1
'  chkMembFees.Value = 1
'  chkFlatInfo.Value = 1
'  chkMeterInfo.Value = 1
'  chkCustMsgs.Value = 1
'
End Sub

Private Sub cmdUntagAll_Click()
'  chkAcctNum.Value = 0
'  chkLocation.Value = 0
'  chkSearchName.Value = 0
'  chkCustName.Value = 0
'  chkAddr1.Value = 0
'  chkAddr2.Value = 0
'  chkServAddr.Value = 0
'  chkCity.Value = 0
'  chkState.Value = 0
'  chkZip.Value = 0
'  chkDPCode.Value = 0
'  chkHPhone.Value = 0
'  chkWPhone.Value = 0
'  chkSoSec.Value = 0
'  chkDrLic.Value = 0
'  chkCustType.Value = 0
'  chkAddr911.Value = 0
'  chkStatus.Value = 0
'  chkopendate.Value = 0
'  chkGroup.Value = 0
'  chkPumpcode.Value = 0
'  chkPayCmnt.Value = 0
'  chkBillCmnt.Value = 0
'  chkCashOnly.Value = 0
'  chkLateFee.Value = 0
'  chkCutOff.Value = 0
'  chkTaxExpt.Value = 0
'  chkSrCit.Value = 0
'  chkOwner.Value = 0
'  ChkBank.Value = 0
'  chkCurrBal.Value = 0
'  chkPrevBal.Value = 0
'  chkRevBal.Value = 0
'  chkDeposit.Value = 0
'  chkUsercode1.Value = 0
'  chkUsercode2.Value = 0
'  chkProRatePCT.Value = 0
'  chkHHMsg1.Value = 0
'  chkHHMsg2.Value = 0
'  chkHHMsg3.Value = 0
'  chkSeq.Value = 0
'  chkPostrte.Value = 0
'  chkZone.Value = 0
'  chkBillCycl.Value = 0
'  chkBillTo.Value = 0
'  chkRevInfo.Value = 0
'  chkMonthly.Value = 0
'  chkMembFees.Value = 0
'  chkFlatInfo.Value = 0
'  chkMeterInfo.Value = 0
'  chkCustMsgs.Value = 0
'
End Sub
Private Function chkforchks()
'  Dim chks As Integer
'  chks = 0
'  If chkAcctNum.Value = 1 Then chks = chks + 1
'  If chkLocation.Value = 1 Then chks = chks + 1
'  If chkSearchName.Value = 1 Then chks = chks + 1
'  If chkCustName.Value = 1 Then chks = chks + 1
'  If chkAddr1.Value = 1 Then chks = chks + 1
'  If chkAddr2.Value = 1 Then chks = chks + 1
'  If chkServAddr.Value = 1 Then chks = chks + 1
'  If chkCity.Value = 1 Then chks = chks + 1
'  If chkState.Value = 1 Then chks = chks + 1
'  If chkZip.Value = 1 Then chks = chks + 1
'  If chkDPCode.Value = 1 Then chks = chks + 1
'  If chkHPhone.Value = 1 Then chks = chks + 1
'  If chkWPhone.Value = 1 Then chks = chks + 1
'  If chkSoSec.Value = 1 Then chks = chks + 1
'  If chkDrLic.Value = 1 Then chks = chks + 1
'  If chkCustType.Value = 1 Then chks = chks + 1
'  If chkAddr911.Value = 1 Then chks = chks + 1
'  If chkStatus.Value = 1 Then chks = chks + 1
'  If chkopendate.Value = 1 Then chks = chks + 1
'  If chkGroup.Value = 1 Then chks = chks + 1
'  If chkPumpcode.Value = 1 Then chks = chks + 1
'  If chkPayCmnt.Value = 1 Then chks = chks + 1
'  If chkBillCmnt.Value = 1 Then chks = chks + 1
'  If chkCashOnly.Value = 1 Then chks = chks + 1
'  If chkLateFee.Value = 1 Then chks = chks + 1
'  If chkCutOff.Value = 1 Then chks = chks + 1
'  If chkTaxExpt.Value = 1 Then chks = chks + 1
'  If chkSrCit.Value = 1 Then chks = chks + 1
'  If chkOwner.Value = 1 Then chks = chks + 1
'  If ChkBank.Value = 1 Then chks = chks + 1
'  If chkCurrBal.Value = 1 Then chks = chks + 1
'  If chkPrevBal.Value = 1 Then chks = chks + 1
'  If chkRevBal.Value = 1 Then chks = chks + 1
'  If chkDeposit.Value = 1 Then chks = chks + 1
'  If chkUsercode1.Value = 1 Then chks = chks + 1
'  If chkUsercode2.Value = 1 Then chks = chks + 1
'  If chkProRatePCT.Value = 1 Then chks = chks + 1
'  If chkHHMsg1.Value = 1 Then chks = chks + 1
'  If chkHHMsg2.Value = 1 Then chks = chks + 1
'  If chkHHMsg3.Value = 1 Then chks = chks + 1
'  If chkSeq.Value = 1 Then chks = chks + 1
'  If chkPostrte.Value = 1 Then chks = chks + 1
'  If chkZone.Value = 1 Then chks = chks + 1
'  If chkBillCycl.Value = 1 Then chks = chks + 1
'  If chkBillTo.Value = 1 Then chks = chks + 1
'  If chkRevInfo.Value = 1 Then chks = chks + 1
'  If chkMonthly.Value = 1 Then chks = chks + 1
'  If chkMembFees.Value = 1 Then chks = chks + 1
'  If chkFlatInfo.Value = 1 Then chks = chks + 1
'  If chkMeterInfo.Value = 1 Then chks = chks + 1
'  If chkCustMsgs.Value = 1 Then chks = chks + 1
' chkforchks = chks
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
       ' UBLog "Closed via Expcustinfo by " + PWUser$
       ' CitiTerminate
      End If
    End If
  End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
'    Case vbKeyUp:
'      SendKeys "+{Tab}"
'      KeyCode = 0
    Case vbKeyEscape:
      cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      cmdOk_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

'
'Private Sub fptxtCycle1_ChangeMode(EditMode As Integer)
'  EditMode = True
'End Sub
'
'Private Sub fptxtCycle1_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeyReturn Then
'    fptxtCycle2.SetFocus
'  End If
'End Sub
'Private Sub fptxtCycle2_ChangeMode(EditMode As Integer)
'  EditMode = True
'End Sub
'
'Private Sub fptxtCycle2_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeyReturn Then
'    optgroup.SetFocus
'  End If
'End Sub
'Private Sub fptxtGroup1_ChangeMode(EditMode As Integer)
'  EditMode = True
'End Sub
'
'Private Sub fptxtGroup1_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeyReturn Then
'    fptxtGroup2.SetFocus
'  End If
'End Sub
'Private Sub fptxtGroup2_ChangeMode(EditMode As Integer)
'  EditMode = True
'End Sub
'
'Private Sub fptxtGroup2_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeyReturn Then
'    cmdTagAll.SetFocus
'  End If
'End Sub
'
'Private Sub fptxtRoute1_ChangeMode(EditMode As Integer)
'  EditMode = True
'End Sub
'Private Sub fptxtRoute1_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeyReturn Then
'    fptxtRoute2.SetFocus
'  End If
'End Sub
'Private Sub fptxtRoute2_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeyReturn Then
'    optCycle.SetFocus
'  End If
'End Sub
'Private Sub fptxtRoute2_ChangeMode(EditMode As Integer)
'  EditMode = True
'End Sub
'Private Function ValidRoutes()
'  If optbook.Value = 1 Then
'    If fptxtRoute1 <> "" And fptxtRoute2 <> "" Then
'      If fptxtRoute1 > fptxtRoute2 Then
'        MsgBox "Invalid Selection, The Beginning Book Should Be Less or Equal to Ending Book.", vbOKOnly, "Invalid Selection"
'        ValidRoutes = False
'      Else
'        ValidRoutes = True
'      End If
'    Else
'      MsgBox "Book May Not Be Left Blank.", vbOKOnly, "Invalid Selection"
'    End If
'  ElseIf optCycle.Value = 1 Then
'    If fptxtCycle1 <> "" And fptxtCycle2 <> "" Then
'      If fptxtCycle1 > fptxtCycle2 Then
'        MsgBox "Invalid Selection, The Beginning Cycle Should Be Less or Equal to Ending Cycle.", vbOKOnly, "Invalid Selection"
'        ValidRoutes = False
'      Else
'        ValidRoutes = True
'      End If
'    Else
'      MsgBox "Cycle May Not Be Left Blank.", vbOKOnly, "Invalid Selection"
'    End If
'  ElseIf optgroup.Value = 1 Then
'    If fptxtGroup1 <> "" And fptxtGroup2 <> "" Then
'      If fptxtGroup1 > fptxtGroup2 Then
'        MsgBox "Invalid Selection, The Beginning Group Should Be Less or Equal to Ending Group.", vbOKOnly, "Invalid Selection"
'        ValidRoutes = False
'      Else
'        ValidRoutes = True
'      End If
'    Else
'      MsgBox "Group May Not Be Left Blank.", vbOKOnly, "Invalid Selection"
'    End If
'  Else
'    MsgBox "You Must Select A Range for Export.", vbOKOnly, "Invalid Selection"
'  End If
'End Function

Private Sub cmdExit_Click()
  
  Unload Me
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
'  StatusBar1.Panels.Item(1).Text = TOWNNAME$
 ' Me.HelpContextID = hlpExportCustomer
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
Private Sub cmdOk_Click()
'
'  If ValidRoutes Then
'    If chkforchks > 0 Then
'      DeActivateControls Me, True
'do consumpstuff here
'      ExpCustStuff
'      ActivateControls Me, True
'    Else
'      MsgBox "You Must Select at least One Field to Export.", vbOKOnly, "Invalid Selection"
'    End If
'  End If
End Sub

'Private Sub ExpCustStuff()
'  Dim Dash80 As String, IndexName As String, IdxRecLen As Integer
'  Dim UBCustRecLen As Integer, UBCust As Integer, UBTran As Integer
'  Dim IdxFileSize As Long, IdxNumOfRecs As Long, cnt As Long
'  Dim UBTranRecLen As Integer, NumOfRecs As Long, NumOfCust As Long
'  Dim Handle As Integer, UsingBook As Boolean, NumOfPeriods As Integer
'  Dim RecNo As Long, DidCnt As Long, ThisTrans As Long, FMonth As Integer
'  Dim FYear As Integer, TYear As Integer, TMonth As Integer
'  Dim FMCnt As Integer, DidAMeter As Boolean, MtrCnt As Integer
'  Dim MeterType As String, MeterConsp As Long, MaxMeterAmt As Long
'  Dim TotalConsump As Long, QPos As Integer, LocationNumber As String
'  Dim Zip As String, CCCnt As Long, NumofRevs As Integer, BuckFmt As String
'  Dim Bookone As Integer, Bookto As Integer, qc As String, ThisBook As String
'  Dim q As String, C As String, qcq As String, OKFlag As Boolean
'  Dim UBRpt As String, zz As String, zzN As Integer, CCnt As Long
'  Dim UBOwnerRecLen As Integer, UBFile As Integer, AcctNumber As Long
'  Dim WhatBook As Integer, Export As Long, RCnt As Integer, FCnt As Integer
'  Dim MCnt As Integer, tempTot As Double, MessageRec As Integer
'  Dim UBMessRecLen As Integer, ThisFile As String, Today As String
'  Dim UBMFile As Integer, Ext As String, GCode As String
'  Dim Txt As String, ChkName As String
'  Dim GroupCde As GroupCodeRecType
'  Dim GrpCodeRecLen As Integer, ghandle As Integer
'  FrmShowPctComp.Label1 = "Creating Export Files"
'  FrmShowPctComp.cmdCancel.Enabled = False
'  FrmShowPctComp.Show , Me
'  q$ = ""
'  qc$ = ""
'  qcq$ = ""
'  If OptDelimiter1.Value = True And chkQuotes.Value <> 1 Then
'    qcq$ = ","
'  ElseIf OptDelimiter2.Value = True And chkQuotes.Value <> 1 Then
'    qcq$ = "|"   'special for johnston co
'  ElseIf OptDelimiter3.Value = True And chkQuotes.Value <> 1 Then
'    qcq$ = Chr$(9)  'this is tab
'  ElseIf chkQuotes.Value = 1 Then
'    q$ = Chr$(34)
'    If OptDelimiter1.Value = True Then
'      qc$ = q$ + ","
'      qcq$ = q$ + "," + q$
'    ElseIf OptDelimiter2.Value = True Then
'      qc$ = q$ + "|"
'      qcq$ = q$ + "|" + q$
'    ElseIf OptDelimiter3.Value = True Then
'      qc$ = q$ + Chr$(9)
'      qcq$ = q$ + Chr$(9) + q$ 'this is tab
'    End If
'  End If
'  If optbook.Value = 1 Then
'    Bookone = Val(QPTrim(fptxtRoute1))
'    Bookto = Val(QPTrim(fptxtRoute2))
'  ElseIf optCycle.Value = 1 Then
'    Bookone = Val(QPTrim(fptxtCycle1))
'    Bookto = Val(QPTrim(fptxtCycle2))
'  ElseIf optgroup.Value = 1 Then
'    Bookone = Val(QPTrim(fptxtGroup1))
'    Bookto = Val(QPTrim(fptxtGroup2))
'  End If
'  IndexName$ = BookIndexFile
'  UsingBook = True
'  OKFlag = True
'  BuckFmt$ = "########.##"
'  NumofRevs = GetNumOfRevs%
'  ReDim UBCustRec(1) As NewUBCustRecType
'  UBCustRecLen = Len(UBCustRec(1))
'
'  ReDim UBOwnerRec(1) As UBOwnerRecType
'  UBOwnerRecLen = Len(UBOwnerRec(1))
'
'  ReDim UBMessRec(1) As UBMessRecType
'  UBMessRecLen = Len(UBMessRec(1))
'
'
'  IdxRecLen = 4               'we are using a long integer
'  IdxFileSize& = FileSize(IndexName$)
'  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
'  ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
'  'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
'  NumOfRecs = IdxNumOfRecs
'
'  Handle = FreeFile
'  Open IndexName$ For Random Shared As Handle Len = IdxRecLen
'  For cnt = 1 To IdxNumOfRecs
'    Get #Handle, cnt, IdxBuff(cnt)
'  Next
'  Close Handle
'
'  UBFile = FreeFile
'  Open UBPath$ + "UBOWNER.DAT" For Random Shared As UBFile Len = UBOwnerRecLen
'
'  UBCust = FreeFile
'  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
'  If chkFileUnique.Value = 1 Then
'    Ext$ = ".ASC"
'    ThisFile$ = "UBC"
'    For cnt = 1 To 5
'      GetRPTName ThisFile$
'      ChkName$ = ThisFile$ + Ext$
'      If Exist(ChkName$) = False Then
'        ThisFile$ = ChkName$
'        Exit For
'      End If
'    Next
'  Else
'    ThisFile$ = "UBCustEx.ASC"
'    KillFile (UBPath$ + ThisFile$)
'  End If
'  If optgroup.Value = 1 Then
'    GrpCodeRecLen = Len(GroupCde)
'    ghandle = FreeFile
'    Open UBPath$ + "UBGrpCde.DAT" For Random Shared As ghandle Len = GrpCodeRecLen
'  End If
'
'  UBRpt = FreeFile
'  Open UBPath$ + ThisFile$ For Output As UBRpt
'  GoSub DoHeaders
'  For cnt = 1 To NumOfRecs
'    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
'    If FrmShowPctComp.Out Then
'      Close
'      Unload FrmShowPctComp
'      GoTo ExitMastCustListing
'    End If
'
'    AcctNumber = IdxBuff(cnt).RecNum
'    Get UBCust, AcctNumber, UBCustRec(1)
'    Get UBFile, AcctNumber, UBOwnerRec(1)
'    '*************************************
'    '   Main body of Printing goes here
'
'' 'This is just to temp export for Johnston Co.
''    If UBCustRec(1).DelFlag <> -1 Then
''        GoSub ExportThisAccount
''    End If
'    If chkActive.Value = 1 Then
'      If UBCustRec(1).DelFlag <> -1 And UBCustRec(1).Status = "A" Then
'        If optbook.Value = 1 Then
'          ThisBook$ = UBCustRec(1).BOOK
'        ElseIf optCycle.Value = 1 Then
'          ThisBook$ = UBCustRec(1).BILLCYCL
'        ElseIf optgroup.Value = 1 Then
'          If UBCustRec(1).GroupCodeRec > 0 Then
'            Get #ghandle, UBCustRec(1).GroupCodeRec, GroupCde
'            ThisBook$ = QPTrim$(GroupCde.GroupCode)
'          End If
'        End If
'        If Left$(ThisBook$, 1) = "0" Then
'          WhatBook = Val(Right$(ThisBook$, 1))
'        Else
'          WhatBook = Val(ThisBook$)
'        End If
'        If WhatBook <= Bookto And WhatBook >= Bookone Then
'          GoSub ExportThisAccount
'        End If
'      End If
'    Else
'      If UBCustRec(1).DelFlag <> -1 Then
'        If optbook.Value = 1 Then
'          ThisBook$ = UBCustRec(1).BOOK
'        ElseIf optCycle.Value = 1 Then
'          ThisBook$ = UBCustRec(1).BILLCYCL
'        ElseIf optgroup.Value = 1 Then
'          If UBCustRec(1).GroupCodeRec > 0 Then
'            Get #ghandle, UBCustRec(1).GroupCodeRec, GroupCde
'            ThisBook$ = QPTrim$(GroupCde.GroupCode)
'          End If
'        End If
'        If Left$(ThisBook$, 1) = "0" Then
'          WhatBook = Val(Right$(ThisBook$, 1))
'        Else
'          WhatBook = Val(ThisBook$)
'        End If
'        If WhatBook <= Bookto And WhatBook >= Bookone Then
'          GoSub ExportThisAccount
'        End If
'      End If
'    End If
'  Next
'
'  Close
'  If Export& > 0 Then
'    MsgBox "File " & UBPath$ & ThisFile$ & " Exported with " & Export& & " Accounts.", vbOKOnly, "Export Completed."
'  Else
'    MsgBox "No Information Found to Export.", vbOKOnly, "Procedure Ended"
'  End If
'GoTo ExitMastCustListing
'
'ExportThisAccount:
'  MessageRec = UBCustRec(1).MessageRec
'  Export& = Export& + 1
'  LocationNumber$ = QPTrim$(UBCustRec(1).BOOK + "-" + UBCustRec(1).SEQNUMB)
'  Zip$ = QPTrim$(UBCustRec(1).ZIPCODE)
'  If Len(Zip$) > 5 Then
'    Zip$ = Left$(Zip$, 5) + "-" + Mid$(Zip$, 6)
'  End If
'  If chkGroup.Value = 1 Then
'    If UBCustRec(1).GroupCodeRec > 0 Then
'     Get #ghandle, UBCustRec(1).GroupCodeRec, GroupCde
'     GCode$ = QPTrim$(GroupCde.GroupCode)
'    Else
'     GCode$ = "None"
'    End If
'  End If
'
'  If chkAcctNum.Value = 1 Then Print #UBRpt, q$; QPTrim$(Str$(AcctNumber));
'  If chkLocation.Value = 1 Then Print #UBRpt, qcq$; LocationNumber$;
'  If chkSearchName.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).SEARCH);
'  If chkCustName.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).CustName);
'  If chkAddr1.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).Addr1);
'  If chkAddr2.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).Addr2);
'  If chkServAddr.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).ServAddr);
'  If chkCity.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).City);
'  If chkState.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).State);
'  If chkZip.Value = 1 Then Print #UBRpt, qcq$; Zip$;
'  If chkDPCode.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).DPCode);
'  If chkHPhone.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).HPHONE);
'  If chkWPhone.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).WPHONE);
'  If chkSoSec.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).SOSEC);
'  If chkDrLic.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).DRVLIC);
'  If chkCustType.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).CUSTTYPE);
'  If chkAddr911.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).Addr911);
' '
'  If chkStatus.Value = 1 Then Print #UBRpt, qcq$; UBCustRec(1).Status;
'  If chkopendate.Value = 1 Then Print #UBRpt, qcq$; Num2Date$(UBCustRec(1).OPENDATE);
'  If chkGroup.Value = 1 Then Print #UBRpt, qcq$; GCode$;
'  If chkPumpcode.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).PumpCode);
'  If chkPayCmnt.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).PAYCMNT);
'  If chkBillCmnt.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).BILLCMNT);
'  If chkCashOnly.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).CASHONLY);
'  If chkLateFee.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).LATEFEE);
'  If chkCutOff.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).CUTOFFYN);
'  If chkTaxExpt.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).TAXEXPT);
'  If chkSrCit.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).SRCIT);
'  If chkOwner.Value = 1 Then
'    Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).OwnLName);
'    Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).OwnFName);
'    Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).Addr1);
'    Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).Addr2);
'    Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).City);
'    Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).State);
'    Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).ZIPCODE);
'    Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).HPHONE);
'    Print #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).WPHONE);
'  End If
'  If ChkBank.Value = 1 Then
'    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).USEDRAFT);
'    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).AcctType);
'    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).BankName);
'    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).BANKLOC);
'    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).TRANSIT);
'    Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).BankAcct);
'  End If
'  If chkCurrBal.Value = 1 Then Print #UBRpt, qcq$; Using(BuckFmt$, UBCustRec(1).CurrBalance);
'  If chkPrevBal.Value = 1 Then Print #UBRpt, qcq$; Using(BuckFmt$, UBCustRec(1).PrevBalance);
'  If chkRevBal.Value = 1 Then
'    For RCnt = 1 To NumofRevs
'      Print #UBRpt, qcq$; Using(BuckFmt$, UBCustRec(1).CurrRevAmts(RCnt));
'    Next
'  End If
'  If chkDeposit.Value = 1 Then Print #UBRpt, qcq$; Using(BuckFmt$, UBCustRec(1).DepositAmt);
''
'  If chkUsercode1.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).USERCODE1);
'  If chkUsercode2.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).USERCODE2);
'  If chkProRatePCT.Value = 1 Then
'    If UBCustRec(1).ProRatePCT > 0 Then
'      Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).ProRatePCT));
'    Else
'      Print #UBRpt, qcq$; "100";
'    End If
'  End If
'  If chkHHMsg1.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).HHMSG1);
'  If chkHHMsg2.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).HHMSG2);
'  If chkHHMsg3.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).HHMSG3);
'  If chkSeq.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).Seq));
'  If chkPostrte.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).POSTRTE);
'  If chkZone.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).ZONE);
'  If chkBillCycl.Value = 1 Then Print #UBRpt, qcq$; UBCustRec(1).BILLCYCL;
'  If chkBillTo.Value = 1 Then Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).BillTo);
'  If chkRevInfo.Value = 1 Then
'    For RCnt = 1 To NumofRevs
'      Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).serv(RCnt).Ratecode);
'      Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).serv(RCnt).RMtrType);
'    Next
'  End If
'  If chkMonthly.Value = 1 Then
'    For MCnt = 1 To 2
'      Print #UBRpt, qcq$; QPTrim$(Str(UBCustRec(1).Monthly(MCnt).AMTOWED));
'      Print #UBRpt, qcq$; QPTrim$(Str(UBCustRec(1).Monthly(MCnt).TotAmtPD));
'      Print #UBRpt, qcq$; QPTrim$(Str(UBCustRec(1).Monthly(MCnt).PayAmt));
'      Print #UBRpt, qcq$; QPTrim$(Str(UBCustRec(1).Monthly(MCnt).RevSource));
'    Next
'  End If
'  If chkMembFees.Value = 1 Then
'    Print #UBRpt, qcq$; QPTrim$(Str(UBCustRec(1).MFEE1));
'    Print #UBRpt, qcq$; QPTrim$(Str(UBCustRec(1).MFEE2));
'  End If
''flatrates
'  If chkFlatInfo.Value = 1 Then
'    For FCnt = 1 To 4
'      Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).FlatRates(FCnt).FRDESC);
'      If UBCustRec(1).FlatRates(FCnt).FRAMT > 0 Then
'        Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).FlatRates(FCnt).FRAMT));
'      Else
'        Print #UBRpt, qcq$; "0.00";
'      End If
'      Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).FlatRates(FCnt).FRFREQ);
'
'      If UBCustRec(1).FlatRates(FCnt).REVSRC > 0 Then
'        Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).FlatRates(FCnt).REVSRC));
'      Else
'        Print #UBRpt, qcq$; "0";
'      End If
'      If UBCustRec(1).FlatRates(FCnt).NumMin > 0 Then
'        Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).FlatRates(FCnt).NumMin));
'      Else
'        Print #UBRpt, qcq$; "0";
'      End If
'    Next
'  End If
''meters
'  If chkMeterInfo.Value = 1 Then
'    For MCnt = 1 To 7
'      Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).LocMeters(MCnt).MtrNum);
'      If UBCustRec(1).LocMeters(MCnt).MTRMulti > 0 Then
'        Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).LocMeters(MCnt).MTRMulti));
'      Else
'        Print #UBRpt, qcq$; "0";
'      End If
'      Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).LocMeters(MCnt).MTRType);
'      Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).LocMeters(MCnt).MtrUnit);
'
'      If UBCustRec(1).LocMeters(MCnt).NumUser > 0 Then
'        Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).LocMeters(MCnt).NumUser));
'      Else
'        Print #UBRpt, qcq$; "0";
'      End If
'      If UBCustRec(1).LocMeters(MCnt).InsDate > 0 Then
'        Print #UBRpt, qcq$; Num2Date$(UBCustRec(1).LocMeters(MCnt).InsDate);
'      Else
'        Print #UBRpt, qcq$; "??/??/????";
'      End If
'      If UBCustRec(1).LocMeters(MCnt).CurRead > 0 Then
'        Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).LocMeters(MCnt).CurRead));
'      Else
'        Print #UBRpt, qcq$; "0";
'      End If
'
'      If UBCustRec(1).LocMeters(MCnt).PrevRead > 0 Then
'        Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).LocMeters(MCnt).PrevRead));
'      Else
'        Print #UBRpt, qcq$; "0";
'      End If
'
'      If UBCustRec(1).LocMeters(MCnt).CurDate > 0 Then
'        Print #UBRpt, qcq$; Num2Date$(UBCustRec(1).LocMeters(MCnt).CurDate);
'      Else
'        Print #UBRpt, qcq$; "??/??/????";
'      End If
'      If UBCustRec(1).LocMeters(MCnt).PastDate > 0 Then
'        Print #UBRpt, qcq$; Num2Date$(UBCustRec(1).LocMeters(MCnt).PastDate);
'      Else
'        Print #UBRpt, qcq$; "??/??/????";
'      End If
'        Print #UBRpt, qcq$; QPTrim$(UBCustRec(1).LocMeters(MCnt).MtrIDNO);
'        Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).LocMeters(MCnt).MtrLat));
'        Print #UBRpt, qcq$; QPTrim$(Str$(UBCustRec(1).LocMeters(MCnt).MtrLng));
'    '
'
'    '    ReadFlag  AS STRING * 1    'hidden & protected
'    '    AvgUse    AS LONG          'hidden & protected
'    '    UseCnt    AS INTEGER       'hidden & protected
'    Next
'  End If
'  If chkCustMsgs.Value = 1 Then
'    If MessageRec > 0 Then
'      UBMFile = FreeFile
'      Open UBPath + "UBMESAGE.DAT" For Random Shared As UBMFile Len = UBMessRecLen
'      Get UBMFile, MessageRec, UBMessRec(1)
'      Close UBMFile
'      For zzN = 1 To 15
'        Txt = UBMessRec(1).MessLine(zzN).msg
'        Txt = RTrim(RemNulls(Txt))
'        Print #UBRpt, qcq$; Txt;
'      Next
'    Else
'      For zzN = 1 To 15
'        Print #UBRpt, qcq$; "";
'      Next
'    End If
'  End If
'  Print #UBRpt, q$
'
'Return
'
'DoHeaders:
'  If chkAcctNum.Value = 1 Then Print #UBRpt, q$; "Account";
'  If chkLocation.Value = 1 Then Print #UBRpt, qcq$; "Location";
'  If chkSearchName.Value = 1 Then Print #UBRpt, qcq$; "SearchName";
'  If chkCustName.Value = 1 Then Print #UBRpt, qcq$; "CustName";
'  If chkAddr1.Value = 1 Then Print #UBRpt, qcq$; "ADDR1";
'  If chkAddr2.Value = 1 Then Print #UBRpt, qcq$; "ADDR2";
'  If chkServAddr.Value = 1 Then Print #UBRpt, qcq$; "ServAddr";
'  If chkCity.Value = 1 Then Print #UBRpt, qcq$; "CITY";
'  If chkState.Value = 1 Then Print #UBRpt, qcq$; "STATE";
'  If chkZip.Value = 1 Then Print #UBRpt, qcq$; "Zip";
'  If chkDPCode.Value = 1 Then Print #UBRpt, qcq$; "DP Code";
'  If chkHPhone.Value = 1 Then Print #UBRpt, qcq$; "HPHONE";
'  If chkWPhone.Value = 1 Then Print #UBRpt, qcq$; "WPHONE";
'  If chkSoSec.Value = 1 Then Print #UBRpt, qcq$; "SOSEC";
'  If chkDrLic.Value = 1 Then Print #UBRpt, qcq$; "DRVLIC";
'  If chkCustType.Value = 1 Then Print #UBRpt, qcq$; "CUSTTYPE";
'  If chkAddr911.Value = 1 Then Print #UBRpt, qcq$; "Addr911";
'  '
'  If chkStatus.Value = 1 Then Print #UBRpt, qcq$; "Status";
'  If chkopendate.Value = 1 Then Print #UBRpt, qcq$; "OpenDate";
'  If chkGroup.Value = 1 Then Print #UBRpt, qcq$; "Group";
'  If chkPumpcode.Value = 1 Then Print #UBRpt, qcq$; "PumpCode";
'  If chkPayCmnt.Value = 1 Then Print #UBRpt, qcq$; "PAYCMNT";
'  If chkBillCmnt.Value = 1 Then Print #UBRpt, qcq$; "BILLCMNT";
'  If chkCashOnly.Value = 1 Then Print #UBRpt, qcq$; "CASHONLY";
'  If chkLateFee.Value = 1 Then Print #UBRpt, qcq$; "LATEFEE";
'  If chkCutOff.Value = 1 Then Print #UBRpt, qcq$; "CUTOFFYN";
'  If chkTaxExpt.Value = 1 Then Print #UBRpt, qcq$; "TAXEXPT";
'  If chkSrCit.Value = 1 Then Print #UBRpt, qcq$; "SRCIT";
'  If chkOwner.Value = 1 Then
'    Print #UBRpt, qcq$; "OwnLName";
'    Print #UBRpt, qcq$; "OwnFName";
'    Print #UBRpt, qcq$; "ADDR1";
'    Print #UBRpt, qcq$; "ADDR2";
'    Print #UBRpt, qcq$; "CITY";
'    Print #UBRpt, qcq$; "STATE";
'    Print #UBRpt, qcq$; "ZIPCODE";
'    Print #UBRpt, qcq$; "HPHONE";
'    Print #UBRpt, qcq$; "WPHONE";
'  End If
'  If ChkBank.Value = 1 Then
'    Print #UBRpt, qcq$; "USEDRAFT";
'    Print #UBRpt, qcq$; "AcctType";
'    Print #UBRpt, qcq$; "BankName";
'    Print #UBRpt, qcq$; "BANKLOC";
'    Print #UBRpt, qcq$; "TRANSIT";
'    Print #UBRpt, qcq$; "BankAcct";
'  End If
'  If chkCurrBal.Value = 1 Then Print #UBRpt, qcq$; "CurrBalance";
'  If chkPrevBal.Value = 1 Then Print #UBRpt, qcq$; "PrevBalance";
'  If chkRevBal.Value = 1 Then
'    For RCnt = 1 To NumofRevs
'      Print #UBRpt, qcq$; "CurrRevAmts";
'    Next
'  End If
'  If chkDeposit.Value = 1 Then Print #UBRpt, qcq$; "Deposit";
'  '
'  If chkUsercode1.Value = 1 Then Print #UBRpt, qcq$; "USERCODE1";
'  If chkUsercode2.Value = 1 Then Print #UBRpt, qcq$; "USERCODE2";
'  If chkProRatePCT.Value = 1 Then Print #UBRpt, qcq$; "ProRate";
'  If chkHHMsg1.Value = 1 Then Print #UBRpt, qcq$; "HHMSG1";
'  If chkHHMsg2.Value = 1 Then Print #UBRpt, qcq$; "HHMSG2";
'  If chkHHMsg3.Value = 1 Then Print #UBRpt, qcq$; "HHMSG3";
'  If chkSeq.Value = 1 Then Print #UBRpt, qcq$; "Read Seq";
'  If chkPostrte.Value = 1 Then Print #UBRpt, qcq$; "POSTRTE";
'  If chkZone.Value = 1 Then Print #UBRpt, qcq$; "ZONE";
'  If chkBillCycl.Value = 1 Then Print #UBRpt, qcq$; "Bill Cycle";
'  If chkBillTo.Value = 1 Then Print #UBRpt, qcq$; "BillTo";
'  If chkRevInfo.Value = 1 Then
'    For RCnt = 1 To NumofRevs
'      Print #UBRpt, qcq$; Str(RCnt) + "RATECODE";
'      Print #UBRpt, qcq$; "RMtrType";
'    Next
'  End If
'  If chkMonthly.Value = 1 Then
'    For MCnt = 1 To 2
'      Print #UBRpt, qcq$; "M-AmtOwed";
'      Print #UBRpt, qcq$; "M-TAmtPaid";
'      Print #UBRpt, qcq$; "M-PayAmt";
'      Print #UBRpt, qcq$; "M-RevSrc";
'    Next
'  End If
'  If chkMembFees.Value = 1 Then
'    Print #UBRpt, qcq$; "MembFEE1";
'    Print #UBRpt, qcq$; "MembFEE2";
'  End If
''flatrates
'  If chkFlatInfo.Value = 1 Then
'    For FCnt = 1 To 4
'      Print #UBRpt, qcq$; Str(FCnt) + "FRDESC";
'      Print #UBRpt, qcq$; "FRAMT";
'      Print #UBRpt, qcq$; "FRFREQ";
'      Print #UBRpt, qcq$; "REVSRC";
'      Print #UBRpt, qcq$; "NumMin";
'    Next
'  End If
''meters
'  If chkMeterInfo.Value = 1 Then
'    For MCnt = 1 To 7
'      Print #UBRpt, qcq$; Str(MCnt) + "MtrNum";
'      Print #UBRpt, qcq$; "MTRMulti";
'      Print #UBRpt, qcq$; "MtrType";
'      Print #UBRpt, qcq$; "MTRUnit";
'      Print #UBRpt, qcq$; "NumUser";
'      Print #UBRpt, qcq$; "InsDate";
'      Print #UBRpt, qcq$; "CurRead";
'      Print #UBRpt, qcq$; "PrevRead";
'      Print #UBRpt, qcq$; "CurDate";
'      Print #UBRpt, qcq$; "PastDate";
'      Print #UBRpt, qcq$; "MeterID";
'      Print #UBRpt, qcq$; "Latitude";
'      Print #UBRpt, qcq$; "Longitude";
'    Next
'  End If
'  If chkCustMsgs.Value = 1 Then
'    For zzN = 1 To 15
'      Print #UBRpt, qcq$; "Message Line " + Str(zzN);
'    Next
'  End If
'  Print #UBRpt, q$
'Return
'ExitMastCustListing:
'
'End Sub
'
'Private Sub optbook_Click()
'  If optbook.Value = 1 Then
'    optCycle.Value = 0
'    optgroup.Value = 0
'  End If
'End Sub
'
'Private Sub optCycle_Click()
'  If optCycle.Value = 1 Then
'    optbook.Value = 0
'    optgroup.Value = 0
'  End If
'End Sub
'
'Private Sub optgroup_Click()
'  If optgroup.Value = 1 Then
'    optbook.Value = 0
'    optCycle.Value = 0
'  End If
'End Sub
Public Sub ExpDecalCust()

  Dim DCVehLen As Integer
  Dim CFile As Integer, VFile As Integer
  Dim NumOfCust As Long, NumOfVeh As Long
  Dim NumOfRecs As Long, DCCustLen As Integer
  Dim cnt As Long, numnum As Long, CustRec As Long
  Dim DCRpt As String, qcq$
  DCRpt = FreeFile
  Open DCPath$ + "DCCustex.ASC" For Output As DCRpt

  ReDim DCCustRec(1) As DCCustRecType
  ReDim DCVehRec(1 To 2) As DCVehType
  DCCustLen = Len(DCCustRec(1))
  DCVehLen = Len(DCVehRec(1))
  qcq$ = "|"
  If Exist(DCPath$ + "DCCust.dat") Then
    CFile = FreeFile
    Open "DCCust.dat" For Random Shared As CFile Len = DCCustLen
    NumOfCust& = LOF(CFile) / DCCustLen
   For cnt = 1 To NumOfCust
    Get CFile, cnt, DCCustRec(1)
   Print #DCRpt, Str(cnt);
   Print #DCRpt, qcq$; QPTrim$(DCCustRec(1).CUSTNUMB);
   Print #DCRpt, qcq$; QPTrim$(DCCustRec(1).SORTNAME);
   Print #DCRpt, qcq$; QPTrim$(DCCustRec(1).BILLNAME);
   Print #DCRpt, qcq$; QPTrim$(DCCustRec(1).ADDRESS1);
   Print #DCRpt, qcq$; QPTrim$(DCCustRec(1).ADDRESS2);
   Print #DCRpt, qcq$; QPTrim$(DCCustRec(1).City);
   Print #DCRpt, qcq$; QPTrim$(DCCustRec(1).State);
   Print #DCRpt, qcq$; QPTrim$(DCCustRec(1).ZIPCODE);
   Print #DCRpt, qcq$; QPTrim$(DCCustRec(1).SOSEC);
   Print #DCRpt, qcq$; QPTrim$(DCCustRec(1).DRVLIC);
   Print #DCRpt, qcq$; Num2Date$(DCCustRec(1).DATEOPED);
   Print #DCRpt, qcq$; QPTrim$(DCCustRec(1).CASHONLY);
   Print #DCRpt, qcq$; QPTrim$(DCCustRec(1).resident);
   Print #DCRpt, qcq$; QPTrim$(DCCustRec(1).Owner);
   Print #DCRpt, qcq$; QPTrim$(DCCustRec(1).HPHONE);
   Print #DCRpt, qcq$; QPTrim$(DCCustRec(1).WPHONE);
   Print #DCRpt, qcq$; QPTrim$(DCCustRec(1).LICENSE);
   Print #DCRpt, qcq$; Str(DCCustRec(1).Valid);
   Print #DCRpt, qcq$; Str(DCCustRec(1).AcctBal);
   Print #DCRpt, qcq$; QPTrim$(DCCustRec(1).Deleted);
   Print #DCRpt, qcq$; Str(DCCustRec(1).FirstTrans);
   Print #DCRpt, qcq$; Str(DCCustRec(1).LastTrans);
   Print #DCRpt, qcq$; Str(DCCustRec(1).FirstCar);
   Print #DCRpt, qcq$; Str(DCCustRec(1).LastCar);
   Print #DCRpt, qcq$; QPTrim$(DCCustRec(1).SocSec1);
   Print #DCRpt, qcq$; QPTrim$(DCCustRec(1).OtherName);
   Print #DCRpt,
  Next
  End If
End Sub
Public Sub ExpDecalVeh()
  Dim DCVehLen As Integer
  Dim VFile As Integer
  Dim NumOfVeh As Long
  Dim NumOfRecs As Long
  Dim cnt As Long
  Dim DCRpt As String, qcq$
  DCRpt = FreeFile
  Open DCPath$ + "DCVehex.ASC" For Output As DCRpt
  ReDim DCVRec(1) As DCVehType
  DCVehLen = Len(DCVRec(1))
  qcq$ = "|"
  If Exist(DCPath$ + "DCVeh.dat") Then
    VFile = FreeFile
    Open "DCVeh.dat" For Random Shared As VFile Len = DCVehLen
    NumOfVeh& = LOF(VFile) / DCVehLen
    For cnt = 1 To NumOfVeh
      Get VFile, cnt, DCVRec(1)
      Print #DCRpt, Str(cnt);
      Print #DCRpt, qcq$; QPTrim$(DCVRec(1).DecalCat);
      Print #DCRpt, qcq$; Str(DCVRec(1).Fee);
      Print #DCRpt, qcq$; QPTrim$(DCVRec(1).makemodel);
      Print #DCRpt, qcq$; QPTrim$(DCVRec(1).StateTag);
      Print #DCRpt, qcq$; Num2Date$(DCVRec(1).ExpireDate);
      Print #DCRpt, qcq$; QPTrim$(DCVRec(1).Sticker);
      Print #DCRpt, qcq$; QPTrim$(DCVRec(1).Valid);
      Print #DCRpt, qcq$; QPTrim$(DCVRec(1).Active);
      Print #DCRpt, qcq$; QPTrim$(DCVRec(1).Desc);
      Print #DCRpt, qcq$; QPTrim$(DCVRec(1).Notes);
      Print #DCRpt, qcq$; QPTrim$(DCVRec(1).PBFlag);
      Print #DCRpt, qcq$; Str(DCVRec(1).MasterRecord);
      Print #DCRpt, qcq$; Str(DCVRec(1).NextRec);
      Print #DCRpt,
    Next
  End If
End Sub

