VERSION 5.00
Begin VB.Form colorchg 
   Caption         =   "Color Change"
   ClientHeight    =   4020
   ClientLeft      =   2100
   ClientTop       =   1710
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   8880
   Begin VB.CommandButton Command5 
      Caption         =   "R E S E T   T O   S T A N D A R D   C O L O R S"
      Height          =   375
      Left            =   120
      TabIndex        =   106
      Top             =   3600
      Width           =   8655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Choose"
      Height          =   375
      Left            =   3960
      TabIndex        =   105
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Change"
      Height          =   375
      Left            =   7680
      TabIndex        =   102
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Try"
      Height          =   375
      Left            =   6360
      TabIndex        =   101
      Top             =   840
      Width           =   1095
   End
   Begin VB.Frame ncframe 
      Caption         =   "New Color"
      Height          =   2055
      Left            =   4800
      TabIndex        =   52
      Top             =   1440
      Visible         =   0   'False
      Width           =   3495
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   100
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   47
         Left            =   960
         TabIndex        =   99
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   98
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   97
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   96
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   95
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   2160
         TabIndex        =   94
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   6
         Left            =   2400
         TabIndex        =   93
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   7
         Left            =   720
         TabIndex        =   92
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   8
         Left            =   960
         TabIndex        =   91
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   9
         Left            =   1200
         TabIndex        =   90
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   10
         Left            =   1440
         TabIndex        =   89
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   11
         Left            =   1680
         TabIndex        =   88
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   12
         Left            =   1920
         TabIndex        =   87
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   13
         Left            =   2160
         TabIndex        =   86
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   14
         Left            =   2400
         TabIndex        =   85
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   15
         Left            =   720
         TabIndex        =   84
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   16
         Left            =   960
         TabIndex        =   83
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   17
         Left            =   1200
         TabIndex        =   82
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   18
         Left            =   1440
         TabIndex        =   81
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   19
         Left            =   1680
         TabIndex        =   80
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   20
         Left            =   1920
         TabIndex        =   79
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   21
         Left            =   2160
         TabIndex        =   78
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF00FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   22
         Left            =   2400
         TabIndex        =   77
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   23
         Left            =   720
         TabIndex        =   76
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   24
         Left            =   960
         TabIndex        =   75
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H000040C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   25
         Left            =   1200
         TabIndex        =   74
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   26
         Left            =   1440
         TabIndex        =   73
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   27
         Left            =   1680
         TabIndex        =   72
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   28
         Left            =   1920
         TabIndex        =   71
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   29
         Left            =   2160
         TabIndex        =   70
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C000C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   30
         Left            =   2400
         TabIndex        =   69
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   31
         Left            =   720
         TabIndex        =   68
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   32
         Left            =   960
         TabIndex        =   67
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00004080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   33
         Left            =   1200
         TabIndex        =   66
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   34
         Left            =   1440
         TabIndex        =   65
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   35
         Left            =   1680
         TabIndex        =   64
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   36
         Left            =   1920
         TabIndex        =   63
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   37
         Left            =   2160
         TabIndex        =   62
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   38
         Left            =   2400
         TabIndex        =   61
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   39
         Left            =   720
         TabIndex        =   60
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   40
         Left            =   960
         TabIndex        =   59
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   41
         Left            =   1200
         TabIndex        =   58
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00004040&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   42
         Left            =   1440
         TabIndex        =   57
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00004000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   43
         Left            =   1680
         TabIndex        =   56
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   44
         Left            =   1920
         TabIndex        =   55
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   45
         Left            =   2160
         TabIndex        =   54
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00400040&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   46
         Left            =   2400
         TabIndex        =   53
         Top             =   1560
         Width           =   255
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Choose"
      Height          =   375
      Left            =   240
      TabIndex        =   51
      Top             =   840
      Width           =   1095
   End
   Begin VB.Frame ocframe 
      Caption         =   "Old Color"
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   47
         Left            =   960
         TabIndex        =   4
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   7
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   8
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   2160
         TabIndex        =   9
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   6
         Left            =   2400
         TabIndex        =   10
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   7
         Left            =   720
         TabIndex        =   11
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   8
         Left            =   960
         TabIndex        =   12
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   9
         Left            =   1200
         TabIndex        =   13
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   10
         Left            =   1440
         TabIndex        =   14
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   11
         Left            =   1680
         TabIndex        =   15
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   12
         Left            =   1920
         TabIndex        =   16
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   13
         Left            =   2160
         TabIndex        =   17
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   14
         Left            =   2400
         TabIndex        =   18
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   15
         Left            =   720
         TabIndex        =   19
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   16
         Left            =   960
         TabIndex        =   20
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   17
         Left            =   1200
         TabIndex        =   21
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   18
         Left            =   1440
         TabIndex        =   22
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   19
         Left            =   1680
         TabIndex        =   23
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   20
         Left            =   1920
         TabIndex        =   24
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   21
         Left            =   2160
         TabIndex        =   25
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF00FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   22
         Left            =   2400
         TabIndex        =   26
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   23
         Left            =   720
         TabIndex        =   27
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   24
         Left            =   960
         TabIndex        =   28
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H000040C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   25
         Left            =   1200
         TabIndex        =   29
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   26
         Left            =   1440
         TabIndex        =   30
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   27
         Left            =   1680
         TabIndex        =   31
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   28
         Left            =   1920
         TabIndex        =   32
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   29
         Left            =   2160
         TabIndex        =   33
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C000C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   30
         Left            =   2400
         TabIndex        =   34
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   31
         Left            =   720
         TabIndex        =   35
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   32
         Left            =   960
         TabIndex        =   36
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00004080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   33
         Left            =   1200
         TabIndex        =   37
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   34
         Left            =   1440
         TabIndex        =   38
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   35
         Left            =   1680
         TabIndex        =   39
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   36
         Left            =   1920
         TabIndex        =   40
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   37
         Left            =   2160
         TabIndex        =   41
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   38
         Left            =   2400
         TabIndex        =   42
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   39
         Left            =   720
         TabIndex        =   43
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000040&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   40
         Left            =   960
         TabIndex        =   44
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   41
         Left            =   1200
         TabIndex        =   45
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00004040&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   42
         Left            =   1440
         TabIndex        =   46
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00004000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   43
         Left            =   1680
         TabIndex        =   47
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   44
         Left            =   1920
         TabIndex        =   48
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   45
         Left            =   2160
         TabIndex        =   49
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00400040&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   46
         Left            =   2400
         TabIndex        =   50
         Top             =   1560
         Width           =   255
      End
   End
   Begin VB.Label tc 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "none"
      Height          =   375
      Left            =   5160
      TabIndex        =   104
      Top             =   840
      Width           =   975
   End
   Begin VB.Label cc 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "none"
      Height          =   375
      Left            =   1440
      TabIndex        =   103
      Top             =   840
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8880
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "To Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label label20 
      BackStyle       =   0  'Transparent
      Caption         =   "From Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "colorchg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If ocframe.Visible = False Then
    Command1.Caption = "Close"
    ocframe.Visible = True
Else
    Command1.Caption = "Choose"
    ocframe.Visible = False
End If
End Sub

Private Sub Command2_Click()
For i = 0 To CIVIL.Controls.Count - 1
    a = CIVIL.Controls(i).Name
    On Error GoTo errrtn1
    If CIVIL.Controls(i).ForeColor = cc.ForeColor Then
        CIVIL.Controls(i).ForeColor = tc.ForeColor
        CIVIL.Controls(i).Refresh
    End If
n1:
Next i
Call Command4_Click
Exit Sub
errrtn1:
If Err = 438 Or Err = 458 Then
    Resume n1
End If
Resume Next

End Sub

Private Sub Command3_Click()
Dim a, b As Long, tofrom(1000, 2) As Long, ct As Integer
ct = 0
a = Dir("cc.tag")
If a > "" Then
    Open "CC.TAG" For Input As #1
    While Not EOF(1)
        ct = ct + 1
        Input #1, a, b
        tofrom(ct, 1) = a
        tofrom(ct, 2) = b
    Wend
    Close #1
End If
fm% = 0
For t% = 1 To ct
    If a = cc.ForeColor Then
        fm% = 1
        tofrom(t%, 2) = tc.ForeColor
        t% = ct
    End If
    If b = cc.ForeColor Then
        fm% = 1
        tofrom(t%, 2) = tc.ForeColor
        t% = ct
    End If
Next t%
If fm% = 0 Then
    ct = ct + 1
    tofrom(ct, 1) = cc.ForeColor
    tofrom(ct, 2) = tc.ForeColor
End If
Open "CC.TAG" For Output As #1
For t% = 1 To ct
    Print #1, tofrom(t%, 1), tofrom(t%, 2)
Next t%
Close #1
msg = MsgBox("Change Completed.", 48, "Genesis Information Log")
End Sub

Private Sub Command4_Click()
If ncframe.Visible = False Then
    Command4.Caption = "Close"
    ncframe.Visible = True
Else
    Command4.Caption = "Choose"
    ncframe.Visible = False
End If
End Sub

Private Sub Command5_Click()
On Error Resume Next
Kill "cc.tag"
On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload CIVIL
CIVIL.Show
End Sub

Private Sub Label1_Click(index As Integer)
tc.Caption = ""
tc.BackColor = Label1(index).BackColor
tc.ForeColor = Label1(index).BackColor
tc.Refresh
End Sub

Private Sub label2_Click(index As Integer)
cc.Caption = ""
cc.BackColor = Label2(index).BackColor
cc.ForeColor = Label2(index).BackColor
cc.Refresh
Call Command1_Click
End Sub
