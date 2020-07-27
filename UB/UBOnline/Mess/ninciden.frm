VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form incident 
   BackColor       =   &H00808000&
   Caption         =   "Genesis Incident Report version 1.0                        "
   ClientHeight    =   7905
   ClientLeft      =   30
   ClientTop       =   465
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7905
   ScaleWidth      =   11895
   WindowState     =   1  'Minimized
   Begin VB.Frame lookupframe 
      BackColor       =   &H00808000&
      Caption         =   "UCR Lookup-To search for multiple words in a description, type in the words separated by a comma."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   11880
      TabIndex        =   216
      Top             =   6840
      Visible         =   0   'False
      Width           =   7260
      Begin VB.ListBox lookuplist 
         Height          =   840
         Left            =   120
         TabIndex        =   266
         Top             =   1200
         Width           =   6975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&FIND"
         Height          =   255
         Left            =   120
         TabIndex        =   264
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&CLOSE"
         Height          =   255
         Left            =   5520
         TabIndex        =   265
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox lookup 
         Height          =   285
         Left            =   120
         TabIndex        =   263
         Top             =   480
         Width           =   6975
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   7185
      Left            =   0
      ScaleHeight     =   7125
      ScaleWidth      =   11565
      TabIndex        =   582
      Top             =   600
      Width           =   11625
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         Height          =   18795
         Left            =   360
         Picture         =   "ninciden.frx":0000
         ScaleHeight     =   18765
         ScaleWidth      =   11565
         TabIndex        =   583
         Top             =   -1680
         Width           =   11595
         Begin VB.Frame vucrf 
            BackColor       =   &H00808000&
            Caption         =   "Victim Connected To These UCR's"
            ForeColor       =   &H00FFFFFF&
            Height          =   2055
            Left            =   360
            TabIndex        =   280
            Top             =   3360
            Visible         =   0   'False
            Width           =   3375
            Begin VB.CommandButton closevucrf 
               Caption         =   "Close"
               Height          =   315
               Left            =   120
               TabIndex        =   282
               Top             =   1695
               Width           =   3135
            End
            Begin MSComctlLib.ListView vucrlist 
               Height          =   1335
               Left            =   120
               TabIndex        =   281
               Top             =   240
               Width           =   3135
               _ExtentX        =   5530
               _ExtentY        =   2355
               View            =   3
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               HideColumnHeaders=   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Object.Width           =   5115
               EndProperty
            End
         End
         Begin VB.Frame flagframe 
            BackColor       =   &H00808000&
            Caption         =   "Incident Flag Frame"
            ForeColor       =   &H00FFFFFF&
            Height          =   4815
            Left            =   3960
            TabIndex        =   588
            Top             =   5880
            Visible         =   0   'False
            Width           =   2775
            Begin VB.CommandButton Command25 
               Caption         =   "Restore List"
               Height          =   495
               Left            =   1560
               TabIndex        =   593
               Top             =   4200
               Width           =   1095
            End
            Begin VB.CommandButton Command24 
               Caption         =   "Save List"
               Height          =   495
               Left            =   120
               TabIndex        =   592
               Top             =   4200
               Width           =   1095
            End
            Begin VB.CommandButton Command23 
               Caption         =   "Close"
               Height          =   495
               Left            =   1560
               TabIndex        =   591
               Top             =   3600
               Width           =   1095
            End
            Begin VB.CommandButton flagbutton 
               Caption         =   "Flag"
               Height          =   495
               Left            =   120
               TabIndex        =   590
               Top             =   3600
               Width           =   1095
            End
            Begin MSComctlLib.ListView incidentlist 
               Height          =   3255
               Left            =   120
               TabIndex        =   589
               Top             =   240
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   5741
               View            =   3
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "IncidentNumber"
                  Object.Width           =   3528
               EndProperty
            End
         End
         Begin VB.Frame sdrugframe 
            BackColor       =   &H00808000&
            BorderStyle     =   0  'None
            Height          =   2460
            Index           =   7
            Left            =   1
            TabIndex        =   385
            Top             =   8000
            Visible         =   0   'False
            Width           =   6855
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   21
               Left            =   120
               TabIndex        =   392
               Top             =   240
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   22
               Left            =   120
               TabIndex        =   389
               Top             =   960
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   23
               Left            =   120
               TabIndex        =   386
               Top             =   1680
               Width           =   2895
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   21
               Left            =   4440
               TabIndex        =   394
               Top             =   240
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   21
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   393
               Top             =   240
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   22
               Left            =   4440
               TabIndex        =   391
               Top             =   960
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   22
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   390
               Top             =   960
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   23
               Left            =   4440
               TabIndex        =   388
               Top             =   1680
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   23
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   387
               Top             =   1680
               Width           =   1215
            End
            Begin VB.CommandButton closedrugframes 
               Caption         =   "Close"
               Height          =   255
               Index           =   7
               Left            =   3120
               TabIndex        =   395
               Top             =   2145
               Width           =   1215
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Type Of Drug                                              Amount                 Measurement"
               ForeColor       =   &H0000FFFF&
               Height          =   255
               Index           =   7
               Left            =   105
               TabIndex        =   396
               Top             =   45
               Width           =   6615
            End
         End
         Begin VB.Frame sdrugframe 
            BackColor       =   &H00808000&
            BorderStyle     =   0  'None
            Height          =   2460
            Index           =   5
            Left            =   10080
            TabIndex        =   409
            Top             =   5000
            Visible         =   0   'False
            Width           =   6855
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   15
               Left            =   120
               TabIndex        =   416
               Top             =   240
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   16
               Left            =   120
               TabIndex        =   413
               Top             =   960
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   17
               Left            =   120
               TabIndex        =   410
               Top             =   1680
               Width           =   2895
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   15
               Left            =   4440
               TabIndex        =   284
               Top             =   240
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   15
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   417
               Top             =   240
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   16
               Left            =   4440
               TabIndex        =   415
               Top             =   960
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   16
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   414
               Top             =   960
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   17
               Left            =   4440
               TabIndex        =   412
               Top             =   1680
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   17
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   411
               Top             =   1680
               Width           =   1215
            End
            Begin VB.CommandButton closedrugframes 
               Caption         =   "Close"
               Height          =   255
               Index           =   5
               Left            =   3120
               TabIndex        =   285
               Top             =   2160
               Width           =   1215
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Type Of Drug                                              Amount                 Measurement"
               ForeColor       =   &H0000FFFF&
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   286
               Top             =   0
               Width           =   6615
            End
         End
         Begin VB.Frame sdrugframe 
            BackColor       =   &H00808000&
            BorderStyle     =   0  'None
            Height          =   2460
            Index           =   4
            Left            =   10080
            TabIndex        =   287
            Top             =   7000
            Visible         =   0   'False
            Width           =   6855
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   12
               Left            =   120
               TabIndex        =   288
               Top             =   240
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   13
               Left            =   120
               TabIndex        =   289
               Top             =   960
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   14
               Left            =   120
               TabIndex        =   290
               Top             =   1680
               Width           =   2895
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   12
               Left            =   4440
               TabIndex        =   291
               Top             =   240
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   12
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   292
               Top             =   240
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   13
               Left            =   4440
               TabIndex        =   293
               Top             =   960
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   13
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   294
               Top             =   960
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   14
               Left            =   4440
               TabIndex        =   295
               Top             =   1680
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   14
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   296
               Top             =   1680
               Width           =   1215
            End
            Begin VB.CommandButton closedrugframes 
               Caption         =   "Close"
               Height          =   255
               Index           =   4
               Left            =   3120
               TabIndex        =   297
               Top             =   2160
               Width           =   1215
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Type Of Drug                                              Amount                 Measurement"
               ForeColor       =   &H0000FFFF&
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   298
               Top             =   -15
               Width           =   6615
            End
         End
         Begin VB.Frame sdrugframe 
            BackColor       =   &H00808000&
            BorderStyle     =   0  'None
            Height          =   2460
            Index           =   3
            Left            =   11040
            TabIndex        =   299
            Top             =   7000
            Visible         =   0   'False
            Width           =   6855
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   9
               Left            =   120
               TabIndex        =   300
               Top             =   240
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   10
               Left            =   120
               TabIndex        =   301
               Top             =   960
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   11
               Left            =   120
               TabIndex        =   302
               Top             =   1680
               Width           =   2895
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   9
               Left            =   4440
               TabIndex        =   306
               Top             =   240
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   9
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   303
               Top             =   240
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   10
               Left            =   4440
               TabIndex        =   307
               Top             =   960
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   10
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   304
               Top             =   960
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   11
               Left            =   4440
               TabIndex        =   308
               Top             =   1680
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   11
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   305
               Top             =   1680
               Width           =   1215
            End
            Begin VB.CommandButton closedrugframes 
               Caption         =   "Close"
               Height          =   255
               Index           =   3
               Left            =   3120
               TabIndex        =   309
               Top             =   2160
               Width           =   1215
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Type Of Drug                                              Amount                 Measurement"
               ForeColor       =   &H0000FFFF&
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   310
               Top             =   0
               Width           =   6615
            End
         End
         Begin VB.Frame sdrugframe 
            BackColor       =   &H00808000&
            BorderStyle     =   0  'None
            Height          =   2520
            Index           =   2
            Left            =   1
            TabIndex        =   373
            Top             =   6000
            Visible         =   0   'False
            Width           =   6855
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   8
               Left            =   120
               TabIndex        =   380
               Top             =   1680
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   7
               Left            =   120
               TabIndex        =   377
               Top             =   960
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   6
               Left            =   120
               TabIndex        =   374
               Top             =   240
               Width           =   2895
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   8
               Left            =   4440
               TabIndex        =   382
               Top             =   1680
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   8
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   381
               Top             =   1680
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   7
               Left            =   4440
               TabIndex        =   379
               Top             =   960
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   7
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   378
               Top             =   960
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   6
               Left            =   4440
               TabIndex        =   376
               Top             =   240
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   6
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   375
               Top             =   240
               Width           =   1215
            End
            Begin VB.CommandButton closedrugframes 
               Caption         =   "Close"
               Height          =   255
               Index           =   2
               Left            =   3120
               TabIndex        =   383
               Top             =   2160
               Width           =   1215
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Type Of Drug                                              Amount                 Measurement"
               ForeColor       =   &H0000FFFF&
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   384
               Top             =   15
               Width           =   6615
            End
         End
         Begin VB.Frame sdrugframe 
            BackColor       =   &H00808000&
            BorderStyle     =   0  'None
            Caption         =   "s"
            Height          =   2460
            Index           =   1
            Left            =   6960
            TabIndex        =   320
            Top             =   5040
            Visible         =   0   'False
            Width           =   6855
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   5
               Left            =   120
               TabIndex        =   327
               Top             =   1680
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   4
               Left            =   120
               TabIndex        =   324
               Top             =   960
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   3
               Left            =   120
               TabIndex        =   321
               Top             =   240
               Width           =   2895
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   5
               Left            =   4440
               TabIndex        =   329
               Top             =   1680
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   5
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   328
               Top             =   1680
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   4
               Left            =   4440
               TabIndex        =   326
               Top             =   960
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   4
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   325
               Top             =   960
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   3
               Left            =   4440
               TabIndex        =   323
               Top             =   240
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   3
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   322
               Top             =   240
               Width           =   1215
            End
            Begin VB.CommandButton closedrugframes 
               Caption         =   "Close"
               Height          =   255
               Index           =   1
               Left            =   3120
               TabIndex        =   330
               Top             =   2160
               Width           =   1215
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Type Of Drug                                              Amount                 Measurement"
               ForeColor       =   &H0000FFFF&
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   331
               Top             =   0
               Width           =   6615
            End
         End
         Begin VB.Frame sdrugframe 
            BackColor       =   &H00808000&
            BorderStyle     =   0  'None
            Height          =   2460
            Index           =   0
            Left            =   1
            TabIndex        =   345
            Top             =   9720
            Visible         =   0   'False
            Width           =   6855
            Begin VB.CommandButton closedrugframes 
               Caption         =   "Close"
               Height          =   255
               Index           =   0
               Left            =   3120
               TabIndex        =   355
               Top             =   2160
               Width           =   1215
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   0
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   347
               Top             =   240
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   0
               Left            =   4440
               TabIndex        =   348
               Top             =   240
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   1
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   350
               Top             =   960
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   1
               Left            =   4440
               TabIndex        =   351
               Top             =   960
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   2
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   353
               Top             =   1680
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   2
               Left            =   4440
               TabIndex        =   354
               Top             =   1680
               Width           =   2295
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   0
               Left            =   120
               TabIndex        =   346
               Top             =   240
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   1
               Left            =   120
               TabIndex        =   349
               Top             =   960
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   2
               Left            =   120
               TabIndex        =   352
               Top             =   1680
               Width           =   2895
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Type Of Drug                                              Amount                 Measurement"
               ForeColor       =   &H0000FFFF&
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   356
               Top             =   0
               Width           =   6615
            End
         End
         Begin VB.Frame sdrugframe 
            BackColor       =   &H00808000&
            BorderStyle     =   0  'None
            Height          =   2460
            Index           =   6
            Left            =   1920
            TabIndex        =   397
            Top             =   7920
            Visible         =   0   'False
            Width           =   6855
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   18
               Left            =   120
               TabIndex        =   404
               Top             =   240
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   19
               Left            =   120
               TabIndex        =   401
               Top             =   960
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   20
               Left            =   120
               TabIndex        =   398
               Top             =   1680
               Width           =   2895
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   18
               Left            =   4440
               TabIndex        =   406
               Top             =   240
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   18
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   405
               Top             =   240
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   19
               Left            =   4440
               TabIndex        =   403
               Top             =   960
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   19
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   402
               Top             =   960
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   20
               Left            =   4440
               TabIndex        =   400
               Top             =   1680
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   20
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   399
               Top             =   1680
               Width           =   1215
            End
            Begin VB.CommandButton closedrugframes 
               Caption         =   "Close"
               Height          =   255
               Index           =   6
               Left            =   3120
               TabIndex        =   407
               Top             =   2160
               Width           =   1215
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Type Of Drug                                              Amount                 Measurement"
               ForeColor       =   &H0000FFFF&
               Height          =   255
               Index           =   6
               Left            =   120
               TabIndex        =   408
               Top             =   0
               Width           =   6615
            End
         End
         Begin VB.Frame pinfoframe 
            BackColor       =   &H00808000&
            Caption         =   "Property Info"
            Height          =   3900
            Index           =   5
            Left            =   11160
            TabIndex        =   573
            Top             =   8880
            Visible         =   0   'False
            Width           =   4325
            Begin VB.CommandButton Command22 
               Caption         =   "Clear"
               Height          =   270
               Index           =   5
               Left            =   3420
               TabIndex        =   579
               Top             =   3525
               Width           =   645
            End
            Begin VB.CommandButton Command10 
               Caption         =   "C   L   O   S   E"
               Height          =   300
               Index           =   5
               Left            =   60
               TabIndex        =   578
               Top             =   3495
               Width           =   3060
            End
            Begin VB.ListBox minorlist 
               Height          =   645
               Index           =   5
               Left            =   915
               TabIndex        =   577
               Top             =   2700
               Width           =   3105
            End
            Begin VB.ListBox majorlist 
               Height          =   645
               Index           =   5
               Left            =   915
               TabIndex        =   576
               Top             =   1890
               Width           =   3105
            End
            Begin VB.ListBox group 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   645
               Index           =   5
               Left            =   915
               TabIndex        =   575
               Top             =   1080
               Width           =   3105
            End
            Begin VB.ListBox pucrlist 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   645
               Index           =   5
               Left            =   915
               Sorted          =   -1  'True
               TabIndex        =   574
               Top             =   240
               Width           =   3105
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Minor Grouping"
               Height          =   450
               Index           =   5
               Left            =   120
               TabIndex        =   269
               Top             =   2730
               Width           =   810
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Major Grouping"
               Height          =   450
               Index           =   5
               Left            =   105
               TabIndex        =   270
               Top             =   1920
               Width           =   810
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Group"
               Height          =   270
               Index           =   5
               Left            =   120
               TabIndex        =   581
               Top             =   1080
               Width           =   810
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "UCR"
               Height          =   270
               Index           =   5
               Left            =   120
               TabIndex        =   580
               Top             =   240
               Width           =   810
            End
         End
         Begin VB.Frame pinfoframe 
            BackColor       =   &H00808000&
            Caption         =   "Property Info"
            Height          =   3900
            Index           =   4
            Left            =   11160
            TabIndex        =   562
            Top             =   9840
            Visible         =   0   'False
            Width           =   4325
            Begin VB.CommandButton Command22 
               Caption         =   "Clear"
               Height          =   270
               Index           =   4
               Left            =   3420
               TabIndex        =   568
               Top             =   3525
               Width           =   645
            End
            Begin VB.CommandButton Command10 
               Caption         =   "C   L   O   S   E"
               Height          =   300
               Index           =   4
               Left            =   60
               TabIndex        =   567
               Top             =   3495
               Width           =   3060
            End
            Begin VB.ListBox minorlist 
               Height          =   645
               Index           =   4
               Left            =   915
               TabIndex        =   566
               Top             =   2700
               Width           =   3105
            End
            Begin VB.ListBox majorlist 
               Height          =   645
               Index           =   4
               Left            =   915
               TabIndex        =   565
               Top             =   1890
               Width           =   3105
            End
            Begin VB.ListBox group 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   645
               Index           =   4
               Left            =   915
               TabIndex        =   564
               Top             =   1080
               Width           =   3105
            End
            Begin VB.ListBox pucrlist 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   645
               Index           =   4
               Left            =   915
               Sorted          =   -1  'True
               TabIndex        =   563
               Top             =   240
               Width           =   3105
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Minor Grouping"
               Height          =   450
               Index           =   4
               Left            =   120
               TabIndex        =   572
               Top             =   2730
               Width           =   810
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Major Grouping"
               Height          =   450
               Index           =   4
               Left            =   105
               TabIndex        =   571
               Top             =   1920
               Width           =   810
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Group"
               Height          =   270
               Index           =   4
               Left            =   120
               TabIndex        =   570
               Top             =   1080
               Width           =   810
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "UCR"
               Height          =   270
               Index           =   4
               Left            =   120
               TabIndex        =   569
               Top             =   240
               Width           =   810
            End
         End
         Begin VB.Frame pinfoframe 
            BackColor       =   &H00808000&
            Caption         =   "Property Info"
            Height          =   3900
            Index           =   3
            Left            =   11160
            TabIndex        =   551
            Top             =   9120
            Visible         =   0   'False
            Width           =   4325
            Begin VB.CommandButton Command22 
               Caption         =   "Clear"
               Height          =   270
               Index           =   3
               Left            =   3420
               TabIndex        =   557
               Top             =   3525
               Width           =   645
            End
            Begin VB.CommandButton Command10 
               Caption         =   "C   L   O   S   E"
               Height          =   300
               Index           =   3
               Left            =   60
               TabIndex        =   556
               Top             =   3495
               Width           =   3060
            End
            Begin VB.ListBox minorlist 
               Height          =   645
               Index           =   3
               Left            =   915
               TabIndex        =   555
               Top             =   2700
               Width           =   3105
            End
            Begin VB.ListBox majorlist 
               Height          =   645
               Index           =   3
               Left            =   915
               TabIndex        =   554
               Top             =   1890
               Width           =   3105
            End
            Begin VB.ListBox group 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   645
               Index           =   3
               Left            =   915
               TabIndex        =   553
               Top             =   1080
               Width           =   3105
            End
            Begin VB.ListBox pucrlist 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   645
               Index           =   3
               Left            =   915
               Sorted          =   -1  'True
               TabIndex        =   552
               Top             =   270
               Width           =   3105
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Minor Grouping"
               Height          =   450
               Index           =   3
               Left            =   120
               TabIndex        =   561
               Top             =   2730
               Width           =   810
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Major Grouping"
               Height          =   450
               Index           =   3
               Left            =   105
               TabIndex        =   560
               Top             =   1920
               Width           =   810
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Group"
               Height          =   270
               Index           =   3
               Left            =   120
               TabIndex        =   559
               Top             =   1080
               Width           =   810
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "UCR"
               Height          =   270
               Index           =   3
               Left            =   120
               TabIndex        =   558
               Top             =   240
               Width           =   810
            End
         End
         Begin VB.Frame pinfoframe 
            BackColor       =   &H00808000&
            Caption         =   "Property Info"
            Height          =   3900
            Index           =   2
            Left            =   11040
            TabIndex        =   540
            Top             =   7800
            Visible         =   0   'False
            Width           =   4325
            Begin VB.CommandButton Command22 
               Caption         =   "Clear"
               Height          =   270
               Index           =   2
               Left            =   3420
               TabIndex        =   546
               Top             =   3525
               Width           =   645
            End
            Begin VB.CommandButton Command10 
               Caption         =   "C   L   O   S   E"
               Height          =   300
               Index           =   2
               Left            =   60
               TabIndex        =   545
               Top             =   3495
               Width           =   3060
            End
            Begin VB.ListBox minorlist 
               Height          =   645
               Index           =   2
               Left            =   915
               TabIndex        =   544
               Top             =   2700
               Width           =   3105
            End
            Begin VB.ListBox majorlist 
               Height          =   645
               Index           =   2
               Left            =   915
               TabIndex        =   543
               Top             =   1890
               Width           =   3105
            End
            Begin VB.ListBox group 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   645
               Index           =   2
               Left            =   915
               TabIndex        =   542
               Top             =   1080
               Width           =   3105
            End
            Begin VB.ListBox pucrlist 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   645
               Index           =   2
               Left            =   915
               Sorted          =   -1  'True
               TabIndex        =   541
               Top             =   270
               Width           =   3105
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Minor Grouping"
               Height          =   450
               Index           =   2
               Left            =   120
               TabIndex        =   550
               Top             =   2730
               Width           =   810
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Major Grouping"
               Height          =   450
               Index           =   2
               Left            =   105
               TabIndex        =   549
               Top             =   1920
               Width           =   810
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Group"
               Height          =   270
               Index           =   2
               Left            =   120
               TabIndex        =   548
               Top             =   1080
               Width           =   810
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "UCR"
               Height          =   270
               Index           =   2
               Left            =   120
               TabIndex        =   547
               Top             =   240
               Width           =   810
            End
         End
         Begin VB.Frame pinfoframe 
            BackColor       =   &H00808000&
            Caption         =   "Property Info"
            Height          =   3900
            Index           =   0
            Left            =   1080
            TabIndex        =   479
            Top             =   2520
            Visible         =   0   'False
            Width           =   4325
            Begin VB.CommandButton Command22 
               Caption         =   "Clear"
               Height          =   270
               Index           =   0
               Left            =   3420
               TabIndex        =   528
               Top             =   3510
               Width           =   645
            End
            Begin VB.CommandButton Command10 
               Caption         =   "C   L   O   S   E"
               Height          =   300
               Index           =   0
               Left            =   60
               TabIndex        =   484
               Top             =   3495
               Width           =   3060
            End
            Begin VB.ListBox minorlist 
               Height          =   645
               Index           =   0
               Left            =   915
               TabIndex        =   483
               Top             =   2700
               Width           =   3105
            End
            Begin VB.ListBox majorlist 
               Height          =   645
               Index           =   0
               Left            =   915
               TabIndex        =   482
               Top             =   1890
               Width           =   3105
            End
            Begin VB.ListBox group 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   645
               Index           =   0
               Left            =   915
               TabIndex        =   481
               Top             =   1080
               Width           =   3105
            End
            Begin VB.ListBox pucrlist 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   645
               Index           =   0
               Left            =   915
               Sorted          =   -1  'True
               TabIndex        =   480
               Top             =   270
               Width           =   3105
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Minor Grouping"
               Height          =   450
               Index           =   0
               Left            =   120
               TabIndex        =   488
               Top             =   2730
               Width           =   810
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Major Grouping"
               Height          =   450
               Index           =   0
               Left            =   105
               TabIndex        =   487
               Top             =   1920
               Width           =   810
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Group"
               Height          =   270
               Index           =   0
               Left            =   120
               TabIndex        =   486
               Top             =   1080
               Width           =   810
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "UCR"
               Height          =   270
               Index           =   0
               Left            =   120
               TabIndex        =   485
               Top             =   240
               Width           =   810
            End
         End
         Begin VB.ListBox BIAS 
            BackColor       =   &H00404000&
            ForeColor       =   &H00FFFFFF&
            Height          =   840
            Left            =   1440
            Sorted          =   -1  'True
            TabIndex        =   317
            Top             =   16200
            Visible         =   0   'False
            Width           =   2265
         End
         Begin MSComctlLib.ListView gactivity 
            Height          =   735
            Index           =   1
            Left            =   4200
            TabIndex        =   371
            Top             =   5280
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   1296
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   16777215
            BackColor       =   4210688
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   5644
            EndProperty
         End
         Begin VB.ListBox UCRLIST 
            BackColor       =   &H00404000&
            ForeColor       =   &H00FFFFFF&
            Height          =   735
            Index           =   2
            Left            =   3840
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   274
            Top             =   4080
            Visible         =   0   'False
            Width           =   4170
         End
         Begin MSComctlLib.ListView gactivity 
            Height          =   735
            Index           =   0
            Left            =   4320
            TabIndex        =   370
            Top             =   3000
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   1296
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   16777215
            BackColor       =   4210688
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   5644
            EndProperty
         End
         Begin VB.ListBox UCRLIST 
            BackColor       =   &H00404000&
            ForeColor       =   &H00FFFFFF&
            Height          =   735
            Index           =   1
            Left            =   5160
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   273
            Top             =   2280
            Visible         =   0   'False
            Width           =   4170
         End
         Begin VB.ListBox UCRLIST 
            BackColor       =   &H00404000&
            ForeColor       =   &H00FFFFFF&
            Height          =   735
            Index           =   0
            Left            =   4080
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   272
            Top             =   2280
            Visible         =   0   'False
            Width           =   4170
         End
         Begin MSComctlLib.ListView activity 
            Height          =   735
            Index           =   2
            Left            =   1800
            TabIndex        =   277
            Top             =   2280
            Visible         =   0   'False
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   1296
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   16777215
            BackColor       =   4210688
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   5644
            EndProperty
         End
         Begin MSComctlLib.ListView activity 
            Height          =   735
            Index           =   1
            Left            =   960
            TabIndex        =   276
            Top             =   2280
            Visible         =   0   'False
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   1296
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   16777215
            BackColor       =   4210688
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   5644
            EndProperty
         End
         Begin MSComctlLib.ListView sublist 
            Height          =   1425
            Index           =   0
            Left            =   360
            TabIndex        =   418
            Top             =   1080
            Visible         =   0   'False
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2514
            View            =   3
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   16777215
            BackColor       =   4210688
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   5644
            EndProperty
         End
         Begin MSComctlLib.ListView activity 
            Height          =   735
            Index           =   0
            Left            =   1680
            TabIndex        =   275
            Top             =   1320
            Visible         =   0   'False
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   1296
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   16777215
            BackColor       =   4210688
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   5644
            EndProperty
         End
         Begin MSComctlLib.ListView HOMOCIDE 
            Height          =   735
            Index           =   2
            Left            =   4200
            TabIndex        =   430
            Top             =   720
            Visible         =   0   'False
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   1296
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   16777215
            BackColor       =   0
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   5733
            EndProperty
         End
         Begin VB.ListBox lactivity 
            BackColor       =   &H00404000&
            ForeColor       =   &H00FFFFFF&
            Height          =   645
            Left            =   3840
            Sorted          =   -1  'True
            TabIndex        =   434
            Top             =   960
            Visible         =   0   'False
            Width           =   5000
         End
         Begin MSComctlLib.ListView HOMOCIDE 
            Height          =   735
            Index           =   0
            Left            =   3840
            TabIndex        =   428
            Top             =   720
            Visible         =   0   'False
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   1296
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   16777215
            BackColor       =   0
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   5733
            EndProperty
         End
         Begin MSComctlLib.ListView HOMOCIDE 
            Height          =   735
            Index           =   1
            Left            =   3840
            TabIndex        =   429
            Top             =   600
            Visible         =   0   'False
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   1296
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   16777215
            BackColor       =   0
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   5733
            EndProperty
         End
         Begin VB.ListBox additional 
            BackColor       =   &H00404000&
            ForeColor       =   &H00FFFFFF&
            Height          =   645
            Index           =   2
            Left            =   240
            Sorted          =   -1  'True
            TabIndex        =   432
            Top             =   1080
            Visible         =   0   'False
            Width           =   3465
         End
         Begin VB.ListBox additional 
            BackColor       =   &H00404000&
            ForeColor       =   &H00FFFFFF&
            Height          =   645
            Index           =   0
            Left            =   240
            Sorted          =   -1  'True
            TabIndex        =   283
            Top             =   840
            Visible         =   0   'False
            Width           =   3465
         End
         Begin VB.ListBox additional 
            BackColor       =   &H00404000&
            ForeColor       =   &H00FFFFFF&
            Height          =   645
            Index           =   1
            Left            =   240
            Sorted          =   -1  'True
            TabIndex        =   431
            Top             =   600
            Visible         =   0   'False
            Width           =   3465
         End
         Begin VB.Frame offenseframe 
            Caption         =   "Additional Offenses"
            Height          =   3030
            Left            =   11400
            TabIndex        =   489
            Top             =   2280
            Visible         =   0   'False
            Width           =   11040
            Begin MSComctlLib.ListView gactivity 
               Height          =   735
               Index           =   4
               Left            =   2040
               TabIndex        =   525
               Top             =   2520
               Visible         =   0   'False
               Width           =   4245
               _ExtentX        =   7488
               _ExtentY        =   1296
               View            =   3
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               HideColumnHeaders=   -1  'True
               _Version        =   393217
               ForeColor       =   16777215
               BackColor       =   4210688
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Object.Width           =   5644
               EndProperty
            End
            Begin VB.ListBox additional 
               BackColor       =   &H00404000&
               ForeColor       =   &H00FFFFFF&
               Height          =   645
               Index           =   4
               Left            =   4200
               Sorted          =   -1  'True
               TabIndex        =   521
               Top             =   2520
               Visible         =   0   'False
               Width           =   3465
            End
            Begin MSComctlLib.ListView HOMOCIDE 
               Height          =   735
               Index           =   4
               Left            =   6720
               TabIndex        =   527
               Top             =   2400
               Visible         =   0   'False
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   1296
               View            =   3
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               HideColumnHeaders=   -1  'True
               _Version        =   393217
               ForeColor       =   16777215
               BackColor       =   0
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Object.Width           =   5733
               EndProperty
            End
            Begin MSComctlLib.ListView HOMOCIDE 
               Height          =   735
               Index           =   3
               Left            =   6480
               TabIndex        =   526
               Top             =   1800
               Visible         =   0   'False
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   1296
               View            =   3
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               HideColumnHeaders=   -1  'True
               _Version        =   393217
               ForeColor       =   16777215
               BackColor       =   0
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Object.Width           =   5733
               EndProperty
            End
            Begin MSComctlLib.ListView gactivity 
               Height          =   735
               Index           =   3
               Left            =   240
               TabIndex        =   524
               Top             =   1800
               Visible         =   0   'False
               Width           =   4245
               _ExtentX        =   7488
               _ExtentY        =   1296
               View            =   3
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               HideColumnHeaders=   -1  'True
               _Version        =   393217
               ForeColor       =   16777215
               BackColor       =   4210688
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Object.Width           =   5644
               EndProperty
            End
            Begin MSComctlLib.ListView activity 
               Height          =   735
               Index           =   4
               Left            =   240
               TabIndex        =   522
               Top             =   2040
               Visible         =   0   'False
               Width           =   3495
               _ExtentX        =   6165
               _ExtentY        =   1296
               View            =   3
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               HideColumnHeaders=   -1  'True
               _Version        =   393217
               ForeColor       =   16777215
               BackColor       =   4210688
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Object.Width           =   5644
               EndProperty
            End
            Begin MSComctlLib.ListView activity 
               Height          =   735
               Index           =   3
               Left            =   5760
               TabIndex        =   523
               Top             =   2040
               Visible         =   0   'False
               Width           =   3495
               _ExtentX        =   6165
               _ExtentY        =   1296
               View            =   3
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               HideColumnHeaders=   -1  'True
               _Version        =   393217
               ForeColor       =   16777215
               BackColor       =   4210688
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Object.Width           =   5644
               EndProperty
            End
            Begin VB.ListBox additional 
               BackColor       =   &H00404000&
               ForeColor       =   &H00FFFFFF&
               Height          =   645
               Index           =   3
               Left            =   3720
               Sorted          =   -1  'True
               TabIndex        =   520
               Top             =   2160
               Visible         =   0   'False
               Width           =   3465
            End
            Begin VB.ListBox UCRLIST 
               BackColor       =   &H00404000&
               ForeColor       =   &H00FFFFFF&
               Height          =   735
               Index           =   4
               Left            =   1560
               Sorted          =   -1  'True
               Style           =   1  'Checkbox
               TabIndex        =   519
               Top             =   2040
               Visible         =   0   'False
               Width           =   4170
            End
            Begin VB.ListBox UCRLIST 
               BackColor       =   &H00404000&
               ForeColor       =   &H00FFFFFF&
               Height          =   735
               Index           =   3
               Left            =   960
               Sorted          =   -1  'True
               Style           =   1  'Checkbox
               TabIndex        =   518
               Top             =   1920
               Visible         =   0   'False
               Width           =   4170
            End
            Begin MSComctlLib.ListView premise 
               Height          =   300
               Index           =   4
               Left            =   6890
               TabIndex        =   512
               Top             =   1425
               Width           =   1520
               _ExtentX        =   2699
               _ExtentY        =   529
               View            =   3
               Sorted          =   -1  'True
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               HideColumnHeaders=   -1  'True
               _Version        =   393217
               ForeColor       =   16777215
               BackColor       =   4210688
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Object.Width           =   5644
               EndProperty
            End
            Begin MSComctlLib.ListView premise 
               Height          =   300
               Index           =   3
               Left            =   6890
               TabIndex        =   500
               Top             =   1100
               Width           =   1520
               _ExtentX        =   2699
               _ExtentY        =   529
               View            =   3
               Sorted          =   -1  'True
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               HideColumnHeaders=   -1  'True
               _Version        =   393217
               ForeColor       =   16777215
               BackColor       =   4210688
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Object.Width           =   5644
               EndProperty
            End
            Begin VB.CommandButton subcode 
               Caption         =   "Sub"
               Height          =   255
               Index           =   4
               Left            =   4050
               TabIndex        =   507
               Top             =   1440
               Width           =   400
            End
            Begin VB.CommandButton subcode 
               Caption         =   "Sub"
               Height          =   255
               Index           =   3
               Left            =   4050
               TabIndex        =   495
               Top             =   1110
               Width           =   400
            End
            Begin VB.CommandButton Command13 
               Caption         =   "Rsn"
               Height          =   255
               Index           =   4
               Left            =   3645
               TabIndex        =   506
               Top             =   1440
               Width           =   400
            End
            Begin VB.CommandButton Command12 
               Caption         =   "Act"
               Height          =   255
               Index           =   4
               Left            =   3240
               TabIndex        =   505
               Top             =   1440
               Width           =   375
            End
            Begin VB.CommandButton Command11 
               Caption         =   "UCR"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   2800
               TabIndex        =   504
               Top             =   1440
               Width           =   450
            End
            Begin VB.CommandButton Command13 
               Caption         =   "Rsn"
               Height          =   255
               Index           =   3
               Left            =   3645
               TabIndex        =   494
               Top             =   1110
               Width           =   400
            End
            Begin VB.CommandButton Command12 
               Caption         =   "Act"
               Height          =   255
               Index           =   3
               Left            =   3270
               TabIndex        =   493
               Top             =   1110
               Width           =   375
            End
            Begin VB.CommandButton Command11 
               Caption         =   "UCR"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   2800
               TabIndex        =   492
               Top             =   1110
               Width           =   450
            End
            Begin VB.TextBox entered 
               BackColor       =   &H00C0C0C0&
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   4
               Left            =   8475
               TabIndex        =   513
               Top             =   1425
               Width           =   750
            End
            Begin VB.TextBox entered 
               BackColor       =   &H00C0C0C0&
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   3
               Left            =   8475
               TabIndex        =   501
               Top             =   1080
               Width           =   750
            End
            Begin VB.Frame Frame12 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Caption         =   "Frame12"
               Height          =   255
               Index           =   7
               Left            =   4520
               TabIndex        =   517
               Top             =   1425
               Width           =   1100
               Begin VB.OptionButton completedy 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "YES"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   6
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   4
                  Left            =   0
                  TabIndex        =   508
                  Top             =   0
                  Width           =   615
               End
               Begin VB.OptionButton completedn 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "NO"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   6
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   4
                  Left            =   600
                  TabIndex        =   509
                  Top             =   0
                  Width           =   495
               End
            End
            Begin VB.Frame Frame12 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Caption         =   "Frame12"
               Height          =   255
               Index           =   6
               Left            =   4520
               TabIndex        =   516
               Top             =   1140
               Width           =   1100
               Begin VB.OptionButton completedy 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "YES"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   6
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   3
                  Left            =   0
                  TabIndex        =   496
                  Top             =   0
                  Width           =   615
               End
               Begin VB.OptionButton completedn 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "NO"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   6
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   3
                  Left            =   600
                  TabIndex        =   497
                  Top             =   0
                  Width           =   495
               End
            End
            Begin VB.CommandButton closeoffense 
               Caption         =   "Close"
               Height          =   315
               Left            =   8220
               TabIndex        =   514
               Top             =   270
               Width           =   1095
            End
            Begin VB.CommandButton Command3 
               Caption         =   "Find"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   5
               Left            =   165
               TabIndex        =   515
               TabStop         =   0   'False
               Top             =   1405
               Width           =   400
            End
            Begin VB.CommandButton Command3 
               Caption         =   "Find"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   4
               Left            =   165
               TabIndex        =   502
               TabStop         =   0   'False
               Top             =   1095
               Width           =   400
            End
            Begin VB.ListBox pickoffense 
               Height          =   255
               Index           =   3
               Left            =   600
               TabIndex        =   491
               Top             =   1110
               Width           =   2200
            End
            Begin VB.ListBox pickoffense 
               Height          =   255
               Index           =   4
               Left            =   600
               TabIndex        =   503
               Top             =   1420
               Width           =   2200
            End
            Begin VB.CheckBox FORCEDENTRYN 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Check1"
               Height          =   255
               Index           =   3
               Left            =   6615
               TabIndex        =   499
               Top             =   1125
               Width           =   150
            End
            Begin VB.CheckBox FORCEDENTRYN 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Check1"
               Height          =   255
               Index           =   4
               Left            =   6600
               TabIndex        =   511
               Top             =   1440
               Width           =   150
            End
            Begin VB.CheckBox FORCEDENTRYY 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Check1"
               Height          =   255
               Index           =   4
               Left            =   5720
               TabIndex        =   510
               Top             =   1440
               Width           =   150
            End
            Begin VB.CheckBox FORCEDENTRYY 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Check1"
               Height          =   255
               Index           =   3
               Left            =   5720
               TabIndex        =   498
               Top             =   1125
               Width           =   150
            End
            Begin VB.PictureBox Picture3 
               Height          =   1000
               Left            =   555
               Picture         =   "ninciden.frx":13726
               ScaleHeight     =   945
               ScaleWidth      =   8715
               TabIndex        =   490
               Top             =   720
               Width           =   8775
            End
         End
         Begin VB.Frame pinfoframe 
            BackColor       =   &H00808000&
            Caption         =   "Property Info"
            Height          =   3900
            Index           =   1
            Left            =   11040
            TabIndex        =   529
            Top             =   9840
            Visible         =   0   'False
            Width           =   4325
            Begin VB.ListBox pucrlist 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   645
               Index           =   1
               Left            =   915
               Sorted          =   -1  'True
               TabIndex        =   530
               Top             =   270
               Width           =   3105
            End
            Begin VB.ListBox group 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   645
               Index           =   1
               ItemData        =   "ninciden.frx":5C704
               Left            =   915
               List            =   "ninciden.frx":5C706
               TabIndex        =   531
               Top             =   1080
               Width           =   3105
            End
            Begin VB.ListBox majorlist 
               Height          =   645
               Index           =   1
               Left            =   915
               TabIndex        =   532
               Top             =   1890
               Width           =   3105
            End
            Begin VB.ListBox minorlist 
               Height          =   645
               Index           =   1
               Left            =   840
               TabIndex        =   533
               Top             =   2700
               Width           =   3105
            End
            Begin VB.CommandButton Command10 
               Caption         =   "C   L   O   S   E"
               Height          =   300
               Index           =   1
               Left            =   60
               TabIndex        =   534
               Top             =   3495
               Width           =   3060
            End
            Begin VB.CommandButton Command22 
               Caption         =   "Clear"
               Height          =   270
               Index           =   1
               Left            =   3420
               TabIndex        =   535
               Top             =   3525
               Width           =   645
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "UCR"
               Height          =   270
               Index           =   1
               Left            =   120
               TabIndex        =   539
               Top             =   240
               Width           =   810
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Group"
               Height          =   270
               Index           =   1
               Left            =   120
               TabIndex        =   538
               Top             =   1080
               Width           =   810
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Major Grouping"
               Height          =   450
               Index           =   1
               Left            =   105
               TabIndex        =   537
               Top             =   1920
               Width           =   810
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Minor Grouping"
               Height          =   450
               Index           =   1
               Left            =   120
               TabIndex        =   536
               Top             =   2730
               Width           =   810
            End
         End
         Begin VB.Frame relationshipframe 
            BackColor       =   &H00808000&
            Caption         =   "Relationship to Subject Number:"
            ForeColor       =   &H000000FF&
            Height          =   3255
            Index           =   1
            Left            =   11400
            TabIndex        =   466
            Top             =   6120
            Visible         =   0   'False
            Width           =   3975
            Begin VB.ListBox relationship 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   10
               Left            =   150
               Style           =   1  'Checkbox
               TabIndex        =   467
               Top             =   175
               Width           =   1800
            End
            Begin VB.ListBox relationship 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   11
               Left            =   150
               Style           =   1  'Checkbox
               TabIndex        =   468
               Top             =   725
               Width           =   1800
            End
            Begin VB.ListBox relationship 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   12
               Left            =   150
               Style           =   1  'Checkbox
               TabIndex        =   469
               Top             =   1275
               Width           =   1800
            End
            Begin VB.ListBox relationship 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   19
               Left            =   2075
               Style           =   1  'Checkbox
               TabIndex        =   476
               Top             =   2375
               Width           =   1800
            End
            Begin VB.ListBox relationship 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   18
               Left            =   2075
               Style           =   1  'Checkbox
               TabIndex        =   475
               Top             =   1825
               Width           =   1800
            End
            Begin VB.ListBox relationship 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   17
               Left            =   2075
               Style           =   1  'Checkbox
               TabIndex        =   474
               Top             =   1275
               Width           =   1800
            End
            Begin VB.ListBox relationship 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   16
               Left            =   2075
               Style           =   1  'Checkbox
               TabIndex        =   473
               Top             =   725
               Width           =   1800
            End
            Begin VB.ListBox relationship 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   15
               Left            =   2085
               Style           =   1  'Checkbox
               TabIndex        =   472
               Top             =   195
               Width           =   1800
            End
            Begin VB.ListBox relationship 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   14
               Left            =   150
               Style           =   1  'Checkbox
               TabIndex        =   471
               Top             =   2375
               Width           =   1800
            End
            Begin VB.ListBox relationship 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   13
               Left            =   150
               Style           =   1  'Checkbox
               TabIndex        =   470
               Top             =   1825
               Width           =   1800
            End
            Begin VB.CommandButton Command8 
               Caption         =   "Close"
               Height          =   195
               Index           =   1
               Left            =   1080
               TabIndex        =   477
               Top             =   3000
               Width           =   1815
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   $"ninciden.frx":5C708
               ForeColor       =   &H0000FFFF&
               Height          =   2895
               Index           =   1
               Left            =   0
               TabIndex        =   478
               Top             =   120
               Width           =   2175
            End
         End
         Begin VB.Frame relationshipframe 
            BackColor       =   &H00808000&
            Caption         =   "Relationship to Subject Number:"
            ForeColor       =   &H000000FF&
            Height          =   3255
            Index           =   0
            Left            =   11280
            TabIndex        =   332
            Top             =   6240
            Visible         =   0   'False
            Width           =   3975
            Begin VB.CommandButton Command8 
               Caption         =   "Close"
               Height          =   195
               Index           =   0
               Left            =   1080
               TabIndex        =   343
               Top             =   3000
               Width           =   1815
            End
            Begin VB.ListBox relationship 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   3
               Left            =   150
               Style           =   1  'Checkbox
               TabIndex        =   336
               Top             =   1825
               Width           =   1800
            End
            Begin VB.ListBox relationship 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   4
               Left            =   150
               Style           =   1  'Checkbox
               TabIndex        =   337
               Top             =   2375
               Width           =   1800
            End
            Begin VB.ListBox relationship 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   5
               Left            =   2075
               Style           =   1  'Checkbox
               TabIndex        =   338
               Top             =   175
               Width           =   1800
            End
            Begin VB.ListBox relationship 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   6
               Left            =   2075
               Style           =   1  'Checkbox
               TabIndex        =   339
               Top             =   725
               Width           =   1800
            End
            Begin VB.ListBox relationship 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   7
               Left            =   2075
               Style           =   1  'Checkbox
               TabIndex        =   340
               Top             =   1275
               Width           =   1800
            End
            Begin VB.ListBox relationship 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   8
               Left            =   2075
               Style           =   1  'Checkbox
               TabIndex        =   341
               Top             =   1825
               Width           =   1800
            End
            Begin VB.ListBox relationship 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   9
               Left            =   2075
               Style           =   1  'Checkbox
               TabIndex        =   342
               Top             =   2375
               Width           =   1800
            End
            Begin VB.ListBox relationship 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   2
               Left            =   150
               Style           =   1  'Checkbox
               TabIndex        =   335
               Top             =   1275
               Width           =   1800
            End
            Begin VB.ListBox relationship 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   1
               Left            =   150
               Style           =   1  'Checkbox
               TabIndex        =   334
               Top             =   725
               Width           =   1800
            End
            Begin VB.ListBox relationship 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   0
               Left            =   150
               Style           =   1  'Checkbox
               TabIndex        =   333
               Top             =   175
               Width           =   1800
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   $"ninciden.frx":5C9AC
               ForeColor       =   &H0000FFFF&
               Height          =   2895
               Index           =   0
               Left            =   3960
               TabIndex        =   344
               Top             =   240
               Width           =   2175
            End
         End
         Begin VB.Frame msframe 
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   3240
            TabIndex        =   584
            Top             =   8745
            Width           =   750
            Begin VB.Image mugshot 
               BorderStyle     =   1  'Fixed Single
               Height          =   645
               Left            =   0
               Stretch         =   -1  'True
               Top             =   -15
               Width           =   750
            End
         End
         Begin VB.Frame caseframe 
            Caption         =   "Case Setup Frame"
            Height          =   2895
            Left            =   11070
            TabIndex        =   446
            Top             =   13380
            Visible         =   0   'False
            Width           =   8175
            Begin VB.Frame Frame8 
               Height          =   1095
               Left            =   3480
               TabIndex        =   464
               Top             =   600
               Width           =   1575
               Begin VB.OptionButton year45 
                  Caption         =   "2-digit Year"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   452
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.OptionButton month45 
                  Caption         =   "2-digit Month"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   453
                  Top             =   600
                  Value           =   -1  'True
                  Width           =   1335
               End
            End
            Begin VB.CommandButton Command21 
               Caption         =   "Close"
               Height          =   375
               Left            =   6480
               TabIndex        =   456
               Top             =   2400
               Width           =   1575
            End
            Begin VB.CommandButton Command20 
               Caption         =   "Update"
               Height          =   375
               Left            =   120
               TabIndex        =   455
               Top             =   2280
               Width           =   1575
            End
            Begin VB.Frame Frame7 
               Height          =   1095
               Left            =   6720
               TabIndex        =   463
               Top             =   600
               Width           =   1335
               Begin VB.TextBox suffix 
                  Height          =   285
                  Left            =   120
                  MaxLength       =   5
                  TabIndex        =   454
                  Top             =   480
                  Width           =   1095
               End
            End
            Begin VB.Frame Frame6 
               Height          =   1095
               Left            =   5160
               TabIndex        =   460
               Top             =   600
               Width           =   1455
               Begin VB.Label Label8 
                  Caption         =   "Incremented 5-digit number"
                  ForeColor       =   &H00000000&
                  Height          =   615
                  Left            =   240
                  TabIndex        =   461
                  Top             =   360
                  Width           =   1095
               End
            End
            Begin VB.Frame Frame4 
               Height          =   1095
               Left            =   1800
               TabIndex        =   458
               Top             =   600
               Width           =   1575
               Begin VB.CheckBox dash3 
                  Caption         =   "Optional Dash"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   451
                  Top             =   480
                  Width           =   1335
               End
            End
            Begin VB.Frame Frame3 
               Height          =   1095
               Left            =   120
               TabIndex        =   448
               Top             =   600
               Width           =   1575
               Begin VB.OptionButton month12 
                  Caption         =   "2-digit Month"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   450
                  Top             =   600
                  Value           =   -1  'True
                  Width           =   1335
               End
               Begin VB.OptionButton year12 
                  Caption         =   "2-digit Year"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   449
                  Top             =   240
                  Width           =   1335
               End
            End
            Begin VB.Label Label10 
               Caption         =   "Digits 4-5"
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   3840
               TabIndex        =   465
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label9 
               Caption         =   "Optional Suffix"
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   6840
               TabIndex        =   462
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label7 
               Caption         =   "Digits 6-10 or 5-9 (based on Dash)"
               ForeColor       =   &H000000C0&
               Height          =   495
               Left            =   5280
               TabIndex        =   459
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Digit 3"
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   2400
               TabIndex        =   457
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label5 
               Caption         =   "Digits 1-2"
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   480
               TabIndex        =   447
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.ComboBox incidentnumber 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   9255
            TabIndex        =   0
            Top             =   285
            Width           =   2295
         End
         Begin VB.Frame lstframe 
            Caption         =   "Spelling Suggestions"
            Height          =   2415
            Left            =   11280
            TabIndex        =   439
            Top             =   14895
            Visible         =   0   'False
            Width           =   4935
            Begin VB.CommandButton Command18 
               Caption         =   "Skip"
               Height          =   375
               Left            =   3720
               TabIndex        =   443
               Top             =   960
               Width           =   1095
            End
            Begin VB.CommandButton Command17 
               Caption         =   "Change"
               Height          =   375
               Left            =   3720
               TabIndex        =   442
               Top             =   480
               Width           =   1095
            End
            Begin VB.CommandButton Command16 
               Caption         =   "Close"
               Height          =   375
               Left            =   3720
               TabIndex        =   444
               Top             =   1920
               Width           =   1095
            End
            Begin VB.ListBox lstsuggestions 
               Height          =   1815
               Left            =   120
               TabIndex        =   440
               Top             =   480
               Width           =   3495
            End
            Begin VB.Label checkword 
               Height          =   255
               Left            =   240
               TabIndex        =   441
               Top             =   240
               Width           =   3495
            End
         End
         Begin VB.CheckBox computerequipment 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Computer"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   1
            Left            =   4320
            TabIndex        =   149
            TabStop         =   0   'False
            Top             =   11280
            Width           =   1000
         End
         Begin VB.CommandButton Command14 
            BackColor       =   &H00808000&
            Caption         =   "Act"
            Height          =   255
            Left            =   11040
            Style           =   1  'Graphical
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   2520
            Width           =   495
         End
         Begin VB.Timer optimer 
            Enabled         =   0   'False
            Interval        =   500
            Left            =   4560
            Top             =   360
         End
         Begin MSComctlLib.ListView gactivity 
            Height          =   735
            Index           =   2
            Left            =   11520
            TabIndex        =   372
            Top             =   1440
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   1296
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   16777215
            BackColor       =   4210688
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   5644
            EndProperty
         End
         Begin MSComctlLib.ListView premise 
            Height          =   300
            Index           =   0
            Left            =   6910
            TabIndex        =   14
            Top             =   1920
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   529
            View            =   3
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   16777215
            BackColor       =   8421376
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   5644
            EndProperty
         End
         Begin MSComctlLib.ListView premise 
            Height          =   300
            Index           =   1
            Left            =   6910
            TabIndex        =   26
            Top             =   2250
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   529
            View            =   3
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   16777215
            BackColor       =   8421376
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   5644
            EndProperty
         End
         Begin MSComctlLib.ListView premise 
            Height          =   300
            Index           =   2
            Left            =   6910
            TabIndex        =   38
            Top             =   2575
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   529
            View            =   3
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   16777215
            BackColor       =   8421376
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   5644
            EndProperty
         End
         Begin MSComctlLib.ListView sublist 
            Height          =   1425
            Index           =   1
            Left            =   4320
            TabIndex        =   419
            Top             =   -1080
            Visible         =   0   'False
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2514
            View            =   3
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   16777215
            BackColor       =   4210688
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   5644
            EndProperty
         End
         Begin MSComctlLib.ListView sublist 
            Height          =   1425
            Index           =   2
            Left            =   4320
            TabIndex        =   420
            Top             =   -1080
            Visible         =   0   'False
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2514
            View            =   3
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   16777215
            BackColor       =   4210688
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   5644
            EndProperty
         End
         Begin MSComctlLib.ListView sublist 
            Height          =   1425
            Index           =   3
            Left            =   5280
            TabIndex        =   421
            Top             =   -1080
            Visible         =   0   'False
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2514
            View            =   3
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   16777215
            BackColor       =   4210688
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   5644
            EndProperty
         End
         Begin MSComctlLib.ListView sublist 
            Height          =   1425
            Index           =   4
            Left            =   4920
            TabIndex        =   422
            Top             =   -1200
            Visible         =   0   'False
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2514
            View            =   3
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   16777215
            BackColor       =   4210688
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   5644
            EndProperty
         End
         Begin MSComctlLib.ListView sublist 
            Height          =   1425
            Index           =   5
            Left            =   3960
            TabIndex        =   423
            Top             =   -1080
            Visible         =   0   'False
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2514
            View            =   3
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   16777215
            BackColor       =   4210688
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   5644
            EndProperty
         End
         Begin MSComctlLib.ListView sublist 
            Height          =   1425
            Index           =   6
            Left            =   4440
            TabIndex        =   424
            Top             =   -1080
            Visible         =   0   'False
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2514
            View            =   3
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   16777215
            BackColor       =   4210688
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   5644
            EndProperty
         End
         Begin MSComctlLib.ListView sublist 
            Height          =   1425
            Index           =   7
            Left            =   4680
            TabIndex        =   425
            Top             =   -1080
            Visible         =   0   'False
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2514
            View            =   3
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   16777215
            BackColor       =   4210688
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   5644
            EndProperty
         End
         Begin MSComctlLib.ListView sublist 
            Height          =   1425
            Index           =   8
            Left            =   5040
            TabIndex        =   426
            Top             =   -1080
            Visible         =   0   'False
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2514
            View            =   3
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   16777215
            BackColor       =   4210688
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   5644
            EndProperty
         End
         Begin MSComctlLib.ListView sublist 
            Height          =   1425
            Index           =   9
            Left            =   5400
            TabIndex        =   427
            Top             =   -960
            Visible         =   0   'False
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2514
            View            =   3
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   16777215
            BackColor       =   4210688
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   5644
            EndProperty
         End
         Begin VB.CommandButton Command13 
            BackColor       =   &H00808000&
            Caption         =   "Rsn"
            Height          =   255
            Index           =   2
            Left            =   3615
            MaskColor       =   &H00808000&
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   2550
            Width           =   400
         End
         Begin VB.CommandButton Command13 
            BackColor       =   &H00808000&
            Caption         =   "Rsn"
            Height          =   255
            Index           =   1
            Left            =   3615
            MaskColor       =   &H00808000&
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   2250
            Width           =   400
         End
         Begin VB.CommandButton Command13 
            BackColor       =   &H00808000&
            Caption         =   "Rsn"
            Height          =   255
            Index           =   0
            Left            =   3615
            MaskColor       =   &H00808000&
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1930
            Width           =   400
         End
         Begin VB.CommandButton Command12 
            BackColor       =   &H00808000&
            Caption         =   "Act"
            Height          =   255
            Index           =   2
            Left            =   3240
            MaskColor       =   &H00808000&
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   2550
            Width           =   375
         End
         Begin VB.CommandButton Command12 
            BackColor       =   &H00808000&
            Caption         =   "Act"
            Height          =   255
            Index           =   1
            Left            =   3225
            MaskColor       =   &H00808000&
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   2250
            Width           =   375
         End
         Begin VB.CommandButton Command12 
            BackColor       =   &H00808000&
            Caption         =   "Act"
            Height          =   255
            Index           =   0
            Left            =   3240
            MaskColor       =   &H00808000&
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   1930
            Width           =   375
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H00808000&
            Caption         =   "UCR"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   2760
            MaskColor       =   &H00808000&
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   2550
            Width           =   450
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H00808000&
            Caption         =   "UCR"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   2760
            MaskColor       =   &H00808000&
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   2250
            Width           =   450
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H00808000&
            Caption         =   "UCR"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   2760
            MaskColor       =   &H00808000&
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1930
            Width           =   450
         End
         Begin Crystal.CrystalReport report 
            Left            =   120
            Top             =   12240
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            Destination     =   1
            PrintFileLinesPerPage=   60
         End
         Begin VB.Data Data1 
            Caption         =   "Data1"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   300
            Left            =   2040
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   0
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.TextBox age 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   5880
            TabIndex        =   129
            Top             =   9405
            Width           =   645
         End
         Begin VB.TextBox age 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   7560
            TabIndex        =   84
            Top             =   6120
            Width           =   525
         End
         Begin VB.TextBox age 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   7600
            TabIndex        =   67
            Top             =   4560
            Width           =   525
         End
         Begin VB.TextBox FOLLOWUPOFFICERDATE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   9960
            TabIndex        =   243
            Top             =   18430
            Width           =   960
         End
         Begin VB.TextBox FOLLOWUPOFFICERUNIT 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   10920
            TabIndex        =   244
            Top             =   18430
            Width           =   660
         End
         Begin VB.TextBox JURISDICTIONRECOVERY 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   9840
            TabIndex        =   157
            Top             =   13900
            Width           =   1575
         End
         Begin VB.TextBox JURISDICTIONTHEFT 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   7920
            TabIndex        =   156
            Top             =   13900
            Width           =   1695
         End
         Begin RichTextLib.RichTextBox NARRATIVE 
            Height          =   2055
            Left            =   480
            TabIndex        =   155
            Top             =   11640
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   3625
            _Version        =   393217
            ScrollBars      =   2
            TextRTF         =   $"ninciden.frx":5CC50
         End
         Begin VB.TextBox TIMEOFARREST 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   10920
            TabIndex        =   154
            Top             =   11270
            Width           =   615
         End
         Begin VB.TextBox DATEOFARREST 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   10065
            TabIndex        =   153
            Top             =   11270
            Width           =   855
         End
         Begin VB.TextBox TOTALARRESTED 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   6480
            TabIndex        =   152
            Top             =   11270
            Width           =   615
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   7320
            TabIndex        =   361
            Top             =   11040
            Width           =   950
            Begin VB.OptionButton ARRESTEDNEARYES 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Yes"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   0
               TabIndex        =   150
               Top             =   0
               Width           =   480
            End
            Begin VB.OptionButton ARRESTEDNEARNO 
               BackColor       =   &H00FFFFFF&
               Caption         =   "No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   480
               TabIndex        =   151
               Top             =   0
               Value           =   -1  'True
               Width           =   510
            End
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   1
            Left            =   1780
            TabIndex        =   360
            Top             =   11200
            Width           =   1600
            Begin VB.OptionButton drugsunknown 
               BackColor       =   &H00FFFFFF&
               Caption         =   "UNK"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   300
               Index           =   1
               Left            =   1080
               TabIndex        =   147
               Top             =   0
               Value           =   -1  'True
               Width           =   585
            End
            Begin VB.OptionButton drugsyes 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Yes"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   300
               Index           =   1
               Left            =   0
               TabIndex        =   145
               Top             =   0
               Width           =   585
            End
            Begin VB.OptionButton drugsno 
               BackColor       =   &H00FFFFFF&
               Caption         =   "No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   300
               Index           =   1
               Left            =   600
               TabIndex        =   146
               Top             =   0
               Width           =   510
            End
         End
         Begin VB.Frame alcoholframe 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   1
            Left            =   3240
            TabIndex        =   359
            Top             =   11040
            Width           =   1575
            Begin VB.OptionButton alcoholno 
               BackColor       =   &H00FFFFFF&
               Caption         =   "No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   195
               Index           =   1
               Left            =   520
               TabIndex        =   143
               Top             =   0
               Width           =   510
            End
            Begin VB.OptionButton alcoholyes 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Yes"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   195
               Index           =   1
               Left            =   0
               TabIndex        =   142
               Top             =   0
               Width           =   600
            End
            Begin VB.OptionButton alcoholunknown 
               BackColor       =   &H00FFFFFF&
               Caption         =   "UNK"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   195
               Index           =   1
               Left            =   1000
               TabIndex        =   144
               Top             =   0
               Value           =   -1  'True
               Width           =   555
            End
         End
         Begin VB.CheckBox SUMMONS 
            BackColor       =   &H00FFFFFF&
            Caption         =   "SUMMONS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   350
            MaskColor       =   &H00808000&
            TabIndex        =   125
            Top             =   11040
            Width           =   910
         End
         Begin VB.CheckBox JAIL 
            BackColor       =   &H00FFFFFF&
            Caption         =   "JAIL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   350
            MaskColor       =   &H00808000&
            TabIndex        =   124
            Top             =   10680
            Width           =   735
         End
         Begin VB.CheckBox ARREST 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ARREST"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   350
            MaskColor       =   &H00808000&
            TabIndex        =   123
            Top             =   10320
            Width           =   850
         End
         Begin VB.CheckBox WARRANT 
            BackColor       =   &H00FFFFFF&
            Caption         =   "WARRANT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   350
            MaskColor       =   &H00808000&
            TabIndex        =   122
            Top             =   9960
            Width           =   905
         End
         Begin VB.CheckBox WANTED 
            BackColor       =   &H00FFFFFF&
            Caption         =   "WANTED"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   350
            MaskColor       =   &H00808000&
            TabIndex        =   121
            Top             =   9600
            Width           =   850
         End
         Begin VB.CheckBox RUNAWAY 
            BackColor       =   &H00FFFFFF&
            Caption         =   "RUNAWAY"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   350
            MaskColor       =   &H00808000&
            TabIndex        =   120
            Top             =   9240
            Width           =   905
         End
         Begin VB.CheckBox SUSPECT 
            BackColor       =   &H00FFFFFF&
            Caption         =   "SUSPECT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            MaskColor       =   &H00808000&
            TabIndex        =   119
            Top             =   8880
            Width           =   850
         End
         Begin VB.CheckBox ALONE 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ALONE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Left            =   5865
            TabIndex        =   117
            Top             =   8475
            Width           =   720
         End
         Begin VB.CheckBox ASSISTED 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ASSISTED"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Left            =   6720
            TabIndex        =   118
            Top             =   8480
            Width           =   975
         End
         Begin VB.CheckBox TODOTHER 
            BackColor       =   &H00FFFFFF&
            Caption         =   "OTHER"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Left            =   4800
            TabIndex        =   116
            Top             =   8480
            Width           =   765
         End
         Begin VB.CheckBox ONEMANVEHICLE 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ONE-MAN VEH"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Left            =   1680
            TabIndex        =   114
            Top             =   8480
            Width           =   1215
         End
         Begin VB.CheckBox TWOMANVEHICLE 
            BackColor       =   &H00FFFFFF&
            Caption         =   "TWO-MAN VEH"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Left            =   360
            TabIndex        =   113
            Top             =   8480
            Width           =   1200
         End
         Begin VB.CheckBox DETECTIVE 
            BackColor       =   &H00FFFFFF&
            Caption         =   "DETECTIVE/SPL ASMNT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Left            =   3000
            TabIndex        =   115
            Top             =   8480
            Width           =   1725
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   10200
            TabIndex        =   358
            Top             =   7860
            Width           =   1050
            Begin VB.OptionButton NONVISIBLEINJURYYES 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Yes"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   0
               TabIndex        =   103
               Top             =   0
               Width           =   480
            End
            Begin VB.OptionButton NONVISIBLEINJURYNO 
               BackColor       =   &H00FFFFFF&
               Caption         =   "No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   550
               TabIndex        =   104
               Top             =   0
               Value           =   -1  'True
               Width           =   495
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   1680
            TabIndex        =   357
            Top             =   7860
            Width           =   1215
            Begin VB.OptionButton VISIBLEINJURYNO 
               BackColor       =   &H00FFFFFF&
               Caption         =   "No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   600
               TabIndex        =   101
               Top             =   0
               Value           =   -1  'True
               Width           =   510
            End
            Begin VB.OptionButton VISIBLEINJURYYES 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Yes"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   0
               TabIndex        =   100
               Top             =   0
               Width           =   600
            End
         End
         Begin VB.TextBox peculiarities 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   1320
            TabIndex        =   136
            Top             =   9960
            Width           =   10065
         End
         Begin VB.TextBox eyes 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   10800
            TabIndex        =   135
            Top             =   9405
            Width           =   780
         End
         Begin VB.TextBox hair 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   10020
            TabIndex        =   134
            Top             =   9405
            Width           =   780
         End
         Begin VB.TextBox weight 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   9240
            TabIndex        =   133
            Top             =   9405
            Width           =   780
         End
         Begin VB.TextBox ht 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   8400
            TabIndex        =   132
            Top             =   9405
            Width           =   780
         End
         Begin VB.TextBox BIRTHDATE 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   7450
            TabIndex        =   131
            Top             =   9405
            Width           =   900
         End
         Begin VB.TextBox peculiarities 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   3480
            TabIndex        =   94
            Top             =   6840
            Width           =   8025
         End
         Begin VB.TextBox eyes 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   2700
            TabIndex        =   93
            Top             =   6840
            Width           =   785
         End
         Begin VB.TextBox hair 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   1920
            TabIndex        =   92
            Top             =   6840
            Width           =   785
         End
         Begin VB.TextBox weight 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   1130
            TabIndex        =   91
            Top             =   6840
            Width           =   785
         End
         Begin VB.TextBox ht 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   360
            TabIndex        =   90
            Top             =   6840
            Width           =   785
         End
         Begin VB.TextBox WORKDAYPHONE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   1
            Left            =   9120
            TabIndex        =   87
            Top             =   6300
            Width           =   1020
         End
         Begin VB.TextBox WORKNIGHTPHONE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   10320
            TabIndex        =   89
            Top             =   6300
            Width           =   1020
         End
         Begin VB.TextBox HOMENIGHTPHONE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   1
            Left            =   10320
            TabIndex        =   88
            Top             =   6000
            Width           =   1020
         End
         Begin VB.TextBox HOMEDAYPHONE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   1
            Left            =   9120
            TabIndex        =   86
            Top             =   6000
            Width           =   1020
         End
         Begin VB.TextBox LOCATIONNUMBER 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   10350
            TabIndex        =   141
            Top             =   10560
            Width           =   1080
         End
         Begin VB.TextBox LOCATIONNUMBER 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   10440
            TabIndex        =   99
            Top             =   7360
            Width           =   1080
         End
         Begin VB.TextBox LOCATIONNUMBER 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   10440
            TabIndex        =   77
            Top             =   5280
            Width           =   1080
         End
         Begin VB.TextBox WORKNIGHTPHONE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   0
            Left            =   10350
            TabIndex        =   72
            Top             =   4700
            Width           =   1020
         End
         Begin VB.TextBox HOMENIGHTPHONE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   0
            Left            =   10350
            TabIndex        =   71
            Top             =   4380
            Width           =   1020
         End
         Begin VB.TextBox WORKDAYPHONE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   0
            Left            =   9060
            TabIndex        =   70
            Top             =   4700
            Width           =   1020
         End
         Begin VB.TextBox HOMEDAYPHONE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   0
            Left            =   9060
            TabIndex        =   69
            Top             =   4380
            Width           =   1020
         End
         Begin VB.TextBox DEPARTINGTIME 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   9400
            TabIndex        =   60
            Top             =   3800
            Width           =   1260
         End
         Begin VB.TextBox TIMEARRIVED 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   8160
            TabIndex        =   59
            Top             =   3800
            Width           =   1260
         End
         Begin VB.TextBox DISPATCHTIME 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   6880
            TabIndex        =   58
            Top             =   3800
            Width           =   1260
         End
         Begin VB.TextBox dispatchdate 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   5640
            TabIndex        =   57
            Top             =   3800
            Width           =   1260
         End
         Begin VB.TextBox numvehicle 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   11
            Left            =   9960
            MaxLength       =   2
            TabIndex        =   262
            Top             =   15600
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox numvehicle 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   10
            Left            =   8520
            MaxLength       =   2
            TabIndex        =   258
            Top             =   15600
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox numvehicle 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   9
            Left            =   6840
            MaxLength       =   2
            TabIndex        =   255
            Top             =   15600
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox numvehicle 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   8
            Left            =   5400
            MaxLength       =   2
            TabIndex        =   252
            Top             =   15600
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox numvehicle 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   7
            Left            =   3720
            MaxLength       =   2
            TabIndex        =   249
            Top             =   15600
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox numvehicle 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   6
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   246
            Top             =   15600
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox numvehicle 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   5
            Left            =   9720
            MaxLength       =   2
            TabIndex        =   260
            Top             =   14640
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox numvehicle 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   4
            Left            =   8160
            MaxLength       =   2
            TabIndex        =   257
            Top             =   14640
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox numvehicle 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   3
            Left            =   6720
            MaxLength       =   2
            TabIndex        =   254
            Top             =   14640
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox numvehicle 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   2
            Left            =   5160
            MaxLength       =   2
            TabIndex        =   251
            Top             =   14640
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox numvehicle 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   1
            Left            =   3480
            MaxLength       =   2
            TabIndex        =   248
            Top             =   14640
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox numvehicle 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   0
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   245
            Top             =   14640
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   36
            Left            =   1320
            TabIndex        =   166
            Top             =   16560
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   37
            Left            =   2880
            TabIndex        =   175
            Top             =   16560
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   38
            Left            =   4440
            TabIndex        =   184
            Top             =   16560
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   39
            Left            =   6000
            TabIndex        =   193
            Top             =   16560
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   40
            Left            =   7560
            TabIndex        =   202
            Top             =   16560
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   41
            Left            =   9120
            TabIndex        =   211
            Top             =   16560
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   30
            Left            =   1320
            TabIndex        =   165
            Top             =   16260
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   31
            Left            =   2880
            TabIndex        =   174
            Top             =   16260
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   32
            Left            =   4440
            TabIndex        =   183
            Top             =   16260
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   33
            Left            =   6000
            TabIndex        =   192
            Top             =   16260
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   34
            Left            =   7560
            TabIndex        =   201
            Top             =   16260
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   35
            Left            =   9120
            TabIndex        =   210
            Top             =   16260
            Width           =   1440
         End
         Begin VB.TextBox DATERECOVERED 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   5
            Left            =   9600
            TabIndex        =   261
            Top             =   15600
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.TextBox DATERECOVERED 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   4
            Left            =   8640
            TabIndex        =   259
            Top             =   15600
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.TextBox DATERECOVERED 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   3
            Left            =   6840
            TabIndex        =   256
            Top             =   15600
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.TextBox DATERECOVERED 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   2
            Left            =   5520
            TabIndex        =   253
            Top             =   15600
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.TextBox DATERECOVERED 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   1
            Left            =   3840
            TabIndex        =   250
            Top             =   15600
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.TextBox DATERECOVERED 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   0
            Left            =   2160
            TabIndex        =   247
            Top             =   15600
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.CommandButton setrel 
            BackColor       =   &H00808000&
            Caption         =   "Set Relationship"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   1
            Left            =   3740
            Style           =   1  'Graphical
            TabIndex        =   80
            Top             =   5950
            Width           =   985
         End
         Begin VB.CommandButton setrel 
            Caption         =   "Set Relationship"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   0
            Left            =   3720
            TabIndex        =   63
            Top             =   4400
            Width           =   995
         End
         Begin VB.ListBox ethnicity 
            BackColor       =   &H00808000&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   2
            Left            =   6500
            TabIndex        =   130
            Top             =   9270
            Width           =   930
         End
         Begin VB.ListBox race 
            BackColor       =   &H00808000&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   2
            Left            =   4030
            TabIndex        =   127
            Top             =   9270
            Width           =   900
         End
         Begin VB.ListBox sex 
            BackColor       =   &H00808000&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   2
            Left            =   4980
            TabIndex        =   128
            Top             =   9270
            Width           =   900
         End
         Begin VB.ListBox ethnicity 
            BackColor       =   &H00808000&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   1
            Left            =   8160
            TabIndex        =   85
            Top             =   6000
            Width           =   930
         End
         Begin VB.ListBox race 
            BackColor       =   &H00808000&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   1
            Left            =   5700
            TabIndex        =   82
            Top             =   6000
            Width           =   855
         End
         Begin VB.ListBox sex 
            BackColor       =   &H00808000&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   1
            Left            =   6630
            TabIndex        =   83
            Top             =   6000
            Width           =   855
         End
         Begin VB.ListBox resident 
            BackColor       =   &H00808000&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   1
            Left            =   4750
            TabIndex        =   81
            Top             =   6000
            Width           =   900
         End
         Begin VB.ListBox resident 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   420
            Index           =   0
            Left            =   4680
            TabIndex        =   64
            Top             =   4440
            Width           =   960
         End
         Begin VB.ListBox sex 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   420
            Index           =   0
            Left            =   6630
            TabIndex        =   66
            Top             =   4440
            Width           =   880
         End
         Begin VB.ListBox race 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   420
            Index           =   0
            Left            =   5640
            TabIndex        =   65
            Top             =   4440
            Width           =   975
         End
         Begin VB.ListBox ethnicity 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   420
            Index           =   0
            Left            =   8160
            TabIndex        =   68
            Top             =   4440
            Width           =   900
         End
         Begin VB.CommandButton Command9 
            BackColor       =   &H00808000&
            Caption         =   "Drug"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   148
            TabStop         =   0   'False
            Top             =   11280
            Width           =   495
         End
         Begin VB.CommandButton Command7 
            BackColor       =   &H00808000&
            Caption         =   "Drug"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8160
            Style           =   1  'Graphical
            TabIndex        =   111
            TabStop         =   0   'False
            Top             =   8160
            Width           =   495
         End
         Begin VB.CheckBox computerequipment 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Computer"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   0
            Left            =   10550
            TabIndex        =   112
            TabStop         =   0   'False
            Top             =   8160
            Width           =   1000
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   0
            Left            =   5520
            TabIndex        =   319
            Top             =   8160
            Width           =   2055
            Begin VB.OptionButton drugsno 
               BackColor       =   &H00FFFFFF&
               Caption         =   "No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   300
               Index           =   0
               Left            =   600
               TabIndex        =   109
               Top             =   0
               Width           =   510
            End
            Begin VB.OptionButton drugsyes 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Yes"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   300
               Index           =   0
               Left            =   0
               TabIndex        =   108
               Top             =   0
               Width           =   600
            End
            Begin VB.OptionButton drugsunknown 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Unknown"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   300
               Index           =   0
               Left            =   1200
               TabIndex        =   110
               Top             =   0
               Value           =   -1  'True
               Width           =   885
            End
         End
         Begin VB.Frame alcoholframe 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   0
            Left            =   2160
            TabIndex        =   318
            Top             =   8180
            Width           =   2655
            Begin VB.OptionButton alcoholunknown 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Unknown"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   195
               Index           =   0
               Left            =   1440
               TabIndex        =   107
               Top             =   0
               Value           =   -1  'True
               Width           =   915
            End
            Begin VB.OptionButton alcoholyes 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Yes"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   105
               Top             =   0
               Width           =   600
            End
            Begin VB.OptionButton alcoholno 
               BackColor       =   &H00FFFFFF&
               Caption         =   "No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   195
               Index           =   0
               Left            =   840
               TabIndex        =   106
               Top             =   0
               Width           =   510
            End
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H00808000&
            Caption         =   "Bias"
            Height          =   255
            Left            =   50
            Style           =   1  'Graphical
            TabIndex        =   213
            TabStop         =   0   'False
            Top             =   16920
            Width           =   615
         End
         Begin VB.TextBox APPROVINGOFFICERDATE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   9960
            TabIndex        =   238
            Top             =   18120
            Width           =   960
         End
         Begin VB.ComboBox followupofficer 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   7800
            Sorted          =   -1  'True
            TabIndex        =   242
            Top             =   18430
            Width           =   2145
         End
         Begin VB.ComboBox approvingofficer 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   6000
            Sorted          =   -1  'True
            TabIndex        =   237
            Top             =   18120
            Width           =   3945
         End
         Begin VB.Frame Frame23 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   6840
            TabIndex        =   316
            Top             =   18480
            Width           =   975
            Begin VB.OptionButton followupyes 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Yes"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   195
               Left            =   0
               TabIndex        =   240
               Top             =   0
               Width           =   480
            End
            Begin VB.OptionButton followupno 
               BackColor       =   &H00FFFFFF&
               Caption         =   "No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   195
               Left            =   480
               TabIndex        =   241
               Top             =   0
               Width           =   510
            End
         End
         Begin VB.TextBox approvingofficeRunit 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   10920
            TabIndex        =   239
            Top             =   18120
            Width           =   660
         End
         Begin VB.TextBox EXCEPTIONALCLEARANCEDATE 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   10080
            TabIndex        =   315
            Top             =   17490
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.TextBox REPORTINGOFFICERDATE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   4430
            TabIndex        =   235
            Top             =   18430
            Width           =   960
         End
         Begin VB.ComboBox reportingofficer 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   1
            Left            =   360
            Sorted          =   -1  'True
            TabIndex        =   234
            Top             =   18430
            Width           =   3945
         End
         Begin VB.TextBox reportingofficeRunit 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   5360
            TabIndex        =   236
            Top             =   18430
            Width           =   600
         End
         Begin VB.TextBox REPORTINGOFFICERDATE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   4410
            TabIndex        =   232
            Top             =   18120
            Width           =   960
         End
         Begin VB.ComboBox reportingofficer 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   0
            Left            =   360
            Sorted          =   -1  'True
            TabIndex        =   231
            Top             =   18120
            Width           =   3945
         End
         Begin VB.TextBox reportingofficeRunit 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   5360
            TabIndex        =   233
            Top             =   18120
            Width           =   600
         End
         Begin VB.CheckBox arrestedunder18 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Arrested Under 18"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Left            =   7200
            TabIndex        =   222
            Top             =   16920
            Width           =   1755
         End
         Begin VB.CheckBox arrested18andover 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Arrested 18 and Over"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Left            =   7200
            TabIndex        =   223
            Top             =   17160
            Width           =   1905
         End
         Begin VB.CheckBox exclear18andover 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Ex-Clear 18 and Over"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Left            =   9600
            TabIndex        =   225
            Top             =   17175
            Width           =   1905
         End
         Begin VB.CheckBox exclearunder18 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Ex-Clear Under 18"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Left            =   9600
            TabIndex        =   224
            Top             =   16920
            Width           =   1755
         End
         Begin VB.Frame Frame19 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   720
            TabIndex        =   314
            Top             =   17160
            Width           =   1335
            Begin VB.OptionButton subjectidentifiedyes 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Yes"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   0
               TabIndex        =   214
               Top             =   0
               Width           =   600
            End
            Begin VB.OptionButton subjectidentifiedno 
               BackColor       =   &H00FFFFFF&
               Caption         =   "No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   720
               TabIndex        =   215
               Top             =   0
               Width           =   510
            End
         End
         Begin VB.Frame Frame17 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   3000
            TabIndex        =   313
            Top             =   17160
            Width           =   1335
            Begin VB.OptionButton subjectlocatedyes 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Yes"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   0
               TabIndex        =   217
               Top             =   0
               Width           =   600
            End
            Begin VB.OptionButton subjectlocatedno 
               BackColor       =   &H00FFFFFF&
               Caption         =   "No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   720
               TabIndex        =   218
               Top             =   0
               Width           =   510
            End
         End
         Begin VB.Frame Frame18 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   2640
            TabIndex        =   311
            Top             =   17520
            Width           =   7440
            Begin VB.OptionButton na 
               BackColor       =   &H00808000&
               Caption         =   "N/A"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   7440
               TabIndex        =   312
               Top             =   120
               Value           =   -1  'True
               Width           =   705
            End
            Begin VB.OptionButton offenderdeath 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Offender Death"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   195
               Left            =   0
               TabIndex        =   226
               Top             =   0
               Width           =   1215
            End
            Begin VB.OptionButton noprosecution 
               BackColor       =   &H00FFFFFF&
               Caption         =   "No Prosecution"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   195
               Left            =   1320
               TabIndex        =   227
               Top             =   0
               Width           =   1305
            End
            Begin VB.OptionButton extraditiondenied 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Extradition Denied"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   195
               Left            =   2590
               TabIndex        =   228
               Top             =   0
               Width           =   1320
            End
            Begin VB.OptionButton victimdeclinescooperation 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Victim Declines Cooperation"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   195
               Left            =   3960
               TabIndex        =   229
               Top             =   0
               Width           =   1920
            End
            Begin VB.OptionButton juvenilenocustody 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Juvenile - No Custody"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   195
               Left            =   5880
               TabIndex        =   230
               Top             =   0
               Width           =   1665
            End
         End
         Begin VB.CheckBox unfounded 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Unfounded"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Left            =   4950
            TabIndex        =   221
            Top             =   17200
            Width           =   1125
         End
         Begin VB.CheckBox admclosed 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Adm. Closed"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Left            =   5620
            TabIndex        =   220
            Top             =   16950
            Width           =   1185
         End
         Begin VB.CheckBox active 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Active"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Left            =   4950
            TabIndex        =   219
            Top             =   16950
            Width           =   795
         End
         Begin VB.TextBox TIMEOFOFFENSE 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   4440
            TabIndex        =   56
            Top             =   3800
            Width           =   1155
         End
         Begin VB.TextBox incidentdate 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   3240
            TabIndex        =   55
            Top             =   3800
            Width           =   1140
         End
         Begin VB.TextBox TIMEOFOFFENSE 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   1600
            TabIndex        =   54
            Top             =   3800
            Width           =   1155
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   29
            Left            =   9120
            TabIndex        =   209
            Top             =   15960
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   28
            Left            =   7560
            TabIndex        =   200
            Top             =   15960
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   27
            Left            =   6000
            TabIndex        =   191
            Top             =   15960
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   26
            Left            =   4455
            TabIndex        =   182
            Top             =   15960
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   25
            Left            =   2880
            TabIndex        =   173
            Top             =   15945
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   24
            Left            =   1320
            TabIndex        =   164
            Top             =   15960
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   23
            Left            =   9120
            TabIndex        =   208
            Top             =   15600
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   22
            Left            =   7560
            TabIndex        =   199
            Top             =   15600
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   21
            Left            =   6000
            TabIndex        =   190
            Top             =   15600
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   20
            Left            =   4440
            TabIndex        =   181
            Top             =   15600
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   19
            Left            =   2880
            TabIndex        =   172
            Top             =   15600
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   18
            Left            =   1320
            TabIndex        =   163
            Top             =   15600
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   17
            Left            =   9120
            TabIndex        =   207
            Top             =   15310
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   16
            Left            =   7560
            TabIndex        =   198
            Top             =   15310
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   15
            Left            =   6000
            TabIndex        =   189
            Top             =   15310
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   14
            Left            =   4440
            TabIndex        =   180
            Top             =   15310
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   13
            Left            =   2880
            TabIndex        =   171
            Top             =   15310
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   12
            Left            =   1320
            TabIndex        =   162
            Top             =   15310
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   11
            Left            =   9120
            TabIndex        =   206
            Top             =   15000
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   10
            Left            =   7560
            TabIndex        =   197
            Top             =   15000
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   9
            Left            =   6000
            TabIndex        =   188
            Top             =   15000
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   8
            Left            =   4440
            TabIndex        =   179
            Top             =   15000
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   7
            Left            =   2880
            TabIndex        =   170
            Top             =   15000
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   6
            Left            =   1320
            TabIndex        =   161
            Top             =   15000
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   5
            Left            =   9120
            TabIndex        =   205
            Top             =   14700
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   4
            Left            =   7560
            TabIndex        =   196
            Top             =   14700
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   3
            Left            =   6000
            TabIndex        =   187
            Top             =   14700
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   4440
            TabIndex        =   178
            Top             =   14700
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   2880
            TabIndex        =   169
            Top             =   14700
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   1320
            TabIndex        =   160
            Top             =   14700
            Width           =   1440
         End
         Begin VB.TextBox description 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   5
            Left            =   9120
            MaxLength       =   30
            MultiLine       =   -1  'True
            TabIndex        =   204
            Top             =   14400
            Width           =   1440
         End
         Begin VB.TextBox description 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   4
            Left            =   7560
            MaxLength       =   30
            MultiLine       =   -1  'True
            TabIndex        =   195
            Top             =   14400
            Width           =   1440
         End
         Begin VB.TextBox description 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   3
            Left            =   6000
            MaxLength       =   30
            MultiLine       =   -1  'True
            TabIndex        =   186
            Top             =   14400
            Width           =   1440
         End
         Begin VB.TextBox description 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   4440
            MaxLength       =   30
            MultiLine       =   -1  'True
            TabIndex        =   177
            Top             =   14400
            Width           =   1440
         End
         Begin VB.TextBox description 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   2880
            MaxLength       =   30
            MultiLine       =   -1  'True
            TabIndex        =   168
            Top             =   14400
            Width           =   1440
         End
         Begin VB.TextBox description 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   1320
            MaxLength       =   30
            MultiLine       =   -1  'True
            TabIndex        =   159
            Top             =   14400
            Width           =   1440
         End
         Begin VB.TextBox incidentlocation 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   360
            MultiLine       =   -1  'True
            TabIndex        =   50
            Top             =   3120
            Width           =   6435
         End
         Begin VB.TextBox incidentzipcode 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   6960
            TabIndex        =   51
            Top             =   3120
            Width           =   1380
         End
         Begin VB.TextBox ELOCATIONNUMBER 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   10680
            TabIndex        =   61
            Top             =   3800
            Width           =   840
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00808000&
            Caption         =   "Property Info"
            Height          =   255
            Index           =   5
            Left            =   9120
            Style           =   1  'Graphical
            TabIndex        =   203
            Top             =   14160
            Width           =   1100
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00808000&
            Caption         =   "Property Info"
            Height          =   255
            Index           =   4
            Left            =   7560
            Style           =   1  'Graphical
            TabIndex        =   194
            Top             =   14160
            Width           =   1100
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00808000&
            Caption         =   "Property Info"
            Height          =   255
            Index           =   3
            Left            =   6000
            Style           =   1  'Graphical
            TabIndex        =   185
            Top             =   14160
            Width           =   1100
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00808000&
            Caption         =   "Property Info"
            Height          =   255
            Index           =   2
            Left            =   4440
            Style           =   1  'Graphical
            TabIndex        =   176
            Top             =   14160
            Width           =   1100
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00808000&
            Caption         =   "Property Info"
            Height          =   255
            Index           =   1
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   167
            Top             =   14160
            Width           =   1100
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00808000&
            Caption         =   "xUCR"
            Height          =   255
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   5760
            Width           =   650
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00808000&
            Caption         =   "Property Info"
            Height          =   255
            Index           =   0
            Left            =   1350
            Style           =   1  'Graphical
            TabIndex        =   158
            Top             =   14160
            Width           =   1100
         End
         Begin VB.OptionButton business 
            BackColor       =   &H00808000&
            Caption         =   "Business"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   9800
            TabIndex        =   41
            Top             =   1100
            Width           =   1665
         End
         Begin VB.OptionButton financialinstitution 
            BackColor       =   &H00808000&
            Caption         =   "Financial Inst."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   9800
            TabIndex        =   42
            Top             =   1300
            Width           =   1665
         End
         Begin VB.OptionButton government 
            BackColor       =   &H00808000&
            Caption         =   "Government"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   9800
            TabIndex        =   43
            Top             =   1500
            Width           =   1665
         End
         Begin VB.OptionButton religiousorganization 
            BackColor       =   &H00808000&
            Caption         =   "Religious Org."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   9800
            TabIndex        =   44
            Top             =   1700
            Width           =   1665
         End
         Begin VB.OptionButton societypublic 
            BackColor       =   &H00808000&
            Caption         =   "Society/Public"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   9800
            TabIndex        =   45
            Top             =   1900
            Width           =   1665
         End
         Begin VB.OptionButton other 
            BackColor       =   &H00808000&
            Caption         =   "Other"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   9800
            TabIndex        =   46
            Top             =   2100
            Width           =   1665
         End
         Begin VB.OptionButton unknown 
            BackColor       =   &H00808000&
            Caption         =   "Unknown"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   9800
            TabIndex        =   47
            Top             =   2300
            Width           =   1665
         End
         Begin VB.OptionButton policeofficer 
            BackColor       =   &H00808000&
            Caption         =   "Police Officer"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   9800
            TabIndex        =   48
            Top             =   2500
            Width           =   1665
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame12"
            Height          =   255
            Index           =   2
            Left            =   4440
            TabIndex        =   279
            Top             =   2550
            Width           =   1145
            Begin VB.OptionButton completedn 
               BackColor       =   &H00FFFFFF&
               Caption         =   "NO"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   255
               Index           =   2
               Left            =   600
               TabIndex        =   35
               Top             =   0
               Width           =   495
            End
            Begin VB.OptionButton completedy 
               BackColor       =   &H00FFFFFF&
               Caption         =   "YES"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   255
               Index           =   2
               Left            =   0
               TabIndex        =   34
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame12"
            Height          =   255
            Index           =   1
            Left            =   4440
            TabIndex        =   278
            Top             =   2250
            Width           =   1145
            Begin VB.OptionButton completedy 
               BackColor       =   &H00FFFFFF&
               Caption         =   "YES"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   22
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton completedn 
               BackColor       =   &H00FFFFFF&
               Caption         =   "NO"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   255
               Index           =   1
               Left            =   600
               TabIndex        =   23
               Top             =   0
               Width           =   495
            End
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame12"
            Height          =   255
            Index           =   0
            Left            =   4440
            TabIndex        =   9
            Top             =   1920
            Width           =   1145
            Begin VB.OptionButton completedn 
               BackColor       =   &H00FFFFFF&
               Caption         =   "NO"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   255
               Index           =   0
               Left            =   600
               TabIndex        =   11
               Top             =   0
               Width           =   495
            End
            Begin VB.OptionButton completedy 
               BackColor       =   &H00FFFFFF&
               Caption         =   "YES"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   10
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.CommandButton ao 
            Caption         =   "Additional Offenses"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Left            =   3000
            TabIndex        =   271
            Top             =   1600
            Width           =   1335
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Find"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   325
            Index           =   2
            Left            =   0
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   2520
            Width           =   400
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Find"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   325
            Index           =   1
            Left            =   0
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   2180
            Width           =   400
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00808000&
            Caption         =   "Find"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   325
            Index           =   0
            Left            =   0
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   1800
            Width           =   400
         End
         Begin VB.ComboBox city 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   2
            Left            =   6360
            Sorted          =   -1  'True
            TabIndex        =   138
            Top             =   10560
            Width           =   1740
         End
         Begin VB.TextBox address 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   1320
            TabIndex        =   137
            Top             =   10560
            Width           =   4815
         End
         Begin VB.ComboBox state 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   2
            Left            =   8200
            Sorted          =   -1  'True
            TabIndex        =   139
            Top             =   10560
            Width           =   765
         End
         Begin VB.TextBox zipcode 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   9120
            TabIndex        =   140
            Top             =   10560
            Width           =   1140
         End
         Begin VB.ComboBox city 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   1
            Left            =   5640
            Sorted          =   -1  'True
            TabIndex        =   96
            Top             =   7360
            Width           =   2460
         End
         Begin VB.TextBox address 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   360
            TabIndex        =   95
            Top             =   7360
            Width           =   5175
         End
         Begin VB.ComboBox state 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   1
            Left            =   8160
            Sorted          =   -1  'True
            TabIndex        =   97
            Top             =   7360
            Width           =   885
         End
         Begin VB.TextBox zipcode 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   1
            Left            =   9120
            TabIndex        =   98
            Top             =   7360
            Width           =   1140
         End
         Begin VB.ComboBox city 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   0
            Left            =   5760
            Sorted          =   -1  'True
            TabIndex        =   74
            Top             =   5280
            Width           =   2340
         End
         Begin VB.TextBox address 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   360
            TabIndex        =   73
            Top             =   5280
            Width           =   5175
         End
         Begin VB.ComboBox state 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   0
            Left            =   8160
            Sorted          =   -1  'True
            TabIndex        =   75
            Top             =   5280
            Width           =   885
         End
         Begin VB.TextBox zipcode 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   0
            Left            =   9120
            TabIndex        =   76
            Top             =   5280
            Width           =   1140
         End
         Begin VB.ComboBox vsname 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   2
            Left            =   1305
            Sorted          =   -1  'True
            TabIndex        =   126
            Top             =   9375
            Width           =   2685
         End
         Begin VB.ComboBox vsname 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   1
            Left            =   375
            Sorted          =   -1  'True
            TabIndex        =   79
            Top             =   6120
            Width           =   3285
         End
         Begin VB.ComboBox vsname 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   0
            Left            =   360
            Sorted          =   -1  'True
            TabIndex        =   62
            Top             =   4560
            Width           =   3285
         End
         Begin VB.TextBox incidentdate 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   360
            TabIndex        =   53
            Top             =   3800
            Width           =   1140
         End
         Begin MSComctlLib.ListView weapontype 
            Height          =   495
            Left            =   8460
            TabIndex        =   52
            Top             =   3000
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   873
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   16777215
            BackColor       =   8421376
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   4410
            EndProperty
         End
         Begin MSComctlLib.ListView injury 
            Height          =   495
            Left            =   3840
            TabIndex        =   102
            Top             =   7650
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   873
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            ForeColor       =   16777215
            BackColor       =   8421376
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   4939
            EndProperty
         End
         Begin VB.CommandButton subcode 
            BackColor       =   &H00808000&
            Caption         =   "Sub"
            Height          =   255
            Index           =   0
            Left            =   4030
            MaskColor       =   &H00808000&
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   1930
            Width           =   400
         End
         Begin VB.CommandButton subcode 
            BackColor       =   &H00808000&
            Caption         =   "Sub"
            Height          =   255
            Index           =   1
            Left            =   4030
            MaskColor       =   &H00808000&
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   2250
            Width           =   400
         End
         Begin VB.CommandButton subcode 
            BackColor       =   &H00808000&
            Caption         =   "Sub"
            Height          =   255
            Index           =   2
            Left            =   4030
            MaskColor       =   &H00808000&
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   2550
            Width           =   400
         End
         Begin VB.OptionButton individual 
            BackColor       =   &H00808000&
            Caption         =   "Individual"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   9800
            TabIndex        =   40
            Top             =   900
            Value           =   -1  'True
            Width           =   1665
         End
         Begin VB.CheckBox onpaper 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Incident Report Previously Submitted on Paper"
            ForeColor       =   &H00808000&
            Height          =   615
            Left            =   6840
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   120
            Width           =   2160
         End
         Begin VB.CheckBox locali 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Local Information Only"
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   6840
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   840
            Width           =   2040
         End
         Begin VB.TextBox entered 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   8880
            TabIndex        =   39
            Top             =   2575
            Width           =   750
         End
         Begin VB.TextBox entered 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   8880
            TabIndex        =   27
            Top             =   2250
            Width           =   750
         End
         Begin VB.TextBox entered 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   8880
            TabIndex        =   15
            Top             =   1920
            Width           =   750
         End
         Begin VB.CheckBox FORCEDENTRYY 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Check1"
            Height          =   255
            Index           =   0
            Left            =   5640
            TabIndex        =   12
            Top             =   1920
            Width           =   150
         End
         Begin VB.CheckBox FORCEDENTRYY 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Check1"
            Height          =   255
            Index           =   1
            Left            =   5640
            TabIndex        =   24
            Top             =   2225
            Width           =   150
         End
         Begin VB.CheckBox FORCEDENTRYY 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Check1"
            Height          =   255
            Index           =   2
            Left            =   5640
            TabIndex        =   36
            Top             =   2550
            Width           =   150
         End
         Begin VB.CheckBox FORCEDENTRYN 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Check1"
            Height          =   255
            Index           =   0
            Left            =   6480
            TabIndex        =   13
            Top             =   1920
            Width           =   150
         End
         Begin VB.CheckBox FORCEDENTRYN 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Check1"
            Height          =   255
            Index           =   1
            Left            =   6480
            TabIndex        =   25
            Top             =   2225
            Width           =   150
         End
         Begin VB.CheckBox FORCEDENTRYN 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Check1"
            Height          =   255
            Index           =   2
            Left            =   6480
            TabIndex        =   37
            Top             =   2550
            Width           =   150
         End
         Begin VB.CommandButton Command15 
            Caption         =   "Check Spelling"
            Height          =   855
            Left            =   10560
            TabIndex        =   438
            Top             =   11640
            Width           =   975
         End
         Begin VB.CommandButton Command19 
            Caption         =   "CaseSetup"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   9180
            TabIndex        =   445
            TabStop         =   0   'False
            Top             =   600
            Width           =   495
         End
         Begin VB.ListBox pickoffense 
            Height          =   255
            Index           =   1
            Left            =   405
            TabIndex        =   17
            Top             =   2235
            Width           =   2295
         End
         Begin VB.ListBox pickoffense 
            Height          =   255
            Index           =   2
            Left            =   405
            TabIndex        =   30
            Top             =   2550
            Width           =   2295
         End
         Begin VB.ListBox pickoffense 
            Height          =   255
            Index           =   0
            Left            =   405
            TabIndex        =   4
            Top             =   1935
            Width           =   2295
         End
         Begin VB.Label pgof 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   120
            TabIndex        =   585
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "YES      NO"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Left            =   7020
            TabIndex        =   437
            Top             =   -29055
            Width           =   615
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "YES      NO"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Left            =   6930
            TabIndex        =   436
            Top             =   -26880
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "YES      NO"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   6840
            TabIndex        =   435
            Top             =   -24720
            Width           =   615
         End
         Begin VB.Label holdoff 
            BackStyle       =   0  'Transparent
            Height          =   135
            Left            =   4860
            TabIndex        =   433
            Top             =   -24060
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label orinumber 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   930
            TabIndex        =   369
            Top             =   -21330
            Width           =   1695
         End
         Begin VB.Label dtoffense 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   9240
            TabIndex        =   368
            Top             =   -8040
            Width           =   1440
         End
         Begin VB.Label TVUNKNOWN 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   11310
            TabIndex        =   367
            Top             =   -150
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label TVCOUNTERFEIT 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   11220
            TabIndex        =   366
            Top             =   2040
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label TVSEIZED 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   10755
            TabIndex        =   365
            Top             =   4170
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label TVRECOVERED 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   11040
            TabIndex        =   364
            Top             =   6360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label TVBURNED 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   10860
            TabIndex        =   363
            Top             =   8475
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label TVDAMAGED 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   10860
            TabIndex        =   362
            Top             =   10740
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label TVSTOLEN 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   10710
            TabIndex        =   212
            Top             =   12975
            Visible         =   0   'False
            Width           =   855
         End
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   7215
      Left            =   11640
      TabIndex        =   268
      Top             =   705
      Width           =   250
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   267
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   1111
      ButtonWidth     =   1111
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "TmpSv"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clear"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Send"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "TmpLst"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "NxtPg"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Book"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "MthRpt"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Email"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Setup"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Trans"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SrvCall"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "BadCk"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "NCrim"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Flag"
            Object.ToolTipText     =   "FlagSend"
            ImageIndex      =   20
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ninciden.frx":5CCD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ninciden.frx":5D12A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ninciden.frx":5D57E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ninciden.frx":5D9D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ninciden.frx":5DE26
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ninciden.frx":5E27A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ninciden.frx":5E6CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ninciden.frx":5EB22
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ninciden.frx":5EF76
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ninciden.frx":5F3CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ninciden.frx":5F81E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ninciden.frx":5FC72
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ninciden.frx":600C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ninciden.frx":6051A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ninciden.frx":62CCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ninciden.frx":63122
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ninciden.frx":63576
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ninciden.frx":639CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ninciden.frx":63E1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ninciden.frx":64272
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame dvipframe 
      BackColor       =   &H00800000&
      Height          =   1440
      Left            =   1920
      TabIndex        =   586
      Top             =   3120
      Width           =   8775
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Data Verification In Progress . . ."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   435
         Left            =   1080
         TabIndex        =   587
         Top             =   600
         Width           =   6660
      End
   End
End
Attribute VB_Name = "incident"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const NUM_OF_SUGS = "Number of Spelling Suggestions: "
'RLB Code
Private IncidentRptLoadedFromDB As Boolean
Private RelatedOfficers(3) As String
'********
Dim VUCRSEL(5) As String
Dim ichanged As Boolean
Dim stra As String, nametype As Integer
Dim casesetup As String
Dim TOBOOK, fromfind, schanged As Integer, FROMXREF As Integer, HOLDINDEX As Integer
Dim FOUNDSELECT As Integer
Dim begindate, EndDate As String, ecc As Integer
Dim holdrecv As Integer, HI As String, itmx As ListItem
Dim FROMKEY, lookupc As Integer
Dim automatic(100) As String, BACKTAB As Integer
Dim found35a, nolarceny, tempsave As Integer, tempword As String
Public Sub loadcodes()
Dim db As Database, rs As Recordset, itmx As ListItem
On Error GoTo oderror1
od1:
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("select * from codes")
On Error Resume Next
lactivity.clear
For t% = 0 To 4
    UCRLIST(t%).clear
    sublist(t%).ListItems.clear
Next t%
For t% = 0 To 23
    drugmeasurement(t%).clear
    drugtype(t%).clear
Next t%
For t% = 0 To 4
    activity(t%).ListItems.clear
    gactivity(t%).ListItems.clear
    HOMOCIDE(t%).ListItems.clear
    additional(t%).clear
Next t%
For t% = 0 To 5
    group(t%).clear
Next t%
For t% = 0 To 19
    relationship(t%).clear
Next t%
reportingofficer(0).clear
reportingofficer(1).clear
followupofficer.clear
approvingofficer.clear
If rs.EOF Then
    db.Close
    Exit Sub
End If
rs.MoveFirst
While Not rs.EOF
    Select Case rs("type")
        Case "ucr"
            For t% = 0 To 4
                UCRLIST(t%).AddItem rs("code")
            Next t%
        Case "lactivity"
            lactivity.AddItem rs("code")
        Case "activity"
            For t% = 0 To 4
                Set itmx = activity(t%).ListItems.add(, , rs("code"))
            Next t%
        Case "additional"
            For t% = 0 To 4
                additional(t%).AddItem rs("code")
            Next t%
        Case "homicide"
                For t% = 0 To 4
                    Set itmx = HOMOCIDE(t%).ListItems.add(, , rs("code"))
                Next t%
        Case "subcodes"
            For t% = 0 To 4
                Set itmx = sublist(t%).ListItems.add(, , rs("code"))
            Next t%
        Case "group"
            For t% = 0 To 5
                group(t%).AddItem rs("code")
            Next t%
        Case "drugtype"
            For t% = 0 To 23
                drugtype(t%).AddItem rs("code")
            Next t%
        Case "measure"
            For t% = 0 To 23
                drugmeasurement(t%).AddItem rs("code")
            Next t%
    End Select
    rs.MoveNext
Wend
For t% = 0 To 4
    Set itmx = gactivity(t%).ListItems.add(, , "Juvenile Gang(J)")
    Set itmx = gactivity(t%).ListItems.add(, , "Other Gang(G)")
    Set itmx = gactivity(t%).ListItems.add(, , "No Gang Involvement(N)")
Next t%
Set rs = db.OpenRecordset("select code from relationship order by code")
If Not rs.EOF Then
    rs.MoveFirst
End If
While Not rs.EOF
    For tt% = 0 To 19
        relationship(tt%).AddItem rs("code")
    Next tt%
    rs.MoveNext
Wend
db.Close
On Error GoTo oderror2
od2:
Set db = OpenDatabase(nwl + "lawsuite.mdb")
Set rs = db.OpenRecordset("select profname from [Deputy Query]")
If Not rs.EOF Then
    rs.MoveFirst
End If
While Not rs.EOF
    If Not IsNull(rs("profname")) Then
        reportingofficer(0).AddItem rs("profname")
        reportingofficer(1).AddItem rs("profname")
        followupofficer.AddItem rs("profname")
        approvingofficer.AddItem rs("profname")
    End If
    rs.MoveNext
Wend
Call LOADMAJOR
Call loadoffense
On Error Resume Next
db.Close
Exit Sub
oderror1:
If Err > 3200 Then
    Resume od1
Else
    Resume Next
End If
oderror2:
If Err > 3200 Then
    Resume od2
Else
    Resume Next
End If
End Sub

Private Sub SENDCHAR(CH As String)
SendKeys CH
End Sub

Private Sub sendclosepara()
SendKeys "{)}"
End Sub
Private Sub senddash()
SendKeys "-"
End Sub

Private Sub sendopenpara()
SendKeys "{(}"
End Sub
Private Sub sendslash()
SendKeys "/"
End Sub

Private Sub sendspace()
SendKeys " "
End Sub

Private Sub active_Click()
ichanged = True
End Sub

Private Sub activity_Click(Index As Integer)


If FROMKEY = 1 Then
    FROMKEY = 0
    Exit Sub
End If
For t% = 0 To UCRLIST(Index).ListCount - 1
    If UCRLIST(Index).Selected(t%) = True Then
        tempucr = Mid$(UCRLIST(Index).List(t%), InStr(UCRLIST(Index).List(t%), "(") + 1, 3)
        t% = UCRLIST(Index).ListCount - 1
    End If
Next t%
Dim countitems As Variant
countitems = 0
If tempucr = "250" Or tempucr = "280" Or tempucr = "35A" Or tempucr = "35B" Or tempucr = "39C" Or tempucr = "370" Or tempucr = "520" Then
    For bb% = 1 To activity(Index).ListItems.Count
            If activity(Index).ListItems(bb%).Selected = True Then
               activity(Index).ListItems(bb%).EnsureVisible
               countitems = countitems + 1
               If countitems > 3 Then
                  msg = MsgBox("For offense codes 250, 280, 35A, 35B, 39C, 370 or 520 you may enter only up to three activities.", 48, "Genesis Error Log")
                  activity(Index).ListItems(bb%).Selected = False
               End If
            End If
    Next bb%
End If
activity(Index).Visible = False
If fromfind = 1 Then
    Exit Sub
End If
'---- setfocus logic ----
'         Command13(index).SetFocus
          If Command13(Index).Visible Then
              Command13(Index).SetFocus
           End If
Select Case tempucr
    Case "09C", "13A", "09A", "09B"
        HOMOCIDE(Index).Left = 5500
        HOMOCIDE(Index).Top = pickoffense(Index).Top - 1000
        HOMOCIDE(Index).Visible = True
'---- setfocus logic ----
'                 HOMOCIDE(index).SetFocus
          If HOMOCIDE(Index).Visible Then
              HOMOCIDE(Index).SetFocus
           End If
    Case Else
'---- setfocus logic ----
'                 completedy(index).SetFocus
          If completedy(Index).Visible Then
              completedy(Index).SetFocus
           End If
End Select
End Sub

Private Sub activity_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
ichanged = True
End Sub

Private Sub activity_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    activity(Index).Visible = False
    If fromfind = 0 Then
'---- setfocus logic ----
'                 completedy(index).SetFocus
          If completedy(Index).Visible Then
              completedy(Index).SetFocus
           End If
    End If
Else
If KeyAscii <> Asc(" ") Then
    FROMKEY = 1
End If
End If

End Sub

Private Sub activity_LostFocus(Index As Integer)
activity(Index).Visible = False
If activity(Index).SelectedItem > "" Then
    FROMKEY = 0
    Call activity_Click(Index)
End If
End Sub

Private Sub additional_Click(Index As Integer)
ichanged = True

If FROMKEY = 1 Then
    FROMKEY = 0
    Exit Sub
End If
additional(Index).Visible = False
If fromfind = 1 Then
'---- setfocus logic ----
'             completedy(index).SetFocus
          If completedy(Index).Visible Then
              completedy(Index).SetFocus
           End If
End If

End Sub

Private Sub additional_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> Asc(" ") Then
    FROMKEY = 1
End If

End Sub

Private Sub additional_LostFocus(Index As Integer)
additional(Index).Visible = False
End Sub

Private Sub address_Change(Index As Integer)
ichanged = True
End Sub

Private Sub address_GotFocus(Index As Integer)
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
If address(Index) > "" Then
    Exit Sub
End If
If Index = 1 Then
    If vsname(1) = vsname(0) Then
        address(1) = address(0)
        city(1) = city(0)
        state(1) = state(0)
        zipcode(1) = zipcode(0)
        resident(1).ListIndex = resident(0).ListIndex
        sex(1).ListIndex = sex(0).ListIndex
        ethnicity(1).ListIndex = ethnicity(0).ListIndex
        age(1) = age(0)
        race(1).ListIndex = race(0).ListIndex
        For i% = 10 To 19
            For tv% = 0 To relationship(i% - 10).ListCount - 1
                If relationship(i% - 10).Selected(tv%) = True Then
                    relationship(i% - 10).ListIndex = tv%
                    tv% = relationship(i% - 10).ListCount - 1
                End If
            Next tv%
            relationship(i%).ListIndex = relationship(i% - 10).ListIndex
        Next i%
        HOMEDAYPHONE(1) = HOMEDAYPHONE(0)
        HOMENIGHTPHONE(1) = HOMENIGHTPHONE(0)
        WORKDAYPHONE(1) = WORKDAYPHONE(0)
        WORKNIGHTPHONE(1) = WORKNIGHTPHONE(0)
        Exit Sub
    End If
End If
End Sub


Private Sub address_LostFocus(Index As Integer)
additional(Index).Visible = False
End Sub


Private Sub admclosed_Click()
ichanged = True
End Sub

Private Sub age_Change(Index As Integer)
ichanged = True
End Sub

Private Sub age_Click(Index As Integer)

If Index <> 0 Then
    
End If
age(Index).Refresh
End Sub


Private Sub age_LostFocus(Index As Integer)
If age(Index) = "" Then
    age(Index) = "00"
End If
End Sub

Private Sub alcoholno_Click(Index As Integer)
ichanged = True
End Sub

Private Sub alcoholno_GotFocus(Index As Integer)
If alcoholframe(Index).Top > (-1 * Picture2.Top) And alcoholframe(Index).Top + alcoholframe(Index).Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If alcoholframe(Index).Top > 500 Then
    VScroll1 = alcoholframe(Index).Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub alcoholunknown_Click(Index As Integer)
ichanged = True
End Sub

Private Sub alcoholunknown_GotFocus(Index As Integer)
If alcoholframe(Index).Top > (-1 * Picture2.Top) And alcoholframe(Index).Top + alcoholframe(Index).Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If alcoholframe(Index).Top > 500 Then
    VScroll1 = alcoholframe(Index).Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub alcoholyes_Click(Index As Integer)
ichanged = True
End Sub

Private Sub alcoholyes_GotFocus(Index As Integer)
If alcoholframe(Index).Top > (-1 * Picture2.Top) And alcoholframe(Index).Top + alcoholframe(Index).Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If alcoholframe(Index).Top > 500 Then
    VScroll1 = alcoholframe(Index).Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub ALONE_Click()
ichanged = True
End Sub

Private Sub ao_Click()
If fromfind = 1 Then
    Exit Sub
End If
offenseframe.Left = 750
offenseframe.Top = 500
offenseframe.Visible = True
offenseframe.ZOrder
'pickoffense(3).SetFocus
End Sub

Private Sub approvingofficer_Change()
ichanged = True
End Sub

Private Sub APPROVINGOFFICERDATE_Change()
ichanged = True
End Sub

Private Sub approvingofficerdate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(APPROVINGOFFICERDATE) = 1 Or Len(APPROVINGOFFICERDATE) = 4 Then
    Call sendslash
End If
End If
End Sub

Private Sub approvingofficerdate_GotFocus()
If approvingofficer.ListIndex > -1 And APPROVINGOFFICERDATE = "" Then
    APPROVINGOFFICERDATE = incidentdate(1)
End If

End Sub

Private Sub APPROVINGOFFICERDATE_LostFocus()
If APPROVINGOFFICERDATE > "" And Not IsDate(APPROVINGOFFICERDATE) Then
    msg = MsgBox("Date/Time entered is invalid.", 48, "Genesis Error Log")
    If fromfind = 0 Then
'---- setfocus logic ----
'                 APPROVINGOFFICERDATE.SetFocus
          If APPROVINGOFFICERDATE.Visible Then
              APPROVINGOFFICERDATE.SetFocus
           End If
    End If
End If
APPROVINGOFFICERDATE = Format$(APPROVINGOFFICERDATE, "mm/dd/yyyy")
End Sub



Private Sub approvingofficeRunit_Change()
ichanged = True
End Sub

Private Sub approvingofficeRunit_GotFocus()
If approvingofficeRunit > "" Or approvingofficer = "" Then
    Exit Sub
End If
Dim db As DAO.Database, rs As DAO.Recordset
Set db = DAO.OpenDatabase(nwl + "LAWSUITE.MDB")
Set rs = db.OpenRecordset("SELECT PROFIDNUM FROM PROFESSIONALS WHERE PROFNAME = '" + approvingofficer + "' AND TYPE = 'D'")
If Not rs.EOF Then
    rs.MoveFirst
    If Not IsNull(rs("PROFIDNUM")) Then
        approvingofficeRunit = rs("PROFIDNUM")
    End If
End If
db.Close
End Sub

Private Sub ARREST_Click()
ichanged = True
End Sub

Private Sub ARREST_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If

End Sub

Private Sub arrested18andover_Click()
ichanged = True
End Sub

Private Sub ARRESTEDNEARNO_Click()
ichanged = True
End Sub

Private Sub ARRESTEDNEARYES_Click()
ichanged = True
End Sub

Private Sub arrestedunder18_Click()
ichanged = True
End Sub

Private Sub ASSISTED_Click()
ichanged = True
End Sub

Private Sub BIAS_Click()
ichanged = True

BIAS.Visible = False
On Error Resume Next
If fromfind = 0 Then
'---- setfocus logic ----
'             subjectidentifiedyes.SetFocus
          If subjectidentifiedyes.Visible Then
              subjectidentifiedyes.SetFocus
           End If
End If
On Error GoTo 0
End Sub

Private Sub birthdate_Change()
ichanged = True
If IsDate(BIRTHDATE) Then
    If age(2) > "" Then
        If fromfind = 0 Then
            msg = MsgBox("Age contradicts birthdate.  Would you like to automatically calculate age?", 4, "Genesis Information Log")
            If msg = 6 Then
                age(2) = DateDiff("yyyy", CDate(BIRTHDATE), CDate(Date$))
            End If
        End If
    End If
End If
End Sub

Private Sub BIRTHDATE_LostFocus()
If BIRTHDATE > "" And Not IsDate(BIRTHDATE) Then
    msg = MsgBox("Date/Time entered is invalid.", 48, "Genesis Error Log")
    If fromfind = 0 Then
'---- setfocus logic ----
'                 BIRTHDATE.SetFocus
          If BIRTHDATE.Visible Then
              BIRTHDATE.SetFocus
           End If
    End If
End If
BIRTHDATE = Format$(BIRTHDATE, "mm/dd/yyyy")
End Sub




Private Sub business_Click()
ichanged = True
End Sub

Private Sub city_Change(Index As Integer)
ichanged = True
End Sub

Private Sub city_GotFocus(Index As Integer)
If Index = 2 And vsname(2) = "UNKNOWN" Then
    city(Index) = ""
End If
    
End Sub

Private Sub closedrugframes_Click(Index As Integer)
sdrugframe(Index).Visible = False
If fromfind = 1 Then
    Exit Sub
End If
If Index > 1 Then
'---- setfocus logic ----
'             description(index - 2).SetFocus
          If description(Index - 2).Visible Then
              description(Index - 2).SetFocus
           End If
End If
If Index = 0 Then
'---- setfocus logic ----
'             computerequipment(index).SetFocus
          If computerequipment(Index).Visible Then
              computerequipment(Index).SetFocus
           End If
End If
If Index = 1 Then
'---- setfocus logic ----
'             ARRESTEDNEARYES.SetFocus
          If ARRESTEDNEARYES.Visible Then
              ARRESTEDNEARYES.SetFocus
           End If
End If
End Sub

Private Sub closeoffense_Click()
offenseframe.Visible = False
End Sub

Private Sub closevucrf_Click()
vucrf.Visible = False
If fromfind = 0 Then
'---- setfocus logic ----
'             vsname(1).SetFocus
          If vsname(1).Visible Then
              vsname(1).SetFocus
           End If
End If
End Sub

Private Sub birthdate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(BIRTHDATE) = 1 Or Len(BIRTHDATE) = 4 Then
    Call sendslash
End If
End If
End Sub



Private Sub Command1_Click()
Dim itmx2 As ListItem
Dim selv(999) As String
siDX% = 0
'If vucrlist.ListItems.Count > 0 Then
'    For vv% = 1 To vucrlist.ListItems.Count
'        If vucrlist.ListItems(vv%).Selected Then
'            siDX% = siDX% + 1
'            selv(siDX%) = vucrlist.ListItems(vv%)
'        End If
'    Next vv%
'End If
vucrlist.ListItems.clear
For vv% = 0 To 4
    For vvv% = 0 To UCRLIST(vv%).ListCount - 1
        If UCRLIST(vv%).Selected(vvv%) = True Then
            foundit% = 0
            For ZZ% = 1 To vucrlist.ListItems.Count
                If vucrlist.ListItems(ZZ%) = UCRLIST(vv%).List(vvv%) Then
                    foundit% = 1
                    ZZ% = vucrlist.ListItems.Count
                End If
            Next ZZ%
            If foundit% = 0 Then
                Set itmx2 = vucrlist.ListItems.add(, , UCRLIST(vv%).List(vvv%))
            End If
            vvv% = UCRLIST(vv%).ListCount - 1
        End If
    Next vvv%
Next vv%
For Z% = 1 To vucrlist.ListItems.Count
    vucrlist.ListItems(Z%).Selected = False
Next Z%
For Z% = 1 To vucrlist.ListItems.Count
    'For zz% = 1 To siDX%
    For ZZ% = 1 To 5
        'If vucrlist.ListItems(Z%) = selv(ZZ%) Then
        If InStr(vucrlist.ListItems(Z%), "(" + VUCRSEL(ZZ%) + ")") > 0 And VUCRSEL(ZZ%) > "" Then
            vucrlist.ListItems(Z%).Selected = True
            ZZ% = 5
        End If
    Next ZZ%
Next Z%
If fromfind = 1 Then
    Exit Sub
End If
showit:
vucrf.Top = 4000
vucrf.Left = 500
vucrf.Visible = True
'---- setfocus logic ----
'         vucrlist.SetFocus
          If vucrlist.Visible Then
              vucrlist.SetFocus
           End If
End Sub



Private Sub Command10_Click(Index As Integer)
optimer.Enabled = False
pinfoframe(Index).Visible = False
'description(Index).SetFocus
End Sub

Private Sub Command11_Click(Index As Integer)
If UCRLIST(Index).Visible = False Then
    Call ucrselect(Index)
Else
    UCRLIST(Index).Visible = False
End If
End Sub

Private Sub Command12_Click(Index As Integer)
If fromfind = 1 Then
    Exit Sub
End If
Me.Caption = Command12.Count
For t% = 0 To UCRLIST(Index).ListCount - 1
    If UCRLIST(Index).Selected(t%) = True Then
        tempucr = Mid$(UCRLIST(Index).List(t%), InStr(UCRLIST(Index).List(t%), "(") + 1, 3)
        t% = UCRLIST(Index).ListCount - 1
    End If
Next t%
If tempucr = "09A" Or tempucr = "09B" Or tempucr = "100" Or tempucr = "120" Or tempucr = "11A" Or tempucr = "11B" Or tempucr = "11C" Or tempucr = "11D" Or tempucr = "13A" Or tempucr = "13B" Or tempucr = "13C" Then
    gactivity(Index).Left = 2000
    gactivity(Index).Top = pickoffense(Index).Top - 1000
    gactivity(Index).Visible = True
'---- setfocus logic ----
'             gactivity(index).SetFocus
          If gactivity(Index).Visible Then
              gactivity(Index).SetFocus
           End If
Else
'===== Error 219
If tempucr = "250" Or tempucr = "280" Or tempucr = "35A" Or tempucr = "35B" Or tempucr = "39C" Or tempucr = "370" Or tempucr = "520" Then
    activity(Index).Left = 2000
    activity(Index).Top = pickoffense(Index).Top - 1000
    activity(Index).Visible = True
    activity(Index).SelectedItem.EnsureVisible
'---- setfocus logic ----
'             activity(index).SetFocus
          If activity(Index).Visible Then
              activity(Index).SetFocus
           End If
End If
End If
End Sub

Private Sub Command13_Click(Index As Integer)
For t% = 0 To UCRLIST(Index).ListCount - 1
    If UCRLIST(Index).Selected(t%) = True Then
        tempucr = Mid$(UCRLIST(Index).List(t%), InStr(UCRLIST(Index).List(t%), "(") + 1, 3)
        t% = UCRLIST(Index).ListCount - 1
    End If
Next t%
Dim countitems As Variant
countitems = 0
For aa% = 1 To gactivity(Index).ListItems.Count
    If gactivity(Index).ListItems(aa%).Selected = True Then
        gactivity(Index).ListItems(aa%).EnsureVisible
        countitems = countitems + 1
        If countitems > 2 Then
            msg = MsgBox("For offense codes 09A, 09B, 100, 120, 11A, 11B, 11C, 11D, 13A, 13B Or 13C you may enter only up to two gang related activities.", 48, "Genesis Error Log")
            gactivity(Index).ListItems(aa%).Selected = False
            Exit Sub
        End If
    End If
Next aa%
gactivity(Index).Visible = False
If fromfind = 1 Then
    Exit Sub
End If
Select Case tempucr
    Case "09C", "13A", "09A", "09B"
        HOMOCIDE(Index).Left = 5500
        HOMOCIDE(Index).Top = pickoffense(Index).Top - 1000
        HOMOCIDE(Index).Visible = True
'---- setfocus logic ----
'                 HOMOCIDE(index).SetFocus
          If HOMOCIDE(Index).Visible Then
              HOMOCIDE(Index).SetFocus
           End If
    Case Else
'---- setfocus logic ----
'                 subcode(index).SetFocus
          If subcode(Index).Visible Then
              subcode(Index).SetFocus
           End If
End Select
On Error GoTo 0

End Sub

Private Sub Command14_Click()
If lactivity.Visible = False Then
    lactivity.Left = 5000
    lactivity.Top = policeofficer.Top - 1000
    lactivity.Visible = True
'---- setfocus logic ----
'             lactivity.SetFocus
          If lactivity.Visible Then
              lactivity.SetFocus
           End If
Else
    lactivity.Visible = False
End If
End Sub

Private Sub Command15_Click()
BeginSpellCheck NARRATIVE.Text, NARRATIVE
End Sub

Private Sub Command16_Click()
lstframe.Visible = False
End Sub

Private Sub Command17_Click()
Dim DONE As Boolean
If lstsuggestions.ListIndex = -1 Then
    Exit Sub
End If
For t% = 1 To Len(NARRATIVE.Text)
    If UCase(Mid$(NARRATIVE.Text, t%, Len(checkword))) = UCase(checkword) Then
        If UCase(Mid$(NARRATIVE.Text, t%, Len(checkword))) = Mid$(NARRATIVE.Text, t%, Len(checkword)) Then
            lstsuggestions.List(lstsuggestions.ListIndex) = UCase(lstsuggestions.List(lstsuggestions.ListIndex))
        End If
        NARRATIVE.Text = Left$(NARRATIVE.Text, t% - 1) + lstsuggestions.List(lstsuggestions.ListIndex) + Mid$(NARRATIVE.Text, t% + Len(checkword))
        t% = t% + Len(checkword)
    End If
Next t%

'RLB CODE
For t% = 1 To Len(tempword)
    If UCase(Mid$(tempword, t%, Len(checkword))) = UCase(checkword) Then
        If UCase(Mid$(tempword, t%, Len(checkword))) = Mid$(tempword, t%, Len(checkword)) Then
            lstsuggestions.List(lstsuggestions.ListIndex) = UCase(lstsuggestions.List(lstsuggestions.ListIndex))
        End If
        tempword = Left$(tempword, t% - 1) + lstsuggestions.List(lstsuggestions.ListIndex) + Mid$(tempword, t% + Len(checkword))
        t% = t% + Len(checkword)
    End If
Next t%
'*********


Call spellcheck(DONE)
If DONE Then
    msg = MsgBox("Spelling check complete.", 48, "Genesis Information Log")
End If
End Sub

Private Sub Command18_Click()
Dim DONE As Boolean

Call spellcheck(DONE)
If DONE Then
    msg = MsgBox("Spelling check complete.", 48, "Genesis Information Log")
End If
End Sub

Private Sub Command19_Click()
If Left$(casesetup, 1) = "1" Then
    year12 = True
    month12 = False
Else
    year12 = False
    month12 = True
End If
dash3 = Val(Mid$(casesetup, 3, 1))
If Mid$(casesetup, 4, 1) = "1" Then
    year45 = True
    month45 = False
Else
    year45 = False
    month45 = True
End If
If Len(casesetup) > 5 Then
    suffix = Mid$(casesetup, 6)
Else
    suffix = ""
End If
    
caseframe.Top = 1000
caseframe.Left = 2600
caseframe.Visible = True
End Sub

Private Sub Command2_Click(Index As Integer)
If pucrlist(Index).ListIndex = -1 Then
    pucrlist(Index).clear
    For vv% = 0 To 4
        For vvv% = 0 To UCRLIST(vv%).ListCount - 1
            If UCRLIST(vv%).Selected(vvv%) = True Then
                foundit% = 0
                For yy% = 0 To pucrlist(Index).ListCount - 1
                    If UCRLIST(vv%).List(vvv%) = pucrlist(Index).List(yy%) Then
                        foundit% = 1
                        yy% = pucrlist(Index).ListCount - 1
                    End If
                Next yy%
                If foundit% = 0 Then
                    '===== 081
                    tempucr = Mid$(UCRLIST(vv%).List(vvv%), InStr(UCRLIST(vv%).List(vvv%), "(") + 1, 3)
                    Select Case tempucr
                        Case "100", "120", "200", "210", "220", "23A", "23B", "23C", "23D", "23E", "23F", "23G", "23H", "240", "250", "26A", "26B", "26C", "26D", "26E", "270", "280", "290", "35A", "35B", "39A", "39B", "39C", "39D", "510"
                            pucrlist(Index).AddItem UCRLIST(vv%).List(vvv%)
                    End Select
                End If
                vvv% = UCRLIST(vv%).ListCount - 1
            End If
        Next vvv%
    Next vv%
Else
    HP = pucrlist(Index).ListIndex
    HC = pucrlist(Index).ListCount
    pucrlist(Index).clear
    For vv% = 0 To 4
        For vvv% = 0 To UCRLIST(vv%).ListCount - 1
            If UCRLIST(vv%).Selected(vvv%) = True Then
                foundit% = 0
                For yy% = 0 To pucrlist(Index).ListCount - 1
                    If UCRLIST(vv%).List(vvv%) = pucrlist(Index).List(yy%) Then
                        foundit% = 1
                        yy% = pucrlist(Index).ListCount - 1
                    End If
                Next yy%
                If foundit% = 0 Then
                    '===== 081
                    tempucr = Mid$(UCRLIST(vv%).List(vvv%), InStr(UCRLIST(vv%).List(vvv%), "(") + 1, 3)
                    Select Case tempucr
                        Case "100", "120", "200", "210", "220", "23A", "23B", "23C", "23D", "23E", "23F", "23G", "23H", "240", "250", "26A", "26B", "26C", "26D", "26E", "270", "280", "290", "35A", "35B", "39A", "39B", "39C", "39D", "510"
                            pucrlist(Index).AddItem UCRLIST(vv%).List(vvv%)
                    End Select
                End If
                vvv% = UCRLIST(vv%).ListCount - 1
            End If
        Next vvv%
    Next vv%
    If HC = pucrlist(Index).ListCount Then
        FROMKEY = 1
        pucrlist(Index).ListIndex = HP
    End If
End If
If pucrlist(Index).ListCount = 0 Then
    Exit Sub
End If
'If fromfind = 1 Then
'    Exit Sub
'End If
pinfoframe(Index).Left = Command2(Index).Left + 1200
pinfoframe(Index).Top = 12300
pinfoframe(Index).Visible = True
'rlb code
pinfoframe(Index).Top = (Command2(Index).Top - CLng(0.5 * pinfoframe(Index).Height))
pinfoframe(Index).Left = Command2(Index).Left
If ((pinfoframe(Index).Left + pinfoframe(Index).Width) > pinfoframe(Index).Container.Width) Then
    pinfoframe(Index).Left = pinfoframe(Index).Container.Width - pinfoframe(Index).Width
End If
If pinfoframe(Index).Top > (-1 * Picture2.Top) And pinfoframe(Index).Top + pinfoframe(Index).Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If pinfoframe(Index).Top > 500 Then
    VScroll1 = pinfoframe(Index).Top - 500
Else
    VScroll1 = 0
End If
End If
'********
pucrlist(Index).Visible = True
'---- setfocus logic ----
'         pucrlist(index).SetFocus
          If pucrlist(Index).Visible Then
              pucrlist(Index).SetFocus
           End If
End Sub



Private Sub Command22_Click(Index As Integer)
pucrlist(Index).ListIndex = -1
group(Index).ListIndex = -1
majorlist(Index).ListIndex = -1
minorlist(Index).ListIndex = -1
End Sub

Private Sub Command23_Click()
flagframe.Visible = False
End Sub

Private Sub Command24_Click()
Open "c:\flaglist" For Output As #1
For t% = 1 To incidentlist.ListItems.Count
    If incidentlist.ListItems(t%).Selected Then
        Print #1, incidentlist.ListItems(t%)
    End If
Next t%
Close #1
MsgBox "Flag list saved.", 48, "Genesis Information Log"
End Sub

Private Sub Command25_Click()
Dim flaglist(9999) As String, fidx As Integer
fidx = 0
Open "c:\flaglist" For Input As #1
While Not EOF(1)
    Line Input #1, a$
    fidx = fidx + 1
    flaglist(fidx) = a$
Wend
Close #1
For t% = 1 To incidentlist.ListItems.Count
    incidentlist.ListItems(t%).Selected = False
    For tt% = 1 To fidx
        If incidentlist.ListItems(t%) = flaglist(tt%) Then
            incidentlist.ListItems(t%).Selected = True
            tt% = fidx
        End If
    Next tt%
Next t%
MsgBox "Flag list restored.", 48, "Genesis Information Log"

End Sub

Private Sub completedn_Click(Index As Integer)
ichanged = True
End Sub

Private Sub completedy_Click(Index As Integer)
ichanged = True
End Sub

Private Sub computerequipment_Click(Index As Integer)
ichanged = True
End Sub

Private Sub DATEOFARREST_Change()
ichanged = True
End Sub

Private Sub DATEOFARREST_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If

End Sub

Private Sub DATERECOVERED_Change(Index As Integer)
ichanged = True
End Sub

Private Sub DEPARTINGTIME_Change()
ichanged = True
End Sub

Private Sub description_Change(Index As Integer)
ichanged = True
End Sub

Private Sub Description_LostFocus(Index As Integer)
If description(Index) = "" And sdrugframe(Index + 2).Visible = False Then
'---- setfocus logic ----
'             subjectidentifiedyes.SetFocus
          If subjectidentifiedyes.Visible Then
              subjectidentifiedyes.SetFocus
           End If
End If
End Sub

Private Sub DETECTIVE_Click()
ichanged = True
End Sub

Private Sub dispatchdate_Change()
ichanged = True
End Sub

Private Sub DISPATCHTIME_Change()
ichanged = True
End Sub

Private Sub drugmeasurement_Click(Index As Integer)
ichanged = True

End Sub

Private Sub drugmeasurement_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    If drugmeasurement(Index).ListIndex > -1 Then
        drugmeasurement(Index).ListIndex = -1
    End If
End If

End Sub

Private Sub drugsunknown_Click(Index As Integer)
ichanged = True

End Sub

Private Sub drugsyes_Click(Index As Integer)
ichanged = True

End Sub

Private Sub drugtype_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    If drugtype(Index).ListIndex > -1 Then
        drugtype(Index).ListIndex = -1
    End If
End If
End Sub

Private Sub dtoffense_Click()
ichanged = True

End Sub

Private Sub ELOCATIONNUMBER_Change()
ichanged = True

End Sub

Private Sub entered_Change(Index As Integer)
ichanged = True

End Sub

Private Sub ethnicity_Click(Index As Integer)
ichanged = True

End Sub

Private Sub ethnicity_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        ethnicity(Index).ListIndex = -1
    End If

End Sub

Private Sub EXCEPTIONALCLEARANCEDATE_Change()
ichanged = True

End Sub

Private Sub EXCEPTIONALCLEARANCEDATE_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If

End Sub

Private Sub exclear18andover_Click()
ichanged = True

End Sub

Private Sub exclearunder18_Click()
ichanged = True

End Sub

Private Sub extraditiondenied_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 And Shift = 0 Then
    extraditiondenied = False
End If

End Sub

Private Sub eyes_Change(Index As Integer)
ichanged = True

End Sub

Private Sub financialinstitution_Click()
ichanged = True

End Sub

Private Sub flagbutton_Click()
Dim db As Database, rs As Recordset
For t% = 1 To incidentlist.ListItems.Count
    If incidentlist.ListItems(t%).Selected Then
        Set db = OpenDatabase(nwi + "incident.mdb")
        Set rs = db.OpenRecordset("select exportfile from incidentsupport where incidentnumber = '" + incidentlist.ListItems(t%) + "'")
        If Not rs.EOF Then
            rs.Edit
            rs("exportfile") = "FLAG"
            rs.Update
        End If
    End If
Next t%
flagframe.Visible = False
MsgBox "Incidents flagged for send on next submission.", 48, "Genesis Information Log"

End Sub

Private Sub followupno_Click()
ichanged = True

End Sub

Private Sub followupofficer_Change()
ichanged = True

End Sub

Private Sub FOLLOWUPOFFICERDATE_Change()
ichanged = True

End Sub

Private Sub FOLLOWUPOFFICERUNIT_Change()
ichanged = True

End Sub

Private Sub FOLLOWUPOFFICERUNIT_GotFocus()
If FOLLOWUPOFFICERUNIT > "" Or followupofficer = "" Then
    Exit Sub
End If
Dim db As DAO.Database, rs As DAO.Recordset
Set db = DAO.OpenDatabase(nwl + "LAWSUITE.MDB")
Set rs = db.OpenRecordset("SELECT PROFIDNUM FROM PROFESSIONALS WHERE PROFNAME = '" + followupofficer + "' AND TYPE = 'D'")
If Not rs.EOF Then
    rs.MoveFirst
    If Not IsNull(rs("PROFIDNUM")) Then
        FOLLOWUPOFFICERUNIT = rs("PROFIDNUM")
    End If
End If
db.Close
End Sub

Private Sub followupyes_Click()
ichanged = True

End Sub

Private Sub FORCEDENTRYN_Click(Index As Integer)
ichanged = True

End Sub

Private Sub FORCEDENTRYY_Click(Index As Integer)
ichanged = True

End Sub

Private Sub government_Click()
ichanged = True

End Sub

Private Sub hair_Change(Index As Integer)
ichanged = True

End Sub

Private Sub holdoff_Click()
ichanged = True

End Sub

Private Sub HOMEDAYPHONE_Change(Index As Integer)
ichanged = True

End Sub

Private Sub HOMENIGHTPHONE_Change(Index As Integer)
ichanged = True

End Sub

Private Sub HOMOCIDE_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
ichanged = True
If FROMKEY = 1 Then
    FROMKEY = 0
    Exit Sub
End If
If HOMOCIDE(Index).SelectedItem Is Nothing Then
    Exit Sub
End If
If fromfind = 1 Then
    Exit Sub
End If

End Sub

Private Sub HOMOCIDE_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> Asc(" ") Then
    FROMKEY = 1
End If
End Sub

Private Sub HOMOCIDE_LostFocus(Index As Integer)
HOMOCIDE(Index).Visible = False
For xx% = 1 To HOMOCIDE(Index).ListItems.Count
    If HOMOCIDE(Index).ListItems(xx%).Selected Then
        If InStr(HOMOCIDE(Index).ListItems(xx%), "(20)") > 0 Or InStr(HOMOCIDE(Index).ListItems(xx%), "(21)") > 0 Then
            additional(Index).Left = 7000
            additional(Index).Top = 1000
            additional(Index).Visible = True
            If fromfind = 0 Then
'---- setfocus logic ----
'                         additional(index).SetFocus
          If additional(Index).Visible Then
              additional(Index).SetFocus
           End If
            End If
            xx% = HOMOCIDE(Index).ListItems.Count
        End If
    End If
Next xx%
If additional(Index).Visible = True Then
    Exit Sub
End If
On Error Resume Next
If fromfind = 1 Then
    Exit Sub
End If
If BACKTAB = 0 Then
'---- setfocus logic ----
'             subcode(index).SetFocus
          If subcode(Index).Visible Then
              subcode(Index).SetFocus
           End If
End If
On Error GoTo 0
End Sub

Private Sub ht_Change(Index As Integer)
ichanged = True

End Sub

Private Sub incidentdate_Change(Index As Integer)
ichanged = True

End Sub

Private Sub incidentlocation_Change()
ichanged = True

End Sub

Private Sub incidentnumber_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Or incidentnumber = "" Then
    Exit Sub
End If
If Left$(UCase(incidentnumber), 8) = "SAVELIST" Then
    Call forcesavelist
    Exit Sub
End If
If Left$(UCase(incidentnumber), 4) = "SAVE" Then
    Call forcesave
    Exit Sub
End If
If Len(incidentnumber) < 12 And Len(incidentnumber) > 0 Then
    incidentnumber = incidentnumber + Space$(12 - Len(incidentnumber))
End If
incidentnumber = UCase(incidentnumber)
If pickoffense(0).ListIndex = -1 And vsname(0) = "" And vsname(1) = "" Then
    Call incidentnumber_Click
End If
End Sub

Private Sub incidentzipcode_Change()
ichanged = True

End Sub

Private Sub individual_Click()
ichanged = True

End Sub

Private Sub injury_ItemClick(ByVal Item As MSComctlLib.ListItem)
ichanged = True

End Sub

Private Sub JAIL_Click()
ichanged = True

End Sub

Private Sub JURISDICTIONRECOVERY_Change()
ichanged = True

End Sub

Private Sub JURISDICTIONTHEFT_Change()
ichanged = True

End Sub

Private Sub juvenilenocustody_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 And Shift = 0 Then
    juvenilenocustody = False
End If
End Sub

Private Sub lactivity_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        lactivity.ListIndex = -1
    End If
End Sub

Private Sub locali_Click()
ichanged = True

End Sub

Private Sub LOCATIONNUMBER_Change(Index As Integer)
ichanged = True

End Sub

Private Sub minorlist_Click(Index As Integer)
ichanged = True

End Sub

Private Sub month12_Click()
ichanged = True

End Sub

Private Sub month45_Click()
ichanged = True

End Sub

Private Sub mugshot_DblClick()
If mugshot.Height = 650 Then
    mugshot.Height = 2600
    mugshot.Width = 3000
Else
    mugshot.Height = 650
    mugshot.Width = 750
End If
msframe.Height = mugshot.Height
msframe.Width = mugshot.Width
End Sub

Private Sub na_Click()
ichanged = True

End Sub

Private Sub NARRATIVE_Change()
ichanged = True

End Sub

Private Sub NARRATIVE_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift = vbCtrlMask) And (KeyCode = vbKeyF2) Then
        Call Command15_Click
    End If
End Sub

Private Sub NONVISIBLEINJURYNO_Click()
ichanged = True

End Sub

Private Sub NONVISIBLEINJURYYES_Click()
ichanged = True

End Sub

Private Sub noprosecution_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 And Shift = 0 Then
    noprosecution = False
End If

End Sub

Private Sub numvehicle_Change(Index As Integer)
ichanged = True

End Sub

Private Sub offenderdeath_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 And Shift = 0 Then
    offenderdeath = False
End If
End Sub

Private Sub ONEMANVEHICLE_Click()
ichanged = True

End Sub

Private Sub onpaper_Click()
ichanged = True

End Sub

Private Sub other_Click()
ichanged = True

End Sub

Private Sub peculiarities_Change(Index As Integer)
ichanged = True

End Sub

'RLB Code



'Private Sub majorlist_Click(Index As Integer)
'
'    Dim db As Database, rs As Recordset
'
'    Set db = OpenDatabase(networkpath + "incident.mdb")
'    Set rs = db.OpenRecordset("select minor from pgroup where major = '" & majorlist(Index).List(majorlist(Index).ListIndex) & "'")
'
'    minorlist(Index).clear
'
'    If Not rs.EOF Then
'        rs.MoveFirst
'
'        While Not rs.EOF
'            minorlist(Index).AddItem rs("minor")
'            rs.MoveNext
'        Wend
'    End If
'
 '    db.Close

'End Sub
'********

Private Sub pickoffense_GotFocus(Index As Integer)
    pickoffense(Index).Height = 6 * pickoffense(Index).Height
    pickoffense(Index).Width = 2 * pickoffense(Index).Width
    pickoffense(Index).ZOrder
End Sub

Private Sub pickoffense_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        pickoffense(Index).ListIndex = -1
    End If
        
End Sub



Private Sub pinfoframe_click(Index As Integer)
'RLB CODE
pinfoframe(Index).ZOrder
'*********
End Sub

Private Sub Command2_GotFocus(Index As Integer)
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub Command20_Click()
Screen.MousePointer = 11
Open nwi + "caseset.tag" For Output As #1
If year12 Then
    casesetup = "1"
Else
    casesetup = "0"
End If
If month12 Then
    casesetup = casesetup + "1"
Else
    casesetup = casesetup + "0"
End If
If dash3 = 1 Then
    casesetup = casesetup + "1"
Else
    casesetup = casesetup + "0"
End If
If year45 Then
    casesetup = casesetup + "1"
Else
    casesetup = casesetup + "0"
End If
If moth45 Then
    casesetup = casesetup + "1"
Else
    casesetup = casesetup + "0"
End If
If suffix > "" Then
    casesetup = casesetup + suffix
End If
Print #1, casesetup
Close #1
Screen.MousePointer = 0
End Sub

Private Sub Command21_Click()
caseframe.Visible = False
End Sub

Private Sub Command3_Click(Index As Integer)
'If pickoffense(index).ListIndex = -1 Then
'    Exit Sub
'End If

lookupframe.Left = 2600
lookupc = Index
lookupframe.Top = 3600
'lookupc = ""
lookupframe.Visible = True
lookup = pickoffense(Index).List(pickoffense(Index).ListIndex)
If fromfind = 0 Then
'---- setfocus logic ----
'             lookup.SetFocus
          If lookup.Visible Then
              lookup.SetFocus
           End If
End If
End Sub

Private Sub Command3_GotFocus(Index As Integer)
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If

End Sub

Private Sub Command4_Click()
If lookuplist.ListIndex > -1 Then
    For t% = 0 To UCRLIST(lookupc).ListCount - 1
        If Mid$(UCRLIST(lookupc).List(t%), InStr(UCRLIST(lookupc).List(t%), "(") + 1, 3) = Mid$(lookuplist.List(lookuplist.ListIndex), InStr(lookuplist.List(lookuplist.ListIndex), "(") + 1, 3) Then
            UCRLIST(lookupc).Selected(t%) = True
            t% = UCRLIST(lookupc).ListCount
            'RLB Code
            If Not (AutoSelectOffense((lookupc), _
                (Mid$(lookuplist.List(lookuplist.ListIndex), _
                        InStr(lookuplist.List(lookuplist.ListIndex), " - ") + 3)))) Then
                    
                    pickoffense(lookupc).ListIndex = -1
                    
            End If
            '***********
        End If
    Next t%
Else
    msg = MsgBox("No OFFENSE has been selected.", vbOK, "Genesis Information Log")
    pickoffense(lookupc).ListIndex = -1
End If

lookupframe.Visible = False
'UCRLIST(lookupc).SetFocus
End Sub

Private Sub Command5_Click()
If lookup = "" Then
    Exit Sub
End If
qaa$ = ""
numcomma% = 0
For t% = 1 To Len(lookup)
    If Mid$(lookup, t%, 1) = "," Then
        numcomma% = numcomma% + 1
    End If
Next t%
If numcomma% > 9 Then
    msg = MsgBox("This feature will only allow up to 10 words.", 48, "Genesis Error Log")
    Exit Sub
End If
temp$ = lookup
For t% = 1 To numcomma%
    If t% = 1 Then
        qaa$ = "offense like '*" + UCase(Left$(temp$, InStr(temp$, ",") - 1)) + "*' "
        qaC$ = "CODE like '*" + UCase(Left$(temp$, InStr(temp$, ",") - 1)) + "*' "
    Else
        qaa$ = qaa$ + " and oFfense like '*" + UCase(Left$(temp$, InStr(temp$, ",") - 1)) + "*' "
        qaC$ = qaC$ + " and CODE like '*" + UCase(Left$(temp$, InStr(temp$, ",") - 1)) + "*' "
    End If
    temp$ = Mid$(lookup, InStr(lookup, ",") + 1)
Next t%
If numcomma% = 0 Then
    qaa$ = "offense like '*" + UCase(lookup) + "*' "
    qaC$ = "CODE like '*" + UCase(lookup) + "*' "
Else
    qaa$ = qaa$ + "and offense like '*" + UCase(temp$) + "*'"
    qaC$ = qaC$ + "and CODE like '*" + UCase(temp$) + "*'"
End If
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("select offense,UCR from OFFENSE where " + qaa$ + " order by offense")
On Error Resume Next
lookuplist.clear
TP% = 1
If rs.EOF Then
    Set rs = db.OpenRecordset("select CODE from UCR where " + qaC$ + " order by CODE")
    If rs.EOF Then
        lookuplist.AddItem "NO MATCHING OFFENSES FOUND."
        db.Close
        Exit Sub
    Else
        TP% = 2
    End If
End If
rs.MoveFirst
While Not rs.EOF
    If TP% = 2 Then
        lookuplist.AddItem rs("CODE")
    Else
        lookuplist.AddItem rs("ucr") + " - " + rs("offense")
    End If
    rs.MoveNext
Wend
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub


Private Sub Command6_Click()
If fromfind = 1 Then
    Exit Sub
End If
If BIAS.Visible = True Then
    BIAS.Visible = False
Else
    BIAS.Top = 16500
    BIAS.Left = 7000
    BIAS.Visible = True
'---- setfocus logic ----
'             BIAS.SetFocus
          If BIAS.Visible Then
              BIAS.SetFocus
           End If
End If
End Sub

Private Sub Command7_Click()
If fromfind = 1 Then
    Exit Sub
End If
sdrugframe(0).Left = 2000
sdrugframe(0).Top = Command7.Top
sdrugframe(0).Visible = True
'---- setfocus logic ----
'         drugtype(0).SetFocus
          If drugtype(0).Visible Then
              drugtype(0).SetFocus
           End If
End Sub

Private Sub Command8_Click(Index As Integer)
relationshipframe(Index).Visible = False
If fromfind = 0 Then
'---- setfocus logic ----
'             resident(index).SetFocus
          If resident(Index).Visible Then
              resident(Index).SetFocus
           End If
End If
End Sub

Private Sub Command9_Click()
If fromfind = 1 Then
    Exit Sub
End If
sdrugframe(1).Left = 2000
sdrugframe(1).Top = Command9.Top
sdrugframe(1).Visible = True
'---- setfocus logic ----
'         drugtype(3).SetFocus
          If drugtype(3).Visible Then
              drugtype(3).SetFocus
           End If
End Sub




Private Sub completedy_GotFocus(Index As Integer)
a = 1
End Sub

Private Sub DATEOFARREST_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(DATEOFARREST) = 1 Or Len(DATEOFARREST) = 4 Then
    Call sendslash
End If
End If

End Sub

Private Sub DATERECOVERED_GotFocus(Index As Integer)
If DATERECOVERED(Index) = "" Then
    DATERECOVERED(Index) = Format$(Date$, "mm/dd/yyyy")
End If
End Sub

Private Sub daterecovered_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(DATERECOVERED(Index)) = 1 Or Len(DATERECOVERED(Index)) = 4 Then
    Call sendslash
End If
End If
End Sub

Private Sub DATERECOVERED_LostFocus(Index As Integer)
If DATERECOVERED(Index) > "" And Not IsDate(DATERECOVERED(Index)) Then
    msg = MsgBox("Date/Time entered is invalid.", 48, "Genesis Error Log")
    If fromfind = 0 Then
'---- setfocus logic ----
'                 DATERECOVERED(index).SetFocus
          If DATERECOVERED(Index).Visible Then
              DATERECOVERED(Index).SetFocus
           End If
    End If
End If
DATERECOVERED(Index) = Format$(DATERECOVERED(Index), "mm/dd/yyyy")
DATERECOVERED(Index).Visible = False
If DATERECOVERED(Index) = "Date" Then
    DATERECOVERED(Index) = ""
End If
'=====Data Item 18 and 19
If fromfind = 1 Then
    Exit Sub
End If
If BACKTAB = 0 Then
    If Mid$(pucrlist(Index).List(pucrlist(Index).ListIndex), InStr(pucrlist(Index).List(pucrlist(Index).ListIndex), "(") + 1, 3) = "240" Then
        tempgroup = Val(Mid$(group(Index).List(group(Index).ListIndex), InStr(group(Index).List(group(Index).ListIndex), "(") + 1, 2))
        If tempgroup = 3 Or tempgroup = 5 Or tempgroup = 24 Or tempgroup = 28 Or tempgroup = 37 Then
            '===== Error 357
            For u% = 0 To 4
                For uu% = 0 To UCRLIST(u%).ListCount - 1
                    If pucrlist(Index).List(pucrlist(Index).ListIndex) = UCRLIST(u%).List(uu%) Then
                        If completedy(u%) Then
                            numvehicle((Index Mod 6) + 6).Visible = True
'---- setfocus logic ----
'                                     numvehicle((index Mod 6) + 6).SetFocus
          If numvehicle((indexMod6) + 6).Visible Then
              numvehicle((indexMod6) + 6).SetFocus
           End If
                            uu% = UCRLIST(u%).ListCount - 1
                            u% = 9
                        End If
                    End If
                Next uu%
            Next u%
            If numvehicle((Index Mod 6) + 6).Visible = False Then
'---- setfocus logic ----
'                         totalvalue(holdrecv + 6).SetFocus
          If totalvalue(holdrecv + 6).Visible Then
              totalvalue(holdrecv + 6).SetFocus
           End If
            End If
        Else
'---- setfocus logic ----
'                     totalvalue(holdrecv + 6).SetFocus
          If totalvalue(holdrecv + 6).Visible Then
              totalvalue(holdrecv + 6).SetFocus
           End If
        End If
    Else
'---- setfocus logic ----
'                 totalvalue(holdrecv + 6).SetFocus
          If totalvalue(holdrecv + 6).Visible Then
              totalvalue(holdrecv + 6).SetFocus
           End If
    End If
End If
End Sub


Private Sub DEPARTINGTIME_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(DEPARTINGTIME) = 1 Then
    Call sendcolon
End If
End If

End Sub

Private Sub DEPARTINGTIME_LostFocus()
If Len(DEPARTINGTIME) = 4 And InStr("0123456789", Left$(DEPARTINGTIME, 1)) > 0 And InStr("0123456789", Mid$(DEPARTINGTIME, 2, 1)) > 0 And InStr("0123456789", Mid$(DEPARTINGTIME, 3, 1)) > 0 And InStr("0123456789", Right$(DEPARTINGTIME, 1)) > 0 Then
    DEPARTINGTIME = Left$(DEPARTINGTIME, 2) + ":" + Right$(DEPARTINGTIME, 2)
End If
If DEPARTINGTIME > "" And Not IsDate(DEPARTINGTIME) Then
    msg = MsgBox("Date/Time entered is invalid.", 48, "Genesis Error Log")
    If fromfind = 0 Then
'---- setfocus logic ----
'                 DEPARTINGTIME.SetFocus
          If DEPARTINGTIME.Visible Then
              DEPARTINGTIME.SetFocus
           End If
    End If
End If

End Sub


Private Sub description_GotFocus(Index As Integer)
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub dispatchdate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(dispatchdate) = 1 Or Len(dispatchdate) = 4 Then
    Call sendslash
End If
End If
End Sub
Private Sub dispatchdate_GotFocus()
If dispatchdate = "" Then
    If IsDate(incidentdate(1)) Then
        dispatchdate = incidentdate(1)
    End If
End If
End Sub

Private Sub DISPATCHDATE_LostFocus()
If dispatchdate > "" And Not IsDate(dispatchdate) Then
    msg = MsgBox("Date/Time entered is invalid.", 48, "Genesis Error Log")
    If fromfind = 0 Then
'---- setfocus logic ----
'                 dispatchdate.SetFocus
          If dispatchdate.Visible Then
              dispatchdate.SetFocus
           End If
    End If
End If
dispatchdate = Format$(dispatchdate, "mm/dd/yyyy")
End Sub

Private Sub DISPATCHTIME_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(DISPATCHTIME) = 1 Then
    Call sendcolon
End If
End If

End Sub

Private Sub DISPATCHTIME_LostFocus()
If Len(DISPATCHTIME) = 4 And InStr("0123456789", Left$(DISPATCHTIME, 1)) > 0 And InStr("0123456789", Mid$(DISPATCHTIME, 2, 1)) > 0 And InStr("0123456789", Mid$(DISPATCHTIME, 3, 1)) > 0 And InStr("0123456789", Right$(DISPATCHTIME, 1)) > 0 Then
    DISPATCHTIME = Left$(DISPATCHTIME, 2) + ":" + Right$(DISPATCHTIME, 2)
End If
If DISPATCHTIME > "" And Not IsDate(DISPATCHTIME) Then
    msg = MsgBox("Date/Time entered is invalid.", 48, "Genesis Error Log")
    If fromfind = 0 Then
'---- setfocus logic ----
'                 DISPATCHTIME.SetFocus
          If DISPATCHTIME.Visible Then
              DISPATCHTIME.SetFocus
           End If
    End If
End If

End Sub


Private Sub drugsno_Click(Index As Integer)
ichanged = True

If drugsno(Index) Then
    For y% = (Index * 3) To (Index * 3) + 2
        drugtype(y%).ListIndex = -1
        drugamt(y%) = ""
        drugmeasurement(y%).ListIndex = -1
    Next y%
End If
End Sub

Private Sub drugsno_GotFocus(Index As Integer)
If Frame11(Index).Top > (-1 * Picture2.Top) And Frame11(Index).Top + Frame11(Index).Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Frame11(Index).Top > 500 Then
    VScroll1 = Frame11(Index).Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub drugsunknown_GotFocus(Index As Integer)
If Frame11(Index).Top > (-1 * Picture2.Top) And Frame11(Index).Top + Frame11(Index).Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Frame11(Index).Top > 500 Then
    VScroll1 = Frame11(Index).Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub drugsyes_GotFocus(Index As Integer)
If Frame11(Index).Top > (-1 * Picture2.Top) And Frame11(Index).Top + Frame11(Index).Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Frame11(Index).Top > 500 Then
    VScroll1 = Frame11(Index).Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub drugtype_Click(Index As Integer)
ichanged = True

If Index > 2 Then
    
End If
If fromfind = 1 Then
    Exit Sub
End If
If drugamt(Index).Visible = True Then
'---- setfocus logic ----
'             drugamt(index).SetFocus
          If drugamt(Index).Visible Then
              drugamt(Index).SetFocus
           End If
End If
End Sub

Private Sub drugtype_GotFocus(Index As Integer)
If Index < 3 Then
    If sdrugframe(0).Top > (-1 * Picture2.Top) And sdrugframe(0).Top + sdrugframe(0).Height < (-1 * Picture2.Top) + VScroll1.Height Then
    Else
    If sdrugframe(0).Top > 500 Then
        VScroll1 = sdrugframe(0).Top - 500
    Else
        VScroll1 = 0
    End If
    End If
Else
    If sdrugframe(1).Top > (-1 * Picture2.Top) And sdrugframe(1).Top + sdrugframe(1).Height < (-1 * Picture2.Top) + VScroll1.Height Then
    Else
    If sdrugframe(1).Top > 500 Then
        VScroll1 = sdrugframe(1).Top - 500
    Else
        VScroll1 = 0
    End If
    End If
End If
End Sub

Private Sub exceptionalclearancedate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(EXCEPTIONALCLEARANCEDATE) = 1 Or Len(EXCEPTIONALCLEARANCEDATE) = 4 Then
    Call sendslash
End If
End If
End Sub
Private Sub EXCEPTIONALCLEARANCEDATE_LostFocus()
If EXCEPTIONALCLEARANCEDATE > "" And Not IsDate(EXCEPTIONALCLEARANCEDATE) Then
    msg = MsgBox("Date/Time entered is invalid.", 48, "Genesis Error Log")
    If fromfind = 0 Then
'---- setfocus logic ----
'                 EXCEPTIONALCLEARANCEDATE.SetFocus
          If EXCEPTIONALCLEARANCEDATE.Visible Then
              EXCEPTIONALCLEARANCEDATE.SetFocus
           End If
    End If
End If
EXCEPTIONALCLEARANCEDATE = Format$(EXCEPTIONALCLEARANCEDATE, "mm/dd/yyyy")
EXCEPTIONALCLEARANCEDATE.Visible = False
End Sub


Private Sub extraditiondenied_Click()
ichanged = True

If fromfind = 1 Then
    Exit Sub
End If
EXCEPTIONALCLEARANCEDATE.Visible = True
'---- setfocus logic ----
'         EXCEPTIONALCLEARANCEDATE.SetFocus
          If EXCEPTIONALCLEARANCEDATE.Visible Then
              EXCEPTIONALCLEARANCEDATE.SetFocus
           End If

End Sub






Private Sub FOLLOWUPOFFICERDATE_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(FOLLOWUPOFFICERDATE) = 1 Or Len(FOLLOWUPOFFICERDATE) = 4 Then
    Call sendslash
End If
End If

End Sub

Private Sub FOLLOWUPOFFICERUNIT_LostFocus()
If BACKTAB = 0 Then
'---- setfocus logic ----
'             onpaper.SetFocus
          If onpaper.Visible Then
              onpaper.SetFocus
           End If
End If
End Sub

Private Sub Form_Resize()
Picture2.Move 0, 0
'With VScroll1
'    .Max = Picture2.Height - Picture1.Height
'End With
'VScroll1.Visible = (Picture1.Height < Picture2.Height)

End Sub

Private Sub gactivity_Click(Index As Integer)
ichanged = True


If FROMKEY = 1 Then
    FROMKEY = 0
    Exit Sub
End If
If gactivity(Index).SelectedItem Is Nothing Then
    Exit Sub
Else
    gactivity(Index).Visible = False
    If fromfind = 0 Then
'---- setfocus logic ----
'                 Command13(index).SetFocus
          If Command13(Index).Visible Then
              Command13(Index).SetFocus
           End If
    End If
End If
End Sub

Private Sub gactivity_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    gactivity(Index).Visible = False
    If fromfind = 0 Then
'---- setfocus logic ----
'                 completedy(index).SetFocus
          If completedy(Index).Visible Then
              completedy(Index).SetFocus
           End If
    End If
Else
If KeyAscii <> Asc(" ") Then
    FROMKEY = 1
End If
End If

End Sub

Private Sub gactivity_LostFocus(Index As Integer)
gactivity(Index).Visible = False
If gactivity(Index).SelectedItem > "" Then
    FROMKEY = 0
    Call gactivity_Click(Index)
End If
End Sub

Private Sub group_Click(Index As Integer)
ichanged = True
If FROMKEY = 1 Then
    FROMKEY = 0
    Exit Sub
End If


On Error Resume Next
If fromfind = 1 Then
    Exit Sub
End If
If Mid$(group(Index).List(group(Index).ListIndex), InStr(group(Index).List(group(Index).ListIndex), "(") + 1, 2) = "10" Then
    sdrugframe(Index + 2).Left = 2000
    sdrugframe(Index + 2).Top = description(Index).Top - sdrugframe(Index + 2).Height - 100
    sdrugframe(Index + 2).Visible = True
    'RLB Code
   ' sdrugframe(index + 2).ZOrder
    '********
'---- setfocus logic ----
'             drugtype(index + 2).SetFocus
          If drugtype(Index + 2).Visible Then
              drugtype(Index + 2).SetFocus
           End If
Else
    sdrugframe(Index + 2).Visible = False
'---- setfocus logic ----
'             totalvalue(index).SetFocus
          If totalvalue(Index).Visible Then
              totalvalue(Index).SetFocus
           End If
End If
On Error GoTo 0
End Sub

Private Sub group_GotFocus(Index As Integer)
a = 1
End Sub

Private Sub group_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> Asc(" ") Then
    FROMKEY = 1
End If
End Sub

Private Sub group_LostFocus(Index As Integer)
'GLENN
If BACKTAB = 0 Then
    If fromfind = 0 Then
'---- setfocus logic ----
'                 description(index).SetFocus
          If description(Index).Visible Then
              description(Index).SetFocus
           End If
    End If
End If
'GLENN
End Sub


Private Sub Form_Load()
'lblSugs.Caption = NUM_OF_SUGS & CStr(0)
goingelsewhere = False
nametype = 1
'RLB Code
    Dim blnSprvsr As Boolean
    
    If (frmLogin.SUPERVISOR = 1) Or (frmLogin.ISUPERVISOR = 1) Then blnSprvsr = True Else blnSprvsr = False
           
    For x% = 1 To Toolbar1.Buttons.Count
        If UCase(Toolbar1.Buttons(x%).Caption) = "SETUP" Then
            Toolbar1.Buttons(x%).Enabled = blnSprvsr
            Exit For
        End If
    Next x%
'*********
    
        
fromexport = 0
For t% = 0 To Forms.Count - 1
    If Forms(t%).Name = "xref" Then
        FROMXREF = 1
    End If
    If Forms(t%).Name = "iexport" Then
        fromexport = 1
    End If
Next t%
If fromexport = 0 Then
    incident.WindowState = 2
Else
    incident.WindowState = 1
End If
Dim db As Database, rs As Recordset
FROMKEY = 0
FOUNDSELECT = 0
On Error Resume Next
Kill "*.dsk"
casesetup = "10001"
Open nwi + "caseset.tag" For Input As #1
Line Input #1, ABC$
Close #1
If ABC$ > "" Then
    casesetup = ABC$
End If
schangeds = 0
Me.Height = 7600
Me.Width = 11700
Me.Top = 0
Me.Left = 0
On Error GoTo oderror
od:
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("Select state,orinumber from system")
On Error Resume Next
If rs.EOF Then
'===== Error 001, 002, 201, 301, 401, 501, 601, 701
    msg = MsgBox("ORI number is missing.  No exports can be generated.  Contact technical support.", 48, "Genesis Error Log")
    db.Close
    Exit Sub
End If
rs.MoveFirst
orinumber = rs("orinumber")
'===== Error 052, 086, 104, 204, 304, 404, 504, 604, 704
If Len(orinumber) <> 9 Then
    msg = MsgBox("ORI number format is incorrect.  No exports can be generated.  Contact technical support.", 48, "Genesis Error Log")
    db.Close
    Exit Sub
End If
'===== Error 059
If Not IsNull(rs("state")) Then
    If rs("state") <> Left$(orinumber, 2) Then
        msg = MsgBox("ORI number format is incorrect.  No exports can be generated.  Contact technical support.", 48, "Genesis Error Log")
        db.Close
        Exit Sub
    End If
Else
    msg = MsgBox("ORI number format is incorrect.  No exports can be generated.  Contact technical support.", 48, "Genesis Error Log")
    db.Close
    Exit Sub
End If
Screen.MousePointer = 11
Call loadcodes
On Error GoTo 0
Call clearroutine(0)
If fromexport = 0 Then
    Call loadupkey
End If
With Picture2
    .AutoSize = True
    .Move 0, 0
End With
Picture1.Height = Picture2.Height
VScroll1.Max = Picture1.Height
'VScroll1.Max = Picture2.Height - Picture1.Height
VScroll1.LargeChange = VScroll1.Max / 10
VScroll1.SmallChange = VScroll1.Max / 100
'VScroll1.Visible = (Picture1.Height < Picture2.Height)
On Error Resume Next
If fromexport = 0 Then
    Call defaultcodes
End If
On Error GoTo getoutf
Open "pP.TAG" For Input As #1
Line Input #1, a$
incidentnumber = a$
Close #1
On Error GoTo 0
Call incidentnumber_Click
Kill "pP.TAG"
On Error Resume Next
'---- setfocus logic ----
'         incidentnumber.SetFocus
          If incidentnumber.Visible Then
              incidentnumber.SetFocus
           End If
On Error GoTo 0
getoutf:

schanged = 0
Screen.MousePointer = 0
ichanged = False
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Kill "*.dsk"
On Error GoTo 0
Set incident = Nothing
End Sub




Private Sub incidentdate_GotFocus(Index As Integer)
If Index = 1 And incidentdate(0) > "" And incidentdate(1) = "" Then
    incidentdate(1) = incidentdate(0)
End If
End Sub

Private Sub incidentdate_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(incidentdate(Index)) = 1 Or Len(incidentdate(Index)) = 4 Then
    Call sendslash
End If
End If
End Sub
Private Sub incidentdate_LostFocus(Index As Integer)
If incidentdate(Index) > "" And Not IsDate(incidentdate(Index)) Then
    msg = MsgBox("Date/Time entered is invalid.", 48, "Genesis Error Log")
    If fromfind = 0 Then
'---- setfocus logic ----
'                 incidentdate(index).SetFocus
          If incidentdate(Index).Visible Then
              incidentdate(Index).SetFocus
           End If
    End If
End If
incidentdate(Index) = Format$(incidentdate(Index), "mm/dd/yyyy")
dtoffense = (incidentdate(0)) + " " + TIMEOFOFFENSE(0)
End Sub

Friend Sub incidentnumber_Click()
For t% = 0 To Forms.Count - 1
    If LCase(Forms(t%).Name) = "search" Or LCase(Forms(t%).Name) = "temprevw" Then
        Exit Sub
        t% = Forms.Count - 1
    End If
Next t%
If incidentnumber = "" Then
    Exit Sub
End If
If Len(incidentnumber) < 12 And Len(incidentnumber) > 0 Then
    incidentnumber = incidentnumber + Space$(12 - Len(incidentnumber))
End If
HI = incidentnumber
Screen.MousePointer = 11
Call clearroutine(1)
incidentnumber = HI
Call findincident(1)
ichanged = False
Picture2.Refresh
Screen.MousePointer = 0
On Error Resume Next
holdoff.Visible = True
'---- setfocus logic ----
'         onpaper.SetFocus
          If onpaper.Visible Then
              onpaper.SetFocus
           End If
optimer.Enabled = False
If holdoff.Visible = True Then
    optimer.Enabled = True
End If
If incidentnumber = "" Then
    incidentnumber = HI
End If
'---- setfocus logic ----
'         pickoffense(0).SetFocus
          If pickoffense(0).Visible Then
              pickoffense(0).SetFocus
           End If
holdoff.Visible = False
On Error GoTo 0
VScroll1.Value = VScroll1.Min
End Sub


Private Sub Incidentnumber_GotFocus()
If HI > "" Then
    incidentnumber = HI
    HI = ""
Else
If incidentnumber = "" Then
    Dim db As Database, rs, rs2 As Recordset, mi As String
    On Error GoTo oderror
od:
    Set db = OpenDatabase(nwi + "INCIDENT.MDB")
    If Left(casesetup, 1) = "1" Then
        likepattern$ = Right(Date, 2)
    Else
        likepattern$ = Left(Date, 2)
    End If
    If Mid(casesetup, 3, 1) = "1" Then
        likepattern$ = likepattern$ + "-"
    End If
    If Left(casesetup, 4) = "1" Then
        likepattern$ = likepattern$ + Right(Date, 2) + "?????"
    Else
        likepattern$ = likepattern$ + Left(Date, 2) + "?????"
    End If
    If Len(casesetup$) > 5 Then
        likepattern$ = likepattern$ + Mid(casesetup, 6)
    End If
    'If Mid$(casesetup, 5, 1) = "0" Then
    '    If Len(casesetup) > 5 Then
    '        likepattern$ = String$(9, "?") + Mid$(casesetup, 6)
    '    Else
    '        likepattern$ = String$(9, "?")
    '    End If
    'Else
    '    If Len(casesetup) > 5 Then
    '        likepattern$ = String$(10, "?") + Mid$(casesetup, 6)
    '    Else
    '        likepattern$ = String$(10, "?")
    '    End If
    'End If
    likepattern$ = likepattern$ + String$(12 - Len(likepattern$), " ")
    Set rs = db.OpenRecordset("SELECT MAX(INCIDENTNUMBER) AS MI FROM INCIDENTREPORTC WHERE INCIDENTNUMBER LIKE '" + likepattern$ + "'")
    If Not rs.EOF Then
        rs.MoveFirst
        If Not IsNull(rs("MI")) Then
            tempi$ = ""
            mi = rs("mi")
            If Mid$(mi, 3, 1) = "-" Then
                starty% = 6
            Else
                starty% = 5
            End If
            For yy% = starty% To Len(mi)
                If InStr("0123456789", Mid$(mi, yy%, 1)) > 0 Then
                    tempi$ = tempi$ + Mid$(mi, yy%, 1)
                End If
            Next yy%
            Set rs2 = db.OpenRecordset("SELECT MAX(INCIDENTNUMBER) AS bI FROM badcheck WHERE INCIDENTNUMBER LIKE '" + likepattern$ + "'")
            If Not rs2.EOF Then
                rs2.MoveFirst
                If Not IsNull(rs2("bi")) Then
                    tempb$ = ""
                    If Mid$(rs2("bi"), 3, 1) = "-" Then
                        starty% = 6
                    Else
                        starty% = 5
                    End If
                    For yy% = starty% To Len(rs2("bi"))
                        If InStr("0123456789", Mid$(rs2("bi"), yy%, 1)) > 0 Then
                            tempb$ = tempb$ + Mid$(rs2("bi"), yy%, 1)
                        End If
                    Next yy%
                    If Val(tempb$) > Val(tempi$) Then
                        mi = rs2("bi")
                        tempi$ = tempb$
                    End If
                End If
            End If
            Select Case Left$(casesetup, 1)
                Case "1"
                    If Left$(mi, 2) <> Right$(Date$, 2) Then
                        tempi$ = "00000"
                    End If
                    Select Case Mid$(casesetup, 3, 1)
                        Case "1"
                            Select Case Val(Mid$(casesetup, 4, 1))
                                Case 0
                                    Select Case Len(casesetup)
                                        Case 5
                                            incidentnumber = Right$(Date$, 2) + "-" + Left$(Date$, 2) + Format$(Val(tempi$) + 1, "00000")
                                        Case Else
                                            incidentnumber = Right$(Date$, 2) + "-" + Left$(Date$, 2) + Format$(Val(tempi$) + 1, "00000") + Mid$(casesetup, 6)
                                    End Select
                                Case 1
                                    Select Case Len(casesetup)
                                        Case 5
                                            incidentnumber = Right$(Date$, 2) + "-" + Right$(Date$, 2) + Format$(Val(tempi$) + 1, "00000")
                                        Case Else
                                            incidentnumber = Right$(Date$, 2) + "-" + Right$(Date$, 2) + Format$(Val(tempi$) + 1, "00000") + Mid$(casesetup, 6)
                                    End Select
                            End Select
                        Case "0"
                            Select Case Val(Mid$(casesetup, 4, 1))
                                Case 0
                                    Select Case Len(casesetup)
                                        Case 5
                                            incidentnumber = Right$(Date$, 2) + Left$(Date$, 2) + Format$(Val(tempi$) + 1, "00000")
                                        Case Else
                                            incidentnumber = Right$(Date$, 2) + Left$(Date$, 2) + Format$(Val(tempi$) + 1, "00000") + Mid$(casesetup, 6)
                                    End Select
                                Case 1
                                    Select Case Len(casesetup)
                                        Case 5
                                            incidentnumber = Right$(Date$, 2) + Right$(Date$, 2) + Format$(Val(tempi$) + 1, "00000")
                                        Case Else
                                            incidentnumber = Right$(Date$, 2) + Right$(Date$, 2) + Format$(Val(tempi$) + 1, "00000") + Mid$(casesetup, 6)
                                    End Select
                            End Select
                    End Select
                Case "0"
                    If Left$(mi, 2) <> Left$(Date$, 2) Then
                        tempi$ = "00000"
                    End If
                    Select Case Mid$(casesetup, 3, 1)
                        Case "1"
                            Select Case Val(Mid$(casesetup, 4, 1))
                                Case 0
                                    Select Case Len(casesetup)
                                        Case 5
                                            incidentnumber = Left$(Date$, 2) + "-" + Left$(Date$, 2) + Format$(Val(tempi$) + 1, "00000")
                                        Case Else
                                            incidentnumber = Left$(Date$, 2) + "-" + Left$(Date$, 2) + Format$(Val(tempi$) + 1, "00000") + Mid$(casesetup, 6)
                                    End Select
                                Case 1
                                    Select Case Len(casesetup)
                                        Case 5
                                            incidentnumber = Left$(Date$, 2) + "-" + Right$(Date$, 2) + Format$(Val(tempi$) + 1, "00000")
                                        Case Else
                                            incidentnumber = Left$(Date$, 2) + "-" + Right$(Date$, 2) + Format$(Val(tempi$) + 1, "00000") + Mid$(casesetup, 6)
                                    End Select
                            End Select
                        Case "0"
                            Select Case Val(Mid$(casesetup, 4, 1))
                                Case 0
                                    Select Case Len(casesetup)
                                        Case 5
                                            incidentnumber = Left$(Date$, 2) + Left$(Date$, 2) + Format$(Val(tempi$) + 1, "00000")
                                        Case Else
                                            incidentnumber = Left$(Date$, 2) + Left$(Date$, 2) + Format$(Val(tempi$) + 1, "00000") + Mid$(casesetup, 6)
                                    End Select
                                Case 1
                                    Select Case Len(casesetup)
                                        Case 5
                                            incidentnumber = Left$(Date$, 2) + Right$(Date$, 2) + Format$(Val(tempi$) + 1, "00000")
                                        Case Else
                                            incidentnumber = Left$(Date$, 2) + Right$(Date$, 2) + Format$(Val(tempi$) + 1, "00000") + Mid$(casesetup, 6)
                                    End Select
                            End Select
                    End Select
            End Select
    End If
    On Error Resume Next
    db.Close
End If
End If
End If
optimer.Enabled = False
If holdoff.Visible = True Then
    optimer.Enabled = True
End If
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
    Resume
End If

End Sub

Private Sub Incidentnumber_LostFocus()
If goingelsewhere = True Then
    Exit Sub
End If

End Sub


Private Sub JAIL_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If

End Sub

Private Sub JURISDICTIONTHEFT_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If

End Sub

Private Sub juvenilenocustody_Click()
ichanged = True

If fromfind = 1 Then
    Exit Sub
End If
EXCEPTIONALCLEARANCEDATE.Visible = True
'---- setfocus logic ----
'         EXCEPTIONALCLEARANCEDATE.SetFocus
          If EXCEPTIONALCLEARANCEDATE.Visible Then
              EXCEPTIONALCLEARANCEDATE.SetFocus
           End If

End Sub



Private Sub lactivity_Click()
ichanged = True
If lactivity.ListIndex > -1 Then
    lactivity.Visible = False
End If
End Sub

Private Sub lookuplist_Click()
If lookuplist.ListIndex > -1 Then
    If InStr(lookuplist.List(lookuplist.ListIndex), "-") > 0 Then
        For t% = 0 To UCRLIST(lookupc).ListCount - 1
            If Mid$(UCRLIST(lookupc).List(t%), InStr(UCRLIST(lookupc).List(t%), "(") + 1, 3) = Mid$(lookuplist.List(lookuplist.ListIndex), InStr(lookuplist.List(lookuplist.ListIndex), "(") + 1, 3) Then
                UCRLIST(lookupc).Selected(t%) = True
                t% = UCRLIST(lookupc).ListCount
                'RLB Code
                If Not (AutoSelectOffense((lookupc), _
                    (Mid$(lookuplist.List(lookuplist.ListIndex), _
                        InStr(lookuplist.List(lookuplist.ListIndex), " - ") + 3)))) Then
                    
                    pickoffense(lookupc).ListIndex = -1
                
                End If
                '*********
                
            End If
        Next t%
    Else
        For t% = 0 To UCRLIST(lookupc).ListCount - 1
            If Mid$(UCRLIST(lookupc).List(t%), InStr(UCRLIST(lookupc).List(t%), "(") + 1, 3) = Mid$(lookuplist.List(lookuplist.ListIndex), InStr(lookuplist.List(lookuplist.ListIndex), "(") + 1, 3) Then
                UCRLIST(lookupc).Selected(t%) = True
                t% = UCRLIST(lookupc).ListCount
                'RLB Code
                If Not (AutoSelectOffense(lookupc, lookuplist.List(lookuplist.ListIndex))) Then
                    pickoffense(lookupc).ListIndex = -1
                End If
                '********
                
            End If
        Next t%
    End If
    lookupframe.Visible = False
End If
End Sub






Private Sub majorlist_Click(Index As Integer)
ichanged = True
If majorlist(Index).ListIndex = -1 Then
    Exit Sub
End If
Call setminorlist(Index)
End Sub

Private Sub narrative_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If

End Sub

Private Sub noprosecution_Click()
ichanged = True

If fromfind = 1 Then
    Exit Sub
End If

EXCEPTIONALCLEARANCEDATE.Visible = True
'---- setfocus logic ----
'         EXCEPTIONALCLEARANCEDATE.SetFocus
          If EXCEPTIONALCLEARANCEDATE.Visible Then
              EXCEPTIONALCLEARANCEDATE.SetFocus
           End If

End Sub

Private Sub numvehicle_GotFocus(Index As Integer)
If numvehicle(Index) = "" Then
    numvehicle(Index) = "#Veh"
End If
End Sub

Private Sub numvehicle_LostFocus(Index As Integer)
numvehicle(Index).Visible = False
If numvehicle(Index) = "#Veh" Then
    numvehicle(Index) = ""
End If
If fromfind = 1 Then
    Exit Sub
End If
If BACKTAB = 0 Then
'---- setfocus logic ----
'             totalvalue(holdrecv + 6).SetFocus
          If totalvalue(holdrecv + 6).Visible Then
              totalvalue(holdrecv + 6).SetFocus
           End If
Else
    BACKTAB = 0
End If
End Sub

Private Sub offenderdeath_Click()
ichanged = True

If fromfind = 1 Then
    Exit Sub
End If

EXCEPTIONALCLEARANCEDATE.Visible = True
'---- setfocus logic ----
'         EXCEPTIONALCLEARANCEDATE.SetFocus
          If EXCEPTIONALCLEARANCEDATE.Visible Then
              EXCEPTIONALCLEARANCEDATE.SetFocus
           End If

End Sub




Private Sub offenderdeath_GotFocus()
If Frame18.Top > (-1 * Picture2.Top) And Frame18.Top + Frame18.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Frame18.Top > 500 Then
    VScroll1 = Frame18.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub onpaper_GotFocus()
'optimer.Enabled = False
'If holdoff.Visible = True Then
'    optimer.Enabled = True
'End If
'If incidentnumber = "" Then
'    incidentnumber = HI
'End If
'pickoffense(0).SetFocus
End Sub

Private Sub optimer_Timer()
VScroll1 = VScroll1.Min
optimer.Enabled = False
holdoff.Visible = False
pickoffense(Index).Height = 255
pickoffense(Index).Width = 2295

End Sub

Private Sub peculiarities_GotFocus(Index As Integer)
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If

End Sub

Private Sub pickoffense_Click(Index As Integer)
ichanged = True
If FROMKEY = 1 Then
    FROMKEY = 0
    Exit Sub
End If
End Sub

Private Sub pickoffense_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyAscii <> Asc(" ") Then
    FROMKEY = 1
End If
End Sub


Private Sub pickoffense_LostFocus(Index As Integer)
If pickoffense(Index).ListIndex > -1 Then
    pickoffense(Index).TopIndex = pickoffense(Index).ListIndex
End If
If pickoffense(Index).Width > 2296 Then
    pickoffense(Index).Height = pickoffense(Index).Height / 6
    pickoffense(Index).Width = 2296
End If

FROMKEY = 0
If Len(pickoffense(Index).List(pickoffense(Index).ListIndex)) > 50 Then
    msg = MsgBox("Offense is being truncated to maximum of 50 characters.", 48, "Too Long")
    pickoffense(Index).List(pickoffense(Index).ListIndex) = Left$(pickoffense(Index).List(pickoffense(Index).ListIndex), 50)
End If
Dim itmx2 As ListItem
Dim selv(999) As String
siDX% = 0
If vucrlist.ListItems.Count > 0 Then
    For vv% = 1 To vucrlist.ListItems.Count
        If vucrlist.ListItems(vv%).Selected Then
            siDX% = siDX% + 1
            selv(siDX%) = vucrlist.ListItems(vv%)
        End If
    Next vv%
End If
vucrlist.ListItems.clear
For vv% = 0 To 4
    For vvv% = 0 To UCRLIST(vv%).ListCount - 1
        If UCRLIST(vv%).Selected(vvv%) = True Then
            foundit% = 0
            For ZZ% = 1 To vucrlist.ListItems.Count
                If vucrlist.ListItems(ZZ%) = UCRLIST(vv%).List(vvv%) Then
                    foundit% = 1
                    ZZ% = vucrlist.ListItems.Count
                End If
            Next ZZ%
            If foundit% = 0 Then
                Set itmx2 = vucrlist.ListItems.add(, , UCRLIST(vv%).List(vvv%))
            End If
            vvv% = UCRLIST(vv%).ListCount - 1
        End If
    Next vvv%
Next vv%
For Z% = 1 To vucrlist.ListItems.Count
    For ZZ% = 1 To siDX%
        If vucrlist.ListItems(Z%) = selv(ZZ%) Then
            vucrlist.ListItems(Z%).Selected = True
            ZZ% = siDX%
        End If
    Next ZZ%
Next Z%
If fromfind = 1 Then
    Exit Sub
End If
If BACKTAB = 0 Then
    If pickoffense(Index).ListIndex = -1 Then
'---- setfocus logic ----
'                 individual.SetFocus
          If individual.Visible Then
              individual.SetFocus
           End If
    End If
Else
    BACKTAB = 0
End If
End Sub

Private Sub Picture1_Click()
'dateframe.Left = 2400
'dateframe.Top = 1560
'    If IsDate(incidentdate(0)) Then
'        dateselect = incidentdate(0)
'    End If
'If fromfind = 1 Then
'    Exit Sub
'End If''

'dateselect.Visible = True
'dateframe.Caption = "Date Selector"
'dateframe.Visible = True
'dateselect.SetFocus
End Sub




Private Sub policeofficer_Click()
ichanged = True

End Sub

Private Sub premise_GotFocus(Index As Integer)
If Index = 3 Then
    premise(4).Visible = False
End If
premise(Index).Height = 1500
premise(Index).Width = 3500

End Sub

Private Sub premise_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
ichanged = True
If Index = 3 Then
    premise(4).Visible = True
End If

'premise(index).Top = premise(index).Top + 1500
End Sub

Private Sub premise_LostFocus(Index As Integer)
If premise(Index).Height = 1500 Then
    'premise(index).Top = premise(index).Top + 1500
    premise(Index).Height = 300
    premise(Index).Width = 1875
    premise(Index).SelectedItem.EnsureVisible
End If
If Index = 3 Then
    premise(4).Visible = True
End If

End Sub

'RLB Code
Private Sub pucrlist_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 2 Then
       pucrlist(Index).ListIndex = -1
    End If
    
End Sub
'********


Private Sub pucrlist_Click(Index As Integer)

If FROMKEY = 1 Then
    FROMKEY = 0
    Exit Sub
End If
ichanged = True
'===== Data Element 18, 19
Dim tempgroup As Integer, itmx As String
tempgroup = Val(Mid$(group(Index).List(group(Index).ListIndex), InStr(group(Index).List(group(Index).ListIndex), "(") + 1, 2))
On Error Resume Next
Call XGROUP_Click(Index)
'---- setfocus logic ----
'         group(index).SetFocus
          If group(Index).Visible Then
              group(Index).SetFocus
           End If
On Error GoTo 0

End Sub



Private Sub pucrlist_LostFocus(Index As Integer)
If fromfind = 1 Then
    Exit Sub
End If
If BACKTAB = 0 Then
'---- setfocus logic ----
'             group(index).SetFocus
          If group(Index).Visible Then
              group(Index).SetFocus
           End If
Else
    BACKTAB = 0
End If
End Sub

Private Sub race_Click(Index As Integer)
ichanged = True

End Sub

Private Sub race_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        race(Index).ListIndex = -1
    End If

End Sub

Private Sub relationship_Click(Index As Integer)
If fromfind = 0 Then
    ichanged = True
Else
    fromfind = 0
End If
End Sub

Private Sub relationship_LostFocus(Index As Integer)
For tv% = 0 To relationship(Index).ListCount - 1
    If relationship(Index).Selected(tv%) = True Then
        relationship(Index).ListIndex = tv%
        tv% = relationship(Index).ListCount - 1
    End If
Next tv%
If relationship(Index).ListIndex = -1 Then
    If Index < 10 Then
        Call Command8_Click(0)
    Else
        Call Command8_Click(1)
    End If
End If
    
End Sub

Private Sub religiousorganization_Click()
ichanged = True

End Sub

Private Sub reportingofficer_Click(Index As Integer)
ichanged = True

End Sub

Private Sub reportingofficer_GotFocus(Index As Integer)
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If

End Sub

Private Sub reportingofficer_LostFocus(Index As Integer)
If Index = 1 And reportingofficer(Index) = "" Then
'---- setfocus logic ----
'             approvingofficer.SetFocus
          If approvingofficer.Visible Then
              approvingofficer.SetFocus
           End If
End If
End Sub

Private Sub REPORTINGOFFICERDATE_Change(Index As Integer)
ichanged = True

End Sub

Private Sub reportingofficerdate_GotFocus(Index As Integer)
If reportingofficer(Index).ListIndex > -1 And REPORTINGOFFICERDATE(Index) = "" Then
    REPORTINGOFFICERDATE(Index) = incidentdate(1)
End If
End Sub

Private Sub REPORTINGOFFICERDATE_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(REPORTINGOFFICERDATE(Index)) = 1 Or Len(REPORTINGOFFICERDATE(Index)) = 4 Then
    Call sendslash
End If
End If

End Sub

Private Sub REPORTINGOFFICERDATE_LostFocus(Index As Integer)
If REPORTINGOFFICERDATE(Index) > "" And Not IsDate(REPORTINGOFFICERDATE(Index)) Then
    msg = MsgBox("Date/Time entered is invalid.", 48, "Genesis Error Log")
    If fromfind = 0 Then
'---- setfocus logic ----
'                 REPORTINGOFFICERDATE(index).SetFocus
          If REPORTINGOFFICERDATE(Index).Visible Then
              REPORTINGOFFICERDATE(Index).SetFocus
           End If
    End If
End If
REPORTINGOFFICERDATE(Index) = Format$(REPORTINGOFFICERDATE(Index), "mm/dd/yyyy")
End Sub


Private Sub reportingofficeRunit_Change(Index As Integer)
ichanged = True

End Sub

Private Sub reportingofficeRunit_GotFocus(Index As Integer)
If reportingofficeRunit(Index) > "" Or reportingofficer(Index) = "" Then
    Exit Sub
End If
Dim db As DAO.Database, rs As DAO.Recordset
Set db = DAO.OpenDatabase(nwl + "LAWSUITE.MDB")
Set rs = db.OpenRecordset("SELECT PROFIDNUM FROM PROFESSIONALS WHERE PROFNAME = '" + reportingofficer(Index) + "' AND TYPE = 'D'")
If Not rs.EOF Then
    rs.MoveFirst
    If Not IsNull(rs("PROFIDNUM")) Then
        reportingofficeRunit(Index) = rs("PROFIDNUM")
    End If
End If
db.Close
End Sub

Private Sub resident_Click(Index As Integer)
ichanged = True

End Sub

Private Sub resident_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        resident(Index).ListIndex = -1
    End If

End Sub

Private Sub RUNAWAY_Click()
ichanged = True

End Sub

Private Sub RUNAWAY_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If

End Sub

Private Sub setrel_Click(Index As Integer)
On Error Resume Next
relationshipframe(Index).Left = 7000
relationshipframe(Index).Top = setrel(Index).Top - 3000
If fromfind = 1 Then
    Exit Sub
End If

relationshipframe(Index).Visible = True
'---- setfocus logic ----
'         relationship(0 + (index * 10)).SetFocus
          If relationship(0 + (Index * 10)).Visible Then
              relationship(0 + (Index * 10)).SetFocus
           End If
On Error GoTo 0
End Sub



Private Sub setrel_GotFocus(Index As Integer)
If vsname(0) = vsname(1) Then
    For t% = 10 To 19
        If relationship(t%).ListIndex = -1 Then
            relationship(t%).ListIndex = relationship(t% - 10).ListIndex
        End If
    Next t%
End If

End Sub

Private Sub sex_Click(Index As Integer)
ichanged = True

End Sub

Private Sub sex_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        sex(Index).ListIndex = -1
    End If

End Sub

Private Sub societypublic_Click()
ichanged = True

End Sub

Private Sub state_Change(Index As Integer)
ichanged = True

End Sub

Private Sub state_GotFocus(Index As Integer)
If Index = 2 And vsname(2) = "UNKNOWN" Then
    state(Index) = ""
End If

End Sub

Private Sub state_LostFocus(Index As Integer)
'X% = SendMessage(state.hwnd, CB_SHOWDROPDOWN, 0, 0&)
If Len(state(Index)) > 2 Then
    msg = MsgBox("Only 2 character states are allowed in this field.", 48, "Genesis Error Log")
    state(Index) = ""
    Exit Sub
End If
End Sub




Private Sub subcode_Click(Index As Integer)
ichanged = True
If sublist(Index).Visible = True Then
    sublist(Index).Visible = False
Else
    sublist(Index).Left = 4500
    sublist(Index).Top = pickoffense(Index).Top - 1500
    sublist(Index).Visible = True
    If fromfind = 0 Then
'---- setfocus logic ----
'                 sublist(index).SetFocus
          If sublist(Index).Visible Then
              sublist(Index).SetFocus
           End If
    End If
End If
End Sub

Private Sub subcode_LostFocus(Index As Integer)
If BACKTAB = 1 Then
'---- setfocus logic ----
'             Command13(index).SetFocus
          If Command13(Index).Visible Then
              Command13(Index).SetFocus
           End If
    BACKTAB = 0
Else
'---- setfocus logic ----
'             completedy(index).SetFocus
          If completedy(Index).Visible Then
              completedy(Index).SetFocus
           End If
End If
End Sub

Private Sub subjectidentifiedno_Click()
ichanged = True

End Sub

Private Sub subjectidentifiedno_GotFocus()
If Frame19.Top > (-1 * Picture2.Top) And Frame19.Top + Frame19.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Frame19.Top > 500 Then
    VScroll1 = Frame19.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub subjectidentifiedyes_Click()
ichanged = True

End Sub

Private Sub subjectidentifiedyes_GotFocus()
If Frame19.Top > (-1 * Picture2.Top) And Frame19.Top + Frame19.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Frame19.Top > 500 Then
    VScroll1 = Frame19.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub subjectlocatedno_Click()
ichanged = True

End Sub

Private Sub subjectlocatedyes_Click()
ichanged = True

End Sub

Private Sub sublist_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
ichanged = True
On Error GoTo 0

If FROMKEY = 1 Then
    FROMKEY = 0
    Exit Sub
End If
If sublist(Index).SelectedItem Is Nothing Then
    Exit Sub
End If
If fromfind = 1 Then
    Exit Sub
End If
sublist(Index).Visible = False
'---- setfocus logic ----
'         completedy(index).SetFocus
          If completedy(Index).Visible Then
              completedy(Index).SetFocus
           End If

End Sub

Private Sub sublist_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> Asc(" ") Then
    FROMKEY = 1
End If

End Sub

Private Sub sublist_LostFocus(Index As Integer)
If sublist(Index).SelectedItem Is Nothing Then
    Exit Sub
End If
'sublist(Index).Visible = False
On Error Resume Next
If fromfind = 1 Then
    Exit Sub
End If
On Error GoTo 0

End Sub

Private Sub suffix_Change()
ichanged = True

End Sub

Private Sub SUMMONS_Click()
ichanged = True

End Sub

Private Sub SUMMONS_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If

End Sub

Private Sub SUSPECT_Click()
ichanged = True

End Sub

Private Sub SUSPECT_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If

End Sub

Private Sub TIMEARRIVED_Change()
ichanged = True

End Sub

Private Sub TIMEARRIVED_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(TIMEARRIVED) = 1 Then
    Call sendcolon
End If
End If

End Sub

Private Sub TIMEARRIVED_LostFocus()
If Len(TIMEARRIVED) = 4 And InStr("0123456789", Left$(TIMEARRIVED, 1)) > 0 And InStr("0123456789", Mid$(TIMEARRIVED, 2, 1)) > 0 And InStr("0123456789", Mid$(TIMEARRIVED, 3, 1)) > 0 And InStr("0123456789", Right$(TIMEARRIVED, 1)) > 0 Then
    TIMEARRIVED = Left$(TIMEARRIVED, 2) + ":" + Right$(TIMEARRIVED, 2)
End If
If TIMEARRIVED > "" And Not IsDate(TIMEARRIVED) Then
    msg = MsgBox("Date/Time entered is invalid.", 48, "Genesis Error Log")
    If fromfind = 0 Then
'---- setfocus logic ----
'                 TIMEARRIVED.SetFocus
          If TIMEARRIVED.Visible Then
              TIMEARRIVED.SetFocus
           End If
    End If
End If

End Sub


Private Sub TIMEOFARREST_Change()
ichanged = True

End Sub

Private Sub TIMEOFARREST_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(TIMEOFARREST) = 1 Then
    Call sendcolon
End If
End If

End Sub

Private Sub TIMEOFOFFENSE_Change(Index As Integer)
ichanged = True

End Sub

Private Sub TIMEOFOFFENSE_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(TIMEOFOFFENSE(Index)) = 1 Then
    Call sendcolon
End If
End If

End Sub

Private Sub TIMEOFOFFENSE_LostFocus(Index As Integer)
If Len(TIMEOFOFFENSE(Index)) = 4 And InStr("0123456789", Left$(TIMEOFOFFENSE(Index), 1)) > 0 And InStr("0123456789", Mid$(TIMEOFOFFENSE(Index), 2, 1)) > 0 And InStr("0123456789", Mid$(TIMEOFOFFENSE(Index), 3, 1)) > 0 And InStr("0123456789", Right$(TIMEOFOFFENSE(Index), 1)) > 0 Then
    TIMEOFOFFENSE(Index) = Left$(TIMEOFOFFENSE(Index), 2) + ":" + Right$(TIMEOFOFFENSE(Index), 2)
End If
If TIMEOFOFFENSE(Index) > "" And Not IsDate(TIMEOFOFFENSE(Index)) Then
    msg = MsgBox("Date/Time entered is invalid.", 48, "Genesis Error Log")
    If fromfind = 0 Then
'---- setfocus logic ----
'                 TIMEOFOFFENSE(index).SetFocus
          If TIMEOFOFFENSE(Index).Visible Then
              TIMEOFOFFENSE(Index).SetFocus
           End If
    End If
End If
dtoffense = incidentdate(0) + " " + TIMEOFOFFENSE(0)
End Sub

Private Sub TODOTHER_Click()
ichanged = True

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim editerr As Integer
Dim db As Database, rs, rs2 As Recordset
Dim pp(100), ppct As Integer
editerr% = 0
goingelsewhere = False
Select Case Button
    Case "Trans"
        If Dir("c:\mobile.lic") > "" Then
            Screen.MousePointer = 11
            trans.Show
            Screen.MousePointer = 0
        Else
            MsgBox "Licensing for mobile products not found on this installation.", 48, "Genesis Error Log"
            Exit Sub
        End If
    Case "SrvCall"
        goingelsewhere = True
        optimer.Enabled = False
        Screen.MousePointer = 11
        service.Show
        Screen.MousePointer = 0
    Case "Victim"
        MsgBox "Victim's Advocate Package not available in current installation.", 48, "Genesis Error Log"
        Screen.MousePointer = 0
        Exit Sub
        goingelsewhere = True
        optimer.Enabled = False
        Screen.MousePointer = 11
        advocate.Show
        Screen.MousePointer = 0
    Case "Flag"
        Set db = OpenDatabase(nwi + "incident.mdb")
        Set rs = db.OpenRecordset("select distinct incidentnumber from incidentsupport where incidentnumber is not null order by incidentnumber")
        incidentlist.ListItems.clear
        While Not rs.EOF
            Set itmx = incidentlist.ListItems.add(, , rs("incidentnumber"))
            rs.MoveNext
        Wend
        rs.Close
        db.Close
        flagframe.Top = Picture2.Top + 200
        flagframe.Visible = True
        incidentlist.SetFocus
        Screen.MousePointer = 0
    Case "BadCk"
        goingelsewhere = True
        optimer.Enabled = False
        Screen.MousePointer = 11
        badcheck.Show
        Screen.MousePointer = 0
    Case "NCrim"
        goingelsewhere = True
        optimer.Enabled = False
        Screen.MousePointer = 11
        noncrim.Show
        Screen.MousePointer = 0
    Case "Flag"
        If UCase(frmLogin.txtUserName) = "DEMO" And UCase(frmLogin.txtPassword) = "DEMO" Then
            msg = MsgBox("Not available in DEMO version.", 48, "Genesis Information Log")
            Screen.MousePointer = 0
            Exit Sub
        End If
        Set db = OpenDatabase(nwi + "incident.mdb")
        Set rs = db.OpenRecordset("select LASTUPDATE, schanged, exportdate, lastexportdate from incidentsupport where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
        If Not rs.EOF Then
            rs.MoveFirst
            rs.Edit
            'Set rs2 = db.OpenRecordset("select LASTUPDATE, schanged, exportdate, lastexportdate from booking where incidentnumber = " + Chr$(34) + Incidentnumber + Chr$(34))
            'If Not rs2.EOF Then
            '    rs2.MoveFirst
            '    rs2.Edit
            '    If rs2("exportdate") = rs("exportdate") Then
            '        rs2("schanged") = 1
            '        rs2("exportdate") = rs2("lastexportdate")
            '        rs2("LASTUPDATE") = Date$
            '        rs2.Update
            '    End If
            'End If
            rs("schanged") = 1
            rs("exportdate") = rs("lastexportdate")
            rs("LASTUPDATE") = Date$
            rs.Update
        End If
        msg = MsgBox("Errored Incident flagged for re-export.", 48, "Genesis Information Log")
    
    Case "NxtPg"
        If incidentnumber = "" Then
            msg = MsgBox("A valid incidentnumber must be present.", 48, "Genesis Error Log")
            Exit Sub
        End If
        schanged = 1
        If schanged = 1 Then
            tempsave = 0
            Screen.MousePointer = 11
            editerr% = 0
            fromfind = 1
            For ty% = 0 To 19
                relationship(ty%).ListIndex = t - 1
                For tv% = 0 To relationship(ty%).ListCount - 1
                    If relationship(ty%).Selected(tv%) = True Then
                        relationship(ty%).ListIndex = tv%
                        tv% = relationship(ty%).ListCount - 1
                    End If
                Next tv%
            Next ty%
            Picture1.Visible = False
            VScroll1.Visible = False
            POPMSG$ = ""
            If ichanged Then
                Call editevent(editerr%, POPMSG$)
                If editerr% = 0 Then
                    Call editvictim(editerr%, POPMSG$)
                    If editerr% = 0 Then
                        Call editsubject(editerr%, POPMSG$)
                        If editerr% = 0 Then
                            Call editadministrative(editerr%, POPMSG$)
                            If editerr% = 0 Then
                                Call editproperty(editerr%, POPMSG$)
                            End If
                        End If
                    End If
                End If
                If editerr% = 0 Then
                    tempsave = 0
                    If UCase(frmLogin.txtUserName) = "DEMO" And UCase(frmLogin.txtPassword) = "DEMO" Then
                    Else
                        Call saveincident
                        ichanged = False
                    End If
                Else
                    MsgBox POPMSG$, 48, "Genesis Error Log"
                    Picture1.Visible = True
                    VScroll1.Visible = True
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
        End If
        Picture1.Visible = True
        VScroll1.Visible = True
        HI = incidentnumber
        HP = Val(PAGE)
        HP = HP + 1
        hid = incidentdate(0)
        Open "NP.TAG" For Output As #1
        Print #1, HI
        Print #1, "1"
        Print #1, hid
        Close #1
        Unload incident
        sinciden.WindowState = vbMaximized
        sinciden.Show
    Case "Save"
        If UCase(frmLogin.txtUserName) = "DEMO" And UCase(frmLogin.txtPassword) = "DEMO" Then
            msg = MsgBox("Not available in DEMO version.", 48, "Genesis Information Log")
            Screen.MousePointer = 0
            Exit Sub
        End If
        tempsave = 0
        Screen.MousePointer = 11
        editerr% = 0
        For ty% = 0 To 19
            relationship(ty%).ListIndex = t - 1
            For tv% = 0 To relationship(ty%).ListCount - 1
                If relationship(ty%).Selected(tv%) = True Then
                    relationship(ty%).ListIndex = tv%
                    tv% = relationship(ty%).ListCount - 1
                End If
            Next tv%
        Next ty%
        Picture1.Visible = False
        VScroll1.Visible = False
        POPMSG$ = ""
        Call editevent(editerr%, POPMSG$)
        If editerr% = 0 Then
            Call editvictim(editerr%, POPMSG$)
            If editerr% = 0 Then
                Call editsubject(editerr%, POPMSG$)
                If editerr% = 0 Then
                    Call editadministrative(editerr%, POPMSG$)
                    If editerr% = 0 Then
                        Call editproperty(editerr%, POPMSG$)
                    End If
                End If
            End If
        End If
        If editerr% = 0 Then
            tempsave = 0
            Call saveincident
            Call clearroutine(0)
            Call loadupkey
            ichanged = False
            incidentnumber = ""
            HI = ""
            If Picture2.Visible = True Then
'---- setfocus logic ----
'                         incidentnumber.SetFocus
          If incidentnumber.Visible Then
              incidentnumber.SetFocus
           End If
            End If
        Else
            MsgBox POPMSG$, 48, "Genesis Error Log"
        End If
        Picture1.Visible = True
        VScroll1.Visible = True
        Screen.MousePointer = 0
    Case "Clear"
        
        Screen.MousePointer = 11
        Call clearroutine(0)
        incidentnumber = ""
        Screen.MousePointer = 0
        holdoff.Visible = True
'---- setfocus logic ----
'                 onpaper.SetFocus
          If onpaper.Visible Then
              onpaper.SetFocus
           End If
        optimer.Enabled = False
        If holdoff.Visible = True Then
            optimer.Enabled = True
        End If
        ichanged = False

    Case "Delete"
        If UCase(frmLogin.txtUserName) = "DEMO" And UCase(frmLogin.txtPassword) = "DEMO" Then
            msg = MsgBox("Not available in DEMO version.", 48, "Genesis Information Log")
            Screen.MousePointer = 0
            Exit Sub
        End If
        Screen.MousePointer = 11
        Call deleteroutine
        Screen.MousePointer = 0
    Case "MthRpt"
        If Val(Left$(Date$, 2)) = 1 Then
            dm$ = "12"
            dy$ = Mid$(Str$(Val(Right$(Date$, 4)) - 1), 2)
        Else
            dm$ = Mid$(Str$(Val(Left$(Date$, 2)) - 1), 2)
            dy$ = Right$(Date$, 4)
        End If
        Dim inpm, inpy As String
        inpm = InputBox("Enter numeric month for report.", "Genesis Information Log", dm$)
        If Val(inpm) < 1 Or Val(inpm) > 12 Then
            msg = MsgBox("Invalid month entry.", 48, "Genesis Error Log")
            Exit Sub
        End If
        inpy = InputBox("Enter numeric year for report.", "Genesis Information Log", dy$)
        If Val(inpy) = 0 Then
            msg = MsgBox("Invalid year entry.", 48, "Genesis Error Log")
            Exit Sub
        End If
        Call monthlyreport(inpm, inpy)
'start print button code
    Case "Print"
        'msg = MsgBox("Function not available at this time.", 48, "Not Available")
        'Exit Sub
        If frmLogin.IPRINT <> 1 And frmLogin.ISUPERVISOR <> 1 And frmLogin.SUPERVISOR <> 1 Then
            msg = MsgBox("Insufficient authority for this operation.", 48, "Genesis Error Log")
            Exit Sub
        End If
        If incidentnumber = "" Then
            msg = MsgBox("A valid incident number must be entered.", 48, "Genesis Error Log")
            Exit Sub
        End If
        schanged = 1
        If schanged = 1 Then
            Screen.MousePointer = 11
            editerr% = 0
            For ty% = 0 To 19
                relationship(ty%).ListIndex = t - 1
                For tv% = 0 To relationship(ty%).ListCount - 1
                    If relationship(ty%).Selected(tv%) = True Then
                        relationship(ty%).ListIndex = tv%
                        tv% = relationship(ty%).ListCount - 1
                    End If
                Next tv%
            Next ty%
            Picture1.Visible = False
            VScroll1.Visible = False
            POPMSG$ = ""
            If ichanged Then
                Call editevent(editerr%, POPMSG$)
                If editerr% = 0 Then
                    Call editvictim(editerr%, POPMSG$)
                    If editerr% = 0 Then
                        Call editsubject(editerr%, POPMSG$)
                        If editerr% = 0 Then
                            Call editadministrative(editerr%, POPMSG$)
                            If editerr% = 0 Then
                                Call editproperty(editerr%, POPMSG$)
                            End If
                        End If
                    End If
                End If
                If editerr% = 0 Then
                    tempsave = 0
                    Call saveincident
                    HI = incidentnumber
                    Call loadupkey
                    ichanged = False
                    incidentnumber = HI
                Else
                    MsgBox POPMSG$, 48, "Genesis Error Log"
                    Picture1.Visible = True
                    VScroll1.Visible = True
                    msg = MsgBox("An incident report cannot be printed with errors.", 48, "Genesis Error Log")
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
        End If
        Picture1.Visible = True
        VScroll1.Visible = True
        On Error GoTo bbhtest
        report.ReportFileName = nwi + "incident.rpt"
        report.SelectionFormula = "{incidentreportc.incidentnumber} = '" + incidentnumber + "'"
        report.PrintFileType = crptCrystal
        report.Destination = crptToPrinter
        report.Action = 1
        ppct = 0
        On Error GoTo oderror2
od2:
        Set db = OpenDatabase(nwi + "incident.mdb")
        Set rs = db.OpenRecordset("select page from supplemental WHERE INCIDENTNUMBER = " + Chr$(34) + incidentnumber + Chr$(34) + " order by page union select page from supplementalSUPPORT WHERE INCIDENTNUMBER = " + Chr$(34) + incidentnumber + Chr$(34) + " order by page")
        If Not rs.EOF Then
            msg = MsgBox("Do you wish to print a supplemental report?", 4, "Genesis Information Log")
            If msg = 6 Then
                rs.MoveFirst
                While Not rs.EOF
                    ppct = ppct + 1
                    pp(ppct) = rs("page")
                    rs.MoveNext
                Wend
                db.Close
                For TP% = 1 To ppct
                    report.ReportFileName = nwi + "sincident.rpt"
                    report.SelectionFormula = "{supplemental.incidentnumber} = '" + incidentnumber + "' and {supplemental.page} = " + Str$(pp(TP%))
                    report.PrintFileType = crptCrystal
                    report.Destination = crptToPrinter
                    report.Action = 1
                Next TP%
            End If
        End If
    Screen.MousePointer = 0

'end print button code
        
    Case "Email"
        If frmLogin.IPRINT <> 1 And frmLogin.ISUPERVISOR <> 1 And frmLogin.SUPERVISOR <> 1 Then
            msg = MsgBox("Insufficient authority for this operation.", 48, "Genesis Error Log")
            Exit Sub
        End If
        If incidentnumber = "" Then
            msg = MsgBox("A valid incident number must be entered.", 48, "Genesis Error Log")
            Exit Sub
        End If
        On Error GoTo bbhtest
        inp = InputBox("Email Address", "Genesis Information Log", "")
        If inp = "" Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        On Error GoTo mailerr
        report.ReportFileName = nwi + "incident.rpt"
        report.SelectionFormula = "{incidentreportc.incidentnumber} = '" + incidentnumber + "'"
        report.Destination = crptMapi
        report.PrintFileType = crptRTF
        report.EmailSubject = "Incident Report " + incidentnumber
        report.EMailToList = inp
        report.Action = 1
        ppct = 0
        On Error GoTo oderror2a
od2a:
        Set db = OpenDatabase(nwi + "incident.mdb")
        Set rs = db.OpenRecordset("select page from supplemental WHERE INCIDENTNUMBER = " + Chr$(34) + incidentnumber + Chr$(34) + " order by page union select page from supplementalSUPPORT WHERE INCIDENTNUMBER = " + Chr$(34) + incidentnumber + Chr$(34) + " order by page")
        On Error Resume Next
        If Not rs.EOF Then
            rs.MoveFirst
            While Not rs.EOF
                ppct = ppct + 1
                pp(ppct) = rs("page")
                rs.MoveNext
            Wend
        End If
        db.Close
        For TP% = 1 To ppct
            report.ReportFileName = nwi + "sincident.rpt"
            report.PrintFileType = crptRTF
            report.SelectionFormula = "{supplemental.incidentnumber} = '" + incidentnumber + "' and {supplemental.page} = " + Str$(pp(TP%))
            report.Destination = crptMapi
            report.Action = 1
        Next TP%
        Screen.MousePointer = 0

    Case "Book"
        If (arrestedunder18 = 0 And arrested18andover = 0) Or DATEOFARREST = "" Or TIMEOFARREST = "" Then
            msg = MsgBox("All arrest information (Date/Time of Arrest, Arrested Under 18, and Arrested 18 and Over) must be entered prior to accessing the Booking Report.", 48, "Genesis Error Log")
'---- setfocus logic ----
'                     DATEOFARREST.SetFocus
          If DATEOFARREST.Visible Then
              DATEOFARREST.SetFocus
           End If
            If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
            Else
            If Me.ActiveControl.Top > 500 Then
                VScroll1 = Me.ActiveControl.Top - 500
            Else
                VScroll1 = 0
            End If
            End If
            Exit Sub
        End If
        If IsDate(EXCEPTIONALCLEARANCEDATE) Then
            msg = MsgBox("Bookings are not allowed for Exceptional Clearance.", 48, "Genesis Error Log")
            EXCEPTIONALCLEARANCEDATE.Visible = True
'---- setfocus logic ----
'                     EXCEPTIONALCLEARANCEDATE.SetFocus
          If EXCEPTIONALCLEARANCEDATE.Visible Then
              EXCEPTIONALCLEARANCEDATE.SetFocus
           End If
            If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
            Else
            If Me.ActiveControl.Top > 500 Then
                VScroll1 = Me.ActiveControl.Top - 500
            Else
                VScroll1 = 0
            End If
            End If
            Exit Sub
        End If
        TOBOOK = 1
        'If Changed = 1 Then
            tempsave = 0
            Screen.MousePointer = 11
            editerr% = 0
            Picture1.Visible = False
            VScroll1.Visible = False
            POPMSG$ = ""
            If ichanged Then
                Call editevent(editerr%, POPMSG$)
                If editerr% = 0 Then
                    Call editvictim(editerr%, POPMSG$)
                    If editerr% = 0 Then
                        Call editsubject(editerr%, POPMSG$)
                        If editerr% = 0 Then
                            Call editadministrative(editerr%, POPMSG$)
                            If editerr% = 0 Then
                                Call editproperty(editerr%, POPMSG$)
                            End If
                        End If
                    End If
                End If
                If editerr% = 0 Then
                    Screen.MousePointer = 0
                    tempsave = 0
                    If UCase(frmLogin.txtUserName) = "DEMO" And UCase(frmLogin.txtPassword) = "DEMO" Then
                    Else
                        Call saveincident
                        ichanged = False
                    End If
                Else
                    MsgBox POPMSG$, 48, "Genesis Error Log"
                    Picture1.Visible = True
                    VScroll1.Visible = True
                    Screen.MousePointer = 0
                    Exit Sub
                End If
        End If
        'End If
        Picture1.Visible = True
        VScroll1.Visible = True
        TOBOOK = 0
        Set db = OpenDatabase(nwi + "incident.mdb")
        Set rs = db.OpenRecordset("select max(subject1) as s1 from supplemental where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
        Set rs2 = db.OpenRecordset("select max(subject2) as s2 from supplemental where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
        On Error Resume Next
        m1 = -1
        m2 = -1
        If Not rs.EOF Then
            rs.MoveFirst
            If Not IsNull(rs("s1")) Then
                m1 = rs("s1")
            End If
        Else
            m1 = -1
        End If
        If Not rs2.EOF Then
            rs2.MoveFirst
            If Not IsNull(rs2("s2")) Then
                m2 = rs2("s2")
            End If
        Else
            m2 = -1
        End If
        On Error GoTo 0
        If m1 > m2 Then
            inp = InputBox("Enter the Subject Number (1 - " + Mid$(Str$(m1), 2) + ") for Booking Report.", "Subject #", "1")
            If Val(inp) = 0 Or Val(inp) > m1 Then
                msg = MsgBox("Invalid Subject Number entered.", 48, "Invalid Data")
                Exit Sub
            End If
        Else
        If m2 > m1 Then
            inp = InputBox("Enter the Subject Number (1 - " + Mid$(Str$(m2), 2) + ") for Booking Report.", "Subject #", "1")
            If Val(inp) = 0 Or Val(inp) > m2 Then
                msg = MsgBox("Invalid Subject Number entered.", 48, "Invalid Data")
                Exit Sub
            End If
        Else
            inp = "1"
        End If
        End If
        Dim tname, trace, tsex, tbirthdate, tage, tethnicity, tht, tweight, thair, teyes, tpeculiarities, taddress, tcity, tstate, tzipcode As String
        If inp = "1" Then
            If vsname(2) = "UNKNOWN" Then
                msg = MsgBox("Booking is not allowed for Subject UNKNOWN.", 48, "Invalid Data")
                Exit Sub
            End If
            tname = vsname(2)
            If race(2).ListIndex > -1 Then
                trace = race(2).List(race(2).ListIndex)
            End If
            If sex(2).ListIndex > -1 Then
                tsex = sex(2).List(sex(2).ListIndex)
            End If
            tbirthdate = BIRTHDATE
            tage = age(2)
            If ethnicity(2).ListIndex > -1 Then
                tethnicity = ethnicity(2).List(ethnicity(2).ListIndex)
            End If
            tht = ht(1)
            tweight = weight(1)
            thair = hair(1)
            teyes = eyes(1)
            tpeculiarities = peculiarities(1)
            taddress = address(2)
            tcity = city(2)
            tstate = state(2)
            tzipcode = zipcode(2)
        Else
            Set rs = db.OpenRecordset("select subject1, subject2, victim1, victim2, name1, name2, race1, race2, sex1, sex2, birthdate1, birthdate2, age1, age2, ethnicity1, ethnicity2, height1, height2,weight1, weight2, hair1, hair2, eyes1, eyes2, peculiarities1, peculiarities2, address1, address2, city1, city2, state1, state2, zipcode1, zipcode2 from supplemental where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34) + " and (subject1 = " + inp + " or subject2 =" + inp + ")")
            rs.MoveFirst
            If rs("subject1") = Val(inp) Then
                lu% = 1
            Else
                lu% = 2
            End If
            If rs("name" + Mid$(Str$(lu%), 2)) = "UNKNOWN" Then
                msg = MsgBox("Booking is not allowed for Subject UNKNOWN.", 48, "Invalid Data")
                Exit Sub
            End If
            tname = rs("name" + Mid$(Str$(lu%), 2))
            trace = rs("race" + Mid$(Str$(lu%), 2))
            tsex = rs("sex" + Mid$(Str$(lu%), 2))
            tbirthdate = rs("birthdate" + Mid$(Str$(lu%), 2))
            tage = rs("age" + Mid$(Str$(lu%), 2))
            tethnicity = rs("ethnicity" + Mid$(Str$(lu%), 2))
            tht = rs("height" + Mid$(Str$(lu%), 2))
            tweight = rs("weight" + Mid$(Str$(lu%), 2))
            thair = rs("hair" + Mid$(Str$(lu%), 2))
            teyes = rs("eyes" + Mid$(Str$(lu%), 2))
            tpeculiarities = rs("peculiarities" + Mid$(Str$(lu%), 2))
            taddress = rs("address" + Mid$(Str$(lu%), 2))
            tcity = rs("city" + Mid$(Str$(lu%), 2))
            tstate = rs("state" + Mid$(Str$(lu%), 2))
            tzipcode = rs("zipcode" + Mid$(Str$(lu%), 2))
        End If
        Set db = OpenDatabase(nwl + "lawsuite.mdb")
        Set rs = db.OpenRecordset("select * from people where dpnamelf = " + Chr$(34) + tname + Chr$(34))
        tssn = ""
        tdl = ""
        tdlstate = ""
        talias = ""
        tidnumber = ""
        If Not rs.EOF Then
            rs.MoveFirst
            If Not IsNull(rs("ssn")) Then
                tssn = rs("ssn")
            End If
            If Not IsNull(rs("dl")) Then
                tdl = rs("dl")
            End If
            If Not IsNull(rs("dlstate")) Then
                tdlstate = rs("dlstate")
            End If
            If Not IsNull(rs("alias")) Then
                talias = rs("alias")
            End If
            If Not IsNull(rs("idnumber")) Then
                tidnumber = rs("idnumber")
            End If
        End If
        On Error Resume Next

        'Call CLEARBOOKING
        On Error GoTo 0
'        booking.Hide
        Open "C:\TOBOOKING" For Output As #1
        Print #1, incidentnumber
        Print #1, inp
        Print #1, DATEOFARREST
        Print #1, TIMEOFARREST
        Print #1, tname
        Print #1, trace
        Print #1, tsex
        Print #1, tbirthdate
        Print #1, tage
        Print #1, tethnicity
        Print #1, tht
        Print #1, tweight
        Print #1, thair
        Print #1, teyes
        Print #1, tpeculiarities
        Print #1, taddress
        Print #1, tcity
        Print #1, tstate
        Print #1, tzipcode
        Print #1, tssn
        Print #1, tdl
        Print #1, tdlstate
        Print #1, alias
        Print #1, tidnumber
        Close #1
        On Error GoTo 0
        Unload incident
        booking.Show
        

    Case "Send"
        If UCase(frmLogin.txtUserName) = "DEMO" And UCase(frmLogin.txtPassword) = "DEMO" Then
            msg = MsgBox("Not available in DEMO version.", 48, "Genesis Information Log")
            Screen.MousePointer = 0
            Exit Sub
        End If
        Screen.MousePointer = 11
        Unload incident
        Unload iexport
        iexport.Show
        Screen.MousePointer = 0
    Case "TmpSv"
        If UCase(frmLogin.txtUserName) = "DEMO" And UCase(frmLogin.txtPassword) = "DEMO" Then
            msg = MsgBox("Not available in DEMO version.", 48, "Genesis Information Log")
            Screen.MousePointer = 0
            Exit Sub
        End If
        If incidentnumber = "" Or Not IsDate(incidentdate(0)) Or vsname(1) = "" Then
            msg = MsgBox("An incident number and date and name must be entered for a TEMP SAVE.", 48, "Genesis Error Log")
            Exit Sub
        End If
        tempsave = 1
        Screen.MousePointer = 11
        Call saveincident
        Call clearroutine(0)
        ichanged = False
        incidentnumber = ""
        If Picture2.Visible = True Then
'---- setfocus logic ----
'                     incidentnumber.SetFocus
          If incidentnumber.Visible Then
              incidentnumber.SetFocus
           End If
        End If
        Screen.MousePointer = 0
    Case "TmpLst"
        goingelsewhere = True
        temprevw.WindowState = vbMaximized
        temprevw.Show
    Case "Search"
        Unload Me
        goingelsewhere = True
        Search.Show
    Case "Exit"
        Unload incident
    Case "Setup"
        goingelsewhere = True
        isetup.Show
    
     
End Select

Exit Sub

'CES Code add
checkload:
If Err = 424 Then
    msg = MsgBox("Selected module is not a part of this installation.", 48, "Genesis Error Log")
Else
    msg = MsgBox(Error$(Err), 48, "Genesis Error Log")
End If
Resume Next
'****

bbhtest:
Resume
oderror2:
If Err > 3200 Then
    Resume od2
Else
    Resume Next
End If
oderror2a:
If Err > 3200 Then
    Resume od2a
Else
    Resume Next
End If
mailerr:
msg = MsgBox("Your standard email program (Outlook, Outlook Express, etc.) must be running in order to send an incident report.", 48, "Genesis Error Log")
Resume Next
End Sub
Private Sub loadupkey()
Call loadincident
On Error GoTo oderror2
od2:
Set db = OpenDatabase(nwl + "LAWSUITE.mdb")
Set rs = db.OpenRecordset("select DPnameLF from people")
On Error Resume Next
If Not rs.EOF Then
    rs.MoveFirst
Else
    db.Close
    Exit Sub
End If
V0 = vsname(0)
V1 = vsname(1)
v2 = vsname(2)
vsname(0).clear
vsname(1).clear
vsname(2).clear
While Not rs.EOF
    If Not IsNull(rs("DPnameLF")) Then
        vsname(0).AddItem rs("DPnameLF")
        vsname(1).AddItem rs("DPnameLF")
        vsname(2).AddItem rs("DPnameLF")
    End If
    rs.MoveNext
Wend
db.Close
vsname(0) = V0
vsname(1) = V1
vsname(2) = v2
Exit Sub
oderror2:
If Err > 3200 Then
    Resume od2
Else
    Resume Next
End If
End Sub


Friend Sub editadministrative(editerr, msg As String)
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwi + "incident.mdb")
'If TOBOOK = 0 And (arrestedunder18 = 1 Or arrested18andover = 1) Then
'    Set rs = db.OpenRecordset("select incidentnumber from booking where INCIDENTnumber = " + Chr$(34) + Incidentnumber + Chr$(34))
'    If rs.EOF Then
'        On Error Resume Next
'        msg   ="A valid arrest number must be entered on Booking Report.  If Booking Report has not yet been received, clear all Arrest fields until such time as paperwork is received."
'        GoTo exitedita
'    End If
'End If
On Error Resume Next

'RLB Code
    Dim dateBegin As Date, dateEnd As Date
    
    
    dateBegin = DateValue(incidentdate(0)) + TimeValue(TIMEOFOFFENSE(0))
    dateEnd = DateValue(incidentdate(1)) + TimeValue(TIMEOFOFFENSE(1))
    
    
    If dateBegin > dateEnd Then
        msg = "The date and time of the beginning of the offense, must take place before the date and time of the end of the offense."
'---- setfocus logic ----
'                 incidentdate(0).SetFocus
          If incidentdate(0).Visible Then
              incidentdate(0).SetFocus
           End If
        GoTo exitedita
    End If
'********

'===== Data Element 5
'===== Error 153
If IsDate(EXCEPTIONALCLEARANCEDATE) And Not (offenderdeath Or noprosecution Or extraditiondenied Or victimdeclinescooperation Or juvenilenocustody) Then
    msg = "The combination of an entered Exceptional Clearance Date and N/A is invalid."
    Call ShowApplicableContainers(EXCEPTIONALCLEARANCEDATE)
'---- setfocus logic ----
'             EXCEPTIONALCLEARANCEDATE.SetFocus
          If EXCEPTIONALCLEARANCEDATE.Visible Then
              EXCEPTIONALCLEARANCEDATE.SetFocus
           End If
    GoTo exitedita
End If
'===== Error 156
If Not IsDate(EXCEPTIONALCLEARANCEDATE) And (offenderdeath Or noprosecution Or extraditiondenied Or victimdeclinescooperation Or juvenilenocustody) Then
    msg = "An Exceptional Clearance Date must be entered for this reason."
    Call ShowApplicableContainers(EXCEPTIONALCLEARANCEDATE)
'---- setfocus logic ----
'             EXCEPTIONALCLEARANCEDATE.SetFocus
          If EXCEPTIONALCLEARANCEDATE.Visible Then
              EXCEPTIONALCLEARANCEDATE.SetFocus
           End If
    GoTo exitedita
End If
'===== Error 155
If IsDate(EXCEPTIONALCLEARANCEDATE) Then
    If CVDate(EXCEPTIONALCLEARANCEDATE) < CVDate(incidentdate(0)) Then
        msg = "Exceptional Clearance Date cannot be earlier than Incident Date."
        Call ShowApplicableContainers(EXCEPTIONALCLEARANCEDATE)
'---- setfocus logic ----
'                 EXCEPTIONALCLEARANCEDATE.SetFocus
          If EXCEPTIONALCLEARANCEDATE.Visible Then
              EXCEPTIONALCLEARANCEDATE.SetFocus
           End If
        GoTo exitedita
    End If
End If
'==== Mandatories E - 4 = GIVEN
'==== Mandatories E - 5
If (exclearunder18 = 1 Or exclearover18 = 1) And Not IsDate(EXCEPTIONALCLEARANCEDATE) Then
    msg = "Exceptional Clearance Date is not valid."
    Call ShowApplicableContainers(exclearunder18)
'---- setfocus logic ----
'             exclearunder18.SetFocus
          If exclearunder18.Visible Then
              exclearunder18.SetFocus
           End If
    GoTo exitedita
End If
'===== Error 105
If IsDate(EXCEPTIONALCLEARANCEDATE) Then
    EXCEPTIONALCLEARANCEDATE = Format$(EXCEPTIONALCLEARANCEDATE, "mm/dd/yyyy")
End If
If IsDate(EXCEPTIONALCLEARANCEDATE) Then
    If CDate(EXCEPTIONALCLEARANCEDATE) < CDate(incidentdate(1)) Then
        msg = "Exceptional Clearance Date cannot be prior to offense date."
        GoTo exitedita
    End If
End If
'===== Data Element 4
If Not (offenderdeath Or noprosecution Or extraditiondenied Or victimdeclinescooperation Or juvenilenocustody) And IsDate(EXCEPTIONALCLEARANCEDATE) Then
    msg = "A type of Exceptional Clearance other than NA must be chosen if date entered."
    Call ShowApplicableContainers(offenderdeath)
'---- setfocus logic ----
'             offenderdeath.SetFocus
          If offenderdeath.Visible Then
              offenderdeath.SetFocus
           End If
    GoTo exitedita
End If
'===== Data Element 5
If offenderdeath Or noprosecution Or extraditiondenied Or victimdeclines Or juvenilenocustody Then
    If Not IsDate(EXCEPTIONALCLEARANCEDATE) Then
        msg = "An exceptional clearance date must be entered if exceptionally cleared."
        Call ShowApplicableContainers(EXCEPTIONALCLEARANCEDATE)
'---- setfocus logic ----
'                 EXCEPTIONALCLEARANCEDATE.SetFocus
          If EXCEPTIONALCLEARANCEDATE.Visible Then
              EXCEPTIONALCLEARANCEDATE.SetFocus
           End If
        GoTo exitedita
    End If
End If
If offenderdeath Or noprosecution Or extraditiondenied Or victimdeclines Or juvenilenocustody Then
    Set rs = db.OpenRecordset("select incidentnumber from booking where INCIDENTnumber = " + Chr$(34) + incidentnumber + Chr$(34))
    If Not rs.EOF Then
        msg = "No arrest data allowed for an exceptional clearance."
        Call ShowApplicableContainers(EXCEPTIONALCLEARANCEDATE)
'---- setfocus logic ----
'                 EXCEPTIONALCLEARANCEDATE.SetFocus
          If EXCEPTIONALCLEARANCEDATE.Visible Then
              EXCEPTIONALCLEARANCEDATE.SetFocus
           End If
        GoTo exitedita
    End If
End If
'===== Mandatories E - 2
If incidentnumber = "" Or Len(incidentnumber) > 12 Then
    msg = "Incident number must be entered and be 12 or less characters long."
    Call ShowApplicableContainers(incidentnumber)
'---- setfocus logic ----
'             incidentnumber.SetFocus
          If incidentnumber.Visible Then
              incidentnumber.SetFocus
           End If
    GoTo exitedita
End If
'===== Error 172
If CVDate(incidentdate(0)) < CVDate("1/1/1991") And onpaper = 0 Then
    msg = "Incident date is not valid."
    Call ShowApplicableContainers(incidentdate(0))
'---- setfocus logic ----
'             incidentdate(0).SetFocus
          If incidentdate(0).Visible Then
              incidentdate(0).SetFocus
           End If
    GoTo exitedita
End If
'===== Error 105
incidentdate(0) = Format$(incidentdate(0), "mm/dd/yyyy")
'===== Mandatories E - 3
'===== Data Element 3
'===== Error 151
If Not IsDate(TIMEOFOFFENSE(0)) Then
    msg = "Time of Offense is not valid."
    Call ShowApplicableContainers(TIMEOFOFFENSE(0))
'---- setfocus logic ----
'             TIMEOFOFFENSE(0).SetFocus
          If TIMEOFOFFENSE(0).Visible Then
              TIMEOFOFFENSE(0).SetFocus
           End If
    GoTo exitedita
End If
'===== SC Enhancements - Administrative
If Not IsDate(TIMEOFOFFENSE(1)) Then
    msg = "Time of Offense is not valid."
    Call ShowApplicableContainers(TIMEOFOFFENSE(1))
'---- setfocus logic ----
'             TIMEOFOFFENSE(1).SetFocus
          If TIMEOFOFFENSE(1).Visible Then
              TIMEOFOFFENSE(1).SetFocus
           End If
    GoTo exitedita
End If
If IsDate(REPORTINGOFFICERDATE(0)) Then
    REPORTINGOFFICERDATE(0) = Format$(REPORTINGOFFICERDATE(0), "mm/dd/yyyy")
End If
If reportingofficer(1) > "" Then
    If Not IsDate(REPORTINGOFFICERDATE(1)) Then
        REPORTINGOFFICERDATE(1) = REPORTINGOFFICERDATE(0)
    End If
    REPORTINGOFFICERDATE(1) = Format$(REPORTINGOFFICERDATE(1), "mm/dd/yyyy")
End If
If APPROVINGOFFICERDATE > "" And Not IsDate(APPROVINGOFFICERDATE) Then
    msg = "Approving Date is not valid."
    Call ShowApplicableContainers(APPROVINGOFFICERDATE)
'---- setfocus logic ----
'             APPROVINGOFFICERDATE.SetFocus
          If APPROVINGOFFICERDATE.Visible Then
              APPROVINGOFFICERDATE.SetFocus
           End If
    GoTo exitedita
End If
'===== SCEnhancements - Administrative
If active = 0 And admclosed = 0 And unfounded = 0 Then
    msg = "Either Active, Adm Closed, or Unfounded must be selected."
    Call ShowApplicableContainers(active)
'---- setfocus logic ----
'             active.SetFocus
          If active.Visible Then
              active.SetFocus
           End If
    GoTo exitedita
End If
'==== Mandatories E - 8A
If BIAS.ListIndex = -1 Then
    msg = "Bias Motivation must be selected."
    Call ShowApplicableContainers(BIAS)
'---- setfocus logic ----
'             BIAS.SetFocus
          If BIAS.Visible Then
              BIAS.SetFocus
           End If
    GoTo exitedita
End If
GoTo goodedita
Exit Sub
exitedita:
editerr = 1
goodedita:
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub
Friend Sub editevent(editerr, msg As String)
Dim testgroup, totgroup As String, temperr As Integer, tempucr, tempgroup, typeselect As String, tempvalue As Single, tempdate As String
'RLB Bandaid
On Error GoTo rlbErre
Screen.MousePointer = 11
'===== Error 101
If Not IsDate(incidentdate(0)) Then
    msg = "Beginning incident date is not valid."
    Call ShowApplicableContainers(incidentdate(0))
'---- setfocus logic ----
'             incidentdate(0).SetFocus
          If incidentdate(0).Visible Then
              incidentdate(0).SetFocus
           End If
    GoTo exitedite
End If
'===== SC Enhancements - Administrative
If Not IsDate(incidentdate(1)) Then
    msg = "Ending incident date is not valid."
    Call ShowApplicableContainers(incidentdate(1))
'---- setfocus logic ----
'             incidentdate(1).SetFocus
          If incidentdate(1).Visible Then
              incidentdate(1).SetFocus
           End If
    GoTo exitedite
End If
If CVDate(incidentdate(0)) > CVDate(Date$) Then
    msg = "Incident date cannot be in the future."
    Call ShowApplicableContainers(incidentdate(0))
'---- setfocus logic ----
'             incidentdate(0).SetFocus
          If incidentdate(0).Visible Then
              incidentdate(0).SetFocus
           End If
    GoTo exitedite
End If
If CVDate(incidentdate(1)) > CVDate(Date$) Then
    msg = "Incident date cannot be in the future."
    Call ShowApplicableContainers(incidentdate(1))
'---- setfocus logic ----
'             incidentdate(1).SetFocus
          If incidentdate(1).Visible Then
              incidentdate(1).SetFocus
           End If
    GoTo exitedite
End If

'===== Error 016, 116, 216, 316, 416, 516, 616, 716
If Len(incidentnumber) < 12 And Len(incidentnumber) > 0 Then
    incidentnumber = incidentnumber + Space$(12 - Len(incidentnumber))
End If
If age(0) > "" Then
    If Val(age(0)) = 0 And age(0) <> "00" Then
        msg = "Age must be entered in the format of a single age: X or XX (i.e. 3 or 42) or in the format of a range of ages: XXXX (i.e. 2025, meaning 20 - 25 years old)."
        Call ShowApplicableContainers(age(0))
'---- setfocus logic ----
'                 age(0).SetFocus
          If age(0).Visible Then
              age(0).SetFocus
           End If
        GoTo exitedite
    End If
End If
For t% = 1 To Len(age(0))
    If InStr("0123456789-", Mid$(age(0), t%, 1)) = 0 Then
        msg = "An invalid entry in age has been found.  Valid Entry Formats:  X, XX, X-Y, XX-Y, X-YY, XX-YY, XXYY"
        t% = Len(age(0))
        Call ShowApplicableContainers(age(0))
'---- setfocus logic ----
'                 age(0).SetFocus
          If age(0).Visible Then
              age(0).SetFocus
           End If
        GoTo exitedite
    End If
Next t%
If InStr(age(0), "-") > 0 Then
    ag1 = Val(Left$(age(0), InStr(age(0), "-") - 1))
    ag2 = Val(Mid$(age(0), InStr(age(0), "-") + 1))
    age(0) = Format$(ag1, "00") + Format$(ag2, "00")
End If
editeventdetail:
'===== SC Edit - Justifiable Homocide (09C) must be on separate casenumber =====
fo% = 0
fj% = 0
nolarceny = 0
tempucr = ""
ucrct% = 0
found35a = 0
ucrs$ = ""
foundweapon = 0
FOUNDLEOKA = 0
For t% = 0 To 4
    For tt% = 0 To UCRLIST(t%).ListCount - 1
        If UCRLIST(t%).Selected(tt%) = True Then
            ucrct% = ucrct% + 1
            tempucr = Mid$(UCRLIST(t%).List(tt%), InStr(UCRLIST(t%).List(tt%), "(") + 1, 3)
            For p% = 1 To Len(ucrs$) Step 3
                If Mid$(ucrs$, p%, 3) = tempucr Then
                    msg = "A duplicate UCR reference for " + tempucr + " has been found.  Do not enter separate offense lines for the same UCR."
                    Call ShowApplicableContainers(UCRLIST(t%))
'---- setfocus logic ----
'                             UCRLIST(t%).SetFocus
          If UCRLIST(t%).Visible Then
              UCRLIST(t%).SetFocus
           End If
                     
                    GoTo exitedite
                End If
            Next p%
            ucrs$ = ucrs$ + tempucr
            Select Case tempucr
                Case "13A", "13B", "09A"
                    FOUNDLEOKA = 1
                Case "09A", "09B", "09C", "100", "11A", "11B", "11C", "11D", "120", "13A", "13B", "210", "520"
                    foundweapon = 1
                Case "09C"
                    fj% = 1
                    nolarceny = 1
                Case "23A", "23B", "23C", "23D", "23E", "23F", "23G", "23H"
                    fo% = 1
                Case "35A"
                    found35a = 1
                    fo% = 1
                    nolarceny = 1
                Case Else
                    nolarceny = 1
                    fo% = 1
            End Select
            tt% = UCRLIST(t%).ListCount - 1
        End If
    Next tt%
Next t%
'===== Error 201
'===== SCEdit 4/21/92 P28
If tempucr = "" Then
    msg = "A valid UCR code must be selected."
    Call ShowApplicableContainers(UCRLIST(0))
'---- setfocus logic ----
'             UCRLIST(0).SetFocus
          If UCRLIST(0).Visible Then
              UCRLIST(0).SetFocus
           End If
    GoTo exitedite
End If
If fo% = 1 And fj% = 1 Then
    msg = "Justifiable Homocide must be submitted on a separate case number."
    Call ShowApplicableContainers(UCRLIST(0))
'---- setfocus logic ----
'             UCRLIST(0).SetFocus
          If UCRLIST(0).Visible Then
              UCRLIST(0).SetFocus
           End If
    GoTo exitedite
End If
'===== SC LEOKA
'--- p 11-21 - User Manual
If policeofficer Or lactivity.ListIndex > -1 Then
    If FOUNDLEOKA = 0 Then
        msg = "LEOKA designation only valid for Aggravated Assault, Simple Assault, or Murder/Non-negligent Manslaughter committed against a police officer."
        Call ShowApplicableContainers(policeofficer)
'---- setfocus logic ----
'                 policeofficer.SetFocus
          If policeofficer.Visible Then
              policeofficer.SetFocus
           End If
        GoTo exitedite
    End If
End If
FSP% = 0
For t% = 0 To 4
    For tt% = 1 To sublist(t%).ListItems.Count
        If sublist(t%).ListItems(tt%).Selected Then
            If Left$(sublist(t%).ListItems(tt%), 1) = "P" Then
                FSP% = 1
                tt% = sublist(t%).ListItems.Count
                t% = 4
            End If
        End If
    Next tt%
Next t%
If lactivity.ListIndex = -1 Then
    If policeofficer Then
        msg = "A LEOKA activity must be selected for Type Victim Police Officer."
        Call ShowApplicableContainers(policeofficer)
'---- setfocus logic ----
'                 policeofficer.SetFocus
          If policeofficer.Visible Then
              policeofficer.SetFocus
           End If
        GoTo exitedite
    End If
    If FSP% = 1 Then
        msg = "A LEOKA activity must be selected for Type Victim Police Officer."
        Call ShowApplicableContainers(sublist(0))
'---- setfocus logic ----
'                 sublist(0).SetFocus
          If sublist(0).Visible Then
              sublist(0).SetFocus
           End If
        GoTo exitedite
    End If
Else
    If Not policeofficer Then
        msg = "If an Activity of selected for LEOKA, the type victim must be police officer."
        Call ShowApplicableContainers(policeofficer)
'---- setfocus logic ----
'                 policeofficer.SetFocus
          If policeofficer.Visible Then
              policeofficer.SetFocus
           End If
        GoTo exitedite
    End If
End If
If FSP% = 0 And policeofficer Then
    msg = "Type Victim is Police Officer, but Police Officer is not indicated in Subcodes."
    Call ShowApplicableContainers(sublist(0))
'---- setfocus logic ----
'             sublist(0).SetFocus
          If sublist(0).Visible Then
              sublist(0).SetFocus
           End If
    GoTo exitedite
End If
If FSP% = 1 And Not policeofficer Then
    msg = "Type Victim must be Police Officer if indicated in Subcodes."
    Call ShowApplicableContainers(policeofficer)
'---- setfocus logic ----
'             policeofficer.SetFocus
          If policeofficer.Visible Then
              policeofficer.SetFocus
           End If
    GoTo exitedite
End If
For t% = 0 To 4
    '===== Mandatories E - 6, 24
    '===== Data Element 6
    UCRLIST(t%).ListIndex = -1
    For tt% = 0 To UCRLIST(t%).ListCount - 1
        If UCRLIST(t%).Selected(tt%) = True Then
            UCRLIST(t%).ListIndex = tt%
            tt% = UCRLIST(t%).ListCount - 1
        End If
    Next tt%
    If pickoffense(t%).ListIndex <> -1 And UCRLIST(t%).ListIndex = -1 Then
        msg = "A valid UCR code must be selected."
        Call ShowApplicableContainers(UCRLIST(t%))
'---- setfocus logic ----
'                 UCRLIST(t%).SetFocus
          If UCRLIST(t%).Visible Then
              UCRLIST(t%).SetFocus
           End If
        GoTo exitedite
    End If
    If InStr(incidentnumber, "-458") > 0 Then
        abc123 = 1
    End If
    If UCRLIST(t%).ListIndex > -1 Then
        tempucr = Mid$(UCRLIST(t%).List(UCRLIST(t%).ListIndex), InStr(UCRLIST(t%).List(UCRLIST(t%).ListIndex), "(") + 1, 3)
        '===== SC LEOKA
        If FSP% = 1 Or policeofficer Then
            For ii% = 1 To HOMOCIDE(t%).ListItems.Count
                If HOMOCIDE(t%).ListItems(ii%).Selected Then
                    If Mid$(HOMOCIDE(t%).ListItems(ii%), InStr(HOMOCIDE(t%).ListItems(ii%), "(") + 1, 2) <> "02" Then
                        msg = "For LEOKA, only homocide/aggravated assault circumstance of 02 is allowed."
                        Call ShowApplicableContainers(HOMOCIDE(t%))
'---- setfocus logic ----
'                                 HOMOCIDE(t%).SetFocus
          If HOMOCIDE(t%).Visible Then
              HOMOCIDE(t%).SetFocus
           End If
                        GoTo exitedite
                    End If
                End If
            Next ii%
        Else
            If Not HOMOCIDE(t%).SelectedItem Is Nothing Then
                For ii% = 1 To HOMOCIDE(t%).ListItems.Count
                    If HOMOCIDE(t%).ListItems(ii%).Selected Then
                        If Mid$(HOMOCIDE(t%).ListItems(ii%), InStr(HOMOCIDE(t%).ListItems(ii%), "(") + 1, 2) = "02" Then
                            msg = "For Homocide Circumstance 02, LEOKA information must be entered."
                            Call ShowApplicableContainers(HOMOCIDE(t%))
'---- setfocus logic ----
'                                     HOMOCIDE(t%).SetFocus
          If HOMOCIDE(t%).Visible Then
              HOMOCIDE(t%).SetFocus
           End If
                            GoTo exitedite
                        End If
                    End If
                Next ii%
            End If
        End If
        '===== Error 469
        If tempucr = "11A" Or tempucr = "36B" Then
            If sex(1).ListIndex = -1 Or Left$(sex(1).List(sex(1).ListIndex), 1) <> "M" And Left$(sex(1).List(sex(1).ListIndex), 1) <> "F" Then
                msg = "For Forcible Rape, a victim sex of M or F are allowed."
                Call ShowApplicableContainers(sex(1))
'---- setfocus logic ----
'                         sex(1).SetFocus
          If sex(1).Visible Then
              sex(1).SetFocus
           End If
                GoTo exitedite
            End If
        End If
        If tempucr = "250" Or tempucr = "280" Or tempucr = "35A" Or tempucr = "35B" Or tempucr = "39C" Or tempucr = "370" Or tempucr = "520" Or tempucr = "09A" Or tempucr = "09B" Or tempucr = "100" Or tempucr = "120" Or tempucr = "11A" Or tempucr = "11B" Or tempucr = "11C" Or tempucr = "11D" Or tempucr = "13A" Or tempucr = "13B" Or tempucr = "13C" Then
            If tempucr = "09A" Or tempucr = "09B" Or tempucr = "100" Or tempucr = "120" Or tempucr = "11A" Or tempucr = "11B" Or tempucr = "11C" Or tempucr = "11D" Or tempucr = "13A" Or tempucr = "13B" Or tempucr = "13C" Then
                foundactivity = 0
                For i% = 1 To gactivity(t%).ListItems.Count
                    If gactivity(t%).ListItems(i%).Selected = True Then
                        gactivity(t%).ListItems(i%).EnsureVisible
                        foundactivity = 1
                        i% = gactivity(t%).ListItems.Count
                    End If
                Next i%
                If foundactivity = 0 Then
                    msg = "A valid activity type must be selected for UCR " + tempucr + "."
                    Call ShowApplicableContainers(gactivity(t%))
'---- setfocus logic ----
'                             gactivity(t%).SetFocus
          If gactivity(t%).Visible Then
              gactivity(t%).SetFocus
           End If
                    GoTo exitedite
                End If
            Else
                '===== Error 201
                foundactivity = 0
                For i% = 1 To activity(t%).ListItems.Count
                    If activity(t%).ListItems(i%).Selected = True Then
                        activity(t%).ListItems(i%).EnsureVisible
                        foundactivity = 1
                        i% = activity(t%).ListItems.Count
                    End If
                Next i%
                If foundactivity = 0 Then
                    msg = "A valid activity type must be selected for UCR " + tempucr + "."
                    Call ShowApplicableContainers(activity(t%))
'---- setfocus logic ----
'                             activity(t%).SetFocus
          If activity(t%).Visible Then
              activity(t%).SetFocus
           End If
                    GoTo exitedite
                End If
            End If
        End If
        foundweapon = 0
        foundweapon99 = 0
        idx5 = 0
        For i% = 1 To weapontype.ListItems.Count
            If weapontype.ListItems(i%).Selected = True Then
                weapontype.ListItems(i%).EnsureVisible
                idx5 = idx5 + 1
                If idx5 > 3 Then
                    i% = weapontype.ListItems.Count
                Else
                    foundweapon = foundweapon + 1
                    If InStr(weapontype.ListItems(i%), "(99)") > 0 Then
                        foundweapon99 = 1
                    End If
                    '----edit by Ed Sloan 11/30/99 from SCLED 3/5/97 update----------------
                    '===== Error 265, 269
                    If tempucr = "13B" Then
                        If (Mid$(weapontype.ListItems(i%), InStr(weapontype.ListItems(i%), "(") + 1, 2) <> "40" And Mid$(weapontype.ListItems(i%), InStr(weapontype.ListItems(i%), "(") + 1, 2) <> "95" And Mid$(weapontype.ListItems(i%), InStr(weapontype.ListItems(i%), "(") + 1, 2) <> "99" And Mid$(weapontype.ListItems(i%), InStr(weapontype.ListItems(i%), "(") + 1, 2) <> "90") Then
                            msg = "For simple assault, weapon codes must be 40, 90, 95 or 99"
                            weapontype.ListItems(i%).Selected = False
                            Call ShowApplicableContainers(weapontype)
'---- setfocus logic ----
'                                     weapontype.SetFocus
          If weapontype.Visible Then
              weapontype.SetFocus
           End If
                            GoTo exitedite
                        End If
                    End If
                    '===== Data Element 13
                    If Mid$(weapontype.ListItems(i%), InStr(weapontype.ListItems(i%), "(") + 1, 2) <> "11" And _
                        Mid$(weapontype.ListItems(i%), InStr(weapontype.ListItems(i%), "(") + 1, 2) <> "12" And _
                        Mid$(weapontype.ListItems(i%), InStr(weapontype.ListItems(i%), "(") + 1, 2) <> "13" And _
                        Mid$(weapontype.ListItems(i%), InStr(weapontype.ListItems(i%), "(") + 1, 2) <> "14" And _
                        Mid$(weapontype.ListItems(i%), InStr(weapontype.ListItems(i%), "(") + 1, 2) <> "15" Then
                            If automatic(i%) <> "N" Then
                                msg = "Weapon Type/Automatic combination invalid."
                                Call ShowApplicableContainers(weapontype)
'---- setfocus logic ----
'                                         weapontype.SetFocus
          If weapontype.Visible Then
              weapontype.SetFocus
           End If
                                GoTo exitedite
                            End If
                    End If
                End If
                End If
        Next i%
        '===== Error 201
        '===== SCEdit 8/9/95
        Select Case tempucr
            Case "09A", "09B", "09C", "100", "11A", "11B", "11C", "11D", "120", "13A", "13B", "210", "520", "980"
                If foundweapon = 0 Then
                    msg = "A valid weapon type must be selected."
                    Call ShowApplicableContainers(weapontype)
'---- setfocus logic ----
'                             weapontype.SetFocus
          If weapontype.Visible Then
              weapontype.SetFocus
           End If
                    GoTo exitedite
                End If
        End Select
        If foundweapon > 0 Then
            Select Case tempucr
                Case "09A", "09B", "09C", "100", "11A", "11B", "11C", "11D", "120", "13A", "13B", "210", "520"
                '===== Error 219
                Case Else
                    If Left$(tempucr, 2) <> "90" Then
                        foundother% = 0
                        For q% = 0 To 4
                            If UCRLIST(q%).ListIndex > -1 Then
                                Select Case Mid$(UCRLIST(q%).List(UCRLIST(q%).ListIndex), InStr(UCRLIST(q%).List(UCRLIST(q%).ListIndex), "(") + 1, 3)
                                    Case "09A", "09B", "09C", "100", "11A", "11B", "11C", "11D", "120", "13A", "13B", "210", "520"
                                        foundother% = 1
                                        q% = 9
                                End Select
                            End If
                        Next q%
                        If foundother% = 0 Then
                            msg = "A weapon cannot be selected with this UCR: " + tempucr
                            Call ShowApplicableContainers(weapontype)
'---- setfocus logic ----
'                                     weapontype.SetFocus
          If weapontype.Visible Then
              weapontype.SetFocus
           End If
                            GoTo exitedite
                        End If
                    End If
            End Select
        End If
        '===== Error 267
        If tempucr = "09A" Or tempucr = "09B" Or tempucr = "09C" Then
            If foundweapon99 > 0 Then
                msg = "A weapon type other than 99 = None must be selected for this UCR " + tempucr + "."
                Call ShowApplicableContainers(weapontype)
'---- setfocus logic ----
'                         weapontype.SetFocus
          If weapontype.Visible Then
              weapontype.SetFocus
           End If
                GoTo exitedite
            End If
        End If
        '===== Data Element 13
        '===== Error 207
        If foundweapon > 1 And foundweapon99 = 1 Then
            msg = "If Weapon Type NONE is chosen, no other weapon types can be selected as well."
            Call ShowApplicableContainers(weapontype)
'---- setfocus logic ----
'                     weapontype.SetFocus
          If weapontype.Visible Then
              weapontype.SetFocus
           End If
            GoTo exitedite
        End If
        '===== SCEdit 8/9/95
        ctpremise% = 0
        For ZZ% = 1 To premise(t%).ListItems.Count
            If premise(t%).ListItems(ZZ%).Selected Then
                ctpremise% = ctpremise% + 1
                temppremise = Mid$(premise(t%).ListItems(ZZ%), InStr(premise(t%).ListItems(ZZ%), "(") + 1, 2)
                If tempucr = "220" Then
                    If Not (temppremise = "01" Or temppremise = "02" Or temppremise = "03" Or temppremise = "04" Or temppremise = "05" Or temppremise = "06" Or temppremise = "07" Or temppremise = "08" Or temppremise = "09" Or temppremise = "11" Or temppremise = "12" Or temppremise = "14" Or temppremise = "15" Or temppremise = "17" Or temppremise = "19" Or temppremise = "20" Or temppremise = "21" Or temppremise = "22" Or temppremise = "23" Or temppremise = "24" Or temppremise = "25" Or temppremise = "26" Or temppremise = "27" Or temppremise = "28") Then
                        msg = "For breaking and entering, premise code must be one of 01, 02, 03, 04, 05, 06, 07, 08, 09, 11, 12, 14, 15, 17, 19, 20, 21, 22, 23, 24, 25, 26, 27 or 28."
                        Call ShowApplicableContainers(premise(t%))
'---- setfocus logic ----
'                                 premise(t%).SetFocus
          If premise(t%).Visible Then
              premise(t%).SetFocus
           End If
                        GoTo exitedite
                    End If
                End If
                If tempucr = "23C" Then
                    If Not (temppremise = "01" Or temppremise = "03" Or temppremise = "04" Or temppremise = "05" Or temppremise = "07" Or temppremise = "08" Or temppremise = "11" Or temppremise = "12" Or temppremise = "14" Or temppremise = "17" Or temppremise = "21" Or temppremise = "22" Or temppremise = "23" Or temppremise = "24" Or temppremise = "25" Or temppremise = "26" Or temppremise = "27") Then
                        msg = "For shoplifting, premise code must be one of 01, 03, 04, 05, 07, 08, 09, 11, 12, 14, 17, 21, 22, 23, 24, 25, 26 or 27."
                        Call ShowApplicableContainers(premise(t%))
'---- setfocus logic ----
'                                 premise(t%).SetFocus
          If premise(t%).Visible Then
              premise(t%).SetFocus
           End If
                        GoTo exitedite
                    End If
                End If
            End If
        Next ZZ%
        '===== SCEdit 1/31/1
        If ctpremise% = 1 Then
            If temppremise = "18" Then
                msg = "If Premise Type 18 (Parking Lot/Parking Garage) is selected, then another premise type must also be selected to further describe it."
                Call ShowApplicableContainers(premise(t%))
'---- setfocus logic ----
'                         premise(t%).SetFocus
          If premise(t%).Visible Then
              premise(t%).SetFocus
           End If
                GoTo exitedite
            End If
        End If
        foundother% = 0
        If tempucr = "09C" Then
            For q% = 0 To t% - 1
                If UCRLIST(q%).ListIndex > -1 Then
                    foundother% = 1
                    q% = t% - 1
                End If
            Next q%
            If foundother% = 0 Then
                For q% = t% + 1 To 9
                    If UCRLIST(q%).ListIndex > -1 Then
                        foundother% = 1
                        q% = 9
                    End If
                Next q%
            End If
            If foundother% = 1 Then
                msg = "For Justifiable Homocide there can be no other offenses."
                Call ShowApplicableContainers(UCRLIST(0))
'---- setfocus logic ----
'                         UCRLIST(0).SetFocus
          If UCRLIST(0).Visible Then
              UCRLIST(0).SetFocus
           End If
                GoTo exitedite
            End If
        End If
        '===== SCEdit 8/9/95
        Select Case tempucr
            Case "09C", "979", "992", "980", "978"
                If IsDate(EXCEPTIONALCLEARANCEDATE) Then
                    msg = "Exceptional clearance is not allowed on a " + tempucr + " offense."
                    Call ShowApplicableContainers(UCRLIST(0))
'---- setfocus logic ----
'                             UCRLIST(0).SetFocus
          If UCRLIST(0).Visible Then
              UCRLIST(0).SetFocus
           End If
                    GoTo exitedite
                End If
        End Select
        '-------------------------------------
        '===== Error 257
        If tempucr = "220" Then
            For ZZ% = 1 To premise(t%).ListItems.Count
                If premise(t%).ListItems(ZZ%).Selected Then
                    If Mid$(premise(t%).ListItems(ZZ%), InStr(premise(t%).ListItems(ZZ%), "(") + 1, 2) = "14" Or Mid$(premise(t%).ListItems(ZZ%), InStr(premise(t%).ListItems(ZZ%), "(") + 1, 2) = "19" Then
                        If Val(entered(t%)) < 1 Or Val(entered(t%)) > 99 Then
                            msg = "Number of premises entered must be entered (01-99)."
                            Call ShowApplicableContainers(premise(t%))
'---- setfocus logic ----
'                                     premise(t%).SetFocus
          If premise(t%).Visible Then
              premise(t%).SetFocus
           End If
                            GoTo exitedite
                        End If
                    End If
                End If
            Next ZZ%
        End If
        '===== Error 252
        If Val(entered(t%)) > 0 Then
            If tempucr <> "220" Then
                msg = "If Number of Premises Entered is valued, the UCR code must be 220 (Burglary/B&E)."
                Call ShowApplicableContainers(UCRLIST(0))
'---- setfocus logic ----
'                         UCRLIST(0).SetFocus
          If UCRLIST(0).Visible Then
              UCRLIST(0).SetFocus
           End If
                GoTo exitedite
            End If
            found1419% = 0
            For ZZ% = 1 To premise(t%).ListItems.Count
                If premise(t%).ListItems(ZZ%).Selected Then
                    If Mid$(premise(t%).ListItems(ZZ%), InStr(premise(t%).ListItems(ZZ%), "(") + 1, 2) = "14" Or Mid$(premise(t%).ListItems(ZZ%), InStr(premise(t%).ListItems(ZZ%), "(") + 1, 2) = "19" Then
                        found1419% = 1
                        ZZ% = premise(t%).ListItems.Count
                    End If
                End If
            Next ZZ%
            If found1419% = 0 Then
                msg = "If Number of Premises Entered is valued, the Premise Type must be 14 (Hotel/Motel) or 19 (Rental/Storage Facility)."
                Call ShowApplicableContainers(premise(t%))
'---- setfocus logic ----
'                         premise(t%).SetFocus
          If premise(t%).Visible Then
              premise(t%).SetFocus
           End If
                GoTo exitedite
            End If
        End If
        '===== Mandatories E - 7
        '===== Data Element 7
        '===== Error 201
        If Not completedy(t%) And Not completedn(t%) Then
            msg = "A valid completed code of Yes or No must be selected in COMPLETED."
            Call ShowApplicableContainers(completedy(t%))
'---- setfocus logic ----
'                     completedy(t%).SetFocus
          If completedy(t%).Visible Then
              completedy(t%).SetFocus
           End If
            GoTo exitedite
        End If
        '=== Additional F 2
        '===== Data Element 7
        '===== Error 256
        If tempucr = "13A" Or tempucr = "13B" Or tempucr = "13C" Then
            If Not completedy(t%) Then
                msg = "Crimes Against Persons must show Completed."
                Call ShowApplicableContainers(completedy(t%))
'---- setfocus logic ----
'                         completedy(t%).SetFocus
          If completedy(t%).Visible Then
              completedy(t%).SetFocus
           End If
                GoTo exitedite
            End If
            If Not individual And Not policeofficer Then
                msg = "Crimes Against Persons must be associated with an Individual or Police Officer."
                Call ShowApplicableContainers(individual)
'---- setfocus logic ----
'                         individual.SetFocus
          If individual.Visible Then
              individual.SetFocus
           End If
                GoTo exitedite
            End If
        End If
        If tempucr = "13A" Or tempucr = "13B" Then
            selw% = 0
            For iop% = 1 To weapontype.ListItems.Count
                If weapontype.ListItems(iop%).Selected Then
                    selw% = 1
                    iop% = weapontype.ListItems.Count
                End If
            Next iop%
            If selw% = 0 Then
                msg = "A weapon type must be selected for this UCR " + tempucr + "."
                Call ShowApplicableContainers(weapontype)
'---- setfocus logic ----
'                         weapontype.SetFocus
          If weapontype.Visible Then
              weapontype.SetFocus
           End If
                GoTo exitedite
            End If
            '---MOVED TO EXPORT ---
            'If injury.SelectedItem Is Nothing Then
            '    msg = "An injury type must be selected for this UCR " + tempucr + "."
            '    GoTo exitedite
            'End If
        End If
        '===== Error 462
        If tempucr = "13A" Then
            selhom% = 0
            For xx% = 1 To HOMOCIDE(t%).ListItems.Count
                If HOMOCIDE(t%).ListItems(xx%).Selected Then
                    selhom% = 1
                    If Mid$(HOMOCIDE(t%).ListItems(xx%), InStr(HOMOCIDE(t%).ListItems(xx%), "(") + 1, 2) = "07" Then
                        msg = "Reason 07 = Mercy Killing not allowed for UCR 13A."
                        Call ShowApplicableContainers(HOMOCIDE(t%))
'---- setfocus logic ----
'                                 HOMOCIDE(t%).SetFocus
          If HOMOCIDE(t%).Visible Then
              HOMOCIDE(t%).SetFocus
           End If
                        GoTo exitedite
                    End If
                End If
            Next xx%
            If selhom% = 0 Then
                msg = "Aggravated Assault/Homocide Circumstances must be selected for Aggravated Assault."
                Call ShowApplicableContainers(HOMOCIDE(t%))
'---- setfocus logic ----
'                         HOMOCIDE(t%).SetFocus
          If HOMOCIDE(t%).Visible Then
              HOMOCIDE(t%).SetFocus
           End If
                GoTo exitedite
            End If
        End If
        '===== Data Element 13
        If tempucr = "13B" Then
            For TU% = 1 To weapontype.ListItems.Count
                If weapontype.ListItems(TU%).Selected Then
                    TESTWEAPON = Mid$(weapontype.ListItems(TU%), InStr(weapontype.ListItems(TU%), "(") + 1, 2)
                    If TESTWEAPON <> "40" And TESTWEAPON <> "90" And TESTWEAPON <> "95" And TESTWEAPON <> "99" And TESTWEAPON <> "  " Then
                        msg = "Weapon type invalid for simple assault."
                        Call ShowApplicableContainers(weapontype)
'---- setfocus logic ----
'                                 weapontype.SetFocus
          If weapontype.Visible Then
              weapontype.SetFocus
           End If
                        GoTo exitedite
                    End If
                End If
            Next TU%
        End If
            
        
        '===== Additional F 4
        '===== Data Element 10
        '===== Data Element 11
        '===== Error 204,253,254
        '===== SCEdit 7.2 Amendment P7-4
        If tempucr = "220" Then
            If societypublic Then
                msg = "Burglary/B&E cannot be Society/Public."
                Call ShowApplicableContainers(societypublic)
'---- setfocus logic ----
'                         societypublic.SetFocus
          If societypublic.Visible Then
              societypublic.SetFocus
           End If
                GoTo exitedite
            End If
        End If
        If tempucr = "220" Or tempucr = "23F" Or tempucr = "23G" Or tempucr = "240" Then
            If FORCEDENTRYY(t%) = 0 And FORCEDENTRYN(t%) = 0 Then
                msg = "A value of Y or N must be entered for Forced Entry."
                Call ShowApplicableContainers(FORCEDENTRYY(t%))
'---- setfocus logic ----
'                         FORCEDENTRYY(t%).SetFocus
          If FORCEDENTRYY(t%).Visible Then
              FORCEDENTRYY(t%).SetFocus
           End If
                GoTo exitedite
            End If
        Else
            If FORCEDENTRYY(t%) = 1 Or FORCEDENTRYN(t%) = 1 Then
                msg = "A value of Y or N CANNOT be entered for Forced Entry."
                Call ShowApplicableContainers(FORCEDENTRYY(t%))
'---- setfocus logic ----
'                         FORCEDENTRYY(t%).SetFocus
          If FORCEDENTRYY(t%).Visible Then
              FORCEDENTRYY(t%).SetFocus
           End If
                GoTo exitedite
            End If
        End If

        

        '===Additional F 5
        '===== Data Element 12
        If tempucr = "250" Then
            sela% = 0
            For iop% = 1 To activity(t%).ListItems.Count
                If activity(t%).ListItems(iop%).Selected Then
                    sela% = 1
                    iop% = activity(t%).ListItems.Count
                End If
            Next iop%
            If sela% = 0 Then
                msg = "Type of Activity must be selected for Counterfeiting."
                Call ShowApplicableContainers(activity(t%))
'---- setfocus logic ----
'                         activity(t%).SetFocus
          If activity(t%).Visible Then
              activity(t%).SetFocus
           End If
                GoTo exitedite
            End If
        End If
            
        '===Additional F 7
        '===== Data Element 7
        '===== Data Element 12
        If tempucr = "35A" Or tempucr = "35B" Then
            sela% = 0
            For iop% = 1 To activity(t%).ListItems.Count
                If activity(t%).ListItems(iop%).Selected Then
                    sela% = 1
                    iop% = activity(t%).ListItems.Count
                End If
            Next iop%
            If sela% = 0 Then
                msg = "Type of Activity must be selected for Drug/Narcotic Offenses."
                Call ShowApplicableContainers(activity(t%))
'---- setfocus logic ----
'                         activity(t%).SetFocus
          If activity(t%).Visible Then
              activity(t%).SetFocus
           End If
                GoTo exitedite
            End If
            If Not societypublic Then
                msg = "Drug crimes victims should be Society/Public."
                Call ShowApplicableContainers(societypublic)
'---- setfocus logic ----
'                         societypublic.SetFocus
          If societypublic.Visible Then
              societypublic.SetFocus
           End If
                GoTo exitedite
            End If
        End If

        
        '===== Additional F 9
        If tempucr = "210" Then
            sela% = 0
            For iop% = 1 To weapontype.ListItems.Count
                If weapontype.ListItems(iop%).Selected Then
                    sela% = 1
                    iop% = weapontype.ListItems.Count
                End If
            Next iop%
            If sela% = 0 Then
                msg = "A weapon type must be selected for this UCR " + tempucr + "."
                Call ShowApplicableContainers(weapontype)
'---- setfocus logic ----
'                         weapontype.SetFocus
          If weapontype.Visible Then
              weapontype.SetFocus
           End If
                GoTo exitedite
            End If
            'If individual Then
            '    If injury.SelectedItem Is Nothing Then
            '        msg = "An injury type must be selected for this UCR " + tempucr + "."
            '        GoTo exitedite
            '    End If
            'End If
        End If
        
        '===== Additional F 11
        '===== Data Element 7
        If tempucr = "39A" Or tempucr = "39B" Or tempucr = "39C" Or tempucr = "39D" Then
            If Not societypublic Then
                msg = "Drug crimes victims should be Society/Public."
                Call ShowApplicableContainers(societypublic)
'---- setfocus logic ----
'                         societypublic.SetFocus
          If societypublic.Visible Then
              societypublic.SetFocus
           End If
                GoTo exitedite
            End If
            '===== Data Element 12
            If tempucr = "39C" Then
                sela% = 0
                For iop% = 1 To activity(t%).ListItems.Count
                    If activity(t%).ListItems(iop%).Selected Then
                        sela% = 1
                        iop% = activity(t%).ListItems.Count
                    End If
                Next iop%
                If sela% = 0 Then
                    msg = "A type of activity must be entered for UCR " + tempucr + "."
                    Call ShowApplicableContainers(activity(t%))
'---- setfocus logic ----
'                             activity(t%).SetFocus
          If activity(t%).Visible Then
              activity(t%).SetFocus
           End If
                    GoTo exitedite
                End If
            End If
        End If
        
        '===== Additional F 12
        '===== Data Element 7, 13
        '===== Error 256
        If tempucr = "09A" Or tempucr = "09B" Or tempucr = "09C" Then
            If Not completedy(t%) Then
                msg = "Homocide must show as completed."
                Call ShowApplicableContainers(completedy(t%))
'---- setfocus logic ----
'                         completedy(t%).SetFocus
          If completedy(t%).Visible Then
              completedy(t%).SetFocus
           End If
                GoTo exitedite
            End If
            sela% = 0
            For iop% = 1 To weapontype.ListItems.Count
                If weapontype.ListItems(iop%).Selected Then
                    sela% = 1
                    iop% = weapontype.ListItems.Count
                End If
            Next iop%
            If sela% = 0 Then
                msg = "A weapon type must be selected for this UCR " + tempucr + "."
                Call ShowApplicableContainers(weapontype)
'---- setfocus logic ----
'                         weapontype.SetFocus
          If weapontype.Visible Then
              weapontype.SetFocus
           End If
                GoTo exitedite
            End If
            If Not individual Then
                msg = "Homocide victims should be Individual."
                Call ShowApplicableContainers(individual)
'---- setfocus logic ----
'                         individual.SetFocus
          If individual.Visible Then
              individual.SetFocus
           End If
                GoTo exitedite
            End If
            sela% = 0
            For iop% = 1 To HOMOCIDE(t%).ListItems.Count
                If HOMOCIDE(t%).ListItems(iop%).Selected Then
                    sela% = 1
                    iop% = HOMOCIDE(t%).ListItems.Count
                End If
            Next iop%
            If sela% = 0 Then
                msg = "Aggravated Assault/Homocide Circumstances must be selected for UCR " + tempucr + "."
                Call ShowApplicableContainers(HOMOCIDE(t%))
'---- setfocus logic ----
'                         HOMOCIDE(t%).SetFocus
          If HOMOCIDE(t%).Visible Then
              HOMOCIDE(t%).SetFocus
           End If
                GoTo exitedite
            End If
            '===== Error 463
            If tempucr = "09C" Then
                hct% = 0
                For xx% = 1 To HOMOCIDE(t%).ListItems.Count
                    If HOMOCIDE(t%).ListItems(xx%).Selected Then
                        hct% = hct% + 1
                        If Mid$(HOMOCIDE(t%).ListItems(xx%), InStr(HOMOCIDE(t%).ListItems(xx%), "(") + 1, 2) <> "20" And Mid$(HOMOCIDE(t%).ListItems(xx%), InStr(HOMOCIDE(t%).ListItems(xx%), "(") + 1, 2) <> "21" Then
                            msg = "Aggravated Assault/Homocide Circumstances must be 20 or 21."
                            Call ShowApplicableContainers(HOMOCIDE(t%))
'---- setfocus logic ----
'                                     HOMOCIDE(t%).SetFocus
          If HOMOCIDE(t%).Visible Then
              HOMOCIDE(t%).SetFocus
           End If
                            GoTo exitedite
                        End If
                    End If
                Next xx%
                If hct% > 2 Then
                    msg = "A maximum of 2 Aggravated Assault/Homocide Circumstances may be selected."
                    Call ShowApplicableContainers(HOMOCIDE(t%))
'---- setfocus logic ----
'                             HOMOCIDE(t%).SetFocus
          If HOMOCIDE(t%).Visible Then
              HOMOCIDE(t%).SetFocus
           End If
                    GoTo exitedite
                End If
                If additional(t%).ListIndex = -1 Then
                    msg = "Additional Circumstances must be selected."
                    Call ShowApplicableContainers(additional(t%))
'---- setfocus logic ----
'                             additional(t%).SetFocus
          If additional(t%).Visible Then
              additional(t%).SetFocus
           End If
                    GoTo exitedite
                End If
            End If
            
        '===== Error 456
        For xx% = 1 To HOMOCIDE(t%).ListItems.Count
            If HOMOCIDE(t%).ListItems(xx%).Selected Then
            '===== Error 480
                If Mid$(HOMOCIDE(t%).ListItems(xx%), InStr(HOMOCIDE(t%).ListItems(xx%), "(") + 1, 2) <> "20" And Mid$(HOMOCIDE(t%).ListItems(xx%), InStr(HOMOCIDE(t%).ListItems(xx%), "(") + 1, 2) = "08" And ucrct% < 2 Then
                    msg = "If reason 08=Other Felony Involved is selected, there must be at least 2 UCR's entered."
                    Call ShowApplicableContainers(HOMOCIDE(t%))
'---- setfocus logic ----
'                             HOMOCIDE(t%).SetFocus
          If HOMOCIDE(t%).Visible Then
              HOMOCIDE(t%).SetFocus
           End If
                    GoTo exitedite
                End If
                If Mid$(HOMOCIDE(t%).ListItems(xx%), InStr(HOMOCIDE(t%).ListItems(xx%), "(") + 1, 2) <> "20" And Mid$(HOMOCIDE(t%).ListItems(xx%), InStr(HOMOCIDE(t%).ListItems(xx%), "(") + 1, 2) = "10" Then
                    For xxx% = 1 To xx% - 1
                        If HOMOCIDE(t%).ListItems(xxx%).Selected Then
                            msg = "If 10=Unknown Circumstances is selected, no other circumstances can be selected."
                            Call ShowApplicableContainers(HOMOCIDE(t%))
'---- setfocus logic ----
'                                     HOMOCIDE(t%).SetFocus
          If HOMOCIDE(t%).Visible Then
              HOMOCIDE(t%).SetFocus
           End If
                            GoTo exitedite
                        End If
                    Next xxx%
                    For xxx% = xx% + 1 To HOMOCIDE(t%).ListItems.Count
                        If HOMOCIDE(t%).ListItems(xxx%).Selected Then
                            msg = "If 10=Unknown Circumstances is selected, no other circumstances can be selected."
                            Call ShowApplicableContainers(HOMOCIDE(t%))
'---- setfocus logic ----
'                                     HOMOCIDE(t%).SetFocus
          If HOMOCIDE(t%).Visible Then
              HOMOCIDE(t%).SetFocus
           End If
                            GoTo exitedite
                        End If
                    Next xxx%
                End If
                '===== Error 456
                Select Case Val(Mid$(HOMOCIDE(t%).ListItems(xx%), InStr(HOMOCIDE(t%).ListItems(xx%), "(") + 1, 2))
                    Case 1 To 10
                        '===== Error 477
                        If tempucr <> "13A" And tempucr <> "09A" Then
                            msg = "The Aggravated Assault/Homocide reason chosen is invalid for the UCR."
                            Call ShowApplicableContainers(HOMOCIDE(t%))
'---- setfocus logic ----
'                                     HOMOCIDE(t%).SetFocus
          If HOMOCIDE(t%).Visible Then
              HOMOCIDE(t%).SetFocus
           End If
                            GoTo exitedite
                        End If
                        For xxx% = 1 To xx% - 1
                            If HOMOCIDE(t%).ListItems(xxx%).Selected Then
                                Select Case Val(Mid$(HOMOCIDE(t%).ListItems(xxx%), InStr(HOMOCIDE(t%).ListItems(xxx%), "(") + 1, 2))
                                    Case 20 To 34
                                        msg = "Only one category may be selected within Aggravted Assault/Homocide reasons."
                                        Call ShowApplicableContainers(HOMOCIDE(t%))
'---- setfocus logic ----
'                                                 HOMOCIDE(t%).SetFocus
          If HOMOCIDE(t%).Visible Then
              HOMOCIDE(t%).SetFocus
           End If
                                        GoTo exitedite
                                End Select
                            End If
                        Next xxx%
                        For xxx% = xx% + 1 To HOMOCIDE(t%).ListItems.Count
                            If HOMOCIDE(t%).ListItems(xxx%).Selected Then
                                Select Case Val(Mid$(HOMOCIDE(t%).ListItems(xxx%), InStr(HOMOCIDE(t%).ListItems(xxx%), "(") + 1, 2))
                                    Case 20 To 34
                                        msg = "Only one category may be selected within Aggravted Assault/Homocide reasons."
                                        Call ShowApplicableContainers(HOMOCIDE(t%))
'---- setfocus logic ----
'                                                 HOMOCIDE(t%).SetFocus
          If HOMOCIDE(t%).Visible Then
              HOMOCIDE(t%).SetFocus
           End If
                                        GoTo exitedite
                                End Select
                            End If
                        Next xxx%
                    Case 20 To 21
                        '===== Error 477
                        If tempucr <> "09C" Then
                            msg = "The Aggravated Assault/Homocide reason chosen is invalid for the UCR."
                            Call ShowApplicableContainers(HOMOCIDE(t%))
'---- setfocus logic ----
'                                     HOMOCIDE(t%).SetFocus
          If HOMOCIDE(t%).Visible Then
              HOMOCIDE(t%).SetFocus
           End If
                            GoTo exitedite
                        End If
                        For xxx% = 1 To xx% - 1
                            If HOMOCIDE(t%).ListItems(xxx%).Selected Then
                                Select Case Val(Mid$(HOMOCIDE(t%).ListItems(xxx%), InStr(HOMOCIDE(t%).ListItems(xxx%), "(") + 1, 2))
                                    Case 1 To 10
                                        msg = "Only one category may be selected within Aggravted Assault/Homocide reasons."
                                        Call ShowApplicableContainers(HOMOCIDE(t%))
'---- setfocus logic ----
'                                                 HOMOCIDE(t%).SetFocus
          If HOMOCIDE(t%).Visible Then
              HOMOCIDE(t%).SetFocus
           End If
                                        GoTo exitedite
                                    Case 30 To 34
                                        msg = "Only one category may be selected within Aggravted Assault/Homocide reasons."
                                        Call ShowApplicableContainers(HOMOCIDE(t%))
'---- setfocus logic ----
'                                                 HOMOCIDE(t%).SetFocus
          If HOMOCIDE(t%).Visible Then
              HOMOCIDE(t%).SetFocus
           End If
                                        GoTo exitedite
                                End Select
                            End If
                        Next xxx%
                        For xxx% = xx% + 1 To HOMOCIDE(t%).ListItems.Count
                            If HOMOCIDE(t%).ListItems(xxx%).Selected Then
                                Select Case Val(Mid$(HOMOCIDE(t%).ListItems(xxx%), InStr(HOMOCIDE(t%).ListItems(xxx%), "(") + 1, 2))
                                    Case 1 To 10
                                        msg = "Only one category may be selected within Aggravted Assault/Homocide reasons."
                                        Call ShowApplicableContainers(HOMOCIDE(t%))
'---- setfocus logic ----
'                                                 HOMOCIDE(t%).SetFocus
          If HOMOCIDE(t%).Visible Then
              HOMOCIDE(t%).SetFocus
           End If
                                        GoTo exitedite
                                    Case 30 To 34
                                        msg = "Only one category may be selected within Aggravted Assault/Homocide reasons."
                                        Call ShowApplicableContainers(HOMOCIDE(t%))
'---- setfocus logic ----
'                                                 HOMOCIDE(t%).SetFocus
          If HOMOCIDE(t%).Visible Then
              HOMOCIDE(t%).SetFocus
           End If
                                        GoTo exitedite
                                End Select
                            End If
                        Next xxx%
                    Case 30 To 34
                        '===== Error 477
                        If tempucr <> "09B" Then
                            msg = "The Aggravated Assault/Homocide reason chosen is invalid for the UCR."
                            Call ShowApplicableContainers(HOMOCIDE(t%))
'---- setfocus logic ----
'                                     HOMOCIDE(t%).SetFocus
          If HOMOCIDE(t%).Visible Then
              HOMOCIDE(t%).SetFocus
           End If
                            GoTo exitedite
                        End If
                        For xxx% = 1 To xx% - 1
                            If HOMOCIDE(t%).ListItems(xxx%).Selected Then
                                Select Case Val(Mid$(HOMOCIDE(t%).ListItems(xxx%), InStr(HOMOCIDE(t%).ListItems(xxx%), "(") + 1, 2))
                                    Case 1 To 21
                                        msg = "Only one category may be selected within Aggravted Assault/Homocide reasons."
                                        Call ShowApplicableContainers(HOMOCIDE(t%))
'---- setfocus logic ----
'                                                 HOMOCIDE(t%).SetFocus
          If HOMOCIDE(t%).Visible Then
              HOMOCIDE(t%).SetFocus
           End If
                                        GoTo exitedite
                                End Select
                            End If
                        Next xxx%
                        For xxx% = xx% + 1 To HOMOCIDE(t%).ListItems.Count
                            If HOMOCIDE(t%).ListItems(xxx%).Selected Then
                                Select Case Val(Mid$(HOMOCIDE(t%).ListItems(xxx%), InStr(HOMOCIDE(t%).ListItems(xxx%), "(") + 1, 2))
                                    Case 1 To 21
                                        msg = "Only one category may be selected within Aggravted Assault/Homocide reasons."
                                        Call ShowApplicableContainers(HOMOCIDE(t%))
'---- setfocus logic ----
'                                                 HOMOCIDE(t%).SetFocus
          If HOMOCIDE(t%).Visible Then
              HOMOCIDE(t%).SetFocus
           End If
                                        GoTo exitedite
                                End Select
                            End If
                        Next xxx%
                End Select
            End If
        Next xx%

        End If
                

        '===== Error 455
        If tempucr = "09A" Or tempucr = "09B" Or tempucr = "09C" Or tempucr = "13A" Then
            For xx% = 1 To HOMOCIDE(t%).ListItems.Count
                If HOMOCIDE(t%).ListItems(xx%).Selected Then
                    If Mid$(HOMOCIDE(t%).ListItems(xx%), InStr(HOMOCIDE(t%).ListItems(xx%), "(") + 1, 2) <> "20" And Mid$(HOMOCIDE(t%).ListItems(xx%), InStr(HOMOCIDE(t%).ListItems(xx%), "(") + 1, 2) = "20" Or _
                       Mid$(HOMOCIDE(t%).ListItems(xx%), InStr(HOMOCIDE(t%).ListItems(xx%), "(") + 1, 2) <> "20" And Mid$(HOMOCIDE(t%).ListItems(xx%), InStr(HOMOCIDE(t%).ListItems(xx%), "(") + 1, 2) = "21" Then
                        If additional(t%).ListIndex = -1 Then
                            msg = "Additional Circumstances must be selected."
                            Call ShowApplicableContainers(HOMOCIDE(t%))
'---- setfocus logic ----
'                                     HOMOCIDE(t%).SetFocus
          If HOMOCIDE(t%).Visible Then
              HOMOCIDE(t%).SetFocus
           End If
                            GoTo exitedite
                        End If
                    End If
                End If
            Next xx%
        End If
        
        '===== Additional F 13
        '===== Data Element 7
        If tempucr = "100" Then
            sela% = 0
            For iop% = 1 To weapontype.ListItems.Count
                If weapontype.ListItems(iop%).Selected Then
                    sela% = 1
                    iop% = weapontype.ListItems.Count
                End If
            Next iop%
            If sela% = 0 Then
                msg = "A weapon type must be selected for this UCR " + tempucr + "."
                Call ShowApplicableContainers(weapontype)
'---- setfocus logic ----
'                         weapontype.SetFocus
          If weapontype.Visible Then
              weapontype.SetFocus
           End If
                GoTo exitedite
            End If
            'If injury.SelectedItem Is Nothing Then
            '    msg = "An injury type must be selected for this UCR " + tempucr + "."
            '    GoTo exitedite
            'End If
            If Not individual Then
                msg = "Kidnaping/abduction crimes victims should be Individual."
                Call ShowApplicableContainers(individual)
'---- setfocus logic ----
'                         individual.SetFocus
          If individual.Visible Then
              individual.SetFocus
           End If
                GoTo exitedite
            End If
        End If

        '===== Additional F 16
        '===== Data Element 12
        If tempucr = "370" Then
            sela% = 0
            For iop% = 1 To activity(t%).ListItems.Count
                If activity(t%).ListItems(iop%).Selected Then
                    sela% = 1
                    iop% = activity(t%).ListItems.Count
                End If
            Next iop%
            If sela% = 0 Then
                msg = "Type of Activity must be selected for Pornography."
                Call ShowApplicableContainers(activity(t%))
'---- setfocus logic ----
'                         activity(t%).SetFocus
          If activity(t%).Visible Then
              activity(t%).SetFocus
           End If
                GoTo exitedite
            End If
            If Not societypublic Then
                msg = "Pornography crimes victims should be Society/Public."
                Call ShowApplicableContainers(societypublic)
'---- setfocus logic ----
'                         societypublic.SetFocus
          If societypublic.Visible Then
              societypublic.SetFocus
           End If
                GoTo exitedite
            End If
        End If
            
        '===== Additional F 17
        If tempucr = "40A" Then
            If Not societypublic Then
                msg = "Prostitution crimes victims should be Society/Public."
                Call ShowApplicableContainers(societypublic)
'---- setfocus logic ----
'                         societypublic.SetFocus
          If societypublic.Visible Then
              societypublic.SetFocus
           End If
                GoTo exitedite
            End If
        End If
            
        '===== Additional F 18
        If tempucr = "120" Then
            sela% = 0
            For iop% = 1 To weapontype.ListItems.Count
                If weapontype.ListItems(iop%).Selected Then
                    sela% = 1
                    iop% = weapontype.ListItems.Count
                End If
            Next iop%
            If sela% = 0 Then
                msg = "A weapon type must be selected for this UCR " + tempucr + "."
                Call ShowApplicableContainers(weapontype)
'---- setfocus logic ----
'                         weapontype.SetFocus
          If weapontype.Visible Then
              weapontype.SetFocus
           End If
                GoTo exitedite
            End If
            'If individual Then
            '    If injury.SelectedItem Is Nothing Then
            '        msg = "A type of injury must be selected."
            '        GoTo exitedite
            '    End If
            'End If
        End If
            
        '===== Additional F 19
        If tempucr = "11A" Or tempucr = "11B" Or tempucr = "11C" Or tempucr = "11D" Then
            sela% = 0
            For iop% = 1 To weapontype.ListItems.Count
                If weapontype.ListItems(iop%).Selected Then
                    sela% = 1
                    iop% = weapontype.ListItems.Count
                End If
            Next iop%
            If sela% = 0 Then
                msg = "A weapon type must be selected for this UCR " + tempucr + "."
                Call ShowApplicableContainers(weapontype)
'---- setfocus logic ----
'                         weapontype.SetFocus
          If weapontype.Visible Then
              weapontype.SetFocus
           End If
                GoTo exitedite
            End If
            If Not individual Then
                msg = "Individual must be chosen for UCR " + tempucr + "."
                Call ShowApplicableContainers(individual)
'---- setfocus logic ----
'                         individual.SetFocus
          If individual.Visible Then
              individual.SetFocus
           End If
                GoTo exitedite
            End If
        End If
        
        '===== Additional F 20
        If tempucr = "36A" Or tempucr = "36B" Then
            If Not individual Then
                msg = "Individual must be chosen for UCR " + tempucr + "."
                Call ShowApplicableContainers(individual)
'---- setfocus logic ----
'                         individual.SetFocus
          If individual.Visible Then
              individual.SetFocus
           End If
                GoTo exitedite
            End If
        End If
        
        '===== Additional F 21
        '===== Data Element 12
        If tempucr = "280" Then
            sela% = 0
            For iop% = 1 To activity(t%).ListItems.Count
                If activity(t%).ListItems(iop%).Selected Then
                    sela% = 1
                    iop% = activity(t%).ListItems.Count
                End If
            Next iop%
            If sela% = 0 Then
                msg = "Type of Activity must be selected for Stolen Property."
                Call ShowApplicableContainers(activity(t%))
'---- setfocus logic ----
'                         activity(t%).SetFocus
          If activity(t%).Visible Then
              activity(t%).SetFocus
           End If
                GoTo exitedite
            End If
        End If
        
        '===== Additional F 22
        '===== Data Element 12
        If tempucr = "520" Then
            sela% = 0
            For iop% = 1 To activity(t%).ListItems.Count
                If activity(t%).ListItems(iop%).Selected Then
                    sela% = 1
                    iop% = activity(t%).ListItems.Count
                End If
            Next iop%
            If sela% = 0 Then
                msg = "Type of Activity must be selected for Weapons Law violations."
                Call ShowApplicableContainers(activity(t%))
'---- setfocus logic ----
'                         activity(t%).SetFocus
          If activity(t%).Visible Then
              activity(t%).SetFocus
           End If
                GoTo exitedite
            End If
            sela% = 0
            For iop% = 1 To weapontype.ListItems.Count
                If weapontype.ListItems(iop%).Selected Then
                    sela% = 1
                    iop% = weapontype.ListItems.Count
                End If
            Next iop%
            If sela% = 0 Then
                msg = "A weapon type must be selected for this UCR " + tempucr + "."
                Call ShowApplicableContainers(weapontype)
'---- setfocus logic ----
'                         weapontype.SetFocus
          If weapontype.Visible Then
              weapontype.SetFocus
           End If
                GoTo exitedite
            End If
            If Not societypublic Then
                msg = "Society must be chosen for UCR " + tempucr + "."
                Call ShowApplicableContainers(societypublic)
'---- setfocus logic ----
'                         societypublic.SetFocus
          If societypublic.Visible Then
              societypublic.SetFocus
           End If
                GoTo exitedite
            End If
        End If
        
        '==== Mandatories E - 9
        '===== Error 201
        If ctpremise% = 0 Then
            msg = "A valid premise type must be selected."
            Call ShowApplicableContainers(entered(t%))
'---- setfocus logic ----
'                     entered(t%).SetFocus
          If entered(t%).Visible Then
              entered(t%).SetFocus
           End If
            GoTo exitedite
        End If
        If entered(t%) = "" Then
            entered(t%) = "0"
        End If
        If tempucr = "09A" Or tempucr = "09B" Or tempucr = "100" Or tempucr = "120" Or tempucr = "11A" Or tempucr = "11B" Or tempucr = "11C" Or tempucr = "11D" Or tempucr = "13A" Or tempucr = "13B" Or tempucr = "13C" Then
            tempactivity = ""
            foundn% = 0
            foundg% = 0
            foundj% = 0
            For q% = 1 To gactivity(t%).ListItems.Count
                If gactivity(t%).ListItems(q%).Selected Then
                    Select Case Mid$(gactivity(t%).ListItems(q%), InStr(gactivity(t%).ListItems(q%), "(") + 1, 1)
                        Case "N"
                            foundn% = 1
                        Case "G"
                            foundg% = 1
                        Case "J"
                            foundj% = 1
                    End Select
                End If
            Next q%
            If foundn% = 1 And (foundj% = 1 Or foundg% = 1) Then
                msg = "If activity code is 'N', codes 'J' and 'G' are not permitted."
                Call ShowApplicableContainers(gactivity(t%))
'---- setfocus logic ----
'                         gactivity(t%).SetFocus
          If gactivity(t%).Visible Then
              gactivity(t%).SetFocus
           End If
                GoTo exitedite
            End If
        End If
    End If
Next t%
'===== Data Element 2
'===== Error 001, 101, 201, 301, 401, 501, 601, 701
If incidentnumber = "" Then
    msg = "A valid incidentnumber must be entered."
    Call ShowApplicableContainers(incidentnumber)
'---- setfocus logic ----
'             incidentnumber.SetFocus
          If incidentnumber.Visible Then
              incidentnumber.SetFocus
           End If
    GoTo exitedite
End If
If Len(incidentnumber) > 12 Then
    msg = "The Incident Number cannot be over 12 characters long."
    Call ShowApplicableContainers(incidentnumber)
'---- setfocus logic ----
'             incidentnumber.SetFocus
          If incidentnumber.Visible Then
              incidentnumber.SetFocus
           End If
    GoTo exitedite
End If
incidentnumber = UCase(incidentnumber)
'===== Error 017, 117, 217, 317, 417, 517, 617, 717
For t% = 1 To Len(incidentnumber)
    If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789- ", Mid$(incidentnumber, t%, 1)) = 0 Then
        msg = "An invalid character has been found in the Incident Number field.  Valid characters are A-Z, 0-9, and Hyphen.  Do not enter any Blanks because these are computer generated."
        t% = Len(incidentnumber)
        Call ShowApplicableContainers(incidentnumber)
'---- setfocus logic ----
'                 incidentnumber.SetFocus
          If incidentnumber.Visible Then
              incidentnumber.SetFocus
           End If
        GoTo exitedite
    End If
Next t%
'===== Error 015, 115, 215, 315, 415, 515, 615, 715
temp$ = ""
For yy% = Len(incidentnumber) To 1 Step -1
    If Mid$(incidentnumber, yy%, 1) <> " " Then
        temp$ = Left$(incidentnumber, yy%)
        yy% = 1
    End If
Next yy%
If InStr(temp$, " ") > 0 Then
    msg = "An invalid character has been found in the Incident Number field.  Valid characters are A-Z, 0-9, and Hyphen.  Do not enter any Blanks because these are computer generated."
    t% = Len(incidentnumber)
    Call ShowApplicableContainers(incidentnumber)
'---- setfocus logic ----
'             incidentnumber.SetFocus
          If incidentnumber.Visible Then
              incidentnumber.SetFocus
           End If
    GoTo exitedite
End If
If Len(incidentnumber) < 12 And Len(incidentnumber) > 0 Then
    incidentnumber = incidentnumber + Space$(12 - Len(incidentnumber))
End If
'If Not IsDate(dispatchdate) Then
'    msg = "Dispatch date is invalid."
'    GoTo exitedite
'End If
'If Not IsDate(DISPATCHTIME) Then
'    msg = "Dispatch time is invalid."
'    GoTo exitedite
'End If
'If Not IsDate(TIMEARRIVED) Then
'    msg = "Arrival time is invalid."
'    GoTo exitedite
'End If
'If Not IsDate(DEPARTINGTIME) Then
'    msg = "Departure time is invalid."
'    GoTo exitedite
'End If
'Call ShowApplicableContainers(incidentnumber)
'---- setfocus logic ----
'         incidentnumber.SetFocus
          If incidentnumber.Visible Then
              incidentnumber.SetFocus
           End If
GoTo goodedite
exitedite:
editerr = 1

goodedite:

Exit Sub
'RLB Bandaid
rlbErre:
    If Err.Number = 5 Then Resume Next
    Resume
    
End Sub

Friend Sub editvictim(editerr, msg As String)
Dim db As Database, rs, rs2 As Recordset
'RLB Bandaid
On Error GoTo rlbErrv
If Len(incidentnumber) < 12 And Len(incidentnumber) > 0 Then
    incidentnumber = incidentnumber + Space$(12 - Len(incidentnumber))
End If
If age(1) > "" Then
    If Val(age(1)) = 0 And age(1) <> "00" Then
        msg = "Age must be entered in the format of a single age: X or XX (i.e. 3 or 42) or in the format of a range of ages: XXXX (i.e. 2025, meaning 20 - 25 years old)."
        Call ShowApplicableContainers(age(1))
'---- setfocus logic ----
'                 age(1).SetFocus
          If age(1).Visible Then
              age(1).SetFocus
           End If
        GoTo exiteditv
    End If
End If
'===== Error 404
If age(1) <> "NN" And age(1) <> "NB" And age(1) <> "BB" Then
    For t% = 1 To Len(age(1))
        If InStr("0123456789-", Mid$(age(1), t%, 1)) = 0 Then
            msg = "An invalid entry in age has been found.  Valid Entry Formats:  X, XX, X-Y, XX-Y, X-YY, XX-YY, XXYY"
            t% = Len(age(1))
            Call ShowApplicableContainers(age(1))
'---- setfocus logic ----
'                     age(1).SetFocus
          If age(1).Visible Then
              age(1).SetFocus
           End If
            GoTo exiteditv
        End If

    Next t%
End If
If InStr(age(1), "-") > 0 Then
    ag1 = Val(Left$(age(1), InStr(age(1), "-") - 1))
    ag2 = Val(Mid$(age(1), InStr(age(1), "-") + 1))
    age(1) = Format$(ag1, "00") + Format$(ag2, "00")
End If
'===== Error 410,422
If Len(age(1)) = 4 Then
    If Val(Left$(age(1), 2)) >= Val(Right$(age(1), 2)) Then
        msg = "For an age range, the first age must be less than the second age."
        Call ShowApplicableContainers(age(1))
'---- setfocus logic ----
'                 age(1).SetFocus
          If age(1).Visible Then
              age(1).SetFocus
           End If
        GoTo exiteditv
    End If
    If Val(Left$(age(1), 2)) = 0 Then
        msg = "The low value in an age range cannot be 0."
        Call ShowApplicableContainers(age(1))
'---- setfocus logic ----
'                 age(1).SetFocus
          If age(1).Visible Then
              age(1).SetFocus
           End If
        GoTo exiteditv
    End If
End If
'===== Error 450
For r% = 10 To 19
    If relationship(r%).ListIndex > -1 Then
        temprel = Mid$(relationship(r%).List(relationship(r%).ListIndex), InStr(relationship(r%).List(relationship(r%).ListIndex), "(") + 1, 2)
        If temprel = "SE" Then
            If Val(age(1)) < 10 Then
                msg = "The relationship of victim to subject cannot be 'SE' when victim's age is less than 10."
                Call ShowApplicableContainers(age(1))
'---- setfocus logic ----
'                         age(1).SetFocus
          If age(1).Visible Then
              age(1).SetFocus
           End If
                GoTo exiteditv
            End If
        End If
        '===== Error 553
        If r% = 10 Then
            Select Case temprel
                Case "BG", "XS", "SE", "HR"
                    If sex(1).List(sex(1).ListIndex) = sex(2).List(sex(2).ListIndex) Then
                        msg = "For relationships of BG, XS, SE, and CS, the sexes of the victim and subject must be different."
                        Call ShowApplicableContainers(sex(1))
'---- setfocus logic ----
'                                 sex(1).SetFocus
          If sex(1).Visible Then
              sex(1).SetFocus
           End If
                        GoTo exiteditv
                    End If
                Case "HR"
                    If sex(1).List(sex(1).ListIndex) <> sex(2).List(sex(2).ListIndex) Then
                        msg = "For relationships of HR, the sexes of the victim and subject must be the same."
                        Call ShowApplicableContainers(sex(1))
'---- setfocus logic ----
'                                 sex(1).SetFocus
          If sex(1).Visible Then
              sex(1).SetFocus
           End If
                        GoTo exiteditv
                    End If
            End Select
        End If
    End If
Next r%
'===== Error 472
If UCase(vsname(2)) = "UNKNOWN" Then
    If (race(1).ListIndex > -1 And Left$(race(1).List(race(1).ListIndex), 1) = "U") And _
        (sex(1).ListIndex > -1 And Left$(sex(1).List(sex(1).ListIndex), 1) = "U") And _
        (ethnicity(1).ListIndex > -1 And Left$(ethnicity(1).List(ethnicity(1).ListIndex), 1) = "U") And _
        Val(age(1)) = 0 Then
    Else
    If relationship(10).ListIndex > -1 Then
        temprel = Mid$(relationship(10).List(relationship(10).ListIndex), InStr(relationship(10).List(relationship(10).ListIndex), "(") + 1, 2)
        If temprel <> "RU" Then
            msg = "The relationship of victim to subject must be Relationship Unknown if subject is UNKNOWN."
            Call ShowApplicableContainers(relationship(10))
'---- setfocus logic ----
'                     relationship(10).SetFocus
          If relationship(10).Visible Then
              relationship(10).SetFocus
           End If
            GoTo exiteditv
        End If
    Else
        msg = "The relationship of victim to subject must be Relationship Unknown if subject is UNKNOWN."
        Call ShowApplicableContainers(relationship(10))
'---- setfocus logic ----
'                 relationship(10).SetFocus
          If relationship(10).Visible Then
              relationship(10).SetFocus
           End If
        GoTo exiteditv
    End If
    End If
End If
'===== Error 401
If Not individual And Not business And Not financialinstitution And Not government And Not religiousorganization And Not societypublic And Not other And Not unknown And Not policeofficer Then
    msg = "A Type of Victim must be chosen."
    Call ShowApplicableContainers(individual)
'---- setfocus logic ----
'             individual.SetFocus
          If individual.Visible Then
              individual.SetFocus
           End If
    GoTo exiteditv
End If
editvictimdetail:
'===== Data Element 20, 21, 22
Dim drugs(2)
drugs(0) = ""
drugs(1) = ""
drugs(2) = ""
dct% = 0
For ii% = 0 To 2
    If drugtype(ii%).ListIndex > -1 Then
        drugs(dct%) = drugtype(ii%).List(drugtype(ii%).ListIndex)
        dct% = dct% + 1
    End If
Next ii%
If dct% > 0 Then
    For Z% = 0 To dct%
        '===== Error 306
        If Z% > 0 Then
            For ZZ% = 0 To Z% - 1
                If drugs(ZZ%) = drugs(Z%) Then
                    If InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "GM=") > 0 Or _
                       InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "KG=") > 0 Or _
                       InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "OZ=") > 0 Or _
                       InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "LB=") > 0 Then
                        If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "GM=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "KG=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "OZ=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "LB=") > 0 Then
                                msg = "Two different measurements within the same weight category cannot be entered for the same type of drug."
                                Call ShowApplicableContainers(drugmeasurement(ZZ%))
'---- setfocus logic ----
'                                         drugmeasurement(zz%).SetFocus
          If drugmeasurement(ZZ%).Visible Then
              drugmeasurement(ZZ%).SetFocus
           End If
                                GoTo exiteditv
                        End If
                    End If
                    If InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "ML=") > 0 Or _
                       InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "LT=") > 0 Or _
                       InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "FO=") > 0 Or _
                       InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "GL=") > 0 Then
                        If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "ML=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "LT=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "FO=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "GL=") > 0 Then
                                msg = "Two different measurements within the same weight category cannot be entered for the same type of drug."
                                Call ShowApplicableContainers(drugmeasurement(ZZ%))
'---- setfocus logic ----
'                                         drugmeasurement(zz%).SetFocus
          If drugmeasurement(ZZ%).Visible Then
              drugmeasurement(ZZ%).SetFocus
           End If
                                GoTo exiteditv
                        End If
                    End If
                    If InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "DU=") > 0 Or _
                       InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "NP=") > 0 Then
                        If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "DU=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "NP=") > 0 Then
                                msg = "Two different measurements within the same weight category cannot be entered for the same type of drug."
                                Call ShowApplicableContainers(drugmeasurement(ZZ%))
'---- setfocus logic ----
'                                         drugmeasurement(zz%).SetFocus
          If drugmeasurement(ZZ%).Visible Then
              drugmeasurement(ZZ%).SetFocus
           End If
                                GoTo exiteditv
                        End If
                    End If
                End If
            Next ZZ%
        End If
        For t% = 1 To Len(drugamt(Z%))
            If InStr("0123456789.", Mid$(drugamt(Z%), t%, 1)) = 0 Then
                msg = "Only 0123456789. are allowed in drug amount.  Enter all fractions as decimals (i.e. 1/2 = .5)."
                Call ShowApplicableContainers(drugamt(Z%))
'---- setfocus logic ----
'                         drugamt(Z%).SetFocus
          If drugamt(Z%).Visible Then
              drugamt(Z%).SetFocus
           End If
                GoTo exiteditv
            End If
        Next t%
        '=====SCEdit 4/21/92 P29  Overtuend drug amt and measurement error
        'If drugamt(z%) = "" Or drugmeasurement(z%).ListIndex = -1 Then
        '    If drugs(z%) > "" And Left$(drugs(z%), 1) <> "X" And Left$(drugs(z%), 1) <> "U" Then
        '        msg = "Drug Quantity and Measurement Type must be entered/selected."
        '        GoTo exiteditv
        '    End If
        'End If
        '===== Error 366
        If drugamt(Z%) > "" Then
            If drugs(Z%) = "" Or drugmeasurement(Z%).ListIndex = -1 Then
                msg = "If a drug quantity is entered, then drug type and measurement type must also be entered."
                Call ShowApplicableContainers(drugamt(Z%))
'---- setfocus logic ----
'                         drugamt(Z%).SetFocus
          If drugamt(Z%).Visible Then
              drugamt(Z%).SetFocus
           End If
                GoTo exiteditv
            End If
        End If
        '===== Error 367
        If drugmeasurement(Z%).ListIndex > -1 Then
            If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "NP=)") > 0 Then
                If Left$(drugs(Z%), 1) <> "E" And Left$(drugs(Z%), 1) <> "G" And Left$(drugs(Z%), 1) <> "K" Then
                    msg = "Number of Plants is only allowabled with Marijuana, Opium, and Other Hallucinogens."
                    Call ShowApplicableContainers(drugmeasurement(Z%))
'---- setfocus logic ----
'                             drugmeasurement(Z%).SetFocus
          If drugmeasurement(Z%).Visible Then
              drugmeasurement(Z%).SetFocus
           End If
                    GoTo exiteditv
                End If
            End If
            '===== Error 384
            If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "XX=") > 0 Then
                If Val(drugamt(Z%)) <> 1 Then
                    msg = "If drug measurement is NOT REPORTED, drug amount must be 1."
                    Call ShowApplicableContainers(drugamt(Z%))
'---- setfocus logic ----
'                             drugamt(Z%).SetFocus
          If drugamt(Z%).Visible Then
              drugamt(Z%).SetFocus
           End If
                    GoTo exiteditv
                End If
            End If
            '===== Error 368
            If drugs(Z%) = "" Or drugamt(Z%) = "" Then
                msg = "If a drug measurement is entered, then drug type and quantity must also be entered."
                Call ShowApplicableContainers(drugamt(Z%))
'---- setfocus logic ----
'                         drugamt(Z%).SetFocus
          If drugamt(Z%).Visible Then
              drugamt(Z%).SetFocus
           End If
                GoTo exiteditv
            End If
        End If
        '===== Error 362
        If Left$(drugs(Z%), 1) = "X" Then
            If drugtype(0).ListIndex = -1 Or drugtype(1).ListIndex = -1 Or drugtype(2).ListIndex = -1 Then
                msg = "If X = Over 3 Drug Types is entered, then the other 2 drug types must also be entered."
                Call ShowApplicableContainers(drugtype(0))
'---- setfocus logic ----
'                         drugtype(0).SetFocus
          If drugtype(0).Visible Then
              drugtype(0).SetFocus
           End If
                GoTo exiteditv
            End If
            '===== Error 363
            If drugamt(Z%) > "" Or drugmeasurement(Z%).ListIndex > -1 Then
                msg = "Drug Amount and Measurement are not valid entries for X=Over 3 Drug Types."
                Call ShowApplicableContainers(drugamt(Z%))
'---- setfocus logic ----
'                         drugamt(Z%).SetFocus
          If drugamt(Z%).Visible Then
              drugamt(Z%).SetFocus
           End If
                GoTo exiteditv
            End If
        End If
    Next Z%
End If
        
'==== Mandatories E - 25 = GIVEN
'==== Mandatories E - 26, 27, 28
'===== Error 453
If individual Or policeofficer Then
    If (Val(age(1)) = 0 And age(1) <> "00") Then
        msg = "Invalid age entered."
        Call ShowApplicableContainers(age(1))
'---- setfocus logic ----
'                 age(1).SetFocus
          If age(1).Visible Then
              age(1).SetFocus
           End If
        GoTo exiteditv
    End If
    If sex(1).ListIndex = -1 Then
        msg = "Invalid sex entered."
        Call ShowApplicableContainers(sex(1))
'---- setfocus logic ----
'                 sex(1).SetFocus
          If sex(1).Visible Then
              sex(1).SetFocus
           End If
        GoTo exiteditv
    End If
    If race(1).ListIndex = -1 Then
        msg = "Invalid race entered."
        Call ShowApplicableContainers(race(1))
'---- setfocus logic ----
'                 race(1).SetFocus
          If race(1).Visible Then
              race(1).SetFocus
           End If
        GoTo exiteditv
    End If
    If ethnicity(1).ListIndex = -1 Then
        msg = "Ethnicity is a required entry."
        Call ShowApplicableContainers(ethnicity(1))
'---- setfocus logic ----
'                 ethnicity(1).SetFocus
          If ethnicity(1).Visible Then
              ethnicity(1).SetFocus
           End If
        GoTo exiteditv
    End If
    If resident(1).ListIndex = -1 Then
        msg = "Resident Status is a required entry."
        Call ShowApplicableContainers(resident(1))
'---- setfocus logic ----
'                 resident(1).SetFocus
          If resident(1).Visible Then
              resident(1).SetFocus
           End If
        GoTo exiteditv
    End If
Else
    '===== Error 458
    If Val(age(1)) > 0 Then
        msg = "Age is not a valid entry for Victim if Type of Victim is not Individual."
        Call ShowApplicableContainers(age(1))
'---- setfocus logic ----
'                 age(1).SetFocus
          If age(1).Visible Then
              age(1).SetFocus
           End If
        GoTo exiteditv
    End If
    If sex(1).ListIndex > -1 And Left$(sex(1).List(sex(1).ListIndex), 1) <> "U" Then
        msg = "Sex is not a valid entry for Victim if Type of Victim is not Individual."
        Call ShowApplicableContainers(sex(1))
'---- setfocus logic ----
'                 sex(1).SetFocus
          If sex(1).Visible Then
              sex(1).SetFocus
           End If
        GoTo exiteditv
    End If
    If race(1).ListIndex > -1 And Left$(race(1).List(race(1).ListIndex), 1) <> "U" Then
        msg = "Race is not a valid entry for Victim if Type of Victim is not Individual."
        Call ShowApplicableContainers(race(1))
'---- setfocus logic ----
'                 race(1).SetFocus
          If race(1).Visible Then
              race(1).SetFocus
           End If
        GoTo exiteditv
    End If
    If ethnicity(1).ListIndex > -1 And Left$(ethnicity(1).List(ethnicity(1).ListIndex), 1) <> "U" Then
        msg = "Ethnicity is not a valid entry for Victim if Type of Victim is not Individual."
        Call ShowApplicableContainers(ethnicity(1))
'---- setfocus logic ----
'                 ethnicity(1).SetFocus
          If ethnicity(1).Visible Then
              ethnicity(1).SetFocus
           End If
        GoTo exiteditv
    End If
    If resident(1).ListIndex > -1 And Left$(resident(1).List(resident(1).ListIndex), 1) <> "U" Then
        msg = "Resident Status is not a valid entry for Victim if Type of Victim is not Individual."
        Call ShowApplicableContainers(resident(1))
'---- setfocus logic ----
'                 resident(1).SetFocus
          If resident(1).Visible Then
              resident(1).SetFocus
           End If
        GoTo exiteditv
    End If
End If
'===== Data Element 24
'===== Error 401
If vucrlist.ListItems.Count = 0 Then
    msg = "Victim/UCR assignment not completed."
    Call Command1_Click
    GoTo exiteditv
Else
If vucrlist.SelectedItem Is Nothing Then
    For tuv = 1 To vucrlist.ListItems.Count
        vucrlist.ListItems(tuv).Selected = True
    Next tuv
End If
End If
FOUNDVUCR = 0
For t% = 1 To vucrlist.ListItems.Count
    If vucrlist.ListItems(t%).Selected Then
        FOUNDVUCR = 1
        t% = vucrlist.ListItems.Count
    End If
Next t%
If FOUNDVUCR = 0 Then
    fromfind = 1
    Call Command1_Click
    fromfind = 0
    For t% = 1 To vucrlist.ListItems.Count
        If vucrlist.ListItems(t%).Selected Then
            FOUNDVUCR = 1
            t% = vucrlist.ListItems.Count
        End If
    Next t%
End If
If FOUNDVUCR = 0 Then
    msg = "At least one UCR code must be connected to the victim."
    Call ShowApplicableContainers(vucrlist)
'---- setfocus logic ----
'             vucrlist.SetFocus
          If vucrlist.Visible Then
              vucrlist.SetFocus
           End If
    GoTo exiteditv
End If
'===== SCEdit 3/5/97 Attachment C
For t% = 1 To vucrlist.ListItems.Count
    If vucrlist.ListItems(t%).Selected Then
        tempvucr = Mid$(vucrlist.ListItems(t%), InStr(vucrlist.ListItems(t%), "(") + 1, 3)
        Select Case tempvucr
            Case "09A"
                For tt% = 1 To vucrlist.ListItems.Count
                    If tt% <> t% Then
                        If vucrlist.ListItems(tt%).Selected Then
                            tempvucr2 = Mid$(vucrlist.ListItems(tt%), InStr(vucrlist.ListItems(tt%), "(") + 1, 3)
                            Select Case tempvucr2
                                Case "09B", "13A", "13B", "13C"
                                    msg = "Invalid UCR Combination for Victim"
                                    Call ShowApplicableContainers(vucrlist)
'---- setfocus logic ----
'                                             vucrlist.SetFocus
          If vucrlist.Visible Then
              vucrlist.SetFocus
           End If
                                    GoTo exiteditv
                            End Select
                        End If
                    End If
                Next tt%
            Case "09B"
                For tt% = 1 To vucrlist.ListItems.Count
                    If tt% <> t% Then
                        If vucrlist.ListItems(tt%).Selected Then
                            tempvucr2 = Mid$(vucrlist.ListItems(tt%), InStr(vucrlist.ListItems(tt%), "(") + 1, 3)
                            Select Case tempvucr2
                                Case "09A", "13A", "13B", "13C"
                                    msg = "Invalid UCR Combination for Victim"
                                    Call ShowApplicableContainers(vucrlist)
'---- setfocus logic ----
'                                             vucrlist.SetFocus
          If vucrlist.Visible Then
              vucrlist.SetFocus
           End If
                                    GoTo exiteditv
                            End Select
                        End If
                    End If
                Next tt%
            Case "11A"
                For tt% = 1 To vucrlist.ListItems.Count
                    If tt% <> t% Then
                        If vucrlist.ListItems(tt%).Selected Then
                            tempvucr2 = Mid$(vucrlist.ListItems(tt%), InStr(vucrlist.ListItems(tt%), "(") + 1, 3)
                            Select Case tempvucr2
                                Case "11D", "13A", "13B", "13C", "36A", "36B"
                                    msg = "Invalid UCR Combination for Victim"
                                    Call ShowApplicableContainers(vucrlist)
'---- setfocus logic ----
'                                             vucrlist.SetFocus
          If vucrlist.Visible Then
              vucrlist.SetFocus
           End If
                                    GoTo exiteditv
                            End Select
                        End If
                    End If
                Next tt%
            Case "11B"
                For tt% = 1 To vucrlist.ListItems.Count
                    If tt% <> t% Then
                        If vucrlist.ListItems(tt%).Selected Then
                            tempvucr2 = Mid$(vucrlist.ListItems(tt%), InStr(vucrlist.ListItems(tt%), "(") + 1, 3)
                            Select Case tempvucr2
                                Case "11D", "13A", "13B", "13C", "36A", "36B"
                                    msg = "Invalid UCR Combination for Victim"
                                    Call ShowApplicableContainers(vucrlist)
'---- setfocus logic ----
'                                             vucrlist.SetFocus
          If vucrlist.Visible Then
              vucrlist.SetFocus
           End If
                                    GoTo exiteditv
                            End Select
                        End If
                    End If
                Next tt%
            Case "11C"
                For tt% = 1 To vucrlist.ListItems.Count
                    If tt% <> t% Then
                        If vucrlist.ListItems(tt%).Selected Then
                            tempvucr2 = Mid$(vucrlist.ListItems(tt%), InStr(vucrlist.ListItems(tt%), "(") + 1, 3)
                            Select Case tempvucr2
                                Case "11D", "13A", "13B", "13C", "36A", "36B"
                                    msg = "Invalid UCR Combination for Victim"
                                    Call ShowApplicableContainers(vucrlist)
'---- setfocus logic ----
'                                             vucrlist.SetFocus
          If vucrlist.Visible Then
              vucrlist.SetFocus
           End If
                                    GoTo exiteditv
                            End Select
                        End If
                    End If
                Next tt%
            Case "11D"
                For tt% = 1 To vucrlist.ListItems.Count
                    If tt% <> t% Then
                        If vucrlist.ListItems(tt%).Selected Then
                            tempvucr2 = Mid$(vucrlist.ListItems(tt%), InStr(vucrlist.ListItems(tt%), "(") + 1, 3)
                            Select Case tempvucr2
                                Case "11A", "11B", "11C", "13A", "13B", "13C", "36A", "36B"
                                    msg = "Invalid UCR Combination for Victim"
                                    Call ShowApplicableContainers(vucrlist)
'---- setfocus logic ----
'                                             vucrlist.SetFocus
          If vucrlist.Visible Then
              vucrlist.SetFocus
           End If
                                    GoTo exiteditv
                            End Select
                        End If
                    End If
                Next tt%
            Case "120"
                For tt% = 1 To vucrlist.ListItems.Count
                    If tt% <> t% Then
                        If vucrlist.ListItems(tt%).Selected Then
                            tempvucr2 = Mid$(vucrlist.ListItems(tt%), InStr(vucrlist.ListItems(tt%), "(") + 1, 3)
                            Select Case tempvucr2
                                Case "13A", "13B", "13C", "23A", "23B", "23C", "23D", "23E", "23F", "23G", "23H", "240"
                                    msg = "Invalid UCR Combination for Victim"
                                    Call ShowApplicableContainers(vucrlist)
'---- setfocus logic ----
'                                             vucrlist.SetFocus
          If vucrlist.Visible Then
              vucrlist.SetFocus
           End If
                                    GoTo exiteditv
                            End Select
                        End If
                    End If
                Next tt%
            Case "13A"
                For tt% = 1 To vucrlist.ListItems.Count
                    If tt% <> t% Then
                        If vucrlist.ListItems(tt%).Selected Then
                            tempvucr2 = Mid$(vucrlist.ListItems(tt%), InStr(vucrlist.ListItems(tt%), "(") + 1, 3)
                            Select Case tempvucr2
                                Case "09A", "09B", "11A", "11B", "11C", "120", "13B", "13C"
                                    msg = "Invalid UCR Combination for Victim"
                                    Call ShowApplicableContainers(vucrlist)
'---- setfocus logic ----
'                                             vucrlist.SetFocus
          If vucrlist.Visible Then
              vucrlist.SetFocus
           End If
                                    GoTo exiteditv
                            End Select
                        End If
                    End If
                Next tt%
            Case "13B"
                For tt% = 1 To vucrlist.ListItems.Count
                    If tt% <> t% Then
                        If vucrlist.ListItems(tt%).Selected Then
                            tempvucr2 = Mid$(vucrlist.ListItems(tt%), InStr(vucrlist.ListItems(tt%), "(") + 1, 3)
                            Select Case tempvucr2
                                Case "09A", "09B", "11A", "11B", "11C", "120", "13A", "13C", "11D"
                                    msg = "Invalid UCR Combination for Victim"
                                    Call ShowApplicableContainers(vucrlist)
'---- setfocus logic ----
'                                             vucrlist.SetFocus
          If vucrlist.Visible Then
              vucrlist.SetFocus
           End If
                                    GoTo exiteditv
                            End Select
                        End If
                    End If
                Next tt%
            Case "13C"
                For tt% = 1 To vucrlist.ListItems.Count
                    If tt% <> t% Then
                        If vucrlist.ListItems(tt%).Selected Then
                            tempvucr2 = Mid$(vucrlist.ListItems(tt%), InStr(vucrlist.ListItems(tt%), "(") + 1, 3)
                            Select Case tempvucr2
                                Case "09A", "09B", "11A", "11B", "11C", "120", "13A", "13B", "11D"
                                    msg = "Invalid UCR Combination for Victim"
                                    Call ShowApplicableContainers(vucrlist)
'---- setfocus logic ----
'                                             vucrlist.SetFocus
          If vucrlist.Visible Then
              vucrlist.SetFocus
           End If
                                    GoTo exiteditv
                            End Select
                        End If
                    End If
                Next tt%
            Case "23A", "23B", "23C", "23D", "23E", "23F", "23G", "23H", "240"
                For tt% = 1 To vucrlist.ListItems.Count
                    If tt% <> t% Then
                        If vucrlist.ListItems(tt%).Selected Then
                            tempvucr2 = Mid$(vucrlist.ListItems(tt%), InStr(vucrlist.ListItems(tt%), "(") + 1, 3)
                            Select Case tempvucr2
                                Case "120"
                                    msg = "Invalid UCR Combination for Victim"
                                    Call ShowApplicableContainers(vucrlist)
'---- setfocus logic ----
'                                             vucrlist.SetFocus
          If vucrlist.Visible Then
              vucrlist.SetFocus
           End If
                                    GoTo exiteditv
                            End Select
                        End If
                    End If
                Next tt%
            Case "36A", "36B"
                For tt% = 1 To vucrlist.ListItems.Count
                    If tt% <> t% Then
                        If vucrlist.ListItems(tt%).Selected Then
                            tempvucr2 = Mid$(vucrlist.ListItems(tt%), InStr(vucrlist.ListItems(tt%), "(") + 1, 3)
                            Select Case tempvucr2
                                Case "11A", "11B", "11C", "11D"
                                    msg = "Invalid UCR Combination for Victim"
                                    Call ShowApplicableContainers(vucrlist)
'---- setfocus logic ----
'                                             vucrlist.SetFocus
          If vucrlist.Visible Then
              vucrlist.SetFocus
           End If
                                    GoTo exiteditv
                            End Select
                        End If
                    End If
                Next tt%
        End Select
    End If
Next t%
'----edit by Ed Sloan 11/30/99 from SCLED 8/9/95 update----------------
For t% = 1 To vucrlist.ListItems.Count
    If vucrlist.ListItems(t%).Selected Then
        tempvucr = Mid$(vucrlist.ListItems(t%), InStr(vucrlist.ListItems(t%), "(") + 1, 3)
        '===== Error 464,465,467
        Select Case tempvucr
                '==== SC edit p11
                Case "120"
                    If Not individual Then
                        msg = "Individual must be selected for Robbery."
                        Call ShowApplicableContainers(individual)
'---- setfocus logic ----
'                                 individual.SetFocus
          If individual.Visible Then
              individual.SetFocus
           End If
                        GoTo exiteditv
                    End If
                Case "13A", "13B", "13C", "09A", "09B", "09C", "100", "11A", "11B", "11C", "11D", "36A", "36B", "36C"
                    If Not individual And Not policeofficer Then
                        msg = "Individual must be selected for Crimes Against Person."
                        Call ShowApplicableContainers(individual)
'---- setfocus logic ----
'                                 individual.SetFocus
          If individual.Visible Then
              individual.SetFocus
           End If
                        GoTo exiteditv
                    End If
                Case "90F"
                    If Not individual And Not policeofficer And Not societypublic Then
                        msg = "Individual or Society must be selected for Family Offenses/Nonviolent."
                        Call ShowApplicableContainers(individual)
'---- setfocus logic ----
'                                 individual.SetFocus
          If individual.Visible Then
              individual.SetFocus
           End If
                        GoTo exiteditv
                    End If
                Case "90Z", "90K", "90N", "90L"
                Case "90B", "90C", "90D", "90G", "90H", "90I", "90E", "90F"
                Case "90J", "36C", "980", "978", "753", "756"
                Case "35A", "35B", "39A", "39B", "39C", "39D", "370", "40A", "40B", "520"
                    If Not societypublic Then
                        msg = "Society must be selected for Crimes Against Society."
                        Call ShowApplicableContainers(societypublic)
'---- setfocus logic ----
'                                 societypublic.SetFocus
          If societypublic.Visible Then
              societypublic.SetFocus
           End If
                        GoTo exiteditv
                    End If
                Case Else
                    If societypublic Then
                        msg = "Society cannot be selected for Crimes Against Property."
                        Call ShowApplicableContainers(societypublic)
'---- setfocus logic ----
'                                 societypublic.SetFocus
          If societypublic.Visible Then
              societypublic.SetFocus
           End If
                        GoTo exiteditv
                    End If
        End Select
        '===== Error 481
        If tempvucr = "36B" And Val(age(1)) > 15 And vucrlist.ListItems(t%).Selected = True Then
            msg = "For statutory rape, the victim must be less than or equal to 15 years of age."
            vucrlist.ListItems(t%).Selected = False
            Call ShowApplicableContainers(vucrlist)
'---- setfocus logic ----
'                     vucrlist.SetFocus
          If vucrlist.Visible Then
              vucrlist.SetFocus
           End If
            GoTo exiteditv
        End If
        '===== SCEdit 8/9/95
        If tempvucr = "23C" Then
            If Len(age(1)) = 4 Then
                If Val(Right$(age(1), 2)) > 15 Then
                    msg = "For Offense 23C, the victim age must be 15 years old or less."
                    Call ShowApplicableContainers(age(1))
'---- setfocus logic ----
'                             age(1).SetFocus
          If age(1).Visible Then
              age(1).SetFocus
           End If
                    GoTo exiteditv
                End If
            Else
            If Len(age(1)) = 2 Then
                If Val(age(1)) > 15 Then
                    msg = "For Offense 23C, the victim age must be 15 years old or less."
                    Call ShowApplicableContainers(age(1))
'---- setfocus logic ----
'                             age(1).SetFocus
          If age(1).Visible Then
              age(1).SetFocus
           End If
                    GoTo exiteditv
                End If
            Else
                msg = "For Offense 23C, the victim age must be 15 years old or less."
                Call ShowApplicableContainers(age(1))
'---- setfocus logic ----
'                         age(1).SetFocus
          If age(1).Visible Then
              age(1).SetFocus
           End If
                GoTo exiteditv
            End If
            End If
        End If
    End If
Next t%
'----------------------------------------------------------------------
'----edit by Ed Sloan 12/03/99 from SCLED 3/5/97 update----------------
Dim ucrexists, valinj As Boolean
ucrexists = False
valinj = False
foundinjtype% = 0
For r% = 1 To vucrlist.ListItems.Count
    If vucrlist.ListItems(r%).Selected Then
        tempvucr = Mid$(vucrlist.ListItems(r%), InStr(vucrlist.ListItems(r%), "(") + 1, 3)
        Select Case tempvucr
            Case "100", "11A", "11B", "11C", "11D", "120", "13A", "13B", "210"
                foundinjtype% = 1
        End Select
    End If
Next r%
For r% = 1 To vucrlist.ListItems.Count
    If vucrlist.ListItems(r%).Selected Then
        tempvucr = Mid$(vucrlist.ListItems(r%), InStr(vucrlist.ListItems(r%), "(") + 1, 3)
        ICT% = 0
        For rr% = 1 To injury.ListItems.Count
            If injury.ListItems(rr%).Selected Then
                ICT% = ICT% + 1
            End If
        Next rr%
        Select Case tempvucr
            '===== Error 479
            Case "13B"
                ucrexists = True
                For q% = 1 To injury.ListItems.Count
                    If injury.ListItems(q%).Selected Then
                        tempinj = Mid$(injury.ListItems(q%), InStr(injury.ListItems(q%), "(") + 1, 1)
                        If (tempinj = "M" Or tempinj = "N") Then
                            validinj = True
                        End If
                    End If
                Next q%
                If ucrexists And Not validinj Then
                    msg = "For simple assault, the only injury types can be minor or none."
                    Call ShowApplicableContainers(vucrlist)
'---- setfocus logic ----
'                             vucrlist.SetFocus
          If vucrlist.Visible Then
              vucrlist.SetFocus
           End If
                    GoTo exiteditv
                End If
            '===== Error 401
            Case "100", "11A", "11B", "11C", "11D", "120", "13A", "13B", "210"
                If ICT% = 0 Then
                    msg = "Type of injury must be selected for UCR " + tempvcur + "."
                    Call ShowApplicableContainers(vucrlist)
'---- setfocus logic ----
'                             vucrlist.SetFocus
          If vucrlist.Visible Then
              vucrlist.SetFocus
           End If
                    GoTo exiteditv
                End If
            '===== Error 419
            Case Else
                If ICT% > 0 And foundinjtype% = 0 Then
                    foundother% = 0
                    For q% = 1 To vucrlist.ListItems.Count
                        If vucrlist.ListItems(q%).Selected Then
                            Select Case Mid$(vucrlist.ListItems(q%), InStr(vucrlist.ListItems(q%), "(") + 1, 3)
                                Case "100", "11A", "11B", "11C", "11D", "120", "13A", "13B", "210"
                                    foundother% = 1
                                    q% = vucrlist.ListItems.Count
                            End Select
                        End If
                    Next q%
                    If foundother% = 0 Then
                        msg = "Type of injury is not applicable for UCR " + tempvcur + "."
                        Call ShowApplicableContainers(injury)
'---- setfocus logic ----
'                                 injury.SetFocus
          If injury.Visible Then
              injury.SetFocus
           End If
                        GoTo exiteditv
                    End If
                End If
        End Select
    End If
Next r%

'===== Error 407
For r% = 1 To injury.ListItems.Count
    If injury.ListItems(r%).Selected Then
        If Mid$(injury.ListItems(r%), InStr(injury.ListItems(r%), "(") + 1, 1) = "N" Then
            For rr% = 1 To r% - 1
                If injury.ListItems(rr%).Selected Then
                    msg = "When Injury Type N=None is selected, no other values may be selected."
                    Call ShowApplicableContainers(injury)
'---- setfocus logic ----
'                             injury.SetFocus
          If injury.Visible Then
              injury.SetFocus
           End If
                    GoTo exiteditv
                End If
            Next rr%
            For rr% = r% + 1 To injury.ListItems.Count
                If injury.ListItems(rr%).Selected Then
                    msg = "When Injury Type N=None is selected, no other values may be selected."
                    Call ShowApplicableContainers(injury)
'---- setfocus logic ----
'                             injury.SetFocus
          If injury.Visible Then
              injury.SetFocus
           End If
                    GoTo exiteditv
                End If
            Next rr%
        End If
    End If
Next r%
For t% = 0 To 4
    UCRLIST(t%).ListIndex = -1
    For tt% = 0 To UCRLIST(t%).ListCount - 1
        If UCRLIST(t%).Selected(tt%) = True Then
            UCRLIST(t%).ListIndex = tt%
            tt% = UCRLIST(t%).ListCount - 1
        End If
    Next tt%
    If UCRLIST(t%).ListIndex > -1 Then
        tempucr = Mid$(UCRLIST(t%).List(UCRLIST(t%).ListIndex), InStr(UCRLIST(t%).List(UCRLIST(t%).ListIndex), "(") + 1, 3)
        If tempucr = "13A" Or tempucr = "13B" Or tempucr = "13C" Then
            If UCase(vsname(2)) <> "UNKNOWN" Then
                If relationship(10).ListIndex = -1 And relationship(11).ListIndex = -1 And relationship(12).ListIndex = -1 And relationship(13).ListIndex = -1 And relationship(14).ListIndex = -1 And relationship(15).ListIndex = -1 And relationship(16).ListIndex = -1 And relationship(17).ListIndex = -1 And relationship(18).ListIndex = -1 And relationship(19).ListIndex = -1 Then
                    msg = "A relationship to subject must be selected."
                    Call ShowApplicableContainers(UCRLIST(t%))
'---- setfocus logic ----
'                             UCRLIST(t%).SetFocus
          If UCRLIST(t%).Visible Then
              UCRLIST(t%).SetFocus
           End If
                    GoTo exiteditv
                End If
            End If
        End If
        '===== Additional F 12
        '===== Data Element 7, 13
        '===== eRROR 458
        Select Case tempucr
            Case "09A", "09B", "09C", "100", "11A", "11B", "11C", "11D", "120", "13A", "13B", "13C", "36A", "36B", "36C"
                If UCase(vsname(2)) <> "UNKNOWN" Then
                    If relationship(10).ListIndex = -1 And relationship(11).ListIndex = -1 And relationship(12).ListIndex = -1 And relationship(13).ListIndex = -1 And relationship(14).ListIndex = -1 And relationship(15).ListIndex = -1 And relationship(16).ListIndex = -1 And relationship(17).ListIndex = -1 And relationship(18).ListIndex = -1 And relationship(19).ListIndex = -1 Then
                        msg = "A relationship to subject must be selected."
                        Call ShowApplicableContainers(relationship(10))
'---- setfocus logic ----
'                                 relationship(10).SetFocus
          If relationship(10).Visible Then
              relationship(10).SetFocus
           End If
                        GoTo exiteditv
                    End If
                End If
        End Select
        '===== Additional F 13
        '===== Data Element 7
        If tempucr = "100" Then
            If UCase(vsname(2)) <> "UNKNOWN" Then
                If relationship(10).ListIndex = -1 And relationship(11).ListIndex = -1 And relationship(12).ListIndex = -1 And relationship(13).ListIndex = -1 And relationship(14).ListIndex = -1 And relationship(15).ListIndex = -1 And relationship(16).ListIndex = -1 And relationship(17).ListIndex = -1 And relationship(18).ListIndex = -1 And relationship(19).ListIndex = -1 Then
                    msg = "A relationship to subject must be selected."
                    Call ShowApplicableContainers(relationship(10))
'---- setfocus logic ----
'                             relationship(10).SetFocus
          If relationship(10).Visible Then
              relationship(10).SetFocus
           End If
                    GoTo exiteditv
                End If
            End If
        End If
        '===== Additional F 18
        If tempucr = "120" Then
            If individual Then
            If UCase(vsname(2)) <> "UNKNOWN" Then
                If relationship(10).ListIndex = -1 And relationship(11).ListIndex = -1 And relationship(12).ListIndex = -1 And relationship(13).ListIndex = -1 And relationship(14).ListIndex = -1 And relationship(15).ListIndex = -1 And relationship(16).ListIndex = -1 And relationship(17).ListIndex = -1 And relationship(18).ListIndex = -1 And relationship(19).ListIndex = -1 Then
                    msg = "A relationship to subject must be selected."
                    Call ShowApplicableContainers(relationship(10))
'---- setfocus logic ----
'                             relationship(10).SetFocus
          If relationship(10).Visible Then
              relationship(10).SetFocus
           End If
                    GoTo exiteditv
                End If
            End If
            End If
        End If
        '===== Additional F 19
        If tempucr = "11A" Or tempucr = "11B" Or tempucr = "11C" Or tempucr = "11D" Then
            If UCase(vsname(2)) <> "UNKNOWN" Then
                If relationship(10).ListIndex = -1 And relationship(11).ListIndex = -1 And relationship(12).ListIndex = -1 And relationship(13).ListIndex = -1 And relationship(14).ListIndex = -1 And relationship(15).ListIndex = -1 And relationship(16).ListIndex = -1 And relationship(17).ListIndex = -1 And relationship(18).ListIndex = -1 And relationship(19).ListIndex = -1 Then
                    msg = "A relationship to subject must be selected."
                    Call ShowApplicableContainers(relationship(10))
'---- setfocus logic ----
'                             relationship(10).SetFocus
          If relationship(10).Visible Then
              relationship(10).SetFocus
           End If
                    GoTo exiteditv
                End If
            End If
        End If
        '===== Additional F 20
        If tempucr = "36A" Or tempucr = "36B" Then
            If UCase(vsname(2)) <> "UNKNOWN" Then
                If relationship(10).ListIndex = -1 And relationship(11).ListIndex = -1 And relationship(12).ListIndex = -1 And relationship(13).ListIndex = -1 And relationship(14).ListIndex = -1 And relationship(15).ListIndex = -1 And relationship(16).ListIndex = -1 And relationship(17).ListIndex = -1 And relationship(18).ListIndex = -1 And relationship(19).ListIndex = -1 Then
                    msg = "A relationship to subject must be selected."
                    Call ShowApplicableContainers(relationship(10))
'---- setfocus logic ----
'                             relationship(10).SetFocus
          If relationship(10).Visible Then
              relationship(10).SetFocus
           End If
                    GoTo exiteditv
                End If
            End If
        End If
    End If
Next t%

GoTo goodeditv
exiteditv:
editerr = 1
goodeditv:
Exit Sub
'RLB Bandaid
rlbErrv:
'    If Err.Number = 5 Then
        Resume Next
End Sub
Friend Sub editsubject(editerr, msg As String)
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwi + "incident.mdb")
'===== SC LEOKA
If policeofficer Then
    If TWOMANVEHICLE = 0 And ONEMANVEHICLE = 0 And DETECTIVE = 0 And TODOTHER = 0 Then
        msg = "If Type Victim is Police Officer, a selection must be made for Two Man Vehicle, One Man Vehicle, Detective/Special Assignment, or Other."
        Call ShowApplicableContainers(policeofficer)
'---- setfocus logic ----
'                 policeofficer.SetFocus
          If policeofficer.Visible Then
              policeofficer.SetFocus
           End If
        GoTo exitedits
    End If
    If Not TWOMANVEHICLE Then
        If ALONE = 0 And ASSISTED = 0 Then
            msg = "If Type Victim is Police Officer and not Two Man Vehicle, a selection must be made for Alone or Assisted."
            Call ShowApplicableContainers(TWOMANVEHICLE)
'---- setfocus logic ----
'                     TWOMANVEHICLE.SetFocus
          If TWOMANVEHICLE.Visible Then
              TWOMANVEHICLE.SetFocus
           End If
            GoTo exitedits
        End If
    End If
End If
'===== Error 761
If RUNAWAY = 1 Then
    If Len(age(2)) = 2 Then
        If Val(age(2)) > 17 Then
            msg = "A runaway must be under the age of 18."
            Call ShowApplicableContainers(RUNAWAY)
'---- setfocus logic ----
'                     RUNAWAY.SetFocus
          If RUNAWAY.Visible Then
              RUNAWAY.SetFocus
           End If
            GoTo exitedits
        End If
    Else
    If Len(age(2)) = 4 Then
        If Val(Right$(age(2), 2)) > 17 Then
            msg = "A runaway must be under the age of 18."
            Call ShowApplicableContainers(age(2))
'---- setfocus logic ----
'                     age(2).SetFocus
          If age(2).Visible Then
              age(2).SetFocus
           End If
            GoTo exitedits
        End If
    Else
        msg = "A runaway must be under the age of 18."
        Call ShowApplicableContainers(age(2))
'---- setfocus logic ----
'                 age(2).SetFocus
          If age(2).Visible Then
              age(2).SetFocus
           End If
        GoTo exitedits
    End If
    End If
End If
'===== Error 665
If IsDate(DATEOFARREST) And IsDate(incidentdate(0)) Then
    If CVDate(incidentdate(0)) > CVDate(DATEOFARREST) Then
        msg = "Date of Arrest cannot be before incident date."
        Call ShowApplicableContainers(incidentdate(0))
'---- setfocus logic ----
'                 incidentdate(0).SetFocus
          If incidentdate(0).Visible Then
              incidentdate(0).SetFocus
           End If
        GoTo exitedits
    End If
End If
If age(2) > "" Then
    If Val(age(2)) = 0 And age(2) <> "00" Then
        msg = "Age must be entered in the format of a single age: X or XX (i.e. 3 or 42) or in the format of a range of ages: XXXX (i.e. 2025, meaning 20 - 25 years old)."
        Call ShowApplicableContainers(age(2))
'---- setfocus logic ----
'                 age(2).SetFocus
          If age(2).Visible Then
              age(2).SetFocus
           End If
        GoTo exitedits
    End If
Else
    '===== error 504
    msg = "Subject age must be entered. (00 = unknown)"
    Call ShowApplicableContainers(age(2))
'---- setfocus logic ----
'             age(2).SetFocus
          If age(2).Visible Then
              age(2).SetFocus
           End If
    GoTo exitedits
End If
'===== Error 504
If sex(2).ListIndex = -1 Then
    msg = "A value for Sex in subject data must be entered."
    Call ShowApplicableContainers(sex(2))
'---- setfocus logic ----
'             sex(2).SetFocus
          If sex(2).Visible Then
              sex(2).SetFocus
           End If
    GoTo exitedits
End If
If race(2).ListIndex = -1 Then
    msg = "A value for race in subject data must be entered."
    Call ShowApplicableContainers(race(2))
'---- setfocus logic ----
'             race(2).SetFocus
          If race(2).Visible Then
              race(2).SetFocus
           End If
    GoTo exitedits
End If
If ethnicity(2).ListIndex = -1 Then
    msg = "A value for ethnicity in subject data must be entered."
    Call ShowApplicableContainers(ethnicity(2))
'---- setfocus logic ----
'             ethnicity(2).SetFocus
          If ethnicity(2).Visible Then
              ethnicity(2).SetFocus
           End If
    GoTo exitedits
End If
'===== Error 404,556
If age(2) <> "NN" And age(2) <> "NB" And age(2) <> "BB" Then
    For t% = 1 To Len(age(2))
        If InStr("0123456789-", Mid$(age(2), t%, 1)) = 0 Then
            msg = "An invalid entry in age has been found.  Valid Entry Formats:  X, XX, X-Y, XX-Y, X-YY, XX-YY, XXYY"
            t% = Len(age(2))
            Call ShowApplicableContainers(age(2))
'---- setfocus logic ----
'                     age(2).SetFocus
          If age(2).Visible Then
              age(2).SetFocus
           End If
            GoTo exitedits
        End If
    Next t%
End If
If InStr(age(2), "-") > 0 Then
    ag1 = Val(Left$(age(2), InStr(age(2), "-") - 1))
    ag2 = Val(Mid$(age(2), InStr(age(2), "-") + 1))
    age(2) = Format$(ag1, "00") + Format$(ag2, "00")
End If
'===== Error 410,422,509,510,522
If Len(age(2)) = 4 Then
    If Val(Left$(age(2), 2)) >= Val(Right$(age(2), 2)) Then
        msg = "For an age range, the first age must be less than the second age."
        Call ShowApplicableContainers(age(2))
        GoTo exitedits
    End If
    If Val(Left$(age(2), 2)) = 0 Then
        msg = "The low value in an age range cannot be 0."
        Call ShowApplicableContainers(age(2))
'---- setfocus logic ----
'                 age(2).SetFocus
          If age(2).Visible Then
              age(2).SetFocus
           End If
        GoTo exitedits
    End If
End If
editsubjectdetail:
'===== Data Element 20, 21, 22
Dim drugs(5)
drugs(0) = ""
drugs(1) = ""
drugs(2) = ""
drugs(3) = ""
drugs(4) = ""
drugs(5) = ""
dct% = 3
For ii% = 3 To 5
    If drugtype(ii%).ListIndex > -1 Then
        drugs(dct%) = drugtype(ii%).List(drugtype(ii%).ListIndex)
        dct% = dct% + 1
    End If
Next ii%
If dct% > 0 And dct% > 3 Then
    For Z% = 3 To dct%
        '===== Error 306
        If Z% > 0 Then
            For ZZ% = 0 To Z% - 1
                If drugs(ZZ%) = drugs(Z%) Then
                    If InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "GM=") > 0 Or _
                       InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "KG=") > 0 Or _
                       InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "OZ=") > 0 Or _
                       InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "LB=") > 0 Then
                        If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "GM=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "KG=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "OZ=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "LB=") > 0 Then
                                msg = "Two different measurements within the same weight category cannot be entered for the same type of drug."
                                Call ShowApplicableContainers(drugmeasurement(ZZ%))
'---- setfocus logic ----
'                                         drugmeasurement(zz%).SetFocus
          If drugmeasurement(ZZ%).Visible Then
              drugmeasurement(ZZ%).SetFocus
           End If
                                GoTo exitedits
                        End If
                    End If
                    If InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "ML=") > 0 Or _
                       InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "LT=") > 0 Or _
                       InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "FO=") > 0 Or _
                       InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "GL=") > 0 Then
                        If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "ML=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "LT=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "FO=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "GL=") > 0 Then
                                msg = "Two different measurements within the same weight category cannot be entered for the same type of drug."
                                Call ShowApplicableContainers(drugmeasurement(ZZ%))
'---- setfocus logic ----
'                                         drugmeasurement(zz%).SetFocus
          If drugmeasurement(ZZ%).Visible Then
              drugmeasurement(ZZ%).SetFocus
           End If
                                GoTo exitedits
                        End If
                    End If
                    If InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "DU=") > 0 Or _
                       InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "NP=") > 0 Then
                        If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "DU=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "NP=") > 0 Then
                                msg = "Two different measurements within the same weight category cannot be entered for the same type of drug."
                                Call ShowApplicableContainers(drugmeasurement(ZZ%))
'---- setfocus logic ----
'                                         drugmeasurement(zz%).SetFocus
          If drugmeasurement(ZZ%).Visible Then
              drugmeasurement(ZZ%).SetFocus
           End If
                                GoTo exitedits
                        End If
                    End If
                End If
            Next ZZ%
        End If
        For t% = 1 To Len(drugamt(Z%))
            If InStr("0123456789.", Mid$(drugamt(Z%), t%, 1)) = 0 Then
                msg = "Only 0123456789. are allowed in drug amount.  Enter all fractions as decimals (i.e. 1/2 = .5)."
                Call ShowApplicableContainers(drugamt(Z%))
'---- setfocus logic ----
'                         drugamt(Z%).SetFocus
          If drugamt(Z%).Visible Then
              drugamt(Z%).SetFocus
           End If
                GoTo exitedits
            End If
        Next t%
        '=====SCEdit 4/21/92 P29  Overturned drug amt and measurement error
        'If drugamt(z%) = "" Or drugmeasurement(z%).ListIndex = -1 Then
        '    If drugs(z%) > "" And Left$(drugs(z%), 1) <> "X" And Left$(drugs(z%), 1) <> "U" Then
        '        msg = "Drug Quantity and Measurement Type must be entered/selected."
        '        GoTo exitedits
        '    End If
        'End If
        '===== Error 366
        If drugamt(Z%) > "" Then
            If drugs(Z%) = "" Or drugmeasurement(Z%).ListIndex = -1 Then
                msg = "If a drug quantity is entered, then drug type and measurement type must also be entered."
                Call ShowApplicableContainers(drugamt(Z%))
'---- setfocus logic ----
'                         drugamt(Z%).SetFocus
          If drugamt(Z%).Visible Then
              drugamt(Z%).SetFocus
           End If
                GoTo exitedits
            End If
        End If
        '===== Error 367
        If drugmeasurement(Z%).ListIndex > -1 Then
            If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "NP=") > 0 Then
                If Left$(drugs(Z%), 1) <> "E" And Left$(drugs(Z%), 1) <> "G" And Left$(drugs(Z%), 1) <> "K" Then
                    msg = "Number of Plants is only allowabled with Marijuana, Opium, and Other Hallucinogens."
                    Call ShowApplicableContainers(drugmeasurement(Z%))
'---- setfocus logic ----
'                             drugmeasurement(Z%).SetFocus
          If drugmeasurement(Z%).Visible Then
              drugmeasurement(Z%).SetFocus
           End If
                    GoTo exitedits
                End If
            End If
            '===== Error 384
            If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "XX=") > 0 Then
                If Val(drugamt(Z%)) <> 1 Then
                    msg = "If drug measurement is NOT REPORTED, drug amount must be 1."
                    Call ShowApplicableContainers(drugamt(Z%))
'---- setfocus logic ----
'                             drugamt(Z%).SetFocus
          If drugamt(Z%).Visible Then
              drugamt(Z%).SetFocus
           End If
                    GoTo exitedits
                End If
            End If
            '===== Error 368
            If drugs(Z%) = "" Or drugamt(Z%) = "" Then
                msg = "If a drug measurement is entered, then drug type and quantity must also be entered."
                Call ShowApplicableContainers(drugamt(Z%))
'---- setfocus logic ----
'                         drugamt(Z%).SetFocus
          If drugamt(Z%).Visible Then
              drugamt(Z%).SetFocus
           End If
                GoTo exitedits
            End If
        End If
        '===== Error 362
        If Left$(drugs(Z%), 1) = "X" Then
            If drugtype(3).ListIndex = -1 Or drugtype(4).ListIndex = -1 Or drugtype(5).ListIndex = -1 Then
                msg = "If X = Over 3 Drug Types is entered, then the other 2 drug types must also be entered."
                Call ShowApplicableContainers(drugtype(3))
'---- setfocus logic ----
'                         drugtype(3).SetFocus
          If drugtype(3).Visible Then
              drugtype(3).SetFocus
           End If
                GoTo exitedits
            End If
            '===== Error 363
            If drugamt(Z%) > "" Or drugmeasurement(Z%).ListIndex > -1 Then
                msg = "Drug Amount and Measurement are not valid entries for X=Over 3 Drug Types."
                Call ShowApplicableContainers(drugamt(Z%))
'---- setfocus logic ----
'                         drugamt(Z%).SetFocus
          If drugamt(Z%).Visible Then
              drugamt(Z%).SetFocus
           End If
                GoTo exitedits
            End If
        End If
    Next Z%
End If
For pp% = 10 To 19
    temprel = Mid$(relationship(pp%).List(relationship(pp%).ListIndex), InStr(relationship(pp%).List(relationship(pp%).ListIndex), "(") + 1, 2)
    If temprel = "CH" Or temprel = "GC" Or temprel = "SC" Then
        If Val(age(2)) < Val(age(1)) Then
            msg = "The relationship of victim to subject cannot be 'PA', 'GP' or 'SP' when victim's age is less than subject's age."
            Call ShowApplicableContainers(age(2))
'---- setfocus logic ----
'                     age(2).SetFocus
          If age(2).Visible Then
              age(2).SetFocus
           End If
            GoTo exitedits
        End If
    End If
    If temprel = "PA" Or temprel = "GP" Or temprel = "SP" Then
        If Val(age(1)) < Val(age(2)) Then
            msg = "The relationship of victim to subject cannot be 'PA', 'GP' or 'SP' when victim's age is less than subject's age."
            Call ShowApplicableContainers(age(1))
'---- setfocus logic ----
'                     age(1).SetFocus
          If age(1).Visible Then
              age(1).SetFocus
           End If
            GoTo exitedits
        End If
    End If
Next pp%
'==== Mandatories E - 37, 38, 39
'===== Error 501
'If individual And UCase(vsname(2)) <> "UNKNOWN" Then
If UCase(vsname(2)) <> "UNKNOWN" Then
    If Val(age(2)) = 0 And age(2) <> "00" Then
        msg = "Invalid age entered."
        Call ShowApplicableContainers(age(2))
'---- setfocus logic ----
'                 age(2).SetFocus
          If age(2).Visible Then
              age(2).SetFocus
           End If
        GoTo exitedits
    End If
    If sex(2).ListIndex = -1 Then
        msg = "Invalid sex entered."
        Call ShowApplicableContainers(sex(2))
'---- setfocus logic ----
'                 sex(2).SetFocus
          If sex(2).Visible Then
              sex(2).SetFocus
           End If
        GoTo exitedits
    End If
    If race(2).ListIndex = -1 Then
        msg = "Invalid race entered."
        Call ShowApplicableContainers(race(2))
'---- setfocus logic ----
'                 race(2).SetFocus
          If race(2).Visible Then
              race(2).SetFocus
           End If
        GoTo exitedits
    End If
    If ethnicity(2).ListIndex = -1 Then
        msg = "Invalid ETHNICITY entered."
        Call ShowApplicableContainers(ethnicity(2))
'---- setfocus logic ----
'                 ethnicity(2).SetFocus
          If ethnicity(2).Visible Then
              ethnicity(2).SetFocus
           End If
        GoTo exitedits
    End If
End If
'==== Mandatories E - 48, 49
'===== Mandatories F - 41, 42, 43, 44, 45, 46, 47, 48, 49, 52

If DATEOFARREST > "" Then
    If Not IsDate(DATEOFARREST) Then
        msg = "Valid date of arrest must be entered."
        Call ShowApplicableContainers(DATEOFARREST)
'---- setfocus logic ----
'                 DATEOFARREST.SetFocus
          If DATEOFARREST.Visible Then
              DATEOFARREST.SetFocus
           End If
        GoTo exitedits
    End If
    '===== Error 601,701
    If Val(age(2)) = 0 And age(2) <> "00" Then
        msg = "Invalid age entered."
        Call ShowApplicableContainers(age(2))
'---- setfocus logic ----
'                 age(2).SetFocus
          If age(2).Visible Then
              age(2).SetFocus
           End If
        GoTo exitedits
    End If
    If sex(2).ListIndex = -1 Then
        msg = "Invalid sex entered."
        Call ShowApplicableContainers(sex(2))
'---- setfocus logic ----
'                 sex(2).SetFocus
          If sex(2).Visible Then
              sex(2).SetFocus
           End If
        GoTo exitedits
    End If
    If race(2).ListIndex = -1 Then
        msg = "Invalid race entered."
        Call ShowApplicableContainers(race(2))
'---- setfocus logic ----
'                 race(2).SetFocus
          If race(2).Visible Then
              race(2).SetFocus
           End If
        GoTo exitedits
    End If
    '===== Error 665
    If CVDate(incidentdate(0)) > CVDate(DATEOFARREST) Then
        msg = "Date of Arrest cannot be before incident date."
        Call ShowApplicableContainers(incidentdate(0))
'---- setfocus logic ----
'                 incidentdate(0).SetFocus
          If incidentdate(0).Visible Then
              incidentdate(0).SetFocus
           End If
        GoTo exitedits
    End If
End If
'===== SCEdit 4/21/92 P28
If (offenderdeath Or noprosecution Or extraditiondenied Or victimdeclinescooperation Or juvenilenocustody) Then
    If age(2) = "00" Then
        msg = "For an exceptional clearance, the subject's age (other than 00) must be selected."
        Call ShowApplicableContainers(age(2))
'---- setfocus logic ----
'                 age(2).SetFocus
          If age(2).Visible Then
              age(2).SetFocus
           End If
        GoTo exitedits
    End If
    If race(2).ListIndex = -1 Or race(2).List(race(2).ListIndex) = "Unknown" Then
        msg = "For an exceptional clearance, the subject's race (other than unknown) must be selected."
        Call ShowApplicableContainers(race(2))
'---- setfocus logic ----
'                 race(2).SetFocus
          If race(2).Visible Then
              race(2).SetFocus
           End If
        GoTo exitedits
    End If
    If sex(2).ListIndex = -1 Or sex(2).List(sex(2).ListIndex) = "Unknown" Then
        msg = "For an exceptional clearance, the subject's sex (other than unknown) must be selected."
        Call ShowApplicableContainers(sex(2))
'---- setfocus logic ----
'                 sex(2).SetFocus
          If sex(2).Visible Then
              sex(2).SetFocus
           End If
        GoTo exitedits
    End If
    If ethnicity(2).ListIndex = -1 Or ethnicity(2).List(ethnicity(2).ListIndex) = "Unknown" Then
        msg = "For an exceptional clearance, the subject's ethnicity (other than unknown) must be selected."
        Call ShowApplicableContainers(ethnicity(2))
'---- setfocus logic ----
'                 ethnicity(2).SetFocus
          If ethnicity(2).Visible Then
              ethnicity(2).SetFocus
           End If
        GoTo exitedits
    End If
End If
GoTo goodedits
exitedits:
editerr = 1
goodedits:
On Error Resume Next
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub

Private Sub TOTALARRESTED_Change()
ichanged = True

End Sub

Private Sub totalvalue_Change(Index As Integer)
ichanged = True

End Sub

Private Sub totalvalue_DblClick(Index As Integer)
If Index >= 18 And Index <= 23 Then
    holdrecv = Index
    DATERECOVERED(Index Mod 6).Visible = True
'---- setfocus logic ----
'             DATERECOVERED(index Mod 6).SetFocus
          If DATERECOVERED(indexMod6).Visible Then
              DATERECOVERED(indexMod6).SetFocus
           End If
    numvehicle((Index Mod 6) + 6).Visible = True
'---- setfocus logic ----
'             numvehicle((index Mod 6) + 6).SetFocus
          If numvehicle((indexMod6) + 6).Visible Then
              numvehicle((indexMod6) + 6).SetFocus
           End If
End If
'=====Data Item 18 and 19
If (Index >= 0 And Index <= 5) Then
    If Mid$(pucrlist(Index).List(pucrlist(Index).ListIndex), InStr(pucrlist(Index).List(pucrlist(Index).ListIndex), "(") + 1, 3) = "240" Then
        tempgroup = Val(Mid$(group(Index).List(group(Index).ListIndex), InStr(group(Index).List(group(Index).ListIndex), "(") + 1, 2))
        If (tempgroup = 3 Or tempgroup = 5 Or tempgroup = 24 Or tempgroup = 28 Or tempgroup = 37) And fromfind = 0 Then
            numvehicle(Index Mod 6).Visible = True
'---- setfocus logic ----
'                     numvehicle(index Mod 6).SetFocus
          If numvehicle(indexMod6).Visible Then
              numvehicle(indexMod6).SetFocus
           End If
        End If
    End If
End If


End Sub

Private Sub totalvalue_GotFocus(Index As Integer)
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If

End Sub

Private Sub totalvalue_LostFocus(Index As Integer)
If fromfind = 1 Then
    Exit Sub
End If
If BACKTAB = 0 Then
If Val(totalvalue(Index)) > 0 And Index >= 18 And Index <= 23 Then
    holdrecv = Index
    DATERECOVERED(Index Mod 6).Visible = True
'---- setfocus logic ----
'             DATERECOVERED(index Mod 6).SetFocus
          If DATERECOVERED(indexMod6).Visible Then
              DATERECOVERED(indexMod6).SetFocus
           End If
Else
'=====Data Item 18 and 19
If Val(totalvalue(Index)) > 0 And (Index >= 0 And Index <= 5) Then
    If Mid$(pucrlist(Index).List(pucrlist(Index).ListIndex), InStr(pucrlist(Index).List(pucrlist(Index).ListIndex), "(") + 1, 3) = "240" Then
        tempgroup = Val(Mid$(group(Index).List(group(Index).ListIndex), InStr(group(Index).List(group(Index).ListIndex), "(") + 1, 2))
        If (tempgroup = 3 Or tempgroup = 5 Or tempgroup = 24 Or tempgroup = 28 Or tempgroup = 37) And fromfind = 0 Then
            numvehicle(Index Mod 6).Visible = True
'---- setfocus logic ----
'                     numvehicle(index Mod 6).SetFocus
          If numvehicle(indexMod6).Visible Then
              numvehicle(indexMod6).SetFocus
           End If
        End If
    End If
End If
End If
Else
    BACKTAB = 0
End If
Call figure
End Sub


Private Sub TVBURNED_Click()
ichanged = True

End Sub

Private Sub TVCOUNTERFEIT_Click()
ichanged = True

End Sub

Private Sub TVDAMAGED_Click()
ichanged = True

End Sub

Private Sub TVRECOVERED_Click()
ichanged = True

End Sub

Private Sub TVSEIZED_Click()
ichanged = True

End Sub

Private Sub TVSTOLEN_Click()
ichanged = True

End Sub

Private Sub TVUNKNOWN_Click()
ichanged = True

End Sub

Private Sub TWOMANVEHICLE_Click()
ichanged = True

End Sub

Private Sub TWOMANVEHICLE_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub ucrlist_Click(Index As Integer)
ichanged = True

If FROMKEY = 1 Then
    FROMKEY = 0
    Exit Sub
End If
Dim itmx As ListItem
If UCRLIST(Index).ListIndex = -1 Then
    Exit Sub
End If
UCRLIST(Index).Visible = False
If fromfind = 1 Then
    Exit Sub
End If
On Error Resume Next
'---- setfocus logic ----
'         Command12(index).SetFocus
          If Command12(Index).Visible Then
              Command12(Index).SetFocus
           End If
On Error GoTo 0
End Sub

Private Sub UCRLIST_ItemCheck(Index As Integer, Item As Integer)
a = 1
End Sub

Private Sub UCRLIST_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode <> Asc(" ") Then
    FROMKEY = 1
End If
End Sub

Private Sub UCRLIST_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    UCRLIST(Index).Visible = False
    If fromfind = 0 Then
'---- setfocus logic ----
'                 Command12(index).SetFocus
          If Command12(Index).Visible Then
              Command12(Index).SetFocus
           End If
    End If
Else
If KeyAscii <> Asc(" ") Then
    FROMKEY = 1
End If
End If
    
End Sub

Private Sub UCRLIST_LostFocus(Index As Integer)
If FROMKEY = 1 Then
    FROMKEY = 0
    Exit Sub
End If
Dim itmx As ListItem
If UCRLIST(Index).ListIndex = -1 Then
    Exit Sub
End If
UCRLIST(Index).Visible = False
If fromfind = 1 Then
    Exit Sub
End If
If BACKTAB = 0 Then
'---- setfocus logic ----
'             Command12(index).SetFocus
          If Command12(Index).Visible Then
              Command12(Index).SetFocus
           End If
Else
    BACKTAB = 0
End If
End Sub

Private Sub UCRLIST_Validate(Index As Integer, Cancel As Boolean)
a = 1
End Sub

Private Sub unfounded_Click()
ichanged = True

End Sub

Private Sub unknown_Click()
ichanged = True

End Sub

Private Sub victimdeclinescooperation_Click()
ichanged = True

If fromfind = 1 Then
    Exit Sub
End If

EXCEPTIONALCLEARANCEDATE.Visible = True
'---- setfocus logic ----
'         EXCEPTIONALCLEARANCEDATE.SetFocus
          If EXCEPTIONALCLEARANCEDATE.Visible Then
              EXCEPTIONALCLEARANCEDATE.SetFocus
           End If

End Sub




Private Sub victimdeclinescooperation_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 And Shift = 0 Then
    victimdeclinescooperation = False
End If

End Sub

Private Sub VISIBLEINJURYNO_Click()
ichanged = True

End Sub

Private Sub VISIBLEINJURYNO_GotFocus()
If Frame1.Top > (-1 * Picture2.Top) And Frame1.Top + Frame1.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Frame1.Top > 500 Then
    VScroll1 = Frame1.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub VISIBLEINJURYYES_Click()
ichanged = True

End Sub

Private Sub VISIBLEINJURYYES_GotFocus()
If Frame1.Top > (-1 * Picture2.Top) And Frame1.Top + Frame1.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Frame1.Top > 500 Then
    VScroll1 = Frame1.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub VScroll1_Change()
Picture2.Top = -1 * VScroll1.Value
End Sub

Private Sub vsname_Change(Index As Integer)
ichanged = True

End Sub

Private Sub vsname_Click(Index As Integer)
If vsname(Index) > "" Then
    Call FILLDATA(Index)
    Call setpopup(vsname(Index), "L")
End If
End Sub

Private Sub vsname_GotFocus(Index As Integer)
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If

End Sub

Private Sub vsname_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 And Shift = vbCtrlMask Then
    If nametype = 0 Then
        cleanup.fname = Me.ActiveControl.Text
        cleanup.lname = ""
    Else
        cleanup.lname = Me.ActiveControl.Text
        cleanup.fname = ""
    End If
    cleanup.Show
End If
End Sub

Private Sub vsname_LostFocus(Index As Integer)
If Len(vsname(Index)) > 60 Then
    msg = MsgBox("A maximum of 60 characters is allowed for name.  Your entry is being truncated to 60 characters.", 48, "Genesis INformation Log")
    vsname(Index) = Left$(vsname(Index), 60)
End If
If individual And UCase(vsname(Index)) <> "SAME AS VICTIM" And vsname(Index) > "" And InStr(vsname(Index), ",") = 0 And UCase(vsname(Index)) <> "UNKNOWN" Then
    msg = MsgBox("All names in the Incident Report system should be entered in the format last name + comma + firstname.", 48, "Invalid Data Format")
'    vsname(index).SetFocus
End If
If BACKTAB = 0 Then
If Index = 1 And vsname(1) = vsname(0) Then
    For ty% = 0 To 9
        For tv% = 0 To relationship(ty%).ListCount - 1
            If relationship(ty%).Selected(tv%) = True Then
                relationship(ty% + 10).ListIndex = tv%
                tv% = relationship(ty%).ListCount - 1
            End If
        Next tv%
    Next ty%
    resident(1).ListIndex = resident(0).ListIndex
    For t% = 10 To 19
        relationship(t%).ListIndex = relationship(t% - 10).ListIndex
    Next t%
    race(1).ListIndex = race(0).ListIndex
    sex(1).ListIndex = sex(0).ListIndex
    age(1) = age(0)
    ethnicity(1).ListIndex = ethnicity(0).ListIndex
    HOMEDAYPHONE(1) = HOMEDAYPHONE(0)
    HOMENIGHTPHONE(1) = HOMENIGHTPHONE(0)
    WORKDAYPHONE(1) = WORKDAYPHONE(0)
    WORKNIGHTPHONE(1) = WORKNIGHTPHONE(0)
    address(1) = address(0)
    city(1) = city(0)
    state(1) = state(0)
    zipcode(1) = zipcode(0)
    LOCATIONNUMBER(1) = LOCATIONNUMBER(0)
    If fromfind = 0 Then
'---- setfocus logic ----
'                 ht(0).SetFocus
          If ht(0).Visible Then
              ht(0).SetFocus
           End If
    End If
Else
If vsname(Index) > "" Then
    Call FILLDATA(Index)
End If
End If
Else
    BACKTAB = 0
End If
End Sub


Private Sub vucrlist_ItemClick(ByVal Item As MSComctlLib.ListItem)
ichanged = True
For p% = 1 To 5
    VUCRSEL(p%) = ""
Next p%
vidx% = 0
For p% = 1 To vucrlist.ListItems.Count
    If vucrlist.ListItems(p%).Selected Then
        vidx% = vidx% + 1
        If vidx% < 6 Then
            VUCRSEL(vidx%) = Mid(vucrlist.ListItems(p%), InStr(vucrlist.ListItems(p%), "(") + 1, 3)
        End If
    End If
Next p%
End Sub

Private Sub vucrlist_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

For p% = 1 To 5
    VUCRSEL(p%) = ""
Next p%
vidx% = 0
For p% = 1 To vucrlist.ListItems.Count
    If vucrlist.ListItems(p%).Selected Then
        vidx% = vidx% + 1
        If vidx% < 6 Then
            VUCRSEL(vidx%) = Mid(vucrlist.ListItems(p%), InStr(vucrlist.ListItems(p%), "(") + 1, 3)
        End If
    End If
Next p%
End Sub

Private Sub WANTED_Click()
ichanged = True

End Sub

Private Sub WANTED_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If

End Sub

Private Sub warrant_Click()
ichanged = True

End Sub

Private Sub WARRANT_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If

End Sub

Private Sub weapontype_ItemClick(ByVal Item As MSComctlLib.ListItem)
ichanged = True


If FROMKEY = 1 Then
    FROMKEY = 0
    Exit Sub
End If
IIDX = 0
CTW% = 0
If Item > "" Then
    For t% = 1 To weapontype.ListItems.Count
        If Item = weapontype.ListItems(t%) Then
            IIDX = t%
            t% = weapontype.ListItems.Count
        End If
    Next t%
End If
If IIDX > 0 Then
    '===== Error 258
    Select Case Val(Mid$(weapontype.ListItems(IIDX), InStr(weapontype.ListItems(IIDX), "(") + 1, 2))
        Case 11 To 15
        inp = InputBox("Automatic or SemiAutomatic (A/S/N)?", "Genesis Information Log", automatic(IIDX))
        inp = UCase(inp)
        If inp = "A" Then
            automatic(IIDX) = "A"
        Else
        If inp = "S" Then
            automatic(IIDX) = "S"
        Else
            automatic(IIDX) = "N"
        End If
        End If
        Case Else
            automatic(IIDX) = "N"
    End Select
End If
End Sub




Friend Sub clearroutine(TP As Integer)
mugshot.Picture = LoadPicture()
Dim itmx As ListItem
For pp% = 1 To 5
    VUCRSEL(pp%) = ""
Next pp%
FOUNDSELECT = 0
ARRESTEDNEARYES = 0
ARERSTEDNEARNO = 1
DATEOFARREST = ""
TIMEOFARREST = ""
vucrlist.ListItems.clear

    'RLB Code
    IncidentRptLoadedFromDB = False
    RelatedOfficers(0) = ""
    RelatedOfficers(1) = ""
    RelatedOfficers(2) = ""
    RelatedOfficers(3) = ""
    '*******
    
For t% = 0 To 5
    minorlist(t%).clear
Next t%
For t% = 0 To 1
    alcoholyes(t%) = 0
    alcoholno(t%) = 0
    alcoholunknown(t%) = 1
    drugsyes(t%) = 0
    drugsno(t%) = 0
    drugsunknown(t%) = 1
Next t%
For t% = 0 To 4
    activity(t%).Visible = False
    gactivity(t%).Visible = False
    UCRLIST(t%).Visible = False
    sublist(t%).Visible = False
    HOMOCIDE(t%).Visible = False
    additional(t%).Visible = False
    premise(Index).Height = 300
    premise(Index).Width = 1875
Next t%
For t% = 0 To 41
    totalvalue(t%) = ""
Next t%
For t% = 0 To 19
    relationship(t%).ListIndex = -1
    For tt% = 0 To relationship(t%).ListCount - 1
        relationship(t%).Selected(tt%) = False
    Next tt%
Next t%
TVSTOLEN = ""
TVDAMAGED = ""
TVBURNED = ""
TVRECOVERED = ""
TVSEIZED = ""
TVCOUNTERFEIT = ""
TVUNKNOWN = ""
onpaper = 0
lookupframe.Visible = False
For t% = 0 To 5
    DATERECOVERED(t%).Visible = False
    group(t%).ListIndex = -1
    numvehicle(t%).Visible = False
    numvehicle(t% + 6).Visible = False
    pucrlist(t%).clear
    pucrlist(t%).ListIndex = -1
    pinfoframe(t%).Visible = False
Next t%
EXCEPTIONALCLEARANCEDATE.Visible = False
relationshipframe(0).Visible = False
relationshipframe(1).Visible = False
sdrugframe(0).Visible = False
sdrugframe(1).Visible = False
sdrugframe(2).Visible = False
sdrugframe(3).Visible = False
sdrugframe(4).Visible = False
sdrugframe(5).Visible = False
sdrugframe(6).Visible = False
sdrugframe(7).Visible = False
BIAS.Visible = False
offenseframe.Visible = False
vucrf.Visible = False
Dim db As Database, rs As Recordset
For t% = 1 To 100
    automatic(t%) = "N"
Next t%
For v% = 0 To 4
    pickoffense(v%).ListIndex = -1
    completedy(v%) = False
    FORCEDENTRYY(v%) = 0
    completedn(v%) = False
    FORCEDENTRYN(v%) = 0
    For ii% = 1 To activity(v%).ListItems.Count
        activity(v%).ListItems(ii%).Selected = False
    Next ii%
    For ii% = 1 To gactivity(v%).ListItems.Count
        gactivity(v%).ListItems(ii%).Selected = False
    Next ii%
    UCRLIST(v%).ListIndex = -1
    For vv% = 0 To UCRLIST(v%).ListCount - 1
        UCRLIST(v%).Selected(vv%) = False
    Next vv%
    For vv% = 1 To premise(v%).ListItems.Count
        premise(v%).ListItems(vv%).Selected = False
    Next vv%
    entered(v%) = ""
    For uu% = 1 To sublist(v%).ListItems.Count
        sublist(v%).ListItems(uu%).Selected = False
    Next uu%
    For xx% = 1 To HOMOCIDE(v%).ListItems.Count
        HOMOCIDE(v%).ListItems(xx%).Selected = False
    Next xx%
    additional(v%).ListIndex = -1
Next v%
incidentlocation = ""
CLOCATIONNUMBER = ""
incidentzipcode = ""
subjectidentifiedyes = False
subjectidentifiedno = True
subjectlocatedyes = False
subjectlocatedno = True
active = 1
admclosed = 0
unfounded = 0
arrestedunder18 = 0
arrested18andover = 0
exclearunder18 = 0
exclear18andover = 0
offenderdeath = False
noprosecution = False
extraditiondenied = False
victimdeclinescooperation = False
juvenilenocustody = False
EXCEPTIONALCLEARANCEDATE = ""
reportingofficer(0) = ""
REPORTINGOFFICERDATE(0) = ""
reportingofficeRunit(0) = ""
reportingofficer(1) = ""
REPORTINGOFFICERDATE(1) = ""
reportingofficeRunit(1) = ""
followupyes = False
followupno = True
followupofficer = ""
approvingofficer = ""
APPROVINGOFFICERDATE = ""
approvingofficeRunit = ""
BIAS.ListIndex = -1
For t% = 0 To 5
    group(t%).ListIndex = -1
    totalvalue(t%) = ""
    description(t%) = ""
    majorlist(t%).ListIndex = -1
    minorlist(t%).clear
    numvehicle(t%) = ""
    numvehicle(t% + 6) = ""
    DATERECOVERED(t%) = ""
Next t%
stolen = 0
damaged = 0
burned = 0
recovered = 0
seized = 0
counterfeited = 0
typeunknown = 0
NARRATIVE.Text = ""
JURISDICTIONTHEFT = ""
JURISDICTIONRECOVERY = ""
vsname(0) = ""
vsname(1) = ""
vsname(2) = "UNKNOWN"
For t% = 0 To 2
    address(t%) = ""
    city(t%) = ""
    state(t%) = ""
    zipcode(t%) = ""
    race(t%).ListIndex = -1
    sex(t%).ListIndex = -1
    age(t%) = ""
    ethnicity(t%).ListIndex = -1
Next t%
BIRTHDATE = ""
For t% = 0 To 1
    ht(t%) = ""
    weight(t%) = ""
    resident(t%).ListIndex = -1
    hair(t%) = ""
    eyes(t%) = ""
    peculiarities(t%) = ""
Next t%
SUSPECT = 0
WARRANT = 0
WANTED = 0
RUNAWAY = 0
ARREST = 0
JAIL = 0
SUMMONS = 0
nearoffenseyes = False
nearoffenseno = False
TOTALARRESTED = ""
TIMEOFOFFENSE(0) = ""
TIMEOFOFFENSE(1) = ""
VISIBLEINJURYYES = False
VISIBLEINJURYNO = True
NONVISIBLEINJURYYES = False
NONVISIBLEINJURYNO = True
TWOMANVEHICLE = 0
ONEMANVEHICLE = 0
DETECTIVE = 0
TODOTHER = 0
ALONE = 0
ASSISTED = 0
For t% = 1 To injury.ListItems.Count
    injury.ListItems(t%).Selected = False
Next t%
For t% = 0 To 23
    drugtype(t%).ListIndex = -1
    drugmeasurement(t%).ListIndex = -1
    drugamt(t%) = ""
Next t%
dispatchdate = ""
DISPATCHTIME = ""
TIMEARRIVED = ""
DEPARTINGTIME = ""
For t% = 1 To weapontype.ListItems.Count
    weapontype.ListItems(t%).Selected = False
Next t%
individual = True
government = False
financialinstitution = False
policeofficer = False
societypublic = False
other = False
business = False
religiousorganziation = False
unknown = False
VGSTOLEN = 0
VGRECOVERED = 0
VGFOUND = 0
VGTOWED = 0
VGSUSPECT = 0
VGVICTIM = 0
VGVEHICLE = 0
VGGUN = 0
VGBOAT = 0
VgLICENSEPLATE = 0
vgsecurities = 0
VGARTICLE = 0
VIN = ""
HULL = ""
SERIAL = ""
VEHGUNSTATE = ""
YEARREG = ""
YEAREXP = ""
vYear = ""
make = ""
vgTYPE = ""
model = ""
STYLE = ""
vColor = ""
BRANDNAME = ""
CALIBER = ""
NIC = ""
DENOMINATION = ""
ISSUER = ""
SECURITIESDATE = ""
MISCELLANEOUS = ""

For t% = 0 To 1
    HOMEDAYPHONE(t%) = ""
    HOMENIGHTPHONE(t%) = ""
    WORKDAYPHONE(t%) = ""
    WORKNIGHTPHONE(t%) = ""
    LOCATIONNUMBER(t%) = ""
    incidentdate(t%) = ""
Next t%
ELOCATIONNUMBER = ""
dtoffense = ""

schanged = 0
VScroll1 = VScroll1.Min
If TP = 0 Then
    Call defaultcodes
Else
    Call DEFAULTCODESS
End If
On Error Resume Next
If TP < 3 Then
'---- setfocus logic ----
'             onpaper.SetFocus
          If onpaper.Visible Then
              onpaper.SetFocus
           End If
End If
On Error GoTo 0

End Sub


Private Sub defaultcodes()
Dim db As Database, rs As Recordset, itmx As ListItem
On Error GoTo oderror
od:
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("select * from codes")
If rs.EOF Then
    db.Close
    On Error Resume Next
    Exit Sub
End If
rs.MoveFirst
On Error Resume Next
For t% = 0 To 4
    premise(t%).ListItems.clear
Next t%
For t% = 0 To 2
    state(t%).clear
    city(t%).clear
    sex(t%).clear
    race(t%).clear
    ethnicity(t%).clear
Next t%
For t% = 0 To 1
    resident(t%).clear
Next t%
BIAS.clear
injury.ListItems.clear
weapontype.ListItems.clear
widx% = 0
While Not rs.EOF
    Select Case rs("type")
        Case "premise"
            For t% = 0 To 4
                Set itmx = premise(t%).ListItems.add(, , rs("code"))
                If UCase(rs("default")) = "Y" And t% = 0 Then
                    premise(t%).ListItems(premise(t%).ListItems.Count).Selected = True
                    premise(t%).ListItems(premise(t%).ListItems.Count).EnsureVisible
                End If
            Next t%
        Case "state"
            For t% = 0 To 2
                state(t%).AddItem rs("code")
'                If UCase(rs("default")) = "Y" Then
'                    state(t%).ListIndex = state(t%).ListCount - 1
'                End If
            Next t%
        Case "city"
            For t% = 0 To 2
                city(t%).AddItem rs("code")
'                If UCase(rs("default")) = "Y" Then
'                    city(t%).ListIndex = city(t%).ListCount - 1
'                End If
            Next t%
        Case "injury"
            Set itmx = injury.ListItems.add(, , rs("code"))
            If UCase(rs("default")) = "Y" Then
                injury.ListItems(injury.ListItems.Count).Selected = True
                injury.ListItems(injury.ListItems.Count).EnsureVisible
            Else
                injury.ListItems(injury.ListItems.Count).Selected = False
            End If
        Case "bias"
            BIAS.AddItem rs("code")
            If UCase(rs("default")) = "Y" Then
                BIAS.ListIndex = BIAS.ListCount - 1
            End If
        Case "sex"
            For t% = 0 To 2
                sex(t%).AddItem rs("code")
                If UCase(rs("default")) = "Y" Then
                    sex(t%).ListIndex = sex(t%).ListCount - 1
                End If
            Next t%
        Case "race"
            For t% = 0 To 2
                race(t%).AddItem rs("code")
                If UCase(rs("default")) = "Y" Then
                    race(t%).ListIndex = race(t%).ListCount - 1
                End If
            Next t%
        Case "ethnicity"
            For t% = 0 To 2
                ethnicity(t%).AddItem rs("code")
                If UCase(rs("default")) = "Y" Then
                    ethnicity(t%).ListIndex = ethnicity(t%).ListCount - 1
                End If
            Next t%
        Case "resident"
            For t% = 0 To 1
                resident(t%).AddItem rs("code")
                If UCase(rs("default")) = "Y" Then
                    resident(t%).ListIndex = resident(t%).ListCount - 1
                End If
            Next t%
        Case "weapon"
            Set itmx = weapontype.ListItems.add(, , rs("code"))
            widx% = widx% + 1
            automatic(widx%) = "N"
            If UCase(rs("default")) = "Y" Then
                weapontype.ListItems(weapontype.ListItems.Count).Selected = True
                weapontype.ListItems(weapontype.ListItems.Count).EnsureVisible
            Else
                weapontype.ListItems(weapontype.ListItems.Count).Selected = False
            End If
    End Select
    rs.MoveNext
Wend
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub
Friend Sub editproperty(editerr, msg As String)
Dim itmx As ListItem
'RLB Bandaid
On Error GoTo rlbErrp
'===== Data Element 20, 21, 22
Dim drugs(23)
drugs(6) = ""
drugs(7) = ""
drugs(8) = ""
drugs(9) = ""
drugs(10) = ""
drugs(11) = ""
drugs(12) = ""
drugs(13) = ""
drugs(14) = ""
drugs(15) = ""
drugs(16) = ""
drugs(17) = ""
drugs(18) = ""
drugs(19) = ""
drugs(20) = ""
drugs(21) = ""
drugs(22) = ""
drugs(23) = ""
For oo% = 6 To 23 Step 3
    dct% = oo%
    GONTONE = False
    For ii% = oo% To oo% + 2
        If drugtype(ii%).ListIndex > -1 Then
            drugs(dct%) = drugtype(ii%).List(drugtype(ii%).ListIndex)
            dct% = dct% + 1
            GOTONE = True
        End If
    Next ii%
    If dct% > 0 And GOTONE Then
        For Z% = oo% To dct%
            '===== Error 306
            '===== SCEdit 4/21/92 P29
            If Z% > 0 Then
                For ZZ% = 6 To Z% - 1
                    If drugs(ZZ%) > "" And drugs(ZZ%) = drugs(Z%) Then
                        If InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "GM=") > 0 Or _
                            InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "KG=") > 0 Or _
                            InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "OZ=") > 0 Or _
                            InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "LB=") > 0 Then
                                If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "GM=") > 0 Or _
                                    InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "KG=") > 0 Or _
                                    InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "OZ=") > 0 Or _
                                    InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "LB=") > 0 Then
                                        msg = "Two different measurements within the same weight category cannot be entered for the same type of drug."
                                        Call ShowApplicableContainers(drugmeasurement(ZZ%))
'---- setfocus logic ----
'                                                 drugmeasurement(zz%).SetFocus
          If drugmeasurement(ZZ%).Visible Then
              drugmeasurement(ZZ%).SetFocus
           End If
                                        GoTo exiteditp
                                End If
                        End If
                        If drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex) = drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex) Then
                            msg = "Two different measurements within the same weight category cannot be entered for the same type of drug."
                            Call ShowApplicableContainers(drugmeasurement(ZZ%))
'---- setfocus logic ----
'                                     drugmeasurement(zz%).SetFocus
          If drugmeasurement(ZZ%).Visible Then
              drugmeasurement(ZZ%).SetFocus
           End If
                            GoTo exiteditp
                        End If
                        If InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "ML=") > 0 Or _
                            InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "LT=") > 0 Or _
                            InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "FO=") > 0 Or _
                            InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "GL=") > 0 Then
                                If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "ML=") > 0 Or _
                                    InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "LT=") > 0 Or _
                                    InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "FO=") > 0 Or _
                                    InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "GL=") > 0 Then
                                        msg = "Two different measurements within the same weight category cannot be entered for the same type of drug."
                                        Call ShowApplicableContainers(drugmeasurement(ZZ%))
'---- setfocus logic ----
'                                                 drugmeasurement(zz%).SetFocus
          If drugmeasurement(ZZ%).Visible Then
              drugmeasurement(ZZ%).SetFocus
           End If
                                        GoTo exiteditp
                                End If
                        End If
                        If InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "DU=") > 0 Or _
                            InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "NP=") > 0 Then
                                If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "DU=") > 0 Or _
                                    InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "NP=") > 0 Then
                                        msg = "Two different measurements within the same weight category cannot be entered for the same type of drug."
                                        Call ShowApplicableContainers(drugmeasurement(ZZ%))
'---- setfocus logic ----
'                                                 drugmeasurement(zz%).SetFocus
          If drugmeasurement(ZZ%).Visible Then
              drugmeasurement(ZZ%).SetFocus
           End If
                                        GoTo exiteditp
                                End If
                        End If
                    End If
                Next ZZ%
            End If
            For t% = 1 To Len(drugamt(Z%))
                If InStr("0123456789.", Mid$(drugamt(Z%), t%, 1)) = 0 Then
                    msg = "Only 0123456789. are allowed in drug amount.  Enter all fractions as decimals (i.e. 1/2 = .5)."
                    Call ShowApplicableContainers(drugamt(Z%))
'---- setfocus logic ----
'                             drugamt(Z%).SetFocus
          If drugamt(Z%).Visible Then
              drugamt(Z%).SetFocus
           End If
                    GoTo exiteditp
                End If
            Next t%
            '===== Error 364
            If drugamt(Z%) = "" Or drugmeasurement(Z%).ListIndex = -1 Then
                If drugs(Z%) > "" And Left$(drugs(Z%), 1) <> "X" And Left$(drugs(Z%), 1) <> "U" Then
                    msg = "Drug Quantity and Measurement Type must be entered/selected."
                    Call ShowApplicableContainers(drugamt(Z%))
'---- setfocus logic ----
'                             drugamt(Z%).SetFocus
          If drugamt(Z%).Visible Then
              drugamt(Z%).SetFocus
           End If
                    GoTo exiteditp
                End If
            End If
            '===== Error 366
            If drugamt(Z%) > "" Then
                If drugs(Z%) = "" Or drugmeasurement(Z%).ListIndex = -1 Then
                    msg = "If a drug quantity is entered, then drug type and measurement type must also be entered."
                    Call ShowApplicableContainers(drugamt(Z%))
'---- setfocus logic ----
'                             drugamt(Z%).SetFocus
          If drugamt(Z%).Visible Then
              drugamt(Z%).SetFocus
           End If
                    GoTo exiteditp
                End If
            End If
            '===== Error 367
            If drugmeasurement(Z%).ListIndex > -1 Then
                If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "NP=") > 0 Then
                    If Left$(drugs(Z%), 1) <> "E" And Left$(drugs(Z%), 1) <> "G" And Left$(drugs(Z%), 1) <> "K" Then
                        msg = "Number of Plants is only allowabled with Marijuana, Opium, and Other Hallucinogens."
                        Call ShowApplicableContainers(drugmeasurement(Z%))
'---- setfocus logic ----
'                                 drugmeasurement(Z%).SetFocus
          If drugmeasurement(Z%).Visible Then
              drugmeasurement(Z%).SetFocus
           End If
                        GoTo exiteditp
                    End If
                End If
                '===== Error 384
                If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "XX=") > 0 Then
                    If Val(drugamt(Z%)) <> 1 Then
                        msg = "If drug measurement is NOT REPORTED, drug amount must be 1."
                        Call ShowApplicableContainers(drugmeasurement(Z%))
'---- setfocus logic ----
'                                 drugmeasurement(Z%).SetFocus
          If drugmeasurement(Z%).Visible Then
              drugmeasurement(Z%).SetFocus
           End If
                        GoTo exiteditp
                    End If
                End If
                '===== Error 368
                If drugs(Z%) = "" Or drugamt(Z%) = "" Then
                    msg = "If a drug measurement is entered, then drug type and quantity must also be entered."
                    Call ShowApplicableContainers(drugmeasurement(Z%))
'---- setfocus logic ----
'                             drugmeasurement(Z%).SetFocus
          If drugmeasurement(Z%).Visible Then
              drugmeasurement(Z%).SetFocus
           End If
                    GoTo exiteditp
                End If
            End If
            '===== Error 362
            If Left$(drugs(Z%), 1) = "X" Then
                If drugtype(oo%).ListIndex = -1 Or drugtype(oo% + 1).ListIndex = -1 Or drugtype(oo% + 2).ListIndex = -1 Then
                    msg = "If X = Over 3 Drug Types is entered, then the other 2 drug types must also be entered."
                    Call ShowApplicableContainers(drugtype(oo%))
'---- setfocus logic ----
'                             drugtype(oo%).SetFocus
          If drugtype(oo%).Visible Then
              drugtype(oo%).SetFocus
           End If
                    GoTo exiteditp
                End If
                '===== Error 363
                If drugamt(Z%) > "" Or drugmeasurement(Z%).ListIndex > -1 Then
                    msg = "Drug Amount and Measurement are not valid entries for X=Over 3 Drug Types."
                    Call ShowApplicableContainers(drugamt(Z%))
'---- setfocus logic ----
'                             drugamt(Z%).SetFocus
          If drugamt(Z%).Visible Then
              drugamt(Z%).SetFocus
           End If
                    GoTo exiteditp
                End If
            End If
        Next Z%
    End If
Next oo%
'GLENN
FOUNDPROP = False
For d% = 0 To 5
    If group(d%).ListIndex > -1 Or pucrlist(d%).ListIndex > -1 Then
        FOUNDPROP = True
    End If
Next d%
If FOUNDPROP Then
    For d% = 0 To 5
        'Data Element 16
        If group(d%).ListIndex > -1 Then
            If pucrlist(d%).ListIndex = -1 And pucrlist(d%).ListCount > 0 Then
                msg = "A UCR must be associated with the property described."
                Call ShowApplicableContainers(pucrlist(d%))
                If pucrlist(d%).Visible = True Then
'---- setfocus logic ----
'                             pucrlist(d%).SetFocus
          If pucrlist(d%).Visible Then
              pucrlist(d%).SetFocus
           End If
                End If
                GoTo exiteditp
            End If
            Select Case Mid$(group(d%).List(group(d%).ListIndex), InStr(group(d%).List(group(d%).ListIndex), "(") + 1, 2)
                Case "09", "22"
                    ctx% = 0
                    For dd% = 0 To 6
                        If totalvalue(d% + (dd% * 6)) = "X" Then
                            ctx% = ctx% + 1
                        End If
                    Next dd%
                    If ctx% = 0 Then
                        msg = "For Credit/Debit Cards and Nonnegotiable Instruments, no value is allowed, so an X must be entered to show the nature of the crime (stolen, recovered, etc.)"
                        Call ShowApplicableContainers(totalvalue(d%))
'---- setfocus logic ----
'                                 totalvalue(d%).SetFocus
          If totalvalue(d%).Visible Then
              totalvalue(d%).SetFocus
           End If
                        GoTo exiteditp
                    Else
                    If ctx% > 1 Then
                        msg = "For Credit/Debit Cards and Nonnegotiable Instruments, no value is allowed, so an X (only 1) must be entered to show the nature of the crime (stolen, recovered, etc.)"
                        Call ShowApplicableContainers(totalvalue(d%))
'---- setfocus logic ----
'                                 totalvalue(d%).SetFocus
          If totalvalue(d%).Visible Then
              totalvalue(d%).SetFocus
           End If
                        GoTo exiteditp
                    End If
                    End If
                Case "77", "99"
                    ctx% = 0
                    ctv% = 0
                    For dd% = 0 To 6
                        If totalvalue(d% + (dd% * 6)) = "X" Then
                            ctx% = ctx% + 1
                        End If
                        If Val(totalvalue(d% + (dd% * 6))) > 0 Then
                            ctv% = ctv% + 1
                        End If
                    Next dd%
                    If ctv% = 0 Then
                        If ctx% = 0 Then
                            msg = "For Other and Special Category, if no value is entered, an X must be entered to show the nature of the crime (stolen, recovered, etc.)"
                            Call ShowApplicableContainers(totalvalue(d%))
'---- setfocus logic ----
'                                     totalvalue(d%).SetFocus
          If totalvalue(d%).Visible Then
              totalvalue(d%).SetFocus
           End If
                            GoTo exiteditp
                        Else
                        If ctx% > 1 Then
                            msg = "For Other and Special Category, if no value is entered, an X (only 1) must be entered to show the nature of the crime (stolen, recovered, etc.)"
                            Call ShowApplicableContainers(totalvalue(d%))
'---- setfocus logic ----
'                                     totalvalue(d%).SetFocus
          If totalvalue(d%).Visible Then
              totalvalue(d%).SetFocus
           End If
                            GoTo exiteditp
                        End If
                        End If
                    End If
                Case "10"
                    If pucrlist(d%).ListIndex > -1 Then
                        ctx% = 0
                        ctv% = 0
                        For dd% = 0 To 6
                            If totalvalue(d% + (dd% * 6)) = "X" Then
                                ctx% = ctx% + 1
                            End If
                            If Val(totalvalue(d% + (dd% * 6))) > 0 Then
                                ctv% = ctv% + 1
                            End If
                        Next dd%
                        If InStr(pucrlist(d%).List(pucrlist(d%).ListIndex), "(35A)") = 0 Then
                            If ctx% > 0 Then
                                msg = "An X cannot be entered for group 10 (Drug/Narcotics) unless the UCR selected is 35A (Drug/Narcotic Violations)."
                                Call ShowApplicableContainers(totalvalue(d%))
'---- setfocus logic ----
'                                         totalvalue(d%).SetFocus
          If totalvalue(d%).Visible Then
              totalvalue(d%).SetFocus
           End If
                                GoTo exiteditp
                            End If
                            If ctv% = 0 Then
                                msg = "An amount must be entered for group 10 (Drug/Narcotics) when the UCR selected is not 35A (Drug/Narcotic Violations)."
                                Call ShowApplicableContainers(totalvalue(d%))
'---- setfocus logic ----
'                                         totalvalue(d%).SetFocus
          If totalvalue(d%).Visible Then
              totalvalue(d%).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        Else
                            Select Case d%
                                Case 0
                                    didx% = 6
                                Case 1
                                    didx% = 9
                                Case 2
                                    didx% = 12
                                Case 3
                                    didx% = 15
                                Case 4
                                    didx% = 18
                                Case 5
                                    didx% = 21
                            End Select
                            If drugtype(didx%).ListIndex = -1 Or Val(drugamt(didx%)) = 0 Or drugmeasurement(didx%).ListIndex = -1 Then
                                msg = "Drug information must be entered group 10 (Drug/Narcotics) when the UCR selected is 35A (Drug/Narcotic Violations)."
                                GoTo exiteditp
                            End If
                            If ctx% = 0 Then
                                msg = "An X must be entered for group 10 (Drug/Narcotics) when the UCR selected is 35A (Drug/Narcotic Violations)."
                                Call ShowApplicableContainers(totalvalue(d% + 24))
'---- setfocus logic ----
'                                         totalvalue(d%).SetFocus
          If totalvalue(d% + 24).Visible Then
              totalvalue(d% + 24).SetFocus
           End If
                                GoTo exiteditp
                            End If
                            If ctx% > 1 Then
                                msg = "An X (only 1) must be entered for group 10 (Drug/Narcotics) when the UCR selected is 35A (Drug/Narcotic Violations)."
                                Call ShowApplicableContainers(totalvalue(d% + 24))
'---- setfocus logic ----
'                                         totalvalue(d%).SetFocus
          If totalvalue(d% + 24).Visible Then
              totalvalue(d% + 24).SetFocus
           End If
                                GoTo exiteditp
                            End If
                            If ctv% > 0 Then
                                msg = "No amount must be entered for group 10 (Drug/Narcotics) when the UCR selected is 35A (Drug/Narcotic Violations). Use an X to note the type of property crime."
                                Call ShowApplicableContainers(totalvalue(d% + 24))
'---- setfocus logic ----
'                                         totalvalue(d%).SetFocus
          If totalvalue(d% + 24).Visible Then
              totalvalue(d% + 24).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                    End If
                Case Else
                    For dd% = 0 To 6
                        If totalvalue(d% + (dd% * 6)) = "X" Then
                            msg = "An entry of X is only allowed for groups 09, 22, 77, and 99."
                            Call ShowApplicableContainers(totalvalue(d% + (dd% * 6)))
'---- setfocus logic ----
'                                     totalvalue(d% + (dd% * 6)).SetFocus
          If totalvalue(d% + (dd% * 6)).Visible Then
              totalvalue(d% + (dd% * 6)).SetFocus
           End If
                            GoTo exiteditp
                        End If
                    Next dd%
            End Select
        End If
        If Not description(d%).Text = "" Then
            If pucrlist(d%).ListIndex = -1 And pucrlist(d%).ListCount > 0 Then
                msg = "A UCR must be associated with the property described."
                Call ShowApplicableContainers(pucrlist(d%))
                If pucrlist(d%).Visible = True Then
'---- setfocus logic ----
'                             pucrlist(d%).SetFocus
          If pucrlist(d%).Visible Then
              pucrlist(d%).SetFocus
           End If
                End If
                GoTo exiteditp
            Else
                If group(d%).ListIndex = -1 And pucrlist(d%).ListCount > 0 Then
                    msg = "A Group must be associated with the property described."
                    Call ShowApplicableContainers(group(d%))
'---- setfocus logic ----
'                             group(d%).SetFocus
          If group(d%).Visible Then
              group(d%).SetFocus
           End If
                    GoTo exiteditp
                End If
            End If
        'Else
        '    If majorlist(d%).ListIndex = -1 Or minorlist(d%).ListIndex = -1 Then
        '        msg = "A major and minor property type must be selected."
        '        majorlist(d%).ListIndex = 0
        '        majorlist(d%).ListIndex = -1
        '        Call ShowApplicableContainers(majorlist(d%))
    'RLB Code
            If majorlist(d%).ListIndex = -1 Or minorlist(d%).ListIndex = -1 Then
                If pucrlist(d%).ListCount > 0 Then
                    msg = "A major and minor property type must be selected."
                    Call ShowApplicableContainers(majorlist(d%))
'---- setfocus logic ----
'                             majorlist(d%).SetFocus
          If majorlist(d%).Visible Then
              majorlist(d%).SetFocus
           End If
                    GoTo exiteditp
                End If
            End If
    '********
        End If
    Next d%
End If
'GLENN
For t% = 0 To 41
    
    tt% = t% Mod 6
    '=== 081
    temppucr = Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3)
    If Val(totalvalue(t%)) > 0 Or totalvalue(t%) = "X" Then
        Select Case t%
            Case 0, 1, 2, 3, 4, 5
                Select Case temppucr
                    Case "510", "220", "270", "210", "26A", "26B", "26C", "26D", "26E", "23A", "23B", "23C", "23D", "23E", "23F", "23G", "23H", "100", "240", "120"
                    Case "90A", "90P", "90B", "90C", "90D", "90E", "90F", "90K", "90G", "90H", "90N", "90I", "90J", "90L", "90Z", ""
                    Case Else
                        msg = "Stolen value not allowed for ucr " + tempucr
                        Call ShowApplicableContainers(totalvalue(t%))
'---- setfocus logic ----
'                                 totalvalue(t%).SetFocus
          If totalvalue(t%).Visible Then
              totalvalue(t%).SetFocus
           End If
                        GoTo exiteditp
                End Select
            Case 6, 7, 8, 9, 10, 11
                Select Case temppucr
                    Case "290"
                    Case "90A", "90P", "90B", "90C", "90D", "90E", "90F", "90K", "90G", "90H", "90N", "90I", "90J", "90L", "90Z", ""
                    Case Else
                        msg = "Damaged value not allowed for ucr " + tempucr
                        Call ShowApplicableContainers(totalvalue(t%))
'---- setfocus logic ----
'                                 totalvalue(t%).SetFocus
          If totalvalue(t%).Visible Then
              totalvalue(t%).SetFocus
           End If
                        GoTo exiteditp
                End Select
            Case 12, 13, 14, 15, 16, 17
                Select Case temppucr
                    Case "200"
                    Case "90A", "90P", "90B", "90C", "90D", "90E", "90F", "90K", "90G", "90H", "90N", "90I", "90J", "90L", "90Z", ""
                    Case Else
                        msg = "Burned value not allowed for ucr " + tempucr
                        Call ShowApplicableContainers(totalvalue(t%))
'---- setfocus logic ----
'                                 totalvalue(t%).SetFocus
          If totalvalue(t%).Visible Then
              totalvalue(t%).SetFocus
           End If
                        GoTo exiteditp
                End Select
            Case 18, 19, 20, 21, 22, 23
                Select Case temppucr
                    Case "510", "220", "250", "270", "210", "26A", "26B", "26C", "26D", "26E", "23A", "23B", "23C", "23D", "23E", "23F", "23G", "23H", "100", "240", "120", "280"
                    Case "90A", "90P", "90B", "90C", "90D", "90E", "90F", "90K", "90G", "90H", "90N", "90I", "90J", "90L", "90Z", ""
                    Case Else
                        msg = "Recovered value not allowed for ucr " + tempucr
                        Call ShowApplicableContainers(totalvalue(t%))
'---- setfocus logic ----
'                                 totalvalue(t%).SetFocus
          If totalvalue(t%).Visible Then
              totalvalue(t%).SetFocus
           End If
                        GoTo exiteditp
                End Select
            Case 24, 25, 26, 27, 28, 29
                Select Case temppucr
                    Case "250", "35A", "35B", "39A", "39B", "39C", "39D"
                    Case "90A", "90P", "90B", "90C", "90D", "90E", "90F", "90K", "90G", "90H", "90N", "90I", "90J", "90L", "90Z", ""
                    Case Else
                        msg = "Seized value not allowed for ucr " + tempucr
                        Call ShowApplicableContainers(totalvalue(t%))
'---- setfocus logic ----
'                                 totalvalue(t%).SetFocus
          If totalvalue(t%).Visible Then
              totalvalue(t%).SetFocus
           End If
                        GoTo exiteditp
                End Select
            Case 30, 31, 32, 33, 34, 35
                Select Case temppucr
                    Case "250"
                    Case "90A", "90P", "90B", "90C", "90D", "90E", "90F", "90K", "90G", "90H", "90N", "90I", "90J", "90L", "90Z", ""
                    Case Else
                        msg = "Counterfeit value not allowed for ucr " + tempucr
                        Call ShowApplicableContainers(totalvalue(t%))
'---- setfocus logic ----
'                                 totalvalue(t%).SetFocus
          If totalvalue(t%).Visible Then
              totalvalue(t%).SetFocus
           End If
                        GoTo exiteditp
                End Select
            Case 36, 37, 38, 39, 40, 41
                Select Case temppucr
                    Case "90A", "90P", "90B", "90C", "90D", "90E", "90F", "90K", "90G", "90H", "90N", "90I", "90J", "90L", "90Z", ""
                    Case "200", "510", "220", "250", "35A", "35B", "290", "270", "210", "26A", "26B", "26C", "26D", "26E", "23A", "23B", "23C", "23D", "23E", "23F", "23G", "23H", "100", "240", "120", "39A", "39B", "39C", "39D", "280"
                    Case Else
                        msg = "Unknown value not allowed for ucr " + tempucr
                        Call ShowApplicableContainers(totalvalue(t%))
'---- setfocus logic ----
'                                 totalvalue(t%).SetFocus
          If totalvalue(t%).Visible Then
              totalvalue(t%).SetFocus
           End If
                        GoTo exiteditp
                End Select
        End Select
    End If
    
    If (Val(totalvalue(t%)) > 0 Or description(tt%) > "" Or group(tt%).ListIndex > -1) And pucrlist(tt%).ListCount > 0 Then
        '===== Error 342
        If Val(totalvalue(t%)) >= 1000000 Then
            msg = "WARNING:  A value of $1,000,000 or greater has been entered in the property value section.  Is this correct?"
            If msg = 7 Then
                Call ShowApplicableContainers(totalvalue(t%))
'---- setfocus logic ----
'                         totalvalue(t%).SetFocus
          If totalvalue(t%).Visible Then
              totalvalue(t%).SetFocus
           End If
                GoTo exiteditp
            End If
        End If
        Select Case t%
            Case 0, 6, 12, 18, 24
                If pucrlist(0).ListIndex = -1 Then
                    msg = "A UCR must be associated with the property described."
                    Call ShowApplicableContainers(pucrlist(0))
'---- setfocus logic ----
'                             pucrlist(0).SetFocus
          If pucrlist(0).Visible Then
              pucrlist(0).SetFocus
           End If
                    GoTo exiteditp
                End If
            Case 1, 7, 13, 19, 25
                If pucrlist(1).ListIndex = -1 Then
                    msg = "A UCR must be associated with the property described."
                    Call ShowApplicableContainers(pucrlist(1))
'---- setfocus logic ----
'                             pucrlist(1).SetFocus
          If pucrlist(1).Visible Then
              pucrlist(1).SetFocus
           End If
                    GoTo exiteditp
                End If
            Case 2, 8, 14, 20, 26
                If pucrlist(2).ListIndex = -1 Then
                    msg = "A UCR must be associated with the property described."
                    Call ShowApplicableContainers(pucrlist(2))
'---- setfocus logic ----
'                             pucrlist(2).SetFocus
          If pucrlist(2).Visible Then
              pucrlist(2).SetFocus
           End If
                    GoTo exiteditp
                End If
            Case 3, 9, 15, 21, 27
                If pucrlist(3).ListIndex = -1 Then
                    msg = "A UCR must be associated with the property described."
                    Call ShowApplicableContainers(pucrlist(3))
'---- setfocus logic ----
'                             pucrlist(3).SetFocus
          If pucrlist(3).Visible Then
              pucrlist(3).SetFocus
           End If
                    GoTo exiteditp
                End If
            Case 4, 10, 16, 22, 28
                If pucrlist(4).ListIndex = -1 Then
                    msg = "A UCR must be associated with the property described."
                    Call ShowApplicableContainers(pucrlist(4))
'---- setfocus logic ----
'                             pucrlist(4).SetFocus
          If pucrlist(4).Visible Then
              pucrlist(4).SetFocus
           End If
                    GoTo exiteditp
                End If
            Case 5, 11, 17, 23, 29
                If pucrlist(5).ListIndex = -1 Then
                    msg = "A UCR must be associated with the property described."
                    Call ShowApplicableContainers(pucrlist(5))
'---- setfocus logic ----
'                             pucrlist(5).SetFocus
          If pucrlist(5).Visible Then
              pucrlist(5).SetFocus
           End If
                    GoTo exiteditp
                End If
        End Select
        
            
    '===== Error 352
    If Val(totalvalue(tt%)) = 0 And Val(totalvalue(tt% + 6)) = 0 And Val(totalvalue(tt% + 12)) = 0 And Val(totalvalue(tt% + 18)) = 0 And Val(totalvalue(tt% + 24)) = 0 And Val(totalvalue(tt% + 30)) = 0 Then
        If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) <> "35A" Then
            If group(tt%).ListIndex > -1 Or DATERECOVERED(tt%) > "" Or numvehicle(tt%) > "" Or numvehicle(tt% + 6) > "" Or drugtype(tt% + 6).ListIndex > -1 Or drugamt(tt% + 6) > "" Or drugmeasurement(tt% + 6).ListIndex > -1 Then
                tempgrp = Mid$(group(tt%).List(group(tt%).ListIndex), InStr(group(tt%).List(group(tt%).ListIndex), "(") + 1, 2)
                If tempgrp <> "09" And tempgrp <> "22" And tempgrp <> "77" And tempgrp <> "99" Then
                    msg = "When Type Property Loss = None, no other applicable values are allowed."
                    Call ShowApplicableContainers(totalvalue(tt%))
'---- setfocus logic ----
'                             totalvalue(tt%).SetFocus
          If totalvalue(tt%).Visible Then
              totalvalue(tt%).SetFocus
           End If
                    GoTo exiteditp
                End If
            End If
        Else
            If DATERECOVERED(tt%) > "" Or numvehicle(tt%) > "" Or numvehicle(tt% + 6) > "" Then
                If Val(numvehicle(tt%)) = 0 And Val(numvehicle(tt% + 6)) = 0 Then
                    numvehicle(tt%) = ""
                    numvehicle(tt% + 6) = ""
                Else
                    msg = "When Type Property Loss = None for UCR 35A, no other applicable values are allowed, except drug-related values."
                    Call ShowApplicableContainers(group(tt%))
'---- setfocus logic ----
'                             group(tt%).SetFocus
          If group(tt%).Visible Then
              group(tt%).SetFocus
           End If
                    GoTo exiteditp
                End If
            End If
        End If
    End If
    
    If Val(totalvalue(tt% + 36)) > 0 Then
        If group(tt%).ListIndex > -1 Or DATERECOVERED(tt%) > "" Or numvehicle(tt%) > "" Or numvehicle(tt% + 6) > "" Or drugtype(tt% + 6).ListIndex > -1 Or drugamt(tt% + 6) > "" Or drugmeasurement(tt% + 6).ListIndex > -1 Then
            msg = "When Type Property Loss = Unknown, no other applicable values are allowed."
            Call ShowApplicableContainers(group(tt%))
'---- setfocus logic ----
'                     group(tt%).SetFocus
          If group(tt%).Visible Then
              group(tt%).SetFocus
           End If
            GoTo exiteditp
        End If
    End If
    
    '===== Edit 84
    If Val(totalvalue(tt%)) > 0 And Val(totalvalue(tt% + 18)) > 0 Then
        If Val(totalvalue(tt%)) < Val(totalvalue(tt% + 18)) Then
            msg = "Recovered value cannot be greater than stolen value."
            Call ShowApplicableContainers(totalvalue(tt%))
'---- setfocus logic ----
'                     totalvalue(tt%).SetFocus
          If totalvalue(tt%).Visible Then
              totalvalue(tt%).SetFocus
           End If
            GoTo exiteditp
        End If
    End If
    
    '===== Data Element 14, 15
    '===== Error 372,375
    If Val(totalvalue(tt%)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
        If description(tt%) = "" Or pucrlist(tt%).ListIndex = -1 Or group(tt%).ListIndex = -1 Then
            msg = "If Burned, Counterfeited, Damaged, Recovered, Seized, or Stolen are selected, then all other PROPERTY values must be entered."
            Call ShowApplicableContainers(totalvalue(tt%))
'---- setfocus logic ----
'                     totalvalue(tt%).SetFocus
          If totalvalue(tt%).Visible Then
              totalvalue(tt%).SetFocus
           End If
            GoTo exiteditp
        End If
        If Val(totalvalue(tt% + 18)) > 0 Then
            If Not IsDate(DATERECOVERED(tt%)) Then
                msg = "If Burned, Counterfeited, Damaged, Recovered, Seized, or Stolen are selected, the all other PROPERTY values must be entered."
                Call ShowApplicableContainers(totalvalue(tt% + 18))
'---- setfocus logic ----
'                         totalvalue(tt% + 18).SetFocus
          If totalvalue(tt% + 18).Visible Then
              totalvalue(tt% + 18).SetFocus
           End If
                GoTo exiteditp
            Else
                If CDate(DATERECOVERED(tt%)) < CDate(incidentdate(1)) Then
                    msg = "Recovery date cannot be before incident date."
                    GoTo exiteditp
                End If
            End If
        End If
    End If
    
    For uu% = 0 To 5
        '===== Error 268
        If group(uu%).ListIndex > -1 And nolarceny = 0 Then
            Select Case Val(Mid$(group(uu%).List(group(uu%).ListIndex), InStr(group(uu%).List(group(uu%).ListIndex), "(") + 1, 2))
                Case 1, 5, 24, 28, 37
                    msg = "Illogical property group/ucr combination."
                    Call ShowApplicableContainers(group(uu%))
'---- setfocus logic ----
'                             group(uu%).SetFocus
          If group(uu%).Visible Then
              group(uu%).SetFocus
           End If
                    GoTo exiteditp
            End Select
        End If
        If pucrlist(uu%).ListIndex > -1 Then
            '===== Error 390
            '===== SCEdit 4/21/92 P30
            Select Case Mid$(pucrlist(uu%).List(pucrlist(uu%).ListIndex), InStr(pucrlist(uu%).List(pucrlist(uu%).ListIndex), "(") + 1, 3)
                Case "240", "220"
                    Select Case Val(Mid$(group(uu%).List(group(uu%).ListIndex), InStr(group(uu%).List(group(uu%).ListIndex), "(") + 1, 2))
                        Case 29 To 35
                            msg = "Illogical property group/ucr combination."
                            Call ShowApplicableContainers(pucrlist(uu%))
'---- setfocus logic ----
'                                     pucrlist(uu%).SetFocus
          If pucrlist(uu%).Visible Then
              pucrlist(uu%).SetFocus
           End If
                            GoTo exiteditp
                    End Select
                    '===== Data Element 18
                    '===== Error 357,358,359
                    If Val(totalvalue(tt%)) > 0 And Val(numvehicle(tt%)) = 0 Then
                        If group(tt%).ListIndex > -1 Then
                            tempgrp = Mid$(group(tt%).List(group(tt%).ListIndex), InStr(group(tt%).List(group(tt%).ListIndex), "(") + 1, 2)
                            If tempgrp = "03" Or tempgrp = "05" Or tempgrp = "24" Or tempgrp = "28" Or tempgrp = "37" Then
                                msg = "A number of stolen vehicles must be entered."
                                Call ShowApplicableContainers(numvehicle(tt%))
'---- setfocus logic ----
'                                         numvehicle(tt%).SetFocus
          If numvehicle(tt%).Visible Then
              numvehicle(tt%).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                    End If
                    '===== Data Element 19
                    '===== Error 360,361,359
                    '===== SCEdit 12/19/92 P35-C  allow tempgrp = 38
                    If Val(totalvalue(tt% + 18)) > 0 And Val(numvehicle(tt% + 6)) = 0 Then
                        If group(tt%).ListIndex > -1 Then
                            tempgrp = Mid$(group(tt%).List(group(tt%).ListIndex), InStr(group(tt%).List(group(tt%).ListIndex), "(") + 1, 2)
                            If tempgrp = "03" Or tempgrp = "05" Or tempgrp = "24" Or tempgrp = "28" Or tempgrp = "37" Or tempgrp = "38" Then
                                msg = "A number of recovered vehicles must be entered."
                                Call ShowApplicableContainers(numvehicle(tt% + 6))
'---- setfocus logic ----
'                                         numvehicle(tt% + 6).SetFocus
          If numvehicle(tt% + 6).Visible Then
              numvehicle(tt% + 6).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                    End If
                Case "23B", "23A", "23C"
                    Select Case Val(Mid$(group(uu%).List(group(uu%).ListIndex), InStr(group(uu%).List(group(uu%).ListIndex), "(") + 1, 2))
                        Case 1, 3, 4, 5, 12, 15, 18, 24, 28, 29, 30, 31, 32, 33, 34, 35, 37, 39
                            msg = "Illogical property group/ucr combination."
                            Call ShowApplicableContainers(pucrlist(uu%))
'---- setfocus logic ----
'                                     pucrlist(uu%).SetFocus
          If pucrlist(uu%).Visible Then
              pucrlist(uu%).SetFocus
           End If
                            GoTo exiteditp
                    End Select
                Case "23C"
                    Select Case Val(Mid$(group(uu%).List(group(uu%).ListIndex), InStr(group(uu%).List(group(uu%).ListIndex), "(") + 1, 2))
                        Case 1, 3, 5, 12, 15, 18, 24, 28, 29, 30, 31, 32, 33, 34, 35, 37, 39
                            msg = "Illogical property group/ucr combination."
                            Call ShowApplicableContainers(pucrlist(uu%))
'---- setfocus logic ----
'                                     pucrlist(uu%).SetFocus
          If pucrlist(uu%).Visible Then
              pucrlist(uu%).SetFocus
           End If
                            GoTo exiteditp
                    End Select
                Case "23F", "23D", "23E"
                    Select Case Val(Mid$(group(uu%).List(group(uu%).ListIndex), InStr(group(uu%).List(group(uu%).ListIndex), "(") + 1, 2))
                        Case 3, 5, 24, 28, 29, 30, 31, 32, 33, 34, 35, 37
                            msg = "Illogical property group/ucr combination."
                            Call ShowApplicableContainers(pucrlist(uu%))
'---- setfocus logic ----
'                                     pucrlist(uu%).SetFocus
          If pucrlist(uu%).Visible Then
              pucrlist(uu%).SetFocus
           End If
                            GoTo exiteditp
                    End Select
                Case "23G"
                    Select Case Val(Mid$(group(uu%).List(group(uu%).ListIndex), InStr(group(uu%).List(group(uu%).ListIndex), "(") + 1, 2))
                        Case 38, 88
                        Case Else
                            msg = "Illogical property group/ucr combination."
                            Call ShowApplicableContainers(pucrlist(uu%))
'---- setfocus logic ----
'                                     pucrlist(uu%).SetFocus
          If pucrlist(uu%).Visible Then
              pucrlist(uu%).SetFocus
           End If
                            GoTo exiteditp
                    End Select
                Case "23H"
                    Select Case Val(Mid$(group(uu%).List(group(uu%).ListIndex), InStr(group(uu%).List(group(uu%).ListIndex), "(") + 1, 2))
                        Case 3, 5, 24, 28, 37
                            msg = "Illogical property group/ucr combination."
                            Call ShowApplicableContainers(pucrlist(uu%))
'---- setfocus logic ----
'                                     pucrlist(uu%).SetFocus
          If pucrlist(uu%).Visible Then
              pucrlist(uu%).SetFocus
           End If
                            GoTo exiteditp
                    End Select
            End Select
            If Mid$(pucrlist(uu%).List(pucrlist(uu%).ListIndex), InStr(pucrlist(uu%).List(pucrlist(uu%).ListIndex), "(") + 1, 3) = "35A" Then
                If (Val(totalvalue(uu% + 0)) = 0 And Val(totalvalue(uu% + 6)) = 0 And Val(totalvalue(uu% + 12)) = 0 And Val(totalvalue(uu% + 18)) = 0 And Val(totalvalue(uu% + 24)) = 0 And Val(totalvalue(uu% + 30)) = 0) Then
                    If drugtype((uu% * 3) + 6).ListIndex = -1 Then
                        msg = "A suspected drug type must be selected for this property."
                        Call ShowApplicableContainers(pucrlist(uu%))
'---- setfocus logic ----
'                                 pucrlist(uu%).SetFocus
          If pucrlist(uu%).Visible Then
              pucrlist(uu%).SetFocus
           End If
                        GoTo exiteditp
                    End If
                End If
                If Val(totalvalue(uu% + 24)) > 0 Then
                    If group(uu%).ListIndex > -1 Then
                        If Mid$(group(uu%).List(group(uu%).ListIndex), InStr(group(uu%).List(group(uu%).ListIndex), "(") + 1, 2) = "10" Then
                            If drugtype((uu% * 3) + 6).ListIndex = -1 Then
                                msg = "A suspected drug type must be selected for this property."
                                Call ShowApplicableContainers(totalvalue(uu% + 24))
'---- setfocus logic ----
'                                         totalvalue(uu% + 24).SetFocus
          If totalvalue(uu% + 24).Visible Then
              totalvalue(uu% + 24).SetFocus
           End If
                                GoTo exiteditp
                            End If
                            If Val(drugamt((uu% * 3) + 6)) = 0 Or drugmeasurement((uu% * 3) + 6).ListIndex = -1 Then
                                msg = "An amount of suspected drugs must be entered for this property."
                                Call ShowApplicableContainers(drugamt((uu% * 3) + 6))
'---- setfocus logic ----
'                                         drugamt((uu% * 3) + 6).SetFocus
          If drugamt((uu% * 3) + 6).Visible Then
              drugamt((uu% * 3) + 6).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next uu%
        
    '===== Data Element 15
    '===== Error 353
    If InStr(group(tt%).List(group(tt%).ListIndex), "(88)") > 0 Then
        If Val(totalvalue(tt%)) <> 1 Then
            msg = "If Type of Property = Pending Inventory(88), a value of 1 must be entered."
            Call ShowApplicableContainers(group(tt%))
'---- setfocus logic ----
'                     group(tt%).SetFocus
          If group(tt%).Visible Then
              group(tt%).SetFocus
           End If
            GoTo exiteditp
        End If
    End If
    
    '===== Data Element 16
    '===== Error 351
    If group(tt%).ListIndex > -1 Then
        tempgrp = Mid$(group(tt%).List(group(tt%).ListIndex), InStr(group(tt%).List(group(tt%).ListIndex), "(") + 1, 2)
        If (Val(totalvalue(tt%)) = 0 And Val(totalvalue(tt% + 6)) = 0 And Val(totalvalue(tt% + 12)) = 0 And Val(totalvalue(tt% + 18)) = 0 And Val(totalvalue(tt% + 24)) = 0 And Val(totalvalue(tt% + 30)) = 0) And _
            tempgrp <> "09" And tempgrp <> "22" And tempgrp <> "77" And tempgrp <> "99" And tempgrp <> "10" Then
            msg = "A property value of 0 is only allowed for Credit/Debit Cards, Nonnegotiable Instruments, Other, Drug/Narcotics and Special Category."
            Call ShowApplicableContainers(group(tt%))
'---- setfocus logic ----
'                     group(tt%).SetFocus
          If group(tt%).Visible Then
              group(tt%).SetFocus
           End If
            GoTo exiteditp
        End If
        '---Ed Sloan added "Or Val(itmx.SubItems(2)) <> 1"-----------
        If Not (Val(totalvalue(tt%)) = 0 And Val(totalvalue(tt% + 6)) = 0 And Val(totalvalue(tt% + 12)) = 0 And Val(totalvalue(tt% + 18)) = 0 And Val(totalvalue(tt% + 24)) = 0 And Val(totalvalue(tt% + 30)) = 0) And _
            (tempgrp = "09" Or tempgrp = "22") Then
            msg = "A property value of 0 is required for Credit/Debit Cards and Nonnegotiable Instruments."
            Call ShowApplicableContainers(totalvalue(tt%))
'---- setfocus logic ----
'                     totalvalue(tt%).SetFocus
          If totalvalue(tt%).Visible Then
              totalvalue(tt%).SetFocus
           End If
            GoTo exiteditp
        End If
        '===== Error 383
        If tempgrp = "10" And InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(35A)") > 0 And _
            (Val(totalvalue(tt%)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0) Then
            msg = "A value is not valid for the Drugs/Narcotics and Drug/Narcotic Violations combination."
            Call ShowApplicableContainers(totalvalue(tt%))
'---- setfocus logic ----
'                     totalvalue(tt%).SetFocus
          If totalvalue(tt%).Visible Then
              totalvalue(tt%).SetFocus
           End If
            GoTo exiteditp
        End If
    Else
    '===== Error 354
        If (Val(totalvalue(tt%)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0) Then
            msg = "If a value greater than 0 is entered, an associated property type must be selected."
            Call ShowApplicableContainers(totalvalue(tt%))
'---- setfocus logic ----
'                     totalvalue(tt%).SetFocus
          If totalvalue(tt%).Visible Then
              totalvalue(tt%).SetFocus
           End If
            GoTo exiteditp
        End If
    End If
    
    '===== Data Element 17
    '===== Error 305
    If DATERECOVERED(tt%) > "" And Not IsDate(DATERECOVERED(tt%)) Then
        msg = "Date Recovered is not a valid date."
        Call ShowApplicableContainers(DATERECOVERED(tt%))
'---- setfocus logic ----
'                 DATERECOVERED(tt%).SetFocus
          If DATERECOVERED(tt%).Visible Then
              DATERECOVERED(tt%).SetFocus
           End If
        GoTo exiteditp
    End If
    If IsDate(DATERECOVERED(tt%)) Then
        If CVDate(DATERECOVERED(tt%)) < CVDate(incidentdate(0)) Then
            msg = "Date Recovered cannot be earlier that Date of Offense."
            Call ShowApplicableContainers(DATERECOVERED(tt%))
            DATERECOVERED(tt%).ZOrder
'---- setfocus logic ----
'                     DATERECOVERED(tt%).SetFocus
          If DATERECOVERED(tt%).Visible Then
              DATERECOVERED(tt%).SetFocus
           End If
            GoTo exiteditp
        End If
        '===== Error 356
        If Val(totalvalue(tt% + 18)) = 0 Or group(tt%).ListIndex = -1 Then
            msg = "If Date Recovered is entered, both type and value of property must be entered."
            Call ShowApplicableContainers(totalvalue(tt%))
'---- setfocus logic ----
'                     totalvalue(tt%).SetFocus
          If totalvalue(tt%).Visible Then
              totalvalue(tt%).SetFocus
           End If
            GoTo exiteditp
        End If
        If Val(totalvalue(tt% + 18)) = 0 Then
            msg = "If Date Recovered is entered, then Recovered must be selected."
            Call ShowApplicableContainers(DATERECOVERED(tt%))
'---- setfocus logic ----
'                     totalvalue(tt%).SetFocus
          If totalvalue(tt%).Visible Then
              totalvalue(tt%).SetFocus
           End If
            GoTo exiteditp
        End If
    End If
    
    
    '===== Data Element 20
    '===== Error 362
    If Left$(drugtype(((tt% + 2) * 3)).List(drugtype(((tt% + 2) * 3)).ListIndex), 1) = "X" Or Left$(drugtype(((tt% + 2) * 3) + 1).List(drugtype(((tt% + 2) * 3) + 1).ListIndex), 1) = "X" Or Left$(drugtype(((tt% + 2) * 3) + 2).List(drugtype(((tt% + 2) * 3) + 2).ListIndex), 1) = "X" Then
        If drugtype((tt% + 2) * 3).ListIndex = -1 Or drugtype(((tt% + 2) * 3) + 1).ListIndex = -1 Or drugtype(((tt% + 2) * 3) + 2).ListIndex = -1 Then
            msg = "If X = Over 3 Drug Types is entered, then the other 2 drug types must also be entered."
            Call ShowApplicableContainers(drugtype((tt% + 2) * 3))
'---- setfocus logic ----
'                     drugtype((tt% + 2) * 3).SetFocus
          If drugtype((tt% + 2) * 3).Visible Then
              drugtype((tt% + 2) * 3).SetFocus
           End If
            GoTo exiteditp
        End If
    End If
        
    End If
    
    
Next t%

For t% = 0 To 4
    UCRLIST(t%).ListIndex = -1
    For tt% = 0 To UCRLIST(t%).ListCount - 1
        If UCRLIST(t%).Selected(tt%) = True Then
            UCRLIST(t%).ListIndex = tt%
            tt% = UCRLIST(t%).ListCount - 1
        End If
    Next tt%
    
    If UCRLIST(t%).ListIndex > -1 Then
        tempucr = Mid$(UCRLIST(t%).List(UCRLIST(t%).ListIndex), InStr(UCRLIST(t%).List(UCRLIST(t%).ListIndex), "(") + 1, 3)
    
        '===Additional F 7
        '===== Data Element 7
        '===== Data Element 12
        If tempucr = "35A" Or tempucr = "35B" Then
            TP$ = "Drug Offenses"
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Not completedy(t%) Then
                            If Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                msg = TP$ + " not completed must have associated value of None or Unknown on property tab."
                                Call ShowApplicableContainers(totalvalue(tt% + 0))
'---- setfocus logic ----
'                                         totalvalue(tt% + 0).SetFocus
          If totalvalue(tt% + 0).Visible Then
              totalvalue(tt% + 0).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                        If completedy(t%) Then
                            'GLENN
                            '===== Error 301
                            If (Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) = 0 Or Val(totalvalue(tt% + 30)) > 0) Then
                                If Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                    msg = TP$ + " completed must have associated information of Seized property entered."
                                    Call ShowApplicableContainers(completedy(t%))
'---- setfocus logic ----
'                                             completedy(t%).SetFocus
          If completedy(t%).Visible Then
              completedy(t%).SetFocus
           End If
                                    GoTo exiteditp
                                End If
                            End If
                        End If
                    End If
                End If
            Next tt%
        End If
        If tempucr = "35A" Then
            TP$ = "Drug Offenses"
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Val(totalvalue(tt% + 24)) > 0 Or totalvalue(tt% + 24) = "X" Then
                            If group(tt%).ListIndex > -1 Then
                                '===== SCEdit 4/21/92 P29
                                If Mid$(group(tt%).List(group(tt%).ListIndex), InStr(group(tt%).List(group(tt%).ListIndex), "(") + 1, 2) <> "10" Then
                                    msg = "Property Description must have a value of 10 for Seized on Drug offense."
                                    Call ShowApplicableContainers(group(tt%))
'---- setfocus logic ----
'                                             group(tt%).SetFocus
          If group(tt%).Visible Then
              group(tt%).SetFocus
           End If
                                    GoTo exiteditp
                                End If
                            End If
                        End If
                    End If
                End If
            Next tt%
        End If
        If tempucr = "35B" Then
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Val(totalvalue(tt% + 24)) > 0 Then
                            If group(tt%).ListIndex > -1 Then
                                '===== SCEdit 4/21/92 P29
                                If Mid$(group(tt%).List(group(tt%).ListIndex), InStr(group(tt%).List(group(tt%).ListIndex), "(") + 1, 2) <> "11" Then
                                    msg = "Property Description must have a value of 11 for Seized on Drug Equipment offense."
                                    Call ShowApplicableContainers(group(tt%))
'---- setfocus logic ----
'                                             group(tt%).SetFocus
          If group(tt%).Visible Then
              group(tt%).SetFocus
           End If
                                    GoTo exiteditp
                                End If
                            End If
                        End If
                    End If
                End If
            Next tt%
        End If
        ' === Additional F 1
        If tempucr = "200" Then
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Val(totalvalue(tt%)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                            If Not completedy(t%) Then
                                If Val(totalvalue(tt%)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                    msg = "Arson not completed must have associated value of None or Unknown on property tab."
                                    Call ShowApplicableContainers(totalvalue(tt%))
'---- setfocus logic ----
'                                             totalvalue(tt%).SetFocus
          If totalvalue(tt%).Visible Then
              totalvalue(tt%).SetFocus
           End If
                                    GoTo exiteditp
                                End If
                            End If
                            If completedy(t%) Then
                                If Val(totalvalue(tt% + 12)) = 0 Then
                                    msg = "Arson completed must have associated value of Burned on property tab."
                                    Call ShowApplicableContainers(totalvalue(tt% + 12))
'---- setfocus logic ----
'                                             totalvalue(tt% + 12).SetFocus
          If totalvalue(tt% + 12).Visible Then
              totalvalue(tt% + 12).SetFocus
           End If
                                    GoTo exiteditp
                                End If
                                If group(tt%).ListIndex = -1 Then
                                    msg = "Type(Group) and Total Value must be entered."
                                    Call ShowApplicableContainers(group(tt%))
'---- setfocus logic ----
'                                             group(tt%).SetFocus
          If group(tt%).Visible Then
              group(tt%).SetFocus
           End If
                                    GoTo exiteditp
                                End If
                            End If
                        End If
                    End If
                End If
            Next tt%
        End If
        
        ' === Additional F 3
        '===== Additional F 4
        '===== Data Element 10
        '===== Data Element 11
        If tempucr = "510" Or tempucr = "220" Then
            Select Case tempucr
                Case "510"
                    TP$ = "Bribery"
                Case "220"
                    TP$ = "Burglary"
            End Select
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Not completedy(t%) Then
                            If Val(totalvalue(tt%)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                msg = TP$ + " not completed must have associated value of None or Unknown on property tab."
                                Call ShowApplicableContainers(totalvalue(tt%))
'---- setfocus logic ----
'                                         totalvalue(tt%).SetFocus
          If totalvalue(tt%).Visible Then
              totalvalue(tt%).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                        If completedy(t%) Then
                            If (Val(totalvalue(tt%)) = 0 And Val(totalvalue(tt% + 6)) = 0 And Val(totalvalue(tt% + 12)) = 0 And Val(totalvalue(tt% + 18)) = 0 And Val(totalvalue(tt% + 24)) = 0 And Val(totalvalue(tt% + 30)) = 0) Or _
                                ((Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt%)) > 0) And Val(totalvalue(tt% + 6)) = 0 And Val(totalvalue(tt% + 12)) = 0 And Val(totalvalue(tt% + 24)) = 0 And Val(totalvalue(tt% + 30)) = 0) Then
                                temperr = 0
                            Else
                                msg = TP$ + " completed must have associated value of None, Recovered, Stolen, or Unknown on property tab."
                                Call ShowApplicableContainers(totalvalue(tt%))
'---- setfocus logic ----
'                                         totalvalue(tt%).SetFocus
          If totalvalue(tt%).Visible Then
              totalvalue(tt%).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                        If Val(totalvalue(tt%)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Then
                            If group(tt%).ListIndex = -1 Then
                                msg = "Type(Group) and Total Value must be entered."
                                Call ShowApplicableContainers(totalvalue(tt%))
'---- setfocus logic ----
'                                         totalvalue(tt%).SetFocus
          If totalvalue(tt%).Visible Then
              totalvalue(tt%).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                        If Val(totalvalue(tt% + 18)) > 0 Then
                            If Not IsDate(DATERECOVERED(tt%)) Then
                                msg = "A valid DATE RECOVERED must be entered."
                                Call ShowApplicableContainers(totalvalue(tt%))
'---- setfocus logic ----
'                                         totalvalue(tt%).SetFocus
          If totalvalue(tt%).Visible Then
              totalvalue(tt%).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                    End If
                End If
            Next tt%
        End If
        
        '===Additional F 5
        '===== Data Element 12
        If tempucr = "250" Then
            TP$ = "Counterfeiting"
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Not completedy(t%) Then
                            If Val(totalvalue(tt%)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                msg = TP$ + " not completed must have associated value of None or Unknown on property tab."
                                Call ShowApplicableContainers(totalvalue(tt%))
'---- setfocus logic ----
'                                         totalvalue(tt%).SetFocus
          If totalvalue(tt%).Visible Then
              totalvalue(tt%).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                        If completedy(t%) Then
                            If Val(totalvalue(tt% + 18)) = 0 And Val(totalvalue(tt% + 24)) = 0 And Val(totalvalue(tt% + 30)) = 0 Then
                                msg = TP$ + " completed must have associated value of Counterfeited, Recovered, or Seized on property tab."
                                Call ShowApplicableContainers(totalvalue(tt%))
'---- setfocus logic ----
'                                         totalvalue(tt%).SetFocus
          If totalvalue(tt%).Visible Then
              totalvalue(tt%).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                        If Val(totalvalue(tt% + 18)) > 0 Then
                            If Not IsDate(DATERECOVERED(tt%)) Then
                                msg = "A valid DATE RECOVERED must be entered."
                                Call ShowApplicableContainers(totalvalue(tt% + 18))
'---- setfocus logic ----
'                                         totalvalue(tt% + 18).SetFocus
          If totalvalue(tt% + 18).Visible Then
              totalvalue(tt% + 18).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                    End If
                End If
            Next tt%
        End If
        
        '===== Additional F 6
        '===== Data Element 7
        If tempucr = "290" Then
            TP$ = "Destruction"
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Not completedy(t%) Then
                            If Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                msg = TP$ + " not completed must have associated value of None or Unknown on property tab."
                                Call ShowApplicableContainers(completedy(t%))
'---- setfocus logic ----
'                                         completedy(t%).SetFocus
          If completedy(t%).Visible Then
              completedy(t%).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                        If completedy(t%) Then
                            If Val(totalvalue(tt% + 6)) = 0 Or Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                msg = TP$ + " completed must have associated value of Damaged on property tab."
                                Call ShowApplicableContainers(completedy(t%))
'---- setfocus logic ----
'                                         completedy(t%).SetFocus
          If completedy(t%).Visible Then
              completedy(t%).SetFocus
           End If
                                GoTo exiteditp
                            End If
                            If group(tt%).ListIndex = -1 Then
                                msg = "Type(Group) and Total Value must be entered."
                                Call ShowApplicableContainers(group(tt%))
'---- setfocus logic ----
'                                         group(tt%).SetFocus
          If group(tt%).Visible Then
              group(tt%).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                    End If
                Else
                'If TT% = 0 Then
                '    msg = "Valid property must be entered."
                '    GoTo exiteditp
                'End If
                End If
            Next tt%
        End If
        
        '===== Additional F 8
        '===== Additional F 9
        '===== Additional F 10
        '===== Additional F 14
        '===== Additional F 18
        '===== Error 077, 078
        TP$ = ""
        Select Case tempucr
            Case "270", "210", "26A", "26B", "26C", "26D", "26E", "23A", "23B", "23C", "23D", "23E", "23F", "23G", "23H", "240", "120"
                TP$ = "Crimes Against Property"
        End Select
        '===== Error 074
        If TP$ > "" Then
            foundmatch% = 0
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        foundmatch% = 1
                    End If
                End If
            Next tt%
            If foundmatch% = 0 And completedy(t%) Then
                msg = TP$ + " must have property data."
                Call ShowApplicableContainers(pucrlist(0))
'---- setfocus logic ----
'                         pucrlist(0).SetFocus
          If pucrlist(0).Visible Then
              pucrlist(0).SetFocus
           End If
                GoTo exiteditp
            End If
        End If
        
        If TP$ > "" Then
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Not completedy(t%) Then
                            If Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                msg = TP$ + " not completed must have associated value of None or Unknown on property tab."
                                Call ShowApplicableContainers(completedy(t%))
'---- setfocus logic ----
'                                         completedy(t%).SetFocus
          If completedy(t%).Visible Then
              completedy(t%).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                        If group(tt%).ListIndex > -1 Then
                            tempgrp = Mid$(group(tt%).List(group(tt%).ListIndex), InStr(group(tt%).List(group(tt%).ListIndex), "(") + 1, 2)
                        End If
                        If completedy(t%) Then
                            If Val(totalvalue(tt%)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Then
                                temperr = 0
                            Else
                                If tempgrp <> "09" And tempgrp <> "22" And tempgrp <> "77" And tempgrp <> "99" Then
                                    msg = TP$ + " completed must have associated value of Recovered or Stolen on property tab."
                                    Call ShowApplicableContainers(completedy(t%))
'---- setfocus logic ----
'                                             completedy(t%).SetFocus
          If completedy(t%).Visible Then
              completedy(t%).SetFocus
           End If
                                    GoTo exiteditp
                                End If
                            End If
                            If group(tt%).ListIndex = -1 Then
                                msg = "Type(Group) and Total Value must be entered."
                                Call ShowApplicableContainers(completedy(t%))
'---- setfocus logic ----
'                                         completedy(t%).SetFocus
          If completedy(t%).Visible Then
              completedy(t%).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                        If Val(totalvalue(tt% + 18)) > 0 Then
                            If Not IsDate(DATERECOVERED(tt%)) Then
                                msg = "A valid DATE RECOVERED must be entered."
                                Call ShowApplicableContainers(completedy(t%))
'---- setfocus logic ----
'                                         completedy(t%).SetFocus
          If completedy(t%).Visible Then
              completedy(t%).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                    End If
                End If
            Next tt%
        End If
        TP$ = ""
        
        Select Case tempucr
            Case "200", "220", "250", "290", "280"
                TP$ = "Crimes Against Property"
            Case "39A", "39B", "39C", "39D"
                TP$ = "Gambling"
            Case "100"
                TP$ = "Kidnaping"
            Case "35A", "35B"
                TP$ = "Drug/Narcotic Offenses"
        End Select
        '===== Error 074
        If TP$ > "" Then
            foundmatch% = 0
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        foundmatch% = 1
                    End If
                End If
            Next tt%
            If foundmatch% = 0 And completedy(t%) Then
                msg = TP$ + " must have property data."
                Call ShowApplicableContainers(pucrlist(0))
'---- setfocus logic ----
'                         pucrlist(0).SetFocus
          If pucrlist(0).Visible Then
              pucrlist(0).SetFocus
           End If
                GoTo exiteditp
            End If
        End If
        
        If TP$ > "" Then
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Not completedy(t%) Then
                            If Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                msg = TP$ + " not completed must have associated value of None or Unknown on property tab."
                                Call ShowApplicableContainers(totalvalue(tt% + 0))
'---- setfocus logic ----
'                                         totalvalue(tt% + 0).SetFocus
          If totalvalue(tt% + 0).Visible Then
              totalvalue(tt% + 0).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                        If completedy(t%) Then
                            If Val(totalvalue(tt%)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Then
                                temperr = 0
                            Else
'                            If tempucr <> "35A" Then
'                                msg = TP$ + " completed must have associated value of Burned, Recovered, or Stolen on property tab."
'                                GoTo exiteditp
'                            End If
                            End If
                            If group(tt%).ListIndex = -1 Then
                                msg = "Type(Group) and Total Value must be entered."
                                Call ShowApplicableContainers(completedy(t%))
'---- setfocus logic ----
'                                         completedy(t%).SetFocus
          If completedy(t%).Visible Then
              completedy(t%).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                        If Val(totalvalue(tt% + 18)) > 0 Then
                            If Not IsDate(DATERECOVERED(tt%)) Then
                                msg = "A valid DATE RECOVERED must be entered."
                                Call ShowApplicableContainers(totalvalue(tt% + 18))
'---- setfocus logic ----
'                                         totalvalue(tt% + 18).SetFocus
          If totalvalue(tt% + 18).Visible Then
              totalvalue(tt% + 18).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                    End If
                End If
            Next tt%
        End If

        
        '===== Additional F 11
        '===== Data Element 7
        If tempucr = "39A" Or tempucr = "39B" Or tempucr = "39C" Or tempucr = "39D" Then
            TP$ = "Gambling Offenses"
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Not completedy(t%) Then
                            If Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                msg = TP$ + " not completed must have associated value of None or Unknown on property tab."
                                Call ShowApplicableContainers(totalvalue(tt% + 0))
'---- setfocus logic ----
'                                         totalvalue(tt% + 0).SetFocus
          If totalvalue(tt% + 0).Visible Then
              totalvalue(tt% + 0).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                        If completedy(t%) Then
                            If Val(totalvalue(tt% + 24)) = 0 Or Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                msg = TP$ + " completed must have associated value of Seized on property tab."
                                Call ShowApplicableContainers(totalvalue(tt% + 24))
'---- setfocus logic ----
'                                         totalvalue(tt% + 24).SetFocus
          If totalvalue(tt% + 24).Visible Then
              totalvalue(tt% + 24).SetFocus
           End If
                                GoTo exiteditp
                            End If
                            If group(tt%).ListIndex = -1 Then
                                msg = "Type(Group) and Total Value must be entered."
                                Call ShowApplicableContainers(group(tt%))
'---- setfocus logic ----
'                                         group(tt%).SetFocus
          If group(tt%).Visible Then
              group(tt%).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                    End If
                End If
            Next tt%
        End If
        
        '===== Additional F 13
        '===== Data Element 7
        If tempucr = "100" Then
            TP$ = "Kidnapping"
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Not completedy(t%) Then
                            If Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                msg = TP$ + " not completed must have associated value of None or Unknown on property tab."
                                Call ShowApplicableContainers(totalvalue(tt% + 0))
'---- setfocus logic ----
'                                         totalvalue(tt% + 0).SetFocus
          If totalvalue(tt% + 0).Visible Then
              totalvalue(tt% + 0).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                        If completedy(t%) Then
                            If (Val(totalvalue(tt%)) = 0 And Val(totalvalue(tt% + 6)) = 0 And Val(totalvalue(tt% + 12)) = 0 And Val(totalvalue(tt% + 18)) = 0 And Val(totalvalue(tt% + 24)) = 0 And Val(totalvalue(tt% + 30)) = 0) Or _
                                ((Val(totalvalue(tt%)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 30)) > 0) And Val(totalvalue(tt% + 6)) = 0 And Val(totalvalue(tt% + 12)) = 0 And Val(totalvalue(tt% + 24)) = 0) Then
                                temperr = 0
                            Else
                                msg = TP$ + " completed must have associated value of None, Recovered, Stolen, or Unknown on property tab."
                                Call ShowApplicableContainers(totalvalue(tt%))
'---- setfocus logic ----
'                                         totalvalue(tt%).SetFocus
          If totalvalue(tt%).Visible Then
              totalvalue(tt%).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                        If Val(totalvalue(tt%)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Then
                            If group(tt%).ListIndex = -1 Then
                                msg = "Type(Group) and Total Value must be entered."
                                Call ShowApplicableContainers(totalvalue(tt%))
'---- setfocus logic ----
'                                         totalvalue(tt%).SetFocus
          If totalvalue(tt%).Visible Then
              totalvalue(tt%).SetFocus
           End If
                                GoTo exiteditp
                            Else
                            If InStr(group(tt%).List(group(tt%).ListIndex), "77") = 0 Then
                                msg = "Group 77 must be selected for Kidnapping."
                                GoTo exiteditp
                            End If
                            End If
                        End If
                        If Val(totalvalue(tt% + 18)) > 0 Then
                            If Not IsDate(DATERECOVERED(tt%)) Then
                                msg = "A valid DATE RECOVERED must be entered."
                                Call ShowApplicableContainers(totalvalue(tt% + 18))
'---- setfocus logic ----
'                                         totalvalue(tt% + 18).SetFocus
          If totalvalue(tt% + 18).Visible Then
              totalvalue(tt% + 18).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                    End If
                End If
            Next tt%
        End If
        
        '===== Additional F 15
        If tempucr = "240" Then
            TP$ = "Motor Vehicle Theft"
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Not completedy(t%) Then
                            If Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                msg = TP$ + " not completed must have associated value of None or Unknown on property tab."
                                Call ShowApplicableContainers(totalvalue(tt% + 0))
'---- setfocus logic ----
'                                         totalvalue(tt% + 0).SetFocus
          If totalvalue(tt% + 0).Visible Then
              totalvalue(tt% + 0).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                        If completedy(t%) Then
                            If (Val(totalvalue(tt%)) > 0 Or Val(totalvalue(tt% + 18)) > 0) And Val(totalvalue(tt% + 6)) = 0 And Val(totalvalue(tt% + 12)) = 0 And Val(totalvalue(tt% + 24)) = 0 And Val(totalvalue(tt% + 30)) = 0 Then
                                temperr = 0
                            Else
                                msg = TP$ + " completed must have associated value of Recovered or Stolen on property tab."
                                Call ShowApplicableContainers(totalvalue(tt%))
'---- setfocus logic ----
'                                         totalvalue(tt%).SetFocus
          If totalvalue(tt%).Visible Then
              totalvalue(tt%).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                        If Val(totalvalue(tt% + 18)) > 0 Then
                            If group(tt%).ListIndex = -1 Then
                                msg = "Type(Group) and Total Value must be entered."
                                Call ShowApplicableContainers(totalvalue(tt% + 18))
'---- setfocus logic ----
'                                         totalvalue(tt% + 18).SetFocus
          If totalvalue(tt% + 18).Visible Then
              totalvalue(tt% + 18).SetFocus
           End If
                                GoTo exiteditp
                            End If
                            tempgroup = Mid$(group(tt%).List(group(tt%).ListIndex), InStr(group(tt%).List(group(tt%).ListIndex), "(") + 1, 2)
                            If tempgroup = "03" Or tempgroup = "05" Or tempgroup = "24" Or tempgroup = "28" Or tempgroup = "37" Then
                                If numvehicle(tt% + 6) = 0 Then
                                    msg = "A number of vehicles recovered must be entered."
                                    Call ShowApplicableContainers(group(tt%))
'---- setfocus logic ----
'                                             group(tt%).SetFocus
          If group(tt%).Visible Then
              group(tt%).SetFocus
           End If
                                    GoTo exiteditp
                                End If
                            End If
                            If Not IsDate(DATERECOVERED(tt%)) Then
                                msg = "A valid DATE RECOVERED must be entered."
                                Call ShowApplicableContainers(DATERECOVERED(tt%))
'---- setfocus logic ----
'                                         DATERECOVERED(tt%).SetFocus
          If DATERECOVERED(tt%).Visible Then
              DATERECOVERED(tt%).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                        If Val(totalvalue(tt%)) > 0 Then
                            If group(tt%).ListIndex = -1 Then
                                msg = "Type(Group) and Total Value must be entered."
                                Call ShowApplicableContainers(totalvalue(tt%))
'---- setfocus logic ----
'                                         totalvalue(tt%).SetFocus
          If totalvalue(tt%).Visible Then
              totalvalue(tt%).SetFocus
           End If
                                GoTo exiteditp
                            End If
                            tempgroup = Mid$(group(tt%).List(group(tt%).ListIndex), InStr(group(tt%).List(group(tt%).ListIndex), "(") + 1, 2)
                            If tempgroup = "03" Or tempgroup = "05" Or tempgroup = "24" Or tempgroup = "28" Or tempgroup = "37" Then
                                If numvehicle(tt%) = 0 Then
                                    msg = "A number of vehicles stolen must be entered."
                                    Call ShowApplicableContainers(totalvalue(tt%))
'---- setfocus logic ----
'                                             totalvalue(tt%).SetFocus
          If totalvalue(tt%).Visible Then
              totalvalue(tt%).SetFocus
           End If
                                    GoTo exiteditp
                                End If
                            Else
                                msg = "Invalid type (group) entered for Motor Vehicle Theft crime."
                                Call ShowApplicableContainers(totalvalue(tt%))
'---- setfocus logic ----
'                                         totalvalue(tt%).SetFocus
          If totalvalue(tt%).Visible Then
              totalvalue(tt%).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                    End If
                End If
            Next tt%
        End If
        
        '===== Additional F 21
        '===== Data Element 12
        If tempucr = "280" Then
            TP$ = "Stolen Property"
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Not completedy(t%) Then
                            If Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                msg = TP$ + " not completed must have associated value of None or Unknown on property tab."
                                Call ShowApplicableContainers(totalvalue(tt% + 0))
'---- setfocus logic ----
'                                         totalvalue(tt% + 0).SetFocus
          If totalvalue(tt% + 0).Visible Then
              totalvalue(tt% + 0).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                        If completedy(t%) Then
                            If (Val(totalvalue(tt%)) = 0 And Val(totalvalue(tt% + 6)) = 0 And Val(totalvalue(tt% + 12)) = 0 And Val(totalvalue(tt% + 18)) = 0 And Val(totalvalue(tt% + 24)) = 0 And Val(totalvalue(tt% + 30)) = 0) Or _
                                (Val(totalvalue(tt% + 18)) > 0 And Val(totalvalue(tt%)) = 0 And Val(totalvalue(tt% + 6)) = 0 And Val(totalvalue(tt% + 12)) = 0 And Val(totalvalue(tt% + 24)) = 0 And Val(totalvalue(tt% + 30)) = 0) Then
                                temperr = 0
                            Else
                                msg = TP$ + " completed must have associated value of None or Recovered on property tab."
                                Call ShowApplicableContainers(totalvalue(tt%))
'---- setfocus logic ----
'                                         totalvalue(tt%).SetFocus
          If totalvalue(tt%).Visible Then
              totalvalue(tt%).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                        If Val(totalvalue(tt% + 18)) > 0 Then
                            If group(tt%).ListIndex = -1 Then
                                msg = "Type(Group) and Total Value must be entered."
                                Call ShowApplicableContainers(totalvalue(tt% + 18))
'---- setfocus logic ----
'                                         totalvalue(tt% + 18).SetFocus
          If totalvalue(tt% + 18).Visible Then
              totalvalue(tt% + 18).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                        If Val(totalvalue(tt% + 18)) > 0 Then
                            If Not IsDate(DATERECOVERED(tt%)) Then
                                msg = "A valid DATE RECOVERED must be entered."
                                Call ShowApplicableContainers(totalvalue(tt% + 18))
'---- setfocus logic ----
'                                         totalvalue(tt% + 18).SetFocus
          If totalvalue(tt% + 18).Visible Then
              totalvalue(tt% + 18).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                    End If
                End If
            Next tt%
        End If
        
        For tt% = 0 To 5
            
            '===== Data Element 14
            If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) <> "35A" And Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) > "" Then
                If Val(totalvalue(tt% + 36)) > 0 Or _
                    (Val(totalvalue(tt% + 0)) = 0 And Val(totalvalue(tt% + 6)) = 0 And Val(totalvalue(tt% + 12)) = 0 And Val(totalvalue(tt% + 18)) = 0 And Val(totalvalue(tt% + 24)) = 0 And Val(totalvalue(tt% + 30)) = 0) Then
                    If group(tt%).ListIndex > -1 Or pucrlist(tt%).ListIndex > -1 Or description(tt%) > "" Then
                        If group(tt%).ListIndex > -1 Then
                            tempgrp = Mid$(group(tt%).List(group(tt%).ListIndex), InStr(group(tt%).List(group(tt%).ListIndex), "(") + 1, 2)
                        End If
                        If tempgrp <> "09" And tempgrp <> "22" And tempgrp <> "77" And tempgrp <> "99" Then
                            msg = "If UNKNOWN or Nothing selected in entry of PROPERTY tab, no other associated values may be selected (i.e. Type, Value, etc.)."
'---- setfocus logic ----
'                                     totalvalue(tt% + 36).SetFocus
          If totalvalue(tt% + 36).Visible Then
              totalvalue(tt% + 36).SetFocus
           End If
                            GoTo exiteditp
                        End If
                    End If
                End If
            End If
        
        Next tt%

        
    
        
    End If
    
Next t%

GoTo goodeditp

exiteditp:

editerr = 1
goodeditp:


Exit Sub
'RLB Bandaid
rlbErrp:
    If Err.Number = 5 Then Resume Next
    Resume
End Sub

Private Sub deleteroutine()
If frmLogin.IDELETE <> 1 And frmLogin.ISUPERVISOR <> 1 And frmLogin.SUPERVISOR <> 1 Then
    msg = MsgBox("Insufficient authority for this operation.", 48, "Genesis Error Log")
    Exit Sub
End If

    'RLB Code
    If UCase(reportingofficer(0).Text) <> UCase(frmLogin.userfullname) Then
        If UCase(reportingofficer(1).Text) <> UCase(frmLogin.userfullname) Then
            If UCase(approvingofficer.Text) <> UCase(frmLogin.userfullname) Then
                 If UCase(followupofficer.Text) <> UCase(frmLogin.userfullname) Then
                    
                    If (frmLogin.ISUPERVISOR <> 1) And (frmLogin.SUPERVISOR <> 1) Then
                        MsgBox "Since you are not recognized as an officer related to this report, you cannot alter this incident report.", vbOKOnly, "Genesis Error Log"
                        Exit Sub
                    End If
                    
                 End If
            End If
        End If
    End If
    '********
    
If incidentnumber = "" Then
    Exit Sub
End If
If Len(incidentnumber) < 12 And Len(incidentnumber) > 0 Then
    incidentnumber = incidentnumber + Space$(12 - Len(incidentnumber))
End If
msg = MsgBox("Are you sure you wish to delete this incident report?", 4, "Genesis Information Log")
If msg <> 6 Then
    Exit Sub
End If
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("select * from INCIDENTsuppORT where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
End If
On Error Resume Next
If IsDate(rs("exportdate")) Then
    msg = MsgBox("This incident report cannot be deleted because it has already been exported.", 48, "Genesis Error Log")
    db.Close
    Exit Sub
End If
While Not rs.EOF
    rs.Delete
    rs.MoveNext
Wend
Set rs = db.OpenRecordset("select * from incidentreportC where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
End If
While Not rs.EOF
    rs.Delete
    rs.MoveNext
Wend
Set rs = db.OpenRecordset("select * from incidentreportS where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
End If
While Not rs.EOF
    rs.Delete
    rs.MoveNext
Wend
Set rs = db.OpenRecordset("select * from incidentreportV where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
End If
While Not rs.EOF
    rs.Delete
    rs.MoveNext
Wend
Set rs = db.OpenRecordset("select * from incidentreportO where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
End If
While Not rs.EOF
    rs.Delete
    rs.MoveNext
Wend
Set rs = db.OpenRecordset("select * from supplemental where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
End If
While Not rs.EOF
    rs.Delete
    rs.MoveNext
Wend
Set rs = db.OpenRecordset("select * from supplementalSUPPORT where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
End If
While Not rs.EOF
    rs.Delete
    rs.MoveNext
Wend
Set rs = db.OpenRecordset("select * from VEHGUN where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
End If
While Not rs.EOF
    rs.Delete
    rs.MoveNext
Wend
Set rs = db.OpenRecordset("select * from BOOKING where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
End If
While Not rs.EOF
    rs.Delete
    rs.MoveNext
Wend
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub
Friend Sub findincident(op As Integer)
Dim db As Database, rs, rs2 As Recordset, ecc As Integer, lu As String
On Error GoTo oderror
od:
Set db = OpenDatabase(nwi + "incident.mdb")
incidentnumber = UCase(incidentnumber)
Set rs = db.OpenRecordset("Select * from incidentreporto where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
If Not rs.EOF Then
    'RLB Code
    IncidentRptLoadedFromDB = True
    '********
    rs.MoveFirst
Else
    'RLB Code
    IncidentRptLoadedFromDB = False
    '********
    On Error Resume Next
    db.Close
    Exit Sub
End If
fromfind = 1
incidentnumber = rs("incidentnumber")
For pp% = 1 To 5
    VUCRSEL(pp%) = ""
Next pp%
ct% = 0
ct2% = 0
For t% = 6 To 23
    ct% = ct% + 1
    ct2% = ct2% + 1
    If ct2% > 3 Then
        ct2% = 1
    End If
    For tt% = 1 To 3
        st% = tt%
        If Not IsNull(rs("ptypeofdrug" + Mid$(Str$(ct%), 2) + Mid$(Str$(tt%), 2))) Then
            For ttt% = 0 To drugtype(t%).ListCount - 1
                If rs("ptypeofdrug" + Mid$(Str$(ct%), 2) + Mid$(Str$(tt%), 2)) = Left$(drugtype(t%).List(ttt%), 1) Then
                    drugtype(t%).ListIndex = ttt%
                    ttt% = drugtype(t%).ListCount - 1
                End If
            Next ttt%
            drugamt(t%) = rs("pdrugamt" + Mid$(Str$(ct%), 2) + Mid$(Str$(tt%), 2))
            For ttt% = 0 To drugmeasurement(t%).ListCount - 1
                If rs("pdrugmeasurement" + Mid$(Str$(ct%), 2) + Mid$(Str$(tt%), 2)) = Left$(drugmeasurement(t%).List(ttt%), 2) Then
                    drugmeasurement(t%).ListIndex = ttt%
                    ttt% = drugmeasurement(t%).ListCount - 1
                End If
            Next ttt%
        End If
        t% = t% + 1
    Next tt%
    t% = t% - 1
Next t%
If Not IsNull(rs("JURISDICTIONTHEFT")) Then
    JURISDICTIONTHEFT = rs("JURISDICTIONTHEFT")
End If
If Not IsNull(rs("JURISDICTIONRECOVERY")) Then
    JURISDICTIONRECOVERY = rs("JURISDICTIONRECOVERY")
End If
subjectidentifiedyes = rs("subjectidentifiedyes")
subjectidentifiedno = rs("subjectidentifiedno")
subjectlocatedyes = rs("subjectlocatedyes")
subjectlocatedno = rs("subjectlocatedno")
If rs("active") = "X" Then
    active = 1
End If
If rs("admclosed") = "X" Then
    admclosed = 1
End If
If rs("unfounded") = "X" Then
    unfounded = 1
End If
If rs("arrestedunder18") = "X" Then
    arrestedunder18 = 1
End If
If rs("arrested18over") = "X" Then
    arrested18andover = 1
End If
If rs("exclearunder18") = "X" Then
    exclearunder18 = 1
End If
If rs("exclear18over") = "X" Then
    exclear18andover = 1
End If
offenderdeath = rs("offenderdeath")
noprosecution = rs("noprosecution")
extraditiondenied = rs("extraditiondenied")
victimdeclinescooperation = rs("victimdeclinescooperation")
juvenilenocustody = rs("juvenilenocustody")
If Not IsNull(rs("excleardate")) Then
    EXCEPTIONALCLEARANCEDATE = rs("excleardate")
End If
If Not IsNull(rs("reportingofficer1")) Then
    reportingofficer(0) = rs("reportingofficer1")
End If
If Not IsNull(rs("reportingdate1")) Then
    REPORTINGOFFICERDATE(0) = rs("reportingdate1")
End If
If Not IsNull(rs("reportingunit1")) Then
    reportingofficeRunit(0) = rs("reportingunit1")
End If
If Not IsNull(rs("reportingofficer2")) Then
    reportingofficer(1) = rs("reportingofficer2")
End If
If Not IsNull(rs("reportingdate2")) Then
    REPORTINGOFFICERDATE(1) = rs("reportingdate2")
End If
If Not IsNull(rs("reportingunit2")) Then
    reportingofficeRunit(1) = rs("reportingunit2")
End If
followupyes = rs("followupyes")
followupno = rs("followupno")
If Not IsNull(rs("followupofficer")) Then
    followupofficer = rs("followupofficer")
End If
If Not IsNull(rs("approvingofficer")) Then
    approvingofficer = rs("approvingofficer")
End If
If Not IsNull(rs("approvingdate")) Then
    APPROVINGOFFICERDATE = rs("approvingdate")
End If
If Not IsNull(rs("approvingunit")) Then
    approvingofficeRunit = rs("approvingunit")
End If
If Not IsNull(rs("followupdate")) Then
    FOLLOWUPOFFICERDATE = rs("followupdate")
End If
If Not IsNull(rs("followupunit")) Then
    FOLLOWUPOFFICERUNIT = rs("followupunit")
End If
BIAS.ListIndex = -1
If rs("BIAS") > "" Then
    For t% = 0 To BIAS.ListCount - 1
        If rs("bias") = Mid$(BIAS.List(t%), InStr(BIAS.List(t%), "(") + 1, 2) Then
            BIAS.ListIndex = t%
            t% = BIAS.ListCount - 1
        End If
    Next t%
End If
NARRATIVE.Text = rs("narrative")
For t% = 0 To 41
    Select Case t%
        Case 0 To 5
            If Not IsNull(rs("stolenvalue" + Mid$(Str$((t% Mod 6) + 1), 2))) Then
                If rs("stolenvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = 9999999 Then
                    totalvalue(t%) = "X"
                Else
                    totalvalue(t%) = rs("stolenvalue" + Mid$(Str$((t% Mod 6) + 1), 2))
                End If
            End If
        Case 6 To 11
            If Not IsNull(rs("damagedvalue" + Mid$(Str$((t% Mod 6) + 1), 2))) Then
                If rs("damagedvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = 9999999 Then
                    totalvalue(t%) = "X"
                Else
                    totalvalue(t%) = rs("damagedvalue" + Mid$(Str$((t% Mod 6) + 1), 2))
                End If
            End If
        Case 12 To 17
            If Not IsNull(rs("burnedvalue" + Mid$(Str$((t% Mod 6) + 1), 2))) Then
                If rs("burnedvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = 9999999 Then
                    totalvalue(t%) = "X"
                Else
                    totalvalue(t%) = rs("burnedvalue" + Mid$(Str$((t% Mod 6) + 1), 2))
                End If
            End If
        Case 18 To 23
            If Not IsNull(rs("recoveredvalue" + Mid$(Str$((t% Mod 6) + 1), 2))) Then
                If rs("recoveredvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = 9999999 Then
                    totalvalue(t%) = "X"
                Else
                    totalvalue(t%) = rs("recoveredvalue" + Mid$(Str$((t% Mod 6) + 1), 2))
                End If
            End If
        Case 24 To 29
            If Not IsNull(rs("seizedvalue" + Mid$(Str$((t% Mod 6) + 1), 2))) Then
                If rs("seizedvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = 9999999 Then
                    totalvalue(t%) = "X"
                Else
                    totalvalue(t%) = rs("seizedvalue" + Mid$(Str$((t% Mod 6) + 1), 2))
                End If
            End If
        Case 30 To 35
            If Not IsNull(rs("counterfeitvalue" + Mid$(Str$((t% Mod 6) + 1), 2))) Then
                If rs("counterfeitvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = 9999999 Then
                    totalvalue(t%) = "X"
                Else
                    totalvalue(t%) = rs("counterfeitvalue" + Mid$(Str$((t% Mod 6) + 1), 2))
                End If
            End If
        Case 36 To 41
            If Not IsNull(rs("unknownvalue" + Mid$(Str$((t% Mod 6) + 1), 2))) Then
                If rs("unknownvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = 9999999 Then
                    totalvalue(t%) = "X"
                Else
                    totalvalue(t%) = rs("unknownvalue" + Mid$(Str$((t% Mod 6) + 1), 2))
                End If
            End If
    End Select
Next t%
For t% = 0 To 5
    If Not IsNull(rs("type" + Mid$(Str$(t% + 1), 2))) Then
        description(t%) = rs("type" + Mid$(Str$(t% + 1), 2))
        FOUNDMAJMIN = False
        For tt% = 0 To majorlist(t%).ListCount - 1
            If majorlist(t%).List(tt%) = rs("major" + Mid$(Str$(t% + 1), 2)) Then
                FOUNDMAJMIN = True
                majorlist(t%).ListIndex = tt%
                tt% = majorlist(t%).ListCount - 1
            End If
        Next tt%
        If Not FOUNDMAJMIN Then
            majorlist(t%).AddItem rs("major" + Mid$(Str$(t% + 1), 2))
            majorlist(t%).ListIndex = majorlist(t%).ListCount - 1
        End If
        Call setminorlist(t%)
        FOUNDMAJMIN = False
        For tt% = 0 To minorlist(t%).ListCount - 1
            If minorlist(t%).List(tt%) = rs("minor" + Mid$(Str$(t% + 1), 2)) Then
                FOUNDMAJMIN = True
                minorlist(t%).ListIndex = tt%
                tt% = minorlist(t%).ListCount - 1
            End If
        Next tt%
        If Not FOUNDMAJMIN Then
            minorlist(t%).AddItem rs("MINOR" + Mid$(Str$(t% + 1), 2))
            minorlist(t%).ListIndex = minorlist(t%).ListCount - 1
        End If
       
    End If
Next t%
Call figure

Set rs = db.OpenRecordset("Select * from incidentreportc where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    rs.Edit
    On Error Resume Next
Else
    GoTo DOV
    On Error Resume Next
    db.Close
    fromfind = 0
    Exit Sub
End If
TIMEOFOFFENSE(0) = rs("timeofoffense1")
incidentdate(0) = rs("dateofoffense1")
TIMEOFOFFENSE(1) = rs("timeofoffense2")
If Not IsNull(rs("dateofoffense2")) Then
    incidentdate(1) = rs("dateofoffense2")
End If
For t% = 0 To 2
    If Not IsNull(rs("offense" + Mid$(Str$(t% + 1), 2))) Or Not IsNull(rs("premise" + Mid$(Str$(t% + 1), 2))) Then
        
        'RLC Code
        If Not (AutoSelectOffense(t%, rs("offense" + Mid$(Str$(t% + 1), 2)))) Then
            pickoffense(t%).ListIndex = -1
        End If
        '*******
        
        completedy(t%) = rs("completedyes" + Mid$(Str$(t% + 1), 2))
        FORCEDENTRYY(t%) = rs("forcedentryyes" + Mid$(Str$(t% + 1), 2))
        completedn(t%) = rs("completedno" + Mid$(Str$(t% + 1), 2))
        FORCEDENTRYN(t%) = rs("forcedentryno" + Mid$(Str$(t% + 1), 2))
        For vv% = 1 To premise(t%).ListItems.Count
            premise(t%).ListItems(vv%).Selected = False
        Next vv%
        If rs("premise" + Mid$(Str$(t% + 1), 2)) > "" Then
            For tt% = 1 To premise(t%).ListItems.Count
                If rs("premise" + Mid$(Str$(t% + 1), 2)) = Mid$(premise(t%).ListItems(tt%), InStr(premise(t%).ListItems(tt%), "(") + 1, 2) Then
                    premise(t%).ListItems(tt%).Selected = True
                    premise(t%).ListItems(tt%).EnsureVisible
                End If
            Next tt%
        End If
        If rs("premise" + Mid$(Str$(t% + 1), 2) + "a") > "" Then
            For tt% = 1 To premise(t%).ListItems.Count
                If rs("premise" + Mid$(Str$(t% + 1), 2) + "a") = Mid$(premise(t%).ListItems(tt%), InStr(premise(t%).ListItems(tt%), "(") + 1, 2) Then
                    premise(t%).ListItems(tt%).Selected = True
                    premise(t%).ListItems(tt%).EnsureVisible
                End If
            Next tt%
        End If
        entered(t%) = rs("entered" + Mid$(Str$(t% + 1), 2))
    End If
Next t%
individual = rs("individual")
business = rs("business")
financialinstitution = rs("financialinstitution")
other = rs("other")
government = rs("government")
unknown = rs("unknown")
religiousorganization = rs("religiousorganization")
policeofficer = rs("policeofficer")
societypublic = rs("societypublic")
incidentlocation = rs("incidentlocation")
CLOCATIONNUMBER = rs("clocationnumber")
incidentzipcode = rs("incidentzipcode")
idx5 = 0
For i% = 1 To weapontype.ListItems.Count
    If rs("weapontype1") = Mid$(weapontype.ListItems(i%), InStr(weapontype.ListItems(i%), "(") + 1, 2) Then
        idx5 = idx5 + 1
        weapontype.ListItems(i%).Selected = True
        weapontype.ListItems(i%).EnsureVisible
        If Not IsNull(rs("automatic1")) Then
            automatic(idx5) = rs("automatic1")
        End If
    End If
    If rs("weapontype2") = Mid$(weapontype.ListItems(i%), InStr(weapontype.ListItems(i%), "(") + 1, 2) Then
        idx5 = idx5 + 1
        weapontype.ListItems(i%).Selected = True
        weapontype.ListItems(i%).EnsureVisible
        If Not IsNull(rs("automatic2")) Then
            automatic(idx5) = rs("automatic2")
        End If
    End If
    If rs("weapontype3") = Mid$(weapontype.ListItems(i%), InStr(weapontype.ListItems(i%), "(") + 1, 2) Then
        idx5 = idx5 + 1
        weapontype.ListItems(i%).Selected = True
        weapontype.ListItems(i%).EnsureVisible
        If Not IsNull(rs("automatic3")) Then
            automatic(idx5) = rs("automatic3")
        End If
    End If
    If rs("weapontype1") = Mid$(weapontype.ListItems(i%), InStr(weapontype.ListItems(i%), "(") + 1, 2) Or _
       rs("weapontype2") = Mid$(weapontype.ListItems(i%), InStr(weapontype.ListItems(i%), "(") + 1, 2) Or _
       rs("weapontype3") = Mid$(weapontype.ListItems(i%), InStr(weapontype.ListItems(i%), "(") + 1, 2) Then
    Else
        weapontype.ListItems(i%).Selected = False
    End If
Next i%
If Not IsNull(rs("automatic1")) Then
    automatic(1) = rs("automatic1")
Else
    automatic(1) = ""
End If
If Not IsNull(rs("automatic2")) Then
    automatic(2) = rs("automatic2")
Else
    automatic(2) = ""
End If
If Not IsNull(rs("automatic3")) Then
    automatic(3) = rs("automatic3")
Else
    automatic(3) = ""
End If
If Not IsNull(rs("dispatchdate")) Then
    dispatchdate = rs("dispatchdate")
End If
If Not IsNull(rs("dispatchtime")) Then
    DISPATCHTIME = rs("dispatchtime")
End If
If Not IsNull(rs("arrivaltime")) Then
    TIMEARRIVED = rs("arrivaltime")
End If
If Not IsNull(rs("departuretime")) Then
    DEPARTINGTIME = rs("departuretime")
End If
CLOCATIONNUMBER = rs("incidentlocatioNnumber")
vsname(0) = rs("cname")
address(0) = rs("caddress")
city(0) = rs("ccity")
state(0) = rs("cstate")
zipcode(0) = rs("czipcode")
resident(0).ListIndex = -1
race(0).ListIndex = -1
sex(0).ListIndex = -1
ethnicity(0).ListIndex = -1
For t% = 0 To resident(0).ListCount - 1
    If Left$(resident(0).List(t%), 1) = rs("cresident") Then
        resident(0).ListIndex = t%
        t% = resident(0).ListCount - 1
    End If
Next t%
For t% = 0 To race(0).ListCount - 1
    If Left$(race(0).List(t%), 1) = rs("cRACE") Then
        race(0).ListIndex = t%
        t% = race(0).ListCount - 1
    End If
Next t%
For t% = 0 To sex(0).ListCount - 1
    If Left$(sex(0).List(t%), 1) = rs("cSEX") Then
        sex(0).ListIndex = t%
        t% = sex(0).ListCount - 1
    End If
Next t%
age(0) = rs("cage")

For t% = 0 To ethnicity(0).ListCount - 1
    If Left$(ethnicity(0).List(t%), 1) = rs("cETHNICITY") Then
        ethnicity(0).ListIndex = t%
        t% = ethnicity(0).ListCount - 1
    End If
Next t%
For t% = 0 To 2
    relationship(t%).ListIndex = -1
    For tt% = 0 To relationship(t%).ListCount - 1
        relationship(t%).Selected(tt%) = False
    Next tt%
    If rs("crelationship" + Mid$(Str$(t% + 1), 2)) > "" Then
        For tt% = 0 To relationship(t%).ListCount - 1
            If rs("crelationship" + Mid$(Str$(t% + 1), 2)) = Mid$(relationship(t%).List(tt%), InStr(relationship(t%).List(tt%), "(") + 1, 2) Then
                relationship(t%).ListIndex = tt%
                relationship(t%).Selected(tt%) = True
                tt% = relationship(t%).ListCount - 1
            End If
        Next tt%
    End If
Next t%
HOMEDAYPHONE(0) = rs("cdayhomephone")
WORKDAYPHONE(0) = rs("cdayworkphone")
HOMENIGHTPHONE(0) = rs("cnighthomephone")
WORKNIGHTPHONE(0) = rs("cnightworkphone")
DOV:
Set rs = db.OpenRecordset("Select * from incidentreportv where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
Else
    db.Close
    fromfind = 0
    Exit Sub
End If
vsname(1) = rs("vname")
address(1) = rs("vaddress")
city(1) = rs("vcity")
state(1) = rs("vstate")
zipcode(1) = rs("vzipcode")
ht(0) = rs("VHEIGHT")
weight(0) = rs("VWEIGHT")
hair(0) = rs("VHAIR")
eyes(0) = rs("VEYES")
peculiarities(0) = rs("VPECULIARITIES")
resident(1).ListIndex = -1
race(1).ListIndex = -1
sex(1).ListIndex = -1
ethnicity(1).ListIndex = -1
For t% = 0 To resident(1).ListCount - 1
    If Left$(resident(1).List(t%), 1) = rs("Vresident") Then
        resident(1).ListIndex = t%
        t% = resident(1).ListCount - 1
    End If
Next t%
For t% = 0 To race(1).ListCount - 1
    If Left$(race(1).List(t%), 1) = rs("VRACE") Then
        race(1).ListIndex = t%
        t% = race(1).ListCount - 1
    End If
Next t%
For t% = 0 To sex(1).ListCount - 1
    If Left$(sex(1).List(t%), 1) = rs("VSEX") Then
        sex(1).ListIndex = t%
        t% = sex(1).ListCount - 1
    End If
Next t%
age(1) = rs("Vage")
For t% = 0 To ethnicity(1).ListCount - 1
    If Left$(ethnicity(1).List(t%), 1) = rs("VETHNICITY") Then
        ethnicity(1).ListIndex = t%
        t% = ethnicity(1).ListCount - 1
    End If
Next t%
For t% = 10 To 12
    For tt% = 0 To relationship(t%).ListCount - 1
        relationship(t%).Selected(tt%) = False
    Next tt%
    relationship(t%).ListIndex = -1
    If rs("Vrelationship" + Mid$(Str$(t% - 9), 2)) > "" Then
        For tt% = 0 To relationship(t%).ListCount - 1
            If rs("Vrelationship" + Mid$(Str$(t% - 9), 2)) = Mid$(relationship(t%).List(tt%), InStr(relationship(t%).List(tt%), "(") + 1, 2) Then
                relationship(t%).ListIndex = tt%
                relationship(t%).Selected(tt%) = True
                tt% = relationship(t%).ListCount - 1
            End If
        Next tt%
    End If
Next t%
vlocationumber = rs("vlocationnumber")
If Not IsNull(rs("vhomephoneDAY")) Then
    HOMEDAYPHONE(1) = rs("vhomephoneDAY")
End If
If Not IsNull(rs("vworkphoneDAY")) Then
    WORKDAYPHONE(1) = rs("vworkphoneDAY")
End If
If Not IsNull(rs("vhomephoneNIGHT")) Then
    HOMENIGHTPHONE(1) = rs("vhomephoneNIGHT")
End If
If Not IsNull(rs("vworkphoneNIGHT")) Then
    WORKNIGHTPHONE(1) = rs("vworkphoneNIGHT")
End If
If Not IsNull(rs("computerequipment")) Then
    computerequipment(0) = rs("computerequipment")
Else
    computerequipment(0) = 0
End If
For t% = 1 To injury.ListItems.Count
    For tt% = 1 To 5
        If Not IsNull(rs("typeofinjury" + Mid$(Str$(tt%), 2))) Then
            If rs("typeofinjury" + Mid$(Str$(tt%), 2)) = Mid$(injury.ListItems(t%), InStr(injury.ListItems(t%), "(") + 1, 1) Then
                injury.ListItems(t%).Selected = True
                injury.ListItems(t%).EnsureVisible
            End If
        End If
    Next tt%
Next t%
If rs("vvisibleinjuryyes") = "X" Then
    VISIBLEINJURYYES = True
End If
If rs("vvisibleinjuryno") = "X" Then
    VISIBLEINJURYNO = True
End If
If rs("vNONVISibleinjuryyes") = "X" Then
    NONVISIBLEINJURYYES = True
End If
If rs("vNONVISibleinjuryno") = "X" Then
    NONVISIBLEINJURYNO = True
End If
If rs("valcoholyes") = "X" Then
    alcoholyes(0) = True
End If
If rs("valcoholNO") = "X" Then
    alcoholno(0) = True
End If
If rs("valcoholUNKNOWN") = "X" Then
    alcoholunknown(0) = True
End If
If rs("vDRUGSyes") = "X" Then
    drugsyes(0) = True
End If
If rs("vDRUGSNO") = "X" Then
    drugsno(0) = True
End If
If rs("vDRUGSUNKNOWN") = "X" Then
    drugsunknown(0) = True
End If
If rs("vtwomanvehicle") = "X" Then
    TWOMANVEHICLE = 1
End If
If rs("vonemanvehicle") = "X" Then
    ONEMANVEHICLE = 1
End If
If Not IsNull(rs("vtypeofdrug")) Then
    For t% = 0 To drugtype(0).ListCount - 1
        If Left$(drugtype(0).List(t%), 1) = rs("vtypeofdrug") Then
            drugtype(0).ListIndex = t%
            t% = drugtype(0).ListCount
        End If
    Next t%
End If
If Not IsNull(rs("Vdrugamt")) Then
    drugamt(0) = rs("Vdrugamt")
End If
For ttt% = 0 To drugmeasurement(0).ListCount - 1
    If rs("Vdrugmeasurement") = Left$(drugmeasurement(0).List(ttt%), 2) Then
        drugmeasurement(0).ListIndex = ttt%
        ttt% = drugmeasurement(0).ListCount - 1
    End If
Next ttt%
If rs("vdetective") = "X" Then
    DETECTIVE = 1
End If
If rs("vother") = "X" Then
    TODOTHER = 1
End If
If rs("valone") = "X" Then
    ALONE = 1
End If
If rs("vassisted") = "X" Then
    ASSISTED = 1
End If

Set rs = db.OpenRecordset("Select * from incidentreports where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
Else
    db.Close
    fromfind = 0
    Exit Sub
End If
vsname(2) = rs("sname")
address(2) = rs("saddress")
city(2) = rs("scity")
state(2) = rs("sstate")
zipcode(2) = rs("szipcode")
ht(1) = rs("SHEIGHT")
weight(1) = rs("SWEIGHT")
hair(1) = rs("SHAIR")
eyes(1) = rs("SEYES")
If Not IsNull(rs("sbirthdate")) Then
    BIRTHDATE = rs("SBIRTHDATE")
End If
peculiarities(1) = rs("SPECULIARITIES")
If Not IsNull(rs("computerequipment")) Then
    computerequipment(1) = rs("computerequipment")
Else
    computerequipment(1) = 0
End If
For t% = 0 To race(2).ListCount - 1
    If Left$(race(2).List(t%), 1) = rs("SRACE") Then
        race(2).ListIndex = t%
        t% = race(2).ListCount - 1
    End If
Next t%
For t% = 0 To sex(2).ListCount - 1
    If Left$(sex(2).List(t%), 1) = rs("SSEX") Then
        sex(2).ListIndex = t%
        t% = sex(2).ListCount - 1
    End If
Next t%
age(2) = rs("Sage")
For t% = 0 To ethnicity(2).ListCount - 1
    If Left$(ethnicity(2).List(t%), 1) = rs("SETHNICITY") Then
        ethnicity(2).ListIndex = t%
        t% = ethnicity(2).ListCount - 1
    End If
Next t%
vlocationumber = rs("Slocationnumber")
If rs("Salcoholyes") = "X" Then
    alcoholyes(1) = True
End If
If rs("SalcoholNO") = "X" Then
    alcoholno(1) = True
End If
If rs("SalcoholUNKNOWN") = "X" Then
    alcoholunknown(1) = True
End If
If rs("SDRUGSyes") = "X" Then
    drugsyes(1) = True
End If
If rs("SDRUGSNO") = "X" Then
    drugsno(1) = True
End If
If rs("SDRUGSUNKNOWN") = "X" Then
    drugsunknown(1) = True
End If
If Not IsNull(rs("stypeofdrug")) Then
    For t% = 0 To drugtype(3).ListCount - 1
        If Left$(drugtype(3).List(t%), 1) = rs("stypeofdrug") Then
            drugtype(3).ListIndex = t%
            t% = drugtype(3).ListCount
        End If
    Next t%
End If
If Not IsNull(rs("Sdrugamt")) Then
    drugamt(3) = rs("Sdrugamt")
End If
For ttt% = 0 To drugmeasurement(3).ListCount - 1
    If rs("Sdrugmeasurement") = Left$(drugmeasurement(3).List(ttt%), 2) Then
        drugmeasurement(3).ListIndex = ttt%
        ttt% = drugmeasurement(3).ListCount - 1
    End If
Next ttt%
If rs("ssuspect") = "X" Then
    SUSPECT = 1
End If
If rs("srunaway") = "X" Then
    RUNAWAY = 1
End If
If rs("swanted") = "X" Then
    WANTED = 1
End If
If rs("sarrest") = "X" Then
    ARREST = 1
End If
If rs("swarrant") = "X" Then
    WARRANT = 1
End If
If rs("sjail") = "X" Then
    JAIL = 1
End If
If rs("ssummons") = "X" Then
    SUMMONS = 1
End If
If rs("swarrant") = "X" Then
    WARRANT = 1
End If
If rs("sarrestednearoffenseyes") = "X" Then
    ARRESTEDNEARYES = True
End If
If rs("sarrestednearoffenseno") = "X" Then
    ARRESTEDNEARNO = True
End If
totalnumberarrested = rs("totalarrested")
If Not IsNull(rs("dateofarrest")) Then
    DATEOFARREST = rs("dateofarrest")
End If
TIMEOFARREST = rs("timeofarrest")
'----- support
'Set RS = DB.OpenRecordset("Select * from incidentsupport where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
Set rs = db.OpenRecordset("Select * from incidentsupport where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
Else
    db.Close
    fromfind = 0
    Exit Sub
End If
For t% = 0 To 5
    pucrlist(t%).clear
Next t%
If Not rs.EOF Then
    For t% = 1 To 2
        If Not IsNull(rs("Vdrugamt" + Mid$(Str$(t% + 1), 2))) Then
            drugamt(t%) = rs("Vdrugamt" + Mid$(Str$(t% + 1), 2))
        End If
        For ttt% = 0 To drugmeasurement(t%).ListCount - 1
            If rs("Vdrugmeasurement" + Mid$(Str$(t% + 1), 2)) = Left$(drugmeasurement(t%).List(ttt%), 2) Then
                drugmeasurement(t%).ListIndex = ttt%
                ttt% = drugmeasurement(t%).ListCount - 1
            End If
        Next ttt%
    Next t%
    For t% = 4 To 5
        If Not IsNull(rs("Sdrugamt" + Mid$(Str$(t% - 2), 2))) Then
            drugamt(t%) = rs("Sdrugamt" + Mid$(Str$(t% - 2), 2))
        End If
        For ttt% = 0 To drugmeasurement(t%).ListCount - 1
            If rs("Sdrugmeasurement" + Mid$(Str$(t% - 2), 2)) = Left$(drugmeasurement(t%).List(ttt%), 2) Then
                drugmeasurement(t%).ListIndex = ttt%
                ttt% = drugmeasurement(t%).ListCount - 1
            End If
        Next ttt%
    Next t%
    If Not IsNull(rs("vtypeofdrug2")) Then
        For t% = 0 To drugtype(1).ListCount - 1
            If Left$(drugtype(1).List(t%), 1) = rs("vtypeofdrug2") Then
                drugtype(1).ListIndex = t%
                t% = drugtype(1).ListCount
            End If
        Next t%
    End If
    If Not IsNull(rs("vtypeofdrug3")) Then
        For t% = 0 To drugtype(2).ListCount - 1
            If Left$(drugtype(2).List(t%), 1) = rs("vtypeofdrug3") Then
                drugtype(2).ListIndex = t%
                t% = drugtype(2).ListCount
            End If
        Next t%
    End If
    If Not IsNull(rs("stypeofdrug2")) Then
        For t% = 0 To drugtype(4).ListCount - 1
            If Left$(drugtype(4).List(t%), 1) = rs("stypeofdrug2") Then
                drugtype(4).ListIndex = t%
                t% = drugtype(4).ListCount
            End If
        Next t%
    End If
    If Not IsNull(rs("stypeofdrug3")) Then
        For t% = 0 To drugtype(5).ListCount - 1
            If Left$(drugtype(5).List(t%), 1) = rs("stypeofdrug3") Then
                drugtype(5).ListIndex = t%
                t% = drugtype(5).ListCount
            End If
        Next t%
    End If
    onpaper = rs("onpaper")
    locali = rs("local")
    For t% = 0 To 4
        For xx% = 1 To HOMOCIDE(t%).ListItems.Count
            HOMOCIDE(t%).ListItems(xx%).Selected = False
        Next xx%
        If Not IsNull(rs("homocide1" + Mid$(Str$(t% + 1), 2))) And rs("homocide1" + Mid$(Str$(t% + 1), 2)) > "" Then
            For tt% = 1 To HOMOCIDE(t%).ListItems.Count
                If rs("homocide1" + Mid$(Str$(t% + 1), 2)) = Mid$(HOMOCIDE(t%).ListItems(tt%), InStr(HOMOCIDE(t%).ListItems(tt%), "(") + 1, 2) Then
                    FROMKEY = 1
                    HOMOCIDE(t%).ListItems(tt%).Selected = True
                    tt% = HOMOCIDE(t%).ListItems.Count
                End If
            Next tt%
        End If
        If Not IsNull(rs("homocide2" + Mid$(Str$(t% + 1), 2))) And rs("homocide2" + Mid$(Str$(t% + 1), 2)) > "" Then
            For tt% = 1 To HOMOCIDE(t%).ListItems.Count
                If rs("homocide2" + Mid$(Str$(t% + 1), 2)) = Mid$(HOMOCIDE(t%).ListItems(tt%), InStr(HOMOCIDE(t%).ListItems(tt%), "(") + 1, 2) Then
                    FROMKEY = 1
                    HOMOCIDE(t%).ListItems(tt%).Selected = True
                    tt% = HOMOCIDE(t%).ListItems.Count
                End If
            Next tt%
        End If
        additional(t%).ListIndex = -1
        If rs("additional" + Mid$(Str$(t% + 1), 2)) > "" Then
            For tt% = 0 To additional(t%).ListCount - 1
                If rs("additional" + Mid$(Str$(t% + 1), 2)) = Mid$(additional(t%).List(tt%), InStr(additional(t%).List(tt%), "(") + 1, 1) Then
                    FROMKEY = 1
                    additional(t%).ListIndex = tt%
                    tt% = additional(t%).ListCount - 1
                End If
            Next tt%
        End If
        UCRLIST(t%).ListIndex = -1
        For uu% = 1 To sublist(t%).ListItems.Count
            sublist(t%).ListItems(uu%).Selected = False
        Next uu%
        tempucr = ""
        If rs("ucr" + Mid$(Str$(t% + 1), 2)) > "" Then
            tempucr = rs("ucr" + Mid$(Str$(t% + 1), 2))
            For tt% = 0 To UCRLIST(t%).ListCount - 1
                If tempucr = Mid$(UCRLIST(t%).List(tt%), InStr(UCRLIST(t%).List(tt%), "(") + 1, 3) Then
                    FROMKEY = 1
                    UCRLIST(t%).ListIndex = tt%
                    UCRLIST(t%).Selected(tt%) = True
                    For ttt% = 0 To 5
                        foundit% = 0
                        For yy% = 0 To pucrlist(ttt%).ListCount - 1
                            If UCRLIST(t%).List(tt%) = pucrlist(ttt%).List(yy%) Then
                                foundit% = 1
                                yy% = pucrlist(ttt%).ListCount - 1
                            End If
                        Next yy%
                        If foundit% = 0 Then
                            tempucr = Mid$(UCRLIST(t%).List(tt%), InStr(UCRLIST(t%).List(tt%), "(") + 1, 3)
                            Select Case tempucr
                                Case "100", "120", "200", "210", "220", "23A", "23B", "23C", "23D", "23E", "23F", "23G", "23H", "240", "250", "26A", "26B", "26C", "26D", "26E", "270", "280", "290", "35A", "35B", "39A", "39B", "39C", "39D", "510"
                                    pucrlist(ttt%).AddItem UCRLIST(t%).List(tt%)
                            End Select
                        End If
                    Next ttt%
                    tt% = UCRLIST(t%).ListCount - 1
                End If
            Next tt%
        End If
        For uu% = 1 To 3
            If rs("subcodes" + Mid$(Str$(t% + 1), 2) + Mid$(Str$(uu%), 2)) > "" Then
                tempsub = rs("subcodes" + Mid$(Str$(t% + 1), 2) + Mid$(Str$(uu%), 2))
                For tt% = 1 To sublist(t%).ListItems.Count
                    If tempsub = Left(sublist(t%).ListItems(tt%), 1) Then
                        FROMKEY = 1
                        sublist(t%).ListItems(tt%).Selected = True
                        tt% = sublist(t%).ListItems.Count
                    End If
                Next tt%
            End If
        Next uu%
        For tt% = 1 To 3
            If tempucr = "09A" Or tempucr = "09B" Or tempucr = "100" Or tempucr = "120" Or tempucr = "11A" Or tempucr = "11B" Or tempucr = "11C" Or tempucr = "11D" Or tempucr = "13A" Or tempucr = "13B" Or tempucr = "13C" Then
                If rs("activity" + Mid$(Str$(t% + 1), 2) + Mid$(Str$(tt%), 2)) > "" Then
                  For ttt% = 1 To gactivity(t%).ListItems.Count
                        If rs("activity" + Mid$(Str$(t% + 1), 2) + Mid$(Str$(tt%), 2)) = Mid$(gactivity(t%).ListItems(ttt%), InStr(gactivity(t%).ListItems(ttt%), "(") + 1, 1) Then
                            gactivity(t%).ListItems(ttt%).Selected = True
                            gactivity(t%).ListItems(ttt%).EnsureVisible
                        End If
                    Next ttt%
                End If
            Else
                If rs("activity" + Mid$(Str$(t% + 1), 2) + Mid$(Str$(tt%), 2)) > "" Then
                    For ttt% = 1 To activity(t%).ListItems.Count
                        If rs("activity" + Mid$(Str$(t% + 1), 2) + Mid$(Str$(tt%), 2)) = Mid$(activity(t%).ListItems(ttt%), InStr(activity(t%).ListItems(ttt%), "(") + 1, 1) Then
                            activity(t%).ListItems(ttt%).Selected = True
                            activity(t%).ListItems(ttt%).EnsureVisible
                        End If
                    Next ttt%
                End If
            End If
        Next tt%
    Next t%
    For t% = 0 To 5
        pucrlist(t%).ListIndex = -1
        If rs("pucr" + Mid$(Str$(t% + 1), 2)) > "" Then
            For tt% = 0 To pucrlist(t%).ListCount - 1
                If rs("pucr" + Mid$(Str$(t% + 1), 2)) = Mid$(pucrlist(t%).List(tt%), InStr(pucrlist(t%).List(tt%), "(") + 1, 3) Then
                    pucrlist(t%).ListIndex = tt%
                    tt% = pucrlist(t%).ListCount - 1
                End If
            Next tt%
        End If
        group(t%).ListIndex = -1
        For tt% = 0 To group(t%).ListCount - 1
            If rs("GROUP" + Mid$(Str$(t% + 1), 2)) = Mid$(group(t%).List(tt%), InStr(group(t%).List(tt%), "(") + 1, 2) Then
                group(t%).ListIndex = tt%
                tt% = group(t%).ListCount - 1
            End If
        Next tt%
        If Not IsNull(rs("numvehicles" + Mid$(Str$(t% + 1), 2))) Then
            numvehicle(t%) = rs("numvehicles" + Mid$(Str$(t% + 1), 2))
        End If
        If Not IsNull(rs("daterecovered" + Mid$(Str$(t% + 1), 2))) Then
            DATERECOVERED(t%) = rs("daterecovered" + Mid$(Str$(t% + 1), 2))
        End If
    Next t%
    For t% = 6 To 11
        If Not IsNull(rs("numvehicles" + Mid$(Str$(t% - 5), 2))) Then
            numvehicle(t%) = rs("numvehicler" + Mid$(Str$(t% - 5), 2))
        End If
    Next t%
    vucrlist.ListItems.clear
    For vv% = 0 To 4
        For vvv% = 0 To UCRLIST(vv%).ListCount - 1
            If UCRLIST(vv%).Selected(vvv%) = True Then
                Set itmx2 = vucrlist.ListItems.add(, , UCRLIST(vv%).List(vvv%))
                vvv% = UCRLIST(vv%).ListCount - 1
            End If
        Next vvv%
    Next vv%
    For tt% = 1 To vucrlist.ListItems.Count
        vucrlist.ListItems(tt%).Selected = False
    Next tt%

    For t% = 0 To 4
        If rs("vucr1" + Mid$(Str$(t% + 1), 2)) > "" Then
            'For tt% = 1 To vucrlist.ListItems.Count
            '    If rs("vucr1" + Mid$(Str$(t% + 1), 2)) = Mid$(vucrlist.ListItems(tt%), InStr(vucrlist.ListItems(tt%), "(") + 1, 3) Then
                    VUCRSEL(t% + 1) = rs("vucr1" + Mid$(Str$(t% + 1), 2))
                    'vucrlist.ListItems(tt%).Selected = True
                    'vucrlist.ListItems(tt%).EnsureVisible
                   ' tt% = vucrlist.ListItems.Count
            '    End If
            'Next tt%
        End If
    Next t%
    fromfind = 0
    Call Command1_Click
    vucrf.Visible = False
    fromfind = 1
    For t% = 3 To 9
        relationship(t%).ListIndex = -1
        For tt% = 0 To relationship(t%).ListCount - 1
            relationship(t%).Selected(tt%) = False
        Next tt%
        If rs("Crelationship" + Mid$(Str$(t% + 1), 2)) > "" Then
            For tt% = 0 To relationship(t%).ListCount - 1
                If rs("Crelationship" + Mid$(Str$(t% + 1), 2)) = Mid$(relationship(t%).List(tt%), InStr(relationship(t%).List(tt%), "(") + 1, 2) Then
                    relationship(t%).ListIndex = tt%
                    relationship(t%).Selected(tt%) = True
                    tt% = relationship(t%).ListCount - 1
                End If
            Next tt%
        End If
    Next t%
    For t% = 13 To 19
        For tt% = 0 To relationship(t%).ListCount - 1
            relationship(t%).Selected(tt%) = False
        Next tt%
        relationship(t%).ListIndex = -1
        If rs("Vrelationship" + Mid$(Str$(t% - 9), 2)) > "" Then
            For tt% = 0 To relationship(t%).ListCount - 1
                If rs("Vrelationship" + Mid$(Str$(t% - 9), 2)) = Mid$(relationship(t%).List(tt%), InStr(relationship(t%).List(tt%), "(") + 1, 2) Then
                    relationship(t%).ListIndex = tt%
                    relationship(t%).Selected(tt%) = True
                    tt% = relationship(t%).ListCount - 1
                End If
            Next tt%
        End If
    Next t%
    For t% = 3 To 4
        If Not IsNull(rs("offense" + Mid$(Str$(t% + 1), 2))) Then
            'RLB Code
            If Not AutoSelectOffense(t%, rs("offense" + Mid$(Str$(t% + 1), 2))) Then
                pickoffense(t%).ListIndex = -1
            End If
            '********
            completedy(t%) = rs("completedyes" + Mid$(Str$(t% + 1), 2))
            FORCEDENTRYY(t%) = rs("forcedentryyes" + Mid$(Str$(t% + 1), 2))
            completedn(t%) = rs("completedno" + Mid$(Str$(t% + 1), 2))
            FORCEDENTRYN(t%) = rs("forcedentryno" + Mid$(Str$(t% + 1), 2))
            For vv% = 1 To premise(t%).ListItems.Count
                premise(t%).ListItems(vv%).Selected = False
            Next vv%
            If rs("premise" + Mid$(Str$(t% + 1), 2)) > "" Then
                For tt% = 1 To premise(t%).ListItems.Count
                    If rs("premise" + Mid$(Str$(t% + 1), 2)) = Mid$(premise(t%).ListItems(tt%), InStr(premise(t%).ListItems(tt%), "(") + 1, 2) Then
                        premise(t%).ListItems(tt%).Selected = True
                        premise(t%).ListItems(tt%).EnsureVisible
                    End If
                Next tt%
            End If
            If rs("premise" + Mid$(Str$(t% + 1), 2) + "a") > "" Then
                For tt% = 1 To premise(t%).ListItems.Count
                    If rs("premise" + Mid$(Str$(t% + 1), 2) + "a") = Mid$(premise(t%).ListItems(tt%), InStr(premise(t%).ListItems(tt%), "(") + 1, 2) Then
                        premise(t%).ListItems(tt%).Selected = True
                        premise(t%).ListItems(tt%).EnsureVisible
                    End If
                Next tt%
            End If
            entered(t%) = rs("entered" + Mid$(Str$(t% + 1), 2))
        End If
    Next t%
End If
lactivity.ListIndex = -1
If Not IsNull(rs("lactivity")) Then
    For t% = 0 To lactivity.ListCount - 1
        If Mid$(lactivity.List(t%), InStr(lactivity.List(t%), "(") + 1, 1) = rs("lactivity") Then
            lactivity.ListIndex = t%
            t% = lactivity.ListCount - 1
        End If
    Next t%
End If
dtoffense = (incidentdate(0)) + " " + TIMEOFOFFENSE(0)
incidentnumber = rs("incidentnumber")
If UCase(vsname(2)) <> "UNKNOWN" Then
    Set db = OpenDatabase(nwl + "lawsuite.mdb")
    ssql = ""
    If IsDate(BIRTHDATE) Then
        ssql = ssql + " and birthdate = #" + BIRTHDATE + "#"
    End If
    Set rs = db.OpenRecordset("select mugshot from people where dpnamelf = " + Chr$(34) + vsname(2) + Chr$(34) + ssql + " and not mugshot is null")
    If Not rs.EOF Then
        rs.MoveFirst
        mugshot.Picture = LoadPicture(rs("mugshot"))
    End If
End If
        
schanged = 0
fromfind = 0
On Error Resume Next
If op = 1 Then
'---- setfocus logic ----
'             onpaper.SetFocus
          If onpaper.Visible Then
              onpaper.SetFocus
           End If
End If
On Error GoTo 0
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("Select max(page) as ctpg from supplemental where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
If rs.EOF Then
    pgof = "Page 1 of 1"
Else
    If Not IsNull(rs("ctpg")) Then
        pgof = "Page 1 of " + CStr(1 + rs("ctpg"))
    Else
        pgof = "Page 1 of 1"
    End If
End If
    
db.Close
On Error Resume Next
If op = 1 Then
'---- setfocus logic ----
'             onpaper.SetFocus
          If onpaper.Visible Then
              onpaper.SetFocus
           End If
End If
On Error GoTo 0
'RLB Code
    RelatedOfficers(0) = reportingofficer(0)
    RelatedOfficers(1) = reportingofficer(1)
    RelatedOfficers(2) = approvingofficer
    RelatedOfficers(3) = followupofficer
'********
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub
Private Sub ucrselect(Index As Integer)
Dim db As Database, rs As Recordset
FOUNDSELECT = 0
For t% = 0 To UCRLIST(Index).ListCount - 1
    If UCRLIST(Index).Selected(t%) Then
        FOUNDSELECT = 1
        t% = UCRLIST(Index).ListCount
    End If
Next t%
If FOUNDSELECT = 1 Then
    GoTo SKIPIT
End If
On Error GoTo oderror
od:
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("Select * from offense where offense = " + Chr$(34) + pickoffense(Index).List(pickoffense(Index).ListIndex) + Chr$(34))
If rs.EOF Then
    ts = ""
    tsct% = 0
    ts1% = 0
    ts2% = 0
    ts3% = 0
    ts4% = 0
    ts5% = 0
    For t% = 1 To Len(pickoffense(Index).List(pickoffense(Index).ListIndex))
        If Mid$(pickoffense(Index).List(pickoffense(Index).ListIndex), t%, 1) = " " Then
            tsct% = tsct% + 1
            If tsct% < 6 Then
                Select Case tsct%
                    Case 1
                        ts1% = t%
                    Case 2
                        ts2% = t%
                    Case 3
                        ts3% = t%
                    Case 4
                        ts4% = t%
                    Case 5
                        ts5% = t%
                End Select
            Else
                t% = Len(pickoffense(Index).List(pickoffense(Index).ListIndex))
            End If
        End If
    Next t%
    If ts1% > 0 Then
        ts = "offense like '*" + Mid$(pickoffense(Index).List(pickoffense(Index).ListIndex), 1, ts1% - 1) + "*'"
    Else
        ts = "offense like '*" + pickoffense(Index).List(pickoffense(Index).ListIndex) + "*'"
    End If
    If ts2% > 0 Then
        ts = ts + "and offense like '*" + Mid$(pickoffense(Index).List(pickoffense(Index).ListIndex), ts1% + 1, ts2% - ts1% - 1) + "*'"
    Else
        ts = ts + "and offense like '*" + Mid$(pickoffense(Index).List(pickoffense(Index).ListIndex), ts1% + 1) + "*'"
        GoTo tsdone
    End If
    If ts3% > 0 Then
        ts = ts + "and offense like '*" + Mid$(pickoffense(Index).List(pickoffense(Index).ListIndex), ts2% + 1, ts3% - ts2% - 1) + "*'"
    Else
        ts = ts + "and offense like '*" + Mid$(pickoffense(Index).List(pickoffense(Index).ListIndex), ts2% + 1) + "*'"
        GoTo tsdone
    End If
    If ts4% > 0 Then
        ts = ts + "and offense like '*" + Mid$(pickoffense(Index).List(pickoffense(Index).ListIndex), ts3% + 1, ts4% - ts3% - 1) + "*'"
    Else
        ts = ts + "and offense like '*" + Mid$(pickoffense(Index).List(pickoffense(Index).ListIndex), ts3% + 1) + "*'"
        GoTo tsdone
    End If
    If ts5% > 0 Then
        ts = ts + "and offense like '*" + Mid$(pickoffense(Index).List(pickoffense(Index).ListIndex), ts4% + 1, ts5% - ts4% - 1) + "*'"
    Else
        ts = ts + "and offense like '*" + Mid$(pickoffense(Index).List(pickoffense(Index).ListIndex), ts4% + 1) + "*'"
    End If
tsdone:
    Set rs = db.OpenRecordset("Select * from offense where " + ts)
    If Not rs.EOF Then
        rs.MoveFirst
        If rs.EOF Then
            For t% = 0 To UCRLIST(Index).ListCount - 1
                If UCRLIST(Index).List(t%) = rs("ucr") Then
                    UCRLIST(Index).ListIndex = t%
                    UCRLIST(Index).Selected(t%) = True
                    t% = UCRLIST(Index).ListCount
                End If
            Next t%
        End If
    End If
Else
    rs.MoveFirst
    If Not rs.EOF Then
        For t% = 0 To UCRLIST(Index).ListCount - 1
            If UCRLIST(Index).List(t%) = rs("ucr") Then
                UCRLIST(Index).Selected(t%) = True
                UCRLIST(Index).ListIndex = t%
                t% = UCRLIST(Index).ListCount
            End If
        Next t%
    End If
End If
db.Close
SKIPIT:
On Error Resume Next
If fromfind = 1 Then
    Exit Sub
End If

UCRLIST(Index).Height = 2000
UCRLIST(Index).Left = 4500
UCRLIST(Index).Top = pickoffense(Index).Top - 1000
UCRLIST(Index).Visible = True
'---- setfocus logic ----
'         UCRLIST(index).SetFocus
          If UCRLIST(Index).Visible Then
              UCRLIST(Index).SetFocus
           End If
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub
Friend Sub saveincident()
If frmLogin.IEDIT <> 1 And frmLogin.ISUPERVISOR <> 1 And frmLogin.SUPERVISOR <> 1 Then
    msg = MsgBox("Insufficient authority for this operation.", 48, "Genesis Error Log")
    Exit Sub
End If
If vucrlist.ListItems.Count = 0 Or vucrlist.SelectedItem Is Nothing Then
    msg = MsgBox("The victim must be associated with UCR's for this report.", 48, "Genesis Error Log")
    Call Command1_Click
    Exit Sub
End If
incident.Refresh
Dim db As Database, rs, rs2 As Recordset, lu As String, luu As String
luu = ""
On Error GoTo oderror1

    'RLB Code
    If IncidentRptLoadedFromDB Then
        If UCase(reportingofficer(0).Text) <> UCase(frmLogin.userfullname) Then
            If UCase(reportingofficer(1).Text) <> UCase(frmLogin.userfullname) Then
                If UCase(approvingofficer.Text) <> UCase(frmLogin.userfullname) Then
                     If UCase(followupofficer.Text) <> UCase(frmLogin.userfullname) Then
                        If (frmLogin.ISUPERVISOR <> 1) And (frmLogin.SUPERVISOR <> 1) Then
                            MsgBox "Since you are not recognized as an officer related to this report, you cannot alter this incident report.", vbOKOnly, "Genesis Error Log"
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    End If
    '********

od1:
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("Select * from incidentreporto where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
ecc = 0

If Not rs.EOF Then
    
       
    rs.MoveFirst
    ecc = 0
    If rs("excleardate") <> EXCEPTIONALCLEARANCEDATE Then
        ecc = 1
    End If
    If (rs("exclearunder18") = "X" And exclearunder18 = 0) Or (rs("exclearunder18") <> "X" And exclearunder18 = 1) Then
        ecc = 1
    End If
    If (rs("exclear18over") = "X" And exclear18andover = 0) Or (rs("exclear18over") <> "X" And exclear18andover = 1) Then
        ecc = 1
    End If
    rs.Edit
    schanged = 0
Else
    rs.AddNew
    schanged = 1
End If
schanged = 1
rs("incidentnumber") = incidentnumber
rs("narrative") = NARRATIVE.Text
rs("JURISDICTIONTHEFT") = JURISDICTIONTHEFT
rs("JURISDICTIONRECOVERY") = JURISDICTIONRECOVERY
For t% = 0 To 41
    tt% = t% Mod 6
    If description(tt%) > "" Or Val(totalvalue(t%)) > 0 Or totalvalue(t%) = "X" Or group(tt%).ListIndex > -1 Then
        rs("type" + Mid$(Str$(tt% + 1), 2)) = description(tt%)
        rs("major" + Mid$(Str$(tt% + 1), 2)) = majorlist(tt%).List(majorlist(tt%).ListIndex)
        rs("minor" + Mid$(Str$(tt% + 1), 2)) = minorlist(tt%).List(minorlist(tt%).ListIndex)
        If totalvalue(t%) = "X" Then
            Select Case t%
                Case 0 To 5
                    rs("stolenvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = 9999999
                Case 6 To 11
                    rs("damagedvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = 9999999
                Case 12 To 17
                    rs("burnedvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = 9999999
                Case 18 To 23
                    rs("recoveredvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = 9999999
                Case 24 To 29
                    rs("seizedvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = 9999999
                Case 30 To 35
                    rs("COUNTERFEITvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = 9999999
                Case 36 To 41
                    rs("UNKNOWNvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = 9999999
            End Select
        Else
            Select Case t%
                Case 0 To 5
                    rs("stolenvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = Val(totalvalue(t%))
                Case 6 To 11
                    rs("damagedvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = Val(totalvalue(t%))
                Case 12 To 17
                    rs("burnedvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = Val(totalvalue(t%))
                Case 18 To 23
                    rs("recoveredvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = Val(totalvalue(t%))
                Case 24 To 29
                    rs("seizedvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = Val(totalvalue(t%))
                Case 30 To 35
                    rs("COUNTERFEITvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = Val(totalvalue(t%))
                Case 36 To 41
                    rs("UNKNOWNvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = Val(totalvalue(t%))
            End Select
        End If
    Else
        rs("type" + Mid$(Str$(tt% + 1), 2)) = Null
        rs("major" + Mid$(Str$(tt% + 1), 2)) = Null
        rs("minor" + Mid$(Str$(tt% + 1), 2)) = Null
        Select Case t%
            Case 0 To 5
                rs("stolenvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = Null
            Case 6 To 11
                rs("damagedvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = Null
            Case 12 To 17
                rs("burnedvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = Null
            Case 18 To 23
                rs("recoveredvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = Null
            Case 24 To 29
                rs("seizedvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = Null
            Case 30 To 35
                rs("COUNTERFEITvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = Null
            Case 36 To 41
                rs("UNKNOWNvalue" + Mid$(Str$((t% Mod 6) + 1), 2)) = Null
        End Select
    End If
Next t%
ct% = 0
For t% = 6 To 23
    ct% = ct% + 1
    For tt% = 1 To 3
        st% = tt%
        If drugtype(t%).ListIndex > -1 Then
            rs("ptypeofdrug" + Mid$(Str$(ct%), 2) + Mid$(Str$(tt%), 2)) = Left$(drugtype(t%).List(drugtype(t%).ListIndex), 1)
            rs("pdrugamt" + Mid$(Str$(ct%), 2) + Mid$(Str$(tt%), 2)) = Val(drugamt(t%))
            rs("pdrugmeasurement" + Mid$(Str$(ct%), 2) + Mid$(Str$(tt%), 2)) = Left$(drugmeasurement(t%).List(drugmeasurement(t%).ListIndex), 2)
        Else
            rs("ptypeofdrug" + Mid$(Str$(ct%), 2) + Mid$(Str$(tt%), 2)) = Nrull
            rs("pdrugamt" + Mid$(Str$(ct%), 2) + Mid$(Str$(tt%), 2)) = Null
            rs("pdrugmeasurement" + Mid$(Str$(ct%), 2) + Mid$(Str$(tt%), 2)) = Null
        End If
        t% = t% + 1
    Next tt%
    t% = t% - 1
Next t%
rs("subjectidentifiedyes") = subjectidentifiedyes
rs("subjectidentifiedno") = subjectidentifiedno
rs("subjectlocatedyes") = subjectlocatedyes
rs("subjectlocatedno") = subjectlocatedno
If active = 1 And (IsNull(rs("ACTIVE")) Or rs("ACTIVE") <> "X") Then
    rs("active") = "X"
    rs("statuschange") = incidentdate(1)
End If
If admclosed = 1 And (IsNull(rs("ADMCLOSED")) Or rs("ADMCLOSED") <> "X") Then
    rs("admclosed") = "X"
    rs("statuschange") = Date$
End If
If unfounded = 1 And (IsNull(rs("UNFOUNDED")) Or rs("UNFOUNDED") <> "X") Then
    rs("unfounded") = "X"
    rs("statuschange") = Date$
End If
If arrestedunder18 = 1 Then
    rs("arrestedunder18") = "X"
End If
If arrested18andover = 1 Then
    rs("arrested18over") = "X"
End If
If exclearunder18 = 1 Then
    rs("exclearunder18") = "X"
End If
If exclear18andover = 1 Then
    rs("exclear18over") = "X"
End If
rs("offenderdeath") = offenderdeath
rs("noprosecution") = noprosecution
rs("extraditiondenied") = extraditiondenied
rs("victimdeclinescooperation") = victimdeclinescooperation
rs("juvenilenocustody") = juvenilenocustody
If IsDate(EXCEPTIONALCLEARANCEDATE) Then
    rs("excleardate") = EXCEPTIONALCLEARANCEDATE
Else
    rs("excleardate") = Null
End If
rs("reportingofficer1") = reportingofficer(0)
If IsDate(REPORTINGOFFICERDATE(0)) Then
    rs("reportingdate1") = REPORTINGOFFICERDATE(0)
End If
rs("reportingunit1") = reportingofficeRunit(0)
If reportingofficer(1) > "" Then
    rs("reportingofficer2") = reportingofficer(1)
    If IsDate(REPORTINGOFFICERDATE(1)) Then
        rs("reportingdate2") = REPORTINGOFFICERDATE(1)
    End If
    rs("reportingunit2") = reportingofficeRunit(1)
End If
rs("followupyes") = followupyes
rs("followupno") = followupno
If followupofficer > "" Then
    rs("followupofficer") = followupofficer
End If
rs("approvingofficer") = approvingofficer
If IsDate(APPROVINGOFFICERDATE) Then
    rs("approvingdate") = APPROVINGOFFICERDATE
End If
rs("approvingunit") = approvingofficeRunit
If IsDate(FOLLOWUPOFFICERDATE) Then
    rs("followupdate") = FOLLOWUPOFFICERDATE
End If
rs("followupunit") = FOLLOWUPOFFICERUNIT
If BIAS > "" Then
    rs("bias") = Mid$(BIAS.List(BIAS.ListIndex), InStr(BIAS.List(BIAS.ListIndex), "(") + 1, 2)
End If
rs.Update

Set rs = db.OpenRecordset("Select * from incidentreports where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
ecc = 0
If Not rs.EOF Then
    rs.MoveFirst
    rs.Edit
Else
    rs.AddNew
End If
schanged = 1
rs("incidentnumber") = incidentnumber
If IsDate(BIRTHDATE) Then
    rs("sbirthdate") = BIRTHDATE
Else
    rs("sbirthdate") = Null
End If
rs("arrestnumber") = incidentnumber
rs("sname") = vsname(2)
rs("saddress") = address(2)
rs("scity") = city(2)
rs("sstate") = state(2)
rs("szipcode") = zipcode(2)
rs("SHEIGHT") = ht(1)
rs("SWEIGHT") = weight(1)
rs("SHAIR") = hair(1)
rs("SEYES") = eyes(1)
rs("SPECULIARITIES") = peculiarities(1)
rs("Srace") = Left$(race(2).List(race(2).ListIndex), 1)
rs("Ssex") = Left$(sex(2).List(sex(2).ListIndex), 1)
rs("Sage") = age(2)
rs("Sethnicity") = Left$(ethnicity(2).List(ethnicity(2).ListIndex), 1)
rs("Slocationnumber") = LOCATIONNUMBER(2)
rs("computerequipment") = computerequipment(1)
rs("Salcoholyes") = ""
rs("Salcoholno") = ""
rs("Salcoholunknown") = ""
If alcoholyes(1) Then
    rs("Salcoholyes") = "X"
Else
If alcoholno(1) Then
    rs("Salcoholno") = "X"
Else
If alcoholunknown(1) Then
    rs("Salcoholunknown") = "X"
End If
End If
End If
rs("Sdrugsyes") = ""
rs("Sdrugsno") = ""
rs("Sdrugsunknown") = ""
If drugsyes(1) Then
    rs("Sdrugsyes") = "X"
Else
If drugsno(1) Then
    rs("Sdrugsno") = "X"
Else
If drugsunknown(1) Then
    rs("Sdrugsunknown") = "X"
End If
End If
End If
If drugtype(3).ListIndex > -1 Then
    rs("stypeofdrug") = Left$(drugtype(3).List(drugtype(3).ListIndex), 1)
    rs("SDRUGAMT") = drugamt(3)
    rs("SDRUGMEASUREMENT") = Left$(drugmeasurement(3).List(drugmeasurement(3).ListIndex), 2)
Else
    rs("stypeofdrug") = Null
    rs("SDRUGAMT") = Null
    rs("SDRUGMEASUREMENT") = Null
End If
If SUSPECT = 1 Then
    rs("ssuspect") = "X"
End If
If RUNAWAY = 1 Then
    rs("srunaway") = "X"
End If
If WANTED = 1 Then
    rs("swanted") = "X"
End If
If ARREST = 1 Then
    rs("sarrest") = "X"
End If
If WARRANT = 1 Then
    rs("swarrant") = "X"
End If
If JAIL = 1 Then
    rs("sjail") = "X"
End If
If SUMMONS = 1 Then
    rs("ssummons") = "X"
End If
If WARRANT = 1 Then
    rs("swarrant") = "X"
End If
If ARRESTEDNEARYES Then
    rs("sarrestednearoffenseyes") = "X"
End If
If ARRESTEDNEARNO Then
    rs("sarrestednearoffenseno") = "X"
End If
rs("totalarrested") = totalnumberarrested
If IsDate(DATEOFARREST) Then
    rs("dateofarrest") = DATEOFARREST
End If
rs("timeofarrest") = TIMEOFARREST
rs.Update

Set rs = db.OpenRecordset("Select * from incidentreportc where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    rs.Edit
Else
    rs.AddNew
End If
schanged = 1
rs("incidentnumber") = incidentnumber
rs("timeofoffense1") = Format$(TIMEOFOFFENSE(0), "hh:mm")
rs("dateofoffense1") = incidentdate(0)
rs("timeofoffense2") = Format$(TIMEOFOFFENSE(1), "hh:mm")
If IsDate(incidentdate(1)) Then
    rs("dateofoffense2") = incidentdate(1)
End If
For t% = 0 To 2
    UCRLIST(t%).ListIndex = -1
    For tt% = 0 To UCRLIST(t%).ListCount - 1
        If UCRLIST(t%).Selected(tt%) = True Then
            UCRLIST(t%).ListIndex = tt%
            tt% = UCRLIST(t%).ListCount - 1
        End If
    Next tt%
    If UCRLIST(t%).ListIndex > -1 Then
        rs("offense" + Mid$(Str$(t% + 1), 2)) = pickoffense(t%).List(pickoffense(t%).ListIndex)
        rs("completedyes" + Mid$(Str$(t% + 1), 2)) = completedy(t%)
        rs("forcedentryyes" + Mid$(Str$(t% + 1), 2)) = FORCEDENTRYY(t%)
        rs("completedno" + Mid$(Str$(t% + 1), 2)) = completedn(t%)
        rs("forcedentryno" + Mid$(Str$(t% + 1), 2)) = FORCEDENTRYN(t%)
        If premise(t%).SelectedItem.Index > 0 Then
            firstone% = 0
            For ZZ% = 1 To premise(t%).ListItems.Count
                If premise(t%).ListItems(ZZ%).Selected Then
                    firstone% = ZZ%
                    rs("premise" + Mid$(Str$(t% + 1), 2)) = Mid$(premise(t%).ListItems(ZZ%), InStr(premise(t%).ListItems(ZZ%), "(") + 1, 2)
                    ZZ% = premise(t%).ListItems.Count
                End If
            Next ZZ%
            If firstone% > 0 Then
                For ZZ% = firstone% + 1 To premise(t%).ListItems.Count
                    If premise(t%).ListItems(ZZ%).Selected Then
                        rs("premise" + Mid$(Str$(t% + 1), 2) + "a") = Mid$(premise(t%).ListItems(ZZ%), InStr(premise(t%).ListItems(ZZ%), "(") + 1, 2)
                        ZZ% = premise(t%).ListItems.Count
                    End If
                Next ZZ%
            End If
        End If
        rs("entered" + Mid$(Str$(t% + 1), 2)) = Val(entered(t%))
    Else
        rs("offense" + Mid$(Str$(t% + 1), 2)) = Null
        rs("completedyes" + Mid$(Str$(t% + 1), 2)) = Null
        rs("forcedentryyes" + Mid$(Str$(t% + 1), 2)) = Null
        rs("completedno" + Mid$(Str$(t% + 1), 2)) = Null
        rs("forcedentryno" + Mid$(Str$(t% + 1), 2)) = Null
        rs("premise" + Mid$(Str$(t% + 1), 2)) = Null
        rs("entered" + Mid$(Str$(t% + 1), 2)) = Null
    End If
Next t%
rs("individual") = individual
rs("business") = business
rs("financialinstitution") = financialinstitution
rs("other") = other
rs("government") = government
rs("unknown") = unknown
rs("religiousorganization") = religiousorganization
rs("policeofficer") = policeofficer
rs("societypublic") = societypublic
rs("incidentlocation") = incidentlocation
rs("clocationnumber") = LOCATIONNUMBER(0)
rs("incidentzipcode") = incidentzipcode
For i% = 1 To 3
    rs("weapontype" + Mid$(Str$(i%), 2)) = ""
    rs("automatic" + Mid$(Str$(i%), 2)) = ""
Next i%
idx5 = 0
For i% = 1 To weapontype.ListItems.Count
    If weapontype.ListItems(i%).Selected = True Then
        idx5 = idx5 + 1
        If idx5 < 4 Then
            rs("weapontype" + Mid$(Str$(idx5), 2)) = Mid$(weapontype.ListItems(i%), InStr(weapontype.ListItems(i%), "(") + 1, 2)
            rs("automatic" + Mid$(Str$(idx5), 2)) = automatic(i%)
        Else
            i% = weapontype.ListItems.Count
        End If
    End If
Next i%
If IsDate(dispatchdate) Then
    rs("dispatchdate") = dispatchdate
End If
If IsDate(DISPATCHTIME) Then
    rs("dispatchtime") = DISPATCHTIME
End If
If IsDate(TIMEARRIVED) Then
    rs("arrivaltime") = TIMEARRIVED
End If
If IsDate(DEPARTINGTIME) Then
    rs("departuretime") = DEPARTINGTIME
End If
rs("incidentlocatioNnumber") = ELOCATIONNUMBER
rs("cname") = vsname(0)
rs("caddress") = address(0)
rs("ccity") = city(0)
rs("cstate") = state(0)
rs("czipcode") = zipcode(0)
rs("cresident") = Left$(resident(0).List(resident(0).ListIndex), 1)
rs("crace") = Left$(race(0).List(race(0).ListIndex), 1)
rs("csex") = Left$(sex(0).List(sex(0).ListIndex), 1)
rs("cage") = age(0)
rs("cethnicity") = Left$(ethnicity(0).List(ethnicity(0).ListIndex), 1)
For t% = 0 To 2
    If relationship(t%).ListIndex > -1 Then
        rs("crelationship" + Mid$(Str$(t% + 1), 2)) = Mid$(relationship(t%).List(relationship(t%).ListIndex), InStr(relationship(t%).List(relationship(t%).ListIndex), "(") + 1, 2)
    Else
        rs("crelationship" + Mid$(Str$(t% + 1), 2)) = Null
    End If
Next t%
rs("cdayhomephone") = HOMEDAYPHONE(0)
rs("cdayworkphone") = WORKDAYPHONE(0)
rs("cnighthomephone") = HOMENIGHTPHONE(0)
rs("cnightworkphone") = WORKNIGHTPHONE(0)

rs.Update

Set rs = db.OpenRecordset("Select * from incidentreportv where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    rs.Edit
Else
    rs.AddNew
End If
schanged = 1
rs("incidentnumber") = incidentnumber
rs("vname") = vsname(1)
rs("vaddress") = address(1)
rs("vcity") = city(1)
rs("computerequipment") = computerequipment(0)
rs("vstate") = state(1)
rs("vzipcode") = zipcode(1)
For t% = 0 To 2
    If relationship(t% + 10).ListIndex > -1 Then
        rs("Vrelationship" + Mid$(Str$(t% + 1), 2)) = Mid$(relationship(t% + 10).List(relationship(t% + 10).ListIndex), InStr(relationship(t% + 10).List(relationship(t% + 10).ListIndex), "(") + 1, 2)
    Else
        rs("Vrelationship" + Mid$(Str$(t% + 1), 2)) = Null
    End If
Next t%
rs("VHEIGHT") = ht(0)
rs("VWEIGHT") = weight(0)
rs("VHAIR") = hair(0)
rs("VEYES") = eyes(0)
rs("VPECULIARITIES") = peculiarities(0)
rs("vresident") = Left$(resident(1).List(resident(1).ListIndex), 1)
rs("vrace") = Left$(race(1).List(race(1).ListIndex), 1)
rs("vsex") = Left$(sex(1).List(sex(1).ListIndex), 1)
rs("vage") = age(1)
rs("vethnicity") = Left$(ethnicity(1).List(ethnicity(1).ListIndex), 1)
rs("vlocationnumber") = LOCATIONNUMBER(1)
rs("vhomephoneDAY") = HOMEDAYPHONE(1)
rs("vworkphoneDAY") = WORKDAYPHONE(1)
rs("vhomephoneNIGHT") = HOMENIGHTPHONE(1)
rs("vworkphoneNIGHT") = WORKNIGHTPHONE(1)
IIDX% = 0
For t% = 1 To 5
    rs("typeofinjury" + Mid$(Str$(t%), 2)) = ""
Next t%
For t% = 1 To injury.ListItems.Count
    If injury.ListItems(t%).Selected = True Then
        IIDX% = IIDX% + 1
        If IIDX% < 6 Then
            rs("typeofinjury" + Mid$(Str$(IIDX%), 2)) = Mid$(injury.ListItems(t%), InStr(injury.ListItems(t%), "(") + 1, 1)
        Else
            t% = injury.ListItems.Count
        End If
    End If
Next t%
If VISIBLEINJURYYES Then
    rs("vvisibleinjuryyes") = "X"
Else
    rs("vvisibleinjuryno") = "X"
End If
If NONVISIBLEINJURYYES Then
    rs("vnonvisibleinjuryyes") = "X"
Else
    rs("vnonvisibleinjuryno") = "X"
End If
rs("valcoholyes") = ""
rs("valcoholno") = ""
rs("valcoholunknown") = ""
If alcoholyes(0) Then
    rs("valcoholyes") = "X"
Else
If alcoholno(0) Then
    rs("valcoholno") = "X"
Else
If alcoholunknown(0) Then
    rs("valcoholunknown") = "X"
End If
End If
End If
rs("vdrugsyes") = ""
rs("vdrugsno") = ""
rs("vdrugsunknown") = ""
If drugsyes(0) Then
    rs("vdrugsyes") = "X"
Else
If drugsno(0) Then
    rs("vdrugsno") = "X"
Else
If drugsunknown(0) Then
    rs("vdrugsunknown") = "X"
End If
End If
End If
If TWOMANVEHICLE = 1 Then
    rs("vtwomanvehicle") = "X"
End If
If ONEMANVEHICLE = 1 Then
    rs("vonemanvehicle") = "X"
End If
For t% = 10 To 12
    If relationship(t%).ListIndex > -1 Then
        rs("Vrelationship" + Mid$(Str$(t% - 9), 2)) = Mid$(relationship(t%).List(relationship(t%).ListIndex), InStr(relationship(t%).List(relationship(t%).ListIndex), "(") + 1, 2)
    Else
        rs("Vrelationship" + Mid$(Str$(t% - 9), 2)) = Null
    End If
Next t%
If drugtype(0).ListIndex > -1 Then
    rs("vtypeofdrug") = Left$(drugtype(0).List(drugtype(0).ListIndex), 1)
    rs("VDRUGAMT") = drugamt(0)
    rs("VDRUGMEASUREMENT") = Left$(drugmeasurement(0).List(drugmeasurement(0).ListIndex), 2)
Else
    rs("vtypeofdrug") = Null
    rs("VDRUGAMT") = Null
    rs("VDRUGMEASUREMENT") = Null
End If
If DETECTIVE = 1 Then
    rs("vdetective") = "X"
End If
If TODOTHER = 1 Then
    rs("vother") = "X"
End If
If ALONE = 1 Then
    rs("valone") = "X"
End If
If ASSISTED = 1 Then
    rs("vassisted") = "X"
End If
rs.Update




'----- support
Set rs = db.OpenRecordset("Select * from incidentSUPPORT where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    od = rs("original")
    If rs("schanged") = 1 Then
        schanged = 1
    End If
    rs.Edit
Else
    lu = Date$
    od = Date$
    rs.AddNew
End If
schanged = 1
rs("original") = od
rs("LASTUPDATE") = Date$
rs("schanged") = schanged
rs("exclearchange") = ecc
rs("incidentnumber") = incidentnumber
If tempsave = 1 Then
    rs("temp") = "Y"
    rs("tempreason") = "TEMP SAVE"
Else
    rs("temp") = "N"
    rs("tempreason") = ""
End If
rs("onpaper") = onpaper
rs("local") = locali
If lactivity.ListIndex > -1 Then
    rs("lactivity") = Mid$(lactivity.List(lactivity.ListIndex), InStr(lactivity.List(lactivity.ListIndex), "(") + 1, 1)
Else
    rs("lactivity") = Null
End If
For t% = 3 To 9
    If relationship(t%).ListIndex > -1 Then
        rs("crelationship" + Mid$(Str$(t% + 1), 2)) = Mid$(relationship(t%).List(relationship(t%).ListIndex), InStr(relationship(t%).List(relationship(t%).ListIndex), "(") + 1, 2)
    Else
        rs("crelationship" + Mid$(Str$(t% + 1), 2)) = Null
    End If
Next t%
For t% = 13 To 19
    If relationship(t%).ListIndex > -1 Then
        rs("Vrelationship" + Mid$(Str$(t% - 9), 2)) = Mid$(relationship(t%).List(relationship(t%).ListIndex), InStr(relationship(t%).List(relationship(t%).ListIndex), "(") + 1, 2)
    Else
        rs("Vrelationship" + Mid$(Str$(t% - 9), 2)) = Null
    End If
Next t%
For t% = 0 To 5
    If numvehicle(t%) > "" Then
        rs("numvehicles" + Mid$(Str$(t% + 1), 2)) = Val(numvehicle(t%))
    Else
        rs("numvehicles" + Mid$(Str$(t% + 1), 2)) = Null
    End If
    If IsDate(DATERECOVERED(t%)) Then
        rs("daterecovered" + Mid$(Str$(t% + 1), 2)) = DATERECOVERED(t%)
    Else
        rs("daterecovered" + Mid$(Str$(t% + 1), 2)) = Null
    End If
Next t%
For t% = 6 To 11
    If numvehicle(t%) > "" Then
        rs("numvehicler" + Mid$(Str$(t% - 5), 2)) = Val(numvehicle(t%))
    Else
        rs("numvehicler" + Mid$(Str$(t% - 5), 2)) = Null
    End If
Next t%

For tt% = 1 To 10
    rs("additional" + Mid$(Str$(tt%), 2)) = Null
    rs("ucr" + Mid$(Str$(tt%), 2)) = Null
    For t% = 1 To 2
        rs("homocide" + Mid$(Str$(t%), 2) + Mid$(Str$(tt%), 2)) = Null
        rs("activity" + Mid$(Str$(tt%), 2) + Mid$(Str$(t%), 2)) = Null
        rs("subcodes" + Mid$(Str$(tt%), 2) + Mid$(Str$(t%), 2)) = Null
    Next t%
    rs("activity" + Mid$(Str$(tt%), 2) + "3") = Null
    rs("subcodes" + Mid$(Str$(tt%), 2) + "3") = Null
Next tt%



For t% = 0 To 4
    hct% = 0
    For xx% = 1 To HOMOCIDE(t%).ListItems.Count
        If HOMOCIDE(t%).ListItems(xx%).Selected Then
            nohomo% = 1
            hct% = hct% + 1
            rs("homocide" + Mid$(Str$(hct%), 2) + Mid$(Str$(t% + 1), 2)) = Mid$(HOMOCIDE(t%).ListItems(xx%), InStr(HOMOCIDE(t%).ListItems(xx%), "(") + 1, 2)
        End If
        If hct% = 2 Then
            xx% = HOMOCIDE(t%).ListItems.Count
        End If
    Next xx%
    If additional(t%).ListIndex > -1 Then
        rs("additional" + Mid$(Str$(t% + 1), 2)) = Mid$(additional(t%).List(additional(t%).ListIndex), InStr(additional(t%).List(additional(t%).ListIndex), "(") + 1, 1)
    End If
    UCRLIST(t%).ListIndex = -1
    For tt% = 0 To UCRLIST(t%).ListCount - 1
        If UCRLIST(t%).Selected(tt%) = True Then
            UCRLIST(t%).ListIndex = tt%
            tt% = UCRLIST(t%).ListCount - 1
        End If
    Next tt%
    If UCRLIST(t%).ListIndex > -1 Then
        rs("ucr" + Mid$(Str$(t% + 1), 2)) = Mid$(UCRLIST(t%).List(UCRLIST(t%).ListIndex), InStr(UCRLIST(t%).List(UCRLIST(t%).ListIndex), "(") + 1, 3)
        tempucr = rs("ucr" + Mid$(Str$(t% + 1), 2))
        IDX = 1
        If tempucr = "09A" Or tempucr = "09B" Or tempucr = "100" Or tempucr = "120" Or tempucr = "11A" Or tempucr = "11B" Or tempucr = "11C" Or tempucr = "11D" Or tempucr = "13A" Or tempucr = "13B" Or tempucr = "13C" Then
            For tt% = 1 To gactivity(t%).ListItems.Count
                If gactivity(t%).ListItems(tt%).Selected Then
                    rs("activity" + Mid$(Str$(t% + 1), 2) + Mid$(Str$(IDX), 2)) = Mid$(gactivity(t%).ListItems(tt%), InStr(gactivity(t%).ListItems(tt%), "(") + 1, 1)
                    IDX = IDX + 1
                End If
                If IDX > 3 Then
                    tt% = gactivity(t%).ListItems.Count
                End If
            Next tt%
        Else
            For tt% = 1 To activity(t%).ListItems.Count
                If activity(t%).ListItems(tt%).Selected Then
                    rs("activity" + Mid$(Str$(t% + 1), 2) + Mid$(Str$(IDX), 2)) = Mid$(activity(t%).ListItems(tt%), InStr(activity(t%).ListItems(tt%), "(") + 1, 1)
                    IDX = IDX + 1
                End If
                If IDX > 3 Then
                    tt% = activity(t%).ListItems.Count
                End If
            Next tt%
        End If
    Else
        rs("ucr" + Mid$(Str$(t% + 1), 2)) = Null
    End If
    uct% = 0
    rs("subcodes" + Mid$(Str$(t% + 1), 2) + "1") = Null
    rs("subcodes" + Mid$(Str$(t% + 1), 2) + "2") = Null
    rs("subcodes" + Mid$(Str$(t% + 1), 2) + "3") = Null
    For uuu% = 1 To sublist(t%).ListItems.Count
        If sublist(t%).ListItems(uuu%).Selected Then
            uct% = uct% + 1
            If uct% < 4 Then
                rs("subcodes" + Mid$(Str$(t% + 1), 2) + Mid$(Str$(uct%), 2)) = Left$(sublist(t%).ListItems(uuu%), 1)
            Else
                uuu% = sublist(t%).ListItems.Count
            End If
        End If
    Next uuu%
Next t%
If drugtype(1).ListIndex > -1 Then
    rs("vtypeofdrug2") = Left$(drugtype(1).List(drugtype(1).ListIndex), 1)
    rs("VDRUGAMT2") = drugamt(1)
    rs("VDRUGMEASUREMENT2") = Left$(drugmeasurement(1).List(drugmeasurement(1).ListIndex), 2)
Else
    rs("vtypeofdrug2") = Null
    rs("VDRUGAMT2") = Null
    rs("VDRUGMEASUREMENT2") = Null
End If
If drugtype(2).ListIndex > -1 Then
    rs("vtypeofdrug3") = Left$(drugtype(2).List(drugtype(2).ListIndex), 1)
    rs("VDRUGAMT3") = drugamt(2)
    rs("VDRUGMEASUREMENT3") = Left$(drugmeasurement(2).List(drugmeasurement(2).ListIndex), 2)
Else
    rs("vtypeofdrug3") = Null
    rs("VDRUGAMT3") = Null
    rs("VDRUGMEASUREMENT3") = Null
End If
If drugtype(4).ListIndex > -1 Then
    rs("stypeofdrug2") = Left$(drugtype(4).List(drugtype(4).ListIndex), 1)
    rs("SDRUGAMT2") = drugamt(4)
    'rs("SDRUGAMT2") = drugamt(2)
    rs("SDRUGMEASUREMENT2") = Left$(drugmeasurement(4).List(drugmeasurement(4).ListIndex), 2)
Else
    rs("stypeofdrug2") = Null
    rs("SDRUGAMT2") = Null
    rs("SDRUGMEASUREMENT2") = Null
End If
If drugtype(5).ListIndex > -1 Then
    rs("stypeofdrug3") = Left$(drugtype(5).List(drugtype(5).ListIndex), 1)
    rs("SDRUGAMT3") = drugamt(5)
    rs("SDRUGMEASUREMENT3") = Left$(drugmeasurement(5).List(drugmeasurement(5).ListIndex), 2)
Else
    rs("stypeofdrug3") = Null
    rs("SDRUGAMT3") = Null
    rs("SDRUGMEASUREMENT3") = Null
End If
For t% = 0 To 5
    If group(t%).ListIndex > -1 Then
        rs("GROUP" + Mid$(Str$(t% + 1), 2)) = Mid$(group(t%).List(group(t%).ListIndex), InStr(group(t%).List(group(t%).ListIndex), "(") + 1, 2)
        If pucrlist(t%).ListIndex > -1 Then
            rs("pucr" + Mid$(Str$(t% + 1), 2)) = Mid$(pucrlist(t%).List(pucrlist(t%).ListIndex), InStr(pucrlist(t%).List(pucrlist(t%).ListIndex), "(") + 1, 3)
        Else
            rs("pucr" + Mid$(Str$(t% + 1), 2)) = Null
        End If
    Else
        rs("GROUP" + Mid$(Str$(t% + 1), 2)) = Null
        rs("pucr" + Mid$(Str$(t% + 1), 2)) = Null
    End If
Next t%
vucridx = 1
For t% = 1 To vucrlist.ListItems.Count
    If vucrlist.ListItems(t%).Selected = True Then
        rs("vucr1" + Mid$(Str$(vucridx), 2)) = Mid$(vucrlist.ListItems(t%), InStr(vucrlist.ListItems(t%), "(") + 1, 3)
        vucridx = vucridx + 1
    End If
    If vucridx > 5 Then
        t% = vucrlist.ListItems.Count
    End If
Next t%
If vucridx > 1 Then
    For t% = vucridx To 5
        rs("vucr1" + Mid$(Str$(t%), 2)) = Null
    Next t%
End If
If vucridx = 1 And vucrlist.ListItems.Count = 1 Then
    rs("vucr1" + Mid$(Str$(vucridx), 2)) = Mid$(vucrlist.ListItems(1), InStr(vucrlist.ListItems(1), "(") + 1, 3)
End If
For t% = 3 To 4
    If pickoffense(t%).ListIndex <> -1 Then
        rs("offense" + Mid$(Str$(t% + 1), 2)) = pickoffense(t%)
        rs("completedyes" + Mid$(Str$(t% + 1), 2)) = completedy(t%)
        rs("forcedentryyes" + Mid$(Str$(t% + 1), 2)) = FORCEDENTRYY(t%)
        rs("completedno" + Mid$(Str$(t% + 1), 2)) = completedn(t%)
        rs("forcedentryno" + Mid$(Str$(t% + 1), 2)) = FORCEDENTRYN(t%)
        If premise(t%).SelectedItem.Index > 0 Then
            firstone% = 0
            For ZZ% = 1 To premise(t%).ListItems.Count
                If premise(t%).ListItems(ZZ%).Selected Then
                    firstone% = ZZ%
                    rs("premise" + Mid$(Str$(t% + 1), 2)) = Mid$(premise(t%).ListItems(ZZ%), InStr(premise(t%).ListItems(ZZ%), "(") + 1, 2)
                    ZZ% = premise(t%).ListItems.Count
                End If
            Next ZZ%
            If firstone% > 0 Then
                For ZZ% = firstone% + 1 To premise(t%).ListItems.Count
                    If premise(t%).ListItems(ZZ%).Selected Then
                        rs("premise" + Mid$(Str$(t% + 1), 2) + "a") = Mid$(premise(t%).ListItems(ZZ%), InStr(premise(t%).ListItems(ZZ%), "(") + 1, 2)
                        ZZ% = premise(t%).ListItems.Count
                    End If
                Next ZZ%
            End If
        End If
        rs("entered" + Mid$(Str$(t% + 1), 2)) = Val(entered(t%))
    Else
        rs("offense" + Mid$(Str$(t% + 1), 2)) = Null
        rs("completedyes" + Mid$(Str$(t% + 1), 2)) = Null
        rs("forcedentryyes" + Mid$(Str$(t% + 1), 2)) = Null
        rs("completedno" + Mid$(Str$(t% + 1), 2)) = Null
        rs("forcedentryno" + Mid$(Str$(t% + 1), 2)) = Null
        rs("premise" + Mid$(Str$(t% + 1), 2)) = Null
        rs("entered" + Mid$(Str$(t% + 1), 2)) = Null
    End If
Next t%
'CES Code
rs("userfullname") = frmLogin.userfullname
rs("userid") = frmLogin.userid
rs("ORINUMBER") = frmLogin.orinumber
rs("udate") = Format$(Now, "mm/dd/yyyy")
rs("utime") = Format$(Now, "hh:mm:ss")
'********
rs.Update
HI = incidentnumber
Call loadincident
incidentnumber = HI
'----- LOCATION INFO
For t% = 0 To 1
    Set rs = db.OpenRecordset("SELECT * FROM CODES WHERE CODE = '" + city(t%) + "' AND TYPE = 'city'")
    If rs.EOF Then
        rs.AddNew
        rs("CODE") = city(t%)
        For tt% = 0 To 2
            city(tt%).AddItem city(t%)
        Next tt%
        rs("TYPE") = "city"
        rs("DEFAULT") = "N"
        rs.Update
    End If
    Set rs = db.OpenRecordset("SELECT * FROM CODES WHERE CODE = '" + state(t%) + "' AND TYPE = 'state'")
    If rs.EOF Then
        rs.AddNew
        rs("CODE") = state(t%)
        For tt% = 0 To 2
            state(tt%).AddItem state(t%)
        Next tt%
        rs("TYPE") = "state"
        rs("DEFAULT") = "N"
        rs.Update
    End If
Next t%
    
'---- OFFENSES
For t% = 0 To 4
    If pickoffense(t%).ListIndex <> -1 Then
        UCRLIST(t%).ListIndex = -1
        For tt% = 0 To UCRLIST(t%).ListCount - 1
            If UCRLIST(t%).Selected(tt%) = True Then
                UCRLIST(t%).ListIndex = tt%
                tt% = UCRLIST(t%).ListCount - 1
            End If
        Next tt%
        Set rs = db.OpenRecordset("select * from offense where offense = " + Chr$(34) + pickoffense(t%).List(pickoffense(t%).ListIndex) + Chr$(34))
        If rs.EOF Then
            rs.AddNew
            rs("offense") = pickoffense(t%).List(pickoffense(t%).ListIndex)
            For tt% = 0 To 4
                pickoffense(tt%).AddItem pickoffense(t%).List(pickoffense(t%).ListIndex)
            Next tt%
            rs("ucr") = UCRLIST(t%).List(UCRLIST(t%).ListIndex)
            rs.Update
        Else
            rs.MoveFirst
            If rs("ucr") <> UCRLIST(t%).List(UCRLIST(t%).ListIndex) Then
                rs.Edit
                rs("offense") = pickoffense(t%).List(pickoffense(t%).ListIndex)
                rs("ucr") = UCRLIST(t%).List(UCRLIST(t%).ListIndex)
                rs.Update
            End If
        End If
    End If
Next t%
db.Close

On Error GoTo oderror2
od2:
Set db = OpenDatabase(nwl + "lawsuite.mdb")

'----- OFFICERS
If reportingofficer(0) > "" Then
    Set rs = db.OpenRecordset("select profidnum,profname, type from professionals where profname =" + Chr$(34) + reportingofficer(0) + Chr$(34))
    If rs.EOF Then
        rs.AddNew
    Else
        rs.MoveFirst
        rs.Edit
    End If
    rs("profname") = reportingofficer(0)
    rs("profidnum") = reportingofficeRunit(0)
    If rs.EOF Then
        reportingofficer(0).AddItem reportingofficer(0)
        reportingofficer(1).AddItem reportingofficer(0)
        approvingofficer.AddItem reportingofficer(0)
        followupofficer.AddItem reportingofficer(0)
    End If
    rs("type") = "D"
    rs.Update
End If
If reportingofficer(1) > "" Then
    Set rs = db.OpenRecordset("select profidnum,profname, type from professionals where profname =" + Chr$(34) + reportingofficer(1) + Chr$(34))
    If rs.EOF Then
        rs.AddNew
    Else
        rs.MoveFirst
        rs.Edit
    End If
    rs("profname") = reportingofficer(1)
    rs("profidnum") = reportingofficeRunit(1)
    If rs.EOF Then
        reportingofficer(0).AddItem reportingofficer(1)
        reportingofficer(1).AddItem reportingofficer(1)
        approvingofficer.AddItem reportingofficer(1)
        followupofficer.AddItem reportingofficer(1)
    End If
    rs("type") = "D"
    rs.Update
End If
If approvingofficer > "" Then
    Set rs = db.OpenRecordset("select profidnum,profname, type from professionals where profname =" + Chr$(34) + approvingofficer + Chr$(34))
    If rs.EOF Then
        rs.AddNew
    Else
        rs.MoveFirst
        rs.Edit
    End If
    rs("profname") = approvingofficer
    rs("profidnum") = approvingofficeRunit
    If rs.EOF Then
        reportingofficer(0).AddItem approvingofficer
        reportingofficer(1).AddItem approvingofficer
        approvingofficer.AddItem approvingofficer
        followupofficer.AddItem approvingofficer
    End If
    rs("type") = "D"
    rs.Update
End If
If followupofficer > "" Then
    Set rs = db.OpenRecordset("select profidnum,profname, type from professionals where profname =" + Chr$(34) + followupofficer + Chr$(34))
    If rs.EOF Then
        rs.AddNew
    Else
        rs.MoveFirst
        rs.Edit
    End If
    rs("profname") = followupofficer
    rs("profidnum") = FOLLOWUPOFFICERUNIT
    If rs.EOF Then
        reportingofficer(0).AddItem followupofficer
        reportingofficer(1).AddItem followupofficer
        approvingofficer.AddItem followupofficer
        followupofficer.AddItem followupofficer
    End If
    rs("type") = "D"
    rs.Update
End If


'-----PEOPLE
On Error GoTo 0
For t% = 0 To 2
    If (t% = 1 And vsname(1) = vsname(0)) Or vsname(t%) = "UNKNOWN" Then
        GoTo nextt
    End If
    Set rs = db.OpenRecordset("select * from people where dpnamelf =" + Chr$(34) + vsname(t%) + Chr$(34))
    If rs.EOF Then
        rs.AddNew
        For tt% = 0 To 2
            vsname(tt%).AddItem vsname(t%)
        Next tt%
    Else
        rs.MoveFirst
        rs.Edit
    End If
    rs("dpnamelf") = vsname(t%)
    rs("dphaddress") = address(t%)
    rs("dphaddress2") = city(t%)
    rs("hstate") = state(t%)
    rs("hzipcode") = zipcode(t%)
    rs("dpsort") = Left$(vsname(t%), 15)
    If t% < 2 Then
        If HOMEDAYPHONE(t%) > "" Then
            rs("dphphone") = HOMEDAYPHONE(t%)
        Else
        If HOMENIGHTPHONE(t%) > "" Then
            rs("dphphone") = HOMENIGHTPHONE(t%)
        End If
        End If
        If WORKDAYPHONE(t%) > "" Then
            rs("dpwphone") = HOMEDAYPHONE(t%)
        Else
        If WORKNIGHTPHONE(t%) > "" Then
            rs("dpwphone") = HOMENIGHTPHONE(t%)
        End If
        End If
        rs("resident") = Left$(resident(t%).List(resident(t%).ListIndex), 1)
    End If
    If t% > 0 Then
        rs("HEIGHT") = ht(t% - 1)
        rs("WEIGHT") = weight(t% - 1)
        rs("HAIR") = hair(t% - 1)
        rs("EYES") = eyes(t% - 1)
        rs("PECULIARITIES") = peculiarities(t% - 1)
    End If
    rs("race") = Left$(race(t%).List(race(t%).ListIndex), 1)
    rs("sex") = Left$(sex(t%).List(sex(t%).ListIndex), 1)
    rs("age") = age(t%)
    rs("ethnicity") = Left$(ethnicity(t%).List(ethnicity(t%).ListIndex), 1)
    hoLdname = vsname(t%)
    osort1$ = ""
    If Left$(hoLdname, 1) = " " Then
        hoLdname = Mid$(hoLdname, 2)
    End If
    If InStr(hoLdname, " CORP") > 0 Or InStr(hoLdname, ",INC") > 0 Or InStr(hoLdname, "COMPANY") > 0 Or InStr(hoLdname, "INC.") > 0 Then
        osort1$ = hoLdname
    End If
    tso$ = hoLdname
    If InStr(tso$, " et al") > 0 Then
        tso$ = Left$(tso$, InStr(tso$, " et al") - 1)
    End If
    If InStr(tso$, " et. al.") > 0 Then
        tso$ = Left$(tso$, InStr(tso$, " et. al.") - 1)
    End If
    If InStr(tso$, ",et al") > 0 Then
        tso$ = Left$(tso$, InStr(tso$, ",et al") - 1)
    End If
    If InStr(tso$, ",et. al.") > 0 Then
        tso$ = Left$(tso$, InStr(tso$, ",et. al.") - 1)
    End If
    If Right$(tso$, 1) = "," Then
        tso$ = Left$(tso$, Len(tso$) - 1)
    End If
    If InStr(tso$, "&") > 0 Then
        tso$ = Left$(tso$, InStr(tso$, "&") - 1)
    End If
    If Right$(tso$, 1) = "," Then
        tso$ = Left$(tso$, Len(tso$) - 1)
    End If
    firstspace% = 0
    While Right$(tso$, 1) = " " And Len(tso$) > 1
        tso$ = Left$(tso$, Len(tso$) - 1)
    Wend
    For tt% = 1 To Len(tso$)
        If Mid$(tso$, tt%, 1) = "," Then
            firstspace% = tt%
            tt% = Len(tso$)
        End If
    Next tt%
    If firstspace% = 0 Then
        If osort1$ = "" Then
            osort1$ = tso$
        End If
        GoTo rsupdate
    End If
    tempsort$ = Mid$(tso$, firstspace% + 1)
    If Left$(tempsort$, 1) = " " Then
        tempsort$ = Mid$(tempsort$, 2)
    End If
    tso$ = Left$(tso$, firstspace% - 1)
    If Right$(tso$, 1) = " " Then
        tso$ = Left$(tso$, Len(tso$) - 1)
    End If
    tempsort$ = tempsort$ + " " + tso$
    If osort1$ = "" Then
        osort1$ = tempsort$
    End If
    If InStr(osort1$, "JR.") Then
        If Mid$(osort1$, InStr(osort1$, "JR.") + 3, 1) = " " Then
            osort1$ = Left$(osort1$, InStr(osort1$, "JR.") - 1) + Mid$(osort1$, InStr(osort1$, "JR.") + 4) + ", JR."
        Else
            osort1$ = Left$(osort1$, InStr(osort1$, "JR.") - 1) + Mid$(osort1$, InStr(osort1$, "JR.") + 3) + ", JR."
    End If
    End If
    If InStr(osort1$, "SR.") Then
        If Mid$(osort1$, InStr(osort1$, "SR.") + 3, 1) = " " Then
            osort1$ = Left$(osort1$, InStr(osort1$, "SR.") - 1) + Mid$(osort1$, InStr(osort1$, "SR.") + 4) + ", SR."
        Else
            osort1$ = Left$(osort1$, InStr(osort1$, "SR.") - 1) + Mid$(osort1$, InStr(osort1$, "SR.") + 3) + ", SR."
    End If
    End If
    If InStr(osort1$, "III") Then
        If Mid$(osort1$, InStr(osort1$, "III") + 3, 1) = " " Then
            osort1$ = Left$(osort1$, InStr(osort1$, "III") - 1) + Mid$(osort1$, InStr(osort1$, "III") + 4) + ", III"
        Else
            osort1$ = Left$(osort1$, InStr(osort1$, "III") - 1) + Mid$(osort1$, InStr(osort1$, "III") + 3) + ", III"
        End If
    End If
    If InStr(osort1$, "IV") Then
        If Mid$(osort1$, InStr(osort1$, "IV") + 2, 1) = " " Then
            osort1$ = Left$(osort1$, InStr(osort1$, "IV") - 1) + Mid$(osort1$, InStr(osort1$, "IV") + 3) + ", III"
        Else
            osort1$ = Left$(osort1$, InStr(osort1$, "IV") - 1) + Mid$(osort1$, InStr(osort1$, "IV") + 2) + ", III"
        End If
    End If
    If Left$(osort1$, 1) = " " Then
        osort1$ = Mid$(osort1$, 2)
    End If
rsupdate:
    rs("dpname") = osort1$
    rs.Update
nextt:
Next t%
On Error Resume Next

'---- setfocus logic ----
'         onpaper.SetFocus
          If onpaper.Visible Then
              onpaper.SetFocus
           End If
db.Close
Exit Sub
oderror1:
If Err > 3200 Then
    Resume od1
Else
    Resume Next
End If
oderror2:
If Err > 3200 Then
    Resume od2
Else
    Resume Next
End If
End Sub
Private Sub sendcolon()
SendKeys ":"
End Sub

Private Sub weapontype_KeyDown(KeyCode As Integer, Shift As Integer)
FROMKEY = 1
End Sub
Private Sub drugamt_Change(Index As Integer)
ichanged = True
stra = drugamt(4).Text
If Index > 2 Then

End If
End Sub
Private Sub XGROUP_Click(Index As Integer)
'group(Index).Top = description(Index).Top - 1000
'group(Index).Left = description(Index).Left
On Error GoTo errh:
group(Index).Visible = True
If fromfind = 0 Then
'---- setfocus logic ----
'             group(index).SetFocus
          If group(Index).Visible Then
              group(Index).SetFocus
           End If
End If
Exit Sub
errh:
If Err.Number = 5 Then
    Resume Next
End If
End Sub

Private Sub incidentnumber_Change()
ichanged = True
If FROMXREF = 1 Then
    Call incidentnumber_Click
    FROMXREF = 0
End If
IncidentRptLoadedFromDB = False
VScroll1.Value = VScroll1.Min
End Sub
Private Sub FILLDATA(IDX As Integer)
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
Set rs = db.OpenRecordset("SELECT * FROM PEOPLE WHERE DPnameLF = " + Chr$(34) + vsname(IDX) + Chr$(34))
On Error Resume Next
If Not rs.EOF Then
    rs.MoveFirst
    If Not IsNull(rs("DPHADDRESS")) Then
        address(IDX) = rs("DPHADDRESS")
    Else
        address(IDX) = ""
    End If
    If Not IsNull(rs("DPHADDRESS2")) Then
        If InStr(rs("DPHADDRESS2"), ",") Then
            city(IDX) = Left$(rs("DPHADDRESS2"), InStr(rs("DPHADDRESS2"), ",") - 1)
            a$ = Mid$(rs("DPHADDRESS2"), InStr(rs("DPHADDRESS2"), ",") + 1)
            For t% = 1 To Len(a$)
                If Mid$(a$, t%, 1) <> " " Then
                    a$ = Mid$(a$, t%)
                    t% = t% + Len(a$)
                End If
            Next t%
            state(IDX) = Left$(a$, 2)
            a$ = Mid$(a$, 3)
            For t% = Len(a$) To 1 Step -1
                If Mid$(a$, t%, 1) = " " Then
                    zipcode(IDX) = Mid$(a$, t% + 1)
                    t% = 1
                End If
            Next t%
        Else
            city(IDX) = rs("DPHADDRESS2")
            state(IDX) = rs("HSTATE")
            zipcode(IDX) = rs("HZIPCODE")
        End If
    Else
        city(IDX) = ""
        state(IDX) = ""
    End If
    t% = IDX
    If t% < 2 Then
        HOMEDAYPHONE(t%) = ""
        HOMENIGHTPHONE(t%) = ""
        HOMEDAYPHONE(t%) = ""
        HOMENIGHTPHONE(t%) = ""
        ht(t%) = ""
        weight(t%) = ""
        hair(t%) = ""
        eyes(t%) = ""
        peculiarities(t%) = ""
        If Not IsNull(rs("dphphone")) Then
            HOMEDAYPHONE(t%) = rs("dphphone")
        Else
        If Not IsNull(rs("dphphone")) Then
            HOMENIGHTPHONE(t%) = rs("dphphone")
        End If
        End If
        If Not IsNull(rs("dpwphone")) Then
            HOMEDAYPHONE(t%) = rs("dpwphone")
        Else
        If Not IsNull(rs("dpwphone")) Then
            HOMENIGHTPHONE(t%) = rs("dpwphone")
        End If
        End If
        If Not IsNull(rs("HEIGHT")) Then
            ht(t%) = rs("HEIGHT")
        End If
        If Not IsNull(rs("WEIGHT")) Then
            weight(t%) = rs("WEIGHT")
        End If
        If Not IsNull(rs("HAIR")) Then
            hair(t%) = rs("HAIR")
        End If
        If Not IsNull(rs("EYES")) Then
            eyes(t%) = rs("EYES")
        End If
        If Not IsNull(rs("PECULIARITIES")) Then
            peculiarities(t%) = rs("PECULIARITIES")
        End If
        resident(t%).ListIndex = -1
        If Not IsNull(rs("resident")) Then
            For tt% = 0 To resident(t%).ListCount - 1
                If Left$(resident(t%).List(tt%), 1) = rs("resident") Then
                    resident(t%).ListIndex = tt%
                    tt% = resident(t%).ListCount - 1
                End If
            Next tt%
        End If
    End If
    If t% = 2 Then
        If IsDate(rs("birthdate")) Then
            BIRTHDATE = CStr(rs("birthdate"))
        Else
            BIRTHDATE = ""
        End If
        If Not IsNull(rs("mugshot")) Then
            mugshot.Picture = LoadPicture(rs("mugshot"))
        Else
            mugshot.Picture = LoadPicture()
        End If
    End If
    race(t%).ListIndex = -1
    If Not IsNull(rs("RACE")) Then
        For tt% = 0 To race(t%).ListCount - 1
            If Left$(race(t%).List(tt%), 1) = rs("RACE") Then
                race(t%).ListIndex = tt%
                tt% = race(t%).ListCount - 1
            End If
        Next tt%
    End If
    sex(t%).ListIndex = -1
    If Not IsNull(rs("SEX")) Then
        For tt% = 0 To sex(t%).ListCount - 1
            If Left$(sex(t%).List(tt%), 1) = rs("SEX") Then
                sex(t%).ListIndex = tt%
                tt% = sex(t%).ListCount - 1
            End If
        Next tt%
    End If
    ethnicity(t%).ListIndex = -1
    If Not IsNull(rs("ETHNICITY")) Then
        For tt% = 0 To ethnicity(t%).ListCount - 1
            If Left$(ethnicity(t%).List(tt%), 1) = rs("ETHNICITY") Then
                ethnicity(t%).ListIndex = tt%
                tt% = ethnicity(t%).ListCount - 1
            End If
        Next tt%
    End If
    age(t%) = ""
    If Not IsNull(rs("age")) Then
        age(t%) = rs("age")
    End If
End If
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If

End Sub
Private Sub DEFAULTCODESS()
Dim db As Database, rs As Recordset, itmx As ListItem
On Error GoTo oderror
od:
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("select * from codes WHERE TYPE = 'bias'")
On Error Resume Next
If rs.EOF Then
    db.Close
    Exit Sub
End If
rs.MoveFirst
BIAS.clear
widx% = 0
While Not rs.EOF
    BIAS.AddItem rs("code")
    If UCase(rs("default")) = "Y" Then
        BIAS.ListIndex = BIAS.ListCount - 1
    End If
    rs.MoveNext
Wend
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If

End Sub
Private Sub figure()
For t% = 0 To 41 Step 6
    temp = Val(totalvalue(t%))
    For tt% = t% + 1 To t% + 5
        temp = temp + Val(totalvalue(tt%))
    Next tt%
    Select Case t%
        Case 0 To 5
            TVSTOLEN = Format$(temp, "########0.00")
        Case 6 To 11
            TVDAMAGED = Format$(temp, "########0.00")
        Case 12 To 17
            TVBURNED = Format$(temp, "########0.00")
        Case 18 To 23
            TVRECOVERED = Format$(temp, "########0.00")
        Case 24 To 29
            TVSEIZED = Format$(temp, "########0.00")
        Case 30 To 35
            TVCOUNTERFEIT = Format$(temp, "########0.00")
        Case 36 To 41
            TVUNKNOWN = Format$(temp, "########0.00")
    End Select
Next t%

End Sub
Private Sub monthlyreport(inpm, inpy As String)
On Error GoTo 0
Screen.MousePointer = 11
daa$ = inpm + "/1/" + inpy
Select Case Val(inpm)
    Case 1, 3, 5, 7, 8, 10, 12
        dbb$ = inpm + "/31/" + inpy
    Case 4, 6, 9, 11
        dbb$ = inpm + "/30/" + inpy
    Case 2
        r1% = Val(inpy) / 4
        r2! = Val(inpy) / 4
        If r1% <> r2! Then
            dbb$ = inpm + "/28/" + inpy$
        Else
            dbb$ = inpm + "/29/" + inpy$
        End If
End Select
Dim crime(200) As String, cidx As Integer, crimetot(200) As Single
Dim db, db2, db3 As Database, rs, rs2, rs3 As Recordset
cidx = 0
Set db = OpenDatabase(nwi + "incident.mdb")
Set db2 = OpenDatabase(nws + "service.mdb")
Set db3 = OpenDatabase(nwb + "booking.mdb")
Set rs = db.OpenRecordset("select ucr1,ucr2,ucr3,ucr4,ucr5,ucr6,ucr7,ucr8,ucr9,ucr10 from incidentsupport where incidentnumber in (select incidentnumber from incidentreportc where dateofoffense1 between #" + daa$ + "# and #" + dbb$ + "#)")
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        For t% = 1 To 10
            If Not IsNull(rs("ucr" + Mid$(Str$(t%), 2))) Then
                tempucr$ = rs("ucr" + Mid$(Str$(t%), 2))
                foundit% = 0
                For tt% = 1 To cidx
                    If tempucr$ = crime(tt%) Then
                        foundit% = tt%
                        tt% = cidx
                    End If
                Next tt%
                If foundit% = 0 Then
                    cidx = cidx + 1
                    crime(cidx) = tempucr$
                    crimetot(cidx) = 1
                Else
                    crimetot(foundit%) = crimetot(foundit) + 1
                End If
            End If
        Next t%
        rs.MoveNext
    Wend
End If
PAGE% = 0
Printer.Orientation = 1
GoSub header
Printer.FontSize = 10
Printer.FontBold = True
Printer.Print "INCIDENT OFFENSE BREAKDOWN"
Printer.Print
Printer.FontUnderline = True
Printer.Print "Category";
Printer.FontBold = False
Printer.Print Tab(100);
Printer.FontBold = True
Printer.Print "Count"
linect% = linect% + 2
Printer.FontBold = False
Printer.FontUnderline = False
For t% = 1 To cidx
    Set rs = db.OpenRecordset("select code from ucr where abbrev = '" + crime(t%) + "'")
    If Not rs.EOF Then
        rs.MoveFirst
        Printer.Print rs("code");
    Else
        Printer.Print crime(t%);
    End If
    Printer.Print Tab(100); Format$(crimetot(t%), "####0")
    linect% = linect% + 1
    If linect% > 55 Then
        Printer.NewPage
        GoSub header
    End If
Next t%

Printer.Print
Printer.FontSize = 10
Printer.FontBold = True
Printer.Print "SERVICE CALL BREAKDOWN"
Printer.Print
Printer.FontUnderline = True
Printer.Print "Category";
Printer.FontBold = False
Printer.Print Tab(100);
Printer.FontBold = True
Printer.Print "Count"
linect% = linect% + 3
Printer.FontBold = False
Printer.FontUnderline = False
cidx = 0
Set rs = db2.OpenRecordset("select * from service where currentdate between #" + daa$ + "# and #" + dbb$ + "#")
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        If rs("alarm") = 1 Then
            tempucr$ = "Alarm"
        End If
        If rs("unlocking") = 1 Then
            tempucr$ = "Unlocking Vehicle"
        End If
        If rs("property") = 1 Then
            tempucr$ = "Property Check"
        End If
        If rs("funeral") = True Then
            tempucr$ = "Funeral Escort"
        End If
        If rs("house") = True Then
            tempucr$ = "House Moving Escort"
        End If
        If rs("mental") = True Then
            tempucr$ = "Mental Transport"
        End If
        If rs("escort") = 1 And Not IsNull(rs("escortother")) And rs("escortother") > "" Then
            tempucr$ = rs("escortother")
        End If
        If rs("warrant") = 1 Then
            tempucr$ = "Warrant"
        End If
        If rs("unfounded") = 1 Then
            tempucr$ = "Unfounded"
        End If
        If rs("other") = 1 Then
            tempucr$ = rs("otherspecify")
        End If
        a = rs("casenumber")
        foundit% = 0
        For tt% = 1 To cidx
            If tempucr$ = crime(tt%) Then
                foundit% = tt%
                tt% = cidx
            End If
        Next tt%
        If foundit% = 0 Then
            cidx = cidx + 1
            crime(cidx) = tempucr$
            crimetot(cidx) = 1
        Else
            crimetot(foundit%) = crimetot(foundit) + 1
        End If
        rs.MoveNext
    Wend
End If
For t% = 1 To cidx
    Set rs = db.OpenRecordset("select code from ucr where abbrev = '" + crime(t%) + "'")
    If Not rs.EOF Then
        rs.MoveFirst
        Printer.Print rs("code");
    Else
        Printer.Print crime(t%);
    End If
    Printer.Print Tab(100); Format$(crimetot(t%), "####0")
    linect% = linect% + 1
    If linect% > 55 Then
        Printer.NewPage
        GoSub header
    End If
Next t%

Printer.Print
Printer.FontSize = 10
Printer.FontBold = True
Printer.Print "DRUG SEIZURE ACTIVITY"
Printer.Print
Printer.FontUnderline = True
Printer.Print "Type of Drug";
Printer.FontBold = False
Printer.Print Tab(70);
Printer.FontBold = True
Printer.Print "Amount";
Printer.FontBold = False
Printer.Print Tab(100);
Printer.FontBold = True
Printer.Print "Measurement"
linect% = linect% + 3
Printer.FontBold = False
Printer.FontUnderline = False
cidx = 0
Set rs = db.OpenRecordset("select * from incidentreporto where incidentnumber in (select incidentnumber from incidentreportc where dateofoffense1 between #" + daa$ + "# and #" + dbb$ + "#)")
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        For t% = 1 To 6
            For tt% = 1 To 3
                If Not IsNull(rs("ptypeofdrug" + Mid$(Str$(t%), 2) + Mid$(Str$(tt%), 2))) Then
                    tempdrug$ = rs("ptypeofdrug" + Mid$(Str$(t%), 2) + Mid$(Str$(tt%), 2)) + "/" + rs("pdrugmeasurement" + Mid$(Str$(t%), 2) + Mid$(Str$(tt%), 2))
                    foundit% = 0
                    For ttt% = 1 To cidx
                        If tempdrug$ = crime(ttt%) Then
                            foundit% = ttt%
                            ttt% = cidx
                        End If
                    Next ttt%
                    If foundit% = 0 Then
                        cidx = cidx + 1
                        crime(cidx) = tempdrug$
                        crimetot(cidx) = rs("pdrugamt" + Mid$(Str$(t%), 2) + Mid$(Str$(tt%), 2))
                    Else
                        crimetot(foundit%) = crimetot(foundit) + rs("pdrugamt" + Mid$(Str$(t%), 2) + Mid$(Str$(tt%), 2))
                    End If
                End If
            Next tt%
        Next t%
        rs.MoveNext
    Wend
End If
Set rs = db.OpenRecordset("select * from supplementalsupport where incidentnumber in (select incidentnumber from incidentreportc where dateofoffense1 between #" + daa$ + "# and #" + dbb$ + "#)")
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        For t% = 1 To 6
            For tt% = 1 To 3
                If Not IsNull(rs("ptypeofdrug" + Mid$(Str$(t%), 2) + Mid$(Str$(tt%), 2))) Then
                    tempdrug$ = rs("ptypeofdrug" + Mid$(Str$(t%), 2) + Mid$(Str$(tt%), 2)) + "/" + rs("pdrugmeasurement" + Mid$(Str$(t%), 2) + Mid$(Str$(tt%), 2))
                    foundit% = 0
                    For ttt% = 1 To cidx
                        If tempdrug$ = crime(ttt%) Then
                            foundit% = ttt%
                            ttt% = cidx
                        End If
                    Next ttt%
                    If foundit% = 0 Then
                        cidx = cidx + 1
                        crime(cidx) = tempdrug$
                        crimetot(cidx) = rs("pdrugamt" + Mid$(Str$(t%), 2) + Mid$(Str$(tt%), 2))
                    Else
                        crimetot(foundit%) = crimetot(foundit) + rs("pdrugamt" + Mid$(Str$(t%), 2) + Mid$(Str$(tt%), 2))
                    End If
                End If
            Next tt%
        Next t%
        rs.MoveNext
    Wend
End If
For t% = 1 To cidx
    Set rs = db.OpenRecordset("select code from codes where code like '" + Left$(crime(t%), 1) + " = *'")
    If Not rs.EOF Then
        rs.MoveFirst
        Printer.Print Mid$(rs("code"), 5);
    Else
        Printer.Print crime(t%);
    End If
    Printer.Print Tab(70); Format$(crimetot(t%), "####0.00"); Tab(100); Mid$(crime(t%), InStr(crime(t%), "/") + 1)
    linect% = linect% + 1
    If linect% > 55 Then
        Printer.NewPage
        GoSub header
    End If
Next t%

Printer.Print
Printer.FontSize = 10
Printer.FontBold = True
Printer.Print "WEAPON SEIZURE ACTIVITY"
Printer.Print
Printer.FontUnderline = True
Printer.Print "Type of Weapon";
Printer.FontBold = False
Printer.Print Tab(100);
Printer.FontBold = True
Printer.Print "Value"
linect% = linect% + 3
Printer.FontBold = False
Printer.FontUnderline = False
cidx = 0
Set rs = db.OpenRecordset("select * from incidentreporto where (seizedvalue1 > 0 or seizedvalue2 > 0 or seizedvalue3 > 0 or seizedvalue4 > 0 or seizedvalue5 > 0 or seizedvalue6 > 0) and incidentnumber in (select incidentnumber from incidentsupport where group1 = '13' or group2 = '13' or group3 = '13' or group4 = '13' or group5 = '13' or group6 = '13') and incidentnumber in (select incidentnumber from incidentreportc where dateofoffense1 between #" + daa$ + "# and #" + dbb$ + "#)")
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        For t% = 1 To 6
            If rs("seizedvalue" + Mid$(Str$(t%), 2)) > 0 Then
                Set rs2 = db.OpenRecordset("select * from incidentsupport where incidentnumber = '" + rs("incidentnumber") + "'")
                If Not rs2.EOF Then
                    rs2.MoveFirst
                    If rs2("group" + Mid$(Str$(t%), 2)) = "13" Then
                        tempweap$ = rs("type" + Mid$(Str$(t%), 2))
                        foundit% = 0
                        For tt% = 1 To cidx
                            If tempweap$ = crime(tt%) Then
                                foundit% = tt%
                                tt% = cidx
                            End If
                        Next tt%
                        If foundit% = 0 Then
                            cidx = cidx + 1
                            crime(cidx) = tempweap$
                            crimetot(cidx) = rs("seizedvalue" + Mid$(Str$(t%), 2))
                        Else
                            crimetot(foundit%) = crimetot(foundit) + rs("pdrugamt" + Mid$(Str$(t%), 2) + Mid$(Str$(tt%), 2))
                        End If
                    End If
                End If
            End If
        Next t%
        rs.MoveNext
    Wend
End If
Set rs = db.OpenRecordset("select * from supplemental where (seizedvalue1 > 0 or seizedvalue2 > 0 or seizedvalue3 > 0 or seizedvalue4 > 0 or seizedvalue5 > 0 or seizedvalue6 > 0) and incidentnumber in (select incidentnumber from supplementalsupport where group1 = '13' or group2 = '13' or group3 = '13' or group4 = '13' or group5 = '13' or group6 = '13') and incidentnumber in (select incidentnumber from incidentreportc where dateofoffense1 between #" + daa$ + "# and #" + dbb$ + "#)")
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        For t% = 1 To 6
            If rs("seizedvalue" + Mid$(Str$(t%), 2)) > 0 Then
                Set rs2 = db.OpenRecordset("select * from supplementalsupport where incidentnumber = '" + rs("incidentnumber") + "'")
                If Not rs2.EOF Then
                    rs2.MoveFirst
                    If rs2("group" + Mid$(Str$(t%), 2)) = "13" Then
                        tempweap$ = rs("type" + Mid$(Str$(t%), 2))
                        foundit% = 0
                        For tt% = 1 To cidx
                            If tempweap$ = crime(tt%) Then
                                foundit% = tt%
                                tt% = cidx
                            End If
                        Next tt%
                        If foundit% = 0 Then
                            cidx = cidx + 1
                            crime(cidx) = tempweap$
                            crimetot(cidx) = rs("seizedvalue" + Mid$(Str$(t%), 2))
                        Else
                            crimetot(foundit%) = crimetot(foundit) + rs("pdrugamt" + Mid$(Str$(t%), 2) + Mid$(Str$(tt%), 2))
                        End If
                    End If
                End If
            End If
        Next t%
        rs.MoveNext
    Wend
End If
For t% = 1 To cidx
    Printer.Print crime(t%); Tab(100); Format$(crimetot(t%), "######0.00")
    linect% = linect% + 1
    If linect% > 55 Then
        Printer.NewPage
        GoSub header
    End If
Next t%

Printer.Print
Printer.FontSize = 10
Printer.FontBold = True
Printer.Print "BOOKINGS"
Printer.Print
Printer.FontUnderline = True
Printer.Print "Primary Offense";
Printer.FontBold = False
Printer.Print Tab(100);
Printer.FontBold = True
Printer.Print "Count"
linect% = linect% + 3
Printer.FontBold = False
Printer.FontUnderline = False
cidx = 0
Set rs = db3.OpenRecordset("select * from booking where dateofarrest between #" + daa$ + "# and #" + dbb$ + "#")
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        If Not IsNull(rs("chargea")) Then
                tempucr$ = rs("chargea")
                foundit% = 0
                For tt% = 1 To cidx
                    If tempucr$ = crime(tt%) Then
                        foundit% = tt%
                        tt% = cidx
                    End If
                Next tt%
                If foundit% = 0 Then
                    cidx = cidx + 1
                    crime(cidx) = tempucr$
                    crimetot(cidx) = 1
                Else
                    crimetot(foundit%) = crimetot(foundit) + 1
                End If
        End If
        rs.MoveNext
    Wend
End If
For t% = 1 To cidx
    Printer.Print crime(t%); Tab(100); Format$(crimetot(t%), "####0")
    linect% = linect% + 1
    If linect% > 55 Then
        Printer.NewPage
        GoSub header
    End If
Next t%

Printer.EndDoc
db.Close
Screen.MousePointer = 0
Exit Sub
header:
PAGE% = PAGE% + 1
Printer.FontName = "MS Sans Serif"
Printer.FontSize = 14
Printer.FontBold = True
Set rs = db.OpenRecordset("select * from system")
If Not rs.EOF Then
    rs.MoveFirst
    Printer.Print rs("office");
End If
Printer.Print Tab(75); PAGE%
Printer.Print "Monthly Incident, Booking, and Service Call Report"
Printer.FontSize = 12
Select Case Val(inpm)
    Case 1
        Printer.Print "January " + inpy
    Case 2
        Printer.Print "February " + inpy
    Case 3
        Printer.Print "March " + inpy
    Case 4
        Printer.Print "April " + inpy
    Case 5
        Printer.Print "May " + inpy
    Case 6
        Printer.Print "June " + inpy
    Case 7
        Printer.Print "July " + inpy
    Case 8
        Printer.Print "August " + inpy
    Case 9
        Printer.Print "September " + inpy
    Case 10
        Printer.Print "October " + inpy
    Case 11
        Printer.Print "November " + inpy
    Case 12
        Printer.Print "December " + inpy
End Select
Printer.Print
Printer.Print
linect% = 5
Printer.FontSize = 10
Printer.FontBold = False
Return
End Sub
Private Sub spellcheck(DONE As Boolean)
Dim wd As New Word.Application
Dim wdsp As Word.SpellingSuggestions
On Error GoTo cmdCheckErr
GETOUT% = 0
lstframe.Visible = False
While tempword > "" And GETOUT% = 0
    wd.Visible = False
    While Left$(tempword, 1) = " "
        tempword = Mid$(tempword, 2)
    Wend
    stopper% = Len(tempword) + 1
    For t% = 1 To Len(tempword)
        If InStr("!()[{]};:,./? " + Chr$(34), Mid$(tempword, t%, 1)) Then
            stopper% = t%
            t% = Len(tempword)
        End If
    Next t%
    checkword = LCase(Left$(tempword, stopper% - 1))
    
    
    If stopper% < Len(tempword) Then
        tempword = Mid$(tempword, stopper% + 1)
    Else
        tempword = ""
    End If
    
    
    
    
    If checkword > "" Then
        wd.Documents.add
        Set wdsp = wd.GetSpellingSuggestions(checkword)
        
        If wdsp.Count > 0 Or wdsp.SpellingErrorType = wdSpellingNotInDictionary Then
            lstsuggestions.clear
            lstframe.Visible = True
            'RLB code
            lstframe.Top = NARRATIVE.Top - lstframe.Height
            lstframe.Left = NARRATIVE.Left + CLng(lstframe.Width * 0.5)
            '***********
            GETOUT% = 1
        Else
            lstframe.Visible = False
        End If
        
        For i% = 1 To wdsp.Count
            lstsuggestions.AddItem wdsp(i%).Name
        Next i%
        If wdsp.SpellingErrorType = wdSpellingNotInDictionary And wdsp.Count = 0 Then
            lstsuggestions.AddItem "Not found in dictionary."
        End If
        wd.Documents.Close
        wd.Visible = False
    Else
        
    End If
Wend
wd.Quit
Set wd = Nothing
Screen.MousePointer = 0
If tempword = "" And lstframe.Visible = False Then
    DONE = True
Else
    DONE = False
End If
Exit Sub
cmdCheckErr:
'MsgBox Err.description
Resume Next
End Sub


Private Sub loadincident()
Dim db As Database, rs As Recordset, tabname As String
On Error GoTo oderror1
od1:
Set db = OpenDatabase(nwi + "incident.mdb")
incidentnumber.clear
Set rs = db.OpenRecordset("select distinct incidentnumber from incidentSUPPort WHERE TEMP = 'N' order by incidentnumber desc")
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        incidentnumber.AddItem rs("incidentnumber")
        rs.MoveNext
    Wend
End If
On Error Resume Next
db.Close
Exit Sub
oderror1:
If Err > 3200 Then
    Resume od1
Else
    Resume Next
End If

End Sub
Friend Sub getincident()
HI = incidentnumber
Screen.MousePointer = 11
Call clearroutine(3)
incidentnumber = HI
holdoff.Visible = True
Call findincident(2)
ichanged = False
Picture2.Refresh
Screen.MousePointer = 0
holdoff.Visible = True
'---- setfocus logic ----
'         onpaper.SetFocus
          If onpaper.Visible Then
              onpaper.SetFocus
           End If
optimer.Enabled = False
If holdoff.Visible = True Then
    optimer.Enabled = True
End If
On Error GoTo 0
VScroll1.Value = VScroll1.Min
End Sub
Private Sub forcesave()
savemask$ = Mid$(incidentnumber, 5)
For t% = 0 To incidentnumber.ListCount - 1
    If Left$(incidentnumber.List(t%), Len(savemask$)) = savemask$ Then
        incidentnumber = incidentnumber.List(t%)
        incidentnumber.Refresh
        Call incidentnumber_Click
        editerr% = 0
        Picture1.Visible = False
        VScroll1.Visible = False
        POPMSG$ = ""
        Call editevent(editerr%, POPMSG$)
        If editerr% = 0 Then
            Call editvictim(editerr%, POPMSG$)
            If editerr% = 0 Then
                Call editsubject(editerr%, POPMSG$)
                If editerr% = 0 Then
                    Call editadministrative(editerr%, POPMSG$)
                    If editerr% = 0 Then
                        Call editproperty(editerr%, POPMSG$)
                    End If
                End If
            End If
        End If
        If editerr% = 0 Then
            Call saveincident
            ichanged = False
        Else
            MsgBox POPMSG$, 48, "Genesis Error Log"
            Picture1.Visible = True
            VScroll1.Visible = True
            msg = MsgBox("Error detected. Continue?", 4, "")
            If msg <> 6 Then
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If
    End If
Next t%
Picture1.Visible = True
VScroll1.Visible = True
Screen.MousePointer = 0
End Sub
Private Sub setminorlist(Index As Integer)
HOLDINDEX = minorlist(Index).ListIndex
minorlist(Index).clear
Dim db As Database, rs As Recordset
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("select minor from pgroup where major = " + Chr$(34) + majorlist(Index).List(majorlist(Index).ListIndex) + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        minorlist(Index).AddItem rs("minor")
        rs.MoveNext
    Wend
End If
If (HOLDINDEX >= -1) And (HOLDINDEX < minorlist(Index).ListCount) Then
    minorlist(Index).ListIndex = HOLDINDEX
Else
    minorlist(Index).ListIndex = -1
End If
db.Close

End Sub

'RLB Code
Private Function AutoSelectOffense(ByVal intIndex As Integer, ByVal strSelect As String) As Boolean
    Dim intx As Integer
    Dim lstOffense As ListBox
    Dim blnFound As Boolean
    
    On Error GoTo errh:
    
    Set lstOffense = pickoffense(intIndex)
    
    For intx = 0 To lstOffense.ListCount - 1
        If lstOffense.List(intx) = strSelect Then
            lstOffense.ListIndex = intx
            blnFound = True
        End If
    Next intx
    
    AutoSelectOffense = blnFound
    
errh:
    Set lstOffense = Nothing
End Function
Private Sub ShowApplicableContainers(objControl As Object)
  
    Picture1.Visible = True
    VScroll1.Visible = True

    Dim intNumContainerLevelsInPicture1 As Integer
    Dim objHoldOriginal As Object
    
    Set objHoldOriginal = objControl
    
    On Error GoTo errh
    
    For xx% = 0 To 10
        DoEvents
    Next xx%
    
    intNumContainerLevelsInPicture1 = 1
    
    Do
    
        If Not (objControl.Container Is Picture2) Then
            intNumContainerLevelsInPicture1 = intNumContainerLevelsInPicture1 + 1
            Set objControl = objControl.Container
        Else
            Exit Do
        End If
        
    Loop While True
    
    Set objControl = objHoldOriginal
    
    For x% = 1 To intNumContainerLevelsInPicture1
        Select Case UCase(objControl.Name)
            Case "PINFOFRAME"
                objControl.Visible = False
                Call Command2_Click(objControl.Index)
            Case "HOMOCIDE"
                objControl.Visible = False
                Call Command13_Click(objControl.Index)
            Case "OFFENSEFRAME"
                objControl.Visible = False
                Call ao_Click
            Case "VUCRF"
                objControl.Visible = False
                Call Command1_Click
            Case "GACTIVITY"
                objControl.Visible = False
                Call Command12_Click(objControl.Index)
            Case "ACTIVITY"
                objControl.Visible = False
                Call Command12_Click(objControl.Index)
            Case "LACTIVITY"
                objControl.Visible = False
                Call Command14_Click
            Case "CASEFRAME"
                objControl.Visible = False
                Call Command19_Click
            Case "PUCRLIST"
                a = 1
         '''''       objControl.Visible = False
         '''''       Call Command2_Click(objControl.index)
            Case "LOOKUPFRAME"
                objControl.Visible = False
                Call Command3_Click(objControl.Index)
            Case "BIAS"
                objControl.Visible = False
                Call Command6_Click
            Case "SDRUGFRAME"
                objControl.Visible = False
                If objControl.Index = 0 Then
                    Call Command7_Click
                ElseIf objControl.Index = 1 Then
                    Call Command9_Click
                Else
'                    objControl.Visible = True
 '                   VScroll1 = 10000
                    sdrugframe(objControl.Index).Left = 2000
                    sdrugframe(objControl.Index).Top = description(objControl.Index).Top - sdrugframe(objControl.Index + 2).Height - 100
                    sdrugframe(objControl.Index).Visible = True
                End If
            Case "NUMVEHICLE"
                objControl.Visible = False
                numvehicle(objControl.Index).Visible = True
            Case "DRUGAMT"
                objControl.Visible = True
                Call drugtype_Click(objControl.Index)
            Case "EXCEPTIONALCLEARANCEDATE"
                objControl.Visible = False
                Call extraditiondenied_Click
            Case "ADDITIONAL"
                objControl.Visible = False
                Call HOMOCIDE_LostFocus(objControl.Index)
            Case "HOLDOFF"
                objControl.Visible = False
                Call incidentnumber_Click
            Case "RELATIONSHIPFRAME"
                objControl.Visible = False
                Call setrel_Click(objControl.Index)
            Case "SUBLIST"
                objControl.Visible = False
                Call subcode_Click(objControl.Index)
            Case "DATERECOVERED"
                objControl.Visible = True
                'Call totalvalue_LostFocus(objControl.index)
            Case "UCRLIST"
                objControl.Visible = False
                Call ucrselect(objControl.Index)
            Case "GROUP"
                fromfind = 1
                objControl.Visible = False
                Call XGROUP_Click(objControl.Index)
                fromfind = 0
        End Select
        If objControl.Top > (-1 * Picture2.Top) And objControl.Top + objControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
        Else
        If objControl.Top > 500 Then
            VScroll1 = objControl.Top - 500
        Else
            VScroll1 = 0
        End If
        End If
        Set objControl = objControl.Container
    Next x%
GETOUT:
On Error Resume Next
Exit Sub
errh:
    
    Set objControl = Nothing
    Set objHoldOriginal = Nothing
    Resume GETOUT

End Sub
'*************
Private Sub weight_Change(Index As Integer)
ichanged = True

End Sub



Private Sub forcesavelist()
Open "c:\savelist" For Input As #1
While Not EOF(1)
    Line Input #1, savemask$
    incidentnumber = savemask$
    incidentnumber.Refresh
    Call incidentnumber_Click
    editerr% = 0
    Picture1.Visible = False
    VScroll1.Visible = False
    POPMSG$ = ""
    Call editevent(editerr%, POPMSG$)
    If editerr% = 0 Then
        Call editvictim(editerr%, POPMSG$)
        If editerr% = 0 Then
            Call editsubject(editerr%, POPMSG$)
            If editerr% = 0 Then
                Call editadministrative(editerr%, POPMSG$)
                If editerr% = 0 Then
                    Call editproperty(editerr%, POPMSG$)
                End If
            End If
        End If
    End If
    If editerr% = 0 Then
        Call saveincident
        ichanged = False
    Else
        MsgBox POPMSG$, 48, "Genesis Error Log"
        Picture1.Visible = True
        VScroll1.Visible = True
        msg = MsgBox("Error detected. Continue?", 4, "")
        If msg <> 6 Then
            Screen.MousePointer = 0
            Close #1
            Exit Sub
        End If
    End If
Wend
Close #1
Picture1.Visible = True
VScroll1.Visible = True
Screen.MousePointer = 0

End Sub

Private Sub WORKDAYPHONE_Change(Index As Integer)
ichanged = True

End Sub

Private Sub WORKNIGHTPHONE_Change(Index As Integer)
ichanged = True

End Sub

Private Sub year12_Click()
ichanged = True

End Sub

Private Sub year45_Click()
ichanged = True

End Sub

Private Sub zipcode_Change(Index As Integer)
ichanged = True

End Sub
Friend Sub LOADMAJOR()
Dim db As Database, rs As Recordset, itmx As ListItem
On Error GoTo oderror1
od1:
Set db = OpenDatabase(nwi + "incident.mdb")
On Error Resume Next
For t% = 0 To 5
    majorlist(t%).clear
Next t%
Set rs = db.OpenRecordset("select distinct major from pgroup order by major")
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        For t% = 0 To 5
            majorlist(t%).AddItem rs("major")
        Next t%
        rs.MoveNext
    Wend
End If
On Error Resume Next
db.Close
Exit Sub
oderror1:
If Err > 3200 Then
    Resume od1
Else
    Resume Next
End If

End Sub
Friend Sub loadoffense()
Dim db As Database, rs As Recordset, itmx As ListItem, HO(4) As String
On Error GoTo oderror1
od1:
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("select * from codes")
For t% = 0 To 4
    If pickoffense(t%).ListIndex > -1 Then
        HO(t%) = pickoffense(t%).List(pickoffense(t%).ListIndex)
    End If
    pickoffense(t%).clear
Next t%
Set rs = db.OpenRecordset("select DISTINCT offense from offense")
If Not rs.EOF Then
    rs.MoveFirst
End If
While Not rs.EOF
    For tt% = 0 To 4
        pickoffense(tt%).AddItem rs("offense")
        If rs("OFFENSE") = HO(tt%) Then
            pickoffense(tt%).ListIndex = pickoffense(tt%).ListCount - 1
        End If
    Next tt%
    rs.MoveNext
Wend
db.Close
Exit Sub
oderror1:
If Err > 3200 Then
    Resume od1
Else
    Resume Next
End If

End Sub
