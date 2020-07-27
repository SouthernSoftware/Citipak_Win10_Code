VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form sinciden 
   BackColor       =   &H00808000&
   Caption         =   "Genesis Supplemental Incident Report version 1.0                        "
   ClientHeight    =   7860
   ClientLeft      =   165
   ClientTop       =   525
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   11925
   WindowState     =   1  'Minimized
   Begin VB.PictureBox Picture1 
      Height          =   7200
      Left            =   0
      ScaleHeight     =   7140
      ScaleWidth      =   11595
      TabIndex        =   231
      Top             =   615
      Width           =   11655
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   19800
         Left            =   240
         Picture         =   "sinciden.frx":0000
         ScaleHeight     =   19800
         ScaleWidth      =   11565
         TabIndex        =   72
         Top             =   -1920
         Width           =   11565
         Begin VB.Frame sdrugframe 
            BackColor       =   &H00808000&
            BorderStyle     =   0  'None
            Height          =   2460
            Index           =   9
            Left            =   6960
            TabIndex        =   391
            Top             =   11760
            Visible         =   0   'False
            Width           =   6855
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   27
               Left            =   120
               TabIndex        =   392
               Top             =   240
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   28
               Left            =   120
               TabIndex        =   393
               Top             =   960
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   29
               Left            =   120
               TabIndex        =   394
               Top             =   1680
               Width           =   2895
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   27
               Left            =   4440
               TabIndex        =   395
               Top             =   240
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   27
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   396
               Top             =   240
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   28
               Left            =   4440
               TabIndex        =   397
               Top             =   960
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   28
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   398
               Top             =   960
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   29
               Left            =   4440
               TabIndex        =   399
               Top             =   1680
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   29
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   400
               Top             =   1680
               Width           =   1215
            End
            Begin VB.CommandButton closedrugframes 
               Caption         =   "Close"
               Height          =   255
               Index           =   9
               Left            =   3120
               TabIndex        =   401
               Top             =   2160
               Width           =   1215
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Type Of Drug                                              Amount                 Measurement"
               ForeColor       =   &H0000FFFF&
               Height          =   255
               Index           =   9
               Left            =   120
               TabIndex        =   402
               Top             =   0
               Width           =   6615
            End
         End
         Begin VB.Frame sdrugframe 
            BackColor       =   &H00808000&
            BorderStyle     =   0  'None
            Height          =   2460
            Index           =   8
            Left            =   6720
            TabIndex        =   379
            Top             =   12120
            Visible         =   0   'False
            Width           =   6855
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   24
               Left            =   120
               TabIndex        =   380
               Top             =   240
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   25
               Left            =   120
               TabIndex        =   381
               Top             =   960
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   26
               Left            =   120
               TabIndex        =   382
               Top             =   1680
               Width           =   2895
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   24
               Left            =   4440
               TabIndex        =   383
               Top             =   240
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   24
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   384
               Top             =   240
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   25
               Left            =   4440
               TabIndex        =   385
               Top             =   960
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   25
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   386
               Top             =   960
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   26
               Left            =   4440
               TabIndex        =   387
               Top             =   1680
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   26
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   388
               Top             =   1680
               Width           =   1215
            End
            Begin VB.CommandButton closedrugframes 
               Caption         =   "Close"
               Height          =   255
               Index           =   8
               Left            =   3120
               TabIndex        =   389
               Top             =   2160
               Width           =   1215
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Type Of Drug                                              Amount                 Measurement"
               ForeColor       =   &H0000FFFF&
               Height          =   255
               Index           =   8
               Left            =   120
               TabIndex        =   390
               Top             =   0
               Width           =   6615
            End
         End
         Begin VB.Frame sdrugframe 
            BackColor       =   &H00808000&
            BorderStyle     =   0  'None
            Height          =   2460
            Index           =   7
            Left            =   7920
            TabIndex        =   367
            Top             =   12120
            Visible         =   0   'False
            Width           =   6855
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   21
               Left            =   120
               TabIndex        =   368
               Top             =   240
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   22
               Left            =   120
               TabIndex        =   369
               Top             =   960
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   23
               Left            =   120
               TabIndex        =   370
               Top             =   1680
               Width           =   2895
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   21
               Left            =   4440
               TabIndex        =   371
               Top             =   240
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   21
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   372
               Top             =   240
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   22
               Left            =   4440
               TabIndex        =   373
               Top             =   960
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   22
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   374
               Top             =   960
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   23
               Left            =   4440
               TabIndex        =   375
               Top             =   1680
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   23
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   376
               Top             =   1680
               Width           =   1215
            End
            Begin VB.CommandButton closedrugframes 
               Caption         =   "Close"
               Height          =   255
               Index           =   7
               Left            =   3120
               TabIndex        =   377
               Top             =   2160
               Width           =   1215
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Type Of Drug                                              Amount                 Measurement"
               ForeColor       =   &H0000FFFF&
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   378
               Top             =   0
               Width           =   6615
            End
         End
         Begin VB.Frame sdrugframe 
            BackColor       =   &H00808000&
            BorderStyle     =   0  'None
            Height          =   2460
            Index           =   6
            Left            =   7800
            TabIndex        =   355
            Top             =   11880
            Visible         =   0   'False
            Width           =   6855
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   18
               Left            =   120
               TabIndex        =   356
               Top             =   240
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   19
               Left            =   120
               TabIndex        =   357
               Top             =   960
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   20
               Left            =   120
               TabIndex        =   358
               Top             =   1680
               Width           =   2895
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   18
               Left            =   4440
               TabIndex        =   359
               Top             =   240
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   18
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   360
               Top             =   240
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   19
               Left            =   4440
               TabIndex        =   361
               Top             =   960
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   19
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   362
               Top             =   960
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   20
               Left            =   4440
               TabIndex        =   363
               Top             =   1680
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   20
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   364
               Top             =   1680
               Width           =   1215
            End
            Begin VB.CommandButton closedrugframes 
               Caption         =   "Close"
               Height          =   255
               Index           =   6
               Left            =   3120
               TabIndex        =   365
               Top             =   2160
               Width           =   1215
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Type Of Drug                                              Amount                 Measurement"
               ForeColor       =   &H0000FFFF&
               Height          =   255
               Index           =   6
               Left            =   105
               TabIndex        =   366
               Top             =   30
               Width           =   6615
            End
         End
         Begin VB.Frame sdrugframe 
            BackColor       =   &H00808000&
            BorderStyle     =   0  'None
            Height          =   2460
            Index           =   5
            Left            =   8040
            TabIndex        =   343
            Top             =   11760
            Visible         =   0   'False
            Width           =   6855
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   15
               Left            =   120
               TabIndex        =   344
               Top             =   240
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   16
               Left            =   120
               TabIndex        =   345
               Top             =   960
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   17
               Left            =   120
               TabIndex        =   346
               Top             =   1680
               Width           =   2895
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   15
               Left            =   4440
               TabIndex        =   347
               Top             =   240
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   15
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   348
               Top             =   240
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   16
               Left            =   4440
               TabIndex        =   349
               Top             =   960
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   16
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   350
               Top             =   960
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   17
               Left            =   4440
               TabIndex        =   351
               Top             =   1680
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   17
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   352
               Top             =   1680
               Width           =   1215
            End
            Begin VB.CommandButton closedrugframes 
               Caption         =   "Close"
               Height          =   255
               Index           =   5
               Left            =   3120
               TabIndex        =   353
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
               TabIndex        =   354
               Top             =   0
               Width           =   6615
            End
         End
         Begin VB.Frame sdrugframe 
            BackColor       =   &H00808000&
            BorderStyle     =   0  'None
            Height          =   2460
            Index           =   4
            Left            =   8160
            TabIndex        =   331
            Top             =   10800
            Visible         =   0   'False
            Width           =   6855
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   12
               Left            =   120
               TabIndex        =   332
               Top             =   240
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   13
               Left            =   120
               TabIndex        =   333
               Top             =   960
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   14
               Left            =   120
               TabIndex        =   334
               Top             =   1680
               Width           =   2895
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   12
               Left            =   4440
               TabIndex        =   335
               Top             =   240
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   12
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   336
               Top             =   240
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   13
               Left            =   4440
               TabIndex        =   337
               Top             =   960
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   13
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   338
               Top             =   960
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   14
               Left            =   4440
               TabIndex        =   339
               Top             =   1680
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   14
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   340
               Top             =   1680
               Width           =   1215
            End
            Begin VB.CommandButton closedrugframes 
               Caption         =   "Close"
               Height          =   255
               Index           =   4
               Left            =   3120
               TabIndex        =   341
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
               TabIndex        =   342
               Top             =   0
               Width           =   6615
            End
         End
         Begin VB.Frame sdrugframe 
            BackColor       =   &H00808000&
            BorderStyle     =   0  'None
            Caption         =   "s"
            Height          =   2460
            Index           =   3
            Left            =   7320
            TabIndex        =   495
            Top             =   11880
            Visible         =   0   'False
            Width           =   6855
            Begin VB.CommandButton closedrugframes 
               Caption         =   "Close"
               Height          =   255
               Index           =   3
               Left            =   3120
               TabIndex        =   505
               Top             =   2160
               Width           =   1215
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   9
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   504
               Top             =   240
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   9
               Left            =   4440
               TabIndex        =   503
               Top             =   240
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   10
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   502
               Top             =   960
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   10
               Left            =   4440
               TabIndex        =   501
               Top             =   960
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   11
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   500
               Top             =   1680
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   11
               Left            =   4440
               TabIndex        =   499
               Top             =   1680
               Width           =   2295
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   9
               Left            =   120
               TabIndex        =   498
               Top             =   240
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   10
               Left            =   120
               TabIndex        =   497
               Top             =   960
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   11
               Left            =   120
               TabIndex        =   496
               Top             =   1680
               Width           =   2895
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Type Of Drug                                              Amount                 Measurement"
               ForeColor       =   &H0000FFFF&
               Height          =   255
               Index           =   3
               Left            =   90
               TabIndex        =   506
               Top             =   0
               Width           =   6615
            End
         End
         Begin VB.Frame sdrugframe 
            BackColor       =   &H00808000&
            BorderStyle     =   0  'None
            Caption         =   "s"
            Height          =   2460
            Index           =   2
            Left            =   7440
            TabIndex        =   280
            Top             =   11760
            Visible         =   0   'False
            Width           =   6855
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   8
               Left            =   120
               TabIndex        =   281
               Top             =   1680
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   7
               Left            =   120
               TabIndex        =   282
               Top             =   960
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   6
               Left            =   120
               TabIndex        =   283
               Top             =   240
               Width           =   2895
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   8
               Left            =   4440
               TabIndex        =   284
               Top             =   1680
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   8
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   285
               Top             =   1680
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   7
               Left            =   4440
               TabIndex        =   286
               Top             =   960
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   7
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   287
               Top             =   960
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   6
               Left            =   4440
               TabIndex        =   288
               Top             =   240
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   6
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   289
               Top             =   240
               Width           =   1215
            End
            Begin VB.CommandButton closedrugframes 
               Caption         =   "Close"
               Height          =   255
               Index           =   2
               Left            =   3120
               TabIndex        =   290
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
               TabIndex        =   291
               Top             =   0
               Width           =   6615
            End
         End
         Begin VB.Frame sdrugframe 
            BackColor       =   &H00808000&
            BorderStyle     =   0  'None
            Caption         =   "s"
            Height          =   2460
            Index           =   1
            Left            =   7560
            TabIndex        =   241
            Top             =   11400
            Visible         =   0   'False
            Width           =   6855
            Begin VB.CommandButton closedrugframes 
               Caption         =   "Close"
               Height          =   255
               Index           =   1
               Left            =   3120
               TabIndex        =   251
               Top             =   2160
               Width           =   1215
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   3
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   250
               Top             =   240
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   3
               Left            =   4440
               TabIndex        =   249
               Top             =   240
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   4
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   248
               Top             =   960
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   4
               Left            =   4440
               TabIndex        =   247
               Top             =   960
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   5
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   246
               Top             =   1680
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   5
               Left            =   4440
               TabIndex        =   245
               Top             =   1680
               Width           =   2295
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   3
               Left            =   120
               TabIndex        =   244
               Top             =   240
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   4
               Left            =   120
               TabIndex        =   243
               Top             =   960
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   5
               Left            =   120
               TabIndex        =   242
               Top             =   1680
               Width           =   2895
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Type Of Drug                                              Amount                 Measurement"
               ForeColor       =   &H0000FFFF&
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   252
               Top             =   0
               Width           =   6615
            End
         End
         Begin VB.Frame sdrugframe 
            BackColor       =   &H00808000&
            BorderStyle     =   0  'None
            Height          =   2460
            Index           =   0
            Left            =   7200
            TabIndex        =   253
            Top             =   11880
            Visible         =   0   'False
            Width           =   6855
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   2
               Left            =   120
               TabIndex        =   263
               Top             =   1680
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   1
               Left            =   120
               TabIndex        =   261
               Top             =   960
               Width           =   2895
            End
            Begin VB.ListBox drugtype 
               Height          =   645
               Index           =   0
               Left            =   120
               TabIndex        =   260
               Top             =   240
               Width           =   2895
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   2
               Left            =   4440
               TabIndex        =   259
               Top             =   1680
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   2
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   258
               Top             =   1680
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   1
               Left            =   4440
               TabIndex        =   257
               Top             =   960
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   1
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   256
               Top             =   960
               Width           =   1215
            End
            Begin VB.ListBox drugmeasurement 
               Height          =   645
               Index           =   0
               Left            =   4440
               TabIndex        =   255
               Top             =   240
               Width           =   2295
            End
            Begin VB.TextBox drugamt 
               Height          =   285
               Index           =   0
               Left            =   3120
               MaxLength       =   12
               TabIndex        =   262
               Top             =   240
               Width           =   1215
            End
            Begin VB.CommandButton closedrugframes 
               Caption         =   "Close"
               Height          =   255
               Index           =   0
               Left            =   3120
               TabIndex        =   254
               Top             =   2160
               Width           =   1215
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Type Of Drug                                              Amount                 Measurement"
               ForeColor       =   &H0000FFFF&
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   264
               Top             =   0
               Width           =   6615
            End
         End
         Begin VB.Frame relationshipframe 
            BackColor       =   &H00808000&
            Caption         =   "Relationship to Subject Number:"
            ForeColor       =   &H000000FF&
            Height          =   3375
            Index           =   0
            Left            =   1000
            TabIndex        =   265
            Top             =   1000
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
               Index           =   0
               Left            =   150
               Style           =   1  'Checkbox
               TabIndex        =   266
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
               Index           =   1
               Left            =   150
               Style           =   1  'Checkbox
               TabIndex        =   267
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
               Index           =   2
               Left            =   150
               Style           =   1  'Checkbox
               TabIndex        =   268
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
               Index           =   9
               Left            =   2075
               Style           =   1  'Checkbox
               TabIndex        =   275
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
               Index           =   8
               Left            =   2075
               Style           =   1  'Checkbox
               TabIndex        =   274
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
               Index           =   7
               Left            =   2075
               Style           =   1  'Checkbox
               TabIndex        =   273
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
               Index           =   6
               Left            =   2075
               Style           =   1  'Checkbox
               TabIndex        =   272
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
               Index           =   5
               Left            =   2075
               Style           =   1  'Checkbox
               TabIndex        =   271
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
               Index           =   4
               Left            =   150
               Style           =   1  'Checkbox
               TabIndex        =   270
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
               Index           =   3
               Left            =   150
               Style           =   1  'Checkbox
               TabIndex        =   269
               Top             =   1825
               Width           =   1800
            End
            Begin VB.CommandButton Command8 
               Caption         =   "Close"
               Height          =   315
               Index           =   0
               Left            =   1080
               TabIndex        =   276
               Top             =   3000
               Width           =   1815
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   $"sinciden.frx":6F7D
               ForeColor       =   &H0000FFFF&
               Height          =   2895
               Index           =   0
               Left            =   0
               TabIndex        =   403
               Top             =   165
               Width           =   2175
            End
         End
         Begin VB.Frame pinfoframe 
            BackColor       =   &H00808000&
            Caption         =   "Property Info"
            Height          =   3900
            Index           =   5
            Left            =   11280
            TabIndex        =   484
            Top             =   2160
            Visible         =   0   'False
            Width           =   4325
            Begin VB.CommandButton Command22 
               Caption         =   "Clear"
               Height          =   270
               Index           =   5
               Left            =   3420
               TabIndex        =   490
               Top             =   3525
               Width           =   645
            End
            Begin VB.CommandButton Command10 
               Caption         =   "C   L   O   S   E"
               Height          =   300
               Index           =   5
               Left            =   60
               TabIndex        =   489
               Top             =   3495
               Width           =   3060
            End
            Begin VB.ListBox minorlist 
               Height          =   645
               Index           =   5
               Left            =   915
               TabIndex        =   488
               Top             =   2700
               Width           =   3105
            End
            Begin VB.ListBox majorlist 
               Height          =   645
               Index           =   5
               Left            =   915
               TabIndex        =   487
               Top             =   1890
               Width           =   3105
            End
            Begin VB.ListBox group 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   645
               Index           =   5
               Left            =   915
               TabIndex        =   486
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
               TabIndex        =   485
               Top             =   270
               Width           =   3105
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Minor Grouping"
               Height          =   450
               Index           =   5
               Left            =   120
               TabIndex        =   494
               Top             =   2730
               Width           =   810
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Major Grouping"
               Height          =   450
               Index           =   5
               Left            =   105
               TabIndex        =   493
               Top             =   1920
               Width           =   810
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Group"
               Height          =   270
               Index           =   5
               Left            =   120
               TabIndex        =   492
               Top             =   1080
               Width           =   810
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "UCR"
               Height          =   270
               Index           =   5
               Left            =   120
               TabIndex        =   491
               Top             =   240
               Width           =   810
            End
         End
         Begin VB.Frame lstframe 
            Caption         =   "Spelling Suggestions"
            Height          =   2415
            Left            =   5640
            TabIndex        =   417
            Top             =   2040
            Visible         =   0   'False
            Width           =   4935
            Begin VB.CommandButton Command18 
               Caption         =   "Skip"
               Height          =   375
               Left            =   3720
               TabIndex        =   421
               Top             =   960
               Width           =   1095
            End
            Begin VB.CommandButton Command17 
               Caption         =   "Change"
               Height          =   375
               Left            =   3720
               TabIndex        =   420
               Top             =   480
               Width           =   1095
            End
            Begin VB.CommandButton Command16 
               Caption         =   "Close"
               Height          =   375
               Left            =   3720
               TabIndex        =   419
               Top             =   1920
               Width           =   1095
            End
            Begin VB.ListBox lstsuggestions 
               Height          =   1815
               Left            =   120
               TabIndex        =   418
               Top             =   480
               Width           =   3495
            End
            Begin VB.Label checkword 
               Height          =   255
               Left            =   240
               TabIndex        =   422
               Top             =   240
               Width           =   3495
            End
         End
         Begin VB.Frame pinfoframe 
            BackColor       =   &H00808000&
            Caption         =   "Property Info"
            Height          =   3900
            Index           =   1
            Left            =   10920
            TabIndex        =   451
            Top             =   1200
            Visible         =   0   'False
            Width           =   4325
            Begin VB.CommandButton Command22 
               Caption         =   "Clear"
               Height          =   270
               Index           =   1
               Left            =   3420
               TabIndex        =   457
               Top             =   3525
               Width           =   645
            End
            Begin VB.CommandButton Command10 
               Caption         =   "C   L   O   S   E"
               Height          =   300
               Index           =   1
               Left            =   60
               TabIndex        =   456
               Top             =   3495
               Width           =   3060
            End
            Begin VB.ListBox minorlist 
               Height          =   645
               Index           =   1
               Left            =   915
               TabIndex        =   455
               Top             =   2700
               Width           =   3105
            End
            Begin VB.ListBox majorlist 
               Height          =   645
               Index           =   1
               Left            =   915
               TabIndex        =   454
               Top             =   1890
               Width           =   3105
            End
            Begin VB.ListBox group 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   645
               Index           =   1
               Left            =   915
               TabIndex        =   453
               Top             =   1080
               Width           =   3105
            End
            Begin VB.ListBox pucrlist 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   645
               Index           =   1
               Left            =   915
               Sorted          =   -1  'True
               TabIndex        =   452
               Top             =   270
               Width           =   3105
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Minor Grouping"
               Height          =   450
               Index           =   1
               Left            =   120
               TabIndex        =   461
               Top             =   2730
               Width           =   810
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Major Grouping"
               Height          =   450
               Index           =   1
               Left            =   105
               TabIndex        =   460
               Top             =   1920
               Width           =   810
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Group"
               Height          =   270
               Index           =   1
               Left            =   120
               TabIndex        =   459
               Top             =   1080
               Width           =   810
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "UCR"
               Height          =   270
               Index           =   1
               Left            =   120
               TabIndex        =   458
               Top             =   240
               Width           =   810
            End
         End
         Begin VB.Frame pinfoframe 
            BackColor       =   &H00808000&
            Caption         =   "Property Info"
            Height          =   3900
            Index           =   0
            Left            =   6360
            TabIndex        =   440
            Top             =   10200
            Visible         =   0   'False
            Width           =   4325
            Begin VB.CommandButton Command22 
               Caption         =   "Clear"
               Height          =   270
               Index           =   0
               Left            =   3420
               TabIndex        =   446
               Top             =   3525
               Width           =   645
            End
            Begin VB.CommandButton Command10 
               Caption         =   "C   L   O   S   E"
               Height          =   300
               Index           =   0
               Left            =   60
               TabIndex        =   445
               Top             =   3495
               Width           =   3060
            End
            Begin VB.ListBox minorlist 
               Height          =   645
               Index           =   0
               Left            =   915
               TabIndex        =   444
               Top             =   2700
               Width           =   3105
            End
            Begin VB.ListBox majorlist 
               Height          =   645
               Index           =   0
               Left            =   915
               TabIndex        =   443
               Top             =   1890
               Width           =   3105
            End
            Begin VB.ListBox group 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   645
               Index           =   0
               Left            =   915
               TabIndex        =   442
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
               TabIndex        =   441
               Top             =   270
               Width           =   3105
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Minor Grouping"
               Height          =   450
               Index           =   0
               Left            =   120
               TabIndex        =   450
               Top             =   2730
               Width           =   810
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Major Grouping"
               Height          =   450
               Index           =   0
               Left            =   105
               TabIndex        =   449
               Top             =   1920
               Width           =   810
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Group"
               Height          =   270
               Index           =   0
               Left            =   120
               TabIndex        =   448
               Top             =   1080
               Width           =   810
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "UCR"
               Height          =   270
               Index           =   0
               Left            =   120
               TabIndex        =   447
               Top             =   240
               Width           =   810
            End
         End
         Begin VB.Frame vucrf 
            BackColor       =   &H00808000&
            Caption         =   "Victim Connected To These UCR's"
            ForeColor       =   &H00FFFFFF&
            Height          =   2055
            Index           =   0
            Left            =   1000
            TabIndex        =   277
            Top             =   5000
            Visible         =   0   'False
            Width           =   3375
            Begin VB.CommandButton closevucrf 
               Caption         =   "Close"
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   279
               Top             =   1680
               Width           =   3135
            End
            Begin MSComctlLib.ListView vucrlist 
               Height          =   1335
               Index           =   0
               Left            =   120
               TabIndex        =   278
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
                  Object.Width           =   4233
               EndProperty
            End
         End
         Begin VB.Frame vucrf 
            BackColor       =   &H00808000&
            Caption         =   "Victim Connected To These UCR's"
            ForeColor       =   &H00FFFFFF&
            Height          =   2055
            Index           =   1
            Left            =   2130
            TabIndex        =   238
            Top             =   780
            Visible         =   0   'False
            Width           =   3375
            Begin VB.CommandButton closevucrf 
               Caption         =   "Close"
               Height          =   315
               Index           =   1
               Left            =   120
               TabIndex        =   239
               Top             =   1680
               Width           =   3135
            End
            Begin MSComctlLib.ListView vucrlist 
               Height          =   1335
               Index           =   1
               Left            =   120
               TabIndex        =   240
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
                  Object.Width           =   4233
               EndProperty
            End
         End
         Begin VB.Frame noframe 
            BackColor       =   &H00808000&
            Caption         =   "NARRATIVE ONLY - DETAIL FIELDS NOT ACCESSIBLE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   6735
            Index           =   1
            Left            =   11520
            TabIndex        =   513
            Top             =   12570
            Visible         =   0   'False
            Width           =   11340
         End
         Begin VB.Frame noframe 
            BackColor       =   &H00808000&
            Caption         =   "NARRATIVE ONLY - DETAIL FIELDS NOT ACCESSIBLE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   7935
            Index           =   0
            Left            =   11340
            TabIndex        =   512
            Top             =   4155
            Visible         =   0   'False
            Width           =   11340
         End
         Begin VB.CheckBox narrativeonly 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Narrative Only"
            Height          =   270
            Left            =   495
            TabIndex        =   0
            Top             =   870
            Width           =   1755
         End
         Begin VB.Frame Frame23 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   7110
            TabIndex        =   511
            Top             =   19350
            Width           =   975
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
               TabIndex        =   226
               Top             =   0
               Width           =   510
            End
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
               Left            =   15
               TabIndex        =   225
               Top             =   0
               Width           =   480
            End
         End
         Begin VB.Frame Frame17 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   3105
            TabIndex        =   510
            Top             =   18165
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
               TabIndex        =   210
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
               TabIndex        =   211
               Top             =   0
               Width           =   510
            End
         End
         Begin VB.Frame Frame19 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   900
            TabIndex        =   509
            Top             =   18165
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
               Left            =   15
               TabIndex        =   208
               Top             =   -15
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
               Left            =   735
               TabIndex        =   209
               Top             =   0
               Width           =   510
            End
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
            Left            =   9660
            TabIndex        =   214
            Top             =   17925
            Width           =   1755
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
            Left            =   9660
            TabIndex        =   215
            Top             =   18180
            Width           =   1905
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
            Left            =   7305
            TabIndex        =   213
            Top             =   18180
            Width           =   1905
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
            Left            =   7305
            TabIndex        =   212
            Top             =   17925
            Width           =   1755
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
            Left            =   8155
            Sorted          =   -1  'True
            TabIndex        =   227
            Top             =   19350
            Width           =   1845
         End
         Begin VB.TextBox FOLLOWUPOFFICERUNIT 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   10950
            TabIndex        =   229
            Top             =   19350
            Width           =   660
         End
         Begin VB.TextBox FOLLOWUPOFFICERDATE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   9990
            TabIndex        =   228
            Top             =   19365
            Width           =   960
         End
         Begin VB.Frame MSFRAME 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   650
            Index           =   1
            Left            =   4335
            TabIndex        =   508
            Top             =   6015
            Width           =   750
            Begin VB.Image MUGSHOT 
               BorderStyle     =   1  'Fixed Single
               Height          =   645
               Index           =   1
               Left            =   0
               Stretch         =   -1  'True
               Top             =   -15
               Width           =   750
            End
         End
         Begin VB.Frame MSFRAME 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   650
            Index           =   0
            Left            =   4320
            TabIndex        =   507
            Top             =   1995
            Width           =   750
            Begin VB.Image MUGSHOT 
               BorderStyle     =   1  'Fixed Single
               Height          =   645
               Index           =   0
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   750
            End
         End
         Begin VB.Frame pinfoframe 
            BackColor       =   &H00808000&
            Caption         =   "Property Info"
            Height          =   3900
            Index           =   4
            Left            =   11130
            TabIndex        =   473
            Top             =   4185
            Visible         =   0   'False
            Width           =   4325
            Begin VB.CommandButton Command22 
               Caption         =   "Clear"
               Height          =   270
               Index           =   4
               Left            =   3420
               TabIndex        =   479
               Top             =   3525
               Width           =   645
            End
            Begin VB.CommandButton Command10 
               Caption         =   "C   L   O   S   E"
               Height          =   300
               Index           =   4
               Left            =   60
               TabIndex        =   478
               Top             =   3495
               Width           =   3060
            End
            Begin VB.ListBox minorlist 
               Height          =   645
               Index           =   4
               Left            =   915
               TabIndex        =   477
               Top             =   2700
               Width           =   3105
            End
            Begin VB.ListBox majorlist 
               Height          =   645
               Index           =   4
               Left            =   915
               TabIndex        =   476
               Top             =   1890
               Width           =   3105
            End
            Begin VB.ListBox group 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   645
               Index           =   4
               Left            =   915
               TabIndex        =   475
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
               TabIndex        =   474
               Top             =   270
               Width           =   3105
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Minor Grouping"
               Height          =   450
               Index           =   4
               Left            =   120
               TabIndex        =   483
               Top             =   2730
               Width           =   810
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Major Grouping"
               Height          =   450
               Index           =   4
               Left            =   105
               TabIndex        =   482
               Top             =   1920
               Width           =   810
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Group"
               Height          =   270
               Index           =   4
               Left            =   120
               TabIndex        =   481
               Top             =   1080
               Width           =   810
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "UCR"
               Height          =   270
               Index           =   4
               Left            =   120
               TabIndex        =   480
               Top             =   240
               Width           =   810
            End
         End
         Begin VB.Frame pinfoframe 
            BackColor       =   &H00808000&
            Caption         =   "Property Info"
            Height          =   3900
            Index           =   2
            Left            =   11085
            TabIndex        =   462
            Top             =   8595
            Visible         =   0   'False
            Width           =   4325
            Begin VB.CommandButton Command22 
               Caption         =   "Clear"
               Height          =   270
               Index           =   2
               Left            =   3420
               TabIndex        =   468
               Top             =   3525
               Width           =   645
            End
            Begin VB.CommandButton Command10 
               Caption         =   "C   L   O   S   E"
               Height          =   300
               Index           =   2
               Left            =   60
               TabIndex        =   467
               Top             =   3495
               Width           =   3060
            End
            Begin VB.ListBox minorlist 
               Height          =   645
               Index           =   2
               Left            =   915
               TabIndex        =   466
               Top             =   2700
               Width           =   3105
            End
            Begin VB.ListBox majorlist 
               Height          =   645
               Index           =   2
               Left            =   915
               TabIndex        =   465
               Top             =   1890
               Width           =   3105
            End
            Begin VB.ListBox group 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   645
               Index           =   2
               Left            =   915
               TabIndex        =   464
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
               TabIndex        =   463
               Top             =   270
               Width           =   3105
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Minor Grouping"
               Height          =   450
               Index           =   2
               Left            =   120
               TabIndex        =   472
               Top             =   2730
               Width           =   810
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Major Grouping"
               Height          =   450
               Index           =   2
               Left            =   105
               TabIndex        =   471
               Top             =   1920
               Width           =   810
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Group"
               Height          =   270
               Index           =   2
               Left            =   120
               TabIndex        =   470
               Top             =   1080
               Width           =   810
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "UCR"
               Height          =   270
               Index           =   2
               Left            =   120
               TabIndex        =   469
               Top             =   240
               Width           =   810
            End
         End
         Begin VB.Frame pinfoframe 
            BackColor       =   &H00808000&
            Caption         =   "Property Info"
            Height          =   3900
            Index           =   3
            Left            =   11265
            TabIndex        =   429
            Top             =   6840
            Visible         =   0   'False
            Width           =   4325
            Begin VB.CommandButton Command22 
               Caption         =   "Clear"
               Height          =   270
               Index           =   3
               Left            =   3420
               TabIndex        =   435
               Top             =   3525
               Width           =   645
            End
            Begin VB.CommandButton Command10 
               Caption         =   "C   L   O   S   E"
               Height          =   300
               Index           =   3
               Left            =   60
               TabIndex        =   434
               Top             =   3495
               Width           =   3060
            End
            Begin VB.ListBox minorlist 
               Height          =   645
               Index           =   3
               Left            =   915
               TabIndex        =   433
               Top             =   2700
               Width           =   3105
            End
            Begin VB.ListBox majorlist 
               Height          =   645
               Index           =   3
               Left            =   915
               TabIndex        =   432
               Top             =   1890
               Width           =   3105
            End
            Begin VB.ListBox group 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   645
               Index           =   3
               Left            =   915
               TabIndex        =   431
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
               TabIndex        =   430
               Top             =   270
               Width           =   3105
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Minor Grouping"
               Height          =   450
               Index           =   3
               Left            =   120
               TabIndex        =   439
               Top             =   2730
               Width           =   810
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Major Grouping"
               Height          =   450
               Index           =   3
               Left            =   105
               TabIndex        =   438
               Top             =   1920
               Width           =   810
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Group"
               Height          =   270
               Index           =   3
               Left            =   120
               TabIndex        =   437
               Top             =   1080
               Width           =   810
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "UCR"
               Height          =   270
               Index           =   3
               Left            =   120
               TabIndex        =   436
               Top             =   240
               Width           =   810
            End
         End
         Begin VB.CommandButton PropertyCommand 
            Caption         =   "Property Info"
            Height          =   255
            Index           =   5
            Left            =   9233
            TabIndex        =   428
            Top             =   15090
            Width           =   1215
         End
         Begin VB.CommandButton PropertyCommand 
            Caption         =   "Property Info"
            Height          =   255
            Index           =   4
            Left            =   7713
            TabIndex        =   427
            Top             =   15090
            Width           =   1215
         End
         Begin VB.CommandButton PropertyCommand 
            Caption         =   "Property Info"
            Height          =   255
            Index           =   3
            Left            =   6193
            TabIndex        =   426
            Top             =   15090
            Width           =   1215
         End
         Begin VB.CommandButton PropertyCommand 
            Caption         =   "Property Info"
            Height          =   255
            Index           =   2
            Left            =   4673
            TabIndex        =   425
            Top             =   15090
            Width           =   1215
         End
         Begin VB.CommandButton PropertyCommand 
            Caption         =   "Property Info"
            Height          =   255
            Index           =   1
            Left            =   3135
            TabIndex        =   424
            Top             =   15090
            Width           =   1215
         End
         Begin VB.CommandButton PropertyCommand 
            Caption         =   "Property Info"
            Height          =   255
            Index           =   0
            Left            =   1693
            TabIndex        =   423
            Top             =   15090
            Width           =   1215
         End
         Begin VB.CommandButton Command15 
            Caption         =   "Check Spelling"
            Height          =   735
            Left            =   720
            TabIndex        =   127
            Top             =   11280
            Width           =   735
         End
         Begin Crystal.CrystalReport report 
            Left            =   2760
            Top             =   720
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin RichTextLib.RichTextBox narrative 
            Height          =   1095
            Left            =   675
            TabIndex        =   126
            Top             =   10140
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   1931
            _Version        =   393217
            ScrollBars      =   2
            TextRTF         =   $"sinciden.frx":7221
         End
         Begin VB.Frame relationshipframe 
            BackColor       =   &H00808000&
            Caption         =   "Relationship to Subject Number:"
            ForeColor       =   &H000000FF&
            Height          =   3375
            Index           =   1
            Left            =   1000
            TabIndex        =   404
            Top             =   1000
            Visible         =   0   'False
            Width           =   3975
            Begin VB.CommandButton Command8 
               Caption         =   "Close"
               Height          =   315
               Index           =   1
               Left            =   1080
               TabIndex        =   415
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
               Index           =   13
               Left            =   150
               Style           =   1  'Checkbox
               TabIndex        =   408
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
               Index           =   14
               Left            =   150
               Style           =   1  'Checkbox
               TabIndex        =   409
               Top             =   2400
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
               Left            =   2075
               Style           =   1  'Checkbox
               TabIndex        =   410
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
               Index           =   16
               Left            =   2075
               Style           =   1  'Checkbox
               TabIndex        =   411
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
               Index           =   17
               Left            =   2075
               Style           =   1  'Checkbox
               TabIndex        =   412
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
               Index           =   18
               Left            =   2075
               Style           =   1  'Checkbox
               TabIndex        =   413
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
               Index           =   19
               Left            =   2075
               Style           =   1  'Checkbox
               TabIndex        =   414
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
               Index           =   12
               Left            =   150
               Style           =   1  'Checkbox
               TabIndex        =   407
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
               Index           =   11
               Left            =   150
               Style           =   1  'Checkbox
               TabIndex        =   406
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
               Index           =   10
               Left            =   150
               Style           =   1  'Checkbox
               TabIndex        =   405
               Top             =   175
               Width           =   1800
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   $"sinciden.frx":72A3
               ForeColor       =   &H0000FFFF&
               Height          =   2895
               Index           =   1
               Left            =   0
               TabIndex        =   416
               Top             =   150
               Width           =   2175
            End
         End
         Begin VB.TextBox reportingofficeRunit 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   5425
            TabIndex        =   218
            Top             =   19035
            Width           =   600
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
            Left            =   480
            Sorted          =   -1  'True
            TabIndex        =   216
            Top             =   19035
            Width           =   3945
         End
         Begin VB.TextBox REPORTINGOFFICERDATE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   4450
            TabIndex        =   217
            Top             =   19035
            Width           =   960
         End
         Begin VB.TextBox reportingofficeRunit 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   5425
            TabIndex        =   221
            Top             =   19365
            Width           =   600
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
            Left            =   480
            Sorted          =   -1  'True
            TabIndex        =   219
            Top             =   19380
            Width           =   3945
         End
         Begin VB.TextBox REPORTINGOFFICERDATE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   4450
            TabIndex        =   220
            Top             =   19365
            Width           =   960
         End
         Begin VB.TextBox approvingofficeRunit 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   10950
            TabIndex        =   224
            Top             =   19035
            Width           =   600
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
            Left            =   6055
            Sorted          =   -1  'True
            TabIndex        =   222
            Top             =   19035
            Width           =   3945
         End
         Begin VB.TextBox APPROVINGOFFICERDATE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   9990
            TabIndex        =   223
            Top             =   19035
            Width           =   960
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
            Height          =   315
            Index           =   3
            Left            =   7800
            Style           =   1  'Graphical
            TabIndex        =   292
            Top             =   9675
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
            Height          =   285
            Index           =   2
            Left            =   8520
            Style           =   1  'Graphical
            TabIndex        =   293
            Top             =   9120
            Width           =   495
         End
         Begin VB.TextBox MISC 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   5760
            TabIndex        =   159
            Top             =   14800
            Width           =   5655
         End
         Begin VB.TextBox ISSUER 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   8200
            TabIndex        =   157
            Top             =   14400
            Width           =   1150
         End
         Begin VB.TextBox NIC 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4560
            TabIndex        =   155
            Top             =   14400
            Width           =   2295
         End
         Begin VB.TextBox BRANDNAME 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   9240
            TabIndex        =   153
            Top             =   13800
            Width           =   1095
         End
         Begin VB.TextBox STYLE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   6840
            TabIndex        =   151
            Top             =   13800
            Width           =   1095
         End
         Begin VB.TextBox stype 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   10050
            TabIndex        =   149
            Top             =   13200
            Width           =   1455
         End
         Begin VB.TextBox YEARN 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   7920
            TabIndex        =   147
            Top             =   13200
            Width           =   975
         End
         Begin VB.TextBox YEAREXP 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   6360
            TabIndex        =   146
            Top             =   13200
            Width           =   1455
         End
         Begin VB.TextBox SERIALSTATE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   10800
            TabIndex        =   143
            Top             =   12600
            Width           =   735
         End
         Begin VB.TextBox SECURITIESDATE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   9450
            TabIndex        =   158
            Top             =   14400
            Width           =   2055
         End
         Begin VB.TextBox DENOMINATION 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   6960
            TabIndex        =   156
            Top             =   14400
            Width           =   1200
         End
         Begin VB.TextBox CALIBER 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   10515
            TabIndex        =   154
            Top             =   13800
            Width           =   975
         End
         Begin VB.TextBox scolor 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   8040
            TabIndex        =   152
            Top             =   13800
            Width           =   1095
         End
         Begin VB.TextBox MODEL 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4560
            TabIndex        =   150
            Top             =   13800
            Width           =   2175
         End
         Begin VB.TextBox MAKE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   8880
            TabIndex        =   148
            Top             =   13200
            Width           =   1095
         End
         Begin VB.TextBox YEARREG 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4560
            TabIndex        =   144
            Top             =   13200
            Width           =   1695
         End
         Begin VB.TextBox SERIAL 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   7200
            TabIndex        =   142
            Top             =   12600
            Width           =   3135
         End
         Begin VB.TextBox HULL 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   7920
            TabIndex        =   141
            Top             =   12240
            Width           =   3495
         End
         Begin VB.TextBox VIN 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4560
            TabIndex        =   140
            Top             =   12255
            Width           =   3255
         End
         Begin VB.CheckBox SARTICLE 
            Caption         =   "Check1"
            Height          =   255
            Left            =   2100
            TabIndex        =   139
            Top             =   14760
            Width           =   225
         End
         Begin VB.CheckBox SSECURITIES 
            Caption         =   "Check1"
            Height          =   255
            Left            =   2100
            TabIndex        =   138
            Top             =   14450
            Width           =   225
         End
         Begin VB.CheckBox SLICENSE 
            Caption         =   "Check1"
            Height          =   255
            Left            =   2100
            TabIndex        =   137
            Top             =   14100
            Width           =   225
         End
         Begin VB.CheckBox SBOAT 
            Caption         =   "Check1"
            Height          =   255
            Left            =   2100
            TabIndex        =   136
            Top             =   13800
            Width           =   225
         End
         Begin VB.CheckBox SGUN 
            Caption         =   "Check1"
            Height          =   255
            Left            =   2100
            TabIndex        =   135
            Top             =   13480
            Width           =   225
         End
         Begin VB.CheckBox SVEHICLE 
            Caption         =   "Check1"
            Height          =   255
            Left            =   2100
            TabIndex        =   134
            Top             =   13200
            Width           =   225
         End
         Begin VB.CheckBox SVICTIM 
            Caption         =   "Check1"
            Height          =   255
            Left            =   840
            TabIndex        =   133
            Top             =   14760
            Width           =   225
         End
         Begin VB.CheckBox SSUSPECT 
            Caption         =   "Check1"
            Height          =   255
            Left            =   840
            TabIndex        =   132
            Top             =   14450
            Width           =   225
         End
         Begin VB.CheckBox STOWED 
            Caption         =   "Check1"
            Height          =   255
            Left            =   840
            TabIndex        =   131
            Top             =   14100
            Width           =   225
         End
         Begin VB.CheckBox SFOUND 
            Caption         =   "Check1"
            Height          =   255
            Left            =   840
            TabIndex        =   130
            Top             =   13800
            Width           =   225
         End
         Begin VB.CheckBox SRECOVERED 
            Caption         =   "Check1"
            Height          =   255
            Left            =   840
            TabIndex        =   129
            Top             =   13480
            Width           =   225
         End
         Begin VB.CheckBox SSTOLEN 
            Caption         =   "Check1"
            Height          =   255
            Left            =   840
            TabIndex        =   128
            Top             =   13170
            Width           =   225
         End
         Begin VB.TextBox numvehicle 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   0
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   294
            Top             =   15600
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox numvehicle 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   1
            Left            =   3480
            MaxLength       =   2
            TabIndex        =   295
            Top             =   15600
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox numvehicle 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   2
            Left            =   5160
            MaxLength       =   2
            TabIndex        =   296
            Top             =   15600
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox numvehicle 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   3
            Left            =   6720
            MaxLength       =   2
            TabIndex        =   297
            Top             =   15600
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox numvehicle 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   4
            Left            =   8160
            MaxLength       =   2
            TabIndex        =   298
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
            TabIndex        =   299
            Top             =   15600
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox DATERECOVERED 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   0
            Left            =   2160
            TabIndex        =   300
            Top             =   16440
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.TextBox DATERECOVERED 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   1
            Left            =   3840
            TabIndex        =   301
            Top             =   16440
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.TextBox DATERECOVERED 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   2
            Left            =   5520
            TabIndex        =   302
            Top             =   16440
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.TextBox DATERECOVERED 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   3
            Left            =   6960
            TabIndex        =   303
            Top             =   16440
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.TextBox DATERECOVERED 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   4
            Left            =   8640
            TabIndex        =   304
            Top             =   16440
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.TextBox DATERECOVERED 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   5
            Left            =   10080
            TabIndex        =   305
            Top             =   16440
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.TextBox numvehicle 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   6
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   306
            Top             =   16440
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox numvehicle 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   7
            Left            =   3720
            MaxLength       =   2
            TabIndex        =   307
            Top             =   16440
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox numvehicle 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   8
            Left            =   5400
            MaxLength       =   2
            TabIndex        =   308
            Top             =   16440
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox numvehicle 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   9
            Left            =   6840
            MaxLength       =   2
            TabIndex        =   309
            Top             =   16440
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox numvehicle 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   10
            Left            =   8520
            MaxLength       =   2
            TabIndex        =   310
            Top             =   16440
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox numvehicle 
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   11
            Left            =   9960
            MaxLength       =   2
            TabIndex        =   311
            Top             =   16440
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.ComboBox vsname 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   1800
            Sorted          =   -1  'True
            TabIndex        =   20
            Top             =   2640
            Width           =   3285
         End
         Begin VB.ComboBox vsname 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   1
            Left            =   1800
            Sorted          =   -1  'True
            TabIndex        =   80
            Top             =   6660
            Width           =   3285
         End
         Begin VB.TextBox zipcode 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   7320
            TabIndex        =   36
            Top             =   3960
            Width           =   780
         End
         Begin VB.ComboBox state 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   6600
            Sorted          =   -1  'True
            TabIndex        =   35
            Top             =   3960
            Width           =   645
         End
         Begin VB.TextBox address 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   1800
            TabIndex        =   33
            Top             =   3960
            Width           =   3015
         End
         Begin VB.ComboBox city 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   4800
            Sorted          =   -1  'True
            TabIndex        =   34
            Top             =   3960
            Width           =   1740
         End
         Begin VB.TextBox zipcode 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   1
            Left            =   7320
            TabIndex        =   96
            Top             =   8040
            Width           =   780
         End
         Begin VB.ComboBox state 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   1
            Left            =   6600
            Sorted          =   -1  'True
            TabIndex        =   95
            Top             =   8040
            Width           =   645
         End
         Begin VB.TextBox address 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   93
            Top             =   8040
            Width           =   3015
         End
         Begin VB.ComboBox city 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   1
            Left            =   4800
            Sorted          =   -1  'True
            TabIndex        =   94
            Top             =   8040
            Width           =   1740
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00808000&
            Caption         =   "VUCR"
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
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   2805
            Width           =   645
         End
         Begin VB.TextBox description 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   1650
            MaxLength       =   30
            MultiLine       =   -1  'True
            TabIndex        =   160
            Top             =   15360
            Width           =   1300
         End
         Begin VB.TextBox description 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   3000
            MaxLength       =   30
            MultiLine       =   -1  'True
            TabIndex        =   168
            Top             =   15360
            Width           =   1440
         End
         Begin VB.TextBox description 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   4560
            MaxLength       =   30
            MultiLine       =   -1  'True
            TabIndex        =   176
            Top             =   15360
            Width           =   1440
         End
         Begin VB.TextBox description 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   3
            Left            =   6080
            MaxLength       =   30
            MultiLine       =   -1  'True
            TabIndex        =   184
            Top             =   15360
            Width           =   1440
         End
         Begin VB.TextBox description 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   4
            Left            =   7600
            MaxLength       =   30
            MultiLine       =   -1  'True
            TabIndex        =   192
            Top             =   15360
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
            TabIndex        =   200
            Top             =   15360
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   1650
            TabIndex        =   161
            Top             =   15660
            Width           =   1300
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   3000
            TabIndex        =   169
            Top             =   15660
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   4560
            TabIndex        =   177
            Top             =   15660
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   3
            Left            =   6080
            TabIndex        =   185
            Top             =   15660
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   4
            Left            =   7600
            TabIndex        =   193
            Top             =   15660
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   5
            Left            =   9120
            TabIndex        =   201
            Top             =   15660
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   6
            Left            =   1650
            TabIndex        =   162
            Top             =   15960
            Width           =   1300
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   7
            Left            =   3000
            TabIndex        =   170
            Top             =   15960
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   8
            Left            =   4560
            TabIndex        =   178
            Top             =   15960
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   9
            Left            =   6080
            TabIndex        =   186
            Top             =   15960
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   10
            Left            =   7600
            TabIndex        =   194
            Top             =   15960
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   11
            Left            =   9120
            TabIndex        =   202
            Top             =   15960
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   12
            Left            =   1650
            TabIndex        =   163
            Top             =   16275
            Width           =   1300
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   13
            Left            =   3000
            TabIndex        =   171
            Top             =   16275
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   14
            Left            =   4560
            TabIndex        =   179
            Top             =   16275
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   15
            Left            =   6080
            TabIndex        =   187
            Top             =   16275
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   16
            Left            =   7600
            TabIndex        =   195
            Top             =   16275
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   17
            Left            =   9120
            TabIndex        =   203
            Top             =   16275
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   18
            Left            =   1650
            TabIndex        =   164
            Top             =   16590
            Width           =   1300
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   19
            Left            =   3000
            TabIndex        =   172
            Top             =   16590
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   20
            Left            =   4560
            TabIndex        =   180
            Top             =   16590
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   21
            Left            =   6080
            TabIndex        =   188
            Top             =   16590
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   22
            Left            =   7600
            TabIndex        =   196
            Top             =   16590
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   23
            Left            =   9120
            TabIndex        =   204
            Top             =   16590
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   24
            Left            =   1650
            TabIndex        =   165
            Top             =   16920
            Width           =   1300
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   25
            Left            =   3000
            TabIndex        =   173
            Top             =   16920
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   26
            Left            =   4560
            TabIndex        =   181
            Top             =   16920
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   27
            Left            =   6080
            TabIndex        =   189
            Top             =   16920
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   28
            Left            =   7600
            TabIndex        =   197
            Top             =   16920
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   29
            Left            =   9120
            TabIndex        =   205
            Top             =   16920
            Width           =   1440
         End
         Begin VB.Frame alcoholframe 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   0
            Left            =   7800
            TabIndex        =   312
            Top             =   4500
            Width           =   1215
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
               Left            =   720
               TabIndex        =   48
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
               Index           =   0
               Left            =   120
               TabIndex        =   47
               Top             =   0
               Width           =   480
            End
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
               Left            =   120
               TabIndex        =   49
               Top             =   240
               Value           =   -1  'True
               Width           =   915
            End
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   0
            Left            =   6720
            TabIndex        =   313
            Top             =   4890
            Width           =   1095
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
               Left            =   0
               TabIndex        =   52
               Top             =   240
               Value           =   -1  'True
               Width           =   885
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
               TabIndex        =   50
               Top             =   0
               Width           =   600
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
               Index           =   0
               Left            =   600
               TabIndex        =   51
               Top             =   0
               Width           =   510
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
            Index           =   0
            Left            =   9600
            TabIndex        =   67
            Top             =   5640
            Width           =   1000
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
            Height          =   285
            Index           =   0
            Left            =   8520
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   5115
            Width           =   495
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
            Index           =   0
            Left            =   10600
            TabIndex        =   27
            Top             =   2535
            Width           =   900
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
            Index           =   0
            Left            =   7250
            TabIndex        =   23
            Top             =   2535
            Width           =   925
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
            Index           =   0
            Left            =   8160
            TabIndex        =   24
            Top             =   2535
            Width           =   880
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
            Index           =   0
            Left            =   6310
            TabIndex        =   22
            Top             =   2535
            Width           =   925
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
            Left            =   6360
            TabIndex        =   82
            Top             =   6555
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
            Index           =   1
            Left            =   8220
            TabIndex        =   84
            Top             =   6555
            Width           =   855
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
            Left            =   7320
            TabIndex        =   83
            Top             =   6555
            Width           =   855
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
            Left            =   10680
            TabIndex        =   87
            Top             =   6555
            Width           =   930
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
            Index           =   0
            Left            =   5210
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   2340
            Width           =   995
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
            Left            =   5280
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   6360
            Width           =   985
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   35
            Left            =   9120
            TabIndex        =   206
            Top             =   17220
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   34
            Left            =   7600
            TabIndex        =   198
            Top             =   17220
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   33
            Left            =   6080
            TabIndex        =   190
            Top             =   17220
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   32
            Left            =   4560
            TabIndex        =   182
            Top             =   17220
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   31
            Left            =   3000
            TabIndex        =   174
            Top             =   17220
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   30
            Left            =   1650
            TabIndex        =   166
            Top             =   17220
            Width           =   1300
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   41
            Left            =   9120
            TabIndex        =   207
            Top             =   17520
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   40
            Left            =   7600
            TabIndex        =   199
            Top             =   17520
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   39
            Left            =   6080
            TabIndex        =   191
            Top             =   17520
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   38
            Left            =   4560
            TabIndex        =   183
            Top             =   17520
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   37
            Left            =   3000
            TabIndex        =   175
            Top             =   17520
            Width           =   1440
         End
         Begin VB.TextBox totalvalue 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   36
            Left            =   1650
            TabIndex        =   167
            Top             =   17520
            Width           =   1300
         End
         Begin VB.TextBox HOMEDAYPHONE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   0
            Left            =   9060
            TabIndex        =   38
            Top             =   3890
            Width           =   1020
         End
         Begin VB.TextBox WORKDAYPHONE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   0
            Left            =   9060
            TabIndex        =   39
            Top             =   4200
            Width           =   1020
         End
         Begin VB.TextBox HOMENIGHTPHONE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   0
            Left            =   10320
            TabIndex        =   40
            Top             =   3890
            Width           =   1020
         End
         Begin VB.TextBox WORKNIGHTPHONE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   0
            Left            =   10320
            TabIndex        =   41
            Top             =   4200
            Width           =   1020
         End
         Begin VB.TextBox LOCATIONNUMBER 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   8160
            TabIndex        =   37
            Top             =   3960
            Width           =   840
         End
         Begin VB.TextBox LOCATIONNUMBER 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   8160
            TabIndex        =   97
            Top             =   8040
            Width           =   840
         End
         Begin VB.TextBox HOMEDAYPHONE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   1
            Left            =   9120
            TabIndex        =   98
            Top             =   7850
            Width           =   1020
         End
         Begin VB.TextBox HOMENIGHTPHONE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   1
            Left            =   10320
            TabIndex        =   100
            Top             =   7850
            Width           =   1020
         End
         Begin VB.TextBox WORKNIGHTPHONE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   10320
            TabIndex        =   101
            Top             =   8160
            Width           =   1020
         End
         Begin VB.TextBox WORKDAYPHONE 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   1
            Left            =   9120
            TabIndex        =   99
            Top             =   8160
            Width           =   1020
         End
         Begin VB.TextBox ht 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   1680
            TabIndex        =   28
            Top             =   3240
            Width           =   785
         End
         Begin VB.TextBox weight 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   2520
            TabIndex        =   29
            Top             =   3240
            Width           =   785
         End
         Begin VB.TextBox hair 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   3240
            TabIndex        =   30
            Top             =   3240
            Width           =   785
         End
         Begin VB.TextBox eyes 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   4050
            TabIndex        =   31
            Top             =   3240
            Width           =   785
         End
         Begin VB.TextBox peculiarities 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   4920
            TabIndex        =   32
            Top             =   3240
            Width           =   6585
         End
         Begin VB.TextBox BIRTHDATE 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   9720
            TabIndex        =   26
            Top             =   2670
            Width           =   900
         End
         Begin VB.TextBox ht 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   88
            Top             =   7200
            Width           =   780
         End
         Begin VB.TextBox weight 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   2480
            TabIndex        =   89
            Top             =   7200
            Width           =   780
         End
         Begin VB.TextBox hair 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   3240
            TabIndex        =   90
            Top             =   7200
            Width           =   780
         End
         Begin VB.TextBox eyes 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   4080
            TabIndex        =   91
            Top             =   7200
            Width           =   780
         End
         Begin VB.TextBox peculiarities 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   4920
            TabIndex        =   92
            Top             =   7200
            Width           =   6585
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   0
            Left            =   3000
            TabIndex        =   314
            Top             =   4680
            Width           =   1160
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
               Index           =   0
               Left            =   0
               TabIndex        =   42
               Top             =   0
               Width           =   480
            End
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
               Index           =   0
               Left            =   480
               TabIndex        =   43
               Top             =   0
               Value           =   -1  'True
               Width           =   615
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   0
            Left            =   4560
            TabIndex        =   315
            Top             =   5100
            Width           =   1050
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
               Index           =   0
               Left            =   550
               TabIndex        =   46
               Top             =   0
               Value           =   -1  'True
               Width           =   495
            End
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
               Index           =   0
               Left            =   0
               TabIndex        =   45
               Top             =   0
               Width           =   480
            End
         End
         Begin VB.CheckBox DETECTIVE 
            BackColor       =   &H00FFFFFF&
            Caption         =   "DETECTIVE/SPL"
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
            Left            =   9240
            TabIndex        =   56
            Top             =   4950
            Width           =   1245
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
            Index           =   0
            Left            =   9240
            TabIndex        =   54
            Top             =   4510
            Width           =   1200
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
            Index           =   0
            Left            =   9240
            TabIndex        =   55
            Top             =   4720
            Width           =   1215
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
            Index           =   0
            Left            =   9240
            TabIndex        =   57
            Top             =   5160
            Width           =   765
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
            Index           =   0
            Left            =   10620
            TabIndex        =   59
            Top             =   5040
            Width           =   865
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
            Index           =   0
            Left            =   10620
            TabIndex        =   58
            Top             =   4680
            Width           =   720
         End
         Begin VB.Frame alcoholframe 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   1
            Left            =   4440
            TabIndex        =   316
            Top             =   5450
            Width           =   1575
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
               TabIndex        =   62
               Top             =   0
               Value           =   -1  'True
               Width           =   555
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
               TabIndex        =   60
               Top             =   0
               Width           =   480
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
               Index           =   1
               Left            =   520
               TabIndex        =   61
               Top             =   0
               Width           =   510
            End
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   1
            Left            =   4440
            TabIndex        =   317
            Top             =   5750
            Width           =   1845
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
               TabIndex        =   64
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
               Index           =   1
               Left            =   0
               TabIndex        =   63
               Top             =   0
               Width           =   585
            End
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
               TabIndex        =   65
               Top             =   0
               Value           =   -1  'True
               Width           =   585
            End
         End
         Begin VB.TextBox age 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   9135
            TabIndex        =   25
            Top             =   2670
            Width           =   525
         End
         Begin VB.TextBox age 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   9120
            TabIndex        =   85
            Top             =   6690
            Width           =   525
         End
         Begin VB.CheckBox ORiGINAL 
            Caption         =   "Check1"
            Height          =   255
            Left            =   480
            TabIndex        =   1
            Top             =   1320
            Width           =   225
         End
         Begin VB.CheckBox MODIFIES 
            Caption         =   "Check1"
            Height          =   255
            Left            =   480
            TabIndex        =   2
            Top             =   1680
            Width           =   225
         End
         Begin VB.CheckBox SUPPLEMENTAL 
            Caption         =   "Check1"
            Height          =   255
            Left            =   2400
            TabIndex        =   3
            Top             =   1320
            Width           =   225
         End
         Begin VB.CheckBox CASEst 
            Caption         =   "Check1"
            Height          =   255
            Left            =   2400
            TabIndex        =   4
            Top             =   1680
            Width           =   225
         End
         Begin VB.CheckBox ADDITIONALV 
            Caption         =   "Check1"
            Height          =   255
            Left            =   4920
            TabIndex        =   5
            Top             =   1320
            Width           =   225
         End
         Begin VB.CheckBox additionalo 
            Caption         =   "Check1"
            Height          =   255
            Left            =   4920
            TabIndex        =   6
            Top             =   1680
            Width           =   225
         End
         Begin VB.CheckBox ADDITIONALS 
            Caption         =   "Check1"
            Height          =   255
            Left            =   7320
            TabIndex        =   7
            Top             =   1320
            Width           =   225
         End
         Begin VB.CheckBox ADDITIONALR 
            Caption         =   "Check1"
            Height          =   255
            Left            =   7320
            TabIndex        =   8
            Top             =   1680
            Width           =   225
         End
         Begin VB.TextBox victim 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   1200
            TabIndex        =   10
            Top             =   2640
            Width           =   495
         End
         Begin VB.CheckBox complainant 
            Caption         =   "Check1"
            Height          =   255
            Index           =   0
            Left            =   1500
            TabIndex        =   9
            Top             =   2280
            Width           =   200
         End
         Begin VB.TextBox subject 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   1200
            TabIndex        =   12
            Top             =   3000
            Width           =   495
         End
         Begin VB.CheckBox runaway 
            Caption         =   "Check1"
            Height          =   255
            Index           =   0
            Left            =   1500
            TabIndex        =   13
            Top             =   3480
            Width           =   200
         End
         Begin VB.CheckBox wanted 
            Caption         =   "Check1"
            Height          =   255
            Index           =   0
            Left            =   1500
            TabIndex        =   14
            Top             =   3900
            Width           =   200
         End
         Begin VB.CheckBox warrant 
            Caption         =   "Check1"
            Height          =   255
            Index           =   0
            Left            =   1500
            TabIndex        =   15
            Top             =   4300
            Width           =   200
         End
         Begin VB.CheckBox arrest 
            Caption         =   "Check1"
            Height          =   255
            Index           =   0
            Left            =   1500
            TabIndex        =   16
            Top             =   4700
            Width           =   200
         End
         Begin VB.CheckBox jail 
            Caption         =   "Check1"
            Height          =   255
            Index           =   0
            Left            =   1500
            TabIndex        =   17
            Top             =   5100
            Width           =   200
         End
         Begin VB.CheckBox summons 
            Caption         =   "Check1"
            Height          =   255
            Index           =   0
            Left            =   1500
            TabIndex        =   18
            Top             =   5520
            Width           =   200
         End
         Begin VB.TextBox typeother 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   600
            MaxLength       =   15
            TabIndex        =   19
            Top             =   5760
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00808000&
            Caption         =   "VUCR"
            Height          =   255
            Index           =   1
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   6775
            Width           =   645
         End
         Begin VB.TextBox victim 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   1200
            TabIndex        =   69
            Top             =   6600
            Width           =   495
         End
         Begin VB.CheckBox complainant 
            Caption         =   "Check1"
            Height          =   255
            Index           =   1
            Left            =   1500
            TabIndex        =   68
            Top             =   6240
            Width           =   200
         End
         Begin VB.TextBox subject 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   1200
            TabIndex        =   71
            Top             =   6960
            Width           =   495
         End
         Begin VB.CheckBox runaway 
            Caption         =   "Check1"
            Height          =   255
            Index           =   1
            Left            =   1500
            TabIndex        =   73
            Top             =   7440
            Width           =   200
         End
         Begin VB.CheckBox wanted 
            Caption         =   "Check1"
            Height          =   255
            Index           =   1
            Left            =   1500
            TabIndex        =   74
            Top             =   7860
            Width           =   200
         End
         Begin VB.CheckBox warrant 
            Caption         =   "Check1"
            Height          =   255
            Index           =   1
            Left            =   1500
            TabIndex        =   75
            Top             =   8265
            Width           =   200
         End
         Begin VB.CheckBox arrest 
            Caption         =   "Check1"
            Height          =   255
            Index           =   1
            Left            =   1500
            TabIndex        =   76
            Top             =   8655
            Width           =   200
         End
         Begin VB.CheckBox jail 
            Caption         =   "Check1"
            Height          =   255
            Index           =   1
            Left            =   1500
            TabIndex        =   77
            Top             =   9060
            Width           =   200
         End
         Begin VB.CheckBox summons 
            Caption         =   "Check1"
            Height          =   255
            Index           =   1
            Left            =   1500
            TabIndex        =   78
            Top             =   9480
            Width           =   200
         End
         Begin VB.TextBox typeother 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   600
            TabIndex        =   79
            Top             =   9720
            Width           =   855
         End
         Begin VB.TextBox BIRTHDATE 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   9720
            TabIndex        =   86
            Top             =   6690
            Width           =   900
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
            Index           =   1
            Left            =   10620
            TabIndex        =   117
            Top             =   8685
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
            Index           =   1
            Left            =   10620
            TabIndex        =   118
            Top             =   9045
            Width           =   865
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
            Index           =   1
            Left            =   9240
            TabIndex        =   114
            Top             =   8730
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
            Index           =   1
            Left            =   9240
            TabIndex        =   113
            Top             =   8520
            Width           =   1200
         End
         Begin VB.CheckBox DETECTIVE 
            BackColor       =   &H00FFFFFF&
            Caption         =   "DETECTIVE/SPL"
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
            Left            =   9240
            TabIndex        =   115
            Top             =   8955
            Width           =   1245
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   1
            Left            =   4560
            TabIndex        =   237
            Top             =   9120
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
               Index           =   1
               Left            =   0
               TabIndex        =   105
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
               Index           =   1
               Left            =   550
               TabIndex        =   106
               Top             =   0
               Value           =   -1  'True
               Width           =   495
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   236
            Top             =   8640
            Width           =   1160
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
               Index           =   1
               Left            =   600
               TabIndex        =   103
               Top             =   0
               Value           =   -1  'True
               Width           =   615
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
               Index           =   1
               Left            =   0
               TabIndex        =   102
               Top             =   0
               Width           =   585
            End
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
            Height          =   315
            Index           =   1
            Left            =   7800
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   5640
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
            Index           =   1
            Left            =   9720
            TabIndex        =   125
            Top             =   9720
            Width           =   1000
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   2
            Left            =   6720
            TabIndex        =   235
            Top             =   8880
            Width           =   1095
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
               Index           =   2
               Left            =   600
               TabIndex        =   111
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
               Index           =   2
               Left            =   0
               TabIndex        =   110
               Top             =   0
               Width           =   480
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
               Index           =   2
               Left            =   0
               TabIndex        =   112
               Top             =   240
               Value           =   -1  'True
               Width           =   885
            End
         End
         Begin VB.Frame alcoholframe 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   2
            Left            =   7800
            TabIndex        =   234
            Top             =   8520
            Width           =   1215
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
               Index           =   2
               Left            =   120
               TabIndex        =   109
               Top             =   240
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
               Index           =   2
               Left            =   120
               TabIndex        =   107
               Top             =   0
               Width           =   480
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
               Index           =   2
               Left            =   720
               TabIndex        =   108
               Top             =   0
               Width           =   510
            End
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   3
            Left            =   4440
            TabIndex        =   233
            Top             =   9705
            Width           =   1845
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
               Index           =   3
               Left            =   1080
               TabIndex        =   124
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
               Index           =   3
               Left            =   0
               TabIndex        =   122
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
               Index           =   3
               Left            =   600
               TabIndex        =   123
               Top             =   0
               Width           =   510
            End
         End
         Begin VB.Frame alcoholframe 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   3
            Left            =   4440
            TabIndex        =   232
            Top             =   9405
            Width           =   1815
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
               Index           =   3
               Left            =   600
               TabIndex        =   120
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
               Index           =   3
               Left            =   0
               TabIndex        =   119
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
               Index           =   3
               Left            =   1080
               TabIndex        =   121
               Top             =   0
               Value           =   -1  'True
               Width           =   555
            End
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
            Index           =   1
            Left            =   9240
            TabIndex        =   116
            Top             =   9160
            Width           =   765
         End
         Begin MSComctlLib.ListView injury 
            Height          =   495
            Index           =   0
            Left            =   1680
            TabIndex        =   44
            Top             =   4920
            Width           =   2535
            _ExtentX        =   4471
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
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   3881
            EndProperty
         End
         Begin MSComctlLib.ListView injury 
            Height          =   495
            Index           =   1
            Left            =   1680
            TabIndex        =   104
            Top             =   8880
            Width           =   2535
            _ExtentX        =   4471
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
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   3881
            EndProperty
         End
         Begin VB.Label pgof 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   120
            TabIndex        =   514
            Top             =   50
            Width           =   1935
         End
         Begin VB.Label incidentnumber 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   375
            Left            =   9240
            TabIndex        =   329
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label PAGE 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   10680
            TabIndex        =   318
            Top             =   1680
            Width           =   375
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "PAGE"
            Height          =   255
            Left            =   10680
            TabIndex        =   319
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label TVSTOLEN 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   10680
            TabIndex        =   320
            Top             =   15660
            Width           =   855
         End
         Begin VB.Label TVDAMAGED 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   10680
            TabIndex        =   321
            Top             =   15960
            Width           =   855
         End
         Begin VB.Label TVBURNED 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   10680
            TabIndex        =   322
            Top             =   16275
            Width           =   855
         End
         Begin VB.Label TVRECOVERED 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   10680
            TabIndex        =   323
            Top             =   16590
            Width           =   855
         End
         Begin VB.Label TVSEIZED 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   10680
            TabIndex        =   324
            Top             =   16920
            Width           =   855
         End
         Begin VB.Label TVCOUNTERFEIT 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   10680
            TabIndex        =   325
            Top             =   17220
            Width           =   855
         End
         Begin VB.Label TVUNKNOWN 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   10680
            TabIndex        =   326
            Top             =   17520
            Width           =   855
         End
         Begin VB.Label orinumber 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   360
            TabIndex        =   327
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "TYPE:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   7320
            TabIndex        =   328
            Top             =   5760
            Width           =   615
         End
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   7215
      Left            =   11655
      TabIndex        =   230
      Top             =   630
      Width           =   250
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   145
      Top             =   0
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   1111
      ButtonWidth     =   1508
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "TempSave"
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
            Caption         =   "TempList"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "NextPage"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "PriorPage"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            ImageIndex      =   12
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
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sinciden.frx":7547
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sinciden.frx":799B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sinciden.frx":7DEF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sinciden.frx":8243
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sinciden.frx":8697
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sinciden.frx":8AEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sinciden.frx":8F3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sinciden.frx":9393
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sinciden.frx":97E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sinciden.frx":9C3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sinciden.frx":A08F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sinciden.frx":A4E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sinciden.frx":A937
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label INCIDENTDATE 
      Caption         =   "Label4"
      Height          =   135
      Left            =   360
      TabIndex        =   330
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "sinciden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fromfind, schanged As Integer, ecc As Integer, incidentfound As Boolean
Dim nametype As Integer
Dim holdrecv As Integer, HI As String
Dim FROMKEY As Integer, BACKTAB As Integer
Dim VUCRSEL(5, 1) As String
Dim tempsave As Integer, TV1, TV2 As String, tempword As String
Private Sub loadcodes()
Dim db As Database, rs As Recordset, itmx As ListItem
On Error GoTo oderror
od:
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("select * from codes")
On Error Resume Next
For t% = 0 To 29
    drugmeasurement(t%).clear
    drugtype(t%).clear
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
        Case "group"
            For t% = 0 To 5
                group(t%).AddItem rs("code")
            Next t%
        Case "drugtype"
            For t% = 0 To 29
                drugtype(t%).AddItem rs("code")
            Next t%
        Case "measure"
            For t% = 0 To 29
                drugmeasurement(t%).AddItem rs("code")
            Next t%
    End Select
    rs.MoveNext
Wend
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
On Error Resume Next
db.Close
Exit Sub
oderror2:
If Err > 3200 Then
    Resume od2
Else
    Resume Next
End If
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
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

Private Sub active_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
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
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
Set rs = db.OpenRecordset("SELECT * FROM PEOPLE WHERE DPnameLF = " + Chr$(34) + vsname(Index) + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    If Not IsNull(rs("DPHADDRESS")) Then
        address(Index) = rs("DPHADDRESS")
    Else
        address(Index) = ""
    End If
End If
db.Close
On Error Resume Next
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub


Private Sub age_Click(Index As Integer)

If Index > 0 Then
    
End If
age(Index).Refresh
End Sub


Private Sub age_LostFocus(Index As Integer)
If age(Index) = "" Then
    age(Index) = "00"
End If
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

Private Sub approvingofficer_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub approvingofficerdate_GotFocus()
If approvingofficer.ListIndex > -1 And APPROVINGOFFICERDATE = "" Then
    APPROVINGOFFICERDATE = Format$(Date$, "mm/dd/yyyy")
End If
End Sub

Private Sub approvingofficerdate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(APPROVINGOFFICERDATE) = 1 Or Len(APPROVINGOFFICERDATE) = 4 Then
    Call sendslash
End If
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

Private Sub ARREST_GotFocus(Index As Integer)
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If

End Sub

Private Sub arrested18andover_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub arrestedunder18_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub birthdate_Change(Index As Integer)
If IsDate(BIRTHDATE(Index)) Then
    If age(Index) > "" Then
        If fromfind = 0 Then
            msg = MsgBox("Age contradicts birthdate.  Would you like to automatically calculate age?", 4, "Genesis Information Log")
            If msg = 6 Then
                age(Index) = DateDiff("yyyy", CDate(BIRTHDATE(Index)), CDate(Date$))
            End If
        End If
    End If
End If
End Sub

Private Sub BIRTHDATE_LostFocus(Index As Integer)
If BIRTHDATE(Index) > "" And Not IsDate(BIRTHDATE(Index)) Then
    msg = MsgBox("Date/Time entered if invalid.", 48, "Genesis Error Log")
'---- setfocus logic ----
'             BIRTHDATE(index).SetFocus
          If BIRTHDATE(Index).Visible Then
              BIRTHDATE(Index).SetFocus
           End If
End If
BIRTHDATE(Index) = Format$(BIRTHDATE(Index), "mm/dd/yyyy")
End Sub

Private Sub BRANDNAME_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub CALIBER_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub city_GotFocus(Index As Integer)
If vsname(Index) = "UNKNOWN" Then
    city(Index) = ""
End If
    
End Sub
Private Sub closedrugframes_Click(Index As Integer)
sdrugframe(Index).Visible = False
If Index > 3 Then
'---- setfocus logic ----
'             totalvalue(index - 2).SetFocus
          If totalvalue(Index - 2).Visible Then
              totalvalue(Index - 2).SetFocus
           End If
End If

End Sub
Private Sub closevucrf_Click(Index As Integer)

vucrf(Index).Visible = False
'---- setfocus logic ----
'         complainant(index).SetFocus
          If complainant(Index).Visible Then
              complainant(Index).SetFocus
           End If
End Sub
Private Sub birthdate_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(BIRTHDATE(Index)) = 1 Or Len(BIRTHDATE(Index)) = 4 Then
    Call sendslash
End If
End If
End Sub
Private Sub Command1_Click(Index As Integer)
If fromfind = 1 Then
    Exit Sub
End If
For r% = 1 To vucrlist(Index).ListItems.Count
    vucrlist(Index).ListItems(r%).Selected = False
Next r%
For r% = 1 To vucrlist(Index).ListItems.Count
    For rr% = 1 To 5
        If VUCRSEL(rr%, Index) > "" And InStr(vucrlist(Index).ListItems(r%), "(" + VUCRSEL(rr%, Index) + ")") > 0 Then
            vucrlist(Index).ListItems(r%).Selected = True
            rr% = 5
        End If
    Next rr%
Next r%
vucrf(Index).Top = Command1(Index).Top - 1000
vucrf(Index).Left = 500
vucrf(Index).Visible = True
'---- setfocus logic ----
'         vucrlist(index).SetFocus
          If vucrlist(Index).Visible Then
              vucrlist(Index).SetFocus
           End If
End Sub
'Private Sub Command10_Click()
'relationshipframe(1).Visible = False
'resident(1).SetFocus
'wEnd Sub

Private Sub Command2_Click(Index As Integer)
If fromfind = 1 Then
    Exit Sub
End If
pucrlist(Index).Top = description(Index).Top - 1000
pucrlist(Index).Left = description(Index).Left + 1000
pucrlist(Index).Visible = True
'---- setfocus logic ----
'         pucrlist(index).SetFocus
          If pucrlist(Index).Visible Then
              pucrlist(Index).SetFocus
           End If
End Sub

Private Sub Command10_Click(Index As Integer)

pinfoframe(Index).Visible = False
'description(Index).SetFocus

End Sub

Private Sub Command22_Click(Index As Integer)
pucrlist(Index).ListIndex = -1
group(Index).ListIndex = -1
majorlist(Index).ListIndex = -1
minorlist(Index).ListIndex = -1
End Sub

Private Sub Command7_Click(Index As Integer)
If fromfind = 1 Then
    Exit Sub
End If
sdrugframe(Index).Left = 2000
sdrugframe(Index).Top = Command7(Index).Top
sdrugframe(Index).Visible = True
'drugtype(1 + (Index * 3)).SetFocus
End Sub

Private Sub Command8_Click(Index As Integer)
relationshipframe(Index).Visible = False
'---- setfocus logic ----
'         resident(index).SetFocus
          If resident(Index).Visible Then
              resident(Index).SetFocus
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
    msg = MsgBox("Date/Time entered if invalid.", 48, "Genesis Error Log")
'---- setfocus logic ----
'             DATERECOVERED(index).SetFocus
          If DATERECOVERED(Index).Visible Then
              DATERECOVERED(Index).SetFocus
           End If
End If
DATERECOVERED(Index) = Format$(DATERECOVERED(Index), "mm/dd/yyyy")
DATERECOVERED(Index).Visible = False
If DATERECOVERED(Index) = "Date" Then
    DATERECOVERED(Index) = ""
End If
'=====Data Item 18 and 19
If BACKTAB = 0 Then
If Mid$(pucrlist(Index).List(pucrlist(Index).ListIndex), InStr(pucrlist(Index).List(pucrlist(Index).ListIndex), "(") + 1, 3) = "240" Then
    tempgroup = Val(Mid$(group(Index).List(group(Index).ListIndex), InStr(group(Index).List(group(Index).ListIndex), "(") + 1, 2))
    If (tempgroup = 3 Or tempgroup = 5 Or tempgroup = 24 Or tempgroup = 28 Or tempgroup = 37) And fromfind = 0 Then
        numvehicle((Index Mod 6) + 6).Visible = True
'---- setfocus logic ----
'                 numvehicle((index Mod 6) + 6).SetFocus
          If numvehicle((indexMod6) + 6).Visible Then
              numvehicle((indexMod6) + 6).SetFocus
           End If
    Else
'---- setfocus logic ----
'                 totalvalue(holdrecv + 6).SetFocus
          If totalvalue(holdrecv + 6).Visible Then
              totalvalue(holdrecv + 6).SetFocus
           End If
    End If
Else
'---- setfocus logic ----
'             totalvalue(holdrecv + 6).SetFocus
          If totalvalue(holdrecv + 6).Visible Then
              totalvalue(holdrecv + 6).SetFocus
           End If
End If
Else
    BACKTAB = 0
End If

End Sub

Private Sub DENOMINATION_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
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

Private Sub drugmeasurement_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    If drugmeasurement(Index).ListIndex > -1 Then
        drugmeasurement(Index).ListIndex = -1
    End If
End If

End Sub

Private Sub drugsno_Click(Index As Integer)

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

If Index > 11 Or _
    (Index > 2 And Index < 5 And Val(subject(0)) > 0) Or _
    (Index > 8 And Index < 12 And Val(subject(1)) > 0) Then
            
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
Select Case Index
    Case 0 To 2
        idx2 = 0
    Case 3 To 5
        idx2 = 1
    Case 6 To 8
        idx2 = 2
    Case 9 To 11
        idx2 = 3
    Case 12 To 14
        idx2 = 4
    Case 15 To 17
        idx2 = 5
    Case 18 To 20
        idx2 = 6
    Case 21 To 23
        idx2 = 7
    Case 24 To 26
        idx2 = 8
    Case 27 To 29
        idx2 = 9
End Select
If sdrugframe(idx2).Top > (-1 * Picture2.Top) And sdrugframe(idx2).Top + sdrugframe(idx2).Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If sdrugframe(idx2).Top > 500 Then
    VScroll1 = sdrugframe(idx2).Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub drugtype_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    If drugtype(Index).ListIndex > -1 Then
        drugtype(Index).ListIndex = -1
    End If
End If

End Sub

Private Sub ethnicity_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        ethnicity(Index).ListIndex = -1
    End If

End Sub

Private Sub exclear18andover_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub exclearunder18_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub extraditiondenied_GotFocus()
If Frame18.Top > (-1 * Picture2.Top) And Frame18.Top + Frame18.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Frame18.Top > 500 Then
    VScroll1 = Frame18.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub FOLLOWUPOFFICERDATE_GotFocus()
If followupofficer.ListIndex > -1 And FOLLOWUPOFFICERDATE = "" Then
    FOLLOWUPOFFICERDATE = Format$(Date$, "mm/dd/yyyy")
End If
End Sub

Private Sub FOLLOWUPOFFICERDATE_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(FOLLOWUPOFFICERDATE) = 1 Or Len(FOLLOWUPOFFICERDATE) = 4 Then
    Call sendslash
End If
End If
End Sub

Private Sub FOLLOWUPOFFICERDATE_LostFocus()
If FOLLOWUPOFFICERDATE > "" And Not IsDate(FOLLOWUPOFFICERDATE) Then
    msg = MsgBox("Date/Time entered is invalid.", 48, "Genesis Error Log")
    If fromfind = 0 Then
'---- setfocus logic ----
'                 FOLLOWUPOFFICERDATE.SetFocus
          If FOLLOWUPOFFICERDATE.Visible Then
              FOLLOWUPOFFICERDATE.SetFocus
           End If
    End If
End If
FOLLOWUPOFFICERDATE = Format$(FOLLOWUPOFFICERDATE, "mm/dd/yyyy")
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

Private Sub Form_Resize()
Picture2.Move 0, 0
'With VScroll1
'    .Max = Picture2.Height - Picture1.Height
'End With
'VScroll1.Visible = (Picture1.Height < Picture2.Height)

End Sub

'Private Sub group_Click(Index As Integer)
'If FROMKEY = 1 Then
'    FROMKEY = 0
'    Exit Sub
'End If
'On Error Resume Next


'group(Index).Visible = False
'On Error GoTo 0
'End Sub


Private Sub Form_Load()
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
    sinciden.WindowState = 2
Else
    sinciden.WindowState = 1
End If
SavePicture Picture2.Picture, "c:\sinc.bmp"
nametype = 1
On Error Resume Next
Kill "*.dsk"
Dim db As Database, rs As Recordset
On Error GoTo oderror1
od1:
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("Select orinumber from system")
If rs.EOF Then
    msg = MsgBox("ORI number is missing.  No exports can be generated.  Contact technical support.", 48, "Genesis Error Log")
    db.Close
    On Error Resume Next
    Exit Sub
End If
rs.MoveFirst
orinumber = rs("orinumber")
FROMKEY = 0
On Error Resume Next


Me.Height = 7600
Me.Width = 11700
Me.Top = 0
Me.Left = 0
Screen.MousePointer = 11
Call loadcodes
On Error GoTo 0
HI = incidentnumber
Call clearroutine(0)
If fromexport = 0 Then
    Call loadupkey
End If
incidentnumber = HI
With Picture2
    .AutoSize = True
    .Move 0, 0
End With
Picture1.Height = Picture2.Height
VScroll1.Max = Picture2.Height
VScroll1.LargeChange = VScroll1.Max / 10
VScroll1.SmallChange = VScroll1.Max / 100
'VScroll1.Visible = (Picture1.Height < Picture2.Height)
getoutf:
If fromexport = 0 Then
    Call defaultcodes
End If
On Error GoTo getoutf
Open "NP.TAG" For Input As #1
Line Input #1, a$
incidentnumber = a$
Line Input #1, a$
PAGE = a$
Line Input #1, a$
incidentdate = a$
Close #1
Kill "NP.TAG"
On Error GoTo oderror2
od2:
Set db = OpenDatabase(nwi + "INCIDENT.MDB")
Set rs = db.OpenRecordset("select ucr1, ucr2, ucr3,ucr4, ucr5, ucr6,ucr7, ucr8, ucr9, ucr10 from incidentSUPPORT where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
rs.MoveFirst
On Error Resume Next
For Index = 0 To 1
    vucrlist(Index).ListItems.clear
    For vv% = 1 To 10
        If Not IsNull(rs("ucr" + Mid$(Str$(vv%), 2))) Then
            Set rs2 = db.OpenRecordset("select code from ucr where abbrev = '" + rs("ucr" + Mid$(Str$(vv%), 2)) + "'")
            rs2.MoveFirst
            Set itmx2 = vucrlist(Index).ListItems.add(, , rs2("code"))
        End If
    Next vv%
Next Index
FOUNDPROP = False
For Index = 0 To 5
    For t% = 1 To vucrlist(0).ListItems.Count
        '===== 081
        tempucr = Mid$(vucrlist(0).ListItems(t%), InStr(vucrlist(0).ListItems(t%), "(") + 1, 3)
        Select Case tempucr
            Case "240", "100", "120", "200", "210", "220", "23A", "23B", "23C", "23D", "23E", "23F", "23G", "23H", "240", "250", "26A", "26B", "26C", "26D", "26E", "270", "280", "290", "35A", "35B", "39A", "39B", "39C", "39D", "510", "26A", "26B", "26C", "26D", "26E"
                pucrlist(Index).AddItem vucrlist(0).ListItems(t%)
                FOUNDPROP = True
        End Select
    Next t%
Next Index
'For t% = 0 To 41
'    totalvalue(t%).Enabled = FOUNDPROP
'Next t%
For t% = 0 To 5
    Command10(t%).Enabled = FOUNDPROP
Next t%
db.Close
Call incidentnumber_Click


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

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
On Error GoTo 0
End Sub

Private Sub ht_GotFocus(Index As Integer)
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If

End Sub

Private Sub HULL_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

'Private Sub group_KeyPress(Index As Integer, KeyAscii As Integer)
'FROMKEY = 1
'End Sub

'Private Sub group_LostFocus(Index As Integer)
'group(Index).Visible = False
'End Sub

Friend Sub incidentnumber_Click()
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
incidentfound = False
Call findincident(incidentfound)
Picture2.Refresh
Screen.MousePointer = 0
VScroll1 = 0
On Error Resume Next
'---- setfocus logic ----
'         narrativeonly.SetFocus
          If narrativeonly.Visible Then
              narrativeonly.SetFocus
           End If
On Error GoTo 0
End Sub

Private Sub Incidentnumber_GotFocus()
If HI > "" Then
    incidentnumber = HI
    HI = ""
End If
End Sub

Private Sub Incidentnumber_LostFocus()
If Len(incidentnumber) < 12 And Len(incidentnumber) > 0 Then
    incidentnumber = incidentnumber + Space$(12 - Len(incidentnumber))
End If
End Sub

Private Sub ISSUER_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub JAIL_GotFocus(Index As Integer)
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If

End Sub

Private Sub juvenilenocustody_GotFocus()
If Frame18.Top > (-1 * Picture2.Top) And Frame18.Top + Frame18.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Frame18.Top > 500 Then
    VScroll1 = Frame18.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub majorlist_Click(Index As Integer)
If majorlist(Index).ListIndex = -1 Then
    Exit Sub
End If
Call setminorlist(Index)
End Sub




Private Sub MAKE_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub MISC_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub MODEL_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
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

Private Sub NARRATIVE_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift = vbCtrlMask) And (KeyCode = vbKeyF2) Then
        Call Command15_Click
    End If
End Sub

Private Sub narrative_LostFocus()
If narrativeonly = 1 Then
'---- setfocus logic ----
'             reportingofficer(0).SetFocus
          If reportingofficer(0).Visible Then
              reportingofficer(0).SetFocus
           End If
End If
End Sub

Private Sub narrativeonly_Click()
If narrativeonly = 1 Then
    noframe(0).Top = 2070
    noframe(0).Left = 240
    noframe(0).Visible = True
    noframe(1).Top = 12030
    noframe(1).Left = 240
    noframe(1).Visible = True
    If NARRATIVE.Visible = True Then
'---- setfocus logic ----
'                 narrative.SetFocus
          If NARRATIVE.Visible Then
              NARRATIVE.SetFocus
           End If
    End If
Else
    noframe(0).Visible = False
    noframe(1).Visible = False
End If
End Sub

Private Sub noprosecution_GotFocus()
If Frame18.Top > (-1 * Picture2.Top) And Frame18.Top + Frame18.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Frame18.Top > 500 Then
    VScroll1 = Frame18.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub NIC_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
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
If BACKTAB = 0 Then
'---- setfocus logic ----
'         totalvalue(holdrecv + 6).SetFocus
          If totalvalue(holdrecv + 6).Visible Then
              totalvalue(holdrecv + 6).SetFocus
           End If
Else
    BACKTAB = 0
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

Private Sub pinfoframe_click(Index As Integer)
    pinfoframe(Index).ZOrder
End Sub



Private Sub PropertyCommand_Click(Index As Integer)
If pucrlist(Index).ListCount = 0 Then
    Exit Sub
End If
'If vucrlist(Index).SelectedItem = -1 Then
'    vucrlist(Index).Clear
'    For vv% = 0 To vucrlist.Count - 1
'        For VVV% = 0 To vucrlist(vv%).ListCount - 1
'            If vucrlist(vv%).Selected(VVV%) = True Then
'                foundit% = 0
'                For yy% = 0 To pvucrlist(Index).ListItems.Count - 1
'                    If vucrlist(vv%).List(VVV%) = pvucrlist(Index).List(yy%) Then
'                        foundit% = 1
'                        yy% = pvucrlist(Index).ListItems.Count - 1
'                    End If
'                Next yy%
'                If foundit% = 0 Then
'                    vucrlist(Index).ListItems.Add vucrlist(vv%).List(VVV%)
'                End If
'                VVV% = vucrlist(vv%).ListCount - 1
'            End If
'        Next VVV%
'    Next vv%
'Else
'    HP = vucrlist(Index).SelectedItem.Index
'    HC = vucrlist(Index).ListItems.Count
'    For vv% = 0 To vucrlist.Count - 1
'        For VVV% = 0 To vucrlist(vv%).ListCount - 1
'            If vucrlist(vv%).Selected(VVV%) = True Then
'                foundit% = 0
'                For yy% = 0 To vucrlist(Index).ListItems.Count - 1
'                    If vucrlist(vv%).List(VVV%) = vucrlist(Index).List(yy%) Then
'                        foundit% = 1
'                        yy% = vucrlist(Index).ListItems.Count - 1
'                    End If
'                Next yy%
'                If foundit% = 0 Then
'                    vucrlist(Index).AddItem vucrlist(vv%).List(VVV%)
'                End If
'                VVV% = vucrlist(vv%).ListCount - 1
'            End If
'        Next VVV%
'    Next vv%
'    If HC = vucrlist(Index).ListItems.Count Then
'        FROMKEY = 1
'        vucrlist(Index).SelectedItem.Index = HP
'    End If
'End If
'If fromfind = 1 Then
'    Exit Sub
'End If
FROMKEY = 0
pinfoframe(Index).Left = PropertyCommand(Index).Left + 1200
pinfoframe(Index).Top = 12300
pinfoframe(Index).Visible = True
group(Index).Visible = True
pucrlist(Index).Visible = True
'---- setfocus logic ----
'         pucrlist(index).SetFocus
          If pucrlist(Index).Visible Then
              pucrlist(Index).SetFocus
           End If
'rlb code

If Index = 5 Then
    pinfoframe(Index).Top = (PropertyCommand(Index).Top - CLng(pinfoframe(Index).Height))
    pinfoframe(Index).Left = (PropertyCommand(Index).Left + PropertyCommand(Index).Width) - pinfoframe(Index).Width
Else
    pinfoframe(Index).Top = (PropertyCommand(Index).Top - CLng(0.5 * pinfoframe(Index).Height))
    pinfoframe(Index).Left = PropertyCommand(Index).Left - (0.25 * (pinfoframe(Index).Width))
End If

pinfoframe(Index).ZOrder


End Sub

Private Sub race_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        race(Index).ListIndex = -1
    End If

End Sub

'Private Sub pucrlist_Click(Index As Integer)'''


'If FROMKEY = 1 Then
'    FROMKEY = 0
'    Exit Sub
'End If
''===== Data Element 18, 19
'Dim tempgroup As Integer, ITMX As String
'tempgroup = Val(Mid$(group(Index).List(group(Index).ListIndex), InStr(group(Index).List(group(Index).ListIndex), "(") + 1, 2))
'pucrlist(Index).Visible = False

'End Sub

'Private Sub pucrlist_LostFocus(Index As Integer)
'pucrlist(Index).Visible = False
'If BACKTAB = 0 Then
'description(Index).SetFocus
'Else
'    BACKTAB = 0
'End If
'End Sub
Private Sub relationship_LostFocus(Index As Integer)
For ty% = Index To Index
    For tv% = 0 To relationship(ty%).ListCount - 1
        If relationship(ty%).Selected(tv%) = True Then
            relationship(ty%).ListIndex = tv%
            tv% = relationship(ty%).ListCount - 1
        End If
    Next tv%
Next ty%
If relationship(Index).ListIndex = -1 Then
    If Index < 10 Then
        Call Command8_Click(0)
    Else
        Call Command8_Click(1)
    End If
End If
    
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

Private Sub resident_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        resident(Index).ListIndex = -1
    End If

End Sub

Private Sub RUNAWAY_GotFocus(Index As Integer)
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If

End Sub

Private Sub SARTICLE_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub SBOAT_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub scolor_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub SECURITIESDATE_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub SECURITIESDATE_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(SECURITIESDATE) = 1 Or Len(SECURITIESDATE) = 4 Then
    Call sendslash
End If
End If

End Sub

Private Sub SERIAL_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub SERIALSTATE_GotFocus()
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
If Val(subject(Index)) > 0 Then
    Exit Sub
End If
If fromfind = 1 Then
    Exit Sub
End If
relationshipframe(Index).Left = 7000
relationshipframe(Index).Top = setrel(Index).Top - 1000
relationshipframe(Index).Visible = True
'---- setfocus logic ----
'         relationship(0 + (index * 10)).SetFocus
          If relationship(0 + (Index * 10)).Visible Then
              relationship(0 + (Index * 10)).SetFocus
           End If
End Sub



Private Sub sex_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        sex(Index).ListIndex = -1
    End If

End Sub

Private Sub SFOUND_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub SGUN_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub SLICENSE_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub SRECOVERED_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub SSECURITIES_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub SSTOLEN_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub SSUSPECT_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub state_GotFocus(Index As Integer)
If vsname(Index) = "UNKNOWN" Then
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

Private Sub STOWED_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub STYLE_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub stype_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
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

Private Sub subjectlocatedno_GotFocus()
If Frame17.Top > (-1 * Picture2.Top) And Frame17.Top + Frame17.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Frame17.Top > 500 Then
    VScroll1 = Frame17.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub subjectlocatedyes_GotFocus()
If Frame17.Top > (-1 * Picture2.Top) And Frame17.Top + Frame17.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Frame17.Top > 500 Then
    VScroll1 = Frame17.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub SUMMONS_GotFocus(Index As Integer)
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If

End Sub

Private Sub SVEHICLE_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub SVICTIM_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim editerr As Integer, begindate, EndDate As String
editerr% = 0
Select Case Button
    
    Case "NextPage"
        If incidentnumber = "" Then
            msg = MsgBox("A valid incidentnumber must be present.", 48, "Genesis Error Log")
            Exit Sub
        End If
        schanged = 1
        If schanged = 1 Then
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
            POPMSG$ = ""
            Call editevent(editerr%, POPMSG$)
            If editerr% = 0 Then
                Call editvictim(editerr%, POPMSG$)
                If editerr% = 0 Then
                    Call editsubject(editerr%, POPMSG$)
                    If editerr% = 0 Then
                        Call editproperty(editerr%, POPMSG$)
                    End If
                End If
            End If
            If editerr% = 0 Then
                tempsave = 0
                If UCase(frmLogin.txtUserName) = "DEMO" And UCase(frmLogin.txtPassword) = "DEMO" Then
                Else
                    Call saveincident
                End If
            Else
                MsgBox POPMSG$, 48, "Genesis Error Log"
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If
        HI = incidentnumber
        HP = Val(PAGE)
        HP = HP + 1
        Call clearroutine(1)
        incidentnumber = ""
        sinciden.incidentnumber = HI
        PAGE = Mid$(Str$(HP), 2)
        Call sinciden.incidentnumber_Click
        
    
    Case "PriorPage"
        If incidentnumber = "" Then
            msg = MsgBox("A valid incidentnumber must be present.", 48, "Genesis Error Log")
            Exit Sub
        End If
        If vsname(0) = "" And vsname(1) = "" And description(0) = "" And totalvalue(0) = "" And VIN = "" And HULL = "" And SERIAL = "" And model = "" Then
        Else
            schanged = 1
            If schanged = 1 Then
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
                POPMSG$ = ""
                Call editevent(editerr%, POPMSG$)
                If editerr% = 0 Then
                    Call editvictim(editerr%, POPMSG$)
                    If editerr% = 0 Then
                        Call editsubject(editerr%, POPMSG$)
                        If editerr% = 0 Then
                            Call editproperty(editerr%, POPMSG$)
                        End If
                    End If
                End If
                If editerr% = 0 Then
                    tempsave = 0
                    If UCase(frmLogin.txtUserName) = "DEMO" And UCase(frmLogin.txtPassword) = "DEMO" Then
                    Else
                      Call saveincident
                    End If
                Else
                    MsgBox POPMSG$, 48, "Genesis Error Log"
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
        End If
        HI = incidentnumber
        HP = Val(PAGE)
        HP = HP - 1
        If HP < 1 Then
            Unload sinciden
            Open "pp.tag" For Output As #1
            Print #1, HI
            Close #1
            incident.WindowState = vbMaximized
            incident.Show
        Else
            Call clearroutine(1)
            sinciden.incidentnumber = HI
            PAGE = Mid$(Str$(HP), 2)
            Call sinciden.incidentnumber_Click
            
        End If
        
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
        POPMSG$ = ""
        Call editevent(editerr%, POPMSG$)
        If editerr% = 0 Then
            Call editvictim(editerr%, POPMSG$)
            If editerr% = 0 Then
                Call editsubject(editerr%, POPMSG$)
                If editerr% = 0 Then
                    Call editproperty(editerr%, POPMSG$)
                End If
            End If
        End If
        If editerr% = 0 Then
            tempsave = 0
            Call saveincident
        Else
            MsgBox POPMSG$, 48, "Genesis Error Log"
        End If
        On Error Resume Next
        Screen.MousePointer = 0
    Case "Clear"
        Screen.MousePointer = 11
        Call clearroutine(0)
        Screen.MousePointer = 0
    Case "Delete"
        If UCase(frmLogin.txtUserName) = "DEMO" And UCase(frmLogin.txtPassword) = "DEMO" Then
            msg = MsgBox("Not available in DEMO version.", 48, "Genesis Information Log")
            Screen.MousePointer = 0
            Exit Sub
        End If
        Screen.MousePointer = 11
        Call deleteroutine
        Call clearroutine(0)
        Screen.MousePointer = 0
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
            POPMSG$ = ""
            Call editevent(editerr%, POPMSG$)
            If editerr% = 0 Then
                Call editvictim(editerr%, POPMSG$)
                If editerr% = 0 Then
                    Call editsubject(editerr%, POPMSG$)
                    If editerr% = 0 Then
                            Call editproperty(editerr%, POPMSG$)
                    End If
                End If
            End If
            If editerr% = 0 Then
                tempsave = 0
                If UCase(frmLogin.txtUserName) = "DEMO" And UCase(frmLogin.txtPassword) = "DEMO" Then
                Else
                    Call saveincident
                End If
            Else
                MsgBox POPMSG$, 48, "Genesis Error Log"
                msg = MsgBox("An incident report cannot be printed with errors.", 48, "Genesis Error Log")
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If
        On Error GoTo bbhtest
'ian's code added(page)
        If narrativeonly = 1 Then
            report.ReportFileName = nwi + "snincident.rpt"
        Else
            report.ReportFileName = nwi + "sincident.rpt"
        End If
        'added (page) here
        report.SelectionFormula = "{supplemental.incidentnumber} = '" + incidentnumber + "' and {supplemental.page} = " + PAGE
        report.PrintFileType = crptCrystal
        report.Destination = crptToPrinter
        report.Action = 1
        Screen.MousePointer = 0
'end ian's code
'end print button code

    Case "TempSave"
        If UCase(frmLogin.txtUserName) = "DEMO" And UCase(frmLogin.txtPassword) = "DEMO" Then
            msg = MsgBox("Not available in DEMO version.", 48, "Genesis Information Log")
            Screen.MousePointer = 0
            Exit Sub
        End If
        If incidentnumber = "" Then
            msg = MsgBox("An incident number and date and name must be entered for a TEMP SAVE.", 48, "Genesis Error Log")
            Exit Sub
        End If
        tempsave = 1
        Screen.MousePointer = 11
        Call saveincident
        Call clearroutine(0)
        Screen.MousePointer = 0
    Case "TempList"
        temprevw.WindowState = vbMaximized
        temprevw.Show
    Case "Defaults"
        defaults.Show
    Case "Search"
        Unload Me
        Search.Show
    Case "Exit"
        Unload sinciden
End Select
Exit Sub
bbhtest:
Resume Next
End Sub
Private Sub loadupkey()
Dim db As Database, rs As Recordset, tabname As String
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "LAWSUITE.mdb")
Set rs = db.OpenRecordset("select DPnameLF from people")
If Not rs.EOF Then
    rs.MoveFirst
Else
    On Error Resume Next
    db.Close
    Exit Sub
End If
On Error Resume Next
vsname(0).clear
vsname(1).clear
While Not rs.EOF
    If Not IsNull(rs("DPnameLF")) Then
        vsname(0).AddItem rs("DPnameLF")
        vsname(1).AddItem rs("DPnameLF")
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

Friend Sub editevent(editerr As Integer, msg As String)
'RLB Bandaid
On Error GoTo rlbErr:
Dim testgroup, totgroup As String, temperr As Integer, tempucr, tempgroup, typeselect As String, tempvalue As Single, tempdate As String
Screen.MousePointer = 11
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
    GoTo exitedite
End If
editeventdetail:
'===== Data Element 2
If incidentnumber = "" Then
    msg = "A valid incidentnumber must be entered."
    Call ShowApplicableContainers(APPROVINGOFFICERDATE)
'---- setfocus logic ----
'             APPROVINGOFFICERDATE.SetFocus
          If APPROVINGOFFICERDATE.Visible Then
              APPROVINGOFFICERDATE.SetFocus
           End If
    GoTo exitedite
End If
If Len(incidentnumber) > 12 Then
    msg = "The Incident Number cannot be over 12 characters long."
    Call ShowApplicableContainers(APPROVINGOFFICERDATE)
'---- setfocus logic ----
'             APPROVINGOFFICERDATE.SetFocus
          If APPROVINGOFFICERDATE.Visible Then
              APPROVINGOFFICERDATE.SetFocus
           End If
    GoTo exitedite
End If
For t% = 1 To Len(incidentnumber)
    If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789- ", Mid$(incidentnumber, t%, 1)) = 0 Then
        msg = "An invalid character has been found in the Incident Number field.  Valid characters are A-Z, 0-9, and Hyphen.  Do not enter any Blanks becuase these are computer generated."
        t% = Len(incidentnumber)
        Call ShowApplicableContainers(APPROVINGOFFICERDATE)
'---- setfocus logic ----
'                 APPROVINGOFFICERDATE.SetFocus
          If APPROVINGOFFICERDATE.Visible Then
              APPROVINGOFFICERDATE.SetFocus
           End If
        GoTo exitedite
    End If
Next t%
If Len(incidentnumber) < 12 And Len(incidentnumber) > 0 Then
    incidentnumber = incidentnumber + Space$(12 - Len(incidentnumber))
End If
Call ShowApplicableContainers(APPROVINGOFFICERDATE)
'---- setfocus logic ----
'         APPROVINGOFFICERDATE.SetFocus
          If APPROVINGOFFICERDATE.Visible Then
              APPROVINGOFFICERDATE.SetFocus
           End If
GoTo goodedite
exitedite:
editerr = 1
goodedite:
Exit Sub
'RLB Bandaid
rlbErr:
    If Err.Number = 5 Then Resume Next
End Sub
Friend Sub editvictim(editerr As Integer, msg As String)
Dim db As Database, rs, rs2, rs3 As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("select individual,SOCIETYPUBLIC from incidentreportC where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
Set rs2 = db.OpenRecordset("select offenderdeath, noprosecution, extraditiondenied, victimdeclinescooperation, juvenilenocustody from incidentreportO where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
Set rs3 = db.OpenRecordset("select ucr1, ucr2, ucr3, ucr4, ucr5, ucr6, ucr7, ucr8, ucr9, ucr10 from incidentsupport where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
rs.MoveFirst
rs2.MoveFirst

For t% = 0 To 1
    If Val(victim(t%)) = 0 Then
        ctsel = 0
        For tt% = 1 To vucrlist(t%).ListItems.Count
            If vucrlist(t%).ListItems(tt%).Selected Then
                ctsel = ctsel + 1
                tt% = vucrlist(t%).ListItems.Count
            End If
        Next tt%
        If ctsel > 0 Then
            msg = "Victim information entered in slot " + CStr(t% + 1) + " with no victim number entered."
            GoTo exiteditv
        End If
        GoTo vnextT
    End If
    If vucrlist(t%).ListItems.Count = 0 Then
        msg = "Victim/UCR assignment not completed."
        Call Command1_Click(t%)
        GoTo exiteditv
    Else
    If vucrlist(t%).SelectedItem Is Nothing Then
        For tuv = 1 To vucrlist(t%).ListItems.Count
            vucrlist(t%).ListItems(tuv).Selected = True
        Next tuv
    End If
    End If
    If t% = 0 Then
        If TV1 = "" Then
            msg = "A Type of Victim must be entered."
            Call ShowApplicableContainers(victim(1))
'---- setfocus logic ----
'                     victim(1).SetFocus
          If victim(1).Visible Then
              victim(1).SetFocus
           End If
            GoTo exiteditv
        End If
    Else
        If TV2 = "" Then
            msg = "A Type of Victim must be entered."
            Call ShowApplicableContainers(victim(2))
'---- setfocus logic ----
'                     victim(2).SetFocus
          If victim(2).Visible Then
              victim(2).SetFocus
           End If
            GoTo exiteditv
        End If
    End If
    If age(t%) > "" Then
        If Val(age(t%)) = 0 And age(t%) <> "00" Then
            msg = "Age must be entered in the format of a single age: X or XX (i.e. 3 or 42) or in the format of a range of ages: XXXX (i.e. 2025, meaning 20 - 25 years old)."
        Call ShowApplicableContainers(age(t%))
'---- setfocus logic ----
'                 age(t%).SetFocus
          If age(t%).Visible Then
              age(t%).SetFocus
           End If
        GoTo exiteditv
        End If
    End If
    '===== Error 404
    If age(t%) <> "NN" And age(t%) <> "NB" And age(t%) <> "BB" Then
        For tt% = 1 To Len(age(t%))
            If InStr("0123456789-", Mid$(age(t%), tt%, 1)) = 0 Then
                msg = "An invalid entry in age has been found.  Valid Entry Formats:  X, XX, X-Y, XX-Y, X-YY, XX-YY, XXYY"
                tt% = Len(age(t%))
                Call ShowApplicableContainers(age(t%))
'---- setfocus logic ----
'                         age(t%).SetFocus
          If age(t%).Visible Then
              age(t%).SetFocus
           End If
                GoTo exiteditv
            End If
        Next tt%
    End If
    If InStr(age(t%), "-") > 0 Then
        ag1 = Val(Left$(age(t%), InStr(age(t%), "-") - 1))
        ag2 = Val(Mid$(age(t%), InStr(age(t%), "-") + 1))
        age(t%) = Format$(ag1, "00") + Format$(ag2, "00")
    End If
    '===== Error 410,422
    If Len(age(t%)) = 4 Then
        If Val(Left$(age(t%), 2)) >= Val(Right$(age(t%), 2)) Then
            msg = "For an age range, the first age must be less than the second age."
            Call ShowApplicableContainers(age(t%))
'---- setfocus logic ----
'                     age(t%).SetFocus
          If age(t%).Visible Then
              age(t%).SetFocus
           End If
            GoTo exiteditv
        End If
        If Val(Left$(age(t%), 2)) = 0 Then
            msg = "The low value in an age range cannot be 0."
            Call ShowApplicableContainers(age(t%))
'---- setfocus logic ----
'                     age(t%).SetFocus
          If age(t%).Visible Then
              age(t%).SetFocus
           End If
            GoTo exiteditv
        End If
    End If

    '===== Data Element 20, 21, 22
    '===== Data Element 20, 21, 22
    Dim drugs(2)
    drugs(0) = ""
    drugs(1) = ""
    drugs(2) = ""
    dct% = t% * 6
    For ii% = (t% * 6) To (t% * 6) + 2
        If drugtype(ii%).ListIndex > -1 Then
            drugs(dct%) = drugtype(ii%).List(drugtype(ii%).ListIndex)
            dct% = dct% + 1
        End If
    Next ii%
    If dct% > t% + 6 Then
        For Z% = (t% * 6) To dct%
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
                                Call ShowApplicableContainers(drugmeasurement(Z%))
'---- setfocus logic ----
'                                         drugmeasurement(Z%).SetFocus
          If drugmeasurement(Z%).Visible Then
              drugmeasurement(Z%).SetFocus
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
                                Call ShowApplicableContainers(drugmeasurement(Z%))
'---- setfocus logic ----
'                                         drugmeasurement(Z%).SetFocus
          If drugmeasurement(Z%).Visible Then
              drugmeasurement(Z%).SetFocus
           End If
                                GoTo exiteditv
                        End If
                    End If
                    If InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "DU=") > 0 Or _
                       InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "NP=") > 0 Then
                        If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "DU=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "NP=") > 0 Then
                                msg = "Two different measurements within the same weight category cannot be entered for the same type of drug."
                                Call ShowApplicableContainers(drugmeasurement(Z%))
'---- setfocus logic ----
'                                         drugmeasurement(Z%).SetFocus
          If drugmeasurement(Z%).Visible Then
              drugmeasurement(Z%).SetFocus
           End If
                                GoTo exiteditv
                        End If
                    End If
                End If
            Next ZZ%
        End If
            For ttt% = 1 To Len(drugamt(Z%))
                If InStr("0123456789.", Mid$(drugamt(Z%), ttt%, 1)) = 0 Then
                    msg = "Only 0123456789. are allowed in drug amount.  Enter all fractions as decimals (i.e. 1/2 = .5)."
                    Call ShowApplicableContainers(drugamt(Z%))
'---- setfocus logic ----
'                             drugamt(Z%).SetFocus
          If drugamt(Z%).Visible Then
              drugamt(Z%).SetFocus
           End If
                    GoTo exiteditv
                End If
            Next ttt%
            If drugamt(Z%) = "" Or drugmeasurement(Z%).ListIndex = -1 Then
                If drugs(Z%) > "" And Left$(drugs(Z%), 1) <> "X" And Left$(drugs(Z%), 1) <> "U" Then
                    msg = "Drug Quantity and Measurement Type must be entered/selected."
                    Call ShowApplicableContainers(drugamt(Z%))
'---- setfocus logic ----
'                             drugamt(Z%).SetFocus
          If drugamt(Z%).Visible Then
              drugamt(Z%).SetFocus
           End If
                    GoTo exiteditv
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
                    GoTo exiteditv
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
                        GoTo exiteditv
                    End If
                End If
                '===== Error 384
                If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "XX=") > 0 Then
                    If Val(drugamt(Z%)) <> 1 Then
                        msg = "If drug measurement is NOT REPORTED, drug amount must be 1."
                        Call ShowApplicableContainers(drugamt(Z%))
'---- setfocus logic ----
'                                 drugamt(Z%).SetFocus
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
'                             drugamt(Z%).SetFocus
          If drugamt(Z%).Visible Then
              drugamt(Z%).SetFocus
           End If
                    GoTo exiteditv
                End If
            End If
            '===== Error 362
            If Left$(drugs(Z%), 1) = "X" Then
                If drugtype((t% * 6)).ListIndex = -1 Or drugtype((t% * 6) + 1).ListIndex = -1 Or drugtype((t% * 6) + 2).ListIndex = -1 Then
                    msg = "If X = Over 3 Drug Types is entered, then the other 2 drug types must also be entered."
                    Call ShowApplicableContainers(drugtype(t% * 6))
'---- setfocus logic ----
'                             drugtype((t% * 6)).SetFocus
          If drugtype((t% * 6)).Visible Then
              drugtype((t% * 6)).SetFocus
           End If
                    GoTo exiteditv
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
                    GoTo exiteditv
                End If
            End If
        Next Z%
    End If
        
    '==== Mandatories E - 25 = GIVEN
    '==== Mandatories E - 26, 27, 28
    If (t% = 0 And (TV1 = "I" Or TV1 = "P")) Or (t% = 1 And (TV2 = "I" Or TV2 = "P")) Then
        If Val(age(t%)) = 0 And age(t%) <> "00" Then
            msg = "Invalid age entered."
            Call ShowApplicableContainers(age(t%))
'---- setfocus logic ----
'                     age(t%).SetFocus
          If age(t%).Visible Then
              age(t%).SetFocus
           End If
            GoTo exiteditv
        End If
        If sex(t%).ListIndex = -1 Then
            msg = "Invalid sex entered."
            Call ShowApplicableContainers(sex(t%))
'---- setfocus logic ----
'                     sex(t%).SetFocus
          If sex(t%).Visible Then
              sex(t%).SetFocus
           End If
            GoTo exiteditv
        End If
        If race(t%).ListIndex = -1 Then
            msg = "Invalid race entered."
            Call ShowApplicableContainers(race(t%))
'---- setfocus logic ----
'                     race(t%).SetFocus
          If race(t%).Visible Then
              race(t%).SetFocus
           End If
            GoTo exiteditv
        End If
        If ethnicity(t%).ListIndex = -1 Then
            msg = "Ethnicity is a required entry."
            Call ShowApplicableContainers(ethnicity(t%))
'---- setfocus logic ----
'                     ethnicity(t%).SetFocus
          If ethnicity(t%).Visible Then
              ethnicity(t%).SetFocus
           End If
            GoTo exiteditv
        End If
        If resident(t%).ListIndex = -1 Then
            msg = "Resident Status is a required entry."
            Call ShowApplicableContainers(resident(t%))
'---- setfocus logic ----
'                     resident(t%).SetFocus
          If resident(t%).Visible Then
              resident(t%).SetFocus
           End If
            GoTo exiteditv
        End If
    Else
        '===== Error 458
        If age(t%) > "" Then
            msg = "Age is not a valid entry for Victim if Type of Victim is not Individual."
            Call ShowApplicableContainers(age(t%))
'---- setfocus logic ----
'                     age(t%).SetFocus
          If age(t%).Visible Then
              age(t%).SetFocus
           End If
            GoTo exiteditv
        End If
        If sex(t%).ListIndex > -1 Then
            msg = "Sex is not a valid entry for Victim if Type of Victim is not Individual."
            Call ShowApplicableContainers(sex(t%))
'---- setfocus logic ----
'                     sex(t%).SetFocus
          If sex(t%).Visible Then
              sex(t%).SetFocus
           End If
            GoTo exiteditv
        End If
        If race(t%).ListIndex > -1 Then
            msg = "Race is not a valid entry for Victim if Type of Victim is not Individual."
            Call ShowApplicableContainers(race(t%))
'---- setfocus logic ----
'                     race(t%).SetFocus
          If race(t%).Visible Then
              race(t%).SetFocus
           End If
            GoTo exiteditv
        End If
        If ethnicity(t%).ListIndex > -1 Then
            msg = "Ethnicity is not a valid entry for Victim if Type of Victim is not Individual."
            Call ShowApplicableContainers(ethnicity(t%))
'---- setfocus logic ----
'                     ethnicity(t%).SetFocus
          If ethnicity(t%).Visible Then
              ethnicity(t%).SetFocus
           End If
            GoTo exiteditv
        End If
        If resident(t%).ListIndex > -1 Then
            msg = "Resident Status is not a valid entry for Victim if Type of Victim is not Individual."
            Call ShowApplicableContainers(resident(t%))
'---- setfocus logic ----
'                     resident(t%).SetFocus
          If resident(t%).Visible Then
              resident(t%).SetFocus
           End If
            GoTo exiteditv
        End If
    End If

    '===== Data Element 24
    FOUNDVUCR = 0
    For tt% = 1 To vucrlist(t%).ListItems.Count
        If vucrlist(t%).ListItems(tt%).Selected Then
            FOUNDVUCR = 1
            tt% = vucrlist(t%).ListItems.Count
        End If
    Next tt%
    If FOUNDVUCR = 0 Then
        msg = "At least one UCR code must be connected to the victim."
'---- setfocus logic ----
'                 vucrlist(t%).SetFocus
          If vucrlist(t%).Visible Then
              vucrlist(t%).SetFocus
           End If
        GoTo exiteditv
    End If
    '----edit by Ed Sloan 11/30/99 from SCLED 8/9/95 update----------------
    For tt% = 1 To vucrlist(t%).ListItems.Count
        If vucrlist(t%).ListItems(tt%).Selected Then
            tempvucr = Mid$(vucrlist(t%).ListItems(tt%), InStr(vucrlist(t%).ListItems(tt%), "(") + 1, 3)
            '===== Error 464,465,467
            Select Case tempvucr
                Case "13A", "13B", "13C", "09A", "09B", "09C", "100", "11A", "11B", "11C", "11D", "36A", "36B", "36C", "90A", "90J"
                    If Not ((t% = 0 And (TV1 = "I" Or TV1 = "P")) Or (t% = 1 And (TV2 = "P" Or TV2 = "I"))) Then
                        msg = "Individual or Police Officer must be selected for Crimes Against Person."
                        Call ShowApplicableContainers(vucrlist(t%))
'---- setfocus logic ----
'                                 vucrlist(t%).SetFocus
          If vucrlist(t%).Visible Then
              vucrlist(t%).SetFocus
           End If
                        GoTo exiteditv
                    End If
                Case "90Z"
                Case "90J", "36C", "980", "978", "753", "756"
                Case "35A", "35B", "39A", "39B", "39C", "39D", "370", "40A", "40B", "520", "90B", "90C", "90D", "90G", "90H", "90I", "90E", "90F"
                    If Not rs("SOCIETYPUBLIC") Then
                        msg = "Society must be selected for Crimes Against Society."
                        Call ShowApplicableContainers(vucrlist(t%))
'---- setfocus logic ----
'                                 vucrlist(t%).SetFocus
          If vucrlist(t%).Visible Then
              vucrlist(t%).SetFocus
           End If
                        GoTo exiteditv
                    End If
                Case Else
                    If rs("societypublic") Then
                        msg = "Society cannot be selected for Crimes Against Property."
                        Call ShowApplicableContainers(vucrlist(t%))
'---- setfocus logic ----
'                                 vucrlist(t%).SetFocus
          If vucrlist(t%).Visible Then
              vucrlist(t%).SetFocus
           End If
                        GoTo exiteditv
                    End If
            End Select
            Select Case tempvucr
                    Case "13A", "13B", "13C", "09A", "09B", "09C", "100", "11A", "11B", "11C", "11D", "36A", "36B"
                        If Not ((t% = 0 And (TV1 = "I" Or TV1 = "P")) Or (t% = 1 And (TV2 = "P" Or TV2 = "I"))) Then
                            msg = "Individual must be selected for Crimes Against Person."
                            Call ShowApplicableContainers(vucrlist(t%))
'---- setfocus logic ----
'                                     vucrlist(t%).SetFocus
          If vucrlist(t%).Visible Then
              vucrlist(t%).SetFocus
           End If
                            GoTo exiteditv
                        End If
                    Case "90F"
                        If Not ((t% = 0 And (TV1 = "I" Or TV1 = "P" Or TV1 = "S")) Or (t% = 1 And (TV2 = "P" Or TV2 = "I" Or TV2 = "S"))) Then
                            msg = "Individual or Society must be selected for Family Offenses/Nonviolent."
                            Call ShowApplicableContainers(vucrlist(t%))
'---- setfocus logic ----
'                                     vucrlist(t%).SetFocus
          If vucrlist(t%).Visible Then
              vucrlist(t%).SetFocus
           End If
                            GoTo exiteditv
                        End If
                    Case "90J", "36C", "980", "978", "753", "756"
                    Case "35A", "35B", "39A", "39B", "39C", "39D", "370", "40A", "40B", "520", "90B", "90C", "90D", "90G", "90H", "90I"
                        If Not ((t% = 0 And (TV1 = "S")) Or (t% = 1 And (TV2 = "S"))) Then
                            msg = "Society must be selected for Crimes Against Society."
                            Call ShowApplicableContainers(vucrlist(t%))
'---- setfocus logic ----
'                                     vucrlist(t%).SetFocus
          If vucrlist(t%).Visible Then
              vucrlist(t%).SetFocus
           End If
                            GoTo exiteditv
                        End If
                    Case Else
                        If ((t% = 0 And (TV1 = "S")) Or (t% = 1 And (TV2 = "S"))) Then
                            msg = "Society cannot be selected for Crimes Against Property."
                            Call ShowApplicableContainers(vucrlist(t%))
'---- setfocus logic ----
'                                     vucrlist(t%).SetFocus
          If vucrlist(t%).Visible Then
              vucrlist(t%).SetFocus
           End If
                            GoTo exiteditv
                        End If
            End Select
            '===== Error 481
            If tempvucr = "36B" And Val(age(t%)) > 15 And vucrlist(t%).ListItems(tt%).Selected = True Then
                msg = "For statutory rape, the victim must be less than or equal to 15 years of age."
                Call ShowApplicableContainers(age(t%))
'---- setfocus logic ----
'                         age(t%).SetFocus
          If age(t%).Visible Then
              age(t%).SetFocus
           End If
                GoTo exiteditv
            End If
            '===== SCEdit 8/9/95
            If tempvucr = "23C" Then
                If Len(age(t%)) = 4 Then
                    If Val(Right$(age(t%), 2)) > 15 Then
                        msg = "For Offense 23C, the victim age must be 15 years old or less."
                        Call ShowApplicableContainers(age(t%))
'---- setfocus logic ----
'                                 age(t%).SetFocus
          If age(t%).Visible Then
              age(t%).SetFocus
           End If
                        GoTo exiteditv
                    End If
                Else
                If Len(age(t%)) = 2 Then
                    If Val(age(t%)) > 15 Then
                        msg = "For Offense 23C, the victim age must be 15 years old or less."
                        Call ShowApplicableContainers(age(t%))
'---- setfocus logic ----
'                                 age(t%).SetFocus
          If age(t%).Visible Then
              age(t%).SetFocus
           End If
                        GoTo exiteditv
                    End If
                Else
                    msg = "For Offense 23C, the victim age must be 15 years old or less."
                    Call ShowApplicableContainers(age(t%))
'---- setfocus logic ----
'                             age(t%).SetFocus
          If age(t%).Visible Then
              age(t%).SetFocus
           End If
                    GoTo exiteditv
                End If
                End If
            End If
        End If
    Next tt%
    '----------------------------------------------------------------------
    '----edit by Ed Sloan 12/03/99 from SCLED 3/5/97 update----------------
    Dim ucrexists, valinj As Boolean
    ucrexists = False
    valinj = False
    foundinjtype% = 0
    For r% = 1 To vucrlist(t%).ListItems.Count
        If vucrlist(t%).ListItems(r%).Selected Then
            tempvucr = Mid$(vucrlist(t%).ListItems(r%), InStr(vucrlist(t%).ListItems(r%), "(") + 1, 3)
            Select Case tempvucr
                Case "100", "11A", "11B", "11C", "11D", "120", "13A", "13B", "210"
                    foundinjtype% = 1
            End Select
        End If
    Next r%
    For r% = 1 To vucrlist(t%).ListItems.Count
        If vucrlist(t%).ListItems(r%).Selected Then
            tempvucr = Mid$(vucrlist(t%).ListItems(r%), InStr(vucrlist(t%).ListItems(r%), "(") + 1, 3)
            ICT% = 0
            For rr% = 1 To injury(t%).ListItems.Count
                If injury(t%).ListItems(rr%).Selected Then
                    ICT% = ICT% + 1
                End If
            Next rr%
            Select Case tempvucr
                '===== Error 479
                Case "13B"
                    ucrexists = True
                    For q% = 1 To injury(t%).ListItems.Count
                        If injury(t%).ListItems(q%).Selected Then
                            tempinj = Mid$(injury(t%).ListItems(q%), InStr(injury(t%).ListItems(q%), "(") + 1, 1)
                            If (tempinj = "M" Or tempinj = "N") Then
                                validinj = True
                            End If
                        End If
                    Next q%
                    If ucrexists And Not validinj Then
                        msg = "For simple assault, the only injury types can be minor or none."
                        Call ShowApplicableContainers(vucrlist(t%))
'---- setfocus logic ----
'                                 vucrlist(t%).SetFocus
          If vucrlist(t%).Visible Then
              vucrlist(t%).SetFocus
           End If
                        GoTo exiteditv
                    End If
                '===== Error 401
                Case "100", "11A", "11B", "11C", "11D", "120", "13A", "13B", "210"
                    If ICT% = 0 Then
                        msg = "Type of injury must be selected for UCR " + tempvcur + "."
                        Call ShowApplicableContainers(vucrlist(t%))
'---- setfocus logic ----
'                                 vucrlist(t%).SetFocus
          If vucrlist(t%).Visible Then
              vucrlist(t%).SetFocus
           End If
                        GoTo exiteditv
                    End If
               '===== Error 419
                Case Else
                If ICT% > 0 And foundinjtype% = 0 Then
                    foundother% = 0
                    For q% = 1 To vucrlist(t%).ListItems.Count
                        If vucrlist(t%).ListItems(q%).Selected Then
                            Select Case Mid$(vucrlist(t%).ListItems(q%), InStr(vucrlist(t%).ListItems(q%), "(") + 1, 3)
                                Case "100", "11A", "11B", "11C", "11D", "120", "13A", "13B", "210"
                                    foundother% = 1
                                    q% = t% - 1
                            End Select
                        End If
                    Next q%
                    If foundother% = 0 Then
                        msg = "Type of injury is not applicable for UCR " + tempvcur + "."
                        Call ShowApplicableContainers(injury(t%))
'---- setfocus logic ----
'                                 injury(t%).SetFocus
          If injury(t%).Visible Then
              injury(t%).SetFocus
           End If
                        GoTo exiteditv
                    End If
                End If
            End Select
        End If
    Next r%
    '===== Error 407
    For r% = 1 To injury(t%).ListItems.Count
        If injury(t%).ListItems(r%).Selected Then
            If Mid$(injury(t%).ListItems(r%), InStr(injury(t%).ListItems(r%), "(") + 1, 1) = "N" Then
                For rr% = 1 To r% - 1
                    If injury(t%).ListItems(rr%).Selected Then
                        msg = "When Ijury Type N=None is selected, no other values may be selected."
                        Call ShowApplicableContainers(injury(t%))
'---- setfocus logic ----
'                                 injury(t%).SetFocus
          If injury(t%).Visible Then
              injury(t%).SetFocus
           End If
                        GoTo exiteditv
                    End If
                Next rr%
                For rr% = r% + 1 To injury(t%).ListItems.Count
                    If injury(t%).ListItems(rr%).Selected Then
                        msg = "When Ijury Type N=None is selected, no other values may be selected."
                        Call ShowApplicableContainers(injury(t%))
'---- setfocus logic ----
'                                 injury(t%).SetFocus
          If injury(t%).Visible Then
              injury(t%).SetFocus
           End If
                        GoTo exiteditv
                    End If
                Next rr%
            End If
        End If
    Next r%
    '===== Error 450
    For r% = t% * 10 To (t% * 10) + 9
        If relationship(r%).ListIndex > -1 Then
            temprel = Mid$(relationship(r%).List(relationship(r%).ListIndex), InStr(relationship(r%).List(relationship(r%).ListIndex), "(") + 1, 2)
            If temprel = "SE" Then
                If Val(age(t%)) < 10 Then
                    msg = "The relationship of victim to subject cannot be 'SE' when victim's age is less than 10."
                    Call ShowApplicableContainers(age(t%))
'---- setfocus logic ----
'                             age(t%).SetFocus
          If age(t%).Visible Then
              age(t%).SetFocus
           End If
                    GoTo exiteditv
                End If
            End If
        End If
    Next r%

    For tt% = 1 To 10
    
        If Not IsNull(rs3("ucr" + Mid$(Str$(tt%), 2))) Then
            tempucr = rs3("ucr" + Mid$(Str$(tt%), 2))
            If tempucr = "13A" Or tempucr = "13B" Or tempucr = "13C" Then
                If UCase(vsname(t%)) <> "UNKNOWN" Then
                    If relationship(t% * 10).ListIndex = -1 And relationship((t% * 10) + 1).ListIndex = -1 And relationship((t% * 10) + 2).ListIndex = -1 And relationship((t% * 10) + 3).ListIndex = -1 And relationship((t% * 10) + 4).ListIndex = -1 And relationship((t% * 10) + 5).ListIndex = -1 And relationship((t% * 10) + 6).ListIndex = -1 And relationship((t% * 10) + 7).ListIndex = -1 And relationship((t% * 10) + 8).ListIndex = -1 And relationship((t% * 10) + 9).ListIndex = -1 Then
                        msg = "A relationship to subject must be selected."
                        Call ShowApplicableContainers(relationship(t% + 10))
'---- setfocus logic ----
'                                 relationship(t% * 10).SetFocus
          If relationship(t% * 10).Visible Then
              relationship(t% * 10).SetFocus
           End If
                        GoTo exiteditv
                    End If
                End If
            End If
            '===== Additional F 12
            '===== Data Element 7, 13
            If tempucr = "09A" Or tempucr = "09B" Or tempucr = "09C" Then
                If UCase(vsname(t%)) <> "UNKNOWN" Then
                    If relationship(t% * 10).ListIndex = -1 And relationship((t% * 10) + 1).ListIndex = -1 And relationship((t% * 10) + 2).ListIndex = -1 And relationship((t% * 10) + 3).ListIndex = -1 And relationship((t% * 10) + 4).ListIndex = -1 And relationship((t% * 10) + 5).ListIndex = -1 And relationship((t% * 10) + 6).ListIndex = -1 And relationship((t% * 10) + 7).ListIndex = -1 And relationship((t% * 10) + 8).ListIndex = -1 And relationship((t% * 10) + 9).ListIndex = -1 Then
                        msg = "A relationship to subject must be selected."
                        Call ShowApplicableContainers(relationship(t% + 10))
'---- setfocus logic ----
'                                 relationship(t% * 10).SetFocus
          If relationship(t% * 10).Visible Then
              relationship(t% * 10).SetFocus
           End If
                        GoTo exiteditv
                    End If
                End If
            End If
            '===== Additional F 13
            '===== Data Element 7
            If tempucr = "100" Then
                If UCase(vsname(t%)) <> "UNKNOWN" Then
                    If relationship(t% * 10).ListIndex = -1 And relationship((t% * 10) + 1).ListIndex = -1 And relationship((t% * 10) + 2).ListIndex = -1 And relationship((t% * 10) + 3).ListIndex = -1 And relationship((t% * 10) + 4).ListIndex = -1 And relationship((t% * 10) + 5).ListIndex = -1 And relationship((t% * 10) + 6).ListIndex = -1 And relationship((t% * 10) + 7).ListIndex = -1 And relationship((t% * 10) + 8).ListIndex = -1 And relationship((t% * 10) + 9).ListIndex = -1 Then
                        msg = "A relationship to subject must be selected."
                        Call ShowApplicableContainers(relationship(t% + 10))
'---- setfocus logic ----
'                                 relationship(t% * 10).SetFocus
          If relationship(t% * 10).Visible Then
              relationship(t% * 10).SetFocus
           End If
                        GoTo exiteditv
                    End If
                End If
            End If
            '===== Additional F 18
            If tempucr = "120" Then
                If (t% = 0 And (TV1 = "I" Or TV1 = "P")) Or (t% = 1 And (TV2 = "I" Or TV2 = "P")) Then
                    If UCase(vsname(t%)) <> "UNKNOWN" Then
                        If relationship(t% * 10).ListIndex = -1 And relationship((t% * 10) + 1).ListIndex = -1 And relationship((t% * 10) + 2).ListIndex = -1 And relationship((t% * 10) + 3).ListIndex = -1 And relationship((t% * 10) + 4).ListIndex = -1 And relationship((t% * 10) + 5).ListIndex = -1 And relationship((t% * 10) + 6).ListIndex = -1 And relationship((t% * 10) + 7).ListIndex = -1 And relationship((t% * 10) + 8).ListIndex = -1 And relationship((t% * 10) + 9).ListIndex = -1 Then
                            msg = "A relationship to subject must be selected."
                            Call ShowApplicableContainers(relationship(t% + 10))
'---- setfocus logic ----
'                                     relationship(t% * 10).SetFocus
          If relationship(t% * 10).Visible Then
              relationship(t% * 10).SetFocus
           End If
                            GoTo exiteditv
                        End If
                    End If
                End If
            End If
            '===== Additional F 19
            If tempucr = "11A" Or tempucr = "11B" Or tempucr = "11C" Or tempucr = "11D" Then
                If UCase(vsname(t%)) <> "UNKNOWN" Then
                    If relationship(t% * 10).ListIndex = -1 And relationship((t% * 10) + 1).ListIndex = -1 And relationship((t% * 10) + 2).ListIndex = -1 And relationship((t% * 10) + 3).ListIndex = -1 And relationship((t% * 10) + 4).ListIndex = -1 And relationship((t% * 10) + 5).ListIndex = -1 And relationship((t% * 10) + 6).ListIndex = -1 And relationship((t% * 10) + 7).ListIndex = -1 And relationship((t% * 10) + 8).ListIndex = -1 And relationship((t% * 10) + 9).ListIndex = -1 Then
                        msg = "A relationship to subject must be selected."
                        Call ShowApplicableContainers(relationship(t% + 10))
'---- setfocus logic ----
'                                 relationship(t% * 10).SetFocus
          If relationship(t% * 10).Visible Then
              relationship(t% * 10).SetFocus
           End If
                        GoTo exiteditv
                    End If
                End If
            End If
            '===== Additional F 20
            If tempucr = "36A" Or tempucr = "36B" Or tempucr = "36C" Then
                If UCase(vsname(t%)) <> "UNKNOWN" Then
                    If relationship(t% * 10).ListIndex = -1 And relationship((t% * 10) + 1).ListIndex = -1 And relationship((t% * 10) + 2).ListIndex = -1 And relationship((t% * 10) + 3).ListIndex = -1 And relationship((t% * 10) + 4).ListIndex = -1 And relationship((t% * 10) + 5).ListIndex = -1 And relationship((t% * 10) + 6).ListIndex = -1 And relationship((t% * 10) + 7).ListIndex = -1 And relationship((t% * 10) + 8).ListIndex = -1 And relationship((t% * 10) + 9).ListIndex = -1 Then
                        msg = "A relationship to subject must be selected."
                        Call ShowApplicableContainers(relationship(t% + 10))
'---- setfocus logic ----
'                                 relationship(t% * 10).SetFocus
          If relationship(t% * 10).Visible Then
              relationship(t% * 10).SetFocus
           End If
                        GoTo exiteditv
                    End If
                End If
            End If
        End If
    Next tt%

vnextT:
Next t%
GoTo goodeditv
exiteditv:
editerr = 1
goodeditv:
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
Friend Sub editsubject(editerr As Integer, msg As String)

Dim db As Database, rs, rs2 As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("select policeofficer,individual from incidentreportC where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
Set rs2 = db.OpenRecordset("select offenderdeath, noprosecution, extraditiondenied, victimdeclinescooperation, juvenilenocustody from incidentreportO where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
rs.MoveFirst

For t% = 0 To 1
    If Val(subject(t%)) = 0 Then
        Call ShowApplicableContainers(subject(t%))
'---- setfocus logic ----
'                 subject(t%).SetFocus
          If subject(t%).Visible Then
              subject(t%).SetFocus
           End If
        GoTo snextt
    End If
    '===== SC LEOKA
    If rs("policeofficer") Then
        If TWOMANVEHICLE(t%) = 0 And ONEMANVEHICLE(t%) = 0 And DETECTIVE(t%) = 0 And TODOTHER(t%) = 0 Then
            msg = "If Type Victim is Police Officer, a selection must be made for Two Man Vehicle, One Man Vehicle, Detective/Special Assignment, or Other."
            Call ShowApplicableContainers(TWOMANVEHICLE(t%))
'---- setfocus logic ----
'                     TWOMANVEHICLE(t%).SetFocus
          If TWOMANVEHICLE(t%).Visible Then
              TWOMANVEHICLE(t%).SetFocus
           End If
            GoTo exitedits
        End If
        If Not TWOMANVEHICLE(t%) Then
            If ALONE(t%) = 0 And ASSISTED(t%) = 0 Then
                msg = "If Type Victim is Police Officer and not Two Man Vehicle, a selection must be made for Alone or Assisted."
                Call ShowApplicableContainers(TWOMANVEHICLE(t%))
'---- setfocus logic ----
'                         TWOMANVEHICLE(t%).SetFocus
          If TWOMANVEHICLE(t%).Visible Then
              TWOMANVEHICLE(t%).SetFocus
           End If
                GoTo exitedits
            End If
        End If
    End If
    tage = age(t%)
    '===== Error 761
    If RUNAWAY(t%) = 1 Then
        If Len(tage) = 2 Then
            If Val(tage) > 17 Then
                msg = "A runaway must be under the age of 18."
                Call ShowApplicableContainers(RUNAWAY(t%))
'---- setfocus logic ----
'                         runaway(t%).SetFocus
          If RUNAWAY(t%).Visible Then
              RUNAWAY(t%).SetFocus
           End If
                GoTo exitedits
            End If
        Else
        If Len(tage) = 4 Then
            If Val(Right$(tage, 2)) > 17 Then
                msg = "A runaway must be under the age of 18."
                Call ShowApplicableContainers(RUNAWAY(t%))
'---- setfocus logic ----
'                         runaway(t%).SetFocus
          If RUNAWAY(t%).Visible Then
              RUNAWAY(t%).SetFocus
           End If
                GoTo exitedits
            End If
        Else
            msg = "A runaway must be under the age of 18."
            Call ShowApplicableContainers(RUNAWAY(t%))
'---- setfocus logic ----
'                     runaway(t%).SetFocus
          If RUNAWAY(t%).Visible Then
              RUNAWAY(t%).SetFocus
           End If
            GoTo exitedits
        End If
        End If
    End If
    If age(t%) > "" Then
        If Val(age(t%)) = 0 And age(t%) <> "00" Then
            msg = "Age must be entered in the format of a single age: X or XX (i.e. 3 or 42) or in the format of a range of ages: XXXX (i.e. 2025, meaning 20 - 25 years old)."
            Call ShowApplicableContainers(age(t%))
'---- setfocus logic ----
'                     age(t%).SetFocus
          If age(t%).Visible Then
              age(t%).SetFocus
           End If
            GoTo exitedits
        End If
    Else
        '===== error 504
        msg = "Subject age must be entered. (00 = unknown)"
        Call ShowApplicableContainers(age(t%))
'---- setfocus logic ----
'                 age(t%).SetFocus
          If age(t%).Visible Then
              age(t%).SetFocus
           End If
        GoTo exitedits
    End If
    '===== Error 504
    If sex(t%).ListIndex = -1 Then
        msg = "A value for Sex in subject data must be entered."
        Call ShowApplicableContainers(sex(t%))
'---- setfocus logic ----
'                 sex(t%).SetFocus
          If sex(t%).Visible Then
              sex(t%).SetFocus
           End If
        GoTo exitedits
    End If
    If race(t%).ListIndex = -1 Then
        msg = "A value for race in subject data must be entered."
        Call ShowApplicableContainers(race(t%))
'---- setfocus logic ----
'                 race(t%).SetFocus
          If race(t%).Visible Then
              race(t%).SetFocus
           End If
        GoTo exitedits
    End If
    If ethnicity(t%).ListIndex = -1 Then
        msg = "A value for ethnicity in subject data must be entered."
        Call ShowApplicableContainers(ethnicity(t%))
'---- setfocus logic ----
'                 ethnicity(t%).SetFocus
          If ethnicity(t%).Visible Then
              ethnicity(t%).SetFocus
           End If
        GoTo exitedits
    End If
    '===== Error 404,556
    If age(t%) <> "NN" And age(t%) <> "NB" And age(t%) <> "BB" Then
        For tt% = 1 To Len(age(t%))
            If InStr("0123456789-", Mid$(age(t%), tt%, 1)) = 0 Then
                msg = "An invalid entry in age has been found.  Valid Entry Formats:  X, XX, X-Y, XX-Y, X-YY, XX-YY, XXYY"
                tt% = Len(age(t%))
                Call ShowApplicableContainers(age(t%))
'---- setfocus logic ----
'                         age(t%).SetFocus
          If age(t%).Visible Then
              age(t%).SetFocus
           End If
                GoTo exitedits
            End If
        Next tt%
    End If
    If InStr(age(t%), "-") > 0 Then
        ag1 = Val(Left$(age(t%), InStr(age(t%), "-") - 1))
        ag2 = Val(Mid$(age(t%), InStr(age(t%), "-") + 1))
        age(t%) = Format$(ag1, "00") + Format$(ag2, "00")
    End If
    '===== Error 410,422,509,510,522
    If Len(age(t%)) = 4 Then
        If Val(Left$(age(t%), 2)) >= Val(Right$(age(t%), 2)) Then
            msg = "For an age range, the first age must be less than the second age."
            Call ShowApplicableContainers(age(t%))
'---- setfocus logic ----
'                     age(t%).SetFocus
          If age(t%).Visible Then
              age(t%).SetFocus
           End If
            GoTo exitedits
        End If
        If Val(Left$(age(t%), 2)) = 0 Then
            msg = "The low value in an age range cannot be 0."
            Call ShowApplicableContainers(age(t%))
'---- setfocus logic ----
'                     age(t%).SetFocus
          If age(t%).Visible Then
              age(t%).SetFocus
           End If
            GoTo exitedits
        End If
    End If
    '===== Data Element 20, 21, 22
    Dim drugs(2)
    drugs(0) = ""
    drugs(1) = ""
    drugs(2) = ""
    dct% = t% * 3
    For ii% = (t% * 3) To (t% * 3) + 2
        If drugtype(ii%).ListIndex > -1 Then
            drugs(dct%) = drugtype(ii%).List(drugtype(ii%).ListIndex)
            dct% = dct% + 1
        End If
    Next ii%
    If dct% > 0 And dct% > t% * 3 Then
        For Z% = (t% * 3) To dct%
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
                                Call ShowApplicableContainers(drugmeasurement(Z%))
'---- setfocus logic ----
'                                         drugmeasurement(Z%).SetFocus
          If drugmeasurement(Z%).Visible Then
              drugmeasurement(Z%).SetFocus
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
                                Call ShowApplicableContainers(drugmeasurement(Z%))
'---- setfocus logic ----
'                                         drugmeasurement(Z%).SetFocus
          If drugmeasurement(Z%).Visible Then
              drugmeasurement(Z%).SetFocus
           End If
                                GoTo exitedits
                        End If
                    End If
                    If InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "DU=") > 0 Or _
                       InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "NP=") > 0 Then
                        If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "DU=") > 0 Or _
                           InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "NP=") > 0 Then
                                msg = "Two different measurements within the same weight category cannot be entered for the same type of drug."
                                Call ShowApplicableContainers(drugmeasurement(Z%))
'---- setfocus logic ----
'                                         drugmeasurement(Z%).SetFocus
          If drugmeasurement(Z%).Visible Then
              drugmeasurement(Z%).SetFocus
           End If
                                GoTo exitedits
                        End If
                    End If
                End If
            Next ZZ%
        End If
            For ttt% = 1 To Len(drugamt(Z%))
                If InStr("0123456789.", Mid$(drugamt(Z%), ttt%, 1)) = 0 Then
                    msg = "Only 0123456789. are allowed in drug amount.  Enter all fractions as decimals (i.e. 1/2 = .5)."
                    Call ShowApplicableContainers(drugamt(Z%))
'---- setfocus logic ----
'                             drugamt(Z%).SetFocus
          If drugamt(Z%).Visible Then
              drugamt(Z%).SetFocus
           End If
                    GoTo exitedits
                End If
            Next ttt%
            If drugamt(Z%) = "" Or drugmeasurement(Z%).ListIndex = -1 Then
                If drugs(Z%) > "" And Left$(drugs(Z%), 1) <> "X" And Left$(drugs(Z%), 1) <> "U" Then
                    msg = "Drug Quantity and Measurement Type must be entered/selected."
                    Call ShowApplicableContainers(drugamt(Z%))
'---- setfocus logic ----
'                             drugamt(Z%).SetFocus
          If drugamt(Z%).Visible Then
              drugamt(Z%).SetFocus
           End If
                    GoTo exitedits
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
'                                 drugmeasurement(Z%).SetFocus
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
                        Call ShowApplicableContainers(drugmeasurement(Z%))
'---- setfocus logic ----
'                                 drugmeasurement(Z%).SetFocus
          If drugmeasurement(Z%).Visible Then
              drugmeasurement(Z%).SetFocus
           End If
                        GoTo exitedits
                    End If
                End If
                '===== Error 368
                If drugs(Z%) = "" Or drugamt(Z%) = "" Then
                    msg = "If a drug measurement is entered, then drug type and quantity must also be entered."
                    Call ShowApplicableContainers(drugamt(Z%))
'---- setfocus logic ----
'                             drugamt(Z%).SetFocus
          If drugamt(Z%).Visible Then
              drugamt(Z%).SetFocus
           End If
                    GoTo exitedits
                End If
            End If
            '===== Error 362
            If Left$(drugs(Z%), 1) = "X" Then
                If drugtype(t% * 3).ListIndex = -1 Or drugtype((t% * 3) + 1).ListIndex = -1 Or drugtype((t% * 3) + 2).ListIndex = -1 Then
                    msg = "If X = Over 3 Drug Types is entered, then the other 2 drug types must also be entered."
                    Call ShowApplicableContainers(drugtype(t% * 3))
'---- setfocus logic ----
'                             drugtype(t% * 3).SetFocus
          If drugtype(t% * 3).Visible Then
              drugtype(t% * 3).SetFocus
           End If
                    GoTo exitedits
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
                    GoTo exitedits
                End If
            End If
        Next Z%
    End If
    
    '==== Mandatories E - 37, 38, 39
    '===== Error 501
'    If (t% = 0 And (TV1 = "I" Or TV1 = "P")) Or (t% = 1 And (TV2 = "I" Or TV2 = "P")) Or UCase(vsname(t%)) <> "UNKNOWN" Then
    If UCase(vsname(t%)) <> "UNKNOWN" Then
        If Val(tage) = 0 And tage <> "00" Then
            msg = "Invalid age entered."
            Call ShowApplicableContainers(vsname(t%))
'---- setfocus logic ----
'                     vsname(t%).SetFocus
          If vsname(t%).Visible Then
              vsname(t%).SetFocus
           End If
            GoTo exitedits
        End If
        If sex(t%).ListIndex = -1 Then
            msg = "Invalid sex entered."
            Call ShowApplicableContainers(sex(t%))
'---- setfocus logic ----
'                     sex(t%).SetFocus
          If sex(t%).Visible Then
              sex(t%).SetFocus
           End If
            GoTo exitedits
        End If
        If race(t%).ListIndex = -1 Then
            msg = "Invalid race entered."
            Call ShowApplicableContainers(race(t%))
'---- setfocus logic ----
'                     race(t%).SetFocus
          If race(t%).Visible Then
              race(t%).SetFocus
           End If
            GoTo exitedits
        End If
        If ethnicity(t%).ListIndex = -1 Then
            msg = "Invalid ETHNICITY entered."
            Call ShowApplicableContainers(ethnicity(t%))
'---- setfocus logic ----
'                     ethnicity(t%).SetFocus
          If ethnicity(t%).Visible Then
              ethnicity(t%).SetFocus
           End If
            GoTo exitedits
        End If
    End If
    
    If (rs2("offenderdeath") Or rs2("noprosecution") Or rs2("extraditiondenied") Or rs2("victimdeclinescooperation") Or rs2("juvenilenocustody")) Then
        If age(t%) = "00" Then
            msg = "For an exceptional clearance, the subjects age (other than 00) must be selected."
            Call ShowApplicableContainers(age(t%))
'---- setfocus logic ----
'                     age(t%).SetFocus
          If age(t%).Visible Then
              age(t%).SetFocus
           End If
            GoTo exitedits
        End If
        If race(t%).ListIndex = -1 Or race(t%).List(race(t%).ListIndex) = "Unknown" Then
            msg = "For an exceptional clearance, the subject's race (other than unknown) must be selected."
            Call ShowApplicableContainers(race(t%))
'---- setfocus logic ----
'                     race(t%).SetFocus
          If race(t%).Visible Then
              race(t%).SetFocus
           End If
            GoTo exitedits
        End If
        If sex(t%).ListIndex = -1 Or sex(t%).List(sex(t%).ListIndex) = "Unknown" Then
            msg = "For an exceptional clearance, the subject's sex (other than unknown) must be selected."
            Call ShowApplicableContainers(sex(t%))
'---- setfocus logic ----
'                     sex(t%).SetFocus
          If sex(t%).Visible Then
              sex(t%).SetFocus
           End If
            GoTo exitedits
        End If
        If ethnicity(t%).ListIndex = -1 Or ethnicity(t%).List(ethnicity(t%).ListIndex) = "Unknown" Then
            msg = "For an exceptional clearance, the subject's ethnicity (other than unknown) must be selected."
            Call ShowApplicableContainers(ethnicity(t%))
'---- setfocus logic ----
'                     ethnicity(t%).SetFocus
          If ethnicity(t%).Visible Then
              ethnicity(t%).SetFocus
           End If
            GoTo exitedits
        End If
    End If

snextt:
Next t%
GoTo goodedits
exitedits:
editerr = 1
goodedits:
db.Close
On Error Resume Next
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
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


Private Sub TWOMANVEHICLE_GotFocus(Index As Integer)
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub unfounded_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub victim_LostFocus(Index As Integer)
If Val(victim(Index)) = 1 Then
    msg = MsgBox("Victim Number must be greater than 1.", 48, "Genesis Error Log")
End If
If Val(victim(Index)) > 1 Then
    If Index = 0 Then
        inp = InputBox("Enter Type of Victim (I=Individual/B=Business/F=Financial Inst./G=Government/R=Relig. Org./S=Society/O=Other/U=Unknown/P=Police Officer", "", TV1)
    Else
        inp = InputBox("Enter Type of Victim (I=Individual/B=Business/F=Financial Inst./G=Government/R=Relig. Org./S=Society/O=Other/U=Unknown/P=Police Officer", "", TV2)
    End If
    inp = UCase(inp)
    If InStr("IBFGRSOUP", inp) = 0 Then
        msg = MsgBox("Invalid Type of Victim entered.", 48, "Genesis Error Log")
    End If
    If Index = 0 Then
        TV1 = inp
    Else
        TV2 = inp
    End If
End If
End Sub

Private Sub victimdeclinescooperation_GotFocus()
If Frame18.Top > (-1 * Picture2.Top) And Frame18.Top + Frame18.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Frame18.Top > 500 Then
    VScroll1 = Frame18.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub VIN_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub VISIBLEINJURYNO_GotFocus(Index As Integer)
If Frame1(Index).Top > (-1 * Picture2.Top) And Frame1(Index).Top + Frame1(Index).Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Frame1(Index).Top > 500 Then
    VScroll1 = Frame1(Index).Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub VISIBLEINJURYYES_GotFocus(Index As Integer)
If Frame1(Index).Top > (-1 * Picture2.Top) And Frame1(Index).Top + Frame1(Index).Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Frame1(Index).Top > 500 Then
    VScroll1 = Frame1(Index).Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub VScroll1_Change()
Picture2.Top = -VScroll1.Value
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
If UCase(vsname(Index)) <> "SAME AS VICTIM" And vsname(Index) > "" And InStr(vsname(Index), ",") = 0 Then
    If (Index = 0 And TV1 <> "I") Or (Index = 1 And TV2 <> "I") Then
    Else
        msg = MsgBox("All names in the Incident Report system should be entered in the format last name + comma + firstname.", 48, "Invalid Data Format")
'---- setfocus logic ----
'                 vsname(index).SetFocus
          If vsname(Index).Visible Then
              vsname(Index).SetFocus
           End If
    End If
End If
If vsname(Index) > "" Then
    Call FILLDATA(Index)
End If

End Sub


Private Sub vucrlist_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
For p% = 1 To 5
    VUCRSEL(p%, Index) = ""
Next p%
vidx% = 0
For p% = 1 To vucrlist(Index).ListItems.Count
    If vucrlist(Index).ListItems(p%).Selected Then
        vidx% = vidx% + 1
        If vidx% < 6 Then
            VUCRSEL(vidx%, Index) = Mid(vucrlist(Index).ListItems(p%), InStr(vucrlist(Index).ListItems(p%), "(") + 1, 3)
        End If
    End If
Next p%
Call poppucr
End Sub

Private Sub vucrlist_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
For p% = 1 To 5
    VUCRSEL(p%, Index) = ""
Next p%
vidx% = 0
For p% = 1 To vucrlist(Index).ListItems.Count
    If vucrlist(Index).ListItems(p%).Selected Then
        vidx% = vidx% + 1
        If vidx% < 6 Then
            VUCRSEL(vidx%, Index) = Mid(vucrlist(Index).ListItems(p%), InStr(vucrlist(Index).ListItems(p%), "(") + 1, 3)
        End If
    End If
Next p%
End Sub

Private Sub WANTED_GotFocus(Index As Integer)
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If

End Sub

Private Sub WARRANT_GotFocus(Index As Integer)
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If

End Sub


Friend Sub clearroutine(TP As Integer)
For pp% = 1 To 5
    VUCRSEL(pp%, 0) = ""
    VUCRSEL(pp%, 1) = ""
Next pp%
subjectidentifiedyes = False
subjectidentifiedno = True
subjectlocatedyes = False
subjectlocatedno = True
arrestedunder18 = 0
arrested18andover = 0
exclearunder18 = 0
exclear18andover = 0
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
mugshot(0).Picture = LoadPicture()
mugshot(1).Picture = LoadPicture()
Dim itmx As ListItem
TV1 = ""
TV2 = ""
For t% = 0 To 1
    alcoholyes(t%) = 0
    alcoholno(t%) = 0
    alcoholunknown(t%) = 1
    drugsyes(t%) = 0
    drugsno(t%) = 0
    drugsunknown(t%) = 1
    'vucrlist(t%).ListItems.Clear
Next t%
For t% = 0 To 5
    DATERECOVERED(t%).Visible = False
    group(t%).Visible = False
    group(t%).ListIndex = -1
    numvehicle(t%).Visible = False
    numvehicle(t% + 6).Visible = False
    pucrlist(t%).Visible = False
    'pucrlist(t%).clear
    pucrlist(t%).ListIndex = -1
    pinfoframe(t%).Visible = False
Next t%
'ian's code
NARRATIVE.Text = ""
'end ian's code
reportingofficer(0) = ""
REPORTINGOFFICERDATE(0) = ""
reportingofficeRunit(0) = ""
reportingofficer(1) = ""
REPORTINGOFFICERDATE(1) = ""
reportingofficeRunit(1) = ""
approvingofficer = ""
APPROVINGOFFICERDATE = ""
approvingofficeRunit = ""
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
relationshipframe(0).Visible = False
relationshipframe(1).Visible = False
For t% = 0 To 9
    sdrugframe(t%).Visible = False
Next t%
vucrf(0).Visible = False
vucrf(1).Visible = False
LOCATIONNUMBER(0) = ""
LOCATIONNUMBER(1) = ""
For t% = 0 To 5
    group(t%).ListIndex = -1
    totalvalue(t%) = ""
    description(t%) = ""
    numvehicle(t%) = ""
    DATERECOVERED(t%) = ""
Next t%
stolen = 0
damaged = 0
burned = 0
recovered = 0
seized = 0
counterfeited = 0
typeunknown = 0
vsname(0) = ""
vsname(1) = ""
For t% = 0 To 1
    address(t%) = ""
    city(t%) = ""
    state(t%) = ""
    zipcode(t%) = ""
    race(t%).ListIndex = -1
    sex(t%).ListIndex = -1
    age(t%) = ""
    ethnicity(t%).ListIndex = -1
    BIRTHDATE(Index) = ""
    ht(t%) = ""
    weight(t%) = ""
    resident(t%).ListIndex = -1
    hair(t%) = ""
    eyes(t%) = ""
    peculiarities(t%) = ""
    VISIBLEINJURYYES(t%) = False
    VISIBLEINJURYNO(t%) = True
    NONVISIBLEINJURYYES(t%) = False
    NONVISIBLEINJURYNO(t%) = True
    TWOMANVEHICLE(t%) = 0
    ONEMANVEHICLE(t%) = 0
    DETECTIVE(t%) = 0
    TODOTHER(t%) = 0
    ALONE(t%) = 0
    ASSISTED(t%) = 0
    For tt% = 1 To injury(t%).ListItems.Count
        injury(t%).ListItems(tt%).Selected = False
    Next tt%
Next t%
WARRANT(0) = 0
WANTED(0) = 0
RUNAWAY(0) = 0
ARREST(0) = 0
JAIL(0) = 0
SUMMONS(0) = 0
WARRANT(1) = 0
WANTED(1) = 0
RUNAWAY(1) = 0
ARREST(1) = 0
JAIL(1) = 0
SUMMONS(1) = 0
For t% = 0 To 29
    drugtype(t%).ListIndex = -1
    drugmeasurement(t%).ListIndex = -1
    drugamt(t%) = ""
Next t%
SSTOLEN = 0
SRECOVERED = 0
SFOUND = 0
STOWED = 0
SSUSPECT = 0
SVICTIM = 0
SVEHICLE = 0
SGUN = 0
SBOAT = 0
sLICENSEPLATE = 0
SSECURITIES = 0
SARTICLE = 0
VIN = ""
HULL = ""
SERIAL = ""
SERIALSTATE = ""
YEARREG = ""
YEAREXP = ""
YEARN = ""
make = ""
stype = ""
model = ""
STYLE = ""
scolor = ""
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
Next t%
ORiGINAL = 0
narrativeonly = 0
MODIFIES = 0
SUPPLEMENTAL = 0
CASEst = 0
ADDITIONALV = 0
additionalo = 0
ADDITIONALS = 0
ADDITIONALR = 0
For t% = 0 To 1
    complainant(t%) = 0
    victim(t%) = ""
    subject(t%) = ""
Next t%
For t% = 0 To 3
    alcoholyes(t%) = False
    alcoholno(t%) = True
    alcoholunknown(t%) = False
    drugsyes(t%) = False
    drugsno(t%) = True
    drugsunknown(t%) = False
Next t%


VScroll1 = 0
If TP = 0 Then
    Call defaultcodes
End If
End Sub
Private Sub defaultcodes()
Dim db As Database, rs As Recordset, itmx As ListItem
On Error GoTo oderror
od:
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("select * from codes")
If rs.EOF Then
    On Error Resume Next
    db.Close
    Exit Sub
End If
rs.MoveFirst
On Error Resume Next
For t% = 0 To 1
    state(t%).clear
    city(t%).clear
    sex(t%).clear
    race(t%).clear
    ethnicity(t%).clear
    resident(t%).clear
    injury(t%).ListItems.clear
Next t%
While Not rs.EOF
    Select Case rs("type")
        Case "state"
            For t% = 0 To 1
                state(t%).AddItem rs("code")
               'If UCase(rs("default")) = "Y" Then
               '     state(t%).ListIndex = state(t%).ListCount - 1
               ' End If
            Next t%
        Case "city"
            For t% = 0 To 1
                city(t%).AddItem rs("code")
                'If UCase(rs("default")) = "Y" Then
                '    city(t%).ListIndex = city(t%).ListCount - 1
                'End If
            Next t%
        Case "injury"
            For t% = 0 To 1
                Set itmx = injury(t%).ListItems.add(, , rs("code"))
                If UCase(rs("default")) = "Y" Then
                    injury(t%).ListItems(injury(t%).ListItems.Count).Selected = True
                    injury(t%).ListItems(injury(t%).ListItems.Count).EnsureVisible
                Else
                    injury(t%).ListItems(injury(t%).ListItems.Count).Selected = False
                End If
            Next t%
        Case "sex"
            For t% = 0 To 1
                sex(t%).AddItem rs("code")
                If UCase(rs("default")) = "Y" Then
                    sex(t%).ListIndex = sex(t%).ListCount - 1
                End If
            Next t%
        Case "race"
            For t% = 0 To 1
                race(t%).AddItem rs("code")
                If UCase(rs("default")) = "Y" Then
                    race(t%).ListIndex = race(t%).ListCount - 1
                End If
            Next t%
        Case "ethnicity"
            For t% = 0 To 1
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
Friend Sub editproperty(editerr As Integer, msg As String)
Dim itmx As ListItem, db As Database, rs, rs2 As Recordset
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("Select dateofoffense2 from incidentreportc where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
If Not rs.EOF Then
    If IsDate(rs("dateofoffense2")) Then
        offensedate = rs("dateofoffense2")
    Else
        msg = "Unable to find incident date."
        GoTo exiteditp
    End If
Else
    msg = "Unable to find incident date."
    GoTo exiteditp
End If
    
'RLB Bandaid
On Error GoTo rlbErr
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
    GOTONE = False
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
                                        Call ShowApplicableContainers(drugmeasurement(Z%))
'---- setfocus logic ----
'                                                 drugmeasurement(Z%).SetFocus
          If drugmeasurement(Z%).Visible Then
              drugmeasurement(Z%).SetFocus
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
                                        Call ShowApplicableContainers(drugmeasurement(Z%))
'---- setfocus logic ----
'                                                 drugmeasurement(Z%).SetFocus
          If drugmeasurement(Z%).Visible Then
              drugmeasurement(Z%).SetFocus
           End If
                                        GoTo exiteditp
                                End If
                        End If
                        If InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "DU=") > 0 Or _
                            InStr(drugmeasurement(ZZ%).List(drugmeasurement(ZZ%).ListIndex), "NP=") > 0 Then
                                If InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "DU=") > 0 Or _
                                    InStr(drugmeasurement(Z%).List(drugmeasurement(Z%).ListIndex), "NP=") > 0 Then
                                        msg = "Two different measurements within the same weight category cannot be entered for the same type of drug."
                                        Call ShowApplicableContainers(drugmeasurement(Z%))
'---- setfocus logic ----
'                                                 drugmeasurement(Z%).SetFocus
          If drugmeasurement(Z%).Visible Then
              drugmeasurement(Z%).SetFocus
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
                    Call ShowApplicableContainers(drugamt(Z%))
'---- setfocus logic ----
'                             drugamt(Z%).SetFocus
          If drugamt(Z%).Visible Then
              drugamt(Z%).SetFocus
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
    'Data Element 16
    If group(d%).ListIndex > -1 Then
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
                    Call ShowApplicableContainers(totalvalue(d% + 24))
'---- setfocus logic ----
'                             totalvalue(d%).SetFocus
          If totalvalue(d% + 24).Visible Then
              totalvalue(d% + 24).SetFocus
           End If
                    GoTo exiteditp
                Else
                If ctx% > 1 Then
                    msg = "For Credit/Debit Cards and Nonnegotiable Instruments, no value is allowed, so an X (only 1) must be entered to show the nature of the crime (stolen, recovered, etc.)"
                    Call ShowApplicableContainers(totalvalue(d%))
'---- setfocus logic ----
'                             totalvalue(d%).SetFocus
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
'                                 totalvalue(d%).SetFocus
          If totalvalue(d%).Visible Then
              totalvalue(d%).SetFocus
           End If
                        GoTo exiteditp
                    Else
                    If ctx% > 1 Then
                        msg = "For Other and Special Category, if no value is entered, an X (only 1) must be entered to show the nature of the crime (stolen, recovered, etc.)"
                        Call ShowApplicableContainers(totalvalue(d%))
'---- setfocus logic ----
'                                 totalvalue(d%).SetFocus
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
'                                     totalvalue(d%).SetFocus
          If totalvalue(d%).Visible Then
              totalvalue(d%).SetFocus
           End If
                            GoTo exiteditp
                        End If
                        If ctv% = 0 Then
                            msg = "An amount must be entered for group 10 (Drug/Narcotics) when the UCR selected is not 35A (Drug/Narcotic Violations)."
                            Call ShowApplicableContainers(totalvalue(d%))
'---- setfocus logic ----
'                                     totalvalue(d%).SetFocus
          If totalvalue(d%).Visible Then
              totalvalue(d%).SetFocus
           End If
                            GoTo exiteditp
                        End If
                    Else
                        Select Case d%
                            Case 0
                                didx% = 12
                            Case 1
                                didx% = 15
                            Case 2
                                didx% = 18
                            Case 3
                                didx% = 21
                            Case 4
                                didx% = 24
                            Case 5
                                didx% = 27
                        End Select
                        If drugtype(didx%).ListIndex = -1 Or Val(drugamt(didx%)) = 0 Or drugmeasurement(didx%).ListIndex = -1 Then
                            msg = "Drug information must be entered group 10 (Drug/Narcotics) when the UCR selected is 35A (Drug/Narcotic Violations)."
                            GoTo exiteditp
                        End If
                        If ctx% = 0 Then
                            msg = "An X must be entered for group 10 (Drug/Narcotics) when the UCR selected is 35A (Drug/Narcotic Violations)."
                            Call ShowApplicableContainers(totalvalue(d% + 24))
'---- setfocus logic ----
'                                     totalvalue(d%).SetFocus
          If totalvalue(d% + 24).Visible Then
              totalvalue(d% + 24).SetFocus
           End If
                            GoTo exiteditp
                        End If
                        If ctx% > 1 Then
                            msg = "An X (only 1) must be entered for group 10 (Drug/Narcotics) when the UCR selected is 35A (Drug/Narcotic Violations)."
                            Call ShowApplicableContainers(totalvalue(d% + 24))
'---- setfocus logic ----
'                                     totalvalue(d%).SetFocus
          If totalvalue(d% + 24).Visible Then
              totalvalue(d% + 24).SetFocus
           End If
                            GoTo exiteditp
                        End If
                        If ctv% > 0 Then
                            msg = "No amount must be entered for group 10 (Drug/Narcotics) when the UCR selected is 35A (Drug/Narcotic Violations). Use an X to note the type of property crime."
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
            Case Else
                For dd% = 0 To 6
                    If totalvalue(d% + (dd% * 6)) = "X" Then
                        msg = "An entry of X is only allowed for groups 09, 22, 77, and 99."
                        Call ShowApplicableContainers(totalvalue(d% + (dd% * 6)))
'---- setfocus logic ----
'                                 totalvalue(d% + (dd% * 6)).SetFocus
          If totalvalue(d% + (dd% * 6)).Visible Then
              totalvalue(d% + (dd% * 6)).SetFocus
           End If
                        GoTo exiteditp
                    End If
                Next dd%
        End Select
    End If
    If Not description(d%).Text = "" Then
        FOUNDPROP = True
        If pucrlist(d%).ListIndex = -1 And pucrlist(d%).ListCount > 0 Then
            msg = "A UCR must be associated with the property described."
            Call ShowApplicableContainers(pucrlist(d%))
'---- setfocus logic ----
'                     pucrlist(d%).SetFocus
          If pucrlist(d%).Visible Then
              pucrlist(d%).SetFocus
           End If
            GoTo exiteditp
        Else
            If group(d%).ListIndex = -1 And pucrlist(d%).ListCount > 0 Then
                msg = "A Group must be associated with the property described."
                Call ShowApplicableContainers(group(d%))
'---- setfocus logic ----
'                         group(d%).SetFocus
          If group(d%).Visible Then
              group(d%).SetFocus
           End If
                GoTo exiteditp
            End If
        End If
    End If
Next d%
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
                    Case "200", "510", "220", "250", "35A", "35B", "290", "270", "210", "26A", "26B", "26C", "26D", "26E", "23A", "23B", "23C", "23D", "23E", "23F", "23G", "23H", "100", "240", "120", "39A", "39B", "39C", "39D", "280"
                    Case "90A", "90P", "90B", "90C", "90D", "90E", "90F", "90K", "90G", "90H", "90N", "90I", "90J", "90L", "90Z", ""
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
        
    tt% = t% Mod 6
            
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
    If Val(totalvalue(tt% + 30)) > 0 Then
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
                Call ShowApplicableContainers(DATERECOVERED(tt%))
'---- setfocus logic ----
'                         DATERECOVERED(tt%).SetFocus
          If DATERECOVERED(tt%).Visible Then
              DATERECOVERED(tt%).SetFocus
           End If
                GoTo exiteditp
            Else
                If CDate(DATERECOVERED(tt%)) < offensedate Then
                    msg = "Recovery date cannot be befroe incident date."
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
                            Call ShowApplicableContainers(group(uu%))
'---- setfocus logic ----
'                                     group(uu%).SetFocus
          If group(uu%).Visible Then
              group(uu%).SetFocus
           End If
                            GoTo exiteditp
                    End Select
                Case "23C"
                    Select Case Val(Mid$(group(uu%).List(group(uu%).ListIndex), InStr(group(uu%).List(group(uu%).ListIndex), "(") + 1, 2))
                        Case 1, 3, 5, 12, 15, 18, 24, 28, 29, 30, 31, 32, 33, 34, 35, 37, 39
                            msg = "Illogical property group/ucr combination."
                            Call ShowApplicableContainers(group(uu%))
'---- setfocus logic ----
'                                     group(uu%).SetFocus
          If group(uu%).Visible Then
              group(uu%).SetFocus
           End If
                            GoTo exiteditp
                    End Select
                Case "23F", "23D", "23E"
                    Select Case Val(Mid$(group(uu%).List(group(uu%).ListIndex), InStr(group(uu%).List(group(uu%).ListIndex), "(") + 1, 2))
                        Case 3, 5, 24, 28, 29, 30, 31, 32, 33, 34, 35, 37
                            msg = "Illogical property group/ucr combination."
                            Call ShowApplicableContainers(group(uu%))
'---- setfocus logic ----
'                                     group(uu%).SetFocus
          If group(uu%).Visible Then
              group(uu%).SetFocus
           End If
                            GoTo exiteditp
                    End Select
                Case "23G"
                    Select Case Val(Mid$(group(uu%).List(group(uu%).ListIndex), InStr(group(uu%).List(group(uu%).ListIndex), "(") + 1, 2))
                        Case 38, 88
                        Case Else
                            msg = "Illogical property group/ucr combination."
                            Call ShowApplicableContainers(group(uu%))
'---- setfocus logic ----
'                                     group(uu%).SetFocus
          If group(uu%).Visible Then
              group(uu%).SetFocus
           End If
                            GoTo exiteditp
                    End Select
                Case "23H"
                    Select Case Val(Mid$(group(uu%).List(group(uu%).ListIndex), InStr(group(uu%).List(group(uu%).ListIndex), "(") + 1, 2))
                        Case 3, 5, 24, 28, 37
                            msg = "Illogical property group/ucr combination."
                            Call ShowApplicableContainers(group(uu%))
'---- setfocus logic ----
'                                     group(uu%).SetFocus
          If group(uu%).Visible Then
              group(uu%).SetFocus
           End If
                            GoTo exiteditp
                    End Select
            End Select
            If Mid$(pucrlist(uu%).List(pucrlist(uu%).ListIndex), InStr(pucrlist(uu%).List(pucrlist(uu%).ListIndex), "(") + 1, 3) = "35A" Then
                If (Val(totalvalue(uu% + 0)) = 0 And Val(totalvalue(uu% + 6)) = 0 And Val(totalvalue(uu% + 12)) = 0 And Val(totalvalue(uu% + 18)) = 0 And Val(totalvalue(uu% + 24)) = 0 And Val(totalvalue(uu% + 30)) = 0) Then
                    If drugtype((uu% * 3) + 12).ListIndex = -1 Then
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
                                Call ShowApplicableContainers(group(uu%))
'---- setfocus logic ----
'                                         group(uu%).SetFocus
          If group(uu%).Visible Then
              group(uu%).SetFocus
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
            Call ShowApplicableContainers(totalvalue(tt%))
'---- setfocus logic ----
'                     totalvalue(tt%).SetFocus
          If totalvalue(tt%).Visible Then
              totalvalue(tt%).SetFocus
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
            msg = "A property value of 0 is only allowed for Credit/Debit Cards, Nonnegotiable Instruments, Other, Drug/Narcotics, and Special Category."
            Call ShowApplicableContainers(totalvalue(tt%))
'---- setfocus logic ----
'                     totalvalue(tt%).SetFocus
          If totalvalue(tt%).Visible Then
              totalvalue(tt%).SetFocus
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
        If CVDate(DATERECOVERED(tt%)) < CVDate(incidentdate) Then
            msg = "Date Recovered cannot be earlier that Date of Offense."
            Call ShowApplicableContainers(DATERECOVERED(tt%))
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
            Call ShowApplicableContainers(totalvalue(tt%))
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
            Call ShowApplicableContainers(drugtype((tt% + 2)))
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
On Error GoTo oderror
od:
Set db = OpenDatabase(nwi + "INCIDENT.MDB")
Set rs = db.OpenRecordset("SELECT UCR1, UCR2, UCR3, UCR4, UCR5, UCR6, UCR7, UCR8, UCR9, UCR10 FROM INCIDENTSUPPORT WHERE INCIDENTNUMBER = '" + incidentnumber + "'")
Set rs2 = db.OpenRecordset("SELECT * FROM INCIDENTREPORTC WHERE INCIDENTNUMBER = '" + incidentnumber + "'")
If Not rs.EOF Then
    rs.MoveFirst
Else
    msg = "Invalid incident report data."
    Call ShowApplicableContainers(incident)
'---- setfocus logic ----
'             incident.SetFocus
          If incident.Visible Then
              incident.SetFocus
           End If
    GoTo exiteditp
End If
If Not rs2.EOF Then
    rs2.MoveFirst
Else
    msg = "Invalid incident report data."
    Call ShowApplicableContainers(incident)
'---- setfocus logic ----
'             incident.SetFocus
          If incident.Visible Then
              incident.SetFocus
           End If
    GoTo exiteditp
End If

For t% = 0 To 9
    
    If Not IsNull(rs("ucr" + Mid$(Str$(t% + 1), 2))) Then
        tempucr = rs("ucr" + Mid$(Str$(t% + 1), 2))
    
        '===Additional F 7
        '===== Data Element 7
        '===== Data Element 12
        If tempucr = "35A" Or tempucr = "35B" Then
            TP$ = "Drug Offenses"
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Not rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
                            If Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                msg = TP$ + " not completed must have associated value of None or Unknown on property tab."
                                Call ShowApplicableContainers(pucrlist(tt%))
'---- setfocus logic ----
'                                         pucrlist(tt%).SetFocus
          If pucrlist(tt%).Visible Then
              pucrlist(tt%).SetFocus
           End If
                                GoTo exiteditp
                            End If
                        End If
                        If rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
                            'GLENN
                            '===== Error 301
                            If (Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) = 0 Or Val(totalvalue(tt% + 30)) > 0) Then
                                If Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                    msg = TP$ + " completed must have associated information of Seized property entered."
                                    Call ShowApplicableContainers(totalvalue(tt% + 0))
'---- setfocus logic ----
'                                             totalvalue(tt% + 0).SetFocus
          If totalvalue(tt% + 0).Visible Then
              totalvalue(tt% + 0).SetFocus
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
                            If Not rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
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
                            If rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
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
                        If Not rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
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
                        If rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
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
                                Call ShowApplicableContainers(group(tt%))
'---- setfocus logic ----
'                                         group(tt%).SetFocus
          If group(tt%).Visible Then
              group(tt%).SetFocus
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
        
        '===Additional F 5
        '===== Data Element 12
        If tempucr = "250" Then
            TP$ = "Counterfeiting"
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Not rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
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
                        If rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
                            If Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 6)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) = 0 Or Val(totalvalue(tt% + 24)) = 0 Or Val(totalvalue(tt% + 30)) = 0 Then
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
                        If Not rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
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
                        If rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
                            If Val(totalvalue(tt% + 6)) = 0 Or Val(totalvalue(tt% + 0)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 24)) > 0 Or Val(totalvalue(tt% + 30)) > 0 Then
                                msg = TP$ + " completed must have associated value of Damaged on property tab."
                                Call ShowApplicableContainers(totalvalue(tt% + 6))
'---- setfocus logic ----
'                                         totalvalue(tt% + 6).SetFocus
          If totalvalue(tt% + 6).Visible Then
              totalvalue(tt% + 6).SetFocus
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
                '    msg ="Valid property must be entered."
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
            Case "270", "200", "220", "250", "290", "210", "26A", "26B", "26C", "26D", "26E", "23A", "23B", "23C", "23D", "23E", "23F", "23G", "23H", "240", "120", "280"
                TP$ = "Crimes Against Property"
            Case "39A", "39B", "39C", "39D"
                TP$ = "Gambling"
            Case "100"
                TP$ = "Kidnaping"
            Case "35A", "35B"
                TP$ = "Drug/Narcotic Offenses"
        End Select
        '===== Error 074
'        If TP$ > "" Then
'            If FOUNDPROP = False Then
'                msg =TP$ + " must have property data."
'                GoTo exiteditp
'            End If
'        End If
        If TP$ > "" Then
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Not rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
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
                        If rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
                            If Val(totalvalue(tt%)) > 0 Or Val(totalvalue(tt% + 18)) > 0 Or Val(totalvalue(tt% + 12)) > 0 Then
                                temperr = 0
                            Else
'                            If tempucr <> "35A" Then
'                                msg =TP$ + " completed must have associated value of Burned, Recovered, or Stolen on property tab."
'                                GoTo exiteditp
'                            End If
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
                        If Val(totalvalue(tt% + 18)) > 0 Then
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
                    End If
                End If
            Next tt%
        End If
        'End If
        
        '===== Additional F 11
        '===== Data Element 7
        If tempucr = "39A" Or tempucr = "39B" Or tempucr = "39C" Or tempucr = "39D" Then
            TP$ = "Gambling Offenses"
            For tt% = 0 To 5
                If pucrlist(tt%).ListIndex > -1 Then
                    If Mid$(pucrlist(tt%).List(pucrlist(tt%).ListIndex), InStr(pucrlist(tt%).List(pucrlist(tt%).ListIndex), "(") + 1, 3) = tempucr Then
                        If Not rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
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
                        If rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
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
                        If Not rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
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
                        If rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
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
                                Call ShowApplicableContainers(DATERECOVERED(tt%))
'---- setfocus logic ----
'                                         DATERECOVERED(tt%).SetFocus
          If DATERECOVERED(tt%).Visible Then
              DATERECOVERED(tt%).SetFocus
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
                        If Not rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
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
                        If rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
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
                                Call ShowApplicableContainers(group(tt%))
'---- setfocus logic ----
'                                         group(tt%).SetFocus
          If group(tt%).Visible Then
              group(tt%).SetFocus
           End If
                                GoTo exiteditp
                            End If
                            tempgroup = Mid$(group(tt%).List(group(tt%).ListIndex), InStr(group(tt%).List(group(tt%).ListIndex), "(") + 1, 2)
                            If tempgroup = "03" Or tempgroup = "05" Or tempgroup = "24" Or tempgroup = "28" Or tempgroup = "37" Then
                                If numvehicle(tt% + 6) = 0 Then
                                    msg = "A number of vehicles recovered must be entered."
                                    Call ShowApplicableContainers(numvehicle(tt% + 6))
'---- setfocus logic ----
'                                             numvehicle(tt% + 6).SetFocus
          If numvehicle(tt% + 6).Visible Then
              numvehicle(tt% + 6).SetFocus
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
                                Call ShowApplicableContainers(group(tt%))
'---- setfocus logic ----
'                                         group(tt%).SetFocus
          If group(tt%).Visible Then
              group(tt%).SetFocus
           End If
                                GoTo exiteditp
                            End If
                            tempgroup = Mid$(group(tt%).List(group(tt%).ListIndex), InStr(group(tt%).List(group(tt%).ListIndex), "(") + 1, 2)
                            If tempgroup = "03" Or tempgroup = "05" Or tempgroup = "24" Or tempgroup = "28" Or tempgroup = "37" Then
                                If numvehicle(tt%) = 0 Then
                                    msg = "A number of vehicles stolen must be entered."
                                    Call ShowApplicableContainers(numvehicle(tt%))
'---- setfocus logic ----
'                                             numvehicle(tt%).SetFocus
          If numvehicle(tt%).Visible Then
              numvehicle(tt%).SetFocus
           End If
                                    GoTo exiteditp
                                End If
                            Else
                                msg = "Invalid type (group) entered for Motor Vehicle Theft crime."
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
                        If Not rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
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
                        If rs2("completedyES" + Mid$(Str$(t% + 1), 2)) Then
                            If (Val(totalvalue(tt%)) = 0 And Val(totalvalue(tt% + 6)) = 0 And Val(totalvalue(tt% + 12)) = 0 And Val(totalvalue(tt% + 18)) = 0 And Val(totalvalue(tt% + 24)) = 0 And Val(totalvalue(tt% + 30)) = 0) Or _
                                (Val(totalvalue(tt% + 18)) > 0 And Val(totalvalue(tt%)) = 0 And Val(totalvalue(tt% + 6)) = 0 And Val(totalvalue(tt% + 12)) = 0 And Val(totalvalue(tt% + 24)) = 0 And Val(totalvalue(tt% + 30)) = 0) Then
                                temperr = 0
                            Else
                                msg = TP$ + " completed must have associated value of None or Recovered on property tab."
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
                                Call ShowApplicableContainers(DATERECOVERED(tt%))
'---- setfocus logic ----
'                                         DATERECOVERED(tt%).SetFocus
          If DATERECOVERED(tt%).Visible Then
              DATERECOVERED(tt%).SetFocus
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
            If tempucr <> "35A" Then
                If Val(totalvalue(tt% + 36)) > 0 Or _
                    (Val(totalvalue(tt% + 0)) = 0 And Val(totalvalue(tt% + 6)) = 0 And Val(totalvalue(tt% + 12)) = 0 And Val(totalvalue(tt% + 18)) = 0 And Val(totalvalue(tt% + 24)) = 0 And Val(totalvalue(tt% + 30)) = 0) Then
                    If group(tt%).ListIndex > -1 Or pucrlist(tt%).ListIndex > -1 Or description(tt%) > "" Then
                        msg = "If UNKNOWN or Nothing selected in entry of PROPERTY tab, no other associated values may be selected (i.e. Type, Value, etc.)."
                        Call ShowApplicableContainers(group(tt%))
'---- setfocus logic ----
'                                 group(tt%).SetFocus
          If group(tt%).Visible Then
              group(tt%).SetFocus
           End If
                        GoTo exiteditp
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
On Error Resume Next
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
Exit Sub
'RLB Bandaid
rlbErr:
    If Err.Number = 5 Then Resume Next
End Sub

Private Sub deleteroutine()
If incidentnumber = "" Then
    Exit Sub
End If
If Len(incidentnumber) < 12 And Len(incidentnumber) > 0 Then
    incidentnumber = incidentnumber + Space$(12 - Len(incidentnumber))
End If
msg = MsgBox("Are you sure you wish to delete this supplemental incident report page?", 4, "Genesis Information Log")
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
Set rs = db.OpenRecordset("select * from supplemental where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34) + " and page = " + PAGE)
If Not rs.EOF Then
    rs.MoveFirst
End If
While Not rs.EOF
    rs.Delete
    rs.MoveNext
Wend
Set rs = db.OpenRecordset("select * from supplementalsupport where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34) + " and page = " + PAGE)
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

Friend Sub findincident(incfound As Boolean)
Dim db As Database, rs, rs2 As Recordset, ecc As Integer, lu As String
On Error GoTo oderror
od:
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("Select * from supplemental where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34) + " AND PAGE = " + PAGE)
If Not rs.EOF Then
    rs.MoveFirst
    On Error Resume Next
Else
    db.Close
    Exit Sub
End If
incfound = True
On Error Resume Next
fromfind = 1
incidentnumber = rs("incidentnumber")
LOCATIONNUMBER(0) = rs("locatioNnumber1")
LOCATIONNUMBER(1) = rs("locatioNnumber2")
vsname(0) = rs("name1")
address(0) = rs("address1")
city(0) = rs("city1")
state(0) = rs("state1")
zipcode(0) = rs("zipcode1")
computerequipment(0) = rs("computerequipment1")
computerequipment(1) = rs("computerequipment2")
'Ian's code
NARRATIVE.Text = rs("narrative")
'end Ian's code
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
If Not IsNull(rs("approvingofficer")) Then
    approvingofficer = rs("approvingofficer")
End If
If Not IsNull(rs("approvingdate")) Then
    APPROVINGOFFICERDATE = rs("approvingdate")
End If
If Not IsNull(rs("approvingunit")) Then
    approvingofficeRunit = rs("approvingunit")
End If
TV1 = ""
TV2 = ""
If rs("individual1") Then
    TV1 = "I"
End If
If rs("business1") Then
    TV1 = "B"
End If
If rs("financialinstitution1") Then
    TV1 = "F"
End If
If rs("government1") Then
    TV1 = "G"
End If
If rs("religiousorganization1") Then
    TV1 = "R"
End If
If rs("societypublic1") Then
    TV1 = "S"
End If
If rs("tvother1") Then
    TV1 = "O"
End If
If rs("unknown1") Then
    TV1 = "U"
End If
If rs("policeofficer1") Then
    TV1 = "P"
End If
If rs("individual2") Then
    TV2 = "I"
End If
If rs("business2") Then
    TV2 = "B"
End If
If rs("financialinstitution2") Then
    TV2 = "F"
End If
If rs("government2") Then
    TV2 = "G"
End If
If rs("religiousorganization2") Then
    TV2 = "R"
End If
If rs("societypublic2") Then
    TV2 = "S"
End If
If rs("tvother2") Then
    TV2 = "O"
End If
If rs("unknown2") Then
    TV2 = "U"
End If
If rs("policeofficer2") Then
    TV2 = "P"
End If
resident(0).ListIndex = -1
race(0).ListIndex = -1
sex(0).ListIndex = -1
ethnicity(0).ListIndex = -1
For t% = 0 To resident(0).ListCount - 1
    If Left$(resident(0).List(t%), 1) = rs("resident1") Then
        resident(0).ListIndex = t%
        t% = resident(0).ListCount - 1
    End If
Next t%
For t% = 0 To race(0).ListCount - 1
    If Left$(race(0).List(t%), 1) = rs("RACE1") Then
        race(0).ListIndex = t%
        t% = race(0).ListCount - 1
    End If
Next t%
For t% = 0 To sex(0).ListCount - 1
    If Left$(sex(0).List(t%), 1) = rs("SEX1") Then
        sex(0).ListIndex = t%
        t% = sex(0).ListCount - 1
    End If
Next t%
age(0) = rs("age1")
For t% = 0 To ethnicity(0).ListCount - 1
    If Left$(ethnicity(0).List(t%), 1) = rs("ETHNICITY1") Then
        ethnicity(0).ListIndex = t%
        t% = ethnicity(0).ListCount - 1
    End If
Next t%
For t% = 0 To 2
    For tt% = 0 To relationship(t%).ListCount - 1
        relationship(t%).Selected(tt%) = False
    Next tt%
    relationship(t%).ListIndex = -1
    If rs("relationship1" + Mid$(Str$(t% + 1), 2)) > "" Then
        For tt% = 0 To relationship(t%).ListCount - 1
            If rs("relationship1" + Mid$(Str$(t% + 1), 2)) = Mid$(relationship(t%).List(tt%), InStr(relationship(t%).List(tt%), "(") + 1, 2) Then
                relationship(t%).ListIndex = tt%
                relationship(t%).Selected(tt%) = True
                tt% = relationship(t%).ListCount - 1
            End If
        Next tt%
    End If
Next t%
HOMEDAYPHONE(0) = rs("homephoneDAY1")
WORKDAYPHONE(0) = rs("WORKphoneDAY1")
HOMENIGHTPHONE(0) = rs("HOMEphoneNIGHT1")
WORKNIGHTPHONE(0) = rs("WORKphoneNIGHT1")
For s% = 0 To 1
    For t% = 1 To injury(s%).ListItems.Count
        For tt% = 1 To 5
            If Not IsNull(rs("typeofinjury" + Mid$(Str$(s% + 1), 2) + Mid$(Str$(tt%), 2))) Then
                If rs("typeofinjury" + Mid$(Str$(s% + 1), 2) + Mid$(Str$(tt%), 2)) = Mid$(injury(s%).ListItems(t%), InStr(injury(s%).ListItems(t%), "(") + 1, 1) Then
                    injury(s%).ListItems(t%).Selected = True
                End If
            End If
        Next tt%
    Next t%
Next s%
If rs("visibleinjuryyes1") = "X" Then
    VISIBLEINJURYYES(0) = True
End If
If rs("visibleinjuryno1") = "X" Then
    VISIBLEINJURYNO(0) = True
End If
If rs("NONVISibleinjuryyes1") = "X" Then
    NONVISIBLEINJURYYES(0) = True
End If
If rs("NONVISibleinjuryno1") = "X" Then
    NONVISIBLEINJURYNO(0) = True
End If
ht(0) = rs("HEIGHT1")
weight(0) = rs("WEIGHT1")
hair(0) = rs("HAIR1")
eyes(0) = rs("EYES1")
peculiarities(0) = rs("PECULIARITIES1")
If rs("valcoholyes1") = "X" Then
    alcoholyes(0) = True
End If
If rs("valcoholNO1") = "X" Then
    alcoholno(0) = True
End If
If rs("valcoholUNKNOWN1") = "X" Then
    alcoholunknown(0) = True
End If
If rs("vDRUGSyes1") = "X" Then
    drugsyes(0) = True
End If
If rs("vDRUGSNO1") = "X" Then
    drugsno(0) = True
End If
If rs("vDRUGSUNKNOWN1") = "X" Then
    drugsunknown(0) = True
End If
If rs("salcoholyes1") = "X" Then
    alcoholyes(1) = True
End If
If rs("salcoholNO1") = "X" Then
    alcoholno(1) = True
End If
If rs("salcoholUNKNOWN1") = "X" Then
    alcoholunknown(1) = True
End If
If rs("sDRUGSyes1") = "X" Then
    drugsyes(1) = True
End If
If rs("sDRUGSNO1") = "X" Then
    drugsno(1) = True
End If
If rs("sDRUGSUNKNOWN1") = "X" Then
    drugsunknown(1) = True
End If


vsname(1) = rs("name2")
address(1) = rs("address2")
city(1) = rs("city2")
state(1) = rs("state2")
zipcode(1) = rs("zipcode2")
ht(1) = rs("HEIGHT2")
weight(1) = rs("WEIGHT2")
hair(1) = rs("HAIR2")
eyes(1) = rs("EYES2")
peculiarities(1) = rs("PECULIARITIES2")
resident(1).ListIndex = -1
race(1).ListIndex = -1
sex(1).ListIndex = -1
ethnicity(1).ListIndex = -1
For t% = 0 To resident(1).ListCount - 1
    If Left$(resident(1).List(t%), 1) = rs("resident2") Then
        resident(1).ListIndex = t%
        t% = resident(1).ListCount - 1
    End If
Next t%
For t% = 0 To race(1).ListCount - 1
    If Left$(race(1).List(t%), 1) = rs("RACE2") Then
        race(1).ListIndex = t%
        t% = race(1).ListCount - 1
    End If
Next t%
For t% = 0 To sex(1).ListCount - 1
    If Left$(sex(1).List(t%), 1) = rs("SEX2") Then
        sex(1).ListIndex = t%
        t% = sex(1).ListCount - 1
    End If
Next t%
age(1) = rs("age2")
For t% = 0 To ethnicity(1).ListCount - 1
    If Left$(ethnicity(1).List(t%), 1) = rs("ETHNICITY2") Then
        ethnicity(1).ListIndex = t%
        t% = ethnicity(1).ListCount - 1
    End If
Next t%
For t% = 10 To 12
    For tt% = 0 To relationship(t%).ListCount - 1
        relationship(t%).Selected(tt%) = False
    Next tt%
    relationship(t%).ListIndex = -1
    If rs("relationship2" + Mid$(Str$(t% - 9), 2)) > "" Then
        For tt% = 0 To relationship(t%).ListCount - 1
            If rs("relationship2" + Mid$(Str$(t% - 9), 2)) = Mid$(relationship(t%).List(tt%), InStr(relationship(t%).List(tt%), "(") + 1, 2) Then
                relationship(t%).ListIndex = tt%
                relationship(t%).Selected(tt%) = True
                tt% = relationship(t%).ListCount - 1
            End If
        Next tt%
    End If
Next t%
If Not IsNull(rs("homephoneDAY2")) Then
    HOMEDAYPHONE(1) = rs("homephoneDAY2")
End If
If Not IsNull(rs("workphoneDAY2")) Then
    WORKDAYPHONE(1) = rs("workphoneDAY2")
End If
If Not IsNull(rs("homephoneNIGHT2")) Then
    HOMENIGHTPHONE(1) = rs("homephoneNIGHT2")
End If
If Not IsNull(rs("workphoneNIGHT2")) Then
    WORKNIGHTPHONE(1) = rs("workphoneNIGHT2")
End If
If rs("visibleinjuryyes2") = "X" Then
    VISIBLEINJURYYES(1) = True
End If
If rs("visibleinjuryno2") = "X" Then
    VISIBLEINJURYNO(1) = True
End If
If rs("NONVISibleinjuryyes2") = "X" Then
    NONVISIBLEINJURYYES(1) = True
End If
If rs("NONVISibleinjuryno2") = "X" Then
    NONVISIBLEINJURYNO(1) = True
End If
If rs("valcoholyes2") = "X" Then
    alcoholyes(2) = True
End If
If rs("valcoholNO2") = "X" Then
    alcoholno(2) = True
End If
If rs("valcoholUNKNOWN2") = "X" Then
    alcoholunknown(2) = True
End If
If rs("vDRUGSyes2") = "X" Then
    drugsyes(2) = True
End If
If rs("vDRUGSNO2") = "X" Then
    drugsno(2) = True
End If
If rs("vDRUGSUNKNOWN2") = "X" Then
    drugsunknown(2) = True
End If
If rs("salcoholyes2") = "X" Then
    alcoholyes(3) = True
End If
If rs("salcoholNO2") = "X" Then
    alcoholno(3) = True
End If
If rs("salcoholUNKNOWN2") = "X" Then
    alcoholunknown(3) = True
End If
If rs("sDRUGSyes2") = "X" Then
    drugsyes(3) = True
End If
If rs("sDRUGSNO2") = "X" Then
    drugsno(3) = True
End If
If rs("sDRUGSUNKNOWN2") = "X" Then
    drugsunknown(3) = True
End If
If rs("twomanvehicle1") = "X" Then
    TWOMANVEHICLE(0) = 1
End If
If rs("onemanvehicle1") = "X" Then
    ONEMANVEHICLE(0) = 1
End If
If rs("twomanvehicle2") = "X" Then
    TWOMANVEHICLE(1) = 1
End If
If rs("onemanvehicle2") = "X" Then
    ONEMANVEHICLE(1) = 1
End If
If Not IsNull(rs("vtypeofdrug11")) Then
    For t% = 0 To drugtype(0).ListCount - 1
        If drugtype(0) = Left$(rs("vtypeofdrug11"), 1) Then
            drugtype(0).ListIndex = t%
            t% = drugtype(0).ListCount
        End If
    Next t%
End If
If Not IsNull(rs("stypeofdrug11")) Then
    For t% = 0 To drugtype(3).ListCount - 1
        If drugtype(3) = Left$(rs("stypeofdrug11"), 1) Then
            drugtype(3).ListIndex = t%
            t% = drugtype(3).ListCount
        End If
    Next t%
End If
If Not IsNull(rs("vtypeofdrug21")) Then
    For t% = 0 To drugtype(6).ListCount - 1
        If drugtype(6) = Left$(rs("vtypeofdrug21"), 1) Then
            drugtype(6).ListIndex = t%
            t% = drugtype(6).ListCount
        End If
    Next t%
End If
If Not IsNull(rs("stypeofdrug21")) Then
    For t% = 0 To drugtype(9).ListCount - 1
        If drugtype(9) = Left$(rs("stypeofdrug21"), 1) Then
            drugtype(9).ListIndex = t%
            t% = drugtype(9).ListCount
        End If
    Next t%
End If

If rs("detective1") = "X" Then
    DETECTIVE(0) = 1
End If
If rs("other1") = "X" Then
    TODOTHER(0) = 1
End If
If rs("alone1") = "X" Then
    ALONE(0) = 1
End If
If rs("assisted1") = "X" Then
    ASSISTED(0) = 1
End If
If rs("detective2") = "X" Then
    DETECTIVE(1) = 1
End If
If rs("other2") = "X" Then
    TODOTHER(1) = 1
End If
If rs("alone2") = "X" Then
    ALONE(1) = 1
End If
If rs("assisted2") = "X" Then
    ASSISTED(1) = 1
End If
If rs("runaway1") = 1 Then
    RUNAWAY(0) = 1
End If
If rs("wanted1") = 1 Then
    WANTED(0) = 1
End If
If rs("arrest1") = 1 Then
    ARREST(0) = 1
End If
If rs("warrant1") = 1 Then
    WARRANT(0) = 1
End If
If rs("jail1") = 1 Then
    JAIL(0) = 1
End If
If rs("summons1") = "X" Then
    SUMMONS(0) = 1
End If
If rs("warrant1") = 1 Then
    WARRANT(0) = 1
End If
If rs("runaway2") = 1 Then
    RUNAWAY(1) = 1
End If
If rs("wanted2") = 1 Then
    WANTED(1) = 1
End If
If rs("arrest2") = 1 Then
    ARREST(1) = 1
End If
If rs("warrant2") = 1 Then
    WARRANT(1) = 1
End If
If rs("jail2") = 1 Then
    JAIL(1) = 1
End If
If rs("summons2") = 1 Then
    SUMMONS(1) = 1
End If
If rs("warrant2") = 1 Then
    WARRANT(1) = 1
End If
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
'For t% = 0 To 5
'    If Not IsNull(rs("type" + Mid$(Str$(t% + 1), 2))) Then
'        description(t%) = rs("type" + Mid$(Str$(t% + 1), 2))
'        For tt% = 0 To majorlist(t%).ListCount - 1
'            If majorlist(t%).List(tt%) = rs("major" + Mid$(Str$(t% + 1), 2)) Then
'                majorlist(t%).ListIndex = tt%
'                tt% = majorlist(t%).ListCount - 1
'            End If
'        Next tt%
'        Call setminorlist(t%)
'        For tt% = 0 To minorlist(t%).ListCount - 1
'            If minorlist(t%).List(tt%) = rs("minor" + Mid$(Str$(t% + 1), 2)) Then
'                minorlist(t%).ListIndex = tt%
'                tt% = minorlist(t%).ListCount - 1
'            End If
'        Next tt%
'    End If
'Next t%
ORiGINAL = rs("original")
MODIFIES = rs("modifies")
SUPPLEMENTAL = rs("supplemental")
CASEst = rs("case")
ADDITIONALV = rs("additionalv")
additionalo = rs("additionalo")
additions = rs("additionals")
ADDITIONALR = rs("additionalr")
If rs("complainant1") = 1 Then
    complainant(0) = 1
End If
If rs("complainant2") = 1 Then
    complainant(1) = 1
End If
If Not IsNull(rs("victim1")) Then
    victim(0) = rs("victim1")
End If
If Not IsNull(rs("victim2")) Then
    victim(1) = rs("victim2")
End If
If Not IsNull(rs("subject1")) Then
    subject(0) = rs("subject1")
End If
If Not IsNull(rs("subject2")) Then
    subject(1) = rs("subject2")
End If
typeother(0) = rs("typeother1")
typeother(1) = rs("typeother2")
If Not IsNull(rs("birthdate1")) Then
    BIRTHDATE(0) = rs("birthdate1")
End If
If Not IsNull(rs("birthdate2")) Then
    BIRTHDATE(1) = rs("birthdate2")
End If

'---support
Set rs = db.OpenRecordset("Select * from supplementalsupport where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34) + " AND PAGE = " + PAGE)
If Not rs.EOF Then
    rs.MoveFirst
Else
    GoTo vehgun
End If
If Not IsNull(rs("narrativeonly")) Then
    narrativeonly = rs("narrativeonly")
Else
    narrativeonly = 0
End If
For r% = 0 To 1
  For t% = 1 To vucrlist(r%).ListItems.Count
        vucrlist(r%).ListItems(t%).Selected = fasle
  Next t%
Next r%
For pp% = 1 To 5
    VUCRSEL(pp%, 0) = ""
    VUCRSEL(pp%, 1) = ""
Next pp%
For r% = 0 To 1
    For rr% = 1 To 5
        If Not IsNull(rs("vucr" + Mid$(Str$(r% + 1), 2) + CStr(rr%))) Then
            VUCRSEL(rr%, r%) = rs("vucr" + Mid$(Str$(r% + 1), 2) + CStr(rr%))
        End If
    Next rr%
Next r%
fromfind = 0
Call Command1_Click(0)
vucrf(0).Visible = False
Call Command1_Click(1)
vucrf(1).Visible = False
fromfind = 1
subjectidentifiedyes = rs("subjectidentifiedyes")
subjectidentifiedno = rs("subjectidentifiedno")
subjectlocatedyes = rs("subjectlocatedyes")
subjectlocatedno = rs("subjectlocatedno")
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
reportingofficer(0) = rs("reportingofficer1")
If Not IsNull(rs("reportingdate1")) Then
    REPORTINGOFFICERDATE(0) = rs("reportingdate1")
End If
reportingofficeRunit(0) = rs("reportingunit1")
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
approvingofficer = rs("approvingofficer")
If Not IsNull(rs("approvingdate")) Then
    APPROVINGOFFICERDATE = rs("approvingdate")
End If
approvingofficeRunit = rs("approvingunit")
If Not IsNull(rs("followupdate")) Then
    FOLLOWUPOFFICERDATE = rs("followupdate")
End If
FOLLOWUPOFFICERUNIT = rs("followupunit")
For r% = 0 To 5
    pucrlist(r%).ListIndex = -1
    For rr% = 0 To pucrlist(r%).ListCount - 1
        If Mid$(pucrlist(r%).List(rr%), InStr(pucrlist(r%).List(rr%), "(") + 1, 3) = rs("PUCR" + Mid$(Str$(r% + 1), 2)) Then
            pucrlist(r%).ListIndex = rr%
            rr% = pucrlist(r%).ListCount - 1
        End If
    Next rr%
Next r%
ct% = 0
For t% = 12 To 29 Step 3
    ct% = ct% + 1
    If ct% > 3 Then
        ct% = 1
    End If
    st% = (t% Mod 3) + ct%
    For tt% = 1 To 3
        If Not IsNull(rs("PTYPEOFDRUG" + Mid$(Str$(ct%), 2) + Mid$(Str$(tt%), 2))) Then
            For ttt% = 0 To drugtype(t%).ListCount - 1
                If rs("PTYPEOFDRUG" + Mid$(Str$(ct%), 2) + Mid$(Str$(tt%), 2)) = Left$(drugtype(t%).List(ttt%), 1) Then
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
    Next tt%
Next t%
For t% = 0 To 5
    If Not IsNull(rs("group" + Mid$(Str$(t% + 1), 2))) Then
        For tt% = 0 To group(t%).ListCount - 1
            If Mid$(group(t%).List(tt%), InStr(group(t%).List(tt%), "(") + 1, 2) = rs("group" + Mid$(Str$(t% + 1), 2)) Then
                group(t%).ListIndex = tt%
                tt% = group(t%).ListCount - 1
            End If
        Next tt%
    Else
        group(t%).ListIndex = -1
    End If
    group(t%).Visible = False
    If Not IsNull(rs("numvehicles" + Mid$(Str$(t% + 1), 2))) Then
        numvehicle(t%) = rs("numvehicles" + Mid$(Str$(t% + 1), 2))
    End If
        If Not IsNull(rs("daterecovered" + Mid$(Str$(t% + 1), 2))) Then
            DATERECOVERED(t%) = rs("daterecovered" + Mid$(Str$(t% + 1), 2))
        End If
Next t%
For t% = 6 To 11
    If Not IsNull(rs("numvehicler" + Mid$(Str$(t% - 5), 2))) Then
        numvehicle(t%) = rs("numvehicler" + Mid$(Str$(t% - 5), 2))
    End If
Next t%

For t% = 1 To 2
    For tt% = 3 To 9
        For ttt% = 0 To relationship(((t% - 1) * 10) + tt%).ListCount - 1
            relationship(((t% - 1) * 10) + tt%).Selected(ttt%) = False
        Next ttt%
        If Not IsNull(rs("relationship" + Mid$(Str$(t%), 2) + Mid$(Str$(tt% + 1), 2))) Then
            For ttt% = 0 To relationship(((t% - 1) * 10) + tt%).ListCount - 1
                If rs("relationship" + Mid$(Str$(t%), 2) + Mid$(Str$(tt% + 1), 2)) = Mid$(relationship(((t% - 1) * 10) + tt%).List(ttt%), InStr(relationship(((t% - 1) * 10) + tt%).List(ttt%), "(") + 1, 2) Then
                    relationship(((t% - 1) * 10) + tt%).ListIndex = ttt%
                    relationship(((t% - 1) * 10) + tt%).Selected(ttt%) = True
                    ttt% = relationship(tt%).ListCount - 1
                End If
            Next ttt%
        Else
            relationship(((t% - 1) * 10) + tt%).ListIndex = -1
        End If
    Next tt%
Next t%
If Not IsNull(rs("vtypeofdrug12")) Then
    For t% = 0 To drugtype(1).ListCount - 1
        If drugtype(1) = Left$(rs("vtypeofdrug12"), 1) Then
            drugtype(1).ListIndex = t%
            t% = drugtype(1).ListCount
        End If
    Next t%
End If
If Not IsNull(rs("vtypeofdrug13")) Then
    For t% = 0 To drugtype(2).ListCount - 1
        If drugtype(2) = Left$(rs("vtypeofdrug13"), 1) Then
            drugtype(2).ListIndex = t%
            t% = drugtype(2).ListCount
        End If
    Next t%
End If
If Not IsNull(rs("stypeofdrug12")) Then
    For t% = 0 To drugtype(4).ListCount - 1
        If drugtype(4) = Left$(rs("stypeofdrug12"), 1) Then
            drugtype(4).ListIndex = t%
            t% = drugtype(4).ListCount
        End If
    Next t%
End If
If Not IsNull(rs("stypeofdrug13")) Then
    For t% = 0 To drugtype(5).ListCount - 1
        If drugtype(5) = Left$(rs("stypeofdrug13"), 1) Then
            drugtype(5).ListIndex = t%
            t% = drugtype(5).ListCount
        End If
    Next t%
End If
If Not IsNull(rs("vtypeofdrug22")) Then
    For t% = 0 To drugtype(7).ListCount - 1
        If drugtype(7) = Left$(rs("vtypeofdrug22"), 1) Then
            drugtype(7).ListIndex = t%
            t% = drugtype(7).ListCount
        End If
    Next t%
End If
If Not IsNull(rs("vtypeofdrug23")) Then
    For t% = 0 To drugtype(8).ListCount - 1
        If drugtype(8) = Left$(rs("vtypeofdrug23"), 1) Then
            drugtype(8).ListIndex = t%
            t% = drugtype(8).ListCount
        End If
    Next t%
End If
If Not IsNull(rs("stypeofdrug22")) Then
    For t% = 0 To drugtype(10).ListCount - 1
        If drugtype(10) = Left$(rs("stypeofdrug22"), 1) Then
            drugtype(10).ListIndex = t%
            t% = drugtype(10).ListCount
        End If
    Next t%
End If
If Not IsNull(rs("stypeofdrug23")) Then
    For t% = 0 To drugtype(11).ListCount - 1
        If drugtype(11) = Left$(rs("stypeofdrug23"), 1) Then
            drugtype(11).ListIndex = t%
            t% = drugtype(11).ListCount
        End If
    Next t%
End If

vehgun:
'---vehgun
Set rs = db.OpenRecordset("Select * from vehgun where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34) + " AND PAGE = " + PAGE)
If Not rs.EOF Then
    rs.MoveFirst
Else
    GoTo MS
End If
SSTOLEN = rs("stolen")
SRECOVERED = rs("recovered")
SFOUND = rs("found")
STOWED = rs("towed")
SSUSPECT = rs("suspect")
SVICTIM = rs("victim")
SVEHICLE = rs("vehicle")
SGUN = rs("gun")
SBOAT = rs("boat")
sLICENSEPLATE = rs("licenseplate")
SSECURITIES = rs("securities")
SARTICLE = rs("article")
VIN = rs("vin")
HULL = rs("hull")
SERIAL = rs("serial")
SERIALSTATE = rs("serialstate")
YEARREG = rs("yearreg")
YEAREXP = rs("yearexp")
YEARN = rs("year")
make = rs("make")
stype = rs("type")
model = rs("model")
STYLE = rs("style")
scolor = rs("color")
BRANDNAME = rs("brandname")
CALIBER = rs("caliber")
NIC = rs("nic")
DENOMINATION = rs("denomination")
ISSUER = rs("issuer")
If Not IsNull(rs("securitiesdate")) Then
    SECURITIESDATE = rs("securitiesdate")
Else
    SECURITIESDATE = ""
End If
MISCELLANEOUS = rs("miscellaneous")
MS:
For yy% = 0 To 1
    If Val(subject(yy%)) > 0 Then
        Set db = OpenDatabase(nwl + "lawsuite.mdb")
        ssql = ""
        If IsDate(BIRTHDATE(0)) Then
            ssql = ssql + " and birthdate = #" + BIRTHDATE(0) + "#"
        End If
        Set rs = db.OpenRecordset("select mugshot from people where dpnamelf = " + Chr$(34) + vsname(yy%) + Chr$(34) + ssql + " and not mugshot is null")
        If Not rs.EOF Then
            rs.MoveFirst
            mugshot(yy%).Picture = LoadPicture(rs("mugshot"))
        End If
    End If
Next yy%
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("Select max(page) as ctpg from supplemental where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
If rs.EOF Then
    pgof = "Page 2 (New)"
Else
    If Not IsNull(rs("ctpg")) Then
        pgof = "Page " + CStr(Val(PAGE) + 1) + " of " + CStr(1 + rs("ctpg"))
    Else
        pgof = "Page 2 (New)"
    End If
End If

fromfind = 0
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub

Private Sub saveincident()
Dim db As Database, rs, rs2 As Recordset, lu As String
On Error GoTo oderror1
od1:
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("Select * from SUPPLEMENTAL where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34) + " and page = " + PAGE)
ecc = 0
If Not rs.EOF Then
    rs.MoveFirst
    If rs("schanged") = 1 Then
        schanged = 1
    Else
        schanged = 0
    End If
    rs.Edit
Else
    rs.AddNew
    schanged = 1
End If
schanged = 1
rs("PAGE") = Val(PAGE)
rs("incidentnumber") = incidentnumber
rs("computerequipment1") = computerequipment(0)
rs("computerequipment2") = computerequipment(1)
'ian's code
rs("narrative") = NARRATIVE.Text
'end ian's code
rs("schanged") = schanged
rs("locationnumber1") = LOCATIONNUMBER(0)
rs("locationnumber2") = LOCATIONNUMBER(1)
rs("name1") = vsname(0)
rs("address1") = address(0)
rs("city1") = city(0)
rs("state1") = state(0)
rs("zipcode1") = zipcode(0)
rs("resident1") = Left$(resident(0).List(resident(0).ListIndex), 1)
rs("race1") = Left$(race(0).List(race(0).ListIndex), 1)
rs("sex1") = Left$(sex(0).List(sex(0).ListIndex), 1)
rs("age1") = age(0)
rs("ethnicity1") = Left$(ethnicity(0).List(ethnicity(0).ListIndex), 1)
For t% = 0 To 2
    If relationship(t%).ListIndex > -1 Then
        rs("relationship1" + Mid$(Str$(t% + 1), 2)) = Mid$(relationship(t%).List(relationship(t%).ListIndex), InStr(relationship(t%).List(relationship(t%).ListIndex), "(") + 1, 2)
    Else
        rs("relationship1" + Mid$(Str$(t% + 1), 2)) = Null
    End If
Next t%
rs("homephoneDAY1") = HOMEDAYPHONE(0)
rs("WORKphoneDAY1") = WORKDAYPHONE(0)
rs("HOMEphoneNIGHT1") = HOMENIGHTPHONE(0)
rs("WORKphoneNIGHT1") = WORKNIGHTPHONE(0)
rs("HEIGHT1") = ht(0)
rs("WEIGHT1") = weight(0)
rs("HAIR1") = hair(0)
rs("EYES1") = eyes(0)
rs("PECULIARITIES1") = peculiarities(0)
For s% = 1 To 2
    For t% = 1 To 5
        rs("typeofinjury" + Mid$(Str$(s%), 2) + Mid$(Str$(t%), 2)) = ""
    Next t%
Next s%
For s% = 0 To 1
    IIDX% = 0
    For t% = 1 To injury(s%).ListItems.Count
        If injury(s%).ListItems(t%).Selected = True Then
            IIDX% = IIDX% + 1
            If IIDX% < 6 Then
                rs("typeofinjury" + Mid$(Str$(s% + 1), 2) + Mid$(Str$(IIDX%), 2)) = Mid$(injury(s%).ListItems(t%), InStr(injury(s%).ListItems(t%), "(") + 1, 1)
            Else
                t% = injury(s%).ListItems.Count
            End If
        End If
    Next t%
Next s%
If VISIBLEINJURYYES(0) Then
    rs("visibleinjuryyes1") = "X"
Else
    rs("visibleinjuryno1") = "X"
End If
If NONVISIBLEINJURYYES(0) Then
    rs("nonvisibleinjuryyes1") = "X"
Else
    rs("nonvisibleinjuryno1") = "X"
End If
rs("name2") = vsname(1)
rs("address2") = address(1)
rs("city2") = city(1)
rs("state2") = state(1)
rs("zipcode2") = zipcode(1)
rs("resident2") = Left$(resident(1).List(resident(1).ListIndex), 1)
rs("race2") = Left$(race(1).List(race(1).ListIndex), 1)
rs("sex2") = Left$(sex(1).List(sex(1).ListIndex), 1)
rs("age2") = age(1)
rs("ethnicity2") = Left$(ethnicity(1).List(ethnicity(1).ListIndex), 1)
For t% = 10 To 12
    If relationship(t%).ListIndex > -1 Then
        rs("relationship2" + Mid$(Str$(t% - 9), 2)) = Mid$(relationship(t%).List(relationship(t%).ListIndex), InStr(relationship(t%).List(relationship(t%).ListIndex), "(") + 1, 2)
    Else
        rs("relationship2" + Mid$(Str$(t% - 9), 2)) = Null
    End If
Next t%
rs("homephoneDAY2") = HOMEDAYPHONE(1)
rs("WORKphoneDAY2") = WORKDAYPHONE(1)
rs("HOMEphoneNIGHT2") = HOMENIGHTPHONE(1)
rs("WORKphoneNIGHT2") = WORKNIGHTPHONE(1)
rs("HEIGHT2") = ht(1)
rs("WEIGHT2") = weight(1)
rs("HAIR2") = hair(1)
rs("EYES2") = eyes(1)
rs("PECULIARITIES2") = peculiarities(1)
If VISIBLEINJURYYES(1) Then
    rs("visibleinjuryyes2") = "X"
Else
    rs("visibleinjuryno2") = "X"
End If
If NONVISIBLEINJURYYES(1) Then
    rs("nonvisibleinjuryyes2") = "X"
Else
    rs("nonvisibleinjuryno2") = "X"
End If

If alcoholyes(0) Then
    rs("valcoholyes1") = "X"
Else
If alcoholno(0) Then
    rs("valcoholno1") = "X"
Else
If alcoholunknown(0) Then
    rs("valcoholunknown1") = "X"
End If
End If
End If
If drugsyes(0) Then
    rs("vdrugsyes1") = "X"
Else
If drugsno(0) Then
    rs("vdrugsno1") = "X"
Else
If drugsunknown(0) Then
    rs("vdrugsunknown1") = "X"
End If
End If
End If
If alcoholyes(1) Then
    rs("salcoholyes1") = "X"
Else
If alcoholno(1) Then
    rs("salcoholno1") = "X"
Else
If alcoholunknown(1) Then
    rs("salcoholunknown1") = "X"
End If
End If
End If
If drugsyes(1) Then
    rs("sdrugsyes1") = "X"
Else
If drugsno(1) Then
    rs("sdrugsno1") = "X"
Else
If drugsunknown(1) Then
    rs("sdrugsunknown1") = "X"
End If
End If
End If
If alcoholyes(2) Then
    rs("valcoholyes2") = "X"
Else
If alcoholno(2) Then
    rs("valcoholno2") = "X"
Else
If alcoholunknown(2) Then
    rs("valcoholunknown2") = "X"
End If
End If
End If
If drugsyes(2) Then
    rs("vdrugsyes2") = "X"
Else
If drugsno(2) Then
    rs("vdrugsno2") = "X"
Else
If drugsunknown(2) Then
    rs("vdrugsunknown2") = "X"
End If
End If
End If
If alcoholyes(3) Then
    rs("salcoholyes2") = "X"
Else
If alcoholno(3) Then
    rs("salcoholno2") = "X"
Else
If alcoholunknown(3) Then
    rs("salcoholunknown2") = "X"
End If
End If
End If
If drugsyes(3) Then
    rs("sdrugsyes2") = "X"
Else
If drugsno(3) Then
    rs("sdrugsno2") = "X"
Else
If drugsunknown(3) Then
    rs("sdrugsunknown2") = "X"
End If
End If
End If

If TWOMANVEHICLE(0) = 1 Then
    rs("twomanvehicle1") = "X"
End If
If ONEMANVEHICLE(0) = 1 Then
    rs("onemanvehicle1") = "X"
End If
If DETECTIVE(0) = 1 Then
    rs("detective1") = "X"
End If
If TODOTHER(0) = 1 Then
    rs("other1") = "X"
End If
If ALONE(0) = 1 Then
    rs("alone1") = "X"
End If
If ASSISTED(0) = 1 Then
    rs("assisted1") = "X"
End If

If TWOMANVEHICLE(1) = 1 Then
    rs("twomanvehicle2") = "X"
End If
If ONEMANVEHICLE(1) = 1 Then
    rs("onemanvehicle2") = "X"
End If
If DETECTIVE(1) = 1 Then
    rs("detective2") = "X"
End If
If TODOTHER(1) = 1 Then
    rs("other2") = "X"
End If
If ALONE(1) = 1 Then
    rs("alone2") = "X"
End If
If ASSISTED(1) = 1 Then
    rs("assisted2") = "X"
End If

If drugtype(0).ListIndex > -1 Then
    rs("vtypeofdrug11") = Left$(drugtype(0).List(drugtype(0).ListIndex), 1)
Else
    rs("vtypeofdrug11") = Null
End If
If drugtype(3).ListIndex > -1 Then
    rs("stypeofdrug11") = Left$(drugtype(3).List(drugtype(3).ListIndex), 1)
Else
    rs("stypeofdrug11") = Null
End If
If drugtype(6).ListIndex > -1 Then
    rs("vtypeofdrug21") = Left$(drugtype(6).List(drugtype(6).ListIndex), 1)
Else
    rs("vtypeofdrug21") = Null
End If
If drugtype(9).ListIndex > -1 Then
    rs("stypeofdrug21") = Left$(drugtype(9).List(drugtype(9).ListIndex), 1)
Else
    rs("stypeofdrug21") = Null
End If

For yy% = 0 To 1
    rs("runaway" + Mid$(Str$(yy% + 1), 2)) = RUNAWAY(yy%)
    rs("wanted" + Mid$(Str$(yy% + 1), 2)) = WANTED(yy%)
    rs("arrest" + Mid$(Str$(yy% + 1), 2)) = ARREST(yy%)
    rs("warrant" + Mid$(Str$(yy% + 1), 2)) = WARRANT(yy%)
    rs("jail" + Mid$(Str$(yy% + 1), 2)) = JAIL(yy%)
    rs("summons" + Mid$(Str$(yy% + 1), 2)) = SUMMONS(yy%)
Next yy%

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
rs("original") = ORiGINAL.Value

rs("modifies") = MODIFIES.Value
rs("supplemental") = SUPPLEMENTAL.Value
rs("case") = CASEst.Value
rs("additionalv") = ADDITIONALV.Value
rs("additionalo") = additionalo.Value
rs("additionals") = ADDITIONALS.Value
rs("additionalr") = ADDITIONALR.Value

If complainant(0) = 1 Then
    rs("complainant1") = "X"
End If
If complainant(1) = 1 Then
    rs("complainant2") = "X"
End If
If victim(0) > "" Then
    rs("victim1") = Val(victim(0))
    Select Case UCase(TV1)
        Case "I"
            rs("individual1") = True
        Case "B"
            rs("business1") = True
        Case "F"
            rs("financialinstitution1") = True
        Case "G"
            rs("government1") = True
        Case "R"
            rs("religiousorganization1") = True
        Case "S"
            rs("societypublic1") = True
        Case "O"
            rs("tvother1") = True
        Case "U"
            rs("unknown1") = True
        Case "P"
            rs("policeofficer1") = True
    End Select
Else
    rs("victim1") = Null
End If
If victim(1) > "" Then
    rs("victim2") = Val(victim(1))
    Select Case UCase(TV2)
        Case "I"
            rs("individual2") = True
        Case "B"
            rs("business2") = True
        Case "F"
            rs("financialinstitution2") = True
        Case "G"
            rs("government2") = True
        Case "R"
            rs("religiousorganization2") = True
        Case "S"
            rs("societypublic2") = True
        Case "O"
            rs("tvother2") = True
        Case "U"
            rs("unknown2") = True
        Case "P"
            rs("policeofficer2") = True
    End Select
Else
    rs("victim2") = Null
End If
If subject(0) > "" Then
    rs("subject1") = Val(subject(0))
Else
    rs("subject1") = Null
End If
If subject(1) > "" Then
    rs("subject2") = Val(subject(1))
Else
    rs("subject2") = Null
End If

rs("typeother1") = typeother(0)
rs("typeother2") = typeother(1)
If IsDate(BIRTHDATE(0)) Then
    rs("birthdate1") = BIRTHDATE(0)
End If
If IsDate(BIRTHDATE(1)) Then
    rs("birthdate2") = BIRTHDATE(1)
End If
rs.Update

'---support
Set rs = db.OpenRecordset("Select * from supplementalsupport where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34) + " and page =" + PAGE)
ecc = 0
If Not rs.EOF Then
    rs.MoveFirst
    rs.Edit
Else
    rs.AddNew
End If
schanged = 1
If tempsave = 1 Then
    rs("TEMP") = "Y"
Else
    rs("TEMP") = "N"
End If
rs("narrativeonly") = narrativeonly
rs("INCIDENTNUMBER") = incidentnumber
rs("PAGE") = Val(PAGE)
For t% = 0 To 5
    If Val(numvehicle(t%)) > 0 Then
        rs("numvehicles" + Mid$(Str$(t% + 1), 2)) = Val(numvehicle(t%))
    Else
        rs("numvehicles" + Mid$(Str$(t% + 1), 2)) = Null
    End If
Next t%
ct% = 0
For t% = 6 To 11
    If Val(numvehicle(t%)) > 0 Then
        rs("numvehicler" + Mid$(Str$(ct% + 1), 2)) = Val(numvehicle(t%))
    Else
        rs("numvehicler" + Mid$(Str$(ct% + 1), 2)) = Null
    End If
    ct% = ct% + 1
Next t%
For t% = 0 To 5
    If IsDate(DATERECOVERED(t%)) Then
        rs("daterecovered" + Mid$(Str$(t% + 1), 2)) = DATERECOVERED(t%)
    Else
        rs("daterecovered" + Mid$(Str$(t% + 1), 2)) = Null
    End If
Next t%
rs("subjectidentifiedyes") = subjectidentifiedyes
rs("subjectidentifiedno") = subjectidentifiedno
rs("subjectlocatedyes") = subjectlocatedyes
rs("subjectlocatedno") = subjectlocatedno
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
For t% = 6 To 11
    If numvehicle(t%) > "" Then
        rs("numvehicler" + Mid$(Str$(t% - 5), 2)) = Val(numvehicle(t%))
    Else
        rs("numvehicler" + Mid$(Str$(t% - 5), 2)) = Null
    End If
Next t%
ct% = 0
For t% = 12 To 23
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

If drugtype(1).ListIndex > -1 Then
    rs("vtypeofdrug12") = Left$(drugtype(1).List(drugtype(1).ListIndex), 1)
Else
    rs("vtypeofdrug12") = Null
End If
If drugtype(2).ListIndex > -1 Then
    rs("vtypeofdrug13") = Left$(drugtype(2).List(drugtype(2).ListIndex), 1)
Else
    rs("vtypeofdrug13") = Null
End If
If drugtype(4).ListIndex > -1 Then
    rs("stypeofdrug12") = Left$(drugtype(4).List(drugtype(4).ListIndex), 1)
Else
    rs("stypeofdrug12") = Null
End If
If drugtype(5).ListIndex > -1 Then
    rs("stypeofdrug13") = Left$(drugtype(5).List(drugtype(5).ListIndex), 1)
Else
    rs("stypeofdrug13") = Null
End If
If drugtype(7).ListIndex > -1 Then
    rs("vtypeofdrug22") = Left$(drugtype(7).List(drugtype(7).ListIndex), 1)
Else
    rs("vtypeofdrug22") = Null
End If
If drugtype(8).ListIndex > -1 Then
    rs("vtypeofdrug23") = Left$(drugtype(8).List(drugtype(8).ListIndex), 1)
Else
    rs("vtypeofdrug23") = Null
End If
If drugtype(10).ListIndex > -1 Then
    rs("stypeofdrug22") = Left$(drugtype(10).List(drugtype(10).ListIndex), 1)
Else
    rs("stypeofdrug22") = Null
End If
If drugtype(11).ListIndex > -1 Then
    rs("stypeofdrug23") = Left$(drugtype(11).List(drugtype(11).ListIndex), 1)
Else
    rs("stypeofdrug23") = Null
End If
For t% = 1 To 2
    TCT% = 0
    If Not (vucrlist(t% - 1).SelectedItem Is Nothing) Then
        For tt% = 1 To vucrlist(t% - 1).ListItems.Count
            If vucrlist(t% - 1).ListItems(tt%).Selected = True Then
                TCT% = TCT% + 1
                rs("vucr" + Mid$(Str$(t%), 2) + Mid$(Str$(TCT%), 2)) = Mid$(vucrlist(t% - 1).ListItems(tt%), InStr(vucrlist(t% - 1).ListItems(tt%), "(") + 1, 3)
            End If
        Next tt%
    End If
    For tt% = TCT% + 1 To 5
        rs("vucr" + Mid$(Str$(t%), 2) + Mid$(Str$(tt%), 2)) = Null
    Next tt%
Next t%
For t% = 1 To 6
    If pucrlist(t% - 1).ListIndex > -1 Then
        rs("Pucr" + Mid$(Str$(t%), 2)) = Mid$(pucrlist(t% - 1).List(pucrlist(t% - 1).ListIndex), InStr(pucrlist(t% - 1).List(pucrlist(t% - 1).ListIndex), "(") + 1, 3)
    Else
        rs("Pucr" + Mid$(Str$(t%), 2)) = Null
    End If
Next t%
For t% = 1 To 6
    If group(t% - 1).ListIndex > -1 Then
        rs("group" + Mid$(Str$(t%), 2)) = Mid$(group(t% - 1).List(group(t% - 1).ListIndex), InStr(group(t% - 1).List(group(t% - 1).ListIndex), "(") + 1, 2)
    Else
        rs("group" + Mid$(Str$(t%), 2)) = Null
    End If
Next t%
For t% = 1 To 2
    For tt% = 3 To 9
        If relationship(((t% - 1) * 10) + tt%).ListIndex > -1 Then
            rs("relationship" + Mid$(Str$(t%), 2) + Mid$(Str$(tt% + 1), 2)) = Mid$(relationship(((t% - 1) * 10) + tt%).List(relationship(((t% - 1) * 10) + tt%).ListIndex), InStr(relationship(((t% - 1) * 10) + tt%).List(relationship(((t% - 1) * 10) + tt%).ListIndex), "(") + 1, 2)
        Else
            rs("relationship" + Mid$(Str$(t%), 2) + Mid$(Str$(tt% + 1), 2)) = Null
        End If
    Next tt%
Next t%
'CES Code
rs("userfullname") = frmLogin.userfullname
rs("userid") = frmLogin.userid
rs("ORINUMBER") = frmLogin.orinumber
rs("udate") = Format$(Now, "mm/dd/yyyy")
rs("utime") = Format$(Now, "hh:mm:ss")
'********
rs.Update

'---vehgun
If SSTOLEN = 1 Or SRECOVERED = 1 Or SFOUND = 1 Or STOWED = 1 Or SSUSPECT = 1 Or SVICTIM = 1 Or SVEHICLE = 1 Or SGUN = 1 Or SBOAT = 1 Or sLICENSEPLATE = 1 Or SSECURITIES = 1 Or SARTICLE = 1 Or VIN > "" Or HULL > "" Or SERIAL > "" Or SERIALSTATE > "" Or YEARREG > "" Or YEAREXP > "" Or YEARN > "" Or make > "" Or stype > "" Or model > "" Or STYLE > "" Or scolor > "" Or BRANDNAME > "" Or CALIBER > "" Or NIC > "" Or DENOMINATION > "" Or ISSUER > "" Or IsDate(SECURITIESDATE) Or MISCELLANEOUS > "" Then
    Set rs = db.OpenRecordset("Select * from vehgun where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34) + " and page = " + PAGE)
    ecc = 0
    If Not rs.EOF Then
        rs.MoveFirst
        rs.Delete
    End If
    rs.AddNew
    rs("PAGE") = Val(PAGE)
    rs("incidentnumber") = incidentnumber
    rs("stolen") = SSTOLEN
    rs("recovered") = SRECOVERED
    rs("found") = SFOUND
    rs("towed") = STOWED
    rs("suspect") = SSUSPECT
    rs("victim") = SVICTIM
    rs("vehicle") = SVEHICLE
    rs("gun") = SGUN
    rs("boat") = SBOAT
    rs("licenseplate") = sLICENSEPLATE
    rs("securities") = SSECURITIES
    rs("article") = SARTICLE
    rs("vin") = VIN
    rs("hull") = HULL
    rs("serial") = SERIAL
    rs("serialstate") = SERIALSTATE
    rs("yearreg") = YEARREG
    rs("yearexp") = YEAREXP
    rs("year") = YEARN
    rs("make") = make
    rs("type") = stype
    rs("model") = model
    rs("style") = STYLE
    rs("color") = scolor
    rs("brandname") = BRANDNAME
    rs("caliber") = CALIBER
    rs("nic") = NIC
    rs("denomination") = DENOMINATION
    rs("issuer") = ISSUER
    If Not IsDate(SECURITIESDATE) Then
        rs("SECURITIESDATE") = Null
    Else
        rs("securitiesdate") = SECURITIESDATE
    End If
    rs("miscellaneous") = MISCELLANEOUS
    rs("userfullname") = frmLogin.userfullname
    rs("userid") = frmLogin.userid
    rs("ORINUMBER") = frmLogin.orinumber
    rs("udate") = Format$(Now, "mm/dd/yyyy")
    rs("utime") = Format$(Now, "hh:mm:ss")
    rs.Update
End If

For t% = 0 To 1
    Set rs = db.OpenRecordset("SELECT * FROM CODES WHERE CODE = '" + city(t%) + "' AND TYPE = 'city'")
    If rs.EOF Then
        rs.AddNew
        rs("CODE") = city(t%)
        For tt% = 0 To 1
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
        For tt% = 0 To 1
            state(tt%).AddItem state(t%)
        Next tt%
        rs("TYPE") = "state"
        rs("DEFAULT") = "N"
        rs.Update
    End If
Next t%
On Error GoTo oderror2
od2:
Set db = OpenDatabase(nwl + "LAWSUITE.MDB")

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

'--- PEOPLE
For t% = 0 To 1
    If vsname(t%) = "UNKNOWN" Or vsname(t%) = "" Then
        GoTo nextt
    End If
    Set rs = db.OpenRecordset("select * from people where dpnamelf =" + Chr$(34) + vsname(t%) + Chr$(34))
    If rs.EOF Then
        rs.AddNew
        For tt% = 0 To 1
            vsname(tt%).AddItem vsname(t%)
        Next tt%
    Else
        rs.MoveFirst
        rs.Edit
    End If
    rs("dpnamelf") = vsname(t%)
    rs("dphaddress") = address(t%)
    rs("dphaddress2") = city(t%) + ", " + state(t%) + " " + zipcode(t%)
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
        rs("HEIGHT") = ht(t%)
        rs("WEIGHT") = weight(t%)
        rs("HAIR") = hair(t%)
        rs("EYES") = eyes(t%)
        rs("PECULIARITIES") = peculiarities(t%)
        rs("resident") = Left$(resident(t%).List(resident(t%).ListIndex), 1)
        rs("race") = Left$(race(t%).List(race(t%).ListIndex), 1)
        rs("sex") = Left$(sex(t%).List(sex(t%).ListIndex), 1)
        rs("age") = age(t%)
        rs("ethnicity") = Left$(ethnicity(t%).List(ethnicity(t%).ListIndex), 1)
        If IsDate(BIRTHDATE(t%)) Then
            rs("birthdate") = BIRTHDATE(t%)
        End If
    End If
    hoLdname = vsname(t%)
    osort1$ = ""
    If Left$(hoLdname, 1) = " " Then
        hoLdname = Mid$(hoLdname, 2)
        osort1$ = hoLdname
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
        If Mid$(tso$, tt%, 1) = " " Then
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
    tso$ = Left$(tso$, firstspace% - 1)
    If Right$(tso$, 1) = "," Then
        tso$ = Left$(tso$, Len(tso$) - 1)
    End If
    tempsort$ = tempsort$ + " " + tso$
    If osort1$ = "" Then
        osort1$ = tempsort$
    End If
rsupdate:
    rs("dpname") = osort1$
    rs.Update
nextt:
Next t%


db.Close
On Error Resume Next
Exit Sub
oderror1:
If Err > 3200 Then
    Resume od1
Else
    Resume Next
    Resume
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
Private Sub drugmeasurement_Click(Index As Integer)

Select Case Index
    Case 3, 4, 5
        
    Case 9 To 29
        
End Select
    
End Sub
Private Sub drugamt_Change(Index As Integer)

Select Case Index
    Case 3, 4, 5
        
    Case 9 To 29
        
End Select

End Sub
Private Sub XGROUP_Click(Index As Integer)
'group(Index).Top = description(Index).Top - 1000
'group(Index).Left = description(Index).Left
group(Index).Visible = True
'---- setfocus logic ----
'         group(index).SetFocus
          If group(Index).Visible Then
              group(Index).SetFocus
           End If

End Sub

Private Sub ORGINAL_Click()

End Sub
Private Sub OFFENDERO_Click()

End Sub
Private Sub FILLDATA(IDX As Integer)
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
Set rs = db.OpenRecordset("SELECT * FROM PEOPLE WHERE DPnameLF = " + Chr$(34) + vsname(IDX) + Chr$(34))
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
        If Not IsNull(rs("resident")) Then
            For tt% = 0 To resident(t%).ListCount - 1
                If Left$(resident(t%).List(tt%), 1) = rs("resident") Then
                    resident(t%).ListIndex = tt%
                    tt% = resident(t%).ListCount - 1
                End If
            Next tt%
        Else
            resident(t%).ListIndex = -1
        End If
    End If
    If Not IsNull(rs("RACE")) Then
        For tt% = 0 To race(t%).ListCount - 1
            If Left$(race(t%).List(tt%), 1) = rs("RACE") Then
                race(t%).ListIndex = tt%
                tt% = race(t%).ListCount - 1
            End If
        Next tt%
    Else
        race(t%).ListIndex = -1
    End If
    If Not IsNull(rs("SEX")) Then
        For tt% = 0 To sex(t%).ListCount - 1
            If Left$(sex(t%).List(tt%), 1) = rs("SEX") Then
                sex(t%).ListIndex = tt%
                tt% = sex(t%).ListCount - 1
            End If
        Next tt%
    Else
        sex(t%).ListIndex = -1
    End If
    If Not IsNull(rs("ETHNICITY")) Then
        For tt% = 0 To ethnicity(t%).ListCount - 1
            If Left$(ethnicity(t%).List(tt%), 1) = rs("ETHNICITY") Then
                ethnicity(t%).ListIndex = tt%
                tt% = ethnicity(t%).ListCount - 1
            End If
        Next tt%
    Else
        ethnicity(t%).ListIndex = -1
    End If
    If Not IsNull(rs("age")) Then
        age(t%) = rs("age")
    Else
        age(t%) = ""
    End If
    If IsDate(rs("birthdate")) Then
        BIRTHDATE(t%) = CStr(rs("Birthdate"))
    Else
        BIRTHDATE(t%) = ""
    End If
    If Not IsNull(rs("mugshot")) Then
        mugshot(t%).Picture = LoadPicture(rs("mugshot"))
    Else
        mugshot(t%).Picture = LoadPicture()
    End If
End If
db.Close
On Error Resume Next
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

Private Sub reportingofficerdate_GotFocus(Index As Integer)
If reportingofficer(Index).ListIndex > -1 And REPORTINGOFFICERDATE(Index) = "" Then
    REPORTINGOFFICERDATE(Index) = Format$(Date$, "mm/dd/yyyy")
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



Private Sub spellcheck(DONE As Boolean)
Dim wd As New Word.Application
Dim wdsp As Word.SpellingSuggestions
On Error GoTo cmdCheckErr
GETOUT% = 0
lstframe.Visible = False
While tempword > "" And GETOUT% = 0
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
Private Sub ShowApplicableContainers(objControl As Object)

    Dim intNumContainerLevelsInPicture1 As Integer
    Dim objHoldOriginal As Object
        
    Set objHoldOriginal = objControl
    
    On Error GoTo errh
    
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
            Case "VUCRF"
                Call Command1_Click(objControl.Index)
            Case "PUCRLIST"
                Call Command2_Click(objControl.Index)
            Case "SDRUGFRAME"
                Call Command7_Click(objControl.Index)
            Case "NUMVEHICLE"
                'objControl.Visible = False
                numvehicle(objControl.Index).Visible = True
            Case "DRUGAMT"
                Call drugtype_Click(objControl.Index)
            Case "RELATIONSHIPFRAME"
                Call setrel_Click(objControl.Index)
            Case "DATERECOVERED"
                Call totalvalue_LostFocus(objControl.Index)
            Case "XGROUP"
                Call group_Click(objControl.Index)
            
        End Select
        
        Set objControl = objControl.Container
    Next x%
GETOUT:
Exit Sub
errh:
    Set objControl = Nothing
    Set objHoldOriginal = Nothing
    Resume GETOUT

End Sub
'*************

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
'===== Data Element 18, 19
Dim tempgroup As Integer, itmx As String
tempgroup = Val(Mid$(group(Index).List(group(Index).ListIndex), InStr(group(Index).List(group(Index).ListIndex), "(") + 1, 2))
On Error Resume Next
Call XGROUP_Click(Index)
On Error GoTo 0
End Sub



Private Sub pucrlist_LostFocus(Index As Integer)
If fromfind = 1 Then
    Exit Sub
End If
If BACKTAB = 0 Then
'---- setfocus logic ----
'             description(index).SetFocus
          If description(Index).Visible Then
              description(Index).SetFocus
           End If
Else
    BACKTAB = 0
End If
End Sub
Private Sub group_Click(Index As Integer)
If FROMKEY = 1 Then
    FROMKEY = 0
    Exit Sub
End If


On Error Resume Next
If fromfind = 1 Then
    Exit Sub
End If
If Mid$(group(Index).List(group(Index).ListIndex), InStr(group(Index).List(group(Index).ListIndex), "(") + 1, 2) = "10" Then
    sdrugframe(Index + 4).Left = 2000
    sdrugframe(Index + 4).Top = description(Index).Top - sdrugframe(Index + 2).Height - 100
    sdrugframe(Index + 4).Visible = True
    'RLB Code
    sdrugframe(Index + 4).ZOrder
    '********
'---- setfocus logic ----
'             drugtype(index + 4).SetFocus
          If drugtype(Index + 4).Visible Then
              drugtype(Index + 4).SetFocus
           End If
Else
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


'===== Mandatories E - 2
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
'==== Mandatories E - 8A

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

Private Sub YEAREXP_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub YEARN_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub

Private Sub YEARREG_GotFocus()
If Me.ActiveControl.Top > (-1 * Picture2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * Picture2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If
End Sub
Private Sub poppucr()
For t% = 0 To 1
    For tt% = 1 To vucrlist(t%).ListItems.Count
        alreadythere = False
        For ttt% = 0 To pucrlist(0).ListCount - 1
            If pucrlist(0).List(ttt%) = vucrlist(t%).ListItems(tt%) Then
                alreadythere = True
                ttt% = pucrlist(0).ListCount - 1
                ttt% = 5
            End If
        Next ttt%
        If Not alreadythere Then
            For ttt% = 0 To 5
                pucrlist(ttt%).AddItem vucrlist(t%).ListItems(tt%)
            Next ttt%
        End If
    Next tt%
Next t%
End Sub

