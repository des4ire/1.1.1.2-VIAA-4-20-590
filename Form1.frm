VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{49F811F7-6005-4AAF-AE00-9D98766A6E26}#1.0#0"; "NTGraph.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FET"
   ClientHeight    =   13830
   ClientLeft      =   240
   ClientTop       =   -21150
   ClientWidth     =   23430
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13830
   ScaleWidth      =   23430
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000C&
      Caption         =   "FET TEST SOFTWARE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11535
      Left            =   1440
      TabIndex        =   0
      Top             =   720
      Width           =   18975
      Begin VB.CommandButton Command43 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   16560
         TabIndex        =   113
         Top             =   8160
         Width           =   735
      End
      Begin MSComDlg.CommonDialog CommonDialog5 
         Left            =   15720
         Top             =   8040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox Text25 
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   15960
         TabIndex        =   112
         Top             =   3840
         Width           =   1815
      End
      Begin VB.TextBox Text24 
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   13920
         TabIndex        =   111
         Top             =   3840
         Width           =   1935
      End
      Begin VB.TextBox Text23 
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   11880
         TabIndex        =   110
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Timer Timer6 
         Enabled         =   0   'False
         Left            =   18120
         Top             =   5040
      End
      Begin VB.CommandButton Command42 
         Caption         =   "case2"
         Height          =   375
         Left            =   17880
         TabIndex        =   109
         Top             =   4440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Command41 
         Caption         =   "case_1"
         Height          =   375
         Left            =   17880
         TabIndex        =   108
         Top             =   3360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Command36 
         Caption         =   "2450 control"
         Height          =   495
         Left            =   8880
         TabIndex        =   107
         Top             =   10080
         Visible         =   0   'False
         Width           =   1695
      End
      Begin NTGRAPHLib.NTGraph NTGraph1 
         Height          =   5415
         Left            =   240
         TabIndex        =   106
         Top             =   4320
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   9551
         _StockProps     =   194
         ShowGrid        =   -1  'True
         BeginProperty TickFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty IdentFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PlotAreaPicture =   "Form1.frx":15A1C
         ControlFramePicture=   "Form1.frx":15A38
      End
      Begin VB.TextBox Text22 
         Height          =   375
         Left            =   2400
         TabIndex        =   105
         Text            =   "Text22"
         Top             =   10680
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command40 
         Caption         =   "graph3"
         Height          =   375
         Left            =   15360
         TabIndex        =   104
         Top             =   10080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command39 
         Caption         =   "Read"
         Height          =   495
         Left            =   12840
         TabIndex        =   103
         Top             =   10440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Timer Timer5 
         Left            =   18000
         Top             =   6720
      End
      Begin VB.TextBox Text21 
         Height          =   495
         Left            =   4560
         TabIndex        =   102
         Top             =   9840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text17 
         Height          =   495
         Left            =   3840
         TabIndex        =   101
         Top             =   9840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text16 
         Height          =   375
         Left            =   16440
         TabIndex        =   100
         Top             =   9360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   16440
         TabIndex        =   99
         Top             =   10800
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command38 
         Caption         =   "STOP"
         Enabled         =   0   'False
         Height          =   975
         Left            =   14280
         TabIndex        =   97
         Top             =   8640
         Width           =   1575
      End
      Begin VB.CommandButton Command37 
         Caption         =   "START"
         Height          =   975
         Left            =   12360
         TabIndex        =   96
         Top             =   8640
         Width           =   1575
      End
      Begin VB.ListBox List6 
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   2595
         ItemData        =   "Form1.frx":15A54
         Left            =   11880
         List            =   "Form1.frx":15A56
         TabIndex        =   95
         Top             =   5400
         Width           =   6015
      End
      Begin VB.CommandButton Command34 
         Caption         =   "CLEAR GRAPH"
         Height          =   615
         Left            =   2400
         TabIndex        =   94
         Top             =   9960
         Width           =   975
      End
      Begin MSComDlg.CommonDialog CommonDialog4 
         Left            =   1680
         Top             =   9960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command35 
         Caption         =   "SAVE AS"
         Height          =   615
         Left            =   600
         TabIndex        =   93
         Top             =   9960
         Width           =   975
      End
      Begin VB.Timer Timer4 
         Interval        =   100
         Left            =   9720
         Top             =   10800
      End
      Begin VB.CommandButton Command33 
         Caption         =   "SMU_conf"
         Height          =   495
         Left            =   6480
         TabIndex        =   92
         Top             =   10800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H8000000C&
         Caption         =   "Output characteristics"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   9120
         TabIndex        =   67
         Top             =   240
         Width           =   2655
         Begin VB.Frame Frame9 
            BackColor       =   &H8000000C&
            Caption         =   "Gate-Source Voltage"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   120
            TabIndex        =   81
            Top             =   2520
            Width           =   2415
            Begin VB.ComboBox Combo14 
               BackColor       =   &H80000001&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   315
               ItemData        =   "Form1.frx":15A58
               Left            =   960
               List            =   "Form1.frx":15AB3
               TabIndex        =   86
               Text            =   "1"
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label Label12 
               BackColor       =   &H8000000C&
               Caption         =   "Fixed, V"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   87
               Top             =   600
               Width           =   735
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H8000000C&
            Caption         =   "Drain-Source Voltage Sweep"
            Height          =   2295
            Left            =   120
            TabIndex        =   79
            Top             =   240
            Width           =   2415
            Begin VB.ComboBox Combo13 
               BackColor       =   &H80000001&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   315
               ItemData        =   "Form1.frx":15B40
               Left            =   960
               List            =   "Form1.frx":15B50
               TabIndex        =   85
               Text            =   "50"
               Top             =   1800
               Width           =   1215
            End
            Begin VB.ComboBox Combo12 
               BackColor       =   &H80000001&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   315
               ItemData        =   "Form1.frx":15B68
               Left            =   960
               List            =   "Form1.frx":15B7E
               TabIndex        =   84
               Text            =   "10"
               Top             =   1320
               Width           =   1215
            End
            Begin VB.ComboBox Combo11 
               BackColor       =   &H80000001&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   315
               ItemData        =   "Form1.frx":15B9A
               Left            =   960
               List            =   "Form1.frx":15C19
               TabIndex        =   83
               Text            =   "1"
               Top             =   840
               Width           =   1215
            End
            Begin VB.ComboBox Combo10 
               BackColor       =   &H80000001&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   315
               ItemData        =   "Form1.frx":15CC2
               Left            =   960
               List            =   "Form1.frx":15D41
               TabIndex        =   82
               Text            =   "-1"
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label17 
               BackColor       =   &H8000000C&
               Caption         =   "Time, s"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   91
               Top             =   1800
               Width           =   855
            End
            Begin VB.Label Label16 
               BackColor       =   &H8000000C&
               Caption         =   "Step, mV"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   90
               Top             =   1320
               Width           =   855
            End
            Begin VB.Label Label15 
               BackColor       =   &H8000000C&
               Caption         =   "End, V"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   89
               Top             =   840
               Width           =   615
            End
            Begin VB.Label Label14 
               BackColor       =   &H8000000C&
               Caption         =   "Start, V"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   88
               Top             =   480
               Width           =   735
            End
         End
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H8000000C&
         Height          =   255
         Left            =   8640
         TabIndex        =   65
         Top             =   360
         Width           =   375
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H8000000C&
         Height          =   375
         Left            =   5400
         TabIndex        =   64
         Top             =   360
         Value           =   -1  'True
         Width           =   375
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   18120
         Top             =   3840
      End
      Begin VB.CommandButton Command32 
         Caption         =   "Command32"
         Height          =   495
         Left            =   13920
         TabIndex        =   63
         Top             =   9960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text20 
         Height          =   375
         Left            =   10800
         TabIndex        =   62
         Text            =   "2"
         Top             =   10080
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command31 
         Caption         =   "output off"
         Height          =   495
         Left            =   8640
         TabIndex        =   61
         Top             =   9960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command30 
         Caption         =   "SET OFFSET"
         Height          =   495
         Left            =   15120
         TabIndex        =   60
         Top             =   4680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H8000000C&
         Caption         =   "Tektronix AFG3022C control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   240
         TabIndex        =   54
         Top             =   2280
         Width           =   4935
         Begin VB.CommandButton Command29 
            Caption         =   "RST_DEV"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3480
            TabIndex        =   59
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton Command28 
            Caption         =   "CLR_DEV"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   58
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton Command27 
            Caption         =   "INIT AFG"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   57
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox Text19 
            BackColor       =   &H80000001&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   285
            Left            =   240
            TabIndex        =   56
            Top             =   720
            Width           =   4455
         End
         Begin VB.TextBox Text18 
            BackColor       =   &H80000001&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   285
            Left            =   240
            TabIndex        =   55
            Top             =   360
            Width           =   4455
         End
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   480
         TabIndex        =   4
         Top             =   600
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   480
         TabIndex        =   8
         Top             =   960
         Width           =   4455
      End
      Begin VB.CommandButton Command26 
         Caption         =   "graph"
         Height          =   495
         Left            =   13680
         TabIndex        =   52
         Top             =   10440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command21 
         Caption         =   "output on"
         Height          =   495
         Left            =   14160
         TabIndex        =   51
         Top             =   11040
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command20 
         Caption         =   "set voltage"
         Height          =   615
         Left            =   6840
         TabIndex        =   50
         Top             =   9960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   6840
         TabIndex        =   49
         Text            =   "Combo4"
         Top             =   9960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   16440
         TabIndex        =   48
         Text            =   "Combo3"
         Top             =   9960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   15360
         TabIndex        =   47
         Text            =   "Combo2"
         Top             =   9960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command19 
         Caption         =   "SWEEP"
         Height          =   615
         Left            =   3720
         TabIndex        =   46
         Top             =   9840
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command18 
         Caption         =   "stop2"
         Height          =   375
         Left            =   12840
         TabIndex        =   45
         Top             =   10200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   6600
         TabIndex        =   44
         Text            =   "100"
         Top             =   10440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text9 
         Height          =   495
         Left            =   1440
         TabIndex        =   43
         Top             =   10560
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Start2"
         Height          =   495
         Left            =   10320
         TabIndex        =   42
         Top             =   10680
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   120
         Top             =   9960
      End
      Begin VB.TextBox Text15 
         Height          =   975
         Left            =   3960
         TabIndex        =   41
         Top             =   10200
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton Command25 
         Caption         =   "BEEP"
         Height          =   495
         Left            =   11520
         TabIndex        =   40
         Top             =   10440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ListBox List5 
         Height          =   645
         ItemData        =   "Form1.frx":15DEA
         Left            =   10080
         List            =   "Form1.frx":15DEC
         TabIndex        =   39
         Top             =   10560
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text14 
         Height          =   495
         Left            =   10440
         TabIndex        =   38
         Top             =   10440
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Read"
         Height          =   615
         Left            =   15720
         TabIndex        =   37
         Top             =   10200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command23 
         Caption         =   "SEND MESSAGE"
         Height          =   615
         Left            =   10080
         TabIndex        =   36
         Top             =   9840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command22 
         Caption         =   "SEND "
         Height          =   375
         Left            =   12360
         TabIndex        =   35
         Top             =   10440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text13 
         Height          =   375
         Left            =   6840
         TabIndex        =   34
         Top             =   10440
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Show graph"
         Height          =   495
         Left            =   9600
         TabIndex        =   33
         Top             =   10200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   9600
         TabIndex        =   32
         Top             =   10320
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   7560
         TabIndex        =   31
         Top             =   10320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ListBox List4 
         Height          =   255
         ItemData        =   "Form1.frx":15DEE
         Left            =   8040
         List            =   "Form1.frx":15DF0
         TabIndex        =   30
         Top             =   9960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   14520
         TabIndex        =   29
         Text            =   " #01RD<CR>"
         Top             =   10320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command14 
         Caption         =   "SET"
         Height          =   375
         Left            =   5160
         TabIndex        =   28
         Top             =   10440
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Command12"
         Height          =   375
         Left            =   12240
         TabIndex        =   27
         Top             =   10320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   2520
         TabIndex        =   26
         Top             =   10440
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command11 
         Caption         =   "CLOSE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7920
         TabIndex        =   25
         Top             =   10680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   9000
         TabIndex        =   24
         Top             =   10800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   6120
         Top             =   9840
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         RThreshold      =   1
         BaudRate        =   19200
         SThreshold      =   1
      End
      Begin VB.CommandButton Command10 
         Caption         =   "OPEN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8280
         TabIndex        =   23
         Top             =   10080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   14760
         TabIndex        =   22
         Top             =   4320
         Width           =   735
      End
      Begin MSComDlg.CommonDialog CommonDialog3 
         Left            =   14160
         Top             =   4320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":15DF2
         Left            =   7920
         List            =   "Form1.frx":15E32
         TabIndex        =   21
         Text            =   "1"
         Top             =   10440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ListBox List3 
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   3180
         ItemData        =   "Form1.frx":15E7D
         Left            =   13920
         List            =   "Form1.frx":15E7F
         TabIndex        =   19
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton Command6 
         Caption         =   "CLEAR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12360
         TabIndex        =   18
         Top             =   11040
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   12720
         TabIndex        =   15
         Text            =   "0"
         Top             =   9960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "STOP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   14
         Top             =   9720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         Caption         =   "START"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   13
         Top             =   10440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   18000
         Top             =   6120
      End
      Begin VB.TextBox Text2 
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   9720
         TabIndex        =   12
         Top             =   9960
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "SET_DEV"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13560
         TabIndex        =   11
         Top             =   9960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "RST_DEV"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   10
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CLR_DEV"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   16200
         Top             =   4320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   12000
         Top             =   4320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command16 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   16680
         TabIndex        =   7
         Top             =   4320
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12480
         TabIndex        =   6
         Top             =   4320
         Width           =   735
      End
      Begin VB.CommandButton Command13 
         Caption         =   "INIT SMU"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ListBox List2 
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   3180
         ItemData        =   "Form1.frx":15E81
         Left            =   15960
         List            =   "Form1.frx":15E83
         TabIndex        =   3
         Top             =   480
         Width           =   1815
      End
      Begin VB.ListBox List1 
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   3180
         ItemData        =   "Form1.frx":15E85
         Left            =   11880
         List            =   "Form1.frx":15E87
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10080
         TabIndex        =   1
         Top             =   10320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000C&
         Caption         =   "Keithley 2400 SMU control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   240
         TabIndex        =   53
         Top             =   240
         Width           =   4935
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H8000000C&
         Caption         =   "Transfer characteristics"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   5880
         TabIndex        =   66
         Top             =   240
         Width           =   2655
         Begin VB.Frame Frame3 
            BackColor       =   &H8000000C&
            Caption         =   "Drain-Source Voltage"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   120
            TabIndex        =   75
            Top             =   2520
            Width           =   2415
            Begin VB.ComboBox Combo8 
               BackColor       =   &H80000001&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   315
               ItemData        =   "Form1.frx":15E89
               Left            =   960
               List            =   "Form1.frx":15F08
               TabIndex        =   76
               Text            =   "1"
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label Label11 
               BackColor       =   &H8000000C&
               Caption         =   "Fixed, V"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   80
               Top             =   600
               Width           =   735
            End
         End
         Begin VB.ComboBox Combo6 
            BackColor       =   &H80000001&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   315
            ItemData        =   "Form1.frx":15FB1
            Left            =   1080
            List            =   "Form1.frx":1600C
            TabIndex        =   69
            Text            =   "1000"
            Top             =   1080
            Width           =   1215
         End
         Begin VB.ComboBox Combo5 
            BackColor       =   &H80000001&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   315
            ItemData        =   "Form1.frx":160B7
            Left            =   1080
            List            =   "Form1.frx":16112
            TabIndex        =   68
            Text            =   "-1000"
            Top             =   600
            Width           =   1215
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H8000000C&
            Caption         =   "Gate-Source Voltage Sweep"
            Height          =   2295
            Left            =   120
            TabIndex        =   70
            Top             =   240
            Width           =   2415
            Begin VB.ComboBox Combo9 
               BackColor       =   &H80000001&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   315
               ItemData        =   "Form1.frx":161BD
               Left            =   960
               List            =   "Form1.frx":161CD
               TabIndex        =   78
               Text            =   "50"
               Top             =   1800
               Width           =   1215
            End
            Begin VB.ComboBox Combo7 
               BackColor       =   &H80000001&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   315
               ItemData        =   "Form1.frx":161E5
               Left            =   960
               List            =   "Form1.frx":161FB
               TabIndex        =   73
               Text            =   "10"
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label Label10 
               BackColor       =   &H8000000C&
               Caption         =   "Time,ms "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   77
               Top             =   1800
               Width           =   615
            End
            Begin VB.Label Label6 
               BackColor       =   &H8000000C&
               Caption         =   "Step,mV"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   74
               Top             =   1320
               Width           =   735
            End
            Begin VB.Label Label13 
               BackColor       =   &H8000000C&
               Caption         =   "End, mV"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   72
               Top             =   840
               Width           =   735
            End
            Begin VB.Label Label9 
               BackColor       =   &H8000000C&
               Caption         =   "Start,mV"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   71
               Top             =   480
               Width           =   975
            End
         End
      End
      Begin VB.Label Label18 
         BackColor       =   &H8000000C&
         Caption         =   "STATUS WINDOW"
         Height          =   375
         Left            =   12000
         TabIndex        =   98
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000C&
         Caption         =   "Amperage, A"
         Height          =   255
         Left            =   14280
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000C&
         Caption         =   "Time, ms"
         Height          =   255
         Left            =   16440
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000C&
         Caption         =   "Voltage, V"
         Height          =   255
         Left            =   12000
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Const vbKeyDecPt = 46
'Const ADDRESS = "USB0::0x05E6::0x2450::04456988::0::INSTR" 'keithley 2450
Const ADDRESS = "GPIB0::24::INSTR" 'keithley 2400 DU
'Const ADDRESS1 = "USB0::0x0957::0x0407::my43002948::0::INSTR" 'agilent 33220A
Const ADDRESS1 = "USB0::0x0699::0x034A::C012939::0::INSTR" 'afg3022C DU
Public chht As Long '***************
Public chwd As Long  '**************
Dim chht1 As Long  '************
Dim chwd1 As Long  '**************
Dim number As Integer
Dim start_v As Integer
Dim start_v2 As Double
Dim data_points As Integer
Dim step As Integer
Dim step2 As Double
Dim amplitude As Integer
Dim counter3 As Integer
Dim Ymin As Double
Dim Ymax As Double
Dim X As Integer
Dim Y As Double
Dim offset As Double
Dim Buffer As Double
Dim Buffer1 As Double
Dim time As Long
Dim current As String
Dim off As Integer
Dim st As Integer 'test*******************
Dim Fgen As VisaComLib.FormattedIO488
Dim rm As VisaComLib.FormattedIO488
Dim afg As VisaComLib.FormattedIO488
Dim m_ioAddress As String
Dim ID1 As String
Dim ID2 As String
Dim counter As Integer
Dim counter2 As Integer
Dim interval As Integer
Dim str1 As String
Dim Data As String
Dim data1 As String
Dim data3 As String
Dim data4 As String
Dim ReturnedData As String
Dim multi As VisaComLib.FormattedIO488
Dim RS As VisaComLib.FormattedIO488
Dim ioAddress As String
Dim io_Address As String
Private Const STATUS_TIMEOUT = &H102&
Private Const INFINITE = -1& ' Бесконечный интервал
Private Const QS_KEY = &H1&
Private Const QS_MOUSEMOVE = &H2&
Private Const QS_MOUSEBUTTON = &H4&
Private Const QS_POSTMESSAGE = &H8&
Private Const QS_TIMER = &H10&
Private Const QS_PAINT = &H20&
Private Const QS_SENDMESSAGE = &H40&
Private Const QS_HOTKEY = &H80&
Private Const QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT _
        Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON _
       Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)
Private Declare Function MsgWaitForMultipleObjects Lib "user32" _
        (ByVal nCount As Long, pHandles As Long, _
        ByVal fWaitAll As Long, ByVal dwMilliseconds _
        As Long, ByVal dwWakeMask As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long


Public Function MsgWaitObj(interval As Long, _
            Optional hObj As Long = 0&, _
          Optional nObj As Long = 0&) As Long
Dim t As Long, T1 As Long
If interval <> INFINITE Then
    t = GetTickCount()
    On Error Resume Next
    t = t + interval
    ' Предотвращение переполнения
    If Err <> 0& Then
        If t > 0& Then
            t = ((t + &H80000000) _
          + interval) + &H80000000
        Else
            t = ((t - &H80000000) _
            + interval) - &H80000000
        End If
    End If
    On Error GoTo 0
    ' В переменной T - абсолютное время окончания интервала
Else
  T1 = INFINITE
End If
Do
    If interval <> INFINITE Then
        T1 = GetTickCount()
        On Error Resume Next
     T1 = t - T1
        ' Предотвращение переполнения
        If Err <> 0& Then
            If t > 0& Then
                T1 = ((t + &H80000000) _
                - (T1 - &H80000000))
            Else
                T1 = ((t - &H80000000) _
                - (T1 + &H80000000))
            End If
        End If
        On Error GoTo 0
        ' В переменной T1 - оставшаяся часть интервала
        If IIf((T1 Xor interval) > 0&, _
          T1 > interval, T1 < 0&) Then
            ' Интервал истек, пока
            ' выполнялась DoEvents
            MsgWaitObj = STATUS_TIMEOUT
            Exit Function
        End If
    End If
    ' Ждем события, истечения интервала или
    ' появления сообщения в очереди потока
    MsgWaitObj = MsgWaitForMultipleObjects(nObj, _
           hObj, 0&, T1, QS_ALLINPUT)
    ' Даем возможность сообщению обработаться
    DoEvents
    If MsgWaitObj <> nObj Then Exit Function
    ' Было сообщение в очереди потока - продолжаем ждать
Loop
End Function
 




Private Sub Command1_Click()
With multi
.WriteString "*CLS"
End With
End Sub

Private Sub Command10_Click()
Text5.Text = Combo1
If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
MSComm1.Settings = "9600,n,8,1"
 'MSComm1.CommEvent = 2
MSComm1.CommPort = Text5.Text

MSComm1.RThreshold = 1 ' Length paket for CommEvent (OnComm)
MSComm1.SThreshold = 1
MSComm1.InputLen = 0
MSComm1.PortOpen = Not MSComm1.PortOpen
If MSComm1.PortOpen Then
Text4.Text = "PORT OPEN"
Else
Text4.Text = "ERROR OPEN PORT"
End If

End Sub



Private Sub Command15_Click()
With multi
'.WriteString "DISP:SCR SWIPE_GRAPh"
.WriteString "DISP:CLE"
.WriteString "DISP:SCR SWIPE_USER"
'.WriteString "DISP:USER1:TEXT 'Hello!'"
.WriteString "disp:user1:text '" & Text13.Text & "'"
'Sleep (1500)
.WriteString "DISP:CLE"
.WriteString "disp:user1:text 'By-By!'"
End With
End Sub




  










Private Function f(X As Single) As Double
f = -0.25 * X ^ 3 + 1.4 * X ^ 2 + 1.8 * X - 6

End Function

Private Sub Command17_Click()
With multi
'.WriteString "*RST"
.WriteString "*CLS"
.WriteString ":SOURce:FUNCtion VOLTage"
'.WriteString ":Source: VOLTage: Range 20"

.WriteString "SENSe:FUNCtion 'curr'"

.WriteString "SENSe:CURRent:RANGe:AUTO ON"
'.WriteString "SENSe:CURRent:UNIT OHM "
'.WriteString "SENSe:CURRent:OCOM ON"
'.WriteString "SOURce:FUNCtion VOLT"
'.WriteString "SOURce:VOLT 5"
'.WriteString "SOURce:VOLT:ILIM 0.01"





'.WriteString "SENSe:FUNCtion 'res'"
'.WriteString "SENSe:RESistance:RANGe:AUTO ON"
'.WriteString "SENSe:RESistance:OCOMpensated ON"


'.WriteString "SENSe:COUNt 5"
.WriteString "OUTPut ON"
'.WriteString "DISP: SCR SWIPE_GRAPh"
'.WriteString "TRACe:TRIGger 'defbuffer1'"
'For i = 1 To 10
End With

counter = 0
Timer2.interval = CInt(Text11.Text)
Timer2.Enabled = True



End Sub

Private Sub Command18_Click()
Timer2.Enabled = False
End Sub

Private Sub Command19_Click()
With multi
.WriteString "*RST"
.WriteString "SOUR:FUNC VOLT"
.WriteString "SOUR:VOLT:RANGe 20"
.WriteString "SOUR:VOLT:ILIM 0.01"
.WriteString "SENS:FUNC 'curr'"
.WriteString "SENSe:CURRent:RANGe 0.01"
'.WriteString "SOUR:SWE:volt:LIN:STEP 0.00, 10.00, .1, 10e-3, 1, FIXED"
.WriteString "SOUR:SWE:VOLT:LIN 10, -10, 100, 0.1"
.WriteString "INIT"
'SOUR:CURR:RANGE 1 SENS:FUNC "VOLT" SENS:VOLT:RANGE 20 SOUR:SWE:CURR:LIN:STEP -1.05, 1.05, .25, 10e-3, 1, FIXED INIT
End With
End Sub

Private Sub Command20_Click()
With multi
.WriteString "SOURce:FUNC VOLT "
.WriteString "SOURce:VOLT:LEV 2"
End With

End Sub

Private Sub Command21_Click()
With multi
.WriteString "OUTPut ON"
End With
End Sub


Private Sub Command22_Click()

Dim i As Integer
With multi
'.WriteString " " & Text13.Text & " "
.WriteString "*RST"
.WriteString "SENSe:FUNCtion 'res'"
.WriteString "SENSe:RESistance:RANGe:AUTO ON"
.WriteString "SENSe:RESistance:OCOMpensated ON"
'.WriteString "SENSe:COUNt 5"
.WriteString "OUTPut ON"
'.WriteString "DISP: SCR SWIPE_GRAPh"
'.WriteString "TRACe:TRIGger 'defbuffer1'"
For i = 1 To 10
.WriteString "measure:resistance?"
str1 = .ReadString
Text13.Text = str1
List5.AddItem str1

'Sleep (1000)
Next
'.WriteString "TRACe:DATA? 1, 5, 'defbuffer1', SOUR, READ"
.WriteString "OUTP OFF"

End With
End Sub

Private Sub Command23_Click()


With multi

.WriteString "*CLS"
.WriteString "SENSe:FUNCtion 'curr'"
.WriteString "SENSe:CURRent:RANGe:AUTO ON"
.WriteString "SENSe:CURRent:UNIT amp "
.WriteString "SENSe:CURRent:OCOM ON"
.WriteString "SOURce:FUNCtion VOLT"
.WriteString "SOURce:VOLT 5"
.WriteString "SOURce:VOLT:ILIM 0.01"
.WriteString "SENSe:COUNT 500 "
.WriteString "OUTPut ON"
.WriteString "TRACe:TRIGger 'defbuffer1'"
.WriteString "DISP:CLE"
.WriteString "DISP:SCR GRAPH"
.WriteString "TRACe:DATA? 1, 500, 'defbuffer1', SOUR, READ "
.WriteString "*WAI"
 'data3 = .ReadString
 'Text14.Text = data3
.WriteString "OUTPut OFF"

'.WriteString ":TRAC:DATA? 1, 10, 'defbuffer1', READ, REL"
'.WriteString ":OUTP OFF"

'.WriteString ":MEASure:Voltage?"
'.WriteString ":SENSe:res:NPLCycles .5"
 'SENS:FUNC "RES" SENS:RES:RANG:AUTO ON
End With
End Sub

Private Sub Command24_Click()
'Dim csv As New ChilkatCsv
Dim pos As Double
Dim neg As Double
Dim diff As Double
Dim str As String
Dim Value() As Double

'ifileno = FreeFile
'Open "C:\Test.txt" For Output As #ifileno
Dim i As Integer
Dim fields() As String
'ReDim value(1 To UBound(fields), 1 To 2)

 With multi
'.WriteString ":SENSe:FUNCtion 'CURRent'"
'.WriteString "OUTPut ON"
.WriteString ":TRACe:DATA? 1, 50, 'defbuffer1'"
str = .ReadString
'Text14.Text = str
fields() = Split(str, ",")
For i = 0 To UBound(fields)
If i Mod 2 = 0 Then
List1.AddItem Trim$(fields(i))
'value(i, 1) = CDbl(fields(i))
'pos = CDbl(fields(i))
'pos = Abs(pos)
'List1.AddItem pos
Else
List3.AddItem Trim$(fields(i))
'value(i, 2) = CDbl(fields(i))
'neg = CDbl(fields(i))
'neg = Abs(neg)
'List3.AddItem neg
'diff = pos - neg
'diff = Abs(diff)
'List8.AddItem diff
End If

Next
'.WriteString ":TRACe:DATA? 25, 50, 'defbuffer1'"
'data4 = .ReadString
'Text15.Text = data4

.WriteString ":TRACE:clear 'defbuffer1'"

'.WriteString "*CLS"
End With
'MSChart1.ChartType = VtChChartType2dXY
 '   MSChart1.RowCount = 2
  '  MSChart1.ColumnCount = UBound(fields)
   ' MSChart1.ChartData = Value
End Sub

Private Sub Command25_Click()
With multi
.WriteString "system:beeper 1000, 1"
.WriteString "system:beeper 1500, 1"
.WriteString "system:beeper 1000, 1"
End With
End Sub










Private Sub Command27_Click()
ioAddress = ADDRESS1
  Dim mgs As VisaComLib.ResourceManager

    On Error GoTo ioError

    ioAddress = InputBox("Enter the IO address of the DEVICE", "Set IO address", ioAddress)

    If Len(ioAddress) > 3 Then
        Set mgs = New VisaComLib.ResourceManager
        Set afg = New VisaComLib.FormattedIO488
        Set afg.IO = mgs.Open(ioAddress)
    End If
    With afg
    .WriteString "*RST"
.WriteString "*CLS"
'.WriteString ":TRACE:clear 'defbuffer1'"
.WriteString "*IDN?"
ID2 = .ReadString
End With
Text19.Text = ID2
Text18.Text = ioAddress
    Exit Sub
ioError:
    MsgBox "Set IO error:" & vbCrLf & Err.Description
End Sub

Private Sub Command30_Click()
With afg
'.WriteString "SOURce1:FUNCtion:shape DC"
'.WriteString "SOURce1:VOLTAGE:LEVEL:IMMEDIATE:OFFSet " & offset & "mV
End With
'counter2 = 0
off = Val(start_v)
st = Val(step)
Timer3.Enabled = True
'offset = 0
End Sub

Private Sub Command31_Click()
With afg
.WriteString "output off"
End With
End Sub


Private Sub Command32_Click()
Dim X As Single
Dim i As Integer

    X = 0
    For i = 1 To 17
        X = X + 10 / 17
    Next i

    If X = 10 Then
        MsgBox X & " = 10"
    Else
        MsgBox X & " <> 10"
    End If
End Sub







Private Sub Command33_Click()
With multi
.WriteString "*CLS"
.WriteString ":SOUR:FUNC VOLT"
'.WriteString ":SOUR:VOLT:MODE FIXED"
.WriteString ":SOUR:VOLT:RANG 20"
'.WriteString ":SOUR:VOLT:LEV 10"
.WriteString ":SOUR:VOLT:LEV 0.001"
.WriteString "SOURCe:VOLTage:ILIMit 0.01"
.WriteString ":SENS:FUNC 'CURR'"
.WriteString ":SENS:CURR:RANG 0.01" '0.01"
'.WriteString ":FORM:ELEM CURR" 'read only current
.WriteString ":OUTP ON"
.WriteString ":READ?"
'.WriteString "DISP:CLE"
'.WriteString "DISP:SCR SWIPE_USER"

current = .ReadString
List3.AddItem current

End With
End Sub

Private Sub Command35_Click()
 On Error GoTo saverr
  Dim strsavefile As String
  With CommonDialog4 ' CommonDialog object
    .Filter = "Pictures (*.bmp)|*.bmp"
    .DefaultExt = "bmp"
    .CancelError = True
    .ShowSave
    strsavefile = .FileName
    If strsavefile = "" Then Exit Sub
  End With
  'Picture1.Picture = Picture.Image
  SavePicture Clipboard.GetData, strsavefile
  Exit Sub
saverr:
  MsgBox Err.Description
End Sub











Private Sub Command36_Click()
With multi
.WriteString "*CLS"
.WriteString ":SOUR:FUNC VOLT"
.WriteString ":SOUR:VOLT:MODE FIXED"
.WriteString ":SOUR:VOLT:RANG 20"
.WriteString ":SOUR:VOLT:LEV " & Combo8.Text & ""
.WriteString ":SENS:CURR:PROT 10E-3"
.WriteString ":SENS:FUNC 'CURR'"
.WriteString ":SENS:CURR:RANG Auto on"
.WriteString ":FORM:ELEM CURR" 'read only current

'.WriteString "DISP:CLE"
'.WriteString "DISP:SCR SWIPE_USER"
.WriteString ":OUTP ON"

End With
End Sub

Private Sub Command37_Click()
List6.Clear
Command38.Enabled = True
If Option3.Value = True Then
List6.AddItem "Transfer mode selected"
List6.AddItem ("Start Gate-Source voltage: " & Combo5.Text & " mV")
List6.AddItem ("End Gate-Source voltage: " & Combo6.Text & " mV")
List6.AddItem ("Gate-Source sweep step: " & Combo7.Text & " mV")
List6.AddItem ("Drain-Source fixed voltage: " & Combo8.Text & " V")
List6.AddItem ("Measurement started")
Command41_Click
ElseIf Option4.Value = True Then
List6.AddItem "Output mode selected"
List6.AddItem ("Start Drain-Source voltage: " & Combo10.Text & " V")
List6.AddItem ("End Drain-Source voltage: " & Combo11.Text & " V")
List6.AddItem ("Darin-Source sweep step: " & Combo12.Text & " mV")
List6.AddItem ("Gate-Source fixed voltage: " & Combo13.Text & " V")
List6.AddItem ("Measurement started")
Command42_Click
End If
End Sub

Private Sub Command38_Click()
List6.AddItem ("Aborted by user!")
Timer3.Enabled = False
Timer6.Enabled = False
Command37.Enabled = True
Command38.Enabled = False
With multi
.WriteString "OUTPUT OFF"
End With
With afg
.WriteString "OUTPUT OFF"
End With
End Sub

Private Sub Command39_Click()
With multi
.WriteString ":READ?"
current = .ReadString
List3.AddItem current
End With
End Sub




Private Sub Command40_Click()
With NTGraph1
     
     .PlotAreaColor = vbBlack
    ' .FrameStyle = Frame
     .Caption = ""
     .XLabel = ""
     .YLabel = ""

     '.ClearGraph 'delete all elements and create a new one
     .ElementLineColor = RGB(255, 255, 0)
     .AddElement  ' Add second elements
     .ElementLineColor = vbGreen

     For X = 0 To 1
          Y = CDbl(Text22.Text)
          .PlotY Y, 0
           'Y = Cos(X / 3.15) * Rnd + 1
          '.PlotXY X, Y, 1
          .SetRange 0, 100, -0.05, 0.05
      Next X

End With
End Sub

Private Sub Command41_Click()

Dim gsweep As String
Dim time As Integer

Command37.Enabled = False
If CInt(Combo5.Text) < 0 And CInt(Combo6.Text) > 0 Then
amplitude = Abs(CInt(Combo5.Text) - CInt(Combo6.Text))
ElseIf CInt(Combo5.Text) < 0 And CInt(Combo6.Text) < 0 Then
amplitude = Abs(CDbl(Combo5.Text) - CDbl(Combo6.Text))
ElseIf CInt(Combo5.Text) > 0 And CInt(Combo6.Text) < 0 Then
amplitude = Abs(CInt(Combo6.Text) - CInt(Combo5.Text))
ElseIf CInt(Combo5.Text) = 0 And CInt(Combo6.Text) > 0 Then
amplitude = Abs(CInt(Combo6.Text) - CInt(Combo5.Text))
ElseIf CInt(Combo5.Text) = 0 And CInt(Combo6.Text) < 0 Then
amplitude = Abs(CInt(Combo6.Text) - CInt(Combo5.Text))
End If
'List6.Clear
List1.Clear
List2.Clear
List3.Clear
Command41.Enabled = False
Text16.Text = (amplitude / CInt(Combo7.Text))
data_points = CInt(Text16.Text)
start_v = (CDbl(Combo5.Text))
Text21.Text = start_v

step = CInt(Combo7.Text)
If CInt(Combo6.Text) < 0 Then
step = step * (-1)
End If
With multi
.WriteString "*CLS"
.WriteString ":ROUTe:TERMinals REAR"
.WriteString ":SOUR:FUNC VOLT"
'.WriteString ":SOUR:VOLT:MODE FIXED"
.WriteString ":SOUR:VOLT:RANG 20"
.WriteString ":SOUR:VOLT:LEV " & Combo8.Text & ""
.WriteString "SOURCe:VOLTage:ILIMit 0.01"
.WriteString ":SENS:FUNC 'CURR'"
.WriteString ":SENS:CURR:RANG 0.01" '0.01"
.WriteString ":FORM:ELEM CURR" 'read only current
.WriteString ":OUTP ON"
Sleep (300)
.WriteString ":READ?"
'.WriteString "DISP:CLE"
'.WriteString "DISP:SCR SWIPE_USER"
current = .ReadString
Buffer = Abs(CDbl(Val(Replace((current), ",", "."))))
Ymin = Buffer - (Buffer * 0.5) 'Buffer
Ymax = (Buffer * 0.5) + Buffer
End With

List6.AddItem ("Data points: " + Text16.Text)
'Timer3.interval = CInt(Combo9.Text) * 1000

'End With
'With multi 'keithley 2450
'.WriteString "*CLS"
'.WriteString ":SOUR:FUNC VOLT"
'.WriteString ":SOUR:VOLT:MODE FIXED"
'.WriteString ":SOUR:VOLT:RANG 20"
'.WriteString ":SOUR:VOLT:LEV " & Combo8.Text & ""
'.WriteString ":SENS:CURR:PROT 10E-3"
'.WriteString ":SENS:FUNC 'CURR'"
'.WriteString ":SENS:CURR:RANG Auto on"
'.WriteString ":FORM:ELEM CURR" 'read only current
'.WriteString ":OUTP ON"
'.WriteString ":READ?"
'.WriteString ":OUTP OFF"
'current = .ReadString
'Buffer = Abs(CDbl(Val(Replace((current), ",", "."))))
'Ymin = Buffer
'Ymax = (Buffer * 0.0000005) + Buffer
'End With
'List6.AddItem ("Data points: " + Text16.Text)
'Timer3.Interval = CInt(Combo9.Text) * 1000
counter3 = 0
With afg
.WriteString "*CLS"
.WriteString "SOURce1:FUNCtion:shape DC"
.WriteString "output on"
'.WriteString "SOURce1:VOLTAGE:LEVEL:IMMEDIATE:OFFSet " & start_v & "mV"
End With

With NTGraph1
 
     
     .PlotAreaColor = vbBlack
    ' .FrameStyle = Frame
     .Caption = ""
     .XLabel = ""
     .YLabel = ""

     .ClearGraph 'delete all elements and create a new one
     .ElementLineColor = RGB(255, 255, 0)
     .AddElement  ' Add second elements
     .ElementLineColor = vbGreen

     For X = 0 To 1
          Y = CDbl(Buffer)
          .PlotY Y, 0
           'Y = Cos(X / 3.15) * Rnd + 1
          '.PlotXY X, Y, 1
          .SetRange 0, data_points, -0.002, Ymax
      Next X

End With
time = 0
interval = CInt((Combo9.Text))
Timer3.interval = interval
Timer3.Enabled = True

End Sub



Private Sub Command42_Click()

Dim gsweep As String
Dim time As Integer

Command37.Enabled = False
'amplitude = Abs(CInt(Combo10.Text) - CInt(Combo11.Text))

If CInt(Combo10.Text) < 0 And CInt(Combo11.Text) > 0 Then
amplitude = Abs(CInt(Combo10.Text) - CInt(Combo11.Text))
ElseIf CInt(Combo10.Text) < 0 And CInt(Combo11.Text) < 0 Then
amplitude = Abs(CDbl(Combo10.Text) - CDbl(Combo11.Text))
ElseIf CInt(Combo10.Text) > 0 And CInt(Combo11.Text) < 0 Then
amplitude = Abs(CInt(Combo10.Text) - CInt(Combo11.Text))
ElseIf CInt(Combo10.Text) = 0 And CInt(Combo11.Text) > 0 Then
amplitude = Abs(CInt(Combo11.Text) - CInt(Combo10.Text))
ElseIf CInt(Combo10.Text) = 0 And CInt(Combo11.Text) < 0 Then
amplitude = Abs(CInt(Combo11.Text) - CInt(Combo10.Text))
End If
'List6.Clear
List1.Clear
List2.Clear
List3.Clear
Command37.Enabled = False
Text16.Text = (amplitude / (CInt(Combo12.Text) / 1000))
data_points = CInt(Text16.Text)
start_v2 = CDbl(Combo10.Text)
Text21.Text = start_v2
step2 = CDbl(Combo12.Text) / 1000
If CInt(Combo11.Text) < 0 Then
step2 = step2 * (-1)
End If
With multi
.WriteString "*CLS"
.WriteString ":ROUTe:TERMinals REAR"
.WriteString ":SOUR:FUNC VOLT"
'.WriteString ":SOUR:VOLT:MODE FIXED"
.WriteString ":SOUR:VOLT:RANG 20"
.WriteString ":SOUR:VOLT:LEV " & Combo10.Text & ""
.WriteString "SOURCe:VOLTage:ILIMit 0.01"
.WriteString ":SENS:FUNC 'CURR'"
.WriteString ":SENS:CURR:RANG 0.01" '0.01"
.WriteString ":FORM:ELEM CURR" 'read only current
.WriteString ":OUTP ON"
Sleep (300)
.WriteString ":READ?"
'.WriteString "DISP:CLE"
'.WriteString "DISP:SCR SWIPE_USER"
current = .ReadString
Buffer = Abs(CDbl(Val(Replace((current), ",", "."))))
Ymin = Buffer - (Buffer * 0.5) 'Buffer
Ymax = (Buffer * 0.5) + Buffer
End With
List6.AddItem ("Data points: " + Text16.Text)
'Timer3.interval = CInt(Combo9.Text) * 1000

'End With
'With multi 'keithley 2450
'.WriteString "*CLS"
'.WriteString ":SOUR:FUNC VOLT"
'.WriteString ":SOUR:VOLT:MODE FIXED"
'.WriteString ":SOUR:VOLT:RANG 20"
'.WriteString ":SOUR:VOLT:LEV " & Combo8.Text & ""
'.WriteString ":SENS:CURR:PROT 10E-3"
'.WriteString ":SENS:FUNC 'CURR'"
'.WriteString ":SENS:CURR:RANG Auto on"
'.WriteString ":FORM:ELEM CURR" 'read only current
'.WriteString ":OUTP ON"
'.WriteString ":READ?"
'.WriteString ":OUTP OFF"
'current = .ReadString
'Buffer = Abs(CDbl(Val(Replace((current), ",", "."))))
'Ymin = Buffer
'Ymax = (Buffer * 0.0000005) + Buffer
'End With
'List6.AddItem ("Data points: " + Text16.Text)
'Timer3.Interval = CInt(Combo9.Text) * 1000
counter3 = 0
With afg
.WriteString "*CLS"
.WriteString "SOURce1:FUNCtion:shape DC"
.WriteString "SOURce1:VOLTAGE:LEVEL:IMMEDIATE:OFFSet " & Combo14.Text & "V"
.WriteString "output on"

End With

With NTGraph1
 
     
     .PlotAreaColor = vbBlack
    ' .FrameStyle = Frame
     .Caption = ""
     .XLabel = "Voltage"
     .YLabel = "Current"

     .ClearGraph 'delete all elements and create a new one
     .ElementLineColor = RGB(255, 255, 0)
     .AddElement  ' Add second elements
     .ElementLineColor = vbGreen

     For X = 0 To 1
          Y = CDbl(Buffer)
          .PlotY Y, 0
           'Y = Cos(X / 3.15) * Rnd + 1
          '.PlotXY X, Y, 1
          .SetRange 0, data_points, -0.002, Ymax
      Next X

End With
time = 0
interval = CInt((Combo13.Text))
Timer6.interval = interval
Timer6.Enabled = True


End Sub



Private Sub Command43_Click()
Dim FileNum6 As Integer
Dim N6 As Integer
On Error GoTo ErrHandler

'если в поле ListBox ничего нет, то не сохраняется
If List6.ListCount = 0 Then
  MsgBox "No data"
  Exit Sub
End If

With CommonDialog5
 'если нажимается кнопка отмены, запускается процедура ошибки
 .CancelError = True
  'очищается имя файла
  .FileName = ""
  'назначается фильтр файла
  .Filter = "Text File |*.txt| НТМL страницы |*.htm; *.html|"
  'назначается расширение, если расширение не назначено, то по умолчанию txt
  .DefaultExt = "txt"
  'назначается функция запроса на перезапись файла и пути
  .Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist
  'задается название диалогу
  .DialogTitle = "Save file"
  'отображается диалог сохранения
  .ShowSave

  FileNum6 = FreeFile()
  Open .FileName For Output As #FileNum6
  'каждая линия ListBox является отдельным объектом, по-этому оператор ListCount считает количество линий -1
  'List(Number) выдает текст на текущей линии
  For N6 = 0 To List6.ListCount - 1
      Print #FileNum6, List6.List(N6)
  Next
  Close #FileNum6
End With

Exit Sub
ErrHandler:
If Err <> cdlCancel Then
 'в случае ошибки закрывается открытый файл
 Close #FileNum6
 MsgBox Err.Description
End If
End Sub

Private Sub MSComm1_OnComm()

' Check for charactersin the buffer

'Dim StrData As Variant 'define variable type as it is variant
      ' Text7.Text = ""
 '     If MSComm1.InBufferCount Then
  ' If MSComm1.CommEvent = comEvReceive Then
   '    StrData = MSComm1.Input
    'List4.AddItem (StrData)
   ' Text8.SelText = Text7.Text & StrData & " "
   ' Text8.Text = StrData
  ' List3.AddItem (Text7.Text)
'End If

 'Static strBuff As String

  '   Select Case MSComm1.CommEvent
   '     Case comEvReceive
    '        Do
     '           DoEvents
      '          strBuff = strBuff & MSComm1.Input
       '     Loop Until InStr(strBuff, Chr(13))
            
        '    If InStr(strBuff, Chr(13)) Then
         '       strBuff = Left(strBuff, Len(strBuff) - 1)
          '      strBuff = Right(strBuff, Len(strBuff) - 3)
           '     Text8.Text = strBuff
            '    List3.AddItem (strBuff)
             '   strBuff = ""
                'txtShortNumber.SetFocus -------------- bo used
           ' End If
       ' End Select
    


End Sub
Private Sub MSComm_OnComm()
Select Case MSComm1.CommEvent
Case comBreak
' A Break was received.
MsgBox ("Break received")
Case comCDTO
' CD(RLSD) Timeout.
Case comCTSTO
' CTSTimeout.
Case comDSRTO
' DSRTimeout.
Case comFrame
' Framing Error
Case comOverrun
' Data Lost.
Case comRxOver
' Receive bufferoverflow.
Case comRxParity
' ParityError.
Case comTxFull
' Transmit bufferfull.
Case comEvCD
' Change in the CD
Case comEvCTS
' Change in the CTS
Case comEvDSR
' Change in the DSR
Case comEvRing
' Change in the RI
Case comEvReceive
Case comEvSend
End Select
End Sub



Private Sub Command11_Click()

Text4.Text = "PORT CLOSE"
If MSComm1.PortOpen Then
MSComm1.PortOpen = False
End If
End
'Stop

End Sub

Private Sub Command12_Click()
Text5.Text = Combo1
End Sub

Private Sub Command13_Click()
ioAddress = ADDRESS
  Dim mgs As VisaComLib.ResourceManager

    On Error GoTo ioError

    ioAddress = InputBox("Enter the IO address of the DEVICE", "Set IO address", ioAddress)

    If Len(ioAddress) > 3 Then
        Set mgs = New VisaComLib.ResourceManager
        Set multi = New VisaComLib.FormattedIO488
        Set multi.IO = mgs.Open(ioAddress)
    End If
    With multi
    .WriteString "*RST"
.WriteString "*CLS"
'.WriteString ":TRACE:clear 'defbuffer1'"
.WriteString "*IDN?"
ID2 = .ReadString
End With
Text1.Text = ID2
Text10.Text = ioAddress
    Exit Sub
ioError:
    MsgBox "Set IO error:" & vbCrLf & Err.Description
End Sub


Private Sub Command14_Click()
MSComm1.Output = Text12.Text + vbCr
End Sub

Private Sub Command16_Click()
Dim FileNum1 As Integer
Dim N1 As Integer
On Error GoTo ErrHandler

'если в поле ListBox ничего нет, то не сохраняется
If List2.ListCount = 0 Then
  MsgBox "No data"
  Exit Sub
End If

With CommonDialog2
 'если нажимается кнопка отмены, запускается процедура ошибки
 .CancelError = True
  'очищается имя файла
  .FileName = ""
  'назначается фильтр файла
  .Filter = "Text File |*.txt| НТМL страницы |*.htm; *.html|"
  'назначается расширение, если расширение не назначено, то по умолчанию txt
  .DefaultExt = "txt"
  'назначается функция запроса на перезапись файла и пути
  .Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist
  'задается название диалогу
  .DialogTitle = "Save file"
  'отображается диалог сохранения
  .ShowSave

  FileNum1 = FreeFile()
  Open .FileName For Output As #FileNum1
  'каждая линия ListBox является отдельным объектом, по-этому оператор ListCount считает количество линий -1
  'List(Number) выдает текст на текущей линии
  For N1 = 0 To List2.ListCount - 1
      Print #FileNum1, List2.List(N1)
  Next
  Close #FileNum1
End With

Exit Sub
ErrHandler:
If Err <> cdlCancel Then
 'в случае ошибки закрывается открытый файл
 Close #FileNum1
 MsgBox Err.Description
End If
End Sub

Private Sub Command2_Click()
With multi
.WriteString "*CLS"
'.WriteString "*RST"
End With
'MSComm1.PortOpen = False
Form1.Hide
Unload Form1
End
End Sub
Private Sub Command3_Click()
With multi
.WriteString "*RST"
End With
End Sub

Private Sub Command4_Click()
With multi
'.WriteString "conf:resistance 100000000.0"
.WriteString "*CLS"
.WriteString "Conf:res"
'.WriteString "res:filt on"
'.WriteString "res:nplc 0.2"
.WriteString "res:rang:auto on"
'.WriteString "system:remote"
End With
End Sub

Private Sub Command5_Click()
 Dim FileNum As Integer
Dim N As Integer
On Error GoTo ErrHandler

'если в поле ListBox ничего нет, то не сохраняется
If List1.ListCount = 0 Then
  MsgBox "No data"
  Exit Sub
End If

With CommonDialog1
 'если нажимается кнопка отмены, запускается процедура ошибки
 .CancelError = True
  'очищается имя файла
  .FileName = ""
  'назначается фильтр файла
  .Filter = "Text File |*.txt| НТМL страницы |*.htm; *.html|"
  'назначается расширение, если расширение не назначено, то по умолчанию txt
  .DefaultExt = "txt"
  'назначается функция запроса на перезапись файла и пути
  .Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist
  'задается название диалогу
  .DialogTitle = "Save file"
  'отображается диалог сохранения
  .ShowSave

  FileNum = FreeFile()
  Open .FileName For Output As #FileNum
  'каждая линия ListBox является отдельным объектом, по-этому оператор ListCount считает количество линий -1
  'List(Number) выдает текст на текущей линии
  For N = 0 To List1.ListCount - 1
      Print #FileNum, List1.List(N)
  Next
  Close #FileNum
End With

Exit Sub
ErrHandler:
If Err <> cdlCancel Then
 'в случае ошибки закрывается открытый файл
 Close #FileNum
 MsgBox Err.Description
End If
End Sub


Private Sub Command6_Click()
List1.Clear
List2.Clear
List3.Clear
Text2.Text = 0
Text3.Text = 0
Text8.Text = 0
End Sub

Private Sub Command7_Click()
Timer1.Enabled = True
With multi
'.WriteString "*CLS"
'Sleep (100) ------------- not used
End With
End Sub

Private Sub Command8_Click()
Timer1.Enabled = False
End Sub

Private Sub Command9_Click()
 Dim FileNum2 As Integer
Dim N2 As Integer
On Error GoTo ErrHandler

'если в поле ListBox ничего нет, то не сохраняется
If List3.ListCount = 0 Then
  MsgBox "No data"
  Exit Sub
End If

With CommonDialog3
 'если нажимается кнопка отмены, запускается процедура ошибки
 .CancelError = True
  'очищается имя файла
  .FileName = ""
  'назначается фильтр файла
  .Filter = "Text File |*.txt| НТМL страницы |*.htm; *.html|"
  'назначается расширение, если расширение не назначено, то по умолчанию txt
  .DefaultExt = "txt"
  'назначается функция запроса на перезапись файла и пути
  .Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist
  'задается название диалогу
  .DialogTitle = "Save file"
  'отображается диалог сохранения
  .ShowSave

  FileNum2 = FreeFile()
  Open .FileName For Output As #FileNum2
  'каждая линия ListBox является отдельным объектом, по-этому оператор ListCount считает количество линий -1
  'List(Number) выдает текст на текущей линии
  For N2 = 0 To List3.ListCount - 1
      Print #FileNum2, List3.List(N2)
  Next
  Close #FileNum2
End With

Exit Sub
ErrHandler:
If Err <> cdlCancel Then
 'в случае ошибки закрывается открытый файл
 Close #FileNum2
 MsgBox Err.Description
End If
End Sub

Private Sub Form_Load()
Timer1.Enabled = False
Option3.Value = True
'Form1.Show
End Sub




Private Sub Timer1_Timer()

With multi
.WriteString "*CLS"
.WriteString "READ?"
'Sleep (500)


'Sleep (100)
Data = .ReadString
'If InStr(data, Chr(13)) Then ------------not used
                Data = Left(Data, Len(Data) - 1)
                Text2.Text = Data
                List1.AddItem Data
                'End If ------------ not used
End With
MSComm1.Output = Text6.Text + vbCr
Static strBuff As String

     Select Case MSComm1.CommEvent
        Case comEvReceive
            Do
                DoEvents
                strBuff = strBuff & MSComm1.Input
            Loop Until InStr(strBuff, Chr(13))
            
            If InStr(strBuff, Chr(13)) Then
                strBuff = Left(strBuff, Len(strBuff) - 1)
                strBuff = Right(strBuff, Len(strBuff) - 3)
                Text8.Text = strBuff
                List3.AddItem (strBuff)
                strBuff = ""
                'txtShortNumber.SetFocus -------------- bo used
            End If
        End Select

Text3.Text = Text3.Text + 1
List2.AddItem Text3.Text
'MSComm1.Output = Text6.Text + vbCr
'Sleep (500)
End Sub
'=================================================================
' Return True if test_string ends with target.
'Private Function EndsWith(ByVal test_string As String, _
 '   ByVal target As String) As Boolean
  '  EndsWith = (Right$(test_string, Len(target)) = target)
'End Function

' Return True if test_string starts with target.
'Private Function StartsWith(ByVal test_string As String, _
 '   ByVal target As String) As Boolean
  '  StartsWith = (Left$(test_string, Len(target)) = target)
'End Function

'=================================================================
Private Sub Timer2_Timer()

counter = counter + 1
Text9.Text = counter


'.WriteString " " & Text13.Text & " "
'.WriteString "*RST"
'.WriteString "SENSe:FUNCtion 'res'"
'.WriteString "SENSe:RESistance:RANGe:AUTO ON"
'.WriteString "SENSe:RESistance:OCOMpensated ON"
'.WriteString "SENSe:COUNt 5"
'.WriteString "OUTPut ON"
'.WriteString "DISP: SCR SWIPE_GRAPh"
'.WriteString "TRACe:TRIGger 'defbuffer1'"
'For i = 1 To 10
If counter = 100 Then
Timer2.Enabled = False
With multi
.WriteString "OUTP OFF"
End With
Else
With multi
.WriteString "measure:current?"
str1 = .ReadString
Text13.Text = str1
List5.AddItem str1

'Sleep (1000
'.WriteString "TRACe:DATA? 1, 5, 'defbuffer1', SOUR, READ"
End With
End If
End Sub

Private Sub Timer3_Timer()


Text12.Text = start_v
Text16.Text = step
'offset = Val(off) 'step
'value = CStr(offset)
'Text17.Text = off
List1.AddItem start_v
Text23.Text = start_v
start_v = start_v + step
time = time + interval
With afg
'.WriteString "FUNCtion DC"
'.WriteString "DC"

.WriteString "SOURce1:VOLTAGE:LEVEL:IMMEDIATE:OFFSet " & start_v & "mV"
'.WriteString "voltage:offset " & Text17.Text & " "

'Next
'.WriteString "voltage:offset 2.5"
'.WriteString "output on"

End With
With multi
.WriteString ":READ?"
current = .ReadString
List3.AddItem current
Text24.Text = current
Buffer = Abs(CDbl(Val(Replace((current), ",", "."))))
'buffer = Val(current)
'buffer = (Replace(Val(CDbl(current), ",", ".")))
Text22.Text = Buffer
Ymax = (Buffer * 0.5) + Buffer
End With

With NTGraph1
 
     
     .PlotAreaColor = vbBlack
    ' .FrameStyle = Frame
     .Caption = ""
     .XLabel = "Voltage"
     .YLabel = "Current"

     '.ClearGraph 'delete all elements and create a new one
     .ElementLineColor = RGB(255, 255, 0)
     .AddElement  ' Add second elements
     .ElementLineColor = vbGreen

     'For X = 0 To 1
          Y = CDbl(Buffer)
          .PlotY Y, 0
           'Y = Cos(X / 3.15) * Rnd + 1
          '.PlotXY X, Y, 1
          .SetRange 0, data_points, Ymin, Ymax
      'Next X

End With




counter3 = counter3 + 1

List2.AddItem time
Text25.Text = time
If counter3 = data_points Then
Timer3.Enabled = False
Command37.Enabled = True
Command38.Enabled = False
List6.AddItem ("Done!")
With afg
.WriteString "Output off"
End With
With multi
.WriteString "output off"
End With
End If
End Sub


Private Sub Timer4_Timer()
If Option4.Value = True Then
Combo5.Enabled = False
Combo6.Enabled = False
Combo7.Enabled = False
Combo8.Enabled = False
Combo9.Enabled = False
Else
Combo5.Enabled = True
Combo6.Enabled = True
Combo7.Enabled = True
Combo8.Enabled = True
Combo9.Enabled = True
End If
If Option3.Value = True Then
Combo10.Enabled = False
Combo11.Enabled = False
Combo12.Enabled = False
Combo13.Enabled = False
Combo14.Enabled = False
Else
Combo10.Enabled = True
Combo11.Enabled = True
Combo12.Enabled = True
Combo13.Enabled = True
Combo14.Enabled = True
End If

End Sub


Private Sub Timer6_Timer()

Text12.Text = start_v2
Text16.Text = step2
Text23.Text = start_v2
'offset = Val(off) 'step
'value = CStr(offset)
'Text17.Text = off
'List1.AddItem start_v2
start_v2 = start_v2 + step2
time = time + interval

With multi
.WriteString ":SOUR:VOLT:LEV " & CStr(start_v2) & ""
List1.AddItem start_v2
.WriteString ":READ?"
current = .ReadString
'List3.AddItem current
Buffer = Abs(CDbl(Val(Replace((current), ",", "."))))
List3.AddItem Buffer
Text24.Text = current
'buffer = Val(current)
'buffer = (Replace(Val(CDbl(current), ",", ".")))
Text22.Text = Buffer
Ymax = (Buffer * 3) + Buffer
End With

With NTGraph1
 
     
     .PlotAreaColor = vbBlack
    ' .FrameStyle = Frame
     .Caption = ""
     .XLabel = "Voltage"
     .YLabel = "Current"

     '.ClearGraph 'delete all elements and create a new one
     .ElementLineColor = RGB(255, 255, 0)
     .AddElement  ' Add second elements
     .ElementLineColor = vbGreen

     'For X = 0 To 1
          Y = CDbl(Buffer)
          .PlotY Y, 0
           'Y = Cos(X / 3.15) * Rnd + 1
          '.PlotXY X, Y, 1
          .SetRange 0, data_points, Ymin, Ymax
      'Next X

End With




counter3 = counter3 + 1

List2.AddItem time
Text25.Text = time
If counter3 = data_points Then
Command37.Enabled = True
Timer6.Enabled = False
Command38.Enabled = False
List6.AddItem ("Done!")
With afg
.WriteString "Output off"
End With
With multi
.WriteString "output off"
End With
End If

End Sub
