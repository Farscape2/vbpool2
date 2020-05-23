VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDeelnems 
   BackColor       =   &H00008000&
   ClientHeight    =   10695
   ClientLeft      =   3555
   ClientTop       =   2475
   ClientWidth     =   12720
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10695
   ScaleWidth      =   12720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc dtaVoorspGroepStnd 
      Height          =   330
      Left            =   105
      Top             =   9870
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Groepstand"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmbZetFinals 
      Caption         =   "Automatisch finaleronde invullen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   -15
      TabIndex        =   74
      Top             =   3915
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.Frame frmFin8 
      BackColor       =   &H00008000&
      Caption         =   "Achtste finalisten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1395
      Left            =   15
      TabIndex        =   67
      Top             =   4125
      Width           =   5670
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   15
         Left            =   4440
         TabIndex        =   135
         Top             =   990
         Width           =   270
      End
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   14
         Left            =   4440
         TabIndex        =   134
         Top             =   750
         Width           =   270
      End
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   13
         Left            =   3120
         TabIndex        =   133
         Top             =   990
         Width           =   270
      End
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   12
         Left            =   3120
         TabIndex        =   132
         Top             =   750
         Width           =   270
      End
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   11
         Left            =   1680
         TabIndex        =   131
         Top             =   990
         Width           =   270
      End
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   10
         Left            =   1680
         TabIndex        =   130
         Top             =   750
         Width           =   270
      End
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   9
         Left            =   360
         TabIndex        =   129
         Top             =   990
         Width           =   270
      End
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   8
         Left            =   360
         TabIndex        =   128
         Top             =   750
         Width           =   270
      End
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   7
         Left            =   4440
         TabIndex        =   127
         Top             =   480
         Width           =   270
      End
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   6
         Left            =   4440
         TabIndex        =   126
         Top             =   240
         Width           =   270
      End
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   5
         Left            =   3120
         TabIndex        =   125
         Top             =   480
         Width           =   270
      End
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   4
         Left            =   3120
         TabIndex        =   124
         Top             =   240
         Width           =   270
      End
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   3
         Left            =   1680
         TabIndex        =   123
         Top             =   480
         Width           =   270
      End
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   2
         Left            =   1680
         TabIndex        =   122
         Top             =   240
         Width           =   270
      End
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   1
         Left            =   360
         TabIndex        =   121
         Top             =   480
         Width           =   270
      End
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   0
         Left            =   360
         TabIndex        =   120
         Top             =   240
         Width           =   270
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000004&
         Height          =   525
         Index           =   2
         Left            =   2880
         Top             =   195
         Width           =   1320
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   15
         Left            =   4755
         TabIndex        =   98
         Top             =   990
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   14
         Left            =   4755
         TabIndex        =   97
         Top             =   750
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   13
         Left            =   3375
         TabIndex        =   96
         Top             =   1005
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   12
         Left            =   3375
         TabIndex        =   95
         Top             =   750
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   1995
         TabIndex        =   94
         Top             =   1005
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   1995
         TabIndex        =   93
         Top             =   750
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   630
         TabIndex        =   92
         Top             =   990
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   630
         TabIndex        =   91
         Top             =   750
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   4755
         TabIndex        =   90
         Top             =   450
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   4755
         TabIndex        =   89
         Top             =   210
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   3375
         TabIndex        =   88
         Top             =   465
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   3375
         TabIndex        =   87
         Top             =   225
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   1995
         TabIndex        =   86
         Top             =   450
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   1995
         TabIndex        =   85
         Top             =   210
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   630
         TabIndex        =   84
         Top             =   465
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   630
         TabIndex        =   9
         Top             =   225
         Width           =   750
      End
      Begin VB.Label lblWedNum 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Index           =   7
         Left            =   4275
         TabIndex        =   83
         Top             =   915
         Width           =   135
      End
      Begin VB.Label lblWedNum 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Index           =   6
         Left            =   2895
         TabIndex        =   82
         Top             =   915
         Width           =   135
      End
      Begin VB.Label lblWedNum 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Index           =   5
         Left            =   1515
         TabIndex        =   81
         Top             =   915
         Width           =   135
      End
      Begin VB.Label lblWedNum 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Index           =   4
         Left            =   135
         TabIndex        =   80
         Top             =   915
         Width           =   135
      End
      Begin VB.Label lblWedNum 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Index           =   3
         Left            =   4275
         TabIndex        =   79
         Top             =   360
         Width           =   135
      End
      Begin VB.Label lblWedNum 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Index           =   2
         Left            =   2895
         TabIndex        =   78
         Top             =   360
         Width           =   135
      End
      Begin VB.Label lblWedNum 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Index           =   1
         Left            =   1515
         TabIndex        =   77
         Top             =   360
         Width           =   135
      End
      Begin VB.Label lblWedNum 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Index           =   0
         Left            =   135
         TabIndex        =   76
         Top             =   360
         Width           =   135
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000004&
         Height          =   525
         Index           =   0
         Left            =   105
         Top             =   195
         Width           =   1320
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000004&
         Height          =   525
         Index           =   1
         Left            =   1485
         Top             =   195
         Width           =   1320
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000004&
         Height          =   525
         Index           =   3
         Left            =   4260
         Top             =   195
         Width           =   1320
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000004&
         Height          =   525
         Index           =   4
         Left            =   105
         Top             =   735
         Width           =   1320
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000004&
         Height          =   525
         Index           =   5
         Left            =   1485
         Top             =   735
         Width           =   1320
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000004&
         Height          =   525
         Index           =   6
         Left            =   2880
         Top             =   735
         Width           =   1320
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000004&
         Height          =   525
         Index           =   7
         Left            =   4260
         Top             =   735
         Width           =   1320
      End
   End
   Begin VB.Frame frmFin4 
      BackColor       =   &H00008000&
      Caption         =   "Kwart finalisten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   870
      Left            =   15
      TabIndex        =   62
      Top             =   5565
      Width           =   5655
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   23
         Left            =   4440
         TabIndex        =   143
         Top             =   480
         Width           =   270
      End
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   22
         Left            =   4440
         TabIndex        =   142
         Top             =   240
         Width           =   270
      End
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   21
         Left            =   3120
         TabIndex        =   141
         Top             =   480
         Width           =   270
      End
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   20
         Left            =   3120
         TabIndex        =   140
         Top             =   240
         Width           =   270
      End
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   19
         Left            =   1680
         TabIndex        =   139
         Top             =   480
         Width           =   270
      End
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   18
         Left            =   1680
         TabIndex        =   138
         Top             =   240
         Width           =   270
      End
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   17
         Left            =   360
         TabIndex        =   137
         Top             =   480
         Width           =   270
      End
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   16
         Left            =   360
         TabIndex        =   136
         Top             =   240
         Width           =   270
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "kw"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   23
         Left            =   4770
         TabIndex        =   105
         Top             =   465
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "kw"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   22
         Left            =   4770
         TabIndex        =   104
         Top             =   225
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "kw"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   21
         Left            =   3390
         TabIndex        =   103
         Top             =   465
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "kw"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   20
         Left            =   3390
         TabIndex        =   102
         Top             =   225
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "kw"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   19
         Left            =   2010
         TabIndex        =   101
         Top             =   465
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "kw"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   18
         Left            =   2010
         TabIndex        =   100
         Top             =   225
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "kw"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   17
         Left            =   645
         TabIndex        =   99
         Top             =   465
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "kw"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   16
         Left            =   645
         TabIndex        =   10
         Top             =   225
         Width           =   750
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000004&
         Height          =   600
         Index           =   11
         Left            =   120
         Top             =   180
         Width           =   1320
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000004&
         Height          =   600
         Index           =   10
         Left            =   1500
         Top             =   180
         Width           =   1320
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000004&
         Height          =   600
         Index           =   9
         Left            =   2895
         Top             =   180
         Width           =   1320
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000004&
         Height          =   600
         Index           =   8
         Left            =   4275
         Top             =   180
         Width           =   1320
      End
      Begin VB.Label lblWedNum 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Index           =   11
         Left            =   4290
         TabIndex        =   66
         Top             =   390
         Width           =   135
      End
      Begin VB.Label lblWedNum 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Index           =   10
         Left            =   2910
         TabIndex        =   65
         Top             =   375
         Width           =   135
      End
      Begin VB.Label lblWedNum 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Index           =   9
         Left            =   1515
         TabIndex        =   64
         Top             =   360
         Width           =   135
      End
      Begin VB.Label lblWedNum 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Index           =   8
         Left            =   195
         TabIndex        =   63
         Top             =   360
         Width           =   135
      End
   End
   Begin VB.Frame frmFin2 
      BackColor       =   &H00008000&
      Caption         =   "Halve finalisten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   59
      Top             =   6450
      Width           =   3135
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   27
         Left            =   1800
         TabIndex        =   147
         Top             =   540
         Width           =   270
      End
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   26
         Left            =   1800
         TabIndex        =   146
         Top             =   285
         Width           =   270
      End
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   25
         Left            =   240
         TabIndex        =   145
         Top             =   540
         Width           =   270
      End
      Begin VB.Label lblFin 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   24
         Left            =   240
         TabIndex        =   144
         Top             =   285
         Width           =   270
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000005&
         Height          =   645
         Index           =   13
         Left            =   1575
         Top             =   210
         Width           =   1410
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000005&
         Height          =   645
         Index           =   12
         Left            =   45
         Top             =   210
         Width           =   1410
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "kw"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   27
         Left            =   2160
         TabIndex        =   108
         Top             =   540
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "kw"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   26
         Left            =   2160
         TabIndex        =   107
         Top             =   285
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "kw"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   25
         Left            =   645
         TabIndex        =   106
         Top             =   540
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "kw"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   24
         Left            =   645
         TabIndex        =   11
         Top             =   285
         Width           =   750
      End
      Begin VB.Label lblWedNum 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Index           =   13
         Left            =   1605
         TabIndex        =   61
         Top             =   435
         Width           =   135
      End
      Begin VB.Label lblWedNum 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Index           =   12
         Left            =   90
         TabIndex        =   60
         Top             =   435
         Width           =   135
      End
   End
   Begin VB.Frame frm3ePl 
      BackColor       =   &H00008000&
      Caption         =   "Derde plaats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   990
      Left            =   3315
      TabIndex        =   57
      Top             =   6450
      Width           =   1140
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "kw"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   29
         Left            =   285
         TabIndex        =   109
         Top             =   570
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "kw"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   28
         Left            =   285
         TabIndex        =   12
         Top             =   300
         Width           =   750
      End
      Begin VB.Label lblWedNum 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Index           =   14
         Left            =   90
         TabIndex        =   58
         Top             =   480
         Width           =   135
      End
   End
   Begin VB.Frame frmFin 
      BackColor       =   &H00008000&
      Caption         =   "Finalisten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   990
      Left            =   4515
      TabIndex        =   55
      Top             =   6450
      Width           =   1140
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "kw"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   31
         Left            =   285
         TabIndex        =   110
         Top             =   570
         Width           =   750
      End
      Begin VB.Label kwFin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "kw"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   30
         Left            =   285
         TabIndex        =   13
         Top             =   315
         Width           =   750
      End
      Begin VB.Label lblWedNum 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Index           =   15
         Left            =   105
         TabIndex        =   56
         Top             =   450
         Width           =   135
      End
   End
   Begin VB.Frame frmGrp 
      BackColor       =   &H00008000&
      Caption         =   "Groepstand"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1890
      Left            =   0
      TabIndex        =   52
      Top             =   2010
      Width           =   2655
      Begin MSAdodcLib.Adodc dtaGrpTeams 
         Height          =   330
         Left            =   840
         Top             =   0
         Visible         =   0   'False
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "GrpTeams"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Frame frmGroep 
         BackColor       =   &H00008000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1560
         Index           =   0
         Left            =   840
         TabIndex        =   53
         Top             =   240
         Width           =   1680
         Begin MSDataListLib.DataCombo groep 
            Bindings        =   "frmPoolParticipants.frx":0000
            DataField       =   "pos1"
            DataSource      =   "dtaVoorspGroepStnd"
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Naam"
            BoundColumn     =   "id"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo groep 
            Bindings        =   "frmPoolParticipants.frx":0024
            DataField       =   "pos2"
            DataSource      =   "dtaVoorspGroepStnd"
            Height          =   315
            Index           =   1
            Left            =   240
            TabIndex        =   6
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Naam"
            BoundColumn     =   "id"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo groep 
            Bindings        =   "frmPoolParticipants.frx":003E
            DataField       =   "pos3"
            DataSource      =   "dtaVoorspGroepStnd"
            Height          =   315
            Index           =   2
            Left            =   240
            TabIndex        =   7
            Top             =   840
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Naam"
            BoundColumn     =   "id"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo groep 
            Bindings        =   "frmPoolParticipants.frx":0058
            DataField       =   "pos4"
            DataSource      =   "dtaVoorspGroepStnd"
            Height          =   315
            Index           =   3
            Left            =   240
            TabIndex        =   8
            Top             =   1200
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Naam"
            BoundColumn     =   "id"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   117
            Top             =   150
            Width           =   180
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   1
            Left            =   75
            TabIndex        =   116
            Top             =   510
            Width           =   180
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   2
            Left            =   75
            TabIndex        =   115
            Top             =   870
            Width           =   180
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   3
            Left            =   75
            TabIndex        =   114
            Top             =   1230
            Width           =   180
         End
      End
      Begin VB.ListBox cmbGroepskeus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1620
         Left            =   120
         TabIndex        =   4
         Top             =   195
         Width           =   660
      End
   End
   Begin VB.Frame frmTS 
      BackColor       =   &H00008000&
      Caption         =   "Topscorers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2070
      Left            =   2670
      TabIndex        =   47
      Top             =   2055
      Width           =   3015
      Begin MSDataListLib.DataCombo cmbTS 
         Bindings        =   "frmPoolParticipants.frx":0072
         DataSource      =   "dtaDezeDeelnemer"
         Height          =   315
         Index           =   0
         Left            =   480
         TabIndex        =   19
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "rnaam"
         BoundColumn     =   "id"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSAdodcLib.Adodc dtaSpelers 
         Height          =   330
         Left            =   600
         Top             =   1560
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   60
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Spelers"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CommandButton btnNweSpeler 
         Caption         =   "Nieuw"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1650
         TabIndex        =   72
         Top             =   165
         Width           =   840
      End
      Begin VB.TextBox dpTS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2565
         TabIndex        =   20
         Top             =   450
         Width           =   345
      End
      Begin VB.TextBox dpTS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   2565
         TabIndex        =   24
         Top             =   1140
         Width           =   345
      End
      Begin VB.TextBox dpTS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2565
         TabIndex        =   23
         Top             =   795
         Width           =   345
      End
      Begin MSDataListLib.DataCombo cmbTS 
         Bindings        =   "frmPoolParticipants.frx":008B
         DataSource      =   "dtaDezeDeelnemer"
         Height          =   315
         Index           =   1
         Left            =   480
         TabIndex        =   118
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "rnaam"
         BoundColumn     =   "id"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbTS 
         Bindings        =   "frmPoolParticipants.frx":00A4
         DataSource      =   "dtaDezeDeelnemer"
         Height          =   315
         Index           =   2
         Left            =   480
         TabIndex        =   119
         Top             =   1200
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "rnaam"
         BoundColumn     =   "id"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Speler"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   24
         Left            =   345
         TabIndex        =   54
         Top             =   225
         Width           =   885
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   4
         Left            =   105
         TabIndex        =   51
         Top             =   495
         Width           =   180
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   5
         Left            =   90
         TabIndex        =   50
         Top             =   825
         Width           =   180
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   6
         Left            =   90
         TabIndex        =   49
         Top             =   1170
         Width           =   180
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "dp"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   7
         Left            =   2625
         TabIndex        =   48
         Top             =   210
         Width           =   180
      End
   End
   Begin VB.Frame frmDiv 
      BackColor       =   &H00008000&
      Caption         =   "Diversen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2385
      Left            =   2040
      TabIndex        =   18
      Top             =   7440
      Width           =   3660
      Begin MSAdodcLib.Adodc dtaVoorspAantallen 
         Height          =   330
         Left            =   240
         Top             =   1560
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "VoorspAantallen"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.TextBox txtVoorsp 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3045
         TabIndex        =   21
         Top             =   465
         Width           =   510
      End
      Begin VB.Label lblGrp 
         BackStyle       =   0  'Transparent
         Caption         =   "Voorspelling"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   46
         Top             =   210
         Width           =   1005
      End
      Begin VB.Label lblGrp 
         BackStyle       =   0  'Transparent
         Caption         =   "Aantal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   1
         Left            =   3030
         TabIndex        =   45
         Top             =   210
         Width           =   495
      End
      Begin VB.Label lblVoorsp 
         BackStyle       =   0  'Transparent
         Caption         =   "Voorspelling"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   44
         Top             =   465
         Width           =   2520
      End
   End
   Begin VB.Frame frmEindstand 
      BackColor       =   &H00008000&
      Caption         =   "Eindstand"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2385
      Left            =   15
      TabIndex        =   39
      Top             =   7440
      Width           =   1965
      Begin MSDataListLib.DataCombo eind 
         Bindings        =   "frmPoolParticipants.frx":00BD
         DataField       =   "kampioen"
         DataSource      =   "dtaDezeDeelnemer"
         Height          =   315
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Naam"
         BoundColumn     =   "id"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo eind 
         Bindings        =   "frmPoolParticipants.frx":00D4
         DataField       =   "pltwee"
         DataSource      =   "dtaDezeDeelnemer"
         Height          =   315
         Index           =   1
         Left            =   360
         TabIndex        =   15
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Naam"
         BoundColumn     =   "id"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo eind 
         Bindings        =   "frmPoolParticipants.frx":00EB
         DataField       =   "plDrie"
         DataSource      =   "dtaDezeDeelnemer"
         Height          =   315
         Index           =   2
         Left            =   360
         TabIndex        =   16
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Naam"
         BoundColumn     =   "id"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo eind 
         Bindings        =   "frmPoolParticipants.frx":0102
         DataField       =   "plVier"
         DataSource      =   "dtaDezeDeelnemer"
         Height          =   315
         Index           =   3
         Left            =   360
         TabIndex        =   17
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Naam"
         BoundColumn     =   "id"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblPl 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   0
         Left            =   135
         TabIndex        =   43
         Top             =   330
         Width           =   270
      End
      Begin VB.Label lblPl 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   1
         Left            =   135
         TabIndex        =   42
         Top             =   660
         Width           =   270
      End
      Begin VB.Label lblPl 
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   2
         Left            =   135
         TabIndex        =   41
         Top             =   990
         Width           =   270
      End
      Begin VB.Label lblPl 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   3
         Left            =   135
         TabIndex        =   40
         Top             =   1335
         Width           =   270
      End
   End
   Begin VB.Frame frmDeeln 
      BackColor       =   &H00008000&
      Caption         =   "Deelnemer gegevens"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1455
      Left            =   -15
      TabIndex        =   1
      Top             =   555
      Width           =   5685
      Begin VB.CheckBox chkBetaald 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Betaald:"
         DataField       =   "betaald"
         DataSource      =   "dtaDezeDeelnemer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Left            =   4410
         TabIndex        =   3
         Top             =   285
         Width           =   975
      End
      Begin VB.TextBox txtBijnaam 
         DataField       =   "bijnaam"
         DataSource      =   "dtaDezeDeelnemer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   870
         TabIndex        =   2
         Top             =   615
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Mobiel:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   3840
         TabIndex        =   37
         Top             =   645
         Width           =   585
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         DataField       =   "mob"
         DataSource      =   "dtaDeelnems"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4425
         TabIndex        =   36
         Top             =   615
         Width           =   1200
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         DataField       =   "email"
         DataSource      =   "dtaDeelnems"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   870
         TabIndex        =   34
         Top             =   990
         Width           =   4755
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Tel:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   2205
         TabIndex        =   33
         Top             =   645
         Width           =   360
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         DataField       =   "tel"
         DataSource      =   "dtaDeelnems"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2565
         TabIndex        =   32
         Top             =   615
         Width           =   1275
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   90
         TabIndex        =   31
         Top             =   1020
         Width           =   495
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Naam:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   105
         TabIndex        =   30
         Top             =   285
         Width           =   765
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Echte naam: "
         DataField       =   "naam"
         DataSource      =   "dtaDeelnems"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   870
         TabIndex        =   29
         Top             =   255
         Width           =   2985
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Bijnaam:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   90
         TabIndex        =   28
         Top             =   645
         Width           =   765
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   12660
      TabIndex        =   26
      Top             =   0
      Width           =   12720
      Begin MSAdodcLib.Adodc dtaDeelnems 
         Height          =   375
         Left            =   4920
         Top             =   0
         Visible         =   0   'False
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   60
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Deelnems"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo cmbDeelnems 
         Bindings        =   "frmPoolParticipants.frx":0119
         Height          =   315
         Left            =   1080
         TabIndex        =   113
         Top             =   120
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "bijnaam"
         BoundColumn     =   "deelnemID"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbAdressen 
         Bindings        =   "frmPoolParticipants.frx":0133
         Height          =   315
         Left            =   6960
         TabIndex        =   112
         Top             =   120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "naam"
         BoundColumn     =   "id"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSAdodcLib.Adodc dtaAdressen 
         Height          =   375
         Left            =   10800
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adressen"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CommandButton btnWisDeelnem 
         Caption         =   "Verwijderen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3375
         TabIndex        =   75
         ToolTipText     =   "Verwijder de huidige deelnemer"
         Top             =   120
         Width           =   1410
      End
      Begin VB.CommandButton btnNewAdres 
         Caption         =   "Adressenlijst"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9240
         TabIndex        =   0
         ToolTipText     =   "Klik hier om nieuw adres toe te voegen"
         Top             =   120
         Width           =   1470
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Deelnemer toevoegen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Left            =   5220
         TabIndex        =   38
         Top             =   150
         Width           =   1740
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Deelnemer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   45
         TabIndex        =   27
         Top             =   165
         Width           =   960
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   510
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   12720
      TabIndex        =   73
      Top             =   10185
      Width           =   12720
      Begin MSAdodcLib.Adodc dtaDeelnemWeds 
         Height          =   330
         Left            =   7680
         Top             =   0
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         ConnectMode     =   3
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   60
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "aDeelnemWeds"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc dtaDezeDeelnemer 
         Height          =   375
         Left            =   5880
         Top             =   0
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "DezeDeelnemer"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CommandButton btnClose 
         Caption         =   "Sluiten"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10935
         TabIndex        =   35
         Top             =   45
         Width           =   1395
      End
      Begin MSAdodcLib.Adodc dtaTeams 
         Height          =   375
         Left            =   4080
         Top             =   0
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Teams"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label lblDeelnemAant 
         BackStyle       =   0  'Transparent
         Caption         =   "Aantal deelnemers:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   45
         TabIndex        =   111
         Top             =   45
         Width           =   2520
      End
   End
   Begin VB.Frame frmUitslagen 
      BackColor       =   &H00008000&
      Caption         =   "Voorspellingen uitslagen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   9330
      Left            =   5745
      TabIndex        =   25
      Top             =   555
      Width           =   6615
      Begin MSDataGridLib.DataGrid grdDeelnWeds 
         Bindings        =   "frmPoolParticipants.frx":014D
         Height          =   8415
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   14843
         _Version        =   393216
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
         WrapCellPointer =   -1  'True
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "ksid"
            Caption         =   "ksid"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1043
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "deelnem"
            Caption         =   "deelnem"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1043
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "wednum"
            Caption         =   "nr"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1043
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "weddatum"
            Caption         =   "datum"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1043
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "wedoms"
            Caption         =   "Wedstrijd"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1043
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "r1"
            Caption         =   "Rust"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1043
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "r2"
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1043
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "e1"
            Caption         =   "Eind"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1043
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "e2"
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1043
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "toto"
            Caption         =   "toto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1043
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   2
         BeginProperty Split0 
            MarqueeStyle    =   2
            ScrollBars      =   0
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            RecordSelectors =   0   'False
            Size            =   2
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   345,26
            EndProperty
            BeginProperty Column03 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1500,095
            EndProperty
            BeginProperty Column04 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   2505,26
            EndProperty
            BeginProperty Column05 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   390,047
            EndProperty
            BeginProperty Column06 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   390,047
            EndProperty
            BeginProperty Column07 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   390,047
            EndProperty
            BeginProperty Column08 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   390,047
            EndProperty
            BeginProperty Column09 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   390,047
            EndProperty
         EndProperty
         BeginProperty Split1 
            MarqueeStyle    =   2
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            RecordSelectors =   0   'False
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   345,26
            EndProperty
            BeginProperty Column03 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1500,095
            EndProperty
            BeginProperty Column04 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   2505,26
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               DividerStyle    =   0
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   390,047
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   345,26
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               DividerStyle    =   0
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   390,047
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   345,26
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   345,26
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtTtl 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   8895
         Width           =   480
      End
      Begin VB.TextBox txtTtl 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2175
         Locked          =   -1  'True
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   8895
         Width           =   480
      End
      Begin VB.Label lblTtl 
         BackStyle       =   0  'Transparent
         Caption         =   "Doelpunten"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   1
         Left            =   2925
         TabIndex        =   70
         Top             =   8910
         Width           =   990
      End
      Begin VB.Label lblTtl 
         BackStyle       =   0  'Transparent
         Caption         =   "Aantal gelijke spelen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   0
         Left            =   630
         TabIndex        =   68
         Top             =   8910
         Width           =   1560
      End
   End
End
Attribute VB_Name = "frmDeelnems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim deeln As Long

Dim WithEvents rsGroepStanden As ADODB.Recordset
Attribute rsGroepStanden.VB_VarHelpID = -1



Private Sub btn8FinEdit_Click()
    curDeeln = deeln
    frm8Fin.Show 1
    deelnFinalePlaatsen
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnFinEdit_Click()
    curDeeln = deeln
    frm1Fin.Show 1
    deelnFinalePlaatsen
End Sub

Private Sub btnKwFinEdit_Click()
    curDeeln = deeln
    frm4Fin.Show 1
    deelnFinalePlaatsen
End Sub

Private Sub btnHalveFinEdit_Click()
    curDeeln = deeln
    frm2Fin.Show 1
    deelnFinalePlaatsen
End Sub

Private Sub btnNewAdres_Click()
    frmAdressen.Show 1
    Me.dtaAdressen.Refresh
End Sub

Private Sub btnNweSpeler_Click()
    frmTeams.Show 1
    Me.dtaSpelers.Refresh
End Sub

Private Sub btnWisDeelnem_Click()
Dim delStr As String
Dim tbl As String
Dim selStr As String
Dim msg As String
msg = "Weet je zeker dat deelnemer met bijnaam: " & Me.cmbDeelnems.Text
msg = msg & vbNewLine & "verwijderd moet worden?"
    If MsgBox(msg, vbYesNo, "Deelnemer verwijderen") = vbYes Then
        delStr = "Delete from "
        selStr = " where deelnem = " & deeln
        cn.Execute delStr & " pooldeelnems where deelnemid= " & deeln
        tbl = "voorspelling_aantallen"
        cn.Execute delStr & tbl & selStr
        tbl = "voorspelling_finales"
        cn.Execute delStr & tbl & selStr
        tbl = "voorspelling_groepstand"
        cn.Execute delStr & tbl & selStr
        tbl = "voorspelling_ts"
        cn.Execute delStr & tbl & selStr
        tbl = "voorspelling_uitsl"
        cn.Execute delStr & tbl & selStr
        Me.dtaDeelnems.Refresh
        Me.dtaDezeDeelnemer.Refresh
        If Me.dtaDezeDeelnemer.Recordset.RecordCount > 0 Then
            deeln = Me.dtaDezeDeelnemer.Recordset!deelnemID
            Me.cmbDeelnems.BoundText = deeln
            cmbDeelnems_Click 2
        End If
    End If
End Sub

Private Sub UpdateDeelnAant()
'maak de voorspelling_aantallen tabel aan voor deze deelnemer
Dim rs As New ADODB.Recordset
Dim poolVoorsp As New ADODB.Recordset
Dim selStr As String
Dim opdr As String
    opdr = "Delete "
    selStr = "from voorspelling_aantallen WHERE deelnem = " & deeln
    cn.Execute opdr & selStr
    selStr = "Select * " & selStr
    rs.Open selStr, cn, adOpenDynamic, adLockOptimistic
    selStr = "Select * from pntToek WHERE poolid=" & poolID
    selStr = selStr & " AND voorspeltype IN (Select id from voorspeltypes WHERE cat = 1)"
    poolVoorsp.Open selStr, cn, adOpenStatic, adLockReadOnly
    Do While Not poolVoorsp.EOF
        rs.AddNew
            rs!deelnem = deeln
            rs!voorspel_type = poolVoorsp!voorspelType
            rs!Aantal = 0
        rs.Update
        poolVoorsp.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Private Sub updateDeelnFin()
Dim rs As New ADODB.Recordset
Dim rsWeds As New ADODB.Recordset
Dim selStr As String
Dim opdr As String
    selStr = "Select * from toernschema where ksID = " & kampID
    selStr = selStr & " AND wedtype > 1"
    rsWeds.Open selStr, cn, adOpenStatic, adLockReadOnly
    opdr = "Delete "
    selStr = "from voorspelling_finales WHERE deelnem = " & deeln
    cn.Execute opdr & selStr
    selStr = "Select * " & selStr
    rs.Open selStr, cn, adOpenDynamic, adLockOptimistic
    Do While Not rsWeds.EOF
        rs.AddNew
            rs!deelnem = deeln
            rs!wed = rsWeds!wedNum
            rs!t1 = 0
            rs!t2 = 0
        rs.Update
        rsWeds.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    rsWeds.Close
    Set rsWeds = Nothing
End Sub

Private Sub updateDeelnGrpStnd()
Dim rs As New ADODB.Recordset
Dim selStr As String
Dim opdr As String
Dim grp As Integer
Dim i As Integer
    opdr = "Delete "
    selStr = "from voorspelling_groepstand WHERE deelnem = " & deeln
    cn.Execute opdr & selStr
    selStr = "Select * " & selStr
    rs.Open selStr, cn, adOpenDynamic, adLockOptimistic
    grp = getKampInfo("groepen")
    For i = 1 To grp
        rs.AddNew
            rs!deelnem = deeln
            rs!groep = Chr(64 + i)
            rs!pos1 = 0
            rs!pos2 = 0
            rs!pos3 = 0
            rs!pos4 = 0
        rs.Update
    Next
    rs.Close
    Set rs = Nothing
End Sub

Private Sub UpdateDeelnTS()
Dim rs As New ADODB.Recordset
Dim poolVoorsp As New ADODB.Recordset
Dim selStr As String
Dim opdr As String
Dim i As Integer
    opdr = "Delete "
    selStr = "from voorspelling_ts WHERE deelnem = " & deeln
    cn.Execute opdr & selStr
    selStr = "Select * " & selStr
    rs.Open selStr, cn, adOpenDynamic, adLockOptimistic
    selStr = "Select * from pntToek WHERE poolid=" & poolID
    selStr = selStr & " AND voorspeltype IN (Select id from voorspeltypes WHERE cat = 5)"
    poolVoorsp.Open selStr, cn, adOpenStatic, adLockReadOnly
    Do While Not poolVoorsp.EOF
        i = i + 1
        rs.AddNew
            rs!deelnem = deeln
            rs!tsnr = i
            rs!ts = 0
            rs!dp = 0
        rs.Update
        poolVoorsp.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Private Sub updateDeelnUitsl()
Dim rs As New ADODB.Recordset
Dim rsWeds As New ADODB.Recordset
Dim selStr As String
Dim opdr As String
    selStr = "Select * from qryweds where ksID = " & kampID
    rsWeds.Open selStr, cn, adOpenDynamic, adLockReadOnly
    opdr = "Delete "
    selStr = "from voorspelling_uitsl WHERE deelnem = " & deeln
    cn.Execute opdr & selStr
    selStr = "Select * " & selStr
    rs.Open selStr, cn, adOpenStatic, adLockOptimistic
    Do While Not rsWeds.EOF
        rs.AddNew
            rs!deelnem = deeln
            rs!wedNum = rsWeds!wedNum
            rs!r1 = 0
            rs!r2 = 0
            rs!e1 = 0
            rs!e2 = 0
            rs!toto = 3
        rs.Update
        rsWeds.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    rsWeds.Close
    Set rsWeds = Nothing
End Sub

Private Sub cmbAdressen_Click(Area As Integer)
Dim rsDeeln As New ADODB.Recordset
Dim selStr As String
Dim msg As String
Dim deelNaam As String
Dim nwDeeln As Boolean
    If Area <> 2 Then Exit Sub
    
    selStr = "Select * from pooldeelnems WHERE poolID = " & poolID
    selStr = selStr & " AND adrID =" & Me.cmbAdressen.BoundText
    rsDeeln.Open selStr, cn, adOpenStatic, adLockOptimistic
    If rsDeeln.RecordCount <= 0 Then
        msg = "Wil je " & Me.cmbAdressen.Text & " toevoegen?"
    Else
        msg = "Wil je nog een pool voor " & Me.cmbAdressen.Text & " toevoegen?"
    End If
    nwDeeln = MsgBox(msg, vbYesNo + vbQuestion, "Deelnemer toevoegen") = vbYes
    
    If nwDeeln Then
        'voeg de tabellen toe
        deelNaam = getDefaultBijnaam(Me.cmbAdressen.BoundText)
        With rsDeeln
            .AddNew
            !poolID = poolID
            !adrID = Me.cmbAdressen.BoundText
            !betaald = False
            If .RecordCount = 0 Then
                !bijnaam = deelNaam
            Else
                !bijnaam = deelNaam & Format(.RecordCount)
            End If
            .Update
            deeln = !deelnemID
        End With
        Me.lblDeelnemAant.Caption = "Deelnemer: " & rsDeeln.AbsolutePosition & "/" & rsDeeln.RecordCount
        Me.dtaDezeDeelnemer.Refresh
        Me.dtaDeelnems.Refresh
        'werk de tabellen in de database bij
        UpdateDeelnAant
        updateDeelnFin
        updateDeelnGrpStnd
        UpdateDeelnTS
        updateDeelnUitsl
        'en laat deelnemer zien
        Me.cmbDeelnems.BoundText = deeln
        cmbDeelnems_Click 2
    End If
    rsDeeln.Close
    Set rsDeeln = Nothing
End Sub

Private Sub cmbDeelnems_Click(Area As Integer)
If Area = 2 And Me.cmbDeelnems.Text > "" Then

    deeln = Me.cmbDeelnems.BoundText
    Me.dtaDezeDeelnemer.Refresh
    Me.dtaDezeDeelnemer.Recordset.MoveFirst
    Me.dtaDezeDeelnemer.Recordset.Find "deelnemID = " & deeln
    Me.dtaDeelnems.Recordset.MoveFirst
    Me.dtaDeelnems.Recordset.Find "deelnemID = " & deeln
    Me.cmbGroepskeus.Text = "A"
    deelnFinalePlaatsen
    deelnAantallen
    deelnTS

    deelnWeds
    Me.lblDeelnemAant.Caption = "Deelnemer: " & Me.dtaDeelnems.Recordset.AbsolutePosition & "/" & Me.dtaDeelnems.Recordset.RecordCount
End If
End Sub

Private Sub deelnWeds()
Dim sqlstr As String
Dim recs As Long
Dim rs As New ADODB.Recordset
    cn.Execute "Delete from tmpDeelnUitsl"
    sqlstr = "INSERT INTO tmpDeelnUitsl ( ksid, wedNum, wedDatum, wedOms, deelnem, r1, r2, e1, e2, toto ) "
    sqlstr = sqlstr & "SELECT qryWeds.ksid, qryWeds.wedNum, Format(qryWeds.datum,'ddd d mmm') & ' om ' & "
    sqlstr = sqlstr & "Format(qryweds.tijd,'Short Time') AS wedDatum, code1 & ': ' & qryweds.naam1 & ' - ' & code2 & ': ' & qryWeds.Naam2 "
    sqlstr = sqlstr & "AS wedOms, voorspelling_uitsl.deelnem, voorspelling_uitsl.r1, voorspelling_uitsl.r2, "
    sqlstr = sqlstr & "voorspelling_uitsl.e1, voorspelling_uitsl.e2, voorspelling_uitsl.toto "
    sqlstr = sqlstr & "FROM voorspelling_uitsl LEFT JOIN qryWeds ON voorspelling_uitsl.wednum = qryWeds.wedNum "
    sqlstr = sqlstr & "WHERE voorspelling_uitsl.deelnem = " & deeln
    sqlstr = sqlstr & " AND ksID = " & kampID
    sqlstr = sqlstr & " ORDER BY qryWeds.ksid, qryweds.datum, qryweds.tijd,qryWeds.wedNum"
    cn.Execute sqlstr, recs
    'cn.CommitTrans
    rs.Open "Select * from tmpDeelnUitsl", cn, adOpenDynamic, adLockOptimistic
    DoEvents
    Set Me.dtaDeelnemWeds.Recordset = rs
    Me.grdDeelnWeds.Refresh
    BerekenTtl
End Sub

Sub deelnGroepStand()
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim selStr As String
    selStr = "Select * from voorspelling_groepstand Where deelnem = " & deeln
    selStr = selStr & " AND groep = '" & Me.cmbGroepskeus.Text & "'"
    rs.Open selStr, cn, adOpenDynamic, adLockOptimistic
    If rs.RecordCount = 1 Then
        For i = 0 To 3
            If rs("pos" & Format(i + 1, "0")) > 0 Then
                Me.groep(i).BoundText = rs("pos" & Format(i + 1, "0"))
            End If
        Next
    End If
    rs.Close
    Set rs = Nothing
End Sub

Sub deelnFinalePlaatsen()
Dim rs As New ADODB.Recordset
Dim sqlstr As String
Dim i As Integer
    sqlstr = "Select * from voorspelling_finales"
    sqlstr = sqlstr & " WHERE deelnem = " & deeln
    sqlstr = sqlstr & " ORDER BY wed"
    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        rs.MoveLast
        rs.MoveFirst
        'achtste finales
        If getKampInfo("groepen") > 4 Then
            For i = 0 To 15 Step 2
                If Not rs.EOF Then
                    Me.kwFin(i).Caption = GetTeam(rs!t1, True)
                    Me.kwFin(i + 1).Caption = GetTeam(rs!t2, True)
                    rs.MoveNext
                End If
            Next
        End If
        'kwartfinales
        For i = 16 To 23 Step 2
            If Not rs.EOF Then
                Me.kwFin(i).Caption = GetTeam(rs!t1, True)
                Me.kwFin(i + 1).Caption = GetTeam(rs!t2, True)
                rs.MoveNext
            End If
        Next
        'halve finales
        For i = 24 To 27 Step 2
            If Not rs.EOF Then
                Me.kwFin(i).Caption = GetTeam(rs!t1, True)
                Me.kwFin(i + 1).Caption = GetTeam(rs!t2, True)
                rs.MoveNext
            End If
        Next
        'kleine finale
        If getKampInfo("derdeplaats") Then
            i = 28
            If Not rs.EOF Then
                Me.kwFin(i).Caption = GetTeam(rs!t1, True)
                Me.kwFin(i + 1).Caption = GetTeam(rs!t2, True)
                rs.MoveNext
            End If
            i = 30
        Else
            i = 30
        End If
        'finale
        If Not rs.EOF Then
                Me.kwFin(i).Caption = GetTeam(rs!t1, True)
                Me.kwFin(i + 1).Caption = GetTeam(rs!t2, True)
            rs.MoveNext
        End If
    End If
    rs.Close
End Sub
Sub deelnTS()
    Dim rs As New ADODB.Recordset
    Dim sqlstr As String
    sqlstr = "Select * from voorspelling_ts Where deelnem = " & deeln
    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            Me.cmbTS(rs!tsnr - 1).BoundText = rs!ts
            Me.dpTS(rs!tsnr - 1) = rs!dp
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub
Sub deelnAantallen()
Dim rs As New ADODB.Recordset
Dim sqlstr As String
Dim i As Integer
    sqlstr = "Select * from voorspelling_aantallen"
    sqlstr = sqlstr & " WHERE deelnem = " & deeln
    sqlstr = sqlstr & " ORDER BY voorspel_type"
    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    i = 0
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
          If i <= Me.txtVoorsp.UBound Then
            Me.txtVoorsp(i).Text = rs!Aantal
          End If
          i = i + 1
          rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub


Private Sub cmbGroepsKeus_Click()
Dim i As Integer
Dim sqlstr As String
Dim rs As ADODB.Recordset

    sqlstr = "Select id, naam, afkort from qryGroepTeams "
    sqlstr = sqlstr & " WHERE ksid = " & kampID
    sqlstr = sqlstr & " AND groep = '" & Me.cmbGroepskeus.Text & "'"
    sqlstr = sqlstr & " ORDER BY plaatsing"
    Set rs = New ADODB.Recordset
    rs.Open sqlstr, cn
    Set Me.dtaGrpTeams.Recordset = rs
    Me.dtaGrpTeams.Refresh
    rs.Close
    sqlstr = "Select * from voorspelling_groepstand"
    sqlstr = sqlstr & " WHERE deelnem = " & deeln
    sqlstr = sqlstr & " AND groep = '" & Me.cmbGroepskeus.Text & "'"
    'rs.Open sqlstr, cn
    Set rsGroepStanden = New ADODB.Recordset
    rsGroepStanden.Open sqlstr, cn
    Set Me.dtaVoorspGroepStnd.Recordset = rsGroepStanden
    
'    For i = 0 To 3
'        Set Me.groep(i).DataSource = rsGroepStanden
'        Set Me.groep(i).RowSource = rs
'    Next
'    deelnGroepStand
   ' rs.Close
    Set rs = Nothing
End Sub

Private Sub cmbTS_Click(Index As Integer, Area As Integer)
    If Area = 2 Then
        updateTS Index
    End If
End Sub

Sub updateTS(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim sqlstr As String
    sqlstr = "Select * from voorspelling_ts Where deelnem = " & deeln
    rs.Open sqlstr, cn, adOpenStatic, adLockOptimistic
    If rs.RecordCount > 0 Then
        rs.Find "tsnr=" & Index + 1
        If Not rs.EOF Then
          rs!ts = Me.cmbTS(Index).BoundText
          rs!dp = Me.dpTS(Index)
          rs.Update
        End If
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub cmbZetFinals_Click()
Dim selStr As String
Dim Index As Integer
Dim grp As Integer
    'update de groepsvoorspelling bij deze deelnemer
    For grp = 1 To getKampInfo("groepen")
        For Index = 0 To 1
            ZetStandaardFinales Chr(grp + 64), Index + 1
        Next
    Next
End Sub


Private Sub dpTS_GotFocus(Index As Integer)
    Me.dpTS(Index).SelStart = 0
    Me.dpTS(Index).SelLength = Len(Me.txtVoorsp(Index).Text)
End Sub

Private Sub dpTS_LostFocus(Index As Integer)
    If nz(Me.dpTS(Index), 0) > 0 Then
        updateTS Index
    End If
End Sub

Private Sub XXXdtaDeelnems_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Me.lblDeelnemAant.Caption = "Deelnemer: " & dtaDeelnems.Recordset.AbsolutePosition & "/" & dtaDeelnems.Recordset.RecordCount
End Sub

Private Sub dtaGrpTeams_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
    MsgBox ErrorNumber & ":" & Description
End Sub

Private Sub dtaVoorspGroepstand_RecordChangeComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    'Me.dtaVoorspGroepstand.Recordset
End Sub

Sub savEindstand(i As Integer)
Dim sqlstr As String
    sqlstr = "UPDATE pooldeelnems SET "
    Select Case i
    Case 0
        sqlstr = sqlstr & " kampioen = "
    Case 1
        sqlstr = sqlstr & " pltwee = "
    Case 2
        sqlstr = sqlstr & " pldrie = "
    Case 3
        sqlstr = sqlstr & " plvier = "
    End Select
    sqlstr = sqlstr & Me.eind(i).BoundText
    sqlstr = sqlstr & " WHERE deelnemid = " & deeln
    cn.Execute sqlstr
End Sub

Private Sub eind_Click(Index As Integer, Area As Integer)
If Area = 2 Then
    savEindstand Index
End If
End Sub

Private Sub Form_Activate()
If Not Me.frmFin8.Visible Then
    Me.frmFin4.Top = Me.frmFin8.Top
    Me.frmFin2.Top = Me.frmFin4.Top + Me.frmFin4.Height + 100
    Me.frmFin.Top = Me.frmFin2.Top
    Me.frmEindstand.Top = Me.frmFin.Top + Me.frmFin.Height + 100
    Me.frmDiv.Top = Me.frmEindstand.Top
    Me.grdDeelnWeds.Height = Me.grdDeelnWeds.RowHeight * 35
    Me.frmUitslagen.Height = Me.grdDeelnWeds.Height + 800
    Me.txtTtl(0).Top = Me.grdDeelnWeds.Top + Me.grdDeelnWeds.Height + 50
    Me.txtTtl(1).Top = Me.grdDeelnWeds.Top + Me.grdDeelnWeds.Height + 50
    Me.lblTtl(0).Top = Me.grdDeelnWeds.Top + Me.grdDeelnWeds.Height + 50
    Me.lblTtl(1).Top = Me.grdDeelnWeds.Top + Me.grdDeelnWeds.Height + 50
    'Me.btnClose.Top = Me.frmUitslagen.Top + Me.frmUitslagen.Height + 50
End If
Me.Height = Me.frmUitslagen.Top + Me.frmUitslagen.Height + 720 + Me.Picture2.Height
Me.width = 12660
'CenterForm Me
'If Me.Height < 10980 Then Me.Height = 10980
End Sub

Private Sub Form_Load()
Dim dta As Control
Dim grp As Integer
Dim i As Integer
Dim j As Integer
Dim sqlstr As String
Dim wdnr As Integer
Dim lastwednr As Integer
Dim strtWedNumLabel As Integer
On Error GoTo formErr
'    Debug.Print "2"
'   chkDBnaam Me
'    sqlstr = "Select * from qryAdressen ORDER BY anaam"
     Me.Height = 10950
    'Me.Width = 12525
    cn.Execute "Delete from tmpDeelnUitsl"
    Screen.MousePointer = vbHourglass
    grp = getKampInfo("groepen")
    If grp = 0 Then
        MsgBox "BIIIIG ERRRRORRR"
        End
    End If
    
    Me.dtaAdressen.ConnectionString = cn.ConnectionString
    Me.dtaAdressen.CommandType = adCmdText
    Me.dtaAdressen.RecordSource = "select * from qryAdressen"
    Me.dtaAdressen.Refresh
    Me.cmbAdressen.ReFill
    
    sqlstr = "Select * from qrydeelnems Where poolid = " & poolID
    sqlstr = sqlstr & " ORDER BY bijnaam"
    Me.dtaDeelnems.ConnectionString = cn.ConnectionString
    Me.dtaDeelnems.RecordSource = sqlstr
    Me.dtaDeelnems.Refresh
    Me.cmbDeelnems.ReFill
    
    Me.dtaDeelnemWeds.ConnectionString = cn.ConnectionString
    Me.dtaDeelnemWeds.RecordSource = "Select * from tmpDeelnUitsl"
    
    Me.dtaDezeDeelnemer.ConnectionString = cn.ConnectionString
    Me.dtaDezeDeelnemer.RecordSource = "Select * from pooldeelnems Where poolid = " & poolID
    Me.dtaDezeDeelnemer.Refresh
    
    sqlstr = "Select * from qryGroepTeams"
    Me.dtaGrpTeams.ConnectionString = cn.ConnectionString
    Me.dtaGrpTeams.CommandType = adCmdText
    Me.dtaGrpTeams.RecordSource = sqlstr
    Me.dtaGrpTeams.Refresh
    For i = 1 To grp
        Me.cmbGroepskeus.AddItem Chr(i + 64)
    Next
    Me.frmFin8.Visible = grp > 4
    
    Me.dtaSpelers.ConnectionString = cn.ConnectionString
    Me.dtaSpelers.RecordSource = "Select * from qryVoetballers WHERE kampID = " & kampID
    Me.dtaSpelers.Refresh
    
    sqlstr = "SELECT groepsindeling.ksID, groepsindeling.groep, teams.id, TeamNamen.Naam"
    sqlstr = sqlstr & " FROM (groepsindeling INNER JOIN teams ON groepsindeling.team = teams.id) "
    sqlstr = sqlstr & " INNER JOIN TeamNamen ON teams.team = TeamNamen.id"
    sqlstr = sqlstr & " Where groepsindeling.ksID = " & kampID
    sqlstr = sqlstr & " ORDER BY TeamNamen.naam"
    Me.dtaTeams.ConnectionString = cn.ConnectionString
    Me.dtaTeams.RecordSource = sqlstr
    Me.dtaTeams.Refresh
    
    If grp < 8 Then
        wdnr = GetFirstFinaleMatch(KwartFinale)
        strtWedNumLabel = 8
    Else
        wdnr = GetFirstFinaleMatch(AchtsteFinale)
        strtWedNumLabel = 0
    End If
    lastwednr = GetFirstFinaleMatch(Finale)
    j = 0
    For i = 0 To lastwednr - wdnr
        Me.lblWedNum(i + strtWedNumLabel).Caption = wdnr + i
        If Not getKampInfo("derdeplaats") Then
            Me.lblWedNum(Me.lblWedNum.UBound) = wdnr + i
            If j = 0 Then j = 16
        End If
        If j < 28 Then
            Me.lblFin(j).Caption = GetWedCode(i + wdnr, 1)
            j = j + 1
            Me.lblFin(j).Caption = GetWedCode(i + wdnr, 0)
            j = j + 1
        End If
    Next
    If grp > 4 Then
        For i = 0 To grp - 1 Step 2
'            Me.lblKwFin(i).Caption = wdnr + i
'            Me.lblKwFin(i + 1).Caption = wdnr + i + 1
        Next
    End If
    wdnr = GetFirstFinaleMatch(KwartFinale)
    For i = 0 To 3 Step 2
'        Me.lblHvFin(i).Caption = wdnr + i
'        Me.lblHvFin(i + 1).Caption = wdnr + i + 1
    Next
    Me.frm3ePl.Visible = getKampInfo("derdeplaats")
    sqlstr = "Select * from qryVoorspelAantalllen WHERE poolID = " & poolID
    sqlstr = sqlstr & " AND cat =1"
    
    Me.dtaVoorspAantallen.ConnectionString = cn.ConnectionString
    Me.dtaVoorspAantallen.RecordSource = sqlstr
    Me.dtaVoorspAantallen.Refresh
    With Me.dtaVoorspAantallen.Recordset
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            Me.lblVoorsp(0).Caption = !omschrijving
            If .RecordCount > 1 Then
                .MoveNext
                Do While Not .EOF
                i = Me.lblVoorsp.UBound + 1
                    Load Me.lblVoorsp(i)
                    Me.lblVoorsp(i).Top = Me.lblVoorsp(i - 1).Top + Me.lblVoorsp(i - 1).Height
                    Me.lblVoorsp(i).Visible = True
                    Me.lblVoorsp(i).Caption = !omschrijving
                    .MoveNext
                    Load Me.txtVoorsp(i)
                    Me.txtVoorsp(i).Top = Me.txtVoorsp(i - 1).Top + Me.txtVoorsp(i - 1).Height
                    Me.txtVoorsp(i).Visible = True
                    Me.txtVoorsp(i).TabIndex = Me.txtVoorsp(i - 1).TabIndex + 1
                Loop
            End If
        Else
            Me.lblVoorsp(0).Visible = False
            Me.txtVoorsp(0).Visible = False
        End If
    End With
    Me.dtaVoorspGroepStnd.ConnectionString = cn.ConnectionString
    Me.dtaVoorspGroepStnd.RecordSource = "Select * from voorspelling_groepstand"
    
    For i = 2 To 3
        Me.eind(i).Visible = CBool(getKampInfo("derdeplaats"))
        Me.lblPl(i).Visible = CBool(getKampInfo("derdeplaats"))
    Next
    For i = 0 To 2
        Me.cmbTS(i).Visible = getPntToek("topscorer " & Format(i + 1, "0")) > 0
        Me.Label11(i + 4).Visible = getPntToek("topscorer " & Format(i + 1, "0")) > 0
        Me.dpTS(i).Visible = getPntToek("topscorer " & Format(i + 1, "0")) > 0
    Next
    If Me.dtaDezeDeelnemer.Recordset.RecordCount > 0 Then
        Me.dtaDezeDeelnemer.Recordset.MoveFirst
        deeln = Me.dtaDezeDeelnemer.Recordset!deelnemID
        Me.dtaDeelnems.Refresh
        If Me.dtaDeelnems.Recordset.RecordCount > 0 Then
            deeln = Me.dtaDeelnems.Recordset!deelnemID
        End If
        Me.cmbDeelnems.BoundText = deeln
        cmbDeelnems_Click 2
        BerekenTtl
    End If
    'Me.cmbAdressen.Enabled = getPoolInfo("eindinschr") >= Now()
    Screen.MousePointer = vbNormal
frmExit:
    Exit Sub
formErr:
If Err = -2147217908 Then
    
Else
    MsgBox Err & ":" & Error
End If
Resume Next
End Sub

Private Sub grdDeelnWeds_AfterColEdit(ByVal ColIndex As Integer)
Dim rs As New ADODB.Recordset
Dim selStr As String
Dim waarsch As Boolean
Dim ans As Boolean
    waarsch = Me.grdDeelnWeds.Columns(5).value > 10
    waarsch = waarsch Or nz(Me.grdDeelnWeds.Columns(6), 11) > 10
    waarsch = waarsch Or nz(Me.grdDeelnWeds.Columns(7), 11) > 10
    waarsch = waarsch Or nz(Me.grdDeelnWeds.Columns(8), 11) > 10
    waarsch = waarsch Or nz(Me.grdDeelnWeds.Columns(9), 4) > 3
    waarsch = waarsch Or nz(Me.grdDeelnWeds.Columns(9), 0) = 0
    ans = True
    If waarsch Then
        ans = MsgBox("Klopt dit wel", vbYesNo + vbQuestion, "Uitslag") = vbYes
    End If
    selStr = "select * from voorspelling_uitsl WHERE deelnem = " & deeln
    selStr = selStr & " AND wednum = " & Me.grdDeelnWeds.Columns(2)
    rs.Open selStr, cn, adOpenStatic, adLockOptimistic
    If ans Then
        rs!r1 = Me.grdDeelnWeds.Columns(5)
        rs!r2 = Me.grdDeelnWeds.Columns(6)
        rs!e1 = Me.grdDeelnWeds.Columns(7)
        rs!e2 = Me.grdDeelnWeds.Columns(8)
        rs!toto = Me.grdDeelnWeds.Columns(9)
        rs.Update
        BerekenTtl
    Else
        Me.grdDeelnWeds.Columns(5) = rs!r1
        Me.grdDeelnWeds.Columns(6) = rs!r2
        Me.grdDeelnWeds.Columns(7) = rs!e1
        Me.grdDeelnWeds.Columns(8) = rs!e2
        Me.grdDeelnWeds.Columns(9) = rs!toto
    End If

    rs.Close
    Set rs = Nothing
End Sub


Sub BerekenTtl()
Dim rs As New ADODB.Recordset
Dim selStr As String
Dim gel As Integer
Dim dp As Integer
    selStr = "select * from voorspelling_uitsl WHERE deelnem = " & deeln
    rs.Open selStr, cn, adOpenStatic, adLockReadOnly
    Do While Not rs.EOF
        dp = dp + rs!e1 + rs!e2
        If rs!e1 = rs!e2 Then gel = gel + 1
        rs.MoveNext
    Loop
    Me.txtTtl(0) = gel
    Me.txtTtl(1) = dp
    rs.Close
    Set rs = Nothing
End Sub

Sub ZetStandaardFinales(grp As String, pl As Integer)
    'zet al vast de standaard waarden bij de achtste (of kwart)finales
    Dim rs As New ADODB.Recordset
    Dim rsFin As New ADODB.Recordset
    Dim selStr As String
    Dim grps As Integer
    Dim i As Integer
    Dim ctl(15) As Control
    selStr = "Select * from voorspelling_groepstand Where deelnem = " & deeln
    selStr = selStr & " AND groep = '" & grp & "'"
    rs.Open selStr, cn, adOpenStatic, adLockOptimistic
    grps = getKampInfo("groepen")
    selStr = "Select * from voorspelling_finales"
    selStr = selStr & " WHERE deelnem = " & deeln
    rsFin.Open selStr, cn, adOpenStatic, adLockOptimistic
    Select Case grp
    Case "A"
        If pl = 1 Then 'A1
            If getKampInfo("Groepen") > 4 Then
                rsFin.Find "wed =" & GetFirstFinaleMatch(AchtsteFinale)
            Else
                rsFin.Find "wed =" & GetFirstFinaleMatch(KwartFinale)
            End If
            rsFin!t1 = rs!pos1
            rsFin.Update
        Else 'A2
            rsFin.Find "wed =" & GetFirstFinaleMatch(AchtsteFinale) + 2
            rsFin!t2 = rs!pos2
            rsFin.Update
        End If
    Case "B"
        If pl = 1 Then 'B1
            rsFin.Find "wed =" & GetFirstFinaleMatch(AchtsteFinale) + 2
            rsFin!t1 = rs!pos1
            rsFin.Update
        Else        'B2
            rsFin.Find "wed =" & GetFirstFinaleMatch(AchtsteFinale)
            rsFin!t2 = rs!pos2
            rsFin.Update
        End If
    Case "C"
        If pl = 1 Then 'C1
            rsFin.Find "wed =" & GetFirstFinaleMatch(AchtsteFinale) + 1
            rsFin!t1 = rs!pos1
            rsFin.Update
        Else        'C2
            rsFin.Find "wed =" & GetFirstFinaleMatch(AchtsteFinale) + 3
            rsFin!t2 = rs!pos2
            rsFin.Update
        End If
    Case "D"
        If pl = 1 Then 'D1
            rsFin.Find "wed =" & GetFirstFinaleMatch(AchtsteFinale) + 3
            rsFin!t1 = rs!pos1
            rsFin.Update
        Else        'D2
            rsFin.Find "wed =" & GetFirstFinaleMatch(AchtsteFinale) + 1
            rsFin!t2 = rs!pos2
            rsFin.Update
        End If
    Case "E"
        If pl = 1 Then 'E1
            rsFin.Find "wed =" & GetFirstFinaleMatch(AchtsteFinale) + 4
            rsFin!t1 = rs!pos1
            rsFin.Update
        Else        'E2
            rsFin.Find "wed =" & GetFirstFinaleMatch(AchtsteFinale) + 6
            rsFin!t2 = rs!pos2
            rsFin.Update
        End If
    Case "F"
        If pl = 1 Then 'F1
            rsFin.Find "wed =" & GetFirstFinaleMatch(AchtsteFinale) + 6
            rsFin!t1 = rs!pos1
            rsFin.Update
        Else        'F2
            rsFin.Find "wed =" & GetFirstFinaleMatch(AchtsteFinale) + 4
            rsFin!t2 = rs!pos2
            rsFin.Update
        End If
    Case "G"
        If pl = 1 Then 'G1
            rsFin.Find "wed =" & GetFirstFinaleMatch(AchtsteFinale) + 5
            rsFin!t1 = rs!pos1
            rsFin.Update
        Else        'G2
            rsFin.Find "wed =" & GetFirstFinaleMatch(AchtsteFinale) + 7
            rsFin!t2 = rs!pos2
            rsFin.Update
        End If
    Case "H"
        If pl = 1 Then 'H1
            rsFin.Find "wed =" & GetFirstFinaleMatch(AchtsteFinale) + 7
            rsFin!t1 = rs!pos1
            rsFin.Update
        Else        'H2
            rsFin.Find "wed =" & GetFirstFinaleMatch(AchtsteFinale) + 5
            rsFin!t2 = rs!pos2
            rsFin.Update
        End If
    End Select
    rsFin.Close
    rs.Close
    Set rs = Nothing
    Set rsFin = Nothing
    deelnFinalePlaatsen
End Sub

Private Sub grdDeelnWedsOld_Click()

End Sub

Private Sub groep_Click(Index As Integer, Area As Integer)
    If Area = 2 Then
        If nz(Me.groep(Index).Text, "") > "" Then
            Dim selStr As String
'            Dim rs As New ADODB.Recordset
            selStr = "UPDATE voorspelling_groepstand"
            selStr = selStr & " SET pos" & Index + 1
            selStr = selStr & " = " & groep(Index).BoundText
            selStr = selStr & " WHERE deelnem = " & deeln
            selStr = selStr & " AND groep = '" & Me.cmbGroepskeus.Text & "'"
            cn.Execute selStr
'            Me.dtaVoorspGroepStnd.Refresh
'            selStr = "Select * from voorspelling_groepstand"
'            selStr = selStr & " WHERE deelnem = " & deeln
'            selStr = selStr & " AND groep = '" & Me.cmbGroepskeus.Text & "'"
'            rs.Open selStr, cn, adOpenDynamic, adLockOptimistic
'            rs("pos" & Format(Index + 1, "0")) = groep(Index).BoundText
'            rs.Update
'            rs.Close
'            Set rs = Nothing
        End If
   End If
End Sub

Private Sub groep_LostFocus(Index As Integer)
'groep_Click Index, 2
End Sub

Private Sub kwFin_Click(Index As Integer)
    curDeeln = deeln
    Select Case Index
    Case Is < 16
        frm8Fin.Show 1
    Case Is < 24
        frm4Fin.Show 1
    Case Is < 28
        frm2Fin.Show 1
    Case Is >= 28
        frm1Fin.Show 1
    End Select
    deelnFinalePlaatsen
End Sub

Private Sub txtBijnaam_LostFocus()
    Me.dtaDezeDeelnemer.Recordset.Update
    Me.dtaDeelnems.Refresh
    
    Me.dtaDeelnems.Recordset.Find "deelnemID=" & deeln
    Me.cmbDeelnems.Text = Me.txtBijnaam
End Sub

Private Sub txtVoorsp_GotFocus(Index As Integer)
Dim i As Integer
    Me.txtVoorsp(Index).SelStart = 0
    Me.txtVoorsp(Index).SelLength = Len(Me.txtVoorsp(Index).Text)
    'voor de zekerheid
    For i = 0 To 3
        If Me.eind(i).Visible Then
            savEindstand i
        End If
    Next
End Sub

Private Sub txtVoorsp_LostFocus(Index As Integer)
'update aantallen tabel
Dim rs As New ADODB.Recordset
Dim selStr As String
    selStr = "Select * from voorspelling_aantallen Where deelnem = " & deeln
    selStr = selStr & " AND voorspel_type = " & getVoorspelTypeID(Me.lblVoorsp(Index).Caption)
    rs.Open selStr, cn, adOpenStatic, adLockOptimistic
    If rs.RecordCount = 1 Then
        rs!Aantal = val(Me.txtVoorsp(Index).Text)
        rs.Update
    End If
rs.Close
Set rs = Nothing
End Sub
