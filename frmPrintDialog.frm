VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPrintDialog 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Afdrukken"
   ClientHeight    =   4815
   ClientLeft      =   1665
   ClientTop       =   2430
   ClientWidth     =   6540
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   HelpContextID   =   450
   Icon            =   "frmPrintDialog.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4815
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picCompetitorList 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   3120
      ScaleHeight     =   2145
      ScaleWidth      =   3315
      TabIndex        =   35
      Top             =   0
      Width           =   3345
      Begin VB.ListBox lstCompetitorPools 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   15
         MultiSelect     =   1  'Simple
         TabIndex        =   38
         Top             =   0
         Width           =   3240
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Allemaal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   240
         Left            =   240
         TabIndex        =   37
         Top             =   1830
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Selectie"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   240
         Left            =   1635
         TabIndex        =   36
         Top             =   1830
         Width           =   1230
      End
   End
   Begin VB.PictureBox picVolgorde 
      Appearance      =   0  'Flat
      FillColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3090
      ScaleHeight     =   450
      ScaleWidth      =   3315
      TabIndex        =   32
      Top             =   90
      Width           =   3345
      Begin VB.OptionButton ScoreVolg 
         Appearance      =   0  'Flat
         Caption         =   "Op score"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   390
         Index           =   1
         Left            =   1560
         TabIndex        =   34
         Top             =   -30
         Width           =   1080
      End
      Begin VB.OptionButton ScoreVolg 
         Appearance      =   0  'Flat
         Caption         =   "Alfabetisch"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   330
         Index           =   0
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.PictureBox picPrnterSettings 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1605
      Left            =   3090
      ScaleHeight     =   1575
      ScaleWidth      =   3315
      TabIndex        =   14
      Top             =   2280
      Width           =   3345
      Begin MSComCtl2.UpDown upDnCopies 
         Height          =   375
         Left            =   3000
         TabIndex        =   42
         Top             =   1125
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         BuddyControl    =   "txtCopies"
         BuddyDispid     =   196612
         OrigLeft        =   2760
         OrigTop         =   1200
         OrigRight       =   3015
         OrigBottom      =   1575
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.CheckBox chkNwePagKop 
         Alignment       =   1  'Right Justify
         Caption         =   "Nwe pag kop"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   1560
         TabIndex        =   31
         ToolTipText     =   "Print wel/niet de kopregels op een nieuwe pagina"
         Top             =   375
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.TextBox txtCopies 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   21
         Text            =   "1"
         Top             =   1125
         Width           =   480
      End
      Begin VB.ComboBox cmbPrinters 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Tag             =   "printer"
         Top             =   720
         Width           =   3135
      End
      Begin VB.OptionButton optLandscape 
         Caption         =   "Liggend"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.OptionButton optPortrait 
         Caption         =   "Staand"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Top             =   120
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CheckBox chkDblSide 
         Alignment       =   1  'Right Justify
         Caption         =   "Dubbelzijdig"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   1155
         Width           =   1425
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Afdruk opties"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   120
         TabIndex        =   41
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aantal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   240
         Left            =   1845
         TabIndex        =   20
         Top             =   1192
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Printer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   225
         TabIndex        =   19
         Top             =   435
         Width           =   570
      End
   End
   Begin VB.PictureBox Picture4 
      Align           =   2  'Align Bottom
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   0
      ScaleHeight     =   810
      ScaleWidth      =   6480
      TabIndex        =   12
      Top             =   3945
      Width           =   6540
      Begin VB.CommandButton KlaarButton 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sluiten"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5280
         TabIndex        =   30
         Tag             =   "SluitPrintDial"
         Top             =   360
         Width           =   1125
      End
      Begin VB.CommandButton btnPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Voorbeeld"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   2895
         TabIndex        =   29
         ToolTipText     =   "Bekijk een voorbeeld op het scherm"
         Top             =   360
         Width           =   1125
      End
      Begin VB.CommandButton btnPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Afdrukken"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   4080
         TabIndex        =   28
         ToolTipText     =   "Stuur dit rapport naar de printer"
         Top             =   360
         Width           =   1125
      End
      Begin VB.CommandButton btnFinalPlayerPrint 
         Caption         =   "Eindstand voor deelnemers  afdrukken"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2895
         TabIndex        =   27
         Top             =   60
         Width           =   3510
      End
      Begin VB.CheckBox chkEindstand 
         Appearance      =   0  'Flat
         Caption         =   "Eind stand"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   615
         Left            =   1800
         TabIndex        =   22
         Tag             =   "chkEinstand"
         Top             =   120
         Width           =   915
      End
      Begin VB.CommandButton btnPrntAllAfterDay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Alles aan einde dag afdrukken"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   75
         TabIndex        =   13
         Top             =   135
         Width           =   1620
      End
   End
   Begin VB.PictureBox picVoorWed 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3090
      ScaleHeight     =   450
      ScaleWidth      =   3315
      TabIndex        =   9
      Top             =   1050
      Width           =   3345
      Begin MSComCtl2.UpDown upDnForMatch 
         Height          =   375
         Left            =   2251
         TabIndex        =   40
         Top             =   15
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         BuddyControl    =   "txtForMatch"
         BuddyDispid     =   196629
         OrigLeft        =   2520
         OrigRight       =   2775
         OrigBottom      =   375
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtForMatch 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Text            =   "1"
         Top             =   15
         Width           =   450
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "Voor wedstrijd nr:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   75
         TabIndex        =   11
         Top             =   45
         Width           =   1875
      End
   End
   Begin VB.PictureBox picTMwed 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3090
      ScaleHeight     =   450
      ScaleWidth      =   3315
      TabIndex        =   6
      Top             =   570
      Width           =   3345
      Begin MSComCtl2.UpDown upDnToMatch 
         Height          =   375
         Left            =   2251
         TabIndex        =   39
         Top             =   30
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         BuddyControl    =   "txtToMatch"
         BuddyDispid     =   196633
         OrigLeft        =   2520
         OrigTop         =   30
         OrigRight       =   2775
         OrigBottom      =   405
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtToMatch 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Text            =   "1"
         Top             =   30
         Width           =   450
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "T/m wedstrijd nr:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   60
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3795
      Left            =   75
      ScaleHeight     =   3765
      ScaleWidth      =   2850
      TabIndex        =   0
      Tag             =   "afdruk"
      Top             =   90
      Width           =   2880
      Begin VB.OptionButton optPrintDoc 
         Appearance      =   0  'Flat
         Caption         =   "Punten samenstelling"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   315
         Index           =   8
         Left            =   90
         TabIndex        =   26
         Top             =   2135
         Width           =   2670
      End
      Begin VB.OptionButton optPrintDoc 
         Appearance      =   0  'Flat
         Caption         =   "Voorspelling"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   315
         Index           =   7
         Left            =   90
         TabIndex        =   25
         Top             =   1323
         Width           =   2670
      End
      Begin VB.OptionButton optPrintDoc 
         Appearance      =   0  'Flat
         Caption         =   "Punten per wedstrijd"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   315
         Index           =   6
         Left            =   90
         TabIndex        =   24
         Top             =   2541
         Width           =   2670
      End
      Begin VB.OptionButton optPrintDoc 
         Appearance      =   0  'Flat
         Caption         =   "Stand in de Pool"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   315
         Index           =   2
         Left            =   90
         TabIndex        =   23
         Top             =   1729
         Width           =   2670
      End
      Begin VB.OptionButton optPrintDoc 
         Appearance      =   0  'Flat
         Caption         =   "Inschrijffomulieren"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Top             =   105
         Width           =   2670
      End
      Begin VB.OptionButton optPrintDoc 
         Appearance      =   0  'Flat
         Caption         =   "Ingevulde Pools"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   315
         Index           =   1
         Left            =   90
         TabIndex        =   4
         Top             =   511
         Width           =   2670
      End
      Begin VB.OptionButton optPrintDoc 
         Appearance      =   0  'Flat
         Caption         =   "Stand in toernooi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   315
         Index           =   4
         Left            =   90
         TabIndex        =   3
         Top             =   3360
         Width           =   2670
      End
      Begin VB.OptionButton optPrintDoc 
         Appearance      =   0  'Flat
         Caption         =   "Grafiek pool stand"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   315
         Index           =   5
         Left            =   90
         TabIndex        =   2
         Top             =   2947
         Width           =   2670
      End
      Begin VB.OptionButton optPrintDoc 
         Appearance      =   0  'Flat
         Caption         =   "Favorieten"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   315
         Index           =   3
         Left            =   90
         TabIndex        =   1
         Top             =   917
         Width           =   2670
      End
   End
   Begin MSComDlg.CommonDialog printerDialog 
      Left            =   2760
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FontName        =   "Tahoma"
   End
End
Attribute VB_Name = "frmPrintDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

'om af te drukken op gekleurde achtergrond
 Private Declare Function SetBkMode Lib "gdi32" _
  (ByVal hdc As Long, ByVal nBkMode As Long) As Long

Private Declare Function GetBkMode Lib "gdi32" _
  (ByVal hdc As Long) As Long

Private Const TRANSPARENT = 1
Private Const OPAQUE = 2
Dim afdrratio

Private iBKMode As Long

'OLD STUFF
Dim headerText
Dim tmwed As Integer
Dim KolHeight As Integer
Dim kolwidth As Integer
Dim kop As String
Dim kol As Integer
Dim Y As Integer
'voor de favorieten afdruk
Dim favYpos As Integer
Dim favXpos As Integer

Dim kopFont As String
Dim txtFont As String

Dim rot As New rotator

Dim X As Integer

Dim RegHeight As Integer
Dim RenScore()
Dim RenPntidx() As Integer
Dim NormHoog As Integer
Dim GrootHoog As Integer
Dim KleinHoog As Integer
Dim SmallHoog As Integer
Dim kophoog As Integer
Dim voethoog As Integer
Dim prnFont As String
Dim wedNu As Integer
Dim obj As Object
Dim maxY As Integer 'voor afdrukken van Favorieten

Dim kleur(64) As Long 'voor grafiek

Dim printPrev As printPreview

Private Sub optPrintDoc_Click(Index As Integer)
Dim i As Integer
Me.picCompetitorList.Visible = False
Me.optPrintDoc(Index).value = True
Select Case Index
  Case 0
    Me.picTMwed.Visible = False
    Me.picVoorWed.Visible = False
    Me.picVolgorde.Visible = False
    Me.optPortrait.value = True
    Me.picCompetitorList.Visible = False
   ' Me.chkDblSide.Value = 1

  Case 1
   'deelnemers met voorspellingen
    Me.picTMwed.Visible = False
    Me.picVoorWed.Visible = False
    Me.picVolgorde.Visible = False
    Me.picCompetitorList.Visible = True
    'Me.txtVoorwed.SetFocus
    Me.optPortrait.value = True
    'txtVoorwed.SetFocus
  Case 2
    'score/ stand in de pool
    picVolgorde.Visible = True 'GetDeelnemAant(thisPool) > 32
    picVoorWed.Visible = False
    picTMwed.Visible = True
    Me.optPortrait.value = True
    'Me.vscrlTM = GetMyNum(GetLastPlayed)
    If tmwed > 0 Then
        Me.upDnToMatch.SetFocus
    End If
  Case 3
    ' Favorieten
    Me.picTMwed.Visible = False
    Me.picVoorWed.Visible = False
    Me.picVolgorde.Visible = False
    Me.optPortrait.value = True
    Me.picCompetitorList.Visible = False
  Case 4
    'Stand in toernooi
    'score/ stand in de pool
    Me.picTMwed.Visible = False
    Me.picVoorWed.Visible = False
    Me.picVolgorde.Visible = False
    Me.optPortrait.value = True
    Me.picCompetitorList.Visible = False
    'Me.vscrlTM = GetMyNum(GetLastPlayed())
    DoEvents
 Case 5
    'grafiek
    Me.picVolgorde.Visible = False
    Me.picVoorWed.Visible = False
    Me.picTMwed.Visible = True
    Me.optLandscape.value = True
    Me.ScoreVolg(1) = True
    'Me.vscrlTM = GetMyNum(GetLastPlayed())
    DoEvents
    tmwed = Me.upDnToMatch
  Case 6
    'punten per wedstrijd
    picVolgorde.Visible = True
    picVoorWed.Visible = False
    picTMwed.Visible = True
    'Me.vscrlTM = GetMyNum(GetLastPlayed())
    tmwed = Me.upDnToMatch
    DoEvents
    Me.picCompetitorList.Visible = False  'getTournamentInfo("groepen")
    Me.optLandscape.value = getTournamentInfo("tournamentGroupCount", cn) > 4
    Me.optPortrait.value = Not Me.optLandscape.value
  Case 7
    'voorspelling per wedstrijd
    picVolgorde.Visible = False
    picVoorWed.Visible = True
    picTMwed.Visible = False
    Me.optPortrait.value = True
    Me.optLandscape.value = False
    'Me.vscrlVoor = GetMyNum(GetLastPlayed()) + 1
    Me.picCompetitorList.Visible = False
  Case 8
    'samenvatting stand
    'Stand in toernooi
    'score/ stand in de pool
    Me.picTMwed.Visible = True
    Me.picVoorWed.Visible = False
    Me.picVolgorde.Visible = True
    Me.optLandscape.value = True
    Me.picCompetitorList.Visible = False
    'Me.vscrlTM = GetMyNum(GetLastPlayed())
  End Select
End Sub

Sub horline(kleur As Integer)
    obj.Line (0, obj.CurrentY)-(obj.ScaleWidth - 50, obj.CurrentY), kleur
End Sub

Sub FormulierenAfdrukken()
Dim txt As String
Dim i As Integer
Dim aant As Integer
Dim amount As Integer
Dim topY As Integer
Dim ypos As Integer
Dim xpos As Integer
Dim kopAnaam As String
Dim kopVnaam As String
Dim sqlstr As String
'Dim rs As New ADODB.Recordset
    obj.FillStyle = vbFSTransparent
    headerText = getOrganisation(cn) & getTournamentInfo("description", cn) & " voetbalpool"
    kop$ = "Inschrijfformulier     inleg: " & Format(getPoolInfo("poolCost", cn), "currency")
    obj.FontName = "Times New Roman"
    InitPage False, True
    obj.Print
    FontGr 12
    topY = obj.CurrentY
    obj.ForeColor = vbBlack
    Vet False
    FontGr 12
    obj.CurrentY = topY
    FontGr 18
    obj.Line (0, topY - 200)-(obj.ScaleWidth + 2 * obj.ScaleLeft, topY + obj.TextHeight("WW") * 4 + 200), , B
    obj.Print
    xpos = obj.CurrentX + 200
    obj.CurrentY = topY
    obj.CurrentX = xpos
    obj.Print "Naam: ....................................................... Telefoon....................................."
    obj.CurrentY = topY + obj.TextWidth("WW")
    obj.CurrentX = xpos
    obj.Print "Adres: ....................................................... Plaats.........................................."
    obj.CurrentY = topY + obj.TextWidth("WW") * 2
    obj.CurrentX = xpos
    obj.Print "Email: ....................................................... Betaald ";
    xpos = obj.CurrentX
    ypos = obj.CurrentY
    obj.DrawWidth = 3
    obj.Line (xpos, ypos)-(xpos + obj.TextWidth("W"), ypos + obj.TextHeight("W")), , B
    obj.DrawWidth = 1
    obj.CurrentY = ypos
    obj.CurrentX = obj.CurrentX + 30
    obj.Print " bij............................"
    FontGr 4
    obj.Print
    'sqlstr = "Select * from poolpnt Where thisPool = " & thisPool
    'rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    FontGr 16
    Vet True
    obj.ForeColor = vbBlue
    obj.Print "Instructies"
    FontGr 11
    Vet False
    obj.ForeColor = vbBlack
    obj.Print "Hier onder (en op de achterkant) kun je voorspellingen invoeren voor de "; getTournamentInfo("description", cn);
    obj.Print " van "; Format(getTournamentInfo("tournamentstartdate", cn), "d MMMM yyyy"); " tot "; Format(getTournamentInfo("tournamentEnddate", cn), "d MMMM yyyy")
    obj.Print "Voor elke juiste voorspelling krijg je punten, bij de verschillende onderdelen staat hoeveel."
    obj.Print "De voorspellingen hoeven niet te kloppen, bij een uitslag kun je bijvoorbeeld 1-0 bij de rust, 0-2 bij de eindstand en een 3 "
    obj.Print "bij de toto invullen. Of je kunt een team dat je uitgeschakeld hebt in een volgende ronde toch weer opnemen."
    If getTournamentInfo("tournamentGroupCount", cn) = 6 And getTournamentInfo("tournamentTeamCount", cn) = 24 Then ' de vier beste derde plaatsen naar kwart finales
      obj.Print "De beste 4 derde plaatsen kwalificeren zich ook voor de 8e finales."
    End If
    FontGr 16
    Vet True
    obj.ForeColor = vbBlue
    obj.Print "Prijzen"
    FontGr 11
    Vet False
    obj.ForeColor = vbBlack
    'obj.Print "Na de finale worden de hoofdprijzen te verdeeld, maar ook per dag zijn er geldprijzen te winnen."
    Vet True
    obj.Print "-  Per dag";
    Vet False
    obj.Print " zijn de volgende geldprijzen te verdienen:"
    obj.Print "  -  ";
    obj.Print "Degene die op ";
    Ital True
    obj.Print "één dag de meeste punten";
    Ital False
    obj.Print " heeft verzameld, ";
    obj.Print " krijgt daarvoor ";
    Vet True
    obj.Print Format(getPoolInfo("prizeHighDayScore", cn), "currency")
    Vet False
    obj.Print "  -  ";
    obj.Print "Degene die na een dag in de ";
    Ital True
    obj.Print "totaalstand bovenaan";
    Ital False
    obj.Print " staat, ";
    obj.Print " krijgt daarvoor ";
    Vet True
    obj.Print Format(getPoolInfo("prizeHighDayPosition", cn), "currency")
    Vet False
    obj.Print "  -  ";
    obj.Print "Degene die na een dag in de ";
    Ital True
    obj.Print "totaalstand onderaan";
    Ital False
    obj.Print " staat, ";
    obj.Print " krijgt daarvoor als troost ";
    Vet True
    obj.Print Format(getPoolInfo("prizeLowDayPosition", cn), "currency")
    Vet False
    obj.Print "  -  ";
    xpos = obj.CurrentX
    obj.Print "De punten voor de finalerondes tellen mee voor de dagprijs op de dag dat de teams bekend zijn"
    obj.CurrentX = xpos
    obj.Print "De punten voor de eindstand, topscorers en aantallen tellen op de dag van de finale mee voor de dagprijs"
    obj.Print "-  ";
    Vet True
    obj.Print "Aan het eind van het toernooi";
    Vet False
    obj.Print " zijn de volgende geldprijzen te verdienen:"
    amount = getPoolInfo("prizeLowFinalPosition", cn)
    If amount > 0 Then
        obj.Print "  -  ";
        xpos = obj.CurrentX
        obj.Print "De ";
        Ital True
        obj.ForeColor = vbRed
        obj.Print "rode lantaarn";
        obj.ForeColor = vbBlack
        Ital False
        obj.Print " ontvangt als troostprijs "; Format(amount, "currency")
    End If
    
    obj.Print "  -  ";
    xpos = obj.CurrentX
    obj.Print "De ";
    Ital True
    obj.Print "hoogste";
    Ital False
    obj.Print " deelnemers in de totaalstand krijgen de volgende prijzen:"
    obj.CurrentX = xpos
    
    obj.Print "1e pl: ";
    Vet True
    obj.Print Format(getPoolInfo("prizePercentageFirst", cn) / 100, "0%");
    Vet False
    amount = getPoolInfo("prizePercentageSecond", cn)
    If amount > 0 Then
        obj.Print ", 2e pl: ";
        Vet True
        obj.Print Format(amount / 100, "0%");
        Vet False
    End If
    amount = getPoolInfo("prizePercentageThird", cn)
    If amount > 0 Then
        obj.Print ", 3e pl: ";
        Vet True
        obj.Print Format(amount / 100, "0%");
        Vet False
    End If
    amount = getPoolInfo("prizePercentageFourth", cn)
    If amount > 0 Then
        obj.Print ", 4e pl: ";
        Vet True
        obj.Print Format(amount / 100, "0%");
        Vet False
    End If
    obj.Print " van de totale inleg (minus de dagprijzen en de rode lantaarn)"
    obj.Print "-  ";
    Ital True
    obj.Print "Bij een gelijk aantal punten wordt de betreffende prijs verdeeld"
    Ital False
    'horline 1
    'groepsstanden
    FontGr 10
    obj.Print
    Y = obj.CurrentY
    X = obj.CurrentX
    FontGr 14
    Vet True
    obj.FillColor = &H808080
    obj.FillStyle = vbFSSolid
    'obj.BackColor = obj.FillColor
    obj.Line (X, Y - 10)-(obj.ScaleWidth, Y + obj.TextHeight("W") + 10), vbBlack, B
    obj.CurrentY = Y
    obj.CurrentX = X + 50
    iBKMode = SetBkMode(obj.hdc, TRANSPARENT)
    obj.ForeColor = vbWhite
    obj.Print "Groepstanden";
    FontGr 10
    Vet False
'    txt = " Vul in: 1 t/m 4 (" & getPntToek("groepstand per juist team") & " pnt per correcte invoer)"
'    'obj.CurrentX = obj.ScaleWidth - obj.TextWidth(txt)
'    obj.CurrentY = Y + 40
'    obj.Print txt;
'    obj.CurrentY = Y
'    FontGr 14
'    obj.Print
'    Y = obj.CurrentY
'    X = obj.CurrentX
'    FontGr 12
'    obj.FillStyle = vbFSTransparent
'    obj.Line (X, Y)-(obj.ScaleWidth, Y + obj.TextHeight("W") * 5), vbBlack, B
'    obj.FillStyle = vbFSTransparent
'    kolwidth = obj.ScaleWidth / getTournamentInfo("groepen")
'    obj.ForeColor = vbBlack
'    For i = 1 To getTournamentInfo("groepen")
'        FontGr 12
'        X = kolwidth * (i - 1) + 50
'        obj.CurrentY = Y + 10
'        obj.CurrentX = X
'        Vet True
'        obj.Print "Groep " & Chr(i + 64)
'        Vet False
'        printgroep i
'    Next
'    obj.Print
'    obj.Font = "Times New Roman"
'    FontGr 2
'    obj.Print
'    FontGr 12
'    printFinals
'    printOverige
'    kop$ = "Wedstrijdvoorspellingen"
'    DoNewPage False, True
'    formulierWeds
'    InvulFormAfdrukken
End Sub

Sub printOverige()
'invulformulier
Dim rs As New ADODB.Recordset
Dim topscAant As Integer
Dim ypos As Integer
Dim xpos As Integer
Dim newlinepos As Integer
Dim kolwidth As Integer
Dim i As Integer

Dim Y As Integer
Dim X As Integer
Dim pnt As Integer
Dim txt As String
    newlinepos = 0
    obj.Print
    kolwidth = obj.ScaleWidth / 4
    'eerst de eindstand
    ypos = obj.CurrentY
    i = getPntToek("1e plaats(Kampioen)")
    
    If i > 0 Then
        'print 1e
        txt = "(" & i & "p)"
        obj.Font = "Tahoma"
        Y = ypos
        obj.CurrentY = Y
        obj.CurrentX = 0
        X = obj.CurrentX
        FontGr 14
        Vet True
        obj.FillColor = &H808080
        obj.FillStyle = vbFSSolid
        obj.Line (X + 30, Y - 10)-(kolwidth - 30, Y + obj.TextHeight("W")), vbBlack, B
        obj.CurrentY = Y
        obj.CurrentX = X + 80
        obj.ForeColor = vbWhite
        obj.Print "Eindstand "
        Vet False
        obj.FillStyle = vbFSTransparent
        Y = obj.CurrentY
        obj.CurrentX = X + 80
        obj.ForeColor = vbBlack
        FontGr 12
        obj.Print "1e:";
        FontGr 14
        obj.Line (X + 30, Y)-(kolwidth - 30, Y + obj.TextHeight("W")), vbBlack, B
        obj.CurrentY = Y + 20
        obj.CurrentX = X + kolwidth - obj.TextWidth(txt) + 20
        FontGr 10
        obj.Print txt;
        obj.CurrentY = Y
        FontGr 14
        obj.Print
        For i = 2 To 4
            pnt = getPntToek(Format(i, "0") & "e plaats")
            If pnt > 0 Then
                Y = obj.CurrentY
                txt = "(" & pnt & "p)"
                obj.CurrentX = X + 80
                FontGr 12
                obj.Print Format(i, "0") & "e:";
                FontGr 14
                obj.Line (X + 30, Y)-(kolwidth - 30, Y + obj.TextHeight("W")), vbBlack, B
                obj.CurrentY = Y + 20
                obj.CurrentX = X + kolwidth - obj.TextWidth(txt) + 20
                FontGr 10
                obj.Print txt;
                obj.CurrentY = Y
                FontGr 14
                obj.Print
                If newlinepos < obj.CurrentY Then newlinepos = obj.CurrentY
            End If
        Next
    End If
    'topscorers
    obj.CurrentY = ypos
    i = getPntToek("topscorer 1")
    If i > 0 Then
        'print 1e
        txt = "(" & i & "p)"
        obj.Font = "Tahoma"
        Y = ypos
        obj.CurrentY = Y
        obj.CurrentX = kolwidth
        X = obj.CurrentX
        FontGr 14
        Vet True
        obj.FillColor = &H808080
        obj.FillStyle = vbFSSolid
        obj.Line (X, Y - 10)-(X + kolwidth * 1.3, Y + obj.TextHeight("W")), vbBlack, B '(X + kolwidth - 30, Y + obj.TextHeight("W")), vbBlack, B
        obj.CurrentY = Y
        obj.CurrentX = X + 50
        obj.ForeColor = vbWhite
        obj.Print "Topscorer";
        If getPntToek("topscorer 2") > 0 Then obj.Print "s";
        
        pnt = getPntToek("doelpunten topscorer 1")
        
        If pnt > 0 Then
            FontGr 14
            'obj.Line (X + kolwidth - 30, Y - 10)-(X + kolwidth * 1.3, Y + obj.TextHeight("W")), vbBlack, B
            obj.CurrentY = Y
            'obj.CurrentX = X + kolwidth + 20
            obj.Print " & aantal goals"
        Else
            obj.Print
        End If
        obj.FillStyle = vbFSTransparent
        obj.ForeColor = vbBlack
        Vet False
        For i = 1 To 3
            pnt = getPntToek("topscorer " & Format(i, "0"))
            If pnt > 0 Then
                Y = obj.CurrentY
                txt = "(" & pnt & "p)"
                obj.CurrentX = X + 50
                FontGr 12
                obj.Print Format(i, "0") & ":";
                FontGr 14
                obj.Line (X, Y)-(X + kolwidth - 30, Y + obj.TextHeight("W")), vbBlack, B
                obj.CurrentY = Y + 20
                obj.CurrentX = X + kolwidth + 20 - obj.TextWidth(txt)
                FontGr 10
                obj.Print txt;
                FontGr 14
                pnt = getPntToek("doelpunten topscorer " & Format(i, "0"))
                If pnt > 0 Then
                    obj.Line (X + kolwidth - 30, Y)-(X + kolwidth * 1.3, Y + obj.TextHeight("W")), vbBlack, B
                    obj.CurrentY = Y + 20
                    obj.CurrentX = X + kolwidth * 1.3 - obj.TextWidth("(" & pnt & "p)") + 50
                    FontGr 10
                    obj.Print "("; Format(pnt, pntFormat); "p)"
                    If newlinepos < obj.CurrentY Then newlinepos = obj.CurrentY
                Else
                    obj.Print
                End If
                obj.CurrentY = Y
                FontGr 14
                obj.Print
            End If
        Next
    End If
'overigen
Dim sqlstr As String
  sqlstr = "Select omschrijving, pnt, marge from voorspeltypes INNER JOIN pnttoek ON voorspeltypes.id = pnttoek.voorspeltype"
  sqlstr = sqlstr & " WHERE voorspeltypes.cat = 1 and pnttoek.poolid = " & thisPool
  sqlstr = sqlstr & " ORDER BY pnt, volgorde"
  rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    rs.Open "Select * from voorspeltypes where cat =1 order by volgorde", cn, adOpenStatic, adLockReadOnly
    Y = ypos
    obj.CurrentY = Y
    obj.CurrentX = X + kolwidth * 1.3 + 30
    X = obj.CurrentX
    FontGr 14
    Vet True
    obj.FillColor = &H808080
    obj.FillStyle = vbFSSolid
    obj.Line (X, Y - 10)-(obj.ScaleWidth - 50, Y + obj.TextHeight("W")), vbBlack, B
    obj.CurrentY = Y
    obj.CurrentX = X + 50
    obj.ForeColor = vbWhite
    obj.Print "Overigen "
    obj.FillStyle = vbFSTransparent
    obj.ForeColor = vbBlack
    Vet False
    Do While Not rs.EOF
        pnt = rs!pnt
        Y = obj.CurrentY
        txt = "(" & pnt & "p)"
        If nz(rs!marge, 0) > 0 Then
          txt = "(±" & rs!marge & ", " & pnt & "p)"
        End If
        obj.CurrentX = X + 50
        FontGr 10
        obj.Print rs!omschrijving; " "; txt; ":";
        FontGr 14
        obj.Line (X, Y)-(obj.ScaleWidth - 50, Y + obj.TextHeight("W")), vbBlack, B
        rs.MoveNext
        If newlinepos < obj.CurrentY Then newlinepos = obj.CurrentY
    Loop
    rs.Close
    Set rs = Nothing
    obj.Line (obj.ScaleWidth - 30 - obj.TextWidth("1234"), ypos + 360)-(obj.ScaleWidth - 30 - obj.TextWidth("1234"), obj.CurrentY)
    obj.Line (0, ypos - 50)-(obj.ScaleWidth - 10, newlinepos + 30), , B
    
End Sub
Sub printFinals()
'onderdeel van formulieren
Dim rs As New ADODB.Recordset
Dim sqlstr As String
Dim xpos As Integer
Dim ypos As Integer
Dim i As Integer
Dim X As Integer
Dim Y As Integer
Dim txt As String
Dim intP As Integer
Dim intQ As Integer
Dim hvFinYpos As Integer
Dim HeeftKlFin As Boolean
    i = getPntToek("achtste finaleplaats") + getPntToek("achtste finalepositie")
    If i > 0 Then
        'print achtste finales
        txt = "("
        intP = getPntToek("achtste finaleplaats")
        intQ = getPntToek("achtste finalepositie")
        If intP > 0 Then txt = txt & intP & " pnt voor elk genoemd team"
        
        If intQ > 0 Then
            If txt > "(" Then txt = txt & " of "
            txt = txt & intQ & " pnt als het ook nog op de juiste plaats staat"
        Else
            txt = txt & ", juiste plaats niet van belang"
        End If
        txt = txt & ")"
        obj.Font = "Tahoma"
        Y = obj.CurrentY
        X = obj.CurrentX
        FontGr 14
        Vet True
        obj.FillColor = &H808080
        obj.FillStyle = vbFSSolid
        obj.Line (X, Y - 10)-(obj.ScaleWidth, Y + obj.TextHeight("W")), vbBlack, B
        'obj.BackColor = obj.FillColor
        iBKMode = SetBkMode(obj.hdc, TRANSPARENT)
        obj.ForeColor = vbWhite
        obj.CurrentY = Y
        obj.CurrentX = X + 50
        obj.Print "Achtstefinales ";
        obj.FillStyle = vbFSTransparent
        FontGr 10
        Vet False
'        obj.CurrentX = obj.ScaleWidth - obj.TextWidth(txt)
        obj.CurrentY = Y + 40
        obj.Print txt;
        obj.ForeColor = vbBlack
        obj.CurrentY = Y
        FontGr 14
        obj.Print
        Y = obj.CurrentY
        X = obj.CurrentX
        FontGr 12
        obj.Line (X, Y)-(obj.ScaleWidth, Y + obj.TextHeight("W") * 4.7), vbBlack, B
        Y = Y + 50
        obj.CurrentY = Y
        obj.FillStyle = vbFSTransparent
        kolwidth = obj.ScaleWidth / 4
        sqlstr = "Select * from qryWeds where  ksid = " & kampID
        sqlstr = sqlstr & " and wedtype = 5"
        sqlstr = sqlstr & " ORDER BY wednum"
        rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
        xpos = 0
        With rs
            If .RecordCount > 0 Then
                i = 0
                Do While Not .EOF
                    ypos = Y
                    FontGr 8
                    obj.CurrentX = xpos + 50
                    obj.CurrentY = ypos + obj.TextHeight("99") * 0.5
                    obj.Print Format(!wedNum, "0"); ":";
                    FontGr 12
                    obj.CurrentX = xpos + obj.TextWidth("00:") + 30
                    obj.CurrentY = ypos
                    FontGr 10
                    obj.Print !code1; ":";
                    FontGr 12
                    obj.DrawWidth = 1
                    obj.Line (xpos + obj.TextWidth("00:"), ypos)-(xpos + kolwidth - 50, ypos + obj.TextHeight("W")), vbBlack, B
                    ypos = obj.CurrentY
                    obj.CurrentX = xpos + obj.TextWidth("00:") + 30
                    FontGr 10
                    obj.Print !code2; ":";
                    FontGr 12
                    obj.Line (xpos + obj.TextWidth("00:"), ypos)-(xpos + kolwidth - 50, ypos + obj.TextHeight("W")), vbBlack, B
                    'wedstrijd nr
                    obj.CurrentY = ypos
                    .MoveNext
                    i = i + 1
                    xpos = kolwidth * i
                    If xpos > obj.ScaleWidth - kolwidth + 100 Then
                        FontGr 8
                        obj.Print
                        obj.Print
                        FontGr 10
                        Y = obj.CurrentY
                        i = 0
                        xpos = 0
                    End If
                Loop
            End If
        End With
    End If
    FontGr 2
    obj.Print
    FontGr 12
    i = getPntToek("kwart finaleplaats") + getPntToek("kwart finalepositie")
    If i > 0 Then
        'print kwart finales
        txt = "("
        intP = getPntToek("kwart finaleplaats")
        intQ = getPntToek("kwart finalepositie")
        If intP > 0 Then txt = txt & intP & " pnt voor elk genoemd team"
        If intQ > 0 Then
            If txt > "(" Then txt = txt & " of "
            txt = txt & intQ & " pnt als het ook nog op de juiste plaats staat"
        Else
            txt = txt & ", juiste plaats hoeft niet"
        End If
        txt = txt & ")"
        obj.Font = "Tahoma"
        Y = obj.CurrentY
        X = obj.CurrentX
        FontGr 14
        Vet True
        obj.FillColor = &H808080
        obj.FillStyle = vbFSSolid
        obj.Line (X, Y - 10)-(obj.ScaleWidth, Y + obj.TextHeight("W")), vbBlack, B
        obj.CurrentY = Y
        obj.CurrentX = X + 50
        obj.ForeColor = vbWhite
        obj.Print "Kwartfinales ";
        FontGr 10
        Vet False
'        obj.CurrentX = obj.ScaleWidth - obj.TextWidth(txt)
        obj.CurrentY = Y + 40
        obj.Print txt;
        obj.ForeColor = vbBlack
        obj.FillStyle = vbFSTransparent
        obj.CurrentY = Y
        FontGr 14
        obj.Print
        Y = obj.CurrentY
        X = obj.CurrentX
        FontGr 12
        obj.Line (X, Y)-(obj.ScaleWidth, Y + obj.TextHeight("W") * 2.5), vbBlack, B
        Y = Y + 50
        obj.CurrentY = Y
        obj.FillStyle = vbFSTransparent
        kolwidth = (obj.ScaleWidth / 8) * 2
        sqlstr = "Select * from qryWeds where  ksid = " & kampID
        sqlstr = sqlstr & " and wedtype = 2"
        sqlstr = sqlstr & " ORDER BY wednum"
        rs.Close
        rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
        xpos = 0
        With rs
            If .RecordCount > 0 Then
                i = 0
                Do While Not .EOF
                    ypos = Y
                    FontGr 8
                    obj.CurrentX = xpos + 50
                    obj.CurrentY = ypos + obj.TextHeight("99") * 0.5
                    obj.Print Format(!wedNum, "0"); ":";
                    FontGr 12
                    obj.CurrentX = xpos + obj.TextWidth("00:") + 30
                    obj.CurrentY = ypos
                    FontGr 10
                    obj.Print !code1; ":";
                    FontGr 12
                    obj.DrawWidth = 1
                    obj.Line (xpos + obj.TextWidth("00:"), ypos)-(xpos + kolwidth - 50, ypos + obj.TextHeight("W")), vbBlack, B
                    ypos = obj.CurrentY
                    obj.CurrentX = xpos + obj.TextWidth("00:") + 30
                    FontGr 10
                    obj.Print !code2; ":";
                    FontGr 12
                    obj.Line (xpos + obj.TextWidth("00:"), ypos)-(xpos + kolwidth - 50, ypos + obj.TextHeight("W")), vbBlack, B
                    'wedstrijd nr
                    obj.CurrentY = ypos
                    .MoveNext
                    i = i + 1
                    xpos = kolwidth * i
                    If xpos > obj.ScaleWidth - kolwidth + 100 Then
                        FontGr 8
                        obj.Print
                        obj.Print
                        FontGr 12
                        Y = obj.CurrentY
                        i = 0
                        xpos = 0
                    End If
                Loop
            End If
        End With
    End If
    FontGr 2
    obj.Print
    FontGr 12
    hvFinYpos = obj.CurrentY
    i = getPntToek("halve finaleplaats") + getPntToek("halve finalepositie")
    If i > 0 Then
        'print halve finales
        txt = "("
        intP = getPntToek("halve finaleplaats")
        If intP > 0 Then txt = txt & intP & ""
        intQ = getPntToek("halve finalepositie")
        If intQ > 0 Then
            If txt > "(" Then txt = txt & "/"
            txt = txt & intQ & " pnt"
        Else
            txt = txt & " pnt"
        End If
        txt = txt & ")"
        obj.Font = "Tahoma"
        Y = obj.CurrentY
        X = obj.CurrentX
        FontGr 14
        Vet True
        obj.FillColor = &H808080
        obj.FillStyle = vbFSSolid
        obj.Line (X, Y - 10)-(obj.ScaleWidth / 2 - 30, Y + obj.TextHeight("W")), vbBlack, B
        obj.CurrentY = Y
        obj.CurrentX = X + 50
        obj.ForeColor = vbWhite
        obj.Print "Halve finales ";
        FontGr 10
        Vet False
        'obj.CurrentX = obj.ScaleWidth / 2 - 30 - obj.TextWidth(txt)
        obj.CurrentY = Y + 40
        obj.Print txt;
        obj.ForeColor = vbBlack
        obj.CurrentY = Y
        FontGr 14
        obj.Print
        Y = obj.CurrentY
        X = obj.CurrentX
        obj.FillStyle = vbFSTransparent
        FontGr 12
        obj.Line (X, Y)-(obj.ScaleWidth / 2 - 30, Y + obj.TextHeight("W") * 2.5), vbBlack, B
        Y = Y + 50
        obj.CurrentY = Y
        obj.FillStyle = vbFSTransparent
        kolwidth = (obj.ScaleWidth / 8) * 2
        sqlstr = "Select * from qryWeds where  ksid = " & kampID
        sqlstr = sqlstr & " and wedtype = 3"
        sqlstr = sqlstr & " ORDER BY wednum"
        rs.Close
        rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
        xpos = 0
        With rs
            If .RecordCount > 0 Then
                i = 0
                Do While Not .EOF
                    ypos = Y
                    FontGr 8
                    obj.CurrentX = xpos + 50
                    obj.CurrentY = ypos + obj.TextHeight("99") * 0.5
                    obj.Print Format(!wedNum, "0"); ":";
                    FontGr 12
                    obj.CurrentX = xpos + obj.TextWidth("00:") + 30
                    obj.CurrentY = ypos
                    FontGr 10
                    obj.Print !code1; ":";
                    FontGr 12
                    obj.DrawWidth = 1
                    obj.Line (xpos + obj.TextWidth("00:"), ypos)-(xpos + kolwidth - 50, ypos + obj.TextHeight("W")), vbBlack, B
                    ypos = obj.CurrentY
                    obj.CurrentX = xpos + obj.TextWidth("00:") + 30
                    FontGr 10
                    obj.Print !code2; ":";
                    FontGr 12
                    obj.Line (xpos + obj.TextWidth("00:"), ypos)-(xpos + kolwidth - 50, ypos + obj.TextHeight("W")), vbBlack, B
                    'wedstrijd nr
                    obj.CurrentY = ypos
                    .MoveNext
                    i = i + 1
                    xpos = kolwidth * i
                    If xpos > obj.ScaleWidth - kolwidth + 100 Then
                        FontGr 8
                        obj.Print
                        obj.Print
                        FontGr 12
                        Y = obj.CurrentY
                        i = 0
                        xpos = 0
                    End If
                Loop
            End If
        End With
    End If
    obj.CurrentY = hvFinYpos
    i = getPntToek("kleine finaleplaats") + getPntToek("kleine finalepositie")
    If i > 0 Then
        HeeftKlFin = True
        'print kleine finale
        txt = "("
        intP = getPntToek("kleine finaleplaats")
        If intP > 0 Then txt = txt & intP & ""
        intP = getPntToek("kleine finalepositie")
        If intP > 0 Then
            If txt > "(" Then txt = txt & "/"
            txt = txt & intP & " pnt"
        Else
            txt = txt & " pnt"
        End If
        txt = txt & ")"
        obj.Font = "Tahoma"
        Y = hvFinYpos
        obj.CurrentY = Y
        obj.CurrentX = obj.ScaleWidth / 2 + 30
        X = obj.CurrentX
        FontGr 14
        Vet True
        obj.FillColor = &H808080
        obj.FillStyle = vbFSSolid
        obj.Line (X, Y - 10)-(obj.ScaleWidth * 0.75, Y + obj.TextHeight("W")), vbBlack, B
        obj.CurrentY = Y
        obj.CurrentX = X + 50
        obj.ForeColor = vbWhite
        obj.Print "3e plaats ";
        FontGr 10
        Vet False
        obj.CurrentY = Y + 40
        obj.Print txt;
        obj.ForeColor = vbBlack
        obj.CurrentY = Y
        FontGr 14
        obj.Print
        obj.FillStyle = vbFSTransparent
        Y = obj.CurrentY
        X = obj.ScaleWidth / 2 + 30
        FontGr 12
        obj.Line (X, Y)-(obj.ScaleWidth * 0.75, Y + obj.TextHeight("W") * 2.5), vbBlack, B
        Y = Y + 50
        obj.CurrentY = Y
        obj.FillStyle = vbFSTransparent
        kolwidth = (obj.ScaleWidth / 8) * 2
        sqlstr = "Select * from qryWeds where  ksid = " & kampID
        sqlstr = sqlstr & " and wedtype = 7"
        sqlstr = sqlstr & " ORDER BY wednum"
        rs.Close
        rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
        xpos = obj.ScaleWidth / 2 + 30
        With rs
            If .RecordCount > 0 Then
                i = 0
                Do While Not .EOF
                    ypos = Y
                    FontGr 8
                    obj.CurrentX = xpos + 50
                    obj.CurrentY = ypos + obj.TextHeight("99") * 0.5
                    obj.Print Format(!wedNum, "0"); ":";
                    FontGr 12
                    obj.CurrentX = xpos + obj.TextWidth("00:") + 30
                    obj.CurrentY = ypos
                    obj.Print !code1; ":";
                    obj.DrawWidth = 1
                    obj.Line (xpos + obj.TextWidth("00:"), ypos)-(xpos + kolwidth - 50, ypos + obj.TextHeight("W")), vbBlack, B
                    ypos = obj.CurrentY
                    obj.CurrentX = xpos + obj.TextWidth("00:") + 30
                    obj.Print !code2; ":";
                    obj.Line (xpos + obj.TextWidth("00:"), ypos)-(xpos + kolwidth - 50, ypos + obj.TextHeight("W")), vbBlack, B
                    'wedstrijd nr
                    obj.CurrentY = ypos
                    .MoveNext
                    i = i + 1
                    xpos = kolwidth * i
                    If xpos > obj.ScaleWidth - kolwidth + 100 Then
                        FontGr 8
                        obj.Print
                        obj.Print
                        FontGr 12
                        Y = obj.CurrentY
                        i = 0
                        xpos = 0
                    End If
                Loop
            End If
        End With
    Else
        HeeftKlFin = False
    End If
    obj.CurrentY = hvFinYpos
    i = getPntToek("finaleplaats") + getPntToek("finalepositie")
    If i > 0 Then
        'print finale
        txt = "("
        intP = getPntToek("finaleplaats")
        If intP > 0 Then txt = txt & intP
        intP = getPntToek("finalepositie")
        If intP > 0 Then
            If txt > "(" Then txt = txt & "/"
            txt = txt & intP & " pnt"
        Else
            txt = txt & " pnt"
        End If
        txt = txt & ")"
        obj.Font = "Tahoma"
        Y = hvFinYpos
        obj.CurrentY = Y
        If HeeftKlFin Then
            obj.CurrentX = obj.ScaleWidth * 0.75 + 30
        Else
            obj.CurrentX = obj.ScaleWidth * 0.5 + 30
        End If
        X = obj.CurrentX
        FontGr 14
        Vet True
        obj.FillColor = &H808080
        obj.FillStyle = vbFSSolid
        obj.Line (X, Y - 10)-(obj.ScaleWidth, Y + obj.TextHeight("W")), vbBlack, B
        obj.ForeColor = vbWhite
        obj.CurrentY = Y
        obj.CurrentX = X + 50
        obj.Print "Finale ";
        FontGr 10
        Vet False
        obj.CurrentY = Y + 40
        obj.Print txt;
        obj.ForeColor = vbBlack
        obj.CurrentY = Y
        FontGr 14
        obj.Print
        obj.FillStyle = vbFSTransparent
        Y = obj.CurrentY
        If HeeftKlFin Then
            X = obj.ScaleWidth * 0.75 + 30
        Else
            X = obj.ScaleWidth * 0.5 + 30
        End If
        FontGr 12
        obj.Line (X, Y)-(obj.ScaleWidth, Y + obj.TextHeight("W") * 2.5), vbBlack, B
        Y = Y + 50
        obj.CurrentY = Y
        obj.FillStyle = vbFSTransparent
        kolwidth = (obj.ScaleWidth / 8) * 2
        sqlstr = "Select * from qryWeds where  ksid = " & kampID
        sqlstr = sqlstr & " and wedtype = 4"
        sqlstr = sqlstr & " ORDER BY wednum"
        rs.Close
        rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
        If HeeftKlFin Then
            xpos = obj.ScaleWidth * 0.75 + 30
        Else
            xpos = obj.ScaleWidth * 0.5 + 30
            kolwidth = kolwidth * 2
        End If
        With rs
            If .RecordCount > 0 Then
                i = 0
                Do While Not .EOF
                    ypos = Y
                    FontGr 8
                    obj.CurrentX = xpos + 50
                    obj.CurrentY = ypos + obj.TextHeight("99") * 0.5
                    'wedstrijd nr
                    obj.Print Format(!wedNum, "0"); ":";
                    FontGr 12
                    obj.CurrentX = xpos + obj.TextWidth("00:") + 30
                    obj.CurrentY = ypos
                    FontGr 10
                    obj.Print !code1; ":";
                    FontGr 12
                    obj.DrawWidth = 1
                    obj.Line (xpos + obj.TextWidth("00:"), ypos)-(xpos + kolwidth - 50, ypos + obj.TextHeight("W")), vbBlack, B
                    ypos = obj.CurrentY
                    obj.CurrentX = xpos + obj.TextWidth("00:") + 30
                    FontGr 10
                    obj.Print !code2; ":";
                    FontGr 12
                    obj.Line (xpos + obj.TextWidth("00:"), ypos)-(xpos + kolwidth - 50, ypos + obj.TextHeight("W")), vbBlack, B
                    obj.CurrentY = ypos
                    .MoveNext
                    i = i + 1
                    xpos = kolwidth * i
                    If xpos > obj.ScaleWidth - kolwidth + 100 Then
                        FontGr 8
                        obj.Print
                        obj.Print
                        FontGr 12
                        Y = obj.CurrentY
                        i = 0
                        xpos = 0
                    End If
                Loop
            End If
            .Close
        End With
        Set rs = Nothing
    End If
    FontGr 8
    obj.Print
    FontGr 12
End Sub

Sub printgroep(nr As Integer)
Dim rs As New ADODB.Recordset
Dim sqlstr As String
Dim xLinePos As Integer
Dim yLinePos As Integer
Dim xpos As Integer
Dim txt As String
Dim vakPos(1, 1)
Dim grp As String * 1
Dim iGrp As Integer
grp = Chr(nr + 64)
FontGr 10
sqlstr = "Select * from qrygroepteams where ksid=" & kampID
sqlstr = sqlstr & " AND groep = '" & grp & "'"
sqlstr = sqlstr & " ORDER BY groep, plaatsing"
rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
yLinePos = obj.CurrentY
iGrp = getTournamentInfo("groepen")
xLinePos = (obj.ScaleWidth / iGrp) * (nr - 1)
xpos = xLinePos + 50
Do While Not rs.EOF
    vakPos(0, 0) = xpos + obj.ScaleWidth / iGrp - obj.TextHeight("W") - obj.TextWidth("W")
    vakPos(0, 1) = obj.CurrentY
    vakPos(1, 0) = vakPos(0, 0) + obj.TextHeight("W")
    vakPos(1, 1) = vakPos(0, 1) + obj.TextHeight("W")
    
    txt = rs!naam
    Do While xpos + obj.TextWidth(txt) > vakPos(0, 0)
        txt = Left(txt, Len(txt) - 1)
    Loop
    obj.CurrentX = xpos
    obj.Print txt;
    obj.FillStyle = vbFSTransparent
    obj.FillColor = vbWhite
    obj.DrawWidth = 1
    
    obj.Line (vakPos(0, 0), vakPos(0, 1))-(vakPos(1, 0), vakPos(1, 1)), vbBlack, B
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
'obj.CurrentY = yLinePos
End Sub

Sub formulierWeds()
'wedstrijden op het poolformulier
Dim fontBas As Integer
Dim rs As New ADODB.Recordset
Dim sqlstr As String
Dim posWednr As Integer
Dim posDatum As Integer
Dim posTijd As Integer
Dim posWedOms As Integer
Dim posRust As Integer
Dim PosEind As Integer
Dim posToto As Integer
Dim wedOms As String
Dim kolwidth As Integer
Dim kolom As Integer
Dim ypos As Integer
Dim curYpos As Integer
Dim X As Integer
Dim Y As Integer
Dim i As Integer
Dim vertLineYPos As Integer
Dim vertLineYPos2 As Integer
Dim topY As String
Dim savdat As Date
Dim vertLineEndPos As Integer
    sqlstr = "Select * from qryweds where ksid = " & kampID
    sqlstr = sqlstr & " ORDER BY datum,tijd,wednum"
    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    If rs.RecordCount = 0 Then
        rs.Close
        Exit Sub
    End If
    fontBas = 10
    FontGr fontBas + 2
    topY = obj.CurrentY
    obj.CurrentY = voethoog - GrootHoog
    ypos = obj.CurrentY
    obj.FillColor = &H808080
    obj.FillStyle = vbFSSolid
    obj.Line (0, ypos)-(obj.ScaleWidth + 2 * obj.ScaleLeft, voethoog), vbBlack, B
    obj.CurrentY = ypos + 30
    FontGr 16
    Vet True
    obj.ForeColor = vbWhite
    iBKMode = SetBkMode(obj.hdc, TRANSPARENT)
    Centreer "UITERLIJK INLEVEREN OP " & UCase(Format(getPoolInfo("eindinschr"), "dddd d mmmm yyyy"))
    obj.ForeColor = vbBlack
    obj.FillStyle = vbFSTransparent
    Vet False
    FontGr fontBas + 2
    obj.CurrentY = topY
    kolom = 0
    kolwidth = obj.ScaleWidth / 2 - obj.TextWidth("w")
    obj.FontName = "Times New Roman"
    FontGr 2
    obj.Print
    FontGr fontBas + 2
    obj.CurrentY = obj.CurrentY + 20
    FontGr fontBas + 4
    Vet True
    obj.Print "Uitleg"
    FontGr fontBas + 2
    Vet False
    obj.Print "Vul hieronder voor alle wedstrijden jouw uitslagen in. ";
    Vet True
    obj.Print "Ook daar waar de teams nog niet bekend zijn."
    Vet False
    obj.Print "(Ook al heb je een ander team op die plaats dan kan je uitslag nog steeds goed zijn)"
    obj.Print "De uitslag hoeft onderling niet te kloppen. ";
    obj.Print "Je krijgt punten voor elk vak dat achteraf juist blijkt te zijn ingevuld."
    obj.Print "Bij 'toto' vul je een 1 in voor winst linker team, een 2 voor winst rechter team en een 3 voor een gelijkspel"
    Vet True
    Centreer "Alle uitslagen, ook de toto, gelden na 90 minuten voetbal!"
    Vet False
    FontGr fontBas
    obj.Print
    Centreer "(plus de eventuele blessuretijd)"
    obj.Print
    FontGr fontBas + 4
    Vet True
    obj.Print "Punten"
    Vet False
    FontGr fontBas + 2
    obj.Print "Ruststand goed: ";
    Vet True
    obj.Print getPntToek("ruststand goed"); "pnt, ";
    Vet False
    obj.Print "Eindstand goed: ";
    Vet True
    obj.Print getPntToek("eindstand goed"); "pnt, ";
    Vet False
    obj.Print "Toto goed: ";
    Vet True
    obj.Print getPntToek("toto goed"); "pnt.";
    Vet False
    If getPntToek("doelpunten op een dag") > 0 Then
        obj.Print "Totaal aantal doelpunten op één dag goed: ";
        Vet True
        obj.Print getPntToek("doelpunten op een dag"); " pnt"
        Vet False
    End If
    obj.Print
    FontGr fontBas
    posDatum = 50
    posTijd = posDatum + obj.TextWidth("MA 26-6") + 10
    posWednr = posTijd + obj.TextWidth("00:000") + 10
    posWedOms = posWednr + obj.TextWidth("199:")
    posRust = posWedOms + obj.TextWidth("Nederland - Zwitserland")
    PosEind = posRust + obj.TextWidth("123456")
    posToto = PosEind + obj.TextWidth("123456")
    
    vertLineYPos = obj.CurrentY
    FontGr fontBas
    obj.Line (0, vertLineYPos - 20)-(kolwidth * 2, vertLineYPos - 20)
    obj.CurrentY = vertLineYPos
    For i = 0 To 1
        obj.CurrentX = posDatum + i * kolwidth
        obj.Print " Datum";
        obj.CurrentX = posTijd + i * kolwidth
        obj.Print " tijd";
        obj.CurrentX = posWednr + i * kolwidth
        obj.Print " nr";
        obj.CurrentX = posWedOms + i * kolwidth
        obj.Print " Wedstrijd";
        obj.CurrentX = posRust + i * kolwidth
        obj.Print " rust";
        obj.CurrentX = PosEind + i * kolwidth
        obj.Print " eind";
        obj.CurrentX = posToto + i * kolwidth
        obj.Print " toto";
    Next
    obj.Print
    obj.Line (0, obj.CurrentY)-(kolwidth * 2, obj.CurrentY), 1
    vertLineYPos2 = obj.CurrentY
    
    ypos = obj.CurrentY
    
    With rs
        .MoveLast
        .MoveFirst
        
        Do While Not .EOF
            If (nz(!naam1, "")) > "" Then
                wedOms = !code1 & ":" & !naam1 & " - " & !code2 & ":" & !naam2
            Else
                wedOms = !code1 & " - " & !code2
            End If
            
            obj.CurrentY = obj.CurrentY + 40
            obj.CurrentX = posWednr + kolom * kolwidth + (posWedOms - posWednr - obj.TextWidth(Format(!wedNum, "0"))) / 2
            obj.Print Format(!wedNum, "0");
            obj.CurrentX = posDatum + kolom * kolwidth
            If savdat <> !Datum Then
                obj.Print Format(!Datum, "ddd d-M"); " ";
                savdat = !Datum
            End If
            obj.CurrentX = posTijd + kolom * kolwidth + (posWednr - posTijd - obj.TextWidth(Format(!tijd, "HH:NN"))) / 2
            obj.Print tijdFormat(!tijd); '  , "HH:NN");
            obj.CurrentX = posWedOms + kolom * kolwidth + 30
            curYpos = obj.CurrentY
            If (nz(!naam1, "")) > "" Then
                FontGr fontBas - 3
                obj.CurrentY = curYpos + 20
                Do While obj.TextWidth(wedOms) > posRust - posWedOms
                    wedOms = Left(wedOms, Len(wedOms) - 1)
                Loop
            Else
                FontGr fontBas
                obj.CurrentY = curYpos
            End If
            obj.Print wedOms;
            obj.CurrentY = curYpos
            FontGr fontBas
            X = posRust + kolom * kolwidth
            Y = obj.CurrentY - 20
            obj.Line (X, Y)-(PosEind + kolom * kolwidth - 10, Y + obj.TextHeight("W") + 50), , B
            obj.CurrentX = posRust + (PosEind - posRust - obj.TextWidth("-")) / 2 + kolom * kolwidth
            obj.CurrentY = Y + 30
            obj.Print "-";
            X = PosEind + kolom * kolwidth + 10
            obj.Line (X, Y)-(posToto + kolom * kolwidth - 10, Y + obj.TextHeight("W") + 50), , B
            obj.CurrentX = PosEind + (posToto - PosEind - obj.TextWidth("-")) / 2 + kolom * kolwidth
            obj.CurrentY = Y + 30
            obj.Print "-";
            X = posToto + kolom * kolwidth + 10
            obj.Line (X, Y)-(kolwidth * (kolom + 1) - obj.TextWidth("0"), Y + obj.TextHeight("W") + 50), , B
            obj.CurrentX = PosEind + (posToto - PosEind - obj.TextWidth("-")) / 2
            obj.CurrentY = Y
            
            FontGr 14
            obj.Print
            FontGr fontBas
            obj.Line (0, obj.CurrentY)-(kolwidth * 2, obj.CurrentY), 1

            .MoveNext
            If (.AbsolutePosition - 1) = Int(rs.RecordCount / 2 + 0.5) Then
                kolom = 1
                vertLineEndPos = obj.CurrentY
                obj.CurrentY = ypos
            End If
        Loop
        .Close
    End With
    Set rs = Nothing
    For i = 0 To 1
        obj.Line (0 + kolwidth * i, vertLineYPos - 10)-(0 + kolwidth * i, vertLineEndPos)
        obj.Line (posWednr + kolwidth * i - 10, vertLineYPos2)-(posWednr + kolwidth * i - 10, vertLineEndPos)
        obj.Line (posTijd + kolwidth * i, vertLineYPos2)-(posTijd + kolwidth * i, vertLineEndPos)
        obj.Line (posWedOms + kolwidth * i - 10, vertLineYPos2)-(posWedOms + kolwidth * i - 10, vertLineEndPos)
    Next
    obj.Line (kolwidth - 50, vertLineYPos - 10)-(kolwidth - 50, vertLineEndPos)
    obj.Line (kolwidth * 2, vertLineYPos - 10)-(kolwidth * 2, vertLineEndPos)
End Sub


Private Sub PrijsAfdr(wat As String, eind As Boolean)
Dim aant As Integer
Dim i As Integer
End Sub

Private Sub Centreer(Tekst$)
    obj.CurrentX = (obj.ScaleWidth - obj.TextWidth(Trim$(Tekst$))) \ 2
    obj.Print Tekst$;
End Sub

Function sqlDeelnems(poule As Long) As String
Dim sqlstr As String
    sqlstr = "Select * from pooldeelnems"
    sqlstr = sqlstr & " WHERE PoolID = " & poule
    sqlstr = sqlstr & " ORDER BY bijnaam "
    sqlDeelnems = sqlstr
End Function

Private Sub Favorieten()
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim aantgroep As Integer
Dim i As Integer
Dim j As Integer
Dim aant As Integer
Dim savX As Integer
Dim savy As Integer
Dim xpos As Integer
Dim col(4) As Integer
Dim yStart As Integer
Dim maxrows As Integer
Dim bewYPos As Integer
Dim deelnAant As Integer
Dim fntGr As Double
Dim sqlstr As String

deelnAant = GetDeelnemAant(thisPool)
headerText = GetOrgNaam(thisPool) & " " & getTournamentInfo("toernooi") & " voetbalpool" & " - Favorieten" & " (" & GetDeelnemAant(thisPool) & " deelnemers)"
'obj.Line (0, obj.CurrentY)-(obj.ScaleWidth, obj.CurrentY)
kop$ = "Groepstanden"
InitPage False, False
'intro
yStart = obj.CurrentY

'groepen
fntGr = obj.Font.Size
sqlstr = "Select groepen from ks WHERE id = " & kampID
rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
aantgroep = rs!groepen
rs.Close
obj.CurrentX = obj.TextWidth("12345678901234567890123456")
For i = 1 To 4
    obj.CurrentX = obj.CurrentX - obj.TextWidth(Format(i, "0") & "e pl")
    obj.Print Format(i, "0"); "e pl";
    col(i) = obj.CurrentX - 50
    obj.CurrentX = obj.CurrentX + obj.TextWidth("123456")
Next
obj.CurrentX = obj.ScaleWidth / 2 + obj.TextWidth("12345678901234567890123456")
For i = 1 To 4
    obj.CurrentX = obj.CurrentX - obj.TextWidth(Format(i, "0") & "e pl")
    obj.Print Format(i, "0"); "e pl";
    obj.CurrentX = obj.CurrentX + obj.TextWidth("123456")
Next
obj.CurrentX = 0
obj.Print
xpos = 0
savy = obj.CurrentY
For i = 1 To aantgroep
    If i = aantgroep / 2 + 1 Then
        xpos = obj.ScaleWidth / 2
        obj.CurrentY = savy
    End If
    sqlstr = "Select * from groepsindeling where ksid = " & kampID
    sqlstr = sqlstr & " AND groep = '" & Chr(i + 64) & "'"
    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    rs.MoveFirst
    obj.CurrentX = xpos
    obj.Print "Groep " & rs!groep; ": ";
    savX = obj.CurrentX
    Do While Not rs.EOF
        obj.CurrentX = savX
        obj.Print GetTeam(rs!team); " ";
        obj.CurrentX = obj.TextWidth("12345678901234567890")
        For j = 1 To 4
            aant = getAantalGrpVoorsp(j, rs!team)
            FontGr 9
            obj.CurrentY = obj.CurrentY + 30
            obj.CurrentX = xpos + col(j) - obj.TextWidth(Format(aant / deelnAant, "0.0%"))
'            obj.Print aant;
'            FontGr 8
            obj.Print Format(aant / deelnAant, "0.0%");
            obj.CurrentY = obj.CurrentY - 30
            FontGr CInt(fntGr)
            'If j < 4 Then obj.Print ", ";
        Next
        obj.Print
        rs.MoveNext
    Loop
    rs.Close
Next
savy = obj.CurrentY
On Error Resume Next
obj.Line (0, yStart)-(obj.ScaleWidth - 50, savy), , B
On Error GoTo 0
maxY = savy
'achtste finales
i = getPntToek("achtste finaleplaats") + getPntToek("achtste finalepositie")
If i > 0 Then
    Fav_Finals 5, 4, "Achtste finales"
    savy = obj.CurrentY
End If
obj.CurrentY = savy
'kwart finales
i = getPntToek("kwart finaleplaats") + getPntToek("kwart finalepositie")
If i > 0 Then
    Fav_Finals 2, 4, "Kwart finales"
    savy = obj.CurrentY
End If
obj.CurrentY = savy
'halve finales
i = getPntToek("halve finaleplaats") + getPntToek("halve finalepositie")
If i > 0 Then
    Fav_Finals 3, 4, "Halve finales"
    savy = obj.CurrentY
    maxY = savy
End If
obj.CurrentY = savy
'kleine finale
i = getPntToek("kleine finaleplaats") + getPntToek("kleine finalepositie")
If i > 0 Then
    bewYPos = obj.CurrentY
    Fav_Finals 7, 4, "Kleine finale"
    savy = maxY
    'maxY = savy
    savX = 3
Else
    bewYPos = obj.CurrentY
    savX = 1
End If

'finale
i = getPntToek("finaleplaats") + getPntToek("finalepositie")
If i > 0 Then
    Fav_Finals 4, 4, "Finale", savy, savX
    If savX = 3 Then
        savX = 1
        savy = obj.CurrentY
    Else
        savy = bewYPos
        savX = 3
    End If
'    savy = obj.CurrentY
    maxY = savy
End If
obj.CurrentY = savy
Fav_Eindstand savy, savX
Fav_Topscorers
Set rs = Nothing
obj.Print
obj.Print
End Sub

Sub Fav_Topscorers()
Dim aant As Integer
Dim cols(5) As Integer
Dim sqlstr As String
Dim savy As Integer
Dim savFntgr As Integer
Dim rs As New ADODB.Recordset
Dim i As Integer
Dim j As Integer
For i = 1 To 4
    cols(i) = Int(obj.ScaleWidth / 4) * (i - 1)
Next
cols(5) = obj.ScaleWidth - 10
sqlstr = "SELECT personen.rnaam, Count(voorspelling_ts.deelnem) AS aantal"
sqlstr = sqlstr & " FROM voorspelling_ts LEFT JOIN personen ON voorspelling_ts.ts = personen.ID"
sqlstr = sqlstr & " WHERE voorspelling_ts.deelnem In (select deelnemid from pooldeelnems where poolid= " & thisPool
sqlstr = sqlstr & " ) GROUP BY personen.rnaam, voorspelling_ts.ts"
sqlstr = sqlstr & " ORDER BY Count(voorspelling_ts.deelnem) DESC, personen.rnaam "
rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
If rs.RecordCount > 0 Then
    rs.MoveLast
End If
aant = rs.RecordCount
i = 1
j = 0

obj.CurrentX = favXpos
If favYpos > voethoog - Int(aant / 4) * obj.TextHeight("tekst") - 120 Then
  kop$ = "Topscorers"
  DoNewPage False, False, 0
  favYpos = obj.CurrentY
Else
  obj.CurrentY = favYpos
  koptekst "Topscorers", False, False, favYpos, 0
End If

savy = obj.CurrentY
rs.MoveFirst

Do While Not rs.EOF
    obj.CurrentX = cols(i)
    If nz(rs!rnaam, "") > "" Then
        obj.Print rs!rnaam;
    Else
        obj.Print "Niet ingevuld";
    End If
    obj.CurrentX = cols(i + 1) - 500 - obj.TextWidth(rs!Aantal)
    obj.Print rs!Aantal
    j = j + 1
    rs.MoveNext
    If obj.CurrentY > favYpos Then
        favYpos = obj.CurrentY
    End If
    If j > Int(aant / 4) - 1 Then
        i = i + 1
        j = 0
        obj.CurrentY = savy
    End If
Loop
rs.Close
Set rs = Nothing
obj.Line (cols(1), savy)-(cols(5) - 50, favYpos), , B

End Sub

Function GetRijAant(wedNum As Integer, team)
'om te bepalen of we naar een nieuw pagina moeten in de favorieten afdruk
Dim sqlstr As String
sqlstr = "SELECT wed, " & team
sqlstr = sqlstr & " From voorspelling_finales"
sqlstr = sqlstr & " WHERE deelnem In (select deelnemid from pooldeelnems where poolid =" & thisPool
sqlstr = sqlstr & " ) GROUP BY wed, " & team
sqlstr = sqlstr & " HAVING wed =" & wedNum
sqlstr = sqlstr & " AND " & team & " >0"
Dim rs As New ADODB.Recordset
rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
If Not rs.RecordCount = 0 Then
    rs.MoveLast
End If
GetRijAant = rs.RecordCount
rs.Close
Set rs = Nothing
End Function

Sub PrintEindStandFav(Plaats As String, col As Integer, rs As ADODB.Recordset, veld As String)
Dim sqlstr As String
Dim ypos As Integer
Dim fntGr As Integer
    ypos = obj.CurrentY
    fntGr = obj.Font.Size
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Vet True
        obj.CurrentX = col
        obj.Print Plaats
        Vet False
        Do While Not rs.EOF
            obj.CurrentX = col + 50
            If nz(rs(veld), 0) = 0 Then
                obj.Print "Niet ingevuld";
            Else
                obj.Print GetTeam(rs(veld));
            End If
            obj.CurrentX = col + obj.TextWidth("123456789012345") - obj.TextWidth(rs!Aantal)
            obj.Print rs!Aantal;
            FontGr fntGr - 3
            obj.CurrentY = obj.CurrentY + 30
            obj.Print "(" & Format(rs!Aantal / GetDeelnemAant(thisPool), "0.0%") & ")"
            obj.CurrentY = obj.CurrentY - 30
            FontGr fntGr
            rs.MoveNext
        Loop
    End If
End Sub
Sub Fav_Eindstand(savy As Integer, savX2 As Integer)
Dim sqlstr As String
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim maxaant As Integer
Dim savX As Integer
Dim aantpos As Integer
Dim startY As Integer
Dim maxY As Integer
Dim i As Integer
Dim savFntgr As Integer
Dim aantFav As Integer
Dim cols(5) As Integer
For i = 1 To 4
    cols(i) = Int((obj.ScaleWidth / 4) * (i - 1))
Next

cols(5) = obj.ScaleWidth - 20

    startY = savy

    sqlstr = "SELECT kampioen, Count(pooldeelnems.deelnemID) AS aantal"
    sqlstr = sqlstr & " From pooldeelnems"
    sqlstr = sqlstr & " WHERE poolid = " & thisPool
    sqlstr = sqlstr & " GROUP BY kampioen"
    sqlstr = sqlstr & " ORDER BY count(deelnemID) desc"
    rs1.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    sqlstr = "SELECT pltwee, Count(pooldeelnems.deelnemID) AS aantal"
    sqlstr = sqlstr & " From pooldeelnems"
    sqlstr = sqlstr & " WHERE poolid = " & thisPool
    sqlstr = sqlstr & " GROUP BY pltwee"
    sqlstr = sqlstr & " ORDER BY count(deelnemID) desc"
    rs2.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    sqlstr = "SELECT pldrie, Count(pooldeelnems.deelnemID) AS aantal"
    sqlstr = sqlstr & " From pooldeelnems"
    sqlstr = sqlstr & " WHERE poolid = " & thisPool
    sqlstr = sqlstr & " GROUP BY pldrie"
    sqlstr = sqlstr & " ORDER BY count(deelnemID) desc"
    rs3.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    sqlstr = "SELECT plvier, Count(pooldeelnems.deelnemID) AS aantal"
    sqlstr = sqlstr & " From pooldeelnems"
    sqlstr = sqlstr & " WHERE poolid = " & thisPool
    sqlstr = sqlstr & " GROUP BY plvier"
    sqlstr = sqlstr & " ORDER BY count(deelnemID) desc"
    rs4.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    If rs1.RecordCount > 0 Then
        rs1.MoveLast
    End If
    If savX2 = 1 Then
        aantFav = 3
    Else
        aantFav = 3
    End If
    
    favXpos = cols(savX2)
    maxaant = rs1.RecordCount
    If rs2.RecordCount > 0 Then
        rs2.MoveLast
        If Not IsNull(rs2!pltwee) Then
            aantFav = aantFav + 1
            favXpos = cols(aantFav + 1)
        End If
    End If
    If rs2.RecordCount > maxaant Then
        maxaant = rs2.RecordCount
    End If
    If rs3.RecordCount > 0 Then
        rs3.MoveLast
        If Not IsNull(rs3!pldrie) Then
            aantFav = 3
            favXpos = cols(aantFav + 1)
        End If
    End If
    If rs3.RecordCount > maxaant Then
        maxaant = rs3.RecordCount
    End If
    If rs4.RecordCount > 0 Then
        rs4.MoveLast
        If Not IsNull(rs4!plvier) Then
            aantFav = 0
            favXpos = cols(1)
        End If
    End If
    If rs4.RecordCount > maxaant Then
        maxaant = rs4.RecordCount
    End If
    savFntgr = obj.FontSize
    obj.FontSize = savFntgr - 3
    maxY = maxaant * obj.TextHeight("Q") + savy
    obj.FontSize = savFntgr
    maxY = maxY + obj.TextHeight("Q") + 50
    If maxY > voethoog - 465 Then
        kop$ = "Favorieten einduitslag"
        DoNewPage False, False, aantFav
        'maxY = obj.CurrentY
        savy = obj.CurrentY
        startY = savy
        savFntgr = obj.FontSize
        obj.FontSize = savFntgr - 3
        maxY = maxaant * obj.TextHeight("Q") + savy
        obj.FontSize = savFntgr
        maxY = maxY + obj.TextHeight("Q") + 50
    Else
      If savX2 = 3 Then
        koptekst "Favorieten einduitslag", False, False, savy, savX2 + 1
      Else
        koptekst "Favorieten einduitslag", False, False, savy, savX2 - 1 ' 0 centreert tussenkop
      End If
      savy = obj.CurrentY
      startY = savy
      savFntgr = obj.FontSize
      obj.FontSize = savFntgr - 3
      maxY = maxaant * obj.TextHeight("Q") + savy
      obj.FontSize = savFntgr
      maxY = maxY + obj.TextHeight("Q") + 50
    End If
    If getPntToek("1e plaats(Kampioen)") Then
        obj.CurrentY = savy
        PrintEindStandFav "kampioen", cols(savX2) + 10, rs1, "kampioen"
        obj.Line (cols(savX2), startY)-(cols(savX2 + 1) - 50, maxY), , B
    End If
    If getPntToek("2e plaats") Then
        obj.CurrentY = savy
        PrintEindStandFav "2e plaats", cols(savX2 + 1) + 10, rs2, "plTwee"
        obj.Line (cols(savX2 + 1), startY)-(cols(savX2 + 2) - 50, maxY), , B
    End If
    If getPntToek("3e plaats") Then
        obj.CurrentY = savy
        PrintEindStandFav "3e plaats", obj.ScaleWidth / 2 + 10, rs3, "pldrie"
        obj.Line (cols(3), startY)-(cols(4) - 50, maxY), , B
    End If
    If getPntToek("4e plaats") Then
        obj.CurrentY = savy
        PrintEindStandFav "4e plaats", (obj.ScaleWidth / 4) * 3 + 10, rs4, "plvier"
        obj.Line (cols(4), startY)-(cols(5) - 50, maxY), , B
    End If
    favYpos = maxY
    favXpos = 0
End Sub
Sub Fav_Finals(wedtype As Integer, cols As Integer, koptxt As String, Optional bewaarYpos As Integer, Optional posX As Integer)
Dim sqlstr As String
Dim rs As New ADODB.Recordset
Dim savX As Integer
Dim savy As Integer
Dim aantpos As Integer
Dim startY As Integer
Dim col() As Integer
Dim i As Integer
Dim j As Integer
Dim team As String
Dim fld As field
Dim maxrows As Integer
Dim maxrows1 As Integer
Dim savMaxRows As Integer
Dim savMaxRows1 As Integer
Dim ttlRows As Integer
Dim maxFinpos As Integer
ReDim col(cols + 1) As Integer
    For i = 1 To cols
        col(i) = (i - 1) * obj.ScaleWidth / cols
    Next
    col(cols + 1) = obj.ScaleWidth
    savy = obj.CurrentY
    sqlstr = "Select * from qryWeds where  ksid = " & kampID
    sqlstr = sqlstr & " and wedtype = " & wedtype
    sqlstr = sqlstr & " ORDER BY wednum"
    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    startY = savy
    'startY = 945
    
    If rs.RecordCount > 0 Then
        savMaxRows = 0
        maxrows = 0
        rs.MoveFirst
        'bepaal aantal regels dat nodig is
        Do While Not rs.EOF
            savMaxRows = maxrows + GetRijAant(rs!wedNum, "t1")
            If maxrows < savMaxRows Then
                maxrows = savMaxRows
            End If
            rs.MoveNext
        Loop
        rs.MoveFirst
        'bepaal aantal regels dat nodig is
        Do While Not rs.EOF
            savMaxRows1 = maxrows1 + GetRijAant(rs!wedNum, "t2")
            If maxrows1 < savMaxRows1 Then
                maxrows1 = savMaxRows1
            End If
            rs.MoveNext
        Loop
        ttlRows = maxrows
        If maxrows1 > ttlRows Then ttlRows = maxrows1
        rs.MoveFirst
        If startY + ttlRows * TextHeight("Q") > voethoog - 465 And wedtype <> 4 Then '(465 = hoogte van het tussenkopje)
            kop$ = koptxt
            If wedtype = 7 Then
                DoNewPage False, False, 2
                maxY = obj.CurrentY
                savy = maxY
                startY = 480
                nwPag = True
            Else
                DoNewPage False, False
                maxY = obj.CurrentY
                savy = maxY
                startY = savy
                nwPag = False
            End If
        Else
            If wedtype = klFinale Then
                finYpos = obj.CurrentY
                koptekst koptxt, False, False, maxY, 2
            ElseIf wedtype = Finale Then
                If getPntToek("kleine finaleplaats") + getPntToek("kleine finalepositie") > 0 Then
                    If nwPag Then
                        koptekst koptxt, False, False, 480, 4
                    Else
                        koptekst koptxt, False, False, finYpos, 4
                    End If
                Else
                    koptekst koptxt, False, False, bewaarYpos, 2
                End If
            Else
                koptekst koptxt, False, False, maxY
            End If
            savy = obj.CurrentY
            startY = savy
        End If
        
        i = 1
        If wedtype = Finale Then
            i = posX
        End If
        'If wedtype = 7 Then Stop
        Do While Not rs.EOF
            If i <= cols Then
                obj.CurrentY = savy
            End If
            fav_finalTeams "t1", "code1", rs, col(i)
            If maxY < obj.CurrentY Then maxY = obj.CurrentY
            i = i + 1
            If i <= cols Then
                obj.CurrentY = savy
            End If
            fav_finalTeams "t2", "code2", rs, col(i)
            If maxY < obj.CurrentY Then maxY = obj.CurrentY
            i = i + 1
            
            If wedtype = 7 And maxY < obj.CurrentY Then
                maxY = obj.CurrentY
            ElseIf wedtype = 4 Then
                If obj.CurrentY > maxY Then
                    maxY = obj.CurrentY
                End If
            End If
            maxY = maxY + 50
            If i = 5 Then
                obj.Line (col(1), startY)-(col(3) - 50, maxY), , B
                obj.Line (col(3), startY)-(col(5) - 50, maxY), , B
            End If
            If posX = 1 And i = 3 Then
                obj.Line (col(1), startY)-(col(3) - 50, maxY), , B
            End If
            
            rs.MoveNext
            If i > cols Then
                i = 1
                obj.CurrentY = maxY + 50
                savy = obj.CurrentY
                maxY = savy
                startY = maxY
                favYpos = savy
                favXpos = 0
            End If
            
        Loop
        
    End If
    rs.Close
    Set rs = Nothing
End Sub

Sub fav_finalTeams(team As String, cod As String, rs As ADODB.Recordset, col)
Dim rs1 As New ADODB.Recordset
Dim savX As Integer
Dim savy As Integer
Dim aantpos As Integer
Dim sqlstr As String
Dim fntGr As Integer
    aantpos = obj.TextWidth("NIET INGEVULD  1")
    sqlstr = "SELECT wed, " & team & ", Count(wed) AS ttl From voorspelling_finales"
    sqlstr = sqlstr & " WHERE deelnem In (select deelnemid from pooldeelnems where poolid =" & thisPool
    sqlstr = sqlstr & " ) GROUP BY wed, " & team
    sqlstr = sqlstr & " HAVING wed=" & rs!wedNum
    sqlstr = sqlstr & " AND " & team & " > 0"
    sqlstr = sqlstr & " ORDER BY count(wed) desc"
    rs1.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    obj.CurrentX = col
    obj.Print rs(cod) & ": ";
    savX = obj.CurrentX
    fntGr = obj.Font.Size
    Do While Not rs1.EOF
        obj.CurrentX = savX
        If nz(rs1(team), "") = "" Then
            obj.Print "Niet ingevuld";
        Else
            obj.Print GetTeam(rs1(team));
        End If
        obj.CurrentX = col + aantpos - obj.TextWidth(rs1!ttl)
        obj.Print rs1!ttl;
        FontGr fntGr - 3
        obj.CurrentY = obj.CurrentY + 30
        obj.Print "(" & Format(rs1!ttl / GetDeelnemAant(thisPool), "0.0%") & ")"
        FontGr fntGr
        obj.CurrentY = obj.CurrentY - 30
        If maxY < obj.CurrentY Then maxY = obj.CurrentY
        rs1.MoveNext
    Loop
    rs1.Close
End Sub

Private Sub Deelnems()
Dim Dezedeeln As Integer
Dim tkst$
Dim tmpnaam$
Dim KolomAant As Integer
Dim i As Integer
Dim k As Integer
Dim LineXpos As Integer
Dim LineYPos As Integer
Dim newlinepos As Integer
Dim TopMarg As Integer
Dim pr As String
Dim rsDeelnem As New ADODB.Recordset
Dim rsDeelnemWeds As New ADODB.Recordset
Dim rsDeelnGroepen As New ADODB.Recordset
Dim rsDeelnFinales As New ADODB.Recordset
Dim rsDeelnts As New ADODB.Recordset
Dim rsDeelnEindstand As New ADODB.Recordset
Dim rsDeelnOverig As New ADODB.Recordset
Dim sqlstr As String
Dim naamHeight As Integer
Dim wedHoog As Integer
Dim NaamHoog As Integer
Dim posDatum As Integer
Dim posTijd As Integer
Dim posWedOms As Integer
Dim posRust As Integer
Dim PosEind As Integer
Dim posToto As Integer
Dim posPnt As Integer
Dim wedYpos As Integer
Dim wedKol As Integer
Dim Helft As Integer
Dim oldhelft As Integer
Dim heeft8stFin As Boolean
Dim savdat As Date
Dim savWedType As Integer
Dim kaderPos As Integer
Dim deelnPag As Integer
Dim grpWedsAant As Integer
Dim nwKol As Boolean
Dim grpPnt As Integer
Dim grpPntTTL As Integer
Dim grpPntposY As Integer
Dim grpPntPosX As Integer
Dim endEersteDeelnPos As Integer
Dim tsYpos As Integer
Dim wedPnt As Integer
Dim ttl As Integer
Dim ttlPosX As Integer
Dim ttlPosY As Integer
Dim grpwedsTtlPosX As Integer
Dim grpwedsTtlPosY As Integer
Dim ttlgrpWeds As Integer
Dim Dagpnt As Integer
Dim dagpntposX As Integer
Dim dagpntposY As Integer
Dim savXpos As Integer
Dim savYpos As Integer
Dim toernooiGestart As Boolean
Dim aantalAfgedrukt As Integer
Dim AantalOpPapier As Integer
Dim prntReg As String
    toernooiGestart = KSStarted()
    If obj.ScaleHeight <> Printer.ScaleHeight Then
        Helft = Helft + obj.TextHeight("W") * 2
    End If
    grpWedsAant = AantGrpWeds()
    rot.Angle = 0
    wedHoog = 9
    NaamHoog = 11
    rsDeelnem.Open sqlDeelnems(thisPool), cn, adOpenStatic, adLockReadOnly
    
    If rsDeelnem.RecordCount = 0 Then
        MsgBox "Geen deelnemers in deze pool", vbQuestion + vbOKOnly, "Deelnemers afdrukken"
        Exit Sub
    End If
    KolomAant = 1
    X% = 20
    headerText = GetOrgNaam(thisPool) & " " & getTournamentInfo("toernooi") & " voetbalpool"
    tkst$ = "Deelnemers en Voorspellingen"
    kop$ = tkst$
    
    InitPage True, False
    FontGr NaamHoog
    obj.CurrentY = obj.CurrentY - 50
    kophoog = obj.CurrentY
    TopMarg = obj.CurrentY
    AantalOpPapier = 2
    If grpWedsAant <= 24 Then
        AantalOpPapier = 3
    End If
    Helft = (voethoog - TopMarg) / AantalOpPapier
'    Helft = obj.ScaleHeight / AantalOpPapier + 100 'obj.CurrentY
    FontGr wedHoog
    'Debug.Print obj.FontSize, Printer.FontSize * afdrRatio
    RegHeight% = obj.TextHeight("x") '* afdrRatio
    FontGr NaamHoog
    naamHeight = obj.TextHeight("x") '* afdrRatio
    If getTournamentInfo("groepen") > 4 Then
        KolomAant = getTournamentInfo("groepen")
    Else
        KolomAant = 8
    End If
    
    kolwidth = Int((obj.ScaleWidth / KolomAant) - 50)
    obj.FillStyle = vbFSTransparent
    rsDeelnem.MoveFirst
    FontGr 8
    posDatum = 50
    posWedOms = posDatum + obj.TextWidth("99-99:")
    posRust = posWedOms + obj.TextWidth("WWW-WWW")
    PosEind = posRust + obj.TextWidth("11-11")
    posToto = PosEind + obj.TextWidth("11-11")
    posPnt = posToto + obj.TextWidth("99")
    FontGr 12
    deelnPag = 0
    Do While Not rsDeelnem.EOF
        If Me.lstCompetitorPools.Selected(rsDeelnem.AbsolutePosition - 1) Or Me.Option3 = True Then
            showInfo True, "Afdrukken deelnemers", rsDeelnem!bijnaam, "Record " & rsDeelnem.AbsolutePosition & "/" & rsDeelnem.RecordCount
            
            If deelnPag = 0 Then
                obj.CurrentY = TopMarg
            Else
                obj.CurrentY = deelnPag * (Helft) + TopMarg
            End If
            LineYPos = obj.CurrentY
            obj.CurrentX = 30
            Vet True
            FontGr NaamHoog + 6
            obj.Print
            wedYpos = obj.CurrentY
            
            obj.Line (0, LineYPos)-(obj.ScaleWidth - 10, wedYpos), &H127419, BF
            obj.CurrentY = LineYPos
            obj.ForeColor = vbWhite
            iBKMode = SetBkMode(obj.hdc, TRANSPARENT)
            obj.CurrentX = 30
            obj.Print rsDeelnem!bijnaam;
            ttlPosX = obj.ScaleWidth
            ttlPosY = obj.CurrentY
            obj.Print
            Vet False
            obj.CurrentX = 50
            obj.ForeColor = 1
            'groepswedstrijden
            sqlstr = "Select * from qryDeelnWeds Where deelnem = " & rsDeelnem!deelnemID
            rsDeelnemWeds.Open sqlstr, cn, adOpenStatic, adLockReadOnly
            FontGr 10
            Vet True
            obj.ForeColor = vbBlue
            obj.Print "Groepswedstrijden";
            grpwedsTtlPosX = obj.CurrentX
            grpwedsTtlPosY = obj.CurrentY
            obj.CurrentX = obj.ScaleWidth * 0.75 + 50
            obj.Print "Finales";
            obj.ForeColor = 1
            Vet False
            obj.FontItalic = True
            FontGr 8
            'For i = 1 To 4
                'obj.CurrentX = obj.ScaleWidth / 4 * i - obj.TextWidth("pnt") - 50
                'obj.Print "pnt";
            'Next
            obj.FontItalic = False
            FontGr 10
            obj.Print
            obj.Line (0, wedYpos - 10)-(obj.ScaleWidth - 10, obj.CurrentY + 10), , B
            LineYPos = obj.CurrentY + 10
            obj.CurrentY = LineYPos
            FontGr 8
            LineXpos = 0
            With rsDeelnemWeds
'                showInfo True, "Afdrukken deelnemers", rsDeelnem!bijnaam, "Record " & rsDeelnem.AbsolutePosition  & "/" & rsDeelnem.RecordCount, "Wedstrijden"
                k = 0
                If .RecordCount > 0 Then
                    .MoveLast
                    .MoveFirst
                    wedKol = 1
                    Do While Not .EOF
                        obj.CurrentX = LineXpos + posWedOms - obj.TextWidth(Format(!Datum, "d-m") & ":") - 50
                        If savdat <> !Datum Or obj.CurrentY = LineYPos Then
                            obj.Print Format(!Datum, "d-m"); ":";
                            savdat = !Datum
                        End If
                        If nz(!tm1, "") > "" And !wedtype = 1 Then
                            pr = !tm1
                        Else
                            pr = nz(!code1, "")
                        End If
                        pr = pr & " - "
                        If nz(!tm2, "") > "" And !wedtype = 1 Then
                            pr = pr & !tm2
                        Else
                            pr = pr & !code2
                        End If
                        obj.CurrentX = LineXpos + posWedOms
                        obj.Print pr;
                        obj.CurrentX = LineXpos + posRust
                        obj.Print !r1; "-"; !r2;
                        obj.CurrentX = PosEind + LineXpos
                        obj.Print !e1; "-"; !e2;
                        obj.CurrentX = LineXpos + posToto
                        obj.Print !toto;
                        obj.Print
                        If newlinepos < obj.CurrentY Then newlinepos = obj.CurrentY
                        rsDeelnemWeds.MoveNext
                        If grpWedsAant < 25 Then
                            nwKol = (.AbsolutePosition - 1) Mod (grpWedsAant / 3) = 0 '= Int(grpWedsAant / 2) Or .AbsolutePosition = grpWedsAant
                        Else
                            nwKol = (.AbsolutePosition - 1) Mod 16 = 0
                        End If
                        If nwKol Then
                            obj.CurrentY = LineYPos
                            k = k + 1
                            If (.AbsolutePosition - 1) = grpWedsAant Then k = 3
                            LineXpos = (obj.ScaleWidth / 4) * k
                        End If
                    Loop
                End If
                .Close
            End With
            obj.Line (0, wedYpos)-(0, newlinepos)
            For i = 1 To 4
                obj.Line (obj.ScaleWidth / 4 * i - 20, LineYPos)-(obj.ScaleWidth / 4 * i - 20, newlinepos)
                obj.Line (obj.ScaleWidth / 4 * (i - 1) + posRust - 20, LineYPos)-(obj.ScaleWidth / 4 * (i - 1) + posRust - 20, newlinepos)
                obj.Line (obj.ScaleWidth / 4 * (i - 1) + PosEind - 20, LineYPos)-(obj.ScaleWidth / 4 * (i - 1) + PosEind - 20, newlinepos)
                obj.Line (obj.ScaleWidth / 4 * (i - 1) + posToto - 20, LineYPos)-(obj.ScaleWidth / 4 * (i - 1) + posToto - 20, newlinepos)
                obj.Line (obj.ScaleWidth / 4 * (i - 1) + posPnt - 20, LineYPos)-(obj.ScaleWidth / 4 * (i - 1) + posPnt - 20, newlinepos)
            Next
            FontGr 10
            'groepstanden
'            showInfo True, "Afdrukken deelnemers", rsDeelnem!bijnaam, "Record " & rsDeelnem.AbsolutePosition + 1 & "/" & rsDeelnem.RecordCount, "Groepstanden"
            obj.Line (0, newlinepos)-(obj.ScaleWidth, newlinepos)
            obj.Line (0, newlinepos)-(obj.ScaleWidth - 10, newlinepos + obj.TextHeight("Gr") + 10), , B
            obj.CurrentY = newlinepos + 10
            obj.CurrentX = 50
            Vet True
            obj.ForeColor = vbBlue
            obj.Print "Groepstanden"
            obj.ForeColor = 1
            Vet False
            LineYPos = obj.CurrentY
            kolwidth = Int((obj.ScaleWidth / KolomAant)) - 1
            FontGr 10
            sqlstr = "Select * from voorspelling_groepstand Where deelnem = " & rsDeelnem!deelnemID
            sqlstr = sqlstr & " ORDER BY groep"
            rsDeelnGroepen.Open sqlstr, cn, adOpenStatic, adLockReadOnly
            'LineYPos = obj.CurrentY - 10
            k = 0
            obj.CurrentX = 50
            Do While Not rsDeelnGroepen.EOF
                obj.FontUnderline = True
                obj.ForeColor = &H4000&
                obj.Print "Groep " & rsDeelnGroepen!groep
                obj.ForeColor = 1
                obj.FontUnderline = False
                
'                obj.CurrentX = obj.CurrentX + obj.TextWidth("|00")
                For i = 1 To 4
                    obj.CurrentX = kolwidth * k
                    pr = GetTeam(rsDeelnGroepen("pos" & Format(i, "0")))
                    If pr = "" Then pr = "?"
                    obj.Print i; ":"; pr
                    If newlinepos < obj.CurrentY Then newlinepos = obj.CurrentY
                Next
                k = k + 1
                obj.Line (kolwidth * (k - 1), LineYPos)-(kolwidth * (k), newlinepos), , B
                obj.CurrentX = kolwidth * k + 100
                obj.CurrentY = LineYPos
                rsDeelnGroepen.MoveNext
            Loop
            
            rsDeelnGroepen.Close
            If grpWedsAant > 24 Then
                obj.CurrentX = grpPntPosX
                obj.CurrentY = newlinepos
            Else
                obj.CurrentX = kolwidth * k
            End If
            'finales
            newlinepos = obj.CurrentY
            obj.Line (obj.CurrentX, newlinepos)-(obj.ScaleWidth, newlinepos)
            obj.CurrentY = newlinepos
            LineXpos = 0
            LineYPos = obj.CurrentY
            sqlstr = "Select * from qrydeelnemfinales WHERE deelnem=" & rsDeelnem!deelnemID
            sqlstr = sqlstr & " AND wedtype = " & AchtsteFinale
            sqlstr = sqlstr & " AND ksid= " & kampID
            If rsDeelnFinales.State = adStateOpen Then
                rsDeelnFinales.Close
            End If
            rsDeelnFinales.Open sqlstr, cn, adOpenStatic, adLockReadOnly
            If rsDeelnFinales.RecordCount > 0 Then
                With rsDeelnFinales
                    obj.CurrentX = LineXpos + 20
                    Vet True
                    obj.ForeColor = vbBlue
                    obj.Print "Achtste finales"
                    obj.ForeColor = 1
                    Vet False
                    Do While Not .EOF
                        obj.CurrentX = LineXpos + 50
                        prntReg = Format(!wed, "0") & ": " & !tm1 & " - " & !tm2
                        Do While obj.TextWidth(prntReg) > obj.ScaleWidth / 5 - 100
                          prntReg = Left(prntReg, Len(prntReg) - 1)
                        Loop
                        obj.Print prntReg;
                        obj.Print
                        If .AbsolutePosition = 4 Then
                            If LineYPos < obj.CurrentY Then LineYPos = obj.CurrentY
                            LineXpos = LineXpos + obj.ScaleWidth / 5
                            obj.CurrentY = newlinepos
                            obj.Print
                        End If
                        .MoveNext
                    Loop
                    .Close
                End With
                obj.CurrentY = newlinepos
            End If
            If grpWedsAant > 24 Then
                LineXpos = obj.ScaleWidth / 5 * 2
            Else
                LineXpos = obj.ScaleWidth / 2
            End If
            sqlstr = "Select distinct * from qrydeelnemfinales WHERE deelnem=" & rsDeelnem!deelnemID
            sqlstr = sqlstr & " AND wedtype = " & KwartFinale
            sqlstr = sqlstr & " AND ksid= " & kampID
            If rsDeelnFinales.State <> 0 Then rsDeelnFinales.Close
            rsDeelnFinales.Open sqlstr, cn, adOpenStatic, adLockReadOnly
            If rsDeelnFinales.RecordCount > 0 Then
                With rsDeelnFinales
                    obj.CurrentX = LineXpos + 50
                    Vet True
                    obj.ForeColor = vbBlue
                    obj.Print "Kwart finales"
                    obj.ForeColor = 1
                    Vet False
                    Do While Not .EOF
                        obj.CurrentX = LineXpos + 50
                        prntReg = Format(!wed, "0") & ": " & !tm1 & " - " & !tm2
                        Do While obj.TextWidth(prntReg) > obj.ScaleWidth / 5 - 100
                          prntReg = Left(prntReg, Len(prntReg) - 1)
                        Loop
                        obj.Print prntReg;
                        obj.Print
                        If LineYPos < obj.CurrentY Then LineYPos = obj.CurrentY
                        .MoveNext
                    Loop
                    .Close
                End With
                obj.CurrentY = newlinepos
            End If
            If grpWedsAant > 24 Then
                LineXpos = obj.ScaleWidth / 5 * 3
            Else
                LineXpos = obj.ScaleWidth / 4 * 3
            End If
            sqlstr = "Select DISTINCT * from qrydeelnemfinales WHERE deelnem=" & rsDeelnem!deelnemID
            sqlstr = sqlstr & " AND wedtype = " & HalveFinale
            sqlstr = sqlstr & " AND ksid= " & kampID
            If rsDeelnFinales.State = adStateOpen Then
                rsDeelnFinales.Close
            End If
            rsDeelnFinales.Open sqlstr, cn, adOpenStatic, adLockReadOnly
            If rsDeelnFinales.RecordCount > 0 Then
                With rsDeelnFinales
                    obj.CurrentX = LineXpos + 50
                    Vet True
                    obj.ForeColor = vbBlue
                    obj.Print "Halve finales"
                    obj.ForeColor = 1
                    If LineYPos < obj.CurrentY Then LineYPos = obj.CurrentY
                    Vet False
                   ' obj.Print
                    Do While Not .EOF
                        obj.CurrentX = LineXpos + 50
                        prntReg = Format(!wed, "0") & ": " & !tm1 & " - " & !tm2
                        Do While obj.TextWidth(prntReg) > obj.ScaleWidth / 5 - 100
                          prntReg = Left(prntReg, Len(prntReg) - 1)
                        Loop
                        obj.Print prntReg; ' Format(!wed, "0"); ": "; !tm1; " - "; !tm2;
                        obj.Print
                        .MoveNext
                    Loop
                    .Close
                End With
                If grpWedsAant > 24 Then
                    obj.CurrentY = newlinepos
                End If
            End If
            If grpWedsAant > 24 Then
                LineXpos = obj.ScaleWidth / 5 * 4
            Else
                LineXpos = obj.ScaleWidth / 4 * 3
            End If
            sqlstr = "Select * from qrydeelnemfinales WHERE deelnem=" & rsDeelnem!deelnemID
            sqlstr = sqlstr & " AND wedtype = " & klFinale
            sqlstr = sqlstr & " AND ksid= " & kampID
            If rsDeelnFinales.State = adStateOpen Then
                rsDeelnFinales.Close
            End If
            rsDeelnFinales.Open sqlstr, cn, adOpenStatic, adLockReadOnly
            
            If rsDeelnFinales.RecordCount > 0 Then
                With rsDeelnFinales
                    obj.CurrentX = LineXpos + 50
                    Vet True
                    obj.ForeColor = vbBlue
                    obj.Print "3e plaats"
                    obj.ForeColor = 1
                    Vet False
                    Do While Not .EOF
                        obj.CurrentX = LineXpos + 50
                        prntReg = Format(!wed, "0") & ": " & !tm1 & " - " & !tm2
                        Do While obj.TextWidth(prntReg) > obj.ScaleWidth / 5 - 100
                          prntReg = Left(prntReg, Len(prntReg) - 1)
                        Loop
                        obj.Print prntReg;
                        obj.Print
                        If LineYPos < obj.CurrentY Then LineYPos = obj.CurrentY
                        .MoveNext
                    Loop
                    obj.CurrentY = obj.CurrentY + 120
                    obj.Line (obj.ScaleWidth / 5 * 4, obj.CurrentY - 20)-(obj.ScaleWidth - 10, obj.CurrentY - 20)
                    obj.CurrentY = obj.CurrentY + 10
                End With
            End If
            sqlstr = "Select DISTINCT * from qrydeelnemfinales WHERE deelnem=" & rsDeelnem!deelnemID
            sqlstr = sqlstr & " AND wedtype = " & Finale
            sqlstr = sqlstr & " AND ksid= " & kampID
            If rsDeelnFinales.State = adStateOpen Then
                rsDeelnFinales.Close
            End If
            rsDeelnFinales.Open sqlstr, cn, adOpenStatic, adLockReadOnly
            If rsDeelnFinales.RecordCount > 0 Then
                With rsDeelnFinales
                    obj.CurrentX = LineXpos + 50
                    Vet True
                    obj.ForeColor = vbBlue
                    obj.Print "Finale"
                    obj.ForeColor = 1
                    Vet False
                    Do While Not .EOF
                        obj.CurrentX = LineXpos + 50
                        prntReg = Format(!wed, "0") & ": " & !tm1 & " - " & !tm2
                        Do While obj.TextWidth(prntReg) > obj.ScaleWidth / 5 - 100
                          prntReg = Left(prntReg, Len(prntReg) - 1)
                        Loop
                        obj.Print prntReg;
                        obj.Print
                        If LineYPos < obj.CurrentY Then LineYPos = obj.CurrentY
                    .MoveNext
                    Loop
                    .Close
                End With
            End If
            If grpWedsAant > 24 Then
                For i = 2 To 4
                    obj.Line (obj.ScaleWidth / 5 * i, newlinepos)-(obj.ScaleWidth / 5 * i, LineYPos)
                Next
            End If
            obj.Line (0, newlinepos)-(obj.ScaleWidth - 10, LineYPos), , B
            'uitslag
            LineYPos = obj.CurrentY + 50
            LineXpos = 50
            obj.CurrentX = LineXpos
            obj.CurrentY = LineYPos
            Vet True
            obj.ForeColor = vbBlue
            obj.Print "Eindstand"
            obj.ForeColor = 1
            Vet False
            obj.CurrentX = LineXpos
            pr = GetTeam(nz(rsDeelnem!kampioen, 0))
            If pr = "" Then pr = "?"
            obj.Print "1: "; pr
            obj.CurrentX = LineXpos
            If getPntToek("2e plaats") > 0 Then
                pr = GetTeam(nz(rsDeelnem!pltwee, 0))
                If pr = "" Then pr = "?"
                obj.Print "2: "; pr
            Else
                obj.Print
            End If
            obj.CurrentX = LineXpos
            If getPntToek("3e plaats") > 0 Then
                pr = GetTeam(nz(rsDeelnem!pldrie, 0))
                If pr = "" Then pr = "?"
                obj.Print "3: "; pr
            Else
                obj.Print
            End If
            obj.CurrentX = LineXpos
            If getPntToek("4e plaats") > 0 Then
              pr = GetTeam(nz(rsDeelnem!plvier, 0))
              If pr = "" Then pr = "?"
              obj.Print "4: "; pr
            Else
                obj.Print
            End If
            newlinepos = obj.CurrentY
            If deelnPag = 1 Then
                oldhelft = Helft
            End If
            obj.Line (0, LineYPos - 10)-(obj.ScaleWidth / 8, newlinepos), , B
            'topscorers
            LineXpos = obj.ScaleWidth / 8 + 50
            obj.CurrentX = LineXpos
            obj.CurrentY = LineYPos
            
            Vet True
            obj.CurrentX = LineXpos + 50
            obj.ForeColor = vbBlue
            obj.Print "Topscorer";
            If getPntToek("doelpunten topscorer 1") > 0 Then
                obj.CurrentX = (obj.ScaleWidth / 5 * 2) - obj.TextWidth("doelp") - 100
                obj.Print "doelp"
            Else
                obj.Print
            End If
            obj.ForeColor = 1
            tsYpos = obj.CurrentY
            kaderPos = obj.ScaleWidth / 5 * 2
            obj.Line (LineXpos, LineYPos - 10)-(kaderPos - 10, newlinepos), , B
            Vet False
            obj.CurrentY = tsYpos
            sqlstr = "Select * from voorspelling_ts WHERE deelnem = " & rsDeelnem!deelnemID
            sqlstr = sqlstr & " ORDER BY tsNR"
            rsDeelnts.Open sqlstr, cn, adOpenStatic, adLockReadOnly
            Do While Not rsDeelnts.EOF
                obj.CurrentX = LineXpos + 50
                pr = getSpelerNaam(nz(rsDeelnts!ts, 0))
                obj.Print pr;
                obj.CurrentX = kaderPos - obj.TextWidth(Format(rsDeelnts!dp, "0")) - 150
                If getPntToek("doelpunten topscorer 1") > 0 Then
                    If rsDeelnts!dp > -1 Then
                      obj.Print Format(rsDeelnts!dp, 0)
                    Else
                        obj.Print
                    End If
                Else
                    obj.Print
                End If
                rsDeelnts.MoveNext
            Loop
            rsDeelnts.Close
            'overige
            LineXpos = kaderPos + 20
            kaderPos = obj.ScaleWidth - 30
            obj.Line (LineXpos, LineYPos - 10)-(kaderPos, newlinepos), , B
            sqlstr = "Select * from qryDeelnVoorspAant WHERE deelnem = " & rsDeelnem!deelnemID
            rsDeelnOverig.Open sqlstr, cn, adOpenStatic, adLockReadOnly
            obj.CurrentY = LineYPos
            obj.CurrentX = LineXpos + 50
            Vet True
            obj.ForeColor = vbBlue
            obj.Print "Overigen ";
            obj.ForeColor = 1
            LineXpos = obj.CurrentX
            Vet False
            With rsDeelnOverig
                Do While Not .EOF
                    obj.CurrentX = LineXpos + 50
                    obj.Print !omschrijving; ": ";
                    obj.Print !Aantal
                    .MoveNext
                Loop
                .Close
            End With
            obj.DrawWidth = 2
            obj.Line (0, obj.CurrentY + 50)-(obj.ScaleWidth - 10, obj.CurrentY + 50)
            aantalAfgedrukt = aantalAfgedrukt + 1
        End If 'deeln selected
        rsDeelnem.MoveNext
        obj.CurrentX = 0
        If Not rsDeelnem.EOF Then
            If Me.lstCompetitorPools.Selected(rsDeelnem.AbsolutePosition - 1) Or Me.Option3 = True Then
                If deelnPag = AantalOpPapier - 1 Then
                    'obj.Line (0, Helft + 200)-(obj.ScaleWidth - 10, endEersteDeelnPos + 50), , B
                    deelnPag = 0
                    newlinepos = 0
                    'Exit Do
                    If Not rsDeelnem.EOF Then DoNewPage False, False
                Else
                    endEersteDeelnPos = obj.CurrentY
                    If aantalAfgedrukt > 0 Then deelnPag = deelnPag + 1
                    
                    If aantalAfgedrukt Mod (AantalOpPapier - 1) = 0 And aantalAfgedrukt > 0 Then
'                        Debug.Print "test"
                    End If
                    obj.Line (0, obj.CurrentY + 50)-(obj.ScaleWidth - 10, endEersteDeelnPos + 50)
                    'obj.Line (0, TopMarg)-(obj.ScaleWidth - 10, endEersteDeelnPos + 50), , B
                End If
                obj.DrawWidth = 1
            End If
        End If
    Loop
    rsDeelnem.Close
    showInfo False
End Sub

Private Sub btnPrntAllAfterDay_Click()
Dim i As Integer
Dim curWed As Integer
Dim savdat As Date
Dim msg As String
'stand in toernooi
Me.vscrlTM.value = GetMyNum(GetLastPlayed)
msg = "Voorspellingen afgedrukt"
If Me.vscrlTM.value > 0 Then
  msg = "Dagstanden, grafiek en voorspellingen afgedrukt"
  showInfo True, "Afdrukken", "Stand van zaken in toernooi", "Wedstrijd: " & Me.vscrlTM.value
  DoEvents
  optPrintDoc_Click 4
  btnPrint_Click 0
  'stand op punten
  DoEvents
  optPrintDoc_Click 2
  Me.ScoreVolg(1) = True
  showInfo True, "Afdrukken", "Stand op punten", "Wedstrijd: " & Me.vscrlTM.value
  btnPrint_Click 0
  'stand alfabetisch
  Screen.MousePointer = vbHourglass
  DoEvents
  optPrintDoc_Click 2
  Me.ScoreVolg(0) = True
  showInfo True, "Afdrukken", "Stand alfabetisch", "Wedstrijd: " & Me.vscrlTM.value
  btnPrint_Click 0
  'punten per wedstrijd alfabetisch
  DoEvents
  optPrintDoc_Click 6
  Me.ScoreVolg(0) = True
  showInfo True, "Afdrukken", "Punten per wedstrijd", "Wedstrijd: " & GetLastPlayed
  tmwed = GetLastPlayed
  btnPrint_Click 0
  'punten opbouw alfabetisch
  DoEvents
  optPrintDoc_Click 8
  Me.ScoreVolg(0) = True
  Me.optLandscape = True
  showInfo True, "Afdrukken", "Puntenopbouw", "Wedstrijd: " & GetLastPlayed
  btnPrint_Click 0
  'grafiek alfabetisch
  DoEvents
  optPrintDoc_Click 5
  Me.ScoreVolg(0) = True
  showInfo True, "Afdrukken", "Grafiek", "Wedstrijd: " & Me.vscrlTM.value
  btnPrint_Click 0
End If
'voorspellingen
curWed = GetMyNum(GetLastPlayed)
If curWed < GetWedAant(kampID) Then
    savdat = getWedDatum(GetWedNum(curWed + 1))
    For i = curWed + 1 To GetWedAant(kampID)
        If Format(getWedDatum(GetWedNum(i)), "d-m-yyyy") = Format(savdat, "d-m-yyyy") Then
            optPrintDoc_Click 7
            Me.vscrlVoor.value = i
            showInfo True, "Afdrukken", "Voorspelling", "Wedstrijd: " & i
            btnPrint_Click 0
        End If
    Next
End If
showInfo False
Screen.MousePointer = vbDefault
MsgBox msg, vbOKOnly + vbInformation, "Afdrukken"
End Sub

Sub EindStandAfdrukken()
Dim i As Integer
Dim curWed As Integer
Dim savdat As Date
'stand in toernooi
If MsgBox("Voor alle deelnemers afdrukken?", vbYesNo, "Eindstand") = vbYes Then
    Me.copies = getAantalUniekeDeelnems()
End If
Me.vscrlTM.value = GetMyNum(GetLastPlayed)
showInfo True, "Afdrukken", "Eindstand toernooi", "Wedstrijd: " & Me.vscrlTM.value
DoEvents
optPrintDoc_Click 4
Me.chkDblSide.value = 0
btnPrint_Click 0
'stand op punten
DoEvents
optPrintDoc_Click 2
Me.ScoreVolg(1) = True
showInfo True, "Afdrukken", "Stand op punten", "Wedstrijd: " & Me.vscrlTM.value
Me.chkDblSide.value = 0
btnPrint_Click 0
'punten per wedstrijd alfabetisch
DoEvents
optPrintDoc_Click 6
Me.ScoreVolg(0) = True
Me.chkDblSide.value = 0
showInfo True, "Afdrukken", "Punten per wedstrijd", "Wedstrijd: " & Me.vscrlTM.value
btnPrint_Click 0
'punten opbouw alfabetisch
DoEvents
optPrintDoc_Click 8
Me.ScoreVolg(0) = True
Me.optLandscape = True
Me.chkDblSide.value = 0
showInfo True, "Afdrukken", "Puntenopbouw", "Wedstrijd: " & GetLastPlayed
btnPrint_Click 0
'grafiek alfabetisch
DoEvents
optPrintDoc_Click 5
Me.ScoreVolg(0) = True
Me.chkDblSide.value = 0
showInfo True, "Afdrukken", "Grafiek", "Wedstrijd: " & Me.vscrlTM.value
btnPrint_Click 0

'klaar
showInfo False
Screen.MousePointer = vbDefault
MsgBox "Eindstand afgedrukt", vbOKOnly + vbInformation, "Afdrukken"

End Sub

Private Sub btnFinalPlayerPrint_Click()
    EindStandAfdrukken
End Sub

Private Sub cmbPrinters_Click()
   Dim prntr As Printer
      
   For Each prntr In Printers
      If cmbPrinters.List(cmbPrinters.ListIndex) = prntr.DeviceName Then
         Set Printer = prntr
      End If
   Next
End Sub

Sub btnPrint_Click(Index As Integer)
Dim i As Integer
Dim hoog As Integer
Dim breed As Integer
Dim DitDoen As Integer
Dim savOrient As Integer
   Dim prntr As Printer
      
   For Each prntr In Printers
      If cmbPrinters.List(cmbPrinters.ListIndex) = prntr.DeviceName Then
         Set Printer = prntr
      End If
   Next

    savOrient = Printer.Orientation
    Screen.MousePointer = vbHourglass
    If Me.optPortrait Then
        Printer.Orientation = vbPRORPortrait
    Else
        Printer.Orientation = vbPRORLandscape
    End If
    Init
    If Index = 0 Then
        Set obj = Printer
        If obj.Duplex <> 0 Then
            If Me.chkDblSide Then
                On Error Resume Next
                If Printer.Orientation = vbPRORPortrait Then
                    Printer.Duplex = 2
                Else
                    Printer.Duplex = 3
                End If
            Else
                Printer.Duplex = 1
            End If
        End If
        Printer.FontTransparent = True
        If Me.upDnCopies = 0 Then Me.upDnCopies = 1
        Printer.copies = Me.upDnCopies.value
        'afdrRatio = 1
    Else
        Me.Visible = False
        printPrev.Show
        If printPrev.printPages.UBound = 0 Then
            Set obj = printPrev.printPages(0)
        End If

    End If
    Set rot.Device = obj
    'Meter.Value = Meter.Min
    For i = 0 To 8
        If Me.optPrintDoc(i).value = True Then
            DitDoen = i
            Exit For
        End If
    Next
    DoEvents
    obj.Font = "Times New Roman"
    Select Case DitDoen
    Case 0
        FormulierenAfdrukken
    Case 1
        Deelnems
    Case 2
        'Stand in pool
        deelnemers Me.ScoreVolg(0), val(Me.upDnToMatch)
    Case 3
        'Favorieten
        Favorieten
    Case 4
        'toernooi stand
        ToernooiStand tmwed
    Case 5
        Grafiek
    Case 6
        'punten per wedstrijd
        DeelnemWeds
    Case 7
        'voorspellingen voor wedstrijd
        AfdrukVoorspWed Me.upDnForMatch
    Case 8
        'samenvatting stand
        EindAfrekening Me.ScoreVolg(0)
    End Select
    
    'Melding.Visible = False
    'Picture1.Visible = True
    DoEvents
    If Index = 0 Then
        Printer.EndDoc
    Else
        printPrev.pageContent.PaintPicture obj.Image, 0, 0, obj.Width, obj.Height
        Set obj = Nothing
    End If
    Screen.MousePointer = Default
    
End Sub

Sub ToernooiStand(tmwed As Integer)
Dim kopje As String
    headerText = GetOrgNaam(thisPool) & " " & getTournamentInfo("toernooi") & " voetbalpool - Stand van zaken"
    kopje = Format(GetWedInfo(tmwed, "datum"), "dddd d mmmm") & ": "
    kopje = kopje & GetWedInfo(tmwed, "naam1") & " vs " & GetWedInfo(tmwed, "naam2")
    kop$ = "Na wedstrijd " & tmwed & ", " & kopje
    InitPage False, True
    tnWeds
    tnGroepStanden
    tnFinales
    prnTopScorers
    
    prAantallen tmwed
    
End Sub

Sub prnTopScorers()
Dim sqlstr As String
Dim rs As New ADODB.Recordset
Dim rsED As New ADODB.Recordset 'voor de eigen doelpunten
Dim i As Integer
Dim grps As Integer
Dim colNu As Integer
Dim numpos As Integer
Dim datPos As Integer
Dim wedPos As Integer
Dim uitslPos As Integer
Dim newYpos As Integer
Dim ypos As Integer
Dim aantpos As Integer
Dim col(5) As Integer
    col(0) = 4
    col(1) = obj.ScaleWidth / 5
    col(2) = obj.ScaleWidth / 5 * 2
    col(3) = obj.ScaleWidth / 5 * 3
    col(4) = obj.ScaleWidth / 5 * 4
    col(5) = obj.ScaleWidth
    aantpos = obj.ScaleWidth / 5
    sqlstr = "select rnaam, afkort, count(rnaam) as aantal from qrywedverloop"
    sqlstr = sqlstr & " WHERE gebeurtenis <= 2"
    sqlstr = sqlstr & " AND ksid = " & kampID
    sqlstr = sqlstr & " GROUP BY rnaam, afkort"
    sqlstr = sqlstr & " ORDER BY count(rnaam) DESC, rnaam"
    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    sqlstr = "select rnaam, afkort, count(rnaam) as aantal from qrywedverloop"
    sqlstr = sqlstr & " WHERE gebeurtenis = 3"
    sqlstr = sqlstr & " AND ksid = " & kampID
    sqlstr = sqlstr & " GROUP BY rnaam, afkort"
    sqlstr = sqlstr & " ORDER BY count(rnaam) DESC, rnaam"
    rsED.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    If rsED.RecordCount > 0 Then
        rsED.MoveLast
        rsED.MoveFirst
    End If
    If rs.RecordCount > 0 Then
        rs.MoveLast
        rs.MoveFirst
    End If
    If rs.RecordCount + rsED.RecordCount > 0 Then
        FontGr 12
        ypos = obj.CurrentY
        obj.ForeColor = vbBlue
        Vet True
        obj.Print "Topscorers tot nu toe: "
        ypos = obj.CurrentY
        Vet False
        obj.ForeColor = 1
        FontGr 8
        Do While Not rs.EOF
            i = i + 1
            obj.CurrentX = col(colNu)
            obj.Print FirstPart(rs!rnaam) & " (" & LCase(rs!afkort) & ")";
            obj.CurrentX = col(colNu) + aantpos - obj.TextWidth("1234567890")
            obj.Print rs!Aantal
            
            
            rs.MoveNext
            If i = Int((rs.RecordCount + rsED.RecordCount + 1) / 5) + 1 Then
                i = 0
                colNu = colNu + 1
                newYpos = obj.CurrentY
                obj.CurrentY = ypos
            End If
        Loop
        If rsED.RecordCount > 0 Then
            obj.ForeColor = vbBlue
            Vet True
            i = i + 1
            obj.CurrentX = col(colNu)
            obj.Print "Eigen doelpunten:"
            If i = Int((rs.RecordCount + rsED.RecordCount + 1) / 5) + 1 Then
                i = 0
                colNu = colNu + 1
                newYpos = obj.CurrentY
                obj.CurrentY = ypos
            End If
            Vet False
            obj.ForeColor = 1
            Do While Not rsED.EOF
                i = i + 1
                obj.CurrentX = col(colNu)
                obj.Print FirstPart(rsED!rnaam) & " (" & LCase(rsED!afkort) & ")";
                obj.CurrentX = col(colNu) + aantpos - obj.TextWidth("1234567890")
                obj.Print rsED!Aantal
                
                
                rsED.MoveNext
                If i = Int((rs.RecordCount + rsED.RecordCount + 1) / 5) + 1 Then
                    i = 0
                    colNu = colNu + 1
                    newYpos = obj.CurrentY
                    obj.CurrentY = ypos
                End If
            Loop
            rsED.Close
        End If
        rs.Close
        obj.Line (0, ypos)-(obj.ScaleWidth - 50, newYpos), , B
        obj.CurrentY = newYpos
        obj.Print
    End If
End Sub

Sub prAantallen(tmwed As Integer)
Dim ypos As Integer
Dim prStr As String
Dim col(6) As Integer
    col(0) = 0
    col(1) = obj.ScaleWidth / 6
    col(2) = obj.ScaleWidth / 6 * 2
    col(3) = obj.ScaleWidth / 6 * 3
    col(4) = obj.ScaleWidth / 6 * 4
    col(5) = obj.ScaleWidth / 6 * 5
    col(6) = obj.ScaleWidth - 50
    FontGr 12
    obj.ForeColor = vbBlue
    Vet True
    obj.Print "Statistieken"
    ypos = obj.CurrentY
    Vet False
    obj.ForeColor = 1
    FontGr 10
    obj.CurrentX = col(0)
    prStr = "Doelpunten: " & Format(getAantal(tmwed, 1) + getAantal(tmwed, 2) + getAantal(tmwed, 3), pntFormat)
    obj.Print prStr;
    obj.CurrentX = col(1)
    prStr = "Penalties: " & Format(getAantal(tmwed, 1) + getAantal(tmwed, 6), pntFormat)
    obj.Print prStr;
    obj.CurrentX = col(2)
    prStr = "Gele kaarten: " & Format(getAantal(tmwed, 4), pntFormat)
    obj.Print prStr;
    obj.CurrentX = col(3)
    prStr = "Rode kaarten: " & Format(getAantal(tmwed, 5), pntFormat)
    obj.Print prStr;
    obj.CurrentX = col(4)
    prStr = "Gelijke spelen: " & Format(getAantalGelijkeSpelen(tmwed), pntFormat)
    obj.Print prStr;
    obj.CurrentX = col(5)
    prStr = "Eigen doelpunten: " & Format(getAantal(tmwed, 3), pntFormat)
    obj.Print prStr
    obj.ForeColor = vbBlue
    Ital True
    Centreer GetDeelnemAant(thisPool) & " deelnemers aan de pool"
    obj.Print
    Ital False
    obj.ForeColor = 1
    obj.Line (col(0), ypos)-(col(6), obj.CurrentY), , B
        
End Sub

Sub tnFinales()
Dim sqlstr As String
Dim rs As New ADODB.Recordset
Dim rsUitsl As New ADODB.Recordset
Dim i As Integer
Dim grps As Integer
Dim col(5) As Integer
Dim colNu As Integer
Dim numpos As Integer
Dim datPos As Integer
Dim wedPos As Integer
Dim vsPos As Integer
Dim uitslPos As Integer
Dim newYpos As Integer
Dim ypos As Integer
Dim topYpos As Integer
Dim wed As Integer
Dim uitsl As String
Dim colNr As Integer
Dim grpAant As Integer
grpAant = getTournamentInfo("groepen")
    col(0) = 20
    col(1) = obj.ScaleWidth / 3 + col(0)
    col(2) = obj.ScaleWidth / 3 * 2 + col(0)
    col(3) = obj.ScaleWidth
    col(4) = obj.ScaleWidth / 6 + col(0)
    col(5) = obj.ScaleWidth / 2 + col(0)
    sqlstr = "Select * from qryWeds "
    sqlstr = sqlstr & " WHERE ksid = " & kampID
    sqlstr = sqlstr & " AND wedtype <> 1"
    sqlstr = sqlstr & " order by mynum, wednum"
    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    Vet True
    FontGr 12
    obj.ForeColor = vbBlue
    obj.Print "Finales"
    topYpos = obj.CurrentY
    colNr = 0
    obj.CurrentX = col(colNr)
    FontGr 10
    If grpAant > 4 Then
        obj.Print "Achtste finales";
        colNr = colNr + 1
        obj.CurrentX = col(colNr)
    End If
    obj.Print "Kwart finales";
    colNr = colNr + 1
    obj.CurrentX = col(colNr)
    obj.Print "Halve finales";
    If colNr < 2 Then
        colNr = colNr + 1
        obj.CurrentX = col(colNr)
        obj.Print "Finale";
    End If
    ypos = obj.CurrentY
    Vet False
    obj.ForeColor = 1
    FontGr 8
    numpos = obj.TextWidth("00")
    datPos = numpos + obj.TextWidth("0")
    wedPos = datPos + obj.TextWidth("za 29 jun 20u:")
    vsPos = wedPos + obj.TextWidth("MEX")
    uitslPos = col(1) - obj.TextWidth("0-0(0-0)nvl:0-0(mexxx)")
    obj.Print
    ypos = obj.CurrentY
    Do While Not rs.EOF
        
        wed = rs!wedtype
        
        Select Case wed
        Case AchtsteFinale
            If grpAant > 4 Then
                colNu = 0
            End If
        Case KwartFinale
            If grpAant > 4 Then
                colNu = 1
            Else
                colNu = 0
            End If
        Case Finale
            colNu = 2
            If grpAant <= 4 Then
                obj.CurrentY = ypos
            End If
        Case Else
            If grpAant > 4 Then
                colNu = 2
            Else
                colNu = 1
            End If
        End Select
        obj.CurrentX = col(colNu) + numpos - obj.TextWidth(Format(rs!mynum, "0"))
        obj.Print Format(rs!mynum, "0");
        obj.CurrentX = col(colNu) + wedPos - obj.TextWidth(Format(rs!tijd, "ddd d mmm HHu") & ": ")
        obj.Print Format(rs!Datum, "ddd d mmm"); tijdFormat(rs!tijd, True); ": "; ' : , " HHu"); ": ";
        obj.CurrentX = col(colNu) + wedPos
        If nz(rs!tm1, "") > "" Then
            obj.Print rs!tm1;
        Else
            obj.Print rs!code1;
        End If
        obj.CurrentX = col(colNu) + vsPos
        
        If nz(rs!tm2, "") > "" Then
            obj.Print " - "; rs!tm2;
        Else
            obj.Print " - "; rs!code2;
        End If
        obj.CurrentX = col(colNu) + uitslPos
        If WedGespeeld(rs!wedNum) Then
            obj.Print GetWedUitsl(rs!wedNum)
        Else
            obj.Print
        End If
        rs.MoveNext
        If Not rs.EOF Then
            If rs!wedtype <> wed Then
                If newYpos < obj.CurrentY Then
                    newYpos = obj.CurrentY
                End If
                If rs!wedtype <> klFinale And rs!wedtype <> Finale Then
                    obj.CurrentY = ypos
                Else
                    Vet True
                    FontGr 12
                    obj.ForeColor = vbBlue
                    obj.CurrentX = col(2)
                    If rs!wedtype = klFinale Then
                        obj.Print "Derde plaats"
                    ElseIf grpAant > 4 Then
                        obj.CurrentX = col(2)
                        obj.Print "Finale"
                    End If
                    Vet False
                    obj.ForeColor = 1
                    FontGr 8
                End If
            End If
        End If
        
    Loop
    obj.Line (col(0) - 20, topYpos)-(col(1) - 50, newYpos), , B
    obj.Line (col(1) - 20, topYpos)-(col(2) - 50, newYpos), , B
    obj.Line (col(2) - 20, topYpos)-(col(3) - 50, newYpos), , B
    
    obj.CurrentY = newYpos
    obj.Print
End Sub


Sub tnGroepStanden()
Dim sqlstr As String
Dim rsGrp As New ADODB.Recordset
Dim i As Integer
Dim grps As Integer
Dim col(4) As Integer
Dim colNu As Integer
Dim teampos As Integer
Dim plPos As Integer
Dim wPos As Integer
Dim vPos As Integer
Dim gPos As Integer
Dim pntpos As Integer
Dim voorPos As Integer
Dim tegenPos As Integer
Dim pos As Integer 'de positie van het team in de groep

Dim ypos As Integer

    col(0) = 0
    col(1) = obj.ScaleWidth / 4
    col(2) = obj.ScaleWidth / 2
    col(3) = obj.ScaleWidth / 4 * 3
    col(4) = obj.ScaleWidth
    Vet True
    FontGr 12
    obj.ForeColor = vbBlue
    obj.Print "Groepstanden"
    ypos = obj.CurrentY
    Vet False
    obj.ForeColor = 1
    FontGr 8
    teampos = 10
    plPos = teampos + obj.TextWidth("1234567890123")
    wPos = plPos + obj.TextWidth("000")
    vPos = wPos + obj.TextWidth("000")
    gPos = vPos + obj.TextWidth("000")
    pntpos = gPos + obj.TextWidth("000")
    voorPos = pntpos + obj.TextWidth("000")
    tegenPos = voorPos + obj.TextWidth("000")
    
    
    grps = getTournamentInfo("groepen")
    colNu = 0
    For i = 1 To grps
        obj.CurrentY = ypos
        sqlstr = "Select * from qryGroepTeams"
        sqlstr = sqlstr & " Where ksID = " & kampID
        sqlstr = sqlstr & " AND groep = '" & Chr(i + 64) & "'"
        sqlstr = sqlstr & " order by pnt DESC, gesp, positie, plaatsing"
        rsGrp.Open sqlstr, cn, adOpenStatic, adLockReadOnly
        obj.CurrentX = col(colNu) + teampos
        obj.Print "groep " & Chr(i + 64);
        obj.CurrentX = col(colNu) + plPos
        obj.Print "sp";
        obj.CurrentX = col(colNu) + wPos
        obj.Print "W";
        obj.CurrentX = col(colNu) + vPos
        obj.Print "V";
        obj.CurrentX = col(colNu) + gPos
        obj.Print "G";
        obj.CurrentX = col(colNu) + pntpos
        obj.Print "P";
        obj.CurrentX = col(colNu) + voorPos
        obj.Print "v-t"
        Do While Not rsGrp.EOF
            pos = pos + 1
            obj.CurrentX = col(colNu) + teampos
            If rsGrp!positie <> 0 Then
                obj.Print Format(rsGrp!positie, "0"); ". "; rsGrp!naam;
            Else
                obj.Print Format(pos, "0"); ". "; rsGrp!naam;
            End If
            obj.CurrentX = col(colNu) + plPos
            obj.Print Format(rsGrp!gesp, "0");
            obj.CurrentX = col(colNu) + wPos
            obj.Print Format(rsGrp!gew, "0");
            obj.CurrentX = col(colNu) + vPos
            obj.Print Format(rsGrp!verl, "0");
            obj.CurrentX = col(colNu) + gPos
            obj.Print Format(rsGrp!gel, "0");
            obj.CurrentX = col(colNu) + pntpos
            obj.Print Format(rsGrp!pnt, "0");
            obj.CurrentX = col(colNu) + voorPos
            obj.Print Format(rsGrp!voor, "0"); "-"; Format(rsGrp!tegen, "0")
            rsGrp.MoveNext
        Loop
        obj.Line (col(colNu), ypos)-(col(colNu + 1) - 50, obj.CurrentY), , B
        colNu = colNu + 1
        If colNu > 3 Then
            colNu = 0
            ypos = obj.CurrentY + 50
        End If
        pos = 0
        rsGrp.Close
    Next
    obj.Print
End Sub

Sub tnWeds()
Dim sqlstr As String
Dim rs As New ADODB.Recordset
Dim rsUitsl As New ADODB.Recordset
Dim i As Integer
Dim grps As Integer
Dim col(3) As Integer
Dim colNu As Integer
Dim numpos As Integer
Dim datPos As Integer
Dim wedPos As Integer
Dim uitslPos As Integer
Dim newYpos As Integer
Dim ypos As Integer
    col(0) = 0
    col(1) = obj.ScaleWidth / 3
    col(2) = obj.ScaleWidth / 3 * 2
    col(3) = obj.ScaleWidth
    sqlstr = "Select * from qryWeds "
    sqlstr = sqlstr & " WHERE ksid = " & kampID
    sqlstr = sqlstr & " AND wedtype = 1"
    sqlstr = sqlstr & " order by mynum"
    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    rs.MoveLast
    rs.MoveFirst
    Vet True
    FontGr 12
    obj.ForeColor = vbBlue
    obj.Print "Groepswedstrijden"
    ypos = obj.CurrentY
    Vet False
    FontGr 8
    obj.ForeColor = 1
    numpos = obj.TextWidth("000")
    datPos = numpos + obj.TextWidth("0")
    wedPos = datPos + obj.TextWidth("za 29 jun 20uW")
    uitslPos = col(1) - obj.TextWidth("0-0 (0-0)")
    Do While Not rs.EOF
        i = i + 1
        obj.CurrentX = col(colNu) + numpos - obj.TextWidth(Format(rs!mynum, "0"))
        obj.Print Format(rs!mynum, "0");
        obj.CurrentX = col(colNu) + datPos
'        obj.Print Format(rs!Datum, "ddd d mmm"); Format(rs!tijd, " HHu."); ": ";
        
        obj.Print Format(rs!Datum, "ddd d mmm"); tijdFormat(rs!tijd, True); ": ";
        obj.CurrentX = col(colNu) + wedPos
        obj.Print rs!naam1 & " - " & rs!naam2;
        obj.CurrentX = col(colNu) + uitslPos
        If WedGespeeld(rs!wedNum) Then
            obj.Print GetWedUitsl(rs!wedNum)
        Else
            obj.Print
        End If
        rs.MoveNext
        If i = rs.RecordCount / 3 Then
            If newYpos < obj.CurrentY Then
                newYpos = obj.CurrentY
            End If
            i = 0
            obj.CurrentY = ypos
            colNu = colNu + 1
        End If
    Loop
    obj.Line (10, ypos)-(obj.ScaleWidth - 50, newYpos), , B
    obj.Line (col(1), ypos)-(col(1), newYpos)
    obj.Line (col(2), ypos)-(col(2), newYpos)
    obj.Print
End Sub

Private Sub DoNewPage(pagnr As Boolean, Optional vulKop As Boolean, Optional koppos As Integer)
    If TypeOf obj Is Printer Then
        Printer.NewPage
    Else
        Load printPrev.afdrpic(printPrev.afdrpic.UBound + 1)
        printPrev.afdrpic(printPrev.afdrpic.UBound).Visible = False
        printPrev.afdrpic(printPrev.afdrpic.UBound).AutoRedraw = True
        Set obj = printPrev.afdrpic(printPrev.afdrpic.UBound)
        printPrev.brnNext.Enabled = printPrev.afdrpic.UBound > 0
    End If
    InitPage pagnr, vulKop, koppos, True
End Sub

Private Sub FontGr(grootte%)

    Printer.FontSize = grootte%
    With obj.Font
        .Size = Printer.FontSize '* afdrratio
        
    End With
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim prntr As Printer
Dim rs As New ADODB.Recordset
Dim sqlstr As String
  Set cn = New ADODB.Connection
    With cn
      .ConnectionString = lclConn
      .CursorLocation = adUseClient
      .Open
    End With
    Set printPrev = New printPreview
    Me.picCompetitorList.Top = 90
    Me.picCompetitorList.Left = 3090
    Me.picPrnterSettings.Left = 3090
    
    Me.picPrnterSettings.Top = 2280
    sqlstr = "Select nickName from tblCompetitorPools where poolid=" & thisPool
    sqlstr = sqlstr & " order by nickName"
    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    Me.lstCompetitorPools.Clear
    Do While Not rs.EOF
        Me.lstCompetitorPools.AddItem rs!NickName
        rs.MoveNext
    Loop
    rs.Close
    cmbPrinters.Clear
    'Load the combo with all available printers
    For Each prntr In Printers
      cmbPrinters.AddItem prntr.DeviceName
      If Printer.DeviceName = prntr.DeviceName Then 'Current default
          cmbPrinters.Text = prntr.DeviceName
      End If
    Next
    
    'admin
    For i = 2 To 8
      Me.optPrintDoc(i).Visible = True
    Next
    Me.btnPrntAllAfterDay.Enabled = getAllMatchesPlayedOnDay(Date, cn)
    Me.btnFinalPlayerPrint.Enabled = getLastMatchPlayed(cn) = getMatchCount(cn)
    Me.chkEindstand.Enabled = Me.btnFinalPlayerPrint.Enabled
    
    headerText = getOrganisation(cn)
    
    tmwed = 0
    If wedNu >= 1 Then
        tmwed = wedNu
        Me.txtToMatch.Enabled = True
    Else
        tmwed = wedNu
    End If
   ' Me.chkDblSide.Enabled = printersettings
    Me.upDnToMatch.Max = wedNu
    Me.upDnForMatch.Max = getCount("Select tournamentID from tblTournamentSchedule where tournamentID = " & thisTournament, cn)
    Me.optPrintDoc(7).Enabled = getCount("Select competitorPoolID from tblCompetitorPools where poolID = " & thisPool, cn) > 0
    Me.optPrintDoc(1).Enabled = Me.optPrintDoc(7).Enabled
    Me.optPrintDoc(3).Enabled = Me.optPrintDoc(7).Enabled
    Me.optPrintDoc(2).Enabled = wedNu > 0
    Me.optPrintDoc(4).Enabled = wedNu > 0
    Me.optPrintDoc(5).Enabled = wedNu > 0
    Me.optPrintDoc(6).Enabled = wedNu > 0
    Me.optPrintDoc(8).Enabled = wedNu > 0
    Me.optPrintDoc(0).value = True
    optPrintDoc_Click 0
    Screen.MousePointer = Default
    ' Me.chkDblSide.Visible = true
    'Me.Eindstand.Enabled = GetLastPlayed = getlastWednum()
    'Me.btnFinalPlayerPrint.Visible = Me.Eindstand.Enabled
    Width = 6630
    Height = 5250
    centerForm Me
    UnifyForm Me
End Sub

Function RandomColor() As Long
    RandomColor = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
End Function


Private Sub Grafiek()
Dim rsPnt As New ADODB.Recordset
Dim rsDeeln As New ADODB.Recordset
Dim rsEtaps As New ADODB.Recordset
Dim sqlstr As String
Dim pnt As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
Dim xpos As Integer
Dim ypos As Integer
Dim yBot As Integer
Dim tmpX As Integer
Dim tmpY As Integer
Dim oldYpos As Integer
Dim bottom As Integer
Dim HoogsteNu As Integer
Dim langsteNaam As Integer
Dim wedAant As Integer
Dim deelnAant As Integer
Dim deelnemsOpPag As Integer
Dim pagaant As Integer
Dim deelnpos As Integer
Dim aantpnt As Integer
Dim maximum As Integer
Dim scorepos As Integer
Dim maximaal As Integer
Dim schaal As Double
Dim factor As Integer
Dim curPag As Integer
Dim deelnemsPagEen As Integer
Dim deelnemPagEenPos As Integer

MakeColors

kop$ = "Grafiek t/m wedstrijd " & tmwed
If Me.Eindstand <> 0 Then
    kop$ = "Grafiek Eindstand"
End If
InitPage False, False
FontGr 8
xpos = obj.CurrentX + obj.TextWidth("200") + obj.ScaleLeft
ypos = obj.CurrentY
sqlstr = "Select deelnemid, bijnaam from pooldeelnems"
sqlstr = sqlstr & " WHERE poolid =  " & thisPool
sqlstr = sqlstr & " Order BY bijnaam"
rsDeeln.Open sqlstr, cn, adOpenStatic, adLockReadOnly
rsDeeln.MoveLast
rsDeeln.MoveFirst
langsteNaam = obj.TextWidth(Left(GetLangsteBijNaam, 15))
langsteNaam = langsteNaam + obj.TextWidth("0(99)")
bottom = voethoog - langsteNaam
yBot = voethoog - TextHeight("999")
deelnAant = rsDeeln.RecordCount
If Me.optLandscape Then 'landscape
    deelnemsOpPag = 40
Else
    deelnemsOpPag = 26
End If
pagaant = 1
If deelnAant > deelnemsOpPag Then
    pagaant = (deelnAant / (deelnemsOpPag + 3) + 0.5)
End If

deelnemsOpPag = Int((deelnAant + 3) / pagaant + 0.5)
wedAant = GetWedAant(kampID)
HoogsteNu = getHoogPnt(tmwed)
If HoogsteNu > 250 Then
    factor = 50
ElseIf HoogsteNu > 150 Then
    factor = 25
ElseIf HoogsteNu > 100 Then
    factor = 10
Else
    factor = 5
End If
Do While aantpnt <= HoogsteNu / factor
    aantpnt = aantpnt + factor
Loop
'obj.Scale
maximum = Int(HoogsteNu / aantpnt + 1) * aantpnt
aantpnt = maximum / factor
scorepos = Int((bottom - ypos) / aantpnt)
'legenda
obj.FillStyle = vbSolid
oldYpos = bottom
FontGr 6
deelnemPagEenPos = obj.TextWidth("99: XXX-XXXX") + 20
obj.ForeColor = vbBlack
For i = 0 To tmwed - 1
    obj.FillColor = kleur(i)
    obj.Line (xpos, oldYpos)-(xpos + deelnemPagEenPos - 20, oldYpos - obj.TextHeight("W")), , B
    obj.CurrentX = xpos + 40
    SetForeCol kleur(i)
    obj.Print getWedTeams(i + 1)
    oldYpos = oldYpos - obj.TextHeight("W")
    obj.ForeColor = vbBlack
Next
FontGr 8

obj.Line (xpos + deelnemPagEenPos + 40, ypos)-(obj.ScaleWidth + 2 * obj.ScaleLeft, ypos)
obj.Line -(obj.ScaleWidth + 2 * obj.ScaleLeft, bottom)
obj.Line -(xpos + deelnemPagEenPos + 40, bottom)
obj.Line -(xpos + deelnemPagEenPos + 40, ypos)
For i = 0 To aantpnt
    ypos = bottom - i * scorepos
    FontGr 8
    obj.Line (xpos + deelnemPagEenPos + 40, ypos)-(obj.ScaleWidth + 2 * obj.ScaleLeft, ypos)
    obj.CurrentX = xpos + deelnemPagEenPos + 40 - TextWidth(CStr(i * maximum / aantpnt)) - 20
    obj.CurrentY = ypos - TextHeight("99") / 2
    obj.Print i * maximum / aantpnt
Next
maximaal = (i - 1) * aantpnt
schaal = (bottom - ypos) / maximum
'FontGr 4
Vet False
rsDeeln.MoveFirst
'kleur(0) = 15
curPag = 1
deelnpos = Int((obj.ScaleWidth - (2 * obj.ScaleLeft) - xpos - deelnemPagEenPos) / deelnemsOpPag)
i = 2 'horizontale positie eerste deelnemer
deelnemsPagEen = deelnemsOpPag - i
Do While Not rsDeeln.EOF
    i = i + 1
    oldYpos = bottom
'    If curPag > 1 Then deelnemsPagEen = deelnemsOpPag
    For j = 0 To tmwed - 1
        obj.FillColor = kleur(j)
        pnt = Int(getDeelnPnt(GetWedNum(j + 1), rsDeeln!deelnemID, 1) * (schaal) + 0.5)
        obj.Line (xpos + 10 + deelnpos * (i - 1), oldYpos)-(xpos + deelnpos * (i - 1) + deelnpos - 10, oldYpos - pnt), , B
        
        oldYpos = oldYpos - pnt
    Next
    FontGr 8
    obj.CurrentX = xpos + deelnpos * (i - 1) + (deelnpos - obj.TextWidth(Format(pnt, "999"))) / 2
    obj.CurrentY = oldYpos - obj.TextHeight(Format(pnt, "##"))
    
    obj.Print Int(getDeelnPnt(GetWedNum(j), rsDeeln!deelnemID, 0))
    obj.CurrentX = xpos + deelnpos * (i - 1) + (deelnpos - TextWidth("W")) / 2
    tmpX = obj.CurrentX
    
    obj.CurrentY = bottom + obj.TextWidth(Trim(rsDeeln!bijnaam) & " ")
    tmpY = obj.CurrentY
    Vet False
    FontGr 10
    Set rot.Device = obj
    obj.CurrentY = bottom + 50
    obj.CurrentX = xpos + deelnpos * (i - 1) + (deelnpos + obj.TextWidth("W")) / 2
    rot.Angle = 270
    rot.PrintText rsDeeln!bijnaam & " (" & getDeelnPnt(tmwed, rsDeeln!deelnemID, 8) & ")"
    rsDeeln.MoveNext
    obj.DrawWidth = 1
    If i = deelnemsOpPag And Not rsDeeln.EOF Then
        DoNewPage False, False
        curPag = curPag + 1
        obj.Line (xpos, ypos)-(obj.ScaleWidth + 2 * obj.ScaleLeft, ypos)
        obj.Line -(obj.ScaleWidth + 2 * obj.ScaleLeft, bottom)
        obj.Line -(xpos, bottom)
        obj.Line -(xpos, ypos)

        For i = 0 To aantpnt
            ypos = bottom - i * scorepos
            FontGr 8
            obj.Line (xpos, ypos)-(obj.ScaleWidth + 2 * obj.ScaleLeft, ypos)
            obj.CurrentX = xpos - TextWidth(CStr(i * maximum / aantpnt)) - 10
            obj.CurrentY = ypos - TextHeight("99") / 2
            obj.Print i * maximum / aantpnt
        Next
        i = 0
        Vet False
        obj.FillStyle = vbSolid
    End If
Loop
    
End Sub



Private Sub Init()
    With Printer
        .FontUnderline = 0
        .FontSize = 18
        GrootHoog = .TextHeight("Jota")
        .FontSize = 10
        KleinHoog = .TextHeight("Jota")
        .FontSize = 8
        SmallHoog = .TextHeight("Jota")
        .FontSize = 12
        NormHoog = .TextHeight("Jota")
        .DrawWidth = 2
    End With
End Sub

Private Sub InitPage(pagnr As Boolean, Optional vullen As Boolean, Optional koppos As Integer, Optional vervolg As Boolean)
' boolean 'doorloop' bepaalt of er een voetregel moet komen
    'Me.prnDialog.FontName = txtFont
    
    'If Not vervolg Or (vervolg And Me.chkNwePagKop) Then voetregel
    
    KopRegel
    koptekst kop$, pagnr, vullen, , koppos

End Sub

Private Sub Ital(Aan As Boolean)
    obj.FontItalic = Aan
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    Set obj = Nothing
  On Error GoTo 0
End Sub

Private Sub KlaarButton_Click()
On Error Resume Next

On Error GoTo 0
Printer.KillDoc
Unload printPrev
Unload Me
End Sub

Private Sub KopRegel()
Dim W%
Dim fnt As String
    obj.ForeColor = RGB(0, 51, 0)
    fnt = obj.FontName
    obj.FontName = "Times New Roman"
    W% = obj.DrawWidth
    obj.DrawWidth = 1
    obj.Line (0, 0)-(obj.ScaleWidth, 0), RGB(0, 51, 0)
    FontGr 4
    obj.Print
    FontGr 16
    Vet True
    Centreer CStr(headerText)
    obj.Print
    Y% = obj.CurrentY
    obj.Line (0, Y%)-(obj.ScaleWidth, Y%), RGB(0, 51, 0)
    FontGr 1
    Vet False
    obj.Print
    kophoog = obj.CurrentY
    obj.DrawWidth = W%
    obj.ForeColor = vbBlack
    obj.FontName = fnt
End Sub

Private Sub koptekst(Tekst$, pagnr As Boolean, Optional vul As Boolean, Optional ypos As Integer, Optional xpos As Integer)
    FontGr 16
    
    obj.FillColor = RGB(0, 51, 0)
    If vul Then
        obj.FillStyle = vbFSSolid
        obj.ForeColor = RGB(204, 251, 153)
        obj.Line (0, kophoog)-(obj.ScaleWidth - 20, kophoog + obj.TextHeight("W")), vbBlack, B
    Else
        obj.ForeColor = RGB(0, 51, 0)
        obj.FillStyle = vbFSTransparent
    End If
    Ital True
    Vet True
    obj.CurrentY = kophoog
    If ypos > 0 Then obj.CurrentY = ypos
    
    iBKMode = SetBkMode(obj.hdc, TRANSPARENT)
    Select Case xpos
    Case 0
        Centreer Tekst$
    Case 1
        obj.CurrentX = 0
        obj.Print Tekst$;
    Case 2
        obj.CurrentX = Int(obj.ScaleWidth / 4) - obj.TextWidth(Tekst$) / 2
        obj.Print Tekst$;
    Case 3
        obj.CurrentX = Int(obj.ScaleWidth / 2) - obj.TextWidth(Tekst$) / 2
        obj.Print Tekst$;
    Case 4
        obj.CurrentX = Int(obj.ScaleWidth / 4) * 3 - obj.TextWidth(Tekst$) / 2
        obj.Print Tekst$;
    End Select
    favYpos = obj.CurrentY
    FontGr 9
    obj.CurrentY = obj.CurrentY + GrootHoog - KleinHoog
    obj.CurrentX = obj.ScaleWidth - obj.TextWidth("blad 12")
    If TypeOf obj Is Printer Then
        If obj.Page > 1 And pagnr Then
            obj.Print "blad "; obj.Page;
        End If
    Else
        If obj.Index > 0 And pagnr Then
            obj.Print "blad "; obj.Index + 1;
        End If
    End If
    FontGr 12
    obj.Print
    kophoog = obj.CurrentY
    obj.FillStyle = vbFSTransparent
    obj.ForeColor = vbBlack
    Ital False
    Vet False
End Sub

Function getAant(deeln As Long, vanwat As String)
'haal het aantal scores op van 'vanwat' bij deeln
Dim rsdeelnScore As New ADODB.Recordset
Dim sqlstr As String
    sqlstr = "SELECT * from deelnempnt"
    sqlstr = sqlstr & " Where deelnID =" & deeln
    sqlstr = sqlstr & " AND " & vanwat & " > 0"
    rsdeelnScore.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    If rsdeelnScore.RecordCount > 0 Then
        rsdeelnScore.MoveLast
    End If
    getAant = rsdeelnScore.RecordCount

End Function

Function GetPntDeelnem(deeln As Long, vanwat As String)
Dim rsdeelnScore As New ADODB.Recordset
Dim pnt As Integer
Dim sqlstr As String
    sqlstr = "SELECT * from deelnempnt"
    sqlstr = sqlstr & " Where deelnID =" & deeln
    sqlstr = sqlstr & " AND " & vanwat & " > 0"
    sqlstr = sqlstr & " order by wednum"
    rsdeelnScore.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    If rsdeelnScore.RecordCount > 0 Then
        rsdeelnScore.MoveLast
        GetPntDeelnem = rsdeelnScore(vanwat)
        If UCase(Left(vanwat, 7)) = UCase("pntfin4") Then
            rsdeelnScore.MoveFirst
            pnt = 0
            Do While Not rsdeelnScore.EOF
                pnt = pnt + rsdeelnScore(vanwat)
                rsdeelnScore.MoveNext
            Loop
            GetPntDeelnem = pnt
        ElseIf UCase(Left(vanwat, 7)) = UCase("pntfin2") Then
            rsdeelnScore.MoveFirst
            pnt = 0
            Do While Not rsdeelnScore.EOF
                pnt = pnt + rsdeelnScore(vanwat)
                rsdeelnScore.MoveNext
            Loop
            GetPntDeelnem = pnt
        End If
    Else
        GetPntDeelnem = 0
    End If
End Function

Sub EindAfrekening(alfabet As Boolean)
Dim rsDeeln As New ADODB.Recordset
Dim rsdeelnScore As New ADODB.Recordset
Dim sqlstr As String
Dim bedr As Currency
Dim geldold As Currency
Dim savy As Integer
Dim leftmarge As Integer
Dim pntpos() As Integer
Dim pnt As Integer
Dim aant As Integer
Dim grpPnt As Integer
Dim geld As Double
Dim geldttl As Double
Dim Tekst$
Dim prStr As String
Dim topYpos As Integer
Dim top2Ypos As Integer
Dim botY As Integer
Dim lastDeelnPos As Integer
Dim maxY As Integer
Dim grp As String
Dim i As Integer
Dim j As Integer
Dim ipos As Integer
Dim has8eFin As Boolean
Dim hasKlFin As Boolean
Dim grpAant As Integer
Dim wdNum As Integer
Dim prTtl As Boolean
Dim colbr As Integer
Dim grpStndBegin As Integer '6
Dim fin8Begin As Integer    '15
Dim fin4Begin As Integer    '24
Dim fin2Begin As Integer    '29
Dim finBegin As Integer     '32

Dim EindstBegin As Integer  '34
Dim AantBegin As Integer    '38
Dim TopScBegin As Integer   '43
Dim TTLBegin As Integer     '44
Dim PosBegin As Integer     '45
Dim GeldBegin As Integer    '46

Dim tmp$
Dim yposnu%

    grpAant = getTournamentInfo("groepen")
    If grpAant > 4 Then
        colbr = 140
    Else
        colbr = 250
    End If
    has8eFin = grpAant > 4
    hasKlFin = getTournamentInfo("derdeplaats")
    If GetLastPlayed = getlastWednum Then
        pntFormat = "0"
    Else
        pntFormat = "0;;\ ;-"
    End If

    leftmarge = obj.CurrentX
    FontGr 10
    obj.Print
    
    FontGr 16
    Vet True
    If Me.Eindstand = False Then
        If alfabet Then
            Tekst$ = "Puntenopbouw t/m wedstrijd " & GetMyNum(GetLastPlayed)
        Else
            Tekst$ = "Puntenopbouw t/m wedstrijd (hoog-laag)" & GetMyNum(GetLastPlayed)
        End If
    Else
        If alfabet Then
            Tekst$ = "Eindstand (alfabetisch)"
        Else
            Tekst$ = "Eindstand (op score)"
        End If
    End If
    headerText = GetOrgNaam(thisPool) & " " & getTournamentInfo("toernooi") & " voetbalpool"

    kop$ = Tekst$
    
    
    
    InitPage False, True
    Ital False
    Vet False
    FontGr 8
    topYpos = obj.CurrentY
    obj.Line (0, topYpos)-(obj.ScaleWidth - 50, topYpos)
    obj.CurrentX = leftmarge
    sqlstr = DeelnResultSql(False, GetLastPlayed)
    rsDeeln.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    If rsDeeln.RecordCount > 0 Then
        rsDeeln.MoveLast
        lastDeelnPos = rsDeeln!postotaal
    End If
    rsDeeln.Close
    sqlstr = DeelnResultSql(alfabet, GetLastPlayed)
    
    rsDeeln.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    If rsDeeln.RecordCount = 0 Then
        obj.Print "Geen deelnemers gevonden"
        Exit Sub
    End If
    FontGr 10
    obj.CurrentX = leftmarge
    obj.Print "Naam";
    obj.CurrentX = obj.TextWidth("123456789012345")
    ReDim Preserve pntpos(1)
    pntpos(0) = 0
    pntpos(1) = obj.CurrentX - colbr
    obj.Print
    top2Ypos = obj.CurrentY
    obj.CurrentX = pntpos(1) + colbr
    FontGr 8
    obj.Print "rust"; '("; Format(getPnt(1), pntFormat); "p)";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = obj.CurrentX
    obj.CurrentX = pntpos(UBound(pntpos)) + colbr
    obj.Print "eind"; '("; Format(getPnt(2), pntFormat); "p)";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = obj.CurrentX
    obj.CurrentX = pntpos(UBound(pntpos)) + colbr
    obj.Print "toto"; '("; Format(getPnt(3), pntFormat); "p)";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = obj.CurrentX
    obj.CurrentX = pntpos(UBound(pntpos)) + colbr
    obj.Print "dlp"; '("; Format(getPnt(28), pntFormat); "p)";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = obj.CurrentX
    obj.CurrentX = pntpos(UBound(pntpos)) + colbr
    obj.Print "tot";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = obj.CurrentX
    grpStndBegin = UBound(pntpos)
    
    For i = 1 To grpAant
        obj.CurrentX = pntpos(UBound(pntpos)) + colbr
        obj.Print Chr(i + 64);
        ReDim Preserve pntpos(UBound(pntpos) + 1)
        pntpos(UBound(pntpos)) = obj.CurrentX
    Next
    obj.CurrentX = pntpos(UBound(pntpos)) + colbr
    obj.Print "tot";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = obj.CurrentX
    If grpAant > 4 Then
        fin8Begin = UBound(pntpos)
        For i = 1 To grpAant
            obj.CurrentX = pntpos(UBound(pntpos)) + colbr
            obj.Print Chr(i + 64);
            ReDim Preserve pntpos(UBound(pntpos) + 1)
            pntpos(UBound(pntpos)) = obj.CurrentX
        Next
        obj.CurrentX = pntpos(UBound(pntpos)) + colbr
        obj.Print "tot";
        ReDim Preserve pntpos(UBound(pntpos) + 1)
        pntpos(UBound(pntpos)) = obj.CurrentX
    End If
    fin4Begin = UBound(pntpos)
    For i = 1 To 4
        obj.CurrentX = pntpos(UBound(pntpos)) + colbr
        obj.Print Format(i, "0");
        ReDim Preserve pntpos(UBound(pntpos) + 1)
        pntpos(UBound(pntpos)) = obj.CurrentX
    Next
    obj.CurrentX = pntpos(UBound(pntpos)) + colbr
    obj.Print "tot";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = obj.CurrentX
    fin2Begin = UBound(pntpos)
    For i = 1 To 2
        obj.CurrentX = pntpos(UBound(pntpos)) + colbr
        obj.Print "  "; Format(i, "0"); "e  ";
        ReDim Preserve pntpos(UBound(pntpos) + 1)
        pntpos(UBound(pntpos)) = obj.CurrentX
    Next
    obj.CurrentX = pntpos(UBound(pntpos)) + colbr
    obj.Print "tot";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = obj.CurrentX
    finBegin = UBound(pntpos)
    obj.CurrentX = pntpos(UBound(pntpos)) + colbr
    If hasKlFin Then
        obj.Print "kl("; Format(getPnt(30), pntFormat);
        If getPnt(31) > 0 Then
            obj.Print "/"; Format(getPnt(31), pntFormat);
        End If
        obj.Print ")";
        ReDim Preserve pntpos(UBound(pntpos) + 1)
        pntpos(UBound(pntpos)) = obj.CurrentX
        obj.CurrentX = pntpos(UBound(pntpos)) + colbr
        obj.Print "gr("; Format(getPnt(11), pntFormat);
        If getPnt(12) > 0 Then
            obj.Print "/"; Format(getPnt(12), pntFormat);
        End If
        obj.Print ")";
        ReDim Preserve pntpos(UBound(pntpos) + 1)
        pntpos(UBound(pntpos)) = obj.CurrentX
        obj.CurrentX = pntpos(UBound(pntpos)) + colbr
    Else
        obj.Print "("; Format(getPnt(11), pntFormat);
        If getPnt(12) > 0 Then
            obj.Print "/"; Format(getPnt(12), pntFormat);
        End If
        obj.Print ")";
        ReDim Preserve pntpos(UBound(pntpos) + 1)
        pntpos(UBound(pntpos)) = obj.CurrentX
        obj.CurrentX = pntpos(UBound(pntpos)) + colbr
    End If
    EindstBegin = UBound(pntpos)
    ' Format(getPnt(15), pntFormat); "/"; Format(getPnt(14), pntFormat); "/"; Format(getPnt(13), pntFormat); "/"; Format(getPnt(29), pntFormat); ")";
    obj.Print "1";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = obj.CurrentX
    obj.CurrentX = pntpos(UBound(pntpos)) + colbr
    obj.Print "2";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = obj.CurrentX
    If hasKlFin Then
        obj.CurrentX = pntpos(UBound(pntpos)) + colbr
        obj.Print "3";
        ReDim Preserve pntpos(UBound(pntpos) + 1)
        pntpos(UBound(pntpos)) = obj.CurrentX
        obj.CurrentX = pntpos(UBound(pntpos)) + colbr
        obj.Print "4";
        ReDim Preserve pntpos(UBound(pntpos) + 1)
        pntpos(UBound(pntpos)) = obj.CurrentX
    End If
    AantBegin = UBound(pntpos)
    obj.CurrentX = pntpos(UBound(pntpos)) + colbr
    obj.Print "dp";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = obj.CurrentX
    obj.CurrentX = pntpos(UBound(pntpos)) + colbr
    obj.Print "gel";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = obj.CurrentX
    obj.CurrentX = pntpos(UBound(pntpos)) + colbr
    obj.Print "gl";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = obj.CurrentX
    obj.CurrentX = pntpos(UBound(pntpos)) + colbr
    obj.Print "rd";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = obj.CurrentX
    obj.CurrentX = pntpos(UBound(pntpos)) + colbr
    obj.Print "pn";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = obj.CurrentX
    TopScBegin = UBound(pntpos)
    obj.CurrentX = pntpos(UBound(pntpos)) + colbr
    obj.Print "scor";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = obj.CurrentX
    TTLBegin = UBound(pntpos)
    obj.CurrentX = pntpos(UBound(pntpos)) + colbr + obj.TextWidth("123")
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = obj.CurrentX
    PosBegin = UBound(pntpos)
    obj.CurrentX = pntpos(UBound(pntpos)) + colbr + obj.TextWidth("123")
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = obj.CurrentX
    GeldBegin = UBound(pntpos)
    obj.CurrentX = pntpos(UBound(pntpos)) + colbr
    obj.Print "";
    'laatste kolom
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = obj.ScaleWidth - 50
    
    obj.CurrentY = topYpos
    FontGr 10
    obj.CurrentX = (pntpos(1) + pntpos(grpStndBegin) + colbr - obj.TextWidth("Wedstrijdpunten")) / 2
    obj.Print "Wedstrijdpunten";
    If grpAant > 4 Then
        obj.CurrentX = (pntpos(grpStndBegin) + pntpos(fin8Begin) + colbr - obj.TextWidth("Groepstand (" & Format(getPnt(8), pntFormat) & "p)")) / 2
    Else
        obj.CurrentX = (pntpos(grpStndBegin) + pntpos(fin4Begin) + colbr - obj.TextWidth("Groepstand (" & Format(getPnt(8), pntFormat) & "p)")) / 2
    End If
    obj.Print "Groepstand (" & Format(getPnt(8), pntFormat) & "p)";
    If grpAant > 4 Then
        obj.CurrentX = (pntpos(fin8Begin) + pntpos(fin4Begin) + colbr - obj.TextWidth("8e Finalisten (" & Format(getPnt(6), pntFormat) & "/" & Format(getPnt(7), pntFormat) & "p)")) / 2
        obj.Print "8e Finalisten (" & Format(getPnt(4), pntFormat);
        If getPnt(5) > 0 Then
            obj.Print "/" & Format(getPnt(5), pntFormat);
        End If
        obj.Print "p)";
    End If
    obj.CurrentX = (pntpos(fin4Begin) + pntpos(fin2Begin) + colbr - obj.TextWidth("4e fin.(" & Format(getPnt(6), pntFormat) & "/" & Format(getPnt(7), pntFormat) & "p)")) / 2
    obj.Print "4efin.(" & Format(getPnt(6), pntFormat);
    If getPnt(7) > 0 Then
        obj.Print "/" & Format(getPnt(7), pntFormat);
    End If
    obj.Print "p)";
    obj.CurrentX = (pntpos(fin2Begin) + pntpos(finBegin) + colbr - obj.TextWidth("2efin.(" & Format(getPnt(9), pntFormat) & "/" & Format(getPnt(10), pntFormat) & "p)")) / 2
    obj.Print "1/2fin.(" & Format(getPnt(9), pntFormat);
    If getPnt(10) > 0 Then
        obj.Print "/" & Format(getPnt(10), pntFormat);
    End If
    obj.Print "p)";
    obj.CurrentX = (pntpos(finBegin) + pntpos(EindstBegin) + colbr - obj.TextWidth("Fin")) / 2
    obj.Print "Fin";
    obj.CurrentX = (pntpos(EindstBegin) + pntpos(AantBegin) + colbr - obj.TextWidth("Eind")) / 2
    obj.Print "Eind";
    obj.CurrentX = (pntpos(AantBegin) + pntpos(TopScBegin) + colbr - obj.TextWidth("Aantallen")) / 2
    obj.Print "Aantallen";
    obj.CurrentX = pntpos(TopScBegin) + colbr
    obj.Print "top";
    obj.CurrentX = (pntpos(TTLBegin) + pntpos(PosBegin) + colbr - obj.TextWidth("Ttl")) / 2
    obj.Print "Ttl";
    obj.CurrentX = (pntpos(PosBegin) + pntpos(GeldBegin) + colbr - obj.TextWidth("Pos")) / 2
    obj.Print "Pos";
    obj.CurrentX = (pntpos(GeldBegin) + pntpos(GeldBegin + 1) + colbr - obj.TextWidth("Geld")) / 2
    obj.Print "Geld";
    FontGr 8
    obj.CurrentY = top2Ypos
    obj.CurrentX = pntpos(UBound(pntpos)) + colbr
    obj.Print
    obj.Line (0, obj.CurrentY)-(obj.ScaleWidth - 50, obj.CurrentY)
    With rsDeeln
        Do While Not .EOF
'            If rsDeeln!deelnemID = 251 Then Stop
            obj.CurrentX = leftmarge
            If !postotaal = 1 Then
                obj.ForeColor = vbBlue
                Vet True
            End If
            If !postotaal = lastDeelnPos Then
                obj.ForeColor = vbRed
            End If
            obj.Print !bijnaam;
            obj.ForeColor = 1
            Vet False
            pnt = PrintAant(!deelnemID, pntpos(2), "pntRust")
            pnt = pnt + PrintAant(!deelnemID, pntpos(3), "pntEind")
            pnt = pnt + PrintAant(!deelnemID, pntpos(4), "pntToto")
            pnt = pnt + PrintAant(!deelnemID, pntpos(5), "dpvddag")
            obj.CurrentX = pntpos(6) - obj.TextWidth(Format(pnt, pntFormat))
            Vet True
            obj.Print Format(pnt, pntFormat);
            Vet False
            pnt = 0
            grpPnt = 0
            For i = 1 To grpAant
                If allPlayed(Chr(i + 64)) Then
                    pntFormat = "0"
                Else
                    pntFormat = "0;;\ ;-"
                End If
                grpPnt = GetPntDeelnem(!deelnID, "pntgrp" & Chr(i + 64))
                pnt = pnt + grpPnt
                obj.CurrentX = (pntpos(i + 5) + pntpos(i + 6) + colbr - obj.TextWidth(Format(grpPnt, pntFormat))) / 2
                obj.Print Format(grpPnt, pntFormat);
            Next
            If grpAant > 4 Then
                obj.CurrentX = pntpos(fin8Begin) - obj.TextWidth(Format(pnt, pntFormat))
            Else
                obj.CurrentX = pntpos(fin4Begin) - obj.TextWidth(Format(pnt, pntFormat))
            End If
            Vet True
            obj.Print Format(pnt, pntFormat);
            Vet False
            pnt = 0
            grpPnt = 0
            If grpAant > 4 Then
                For i = 1 To grpAant
                    If allPlayed(Chr(i + 64)) Then
                        pntFormat = "0"
                        grpPnt = GetPntDeelnem(!deelnID, "pntfin8" & Chr(i + 64))
                    Else
                        grpPnt = 0
                        pntFormat = "0;;\ ;-"
                    End If
                    pnt = pnt + grpPnt
                    obj.CurrentX = (pntpos(fin8Begin - 1 + i) + pntpos(i + fin8Begin) + colbr - obj.TextWidth(Format(grpPnt, pntFormat))) / 2
                    obj.Print Format(grpPnt, pntFormat);
                Next
                obj.CurrentX = pntpos(fin4Begin) - obj.TextWidth(Format(pnt, pntFormat))
                Vet True
                If allPlayed("A") Then
                    pntFormat = "0"
                Else
                    pntFormat = "0;;\ ;-"
                End If
                obj.Print Format(pnt, pntFormat);
                Vet False
                pnt = 0
                grpPnt = 0
            Else
                For i = 1 To grpAant
                    If allPlayed(Chr(i + 64)) Then
                        pntFormat = "0"
                        grpPnt = GetPntDeelnem(!deelnID, "pntfin4" & Chr(i + 64))
                    Else
                        grpPnt = 0
                        pntFormat = "0;;\ ;-"
                    End If
                    pnt = pnt + grpPnt
                    obj.CurrentX = (pntpos(fin4Begin - 1 + i) + pntpos(i + fin4Begin) + colbr - obj.TextWidth(Format(grpPnt, pntFormat))) / 2
                    obj.Print Format(grpPnt, pntFormat);
                Next
                Vet True
                If allPlayed("A") Then
                    pntFormat = "0"
                Else
                    pntFormat = "0;;\ ;-"
                End If
                obj.CurrentX = pntpos(fin2Begin) - obj.TextWidth(Format(pnt, pntFormat))
                obj.Print Format(pnt, pntFormat);
                Vet False
                pnt = 0
                grpPnt = 0
            End If
            
            If grpAant > 4 Then
                For i = 1 To 8 Step 2
                    grpPnt = 0
                    'If !deelnID = 139 Then Stop
                    wdNum = i + j + GetFirstFinaleMatch(AchtsteFinale) - 1
                    Select Case wdNum
                    Case 49, 50
                        grp = "B"
                        ipos = 2
                    Case 51, 52
                        grp = "C"
                        ipos = 3
                    Case 53, 54
                        grp = "A"
                        ipos = 1
                    Case 55, 56
                        grp = "D"
                        ipos = 4
                    End Select
                    If GetMyNum(wdNum) <= GetMyNum(GetLastPlayed) Then
                        pntFormat = "0"
                        grpPnt = grpPnt + GetPntDeelnem(!deelnID, "pntFin4" & grp)
                        'grpPnt = grpPnt + getDeelnPnt(GetPrevWednum(wdNum), !deelnID, 9, "4" & grp)
                        prTtl = True
                    Else
                        pntFormat = "0;;\ ;-"
                        grpPnt = 0
                    End If
                    pnt = pnt + grpPnt
                    obj.CurrentX = (pntpos(ipos + fin4Begin - 1) + pntpos(ipos + fin4Begin) + colbr - obj.TextWidth(Format(grpPnt, pntFormat))) / 2
                    obj.Print Format(grpPnt, pntFormat);
                Next
                If prTtl > 0 Then pntFormat = "0"
                obj.CurrentX = pntpos(fin2Begin) - obj.TextWidth(Format(pnt, pntFormat))
                Vet True
                obj.Print Format(pnt, pntFormat);
                Vet False
            End If
            pnt = 0
            grpPnt = 0
            For i = 1 To 2
                If GetMyNum(i + GetFirstFinaleMatch(KwartFinale) - 1) <= GetMyNum(GetLastPlayed) Then
                    pntFormat = "0"
                Else
                    pntFormat = "0;;\ ;-"
                End If
                'If !deelnID = 183 Then Stop
                grpPnt = GetPntDeelnem(!deelnID, "pntfin2" & Chr(i + 64))
                pnt = pnt + grpPnt
                obj.CurrentX = (pntpos(i + fin2Begin - 1) + pntpos(i + fin2Begin) + colbr - obj.TextWidth(Format(grpPnt, pntFormat))) / 2
                obj.Print Format(grpPnt, pntFormat);
            Next
            obj.CurrentX = pntpos(finBegin) - obj.TextWidth(Format(pnt, pntFormat))
            Vet True
            obj.Print Format(pnt, pntFormat);
            Vet False
            If GetMyNum(GetFirstFinaleMatch(HalveFinale)) <= GetMyNum(GetLastPlayed) Then
                pntFormat = "0"
            Else
                pntFormat = "0;;\ ;-"
            End If
            If hasKlFin Then
                grpPnt = GetPntDeelnem(!deelnID, "pntklfin")
                obj.CurrentX = pntpos(32) + (pntpos(33) - pntpos(32) + colbr - obj.TextWidth(Format(grpPnt, pntFormat))) / 2
                obj.Print Format(grpPnt, pntFormat);
            End If
            grpPnt = GetPntDeelnem(!deelnID, "pntfin")
            obj.CurrentX = (pntpos(finBegin + 1) + pntpos(EindstBegin) + colbr - obj.TextWidth(Format(grpPnt, pntFormat))) / 2
            obj.Print Format(grpPnt, pntFormat);
            pntFormat = "0;;\ ;-"
            If GetLastPlayed = getlastWednum Then pntFormat = "0"
            For i = 1 To 2
                grpPnt = getEindStandpnt(!deelnID, i)
                obj.CurrentX = (pntpos(finBegin + 1 + i) + pntpos(EindstBegin + i) + colbr - obj.TextWidth(Format(grpPnt, pntFormat))) / 2
                obj.Print Format(grpPnt, pntFormat);
            Next
            pntFormat = "0;;\ ;-"
            If GetLastPlayed >= getlastWednum - 1 Then pntFormat = "0"
            For i = 3 To 4
                grpPnt = getEindStandpnt(!deelnID, i)
                obj.CurrentX = (pntpos(EindstBegin - 1 + i) + pntpos(EindstBegin + i) + colbr - obj.TextWidth(Format(grpPnt, pntFormat))) / 2
                obj.Print Format(grpPnt, pntFormat);
            Next
            pntFormat = "0;;\ ;-"
            If GetLastPlayed = getlastWednum Then
                pntFormat = "0"
                grpPnt = getDeelnAantPnt(!deelnID, voorspDP)
                obj.CurrentX = (pntpos(AantBegin) + pntpos(AantBegin + 1) + colbr - obj.TextWidth(Format(grpPnt, pntFormat))) / 2
                obj.Print Format(grpPnt, pntFormat);
                grpPnt = getDeelnAantPnt(!deelnID, voorspGelijk)
                obj.CurrentX = (pntpos(AantBegin + 1) + pntpos(AantBegin + 2) + colbr - obj.TextWidth(Format(grpPnt, pntFormat))) / 2
                obj.Print Format(grpPnt, pntFormat);
                grpPnt = getDeelnAantPnt(!deelnID, voorspGeel)
                obj.CurrentX = (pntpos(AantBegin + 2) + pntpos(AantBegin + 3) + colbr - obj.TextWidth(Format(grpPnt, pntFormat))) / 2
                obj.Print Format(grpPnt, pntFormat);
                grpPnt = getDeelnAantPnt(!deelnID, voorspRood)
                obj.CurrentX = (pntpos(AantBegin + 3) + pntpos(AantBegin + 4) + colbr - obj.TextWidth(Format(grpPnt, pntFormat))) / 2
                obj.Print Format(grpPnt, pntFormat);
                grpPnt = getDeelnAantPnt(!deelnID, voorspPens)
                obj.CurrentX = (pntpos(AantBegin + 4) + pntpos(TopScBegin) + colbr - obj.TextWidth(Format(grpPnt, pntFormat))) / 2
                obj.Print Format(grpPnt, pntFormat);
                grpPnt = GetPntDeelnem(!deelnID, "pntTopSc")
                obj.CurrentX = (pntpos(TopScBegin) + pntpos(TTLBegin) + colbr - obj.TextWidth(Format(grpPnt, pntFormat))) / 2
                obj.Print Format(grpPnt, pntFormat);
            End If
            pntFormat = "0"
            If !postotaal = 1 Then
                obj.ForeColor = vbBlue
                Vet True
            End If
            If !postotaal = lastDeelnPos Then
                obj.ForeColor = vbRed
            End If
            'If !deelnID = 125 Then Stop
            grpPnt = GetPntDeelnem(!deelnID, "grandtotaal")
            obj.CurrentX = (pntpos(TTLBegin) + pntpos(PosBegin) + colbr - obj.TextWidth(Format(grpPnt, pntFormat))) / 2
            obj.Print Format(grpPnt, pntFormat);
            grpPnt = GetPntDeelnem(!deelnID, "postotaal")
            obj.CurrentX = (pntpos(PosBegin) + pntpos(GeldBegin) + colbr - obj.TextWidth(Format(grpPnt, pntFormat))) / 2
            obj.Print Format(grpPnt, pntFormat);
            obj.ForeColor = 1
            Vet False
            geld = GetPntDeelnem(!deelnID, "geldttl")
            obj.CurrentX = pntpos(GeldBegin + 1) - colbr - obj.TextWidth(Format(geld, "currency"))
            obj.Print Format(geld, "currency");
            obj.Print
            obj.ForeColor = 1
            obj.Line (0, obj.CurrentY)-(obj.ScaleWidth - 50, obj.CurrentY)
            grpPnt = 0
            
            .MoveNext
'            If .AbsolutePosition >= 53 Then Stop
            If obj.CurrentY >= voethoog Then 'onderkant pagina
              If Not rsDeeln.EOF Then
                botY = obj.CurrentY
                obj.Line (pntpos(1) + 75, topYpos)-(pntpos(1) + 75, top2Ypos)
                obj.Line (pntpos(grpStndBegin) + 75, topYpos)-(pntpos(6) + 75, top2Ypos)
                If grpAant > 4 Then
                    obj.Line (pntpos(fin8Begin) + 75, topYpos)-(pntpos(15) + 75, top2Ypos)
                End If
                obj.Line (pntpos(fin4Begin) + 75, topYpos)-(pntpos(fin4Begin) + 75, top2Ypos)
                obj.Line (pntpos(fin2Begin) + 75, topYpos)-(pntpos(fin2Begin) + 75, top2Ypos)
                obj.Line (pntpos(finBegin) + 75, topYpos)-(pntpos(finBegin) + 75, top2Ypos)
                obj.Line (pntpos(EindstBegin) + 75, topYpos)-(pntpos(EindstBegin) + 75, top2Ypos)
                obj.Line (pntpos(AantBegin) + 75, topYpos)-(pntpos(AantBegin) + 75, top2Ypos)
                obj.Line (pntpos(TopScBegin) + 75, topYpos)-(pntpos(TopScBegin) + 75, top2Ypos)
                obj.Line (pntpos(TTLBegin) + 75, topYpos)-(pntpos(TTLBegin) + 75, top2Ypos)
                obj.Line (pntpos(PosBegin) + 75, topYpos)-(pntpos(PosBegin) + 75, top2Ypos)
                obj.Line (pntpos(GeldBegin) + 75, topYpos)-(pntpos(GeldBegin) + 75, top2Ypos)
                For i = 1 To UBound(pntpos) - 1
                    obj.Line (pntpos(i) + 75, top2Ypos)-(pntpos(i) + 75, botY)
                Next
                obj.Line (obj.ScaleWidth - 50, topYpos)-(obj.ScaleWidth - 50, botY)
                DoNewPage False, True
                obj.Line (0, topYpos)-(obj.ScaleWidth - 50, topYpos)
                obj.CurrentX = leftmarge
                obj.CurrentY = topYpos
                FontGr 10
                obj.Print "Naam";
                obj.CurrentY = top2Ypos
                obj.CurrentX = pntpos(1) + colbr
                FontGr 8
                obj.Print "rust"; '("; Format(getPnt(1), pntFormat); "p)";
                obj.CurrentX = pntpos(2) + colbr
                obj.Print "eind"; '("; Format(getPnt(2), pntFormat); "p)";
                obj.CurrentX = pntpos(3) + colbr
                obj.Print "toto"; '("; Format(getPnt(3), pntFormat); "p)";
                obj.CurrentX = pntpos(4) + colbr
                obj.Print "dlp"; '("; Format(getPnt(28), pntFormat); "p)";
                obj.CurrentX = pntpos(5) + colbr
                obj.Print "tot";
                If grpAant > 4 Then
                    For i = 1 To 8
                        obj.CurrentX = pntpos(5 + i) + colbr
                        obj.Print Chr(i + 64);
                    Next
                    obj.CurrentX = pntpos(14) + colbr
                    obj.Print "tot";
                    For i = 1 To 8
                        obj.CurrentX = pntpos(14 + i) + colbr
                        obj.Print Chr(i + 64);
                    Next
                    obj.CurrentX = pntpos(23) + colbr
                    obj.Print "tot";
                End If
                For i = 1 To 4
                    obj.CurrentX = pntpos(fin4Begin - 1 + i) + colbr
                    obj.Print Format(i, "0");
                Next
                obj.CurrentX = pntpos(fin2Begin - 1) + colbr
                obj.Print "tot";
                For i = 1 To 2
                    obj.CurrentX = pntpos(fin2Begin - 1 + i) + colbr
                    obj.Print "  "; Format(i, "0"); "e  ";
                Next
                obj.CurrentX = pntpos(finBegin - 1) + colbr
                obj.Print "tot";
                obj.CurrentX = pntpos(finBegin) + colbr
                If hasKlFin Then
                    obj.Print "kl("; Format(getPnt(30), pntFormat);
                    If getPnt(31) > 0 Then
                        obj.Print "/"; Format(getPnt(31), pntFormat);
                    End If
                    obj.Print ")";
                    obj.CurrentX = pntpos(EindstBegin - 1) + colbr
                    obj.Print "gr("; Format(getPnt(11), pntFormat);
                    If getPnt(12) > 0 Then
                        obj.Print "/"; Format(getPnt(12), pntFormat);
                    End If
                    obj.Print ")";
                Else
                    obj.Print "("; Format(getPnt(11), pntFormat);
                    If getPnt(12) > 0 Then
                        obj.Print "/"; Format(getPnt(12), pntFormat);
                    End If
                    obj.Print ")";
                End If
                
                For i = 1 To grpAant / 2
                    obj.CurrentX = pntpos(EindstBegin - 1 + i) + colbr
                    ' Format(getPnt(15), pntFormat); "/"; Format(getPnt(14), pntFormat); "/"; Format(getPnt(13), pntFormat); "/"; Format(getPnt(29), pntFormat); ")";
                    obj.Print Format(i, "0");
                Next
                obj.CurrentX = pntpos(AantBegin) + colbr
                obj.Print "dp";
                obj.CurrentX = pntpos(AantBegin + 1) + colbr
                obj.Print "gel";
                obj.CurrentX = pntpos(AantBegin + 2) + colbr
                obj.Print "gl";
                obj.CurrentX = pntpos(AantBegin + 3) + colbr
                obj.Print "rd";
                obj.CurrentX = pntpos(AantBegin + 4) + colbr
                obj.Print "pn";
                obj.CurrentX = pntpos(TopScBegin) + colbr
                obj.Print "scor";
                'laatste kolom
                obj.CurrentX = obj.ScaleWidth - 50
                
                obj.CurrentY = topYpos
                FontGr 10
                obj.CurrentX = (pntpos(1) + pntpos(grpStndBegin) + colbr - obj.TextWidth("Wedstrijdpunten")) / 2
                obj.Print "Wedstrijdpunten";
                If grpAant > 4 Then
                    obj.CurrentX = (pntpos(grpStndBegin) + pntpos(fin8Begin) + colbr - obj.TextWidth("Groepstand (" & Format(getPnt(8), pntFormat) & "p)")) / 2
                Else
                    obj.CurrentX = (pntpos(grpStndBegin) + pntpos(fin4Begin) + colbr - obj.TextWidth("Groepstand (" & Format(getPnt(8), pntFormat) & "p)")) / 2
                End If
                obj.Print "Groepstand (" & Format(getPnt(8), pntFormat) & "p)";
                If grpAant > 4 Then
                    obj.CurrentX = (pntpos(fin8Begin) + pntpos(fin4Begin) + colbr - obj.TextWidth("8e Finalisten (" & Format(getPnt(6), pntFormat) & "/" & Format(getPnt(7), pntFormat) & "p)")) / 2
                    obj.Print "8e Finalisten (" & Format(getPnt(4), pntFormat);
                    If getPnt(5) > 0 Then
                        obj.Print "/" & Format(getPnt(5), pntFormat);
                    End If
                    obj.Print "p)";
                End If
                obj.CurrentX = (pntpos(fin4Begin) + pntpos(fin2Begin) + colbr - obj.TextWidth("4e fin.(" & Format(getPnt(6), pntFormat) & "/" & Format(getPnt(7), pntFormat) & "p)")) / 2
                obj.Print "4efin.(" & Format(getPnt(6), pntFormat);
                If getPnt(7) > 0 Then
                    obj.Print "/" & Format(getPnt(7), pntFormat);
                End If
                obj.Print "p)";
                obj.CurrentX = (pntpos(fin2Begin) + pntpos(finBegin) + colbr - obj.TextWidth("2efin.(" & Format(getPnt(9), pntFormat) & "/" & Format(getPnt(10), pntFormat) & "p)")) / 2
                obj.Print "1/2fin.(" & Format(getPnt(9), pntFormat);
                If getPnt(10) > 0 Then
                    obj.Print "/" & Format(getPnt(10), pntFormat);
                End If
                obj.Print "p)";
                obj.CurrentX = (pntpos(finBegin) + pntpos(EindstBegin) + colbr - obj.TextWidth("Fin")) / 2
                obj.Print "Fin";
                obj.CurrentX = (pntpos(EindstBegin) + pntpos(AantBegin) + colbr - obj.TextWidth("Eind")) / 2
                obj.Print "Eind";
                obj.CurrentX = (pntpos(AantBegin) + pntpos(TopScBegin) + colbr - obj.TextWidth("Aantallen")) / 2
                obj.Print "Aantallen";
                obj.CurrentX = pntpos(TopScBegin) + colbr
                obj.Print "top";
                obj.CurrentX = (pntpos(TTLBegin) + pntpos(PosBegin) + colbr - obj.TextWidth("Ttl")) / 2
                obj.Print "Ttl";
                obj.CurrentX = (pntpos(PosBegin) + pntpos(GeldBegin) + colbr - obj.TextWidth("Pos")) / 2
                obj.Print "Pos";
                obj.CurrentX = (pntpos(GeldBegin) + pntpos(GeldBegin + 1) + colbr - obj.TextWidth("Geld")) / 2
                obj.Print "Geld";
                FontGr 8
                obj.CurrentY = top2Ypos
                obj.CurrentX = pntpos(UBound(pntpos)) + colbr
                obj.Print
                obj.Line (0, obj.CurrentY)-(obj.ScaleWidth - 50, obj.CurrentY)
              End If
            End If
        Loop
    End With
    botY = obj.CurrentY
    obj.Line (pntpos(1) + 75, topYpos)-(pntpos(1) + 75, top2Ypos)
    obj.Line (pntpos(grpStndBegin) + 75, topYpos)-(pntpos(6) + 75, top2Ypos)
    If grpAant > 4 Then
        obj.Line (pntpos(fin8Begin) + 75, topYpos)-(pntpos(15) + 75, top2Ypos)
    End If
    obj.Line (pntpos(fin4Begin) + 75, topYpos)-(pntpos(fin4Begin) + 75, top2Ypos)
    obj.Line (pntpos(fin2Begin) + 75, topYpos)-(pntpos(fin2Begin) + 75, top2Ypos)
    obj.Line (pntpos(finBegin) + 75, topYpos)-(pntpos(finBegin) + 75, top2Ypos)
    obj.Line (pntpos(EindstBegin) + 75, topYpos)-(pntpos(EindstBegin) + 75, top2Ypos)
    obj.Line (pntpos(AantBegin) + 75, topYpos)-(pntpos(AantBegin) + 75, top2Ypos)
    obj.Line (pntpos(TopScBegin) + 75, topYpos)-(pntpos(TopScBegin) + 75, top2Ypos)
    obj.Line (pntpos(TTLBegin) + 75, topYpos)-(pntpos(TTLBegin) + 75, top2Ypos)
    obj.Line (pntpos(PosBegin) + 75, topYpos)-(pntpos(PosBegin) + 75, top2Ypos)
    obj.Line (pntpos(GeldBegin) + 75, topYpos)-(pntpos(GeldBegin) + 75, top2Ypos)
    For i = 1 To UBound(pntpos) - 1
        obj.Line (pntpos(i) + 75, top2Ypos)-(pntpos(i) + 75, botY)
    Next
    obj.Line (obj.ScaleWidth - 50, topYpos)-(obj.ScaleWidth - 50, botY)
End Sub

Function PrintAant(deelnem As Long, pos, vanwat As String)
Dim aant As Integer
Dim pnt As Long
    Select Case vanwat
    Case "pntRust"
    pnt = getPnt(1)
    Case "pntEind"
    pnt = getPnt(2)
    Case "pntToto"
    pnt = getPnt(3)
    Case "dpvddag"
    pnt = getPnt(28)
    End Select
    If LCase(Left(vanwat, 6)) = "pntgrp" Then
        pnt = getPnt(8)
    End If
    
    aant = getAant(deelnem, vanwat)
'    obj.CurrentX = pos - obj.TextWidth("(" & Format(aant, "0") & "x) " & Format(aant * pnt, "0"))
    obj.CurrentX = pos - obj.TextWidth(Format(aant * pnt, "0"))
    Ital True
'    obj.Print "(" & Format(aant, "0"); "x) ";
    Ital False
    obj.Print Format(aant * pnt, "0");
    PrintAant = aant * pnt
End Function

Sub deelnemers(alfabet As Boolean, wedNum As Integer)
' En nu de deelnemers
Dim rsDeeln As New ADODB.Recordset
Dim rsdeelnScore As New ADODB.Recordset
Dim bedr As Currency
Dim pnt As Integer
Dim last As Integer
Dim eerst As Integer
Dim lastttl As Integer
Dim verh As Double
Dim geldold As Currency
Dim savy As Integer
Dim leftmarge As Integer
Dim deelkolwidth%
Dim DeelOldPntPos%
Dim DeelWedPntPos%
Dim DeelNewPntPos%
Dim deelnaampos%
Dim deelgeldpos%
Dim DeelGeldnwPos%
Dim DeelGeldttlPos%
Dim Tekst$
Dim prStr As String
Dim yLinePos%
Dim DeelTopPos%
Dim i As Integer
Dim tmp$
Dim yposnu%
    'wednum = GetWedNum(wednum)
    leftmarge = obj.CurrentX
    deelkolwidth% = (obj.ScaleWidth + 2 * obj.ScaleLeft) \ 2
    FontGr 10
    deelnaampos% = obj.TextWidth("999.")
    DeelOldPntPos% = deelnaampos% + deelkolwidth% / 4 - 200
    DeelWedPntPos% = DeelOldPntPos% + deelkolwidth / 10
    DeelNewPntPos% = DeelWedPntPos% + deelkolwidth / 10
    
    deelgeldpos% = DeelNewPntPos% + deelkolwidth / 6 + 200
    DeelGeldnwPos% = deelgeldpos% + deelkolwidth / 6 - 100
    DeelGeldttlPos% = DeelGeldnwPos% + deelkolwidth / 6 - 100
    
    If alfabet Then
        deelnaampos% = Me.CurrentX + 40
    End If
    
    obj.Print
    
    FontGr 16
    Vet True
    If alfabet Then
        Tekst$ = "Resultaat (A-Z) na " & GetMyNum(wedNum) & "e wed: " & GetWedInfo(wedNum, "naam1") & "-" & GetWedInfo(wedNum, "naam2") & ": " & GetWedUitsl(wedNum)
    Else
        Tekst$ = "Stand na " & GetMyNum(wedNum) & "e wed: " & GetWedInfo(wedNum, "naam1") & "-" & GetWedInfo(wedNum, "naam2") & ": " & GetWedUitsl(wedNum)
    End If
    If Me.Eindstand Then
        If alfabet Then
            Tekst$ = "Eindstand alfabetisch"
        Else
            Tekst$ = "Eindstand"
        End If
    End If
    headerText = GetOrgNaam(thisPool) & " " & getTournamentInfo("toernooi") & " voetbalpool"

    kop$ = Tekst$
    
    InitPage False, True
    Ital False
    Vet False
    FontGr 10
    obj.CurrentX = (obj.ScaleWidth - obj.TextWidth("onderstreept=daghoogste, vet=bovenaan, cursief=onderaan")) / 2
    obj.Print "(";
    obj.FontUnderline = True
    obj.ForeColor = &H8000&
    obj.Print "onderstreept";
    obj.FontUnderline = False
    obj.ForeColor = 0
    obj.Print "= daghoogste, ";
    obj.ForeColor = vbBlue
    Vet True
    obj.Print "vet";
    Vet False
    obj.ForeColor = 0
    obj.Print "= bovenaan, ";
    Ital True
    obj.ForeColor = vbRed
    obj.Print "cursief";
    Ital False
    obj.ForeColor = 0
    obj.Print "= onderaan)"
    
    savy = obj.CurrentY
    For kol% = 0 To 1
        If Not alfabet Then
            obj.CurrentX = kol% * deelkolwidth%
            'obj.Print "pos";
        End If
        obj.CurrentX = deelnaampos% + kol% * deelkolwidth%
        obj.Print "Naam";
        If alfabet Then obj.Print " (pl)";
        obj.CurrentX = DeelOldPntPos% + kol% * deelkolwidth%
        obj.Print "had  +";
        obj.CurrentX = DeelWedPntPos% + kol% * deelkolwidth%
        obj.Print "erbij =";
        obj.CurrentX = DeelNewPntPos% + kol% * deelkolwidth% + obj.TextWidth("999") - obj.TextWidth("nu")
        obj.Print "nu";
        obj.CurrentX = deelgeldpos% - obj.TextWidth("Geld") + kol% * deelkolwidth%
        obj.Print "Geld";
        obj.CurrentX = DeelGeldnwPos% - obj.TextWidth("erbij") + kol% * deelkolwidth%
        obj.Print "erbij";
        obj.CurrentX = DeelGeldttlPos% - obj.TextWidth("totaal") + kol% * deelkolwidth%
        obj.Print "totaal";
    Next
    obj.CurrentY = obj.CurrentY + 50
    yLinePos% = obj.CurrentY + TextHeight("test")
    obj.Line (leftmarge, yLinePos%)-(obj.ScaleWidth + obj.ScaleLeft * 2, yLinePos%)
    obj.CurrentY = obj.CurrentY + 50
    DeelTopPos% = obj.CurrentY
'    obj.Print
    'bepaal hoogste en laagste
    rsDeeln.Open DeelnResultSql(False, wedNum), cn, adOpenStatic, adLockReadOnly 'op punten volgorde dus
    If rsDeeln.RecordCount > 0 Then
        rsDeeln.MoveLast
        last = nz(rsDeeln!grandtotaal, 0)
    Else
        Exit Sub
    End If
    rsDeeln.Close
    obj.CurrentX = 0
    'en nu opnieuw openen
    rsDeeln.Open DeelnResultSql(alfabet, wedNum), cn, adOpenStatic, adLockReadOnly 'op volgorde dus
    With rsDeeln
        If .RecordCount > 0 Then
            .MoveFirst
            lastttl = 0
            kol% = 0
            Do While Not .EOF
                i = i + 1
                If i = Int(.RecordCount / 2 + 0.5) + 1 Then
                    kol% = deelkolwidth%
                    obj.CurrentY = DeelTopPos%
                End If
                obj.CurrentX = obj.CurrentX + deelnaampos% - obj.TextWidth(!postotaal) - obj.TextWidth("..") + kol%
                If Not alfabet Then
                    If lastttl <> !grandtotaal Then obj.Print !postotaal & ".";
                End If
                Vet !postotaal = 1
                Ital nz(!grandtotaal, 0) = last
                prStr = Left(!bijnaam, 12)
                If alfabet Then
                    prStr = prStr & " (" & !postotaal & ")"
                End If
                If !grandtotaal = last Then
                    obj.ForeColor = vbRed
                ElseIf nz(!postotaal, 0) = 1 Then
                    obj.ForeColor = vbBlue
                ElseIf nz(!posdag, 0) = 1 Then
                    obj.ForeColor = &H8000&
                Else
                    obj.ForeColor = 0
                End If
                obj.CurrentX = deelnaampos% + kol%
                obj.FontUnderline = nz(!posdag, 0) = 1
                
                obj.Print prStr;
                Vet False
                Ital False
                obj.ForeColor = 0
                obj.FontUnderline = False
                If wedNum > 1 Then
                    pnt = getTussenstand(!deelnemID, wedNum)
                    geldold = getTussenstandGeld(!deelnemID, GetWedNumPrevDag(wedNum))
                Else
                    pnt = 0
                    geldold = 0
                End If
                
                obj.CurrentX = DeelOldPntPos% + kol% + obj.TextWidth("999") - obj.TextWidth(CStr(pnt))
                obj.Print Format$(pnt, "##0");
                Vet False
                pnt = nz(!Dagpnt, 0)
                obj.CurrentX = DeelWedPntPos% + kol% + obj.TextWidth("999") - obj.TextWidth(CStr(pnt))
                obj.FontUnderline = nz(!posdag, 0) = 1
                If !posdag = 1 Then
                    obj.ForeColor = &H8000&
                Else
                    obj.ForeColor = 0
                End If
                obj.Print Format$(pnt, "##0");
                obj.ForeColor = 0
                obj.FontUnderline = False
                Vet !postotaal = 1
                Ital nz(!grandtotaal, 0) = last
                pnt = nz(!grandtotaal, 0)
                If !grandtotaal = last Then
                    obj.ForeColor = vbRed
                ElseIf !postotaal = 1 Then
                    obj.ForeColor = vbBlue
                Else
                    obj.ForeColor = 0
                End If
                obj.CurrentX = DeelNewPntPos% + kol% + obj.TextWidth("999") - obj.TextWidth(CStr(pnt))
                If !grandtotaal = last Then
                    obj.ForeColor = &H80&
                ElseIf !postotaal = 1 Then
                    obj.ForeColor = &HC00000
                Else
                    obj.ForeColor = 0
                End If
                obj.Print Format$(!grandtotaal, "##0");
                obj.ForeColor = 0
                Vet False
                Ital False
                tmp$ = Format$(geldold, "€ ##0.00")
                obj.CurrentX = deelgeldpos% - obj.TextWidth(tmp$) + kol%
                obj.Print tmp$;   '= geld
                tmp$ = Format$(!daggeldttl, "€ ##0.00")
                obj.CurrentX = DeelGeldnwPos% - obj.TextWidth(tmp$) + kol%
                obj.Print tmp$;
                bedr = 0
                tmp$ = Format$(geldold + !daggeldttl, "€ ##0.00")
                obj.CurrentX = DeelGeldttlPos% - obj.TextWidth(tmp$) + kol%
                obj.Print tmp$;   '= geld
                obj.Print
                lastttl = nz(!grandtotaal, 0)
                rsDeeln.MoveNext
            Loop
            obj.Print
            yposnu% = obj.CurrentY
            obj.Line (deelkolwidth%, savy)-(deelkolwidth%, yposnu%)
            obj.Line (deelgeldpos - obj.TextWidth("Geld") - 400, yLinePos%)-(deelgeldpos - obj.TextWidth("Geld") - 400, yposnu%)
            obj.Line (deelgeldpos - obj.TextWidth("Geld") - 400 + deelkolwidth%, yLinePos%)-(deelgeldpos - obj.TextWidth("Geld") - 400 + deelkolwidth%, yposnu%)
            obj.Line (leftmarge, yposnu%)-(obj.ScaleWidth + obj.ScaleLeft * 2, yposnu%)
        End If
        .Close
    End With
    obj.Print
End Sub

Function DeelnResultSql(alfabet As Boolean, wedNum As Integer) As String
Dim sql As String
    sql = "Select deelnemID, bijnaam, wednum,"
    sql = sql & " deelnempnt.*"
    sql = sql & " from deelnempnt, pooldeelnems"
    sql = sql & " WHERE pooldeelnems!deelnemID = deelnempnt.deelnid"
    sql = sql & " AND pooldeelnems!thisPool = " & thisPool
    sql = sql & " AND wednum = " & wedNum
    If alfabet Then
        sql = sql & " ORDER BY bijnaam"
    Else
        sql = sql & " ORDER BY grandtotaal DESC, bijnaam ASC"
    End If
    DeelnResultSql = sql
End Function


Private Sub lstCompetitorPools_Click()
    Me.Option4.value = True
End Sub

Private Sub txtVoorwed_Change()
chkTxtValue Me.txtVoorWed, Me.vscrlVoor
tmwed = val(txtVoorWed.Text)
End Sub

Private Sub txtTMwed_Change()
chkTxtValue Me.txtTMwed, Me.vscrlTM
tmwed = val(txtTMwed.Text)
End Sub

Private Sub Vet(Aan As Boolean)
    obj.FontBold = Aan
End Sub

Private Sub voetregel()
Dim W%
Dim i As Double
Dim fontnaam As String
    obj.ForeColor = RGB(0, 51, 0)
    W% = obj.DrawWidth
    obj.DrawWidth = 1
    FontGr 8
    Ital True
    Vet False
    fontnaam = obj.FontName
    obj.FontName = "Garamond"
    obj.CurrentY = obj.ScaleHeight - obj.TextHeight("w")
    voethoog = obj.CurrentY - obj.TextHeight("w")
    Y% = obj.CurrentY
    obj.Line (0, Y% - 15 * afdrratio)-(obj.ScaleWidth, Y% - 15 * afdrratio)
    obj.CurrentY = Y%
    Centreer "© 2004-" & Year(Now) & " jota computer assistentie"
    obj.FontName = fontnaam
    obj.Print
    FontGr 12
    Vet False
    Ital False
    Y% = obj.CurrentY + 50 * afdrratio
    'obj.Line (0, y%)-(obj.ScaleWidth, y%)
    obj.ForeColor = vbBlack
    obj.DrawWidth = 1
End Sub



Sub deelnemWedsInfo(inclpnt As Boolean)
Dim infostr As String
Dim pntToto As Integer
Dim pntRust As Integer
Dim pntEind As Integer
Dim pntDp As Integer
pntToto = getPntToek("toto goed")
pntRust = getPntToek("ruststand goed")
pntEind = getPntToek("eindstand goed")
pntDp = getPntToek("doelpunten op een dag")
infostr = "Samenstelling punten: rust goed "
If inclpnt Then infostr = infostr & pntRust & " pnt"
infostr = infostr & ", eindstand goed "
If inclpnt Then infostr = infostr & pntEind & " pnt"
infostr = infostr & ", toto goed "
If inclpnt Then infostr = infostr & pntToto & " pnt, "
infostr = infostr & ", aantal doelpunten van de dag goed "
If inclpnt Then infostr = infostr & pntDp & " pnt"
FontGr 10
obj.CurrentX = (obj.ScaleWidth - obj.TextWidth(infostr)) / 2
obj.Print "Samenstelling punten: ";
Ital True
obj.Print "toto goed";
If inclpnt Then obj.Print pntToto; "pnt";
Ital False
obj.Print ", ";
obj.FontUnderline = True
obj.Print "rust goed";
If inclpnt Then obj.Print pntRust; "pnt";
obj.FontUnderline = False
obj.Print ", ";
Vet True
obj.Print " eindstand goed";
If inclpnt Then obj.Print pntEind; "pnt";
Vet False
obj.Print ", ";
obj.ForeColor = vbBlue
obj.Print "aantal doelpunten van de dag goed";
If inclpnt Then obj.Print pntDp; "pnt"
obj.ForeColor = 1
obj.CurrentY = obj.CurrentY + 50


End Sub

Sub DeelnemWeds()
'print de deelnemers en hun punten per wedstrijd
Dim rsDeeln As New ADODB.Recordset
Dim rsDeelnPnt As New ADODB.Recordset
Dim rsWeds As New ADODB.Recordset
Dim sqlstr As String
Dim xpos As Integer
Dim posX() As Integer
Dim X As Integer
Dim i As Integer
Dim topY As Integer
Dim botY As Integer
Dim topYpos As Integer
Dim kolwidth As Integer
Dim ttlKolWidth As Integer
Dim wedstrijd As String
Dim verttxtHeight 'de hoogte van de verticale text bovenin
Dim infostr As String
headerText = GetOrgNaam(thisPool) & " " & getTournamentInfo("toernooi") & " voetbalpool"
kop$ = "Punten t/m wedstrijd " & tmwed
InitPage False, True
obj.CurrentY = obj.CurrentY - 50
topYpos = obj.CurrentY
deelnemWedsInfo True 'druk de inforegel over de punten toekenning af
topY = obj.CurrentY
obj.Line (0, topY)-(obj.ScaleWidth - 50, topY)
FontGr 8
sqlstr = "SELECT pooldeelnems.deelnemID, pooldeelnems.bijnaam, deelnempnt.grandTotaal"
sqlstr = sqlstr & " FROM (pooldeelnems INNER JOIN deelnempnt ON pooldeelnems.deelnemID = deelnempnt.deelnID) "
sqlstr = sqlstr & " INNER JOIN toernschema ON deelnempnt.wedNum = toernschema.wedNum"
sqlstr = sqlstr & " Where pooldeelnems.thisPool = " & thisPool
sqlstr = sqlstr & " And toernschema.myNum = " & tmwed
sqlstr = sqlstr & " And toernschema.ksid = " & kampID
If Me.ScoreVolg(1) = True Then
    sqlstr = sqlstr & " order by grandtotaal DESC"
Else
    sqlstr = sqlstr & " order by bijnaam"
End If

rsDeeln.Open sqlstr, cn, adOpenStatic, adLockReadOnly
sqlstr = "Select * from qryweds where ksid=" & kampID
'sqlstr = sqlstr & " AND wednum <=" & tmWed
sqlstr = sqlstr & " order by mynum"
rsWeds.Open sqlstr, cn, adOpenStatic, adLockReadOnly
verttxtHeight = obj.TextWidth("1234567890123456789012345")
obj.CurrentY = verttxtHeight

obj.CurrentX = obj.TextWidth(Left(GetLangsteBijNaam, 15))
ReDim posX(1)
posX(1) = obj.CurrentX
With rsWeds
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            rot.Angle = 90
            obj.CurrentX = posX(UBound(posX))
            If !tm1 > "" Then
                wedstrijd = !tm1 & "-"
                If !tm2 > "" Then
                    wedstrijd = wedstrijd & !tm2
                Else
                    wedstrijd = wedstrijd & !code2
                End If
            Else
                wedstrijd = !code1 & "-"
                If !tm2 > "" Then
                    wedstrijd = wedstrijd & !tm2
                Else
                    wedstrijd = wedstrijd & !code2
                End If
            End If
            rot.PrintText !mynum & ": " & wedstrijd
            rot.Angle = 0
            xpos = obj.CurrentX + obj.TextWidth("99") * 1.2
            ReDim Preserve posX(UBound(posX) + 1)
            posX(UBound(posX)) = xpos
            rsWeds.MoveNext
            'Debug.Print UBound(posX), posX(UBound(posX))
        Loop
    End If
End With
rot.Angle = 90
obj.CurrentX = posX(UBound(posX))
rot.PrintText " pnt groepstand"

If getTournamentInfo("groepen") > 4 Then
    xpos = obj.CurrentX + obj.TextWidth("geld") * 1.2
    ReDim Preserve posX(UBound(posX) + 1)
    posX(UBound(posX)) = xpos
    rot.Angle = 90
    obj.CurrentX = posX(UBound(posX))
    rot.PrintText " 8e Finalisten"
End If
xpos = obj.CurrentX + obj.TextWidth("99") * 1.2
ReDim Preserve posX(UBound(posX) + 1)
posX(UBound(posX)) = xpos
rot.Angle = 90
obj.CurrentX = posX(UBound(posX))
rot.PrintText " Kw Finalisten"

xpos = obj.CurrentX + obj.TextWidth("99") * 1.2
ReDim Preserve posX(UBound(posX) + 1)
posX(UBound(posX)) = xpos
rot.Angle = 90
obj.CurrentX = posX(UBound(posX))
rot.PrintText " Hv Finalisten"

xpos = obj.CurrentX + obj.TextWidth("99") * 1.2
ReDim Preserve posX(UBound(posX) + 1)
posX(UBound(posX)) = xpos
rot.Angle = 90
obj.CurrentX = posX(UBound(posX))
rot.PrintText " Finalisten"

xpos = obj.CurrentX + obj.TextWidth("99") * 1.2
ReDim Preserve posX(UBound(posX) + 1)
posX(UBound(posX)) = xpos
rot.Angle = 90
obj.CurrentX = posX(UBound(posX))
rot.PrintText " Eindstand"

xpos = obj.CurrentX + obj.TextWidth("99") * 1.2
ReDim Preserve posX(UBound(posX) + 1)
posX(UBound(posX)) = xpos
rot.Angle = 90
obj.CurrentX = posX(UBound(posX))
rot.PrintText " Topscorers"

xpos = obj.CurrentX + obj.TextWidth("99") * 1.2
ReDim Preserve posX(UBound(posX) + 1)
posX(UBound(posX)) = xpos
rot.Angle = 90
obj.CurrentX = posX(UBound(posX))
rot.PrintText " Overigen"

xpos = obj.CurrentX + obj.TextWidth("99") * 1.2
ReDim Preserve posX(UBound(posX) + 1)
posX(UBound(posX)) = xpos
rot.Angle = 90
obj.CurrentX = posX(UBound(posX))
rot.PrintText " Totaal"

xpos = obj.CurrentX + obj.TextWidth("999") * 1.2
ReDim Preserve posX(UBound(posX) + 1)
posX(UBound(posX)) = xpos
rot.Angle = 90
obj.CurrentX = posX(UBound(posX))
rot.PrintText " positie"

xpos = obj.CurrentX + obj.TextWidth("99") * 1.2
ReDim Preserve posX(UBound(posX) + 1)
posX(UBound(posX)) = xpos
rot.Angle = 90
obj.CurrentX = posX(UBound(posX))
obj.CurrentY = verttxtHeight - obj.TextHeight("Geld")
'obj.Print " geld";

xpos = obj.CurrentX + obj.TextWidth("geld") * 1.2
obj.Print
topYpos = obj.CurrentY + 50
ReDim Preserve posX(UBound(posX) + 1)
posX(UBound(posX)) = xpos
obj.Line (0, topYpos)-(posX(UBound(posX)), topYpos)
obj.CurrentY = topYpos
obj.CurrentX = 0
kolwidth = posX(2) - posX(1)
botY = obj.CurrentY
pntFormat = "0;;\ ;-"

Do While Not rsDeeln.EOF
Dim naam As String
    naam = rsDeeln!bijnaam
   ' If InStr(naam, "Winner") > 0 Then Stop       1234567890
    Do While obj.TextWidth(naam) > obj.TextWidth("123456789012345")
        naam = Left(naam, Len(naam) - 1)
    Loop
    obj.Print naam;
    sqlstr = "SELECT toernschema.tijd, deelnemPnt.*, toernschema.gespeeld"
    sqlstr = sqlstr & " FROM deelnemPnt INNER JOIN toernschema ON deelnemPnt.wedNum = toernschema.wedNum"
    sqlstr = sqlstr & " Where toernschema.mynum <=" & tmwed
    sqlstr = sqlstr & " AND toernschema.ksid = " & kampID
    sqlstr = sqlstr & " AND deelnID = " & rsDeeln!deelnemID
    sqlstr = sqlstr & " ORDER BY toernschema.mynum"
    rsDeelnPnt.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    i = 0
    With rsDeelnPnt
        rot.Angle = 90
        Do While Not .EOF
            i = i + 1
            obj.CurrentX = posX(i) + (kolwidth - obj.TextWidth(Format(nz(!pnttotaal, 0), pntFormat))) / 2
'            rot.Angle = 0
            'If !pnttotaal = 7 Then Stop
            Ital nz(!pntToto, 0) <> 0
            Vet nz(!pntEind, 0) <> 0
            obj.FontUnderline = nz(!pntRust, 0) > 0
            If nz(!dpvddag, 0) > 0 Then
                obj.ForeColor = vbBlue
            End If
            obj.Print Format(nz(!pnttotaal, 0), pntFormat);
            Vet False
            Ital False
            obj.FontUnderline = False
            obj.ForeColor = 1
            
            .MoveNext
            rot.Angle = 90
        Loop
        If Not .RecordCount = 0 Then
            .MoveLast
            If !postotaal = 1 Then
                obj.ForeColor = &HC00000
                obj.FontBold = True
            Else
                obj.ForeColor = vbBlack
                obj.FontBold = False
            End If
            ttlKolWidth = posX(UBound(posX) - 10) - posX(UBound(posX) - 11)

            obj.CurrentX = posX(UBound(posX) - 11) + (ttlKolWidth - obj.TextWidth(Format(nz(!pntgrp, 0), pntFormat))) / 2
            obj.Print Format(nz(!pntgrp, 0), pntFormat);
            ttlKolWidth = posX(UBound(posX) - 9) - posX(UBound(posX) - 10)
            If getTournamentInfo("groepen") > 4 Then
                obj.CurrentX = posX(UBound(posX) - 10) + (ttlKolWidth - obj.TextWidth(Format(nz(!pnt8fin, 0), pntFormat))) / 2
                obj.Print Format(nz(!pnt8fin, 0), pntFormat);
                ttlKolWidth = posX(UBound(posX) - 8) - posX(UBound(posX) - 9)
            End If
            obj.CurrentX = posX(UBound(posX) - 9) + (ttlKolWidth - obj.TextWidth(Format(nz(!pntkwfin, 0), pntFormat))) / 2
            obj.Print Format(nz(!pntkwfin, 0), pntFormat);
            ttlKolWidth = posX(UBound(posX) - 7) - posX(UBound(posX) - 8)
            obj.CurrentX = posX(UBound(posX) - 8) + (ttlKolWidth - obj.TextWidth(Format(nz(!pnthvfin, 0), pntFormat))) / 2
            obj.Print Format(nz(!pnthvfin, 0), pntFormat);
            ttlKolWidth = posX(UBound(posX) - 6) - posX(UBound(posX) - 7)
            obj.CurrentX = posX(UBound(posX) - 7) + (ttlKolWidth - obj.TextWidth(Format(nz(!pntfin, 0) + nz(!pntklfin, 0), pntFormat))) / 2
            obj.Print Format(nz(!pntfin, 0) + nz(!pntklfin, 0), pntFormat);
            ttlKolWidth = posX(UBound(posX) - 5) - posX(UBound(posX) - 6)
            obj.CurrentX = posX(UBound(posX) - 6) + (ttlKolWidth - obj.TextWidth(Format(!pntuitslnaklfin + !pntuitsl, pntFormat))) / 2
            obj.Print Format(!pntuitslnaklfin + !pntuitsl, pntFormat);
            ttlKolWidth = posX(UBound(posX) - 4) - posX(UBound(posX) - 5)
            obj.CurrentX = posX(UBound(posX) - 5) + (ttlKolWidth - obj.TextWidth(Format(nz(!pntTopsc, 0) + nz(!pntOverig, 0), pntFormat))) / 2
            obj.Print Format(nz(!pntTopsc, 0), pntFormat);
            ttlKolWidth = posX(UBound(posX) - 3) - posX(UBound(posX) - 4)
            obj.CurrentX = posX(UBound(posX) - 4) + (ttlKolWidth - obj.TextWidth(Format(nz(!pntTopsc, 0) + nz(!pntOverig, 0), pntFormat))) / 2
            obj.Print Format(nz(!pntOverig, 0), pntFormat);
            ttlKolWidth = posX(UBound(posX) - 2) - posX(UBound(posX) - 3)
            obj.CurrentX = posX(UBound(posX) - 3) + (ttlKolWidth - obj.TextWidth(Format(nz(!grandtotaal, 0), pntFormat))) / 2
            obj.Print Format(nz(!grandtotaal, 0), pntFormat);
            ttlKolWidth = posX(UBound(posX) - 1) - posX(UBound(posX) - 2)
            obj.CurrentX = posX(UBound(posX) - 2) + (ttlKolWidth - obj.TextWidth(Format(nz(!postotaal, 0), pntFormat))) / 2
            obj.Print Format(nz(!postotaal, 0), pntFormat);
            obj.CurrentX = posX(UBound(posX)) - obj.TextWidth(Format(nz(!geldttl, 0), "currency"))
            obj.ForeColor = vbBlack
            obj.FontItalic = False
            obj.FontBold = False
'            obj.Print Format(nz(!geldttl, 0), "currency");
        End If
        obj.Print
    End With
    obj.Line (0, obj.CurrentY + 10)-(posX(UBound(posX)), obj.CurrentY + 10)
    obj.CurrentY = obj.CurrentY + 10
    obj.CurrentX = 0
    botY = obj.CurrentY
'    If rsDeeln.AbsolutePosition = 67 Then Stop
    If botY >= voethoog And rsDeeln.AbsolutePosition < rsDeeln.RecordCount Then
        'nieuwe pagina
        'eerste de lijntjes
        For i = 1 To UBound(posX)
            obj.Line (posX(i), topY)-(posX(i), botY)
        Next
        i = 0
        DoNewPage False, True
        obj.CurrentY = obj.CurrentY - 50
        topYpos = obj.CurrentY
        deelnemWedsInfo True 'druk de inforegel over de punten toekenning af
        topY = obj.CurrentY
        obj.Line (0, topY)-(obj.ScaleWidth - 50, topY)
        FontGr 8
        obj.CurrentY = verttxtHeight
        obj.CurrentX = obj.TextWidth("123456789012345")
        With rsWeds
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    Set rot.Device = obj
                    i = i + 1
                    rot.Angle = 90
                    obj.CurrentX = posX(i)
                    If !tm1 > "" Then
                        rot.PrintText !mynum & ": " & !tm1 & "-" & !tm2
                    Else
                        rot.PrintText !mynum & ": " & !code1 & "-" & !code2
                    End If
                    rot.Angle = 0
                    .MoveNext
                Loop
            End If
        End With
        rot.Angle = 90
        If getTournamentInfo("groepen") > 4 Then
            X = 11
        Else
            X = 10
        End If
        obj.CurrentX = posX(UBound(posX) - X)
        X = X - 1
        rot.PrintText " pnt groepstand"
        If getTournamentInfo("groepen") > 4 Then
            obj.CurrentX = posX(UBound(posX) - X)
            X = X - 1
            rot.PrintText " 8e Finalisten"
        End If
        obj.CurrentX = posX(UBound(posX) - X)
        X = X - 1
        rot.PrintText " Kw Finalisten"
        obj.CurrentX = posX(UBound(posX) - X)
        X = X - 1
        rot.PrintText " Hv Finalisten"
        obj.CurrentX = posX(UBound(posX) - X)
        X = X - 1
        rot.PrintText " Finalisten"
        obj.CurrentX = posX(UBound(posX) - X)
        X = X - 1
        rot.PrintText " Eindstand"
        obj.CurrentX = posX(UBound(posX) - X)
        X = X - 1
        rot.PrintText " Topscorers"
        obj.CurrentX = posX(UBound(posX) - X)
        X = X - 1
        rot.PrintText " Overigen"
        obj.CurrentX = posX(UBound(posX) - X)
        X = X - 1
        rot.PrintText " Totaal"
        obj.CurrentX = posX(UBound(posX) - X)
        X = X - 1
        rot.PrintText " positie"
        obj.CurrentX = posX(UBound(posX) - X)
        obj.CurrentY = verttxtHeight ' - obj.TextHeight("Geld")
 '       obj.Print " geld"
        topYpos = obj.CurrentY + 50
        obj.Line (0, topYpos)-(posX(UBound(posX)), topYpos)
        obj.CurrentY = topYpos
        obj.CurrentX = 0
        i = i + 1
    End If
    rsDeeln.MoveNext
    rsDeelnPnt.Close
Loop
rsDeeln.Close
For i = 1 To UBound(posX)
    obj.Line (posX(i), topY)-(posX(i), botY)
Next
i = 0
Set rsDeeln = Nothing
Set rsDeelnPnt = Nothing
End Sub

Sub DeelnemWedsPos()
Dim rsDeeln As New ADODB.Recordset
Dim rsDeelnPnt As New ADODB.Recordset
Dim rsWeds As New ADODB.Recordset
Dim sqlstr As String
Dim xpos As Integer
Dim posX() As Integer
Dim i As Integer
Dim topY As Integer
Dim botY As Integer
Dim topYpos As Integer
Dim kolwidth As Integer
Dim ttlKolWidth As Integer
Dim verttxtHeight 'de hoogte van de verticale text bovenin
Dim infostr As String
headerText = GetOrgNaam(thisPool) & " " & getTournamentInfo("toernooi") & " voetbalpool"
kop$ = "Positie in de pool na elke wedstrijd"
InitPage False, True
obj.CurrentY = obj.CurrentY - 50
topYpos = obj.CurrentY
deelnemWedsInfo False 'druk de inforegel over de punten toekenning af
topY = obj.CurrentY
obj.Line (0, topY)-(obj.ScaleWidth - 50, topY)
FontGr 8
sqlstr = "SELECT pooldeelnems.deelnemID, pooldeelnems.bijnaam, deelnempnt.grandTotaal"
sqlstr = sqlstr & " FROM (pooldeelnems INNER JOIN deelnempnt ON pooldeelnems.deelnemID = deelnempnt.deelnID) "
sqlstr = sqlstr & " INNER JOIN toernschema ON deelnempnt.wedNum = toernschema.wedNum"
sqlstr = sqlstr & " Where pooldeelnems.thisPool = " & thisPool
sqlstr = sqlstr & " And toernschema.myNum = " & tmwed
sqlstr = sqlstr & " And toernschema.ksid = " & kampID
If Me.ScoreVolg(1) = True Then
    sqlstr = sqlstr & " order by grandtotaal DESC"
Else
    sqlstr = sqlstr & " order by bijnaam"
End If

rsDeeln.Open sqlstr, cn, adOpenStatic, adLockReadOnly
sqlstr = "Select * from qryweds where ksid=" & kampID
'sqlstr = sqlstr & " AND wednum <=" & tmWed
sqlstr = sqlstr & " order by mynum"
rsWeds.Open sqlstr, cn, adOpenStatic, adLockReadOnly
verttxtHeight = obj.TextWidth("1234567890123456789012345")
obj.CurrentY = verttxtHeight
obj.CurrentX = obj.TextWidth("1234567890")
ReDim posX(1)
posX(1) = obj.CurrentX
With rsWeds
    Do While Not .EOF
        rot.Angle = 90
        obj.CurrentX = posX(UBound(posX))
        If !tm1 > "" Then
            rot.PrintText !mynum & ": " & !tm1 & "-" & !tm2
        Else
            rot.PrintText !mynum & ": " & !code1 & "-" & !code2
        End If
        rot.Angle = 0
        xpos = obj.CurrentX + obj.TextWidth("99") * 1.3
        ReDim Preserve posX(UBound(posX) + 1)
        posX(UBound(posX)) = xpos
        .MoveNext
    Loop
End With

'obj.Print
topYpos = obj.CurrentY + 50
ReDim Preserve posX(UBound(posX) + 1)
posX(UBound(posX)) = xpos
obj.Line (0, topYpos)-(posX(UBound(posX)), topYpos)
obj.CurrentY = topYpos
obj.CurrentX = 0
kolwidth = posX(2) - posX(1)
botY = obj.CurrentY
pntFormat = "0;;\ ;-"

Do While Not rsDeeln.EOF
    obj.Print rsDeeln!bijnaam;
    sqlstr = "SELECT toernschema.tijd, deelnemPnt.*, toernschema.gespeeld"
    sqlstr = sqlstr & " FROM deelnemPnt INNER JOIN toernschema ON deelnemPnt.wedNum = toernschema.wedNum"
    sqlstr = sqlstr & " Where toernschema.mynum <=" & tmwed
    sqlstr = sqlstr & " AND toernschema.ksid = " & kampID
    sqlstr = sqlstr & " AND deelnID = " & rsDeeln!deelnemID
    sqlstr = sqlstr & " ORDER BY toernschema.mynum"
    rsDeelnPnt.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    i = 0
    With rsDeelnPnt
        rot.Angle = 90
        Do While Not .EOF
            i = i + 1
            obj.CurrentX = posX(i) + (kolwidth - obj.TextWidth(Format(nz(!postotaal, 0), pntFormat))) / 2
            Ital nz(!pntToto, 0) <> 0
            Vet nz(!pntEind, 0) <> 0
            obj.FontUnderline = nz(!pntRust, 0) > 0
            If nz(!dpvddag, 0) > 0 Then
                obj.ForeColor = vbBlue
            End If
            obj.Print Format(nz(!postotaal, 0), pntFormat);
            Vet False
            Ital False
            obj.FontUnderline = False
            obj.ForeColor = 1
            
            .MoveNext
            rot.Angle = 90
        Loop
        obj.Print
    End With
    obj.Line (0, obj.CurrentY + 10)-(posX(UBound(posX)), obj.CurrentY + 10)
    obj.CurrentY = obj.CurrentY + 10
    obj.CurrentX = 0
    botY = obj.CurrentY
    If botY >= voethoog Then
        'nieuwe pagina
        'eerste de lijntjes
        For i = 1 To UBound(posX)
            obj.Line (posX(i), topY)-(posX(i), botY)
        Next
        i = 0
        DoNewPage False, True
        obj.CurrentY = obj.CurrentY - 50
        topYpos = obj.CurrentY
        deelnemWedsInfo False 'druk de inforegel over de punten toekenning af
        topY = obj.CurrentY
        obj.Line (0, topY)-(obj.ScaleWidth - 50, topY)
        FontGr 8
        obj.CurrentY = verttxtHeight
        obj.CurrentX = obj.TextWidth("123456789012345")
        With rsWeds
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    Set rot.Device = obj
                    i = i + 1
                    rot.Angle = 90
                    obj.CurrentX = posX(i)
                    If !tm1 > "" Then
                        rot.PrintText !mynum & ": " & !tm1 & "-" & !tm2
                    Else
                        rot.PrintText !mynum & ": " & !code1 & "-" & !code2
                    End If
                    rot.Angle = 0
                    .MoveNext
                Loop
            End If
        End With
        'obj.Print
        topYpos = obj.CurrentY + 50
        obj.Line (0, topYpos)-(posX(UBound(posX)), topYpos)
        obj.CurrentY = topYpos
        obj.CurrentX = 0
        i = i + 1
    End If
    rsDeeln.MoveNext
Loop
For i = 1 To UBound(posX)
    obj.Line (posX(i), topY)-(posX(i), botY)
Next
i = 0


End Sub

Sub AfdrukVoorspWed(wedNum As Integer)
Dim sqlstr As String
Dim rs As New ADODB.Recordset
Dim rsDeeln As New ADODB.Recordset
Dim cloneRS As ADODB.Recordset
Dim zoekstr As String
Dim kopje As String
Dim xpos As Integer
Dim cols(4) As Integer
Dim naampos
Dim rijen As Integer
Dim rijnu As Integer
Dim yStart As Integer
Dim lineXstart As Integer
Dim lineYstart As Integer
Dim lineXend As Integer
Dim lineYend As Integer
Dim koppos(3) As Integer
Dim col As Integer
Dim i As Integer
wedNum = GetWedNum(wedNum)
    headerText = GetOrgNaam(thisPool) & " " & getTournamentInfo("toernooi") & " voetbalpool" & " - Voorspelling"
    If Not Me.optPortrait Then
        cols(0) = 0
        cols(1) = obj.ScaleWidth / 4
        cols(2) = obj.ScaleWidth / 2
        cols(3) = obj.ScaleWidth / 4 * 3
        cols(4) = obj.ScaleWidth
        col = 4
    Else
        cols(0) = 0
        cols(1) = obj.ScaleWidth / 3
        cols(2) = obj.ScaleWidth / 3 * 2
        cols(3) = obj.ScaleWidth
        cols(4) = obj.ScaleWidth
        col = 3
    End If
    kopje = Format(GetWedInfo(wedNum, "datum"), "ddd d mmm") & " "
    kopje = kopje & Format(GetWedInfo(wedNum, "tijd"), "HH:MM") & ": "
    kopje = kopje & GetWedInfo(wedNum, "naam1") & " vs " & GetWedInfo(wedNum, "naam2")
    kop$ = "Wedstrijd " & GetMyNum(wedNum) & ": " & kopje
    InitPage False, True
    
    obj.Print
    koppos(0) = 50
    koppos(1) = obj.TextWidth("0-000")
    koppos(2) = koppos(1) + obj.TextWidth("0-000")
    koppos(3) = koppos(2) + obj.TextWidth("0-000")
    obj.ForeColor = RGB(0, 51, 0)
    For i = 0 To col - 1
        obj.CurrentX = cols(i) + koppos(0)
        obj.Print "Rust";
        obj.CurrentX = cols(i) + koppos(1)
        obj.Print "Eind";
        obj.CurrentX = cols(i) + koppos(2)
        obj.Print "Toto";
        obj.CurrentX = cols(i) + koppos(3)
        obj.Print "Wie";
    Next
    obj.ForeColor = 0
    obj.Print
    yStart = obj.CurrentY
    sqlstr = "SELECT e1, e2, r1,r2,toto, wednum "
    sqlstr = sqlstr & " FROM voorspelling_uitsl INNER JOIN "
    sqlstr = sqlstr & " pooldeelnems ON voorspelling_uitsl.deelnem = pooldeelnems.deelnemID"
    sqlstr = sqlstr & " GROUP BY e1, e2, r1, r2, toto, wednum, poolid"
    sqlstr = sqlstr & " HAVING wednum=" & wedNum
    sqlstr = sqlstr & " AND pooldeelnems.poolid= " & thisPool
    sqlstr = sqlstr & " ORDER BY r1,r2,e1,e2,toto"
    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    sqlstr = "SELECT e1, e2, r1,r2,toto, wednum, bijnaam "
    sqlstr = sqlstr & " FROM voorspelling_uitsl INNER JOIN "
    sqlstr = sqlstr & " pooldeelnems ON voorspelling_uitsl.deelnem = pooldeelnems.deelnemID"
    sqlstr = sqlstr & " WHERE wednum = " & wedNum
    sqlstr = sqlstr & " AND poolid = " & thisPool
    sqlstr = sqlstr & " ORDER BY bijnaam"
    rsDeeln.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    rsDeeln.MoveLast
    rijen = Int(rsDeeln.RecordCount / col + 0.5) + 1
    rsDeeln.MoveFirst
    rs.MoveFirst
    i = 0
    Do While Not rs.EOF
        Set cloneRS = rsDeeln.Clone
        zoekstr = "e1 = " & rs!e1
        zoekstr = zoekstr & " and e2 = " & rs!e2
        zoekstr = zoekstr & " and r1 = " & rs!r1
        zoekstr = zoekstr & " and r2 = " & rs!r2
        zoekstr = zoekstr & " and toto = " & rs!toto
        cloneRS.Filter = zoekstr
        If cloneRS.EOF Or cloneRS.BOF Then
            rsDeeln.MoveLast
            rsDeeln.MoveNext
        End If
        'rsDeeln.Find zoekstr, , , 0
        obj.CurrentX = cols(i)
        lineXstart = obj.CurrentX
        lineYstart = obj.CurrentY
        obj.CurrentX = cols(i) + koppos(0)
        obj.Print rs!r1 & "-" & rs!r2;
        obj.CurrentX = cols(i) + koppos(1)
        Vet True
        obj.Print rs!e1 & "-" & rs!e2;
        Vet False
        obj.CurrentX = cols(i) + koppos(2)
        obj.Print rs!toto;
        cloneRS.MoveFirst
        Do While Not cloneRS.EOF
            obj.CurrentX = cols(i) + koppos(3)
            obj.Print cloneRS!bijnaam
            rijnu = rijnu + 1
            cloneRS.MoveNext
        Loop
        lineXend = cols(i + 1) - 100
        lineYend = obj.CurrentY
        obj.Line (lineXstart, lineYstart)-(lineXend, lineYend), , B
        rs.MoveNext
        If rijnu >= rijen Then
            i = i + 1
            obj.CurrentY = yStart
            rijnu = 0
        End If
        cloneRS.Close
        Set cloneRS = Nothing
    Loop
    rs.Close
    rsDeeln.Close
    Set rs = Nothing
    Set rsDeeln = Nothing
End Sub
Sub SetForeCol(kl As Long)
Dim r As Integer
Dim g As Integer
Dim b As Integer
    r = &HFF& And kl
    g = (&HFF00& And kl) \ 256
    b = (&HFF0000 And kl) \ 65536
    If r * 0.3 + g * 0.59 + b * 0.11 < 128 Then
        obj.ForeColor = vbWhite
    Else
        obj.ForeColor = vbBlack
    End If

End Sub

Sub MakeColors()
Dim i As Integer
Dim a As Integer
Dim c As Integer
Dim r As Integer
Dim g As Integer
Dim b As Integer
Dim klCol As Collection
Dim forecol As Integer
    Set klCol = New Collection
    For r = 0 To 255 Step 63
       For g = 0 To 255 Step 63
         For b = 0 To 255 Step 63
            klCol.Add RGB(r, g, b)
         Next
       Next
    Next
For a = 0 To 64
    i = Int(Rnd() * klCol.Count) + 1
    kleur(a) = klCol(i)
    klCol.Remove i
Next
End Sub
