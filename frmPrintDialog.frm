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
      Left            =   240
      ScaleHeight     =   2145
      ScaleWidth      =   3315
      TabIndex        =   35
      Top             =   240
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
         Left            =   30
         MultiSelect     =   1  'Simple
         TabIndex        =   38
         Top             =   30
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
         BuddyDispid     =   196617
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
      Begin VB.CommandButton btnClose 
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
         Value           =   1
         BuddyControl    =   "txtForMatch"
         BuddyDispid     =   196632
         OrigLeft        =   2520
         OrigRight       =   2775
         OrigBottom      =   375
         Max             =   24
         Min             =   1
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
   Begin VB.PictureBox picToMatch 
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
         Value           =   1
         BuddyControl    =   "txtToMatch"
         BuddyDispid     =   196635
         OrigLeft        =   2520
         OrigTop         =   30
         OrigRight       =   2775
         OrigBottom      =   405
         Max             =   24
         Min             =   1
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
         Index           =   5
         Left            =   90
         TabIndex        =   26
         Top             =   2135
         Width           =   2670
      End
      Begin VB.OptionButton optPrintDoc 
         Appearance      =   0  'Flat
         Caption         =   "Voorspellingen"
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
         Index           =   4
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
         Top             =   480
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
         Index           =   8
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
         Index           =   7
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
         Index           =   2
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

'global objects
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Dim rotater As rotator

Dim printPrev As printPreview

'added function to print color on dark background
Private Declare Function SetBkMode Lib "gdi32" _
 (ByVal hdc As Long, ByVal nBkMode As Long) As Long

Private Declare Function GetBkMode Lib "gdi32" _
 (ByVal hdc As Long) As Long

Private Const TRANSPARENT = 1
Private Const OPAQUE = 2
Private iBKMode As Long

'gobals for every print
Dim headerText ' top of the page
Dim heading1 As String 'top of the section
Dim toMatch As Integer  'to store the matchorder number till where we should print
Dim currentMatch As Integer 'the currentMatch ordernumber

'OLD STUFF
Dim KolHeight As Integer
Dim kolwidth As Integer
Dim kol As Integer
Dim y As Integer
'voor de printFavourites afdruk
Dim favYpos As Integer
Dim favXpos As Integer

Dim headingFont As String
Dim textFont As String

Dim x As Integer

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
Dim printobj As Object
Dim maxY As Integer 'voor afdrukken van printFavourites

Dim kleur(64) As Long 'voor grafiek

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'tidy up
If Not rs Is Nothing Then
  If (rs.State And adStateOpen) = adStateOpen Then rs.Close
  Set rs = Nothing
End If
If Not cn Is Nothing Then
  If (cn.State And adStateOpen) = adStateOpen Then cn.Close
  Set cn = Nothing
End If
If Not rotater Is Nothing Then
  Set rs = Nothing
End If
If Not printPrev Is Nothing Then
  Set printPrev = Nothing
End If

End Sub

Private Sub optPrintDoc_Click(Index As Integer)
Dim i As Integer
Me.picCompetitorList.Visible = False
Me.optPrintDoc(Index).value = True
Select Case Index
  Case 0
    Me.picToMatch.Visible = False
    Me.picVoorWed.Visible = False
    Me.picVolgorde.Visible = False
    Me.optPortrait.value = True
    Me.picCompetitorList.Visible = False
   ' Me.chkDblSide.Value = 1

  Case 1
   'deelnemers met voorspellingen
    Me.picToMatch.Visible = False
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
    picToMatch.Visible = True
    Me.optPortrait.value = True
    'Me.vscrlTM = GetMyNum(GetLastPlayed)
    If toMatch > 0 Then
        Me.upDnToMatch.SetFocus
    End If
  Case 3
    ' printFavourites
    Me.picToMatch.Visible = False
    Me.picVoorWed.Visible = False
    Me.picVolgorde.Visible = False
    Me.optPortrait.value = True
    Me.picCompetitorList.Visible = False
  Case 4
    'Stand in toernooi
    'score/ stand in de pool
    Me.picToMatch.Visible = False
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
    Me.picToMatch.Visible = True
    Me.optLandscape.value = True
    Me.ScoreVolg(1) = True
    'Me.vscrlTM = GetMyNum(GetLastPlayed())
    DoEvents
    toMatch = Me.upDnToMatch
  Case 6
    'punten per wedstrijd
    picVolgorde.Visible = True
    picVoorWed.Visible = False
    picToMatch.Visible = True
    'Me.vscrlTM = GetMyNum(GetLastPlayed())
    toMatch = Me.upDnToMatch
    DoEvents
    Me.picCompetitorList.Visible = False  'getTournamentInfo("groepen")
    Me.optLandscape.value = getTournamentInfo("tournamentGroupCount", cn) > 4
    Me.optPortrait.value = Not Me.optLandscape.value
  Case 7
    'voorspelling per wedstrijd
    picVolgorde.Visible = False
    picVoorWed.Visible = True
    picToMatch.Visible = False
    Me.optPortrait.value = True
    Me.optLandscape.value = False
    'Me.vscrlVoor = GetMyNum(GetLastPlayed()) + 1
    Me.picCompetitorList.Visible = False
  Case 8
    'samenvatting stand
    'Stand in toernooi
    'score/ stand in de pool
    Me.picToMatch.Visible = True
    Me.picVoorWed.Visible = False
    Me.picVolgorde.Visible = True
    Me.optLandscape.value = True
    Me.picCompetitorList.Visible = False
    'Me.vscrlTM = GetMyNum(GetLastPlayed())
  End Select
End Sub

Sub horline(kleur As Integer)
    printobj.Line (0, printobj.CurrentY)-(printobj.ScaleWidth - 50, printobj.CurrentY), kleur
End Sub

Sub printPoolForms()
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
    printobj.FillStyle = vbFSTransparent
    headerText = getOrganisation(cn) & getTournamentInfo("description", cn) & " voetbalpool"
    heading1 = "Inschrijfformulier     inleg: " & Format(getPoolInfo("poolCost", cn), "currency")
    printobj.FontName = headingFont
    InitPage False, True
    printobj.Print
    fontSizing 12
    topY = printobj.CurrentY
    printobj.ForeColor = vbBlack
    printobj.FontBold = False
    fontSizing 12
    printobj.CurrentY = topY
    fontSizing 18
    printobj.Line (0, topY - 200)-(printobj.ScaleWidth + 2 * printobj.ScaleLeft, topY + printobj.TextHeight("WW") * 4 + 200), , B
    printobj.Print
    xpos = printobj.CurrentX + 200
    printobj.CurrentY = topY
    printobj.CurrentX = xpos
    printobj.Print "Naam: ....................................................... Telefoon....................................."
    printobj.CurrentY = topY + printobj.TextWidth("WW")
    printobj.CurrentX = xpos
    printobj.Print "Adres: ....................................................... Plaats.........................................."
    printobj.CurrentY = topY + printobj.TextWidth("WW") * 2
    printobj.CurrentX = xpos
    printobj.Print "Email: ....................................................... Betaald ";
    xpos = printobj.CurrentX
    ypos = printobj.CurrentY
    printobj.DrawWidth = 3
    printobj.Line (xpos, ypos)-(xpos + printobj.TextWidth("W"), ypos + printobj.TextHeight("W")), , B
    printobj.DrawWidth = 1
    printobj.CurrentY = ypos
    printobj.CurrentX = printobj.CurrentX + 30
    printobj.Print " bij............................"
    fontSizing 4
    printobj.Print
    'sqlstr = "Select * from poolpnt Where thisPool = " & thisPool
    'rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    fontSizing 16
    printobj.FontBold = True
    printobj.ForeColor = vbBlue
    printobj.Print "Instructies"
    fontSizing 11
    printobj.FontBold = False
    printobj.ForeColor = vbBlack
    printobj.Print "Hier onder (en op de achterkant) kun je voorspellingen invoeren voor de "; getTournamentInfo("description", cn);
    printobj.Print " van "; Format(getTournamentInfo("tournamentstartdate", cn), "d MMMM yyyy"); " tot "; Format(getTournamentInfo("tournamentEnddate", cn), "d MMMM yyyy")
    printobj.Print "Voor elke juiste voorspelling krijg je punten, bij de verschillende onderdelen staat hoeveel."
    printobj.Print "De voorspellingen hoeven niet te kloppen, bij een uitslag kun je bijvoorbeeld 1-0 bij de rust, 0-2 bij de eindstand en een 3 "
    printobj.Print "bij de toto invullen. Of je kunt een team dat je uitgeschakeld hebt in een volgende ronde toch weer opnemen."
    If getTournamentInfo("tournamentGroupCount", cn) = 6 And getTournamentInfo("tournamentTeamCount", cn) = 24 Then ' de vier beste derde plaatsen naar kwart finales
      printobj.Print "De beste 4 derde plaatsen kwalificeren zich ook voor de 8e finales."
    End If
    fontSizing 16
    printobj.FontBold = True
    printobj.ForeColor = vbBlue
    printobj.Print "Prijzen"
    fontSizing 11
    printobj.FontBold = False
    printobj.ForeColor = vbBlack
    'printObj.Print "Na de finale worden de hoofdprijzen te verdeeld, maar ook per dag zijn er geldprijzen te winnen."
    printobj.FontBold = True
    printobj.Print "-  Per dag";
    printobj.FontBold = False
    printobj.Print " zijn de volgende geldprijzen te verdienen:"
    printobj.Print "  -  ";
    printobj.Print "Degene die op ";
    printobj.FontItalic = True
    printobj.Print "één dag de meeste punten";
    printobj.FontItalic = False
    printobj.Print " heeft verzameld, ";
    printobj.Print " krijgt daarvoor ";
    printobj.FontBold = True
    printobj.Print Format(getPoolInfo("prizeHighDayScore", cn), "currency")
    printobj.FontBold = False
    printobj.Print "  -  ";
    printobj.Print "Degene die na een dag in de ";
    printobj.FontItalic = True
    printobj.Print "totaalstand bovenaan";
    printobj.FontItalic = False
    printobj.Print " staat, ";
    printobj.Print " krijgt daarvoor ";
    printobj.FontBold = True
    printobj.Print Format(getPoolInfo("prizeHighDayPosition", cn), "currency")
    printobj.FontBold = False
    printobj.Print "  -  ";
    printobj.Print "Degene die na een dag in de ";
    printobj.FontItalic = True
    printobj.Print "totaalstand onderaan";
    printobj.FontItalic = False
    printobj.Print " staat, ";
    printobj.Print " krijgt daarvoor als troost ";
    printobj.FontBold = True
    printobj.Print Format(getPoolInfo("prizeLowDayPosition", cn), "currency")
    printobj.FontBold = False
    printobj.Print "  -  ";
    xpos = printobj.CurrentX
    printobj.Print "De punten voor de finalerondes tellen mee voor de dagprijs op de dag dat de teams bekend zijn"
    printobj.CurrentX = xpos
    printobj.Print "De punten voor de eindstand, topscorers en aantallen tellen op de dag van de finale mee voor de dagprijs"
    printobj.Print "-  ";
    printobj.FontBold = True
    printobj.Print "Aan het eind van het toernooi";
    printobj.FontBold = False
    printobj.Print " zijn de volgende geldprijzen te verdienen:"
    amount = getPoolInfo("prizeLowFinalPosition", cn)
    If amount > 0 Then
        printobj.Print "  -  ";
        xpos = printobj.CurrentX
        printobj.Print "De ";
        printobj.FontItalic = True
        printobj.ForeColor = vbRed
        printobj.Print "rode lantaarn";
        printobj.ForeColor = vbBlack
        printobj.FontItalic = False
        printobj.Print " ontvangt als troostprijs "; Format(amount, "currency")
    End If
    
    printobj.Print "  -  ";
    xpos = printobj.CurrentX
    printobj.Print "De ";
    printobj.FontItalic = True
    printobj.Print "hoogste";
    printobj.FontItalic = False
    printobj.Print " deelnemers in de totaalstand krijgen de volgende prijzen:"
    printobj.CurrentX = xpos
    
    printobj.Print "1e pl: ";
    printobj.FontBold = True
    printobj.Print Format(getPoolInfo("prizePercentageFirst", cn) / 100, "0%");
    printobj.FontBold = False
    amount = getPoolInfo("prizePercentageSecond", cn)
    If amount > 0 Then
        printobj.Print ", 2e pl: ";
        printobj.FontBold = True
        printobj.Print Format(amount / 100, "0%");
        printobj.FontBold = False
    End If
    amount = getPoolInfo("prizePercentageThird", cn)
    If amount > 0 Then
        printobj.Print ", 3e pl: ";
        printobj.FontBold = True
        printobj.Print Format(amount / 100, "0%");
        printobj.FontBold = False
    End If
    amount = getPoolInfo("prizePercentageFourth", cn)
    If amount > 0 Then
        printobj.Print ", 4e pl: ";
        printobj.FontBold = True
        printobj.Print Format(amount / 100, "0%");
        printobj.FontBold = False
    End If
    printobj.Print " van de totale inleg (minus de dagprijzen en de rode lantaarn)"
    printobj.Print "-  ";
    printobj.FontItalic = True
    printobj.Print "Bij een gelijk aantal punten wordt de betreffende prijs verdeeld"
    printobj.FontItalic = False
    'horline 1
    'groepsstanden
    fontSizing 10
    printobj.Print
    y = printobj.CurrentY
    x = printobj.CurrentX
    fontSizing 14
    printobj.FontBold = True
    printobj.FillColor = &H808080
    printobj.FillStyle = vbFSSolid
    'printObj.BackColor = printObj.FillColor
    printobj.Line (x, y - 10)-(printobj.ScaleWidth, y + printobj.TextHeight("W") + 10), vbBlack, B
    printobj.CurrentY = y
    printobj.CurrentX = x + 50
    iBKMode = SetBkMode(printobj.hdc, TRANSPARENT)
    printobj.ForeColor = vbWhite
    printobj.Print "Groepstanden";
    fontSizing 10
    printobj.FontBold = False
    txt = " Vul in: 1 t/m 4 (" & CStr(getPoolPoints("groepstand per juist team", cn)) & " pnt per correcte invoer)"
    
    'txt = CStr(getPoolPoints("groepstand per juist team", cn))
    'printObj.CurrentX = printObj.ScaleWidth - printObj.TextWidth(txt)
    printobj.CurrentY = y + 40
    printobj.Print txt;
    printobj.CurrentY = y
    fontSizing 14
    printobj.Print
    y = printobj.CurrentY
    x = printobj.CurrentX
    fontSizing 12
    printobj.FillStyle = vbFSTransparent
    printobj.Line (x, y)-(printobj.ScaleWidth, y + printobj.TextHeight("W") * 5), vbBlack, B
    printobj.FillStyle = vbFSTransparent
    kolwidth = printobj.ScaleWidth / getTournamentInfo("tournamentGroupCount", cn)
    printobj.ForeColor = vbBlack
    For i = 1 To getTournamentInfo("tournamentGroupCount", cn)
        fontSizing 12
        x = kolwidth * (i - 1) + 50
        printobj.CurrentY = y + 10
        printobj.CurrentX = x
        printobj.FontBold = True
        printobj.Print "Groep " & Chr(i + 64)
        printobj.FontBold = False
        printgroep i
    Next
    printobj.Print
    printobj.Font = textFont
    fontSizing 2
    printobj.Print
    fontSizing 12
'    printFinals
'    printOverige
'    heading1 = "Wedstrijdvoorspellingen"
'    DoNewPage False, True
'    formulierWeds
    'InvulFormAfdrukken
End Sub

'Sub printOverige()
'invulformulier
'Dim rs As New ADODB.Recordset
'Dim topscAant As Integer
'Dim ypos As Integer
'Dim xpos As Integer
'Dim newlinepos As Integer
'Dim kolwidth As Integer
'Dim i As Integer
'
'Dim y As Integer
'Dim x As Integer
'Dim pnt As Integer
'Dim txt As String
'    newlinepos = 0
'    printObj.Print
'    kolwidth = printObj.ScaleWidth / 4
'    'eerst de eindstand
'    ypos = printObj.CurrentY
'    i = getPntToek("1e plaats(Kampioen)")
'
'    If i > 0 Then
'        'print 1e
'        txt = "(" & i & "p)"
'        printObj.Font = "Tahoma"
'        y = ypos
'        printObj.CurrentY = y
'        printObj.CurrentX = 0
'        x = printObj.CurrentX
'        fontSizing 14
'        printObj.fontBold = True
'        printObj.FillColor = &H808080
'        printObj.FillStyle = vbFSSolid
'        printObj.Line (x + 30, y - 10)-(kolwidth - 30, y + printObj.TextHeight("W")), vbBlack, B
'        printObj.CurrentY = y
'        printObj.CurrentX = x + 80
'        printObj.ForeColor = vbWhite
'        printObj.Print "Eindstand "
'        printObj.fontBold = False
'        printObj.FillStyle = vbFSTransparent
'        y = printObj.CurrentY
'        printObj.CurrentX = x + 80
'        printObj.ForeColor = vbBlack
'        fontSizing 12
'        printObj.Print "1e:";
'        fontSizing 14
'        printObj.Line (x + 30, y)-(kolwidth - 30, y + printObj.TextHeight("W")), vbBlack, B
'        printObj.CurrentY = y + 20
'        printObj.CurrentX = x + kolwidth - printObj.TextWidth(txt) + 20
'        fontSizing 10
'        printObj.Print txt;
'        printObj.CurrentY = y
'        fontSizing 14
'        printObj.Print
'        For i = 2 To 4
'            pnt = getPntToek(Format(i, "0") & "e plaats")
'            If pnt > 0 Then
'                y = printObj.CurrentY
'                txt = "(" & pnt & "p)"
'                printObj.CurrentX = x + 80
'                fontSizing 12
'                printObj.Print Format(i, "0") & "e:";
'                fontSizing 14
'                printObj.Line (x + 30, y)-(kolwidth - 30, y + printObj.TextHeight("W")), vbBlack, B
'                printObj.CurrentY = y + 20
'                printObj.CurrentX = x + kolwidth - printObj.TextWidth(txt) + 20
'                fontSizing 10
'                printObj.Print txt;
'                printObj.CurrentY = y
'                fontSizing 14
'                printObj.Print
'                If newlinepos < printObj.CurrentY Then newlinepos = printObj.CurrentY
'            End If
'        Next
'    End If
'    'topscorers
'    printObj.CurrentY = ypos
'    i = getPntToek("topscorer 1")
'    If i > 0 Then
'        'print 1e
'        txt = "(" & i & "p)"
'        printObj.Font = "Tahoma"
'        y = ypos
'        printObj.CurrentY = y
'        printObj.CurrentX = kolwidth
'        x = printObj.CurrentX
'        fontSizing 14
'        printObj.fontBold = True
'        printObj.FillColor = &H808080
'        printObj.FillStyle = vbFSSolid
'        printObj.Line (x, y - 10)-(x + kolwidth * 1.3, y + printObj.TextHeight("W")), vbBlack, B '(X + kolwidth - 30, Y + printObj.TextHeight("W")), vbBlack, B
'        printObj.CurrentY = y
'        printObj.CurrentX = x + 50
'        printObj.ForeColor = vbWhite
'        printObj.Print "Topscorer";
'        If getPntToek("topscorer 2") > 0 Then printObj.Print "s";
'
'        pnt = getPntToek("doelpunten topscorer 1")
'
'        If pnt > 0 Then
'            fontSizing 14
'            'printObj.Line (X + kolwidth - 30, Y - 10)-(X + kolwidth * 1.3, Y + printObj.TextHeight("W")), vbBlack, B
'            printObj.CurrentY = y
'            'printObj.CurrentX = X + kolwidth + 20
'            printObj.Print " & aantal goals"
'        Else
'            printObj.Print
'        End If
'        printObj.FillStyle = vbFSTransparent
'        printObj.ForeColor = vbBlack
'        printObj.fontBold = False
'        For i = 1 To 3
'            pnt = getPntToek("topscorer " & Format(i, "0"))
'            If pnt > 0 Then
'                y = printObj.CurrentY
'                txt = "(" & pnt & "p)"
'                printObj.CurrentX = x + 50
'                fontSizing 12
'                printObj.Print Format(i, "0") & ":";
'                fontSizing 14
'                printObj.Line (x, y)-(x + kolwidth - 30, y + printObj.TextHeight("W")), vbBlack, B
'                printObj.CurrentY = y + 20
'                printObj.CurrentX = x + kolwidth + 20 - printObj.TextWidth(txt)
'                fontSizing 10
'                printObj.Print txt;
'                fontSizing 14
'                pnt = getPntToek("doelpunten topscorer " & Format(i, "0"))
'                If pnt > 0 Then
'                    printObj.Line (x + kolwidth - 30, y)-(x + kolwidth * 1.3, y + printObj.TextHeight("W")), vbBlack, B
'                    printObj.CurrentY = y + 20
'                    printObj.CurrentX = x + kolwidth * 1.3 - printObj.TextWidth("(" & pnt & "p)") + 50
'                    fontSizing 10
'                    printObj.Print "("; Format(pnt, pntFormat); "p)"
'                    If newlinepos < printObj.CurrentY Then newlinepos = printObj.CurrentY
'                Else
'                    printObj.Print
'                End If
'                printObj.CurrentY = y
'                fontSizing 14
'                printObj.Print
'            End If
'        Next
'    End If
''overigen
'Dim sqlstr As String
'  sqlstr = "Select omschrijving, pnt, marge from voorspeltypes INNER JOIN pnttoek ON voorspeltypes.id = pnttoek.voorspeltype"
'  sqlstr = sqlstr & " WHERE voorspeltypes.cat = 1 and pnttoek.poolid = " & thisPool
'  sqlstr = sqlstr & " ORDER BY pnt, volgorde"
'  rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
''    rs.Open "Select * from voorspeltypes where cat =1 order by volgorde", cn, adOpenStatic, adLockReadOnly
'    y = ypos
'    printObj.CurrentY = y
'    printObj.CurrentX = x + kolwidth * 1.3 + 30
'    x = printObj.CurrentX
'    fontSizing 14
'    printObj.fontBold = True
'    printObj.FillColor = &H808080
'    printObj.FillStyle = vbFSSolid
'    printObj.Line (x, y - 10)-(printObj.ScaleWidth - 50, y + printObj.TextHeight("W")), vbBlack, B
'    printObj.CurrentY = y
'    printObj.CurrentX = x + 50
'    printObj.ForeColor = vbWhite
'    printObj.Print "Overigen "
'    printObj.FillStyle = vbFSTransparent
'    printObj.ForeColor = vbBlack
'    printObj.fontBold = False
'    Do While Not rs.EOF
'        pnt = rs!pnt
'        y = printObj.CurrentY
'        txt = "(" & pnt & "p)"
'        If nz(rs!marge, 0) > 0 Then
'          txt = "(±" & rs!marge & ", " & pnt & "p)"
'        End If
'        printObj.CurrentX = x + 50
'        fontSizing 10
'        printObj.Print rs!omschrijving; " "; txt; ":";
'        fontSizing 14
'        printObj.Line (x, y)-(printObj.ScaleWidth - 50, y + printObj.TextHeight("W")), vbBlack, B
'        rs.MoveNext
'        If newlinepos < printObj.CurrentY Then newlinepos = printObj.CurrentY
'    Loop
'    rs.Close
'    Set rs = Nothing
'    printObj.Line (printObj.ScaleWidth - 30 - printObj.TextWidth("1234"), ypos + 360)-(printObj.ScaleWidth - 30 - printObj.TextWidth("1234"), printObj.CurrentY)
'    printObj.Line (0, ypos - 50)-(printObj.ScaleWidth - 10, newlinepos + 30), , B
'
'End Sub
'Sub printFinals()
''onderdeel van formulieren
'Dim rs As New ADODB.Recordset
'Dim sqlstr As String
'Dim xpos As Integer
'Dim ypos As Integer
'Dim i As Integer
'Dim x As Integer
'Dim y As Integer
'Dim txt As String
'Dim intP As Integer
'Dim intQ As Integer
'Dim hvFinYpos As Integer
'Dim HeeftKlFin As Boolean
'    i = getPntToek("achtste finaleplaats") + getPntToek("achtste finalepositie")
'    If i > 0 Then
'        'print achtste finales
'        txt = "("
'        intP = getPntToek("achtste finaleplaats")
'        intQ = getPntToek("achtste finalepositie")
'        If intP > 0 Then txt = txt & intP & " pnt voor elk genoemd team"
'
'        If intQ > 0 Then
'            If txt > "(" Then txt = txt & " of "
'            txt = txt & intQ & " pnt als het ook nog op de juiste plaats staat"
'        Else
'            txt = txt & ", juiste plaats niet van belang"
'        End If
'        txt = txt & ")"
'        printObj.Font = "Tahoma"
'        y = printObj.CurrentY
'        x = printObj.CurrentX
'        fontSizing 14
'        printObj.fontBold = True
'        printObj.FillColor = &H808080
'        printObj.FillStyle = vbFSSolid
'        printObj.Line (x, y - 10)-(printObj.ScaleWidth, y + printObj.TextHeight("W")), vbBlack, B
'        'printObj.BackColor = printObj.FillColor
'        iBKMode = SetBkMode(printObj.hdc, TRANSPARENT)
'        printObj.ForeColor = vbWhite
'        printObj.CurrentY = y
'        printObj.CurrentX = x + 50
'        printObj.Print "Achtstefinales ";
'        printObj.FillStyle = vbFSTransparent
'        fontSizing 10
'        printObj.fontBold = False
''        printObj.CurrentX = printObj.ScaleWidth - printObj.TextWidth(txt)
'        printObj.CurrentY = y + 40
'        printObj.Print txt;
'        printObj.ForeColor = vbBlack
'        printObj.CurrentY = y
'        fontSizing 14
'        printObj.Print
'        y = printObj.CurrentY
'        x = printObj.CurrentX
'        fontSizing 12
'        printObj.Line (x, y)-(printObj.ScaleWidth, y + printObj.TextHeight("W") * 4.7), vbBlack, B
'        y = y + 50
'        printObj.CurrentY = y
'        printObj.FillStyle = vbFSTransparent
'        kolwidth = printObj.ScaleWidth / 4
'        sqlstr = "Select * from qryWeds where  ksid = " & kampID
'        sqlstr = sqlstr & " and wedtype = 5"
'        sqlstr = sqlstr & " ORDER BY wednum"
'        rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'        xpos = 0
'        With rs
'            If .RecordCount > 0 Then
'                i = 0
'                Do While Not .EOF
'                    ypos = y
'                    fontSizing 8
'                    printObj.CurrentX = xpos + 50
'                    printObj.CurrentY = ypos + printObj.TextHeight("99") * 0.5
'                    printObj.Print Format(!wedNum, "0"); ":";
'                    fontSizing 12
'                    printObj.CurrentX = xpos + printObj.TextWidth("00:") + 30
'                    printObj.CurrentY = ypos
'                    fontSizing 10
'                    printObj.Print !code1; ":";
'                    fontSizing 12
'                    printObj.DrawWidth = 1
'                    printObj.Line (xpos + printObj.TextWidth("00:"), ypos)-(xpos + kolwidth - 50, ypos + printObj.TextHeight("W")), vbBlack, B
'                    ypos = printObj.CurrentY
'                    printObj.CurrentX = xpos + printObj.TextWidth("00:") + 30
'                    fontSizing 10
'                    printObj.Print !code2; ":";
'                    fontSizing 12
'                    printObj.Line (xpos + printObj.TextWidth("00:"), ypos)-(xpos + kolwidth - 50, ypos + printObj.TextHeight("W")), vbBlack, B
'                    'wedstrijd nr
'                    printObj.CurrentY = ypos
'                    .MoveNext
'                    i = i + 1
'                    xpos = kolwidth * i
'                    If xpos > printObj.ScaleWidth - kolwidth + 100 Then
'                        fontSizing 8
'                        printObj.Print
'                        printObj.Print
'                        fontSizing 10
'                        y = printObj.CurrentY
'                        i = 0
'                        xpos = 0
'                    End If
'                Loop
'            End If
'        End With
'    End If
'    fontSizing 2
'    printObj.Print
'    fontSizing 12
'    i = getPntToek("kwart finaleplaats") + getPntToek("kwart finalepositie")
'    If i > 0 Then
'        'print kwart finales
'        txt = "("
'        intP = getPntToek("kwart finaleplaats")
'        intQ = getPntToek("kwart finalepositie")
'        If intP > 0 Then txt = txt & intP & " pnt voor elk genoemd team"
'        If intQ > 0 Then
'            If txt > "(" Then txt = txt & " of "
'            txt = txt & intQ & " pnt als het ook nog op de juiste plaats staat"
'        Else
'            txt = txt & ", juiste plaats hoeft niet"
'        End If
'        txt = txt & ")"
'        printObj.Font = "Tahoma"
'        y = printObj.CurrentY
'        x = printObj.CurrentX
'        fontSizing 14
'        printObj.fontBold = True
'        printObj.FillColor = &H808080
'        printObj.FillStyle = vbFSSolid
'        printObj.Line (x, y - 10)-(printObj.ScaleWidth, y + printObj.TextHeight("W")), vbBlack, B
'        printObj.CurrentY = y
'        printObj.CurrentX = x + 50
'        printObj.ForeColor = vbWhite
'        printObj.Print "Kwartfinales ";
'        fontSizing 10
'        printObj.fontBold = False
''        printObj.CurrentX = printObj.ScaleWidth - printObj.TextWidth(txt)
'        printObj.CurrentY = y + 40
'        printObj.Print txt;
'        printObj.ForeColor = vbBlack
'        printObj.FillStyle = vbFSTransparent
'        printObj.CurrentY = y
'        fontSizing 14
'        printObj.Print
'        y = printObj.CurrentY
'        x = printObj.CurrentX
'        fontSizing 12
'        printObj.Line (x, y)-(printObj.ScaleWidth, y + printObj.TextHeight("W") * 2.5), vbBlack, B
'        y = y + 50
'        printObj.CurrentY = y
'        printObj.FillStyle = vbFSTransparent
'        kolwidth = (printObj.ScaleWidth / 8) * 2
'        sqlstr = "Select * from qryWeds where  ksid = " & kampID
'        sqlstr = sqlstr & " and wedtype = 2"
'        sqlstr = sqlstr & " ORDER BY wednum"
'        rs.Close
'        rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'        xpos = 0
'        With rs
'            If .RecordCount > 0 Then
'                i = 0
'                Do While Not .EOF
'                    ypos = y
'                    fontSizing 8
'                    printObj.CurrentX = xpos + 50
'                    printObj.CurrentY = ypos + printObj.TextHeight("99") * 0.5
'                    printObj.Print Format(!wedNum, "0"); ":";
'                    fontSizing 12
'                    printObj.CurrentX = xpos + printObj.TextWidth("00:") + 30
'                    printObj.CurrentY = ypos
'                    fontSizing 10
'                    printObj.Print !code1; ":";
'                    fontSizing 12
'                    printObj.DrawWidth = 1
'                    printObj.Line (xpos + printObj.TextWidth("00:"), ypos)-(xpos + kolwidth - 50, ypos + printObj.TextHeight("W")), vbBlack, B
'                    ypos = printObj.CurrentY
'                    printObj.CurrentX = xpos + printObj.TextWidth("00:") + 30
'                    fontSizing 10
'                    printObj.Print !code2; ":";
'                    fontSizing 12
'                    printObj.Line (xpos + printObj.TextWidth("00:"), ypos)-(xpos + kolwidth - 50, ypos + printObj.TextHeight("W")), vbBlack, B
'                    'wedstrijd nr
'                    printObj.CurrentY = ypos
'                    .MoveNext
'                    i = i + 1
'                    xpos = kolwidth * i
'                    If xpos > printObj.ScaleWidth - kolwidth + 100 Then
'                        fontSizing 8
'                        printObj.Print
'                        printObj.Print
'                        fontSizing 12
'                        y = printObj.CurrentY
'                        i = 0
'                        xpos = 0
'                    End If
'                Loop
'            End If
'        End With
'    End If
'    fontSizing 2
'    printObj.Print
'    fontSizing 12
'    hvFinYpos = printObj.CurrentY
'    i = getPntToek("halve finaleplaats") + getPntToek("halve finalepositie")
'    If i > 0 Then
'        'print halve finales
'        txt = "("
'        intP = getPntToek("halve finaleplaats")
'        If intP > 0 Then txt = txt & intP & ""
'        intQ = getPntToek("halve finalepositie")
'        If intQ > 0 Then
'            If txt > "(" Then txt = txt & "/"
'            txt = txt & intQ & " pnt"
'        Else
'            txt = txt & " pnt"
'        End If
'        txt = txt & ")"
'        printObj.Font = "Tahoma"
'        y = printObj.CurrentY
'        x = printObj.CurrentX
'        fontSizing 14
'        printObj.fontBold = True
'        printObj.FillColor = &H808080
'        printObj.FillStyle = vbFSSolid
'        printObj.Line (x, y - 10)-(printObj.ScaleWidth / 2 - 30, y + printObj.TextHeight("W")), vbBlack, B
'        printObj.CurrentY = y
'        printObj.CurrentX = x + 50
'        printObj.ForeColor = vbWhite
'        printObj.Print "Halve finales ";
'        fontSizing 10
'        printObj.fontBold = False
'        'printObj.CurrentX = printObj.ScaleWidth / 2 - 30 - printObj.TextWidth(txt)
'        printObj.CurrentY = y + 40
'        printObj.Print txt;
'        printObj.ForeColor = vbBlack
'        printObj.CurrentY = y
'        fontSizing 14
'        printObj.Print
'        y = printObj.CurrentY
'        x = printObj.CurrentX
'        printObj.FillStyle = vbFSTransparent
'        fontSizing 12
'        printObj.Line (x, y)-(printObj.ScaleWidth / 2 - 30, y + printObj.TextHeight("W") * 2.5), vbBlack, B
'        y = y + 50
'        printObj.CurrentY = y
'        printObj.FillStyle = vbFSTransparent
'        kolwidth = (printObj.ScaleWidth / 8) * 2
'        sqlstr = "Select * from qryWeds where  ksid = " & kampID
'        sqlstr = sqlstr & " and wedtype = 3"
'        sqlstr = sqlstr & " ORDER BY wednum"
'        rs.Close
'        rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'        xpos = 0
'        With rs
'            If .RecordCount > 0 Then
'                i = 0
'                Do While Not .EOF
'                    ypos = y
'                    fontSizing 8
'                    printObj.CurrentX = xpos + 50
'                    printObj.CurrentY = ypos + printObj.TextHeight("99") * 0.5
'                    printObj.Print Format(!wedNum, "0"); ":";
'                    fontSizing 12
'                    printObj.CurrentX = xpos + printObj.TextWidth("00:") + 30
'                    printObj.CurrentY = ypos
'                    fontSizing 10
'                    printObj.Print !code1; ":";
'                    fontSizing 12
'                    printObj.DrawWidth = 1
'                    printObj.Line (xpos + printObj.TextWidth("00:"), ypos)-(xpos + kolwidth - 50, ypos + printObj.TextHeight("W")), vbBlack, B
'                    ypos = printObj.CurrentY
'                    printObj.CurrentX = xpos + printObj.TextWidth("00:") + 30
'                    fontSizing 10
'                    printObj.Print !code2; ":";
'                    fontSizing 12
'                    printObj.Line (xpos + printObj.TextWidth("00:"), ypos)-(xpos + kolwidth - 50, ypos + printObj.TextHeight("W")), vbBlack, B
'                    'wedstrijd nr
'                    printObj.CurrentY = ypos
'                    .MoveNext
'                    i = i + 1
'                    xpos = kolwidth * i
'                    If xpos > printObj.ScaleWidth - kolwidth + 100 Then
'                        fontSizing 8
'                        printObj.Print
'                        printObj.Print
'                        fontSizing 12
'                        y = printObj.CurrentY
'                        i = 0
'                        xpos = 0
'                    End If
'                Loop
'            End If
'        End With
'    End If
'    printObj.CurrentY = hvFinYpos
'    i = getPntToek("kleine finaleplaats") + getPntToek("kleine finalepositie")
'    If i > 0 Then
'        HeeftKlFin = True
'        'print kleine finale
'        txt = "("
'        intP = getPntToek("kleine finaleplaats")
'        If intP > 0 Then txt = txt & intP & ""
'        intP = getPntToek("kleine finalepositie")
'        If intP > 0 Then
'            If txt > "(" Then txt = txt & "/"
'            txt = txt & intP & " pnt"
'        Else
'            txt = txt & " pnt"
'        End If
'        txt = txt & ")"
'        printObj.Font = "Tahoma"
'        y = hvFinYpos
'        printObj.CurrentY = y
'        printObj.CurrentX = printObj.ScaleWidth / 2 + 30
'        x = printObj.CurrentX
'        fontSizing 14
'        printObj.fontBold = True
'        printObj.FillColor = &H808080
'        printObj.FillStyle = vbFSSolid
'        printObj.Line (x, y - 10)-(printObj.ScaleWidth * 0.75, y + printObj.TextHeight("W")), vbBlack, B
'        printObj.CurrentY = y
'        printObj.CurrentX = x + 50
'        printObj.ForeColor = vbWhite
'        printObj.Print "3e plaats ";
'        fontSizing 10
'        printObj.fontBold = False
'        printObj.CurrentY = y + 40
'        printObj.Print txt;
'        printObj.ForeColor = vbBlack
'        printObj.CurrentY = y
'        fontSizing 14
'        printObj.Print
'        printObj.FillStyle = vbFSTransparent
'        y = printObj.CurrentY
'        x = printObj.ScaleWidth / 2 + 30
'        fontSizing 12
'        printObj.Line (x, y)-(printObj.ScaleWidth * 0.75, y + printObj.TextHeight("W") * 2.5), vbBlack, B
'        y = y + 50
'        printObj.CurrentY = y
'        printObj.FillStyle = vbFSTransparent
'        kolwidth = (printObj.ScaleWidth / 8) * 2
'        sqlstr = "Select * from qryWeds where  ksid = " & kampID
'        sqlstr = sqlstr & " and wedtype = 7"
'        sqlstr = sqlstr & " ORDER BY wednum"
'        rs.Close
'        rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'        xpos = printObj.ScaleWidth / 2 + 30
'        With rs
'            If .RecordCount > 0 Then
'                i = 0
'                Do While Not .EOF
'                    ypos = y
'                    fontSizing 8
'                    printObj.CurrentX = xpos + 50
'                    printObj.CurrentY = ypos + printObj.TextHeight("99") * 0.5
'                    printObj.Print Format(!wedNum, "0"); ":";
'                    fontSizing 12
'                    printObj.CurrentX = xpos + printObj.TextWidth("00:") + 30
'                    printObj.CurrentY = ypos
'                    printObj.Print !code1; ":";
'                    printObj.DrawWidth = 1
'                    printObj.Line (xpos + printObj.TextWidth("00:"), ypos)-(xpos + kolwidth - 50, ypos + printObj.TextHeight("W")), vbBlack, B
'                    ypos = printObj.CurrentY
'                    printObj.CurrentX = xpos + printObj.TextWidth("00:") + 30
'                    printObj.Print !code2; ":";
'                    printObj.Line (xpos + printObj.TextWidth("00:"), ypos)-(xpos + kolwidth - 50, ypos + printObj.TextHeight("W")), vbBlack, B
'                    'wedstrijd nr
'                    printObj.CurrentY = ypos
'                    .MoveNext
'                    i = i + 1
'                    xpos = kolwidth * i
'                    If xpos > printObj.ScaleWidth - kolwidth + 100 Then
'                        fontSizing 8
'                        printObj.Print
'                        printObj.Print
'                        fontSizing 12
'                        y = printObj.CurrentY
'                        i = 0
'                        xpos = 0
'                    End If
'                Loop
'            End If
'        End With
'    Else
'        HeeftKlFin = False
'    End If
'    printObj.CurrentY = hvFinYpos
'    i = getPntToek("finaleplaats") + getPntToek("finalepositie")
'    If i > 0 Then
'        'print finale
'        txt = "("
'        intP = getPntToek("finaleplaats")
'        If intP > 0 Then txt = txt & intP
'        intP = getPntToek("finalepositie")
'        If intP > 0 Then
'            If txt > "(" Then txt = txt & "/"
'            txt = txt & intP & " pnt"
'        Else
'            txt = txt & " pnt"
'        End If
'        txt = txt & ")"
'        printObj.Font = "Tahoma"
'        y = hvFinYpos
'        printObj.CurrentY = y
'        If HeeftKlFin Then
'            printObj.CurrentX = printObj.ScaleWidth * 0.75 + 30
'        Else
'            printObj.CurrentX = printObj.ScaleWidth * 0.5 + 30
'        End If
'        x = printObj.CurrentX
'        fontSizing 14
'        printObj.fontBold = True
'        printObj.FillColor = &H808080
'        printObj.FillStyle = vbFSSolid
'        printObj.Line (x, y - 10)-(printObj.ScaleWidth, y + printObj.TextHeight("W")), vbBlack, B
'        printObj.ForeColor = vbWhite
'        printObj.CurrentY = y
'        printObj.CurrentX = x + 50
'        printObj.Print "Finale ";
'        fontSizing 10
'        printObj.fontBold = False
'        printObj.CurrentY = y + 40
'        printObj.Print txt;
'        printObj.ForeColor = vbBlack
'        printObj.CurrentY = y
'        fontSizing 14
'        printObj.Print
'        printObj.FillStyle = vbFSTransparent
'        y = printObj.CurrentY
'        If HeeftKlFin Then
'            x = printObj.ScaleWidth * 0.75 + 30
'        Else
'            x = printObj.ScaleWidth * 0.5 + 30
'        End If
'        fontSizing 12
'        printObj.Line (x, y)-(printObj.ScaleWidth, y + printObj.TextHeight("W") * 2.5), vbBlack, B
'        y = y + 50
'        printObj.CurrentY = y
'        printObj.FillStyle = vbFSTransparent
'        kolwidth = (printObj.ScaleWidth / 8) * 2
'        sqlstr = "Select * from qryWeds where  ksid = " & kampID
'        sqlstr = sqlstr & " and wedtype = 4"
'        sqlstr = sqlstr & " ORDER BY wednum"
'        rs.Close
'        rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'        If HeeftKlFin Then
'            xpos = printObj.ScaleWidth * 0.75 + 30
'        Else
'            xpos = printObj.ScaleWidth * 0.5 + 30
'            kolwidth = kolwidth * 2
'        End If
'        With rs
'            If .RecordCount > 0 Then
'                i = 0
'                Do While Not .EOF
'                    ypos = y
'                    fontSizing 8
'                    printObj.CurrentX = xpos + 50
'                    printObj.CurrentY = ypos + printObj.TextHeight("99") * 0.5
'                    'wedstrijd nr
'                    printObj.Print Format(!wedNum, "0"); ":";
'                    fontSizing 12
'                    printObj.CurrentX = xpos + printObj.TextWidth("00:") + 30
'                    printObj.CurrentY = ypos
'                    fontSizing 10
'                    printObj.Print !code1; ":";
'                    fontSizing 12
'                    printObj.DrawWidth = 1
'                    printObj.Line (xpos + printObj.TextWidth("00:"), ypos)-(xpos + kolwidth - 50, ypos + printObj.TextHeight("W")), vbBlack, B
'                    ypos = printObj.CurrentY
'                    printObj.CurrentX = xpos + printObj.TextWidth("00:") + 30
'                    fontSizing 10
'                    printObj.Print !code2; ":";
'                    fontSizing 12
'                    printObj.Line (xpos + printObj.TextWidth("00:"), ypos)-(xpos + kolwidth - 50, ypos + printObj.TextHeight("W")), vbBlack, B
'                    printObj.CurrentY = ypos
'                    .MoveNext
'                    i = i + 1
'                    xpos = kolwidth * i
'                    If xpos > printObj.ScaleWidth - kolwidth + 100 Then
'                        fontSizing 8
'                        printObj.Print
'                        printObj.Print
'                        fontSizing 12
'                        y = printObj.CurrentY
'                        i = 0
'                        xpos = 0
'                    End If
'                Loop
'            End If
'            .Close
'        End With
'        Set rs = Nothing
'    End If
'    fontSizing 8
'    printObj.Print
'    fontSizing 12
'End Sub
'
Sub printgroep(nr As Integer)
Dim sqlstr As String
Dim xLinePos As Integer
Dim yLinePos As Integer
Dim xpos As Integer
Dim txt As String
Dim vakPos(1, 1)
Dim grp As String * 1
Dim iGrp As Integer
fontSizing 10

Set rs = New ADODB.Recordset
sqlstr = "Select groupLetter, groupPlace, teamName from (tblGroupLayout l"
sqlstr = sqlstr & " INNER JOIN tblTournamentTeamCodes c ON (l.teamId = c.teamId) "
sqlstr = sqlstr & " AND (l.tournamentID = c.tournamentId))"
sqlstr = sqlstr & " INNER JOIN tblTeamNames n on n.teamNameId = c.teamid "
sqlstr = sqlstr & " WHERE l.groupLetter = '" & Chr(64 + nr) & "'"
sqlstr = sqlstr & " ORDER BY groupletter, groupPlace"
rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly

yLinePos = printobj.CurrentY
iGrp = getTournamentInfo("tournamentGroupCount", cn)

xLinePos = (printobj.ScaleWidth / iGrp) * (nr - 1)
xpos = xLinePos + 50
Do While Not rs.EOF
    vakPos(0, 0) = xpos + printobj.ScaleWidth / iGrp - printobj.TextHeight("W") - printobj.TextWidth("W")
    vakPos(0, 1) = printobj.CurrentY
    vakPos(1, 0) = vakPos(0, 0) + printobj.TextHeight("W")
    vakPos(1, 1) = vakPos(0, 1) + printobj.TextHeight("W")

    txt = rs!teamName

    Do While xpos + printobj.TextWidth(txt) > vakPos(0, 0)
        txt = Left(txt, Len(txt) - 1)
    Loop
    printobj.CurrentX = xpos
    printobj.Print txt;
    printobj.FillStyle = vbFSTransparent
    printobj.FillColor = vbWhite
    printobj.DrawWidth = 1

    printobj.Line (vakPos(0, 0), vakPos(0, 1))-(vakPos(1, 0), vakPos(1, 1)), vbBlack, B
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
'printObj.CurrentY = yLinePos
End Sub
'
'Sub formulierWeds()
''wedstrijden op het poolformulier
'Dim fontBas As Integer
'Dim rs As New ADODB.Recordset
'Dim sqlstr As String
'Dim posWednr As Integer
'Dim posDatum As Integer
'Dim posTijd As Integer
'Dim posWedOms As Integer
'Dim posRust As Integer
'Dim PosEind As Integer
'Dim posToto As Integer
'Dim wedOms As String
'Dim kolwidth As Integer
'Dim kolom As Integer
'Dim ypos As Integer
'Dim curYpos As Integer
'Dim x As Integer
'Dim y As Integer
'Dim i As Integer
'Dim vertLineYPos As Integer
'Dim vertLineYPos2 As Integer
'Dim topY As String
'Dim savdat As Date
'Dim vertLineEndPos As Integer
'    sqlstr = "Select * from qryweds where ksid = " & kampID
'    sqlstr = sqlstr & " ORDER BY datum,tijd,wednum"
'    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    If rs.RecordCount = 0 Then
'        rs.Close
'        Exit Sub
'    End If
'    fontBas = 10
'    fontSizing fontBas + 2
'    topY = printObj.CurrentY
'    printObj.CurrentY = voethoog - GrootHoog
'    ypos = printObj.CurrentY
'    printObj.FillColor = &H808080
'    printObj.FillStyle = vbFSSolid
'    printObj.Line (0, ypos)-(printObj.ScaleWidth + 2 * printObj.ScaleLeft, voethoog), vbBlack, B
'    printObj.CurrentY = ypos + 30
'    fontSizing 16
'    printObj.fontBold = True
'    printObj.ForeColor = vbWhite
'    iBKMode = SetBkMode(printObj.hdc, TRANSPARENT)
'    Centreer "UITERLIJK INLEVEREN OP " & UCase(Format(getPoolInfo("eindinschr"), "dddd d mmmm yyyy"))
'    printObj.ForeColor = vbBlack
'    printObj.FillStyle = vbFSTransparent
'    printObj.fontBold = False
'    fontSizing fontBas + 2
'    printObj.CurrentY = topY
'    kolom = 0
'    kolwidth = printObj.ScaleWidth / 2 - printObj.TextWidth("w")
'    printObj.FontName = "Times New Roman"
'    fontSizing 2
'    printObj.Print
'    fontSizing fontBas + 2
'    printObj.CurrentY = printObj.CurrentY + 20
'    fontSizing fontBas + 4
'    printObj.fontBold = True
'    printObj.Print "Uitleg"
'    fontSizing fontBas + 2
'    printObj.fontBold = False
'    printObj.Print "Vul hieronder voor alle wedstrijden jouw uitslagen in. ";
'    printObj.fontBold = True
'    printObj.Print "Ook daar waar de teams nog niet bekend zijn."
'    printObj.fontBold = False
'    printObj.Print "(Ook al heb je een ander team op die plaats dan kan je uitslag nog steeds goed zijn)"
'    printObj.Print "De uitslag hoeft onderling niet te kloppen. ";
'    printObj.Print "Je krijgt punten voor elk vak dat achteraf juist blijkt te zijn ingevuld."
'    printObj.Print "Bij 'toto' vul je een 1 in voor winst linker team, een 2 voor winst rechter team en een 3 voor een gelijkspel"
'    printObj.fontBold = True
'    Centreer "Alle uitslagen, ook de toto, gelden na 90 minuten voetbal!"
'    printObj.fontBold = False
'    fontSizing fontBas
'    printObj.Print
'    Centreer "(plus de eventuele blessuretijd)"
'    printObj.Print
'    fontSizing fontBas + 4
'    printObj.fontBold = True
'    printObj.Print "Punten"
'    printObj.fontBold = False
'    fontSizing fontBas + 2
'    printObj.Print "Ruststand goed: ";
'    printObj.fontBold = True
'    printObj.Print getPntToek("ruststand goed"); "pnt, ";
'    printObj.fontBold = False
'    printObj.Print "Eindstand goed: ";
'    printObj.fontBold = True
'    printObj.Print getPntToek("eindstand goed"); "pnt, ";
'    printObj.fontBold = False
'    printObj.Print "Toto goed: ";
'    printObj.fontBold = True
'    printObj.Print getPntToek("toto goed"); "pnt.";
'    printObj.fontBold = False
'    If getPntToek("doelpunten op een dag") > 0 Then
'        printObj.Print "Totaal aantal doelpunten op één dag goed: ";
'        printObj.fontBold = True
'        printObj.Print getPntToek("doelpunten op een dag"); " pnt"
'        printObj.fontBold = False
'    End If
'    printObj.Print
'    fontSizing fontBas
'    posDatum = 50
'    posTijd = posDatum + printObj.TextWidth("MA 26-6") + 10
'    posWednr = posTijd + printObj.TextWidth("00:000") + 10
'    posWedOms = posWednr + printObj.TextWidth("199:")
'    posRust = posWedOms + printObj.TextWidth("Nederland - Zwitserland")
'    PosEind = posRust + printObj.TextWidth("123456")
'    posToto = PosEind + printObj.TextWidth("123456")
'
'    vertLineYPos = printObj.CurrentY
'    fontSizing fontBas
'    printObj.Line (0, vertLineYPos - 20)-(kolwidth * 2, vertLineYPos - 20)
'    printObj.CurrentY = vertLineYPos
'    For i = 0 To 1
'        printObj.CurrentX = posDatum + i * kolwidth
'        printObj.Print " Datum";
'        printObj.CurrentX = posTijd + i * kolwidth
'        printObj.Print " tijd";
'        printObj.CurrentX = posWednr + i * kolwidth
'        printObj.Print " nr";
'        printObj.CurrentX = posWedOms + i * kolwidth
'        printObj.Print " Wedstrijd";
'        printObj.CurrentX = posRust + i * kolwidth
'        printObj.Print " rust";
'        printObj.CurrentX = PosEind + i * kolwidth
'        printObj.Print " eind";
'        printObj.CurrentX = posToto + i * kolwidth
'        printObj.Print " toto";
'    Next
'    printObj.Print
'    printObj.Line (0, printObj.CurrentY)-(kolwidth * 2, printObj.CurrentY), 1
'    vertLineYPos2 = printObj.CurrentY
'
'    ypos = printObj.CurrentY
'
'    With rs
'        .MoveLast
'        .MoveFirst
'
'        Do While Not .EOF
'            If (nz(!naam1, "")) > "" Then
'                wedOms = !code1 & ":" & !naam1 & " - " & !code2 & ":" & !naam2
'            Else
'                wedOms = !code1 & " - " & !code2
'            End If
'
'            printObj.CurrentY = printObj.CurrentY + 40
'            printObj.CurrentX = posWednr + kolom * kolwidth + (posWedOms - posWednr - printObj.TextWidth(Format(!wedNum, "0"))) / 2
'            printObj.Print Format(!wedNum, "0");
'            printObj.CurrentX = posDatum + kolom * kolwidth
'            If savdat <> !datum Then
'                printObj.Print Format(!datum, "ddd d-M"); " ";
'                savdat = !datum
'            End If
'            printObj.CurrentX = posTijd + kolom * kolwidth + (posWednr - posTijd - printObj.TextWidth(Format(!tijd, "HH:NN"))) / 2
'            printObj.Print tijdFormat(!tijd); '  , "HH:NN");
'            printObj.CurrentX = posWedOms + kolom * kolwidth + 30
'            curYpos = printObj.CurrentY
'            If (nz(!naam1, "")) > "" Then
'                fontSizing fontBas - 3
'                printObj.CurrentY = curYpos + 20
'                Do While printObj.TextWidth(wedOms) > posRust - posWedOms
'                    wedOms = Left(wedOms, Len(wedOms) - 1)
'                Loop
'            Else
'                fontSizing fontBas
'                printObj.CurrentY = curYpos
'            End If
'            printObj.Print wedOms;
'            printObj.CurrentY = curYpos
'            fontSizing fontBas
'            x = posRust + kolom * kolwidth
'            y = printObj.CurrentY - 20
'            printObj.Line (x, y)-(PosEind + kolom * kolwidth - 10, y + printObj.TextHeight("W") + 50), , B
'            printObj.CurrentX = posRust + (PosEind - posRust - printObj.TextWidth("-")) / 2 + kolom * kolwidth
'            printObj.CurrentY = y + 30
'            printObj.Print "-";
'            x = PosEind + kolom * kolwidth + 10
'            printObj.Line (x, y)-(posToto + kolom * kolwidth - 10, y + printObj.TextHeight("W") + 50), , B
'            printObj.CurrentX = PosEind + (posToto - PosEind - printObj.TextWidth("-")) / 2 + kolom * kolwidth
'            printObj.CurrentY = y + 30
'            printObj.Print "-";
'            x = posToto + kolom * kolwidth + 10
'            printObj.Line (x, y)-(kolwidth * (kolom + 1) - printObj.TextWidth("0"), y + printObj.TextHeight("W") + 50), , B
'            printObj.CurrentX = PosEind + (posToto - PosEind - printObj.TextWidth("-")) / 2
'            printObj.CurrentY = y
'
'            fontSizing 14
'            printObj.Print
'            fontSizing fontBas
'            printObj.Line (0, printObj.CurrentY)-(kolwidth * 2, printObj.CurrentY), 1
'
'            .MoveNext
'            If (.AbsolutePosition - 1) = Int(rs.RecordCount / 2 + 0.5) Then
'                kolom = 1
'                vertLineEndPos = printObj.CurrentY
'                printObj.CurrentY = ypos
'            End If
'        Loop
'        .Close
'    End With
'    Set rs = Nothing
'    For i = 0 To 1
'        printObj.Line (0 + kolwidth * i, vertLineYPos - 10)-(0 + kolwidth * i, vertLineEndPos)
'        printObj.Line (posWednr + kolwidth * i - 10, vertLineYPos2)-(posWednr + kolwidth * i - 10, vertLineEndPos)
'        printObj.Line (posTijd + kolwidth * i, vertLineYPos2)-(posTijd + kolwidth * i, vertLineEndPos)
'        printObj.Line (posWedOms + kolwidth * i - 10, vertLineYPos2)-(posWedOms + kolwidth * i - 10, vertLineEndPos)
'    Next
'    printObj.Line (kolwidth - 50, vertLineYPos - 10)-(kolwidth - 50, vertLineEndPos)
'    printObj.Line (kolwidth * 2, vertLineYPos - 10)-(kolwidth * 2, vertLineEndPos)
'End Sub
'
'
'Private Sub PrijsAfdr(wat As String, eind As Boolean)
'Dim aant As Integer
'Dim i As Integer
'End Sub
'
Private Sub Centreer(Tekst$)
    printobj.CurrentX = (printobj.ScaleWidth - printobj.TextWidth(Trim$(Tekst$))) \ 2
    printobj.Print Tekst$;
End Sub
'
'Function sqlDeelnems(poule As Long) As String
'Dim sqlstr As String
'    sqlstr = "Select * from pooldeelnems"
'    sqlstr = sqlstr & " WHERE PoolID = " & poule
'    sqlstr = sqlstr & " ORDER BY bijnaam "
'    sqlDeelnems = sqlstr
'End Function
'
'Private Sub printFavourites()
'Dim rs As New ADODB.Recordset
'Dim rs2 As New ADODB.Recordset
'Dim aantgroep As Integer
'Dim i As Integer
'Dim J As Integer
'Dim aant As Integer
'Dim savX As Integer
'Dim savy As Integer
'Dim xpos As Integer
'Dim col(4) As Integer
'Dim yStart As Integer
'Dim maxrows As Integer
'Dim bewYPos As Integer
'Dim deelnAant As Integer
'Dim fntGr As Double
'Dim sqlstr As String
'
'deelnAant = GetDeelnemAant(thisPool)
'headerText = GetOrgNaam(thisPool) & " " & getTournamentInfo("toernooi") & " voetbalpool" & " - Favorieten" & " (" & GetDeelnemAant(thisPool) & " deelnemers)"
''printObj.Line (0, printObj.CurrentY)-(printObj.ScaleWidth, printObj.CurrentY)
'heading1 = "Groepstanden"
'InitPage False, False
''intro
'yStart = printObj.CurrentY
'
''groepen
'fntGr = printObj.Font.Size
'sqlstr = "Select groepen from ks WHERE id = " & kampID
'rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'aantgroep = rs!groepen
'rs.Close
'printObj.CurrentX = printObj.TextWidth("12345678901234567890123456")
'For i = 1 To 4
'    printObj.CurrentX = printObj.CurrentX - printObj.TextWidth(Format(i, "0") & "e pl")
'    printObj.Print Format(i, "0"); "e pl";
'    col(i) = printObj.CurrentX - 50
'    printObj.CurrentX = printObj.CurrentX + printObj.TextWidth("123456")
'Next
'printObj.CurrentX = printObj.ScaleWidth / 2 + printObj.TextWidth("12345678901234567890123456")
'For i = 1 To 4
'    printObj.CurrentX = printObj.CurrentX - printObj.TextWidth(Format(i, "0") & "e pl")
'    printObj.Print Format(i, "0"); "e pl";
'    printObj.CurrentX = printObj.CurrentX + printObj.TextWidth("123456")
'Next
'printObj.CurrentX = 0
'printObj.Print
'xpos = 0
'savy = printObj.CurrentY
'For i = 1 To aantgroep
'    If i = aantgroep / 2 + 1 Then
'        xpos = printObj.ScaleWidth / 2
'        printObj.CurrentY = savy
'    End If
'    sqlstr = "Select * from groepsindeling where ksid = " & kampID
'    sqlstr = sqlstr & " AND groep = '" & Chr(i + 64) & "'"
'    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    rs.MoveFirst
'    printObj.CurrentX = xpos
'    printObj.Print "Groep " & rs!groep; ": ";
'    savX = printObj.CurrentX
'    Do While Not rs.EOF
'        printObj.CurrentX = savX
'        printObj.Print GetTeam(rs!team); " ";
'        printObj.CurrentX = printObj.TextWidth("12345678901234567890")
'        For J = 1 To 4
'            aant = getAantalGrpVoorsp(J, rs!team)
'            fontSizing 9
'            printObj.CurrentY = printObj.CurrentY + 30
'            printObj.CurrentX = xpos + col(J) - printObj.TextWidth(Format(aant / deelnAant, "0.0%"))
''            printObj.Print aant;
''            fontSizing 8
'            printObj.Print Format(aant / deelnAant, "0.0%");
'            printObj.CurrentY = printObj.CurrentY - 30
'            fontSizing CInt(fntGr)
'            'If j < 4 Then printObj.Print ", ";
'        Next
'        printObj.Print
'        rs.MoveNext
'    Loop
'    rs.Close
'Next
'savy = printObj.CurrentY
'On Error Resume Next
'printObj.Line (0, yStart)-(printObj.ScaleWidth - 50, savy), , B
'On Error GoTo 0
'maxY = savy
''achtste finales
'i = getPntToek("achtste finaleplaats") + getPntToek("achtste finalepositie")
'If i > 0 Then
'    Fav_Finals 5, 4, "Achtste finales"
'    savy = printObj.CurrentY
'End If
'printObj.CurrentY = savy
''kwart finales
'i = getPntToek("kwart finaleplaats") + getPntToek("kwart finalepositie")
'If i > 0 Then
'    Fav_Finals 2, 4, "Kwart finales"
'    savy = printObj.CurrentY
'End If
'printObj.CurrentY = savy
''halve finales
'i = getPntToek("halve finaleplaats") + getPntToek("halve finalepositie")
'If i > 0 Then
'    Fav_Finals 3, 4, "Halve finales"
'    savy = printObj.CurrentY
'    maxY = savy
'End If
'printObj.CurrentY = savy
''kleine finale
'i = getPntToek("kleine finaleplaats") + getPntToek("kleine finalepositie")
'If i > 0 Then
'    bewYPos = printObj.CurrentY
'    Fav_Finals 7, 4, "Kleine finale"
'    savy = maxY
'    'maxY = savy
'    savX = 3
'Else
'    bewYPos = printObj.CurrentY
'    savX = 1
'End If
'
''finale
'i = getPntToek("finaleplaats") + getPntToek("finalepositie")
'If i > 0 Then
'    Fav_Finals 4, 4, "Finale", savy, savX
'    If savX = 3 Then
'        savX = 1
'        savy = printObj.CurrentY
'    Else
'        savy = bewYPos
'        savX = 3
'    End If
''    savy = printObj.CurrentY
'    maxY = savy
'End If
'printObj.CurrentY = savy
'Fav_Eindstand savy, savX
'Fav_Topscorers
'Set rs = Nothing
'printObj.Print
'printObj.Print
'End Sub
'
'Sub Fav_Topscorers()
'Dim aant As Integer
'Dim cols(5) As Integer
'Dim sqlstr As String
'Dim savy As Integer
'Dim savFntgr As Integer
'Dim rs As New ADODB.Recordset
'Dim i As Integer
'Dim J As Integer
'For i = 1 To 4
'    cols(i) = Int(printObj.ScaleWidth / 4) * (i - 1)
'Next
'cols(5) = printObj.ScaleWidth - 10
'sqlstr = "SELECT personen.rnaam, Count(voorspelling_ts.deelnem) AS aantal"
'sqlstr = sqlstr & " FROM voorspelling_ts LEFT JOIN personen ON voorspelling_ts.ts = personen.ID"
'sqlstr = sqlstr & " WHERE voorspelling_ts.deelnem In (select deelnemid from pooldeelnems where poolid= " & thisPool
'sqlstr = sqlstr & " ) GROUP BY personen.rnaam, voorspelling_ts.ts"
'sqlstr = sqlstr & " ORDER BY Count(voorspelling_ts.deelnem) DESC, personen.rnaam "
'rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'If rs.RecordCount > 0 Then
'    rs.MoveLast
'End If
'aant = rs.RecordCount
'i = 1
'J = 0
'
'printObj.CurrentX = favXpos
'If favYpos > voethoog - Int(aant / 4) * printObj.TextHeight("tekst") - 120 Then
'  heading1 = "Topscorers"
'  DoNewPage False, False, 0
'  favYpos = printObj.CurrentY
'Else
'  printObj.CurrentY = favYpos
'  headingText "Topscorers", False, False, favYpos, 0
'End If
'
'savy = printObj.CurrentY
'rs.MoveFirst
'
'Do While Not rs.EOF
'    printObj.CurrentX = cols(i)
'    If nz(rs!rnaam, "") > "" Then
'        printObj.Print rs!rnaam;
'    Else
'        printObj.Print "Niet ingevuld";
'    End If
'    printObj.CurrentX = cols(i + 1) - 500 - printObj.TextWidth(rs!Aantal)
'    printObj.Print rs!Aantal
'    J = J + 1
'    rs.MoveNext
'    If printObj.CurrentY > favYpos Then
'        favYpos = printObj.CurrentY
'    End If
'    If J > Int(aant / 4) - 1 Then
'        i = i + 1
'        J = 0
'        printObj.CurrentY = savy
'    End If
'Loop
'rs.Close
'Set rs = Nothing
'printObj.Line (cols(1), savy)-(cols(5) - 50, favYpos), , B
'
'End Sub
'
'Function GetRijAant(wedNum As Integer, team)
''om te bepalen of we naar een nieuw pagina moeten in de favorieten afdruk
'Dim sqlstr As String
'sqlstr = "SELECT wed, " & team
'sqlstr = sqlstr & " From voorspelling_finales"
'sqlstr = sqlstr & " WHERE deelnem In (select deelnemid from pooldeelnems where poolid =" & thisPool
'sqlstr = sqlstr & " ) GROUP BY wed, " & team
'sqlstr = sqlstr & " HAVING wed =" & wedNum
'sqlstr = sqlstr & " AND " & team & " >0"
'Dim rs As New ADODB.Recordset
'rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'If Not rs.RecordCount = 0 Then
'    rs.MoveLast
'End If
'GetRijAant = rs.RecordCount
'rs.Close
'Set rs = Nothing
'End Function
'
'Sub PrintEindStandFav(Plaats As String, col As Integer, rs As ADODB.Recordset, veld As String)
'Dim sqlstr As String
'Dim ypos As Integer
'Dim fntGr As Integer
'    ypos = printObj.CurrentY
'    fntGr = printObj.Font.Size
'    If rs.RecordCount > 0 Then
'        rs.MoveFirst
'        printObj.fontBold = True
'        printObj.CurrentX = col
'        printObj.Print Plaats
'        printObj.fontBold = False
'        Do While Not rs.EOF
'            printObj.CurrentX = col + 50
'            If nz(rs(veld), 0) = 0 Then
'                printObj.Print "Niet ingevuld";
'            Else
'                printObj.Print GetTeam(rs(veld));
'            End If
'            printObj.CurrentX = col + printObj.TextWidth("123456789012345") - printObj.TextWidth(rs!Aantal)
'            printObj.Print rs!Aantal;
'            fontSizing fntGr - 3
'            printObj.CurrentY = printObj.CurrentY + 30
'            printObj.Print "(" & Format(rs!Aantal / GetDeelnemAant(thisPool), "0.0%") & ")"
'            printObj.CurrentY = printObj.CurrentY - 30
'            fontSizing fntGr
'            rs.MoveNext
'        Loop
'    End If
'End Sub
'Sub Fav_Eindstand(savy As Integer, savX2 As Integer)
'Dim sqlstr As String
'Dim rs1 As New ADODB.Recordset
'Dim rs2 As New ADODB.Recordset
'Dim rs3 As New ADODB.Recordset
'Dim rs4 As New ADODB.Recordset
'Dim maxaant As Integer
'Dim savX As Integer
'Dim aantpos As Integer
'Dim startY As Integer
'Dim maxY As Integer
'Dim i As Integer
'Dim savFntgr As Integer
'Dim aantFav As Integer
'Dim cols(5) As Integer
'For i = 1 To 4
'    cols(i) = Int((printObj.ScaleWidth / 4) * (i - 1))
'Next
'
'cols(5) = printObj.ScaleWidth - 20
'
'    startY = savy
'
'    sqlstr = "SELECT kampioen, Count(pooldeelnems.deelnemID) AS aantal"
'    sqlstr = sqlstr & " From pooldeelnems"
'    sqlstr = sqlstr & " WHERE poolid = " & thisPool
'    sqlstr = sqlstr & " GROUP BY kampioen"
'    sqlstr = sqlstr & " ORDER BY count(deelnemID) desc"
'    rs1.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    sqlstr = "SELECT pltwee, Count(pooldeelnems.deelnemID) AS aantal"
'    sqlstr = sqlstr & " From pooldeelnems"
'    sqlstr = sqlstr & " WHERE poolid = " & thisPool
'    sqlstr = sqlstr & " GROUP BY pltwee"
'    sqlstr = sqlstr & " ORDER BY count(deelnemID) desc"
'    rs2.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    sqlstr = "SELECT pldrie, Count(pooldeelnems.deelnemID) AS aantal"
'    sqlstr = sqlstr & " From pooldeelnems"
'    sqlstr = sqlstr & " WHERE poolid = " & thisPool
'    sqlstr = sqlstr & " GROUP BY pldrie"
'    sqlstr = sqlstr & " ORDER BY count(deelnemID) desc"
'    rs3.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    sqlstr = "SELECT plvier, Count(pooldeelnems.deelnemID) AS aantal"
'    sqlstr = sqlstr & " From pooldeelnems"
'    sqlstr = sqlstr & " WHERE poolid = " & thisPool
'    sqlstr = sqlstr & " GROUP BY plvier"
'    sqlstr = sqlstr & " ORDER BY count(deelnemID) desc"
'    rs4.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    If rs1.RecordCount > 0 Then
'        rs1.MoveLast
'    End If
'    If savX2 = 1 Then
'        aantFav = 3
'    Else
'        aantFav = 3
'    End If
'
'    favXpos = cols(savX2)
'    maxaant = rs1.RecordCount
'    If rs2.RecordCount > 0 Then
'        rs2.MoveLast
'        If Not IsNull(rs2!pltwee) Then
'            aantFav = aantFav + 1
'            favXpos = cols(aantFav + 1)
'        End If
'    End If
'    If rs2.RecordCount > maxaant Then
'        maxaant = rs2.RecordCount
'    End If
'    If rs3.RecordCount > 0 Then
'        rs3.MoveLast
'        If Not IsNull(rs3!pldrie) Then
'            aantFav = 3
'            favXpos = cols(aantFav + 1)
'        End If
'    End If
'    If rs3.RecordCount > maxaant Then
'        maxaant = rs3.RecordCount
'    End If
'    If rs4.RecordCount > 0 Then
'        rs4.MoveLast
'        If Not IsNull(rs4!plvier) Then
'            aantFav = 0
'            favXpos = cols(1)
'        End If
'    End If
'    If rs4.RecordCount > maxaant Then
'        maxaant = rs4.RecordCount
'    End If
'    savFntgr = printObj.FontSize
'    printObj.FontSize = savFntgr - 3
'    maxY = maxaant * printObj.TextHeight("Q") + savy
'    printObj.FontSize = savFntgr
'    maxY = maxY + printObj.TextHeight("Q") + 50
'    If maxY > voethoog - 465 Then
'        heading1 = "Favorieten einduitslag"
'        DoNewPage False, False, aantFav
'        'maxY = printObj.CurrentY
'        savy = printObj.CurrentY
'        startY = savy
'        savFntgr = printObj.FontSize
'        printObj.FontSize = savFntgr - 3
'        maxY = maxaant * printObj.TextHeight("Q") + savy
'        printObj.FontSize = savFntgr
'        maxY = maxY + printObj.TextHeight("Q") + 50
'    Else
'      If savX2 = 3 Then
'        headingText "Favorieten einduitslag", False, False, savy, savX2 + 1
'      Else
'        headingText "Favorieten einduitslag", False, False, savy, savX2 - 1 ' 0 centreert tussenkop
'      End If
'      savy = printObj.CurrentY
'      startY = savy
'      savFntgr = printObj.FontSize
'      printObj.FontSize = savFntgr - 3
'      maxY = maxaant * printObj.TextHeight("Q") + savy
'      printObj.FontSize = savFntgr
'      maxY = maxY + printObj.TextHeight("Q") + 50
'    End If
'    If getPntToek("1e plaats(Kampioen)") Then
'        printObj.CurrentY = savy
'        PrintEindStandFav "kampioen", cols(savX2) + 10, rs1, "kampioen"
'        printObj.Line (cols(savX2), startY)-(cols(savX2 + 1) - 50, maxY), , B
'    End If
'    If getPntToek("2e plaats") Then
'        printObj.CurrentY = savy
'        PrintEindStandFav "2e plaats", cols(savX2 + 1) + 10, rs2, "plTwee"
'        printObj.Line (cols(savX2 + 1), startY)-(cols(savX2 + 2) - 50, maxY), , B
'    End If
'    If getPntToek("3e plaats") Then
'        printObj.CurrentY = savy
'        PrintEindStandFav "3e plaats", printObj.ScaleWidth / 2 + 10, rs3, "pldrie"
'        printObj.Line (cols(3), startY)-(cols(4) - 50, maxY), , B
'    End If
'    If getPntToek("4e plaats") Then
'        printObj.CurrentY = savy
'        PrintEindStandFav "4e plaats", (printObj.ScaleWidth / 4) * 3 + 10, rs4, "plvier"
'        printObj.Line (cols(4), startY)-(cols(5) - 50, maxY), , B
'    End If
'    favYpos = maxY
'    favXpos = 0
'End Sub
'Sub Fav_Finals(wedtype As Integer, cols As Integer, koptxt As String, Optional bewaarYpos As Integer, Optional posX As Integer)
'Dim sqlstr As String
'Dim rs As New ADODB.Recordset
'Dim savX As Integer
'Dim savy As Integer
'Dim aantpos As Integer
'Dim startY As Integer
'Dim col() As Integer
'Dim i As Integer
'Dim J As Integer
'Dim team As String
'Dim fld As field
'Dim maxrows As Integer
'Dim maxrows1 As Integer
'Dim savMaxRows As Integer
'Dim savMaxRows1 As Integer
'Dim ttlRows As Integer
'Dim maxFinpos As Integer
'ReDim col(cols + 1) As Integer
'    For i = 1 To cols
'        col(i) = (i - 1) * printObj.ScaleWidth / cols
'    Next
'    col(cols + 1) = printObj.ScaleWidth
'    savy = printObj.CurrentY
'    sqlstr = "Select * from qryWeds where  ksid = " & kampID
'    sqlstr = sqlstr & " and wedtype = " & wedtype
'    sqlstr = sqlstr & " ORDER BY wednum"
'    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    startY = savy
'    'startY = 945
'
'    If rs.RecordCount > 0 Then
'        savMaxRows = 0
'        maxrows = 0
'        rs.MoveFirst
'        'bepaal aantal regels dat nodig is
'        Do While Not rs.EOF
'            savMaxRows = maxrows + GetRijAant(rs!wedNum, "t1")
'            If maxrows < savMaxRows Then
'                maxrows = savMaxRows
'            End If
'            rs.MoveNext
'        Loop
'        rs.MoveFirst
'        'bepaal aantal regels dat nodig is
'        Do While Not rs.EOF
'            savMaxRows1 = maxrows1 + GetRijAant(rs!wedNum, "t2")
'            If maxrows1 < savMaxRows1 Then
'                maxrows1 = savMaxRows1
'            End If
'            rs.MoveNext
'        Loop
'        ttlRows = maxrows
'        If maxrows1 > ttlRows Then ttlRows = maxrows1
'        rs.MoveFirst
'        If startY + ttlRows * TextHeight("Q") > voethoog - 465 And wedtype <> 4 Then '(465 = hoogte van het tussenkopje)
'            heading1 = koptxt
'            If wedtype = 7 Then
'                DoNewPage False, False, 2
'                maxY = printObj.CurrentY
'                savy = maxY
'                startY = 480
'                nwPag = True
'            Else
'                DoNewPage False, False
'                maxY = printObj.CurrentY
'                savy = maxY
'                startY = savy
'                nwPag = False
'            End If
'        Else
'            If wedtype = klFinale Then
'                finYpos = printObj.CurrentY
'                headingText koptxt, False, False, maxY, 2
'            ElseIf wedtype = Finale Then
'                If getPntToek("kleine finaleplaats") + getPntToek("kleine finalepositie") > 0 Then
'                    If nwPag Then
'                        headingText koptxt, False, False, 480, 4
'                    Else
'                        headingText koptxt, False, False, finYpos, 4
'                    End If
'                Else
'                    headingText koptxt, False, False, bewaarYpos, 2
'                End If
'            Else
'                headingText koptxt, False, False, maxY
'            End If
'            savy = printObj.CurrentY
'            startY = savy
'        End If
'
'        i = 1
'        If wedtype = Finale Then
'            i = posX
'        End If
'        'If wedtype = 7 Then Stop
'        Do While Not rs.EOF
'            If i <= cols Then
'                printObj.CurrentY = savy
'            End If
'            fav_finalTeams "t1", "code1", rs, col(i)
'            If maxY < printObj.CurrentY Then maxY = printObj.CurrentY
'            i = i + 1
'            If i <= cols Then
'                printObj.CurrentY = savy
'            End If
'            fav_finalTeams "t2", "code2", rs, col(i)
'            If maxY < printObj.CurrentY Then maxY = printObj.CurrentY
'            i = i + 1
'
'            If wedtype = 7 And maxY < printObj.CurrentY Then
'                maxY = printObj.CurrentY
'            ElseIf wedtype = 4 Then
'                If printObj.CurrentY > maxY Then
'                    maxY = printObj.CurrentY
'                End If
'            End If
'            maxY = maxY + 50
'            If i = 5 Then
'                printObj.Line (col(1), startY)-(col(3) - 50, maxY), , B
'                printObj.Line (col(3), startY)-(col(5) - 50, maxY), , B
'            End If
'            If posX = 1 And i = 3 Then
'                printObj.Line (col(1), startY)-(col(3) - 50, maxY), , B
'            End If
'
'            rs.MoveNext
'            If i > cols Then
'                i = 1
'                printObj.CurrentY = maxY + 50
'                savy = printObj.CurrentY
'                maxY = savy
'                startY = maxY
'                favYpos = savy
'                favXpos = 0
'            End If
'
'        Loop
'
'    End If
'    rs.Close
'    Set rs = Nothing
'End Sub
'
'Sub fav_finalTeams(team As String, cod As String, rs As ADODB.Recordset, col)
'Dim rs1 As New ADODB.Recordset
'Dim savX As Integer
'Dim savy As Integer
'Dim aantpos As Integer
'Dim sqlstr As String
'Dim fntGr As Integer
'    aantpos = printObj.TextWidth("NIET INGEVULD  1")
'    sqlstr = "SELECT wed, " & team & ", Count(wed) AS ttl From voorspelling_finales"
'    sqlstr = sqlstr & " WHERE deelnem In (select deelnemid from pooldeelnems where poolid =" & thisPool
'    sqlstr = sqlstr & " ) GROUP BY wed, " & team
'    sqlstr = sqlstr & " HAVING wed=" & rs!wedNum
'    sqlstr = sqlstr & " AND " & team & " > 0"
'    sqlstr = sqlstr & " ORDER BY count(wed) desc"
'    rs1.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    printObj.CurrentX = col
'    printObj.Print rs(cod) & ": ";
'    savX = printObj.CurrentX
'    fntGr = printObj.Font.Size
'    Do While Not rs1.EOF
'        printObj.CurrentX = savX
'        If nz(rs1(team), "") = "" Then
'            printObj.Print "Niet ingevuld";
'        Else
'            printObj.Print GetTeam(rs1(team));
'        End If
'        printObj.CurrentX = col + aantpos - printObj.TextWidth(rs1!ttl)
'        printObj.Print rs1!ttl;
'        fontSizing fntGr - 3
'        printObj.CurrentY = printObj.CurrentY + 30
'        printObj.Print "(" & Format(rs1!ttl / GetDeelnemAant(thisPool), "0.0%") & ")"
'        fontSizing fntGr
'        printObj.CurrentY = printObj.CurrentY - 30
'        If maxY < printObj.CurrentY Then maxY = printObj.CurrentY
'        rs1.MoveNext
'    Loop
'    rs1.Close
'End Sub
'
'Private Sub printParticipantPools()
'Dim Dezedeeln As Integer
'Dim tkst$
'Dim tmpnaam$
'Dim KolomAant As Integer
'Dim i As Integer
'Dim K As Integer
'Dim LineXpos As Integer
'Dim LineYPos As Integer
'Dim newlinepos As Integer
'Dim TopMarg As Integer
'Dim pr As String
'Dim rsDeelnem As New ADODB.Recordset
'Dim rsDeelnemWeds As New ADODB.Recordset
'Dim rsDeelnGroepen As New ADODB.Recordset
'Dim rsDeelnFinales As New ADODB.Recordset
'Dim rsDeelnts As New ADODB.Recordset
'Dim rsDeelnEindstand As New ADODB.Recordset
'Dim rsDeelnOverig As New ADODB.Recordset
'Dim sqlstr As String
'Dim naamHeight As Integer
'Dim wedHoog As Integer
'Dim NaamHoog As Integer
'Dim posDatum As Integer
'Dim posTijd As Integer
'Dim posWedOms As Integer
'Dim posRust As Integer
'Dim PosEind As Integer
'Dim posToto As Integer
'Dim posPnt As Integer
'Dim wedYpos As Integer
'Dim wedKol As Integer
'Dim Helft As Integer
'Dim oldhelft As Integer
'Dim heeft8stFin As Boolean
'Dim savdat As Date
'Dim savWedType As Integer
'Dim kaderPos As Integer
'Dim deelnPag As Integer
'Dim grpWedsAant As Integer
'Dim nwKol As Boolean
'Dim grpPnt As Integer
'Dim grpPntTTL As Integer
'Dim grpPntposY As Integer
'Dim grpPntPosX As Integer
'Dim endEersteDeelnPos As Integer
'Dim tsYpos As Integer
'Dim wedPnt As Integer
'Dim ttl As Integer
'Dim ttlPosX As Integer
'Dim ttlPosY As Integer
'Dim grpwedsTtlPosX As Integer
'Dim grpwedsTtlPosY As Integer
'Dim ttlgrpWeds As Integer
'Dim Dagpnt As Integer
'Dim dagpntposX As Integer
'Dim dagpntposY As Integer
'Dim savXpos As Integer
'Dim savYpos As Integer
'Dim toernooiGestart As Boolean
'Dim aantalAfgedrukt As Integer
'Dim AantalOpPapier As Integer
'Dim prntReg As String
'    toernooiGestart = KSStarted()
'    If printObj.ScaleHeight <> Printer.ScaleHeight Then
'        Helft = Helft + printObj.TextHeight("W") * 2
'    End If
'    grpWedsAant = AantGrpWeds()
'    rot.Angle = 0
'    wedHoog = 9
'    NaamHoog = 11
'    rsDeelnem.Open sqlDeelnems(thisPool), cn, adOpenStatic, adLockReadOnly
'
'    If rsDeelnem.RecordCount = 0 Then
'        MsgBox "Geen deelnemers in deze pool", vbQuestion + vbOKOnly, "Deelnemers afdrukken"
'        Exit Sub
'    End If
'    KolomAant = 1
'    x% = 20
'    headerText = GetOrgNaam(thisPool) & " " & getTournamentInfo("toernooi") & " voetbalpool"
'    tkst$ = "Deelnemers en Voorspellingen"
'    heading1 = tkst$
'
'    InitPage True, False
'    fontSizing NaamHoog
'    printObj.CurrentY = printObj.CurrentY - 50
'    kophoog = printObj.CurrentY
'    TopMarg = printObj.CurrentY
'    AantalOpPapier = 2
'    If grpWedsAant <= 24 Then
'        AantalOpPapier = 3
'    End If
'    Helft = (voethoog - TopMarg) / AantalOpPapier
''    Helft = printObj.ScaleHeight / AantalOpPapier + 100 'printObj.CurrentY
'    fontSizing wedHoog
'    'Debug.Print printObj.FontSize, Printer.FontSize * printRatio
'    RegHeight% = printObj.TextHeight("x") '* printRatio
'    fontSizing NaamHoog
'    naamHeight = printObj.TextHeight("x") '* printRatio
'    If getTournamentInfo("groepen") > 4 Then
'        KolomAant = getTournamentInfo("groepen")
'    Else
'        KolomAant = 8
'    End If
'
'    kolwidth = Int((printObj.ScaleWidth / KolomAant) - 50)
'    printObj.FillStyle = vbFSTransparent
'    rsDeelnem.MoveFirst
'    fontSizing 8
'    posDatum = 50
'    posWedOms = posDatum + printObj.TextWidth("99-99:")
'    posRust = posWedOms + printObj.TextWidth("WWW-WWW")
'    PosEind = posRust + printObj.TextWidth("11-11")
'    posToto = PosEind + printObj.TextWidth("11-11")
'    posPnt = posToto + printObj.TextWidth("99")
'    fontSizing 12
'    deelnPag = 0
'    Do While Not rsDeelnem.EOF
'        If Me.lstCompetitorPools.Selected(rsDeelnem.AbsolutePosition - 1) Or Me.Option3 = True Then
'            showInfo True, "Afdrukken deelnemers", rsDeelnem!bijnaam, "Record " & rsDeelnem.AbsolutePosition & "/" & rsDeelnem.RecordCount
'
'            If deelnPag = 0 Then
'                printObj.CurrentY = TopMarg
'            Else
'                printObj.CurrentY = deelnPag * (Helft) + TopMarg
'            End If
'            LineYPos = printObj.CurrentY
'            printObj.CurrentX = 30
'            printObj.fontBold = True
'            fontSizing NaamHoog + 6
'            printObj.Print
'            wedYpos = printObj.CurrentY
'
'            printObj.Line (0, LineYPos)-(printObj.ScaleWidth - 10, wedYpos), &H127419, BF
'            printObj.CurrentY = LineYPos
'            printObj.ForeColor = vbWhite
'            iBKMode = SetBkMode(printObj.hdc, TRANSPARENT)
'            printObj.CurrentX = 30
'            printObj.Print rsDeelnem!bijnaam;
'            ttlPosX = printObj.ScaleWidth
'            ttlPosY = printObj.CurrentY
'            printObj.Print
'            printObj.fontBold = False
'            printObj.CurrentX = 50
'            printObj.ForeColor = 1
'            'groepswedstrijden
'            sqlstr = "Select * from qryDeelnWeds Where deelnem = " & rsDeelnem!deelnemID
'            rsDeelnemWeds.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'            fontSizing 10
'            printObj.fontBold = True
'            printObj.ForeColor = vbBlue
'            printObj.Print "Groepswedstrijden";
'            grpwedsTtlPosX = printObj.CurrentX
'            grpwedsTtlPosY = printObj.CurrentY
'            printObj.CurrentX = printObj.ScaleWidth * 0.75 + 50
'            printObj.Print "Finales";
'            printObj.ForeColor = 1
'            printObj.fontBold = False
'            printObj.FontItalic = True
'            fontSizing 8
'            'For i = 1 To 4
'                'printObj.CurrentX = printObj.ScaleWidth / 4 * i - printObj.TextWidth("pnt") - 50
'                'printObj.Print "pnt";
'            'Next
'            printObj.FontItalic = False
'            fontSizing 10
'            printObj.Print
'            printObj.Line (0, wedYpos - 10)-(printObj.ScaleWidth - 10, printObj.CurrentY + 10), , B
'            LineYPos = printObj.CurrentY + 10
'            printObj.CurrentY = LineYPos
'            fontSizing 8
'            LineXpos = 0
'            With rsDeelnemWeds
''                showInfo True, "Afdrukken deelnemers", rsDeelnem!bijnaam, "Record " & rsDeelnem.AbsolutePosition  & "/" & rsDeelnem.RecordCount, "Wedstrijden"
'                K = 0
'                If .RecordCount > 0 Then
'                    .MoveLast
'                    .MoveFirst
'                    wedKol = 1
'                    Do While Not .EOF
'                        printObj.CurrentX = LineXpos + posWedOms - printObj.TextWidth(Format(!datum, "d-m") & ":") - 50
'                        If savdat <> !datum Or printObj.CurrentY = LineYPos Then
'                            printObj.Print Format(!datum, "d-m"); ":";
'                            savdat = !datum
'                        End If
'                        If nz(!tm1, "") > "" And !wedtype = 1 Then
'                            pr = !tm1
'                        Else
'                            pr = nz(!code1, "")
'                        End If
'                        pr = pr & " - "
'                        If nz(!tm2, "") > "" And !wedtype = 1 Then
'                            pr = pr & !tm2
'                        Else
'                            pr = pr & !code2
'                        End If
'                        printObj.CurrentX = LineXpos + posWedOms
'                        printObj.Print pr;
'                        printObj.CurrentX = LineXpos + posRust
'                        printObj.Print !r1; "-"; !r2;
'                        printObj.CurrentX = PosEind + LineXpos
'                        printObj.Print !e1; "-"; !e2;
'                        printObj.CurrentX = LineXpos + posToto
'                        printObj.Print !toto;
'                        printObj.Print
'                        If newlinepos < printObj.CurrentY Then newlinepos = printObj.CurrentY
'                        rsDeelnemWeds.MoveNext
'                        If grpWedsAant < 25 Then
'                            nwKol = (.AbsolutePosition - 1) Mod (grpWedsAant / 3) = 0 '= Int(grpWedsAant / 2) Or .AbsolutePosition = grpWedsAant
'                        Else
'                            nwKol = (.AbsolutePosition - 1) Mod 16 = 0
'                        End If
'                        If nwKol Then
'                            printObj.CurrentY = LineYPos
'                            K = K + 1
'                            If (.AbsolutePosition - 1) = grpWedsAant Then K = 3
'                            LineXpos = (printObj.ScaleWidth / 4) * K
'                        End If
'                    Loop
'                End If
'                .Close
'            End With
'            printObj.Line (0, wedYpos)-(0, newlinepos)
'            For i = 1 To 4
'                printObj.Line (printObj.ScaleWidth / 4 * i - 20, LineYPos)-(printObj.ScaleWidth / 4 * i - 20, newlinepos)
'                printObj.Line (printObj.ScaleWidth / 4 * (i - 1) + posRust - 20, LineYPos)-(printObj.ScaleWidth / 4 * (i - 1) + posRust - 20, newlinepos)
'                printObj.Line (printObj.ScaleWidth / 4 * (i - 1) + PosEind - 20, LineYPos)-(printObj.ScaleWidth / 4 * (i - 1) + PosEind - 20, newlinepos)
'                printObj.Line (printObj.ScaleWidth / 4 * (i - 1) + posToto - 20, LineYPos)-(printObj.ScaleWidth / 4 * (i - 1) + posToto - 20, newlinepos)
'                printObj.Line (printObj.ScaleWidth / 4 * (i - 1) + posPnt - 20, LineYPos)-(printObj.ScaleWidth / 4 * (i - 1) + posPnt - 20, newlinepos)
'            Next
'            fontSizing 10
'            'groepstanden
''            showInfo True, "Afdrukken deelnemers", rsDeelnem!bijnaam, "Record " & rsDeelnem.AbsolutePosition + 1 & "/" & rsDeelnem.RecordCount, "Groepstanden"
'            printObj.Line (0, newlinepos)-(printObj.ScaleWidth, newlinepos)
'            printObj.Line (0, newlinepos)-(printObj.ScaleWidth - 10, newlinepos + printObj.TextHeight("Gr") + 10), , B
'            printObj.CurrentY = newlinepos + 10
'            printObj.CurrentX = 50
'            printObj.fontBold = True
'            printObj.ForeColor = vbBlue
'            printObj.Print "Groepstanden"
'            printObj.ForeColor = 1
'            printObj.fontBold = False
'            LineYPos = printObj.CurrentY
'            kolwidth = Int((printObj.ScaleWidth / KolomAant)) - 1
'            fontSizing 10
'            sqlstr = "Select * from voorspelling_groepstand Where deelnem = " & rsDeelnem!deelnemID
'            sqlstr = sqlstr & " ORDER BY groep"
'            rsDeelnGroepen.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'            'LineYPos = printObj.CurrentY - 10
'            K = 0
'            printObj.CurrentX = 50
'            Do While Not rsDeelnGroepen.EOF
'                printObj.FontUnderline = True
'                printObj.ForeColor = &H4000&
'                printObj.Print "Groep " & rsDeelnGroepen!groep
'                printObj.ForeColor = 1
'                printObj.FontUnderline = False
'
''                printObj.CurrentX = printObj.CurrentX + printObj.TextWidth("|00")
'                For i = 1 To 4
'                    printObj.CurrentX = kolwidth * K
'                    pr = GetTeam(rsDeelnGroepen("pos" & Format(i, "0")))
'                    If pr = "" Then pr = "?"
'                    printObj.Print i; ":"; pr
'                    If newlinepos < printObj.CurrentY Then newlinepos = printObj.CurrentY
'                Next
'                K = K + 1
'                printObj.Line (kolwidth * (K - 1), LineYPos)-(kolwidth * (K), newlinepos), , B
'                printObj.CurrentX = kolwidth * K + 100
'                printObj.CurrentY = LineYPos
'                rsDeelnGroepen.MoveNext
'            Loop
'
'            rsDeelnGroepen.Close
'            If grpWedsAant > 24 Then
'                printObj.CurrentX = grpPntPosX
'                printObj.CurrentY = newlinepos
'            Else
'                printObj.CurrentX = kolwidth * K
'            End If
'            'finales
'            newlinepos = printObj.CurrentY
'            printObj.Line (printObj.CurrentX, newlinepos)-(printObj.ScaleWidth, newlinepos)
'            printObj.CurrentY = newlinepos
'            LineXpos = 0
'            LineYPos = printObj.CurrentY
'            sqlstr = "Select * from qrydeelnemfinales WHERE deelnem=" & rsDeelnem!deelnemID
'            sqlstr = sqlstr & " AND wedtype = " & AchtsteFinale
'            sqlstr = sqlstr & " AND ksid= " & kampID
'            If rsDeelnFinales.State = adStateOpen Then
'                rsDeelnFinales.Close
'            End If
'            rsDeelnFinales.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'            If rsDeelnFinales.RecordCount > 0 Then
'                With rsDeelnFinales
'                    printObj.CurrentX = LineXpos + 20
'                    printObj.fontBold = True
'                    printObj.ForeColor = vbBlue
'                    printObj.Print "Achtste finales"
'                    printObj.ForeColor = 1
'                    printObj.fontBold = False
'                    Do While Not .EOF
'                        printObj.CurrentX = LineXpos + 50
'                        prntReg = Format(!wed, "0") & ": " & !tm1 & " - " & !tm2
'                        Do While printObj.TextWidth(prntReg) > printObj.ScaleWidth / 5 - 100
'                          prntReg = Left(prntReg, Len(prntReg) - 1)
'                        Loop
'                        printObj.Print prntReg;
'                        printObj.Print
'                        If .AbsolutePosition = 4 Then
'                            If LineYPos < printObj.CurrentY Then LineYPos = printObj.CurrentY
'                            LineXpos = LineXpos + printObj.ScaleWidth / 5
'                            printObj.CurrentY = newlinepos
'                            printObj.Print
'                        End If
'                        .MoveNext
'                    Loop
'                    .Close
'                End With
'                printObj.CurrentY = newlinepos
'            End If
'            If grpWedsAant > 24 Then
'                LineXpos = printObj.ScaleWidth / 5 * 2
'            Else
'                LineXpos = printObj.ScaleWidth / 2
'            End If
'            sqlstr = "Select distinct * from qrydeelnemfinales WHERE deelnem=" & rsDeelnem!deelnemID
'            sqlstr = sqlstr & " AND wedtype = " & KwartFinale
'            sqlstr = sqlstr & " AND ksid= " & kampID
'            If rsDeelnFinales.State <> 0 Then rsDeelnFinales.Close
'            rsDeelnFinales.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'            If rsDeelnFinales.RecordCount > 0 Then
'                With rsDeelnFinales
'                    printObj.CurrentX = LineXpos + 50
'                    printObj.fontBold = True
'                    printObj.ForeColor = vbBlue
'                    printObj.Print "Kwart finales"
'                    printObj.ForeColor = 1
'                    printObj.fontBold = False
'                    Do While Not .EOF
'                        printObj.CurrentX = LineXpos + 50
'                        prntReg = Format(!wed, "0") & ": " & !tm1 & " - " & !tm2
'                        Do While printObj.TextWidth(prntReg) > printObj.ScaleWidth / 5 - 100
'                          prntReg = Left(prntReg, Len(prntReg) - 1)
'                        Loop
'                        printObj.Print prntReg;
'                        printObj.Print
'                        If LineYPos < printObj.CurrentY Then LineYPos = printObj.CurrentY
'                        .MoveNext
'                    Loop
'                    .Close
'                End With
'                printObj.CurrentY = newlinepos
'            End If
'            If grpWedsAant > 24 Then
'                LineXpos = printObj.ScaleWidth / 5 * 3
'            Else
'                LineXpos = printObj.ScaleWidth / 4 * 3
'            End If
'            sqlstr = "Select DISTINCT * from qrydeelnemfinales WHERE deelnem=" & rsDeelnem!deelnemID
'            sqlstr = sqlstr & " AND wedtype = " & HalveFinale
'            sqlstr = sqlstr & " AND ksid= " & kampID
'            If rsDeelnFinales.State = adStateOpen Then
'                rsDeelnFinales.Close
'            End If
'            rsDeelnFinales.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'            If rsDeelnFinales.RecordCount > 0 Then
'                With rsDeelnFinales
'                    printObj.CurrentX = LineXpos + 50
'                    printObj.fontBold = True
'                    printObj.ForeColor = vbBlue
'                    printObj.Print "Halve finales"
'                    printObj.ForeColor = 1
'                    If LineYPos < printObj.CurrentY Then LineYPos = printObj.CurrentY
'                    printObj.fontBold = False
'                   ' printObj.Print
'                    Do While Not .EOF
'                        printObj.CurrentX = LineXpos + 50
'                        prntReg = Format(!wed, "0") & ": " & !tm1 & " - " & !tm2
'                        Do While printObj.TextWidth(prntReg) > printObj.ScaleWidth / 5 - 100
'                          prntReg = Left(prntReg, Len(prntReg) - 1)
'                        Loop
'                        printObj.Print prntReg; ' Format(!wed, "0"); ": "; !tm1; " - "; !tm2;
'                        printObj.Print
'                        .MoveNext
'                    Loop
'                    .Close
'                End With
'                If grpWedsAant > 24 Then
'                    printObj.CurrentY = newlinepos
'                End If
'            End If
'            If grpWedsAant > 24 Then
'                LineXpos = printObj.ScaleWidth / 5 * 4
'            Else
'                LineXpos = printObj.ScaleWidth / 4 * 3
'            End If
'            sqlstr = "Select * from qrydeelnemfinales WHERE deelnem=" & rsDeelnem!deelnemID
'            sqlstr = sqlstr & " AND wedtype = " & klFinale
'            sqlstr = sqlstr & " AND ksid= " & kampID
'            If rsDeelnFinales.State = adStateOpen Then
'                rsDeelnFinales.Close
'            End If
'            rsDeelnFinales.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'
'            If rsDeelnFinales.RecordCount > 0 Then
'                With rsDeelnFinales
'                    printObj.CurrentX = LineXpos + 50
'                    printObj.fontBold = True
'                    printObj.ForeColor = vbBlue
'                    printObj.Print "3e plaats"
'                    printObj.ForeColor = 1
'                    printObj.fontBold = False
'                    Do While Not .EOF
'                        printObj.CurrentX = LineXpos + 50
'                        prntReg = Format(!wed, "0") & ": " & !tm1 & " - " & !tm2
'                        Do While printObj.TextWidth(prntReg) > printObj.ScaleWidth / 5 - 100
'                          prntReg = Left(prntReg, Len(prntReg) - 1)
'                        Loop
'                        printObj.Print prntReg;
'                        printObj.Print
'                        If LineYPos < printObj.CurrentY Then LineYPos = printObj.CurrentY
'                        .MoveNext
'                    Loop
'                    printObj.CurrentY = printObj.CurrentY + 120
'                    printObj.Line (printObj.ScaleWidth / 5 * 4, printObj.CurrentY - 20)-(printObj.ScaleWidth - 10, printObj.CurrentY - 20)
'                    printObj.CurrentY = printObj.CurrentY + 10
'                End With
'            End If
'            sqlstr = "Select DISTINCT * from qrydeelnemfinales WHERE deelnem=" & rsDeelnem!deelnemID
'            sqlstr = sqlstr & " AND wedtype = " & Finale
'            sqlstr = sqlstr & " AND ksid= " & kampID
'            If rsDeelnFinales.State = adStateOpen Then
'                rsDeelnFinales.Close
'            End If
'            rsDeelnFinales.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'            If rsDeelnFinales.RecordCount > 0 Then
'                With rsDeelnFinales
'                    printObj.CurrentX = LineXpos + 50
'                    printObj.fontBold = True
'                    printObj.ForeColor = vbBlue
'                    printObj.Print "Finale"
'                    printObj.ForeColor = 1
'                    printObj.fontBold = False
'                    Do While Not .EOF
'                        printObj.CurrentX = LineXpos + 50
'                        prntReg = Format(!wed, "0") & ": " & !tm1 & " - " & !tm2
'                        Do While printObj.TextWidth(prntReg) > printObj.ScaleWidth / 5 - 100
'                          prntReg = Left(prntReg, Len(prntReg) - 1)
'                        Loop
'                        printObj.Print prntReg;
'                        printObj.Print
'                        If LineYPos < printObj.CurrentY Then LineYPos = printObj.CurrentY
'                    .MoveNext
'                    Loop
'                    .Close
'                End With
'            End If
'            If grpWedsAant > 24 Then
'                For i = 2 To 4
'                    printObj.Line (printObj.ScaleWidth / 5 * i, newlinepos)-(printObj.ScaleWidth / 5 * i, LineYPos)
'                Next
'            End If
'            printObj.Line (0, newlinepos)-(printObj.ScaleWidth - 10, LineYPos), , B
'            'uitslag
'            LineYPos = printObj.CurrentY + 50
'            LineXpos = 50
'            printObj.CurrentX = LineXpos
'            printObj.CurrentY = LineYPos
'            printObj.fontBold = True
'            printObj.ForeColor = vbBlue
'            printObj.Print "Eindstand"
'            printObj.ForeColor = 1
'            printObj.fontBold = False
'            printObj.CurrentX = LineXpos
'            pr = GetTeam(nz(rsDeelnem!kampioen, 0))
'            If pr = "" Then pr = "?"
'            printObj.Print "1: "; pr
'            printObj.CurrentX = LineXpos
'            If getPntToek("2e plaats") > 0 Then
'                pr = GetTeam(nz(rsDeelnem!pltwee, 0))
'                If pr = "" Then pr = "?"
'                printObj.Print "2: "; pr
'            Else
'                printObj.Print
'            End If
'            printObj.CurrentX = LineXpos
'            If getPntToek("3e plaats") > 0 Then
'                pr = GetTeam(nz(rsDeelnem!pldrie, 0))
'                If pr = "" Then pr = "?"
'                printObj.Print "3: "; pr
'            Else
'                printObj.Print
'            End If
'            printObj.CurrentX = LineXpos
'            If getPntToek("4e plaats") > 0 Then
'              pr = GetTeam(nz(rsDeelnem!plvier, 0))
'              If pr = "" Then pr = "?"
'              printObj.Print "4: "; pr
'            Else
'                printObj.Print
'            End If
'            newlinepos = printObj.CurrentY
'            If deelnPag = 1 Then
'                oldhelft = Helft
'            End If
'            printObj.Line (0, LineYPos - 10)-(printObj.ScaleWidth / 8, newlinepos), , B
'            'topscorers
'            LineXpos = printObj.ScaleWidth / 8 + 50
'            printObj.CurrentX = LineXpos
'            printObj.CurrentY = LineYPos
'
'            printObj.fontBold = True
'            printObj.CurrentX = LineXpos + 50
'            printObj.ForeColor = vbBlue
'            printObj.Print "Topscorer";
'            If getPntToek("doelpunten topscorer 1") > 0 Then
'                printObj.CurrentX = (printObj.ScaleWidth / 5 * 2) - printObj.TextWidth("doelp") - 100
'                printObj.Print "doelp"
'            Else
'                printObj.Print
'            End If
'            printObj.ForeColor = 1
'            tsYpos = printObj.CurrentY
'            kaderPos = printObj.ScaleWidth / 5 * 2
'            printObj.Line (LineXpos, LineYPos - 10)-(kaderPos - 10, newlinepos), , B
'            printObj.fontBold = False
'            printObj.CurrentY = tsYpos
'            sqlstr = "Select * from voorspelling_ts WHERE deelnem = " & rsDeelnem!deelnemID
'            sqlstr = sqlstr & " ORDER BY tsNR"
'            rsDeelnts.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'            Do While Not rsDeelnts.EOF
'                printObj.CurrentX = LineXpos + 50
'                pr = getSpelerNaam(nz(rsDeelnts!ts, 0))
'                printObj.Print pr;
'                printObj.CurrentX = kaderPos - printObj.TextWidth(Format(rsDeelnts!dp, "0")) - 150
'                If getPntToek("doelpunten topscorer 1") > 0 Then
'                    If rsDeelnts!dp > -1 Then
'                      printObj.Print Format(rsDeelnts!dp, 0)
'                    Else
'                        printObj.Print
'                    End If
'                Else
'                    printObj.Print
'                End If
'                rsDeelnts.MoveNext
'            Loop
'            rsDeelnts.Close
'            'overige
'            LineXpos = kaderPos + 20
'            kaderPos = printObj.ScaleWidth - 30
'            printObj.Line (LineXpos, LineYPos - 10)-(kaderPos, newlinepos), , B
'            sqlstr = "Select * from qryDeelnVoorspAant WHERE deelnem = " & rsDeelnem!deelnemID
'            rsDeelnOverig.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'            printObj.CurrentY = LineYPos
'            printObj.CurrentX = LineXpos + 50
'            printObj.fontBold = True
'            printObj.ForeColor = vbBlue
'            printObj.Print "Overigen ";
'            printObj.ForeColor = 1
'            LineXpos = printObj.CurrentX
'            printObj.fontBold = False
'            With rsDeelnOverig
'                Do While Not .EOF
'                    printObj.CurrentX = LineXpos + 50
'                    printObj.Print !omschrijving; ": ";
'                    printObj.Print !Aantal
'                    .MoveNext
'                Loop
'                .Close
'            End With
'            printObj.DrawWidth = 2
'            printObj.Line (0, printObj.CurrentY + 50)-(printObj.ScaleWidth - 10, printObj.CurrentY + 50)
'            aantalAfgedrukt = aantalAfgedrukt + 1
'        End If 'deeln selected
'        rsDeelnem.MoveNext
'        printObj.CurrentX = 0
'        If Not rsDeelnem.EOF Then
'            If Me.lstCompetitorPools.Selected(rsDeelnem.AbsolutePosition - 1) Or Me.Option3 = True Then
'                If deelnPag = AantalOpPapier - 1 Then
'                    'printObj.Line (0, Helft + 200)-(printObj.ScaleWidth - 10, endEersteDeelnPos + 50), , B
'                    deelnPag = 0
'                    newlinepos = 0
'                    'Exit Do
'                    If Not rsDeelnem.EOF Then DoNewPage False, False
'                Else
'                    endEersteDeelnPos = printObj.CurrentY
'                    If aantalAfgedrukt > 0 Then deelnPag = deelnPag + 1
'
'                    If aantalAfgedrukt Mod (AantalOpPapier - 1) = 0 And aantalAfgedrukt > 0 Then
''                        Debug.Print "test"
'                    End If
'                    printObj.Line (0, printObj.CurrentY + 50)-(printObj.ScaleWidth - 10, endEersteDeelnPos + 50)
'                    'printObj.Line (0, TopMarg)-(printObj.ScaleWidth - 10, endEersteDeelnPos + 50), , B
'                End If
'                printObj.DrawWidth = 1
'            End If
'        End If
'    Loop
'    rsDeelnem.Close
'    showInfo False
'End Sub
'
'Private Sub btnPrntAllAfterDay_Click()
'Dim i As Integer
'Dim curWed As Integer
'Dim savdat As Date
'Dim msg As String
''stand in toernooi
'Me.vscrlTM.value = GetMyNum(GetLastPlayed)
'msg = "Voorspellingen afgedrukt"
'If Me.vscrlTM.value > 0 Then
'  msg = "Dagstanden, grafiek en voorspellingen afgedrukt"
'  showInfo True, "Afdrukken", "Stand van zaken in toernooi", "Wedstrijd: " & Me.vscrlTM.value
'  DoEvents
'  optPrintDoc_Click 4
'  btnPrint_Click 0
'  'stand op punten
'  DoEvents
'  optPrintDoc_Click 2
'  Me.ScoreVolg(1) = True
'  showInfo True, "Afdrukken", "Stand op punten", "Wedstrijd: " & Me.vscrlTM.value
'  btnPrint_Click 0
'  'stand alfabetisch
'  Screen.MousePointer = vbHourglass
'  DoEvents
'  optPrintDoc_Click 2
'  Me.ScoreVolg(0) = True
'  showInfo True, "Afdrukken", "Stand alfabetisch", "Wedstrijd: " & Me.vscrlTM.value
'  btnPrint_Click 0
'  'punten per wedstrijd alfabetisch
'  DoEvents
'  optPrintDoc_Click 6
'  Me.ScoreVolg(0) = True
'  showInfo True, "Afdrukken", "Punten per wedstrijd", "Wedstrijd: " & GetLastPlayed
'  toMatch = getLastMatchPlayed(cn)
'  btnPrint_Click 0
'  'punten opbouw alfabetisch
'  DoEvents
'  optPrintDoc_Click 8
'  Me.ScoreVolg(0) = True
'  Me.optLandscape = True
'  showInfo True, "Afdrukken", "Puntenopbouw", "Wedstrijd: " & GetLastPlayed
'  btnPrint_Click 0
'  'grafiek alfabetisch
'  DoEvents
'  optPrintDoc_Click 5
'  Me.ScoreVolg(0) = True
'  showInfo True, "Afdrukken", "Grafiek", "Wedstrijd: " & Me.vscrlTM.value
'  btnPrint_Click 0
'End If
''voorspellingen
'curWed = GetMyNum(GetLastPlayed)
'If curWed < GetWedAant(kampID) Then
'    savdat = getWedDatum(GetWedNum(curWed + 1))
'    For i = curWed + 1 To GetWedAant(kampID)
'        If Format(getWedDatum(GetWedNum(i)), "d-m-yyyy") = Format(savdat, "d-m-yyyy") Then
'            optPrintDoc_Click 7
'            Me.vscrlVoor.value = i
'            showInfo True, "Afdrukken", "Voorspelling", "Wedstrijd: " & i
'            btnPrint_Click 0
'        End If
'    Next
'End If
'showInfo False
'Screen.MousePointer = vbDefault
'MsgBox msg, vbOKOnly + vbInformation, "Afdrukken"
'End Sub
'
'Sub EindStandAfdrukken()
'Dim i As Integer
'Dim curWed As Integer
'Dim savdat As Date
''stand in toernooi
'If MsgBox("Voor alle deelnemers afdrukken?", vbYesNo, "Eindstand") = vbYes Then
'    Me.Copies = getAantalUniekeDeelnems()
'End If
'Me.vscrlTM.value = GetMyNum(GetLastPlayed)
'showInfo True, "Afdrukken", "Eindstand toernooi", "Wedstrijd: " & Me.vscrlTM.value
'DoEvents
'optPrintDoc_Click 4
'Me.chkDblSide.value = 0
'btnPrint_Click 0
''stand op punten
'DoEvents
'optPrintDoc_Click 2
'Me.ScoreVolg(1) = True
'showInfo True, "Afdrukken", "Stand op punten", "Wedstrijd: " & Me.vscrlTM.value
'Me.chkDblSide.value = 0
'btnPrint_Click 0
''punten per wedstrijd alfabetisch
'DoEvents
'optPrintDoc_Click 6
'Me.ScoreVolg(0) = True
'Me.chkDblSide.value = 0
'showInfo True, "Afdrukken", "Punten per wedstrijd", "Wedstrijd: " & Me.vscrlTM.value
'btnPrint_Click 0
''punten opbouw alfabetisch
'DoEvents
'optPrintDoc_Click 8
'Me.ScoreVolg(0) = True
'Me.optLandscape = True
'Me.chkDblSide.value = 0
'showInfo True, "Afdrukken", "Puntenopbouw", "Wedstrijd: " & GetLastPlayed
'btnPrint_Click 0
''grafiek alfabetisch
'DoEvents
'optPrintDoc_Click 5
'Me.ScoreVolg(0) = True
'Me.chkDblSide.value = 0
'showInfo True, "Afdrukken", "Grafiek", "Wedstrijd: " & Me.vscrlTM.value
'btnPrint_Click 0
'
''klaar
'showInfo False
'Screen.MousePointer = vbDefault
'MsgBox "Eindstand afgedrukt", vbOKOnly + vbInformation, "Afdrukken"
'
'End Sub
'
'Private Sub btnFinalPlayerPrint_Click()
'    EindStandAfdrukken
'End Sub
'
'Private Sub cmbPrinters_Click()
'   Dim prntr As Printer
'
'   For Each prntr In Printers
'      If cmbPrinters.List(cmbPrinters.ListIndex) = prntr.DeviceName Then
'         Set Printer = prntr
'      End If
'   Next
'End Sub
'
Sub btnPrint_Click(Index As Integer)
Dim i As Integer
Dim printSelect As Integer

  Dim prntr As Printer

  For Each prntr In Printers
    If cmbPrinters.List(cmbPrinters.ListIndex) = prntr.DeviceName Then
      Set Printer = prntr
    End If
  Next

  Screen.MousePointer = vbHourglass
  If Me.optPortrait Then
    Printer.Orientation = vbPRORPortrait
  Else
    Printer.Orientation = vbPRORLandscape
  End If
  Init
  If Index = 0 Then
    Set printobj = Printer
    If printobj.Duplex <> 0 Then
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
    Printer.Copies = Me.upDnCopies.value
    'printRatio = 1
  Else
      Me.Visible = False
      printPrev.Show
      If printPrev.printPages.UBound = 0 Then
          Set printobj = printPrev.printPages(0)
      End If

  End If
  Set rotater.Device = printobj
  'Meter.Value = Meter.Min
  For i = 0 To 8
      If Me.optPrintDoc(i).value = True Then
          printSelect = i
          Exit For
      End If
  Next
  DoEvents
  printobj.Font = "Times New Roman"
  Select Case printSelect
  Case 0
      printPoolForms
  Case 1
      'printParticipantPools
  Case 2
      'Favorieten
      'printFavourites
  Case 3
      'voorspellingen voor wedstrijd
      'printMatchPredictions Me.upDnForMatch
  Case 4
      'Stand in pool
     ' printPoolStandings Me.ScoreVolg(0), val(Me.upDnToMatch)
  Case 5
      'samenvatting stand
      'printPoolPoints Me.ScoreVolg(0)
  Case 6
      'punten per wedstrijd
      'printPoolPointsPerMatch
  Case 7
      'printSkyline
  Case 8
      'toernooi stand
      'printTournamentStandings toMatch
  End Select

  'Melding.Visible = False
  'Picture1.Visible = True
  DoEvents
  If Index = 0 Then
      Printer.EndDoc
  Else
      printPrev.pageContent.PaintPicture printobj.Image, 0, 0, printobj.Width, printobj.Height
      Set printobj = Nothing
  End If
  Screen.MousePointer = Default

End Sub

'Sub printTournamentStandings(toMatch As Integer)
'Dim kopje As String
'    headerText = GetOrgNaam(thisPool) & " " & getTournamentInfo("toernooi") & " voetbalpool - Stand van zaken"
'    kopje = Format(GetWedInfo(toMatch, "datum"), "dddd d mmmm") & ": "
'    kopje = kopje & GetWedInfo(toMatch, "naam1") & " vs " & GetWedInfo(toMatch, "naam2")
'    heading1 = "Na wedstrijd " & toMatch & ", " & kopje
'    InitPage False, True
'    tnWeds
'    tnGroepStanden
'    tnFinales
'    prnTopScorers
'
'    prAantallen toMatch
'
'End Sub
'
'Sub prnTopScorers()
'Dim sqlstr As String
'Dim rs As New ADODB.Recordset
'Dim rsED As New ADODB.Recordset 'voor de eigen doelpunten
'Dim i As Integer
'Dim grps As Integer
'Dim colNu As Integer
'Dim numpos As Integer
'Dim datPos As Integer
'Dim wedPos As Integer
'Dim uitslPos As Integer
'Dim newYpos As Integer
'Dim ypos As Integer
'Dim aantpos As Integer
'Dim col(5) As Integer
'    col(0) = 4
'    col(1) = printObj.ScaleWidth / 5
'    col(2) = printObj.ScaleWidth / 5 * 2
'    col(3) = printObj.ScaleWidth / 5 * 3
'    col(4) = printObj.ScaleWidth / 5 * 4
'    col(5) = printObj.ScaleWidth
'    aantpos = printObj.ScaleWidth / 5
'    sqlstr = "select rnaam, afkort, count(rnaam) as aantal from qrywedverloop"
'    sqlstr = sqlstr & " WHERE gebeurtenis <= 2"
'    sqlstr = sqlstr & " AND ksid = " & kampID
'    sqlstr = sqlstr & " GROUP BY rnaam, afkort"
'    sqlstr = sqlstr & " ORDER BY count(rnaam) DESC, rnaam"
'    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    sqlstr = "select rnaam, afkort, count(rnaam) as aantal from qrywedverloop"
'    sqlstr = sqlstr & " WHERE gebeurtenis = 3"
'    sqlstr = sqlstr & " AND ksid = " & kampID
'    sqlstr = sqlstr & " GROUP BY rnaam, afkort"
'    sqlstr = sqlstr & " ORDER BY count(rnaam) DESC, rnaam"
'    rsED.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    If rsED.RecordCount > 0 Then
'        rsED.MoveLast
'        rsED.MoveFirst
'    End If
'    If rs.RecordCount > 0 Then
'        rs.MoveLast
'        rs.MoveFirst
'    End If
'    If rs.RecordCount + rsED.RecordCount > 0 Then
'        fontSizing 12
'        ypos = printObj.CurrentY
'        printObj.ForeColor = vbBlue
'        printObj.fontBold = True
'        printObj.Print "Topscorers tot nu toe: "
'        ypos = printObj.CurrentY
'        printObj.fontBold = False
'        printObj.ForeColor = 1
'        fontSizing 8
'        Do While Not rs.EOF
'            i = i + 1
'            printObj.CurrentX = col(colNu)
'            printObj.Print FirstPart(rs!rnaam) & " (" & LCase(rs!afkort) & ")";
'            printObj.CurrentX = col(colNu) + aantpos - printObj.TextWidth("1234567890")
'            printObj.Print rs!Aantal
'
'
'            rs.MoveNext
'            If i = Int((rs.RecordCount + rsED.RecordCount + 1) / 5) + 1 Then
'                i = 0
'                colNu = colNu + 1
'                newYpos = printObj.CurrentY
'                printObj.CurrentY = ypos
'            End If
'        Loop
'        If rsED.RecordCount > 0 Then
'            printObj.ForeColor = vbBlue
'            printObj.fontBold = True
'            i = i + 1
'            printObj.CurrentX = col(colNu)
'            printObj.Print "Eigen doelpunten:"
'            If i = Int((rs.RecordCount + rsED.RecordCount + 1) / 5) + 1 Then
'                i = 0
'                colNu = colNu + 1
'                newYpos = printObj.CurrentY
'                printObj.CurrentY = ypos
'            End If
'            printObj.fontBold = False
'            printObj.ForeColor = 1
'            Do While Not rsED.EOF
'                i = i + 1
'                printObj.CurrentX = col(colNu)
'                printObj.Print FirstPart(rsED!rnaam) & " (" & LCase(rsED!afkort) & ")";
'                printObj.CurrentX = col(colNu) + aantpos - printObj.TextWidth("1234567890")
'                printObj.Print rsED!Aantal
'
'
'                rsED.MoveNext
'                If i = Int((rs.RecordCount + rsED.RecordCount + 1) / 5) + 1 Then
'                    i = 0
'                    colNu = colNu + 1
'                    newYpos = printObj.CurrentY
'                    printObj.CurrentY = ypos
'                End If
'            Loop
'            rsED.Close
'        End If
'        rs.Close
'        printObj.Line (0, ypos)-(printObj.ScaleWidth - 50, newYpos), , B
'        printObj.CurrentY = newYpos
'        printObj.Print
'    End If
'End Sub
'
'Sub prAantallen(toMatch As Integer)
'Dim ypos As Integer
'Dim prStr As String
'Dim col(6) As Integer
'    col(0) = 0
'    col(1) = printObj.ScaleWidth / 6
'    col(2) = printObj.ScaleWidth / 6 * 2
'    col(3) = printObj.ScaleWidth / 6 * 3
'    col(4) = printObj.ScaleWidth / 6 * 4
'    col(5) = printObj.ScaleWidth / 6 * 5
'    col(6) = printObj.ScaleWidth - 50
'    fontSizing 12
'    printObj.ForeColor = vbBlue
'    printObj.fontBold = True
'    printObj.Print "Statistieken"
'    ypos = printObj.CurrentY
'    printObj.fontBold = False
'    printObj.ForeColor = 1
'    fontSizing 10
'    printObj.CurrentX = col(0)
'    prStr = "Doelpunten: " & Format(getAantal(toMatch, 1) + getAantal(toMatch, 2) + getAantal(toMatch, 3), pntFormat)
'    printObj.Print prStr;
'    printObj.CurrentX = col(1)
'    prStr = "Penalties: " & Format(getAantal(toMatch, 1) + getAantal(toMatch, 6), pntFormat)
'    printObj.Print prStr;
'    printObj.CurrentX = col(2)
'    prStr = "Gele kaarten: " & Format(getAantal(toMatch, 4), pntFormat)
'    printObj.Print prStr;
'    printObj.CurrentX = col(3)
'    prStr = "Rode kaarten: " & Format(getAantal(toMatch, 5), pntFormat)
'    printObj.Print prStr;
'    printObj.CurrentX = col(4)
'    prStr = "Gelijke spelen: " & Format(getAantalGelijkeSpelen(toMatch), pntFormat)
'    printObj.Print prStr;
'    printObj.CurrentX = col(5)
'    prStr = "Eigen doelpunten: " & Format(getAantal(toMatch, 3), pntFormat)
'    printObj.Print prStr
'    printObj.ForeColor = vbBlue
'    printobj.FontItalic = True
'    Centreer GetDeelnemAant(thisPool) & " deelnemers aan de pool"
'    printObj.Print
'    printobj.FontItalic = False
'    printObj.ForeColor = 1
'    printObj.Line (col(0), ypos)-(col(6), printObj.CurrentY), , B
'
'End Sub
'
'Sub tnFinales()
'Dim sqlstr As String
'Dim rs As New ADODB.Recordset
'Dim rsUitsl As New ADODB.Recordset
'Dim i As Integer
'Dim grps As Integer
'Dim col(5) As Integer
'Dim colNu As Integer
'Dim numpos As Integer
'Dim datPos As Integer
'Dim wedPos As Integer
'Dim vsPos As Integer
'Dim uitslPos As Integer
'Dim newYpos As Integer
'Dim ypos As Integer
'Dim topYpos As Integer
'Dim wed As Integer
'Dim uitsl As String
'Dim colNr As Integer
'Dim grpAant As Integer
'grpAant = getTournamentInfo("groepen")
'    col(0) = 20
'    col(1) = printObj.ScaleWidth / 3 + col(0)
'    col(2) = printObj.ScaleWidth / 3 * 2 + col(0)
'    col(3) = printObj.ScaleWidth
'    col(4) = printObj.ScaleWidth / 6 + col(0)
'    col(5) = printObj.ScaleWidth / 2 + col(0)
'    sqlstr = "Select * from qryWeds "
'    sqlstr = sqlstr & " WHERE ksid = " & kampID
'    sqlstr = sqlstr & " AND wedtype <> 1"
'    sqlstr = sqlstr & " order by mynum, wednum"
'    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    printObj.fontBold = True
'    fontSizing 12
'    printObj.ForeColor = vbBlue
'    printObj.Print "Finales"
'    topYpos = printObj.CurrentY
'    colNr = 0
'    printObj.CurrentX = col(colNr)
'    fontSizing 10
'    If grpAant > 4 Then
'        printObj.Print "Achtste finales";
'        colNr = colNr + 1
'        printObj.CurrentX = col(colNr)
'    End If
'    printObj.Print "Kwart finales";
'    colNr = colNr + 1
'    printObj.CurrentX = col(colNr)
'    printObj.Print "Halve finales";
'    If colNr < 2 Then
'        colNr = colNr + 1
'        printObj.CurrentX = col(colNr)
'        printObj.Print "Finale";
'    End If
'    ypos = printObj.CurrentY
'    printObj.fontBold = False
'    printObj.ForeColor = 1
'    fontSizing 8
'    numpos = printObj.TextWidth("00")
'    datPos = numpos + printObj.TextWidth("0")
'    wedPos = datPos + printObj.TextWidth("za 29 jun 20u:")
'    vsPos = wedPos + printObj.TextWidth("MEX")
'    uitslPos = col(1) - printObj.TextWidth("0-0(0-0)nvl:0-0(mexxx)")
'    printObj.Print
'    ypos = printObj.CurrentY
'    Do While Not rs.EOF
'
'        wed = rs!wedtype
'
'        Select Case wed
'        Case AchtsteFinale
'            If grpAant > 4 Then
'                colNu = 0
'            End If
'        Case KwartFinale
'            If grpAant > 4 Then
'                colNu = 1
'            Else
'                colNu = 0
'            End If
'        Case Finale
'            colNu = 2
'            If grpAant <= 4 Then
'                printObj.CurrentY = ypos
'            End If
'        Case Else
'            If grpAant > 4 Then
'                colNu = 2
'            Else
'                colNu = 1
'            End If
'        End Select
'        printObj.CurrentX = col(colNu) + numpos - printObj.TextWidth(Format(rs!mynum, "0"))
'        printObj.Print Format(rs!mynum, "0");
'        printObj.CurrentX = col(colNu) + wedPos - printObj.TextWidth(Format(rs!tijd, "ddd d mmm HHu") & ": ")
'        printObj.Print Format(rs!datum, "ddd d mmm"); tijdFormat(rs!tijd, True); ": "; ' : , " HHu"); ": ";
'        printObj.CurrentX = col(colNu) + wedPos
'        If nz(rs!tm1, "") > "" Then
'            printObj.Print rs!tm1;
'        Else
'            printObj.Print rs!code1;
'        End If
'        printObj.CurrentX = col(colNu) + vsPos
'
'        If nz(rs!tm2, "") > "" Then
'            printObj.Print " - "; rs!tm2;
'        Else
'            printObj.Print " - "; rs!code2;
'        End If
'        printObj.CurrentX = col(colNu) + uitslPos
'        If WedGespeeld(rs!wedNum) Then
'            printObj.Print GetWedUitsl(rs!wedNum)
'        Else
'            printObj.Print
'        End If
'        rs.MoveNext
'        If Not rs.EOF Then
'            If rs!wedtype <> wed Then
'                If newYpos < printObj.CurrentY Then
'                    newYpos = printObj.CurrentY
'                End If
'                If rs!wedtype <> klFinale And rs!wedtype <> Finale Then
'                    printObj.CurrentY = ypos
'                Else
'                    printObj.fontBold = True
'                    fontSizing 12
'                    printObj.ForeColor = vbBlue
'                    printObj.CurrentX = col(2)
'                    If rs!wedtype = klFinale Then
'                        printObj.Print "Derde plaats"
'                    ElseIf grpAant > 4 Then
'                        printObj.CurrentX = col(2)
'                        printObj.Print "Finale"
'                    End If
'                    printObj.fontBold = False
'                    printObj.ForeColor = 1
'                    fontSizing 8
'                End If
'            End If
'        End If
'
'    Loop
'    printObj.Line (col(0) - 20, topYpos)-(col(1) - 50, newYpos), , B
'    printObj.Line (col(1) - 20, topYpos)-(col(2) - 50, newYpos), , B
'    printObj.Line (col(2) - 20, topYpos)-(col(3) - 50, newYpos), , B
'
'    printObj.CurrentY = newYpos
'    printObj.Print
'End Sub
'
'
'Sub tnGroepStanden()
'Dim sqlstr As String
'Dim rsGrp As New ADODB.Recordset
'Dim i As Integer
'Dim grps As Integer
'Dim col(4) As Integer
'Dim colNu As Integer
'Dim teampos As Integer
'Dim plPos As Integer
'Dim wPos As Integer
'Dim vPos As Integer
'Dim gPos As Integer
'Dim pntpos As Integer
'Dim voorPos As Integer
'Dim tegenPos As Integer
'Dim pos As Integer 'de positie van het team in de groep
'
'Dim ypos As Integer
'
'    col(0) = 0
'    col(1) = printObj.ScaleWidth / 4
'    col(2) = printObj.ScaleWidth / 2
'    col(3) = printObj.ScaleWidth / 4 * 3
'    col(4) = printObj.ScaleWidth
'    printObj.fontBold = True
'    fontSizing 12
'    printObj.ForeColor = vbBlue
'    printObj.Print "Groepstanden"
'    ypos = printObj.CurrentY
'    printObj.fontBold = False
'    printObj.ForeColor = 1
'    fontSizing 8
'    teampos = 10
'    plPos = teampos + printObj.TextWidth("1234567890123")
'    wPos = plPos + printObj.TextWidth("000")
'    vPos = wPos + printObj.TextWidth("000")
'    gPos = vPos + printObj.TextWidth("000")
'    pntpos = gPos + printObj.TextWidth("000")
'    voorPos = pntpos + printObj.TextWidth("000")
'    tegenPos = voorPos + printObj.TextWidth("000")
'
'
'    grps = getTournamentInfo("groepen")
'    colNu = 0
'    For i = 1 To grps
'        printObj.CurrentY = ypos
'        sqlstr = "Select * from qryGroepTeams"
'        sqlstr = sqlstr & " Where ksID = " & kampID
'        sqlstr = sqlstr & " AND groep = '" & Chr(i + 64) & "'"
'        sqlstr = sqlstr & " order by pnt DESC, gesp, positie, plaatsing"
'        rsGrp.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'        printObj.CurrentX = col(colNu) + teampos
'        printObj.Print "groep " & Chr(i + 64);
'        printObj.CurrentX = col(colNu) + plPos
'        printObj.Print "sp";
'        printObj.CurrentX = col(colNu) + wPos
'        printObj.Print "W";
'        printObj.CurrentX = col(colNu) + vPos
'        printObj.Print "V";
'        printObj.CurrentX = col(colNu) + gPos
'        printObj.Print "G";
'        printObj.CurrentX = col(colNu) + pntpos
'        printObj.Print "P";
'        printObj.CurrentX = col(colNu) + voorPos
'        printObj.Print "v-t"
'        Do While Not rsGrp.EOF
'            pos = pos + 1
'            printObj.CurrentX = col(colNu) + teampos
'            If rsGrp!positie <> 0 Then
'                printObj.Print Format(rsGrp!positie, "0"); ". "; rsGrp!naam;
'            Else
'                printObj.Print Format(pos, "0"); ". "; rsGrp!naam;
'            End If
'            printObj.CurrentX = col(colNu) + plPos
'            printObj.Print Format(rsGrp!gesp, "0");
'            printObj.CurrentX = col(colNu) + wPos
'            printObj.Print Format(rsGrp!gew, "0");
'            printObj.CurrentX = col(colNu) + vPos
'            printObj.Print Format(rsGrp!verl, "0");
'            printObj.CurrentX = col(colNu) + gPos
'            printObj.Print Format(rsGrp!gel, "0");
'            printObj.CurrentX = col(colNu) + pntpos
'            printObj.Print Format(rsGrp!pnt, "0");
'            printObj.CurrentX = col(colNu) + voorPos
'            printObj.Print Format(rsGrp!voor, "0"); "-"; Format(rsGrp!tegen, "0")
'            rsGrp.MoveNext
'        Loop
'        printObj.Line (col(colNu), ypos)-(col(colNu + 1) - 50, printObj.CurrentY), , B
'        colNu = colNu + 1
'        If colNu > 3 Then
'            colNu = 0
'            ypos = printObj.CurrentY + 50
'        End If
'        pos = 0
'        rsGrp.Close
'    Next
'    printObj.Print
'End Sub
'
'Sub tnWeds()
'Dim sqlstr As String
'Dim rs As New ADODB.Recordset
'Dim rsUitsl As New ADODB.Recordset
'Dim i As Integer
'Dim grps As Integer
'Dim col(3) As Integer
'Dim colNu As Integer
'Dim numpos As Integer
'Dim datPos As Integer
'Dim wedPos As Integer
'Dim uitslPos As Integer
'Dim newYpos As Integer
'Dim ypos As Integer
'    col(0) = 0
'    col(1) = printObj.ScaleWidth / 3
'    col(2) = printObj.ScaleWidth / 3 * 2
'    col(3) = printObj.ScaleWidth
'    sqlstr = "Select * from qryWeds "
'    sqlstr = sqlstr & " WHERE ksid = " & kampID
'    sqlstr = sqlstr & " AND wedtype = 1"
'    sqlstr = sqlstr & " order by mynum"
'    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    rs.MoveLast
'    rs.MoveFirst
'    printObj.fontBold = True
'    fontSizing 12
'    printObj.ForeColor = vbBlue
'    printObj.Print "Groepswedstrijden"
'    ypos = printObj.CurrentY
'    printObj.fontBold = False
'    fontSizing 8
'    printObj.ForeColor = 1
'    numpos = printObj.TextWidth("000")
'    datPos = numpos + printObj.TextWidth("0")
'    wedPos = datPos + printObj.TextWidth("za 29 jun 20uW")
'    uitslPos = col(1) - printObj.TextWidth("0-0 (0-0)")
'    Do While Not rs.EOF
'        i = i + 1
'        printObj.CurrentX = col(colNu) + numpos - printObj.TextWidth(Format(rs!mynum, "0"))
'        printObj.Print Format(rs!mynum, "0");
'        printObj.CurrentX = col(colNu) + datPos
''        printObj.Print Format(rs!Datum, "ddd d mmm"); Format(rs!tijd, " HHu."); ": ";
'
'        printObj.Print Format(rs!datum, "ddd d mmm"); tijdFormat(rs!tijd, True); ": ";
'        printObj.CurrentX = col(colNu) + wedPos
'        printObj.Print rs!naam1 & " - " & rs!naam2;
'        printObj.CurrentX = col(colNu) + uitslPos
'        If WedGespeeld(rs!wedNum) Then
'            printObj.Print GetWedUitsl(rs!wedNum)
'        Else
'            printObj.Print
'        End If
'        rs.MoveNext
'        If i = rs.RecordCount / 3 Then
'            If newYpos < printObj.CurrentY Then
'                newYpos = printObj.CurrentY
'            End If
'            i = 0
'            printObj.CurrentY = ypos
'            colNu = colNu + 1
'        End If
'    Loop
'    printObj.Line (10, ypos)-(printObj.ScaleWidth - 50, newYpos), , B
'    printObj.Line (col(1), ypos)-(col(1), newYpos)
'    printObj.Line (col(2), ypos)-(col(2), newYpos)
'    printObj.Print
'End Sub
'
'Private Sub DoNewPage(pagnr As Boolean, Optional vulKop As Boolean, Optional koppos As Integer)
'    If TypeOf printObj Is Printer Then
'        Printer.NewPage
'    Else
'        Load printPrev.afdrpic(printPrev.afdrpic.UBound + 1)
'        printPrev.afdrpic(printPrev.afdrpic.UBound).Visible = False
'        printPrev.afdrpic(printPrev.afdrpic.UBound).AutoRedraw = True
'        Set printObj = printPrev.afdrpic(printPrev.afdrpic.UBound)
'        printPrev.brnNext.Enabled = printPrev.afdrpic.UBound > 0
'    End If
'    InitPage pagnr, vulKop, koppos, True
'End Sub
'
Private Sub fontSizing(grootte%)

    Printer.fontSize = grootte%
    With printobj.Font
        .Size = Printer.fontSize '* printRatio

    End With
End Sub

Sub initializeVars()

  headerText = getOrganisation(cn)
  currentMatch = getLastMatchPlayed(cn)
  toMatch = currentMatch
  headingFont = "Cambria"
  textFont = "Calibri"

End Sub

Private Sub Form_Load()
Dim i As Integer
Dim prntr As Printer

Dim sqlstr As String
  Set cn = New ADODB.Connection
  With cn
    .ConnectionString = lclConn
    .CursorLocation = adUseClient
    .Open
  End With

  Set rotater = New rotator

  Set printPrev = New printPreview
  sqlstr = "Select nickName from tblCompetitorPools where poolid=" & thisPool
  sqlstr = sqlstr & " order by nickName"
  Set rs = New ADODB.Recordset
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
  initializeVars
  updateForm
  centerForm Me
  UnifyForm Me
End Sub

Sub updateForm()

  Me.picCompetitorList.Top = 90
  Me.picCompetitorList.Left = 3090
  Me.picPrnterSettings.Left = 3090
  Me.picPrnterSettings.Top = 2280

  'button to print everything for the next day
  Me.btnPrntAllAfterDay.Enabled = getAllMatchesPlayedOnDay(Date, cn)
  'nutton to print the results for each participant at end of tournament
  Me.btnFinalPlayerPrint.Enabled = getLastMatchPlayed(cn) = getMatchCount(cn)
  'option to print everything at the end of the tournament
  Me.chkEindstand.Enabled = Me.btnFinalPlayerPrint.Enabled

  If getLastMatchPlayed(cn) >= 1 Then
      Me.txtToMatch.Enabled = True
  End If
  ' Me.chkDblSide.Enabled = printersettings
  Me.upDnToMatch.Max = getCount("Select tournamentID from tblTournamentSchedule where tournamentID = " & thisTournament, cn)
  Me.upDnForMatch.Max = Me.upDnToMatch.Max + 1
  Me.upDnCopies = 1
  Me.optPortrait = True
  Me.optPrintDoc(7).Enabled = getCount("Select competitorPoolID from tblCompetitorPools where poolID = " & thisPool, cn) > 0
  Me.optPrintDoc(1).Enabled = Me.optPrintDoc(7).Enabled
  Me.optPrintDoc(3).Enabled = Me.optPrintDoc(7).Enabled
  Me.optPrintDoc(2).Enabled = currentMatch > 0
  Me.optPrintDoc(4).Enabled = currentMatch > 0
  Me.optPrintDoc(5).Enabled = currentMatch > 0
  Me.optPrintDoc(6).Enabled = currentMatch > 0
  Me.optPrintDoc(8).Enabled = currentMatch > 0
  Me.optPrintDoc(0).value = True
  optPrintDoc_Click 0
  Screen.MousePointer = Default
  ' Me.chkDblSide.Visible = true
  Me.Width = 6630
  Me.Height = 5250

End Sub

Function RandomColor() As Long
    RandomColor = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
End Function


'Private Sub printSkyline()
'Dim rsPnt As New ADODB.Recordset
'Dim rsDeeln As New ADODB.Recordset
'Dim rsEtaps As New ADODB.Recordset
'Dim sqlstr As String
'Dim pnt As Integer
'Dim i As Integer
'Dim J As Integer
'Dim K As Integer
'Dim l As Integer
'Dim xpos As Integer
'Dim ypos As Integer
'Dim yBot As Integer
'Dim tmpX As Integer
'Dim tmpY As Integer
'Dim oldYpos As Integer
'Dim bottom As Integer
'Dim HoogsteNu As Integer
'Dim langsteNaam As Integer
'Dim wedAant As Integer
'Dim deelnAant As Integer
'Dim deelnemsOpPag As Integer
'Dim pagaant As Integer
'Dim deelnpos As Integer
'Dim aantpnt As Integer
'Dim maximum As Integer
'Dim scorepos As Integer
'Dim maximaal As Integer
'Dim schaal As Double
'Dim factor As Integer
'Dim curPag As Integer
'Dim deelnemsPagEen As Integer
'Dim deelnemPagEenPos As Integer
'
'MakeColors
'
'heading1 = "Grafiek t/m wedstrijd " & toMatch
'If Me.Eindstand <> 0 Then
'    heading1 = "Grafiek Eindstand"
'End If
'InitPage False, False
'fontSizing 8
'xpos = printObj.CurrentX + printObj.TextWidth("200") + printObj.ScaleLeft
'ypos = printObj.CurrentY
'sqlstr = "Select deelnemid, bijnaam from pooldeelnems"
'sqlstr = sqlstr & " WHERE poolid =  " & thisPool
'sqlstr = sqlstr & " Order BY bijnaam"
'rsDeeln.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'rsDeeln.MoveLast
'rsDeeln.MoveFirst
'langsteNaam = printObj.TextWidth(Left(GetLangsteBijNaam, 15))
'langsteNaam = langsteNaam + printObj.TextWidth("0(99)")
'bottom = voethoog - langsteNaam
'yBot = voethoog - TextHeight("999")
'deelnAant = rsDeeln.RecordCount
'If Me.optLandscape Then 'landscape
'    deelnemsOpPag = 40
'Else
'    deelnemsOpPag = 26
'End If
'pagaant = 1
'If deelnAant > deelnemsOpPag Then
'    pagaant = (deelnAant / (deelnemsOpPag + 3) + 0.5)
'End If
'
'deelnemsOpPag = Int((deelnAant + 3) / pagaant + 0.5)
'wedAant = GetWedAant(kampID)
'HoogsteNu = getHoogPnt(toMatch)
'If HoogsteNu > 250 Then
'    factor = 50
'ElseIf HoogsteNu > 150 Then
'    factor = 25
'ElseIf HoogsteNu > 100 Then
'    factor = 10
'Else
'    factor = 5
'End If
'Do While aantpnt <= HoogsteNu / factor
'    aantpnt = aantpnt + factor
'Loop
''printObj.Scale
'maximum = Int(HoogsteNu / aantpnt + 1) * aantpnt
'aantpnt = maximum / factor
'scorepos = Int((bottom - ypos) / aantpnt)
''legenda
'printObj.FillStyle = vbSolid
'oldYpos = bottom
'fontSizing 6
'deelnemPagEenPos = printObj.TextWidth("99: XXX-XXXX") + 20
'printObj.ForeColor = vbBlack
'For i = 0 To toMatch - 1
'    printObj.FillColor = kleur(i)
'    printObj.Line (xpos, oldYpos)-(xpos + deelnemPagEenPos - 20, oldYpos - printObj.TextHeight("W")), , B
'    printObj.CurrentX = xpos + 40
'    SetForeCol kleur(i)
'    printObj.Print getWedTeams(i + 1)
'    oldYpos = oldYpos - printObj.TextHeight("W")
'    printObj.ForeColor = vbBlack
'Next
'fontSizing 8
'
'printObj.Line (xpos + deelnemPagEenPos + 40, ypos)-(printObj.ScaleWidth + 2 * printObj.ScaleLeft, ypos)
'printObj.Line -(printObj.ScaleWidth + 2 * printObj.ScaleLeft, bottom)
'printObj.Line -(xpos + deelnemPagEenPos + 40, bottom)
'printObj.Line -(xpos + deelnemPagEenPos + 40, ypos)
'For i = 0 To aantpnt
'    ypos = bottom - i * scorepos
'    fontSizing 8
'    printObj.Line (xpos + deelnemPagEenPos + 40, ypos)-(printObj.ScaleWidth + 2 * printObj.ScaleLeft, ypos)
'    printObj.CurrentX = xpos + deelnemPagEenPos + 40 - TextWidth(CStr(i * maximum / aantpnt)) - 20
'    printObj.CurrentY = ypos - TextHeight("99") / 2
'    printObj.Print i * maximum / aantpnt
'Next
'maximaal = (i - 1) * aantpnt
'schaal = (bottom - ypos) / maximum
''fontSizing 4
'printObj.fontBold = False
'rsDeeln.MoveFirst
''kleur(0) = 15
'curPag = 1
'deelnpos = Int((printObj.ScaleWidth - (2 * printObj.ScaleLeft) - xpos - deelnemPagEenPos) / deelnemsOpPag)
'i = 2 'horizontale positie eerste deelnemer
'deelnemsPagEen = deelnemsOpPag - i
'Do While Not rsDeeln.EOF
'    i = i + 1
'    oldYpos = bottom
''    If curPag > 1 Then deelnemsPagEen = deelnemsOpPag
'    For J = 0 To toMatch - 1
'        printObj.FillColor = kleur(J)
'        pnt = Int(getDeelnPnt(GetWedNum(J + 1), rsDeeln!deelnemID, 1) * (schaal) + 0.5)
'        printObj.Line (xpos + 10 + deelnpos * (i - 1), oldYpos)-(xpos + deelnpos * (i - 1) + deelnpos - 10, oldYpos - pnt), , B
'
'        oldYpos = oldYpos - pnt
'    Next
'    fontSizing 8
'    printObj.CurrentX = xpos + deelnpos * (i - 1) + (deelnpos - printObj.TextWidth(Format(pnt, "999"))) / 2
'    printObj.CurrentY = oldYpos - printObj.TextHeight(Format(pnt, "##"))
'
'    printObj.Print Int(getDeelnPnt(GetWedNum(J), rsDeeln!deelnemID, 0))
'    printObj.CurrentX = xpos + deelnpos * (i - 1) + (deelnpos - TextWidth("W")) / 2
'    tmpX = printObj.CurrentX
'
'    printObj.CurrentY = bottom + printObj.TextWidth(Trim(rsDeeln!bijnaam) & " ")
'    tmpY = printObj.CurrentY
'    printObj.fontBold = False
'    fontSizing 10
'    Set rot.Device = printObj
'    printObj.CurrentY = bottom + 50
'    printObj.CurrentX = xpos + deelnpos * (i - 1) + (deelnpos + printObj.TextWidth("W")) / 2
'    rot.Angle = 270
'    rot.PrintText rsDeeln!bijnaam & " (" & getDeelnPnt(toMatch, rsDeeln!deelnemID, 8) & ")"
'    rsDeeln.MoveNext
'    printObj.DrawWidth = 1
'    If i = deelnemsOpPag And Not rsDeeln.EOF Then
'        DoNewPage False, False
'        curPag = curPag + 1
'        printObj.Line (xpos, ypos)-(printObj.ScaleWidth + 2 * printObj.ScaleLeft, ypos)
'        printObj.Line -(printObj.ScaleWidth + 2 * printObj.ScaleLeft, bottom)
'        printObj.Line -(xpos, bottom)
'        printObj.Line -(xpos, ypos)
'
'        For i = 0 To aantpnt
'            ypos = bottom - i * scorepos
'            fontSizing 8
'            printObj.Line (xpos, ypos)-(printObj.ScaleWidth + 2 * printObj.ScaleLeft, ypos)
'            printObj.CurrentX = xpos - TextWidth(CStr(i * maximum / aantpnt)) - 10
'            printObj.CurrentY = ypos - TextHeight("99") / 2
'            printObj.Print i * maximum / aantpnt
'        Next
'        i = 0
'        printObj.fontBold = False
'        printObj.FillStyle = vbSolid
'    End If
'Loop
'
'End Sub
'
Private Sub Init()
    With Printer
        .FontUnderline = 0
        .fontSize = 18
        GrootHoog = .TextHeight("Jota")
        .fontSize = 10
        KleinHoog = .TextHeight("Jota")
        .fontSize = 8
        SmallHoog = .TextHeight("Jota")
        .fontSize = 12
        NormHoog = .TextHeight("Jota")
        .DrawWidth = 2
    End With
End Sub
'
Private Sub InitPage(pagnr As Boolean, Optional vullen As Boolean, Optional koppos As Integer, Optional vervolg As Boolean)
' boolean 'doorloop' bepaalt of er een voetregel moet komen
    Me.printerDialog.FontName = textFont

    'If Not vervolg Or (vervolg And Me.chkNwePagKop) Then voetregel
    'print the page header
    pageHeader
    ' print the first heading
    headingText heading1, pagnr, vullen, , koppos

End Sub
'

'Private Sub Form_Unload(Cancel As Integer)
'  On Error Resume Next
'    Set printObj = Nothing
'  On Error GoTo 0
'End Sub
'
Private Sub btnClose_Click()
On Error Resume Next

On Error GoTo 0
Printer.KillDoc
Unload printPrev
Unload Me
End Sub
'
Private Sub pageHeader()
Dim printWidth As Integer
Dim fnt As String
  With printobj
    .ForeColor = RGB(0, 51, 0)
    fnt = .FontName
    .FontName = headingFont
    printWidth = .DrawWidth
    .DrawWidth = 1
    printobj.Line (0, 0)-(.ScaleWidth, 0), RGB(0, 51, 0)
    fontSizing 4
    printobj.Print
    fontSizing 16
    printobj.FontBold = True
    Centreer CStr(headerText)
    printobj.Print
    y% = .CurrentY
    printobj.Line (0, y%)-(.ScaleWidth, y%), RGB(0, 51, 0)
    fontSizing 1
    printobj.FontBold = False
    printobj.Print
    kophoog = .CurrentY
    .DrawWidth = printWidth
    .ForeColor = vbBlack
    .FontName = fnt
  End With
End Sub

Private Sub headingText(Tekst$, pagnr As Boolean, Optional vul As Boolean, Optional ypos As Integer, Optional xpos As Integer)
    fontSizing 16

    printobj.FillColor = RGB(0, 51, 0)
    If vul Then
        printobj.FillStyle = vbFSSolid
        printobj.ForeColor = RGB(204, 251, 153)
        printobj.Line (0, kophoog)-(printobj.ScaleWidth - 20, kophoog + printobj.TextHeight("W")), vbBlack, B
    Else
        printobj.ForeColor = RGB(0, 51, 0)
        printobj.FillStyle = vbFSTransparent
    End If
    printobj.FontItalic = True
    printobj.FontBold = True
    printobj.CurrentY = kophoog
    If ypos > 0 Then printobj.CurrentY = ypos

    iBKMode = SetBkMode(printobj.hdc, TRANSPARENT)
    Select Case xpos
    Case 0
        Centreer Tekst$
    Case 1
        printobj.CurrentX = 0
        printobj.Print Tekst$;
    Case 2
        printobj.CurrentX = Int(printobj.ScaleWidth / 4) - printobj.TextWidth(Tekst$) / 2
        printobj.Print Tekst$;
    Case 3
        printobj.CurrentX = Int(printobj.ScaleWidth / 2) - printobj.TextWidth(Tekst$) / 2
        printobj.Print Tekst$;
    Case 4
        printobj.CurrentX = Int(printobj.ScaleWidth / 4) * 3 - printobj.TextWidth(Tekst$) / 2
        printobj.Print Tekst$;
    End Select
    favYpos = printobj.CurrentY
    fontSizing 9
    printobj.CurrentY = printobj.CurrentY + GrootHoog - KleinHoog
    printobj.CurrentX = printobj.ScaleWidth - printobj.TextWidth("blad 12")
    If TypeOf printobj Is Printer Then
        If printobj.Page > 1 And pagnr Then
            printobj.Print "blad "; printobj.Page;
        End If
    Else
        If printobj.Index > 0 And pagnr Then
            printobj.Print "blad "; printobj.Index + 1;
        End If
    End If
    fontSizing 12
    printobj.Print
    kophoog = printobj.CurrentY
    printobj.FillStyle = vbFSTransparent
    printobj.ForeColor = vbBlack
    printobj.FontItalic = False
    printobj.FontBold = False
End Sub

'Function getAant(deeln As Long, vanwat As String)
''haal het aantal scores op van 'vanwat' bij deeln
'Dim rsdeelnScore As New ADODB.Recordset
'Dim sqlstr As String
'    sqlstr = "SELECT * from deelnempnt"
'    sqlstr = sqlstr & " Where deelnID =" & deeln
'    sqlstr = sqlstr & " AND " & vanwat & " > 0"
'    rsdeelnScore.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    If rsdeelnScore.RecordCount > 0 Then
'        rsdeelnScore.MoveLast
'    End If
'    getAant = rsdeelnScore.RecordCount
'
'End Function
'
'Function GetPntDeelnem(deeln As Long, vanwat As String)
'Dim rsdeelnScore As New ADODB.Recordset
'Dim pnt As Integer
'Dim sqlstr As String
'    sqlstr = "SELECT * from deelnempnt"
'    sqlstr = sqlstr & " Where deelnID =" & deeln
'    sqlstr = sqlstr & " AND " & vanwat & " > 0"
'    sqlstr = sqlstr & " order by wednum"
'    rsdeelnScore.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    If rsdeelnScore.RecordCount > 0 Then
'        rsdeelnScore.MoveLast
'        GetPntDeelnem = rsdeelnScore(vanwat)
'        If UCase(Left(vanwat, 7)) = UCase("pntfin4") Then
'            rsdeelnScore.MoveFirst
'            pnt = 0
'            Do While Not rsdeelnScore.EOF
'                pnt = pnt + rsdeelnScore(vanwat)
'                rsdeelnScore.MoveNext
'            Loop
'            GetPntDeelnem = pnt
'        ElseIf UCase(Left(vanwat, 7)) = UCase("pntfin2") Then
'            rsdeelnScore.MoveFirst
'            pnt = 0
'            Do While Not rsdeelnScore.EOF
'                pnt = pnt + rsdeelnScore(vanwat)
'                rsdeelnScore.MoveNext
'            Loop
'            GetPntDeelnem = pnt
'        End If
'    Else
'        GetPntDeelnem = 0
'    End If
'End Function
'
'Sub printPoolPoints(alfabet As Boolean)
'Dim rsDeeln As New ADODB.Recordset
'Dim rsdeelnScore As New ADODB.Recordset
'Dim sqlstr As String
'Dim bedr As Currency
'Dim geldold As Currency
'Dim savy As Integer
'Dim leftmarge As Integer
'Dim pntpos() As Integer
'Dim pnt As Integer
'Dim aant As Integer
'Dim grpPnt As Integer
'Dim geld As Double
'Dim geldttl As Double
'Dim Tekst$
'Dim prStr As String
'Dim topYpos As Integer
'Dim top2Ypos As Integer
'Dim botY As Integer
'Dim lastDeelnPos As Integer
'Dim maxY As Integer
'Dim grp As String
'Dim i As Integer
'Dim J As Integer
'Dim ipos As Integer
'Dim has8eFin As Boolean
'Dim hasKlFin As Boolean
'Dim grpAant As Integer
'Dim wdNum As Integer
'Dim prTtl As Boolean
'Dim colbr As Integer
'Dim grpStndBegin As Integer '6
'Dim fin8Begin As Integer    '15
'Dim fin4Begin As Integer    '24
'Dim fin2Begin As Integer    '29
'Dim finBegin As Integer     '32
'
'Dim EindstBegin As Integer  '34
'Dim AantBegin As Integer    '38
'Dim TopScBegin As Integer   '43
'Dim TTLBegin As Integer     '44
'Dim PosBegin As Integer     '45
'Dim GeldBegin As Integer    '46
'
'Dim tmp$
'Dim yposnu%
'
'    grpAant = getTournamentInfo("groepen")
'    If grpAant > 4 Then
'        colbr = 140
'    Else
'        colbr = 250
'    End If
'    has8eFin = grpAant > 4
'    hasKlFin = getTournamentInfo("derdeplaats")
'    If GetLastPlayed = getlastWednum Then
'        pntFormat = "0"
'    Else
'        pntFormat = "0;;\ ;-"
'    End If
'
'    leftmarge = printObj.CurrentX
'    fontSizing 10
'    printObj.Print
'
'    fontSizing 16
'    printObj.fontBold = True
'    If Me.Eindstand = False Then
'        If alfabet Then
'            Tekst$ = "Puntenopbouw t/m wedstrijd " & GetMyNum(GetLastPlayed)
'        Else
'            Tekst$ = "Puntenopbouw t/m wedstrijd (hoog-laag)" & GetMyNum(GetLastPlayed)
'        End If
'    Else
'        If alfabet Then
'            Tekst$ = "Eindstand (alfabetisch)"
'        Else
'            Tekst$ = "Eindstand (op score)"
'        End If
'    End If
'    headerText = GetOrgNaam(thisPool) & " " & getTournamentInfo("toernooi") & " voetbalpool"
'
'    heading1 = Tekst$
'
'
'
'    InitPage False, True
'    printobj.FontItalic = False
'    printObj.fontBold = False
'    fontSizing 8
'    topYpos = printObj.CurrentY
'    printObj.Line (0, topYpos)-(printObj.ScaleWidth - 50, topYpos)
'    printObj.CurrentX = leftmarge
'    sqlstr = DeelnResultSql(False, GetLastPlayed)
'    rsDeeln.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    If rsDeeln.RecordCount > 0 Then
'        rsDeeln.MoveLast
'        lastDeelnPos = rsDeeln!postotaal
'    End If
'    rsDeeln.Close
'    sqlstr = DeelnResultSql(alfabet, GetLastPlayed)
'
'    rsDeeln.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    If rsDeeln.RecordCount = 0 Then
'        printObj.Print "Geen deelnemers gevonden"
'        Exit Sub
'    End If
'    fontSizing 10
'    printObj.CurrentX = leftmarge
'    printObj.Print "Naam";
'    printObj.CurrentX = printObj.TextWidth("123456789012345")
'    ReDim Preserve pntpos(1)
'    pntpos(0) = 0
'    pntpos(1) = printObj.CurrentX - colbr
'    printObj.Print
'    top2Ypos = printObj.CurrentY
'    printObj.CurrentX = pntpos(1) + colbr
'    fontSizing 8
'    printObj.Print "rust"; '("; Format(getPnt(1), pntFormat); "p)";
'    ReDim Preserve pntpos(UBound(pntpos) + 1)
'    pntpos(UBound(pntpos)) = printObj.CurrentX
'    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'    printObj.Print "eind"; '("; Format(getPnt(2), pntFormat); "p)";
'    ReDim Preserve pntpos(UBound(pntpos) + 1)
'    pntpos(UBound(pntpos)) = printObj.CurrentX
'    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'    printObj.Print "toto"; '("; Format(getPnt(3), pntFormat); "p)";
'    ReDim Preserve pntpos(UBound(pntpos) + 1)
'    pntpos(UBound(pntpos)) = printObj.CurrentX
'    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'    printObj.Print "dlp"; '("; Format(getPnt(28), pntFormat); "p)";
'    ReDim Preserve pntpos(UBound(pntpos) + 1)
'    pntpos(UBound(pntpos)) = printObj.CurrentX
'    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'    printObj.Print "tot";
'    ReDim Preserve pntpos(UBound(pntpos) + 1)
'    pntpos(UBound(pntpos)) = printObj.CurrentX
'    grpStndBegin = UBound(pntpos)
'
'    For i = 1 To grpAant
'        printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'        printObj.Print Chr(i + 64);
'        ReDim Preserve pntpos(UBound(pntpos) + 1)
'        pntpos(UBound(pntpos)) = printObj.CurrentX
'    Next
'    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'    printObj.Print "tot";
'    ReDim Preserve pntpos(UBound(pntpos) + 1)
'    pntpos(UBound(pntpos)) = printObj.CurrentX
'    If grpAant > 4 Then
'        fin8Begin = UBound(pntpos)
'        For i = 1 To grpAant
'            printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'            printObj.Print Chr(i + 64);
'            ReDim Preserve pntpos(UBound(pntpos) + 1)
'            pntpos(UBound(pntpos)) = printObj.CurrentX
'        Next
'        printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'        printObj.Print "tot";
'        ReDim Preserve pntpos(UBound(pntpos) + 1)
'        pntpos(UBound(pntpos)) = printObj.CurrentX
'    End If
'    fin4Begin = UBound(pntpos)
'    For i = 1 To 4
'        printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'        printObj.Print Format(i, "0");
'        ReDim Preserve pntpos(UBound(pntpos) + 1)
'        pntpos(UBound(pntpos)) = printObj.CurrentX
'    Next
'    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'    printObj.Print "tot";
'    ReDim Preserve pntpos(UBound(pntpos) + 1)
'    pntpos(UBound(pntpos)) = printObj.CurrentX
'    fin2Begin = UBound(pntpos)
'    For i = 1 To 2
'        printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'        printObj.Print "  "; Format(i, "0"); "e  ";
'        ReDim Preserve pntpos(UBound(pntpos) + 1)
'        pntpos(UBound(pntpos)) = printObj.CurrentX
'    Next
'    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'    printObj.Print "tot";
'    ReDim Preserve pntpos(UBound(pntpos) + 1)
'    pntpos(UBound(pntpos)) = printObj.CurrentX
'    finBegin = UBound(pntpos)
'    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'    If hasKlFin Then
'        printObj.Print "kl("; Format(getPnt(30), pntFormat);
'        If getPnt(31) > 0 Then
'            printObj.Print "/"; Format(getPnt(31), pntFormat);
'        End If
'        printObj.Print ")";
'        ReDim Preserve pntpos(UBound(pntpos) + 1)
'        pntpos(UBound(pntpos)) = printObj.CurrentX
'        printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'        printObj.Print "gr("; Format(getPnt(11), pntFormat);
'        If getPnt(12) > 0 Then
'            printObj.Print "/"; Format(getPnt(12), pntFormat);
'        End If
'        printObj.Print ")";
'        ReDim Preserve pntpos(UBound(pntpos) + 1)
'        pntpos(UBound(pntpos)) = printObj.CurrentX
'        printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'    Else
'        printObj.Print "("; Format(getPnt(11), pntFormat);
'        If getPnt(12) > 0 Then
'            printObj.Print "/"; Format(getPnt(12), pntFormat);
'        End If
'        printObj.Print ")";
'        ReDim Preserve pntpos(UBound(pntpos) + 1)
'        pntpos(UBound(pntpos)) = printObj.CurrentX
'        printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'    End If
'    EindstBegin = UBound(pntpos)
'    ' Format(getPnt(15), pntFormat); "/"; Format(getPnt(14), pntFormat); "/"; Format(getPnt(13), pntFormat); "/"; Format(getPnt(29), pntFormat); ")";
'    printObj.Print "1";
'    ReDim Preserve pntpos(UBound(pntpos) + 1)
'    pntpos(UBound(pntpos)) = printObj.CurrentX
'    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'    printObj.Print "2";
'    ReDim Preserve pntpos(UBound(pntpos) + 1)
'    pntpos(UBound(pntpos)) = printObj.CurrentX
'    If hasKlFin Then
'        printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'        printObj.Print "3";
'        ReDim Preserve pntpos(UBound(pntpos) + 1)
'        pntpos(UBound(pntpos)) = printObj.CurrentX
'        printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'        printObj.Print "4";
'        ReDim Preserve pntpos(UBound(pntpos) + 1)
'        pntpos(UBound(pntpos)) = printObj.CurrentX
'    End If
'    AantBegin = UBound(pntpos)
'    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'    printObj.Print "dp";
'    ReDim Preserve pntpos(UBound(pntpos) + 1)
'    pntpos(UBound(pntpos)) = printObj.CurrentX
'    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'    printObj.Print "gel";
'    ReDim Preserve pntpos(UBound(pntpos) + 1)
'    pntpos(UBound(pntpos)) = printObj.CurrentX
'    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'    printObj.Print "gl";
'    ReDim Preserve pntpos(UBound(pntpos) + 1)
'    pntpos(UBound(pntpos)) = printObj.CurrentX
'    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'    printObj.Print "rd";
'    ReDim Preserve pntpos(UBound(pntpos) + 1)
'    pntpos(UBound(pntpos)) = printObj.CurrentX
'    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'    printObj.Print "pn";
'    ReDim Preserve pntpos(UBound(pntpos) + 1)
'    pntpos(UBound(pntpos)) = printObj.CurrentX
'    TopScBegin = UBound(pntpos)
'    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'    printObj.Print "scor";
'    ReDim Preserve pntpos(UBound(pntpos) + 1)
'    pntpos(UBound(pntpos)) = printObj.CurrentX
'    TTLBegin = UBound(pntpos)
'    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr + printObj.TextWidth("123")
'    ReDim Preserve pntpos(UBound(pntpos) + 1)
'    pntpos(UBound(pntpos)) = printObj.CurrentX
'    PosBegin = UBound(pntpos)
'    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr + printObj.TextWidth("123")
'    ReDim Preserve pntpos(UBound(pntpos) + 1)
'    pntpos(UBound(pntpos)) = printObj.CurrentX
'    GeldBegin = UBound(pntpos)
'    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'    printObj.Print "";
'    'laatste kolom
'    ReDim Preserve pntpos(UBound(pntpos) + 1)
'    pntpos(UBound(pntpos)) = printObj.ScaleWidth - 50
'
'    printObj.CurrentY = topYpos
'    fontSizing 10
'    printObj.CurrentX = (pntpos(1) + pntpos(grpStndBegin) + colbr - printObj.TextWidth("Wedstrijdpunten")) / 2
'    printObj.Print "Wedstrijdpunten";
'    If grpAant > 4 Then
'        printObj.CurrentX = (pntpos(grpStndBegin) + pntpos(fin8Begin) + colbr - printObj.TextWidth("Groepstand (" & Format(getPnt(8), pntFormat) & "p)")) / 2
'    Else
'        printObj.CurrentX = (pntpos(grpStndBegin) + pntpos(fin4Begin) + colbr - printObj.TextWidth("Groepstand (" & Format(getPnt(8), pntFormat) & "p)")) / 2
'    End If
'    printObj.Print "Groepstand (" & Format(getPnt(8), pntFormat) & "p)";
'    If grpAant > 4 Then
'        printObj.CurrentX = (pntpos(fin8Begin) + pntpos(fin4Begin) + colbr - printObj.TextWidth("8e Finalisten (" & Format(getPnt(6), pntFormat) & "/" & Format(getPnt(7), pntFormat) & "p)")) / 2
'        printObj.Print "8e Finalisten (" & Format(getPnt(4), pntFormat);
'        If getPnt(5) > 0 Then
'            printObj.Print "/" & Format(getPnt(5), pntFormat);
'        End If
'        printObj.Print "p)";
'    End If
'    printObj.CurrentX = (pntpos(fin4Begin) + pntpos(fin2Begin) + colbr - printObj.TextWidth("4e fin.(" & Format(getPnt(6), pntFormat) & "/" & Format(getPnt(7), pntFormat) & "p)")) / 2
'    printObj.Print "4efin.(" & Format(getPnt(6), pntFormat);
'    If getPnt(7) > 0 Then
'        printObj.Print "/" & Format(getPnt(7), pntFormat);
'    End If
'    printObj.Print "p)";
'    printObj.CurrentX = (pntpos(fin2Begin) + pntpos(finBegin) + colbr - printObj.TextWidth("2efin.(" & Format(getPnt(9), pntFormat) & "/" & Format(getPnt(10), pntFormat) & "p)")) / 2
'    printObj.Print "1/2fin.(" & Format(getPnt(9), pntFormat);
'    If getPnt(10) > 0 Then
'        printObj.Print "/" & Format(getPnt(10), pntFormat);
'    End If
'    printObj.Print "p)";
'    printObj.CurrentX = (pntpos(finBegin) + pntpos(EindstBegin) + colbr - printObj.TextWidth("Fin")) / 2
'    printObj.Print "Fin";
'    printObj.CurrentX = (pntpos(EindstBegin) + pntpos(AantBegin) + colbr - printObj.TextWidth("Eind")) / 2
'    printObj.Print "Eind";
'    printObj.CurrentX = (pntpos(AantBegin) + pntpos(TopScBegin) + colbr - printObj.TextWidth("Aantallen")) / 2
'    printObj.Print "Aantallen";
'    printObj.CurrentX = pntpos(TopScBegin) + colbr
'    printObj.Print "top";
'    printObj.CurrentX = (pntpos(TTLBegin) + pntpos(PosBegin) + colbr - printObj.TextWidth("Ttl")) / 2
'    printObj.Print "Ttl";
'    printObj.CurrentX = (pntpos(PosBegin) + pntpos(GeldBegin) + colbr - printObj.TextWidth("Pos")) / 2
'    printObj.Print "Pos";
'    printObj.CurrentX = (pntpos(GeldBegin) + pntpos(GeldBegin + 1) + colbr - printObj.TextWidth("Geld")) / 2
'    printObj.Print "Geld";
'    fontSizing 8
'    printObj.CurrentY = top2Ypos
'    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'    printObj.Print
'    printObj.Line (0, printObj.CurrentY)-(printObj.ScaleWidth - 50, printObj.CurrentY)
'    With rsDeeln
'        Do While Not .EOF
''            If rsDeeln!deelnemID = 251 Then Stop
'            printObj.CurrentX = leftmarge
'            If !postotaal = 1 Then
'                printObj.ForeColor = vbBlue
'                printObj.fontBold = True
'            End If
'            If !postotaal = lastDeelnPos Then
'                printObj.ForeColor = vbRed
'            End If
'            printObj.Print !bijnaam;
'            printObj.ForeColor = 1
'            printObj.fontBold = False
'            pnt = PrintAant(!deelnemID, pntpos(2), "pntRust")
'            pnt = pnt + PrintAant(!deelnemID, pntpos(3), "pntEind")
'            pnt = pnt + PrintAant(!deelnemID, pntpos(4), "pntToto")
'            pnt = pnt + PrintAant(!deelnemID, pntpos(5), "dpvddag")
'            printObj.CurrentX = pntpos(6) - printObj.TextWidth(Format(pnt, pntFormat))
'            printObj.fontBold = True
'            printObj.Print Format(pnt, pntFormat);
'            printObj.fontBold = False
'            pnt = 0
'            grpPnt = 0
'            For i = 1 To grpAant
'                If allPlayed(Chr(i + 64)) Then
'                    pntFormat = "0"
'                Else
'                    pntFormat = "0;;\ ;-"
'                End If
'                grpPnt = GetPntDeelnem(!deelnID, "pntgrp" & Chr(i + 64))
'                pnt = pnt + grpPnt
'                printObj.CurrentX = (pntpos(i + 5) + pntpos(i + 6) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
'                printObj.Print Format(grpPnt, pntFormat);
'            Next
'            If grpAant > 4 Then
'                printObj.CurrentX = pntpos(fin8Begin) - printObj.TextWidth(Format(pnt, pntFormat))
'            Else
'                printObj.CurrentX = pntpos(fin4Begin) - printObj.TextWidth(Format(pnt, pntFormat))
'            End If
'            printObj.fontBold = True
'            printObj.Print Format(pnt, pntFormat);
'            printObj.fontBold = False
'            pnt = 0
'            grpPnt = 0
'            If grpAant > 4 Then
'                For i = 1 To grpAant
'                    If allPlayed(Chr(i + 64)) Then
'                        pntFormat = "0"
'                        grpPnt = GetPntDeelnem(!deelnID, "pntfin8" & Chr(i + 64))
'                    Else
'                        grpPnt = 0
'                        pntFormat = "0;;\ ;-"
'                    End If
'                    pnt = pnt + grpPnt
'                    printObj.CurrentX = (pntpos(fin8Begin - 1 + i) + pntpos(i + fin8Begin) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
'                    printObj.Print Format(grpPnt, pntFormat);
'                Next
'                printObj.CurrentX = pntpos(fin4Begin) - printObj.TextWidth(Format(pnt, pntFormat))
'                printObj.fontBold = True
'                If allPlayed("A") Then
'                    pntFormat = "0"
'                Else
'                    pntFormat = "0;;\ ;-"
'                End If
'                printObj.Print Format(pnt, pntFormat);
'                printObj.fontBold = False
'                pnt = 0
'                grpPnt = 0
'            Else
'                For i = 1 To grpAant
'                    If allPlayed(Chr(i + 64)) Then
'                        pntFormat = "0"
'                        grpPnt = GetPntDeelnem(!deelnID, "pntfin4" & Chr(i + 64))
'                    Else
'                        grpPnt = 0
'                        pntFormat = "0;;\ ;-"
'                    End If
'                    pnt = pnt + grpPnt
'                    printObj.CurrentX = (pntpos(fin4Begin - 1 + i) + pntpos(i + fin4Begin) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
'                    printObj.Print Format(grpPnt, pntFormat);
'                Next
'                printObj.fontBold = True
'                If allPlayed("A") Then
'                    pntFormat = "0"
'                Else
'                    pntFormat = "0;;\ ;-"
'                End If
'                printObj.CurrentX = pntpos(fin2Begin) - printObj.TextWidth(Format(pnt, pntFormat))
'                printObj.Print Format(pnt, pntFormat);
'                printObj.fontBold = False
'                pnt = 0
'                grpPnt = 0
'            End If
'
'            If grpAant > 4 Then
'                For i = 1 To 8 Step 2
'                    grpPnt = 0
'                    'If !deelnID = 139 Then Stop
'                    wdNum = i + J + GetFirstFinaleMatch(AchtsteFinale) - 1
'                    Select Case wdNum
'                    Case 49, 50
'                        grp = "B"
'                        ipos = 2
'                    Case 51, 52
'                        grp = "C"
'                        ipos = 3
'                    Case 53, 54
'                        grp = "A"
'                        ipos = 1
'                    Case 55, 56
'                        grp = "D"
'                        ipos = 4
'                    End Select
'                    If GetMyNum(wdNum) <= GetMyNum(GetLastPlayed) Then
'                        pntFormat = "0"
'                        grpPnt = grpPnt + GetPntDeelnem(!deelnID, "pntFin4" & grp)
'                        'grpPnt = grpPnt + getDeelnPnt(GetPrevWednum(wdNum), !deelnID, 9, "4" & grp)
'                        prTtl = True
'                    Else
'                        pntFormat = "0;;\ ;-"
'                        grpPnt = 0
'                    End If
'                    pnt = pnt + grpPnt
'                    printObj.CurrentX = (pntpos(ipos + fin4Begin - 1) + pntpos(ipos + fin4Begin) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
'                    printObj.Print Format(grpPnt, pntFormat);
'                Next
'                If prTtl > 0 Then pntFormat = "0"
'                printObj.CurrentX = pntpos(fin2Begin) - printObj.TextWidth(Format(pnt, pntFormat))
'                printObj.fontBold = True
'                printObj.Print Format(pnt, pntFormat);
'                printObj.fontBold = False
'            End If
'            pnt = 0
'            grpPnt = 0
'            For i = 1 To 2
'                If GetMyNum(i + GetFirstFinaleMatch(KwartFinale) - 1) <= GetMyNum(GetLastPlayed) Then
'                    pntFormat = "0"
'                Else
'                    pntFormat = "0;;\ ;-"
'                End If
'                'If !deelnID = 183 Then Stop
'                grpPnt = GetPntDeelnem(!deelnID, "pntfin2" & Chr(i + 64))
'                pnt = pnt + grpPnt
'                printObj.CurrentX = (pntpos(i + fin2Begin - 1) + pntpos(i + fin2Begin) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
'                printObj.Print Format(grpPnt, pntFormat);
'            Next
'            printObj.CurrentX = pntpos(finBegin) - printObj.TextWidth(Format(pnt, pntFormat))
'            printObj.fontBold = True
'            printObj.Print Format(pnt, pntFormat);
'            printObj.fontBold = False
'            If GetMyNum(GetFirstFinaleMatch(HalveFinale)) <= GetMyNum(GetLastPlayed) Then
'                pntFormat = "0"
'            Else
'                pntFormat = "0;;\ ;-"
'            End If
'            If hasKlFin Then
'                grpPnt = GetPntDeelnem(!deelnID, "pntklfin")
'                printObj.CurrentX = pntpos(32) + (pntpos(33) - pntpos(32) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
'                printObj.Print Format(grpPnt, pntFormat);
'            End If
'            grpPnt = GetPntDeelnem(!deelnID, "pntfin")
'            printObj.CurrentX = (pntpos(finBegin + 1) + pntpos(EindstBegin) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
'            printObj.Print Format(grpPnt, pntFormat);
'            pntFormat = "0;;\ ;-"
'            If GetLastPlayed = getlastWednum Then pntFormat = "0"
'            For i = 1 To 2
'                grpPnt = getEindStandpnt(!deelnID, i)
'                printObj.CurrentX = (pntpos(finBegin + 1 + i) + pntpos(EindstBegin + i) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
'                printObj.Print Format(grpPnt, pntFormat);
'            Next
'            pntFormat = "0;;\ ;-"
'            If GetLastPlayed >= getlastWednum - 1 Then pntFormat = "0"
'            For i = 3 To 4
'                grpPnt = getEindStandpnt(!deelnID, i)
'                printObj.CurrentX = (pntpos(EindstBegin - 1 + i) + pntpos(EindstBegin + i) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
'                printObj.Print Format(grpPnt, pntFormat);
'            Next
'            pntFormat = "0;;\ ;-"
'            If GetLastPlayed = getlastWednum Then
'                pntFormat = "0"
'                grpPnt = getDeelnAantPnt(!deelnID, voorspDP)
'                printObj.CurrentX = (pntpos(AantBegin) + pntpos(AantBegin + 1) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
'                printObj.Print Format(grpPnt, pntFormat);
'                grpPnt = getDeelnAantPnt(!deelnID, voorspGelijk)
'                printObj.CurrentX = (pntpos(AantBegin + 1) + pntpos(AantBegin + 2) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
'                printObj.Print Format(grpPnt, pntFormat);
'                grpPnt = getDeelnAantPnt(!deelnID, voorspGeel)
'                printObj.CurrentX = (pntpos(AantBegin + 2) + pntpos(AantBegin + 3) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
'                printObj.Print Format(grpPnt, pntFormat);
'                grpPnt = getDeelnAantPnt(!deelnID, voorspRood)
'                printObj.CurrentX = (pntpos(AantBegin + 3) + pntpos(AantBegin + 4) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
'                printObj.Print Format(grpPnt, pntFormat);
'                grpPnt = getDeelnAantPnt(!deelnID, voorspPens)
'                printObj.CurrentX = (pntpos(AantBegin + 4) + pntpos(TopScBegin) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
'                printObj.Print Format(grpPnt, pntFormat);
'                grpPnt = GetPntDeelnem(!deelnID, "pntTopSc")
'                printObj.CurrentX = (pntpos(TopScBegin) + pntpos(TTLBegin) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
'                printObj.Print Format(grpPnt, pntFormat);
'            End If
'            pntFormat = "0"
'            If !postotaal = 1 Then
'                printObj.ForeColor = vbBlue
'                printObj.fontBold = True
'            End If
'            If !postotaal = lastDeelnPos Then
'                printObj.ForeColor = vbRed
'            End If
'            'If !deelnID = 125 Then Stop
'            grpPnt = GetPntDeelnem(!deelnID, "grandtotaal")
'            printObj.CurrentX = (pntpos(TTLBegin) + pntpos(PosBegin) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
'            printObj.Print Format(grpPnt, pntFormat);
'            grpPnt = GetPntDeelnem(!deelnID, "postotaal")
'            printObj.CurrentX = (pntpos(PosBegin) + pntpos(GeldBegin) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
'            printObj.Print Format(grpPnt, pntFormat);
'            printObj.ForeColor = 1
'            printObj.fontBold = False
'            geld = GetPntDeelnem(!deelnID, "geldttl")
'            printObj.CurrentX = pntpos(GeldBegin + 1) - colbr - printObj.TextWidth(Format(geld, "currency"))
'            printObj.Print Format(geld, "currency");
'            printObj.Print
'            printObj.ForeColor = 1
'            printObj.Line (0, printObj.CurrentY)-(printObj.ScaleWidth - 50, printObj.CurrentY)
'            grpPnt = 0
'
'            .MoveNext
''            If .AbsolutePosition >= 53 Then Stop
'            If printObj.CurrentY >= voethoog Then 'onderkant pagina
'              If Not rsDeeln.EOF Then
'                botY = printObj.CurrentY
'                printObj.Line (pntpos(1) + 75, topYpos)-(pntpos(1) + 75, top2Ypos)
'                printObj.Line (pntpos(grpStndBegin) + 75, topYpos)-(pntpos(6) + 75, top2Ypos)
'                If grpAant > 4 Then
'                    printObj.Line (pntpos(fin8Begin) + 75, topYpos)-(pntpos(15) + 75, top2Ypos)
'                End If
'                printObj.Line (pntpos(fin4Begin) + 75, topYpos)-(pntpos(fin4Begin) + 75, top2Ypos)
'                printObj.Line (pntpos(fin2Begin) + 75, topYpos)-(pntpos(fin2Begin) + 75, top2Ypos)
'                printObj.Line (pntpos(finBegin) + 75, topYpos)-(pntpos(finBegin) + 75, top2Ypos)
'                printObj.Line (pntpos(EindstBegin) + 75, topYpos)-(pntpos(EindstBegin) + 75, top2Ypos)
'                printObj.Line (pntpos(AantBegin) + 75, topYpos)-(pntpos(AantBegin) + 75, top2Ypos)
'                printObj.Line (pntpos(TopScBegin) + 75, topYpos)-(pntpos(TopScBegin) + 75, top2Ypos)
'                printObj.Line (pntpos(TTLBegin) + 75, topYpos)-(pntpos(TTLBegin) + 75, top2Ypos)
'                printObj.Line (pntpos(PosBegin) + 75, topYpos)-(pntpos(PosBegin) + 75, top2Ypos)
'                printObj.Line (pntpos(GeldBegin) + 75, topYpos)-(pntpos(GeldBegin) + 75, top2Ypos)
'                For i = 1 To UBound(pntpos) - 1
'                    printObj.Line (pntpos(i) + 75, top2Ypos)-(pntpos(i) + 75, botY)
'                Next
'                printObj.Line (printObj.ScaleWidth - 50, topYpos)-(printObj.ScaleWidth - 50, botY)
'                DoNewPage False, True
'                printObj.Line (0, topYpos)-(printObj.ScaleWidth - 50, topYpos)
'                printObj.CurrentX = leftmarge
'                printObj.CurrentY = topYpos
'                fontSizing 10
'                printObj.Print "Naam";
'                printObj.CurrentY = top2Ypos
'                printObj.CurrentX = pntpos(1) + colbr
'                fontSizing 8
'                printObj.Print "rust"; '("; Format(getPnt(1), pntFormat); "p)";
'                printObj.CurrentX = pntpos(2) + colbr
'                printObj.Print "eind"; '("; Format(getPnt(2), pntFormat); "p)";
'                printObj.CurrentX = pntpos(3) + colbr
'                printObj.Print "toto"; '("; Format(getPnt(3), pntFormat); "p)";
'                printObj.CurrentX = pntpos(4) + colbr
'                printObj.Print "dlp"; '("; Format(getPnt(28), pntFormat); "p)";
'                printObj.CurrentX = pntpos(5) + colbr
'                printObj.Print "tot";
'                If grpAant > 4 Then
'                    For i = 1 To 8
'                        printObj.CurrentX = pntpos(5 + i) + colbr
'                        printObj.Print Chr(i + 64);
'                    Next
'                    printObj.CurrentX = pntpos(14) + colbr
'                    printObj.Print "tot";
'                    For i = 1 To 8
'                        printObj.CurrentX = pntpos(14 + i) + colbr
'                        printObj.Print Chr(i + 64);
'                    Next
'                    printObj.CurrentX = pntpos(23) + colbr
'                    printObj.Print "tot";
'                End If
'                For i = 1 To 4
'                    printObj.CurrentX = pntpos(fin4Begin - 1 + i) + colbr
'                    printObj.Print Format(i, "0");
'                Next
'                printObj.CurrentX = pntpos(fin2Begin - 1) + colbr
'                printObj.Print "tot";
'                For i = 1 To 2
'                    printObj.CurrentX = pntpos(fin2Begin - 1 + i) + colbr
'                    printObj.Print "  "; Format(i, "0"); "e  ";
'                Next
'                printObj.CurrentX = pntpos(finBegin - 1) + colbr
'                printObj.Print "tot";
'                printObj.CurrentX = pntpos(finBegin) + colbr
'                If hasKlFin Then
'                    printObj.Print "kl("; Format(getPnt(30), pntFormat);
'                    If getPnt(31) > 0 Then
'                        printObj.Print "/"; Format(getPnt(31), pntFormat);
'                    End If
'                    printObj.Print ")";
'                    printObj.CurrentX = pntpos(EindstBegin - 1) + colbr
'                    printObj.Print "gr("; Format(getPnt(11), pntFormat);
'                    If getPnt(12) > 0 Then
'                        printObj.Print "/"; Format(getPnt(12), pntFormat);
'                    End If
'                    printObj.Print ")";
'                Else
'                    printObj.Print "("; Format(getPnt(11), pntFormat);
'                    If getPnt(12) > 0 Then
'                        printObj.Print "/"; Format(getPnt(12), pntFormat);
'                    End If
'                    printObj.Print ")";
'                End If
'
'                For i = 1 To grpAant / 2
'                    printObj.CurrentX = pntpos(EindstBegin - 1 + i) + colbr
'                    ' Format(getPnt(15), pntFormat); "/"; Format(getPnt(14), pntFormat); "/"; Format(getPnt(13), pntFormat); "/"; Format(getPnt(29), pntFormat); ")";
'                    printObj.Print Format(i, "0");
'                Next
'                printObj.CurrentX = pntpos(AantBegin) + colbr
'                printObj.Print "dp";
'                printObj.CurrentX = pntpos(AantBegin + 1) + colbr
'                printObj.Print "gel";
'                printObj.CurrentX = pntpos(AantBegin + 2) + colbr
'                printObj.Print "gl";
'                printObj.CurrentX = pntpos(AantBegin + 3) + colbr
'                printObj.Print "rd";
'                printObj.CurrentX = pntpos(AantBegin + 4) + colbr
'                printObj.Print "pn";
'                printObj.CurrentX = pntpos(TopScBegin) + colbr
'                printObj.Print "scor";
'                'laatste kolom
'                printObj.CurrentX = printObj.ScaleWidth - 50
'
'                printObj.CurrentY = topYpos
'                fontSizing 10
'                printObj.CurrentX = (pntpos(1) + pntpos(grpStndBegin) + colbr - printObj.TextWidth("Wedstrijdpunten")) / 2
'                printObj.Print "Wedstrijdpunten";
'                If grpAant > 4 Then
'                    printObj.CurrentX = (pntpos(grpStndBegin) + pntpos(fin8Begin) + colbr - printObj.TextWidth("Groepstand (" & Format(getPnt(8), pntFormat) & "p)")) / 2
'                Else
'                    printObj.CurrentX = (pntpos(grpStndBegin) + pntpos(fin4Begin) + colbr - printObj.TextWidth("Groepstand (" & Format(getPnt(8), pntFormat) & "p)")) / 2
'                End If
'                printObj.Print "Groepstand (" & Format(getPnt(8), pntFormat) & "p)";
'                If grpAant > 4 Then
'                    printObj.CurrentX = (pntpos(fin8Begin) + pntpos(fin4Begin) + colbr - printObj.TextWidth("8e Finalisten (" & Format(getPnt(6), pntFormat) & "/" & Format(getPnt(7), pntFormat) & "p)")) / 2
'                    printObj.Print "8e Finalisten (" & Format(getPnt(4), pntFormat);
'                    If getPnt(5) > 0 Then
'                        printObj.Print "/" & Format(getPnt(5), pntFormat);
'                    End If
'                    printObj.Print "p)";
'                End If
'                printObj.CurrentX = (pntpos(fin4Begin) + pntpos(fin2Begin) + colbr - printObj.TextWidth("4e fin.(" & Format(getPnt(6), pntFormat) & "/" & Format(getPnt(7), pntFormat) & "p)")) / 2
'                printObj.Print "4efin.(" & Format(getPnt(6), pntFormat);
'                If getPnt(7) > 0 Then
'                    printObj.Print "/" & Format(getPnt(7), pntFormat);
'                End If
'                printObj.Print "p)";
'                printObj.CurrentX = (pntpos(fin2Begin) + pntpos(finBegin) + colbr - printObj.TextWidth("2efin.(" & Format(getPnt(9), pntFormat) & "/" & Format(getPnt(10), pntFormat) & "p)")) / 2
'                printObj.Print "1/2fin.(" & Format(getPnt(9), pntFormat);
'                If getPnt(10) > 0 Then
'                    printObj.Print "/" & Format(getPnt(10), pntFormat);
'                End If
'                printObj.Print "p)";
'                printObj.CurrentX = (pntpos(finBegin) + pntpos(EindstBegin) + colbr - printObj.TextWidth("Fin")) / 2
'                printObj.Print "Fin";
'                printObj.CurrentX = (pntpos(EindstBegin) + pntpos(AantBegin) + colbr - printObj.TextWidth("Eind")) / 2
'                printObj.Print "Eind";
'                printObj.CurrentX = (pntpos(AantBegin) + pntpos(TopScBegin) + colbr - printObj.TextWidth("Aantallen")) / 2
'                printObj.Print "Aantallen";
'                printObj.CurrentX = pntpos(TopScBegin) + colbr
'                printObj.Print "top";
'                printObj.CurrentX = (pntpos(TTLBegin) + pntpos(PosBegin) + colbr - printObj.TextWidth("Ttl")) / 2
'                printObj.Print "Ttl";
'                printObj.CurrentX = (pntpos(PosBegin) + pntpos(GeldBegin) + colbr - printObj.TextWidth("Pos")) / 2
'                printObj.Print "Pos";
'                printObj.CurrentX = (pntpos(GeldBegin) + pntpos(GeldBegin + 1) + colbr - printObj.TextWidth("Geld")) / 2
'                printObj.Print "Geld";
'                fontSizing 8
'                printObj.CurrentY = top2Ypos
'                printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
'                printObj.Print
'                printObj.Line (0, printObj.CurrentY)-(printObj.ScaleWidth - 50, printObj.CurrentY)
'              End If
'            End If
'        Loop
'    End With
'    botY = printObj.CurrentY
'    printObj.Line (pntpos(1) + 75, topYpos)-(pntpos(1) + 75, top2Ypos)
'    printObj.Line (pntpos(grpStndBegin) + 75, topYpos)-(pntpos(6) + 75, top2Ypos)
'    If grpAant > 4 Then
'        printObj.Line (pntpos(fin8Begin) + 75, topYpos)-(pntpos(15) + 75, top2Ypos)
'    End If
'    printObj.Line (pntpos(fin4Begin) + 75, topYpos)-(pntpos(fin4Begin) + 75, top2Ypos)
'    printObj.Line (pntpos(fin2Begin) + 75, topYpos)-(pntpos(fin2Begin) + 75, top2Ypos)
'    printObj.Line (pntpos(finBegin) + 75, topYpos)-(pntpos(finBegin) + 75, top2Ypos)
'    printObj.Line (pntpos(EindstBegin) + 75, topYpos)-(pntpos(EindstBegin) + 75, top2Ypos)
'    printObj.Line (pntpos(AantBegin) + 75, topYpos)-(pntpos(AantBegin) + 75, top2Ypos)
'    printObj.Line (pntpos(TopScBegin) + 75, topYpos)-(pntpos(TopScBegin) + 75, top2Ypos)
'    printObj.Line (pntpos(TTLBegin) + 75, topYpos)-(pntpos(TTLBegin) + 75, top2Ypos)
'    printObj.Line (pntpos(PosBegin) + 75, topYpos)-(pntpos(PosBegin) + 75, top2Ypos)
'    printObj.Line (pntpos(GeldBegin) + 75, topYpos)-(pntpos(GeldBegin) + 75, top2Ypos)
'    For i = 1 To UBound(pntpos) - 1
'        printObj.Line (pntpos(i) + 75, top2Ypos)-(pntpos(i) + 75, botY)
'    Next
'    printObj.Line (printObj.ScaleWidth - 50, topYpos)-(printObj.ScaleWidth - 50, botY)
'End Sub
'
'Function PrintAant(deelnem As Long, pos, vanwat As String)
'Dim aant As Integer
'Dim pnt As Long
'    Select Case vanwat
'    Case "pntRust"
'    pnt = getPnt(1)
'    Case "pntEind"
'    pnt = getPnt(2)
'    Case "pntToto"
'    pnt = getPnt(3)
'    Case "dpvddag"
'    pnt = getPnt(28)
'    End Select
'    If LCase(Left(vanwat, 6)) = "pntgrp" Then
'        pnt = getPnt(8)
'    End If
'
'    aant = getAant(deelnem, vanwat)
''    printObj.CurrentX = pos - printObj.TextWidth("(" & Format(aant, "0") & "x) " & Format(aant * pnt, "0"))
'    printObj.CurrentX = pos - printObj.TextWidth(Format(aant * pnt, "0"))
'    printobj.FontItalic = True
''    printObj.Print "(" & Format(aant, "0"); "x) ";
'    printobj.FontItalic = False
'    printObj.Print Format(aant * pnt, "0");
'    PrintAant = aant * pnt
'End Function
'
'Sub printPoolStandings(alfabet As Boolean, wedNum As Integer)
'' En nu de deelnemers
'Dim rsDeeln As New ADODB.Recordset
'Dim rsdeelnScore As New ADODB.Recordset
'Dim bedr As Currency
'Dim pnt As Integer
'Dim last As Integer
'Dim eerst As Integer
'Dim lastttl As Integer
'Dim verh As Double
'Dim geldold As Currency
'Dim savy As Integer
'Dim leftmarge As Integer
'Dim deelkolwidth%
'Dim DeelOldPntPos%
'Dim DeelWedPntPos%
'Dim DeelNewPntPos%
'Dim deelnaampos%
'Dim deelgeldpos%
'Dim DeelGeldnwPos%
'Dim DeelGeldttlPos%
'Dim Tekst$
'Dim prStr As String
'Dim yLinePos%
'Dim DeelTopPos%
'Dim i As Integer
'Dim tmp$
'Dim yposnu%
'    'wednum = GetWedNum(wednum)
'    leftmarge = printObj.CurrentX
'    deelkolwidth% = (printObj.ScaleWidth + 2 * printObj.ScaleLeft) \ 2
'    fontSizing 10
'    deelnaampos% = printObj.TextWidth("999.")
'    DeelOldPntPos% = deelnaampos% + deelkolwidth% / 4 - 200
'    DeelWedPntPos% = DeelOldPntPos% + deelkolwidth / 10
'    DeelNewPntPos% = DeelWedPntPos% + deelkolwidth / 10
'
'    deelgeldpos% = DeelNewPntPos% + deelkolwidth / 6 + 200
'    DeelGeldnwPos% = deelgeldpos% + deelkolwidth / 6 - 100
'    DeelGeldttlPos% = DeelGeldnwPos% + deelkolwidth / 6 - 100
'
'    If alfabet Then
'        deelnaampos% = Me.CurrentX + 40
'    End If
'
'    printObj.Print
'
'    fontSizing 16
'    printObj.fontBold = True
'    If alfabet Then
'        Tekst$ = "Resultaat (A-Z) na " & GetMyNum(wedNum) & "e wed: " & GetWedInfo(wedNum, "naam1") & "-" & GetWedInfo(wedNum, "naam2") & ": " & GetWedUitsl(wedNum)
'    Else
'        Tekst$ = "Stand na " & GetMyNum(wedNum) & "e wed: " & GetWedInfo(wedNum, "naam1") & "-" & GetWedInfo(wedNum, "naam2") & ": " & GetWedUitsl(wedNum)
'    End If
'    If Me.Eindstand Then
'        If alfabet Then
'            Tekst$ = "Eindstand alfabetisch"
'        Else
'            Tekst$ = "Eindstand"
'        End If
'    End If
'    headerText = GetOrgNaam(thisPool) & " " & getTournamentInfo("toernooi") & " voetbalpool"
'
'    heading1 = Tekst$
'
'    InitPage False, True
'    printobj.FontItalic = False
'    printObj.fontBold = False
'    fontSizing 10
'    printObj.CurrentX = (printObj.ScaleWidth - printObj.TextWidth("onderstreept=daghoogste, vet=bovenaan, cursief=onderaan")) / 2
'    printObj.Print "(";
'    printObj.FontUnderline = True
'    printObj.ForeColor = &H8000&
'    printObj.Print "onderstreept";
'    printObj.FontUnderline = False
'    printObj.ForeColor = 0
'    printObj.Print "= daghoogste, ";
'    printObj.ForeColor = vbBlue
'    printObj.fontBold = True
'    printObj.Print "vet";
'    printObj.fontBold = False
'    printObj.ForeColor = 0
'    printObj.Print "= bovenaan, ";
'    printobj.FontItalic = True
'    printObj.ForeColor = vbRed
'    printObj.Print "cursief";
'    printobj.FontItalic = False
'    printObj.ForeColor = 0
'    printObj.Print "= onderaan)"
'
'    savy = printObj.CurrentY
'    For kol% = 0 To 1
'        If Not alfabet Then
'            printObj.CurrentX = kol% * deelkolwidth%
'            'printObj.Print "pos";
'        End If
'        printObj.CurrentX = deelnaampos% + kol% * deelkolwidth%
'        printObj.Print "Naam";
'        If alfabet Then printObj.Print " (pl)";
'        printObj.CurrentX = DeelOldPntPos% + kol% * deelkolwidth%
'        printObj.Print "had  +";
'        printObj.CurrentX = DeelWedPntPos% + kol% * deelkolwidth%
'        printObj.Print "erbij =";
'        printObj.CurrentX = DeelNewPntPos% + kol% * deelkolwidth% + printObj.TextWidth("999") - printObj.TextWidth("nu")
'        printObj.Print "nu";
'        printObj.CurrentX = deelgeldpos% - printObj.TextWidth("Geld") + kol% * deelkolwidth%
'        printObj.Print "Geld";
'        printObj.CurrentX = DeelGeldnwPos% - printObj.TextWidth("erbij") + kol% * deelkolwidth%
'        printObj.Print "erbij";
'        printObj.CurrentX = DeelGeldttlPos% - printObj.TextWidth("totaal") + kol% * deelkolwidth%
'        printObj.Print "totaal";
'    Next
'    printObj.CurrentY = printObj.CurrentY + 50
'    yLinePos% = printObj.CurrentY + TextHeight("test")
'    printObj.Line (leftmarge, yLinePos%)-(printObj.ScaleWidth + printObj.ScaleLeft * 2, yLinePos%)
'    printObj.CurrentY = printObj.CurrentY + 50
'    DeelTopPos% = printObj.CurrentY
''    printObj.Print
'    'bepaal hoogste en laagste
'    rsDeeln.Open DeelnResultSql(False, wedNum), cn, adOpenStatic, adLockReadOnly 'op punten volgorde dus
'    If rsDeeln.RecordCount > 0 Then
'        rsDeeln.MoveLast
'        last = nz(rsDeeln!grandtotaal, 0)
'    Else
'        Exit Sub
'    End If
'    rsDeeln.Close
'    printObj.CurrentX = 0
'    'en nu opnieuw openen
'    rsDeeln.Open DeelnResultSql(alfabet, wedNum), cn, adOpenStatic, adLockReadOnly 'op volgorde dus
'    With rsDeeln
'        If .RecordCount > 0 Then
'            .MoveFirst
'            lastttl = 0
'            kol% = 0
'            Do While Not .EOF
'                i = i + 1
'                If i = Int(.RecordCount / 2 + 0.5) + 1 Then
'                    kol% = deelkolwidth%
'                    printObj.CurrentY = DeelTopPos%
'                End If
'                printObj.CurrentX = printObj.CurrentX + deelnaampos% - printObj.TextWidth(!postotaal) - printObj.TextWidth("..") + kol%
'                If Not alfabet Then
'                    If lastttl <> !grandtotaal Then printObj.Print !postotaal & ".";
'                End If
'                printObj.fontBold = !postotaal = 1
'                printobj.FontItalic = nz(!grandtotaal, 0) = last
'                prStr = Left(!bijnaam, 12)
'                If alfabet Then
'                    prStr = prStr & " (" & !postotaal & ")"
'                End If
'                If !grandtotaal = last Then
'                    printObj.ForeColor = vbRed
'                ElseIf nz(!postotaal, 0) = 1 Then
'                    printObj.ForeColor = vbBlue
'                ElseIf nz(!posdag, 0) = 1 Then
'                    printObj.ForeColor = &H8000&
'                Else
'                    printObj.ForeColor = 0
'                End If
'                printObj.CurrentX = deelnaampos% + kol%
'                printObj.FontUnderline = nz(!posdag, 0) = 1
'
'                printObj.Print prStr;
'                printObj.fontBold = False
'                printobj.FontItalic = False
'                printObj.ForeColor = 0
'                printObj.FontUnderline = False
'                If wedNum > 1 Then
'                    pnt = getTussenstand(!deelnemID, wedNum)
'                    geldold = getTussenstandGeld(!deelnemID, GetWedNumPrevDag(wedNum))
'                Else
'                    pnt = 0
'                    geldold = 0
'                End If
'
'                printObj.CurrentX = DeelOldPntPos% + kol% + printObj.TextWidth("999") - printObj.TextWidth(CStr(pnt))
'                printObj.Print Format$(pnt, "##0");
'                printObj.fontBold = False
'                pnt = nz(!Dagpnt, 0)
'                printObj.CurrentX = DeelWedPntPos% + kol% + printObj.TextWidth("999") - printObj.TextWidth(CStr(pnt))
'                printObj.FontUnderline = nz(!posdag, 0) = 1
'                If !posdag = 1 Then
'                    printObj.ForeColor = &H8000&
'                Else
'                    printObj.ForeColor = 0
'                End If
'                printObj.Print Format$(pnt, "##0");
'                printObj.ForeColor = 0
'                printObj.FontUnderline = False
'                printObj.fontBold = !postotaal = 1
'                printobj.FontItalic = nz(!grandtotaal, 0) = last
'                pnt = nz(!grandtotaal, 0)
'                If !grandtotaal = last Then
'                    printObj.ForeColor = vbRed
'                ElseIf !postotaal = 1 Then
'                    printObj.ForeColor = vbBlue
'                Else
'                    printObj.ForeColor = 0
'                End If
'                printObj.CurrentX = DeelNewPntPos% + kol% + printObj.TextWidth("999") - printObj.TextWidth(CStr(pnt))
'                If !grandtotaal = last Then
'                    printObj.ForeColor = &H80&
'                ElseIf !postotaal = 1 Then
'                    printObj.ForeColor = &HC00000
'                Else
'                    printObj.ForeColor = 0
'                End If
'                printObj.Print Format$(!grandtotaal, "##0");
'                printObj.ForeColor = 0
'                printObj.fontBold = False
'                printobj.FontItalic = False
'                tmp$ = Format$(geldold, "€ ##0.00")
'                printObj.CurrentX = deelgeldpos% - printObj.TextWidth(tmp$) + kol%
'                printObj.Print tmp$;   '= geld
'                tmp$ = Format$(!daggeldttl, "€ ##0.00")
'                printObj.CurrentX = DeelGeldnwPos% - printObj.TextWidth(tmp$) + kol%
'                printObj.Print tmp$;
'                bedr = 0
'                tmp$ = Format$(geldold + !daggeldttl, "€ ##0.00")
'                printObj.CurrentX = DeelGeldttlPos% - printObj.TextWidth(tmp$) + kol%
'                printObj.Print tmp$;   '= geld
'                printObj.Print
'                lastttl = nz(!grandtotaal, 0)
'                rsDeeln.MoveNext
'            Loop
'            printObj.Print
'            yposnu% = printObj.CurrentY
'            printObj.Line (deelkolwidth%, savy)-(deelkolwidth%, yposnu%)
'            printObj.Line (deelgeldpos - printObj.TextWidth("Geld") - 400, yLinePos%)-(deelgeldpos - printObj.TextWidth("Geld") - 400, yposnu%)
'            printObj.Line (deelgeldpos - printObj.TextWidth("Geld") - 400 + deelkolwidth%, yLinePos%)-(deelgeldpos - printObj.TextWidth("Geld") - 400 + deelkolwidth%, yposnu%)
'            printObj.Line (leftmarge, yposnu%)-(printObj.ScaleWidth + printObj.ScaleLeft * 2, yposnu%)
'        End If
'        .Close
'    End With
'    printObj.Print
'End Sub
'
'Function DeelnResultSql(alfabet As Boolean, wedNum As Integer) As String
'Dim sql As String
'    sql = "Select deelnemID, bijnaam, wednum,"
'    sql = sql & " deelnempnt.*"
'    sql = sql & " from deelnempnt, pooldeelnems"
'    sql = sql & " WHERE pooldeelnems!deelnemID = deelnempnt.deelnid"
'    sql = sql & " AND pooldeelnems!thisPool = " & thisPool
'    sql = sql & " AND wednum = " & wedNum
'    If alfabet Then
'        sql = sql & " ORDER BY bijnaam"
'    Else
'        sql = sql & " ORDER BY grandtotaal DESC, bijnaam ASC"
'    End If
'    DeelnResultSql = sql
'End Function
'
'
'Private Sub lstCompetitorPools_Click()
'    Me.Option4.value = True
'End Sub
'
'Private Sub txtVoorwed_Change()
'chkTxtValue Me.txtVoorWed, Me.vscrlVoor
'toMatch = val(txtVoorWed.Text)
'End Sub
'
'Private Sub txttoMatch_Change()
'chkTxtValue Me.txtToMatch, Me.vscrlTM
'toMatch = val(txtToMatch.Text)
'End Sub
'

Private Sub voetregel()
Dim printWidth
Dim i As Double
Dim fontnaam As String
    printobj.ForeColor = RGB(0, 51, 0)
    printWidth = printobj.DrawWidth
    printobj.DrawWidth = 1
    fontSizing 8
    printobj.FontItalic = True
    printobj.FontBold = False
    fontnaam = printobj.FontName
    printobj.FontName = "Garamond"
    printobj.CurrentY = printobj.ScaleHeight - printobj.TextHeight("w")
    voethoog = printobj.CurrentY - printobj.TextHeight("w")
    y% = printobj.CurrentY
    'printObj.Line (0, y% - 15 * printRatio)-(printObj.ScaleWidth, y% - 15 * printRatio)
    printobj.CurrentY = y%
    Centreer "© 2004-" & Year(Now) & " jota computer assistentie"
    printobj.FontName = fontnaam
    printobj.Print
    fontSizing 12
    printobj.FontBold = False
    printobj.FontItalic = False
    'y% = printObj.CurrentY + 50 * printRatio
    'printObj.Line (0, y%)-(printObj.ScaleWidth, y%)
    printobj.ForeColor = vbBlack
    printobj.DrawWidth = 1
End Sub

'
'
'Sub deelnemWedsInfo(inclpnt As Boolean)
'Dim infostr As String
'Dim pntToto As Integer
'Dim pntRust As Integer
'Dim pntEind As Integer
'Dim pntDp As Integer
'pntToto = getPntToek("toto goed")
'pntRust = getPntToek("ruststand goed")
'pntEind = getPntToek("eindstand goed")
'pntDp = getPntToek("doelpunten op een dag")
'infostr = "Samenstelling punten: rust goed "
'If inclpnt Then infostr = infostr & pntRust & " pnt"
'infostr = infostr & ", eindstand goed "
'If inclpnt Then infostr = infostr & pntEind & " pnt"
'infostr = infostr & ", toto goed "
'If inclpnt Then infostr = infostr & pntToto & " pnt, "
'infostr = infostr & ", aantal doelpunten van de dag goed "
'If inclpnt Then infostr = infostr & pntDp & " pnt"
'fontSizing 10
'printObj.CurrentX = (printObj.ScaleWidth - printObj.TextWidth(infostr)) / 2
'printObj.Print "Samenstelling punten: ";
'printobj.FontItalic = True
'printObj.Print "toto goed";
'If inclpnt Then printObj.Print pntToto; "pnt";
'printobj.FontItalic = False
'printObj.Print ", ";
'printObj.FontUnderline = True
'printObj.Print "rust goed";
'If inclpnt Then printObj.Print pntRust; "pnt";
'printObj.FontUnderline = False
'printObj.Print ", ";
'printObj.fontBold = True
'printObj.Print " eindstand goed";
'If inclpnt Then printObj.Print pntEind; "pnt";
'printObj.fontBold = False
'printObj.Print ", ";
'printObj.ForeColor = vbBlue
'printObj.Print "aantal doelpunten van de dag goed";
'If inclpnt Then printObj.Print pntDp; "pnt"
'printObj.ForeColor = 1
'printObj.CurrentY = printObj.CurrentY + 50
'
'
'End Sub
'
'Sub printPoolPointsPerMatch()
''print de deelnemers en hun punten per wedstrijd
'Dim rsDeeln As New ADODB.Recordset
'Dim rsDeelnPnt As New ADODB.Recordset
'Dim rsWeds As New ADODB.Recordset
'Dim sqlstr As String
'Dim xpos As Integer
'Dim posX() As Integer
'Dim x As Integer
'Dim i As Integer
'Dim topY As Integer
'Dim botY As Integer
'Dim topYpos As Integer
'Dim kolwidth As Integer
'Dim ttlKolWidth As Integer
'Dim wedstrijd As String
'Dim verttxtHeight 'de hoogte van de verticale text bovenin
'Dim infostr As String
'headerText = GetOrgNaam(thisPool) & " " & getTournamentInfo("toernooi") & " voetbalpool"
'heading1 = "Punten t/m wedstrijd " & toMatch
'InitPage False, True
'printObj.CurrentY = printObj.CurrentY - 50
'topYpos = printObj.CurrentY
'deelnemWedsInfo True 'druk de inforegel over de punten toekenning af
'topY = printObj.CurrentY
'printObj.Line (0, topY)-(printObj.ScaleWidth - 50, topY)
'fontSizing 8
'sqlstr = "SELECT pooldeelnems.deelnemID, pooldeelnems.bijnaam, deelnempnt.grandTotaal"
'sqlstr = sqlstr & " FROM (pooldeelnems INNER JOIN deelnempnt ON pooldeelnems.deelnemID = deelnempnt.deelnID) "
'sqlstr = sqlstr & " INNER JOIN toernschema ON deelnempnt.wedNum = toernschema.wedNum"
'sqlstr = sqlstr & " Where pooldeelnems.thisPool = " & thisPool
'sqlstr = sqlstr & " And toernschema.myNum = " & toMatch
'sqlstr = sqlstr & " And toernschema.ksid = " & kampID
'If Me.ScoreVolg(1) = True Then
'    sqlstr = sqlstr & " order by grandtotaal DESC"
'Else
'    sqlstr = sqlstr & " order by bijnaam"
'End If
'
'rsDeeln.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'sqlstr = "Select * from qryweds where ksid=" & kampID
''sqlstr = sqlstr & " AND wednum <=" & toMatch
'sqlstr = sqlstr & " order by mynum"
'rsWeds.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'verttxtHeight = printObj.TextWidth("1234567890123456789012345")
'printObj.CurrentY = verttxtHeight
'
'printObj.CurrentX = printObj.TextWidth(Left(GetLangsteBijNaam, 15))
'ReDim posX(1)
'posX(1) = printObj.CurrentX
'With rsWeds
'    If .RecordCount > 0 Then
'        .MoveFirst
'        Do While Not .EOF
'            rot.Angle = 90
'            printObj.CurrentX = posX(UBound(posX))
'            If !tm1 > "" Then
'                wedstrijd = !tm1 & "-"
'                If !tm2 > "" Then
'                    wedstrijd = wedstrijd & !tm2
'                Else
'                    wedstrijd = wedstrijd & !code2
'                End If
'            Else
'                wedstrijd = !code1 & "-"
'                If !tm2 > "" Then
'                    wedstrijd = wedstrijd & !tm2
'                Else
'                    wedstrijd = wedstrijd & !code2
'                End If
'            End If
'            rot.PrintText !mynum & ": " & wedstrijd
'            rot.Angle = 0
'            xpos = printObj.CurrentX + printObj.TextWidth("99") * 1.2
'            ReDim Preserve posX(UBound(posX) + 1)
'            posX(UBound(posX)) = xpos
'            rsWeds.MoveNext
'            'Debug.Print UBound(posX), posX(UBound(posX))
'        Loop
'    End If
'End With
'rot.Angle = 90
'printObj.CurrentX = posX(UBound(posX))
'rot.PrintText " pnt groepstand"
'
'If getTournamentInfo("groepen") > 4 Then
'    xpos = printObj.CurrentX + printObj.TextWidth("geld") * 1.2
'    ReDim Preserve posX(UBound(posX) + 1)
'    posX(UBound(posX)) = xpos
'    rot.Angle = 90
'    printObj.CurrentX = posX(UBound(posX))
'    rot.PrintText " 8e Finalisten"
'End If
'xpos = printObj.CurrentX + printObj.TextWidth("99") * 1.2
'ReDim Preserve posX(UBound(posX) + 1)
'posX(UBound(posX)) = xpos
'rot.Angle = 90
'printObj.CurrentX = posX(UBound(posX))
'rot.PrintText " Kw Finalisten"
'
'xpos = printObj.CurrentX + printObj.TextWidth("99") * 1.2
'ReDim Preserve posX(UBound(posX) + 1)
'posX(UBound(posX)) = xpos
'rot.Angle = 90
'printObj.CurrentX = posX(UBound(posX))
'rot.PrintText " Hv Finalisten"
'
'xpos = printObj.CurrentX + printObj.TextWidth("99") * 1.2
'ReDim Preserve posX(UBound(posX) + 1)
'posX(UBound(posX)) = xpos
'rot.Angle = 90
'printObj.CurrentX = posX(UBound(posX))
'rot.PrintText " Finalisten"
'
'xpos = printObj.CurrentX + printObj.TextWidth("99") * 1.2
'ReDim Preserve posX(UBound(posX) + 1)
'posX(UBound(posX)) = xpos
'rot.Angle = 90
'printObj.CurrentX = posX(UBound(posX))
'rot.PrintText " Eindstand"
'
'xpos = printObj.CurrentX + printObj.TextWidth("99") * 1.2
'ReDim Preserve posX(UBound(posX) + 1)
'posX(UBound(posX)) = xpos
'rot.Angle = 90
'printObj.CurrentX = posX(UBound(posX))
'rot.PrintText " Topscorers"
'
'xpos = printObj.CurrentX + printObj.TextWidth("99") * 1.2
'ReDim Preserve posX(UBound(posX) + 1)
'posX(UBound(posX)) = xpos
'rot.Angle = 90
'printObj.CurrentX = posX(UBound(posX))
'rot.PrintText " Overigen"
'
'xpos = printObj.CurrentX + printObj.TextWidth("99") * 1.2
'ReDim Preserve posX(UBound(posX) + 1)
'posX(UBound(posX)) = xpos
'rot.Angle = 90
'printObj.CurrentX = posX(UBound(posX))
'rot.PrintText " Totaal"
'
'xpos = printObj.CurrentX + printObj.TextWidth("999") * 1.2
'ReDim Preserve posX(UBound(posX) + 1)
'posX(UBound(posX)) = xpos
'rot.Angle = 90
'printObj.CurrentX = posX(UBound(posX))
'rot.PrintText " positie"
'
'xpos = printObj.CurrentX + printObj.TextWidth("99") * 1.2
'ReDim Preserve posX(UBound(posX) + 1)
'posX(UBound(posX)) = xpos
'rot.Angle = 90
'printObj.CurrentX = posX(UBound(posX))
'printObj.CurrentY = verttxtHeight - printObj.TextHeight("Geld")
''printObj.Print " geld";
'
'xpos = printObj.CurrentX + printObj.TextWidth("geld") * 1.2
'printObj.Print
'topYpos = printObj.CurrentY + 50
'ReDim Preserve posX(UBound(posX) + 1)
'posX(UBound(posX)) = xpos
'printObj.Line (0, topYpos)-(posX(UBound(posX)), topYpos)
'printObj.CurrentY = topYpos
'printObj.CurrentX = 0
'kolwidth = posX(2) - posX(1)
'botY = printObj.CurrentY
'pntFormat = "0;;\ ;-"
'
'Do While Not rsDeeln.EOF
'Dim naam As String
'    naam = rsDeeln!bijnaam
'   ' If InStr(naam, "Winner") > 0 Then Stop       1234567890
'    Do While printObj.TextWidth(naam) > printObj.TextWidth("123456789012345")
'        naam = Left(naam, Len(naam) - 1)
'    Loop
'    printObj.Print naam;
'    sqlstr = "SELECT toernschema.tijd, deelnemPnt.*, toernschema.gespeeld"
'    sqlstr = sqlstr & " FROM deelnemPnt INNER JOIN toernschema ON deelnemPnt.wedNum = toernschema.wedNum"
'    sqlstr = sqlstr & " Where toernschema.mynum <=" & toMatch
'    sqlstr = sqlstr & " AND toernschema.ksid = " & kampID
'    sqlstr = sqlstr & " AND deelnID = " & rsDeeln!deelnemID
'    sqlstr = sqlstr & " ORDER BY toernschema.mynum"
'    rsDeelnPnt.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    i = 0
'    With rsDeelnPnt
'        rot.Angle = 90
'        Do While Not .EOF
'            i = i + 1
'            printObj.CurrentX = posX(i) + (kolwidth - printObj.TextWidth(Format(nz(!pnttotaal, 0), pntFormat))) / 2
''            rot.Angle = 0
'            'If !pnttotaal = 7 Then Stop
'            printobj.FontItalic = nz(!pntToto, 0) <> 0
'            printObj.fontBold = nz(!pntEind, 0) <> 0
'            printObj.FontUnderline = nz(!pntRust, 0) > 0
'            If nz(!dpvddag, 0) > 0 Then
'                printObj.ForeColor = vbBlue
'            End If
'            printObj.Print Format(nz(!pnttotaal, 0), pntFormat);
'            printObj.fontBold = False
'            printobj.FontItalic = False
'            printObj.FontUnderline = False
'            printObj.ForeColor = 1
'
'            .MoveNext
'            rot.Angle = 90
'        Loop
'        If Not .RecordCount = 0 Then
'            .MoveLast
'            If !postotaal = 1 Then
'                printObj.ForeColor = &HC00000
'                printObj.FontBold = True
'            Else
'                printObj.ForeColor = vbBlack
'                printObj.FontBold = False
'            End If
'            ttlKolWidth = posX(UBound(posX) - 10) - posX(UBound(posX) - 11)
'
'            printObj.CurrentX = posX(UBound(posX) - 11) + (ttlKolWidth - printObj.TextWidth(Format(nz(!pntgrp, 0), pntFormat))) / 2
'            printObj.Print Format(nz(!pntgrp, 0), pntFormat);
'            ttlKolWidth = posX(UBound(posX) - 9) - posX(UBound(posX) - 10)
'            If getTournamentInfo("groepen") > 4 Then
'                printObj.CurrentX = posX(UBound(posX) - 10) + (ttlKolWidth - printObj.TextWidth(Format(nz(!pnt8fin, 0), pntFormat))) / 2
'                printObj.Print Format(nz(!pnt8fin, 0), pntFormat);
'                ttlKolWidth = posX(UBound(posX) - 8) - posX(UBound(posX) - 9)
'            End If
'            printObj.CurrentX = posX(UBound(posX) - 9) + (ttlKolWidth - printObj.TextWidth(Format(nz(!pntkwfin, 0), pntFormat))) / 2
'            printObj.Print Format(nz(!pntkwfin, 0), pntFormat);
'            ttlKolWidth = posX(UBound(posX) - 7) - posX(UBound(posX) - 8)
'            printObj.CurrentX = posX(UBound(posX) - 8) + (ttlKolWidth - printObj.TextWidth(Format(nz(!pnthvfin, 0), pntFormat))) / 2
'            printObj.Print Format(nz(!pnthvfin, 0), pntFormat);
'            ttlKolWidth = posX(UBound(posX) - 6) - posX(UBound(posX) - 7)
'            printObj.CurrentX = posX(UBound(posX) - 7) + (ttlKolWidth - printObj.TextWidth(Format(nz(!pntfin, 0) + nz(!pntklfin, 0), pntFormat))) / 2
'            printObj.Print Format(nz(!pntfin, 0) + nz(!pntklfin, 0), pntFormat);
'            ttlKolWidth = posX(UBound(posX) - 5) - posX(UBound(posX) - 6)
'            printObj.CurrentX = posX(UBound(posX) - 6) + (ttlKolWidth - printObj.TextWidth(Format(!pntuitslnaklfin + !pntuitsl, pntFormat))) / 2
'            printObj.Print Format(!pntuitslnaklfin + !pntuitsl, pntFormat);
'            ttlKolWidth = posX(UBound(posX) - 4) - posX(UBound(posX) - 5)
'            printObj.CurrentX = posX(UBound(posX) - 5) + (ttlKolWidth - printObj.TextWidth(Format(nz(!pntTopsc, 0) + nz(!pntOverig, 0), pntFormat))) / 2
'            printObj.Print Format(nz(!pntTopsc, 0), pntFormat);
'            ttlKolWidth = posX(UBound(posX) - 3) - posX(UBound(posX) - 4)
'            printObj.CurrentX = posX(UBound(posX) - 4) + (ttlKolWidth - printObj.TextWidth(Format(nz(!pntTopsc, 0) + nz(!pntOverig, 0), pntFormat))) / 2
'            printObj.Print Format(nz(!pntOverig, 0), pntFormat);
'            ttlKolWidth = posX(UBound(posX) - 2) - posX(UBound(posX) - 3)
'            printObj.CurrentX = posX(UBound(posX) - 3) + (ttlKolWidth - printObj.TextWidth(Format(nz(!grandtotaal, 0), pntFormat))) / 2
'            printObj.Print Format(nz(!grandtotaal, 0), pntFormat);
'            ttlKolWidth = posX(UBound(posX) - 1) - posX(UBound(posX) - 2)
'            printObj.CurrentX = posX(UBound(posX) - 2) + (ttlKolWidth - printObj.TextWidth(Format(nz(!postotaal, 0), pntFormat))) / 2
'            printObj.Print Format(nz(!postotaal, 0), pntFormat);
'            printObj.CurrentX = posX(UBound(posX)) - printObj.TextWidth(Format(nz(!geldttl, 0), "currency"))
'            printObj.ForeColor = vbBlack
'            printObj.FontItalic = False
'            printObj.FontBold = False
''            printObj.Print Format(nz(!geldttl, 0), "currency");
'        End If
'        printObj.Print
'    End With
'    printObj.Line (0, printObj.CurrentY + 10)-(posX(UBound(posX)), printObj.CurrentY + 10)
'    printObj.CurrentY = printObj.CurrentY + 10
'    printObj.CurrentX = 0
'    botY = printObj.CurrentY
''    If rsDeeln.AbsolutePosition = 67 Then Stop
'    If botY >= voethoog And rsDeeln.AbsolutePosition < rsDeeln.RecordCount Then
'        'nieuwe pagina
'        'eerste de lijntjes
'        For i = 1 To UBound(posX)
'            printObj.Line (posX(i), topY)-(posX(i), botY)
'        Next
'        i = 0
'        DoNewPage False, True
'        printObj.CurrentY = printObj.CurrentY - 50
'        topYpos = printObj.CurrentY
'        deelnemWedsInfo True 'druk de inforegel over de punten toekenning af
'        topY = printObj.CurrentY
'        printObj.Line (0, topY)-(printObj.ScaleWidth - 50, topY)
'        fontSizing 8
'        printObj.CurrentY = verttxtHeight
'        printObj.CurrentX = printObj.TextWidth("123456789012345")
'        With rsWeds
'            If .RecordCount > 0 Then
'                .MoveFirst
'                Do While Not .EOF
'                    Set rot.Device = printObj
'                    i = i + 1
'                    rot.Angle = 90
'                    printObj.CurrentX = posX(i)
'                    If !tm1 > "" Then
'                        rot.PrintText !mynum & ": " & !tm1 & "-" & !tm2
'                    Else
'                        rot.PrintText !mynum & ": " & !code1 & "-" & !code2
'                    End If
'                    rot.Angle = 0
'                    .MoveNext
'                Loop
'            End If
'        End With
'        rot.Angle = 90
'        If getTournamentInfo("groepen") > 4 Then
'            x = 11
'        Else
'            x = 10
'        End If
'        printObj.CurrentX = posX(UBound(posX) - x)
'        x = x - 1
'        rot.PrintText " pnt groepstand"
'        If getTournamentInfo("groepen") > 4 Then
'            printObj.CurrentX = posX(UBound(posX) - x)
'            x = x - 1
'            rot.PrintText " 8e Finalisten"
'        End If
'        printObj.CurrentX = posX(UBound(posX) - x)
'        x = x - 1
'        rot.PrintText " Kw Finalisten"
'        printObj.CurrentX = posX(UBound(posX) - x)
'        x = x - 1
'        rot.PrintText " Hv Finalisten"
'        printObj.CurrentX = posX(UBound(posX) - x)
'        x = x - 1
'        rot.PrintText " Finalisten"
'        printObj.CurrentX = posX(UBound(posX) - x)
'        x = x - 1
'        rot.PrintText " Eindstand"
'        printObj.CurrentX = posX(UBound(posX) - x)
'        x = x - 1
'        rot.PrintText " Topscorers"
'        printObj.CurrentX = posX(UBound(posX) - x)
'        x = x - 1
'        rot.PrintText " Overigen"
'        printObj.CurrentX = posX(UBound(posX) - x)
'        x = x - 1
'        rot.PrintText " Totaal"
'        printObj.CurrentX = posX(UBound(posX) - x)
'        x = x - 1
'        rot.PrintText " positie"
'        printObj.CurrentX = posX(UBound(posX) - x)
'        printObj.CurrentY = verttxtHeight ' - printObj.TextHeight("Geld")
' '       printObj.Print " geld"
'        topYpos = printObj.CurrentY + 50
'        printObj.Line (0, topYpos)-(posX(UBound(posX)), topYpos)
'        printObj.CurrentY = topYpos
'        printObj.CurrentX = 0
'        i = i + 1
'    End If
'    rsDeeln.MoveNext
'    rsDeelnPnt.Close
'Loop
'rsDeeln.Close
'For i = 1 To UBound(posX)
'    printObj.Line (posX(i), topY)-(posX(i), botY)
'Next
'i = 0
'Set rsDeeln = Nothing
'Set rsDeelnPnt = Nothing
'End Sub
'
'Sub DeelnemWedsPos()
'Dim rsDeeln As New ADODB.Recordset
'Dim rsDeelnPnt As New ADODB.Recordset
'Dim rsWeds As New ADODB.Recordset
'Dim sqlstr As String
'Dim xpos As Integer
'Dim posX() As Integer
'Dim i As Integer
'Dim topY As Integer
'Dim botY As Integer
'Dim topYpos As Integer
'Dim kolwidth As Integer
'Dim ttlKolWidth As Integer
'Dim verttxtHeight 'de hoogte van de verticale text bovenin
'Dim infostr As String
'headerText = GetOrgNaam(thisPool) & " " & getTournamentInfo("toernooi") & " voetbalpool"
'heading1 = "Positie in de pool na elke wedstrijd"
'InitPage False, True
'printObj.CurrentY = printObj.CurrentY - 50
'topYpos = printObj.CurrentY
'deelnemWedsInfo False 'druk de inforegel over de punten toekenning af
'topY = printObj.CurrentY
'printObj.Line (0, topY)-(printObj.ScaleWidth - 50, topY)
'fontSizing 8
'sqlstr = "SELECT pooldeelnems.deelnemID, pooldeelnems.bijnaam, deelnempnt.grandTotaal"
'sqlstr = sqlstr & " FROM (pooldeelnems INNER JOIN deelnempnt ON pooldeelnems.deelnemID = deelnempnt.deelnID) "
'sqlstr = sqlstr & " INNER JOIN toernschema ON deelnempnt.wedNum = toernschema.wedNum"
'sqlstr = sqlstr & " Where pooldeelnems.thisPool = " & thisPool
'sqlstr = sqlstr & " And toernschema.myNum = " & toMatch
'sqlstr = sqlstr & " And toernschema.ksid = " & kampID
'If Me.ScoreVolg(1) = True Then
'    sqlstr = sqlstr & " order by grandtotaal DESC"
'Else
'    sqlstr = sqlstr & " order by bijnaam"
'End If
'
'rsDeeln.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'sqlstr = "Select * from qryweds where ksid=" & kampID
''sqlstr = sqlstr & " AND wednum <=" & toMatch
'sqlstr = sqlstr & " order by mynum"
'rsWeds.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'verttxtHeight = printObj.TextWidth("1234567890123456789012345")
'printObj.CurrentY = verttxtHeight
'printObj.CurrentX = printObj.TextWidth("1234567890")
'ReDim posX(1)
'posX(1) = printObj.CurrentX
'With rsWeds
'    Do While Not .EOF
'        rot.Angle = 90
'        printObj.CurrentX = posX(UBound(posX))
'        If !tm1 > "" Then
'            rot.PrintText !mynum & ": " & !tm1 & "-" & !tm2
'        Else
'            rot.PrintText !mynum & ": " & !code1 & "-" & !code2
'        End If
'        rot.Angle = 0
'        xpos = printObj.CurrentX + printObj.TextWidth("99") * 1.3
'        ReDim Preserve posX(UBound(posX) + 1)
'        posX(UBound(posX)) = xpos
'        .MoveNext
'    Loop
'End With
'
''printObj.Print
'topYpos = printObj.CurrentY + 50
'ReDim Preserve posX(UBound(posX) + 1)
'posX(UBound(posX)) = xpos
'printObj.Line (0, topYpos)-(posX(UBound(posX)), topYpos)
'printObj.CurrentY = topYpos
'printObj.CurrentX = 0
'kolwidth = posX(2) - posX(1)
'botY = printObj.CurrentY
'pntFormat = "0;;\ ;-"
'
'Do While Not rsDeeln.EOF
'    printObj.Print rsDeeln!bijnaam;
'    sqlstr = "SELECT toernschema.tijd, deelnemPnt.*, toernschema.gespeeld"
'    sqlstr = sqlstr & " FROM deelnemPnt INNER JOIN toernschema ON deelnemPnt.wedNum = toernschema.wedNum"
'    sqlstr = sqlstr & " Where toernschema.mynum <=" & toMatch
'    sqlstr = sqlstr & " AND toernschema.ksid = " & kampID
'    sqlstr = sqlstr & " AND deelnID = " & rsDeeln!deelnemID
'    sqlstr = sqlstr & " ORDER BY toernschema.mynum"
'    rsDeelnPnt.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    i = 0
'    With rsDeelnPnt
'        rot.Angle = 90
'        Do While Not .EOF
'            i = i + 1
'            printObj.CurrentX = posX(i) + (kolwidth - printObj.TextWidth(Format(nz(!postotaal, 0), pntFormat))) / 2
'            printobj.FontItalic = nz(!pntToto, 0) <> 0
'            printObj.fontBold = nz(!pntEind, 0) <> 0
'            printObj.FontUnderline = nz(!pntRust, 0) > 0
'            If nz(!dpvddag, 0) > 0 Then
'                printObj.ForeColor = vbBlue
'            End If
'            printObj.Print Format(nz(!postotaal, 0), pntFormat);
'            printObj.fontBold = False
'            printobj.FontItalic = False
'            printObj.FontUnderline = False
'            printObj.ForeColor = 1
'
'            .MoveNext
'            rot.Angle = 90
'        Loop
'        printObj.Print
'    End With
'    printObj.Line (0, printObj.CurrentY + 10)-(posX(UBound(posX)), printObj.CurrentY + 10)
'    printObj.CurrentY = printObj.CurrentY + 10
'    printObj.CurrentX = 0
'    botY = printObj.CurrentY
'    If botY >= voethoog Then
'        'nieuwe pagina
'        'eerste de lijntjes
'        For i = 1 To UBound(posX)
'            printObj.Line (posX(i), topY)-(posX(i), botY)
'        Next
'        i = 0
'        DoNewPage False, True
'        printObj.CurrentY = printObj.CurrentY - 50
'        topYpos = printObj.CurrentY
'        deelnemWedsInfo False 'druk de inforegel over de punten toekenning af
'        topY = printObj.CurrentY
'        printObj.Line (0, topY)-(printObj.ScaleWidth - 50, topY)
'        fontSizing 8
'        printObj.CurrentY = verttxtHeight
'        printObj.CurrentX = printObj.TextWidth("123456789012345")
'        With rsWeds
'            If .RecordCount > 0 Then
'                .MoveFirst
'                Do While Not .EOF
'                    Set rot.Device = printObj
'                    i = i + 1
'                    rot.Angle = 90
'                    printObj.CurrentX = posX(i)
'                    If !tm1 > "" Then
'                        rot.PrintText !mynum & ": " & !tm1 & "-" & !tm2
'                    Else
'                        rot.PrintText !mynum & ": " & !code1 & "-" & !code2
'                    End If
'                    rot.Angle = 0
'                    .MoveNext
'                Loop
'            End If
'        End With
'        'printObj.Print
'        topYpos = printObj.CurrentY + 50
'        printObj.Line (0, topYpos)-(posX(UBound(posX)), topYpos)
'        printObj.CurrentY = topYpos
'        printObj.CurrentX = 0
'        i = i + 1
'    End If
'    rsDeeln.MoveNext
'Loop
'For i = 1 To UBound(posX)
'    printObj.Line (posX(i), topY)-(posX(i), botY)
'Next
'i = 0
'
'
'End Sub
'
'Sub printMatchPredictions(wedNum As Integer)
'Dim sqlstr As String
'Dim rs As New ADODB.Recordset
'Dim rsDeeln As New ADODB.Recordset
'Dim cloneRS As ADODB.Recordset
'Dim zoekstr As String
'Dim kopje As String
'Dim xpos As Integer
'Dim cols(4) As Integer
'Dim naampos
'Dim rijen As Integer
'Dim rijnu As Integer
'Dim yStart As Integer
'Dim lineXstart As Integer
'Dim lineYstart As Integer
'Dim lineXend As Integer
'Dim lineYend As Integer
'Dim koppos(3) As Integer
'Dim col As Integer
'Dim i As Integer
'wedNum = GetWedNum(wedNum)
'    headerText = GetOrgNaam(thisPool) & " " & getTournamentInfo("toernooi") & " voetbalpool" & " - Voorspelling"
'    If Not Me.optPortrait Then
'        cols(0) = 0
'        cols(1) = printObj.ScaleWidth / 4
'        cols(2) = printObj.ScaleWidth / 2
'        cols(3) = printObj.ScaleWidth / 4 * 3
'        cols(4) = printObj.ScaleWidth
'        col = 4
'    Else
'        cols(0) = 0
'        cols(1) = printObj.ScaleWidth / 3
'        cols(2) = printObj.ScaleWidth / 3 * 2
'        cols(3) = printObj.ScaleWidth
'        cols(4) = printObj.ScaleWidth
'        col = 3
'    End If
'    kopje = Format(GetWedInfo(wedNum, "datum"), "ddd d mmm") & " "
'    kopje = kopje & Format(GetWedInfo(wedNum, "tijd"), "HH:MM") & ": "
'    kopje = kopje & GetWedInfo(wedNum, "naam1") & " vs " & GetWedInfo(wedNum, "naam2")
'    heading1 = "Wedstrijd " & GetMyNum(wedNum) & ": " & kopje
'    InitPage False, True
'
'    printObj.Print
'    koppos(0) = 50
'    koppos(1) = printObj.TextWidth("0-000")
'    koppos(2) = koppos(1) + printObj.TextWidth("0-000")
'    koppos(3) = koppos(2) + printObj.TextWidth("0-000")
'    printObj.ForeColor = RGB(0, 51, 0)
'    For i = 0 To col - 1
'        printObj.CurrentX = cols(i) + koppos(0)
'        printObj.Print "Rust";
'        printObj.CurrentX = cols(i) + koppos(1)
'        printObj.Print "Eind";
'        printObj.CurrentX = cols(i) + koppos(2)
'        printObj.Print "Toto";
'        printObj.CurrentX = cols(i) + koppos(3)
'        printObj.Print "Wie";
'    Next
'    printObj.ForeColor = 0
'    printObj.Print
'    yStart = printObj.CurrentY
'    sqlstr = "SELECT e1, e2, r1,r2,toto, wednum "
'    sqlstr = sqlstr & " FROM voorspelling_uitsl INNER JOIN "
'    sqlstr = sqlstr & " pooldeelnems ON voorspelling_uitsl.deelnem = pooldeelnems.deelnemID"
'    sqlstr = sqlstr & " GROUP BY e1, e2, r1, r2, toto, wednum, poolid"
'    sqlstr = sqlstr & " HAVING wednum=" & wedNum
'    sqlstr = sqlstr & " AND pooldeelnems.poolid= " & thisPool
'    sqlstr = sqlstr & " ORDER BY r1,r2,e1,e2,toto"
'    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    sqlstr = "SELECT e1, e2, r1,r2,toto, wednum, bijnaam "
'    sqlstr = sqlstr & " FROM voorspelling_uitsl INNER JOIN "
'    sqlstr = sqlstr & " pooldeelnems ON voorspelling_uitsl.deelnem = pooldeelnems.deelnemID"
'    sqlstr = sqlstr & " WHERE wednum = " & wedNum
'    sqlstr = sqlstr & " AND poolid = " & thisPool
'    sqlstr = sqlstr & " ORDER BY bijnaam"
'    rsDeeln.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    rsDeeln.MoveLast
'    rijen = Int(rsDeeln.RecordCount / col + 0.5) + 1
'    rsDeeln.MoveFirst
'    rs.MoveFirst
'    i = 0
'    Do While Not rs.EOF
'        Set cloneRS = rsDeeln.Clone
'        zoekstr = "e1 = " & rs!e1
'        zoekstr = zoekstr & " and e2 = " & rs!e2
'        zoekstr = zoekstr & " and r1 = " & rs!r1
'        zoekstr = zoekstr & " and r2 = " & rs!r2
'        zoekstr = zoekstr & " and toto = " & rs!toto
'        cloneRS.Filter = zoekstr
'        If cloneRS.EOF Or cloneRS.BOF Then
'            rsDeeln.MoveLast
'            rsDeeln.MoveNext
'        End If
'        'rsDeeln.Find zoekstr, , , 0
'        printObj.CurrentX = cols(i)
'        lineXstart = printObj.CurrentX
'        lineYstart = printObj.CurrentY
'        printObj.CurrentX = cols(i) + koppos(0)
'        printObj.Print rs!r1 & "-" & rs!r2;
'        printObj.CurrentX = cols(i) + koppos(1)
'        printObj.fontBold = True
'        printObj.Print rs!e1 & "-" & rs!e2;
'        printObj.fontBold = False
'        printObj.CurrentX = cols(i) + koppos(2)
'        printObj.Print rs!toto;
'        cloneRS.MoveFirst
'        Do While Not cloneRS.EOF
'            printObj.CurrentX = cols(i) + koppos(3)
'            printObj.Print cloneRS!bijnaam
'            rijnu = rijnu + 1
'            cloneRS.MoveNext
'        Loop
'        lineXend = cols(i + 1) - 100
'        lineYend = printObj.CurrentY
'        printObj.Line (lineXstart, lineYstart)-(lineXend, lineYend), , B
'        rs.MoveNext
'        If rijnu >= rijen Then
'            i = i + 1
'            printObj.CurrentY = yStart
'            rijnu = 0
'        End If
'        cloneRS.Close
'        Set cloneRS = Nothing
'    Loop
'    rs.Close
'    rsDeeln.Close
'    Set rs = Nothing
'    Set rsDeeln = Nothing
'End Sub
Sub SetForeCol(kl As Long)
Dim r As Integer
Dim g As Integer
Dim b As Integer
    r = &HFF& And kl
    g = (&HFF00& And kl) \ 256
    b = (&HFF0000 And kl) \ 65536
    If r * 0.3 + g * 0.59 + b * 0.11 < 128 Then
        printobj.ForeColor = vbWhite
    Else
        printobj.ForeColor = vbBlack
    End If

End Sub

Sub MakeColors()
Dim i As Integer
Dim A As Integer
Dim C As Integer
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
For A = 0 To 64
    i = Int(Rnd() * klCol.Count) + 1
    kleur(A) = klCol(i)
    klCol.Remove i
Next
End Sub
