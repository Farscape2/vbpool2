VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPrinting 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Afdrukken"
   ClientHeight    =   5490
   ClientLeft      =   1665
   ClientTop       =   2430
   ClientWidth     =   9720
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
   Icon            =   "frmPrint.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5490
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton ScoreVolg 
      Appearance      =   0  'Flat
      Caption         =   "Op score"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   390
      Index           =   1
      Left            =   4515
      TabIndex        =   41
      Top             =   0
      Width           =   1080
   End
   Begin VB.OptionButton ScoreVolg 
      Appearance      =   0  'Flat
      Caption         =   "Alfabetisch"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   330
      Index           =   0
      Left            =   3240
      TabIndex        =   40
      Top             =   30
      Value           =   -1  'True
      Width           =   1275
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   3000
      ScaleHeight     =   2145
      ScaleWidth      =   2670
      TabIndex        =   14
      Top             =   2160
      Width           =   2730
      Begin VB.CheckBox chkNwePagKop 
         Alignment       =   1  'Right Justify
         Caption         =   "Nwe pag kop"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   960
         TabIndex        =   39
         ToolTipText     =   "Print wel/niet de kopregels op een nieuwe pagina"
         Top             =   0
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.VScrollBar vscrlCopies 
         Height          =   285
         Left            =   870
         Min             =   1
         TabIndex        =   25
         Top             =   960
         Value           =   1
         Width           =   180
      End
      Begin VB.TextBox copies 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   24
         Text            =   "1"
         Top             =   945
         Width           =   260
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   15
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Tag             =   "printer"
         Top             =   240
         Width           =   2415
      End
      Begin VB.OptionButton optLandscape 
         Caption         =   "Liggend"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   1350
         TabIndex        =   17
         Top             =   660
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.OptionButton optPortrait 
         Caption         =   "Staand"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   60
         TabIndex        =   16
         Top             =   645
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CheckBox chkDblSide 
         Alignment       =   1  'Right Justify
         Caption         =   "Dubbelzijdig"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   1155
         TabIndex        =   15
         Top             =   945
         Width           =   1185
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aantal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   45
         TabIndex        =   23
         Top             =   990
         Width           =   465
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Printer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   300
         Left            =   105
         TabIndex        =   22
         Top             =   0
         Width           =   585
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
      ScaleWidth      =   9660
      TabIndex        =   12
      Top             =   4620
      Width           =   9720
      Begin VB.CommandButton KlaarButton 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sluiten"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5040
         TabIndex        =   38
         Tag             =   "SluitPrintDial"
         Top             =   360
         Width           =   885
      End
      Begin VB.CommandButton btnPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Voorbeeld"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   3000
         TabIndex        =   37
         ToolTipText     =   "Bekijk een voorbeeld op het scherm"
         Top             =   360
         Width           =   885
      End
      Begin VB.CommandButton btnPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Afdrukken"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   3990
         TabIndex        =   36
         ToolTipText     =   "Stuur dit rapport naar de printer"
         Top             =   360
         Width           =   885
      End
      Begin VB.CommandButton cmdEindstand 
         Caption         =   "Eindstand voor deelnemers  afdrukken"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3015
         TabIndex        =   35
         Top             =   60
         Width           =   2910
      End
      Begin VB.CheckBox Eindstand 
         Appearance      =   0  'Flat
         Caption         =   "Eindstand"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   270
         Left            =   1800
         TabIndex        =   26
         Tag             =   "chkEinstand"
         Top             =   360
         Width           =   1065
      End
      Begin VB.CommandButton btnPrnDagResults 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Alles voor einde dag afdrukken"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   390
      Left            =   3360
      ScaleHeight     =   330
      ScaleWidth      =   2430
      TabIndex        =   9
      Top             =   1455
      Width           =   2490
      Begin VB.VScrollBar vscrlVoor 
         Height          =   285
         Left            =   1800
         Max             =   500
         Min             =   1
         TabIndex        =   19
         Top             =   30
         Value           =   1
         Width           =   180
      End
      Begin VB.TextBox txtVoorWed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1470
         TabIndex        =   10
         Text            =   "1"
         Top             =   15
         Width           =   345
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "Voor wedstrijd nr:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
         Width           =   1275
      End
   End
   Begin VB.PictureBox picTMwed 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   390
      Left            =   3360
      ScaleHeight     =   330
      ScaleWidth      =   2430
      TabIndex        =   6
      Top             =   1005
      Width           =   2490
      Begin VB.VScrollBar vscrlTM 
         Height          =   285
         Left            =   1800
         Max             =   500
         TabIndex        =   18
         Top             =   30
         Width           =   180
      End
      Begin VB.TextBox txtTMwed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1470
         TabIndex        =   7
         Text            =   "1"
         Top             =   15
         Width           =   345
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "T/m wedstrijd nr:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   75
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture3 
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      ForeColor       =   &H0000FFFF&
      Height          =   3555
      Left            =   120
      ScaleHeight     =   3495
      ScaleWidth      =   2220
      TabIndex        =   0
      Tag             =   "afdruk"
      Top             =   120
      Width           =   2280
      Begin VB.OptionButton optPrintSelect 
         Appearance      =   0  'Flat
         Caption         =   "Punten samenstelling"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
         TabIndex        =   34
         Top             =   1990
         Width           =   2235
      End
      Begin VB.OptionButton optPrintSelect 
         Appearance      =   0  'Flat
         Caption         =   "Voorspelling"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
         TabIndex        =   33
         Top             =   1260
         Width           =   2295
      End
      Begin VB.OptionButton optPrintSelect 
         Appearance      =   0  'Flat
         Caption         =   "Punten per wedstrijd"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
         TabIndex        =   32
         Top             =   2355
         Width           =   2235
      End
      Begin VB.OptionButton optPrintSelect 
         Appearance      =   0  'Flat
         Caption         =   "Stand in de Pool"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
         TabIndex        =   27
         Top             =   1625
         Width           =   2235
      End
      Begin VB.OptionButton optPrintSelect 
         Appearance      =   0  'Flat
         Caption         =   "Inschrijffomulieren"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
         Width           =   1845
      End
      Begin VB.OptionButton optPrintSelect 
         Appearance      =   0  'Flat
         Caption         =   "Ingevulde Pools"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
         Top             =   470
         Width           =   2670
      End
      Begin VB.OptionButton optPrintSelect 
         Appearance      =   0  'Flat
         Caption         =   "Stand in toernooi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
         Top             =   3090
         Width           =   2175
      End
      Begin VB.OptionButton optPrintSelect 
         Appearance      =   0  'Flat
         Caption         =   "Grafiek pool stand"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
         Top             =   2720
         Width           =   2295
      End
      Begin VB.OptionButton optPrintSelect 
         Appearance      =   0  'Flat
         Caption         =   "Favorieten"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Index           =   3
         Left            =   90
         TabIndex        =   1
         Top             =   835
         Width           =   1170
      End
   End
   Begin MSComDlg.CommonDialog prnDialog 
      Left            =   120
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FontName        =   "Tahoma"
   End
   Begin VB.Frame frmDeelnems 
      ForeColor       =   &H00004000&
      Height          =   2295
      Left            =   6720
      TabIndex        =   28
      Top             =   480
      Visible         =   0   'False
      Width           =   2490
      Begin VB.OptionButton Option4 
         Caption         =   "Selectie"
         ForeColor       =   &H00004000&
         Height          =   330
         Left            =   1395
         TabIndex        =   31
         Top             =   1875
         Width           =   990
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Allemaal"
         ForeColor       =   &H00004000&
         Height          =   330
         Left            =   120
         TabIndex        =   30
         Top             =   1920
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.ListBox lstDeelnems 
         Height          =   1620
         Left            =   135
         MultiSelect     =   1  'Simple
         TabIndex        =   29
         Top             =   165
         Width           =   2055
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Afdruk opties"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   7020
      TabIndex        =   20
      Top             =   1995
      Width           =   1215
   End
End
Attribute VB_Name = "frmPrinting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'om af te drukken op gekleurde achtergrond
Private Declare Function SetBkMode Lib "gdi32" _
  (ByVal hdc As Long, ByVal nBkMode As Long) As Long

Private Declare Function GetBkMode Lib "gdi32" _
  (ByVal hdc As Long) As Long

Private Const TRANSPARENT = 1
Private Const OPAQUE = 2

Private iBKMode As Long

Dim headerText
Dim tillMatch As Integer
Dim columnWidth As Integer
Dim header2 As String
Dim columnNr As Integer
Dim vertPos As Integer
'voor de favorieten afdruk
Dim favYpos As Integer
Dim favXpos As Integer

Dim header2Font As String
Dim txtFont As String

Dim rotate As New rotator

Dim horPos As Integer

Dim lineHeight As Integer
Dim NormalHeight As Integer
Dim largeHeight As Integer
Dim smallHeight As Integer
Dim verySmallHeight As Integer
Dim headerHeight As Integer
Dim footerHeight As Integer
Dim printFont As String
Dim thisMatch As Integer
Dim printObj As Object 'printer or print preview
Dim maxY As Integer 'voor afdrukken van Favorieten

Dim prntColor(64) As Long 'voor grafiek

Dim frmPrnt As frmPreview

Private Sub Afdruk_Click(Index As Integer)
Dim i As Integer
Me.frmDeelnems.Visible = False
Afdruk(Index).value = True
Select Case Index
  Case 0
    Me.pictillMatch.Visible = False
    Me.picVoorWed.Visible = False
    Me.picVolgorde.Visible = False
    Me.optPortrait.value = True
    Me.frmDeelnems.Visible = False
   ' Me.chkDblSide.Value = 1

  Case 1
   'deelnemers met voorspellingen
    Me.pictillMatch.Visible = False
    Me.picVoorWed.Visible = False
    Me.picVolgorde.Visible = False
    Me.frmDeelnems.Visible = True
    'Me.txtVoorwed.SetFocus
    Me.optPortrait.value = True
    'txtVoorwed.SetFocus
  Case 2
    'score/ stand in de pool
    picVolgorde.Visible = True 'GetDeelnemAant(poolID) > 32
    picVoorWed.Visible = False
    pictillMatch.Visible = True
    Me.optPortrait.value = True
    Me.vscrlTM = GetMyNum(GetLastPlayed)
    If tillMatch > 0 Then
        Me.txttillMatch.SetFocus
    End If
  Case 3
    ' Favorieten
    Me.pictillMatch.Visible = False
    Me.picVoorWed.Visible = False
    Me.picVolgorde.Visible = False
    Me.optPortrait.value = True
    Me.frmDeelnems.Visible = False
  Case 4
    'Stand in toernooi
    'score/ stand in de pool
    Me.pictillMatch.Visible = False
    Me.picVoorWed.Visible = False
    Me.picVolgorde.Visible = False
    Me.optPortrait.value = True
    Me.frmDeelnems.Visible = False
    Me.vscrlTM = GetMyNum(GetLastPlayed())
    DoEvents
 Case 5
    'grafiek
    Me.picVolgorde.Visible = False
    Me.picVoorWed.Visible = False
    Me.pictillMatch.Visible = True
    Me.optLandscape.value = True
    Me.ScoreVolg(1) = True
    Me.vscrlTM = GetMyNum(GetLastPlayed())
    DoEvents
    tillMatch = Me.vscrlTM
  Case 6
    'punten per wedstrijd
    picVolgorde.Visible = True
    picVoorWed.Visible = False
    pictillMatch.Visible = True
    Me.vscrlTM = GetMyNum(GetLastPlayed())
    tillMatch = Me.vscrlTM
    DoEvents
    Me.frmDeelnems.Visible = False  'getKampInfo("groepen")
    Me.optLandscape.value = getKampInfo("groepen") > 4
    Me.optPortrait.value = Not Me.optLandscape.value
  Case 7
    'voorspelling per wedstrijd
    picVolgorde.Visible = False
    picVoorWed.Visible = True
    pictillMatch.Visible = False
    Me.optPortrait.value = True
    Me.optLandscape.value = False
    Me.vscrlVoor = GetMyNum(GetLastPlayed()) + 1
    Me.frmDeelnems.Visible = False
  Case 8
    'samenvatting stand
    'Stand in toernooi
    'score/ stand in de pool
    Me.pictillMatch.Visible = True
    Me.picVoorWed.Visible = False
    Me.picVolgorde.Visible = True
    Me.optLandscape.value = True
    Me.frmDeelnems.Visible = False
    Me.vscrlTM = GetMyNum(GetLastPlayed())
  End Select
End Sub

Sub horline(kleur As Integer)
    printObj.Line (0, printObj.CurrentY)-(printObj.ScaleWidth - 50, printObj.CurrentY), kleur
End Sub

Sub printCompetitorForms()
Dim txt As String
Dim i As Integer
Dim aant As Integer
Dim topY As Integer
Dim ypos As Integer
Dim xpos As Integer
Dim header2Anaam As String
Dim header2Vnaam As String
Dim sqlstr As String
'Dim rs As New ADODB.Recordset
    printObj.FillStyle = vbFSTransparent
    headerText = GetOrgNaam(poolID) & getKampInfo("toernooi") & " voetbalpool"
    header2$ = "Inschrijfformulier     inleg: " & Format(getPoolInfo("inleg"), "currency")
    printObj.FontName = "Times New Roman"
    InitPage False, True
    printObj.Print
    FontGr 12
    topY = printObj.CurrentY
    printObj.ForeColor = vbBlack
    Vet False
    FontGr 12
    printObj.CurrentY = topY
    FontGr 18
    printObj.Line (0, topY - 200)-(printObj.ScaleWidth + 2 * printObj.ScaleLeft, topY + printObj.TextHeight("WW") * 4 + 200), , B
    printObj.Print
    xpos = printObj.CurrentX + 200
    printObj.CurrentY = topY
    printObj.CurrentX = xpos
    printObj.Print "Naam: ....................................................... Telefoon....................................."
    printObj.CurrentY = topY + printObj.TextWidth("WW")
    printObj.CurrentX = xpos
    printObj.Print "Adres: ....................................................... Plaats.........................................."
    printObj.CurrentY = topY + printObj.TextWidth("WW") * 2
    printObj.CurrentX = xpos
    printObj.Print "Email: ....................................................... Betaald ";
    xpos = printObj.CurrentX
    ypos = printObj.CurrentY
    printObj.DrawWidth = 3
    printObj.Line (xpos, ypos)-(xpos + printObj.TextWidth("W"), ypos + printObj.TextHeight("W")), , B
    printObj.DrawWidth = 1
    printObj.CurrentY = ypos
    printObj.CurrentX = printObj.CurrentX + 30
    printObj.Print " bij............................"
    FontGr 4
    printObj.Print
    'sqlstr = "Select * from poolpnt Where poolID = " & poolID
    'rs.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
    FontGr 16
    Vet True
    printObj.ForeColor = vbBlue
    printObj.Print "Instructies"
    FontGr 11
    Vet False
    printObj.ForeColor = vbBlack
    printObj.Print "Hier onder (en op de achterkant) kun je voorspellingen invoeren voor de "; getKampInfo("toernooi");
    printObj.Print " van "; Format(getKampInfo("startdate"), "d MMMM yyyy"); " tot "; Format(getKampInfo("einddate"), "d MMMM yyyy")
    printObj.Print "Voor elke juiste voorspelling krijg je punten, bij de verschillende onderdelen staat hoeveel."
    printObj.Print "De voorspellingen hoeven niet te kloppen, bij een uitslag kun je bijvoorbeeld 1-0 bij de rust, 0-2 bij de eindstand en een 3 "
    printObj.Print "bij de toto invullen. Of je kunt een team dat je uitgeschakeld hebt in een volgende ronde toch weer opnemen."
    If getKampInfo("groepen") = 6 And getKampInfo("aantalTeams") = 24 Then ' de vier beste derde plaatsen naar kwart finales
      printObj.Print "De beste 4 derde plaatsen kwalificeren zich ook voor de 8e finales."
    End If
    FontGr 16
    Vet True
    printObj.ForeColor = vbBlue
    printObj.Print "Prijzen"
    FontGr 11
    Vet False
    printObj.ForeColor = vbBlack
    'printObj.Print "Na de finale worden de hoofdprijzen te verdeeld, maar ook per dag zijn er geldprijzen te winnen."
    Vet True
    printObj.Print "-  Per dag";
    Vet False
    printObj.Print " zijn de volgende geldprijzen te verdienen:"
    printObj.Print "  -  ";
    printObj.Print "Degene die op ";
    Ital True
    printObj.Print "één dag de meeste punten";
    Ital False
    printObj.Print " heeft verzameld, ";
    printObj.Print " krijgt daarvoor ";
    Vet True
    printObj.Print Format(getPoolInfo("dagprijs"), "currency")
    Vet False
    printObj.Print "  -  ";
    printObj.Print "Degene die na een dag in de ";
    Ital True
    printObj.Print "totaalstand bovenaan";
    Ital False
    printObj.Print " staat, ";
    printObj.Print " krijgt daarvoor ";
    Vet True
    printObj.Print Format(getPoolInfo("dagstandprijs"), "currency")
    Vet False
    printObj.Print "  -  ";
    printObj.Print "Degene die na een dag in de ";
    Ital True
    printObj.Print "totaalstand onderaan";
    Ital False
    printObj.Print " staat, ";
    printObj.Print " krijgt daarvoor als troost ";
    Vet True
    printObj.Print Format(getPoolInfo("daglaatste"), "currency")
    Vet False
    printObj.Print "  -  ";
    xpos = printObj.CurrentX
    printObj.Print "De punten voor de finalerondes tellen mee voor de dagprijs op de dag dat de teams bekend zijn"
    printObj.CurrentX = xpos
    printObj.Print "De punten voor de eindstand, topscorers en aantallen tellen op de dag van de finale mee voor de dagprijs"
    printObj.Print "-  ";
    Vet True
    printObj.Print "Aan het eind van het toernooi";
    Vet False
    printObj.Print " zijn de volgende geldprijzen te verdienen:"
    aant = getPoolInfo("prijslaatste")
    If aant > 0 Then
        printObj.Print "  -  ";
        xpos = printObj.CurrentX
        printObj.Print "De ";
        Ital True
        printObj.ForeColor = vbRed
        printObj.Print "rode lantaarn";
        printObj.ForeColor = vbBlack
        Ital False
        printObj.Print " ontvangt als troostprijs "; Format(aant, "currency")
    End If
    
    printObj.Print "  -  ";
    xpos = printObj.CurrentX
    printObj.Print "De ";
    Ital True
    printObj.Print "hoogste";
    Ital False
    printObj.Print " deelnemers in de totaalstand krijgen de volgende prijzen:"
    printObj.CurrentX = xpos
    
    printObj.Print "1e pl: ";
    Vet True
    printObj.Print Format(getPoolInfo("eerste"), "0%");
    Vet False
    If getPoolInfo("tweede") > 0 Then
        printObj.Print ", 2e pl: ";
        Vet True
        printObj.Print Format(getPoolInfo("tweede"), "0%");
        Vet False
    End If
    If getPoolInfo("derde") > 0 Then
        printObj.Print ", 3e pl: ";
        Vet True
        printObj.Print Format(getPoolInfo("derde"), "0%");
        Vet False
    End If
    If getPoolInfo("vierde") > 0 Then
        printObj.Print ", 4e pl: ";
        Vet True
        printObj.Print Format(getPoolInfo("vierde"), "0%");
        Vet False
    End If
    printObj.Print " van de totale inleg (minus de dagprijzen en de rode lantaarn)"
    printObj.Print "-  ";
    Ital True
    printObj.Print "Bij een gelijk aantal punten wordt de betreffende prijs verdeeld"
    Ital False
    'horline 1
    'groepsstanden
    FontGr 10
    printObj.Print
    vertPos = printObj.CurrentY
    horPos = printObj.CurrentX
    FontGr 14
    Vet True
    printObj.FillColor = &H808080
    printObj.FillStyle = vbFSSolid
    'printObj.BackColor = printObj.FillColor
    printObj.Line (horPos, vertPos - 10)-(printObj.ScaleWidth, vertPos + printObj.TextHeight("W") + 10), vbBlack, B
    printObj.CurrentY = vertPos
    printObj.CurrentX = horPos + 50
    iBKMode = SetBkMode(printObj.hdc, TRANSPARENT)
    printObj.ForeColor = vbWhite
    printObj.Print "Groepstanden";
    FontGr 10
    Vet False
    txt = " Vul in: 1 t/m 4 (" & getPntToek("groepstand per juist team") & " pnt per correcte invoer)"
    'printObj.CurrentX = printObj.ScaleWidth - printObj.TextWidth(txt)
    printObj.CurrentY = vertPos + 40
    printObj.Print txt;
    printObj.CurrentY = vertPos
    FontGr 14
    printObj.Print
    vertPos = printObj.CurrentY
    horPos = printObj.CurrentX
    FontGr 12
    printObj.FillStyle = vbFSTransparent
    printObj.Line (horPos, vertPos)-(printObj.ScaleWidth, vertPos + printObj.TextHeight("W") * 5), vbBlack, B
    printObj.FillStyle = vbFSTransparent
    columnWidth = printObj.ScaleWidth / getKampInfo("groepen")
    printObj.ForeColor = vbBlack
    For i = 1 To getKampInfo("groepen")
        FontGr 12
        horPos = columnWidth * (i - 1) + 50
        printObj.CurrentY = vertPos + 10
        printObj.CurrentX = horPos
        Vet True
        printObj.Print "Groep " & Chr(i + 64)
        Vet False
        printgroep i
    Next
    printObj.Print
    printObj.Font = "Times New Roman"
    FontGr 2
    printObj.Print
    FontGr 12
    printFinals
    printOverige
    header2$ = "Wedstrijdvoorspellingen"
    DoNewPage False, True
    formulierWeds
'    InvulFormAfdrukken
End Sub

Sub printOverige()
'invulformulier
Dim rs As New ADODB.Recordset
Dim topscAant As Integer
Dim ypos As Integer
Dim xpos As Integer
Dim newlinepos As Integer
Dim columnWidth As Integer
Dim i As Integer

Dim vertPos As Integer
Dim horPos As Integer
Dim pnt As Integer
Dim txt As String
    newlinepos = 0
    printObj.Print
    columnWidth = printObj.ScaleWidth / 4
    'eerst de eindstand
    ypos = printObj.CurrentY
    i = getPntToek("1e plaats(Kampioen)")
    
    If i > 0 Then
        'print 1e
        txt = "(" & i & "p)"
        printObj.Font = "Tahoma"
        vertPos = ypos
        printObj.CurrentY = vertPos
        printObj.CurrentX = 0
        horPos = printObj.CurrentX
        FontGr 14
        Vet True
        printObj.FillColor = &H808080
        printObj.FillStyle = vbFSSolid
        printObj.Line (horPos + 30, vertPos - 10)-(columnWidth - 30, vertPos + printObj.TextHeight("W")), vbBlack, B
        printObj.CurrentY = vertPos
        printObj.CurrentX = horPos + 80
        printObj.ForeColor = vbWhite
        printObj.Print "Eindstand "
        Vet False
        printObj.FillStyle = vbFSTransparent
        vertPos = printObj.CurrentY
        printObj.CurrentX = horPos + 80
        printObj.ForeColor = vbBlack
        FontGr 12
        printObj.Print "1e:";
        FontGr 14
        printObj.Line (horPos + 30, vertPos)-(columnWidth - 30, vertPos + printObj.TextHeight("W")), vbBlack, B
        printObj.CurrentY = vertPos + 20
        printObj.CurrentX = horPos + columnWidth - printObj.TextWidth(txt) + 20
        FontGr 10
        printObj.Print txt;
        printObj.CurrentY = vertPos
        FontGr 14
        printObj.Print
        For i = 2 To 4
            pnt = getPntToek(Format(i, "0") & "e plaats")
            If pnt > 0 Then
                vertPos = printObj.CurrentY
                txt = "(" & pnt & "p)"
                printObj.CurrentX = horPos + 80
                FontGr 12
                printObj.Print Format(i, "0") & "e:";
                FontGr 14
                printObj.Line (horPos + 30, vertPos)-(columnWidth - 30, vertPos + printObj.TextHeight("W")), vbBlack, B
                printObj.CurrentY = vertPos + 20
                printObj.CurrentX = horPos + columnWidth - printObj.TextWidth(txt) + 20
                FontGr 10
                printObj.Print txt;
                printObj.CurrentY = vertPos
                FontGr 14
                printObj.Print
                If newlinepos < printObj.CurrentY Then newlinepos = printObj.CurrentY
            End If
        Next
    End If
    'topscorers
    printObj.CurrentY = ypos
    i = getPntToek("topscorer 1")
    If i > 0 Then
        'print 1e
        txt = "(" & i & "p)"
        printObj.Font = "Tahoma"
        vertPos = ypos
        printObj.CurrentY = vertPos
        printObj.CurrentX = columnWidth
        horPos = printObj.CurrentX
        FontGr 14
        Vet True
        printObj.FillColor = &H808080
        printObj.FillStyle = vbFSSolid
        printObj.Line (horPos, vertPos - 10)-(horPos + columnWidth * 1.3, vertPos + printObj.TextHeight("W")), vbBlack, B '(horPos + columnWidth - 30, vertPos + printObj.TextHeight("W")), vbBlack, B
        printObj.CurrentY = vertPos
        printObj.CurrentX = horPos + 50
        printObj.ForeColor = vbWhite
        printObj.Print "Topscorer";
        If getPntToek("topscorer 2") > 0 Then printObj.Print "s";
        
        pnt = getPntToek("doelpunten topscorer 1")
        
        If pnt > 0 Then
            FontGr 14
            'printObj.Line (horPos + columnWidth - 30, vertPos - 10)-(horPos + columnWidth * 1.3, vertPos + printObj.TextHeight("W")), vbBlack, B
            printObj.CurrentY = vertPos
            'printObj.CurrentX = horPos + columnWidth + 20
            printObj.Print " & aantal goals"
        Else
            printObj.Print
        End If
        printObj.FillStyle = vbFSTransparent
        printObj.ForeColor = vbBlack
        Vet False
        For i = 1 To 3
            pnt = getPntToek("topscorer " & Format(i, "0"))
            If pnt > 0 Then
                vertPos = printObj.CurrentY
                txt = "(" & pnt & "p)"
                printObj.CurrentX = horPos + 50
                FontGr 12
                printObj.Print Format(i, "0") & ":";
                FontGr 14
                printObj.Line (horPos, vertPos)-(horPos + columnWidth - 30, vertPos + printObj.TextHeight("W")), vbBlack, B
                printObj.CurrentY = vertPos + 20
                printObj.CurrentX = horPos + columnWidth + 20 - printObj.TextWidth(txt)
                FontGr 10
                printObj.Print txt;
                FontGr 14
                pnt = getPntToek("doelpunten topscorer " & Format(i, "0"))
                If pnt > 0 Then
                    printObj.Line (horPos + columnWidth - 30, vertPos)-(horPos + columnWidth * 1.3, vertPos + printObj.TextHeight("W")), vbBlack, B
                    printObj.CurrentY = vertPos + 20
                    printObj.CurrentX = horPos + columnWidth * 1.3 - printObj.TextWidth("(" & pnt & "p)") + 50
                    FontGr 10
                    printObj.Print "("; Format(pnt, pntFormat); "p)"
                    If newlinepos < printObj.CurrentY Then newlinepos = printObj.CurrentY
                Else
                    printObj.Print
                End If
                printObj.CurrentY = vertPos
                FontGr 14
                printObj.Print
            End If
        Next
    End If
'overigen
Dim sqlstr As String
  sqlstr = "Select omschrijving, pnt, marge from voorspeltypes INNER JOIN pnttoek ON voorspeltypes.id = pnttoek.voorspeltype"
  sqlstr = sqlstr & " WHERE voorspeltypes.cat = 1 and pnttoek.poolid = " & poolID
  sqlstr = sqlstr & " ORDER BY pnt, volgorde"
  rs.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
'    rs.Open "Select * from voorspeltypes where cat =1 order by volgorde", dbConn, adOpenStatic, adLockReadOnly
    vertPos = ypos
    printObj.CurrentY = vertPos
    printObj.CurrentX = horPos + columnWidth * 1.3 + 30
    horPos = printObj.CurrentX
    FontGr 14
    Vet True
    printObj.FillColor = &H808080
    printObj.FillStyle = vbFSSolid
    printObj.Line (horPos, vertPos - 10)-(printObj.ScaleWidth - 50, vertPos + printObj.TextHeight("W")), vbBlack, B
    printObj.CurrentY = vertPos
    printObj.CurrentX = horPos + 50
    printObj.ForeColor = vbWhite
    printObj.Print "Overigen "
    printObj.FillStyle = vbFSTransparent
    printObj.ForeColor = vbBlack
    Vet False
    Do While Not rs.EOF
        pnt = rs!pnt
        vertPos = printObj.CurrentY
        txt = "(" & pnt & "p)"
        If nz(rs!marge, 0) > 0 Then
          txt = "(±" & rs!marge & ", " & pnt & "p)"
        End If
        printObj.CurrentX = horPos + 50
        FontGr 10
        printObj.Print rs!omschrijving; " "; txt; ":";
        FontGr 14
        printObj.Line (horPos, vertPos)-(printObj.ScaleWidth - 50, vertPos + printObj.TextHeight("W")), vbBlack, B
        rs.MoveNext
        If newlinepos < printObj.CurrentY Then newlinepos = printObj.CurrentY
    Loop
    rs.Close
    Set rs = Nothing
    printObj.Line (printObj.ScaleWidth - 30 - printObj.TextWidth("1234"), ypos + 360)-(printObj.ScaleWidth - 30 - printObj.TextWidth("1234"), printObj.CurrentY)
    printObj.Line (0, ypos - 50)-(printObj.ScaleWidth - 10, newlinepos + 30), , B
    
End Sub
Sub printFinals()
'onderdeel van formulieren
Dim rs As New ADODB.Recordset
Dim sqlstr As String
Dim xpos As Integer
Dim ypos As Integer
Dim i As Integer
Dim horPos As Integer
Dim vertPos As Integer
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
        printObj.Font = "Tahoma"
        vertPos = printObj.CurrentY
        horPos = printObj.CurrentX
        FontGr 14
        Vet True
        printObj.FillColor = &H808080
        printObj.FillStyle = vbFSSolid
        printObj.Line (horPos, vertPos - 10)-(printObj.ScaleWidth, vertPos + printObj.TextHeight("W")), vbBlack, B
        'printObj.BackColor = printObj.FillColor
        iBKMode = SetBkMode(printObj.hdc, TRANSPARENT)
        printObj.ForeColor = vbWhite
        printObj.CurrentY = vertPos
        printObj.CurrentX = horPos + 50
        printObj.Print "Achtstefinales ";
        printObj.FillStyle = vbFSTransparent
        FontGr 10
        Vet False
'        printObj.CurrentX = printObj.ScaleWidth - printObj.TextWidth(txt)
        printObj.CurrentY = vertPos + 40
        printObj.Print txt;
        printObj.ForeColor = vbBlack
        printObj.CurrentY = vertPos
        FontGr 14
        printObj.Print
        vertPos = printObj.CurrentY
        horPos = printObj.CurrentX
        FontGr 12
        printObj.Line (horPos, vertPos)-(printObj.ScaleWidth, vertPos + printObj.TextHeight("W") * 4.7), vbBlack, B
        vertPos = vertPos + 50
        printObj.CurrentY = vertPos
        printObj.FillStyle = vbFSTransparent
        columnWidth = printObj.ScaleWidth / 4
        sqlstr = "Select * from qryWeds where  ksid = " & kampID
        sqlstr = sqlstr & " and wedtype = 5"
        sqlstr = sqlstr & " ORDER BY wednum"
        rs.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
        xpos = 0
        With rs
            If .RecordCount > 0 Then
                i = 0
                Do While Not .EOF
                    ypos = vertPos
                    FontGr 8
                    printObj.CurrentX = xpos + 50
                    printObj.CurrentY = ypos + printObj.TextHeight("99") * 0.5
                    printObj.Print Format(!wedNum, "0"); ":";
                    FontGr 12
                    printObj.CurrentX = xpos + printObj.TextWidth("00:") + 30
                    printObj.CurrentY = ypos
                    FontGr 10
                    printObj.Print !code1; ":";
                    FontGr 12
                    printObj.DrawWidth = 1
                    printObj.Line (xpos + printObj.TextWidth("00:"), ypos)-(xpos + columnWidth - 50, ypos + printObj.TextHeight("W")), vbBlack, B
                    ypos = printObj.CurrentY
                    printObj.CurrentX = xpos + printObj.TextWidth("00:") + 30
                    FontGr 10
                    printObj.Print !code2; ":";
                    FontGr 12
                    printObj.Line (xpos + printObj.TextWidth("00:"), ypos)-(xpos + columnWidth - 50, ypos + printObj.TextHeight("W")), vbBlack, B
                    'wedstrijd nr
                    printObj.CurrentY = ypos
                    .MoveNext
                    i = i + 1
                    xpos = columnWidth * i
                    If xpos > printObj.ScaleWidth - columnWidth + 100 Then
                        FontGr 8
                        printObj.Print
                        printObj.Print
                        FontGr 10
                        vertPos = printObj.CurrentY
                        i = 0
                        xpos = 0
                    End If
                Loop
            End If
        End With
    End If
    FontGr 2
    printObj.Print
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
        printObj.Font = "Tahoma"
        vertPos = printObj.CurrentY
        horPos = printObj.CurrentX
        FontGr 14
        Vet True
        printObj.FillColor = &H808080
        printObj.FillStyle = vbFSSolid
        printObj.Line (horPos, vertPos - 10)-(printObj.ScaleWidth, vertPos + printObj.TextHeight("W")), vbBlack, B
        printObj.CurrentY = vertPos
        printObj.CurrentX = horPos + 50
        printObj.ForeColor = vbWhite
        printObj.Print "Kwartfinales ";
        FontGr 10
        Vet False
'        printObj.CurrentX = printObj.ScaleWidth - printObj.TextWidth(txt)
        printObj.CurrentY = vertPos + 40
        printObj.Print txt;
        printObj.ForeColor = vbBlack
        printObj.FillStyle = vbFSTransparent
        printObj.CurrentY = vertPos
        FontGr 14
        printObj.Print
        vertPos = printObj.CurrentY
        horPos = printObj.CurrentX
        FontGr 12
        printObj.Line (horPos, vertPos)-(printObj.ScaleWidth, vertPos + printObj.TextHeight("W") * 2.5), vbBlack, B
        vertPos = vertPos + 50
        printObj.CurrentY = vertPos
        printObj.FillStyle = vbFSTransparent
        columnWidth = (printObj.ScaleWidth / 8) * 2
        sqlstr = "Select * from qryWeds where  ksid = " & kampID
        sqlstr = sqlstr & " and wedtype = 2"
        sqlstr = sqlstr & " ORDER BY wednum"
        rs.Close
        rs.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
        xpos = 0
        With rs
            If .RecordCount > 0 Then
                i = 0
                Do While Not .EOF
                    ypos = vertPos
                    FontGr 8
                    printObj.CurrentX = xpos + 50
                    printObj.CurrentY = ypos + printObj.TextHeight("99") * 0.5
                    printObj.Print Format(!wedNum, "0"); ":";
                    FontGr 12
                    printObj.CurrentX = xpos + printObj.TextWidth("00:") + 30
                    printObj.CurrentY = ypos
                    FontGr 10
                    printObj.Print !code1; ":";
                    FontGr 12
                    printObj.DrawWidth = 1
                    printObj.Line (xpos + printObj.TextWidth("00:"), ypos)-(xpos + columnWidth - 50, ypos + printObj.TextHeight("W")), vbBlack, B
                    ypos = printObj.CurrentY
                    printObj.CurrentX = xpos + printObj.TextWidth("00:") + 30
                    FontGr 10
                    printObj.Print !code2; ":";
                    FontGr 12
                    printObj.Line (xpos + printObj.TextWidth("00:"), ypos)-(xpos + columnWidth - 50, ypos + printObj.TextHeight("W")), vbBlack, B
                    'wedstrijd nr
                    printObj.CurrentY = ypos
                    .MoveNext
                    i = i + 1
                    xpos = columnWidth * i
                    If xpos > printObj.ScaleWidth - columnWidth + 100 Then
                        FontGr 8
                        printObj.Print
                        printObj.Print
                        FontGr 12
                        vertPos = printObj.CurrentY
                        i = 0
                        xpos = 0
                    End If
                Loop
            End If
        End With
    End If
    FontGr 2
    printObj.Print
    FontGr 12
    hvFinYpos = printObj.CurrentY
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
        printObj.Font = "Tahoma"
        vertPos = printObj.CurrentY
        horPos = printObj.CurrentX
        FontGr 14
        Vet True
        printObj.FillColor = &H808080
        printObj.FillStyle = vbFSSolid
        printObj.Line (horPos, vertPos - 10)-(printObj.ScaleWidth / 2 - 30, vertPos + printObj.TextHeight("W")), vbBlack, B
        printObj.CurrentY = vertPos
        printObj.CurrentX = horPos + 50
        printObj.ForeColor = vbWhite
        printObj.Print "Halve finales ";
        FontGr 10
        Vet False
        'printObj.CurrentX = printObj.ScaleWidth / 2 - 30 - printObj.TextWidth(txt)
        printObj.CurrentY = vertPos + 40
        printObj.Print txt;
        printObj.ForeColor = vbBlack
        printObj.CurrentY = vertPos
        FontGr 14
        printObj.Print
        vertPos = printObj.CurrentY
        horPos = printObj.CurrentX
        printObj.FillStyle = vbFSTransparent
        FontGr 12
        printObj.Line (horPos, vertPos)-(printObj.ScaleWidth / 2 - 30, vertPos + printObj.TextHeight("W") * 2.5), vbBlack, B
        vertPos = vertPos + 50
        printObj.CurrentY = vertPos
        printObj.FillStyle = vbFSTransparent
        columnWidth = (printObj.ScaleWidth / 8) * 2
        sqlstr = "Select * from qryWeds where  ksid = " & kampID
        sqlstr = sqlstr & " and wedtype = 3"
        sqlstr = sqlstr & " ORDER BY wednum"
        rs.Close
        rs.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
        xpos = 0
        With rs
            If .RecordCount > 0 Then
                i = 0
                Do While Not .EOF
                    ypos = vertPos
                    FontGr 8
                    printObj.CurrentX = xpos + 50
                    printObj.CurrentY = ypos + printObj.TextHeight("99") * 0.5
                    printObj.Print Format(!wedNum, "0"); ":";
                    FontGr 12
                    printObj.CurrentX = xpos + printObj.TextWidth("00:") + 30
                    printObj.CurrentY = ypos
                    FontGr 10
                    printObj.Print !code1; ":";
                    FontGr 12
                    printObj.DrawWidth = 1
                    printObj.Line (xpos + printObj.TextWidth("00:"), ypos)-(xpos + columnWidth - 50, ypos + printObj.TextHeight("W")), vbBlack, B
                    ypos = printObj.CurrentY
                    printObj.CurrentX = xpos + printObj.TextWidth("00:") + 30
                    FontGr 10
                    printObj.Print !code2; ":";
                    FontGr 12
                    printObj.Line (xpos + printObj.TextWidth("00:"), ypos)-(xpos + columnWidth - 50, ypos + printObj.TextHeight("W")), vbBlack, B
                    'wedstrijd nr
                    printObj.CurrentY = ypos
                    .MoveNext
                    i = i + 1
                    xpos = columnWidth * i
                    If xpos > printObj.ScaleWidth - columnWidth + 100 Then
                        FontGr 8
                        printObj.Print
                        printObj.Print
                        FontGr 12
                        vertPos = printObj.CurrentY
                        i = 0
                        xpos = 0
                    End If
                Loop
            End If
        End With
    End If
    printObj.CurrentY = hvFinYpos
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
        printObj.Font = "Tahoma"
        vertPos = hvFinYpos
        printObj.CurrentY = vertPos
        printObj.CurrentX = printObj.ScaleWidth / 2 + 30
        horPos = printObj.CurrentX
        FontGr 14
        Vet True
        printObj.FillColor = &H808080
        printObj.FillStyle = vbFSSolid
        printObj.Line (horPos, vertPos - 10)-(printObj.ScaleWidth * 0.75, vertPos + printObj.TextHeight("W")), vbBlack, B
        printObj.CurrentY = vertPos
        printObj.CurrentX = horPos + 50
        printObj.ForeColor = vbWhite
        printObj.Print "3e plaats ";
        FontGr 10
        Vet False
        printObj.CurrentY = vertPos + 40
        printObj.Print txt;
        printObj.ForeColor = vbBlack
        printObj.CurrentY = vertPos
        FontGr 14
        printObj.Print
        printObj.FillStyle = vbFSTransparent
        vertPos = printObj.CurrentY
        horPos = printObj.ScaleWidth / 2 + 30
        FontGr 12
        printObj.Line (horPos, vertPos)-(printObj.ScaleWidth * 0.75, vertPos + printObj.TextHeight("W") * 2.5), vbBlack, B
        vertPos = vertPos + 50
        printObj.CurrentY = vertPos
        printObj.FillStyle = vbFSTransparent
        columnWidth = (printObj.ScaleWidth / 8) * 2
        sqlstr = "Select * from qryWeds where  ksid = " & kampID
        sqlstr = sqlstr & " and wedtype = 7"
        sqlstr = sqlstr & " ORDER BY wednum"
        rs.Close
        rs.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
        xpos = printObj.ScaleWidth / 2 + 30
        With rs
            If .RecordCount > 0 Then
                i = 0
                Do While Not .EOF
                    ypos = vertPos
                    FontGr 8
                    printObj.CurrentX = xpos + 50
                    printObj.CurrentY = ypos + printObj.TextHeight("99") * 0.5
                    printObj.Print Format(!wedNum, "0"); ":";
                    FontGr 12
                    printObj.CurrentX = xpos + printObj.TextWidth("00:") + 30
                    printObj.CurrentY = ypos
                    printObj.Print !code1; ":";
                    printObj.DrawWidth = 1
                    printObj.Line (xpos + printObj.TextWidth("00:"), ypos)-(xpos + columnWidth - 50, ypos + printObj.TextHeight("W")), vbBlack, B
                    ypos = printObj.CurrentY
                    printObj.CurrentX = xpos + printObj.TextWidth("00:") + 30
                    printObj.Print !code2; ":";
                    printObj.Line (xpos + printObj.TextWidth("00:"), ypos)-(xpos + columnWidth - 50, ypos + printObj.TextHeight("W")), vbBlack, B
                    'wedstrijd nr
                    printObj.CurrentY = ypos
                    .MoveNext
                    i = i + 1
                    xpos = columnWidth * i
                    If xpos > printObj.ScaleWidth - columnWidth + 100 Then
                        FontGr 8
                        printObj.Print
                        printObj.Print
                        FontGr 12
                        vertPos = printObj.CurrentY
                        i = 0
                        xpos = 0
                    End If
                Loop
            End If
        End With
    Else
        HeeftKlFin = False
    End If
    printObj.CurrentY = hvFinYpos
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
        printObj.Font = "Tahoma"
        vertPos = hvFinYpos
        printObj.CurrentY = vertPos
        If HeeftKlFin Then
            printObj.CurrentX = printObj.ScaleWidth * 0.75 + 30
        Else
            printObj.CurrentX = printObj.ScaleWidth * 0.5 + 30
        End If
        horPos = printObj.CurrentX
        FontGr 14
        Vet True
        printObj.FillColor = &H808080
        printObj.FillStyle = vbFSSolid
        printObj.Line (horPos, vertPos - 10)-(printObj.ScaleWidth, vertPos + printObj.TextHeight("W")), vbBlack, B
        printObj.ForeColor = vbWhite
        printObj.CurrentY = vertPos
        printObj.CurrentX = horPos + 50
        printObj.Print "Finale ";
        FontGr 10
        Vet False
        printObj.CurrentY = vertPos + 40
        printObj.Print txt;
        printObj.ForeColor = vbBlack
        printObj.CurrentY = vertPos
        FontGr 14
        printObj.Print
        printObj.FillStyle = vbFSTransparent
        vertPos = printObj.CurrentY
        If HeeftKlFin Then
            horPos = printObj.ScaleWidth * 0.75 + 30
        Else
            horPos = printObj.ScaleWidth * 0.5 + 30
        End If
        FontGr 12
        printObj.Line (horPos, vertPos)-(printObj.ScaleWidth, vertPos + printObj.TextHeight("W") * 2.5), vbBlack, B
        vertPos = vertPos + 50
        printObj.CurrentY = vertPos
        printObj.FillStyle = vbFSTransparent
        columnWidth = (printObj.ScaleWidth / 8) * 2
        sqlstr = "Select * from qryWeds where  ksid = " & kampID
        sqlstr = sqlstr & " and wedtype = 4"
        sqlstr = sqlstr & " ORDER BY wednum"
        rs.Close
        rs.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
        If HeeftKlFin Then
            xpos = printObj.ScaleWidth * 0.75 + 30
        Else
            xpos = printObj.ScaleWidth * 0.5 + 30
            columnWidth = columnWidth * 2
        End If
        With rs
            If .RecordCount > 0 Then
                i = 0
                Do While Not .EOF
                    ypos = vertPos
                    FontGr 8
                    printObj.CurrentX = xpos + 50
                    printObj.CurrentY = ypos + printObj.TextHeight("99") * 0.5
                    'wedstrijd nr
                    printObj.Print Format(!wedNum, "0"); ":";
                    FontGr 12
                    printObj.CurrentX = xpos + printObj.TextWidth("00:") + 30
                    printObj.CurrentY = ypos
                    FontGr 10
                    printObj.Print !code1; ":";
                    FontGr 12
                    printObj.DrawWidth = 1
                    printObj.Line (xpos + printObj.TextWidth("00:"), ypos)-(xpos + columnWidth - 50, ypos + printObj.TextHeight("W")), vbBlack, B
                    ypos = printObj.CurrentY
                    printObj.CurrentX = xpos + printObj.TextWidth("00:") + 30
                    FontGr 10
                    printObj.Print !code2; ":";
                    FontGr 12
                    printObj.Line (xpos + printObj.TextWidth("00:"), ypos)-(xpos + columnWidth - 50, ypos + printObj.TextHeight("W")), vbBlack, B
                    printObj.CurrentY = ypos
                    .MoveNext
                    i = i + 1
                    xpos = columnWidth * i
                    If xpos > printObj.ScaleWidth - columnWidth + 100 Then
                        FontGr 8
                        printObj.Print
                        printObj.Print
                        FontGr 12
                        vertPos = printObj.CurrentY
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
    printObj.Print
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
rs.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
yLinePos = printObj.CurrentY
iGrp = getKampInfo("groepen")
xLinePos = (printObj.ScaleWidth / iGrp) * (nr - 1)
xpos = xLinePos + 50
Do While Not rs.EOF
    vakPos(0, 0) = xpos + printObj.ScaleWidth / iGrp - printObj.TextHeight("W") - printObj.TextWidth("W")
    vakPos(0, 1) = printObj.CurrentY
    vakPos(1, 0) = vakPos(0, 0) + printObj.TextHeight("W")
    vakPos(1, 1) = vakPos(0, 1) + printObj.TextHeight("W")
    
    txt = rs!naam
    Do While xpos + printObj.TextWidth(txt) > vakPos(0, 0)
        txt = Left(txt, Len(txt) - 1)
    Loop
    printObj.CurrentX = xpos
    printObj.Print txt;
    printObj.FillStyle = vbFSTransparent
    printObj.FillColor = vbWhite
    printObj.DrawWidth = 1
    
    printObj.Line (vakPos(0, 0), vakPos(0, 1))-(vakPos(1, 0), vakPos(1, 1)), vbBlack, B
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
'printObj.CurrentY = yLinePos
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
Dim columnWidth As Integer
Dim columnNrom As Integer
Dim ypos As Integer
Dim curYpos As Integer
Dim horPos As Integer
Dim vertPos As Integer
Dim i As Integer
Dim vertLineYPos As Integer
Dim vertLineYPos2 As Integer
Dim topY As String
Dim savdat As Date
Dim vertLineEndPos As Integer
    sqlstr = "Select * from qryweds where ksid = " & kampID
    sqlstr = sqlstr & " ORDER BY datum,tijd,wednum"
    rs.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
    If rs.RecordCount = 0 Then
        rs.Close
        Exit Sub
    End If
    fontBas = 10
    FontGr fontBas + 2
    topY = printObj.CurrentY
    printObj.CurrentY = footerHeight - largeHeight
    ypos = printObj.CurrentY
    printObj.FillColor = &H808080
    printObj.FillStyle = vbFSSolid
    printObj.Line (0, ypos)-(printObj.ScaleWidth + 2 * printObj.ScaleLeft, footerHeight), vbBlack, B
    printObj.CurrentY = ypos + 30
    FontGr 16
    Vet True
    printObj.ForeColor = vbWhite
    iBKMode = SetBkMode(printObj.hdc, TRANSPARENT)
    Centreer "UITERLIJK INLEVEREN OP " & UCase(Format(getPoolInfo("eindinschr"), "dddd d mmmm yyyy"))
    printObj.ForeColor = vbBlack
    printObj.FillStyle = vbFSTransparent
    Vet False
    FontGr fontBas + 2
    printObj.CurrentY = topY
    columnNrom = 0
    columnWidth = printObj.ScaleWidth / 2 - printObj.TextWidth("w")
    printObj.FontName = "Times New Roman"
    FontGr 2
    printObj.Print
    FontGr fontBas + 2
    printObj.CurrentY = printObj.CurrentY + 20
    FontGr fontBas + 4
    Vet True
    printObj.Print "Uitleg"
    FontGr fontBas + 2
    Vet False
    printObj.Print "Vul hieronder voor alle wedstrijden jouw uitslagen in. ";
    Vet True
    printObj.Print "Ook daar waar de teams nog niet bekend zijn."
    Vet False
    printObj.Print "(Ook al heb je een ander team op die plaats dan kan je uitslag nog steeds goed zijn)"
    printObj.Print "De uitslag hoeft onderling niet te kloppen. ";
    printObj.Print "Je krijgt punten voor elk vak dat achteraf juist blijkt te zijn ingevuld."
    printObj.Print "Bij 'toto' vul je een 1 in voor winst linker team, een 2 voor winst rechter team en een 3 voor een gelijkspel"
    Vet True
    Centreer "Alle uitslagen, ook de toto, gelden na 90 minuten voetbal!"
    Vet False
    FontGr fontBas
    printObj.Print
    Centreer "(plus de eventuele blessuretijd)"
    printObj.Print
    FontGr fontBas + 4
    Vet True
    printObj.Print "Punten"
    Vet False
    FontGr fontBas + 2
    printObj.Print "Ruststand goed: ";
    Vet True
    printObj.Print getPntToek("ruststand goed"); "pnt, ";
    Vet False
    printObj.Print "Eindstand goed: ";
    Vet True
    printObj.Print getPntToek("eindstand goed"); "pnt, ";
    Vet False
    printObj.Print "Toto goed: ";
    Vet True
    printObj.Print getPntToek("toto goed"); "pnt.";
    Vet False
    If getPntToek("doelpunten op een dag") > 0 Then
        printObj.Print "Totaal aantal doelpunten op één dag goed: ";
        Vet True
        printObj.Print getPntToek("doelpunten op een dag"); " pnt"
        Vet False
    End If
    printObj.Print
    FontGr fontBas
    posDatum = 50
    posTijd = posDatum + printObj.TextWidth("MA 26-6") + 10
    posWednr = posTijd + printObj.TextWidth("00:000") + 10
    posWedOms = posWednr + printObj.TextWidth("199:")
    posRust = posWedOms + printObj.TextWidth("Nederland - Zwitserland")
    PosEind = posRust + printObj.TextWidth("123456")
    posToto = PosEind + printObj.TextWidth("123456")
    
    vertLineYPos = printObj.CurrentY
    FontGr fontBas
    printObj.Line (0, vertLineYPos - 20)-(columnWidth * 2, vertLineYPos - 20)
    printObj.CurrentY = vertLineYPos
    For i = 0 To 1
        printObj.CurrentX = posDatum + i * columnWidth
        printObj.Print " Datum";
        printObj.CurrentX = posTijd + i * columnWidth
        printObj.Print " tijd";
        printObj.CurrentX = posWednr + i * columnWidth
        printObj.Print " nr";
        printObj.CurrentX = posWedOms + i * columnWidth
        printObj.Print " Wedstrijd";
        printObj.CurrentX = posRust + i * columnWidth
        printObj.Print " rust";
        printObj.CurrentX = PosEind + i * columnWidth
        printObj.Print " eind";
        printObj.CurrentX = posToto + i * columnWidth
        printObj.Print " toto";
    Next
    printObj.Print
    printObj.Line (0, printObj.CurrentY)-(columnWidth * 2, printObj.CurrentY), 1
    vertLineYPos2 = printObj.CurrentY
    
    ypos = printObj.CurrentY
    
    With rs
        .MoveLast
        .MoveFirst
        
        Do While Not .EOF
            If (nz(!naam1, "")) > "" Then
                wedOms = !code1 & ":" & !naam1 & " - " & !code2 & ":" & !naam2
            Else
                wedOms = !code1 & " - " & !code2
            End If
            
            printObj.CurrentY = printObj.CurrentY + 40
            printObj.CurrentX = posWednr + columnNrom * columnWidth + (posWedOms - posWednr - printObj.TextWidth(Format(!wedNum, "0"))) / 2
            printObj.Print Format(!wedNum, "0");
            printObj.CurrentX = posDatum + columnNrom * columnWidth
            If savdat <> !Datum Then
                printObj.Print Format(!Datum, "ddd d-M"); " ";
                savdat = !Datum
            End If
            printObj.CurrentX = posTijd + columnNrom * columnWidth + (posWednr - posTijd - printObj.TextWidth(Format(!tijd, "HH:NN"))) / 2
            printObj.Print tijdFormat(!tijd); '  , "HH:NN");
            printObj.CurrentX = posWedOms + columnNrom * columnWidth + 30
            curYpos = printObj.CurrentY
            If (nz(!naam1, "")) > "" Then
                FontGr fontBas - 3
                printObj.CurrentY = curYpos + 20
                Do While printObj.TextWidth(wedOms) > posRust - posWedOms
                    wedOms = Left(wedOms, Len(wedOms) - 1)
                Loop
            Else
                FontGr fontBas
                printObj.CurrentY = curYpos
            End If
            printObj.Print wedOms;
            printObj.CurrentY = curYpos
            FontGr fontBas
            horPos = posRust + columnNrom * columnWidth
            vertPos = printObj.CurrentY - 20
            printObj.Line (horPos, vertPos)-(PosEind + columnNrom * columnWidth - 10, vertPos + printObj.TextHeight("W") + 50), , B
            printObj.CurrentX = posRust + (PosEind - posRust - printObj.TextWidth("-")) / 2 + columnNrom * columnWidth
            printObj.CurrentY = vertPos + 30
            printObj.Print "-";
            horPos = PosEind + columnNrom * columnWidth + 10
            printObj.Line (horPos, vertPos)-(posToto + columnNrom * columnWidth - 10, vertPos + printObj.TextHeight("W") + 50), , B
            printObj.CurrentX = PosEind + (posToto - PosEind - printObj.TextWidth("-")) / 2 + columnNrom * columnWidth
            printObj.CurrentY = vertPos + 30
            printObj.Print "-";
            horPos = posToto + columnNrom * columnWidth + 10
            printObj.Line (horPos, vertPos)-(columnWidth * (columnNrom + 1) - printObj.TextWidth("0"), vertPos + printObj.TextHeight("W") + 50), , B
            printObj.CurrentX = PosEind + (posToto - PosEind - printObj.TextWidth("-")) / 2
            printObj.CurrentY = vertPos
            
            FontGr 14
            printObj.Print
            FontGr fontBas
            printObj.Line (0, printObj.CurrentY)-(columnWidth * 2, printObj.CurrentY), 1

            .MoveNext
            If (.AbsolutePosition - 1) = Int(rs.RecordCount / 2 + 0.5) Then
                columnNrom = 1
                vertLineEndPos = printObj.CurrentY
                printObj.CurrentY = ypos
            End If
        Loop
        .Close
    End With
    Set rs = Nothing
    For i = 0 To 1
        printObj.Line (0 + columnWidth * i, vertLineYPos - 10)-(0 + columnWidth * i, vertLineEndPos)
        printObj.Line (posWednr + columnWidth * i - 10, vertLineYPos2)-(posWednr + columnWidth * i - 10, vertLineEndPos)
        printObj.Line (posTijd + columnWidth * i, vertLineYPos2)-(posTijd + columnWidth * i, vertLineEndPos)
        printObj.Line (posWedOms + columnWidth * i - 10, vertLineYPos2)-(posWedOms + columnWidth * i - 10, vertLineEndPos)
    Next
    printObj.Line (columnWidth - 50, vertLineYPos - 10)-(columnWidth - 50, vertLineEndPos)
    printObj.Line (columnWidth * 2, vertLineYPos - 10)-(columnWidth * 2, vertLineEndPos)
End Sub


Private Sub PrijsAfdr(wat As String, eind As Boolean)
Dim aant As Integer
Dim i As Integer
End Sub

Private Sub Centreer(Tekst$)
    printObj.CurrentX = (printObj.ScaleWidth - printObj.TextWidth(Trim$(Tekst$))) \ 2
    printObj.Print Tekst$;
End Sub

Function sqlDeelnems(poule As Long) As String
Dim sqlstr As String
    sqlstr = "Select * from pooldeelnems"
    sqlstr = sqlstr & " WHERE PoolID = " & poule
    sqlstr = sqlstr & " ORDER BY bijnaam "
    sqlDeelnems = sqlstr
End Function

Private Sub printFavourites()
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

deelnAant = GetDeelnemAant(poolID)
headerText = GetOrgNaam(poolID) & " " & getKampInfo("toernooi") & " voetbalpool" & " - Favorieten" & " (" & GetDeelnemAant(poolID) & " deelnemers)"
'printObj.Line (0, printObj.CurrentY)-(printObj.ScaleWidth, printObj.CurrentY)
header2$ = "Groepstanden"
InitPage False, False
'intro
yStart = printObj.CurrentY

'groepen
fntGr = printObj.Font.Size
sqlstr = "Select groepen from ks WHERE id = " & kampID
rs.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
aantgroep = rs!groepen
rs.Close
printObj.CurrentX = printObj.TextWidth("12345678901234567890123456")
For i = 1 To 4
    printObj.CurrentX = printObj.CurrentX - printObj.TextWidth(Format(i, "0") & "e pl")
    printObj.Print Format(i, "0"); "e pl";
    col(i) = printObj.CurrentX - 50
    printObj.CurrentX = printObj.CurrentX + printObj.TextWidth("123456")
Next
printObj.CurrentX = printObj.ScaleWidth / 2 + printObj.TextWidth("12345678901234567890123456")
For i = 1 To 4
    printObj.CurrentX = printObj.CurrentX - printObj.TextWidth(Format(i, "0") & "e pl")
    printObj.Print Format(i, "0"); "e pl";
    printObj.CurrentX = printObj.CurrentX + printObj.TextWidth("123456")
Next
printObj.CurrentX = 0
printObj.Print
xpos = 0
savy = printObj.CurrentY
For i = 1 To aantgroep
    If i = aantgroep / 2 + 1 Then
        xpos = printObj.ScaleWidth / 2
        printObj.CurrentY = savy
    End If
    sqlstr = "Select * from groepsindeling where ksid = " & kampID
    sqlstr = sqlstr & " AND groep = '" & Chr(i + 64) & "'"
    rs.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
    rs.MoveFirst
    printObj.CurrentX = xpos
    printObj.Print "Groep " & rs!groep; ": ";
    savX = printObj.CurrentX
    Do While Not rs.EOF
        printObj.CurrentX = savX
        printObj.Print GetTeam(rs!team); " ";
        printObj.CurrentX = printObj.TextWidth("12345678901234567890")
        For j = 1 To 4
            aant = getAantalGrpVoorsp(j, rs!team)
            FontGr 9
            printObj.CurrentY = printObj.CurrentY + 30
            printObj.CurrentX = xpos + col(j) - printObj.TextWidth(Format(aant / deelnAant, "0.0%"))
'            printObj.Print aant;
'            FontGr 8
            printObj.Print Format(aant / deelnAant, "0.0%");
            printObj.CurrentY = printObj.CurrentY - 30
            FontGr CInt(fntGr)
            'If j < 4 Then printObj.Print ", ";
        Next
        printObj.Print
        rs.MoveNext
    Loop
    rs.Close
Next
savy = printObj.CurrentY
On Error Resume Next
printObj.Line (0, yStart)-(printObj.ScaleWidth - 50, savy), , B
On Error GoTo 0
maxY = savy
'achtste finales
i = getPntToek("achtste finaleplaats") + getPntToek("achtste finalepositie")
If i > 0 Then
    Fav_Finals 5, 4, "Achtste finales"
    savy = printObj.CurrentY
End If
printObj.CurrentY = savy
'kwart finales
i = getPntToek("kwart finaleplaats") + getPntToek("kwart finalepositie")
If i > 0 Then
    Fav_Finals 2, 4, "Kwart finales"
    savy = printObj.CurrentY
End If
printObj.CurrentY = savy
'halve finales
i = getPntToek("halve finaleplaats") + getPntToek("halve finalepositie")
If i > 0 Then
    Fav_Finals 3, 4, "Halve finales"
    savy = printObj.CurrentY
    maxY = savy
End If
printObj.CurrentY = savy
'kleine finale
i = getPntToek("kleine finaleplaats") + getPntToek("kleine finalepositie")
If i > 0 Then
    bewYPos = printObj.CurrentY
    Fav_Finals 7, 4, "Kleine finale"
    savy = maxY
    'maxY = savy
    savX = 3
Else
    bewYPos = printObj.CurrentY
    savX = 1
End If

'finale
i = getPntToek("finaleplaats") + getPntToek("finalepositie")
If i > 0 Then
    Fav_Finals 4, 4, "Finale", savy, savX
    If savX = 3 Then
        savX = 1
        savy = printObj.CurrentY
    Else
        savy = bewYPos
        savX = 3
    End If
'    savy = printObj.CurrentY
    maxY = savy
End If
printObj.CurrentY = savy
Fav_Eindstand savy, savX
Fav_Topscorers
Set rs = Nothing
printObj.Print
printObj.Print
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
    cols(i) = Int(printObj.ScaleWidth / 4) * (i - 1)
Next
cols(5) = printObj.ScaleWidth - 10
sqlstr = "SELECT personen.rnaam, Count(voorspelling_ts.deelnem) AS aantal"
sqlstr = sqlstr & " FROM voorspelling_ts LEFT JOIN personen ON voorspelling_ts.ts = personen.ID"
sqlstr = sqlstr & " WHERE voorspelling_ts.deelnem In (select deelnemid from pooldeelnems where poolid= " & poolID
sqlstr = sqlstr & " ) GROUP BY personen.rnaam, voorspelling_ts.ts"
sqlstr = sqlstr & " ORDER BY Count(voorspelling_ts.deelnem) DESC, personen.rnaam "
rs.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
If rs.RecordCount > 0 Then
    rs.MoveLast
End If
aant = rs.RecordCount
i = 1
j = 0

printObj.CurrentX = favXpos
If favYpos > footerHeight - Int(aant / 4) * printObj.TextHeight("tekst") - 120 Then
  header2$ = "Topscorers"
  DoNewPage False, False, 0
  favYpos = printObj.CurrentY
Else
  printObj.CurrentY = favYpos
  header2tekst "Topscorers", False, False, favYpos, 0
End If

savy = printObj.CurrentY
rs.MoveFirst

Do While Not rs.EOF
    printObj.CurrentX = cols(i)
    If nz(rs!rnaam, "") > "" Then
        printObj.Print rs!rnaam;
    Else
        printObj.Print "Niet ingevuld";
    End If
    printObj.CurrentX = cols(i + 1) - 500 - printObj.TextWidth(rs!Aantal)
    printObj.Print rs!Aantal
    j = j + 1
    rs.MoveNext
    If printObj.CurrentY > favYpos Then
        favYpos = printObj.CurrentY
    End If
    If j > Int(aant / 4) - 1 Then
        i = i + 1
        j = 0
        printObj.CurrentY = savy
    End If
Loop
rs.Close
Set rs = Nothing
printObj.Line (cols(1), savy)-(cols(5) - 50, favYpos), , B

End Sub

Function GetRijAant(wedNum As Integer, team)
'om te bepalen of we naar een nieuw pagina moeten in de favorieten afdruk
Dim sqlstr As String
sqlstr = "SELECT wed, " & team
sqlstr = sqlstr & " From voorspelling_finales"
sqlstr = sqlstr & " WHERE deelnem In (select deelnemid from pooldeelnems where poolid =" & poolID
sqlstr = sqlstr & " ) GROUP BY wed, " & team
sqlstr = sqlstr & " HAVING wed =" & wedNum
sqlstr = sqlstr & " AND " & team & " >0"
Dim rs As New ADODB.Recordset
rs.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
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
    ypos = printObj.CurrentY
    fntGr = printObj.Font.Size
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Vet True
        printObj.CurrentX = col
        printObj.Print Plaats
        Vet False
        Do While Not rs.EOF
            printObj.CurrentX = col + 50
            If nz(rs(veld), 0) = 0 Then
                printObj.Print "Niet ingevuld";
            Else
                printObj.Print GetTeam(rs(veld));
            End If
            printObj.CurrentX = col + printObj.TextWidth("123456789012345") - printObj.TextWidth(rs!Aantal)
            printObj.Print rs!Aantal;
            FontGr fntGr - 3
            printObj.CurrentY = printObj.CurrentY + 30
            printObj.Print "(" & Format(rs!Aantal / GetDeelnemAant(poolID), "0.0%") & ")"
            printObj.CurrentY = printObj.CurrentY - 30
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
    cols(i) = Int((printObj.ScaleWidth / 4) * (i - 1))
Next

cols(5) = printObj.ScaleWidth - 20

    startY = savy

    sqlstr = "SELECT kampioen, Count(pooldeelnems.deelnemID) AS aantal"
    sqlstr = sqlstr & " From pooldeelnems"
    sqlstr = sqlstr & " WHERE poolid = " & poolID
    sqlstr = sqlstr & " GROUP BY kampioen"
    sqlstr = sqlstr & " ORDER BY count(deelnemID) desc"
    rs1.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
    sqlstr = "SELECT pltwee, Count(pooldeelnems.deelnemID) AS aantal"
    sqlstr = sqlstr & " From pooldeelnems"
    sqlstr = sqlstr & " WHERE poolid = " & poolID
    sqlstr = sqlstr & " GROUP BY pltwee"
    sqlstr = sqlstr & " ORDER BY count(deelnemID) desc"
    rs2.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
    sqlstr = "SELECT pldrie, Count(pooldeelnems.deelnemID) AS aantal"
    sqlstr = sqlstr & " From pooldeelnems"
    sqlstr = sqlstr & " WHERE poolid = " & poolID
    sqlstr = sqlstr & " GROUP BY pldrie"
    sqlstr = sqlstr & " ORDER BY count(deelnemID) desc"
    rs3.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
    sqlstr = "SELECT plvier, Count(pooldeelnems.deelnemID) AS aantal"
    sqlstr = sqlstr & " From pooldeelnems"
    sqlstr = sqlstr & " WHERE poolid = " & poolID
    sqlstr = sqlstr & " GROUP BY plvier"
    sqlstr = sqlstr & " ORDER BY count(deelnemID) desc"
    rs4.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
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
    savFntgr = printObj.FontSize
    printObj.FontSize = savFntgr - 3
    maxY = maxaant * printObj.TextHeight("Q") + savy
    printObj.FontSize = savFntgr
    maxY = maxY + printObj.TextHeight("Q") + 50
    If maxY > footerHeight - 465 Then
        header2$ = "Favorieten einduitslag"
        DoNewPage False, False, aantFav
        'maxY = printObj.CurrentY
        savy = printObj.CurrentY
        startY = savy
        savFntgr = printObj.FontSize
        printObj.FontSize = savFntgr - 3
        maxY = maxaant * printObj.TextHeight("Q") + savy
        printObj.FontSize = savFntgr
        maxY = maxY + printObj.TextHeight("Q") + 50
    Else
      If savX2 = 3 Then
        header2tekst "Favorieten einduitslag", False, False, savy, savX2 + 1
      Else
        header2tekst "Favorieten einduitslag", False, False, savy, savX2 - 1 ' 0 centreert tussenheader2
      End If
      savy = printObj.CurrentY
      startY = savy
      savFntgr = printObj.FontSize
      printObj.FontSize = savFntgr - 3
      maxY = maxaant * printObj.TextHeight("Q") + savy
      printObj.FontSize = savFntgr
      maxY = maxY + printObj.TextHeight("Q") + 50
    End If
    If getPntToek("1e plaats(Kampioen)") Then
        printObj.CurrentY = savy
        PrintEindStandFav "kampioen", cols(savX2) + 10, rs1, "kampioen"
        printObj.Line (cols(savX2), startY)-(cols(savX2 + 1) - 50, maxY), , B
    End If
    If getPntToek("2e plaats") Then
        printObj.CurrentY = savy
        PrintEindStandFav "2e plaats", cols(savX2 + 1) + 10, rs2, "plTwee"
        printObj.Line (cols(savX2 + 1), startY)-(cols(savX2 + 2) - 50, maxY), , B
    End If
    If getPntToek("3e plaats") Then
        printObj.CurrentY = savy
        PrintEindStandFav "3e plaats", printObj.ScaleWidth / 2 + 10, rs3, "pldrie"
        printObj.Line (cols(3), startY)-(cols(4) - 50, maxY), , B
    End If
    If getPntToek("4e plaats") Then
        printObj.CurrentY = savy
        PrintEindStandFav "4e plaats", (printObj.ScaleWidth / 4) * 3 + 10, rs4, "plvier"
        printObj.Line (cols(4), startY)-(cols(5) - 50, maxY), , B
    End If
    favYpos = maxY
    favXpos = 0
End Sub
Sub Fav_Finals(wedtype As Integer, cols As Integer, header2txt As String, Optional bewaarYpos As Integer, Optional posX As Integer)
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
        col(i) = (i - 1) * printObj.ScaleWidth / cols
    Next
    col(cols + 1) = printObj.ScaleWidth
    savy = printObj.CurrentY
    sqlstr = "Select * from qryWeds where  ksid = " & kampID
    sqlstr = sqlstr & " and wedtype = " & wedtype
    sqlstr = sqlstr & " ORDER BY wednum"
    rs.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
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
        If startY + ttlRows * TextHeight("Q") > footerHeight - 465 And wedtype <> 4 Then '(465 = hoogte van het tussenheader2je)
            header2$ = header2txt
            If wedtype = 7 Then
                DoNewPage False, False, 2
                maxY = printObj.CurrentY
                savy = maxY
                startY = 480
                nwPag = True
            Else
                DoNewPage False, False
                maxY = printObj.CurrentY
                savy = maxY
                startY = savy
                nwPag = False
            End If
        Else
            If wedtype = klFinale Then
                finYpos = printObj.CurrentY
                header2tekst header2txt, False, False, maxY, 2
            ElseIf wedtype = Finale Then
                If getPntToek("kleine finaleplaats") + getPntToek("kleine finalepositie") > 0 Then
                    If nwPag Then
                        header2tekst header2txt, False, False, 480, 4
                    Else
                        header2tekst header2txt, False, False, finYpos, 4
                    End If
                Else
                    header2tekst header2txt, False, False, bewaarYpos, 2
                End If
            Else
                header2tekst header2txt, False, False, maxY
            End If
            savy = printObj.CurrentY
            startY = savy
        End If
        
        i = 1
        If wedtype = Finale Then
            i = posX
        End If
        'If wedtype = 7 Then Stop
        Do While Not rs.EOF
            If i <= cols Then
                printObj.CurrentY = savy
            End If
            fav_finalTeams "t1", "code1", rs, col(i)
            If maxY < printObj.CurrentY Then maxY = printObj.CurrentY
            i = i + 1
            If i <= cols Then
                printObj.CurrentY = savy
            End If
            fav_finalTeams "t2", "code2", rs, col(i)
            If maxY < printObj.CurrentY Then maxY = printObj.CurrentY
            i = i + 1
            
            If wedtype = 7 And maxY < printObj.CurrentY Then
                maxY = printObj.CurrentY
            ElseIf wedtype = 4 Then
                If printObj.CurrentY > maxY Then
                    maxY = printObj.CurrentY
                End If
            End If
            maxY = maxY + 50
            If i = 5 Then
                printObj.Line (col(1), startY)-(col(3) - 50, maxY), , B
                printObj.Line (col(3), startY)-(col(5) - 50, maxY), , B
            End If
            If posX = 1 And i = 3 Then
                printObj.Line (col(1), startY)-(col(3) - 50, maxY), , B
            End If
            
            rs.MoveNext
            If i > cols Then
                i = 1
                printObj.CurrentY = maxY + 50
                savy = printObj.CurrentY
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
    aantpos = printObj.TextWidth("NIET INGEVULD  1")
    sqlstr = "SELECT wed, " & team & ", Count(wed) AS ttl From voorspelling_finales"
    sqlstr = sqlstr & " WHERE deelnem In (select deelnemid from pooldeelnems where poolid =" & poolID
    sqlstr = sqlstr & " ) GROUP BY wed, " & team
    sqlstr = sqlstr & " HAVING wed=" & rs!wedNum
    sqlstr = sqlstr & " AND " & team & " > 0"
    sqlstr = sqlstr & " ORDER BY count(wed) desc"
    rs1.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
    printObj.CurrentX = col
    printObj.Print rs(cod) & ": ";
    savX = printObj.CurrentX
    fntGr = printObj.Font.Size
    Do While Not rs1.EOF
        printObj.CurrentX = savX
        If nz(rs1(team), "") = "" Then
            printObj.Print "Niet ingevuld";
        Else
            printObj.Print GetTeam(rs1(team));
        End If
        printObj.CurrentX = col + aantpos - printObj.TextWidth(rs1!ttl)
        printObj.Print rs1!ttl;
        FontGr fntGr - 3
        printObj.CurrentY = printObj.CurrentY + 30
        printObj.Print "(" & Format(rs1!ttl / GetDeelnemAant(poolID), "0.0%") & ")"
        FontGr fntGr
        printObj.CurrentY = printObj.CurrentY - 30
        If maxY < printObj.CurrentY Then maxY = printObj.CurrentY
        rs1.MoveNext
    Loop
    rs1.Close
End Sub

Private Sub printCompetitors()
Dim Dezedeeln As Integer
Dim tkst$
Dim tmpnaam$
Dim columnNromAant As Integer
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
Dim wedcolumnNr As Integer
Dim Helft As Integer
Dim oldhelft As Integer
Dim heeft8stFin As Boolean
Dim savdat As Date
Dim savWedType As Integer
Dim kaderPos As Integer
Dim deelnPag As Integer
Dim grpWedsAant As Integer
Dim nwcolumnNr As Boolean
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
    If printObj.ScaleHeight <> Printer.ScaleHeight Then
        Helft = Helft + printObj.TextHeight("W") * 2
    End If
    grpWedsAant = AantGrpWeds()
    rotate.Angle = 0
    wedHoog = 9
    NaamHoog = 11
    rsDeelnem.Open sqlDeelnems(poolID), dbConn, adOpenStatic, adLockReadOnly
    
    If rsDeelnem.RecordCount = 0 Then
        MsgBox "Geen deelnemers in deze pool", vbQuestion + vbOKOnly, "Deelnemers afdrukken"
        Exit Sub
    End If
    columnNromAant = 1
    horPos% = 20
    headerText = GetOrgNaam(poolID) & " " & getKampInfo("toernooi") & " voetbalpool"
    tkst$ = "Deelnemers en Voorspellingen"
    header2$ = tkst$
    
    InitPage True, False
    FontGr NaamHoog
    printObj.CurrentY = printObj.CurrentY - 50
    headerHeight = printObj.CurrentY
    TopMarg = printObj.CurrentY
    AantalOpPapier = 2
    If grpWedsAant <= 24 Then
        AantalOpPapier = 3
    End If
    Helft = (footerHeight - TopMarg) / AantalOpPapier
'    Helft = printObj.ScaleHeight / AantalOpPapier + 100 'printObj.CurrentY
    FontGr wedHoog
    'Debug.Print printObj.FontSize, Printer.FontSize * afdrRatio
    lineHeight% = printObj.TextHeight("x") '* afdrRatio
    FontGr NaamHoog
    naamHeight = printObj.TextHeight("x") '* afdrRatio
    If getKampInfo("groepen") > 4 Then
        columnNromAant = getKampInfo("groepen")
    Else
        columnNromAant = 8
    End If
    
    columnWidth = Int((printObj.ScaleWidth / columnNromAant) - 50)
    printObj.FillStyle = vbFSTransparent
    rsDeelnem.MoveFirst
    FontGr 8
    posDatum = 50
    posWedOms = posDatum + printObj.TextWidth("99-99:")
    posRust = posWedOms + printObj.TextWidth("WWW-WWW")
    PosEind = posRust + printObj.TextWidth("11-11")
    posToto = PosEind + printObj.TextWidth("11-11")
    posPnt = posToto + printObj.TextWidth("99")
    FontGr 12
    deelnPag = 0
    Do While Not rsDeelnem.EOF
        If Me.lstDeelnems.Selected(rsDeelnem.AbsolutePosition - 1) Or Me.Option3 = True Then
            showInfo True, "Afdrukken deelnemers", rsDeelnem!bijnaam, "Record " & rsDeelnem.AbsolutePosition & "/" & rsDeelnem.RecordCount
            
            If deelnPag = 0 Then
                printObj.CurrentY = TopMarg
            Else
                printObj.CurrentY = deelnPag * (Helft) + TopMarg
            End If
            LineYPos = printObj.CurrentY
            printObj.CurrentX = 30
            Vet True
            FontGr NaamHoog + 6
            printObj.Print
            wedYpos = printObj.CurrentY
            
            printObj.Line (0, LineYPos)-(printObj.ScaleWidth - 10, wedYpos), &H127419, BF
            printObj.CurrentY = LineYPos
            printObj.ForeColor = vbWhite
            iBKMode = SetBkMode(printObj.hdc, TRANSPARENT)
            printObj.CurrentX = 30
            printObj.Print rsDeelnem!bijnaam;
            ttlPosX = printObj.ScaleWidth
            ttlPosY = printObj.CurrentY
            printObj.Print
            Vet False
            printObj.CurrentX = 50
            printObj.ForeColor = 1
            'groepswedstrijden
            sqlstr = "Select * from qryDeelnWeds Where deelnem = " & rsDeelnem!deelnemID
            rsDeelnemWeds.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
            FontGr 10
            Vet True
            printObj.ForeColor = vbBlue
            printObj.Print "Groepswedstrijden";
            grpwedsTtlPosX = printObj.CurrentX
            grpwedsTtlPosY = printObj.CurrentY
            printObj.CurrentX = printObj.ScaleWidth * 0.75 + 50
            printObj.Print "Finales";
            printObj.ForeColor = 1
            Vet False
            printObj.FontItalic = True
            FontGr 8
            'For i = 1 To 4
                'printObj.CurrentX = printObj.ScaleWidth / 4 * i - printObj.TextWidth("pnt") - 50
                'printObj.Print "pnt";
            'Next
            printObj.FontItalic = False
            FontGr 10
            printObj.Print
            printObj.Line (0, wedYpos - 10)-(printObj.ScaleWidth - 10, printObj.CurrentY + 10), , B
            LineYPos = printObj.CurrentY + 10
            printObj.CurrentY = LineYPos
            FontGr 8
            LineXpos = 0
            With rsDeelnemWeds
'                showInfo True, "Afdrukken deelnemers", rsDeelnem!bijnaam, "Record " & rsDeelnem.AbsolutePosition  & "/" & rsDeelnem.RecordCount, "Wedstrijden"
                k = 0
                If .RecordCount > 0 Then
                    .MoveLast
                    .MoveFirst
                    wedcolumnNr = 1
                    Do While Not .EOF
                        printObj.CurrentX = LineXpos + posWedOms - printObj.TextWidth(Format(!Datum, "d-m") & ":") - 50
                        If savdat <> !Datum Or printObj.CurrentY = LineYPos Then
                            printObj.Print Format(!Datum, "d-m"); ":";
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
                        printObj.CurrentX = LineXpos + posWedOms
                        printObj.Print pr;
                        printObj.CurrentX = LineXpos + posRust
                        printObj.Print !r1; "-"; !r2;
                        printObj.CurrentX = PosEind + LineXpos
                        printObj.Print !e1; "-"; !e2;
                        printObj.CurrentX = LineXpos + posToto
                        printObj.Print !toto;
                        printObj.Print
                        If newlinepos < printObj.CurrentY Then newlinepos = printObj.CurrentY
                        rsDeelnemWeds.MoveNext
                        If grpWedsAant < 25 Then
                            nwcolumnNr = (.AbsolutePosition - 1) Mod (grpWedsAant / 3) = 0 '= Int(grpWedsAant / 2) Or .AbsolutePosition = grpWedsAant
                        Else
                            nwcolumnNr = (.AbsolutePosition - 1) Mod 16 = 0
                        End If
                        If nwcolumnNr Then
                            printObj.CurrentY = LineYPos
                            k = k + 1
                            If (.AbsolutePosition - 1) = grpWedsAant Then k = 3
                            LineXpos = (printObj.ScaleWidth / 4) * k
                        End If
                    Loop
                End If
                .Close
            End With
            printObj.Line (0, wedYpos)-(0, newlinepos)
            For i = 1 To 4
                printObj.Line (printObj.ScaleWidth / 4 * i - 20, LineYPos)-(printObj.ScaleWidth / 4 * i - 20, newlinepos)
                printObj.Line (printObj.ScaleWidth / 4 * (i - 1) + posRust - 20, LineYPos)-(printObj.ScaleWidth / 4 * (i - 1) + posRust - 20, newlinepos)
                printObj.Line (printObj.ScaleWidth / 4 * (i - 1) + PosEind - 20, LineYPos)-(printObj.ScaleWidth / 4 * (i - 1) + PosEind - 20, newlinepos)
                printObj.Line (printObj.ScaleWidth / 4 * (i - 1) + posToto - 20, LineYPos)-(printObj.ScaleWidth / 4 * (i - 1) + posToto - 20, newlinepos)
                printObj.Line (printObj.ScaleWidth / 4 * (i - 1) + posPnt - 20, LineYPos)-(printObj.ScaleWidth / 4 * (i - 1) + posPnt - 20, newlinepos)
            Next
            FontGr 10
            'groepstanden
'            showInfo True, "Afdrukken deelnemers", rsDeelnem!bijnaam, "Record " & rsDeelnem.AbsolutePosition + 1 & "/" & rsDeelnem.RecordCount, "Groepstanden"
            printObj.Line (0, newlinepos)-(printObj.ScaleWidth, newlinepos)
            printObj.Line (0, newlinepos)-(printObj.ScaleWidth - 10, newlinepos + printObj.TextHeight("Gr") + 10), , B
            printObj.CurrentY = newlinepos + 10
            printObj.CurrentX = 50
            Vet True
            printObj.ForeColor = vbBlue
            printObj.Print "Groepstanden"
            printObj.ForeColor = 1
            Vet False
            LineYPos = printObj.CurrentY
            columnWidth = Int((printObj.ScaleWidth / columnNromAant)) - 1
            FontGr 10
            sqlstr = "Select * from voorspelling_groepstand Where deelnem = " & rsDeelnem!deelnemID
            sqlstr = sqlstr & " ORDER BY groep"
            rsDeelnGroepen.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
            'LineYPos = printObj.CurrentY - 10
            k = 0
            printObj.CurrentX = 50
            Do While Not rsDeelnGroepen.EOF
                printObj.FontUnderline = True
                printObj.ForeColor = &H4000&
                printObj.Print "Groep " & rsDeelnGroepen!groep
                printObj.ForeColor = 1
                printObj.FontUnderline = False
                
'                printObj.CurrentX = printObj.CurrentX + printObj.TextWidth("|00")
                For i = 1 To 4
                    printObj.CurrentX = columnWidth * k
                    pr = GetTeam(rsDeelnGroepen("pos" & Format(i, "0")))
                    If pr = "" Then pr = "?"
                    printObj.Print i; ":"; pr
                    If newlinepos < printObj.CurrentY Then newlinepos = printObj.CurrentY
                Next
                k = k + 1
                printObj.Line (columnWidth * (k - 1), LineYPos)-(columnWidth * (k), newlinepos), , B
                printObj.CurrentX = columnWidth * k + 100
                printObj.CurrentY = LineYPos
                rsDeelnGroepen.MoveNext
            Loop
            
            rsDeelnGroepen.Close
            If grpWedsAant > 24 Then
                printObj.CurrentX = grpPntPosX
                printObj.CurrentY = newlinepos
            Else
                printObj.CurrentX = columnWidth * k
            End If
            'finales
            newlinepos = printObj.CurrentY
            printObj.Line (printObj.CurrentX, newlinepos)-(printObj.ScaleWidth, newlinepos)
            printObj.CurrentY = newlinepos
            LineXpos = 0
            LineYPos = printObj.CurrentY
            sqlstr = "Select * from qrydeelnemfinales WHERE deelnem=" & rsDeelnem!deelnemID
            sqlstr = sqlstr & " AND wedtype = " & AchtsteFinale
            sqlstr = sqlstr & " AND ksid= " & kampID
            If rsDeelnFinales.State = adStateOpen Then
                rsDeelnFinales.Close
            End If
            rsDeelnFinales.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
            If rsDeelnFinales.RecordCount > 0 Then
                With rsDeelnFinales
                    printObj.CurrentX = LineXpos + 20
                    Vet True
                    printObj.ForeColor = vbBlue
                    printObj.Print "Achtste finales"
                    printObj.ForeColor = 1
                    Vet False
                    Do While Not .EOF
                        printObj.CurrentX = LineXpos + 50
                        prntReg = Format(!wed, "0") & ": " & !tm1 & " - " & !tm2
                        Do While printObj.TextWidth(prntReg) > printObj.ScaleWidth / 5 - 100
                          prntReg = Left(prntReg, Len(prntReg) - 1)
                        Loop
                        printObj.Print prntReg;
                        printObj.Print
                        If .AbsolutePosition = 4 Then
                            If LineYPos < printObj.CurrentY Then LineYPos = printObj.CurrentY
                            LineXpos = LineXpos + printObj.ScaleWidth / 5
                            printObj.CurrentY = newlinepos
                            printObj.Print
                        End If
                        .MoveNext
                    Loop
                    .Close
                End With
                printObj.CurrentY = newlinepos
            End If
            If grpWedsAant > 24 Then
                LineXpos = printObj.ScaleWidth / 5 * 2
            Else
                LineXpos = printObj.ScaleWidth / 2
            End If
            sqlstr = "Select distinct * from qrydeelnemfinales WHERE deelnem=" & rsDeelnem!deelnemID
            sqlstr = sqlstr & " AND wedtype = " & KwartFinale
            sqlstr = sqlstr & " AND ksid= " & kampID
            If rsDeelnFinales.State <> 0 Then rsDeelnFinales.Close
            rsDeelnFinales.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
            If rsDeelnFinales.RecordCount > 0 Then
                With rsDeelnFinales
                    printObj.CurrentX = LineXpos + 50
                    Vet True
                    printObj.ForeColor = vbBlue
                    printObj.Print "Kwart finales"
                    printObj.ForeColor = 1
                    Vet False
                    Do While Not .EOF
                        printObj.CurrentX = LineXpos + 50
                        prntReg = Format(!wed, "0") & ": " & !tm1 & " - " & !tm2
                        Do While printObj.TextWidth(prntReg) > printObj.ScaleWidth / 5 - 100
                          prntReg = Left(prntReg, Len(prntReg) - 1)
                        Loop
                        printObj.Print prntReg;
                        printObj.Print
                        If LineYPos < printObj.CurrentY Then LineYPos = printObj.CurrentY
                        .MoveNext
                    Loop
                    .Close
                End With
                printObj.CurrentY = newlinepos
            End If
            If grpWedsAant > 24 Then
                LineXpos = printObj.ScaleWidth / 5 * 3
            Else
                LineXpos = printObj.ScaleWidth / 4 * 3
            End If
            sqlstr = "Select DISTINCT * from qrydeelnemfinales WHERE deelnem=" & rsDeelnem!deelnemID
            sqlstr = sqlstr & " AND wedtype = " & HalveFinale
            sqlstr = sqlstr & " AND ksid= " & kampID
            If rsDeelnFinales.State = adStateOpen Then
                rsDeelnFinales.Close
            End If
            rsDeelnFinales.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
            If rsDeelnFinales.RecordCount > 0 Then
                With rsDeelnFinales
                    printObj.CurrentX = LineXpos + 50
                    Vet True
                    printObj.ForeColor = vbBlue
                    printObj.Print "Halve finales"
                    printObj.ForeColor = 1
                    If LineYPos < printObj.CurrentY Then LineYPos = printObj.CurrentY
                    Vet False
                   ' printObj.Print
                    Do While Not .EOF
                        printObj.CurrentX = LineXpos + 50
                        prntReg = Format(!wed, "0") & ": " & !tm1 & " - " & !tm2
                        Do While printObj.TextWidth(prntReg) > printObj.ScaleWidth / 5 - 100
                          prntReg = Left(prntReg, Len(prntReg) - 1)
                        Loop
                        printObj.Print prntReg; ' Format(!wed, "0"); ": "; !tm1; " - "; !tm2;
                        printObj.Print
                        .MoveNext
                    Loop
                    .Close
                End With
                If grpWedsAant > 24 Then
                    printObj.CurrentY = newlinepos
                End If
            End If
            If grpWedsAant > 24 Then
                LineXpos = printObj.ScaleWidth / 5 * 4
            Else
                LineXpos = printObj.ScaleWidth / 4 * 3
            End If
            sqlstr = "Select * from qrydeelnemfinales WHERE deelnem=" & rsDeelnem!deelnemID
            sqlstr = sqlstr & " AND wedtype = " & klFinale
            sqlstr = sqlstr & " AND ksid= " & kampID
            If rsDeelnFinales.State = adStateOpen Then
                rsDeelnFinales.Close
            End If
            rsDeelnFinales.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
            
            If rsDeelnFinales.RecordCount > 0 Then
                With rsDeelnFinales
                    printObj.CurrentX = LineXpos + 50
                    Vet True
                    printObj.ForeColor = vbBlue
                    printObj.Print "3e plaats"
                    printObj.ForeColor = 1
                    Vet False
                    Do While Not .EOF
                        printObj.CurrentX = LineXpos + 50
                        prntReg = Format(!wed, "0") & ": " & !tm1 & " - " & !tm2
                        Do While printObj.TextWidth(prntReg) > printObj.ScaleWidth / 5 - 100
                          prntReg = Left(prntReg, Len(prntReg) - 1)
                        Loop
                        printObj.Print prntReg;
                        printObj.Print
                        If LineYPos < printObj.CurrentY Then LineYPos = printObj.CurrentY
                        .MoveNext
                    Loop
                    printObj.CurrentY = printObj.CurrentY + 120
                    printObj.Line (printObj.ScaleWidth / 5 * 4, printObj.CurrentY - 20)-(printObj.ScaleWidth - 10, printObj.CurrentY - 20)
                    printObj.CurrentY = printObj.CurrentY + 10
                End With
            End If
            sqlstr = "Select DISTINCT * from qrydeelnemfinales WHERE deelnem=" & rsDeelnem!deelnemID
            sqlstr = sqlstr & " AND wedtype = " & Finale
            sqlstr = sqlstr & " AND ksid= " & kampID
            If rsDeelnFinales.State = adStateOpen Then
                rsDeelnFinales.Close
            End If
            rsDeelnFinales.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
            If rsDeelnFinales.RecordCount > 0 Then
                With rsDeelnFinales
                    printObj.CurrentX = LineXpos + 50
                    Vet True
                    printObj.ForeColor = vbBlue
                    printObj.Print "Finale"
                    printObj.ForeColor = 1
                    Vet False
                    Do While Not .EOF
                        printObj.CurrentX = LineXpos + 50
                        prntReg = Format(!wed, "0") & ": " & !tm1 & " - " & !tm2
                        Do While printObj.TextWidth(prntReg) > printObj.ScaleWidth / 5 - 100
                          prntReg = Left(prntReg, Len(prntReg) - 1)
                        Loop
                        printObj.Print prntReg;
                        printObj.Print
                        If LineYPos < printObj.CurrentY Then LineYPos = printObj.CurrentY
                    .MoveNext
                    Loop
                    .Close
                End With
            End If
            If grpWedsAant > 24 Then
                For i = 2 To 4
                    printObj.Line (printObj.ScaleWidth / 5 * i, newlinepos)-(printObj.ScaleWidth / 5 * i, LineYPos)
                Next
            End If
            printObj.Line (0, newlinepos)-(printObj.ScaleWidth - 10, LineYPos), , B
            'uitslag
            LineYPos = printObj.CurrentY + 50
            LineXpos = 50
            printObj.CurrentX = LineXpos
            printObj.CurrentY = LineYPos
            Vet True
            printObj.ForeColor = vbBlue
            printObj.Print "Eindstand"
            printObj.ForeColor = 1
            Vet False
            printObj.CurrentX = LineXpos
            pr = GetTeam(nz(rsDeelnem!kampioen, 0))
            If pr = "" Then pr = "?"
            printObj.Print "1: "; pr
            printObj.CurrentX = LineXpos
            If getPntToek("2e plaats") > 0 Then
                pr = GetTeam(nz(rsDeelnem!pltwee, 0))
                If pr = "" Then pr = "?"
                printObj.Print "2: "; pr
            Else
                printObj.Print
            End If
            printObj.CurrentX = LineXpos
            If getPntToek("3e plaats") > 0 Then
                pr = GetTeam(nz(rsDeelnem!pldrie, 0))
                If pr = "" Then pr = "?"
                printObj.Print "3: "; pr
            Else
                printObj.Print
            End If
            printObj.CurrentX = LineXpos
            If getPntToek("4e plaats") > 0 Then
              pr = GetTeam(nz(rsDeelnem!plvier, 0))
              If pr = "" Then pr = "?"
              printObj.Print "4: "; pr
            Else
                printObj.Print
            End If
            newlinepos = printObj.CurrentY
            If deelnPag = 1 Then
                oldhelft = Helft
            End If
            printObj.Line (0, LineYPos - 10)-(printObj.ScaleWidth / 8, newlinepos), , B
            'topscorers
            LineXpos = printObj.ScaleWidth / 8 + 50
            printObj.CurrentX = LineXpos
            printObj.CurrentY = LineYPos
            
            Vet True
            printObj.CurrentX = LineXpos + 50
            printObj.ForeColor = vbBlue
            printObj.Print "Topscorer";
            If getPntToek("doelpunten topscorer 1") > 0 Then
                printObj.CurrentX = (printObj.ScaleWidth / 5 * 2) - printObj.TextWidth("doelp") - 100
                printObj.Print "doelp"
            Else
                printObj.Print
            End If
            printObj.ForeColor = 1
            tsYpos = printObj.CurrentY
            kaderPos = printObj.ScaleWidth / 5 * 2
            printObj.Line (LineXpos, LineYPos - 10)-(kaderPos - 10, newlinepos), , B
            Vet False
            printObj.CurrentY = tsYpos
            sqlstr = "Select * from voorspelling_ts WHERE deelnem = " & rsDeelnem!deelnemID
            sqlstr = sqlstr & " ORDER BY tsNR"
            rsDeelnts.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
            Do While Not rsDeelnts.EOF
                printObj.CurrentX = LineXpos + 50
                pr = getSpelerNaam(nz(rsDeelnts!ts, 0))
                printObj.Print pr;
                printObj.CurrentX = kaderPos - printObj.TextWidth(Format(rsDeelnts!dp, "0")) - 150
                If getPntToek("doelpunten topscorer 1") > 0 Then
                    If rsDeelnts!dp > -1 Then
                      printObj.Print Format(rsDeelnts!dp, 0)
                    Else
                        printObj.Print
                    End If
                Else
                    printObj.Print
                End If
                rsDeelnts.MoveNext
            Loop
            rsDeelnts.Close
            'overige
            LineXpos = kaderPos + 20
            kaderPos = printObj.ScaleWidth - 30
            printObj.Line (LineXpos, LineYPos - 10)-(kaderPos, newlinepos), , B
            sqlstr = "Select * from qryDeelnVoorspAant WHERE deelnem = " & rsDeelnem!deelnemID
            rsDeelnOverig.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
            printObj.CurrentY = LineYPos
            printObj.CurrentX = LineXpos + 50
            Vet True
            printObj.ForeColor = vbBlue
            printObj.Print "Overigen ";
            printObj.ForeColor = 1
            LineXpos = printObj.CurrentX
            Vet False
            With rsDeelnOverig
                Do While Not .EOF
                    printObj.CurrentX = LineXpos + 50
                    printObj.Print !omschrijving; ": ";
                    printObj.Print !Aantal
                    .MoveNext
                Loop
                .Close
            End With
            printObj.DrawWidth = 2
            printObj.Line (0, printObj.CurrentY + 50)-(printObj.ScaleWidth - 10, printObj.CurrentY + 50)
            aantalAfgedrukt = aantalAfgedrukt + 1
        End If 'deeln selected
        rsDeelnem.MoveNext
        printObj.CurrentX = 0
        If Not rsDeelnem.EOF Then
            If Me.lstDeelnems.Selected(rsDeelnem.AbsolutePosition - 1) Or Me.Option3 = True Then
                If deelnPag = AantalOpPapier - 1 Then
                    'printObj.Line (0, Helft + 200)-(printObj.ScaleWidth - 10, endEersteDeelnPos + 50), , B
                    deelnPag = 0
                    newlinepos = 0
                    'Exit Do
                    If Not rsDeelnem.EOF Then DoNewPage False, False
                Else
                    endEersteDeelnPos = printObj.CurrentY
                    If aantalAfgedrukt > 0 Then deelnPag = deelnPag + 1
                    
                    If aantalAfgedrukt Mod (AantalOpPapier - 1) = 0 And aantalAfgedrukt > 0 Then
'                        Debug.Print "test"
                    End If
                    printObj.Line (0, printObj.CurrentY + 50)-(printObj.ScaleWidth - 10, endEersteDeelnPos + 50)
                    'printObj.Line (0, TopMarg)-(printObj.ScaleWidth - 10, endEersteDeelnPos + 50), , B
                End If
                printObj.DrawWidth = 1
            End If
        End If
    Loop
    rsDeelnem.Close
    showInfo False
End Sub

Private Sub btnPrint_Click(Index As Integer)

End Sub

Private Sub btnPrnDagResults_Click()
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
  Afdruk_Click 4
  startPrint_Click 0
  'stand op punten
  DoEvents
  Afdruk_Click 2
  Me.ScoreVolg(1) = True
  showInfo True, "Afdrukken", "Stand op punten", "Wedstrijd: " & Me.vscrlTM.value
  startPrint_Click 0
  'stand alfabetisch
  Screen.MousePointer = vbHourglass
  DoEvents
  Afdruk_Click 2
  Me.ScoreVolg(0) = True
  showInfo True, "Afdrukken", "Stand alfabetisch", "Wedstrijd: " & Me.vscrlTM.value
  startPrint_Click 0
  'punten per wedstrijd alfabetisch
  DoEvents
  Afdruk_Click 6
  Me.ScoreVolg(0) = True
  showInfo True, "Afdrukken", "Punten per wedstrijd", "Wedstrijd: " & GetLastPlayed
  tillMatch = GetLastPlayed
  startPrint_Click 0
  'punten opbouw alfabetisch
  DoEvents
  Afdruk_Click 8
  Me.ScoreVolg(0) = True
  Me.optLandscape = True
  showInfo True, "Afdrukken", "Puntenopbouw", "Wedstrijd: " & GetLastPlayed
  startPrint_Click 0
  'grafiek alfabetisch
  DoEvents
  Afdruk_Click 5
  Me.ScoreVolg(0) = True
  showInfo True, "Afdrukken", "Grafiek", "Wedstrijd: " & Me.vscrlTM.value
  startPrint_Click 0
End If
'voorspellingen
curWed = GetMyNum(GetLastPlayed)
If curWed < GetWedAant(kampID) Then
    savdat = getWedDatum(GetWedNum(curWed + 1))
    For i = curWed + 1 To GetWedAant(kampID)
        If Format(getWedDatum(GetWedNum(i)), "d-m-yyyy") = Format(savdat, "d-m-yyyy") Then
            Afdruk_Click 7
            Me.vscrlVoor.value = i
            showInfo True, "Afdrukken", "Voorspelling", "Wedstrijd: " & i
            startPrint_Click 0
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
Afdruk_Click 4
Me.chkDblSide.value = 0
startPrint_Click 0
'stand op punten
DoEvents
Afdruk_Click 2
Me.ScoreVolg(1) = True
showInfo True, "Afdrukken", "Stand op punten", "Wedstrijd: " & Me.vscrlTM.value
Me.chkDblSide.value = 0
startPrint_Click 0
'punten per wedstrijd alfabetisch
DoEvents
Afdruk_Click 6
Me.ScoreVolg(0) = True
Me.chkDblSide.value = 0
showInfo True, "Afdrukken", "Punten per wedstrijd", "Wedstrijd: " & Me.vscrlTM.value
startPrint_Click 0
'punten opbouw alfabetisch
DoEvents
Afdruk_Click 8
Me.ScoreVolg(0) = True
Me.optLandscape = True
Me.chkDblSide.value = 0
showInfo True, "Afdrukken", "Puntenopbouw", "Wedstrijd: " & GetLastPlayed
startPrint_Click 0
'grafiek alfabetisch
DoEvents
Afdruk_Click 5
Me.ScoreVolg(0) = True
Me.chkDblSide.value = 0
showInfo True, "Afdrukken", "Grafiek", "Wedstrijd: " & Me.vscrlTM.value
startPrint_Click 0

'klaar
showInfo False
Screen.MousePointer = vbDefault
MsgBox "Eindstand afgedrukt", vbOKOnly + vbInformation, "Afdrukken"

End Sub

Private Sub cmdEindstand_Click()
    EindStandAfdrukken
End Sub

Private Sub Combo1_Click()
   Dim horPos As Printer
      
   For Each horPos In Printers
      If Combo1.List(Combo1.ListIndex) = horPos.DeviceName Then
         Set Printer = horPos
      End If
   Next

   '  Me.chkDblSide.Visible = True
End Sub

Sub startPrint_Click(Index As Integer)
Dim i As Integer
Dim hoog As Integer
Dim breed As Integer
Dim printForm As Integer
Dim savOrient As Integer
   Dim horPos As Printer
      
   For Each horPos In Printers
      If Combo1.List(Combo1.ListIndex) = horPos.DeviceName Then
         Set Printer = horPos
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
        Set printObj = Printer
        If printObj.Duplex <> 0 Then
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
        Printer.copies = Me.vscrlCopies.value
        afdrRatio = 1
    Else
        Me.Visible = False
        frmPrnt.Show
        If frmPrnt.afdrpic.UBound = 0 Then
            Set printObj = frmPrnt.afdrpic(0)
        End If

    End If
    Set rotate.Device = printObj
    'Meter.Value = Meter.Min
    For i = 0 To 8
        If optPrintSelect(i).value = True Then
            printForm = i
            Exit For
        End If
    Next
    DoEvents
    printObj.Font = txtFont
    Select Case printForm
    Case 0
        printCompetitorForms
    Case 1
        printCompetitors
    Case 2
        'Stand in pool
        printRanking Me.ScoreVolg(0), GetWedNum(val(Me.txttillMatch))
    Case 3
        'Favorieten
        printFavourites
    Case 4
        'toernooi stand
        printTournamentStandings GetWedNum(tillMatch)
    Case 5
        printSkyline
    Case 6
        'punten per wedstrijd
        printPointsPerMatch
    Case 7
        'voorspellingen voor wedstrijd
        printMatchPredictions Me.vscrlVoor
    Case 8
        'samenvatting stand
        printFinalScore Me.ScoreVolg(0)
    End Select
    
    'Melding.Visible = False
    'Picture1.Visible = True
    DoEvents
    If Index = 0 Then
        Printer.EndDoc
    Else
        frmPrnt.Picture1.PaintPicture printObj.Image, 0, 0, printObj.Width, printObj.Height
        Set printObj = Nothing
    End If
    Screen.MousePointer = Default
    
End Sub

Sub printTournamentStandings(tillMatch As Integer)
Dim header2je As String
    headerText = GetOrgNaam(poolID) & " " & getKampInfo("toernooi") & " voetbalpool - Stand van zaken"
    header2je = Format(GetWedInfo(tillMatch, "datum"), "dddd d mmmm") & ": "
    header2je = header2je & GetWedInfo(tillMatch, "naam1") & " vs " & GetWedInfo(tillMatch, "naam2")
    header2$ = "Na wedstrijd " & tillMatch & ", " & header2je
    InitPage False, True
    tnWeds
    tnGroepStanden
    tnFinales
    prnTopScorers
    
    prAantallen tillMatch
    
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
    col(1) = printObj.ScaleWidth / 5
    col(2) = printObj.ScaleWidth / 5 * 2
    col(3) = printObj.ScaleWidth / 5 * 3
    col(4) = printObj.ScaleWidth / 5 * 4
    col(5) = printObj.ScaleWidth
    aantpos = printObj.ScaleWidth / 5
    sqlstr = "select rnaam, afkort, count(rnaam) as aantal from qrywedverloop"
    sqlstr = sqlstr & " WHERE gebeurtenis <= 2"
    sqlstr = sqlstr & " AND ksid = " & kampID
    sqlstr = sqlstr & " GROUP BY rnaam, afkort"
    sqlstr = sqlstr & " ORDER BY count(rnaam) DESC, rnaam"
    rs.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
    sqlstr = "select rnaam, afkort, count(rnaam) as aantal from qrywedverloop"
    sqlstr = sqlstr & " WHERE gebeurtenis = 3"
    sqlstr = sqlstr & " AND ksid = " & kampID
    sqlstr = sqlstr & " GROUP BY rnaam, afkort"
    sqlstr = sqlstr & " ORDER BY count(rnaam) DESC, rnaam"
    rsED.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
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
        ypos = printObj.CurrentY
        printObj.ForeColor = vbBlue
        Vet True
        printObj.Print "Topscorers tot nu toe: "
        ypos = printObj.CurrentY
        Vet False
        printObj.ForeColor = 1
        FontGr 8
        Do While Not rs.EOF
            i = i + 1
            printObj.CurrentX = col(colNu)
            printObj.Print FirstPart(rs!rnaam) & " (" & LCase(rs!afkort) & ")";
            printObj.CurrentX = col(colNu) + aantpos - printObj.TextWidth("1234567890")
            printObj.Print rs!Aantal
            
            
            rs.MoveNext
            If i = Int((rs.RecordCount + rsED.RecordCount + 1) / 5) + 1 Then
                i = 0
                colNu = colNu + 1
                newYpos = printObj.CurrentY
                printObj.CurrentY = ypos
            End If
        Loop
        If rsED.RecordCount > 0 Then
            printObj.ForeColor = vbBlue
            Vet True
            i = i + 1
            printObj.CurrentX = col(colNu)
            printObj.Print "Eigen doelpunten:"
            If i = Int((rs.RecordCount + rsED.RecordCount + 1) / 5) + 1 Then
                i = 0
                colNu = colNu + 1
                newYpos = printObj.CurrentY
                printObj.CurrentY = ypos
            End If
            Vet False
            printObj.ForeColor = 1
            Do While Not rsED.EOF
                i = i + 1
                printObj.CurrentX = col(colNu)
                printObj.Print FirstPart(rsED!rnaam) & " (" & LCase(rsED!afkort) & ")";
                printObj.CurrentX = col(colNu) + aantpos - printObj.TextWidth("1234567890")
                printObj.Print rsED!Aantal
                
                
                rsED.MoveNext
                If i = Int((rs.RecordCount + rsED.RecordCount + 1) / 5) + 1 Then
                    i = 0
                    colNu = colNu + 1
                    newYpos = printObj.CurrentY
                    printObj.CurrentY = ypos
                End If
            Loop
            rsED.Close
        End If
        rs.Close
        printObj.Line (0, ypos)-(printObj.ScaleWidth - 50, newYpos), , B
        printObj.CurrentY = newYpos
        printObj.Print
    End If
End Sub

Sub prAantallen(tillMatch As Integer)
Dim ypos As Integer
Dim prStr As String
Dim col(6) As Integer
    col(0) = 0
    col(1) = printObj.ScaleWidth / 6
    col(2) = printObj.ScaleWidth / 6 * 2
    col(3) = printObj.ScaleWidth / 6 * 3
    col(4) = printObj.ScaleWidth / 6 * 4
    col(5) = printObj.ScaleWidth / 6 * 5
    col(6) = printObj.ScaleWidth - 50
    FontGr 12
    printObj.ForeColor = vbBlue
    Vet True
    printObj.Print "Statistieken"
    ypos = printObj.CurrentY
    Vet False
    printObj.ForeColor = 1
    FontGr 10
    printObj.CurrentX = col(0)
    prStr = "Doelpunten: " & Format(getAantal(tillMatch, 1) + getAantal(tillMatch, 2) + getAantal(tillMatch, 3), pntFormat)
    printObj.Print prStr;
    printObj.CurrentX = col(1)
    prStr = "Penalties: " & Format(getAantal(tillMatch, 1) + getAantal(tillMatch, 6), pntFormat)
    printObj.Print prStr;
    printObj.CurrentX = col(2)
    prStr = "Gele kaarten: " & Format(getAantal(tillMatch, 4), pntFormat)
    printObj.Print prStr;
    printObj.CurrentX = col(3)
    prStr = "Rode kaarten: " & Format(getAantal(tillMatch, 5), pntFormat)
    printObj.Print prStr;
    printObj.CurrentX = col(4)
    prStr = "Gelijke spelen: " & Format(getAantalGelijkeSpelen(tillMatch), pntFormat)
    printObj.Print prStr;
    printObj.CurrentX = col(5)
    prStr = "Eigen doelpunten: " & Format(getAantal(tillMatch, 3), pntFormat)
    printObj.Print prStr
    printObj.ForeColor = vbBlue
    Ital True
    Centreer GetDeelnemAant(poolID) & " deelnemers aan de pool"
    printObj.Print
    Ital False
    printObj.ForeColor = 1
    printObj.Line (col(0), ypos)-(col(6), printObj.CurrentY), , B
        
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
grpAant = getKampInfo("groepen")
    col(0) = 20
    col(1) = printObj.ScaleWidth / 3 + col(0)
    col(2) = printObj.ScaleWidth / 3 * 2 + col(0)
    col(3) = printObj.ScaleWidth
    col(4) = printObj.ScaleWidth / 6 + col(0)
    col(5) = printObj.ScaleWidth / 2 + col(0)
    sqlstr = "Select * from qryWeds "
    sqlstr = sqlstr & " WHERE ksid = " & kampID
    sqlstr = sqlstr & " AND wedtype <> 1"
    sqlstr = sqlstr & " order by mynum, wednum"
    rs.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
    Vet True
    FontGr 12
    printObj.ForeColor = vbBlue
    printObj.Print "Finales"
    topYpos = printObj.CurrentY
    colNr = 0
    printObj.CurrentX = col(colNr)
    FontGr 10
    If grpAant > 4 Then
        printObj.Print "Achtste finales";
        colNr = colNr + 1
        printObj.CurrentX = col(colNr)
    End If
    printObj.Print "Kwart finales";
    colNr = colNr + 1
    printObj.CurrentX = col(colNr)
    printObj.Print "Halve finales";
    If colNr < 2 Then
        colNr = colNr + 1
        printObj.CurrentX = col(colNr)
        printObj.Print "Finale";
    End If
    ypos = printObj.CurrentY
    Vet False
    printObj.ForeColor = 1
    FontGr 8
    numpos = printObj.TextWidth("00")
    datPos = numpos + printObj.TextWidth("0")
    wedPos = datPos + printObj.TextWidth("za 29 jun 20u:")
    vsPos = wedPos + printObj.TextWidth("MEX")
    uitslPos = col(1) - printObj.TextWidth("0-0(0-0)nvl:0-0(mexxx)")
    printObj.Print
    ypos = printObj.CurrentY
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
                printObj.CurrentY = ypos
            End If
        Case Else
            If grpAant > 4 Then
                colNu = 2
            Else
                colNu = 1
            End If
        End Select
        printObj.CurrentX = col(colNu) + numpos - printObj.TextWidth(Format(rs!mynum, "0"))
        printObj.Print Format(rs!mynum, "0");
        printObj.CurrentX = col(colNu) + wedPos - printObj.TextWidth(Format(rs!tijd, "ddd d mmm HHu") & ": ")
        printObj.Print Format(rs!Datum, "ddd d mmm"); tijdFormat(rs!tijd, True); ": "; ' : , " HHu"); ": ";
        printObj.CurrentX = col(colNu) + wedPos
        If nz(rs!tm1, "") > "" Then
            printObj.Print rs!tm1;
        Else
            printObj.Print rs!code1;
        End If
        printObj.CurrentX = col(colNu) + vsPos
        
        If nz(rs!tm2, "") > "" Then
            printObj.Print " - "; rs!tm2;
        Else
            printObj.Print " - "; rs!code2;
        End If
        printObj.CurrentX = col(colNu) + uitslPos
        If WedGespeeld(rs!wedNum) Then
            printObj.Print GetWedUitsl(rs!wedNum)
        Else
            printObj.Print
        End If
        rs.MoveNext
        If Not rs.EOF Then
            If rs!wedtype <> wed Then
                If newYpos < printObj.CurrentY Then
                    newYpos = printObj.CurrentY
                End If
                If rs!wedtype <> klFinale And rs!wedtype <> Finale Then
                    printObj.CurrentY = ypos
                Else
                    Vet True
                    FontGr 12
                    printObj.ForeColor = vbBlue
                    printObj.CurrentX = col(2)
                    If rs!wedtype = klFinale Then
                        printObj.Print "Derde plaats"
                    ElseIf grpAant > 4 Then
                        printObj.CurrentX = col(2)
                        printObj.Print "Finale"
                    End If
                    Vet False
                    printObj.ForeColor = 1
                    FontGr 8
                End If
            End If
        End If
        
    Loop
    printObj.Line (col(0) - 20, topYpos)-(col(1) - 50, newYpos), , B
    printObj.Line (col(1) - 20, topYpos)-(col(2) - 50, newYpos), , B
    printObj.Line (col(2) - 20, topYpos)-(col(3) - 50, newYpos), , B
    
    printObj.CurrentY = newYpos
    printObj.Print
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
    col(1) = printObj.ScaleWidth / 4
    col(2) = printObj.ScaleWidth / 2
    col(3) = printObj.ScaleWidth / 4 * 3
    col(4) = printObj.ScaleWidth
    Vet True
    FontGr 12
    printObj.ForeColor = vbBlue
    printObj.Print "Groepstanden"
    ypos = printObj.CurrentY
    Vet False
    printObj.ForeColor = 1
    FontGr 8
    teampos = 10
    plPos = teampos + printObj.TextWidth("1234567890123")
    wPos = plPos + printObj.TextWidth("000")
    vPos = wPos + printObj.TextWidth("000")
    gPos = vPos + printObj.TextWidth("000")
    pntpos = gPos + printObj.TextWidth("000")
    voorPos = pntpos + printObj.TextWidth("000")
    tegenPos = voorPos + printObj.TextWidth("000")
    
    
    grps = getKampInfo("groepen")
    colNu = 0
    For i = 1 To grps
        printObj.CurrentY = ypos
        sqlstr = "Select * from qryGroepTeams"
        sqlstr = sqlstr & " Where ksID = " & kampID
        sqlstr = sqlstr & " AND groep = '" & Chr(i + 64) & "'"
        sqlstr = sqlstr & " order by pnt DESC, gesp, positie, plaatsing"
        rsGrp.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
        printObj.CurrentX = col(colNu) + teampos
        printObj.Print "groep " & Chr(i + 64);
        printObj.CurrentX = col(colNu) + plPos
        printObj.Print "sp";
        printObj.CurrentX = col(colNu) + wPos
        printObj.Print "W";
        printObj.CurrentX = col(colNu) + vPos
        printObj.Print "V";
        printObj.CurrentX = col(colNu) + gPos
        printObj.Print "G";
        printObj.CurrentX = col(colNu) + pntpos
        printObj.Print "P";
        printObj.CurrentX = col(colNu) + voorPos
        printObj.Print "v-t"
        Do While Not rsGrp.EOF
            pos = pos + 1
            printObj.CurrentX = col(colNu) + teampos
            If rsGrp!positie <> 0 Then
                printObj.Print Format(rsGrp!positie, "0"); ". "; rsGrp!naam;
            Else
                printObj.Print Format(pos, "0"); ". "; rsGrp!naam;
            End If
            printObj.CurrentX = col(colNu) + plPos
            printObj.Print Format(rsGrp!gesp, "0");
            printObj.CurrentX = col(colNu) + wPos
            printObj.Print Format(rsGrp!gew, "0");
            printObj.CurrentX = col(colNu) + vPos
            printObj.Print Format(rsGrp!verl, "0");
            printObj.CurrentX = col(colNu) + gPos
            printObj.Print Format(rsGrp!gel, "0");
            printObj.CurrentX = col(colNu) + pntpos
            printObj.Print Format(rsGrp!pnt, "0");
            printObj.CurrentX = col(colNu) + voorPos
            printObj.Print Format(rsGrp!voor, "0"); "-"; Format(rsGrp!tegen, "0")
            rsGrp.MoveNext
        Loop
        printObj.Line (col(colNu), ypos)-(col(colNu + 1) - 50, printObj.CurrentY), , B
        colNu = colNu + 1
        If colNu > 3 Then
            colNu = 0
            ypos = printObj.CurrentY + 50
        End If
        pos = 0
        rsGrp.Close
    Next
    printObj.Print
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
    col(1) = printObj.ScaleWidth / 3
    col(2) = printObj.ScaleWidth / 3 * 2
    col(3) = printObj.ScaleWidth
    sqlstr = "Select * from qryWeds "
    sqlstr = sqlstr & " WHERE ksid = " & kampID
    sqlstr = sqlstr & " AND wedtype = 1"
    sqlstr = sqlstr & " order by mynum"
    rs.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
    rs.MoveLast
    rs.MoveFirst
    Vet True
    FontGr 12
    printObj.ForeColor = vbBlue
    printObj.Print "Groepswedstrijden"
    ypos = printObj.CurrentY
    Vet False
    FontGr 8
    printObj.ForeColor = 1
    numpos = printObj.TextWidth("000")
    datPos = numpos + printObj.TextWidth("0")
    wedPos = datPos + printObj.TextWidth("za 29 jun 20uW")
    uitslPos = col(1) - printObj.TextWidth("0-0 (0-0)")
    Do While Not rs.EOF
        i = i + 1
        printObj.CurrentX = col(colNu) + numpos - printObj.TextWidth(Format(rs!mynum, "0"))
        printObj.Print Format(rs!mynum, "0");
        printObj.CurrentX = col(colNu) + datPos
'        printObj.Print Format(rs!Datum, "ddd d mmm"); Format(rs!tijd, " HHu."); ": ";
        
        printObj.Print Format(rs!Datum, "ddd d mmm"); tijdFormat(rs!tijd, True); ": ";
        printObj.CurrentX = col(colNu) + wedPos
        printObj.Print rs!naam1 & " - " & rs!naam2;
        printObj.CurrentX = col(colNu) + uitslPos
        If WedGespeeld(rs!wedNum) Then
            printObj.Print GetWedUitsl(rs!wedNum)
        Else
            printObj.Print
        End If
        rs.MoveNext
        If i = rs.RecordCount / 3 Then
            If newYpos < printObj.CurrentY Then
                newYpos = printObj.CurrentY
            End If
            i = 0
            printObj.CurrentY = ypos
            colNu = colNu + 1
        End If
    Loop
    printObj.Line (10, ypos)-(printObj.ScaleWidth - 50, newYpos), , B
    printObj.Line (col(1), ypos)-(col(1), newYpos)
    printObj.Line (col(2), ypos)-(col(2), newYpos)
    printObj.Print
End Sub

Private Sub DoNewPage(pagnr As Boolean, Optional vulheader2 As Boolean, Optional header2pos As Integer)
    If TypeOf printObj Is Printer Then
        Printer.NewPage
    Else
        Load frmPrnt.afdrpic(frmPrnt.afdrpic.UBound + 1)
        frmPrnt.afdrpic(frmPrnt.afdrpic.UBound).Visible = False
        frmPrnt.afdrpic(frmPrnt.afdrpic.UBound).AutoRedraw = True
        Set printObj = frmPrnt.afdrpic(frmPrnt.afdrpic.UBound)
        frmPrnt.brnNext.Enabled = frmPrnt.afdrpic.UBound > 0
    End If
    InitPage pagnr, vulheader2, header2pos, True
End Sub

Private Sub FontGr(grootte%)
    Printer.FontSize = grootte%
    With printObj.Font
        .Size = Printer.FontSize * afdrRatio
    End With
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim horPos As Printer
Dim rs As New ADODB.Recordset
Dim sqlstr As String
  txtFont = "Times New Roman"
  header2Font = "Times New Roman"
    sqlstr = "Select * from pooldeelnems where poolid=" & poolID
    sqlstr = sqlstr & " order by bijnaam"
    rs.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
    Set frmPrnt = New frmAfdrukVB
    Me.lstDeelnems.Clear
    thisMatch = GetMyNum(GetLastPlayed)
    
    Do While Not rs.EOF
        Me.lstDeelnems.AddItem rs!bijnaam
        rs.MoveNext
    Loop
    rs.Close
    Combo1.Clear
    'Load the combo with all available printers
        For Each horPos In Printers
        Combo1.AddItem horPos.DeviceName
        If Printer.DeviceName = horPos.DeviceName Then 'Current default
            Combo1.Text = horPos.DeviceName
        End If
    Next
    
    Width = 6240
    Height = 4800
    centerForm Me
    'admin
    For i = 2 To 8
      Me.Afdruk(i).Visible = True
    Next
    Me.btnPrnDagResults.Visible = True
    Me.cmdEindstand.Visible = True
    Me.Eindstand.Visible = True
    Screen.MousePointer = vbHourglass
    headerText = GetOrgNaam(poolID)
    tillMatch = 0
    If thisMatch >= 1 Then
        tillMatch = thisMatch
        txttillMatch.Enabled = True
    Else
        tillMatch = thisMatch
    End If
   ' Me.chkDblSide.Enabled = printersettings
    Me.vscrlTM.Max = thisMatch
    Me.vscrlVoor.Max = GetWedAant(kampID)
    Me.Afdruk(7).Enabled = GetDeelnemAant(poolID) > 0
    Me.Afdruk(1).Enabled = Me.Afdruk(7).Enabled
    Me.Afdruk(3).Enabled = Me.Afdruk(7).Enabled
    Me.Afdruk(2).Enabled = thisMatch > 0
    Me.Afdruk(4).Enabled = thisMatch > 0
    Me.Afdruk(5).Enabled = thisMatch > 0
    Me.Afdruk(6).Enabled = thisMatch > 0
    Me.Afdruk(8).Enabled = thisMatch > 0
    Me.Afdruk(0).value = True
    Afdruk_Click 0
    Screen.MousePointer = Default
    ' Me.chkDblSide.Visible = true
    Me.Eindstand.Enabled = GetLastPlayed = getlastWednum()
    Me.cmdEindstand.Visible = Me.Eindstand.Enabled
End Sub

Function RandomColor() As Long
    RandomColor = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
End Function


Private Sub printSkyline()
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

header2$ = "Grafiek t/m wedstrijd " & tillMatch
If Me.Eindstand <> 0 Then
    header2$ = "Grafiek Eindstand"
End If
InitPage False, False
FontGr 8
xpos = printObj.CurrentX + printObj.TextWidth("200") + printObj.ScaleLeft
ypos = printObj.CurrentY
sqlstr = "Select deelnemid, bijnaam from pooldeelnems"
sqlstr = sqlstr & " WHERE poolid =  " & poolID
sqlstr = sqlstr & " Order BY bijnaam"
rsDeeln.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
rsDeeln.MoveLast
rsDeeln.MoveFirst
langsteNaam = printObj.TextWidth(Left(GetLangsteBijNaam, 15))
langsteNaam = langsteNaam + printObj.TextWidth("0(99)")
bottom = footerHeight - langsteNaam
yBot = footerHeight - TextHeight("999")
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
HoogsteNu = getHoogPnt(tillMatch)
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
'printObj.Scale
maximum = Int(HoogsteNu / aantpnt + 1) * aantpnt
aantpnt = maximum / factor
scorepos = Int((bottom - ypos) / aantpnt)
'legenda
printObj.FillStyle = vbSolid
oldYpos = bottom
FontGr 6
deelnemPagEenPos = printObj.TextWidth("99: XXX-XXXX") + 20
printObj.ForeColor = vbBlack
For i = 0 To tillMatch - 1
    printObj.FillColor = prntColor(i)
    printObj.Line (xpos, oldYpos)-(xpos + deelnemPagEenPos - 20, oldYpos - printObj.TextHeight("W")), , B
    printObj.CurrentX = xpos + 40
    SetForeCol prntColor(i)
    printObj.Print getWedTeams(i + 1)
    oldYpos = oldYpos - printObj.TextHeight("W")
    printObj.ForeColor = vbBlack
Next
FontGr 8

printObj.Line (xpos + deelnemPagEenPos + 40, ypos)-(printObj.ScaleWidth + 2 * printObj.ScaleLeft, ypos)
printObj.Line -(printObj.ScaleWidth + 2 * printObj.ScaleLeft, bottom)
printObj.Line -(xpos + deelnemPagEenPos + 40, bottom)
printObj.Line -(xpos + deelnemPagEenPos + 40, ypos)
For i = 0 To aantpnt
    ypos = bottom - i * scorepos
    FontGr 8
    printObj.Line (xpos + deelnemPagEenPos + 40, ypos)-(printObj.ScaleWidth + 2 * printObj.ScaleLeft, ypos)
    printObj.CurrentX = xpos + deelnemPagEenPos + 40 - TextWidth(CStr(i * maximum / aantpnt)) - 20
    printObj.CurrentY = ypos - TextHeight("99") / 2
    printObj.Print i * maximum / aantpnt
Next
maximaal = (i - 1) * aantpnt
schaal = (bottom - ypos) / maximum
'FontGr 4
Vet False
rsDeeln.MoveFirst
'prntColor(0) = 15
curPag = 1
deelnpos = Int((printObj.ScaleWidth - (2 * printObj.ScaleLeft) - xpos - deelnemPagEenPos) / deelnemsOpPag)
i = 2 'horizontale positie eerste deelnemer
deelnemsPagEen = deelnemsOpPag - i
Do While Not rsDeeln.EOF
    i = i + 1
    oldYpos = bottom
'    If curPag > 1 Then deelnemsPagEen = deelnemsOpPag
    For j = 0 To tillMatch - 1
        printObj.FillColor = prntColor(j)
        pnt = Int(getDeelnPnt(GetWedNum(j + 1), rsDeeln!deelnemID, 1) * (schaal) + 0.5)
        printObj.Line (xpos + 10 + deelnpos * (i - 1), oldYpos)-(xpos + deelnpos * (i - 1) + deelnpos - 10, oldYpos - pnt), , B
        
        oldYpos = oldYpos - pnt
    Next
    FontGr 8
    printObj.CurrentX = xpos + deelnpos * (i - 1) + (deelnpos - printObj.TextWidth(Format(pnt, "999"))) / 2
    printObj.CurrentY = oldYpos - printObj.TextHeight(Format(pnt, "##"))
    
    printObj.Print Int(getDeelnPnt(GetWedNum(j), rsDeeln!deelnemID, 0))
    printObj.CurrentX = xpos + deelnpos * (i - 1) + (deelnpos - TextWidth("W")) / 2
    tmpX = printObj.CurrentX
    
    printObj.CurrentY = bottom + printObj.TextWidth(Trim(rsDeeln!bijnaam) & " ")
    tmpY = printObj.CurrentY
    Vet False
    FontGr 10
    Set rotate.Device = printObj
    printObj.CurrentY = bottom + 50
    printObj.CurrentX = xpos + deelnpos * (i - 1) + (deelnpos + printObj.TextWidth("W")) / 2
    rotate.Angle = 270
    rotate.PrintText rsDeeln!bijnaam & " (" & getDeelnPnt(tillMatch, rsDeeln!deelnemID, 8) & ")"
    rsDeeln.MoveNext
    printObj.DrawWidth = 1
    If i = deelnemsOpPag And Not rsDeeln.EOF Then
        DoNewPage False, False
        curPag = curPag + 1
        printObj.Line (xpos, ypos)-(printObj.ScaleWidth + 2 * printObj.ScaleLeft, ypos)
        printObj.Line -(printObj.ScaleWidth + 2 * printObj.ScaleLeft, bottom)
        printObj.Line -(xpos, bottom)
        printObj.Line -(xpos, ypos)

        For i = 0 To aantpnt
            ypos = bottom - i * scorepos
            FontGr 8
            printObj.Line (xpos, ypos)-(printObj.ScaleWidth + 2 * printObj.ScaleLeft, ypos)
            printObj.CurrentX = xpos - TextWidth(CStr(i * maximum / aantpnt)) - 10
            printObj.CurrentY = ypos - TextHeight("99") / 2
            printObj.Print i * maximum / aantpnt
        Next
        i = 0
        Vet False
        printObj.FillStyle = vbSolid
    End If
Loop
    
End Sub



Private Sub Init()
    With Printer
        .FontUnderline = 0
        .FontSize = 18
        largeHeight = .TextHeight("Jota")
        .FontSize = 10
        smallHeight = .TextHeight("Jota")
        .FontSize = 8
        verySmallHeight = .TextHeight("Jota")
        .FontSize = 12
        NormalHeight = .TextHeight("Jota")
        .DrawWidth = 2
    End With
End Sub

Private Sub InitPage(pagnr As Boolean, Optional vullen As Boolean, Optional header2pos As Integer, Optional vervolg As Boolean)
' boolean 'doorloop' bepaalt of er een voetregel moet komen
    Me.prnDialog.FontName = txtFont
    
    If Not vervolg Or (vervolg And Me.chkNwePagheader2) Then voetregel
    
    header2Regel
    header2tekst header2$, pagnr, vullen, , header2pos

End Sub

Private Sub Ital(Aan As Boolean)
    printObj.FontItalic = Aan
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    Set printObj = Nothing
  On Error GoTo 0
End Sub

Private Sub KlaarButton_Click()
On Error Resume Next

On Error GoTo 0
Printer.KillDoc
Unload frmPrnt
Unload Me
End Sub

Private Sub header2Regel()
Dim W%
Dim fnt As String
    printObj.ForeColor = RGB(0, 51, 0)
    fnt = printObj.FontName
    printObj.FontName = header2Font
    W% = printObj.DrawWidth
    printObj.DrawWidth = 1
    printObj.Line (0, 0)-(printObj.ScaleWidth, 0), RGB(0, 51, 0)
    FontGr 4
    printObj.Print
    FontGr 16
    Vet True
    Centreer headerText
    printObj.Print
    vertPos% = printObj.CurrentY
    printObj.Line (0, vertPos%)-(printObj.ScaleWidth, vertPos%), RGB(0, 51, 0)
    FontGr 1
    Vet False
    printObj.Print
    headerHeight = printObj.CurrentY
    printObj.DrawWidth = W%
    printObj.ForeColor = vbBlack
    printObj.FontName = fnt
End Sub

Private Sub header2tekst(Tekst$, pagnr As Boolean, Optional vul As Boolean, Optional ypos As Integer, Optional xpos As Integer)
    FontGr 16
    
    printObj.FillColor = RGB(0, 51, 0)
    If vul Then
        printObj.FillStyle = vbFSSolid
        printObj.ForeColor = RGB(204, 251, 153)
        printObj.Line (0, headerHeight)-(printObj.ScaleWidth - 20, headerHeight + printObj.TextHeight("W")), vbBlack, B
    Else
        printObj.ForeColor = RGB(0, 51, 0)
        printObj.FillStyle = vbFSTransparent
    End If
    Ital True
    Vet True
    printObj.CurrentY = headerHeight
    If ypos > 0 Then printObj.CurrentY = ypos
    
    iBKMode = SetBkMode(printObj.hdc, TRANSPARENT)
    Select Case xpos
    Case 0
        Centreer Tekst$
    Case 1
        printObj.CurrentX = 0
        printObj.Print Tekst$;
    Case 2
        printObj.CurrentX = Int(printObj.ScaleWidth / 4) - printObj.TextWidth(Tekst$) / 2
        printObj.Print Tekst$;
    Case 3
        printObj.CurrentX = Int(printObj.ScaleWidth / 2) - printObj.TextWidth(Tekst$) / 2
        printObj.Print Tekst$;
    Case 4
        printObj.CurrentX = Int(printObj.ScaleWidth / 4) * 3 - printObj.TextWidth(Tekst$) / 2
        printObj.Print Tekst$;
    End Select
    favYpos = printObj.CurrentY
    FontGr 9
    printObj.CurrentY = printObj.CurrentY + largeHeight - smallHeight
    printObj.CurrentX = printObj.ScaleWidth - printObj.TextWidth("blad 12")
    If TypeOf printObj Is Printer Then
        If printObj.Page > 1 And pagnr Then
            printObj.Print "blad "; printObj.Page;
        End If
    Else
        If printObj.Index > 0 And pagnr Then
            printObj.Print "blad "; printObj.Index + 1;
        End If
    End If
    FontGr 12
    printObj.Print
    headerHeight = printObj.CurrentY
    printObj.FillStyle = vbFSTransparent
    printObj.ForeColor = vbBlack
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
    rsdeelnScore.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
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
    rsdeelnScore.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
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

Sub printFinalScore(alfabet As Boolean)
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

    grpAant = getKampInfo("groepen")
    If grpAant > 4 Then
        colbr = 140
    Else
        colbr = 250
    End If
    has8eFin = grpAant > 4
    hasKlFin = getKampInfo("derdeplaats")
    If GetLastPlayed = getlastWednum Then
        pntFormat = "0"
    Else
        pntFormat = "0;;\ ;-"
    End If

    leftmarge = printObj.CurrentX
    FontGr 10
    printObj.Print
    
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
    headerText = GetOrgNaam(poolID) & " " & getKampInfo("toernooi") & " voetbalpool"

    header2$ = Tekst$
    
    
    
    InitPage False, True
    Ital False
    Vet False
    FontGr 8
    topYpos = printObj.CurrentY
    printObj.Line (0, topYpos)-(printObj.ScaleWidth - 50, topYpos)
    printObj.CurrentX = leftmarge
    sqlstr = DeelnResultSql(False, GetLastPlayed)
    rsDeeln.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
    If rsDeeln.RecordCount > 0 Then
        rsDeeln.MoveLast
        lastDeelnPos = rsDeeln!postotaal
    End If
    rsDeeln.Close
    sqlstr = DeelnResultSql(alfabet, GetLastPlayed)
    
    rsDeeln.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
    If rsDeeln.RecordCount = 0 Then
        printObj.Print "Geen deelnemers gevonden"
        Exit Sub
    End If
    FontGr 10
    printObj.CurrentX = leftmarge
    printObj.Print "Naam";
    printObj.CurrentX = printObj.TextWidth("123456789012345")
    ReDim Preserve pntpos(1)
    pntpos(0) = 0
    pntpos(1) = printObj.CurrentX - colbr
    printObj.Print
    top2Ypos = printObj.CurrentY
    printObj.CurrentX = pntpos(1) + colbr
    FontGr 8
    printObj.Print "rust"; '("; Format(getPnt(1), pntFormat); "p)";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = printObj.CurrentX
    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
    printObj.Print "eind"; '("; Format(getPnt(2), pntFormat); "p)";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = printObj.CurrentX
    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
    printObj.Print "toto"; '("; Format(getPnt(3), pntFormat); "p)";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = printObj.CurrentX
    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
    printObj.Print "dlp"; '("; Format(getPnt(28), pntFormat); "p)";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = printObj.CurrentX
    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
    printObj.Print "tot";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = printObj.CurrentX
    grpStndBegin = UBound(pntpos)
    
    For i = 1 To grpAant
        printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
        printObj.Print Chr(i + 64);
        ReDim Preserve pntpos(UBound(pntpos) + 1)
        pntpos(UBound(pntpos)) = printObj.CurrentX
    Next
    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
    printObj.Print "tot";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = printObj.CurrentX
    If grpAant > 4 Then
        fin8Begin = UBound(pntpos)
        For i = 1 To grpAant
            printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
            printObj.Print Chr(i + 64);
            ReDim Preserve pntpos(UBound(pntpos) + 1)
            pntpos(UBound(pntpos)) = printObj.CurrentX
        Next
        printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
        printObj.Print "tot";
        ReDim Preserve pntpos(UBound(pntpos) + 1)
        pntpos(UBound(pntpos)) = printObj.CurrentX
    End If
    fin4Begin = UBound(pntpos)
    For i = 1 To 4
        printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
        printObj.Print Format(i, "0");
        ReDim Preserve pntpos(UBound(pntpos) + 1)
        pntpos(UBound(pntpos)) = printObj.CurrentX
    Next
    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
    printObj.Print "tot";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = printObj.CurrentX
    fin2Begin = UBound(pntpos)
    For i = 1 To 2
        printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
        printObj.Print "  "; Format(i, "0"); "e  ";
        ReDim Preserve pntpos(UBound(pntpos) + 1)
        pntpos(UBound(pntpos)) = printObj.CurrentX
    Next
    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
    printObj.Print "tot";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = printObj.CurrentX
    finBegin = UBound(pntpos)
    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
    If hasKlFin Then
        printObj.Print "kl("; Format(getPnt(30), pntFormat);
        If getPnt(31) > 0 Then
            printObj.Print "/"; Format(getPnt(31), pntFormat);
        End If
        printObj.Print ")";
        ReDim Preserve pntpos(UBound(pntpos) + 1)
        pntpos(UBound(pntpos)) = printObj.CurrentX
        printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
        printObj.Print "gr("; Format(getPnt(11), pntFormat);
        If getPnt(12) > 0 Then
            printObj.Print "/"; Format(getPnt(12), pntFormat);
        End If
        printObj.Print ")";
        ReDim Preserve pntpos(UBound(pntpos) + 1)
        pntpos(UBound(pntpos)) = printObj.CurrentX
        printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
    Else
        printObj.Print "("; Format(getPnt(11), pntFormat);
        If getPnt(12) > 0 Then
            printObj.Print "/"; Format(getPnt(12), pntFormat);
        End If
        printObj.Print ")";
        ReDim Preserve pntpos(UBound(pntpos) + 1)
        pntpos(UBound(pntpos)) = printObj.CurrentX
        printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
    End If
    EindstBegin = UBound(pntpos)
    ' Format(getPnt(15), pntFormat); "/"; Format(getPnt(14), pntFormat); "/"; Format(getPnt(13), pntFormat); "/"; Format(getPnt(29), pntFormat); ")";
    printObj.Print "1";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = printObj.CurrentX
    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
    printObj.Print "2";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = printObj.CurrentX
    If hasKlFin Then
        printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
        printObj.Print "3";
        ReDim Preserve pntpos(UBound(pntpos) + 1)
        pntpos(UBound(pntpos)) = printObj.CurrentX
        printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
        printObj.Print "4";
        ReDim Preserve pntpos(UBound(pntpos) + 1)
        pntpos(UBound(pntpos)) = printObj.CurrentX
    End If
    AantBegin = UBound(pntpos)
    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
    printObj.Print "dp";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = printObj.CurrentX
    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
    printObj.Print "gel";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = printObj.CurrentX
    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
    printObj.Print "gl";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = printObj.CurrentX
    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
    printObj.Print "rd";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = printObj.CurrentX
    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
    printObj.Print "pn";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = printObj.CurrentX
    TopScBegin = UBound(pntpos)
    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
    printObj.Print "scor";
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = printObj.CurrentX
    TTLBegin = UBound(pntpos)
    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr + printObj.TextWidth("123")
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = printObj.CurrentX
    PosBegin = UBound(pntpos)
    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr + printObj.TextWidth("123")
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = printObj.CurrentX
    GeldBegin = UBound(pntpos)
    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
    printObj.Print "";
    'laatste columnNrom
    ReDim Preserve pntpos(UBound(pntpos) + 1)
    pntpos(UBound(pntpos)) = printObj.ScaleWidth - 50
    
    printObj.CurrentY = topYpos
    FontGr 10
    printObj.CurrentX = (pntpos(1) + pntpos(grpStndBegin) + colbr - printObj.TextWidth("Wedstrijdpunten")) / 2
    printObj.Print "Wedstrijdpunten";
    If grpAant > 4 Then
        printObj.CurrentX = (pntpos(grpStndBegin) + pntpos(fin8Begin) + colbr - printObj.TextWidth("Groepstand (" & Format(getPnt(8), pntFormat) & "p)")) / 2
    Else
        printObj.CurrentX = (pntpos(grpStndBegin) + pntpos(fin4Begin) + colbr - printObj.TextWidth("Groepstand (" & Format(getPnt(8), pntFormat) & "p)")) / 2
    End If
    printObj.Print "Groepstand (" & Format(getPnt(8), pntFormat) & "p)";
    If grpAant > 4 Then
        printObj.CurrentX = (pntpos(fin8Begin) + pntpos(fin4Begin) + colbr - printObj.TextWidth("8e Finalisten (" & Format(getPnt(6), pntFormat) & "/" & Format(getPnt(7), pntFormat) & "p)")) / 2
        printObj.Print "8e Finalisten (" & Format(getPnt(4), pntFormat);
        If getPnt(5) > 0 Then
            printObj.Print "/" & Format(getPnt(5), pntFormat);
        End If
        printObj.Print "p)";
    End If
    printObj.CurrentX = (pntpos(fin4Begin) + pntpos(fin2Begin) + colbr - printObj.TextWidth("4e fin.(" & Format(getPnt(6), pntFormat) & "/" & Format(getPnt(7), pntFormat) & "p)")) / 2
    printObj.Print "4efin.(" & Format(getPnt(6), pntFormat);
    If getPnt(7) > 0 Then
        printObj.Print "/" & Format(getPnt(7), pntFormat);
    End If
    printObj.Print "p)";
    printObj.CurrentX = (pntpos(fin2Begin) + pntpos(finBegin) + colbr - printObj.TextWidth("2efin.(" & Format(getPnt(9), pntFormat) & "/" & Format(getPnt(10), pntFormat) & "p)")) / 2
    printObj.Print "1/2fin.(" & Format(getPnt(9), pntFormat);
    If getPnt(10) > 0 Then
        printObj.Print "/" & Format(getPnt(10), pntFormat);
    End If
    printObj.Print "p)";
    printObj.CurrentX = (pntpos(finBegin) + pntpos(EindstBegin) + colbr - printObj.TextWidth("Fin")) / 2
    printObj.Print "Fin";
    printObj.CurrentX = (pntpos(EindstBegin) + pntpos(AantBegin) + colbr - printObj.TextWidth("Eind")) / 2
    printObj.Print "Eind";
    printObj.CurrentX = (pntpos(AantBegin) + pntpos(TopScBegin) + colbr - printObj.TextWidth("Aantallen")) / 2
    printObj.Print "Aantallen";
    printObj.CurrentX = pntpos(TopScBegin) + colbr
    printObj.Print "top";
    printObj.CurrentX = (pntpos(TTLBegin) + pntpos(PosBegin) + colbr - printObj.TextWidth("Ttl")) / 2
    printObj.Print "Ttl";
    printObj.CurrentX = (pntpos(PosBegin) + pntpos(GeldBegin) + colbr - printObj.TextWidth("Pos")) / 2
    printObj.Print "Pos";
    printObj.CurrentX = (pntpos(GeldBegin) + pntpos(GeldBegin + 1) + colbr - printObj.TextWidth("Geld")) / 2
    printObj.Print "Geld";
    FontGr 8
    printObj.CurrentY = top2Ypos
    printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
    printObj.Print
    printObj.Line (0, printObj.CurrentY)-(printObj.ScaleWidth - 50, printObj.CurrentY)
    With rsDeeln
        Do While Not .EOF
'            If rsDeeln!deelnemID = 251 Then Stop
            printObj.CurrentX = leftmarge
            If !postotaal = 1 Then
                printObj.ForeColor = vbBlue
                Vet True
            End If
            If !postotaal = lastDeelnPos Then
                printObj.ForeColor = vbRed
            End If
            printObj.Print !bijnaam;
            printObj.ForeColor = 1
            Vet False
            pnt = PrintAant(!deelnemID, pntpos(2), "pntRust")
            pnt = pnt + PrintAant(!deelnemID, pntpos(3), "pntEind")
            pnt = pnt + PrintAant(!deelnemID, pntpos(4), "pntToto")
            pnt = pnt + PrintAant(!deelnemID, pntpos(5), "dpvddag")
            printObj.CurrentX = pntpos(6) - printObj.TextWidth(Format(pnt, pntFormat))
            Vet True
            printObj.Print Format(pnt, pntFormat);
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
                printObj.CurrentX = (pntpos(i + 5) + pntpos(i + 6) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
                printObj.Print Format(grpPnt, pntFormat);
            Next
            If grpAant > 4 Then
                printObj.CurrentX = pntpos(fin8Begin) - printObj.TextWidth(Format(pnt, pntFormat))
            Else
                printObj.CurrentX = pntpos(fin4Begin) - printObj.TextWidth(Format(pnt, pntFormat))
            End If
            Vet True
            printObj.Print Format(pnt, pntFormat);
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
                    printObj.CurrentX = (pntpos(fin8Begin - 1 + i) + pntpos(i + fin8Begin) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
                    printObj.Print Format(grpPnt, pntFormat);
                Next
                printObj.CurrentX = pntpos(fin4Begin) - printObj.TextWidth(Format(pnt, pntFormat))
                Vet True
                If allPlayed("A") Then
                    pntFormat = "0"
                Else
                    pntFormat = "0;;\ ;-"
                End If
                printObj.Print Format(pnt, pntFormat);
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
                    printObj.CurrentX = (pntpos(fin4Begin - 1 + i) + pntpos(i + fin4Begin) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
                    printObj.Print Format(grpPnt, pntFormat);
                Next
                Vet True
                If allPlayed("A") Then
                    pntFormat = "0"
                Else
                    pntFormat = "0;;\ ;-"
                End If
                printObj.CurrentX = pntpos(fin2Begin) - printObj.TextWidth(Format(pnt, pntFormat))
                printObj.Print Format(pnt, pntFormat);
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
                    printObj.CurrentX = (pntpos(ipos + fin4Begin - 1) + pntpos(ipos + fin4Begin) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
                    printObj.Print Format(grpPnt, pntFormat);
                Next
                If prTtl > 0 Then pntFormat = "0"
                printObj.CurrentX = pntpos(fin2Begin) - printObj.TextWidth(Format(pnt, pntFormat))
                Vet True
                printObj.Print Format(pnt, pntFormat);
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
                printObj.CurrentX = (pntpos(i + fin2Begin - 1) + pntpos(i + fin2Begin) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
                printObj.Print Format(grpPnt, pntFormat);
            Next
            printObj.CurrentX = pntpos(finBegin) - printObj.TextWidth(Format(pnt, pntFormat))
            Vet True
            printObj.Print Format(pnt, pntFormat);
            Vet False
            If GetMyNum(GetFirstFinaleMatch(HalveFinale)) <= GetMyNum(GetLastPlayed) Then
                pntFormat = "0"
            Else
                pntFormat = "0;;\ ;-"
            End If
            If hasKlFin Then
                grpPnt = GetPntDeelnem(!deelnID, "pntklfin")
                printObj.CurrentX = pntpos(32) + (pntpos(33) - pntpos(32) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
                printObj.Print Format(grpPnt, pntFormat);
            End If
            grpPnt = GetPntDeelnem(!deelnID, "pntfin")
            printObj.CurrentX = (pntpos(finBegin + 1) + pntpos(EindstBegin) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
            printObj.Print Format(grpPnt, pntFormat);
            pntFormat = "0;;\ ;-"
            If GetLastPlayed = getlastWednum Then pntFormat = "0"
            For i = 1 To 2
                grpPnt = getEindStandpnt(!deelnID, i)
                printObj.CurrentX = (pntpos(finBegin + 1 + i) + pntpos(EindstBegin + i) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
                printObj.Print Format(grpPnt, pntFormat);
            Next
            pntFormat = "0;;\ ;-"
            If GetLastPlayed >= getlastWednum - 1 Then pntFormat = "0"
            For i = 3 To 4
                grpPnt = getEindStandpnt(!deelnID, i)
                printObj.CurrentX = (pntpos(EindstBegin - 1 + i) + pntpos(EindstBegin + i) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
                printObj.Print Format(grpPnt, pntFormat);
            Next
            pntFormat = "0;;\ ;-"
            If GetLastPlayed = getlastWednum Then
                pntFormat = "0"
                grpPnt = getDeelnAantPnt(!deelnID, voorspDP)
                printObj.CurrentX = (pntpos(AantBegin) + pntpos(AantBegin + 1) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
                printObj.Print Format(grpPnt, pntFormat);
                grpPnt = getDeelnAantPnt(!deelnID, voorspGelijk)
                printObj.CurrentX = (pntpos(AantBegin + 1) + pntpos(AantBegin + 2) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
                printObj.Print Format(grpPnt, pntFormat);
                grpPnt = getDeelnAantPnt(!deelnID, voorspGeel)
                printObj.CurrentX = (pntpos(AantBegin + 2) + pntpos(AantBegin + 3) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
                printObj.Print Format(grpPnt, pntFormat);
                grpPnt = getDeelnAantPnt(!deelnID, voorspRood)
                printObj.CurrentX = (pntpos(AantBegin + 3) + pntpos(AantBegin + 4) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
                printObj.Print Format(grpPnt, pntFormat);
                grpPnt = getDeelnAantPnt(!deelnID, voorspPens)
                printObj.CurrentX = (pntpos(AantBegin + 4) + pntpos(TopScBegin) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
                printObj.Print Format(grpPnt, pntFormat);
                grpPnt = GetPntDeelnem(!deelnID, "pntTopSc")
                printObj.CurrentX = (pntpos(TopScBegin) + pntpos(TTLBegin) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
                printObj.Print Format(grpPnt, pntFormat);
            End If
            pntFormat = "0"
            If !postotaal = 1 Then
                printObj.ForeColor = vbBlue
                Vet True
            End If
            If !postotaal = lastDeelnPos Then
                printObj.ForeColor = vbRed
            End If
            'If !deelnID = 125 Then Stop
            grpPnt = GetPntDeelnem(!deelnID, "grandtotaal")
            printObj.CurrentX = (pntpos(TTLBegin) + pntpos(PosBegin) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
            printObj.Print Format(grpPnt, pntFormat);
            grpPnt = GetPntDeelnem(!deelnID, "postotaal")
            printObj.CurrentX = (pntpos(PosBegin) + pntpos(GeldBegin) + colbr - printObj.TextWidth(Format(grpPnt, pntFormat))) / 2
            printObj.Print Format(grpPnt, pntFormat);
            printObj.ForeColor = 1
            Vet False
            geld = GetPntDeelnem(!deelnID, "geldttl")
            printObj.CurrentX = pntpos(GeldBegin + 1) - colbr - printObj.TextWidth(Format(geld, "currency"))
            printObj.Print Format(geld, "currency");
            printObj.Print
            printObj.ForeColor = 1
            printObj.Line (0, printObj.CurrentY)-(printObj.ScaleWidth - 50, printObj.CurrentY)
            grpPnt = 0
            
            .MoveNext
'            If .AbsolutePosition >= 53 Then Stop
            If printObj.CurrentY >= footerHeight Then 'onderkant pagina
              If Not rsDeeln.EOF Then
                botY = printObj.CurrentY
                printObj.Line (pntpos(1) + 75, topYpos)-(pntpos(1) + 75, top2Ypos)
                printObj.Line (pntpos(grpStndBegin) + 75, topYpos)-(pntpos(6) + 75, top2Ypos)
                If grpAant > 4 Then
                    printObj.Line (pntpos(fin8Begin) + 75, topYpos)-(pntpos(15) + 75, top2Ypos)
                End If
                printObj.Line (pntpos(fin4Begin) + 75, topYpos)-(pntpos(fin4Begin) + 75, top2Ypos)
                printObj.Line (pntpos(fin2Begin) + 75, topYpos)-(pntpos(fin2Begin) + 75, top2Ypos)
                printObj.Line (pntpos(finBegin) + 75, topYpos)-(pntpos(finBegin) + 75, top2Ypos)
                printObj.Line (pntpos(EindstBegin) + 75, topYpos)-(pntpos(EindstBegin) + 75, top2Ypos)
                printObj.Line (pntpos(AantBegin) + 75, topYpos)-(pntpos(AantBegin) + 75, top2Ypos)
                printObj.Line (pntpos(TopScBegin) + 75, topYpos)-(pntpos(TopScBegin) + 75, top2Ypos)
                printObj.Line (pntpos(TTLBegin) + 75, topYpos)-(pntpos(TTLBegin) + 75, top2Ypos)
                printObj.Line (pntpos(PosBegin) + 75, topYpos)-(pntpos(PosBegin) + 75, top2Ypos)
                printObj.Line (pntpos(GeldBegin) + 75, topYpos)-(pntpos(GeldBegin) + 75, top2Ypos)
                For i = 1 To UBound(pntpos) - 1
                    printObj.Line (pntpos(i) + 75, top2Ypos)-(pntpos(i) + 75, botY)
                Next
                printObj.Line (printObj.ScaleWidth - 50, topYpos)-(printObj.ScaleWidth - 50, botY)
                DoNewPage False, True
                printObj.Line (0, topYpos)-(printObj.ScaleWidth - 50, topYpos)
                printObj.CurrentX = leftmarge
                printObj.CurrentY = topYpos
                FontGr 10
                printObj.Print "Naam";
                printObj.CurrentY = top2Ypos
                printObj.CurrentX = pntpos(1) + colbr
                FontGr 8
                printObj.Print "rust"; '("; Format(getPnt(1), pntFormat); "p)";
                printObj.CurrentX = pntpos(2) + colbr
                printObj.Print "eind"; '("; Format(getPnt(2), pntFormat); "p)";
                printObj.CurrentX = pntpos(3) + colbr
                printObj.Print "toto"; '("; Format(getPnt(3), pntFormat); "p)";
                printObj.CurrentX = pntpos(4) + colbr
                printObj.Print "dlp"; '("; Format(getPnt(28), pntFormat); "p)";
                printObj.CurrentX = pntpos(5) + colbr
                printObj.Print "tot";
                If grpAant > 4 Then
                    For i = 1 To 8
                        printObj.CurrentX = pntpos(5 + i) + colbr
                        printObj.Print Chr(i + 64);
                    Next
                    printObj.CurrentX = pntpos(14) + colbr
                    printObj.Print "tot";
                    For i = 1 To 8
                        printObj.CurrentX = pntpos(14 + i) + colbr
                        printObj.Print Chr(i + 64);
                    Next
                    printObj.CurrentX = pntpos(23) + colbr
                    printObj.Print "tot";
                End If
                For i = 1 To 4
                    printObj.CurrentX = pntpos(fin4Begin - 1 + i) + colbr
                    printObj.Print Format(i, "0");
                Next
                printObj.CurrentX = pntpos(fin2Begin - 1) + colbr
                printObj.Print "tot";
                For i = 1 To 2
                    printObj.CurrentX = pntpos(fin2Begin - 1 + i) + colbr
                    printObj.Print "  "; Format(i, "0"); "e  ";
                Next
                printObj.CurrentX = pntpos(finBegin - 1) + colbr
                printObj.Print "tot";
                printObj.CurrentX = pntpos(finBegin) + colbr
                If hasKlFin Then
                    printObj.Print "kl("; Format(getPnt(30), pntFormat);
                    If getPnt(31) > 0 Then
                        printObj.Print "/"; Format(getPnt(31), pntFormat);
                    End If
                    printObj.Print ")";
                    printObj.CurrentX = pntpos(EindstBegin - 1) + colbr
                    printObj.Print "gr("; Format(getPnt(11), pntFormat);
                    If getPnt(12) > 0 Then
                        printObj.Print "/"; Format(getPnt(12), pntFormat);
                    End If
                    printObj.Print ")";
                Else
                    printObj.Print "("; Format(getPnt(11), pntFormat);
                    If getPnt(12) > 0 Then
                        printObj.Print "/"; Format(getPnt(12), pntFormat);
                    End If
                    printObj.Print ")";
                End If
                
                For i = 1 To grpAant / 2
                    printObj.CurrentX = pntpos(EindstBegin - 1 + i) + colbr
                    ' Format(getPnt(15), pntFormat); "/"; Format(getPnt(14), pntFormat); "/"; Format(getPnt(13), pntFormat); "/"; Format(getPnt(29), pntFormat); ")";
                    printObj.Print Format(i, "0");
                Next
                printObj.CurrentX = pntpos(AantBegin) + colbr
                printObj.Print "dp";
                printObj.CurrentX = pntpos(AantBegin + 1) + colbr
                printObj.Print "gel";
                printObj.CurrentX = pntpos(AantBegin + 2) + colbr
                printObj.Print "gl";
                printObj.CurrentX = pntpos(AantBegin + 3) + colbr
                printObj.Print "rd";
                printObj.CurrentX = pntpos(AantBegin + 4) + colbr
                printObj.Print "pn";
                printObj.CurrentX = pntpos(TopScBegin) + colbr
                printObj.Print "scor";
                'laatste columnNrom
                printObj.CurrentX = printObj.ScaleWidth - 50
                
                printObj.CurrentY = topYpos
                FontGr 10
                printObj.CurrentX = (pntpos(1) + pntpos(grpStndBegin) + colbr - printObj.TextWidth("Wedstrijdpunten")) / 2
                printObj.Print "Wedstrijdpunten";
                If grpAant > 4 Then
                    printObj.CurrentX = (pntpos(grpStndBegin) + pntpos(fin8Begin) + colbr - printObj.TextWidth("Groepstand (" & Format(getPnt(8), pntFormat) & "p)")) / 2
                Else
                    printObj.CurrentX = (pntpos(grpStndBegin) + pntpos(fin4Begin) + colbr - printObj.TextWidth("Groepstand (" & Format(getPnt(8), pntFormat) & "p)")) / 2
                End If
                printObj.Print "Groepstand (" & Format(getPnt(8), pntFormat) & "p)";
                If grpAant > 4 Then
                    printObj.CurrentX = (pntpos(fin8Begin) + pntpos(fin4Begin) + colbr - printObj.TextWidth("8e Finalisten (" & Format(getPnt(6), pntFormat) & "/" & Format(getPnt(7), pntFormat) & "p)")) / 2
                    printObj.Print "8e Finalisten (" & Format(getPnt(4), pntFormat);
                    If getPnt(5) > 0 Then
                        printObj.Print "/" & Format(getPnt(5), pntFormat);
                    End If
                    printObj.Print "p)";
                End If
                printObj.CurrentX = (pntpos(fin4Begin) + pntpos(fin2Begin) + colbr - printObj.TextWidth("4e fin.(" & Format(getPnt(6), pntFormat) & "/" & Format(getPnt(7), pntFormat) & "p)")) / 2
                printObj.Print "4efin.(" & Format(getPnt(6), pntFormat);
                If getPnt(7) > 0 Then
                    printObj.Print "/" & Format(getPnt(7), pntFormat);
                End If
                printObj.Print "p)";
                printObj.CurrentX = (pntpos(fin2Begin) + pntpos(finBegin) + colbr - printObj.TextWidth("2efin.(" & Format(getPnt(9), pntFormat) & "/" & Format(getPnt(10), pntFormat) & "p)")) / 2
                printObj.Print "1/2fin.(" & Format(getPnt(9), pntFormat);
                If getPnt(10) > 0 Then
                    printObj.Print "/" & Format(getPnt(10), pntFormat);
                End If
                printObj.Print "p)";
                printObj.CurrentX = (pntpos(finBegin) + pntpos(EindstBegin) + colbr - printObj.TextWidth("Fin")) / 2
                printObj.Print "Fin";
                printObj.CurrentX = (pntpos(EindstBegin) + pntpos(AantBegin) + colbr - printObj.TextWidth("Eind")) / 2
                printObj.Print "Eind";
                printObj.CurrentX = (pntpos(AantBegin) + pntpos(TopScBegin) + colbr - printObj.TextWidth("Aantallen")) / 2
                printObj.Print "Aantallen";
                printObj.CurrentX = pntpos(TopScBegin) + colbr
                printObj.Print "top";
                printObj.CurrentX = (pntpos(TTLBegin) + pntpos(PosBegin) + colbr - printObj.TextWidth("Ttl")) / 2
                printObj.Print "Ttl";
                printObj.CurrentX = (pntpos(PosBegin) + pntpos(GeldBegin) + colbr - printObj.TextWidth("Pos")) / 2
                printObj.Print "Pos";
                printObj.CurrentX = (pntpos(GeldBegin) + pntpos(GeldBegin + 1) + colbr - printObj.TextWidth("Geld")) / 2
                printObj.Print "Geld";
                FontGr 8
                printObj.CurrentY = top2Ypos
                printObj.CurrentX = pntpos(UBound(pntpos)) + colbr
                printObj.Print
                printObj.Line (0, printObj.CurrentY)-(printObj.ScaleWidth - 50, printObj.CurrentY)
              End If
            End If
        Loop
    End With
    botY = printObj.CurrentY
    printObj.Line (pntpos(1) + 75, topYpos)-(pntpos(1) + 75, top2Ypos)
    printObj.Line (pntpos(grpStndBegin) + 75, topYpos)-(pntpos(6) + 75, top2Ypos)
    If grpAant > 4 Then
        printObj.Line (pntpos(fin8Begin) + 75, topYpos)-(pntpos(15) + 75, top2Ypos)
    End If
    printObj.Line (pntpos(fin4Begin) + 75, topYpos)-(pntpos(fin4Begin) + 75, top2Ypos)
    printObj.Line (pntpos(fin2Begin) + 75, topYpos)-(pntpos(fin2Begin) + 75, top2Ypos)
    printObj.Line (pntpos(finBegin) + 75, topYpos)-(pntpos(finBegin) + 75, top2Ypos)
    printObj.Line (pntpos(EindstBegin) + 75, topYpos)-(pntpos(EindstBegin) + 75, top2Ypos)
    printObj.Line (pntpos(AantBegin) + 75, topYpos)-(pntpos(AantBegin) + 75, top2Ypos)
    printObj.Line (pntpos(TopScBegin) + 75, topYpos)-(pntpos(TopScBegin) + 75, top2Ypos)
    printObj.Line (pntpos(TTLBegin) + 75, topYpos)-(pntpos(TTLBegin) + 75, top2Ypos)
    printObj.Line (pntpos(PosBegin) + 75, topYpos)-(pntpos(PosBegin) + 75, top2Ypos)
    printObj.Line (pntpos(GeldBegin) + 75, topYpos)-(pntpos(GeldBegin) + 75, top2Ypos)
    For i = 1 To UBound(pntpos) - 1
        printObj.Line (pntpos(i) + 75, top2Ypos)-(pntpos(i) + 75, botY)
    Next
    printObj.Line (printObj.ScaleWidth - 50, topYpos)-(printObj.ScaleWidth - 50, botY)
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
'    printObj.CurrentX = pos - printObj.TextWidth("(" & Format(aant, "0") & "x) " & Format(aant * pnt, "0"))
    printObj.CurrentX = pos - printObj.TextWidth(Format(aant * pnt, "0"))
    Ital True
'    printObj.Print "(" & Format(aant, "0"); "x) ";
    Ital False
    printObj.Print Format(aant * pnt, "0");
    PrintAant = aant * pnt
End Function

Sub printRanking(alfabet As Boolean, wedNum As Integer)
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
Dim deelcolumnWidth%
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
    leftmarge = printObj.CurrentX
    deelcolumnWidth% = (printObj.ScaleWidth + 2 * printObj.ScaleLeft) \ 2
    FontGr 10
    deelnaampos% = printObj.TextWidth("999.")
    DeelOldPntPos% = deelnaampos% + deelcolumnWidth% / 4 - 200
    DeelWedPntPos% = DeelOldPntPos% + deelcolumnWidth / 10
    DeelNewPntPos% = DeelWedPntPos% + deelcolumnWidth / 10
    
    deelgeldpos% = DeelNewPntPos% + deelcolumnWidth / 6 + 200
    DeelGeldnwPos% = deelgeldpos% + deelcolumnWidth / 6 - 100
    DeelGeldttlPos% = DeelGeldnwPos% + deelcolumnWidth / 6 - 100
    
    If alfabet Then
        deelnaampos% = Me.CurrentX + 40
    End If
    
    printObj.Print
    
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
    headerText = GetOrgNaam(poolID) & " " & getKampInfo("toernooi") & " voetbalpool"

    header2$ = Tekst$
    
    InitPage False, True
    Ital False
    Vet False
    FontGr 10
    printObj.CurrentX = (printObj.ScaleWidth - printObj.TextWidth("onderstreept=daghoogste, vet=bovenaan, cursief=onderaan")) / 2
    printObj.Print "(";
    printObj.FontUnderline = True
    printObj.ForeColor = &H8000&
    printObj.Print "onderstreept";
    printObj.FontUnderline = False
    printObj.ForeColor = 0
    printObj.Print "= daghoogste, ";
    printObj.ForeColor = vbBlue
    Vet True
    printObj.Print "vet";
    Vet False
    printObj.ForeColor = 0
    printObj.Print "= bovenaan, ";
    Ital True
    printObj.ForeColor = vbRed
    printObj.Print "cursief";
    Ital False
    printObj.ForeColor = 0
    printObj.Print "= onderaan)"
    
    savy = printObj.CurrentY
    For columnNr% = 0 To 1
        If Not alfabet Then
            printObj.CurrentX = columnNr% * deelcolumnWidth%
            'printObj.Print "pos";
        End If
        printObj.CurrentX = deelnaampos% + columnNr% * deelcolumnWidth%
        printObj.Print "Naam";
        If alfabet Then printObj.Print " (pl)";
        printObj.CurrentX = DeelOldPntPos% + columnNr% * deelcolumnWidth%
        printObj.Print "had  +";
        printObj.CurrentX = DeelWedPntPos% + columnNr% * deelcolumnWidth%
        printObj.Print "erbij =";
        printObj.CurrentX = DeelNewPntPos% + columnNr% * deelcolumnWidth% + printObj.TextWidth("999") - printObj.TextWidth("nu")
        printObj.Print "nu";
        printObj.CurrentX = deelgeldpos% - printObj.TextWidth("Geld") + columnNr% * deelcolumnWidth%
        printObj.Print "Geld";
        printObj.CurrentX = DeelGeldnwPos% - printObj.TextWidth("erbij") + columnNr% * deelcolumnWidth%
        printObj.Print "erbij";
        printObj.CurrentX = DeelGeldttlPos% - printObj.TextWidth("totaal") + columnNr% * deelcolumnWidth%
        printObj.Print "totaal";
    Next
    printObj.CurrentY = printObj.CurrentY + 50
    yLinePos% = printObj.CurrentY + TextHeight("test")
    printObj.Line (leftmarge, yLinePos%)-(printObj.ScaleWidth + printObj.ScaleLeft * 2, yLinePos%)
    printObj.CurrentY = printObj.CurrentY + 50
    DeelTopPos% = printObj.CurrentY
'    printObj.Print
    'bepaal hoogste en laagste
    rsDeeln.Open DeelnResultSql(False, wedNum), dbConn, adOpenStatic, adLockReadOnly 'op punten volgorde dus
    If rsDeeln.RecordCount > 0 Then
        rsDeeln.MoveLast
        last = nz(rsDeeln!grandtotaal, 0)
    Else
        Exit Sub
    End If
    rsDeeln.Close
    printObj.CurrentX = 0
    'en nu opnieuw openen
    rsDeeln.Open DeelnResultSql(alfabet, wedNum), dbConn, adOpenStatic, adLockReadOnly 'op volgorde dus
    With rsDeeln
        If .RecordCount > 0 Then
            .MoveFirst
            lastttl = 0
            columnNr% = 0
            Do While Not .EOF
                i = i + 1
                If i = Int(.RecordCount / 2 + 0.5) + 1 Then
                    columnNr% = deelcolumnWidth%
                    printObj.CurrentY = DeelTopPos%
                End If
                printObj.CurrentX = printObj.CurrentX + deelnaampos% - printObj.TextWidth(!postotaal) - printObj.TextWidth("..") + columnNr%
                If Not alfabet Then
                    If lastttl <> !grandtotaal Then printObj.Print !postotaal & ".";
                End If
                Vet !postotaal = 1
                Ital nz(!grandtotaal, 0) = last
                prStr = Left(!bijnaam, 12)
                If alfabet Then
                    prStr = prStr & " (" & !postotaal & ")"
                End If
                If !grandtotaal = last Then
                    printObj.ForeColor = vbRed
                ElseIf nz(!postotaal, 0) = 1 Then
                    printObj.ForeColor = vbBlue
                ElseIf nz(!posdag, 0) = 1 Then
                    printObj.ForeColor = &H8000&
                Else
                    printObj.ForeColor = 0
                End If
                printObj.CurrentX = deelnaampos% + columnNr%
                printObj.FontUnderline = nz(!posdag, 0) = 1
                
                printObj.Print prStr;
                Vet False
                Ital False
                printObj.ForeColor = 0
                printObj.FontUnderline = False
                If wedNum > 1 Then
                    pnt = getTussenstand(!deelnemID, wedNum)
                    geldold = getTussenstandGeld(!deelnemID, GetWedNumPrevDag(wedNum))
                Else
                    pnt = 0
                    geldold = 0
                End If
                
                printObj.CurrentX = DeelOldPntPos% + columnNr% + printObj.TextWidth("999") - printObj.TextWidth(CStr(pnt))
                printObj.Print Format$(pnt, "##0");
                Vet False
                pnt = nz(!Dagpnt, 0)
                printObj.CurrentX = DeelWedPntPos% + columnNr% + printObj.TextWidth("999") - printObj.TextWidth(CStr(pnt))
                printObj.FontUnderline = nz(!posdag, 0) = 1
                If !posdag = 1 Then
                    printObj.ForeColor = &H8000&
                Else
                    printObj.ForeColor = 0
                End If
                printObj.Print Format$(pnt, "##0");
                printObj.ForeColor = 0
                printObj.FontUnderline = False
                Vet !postotaal = 1
                Ital nz(!grandtotaal, 0) = last
                pnt = nz(!grandtotaal, 0)
                If !grandtotaal = last Then
                    printObj.ForeColor = vbRed
                ElseIf !postotaal = 1 Then
                    printObj.ForeColor = vbBlue
                Else
                    printObj.ForeColor = 0
                End If
                printObj.CurrentX = DeelNewPntPos% + columnNr% + printObj.TextWidth("999") - printObj.TextWidth(CStr(pnt))
                If !grandtotaal = last Then
                    printObj.ForeColor = &H80&
                ElseIf !postotaal = 1 Then
                    printObj.ForeColor = &HC00000
                Else
                    printObj.ForeColor = 0
                End If
                printObj.Print Format$(!grandtotaal, "##0");
                printObj.ForeColor = 0
                Vet False
                Ital False
                tmp$ = Format$(geldold, "€ ##0.00")
                printObj.CurrentX = deelgeldpos% - printObj.TextWidth(tmp$) + columnNr%
                printObj.Print tmp$;   '= geld
                tmp$ = Format$(!daggeldttl, "€ ##0.00")
                printObj.CurrentX = DeelGeldnwPos% - printObj.TextWidth(tmp$) + columnNr%
                printObj.Print tmp$;
                bedr = 0
                tmp$ = Format$(geldold + !daggeldttl, "€ ##0.00")
                printObj.CurrentX = DeelGeldttlPos% - printObj.TextWidth(tmp$) + columnNr%
                printObj.Print tmp$;   '= geld
                printObj.Print
                lastttl = nz(!grandtotaal, 0)
                rsDeeln.MoveNext
            Loop
            printObj.Print
            yposnu% = printObj.CurrentY
            printObj.Line (deelcolumnWidth%, savy)-(deelcolumnWidth%, yposnu%)
            printObj.Line (deelgeldpos - printObj.TextWidth("Geld") - 400, yLinePos%)-(deelgeldpos - printObj.TextWidth("Geld") - 400, yposnu%)
            printObj.Line (deelgeldpos - printObj.TextWidth("Geld") - 400 + deelcolumnWidth%, yLinePos%)-(deelgeldpos - printObj.TextWidth("Geld") - 400 + deelcolumnWidth%, yposnu%)
            printObj.Line (leftmarge, yposnu%)-(printObj.ScaleWidth + printObj.ScaleLeft * 2, yposnu%)
        End If
        .Close
    End With
    printObj.Print
End Sub

Function DeelnResultSql(alfabet As Boolean, wedNum As Integer) As String
Dim sql As String
    sql = "Select deelnemID, bijnaam, wednum,"
    sql = sql & " deelnempnt.*"
    sql = sql & " from deelnempnt, pooldeelnems"
    sql = sql & " WHERE pooldeelnems!deelnemID = deelnempnt.deelnid"
    sql = sql & " AND pooldeelnems!poolID = " & poolID
    sql = sql & " AND wednum = " & wedNum
    If alfabet Then
        sql = sql & " ORDER BY bijnaam"
    Else
        sql = sql & " ORDER BY grandtotaal DESC, bijnaam ASC"
    End If
    DeelnResultSql = sql
End Function


Private Sub lstDeelnems_Click()
    Me.Option4.value = True
End Sub

Private Sub txtVoorwed_Change()
chkTxtValue Me.txtVoorWed, Me.vscrlVoor
tillMatch = val(txtVoorWed.Text)
End Sub

Private Sub txttillMatch_Change()
chkTxtValue Me.txttillMatch, Me.vscrlTM
tillMatch = val(txttillMatch.Text)
End Sub

Private Sub Vet(Aan As Boolean)
    printObj.FontBold = Aan
End Sub

Private Sub voetregel()
Dim W%
Dim i As Double
Dim fontnaam As String
    printObj.ForeColor = RGB(0, 51, 0)
    W% = printObj.DrawWidth
    printObj.DrawWidth = 1
    FontGr 8
    Ital True
    Vet False
    fontnaam = printObj.FontName
    printObj.FontName = "Garamond"
    printObj.CurrentY = printObj.ScaleHeight - printObj.TextHeight("w")
    footerHeight = printObj.CurrentY - printObj.TextHeight("w")
    vertPos% = printObj.CurrentY
    printObj.Line (0, vertPos% - 15 * afdrRatio)-(printObj.ScaleWidth, vertPos% - 15 * afdrRatio)
    printObj.CurrentY = vertPos%
    Centreer "© 2004-" & Year(Now) & " jota computer assistentie"
    printObj.FontName = fontnaam
    printObj.Print
    FontGr 12
    Vet False
    Ital False
    vertPos% = printObj.CurrentY + 50 * afdrRatio
    'printObj.Line (0, y%)-(printObj.ScaleWidth, y%)
    printObj.ForeColor = vbBlack
    printObj.DrawWidth = 1
End Sub


Private Sub vscrlCopies_Change()
    Me.copies.Text = Me.vscrlCopies.value
End Sub

Private Sub vscrlTM_Change()
Me.txttillMatch = Me.vscrlTM.value
tillMatch = Me.vscrlTM
End Sub

Private Sub vscrlVoor_Change()
Me.txtVoorWed = Me.vscrlVoor.value
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
printObj.CurrentX = (printObj.ScaleWidth - printObj.TextWidth(infostr)) / 2
printObj.Print "Samenstelling punten: ";
Ital True
printObj.Print "toto goed";
If inclpnt Then printObj.Print pntToto; "pnt";
Ital False
printObj.Print ", ";
printObj.FontUnderline = True
printObj.Print "rust goed";
If inclpnt Then printObj.Print pntRust; "pnt";
printObj.FontUnderline = False
printObj.Print ", ";
Vet True
printObj.Print " eindstand goed";
If inclpnt Then printObj.Print pntEind; "pnt";
Vet False
printObj.Print ", ";
printObj.ForeColor = vbBlue
printObj.Print "aantal doelpunten van de dag goed";
If inclpnt Then printObj.Print pntDp; "pnt"
printObj.ForeColor = 1
printObj.CurrentY = printObj.CurrentY + 50


End Sub

Sub printPointsPerMatch()
'print de deelnemers en hun punten per wedstrijd
Dim rsDeeln As New ADODB.Recordset
Dim rsDeelnPnt As New ADODB.Recordset
Dim rsWeds As New ADODB.Recordset
Dim sqlstr As String
Dim xpos As Integer
Dim posX() As Integer
Dim horPos As Integer
Dim i As Integer
Dim topY As Integer
Dim botY As Integer
Dim topYpos As Integer
Dim columnWidth As Integer
Dim ttlcolumnWidth As Integer
Dim wedstrijd As String
Dim verttxtHeight 'de hoogte van de verticale text bovenin
Dim infostr As String
headerText = GetOrgNaam(poolID) & " " & getKampInfo("toernooi") & " voetbalpool"
header2$ = "Punten t/m wedstrijd " & tillMatch
InitPage False, True
printObj.CurrentY = printObj.CurrentY - 50
topYpos = printObj.CurrentY
deelnemWedsInfo True 'druk de inforegel over de punten toekenning af
topY = printObj.CurrentY
printObj.Line (0, topY)-(printObj.ScaleWidth - 50, topY)
FontGr 8
sqlstr = "SELECT pooldeelnems.deelnemID, pooldeelnems.bijnaam, deelnempnt.grandTotaal"
sqlstr = sqlstr & " FROM (pooldeelnems INNER JOIN deelnempnt ON pooldeelnems.deelnemID = deelnempnt.deelnID) "
sqlstr = sqlstr & " INNER JOIN toernschema ON deelnempnt.wedNum = toernschema.wedNum"
sqlstr = sqlstr & " Where pooldeelnems.poolID = " & poolID
sqlstr = sqlstr & " And toernschema.myNum = " & tillMatch
sqlstr = sqlstr & " And toernschema.ksid = " & kampID
If Me.ScoreVolg(1) = True Then
    sqlstr = sqlstr & " order by grandtotaal DESC"
Else
    sqlstr = sqlstr & " order by bijnaam"
End If

rsDeeln.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
sqlstr = "Select * from qryweds where ksid=" & kampID
'sqlstr = sqlstr & " AND wednum <=" & tillMatch
sqlstr = sqlstr & " order by mynum"
rsWeds.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
verttxtHeight = printObj.TextWidth("1234567890123456789012345")
printObj.CurrentY = verttxtHeight

printObj.CurrentX = printObj.TextWidth(Left(GetLangsteBijNaam, 15))
ReDim posX(1)
posX(1) = printObj.CurrentX
With rsWeds
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            rotate.Angle = 90
            printObj.CurrentX = posX(UBound(posX))
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
            rotate.PrintText !mynum & ": " & wedstrijd
            rotate.Angle = 0
            xpos = printObj.CurrentX + printObj.TextWidth("99") * 1.2
            ReDim Preserve posX(UBound(posX) + 1)
            posX(UBound(posX)) = xpos
            rsWeds.MoveNext
            'Debug.Print UBound(posX), posX(UBound(posX))
        Loop
    End If
End With
rotate.Angle = 90
printObj.CurrentX = posX(UBound(posX))
rotate.PrintText " pnt groepstand"

If getKampInfo("groepen") > 4 Then
    xpos = printObj.CurrentX + printObj.TextWidth("geld") * 1.2
    ReDim Preserve posX(UBound(posX) + 1)
    posX(UBound(posX)) = xpos
    rotate.Angle = 90
    printObj.CurrentX = posX(UBound(posX))
    rotate.PrintText " 8e Finalisten"
End If
xpos = printObj.CurrentX + printObj.TextWidth("99") * 1.2
ReDim Preserve posX(UBound(posX) + 1)
posX(UBound(posX)) = xpos
rotate.Angle = 90
printObj.CurrentX = posX(UBound(posX))
rotate.PrintText " Kw Finalisten"

xpos = printObj.CurrentX + printObj.TextWidth("99") * 1.2
ReDim Preserve posX(UBound(posX) + 1)
posX(UBound(posX)) = xpos
rotate.Angle = 90
printObj.CurrentX = posX(UBound(posX))
rotate.PrintText " Hv Finalisten"

xpos = printObj.CurrentX + printObj.TextWidth("99") * 1.2
ReDim Preserve posX(UBound(posX) + 1)
posX(UBound(posX)) = xpos
rotate.Angle = 90
printObj.CurrentX = posX(UBound(posX))
rotate.PrintText " Finalisten"

xpos = printObj.CurrentX + printObj.TextWidth("99") * 1.2
ReDim Preserve posX(UBound(posX) + 1)
posX(UBound(posX)) = xpos
rotate.Angle = 90
printObj.CurrentX = posX(UBound(posX))
rotate.PrintText " Eindstand"

xpos = printObj.CurrentX + printObj.TextWidth("99") * 1.2
ReDim Preserve posX(UBound(posX) + 1)
posX(UBound(posX)) = xpos
rotate.Angle = 90
printObj.CurrentX = posX(UBound(posX))
rotate.PrintText " Topscorers"

xpos = printObj.CurrentX + printObj.TextWidth("99") * 1.2
ReDim Preserve posX(UBound(posX) + 1)
posX(UBound(posX)) = xpos
rotate.Angle = 90
printObj.CurrentX = posX(UBound(posX))
rotate.PrintText " Overigen"

xpos = printObj.CurrentX + printObj.TextWidth("99") * 1.2
ReDim Preserve posX(UBound(posX) + 1)
posX(UBound(posX)) = xpos
rotate.Angle = 90
printObj.CurrentX = posX(UBound(posX))
rotate.PrintText " Totaal"

xpos = printObj.CurrentX + printObj.TextWidth("999") * 1.2
ReDim Preserve posX(UBound(posX) + 1)
posX(UBound(posX)) = xpos
rotate.Angle = 90
printObj.CurrentX = posX(UBound(posX))
rotate.PrintText " positie"

xpos = printObj.CurrentX + printObj.TextWidth("99") * 1.2
ReDim Preserve posX(UBound(posX) + 1)
posX(UBound(posX)) = xpos
rotate.Angle = 90
printObj.CurrentX = posX(UBound(posX))
printObj.CurrentY = verttxtHeight - printObj.TextHeight("Geld")
'printObj.Print " geld";

xpos = printObj.CurrentX + printObj.TextWidth("geld") * 1.2
printObj.Print
topYpos = printObj.CurrentY + 50
ReDim Preserve posX(UBound(posX) + 1)
posX(UBound(posX)) = xpos
printObj.Line (0, topYpos)-(posX(UBound(posX)), topYpos)
printObj.CurrentY = topYpos
printObj.CurrentX = 0
columnWidth = posX(2) - posX(1)
botY = printObj.CurrentY
pntFormat = "0;;\ ;-"

Do While Not rsDeeln.EOF
Dim naam As String
    naam = rsDeeln!bijnaam
   ' If InStr(naam, "Winner") > 0 Then Stop       1234567890
    Do While printObj.TextWidth(naam) > printObj.TextWidth("123456789012345")
        naam = Left(naam, Len(naam) - 1)
    Loop
    printObj.Print naam;
    sqlstr = "SELECT toernschema.tijd, deelnemPnt.*, toernschema.gespeeld"
    sqlstr = sqlstr & " FROM deelnemPnt INNER JOIN toernschema ON deelnemPnt.wedNum = toernschema.wedNum"
    sqlstr = sqlstr & " Where toernschema.mynum <=" & tillMatch
    sqlstr = sqlstr & " AND toernschema.ksid = " & kampID
    sqlstr = sqlstr & " AND deelnID = " & rsDeeln!deelnemID
    sqlstr = sqlstr & " ORDER BY toernschema.mynum"
    rsDeelnPnt.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
    i = 0
    With rsDeelnPnt
        rotate.Angle = 90
        Do While Not .EOF
            i = i + 1
            printObj.CurrentX = posX(i) + (columnWidth - printObj.TextWidth(Format(nz(!pnttotaal, 0), pntFormat))) / 2
'            rotate.Angle = 0
            'If !pnttotaal = 7 Then Stop
            Ital nz(!pntToto, 0) <> 0
            Vet nz(!pntEind, 0) <> 0
            printObj.FontUnderline = nz(!pntRust, 0) > 0
            If nz(!dpvddag, 0) > 0 Then
                printObj.ForeColor = vbBlue
            End If
            printObj.Print Format(nz(!pnttotaal, 0), pntFormat);
            Vet False
            Ital False
            printObj.FontUnderline = False
            printObj.ForeColor = 1
            
            .MoveNext
            rotate.Angle = 90
        Loop
        If Not .RecordCount = 0 Then
            .MoveLast
            If !postotaal = 1 Then
                printObj.ForeColor = &HC00000
                printObj.FontBold = True
            Else
                printObj.ForeColor = vbBlack
                printObj.FontBold = False
            End If
            ttlcolumnWidth = posX(UBound(posX) - 10) - posX(UBound(posX) - 11)

            printObj.CurrentX = posX(UBound(posX) - 11) + (ttlcolumnWidth - printObj.TextWidth(Format(nz(!pntgrp, 0), pntFormat))) / 2
            printObj.Print Format(nz(!pntgrp, 0), pntFormat);
            ttlcolumnWidth = posX(UBound(posX) - 9) - posX(UBound(posX) - 10)
            If getKampInfo("groepen") > 4 Then
                printObj.CurrentX = posX(UBound(posX) - 10) + (ttlcolumnWidth - printObj.TextWidth(Format(nz(!pnt8fin, 0), pntFormat))) / 2
                printObj.Print Format(nz(!pnt8fin, 0), pntFormat);
                ttlcolumnWidth = posX(UBound(posX) - 8) - posX(UBound(posX) - 9)
            End If
            printObj.CurrentX = posX(UBound(posX) - 9) + (ttlcolumnWidth - printObj.TextWidth(Format(nz(!pntkwfin, 0), pntFormat))) / 2
            printObj.Print Format(nz(!pntkwfin, 0), pntFormat);
            ttlcolumnWidth = posX(UBound(posX) - 7) - posX(UBound(posX) - 8)
            printObj.CurrentX = posX(UBound(posX) - 8) + (ttlcolumnWidth - printObj.TextWidth(Format(nz(!pnthvfin, 0), pntFormat))) / 2
            printObj.Print Format(nz(!pnthvfin, 0), pntFormat);
            ttlcolumnWidth = posX(UBound(posX) - 6) - posX(UBound(posX) - 7)
            printObj.CurrentX = posX(UBound(posX) - 7) + (ttlcolumnWidth - printObj.TextWidth(Format(nz(!pntfin, 0) + nz(!pntklfin, 0), pntFormat))) / 2
            printObj.Print Format(nz(!pntfin, 0) + nz(!pntklfin, 0), pntFormat);
            ttlcolumnWidth = posX(UBound(posX) - 5) - posX(UBound(posX) - 6)
            printObj.CurrentX = posX(UBound(posX) - 6) + (ttlcolumnWidth - printObj.TextWidth(Format(!pntuitslnaklfin + !pntuitsl, pntFormat))) / 2
            printObj.Print Format(!pntuitslnaklfin + !pntuitsl, pntFormat);
            ttlcolumnWidth = posX(UBound(posX) - 4) - posX(UBound(posX) - 5)
            printObj.CurrentX = posX(UBound(posX) - 5) + (ttlcolumnWidth - printObj.TextWidth(Format(nz(!pntTopsc, 0) + nz(!pntOverig, 0), pntFormat))) / 2
            printObj.Print Format(nz(!pntTopsc, 0), pntFormat);
            ttlcolumnWidth = posX(UBound(posX) - 3) - posX(UBound(posX) - 4)
            printObj.CurrentX = posX(UBound(posX) - 4) + (ttlcolumnWidth - printObj.TextWidth(Format(nz(!pntTopsc, 0) + nz(!pntOverig, 0), pntFormat))) / 2
            printObj.Print Format(nz(!pntOverig, 0), pntFormat);
            ttlcolumnWidth = posX(UBound(posX) - 2) - posX(UBound(posX) - 3)
            printObj.CurrentX = posX(UBound(posX) - 3) + (ttlcolumnWidth - printObj.TextWidth(Format(nz(!grandtotaal, 0), pntFormat))) / 2
            printObj.Print Format(nz(!grandtotaal, 0), pntFormat);
            ttlcolumnWidth = posX(UBound(posX) - 1) - posX(UBound(posX) - 2)
            printObj.CurrentX = posX(UBound(posX) - 2) + (ttlcolumnWidth - printObj.TextWidth(Format(nz(!postotaal, 0), pntFormat))) / 2
            printObj.Print Format(nz(!postotaal, 0), pntFormat);
            printObj.CurrentX = posX(UBound(posX)) - printObj.TextWidth(Format(nz(!geldttl, 0), "currency"))
            printObj.ForeColor = vbBlack
            printObj.FontItalic = False
            printObj.FontBold = False
'            printObj.Print Format(nz(!geldttl, 0), "currency");
        End If
        printObj.Print
    End With
    printObj.Line (0, printObj.CurrentY + 10)-(posX(UBound(posX)), printObj.CurrentY + 10)
    printObj.CurrentY = printObj.CurrentY + 10
    printObj.CurrentX = 0
    botY = printObj.CurrentY
'    If rsDeeln.AbsolutePosition = 67 Then Stop
    If botY >= footerHeight And rsDeeln.AbsolutePosition < rsDeeln.RecordCount Then
        'nieuwe pagina
        'eerste de lijntjes
        For i = 1 To UBound(posX)
            printObj.Line (posX(i), topY)-(posX(i), botY)
        Next
        i = 0
        DoNewPage False, True
        printObj.CurrentY = printObj.CurrentY - 50
        topYpos = printObj.CurrentY
        deelnemWedsInfo True 'druk de inforegel over de punten toekenning af
        topY = printObj.CurrentY
        printObj.Line (0, topY)-(printObj.ScaleWidth - 50, topY)
        FontGr 8
        printObj.CurrentY = verttxtHeight
        printObj.CurrentX = printObj.TextWidth("123456789012345")
        With rsWeds
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    Set rotate.Device = printObj
                    i = i + 1
                    rotate.Angle = 90
                    printObj.CurrentX = posX(i)
                    If !tm1 > "" Then
                        rotate.PrintText !mynum & ": " & !tm1 & "-" & !tm2
                    Else
                        rotate.PrintText !mynum & ": " & !code1 & "-" & !code2
                    End If
                    rotate.Angle = 0
                    .MoveNext
                Loop
            End If
        End With
        rotate.Angle = 90
        If getKampInfo("groepen") > 4 Then
            horPos = 11
        Else
            horPos = 10
        End If
        printObj.CurrentX = posX(UBound(posX) - horPos)
        horPos = horPos - 1
        rotate.PrintText " pnt groepstand"
        If getKampInfo("groepen") > 4 Then
            printObj.CurrentX = posX(UBound(posX) - horPos)
            horPos = horPos - 1
            rotate.PrintText " 8e Finalisten"
        End If
        printObj.CurrentX = posX(UBound(posX) - horPos)
        horPos = horPos - 1
        rotate.PrintText " Kw Finalisten"
        printObj.CurrentX = posX(UBound(posX) - horPos)
        horPos = horPos - 1
        rotate.PrintText " Hv Finalisten"
        printObj.CurrentX = posX(UBound(posX) - horPos)
        horPos = horPos - 1
        rotate.PrintText " Finalisten"
        printObj.CurrentX = posX(UBound(posX) - horPos)
        horPos = horPos - 1
        rotate.PrintText " Eindstand"
        printObj.CurrentX = posX(UBound(posX) - horPos)
        horPos = horPos - 1
        rotate.PrintText " Topscorers"
        printObj.CurrentX = posX(UBound(posX) - horPos)
        horPos = horPos - 1
        rotate.PrintText " Overigen"
        printObj.CurrentX = posX(UBound(posX) - horPos)
        horPos = horPos - 1
        rotate.PrintText " Totaal"
        printObj.CurrentX = posX(UBound(posX) - horPos)
        horPos = horPos - 1
        rotate.PrintText " positie"
        printObj.CurrentX = posX(UBound(posX) - horPos)
        printObj.CurrentY = verttxtHeight ' - printObj.TextHeight("Geld")
 '       printObj.Print " geld"
        topYpos = printObj.CurrentY + 50
        printObj.Line (0, topYpos)-(posX(UBound(posX)), topYpos)
        printObj.CurrentY = topYpos
        printObj.CurrentX = 0
        i = i + 1
    End If
    rsDeeln.MoveNext
    rsDeelnPnt.Close
Loop
rsDeeln.Close
For i = 1 To UBound(posX)
    printObj.Line (posX(i), topY)-(posX(i), botY)
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
Dim columnWidth As Integer
Dim ttlcolumnWidth As Integer
Dim verttxtHeight 'de hoogte van de verticale text bovenin
Dim infostr As String
headerText = GetOrgNaam(poolID) & " " & getKampInfo("toernooi") & " voetbalpool"
header2$ = "Positie in de pool na elke wedstrijd"
InitPage False, True
printObj.CurrentY = printObj.CurrentY - 50
topYpos = printObj.CurrentY
deelnemWedsInfo False 'druk de inforegel over de punten toekenning af
topY = printObj.CurrentY
printObj.Line (0, topY)-(printObj.ScaleWidth - 50, topY)
FontGr 8
sqlstr = "SELECT pooldeelnems.deelnemID, pooldeelnems.bijnaam, deelnempnt.grandTotaal"
sqlstr = sqlstr & " FROM (pooldeelnems INNER JOIN deelnempnt ON pooldeelnems.deelnemID = deelnempnt.deelnID) "
sqlstr = sqlstr & " INNER JOIN toernschema ON deelnempnt.wedNum = toernschema.wedNum"
sqlstr = sqlstr & " Where pooldeelnems.poolID = " & poolID
sqlstr = sqlstr & " And toernschema.myNum = " & tillMatch
sqlstr = sqlstr & " And toernschema.ksid = " & kampID
If Me.ScoreVolg(1) = True Then
    sqlstr = sqlstr & " order by grandtotaal DESC"
Else
    sqlstr = sqlstr & " order by bijnaam"
End If

rsDeeln.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
sqlstr = "Select * from qryweds where ksid=" & kampID
'sqlstr = sqlstr & " AND wednum <=" & tillMatch
sqlstr = sqlstr & " order by mynum"
rsWeds.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
verttxtHeight = printObj.TextWidth("1234567890123456789012345")
printObj.CurrentY = verttxtHeight
printObj.CurrentX = printObj.TextWidth("1234567890")
ReDim posX(1)
posX(1) = printObj.CurrentX
With rsWeds
    Do While Not .EOF
        rotate.Angle = 90
        printObj.CurrentX = posX(UBound(posX))
        If !tm1 > "" Then
            rotate.PrintText !mynum & ": " & !tm1 & "-" & !tm2
        Else
            rotate.PrintText !mynum & ": " & !code1 & "-" & !code2
        End If
        rotate.Angle = 0
        xpos = printObj.CurrentX + printObj.TextWidth("99") * 1.3
        ReDim Preserve posX(UBound(posX) + 1)
        posX(UBound(posX)) = xpos
        .MoveNext
    Loop
End With

'printObj.Print
topYpos = printObj.CurrentY + 50
ReDim Preserve posX(UBound(posX) + 1)
posX(UBound(posX)) = xpos
printObj.Line (0, topYpos)-(posX(UBound(posX)), topYpos)
printObj.CurrentY = topYpos
printObj.CurrentX = 0
columnWidth = posX(2) - posX(1)
botY = printObj.CurrentY
pntFormat = "0;;\ ;-"

Do While Not rsDeeln.EOF
    printObj.Print rsDeeln!bijnaam;
    sqlstr = "SELECT toernschema.tijd, deelnemPnt.*, toernschema.gespeeld"
    sqlstr = sqlstr & " FROM deelnemPnt INNER JOIN toernschema ON deelnemPnt.wedNum = toernschema.wedNum"
    sqlstr = sqlstr & " Where toernschema.mynum <=" & tillMatch
    sqlstr = sqlstr & " AND toernschema.ksid = " & kampID
    sqlstr = sqlstr & " AND deelnID = " & rsDeeln!deelnemID
    sqlstr = sqlstr & " ORDER BY toernschema.mynum"
    rsDeelnPnt.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
    i = 0
    With rsDeelnPnt
        rotate.Angle = 90
        Do While Not .EOF
            i = i + 1
            printObj.CurrentX = posX(i) + (columnWidth - printObj.TextWidth(Format(nz(!postotaal, 0), pntFormat))) / 2
            Ital nz(!pntToto, 0) <> 0
            Vet nz(!pntEind, 0) <> 0
            printObj.FontUnderline = nz(!pntRust, 0) > 0
            If nz(!dpvddag, 0) > 0 Then
                printObj.ForeColor = vbBlue
            End If
            printObj.Print Format(nz(!postotaal, 0), pntFormat);
            Vet False
            Ital False
            printObj.FontUnderline = False
            printObj.ForeColor = 1
            
            .MoveNext
            rotate.Angle = 90
        Loop
        printObj.Print
    End With
    printObj.Line (0, printObj.CurrentY + 10)-(posX(UBound(posX)), printObj.CurrentY + 10)
    printObj.CurrentY = printObj.CurrentY + 10
    printObj.CurrentX = 0
    botY = printObj.CurrentY
    If botY >= footerHeight Then
        'nieuwe pagina
        'eerste de lijntjes
        For i = 1 To UBound(posX)
            printObj.Line (posX(i), topY)-(posX(i), botY)
        Next
        i = 0
        DoNewPage False, True
        printObj.CurrentY = printObj.CurrentY - 50
        topYpos = printObj.CurrentY
        deelnemWedsInfo False 'druk de inforegel over de punten toekenning af
        topY = printObj.CurrentY
        printObj.Line (0, topY)-(printObj.ScaleWidth - 50, topY)
        FontGr 8
        printObj.CurrentY = verttxtHeight
        printObj.CurrentX = printObj.TextWidth("123456789012345")
        With rsWeds
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    Set rotate.Device = printObj
                    i = i + 1
                    rotate.Angle = 90
                    printObj.CurrentX = posX(i)
                    If !tm1 > "" Then
                        rotate.PrintText !mynum & ": " & !tm1 & "-" & !tm2
                    Else
                        rotate.PrintText !mynum & ": " & !code1 & "-" & !code2
                    End If
                    rotate.Angle = 0
                    .MoveNext
                Loop
            End If
        End With
        'printObj.Print
        topYpos = printObj.CurrentY + 50
        printObj.Line (0, topYpos)-(posX(UBound(posX)), topYpos)
        printObj.CurrentY = topYpos
        printObj.CurrentX = 0
        i = i + 1
    End If
    rsDeeln.MoveNext
Loop
For i = 1 To UBound(posX)
    printObj.Line (posX(i), topY)-(posX(i), botY)
Next
i = 0


End Sub

Sub printMatchPredictions(wedNum As Integer)
Dim sqlstr As String
Dim rs As New ADODB.Recordset
Dim rsDeeln As New ADODB.Recordset
Dim cloneRS As ADODB.Recordset
Dim zoekstr As String
Dim header2je As String
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
Dim header2pos(3) As Integer
Dim col As Integer
Dim i As Integer
wedNum = GetWedNum(wedNum)
    headerText = GetOrgNaam(poolID) & " " & getKampInfo("toernooi") & " voetbalpool" & " - Voorspelling"
    If Not Me.optPortrait Then
        cols(0) = 0
        cols(1) = printObj.ScaleWidth / 4
        cols(2) = printObj.ScaleWidth / 2
        cols(3) = printObj.ScaleWidth / 4 * 3
        cols(4) = printObj.ScaleWidth
        col = 4
    Else
        cols(0) = 0
        cols(1) = printObj.ScaleWidth / 3
        cols(2) = printObj.ScaleWidth / 3 * 2
        cols(3) = printObj.ScaleWidth
        cols(4) = printObj.ScaleWidth
        col = 3
    End If
    header2je = Format(GetWedInfo(wedNum, "datum"), "ddd d mmm") & " "
    header2je = header2je & Format(GetWedInfo(wedNum, "tijd"), "HH:MM") & ": "
    header2je = header2je & GetWedInfo(wedNum, "naam1") & " vs " & GetWedInfo(wedNum, "naam2")
    header2$ = "Wedstrijd " & GetMyNum(wedNum) & ": " & header2je
    InitPage False, True
    
    printObj.Print
    header2pos(0) = 50
    header2pos(1) = printObj.TextWidth("0-000")
    header2pos(2) = header2pos(1) + printObj.TextWidth("0-000")
    header2pos(3) = header2pos(2) + printObj.TextWidth("0-000")
    printObj.ForeColor = RGB(0, 51, 0)
    For i = 0 To col - 1
        printObj.CurrentX = cols(i) + header2pos(0)
        printObj.Print "Rust";
        printObj.CurrentX = cols(i) + header2pos(1)
        printObj.Print "Eind";
        printObj.CurrentX = cols(i) + header2pos(2)
        printObj.Print "Toto";
        printObj.CurrentX = cols(i) + header2pos(3)
        printObj.Print "Wie";
    Next
    printObj.ForeColor = 0
    printObj.Print
    yStart = printObj.CurrentY
    sqlstr = "SELECT e1, e2, r1,r2,toto, wednum "
    sqlstr = sqlstr & " FROM voorspelling_uitsl INNER JOIN "
    sqlstr = sqlstr & " pooldeelnems ON voorspelling_uitsl.deelnem = pooldeelnems.deelnemID"
    sqlstr = sqlstr & " GROUP BY e1, e2, r1, r2, toto, wednum, poolid"
    sqlstr = sqlstr & " HAVING wednum=" & wedNum
    sqlstr = sqlstr & " AND pooldeelnems.poolid= " & poolID
    sqlstr = sqlstr & " ORDER BY r1,r2,e1,e2,toto"
    rs.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
    sqlstr = "SELECT e1, e2, r1,r2,toto, wednum, bijnaam "
    sqlstr = sqlstr & " FROM voorspelling_uitsl INNER JOIN "
    sqlstr = sqlstr & " pooldeelnems ON voorspelling_uitsl.deelnem = pooldeelnems.deelnemID"
    sqlstr = sqlstr & " WHERE wednum = " & wedNum
    sqlstr = sqlstr & " AND poolid = " & poolID
    sqlstr = sqlstr & " ORDER BY bijnaam"
    rsDeeln.Open sqlstr, dbConn, adOpenStatic, adLockReadOnly
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
        printObj.CurrentX = cols(i)
        lineXstart = printObj.CurrentX
        lineYstart = printObj.CurrentY
        printObj.CurrentX = cols(i) + header2pos(0)
        printObj.Print rs!r1 & "-" & rs!r2;
        printObj.CurrentX = cols(i) + header2pos(1)
        Vet True
        printObj.Print rs!e1 & "-" & rs!e2;
        Vet False
        printObj.CurrentX = cols(i) + header2pos(2)
        printObj.Print rs!toto;
        cloneRS.MoveFirst
        Do While Not cloneRS.EOF
            printObj.CurrentX = cols(i) + header2pos(3)
            printObj.Print cloneRS!bijnaam
            rijnu = rijnu + 1
            cloneRS.MoveNext
        Loop
        lineXend = cols(i + 1) - 100
        lineYend = printObj.CurrentY
        printObj.Line (lineXstart, lineYstart)-(lineXend, lineYend), , B
        rs.MoveNext
        If rijnu >= rijen Then
            i = i + 1
            printObj.CurrentY = yStart
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
        printObj.ForeColor = vbWhite
    Else
        printObj.ForeColor = vbBlack
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
    prntColor(a) = klCol(i)
    klCol.Remove i
Next
End Sub
