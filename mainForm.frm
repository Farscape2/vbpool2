VERSION 5.00
Begin VB.Form mainForm 
   BackColor       =   &H00B2EDB0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Voetbalpool"
   ClientHeight    =   4800
   ClientLeft      =   9555
   ClientTop       =   5730
   ClientWidth     =   8445
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
   Icon            =   "mainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "©2004 - 2020 jota services"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3645
      TabIndex        =   2
      Tag             =   "small"
      Top             =   4320
      Width           =   1845
   End
   Begin VB.Label lblPoolName 
      Alignment       =   2  'Center
      BackColor       =   &H00B2EDB0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Tag             =   "kop1"
      Top             =   3480
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.Label lblStartTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00B2EDB0&
      Caption         =   "Voetbalpool"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Tag             =   "kop2"
      Top             =   0
      Width           =   7155
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFC0&
      FillColor       =   &H00B2EDB0&
      Height          =   1815
      Index           =   1
      Left            =   7920
      Top             =   1320
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFC0&
      FillColor       =   &H00B2EDB0&
      Height          =   1815
      Index           =   0
      Left            =   -10
      Top             =   1320
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0FFC0&
      Height          =   1600
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   1440
      Width           =   1600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFC0&
      X1              =   4440
      X2              =   4440
      Y1              =   840
      Y2              =   4440
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Bestand"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open Pool"
      End
      Begin VB.Menu mnuNewPool 
         Caption         =   "&Nieuwe Pool"
      End
      Begin VB.Menu mnuNewTournament 
         Caption         =   "Nieuw &Toernooi"
      End
      Begin VB.Menu mnuSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Af&drukken"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuExitApp 
         Caption         =   "&Afsluiten"
      End
   End
   Begin VB.Menu mnuEditPool 
      Caption         =   "&Pool"
      Begin VB.Menu mnuPoolBasicData 
         Caption         =   "&Basis gegevens"
      End
      Begin VB.Menu mnuPoolSettings 
         Caption         =   "&Instelingen"
      End
      Begin VB.Menu mnuPoolCompetitors 
         Caption         =   "&Deelnemers"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuEditTournaments 
      Caption         =   "&Toernooi"
      Begin VB.Menu mnuTouramentData 
         Caption         =   "&Gegevens"
      End
      Begin VB.Menu mnuTournamentTeams 
         Caption         =   "&Ploegen"
      End
      Begin VB.Menu mnuTournamentSchedule 
         Caption         =   "&Wedstrijdschema"
      End
   End
   Begin VB.Menu mnuWedstrijd 
      Caption         =   "&Wedstrijd"
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Opties"
      Begin VB.Menu mnuOptionsPointTypes 
         Caption         =   "&Voorspelling types"
      End
      Begin VB.Menu mnuDblPlayers 
         Caption         =   "Remove Double Players"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&Over"
      End
   End
End
Attribute VB_Name = "mainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

'set Form defaults

    centerForm Me
    UnifyForm Me
    
    updateForm
    
End Sub

Sub updateForm()
    Me.mnuPrint.Enabled = thisPool
    Me.mnuEditPool.Enabled = thisPool
    'me.mnuEditTournaments
    Me.mnuPoolCompetitors.Enabled = thisPool
    Me.mnuDblPlayers.Visible = False 'just for admin
    Me.Caption = "Jota's Voetbalpool"
    DoEvents
    If thisPool Then
        
        With Me.lblStartTitle
            .Caption = getOrganisation()
            .BackColor = Me.BackColor
            .BackStyle = 1
        End With
        With Me.lblPoolName
            .Caption = getPoolInfo("poolName")
            .Visible = True
            .BackColor = Me.BackColor
            .BackStyle = 1
            .Refresh
        End With
        
    Else
        Me.lblStartTitle.Caption = "Voetbalpool - geen pool geselecteerd"
        Me.lblPoolName.Visible = False
        Me.Caption = "Jota's Voetbalpool"
    End If
    Me.lblCopyright = "© 2004 - " & Year(Now) & " jota computer assistentie"
End Sub

Private Sub Form_Resize()
'set graphics right
'middle line
Dim windowW As Integer 'window width
Dim windowH As Integer 'window height
    If Me.Width < 12000 Then Me.Width = 12000
    If Me.Height < 7600 Then Me.Height = 7600
    windowH = Me.ScaleHeight
    windowW = Me.ScaleWidth
    With Me.Line1
        .X1 = windowW / 2
        .Y1 = 0
        .X2 = .X1
        .Y2 = windowH
    End With
    With Me.Shape1(0)
        .Height = windowH / 2
        .Width = .Height / 2.2
        .Top = (windowH / 2) - (.Height / 2)
        .Left = -10
    End With
    With Me.Shape1(1)
        .Height = Me.Shape1(0).Height
        .Width = Me.Shape1(0).Width
        .Top = Me.Shape1(0).Top
        .Left = windowW - .Width + 10
    End With
    With Me.Shape2
        .Height = windowH / 3
        .Width = .Height
        .Left = (windowW / 2) - (.Width / 2)
        .Top = (windowH / 2) - (.Height / 2)
    End With
    With Me.lblStartTitle
        .Width = windowW
        .Top = 250
        .Left = 0
    End With
    With Me.lblPoolName
        .Width = windowW - Me.Shape1(0).Width * 2 - 30
        .Left = Me.Shape1(0).Width + 20
        .Top = (windowH / 2) - (.Height / 2)
    End With
    With Me.lblCopyright
        .Left = windowW - .Width - 120
        .Top = windowH - .Height - 60
    End With
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuDblPlayers_Click()
    frmRemoveDoubleIds.Show 1
End Sub

Private Sub mnuExitApp_Click()
    Dim objForm As Form
    For Each objForm In Forms
        If objForm.Name <> Me.Name Then
            Unload objForm
            Set objForm = Nothing
        End If
    Next
    Unload Me
End Sub

Private Sub mnuFileOpen_Click()
    openPool.Show 1
    updateForm
End Sub

Private Sub mnuNewPool_Click()
    newPoolForm.Show 1
End Sub

Private Sub mnuOptionsPointTypes_Click()
    pointTypes.Show 1
End Sub

Private Sub mnuPoolBasicData_Click()
    poolsForm.Show 1
    DoEvents
    updateForm
End Sub

Private Sub mnuPoolSettings_Click()
    poolPointsForm.Show 1
End Sub

Private Sub mnuTouramentData_Click()
    tournamentsForm.Show 1
End Sub

Private Sub mnuTournamentSchedule_Click()
      matchlistForm.Show 1
End Sub

Private Sub mnuTournamentTeams_Click()
    teamsForm.Show 1
End Sub
