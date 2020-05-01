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
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   6240
      TabIndex        =   4
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label lblStartText 
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   720
      TabIndex        =   3
      Tag             =   "kop"
      Top             =   1080
      Width           =   7095
   End
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
      Top             =   4440
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
      Height          =   855
      Left            =   720
      TabIndex        =   1
      Tag             =   "kop1"
      Top             =   3240
      Visible         =   0   'False
      Width           =   7215
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
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Tag             =   "kop2"
      Top             =   0
      Width           =   8235
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
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuNewTournament 
         Caption         =   "Nieuw &Toernooi"
         Enabled         =   0   'False
         Visible         =   0   'False
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
      Begin VB.Menu mnuTournamentData 
         Caption         =   "&Gegevens"
      End
      Begin VB.Menu mnuTournamentTeams 
         Caption         =   "&Ploegen"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTournamentSchedule 
         Caption         =   "&Wedstrijdschema"
      End
   End
   Begin VB.Menu mnuWedstrijd 
      Caption         =   "&Wedstrijd"
      Begin VB.Menu mnuMatchOverview 
         Caption         =   "&Overzicht"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Opties"
      Begin VB.Menu mnuStartOver 
         Caption         =   "&Gegevens inlezen"
      End
      Begin VB.Menu mnuOptionsPointTypes 
         Caption         =   "&Voorspelling types"
      End
      Begin VB.Menu mnuAdmin 
         Caption         =   "&Organisatie"
      End
      Begin VB.Menu mnuDblPlayers 
         Caption         =   "Remove Double Players"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuConvert 
         Caption         =   "Convert Tournamentschedule table"
         Enabled         =   0   'False
         Visible         =   0   'False
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

Dim cn As ADODB.Connection

Dim startState As Integer

Function msg1()
Dim msg As String
    msg = "Welkom bij Jota's Voetbalpool"
    msg = msg & vbNewLine
    msg = msg & "We konden nog geen gegevens vinden in het systeem."
    msg = msg & vbNewLine
    msg = msg & "Vul eerst het volgende formulier in. "
    msg = msg & "De gegevens worden gebuikt bij de verschillende afdrukken,"
    msg = msg & vbNewLine
    msg = msg & "dus maak er geen zootje van ;-)"
    msg1 = msg
End Function

Function msg2()
Dim msg As String
    msg = msg & "Dank voor het invullen."
    msg = msg & vbNewLine
    msg = msg & "We gaan nu de gegevens van het laatst bekende toernooi van de server halen"
    msg = msg & " en vullen dan de Voetbalpool met standaard instellingen, "
    msg = msg & "die je later natuurlijk kunt aanpassen."
    msg = msg & vbNewLine & vbNewLine
    msg = msg & "Klik op OK en dan een ogenblik geduld, zo gebeurd..."
    msg2 = msg
End Function

Function msg3()
Dim msg As String
    msg = msg & "Klaar!"
    msg = msg & vbNewLine
    msg = msg & "Je kunt nu in het menu 'Pool' de naam van deze pool "
    msg = msg & "en de puntentoekenning aanpassen."
    msg = msg & vbNewLine
    msg = msg & "Als je daarmee klaar bent kun je via het menu"
    msg = msg & " 'Bestand - Print' "
    msg = msg & "de poolformulieren afdrukken."
    msg = msg & vbNewLine & vbNewLine
    msg = msg & "Veel plezier met Jota's Voetbalpool!"
    msg3 = msg
End Function

Sub firstStart()
Dim msg As String
    If thisPool = 0 Then
        ''get organisation data
         frmOrganisation.Show 1
        ''get tournament data
        DoEvents
        Me.lblStartText = msg2
        msg = "Welkom bij Jota's Voetbalpool"
        msg = msg & vbNewLine & vbNewLine
        MsgBox msg, vbOKOnly + vbInformation, "Nieuwe start"
        DoEvents
        'copy the tournament data
        getTournamentTables
        ''fill tables with default values
        fillDefaultValues
        '
        DoEvents
        MsgBox msg, vbOKOnly + vbInformation, "Nieuwe start"
        DoEvents
    End If
    DoEvents 'why not
    updateForm
End Sub

Private Sub btnOk_Click()
    If startState = 1 Then
        frmOrganisation.Show 1
        DoEvents
        startState = 2
        Me.lblStartText = msg2
        Exit Sub
    End If
    If startState = 2 Then
        getTournamentTables
        ''fill tables with default values
        fillDefaultValues
        startState = 3
        Me.lblStartText.Alignment = 2
        Me.lblStartText = msg3
        Exit Sub
    End If
    If startState = 3 Then
        updateForm
    End If
End Sub

Private Sub Form_Load()
    Dim msg As String
'open db connection
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = lclConn()
        .Open
    End With
    'set Form defaults
    'size form half the screen size
    Me.Width = Screen.Width / 2
    Me.Height = Screen.Height / 2
    write2Log "Main form opened", True
    
    If thisPool = 0 Then
        startState = 1
        Me.lblStartText.Visible = True
        Me.lblStartText.Caption = msg1
    End If
    updateForm
    centerForm Me
    UnifyForm Me
    
End Sub

Sub updateForm()
    Me.lblStartText.Visible = thisPool = 0
    Me.btnOk.Visible = thisPool = 0
    
    Me.mnuFileOpen.Enabled = recordsExist("tblPools", cn)
    Me.mnuPrint.Enabled = thisPool > 0
    Me.mnuEditPool.Enabled = thisPool > 0
    
    Me.mnuNewPool.Enabled = recordsExist("tblTournaments", cn)
    Me.mnuPoolCompetitors.Enabled = thisPool > 0
    Me.mnuDblPlayers.Visible = adminLogin 'just for admin
    Me.mnuConvert.Visible = adminLogin 'just for admin
    
    Me.mnuEditTournaments.Visible = adminLogin
    Me.mnuNewTournament.Visible = adminLogin
    Me.mnuOptionsPointTypes.Visible = adminLogin
    Me.mnuTournamentData.Visible = True
    Me.mnuTournamentSchedule.Visible = adminLogin
    Me.mnuTournamentTeams.Visible = adminLogin
    
    Me.Caption = "Jota's Voetbalpool 2.0"
    DoEvents
    If thisPool Then
        
        With Me.lblStartTitle
            .Caption = getOrganisation(cn)
            .BackColor = Me.BackColor
            .BackStyle = 1
        End With
        With Me.lblPoolName
            .Caption = getPoolInfo("poolName", cn)
            .Visible = True
            .BackColor = Me.BackColor
            .BackStyle = 1
            .Refresh
        End With
        
    Else
        Me.lblStartTitle.Caption = "Jota's Voetbalpool - geen pool geselecteerd"
        Me.lblPoolName.Visible = False
    End If
    Me.lblCopyright = "© 2004 - " & Year(Now) & " jota services"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim objForm As Form
    
    If Not cn Is Nothing Then
        If (cn.State And adStateOpen) = adStateOpen Then
            cn.Close
        End If
        Set cn = Nothing
    End If
    
    For Each objForm In Forms
        If objForm.Name <> Me.Name Then
            Unload objForm
            Set objForm = Nothing
        End If
    Next
    write2Log "App ended", True
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
    With Me.lblStartText
        .Left = Me.Shape1(0).Width
        .Width = Me.lblPoolName.Width
        .Top = (windowH / 2) - (.Height / 2)
    End With
    With Me.btnOk
        .Top = Me.lblStartText.Top + Me.lblStartText.Height + 20
        .Left = Me.Shape1(1).Left - .Width
    End With
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuAdmin_Click()
'open the organisation form
    frmOrganisation.Show 1  '
'    If Not recordsExist("tblOrganisation", cn) Then
'        frmOrganisation.Show 1  'there is no organisation yet
'    Else
'        adminLogin = DoLogin
'        If Not adminLogin Then
'            MsgBox "Admin rechten niet verkregen", vbOKOnly + vbExclamation, "Admin"
'        End If
'        updateForm
'    End If
End Sub

Private Sub mnuConvert_Click()
    convertTournamentScheduleTable
    write2Log "Conversion attempted", True
End Sub

Private Sub mnuDblPlayers_Click()
    'frmRemoveDoubleIds.Show 1
End Sub

Private Sub mnuExitApp_Click()
    Unload Me
End Sub

Private Sub mnuFileOpen_Click()
    openPool.Show 1
    updateForm
    write2Log "Pool opened", True

End Sub

Private Sub mnuMatchOverview_Click()
    matchlistForm.Show 1
End Sub

Private Sub mnuNewPool_Click()
    newPoolForm.Show 1
    updateForm
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

Private Sub mnuStartOver_Click()
    Dim msg As String
    msg = "Hiermee kun je de gegevens van het huidige toernooi (opnieuw) inlezen."
    msg = msg & vbNewLine & "Alle door jou toegevoegde gegevens blijven onveranderd."
    msg = msg & vbNewLine & "Zorg dat je een werkende internet verbinding hebt,"
    msg = msg & vbNewLine & "anders kan het niet"
    msg = msg & vbNewLine & vbNewLine & "Druk op OK als je het zeker weet of anders op Annuleren"
    If MsgBox(msg, vbOKCancel, "Data inlezen") = vbOK Then
        frmCopyData.Show 1
    End If
    updateForm
End Sub

Private Sub mnuTournamentData_Click()
    tournamentsForm.Show 1
End Sub

Private Sub mnuTournamentSchedule_Click()
      matchlistForm.Show 1
End Sub

Private Sub mnuTournamentTeams_Click()
    teamsForm.Show 1
End Sub

