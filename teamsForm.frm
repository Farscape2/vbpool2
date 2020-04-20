VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form teamsForm 
   Caption         =   "Ploegen en spelers"
   ClientHeight    =   1845
   ClientLeft      =   12030
   ClientTop       =   3765
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   5055
   Begin VB.CommandButton btnPlayers 
      Caption         =   "Spelers"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo cmbTeams 
      DataField       =   "teamID"
      DataSource      =   "Adodc1"
      Height          =   360
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Tag             =   "A1"
      Top             =   600
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Ok"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblPoolName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pool A"
      Height          =   360
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Tag             =   "kop"
      Top             =   240
      Width           =   2100
   End
   Begin VB.Label lblPoolNr 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   375
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   600
      Width           =   200
   End
End
Attribute VB_Name = "teamsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub makeForm()

Dim rs As ADODB.Recordset
Dim rsTeamCodes As ADODB.Recordset
Dim sqlstr As String
Dim groups As Integer
Dim teams As Integer
Dim row As Integer, col As Integer
Dim grpCounter As Integer
Dim counter As Integer
Dim groupSize As Integer
Dim grp As Integer
    
    Set rs = New ADODB.Recordset
    Set rsTeamCodes = New ADODB.Recordset
    'check if the base schedule is made
    
    If Not tournamentHasSchedule Then
        generateSchedule
    End If
    
    'fill combobox with teamnames
    
    sqlstr = "Select teamNameId, TeamName, teamShortname, teamType from tblTeamNames "
    If getTournamentInfo("tournamentType") = "EK" Then
        sqlstr = sqlstr & "Where teamtype <= 1"
    End If
    If getTournamentInfo("tournamentType") = "CL" Then
        sqlstr = sqlstr & "Where teamtype > 2"
    End If
    sqlstr = sqlstr & " order by teamname "
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If rs.EOF Then Exit Sub
    
    sqlstr = "Select * from tblTournamentTeamCodes where tournamentid = " & thisTournament
    rsTeamCodes.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
    
    groups = getTournamentInfo("tournamentGroupCount")
    teams = getTournamentInfo("tournamentTeamCount")
    groupSize = teams / groups
    counter = 0
    grp = 0
    For row = 1 To groups / 2
        For col = 1 To 2
            grp = grp + 1
            If lblPoolName.Count < grp Then
                Load lblPoolName(grp - 1)
                lblPoolName(grp - 1).Visible = True
            End If
            With Me.lblPoolName(grp - 1)
                .Caption = "Pool " & Chr(64 + grp)
                .Width = 2100
                .Height = 360
                .Top = 240 + (row - 1) * (.Height + 100 + (groupSize * 360))
                .Left = 300 + (col - 1) * 2200
            End With
            For grpCounter = 1 To groupSize
                counter = counter + 1
                If lblPoolNr.Count < counter Then
                    Load Me.lblPoolNr(counter - 1)
                    Load Me.cmbTeams(counter - 1)
                    Me.lblPoolNr(counter - 1).Visible = True
                    Me.cmbTeams(counter - 1).Visible = True
                    Me.cmbTeams(counter - 1).TabIndex = counter
                End If
                With Me.lblPoolNr(counter - 1)
                    .Caption = grpCounter
                    .Left = 300 + (col - 1) * 2200
                    .Top = 600 + (grpCounter - 1) * 360 + (row - 1) * (.Height + 100 + (groupSize * 360))
                End With
                With Me.cmbTeams(counter - 1)
                    Set .RowSource = rs
                    .ListField = "teamname"
                    .BoundColumn = "teamNameId"
                    .Left = 600 + (col - 1) * 2200
                    .Width = 1800
                    .Top = Me.lblPoolNr(counter - 1).Top
                    'Add tag to find record later in tblTournamentTeamCodes
                    .Tag = Chr(64 + grp) & Format(grpCounter, "0")
                    'if teamId exists in table, show team
                    rsTeamCodes.MoveFirst
                    rsTeamCodes.Find "teamcode = '" & .Tag & "'"
                    .BoundText = Nz(rsTeamCodes!teamId, 0)
                End With
            Next
        Next
        Me.Height = (Me.Height - Me.ScaleHeight) + 640 + row * (groupSize + 1) * 360
    Next
    'ruimte voor knoppen
    Me.Height = Me.Height + Me.btnClose.Height + 240
    Me.btnClose.Top = Me.ScaleHeight - Me.btnClose.Height - 180
    Me.btnPlayers.Top = Me.btnClose.Top
    Me.btnPlayers.Left = Me.btnClose.Left - Me.btnPlayers.Width - 50
    If (rs.State And adStateOpen) = adStateOpen Then rs.Close
    If (rsTeamCodes.State And adStateOpen) = adStateOpen Then rsTeamCodes.Close
    Set rs = Nothing
    Set rsTeamCodes = Nothing
End Sub

Private Sub btnClose_Click()
    Dim ctl As Control
    For Each ctl In lblPoolName
        If ctl.Index <> 0 Then
            Unload ctl
        End If
    Next
    For Each ctl In cmbTeams
        If ctl.Index <> 0 Then
            Unload ctl
        End If
    Next
    For Each ctl In lblPoolNr
        If ctl.Index <> 0 Then
            Unload ctl
        End If
    Next
    Set ctl = Nothing
    Unload Me
End Sub

Private Sub btnPlayers_Click()
    playersForm.Show 1
End Sub

Private Sub cmbTeams_LostFocus(Index As Integer)
Dim sqlstr As String
Dim cmd As New ADODB.Command
    If Me.cmbTeams(Index).Text = "" Then Exit Sub
    'find and update the record based on the tag of the control
    sqlstr = "Update tblTournamentTeamCodes Set teamId = " & Me.cmbTeams(Index).BoundText
    sqlstr = sqlstr & " WHERE tournamentID = " & thisTournament
    sqlstr = sqlstr & " AND teamcode = '" & Me.cmbTeams(Index).Tag & "'"
    On Error GoTo dataerror
    cn.BeginTrans
    With cmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = sqlstr
        .Execute
    End With
    cn.CommitTrans
    Exit Sub
dataerror:
    cn.RollbackTrans
    
End Sub

Private Sub Form_Load()
    makeForm
    centerForm Me
    UnifyForm Me
End Sub

