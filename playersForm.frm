VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form playersForm 
   Caption         =   "Spelers"
   ClientHeight    =   8505
   ClientLeft      =   12540
   ClientTop       =   3435
   ClientWidth     =   3540
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
   ScaleHeight     =   8505
   ScaleWidth      =   3540
   Begin VB.ComboBox cmbTeams 
      Height          =   360
      Left            =   960
      TabIndex        =   5
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton btnNew 
      Caption         =   "Nieuw"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   7920
      Width           =   1335
   End
   Begin MSComctlLib.ListView lstPlayers 
      Height          =   6855
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   12091
      View            =   2
      Arrange         =   2
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Spelers"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Tag             =   "kop"
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Team"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "playersForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'to preserve the tournamentTeamCode
Dim thisTeam As Long

Dim cn As ADODB.Connection

Dim rs As ADODB.Recordset

Private Sub btnNew_Click()
    'add player to database
    playerAddForm.Country = getTeamInfo(Me.cmbTeams.ItemData(Me.cmbTeams.ListIndex), "teamCountryId", cn)
    playerAddForm.Show 1
    updateListview
End Sub

Private Sub btnOk_Click()
Unload Me
End Sub


Private Sub cmbTeams_Click()
    updateListview
End Sub

Private Sub Form_Load()
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = lclConn()
        .Open
    End With
    Set rs = New ADODB.Recordset
    Dim sqlstr As String
    sqlstr = "Select * from tblTeamNames Where teamtype <>0  and teamNameId IN "
    sqlstr = sqlstr & " (Select teamid from tblTournamentTeamCodes where tournamentid = " & thisTournament
    sqlstr = sqlstr & " ) Order by teamName"
    'fill teams combo
        
    FillCombo Me.cmbTeams, sqlstr, cn, "teamName", "teamNameid"
'    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
'    With Me.cmbTeams
'        Set .RowSource = rs
'        .ListField = "teamName"
'        .BoundColumn = "teamNameId"
'        .Refresh
'    End With
    Me.cmbTeams.ListIndex = 0
    UnifyForm Me
'    centerForm Me
    updateListview
    
End Sub

Sub updateListview()
    Dim rsPlayers As ADODB.Recordset
    
    Dim lItem As ListItem
    Dim sqlstr As String
    
    Set rsPlayers = New ADODB.Recordset
    
    'get the tournament teamcode for this team
    thisTeam = Me.cmbTeams.ItemData(Me.cmbTeams.ListIndex)
    
    sqlstr = "Select* from tblPeople "
    sqlstr = sqlstr & " Where countryCode = " & nz(getTeamInfo(thisTeam, "teamCountryId", cn), 0)
    sqlstr = sqlstr & " and function1 >1 and function1 <6"
    sqlstr = sqlstr & " Order by nickname"
    rsPlayers.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    
    With Me.lstPlayers
        .ListItems.Clear
        .ColumnHeaders.Add , , "Bijnaam", 2500
        .ColumnHeaders.Add , , "ID", 0
        .View = lvwReport
        .Checkboxes = True
        .Sorted = True
        .SortKey = 0
        Do While Not rsPlayers.EOF
            Set lItem = .ListItems.Add(1)
            lItem.Text = rsPlayers!NickName
            lItem.Checked = playerInTournamentTeam(rsPlayers!peopleid, thisTeam, cn)
            lItem.SubItems(1) = nz(rsPlayers!peopleid, "")
            rsPlayers.MoveNext
        Loop
    End With
    If (rsPlayers.State And adStateOpen) = adStateOpen Then rsPlayers.Close
    Set rsPlayers = Nothing

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Clean-up procedure
    If Not rs Is Nothing Then
        'first, check if the state is open, if yes then close it
        If (rs.State And adStateOpen) = adStateOpen Then
            rs.Close
        End If
        'set them to nothing
        Set rs = Nothing
    End If
    'same comment with rs
    If Not cn Is Nothing Then
        If (cn.State And adStateOpen) = adStateOpen Then
            cn.Close
        End If
        Set cn = Nothing
    End If
End Sub

Private Sub lstPlayers_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    'add / remove player from tournament team
    Dim sqlstr As String
    If Item.Checked Then
        sqlstr = "Insert into tblTeamPlayers (tournamentId, teamId, playerId) "
        sqlstr = sqlstr & "VALUES (" & thisTournament
        sqlstr = sqlstr & ", " & thisTeam
        sqlstr = sqlstr & ", " & val(Item.SubItems(1))
        sqlstr = sqlstr & ")"
    Else
        sqlstr = "Delete from tblTeamPlayers where tournamentId = " & thisTournament
        sqlstr = sqlstr & " AND teamID = " & thisTeam
        sqlstr = sqlstr & " AND playerId = " & val(Item.SubItems(1))
    End If
    cn.Execute sqlstr
End Sub
