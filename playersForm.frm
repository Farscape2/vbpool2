VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton btnNew 
      Caption         =   "Nieuw"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   7920
      Width           =   1335
   End
   Begin MSComctlLib.ListView lstPlayers 
      Height          =   6855
      Left            =   240
      TabIndex        =   3
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
   Begin MSAdodcLib.Adodc dtcTeams 
      Height          =   330
      Left            =   120
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo cmbTeams 
      Bindings        =   "playersForm.frx":0000
      Height          =   360
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Spelers"
      Height          =   495
      Left            =   120
      TabIndex        =   2
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
      TabIndex        =   1
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

Private Sub btnNew_Click()
    'add player to database
    playerAddForm.Country = getTeamInfo(Me.cmbTeams.BoundText, "teamCountryId")
    playerAddForm.Show 1
    updateListview
End Sub

Private Sub btnOk_Click()
Unload Me
End Sub

Private Sub cmbTeams_Click(Area As Integer)
    If Area <> 0 Then
          updateListview
    End If
End Sub

Private Sub Form_Load()
'fill teams combo
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sqlstr As String
    sqlstr = "Select * from tblTeamNames Where teamtype <>0  and teamNameId IN "
    sqlstr = sqlstr & " (Select teamId from tblTournamentTeamCodes where tournamentid = " & thisTournament
    sqlstr = sqlstr & " ) Order by teamName"
    rsTeams.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    With Me.cmbTeams
        Set .RowSource = rsTeams
        .ListField = "teamName"
        .BoundColumn = "teamNameId"
        .Refresh
    End With
    Me.cmbTeams.Text = rsTeams!teamName
    UnifyForm Me
'    centerForm Me
    updateListview
    
    If (rs.State And adStateOpen) = adStateOpen Then rs.Close
    Set rs = Nothing

End Sub

Sub updateListview()
    Dim rsPlayers As ADODB.Recordset
    
    Dim lItem As ListItem
    Dim sqlstr As String
    
    Set rsPlayers = New ADODB.Recordset
    
    'get the tournament teamcode for this team
    thisTeam = getTournamentTeamCode(Me.cmbTeams.BoundText)
    
    sqlstr = "Select* from tblPeople "
    sqlstr = sqlstr & " Where countryCode = " & Nz(getTeamInfo(Me.cmbTeams.BoundText, "teamCountryId"), 0)
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
            lItem.Checked = playerInTournamentTeam(rsPlayers!peopleid, thisTeam)
            lItem.SubItems(1) = Nz(rsPlayers!peopleid, "")
            rsPlayers.MoveNext
        Loop
    End With
    If (rsPlayers.State And adStateOpen) = adStateOpen Then rsPlayers.Close
    Set rsPlayers = Nothing

End Sub

Private Sub lstPlayers_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    'add / remove player from tournament team
    Dim sqlstr As String
    If Item.Checked Then
        sqlstr = "Insert into tblTeamPlayers (tournamentId, teamId, playerId) "
        sqlstr = sqlstr & "VALUES (" & thisTournament
        sqlstr = sqlstr & ", " & thisTeam
        sqlstr = sqlstr & ", " & Val(Item.SubItems(1))
        sqlstr = sqlstr & ")"
    Else
        sqlstr = "Delete from tblTeamPlayers where tournamentId = " & thisTournament
        sqlstr = sqlstr & " AND teamID = " & thisTeam
        sqlstr = sqlstr & " AND playerId = " & Val(Item.SubItems(1))
    End If
    cn.Execute sqlstr
End Sub
