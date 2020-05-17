VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMatches 
   Caption         =   "Wedstrijden"
   ClientHeight    =   9630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12795
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
   ScaleHeight     =   9630
   ScaleMode       =   0  'User
   ScaleWidth      =   12795
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grdMatches 
      Height          =   7575
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   13361
      _Version        =   393216
      BackColorBkg    =   -2147483633
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin MSDataListLib.DataCombo cmbTeamA 
      Height          =   360
      Left            =   3120
      TabIndex        =   20
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.TextBox txtOrder 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   11040
      TabIndex        =   17
      Top             =   720
      Width           =   420
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   10680
      TabIndex        =   10
      Top             =   9120
      Width           =   1815
   End
   Begin VB.TextBox txtNr 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox txtTime 
      Height          =   375
      Left            =   2205
      TabIndex        =   5
      Text            =   "00:00"
      Top             =   720
      Width           =   630
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Opslaan"
      Height          =   495
      Left            =   11760
      TabIndex        =   7
      Top             =   600
      Width           =   855
   End
   Begin MSComCtl2.UpDown UpDnMinutes 
      Height          =   375
      Left            =   2820
      TabIndex        =   6
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Increment       =   15
      Max             =   45
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDnHours 
      Height          =   375
      Left            =   1950
      TabIndex        =   4
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Max             =   23
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.DTPicker dtDate 
      Height          =   375
      Left            =   855
      TabIndex        =   3
      Top             =   720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MM"
      Format          =   146538499
      CurrentDate     =   43939
   End
   Begin MSComCtl2.UpDown upDnNr 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtNr"
      BuddyDispid     =   196613
      OrigLeft        =   840
      OrigTop         =   480
      OrigRight       =   1095
      OrigBottom      =   855
      Max             =   144
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDnOrder 
      Height          =   375
      Left            =   11460
      TabIndex        =   18
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtOrder"
      BuddyDispid     =   196611
      OrigLeft        =   840
      OrigTop         =   480
      OrigRight       =   1095
      OrigBottom      =   855
      Max             =   144
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSDataListLib.DataCombo cmbTeamB 
      Height          =   360
      Left            =   5040
      TabIndex        =   21
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbTypes 
      Height          =   360
      Left            =   6960
      TabIndex        =   22
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbLocation 
      Height          =   360
      Left            =   9000
      TabIndex        =   23
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Volgorde"
      Height          =   255
      Index           =   7
      Left            =   10800
      TabIndex        =   19
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Locatie"
      Height          =   255
      Index           =   6
      Left            =   9240
      TabIndex        =   16
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Soort wedstrijd"
      Height          =   255
      Index           =   5
      Left            =   7200
      TabIndex        =   15
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Team B"
      Height          =   255
      Index           =   4
      Left            =   5160
      TabIndex        =   14
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Team A"
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   13
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tijd"
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   12
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Datum"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   11
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nr"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Wedstrijden"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Tag             =   "kop"
      Top             =   120
      Width           =   12255
   End
End
Attribute VB_Name = "frmMatches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dontMove As Boolean 'prevent editBar from updateting

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim rsTeams As ADODB.Recordset
Dim rsTypes As ADODB.Recordset
Dim rsLocation As ADODB.Recordset

Dim colWidths(9) As Double

Sub setMatchGrid()
Dim sqlstr As String
Dim dCol As Object
Dim col As Column
Dim i As Integer, j As Integer
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
  sqlstr = "SELECT m.matchNumber as nr, format(m.matchDate,'dd-MM') as Datum, format(matchTime ,'HH:NN') as Tijd, "
  sqlstr = sqlstr & " a.teamcode as A, ta.teamName as Team1, b.teamcode as B, tb.teamName as Team2, "
  sqlstr = sqlstr & " t.matchTypeDescription as Type, s.stadiumName & '/' & s.stadiumLocation as Locatie, m.matchOrder as vlg"
  sqlstr = sqlstr & " FROM ((((tblTournamentSchedule m LEFT JOIN tblStadiums s ON m.matchStadiumID = s.stadiumID) "
  sqlstr = sqlstr & " LEFT JOIN tblTournamentTeamCodes AS b ON m.matchTeamB = b.teamCode) "
  sqlstr = sqlstr & " LEFT JOIN tblTeamNames AS tb ON b.teamID = tb.teamNameID) "
  sqlstr = sqlstr & " LEFT JOIN (tblTournamentTeamCodes a "
  sqlstr = sqlstr & " LEFT JOIN tblTeamNames ta ON a.teamID = ta.teamNameID) ON m.matchTeamA = a.teamCode) "
  sqlstr = sqlstr & " LEFT JOIN tblMatchTypes t ON m.matchType = t.matchTypeID"
  sqlstr = sqlstr & " WHERE m.tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND a.tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND b.tournamentID = " & thisTournament
  sqlstr = sqlstr & " ORDER BY m.matchorder"
  rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
'fill the grid
   With Me.grdMatches
    .SelectionMode = flexSelectionByRow
    .Clear
     .rows = rs.RecordCount + 1
    .cols = rs.Fields.Count
    i = 0
    For j = 0 To rs.Fields.Count - 1
      If Not IsNull(rs.Fields(j).Name) Then
        .TextMatrix(i, j) = rs.Fields(j).Name
      End If
    Next
    rs.MoveFirst
    Do While Not rs.EOF
      i = i + 1
      For j = 0 To rs.Fields.Count - 1
        If Not IsNull(rs.Fields(j).value) Then
          .TextMatrix(i, j) = rs.Fields(j).value
        Else
          .TextMatrix(i, j) = ""
        End If
      Next
      rs.MoveNext
    Loop
    .ColWidth(0) = 400
    .ColAlignment(0) = flexAlignCenterCenter
    .ColWidth(1) = 700
    .ColAlignment(1) = flexAlignCenterCenter
    .ColWidth(2) = 600
    .ColAlignment(2) = flexAlignCenterCenter
    .ColWidth(3) = 600
    .ColAlignment(3) = flexAlignCenterCenter
    .ColWidth(4) = 1750
    .ColWidth(5) = 600
    .ColAlignment(5) = flexAlignCenterCenter
    .ColWidth(6) = 1750
    .ColWidth(7) = 1900
    .ColWidth(8) = 3250
    .ColWidth(9) = 500
    .ColAlignment(9) = flexAlignCenterCenter
    
   End With
   'select entire row
  ' Me.grdMatches.MarqueeStyle = dbgHighlightRow
   'force update of editBar controls
   'grdMatches_RowColChange 1, 1
    
End Sub

Sub setEditBar()

    Set rsTeams = New ADODB.Recordset
    Set rsTypes = New ADODB.Recordset
    Set rsLocation = New ADODB.Recordset

    Dim sqlstr As String

    ' Using DataCombo boxes for a change. Is so much easties in this case
    ' Normal ComboBox.ItemData can only be long data type
    'besides it is doing strange thing when filling and getting the actual value

    sqlstr = "Select teamcode, teamCode & ': ' & teamName as team "
    sqlstr = sqlstr & "from tblTournamentTeamCodes c LEFT JOIN tblTeamNames n on c.teamId = n.teamnameid"
    sqlstr = sqlstr & " Where c.tournamentid = " & thisTournament
    rsTeams.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    With Me.cmbTeamA
        Set .RowSource = rsTeams
        .ListField = "team"
        .BoundColumn = "teamcode"
    End With
    With Me.cmbTeamB
        Set .RowSource = rsTeams
        .ListField = "team"
        .BoundColumn = "teamcode"
    End With
    
    sqlstr = "Select matchtypeId as id , matchtypedescription as descr from tblMatchTypes"
    rsTypes.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    With Me.cmbTypes
        Set .RowSource = rsTypes
        .ListField = "descr"
        .BoundColumn = "id"
    End With
    
    sqlstr = "Select stadiumId as id, stadiumName & '/' & stadiumLocation as name from tblStadiums order by stadiumName"
    rsLocation.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    With Me.cmbLocation
        Set .RowSource = rsLocation
        .ListField = "name"
        .BoundColumn = "id"
    End With
    
    Me.dtDate = getTournamentInfo("tournamentStartDate", cn)
    Me.UpDnHours = 20
    
    
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
    Dim sqlstr As String
    Set rs = New ADODB.Recordset
    'validate
    Dim msg As String
    msg = ""
    If Not IsNumeric(Me.txtNr.Text) Then msg = msg & "Geen wedstrijdnummer" & vbNewLine
    If Me.cmbTeamA.BoundText = "" Then msg = msg & "Geen Team A" & vbNewLine
    If Me.cmbTeamB.BoundText = "" Then msg = msg & "Geen Team B" & vbNewLine
    If Me.cmbTypes.BoundText < 1 Then msg = msg & "Geen soort wedstrijd" & vbNewLine
    If Me.cmbLocation.BoundText < 1 Then msg = msg & "Geen locatie" & vbNewLine
    If Not IsNumeric(Me.txtOrder.Text) Then Me.UpDnOrder = Me.upDnNr
    If msg > "" Then
        msg = "FOUT: " & vbNewLine
        MsgBox msg, vbOKOnly + vbCritical, "Wedstrijd toevoegen"
        Exit Sub
    End If
    
    sqlstr = "Select * from tblTournamentSchedule Where tournamentId = " & thisTournament
    rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
    rs.Find "matchNumber = " & val(Me.txtNr)
    
    If rs.EOF Then 'add new match
        rs.AddNew
    End If
    With rs
        !TournamentId = thisTournament
        !matchnumber = val(Me.txtNr)
        !matchdate = CDbl(Me.dtDate)
        !matchtime = IIf(Me.txtTime = "24:00", "23:59", Me.txtTime)
        !matchteamA = Me.cmbTeamA.BoundText
        !matchteamb = Me.cmbTeamB.BoundText
        !matchtype = Me.cmbTypes.BoundText
        !matchStadiumID = Me.cmbLocation.BoundText
        !matchorder = val(Me.txtOrder)
    End With
    rs.Update
    
    setMatchGrid
    findInGrid Me.upDnNr
    DoEvents
    
End Sub

Sub findInGrid(txt As String)
'find the txt in the grdMatches
  Dim i As Integer
  Do While Not Me.grdMatches.TextMatrix(i, 0) = txt
    i = i + 1
  Loop
  If i <= Me.grdMatches.rows Then
    Me.grdMatches.row = i
  End If
End Sub

Private Sub Form_Load()
    'open the database
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = lclConn()
        .Open
    End With
    
    setEditBar
    setMatchGrid
    Me.grdMatches.row = 2
    Me.grdMatches.row = 1
    setState 'only if admin is logged in is editting possible
    
    UnifyForm Me
    centerForm Me

End Sub

Sub setState()
Dim ctl As Control
Dim col As Object
    For Each ctl In Me.Controls
        If TypeOf ctl Is UpDown _
            Or TypeOf ctl Is ComboBox _
            Or TypeOf ctl Is DataCombo _
            Or TypeOf ctl Is DTPicker _
            Or TypeOf ctl Is TextBox _
            Or ctl.Name = "btnSave" Then
            ctl.Enabled = adminLogin
        End If
'        Me.grdMatches.AllowAddNew = adminLogin
'        Me.grdMatches.AllowDelete = adminLogin
'        Me.grdMatches.AllowUpdate = adminLogin
'        For Each col In Me.grdMatches.Columns
'            col.Locked = Not adminLogin
'        Next
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not rs Is Nothing Then
        If (rs.State And adStateOpen) = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    If Not rsTeams Is Nothing Then
        If (rsTeams.State And adStateOpen) = adStateOpen Then rsTeams.Close
        Set rsTeams = Nothing
    End If
    
    If Not rsLocation Is Nothing Then
        If (rsLocation.State And adStateOpen) = adStateOpen Then rsLocation.Close
        Set rsLocation = Nothing
    End If
    If Not rsTypes Is Nothing Then
        If (rsTypes.State And adStateOpen) = adStateOpen Then rsTypes.Close
        Set rsTypes = Nothing
    End If
    If Not cn Is Nothing Then
        If (cn.State And adStateOpen) = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    
End Sub

Private Sub Form_Resize()
'attempt to resize everything with the form
'if to complicated we will fix the forms borders
Dim initHeigth As Integer, initWidth As Integer
Dim gridHeight As Integer, gridWidth As Integer
Dim col As Column
Dim i As Integer
Const leftPos = 240
'set the initial Height/Width
initHeigth = 10215
initWidth = 12945
gridHeight = 7815
gridWidth = 12360
With Me
  If .Height < initHeigth Then .Height = initHeigth
  If .width < initWidth Then .width = initWidth
  .grdMatches.Height = .Height - 2400
  .grdMatches.width = .width - 585
  .btnClose.Top = .Height - 1095
  .btnClose.Left = .width - 2265
  .btnSave.Left = .width - 1185
  i = 0
'  For Each col In .grdMatches.Columns
'    col.width = .grdMatches.width * colWidths(i)
'    i = i + 1
'  Next
  
End With

End Sub


Private Sub grdMatches_RowColChange()
  Dim sqlstr As String
  Set rs = New ADODB.Recordset
  sqlstr = "Select * from tblTournamentSchedule  where tournamentID = " & thisTournament
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  rs.Find "matchNumber = " & Me.grdMatches.TextMatrix(Me.grdMatches.row, 0)
  With rs
    If Not .EOF Then
      Me.txtNr = !matchnumber
      Me.dtDate = !matchdate
      Me.txtTime = !matchtime
      Me.cmbTypes.BoundText = !matchtype
      Me.cmbLocation.BoundText = !matchStadiumID
      Me.cmbTeamA.BoundText = !matchteamA
      Me.cmbTeamB.BoundText = !matchteamb
      Me.txtOrder = !matchorder
    End If
  End With
End Sub


Private Sub updnMinutes_Change()
  Me.txtTime = Format(Me.UpDnHours, "00") & ":" & Format(Me.UpDnMinutes, "00")
End Sub

Private Sub updnHours_Change()
  Me.txtTime = Format(Me.UpDnHours, "00") & ":" & Format(Me.UpDnMinutes, "00")
End Sub


Private Sub upDnNr_DownClick()
  findInGrid Me.txtNr
End Sub

Private Sub upDnNr_UpClick()
  findInGrid Me.txtNr
End Sub
