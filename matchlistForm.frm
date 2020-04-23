VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form matchlistForm 
   Caption         =   "Wedstrijden"
   ClientHeight    =   9630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12705
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
   ScaleWidth      =   12705
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOrder 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   11040
      TabIndex        =   21
      Top             =   720
      Width           =   420
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   10680
      TabIndex        =   14
      Top             =   9120
      Width           =   1815
   End
   Begin VB.TextBox txtNr 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   555
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
      TabIndex        =   11
      Top             =   600
      Width           =   855
   End
   Begin MSDataListLib.DataCombo cmbTeamA 
      Height          =   360
      Left            =   3120
      TabIndex        =   7
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
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
   Begin MSAdodcLib.Adodc dtcMatches 
      Height          =   330
      Left            =   840
      Top             =   9120
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
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
      Caption         =   ""
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
   Begin MSDataGridLib.DataGrid grdMatches 
      Height          =   7815
      Left            =   240
      TabIndex        =   13
      Top             =   1200
      Width           =   12360
      _ExtentX        =   21802
      _ExtentY        =   13785
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   "Match"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            Alignment       =   2
            ColumnAllowSizing=   0   'False
            Object.Visible         =   -1  'True
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo cmbTeamB 
      Height          =   360
      Left            =   5160
      TabIndex        =   8
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbTypes 
      Height          =   360
      Left            =   7200
      TabIndex        =   9
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbLocation 
      Height          =   360
      Left            =   9240
      TabIndex        =   10
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
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
      Format          =   148176899
      CurrentDate     =   43939
   End
   Begin MSComCtl2.UpDown upDnNr 
      Height          =   375
      Left            =   540
      TabIndex        =   2
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtNr"
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
   Begin MSComCtl2.UpDown UpDnOrder 
      Height          =   375
      Left            =   11460
      TabIndex        =   22
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtOrder"
      BuddyDispid     =   196609
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Volgorde"
      Height          =   255
      Index           =   7
      Left            =   10800
      TabIndex        =   23
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Locatie"
      Height          =   255
      Index           =   6
      Left            =   9240
      TabIndex        =   20
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Soort wedstrijd"
      Height          =   255
      Index           =   5
      Left            =   7200
      TabIndex        =   19
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Team B"
      Height          =   255
      Index           =   4
      Left            =   5160
      TabIndex        =   18
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Team A"
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   17
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
      TabIndex        =   16
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Datum"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   15
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
      TabIndex        =   12
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
      Width           =   11895
   End
End
Attribute VB_Name = "matchlistForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dontMove As Boolean 'prevent editBar from updateting

Dim rs As ADODB.Recordset
Dim rsTypes As ADODB.Recordset
Dim rsLocation As ADODB.Recordset

Sub setMatchGrid()
Dim sqlstr As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim dCol As Column
    
    sqlstr = "SELECT m.matchNumber as Wedstr, m.matchDate as Datum, matchTime as Tijd, "
    sqlstr = sqlstr & " ta.teamName as Team1, tb.teamName as Team2, t.matchTypeDescription as Type, s.stadiumName & '/' & s.stadiumLocation as Locatie,"
    sqlstr = sqlstr & " a.teamcode as CodeA, b.teamcode as CodeB, t.matchtypeId as typeId, s.stadiumId as stadiumId, m.matchOrder as volgorde"
    sqlstr = sqlstr & " FROM ((((tblTournamentSchedule m LEFT JOIN tblStadiums s ON m.matchStadiumID = s.stadiumID) "
    sqlstr = sqlstr & " LEFT JOIN tblTournamentTeamCodes AS b ON m.matchTeamB = b.teamCode) "
    sqlstr = sqlstr & " LEFT JOIN tblTeamNames AS tb ON b.teamID = tb.teamNameID) "
    sqlstr = sqlstr & " LEFT JOIN (tblTournamentTeamCodes a "
    sqlstr = sqlstr & " LEFT JOIN tblTeamNames ta ON a.teamID = ta.teamNameID) ON m.matchTeamA = a.teamCode) "
    sqlstr = sqlstr & " LEFT JOIN tblMatchTypes t ON m.matchType = t.matchTypeID"
    sqlstr = sqlstr & " WHERE m.tournamentID = " & thisTournament
    sqlstr = sqlstr & " AND a.tournamentID = " & thisTournament
    sqlstr = sqlstr & " AND b.tournamentID = " & thisTournament
    sqlstr = sqlstr & " ORDER BY m.matchNumber"
    
    rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
    
    With Me.dtcMatches
        .ConnectionString = cn
        .RecordSource = sqlstr
        .CommandType = adCmdText
        .Refresh
    End With
    
    With Me.grdMatches
        Set .DataSource = Me.dtcMatches
        Set dCol = .Columns(0)
        dCol.Caption = " Nr"
        dCol.DataField = "Wedstr"
        dCol.Alignment = dbgCenter
        dCol.Width = 400
        dCol.Visible = True
        Set dCol = .Columns.Add(1)
        dCol.Caption = "Datum"
        dCol.DataField = "Datum"
        dCol.NumberFormat = "dd-MM"
        dCol.Alignment = dbgCenter
        dCol.Width = 700
        dCol.Visible = True
        Set dCol = .Columns.Add(2)
        dCol.Caption = "Tijd"
        dCol.DataField = "Tijd"
        dCol.NumberFormat = "hh:mm"
        dCol.Width = 600
        dCol.Visible = True
        Set dCol = .Columns.Add(3)
        dCol.DataField = "CodeA"
        dCol.Alignment = dbgCenter
        dCol.Caption = "  A"
        dCol.Width = 600
        dCol.Visible = True
        Set dCol = .Columns.Add(4)
        dCol.Caption = "TeamA"
        dCol.DataField = "Team1"
        dCol.Width = 2000
        dCol.Visible = True
        Set dCol = .Columns.Add(5)
        dCol.Caption = "  B"
        dCol.DataField = "CodeB"
        dCol.Alignment = dbgCenter
        dCol.Width = 600
        dCol.Visible = True
        Set dCol = .Columns.Add(6)
        dCol.Caption = "TeamB"
        dCol.DataField = "Team2"
        dCol.Width = 1900
        dCol.Visible = True
        Set dCol = .Columns.Add(7)
        dCol.Caption = "Type"
        dCol.DataField = "Type"
        dCol.Width = 1900
        dCol.Alignment = dbgCenter
        dCol.Visible = True
        Set dCol = .Columns.Add(8)
        dCol.Caption = "Locatie"
        dCol.DataField = "Locatie"
        dCol.Width = 2650
        dCol.Visible = True
        Set dCol = .Columns.Add(9)
        dCol.Caption = "Volg"
        dCol.DataField = "volgorde"
        dCol.Width = 400
        dCol.Alignment = dbgCenter
        dCol.Visible = True
        
        .Columns(3).Alignment = dbgCenter
        .Columns(5).Alignment = dbgCenter
        .ReBind
        .Refresh
    End With
    'select entire row
    Me.grdMatches.MarqueeStyle = dbgHighlightRow
    'force update of editBar controls
    grdMatches_RowColChange 1, 1
    
End Sub

Sub setEditBar()
Set rs = New ADODB.Recordset
Set rsTypes = New ADODB.Recordset
Set rsLocation = New ADODB.Recordset

Dim sqlstr As String

    sqlstr = "Select teamcode, teamCode & ': ' & teamName as team "
    sqlstr = sqlstr & "from tblTournamentTeamCodes c LEFT JOIN tblTeamNames n on c.teamId = n.teamnameid"
    sqlstr = sqlstr & " Where c.tournamentid = " & thisTournament
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then rs.MoveLast

    With Me.cmbTeamA
        Set .RowSource = rs
        .ListField = "team"
        .BoundColumn = "teamcode"
        .Refresh
    End With
    With Me.cmbTeamB
        Set .RowSource = rs
        .ListField = "team"
        .BoundColumn = "teamcode"
        .Refresh
    End With
    sqlstr = "Select matchtypeId as id , matchtypedescription as descr from tblMatchTypes"
    rsTypes.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    With Me.cmbTypes
        Set .RowSource = rsTypes
        .ListField = "descr"
        .BoundColumn = "id"
        .Refresh
    End With
    sqlstr = "Select stadiumId as id, stadiumName & '/' & stadiumLocation as name from tblStadiums order by stadiumName"
    rsLocation.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    With Me.cmbLocation
        Set .RowSource = rsLocation
        .ListField = "name"
        .BoundColumn = "id"
        .Refresh
    End With
    
    
    Me.dtDate = getTournamentInfo("tournamentStartDate")
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
    rs.Find "matchNumber = " & Val(Me.txtNr)
    
    If rs.EOF Then 'add new match
        rs.AddNew
    End If
    With rs
        !tournamentID = thisTournament
        !matchnumber = Me.upDnNr
        !matchdate = CDbl(Me.dtDate)
        !matchtime = IIf(Me.txtTime = "24:00", "23:59", Me.txtTime)
        !matchTeamA = Me.cmbTeamA.BoundText
        !matchTeamB = Me.cmbTeamB.BoundText
        !matchType = Me.cmbTypes.BoundText
        !matchStadiumId = Me.cmbLocation.BoundText
        !matchOrder = Me.UpDnOrder
    End With
    rs.Update
    
    Me.dtcMatches.Recordset.Requery
    Me.dtcMatches.Refresh
    Set Me.grdMatches.DataSource = Me.dtcMatches
    Me.grdMatches.Refresh
    DoEvents
    Me.dtcMatches.Recordset.Move Val(Me.txtNr) - 1, 0
    
End Sub

Private Sub Form_Load()

    setEditBar
    setMatchGrid
    
    UnifyForm Me
    centerForm Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If (rs.State And adStateOpen) = adStateOpen Then rs.Close
    Set rs = Nothing
    If (rsLocation.State And adStateOpen) = adStateOpen Then rsLocation.Close
    Set rsLocation = Nothing
    If (rsTypes.State And adStateOpen) = adStateOpen Then rsTypes.Close
    Set rsTypes = Nothing

End Sub

Private Sub grdMatches_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
     'update editBar
     On Error Resume Next
     If dontMove Then Exit Sub
     With Me.dtcMatches.Recordset
        Me.upDnNr = !wedstr
        Me.dtDate = !datum
        Me.txtTime = Format(!tijd, "hh:mm")
        Me.cmbTeamA.BoundText = !CodeA
        Me.cmbTeamB.BoundText = !CodeB
        Me.cmbTypes.BoundText = !typeid
        Me.cmbLocation.BoundText = !stadiumId
        Me.UpDnOrder = !volgorde
    End With
End Sub

Private Sub updnMinutes_Change()
  Me.txtTime = Format(Me.UpDnHours, "00") & ":" & Format(Me.UpDnMinutes, "00")
End Sub

Private Sub updnHours_Change()
  Me.txtTime = Format(Me.UpDnHours, "00") & ":" & Format(Me.UpDnMinutes, "00")
End Sub


Private Sub upDnNr_DownClick()
    With Me.dtcMatches.Recordset
        dontMove = True
        .MoveFirst
        dontMove = False
        .Find "Wedstr = " & Me.upDnNr
        
        'when at first match, the rowcolchange event doesn't get fired. so trick it
        If Me.upDnNr = 1 Then
            dontMove = True
            .MoveLast
            dontMove = False
            .MoveFirst
        End If
    End With
    
    
End Sub

Private Sub upDnNr_UpClick()
    With Me.dtcMatches.Recordset
        dontMove = True
        .MoveFirst
        dontMove = False
        .Find "Wedstr = " & Me.upDnNr
        'add new
        If .EOF Then
            Me.cmbTeamA = ""
            Me.cmbTeamB = ""
            Me.cmbTypes.BoundText = "1"
            Me.cmbLocation = ""
        End If
    End With
End Sub
