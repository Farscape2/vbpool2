Attribute VB_Name = "dbfunctions"
Option Explicit

Function getOrganisation(cn As ADODB.Connection, Optional field As String) As String
'get the name for the organisation of this pool / or just the content of field
Dim adoCmd As ADODB.Command
Dim rs As ADODB.Recordset
Dim sqlstr As String
Dim result As String
    sqlstr = "Select * from tblOrganisation"
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = sqlstr
        Set rs = .Execute
    End With

    If Not rs.EOF Then
        If field = "" Then
            result = Trim(rs!firstname)
            If rs!middlename > "" Then
                result = result & " " & Trim(rs!middlename)
            End If
            If rs!lastname > "" Then
                result = result & " " & Trim(rs!lastname)
            End If
'            result = result & vbNewLine & Trim(rs!address) & vbNewLine & Trim(rs!postalcode) & " " & Trim(rs!city)
        Else
            result = rs(field)
        End If
    End If
    getOrganisation = result
    rs.Close
    Set rs = Nothing
End Function

Function getPoolInfo(fldName As String, cn As ADODB.Connection)
'return the value of fieldnmame in tblPools
Dim adoCmd As ADODB.Command
Dim rs As ADODB.Recordset
Dim sqlstr As String

    Set adoCmd = New ADODB.Command
    sqlstr = "Select " & fldName & " from tblPools where poolid = ?"
    With adoCmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = sqlstr
        .Prepared = True
        .Parameters.Append .CreateParameter("id", adInteger, adParamInput)
        .Parameters("id").value = thisPool
        Set rs = .Execute
    End With
    If Not rs.EOF Then
        getPoolInfo = rs(fldName)
    Else
        getPoolInfo = Null
    End If
    
    rs.Close
    Set rs = Nothing
    Set adoCmd = Nothing

End Function

Function getTournamentInfo(fldName As String, cn As ADODB.Connection)
'return the value of fieldnmame in tblTournaments
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command
    Dim sqlstr As String
    Dim result As Variant
    Dim rs As ADODB.Recordset
    
    sqlstr = "Select * from tblTournaments Where tournamentID = ? "
    With adoCmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = sqlstr
        .Prepared = True
        .Parameters.Append .CreateParameter("id", adInteger, adParamInput)
        .Parameters("id").value = thisTournament
        Set rs = .Execute
    End With
    If Not rs.EOF Then
        ' add description as extra - Access doesn't understand concat
        If fldName = "description" Then
            result = rs!tournamenttype & " - " & rs!tournamentYear
        Else
            If rs(fldName).Type = adBoolean Then
                result = CBool(rs(fldName)) * 1
            Else
                result = rs(fldName)
            End If
        End If
    Else
        result = Null
    End If
    getTournamentInfo = result
    rs.Close
    Set rs = Nothing
    Set adoCmd = Nothing
End Function

Function chkPoolHasCompetitors(pool As Long, cn As ADODB.Connection)
'are there competitors for this pool
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sqlstr As String
        
        sqlstr = "Select  poolID from tblCompetitors Where poolid = " & pool
        rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
        chkPoolHasCompetitors = Not rs.EOF
    
    rs.Close
    Set rs = Nothing
End Function

Function chkTournamentHasPools(tournament As Long, cn As ADODB.Connection)
'are there pools for this tournament?
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sqlstr As String
        sqlstr = "Select tournamentID from tblPools Where tournamentid = " & tournament
        rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
        chkTournamentHasPools = Not rs.EOF
    rs.Close
    Set rs = Nothing
End Function

Function getThisPoolTournamentId(cn As ADODB.Connection) As Long
'return the tournament for the current pool
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    getThisPoolTournamentId = 0
    Dim sqlstr As String
    sqlstr = "Select tournamentID from tblPools Where poolid = " & thisPool
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
        getThisPoolTournamentId = rs!TournamentId
    End If
    rs.Close
    Set rs = Nothing
End Function

Function chkTournamentStarted(cn As ADODB.Connection)
'check to see if tournament already started

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sqlstr As String
    chkTournamentStarted = False
    sqlstr = "Select * from tblTournaments Where tournamentid = " & getThisPoolTournamentId(cn)
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
        chkTournamentStarted = CDbl(rs!tournamentStartDate) < CDate(Now())
    End If
    rs.Close
    Set rs = Nothing
End Function

Function supportsTransactions(cn As ADODB.Connection) As Boolean
'check if connection supports transactions
    On Error GoTo err_supportsTransactions:
        Dim lValue As Long
        lValue = cn.Properties("Transaction DDL").value
        supportsTransactions = True
    Exit Function
err_supportsTransactions:
    Select Case Err.number
    Case adErrItemNotFound:
        supportsTransactions = False
    Case Else
        MsgBox Err.Description
    End Select
End Function

Function tournamentHasSchedule(cn As ADODB.Connection) As Boolean
'check if there is already a schedule made
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sqlstr As String
    sqlstr = "select * from tblTournamentSchedule where tournamentid = " & thisTournament
    rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
    tournamentHasSchedule = Not rs.EOF
    rs.Close
    Set rs = Nothing
End Function

Function tournamentBaseSchedule() As Boolean
'check if there is already a base schedule made
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = lclConn
        .Open
    End With
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sqlstr As String
    sqlstr = "select * from tblTournamentTeamCodes where tournamentid = " & thisTournament
    rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
    tournamentBaseSchedule = Not rs.EOF
    rs.Close
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
End Function

Sub generateSchedule()
'this routine builds the teams codes table for later use in Schedule. There we will add teamnames to these codes

Dim rsSchedule As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim sqlstr As String
Dim msg As String
Dim qry As ADODB.Command
Dim makeSchedule As Boolean
Dim letter As Integer
Dim matches As Integer
Dim groupSize  As Integer
Dim i As Integer, j As Integer
Dim teamCode As String

    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = lclConn
        .Open
    End With
    
    Set rsSchedule = New ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set qry = New ADODB.Command
    
    'we will exit
    'if there are is already a base schedule (in tblTournamentTeamCodses) for this tournament
     If tournamentBaseSchedule Then Exit Sub
    ''!!!!!!!!!!!!!!!!!!!
    'this routine gereates all the teamcodes necessary for this tournament. It will OVERWRITE the existing tblTournamentTeamCodes
    '!!!!!!!!!!!!!!!!!!!!
    sqlstr = "Select tournamentTeamCount as teams, tournamentGroupCount as groups from tblTournaments where tournamentId = " & thisTournament
    rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
    If rs.EOF Then Exit Sub
    groupSize = rs!teams / rs!groups
    matches = (groupSize - 1) * 2 * rs!groups 'total matches during groupfase
    'empty the codes table for this tournament
    cn.Execute "Delete from tblTournamentTeamCodes where tournamentid = " & thisTournament
    cn.Execute "Delete from tblTournamentSchedule where tournamentID = " & thisTournament
    sqlstr = "Select * from tblTournamentTeamCodes"
    rsSchedule.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
    With rsSchedule
        For i = 1 To rs!groups
            For j = 1 To groupSize
                .AddNew
                !TournamentId = thisTournament
                teamCode = Chr(i + 64) & Format(j, "0")
                !teamCode = teamCode
                .Update
            Next
        Next
        If rs!groups > 4 Then
        '8th finales (normally I hope), should be 16 teams
            For i = 1 To rs!groups
                .AddNew
                !TournamentId = thisTournament
                !teamCode = "1" & Chr(i + 64)
                .Update
                .AddNew
                !TournamentId = thisTournament
                !teamCode = "2" & Chr(i + 64)
                .Update
            Next
            'if there are 6 groups then we need to add the best 3rd places to gt to 16
            If rs!groups = 6 Then  'add best 3rd places
                .AddNew
                !TournamentId = thisTournament
                !teamCode = "3ABC"
                .Update
                .AddNew
                !TournamentId = thisTournament
                !teamCode = "3ABCD"
                .Update
                .AddNew
                !TournamentId = thisTournament
                !teamCode = "3DEF"
                .Update
                .AddNew
                !TournamentId = thisTournament
                !teamCode = "3ADEF"
                .Update
            End If
        End If
        'other finals just the W(inner) of the matchnumber
        For i = matches + 1 To matches + 15
            .AddNew
            !TournamentId = thisTournament
            !teamCode = "W" & Format(i, "00")
            .Update
        Next
        If getTournamentInfo("tournamentThirdPlace", cn) Then 'add match for third place
            .AddNew
            !TournamentId = thisTournament
            !teamCode = "V" & Format(matches + 14, "00")
            .Update
        End If
    End With
    If (rs.State And adStateOpen) = adStateOpen Then rs.Close
    If (rsSchedule.State And adStateOpen) = adStateOpen Then rsSchedule.Close
    Set rs = Nothing
    Set rsSchedule = Nothing
    cn.Close
    Set cn = Nothing
End Sub

Sub addPlayers(cn As ADODB.Connection)
'add all players in the tblPeople table from a country in this tournament
    Dim sqlstr As String
    Dim rsTeams As ADODB.Recordset
    Dim rsPlayers As ADODB.Recordset
    
    Set rsTeams = New ADODB.Recordset
    Set rsPlayers = New ADODB.Recordset
    
    'remove all players in thistournament first
    sqlstr = "Delete from tblTeamPlayers where tournamentid = " & thisTournament
    cn.Execute sqlstr
    ' now build aqlstr to add players to teams
    sqlstr = "SELECT tournamentID, teamID, a.teamcodeID, teamName, b.teamCountryID, teamType "
    sqlstr = sqlstr & " FROM tblTeamNames b INNER JOIN tblTournamentTeamCodes a ON b.teamNameID = a.teamID"
    sqlstr = sqlstr & " where tournamentID = " & thisTournament
    rsTeams.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    
    If rsTeams.EOF Then Exit Sub 'if there are no eams what are we doing here
    
    rsTeams.MoveFirst
    Do While Not rsTeams.EOF
        'get all football players (function betweeen 2 and 5) from the same country as the team (NOT for clubteams)
        sqlstr = "Select * from tblPeople where function > 1 and function < 6 and countryCode = " & rsTeams!teamCountryId
        rsPlayers.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
        Do While Not rsPlayers.EOF
            sqlstr = "Insert into tblTeamPlayers (tournamentId, teamId, PlayerId) VALUES (" & thisTournament & "," & rsTeams!teamcodeID & ", " & rsPlayers!peopleid & ")"
            cn.Execute sqlstr
            rsPlayers.MoveNext
        Loop
        rsPlayers.Close
        rsTeams.MoveNext
    Loop
    If (rsTeams.State And adStateOpen) = adStateOpen Then rsTeams.Close
    Set rsTeams = Nothing
    If (rsPlayers.State And adStateOpen) = adStateOpen Then rsPlayers.Close
    Set rsPlayers = Nothing
End Sub
 
Function getTeamInfo(teamId As Long, fld As String, cn As ADODB.Connection)
    Dim sqlstr As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sqlstr = "Select * from tblTeamNames where teamNameId = " & teamId
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
        getTeamInfo = rs(fld)
    Else
        getTeamInfo = Null
    End If
    rs.Close
    Set rs = Nothing
End Function

Function getTeamId(tournamentTeamCode As Long, cn As ADODB.Connection)
'get the basic id  of a tournament teamcode
    Dim sqlstr As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sqlstr = "Select * from tblTournamentTeamCodes where teamCodeId = " & tournamentTeamCode
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
        getTeamId = rs(rs!teamId)
    Else
        getTeamId = Null
    End If
    rs.Close
    Set rs = Nothing
End Function

Function getTournamentTeamCode(teamId As Long, cn As ADODB.Connection)
'get the teamId from a tounamentTeamCode
    Dim sqlstr As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sqlstr = "Select * from tblTournamentTeamCodes where tournamentId = " & thisTournament
    sqlstr = sqlstr & " AND teamId = " & teamId
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
        getTournamentTeamCode = rs!teamCode
    Else
        getTournamentTeamCode = Null
    End If
    rs.Close
    Set rs = Nothing

End Function


Function playerInTournamentTeam(playerId As Long, teamId As Long, cn As ADODB.Connection)
    Dim sqlstr As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sqlstr = "Select * from tblTeamPlayers where teamId = " & teamId
    sqlstr = sqlstr & " AND playerId = " & playerId
    sqlstr = sqlstr & " AND tournamentId = " & thisTournament
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    
    playerInTournamentTeam = Not rs.EOF
    
    rs.Close
    Set rs = Nothing
End Function

Function playerExists(fName As String, mName As String, lName As String, NickName As String, cn As ADODB.Connection)
    'check double entries
    Dim sqlstr As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sqlstr = "Select * from tblPeople where (firstname = '" & fName
    sqlstr = sqlstr & "' AND middleName = '" & mName
    sqlstr = sqlstr & "' AND lastName = '" & lName
    sqlstr = sqlstr & "') OR nickName = '" & NickName & "'"
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    
    playerExists = Not rs.EOF
    
    rs.Close
    Set rs = Nothing
End Function


Function convertTournamentScheduleTable()
'change the reference in the tables from teamCodeID(Former primary Key from tblTournamentTeamCodes) to teamCode(string, A1, B2 etc)
'
'this makes the relation between schedule and teamcodes more intuitive, allbeit more complex (on two fields: tournamentID AND teamCode)
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = lclConn
        .Open
    End With
    

    Dim rsTn As ADODB.Recordset
    Dim rsCodes As ADODB.Recordset
    Set rsTn = New ADODB.Recordset
    Set rsCodes = New ADODB.Recordset
    Dim sqlstr As String
    sqlstr = "select * from  tblTournamentTeamCodes where teamCodeID > 0"
    rsCodes.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
    Do While Not rsCodes.EOF
        sqlstr = "UPDATE tblGroupLayout SET teamID = " & rsCodes!teamId
        sqlstr = sqlstr & " WHERE teamId = " & rsCodes!teamcodeID
        
'        sqlstr = "Update tblTournamentSchedule SET matchTeamA = '" & rsCodes!teamCode & "'"
'        sqlstr = sqlstr & " WHERE matchTeamA = '" & CStr(rsCodes!teamcodeID) & "'"
'        cn.Execute sqlstr
'        sqlstr = "Update tblTournamentSchedule SET matchTeamB = '" & rsCodes!teamCode & "'"
'        sqlstr = sqlstr & " WHERE matchTeamB = '" & CStr(rsCodes!teamcodeID) & "'"
'        cn.Execute sqlstr
'        sqlstr = "Update tblMatchResults SET TeamA_ID = '" & rsCodes!teamCode & "'"
'        sqlstr = sqlstr & " WHERE TeamA_ID = '" & CStr(rsCodes!teamcodeID) & "'"
'        cn.Execute sqlstr
'        sqlstr = "Update tblMatchResults SET TeamB_ID = '" & rsCodes!teamCode & "'"
'        sqlstr = sqlstr & " WHERE TeamB_ID = '" & CStr(rsCodes!teamcodeID) & "'"
'        cn.Execute sqlstr
'        sqlstr = "Update tblMatchResults SET TeamWinner = '" & rsCodes!teamCode & "'"
'        sqlstr = sqlstr & " WHERE TeamWinner = '" & CStr(rsCodes!teamcodeID) & "'"
'        cn.Execute sqlstr
'        sqlstr = "Update tblPredictionGroupResults SET predictionGroupPosition1 = '" & rsCodes!teamCode & "'"
'        sqlstr = sqlstr & " WHERE predictionGroupPosition1 = '" & CStr(rsCodes!teamcodeID) & "'"
'        cn.Execute sqlstr
'        sqlstr = "Update tblPredictionGroupResults SET predictionGroupPosition2 = '" & rsCodes!teamCode & "'"
'        sqlstr = sqlstr & " WHERE predictionGroupPosition2 = '" & CStr(rsCodes!teamcodeID) & "'"
'        cn.Execute sqlstr
'        sqlstr = "Update tblPredictionGroupResults SET predictionGroupPosition3 = '" & rsCodes!teamCode & "'"
'        sqlstr = sqlstr & " WHERE predictionGroupPosition3 = '" & CStr(rsCodes!teamcodeID) & "'"
'        cn.Execute sqlstr
'        sqlstr = "Update tblPredictionGroupResults SET predictionGroupPosition4 = '" & rsCodes!teamCode & "'"
'        sqlstr = sqlstr & " WHERE predictionGroupPosition4 = '" & CStr(rsCodes!teamcodeID) & "'"
'        cn.Execute sqlstr
'        sqlstr = "Update tblPrediction_Finals SET teamNameA = '" & rsCodes!teamCode & "'"
'        sqlstr = sqlstr & " WHERE teamNameA = '" & CStr(rsCodes!teamcodeID) & "'"
'        cn.Execute sqlstr
'        sqlstr = "Update tblPrediction_Finals SET teamNameB = '" & rsCodes!teamCode & "'"
'        sqlstr = sqlstr & " WHERE teamNameB = '" & CStr(rsCodes!teamcodeID) & "'"
'        cn.Execute sqlstr
'        sqlstr = "Update tblTeamPlayers SET teamID = " & rsCodes!teamId
'        sqlstr = sqlstr & " WHERE teamId = " & rsCodes!teamcodeId
        If Not IsNull(rsCodes!teamId) Then cn.Execute sqlstr
        
        rsCodes.MoveNext
    Loop
    If Not rsTn Is Nothing Then
        rsTn.Close
        Set rsTn = Nothing
    End If
    If Not rsCodes Is Nothing Then
        rsCodes.Close
        Set rsCodes = Nothing
    End If
    
    If Not cn Is Nothing Then
        If (cn.State And adStateOpen) = adStateOpen Then
            cn.Close
        End If
        Set cn = Nothing
    End If
End Function

Function getMatchCount(cn As ADODB.Connection, Optional TournamentId As Long)
  'return number of matches for current tournament or given tournamentID
  Dim sqlstr As String
  Dim rs As ADODB.Recordset
  If Not TournamentId Then TournamentId = thisTournament
  Set rs = New ADODB.Recordset
  sqlstr = "Select COUNT(*) as recAant from tblTournamentSchedule "
  sqlstr = sqlstr & "WHERE tournamentID = " & TournamentId
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    getMatchCount = rs!recAant
  Else
    getMatchCount = 0
  End If
  rs.Close
  Set rs = Nothing
End Function

Function getLastMatchPlayed(cn As ADODB.Connection)
'return the matchOrder number of the last match played
'!!!!!!!!!!!!!!  DO NOT Use MatchNUMBER becasue it can be different then the order of play
Dim sqlstr As String
Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  sqlstr = "Select matchOrder from tblTournamentSchedule where tournamentId = " & thisTournament
  sqlstr = sqlstr & " AND matchPlayed = True Order by matchOrder"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    rs.MoveLast
    getLastMatchPlayed = rs!matchOrder
  Else
    getLastMatchPlayed = 0
  End If
  rs.Close
  Set rs = Nothing
End Function

Function getAllMatchesPlayedOnDay(thisDay As Date, cn As ADODB.Connection) As Boolean
'returns true if all the matches on date thisDay have a final result
Dim sqlstr As String
Dim rs As ADODB.Recordset
Dim matchesToPlay As Integer
  Set rs = New ADODB.Recordset
  sqlstr = "Select count(matchDate) as NumberOfMatches from tblTournamentSchedule where tournamentId = " & thisTournament
  sqlstr = sqlstr & " AND cdbl(matchDate) = " & CDbl(thisDay)
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  matchesToPlay = rs!NumberofMatches
  
  rs.Close
  sqlstr = sqlstr & " AND matchPlayed = true"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If matchesToPlay > 0 Then
    getAllMatchesPlayedOnDay = rs!NumberofMatches = matchesToPlay
  Else
    getAllMatchesPlayedOnDay = False
  End If

  rs.Close
  Set rs = Nothing

End Function

Function getCount(strSQL As String, cn As ADODB.Connection)
  'return number of records in fromTbl
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  rs.Open strSQL, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    rs.MoveLast
    getCount = rs.RecordCount
  Else
    getCount = 0
  End If
  rs.Close
  Set rs = Nothing
End Function

