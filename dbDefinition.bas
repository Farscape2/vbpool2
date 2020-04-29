Attribute VB_Name = "dbDefinition"
''''''''''
'Routines and functions for the definition and copying the database

Option Explicit

Public cn As ADODB.Connection      'Access connection
Public myConn As ADODB.Connection  'MySql Connection

Public Const dbName = "vbpool2"

Public Const dBaseType = "ACCESS"

Sub openDB()

'open local database connection to the access mdb file
Dim fullPath As String
    fullPath = App.Path & "\" & dbName & ".mdb"
    If Dir(fullPath) = "" Then
        'frmCopyData.Show 1
    End If
    Set cn = New ADODB.Connection
    
    With cn
    '''''''ACCESS Connection
        .Provider = "Microsoft.Jet.OLEDB.4.0;"
        .ConnectionString = "Data Source=" & fullPath
        .CursorLocation = adUseClient
        .Open
    End With
End Sub

Sub openMySql()
'open mySql server connection
    Dim server As String
    Dim driver As String
    Dim cnstr As String
    Dim passwd As String
    passwd = "!xjer56!"
    server = "192.168.178.14"
    'server = "jotaservices.duckdns.org"
    driver = "{MariaDB ODBC 3.1 Driver}"
    Set myConn = New ADODB.Connection
    With myConn
        cnstr = "DRIVER=" & driver & ";TCPIP=1;SERVER=" & server & ";DATABASE=" & dbName & ";UID=jeroen;PWD=" & passwd & ";port=3306"
        .ConnectionString = cnstr
        .CursorLocation = adUseClient
        .Open
    End With
End Sub

Function tableExists(srcTable As String)
'check if table exists in local database
Dim rs As ADODB.Recordset
    If Not cnOpen(cn) Then openDB
    Set rs = cn.OpenSchema(adSchemaColumns, Array(Empty, Empty, srcTable, Empty))
    tableExists = Not (rs.BOF And rs.EOF)
    rs.Close
    Set rs = Nothing
End Function

Function recordsExist(tblName As String)
    Dim rs As ADODB.Recordset
    If tableExists(tblName) Then
        Set rs = New ADODB.Recordset
        rs.Open "Select * from " & tblName, cn, adOpenKeyset, adLockReadOnly
        recordsExist = Not rs.EOF
    Else
        recordsExist = False
    End If
End Function

Function cFieldType(fldType As String) As Integer
'convert mySQL fldType to ADODB type
    Dim returnType As Integer
    If Left(fldType, 7) = "varchar" Then
        returnType = adVarWChar  'default type
    Else
        Select Case LCase(fldType)
        Case "Date", "time", "datetime", "timestamp"
            returnType = adDate
        Case "int(11)"
            returnType = adInteger
        Case "double"
            returnType = adDouble
        Case "decimal(19,4)"
            returnType = adCurrency
        Case "tinyint(3)", "tinyint(3) unsigned"
            returnType = adUnsignedTinyInt
        Case "tinyint(1)"
            returnType = adBoolean
        Case Else
            Stop
        End Select
    End If
    cFieldType = returnType
End Function


'create the database
Sub createDb()
    Dim adoCat As adox.Catalog
    Dim tbl As adox.Table
    Dim setupDb As String
    Dim newDb As String
    Dim msg As String
    Dim fileName As String
    
    ' MDB to be created. In app.path
    newDb = App.Path & "\" & dbName & ".mdb"
    ' Drop the existing database, if any.
    If Dir(newDb) > "" Then
        msg = "Er is al een database " & newDb & vbNewLine
        msg = msg & "Wil je een kopie hiervan bewaren?" & vbNewLine
        If MsgBox(msg, vbYesNo, "Nieuwe database aanmaken") = vbYes Then
            FileCopy newDb, newDb & ".bak"
        End If
        Kill newDb 'remove the old db
    End If
    setupDb = App.Path & "\vbpoolSetup.mdb"
    If Dir(setupDb) = "" Then 'no setupDb, make one
        Set adoCat = New adox.Catalog
        ' Create the db
        adoCat.Create ("Provider='Microsoft.Jet.OLEDB.4.0';Data Source=" & newDb & ";")
        'add local tables to db
        makeTables
    Else
        FileCopy setupDb, newDb
    End If
    MsgBox "Nieuwe database is aangemaakt." & vbNewLine & "Vul de gegevens in en kies een wachtwoord", vbOKOnly + vbInformation, "Nieuwe installatie"
    If Not cnOpen(cn) Then openDB
    frmOrganisation.Show 1
    fillDefaultValues
End Sub


Sub fillDefaultValues()
'fill some tables with default values
    Dim rs As ADODB.Recordset
    Dim sqlstr As String
    Dim cmd As ADODB.Command
    Dim orgID As Long
    If Not cnOpen(cn) Then openDB
    Set rs = New ADODB.Recordset
    sqlstr = "Select * from tblTournaments ORDER BY tournamentStartDate"
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If rs.EOF Then
        MsgBox "Database problem", vbOKOnly + vbCritical, "Contact Jota"
        Exit Sub
    End If
    'get the last tournament ID
    rs.MoveLast
    thisTournament = rs!tournamentId
    rs.Close
    'get the OrganisationID - should be only one organisation
    orgID = getOrganisation("organisationID")
    'create a first record in tblPools
    sqlstr = "INSERT INTO tblPools (tournamentId, OrganisationId) VALUES (" & thisTournament & ", " & orgID & ")"
    cn.Execute sqlstr
    'default data for thispool
    rs.Open "Select * from tblPools"
        rs.MoveLast 'just in case
        rs!Poolname = "Eerste pool"
        rs!poolFormsFrom = Date
        rs!poolformstill = getTournamentInfo("tournamentEndDate") - 7
        rs!poolcost = 10
        rs!prizeHighDayScore = 2.5
        rs!prizeHighDayPosition = 1
        rs!prizeLowDayPosition = 0.1
        rs!prizePercentageFirst = 0.5
        rs!prizePercentageSecond = 0.3
        rs!prizePercentageThird = 0.2
        rs!prizePercentageFourth = 0
        rs!prizeLowFinalPosition = 10
    rs.Update
    rs.Close
    'default data for points
    Set cmd = New ADODB.Command
    'construct command to fill poolpoints with pointtypes
    
    
End Sub
Sub makeTables()
    Dim sqlstr As String
    'new local tables just in case
    If Not cnOpen(cn) Then openDB 'set the connection
    
    'address table
    sqlstr = "CREATE TABLE tblAddress ( "
    sqlstr = sqlstr & "addressID INTEGER NOT NULL, "
    sqlstr = sqlstr & "firstname VARCHAR(50), "
    sqlstr = sqlstr & "middlename VARCHAR(32), "
    sqlstr = sqlstr & "lastname VARCHAR(50), "
    sqlstr = sqlstr & "shortname VARCHAR(24), "
    sqlstr = sqlstr & "address VARCHAR(50), "
    sqlstr = sqlstr & "postalcode VARCHAR(10), "
    sqlstr = sqlstr & "city VARCHAR(50), "
    sqlstr = sqlstr & "telephone VARCHAR(20), "
    sqlstr = sqlstr & "email VarChar(255) "
    sqlstr = sqlstr & ") "
    cn.Execute sqlstr
    sqlstr = "CREATE INDEX PrimaryKey on tblAddress (addressID) WITH PRIMARY"
    cn.Execute sqlstr
    
    'competitorpoints
    sqlstr = "CREATE TABLE tblCompetitorPoints ("
    sqlstr = sqlstr & "competitorID INTEGER NOT NULL,"
    sqlstr = sqlstr & "matchNumber INTEGER NOT NULL,"
    sqlstr = sqlstr & "pointsMatchTeams INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsGroupStanding INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsFinals_8 INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsFinals_4 INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsFinals_2 INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsFinals_34 INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsFinal INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsMatchResults INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTopscorers INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsOther INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsDayTotal INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsGrandTotal INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "poisitionDay INTEGER,"
    sqlstr = sqlstr & "positionTotal INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "moneyDay DECIMAL(19,4) DEFAULT 0,"
    sqlstr = sqlstr & "moneyDayPosition DECIMAL(19,4) DEFAULT 0,"
    sqlstr = sqlstr & "moneyDayLast DECIMAL(19,4),"
    sqlstr = sqlstr & "moneyTotal DECIMAL(19,4) DEFAULT 0,"
    sqlstr = sqlstr & "moneyDayTotal DECIMAL(19,4) DEFAULT 0,"
    sqlstr = sqlstr & "pointsDay INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "positionMatches INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "poinstDayGoals INTEGER,"
    sqlstr = sqlstr & "pointsHalfTime INTEGER,"
    sqlstr = sqlstr & "pointsFulltime INTEGER,"
    sqlstr = sqlstr & "pointsToto INTEGER,"
    sqlstr = sqlstr & "pointsGrpA INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsGrpB INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsGrpC INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsGrpD INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsGrpE INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsGrpF INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsGrpG INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsGrpH INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals8A INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals8B INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals8C INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals8D INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals8E INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals8F INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals8G INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals8H INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals4A INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals4B INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals4C INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals4D INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals2A INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals2B INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals34 INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinal INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTotalAfterFinal34 INTEGER DEFAULT 0"
    sqlstr = sqlstr & ")"
    cn.Execute sqlstr
    
    'deelnemers
    sqlstr = "CREATE TABLE tblCompetitors ("
    sqlstr = sqlstr & "competitorID INTEGER NOT NULL, "
    sqlstr = sqlstr & "poolid INTEGER NOT NULL,"
    sqlstr = sqlstr & "addressID INTEGER NOT NULL,"
    sqlstr = sqlstr & "nickName VARCHAR(50) NOT NULL,"
    sqlstr = sqlstr & "payed YESNO DEFAULT 0,"
    sqlstr = sqlstr & "predictionTeam1 INTEGER,"
    sqlstr = sqlstr & "predictionTeam2 INTEGER,"
    sqlstr = sqlstr & "predictionTeam3 INTEGER,"
    sqlstr = sqlstr & "predictionTeam4 INTEGER"
    sqlstr = sqlstr & ") "
    cn.Execute sqlstr
    sqlstr = "CREATE INDEX PrimaryKey on tblCompetitors (competitorID) WITH PRIMARY"
    cn.Execute sqlstr

    'pool points
    sqlstr = "CREATE TABLE tblPoolPoints ("
    sqlstr = sqlstr & "poolid INTEGER NOT NULL,"
    sqlstr = sqlstr & "pointTypeID INTEGER NOT NULL,"
    sqlstr = sqlstr & "pointPointsAward INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointPointsMargin byte DEFAULT 0 )"
    cn.Execute sqlstr

    'pools
    sqlstr = "CREATE TABLE tblPools ("
    sqlstr = sqlstr & "poolID INTEGER NOT NULL DEFAULT 0,"
    sqlstr = sqlstr & "tournamentID INTEGER DEFAULT NULL,"
    sqlstr = sqlstr & "organisationID INTEGER DEFAULT NULL,"
    sqlstr = sqlstr & "poolName varchar(50) DEFAULT NULL,"
    sqlstr = sqlstr & "poolStartAcceptForms datetime DEFAULT NULL,"
    sqlstr = sqlstr & "poolEndAcceptForms datetime DEFAULT NULL,"
    sqlstr = sqlstr & "poolCost decimal(19,4) DEFAULT 10.0000,"
    sqlstr = sqlstr & "prizeHighDayScore decimal(19,4) DEFAULT 0.0000,"
    sqlstr = sqlstr & "prizeHighDayPosition decimal(19,4) DEFAULT 0.0000,"
    sqlstr = sqlstr & "prizeLowDayPosition decimal(19,4) DEFAULT 0.0000,"
    sqlstr = sqlstr & "prizePercentageFirst double DEFAULT 0,"
    sqlstr = sqlstr & "prizePercentageSecond double DEFAULT 0,"
    sqlstr = sqlstr & "prizePercentageThird double DEFAULT 0,"
    sqlstr = sqlstr & "prizePercentageFourth double DEFAULT 0,"
    sqlstr = sqlstr & "prizeLowFinalPosition decimal(19,4) DEFAULT 0.0000"
    sqlstr = sqlstr & ")"
    cn.Execute sqlstr
    sqlstr = "CREATE INDEX PrimaryKey on tblPools (poolID) WITH PRIMARY"
    cn.Execute sqlstr
    
    'predictions - Groups
    sqlstr = "CREATE TABLE tblPredictionGroupResults ("
    sqlstr = sqlstr & "competitorID INTEGER NOT NULL,"
    sqlstr = sqlstr & "groupLetter varchar(1) DEFAULT NULL,"
    sqlstr = sqlstr & "predictionGroupPosition1 varchar(255) DEFAULT NULL,"
    sqlstr = sqlstr & "predictionGroupPosition2 varchar(255) DEFAULT NULL,"
    sqlstr = sqlstr & "predictionGroupPosition3 varchar(255) DEFAULT NULL,"
    sqlstr = sqlstr & "predictionGroupPosition4 varchar(255) DEFAULT NULL"
    sqlstr = sqlstr & ")"
    cn.Execute sqlstr


    sqlstr = "CREATE TABLE tblPredictionTopscorers ("
    sqlstr = sqlstr & "competitorID INTEGER NOT NULL,"
    sqlstr = sqlstr & "predictionTopscorerPosittion INTEGER DEFAULT NULL,"
    sqlstr = sqlstr & "predictionTopscorePlayerID INTEGER DEFAULT NULL,"
    sqlstr = sqlstr & "predictionTopscoreGoals INTEGER DEFAULT NULL"
    sqlstr = sqlstr & ")"
    cn.Execute sqlstr

    sqlstr = "CREATE TABLE tblPrediction_Finals ("
    sqlstr = sqlstr & "competitorID INTEGER NOT NULL,"
    sqlstr = sqlstr & "matchNumber INTEGER DEFAULT NULL,"
    sqlstr = sqlstr & "teamNameA INTEGER DEFAULT NULL,"
    sqlstr = sqlstr & "teamNameB INTEGER DEFAULT NULL"
    sqlstr = sqlstr & ")"
    cn.Execute sqlstr
    
    sqlstr = "CREATE TABLE tblPrediction_MatchResults ("
    sqlstr = sqlstr & "competitorID INTEGER DEFAULT NULL,"
    sqlstr = sqlstr & "matchNumber INTEGER DEFAULT NULL,"
    sqlstr = sqlstr & "predictionGoalsHalftimeA BYTE DEFAULT NULL,"
    sqlstr = sqlstr & "predictionGoalsHalftimeB BYTE DEFAULT 0,"
    sqlstr = sqlstr & "predictionGoalsFulltimeA BYTE DEFAULT 0,"
    sqlstr = sqlstr & "predictionGoalsFulltimeB BYTE DEFAULT 0,"
    sqlstr = sqlstr & "predictionResultToto BYTE DEFAULT NULL"
    sqlstr = sqlstr & ")"
    cn.Execute sqlstr
    
    sqlstr = "CREATE TABLE tblPrediction_Numbers ("
    sqlstr = sqlstr & "competitorID INTEGER NOT NULL,"
    sqlstr = sqlstr & "predictionTypeID INTEGER DEFAULT NULL,"
    sqlstr = sqlstr & "predictionNumber INTEGER DEFAULT NULL"
    sqlstr = sqlstr & ")"
    cn.Execute sqlstr
    
    sqlstr = "CREATE TABLE tblUsers ("
    sqlstr = sqlstr & "userID INTEGER NOT NULL, "
    sqlstr = sqlstr & "username VARCHAR(50), "
    sqlstr = sqlstr & "Passwd VARCHAR(50) NOT NULL"
    sqlstr = sqlstr & ")"
    cn.Execute sqlstr
    
End Sub

Function cnOpen(adoCn As ADODB.Connection) As Boolean

    '----------------------------------------------------------------
    '#PURPOSE: Checks whether the supplied db connection is alive and
    '          hasn't had it's TCP connection forcibly closed by remote
    '          host, for example, as happens during an undock event
    '#RETURNS: True if the supplied db is connected and error-free,
    '          False otherwise
    '#AUTHOR:  Belladonna
    '----------------------------------------------------------------

    Dim i As Long
    Dim cmd As New ADODB.Command

    'Set up SQL command to return 1
    cmd.CommandText = "SELECT 1"
    On Error GoTo endFunction
    cmd.ActiveConnection = adoCn

    'Run a simple query, to test the connection
    On Error Resume Next
    i = cmd.Execute.Fields(0)
    On Error GoTo 0

    'Tidy up
    Set cmd = Nothing

    'If i is 1, connection is open
    If i = 1 Then
        cnOpen = True
    Else
        cnOpen = False
    End If
    Exit Function
endFunction:
    cnOpen = False
End Function
