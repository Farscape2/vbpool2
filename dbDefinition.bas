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
        createDb
    End If
    
    Set cn = New ADODB.Connection
    If cn.State = 1 Then cn.Close
    
    With cn
    '''''''ACCESS Connection
        .Provider = "Microsoft.Jet.OLEDB.4.0;"
        .ConnectionString = "Data Source=" & fullPath
        .Open
    End With
End Sub

Sub openMySql()
'open mySql server connection
    Dim server As String
    Dim driver As String
    Dim cnstr As String
    server = "192.168.178.14"
    'server = "jotaservices.duckdns.org"
    driver = "{MariaDB ODBC 3.1 Driver}"
    Set myConn = New ADODB.Connection
    If myConn.State = 1 Then myConn.Close
    With myConn
        cnstr = "DRIVER=" & driver & ";TCPIP=1;SERVER=" & server & ";DATABASE=" & dbName & ";UID=jeroen;PWD=!xjer56!;port=3306"
        .ConnectionString = cnstr
        .CursorLocation = adUseClient
        .Open
    End With
End Sub

Sub createDb()
'(re) create an .mdb Access database file, based on the base server MySql connection
Dim adoCatalog As ADOX.Catalog
Dim adoTable As ADOX.Table
Dim newTable As String
Dim newDb As String
Dim newConn As String
Dim rs As ADODB.Recordset
Dim sqlStr As String

On Error GoTo connError
    Set myConn = New ADODB.Connection
    If (myConn.State And adStateClosed) = adStateClosed Then
        openMySql
    End If
    'open connection to mySql
    'get the tables from the mySql table collection
    Set rs = New ADODB.Recordset
    sqlStr = "SHOW TABLES in " & dbName
    rs.Open sqlStr, myConn, adOpenStatic, adLockReadOnly
    If rs.EOF Then
        MsgBox "Geen MySQL tabellen gevonden!", vbOKOnly, "FOUT"
        Exit Sub
    End If
    
    ' MDB to be created. In app.path
    newDb = App.Path & "\" & dbName & ".mdb"
    ' Drop the existing database, if any.
    On Error Resume Next 'in case not found
    If (cn.State And adStateOpen) = adStateOpen Then
        cn.Close
    End If
    Kill newDb
    On Error GoTo 0
    
    'Create instance of the ADOX-object.
    Set adoCatalog = New ADOX.Catalog
    ' Create the db
    newConn = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & newDb & ";"

    adoCatalog.Create (newConn)
    
    Do While Not rs.EOF
        newTable = rs.Fields(0)
        'copy the tabledefs to the mdb
        Set adoTable = New ADOX.Table
        adoTable.Name = newTable
        adoTable.ParentCatalog = adoCatalog
        If newTable = "tblGroupLayout" Then
            'why doesn't this one work
            'Stop
        End If
        duplicateFields adoTable, newTable
        adoCatalog.Tables.Append adoTable
        rs.MoveNext
        Set adoTable = Nothing
    Loop
    
    rs.Close
    Set rs = Nothing
    Set adoCatalog = Nothing
    
    MsgBox "New vbpool.MDB Created - '" & newDb & "'", vbInformation
    Exit Sub
connError:
    MsgBox "Connectie met mySql server is niet gelukt, database is niet aangemaakt/overschreven", vbOKOnly, "Database aanmaken"
End Sub

Sub duplicateFields(toTable As ADOX.Table, fromTbl As String)
    'copy tbl fields to Access database
    Dim rs As ADODB.Recordset  'to store the columns
    Dim col As ADOX.Column
    Dim sqlStr As String
    Dim ln As Integer
    Dim fldName As String
    'get all tables from the server
    Set rs = New ADODB.Recordset
    sqlStr = "SHOW COLUMNS in " & fromTbl & " in " & dbName
    rs.Open sqlStr, myConn, adOpenStatic, adLockReadOnly
    'copy the field defintion
    
    With toTable
        Do While Not rs.EOF
            fldName = rs.Fields(0).Value
            Set col = New ADOX.Column
            col.Name = fldName
            col.Type = cFieldType(rs.Fields("Type"))
            .Columns.Append col
            If InStr(LCase(rs.Fields("Type")), "varchar") Then
                ln = Val(Mid(rs.Fields("Type"), 9, Len(rs.Fields("Type")) - 9))
                .Columns(fldName).DefinedSize = ln
            End If
            If LCase(rs.Fields("Extra")) = "auto_increment" And rs.Fields("Type") = "int(11)" Then
                .Columns(fldName).Properties("AutoIncrement").Value = True
                .Keys.Append "PrimaryKey", adKeyPrimary, fldName
            End If
            rs.MoveNext
        Loop
    End With
    
    'release from memory
    rs.Close
    Set rs = Nothing
    Set col = Nothing
    
End Sub

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


