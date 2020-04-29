VERSION 5.00
Begin VB.Form frmCopyData 
   Caption         =   "Get Data"
   ClientHeight    =   3165
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6045
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
   ScaleHeight     =   3165
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbTournament 
      Height          =   360
      Left            =   4200
      TabIndex        =   6
      Top             =   1027
      Width           =   1575
   End
   Begin VB.CheckBox chkNewDb 
      Alignment       =   1  'Right Justify
      Caption         =   "Nieuwe database aanmaken"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2775
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Sluiten"
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "Start"
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Toernooi"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   1080
      Width           =   855
   End
   Begin VB.Shape shpFill 
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   120
      Top             =   2520
      Width           =   135
   End
   Begin VB.Shape shpBorder 
      Height          =   495
      Left            =   120
      Top             =   2520
      Width           =   4000
   End
   Begin VB.Label lblRecord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Tag             =   "kop"
      Top             =   1800
      Width           =   4000
   End
   Begin VB.Label lblTblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Tag             =   "kop"
      Top             =   1440
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Importeer de tabellen van de mySQL server naar de lokale database vbpool2.mdb"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Tag             =   "kop"
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmCopyData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
Dim msg As String

    msg = "Nieuwe Pool database aanmaken?"
    If chkNewDb Then
        If MsgBox(msg, vbYesNo + vbQuestion, "Nieuwe database") = vbYes Then
        '(re) create an .mdb Access database file, for the local tables
            createDb
        End If
    End If
    If Me.cmbTournament.Text > "" Then
        msg = "Tournooigegevens inlezen van " & Me.cmbTournament.Text
    Else
        msg = "Tournooigegevens inlezen van alle toernooien in de database"
    End If
    If MsgBox(msg, vbYesNo + vbQuestion, "Toernooi gegevens") = vbYes Then
    'add tournament tables to the database from the server
        copyTournamentTables
        'copyData
    End If
End Sub

Private Sub Form_Load()
Dim sqlstr As String
    'fill combobox
    sqlstr = "Select tournamentID, "
    sqlstr = sqlstr & " concat(tournamentYear, ' - ', tournamentType) "
    sqlstr = sqlstr & " as tournament from tblTournaments order by tournamentYear"
    FillCombo Me.cmbTournament, sqlstr, "tournament", "tournamentID", True
    
    UnifyForm Me
    centerForm Me
End Sub

Sub copyTournamentTables()
Dim srcTable As String
Dim newDb As String
Dim newConn As String
Dim rs As ADODB.Recordset
Dim cols As ADODB.Recordset
Dim sqlstr As String
Dim adoTable As adox.Table
Dim adoCatalog As adox.Catalog
Dim tournTable As Boolean

'On Error GoTo connError
    'open connection to mySql
    openMySql
    'get the tables from the mySql table collection
    Set rs = New ADODB.Recordset
    sqlstr = "SHOW TABLES in " & dbName
    rs.Open sqlstr, myConn, adOpenStatic, adLockReadOnly
    If rs.EOF Then
        MsgBox "Geen MySQL tabellen gevonden!", vbOKOnly, "FOUT"
        Exit Sub
    End If
    Set adoCatalog = New adox.Catalog
    adoCatalog.ActiveConnection = cn
    Do While Not rs.EOF
        srcTable = rs.Fields(0)
        If Left(srcTable, 6) <> "local_" Then
            'copy the tabledefs to the mdb
            If Not tableExists(srcTable) Then
                Set adoTable = New adox.Table
                adoTable.Name = srcTable
                Me.lblTblName.Caption = "Tabel: " & rs.AbsolutePosition & "/" & rs.RecordCount
                adoTable.ParentCatalog = adoCatalog
                duplicateFields adoTable, srcTable
                adoCatalog.Tables.Append adoTable
                Set adoTable = Nothing
            End If
        End If
        rs.MoveNext
    Loop
    rs.MoveFirst
    Do While Not rs.EOF
        Set cols = New ADODB.Recordset
        srcTable = rs.Fields(0)
        Me.lblTblName.Caption = "Tabel: " & srcTable
        If Left(srcTable, 6) <> "local_" Then
            cols.Open "SHOW COLUMNS from " & srcTable, myConn, adOpenForwardOnly, adLockReadOnly
            tournTable = False
            Do While Not cols.EOF
                If UCase(cols.Fields(0)) = "TOURNAMENTID" Then
                    tournTable = True
                    Exit Do
                End If
                cols.MoveNext
            Loop
            copyData srcTable, tournTable
        End If
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Set adoCatalog = Nothing
    myConn.Close
    Set myConn = Nothing
    
    Me.lblTblName.Caption = "Klaar! vbpool.MDB ingelezen"
    Me.lblRecord.Caption = ""
    Exit Sub
connError:
    MsgBox "Connectie met mySql server is niet gelukt, database is niet aangemaakt/overschreven", vbOKOnly, "Database aanmaken"
End Sub

Sub copyData(tblName As String, tournTable As Boolean)
    'tournData indicates if only specific tournament data will copied

Dim cmnd As ADODB.Command
Dim rsFrom As ADODB.Recordset
Dim rsTo As ADODB.Recordset
Dim sqlstr As String
Dim dellstr As String
Dim delstr As String
Dim valStr As String
Dim fld As field
    'open mySQL connectin
    openMySql
    
    'open the access database
    If Not cnOpen(cn) Then openDB
    
    Set cmnd = New ADODB.Command
    'open the fromTable
    With cmnd
        .ActiveConnection = myConn
        .CommandType = adCmdText
        sqlstr = "Select * from " & tblName
        delstr = "Delete from " & tblName
        If tournTable And Me.cmbTournament.ListIndex > 0 Then
            'only copy records for seleted tournament
            sqlstr = sqlstr & " WHERE tournamentID = " & Me.cmbTournament.ItemData(Me.cmbTournament.ListIndex)
            delstr = delstr & " WHERE tournamentID = " & Me.cmbTournament.ItemData(Me.cmbTournament.ListIndex)
        End If
        .CommandText = sqlstr
        Set rsFrom = .Execute
    End With
    'delete records from local table
    cn.Execute delstr
    'add to the toTable
    Set rsTo = New ADODB.Recordset
    rsTo.Open "Select * from " & tblName, cn, adOpenKeyset, adLockOptimistic
    Do While Not rsFrom.EOF  'loop through records
        rsTo.AddNew
        Me.shpFill.Width = rsFrom.AbsolutePosition * (Me.shpBorder.Width / rsFrom.RecordCount)
        Me.lblRecord.Caption = "Record " & rsFrom.AbsolutePosition & "/" & rsFrom.RecordCount
        DoEvents
        For Each fld In rsFrom.Fields  'loop through fields
            If fld.Name = "moneydaylast" Then Stop
            If Not IsNull(fld.value) Then
                rsTo(fld.Name) = fld.value
            Else
                If rsTo(fld.Name).Attributes = 70 Or rsTo(fld.Name).Attributes = 86 Then
                    If rsTo(fld.Name).Type = adVarWChar Then
                        rsTo(fld.Name) = ""
                    Else
                        rsTo(fld.Name) = 0
                    End If
                End If
            End If
        Next
        rsTo.Update
        rsFrom.MoveNext 'next record
    Loop
nextTable:
    'tidy up
    rsFrom.Close
    Set rsFrom = Nothing
    Set cmnd = Nothing
    rsTo.Close
    Set rsTo = Nothing
    myConn.Close
    Set myConn = Nothing
End Sub

Sub duplicateFields(toTable As adox.Table, fromTbl As String)
    'copy tbl fields to Access database
    Dim rs As ADODB.Recordset  'to store the columns
    Dim col As adox.Column
    Dim sqlstr As String
    Dim ln As Integer
    Dim fldName As String
    openMySql
    'get all tables from the server
    Set rs = New ADODB.Recordset
    sqlstr = "SHOW COLUMNS in " & fromTbl & " in " & dbName
    rs.Open sqlstr, myConn, adOpenStatic, adLockReadOnly
    'copy the field defintion
    
    With toTable
        Do While Not rs.EOF
            fldName = rs.Fields(0).value
            Set col = New adox.Column
            col.Name = fldName
            col.Type = cFieldType(rs.Fields("Type"))
            .Columns.Append col
            If InStr(LCase(rs.Fields("Type")), "varchar") Then
                ln = Val(Mid(rs.Fields("Type"), 9, Len(rs.Fields("Type")) - 9))
                .Columns(fldName).DefinedSize = ln
            End If
            If LCase(rs.Fields("Extra")) = "auto_increment" And rs.Fields("Type") = "int(11)" Then
                .Columns(fldName).Properties("AutoIncrement").value = True
                .Keys.Append "PrimaryKey", adKeyPrimary, fldName
            End If
            rs.MoveNext
        Loop
    End With
    
    'release from memory
    rs.Close
    Set rs = Nothing
    Set col = Nothing
    myConn.Close
    Set myConn = Nothing
End Sub


