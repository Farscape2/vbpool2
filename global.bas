Attribute VB_Name = "global"
Option Explicit

'currentPool is read and stored in dbFunctions module
Public thisPool As Long
Public thisTournament As Long
'variable to preserve the current active country
Public currentCountry As Long  'used to pass information between forms
Public adminLogin As Boolean

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub Main()
    
    'commandline arguments
    Dim i As Integer
    Dim strArgs() As String
    ' check if we started the app as admin
    strArgs = Split(Command$, " ")
    For i = 0 To UBound(strArgs)
        If strArgs(i) = "admin" Then
            adminLogin = True
            Exit For
        End If
    Next
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    write2Log "App started", True
    'check other instance of app
    If App.PrevInstance = True Then
        MsgBox "VBPool2.0 draait al...."
        Exit Sub
    End If
    'set and open the database
    If Dir(App.Path & "\" & dbName & ".mdb") = "" Then
        createDb
        write2Log "No vbpool2.mdb, dbcreated"
    End If
    'now that the database is crated we can open the connection
    With cn
        .ConnectionString = lclConn()
        .Open
    End With
    
    'if there is a pools table with at least one record
    If recordsExist("tblPools", cn) Then
        ' get last poolID
        thisPool = val(GetSetting(App.EXEName, "global", "lastpool", 0))
    End If
    If thisPool Then
        thisTournament = getThisPoolTournamentId(cn)
    End If
    cn.Close
    Set cn = Nothing
    'open main form
    mainForm.Show
    If Not cn Is Nothing Then
        If (cn.State And adStateOpen) = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
End Sub

Sub UnifyForm(frm As Form, Optional center As Boolean)
'basic format for all forms
    Dim ctl As Control
    For Each ctl In frm.Controls
        On Error Resume Next 'if property does not exist
        ctl.Font.Name = "Tahoma"
        ctl.Font.Size = 10
        
        If InStr(ctl.Tag, "kop") Then 'small heading
            ctl.Font.Name = "Times New Roman"
            ctl.Font.Size = 14
            If InStr(ctl.Tag, "kop2") Then 'larger heading
                ctl.Font.Size = 20
            End If
            If InStr(ctl.Tag, "kop1") Then  'large heading
                ctl.Font.Size = 32
            End If
        End If
        
        If TypeOf ctl Is Label Then
            ctl.ForeColor = &H4000&  'dark green
        End If
        If TypeOf ctl Is CheckBox Then
            ctl.BackColor = frm.BackColor
        End If
        If InStr(ctl.Tag, "small") Then  'used for ©opyright message
 '           ctl.ForeColor = vbBlue
            ctl.Font.Size = 11
            ctl.Font.Name = "Garamond"
        End If
    Next
End Sub

Sub centerForm(frm As Object)
   frm.Move (Screen.Width - frm.Width) / 2, (Screen.Height - frm.Height) / 2
End Sub

Function float(strNumber As String) As String
'convert formatted dutch float number to dot seperated value
    Dim number As String
    If InStr(strNumber, "%") Then
        strNumber = val(Left(strNumber, Len(strNumber) - 1)) / 100
    End If
    
    If Not IsNumeric(strNumber) Then
        Exit Function
    Else
        float = Replace(strNumber, ",", ".")
    End If
End Function

Public Function setCombo(objCmb As ComboBox, val As Variant)
    'set the combo listitem based on val in the listindex
    Dim i As Integer
    With objCmb
        Do While Not .ItemData(i) = val
            i = i + 1
        Loop
        objCmb.ListIndex = i
    End With
End Function


Public Sub FillCombo(objComboBox As ComboBox, _
                     strSql As String, _
                     cn As ADODB.Connection, _
                     strFieldToShow As String, _
                     Optional strFieldForItemData As String)

'Fills a combobox with values from a database
    
    Dim oRS As ADODB.Recordset  'Load the data
    Set oRS = New ADODB.Recordset
    oRS.Open strSql, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If oRS.EOF Then
        MsgBox "Geen records in recordset", vbCritical + vbOKOnly, "FillCombo"
        Exit Sub
    End If
    With objComboBox          'Fill the combo box
        .Clear
        If strFieldForItemData = "" Then
            Do While Not oRS.EOF      '(without ItemData)
                .AddItem oRS.Fields(strFieldToShow).value
                oRS.MoveNext
            Loop
        Else
            Do While Not oRS.EOF      '(with ItemData)
                .AddItem oRS.Fields(strFieldToShow).value
                .ItemData(.NewIndex) = oRS.Fields(strFieldForItemData).value
                oRS.MoveNext
            Loop
        End If
    End With
    
    oRS.Close                 'Tidy up
    Set oRS = Nothing

End Sub

Sub fillList(objListBox As ListBox, _
              strSql As String, _
              cn As ADODB.Connection, _
              strFieldToShow As String, _
              Optional strFieldForItemData As String)

    Dim oRS As ADODB.Recordset  'Load the data
    Set oRS = New ADODB.Recordset
    oRS.Open strSql, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If oRS.EOF Then
        Exit Sub
    End If
    With objListBox          'Fill the list box
        .Clear
        If strFieldForItemData = "" Then
            Do While Not oRS.EOF      '(without ItemData)
                .AddItem oRS.Fields(strFieldToShow).value
                oRS.MoveNext
            Loop
        Else
            Do While Not oRS.EOF      '(with ItemData)
                .AddItem oRS.Fields(strFieldToShow).value
                .ItemData(.NewIndex) = oRS.Fields(strFieldForItemData).value
                oRS.MoveNext
            Loop
        End If
    End With
    
    oRS.Close                 'Tidy up
    Set oRS = Nothing


End Sub

Public Function DoLogin() As Boolean

'login system originally from Michael Ciurescu (CVMichael from vbforums.com)
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = lclConn()
        .Open
    End With
    

    Dim UserName As String, Password As String, Ret As Boolean
    Dim LoginSuccessful As Boolean, rsData As ADODB.Recordset
    Dim MD5 As New clsMD5
    
    Randomize
    
    ' Get the user that last logged in from the registry
    UserName = getOrganisation(cn, "lastname")
        
    ' prompt user to enter username and password
    Ret = frmAdminLogin.GetLogIn(UserName, Password)
    
    Do While Ret
        Set rsData = cn.Execute("SELECT Passwd FROM tblOrganisation WHERE lastname = '" & Replace(UserName, "'", "''") & "'")
        
        ' if a record was found, it means the user exists
        If Not rsData.EOF Then
            ' check if the password is correct
            If UCase(MD5.DigestStrToHexStr(Password)) = UCase(rsData("Passwd").value) Then
                
                LoginSuccessful = True
                Exit Do
            End If
        End If
        
        If Not LoginSuccessful Then
            Ret = False
            
            If MsgBox("Wachtwoord onjuist, nog eens proberen?", vbQuestion + vbYesNo, "Login mislukt") = vbYes Then
                ' to prevent brute force password cracking from the application
                Sleep 200 + 300 * Rnd
                
                ' if login was not successfull, prompt again until Cancel is clicked
                Ret = frmAdminLogin.GetLogIn(UserName, Password)
            End If
        End If
    Loop
    If Not LoginSuccessful Then
        write2Log "Login failed", True
    Else
        write2Log "Login successfull", True
    End If
    DoLogin = LoginSuccessful
    
    cn.Close
    Set cn = Nothing
End Function

'add the nz function
Public Function nz(strValue As Variant, Optional alternative As String = "") As Variant
    If Not IsNull(strValue) Then
        nz = strValue
    Else
        nz = alternative
    End If
End Function

Public Sub write2Log(txt, Optional timekolom As Boolean)
Dim iFileNr As Integer
Dim filenaam As String
Dim timestamp  As String

    iFileNr = FreeFile
    filenaam = App.Path & "\vbpool20.log"
    If timekolom Then
        timestamp = Format(Now(), "YYYY-MM-DD hh:nn:ss")
    Else
        timestamp = Space(20)
    End If
    
    Open filenaam For Append As #iFileNr
        Print #iFileNr, timestamp, txt
    Close #iFileNr

End Sub

Sub getTournamentTables()
Dim srcTable As String
Dim rsTables As ADODB.Recordset
Dim rsCols As ADODB.Recordset
Dim sqlstr As String
Dim tournTable As Boolean
Dim myConn As ADODB.Connection
    
    'get the tables from the mySql table collection
    Set rsTables = New ADODB.Recordset
    Set myConn = New ADODB.Connection
    With myConn
        .CursorLocation = adUseClient
        .ConnectionString = mySqlConn
        .Open
    End With
    sqlstr = "Select tournamentID from tblTournaments order by tournamentStartDate"
    rsTables.Open sqlstr, myConn, adOpenKeyset, adLockReadOnly
    If rsTables.EOF Then
        MsgBox "Geen verbinding gemaakt of geen gegevens gevonden!" & vbNewLine & "Kan niet verder gaan", vbOKOnly + vbCritical, "Database probleem"
        Exit Sub
    End If
    rsTables.MoveLast
    thisTournament = rsTables!TournamentId
    rsTables.Close
    'Use different sql in rsTablses now
    sqlstr = "SHOW TABLES in " & dbName
    rsTables.Open sqlstr, myConn, adOpenStatic, adLockReadOnly
    If rsTables.EOF Then
        MsgBox "Geen MySQL tabellen gevonden!", vbOKOnly, "FOUT"
        Exit Sub
    End If
    'get the id of the last tournament
    rsTables.MoveFirst
    Do While Not rsTables.EOF
        Set rsCols = New ADODB.Recordset
        srcTable = rsTables.Fields(0)
        If Left(srcTable, 6) <> "local_" Then
'            Me.lblTblName.Caption = "Tabel: " & srcTable
            'open connection to mySql
            rsCols.Open "SHOW COLUMNS from " & srcTable, myConn, adOpenForwardOnly, adLockReadOnly
            tournTable = False
            Do While Not rsCols.EOF 'check if there is a field for tournamentID, if so copy only data for this tournament
                If UCase(rsCols.Fields(0)) = "TOURNAMENTID" Then
                    tournTable = True
                    Exit Do
                End If
                rsCols.MoveNext
            Loop
            rsCols.Close
            copyTournamentData srcTable, tournTable, myConn
        End If
        rsTables.MoveNext
    Loop
    
    If Not rsTables Is Nothing Then
        If (rsTables.State And adStateOpen) = adStateOpen Then rsTables.Close
        Set rsTables = Nothing
    End If
    If Not rsCols Is Nothing Then
        If (rsCols.State And adStateOpen) = adStateOpen Then rsCols.Close
        Set rsCols = Nothing
    End If
    If Not myConn Is Nothing Then
        If (myConn.State And adStateOpen) = adStateOpen Then myConn.Close
        Set myConn = Nothing
    End If
'    Me.lblTblName.Caption = "Klaar! Alles ingelezen"
'    Me.lblRecord.Caption = ""
End Sub

Sub copyTournamentData(tblName As String, tournTable As Boolean, myConn As ADODB.Connection)
    'tournTable indicates if only specific tournament data will copied

Dim cmnd As ADODB.Command
Dim rsFrom As ADODB.Recordset
Dim rsTo As ADODB.Recordset
Dim sqlstr As String
Dim dellstr As String
Dim delstr As String
Dim valStr As String
Dim fld As field
Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = lclConn
        .Open
    End With
    
    Set cmnd = New ADODB.Command
    'open the fromTable
    With cmnd
        .ActiveConnection = myConn
        .CommandType = adCmdText
        sqlstr = "Select * from " & tblName
        delstr = "Delete from " & tblName
        If tournTable Then
            'only copy records for seleted tournament
            sqlstr = sqlstr & " WHERE tournamentID = " & thisTournament
            delstr = delstr & " WHERE tournamentID = " & thisTournament
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
        'show info on form
        'Me.shpFill.Width = rsFrom.AbsolutePosition * (Me.shpBorder.Width / rsFrom.RecordCount)
        'Me.lblRecord.Caption = "Record " & rsFrom.AbsolutePosition & "/" & rsFrom.RecordCount
        DoEvents
        For Each fld In rsFrom.Fields  'loop through fields
            If Not IsNull(fld.value) Then
                rsTo(fld.Name) = fld.value
            Else
                If rsTo(fld.Name).Attributes = 70 Or rsTo(fld.Name).Attributes = 86 Then
                'if the field can not be NULL / just in case
                    If rsTo(fld.Name).Type = adVarWChar Then
                        rsTo(fld.Name) = "" 'set it to empty string
                    Else
                        rsTo(fld.Name) = 0 'set it to 0
                    End If
                End If
            End If
        Next
        rsTo.Update
        rsFrom.MoveNext 'next record
    Loop
    'tidy up
    
    If Not rsFrom Is Nothing Then
        If (rsFrom.State And adStateOpen) Then rsFrom.Close
        Set rsFrom = Nothing
    End If
    If Not cmnd Is Nothing Then
        Set cmnd = Nothing
    End If
    
    If Not rsTo Is Nothing Then
        If (rsTo.State And adStateOpen) = adStateOpen Then rsTo.Close
        Set rsTo = Nothing
    End If
    If Not cn Is Nothing Then
        If (cn.State And adStateOpen) = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    
End Sub


