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
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    'check other instance of app
    If App.PrevInstance = True Then
        MsgBox "VBPool2.0 draait al...."
        Exit Sub
    End If
    'set and open the database
    If Dir(App.Path & "\" & dbName & ".mdb") = "" Then
        createDb
    End If
    If Not cnOpen(cn) Then openDB
    If tableExists("tblPools") Then
        ' get last poolID
        If getPoolInfo("poolName") Then
            thisPool = Val(GetSetting(App.EXEName, "global", "lastpool", 0))
        End If
    End If
    If thisPool Then
        thisTournament = getThisPoolTournamentId()
    End If
    'open main form
    mainForm.Show
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
        strNumber = Val(Left(strNumber, Len(strNumber) - 1)) / 100
    End If
    
    If Not IsNumeric(strNumber) Then
        Exit Function
    Else
        float = Replace(strNumber, ",", ".")
    End If
End Function

Public Sub FillCombo(objComboBox As ComboBox, _
                     strSQL As String, _
                     strFieldToShow As String, _
                     Optional strFieldForItemData As String, _
                     Optional mySql As Boolean _
                     )
'Fills a combobox with values from a database
    
    Dim oRS As ADODB.Recordset  'Load the data
    Dim conn As ADODB.Connection
    If mySql Then
        Set conn = myConn
    Else
        Set conn = cn
    End If

    Set oRS = New ADODB.Recordset
    oRS.Open strSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
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
    conn.Close
    Set conn = Nothing
End Sub

Public Function DoLogin() As Boolean

'login system originally from Michael Ciurescu (CVMichael from vbforums.com)

    Dim UserName As String, Password As String, Ret As Boolean
    Dim LoginSuccessful As Boolean, rsData As ADODB.Recordset
    Dim MD5 As New clsMD5
    
    Randomize
    
    ' Get the user that last logged in from the registry
    UserName = getOrganisation("lastname")
        
    ' prompt user to enter username and password
    Ret = frmAdminLogin.GetLogIn(UserName, Password)
    
    Do While Ret
        Set rsData = cn.Execute("SELECT Passwd FROM tblOrganisation WHERE lastname = '" & Replace(UserName, "'", "''") & "'")
        
        ' if a record was found, it means the user exists
        If rsData.RecordCount > 0 Then
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
    
    DoLogin = LoginSuccessful
End Function

'add the nz function
Public Function nz(strValue As Variant, Optional alternative As String = "") As Variant
    If Not IsNull(strValue) Then
        nz = strValue
    Else
        nz = alternative
    End If
End Function


