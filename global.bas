Attribute VB_Name = "global"
Option Explicit

'currentPool is read and stored in dbFunctions module
Public thisPool As Long
Public thisTournament As Long
'variable to preserve the current active country
Public currentCountry As Long  'used to pass information between forms

Sub Main()
    'check other instance of app
    If App.PrevInstance = True Then
        MsgBox "VBPool2.0 draait al...."
        Exit Sub
    End If
    'set and open the database
    openDB
    ' get last poolID
    thisPool = Val(GetSetting(App.EXEName, "global", "lastpool"))
    If thisPool Then
        thisTournament = getThisPoolTournamentId()
    End If
    'open main form
    mainForm.Show
End Sub

'add the nz function
Public Function nz(strValue As Variant, Optional alternative As String = "") As Variant
    If Not IsNull(strValue) Then
        nz = strValue
    Else
        nz = alternative
    End If
End Function

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
                     Optional strFieldForItemData As String)
' There is ONE global connection
'Fills a combobox with values from a database

'Parameters:
  'objComboBox    = the ComboBox to fill
  'oConn          = the connection to the database
  'strSQL         = the SQL to load the data
  'strFieldToShow = the name of the field to show
  'strFieldForItemData (optional) = the name of the field to put into ItemData (Integer type fields only)

'Example usage (standard):
  'Call FillCombo(Combo1, oConn, "SELECT Colour FROM Colours", "Colour")

'Example usage (with ItemData):
  'Call FillCombo(Combo1, oConn, "SELECT Colour, ColourID FROM Colours", "Colour", "ColourID")


Dim oRS As ADODB.Recordset  'Load the data
  Set oRS = New ADODB.Recordset
  oRS.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
  If oRS.EOF Then
    MsgBox "Geen records in recordset", vbCritical + vbOKOnly, "FillCombo"
    Exit Sub
  End If
  With objComboBox          'Fill the combo box
    .Clear
    If strFieldForItemData = "" Then
      Do While Not oRS.EOF      '(without ItemData)
        .AddItem oRS.Fields(strFieldToShow).Value
        oRS.MoveNext
      Loop
    Else
      Do While Not oRS.EOF      '(with ItemData)
        .AddItem oRS.Fields(strFieldToShow).Value
        .ItemData(.NewIndex) = oRS.Fields(strFieldForItemData).Value
        oRS.MoveNext
      Loop
    End If
  End With

  oRS.Close                 'Tidy up
  Set oRS = Nothing

End Sub

