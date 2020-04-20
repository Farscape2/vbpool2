VERSION 5.00
Begin VB.Form basicDatasForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Table"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Annuleren"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Sluiten"
      Default         =   -1  'True
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Opslaan"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   3840
      Width           =   1335
   End
   Begin VB.ComboBox cmbSelect 
      Height          =   360
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   4215
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "Nieuw"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Selecteer"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "basicDatasForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'temporary array to save tournament ID's
Dim dataArray() As String
Dim selectedID As Long
Dim rs As New ADODB.Recordset

Dim dataChanged As Boolean
Dim newRecord As Boolean

Private Sub cmbSelect_Click()
Dim sqlstr As String
    selectedID = dataArray(Me.cmbSelect.ListIndex)
    'build sql string
    sqlstr = "Select * from tournaments where id = " & selectedID
    'open recordset
    rs.Open sqlstr, cn, adOpenDynamic, adLockOptimistic
    
    'do form stuff, fil the fields etc
    
    rs.Close
End Sub

Private Sub cmdClose_Click()
    If checkChanged() Then
        If MsgBox("Gegevens zijn veranderd, opslaan?", vbYesNo + vbQuestion) = vbYes Then
            cmdSave_Click
        End If
    End If
    Set rs = Nothing
    Unload Me
End Sub

Function checkChanged()
'check if any data has been changed
Dim sqlstr As String
Dim ctl As Control
Dim changed As Boolean
    changed = False
    If selectedID = 0 Then
        changed = True
    Else
        sqlstr = "select * from TABLENAME where ID = " & selectedID
        rs.Open sqlstr, cn, adOpenDynamic, adLockOptimistic
    End If
    For Each ctl In Me.Controls
        If ctl.Value <> rs(ctl.DataField).Value Then
            changed = True
            Exit For
        End If
    Next
    checkChanged = changed
End Function

Private Sub cmdNew_Click()
 Dim ctl As Control
 
    'empty all fields to add record to database
    For Each ctl In Me.Controls
        If Not TypeOf ctl Is CommandButton And Not TypeOf ctl Is Label Then
            On Error Resume Next
            ctl.Value = 0
            ctl.Text = ""
        End If
    Next
    newRecord = True
    dataChanged = True
    'setfocus to first control
    'Me.Controls("cmbTournamentType").SetFocus
End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub cmdSave_Click()
    'save the record to the database
    Dim sqlstr As String
    Dim dataOk As Boolean
    Dim errText As String
    On Error GoTo dataErr
    '''
    'data validation goes here
    '''
    'build insert command
    'sqlStr = "insert into tablename (fieldlist) VALUES ( values )"
    'execute the command
    cn.Execute sqlstr
    dataChanged = False
exitSub:
    Exit Sub
dataErr:
    MsgBox errText, "Fout bij opslaan gegevens", vbOKOnly
    Resume exitSub
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim sqlstr As String

Dim strSelect As String

dataChanged = False
'''''TO DO
        'fill default values
'select recordeset from the database
    'sqlStr = "select * from tournaments order by jaar, soort"

'fill select combo
    rs.Open sqlstr, cn, adOpenDynamic, adLockOptimistic
    If rs.RecordCount > 0 Then
        ReDim dataArray(rs.RecordCount - 1)
        rs.MoveFirst
        Do While Not rs.EOF
            With Me.cmbSelect
                'save the id
                dataArray(rs.AbsolutePosition - 1) = rs!id
                'add string to combobox
''''  TO DO
                'build the combobox string
                ' strSelect = rs!soort & "-" & rs!jaar
                .AddItem strSelect
                rs.MoveNext
            End With
        Loop
    End If
'close the recordset
    rs.Close
End Sub

