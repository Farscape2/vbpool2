VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form tournamentsForm 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Torrnooien"
   ClientHeight    =   3660
   ClientLeft      =   12630
   ClientTop       =   6360
   ClientWidth     =   5715
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
   ScaleHeight     =   3660
   ScaleWidth      =   5715
   Begin VB.ComboBox cmbLanden 
      Height          =   360
      Left            =   3120
      TabIndex        =   21
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CheckBox chkThrirdPlace 
      Height          =   255
      Left            =   1560
      TabIndex        =   19
      Top             =   2400
      Width           =   255
   End
   Begin VB.TextBox txtGroupCount 
      DataSource      =   "dtcTournaments"
      Height          =   360
      Left            =   4800
      TabIndex        =   17
      Top             =   1860
      Width           =   420
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   5655
      TabIndex        =   14
      Top             =   2925
      Width           =   5715
      Begin VB.CommandButton btnCancel 
         Cancel          =   -1  'True
         Caption         =   "Annuleren"
         Height          =   495
         Left            =   2910
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "Opslaan"
         Height          =   495
         Left            =   1575
         TabIndex        =   10
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton btnClose 
         Caption         =   "Sluiten"
         Default         =   -1  'True
         Height          =   495
         Left            =   4245
         TabIndex        =   12
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.TextBox txtTeamAantal 
      DataSource      =   "dtcTournaments"
      Height          =   360
      Left            =   1440
      TabIndex        =   7
      Top             =   1860
      Width           =   435
   End
   Begin MSComCtl2.UpDown upDwnTeamAantal 
      Height          =   360
      Left            =   1860
      TabIndex        =   13
      Top             =   1860
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   635
      _Version        =   393216
      Value           =   8
      BuddyControl    =   "txtTeamAantal"
      BuddyDispid     =   196615
      OrigLeft        =   1920
      OrigTop         =   1800
      OrigRight       =   2175
      OrigBottom      =   2175
      Max             =   48
      Min             =   8
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.DTPicker dtpStart 
      DataSource      =   "dtcTournaments"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   1260
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   147456001
      CurrentDate     =   43932
   End
   Begin VB.ComboBox cmbYear 
      DataSource      =   "dtcTournaments"
      Height          =   360
      Left            =   4200
      TabIndex        =   3
      Top             =   780
      Width           =   1335
   End
   Begin VB.ComboBox cmbType 
      DataSource      =   "dtcTournaments"
      Height          =   360
      Left            =   1440
      TabIndex        =   1
      Top             =   780
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpEind 
      DataSource      =   "dtcTournaments"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   1260
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   147456001
      CurrentDate     =   43932
   End
   Begin MSComCtl2.UpDown UpDnGroupCount 
      Height          =   360
      Left            =   5220
      TabIndex        =   18
      Top             =   1860
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   635
      _Version        =   393216
      Value           =   12
      BuddyControl    =   "txtGroupCount"
      BuddyDispid     =   196610
      OrigLeft        =   1920
      OrigTop         =   1800
      OrigRight       =   2175
      OrigBottom      =   2175
      Max             =   12
      Min             =   2
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Derde plaats"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Aantal groepen"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Toernooi gegevens"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Tag             =   "kop"
      Top             =   120
      Width           =   5295
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5640
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Aantal teams"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Locatie"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Van / Tot"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jaar"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "tournamentsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim editState As Boolean

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
Dim msgStr As String
Dim confirmation  As Integer
Dim pos As Long
If chkTournamentHasPools(Me.dtcTournaments.Recordset!tournamentID) Then
    MsgBox "Er zijn pools voor dit toernooi", vbOKOnly + vbCritical, "Kan niet verwijderen"
    Exit Sub
End If
msgStr = "Dit toernooi werkelijk verwijderen? " & vbNewLine & "(kan alleen als er geen pools voor zijn)"
confirmation = MsgBox(msgStr, vbQuestion + vbYesNo, "Toernooi wissen")
With Me.dtcTournaments.Recordset
    If confirmation = vbYes Then
        pos = .AbsolutePosition
        .Delete
        If pos = 1 Then
            .MoveNext
        Else
            .MovePrevious
        End If
    End If
End With
End Sub

Private Sub btnCancel_Click()
    cn.RollbackTrans
    setState False
End Sub

Private Sub btnSave_Click()

    If editState Then
        cn.CommitTrans
        setState False
        'check / generate the tournament schedule
        generateSchedule
        
    Else
        setState True
        cn.BeginTrans
    End If
    
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim ctl As Control
Dim rs As ADODB.Recordset
Dim adoCmd As ADODB.Command

Dim sqlstr As String

Set adoCmd = New ADODB.Command

'set Form defaults
    UnifyForm Me
    sqlstr = "Select * from tblTournaments WHERE tournamentID = ?"
    With adoCmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = sqlstr
        .Parameters.Append .CreateParameter("id", adInteger, adParamInputOutput)
        .Parameters("id").Value = thisTournament
        Set rs = .Execute
    End With
    
'bindings
    With Me.cmbType
        .AddItem "CL"
        .AddItem "EK"
        .AddItem "WK"
    End With
    With Me.cmbYear
        For i = Year(Now) - 10 To Year(Now) + 10
            Me.cmbYear.AddItem i
        Next
    End With
    sqlstr = "Select * from tblCountries order by countryName"
    FillCombo Me.cmbLanden, sqlstr, "countryname", "countryid"
    
    Me.cmbType = rs!tournamenttype
    Me.cmbYear = rs!tournamentYear
    Me.dtpStart = CDbl(rs!tournamentstartdate)
    Me.dtpEind = CDbl(rs!tournamentEnddate)
    Me.UpDnGroupCount.Value = rs!tournamentGroupCount
    Me.upDwnTeamAantal = rs!tournamentTeamCount
    For i = 0 To Me.cmbLanden.ListCount - 1
        If Me.cmbLanden.ItemData(i) = rs!tournamentLocationID Then
            Me.cmbLanden.ListIndex = i
            Exit For
        End If
    Next
        
'    With Me.dtpStart
'        Set .DataSource = Me.dtcTournaments
'        .DataField = "tournamentStartDate"
'    End With
'    With Me.dtpEind
'        Set .DataSource = Me.dtcTournaments
'        .DataField = "tournamentEndDate"
'    End With
'    With Me.txtTeamAantal
'        Set .DataSource = Me.dtcTournaments
'        .DataField = "tournamentTeamCount"
'    End With
'    With Me.txtGroupCount
'        Set .DataSource = Me.dtcTournaments
'        .DataField = "tournamentGroupCount"
'    End With
'    With Me.chkThrirdPlace
'        Set .DataSource = Me.dtcTournaments
'        .DataField = "tournamentThirdPlace"
'    End With
    
    Me.btnSave.Enabled = Not chkTournamentStarted()

    'set form state
    setState False
    If (rs.State And adStateOpen) = adStateOpen Then rs.Close
    Set rs = Nothing
End Sub

Sub setState(edit As Boolean)
Dim ctl As Control
    editState = edit
    With Me
        For Each ctl In .Controls
            If TypeOf ctl Is TextBox Or _
                TypeOf ctl Is DataCombo Or _
                TypeOf ctl Is ComboBox Then
                ctl.Locked = Not edit
            End If
            If TypeOf ctl Is DTPicker Or _
                TypeOf ctl Is UpDown Then
                ctl.Enabled = edit
            End If
        Next
        .btnCancel.Visible = edit
        If edit Then
            .btnSave.Caption = "Opslaan"
        Else
            .btnSave.Caption = "Bewerken"
        End If
        Me.btnClose.Enabled = Not edit
    End With
End Sub


