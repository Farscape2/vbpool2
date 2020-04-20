VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
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
   Begin VB.CheckBox chkThrirdPlace 
      Height          =   255
      Left            =   1560
      TabIndex        =   20
      Top             =   2400
      Width           =   255
   End
   Begin VB.TextBox txtGroupCount 
      DataSource      =   "dtcTournaments"
      Height          =   360
      Left            =   4800
      TabIndex        =   18
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
      TabIndex        =   15
      Top             =   2925
      Width           =   5715
      Begin VB.CommandButton btnCancel 
         Cancel          =   -1  'True
         Caption         =   "Annuleren"
         Height          =   495
         Left            =   2910
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "Opslaan"
         Height          =   495
         Left            =   1575
         TabIndex        =   11
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton btnClose 
         Caption         =   "Sluiten"
         Default         =   -1  'True
         Height          =   495
         Left            =   4245
         TabIndex        =   13
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
      TabIndex        =   14
      Top             =   1860
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   635
      _Version        =   393216
      Value           =   24
      BuddyControl    =   "txtTeamAantal"
      BuddyDispid     =   196613
      OrigLeft        =   1920
      OrigTop         =   1800
      OrigRight       =   2175
      OrigBottom      =   2175
      Max             =   64
      Min             =   8
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSAdodcLib.Adodc dtcTournaments 
      Height          =   360
      Left            =   3600
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   635
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   14737632
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "MS Access Database"
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
   Begin MSDataListLib.DataCombo cmbLanden 
      DataSource      =   "dtcTournaments"
      Height          =   360
      Left            =   3240
      TabIndex        =   10
      Top             =   2340
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
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
      Format          =   51314689
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
      Format          =   51314689
      CurrentDate     =   43932
   End
   Begin MSComCtl2.UpDown UpDnGroupCount 
      Height          =   360
      Left            =   5220
      TabIndex        =   19
      Top             =   1860
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   635
      _Version        =   393216
      Value           =   24
      BuddyControl    =   "txtGroupCount"
      BuddyDispid     =   196629
      OrigLeft        =   1920
      OrigTop         =   1800
      OrigRight       =   2175
      OrigBottom      =   2175
      Max             =   64
      Min             =   8
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
      TabIndex        =   21
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
      TabIndex        =   17
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Toernooi gegevens"
      Height          =   375
      Left            =   240
      TabIndex        =   16
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

Private Sub dtcTournaments_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    With Me.dtcTournaments
        .Caption = " " & .Recordset.AbsolutePosition & "/" & .Recordset.RecordCount
    End With
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim ctl As Control
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

Dim sqlstr As String

'set Form defaults
    UnifyForm Me

'basis tabel
    With Me.dtcTournaments
        .ConnectionString = cn.ConnectionString
        .CommandType = adCmdText
        .RecordSource = "select * from tblTournaments where tournamentID = " & getThisPoolTournamentId()
    End With
    
'bindings
    With Me.cmbType
        .AddItem "CL"
        .AddItem "EK"
        .AddItem "WK"
        Set .DataSource = Me.dtcTournaments
        .DataField = "tournamentType"
    End With
    With Me.cmbYear
        For i = Year(Now) - 10 To Year(Now) + 10
            Me.cmbYear.AddItem i
        Next
        Set .DataSource = Me.dtcTournaments
        .DataField = "tournamentYear"
    End With
    With Me.dtpStart
        Set .DataSource = Me.dtcTournaments
        .DataField = "tournamentStartDate"
    End With
    With Me.dtpEind
        Set .DataSource = Me.dtcTournaments
        .DataField = "tournamentEndDate"
    End With
    With Me.txtTeamAantal
        Set .DataSource = Me.dtcTournaments
        .DataField = "tournamentTeamCount"
    End With
    With Me.txtGroupCount
        Set .DataSource = Me.dtcTournaments
        .DataField = "tournamentGroupCount"
    End With
    With Me.chkThrirdPlace
        Set .DataSource = Me.dtcTournaments
        .DataField = "tournamentThirdPlace"
    End With
    sqlstr = "Select * from tblCountries order by countryName"
    rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
    With Me.cmbLanden
        Set .RowSource = rs
        Set .DataSource = Me.dtcTournaments
        .BoundColumn = "countryId"
        .ListField = "countryName"
        .DataField = "tournamentLocationID"
    End With
    
    Me.dtcTournaments.Recordset.MoveLast
    
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

