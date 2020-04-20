VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form openPool 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open Pool"
   ClientHeight    =   2610
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4155
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataCombo cmbPools 
      Bindings        =   "openPool.frx":0000
      DataSource      =   "dtcPools"
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "poolName"
      BoundColumn     =   "poolID"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc dtcPools 
      Height          =   375
      Left            =   2400
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Pools"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   1935
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Open"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblPoolID 
      Caption         =   "Label3"
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Tot"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   2055
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Van"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblTournamentInfo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Toernooi:"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label lblTournamentInfo 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   2
      Left            =   840
      TabIndex        =   5
      Top             =   1935
      Width           =   1695
   End
   Begin VB.Label lblTournamentInfo 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   3
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Selecteer een pool"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Tag             =   "kop"
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "openPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub cmbPools_Click(Area As Integer)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim tournamentID As Long
    
    If Area <> 0 Then
        rs.Open "Select * from tblPools WHERE poolID = " & Val(Me.cmbPools.BoundText), cn, adOpenKeyset, adLockReadOnly
        If Not rs.EOF Then
            Me.lblPoolID.Caption = rs!poolid
            tournamentID = rs!tournamentID
            rs.Close
            rs.Open "Select * from tblTournaments WHERE tournamentID = " & tournamentID, cn, adOpenKeyset, adLockReadOnly
            If Not rs.EOF Then
                With Me
                    .lblTournamentInfo(0).Caption = "Toernooi: " & rs!tournamenttype & "-" & rs!tournamentYear
                    .lblTournamentInfo(1).Caption = rs!tournamentStartDate
                    .lblTournamentInfo(2).Caption = rs!tournamentEnddate
                End With
            End If
            rs.Close
        End If
    End If

    If (rs.State And adStateOpen) = adStateOpen Then rs.Close
    Set rs = Nothing

End Sub

Private Sub Form_Load()
    Me.dtcPools.ConnectionString = cn.ConnectionString
    Me.dtcPools.RecordSource = "select * from tblPools"
    Me.dtcPools.Refresh
'set Form defaults
    UnifyForm Me
    
End Sub

Private Sub OKButton_Click()
    thisPool = Val(Me.lblPoolID.Caption)
    thisTournament = getThisPoolTournamentId
    SaveSetting App.EXEName, "global", "lastpool", thisPool
    Unload Me
End Sub
