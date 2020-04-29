VERSION 5.00
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
   Begin VB.ComboBox cmbSelPool 
      Height          =   315
      Left            =   240
      TabIndex        =   9
      Top             =   480
      Width           =   3735
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
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Tot"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   2055
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Van"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblTournamentInfo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Toernooi:"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label lblTournamentInfo 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   2
      Left            =   840
      TabIndex        =   4
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

Dim rs As ADODB.Recordset

Private Sub CancelButton_Click()
    Unload Me
End Sub


Private Sub cmbSelPool_Click()
    Dim sqlstr As String
    Dim lstID As Long
    
    Set rs = New ADODB.Recordset
    
    lstID = Me.cmbSelPool.ItemData(Me.cmbSelPool.ListIndex)
    thisPool = lstID
    sqlstr = "Select * from tblTournaments WHERE tournamentID=" & getThisPoolTournamentId()
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
        With Me
            .lblTournamentInfo(0).Caption = "Toernooi: " & rs!tournamenttype & "-" & rs!tournamentYear
            .lblTournamentInfo(1).Caption = rs!tournamentStartDate
            .lblTournamentInfo(2).Caption = rs!tournamentEnddate
        End With
    End If
    

End Sub

Private Sub Form_Load()
    Dim sqlstr As String
    sqlstr = "Select * from tblPools"
    FillCombo Me.cmbSelPool, sqlstr, "poolName", "poolId"
'set Form defaults
    UnifyForm Me
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If (rs.State And adStateOpen) = adStateOpen Then rs.Close
    Set rs = Nothing

End Sub

Private Sub OKButton_Click()
    thisPool = Me.cmbSelPool.ItemData(Me.cmbSelPool.ListIndex)
    thisTournament = getThisPoolTournamentId
    SaveSetting App.EXEName, "global", "lastpool", thisPool
    Unload Me
End Sub
