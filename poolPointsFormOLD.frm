VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPoolPoints 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Punten toekenning"
   ClientHeight    =   9030
   ClientLeft      =   12975
   ClientTop       =   4365
   ClientWidth     =   5805
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMarge 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   600
      Width           =   420
   End
   Begin VB.TextBox txtPnt 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   600
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid grdPoolpoints 
      Height          =   7455
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   13150
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin VB.CommandButton btnCopyDefaultPoints 
      Caption         =   "Beginwaarden"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   8520
      Width           =   1815
   End
   Begin VB.ComboBox cmbPointTypes 
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   8520
      Width           =   1455
   End
   Begin MSComCtl2.UpDown upDnPnt 
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   600
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtPnt"
      BuddyDispid     =   196615
      OrigLeft        =   840
      OrigTop         =   480
      OrigRight       =   1095
      OrigBottom      =   855
      Max             =   150
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown uipDnMarge 
      Height          =   375
      Left            =   5340
      TabIndex        =   8
      Top             =   600
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtMarge"
      BuddyDispid     =   196617
      OrigLeft        =   840
      OrigTop         =   480
      OrigRight       =   1095
      OrigBottom      =   855
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Label Label3 
      Caption         =   "mrg"
      Height          =   255
      Left            =   4440
      TabIndex        =   10
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "pnt"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Punten en marges"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Tag             =   "kop"
      Top             =   120
      Width           =   5670
   End
End
Attribute VB_Name = "frmPoolPoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection

Dim cmbGridSelect As Boolean

'if true then don't update combo box after datagrid.refresh
'to prevent recordsource jump back to first record

Private Sub btnClose_Click()
'check if any row has 0 point and delete this record
    Dim msg As String
    Dim qry As ADODB.Command
    Dim sqlstr As String
    Dim rows As Long
    Set qry = New ADODB.Command
    sqlstr = "Delete from tblPoolPoints where poolid = " & thisPool
    sqlstr = sqlstr & " AND pointPointsAward = 0 "
    qry.ActiveConnection = cn
    qry.CommandText = sqlstr
    qry.CommandType = adCmdText
    qry.Execute rows
    If rows Then
        If rows = 1 Then
            msg = "1 rij zonder punten is verwijderd"
        Else
            msg = rows & " rijen zonder punten, zijn verwijderd"
        End If
        MsgBox msg, vbOKOnly + vbInformation, "Pool instellingen"
        'cn.CommitTrans
    End If
    Unload Me
    Set qry = Nothing
End Sub

Sub insertRecord()
Dim qry As ADODB.Command

Dim sqlstr As String
Dim rows As Long

    Set qry = New ADODB.Command
    sqlstr = "insert into tblPoolPoints (poolID, pointTypeId, pointPointsAward, pointPointsMargin) "
    sqlstr = sqlstr & "VALUES ( " & thisPool
    sqlstr = sqlstr & ", " & Me.cmbPointTypes.ItemData(Me.cmbPointTypes.ListIndex)
    sqlstr = sqlstr & ", 0, 0)"
    qry.CommandType = adCmdText
    qry.CommandText = sqlstr
    qry.ActiveConnection = cn
    cn.BeginTrans
    qry.Execute rows
'    MsgBox rows & "  record toegevoegd", vbOKOnly + vbInformation, "Voorspelling opgeslagen"
    cn.CommitTrans
    Set qry = Nothing
End Sub

Private Sub btnCopyDefaultPoints_Click()
  'copy default points table
  copyDefaultPoints
  fillPointsGrid
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    Dim sqlstr As String
    
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = lclConn()
        .Open
    End With
    
    sqlstr = "Select pointtypeId as id, pointTypeDescription as omschrijving from tblPointtypes "
    If Not getTournamentInfo("tournamentThirdPlace", cn) Then
        'exclude pointtype catgory "Kleine Finale"
        sqlstr = sqlstr & " where pointTypeCategory <> 6"
    End If
    
    sqlstr = sqlstr & " order by pointtypecategory, pointtypelistorder"
    
    FillCombo Me.cmbPointTypes, sqlstr, cn, "omschrijving", "id"
'    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
'    With Me.cmbPointTypes
'        Set .RowSource = rs
'        .BoundColumn = "id"
'        .ListField = "omschrijving"
'    End With
    fillPointsGrid
    UnifyForm Me
    centerForm Me
End Sub

Sub fillPointsGrid()
Dim sqlstr As String
Dim i As Integer, j As Integer
  Set rs = New ADODB.Recordset
  sqlstr = "Select a.pointTypeID as id, a.poolID as poolId, pointTypeDescription as Omschrijving,"
  sqlstr = sqlstr & "pointPointsAward as Punten,"
  sqlstr = sqlstr & "pointPointsMargin as Marge "
  sqlstr = sqlstr & "from tblPoolPoints a inner join tblPointTypes b "
  sqlstr = sqlstr & "on a.pointtypeid = b.pointtypeid"
  sqlstr = sqlstr & " where a.poolID = " & thisPool
  sqlstr = sqlstr & " order by b.pointtypecategory, b.pointtypelistorder"
  
  With rs
      .CursorLocation = adUseClient
      .Open sqlstr, cn, adOpenKeyset, adLockOptimistic
  End With
  With Me.grdPoolpoints
    .Clear
    .cols = rs.Fields.Count
    .rows = rs.RecordCount + 1
    i = 0
    For j = 0 To rs.Fields.Count - 1
      .TextMatrix(i, j) = rs.Fields(j).Name
    Next
    Do While Not rs.EOF
      i = i + 1
      For j = 0 To rs.Fields.Count - 1
        .TextMatrix(i, j) = rs.Fields(j).value
      Next
      rs.MoveNext
    Loop
    .ColWidth(0) = 0
    .ColWidth(1) = 0
    .ColWidth(2) = 3200
    .ColAlignment(2) = flexAlignLeftCenter
    .ColWidth(3) = 800
    .ColAlignment(3) = flexAlignCenterCenter
    .ColWidth(4) = 800
    .ColAlignment(4) = flexAlignCenterCenter
  End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If (rs.State And adStateOpen) = adStateOpen Then rs.Close
    Set rs = Nothing
    If (cn.State And adStateOpen) = adStateOpen Then cn.Close
    Set cn = Nothing
End Sub

Private Sub grdPoolPunten_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If cmbGridSelect Then Exit Sub
    Me.cmbPointTypes = grdPoolPunten.Columns(2)
 End Sub

Private Sub grdPoolpoints_RowColChange()
Dim id As Integer, i As Integer

  id = Me.grdPoolpoints.TextMatrix(Me.grdPoolpoints.row, 0)
  For i = 0 To Me.cmbPointTypes.ListCount - 1
    If Me.cmbPointTypes.ItemData(i) = id Then
      Exit For
    End If
  Next
  If i < Me.cmbPointTypes.ListCount Then
    Me.cmbPointTypes = Me.cmbPointTypes.List(i)
    Me.upDnPnt = Me.grdPoolpoints.TextMatrix(Me.grdPoolpoints.row, 3)
    Me.uipDnMarge = val(Me.grdPoolpoints.TextMatrix(Me.grdPoolpoints.row, 4))
  End If
End Sub
