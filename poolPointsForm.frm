VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form poolPointsForm 
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
   Begin VB.CommandButton btnClose 
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   8520
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc dtcPointList 
      Height          =   375
      Left            =   3240
      Top             =   6840
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
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
   Begin MSDataGridLib.DataGrid grdPoolPunten 
      Bindings        =   "poolPointsForm.frx":0000
      Height          =   7335
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   12938
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "pointTypeId"
         Caption         =   "pointTypeId"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "poolID"
         Caption         =   "poolID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Omschrijving"
         Caption         =   "Omschrijving"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Punten"
         Caption         =   "Punten"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "marge"
         Caption         =   "marge"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column02 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   3495,118
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   734,74
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   734,74
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc dtcPoolPoint 
      Height          =   330
      Left            =   480
      Top             =   8520
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   582
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
   Begin MSDataListLib.DataCombo cmbPointTypes 
      Bindings        =   "poolPointsForm.frx":001B
      Height          =   360
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   "omschrijving"
      BoundColumn     =   "id"
      Text            =   ""
   End
   Begin VB.Label Label2 
      Caption         =   "Zoek voorspelling"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Punten en marges"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Tag             =   "kop"
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "poolPointsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
    cn.BeginTrans
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
        cn.CommitTrans
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
    sqlstr = sqlstr & ", " & Val(Me.cmbPointTypes.BoundText)
    sqlstr = sqlstr & ", 0, 0)"
    qry.CommandType = adCmdText
    qry.CommandText = sqlstr
    qry.ActiveConnection = cn
    cn.BeginTrans
    qry.Execute rows
'    MsgBox rows & "  record toegevoegd", vbOKOnly + vbInformation, "Voorspelling opgeslagen"
    cn.CommitTrans
    Me.dtcPoolPoint.Recordset.Requery
    Me.dtcPoolPoint.Refresh
    Me.grdPoolPunten.Refresh
    Me.dtcPoolPoint.Recordset.Find "id = " & Val(Me.cmbPointTypes.BoundText)
    
    Set qry = Nothing
End Sub

Private Sub cmbPointTypes_Click(Area As Integer)
    
 If Area <> 0 Then
   
    With Me.dtcPoolPoint.Recordset
        cmbGridSelect = True
        .MoveFirst
         cmbGridSelect = False
        .Find "id = " & Val(Me.cmbPointTypes.BoundText)
        If .EOF Then
            'add new record
            insertRecord
        End If
    End With
 End If
End Sub

Private Sub Form_Load()
    Dim sqlstr As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    sqlstr = "Select pointtypeId as id, pointTypeDescription as omschrijving from tblPointtypes "
    If Not getTournamentInfo("tournamentThirdPlace") Then
        'exclude pointtype catgory "Kleine Finale"
        sqlstr = sqlstr & " where pointTypeCategory <> 6"
    End If
    sqlstr = sqlstr & " order by pointtypecategory, pointtypelistorder"
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    With Me.dtcPointList
        .ConnectionString = cn.ConnectionString
        .RecordSource = sqlstr
        .Refresh
    End With
    Me.dtcPoolPoint.ConnectionString = cn.ConnectionString
    
    sqlstr = "Select a.pointTypeID as id, a.poolID as poolId, pointTypeDescription as Omschrijving,"
    sqlstr = sqlstr & "pointPointsAward as Punten,"
    sqlstr = sqlstr & "pointPointsMargin as marge "
    sqlstr = sqlstr & "from tblPoolPoints a inner join tblPointTypes b "
    sqlstr = sqlstr & "on a.pointtypeid = b.pointtypeid"
    sqlstr = sqlstr & " where a.poolID = " & thisPool
    sqlstr = sqlstr & " order by b.pointtypecategory, b.pointtypelistorder"
    Me.dtcPoolPoint.RecordSource = sqlstr
    Me.dtcPoolPoint.Refresh
    Me.grdPoolPunten.Refresh
    UnifyForm Me

    If (rs.State And adStateOpen) = adStateOpen Then rs.Close
    Set rs = Nothing
End Sub

Private Sub grdPoolPunten_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    'check if record has changed
    If cmbGridSelect Then Exit Sub
    With Me.dtcPoolPoint.Recordset
        If Not .EOF Then
            Me.cmbPointTypes = !omschrijving
        End If
    End With
End Sub

