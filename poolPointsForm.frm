VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
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
   Begin VB.CommandButton btnCopyDefaultPoints 
      Caption         =   "Beginwaarden"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   8520
      Width           =   1815
   End
   Begin VB.ComboBox cmbPointTypes 
      Height          =   360
      Left            =   1800
      TabIndex        =   4
      Top             =   600
      Width           =   3735
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   8520
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid grdPoolPunten 
      Height          =   7335
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   12938
      _Version        =   393216
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      WrapCellPointer =   -1  'True
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
      SplitCount      =   2
      BeginProperty Split0 
         MarqueeStyle    =   5
         ScrollBars      =   0
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Size            =   2
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
            Object.Visible         =   0   'False
            ColumnWidth     =   599,811
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
            Object.Visible         =   0   'False
            ColumnWidth     =   599,811
         EndProperty
         BeginProperty Column02 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   3600
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnAllowSizing=   0   'False
            Object.Visible         =   0   'False
            ColumnWidth     =   734,74
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            Object.Visible         =   0   'False
            ColumnWidth     =   734,74
         EndProperty
      EndProperty
      BeginProperty Split1 
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         RecordSelectors =   0   'False
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
            Object.Visible         =   0   'False
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
      Left            =   2880
      Top             =   8520
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin VB.Label Label2 
      Caption         =   "Zoek voorspelling"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
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
    Me.dtcPoolPoint.Recordset.Requery
    Me.dtcPoolPoint.Refresh
    Me.grdPoolPunten.Refresh
    Me.dtcPoolPoint.Recordset.Find "id = " & Me.cmbPointTypes.ItemData(Me.cmbPointTypes.ListIndex)
    
    Set qry = Nothing
End Sub

Private Sub btnCopyDefaultPoints_Click()
  'copy default points table
  copyDefaultPoints
  dtcPoolPoint.Refresh
  Me.grdPoolPunten.Refresh
End Sub

Private Sub cmbPointTypes_Click()
    With Me.dtcPoolPoint.Recordset
        cmbGridSelect = True 'prevent executing rowcolchange
        .MoveFirst 'to start find at first row
        .Find "id = " & val(Me.cmbPointTypes.ItemData(Me.cmbPointTypes.ListIndex))
        If .EOF Then 'not found
            'add new record
            insertRecord
            cmbGridSelect = False
        End If
    End With
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
    
    Set rs = New ADODB.Recordset
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
    
    Me.dtcPoolPoint.ConnectionString = cn.ConnectionString
    Me.dtcPoolPoint.CursorLocation = adUseClient
    sqlstr = "Select a.pointTypeID as id, a.poolID as poolId, pointTypeDescription as Omschrijving,"
    sqlstr = sqlstr & "pointPointsAward as Punten,"
    sqlstr = sqlstr & "pointPointsMargin as marge "
    sqlstr = sqlstr & "from tblPoolPoints a inner join tblPointTypes b "
    sqlstr = sqlstr & "on a.pointtypeid = b.pointtypeid"
    sqlstr = sqlstr & " where a.poolID = " & thisPool
    sqlstr = sqlstr & " order by b.pointtypecategory, b.pointtypelistorder"
    
    With rs
        .CursorLocation = adUseClient
        .Open sqlstr, cn, adOpenKeyset, adLockOptimistic
    End With
'    With Me.grdPoolPunten
'        Set .DataSource = rs
'    End With

'decided to use the adodc control anyway, much easier!
    Me.dtcPoolPoint.RecordSource = sqlstr
    Set Me.grdPoolPunten.DataSource = Me.dtcPoolPoint
    
    UnifyForm Me
    centerForm Me
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

