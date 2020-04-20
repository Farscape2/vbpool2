VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form poolsForm 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pools"
   ClientHeight    =   5355
   ClientLeft      =   12630
   ClientTop       =   6360
   ClientWidth     =   5790
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
   ScaleHeight     =   5355
   ScaleWidth      =   5790
   Begin VB.Frame frmPrizes 
      Caption         =   "Prijzen"
      Height          =   2295
      Left            =   0
      TabIndex        =   10
      Top             =   2280
      Width           =   5775
      Begin MSMask.MaskEdBox txtHighestDayscore 
         DataField       =   "prizeMostDayPoints"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   2
         EndProperty
         DataSource      =   "dtcPools"
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Top             =   660
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         Format          =   "€ #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHighestPosition 
         DataField       =   "prizeBestDayPosition"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   2
         EndProperty
         DataSource      =   "dtcPools"
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Top             =   1132
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         Format          =   "€ #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtLowestPosition 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   2
         EndProperty
         DataSource      =   "dtcPools"
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Top             =   1650
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         Format          =   "€ #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPercentage 
         DataField       =   "prizePercentageFirst"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   5
         EndProperty
         DataSource      =   "dtcPools"
         Height          =   375
         Index           =   0
         Left            =   3240
         TabIndex        =   14
         Top             =   660
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         _Version        =   393216
         Format          =   "0%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPrizeLastOverall 
         DataField       =   "prizeLastOverallPosition"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   2
         EndProperty
         DataSource      =   "dtcPools"
         Height          =   375
         Left            =   3600
         TabIndex        =   18
         Top             =   1650
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         Format          =   "€ #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPercentage 
         DataField       =   "prizePercentageSecond"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   5
         EndProperty
         DataSource      =   "dtcPools"
         Height          =   375
         Index           =   1
         Left            =   4680
         TabIndex        =   15
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         _Version        =   393216
         Format          =   "0%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPercentage 
         DataField       =   "prizePercentageThird"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   5
         EndProperty
         DataSource      =   "dtcPools"
         Height          =   375
         Index           =   2
         Left            =   3240
         TabIndex        =   16
         Top             =   1132
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         _Version        =   393216
         Format          =   "0%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPercentage 
         DataField       =   "prizePercentageFourth"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   5
         EndProperty
         DataSource      =   "dtcPools"
         Height          =   375
         Index           =   3
         Left            =   4680
         TabIndex        =   17
         Top             =   1132
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         _Version        =   393216
         Format          =   "0%"
         PromptChar      =   "_"
      End
      Begin VB.Label lblTotal 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   615
         Left            =   4680
         TabIndex        =   34
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Laatste"
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   2880
         TabIndex        =   33
         Top             =   1710
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Onderaan"
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1710
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bovenaan"
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1192
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Meeste punten"
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   0
         TabIndex        =   30
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "2e"
         Height          =   375
         Left            =   4320
         TabIndex        =   29
         Top             =   660
         Width           =   375
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "1e"
         Height          =   375
         Left            =   2880
         TabIndex        =   28
         Top             =   660
         Width           =   375
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "4e"
         Height          =   255
         Left            =   4320
         TabIndex        =   27
         Top             =   1192
         Width           =   375
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "3e"
         Height          =   255
         Left            =   2880
         TabIndex        =   26
         Top             =   1192
         Width           =   375
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Eindstand"
         Height          =   255
         Left            =   3840
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.Line Line2 
         X1              =   2640
         X2              =   2640
         Y1              =   360
         Y2              =   2040
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Dagprijzen"
         Height          =   255
         Left            =   720
         TabIndex        =   24
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSMask.MaskEdBox txtCosts 
      DataSource      =   "dtcPools"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   1260
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      Format          =   "€ #,##0.00"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtPoolName 
      DataSource      =   "dtcPools"
      Height          =   360
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   4455
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   5730
      TabIndex        =   23
      Top             =   4620
      Width           =   5790
      Begin VB.CommandButton btnCancel 
         Cancel          =   -1  'True
         Caption         =   "Annuleren"
         Height          =   495
         Left            =   2880
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton btnDelete 
         Caption         =   "Wissen"
         Height          =   495
         Left            =   2880
         TabIndex        =   21
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "Opslaan"
         Height          =   495
         Left            =   1575
         TabIndex        =   19
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton btnClose 
         Caption         =   "Sluiten"
         Default         =   -1  'True
         Height          =   495
         Left            =   4245
         TabIndex        =   22
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc dtcPools 
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
      CommandType     =   1
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
   Begin MSDataListLib.DataCombo cmbTournaments 
      DataSource      =   "dtcPools"
      Height          =   360
      Left            =   1080
      TabIndex        =   3
      Top             =   1267
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpStart 
      DataSource      =   "dtcPools"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   1740
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   146800641
      CurrentDate     =   43932
   End
   Begin MSComCtl2.DTPicker dtpEind 
      DataSource      =   "dtcPools"
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   1740
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   146800641
      CurrentDate     =   43932
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pool gegevens aanpassen"
      Height          =   375
      Left            =   240
      TabIndex        =   35
      Tag             =   "kop"
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Inleg "
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pool naam"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   -120
      TabIndex        =   0
      Top             =   780
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5640
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "tot"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Inleveren"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Toernooi"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
End
Attribute VB_Name = "poolsForm"
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

If chkPoolHasCompetitors(Me.dtcPools.Recordset!poolid) Then
    MsgBox "Pool heeft deelnemers", vbOKOnly + vbCritical, "Kan niet verwijderen"
    Exit Sub
End If
msgStr = "Deze pool werkelijk verwijderen? " & vbNewLine & "(kan alleen als er geen deelnemers voor zijn)"
confirmation = MsgBox(msgStr, vbQuestion + vbYesNo, "Toernooi wissen")
With Me.dtcPools.Recordset
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
        Me.dtcPools.Recordset.Update
        cn.CommitTrans
        setState False
        DoEvents
    Else
        setState True
        If (supportsTransactions(cn)) Then
            cn.BeginTrans
        End If
    End If
    
End Sub

Private Sub dtcPools_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    With Me.dtcPools
        .Caption = " " & .Recordset.AbsolutePosition & "/" & .Recordset.RecordCount
    End With
End Sub

Private Sub Form_Load()
Dim ctl As Control
Dim sqlstr As String
Dim i As Integer
Dim rsTournaments As ADODB.Recordset
    Set rsTournaments = New ADODB.Recordset
'set Form defaults
    UnifyForm Me

'back color of frame
    Me.frmPrizes.BackColor = Me.BackColor
'basis tabel
    With Me.dtcPools
        .ConnectionString = cn.ConnectionString
        .CommandType = adCmdText
        .RecordSource = "select * from tblPools where poolid=" & thisPool
    End With

'bindings
    Me.txtPoolName.DataField = "poolName"
    sqlstr = "Select tournamentID, tournamentType & ' - ' & tournamentYear as tournament from tblTournaments order by tournamentYear"
    rsTournaments.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
    With Me.cmbTournaments
        Set .RowSource = rsTournaments
        Set .DataSource = Me.dtcPools
        .BoundColumn = "tournamentId"
        .ListField = "tournament"
        .DataField = "tournamentId"
    End With
    Me.txtCosts.DataField = "poolCost"
    Me.dtpStart.DataField = "poolStartAcceptForms"
    Me.dtpEind.DataField = "poolEndAcceptForms"
    
    'prizes
    Me.txtHighestDayscore.DataField = "prizeHighDayScore"
    Me.txtHighestPosition.DataField = "prizeHighDayOverallPosition"
    Me.txtLowestPosition.DataField = "prizeLowDayOverallPosition"
    
    Me.txtPercentage(0).DataField = "prizePercentageFirst"
    Me.txtPercentage(1).DataField = "prizePercentageSecond"
    Me.txtPercentage(2).DataField = "prizePercentageThird"
    Me.txtPercentage(3).DataField = "prizePercentageFourth"
    Me.txtPrizeLastOverall.DataField = "prizeLowFinalOverallPosition"
    
    
    Me.btnSave.Enabled = Not chkTournamentStarted()
    Me.btnDelete.Enabled = Not chkTournamentStarted()
    'set form state
    setState False
    If (rsTournaments.State And adStateOpen) = adStateOpen Then rsTournaments.Close
    Set rsTournaments = Nothing

End Sub

Sub setState(edit As Boolean)
Dim ctl As Control
    editState = edit
    With Me
        For Each ctl In .Controls
            If TypeOf ctl Is DTPicker Or _
                TypeOf ctl Is TextBox Or _
                TypeOf ctl Is DataCombo Or _
                TypeOf ctl Is ComboBox Or _
                TypeOf ctl Is MaskEdBox Or _
                TypeOf ctl Is UpDown Then
                ctl.Enabled = edit
            End If
        Next
        .btnDelete.Visible = Not edit
        .btnCancel.Visible = edit
        If edit Then
            .btnSave.Caption = "Opslaan"
        Else
            .btnSave.Caption = "Bewerken"
        End If
        .btnClose.Enabled = Not edit
    End With
End Sub

Sub calcTotalPercentage()
'calculate the total of the percentage prizes
    Dim totalPerc As Double
    Dim i As Integer
    totalPerc = 0
    For i = 0 To 3
        totalPerc = totalPerc + Val(float(Me.txtPercentage(i).Text))
    Next
    Me.lblTotal.Caption = Format(totalPerc, "0.0%")
    If totalPerc <> 1 Then
        Me.lblTotal.ForeColor = vbRed
        Me.lblTotal.Caption = "LET OP: " & Format(totalPerc, "0.0%")
    Else
        Me.lblTotal.ForeColor = Me.Label15.ForeColor
    End If
End Sub

Private Sub txtPercentage_LostFocus(Index As Integer)
    calcTotalPercentage
End Sub
