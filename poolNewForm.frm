VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form newPoolForm 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pools"
   ClientHeight    =   5190
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
   ScaleHeight     =   5190
   ScaleWidth      =   5790
   Begin VB.Frame frmPrizes 
      Caption         =   "Prijzen"
      Height          =   2295
      Left            =   0
      TabIndex        =   10
      Top             =   2160
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
         Height          =   375
         Left            =   4680
         TabIndex        =   32
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Laatste"
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   2880
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "2e"
         Height          =   375
         Left            =   4320
         TabIndex        =   27
         Top             =   660
         Width           =   375
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "1e"
         Height          =   375
         Left            =   2880
         TabIndex        =   26
         Top             =   660
         Width           =   375
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "4e"
         Height          =   255
         Left            =   4320
         TabIndex        =   25
         Top             =   1192
         Width           =   375
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "3e"
         Height          =   255
         Left            =   2880
         TabIndex        =   24
         Top             =   1192
         Width           =   375
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Eindstand"
         Height          =   255
         Left            =   3840
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSMask.MaskEdBox txtCosts 
      DataSource      =   "dtcPools"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   1140
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
      Top             =   600
      Width           =   4455
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   5730
      TabIndex        =   21
      Top             =   4455
      Width           =   5790
      Begin VB.CommandButton btnCancel 
         Cancel          =   -1  'True
         Caption         =   "Annuleren"
         Height          =   495
         Left            =   3000
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton btnClose 
         Caption         =   "Opslaan"
         Default         =   -1  'True
         Height          =   495
         Left            =   4320
         TabIndex        =   20
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSDataListLib.DataCombo cmbTournaments 
      DataSource      =   "dtcPools"
      Height          =   360
      Left            =   1080
      TabIndex        =   3
      Top             =   1140
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
      Top             =   1620
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   149094401
      CurrentDate     =   43932
   End
   Begin MSComCtl2.DTPicker dtpEind 
      DataSource      =   "dtcPools"
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   1620
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   149094401
      CurrentDate     =   43932
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Voeg nieuwe pool toe"
      Height          =   375
      Left            =   0
      TabIndex        =   33
      Tag             =   "kop"
      Top             =   120
      Width           =   5775
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Inleg "
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   1200
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
      Top             =   660
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "tot"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   1680
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
      Top             =   1680
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
      Top             =   1200
      Width           =   855
   End
End
Attribute VB_Name = "newPoolForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClose_Click()
Dim sqlstr As String
    'save the record
    'build save string
    Dim i As Integer
    With Me
        sqlstr = "insert into tblPools (tournamentID, poolName, poolStartAcceptForms, poolEndAcceptforms, "
        sqlstr = sqlstr & "poolcost, prizeHighDayScore, prizeHighDayOverallPosition, prizeLowDayOverallPosition, "
        sqlstr = sqlstr & "prizePercentageFirst, prizePercentageSecond, prizePercentageThird, prizePercentageFourth, "
        sqlstr = sqlstr & "prizeLowFinalOverallPosition) VALUES ("
        sqlstr = sqlstr & Val(.cmbTournaments.BoundText) & ", '" & .txtPoolName & "', " & CDbl(.dtpStart) & ", " & CDbl(.dtpStart) & ", "
        sqlstr = sqlstr & float(.txtCosts) & ", " & float(.txtHighestDayscore) & ", " & float(.txtHighestPosition) & ", " & float(.txtLowestPosition) & ", "
        For i = 0 To 3
            sqlstr = sqlstr & float(.txtPercentage(i)) & ", "
        Next
        sqlstr = sqlstr & float(.txtPrizeLastOverall) & ")"
    End With
    cn.Execute sqlstr
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim ctl As Control
Dim sqlstr As String
Dim i As Integer
Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
'set Form defaults
    UnifyForm Me

'back color of frame
    Me.frmPrizes.BackColor = Me.BackColor
'basis tabel

'fill tournament combo
    Me.txtPoolName.DataField = "poolName"
    sqlstr = "Select tournamentID, tournamentType & ' - ' & tournamentYear as tournament from tblTournaments order by tournamentYear"
    rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
    With Me.cmbTournaments
        Set .RowSource = rs
        .BoundColumn = "tournamentId"
        .ListField = "tournament"
    End With
    
    If (rs.State And adStateOpen) = adStateOpen Then rs.Close
    Set rs = Nothing
    
End Sub

Sub calcTotalPercentage()
'calculate the total of the percentage prizes
    Dim totalPerc As Double
    Dim i As Integer
    
    For i = 0 To 3
        totalPerc = totalPerc + Val(Me.txtPercentage(i).Text)
    Next
    Me.lblTotal.Caption = Format(totalPerc / 100, "0%")
    If totalPerc <> 100 Then
        Me.lblTotal.ForeColor = vbRed
    Else
        Me.lblTotal.ForeColor = Me.Label15.ForeColor
    End If
End Sub

Private Sub txtPercentage_Change(Index As Integer)
    calcTotalPercentage
End Sub

