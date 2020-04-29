VERSION 5.00
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
   Begin VB.ComboBox cmbTournaments 
      Height          =   360
      Left            =   1080
      TabIndex        =   26
      Top             =   1147
      Width           =   1575
   End
   Begin VB.Frame frmPrizes 
      Caption         =   "Prijzen"
      Height          =   2295
      Left            =   0
      TabIndex        =   9
      Top             =   2160
      Width           =   5775
      Begin MSComCtl2.UpDown upDnPerc 
         Height          =   375
         Index           =   0
         Left            =   3720
         TabIndex        =   27
         Top             =   660
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         BuddyControl    =   "txtPercentage(0)"
         BuddyDispid     =   196634
         BuddyIndex      =   0
         OrigLeft        =   3720
         OrigTop         =   660
         OrigRight       =   3975
         OrigBottom      =   1035
         Increment       =   5
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   22
         Enabled         =   -1  'True
      End
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
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   13
         Top             =   660
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   661
         _Version        =   393216
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
         TabIndex        =   14
         Top             =   1770
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
         Index           =   1
         Left            =   4800
         TabIndex        =   28
         Top             =   660
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   661
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSComCtl2.UpDown upDnPerc 
         Height          =   375
         Index           =   1
         Left            =   5265
         TabIndex        =   30
         Top             =   660
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         BuddyControl    =   "txtPercentage(1)"
         BuddyDispid     =   196634
         BuddyIndex      =   1
         OrigLeft        =   3720
         OrigTop         =   660
         OrigRight       =   3975
         OrigBottom      =   1035
         Increment       =   5
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   22
         Enabled         =   -1  'True
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
         Index           =   2
         Left            =   3240
         TabIndex        =   31
         Top             =   1200
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   661
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSComCtl2.UpDown upDnPerc 
         Height          =   375
         Index           =   2
         Left            =   3705
         TabIndex        =   33
         Top             =   1200
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         BuddyControl    =   "txtPercentage(2)"
         BuddyDispid     =   196634
         BuddyIndex      =   2
         OrigLeft        =   3720
         OrigTop         =   660
         OrigRight       =   3975
         OrigBottom      =   1035
         Increment       =   5
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   22
         Enabled         =   -1  'True
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
         Index           =   3
         Left            =   4800
         TabIndex        =   34
         Top             =   1200
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   661
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSComCtl2.UpDown upDnPerc 
         Height          =   375
         Index           =   3
         Left            =   5265
         TabIndex        =   36
         Top             =   1200
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         BuddyControl    =   "txtPercentage(3)"
         BuddyDispid     =   196634
         BuddyIndex      =   3
         OrigLeft        =   3720
         OrigTop         =   660
         OrigRight       =   3975
         OrigBottom      =   1035
         Increment       =   5
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   22
         Enabled         =   -1  'True
      End
      Begin VB.Label lblTotal 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   4680
         TabIndex        =   37
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "4e"
         Height          =   375
         Index           =   3
         Left            =   4440
         TabIndex        =   35
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "3e"
         Height          =   375
         Index           =   2
         Left            =   2880
         TabIndex        =   32
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "2e"
         Height          =   375
         Index           =   1
         Left            =   4440
         TabIndex        =   29
         Top             =   660
         Width           =   375
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Laatste"
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   2880
         TabIndex        =   24
         Top             =   1830
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Onderaan"
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   120
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "1e"
         Height          =   375
         Index           =   0
         Left            =   2880
         TabIndex        =   20
         Top             =   660
         Width           =   375
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Eindstand (percentages)"
         Height          =   255
         Left            =   3000
         TabIndex        =   19
         Top             =   240
         Width           =   2295
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
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSMask.MaskEdBox txtCosts 
      DataSource      =   "dtcPools"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
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
      TabIndex        =   17
      Top             =   4455
      Width           =   5790
      Begin VB.CommandButton btnCancel 
         Cancel          =   -1  'True
         Caption         =   "Annuleren"
         Height          =   495
         Left            =   3000
         TabIndex        =   15
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton btnClose 
         Caption         =   "Opslaan"
         Default         =   -1  'True
         Height          =   495
         Left            =   4320
         TabIndex        =   16
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSComCtl2.DTPicker dtpStart 
      DataSource      =   "dtcPools"
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   1620
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   146931713
      CurrentDate     =   43932
   End
   Begin MSComCtl2.DTPicker dtpEind 
      DataSource      =   "dtcPools"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   1620
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   146931713
      CurrentDate     =   43932
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Voeg nieuwe pool toe"
      Height          =   375
      Left            =   0
      TabIndex        =   25
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
      TabIndex        =   3
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
      TabIndex        =   7
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
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Toernooi"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   0
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
    Dim i As Integer
    Dim tournID As Long
    tournID = Me.cmbTournaments.ItemData(Me.cmbTournaments.ListIndex)
    With Me
        'build save string
        sqlstr = "insert into tblPools (tournamentID, poolName, poolStartAcceptForms, poolEndAcceptforms, "
        sqlstr = sqlstr & "poolcost, prizeHighDayScore, prizeHighDayOverallPosition, prizeLowDayOverallPosition, "
        sqlstr = sqlstr & "prizePercentageFirst, prizePercentageSecond, prizePercentageThird, prizePercentageFourth, "
        sqlstr = sqlstr & "prizeLowFinalOverallPosition) VALUES ("
        sqlstr = sqlstr & tournID & ", '" & .txtPoolName & "', " & CDbl(.dtpStart) & ", " & CDbl(.dtpStart) & ", "
        sqlstr = sqlstr & float(.txtCosts) & ", " & float(.txtHighestDayscore) & ", " & float(.txtHighestPosition) & ", " & float(.txtLowestPosition) & ", "
        For i = 0 To 3
            sqlstr = sqlstr & float(.txtPercentage(i) / 100) & ", "
        Next
        sqlstr = sqlstr & float(.txtPrizeLastOverall) & ")"
    End With
    If Not cnOpen(cn) Then openDB
    cn.Execute sqlstr
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim sqlstr As String
Dim i As Integer
Dim rs As adodb.Recordset
    Set rs = New adodb.Recordset
'set Form defaults
    UnifyForm Me

'back color of frame
    Me.frmPrizes.BackColor = Me.BackColor
'basis tabel

'fill tournament combo
    sqlstr = "Select tournamentID, "
    If dBaseType = "ACCESS" Then
        sqlstr = sqlstr & " tournamentYear & ' - '  & tournamentType "
    Else
        sqlstr = sqlstr & " concat(tournamentYear, ' - ', tournamentType) "
    End If
    sqlstr = sqlstr & " as tournament from tblTournaments order by tournamentYear"
    FillCombo Me.cmbTournaments, sqlstr, "tournament", "tournamentID"
    
    If (rs.State And adStateOpen) = adStateOpen Then rs.Close
    Set rs = Nothing
    
End Sub

Sub calcTotalPercentage()
'calculate the total of the percentage prizes
    Dim totalPerc As Double
    Dim i As Integer
    
    For i = 0 To 3
        totalPerc = totalPerc + Val(float(Me.txtPercentage(i).Text))
    Next
    Me.lblTotal.Caption = Format(totalPerc / 100, "0%")
    If totalPerc <> 100 Then
        Me.lblTotal.ForeColor = vbRed
    Else
        Me.lblTotal.ForeColor = Me.Label1.ForeColor
    End If
End Sub

Private Sub txtPercentage_LostFocus(Index As Integer)
    calcTotalPercentage
End Sub
