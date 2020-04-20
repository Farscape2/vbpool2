VERSION 5.00
Begin VB.Form frmRemoveDoubleIds 
   Caption         =   "Dubbele Player Id's"
   ClientHeight    =   2730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
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
   ScaleHeight     =   2730
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnClose 
      Caption         =   "Sluiten"
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Vervang"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dubbele Players ID's"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Tag             =   "kop"
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Vervang door"
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Playe ID"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "frmRemoveDoubleIds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Dim sqlstr As String
    sqlstr = "Update tblTeamPlayers set PlayerId = " & Val(Me.Text2)
    sqlstr = sqlstr & " WHERE playerId = " & Val(Me.Text1)
    cn.Execute sqlstr
    sqlstr = "Update tblMatchEvents set PlayerId = " & Val(Me.Text2)
    sqlstr = sqlstr & " WHERE playerId = " & Val(Me.Text1)
    cn.Execute sqlstr
    sqlstr = "Update tblPredictionTopscorers set predictionTopscorePlayerID = " & Val(Me.Text2)
    sqlstr = sqlstr & " WHERE predictionTopscorePlayerID = " & Val(Me.Text1)
    cn.Execute sqlstr
    sqlstr = "Delete from tblPeople where peopleid = " & Me.Text1
    cn.Execute sqlstr
End Sub

