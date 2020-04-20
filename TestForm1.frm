VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
    If rs.AbsolutePosition < rs.RecordCount Then rs.MoveNext
    'Me.Label1.Caption = "nr: " & rs.CursorLocation
    Me.Label1 = "Nr: " & rs.AbsolutePosition
End Sub

Private Sub Command2_Click()
    If rs.AbsolutePosition > 1 Then rs.MovePrevious
    Me.Label1 = "Nr: " & rs.AbsolutePosition
End Sub

Private Sub Form_Load()
rs.Open "Select * from ks", cn, adOpenKeyset, adLockOptimistic
Set Me.Text1.DataSource = rs
Me.Text1.DataField = "jaar"

End Sub
