VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form playerAddForm 
   Caption         =   "Speler toevoegen"
   ClientHeight    =   3630
   ClientLeft      =   16335
   ClientTop       =   6420
   ClientWidth     =   3690
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
   ScaleHeight     =   3630
   ScaleWidth      =   3690
   Begin MSDataListLib.DataCombo cmbCountry 
      Height          =   360
      Left            =   1440
      TabIndex        =   11
      Top             =   2520
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Opslaan"
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox txtNickName 
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtAName 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtTname 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtVnaam 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Land"
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nieuwe speler"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Tag             =   "kop"
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bijnaam"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Achternaam"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "tussenvg"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Voornaam"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "playerAddForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Private currentCountry As Long

Public Property Get Country() As Long
    Country = currentCountry
End Property

Public Property Let Country(ByVal NewValue As Long)
    currentCountry = NewValue
End Property

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
    Dim sqlstr As String
    If Me.txtNickName = "" Then Me.txtNickName = buildNickName
        
    If Not playerExists(Me.txtVnaam, Me.txtTname, Me.txtAName, Me.txtNickName, cn) Then
        sqlstr = "Insert into tblPeople (firstName, middleName, lastName, nickName, function1, countryCode"
        sqlstr = sqlstr & ") VALUES ('" & Me.txtVnaam
        sqlstr = sqlstr & "','" & Me.txtTname
        sqlstr = sqlstr & "','" & Me.txtAName
        sqlstr = sqlstr & "','" & Me.txtNickName
        sqlstr = sqlstr & "', 3" 'make it a midfielder, for now
        sqlstr = sqlstr & "," & val(Me.cmbCountry.BoundText)
        sqlstr = sqlstr & ")"
        cn.Execute sqlstr
    Else
        MsgBox "Speler bestaat al", vbOKOnly + vbInformation, "Speler toevoegen"
    End If
    Unload Me
End Sub

Private Sub Form_Load()
Dim sqlstr As String
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = lclConn()
        .Open
    End With
    sqlstr = "Select * from tblCountries "
    Set rs = New ADODB.Recordset
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    With Me.cmbCountry
        Set .RowSource = rs
        .ListField = "CountryName"
        .BoundColumn = "countryId"
        .Refresh
    End With
    Me.cmbCountry.BoundText = Country
    UnifyForm Me
End Sub

Function buildNickName()
Dim NickName As String
    NickName = Me.txtAName
    If Me.txtVnaam > "" Or Me.txtTname > "" Then
        If NickName > "" Then NickName = NickName & ","
    End If
    If Me.txtVnaam > "" Then NickName = NickName & " " & Me.txtVnaam
    If Me.txtTname > "" Then NickName = NickName & " " & Me.txtTname
    buildNickName = Trim(NickName)
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Clean-up procedure
    If Not rs Is Nothing Then
        'first, check if the state is open, if yes then close it
        If (rs.State And adStateOpen) = adStateOpen Then
            rs.Close
        End If
        'set them to nothing
        Set rs = Nothing
    End If
    'same comment with rs
    If Not cn Is Nothing Then
        If (cn.State And adStateOpen) = adStateOpen Then
            cn.Close
        End If
        Set cn = Nothing
    End If
End Sub

Private Sub Label4_Click()
    Me.txtNickName = buildNickName
End Sub

