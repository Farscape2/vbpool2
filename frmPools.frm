VERSION 5.00
Begin VB.Form frmPools 
   Caption         =   "Form1"
   ClientHeight    =   3870
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmPools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Caption = getOrganisation()
End Sub
