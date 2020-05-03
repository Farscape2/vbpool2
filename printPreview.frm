VERSION 5.00
Begin VB.Form printPreview 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Afdruk voorbeeld"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   1005
   ClientWidth     =   10860
   FillColor       =   &H000000FF&
   HelpContextID   =   460
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   10860
   Begin VB.PictureBox vscrlHolder 
      Align           =   4  'Align Right
      Height          =   10140
      Left            =   10575
      Negotiate       =   -1  'True
      ScaleHeight     =   10080
      ScaleWidth      =   225
      TabIndex        =   5
      Top             =   450
      Width           =   285
      Begin VB.VScrollBar VScroll1 
         Height          =   10005
         LargeChange     =   5000
         Left            =   0
         SmallChange     =   1000
         TabIndex        =   6
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.PictureBox hscrlHolder 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      Negotiate       =   -1  'True
      ScaleHeight     =   300
      ScaleWidth      =   10800
      TabIndex        =   3
      Top             =   10590
      Width           =   10860
      Begin VB.HScrollBar HScroll1 
         Height          =   285
         LargeChange     =   5000
         Left            =   0
         SmallChange     =   1000
         TabIndex        =   4
         Top             =   0
         Width           =   7200
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      Negotiate       =   -1  'True
      ScaleHeight     =   390
      ScaleWidth      =   10800
      TabIndex        =   0
      Top             =   0
      Width           =   10860
      Begin VB.ComboBox cmbZoom 
         Height          =   315
         Left            =   840
         TabIndex        =   10
         Text            =   "100%"
         Top             =   45
         Width           =   1335
      End
      Begin VB.CommandButton btnPrint 
         Caption         =   "Afdrukken"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7275
         TabIndex        =   7
         Top             =   30
         Width           =   1080
      End
      Begin VB.CommandButton btnNext 
         Caption         =   "Volgende>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6150
         TabIndex        =   2
         Top             =   30
         Width           =   1080
      End
      Begin VB.CommandButton btnPrev 
         Caption         =   "< Vorige"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5025
         TabIndex        =   1
         Top             =   30
         Width           =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "Zoom"
         Height          =   285
         Left            =   225
         TabIndex        =   11
         Top             =   60
         Width           =   495
      End
   End
   Begin VB.PictureBox pageHolder 
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   9930
      Left            =   45
      Negotiate       =   -1  'True
      ScaleHeight     =   9870
      ScaleWidth      =   8310
      TabIndex        =   8
      Top             =   465
      Width           =   8370
      Begin VB.PictureBox pageContent 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9270
         Left            =   120
         ScaleHeight     =   9270
         ScaleWidth      =   7905
         TabIndex        =   12
         Top             =   120
         Width           =   7905
      End
      Begin VB.PictureBox printPages 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   9600
         Index           =   0
         Left            =   105
         ScaleHeight     =   9600
         ScaleLeft       =   -283
         ScaleMode       =   0  'User
         ScaleTop        =   -283
         ScaleWidth      =   8085
         TabIndex        =   9
         Top             =   60
         Visible         =   0   'False
         Width           =   8085
      End
   End
End
Attribute VB_Name = "printPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim currentPage As Integer
Dim printRatio As Double
Dim zoomFactor As Double 'zoom factor

Private Sub btnNext_Click()
    zoomFactor = val(Me.cmbZoom) / 100
    If currentPage < Me.printPages.UBound Then
        currentPage = currentPage + 1
        Me.pageContent.Cls
        With Me.printPages(currentPage)
            Me.pageContent.Move .Left * zoomFactor, .Top * zoomFactor, .Width * zoomFactor, .Height * zoomFactor
            Me.pageContent.PaintPicture .Image, 0, 0, .Width * zoomFactor, .Height * zoomFactor
            Me.pageContent.Refresh
        End With
        Me.printPages(currentPage).ZOrder
    End If
    Me.btnPrev.Enabled = currentPage > 0
    Me.btnNext.Enabled = currentPage < Me.printPages.UBound
End Sub

Private Sub btnPrint_Click()
    With frmPrintDialog
        .btnPrint_Click 0
    End With
End Sub

Private Sub btnPrev_Click()
    zoomFactor = val(Me.cmbZoom) / 100
    If currentPage > 0 Then
        currentPage = currentPage - 1
        Me.pageContent.Cls
        With Me.printPages(currentPage)
            Me.pageContent.Move .Left * zoomFactor, .Top * zoomFactor, .Width * zoomFactor, .Height * zoomFactor
            Me.pageContent.PaintPicture .Image, 0, 0, .Width * zoomFactor, .Height * zoomFactor
            Me.pageContent.Refresh
        End With
        'Me.printPages(currentPage).ZOrder
    End If
    Me.btnPrev.Enabled = currentPage > 0
    Me.btnNext.Enabled = currentPage < Me.printPages.UBound
End Sub

Private Function ScalePicPreviewToPrinter(picPreview As PictureBox) As Double
    
    Dim Ratio As Double ' Ratio between Printer and Picture
    Dim LRGap As Double, TBGap As Double
    Dim HeightRatio As Double, WidthRatio As Double
    Dim PgWidth As Double, PgHeight As Double
    Dim smtemp As Long
    
    ' Get the physical page size in twips ('Inches):
    PgWidth = Printer.Width '/ 1440
    PgHeight = Printer.Height ' / 1440
    
    ' Find the size of the non-printable area on the printer to
    ' use to offset coordinates. These formulas assume the
    ' printable area is centered on the page:
    smtemp = Printer.ScaleMode
    'Printer.ScaleMode = vbInches
    ' Scale PictureBox to Printer's printable area in Inches:
    
    picPreview.ScaleMode = vbTwips
    
    LRGap = (PgWidth - Printer.ScaleWidth) / 2
    TBGap = (PgHeight - Printer.ScaleHeight) / 2
    'Me.printPages(0).Container.BackColor = vbBlue
    Printer.ScaleMode = smtemp
    
    
    ' Compare the height and with ratios to determine the
    ' Ratio to use and how to size the picture box:
    HeightRatio = picPreview.Container.ScaleHeight / PgHeight
    WidthRatio = picPreview.Container.ScaleWidth / PgWidth
    
    If HeightRatio < WidthRatio Then
        Ratio = HeightRatio
    Else
        Ratio = WidthRatio
    End If
    'Ratio = picPreview.FontSize / 8
    picPreview.Container.Height = PgHeight * Ratio
    picPreview.Container.Width = PgWidth * Ratio
    Me.printPages(0).Top = TBGap * Ratio
    Me.printPages(0).Left = LRGap * Ratio
    Me.printPages(0).Height = Me.printPages(0).Container.Height - 1 * TBGap * Ratio
    Me.printPages(0).Width = Me.printPages(0).Container.Width - 2 * LRGap * Ratio
    ' Set default properties of picture box to match printer
    ' There are many that you could add here:
    picPreview.Container.Scale (0, 0)-(PgWidth, PgHeight)
    picPreview.Font.Name = Printer.Font.Name
    picPreview.fontSize = Printer.fontSize * Ratio
    picPreview.ForeColor = Printer.ForeColor
    picPreview.FillStyle = vbTransparent
    picPreview.Cls
    
    ScalePicPreviewToPrinter = Int(Ratio * 100) / 100
'    picPreview.ScaleMode = 1
End Function


Private Sub cmbZoom_Click()
    zoomFactor = val(Me.cmbZoom) / 100 '* 100
    Me.pageHolder.AutoRedraw = True
    Me.pageHolder.Move Me.pageHolder.Left, Me.pageHolder.Top, Printer.Width * zoomFactor, Printer.Height * zoomFactor
    Me.pageContent.Cls
    DoEvents
    With Me.printPages(currentPage)
        Me.pageContent.Move .Left * zoomFactor, .Top * zoomFactor, Printer.ScaleWidth * zoomFactor, Printer.ScaleHeight * zoomFactor
        Me.pageContent.PaintPicture .Image, 0, 0, .Width * zoomFactor, .Height * zoomFactor
        Me.pageContent.Refresh
    End With
End Sub

Private Sub Form_Load()
    Dim prtWidth As Integer
    Dim prtHeight As Integer
    Dim scm As Double
    Dim i As Integer
    
    Me.Font.Size = Printer.fontSize
    Me.Font.Name = Printer.Font.Name
    scm = Me.TextHeight("w") / Printer.TextHeight("w")
    prtWidth = Printer.Width / scm
    prtHeight = Printer.Height / scm
    'eerst op 100 % zetten
    Me.printPages(0).Container.Height = prtHeight
    Me.printPages(0).Container.Width = prtWidth - 10
    printRatio = ScalePicPreviewToPrinter(Me.printPages(0))
    Me.btnPrev.Enabled = False
    currentPage = 0
    For i = 25 To 200 Step 25
        Me.cmbZoom.AddItem i & "%"
    Next
    Me.printPages(0).ScaleMode = vbTwips
    With Me.printPages(0)
        Me.pageContent.Move .Left, .Top, .Width, .Height
        Me.pageContent.PaintPicture .Image, 0, 0, .Width, .Height
    End With
    Me.Width = Me.pageHolder.Width + 1000
    Me.Height = Screen.Height - 1500
    Me.VScroll1.Max = Me.pageHolder.Height - Me.ScaleHeight
    If Me.ScaleWidth > Me.pageHolder.Width Then
      Me.HScroll1.Max = 0
    Else
      Me.HScroll1.Max = Me.pageHolder.Width - Me.ScaleWidth
    End If
    cmbZoom_Click
    centerForm Me
    Me.Visible = True
    
    
End Sub

Private Sub Form_Resize()
Dim i As Integer
    Me.VScroll1.Height = Me.vscrlHolder.Height
    Me.HScroll1.Width = Me.hscrlHolder.Width - Me.vscrlHolder.Width
    Me.pageHolder.Left = 100
    Me.pageHolder.Top = 100 + Me.picButtons.Height
    For i = 0 To Me.printPages.UBound - 1
    Me.printPages(i).Left = Me.HScroll1 * -1 - Me.printPages(0).ScaleLeft
    Me.printPages(i).Top = Me.VScroll1 * -1 + 450
    Me.printPages(i).Top = 240
    Me.printPages(i).Left = 240
    Me.printPages(i).ScaleTop = Printer.ScaleTop - 240 '(Printer.Height - Printer.ScaleHeight) / -2
    Me.printPages(i).ScaleLeft = Printer.ScaleLeft - 240 '(Printer.Width - Printer.ScaleWidth) / -2
    Next
    Me.btnNext.Enabled = Me.printPages.UBound > 0
    If Me.ScaleWidth > Me.pageHolder.Width Then
      Me.HScroll1.Max = 0
    Else
      Me.HScroll1.Max = Me.pageHolder.Width - Me.ScaleWidth
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmPrintDialog.Visible = True
End Sub

Private Sub HScroll1_Change()
    Me.pageHolder.Left = Me.HScroll1 * -1 + 450
End Sub

Private Sub VScroll1_Change()
    Me.pageHolder.Top = Me.VScroll1 * -1 + 450
End Sub


