VERSION 5.00
Begin VB.Form frmBrowse 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Browse Works"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picLionCapital 
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   5760
      Picture         =   "frmBrowse.frx":0000
      ScaleHeight     =   6015
      ScaleWidth      =   3975
      TabIndex        =   14
      Top             =   840
      Width           =   3975
   End
   Begin VB.PictureBox picBustofaman 
      BorderStyle     =   0  'None
      Height          =   6135
      Left            =   5640
      Picture         =   "frmBrowse.frx":486E
      ScaleHeight     =   6135
      ScaleWidth      =   4095
      TabIndex        =   13
      Top             =   840
      Width           =   4095
   End
   Begin VB.PictureBox picTerracottaFigurine 
      BorderStyle     =   0  'None
      Height          =   6135
      Left            =   6000
      Picture         =   "frmBrowse.frx":8BDC
      ScaleHeight     =   6135
      ScaleWidth      =   3375
      TabIndex        =   12
      Top             =   960
      Width           =   3375
   End
   Begin VB.PictureBox picSeatedHumanSeal 
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   3360
      Picture         =   "frmBrowse.frx":DDB5
      ScaleHeight     =   7815
      ScaleWidth      =   8895
      TabIndex        =   11
      Top             =   120
      Width           =   8895
   End
   Begin VB.PictureBox picGreatBath 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   4680
      Picture         =   "frmBrowse.frx":9AE87
      ScaleHeight     =   3495
      ScaleWidth      =   6255
      TabIndex        =   10
      Top             =   2280
      Width           =   6255
   End
   Begin VB.PictureBox picDate 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      ScaleHeight     =   255
      ScaleWidth      =   2055
      TabIndex        =   8
      Top             =   8520
      Width           =   2055
   End
   Begin VB.PictureBox picArtist 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      ScaleHeight     =   255
      ScaleWidth      =   2535
      TabIndex        =   6
      Top             =   8520
      Width           =   2535
   End
   Begin VB.PictureBox picTitle 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      ScaleHeight     =   255
      ScaleWidth      =   7575
      TabIndex        =   4
      Top             =   8160
      Width           =   7575
   End
   Begin VB.CommandButton cmdBack_Choose_Test 
      Caption         =   "Go Back"
      Height          =   615
      Left            =   10320
      TabIndex        =   2
      Top             =   10320
      Width           =   2295
   End
   Begin VB.PictureBox picInfo 
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   15195
      TabIndex        =   1
      Top             =   9000
      Width           =   15255
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next Work"
      Height          =   615
      Left            =   2760
      TabIndex        =   0
      Top             =   10320
      Width           =   2295
   End
   Begin VB.Label lblinfo 
      Caption         =   "Notable Information:"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   8640
      Width           =   2535
   End
   Begin VB.Label lblDate 
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   7
      Top             =   8520
      Width           =   735
   End
   Begin VB.Label lblArtist 
      Caption         =   "Artist:"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   8520
      Width           =   855
   End
   Begin VB.Label lblTitle 
      Caption         =   "Title:"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   3
      Top             =   8160
      Width           =   615
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'shows works and information on works sequentially
Dim work_date As Integer, artist As String, title As String, extrainfo As String, dopos As Integer


Private Sub cmdBack_Choose_Test_Click() 'brings user back to main page
    frmBrowse.Visible = False
    frmChoose_Test.Visible = True
End Sub

Private Sub cmdNext_Click() 'sequentially displays work information and respective picture boxes
picInfo.Cls
picArtist.Cls
picDate.Cls
picTitle.Cls

    dopos = dopos + 1
        
        picTitle.Print titles(dopos + 1)
        picArtist.Print artists(dopos + 1)
        picDate.Print workdate(dopos + 1)
        picInfo.Print extrainfos(dopos + 1)
        picInfo.Print extrainfos2(dopos + 1)

    If dopos = 1 Then
        picGreatBath.Visible = False
        picSeatedHumanSeal.Visible = True
        picTerracottaFigurine.Visible = False
        picBustofaman.Visible = False
        picLionCapital.Visible = False
    End If

    If dopos = 2 Then
        picGreatBath.Visible = False
        picSeatedHumanSeal.Visible = False
        picTerracottaFigurine.Visible = True
        picBustofaman.Visible = False
        picLionCapital.Visible = False
    End If

    If dopos = 3 Then
        picGreatBath.Visible = False
        picSeatedHumanSeal.Visible = False
        picTerracottaFigurine.Visible = False
        picBustofaman.Visible = True
        picLionCapital.Visible = False
    End If

    If dopos = 4 Then
       picGreatBath.Visible = False
        picSeatedHumanSeal.Visible = False
        picTerracottaFigurine.Visible = False
        picBustofaman.Visible = False
        picLionCapital.Visible = True
    End If
    
    If dopos > 4 Then
        frmBrowse.Hide
        frmBrowse2.Show
    End If


End Sub

Private Sub Form_Activate() 'initially displays earliest work


picTitle.Print titles(1)
picArtist.Print artists(1)
picDate.Print workdate(1)
picInfo.Print extrainfos(1)
picInfo.Print extrainfos2(1)
picGreatBath.Visible = True
picSeatedHumanSeal.Visible = False
picTerracottaFigurine.Visible = False
picBustofaman.Visible = False
picLionCapital.Visible = False

dopos = 1

End Sub


