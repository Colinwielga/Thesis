VERSION 5.00
Begin VB.Form frmBrowse2 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Browse Works"
   ClientHeight    =   10500
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   14040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10500
   ScaleWidth      =   14040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picKingKanishka 
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   5760
      Picture         =   "frmBrowse2.frx":0000
      ScaleHeight     =   6015
      ScaleWidth      =   4095
      TabIndex        =   17
      Top             =   1080
      Width           =   4095
   End
   Begin VB.PictureBox picStandingBuddha 
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   2520
      Picture         =   "frmBrowse2.frx":597A
      ScaleHeight     =   7815
      ScaleWidth      =   10335
      TabIndex        =   16
      Top             =   120
      Width           =   10335
   End
   Begin VB.PictureBox picSeatedBuddha 
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   5400
      Picture         =   "frmBrowse2.frx":1CFD8
      ScaleHeight     =   6015
      ScaleWidth      =   4455
      TabIndex        =   15
      Top             =   1080
      Width           =   4455
   End
   Begin VB.PictureBox picGreatMonkeyJataka 
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   5280
      Picture         =   "frmBrowse2.frx":23B03
      ScaleHeight     =   6375
      ScaleWidth      =   4695
      TabIndex        =   14
      Top             =   600
      Width           =   4695
   End
   Begin VB.PictureBox picYakshi 
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   4920
      Picture         =   "frmBrowse2.frx":28FF4
      ScaleHeight     =   5055
      ScaleWidth      =   5295
      TabIndex        =   13
      Top             =   1440
      Width           =   5295
   End
   Begin VB.PictureBox picRailingandGate 
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   2760
      Picture         =   "frmBrowse2.frx":3A5B5
      ScaleHeight     =   5415
      ScaleWidth      =   9015
      TabIndex        =   12
      Top             =   1560
      Width           =   9015
   End
   Begin VB.PictureBox picQueenMayasDream 
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   5400
      Picture         =   "frmBrowse2.frx":567CB
      ScaleHeight     =   6015
      ScaleWidth      =   3975
      TabIndex        =   11
      Top             =   960
      Width           =   3975
   End
   Begin VB.PictureBox picChandraYakshi 
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   5400
      Picture         =   "frmBrowse2.frx":5DC82
      ScaleHeight     =   5655
      ScaleWidth      =   3975
      TabIndex        =   10
      Top             =   1200
      Width           =   3975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next Work"
      Height          =   615
      Left            =   2760
      TabIndex        =   5
      Top             =   10200
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
      TabIndex        =   4
      Top             =   8880
      Width           =   15255
   End
   Begin VB.CommandButton cmdBack_Choose_Test 
      Caption         =   "Go Back"
      Height          =   615
      Left            =   10320
      TabIndex        =   3
      Top             =   10200
      Width           =   2295
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
      TabIndex        =   2
      Top             =   8040
      Width           =   7575
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
      TabIndex        =   1
      Top             =   8400
      Width           =   2535
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
      TabIndex        =   0
      Top             =   8400
      Width           =   2055
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
      TabIndex        =   9
      Top             =   8040
      Width           =   615
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
      TabIndex        =   8
      Top             =   8400
      Width           =   855
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
      Top             =   8400
      Width           =   735
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
      TabIndex        =   6
      Top             =   8520
      Width           =   2535
   End
End
Attribute VB_Name = "frmBrowse2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'second browse form: functions as first browse form, broken into second form for the convenience of the programmer
Dim dopos As Integer

Private Sub cmdNext_Click()
picInfo.Cls
picArtist.Cls
picDate.Cls
picTitle.Cls

    dopos = dopos + 1
        
        picTitle.Print titles(dopos + 6)
        picArtist.Print artists(dopos + 6)
        picDate.Print workdate(dopos + 6)
        picInfo.Print extrainfos(dopos + 6)
        picInfo.Print extrainfos2(dopos + 6)

     If dopos = 1 Then
        picChandraYakshi.Visible = False
        picQueenMayasDream.Visible = True
        picRailingandGate.Visible = False
        picYakshi.Visible = False
        picGreatMonkeyJataka.Visible = False
        picSeatedBuddha.Visible = False
        picStandingBuddha.Visible = False
        picKingKanishka.Visible = False
    End If
    
     If dopos = 2 Then
        picChandraYakshi.Visible = False
        picQueenMayasDream.Visible = False
        picRailingandGate.Visible = True
        picYakshi.Visible = False
        picGreatMonkeyJataka.Visible = False
        picSeatedBuddha.Visible = False
        picStandingBuddha.Visible = False
        picKingKanishka.Visible = False
    End If
    
    If dopos = 3 Then
        picChandraYakshi.Visible = False
        picQueenMayasDream.Visible = False
        picRailingandGate.Visible = False
        picYakshi.Visible = True
        picGreatMonkeyJataka.Visible = False
        picSeatedBuddha.Visible = False
        picStandingBuddha.Visible = False
        picKingKanishka.Visible = False
    End If
    
    If dopos = 4 Then
        picChandraYakshi.Visible = False
        picQueenMayasDream.Visible = False
        picRailingandGate.Visible = False
        picYakshi.Visible = False
        picGreatMonkeyJataka.Visible = True
        picSeatedBuddha.Visible = False
        picStandingBuddha.Visible = False
        picKingKanishka.Visible = False
    End If
    
    If dopos = 5 Then
        picChandraYakshi.Visible = False
        picQueenMayasDream.Visible = False
        picRailingandGate.Visible = False
        picYakshi.Visible = False
        picGreatMonkeyJataka.Visible = False
        picSeatedBuddha.Visible = True
        picStandingBuddha.Visible = False
        picKingKanishka.Visible = False
    End If
    
     If dopos = 6 Then
        picChandraYakshi.Visible = False
        picQueenMayasDream.Visible = False
        picRailingandGate.Visible = False
        picYakshi.Visible = False
        picGreatMonkeyJataka.Visible = False
        picSeatedBuddha.Visible = False
        picStandingBuddha.Visible = True
        picKingKanishka.Visible = False
    End If
    
     If dopos = 7 Then
        picChandraYakshi.Visible = False
        picQueenMayasDream.Visible = False
        picRailingandGate.Visible = False
        picYakshi.Visible = False
        picGreatMonkeyJataka.Visible = False
        picSeatedBuddha.Visible = False
        picStandingBuddha.Visible = False
        picKingKanishka.Visible = True
    End If
    
    If dopos > 7 Then
        frmBrowse2.Hide
        frmChoose_Test.Show
    End If


End Sub


Private Sub Form_Activate()
    dopos = 0
        picTitle.Print titles(6)
        picArtist.Print artists(6)
        picDate.Print workdate(6)
        picInfo.Print extrainfos(6)
        picInfo.Print extrainfos2(6)
        
        picChandraYakshi.Visible = True
        picQueenMayasDream.Visible = False
        picRailingandGate.Visible = False
        picYakshi.Visible = False
        picGreatMonkeyJataka.Visible = False
        picSeatedBuddha.Visible = False
        picStandingBuddha.Visible = False
        picKingKanishka.Visible = False
        
End Sub

