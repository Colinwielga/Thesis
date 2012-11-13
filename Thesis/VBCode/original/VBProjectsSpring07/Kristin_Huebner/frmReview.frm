VERSION 5.00
Begin VB.Form frmReview 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   11070
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   13710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11070
   ScaleWidth      =   13710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   9240
      TabIndex        =   22
      Top             =   10080
      Width           =   2415
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
      TabIndex        =   17
      Top             =   8280
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
      TabIndex        =   16
      Top             =   8280
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
      TabIndex        =   15
      Top             =   7920
      Width           =   7575
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
      TabIndex        =   14
      Top             =   8760
      Width           =   15255
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next Work"
      Height          =   615
      Left            =   2760
      TabIndex        =   13
      Top             =   10080
      Width           =   2295
   End
   Begin VB.PictureBox picChandraYakshi 
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   5400
      Picture         =   "frmReview.frx":0000
      ScaleHeight     =   5655
      ScaleWidth      =   3975
      TabIndex        =   12
      Top             =   1080
      Width           =   3975
   End
   Begin VB.PictureBox picQueenMayasDream 
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   5400
      Picture         =   "frmReview.frx":4E22
      ScaleHeight     =   6015
      ScaleWidth      =   3975
      TabIndex        =   11
      Top             =   840
      Width           =   3975
   End
   Begin VB.PictureBox picRailingandGate 
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   2760
      Picture         =   "frmReview.frx":C2D9
      ScaleHeight     =   5415
      ScaleWidth      =   9015
      TabIndex        =   10
      Top             =   1440
      Width           =   9015
   End
   Begin VB.PictureBox picYakshi 
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   4920
      Picture         =   "frmReview.frx":284EF
      ScaleHeight     =   5055
      ScaleWidth      =   5295
      TabIndex        =   9
      Top             =   1320
      Width           =   5295
   End
   Begin VB.PictureBox picGreatMonkeyJataka 
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   5280
      Picture         =   "frmReview.frx":39AB0
      ScaleHeight     =   6375
      ScaleWidth      =   4695
      TabIndex        =   8
      Top             =   480
      Width           =   4695
   End
   Begin VB.PictureBox picSeatedBuddha 
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   5400
      Picture         =   "frmReview.frx":3EFA1
      ScaleHeight     =   6015
      ScaleWidth      =   4455
      TabIndex        =   7
      Top             =   960
      Width           =   4455
   End
   Begin VB.PictureBox picStandingBuddha 
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   2520
      Picture         =   "frmReview.frx":45ACC
      ScaleHeight     =   7815
      ScaleWidth      =   10335
      TabIndex        =   6
      Top             =   0
      Width           =   10335
   End
   Begin VB.PictureBox picKingKanishka 
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   5760
      Picture         =   "frmReview.frx":5D12A
      ScaleHeight     =   6015
      ScaleWidth      =   4095
      TabIndex        =   5
      Top             =   960
      Width           =   4095
   End
   Begin VB.PictureBox picGreatBath 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   4680
      Picture         =   "frmReview.frx":62AA4
      ScaleHeight     =   3495
      ScaleWidth      =   6255
      TabIndex        =   4
      Top             =   2160
      Width           =   6255
   End
   Begin VB.PictureBox picSeatedHumanSeal 
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   3360
      Picture         =   "frmReview.frx":677F6
      ScaleHeight     =   7815
      ScaleWidth      =   8895
      TabIndex        =   3
      Top             =   0
      Width           =   8895
   End
   Begin VB.PictureBox picTerracottaFigurine 
      BorderStyle     =   0  'None
      Height          =   6135
      Left            =   6000
      Picture         =   "frmReview.frx":F48C8
      ScaleHeight     =   6135
      ScaleWidth      =   3375
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.PictureBox picBustofaman 
      BorderStyle     =   0  'None
      Height          =   6135
      Left            =   5640
      Picture         =   "frmReview.frx":F9AA1
      ScaleHeight     =   6135
      ScaleWidth      =   4095
      TabIndex        =   1
      Top             =   720
      Width           =   4095
   End
   Begin VB.PictureBox picLionCapital 
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   5760
      Picture         =   "frmReview.frx":FDE0F
      ScaleHeight     =   6015
      ScaleWidth      =   3975
      TabIndex        =   0
      Top             =   720
      Width           =   3975
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
      TabIndex        =   21
      Top             =   8400
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
      TabIndex        =   20
      Top             =   8280
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
      TabIndex        =   19
      Top             =   8280
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
      TabIndex        =   18
      Top             =   7920
      Width           =   615
   End
End
Attribute VB_Name = "frmReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'searches items to display those indicated in Quiz as needing review
Dim Pos As Integer


Private Sub cmdBack_Choose_Test_Click() 'goes back to result page
    frmReview.Hide
    frmResults.Show
End Sub

Private Sub cmdNext_Click() 'searches and shows items indicated as needing review
    Dim Counter As Integer
    Pos = Pos + 1
    lblTitle.Visible = True
    lblArtist.Visible = True
    lblinfo.Visible = True
    lblDate.Visible = True
    picArtist.Visible = True
    picTitle.Visible = True
    picDate.Visible = True
    picInfo.Visible = True
    
    picInfo.Cls
    picArtist.Cls
    picDate.Cls
    picTitle.Cls
    
    Do Until Review(Pos) = True Or Pos > 13
        Pos = Pos + 1
    Loop

    If Pos > 13 Then
        MsgBox "I have nothing for you to review.", , "Well done" 'displays message and ends review form when the search has been exhausted
        frmReview.Hide
        frmChoose_Test.Show
    End If
    
    If Pos = 1 And Review(Pos) = True Then
        Counter = Counter + 1
        
        picTitle.Print titles(Pos)
        picArtist.Print artists(Pos)
        picDate.Print workdate(Pos)
        picInfo.Print extrainfos(Pos)
        picInfo.Print extrainfos2(Pos)
    
        picGreatBath.Visible = True
        picSeatedHumanSeal.Visible = False
        picTerracottaFigurine.Visible = False
        picBustofaman.Visible = False
        picLionCapital.Visible = False
        picChandraYakshi.Visible = False
        picQueenMayasDream.Visible = False
        picRailingandGate.Visible = False
        picYakshi.Visible = False
        picGreatMonkeyJataka.Visible = False
        picSeatedBuddha.Visible = False
        picStandingBuddha.Visible = False
        picKingKanishka.Visible = False
    End If
    
    If Pos = 2 And Review(Pos) = True Then
      Counter = Counter + 1
      
        picTitle.Print titles(Pos)
        picArtist.Print artists(Pos)
        picDate.Print workdate(Pos)
        picInfo.Print extrainfos(Pos)
        picInfo.Print extrainfos2(Pos)
    
        picGreatBath.Visible = False
        picSeatedHumanSeal.Visible = True
        picTerracottaFigurine.Visible = False
        picBustofaman.Visible = False
        picLionCapital.Visible = False
        picChandraYakshi.Visible = False
        picQueenMayasDream.Visible = False
        picRailingandGate.Visible = False
        picYakshi.Visible = False
        picGreatMonkeyJataka.Visible = False
        picSeatedBuddha.Visible = False
        picStandingBuddha.Visible = False
        picKingKanishka.Visible = False
    End If
    
    If Review(Pos) = True And Pos = 3 Then
      Counter = Counter + 1
        picTitle.Print titles(Pos)
        picArtist.Print artists(Pos)
        picDate.Print workdate(Pos)
        picInfo.Print extrainfos(Pos)
        picInfo.Print extrainfos2(Pos)
    
        picGreatBath.Visible = False
        picSeatedHumanSeal.Visible = False
        picTerracottaFigurine.Visible = True
        picBustofaman.Visible = False
        picLionCapital.Visible = False
        picChandraYakshi.Visible = False
        picQueenMayasDream.Visible = False
        picRailingandGate.Visible = False
        picYakshi.Visible = False
        picGreatMonkeyJataka.Visible = False
        picSeatedBuddha.Visible = False
        picStandingBuddha.Visible = False
        picKingKanishka.Visible = False
    End If
    
    If Review(Pos) = True And Pos = 4 Then
      Counter = Counter + 1
        picTitle.Print titles(Pos)
        picArtist.Print artists(Pos)
        picDate.Print workdate(Pos)
        picInfo.Print extrainfos(Pos)
        picInfo.Print extrainfos2(Pos)
        
        picGreatBath.Visible = False
        picSeatedHumanSeal.Visible = False
        picTerracottaFigurine.Visible = False
        picBustofaman.Visible = True
        picLionCapital.Visible = False
        picChandraYakshi.Visible = False
        picQueenMayasDream.Visible = False
        picRailingandGate.Visible = False
        picYakshi.Visible = False
        picGreatMonkeyJataka.Visible = False
        picSeatedBuddha.Visible = False
        picStandingBuddha.Visible = False
        picKingKanishka.Visible = False
    End If
    
    If Review(Pos) = True And Pos = 5 Then
      Counter = Counter + 1
        picTitle.Print titles(Pos)
        picArtist.Print artists(Pos)
        picDate.Print workdate(Pos)
        picInfo.Print extrainfos(Pos)
        picInfo.Print extrainfos2(Pos)
        
        picGreatBath.Visible = False
        picSeatedHumanSeal.Visible = False
        picTerracottaFigurine.Visible = False
        picBustofaman.Visible = False
        picLionCapital.Visible = True
        picChandraYakshi.Visible = False
        picQueenMayasDream.Visible = False
        picRailingandGate.Visible = False
        picYakshi.Visible = False
        picGreatMonkeyJataka.Visible = False
        picSeatedBuddha.Visible = False
        picStandingBuddha.Visible = False
        picKingKanishka.Visible = False
    End If
    
    If Review(Pos) = True And Pos = 6 Then
      Counter = Counter + 1
        picTitle.Print titles(Pos)
        picArtist.Print artists(Pos)
        picDate.Print workdate(Pos)
        picInfo.Print extrainfos(Pos)
        picInfo.Print extrainfos2(Pos)
    
        picGreatBath.Visible = False
        picSeatedHumanSeal.Visible = False
        picTerracottaFigurine.Visible = False
        picBustofaman.Visible = False
        picLionCapital.Visible = False
        picChandraYakshi.Visible = True
        picQueenMayasDream.Visible = False
        picRailingandGate.Visible = False
        picYakshi.Visible = False
        picGreatMonkeyJataka.Visible = False
        picSeatedBuddha.Visible = False
        picStandingBuddha.Visible = False
        picKingKanishka.Visible = False
    End If
    
    
    If Review(Pos) = True And Pos = 7 Then
      Counter = Counter + 1
        picTitle.Print titles(Pos)
        picArtist.Print artists(Pos)
        picDate.Print workdate(Pos)
        picInfo.Print extrainfos(Pos)
        picInfo.Print extrainfos2(Pos)
    
        picGreatBath.Visible = False
        picSeatedHumanSeal.Visible = False
        picTerracottaFigurine.Visible = False
        picBustofaman.Visible = False
        picLionCapital.Visible = False
        picChandraYakshi.Visible = False
        picQueenMayasDream.Visible = True
        picRailingandGate.Visible = False
        picYakshi.Visible = False
        picGreatMonkeyJataka.Visible = False
        picSeatedBuddha.Visible = False
        picStandingBuddha.Visible = False
        picKingKanishka.Visible = False
    End If
    
    
    If Review(Pos) = True And Pos = 8 Then
      Counter = Counter + 1
        picTitle.Print titles(Pos)
        picArtist.Print artists(Pos)
        picDate.Print workdate(Pos)
        picInfo.Print extrainfos(Pos)
        picInfo.Print extrainfos2(Pos)
    
        picGreatBath.Visible = False
        picSeatedHumanSeal.Visible = False
        picTerracottaFigurine.Visible = False
        picBustofaman.Visible = False
        picLionCapital.Visible = False
        picChandraYakshi.Visible = False
        picQueenMayasDream.Visible = False
        picRailingandGate.Visible = True
        picYakshi.Visible = False
        picGreatMonkeyJataka.Visible = False
        picSeatedBuddha.Visible = False
        picStandingBuddha.Visible = False
        picKingKanishka.Visible = False
    End If
    
    If Review(Pos) = True And Pos = 9 Then
      Counter = Counter + 1
        picTitle.Print titles(Pos)
        picArtist.Print artists(Pos)
        picDate.Print workdate(Pos)
        picInfo.Print extrainfos(Pos)
        picInfo.Print extrainfos2(Pos)
      
        picGreatBath.Visible = False
        picSeatedHumanSeal.Visible = False
        picTerracottaFigurine.Visible = False
        picBustofaman.Visible = False
        picLionCapital.Visible = False
        picChandraYakshi.Visible = False
        picQueenMayasDream.Visible = False
        picRailingandGate.Visible = False
        picYakshi.Visible = True
        picGreatMonkeyJataka.Visible = False
        picSeatedBuddha.Visible = False
        picStandingBuddha.Visible = False
        picKingKanishka.Visible = False
    End If
    
    If Review(Pos) = True And Pos = 10 Then
      Counter = Counter + 1
        picTitle.Print titles(Pos)
        picArtist.Print artists(Pos)
        picDate.Print workdate(Pos)
        picInfo.Print extrainfos(Pos)
        picInfo.Print extrainfos2(Pos)
      
        picGreatBath.Visible = False
        picSeatedHumanSeal.Visible = False
        picTerracottaFigurine.Visible = False
        picBustofaman.Visible = False
        picLionCapital.Visible = False
        picChandraYakshi.Visible = False
        picQueenMayasDream.Visible = False
        picRailingandGate.Visible = False
        picYakshi.Visible = False
        picGreatMonkeyJataka.Visible = True
        picSeatedBuddha.Visible = False
        picStandingBuddha.Visible = False
        picKingKanishka.Visible = False
    End If
    
    
    If Review(Pos) = True And Pos = 11 Then
      Counter = Counter + 1
        picTitle.Print titles(Pos)
        picArtist.Print artists(Pos)
        picDate.Print workdate(Pos)
        picInfo.Print extrainfos(Pos)
        picInfo.Print extrainfos2(Pos)
    
        picGreatBath.Visible = False
        picSeatedHumanSeal.Visible = False
        picTerracottaFigurine.Visible = False
        picBustofaman.Visible = False
        picLionCapital.Visible = False
        picChandraYakshi.Visible = False
        picQueenMayasDream.Visible = False
        picRailingandGate.Visible = False
        picYakshi.Visible = False
        picGreatMonkeyJataka.Visible = False
        picSeatedBuddha.Visible = True
        picStandingBuddha.Visible = False
        picKingKanishka.Visible = False
    End If
    
    
    If Review(Pos) = True And Pos = 12 Then
      Counter = Counter + 1
        picTitle.Print titles(Pos)
        picArtist.Print artists(Pos)
        picDate.Print workdate(Pos)
        picInfo.Print extrainfos(Pos)
        picInfo.Print extrainfos2(Pos)
    
        picGreatBath.Visible = False
        picSeatedHumanSeal.Visible = False
        picTerracottaFigurine.Visible = False
        picBustofaman.Visible = False
        picLionCapital.Visible = False
        picChandraYakshi.Visible = False
        picQueenMayasDream.Visible = False
        picRailingandGate.Visible = False
        picYakshi.Visible = False
        picGreatMonkeyJataka.Visible = False
        picSeatedBuddha.Visible = False
        picStandingBuddha.Visible = True
        picKingKanishka.Visible = False
    End If
    
    If Review(Pos) = True And Pos = 13 Then
      Counter = Counter + 1
        picTitle.Print titles(Pos)
        picArtist.Print artists(Pos)
        picDate.Print workdate(Pos)
        picInfo.Print extrainfos(Pos)
        picInfo.Print extrainfos2(Pos)
    
        picGreatBath.Visible = False
        picSeatedHumanSeal.Visible = False
        picTerracottaFigurine.Visible = False
        picBustofaman.Visible = False
        picLionCapital.Visible = False
        picChandraYakshi.Visible = False
        picQueenMayasDream.Visible = False
        picRailingandGate.Visible = False
        picYakshi.Visible = False
        picGreatMonkeyJataka.Visible = False
        picSeatedBuddha.Visible = False
        picStandingBuddha.Visible = False
        picKingKanishka.Visible = True
    End If

    
    
End Sub
    

Private Sub Form_Activate() 'hides everything except next work button when form appears
        picGreatBath.Visible = False
        picSeatedHumanSeal.Visible = False
        picTerracottaFigurine.Visible = False
        picBustofaman.Visible = False
        picLionCapital.Visible = False
        picChandraYakshi.Visible = False
        picQueenMayasDream.Visible = False
        picRailingandGate.Visible = False
        picYakshi.Visible = False
        picGreatMonkeyJataka.Visible = False
        picSeatedBuddha.Visible = False
        picStandingBuddha.Visible = False
        picKingKanishka.Visible = False
        lblTitle.Visible = False
        lblArtist.Visible = False
        lblinfo.Visible = False
        lblDate.Visible = False
        picArtist.Visible = False
        picTitle.Visible = False
        picDate.Visible = False
        picInfo.Visible = False
End Sub

