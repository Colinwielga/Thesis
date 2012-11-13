VERSION 5.00
Begin VB.Form frmTop 
   BackColor       =   &H00FF0000&
   Caption         =   "Top Ten Movies"
   ClientHeight    =   8775
   ClientLeft      =   2310
   ClientTop       =   1710
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   10830
   Begin VB.ListBox m 
      BeginProperty Font 
         Name            =   "@Batang"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      ItemData        =   "frmTop.frx":0000
      Left            =   2760
      List            =   "frmTop.frx":0007
      TabIndex        =   28
      Top             =   6960
      Width           =   3615
   End
   Begin VB.CommandButton cmdStats 
      BackColor       =   &H0000FFFF&
      Caption         =   "Click Me for Additional Movie Stats"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6960
      Width           =   3135
   End
   Begin VB.CommandButton cmdSnow 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton cmdDalmations 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton cmdAladdin 
      BackColor       =   &H0000FFFF&
      Height          =   615
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton cmdSleeping 
      BackColor       =   &H0000FFFF&
      Height          =   615
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton cmdMulan 
      BackColor       =   &H0000FFFF&
      Height          =   615
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton cmdMermaid 
      BackColor       =   &H0000FFFF&
      Height          =   615
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdLion 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdCinderella 
      BackColor       =   &H0000FFFF&
      Height          =   615
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdBeauty 
      BackColor       =   &H0000FFFF&
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdBambi 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1680
      Width           =   615
   End
   Begin VB.PictureBox Picture11 
      Height          =   1335
      Left            =   9360
      Picture         =   "frmTop.frx":001B
      ScaleHeight     =   1275
      ScaleWidth      =   795
      TabIndex        =   14
      Top             =   5640
      Width           =   855
   End
   Begin VB.PictureBox Picture10 
      Height          =   1335
      Left            =   9360
      Picture         =   "frmTop.frx":0E61
      ScaleHeight     =   1275
      ScaleWidth      =   915
      TabIndex        =   13
      Top             =   4320
      Width           =   975
   End
   Begin VB.PictureBox Picture9 
      Height          =   1335
      Left            =   6840
      Picture         =   "frmTop.frx":1C36
      ScaleHeight     =   1275
      ScaleWidth      =   795
      TabIndex        =   12
      Top             =   3120
      Width           =   855
   End
   Begin VB.PictureBox Picture8 
      Height          =   1335
      Left            =   4200
      Picture         =   "frmTop.frx":2A98
      ScaleHeight     =   1275
      ScaleWidth      =   915
      TabIndex        =   11
      Top             =   3120
      Width           =   975
   End
   Begin VB.PictureBox Picture7 
      Height          =   1335
      Left            =   9360
      Picture         =   "frmTop.frx":3973
      ScaleHeight     =   1275
      ScaleWidth      =   915
      TabIndex        =   10
      Top             =   3000
      Width           =   975
   End
   Begin VB.PictureBox Picture6 
      Height          =   1335
      Left            =   6840
      Picture         =   "frmTop.frx":4662
      ScaleHeight     =   1275
      ScaleWidth      =   795
      TabIndex        =   9
      Top             =   1800
      Width           =   855
   End
   Begin VB.PictureBox Picture5 
      Height          =   1335
      Left            =   4320
      Picture         =   "frmTop.frx":5458
      ScaleHeight     =   1275
      ScaleWidth      =   795
      TabIndex        =   8
      Top             =   1800
      Width           =   855
   End
   Begin VB.PictureBox Picture4 
      Height          =   1335
      Left            =   9360
      Picture         =   "frmTop.frx":6241
      ScaleHeight     =   1275
      ScaleWidth      =   795
      TabIndex        =   7
      Top             =   1680
      Width           =   855
   End
   Begin VB.PictureBox Picture3 
      Height          =   1335
      Left            =   6720
      Picture         =   "frmTop.frx":6FA6
      ScaleHeight     =   1275
      ScaleWidth      =   915
      TabIndex        =   6
      Top             =   4440
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      Height          =   1335
      Left            =   4200
      Picture         =   "frmTop.frx":7F87
      ScaleHeight     =   1275
      ScaleWidth      =   795
      TabIndex        =   5
      Top             =   4440
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H008080FF&
      Height          =   6615
      Left            =   0
      ScaleHeight     =   6555
      ScaleWidth      =   3315
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.CommandButton cmdGiftShop 
         BackColor       =   &H00FF0000&
         Caption         =   "Gift Shop"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2640
         Width           =   2535
      End
      Begin VB.CommandButton cmdTickets 
         BackColor       =   &H0000C000&
         Caption         =   "Buy Your Tickets Now!"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3960
         Width           =   2775
      End
      Begin VB.CommandButton cmdQuit 
         BackColor       =   &H00800080&
         Caption         =   "Quit "
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5520
         Width           =   2655
      End
      Begin VB.CommandButton cmdTrivia 
         BackColor       =   &H0000FFFF&
         Caption         =   "Trivia Game"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1320
         Width           =   2655
      End
      Begin VB.CommandButton cmdIntro 
         BackColor       =   &H000000FF&
         Caption         =   "Main Page"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Click On Your Favorite Movie Using the Drop Down Box To The Right For Our Personal Research"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   29
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Top 10 Movies Of All Time (Click On A Movie, And Receive A Full Description)"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3720
      TabIndex        =   15
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "frmTop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Disney Land Trivia
'frmAladdin
'Kelly Holmseth and Danny Hansen
'10/28/06
'Objective: The objective of this form is show the user the top ten movies of all time, while allowing them to click on a button to see summaries of the movie that the user chose
Option Explicit

Private Sub cmdAladdin_Click()
frmTop.Hide
frmAladdin.Show

End Sub

Private Sub cmdBambi_Click()
frmTop.Hide
frmBambi.Show

End Sub

Private Sub cmdBeauty_Click()
frmBeauty.Show
frmTop.Hide

End Sub

Private Sub cmdCinderella_Click()
frmTop.Hide
frmCinderella.Show

End Sub

Private Sub cmdDalmations_Click()
frmTop.Hide
frmDalmation.Show

End Sub

Private Sub cmdGiftShop_Click()
frmTrivia.Hide
frmIntro.Hide
frmGiftShop.Show
frmTop.Hide
frmTickets.Hide

End Sub

Private Sub cmdIntro_Click()
frmTrivia.Hide
frmIntro.Show
frmGiftShop.Hide
frmTop.Hide
frmTickets.Hide

End Sub

Private Sub cmdLion_Click()
frmLion.Show
frmTop.Hide

End Sub

Private Sub cmdMermaid_Click()
frmTop.Hide
frmMermaid.Show
'frmTop.Visible = False   if you already have frmTop.Hide and frm Mermaid.Hide you do not need to use the visible command, they do the same thing.
'frmMermaid.Visible = True
End Sub

Private Sub cmdMulan_Click()
frmTop.Hide
frmMulan.Show

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdSleeping_Click()
frmTop.Hide
frmSleeping.Show

End Sub

Private Sub cmdSnow_Click()
frmTop.Hide
frmSnow.Show

End Sub

Private Sub cmdStats_Click()
frmTop.Hide
frmStats.Show
End Sub

Private Sub cmdTickets_Click(Index As Integer)
frmTrivia.Hide
frmIntro.Hide
frmGiftShop.Hide
frmTop.Hide
frmTickets.Show

End Sub

Private Sub cmdTop_Click(Index As Integer)
frmTrivia.Hide
frmIntro.Hide
frmGiftShop.Hide
frmTop.Show
frmTickets.Hide

End Sub

Private Sub cmdTrivia_Click()
frmTrivia.Show
frmIntro.Hide
frmGiftShop.Hide
frmTop.Hide
frmTickets.Hide

End Sub




Private Sub m_Click()
    m.Clear            'this is the code for our drop down menu, kinda cool huh, this is one of our additional codes not learned in class.
    m.AddItem "Bambi", 0
    m.AddItem "Cinderella", 1  'the number corresponds to the position shown on the drop down menu.
    m.AddItem "Lion King", 2
    m.AddItem "The Little Mermaid", 3
    m.AddItem "Beauty and the Beast", 4
    m.AddItem "Mulan", 5
    m.AddItem "Snow White and the Seven Dwarfs", 6
    m.AddItem "101 Dalmations", 7
    m.AddItem "Aladdin", 8
    m.AddItem "Sleeping Beauty", 9
    
End Sub
