VERSION 5.00
Begin VB.Form frmmeetplayers 
   BackColor       =   &H000000C0&
   Caption         =   "Form1"
   ClientHeight    =   7275
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdreturn 
      Caption         =   "return"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   17
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton cmdwennington 
      Caption         =   "Bill Wennington"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   16
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdsimpkins 
      Caption         =   "Dickie Simpkins"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   15
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdsalley 
      Caption         =   "John Salley"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   14
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton cmdrodman 
      Caption         =   "Dennis Rodman"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   13
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton cmdpippen 
      Caption         =   "Scottie Pippen"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton cmdlongley 
      Caption         =   "Luc Longley"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton cmdkukoc 
      Caption         =   "Toni Kukoc"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton cmdkerr 
      Caption         =   "Steve Kerr"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton cmdjordan 
      Caption         =   "Michael Jordan"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton cmdedwards 
      Caption         =   "James Edwards"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton cmdCaffey 
      Caption         =   "Jason Caffey"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   1080
      Width           =   2175
   End
   Begin VB.PictureBox picresults2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6840
      ScaleHeight     =   1155
      ScaleWidth      =   6315
      TabIndex        =   5
      Top             =   4560
      Width           =   6375
   End
   Begin VB.PictureBox picresults 
      Height          =   3735
      Left            =   6840
      ScaleHeight     =   3675
      ScaleWidth      =   6315
      TabIndex        =   4
      Top             =   480
      Width           =   6375
   End
   Begin VB.CommandButton Cmdbuechler 
      Caption         =   "Jud Buechler"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton CmdHarper 
      Caption         =   "Ron Harper"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton cmdbrown 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Randy Brown"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Click Button above to view a picture of each player!"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   600
      TabIndex        =   3
      Top             =   4200
      Width           =   4455
   End
End
Attribute VB_Name = "frmmeetplayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    'Chicago Bulls (Chicagobulls.vbp)
    'frmmeetplayers(frmmeetplayers.frm)
    'Written by: Brian Cullen
    'Written on: March 16, 2008
    'Objective: This form allows the user to meet the Chicago Bulls players by
    'viewing a photo of each player.

    
    
Private Sub cmdbrown_Click()
     'This button loads and prints a picture of brown
     
     'Clears the previous results in the picture box
     picresults.Cls
     picresults2.Cls
     'Loads and prints the picture called brown
     picresults.Picture = LoadPicture(App.Path & "\Brown.jpg")
     picresults2.Print "#0 Randy Brown"
End Sub

Private Sub cmdbuechler_Click()
    'This button loads and prints a picture of buechler
    
    'Clears the previous results in the picture box
    picresults.Cls
    picresults2.Cls
    'Loads and prints the picture called buechler
    picresults.Picture = LoadPicture(App.Path & "\buechler.jpg")
    picresults2.Print "#30 Jud "; "The; Stud"; " Buechler"
End Sub
Private Sub cmdCaffey_Click()
    'This button loads and prints a picture of caffey
    
    'Clears the previous results in the picture box
    picresults.Cls
    picresults2.Cls
    'Loads and prints the picture called
    picresults.Picture = LoadPicture(App.Path & "\caffey.jpg")
    picresults2.Print "#35 Jason Caffey"
End Sub
Private Sub cmdedwards_Click()
    'This button loads and prints a picture of edwards
    
    'Clears the previous results in the picture box
    picresults.Cls
    picresults2.Cls
    'Loads and prints the picture called edwards
    picresults.Picture = LoadPicture(App.Path & "\edwards.jpg")
    picresults2.Print "#53 James Edwards"
End Sub
Private Sub cmdharper_Click()
    'This button loads and prints a picture of Ron harper
    
    'Clears the previous results in the picture box
    picresults.Cls
    picresults2.Cls
    'Loads and prints the picture called Meghan
    picresults.Picture = LoadPicture(App.Path & "\harper.jpg")
    picresults2.Print "#9 Ron Harper"
End Sub
Private Sub cmdjordan_Click()
    'This button loads and prints a picture of jordan
    
    'Clears the previous results in the picture box
    picresults.Cls
    picresults2.Cls
    'Loads and prints the picture called jordan
    picresults.Picture = LoadPicture(App.Path & "\jordan.jpg")
    picresults2.Print "#23 Michael 'Air' Jordan"
End Sub
Private Sub cmdkerr_Click()
    'This button loads and prints a picture of kerr
    
    'Clears the previous results in the picture box
    picresults.Cls
    picresults2.Cls
    'Loads and prints the picture called kerr
    picresults.Picture = LoadPicture(App.Path & "\kerr.jpg")
    picresults2.Print "#25 Steve 'Stevie Wonder' Kerr"
End Sub
Private Sub cmdkukoc_Click()
    'This button loads and prints a picture of toni kukoc
    
    'Clears the previous results in the picture box
    picresults.Cls
    picresults2.Cls
    'Loads and prints the picture called kukoc
    picresults.Picture = LoadPicture(App.Path & "\kukoc.jpg")
    picresults2.Print "#7 Toni Kukoc"
End Sub
Private Sub cmdlongley_Click()
    'This button loads and prints a picture of longley
    
    'Clears the previous results in the picture box
    picresults.Cls
    picresults2.Cls
    'Loads and prints the picture called longley
    picresults.Picture = LoadPicture(App.Path & "\longley.jpg")
    picresults2.Print "#13 Luc  'The Big Aussie'  Longley"
End Sub
Private Sub cmdpippen_Click()
    'This button loads and prints a picture of scottie pippen
    
    'Clears the previous results in the picture box
    picresults.Cls
    picresults2.Cls
    'Loads and prints the picture called pippen
    picresults.Picture = LoadPicture(App.Path & "\pippen.jpg")
    picresults2.Print "#33 Scottie Pippen"
End Sub



Private Sub cmdrodman_Click()
    'This button loads and prints a picture of denny rodman
    
    'Clears the previous results in the picture box
    picresults.Cls
    picresults2.Cls
    'Loads and prints the picture called rodman
    picresults.Picture = LoadPicture(App.Path & "\rodman.jpg")
    picresults2.Print "#91 Dennis  The Worm Rodman"
End Sub
Private Sub cmdsalley_Click()
    'This button loads and prints a picture of salley
    
    'Clears the previous results in the picture box
    picresults.Cls
    picresults2.Cls
    'Loads and prints the picture called salley
    picresults.Picture = LoadPicture(App.Path & "\salley.jpg")
    picresults2.Print "#22 John Salley"
End Sub
Private Sub cmdsimpkins_Click()
    'This button loads and prints a picture of simpkins
    
    'Clears the previous results in the picture box
    picresults.Cls
    picresults2.Cls
    'Loads and prints the picture called simpkins
    picresults.Picture = LoadPicture(App.Path & "\simpkins.jpg")
    picresults2.Print "#8 Dickie Simpkins"
End Sub
Private Sub cmdwennington_Click()
    'This button loads and prints a picture of bill wennington
    
    'Clears the previous results in the picture box
    picresults.Cls
    picresults2.Cls
    'Loads and prints the picture called wennington
    picresults.Picture = LoadPicture(App.Path & "\wennington.jpg")
    picresults2.Print "#34 Bill Wennington"
End Sub
Private Sub cmdquit_Click()
    'This stops the program
    End
End Sub
End Sub
Private Sub cmdreturn_Click()
frmmainpage.Show
frmmeetplayers.Hide

End Sub


