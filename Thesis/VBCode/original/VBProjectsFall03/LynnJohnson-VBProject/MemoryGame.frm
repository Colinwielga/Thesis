VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox card1 
      Height          =   855
      Left            =   480
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   21
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox SecondNumber 
      Height          =   615
      Left            =   2280
      TabIndex        =   19
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox FirstNumber 
      Height          =   615
      Left            =   600
      TabIndex        =   18
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdplay 
      Caption         =   "Play"
      Height          =   1095
      Left            =   7680
      TabIndex        =   17
      Top             =   360
      Width           =   1215
   End
   Begin VB.PictureBox pbxresults 
      Height          =   1095
      Left            =   7680
      ScaleHeight     =   1035
      ScaleWidth      =   1275
      TabIndex        =   16
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdshuffle 
      Caption         =   "Shuffle Cards"
      Height          =   495
      Left            =   7680
      TabIndex        =   15
      Top             =   2280
      Width           =   1215
   End
   Begin VB.PictureBox card116 
      Height          =   855
      Left            =   5760
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   14
      Top             =   5400
      Width           =   1335
   End
   Begin VB.PictureBox card15 
      Height          =   855
      Left            =   3960
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   13
      Top             =   5400
      Width           =   1335
   End
   Begin VB.PictureBox card14 
      Height          =   855
      Left            =   2280
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   12
      Top             =   5400
      Width           =   1335
   End
   Begin VB.PictureBox card13 
      Height          =   855
      Left            =   480
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   11
      Top             =   5400
      Width           =   1335
   End
   Begin VB.PictureBox card12 
      Height          =   855
      Left            =   5760
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   10
      Top             =   4080
      Width           =   1335
   End
   Begin VB.PictureBox card11 
      Height          =   855
      Left            =   3960
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   9
      Top             =   4080
      Width           =   1335
   End
   Begin VB.PictureBox card10 
      Height          =   855
      Left            =   2280
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   8
      Top             =   4080
      Width           =   1335
   End
   Begin VB.PictureBox card9 
      Height          =   855
      Left            =   480
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   7
      Top             =   4080
      Width           =   1335
   End
   Begin VB.PictureBox card8 
      Height          =   855
      Left            =   5760
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   6
      Top             =   2880
      Width           =   1335
   End
   Begin VB.PictureBox card7 
      Height          =   855
      Left            =   3960
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   5
      Top             =   2880
      Width           =   1335
   End
   Begin VB.PictureBox card6 
      Height          =   855
      Left            =   2280
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   4
      Top             =   2880
      Width           =   1335
   End
   Begin VB.PictureBox card5 
      Height          =   855
      Left            =   480
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   3
      Top             =   2880
      Width           =   1335
   End
   Begin VB.PictureBox card4 
      Height          =   855
      Left            =   5760
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.PictureBox card3 
      Height          =   855
      Left            =   3960
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   1
      Top             =   1680
      Width           =   1335
   End
   Begin VB.PictureBox card2 
      Height          =   855
      Left            =   2280
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   0
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblone 
      Caption         =   "1"
      Height          =   255
      Left            =   600
      TabIndex        =   22
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Select two numbers"
      Height          =   255
      Left            =   1200
      TabIndex        =   20
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdshuffle_Click()
    
    Randomize (i)
    
    Dim i As Integer
    For i = 1 To 8
        j = Rnd(i)
    Next i
        pbxresults.Print j
End Sub

Private Sub card1_Click()
    card1.Print 1
    
End Sub

Private Sub cmdplay_Click()
    
    Dim i As Integer
   
    
    Do While Not game
        
        one = First
    

 
Select Case pics
    Case Is = 1
        Load "M:\CS130\VB Project\Images\hummingbird.jpg"
    Case Is = 2
        Load "M:\CS130\VB Project\Images\Bora Bora Palm.jpg"
    Case Is = 3
        Load "M:\CS130\VB Project\Images\Giant Panda.jpg"
    Case Is = 4
        Load "M:\CS130\VB Project\Images\Cropped Panda.jpg"
End Select

End Sub

Private Sub Picture1_Click()
    Dim icount As Integer
    'add to icount'
End Sub
