VERSION 5.00
Begin VB.Form frmpetform 
   BackColor       =   &H80000013&
   Caption         =   "Form1"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   Picture         =   "ChoosePet.frx":0000
   ScaleHeight     =   7890
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGoStore 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Go to Main Page"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5880
      Width           =   2655
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6480
      Width           =   735
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFC0C0&
      Height          =   3615
      Left            =   7080
      ScaleHeight     =   3555
      ScaleWidth      =   3195
      TabIndex        =   12
      Top             =   2040
      Width           =   3255
   End
   Begin VB.PictureBox Picture4 
      Height          =   2175
      Left            =   3720
      Picture         =   "ChoosePet.frx":181C52
      ScaleHeight     =   2115
      ScaleWidth      =   3195
      TabIndex        =   11
      Top             =   4680
      Width           =   3255
   End
   Begin VB.PictureBox Picture3 
      Height          =   1935
      Left            =   3720
      Picture         =   "ChoosePet.frx":19DC14
      ScaleHeight     =   1875
      ScaleWidth      =   3195
      TabIndex        =   10
      Top             =   1800
      Width           =   3255
   End
   Begin VB.PictureBox Picture2 
      Height          =   2175
      Left            =   120
      Picture         =   "ChoosePet.frx":1B3CD6
      ScaleHeight     =   2115
      ScaleWidth      =   3195
      TabIndex        =   9
      Top             =   4680
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   120
      Picture         =   "ChoosePet.frx":1CB37C
      ScaleHeight     =   1875
      ScaleWidth      =   3195
      TabIndex        =   8
      Top             =   1800
      Width           =   3255
   End
   Begin VB.CommandButton cmdInfoTurtle 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Find out more about the Turtle!"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton cmdBuyTurtle 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Purchase Turtle"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdInfoDog 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Find out more about the Dog!"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdBuyDog 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Purchase Dog"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton cmdInfoFish 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Find out more about the Fish!"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdBuyFish 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Purchase Fish"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdInfoKitten 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Find out more about the Kitten!"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdBuyCat 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Purchase Kitten"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label lblChoose 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Your Pet!"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   1575
      Left            =   360
      TabIndex        =   15
      Top             =   0
      Width           =   9495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      BackStyle       =   0  'Transparent
      Caption         =   "Useful Information about your new pet!"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   555
      Left            =   6360
      TabIndex        =   13
      Top             =   1440
      Width           =   4275
   End
End
Attribute VB_Name = "frmpetform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'on this form you can learn more about the pets you can
'choose from and then also purchase any of the pets.

Private Sub Command1_Click()

End Sub



Private Sub cmdBuyCat_Click()
frmpetform.Hide
frmKittenForm.Show
End Sub

Private Sub cmdBuyDog_Click()
frmpetform.Hide
frmDogform.Show
End Sub

Private Sub cmdBuyFish_Click()
frmpetform.Hide
frmFishform.Show
End Sub

Private Sub cmdBuyTurtle_Click()
frmpetform.Hide
frmTurtleform.Show
End Sub

Private Sub cmdGoStore_Click()

frmpetform.Hide
Welcomeform2.Show

End Sub

Private Sub cmdInfoDog_Click()

picresults.Cls

picresults.Print "The dog is a very friendly and people oriented"
picresults.Print "pet. He or she will love to play with people"
picresults.Print "and most of all play fetch. Dogs love to go"
picresults.Print "for walks in the fresh air and also ride in"
picresults.Print "the car with the windows open. Also, if your"
picresults.Print "dog is a puppy, he or she may need to be potty"
picresults.Print "trained. Dogs are typically fed twice a day."
picresults.Print "Once in the morning and once in the evening."
picresults.Print "Dogs also need to be groomed at least twice"
picresults.Print "a month. This includes a bath followed by the"
picresults.Print "brushing of their fur. Some types of dogs may"
picresults.Print "require more than this."
picresults.Print "Good luck with your new Dog!"

End Sub

Private Sub cmdInfoFish_Click()

picresults.Cls

picresults.Print "Your new fish must be kept in water at all"
picresults.Print "times. It prefers a tank with an open top"
picresults.Print "so that breathing is easier. Your new fish"
picresults.Print "loves to swim in and out of many things so"
picresults.Print "it will love you even more if you provide many"
picresults.Print "toys in its new tank. You can purchase this"
picresults.Print "at our store! Also, your new fish will need"
picresults.Print "to be fed twice everyday. Once in the morn-"
picresults.Print "ing and once at night before sleep. The fish"
picresults.Print "is also a very friendly pet."
picresults.Print "Good luck with your new fish!"

End Sub

Private Sub cmdInfoKitten_Click()

picresults.Cls

picresults.Print "The kitten is a very loveable and fluffy pet."
picresults.Print "Your new kitten will love to play tug-o-war"
picresults.Print "and cuddle with its new owner.The kitten will"
picresults.Print "need a litter box in order to keep your house"
picresults.Print "clean and fresh. He or she will also need lots"
picresults.Print "of toys to play with and keep him or her busy."
picresults.Print "Kittens eat at their own pleasure and there-"
picresults.Print "fore are satisfied if you keep their dishes"
picresults.Print "of food and water full at all times."
picresults.Print "Good luck with your new kitten!"

End Sub

Private Sub cmdInfoTurtle_Click()

picresults.Cls

picresults.Print "The turtle is a very hassle-free pet. If you"
picresults.Print "do not want to mess with litter and fur every-"
picresults.Print "where, then the turtle is the pet for you."
picresults.Print "Turtles are also a very interesting pet and"
picresults.Print "can be enjoyed by just watching the tank."
picresults.Print "A turtle will require a moderately large tank"
picresults.Print "with room to swim and play. Turtles will feed"
picresults.Print "on the vegetation in the tank and do not re-"
picresults.Print "quire daily feeding, although this may be"
picresults.Print "given as a treat. "
picresults.Print "Good luck with your new turtle!"

End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub lblChoose_Click()

End Sub
