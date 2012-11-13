VERSION 5.00
Begin VB.Form frmCharacters 
   BackColor       =   &H00800000&
   Caption         =   "FormCharacters"
   ClientHeight    =   10770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16290
   LinkTopic       =   "Form1"
   ScaleHeight     =   10770
   ScaleWidth      =   16290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdpickflava 
      Caption         =   "Pick Me!"
      Height          =   615
      Left            =   10320
      TabIndex        =   16
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdbioflava 
      Caption         =   "Bio"
      Height          =   615
      Left            =   10320
      TabIndex        =   15
      Top             =   3960
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   1935
      Left            =   8400
      ScaleHeight     =   1875
      ScaleWidth      =   1755
      TabIndex        =   14
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton cmdbiofox 
      Caption         =   "Bio"
      Height          =   615
      Left            =   7080
      TabIndex        =   13
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmdpickfox 
      Caption         =   "Pick Me!"
      Height          =   615
      Left            =   7080
      TabIndex        =   12
      Top             =   7680
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      Height          =   1695
      Left            =   4680
      ScaleHeight     =   1635
      ScaleWidth      =   2235
      TabIndex        =   11
      Top             =   6720
      Width           =   2295
   End
   Begin VB.CommandButton cmdbioalba 
      Caption         =   "Bio"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdpickalba 
      Caption         =   "Pick Me!"
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   5640
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   2295
      Left            =   1800
      ScaleHeight     =   2235
      ScaleWidth      =   1755
      TabIndex        =   8
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton cmdpickosama 
      Caption         =   "Pick Me!"
      Height          =   615
      Left            =   10440
      TabIndex        =   7
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdosamabio 
      Caption         =   "Bio"
      Height          =   615
      Left            =   10440
      TabIndex        =   6
      Top             =   720
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   8280
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton cmdpicktrump 
      Caption         =   "Pick Me!"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   3480
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Height          =   2895
      Left            =   360
      ScaleHeight     =   2835
      ScaleWidth      =   2715
      TabIndex        =   3
      Top             =   360
      Width           =   2775
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FFFF&
      Height          =   2895
      Left            =   3960
      ScaleHeight     =   2895
      ScaleWidth      =   4095
      TabIndex        =   2
      Top             =   2280
      Width           =   4095
   End
   Begin VB.CommandButton cmdtrumpbio 
      Caption         =   "Bio"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Text            =   "Which contestant would you like to be?"
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "frmCharacters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdbioalba_Click()
picresults.Cls
picresults.Print "                 Jessica Alba      "
picresults.Print "********************************************************************************************************"
picresults.Print "I'm the hottest girl ever, thats all you"
picresults.Print "need to know about me"
End Sub

Private Sub cmdbioflava_Click()
picresults.Cls
picresults.Print "                  Flava Flav     "
picresults.Print "***************************************************************************************************"
picresults.Print "YEAH BOOOOOIIIII!! Flava Flav is the name and"
picresults.Print "I'm the smartest man alive! I've already got"
picresults.Print "a million dollars, but I always want to make more"
picresults.Print "money...."
picresults.Print "..acutally, the second season of my show bombed"
picresults.Print "and I'm dead broke. I can't even afford a new clock"
picresults.Print "to wear around town. Please, help me get rich again!"

End Sub

Private Sub cmdbiofox_Click()
picresults.Cls
picresults.Print "                  Megan Fox"
picresults.Print "**************************************************************************************************"
picresults.Print "My name is Megan Fox and I'm an actress. You"
picresults.Print "might know me from the recent hit movie"
picresults.Print "Transformers, where I play a ditzy but hot"
picresults.Print "girl who falls in love with a loser because"
picresults.Print "he drives a cool car. In my spare time I like"
picresults.Print "to read comic books, skydive, and collect belly"
picresults.Print "button lint. Pick me and I will totally flirt"
picresults.Print "with the host so he will give me the"
picresults.Print "million dollars!"



End Sub

Private Sub cmdosamabio_Click()
picresults.Cls
picresults.Print "              Osama Bin Laden"
picresults.Print "****************************************************************************************"
picresults.Print "AAAAHH! Why would you stupid Americans ever"
picresults.Print "pick me to do a game show? Don't you know I"
picresults.Print "live in a cave? Have you ever seen a television"
picresults.Print "inside a cave? I DON'T EVEN HAVE RUNNING"
picresults.Print "WATER YOU FOOLS!!! You would have to be an"
picresults.Print "idiot to think I would be good at this game!"

End Sub

Private Sub cmdpickalba_Click()
frmCharacters.Hide
FrmQuestion1.Show
End Sub

Private Sub cmdpickflava_Click()
frmCharacters.Hide
FrmQuestion1.Show

End Sub

Private Sub cmdpickfox_Click()
frmCharacters.Hide
FrmQuestion1.Show
End Sub

Private Sub cmdpickosama_Click()
frmCharacters.Hide
FrmQuestion1.Show
End Sub

Private Sub cmdpicktrump_Click()
frmCharacters.Hide
FrmQuestion1.Show
End Sub

Private Sub cmdtrumpbio_Click()
    picresults.Cls
    picresults.Print "               Donald Trump"
    picresults.Print "***************************************************************************************"
    picresults.Print "Hi, my name is Don Trump. I was"
    picresults.Print "Born on June 14th, 1946 and am"
    picresults.Print "currently Chairmen and CEO of the"
    picresults.Print "Trump Organization.  In my spare time,"
    picresults.Print "I like to drink champagne, take long"
    picresults.Print "walks at sunset, and shop for new hair pieces."
    picresults.Print "Pick me and you won't be dissapointed!"
 
End Sub

Private Sub Form_Load()
'FiftyEnabled = True
'PhoneEnabled = True
'AudienceEnabled = True
End Sub

