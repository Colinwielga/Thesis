VERSION 5.00
Begin VB.Form frmPictures 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Pictures from Mission Trips (Megan Fitzgerald)"
   ClientHeight    =   7035
   ClientLeft      =   3630
   ClientTop       =   2790
   ClientWidth     =   9870
   BeginProperty Font 
      Name            =   "Century Schoolbook"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   9870
   Begin VB.TextBox txtMoreInfo 
      Height          =   615
      Left            =   1320
      TabIndex        =   12
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton cmdNew10 
      BackColor       =   &H00FF8080&
      Caption         =   "Picture 10"
      Height          =   495
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5640
      Width           =   1095
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF8080&
      Height          =   5415
      Left            =   120
      ScaleHeight     =   5415
      ScaleWidth      =   8415
      TabIndex        =   10
      Top             =   120
      Width           =   8415
   End
   Begin VB.CommandButton cmdNew9 
      BackColor       =   &H00FF8080&
      Caption         =   "Picture 9"
      Height          =   495
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew8 
      BackColor       =   &H00FF8080&
      Caption         =   "Picture 8"
      Height          =   495
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew7 
      BackColor       =   &H00FF8080&
      Caption         =   "Picture 7"
      Height          =   495
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew6 
      BackColor       =   &H00FF8080&
      Caption         =   "Picture 6"
      Height          =   495
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew5 
      BackColor       =   &H00FF8080&
      Caption         =   "Picture 5"
      Height          =   495
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew4 
      BackColor       =   &H00FF8080&
      Caption         =   "Picture 4"
      Height          =   495
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew3 
      BackColor       =   &H00FF8080&
      Caption         =   "Picture 3"
      Height          =   495
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew2 
      BackColor       =   &H00FF8080&
      Caption         =   "Picture 2"
      Height          =   495
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew1 
      BackColor       =   &H00FF8080&
      Caption         =   "Picture 1"
      Height          =   495
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FF8080&
      Caption         =   "Return to Mission Trips"
      Height          =   615
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   $"Pictures.frx":0000
      Height          =   1095
      Left            =   360
      TabIndex        =   13
      Top             =   5640
      Width           =   6135
   End
End
Attribute VB_Name = "frmPictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectAmigosForChrist (Megan Fitzgerald's Visual Basic Project.vbp)
'Form Name : frmPictures (Pictures.frm)
'Author: Megan Fitzgerald
'Date Written: October 28, 2003
'Purpose of This Form: To allow the user to interact with this interface and
                        'see pictures from previous mission trips to Nicaragua.
                        'The user will also be given the option of learning more
                        'about each picture by having messages pop up when the
                        'button for a particular picture is clicked.

Option Explicit

Dim PATH As String
Dim Answer As String, Answer2 As String
'For each of the various buttons labeled "New" a picture will be loaded
'onto the form each time a button is clicked.  The picResults box will
'be cleared at the beginning of each command sequence.  The series of
'"If...then" statements refer to the option of the user to have message boxes
'appear each time a picture button is clicked.

Private Sub cmdNew1_Click()
picResults.Cls
picResults.Picture = LoadPicture(PATH & "CuteKids.jpg")
If txtMoreInfo.Text = Answer Or txtMoreInfo.Text = Answer2 Then MsgBox "Even though these kids live in such impoverished conditions, they always seem to find a reason to smile.", , "Adorable Kids"
End Sub


Private Sub cmdNew2_Click()
picResults.Cls
picResults.Picture = LoadPicture(PATH & "Buckets on Heads.jpg")
If txtMoreInfo.Text = Answer Or txtMoreInfo.Text = Answer2 Then MsgBox "During each mission trip, we work on various projects in the villages, such as this one where we are constructing homes.", , "Projects"
End Sub

Private Sub cmdNew3_Click()
picResults.Cls
picResults.Picture = LoadPicture(PATH & "Kids at Church.jpg")
If txtMoreInfo.Text = Answer Or txtMoreInfo.Text = Answer2 Then MsgBox "This picture was taken outside of the Catholic Church that we attend while we are in Nicaragua.", , "Church"
End Sub

Private Sub cmdNew4_Click()
picResults.Cls
picResults.Picture = LoadPicture(PATH & "Home that We Built.jpg")
If txtMoreInfo.Text = Answer Or txtMoreInfo.Text = Answer2 Then MsgBox "This picture gives you an idea of what the homes we build look like. Although this home may look primitive and small to you, it is a mansion in the eyes of someone who has previously been living under a garbage bag held up by sticks.", , "Homes"
End Sub

Private Sub cmdNew5_Click()
picResults.Cls
picResults.Picture = LoadPicture(PATH & "Reading to Nicaraguan Kids.jpg")
If txtMoreInfo.Text = Answer Or txtMoreInfo.Text = Answer2 Then MsgBox "The Nicaraguan children soak up the attention that we give them.  In this picture, one of the high school students on the trip is reading to some of the Nicaraguan children.", , "Reading Time"
End Sub

Private Sub cmdNew6_Click()
picResults.Cls
picResults.Picture = LoadPicture(PATH & "Garbage Dump.jpg")
If txtMoreInfo.Text = Answer Or txtMoreInfo.Text = Answer2 Then MsgBox "After Hurrican Mitch destroyed their village, the government put these people in a village right next to this garbage dump.  The people here are some of the poorest in the country and they survive off of the garbage that comes into the dump.  This place is infested with disease and therefore Amigos for Christ has been working to move these Nicaraguans to a new village a few miles away.", , "Garbage Dump"
End Sub

Private Sub cmdNew7_Click()
picResults.Cls
picResults.Picture = LoadPicture(PATH & "Playing Games.jpg")
If txtMoreInfo.Text = Answer Or txtMoreInfo.Text = Answer2 Then MsgBox "The Nicaraguan children love to play the games that we teach them and throughout each mission trip strong friendships are formed between us and the people of Nicaragua.  It is usually a tear-filled goodbye at the end of the trip.", , "Playing Games"
End Sub

Private Sub cmdNew8_Click()
picResults.Cls
picResults.Picture = LoadPicture(PATH & "Traditional Dance.jpg")
If txtMoreInfo.Text = Answer Or txtMoreInfo.Text = Answer2 Then MsgBox "This picture was taken at one of the schools that Amigos for Christ helps to fund. The children in this picture were putting on a welcome ceremony for us in which some of them danced in traditional costumes.", , "Traditional Dance"
End Sub

Private Sub cmdNew9_Click()
picResults.Cls
picResults.Picture = LoadPicture(PATH & "Dedication Ceremony.jpg")
If txtMoreInfo.Text = Answer Or txtMoreInfo.Text = Answer2 Then MsgBox "After we finish building the homes in a village, our next goal is to build the school so that the children will have an opportunity to learn and hopefully one day they will be able to escape the vicious cycle of poverty. This picture was taken during the dedication ceremony at one of these schools.", , "Dedication Ceremony"
End Sub
Private Sub cmdNew10_Click()
picResults.Cls
picResults.Picture = LoadPicture(PATH & "Hard Days Work.jpg")
If txtMoreInfo.Text = Answer Or txtMoreInfo.Text = Answer2 Then MsgBox "Although the work is exhausting (as this picture reveals), it is such a rewarding feeling to know that you have been able to serve God by serving His poor.", , "Sweet Rewards"
End Sub
Private Sub cmdReturn_Click()

'Takes the user back to Homepage "Amigos for Christ".
frmPictures.Hide
frmMissionTrips.Show

End Sub

Private Sub Form_Load()
PATH = "N:\CS130\handin\Fitzgerald_Megan\Pictures\"
Answer = "Yes"
Answer2 = "yes"
End Sub
