VERSION 5.00
Begin VB.Form frmNature 
   BackColor       =   &H00008000&
   Caption         =   "Nature"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   11340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdListenfrog 
      Caption         =   "Listen"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   6
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmdListenowl 
      Caption         =   "Listen"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmdListenbird 
      Caption         =   "Listen"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   4
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdListenwolf 
      Caption         =   "Listen"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdLoadPic 
      Caption         =   "Load the images."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1635
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Image Image4 
      Height          =   1575
      Left            =   7920
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Image Image3 
      Height          =   1815
      Left            =   3960
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   1455
      Left            =   7920
      Top             =   0
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   3840
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmNature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Music VB Project by Cassiann Procenko
'Form Name is frmNature
'Date written 10/16/2009
'Purpose of this form is to show viewers that animals are musical just like people and inform them about the animals, and also share with the viewer the each animal's song.

'this code is taken from the online Visual Basic Manual in the forum.  http://www.microsoft.com.  On the page, it states the code came from John Walkenbach and his book titled, "Excel 2000 Power Programming."
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Dim returnval As Long
Dim soundfile As String

Private Sub cmdListenbird_Click()
'play sound file
soundfile = App.Path & "canary.wav"
returnval = PlaySound("canary.wav", 0, &H0)
End Sub

Private Sub cmdListenfrog_Click()
'play sound file
soundfile = App.Path & "frog.wav"
returnval = PlaySound("frog.wav", 0, &H0)
End Sub

Private Sub cmdListenowl_Click()
'play sound file
soundfile = App.Path & "owl.wav"
returnval = PlaySound("owl.wav", 0, &H0)
End Sub

Private Sub cmdListenwolf_Click()
'play sound file
soundfile = App.Path & "howl.wav"
returnval = PlaySound("howl.wav", 0, &H0)
End Sub

Private Sub cmdLoadPic_Click()
'load the images of the animals
Image1 = LoadPicture(App.Path & "\howl1.jpg")
Image2 = LoadPicture(App.Path & "\songbird1.jpg")
Image3 = LoadPicture(App.Path & "\owl.jpg")
Image4 = LoadPicture(App.Path & "\frog.jpg")
MsgBox "Please click on the image that represents the animal with the most beautiful song, in your opinion."
End Sub

Private Sub cmdQuit_Click()
'show and hide forms
    frmLeave.Show
    frmNature.Hide
End Sub

Private Sub cmdReturn_Click()
'show and hide forms
    frmMusicTypes.Show
    frmNature.Hide
End Sub


Private Sub Image1_Click()
'information on the gray wolf
'Works cited:  information taken and sound clip from http://www.wolfsongalaska.org/.  wolf picture taken from www.nationalgeographic.com
MsgBox "You have chosen the gray wolf.  Howling is a common way gray wolves communicate with each other when they are apart, such as when another wolf is lost.", , "Gray Wolf"
End Sub


Private Sub Image2_Click()
'information on the canary
'Works cited:  information and sound clip taken from http://www.birds.cornell.edu/.  bird picture taken from www.nationalgeographic.com
MsgBox "You have chosen a bird called a canary.  This is a bird in the finch family.  Most birds sing songs, and the canary is no different, although some people say it has the prettiest song.  Birds sing to proclaim their territory and to win over mates.", , "Canary"
End Sub

Private Sub Image3_Click()
'information on the owl
'works cited:  information and sound clip taken from http://www.owlpages.com/. owl picture taken from www.nationalgeographic.com
MsgBox "You have chosen the Great Horned Owl.  These birds hoot as a form of communication with other birds.", , "Great Horned Owl"
End Sub

Private Sub Image4_Click()
'information on the frog
'works cited: information and sound clip taken from http://allaboutfrogs.org/.  frog picture taken from www.nationalgeographic.com
MsgBox "You have chosen a Frog.  Frogs sing their songs for a variety of different reasons including calling for a mate, sensing a change in weather, communication, and also marking their territory.", , "Frog"
End Sub
