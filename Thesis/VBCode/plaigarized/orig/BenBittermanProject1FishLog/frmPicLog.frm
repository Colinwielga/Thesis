VERSION 5.00
Begin VB.Form frmPicLog 
   Caption         =   "Form2"
   ClientHeight    =   13020
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   18960
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   Picture         =   "frmPicLog.frx":0000
   ScaleHeight     =   13020
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPicName 
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      TabIndex        =   10
      Top             =   10920
      Width           =   4695
   End
   Begin VB.CommandButton cmdSwitch2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Switch Back to  Fishing Log"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   14760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5040
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   14760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7680
      Width           =   2175
   End
   Begin VB.CommandButton cmdPicSelect 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Show Selected Picture"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7080
      Width           =   2655
   End
   Begin VB.TextBox txtPicSelect 
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      TabIndex        =   5
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox txtPicNumber 
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11400
      TabIndex        =   2
      Top             =   9720
      Width           =   975
   End
   Begin VB.PictureBox picSlideShow 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   6840
      ScaleHeight     =   8115
      ScaleWidth      =   6555
      TabIndex        =   1
      Top             =   1440
      Width           =   6615
   End
   Begin VB.CommandButton cmdSlideShow 
      BackColor       =   &H00C0FFC0&
      Caption         =   "View Slide Show   of Ben's Fish"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label lblPicName2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "^ PicName/ CatchDate and Fish Type ^"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   11
      Top             =   11520
      Width           =   4695
   End
   Begin VB.Label lblPictureLog 
      BackStyle       =   0  'Transparent
      Caption         =   "Fishing Pic Log"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   6720
      TabIndex        =   9
      Top             =   0
      Width           =   10215
   End
   Begin VB.Label lblQuestion 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Which picture do you want to view?"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1920
      TabIndex        =   4
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Label lblPicNumber 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Picture Number ---->"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      TabIndex        =   3
      Top             =   9840
      Width           =   3255
   End
End
Attribute VB_Name = "frmPicLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdPicSelect_Click()
'Declare variables
Dim PicSelect As Integer, Ctr As Integer
Dim PicNames(1 To 50) As String, Found As Boolean, PicNumber As Integer
Dim I As Integer

'Open file PicNames
Open App.Path & "\PicNames.txt" For Input As #1
Ctr = 0

'Store names into an array
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, PicNames(Ctr)
Loop
Close #1

'Set data entered in text box PicSelect equal to variable PicSelect
PicSelect = txtPicSelect.Text
I = 0
'set found = to false for match/stop search
Found = False
    
    'Initiate match/stop search using a do while loop
    Do While ((Not Found) And (I < Ctr))
    I = I + 1
        If PicSelect = I Then
            'When data entered into text box is equal to counter I
            Found = True
            'Load up the picture that corresponds with I in the array
            picSlideShow.Picture = LoadPicture(App.Path & "\Cs130FishLog\" & PicNames(I))
            'Display number from PicSelect in the other txt box as well so user is not confused
            txtPicNumber.Text = PicSelect
            txtPicName.Text = PicNames(I)
        End If
    Loop
    
    If (Not Found) Then
        'If data put in inputbox is not found a message box will notify the user that no picture was listed to that number
        MsgBox ("There was no picture listed under that number")
    End If

End Sub

Private Sub cmdQuit2_Click()
    End
End Sub

Private Sub cmdSlideShow_Click()
'Declare variables: stopper is a type of counter that is used during slide show
Dim whichfish As Integer, stopper As Integer, t As Double, oldfish As Integer, ctr2 As Double
'Declare new array as PicNames for the names of pictures contained in folder "Cs130FishLog" and file PicNames
Dim PicNames(1 To 50) As String, Ctr As Integer

'open up the file
Open App.Path & "\PicNames.txt" For Input As #1
Ctr = 0

'do while loop used to store the names of the pictures into an array
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, PicNames(Ctr)
Loop
Close #1

'Start with picture 1 so whichfish = 1
whichfish = 1
stopper = 0

'Slide show format ideas were sythesized from example "N:\Classes\CS130\VB_Examples\Multi_form_Sample_w_pictures_and_Module"
'Use a nested Do While loop to runthe slide show
'First cotinue loop while stopper is less than counter
Do While (stopper < Ctr)
    'display whichfish in the text box as output for each new picture that is shown
    txtPicNumber.Text = whichfish
    txtPicName.Text = PicNames(whichfish)
    'Load up picture from folder Cs130FishLog and displat it in the SlideShow picture box according to the PicNames array
    picSlideShow.Picture = LoadPicture(App.Path & "\Cs130FishLog\" & PicNames(whichfish))
    t = Timer
    
    'set the timing of the slide show
    Do While (Timer - t) < 2
        'DoEvents is necissary for displaying Whichfish in the text box called txtPicNumber
        'John Miller helped me discover that DoEvents was necissary within this loop to display the output counter in the textbox
        DoEvents
        
    Loop
    
    'count up the stopper
    stopper = stopper + 1
    'oldfish becomes wichfish
    oldfish = whichfish
    'set whichfish to stopper "modulus arithmetic" + 1
    whichfish = (stopper Mod Ctr) + 1
    
Loop

    
End Sub



Private Sub cmdSwitch2_Click()
'Switches program back to first form
frmPicLog.Hide
frmFishLog.Show
End Sub

Private Sub Form_Load()
    'When the program starts on form PicLog a journal cover picture diplaying types of freshwater fish is printed in the picture box
    ' Picture "JournalCover" was taken from google images after searching Freshwater fish
    'http://www.google.com/imgres?imgurl=http://www.gold-fish.us/upl/Image/fishes.jpg&imgrefurl=http://www.gold-fish.us/type-fish/fresh-water-fish-species.html&usg=__KarbjTlVGzUGLpCMmtjAu25UZrk=&h=640&w=480&sz=48&hl=en&start=0&zoom=1&tbnid=hACyddvT3LQwEM:&tbnh=141&tbnw=106&prev=/images%3Fq%3Dfreshwater%2Bfish%26hl%3Den%26biw%3D1676%26bih%3D837%26gbv%3D2%26tbs%3Disch:1&itbs=1&iact=hc&vpx=1285&vpy=186&dur=733&hovh=259&hovw=194&tx=110&ty=143&ei=VwrFTLzzGIymnAeCofXXCQ&oei=VwrFTLzzGIymnAeCofXXCQ&esq=1&page=1&ndsp=34&ved=1t:429,r:7,s:0
    picSlideShow.Picture = LoadPicture(App.Path & "\JournalCover.jpg")
End Sub

