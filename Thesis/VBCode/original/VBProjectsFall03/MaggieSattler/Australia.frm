VERSION 5.00
Begin VB.Form Australia1 
   BackColor       =   &H00800000&
   Caption         =   "Form1"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   Picture         =   "Australia.frx":0000
   ScaleHeight     =   8490
   ScaleWidth      =   11655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Restart 
      Caption         =   "See Them Again"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6600
      TabIndex        =   17
      Top             =   7080
      Width           =   1935
   End
   Begin VB.CommandButton okay 
      BackColor       =   &H00E0E0E0&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   16
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox Guess 
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   15
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton Quit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9000
      TabIndex        =   12
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton Alph 
      Caption         =   "View All Cities In Alphabetical Order"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      TabIndex        =   11
      Top             =   7080
      Width           =   2895
   End
   Begin VB.CommandButton View 
      Caption         =   "View All Cities On This Map"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   10
      Top             =   7080
      Width           =   2535
   End
   Begin VB.PictureBox MapCities 
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   3855
      Left            =   7800
      ScaleHeight     =   3795
      ScaleWidth      =   2715
      TabIndex        =   9
      Top             =   2160
      Width           =   2775
   End
   Begin VB.CommandButton Darwin 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Sydney 
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   6
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Cairns 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   5
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton Adelaide 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   4920
      Width           =   255
   End
   Begin VB.CommandButton AliceSprings 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton Hobart 
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   2
      Top             =   6240
      Width           =   255
   End
   Begin VB.CommandButton Perth 
      BackColor       =   &H00E0E0E0&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   4680
      Width           =   255
   End
   Begin VB.CommandButton Broome 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   0
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00800000&
      Caption         =   "Maggie Sattler"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9120
      TabIndex        =   18
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      Caption         =   "Your Guess:"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   8520
      TabIndex        =   14
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Caption         =   "Hello!  Welcome to Australia!  Click on a letter below and find its accompanying city!  Can you guess which city I'm from?"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   1680
      TabIndex        =   13
      Top             =   240
      Width           =   5175
   End
   Begin VB.Image Image2 
      Height          =   1005
      Left            =   360
      Picture         =   "Australia.frx":0342
      Top             =   120
      Width           =   1110
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Cities To Find:"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7920
      TabIndex        =   8
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   5415
      Left            =   720
      Picture         =   "Australia.frx":3E24
      Top             =   1320
      Width           =   6780
   End
End
Attribute VB_Name = "Australia1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'All Aboard Australia (AllAboard.vbp)
'Maggie Sattler
'Date completed: 10/27/03

Option Explicit
Public PATH As String


Private Sub Adelaide_Click()
    'make command button for the particular clicked-on city disappear
    Adelaide.Visible = False
    
    'make message box pop up that tells user which city has been found
    MsgBox "You have found the city of Adelaide!", , "Congratulations!"

    'brings user to correct form
    Australia1.Hide
    Adelaide1.Show
    
End Sub

Private Sub AliceSprings_Click()
    'because the rest of the command buttons with names of Australian cities follow the
    'exact code of the first one, I will refrain from using the same comments
    'over and over
    AliceSprings.Visible = False
    MsgBox "You have found the city of Alice Springs!", , "Congratulations!"
    Australia1.Hide
    AliceSprings1.Show
  
End Sub


Private Sub Alph_Click()

    'Clear the screen of previous information
    MapCities.Cls
        
    'give dimensions to my 8 cities, a temporary storage box, a counter and a variable
    'to keep track of the number of passes to be made
    Dim Cities(1 To 8) As String, Temp As String
    Dim N As Integer
    Dim Pass As Integer
    
    'prepare notepad file to be read
    Open PATH & "AllCities.txt" For Input As #1

    'arrange .txt file into an array
    For N = 1 To 8
        Input #1, Cities(N)
    Next N
    
    'start making passages and, using variable N as counter, go through the array
    'and if the previous city begins with a higher letter in the alphabet than
    'its former, switch the two around.  This should arrange my array into
    'alphabetical order.
    For Pass = 1 To 7
        For N = 1 To 8 - Pass
            If Cities(N) > Cities(N + 1) Then
                Temp = Cities(N)
                Cities(N) = Cities(N + 1)
                Cities(N + 1) = Temp
            End If
        Next N
    Next Pass
    
    'print out the cities in their new alphabetical order
    For N = 1 To 8
        MapCities.Print Cities(N)
    Next N
    
    'close the file
    Close #1
            
    
End Sub

Private Sub Broome_Click()

    Broome.Visible = False
    MsgBox "You have found the city of Broome!", , "Congratulations!"
    Australia1.Hide
    Broome1.Show
End Sub

Private Sub Cairns_Click()

    Cairns.Visible = False
    MsgBox "You have found the city of Cairns!", , "Congratulations!"
    Australia1.Hide
    Cairns1.Show
    
   
End Sub


Private Sub Darwin_Click()
    
    Darwin.Visible = False
    MsgBox "You have found the city of Darwin!", , "Congratulations!"
    Australia1.Hide
    Darwin1.Show
    
End Sub


Private Sub Form_Load()
    PATH = "N:\CS130\handin\Sattler_Maggie\"
End Sub

Private Sub Hobart_Click()

    Hobart.Visible = False
    MsgBox "You have found the city of Hobart!", , "Congratulations!"
    Australia1.Hide
    Hobart1.Show
    
End Sub

Private Sub okay_Click()

    'the user will input a city name into the textbox "Guess"
    'This code allows the computer to read the input and if they have the city spelled
    'correctly, then a message box will pop up with the appropriate message.
    'I've set it up so that the user does not need to capitalize the city name.
    'If the user is incorrect, an appropriate message box will also pop up.
    If Guess.Text = "Sydney" Or Guess.Text = "sydney" Then
            MsgBox "Bloody right you are!  Looks like someone's done their homework!", , "You Are Correct!"
        Else
            MsgBox "No worries, mate!  Go ahead and give it another try!", , "You Are Incorrect!"
            MsgBox "Hint: To learn correct spellings, check out a list of cities by clicking on the button below!", , "HINT"
    End If
    
End Sub

Private Sub Perth_Click()

    Perth.Visible = False
    MsgBox "You have found the city of Perth!", , "Congratulations!"
    Australia1.Hide
    Perth1.Show
End Sub

Private Sub Quit_Click()
    'end the program
    End
End Sub

Private Sub Restart_Click()
    'make all the command buttons visible again so the user may go back to a city
    'if he or she desires to do so.
    'I've set it up so that when the user clicks on a button on the map, the button will
    'disappear so that the user can see which cities he or she has not yet seen.
    'With this command, the user can then go back and review any cities he or she
    'wishes to see again.
    Adelaide.Visible = True
    AliceSprings.Visible = True
    Darwin.Visible = True
    Sydney.Visible = True
    Perth.Visible = True
    Hobart.Visible = True
    Cairns.Visible = True
    Broome.Visible = True
    
End Sub

Private Sub Sydney_Click()

    Sydney.Visible = False
    MsgBox "You have found the city of Sydney!", , "Congratulations!"
    Australia1.Hide
    Sydney1.Show
    
End Sub

Private Sub View_Click()


    'Clear previous information from the picture box
    MapCities.Cls
    
    'prepare file to be read
    Open PATH & "AllCities.txt" For Input As #1
    
    'give dimensions to city names and a counter
    Dim Cities(1 To 8) As String
    Dim N As Integer
    
    'arrange city names into an array
    'and then prints out city names in the picture box
    'This way the user can have a guide that tells him/her which cities to look for
    For N = 1 To 8
        Input #1, Cities(N)
        MapCities.Print Cities(N)
    Next N
    
    'close the file
    Close #1
    
    
End Sub
