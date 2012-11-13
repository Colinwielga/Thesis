VERSION 5.00
Begin VB.Form frmCharacters 
   BackColor       =   &H00000080&
   Caption         =   "Characters"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picNames 
      BackColor       =   &H00FFFFFF&
      Height          =   7815
      Left            =   3120
      ScaleHeight     =   7755
      ScaleWidth      =   2115
      TabIndex        =   9
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton cmdPicture 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click to see a picture of this character"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8040
      Width           =   2175
   End
   Begin VB.CommandButton cmdFacts 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click to learn facts about this character"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7200
      Width           =   2175
   End
   Begin VB.TextBox txtCharacterName 
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   6240
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   7815
      Left            =   5400
      ScaleHeight     =   7755
      ScaleWidth      =   5475
      TabIndex        =   2
      Top             =   240
      Width           =   5535
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click to display character names in alphabetical order"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      MaskColor       =   &H00000080&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8160
      Width           =   1935
   End
   Begin VB.Label lblCharacterFacts 
      BackColor       =   &H00000080&
      Caption         =   "Character Facts:"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   360
      TabIndex        =   10
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label lblCharacters 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmCharacters.frx":0000
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label lblRequest 
      BackColor       =   &H00000080&
      Caption         =   "Please enter the corresponding number next to your desired character's name."
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Label lblQuestion 
      BackColor       =   &H00000080&
      Caption         =   "Which character would you like to learn more about?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   4680
      Width           =   2775
   End
End
Attribute VB_Name = "frmCharacters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Twilight
'Form Name: frmCharacters
'Author: Mollie Land
'Date Written: 3/20/09
'Objective: This form has several options for learning more about some of the main characters in Twilight
'Each button gives the user a different option for learning more about the characters
'Initially only two buttons are available so that the array is read before the user can move
'on to learning facts or seeing a picture of the characters

'Dim variables used for multiple buttons (global variables)
Dim CharacterImages(1 To 20) As String, Characters(1 To 20) As String


Private Sub cmdFacts_Click()
    'This button will display facts about the characters
    'It will do so by using a select case statement which will then message box
    'the character's name using the array and then a fact statement about the character

    'Dim variables subject to just this button
    Dim CharacterFact As Integer

    'get a number from the user to determine which character facts to display
    CharacterFact = txtCharacterName.Text
    
    'Display character facts using select case statement
    'The number displayed next to the character name matches to the fact in each case
    'If the user does not enter a number that is specified, they are told they have entered an invalid number
    Select Case CharacterFact
        Case Is = 1
            MsgBox Characters(CharacterFact) & " is Edward's vampire sister who can see the future.", , "Character Facts"
        Case Is = 2
            MsgBox Characters(CharacterFact) & " is a small town girl who falls in love with Edward.", , "Character Facts"
        Case Is = 3
            MsgBox Characters(CharacterFact) & " is the vampire father of the Cullen family who also serves as the town doctor.", , "Character Facts"
        Case Is = 4
            MsgBox Characters(CharacterFact) & " is Bella's father who works for the Fork's police department.", , "Character Facts"
        Case Is = 5
            MsgBox Characters(CharacterFact) & " is a vampire who can read minds and is in love with Bella.", , "Character Facts"
        Case Is = 6
            MsgBox Characters(CharacterFact) & " is a vampire and a member of the Cullen family.", , "Character Facts"
        Case Is = 7
            MsgBox Characters(CharacterFact) & " is the mother of the Cullen family.", , "Character Facts"
        Case Is = 8
            MsgBox Characters(CharacterFact) & " is a werewolf and Bella's best friend.", , "Character Facts"
        Case Is = 9
            MsgBox Characters(CharacterFact) & " is a vampire tracking Bella.", , "Character Facts"
        Case Is = 10
            MsgBox Characters(CharacterFact) & " is a vampire who can control feelings and is a member of the Cullen family.", , "Character Facts"
        Case Is = 11
            MsgBox Characters(CharacterFact) & " is Bella's mother who lives in Florida.", , "Character Facts"
        Case Is = 12
            MsgBox Characters(CharacterFact) & " is a vampire and a member of the Cullen family.", , "Character Facts"
        Case Is = 13
            MsgBox Characters(CharacterFact) & " is James's mate who is a vampire after Bella as well.", , "Character Facts"
        Case Else
            MsgBox "You have entered an invalid number. Please try again by entering a number between 1 and 13.", , "Error"
    End Select
     
    'clears the text box so the user can enter a new number for a new character
    txtCharacterName.Text = ""
    
End Sub

Private Sub cmdPicture_Click()
    'this button will show a picture from the array of the chosen character
    'by getting an input number from the user
    
    'Dim variables subject to this button
    Dim Picture As Integer
    
    'clear the picture box
    picResults.Picture = LoadPicture("")
    
    'get a number from the user using a text box
    Picture = txtCharacterName.Text
    
    'this code works so that if an invalid number is put in then the user is asked
    'to input a number within the range
    'this loop continues until a number is entered from the user within the correct range
    Do While (Picture < 1 Or Picture > 13)
        txtCharacterName.Text = InputBox("Enter an integer between 1 and 13.")
        Picture = txtCharacterName.Text
    Loop
    
    'find the user's number in the array to display the picture in the picResults box
    picResults.Picture = LoadPicture(App.Path & "/" & CharacterImages(Picture))
    
    'clear the text box so the user can enter a new number for a new character
    txtCharacterName.Text = ""
    
End Sub

Private Sub cmdRead_Click()
    'this button will open the file and display the character names in
    'alphabetical order so the user can know which characters they can
    'learn more about
    'Each character name will have a number printed next to it as well
    'this number is used in the text box so the user can learn more about the characters
    
   
    'Dim Variables subject to this button
    Dim Pass As Integer, Pos As Integer, Temp As String
    
    'Initialize the counter
    CTR = 0
    
    'Open the file
    Open App.Path & "/CharacterList.txt" For Input As #1
    
    'Read the file and close it when it has been read
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, Characters(CTR), CharacterImages(CTR)
    Loop
    Close (1)
    
    'put the list into alphabetical order and match the image with the name
    'this is a bubble sort
    For Pass = 1 To CTR - 1
        For Pos = 1 To (CTR - Pass)
            If Characters(Pos) > Characters(Pos + 1) Then
                Temp = Characters(Pos)
                Characters(Pos) = Characters(Pos + 1)
                Characters(Pos + 1) = Temp
                
                Temp = CharacterImages(Pos)
                CharacterImages(Pos) = CharacterImages(Pos + 1)
                CharacterImages(Pos + 1) = Temp
            End If
        Next Pos
    Next Pass
    
    'print the list with a coordinating number beside the list using For/Next
    For J = 1 To CTR
        picNames.Print J; ".) "; Characters(J)
    Next J
    
    'Enable the other buttons so that the user has the option of what they want to learn about
    cmdPicture.Enabled = True
    cmdFacts.Enabled = True
    
End Sub

Private Sub cmdReturn_Click()
    'clear list of names
    picNames.Cls
    
    'clear the picture box
    picResults.Picture = LoadPicture("")

    
    'disable the picture and facts buttons
    cmdPicture.Enabled = False
    cmdFacts.Enabled = False
    
    'return to main menu, hiding the Charactes form
    frmStart.Show
    frmCharacters.Hide
End Sub

