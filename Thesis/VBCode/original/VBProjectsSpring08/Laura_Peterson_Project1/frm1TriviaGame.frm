VERSION 5.00
Begin VB.Form frm1TriviaGame 
   Caption         =   "Trivia Game"
   ClientHeight    =   7050
   ClientLeft      =   3015
   ClientTop       =   3765
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   Picture         =   "frm1TriviaGame.frx":0000
   ScaleHeight     =   7050
   ScaleWidth      =   8730
   Begin VB.TextBox txtAnswer 
      Height          =   1935
      Left            =   1800
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdRules 
      BackColor       =   &H00FF80FF&
      Caption         =   "Rules"
      BeginProperty Font 
         Name            =   "Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FF80FF&
      Caption         =   "Return to Game Menu"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000009&
      Height          =   3255
      Left            =   1920
      ScaleHeight     =   3195
      ScaleWidth      =   6435
      TabIndex        =   1
      Top             =   3000
      Width           =   6495
   End
   Begin VB.CommandButton cmdAnswer 
      BackColor       =   &H00FF80FF&
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "First! Enter difficulty level==>       1=Easy          2=Medium    3=Hard "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "frm1TriviaGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'option explicit
'Laura's Movie Gallery
'frmTriviaGame
'Laura Peterson
'03/24/2008
'This form allows the user to test the knowledge they have learned about the movies
'in my gallery by playing a trivia game.
Dim Number As Integer
Dim Response As Integer
Dim CTR As Integer
Dim Answer As Integer
Dim Percent As Single

'
Private Sub cmdAnswer_Click()
'this will get a number from the user which indicates a difficulty level for the game
Number = txtAnswer.Text
CTR = 0
Dim QCTR As Integer
QCTR = 0
Select Case Number
        Case 1
        'this will clear the picture box
            picResults.Cls
        'this will print the first question and possible answers
            picResults.Print "#1"
            picResults.Print "In the Wizard of Oz, the main character Dorothy hails from which state?"
            picResults.Print "1) Nebraska"
            picResults.Print "2) Kansas"
            picResults.Print "3) Arkansas"
        'this will allow the user to input the number of their answer based on the choices given
            Response = InputBox("Enter the number of your answer")
        'this will detect whether their answer was correct and give them a response letting them know.
              Select Case Response
                Case 1, 3
                    MsgBox "Sorry, response 2 was correct the correct answer."
                    'This counter will count the number of questions
                    QCTR = QCTR + 1
                Case 2
                    MsgBox "Correct!"
                    'this counter will count the number of correct questions answered
                    CTR = CTR + 1
                    QCTR = QCTR + 1
                'if the user enters a number other than 1, 2 or 3 an error message will appear
                Case Else
                    MsgBox "You must choose 1, 2 or 3", , "Error"
            
            End Select
            
            picResults.Cls
            picResults.Print "#2"
            picResults.Print "Which film sparked controversy with the vulgar line ";
            picResults.Print "Frankly my dear, I don 't give a damn!"
            picResults.Print "1) Gone with the Wind"
            picResults.Print "2) Casablanca"
            picResults.Print "3) On the Waterfront"
            
            Response = InputBox("Enter the number of your answer")
            Select Case Response
                Case 2, 3
                    MsgBox "Sorry, response 1 was correct was the correct answer."
                    QCTR = QCTR + 1
                Case 1
                    MsgBox "Correct!"
                    CTR = CTR + 1
                    QCTR = QCTR + 1
                Case Else
                    MsgBox "You must choose 1, 2 or 3", , "Error"
            
            End Select
            
            picResults.Cls
            picResults.Print "#3"
            picResults.Print "Marlon Brando stars as Don Vito Corleone in this 1970s "
            picResults.Print "crime film."
            picResults.Print "1) One Flew Over the Cuckoo's Nest"
            picResults.Print "2) The Godfather"
            picResults.Print "3) Bonnie and Clyde"
            Response = InputBox("Enter the number of your answer")
            Select Case Response
                Case 1, 3
                    MsgBox "Sorry, response 2 was correct the correct answer."
                    QCTR = QCTR + 1
                Case 2
                    MsgBox "Correct!"
                    CTR = CTR + 1
                    QCTR = QCTR + 1
                Case Else
                    MsgBox "You must choose 1, 2 or 3", , "Error"
            End Select
            
            picResults.Cls
            picResults.Print "#4"
            picResults.Print "This film is based on John Steinbeck's best-selling novel."
            picResults.Print "1) Some Like it Hot"
            picResults.Print "2) The Godfather"
            picResults.Print "3) The Grapes of Wrath"
            Response = InputBox("Enter the number of your answer")
            Select Case Response
                Case 1, 2
                    MsgBox "Sorry, response 3 was correct the correct answer."
                    QCTR = QCTR + 1
                Case 3
                    MsgBox "Correct!"
                    CTR = CTR + 1
                    QCTR = QCTR + 1
                Case Else
                    MsgBox "You must choose 1, 2 or 3", , "Error"
            End Select
            
            picResults.Cls
            picResults.Print "#5"
            picResults.Print "Carrie Fisher stars as Princess Leia in this "
            picResults.Print "blockbuster fantasy film."
            picResults.Print "1) 2001: A Space Odyssey"
            picResults.Print "2) Star Wars"
            picResults.Print "3) Apocalypse Now"
            Response = InputBox("Enter the number of your answer")
            Select Case Response
                Case 1, 3
                    MsgBox "Sorry, response 2 was correct the correct answer."
                    QCTR = QCTR + 1
                Case 2
                    MsgBox "Correct!"
                    CTR = CTR + 1
                    QCTR = QCTR + 1
                Case Else
                    MsgBox "You must choose 1, 2 or 3", , "Error"
                
            End Select
            'the program will then print the number of questions asked and the number answered correctly
            'then it will calculate and print the percentage correct
            Percent = CTR / QCTR
                picResults.Cls
                picResults.Print "Congratulations you got"; CTR; "out of"; QCTR; "answers correct!"
                picResults.Print "or"; Int(Percent * 100); "%"
        
        Case 2
            picResults.Cls
            picResults.Print "#1"
            picResults.Print "Frodo's strange and mighty odyssey to the Cracks of "
            picResults.Print "Doom is highlighted in this fantasy epic."
            picResults.Print "1)Lawrence of Arabia"
            picResults.Print "2)2001: A Space Odyssey"
            picResults.Print "3)The Lord of the Rings: The Fellowship of the Ring"
            
            Response = InputBox("Enter the number of your answer")
            Select Case Response
                Case 1, 2
                    MsgBox "Sorry, response 3 was correct the correct answer."
                    QCTR = QCTR + 1
                Case 3
                    MsgBox "Correct!"
                    CTR = CTR + 1
                    QCTR = QCTR + 1
                Case Else
                    MsgBox "You must choose 1, 2 or 3", , "Error"
            End Select
            picResults.Cls
            picResults.Print "#2"
            picResults.Print "Fay Wray stars as the beauty who killed the beast "
            picResults.Print "in this early film known for its special effects."
            picResults.Print "1)King Kong"
            picResults.Print "2)Double Indemnity"
            picResults.Print "3)The Maltese Falcon"
            
            Response = InputBox("Enter the number of your answer")
            Select Case Response
                Case 2, 3
                    MsgBox "Sorry, response 1 was correct the correct answer."
                    QCTR = QCTR + 1
                Case 1
                    MsgBox "Correct!"
                    CTR = CTR + 1
                    QCTR = QCTR + 1
                Case Else
                    MsgBox "You must choose 1, 2 or 3", , "Error"
            End Select
            picResults.Cls
            picResults.Print "#3"
            picResults.Print "In One Flew Over the Cuckoo's Nest, this actor stars as a "
            picResults.Print "troublemaker who is committed to a mental institution. "
            picResults.Print "1)  Michael Douglas"
            picResults.Print "2)  Jack Nicholson"
            picResults.Print "3)  Dustin Hoffman"
            Response = InputBox("Enter the number of your answer")
            
            Select Case Response
                Case 1, 3
                    MsgBox "Sorry, response 2 was correct the correct answer."
                    QCTR = QCTR + 1
                Case 2
                    MsgBox "Correct!"
                    CTR = CTR + 1
                    QCTR = QCTR + 1
                Case Else
                    MsgBox "You must choose 1, 2 or 3", , "Error"
            End Select
             picResults.Cls
            picResults.Print "#4"
            picResults.Print "This 1960 film is best known for its haunting shower "
            picResults.Print "scene and original score."
            picResults.Print "1)  Dr. Strangelove"
            picResults.Print "2)  Psycho"
            picResults.Print "3)  Rear Window"
            Response = InputBox("Enter the number of your answer")
            
            Select Case Response
                Case 1, 3
                    MsgBox "Sorry, response 2 was correct the correct answer."
                    QCTR = QCTR + 1
                Case 2
                    MsgBox "Correct!"
                    CTR = CTR + 1
                    QCTR = QCTR + 1
                Case Else
                    MsgBox "You must choose 1, 2 or 3", , "Error"
            End Select
              picResults.Cls
            picResults.Print "#5"
            picResults.Print "This film is highlighted by Simon and Garfunkel's"
            picResults.Print "songs, such as the title track 'Mrs. Robinson.'"
            picResults.Print "1)  The Graduate"
            picResults.Print "2)  One Flew Over the Cuckoo's Nest"
            picResults.Print "3)  On the Waterfront"
            Response = InputBox("Enter the number of your answer")
            
            Select Case Response
                Case 2, 3
                    MsgBox "Sorry, response 1 was correct the correct answer."
                    QCTR = QCTR + 1
                Case 1
                    MsgBox "Correct!"
                    CTR = CTR + 1
                    QCTR = QCTR + 1
                Case Else
                    MsgBox "You must choose 1, 2 or 3", , "Error"
            End Select
            picResults.Cls
            picResults.Print "#6"
            picResults.Print "Which science fiction film features a computer named HAL?"
            picResults.Print "1)  E.T. The Extra-Terrestrial"
            picResults.Print "2)  Star Wars"
            picResults.Print "3)  2001: A Space Odyssey"
            Response = InputBox("Enter the number of your answer")
            
            Select Case Response
                Case 2, 1
                    MsgBox "Sorry, response 3 was correct the correct answer."
                    QCTR = QCTR + 1
                Case 3
                    MsgBox "Correct!"
                    CTR = CTR + 1
                    QCTR = QCTR + 1
                Case Else
                    MsgBox "You must choose 1, 2 or 3", , "Error"
            End Select
            picResults.Cls
            picResults.Print "#7"
            picResults.Print "What film shares the idea that each time a bell"
            picResults.Print "rings an angel gets his wings?"
            picResults.Print "1)  It's a Wonderful Life"
            picResults.Print "2)  The Wizard of Oz"
            picResults.Print "3)  The Philadelphia Story"
            Response = InputBox("Enter the number of your answer")
            
            Select Case Response
                Case 2, 3
                    MsgBox "Sorry, response 1 was correct the correct answer."
                    QCTR = QCTR + 1
                Case 1
                    MsgBox "Correct!"
                    CTR = CTR + 1
                    QCTR = QCTR + 1
                Case Else
                    MsgBox "You must choose 1, 2 or 3", , "Error"
            End Select
                Percent = CTR / QCTR
                picResults.Cls
                picResults.Print "Congratulations you got"; CTR; "out of"; QCTR; "answers correct!"
                picResults.Print "or"; Int(Percent * 100); "%"
                
        Case 3
            picResults.Cls
            picResults.Print "#1"
            picResults.Print "He recreates the role which made him a "
            picResults.Print "star on Broadway in A Streetcar Named Desire."
            picResults.Print "1) Robert De Niro"
            picResults.Print "2) Al Pacino"
            picResults.Print "3) Marlon Brando"
            Response = InputBox("Enter the number of your answer")
            Select Case Response
                Case 2, 1
                    MsgBox "Sorry, response 3 was correct the correct answer."
                    QCTR = QCTR + 1
                Case 3
                    MsgBox "Correct!"
                    CTR = CTR + 1
                    QCTR = QCTR + 1
                Case Else
                    MsgBox "You must choose 1, 2 or 3", , "Error"
            End Select
             picResults.Cls
            picResults.Print "#2"
            picResults.Print "Diane Keaton spawned a new fashion trend "
            picResults.Print "in this 1977 film by Woody Allen."
            picResults.Print "1)Annie Hall"
            picResults.Print "2)  The Best Years of Our Lives"
            picResults.Print "3)  Chinatown"
            Response = InputBox("Enter the number of your answer")
            Select Case Response
                Case 2, 3
                    MsgBox "Sorry, response 1 was correct the correct answer."
                    QCTR = QCTR + 1
                Case 1
                    MsgBox "Correct!"
                    CTR = CTR + 1
                    QCTR = QCTR + 1
                Case Else
                    MsgBox "You must choose 1, 2 or 3", , "Error"
            End Select
            picResults.Cls
            picResults.Print "#3"
            picResults.Print "Which of these films does Humphrey Bogart NOT star in?"
            picResults.Print "1)  The Maltese Falcon"
            picResults.Print "2)  The Treasure of the Sierra Madre"
            picResults.Print "3)  It's a Wonderful Life"
            Response = InputBox("Enter the number of your answer")
            Select Case Response
                Case 2, 1
                    MsgBox "Sorry, response 3 was correct the correct answer."
                    QCTR = QCTR + 1
                Case 3
                    MsgBox "Correct!"
                    CTR = CTR + 1
                    QCTR = QCTR + 1
                Case Else
                    MsgBox "You must choose 1, 2 or 3", , "Error"
            End Select
            picResults.Cls
            picResults.Print "#4"
            picResults.Print "In the film Mr. Smith Goes to Washington,James"
            picResults.Print "Stewart's character holds what position in government?"
            picResults.Print "1)  Governor "
            picResults.Print "2)  U.S. Senator"
            picResults.Print "3)  U.S. Representative"
            Response = InputBox("Enter the number of your answer")
            Select Case Response
                Case 1, 3
                    MsgBox "Sorry, response 2 was correct the correct answer."
                    QCTR = QCTR + 1
                Case 2
                    MsgBox "Correct!"
                    CTR = CTR + 1
                    QCTR = QCTR + 1
                Case Else
                    MsgBox "You must choose 1, 2 or 3", , "Error"
             End Select
            picResults.Cls
            picResults.Print "#5"
            picResults.Print "Which of the following films stars Gary Cooper?"""
            picResults.Print "1)  High Noon"
            picResults.Print "2)  The Philadelphia Story"
            picResults.Print "3)  The Grapes of Wrath"
            Response = InputBox("Enter the number of your answer")
            Select Case Response
                Case 2, 3
                    MsgBox "Sorry, response 1 was correct the correct answer."
                    QCTR = QCTR + 1
                Case 1
                    MsgBox "Correct!"
                    CTR = CTR + 1
                    QCTR = QCTR + 1
                Case Else
                    MsgBox "You must choose 1, 2 or 3", , "Error"
            End Select
                Percent = CTR / QCTR
                picResults.Cls
                picResults.Print "Congratulations you got"; CTR; "out of"; QCTR; "answers correct!"
                picResults.Print "or"; Int(Percent * 100); "%"
End Select

End Sub
'this will return to the game menu
Private Sub cmdReturn_Click()
frmGames.Show
End Sub
'this button will print the rules of the game in the picture box
Private Sub cmdRules_Click()
picResults.Cls
picResults.Print "All questions in this game relate to the 50 movies "
picResults.Print "in Laura's Movie Gallery. First, you must chose a"
picResults.Print "level of difficulty. You must then answer the"
picResults.Print "questions in the level by inputing the corresponding"
picResults.Print "number into the input box via the button on the right."
picResults.Print "When you are finished with the round you will recieve"
picResults.Print "a message telling you how you did. Have fun!"
picResults.Print "NOTE* You must finish the round before moving on."
End Sub

