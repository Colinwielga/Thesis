VERSION 5.00
Begin VB.Form frmTestFlashCards 
   BackColor       =   &H00000080&
   Caption         =   "Lingua Vivens - Student Options - Flash Card Test (No Grade)"
   ClientHeight    =   5310
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16125
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   16125
   Begin VB.Frame Frame1 
      BackColor       =   &H00000080&
      Caption         =   "Select a Category to Quiz ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   12360
      TabIndex        =   14
      Top             =   720
      Width           =   3255
      Begin VB.OptionButton optAdjectives 
         BackColor       =   &H00000080&
         Caption         =   "Adjectives"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1500
         TabIndex        =   20
         Top             =   1440
         Width           =   1815
      End
      Begin VB.OptionButton optAdverbs 
         BackColor       =   &H00000080&
         Caption         =   "Adverbs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1500
         TabIndex        =   19
         Top             =   840
         Width           =   1695
      End
      Begin VB.OptionButton optParticles 
         BackColor       =   &H00000080&
         Caption         =   "Particles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1500
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optVerbs 
         BackColor       =   &H00000080&
         Caption         =   "Verbs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   1695
      End
      Begin VB.OptionButton optNouns 
         BackColor       =   &H00000080&
         Caption         =   "Nouns"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   2055
      End
      Begin VB.OptionButton optAll 
         BackColor       =   &H00000080&
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdLogOut 
      BackColor       =   &H00808080&
      Caption         =   "LogOut"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00808080&
      Caption         =   "Return to Student Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00000080&
      Caption         =   "Begin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Frame fraPartOfSpeech 
      BackColor       =   &H00000080&
      Caption         =   "Show Part of Speech with ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8160
      TabIndex        =   7
      Top             =   1920
      Width           =   4095
      Begin VB.OptionButton optClue 
         BackColor       =   &H00000080&
         Caption         =   "Clue"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton optAnswer 
         BackColor       =   &H00000080&
         Caption         =   "Answer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdFlip 
      BackColor       =   &H00000080&
      Caption         =   "Flip Card"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00000080&
      Caption         =   "Next Card"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Frame fraLatinOrEnglish 
      BackColor       =   &H00000080&
      Caption         =   "Card Order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8160
      TabIndex        =   2
      Top             =   720
      Width           =   4095
      Begin VB.OptionButton optEnglish 
         BackColor       =   &H00000080&
         Caption         =   "Show English First"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   4
         Top             =   360
         Width           =   2000
      End
      Begin VB.OptionButton optLatin 
         BackColor       =   &H00000080&
         Caption         =   "Show Latin First"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdRandomize 
      BackColor       =   &H00000080&
      Caption         =   "Randomize Order"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   2175
   End
   Begin VB.PictureBox picCard 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   840
      ScaleHeight     =   2235
      ScaleWidth      =   6915
      TabIndex        =   0
      Top             =   720
      Width           =   6975
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H00000080&
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Flash Card Test"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   36
         Charset         =   255
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   8280
      TabIndex        =   22
      Top             =   3840
      Width           =   4815
   End
   Begin VB.Label lblLanguage 
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Displaying : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   21
      Top             =   240
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   5055
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   15615
   End
End
Attribute VB_Name = "frmTestFlashCards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Form Level variable which keeps track of the position of the array in order to flip the card, and then move to the next one
Dim FlashPos As Integer
'Used to check and see if at least on of a given criteris is found (initialized in cmdStart_Click() and tested in TestandPrint)
Dim FoundOne As Boolean
Dim Flipped As Boolean


Private Sub cmdFlip_Click()
    'This is the code to see the opposite pair of words, or the other side of the flash card
    'Clears picCard
    picCard.Cls
    'Evaluates the options for the flash card (whether the latin is first or english first and then whether the part of speech is with the clue or the answer)
    'and then prints the appropriate response
    'This is the reciprocal of the 'next' button
    If Not Flipped Then
        Flipped = True
        If optEnglish.Value = False Then
            If optAnswer.Value = False Then
                picCard.Print EnglishFlash(FlashPos)
            Else
                picCard.Print EnglishFlash(FlashPos) & ", " & LCase(partSpeechFlash(FlashPos)) & "."
            End If
            lblLanguage.Caption = "Displaying: ENGLISH"
        ElseIf optLatin.Value = False Then
            If optAnswer.Value = False Then
                picCard.Print LatinFlash(FlashPos)
            Else
                picCard.Print LatinFlash(FlashPos) & ", " & LCase(partSpeechFlash(FlashPos)) & "."
            End If
            lblLanguage.Caption = "Displaying: LATIN"
        End If
    Else
        Flipped = False
        If optEnglish.Value = True Then
            If optAnswer.Value = True Then
                picCard.Print EnglishFlash(FlashPos)
            Else
                picCard.Print EnglishFlash(FlashPos) & ", " & LCase(partSpeechFlash(FlashPos)) & "."
            End If
            lblLanguage.Caption = "Displaying: ENGLISH"
        ElseIf optLatin.Value = True Then
            If optAnswer.Value = True Then
                picCard.Print LatinFlash(FlashPos)
            Else
                picCard.Print LatinFlash(FlashPos) & ", " & LCase(partSpeechFlash(FlashPos)) & "."
            End If
            lblLanguage.Caption = "Displaying: LATIN"
        End If
    End If

End Sub

Private Sub cmdLogOut_Click()
    'Runs the public logout subroutine (cf. mdlPublicSubs)
    frmTestFlashCards.Hide
    Call LogOut
End Sub

Private Sub cmdNext_Click()
    ' This button is to move forward to the next position in the array and display the next 'card'
    'Clears picCard
    picCard.Cls
    Flipped = False
        'Checks to see which set of data the user is testing all or any one of the parts of speech
2        If optAll.Value = True Then
            If FlashPos < flashCtr Then
                'Increments the pos of the array with each click (initial value = 2)
                FlashPos = FlashPos + 1
                'Checks to see in which order the user wishes to test his/her vocabulary whether latin first or english first
                If optEnglish.Value = True Then
                    'Checks to see if the user wants the part of speech with the answer or the clue (this is the reciprocal of the flip If-statement)
                    If optAnswer.Value = True Then
                        picCard.Print EnglishFlash(FlashPos)
                    Else
                        picCard.Print EnglishFlash(FlashPos) & ", " & LCase(partSpeechFlash(FlashPos)) & "."
                    End If
                    lblLanguage.Caption = "Displaying: ENGLISH"
                ElseIf optLatin.Value = True Then
                    If optAnswer.Value = True Then
                        picCard.Print LatinFlash(FlashPos)
                    Else
                        picCard.Print LatinFlash(FlashPos) & ", " & LCase(partSpeechFlash(FlashPos)) & "."
                    End If
                    lblLanguage.Caption = "Displaying: LATIN"
                End If
            Else
                FlashPos = 2
                GoTo 2
            End If
        ElseIf optNouns.Value = True Then
            'Uses the form-level subroutine TestandPrint (cf. TestAndPrint)
            TestAndPrint ("N") 'Gives a value to the string variable in the public sub TestAndPrint
            'picCard.Print FlashPos
        ElseIf optVerbs.Value = True Then
            TestAndPrint ("V")
            'picCard.Print FlashPos
        ElseIf optAdverbs.Value = True Then
            TestAndPrint ("ADV")
            'picCard.Print FlashPos
        ElseIf optAdjectives.Value = True Then
            TestAndPrint ("ADJ")
            'picCard.Print FlashPos
        ElseIf optParticles.Value = True Then
            TestAndPrint ("PART")
            'picCard.Print FlashPos
        End If
End Sub

Private Sub cmdRandomize_Click()
    'Randomizes the order of the cards (the order of the arrays) so that the user can get learn better
    'Declares local varaibles
    Dim randomPos1 As Integer
    Dim RandomPos2 As Integer
    Dim Pass As Integer
    Dim pos As Integer
    'Randomizes the random number genearator of Visual Basic
    Randomize
    'Loops a significant number of times in order to properly randomizes (modified bubble-sort algorithm)
    'Does not move the first entry
    For Pass = 2 To flashCtr
        For pos = 2 To flashCtr - Pass
            'Sets the varaibles equal to random numbers from 2 to the end of the array
            randomPos1 = Int((flashCtr - 2 + 1) * Rnd + 2)
99          RandomPos2 = Int((flashCtr - 2 + 1) * Rnd + 2)
            'Checks to see if the two numbers are the same, if they are it goes back and randomizes the second number  until they are different
            If randomPos1 = RandomPos2 Then GoTo 99
            
            'Calls the public subroutines SwapString (cf. mdlPublicSubs)
            SwapString LatinFlash(randomPos1), LatinFlash(RandomPos2)
            SwapString EnglishFlash(randomPos1), EnglishFlash(RandomPos2)
            SwapString partSpeechFlash(randomPos1), partSpeechFlash(RandomPos2)
            
        Next pos
    Next Pass
    'Clears picCard
    picCard.Cls
    'Runs the if-algorithm described above in "Private Sub cmdNext_Click()"
    If optAll.Value = True Then
        FlashPos = 2
        If optEnglish.Value = True Then
            If optAnswer.Value = True Then
                picCard.Print EnglishFlash(FlashPos)
            Else
                picCard.Print EnglishFlash(FlashPos) & ", " & LCase(partSpeechFlash(FlashPos)) & "."
            End If
            lblLanguage.Caption = "Displaying: ENGLISH"
        ElseIf optLatin.Value = True Then
            If optAnswer.Value = True Then
                picCard.Print LatinFlash(FlashPos)
            Else
                picCard.Print LatinFlash(FlashPos) & ", " & LCase(partSpeechFlash(FlashPos)) & "."
            End If
            lblLanguage.Caption = "Displaying: LATIN"
        End If
    ElseIf optNouns.Value = True Then
        FlashPos = 1
        TestAndPrint ("N")
    ElseIf optVerbs.Value = True Then
        FlashPos = 1
        TestAndPrint ("V")
    ElseIf optAdverbs.Value = True Then
        FlashPos = 1
        TestAndPrint ("ADV")
    ElseIf optAdjectives.Value = True Then
        FlashPos = 1
        TestAndPrint ("ADJ")
    ElseIf optParticles.Value = True Then
        FlashPos = 1
        TestAndPrint ("PART")
    End If
    
    
End Sub

Private Sub cmdReturn_Click()
    'Returns to the Student Options Pane
    frmTestFlashCards.Hide
    frmOptionsPage.Show
End Sub
Public Sub TestAndPrint(x As String)
    'This public Subroutine is used to search the array for a given criteria, and to return the pos of the next item matching the criteria and then to print it
    'Declares and initializes the found boolean used for searching
1     Dim Found As Boolean
    
    
    Found = False
    'Uses a match and stop search to find the next item matching the criteria (which is given within the Private Sub cmdNExt_Click() )
    Do Until Found Or FlashPos = flashCtr
        FlashPos = FlashPos + 1
        If x = partSpeechFlash(FlashPos) Then
            Found = True
            FoundOne = True
        End If
    Loop
    'Prints the match using the same algoritm as in the private sub cmdNext_Click()
    If Found Then
       
        If optEnglish.Value = True Then
            If optAnswer.Value = True Then
                picCard.Print EnglishFlash(FlashPos)
            Else
                picCard.Print EnglishFlash(FlashPos) & ", " & LCase(partSpeechFlash(FlashPos)) & "."
            End If
            lblLanguage.Caption = "Displaying: ENGLISH"
        ElseIf optLatin.Value = True Then
            If optAnswer.Value = True Then
                picCard.Print LatinFlash(FlashPos)
            Else
                picCard.Print LatinFlash(FlashPos) & ", " & LCase(partSpeechFlash(FlashPos)) & "."
            End If
            lblLanguage.Caption = "Displaying: LATIN"
        End If
    ElseIf FoundOne And FlashPos = flashCtr Then
        FlashPos = 1
        GoTo 1
    ElseIf Not Found Then 'If there are no matches to the criteria this returns an error message
        MsgBox "There are no " & x & " words in your data file, please select another criterion."
    End If
End Sub
Private Sub cmdStart_Click()
    'Used to begin the flashCard session, enables various buttton and intitalizes the first entry setting FlashPos = 2 and etc.
    'Enables all the buttons and sets itself to invisible
    cmdRandomize.Enabled = True
    cmdNext.Enabled = True
    cmdFlip.Enabled = True
    cmdStop.Visible = True
    cmdStart.Visible = False
    FoundOne = False
    'Reads the data into arrays (essentially to remove any randomization and return it to default order)
    'Opens the text file for the surrent user
    Open App.Path & "\Data\FlashCards\" & userName(StudentPosition) & ".txt" For Input As #1
        'Initializes flashCtr
        flashCtr = 0
        'Loops in order to read the data
        Do Until EOF(1)
            'Increments FlashCtr
            flashCtr = flashCtr + 1
            'Reads the data into the arrays
            Input #1, LatinFlash(flashCtr), EnglishFlash(flashCtr), partSpeechFlash(flashCtr)
        Loop
    Close #1
    'Clears picCard
    picCard.Cls
   'Uses the same if-algorithm as described as above in Private Sub cmdNext_Click()
   If optAll.Value = True Then
        FlashPos = 2 'Initializes the flashCtr at the second poition of the array
        If optEnglish.Value = True Then
            If optAnswer.Value = True Then
                picCard.Print EnglishFlash(FlashPos)
            Else
                picCard.Print EnglishFlash(FlashPos) & ", " & LCase(partSpeechFlash(FlashPos)) & "."
            End If
            lblLanguage.Caption = "Displaying: ENGLISH"
        ElseIf optLatin.Value = True Then
            If optAnswer.Value = True Then
                picCard.Print LatinFlash(FlashPos)
            Else
                picCard.Print LatinFlash(FlashPos) & ", " & LCase(partSpeechFlash(FlashPos)) & "."
            End If
            lblLanguage.Caption = "Displaying: LATIN"
        End If
    ElseIf optNouns.Value = True Then
        FlashPos = 1 'initializes flashCtr at 1 in order to run a proper search of the array
        TestAndPrint ("N") 'Gives a value to the string variable in the public sub TestAndPrint
    ElseIf optVerbs.Value = True Then
        FlashPos = 1
        TestAndPrint ("V")
    ElseIf optAdverbs.Value = True Then
        FlashPos = 1
        TestAndPrint ("ADV")
    ElseIf optAdjectives.Value = True Then
        FlashPos = 1
        TestAndPrint ("ADJ")
    ElseIf optParticles.Value = True Then
        FlashPos = 1
        TestAndPrint ("PART")
    End If
    
End Sub

Private Sub cmdStop_Click()
    'Stops the flashCard session
    'Disables the buttons and makes itself invisible
    cmdRandomize.Enabled = False
    cmdNext.Enabled = False
    cmdFlip.Enabled = False
    cmdStop.Visible = False
    cmdStart.Visible = True
    'reinitializes the flashctr varaible at 2 and clears picCard
    FlashPos = 2
    picCard.Cls
    lblLanguage.Caption = "Displaying: "
End Sub

Private Sub optAnswer_Click()
    'updates the users new selecion and displays a card appropriate to the users option selections
    picCard.Cls
    If optEnglish.Value = True Then
        If optAnswer.Value = True Then
            picCard.Print EnglishFlash(FlashPos)
        Else
            picCard.Print EnglishFlash(FlashPos) & ", " & LCase(partSpeechFlash(FlashPos)) & "."
        End If
        lblLanguage.Caption = "Displaying: ENGLISH"
    ElseIf optLatin.Value = True Then
        If optAnswer.Value = True Then
            picCard.Print LatinFlash(FlashPos)
        Else
            picCard.Print LatinFlash(FlashPos) & ", " & LCase(partSpeechFlash(FlashPos)) & "."
        End If
        lblLanguage.Caption = "Displaying: LATIN"
    End If
End Sub

Private Sub optClue_Click()
    'updates the users new selecion and displays a card appropriate to the users option selections
    picCard.Cls
    If optEnglish.Value = True Then
        If optAnswer.Value = True Then
            picCard.Print EnglishFlash(FlashPos)
        Else
            picCard.Print EnglishFlash(FlashPos) & ", " & LCase(partSpeechFlash(FlashPos)) & "."
        End If
        lblLanguage.Caption = "Displaying: ENGLISH"
    ElseIf optLatin.Value = True Then
        If optAnswer.Value = True Then
            picCard.Print LatinFlash(FlashPos)
        Else
            picCard.Print LatinFlash(FlashPos) & ", " & LCase(partSpeechFlash(FlashPos)) & "."
        End If
        lblLanguage.Caption = "Displaying: LATIN"
    End If
End Sub

Private Sub optEnglish_Click()
    'updates the users new selecion and displays a card appropriate to the users option selections
    picCard.Cls
    If optEnglish.Value = True Then
        If optAnswer.Value = True Then
            picCard.Print EnglishFlash(FlashPos)
        Else
            picCard.Print EnglishFlash(FlashPos) & ", " & LCase(partSpeechFlash(FlashPos)) & "."
        End If
        lblLanguage.Caption = "Displaying: ENGLISH"
    ElseIf optLatin.Value = True Then
        If optAnswer.Value = True Then
            picCard.Print LatinFlash(FlashPos)
        Else
            picCard.Print LatinFlash(FlashPos) & ", " & LCase(partSpeechFlash(FlashPos)) & "."
        End If
        lblLanguage.Caption = "Displaying: LATIN"
    End If
End Sub

Private Sub optLatin_Click()
    'updates the users new selecion and displays a card appropriate to the users option selections
    picCard.Cls
    If optEnglish.Value = True Then
        If optAnswer.Value = True Then
            picCard.Print EnglishFlash(FlashPos)
        Else
            picCard.Print EnglishFlash(FlashPos) & ", " & LCase(partSpeechFlash(FlashPos)) & "."
        End If
        lblLanguage.Caption = "Displaying: ENGLISH"
    ElseIf optLatin.Value = True Then
        If optAnswer.Value = True Then
            picCard.Print LatinFlash(FlashPos)
        Else
            picCard.Print LatinFlash(FlashPos) & ", " & LCase(partSpeechFlash(FlashPos)) & "."
        End If
        lblLanguage.Caption = "Displaying: LATIN"
    End If
End Sub
