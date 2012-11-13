VERSION 5.00
Begin VB.Form frmCreateFlashCards 
   BackColor       =   &H00000080&
   Caption         =   "Lingua Vivens - Student Options - Create Flash Cards"
   ClientHeight    =   10515
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17805
   LinkTopic       =   "Form1"
   ScaleHeight     =   10515
   ScaleWidth      =   17805
   Begin VB.CommandButton cmdAddVocabFromLists 
      BackColor       =   &H00000080&
      Caption         =   "Add Vocabulary from Test Lists"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00000080&
      Caption         =   "Reset Flash Card Data"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2760
      Width           =   2175
   End
   Begin VB.PictureBox picWords 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   8640
      ScaleHeight     =   8955
      ScaleWidth      =   7515
      TabIndex        =   14
      Top             =   600
      Width           =   7575
   End
   Begin VB.Frame fraPartOfSpeech 
      BackColor       =   &H00000080&
      Caption         =   "Select Part of Speech"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   3240
      TabIndex        =   8
      Top             =   2520
      Width           =   5055
      Begin VB.OptionButton optParticle 
         BackColor       =   &H00000080&
         Caption         =   "Particle (prepositions, interjections, etc.)"
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
         TabIndex        =   13
         Top             =   2760
         Width           =   3135
      End
      Begin VB.OptionButton optAdjective 
         BackColor       =   &H00000080&
         Caption         =   "Adjective"
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
         TabIndex        =   12
         Top             =   2160
         Width           =   2295
      End
      Begin VB.OptionButton optAdverb 
         BackColor       =   &H00000080&
         Caption         =   "Adverb"
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
         TabIndex        =   11
         Top             =   1560
         Width           =   2415
      End
      Begin VB.OptionButton optVerb 
         BackColor       =   &H00000080&
         Caption         =   "Verb"
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
         TabIndex        =   10
         Top             =   960
         Width           =   2535
      End
      Begin VB.OptionButton optNoun 
         BackColor       =   &H00000080&
         Caption         =   "Noun"
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
         TabIndex        =   9
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.TextBox txtEnglish 
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox txtLatin 
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   960
      Width           =   2775
   End
   Begin VB.CommandButton cmdShow 
      BackColor       =   &H00000080&
      Caption         =   "Display Current Vocabulary"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton cmdLogout 
      BackColor       =   &H00808080&
      Caption         =   "Log Out"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00000080&
      Caption         =   "Add Vocabulary"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Flash Card Creator"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   39.75
         Charset         =   255
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   1320
      TabIndex        =   18
      Top             =   7920
      Width           =   6375
   End
   Begin VB.Label lblRecommendation 
      BackStyle       =   0  'Transparent
      Caption         =   "(Recomend including gender with nouns)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3240
      TabIndex        =   17
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblEnglish 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter English Translation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lblLatin 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Latin Word"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   960
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   10095
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   17295
   End
End
Attribute VB_Name = "frmCreateFlashCards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'formLevel varaible for use in checking if  a new word has been added before displaying a the word list
Dim newAdded As Boolean

Private Sub cmdAdd_Click()
    'This button will add new vocabulary to the flashcard lists of the user, by geting input via text boxes
    'Declares varaibles
    Dim newLatin As String
    Dim newEnglish As String
    Dim newPartSpeech As String
    Dim verify As String
    'Initializes varaibles
    newAdded = False
    newLatin = txtLatin.Text
    newEnglish = txtEnglish.Text
    'Reads the input from the option buttons and defines varaibles acordingly
    If optNoun.Value = True Then
        newPartSpeech = "N" 'Sets noun part for input into text files
        optNoun.Value = False ' resets optionbutton for future use
    ElseIf optVerb.Value = True Then
        newPartSpeech = "V"
        optVerb.Value = False
    ElseIf optAdverb.Value = True Then
        newPartSpeech = "ADV"
        optAdverb.Value = False
    ElseIf optAdjective.Value = True Then
        newPartSpeech = "ADJ"
        optAdjective.Value = False
    ElseIf optParticle.Value = True Then
        newPartSpeech = "PART"
        optParticle.Value = False
    Else 'Error handling for if user selecs no option button
        MsgBox "Please select a part of speech"
        Exit Sub 'exit subroutine for user to correct mistake
    End If
    'If statement to make sure that the user enters all data
    If newLatin = "" Or newEnglish = "" Then
        MsgBox "Please make sure that all fields are filled in"
    Else
        'Checks to see if the input is correct and verifies using an input box
        verify = InputBox("You are about to enter the " & UCase(newPartSpeech) & ", " & UCase(newLatin) & " meaning: " & UCase(newEnglish) & ". Is this what you wish to add, if yes type 'yes' if no type 'no'.")
        'if the user is sure that he/she want to enter the data then write it to a text file
        If LCase(verify) = "yes" Then
            Open App.Path & "\Data\FlashCards\" & userName(StudentPosition) & ".txt" For Append As #1
                Write #1, newLatin, newEnglish, newPartSpeech
            Close #1
            MsgBox UCase(newLatin) & " has been added to your flash cards."
            newAdded = True
            txtLatin.Text = ""
            txtEnglish.Text = ""
        Else
            MsgBox UCase(newLatin) & " has NOT been added to your flash cards."
        End If
    End If
    
    

End Sub

Private Sub cmdAddVocabFromLists_Click()
    'Adds the vocabulary list from the aministrator to the student's personal list
    'Declares useful varaibles
    Dim newLatin As String
    Dim newEnglish As String
    Dim newPartSpeech As String
    Dim pos As Integer
    Dim posSearch As Integer
    Dim Found As Boolean
    Dim searchWord As String
    Dim searchLength As Integer
    Dim searchSpace As Integer
    Dim genderName As String
    'Opens the unique flash card set for the user
   
    Open App.Path & "\Data\FlashCards\" & userName(StudentPosition) & ".txt" For Append As #1
        'loops over the verb arrays and list out each one with appropriate endings
        For pos = 1 To verbCtr
            Found = False
            posSearch = 1
            Do Until Found Or posSearch = flashCtr
                posSearch = posSearch + 1
                searchLength = Len(LatinFlash(posSearch))
                searchSpace = InStr(LatinFlash(posSearch), " ")
                searchWord = Right(LatinFlash(posSearch), searchLength - searchSpace)
                If searchWord = VerbInfinitive(pos) Then
                    Found = True
                End If
            Loop
            
            If Not Found Then
                Select Case VerbConjugation(pos)
                    Case 1, 3
                        newLatin = VerbPresStem(pos) & "o" & ", " & VerbInfinitive(pos)
                    Case 2
                        newLatin = VerbPresStem(pos) & "eo" & ", " & VerbInfinitive(pos)
                    Case 4, 5
                        newLatin = VerbPresStem(pos) & "io" & ", " & VerbInfinitive(pos)
                End Select
                
                newEnglish = VerbDefinition(pos)
                newPartSpeech = "V"
                Write #1, newLatin, newEnglish, newPartSpeech
            End If
            
        Next pos
        'Loops over the noun arrays and adds the nouns to the flash card txt file
        For pos = 1 To NounCtr
            Found = False
            posSearch = 1
            Do Until Found Or posSearch = flashCtr
                posSearch = posSearch + 1
                searchSpace = InStr(LatinFlash(posSearch), ",")
                searchWord = Left(LatinFlash(posSearch), searchSpace - 1)
                If searchWord = NomSNoun(pos) Then
                    Found = True
                End If
            Loop
            
            If Not Found Then
                newLatin = NomSNoun(pos) & ", " & GenSNoun(pos)
                newEnglish = definitionNoun(pos)
                newPartSpeech = "N"
                Select Case GenderNoun(pos)
                    Case 1
                        genderName = "F."
                    Case 2
                        genderName = "M."
                    Case 3
                        genderName = "N."
                End Select
                
                newLatin = NomSNoun(pos) & ", " & GenSNoun(pos) & " " & genderName
                Write #1, newLatin, newEnglish, newPartSpeech
            End If
        Next pos
        newAdded = True
    Close #1
    
    MsgBox "The vocabulary from the test lists has successfully been added to you flash card list."
End Sub

Private Sub cmdLogOut_Click()
    'Runs the public subroutine logOut, (cf. mdlPublicSubs)
    frmCreateFlashCards.Hide
    Call LogOut
End Sub


Private Sub cmdReset_Click()
    Dim verify As String
    
    verify = InputBox("You are about to clear your flashcard data. Is this what you wished to do?  Enter 'yes' to continue or 'no' to cancel")
    If LCase(verify) = "yes" Then
        Open App.Path & "\Data\FlashCards\" & userName(StudentPosition) & ".txt" For Output As #1
            Write #1, "This is the First Entry", "It will be ignored", "But it must be present"
        Close #1
        MsgBox "Your flashcard data has been cleared"
    ElseIf verify = vbNullString Or LCase(verify) = "no" Or verify <> "yes" Then
        MsgBox "Your flashcard data was not cleared"
    End If
    newAdded = True
    picWords.Cls
End Sub

Private Sub cmdReturn_Click()
    'Returns to the sutdent options pane
    frmCreateFlashCards.Hide
    frmOptionsPage.Show
End Sub

Private Sub cmdShow_Click()
    'Used to disply the entire word list for the user
    Dim pos As Integer
   
    'Checks to see if a new words has been added since form load
    If newAdded Then
        'Rests the flashCtr
        flashCtr = 0
        'Opens text file and rereads the data to arrays
        Open App.Path & "\Data\FlashCards\" & userName(StudentPosition) & ".txt" For Input As #1
            Do Until EOF(1)
                flashCtr = flashCtr + 1
                Input #1, LatinFlash(flashCtr), EnglishFlash(flashCtr), partSpeechFlash(flashCtr)
            Loop
        Close #1
    End If
    'Clears picWords and displays new header
    picWords.Cls
    picWords.Print "Latin Word"; Tab(30); "Part of Speech"; Tab(55); "English Definition"
    picWords.Print "*********************************************************************************************************************************"
    'prints the entire array to picWords as long as there is data in file, if not returns an error message
    If flashCtr <> 1 Then
        For pos = 2 To flashCtr
            picWords.Print LatinFlash(pos); Tab(30); partSpeechFlash(pos) & "."; Tab(55); EnglishFlash(pos)
        Next pos
    Else
        picWords.Print "There is no Vocab to display"
    End If
End Sub

Private Sub Form_Load()
    'Reads the flashCard for the given user into arrays for use by this form
    Open App.Path & "\Data\FlashCards\" & userName(StudentPosition) & ".txt" For Input As #1
        flashCtr = 0
        Do Until EOF(1)
            flashCtr = flashCtr + 1
            Input #1, LatinFlash(flashCtr), EnglishFlash(flashCtr), partSpeechFlash(flashCtr)
        Loop
    Close #1
End Sub
