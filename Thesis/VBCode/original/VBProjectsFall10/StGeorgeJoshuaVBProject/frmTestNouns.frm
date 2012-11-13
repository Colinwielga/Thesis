VERSION 5.00
Begin VB.Form frmTestNouns 
   BackColor       =   &H00000080&
   Caption         =   "Lingua Vivens - Student Tests - Noun Forms"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17310
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   17310
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6120
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6960
      Width           =   2175
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00000080&
      Caption         =   "Start"
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
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6120
      Width           =   2175
   End
   Begin VB.PictureBox picSentence 
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
      Height          =   615
      Left            =   5400
      ScaleHeight     =   555
      ScaleWidth      =   3915
      TabIndex        =   21
      Top             =   720
      Width           =   3975
   End
   Begin VB.PictureBox picNounsTested 
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
      Height          =   6495
      Left            =   10200
      ScaleHeight     =   6435
      ScaleWidth      =   6075
      TabIndex        =   9
      Top             =   1080
      Width           =   6135
   End
   Begin VB.PictureBox picGrade 
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
      Height          =   615
      Left            =   15240
      ScaleHeight     =   555
      ScaleWidth      =   1035
      TabIndex        =   7
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H00000080&
      Caption         =   "End Session"
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
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdSubmit 
      BackColor       =   &H00000080&
      Caption         =   "Submit and Display Next"
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Frame fraNumber 
      BackColor       =   &H00000080&
      Caption         =   "Number ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   3720
      TabIndex        =   4
      Top             =   2280
      Width           =   2655
      Begin VB.OptionButton optPlural 
         BackColor       =   &H00000080&
         Caption         =   "Plural"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   370
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   1935
      End
      Begin VB.OptionButton optSingular 
         BackColor       =   &H00000080&
         Caption         =   "Singular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   370
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame fraCase 
      BackColor       =   &H00000080&
      Caption         =   "Case ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   6720
      TabIndex        =   3
      Top             =   2280
      Width           =   2655
      Begin VB.OptionButton optVocative 
         BackColor       =   &H00000080&
         Caption         =   "Vocative"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   370
         Left            =   240
         TabIndex        =   20
         Top             =   2880
         Width           =   2175
      End
      Begin VB.OptionButton optAblative 
         BackColor       =   &H00000080&
         Caption         =   "Ablative"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   370
         Left            =   240
         TabIndex        =   19
         Top             =   2400
         Width           =   2175
      End
      Begin VB.OptionButton optAccusative 
         BackColor       =   &H00000080&
         Caption         =   "Accusative"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   370
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   2175
      End
      Begin VB.OptionButton optDative 
         BackColor       =   &H00000080&
         Caption         =   "Dative"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   370
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton optGenitive 
         BackColor       =   &H00000080&
         Caption         =   "Genitive"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   370
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   1935
      End
      Begin VB.OptionButton optNominative 
         BackColor       =   &H00000080&
         Caption         =   "Nominative"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   370
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame fraGender 
      BackColor       =   &H00000080&
      Caption         =   "Gender ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   720
      TabIndex        =   2
      Top             =   2280
      Width           =   2655
      Begin VB.OptionButton optNeuter 
         BackColor       =   &H00000080&
         Caption         =   "Neuter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   370
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton optMasculine 
         BackColor       =   &H00000080&
         Caption         =   "Masculine"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   370
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   1935
      End
      Begin VB.OptionButton optFeminine 
         BackColor       =   &H00000080&
         Caption         =   "Feminine"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   370
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.PictureBox picNoun 
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
      Height          =   615
      Left            =   720
      ScaleHeight     =   555
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Noun Test"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   32.25
         Charset         =   255
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   9840
      TabIndex        =   27
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label lblNB 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: the sentences will rarely make sense, do not translate, just observe case and verb endings."
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
      Height          =   615
      Left            =   5400
      TabIndex        =   26
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label lblSentence 
      BackStyle       =   0  'Transparent
      Caption         =   "As used in the sentence:"
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
      Height          =   255
      Left            =   5520
      TabIndex        =   22
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label lblGrade 
      BackStyle       =   0  'Transparent
      Caption         =   "# Correct / # Tested"
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
      Height          =   255
      Left            =   13080
      TabIndex        =   8
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblNoun 
      BackStyle       =   0  'Transparent
      Caption         =   "Please select the GENDER, NUMBER, and CASE of:"
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
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   4575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   7575
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   16575
   End
End
Attribute VB_Name = "frmTestNouns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Form Level Varaibles used across many of the buttons
'used to keep track of score information
Dim gradeMax As Integer, NumCorrect As Integer, NumWrong As Integer
'Used to store the randomly generated numbers which define many of the other variables
Dim randomNoun As Integer, randomCase As Integer
'Variables used to store data concerning the Noun being tested, many of these are defined random with varaibles above
Dim testCase As String, testGender As Integer, testNumber As String
Dim testStem As String, testEnding As String, testDeclension As Integer
Dim testSentence As String, testNoun As String
'Used to store the answers of the user for comparison with the test variables above
Dim answerCase As String, answerGender As String, answerNumber As String
'Booleans to check for the appropriateness and availability of nouns with respect to the student's level
Dim good As Boolean, nounMatch As Boolean
'Used for easy of reading and for storing a string equivalent for a number varaible
Dim GenderLetter As String


'Public Subroutine which dtermines the noun and the case and the number to be tested, ensures that it is appropriate for the student

Public Sub Testing()
    'Used for searching arrays to determine matches for things
    Dim pos As Integer
    'Randomizes the random number generator
    Randomize
    'defines the initial random numbers, one to select a noun from the NounArrays, and another to determine a case and concurrently a number for the noun in the NounCasesArrays
    randomNoun = Int((NounCtr - 1 + 1) * Rnd + 1) 'limited to the number of nouns in the arrays  (i.e. it does not generate a number beyond that of the greatest index of the nounArrays
    randomCase = Int((12 - 1 + 1) * Rnd + 1) 'Limited to the 12 cases in the arrays
    'Initializes variable for use in the subroutine
    good = False
    nounMatch = False
    pos = 0
    'Searches via match and stop to make sure that there are nouns in the arrays which satisfy the user's level
    Do Until nounMatch Or pos = NounCtr
        pos = pos + 1
        If NounDifficulty(pos) <= StudentLevel Then 'checks that there is at least one noun which is less than or equal to the student's level
            nounMatch = True
        End If
    Loop
    'If there is at least one noun then check and make sure that the noun is appropriate for the current user
    If nounMatch Then
        Do Until good
            If NounDifficulty(randomNoun) > StudentLevel Then 'if the noundifficulty of the random number generated is greater than the student level then generate a new number
                randomNoun = Int((NounCtr - 1 + 1) * Rnd + 1)
            Else
                good = True
                
            End If
        Loop
    Else 'Lets the user know that there are no nouns which satisfy the class level requirments and exits the subroutine
        MsgBox "There are no nouns which fall within range of your class level. Please Contact your administrator to remedy this situation."
        Exit Sub
    End If
    'Determine what gender the random case has returned (ordered in the textfile and arrays) and sets the testNumber Varaible with a string value accordingly
    Select Case randomCase
        Case 1 To 6
            testNumber = "Singular"
        Case 7 To 12
            testNumber = "Plural"
        Case Else
            MsgBox "Out of Range, Alert your Administrator" ' if for some reason there are more cases than expected
            Exit Sub
    End Select
    
    'Determines which case the randomCase has returned and sets the testCase to this string title
    Select Case randomCase
        Case 1, 7
            testCase = "Nominative"
        Case 2, 8
            testCase = "Genitive"
        Case 3, 9
            testCase = "Dative"
        Case 4, 10
            testCase = "Accusative"
        Case 5, 11
            testCase = "Ablative"
        Case 6, 12
            testCase = "Vocative"
        Case Else
            MsgBox "Out of Range, Alert your Administrator"
            Exit Sub
    End Select
    'Gets the gender, stem and declension of the nouns( data which is already stored in arrays in the proper fashion) based on the randomNoun varaible and stores them to appropriate variables
    testGender = GenderNoun(randomNoun)
    testStem = stemNoun(randomNoun)
    testDeclension = DeclensionNoun(randomNoun)
    'Uses the declension of the noun and the random case and the gender to determine which case ending will be used
    If testDeclension = 1 Then
        testEnding = First(randomCase)
    ElseIf testDeclension = 2 Then
        If testGender = 2 Then
            testEnding = SecondM(randomCase)
        Else
            testEnding = SecondN(randomCase)
        End If
    ElseIf testDeclension = 3 Then
        If testGender = 1 Or testGender = 2 Then
            testEnding = ThirdMandF(randomCase)
        Else
            testEnding = ThirdN(randomCase)
        End If
    ElseIf testDeclension = 4 Then
        If testGender = 2 Then
            testEnding = FourthM(randomCase)
        Else
            testEnding = FourthN(randomCase)
        End If
    Else
        testEnding = Fifth(randomCase)
    End If
    'Gives the genderLetter variable an appropriate string Varaible for use later
    Select Case testGender
        Case 1
            GenderLetter = "F."
        Case 2
            GenderLetter = "M."
        Case 3
            GenderLetter = "N."
    End Select
    'Clears the display picture boxes for the display of the nouns and sentences to be generated
    picNoun.Cls
    picSentence.Cls
    
    'Case to generate the proper noun and corresponding sentence and display it in the picboxes
    Select Case randomCase
        'Tests the case
        Case 1
            picNoun.Print NomSNoun(randomNoun) 'Uses already stored values for some cases and generated nouns for others
            testNoun = NomSNoun(randomNoun) 'Stores the complete noun in testNoun for display purposes
            Select Case testGender
                'Tests the gender of the noun and generates an appropriate sentence
                Case 1
                    picSentence.Print NomSNoun(randomNoun) & " longa amat pecuniam "
                Case 2
                    picSentence.Print NomSNoun(randomNoun) & " longus amat pecuniam"
                Case 3
                    picSentence.Print NomSNoun(randomNoun) & " longum amat pecuniam"
            End Select
        Case 2
            picNoun.Print GenSNoun(randomNoun)
            testNoun = GenSNoun(randomNoun)
            Select Case testGender
                Case 1
                    picSentence.Print "amat pecuniam " & GenSNoun(randomNoun) & " longae"
                Case 2
                    picSentence.Print "amat pecuniam " & GenSNoun(randomNoun) & " longi"
                Case 3
                    picSentence.Print "amat pecuniam " & GenSNoun(randomNoun) & " longi"
            End Select
        Case 3
            picNoun.Print stemNoun(randomNoun) & testEnding 'Uses the stem of the noun and the random Test ending to concactinate an appropriate noun for testing
            testNoun = stemNoun(randomNoun) & testEnding
            Select Case testGender
                Case 1
                    picSentence.Print "dat pecuniam " & stemNoun(randomNoun) & testEnding & " longae"
                Case 2
                    picSentence.Print "dat pecuniam " & stemNoun(randomNoun) & testEnding & " longo"
                Case 3
                    picSentence.Print "dat pecuniam " & stemNoun(randomNoun) & testEnding & " longo"
            End Select
        Case 4
            Select Case testGender
                Case 1
                    picSentence.Print "pecunia amat " & stemNoun(randomNoun) & testEnding & " longam"
                    picNoun.Print stemNoun(randomNoun) & testEnding
                    testNoun = stemNoun(randomNoun) & testEnding
                Case 2
                    picSentence.Print "pecunia amat " & stemNoun(randomNoun) & testEnding & " longum"
                    picNoun.Print stemNoun(randomNoun) & testEnding
                    testNoun = stemNoun(randomNoun) & testEnding
                Case 3
                    picSentence.Print "pecunia amat " & NomSNoun(randomNoun) & " longum"
                    picNoun.Print NomSNoun(randomNoun)
                    testNoun = NomSNoun(randomNoun)
            End Select
        Case 5
            picNoun.Print stemNoun(randomNoun) & testEnding
            testNoun = stemNoun(randomNoun) & testEnding
            Select Case testGender
                Case 1
                    picSentence.Print "pecunia amatur " & stemNoun(randomNoun) & testEnding & " longa"
                Case 2
                    picSentence.Print "pecunia amatur " & stemNoun(randomNoun) & testEnding & " longo"
                Case 3
                    picSentence.Print "pecunia amatur " & stemNoun(randomNoun) & testEnding & " longo"
            End Select
        Case 6
            Select Case testGender
                Case 1
                    picSentence.Print NomSNoun(randomNoun) & " ama pecuniam"
                    picNoun.Print NomSNoun(randomNoun)
                    testNoun = NomSNoun(randomNoun)
                Case 2 ' Tests a special case of second declension masculine nouns and generates the appriopriate noun/Sentence
                    If testDeclension = 2 Then
                        picSentence.Print stemNoun(randomNoun) & testEnding & " ama pecuniam"
                        picNoun.Print stemNoun(randomNoun) & testEnding
                        testNoun = stemNoun(randomNoun) & testEnding
                    Else
                        picSentence.Print NomSNoun(randomNoun) & " ama pecuniam"
                        picNoun.Print NomSNoun(randomNoun)
                        testNoun = NomSNoun(randomNoun)
                    End If
                Case 3
                    picSentence.Print NomSNoun(randomNoun) & " ama pecuniam"
                    picNoun.Print NomSNoun(randomNoun)
                    testNoun = NomSNoun(randomNoun)
            End Select
        Case 7
            picNoun.Print stemNoun(randomNoun) & testEnding
            testNoun = stemNoun(randomNoun) & testEnding
            Select Case testGender
                Case 1
                    picSentence.Print stemNoun(randomNoun) & testEnding & " longae amant pecuniam"
                Case 2
                    picSentence.Print stemNoun(randomNoun) & testEnding & " longi amant pecuniam"
                Case 3
                    picSentence.Print stemNoun(randomNoun) & testEnding & " longa amant pecuniam"
            End Select
        Case 8
            picNoun.Print stemNoun(randomNoun) & testEnding
            testNoun = stemNoun(randomNoun) & testEnding
            Select Case testGender
                Case 1
                    picSentence.Print "amat pecuniam " & stemNoun(randomNoun) & testEnding & " longarum"
                Case 2
                    picSentence.Print "amat pecuniam " & stemNoun(randomNoun) & testEnding & " longorum"
                Case 3
                    picSentence.Print "amat pecuniam " & stemNoun(randomNoun) & testEnding & " longorum"
            End Select
        Case 9
            picNoun.Print stemNoun(randomNoun) & testEnding
            testNoun = stemNoun(randomNoun) & testEnding
            Select Case testGender
                Case 1
                    picSentence.Print "dat pecuniam " & stemNoun(randomNoun) & testEnding & " longis"
                Case 2
                    picSentence.Print "dat pecuniam " & stemNoun(randomNoun) & testEnding & " longis"
                Case 3
                    picSentence.Print "dat pecuniam " & stemNoun(randomNoun) & testEnding & " longis"
            End Select
        Case 10
            picNoun.Print stemNoun(randomNoun) & testEnding
            testNoun = stemNoun(randomNoun) & testEnding
            Select Case testGender
                Case 1
                    picSentence.Print "pecunia amat " & stemNoun(randomNoun) & testEnding & " longas"
                Case 2
                    picSentence.Print "pecunia amat " & stemNoun(randomNoun) & testEnding & " longos"
                Case 3
                    picSentence.Print "pecunia amat " & stemNoun(randomNoun) & testEnding & " longa"
            End Select
        Case 11
            picNoun.Print stemNoun(randomNoun) & testEnding
            testNoun = stemNoun(randomNoun) & testEnding
            Select Case testGender
                Case 1
                    picSentence.Print "pecunia amatur " & stemNoun(randomNoun) & testEnding & " longis"
                Case 2
                    picSentence.Print "pecunia amatur " & stemNoun(randomNoun) & testEnding & " longis"
                Case 3
                    picSentence.Print "pecunia amatur " & stemNoun(randomNoun) & testEnding & " longis"
            End Select
        Case 12
            picNoun.Print stemNoun(randomNoun) & testEnding
            testNoun = stemNoun(randomNoun) & testEnding
            Select Case testGender
                Case 1
                    picSentence.Print stemNoun(randomNoun) & testEnding & " amate pecuniam!"
                Case 2
                    picSentence.Print stemNoun(randomNoun) & testEnding & " amate pecuniam!"
                Case 3
                    picSentence.Print stemNoun(randomNoun) & testEnding & " amate pecuniam!"
            End Select
        Case Else
            MsgBox "Whoops how did you make a mistake" 'returns an error if possible I missed something
    End Select
    
End Sub


Private Sub cmdEnd_Click()
    'code to end the testing session store appropriate data and clear some text boxes
    'Makes visible and enables/disables certain buttons
    cmdEnd.Visible = False
    cmdStart.Visible = True
    cmdLogOut.Enabled = True
    cmdReturn.Enabled = True
    cmdSubmit.Enabled = False
    'calls the public subroutine for calculating, altering/generating, the student global score for the whole program (cf. mdlPublicSubs)
    Call CalculateGrade(NumCorrect, NumWrong, gradeMax)
    'Clears picture boxes (leaves picNounsTested and picGrade alone for reference purposes until a new session is started)
    picNoun.Cls
    picSentence.Cls
End Sub

Private Sub cmdLogOut_Click()
    'Logouts the user
    frmTestNouns.Hide
    'cf. mdlPublicSubs
    Call LogOut
End Sub

Private Sub cmdReturn_Click()
    'Returns the user to the options page
    frmTestNouns.Hide
    frmOptionsPage.Show
End Sub

Private Sub cmdStart_Click()
    'this button starts the testing session
    'Enables, and makes visible appropriate buttons (end is disabled for the first round in order to avoid errors in the program [especially with grading])
    cmdEnd.Visible = True
    cmdStart.Visible = False
    cmdLogOut.Enabled = False
    cmdReturn.Enabled = False
    cmdSubmit.Enabled = True
    'initializes varaibles
    gradeMax = 0
    NumWrong = 0
    NumCorrect = 0
    'clears and prepares the picBoxes
    picGrade.Cls
    picGrade.Print NumCorrect & "/" & gradeMax
    picNoun.Cls
    picSentence.Cls
    picNounsTested.Cls
    'Calls the pulic subroutine testing cf. above
    Call Testing
    'Prepares more picBoxes
    picNounsTested.Print "Noun"; Tab(20); "Form-Tested"; Tab(40); "Form-Guessed"; Tab(60); "Correct"
    picNounsTested.Print "******************************************************************************************************************************************"
End Sub

Private Sub cmdSubmit_Click()
    'This button test the answer given by user, stores the result in form-level varaibles and displays a new noun
    'Local variables for display purposes
    Dim answerCaseName As String
    Dim answerNumberName As String
    Dim answerGenderName As String
    'Gets the answer from the user concerning the case
    If optNominative.Value = True Then
        answerCase = "Nominative" 'Gives a testable value to answerCase
        answerCaseName = "Nom" 'Gives a display name to the selection
        optNominative.Value = False 'resets the optionButton
    ElseIf optGenitive.Value = True Then
        answerCase = "Genitive"
        answerCaseName = "Gen"
        optGenitive.Value = False
    ElseIf optDative.Value = True Then
        answerCase = "Dative"
        answerCaseName = "Dat"
        optDative.Value = False
    ElseIf optAccusative.Value = True Then
        answerCase = "Accusative"
        answerCaseName = "Acc"
        optAccusative.Value = False
    ElseIf optAblative.Value = True Then
        answerCase = "Ablative"
        answerCaseName = "Abl"
        optAblative.Value = False
    ElseIf optVocative.Value = True Then
        answerCase = "Vocative"
        answerCaseName = "Voc"
        optVocative.Value = False
    Else
        MsgBox "Please Select a Case." 'Ensures the user selects a case
        Exit Sub
    End If
    'Used to test the number of the user's slection
    If optSingular.Value = True Then
        answerNumber = "Singular" 'gives variable a testable value
        answerNumberName = "S" 'gives a display value
        optSingular.Value = False ' resets the optionButton
    ElseIf optPlural.Value = True Then
        answerNumber = "Plural"
        answerNumberName = "P"
        optPlural.Value = False
    Else
        MsgBox "Please select a Number." 'Ensures that the user selects a number
        Exit Sub
    End If
    'defines the user's selection for gender
    If optFeminine.Value = True Then
        answerGender = 1 ' sets a testable value for the user's answer
        answerGenderName = "F." ' gives a display name to the answer
        optFeminine.Value = False ' resets the option button
    ElseIf optMasculine.Value = True Then
        answerGender = 2
        answerGenderName = "M."
        optMasculine.Value = False
    ElseIf optNeuter.Value = True Then
        answerGender = 3
        answerGenderName = "N."
        optNeuter.Value = False
    Else
        MsgBox "Please Select a Gender." ' Ensures that the user selects a gender
        Exit Sub
    End If
    'Tests user answer against the generated test
    If testCase = answerCase And testNumber = answerNumber And testGender = answerGender Then 'if all criteria match, displays as much and adds one to the number correct running total
        picNounsTested.Print testNoun; Tab(20); formName(randomCase) & " " & GenderLetter; Tab(40); answerCaseName & answerNumberName & " " & answerGenderName; Tab(60); "Yes"
        NumCorrect = NumCorrect + 1
    Else 'If not diaplays as much and adds one to the numWrong running total
        picNounsTested.Print testNoun; Tab(20); formName(randomCase) & " " & GenderLetter; Tab(40); answerCaseName & answerNumberName & " " & answerGenderName; Tab(60); "No"
        NumWrong = NumWrong + 1
    End If
    'Adds one to the Max grade for percentaging purposes
    gradeMax = gradeMax + 1
    'Clears and displays new information
    picGrade.Cls
    picGrade.Print NumCorrect & "/" & gradeMax
    
    'Enables the end button so the user may end if he/she so desires
    cmdEnd.Enabled = True
    'Calls the testing subroutine and generates a new noun and ending (cf. above)
    Call Testing
    
End Sub
