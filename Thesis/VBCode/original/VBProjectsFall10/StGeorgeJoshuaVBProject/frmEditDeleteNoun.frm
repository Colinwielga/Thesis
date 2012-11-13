VERSION 5.00
Begin VB.Form frmEditDeleteNoun 
   BackColor       =   &H00000080&
   Caption         =   "Form1"
   ClientHeight    =   6315
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13155
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   13155
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00000080&
      Caption         =   "Delete Noun"
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
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Frame fraGender 
      BackColor       =   &H00000080&
      Caption         =   "Choose a Gender"
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
      Left            =   6720
      TabIndex        =   17
      Top             =   1920
      Width           =   2655
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
         Height          =   495
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   1455
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
         Height          =   495
         Left            =   240
         TabIndex        =   19
         Top             =   1200
         Width           =   1695
      End
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
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   1695
      End
   End
   Begin VB.Frame fraDeclensions 
      BackColor       =   &H00000080&
      Caption         =   "Choose a Declension"
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
      Left            =   3720
      TabIndex        =   11
      Top             =   1920
      Width           =   2655
      Begin VB.OptionButton optFirst 
         BackColor       =   &H00000080&
         Caption         =   "First Declension"
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
         TabIndex        =   16
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton optSecond 
         BackColor       =   &H00000080&
         Caption         =   "Second Declenion"
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
         TabIndex        =   15
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton optThird 
         BackColor       =   &H00000080&
         Caption         =   "Third Declension"
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
         TabIndex        =   14
         Top             =   1560
         Width           =   2295
      End
      Begin VB.OptionButton optFourth 
         BackColor       =   &H00000080&
         Caption         =   "Fourth Declension"
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
         Top             =   2160
         Width           =   2295
      End
      Begin VB.OptionButton optFifth 
         BackColor       =   &H00000080&
         Caption         =   "Fifth Declension"
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
         Top             =   2640
         Width           =   2295
      End
   End
   Begin VB.TextBox txtNom 
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtGen 
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtStem 
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton cmdAddNoun 
      BackColor       =   &H00000080&
      Caption         =   "Submit Noun"
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
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox txtDefinition 
      Height          =   375
      Left            =   7680
      TabIndex        =   6
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00808080&
      Caption         =   "Return to Class Vocab"
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
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Frame fraDifficulty 
      BackColor       =   &H00000080&
      Caption         =   "Select a Difficulty Level"
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
      Left            =   9720
      TabIndex        =   0
      Top             =   1920
      Width           =   2655
      Begin VB.OptionButton optEasy 
         BackColor       =   &H00000080&
         Caption         =   "Easy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton optIntermediate 
         BackColor       =   &H00000080&
         Caption         =   "Intermediate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   240
         TabIndex        =   3
         Top             =   980
         Width           =   1935
      End
      Begin VB.OptionButton optDifficult 
         BackColor       =   &H00000080&
         Caption         =   "Difficult"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   240
         TabIndex        =   2
         Top             =   1480
         Width           =   1935
      End
      Begin VB.OptionButton optCollegeLevel 
         BackColor       =   &H00000080&
         Caption         =   "College Level"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   240
         TabIndex        =   1
         Top             =   1980
         Width           =   1935
      End
   End
   Begin VB.Label lblNom 
      BackStyle       =   0  'Transparent
      Caption         =   "Nomintive Singular"
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
      Left            =   720
      TabIndex        =   25
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label lblGen 
      BackStyle       =   0  'Transparent
      Caption         =   "Genitive Singular"
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
      Left            =   3000
      TabIndex        =   24
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label lblStem 
      BackStyle       =   0  'Transparent
      Caption         =   "Noun Stem"
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
      Left            =   5280
      TabIndex        =   23
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lblDefinition 
      BackStyle       =   0  'Transparent
      Caption         =   "Noun Definition"
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
      Left            =   7680
      TabIndex        =   22
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Nouns"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   30
         Charset         =   255
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   9840
      TabIndex        =   21
      Top             =   360
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   6135
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   12975
   End
End
Attribute VB_Name = "frmEditDeleteNoun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddNoun_Click()
    'Button to add a new noun to the data file and checks to make sure that all the information is inputed
    'defines the varaibles
    Dim nomS As String
    Dim genS As String
    Dim stem As String
    Dim Definition As String
    Dim gender As Integer
    Dim Declension As Integer
    Dim DeclensionCheck(1 To 5) As Boolean, decCtr As Integer, decCtr2 As Integer, decFound As Boolean
    Dim genderCheck(1 To 3) As Boolean, genCtr As Integer, genCtr2 As Integer, genFound As Boolean
    Dim Difficulty As Integer
    Dim DifficultyName As String
    Dim verify As String
    Dim genderName As String
    Dim DeclensionMarker As String
    Dim pos As Integer
    'Initiates variables
    decFound = False
    decCtr = 0
    genFound = False
    genCtr = 0
    
    
        
    'Gets input from user
    nomS = txtNom.Text
    genS = txtGen.Text
    stem = txtStem.Text
    Definition = txtDefinition.Text
    'Reads the declension opt values into an array so it can be searched easily
    DeclensionCheck(1) = optFirst.Value
    DeclensionCheck(2) = optSecond.Value
    DeclensionCheck(3) = optThird.Value
    DeclensionCheck(4) = optFourth.Value
    DeclensionCheck(5) = optFifth.Value
    'reads the gender option button values into an array so it can be searched easily
    genderCheck(1) = optFeminine.Value
    genderCheck(2) = optMasculine.Value
    genderCheck(3) = optNeuter.Value
    'Loops to check if any value is true, if it has it stops and records the position of the true value
    Do Until decFound Or decCtr = 5
        decCtr = decCtr + 1
        If DeclensionCheck(decCtr) = True Then
            decFound = True
        End If
    Loop
    'Checks to see if any true value was found, if not ask the user to select a declension, else store the declension position into the declension variable
    If decFound Then
        Declension = decCtr
    Else
        MsgBox "Please select a declension"
        Exit Sub 'ends the subroutine so user can correct his/her mistake
    End If
    'Loops to check if a gender value has been made true, if it has then it records it position
    Do Until genFound Or genCtr = 3
        genCtr = genCtr + 1
        If genderCheck(genCtr) = True Then
            genFound = True
        End If
    Loop
    'If a true value has been found then stores its position into the  gender varaible, if not it asks the user to select a gender
    If genFound Then
        gender = genCtr
    Else
        MsgBox "Please select a gender"
        Exit Sub 'Ends the sub Routine in order for the user to remedy his/her mistake
    End If
    'Get the information from the Difficulty Option Set and gives appropriate values to the variables
    If optEasy.Value = True Then
        Difficulty = 1
        DifficultyName = "Easy"
        optEasy.Value = False
    ElseIf optIntermediate.Value = True Then
        Difficulty = 2
        DifficultyName = "Intermediate"
        optIntermediate.Value = False
    ElseIf optDifficult.Value = True Then
        Difficulty = 3
        DifficultyName = "Difficult"
        optDifficult.Value = False
    ElseIf optCollegeLevel.Value = True Then
        Difficulty = 4
        DifficultyName = "College Level"
        optCollegeLevel = False
    Else
        MsgBox "Please select a difficulty level"
        Exit Sub
    End If
    
    'Gives a string value to a number value
    If gender = 1 Then
        genderName = "feminine"
    ElseIf gender = 2 Then
        genderName = "masculine"
    Else
        genderName = "neuter"
    End If
    'Provides a easier to read message statement for the user
    If Declension = 1 Then
        DeclensionMarker = "1st"
    ElseIf Declension = 2 Then
        DeclensionMarker = "2nd"
    ElseIf Declension = 3 Then
        DeclensionMarker = "3rd"
    Else
        DeclensionMarker = Declension & "th"
    End If
    'Checks to see if any of the fields are empty
    If nomS = "" Or genS = "" Or stem = "" Or Definition = "" Then
        MsgBox "Please make sure that all fields are filled in." 'error message if fields empty
    Else 'first verifies that the user is inputing valid data, and then appends the new noun into the Data File
            verify = InputBox("You are about to edit " & UCase(nomS) & ", " & UCase(genS) & " a " & DeclensionMarker & " declension " & UCase(genderName) & " noun with a stem of " & UCase(stem) & "-,  a definition of " & UCase(Definition) & " and a difficulty level of: " & UCase(DifficultyName) & ". If this is what you wish type 'yes' below, if not type 'no'.")
            If LCase(verify) = "yes" Then 'checks to see if user thinks the information is correct
                'gives the noun arrays for the noun being edited, the updated noun information
                NomSNoun(NounPosition) = nomS
                GenSNoun(NounPosition) = genS
                stemNoun(NounPosition) = stem
                DeclensionNoun(NounPosition) = Declension
                GenderNoun(NounPosition) = gender
                definitionNoun(NounPosition) = Definition
                NounDifficulty(NounPosition) = Difficulty
                
                'Output the entire noun array parralels into a blank text file
                Open App.Path & "\data\Nouns.txt" For Output As #1
                    For pos = 1 To NounCtr
                        Write #1, NomSNoun(pos), GenSNoun(pos), stemNoun(pos), DeclensionNoun(pos), GenderNoun(pos), definitionNoun(pos), NounDifficulty(pos)
                    Next pos
                Close #1
                MsgBox UCase(nomS) & ", " & UCase(genS) & " has been successfuly edited." 'user feedback for completion
            Else
                MsgBox UCase(nomS) & ", " & UCase(genS) & " will not be edited as specified, please make any chnages and try again"
                Exit Sub
            End If
    End If
    'Clears text boxes for future use
    txtNom.Text = ""
    txtGen.Text = ""
    txtStem.Text = ""
    txtDefinition.Text = ""
    'Returns tot eh class vocab form
    frmEditDeleteNoun.Hide
    frmClassVocab.Show
    Call ReadNouns
    
End Sub

Private Sub cmdDelete_Click()
    'used to delete the noun being edited
    'Declares sueful varaibles
    Dim pos As Integer
    Dim verify As String
    Dim noun As String
    'Gives the noun string the value of the nomS and Gens of the nounb eing edited (used solely for ease of use)
    noun = NomSNoun(NounPosition) & ", " & GenSNoun(NounPosition)
    'Asks the user to verify his/her descision to delete the noun
    verify = InputBox("Please verify that you wish to delete " & UCase(noun) & " by inputing 'yes' into the field below, if not enter 'no'")
    'If user wishes to continue
    If LCase(verify) = "yes" Then
        'Loops from the noun being edited to one less than the total length of the array giveign the next aray position to the next
        For pos = NounPosition To NounCtr - 1
            NomSNoun(pos) = NomSNoun(pos + 1)
            GenSNoun(pos) = GenSNoun(pos + 1)
            stemNoun(pos) = stemNoun(pos + 1)
            DeclensionNoun(pos) = DeclensionNoun(pos + 1)
            GenderNoun(pos) = GenderNoun(pos + 1)
            definitionNoun(pos) = definitionNoun(pos + 1)
            NounDifficulty(pos) = NounDifficulty(pos + 1)
        Next pos
        'Redices the number of nouns for the one jsut deleted
        NounCtr = NounCtr - 1
        'Opens and outputs the entirety of the noun array to the noun text file storing the users deletion
        Open App.Path & "\data\Nouns.txt" For Output As #1
            For pos = 1 To NounCtr
                Write #1, NomSNoun(pos), GenSNoun(pos), stemNoun(pos), DeclensionNoun(pos), GenderNoun(pos), definitionNoun(pos), NounDifficulty(pos)
            Next pos
        Close #1

    Else
        MsgBox UCase(noun) & " will not be deleted"
        Exit Sub
    End If
    'Clears the textboxes
    txtNom.Text = ""
    txtGen.Text = ""
    txtStem.Text = ""
    txtDefinition.Text = ""
    'Returns the user to the class vocab form
    frmEditDeleteNoun.Hide
    frmClassVocab.Show
    're-reads the updated text file
    Call ReadNouns
    
End Sub


Private Sub cmdReturn_Click()
    frmAddNouns.Hide
    frmClassVocab.Show
End Sub

