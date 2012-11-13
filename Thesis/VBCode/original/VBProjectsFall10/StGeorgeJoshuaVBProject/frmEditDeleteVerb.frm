VERSION 5.00
Begin VB.Form frmEditDeleteVerb 
   BackColor       =   &H00000080&
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13335
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   13335
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00000080&
      Caption         =   "Delete Verb"
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3480
      Width           =   2775
   End
   Begin VB.TextBox txtPresentStem 
      Height          =   375
      Left            =   600
      TabIndex        =   23
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox txtInfinitive 
      Height          =   375
      Left            =   3240
      TabIndex        =   22
      Top             =   240
      Width           =   2175
   End
   Begin VB.TextBox txtPerfectStem 
      Height          =   405
      Left            =   5760
      TabIndex        =   21
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox txtParticipleStem 
      Height          =   375
      Left            =   8160
      TabIndex        =   20
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox txtDefinition 
      Height          =   405
      Left            =   10680
      TabIndex        =   19
      Top             =   240
      Width           =   2175
   End
   Begin VB.Frame fraConjugation 
      BackColor       =   &H00000080&
      Caption         =   "Select a Conjugation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   4200
      TabIndex        =   13
      Top             =   2640
      Width           =   2535
      Begin VB.OptionButton optFirst 
         BackColor       =   &H00000080&
         Caption         =   "First "
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
         TabIndex        =   18
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton optSecond 
         BackColor       =   &H00000080&
         Caption         =   "Second"
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
         TabIndex        =   17
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton optThird 
         BackColor       =   &H00000080&
         Caption         =   "Third"
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
         Top             =   1320
         Width           =   1815
      End
      Begin VB.OptionButton optThirdIO 
         BackColor       =   &H00000080&
         Caption         =   "Third - IO"
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
         TabIndex        =   15
         Top             =   1800
         Width           =   1815
      End
      Begin VB.OptionButton optFourth 
         BackColor       =   &H00000080&
         Caption         =   "Fourth"
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
         Top             =   2280
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdSubmit 
      BackColor       =   &H00000080&
      Caption         =   "Submit Verb"
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2640
      Width           =   2775
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Frame fraDificultyLevel 
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
      Height          =   3375
      Left            =   9960
      TabIndex        =   6
      Top             =   2640
      Width           =   2535
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
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   735
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
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1575
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
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1095
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
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   1695
      End
   End
   Begin VB.Frame fraSpecial 
      BackColor       =   &H00000080&
      Caption         =   "Select a Verb Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   7080
      TabIndex        =   1
      Top             =   2640
      Width           =   2535
      Begin VB.OptionButton optRegular 
         BackColor       =   &H00000080&
         Caption         =   "Regular"
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
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optDeponent 
         BackColor       =   &H00000080&
         Caption         =   "Deponent"
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
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton optSemiDeponent 
         BackColor       =   &H00000080&
         Caption         =   "Semi-Deponent"
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
         TabIndex        =   3
         Top             =   1320
         Width           =   1695
      End
      Begin VB.OptionButton optDefective 
         BackColor       =   &H00000080&
         Caption         =   "Defective"
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
         TabIndex        =   2
         Top             =   1800
         Width           =   1215
      End
   End
   Begin VB.TextBox txtPrincipleParts 
      Height          =   405
      Left            =   600
      TabIndex        =   0
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Label lblFirstStem 
      BackStyle       =   0  'Transparent
      Caption         =   "Present Stem                  Example: amo = am ; ago = ag"
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
      Left            =   600
      TabIndex        =   32
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label lblInfinitive 
      BackStyle       =   0  'Transparent
      Caption         =   "Infinitive (2nd conjugation nouns are indicated with e^ as in mone^re)"
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
      Height          =   855
      Left            =   3240
      TabIndex        =   31
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label lblPerfectStem 
      BackStyle       =   0  'Transparent
      Caption         =   "Perfect Stem              Example: amavi =amav ; egi = egi"
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
      Left            =   5760
      TabIndex        =   30
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblPartStem 
      BackStyle       =   0  'Transparent
      Caption         =   "Participle Stem          Example: amatus = amat ; actus = act"
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
      Height          =   735
      Left            =   8160
      TabIndex        =   29
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label lblDefinition 
      BackStyle       =   0  'Transparent
      Caption         =   "Definition"
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
      Left            =   10680
      TabIndex        =   28
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label lblPrincipleParts 
      BackStyle       =   0  'Transparent
      Caption         =   "Principle Parts"
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
      Left            =   600
      TabIndex        =   27
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   135
      Left            =   1200
      TabIndex        =   26
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label lblSecondConjugation 
      BackStyle       =   0  'Transparent
      Caption         =   "(long vowel souns are indicated like e^, as in mone^o, mone^re)"
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
      Left            =   2040
      TabIndex        =   25
      Top             =   2160
      Width           =   5535
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Verbs"
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
      Left            =   9000
      TabIndex        =   24
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   6495
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "frmEditDeleteVerb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLogOut_Click()
    'Logs User out to the login Page
    frmAddVerbs.Hide
    Call LogOut
End Sub

Private Sub cmdQuit_Click()
    'Quits the program
    End
End Sub

Private Sub cmdDelete_Click()
    'Deletes a given verb
    'Declares useful variables
    Dim pos As Integer
    Dim verify As String
    Dim verb As String
    'Sets teh verbname as the infinitive of the verb to be deleted
    verb = VerbInfinitive(VerbPosition)
    'gets input fom user about whether or not he/she want to delete the verb
    verify = InputBox("You are about to delete the verb" & UCase(verb) & ". Do you wish to continue? If yes type 'yes' if not type 'no'.")
    'if the user indicates yes then
    If LCase(verify) = "yes" Then
        'Loops over the array beginning at the verb to be deleted and ending just short of the array, replacing the previous with the next
        For pos = VerbPosition To verbCtr - 1
            VerbPresStem(pos) = VerbPresStem(pos + 1)
            VerbInfinitive(pos) = VerbInfinitive(pos + 1)
            VerbPerfStem(pos) = VerbPerfStem(pos + 1)
            VerbPartStem(pos) = VerbPartStem(pos + 1)
            VerbDefinition(pos) = VerbDefinition(pos + 1)
            VerbConjugation(pos) = VerbConjugation(pos + 1)
            VerbDifficulty(pos) = VerbDifficulty(pos + 1)
            VerbClass(pos) = VerbClass(pos + 1)
            VerbPrincipleParts(pos) = VerbPrincipleParts(pos + 1)
        Next pos
        'resets the verbctr varaible with the appropriate number of verbs
        verbCtr = verbCtr - 1
        'opens the verbs text fiel for output in order to rewrtie the
        Open App.Path & "\Data\Verbs.txt" For Output As #1
            For pos = 1 To verbCtr
                Write #1, VerbPresStem(pos), VerbInfinitive(pos), VerbPerfStem(pos), VerbPartStem(pos), VerbDefinition(pos), VerbConjugation(pos), VerbDifficulty(pos), VerbClass(pos), VerbPrincipleParts(pos)
            Next pos
        Close #1
        MsgBox UCase(verb) & " has been deleted."
    Else
        MsgBox UCase(verb) & " has not been deleted."
    End If
    
    Call ReadVerbs
    frmEditDeleteVerb.Hide
    frmClassVocab.Show
End Sub

Private Sub cmdReturn_Click()
    'Resturns the user to the Class Vocab
    frmAddVerbs.Hide
    frmClassVocab.Show
End Sub

Private Sub cmdSubmit_Click()
    'writes a new verb to the Verbs text file, ensure that all selections are filled
    'Defines the variables used for the button
    Dim newPresStem As String
    Dim newInfinitive As String
    Dim newPerfStem As String
    Dim newPartStem As String
    Dim newDefinition As String
    Dim newPrincipleParts As String
    Dim Conjugation As Integer
    Dim conjugationName As String
    Dim Difficulty As Integer
    Dim DifficultyName As String
    Dim VerbType As Integer
    Dim VerbTypeName As String
    Dim verify As String
    Dim pos As Integer
    'Gets user input from text boxes
    newPresStem = txtPresentStem.Text
    newInfinitive = txtInfinitive.Text
    newPerfStem = txtPerfectStem.Text
    newPartStem = txtParticipleStem.Text
    newDefinition = txtDefinition.Text
    newPrincipleParts = txtPrincipleParts.Text
    'Gets user input from booleans, testing to make sure that all of them are filled, and then resets them for future use
    If optFirst = True Then
        Conjugation = 1 'Gives the conjugation variable an integer value corresponding to which conjugation it is
        conjugationName = "First" 'Gives the conjugationName variable a string value corresponding to which conjugation it is
        optFirst.Value = False 'Resets the boolean which was true
    ElseIf optSecond = True Then
        Conjugation = 2
        conjugationName = "Second"
        optSecond.Value = False
    ElseIf optThird = True Then
        Conjugation = 3
        conjugationName = "Third"
        optThird.Value = False
    ElseIf optThirdIO = True Then
        Conjugation = 4
        conjugationName = "Third-IO"
        optThirdIO.Value = False
    ElseIf optFourth = True Then
        Conjugation = 5
        conjugationName = "Fourth"
        optFourth.Value = False
    Else
        'Handles the situation where no option button is selected
        MsgBox "Please select a conjugation"
        Exit Sub
    End If
    
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
    
    If optRegular.Value = True Then
        VerbType = 1
        VerbTypeName = "Regular"
    ElseIf optDeponent.Value = True Then
        VerbType = 2
        VerbTypeName = "Deponent"
        optDeponent.Value = False
        optRegular.Value = True
    ElseIf optSemiDeponent.Value = True Then
        VerbType = 3
        VerbTypeName = "Semi-Deponent"
        optSemiDeponent.Value = False
        optRegular.Value = True
    Else
        VerbType = 4
        VerbTypeName = "Defective"
        optDefective.Value = False
        optRegular.Value = True
    End If
    
    'Reads the new information into the text file and returns an error if any field was left blank
    If newPresStem = "" Or newInfinitive = "" Or newPerfStem = "" Or newPartStem = "" Or newDefinition = "" Or newPrincipleParts = "" Then
        MsgBox "Please ensure that all fields are filled in"
    Else
        verify = InputBox("You are about to edit the " & UCase(conjugationName) & " conjugation " & UCase(VerbTypeName) & " verb, " & UCase(newInfinitive) & ", meaning " & UCase(newDefinition) & " with a present stem of " & UCase(newPresStem) & ", a perfect stem of " & UCase(newPerfStem) & ", and a participle stem of " & UCase(newPartStem) & " with a difficulty tag of: " & UCase(DifficultyName) & "and principle parts: " & UCase(newPrincipleParts) & ". If this is correct enter 'yes', if not enter 'no.'")
        If LCase(verify) = "yes" Then
            'Gives the verb which is to be edited the new values which the user designated and stores them across all parralel arrays
            VerbPresStem(VerbPosition) = newPresStem
            VerbInfinitive(VerbPosition) = newInfinitive
            VerbPerfStem(VerbPosition) = newPerfStem
            VerbPartStem(VerbPosition) = newPartStem
            VerbDefinition(VerbPosition) = newDefinition
            VerbConjugation(VerbPosition) = Conjugation
            VerbDifficulty(VerbPosition) = Difficulty
            VerbClass(VerbPosition) = VerbType
            VerbPrincipleParts(VerbPosition) = newPrincipleParts
            'Opens the verbs text file for output to rewrite the entire text file
            Open App.Path & "\data\Verbs.txt" For Output As #1
                'loops writing each of the arrays to a blank text file
                For pos = 1 To verbCtr
                    Write #1, VerbPresStem(pos), VerbInfinitive(pos), VerbPerfStem(pos), VerbPartStem(pos), VerbDefinition(pos), VerbConjugation(pos), VerbDifficulty(pos), VerbClass(pos), VerbPrincipleParts(pos)
                Next pos
            Close #1
            MsgBox UCase(newInfinitive) & " has been successfully edited"
        Else
            MsgBox UCase(newInfinitive) & " will not be edited"
            Exit Sub
        End If
    End If
    'Resets the text fields
    txtPresentStem.Text = ""
    txtInfinitive.Text = ""
    txtPerfectStem.Text = ""
    txtParticipleStem.Text = ""
    txtDefinition.Text = ""
    txtPrincipleParts.Text = ""
    Call ReadVerbs
    frmEditDeleteVerb.Hide
    frmClassVocab.Show
End Sub

