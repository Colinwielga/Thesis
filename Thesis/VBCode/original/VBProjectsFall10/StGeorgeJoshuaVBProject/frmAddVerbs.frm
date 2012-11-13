VERSION 5.00
Begin VB.Form frmAddVerbs 
   BackColor       =   &H00000080&
   Caption         =   "Lingu Vivens- Admin Options -Add Verbs"
   ClientHeight    =   6975
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13560
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   13560
   Begin VB.TextBox txtPrincipleParts 
      Height          =   405
      Left            =   720
      TabIndex        =   30
      Top             =   1920
      Width           =   4215
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
      Left            =   7200
      TabIndex        =   25
      Top             =   2880
      Width           =   2535
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
         TabIndex        =   29
         Top             =   1800
         Width           =   1215
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
         TabIndex        =   28
         Top             =   1320
         Width           =   1695
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
         TabIndex        =   27
         Top             =   840
         Width           =   1335
      End
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
         TabIndex        =   26
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
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
      Left            =   10080
      TabIndex        =   20
      Top             =   2880
      Width           =   2535
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
         TabIndex        =   24
         Top             =   1800
         Width           =   1695
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
         TabIndex        =   23
         Top             =   1320
         Width           =   1095
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
         TabIndex        =   22
         Top             =   840
         Width           =   1575
      End
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
         TabIndex        =   21
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00808080&
      Caption         =   "Quit"
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
      TabIndex        =   19
      Top             =   5640
      Width           =   2775
   End
   Begin VB.CommandButton cmdLogOut 
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4800
      Width           =   2775
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00808080&
      Caption         =   "Return to Administration Options"
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
      TabIndex        =   17
      Top             =   3960
      Width           =   2775
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2880
      Width           =   2775
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
      Left            =   4320
      TabIndex        =   10
      Top             =   2880
      Width           =   2535
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
         TabIndex        =   15
         Top             =   2280
         Width           =   975
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
         TabIndex        =   14
         Top             =   1800
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
         TabIndex        =   13
         Top             =   1320
         Width           =   1815
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
         TabIndex        =   12
         Top             =   840
         Width           =   1815
      End
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
         TabIndex        =   11
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.TextBox txtDefinition 
      Height          =   405
      Left            =   10800
      TabIndex        =   4
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtParticipleStem 
      Height          =   375
      Left            =   8280
      TabIndex        =   3
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox txtPerfectStem 
      Height          =   405
      Left            =   5880
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtInfinitive 
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtPresentStem 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Add Verbs"
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
      Left            =   9120
      TabIndex        =   34
      Top             =   1800
      Width           =   3015
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
      Left            =   2160
      TabIndex        =   33
      Top             =   2400
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   135
      Left            =   1320
      TabIndex        =   32
      Top             =   1920
      Width           =   135
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
      Left            =   720
      TabIndex        =   31
      Top             =   2400
      Width           =   2295
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
      Left            =   10800
      TabIndex        =   9
      Top             =   960
      Width           =   2175
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
      Left            =   8280
      TabIndex        =   8
      Top             =   960
      Width           =   2055
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
      Left            =   5880
      TabIndex        =   7
      Top             =   960
      Width           =   1935
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
      Left            =   3360
      TabIndex        =   6
      Top             =   960
      Width           =   2175
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
      Left            =   720
      TabIndex        =   5
      Top             =   960
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   6495
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   13335
   End
End
Attribute VB_Name = "frmAddVerbs"
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

Private Sub cmdReturn_Click()
    'Resturns the user to the Admin Pane
    frmAddVerbs.Hide
    frmAdmin.Show
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
        verify = InputBox("You are about to add the " & UCase(conjugationName) & " conjugation " & UCase(VerbTypeName) & " verb, " & UCase(newInfinitive) & ", meaning " & UCase(newDefinition) & " with a present stem of " & UCase(newPresStem) & ", a perfect stem of " & UCase(newPerfStem) & ", and a participle stem of " & UCase(newPartStem) & "with a difficulty tag of: " & UCase(DifficultyName) & "and principle parts:" & UCase(newPrincipleParts) & ". If this is correct enter 'yes', if not enter 'no.'")
        If LCase(verify) = "yes" Then
            Open App.Path & "\data\Verbs.txt" For Append As #1
                Write #1, newPresStem, newInfinitive, newPerfStem, newPartStem, newDefinition, Conjugation, Difficulty, VerbType, newPrincipleParts
            Close #1
            MsgBox UCase(newInfinitive) & " has been added to the verb list"
        Else
            MsgBox UCase(newInfinitive) & " will not be added to the verb list"
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
End Sub


