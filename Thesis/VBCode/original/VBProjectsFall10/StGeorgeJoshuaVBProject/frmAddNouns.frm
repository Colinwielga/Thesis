VERSION 5.00
Begin VB.Form frmAddNouns 
   BackColor       =   &H00000080&
   Caption         =   "Lingua Vivens - Admin Options - Add Nouns"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13545
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   13545
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
      Left            =   9960
      TabIndex        =   22
      Top             =   2040
      Width           =   2655
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
         TabIndex        =   26
         Top             =   1980
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
         TabIndex        =   25
         Top             =   1480
         Width           =   1935
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
         TabIndex        =   24
         Top             =   980
         Width           =   1935
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
         Height          =   500
         Left            =   240
         TabIndex        =   23
         Top             =   480
         Width           =   1815
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
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5160
      Width           =   2655
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
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4320
      Width           =   2655
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00808080&
      Caption         =   "Return to Administrator Options"
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
      TabIndex        =   19
      Top             =   3480
      Width           =   2655
   End
   Begin VB.TextBox txtDefinition 
      Height          =   375
      Left            =   7920
      TabIndex        =   7
      Top             =   600
      Width           =   1815
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
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox txtStem 
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox txtGen 
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txtNom 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   1935
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
      Left            =   3960
      TabIndex        =   9
      Top             =   2040
      Width           =   2655
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
         TabIndex        =   15
         Top             =   2640
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
         TabIndex        =   14
         Top             =   2160
         Width           =   2295
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
         TabIndex        =   13
         Top             =   1560
         Width           =   2295
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
         TabIndex        =   12
         Top             =   960
         Width           =   2175
      End
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
         TabIndex        =   11
         Top             =   480
         Width           =   2175
      End
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
      Left            =   6960
      TabIndex        =   10
      Top             =   2040
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
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   1695
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
         TabIndex        =   17
         Top             =   1200
         Width           =   1695
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
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Add Nouns"
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
      Left            =   10080
      TabIndex        =   27
      Top             =   480
      Width           =   2895
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
      Left            =   7920
      TabIndex        =   8
      Top             =   1080
      Width           =   1815
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
      Left            =   5520
      TabIndex        =   5
      Top             =   1080
      Width           =   2055
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
      Left            =   3240
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
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
      Left            =   960
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   6135
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   12975
   End
End
Attribute VB_Name = "frmAddNouns"
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
            verify = InputBox("You are about to enter " & UCase(nomS) & ", " & UCase(genS) & " a " & DeclensionMarker & " declension " & UCase(genderName) & " noun with a stem of " & UCase(stem) & "-,  a definition of " & UCase(Definition) & " and a difficulty level of: " & UCase(DifficultyName) & ". If this is what you wish type 'yes' below, if not type 'no'.")
            If LCase(verify) = "yes" Then 'checks to see if user thinks the information is correct
                Open App.Path & "\data\Nouns.txt" For Append As #1
                    Write #1, nomS, genS, stem, Declension, gender, Definition, Difficulty
                Close #1
                MsgBox UCase(nomS) & ", " & UCase(genS) & " has been added to the noun list." 'user feedback for completion
            End If
    End If
    'Clears text boxes for future use
    txtNom.Text = ""
    txtGen.Text = ""
    txtStem.Text = ""
    txtDefinition.Text = ""
    
End Sub

Private Sub cmdLogOut_Click()
    'hides current form and shows login form
    frmAddNouns.Hide
    'Public Subroutine which properly resets values and shows login page (see mdlPublicSubs)
    Call LogOut
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdReturn_Click()
    frmAddNouns.Hide
    frmAdmin.Show
End Sub
