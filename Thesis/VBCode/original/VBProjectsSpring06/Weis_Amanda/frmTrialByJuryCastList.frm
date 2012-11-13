VERSION 5.00
Begin VB.Form frmTrialByJuryCastList 
   BackColor       =   &H00000000&
   Caption         =   "Trial By Jury Cast List"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   Picture         =   "frmTrialByJuryCastList.frx":0000
   ScaleHeight     =   7575
   ScaleWidth      =   14985
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCastList 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   6735
      Left            =   8520
      ScaleHeight     =   6675
      ScaleWidth      =   6195
      TabIndex        =   5
      Top             =   120
      Width           =   6255
   End
   Begin VB.CommandButton cmdCastByYear 
      Caption         =   "Cast By Year"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   4
      ToolTipText     =   "Click to display the cast list with year in front."
      Top             =   5880
      Width           =   2415
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort Cast By Year"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Click to sort the cast list by year."
      Top             =   6720
      Width           =   2655
   End
   Begin VB.CommandButton cmdInputYear 
      Caption         =   "Learn About A Certain Character"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Click to type in the name of the character you would like to learn more about."
      Top             =   5040
      Width           =   2655
   End
   Begin VB.CommandButton cmdCastList 
      Caption         =   "Cast List"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Click to display the cast list."
      Top             =   5880
      Width           =   2655
   End
   Begin VB.CommandButton cmdGoBack 
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   0
      ToolTipText     =   "Click to go back to previous form."
      Top             =   6720
      Width           =   2415
   End
   Begin VB.Label lblDesigned 
      BackColor       =   &H00000000&
      Caption         =   "Designed by Amanda Weis"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmTrialByJuryCastList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'declare all arrays and variables
    Dim Character(1 To 100) As String
    Dim Year(1 To 100) As String
    Dim Names(1 To 100) As String
    Dim Arraysize As Integer
    'open and read and print file displaying cast list
Private Sub cmdCastByYear_Click()
    Dim Pos As Integer
    Pos = 0
    picCastList.Cls
    Open App.Path & "\TrialByJuryYear.txt" For Input As #1
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, Year(Pos), Character(Pos), Names(Pos)
        picCastList.Print Year(Pos); Tab(25); Character(Pos), Names(Pos)
    Loop
    Close #1
End Sub
    'open and read and display another file organized first by year
Private Sub cmdCastList_Click()
    Dim Pos As Integer
    Pos = 0
    picCastList.Cls
    Open App.Path & "\TrialByJury.txt" For Input As #1
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, Character(Pos), Names(Pos), Year(Pos)
        picCastList.Print Character(Pos); Tab(25); Names(Pos), Year(Pos)
    Loop
    Close #1
    Arraysize = Pos
End Sub
    'create button to go back to previous slide
Private Sub cmdGoBack_Click()
    frmTrialByJuryCastList.Hide
    frmTrialByJury.Show
End Sub

    'create inputbox and use if statements for the user to input a character name and a message box will appear showing them information
Private Sub cmdInputYear_Click()
    Dim InputYear As String
    InputYear = InputBox("Which character would you like to learn more about?", "Input")
    If InputYear = "Judge" Then
        MsgBox "The Judge is not the most fair judge around.  He had married an 'elderly, ugly daughter' of a rich old attorney in order to become a successful judge.  Once becoming a judge, he got rid of the ugly daughter.  By the end of Trial By Jury, the judge solves everyone's problem by marrying the beautiful Angelina."
    ElseIf InputYear = "Angelina" Then
        MsgBox "Angelina feels betrayed by Edwin, who left her at the altar for another woman.  Instead of being the hurt and crying little girl that some people might expect, Angelina takes matters into her own hands, and hires a determined attorney to squeeze all the money out of Edwin."
    ElseIf InputYear = "Edwin" Then
        MsgBox "Edwin is the kind of guy that just won't stay put with one woman.  He promises to marry Angelina before he decides to move on to another woman.  Unexpectedly, Angelina takes him to court, where he pleads his case of innocence.  Though throughout the opera everyone is pitted against him, though he wins in the end just as everyone else does and walks out a 'free' man."
   ElseIf InputYear = "Bridesmaid" Then
        MsgBox "The Bridesmaids are Angelina's best friends and wedding party.  They are their to work their charm on the judge and jury and support their friend Angelina."
    ElseIf InputYear = "Usher" Then
        MsgBox "The Usher provides comic relief and is the judge's righthand man."
    ElseIf InputYear = "Court Reporter" Then
       MsgBox "The Court Reporters are there to take note of what is going on and provide further comic relief for the audience."
    Else: InputYear = "Counsel"
        MsgBox "The Counsel woman is hired to defend Angelina's case.  She is a determined feminist who encourages Angelina to play the poor and beautiful victim."
    End If
End Sub
    'declare variables
    'use Next to arrange the cast by year
Private Sub cmdSort_Click()
    Dim Pass As Integer
    Dim I As Integer
    Dim Temp As String
    Dim N As Integer
    Dim Pos As Integer
    Dim TempCharacter As String
    Dim TempNames As String
    N = Arraysize
    Arraysize = Pos
    For Pass = 1 To N - 1
        For I = 1 To N - Pass
            If Year(I) > Year(I + 1) Then
                Temp = Year(I)
                Year(I) = Year(I + 1)
                Year(I + 1) = Temp
            End If
            TempCharacter = Character(I)
            Character(I) = Character(I + 1)
            Character(I + 1) = TempCharacter
            TempNames = Names(I)
            Names(I) = Names(I + 1)
            Names(I + 1) = TempNames
        Next I
    Next Pass
    Pos = 0
    picCastList.Cls
    For Pos = 1 To N
        picCastList.Print Year(Pos); Tab(25); Character(Pos), Names(Pos)
    Next Pos
End Sub



