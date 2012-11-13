VERSION 5.00
Begin VB.Form frmCharacters 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF0000&
   Caption         =   "View the Mario Character Profiles"
   ClientHeight    =   10125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   ScaleHeight     =   10125
   ScaleWidth      =   11640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Search For a Character's Profile"
      Height          =   975
      Left            =   2040
      TabIndex        =   3
      Top             =   4440
      Width           =   2412
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "See the Characters to Choose From"
      Height          =   975
      Left            =   2040
      TabIndex        =   2
      Top             =   1800
      Width           =   2412
   End
   Begin VB.PictureBox picResults 
      Height          =   4452
      Left            =   5400
      ScaleHeight     =   4395
      ScaleWidth      =   4155
      TabIndex        =   1
      Top             =   1680
      Width           =   4212
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to the main page"
      Height          =   732
      Left            =   7200
      TabIndex        =   0
      Top             =   840
      Width           =   2412
   End
   Begin VB.PictureBox picPicture 
      Height          =   3375
      Left            =   600
      Picture         =   "frmLayouts.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   10155
      TabIndex        =   6
      Top             =   6360
      Width           =   10215
   End
   Begin VB.Label lblMyname 
      BackStyle       =   0  'Transparent
      Caption         =   "By Bill Macy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   9840
      Width           =   1935
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "By Bill Macy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   10680
      Width           =   1935
   End
   Begin VB.Label lblspelling 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You must spell and punctualize the Character as it appears in the list to get an appropriate search result"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   10.5
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1680
      TabIndex        =   5
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The Characters of the First Game!"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmCharacters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Mario Madness
'Form name: frmCharacters
'Author: Bill Macy
'Date Written: Tuesday March 14th, 2006
'Objective of form:  This form allows the user to view all the characters in the first
                    'mario game and then search the profiles of each of those characters.


Option Explicit
Dim counter As Integer      'declares my variables
Dim i As Integer
Public size As Integer
Dim Character(1 To 17) As String
Dim Information(1 To 17) As String

Private Sub cmdreturn_Click()
    frmCharacters.Hide      'hides the characters page to go back to the main page
    frmMain.Show        'allows the user to return to the main page
End Sub

Private Sub cmdsearch_Click()
    Dim search As String    'declares my variables
    Dim searchtrue As Boolean
    Dim pos As Integer
    Dim searchoutput As String
    searchtrue = False      'sets the variable so a match isnt found
    search = 0
    search = InputBox("Please enter the name of the character you are looking for.", "Character Name")      'allows the user to enter a name or part of a name to search for a profile of a character
    For pos = 1 To size     'loops through the list looking for the letters that the user typed in the input box
            searchoutput = InStr(Character(pos), search)        'looks in the strings for each letter
            If searchoutput <> 0 Then       'if something was found it loops through to display the match
            searchtrue = True       'sets it so the search was successful
            MsgBox "The character you selected is " & Character(pos) & ".  The profile: " & Information(pos), , "Profile"     'displays the information to the user from their search
        End If
    Next pos
        If searchtrue = False Then      'loops through is the search was unsuccessful
            MsgBox "The entry you made is either spelt or punctualized wrong.  Capitalization matters.  You may have also entered a Character that is not in the first game.  Please try again and refer to the list of characters.", , "Error"      'informs the user that the search was incorrect or not found
        End If
End Sub

Private Sub cmdStart_Click()
    picresults.Cls      'clear the picture box of anything that is in it
    counter = 0     'sets the counter equal to zero
    Open App.Path & "\Characters.txt" For Input As #2       'opens the file that has the character names
    Do Until EOF(2)     'loops through storing the names in an array
        counter = counter + 1   'increments by one so each name is stored in a different spot
        Input #2, Character(counter), Information(counter)  'writes the information into the array
    Loop
    Close #2        'closes the file
    size = counter      'sets the size equal to the number of names placed in the array
    picresults.Print "Character Name"       'prints the words character name
    picresults.Print "**********************"       'prints a bunch of astrics
    For i = 1 To counter        'loops through to print the characters names
        picresults.Print Character(i)       'prints the names
    Next i      'moves to the next name
End Sub
