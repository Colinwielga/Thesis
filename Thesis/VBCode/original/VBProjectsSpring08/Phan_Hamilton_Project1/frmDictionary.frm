VERSION 5.00
Begin VB.Form Dictionary 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H00FF0000&
      Caption         =   "Show all words in alphabetical order"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdTranslate 
      BackColor       =   &H00FF0000&
      Caption         =   "Translate!"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FF0000&
      Caption         =   "Return to Destination"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      Height          =   5895
      Left            =   4440
      ScaleHeight     =   5835
      ScaleWidth      =   4995
      TabIndex        =   1
      Top             =   240
      Width           =   5055
   End
   Begin VB.TextBox txtEnter 
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label lblIntro 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "<-- Works only once! Must go back to Destination to reset!"
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
      Left            =   2280
      TabIndex        =   6
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblEnter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please enter the English word (make sure it is spelled correctly and first letter capitalized!):"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Learning Korean 101
'Form Name: Dictionary
'Amanda Phan and Natalie Hamilton
'Date: March 10
'Objective:  The form the user is on will allow the user to do 2 things.  The first thing
    'the user can do is type in an English word.  Then when the user clicks on the button
    'labeled translate, the Korean  phonetic translation will appear.  The second thing
    'is that the user can ask the program to show all the words in the dictionary in the
    'picture box in alphabetical order.
'Comments:  The form allows the user to type in a word in the textbox, and click on the
    'button labeled translate.  Here, the word is then in a match and stop search within
    'the array. Once the word is found, the Korean translation will appear in a messagebox.
    'If the word is not found, a messagebox will appear saying that the word was not found.
    'If the user wants all the words to be listed, they would clicked on the button labeled,
    'show all. Once the button is clicked, a bubble sort is conducted that will sort the
    'array of words in alphabetical order.  Then, an exhaustive search is created to show
    'all the words.  Here then, all the words will be listed within the picturebox. If the user would like
    'to return to the previous form, they will have to click on the button labeled return. When the
    'the button is clicked on, the present form will hide and the desired previous form
    'will appear.



Option Explicit
Dim EnglishWord(1 To 100) As String, KoreanWord(1 To 100) As String
Dim CTR As Integer

Private Sub cmdAll_Click()
Dim Pos As Integer, Pass As Integer, TempName As String, J As Integer
Dim Found As Boolean

Found = False
'This action will create the labels for the dictionary being listed.
picResults.Print "This file is now put into alphbetical order:"
picResults.Print "English Word", Tab(20), "Korean Word"
picResults.Print "**********************************************************"
'This action will open the Dictionary Array
CTR = 0
    Open App.Path & "\DictionaryList2.txt" For Input As #1
        Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, EnglishWord(CTR), KoreanWord(CTR)
        Loop

'This action will sort the words in the dictionary in alphabetical order
Pos = 0
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If EnglishWord(Pos) > EnglishWord(Pos + 1) Then
                Found = True
                TempName = EnglishWord(Pos)
                EnglishWord(Pos) = EnglishWord(Pos + 1)
                EnglishWord(Pos + 1) = TempName
                TempName = KoreanWord(Pos)
                KoreanWord(Pos) = KoreanWord(Pos + 1)
                KoreanWord(Pos + 1) = TempName
            End If
        Next Pos
    Next Pass
    
    

'This action will display all the words in the dictionary in a table in the picture box
For J = 1 To CTR
    picResults.Print EnglishWord(J), Tab(20), KoreanWord(J)
   
Next J



End Sub

Private Sub cmdReturn_Click()
'This action will cause the present form to hide and the previous form to appear.
Destination.Show
Dictionary.Hide

End Sub

Private Sub cmdTranslate_Click()
Dim Pos As Integer, Found As Boolean, Word As String
'This allows the program to read what the user typed in.
Word = txtEnter.Text
'This opens the Dictionary array
    Found = False
    CTR = 0
    Open App.Path & "\DictionaryList2.txt" For Input As #1
        Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, EnglishWord(CTR), KoreanWord(CTR)
        Loop
'This creates a match and stop search to find the word in the dictionary.
Do While (Not Found And Pos < CTR)
        Pos = Pos + 1
            If EnglishWord(Pos) = Word Then
            Found = True
                If Found Then
                    'This action will have the translation pop up in a messagebox
                    MsgBox ("The Korean translation of the word " & Word & " is " & KoreanWord(Pos) & ".")
                End If
            End If
  
Loop
'This is if the word is not found in the dictionary, a messagebox will pop up telling the user that
If Not Found Then
    MsgBox ("Sorry, the word " & Word & " is not in the dictionary.")
End If
'This closes the list
Close #1

            






End Sub

