VERSION 5.00
Begin VB.Form EuropeanGame 
   BackColor       =   &H80000008&
   Caption         =   "How well do you know Europe?"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   1
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.Image picImage 
      Height          =   3855
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   5175
   End
End
Attribute VB_Name = "EuropeanGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: The Globe Trotter Experience
'Form name: EuropeanGame.frm
'Author: Marta Gago & Brian Downes
'Date Written: Thursday March 27th, 2008
'Objective of form:  The objective of this form is to test the user of his knowledge

Option Explicit
Dim Photo As String
Dim Score As Integer
Dim country(1 To 100) As String
Dim pictures(1 To 100) As Integer       'Dim variables for the entire form

'Hides the EuropeanGame Form and shows the Europe Form
Private Sub cmdBack_Click()
EuropeanGame.Hide
Europe.Show
End Sub
'The game that tests the user with pictures of europe
'and the answer of what country that picture is assocciated with
Private Sub cmdLoad_Click()

ctr = 0     'Set ctr to zero

Open App.Path & "\Europe.txt" For Input As #1   'open the europe data file for country names
'data is put into an array with the Do while function
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, country(ctr)
Loop

    MsgBox ("Let's go!!!")
    

For j = 1 To ctr    'Using exhaustive search, a picture is loaded,
                    'an answer is given (unless the user types 'Finish' to quit), and is scored on his/her performance


    picImage.Picture = LoadPicture(App.Path & "\Europe\" & country(j) & ".jpg")                 'First picture is loaded with the country name
    Photo = InputBox("What Country Do You Associate This With?" & " Type 'Finish' To Quit")     'The user is prompted to answer what country the picture is associated with
   
            If Photo = country(j) Then  'if the user gets the country name right then he is given a +1 to his score
                MsgBox "You're Good"
                Score = Score + 1
            ElseIf Photo = "Finish" Then    'if the user types 'Finish' the game ends and the users score is shown
                GoTo endloop                'Brings the user outside of the loop and the game ends
            Else: MsgBox ("Incorrect.  The Country is " & country(j))
            End If
            
        
Next j
endloop:
MsgBox ("Your score is " & Score & " out of " & ctr)    'The user's score is shown

Close       'Close the array

End Sub


