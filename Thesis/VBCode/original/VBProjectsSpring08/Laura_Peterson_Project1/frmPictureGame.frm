VERSION 5.00
Begin VB.Form frmPictureGame 
   Caption         =   "Picture Game"
   ClientHeight    =   8985
   ClientLeft      =   2820
   ClientTop       =   1425
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   Picture         =   "frmPictureGame.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   8835
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7560
      Width           =   1695
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H80000009&
      Height          =   1815
      Left            =   2160
      ScaleHeight     =   1755
      ScaleWidth      =   4755
      TabIndex        =   3
      Top             =   240
      Width           =   4815
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Start Game"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdRules 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Learn How to Play"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000009&
      Height          =   3735
      Left            =   2040
      ScaleHeight     =   3675
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   5160
      Width           =   5415
   End
End
Attribute VB_Name = "frmPictureGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'option explicit
'Laura's Movie Gallery
'frmPictureGame
'Laura Peterson
'3/24/2008
'Objective: The purpose of this form is to display movie images and allow
'the user to guess which film the image belongs to.

Dim Answer As Integer
Dim CTR As Integer
Dim Percent As Single

Private Sub cmdReturn_Click()
frmGenres.Show
frmPictureGame.Hide
End Sub

Private Sub cmdRules_Click()
'this will print the rules of the game into the picture box
    picResults.Cls
    picResults.Print "A picture will display in picture box. "
    picResults.Print "Select which film the picture is from based on the choices provided."
    picResults.Print "NOTE* You must finish the game before moving on"
End Sub


Private Sub cmdStart_Click()
CTR = 0
    'this will clear the picture box and load the picture into the picture box
    picResults.Cls
    picResults.Picture = LoadPicture(App.Path & "\thegraduate.jpg")
     picResults2.Cls
    'this will print the question and possible answers into the second picture box.
        picResults2.Print "To which film does this picture belong?"
        picResults2.Print "1)All About Eve "
        picResults2.Print "2)The Graduate "
        picResults2.Print "3)On the Waterfront"
        Answer = InputBox("Enter the number of the correct answer")
            If Answer = 2 Then
                    MsgBox "Correct!"
                    CTR = CTR + 1
                ElseIf Answer = 1 Then
                    MsgBox "That is incorrect, please try again."
                ElseIf Answer = 3 Then
                    MsgBox "That is incorrect, please try again."
                Else
                    MsgBox "Error. Invalid Number.", , "Error"
            End If
        
        picResults.Picture = LoadPicture(App.Path & "\grapesofwrath.jpg")
        picResults2.Cls
        picResults2.Print "To which film does this picture belong?"
        picResults2.Print "1) Casablanca"
        picResults2.Print "2) Raging Bull"
        picResults2.Print "3) The Grapes of Wrath"
        Answer = InputBox("Enter the number of the correct answer")
            If Answer = 3 Then
                    MsgBox "Correct!"
                    CTR = CTR + 1
                ElseIf Answer = 2 Then
                    MsgBox "That is incorrect, please try again."
                ElseIf Answer = 1 Then
                    MsgBox "That is incorrect, please try again."
                Else
                    MsgBox "Error. Invalid Number.", , "Error"
            End If
        picResults.Picture = LoadPicture(App.Path & "\citizenkane.jpg")
        picResults2.Cls
        picResults2.Print "To which film does this picture belong?"
        picResults2.Print "1) Citizen Kane"
        picResults2.Print "2) City Lights"
        picResults2.Print "3) The Maltese Falcon"
        Answer = InputBox("Enter the number of the correct answer")
            If Answer = 1 Then
                    MsgBox "Correct!"
                    CTR = CTR + 1
                ElseIf Answer = 2 Then
                    MsgBox "That is incorrect, please try again."
                ElseIf Answer = 3 Then
                    MsgBox "That is incorrect, please try again."
                Else
                    MsgBox "Error. Invalid Number.", , "Error"
            End If
        picResults.Picture = LoadPicture(App.Path & "\soundofmusic1.jpg")
        picResults2.Cls
        picResults2.Print "To which film does this picture belong?"
        picResults2.Print "1) Singin' In The Rain"
        picResults2.Print "2) The Sound of Music"
        picResults2.Print "3) Some Like it Hot"
        Answer = InputBox("Enter the number of the correct answer")
            If Answer = 1 Then
                    MsgBox "Correct!"
                    CTR = CTR + 1
                ElseIf Answer = 2 Then
                    MsgBox "That is incorrect, please try again."
                ElseIf Answer = 3 Then
                    MsgBox "That is incorrect, please try again."
                Else
                    MsgBox "Error. Invalid Number.", , "Error"
            End If
        picResults.Picture = LoadPicture(App.Path & "\rearwindow.jpg")
        picResults2.Cls
        picResults2.Print "To which film does this picture belong?"
        picResults2.Print "1) Double Indemnity"
        picResults2.Print "2) Psycho"
        picResults2.Print "3) Rear Window"
        Answer = InputBox("Enter the number of the correct answer")
            If Answer = 3 Then
                    MsgBox "Correct!"
                    CTR = CTR + 1
                ElseIf Answer = 2 Then
                    MsgBox "That is incorrect, please try again."
                ElseIf Answer = 1 Then
                    MsgBox "That is incorrect, please try again."
                Else
                    MsgBox "Error. Invalid Number.", , "Error"
            End If
        picResults.Picture = LoadPicture(App.Path & "\Schindler1.jpg")
        picResults2.Cls
        picResults2.Print "To which film does this picture belong?"
        picResults2.Print "1) The General"
        picResults2.Print "2) Schindler's List"
        picResults2.Print "3) Apocalypse Now"
        Answer = InputBox("Enter the number of the correct answer")
            If Answer = 2 Then
                    MsgBox "Correct!"
                    CTR = CTR + 1
                ElseIf Answer = 1 Then
                    MsgBox "That is incorrect, please try again."
                ElseIf Answer = 3 Then
                    MsgBox "That is incorrect, please try again."
                Else
                    MsgBox "Error. Invalid Number.", , "Error"
            End If
        picResults.Picture = LoadPicture(App.Path & "\streetcar.jpg")
        picResults2.Cls
        picResults2.Print "To which film does this picture belong?"
        picResults2.Print "1) A Streetcar Named Desire"
        picResults2.Print "2) Casablanca"
        picResults2.Print "3) The Maltese Falcon"
        Answer = InputBox("Enter the number of the correct answer")
            If Answer = 1 Then
                    MsgBox "Correct!"
                    CTR = CTR + 1
                ElseIf Answer = 2 Then
                    MsgBox "That is incorrect, please try again."
                ElseIf Answer = 3 Then
                    MsgBox "That is incorrect, please try again."
                Else
                    MsgBox "Error. Invalid Number.", , "Error"
            End If
        picResults.Picture = LoadPicture(App.Path & "\psycho1.jpg")
        picResults2.Cls
        picResults2.Print "To which film does this picture belong?"
        picResults2.Print "1) Bonnie and Clyde"
        picResults2.Print "2) It Happened One Night"
        picResults2.Print "3) Psycho"
        Answer = InputBox("Enter the number of the correct answer")
            If Answer = 3 Then
                    MsgBox "Correct!"
                    CTR = CTR + 1
                ElseIf Answer = 2 Then
                    MsgBox "That is incorrect, please try again."
                ElseIf Answer = 1 Then
                    MsgBox "That is incorrect, please try again."
                Else
                    MsgBox "Error. Invalid Number.", , "Error"
            End If
        picResults.Picture = LoadPicture(App.Path & "\LOTR.jpg")
        picResults2.Cls
        picResults2.Print "To which film does this picture belong?"
        picResults2.Print "1) Star Wars"
        picResults2.Print "2) The Lord of the Rings"
        picResults2.Print "3) Vertigo"
        Answer = InputBox("Enter the number of the correct answer")
            If Answer = 2 Then
                    MsgBox "Correct!"
                    CTR = CTR + 1
                ElseIf Answer = 1 Then
                    MsgBox "That is incorrect, please try again."
                ElseIf Answer = 3 Then
                    MsgBox "That is incorrect, please try again."
                Else
                    MsgBox "Error. Invalid Number.", , "Error"
            End If
        picResults.Picture = LoadPicture(App.Path & "\Casablanca.jpg")
        picResults2.Cls
        picResults2.Print "To which film does this picture belong?"
        picResults2.Print "1) All About Eve"
        picResults2.Print "2) Casablanca"
        picResults2.Print "3) It Happened One Night"
        Answer = InputBox("Enter the number of the correct answer")
            If Answer = 2 Then
                    MsgBox "Correct!"
                    CTR = CTR + 1
                ElseIf Answer = 1 Then
                    MsgBox "That is incorrect, please try again."
                ElseIf Answer = 3 Then
                    MsgBox "That is incorrect, please try again."
                Else
                    MsgBox "Error. Invalid Number.", , "Error"
            End If
'this will print the user's results and give them their percentage correct.
        Percent = CTR / 9
        picResults2.Cls
        picResults2.Print "Congratulations! You got"; CTR; "out of 9 correct! or"; Int(Percent * 100); "%"
End Sub
