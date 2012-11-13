VERSION 5.00
Begin VB.Form frmregister 
   Caption         =   "Age Verification"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   Picture         =   "frmregister.frx":0000
   ScaleHeight     =   7905
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit from Program"
      Height          =   1215
      Left            =   240
      TabIndex        =   6
      Top             =   5160
      Width           =   2415
   End
   Begin VB.CommandButton cmdsubmit 
      Caption         =   "Submit"
      Height          =   1215
      Left            =   7200
      TabIndex        =   5
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox txtage 
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox txtname 
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Please Enter Your Birth Year (EX- 1976 or 1954)"
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Please Enter Your Name"
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   480
      Picture         =   "frmregister.frx":6756
      Top             =   7080
      Width           =   6675
   End
   Begin VB.Label lblage 
      BackColor       =   &H000000FF&
      Caption         =   "You MUST be atleast 21 to view information about BeerBall. Please enter your name and age below;"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   360
      Width           =   4935
   End
End
Attribute VB_Name = "frmregister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdexit_Click()
End 'allows user to quit program

End Sub

Private Sub cmdsubmit_Click()
' this subroutine asks users to enter a name and age. The name is store globally and the age is verified to make sure the user is not under 21 years of age.
' being that the legal drinking age is 21 noone under the age of 21 will be allowed to make other selections


Dim age As Integer
Dim correct As Boolean

username = txtname 'saves the users input names as a registered name

age = txtage
correct = False 'sets variable to move on after age verifiaction

If Len(age) <> 4 Then 'makes length of input age of only 4 numbers acceptable
    Select Case age
        Case 1984 To 1986 'one case of an acceptable age
            MsgBox username & " congratualtions you barely made it under the wire. With your age this game will be perfect for you." 'displays a message box
            correct = True
            verified = True
        
        Case 1975 To 1984 'one case of an acceptable age
            MsgBox username & " your not to old to enjoy this game! Please Have Fun!"
            correct = True
            verified = True
        
        Case 1960 To 1975 'one case of an acceptable age
            MsgBox username & " you still should enjoy this game but your kids in college might enjoy this one more!"
            correct = True
            verified = True
        
        Case Is >= 1987 'one case of an unacceptable age
            MsgBox username & " sorry you are not old enough to play this game. Please exit the game and go to Confession."
            verified = False
        
        Case Else 'one case of an acceptable age
            MsgBox username & " not Sure why someone your age wants to play this game but be my guest and learn a little more."
            correct = True
            verified = True
        End Select
Else
    MsgBox " Sorry the year you entered is invalid please try again." 'asks user to go back and input a correct year
End If
    
If correct = True Then ' displays the main menu again if age has been verified
    frmregister.Hide
    frmmain.Show
End If

End Sub



