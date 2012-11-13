VERSION 5.00
Begin VB.Form frmRateService 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   11250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRateService 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Click here to Rate the Quality of Service you recieved while on your Safari Adventure"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   10815
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Gungsuh"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Return to Safari HQ"
      BeginProperty Font 
         Name            =   "Gungsuh"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C000C0&
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1275
      ScaleWidth      =   10875
      TabIndex        =   0
      Top             =   1200
      Width           =   10935
   End
End
Attribute VB_Name = "frmRateService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The Great Safari Adventure
'frmRateService
'Kit and Liz Chambers
'Feb 24th 2010
'This form is ment to use a select case statment to find the quality of service

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRateService_Click()
    Dim Points As Single, Grade As String
    
    'clear the pictureBox
    picResults.Cls
    
    'The next line of code uses an InputBox function to get input from the user.
    'A box will pop up and prompt the user for a number.
    Points = InputBox("Enter the points to rate the service 0 being the worst and 100 being the best, (0-100) ")

    Select Case Points
        Case Is >= 90
            Grade = "A! we are so glad you enjoyed your stay"
        Case Is >= 80
            Grade = "B, next time let us know how we can put on our A game and serve you better!"
        Case 70 To 79
            Grade = "C... whoops! We're sorry, please leave us a comment in the comment box "
        Case 60, 61, 62 To 69.5
            Grade = "D, that bad? yikes! thanks for letting us know enjoy your vacation"
        Case Is >= 0
            Grade = "F. We have failed you"
        Case Else
            Grade = "Oops, invalid point total, please check your data"
    End Select
    picResults.Print "With "; Points; " points,"; Chr(10); "the grade you gave us is "; Grade; "."
End Sub


Private Sub cmdReturn_Click()
frmRateService.Hide 'hides this form
FrmWelcome.Show 'brings you to welcome form
End Sub
