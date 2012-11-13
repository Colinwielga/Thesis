VERSION 5.00
Begin VB.Form frmShot 
   BackColor       =   &H0057C0E8&
   Caption         =   "Shoot On The Wild Goalies"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShot 
      BackColor       =   &H00000080&
      Caption         =   "SHOOOOOOT!"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4080
      Width           =   2775
   End
   Begin VB.TextBox txtNumber 
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8880
      TabIndex        =   8
      Top             =   2640
      Width           =   2775
   End
   Begin VB.PictureBox picLogo 
      Height          =   2295
      Left            =   120
      Picture         =   "frmShot.frx":0000
      ScaleHeight     =   2235
      ScaleWidth      =   3075
      TabIndex        =   7
      Top             =   6000
      Width           =   3135
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00004000&
      Height          =   4695
      Left            =   3840
      ScaleHeight     =   4635
      ScaleWidth      =   4755
      TabIndex        =   6
      Top             =   840
      Width           =   4815
   End
   Begin VB.CommandButton cmdChoose 
      BackColor       =   &H00000080&
      Caption         =   "Click Here to Choose A Goalie"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00000080&
      Caption         =   "Go Back To The Excel Energy Center"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6960
      Width           =   3135
   End
   Begin VB.Label lblResponse 
      BackColor       =   &H0057C0E8&
      Caption         =   "Coach's response:"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   3480
      TabIndex        =   12
      Top             =   5880
      Width           =   3975
   End
   Begin VB.Label lblCoach 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1455
      Left            =   3360
      TabIndex        =   11
      Top             =   6360
      Width           =   5055
   End
   Begin VB.Label lblInfoShot 
      BackColor       =   &H0057C0E8&
      Caption         =   "Enter The Number Of The Place Where You Want To Shoot (1 to 5) And Click The Shoot Button:"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1575
      Left            =   8880
      TabIndex        =   9
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label lblChoose 
      BackColor       =   &H0057C0E8&
      Caption         =   "Choose The Goalie You Want To Shoot On:"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label lblHarding 
      BackColor       =   &H0057C0E8&
      Caption         =   "37. Josh Harding"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label lblBackstrom 
      BackColor       =   &H0057C0E8&
      Caption         =   "32. Nicklas Backstrom"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label lblShot 
      BackColor       =   &H0057C0E8&
      Caption         =   "This is your chance to shoot on one of the Minnesota Wild's goalies!"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   615
      Left            =   1440
      TabIndex        =   1
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "frmShot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Minnesota Wild Visual Basic project
'Shot Form
'Authors: Adam Andersson
'22 Feb 2010
'The purpose of this form is to have the user interact with
'the program by shooting on either of the 2 Minnesota Wild
'goaltenders by entering information into input boxes.


Dim pictureNumber As Integer


Private Sub cmdChoose_Click()
'this subroutine loads a picture into a picture box
'the list of picture names has already been loaded into the array
'The filenames were put into the array using a code Module that was
'executed when the program first started running.
'a comment in the label box will appear when the user tries to shoot on the goalie

Dim Found As Boolean

Found = False

'use inputbox to make the user choose the goalie
pictureNumber = InputBox("Enter: 1, if you want to shot at Backstrom or enter: 2, if you want to shot at Harding:")

Do While Not Found
    If pictureNumber = 1 Or pictureNumber = 2 Then
        Found = True
    Else
        pictureNumber = InputBox("Enter 1 if you want to shot at Backstrom or 2 if you want to shot at Harding:")
    End If
Loop
'use the number to choose the desired filename from the array of names

picResults.Picture = LoadPicture(App.Path & "\images\" & Pictures(pictureNumber)) 'load the picture into the picturebox

lblCoach.Caption = "" 'clear the label


End Sub

Private Sub cmdShot_Click()

If pictureNumber = 1 Then 'if the user chose backstrom this different pictures will be load
    If txtNumber.Text = 1 Then
        picResults.Picture = LoadPicture(App.Path & "\images2\" & Saves(1)) 'load picture 1
        lblCoach.Caption = "Backstrom has a sick blocker! That was a terrible spot to shoot at!" 'load comment in the label box
    End If
    If txtNumber.Text = 2 Then
        picResults.Picture = LoadPicture(App.Path & "\images2\" & Saves(2)) 'load picture 2
        lblCoach.Caption = "Backstrom's glove is even better than his blocker, figure it out!" 'load comment in the label box
    End If
    If txtNumber.Text = 3 Then
        picResults.Picture = LoadPicture(App.Path & "\images2\" & Saves(3)) 'load picture 3
        lblCoach.Caption = "Hmm... Come on, you're not even trying. My tip: Shoot Better!" 'load comment in the label box
    End If
    If txtNumber.Text = 4 Then
        picResults.Picture = LoadPicture(App.Path & "\images2\" & Saves(4)) 'load picture 4
        lblCoach.Caption = "Good save by Backstrom, not much you can do there." 'load comment in the label box
    End If
    If txtNumber.Text = 5 Then
        picResults.Picture = LoadPicture(App.Path & "\images2\" & Saves(5)) 'load picture 5
        lblCoach.Caption = "GOAL!!! Hahaha, look how he is crawling on the ice. He never knows what to do when someone shoots at the five hole! Good Job!" 'load comment in the label box
    End If
End If
If pictureNumber = 2 Then 'if the user chooses Harding, a different picture will be loaded
    If txtNumber.Text = 1 Then
        picResults.Picture = LoadPicture(App.Path & "\images2\" & Saves(6)) 'load picture 6
        lblCoach.Caption = "Look up before you shoot, Harding has his glove on his right hand!.. And he loves his glove..." 'load comment in the label box
    End If
    If txtNumber.Text = 2 Then
        picResults.Picture = LoadPicture(App.Path & "\images2\" & Saves(7)) 'load picture 7
        lblCoach.Caption = "Josh's blocker is so fast, he even was able to switch gear before the shot reached the net. Try again bud!" 'load comment in the label box
    End If
    If txtNumber.Text = 3 Then
        picResults.Picture = LoadPicture(App.Path & "\images2\" & Saves(8)) 'load picture 8
        lblCoach.Caption = "GOAL!!! Nice dangle before the shot too! Josh thought that you were shooting at his awesome blocker, so he tried to switch his gear before making the save. Look how dumb he looks now!" 'load comment in the label box
    End If
    If txtNumber.Text = 4 Then
        picResults.Picture = LoadPicture(App.Path & "\images2\" & Saves(9)) 'load picture 9
        lblCoach.Caption = "Ehhh.. If you shoot outside the net you will obviously not score..." 'load comment in the label box
    End If
    If txtNumber.Text = 5 Then
        picResults.Picture = LoadPicture(App.Path & "\images2\" & Saves(10)) 'load picture 10
        lblCoach.Caption = "No... Try Again!" 'load comment in the label box
    End If
End If
    
    
End Sub

Private Sub cmdBack_Click()
    lblCoach.Caption = "" 'clear the label
    picResults.Picture = LoadPicture(App.Path & "\images2\" & Saves(11)) 'load picture 11 (blank page)
    txtNumber.Text = "" 'clear txt box

'show the main form and hide the other forms
    frmWelcome.Hide
    frmMain.Show
    frmRoster.Hide
    frmShot.Hide
    frmShop.Hide
    frmLeague.Hide

End Sub





