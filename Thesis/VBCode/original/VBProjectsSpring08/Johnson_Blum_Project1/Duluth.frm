VERSION 5.00
Begin VB.Form Duluth 
   BackColor       =   &H00800000&
   Caption         =   "Form3"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10215
   LinkTopic       =   "Form3"
   ScaleHeight     =   8880
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Height          =   1575
      Left            =   6120
      Picture         =   "Duluth.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   2235
      TabIndex        =   13
      Top             =   2520
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      Height          =   615
      Left            =   3360
      ScaleHeight     =   555
      ScaleWidth      =   6795
      TabIndex        =   12
      Top             =   1440
      Width           =   6855
   End
   Begin VB.CommandButton cmdActivity 
      Caption         =   "Type in the Text box the $ you would like to spend for an activity==> then press here!"
      Height          =   855
      Left            =   240
      TabIndex        =   11
      Top             =   7920
      Width           =   2055
   End
   Begin VB.TextBox txtNumber 
      Height          =   495
      Left            =   2640
      TabIndex        =   10
      Top             =   8040
      Width           =   1815
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   2520
      Picture         =   "Duluth.frx":C7C6
      ScaleHeight     =   1515
      ScaleWidth      =   1755
      TabIndex        =   8
      Top             =   6000
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   360
      Picture         =   "Duluth.frx":168F8
      ScaleHeight     =   1515
      ScaleWidth      =   1875
      TabIndex        =   7
      Top             =   6000
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   3840
      Picture         =   "Duluth.frx":1FC1A
      ScaleHeight     =   1515
      ScaleWidth      =   1995
      TabIndex        =   6
      Top             =   2520
      Width           =   2055
   End
   Begin VB.PictureBox picBridge 
      Height          =   2535
      Left            =   4560
      ScaleHeight     =   2475
      ScaleWidth      =   4515
      TabIndex        =   5
      Top             =   4200
      Width           =   4575
   End
   Begin VB.CommandButton cmdPicture 
      BackColor       =   &H00808080&
      Caption         =   "Take a glimpse of the Aerial Bridge ==>"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   3375
   End
   Begin VB.CommandButton cmdHomepage 
      BackColor       =   &H8000000D&
      Caption         =   "Back to Homepage"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdTrivia 
      Caption         =   "Answer a Trivia Question"
      BeginProperty Font 
         Name            =   "Berlin Sans FB Demi"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   480
      TabIndex        =   2
      Top             =   2640
      Width           =   2895
   End
   Begin VB.CommandButton cmdFacts 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Press Here For a Fun Fact! "
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   15.75
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      MaskColor       =   &H00404040&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label lbl 
      Caption         =   "Known for Scenic wonders and                  Natural Beauty"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   4800
      TabIndex        =   9
      Top             =   6840
      Width           =   3975
   End
   Begin VB.Label lblDuluth 
      Caption         =   "     \\ Duluth\\"
      BeginProperty Font 
         Name            =   "Kozuka Gothic Pro B"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "Duluth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                'Project Name: Minnesooota
                'Form Name: Duluth
                'Author: Danielle Johnson and Tony Blum
                'Date Written: March 25, 2008
                'in this form, the user can press command buttons to find out facts, look at a picture of a landmark of Duluth, and answer a trivia question
Option Explicit 'declares all variables on this form

Private Sub cmdActivity_Click()
Dim Number As Single

Number = txtNumber.Text

Select Case Number
    Case Is < 5 'if the user types in a number less than 5, this message will pop up
        MsgBox "You can get an ice cream treat at Lake Street.", , "Activity"
    Case Is < 15 'if the user types in a number less then 15 more than 5, this message will pop up
        MsgBox "Go to the I-Max theater or the zoo.", , "Activity"
    Case Is < 20 'if the user types in a number less than 20 more than 15, this message will pop up
        MsgBox "You can go on a Vista Queen/Vista King Ride.", , "Activity"
    Case Is < 40 'if the user types in a number less than 40 more than 20, this message will pop up
        MsgBox "You can go to the Great Lakes Aquarium.", , "Activity"
    Case Is < 50 'if the user types in a number less than 50 more than 40,this message will pop up
        MsgBox " You can go on a train ride from the depot.", , "Activity"
    Case Is < 100 'if the user types in a number less than 100 more than 50, this message will pop up
        MsgBox " Go to Lester Park Golf Course!", , "Activity"
    Case Is > 100 'if the user tpes in number less than 100, this message will pop up
        MsgBox " Rent a Limo and hit D-TOWN!!", , "Activity"
        
End Select
    

End Sub

Private Sub cmdFacts_Click()
Dim Facts(1 To 50) As String, CTR As Integer

Open App.Path & "\Duluthfacts.txt" For Input As #1

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Facts(CTR)
Loop
Close #1
picResults.Cls
picResults.Print Facts(CInt(Int((10 * Rnd()) + 1)))

End Sub


Private Sub cmdHomepage_Click()
Duluth.Hide 'brings user back to Minnesota homepage
Minnesota.Show

End Sub

Private Sub cmdPicture_Click()
picBridge.Picture = LoadPicture(App.Path & "\duluthbridge.jpg") 'when user presses command button, the picture of the aerial bridge will appear in picture box

End Sub

Private Sub cmdTrivia_Click()
Dim Boat As String 'declare the variable

Boat = InputBox("What is the name of the infamous ship that sank in Lake Superior in 1975? (make sure you get the spelling right;)", "Trivia Question")
'user types in answer

If Boat = "Edmund Fitzgerald" Then
    MsgBox "You are correct! " & Boat & " was the name of the ship that sank...what a tradegy.", , "Boat"
Else
    MsgBox "No, I'm sorry, the name was Edmund Fitzgerald.", , "Boat" 'both messages are variations of what to expect when the user types in their answer
    End If
End Sub


