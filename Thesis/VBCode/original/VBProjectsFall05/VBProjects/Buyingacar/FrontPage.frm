VERSION 5.00
Begin VB.Form MainForm1 
   BackColor       =   &H80000012&
   Caption         =   "MainForm1"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   ScaleHeight     =   9045
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H008080FF&
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   2
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton cmdMake 
      BackColor       =   &H008080FF&
      Caption         =   "Choose car by Make/Brand"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton cmdStyle 
      BackColor       =   &H008080FF&
      Caption         =   "Choose car by style"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H008080FF&
      Caption         =   "Searching for a new car?"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2040
      TabIndex        =   4
      Top             =   1080
      Width           =   7455
   End
   Begin VB.Image Image1 
      Height          =   2310
      Left            =   3720
      Picture         =   "FrontPage.frx":0000
      Top             =   2520
      Width           =   4050
   End
End
Attribute VB_Name = "MainForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Buying A Car(VB-project.vbp)
'Form Name: MainForm1 (Vbpofect - frm.frm)
'Author: Katie Lee
'Date : Monday October 31, 2005
'Purpose of the Project: To have the user interact with the program
                    'to decide how much they want to spend on a car
                    ' and which style and brand of car they prefer
                    'so the program can educate them on the car
                    ' that is right for them
'Purpose of the form:  It is the starting blocks of the project from, the
                    ' MainForm1 the user can navigate to different forms to
                    'search different styles and brand of cars he/she might purchase.
                    'The user must input information using InputBoxes.
                    'It also informs the user through message box's that pop up.
Option Explicit

Private Sub cmdLoad_Click()
Dim I As Integer

Open App.Path & "\CarData.txt" For Input As #1 'opens file so arrays can be read in 'Path = "M:\CS130\Buying_A_Car\"
I = 1
For I = 1 To 39
    Input #1, Model(I), Style(I), Company(I), Price(I)
Next I
Close #1 'closes file when done reading in array

End Sub

Private Sub cmdMake_Click()
MainForm1.Hide
MakeForm.Show 'goes to the MakeForm to decide which style of that Brand he/she prefers

End Sub

Private Sub cmdQuit_Click()
End 'allows users to exit the program
End Sub

Private Sub cmdStyle_Click()
MainForm1.Hide
StyleForm.Show 'goes to the  StyleForm to determine which brand is cheapest for the style the user input

End Sub
