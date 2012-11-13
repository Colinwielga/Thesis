VERSION 5.00
Begin VB.Form Instructions 
   BackColor       =   &H00000000&
   Caption         =   "Instructions"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSwitch 
      BackColor       =   &H000000FF&
      Caption         =   "Start Program"
      BeginProperty Font 
         Name            =   "Tekton Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7800
      Width           =   3135
   End
   Begin VB.CommandButton cmdInstructions 
      BackColor       =   &H000000FF&
      Caption         =   "View Instructions"
      BeginProperty Font 
         Name            =   "Tekton Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   2655
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00000000&
      FillColor       =   &H000000FF&
      ForeColor       =   &H000000FF&
      Height          =   5925
      Left            =   480
      ScaleHeight     =   5865
      ScaleWidth      =   10155
      TabIndex        =   1
      Top             =   1800
      Width           =   10215
      Begin VB.Image Image1 
         Height          =   2280
         Left            =   3120
         Picture         =   "Instructions.frx":0000
         Top             =   3480
         Width           =   2865
      End
   End
   Begin VB.Label lblGeneral 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Budget Information For '07-'08 School Year"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   7455
   End
End
Attribute VB_Name = "Instructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdInstructions_Click()
    'this button will print a description and some instructions for the user
    
picResults.Print "Welcome to the Budget Calculation Program"
picResults.Print "****************************************************"
picResults.Print
    
    'print description
picResults.Print "The intent of this program is to help its user determine whether the expenses they have estimated for the year "
picResults.Print "equal the budget amount they have been given."
picResults.Print
    
    'print instructions
picResults.Print "To start the program, click the Start Program button below."
picResults.Print "Next, enter your given budget amount in the text box provided."
picResults.Print "Then, click the Enter Estimated Cost per Month button, which will ask you to input total estimated expenses for each of 9 months"
picResults.Print "Then, click the Total Estimated Cost button, which will calculate and print the sum of your inputs"
picResults.Print "Finally, click the Has Budget Requirement Been Met? button, which will compare the given budget amount to the total estimated cost."
picResults.Print
picResults.Print "The program will then inform the user if their estimated costs are greater,equal to, or less than thier given budget amount."
picResults.Print
picResults.Print "Enjoy!!"

End Sub

Private Sub cmdSwitch_Click()
'this button will move from the initial form to the actual program

Instructions.Hide          'this code hides the 1st form
BudgetInfo.Show     'this code makes the actual program visible

End Sub

