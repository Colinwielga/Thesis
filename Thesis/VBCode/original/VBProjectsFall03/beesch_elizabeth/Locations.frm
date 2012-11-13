VERSION 5.00
Begin VB.Form Locations 
   BackColor       =   &H00C000C0&
   Caption         =   "Number of Locations for Design"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   11385
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton opt4 
      BackColor       =   &H00FF80FF&
      Caption         =   "4"
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   2160
      Width           =   495
   End
   Begin VB.OptionButton opt3 
      BackColor       =   &H00FF80FF&
      Caption         =   "3"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   2160
      Width           =   495
   End
   Begin VB.OptionButton opt2 
      BackColor       =   &H00FF80FF&
      Caption         =   "2"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   2160
      Width           =   495
   End
   Begin VB.OptionButton opt1 
      BackColor       =   &H00FF80FF&
      Caption         =   "1"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton next 
      Caption         =   "NEXT"
      Height          =   615
      Left            =   4440
      TabIndex        =   1
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label enter 
      Alignment       =   2  'Center
      BackColor       =   &H00C000C0&
      Caption         =   "Click the Number of Locations for Design:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1575
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Locations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Screen Printing(Main1.vpb)
'Form Name : locations(locations.frm)
'Author: Elizabeth Beesch
'Date Written: October 27, 2003
'Purpose of Form:  Have the user choose the number of locations for printing
Option Explicit 'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.

Private Sub next_Click() ' when the form loads the following will happen
    finalform.Show ' shows the next form in the sequence
    Locations.Hide ' hides current form
End Sub 'ends the commands of the button

Private Sub opt1_Click() ' when the form loads the following will happen
    D = 1 'changes the value of D
End Sub ''ends the commands of the button

Private Sub opt2_Click() ' when the form loads the following will happen
    D = 2 'changes the value of D
End Sub 'ends the commands of the button

Private Sub opt3_Click() ' when the form loads the following will happen
    D = 3 'changes the value of D
End Sub 'ends the commands of the button

Private Sub opt4_Click() ' when the form loads the following will happen
    D = 4 'changes the value of D
End Sub 'ends the commands of the button
