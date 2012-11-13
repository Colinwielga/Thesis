VERSION 5.00
Begin VB.Form ChoiceForm 
   BackColor       =   &H00000080&
   Caption         =   "Choose Your Race"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9960
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Quitbutton 
      BackColor       =   &H0000C0C0&
      Caption         =   "Quit"
      Height          =   1335
      Left            =   8520
      MaskColor       =   &H000000FF&
      TabIndex        =   2
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command5Kbutton 
      BackColor       =   &H000000FF&
      Caption         =   "5,000 Meters"
      Height          =   855
      Left            =   4800
      TabIndex        =   1
      Top             =   5760
      Width           =   3255
   End
   Begin VB.CommandButton milebutton 
      Caption         =   "1 Mile"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   5760
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000C0&
      Caption         =   "Please Select One of the Following Races:"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   5160
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Steve Prefontaine Split Calculator and Comparison Program"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   360
      Width           =   8415
   End
   Begin VB.Image Image1 
      Height          =   3495
      Left            =   1920
      Picture         =   "RaceChoiceForm.frx":0000
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   5970
   End
End
Attribute VB_Name = "ChoiceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Project Name: TrackandFieldProgram (TrackProgram)'
'Form Name: ChoiceForm (RaceChoiceForm.frm)'
'Written By: Paul Jeske'
'Date Written: March 15th, 2004'
'Purpose of the Project: To display the splits from a given race, to give the
                         'user their fastest lap, slowest lap, average lap and total
                         'time.  Next the total time is compared to the times of
                         'running great Steve Prefontaine to show the user how
                         'far ahead or behind they are from his times
'Purpose of this form: Displays Title of program and allows user
                       'to choose one of two races(the mile or the 5k)
                       'which then advances them to the next form

'Command that forces user to declare a variable as needed'
Option Explicit

Private Sub Command5Kbutton_Click()
'Hides "ChoiceForm" and displays the new "FiveKForm"'
FiveKForm.Show
ChoiceForm.Hide
End Sub
Private Sub milebutton_Click()
'Hides "ChoiceForm" and displays the new "MileForm"'
    MileForm.Show
    ChoiceForm.Hide

End Sub
Private Sub Form_Load()
'Declares the Path for which all inputs will use in this project'
Path = "N:\CS130\handin\Jeske, Paul\"
End Sub
Private Sub Quitbutton_Click()
'Ends Program'
End
End Sub
