VERSION 5.00
Begin VB.Form frmreferences 
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Myriad Pro Light"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   2
      Top             =   9240
      Width           =   4095
   End
   Begin VB.CommandButton cmdreferences 
      Caption         =   "References"
      BeginProperty Font 
         Name            =   "Myriad Pro Light"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   9240
      Width           =   4095
   End
   Begin VB.PictureBox picresults 
      AutoSize        =   -1  'True
      Height          =   8775
      Left            =   240
      ScaleHeight     =   8715
      ScaleWidth      =   8355
      TabIndex        =   0
      Top             =   240
      Width           =   8415
   End
End
Attribute VB_Name = "frmreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: A Review of Theoretical Orientations in Clinical Psychology
'Form name: frmreferences
'Author: Calvin Pipenhagen
'Date Written: March 27, 2008
'Objective: This form displays the refrences used for the project

Private Sub cmdreferences_Click() 'loads references
Dim references(1 To 100) As String
Dim ctr As Integer
Dim n As Integer
picresults.Cls
Open App.Path & "\references.txt" For Input As #1
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, references(ctr)
    Loop
    Close #1
For n = 1 To ctr
    picresults.Print references(n)
Next n
End Sub

Private Sub Command1_Click() 'back to main menu
frmreferences.Hide
frmselectschool.Show
End Sub
