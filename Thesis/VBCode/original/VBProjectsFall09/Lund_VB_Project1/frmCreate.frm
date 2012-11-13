VERSION 5.00
Begin VB.Form frmCreate 
   Caption         =   "Create a Setilist"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8715
   LinkTopic       =   "Form2"
   ScaleHeight     =   7950
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResults 
      Height          =   7215
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   360
      Width           =   4815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Export Setlist"
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Return to Main Form"
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add with Segue"
      Height          =   615
      Left            =   1680
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Song"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Select a Song :"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
End
Attribute VB_Name = "frmCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ctr As Integer, songname(1 To 255) As String, outname As String
    

Private Sub Command1_Click()

    'Prints selection
    txtResults.Text = txtResults.Text & Combo1.Text & vbCrLf

    'Adds selection to Setlist array
    Setlist(a) = Combo1.Text
    
    'advances the setlist array
    a = a + 1

End Sub

Private Sub Command2_Click()
    'Prints selection with segway indication
    txtResults.Text = txtResults.Text & Combo1.Text & " >" & vbCrLf
    
    'Adds selection to Setlist array
    Setlist(a) = Combo1.Text & " >"
    
    'advances the setlist array
    a = a + 1

End Sub

Private Sub Command3_Click()

    'Brings the user back tot he main screen
    frmCreate.Hide

    frmMain.Show

End Sub

Private Sub Command4_Click()
   
   'opens file setlistoutput.txt to receive the setlist
   Open App.Path & "\setlistoutput.txt" For Output As #1
   
   'writes the setlist in the textbox to setlistoutput.txt
   Print #1, txtResults.Text
   
   'Closes channel
   Close #1
   
   'Confirms the file has been saved
   MsgBox ("The file has been saved")

End Sub
