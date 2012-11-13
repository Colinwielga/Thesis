VERSION 5.00
Begin VB.Form frmplayers 
   BackColor       =   &H0000FFFF&
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00800080&
      Caption         =   "Return to the Front Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   2895
   End
   Begin VB.CommandButton cmdsearch 
      BackColor       =   &H00800080&
      Caption         =   "Find Out!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   2895
   End
   Begin VB.TextBox txtSearch 
      Height          =   1695
      Left            =   2640
      TabIndex        =   2
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter a Name to Find out if that person is on the Vikings' Roster:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Is He On the Vikings?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmplayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Project Title: Minnesota Vikings Fan Page

'Form Name: Players

'Written by Jim Zrust

'Date: November 1, 2006

'Form objective: this form was intended to allow the user to input a name to see if they are on the Vikings roster.
'it can help show the user if their favorite player is still on the team, or it can simply be used for
'fun and the user could do something as simple as type their own name in and see what happens

Private Sub cmdReturn_Click() 'button to return to front page
frmplayers.Hide
frmhome.Show
End Sub

Private Sub cmdsearch_Click()
Dim N As String
Dim Roster(1 To 60) As String 'declare array
N = txtSearch.Text
I = 0
Found = False 'useful to make sure that the match was really found
Open App.Path & "\roster.txt" For Input As #1
For I = 1 To 60
    Input #1, Roster(I) 'fill the array
Next I
I = 0 'reset I back to 0
Do While ((Not Found) And (I < 60)) 'searches until found or end of list
    I = I + 1
    If N = Roster(I) Then Found = True
Loop
If (Not Found) Then
        MsgBox N & " is not on the Vikings", , "Sorry" 'if not found tell user through message box
    Else
        MsgBox N & " is on the Vikings", , "Congratulations" 'if found tell
End If
Close #1
End Sub

