VERSION 5.00
Begin VB.Form frmSJUscores 
   BackColor       =   &H00004000&
   Caption         =   "2005 SJU Rugby Scores"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit Program"
      BeginProperty Font 
         Name            =   "Baskerville MT for Brill 01 SC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   4
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton cdmsort 
      Caption         =   "Sort Scores Alphabetically"
      BeginProperty Font 
         Name            =   "Baskerville MT for Brill 01 SC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7440
      TabIndex        =   3
      Top             =   5520
      Width           =   2055
   End
   Begin VB.CommandButton cmddisplay 
      Caption         =   "Display Scores"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Baskerville MT for Brill 01 SC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7440
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.PictureBox picbox 
      Height          =   5655
      Left            =   3000
      ScaleHeight     =   5595
      ScaleWidth      =   4155
      TabIndex        =   1
      Top             =   1320
      Width           =   4215
   End
   Begin VB.CommandButton cdmreturntomainmenu 
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Baskerville MT for Brill 01 SC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8280
      TabIndex        =   0
      Top             =   6960
      Width           =   2055
   End
   Begin VB.Image imgrugby5 
      Height          =   2520
      Left            =   7440
      Picture         =   "frmrules.frx":0000
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Image imgrugby3 
      Height          =   1800
      Left            =   240
      Picture         =   "frmrules.frx":11E62
      Top             =   4560
      Width           =   2520
   End
   Begin VB.Image imgrugby2 
      Height          =   1875
      Left            =   240
      Picture         =   "frmrules.frx":20AE4
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Image imgrugby1 
      Height          =   1905
      Left            =   240
      Picture         =   "frmrules.frx":2F77A
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmSJUscores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : RugbyVBProject (Rugby.vbp)
'Form Name : frmSJUscores(frmSJUscores.frm)
'Author: Alex Forrest
'purpose of the form: This form is designed to display the scores of the SJU rugby
    'team this fall.  The user can also click on another button to sort these scores
    'alphabetically in terms of the opponent.

Option Explicit
Dim sjuarray(1 To 10) As String
Dim opponentarray(1 To 10) As String
Private Sub cdmreturntomainmenu_Click()
    frmSJUscores.Hide
    frmMainmenu.Show 'takes the user back to the main menu
End Sub

Private Sub cdmsort_Click()
    Dim pass As Integer
    Dim X As Integer, r As Integer
    Dim temp As String
    Dim tempop As String
    Dim I As Single
    I = 10 'sets the array to 10
    For pass = 1 To I - 1 'starts the sorting process by making first pass
        For X = 1 To I - pass 'starts comparing each part of the array
            If opponentarray(X) > opponentarray(X + 1) Then 'begins if case to sort array
                temp = sjuarray(X) 'uses a holding spot to match up the correct information
                sjuarray(X) = sjuarray(X + 1) 'sets each spot to the correct information after being compared
                sjuarray(X + 1) = temp
                tempop = opponentarray(X) 'uses a holding spot for the other array
                opponentarray(X) = opponentarray(X + 1) 'sets each spot to the correct information after being compared
                opponentarray(X + 1) = tempop
            End If 'ends the if statement
        Next X ' peforms the above procedure for the next part of the array
    Next pass 'goes to make the next pass
    picbox.Print
    picbox.Print "*****************************************************************************"
    picbox.Print "The scores in alphabetical order are:"
    picbox.Print
    For r = 1 To 10 'begins the loops to print each array
        picbox.Print sjuarray(r); Tab(18); opponentarray(r) 'prints both arrays sorted by alphabetical order in terms of the opponent
    Next r 'goes to the next part of each array to print them
End Sub

Private Sub cmddisplay_Click()
Dim I As Single
Dim j As Single
picbox.Cls

I = 0
    Open App.Path & "\sjuscores.txt" For Input As #1 'opens the corresponding file in Path = M:\CSI130\VB Project
    Do Until EOF(1) 'instructs the program to do the following procedure until the end of the file
        I = I + 1
        Input #1, sjuarray(I), opponentarray(I) 'sets the whole file equal to input #1
    Loop 'loops and performs the same procedure until the end of the file
    picbox.Print "2005 SJU Rugby Scores:"
    picbox.Print
    For j = 1 To I 'begins the loops to print
        picbox.Print sjuarray(j), opponentarray(j) 'prints the corresponding part of the array
    Next j 'goes to next part of array to print
    Close #1
    picbox.Print
    picbox.Print "2005 Minnesota State Rugby Champs!"
End Sub

Private Sub cmdquit_Click()
    End 'ends the program
End Sub
