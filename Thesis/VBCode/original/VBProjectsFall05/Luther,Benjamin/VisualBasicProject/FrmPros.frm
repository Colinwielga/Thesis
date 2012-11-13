VERSION 5.00
Begin VB.Form FrmPros 
   Caption         =   "SURF Professionals"
   ClientHeight    =   7170
   ClientLeft      =   5475
   ClientTop       =   1920
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   Palette         =   "FrmPros.frx":0000
   Picture         =   "FrmPros.frx":74C1
   ScaleHeight     =   7170
   ScaleWidth      =   5295
   Begin VB.CommandButton Cmdclose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close"
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6720
      Width           =   735
   End
   Begin VB.CommandButton CmdRank 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Arrange By Rank"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton CmdName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Arrange By Last Name"
      Height          =   495
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   1815
   End
   Begin VB.PictureBox picpros 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   600
      ScaleHeight     =   3675
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   2280
      Width           =   4095
   End
   Begin VB.Label lbldesciption 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmPros.frx":E982
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label Lblsurfers 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Current 2005 World Cup Tour Standings"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   4815
   End
End
Attribute VB_Name = "FrmPros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: SurfProject (SurfingProject.vbp)
'Form Name: frmDest (frmPros.frm)
'Author: Benjamin Luther
'Purpose of Form: The purpose of this form is to
                        'have the user click on a command button
                        'to see the current top surfers of the world.
                        'They are able to view them by their ranks
                        'or by the order of their last name.
Option Explicit 'Forces explicit declaration of all variables.
Dim I As Integer 'declares storage space for variables for the whole form
Dim ranks(1 To 100) As Integer, names(1 To 100) As String, country(1 To 100) As String ' sets each of the variables to their corresponding input up to 100 inputs

Private Sub cmdclose_Click() 'when command button is clicked
    FrmPros.Hide 'Hides the pros form
End Sub

Private Sub CmdName_Click()
    picpros.Cls 'clears the picture box
    I = 0 ' sets I to zero
    Open App.Path & "\prosurf.txt" For Input As #1 'opens the text file prosurf to be used as input
    picpros.Print Tab(4); "Name"; Tab(19); "Rank"; Tab(27); "Country" 'prints the headings for the picturebox
    Do Until EOF(1) 'the loop runs until it reaches the end of the file
        I = I + 1 'adds one to the value of I
        Input #1, ranks(I), names(I), country(I) 'Inputs from the text file prosurf the ranks, names, and place of origin from the surfers
        picpros.Print names(I); Tab(19); ranks(I); Tab(28), country(I) 'prints the names, ranks, and country in order from the text file
    Loop 'loops back to do until End of File
    Close #1 'Closes the text file prosurf
End Sub

Private Sub CmdRank_Click()
    Dim Temp As Integer, Pass As Integer, Temp1 As String, N As Integer ' declares storage space for variables
    picpros.Cls 'clears the picture box
    I = 0 ' sets variable I to 0
    Open App.Path & "\prosurf.txt" For Input As #1 'opens the text file prosurf for input
    picpros.Print "Rank"; Tab(11); "Name"; Tab(27); "Country" 'prints the headings for the picturebox
    Do Until EOF(1) 'ther loop runs until it reaches the end of the file
        I = I + 1 'adds one to the value of I
        Input #1, ranks(I), names(I), country(I) 'reads the values from the text file and sets them into the array
    Loop 'loops back to do until end of file
        For Pass = 1 To I - 1 'loops until pass is equal to the value I- 1
            For N = 1 To I - Pass 'Loops until value N is eqaul to I (value set when file was read in previous lines)
                If ranks(N) > ranks(N + 1) Then 'If the current value of rank is greater than the next value of rank then continue, if not it continues though the loop
                    Temp = ranks(N) 'sets the rank of N into the variable Temp
                    ranks(N) = ranks(N + 1) 'sets the ran of N to the next N in the array
                    ranks(N + 1) = Temp 'sets the next value of the array equal to the Temp value
                    Temp1 = names(N) 'Sets the variable of temp1 to the input of the current names array N
                    names(N) = names(N + 1) 'Sets the current names array N to the next names array value
                    names(N + 1) = Temp1 'Sets the next value of the names array N to the variable of temp1
                    Temp1 = country(N) 'Sets the variable of temp1 to the input of the current country array N
                    country(N) = country(N + 1) 'Sets the current country array N to the next country array value
                    country(N + 1) = Temp1 'Sets the next value of the country array N to the variable of temp1
                End If 'ends the If statement and continues the loop
            Next N 'continues to the next N in the loop
        Next Pass 'continues to the next Pass in the loop
    For N = 1 To I 'starts the loop for N, from 1 to the value of I (number of inputs from the text file)
            picpros.Print ranks(N); Tab(8); names(N); Tab(28), country(N) 'prints the current values if the ranks, names, and country in the correct order
    Next N 'continues loop to the next N
    Close #1 'closes the file prosurf.txt
End Sub

