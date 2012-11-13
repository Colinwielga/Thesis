VERSION 5.00
Begin VB.Form Frmminmax 
   BackColor       =   &H000000FF&
   Caption         =   "minimax/maximin strategies"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdQuit2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8880
      Width           =   3615
   End
   Begin VB.PictureBox Picture1 
      Height          =   5655
      Left            =   9840
      Picture         =   "Frmminmax.frx":0000
      ScaleHeight     =   5595
      ScaleWidth      =   5115
      TabIndex        =   12
      Top             =   4200
      Width           =   5175
   End
   Begin VB.CommandButton CmdDominance 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back to Dominance Theory Page"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2280
      Width           =   1575
   End
   Begin VB.PictureBox PicResults 
      BackColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   6120
      ScaleHeight     =   3795
      ScaleWidth      =   6075
      TabIndex        =   10
      Top             =   240
      Width           =   6135
   End
   Begin VB.CommandButton Cmdminimax 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click to calculate the maximin and minimax values of your matrix"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7560
      Width           =   5175
   End
   Begin VB.TextBox Txtfour 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4080
      TabIndex        =   8
      Top             =   5760
      Width           =   1815
   End
   Begin VB.TextBox Txtthree 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4080
      TabIndex        =   7
      Top             =   3960
      Width           =   1815
   End
   Begin VB.TextBox Txttwo 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4080
      TabIndex        =   6
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox Txtone 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4080
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Cmdgoback 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back to Game Theory Main Page"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Lblentry4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter any payoff value for the first (row) player that represents the second row, second column entry in the 2 x 2 matrix =>"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      TabIndex        =   4
      Top             =   5760
      Width           =   3135
   End
   Begin VB.Label Lblentry3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter any payoff value for the first (row) player that represents the second row, first column entry in the 2 x 2 matrix =>"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      TabIndex        =   3
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Label Lblentry2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter any payoff value for the first (row) player that represents the first row, second column entry in the 2 x 2 matrix =>"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label Lblentry1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter any payoff value for the first (row) player that represents the first row, first column entry in the 2 x 2 matrix =>"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Frmminmax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Introduction to Game Theory
'Form: Frmminmax
'Carson Sievert
'Aaron Skelly
'March 23, 2008
'This form explains and demonstrates the conept of minimax/maximin with 2x2
'matrices. Saddle points can also be found here.

Private Sub CmdDominance_Click()
Frmminmax.Hide  'moves user to dominance page
FrmDominance.Show

End Sub

Private Sub Cmdgoback_Click()
    Frmminmax.Hide  'moves user to the Game Theory
    FrmGameTheory.Show
End Sub

Private Sub Cmdminimax_Click()
    'This command button creates a 2x2 matrix with values input by the user
    'and then determines the minimax/maximin and it will state the saddle
    'point if there is one.
    Dim oddmatrix(1 To 2, 1 To 2) As Single
    Dim largest(1 To 2) As Single, smallest(1 To 2) As Single
    Dim minimax As Single, maximin As Single
    Dim row As Integer, column As Integer
    
    'Here the user enters the matrix values with a text boxes.
    oddmatrix(1, 1) = Txtone.Text
    oddmatrix(2, 1) = Txttwo.Text
    oddmatrix(1, 2) = Txtthree.Text
    oddmatrix(2, 2) = Txtfour.Text
    'Shows the user the matrix he/she has inputed
    PicResults.Print "You have inputed the matrix=>"; oddmatrix(1, 1); " "; oddmatrix(2, 1)
    PicResults.Print Tab(30); oddmatrix(1, 2); " "; oddmatrix(2, 2)
    PicResults.Print
    Select Case oddmatrix(1, 1)        'this case finds the 1st column maxima
        Case Is >= oddmatrix(1, 2)
            largest(1) = oddmatrix(1, 1)
        Case Is < oddmatrix(1, 2)
            largest(1) = oddmatrix(1, 2)
    End Select
    Select Case oddmatrix(2, 1)        'this case finds the 2nd column maxima
        Case Is >= oddmatrix(2, 2)
            largest(2) = oddmatrix(2, 1)
        Case Is < oddmatrix(2, 2)
            largest(2) = oddmatrix(2, 2)
    End Select
    Select Case oddmatrix(1, 1)        'this case finds the 1st row minima
        Case Is >= oddmatrix(2, 1)
            smallest(1) = oddmatrix(2, 1)
        Case Is < oddmatrix(2, 1)
            smallest(1) = oddmatrix(1, 1)
    End Select
    Select Case oddmatrix(1, 2)        'this case finds the 2nd row minima
        Case Is >= oddmatrix(2, 2)
            smallest(2) = oddmatrix(2, 2)
        Case Is < oddmatrix(2, 2)
            smallest(2) = oddmatrix(1, 2)
    End Select
    Select Case largest(1)             'this case finds the minimax
        Case Is >= largest(2)
            minimax = largest(2)
        Case Is < largest(2)
            minimax = largest(1)
    End Select
    Select Case smallest(1)            'this case finds the maximin
        Case Is >= smallest(2)
            maximin = smallest(1)
        Case Is < smallest(2)
            maximin = smallest(2)
    End Select
    'The following print statements explain how minimax and maximin strategies
    'are computed and used to find any saddle points.
    PicResults.Print "Mini-max and maxi-min strategies are another method used to find pure strategies."
    PicResults.Print "The mini-max strategy is used to describe the strategy the column player should"
    PicResults.Print "follow and is computed by taking the maximum payoff from each column entry."
    PicResults.Print "In this case, the two column maxima are =>"; largest(1); largest(2)
    PicResults.Print "Then the minimum of those payoffs are used for the value of the minimax strategy."
    'The minimax value is printed here
    PicResults.Print "In this case, the minimax strategy has a value of "; minimax
    PicResults.Print
    PicResults.Print "The maxi-min strategy is used to describe the strategy the row player should"
    PicResults.Print "follow, and it is computed by taking the minimum payoff from each row entry."
    PicResults.Print "In this case, the two row minima are =>"; smallest(1); smallest(2)
    PicResults.Print "Then the maximum of those payoffs are used for the value of the maxi-min strategy."
    'The maximin value is printed here
    PicResults.Print "In this case, the maximin has a value of "; maximin
    
    'Here the value of the saddle point is stated if the minimax=maximin
    'and if minimax and maximin not are equal then there is no saddle point.
    If minimax = maximin Then
        PicResults.Print "Therefore; minimax=maximin, so there is a saddle point"
        PicResults.Print "at every entry of the matrix that has a value of "; maximin
    ElseIf minimax > maximin Then
        PicResults.Print "Therefore; minimax>maximin, so no saddle point exists in this case"
    ElseIf minimax < maximin Then
        PicResults.Print "Therefore; minimax<maximin, so no saddle point exists in this case"
    End If
End Sub

Private Sub CmdQuit2_Click()
End
End Sub
