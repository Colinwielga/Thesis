VERSION 5.00
Begin VB.Form FrmDominance 
   BackColor       =   &H000000FF&
   Caption         =   "Dominance Relations"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "FrmDominance.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdQuit3 
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
      Height          =   975
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton Cmdminimax 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click here to move on to minimax/maximin strategies"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   10560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
      Width           =   2055
   End
   Begin VB.CommandButton CmdClear 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clear Screen"
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
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton Cmd 
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
      Height          =   1455
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton Cmd2x2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click here to input your own 2x2 matrix"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   1935
   End
   Begin VB.PictureBox PicResults 
      BackColor       =   &H00FFFFFF&
      Height          =   4935
      Left            =   8040
      ScaleHeight     =   4875
      ScaleWidth      =   6915
      TabIndex        =   0
      Top             =   360
      Width           =   6975
   End
End
Attribute VB_Name = "FrmDominance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Introduction to Game Theory
'Form: FrmDominance
'Carson Sievert
'Aaron Skelly
'March 23, 2008
'This form explains and demonstrates the Dominance Theory. Here the user
'can read and experiment with the dominance theory and 2x2 matrices.


Private Sub Cmd_Click()
    FrmDominance.Hide   'This button moves the user from the Dominance page
    FrmGameTheory.Show  'back to the Game Theory page.
End Sub

'This is where the user will input the values he/she chooses for the matrices
'Once the values have been shown the program will evaluate which rows/colums dominate
'and whether or not there are saddle point(s) according to the values the user inputs.
Private Sub Cmd2x2_Click()
    Dim evenmatrix(1 To 2, 1 To 2) As Single
    Dim X1 As Single, X2 As Single, Y1 As Single, Y2 As Single, V As Single
    'This is where the user inputs values
    evenmatrix(1, 1) = InputBox("Enter any payoff value for the first (row) player that represents the first row, first column entry in the 2 x 2 matrix")
    evenmatrix(2, 1) = InputBox("Enter any payoff value for the first (row) player that represents the first row, second column entry in the 2 x 2 matrix")
    evenmatrix(1, 2) = InputBox("Enter any payoff value for the first (row) player that represents the second row, first column entry in the 2 x 2 matrix")
    evenmatrix(2, 2) = InputBox("Enter any payoff value for the first (row) player that represents the second row, second column entry in the 2 x 2 matrix")
    
    'Here the program print out a 2x2 matrix of the values the user entered.
    PicResults.Print "You have entered the following matrix => "; evenmatrix(1, 1); " "; evenmatrix(2, 1)
    PicResults.Print Tab(40); evenmatrix(1, 2); " "; evenmatrix(2, 2)

    If (evenmatrix(1, 1) = evenmatrix(1, 2)) And (evenmatrix(1, 1) = evenmatrix(2, 1)) And (evenmatrix(1, 1) = evenmatrix(2, 2)) Then
        MsgBox ("All four entries have the same value, since the payoffs will never change any strategy is a optimal strategy for both players, so all four entries are saddle points")
        'This statement is for when the user enters four of the same values.
    
    ElseIf (evenmatrix(1, 1) >= evenmatrix(1, 2)) And (evenmatrix(2, 1) >= evenmatrix(2, 2)) Then
        PicResults.Print "The row player is trying to maximize his/her payoff; since both of the first row entries are"
        PicResults.Print "greater than the respective entries in the bottom row:"
        PicResults.Print "The first row dominates the second row, so we are left with the row => "; evenmatrix(1, 1); " "; evenmatrix(2, 1)
        'Here the first row dominates the second row
        
        'This next if statement determines which column then dominates and where the saddle point is.
        If evenmatrix(1, 1) < evenmatrix(2, 1) Then
            PicResults.Print "The column player is trying to minimize the payoff entry (because their payoff is the"
            PicResults.Print "opposite of the entry). Since the first column entry is smaller than the second:"
            PicResults.Print "The first column dominates the second column, so there is a saddle point at the first"
            PicResults.Print "row and column and the value of the saddle point is "; evenmatrix(1, 1)
            'First column dominates the second
        
        ElseIf evenmatrix(1, 1) > evenmatrix(2, 1) Then
            PicResults.Print "The column player is trying to minimize the payoff entry (because their payoff is the"
            PicResults.Print "opposite of the entry). Since the second column entry is smaller than the first:"
            PicResults.Print "The second column dominates the first column, so there is a saddle point at the first row,"
            PicResults.Print "second column and the value of the saddle point is "; evenmatrix(2, 1)
            'Second column dominates the first
        
        ElseIf evenmatrix(1, 1) = evenmatrix(2, 1) Then
            PicResults.Print "The two remaining entries have the same payoff; therefore,"
            PicResults.Print "there is a saddle point at both entries and their value is "; evenmatrix(1, 1)
            'Payoff here is the same => there is no dominate column.
        End If
    
    ElseIf (evenmatrix(1, 1) <= evenmatrix(1, 2)) And (evenmatrix(2, 1) <= evenmatrix(2, 2)) Then
        PicResults.Print "The row player is trying to maximize the payoff; since both of the second row entries are"
        PicResults.Print "greater than the respective entries in the top row:"
        PicResults.Print "The second row dominates the first row, so we are left with the row =>"; evenmatrix(1, 2); " "; evenmatrix(2, 2)
        'Here the second row dominates the first
        
        'Same as above this next if statement determines the dominate column and saddle point.
        If evenmatrix(1, 2) < evenmatrix(2, 2) Then
            PicResults.Print "The column player is trying to minimize the payoff entry (because their payoff is the"
            PicResults.Print "opposite of the entry). Since the first column entry is smaller than the second:"
            PicResults.Print "The first column dominates the second column, so there is a saddle point at the second row,"
            PicResults.Print "first column and the value of the saddle point is "; evenmatrix(1, 2)
        
        ElseIf evenmatrix(1, 2) > evenmatrix(2, 2) Then
            PicResults.Print "The column player is trying to minimize the payoff entry (because their payoff is the"
            PicResults.Print "opposite of the entry). Since the second column entry is smaller than the first:"
            PicResults.Print "The second column dominates the first column, so there is a saddle point at the second row,"
            PicResults.Print "second column and the value of the saddle point is "; evenmatrix(2, 2)
        
        ElseIf evenmatrix(1, 2) = evenmatrix(2, 2) Then
            PicResults.Print "The two remaining entries have the same payoff; therefore,"
            PicResults.Print "there is a saddle point at both entries and their value is "; evenmatrix(2, 2)
        End If
    
    'Here the second column dominates the first column.
    ElseIf (evenmatrix(1, 1) >= evenmatrix(2, 1)) And (evenmatrix(1, 2) >= evenmatrix(2, 2)) Then
        PicResults.Print "The second column dominates the first column, so we are left with the column entries =>"; evenmatrix(2, 1)
        PicResults.Print Tab(84); evenmatrix(2, 2)
        
        'This if statement then determines the dominante row and the saddle point.
        If evenmatrix(2, 1) > evenmatrix(2, 2) Then
            PicResults.Print "The row player is trying to maximize the payoff entry. "
            PicResults.Print "Since the first row entry is greater than the second:"
            PicResults.Print "The first row dominates the second row, so there is a saddle point at the first row,"
            PicResults.Print "second column and the value of the saddle point is "; evenmatrix(2, 1)
        
        ElseIf evenmatrix(2, 1) < evenmatrix(2, 2) Then
            PicResults.Print "The row player is trying to maximize the payoff entry. "
            PicResults.Print "Since the second row entry is greater than the first:"
            PicResults.Print "The second row dominates the first row, so there is a saddle point at the second row,"
            PicResults.Print "second column and the value of the saddle point is " & evenmatrix(2, 2)
        
        ElseIf evenmatrix(2, 1) = evenmatrix(2, 2) Then
            PicResults.Print "The two remaining entries have the same payoff; therefore,"
            PicResults.Print "there is a saddle point at both entries and their value is "; evenmatrix(2, 2)
        End If
    
    'Here the first column dominates the second column.
    ElseIf (evenmatrix(1, 1) <= evenmatrix(2, 1)) And (evenmatrix(1, 2) <= evenmatrix(2, 2)) Then
        PicResults.Print "The first column dominates the second column, so we are left with the column entries =>"; evenmatrix(1, 1)
        PicResults.Print Tab(84); evenmatrix(1, 2)
        
        'Again this is where row dominance and the saddle points are found.
        If evenmatrix(1, 1) < evenmatrix(1, 2) Then
            PicResults.Print "The row player is trying to maximize the payoff entry. "
            PicResults.Print "Since the second row entry is greater than the first:"
            PicResults.Print "The second row dominates the first row, so there is a saddle point at the second row,"
            PicResults.Print "first column and the value of the saddle point is "; evenmatrix(1, 2)
        
        ElseIf evenmatrix(1, 1) > evenmatrix(1, 2) Then
            PicResults.Print "The row player is trying to maximize the payoff entry. "
            PicResults.Print "Since the first row entry is greater than the second:"
            PicResults.Print "The first row dominates the second row, so there is a saddle point at the first row,"
            PicResults.Print "first column and the value of the saddle point is "; evenmatrix(1, 1)
        
        ElseIf evenmatrix(1, 1) = evenmatrix(1, 2) Then
            PicResults.Print "The two remaining entries have the same payoff; therefore,"
            PicResults.Print "there is a saddle point at both entries and their value is "; evenmatrix(1, 2)
        End If
    
    Else 'This takes care of the case where no pure strategies or dominance relations exist.
        PicResults.Print "No dominance relations exist, so there is no pure strategy for the players."
        PicResults.Print "Although there is no pure strategy, every 2 x 2 matrix has a optimal mixed"
        PicResults.Print "strategies for both players.  Which means that the players should only play"
        PicResults.Print "a certain strategy a certain percent of the time in order to expect an optimal payoff."
        
        X1 = Abs(evenmatrix(2, 2) - evenmatrix(1, 2)) / Abs(evenmatrix(1, 1) - evenmatrix(2, 1) - evenmatrix(1, 2) + evenmatrix(2, 2))
        'Calculates the percent of the time the row player should play the first strategy.
        
        X2 = 1 - X1 'Calculates the percent of the time the row player should play the second strategy.
        
        Y1 = Abs(evenmatrix(2, 2) - evenmatrix(2, 1)) / Abs(evenmatrix(1, 1) - evenmatrix(2, 1) - evenmatrix(1, 2) + evenmatrix(2, 2))
        'Calculates the percent of the time the column player should play the first strategy.
        
        Y2 = 1 - Y1 'Calculates the percent of the time the column player should play the second strategy.
        
        V = ((evenmatrix(2, 2) * evenmatrix(1, 1)) - (evenmatrix(1, 2) * evenmatrix(2, 1))) / (evenmatrix(1, 1) - evenmatrix(2, 1) - evenmatrix(1, 2) + evenmatrix(2, 2))
        
        'Calculates the expected payoff for the optimal strategies.
        'The picresults here prints out a strategy for the row and column players while stating the saddle point.
        PicResults.Print "For the matrix above, the row player should play his/her first strategy"
        PicResults.Print FormatPercent(X1); " of the time and his/her second strategy "; FormatPercent(X2); " of the time."
        PicResults.Print "If the row player follows this strategy (for an infinite amount of times), he/she can expect"
        PicResults.Print "an average payoff greater than or equal to "; V
        PicResults.Print "The column player should play his/her first strategy"; FormatPercent(Y1)
        PicResults.Print "of the time and his/her second strategy "; FormatPercent(Y2); " of the time."
        PicResults.Print "If the column player follows this strategy (for an infinite amount of times), he/she can expect"
        PicResults.Print "an average payoff greater than or equal to "; -V
    End If
End Sub

Private Sub CmdClear_Click()
    PicResults.Cls  'clears the dominance print box
End Sub


Private Sub Cmdminimax_Click()
    FrmDominance.Hide   'moves the user to the mini max page
    Frmminmax.Show
End Sub

Private Sub CmdQuit_Click()
    End 'QUIT
End Sub

Private Sub CmdQuit3_Click()
    End
End Sub
