VERSION 5.00
Begin VB.Form Med 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   9735
   ClientLeft      =   255
   ClientTop       =   1035
   ClientWidth     =   15135
   BeginProperty Font 
      Name            =   "Modern No. 20"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9735
   ScaleWidth      =   15135
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF8080&
      Caption         =   "Home"
      Height          =   975
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8640
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000FF&
      Caption         =   "Easier"
      Height          =   975
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8640
      Width           =   1695
   End
   Begin VB.CommandButton BUT4 
      Caption         =   "D."
      Height          =   495
      Index           =   1
      Left            =   10320
      TabIndex        =   20
      Top             =   8640
      Width           =   495
   End
   Begin VB.CommandButton but3 
      Caption         =   "C."
      Height          =   495
      Index           =   1
      Left            =   10320
      TabIndex        =   19
      Top             =   8040
      Width           =   495
   End
   Begin VB.CommandButton but2 
      Caption         =   "B."
      Height          =   495
      Index           =   1
      Left            =   10320
      TabIndex        =   18
      Top             =   7440
      Width           =   495
   End
   Begin VB.CommandButton but1 
      Caption         =   "A."
      Height          =   495
      Index           =   1
      Left            =   10320
      TabIndex        =   17
      Top             =   6840
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enter"
      Height          =   735
      Left            =   9600
      TabIndex        =   14
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox enterbox 
      Height          =   615
      Left            =   9600
      TabIndex        =   13
      Top             =   4440
      Width           =   1215
   End
   Begin VB.PictureBox done 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12360
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   12
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox nums 
      Height          =   735
      Left            =   3000
      TabIndex        =   11
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton cmdGrade 
      Caption         =   "Answer"
      Height          =   735
      Left            =   3000
      TabIndex        =   10
      Top             =   7920
      Width           =   1815
   End
   Begin VB.CommandButton BUT4 
      Caption         =   "D."
      Height          =   495
      Index           =   0
      Left            =   3120
      TabIndex        =   7
      Top             =   5760
      Width           =   495
   End
   Begin VB.CommandButton BT3 
      Caption         =   "C."
      Height          =   495
      Index           =   0
      Left            =   3120
      TabIndex        =   6
      Top             =   5160
      Width           =   495
   End
   Begin VB.CommandButton bt2 
      Caption         =   "B."
      Height          =   495
      Index           =   0
      Left            =   3120
      TabIndex        =   5
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "A."
      Height          =   495
      Index           =   0
      Left            =   3120
      TabIndex        =   4
      Top             =   3960
      Width           =   495
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5160
      ScaleHeight     =   1035
      ScaleWidth      =   5475
      TabIndex        =   3
      Top             =   2160
      Width           =   5535
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1335
      Index           =   1
      Left            =   13200
      Picture         =   "Med.frx":0000
      ScaleHeight     =   1335
      ScaleWidth      =   1455
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1335
      Index           =   0
      Left            =   480
      Picture         =   "Med.frx":11EA
      ScaleHeight     =   1335
      ScaleWidth      =   1455
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   $"Med.frx":23D4
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   1
      Left            =   11040
      TabIndex        =   16
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   $"Med.frx":24E5
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Index           =   1
      Left            =   11040
      TabIndex        =   15
      Top             =   6840
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   "What was the most amount of people at a Twins Game?"
      Height          =   2295
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   6840
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   $"Med.frx":25C1
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Twins Trivia"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   84.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1695
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   9255
   End
End
Attribute VB_Name = "Med"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'dims variables
Option Explicit
Dim CTR As Single, runningtotal As Integer
'clears picResults and says answer is wrong
Private Sub bt2_Click(Index As Integer)
picResults.Cls

picResults.Print "Incorrect"
End Sub
'clears picResults and says answer is wrong
Private Sub BT3_Click(Index As Integer)
picResults.Cls

picResults.Print "Incorrect"
End Sub
'clears picResults and says answer is wrong
Private Sub BT4_Click()
picResults.Cls

picResults.Print "Incorrect"
End Sub
'clears picResults and says answer is wrong
Private Sub but1_Click(Index As Integer)
picResults.Cls

picResults.Print "Incorrect"
End Sub
'clears picResults and says answer is right and adds one to runningtotal
Private Sub but2_Click(Index As Integer)



picResults.Cls
picResults.Print "Correct"
done.Cls
runningtotal = runningtotal + 1
done.Print runningtotal

End Sub

'clears picResults and says answer is wrong
Private Sub but3_Click(Index As Integer)


picResults.Cls

picResults.Print "Incorrect"
End Sub
'clears picResults and says answer is wrong
Private Sub but4_Click(Index As Integer)


picResults.Cls

picResults.Print "Incorrect"
End Sub
'uses user answer and sees if it is right answer and adds one to runningtotal
Private Sub cmdGrade_Click()
Dim number As Single
 picResults.Cls
number = nums.Text


Select Case number
    Case Is < 100000   'if the user types in a number less than 100000, this message will pop up
        picResults.Print "That's not right."
       
    Case Is < 200000 'if the user types in a number less then 200000, this message will pop up
        picResults.Cls
        picResults.Print "Close but not right."
    Case Is < 300000 > 250000 'if the user types in a number less than 300000 and greater than 250000 more than 15, this message will pop up
        picResults.Cls
        picResults.Print "A little too high."
    Case Is = 250000 'if the user types in a number 250000, this message will pop up
        picResults.Cls
        picResults.Print "Exactly."
        runningtotal = runningtotal + 1
            done.Cls
            done.Print runningtotal
        
End Select
End Sub
'get users answer and declares  it is right and adds one to runningtotal
Private Sub Command1_Click()
Dim find As String, output As String
    
    'Clear the picturebox used for output
    picResults.Cls
    'get points from textbox and assign to variable
    
  
    find = enterbox.Text
    'assign the right or wrong
    If find = "D" Or find = "d" Then
            runningtotal = runningtotal + 1
            done.Cls
            done.Print runningtotal
            
            picResults.Print "Correct"
    
        Else: output = "Incorrect"
            picResults.Print "Incorrect"
    End If
End Sub

'clears picResults and says answer is right and adds one to runningtotal
Private Sub Command2_Click(Index As Integer)

picResults.Cls
picResults.Print "Correct"
done.Cls
runningtotal = runningtotal + 1
done.Print runningtotal
End Sub
'goes to easier page
Private Sub Command3_Click()
Easy.Show
Med.Hide

End Sub
'goes to home page
Private Sub Command4_Click()
main.Show
Med.Hide

End Sub
