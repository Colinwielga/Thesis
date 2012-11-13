VERSION 5.00
Begin VB.Form Easy 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   8880
   ClientLeft      =   2025
   ClientTop       =   2010
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   11955
   Begin VB.CommandButton But4 
      Caption         =   "D."
      Height          =   495
      Index           =   1
      Left            =   8160
      TabIndex        =   22
      Top             =   5760
      Width           =   495
   End
   Begin VB.CommandButton but3 
      Caption         =   "C."
      Height          =   495
      Index           =   1
      Left            =   8160
      TabIndex        =   21
      Top             =   5160
      Width           =   495
   End
   Begin VB.CommandButton but2 
      Caption         =   "B."
      Height          =   495
      Index           =   1
      Left            =   8160
      TabIndex        =   20
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton but1 
      Caption         =   "A."
      Height          =   495
      Index           =   1
      Left            =   8160
      TabIndex        =   19
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox enterbox 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   17
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txtfindOne 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   16
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "Harder"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1560
      Width           =   1095
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
      Left            =   10800
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   13
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdoutput 
      Caption         =   "Answer"
      Default         =   -1  'True
      Height          =   735
      Left            =   6960
      TabIndex        =   12
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton BT4 
      Caption         =   "D."
      Height          =   495
      Index           =   0
      Left            =   3000
      TabIndex        =   10
      Top             =   7920
      Width           =   495
   End
   Begin VB.CommandButton BT3 
      Caption         =   "C."
      Height          =   495
      Index           =   0
      Left            =   3000
      TabIndex        =   9
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton bt2 
      Caption         =   "B."
      Height          =   495
      Index           =   0
      Left            =   3000
      TabIndex        =   8
      Top             =   6720
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "A."
      Height          =   495
      Index           =   0
      Left            =   3000
      TabIndex        =   7
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enter"
      Height          =   735
      Left            =   2880
      TabIndex        =   5
      Top             =   4920
      Width           =   1215
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4680
      ScaleHeight     =   795
      ScaleWidth      =   3075
      TabIndex        =   4
      Top             =   2880
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      Height          =   1455
      Index           =   1
      Left            =   10560
      Picture         =   "Easy.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   3
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      Height          =   1455
      Index           =   0
      Left            =   0
      Picture         =   "Easy.frx":17F9
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   2
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      Caption         =   $"Easy.frx":2FF2
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
      Left            =   8880
      TabIndex        =   18
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   "In which ballpark do the Twins play in? "
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   8880
      TabIndex        =   11
      Top             =   6360
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      Caption         =   $"Easy.frx":30D8
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
      Left            =   0
      TabIndex        =   6
      Top             =   6120
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Twins Trivia"
      BeginProperty Font 
         Name            =   "Niagara Engraved"
         Size            =   150
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   1440
      TabIndex        =   1
      Top             =   0
      Width           =   9135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   $"Easy.frx":31AB
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   3600
      Width           =   2775
   End
End
Attribute VB_Name = "Easy"
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
Private Sub BT4_Click(Index As Integer)
picResults.Cls
picResults.Print "Incorrect"
End Sub
'clears picResults and says answer is wrong
Private Sub but1_Click(Index As Integer)
picResults.Cls

picResults.Print "Incorrect"
End Sub
'clears picResults and says answer is wrong
Private Sub but2_Click(Index As Integer)
picResults.Cls

picResults.Print "Incorrect"
End Sub
'clears picResults and says answer is right and adds one to runningtotal
Private Sub but3_Click(Index As Integer)
picResults.Cls
picResults.Print "Correct"
done.Cls
runningtotal = runningtotal + 1
done.Print runningtotal
End Sub
'clears picResults and says answer is wrong
Private Sub but4_Click(Index As Integer)
picResults.Cls

picResults.Print "Incorrect"
End Sub
'goes through answer and if it is correct adds on to counter or displays incorrect in picResults
Private Sub cmdoutput_Click()
'declare the variables used
    Dim find As String, Grade As String
    
    'Clear the picturebox used for output
    picResults.Cls
    'get points from textbox and assign to variable
    
    find = txtfindOne.Text
    
    'assign the correct name
    If find = "Metrodome" Or find = "dome" Then
            Grade = "Correct"
            runningtotal = runningtotal + 1
            done.Cls
            done.Print runningtotal
            
        Else: Grade = "Incorrect"
    End If

    picResults.Print Grade
End Sub
'goes through answer and if it is correct adds on to counter or displays incorrect in picResults
Private Sub Command1_Click()
Dim find As String, output As String
    
    'Clear the picturebox used for output
    picResults.Cls
    'get points from textbox and assign to variable
    
  
    find = enterbox.Text
    'assign the right or wrong
    If find = "True" Or find = "true" Then
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


    


'hides easy form and goes back to main page
Private Sub Command3_Click()
Easy.Hide
main.Show

End Sub
'hides easy form and goes to med form
Private Sub Command4_Click()
Easy.Hide
Med.Show

End Sub
