VERSION 5.00
Begin VB.Form frmseven 
   BackColor       =   &H00800080&
   Caption         =   "Seven Dwarfs"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdScore 
      Caption         =   "I'm Done, See my Score!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3240
      TabIndex        =   3
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5280
      TabIndex        =   2
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmdseven 
      Caption         =   "Enter Names!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   1
      Top             =   6720
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "by Lisa Hammer and Kate Bancks"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   7920
      Width           =   2895
   End
   Begin VB.Label lblseven 
      BackStyle       =   0  'Transparent
      Caption         =   "The Seven Dwarfs LOVE Circus Fun...How Many            Of Them Can You Name?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   5295
      Left            =   360
      Picture         =   "seven.frx":0000
      Top             =   1320
      Width           =   6465
   End
End
Attribute VB_Name = "frmseven"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdExit_Click()             'the purpose of frmseven is for the user input as many of the seven dwarfs as possible.
    End                                 'this form is level 3 of the game.
                                        'the score is kept on this form as well, and added to the running total.
End Sub                                 'this button allows the user to quit the program

Private Sub cmdScore_Click()            'this button switches forms
    frmseven.Hide
    frmPoints.Show
    frmseven.Visible = False
    frmPoints.Visible = True
End Sub

Private Sub cmdseven_Click()            'this button reads a file into two parallel arrays and asks the user for input to a question and compares that to the answer array
    Dim DwarfArray(1 To 7) As String    'it also increments the score and gives relative feedback
    Dim DwarfRank(1 To 7) As Integer
    Dim searchvalue As String
    Dim I As Integer
    Dim pos As Integer
    Dim found As Boolean
    Dim arraysize As Integer
    arraysize = 7
    found = False
    pos = 0
    Open App.Path & "/seven.txt" For Input As #3
        Do Until EOF(3)
            pos = pos + 1
            Input #3, DwarfArray(pos), DwarfRank(pos)
        Loop
    Close #3
    pos = 0
    searchvalue = InputBox("Name a dwarf", "DWARF")
        Do While found = False And pos < arraysize
            pos = pos + 1
            If searchvalue = DwarfArray(pos) Then
                found = True
                C = C + 1
            End If
     Loop
        If found = True Then
            MsgBox "GOOD JOB!!", , "RIGHT!"
        Else
            MsgBox "TRY AGAIN", , "WHOOPS!!"
        End If

End Sub

        
