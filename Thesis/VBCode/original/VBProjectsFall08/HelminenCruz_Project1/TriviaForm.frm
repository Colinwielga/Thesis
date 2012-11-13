VERSION 5.00
Begin VB.Form NameOptions 
   Caption         =   "Form2"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9600
   LinkTopic       =   "Form2"
   Picture         =   "TriviaForm.frx":0000
   ScaleHeight     =   6975
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton NameOptions 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sort by percentage frequency in ascending order"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sort by  percentage frequency in ascending order"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton cmdmainpage 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Go to Main Page"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmdfrequency2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sort by percentage frequency in descending order"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton cmdfrequency 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sort by  percentage frequency in decending order"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton cmdascending2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sort Alphabetically in descending order"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdascending1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sort Alphabetically in ascending order"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmddescending 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sort Alphabetically in descending order"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdascending 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sort Alphabetically in ascending order"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4680
      Width           =   1695
   End
   Begin VB.PictureBox picresults2 
      BackColor       =   &H00FFC0C0&
      Height          =   3495
      Left            =   6000
      ScaleHeight     =   3435
      ScaleWidth      =   2715
      TabIndex        =   4
      Top             =   1080
      Width           =   2775
   End
   Begin VB.PictureBox picresults1 
      BackColor       =   &H00FFC0C0&
      Height          =   3495
      Left            =   1200
      ScaleHeight     =   3435
      ScaleWidth      =   2715
      TabIndex        =   3
      Top             =   1080
      Width           =   2775
   End
   Begin VB.CommandButton cmdmale 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Read Male Names"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton cmdfemale 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Read Female Names"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      MaskColor       =   &H00FFC0FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name Options"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "NameOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form inputs name options for pets. the user is also
'able to sort according to name in ascending order and
'descending order. and also according to the percentage
'of usage of the names. from the most common to the least
'common.



Option Explicit
Dim Femalenames(1 To 19) As String
Dim femalefrequency(1 To 19) As String
Dim malenames(1 To 19) As String
Dim malefrequency(1 To 19) As String
Dim ctr As Single
Dim N As Integer
Dim pos As Integer
Dim pass As Integer
Dim tempfemalenames As String
Dim tempfemalefrequency As String
Dim tempmalenames As String
Dim tempmalefrequency As String



Private Sub cmdascending_Click()
picresults1.Cls
picresults1.Print "Name", "% Frequency"
picresults1.Print "____________________________________"

For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If Femalenames(pos) > Femalenames(pos + 1) Then
            tempfemalenames = Femalenames(pos)
            Femalenames(pos) = Femalenames(pos + 1)
            Femalenames(pos + 1) = tempfemalenames

            tempfemalefrequency = femalefrequency(pos)
            femalefrequency(pos) = femalefrequency(pos + 1)
            femalefrequency(pos + 1) = tempfemalefrequency
        End If
    Next pos
Next pass

For N = 1 To ctr
    picresults1.Print Tab(1); Femalenames(N); Tab(25); femalefrequency(N)
Next N

End Sub

Private Sub cmdascending1_Click()
picresults2.Cls
picresults2.Print "Name", "% Frequency"
picresults2.Print "____________________________________"

For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If malenames(pos) > malenames(pos + 1) Then
            tempmalenames = malenames(pos)
            malenames(pos) = malenames(pos + 1)
            malenames(pos + 1) = tempmalenames

            tempmalefrequency = malefrequency(pos)
            malefrequency(pos) = malefrequency(pos + 1)
            malefrequency(pos + 1) = tempmalefrequency
        End If
    Next pos
Next pass

For N = 1 To ctr
    picresults2.Print Tab(1); malenames(N); Tab(25); malefrequency(N)
Next N

End Sub

Private Sub cmdascending2_Click()
picresults2.Cls
picresults2.Print "Name", "% Frequency"
picresults2.Print "____________________________________"

For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If malenames(pos) < malenames(pos + 1) Then
            tempmalenames = malenames(pos)
            malenames(pos) = malenames(pos + 1)
            malenames(pos + 1) = tempmalenames

            tempmalefrequency = malefrequency(pos)
            malefrequency(pos) = malefrequency(pos + 1)
            malefrequency(pos + 1) = tempmalefrequency
        End If
    Next pos
Next pass

For N = 1 To ctr
    picresults2.Print Tab(1); malenames(N); Tab(25); malefrequency(N)
Next N

End Sub


Private Sub cmddescending_Click()
picresults1.Cls
picresults1.Print "Name", "% Frequency"
picresults1.Print "____________________________________"

For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If Femalenames(pos) < Femalenames(pos + 1) Then
            tempfemalenames = Femalenames(pos)
            Femalenames(pos) = Femalenames(pos + 1)
            Femalenames(pos + 1) = tempfemalenames

            tempfemalefrequency = femalefrequency(pos)
            femalefrequency(pos) = femalefrequency(pos + 1)
            femalefrequency(pos + 1) = tempfemalefrequency
        End If
    Next pos
Next pass

For N = 1 To ctr
    picresults1.Print Tab(1); Femalenames(N); Tab(25); femalefrequency(N)
Next N

End Sub

Private Sub cmdfemale_Click()
ctr = 0
picresults1.Cls
picresults1.Print "Name", "% frequency"
picresults1.Print "____________________________________"

Open App.Path & "\femalenames.txt" For Input As #1

    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, Femalenames(ctr), femalefrequency(ctr)
        picresults1.Print Tab(2); Femalenames(ctr); Tab(25); femalefrequency(ctr)
    Loop

Close #1

End Sub

Private Sub cmdfrequency_Click()
picresults1.Cls
picresults1.Print "Name", "% Frequency"
picresults1.Print "____________________________________"

For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If femalefrequency(pos) > femalefrequency(pos + 1) Then
            tempfemalefrequency = femalefrequency(pos)
            femalefrequency(pos) = femalefrequency(pos + 1)
            femalefrequency(pos + 1) = tempfemalefrequency

            tempfemalenames = Femalenames(pos)
            Femalenames(pos) = Femalenames(pos + 1)
            Femalenames(pos + 1) = tempfemalenames
        End If
    Next pos
Next pass

For N = 1 To ctr
    picresults1.Print Tab(1); Femalenames(N); Tab(25); femalefrequency(N)
Next N

End Sub

Private Sub cmdfrequency2_Click()
picresults2.Cls
picresults2.Print "Name", "% Frequency"
picresults2.Print "____________________________________"

For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If malefrequency(pos) > malefrequency(pos + 1) Then
            tempmalefrequency = malefrequency(pos)
            malefrequency(pos) = malefrequency(pos + 1)
            malefrequency(pos + 1) = tempmalefrequency

            tempmalenames = malenames(pos)
            malenames(pos) = malenames(pos + 1)
            malenames(pos + 1) = tempmalenames
        End If
    Next pos
Next pass

For N = 1 To ctr
    picresults2.Print Tab(1); malenames(N); Tab(25); malefrequency(N)
Next N

End Sub


Private Sub cmdmainpage_Click()

Form2.Hide
Welcomeform2.Show

End Sub

Private Sub cmdmale_Click()
ctr = 0
picresults2.Cls
picresults2.Print "Name", "% Frequency"
picresults2.Print "____________________________________"

Open App.Path & "\malenames.txt" For Input As #1

    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, malenames(ctr), malefrequency(ctr)
        picresults2.Print Tab(2); malenames(ctr); Tab(25); malefrequency(ctr)
    Loop

Close #1

End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub Command1_Click()
picresults1.Cls
picresults1.Print "Name", "% Frequency"
picresults1.Print "____________________________________"

For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If femalefrequency(pos) < femalefrequency(pos + 1) Then
            tempfemalefrequency = femalefrequency(pos)
            femalefrequency(pos) = femalefrequency(pos + 1)
            femalefrequency(pos + 1) = tempfemalefrequency

            tempfemalenames = Femalenames(pos)
            Femalenames(pos) = Femalenames(pos + 1)
            Femalenames(pos + 1) = tempfemalenames
        End If
    Next pos
Next pass
For N = 1 To ctr

picresults1.Print Tab(1); Femalenames(N); Tab(25); femalefrequency(N)
Next N
End Sub

Private Sub Command2_Click()
picresults2.Cls
picresults2.Print "Name", "% Frequency"
picresults2.Print "____________________________________"

For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If malefrequency(pos) < malefrequency(pos + 1) Then
            tempmalefrequency = malefrequency(pos)
            malefrequency(pos) = malefrequency(pos + 1)
            malefrequency(pos + 1) = tempmalefrequency

            tempmalenames = malenames(pos)
            malenames(pos) = malenames(pos + 1)
            malenames(pos + 1) = tempmalenames
        End If
    Next pos
Next pass

For N = 1 To ctr
    picresults2.Print Tab(2); malenames(N); Tab(25); malefrequency(N)
Next N

End Sub

Private Sub NameOptions_Click()
picresults2.Cls
picresults2.Print "Name", "% Frequency"
picresults2.Print "____________________________________"

For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If malefrequency(pos) < malefrequency(pos + 1) Then
            tempmalefrequency = malefrequency(pos)
            malefrequency(pos) = malefrequency(pos + 1)
            malefrequency(pos + 1) = tempmalefrequency

            tempmalenames = malenames(pos)
            malenames(pos) = malenames(pos + 1)
            malenames(pos + 1) = tempmalenames
        End If
    Next pos
Next pass

For N = 1 To ctr
    picresults2.Print Tab(1); malenames(N); Tab(25); malefrequency(N)
Next N

End Sub

Private Sub picresults1_Click()

End Sub
