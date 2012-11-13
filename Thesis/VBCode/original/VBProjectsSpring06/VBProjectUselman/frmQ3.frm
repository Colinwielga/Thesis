VERSION 5.00
Begin VB.Form frmQ3 
   BackColor       =   &H00008000&
   Caption         =   "Question 3"
   ClientHeight    =   9585
   ClientLeft      =   1350
   ClientTop       =   1125
   ClientWidth     =   12615
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9585
   ScaleWidth      =   12615
   Begin VB.CommandButton cmdQ3f 
      BackColor       =   &H00008000&
      Caption         =   "Go Back To Main Menu"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8280
      Width           =   1815
   End
   Begin VB.CommandButton cmdQ3e 
      BackColor       =   &H00008000&
      Caption         =   "Finish"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton cmdQ3d 
      BackColor       =   &H00008000&
      Caption         =   "Click Here To See The Approximate Weight Of The Lightest Animal"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6960
      Width           =   3375
   End
   Begin VB.CommandButton cmdQ3c 
      BackColor       =   &H00008000&
      Caption         =   "Click Here To See The Approximate Weight Of The Heaviest Animal"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5640
      Width           =   3375
   End
   Begin VB.CommandButton cmdQ3b 
      BackColor       =   &H00008000&
      Caption         =   "Click Here To See The Weight Of Each Animal From Lightest To Heaviest"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6840
      Width           =   3375
   End
   Begin VB.CommandButton cmdQ3a 
      BackColor       =   &H00008000&
      Caption         =   "Click Here If You Think You Know The Answer"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   3375
   End
   Begin VB.CommandButton cmdQ3 
      BackColor       =   &H00008000&
      Caption         =   "Click Here To Load List"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   3375
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   2040
      ScaleHeight     =   3555
      ScaleWidth      =   8595
      TabIndex        =   0
      Top             =   1680
      Width           =   8655
   End
   Begin VB.Label lblMe 
      BackColor       =   &H00008000&
      Caption         =   "Lance Uselman"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   9000
      Width           =   1575
   End
   Begin VB.Label lblq3 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Which of these animals, on average, weighs the most?"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      TabIndex        =   1
      Top             =   480
      Width           =   4695
   End
End
Attribute VB_Name = "frmQ3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Wildlife Challenge (Project1.vbp)
'frmQ3 (frmQ3.frm)
'Lance Uselman
'March 24, 2006
'Purpose of the form: This form asks the user to load a list of animals into a
'                     picture box and then asks to input the name of one of the
'                     animals into an input box. A message associated with the
'                     answer is printed in the picture box. The user can also
'                     view the lightest and heaviest animals as well as the list
'                     sorted by weight.

Option Explicit
    Dim arrayName(1 To 9) As String
    Dim arrayLbs(1 To 9) As Single

Private Sub cmdQ3e_Click()
    frmComplete.Show
    frmQ3.Hide  'This button allows the user to finish the quiz.
End Sub

Private Sub cmdQ3_Click()
    Dim pos, size As Integer    'This button loads a list into arrays and prints them in a picture box.
    picResults.Cls
    picResults.Print "Common Name", "Approximate Weight (LBS)"
    picResults.Print "******************************************************************"
    Open App.Path & "\weights.txt" For Input As #1
        For pos = 1 To 9
            Input #1, arrayName(pos), arrayLbs(pos)
            picResults.Print arrayName(pos), , , "?"
        Next pos
    Close #1
    size = pos
    picResults.Print
End Sub

Private Sub cmdQ3a_Click()
    Dim A As String 'This button asks the user to input an answer into a input box and then prints the associated message in the picture box.
    picResults.Cls
    A = InputBox("Enter the name of the animal you believe is the heaviest.", "Heaviest Animal")
    If A = "Moose" Then
        picResults.Print "Correct!  Moose is the correct answer."
    Else
        picResults.Print "Incorrect.  Please try again."
    End If
End Sub

Private Sub cmdQ3b_Click()
    Dim pos, Pass, Temp, size As Integer    'This button loads the names and weights of the animals from a file into arrays and then prints them in a picture box sorted from lightest to heaviest.
    picResults.Cls
    picResults.Print "Common Name", "Approximate Weight (LBS)"
    picResults.Print "******************************************************************"
    pos = 0
    Open App.Path & "\weights.txt" For Input As #1
    Do Until EOF(1)
        pos = pos + 1
        Input #1, arrayName(pos), arrayLbs(pos)
    Loop
    Close #1
    size = pos
    For Pass = 1 To (size - 1)
        For pos = 1 To (size - Pass)
            If arrayLbs(pos) > arrayLbs(pos + 1) Then
                Temp = arrayLbs(pos)
                arrayLbs(pos) = arrayLbs(pos + 1)
                arrayLbs(pos + 1) = Temp
                Temp = arrayName(pos)
                arrayName(pos) = arrayName(pos + 1)
                arrayName(pos + 1) = Temp
            End If
        Next pos
    Next Pass
    For pos = 1 To size
        picResults.Print arrayName(pos), , , arrayLbs(pos)
    Next pos
End Sub

Private Sub cmdQ3c_Click()
    Dim Max As Single   'This button finds the largest weight on the list and prints it in a picture box.
    Dim pos, size As Integer
    picResults.Cls
    Open App.Path & "\weights.txt" For Input As #1
    pos = 0
    Do Until EOF(1)
        pos = pos + 1
        Input #1, arrayName(pos), arrayLbs(pos)
    Loop
    Close #1
    Max = arrayLbs(1)
    size = pos
    For pos = 1 To size
        If Max < arrayLbs(pos) Then
            Max = arrayLbs(pos)
        End If
    Next pos
    picResults.Print "The heaviest animal on this list weighs approximately"; Max; "pounds."
End Sub

Private Sub cmdQ3d_Click()
    Dim Min As Single   'This button finds the lightest weight on the list and prints it in a picture box.
    Dim pos, size As Integer
    picResults.Cls
    Open App.Path & "\weights.txt" For Input As #1
    pos = 0
    Do Until EOF(1)
        pos = pos + 1
        Input #1, arrayName(pos), arrayLbs(pos)
    Loop
    Close #1
    Min = arrayLbs(1)
    size = pos
    For pos = 1 To size
        If Min > arrayLbs(pos) Then
            Min = arrayLbs(pos)
        End If
    Next pos
    picResults.Print "The lightest animal on this list weighs approximately"; Min; "pounds."
End Sub

Private Sub cmdQ3f_Click()
    frmMain.Show
    frmQ3.Hide  'This button allows the user to go back to the main form.
End Sub

