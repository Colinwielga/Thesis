VERSION 5.00
Begin VB.Form frmSocial 
   BackColor       =   &H00FF8080&
   Caption         =   "Social Psychology Section"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11220
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF8080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   11220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton cmdReview 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Review more information on the Self"
      BeginProperty Font 
         Name            =   "Lithos Pro Regular"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   2415
   End
   Begin VB.PictureBox picOutput 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Lithos Pro Regular"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   2640
      ScaleHeight     =   5415
      ScaleWidth      =   8415
      TabIndex        =   3
      Top             =   1440
      Width           =   8415
   End
   Begin VB.CommandButton cmdUnrealisticOpt 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Self-Serving Bias"
      BeginProperty Font 
         Name            =   "Lithos Pro Regular"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton cmdImplicitEgo 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Self-Schema"
      BeginProperty Font 
         Name            =   "Lithos Pro Regular"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblMainSelf 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Social Psychology and The Self"
      BeginProperty Font 
         Name            =   "Lithos Pro Regular"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   9855
   End
End
Attribute VB_Name = "frmSocial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdImplicitEgo_Click()
'This command discusses implicit egotism and the self schema
    
    'declare variables
    Dim FaveLetter As String
    Dim FaveNum As Integer
    Dim FoundName As Boolean, FoundNum As Boolean
    FoundName = False
    FoundNum = False
    MsgBox "The self schema is how we organize information about ourselves.", , "Self Schema"
    FaveLetter = InputBox("Enter your favorite letter of the alphabet", "Favorite Letter")
    FaveNum = InputBox("Enter your favorite number between 0-99", "Favorite Number")
    
    'determine if favorite letter and number match any in the user's first and last name and birthday
    If InStr(First, FaveLetter) <> 0 Then
        FoundName = True
    End If
    If InStr(Last, FaveLetter) <> 0 Then
        FoundName = True
    End If
    If FaveNum = BMonth Then
        FoundNum = True
    End If
    If FaveNum = BDay Then
        FoundNum = True
    End If
    If FaveNum = BYear Then
        FoundNum = True
    End If
    
    'displays results and explanations of the implicit egotism example
    picOutput.Cls
    picOutput.Print "Sometimes we organize information with preference to items that"
    picOutput.Print "are similar to us."
    picOutput.Print "This is known as implicit egotism."
    If FoundName = True Then
        If FoundNum = True Then
            picOutput.Print "You showed signs of implicit egotism since your favorite number " & FaveNum
            picOutput.Print "and letter of " & FaveLetter & " are in your name and birthday"
        Else
            picOutput.Print "You display a slight tendency towards implicit egotism"
        End If
    Else
        If FoundNum = True Then
            picOutput.Print "You display a slight tendency towards implicit egotism"
        Else
            picOutput.Print "You do not display any signs of implicit egotism since your favorite"
            picOutput.Print "letter and number are not in your name or birthday."
        End If
    End If
End Sub
Private Sub cmdReturn_Click()
'returns user to main menu
    frmSocial.Hide
    frmBegin.Show
End Sub
Private Sub cmdReview_Click()
'this command takes information from a data file and displays the review information on the self
    
    'declare variables
    Dim InfoArray(1 To 30) As String
    Dim Ctr As Integer, Size As Integer
    Ctr = 0
    
    'open file and input data into an array
    Open App.Path & "\SocialInfo.txt" For Input As #2
    Do Until EOF(2)
        Ctr = Ctr + 1
        Input #2, InfoArray(Ctr)
    Loop
    Close #2
    Size = Ctr
    
    'displays review information from the array
    picOutput.Cls
    For Ctr = 1 To Size
        picOutput.Print InfoArray(Ctr)
    Next Ctr
End Sub
Private Sub cmdUnrealisticOpt_Click()
'this command discusses information on unrealistic optimism and self-serving bias
    
    'declare variables
    Dim PersonalAge As Integer
    Dim Difference As Integer
    
    'find specific age that user thinks that he or she will die at
    picOutput.Cls
    MsgBox "Self-serving bias is our tendency to perceive ourselves favorably", , "Self-Serving Bias"
    PersonalAge = InputBox("Enter age that you think that you will die at", "Personal")
    Difference = PersonalAge - 75
    picOutput.Print "Most individuals predict themselves as living beyond the average age of 75."
    picOutput.Print "This is the idea of unrealistic optimism, because most people think"
    picOutput.Print "that something bad will never happen to them."
    picOutput.Print "*************************************************************************"
    
    'this case determines whether or not the user guessed right or has unrealistic optimism
    Select Case Difference
        Case 0
            picOutput.Print "You were the same as the average age of death for most Americans."
        Case 1 To 10
            picOutput.Print "You predicted yourself as living longer than the average American."
        Case Else
            picOutput.Print "You predicted yourself as living shorter than the average American."
    End Select
End Sub


