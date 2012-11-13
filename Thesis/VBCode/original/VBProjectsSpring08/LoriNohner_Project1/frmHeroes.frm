VERSION 5.00
Begin VB.Form frmGame 
   BackColor       =   &H0000C000&
   Caption         =   "Match Game"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   BeginProperty Font 
      Name            =   "Palatino Linotype"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAnswers 
      BackColor       =   &H0080FFFF&
      Caption         =   "Am I Right?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   3015
   End
   Begin VB.ListBox lstPrince 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      ItemData        =   "frmHeroes.frx":0000
      Left            =   7320
      List            =   "frmHeroes.frx":001F
      TabIndex        =   6
      Top             =   2160
      Width           =   2295
   End
   Begin VB.ListBox lstPrincess 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      ItemData        =   "frmHeroes.frx":008E
      Left            =   480
      List            =   "frmHeroes.frx":00B0
      TabIndex        =   2
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000C0&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H000000FF&
      Caption         =   "Retun to Disney Castle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label lblRules 
      BackColor       =   &H0080FFFF&
      Caption         =   "Click on a princess from the box on the left and match her with her prince from the box on the right"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2280
      TabIndex        =   9
      Top             =   1080
      Width           =   5655
   End
   Begin VB.Label lblGame 
      BackColor       =   &H0080FFFF&
      Caption         =   "Match Game"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   8
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label lblPrince 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblequals 
      BackColor       =   &H000000FF&
      Caption         =   "    ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   4
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label lblPrincess 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Disney Project
'Game
'Lori Nohner
'Written MArch 17,2008
'Objective- allows user to match a Disney princess with a Disney prince and to see if they're right.

Private Sub cmdAnswers_Click()
    Dim Found As Boolean, Princess As String 'Declares variable found as a Boolean and Princess as a string
    
    Princess = lblPrincess.Caption 'declares the string in the text box as "Princess"
    
    Select Case Princess 'matches each princess with their prince.  If the answer is correct then found is set to true
        Case "Jasmine"
            If lblPrince.Caption = "Aladdin" Then
                Found = True
            End If
        Case "Belle"
            If lblPrince.Caption = "Beast" Then
                Found = True
            End If
        Case "Cinderella", "Snow White"
            If lblPrince.Caption = "Prince Charming" Then
                Found = True
            End If
        Case "Ariel"
            If lblPrince.Caption = "Prince Eric" Then
                Found = True
            End If
        Case "Meg"
            If lblPrince.Caption = "Hercules" Then
                Found = True
            End If
        Case "Maid Marian"
            If lblPrince.Caption = "Robin Hood" Then
                Found = True
            End If
        Case "Pocahontas"
            If lblPrince.Caption = "John Smith" Then
                Found = True
            End If
        Case "Aurora"
            If lblPrince.Caption = "Prince Phillip" Then
                Found = True
            End If
        Case "Wendy"
            If lblPrince.Caption = "Peter Pan" Then
                Found = True
            End If
        
    End Select
    
    
        If Found Then 'displays a message to the user telling them if they are right or not
            MsgBox "You're right!!", , "Correct!"
        Else
            MsgBox "Sorry, try again.", , "Incorrect"
        End If
        lblPrincess.Caption = "" 'clears the text box
        lblPrince.Caption = "" 'clears the text box
            
End Sub

Private Sub cmdExit_Click()
    End 'quits program
End Sub

Private Sub cmdReturn_Click()
    frmGame.Hide 'hides heroes page
    frmDisneyCastle.Show 'returns to Disney home page
End Sub


Private Sub lstPrince_Click()
    lblPrince.Caption = lstPrince.Text 'loads the name of the prince that the user clicks on to the text box
End Sub

Private Sub lstPrincess_Click()
    lblPrincess.Caption = lstPrincess.Text ' loads the name of the princess that the user clicks to the text box
End Sub
