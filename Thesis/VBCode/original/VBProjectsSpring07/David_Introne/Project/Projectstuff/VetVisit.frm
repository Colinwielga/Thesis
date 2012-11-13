VERSION 5.00
Begin VB.Form VetVisit 
   BackColor       =   &H00000000&
   Caption         =   "Vet Visit"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13755
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   13755
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Text            =   "For total pet care click on everything once!"
      Top             =   0
      Width           =   8055
   End
   Begin VB.CommandButton Bonus 
      BackColor       =   &H0000FF00&
      Caption         =   "Bonus Question"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   3135
   End
   Begin VB.PictureBox PicWeight 
      BackColor       =   &H00FFFFC0&
      FillColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6840
      ScaleHeight     =   1515
      ScaleWidth      =   2475
      TabIndex        =   4
      Top             =   5280
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Click to input Puppy weight in Pounts"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5520
      Width           =   2775
   End
   Begin VB.CommandButton Clip 
      BackColor       =   &H0080FF80&
      Caption         =   "Click To Get Clipped"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "Get Shots"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7800
      Width           =   2055
   End
   Begin VB.CommandButton Back 
      BackColor       =   &H00808080&
      Caption         =   "Back to Profile"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7440
      Width           =   4695
   End
   Begin VB.Image Image3 
      Height          =   7140
      Left            =   5880
      Picture         =   "VetVisit.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7920
   End
   Begin VB.Image Image2 
      Height          =   4155
      Left            =   0
      Picture         =   "VetVisit.frx":D69C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5970
   End
   Begin VB.Image Image1 
      Height          =   4995
      Left            =   0
      Picture         =   "VetVisit.frx":6D89A
      Top             =   4080
      Width           =   6090
   End
End
Attribute VB_Name = "VetVisit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Back_Click()
    Select Case puppick
            Case 11 'moves on to next case if previous did not work
                VetVisit.Hide
                ProShep.Show
            Case 12 'moves on to next case if previous did not work
                VetVisit.Hide
                ProPit.Show
            Case 13 'moves on to next case if previous did not work
                VetVisit.Hide
                ProMtn.Show
            Case 14
                VetVisit.Hide
                Produch.Show
    End Select
End Sub

Private Sub Bonus_Click()
Dim Bonus As Integer 'sets variables
MsgBox "Input the number that cooresonds to the dog correctly and you win a hearty 7 points! Ready?", , "How to Play"
Bonus = InputBox("What kind of dog is this to the right? Is it 1 a Husky, 2 a Lab mix, 3 a Cocker Spaniel or 4 a Malamute?", "Bonus")
If Bonus = 3 Then 'tests input value to see if player passed
    BonusScreen.Show
    Else
    MsgBox "Sorry, Try again", , "Oops"
    End If
End Sub

Private Sub Clip_Click()
Clip.Visible = False ' displays message, gives points and causes button to become invisible
MsgBox "You get 2 points for basic maintinance.", , "Result"
End Sub

Private Sub Command2_Click()
Command2.Visible = False ' displays message, gives points and causes button to become invisible
MsgBox "You recieve 5 points for important shot treatments!", , "Result"
End Sub

Private Sub Command3_Click()
Dim WeightLBS As Integer, Kilograms As Integer, Ounce As Integer, Newtons As Integer
'sets variables
'sets equation
WeightLBS = InputBox("Input weight in Lbs", "Weight")
Kilograms = (WeightLBS * 0.4535924)
Newtons = (WeightLBS * 4.448222)
Ounce = (WeightLBS * 16.69583)
PicWeight.Cls 'clears screen
    
            PicWeight.Print pupick; " weighs..."
            PicWeight.Print FormatNumber(Kilograms); " Kilograms" 'prints results
            PicWeight.Print FormatNumber(Newtons); " Newtons"
            PicWeight.Print FormatNumber(Ounce); " Ounces"
        
    
MsgBox "You get 2 poits" 'displays points gained
Score = Score + 2
End Sub
