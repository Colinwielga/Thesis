VERSION 5.00
Begin VB.Form frmtitle 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   9510
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   9510
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1215
      Left            =   5520
      TabIndex        =   2
      Text            =   "Welcome to Deal or No Deal!!  ,With Howie"
      Top             =   2640
      Width           =   8775
   End
   Begin VB.PictureBox Picture1 
      Height          =   3735
      Left            =   7560
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   3675
      ScaleWidth      =   4875
      TabIndex        =   1
      Top             =   3720
      Width           =   4935
   End
   Begin VB.CommandButton cmdstart 
      BackColor       =   &H0000FFFF&
      Caption         =   "Click here to play"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8400
      Width           =   3255
   End
End
Attribute VB_Name = "frmtitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdstart_Click()

Dim ctr As Integer                      'dims variables
Dim index1 As Integer
Dim index2 As Integer
Dim Temp As Long
Dim I As Integer

ctr = 0                                 'keeps track of how many entries in the array
Open App.Path & "\value.txt" For Input As #1            'opens a txt file
Do Until EOF(1)                                         'reads file until no information is left
    ctr = ctr + 1                                       'keeps counter place
    Input #1, value(ctr)                                'declares what values recieved from file represent
Loop
For I = 1 To 25                                 'Randomizes the Array by taking the value of one array and switching it with another, similar to a bubble sort but is not based on value
    index1 = Int(Rnd * 25) + 1
    index2 = Int(Rnd * 25) + 1
    While index1 = index2
    index2 = Int(Rnd * 25) + 1
    Wend
    Temp = value(index1)
    value(index1) = value(index2)
    value(index2) = Temp
Next I                                          'ends loop
    
    
        
        

player = InputBox("Enter Your Name", "Enter Name")
frmtitle.Hide                                         'Changes form
frmgamescreen.Show
End Sub
