VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Fish Journal"
   ClientHeight    =   6435
   ClientLeft      =   4485
   ClientTop       =   3630
   ClientWidth     =   9150
   LinkTopic       =   "Form2"
   ScaleHeight     =   6435
   ScaleWidth      =   9150
   Visible         =   0   'False
   Begin VB.CommandButton cmderase 
      Caption         =   "Erase all entries"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7920
      TabIndex        =   5
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdj_print 
      Caption         =   "Print Journal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7920
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdleave 
      BackColor       =   &H008080FF&
      Caption         =   "Leave Minnesota"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "Update Journal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7920
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.PictureBox picoutput 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   120
      ScaleHeight     =   5355
      ScaleWidth      =   7515
      TabIndex        =   1
      Top             =   480
      Width           =   7575
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H0080FF80&
      Caption         =   "Back To Main Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4800
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Minnesota Fisher
'Fish Journal
'Eric Glorvigen
'Date= March 5
'this page asks the user for input, which is store into an array
'and then prints the array

Dim ctr As Integer
Dim location(1 To 100) As String
Dim fish(1 To 100) As String
Dim time(1 To 100) As String

Private Sub cmderase_Click()
    'this page resets the counter and clears the picture box
        picoutput.Cls
        ctr = 0
End Sub

Private Sub cmdexit_Click()
    'exit back to home page
        form1.Show
        Form2.Hide
End Sub

Private Sub cmdj_print_Click()
    'this takes the input from the user and prints into picture box
    
        Dim n As Integer
        
            picoutput.Print "The Journal of "; inputname
        
            picoutput.Print "Location:", Tab(25); "Type of Fish:", Tab(45); "Date:"
            picoutput.Print "**************************************************"
    
        For n = 1 To ctr
                    picoutput.Print location(n), Tab(25); fish(n), Tab(45); time(n)
        Next n
    
End Sub

Private Sub cmdleave_Click()
    'exit program
        End
End Sub

Private Sub cmdupdate_Click()
'asks user for input and stores into three parallel arrays

    Dim n As Integer
    Dim myfish As String
    Dim spot As String
    Dim mytime As String
    
       
        Do While LCase(spot) <> "leave"
            spot = InputBox("Enter Location or type Leave to quit", "Location")
                If LCase(spot) <> "leave" Then
                    myfish = InputBox("Enter type of fish:", "Fish")
                    mytime = InputBox("Enter date:", "Date")
                    ctr = ctr + 1
                    location(ctr) = spot
                    fish(ctr) = myfish
                    time(ctr) = mytime
                End If
        Loop
End Sub



