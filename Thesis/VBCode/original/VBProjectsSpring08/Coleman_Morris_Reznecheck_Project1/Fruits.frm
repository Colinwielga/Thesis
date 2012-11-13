VERSION 5.00
Begin VB.Form frm4 
   BackColor       =   &H0080FF80&
   Caption         =   "Fruit"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   12780
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.PictureBox picoutput 
      Height          =   4335
      Left            =   2760
      ScaleHeight     =   4275
      ScaleWidth      =   3555
      TabIndex        =   2
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton cmdshow 
      Caption         =   "Show Fruit"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back To Foor Pyramid"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "It is recommended that you consume about 2 cups of fruit per day"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   1935
   End
End
Attribute VB_Name = "frm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Healthy Living
'frm4
'Ben Morris
'March 21
'displays the different fruits and searches them
Option Explicit
Dim fruit(1 To 15) As String
Dim CTR As Integer
Private Sub cmdback_Click()
    frm1.Show
    frm2.Hide
    frm3.Hide
    frm4.Hide
    frm5.Hide
    frm6.Hide
    frm7.Hide
    frm8.Hide
    'shows the pyramid and hides all others
End Sub

Private Sub cmdsearch_Click()

   'this code is a do while sort, and it finds a entered fruit friom the list and prints it
    
    picoutput.Cls
    Dim found As Boolean, X As String
    CTR = 0
    found = False
    X = InputBox("Search for a type of fruit", "Fruit")
    
    Do While (found = False And CTR < 15)
        CTR = CTR + 1
        If fruit(CTR) = X Then
            found = True
        End If
    Loop
    
    If found Then
        picoutput.Print "Your Search Revealed"
        picoutput.Print X

    Else
        picoutput.Print "Thats not in my list"
    End If


        
    
    
End Sub

Private Sub cmdshow_Click()

    'this code gets the fruit list form the file and prints it in the output
    picoutput.Cls
    CTR = 0
    picoutput.Print "These are some Examples of Fruit"
    picoutput.Print "--------------------------------------------------------------------"
    Open App.Path & "\Fruit.txt" For Input As #1
    Do Until EOF(1)
    
    ' Get the data from the file
        CTR = CTR + 1
        Input #1, fruit(CTR)
        
        picoutput.Print fruit(CTR)
    
    Loop
    Close #1


End Sub

