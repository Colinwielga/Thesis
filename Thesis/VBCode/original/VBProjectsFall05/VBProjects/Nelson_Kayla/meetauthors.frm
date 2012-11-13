VERSION 5.00
Begin VB.Form frmmeetauthors 
   BackColor       =   &H0080FFFF&
   Caption         =   "Meet Authors; Project by Kayla Nelson"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10485
   ForeColor       =   &H00008080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdmeet 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Click Here to Select an Author to Meet"
      Height          =   1335
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5760
      Width           =   4455
   End
   Begin VB.PictureBox picjamesberendt 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2790
      Left            =   7800
      Picture         =   "meetauthors.frx":0000
      ScaleHeight     =   2760
      ScaleWidth      =   2475
      TabIndex        =   12
      Top             =   3720
      Width           =   2505
   End
   Begin VB.PictureBox pickhaledhosseini 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2205
      Left            =   5640
      Picture         =   "meetauthors.frx":17AE
      ScaleHeight     =   2175
      ScaleWidth      =   1650
      TabIndex        =   11
      Top             =   3720
      Width           =   1680
   End
   Begin VB.PictureBox picnicholassparks 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1530
      Left            =   2640
      Picture         =   "meetauthors.frx":28C0
      ScaleHeight     =   1500
      ScaleWidth      =   2205
      TabIndex        =   10
      Top             =   3720
      Width           =   2235
   End
   Begin VB.PictureBox picjamesfrey 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1530
      Left            =   480
      Picture         =   "meetauthors.frx":3489
      ScaleHeight     =   1500
      ScaleWidth      =   1380
      TabIndex        =   9
      Top             =   3720
      Width           =   1410
   End
   Begin VB.PictureBox piccity 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   8400
      Picture         =   "meetauthors.frx":561A
      ScaleHeight     =   2505
      ScaleWidth      =   1425
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.PictureBox picmillion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   240
      Picture         =   "meetauthors.frx":72BB
      ScaleHeight     =   2985
      ScaleWidth      =   1905
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.PictureBox picfirst 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   2880
      Picture         =   "meetauthors.frx":88A2
      ScaleHeight     =   2625
      ScaleWidth      =   1785
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.PictureBox pickite 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   5640
      Picture         =   "meetauthors.frx":A510
      ScaleHeight     =   2265
      ScaleWidth      =   1545
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdreturnmm 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6600
      Width           =   2775
   End
   Begin VB.Label lblberendt 
      BackColor       =   &H00C0FFFF&
      Caption         =   "John Berendt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   8
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lblhosseini 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Khaled Hosseini"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblsparks 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Nicholas Sparks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblfrey 
      BackColor       =   &H00C0FFFF&
      Caption         =   "James Frey"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   3240
      Width           =   1335
   End
End
Attribute VB_Name = "frmmeetauthors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Kayla 's Book Club(MainMenu.vbp)
'Form Name: Meet Authors (frmmeetauthors.vbp)
'Author: Kayla Nelson
'Date: 10-27-05
'purpose of the form: This form allows the reader to input the name of the author they would like to read about.  A Message box then appears with the information.

Dim AN(1 To 4) As String
Option Explicit
Private Sub cmdmeet_Click()
    Dim A As String
    A = InputBox("Please enter the name of the author you would like to meet", "MeetAuthor") 'Gets the name of the author from the user
    Dim NotFound As Boolean
    Dim I As Integer
    NotFound = True
    Open App.Path & "\meau.txt" For Input As #1 'Opens the input file array containing the authors name
    
    
    For I = 1 To 4
        Input #1, AN(I) 'Fills array with information from input#1
    Next I
    
    I = 0
    Do While NotFound And I <= 3 'Begins the loop to search the array
        I = I + 1 'Used as a counter for the array
        If A = AN(I) Then NotFound = False 'If the name the user inputed matches the one within the array, The file is said to be Found and it will continue on to the next section
        Close #1
     Loop

    If NotFound Then 'If the name entered was no found within the array a message box will appear with this message.
        MsgBox "You have entered an invalid name, Please check your spelling."
    ElseIf A = AN(1) Then 'If the name entered was found in the number 1 slot of the array a message box will appear with this message.
        MsgBox "Born in Cleveland, Ohio, 1969. Spent most of childhood in Ohio and Michigan. Has also lived in Boston, Wrightsville Beach NC, Sao Paulo Brazil, London, Paris, Chicago, Los Angeles. Graduated high school in 1988. Further education at Denison University and the Art Institute of Chicago. Worked variety of jobs in Chicago. Worked as a screenwriter, director and producer in Los Angeles. In 2000, took second mortgage on house, spent a year writing A Million Little Pieces. Nan A. Talese/ Doubleday publishes AMLP in May of 2003. Writes My Friend Leonard in 2004. Riverhead Books publishes MFL in June of 2005. Lives in New York with wife, daughter, and two dogs.", , "James Frey's Biography"
    ElseIf A = AN(2) Then 'If the name entered was found in the number 2 slot of the array a message box will appear with this message.
        MsgBox "Khaled Hosseini was born in Kabul, Afghanistan in 1965. He is the oldest of five children. and his mother was a teacher of Farsi and History at a large girls high school in Kabul. In 1976, Khaled’s family was relocated to Paris, France, where his father was assigned a diplomatic post in the Afghan embassy. The assignment would return the Hosseini family in 1980, but by then Afghanistan had already witnessed a bloody communist coup and the Soviet invasion. Khaled’s family, instead, asked for and was granted political asylum in the U.S. He moved to San Jose, CA, with his family in 1980. He attended Santa Clara University and graduated from UC San Diego School of Medicine. He has been in practice as an internist since 1996. He is married, has two children (a boy and a girl, Haris and Farah). The Kite Runner is his first novel.", , "Khaled Hosseini's Biography"
    ElseIf A = AN(3) Then 'If the name entered was found in the number 3 slot of the array a message box will appear with this message.
        MsgBox "John Berendt is an American author, known for writing the best-selling non-fiction book Midnight in the Garden of Good and Evil. Berendt grew up in Syracuse, New York, where both of his parents were writers. As an English major at Harvard University, he worked on the staff of the Harvard Lampoon. He graduated in 1961 and moved to New York City to pursue a journalist career. Berendt was editor of New York Magazine from 1977 to 1979 and a columnist for Esquire from 1982 to 1994. When he penned Midnight in the Garden of Good and Evil in 1994, Berendt became an overnight success. Chronicling the real-life events surrounding a murder trial in Savannah, Georgia, the book spent 216 weeks on the New York Times bestseller list. A movie version directed by Clint Eastwood appeared in 1997 to mixed acclaim. Berendt's next book, The City of Falling Angels, chronicling interwoven lives in Venice, has just been released.", , "John Berendt's Biography"
    Else 'Since the name was found, and it was not one of the first three- then the only thing left for it to be is array slot 4.  A message box will then appear with this message.
        MsgBox "Nicholas Sparks was born on December 31, 1965 in Omaha, Nebraska. He graduated from the University of Notre Dame in 1988 and is one of the more critically acclaimed authors of the past 5 years. He is the author of 5 best-selling books, including The Notebook and The Rescue. Three of his books, Message in a Bottle (1999), A Walk to Remember (2002) and The Notebook (2004), have been adapted into blockbuster movies. Sparks lives in North Carolina with his wife of 13 years; 3 sons, and twin daughters.", , "Nicholas Sparks' Biography"
    End If 'Closes the If statements
    End Sub

Private Sub cmdreturnmm_Click() 'This closes the Meet authors form and opens the Main Menu form.
    frmmeetauthors.Hide
    frmmainmenu.Show
End Sub

