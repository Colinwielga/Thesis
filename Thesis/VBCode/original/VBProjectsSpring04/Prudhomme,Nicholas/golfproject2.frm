VERSION 5.00
Begin VB.Form golfclub 
   BackColor       =   &H00C00000&
   Caption         =   "Form1"
   ClientHeight    =   4665
   ClientLeft      =   3375
   ClientTop       =   195
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   8190
   Begin VB.CommandButton cmdprice 
      Caption         =   "Sort Irons by Price"
      Enabled         =   0   'False
      Height          =   855
      Left            =   5040
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdbrand 
      Caption         =   "What Kind should I buy?"
      Height          =   855
      Left            =   5040
      TabIndex        =   9
      Top             =   480
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Height          =   4455
      Left            =   6480
      Picture         =   "golf project 2.frx":0000
      ScaleHeight     =   4395
      ScaleWidth      =   1275
      TabIndex        =   4
      Top             =   120
      Width           =   1335
      Begin VB.TextBox txtface 
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Text            =   "3"
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txthosel 
         Height          =   285
         Left            =   840
         TabIndex        =   7
         Text            =   "4"
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox txtshaft 
         Height          =   285
         Left            =   360
         TabIndex        =   6
         Text            =   "2"
         Top             =   2160
         Width           =   255
      End
      Begin VB.TextBox txtgrip 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Text            =   "1"
         Top             =   3720
         Width           =   255
      End
   End
   Begin VB.PictureBox picclub 
      Height          =   2775
      Left            =   120
      ScaleHeight     =   2715
      ScaleWidth      =   4635
      TabIndex        =   3
      Top             =   240
      Width           =   4695
   End
   Begin VB.CommandButton cmdcourse 
      Caption         =   "On to the Links!"
      Height          =   855
      Left            =   4320
      TabIndex        =   2
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmdwhat 
      Caption         =   "What is it?"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "By: Nicholas Prudhomme"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   4320
      Width           =   1935
   End
End
Attribute VB_Name = "golfclub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Golf_Project (N_Prudhomme_project.vbp)
'File Name: golfclub "golf project 2.frm"
'Author: Nicholas Prudhomme
'Date Written: March 12, 2004
'Overall Purpose: To introduce concepts of the game of golf
                 'For the user to learn about parts of the golf club
                 'For the user to learn about handicapping
                 'For the user to learn about distance and location on the course
'Purpose of Form: To teach people about parts of the golf club
                 'To inform/recommend types of clubs to the user
                 'Show the user the price range and selection of club
'Option Explicit is a command force that makes the user declare all
'variables before they can be used
Option Explicit
Dim path As String
Dim brand(1 To 9) As String
Dim irons(1 To 9) As String
Dim price(1 To 9) As Integer

Private Sub cmdbrand_Click()
Dim K As Integer
'clear results in picture box
picclub.Cls
'the following are used just as stlye  to set up the print box
picclub.Print "We recommend the following types of irons:"
'tab function used to improve quality of the print box
picclub.Print "Brand"; Tab(20); "Number of Irons"; Tab(40); "Price"
picclub.Print "-------------------------------------------------------------------------------------------------"
'file has names of clubs, how many you receive and their price
Open path & "clubprice.txt" For Input As #1
'loading file into array and printing it
For K = 1 To 9
    Input #1, brand(K), irons(K), price(K)
    picclub.Print brand(K); Tab(20); irons(K); Tab(40); FormatCurrency(price(K))
Next K
Close #1
'makes it so the user must first look at the clubs and then sort them by price
cmdprice.Enabled = True
End Sub

Private Sub cmdcourse_Click()
'will hide the club section of the project and show the course section
golfcourse.Show
golfclub.Hide
End Sub

Private Sub cmdprice_Click()
'will sort the clubs by price and produce a new list in the picture box
'according to price
Dim i As Integer
Dim pass As Integer
Dim temp As Integer
Dim tempa As String
Dim tempb As String
picclub.Cls
'same file as previous has information about clubs and their prices
Open path & "clubprice.txt" For Input As #1
'reloaded because it was closed in the brand command button
For i = 1 To 9
    Input #1, brand(i), irons(i), price(i)
Next i
'the next sequence will allow the irons to be sorted by price in ascending order
For pass = 1 To 9 - 1
    For i = 1 To 9 - pass
        If price(i) > price(i + 1) Then
            temp = price(i)
            price(i) = price(i + 1)
            price(i + 1) = temp
            'so the brand and number of irons get sorted with their corresponding price
            tempa = brand(i)
            brand(i) = brand(i + 1)
            brand(i + 1) = tempa
            tempb = irons(i)
            irons(i) = irons(i + 1)
            irons(i + 1) = tempb
        End If
    Next i
Next pass
'print statements done for style before results are printed
picclub.Print "Brand"; Tab(20); "Number of Irons"; Tab(40); "Price"
picclub.Print "-------------------------------------------------------------------------------------------------"
For i = 1 To 9
    'will print the clubs in order of ascending price
    picclub.Print brand(i); Tab(20); irons(i); Tab(40); FormatCurrency(price(i))
Next i
Close #1 'close the file used in the array
End Sub

Private Sub cmdquit_Click()
    End 'will end the program
End Sub

Private Sub cmdwhat_Click()
'this button will allow the user to find out what each designated piece of the
'golf club is
Dim nam(1 To 4) As String
Dim descript(1 To 4) As String
Dim parts As Integer
Dim choice(1 To 4) As String
Dim user As Integer
'clear the print box for further use
picclub.Cls
'allows the user to choose which piece
user = InputBox("What part of the club would you like to know about?", "Parts of the Club")
'open file containing a name and description for each piece
Open path & "club.txt" For Input As #1
'load file into array
For parts = 1 To 4
    Input #1, choice(parts), nam(parts), descript(parts)
Next parts
'the message box will ensure that the user pick an available number
If user > 4 Then
    MsgBox "Sorry, you must enter a number between 1 and 4.", , "Error"
    ElseIf user < 1 Then
    MsgBox "Sorry, you must enter a number between 1 and 4.", , "Error"
End If
'print statements for the selections
If user = 1 Then
    picclub.Print nam(1); Tab(1); descript(1)
    ElseIf user = 2 Then
    picclub.Print nam(2); Tab(1); descript(2)
    ElseIf user = 3 Then
    picclub.Print nam(3); Tab(1); descript(3)
    ElseIf user = 4 Then
    picclub.Print nam(4); Tab(1); descript(4)
End If
    Close #1
End Sub

Private Sub Form_Load()
'make the file a path so that it can be accessed by other computers
 path = "N:\CS130\handin\Prudhomme, Nicholas\"
End Sub
