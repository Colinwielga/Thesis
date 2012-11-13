VERSION 5.00
Begin VB.Form frmTorreTwo 
   BackColor       =   &H0000FF00&
   Caption         =   "400 I.M. 2004 Olympics Display and Sorting"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLastPage 
      BackColor       =   &H000000FF&
      Caption         =   "Previous Page"
      Height          =   735
      Left            =   480
      TabIndex        =   6
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdNextPage 
      Caption         =   "Next Page"
      Height          =   855
      Left            =   360
      TabIndex        =   5
      Top             =   3360
      Width           =   1575
   End
   Begin VB.PictureBox picOutputIMTwo 
      BackColor       =   &H00800080&
      Height          =   2175
      Left            =   2520
      ScaleHeight     =   2115
      ScaleWidth      =   5955
      TabIndex        =   3
      Top             =   2760
      Width           =   6015
   End
   Begin VB.CommandButton cmdPlace 
      Caption         =   "Show Placing"
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Display Heat"
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.PictureBox picOutputIM 
      BackColor       =   &H00800080&
      Height          =   2175
      Left            =   2520
      ScaleHeight     =   2115
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   480
      Width           =   6015
   End
   Begin VB.Label lblFourIM 
      BackColor       =   &H0000FFFF&
      Caption         =   "2004 Olympic 400 I.M. Finals"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "frmTorreTwo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The Author of this Form was Torre Ahlberg on 11/01/06
'The objective of this form is to take the results of the 400 I.M. from the olympics
'which is my favorite event and display the original results
'And then thake these results and sort them into the correct finishing places
'It also uses multiple forms

'This Subroutine uses the visible true or false to go back to the first Form
Private Sub cmdLastPage_Click()
    frmTorreTwo.Visible = False
    frmTorre.Visible = True
    
End Sub

'This Subroutine uses the visible true or false to go from this Form to the next Form
Private Sub cmdNextPage_Click()
    frmTorreTwo.Visible = False
    TorreThree.Visible = True
    
End Sub

'This Subroutine takes the 400 I.M. results from the 2004 Olympics
'and put the result into a parallel array (Place, Person, and time
'It then takes this array and sorts it so that the places, people and times line up in 1st thru 8th place
'It then takes these results and prints them in a picture box
Private Sub cmdPlace_Click()
   Dim Place(1 To 8) As Integer
    Dim Person(1 To 8) As String
    Dim Time(1 To 8) As Single
    Dim SPlace As Integer
    Dim SPerson As String
    Dim STime As Single
    Dim Counter As Integer
    Dim Pos As Integer
    Dim TempNum As Single
    Dim TempNumOne As String
    Dim Pass As Integer
    Dim Comp As Integer
    Dim Size As Integer
    Counter = 0
    TempNum = 0
    
    Open App.Path & "\400IM.txt" For Input As #1
    Do Until EOF(1)
        Input #1, SPlace, SPerson, STime
        Counter = Counter + 1
        Place(Counter) = SPlace
        Person(Counter) = SPerson
        Time(Counter) = STime
    Loop
    Close #1
    
    For Pass = 1 To Counter - 1
        For Comp = 1 To Counter - Pass
            If Place(Comp) > Place(Comp + 1) Then
                TempNum = Place(Comp)
                Place(Comp) = Place(Comp + 1)
                Place(Comp + 1) = TempNum
            End If
        Next Comp
    Next Pass
    
    For Pass = 1 To Counter - 1
        For Comp = 1 To Counter - Pass
            If Person(Comp) > Person(Comp + 1) Then
                TempNumOne = Person(Comp)
                Person(Comp) = Person(Comp + 1)
                Person(Comp + 1) = TempNumOne
            End If
        Next Comp
    Next Pass
    
    For Pass = 1 To Counter - 1
        For Comp = 1 To Counter - Pass
            If Time(Comp) > Time(Comp + 1) Then
                TempNum = Time(Comp)
                Time(Comp) = Time(Comp + 1)
                Time(Comp + 1) = TempNum
            End If
        Next Comp
    Next Pass
    
    picOutputIMTwo.Print "Place", "Time", "Person"
    For Pos = 1 To 8
    picOutputIMTwo.Print Place(Pos), Time(Pos), Person(Pos)
    Next Pos
                
        
End Sub

'This Subroutine takes the results of the 2004 olymics 400 I.M. finals and puts it into a parallel array
'It then prints this array in a picture box
Private Sub cmdSort_Click()
    Dim Place(1 To 8) As Integer
    Dim Person(1 To 8) As String
    Dim Time(1 To 8) As Single
    Dim SPlace As Integer
    Dim SPerson As String
    Dim STime As Single
    Dim Counter As Integer
    Dim Pos As Integer
    Counter = 0
    
    Open App.Path & "\400IM.txt" For Input As #1
    Do Until EOF(1)
        Input #1, SPlace, SPerson, STime
        Counter = Counter + 1
        Place(Counter) = SPlace
        Person(Counter) = SPerson
        Time(Counter) = STime
    Loop
    Close #1
    
    picOutputIM.Print "Place", "Name", , "Time"
    For Pos = 1 To 8
        picOutputIM.Print Place(Pos), Person(Pos), , FormatNumber(Time(Pos))
    Next Pos
    
End Sub
