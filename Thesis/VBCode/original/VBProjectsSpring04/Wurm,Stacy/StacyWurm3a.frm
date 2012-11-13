VERSION 5.00
Begin VB.Form DateEventMovie 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Movie Options"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optJumbo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0FF&
      Caption         =   "Jumbo Candy"
      Height          =   255
      Left            =   1680
      TabIndex        =   14
      Top             =   2760
      Width           =   1575
   End
   Begin VB.OptionButton optLarge 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0FF&
      Caption         =   "Large Candy"
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   2400
      Width           =   1575
   End
   Begin VB.OptionButton optMedium 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0FF&
      Caption         =   "Medium Candy"
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   2040
      Width           =   1575
   End
   Begin VB.OptionButton optSmall 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0FF&
      Caption         =   "Small Candy"
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   1680
      Width           =   1575
   End
   Begin VB.OptionButton optSSoda 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0FF&
      Caption         =   "Small Soda"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   1455
   End
   Begin VB.OptionButton optLSoda 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0FF&
      Caption         =   "Large Soda"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "I have made my choices"
      Height          =   615
      Left            =   3720
      TabIndex        =   8
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add to my Costs"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1440
      TabIndex        =   7
      Top             =   4680
      Width           =   1095
   End
   Begin VB.OptionButton optSPopcorn 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0FF&
      Caption         =   "Small Popcorn"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   1455
   End
   Begin VB.PictureBox picResults 
      Height          =   4095
      Left            =   4320
      ScaleHeight     =   4035
      ScaleWidth      =   4035
      TabIndex        =   5
      Top             =   360
      Width           =   4095
   End
   Begin VB.OptionButton optLPopcorn 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0FF&
      Caption         =   "Large Popcorn"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search for Candy"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdPriceSort 
      Caption         =   "Sort items by Price"
      Height          =   855
      Left            =   2160
      TabIndex        =   2
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   7920
      TabIndex        =   0
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Options 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   $"StacyWurm3a.frx":0000
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "DateEventMovie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Project Name: Date Chooser (Wurm, Stacy - VB Project)
' Form Name: DateEvent (StacyWurm3.frm)
' Author: Stacy Wurm
' Date Written: Sunday, March 7th, 2004
' Purpose of this Form: ' This allows the user to purchase concessions
                        ' When they go to the movie
                        ' It will also sort the items by price or
                        ' search through the items for something specific

Private Sub cmdAdd_Click()
' Prints how much you have spent on concessions
picResults.Cls
picResults.Print FormatCurrency(MovieCost); " has been added to your cost for snacks."
End Sub

Private Sub cmdDone_Click()
' Goes back to the main event page
DateEventMovie.Hide
TotalCost = TotalCost + MovieCost
End Sub

Private Sub cmdPriceSort_Click()
picResults.Cls
' bubble sort the items by price to see what is most expensive
For PASS = 1 To CTR - 1
    For COMP = 1 To CTR - PASS
        If Price(COMP) < Price(COMP + 1) Then
            
            'switch prices
            tempPrice = Price(COMP)
            Price(COMP) = Price(COMP + 1)
            Price(COMP + 1) = tempPrice
            
            'switch items
            tempItem = Item(COMP)
            Item(COMP) = Item(COMP + 1)
            Item(COMP + 1) = tempItem
            
        End If
    Next COMP
Next PASS
picResults.Print "***Highest to Lowest Price***"
picResults.Print "Item"; Tab(30); "Price"
picResults.Print "________________________________"
For J = 1 To CTR
    picResults.Print Item(J); Tab(30); FormatCurrency(Price(J))
Next J
End Sub

Private Sub cmdQuit_Click()
End
End Sub
Private Sub cmdSearch_Click()
' Searches for specific items on the list
picResults.Cls
Dim Candy As String
Dim Found As Boolean
Found = False
Candy = InputBox("Please enter desired snack")
For J = 1 To CTR
    If Candy = Item(J) Then
        picResults.Print "Yes we have "; Item(J); ".  It costs "; FormatCurrency(Price(J))
        Found = True
    End If
Next J
If Not Found Then
    picResults.Print "Sorry we do not have that item."
End If
End Sub

Private Sub Form_Load()
' Put items in an array
Dim Path As String
Path = "N:\CS130\handin\Wurm, Stacy\"
Open Path + "MovieOptions.txt" For Input As #1
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Item(CTR), Price(CTR)
Loop
Close
End Sub

Private Sub optJumbo_Click()
' Adds concessions onto the cost of the movie
Dim JCandy As Single
JCandy = 2.75
cmdAdd.Enabled = True
MovieCost = MovieCost + JCandy
picResults.Print "Thank-you for having concessions!!!"
End Sub

Private Sub optLarge_Click()
' Adds concessions onto the cost of the movie
Dim LCandy As Single
LCandy = 2.25
cmdAdd.Enabled = True
MovieCost = MovieCost + LCandy
picResults.Print "Thank-you for having concessions!!!"
End Sub

Private Sub optLPopcorn_Click()
' Adds concessions onto the cost of the movie
Dim LPopcorn As Single
LPopcorn = 3.5
cmdAdd.Enabled = True
MovieCost = MovieCost + LPopcorn
picResults.Print "Thank-you for having concessions!!!"
End Sub

Private Sub optLSoda_Click()
' Adds concessions onto the cost of the movie
Dim LSoda As Single
LSoda = 3
cmdAdd.Enabled = True
MovieCost = MovieCost + LSoda
picResults.Print "Thank-you for having concessions!!!"
End Sub

Private Sub optMedium_Click()
' Adds concessions onto the cost of the movie
Dim MCandy As Single
MCandy = 2
cmdAdd.Enabled = True
MovieCost = MovieCost + MCandy
picResults.Print "Thank-you for having concessions!!!"
End Sub

Private Sub optSmall_Click()
' Adds concessions onto the cost of the movie
Dim SCandy As Single
SCandy = 1.75
cmdAdd.Enabled = True
MovieCost = MovieCost + SCandy
picResults.Print "Thank-you for having concessions!!!"
End Sub

Private Sub optSPopcorn_Click()
' Adds concessions onto the cost of the movie
Dim SPopcorn As Single
SPopcorn = 2.25
cmdAdd.Enabled = True
MovieCost = MovieCost + SPopcorn
picResults.Print "Thank-you for having concessions!!!"
End Sub

Private Sub optSSoda_Click()
' Adds concessions onto the cost of the movie
Dim SSoda As Single
SSoda = 1.75
cmdAdd.Enabled = True
MovieCost = MovieCost + SSoda
picResults.Print "Thank-you for having concessions!!!"
End Sub
