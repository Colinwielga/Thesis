VERSION 5.00
Begin VB.Form frmBeerReferences 
   AutoRedraw      =   -1  'True
   Caption         =   "Beer References"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   Picture         =   "frmBeerReferences.frx":0000
   ScaleHeight     =   5385
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDisplay1 
      Caption         =   "Display Beer in order from Lightest to Darkest"
      BeginProperty Font 
         Name            =   "Minion Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      TabIndex        =   4
      Top             =   0
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H000040C0&
      BeginProperty Font 
         Name            =   "Minion Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   2895
      Left            =   120
      ScaleHeight     =   2835
      ScaleWidth      =   4755
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display Beer in order from Darkest to Lightest"
      BeginProperty Font 
         Name            =   "Minion Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      Picture         =   "frmBeerReferences.frx":6AD0
      TabIndex        =   2
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Beer Rankings"
      BeginProperty Font 
         Name            =   "Minion Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Menu"
      BeginProperty Font 
         Name            =   "Minion Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "frmBeerReferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Bartending School
'frmBeerReferences(Beer References)
'By Fred Paul & Michael McKeever
'March 22,2006
'The beer references form orders beers from lightest to darkest
'and from darkest to lightest, in order to educate the user
'in different beer types.

'Declare Beer rankings from numbers 1 to 8
Dim BeerRankings(1 To 8) As Integer
'Declare Pass pos and Size variables as Integers for the loop.
Dim Pass, Pos, Size As Integer
'Declare Numbers as Single according to the number of units
'In the array.
Dim Numbers(1 To 8) As Single
'Declare Flavors as String variable according to numbers in
'txt file
Dim Flavors(1 To 8) As String
'Declare Temp as Single to Store temporary data from txt file
Dim Temp As Single
'Declare Second Tempp as String to Stor temp. String functions
'from txtfile.
Dim Tempp As String




Private Sub cmdBack_Click()
    'When clicked Beer references form disappears and the bar
    'form appears
    frmBeerReferences.Hide
    frmBar.Show
End Sub

Private Sub cmdDisplay_Click()
    'When clicked this Button Sorts the beer loaded from the
    'Beerstyles.txt file
    'It then displays the beers from darkest to lightest in order
    For Pass = 1 To (Size - 1)
        For Pos = 1 To (Size - Pass)
            If Numbers(Pos) > Numbers(Pos + 1) Then
                Temp = Numbers(Pos)
                Numbers(Pos) = Numbers(Pos + 1)
                Numbers(Pos + 1) = Temp
                
                Tempp = Flavors(Pos)
                Flavors(Pos) = Flavors(Pos + 1)
                Flavors(Pos + 1) = Tempp
            End If
        Next Pos
    Next Pass
    picResults.Visible = True
    
    picResults.Cls
    For Pos = 1 To Size
        picResults.Print Flavors(Pos)
    Next Pos
End Sub

Private Sub cmdDisplay1_Click()
'When clicked this Button Sorts the beer loaded from the
    'Beerstyles.txt file
    'It then displays the beers from lightest to darkest in order
    For Pass = 1 To (Size - 1)
        For Pos = 1 To (Size - Pass)
            If Numbers(Pos) < Numbers(Pos + 1) Then
                Temp = Numbers(Pos)
                Numbers(Pos) = Numbers(Pos + 1)
                Numbers(Pos + 1) = Temp
                
                Tempp = Flavors(Pos)
                Flavors(Pos) = Flavors(Pos + 1)
                Flavors(Pos + 1) = Tempp
            End If
        Next Pos
    Next Pass
    picResults.Visible = True
    'Dim Pos As Integer
    picResults.Cls
    For Pos = 1 To Size
        picResults.Print Flavors(Pos)
    Next Pos
End Sub

Private Sub cmdLoad_Click()
    'When clicked this button loads the fil beerstyles.txt for
    'use in other cmd functions

    Open App.Path & "\BeerStyles.txt" For Input As #1
    Pos = 0
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, Numbers(Pos), Flavors(Pos)
    Loop
    Close #1
        
    Size = Pos


End Sub


