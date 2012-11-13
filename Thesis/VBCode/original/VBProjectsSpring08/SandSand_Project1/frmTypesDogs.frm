VERSION 5.00
Begin VB.Form frmTypesDogs 
   BackColor       =   &H00000080&
   Caption         =   "Types of Dogs"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdSelect 
      BackColor       =   &H0000FFFF&
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4080
      Width           =   3015
   End
   Begin VB.TextBox txtDog 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11160
      TabIndex        =   6
      Top             =   2760
      Width           =   3615
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10080
      TabIndex        =   5
      Top             =   8760
      Width           =   4815
   End
   Begin VB.CommandButton cmdAlpha 
      BackColor       =   &H0000FFFF&
      Caption         =   "Sort Dogs Alphabeticaly"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5880
      Width           =   3255
   End
   Begin VB.PictureBox picResults 
      Height          =   6255
      Left            =   5160
      ScaleHeight     =   6195
      ScaleWidth      =   5595
      TabIndex        =   3
      Top             =   1680
      Width           =   5655
   End
   Begin VB.CommandButton cmdPrice 
      BackColor       =   &H0000FFFF&
      Caption         =   "Sort Dogs by Price"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      Width           =   3255
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H0000FFFF&
      Caption         =   " List Dogs"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label LabInstructions 
      BackColor       =   &H00000080&
      Caption         =   "Click below to view the list of dogs."
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   9
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000080&
      Caption         =   "Type the name of the breed of dog that you want to purchase below:     ( Make sure to begin the name with a capital letter)"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1455
      Left            =   11160
      TabIndex        =   7
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      Caption         =   "Types of Dogs"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   48
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1215
      Left            =   600
      TabIndex        =   0
      Top             =   8400
      Width           =   5895
   End
End
Attribute VB_Name = "frmTypesDogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sand's Pet Store
'Name of Form: frmTypesDogs
'Author: Scott Sand and Kate Sand
'Date Written: March 7, 2008
'Objective: The objective of this form is to list the types of dogs
'           Sand's Pet Store has and for the custumer to choose a dog
'           that is right for them.
'Other Comments: This form lists the types of dogs available and their
'                prices.  It also lists the dogs in alphabetical order
'                and in order by their prices.  After the custumer select a
'                dog, the form moves to kennels for the dog.
Option Explicit

Private Sub cmdAlpha_Click()
'The list of dogs offered is sorted into alphabetical order
Dim Pass As Integer, Pos As Integer
Dim tempdog As String, tempdogprice As Integer
Dim I As Integer
PicResults.Cls
For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If Dogs(Pos) > Dogs(Pos + 1) Then
            tempdog = Dogs(Pos)
            Dogs(Pos) = Dogs(Pos + 1)
            Dogs(Pos + 1) = tempdog
            tempdogprice = DogCost(Pos)
            DogCost(Pos) = DogCost(Pos + 1)
            DogCost(Pos + 1) = tempdogprice
        End If
    Next Pos
Next Pass

For I = 1 To CTR
    PicResults.Print Dogs(I); Tab(40); FormatCurrency(DogCost(I))
Next I
End Sub

Private Sub cmdBack_Click()
'The user is directed back to the main menu
frmMainMenu.Show
frmTypesDogs.Hide
End Sub

Private Sub cmdPrice_Click()
'The list of dogs offered is sorted in descending order by price
Dim Pass As Integer, Pos As Integer
Dim tempdog As String, tempdogprice As Integer
Dim I As Integer
PicResults.Cls
For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If DogCost(Pos) < DogCost(Pos + 1) Then
            tempdog = Dogs(Pos)
            Dogs(Pos) = Dogs(Pos + 1)
            Dogs(Pos + 1) = tempdog
            tempdogprice = DogCost(Pos)
            DogCost(Pos) = DogCost(Pos + 1)
            DogCost(Pos + 1) = tempdogprice
        End If
    Next Pos
Next Pass

For I = 1 To CTR
    PicResults.Print Dogs(I); Tab(40); FormatCurrency(DogCost(I))
Next I
End Sub

Private Sub cmdRead_Click()
'A list of dogs is loaded from a document and displayed
PicResults.Cls
CTR = 0
Open App.Path & "\Dogs.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Dogs(CTR), DogCost(CTR)
    PicResults.Print Dogs(CTR); Tab(40); FormatCurrency(DogCost(CTR))
Loop
Close #1
cmdAlpha.Enabled = True
cmdPrice.Enabled = True
cmdRead.Enabled = False
End Sub

Private Sub cmdSelect_Click()
'The customer types the name of the dog they want in the text box and the program
'searches for the dog and if it is found it is added to the shopping cart
Dim Found As Boolean
Dim I As Integer, inputdog As String
inputdog = txtDog.Text
I = 0
Found = False
Do While ((Not Found) And (I < CTR))
    I = I + 1
    If inputdog = Dogs(I) Then Found = True
Loop
If (Not Found) Then
        MsgBox (inputdog & " was not found please check your spelling or select a different dog.")
    Else
        MsgBox ("You have selected a " & Dogs(I) & " and the cost is " & FormatCurrency(DogCost(I)) & ".")
        PetCost = PetCost + DogCost(I)
        frmDogKennels.Show
        frmTypesDogs.Hide
End If
End Sub

