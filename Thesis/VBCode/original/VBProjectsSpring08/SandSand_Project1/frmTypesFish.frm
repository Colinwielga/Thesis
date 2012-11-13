VERSION 5.00
Begin VB.Form frmTypesFish 
   BackColor       =   &H00FF0000&
   Caption         =   "Tyes of Fish"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdSelect 
      BackColor       =   &H0000C000&
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4560
      Width           =   2535
   End
   Begin VB.TextBox txtFish 
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
      Left            =   11280
      TabIndex        =   7
      Top             =   3120
      Width           =   3495
   End
   Begin VB.PictureBox picResults 
      Height          =   5775
      Left            =   4800
      ScaleHeight     =   5715
      ScaleWidth      =   5835
      TabIndex        =   5
      Top             =   1560
      Width           =   5895
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9960
      TabIndex        =   4
      Top             =   8880
      Width           =   4815
   End
   Begin VB.CommandButton cmdAlpha 
      BackColor       =   &H0000C000&
      Caption         =   "Sort Fish Alphabeticaly"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   3615
   End
   Begin VB.CommandButton cmdPrice 
      BackColor       =   &H0000C000&
      Caption         =   "Sort Fish by Price"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   3615
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H0000C000&
      Caption         =   "List Fish"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Label LabInstructions 
      BackColor       =   &H00FF0000&
      Caption         =   "Click below to view the list of fish."
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   975
      Left            =   720
      TabIndex        =   9
      Top             =   360
      Width           =   5415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      Caption         =   "Type the name of the species of fish that you would like to purchase below: ( Make sure you begin the name with a capital letter)"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1815
      Left            =   11160
      TabIndex        =   6
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Types of Fish"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   48
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1455
      Left            =   720
      TabIndex        =   3
      Top             =   8160
      Width           =   6855
   End
End
Attribute VB_Name = "frmTypesFish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sand's Pet Store
'Name of Form: frmTypesFish
'Author: Scott Sand and Kate Sand
'Date Written: March 7, 2008
'Objective: The objective of this form is to list the types of fish
'           Sand's Pet Store has and for the custumer to choose a fish
'           that is right for them.
'Other Comments: This form lists the types of fish available and their
'                prices.  It also lists the fish in alphabetical order
'                and in order by their prices.  After the custumer select a
'                fish, the form moves to tanks for the fish.
Option Explicit

Private Sub cmdAlpha_Click()
'The list of fish is sorted into alphabetical order
Dim Pass As Integer, Pos As Integer
Dim tempfish As String, tempfishprice As Integer
Dim I As Integer
PicResults.Cls
For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If Fish(Pos) > Fish(Pos + 1) Then
            tempfish = Fish(Pos)
            Fish(Pos) = Fish(Pos + 1)
            Fish(Pos + 1) = tempfish
            tempfishprice = FishCost(Pos)
            FishCost(Pos) = FishCost(Pos + 1)
            FishCost(Pos + 1) = tempfishprice
        End If
    Next Pos
Next Pass

For I = 1 To CTR
    PicResults.Print Fish(I); Tab(40); FormatCurrency(FishCost(I))
Next I
End Sub

Private Sub cmdBack_Click()
'The customer is directed back to the main menu
frmMainMenu.Show
frmTypesFish.Hide
End Sub

Private Sub cmdPrice_Click()
'The list of fish is sorted into descending order by price
Dim Pass As Integer, Pos As Integer
Dim tempfish As String, tempfishprice As Integer
Dim I As Integer
PicResults.Cls
For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If FishCost(Pos) < FishCost(Pos + 1) Then
            tempfish = Fish(Pos)
            Fish(Pos) = Fish(Pos + 1)
            Fish(Pos + 1) = tempfish
            tempfishprice = FishCost(Pos)
            FishCost(Pos) = FishCost(Pos + 1)
            FishCost(Pos + 1) = tempfishprice
        End If
    Next Pos
Next Pass

For I = 1 To CTR
    PicResults.Print Fish(I); Tab(40); FormatCurrency(FishCost(I))
Next I
End Sub

Private Sub cmdRead_Click()
'A list of fish is loaded from a document and displayed
PicResults.Cls
CTR = 0
Open App.Path & "\Fish.txt" For Input As #3
Do Until EOF(3)
    CTR = CTR + 1
    Input #3, Fish(CTR), FishCost(CTR)
    PicResults.Print Fish(CTR); Tab(40); FormatCurrency(FishCost(CTR))
Loop
Close #3
cmdAlpha.Enabled = True
cmdPrice.Enabled = True
cmdRead.Enabled = False
End Sub

Private Sub cmdSelect_Click()
'The customer types the name of the fish they wish to purchase in the text box
'the program searches the list of fish for the input fish and if it is found
'it will be added to the shopping cart
Dim Found As Boolean
Dim I As Integer, inputfish As String
inputfish = txtFish.Text
I = 0
Found = False
Do While ((Not Found) And (I < CTR))
    I = I + 1
    If inputfish = Fish(I) Then Found = True
Loop
If (Not Found) Then
        MsgBox (inputfish & " was not found please check your spelling or select a different fish.")
        
    Else
        MsgBox ("You have selected " & Fish(I) & " that costs " & FormatCurrency(FishCost(I)))
        PetCost = PetCost + FishCost(I)
        frmFishTanks.Show
        frmTypesFish.Hide
End If
End Sub
