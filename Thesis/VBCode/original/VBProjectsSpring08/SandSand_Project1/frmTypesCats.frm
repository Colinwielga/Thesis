VERSION 5.00
Begin VB.Form frmTypesCats 
   BackColor       =   &H0080C0FF&
   Caption         =   "Types of Cats"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdSelect 
      BackColor       =   &H00FF8080&
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4080
      Width           =   3015
   End
   Begin VB.TextBox txtCat 
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
      Left            =   10920
      TabIndex        =   7
      Top             =   2880
      Width           =   3735
   End
   Begin VB.PictureBox picResults 
      Height          =   5655
      Left            =   4800
      ScaleHeight     =   5595
      ScaleWidth      =   5595
      TabIndex        =   5
      Top             =   1440
      Width           =   5655
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10680
      TabIndex        =   4
      Top             =   9240
      Width           =   4215
   End
   Begin VB.CommandButton cmdAlpha 
      BackColor       =   &H00FF8080&
      Caption         =   "Sort Cats Alphabeticaly"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   3255
   End
   Begin VB.CommandButton cmdPrice 
      BackColor       =   &H00FF8080&
      Caption         =   "Sort Cats by Price"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   3255
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H00FF8080&
      Caption         =   "List Cats"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label LabInstructions 
      BackColor       =   &H0080C0FF&
      Caption         =   "Click below to view the list of cats."
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   495
      Left            =   720
      TabIndex        =   9
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Type the breed of the cat that you are looking to purchase below: (Make sure to begin the name with a capital letter)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1215
      Left            =   11040
      TabIndex        =   6
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Types of Cats"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   2175
      Left            =   1080
      TabIndex        =   3
      Top             =   7440
      Width           =   5655
   End
End
Attribute VB_Name = "frmTypesCats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sand's Pet Store
'Name of Form: frmTypesCats
'Author: Scott Sand and Kate Sand
'Date Written: March 7, 2008
'Objective: The objective of this form is to list the types of cats
'           Sand's Pet Store has and for the custumer to choose a cat
'           that is right for them.
'Other Comments: This form lists the types of cats available and their
'                prices.  It also lists the dogs in alphabetical order
'                and in order by their price.  After the custumer select a
'                cat, the form moves to litterboxes for the cat.
Option Explicit

Private Sub cmdAlpha_Click()
'The customer can sort the list of cats offered into alphabetical order
Dim Pass As Integer, Pos As Integer
Dim tempcat As String, tempcatprice As Integer
Dim I As Integer
PicResults.Cls
For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If Cats(Pos) > Cats(Pos + 1) Then
            tempcat = Cats(Pos)
            Cats(Pos) = Cats(Pos + 1)
            Cats(Pos + 1) = tempcat
            tempcatprice = CatCost(Pos)
            CatCost(Pos) = CatCost(Pos + 1)
            CatCost(Pos + 1) = tempcatprice
        End If
    Next Pos
Next Pass

For I = 1 To CTR
    PicResults.Print Cats(I); Tab(40); FormatCurrency(CatCost(I))
Next I
End Sub

Private Sub cmdBack_Click()
'The customer is directed back to the main menu
frmMainMenu.Show
frmTypesCats.Hide
End Sub

Private Sub cmdPrice_Click()
'The list of cats offered is sorted into descending order by cost
Dim Pass As Integer, Pos As Integer
Dim tempcat As String, tempcatprice As Integer
Dim I As Integer
PicResults.Cls
For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If CatCost(Pos) < CatCost(Pos + 1) Then
            tempcat = Cats(Pos)
            Cats(Pos) = Cats(Pos + 1)
            Cats(Pos + 1) = tempcat
            tempcatprice = CatCost(Pos)
            CatCost(Pos) = CatCost(Pos + 1)
            CatCost(Pos + 1) = tempcatprice
        End If
    Next Pos
Next Pass

For I = 1 To CTR
    PicResults.Print Cats(I); Tab(40); FormatCurrency(CatCost(I))
Next I
End Sub

Private Sub cmdRead_Click()
'The list of cats is read from a document and displayed
PicResults.Cls
CTR = 0
Open App.Path & "\Cats.txt" For Input As #2
Do Until EOF(2)
    CTR = CTR + 1
    Input #2, Cats(CTR), CatCost(CTR)
    PicResults.Print Cats(CTR); Tab(40); FormatCurrency(CatCost(CTR))
Loop
Close #2
cmdAlpha.Enabled = True
cmdPrice.Enabled = True
cmdRead.Enabled = False
End Sub

Private Sub cmdSelect_Click()
'The customer enters a cat name and then the program will search for the cat that was input by the user
'If the cat is found it will be placed in the shopping cart
Dim Found As Boolean
Dim I As Integer, inputcat As String
inputcat = txtCat.Text
I = 0
Found = False
Do While ((Not Found) And (I < CTR))
    I = I + 1
    If inputcat = Cats(I) Then Found = True
Loop
If (Not Found) Then
        MsgBox (inputcat & " was not found please check your spelling or select a different cat.")
    Else
        MsgBox ("You have selected a " & Cats(I) & " and the cost is " & FormatCurrency(CatCost(I)) & ".")
        PetCost = PetCost + CatCost(I)
        frmLitterBox.Show
        frmTypesCats.Hide
End If

End Sub
