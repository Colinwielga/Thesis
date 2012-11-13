VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find the Right Climbing Shoe For You"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   6750
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdYDSSort 
      Caption         =   "Sort Results using YDS Recommedations"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4920
      TabIndex        =   11
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton cmdSortPrice 
      Caption         =   "Sort Results by Price"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1920
      TabIndex        =   10
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Click toLoad Stock"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      TabIndex        =   8
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      TabIndex        =   6
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "All Around Shoes"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   5
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton cmdSport 
      Caption         =   "Sport Climbing Shoes"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   4
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton cmdBoulder 
      Caption         =   "Bouldering Shoes"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton cmdBig 
      Caption         =   "Big Wall Shoes"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   2
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label lblDesigner 
      BackStyle       =   0  'Transparent
      Caption         =   "Designer: Mike Paisley"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   12
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label lblSecond 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Second click on how you want your results sorted for your ease."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   9
      Top             =   2160
      Width           =   5055
   End
   Begin VB.Label lblFirst 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "First Load our Current Shoe Selection"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   960
      Width           =   5655
   End
   Begin VB.Label lblThird 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Third Click on Desired Climbing Style"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   1
      Top             =   3960
      Width           =   4575
   End
   Begin VB.Label lblMain 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Thanks for choosing Split Rock Climbing for your shoe Recommendation Source."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is the main form that the User sees upon running the Program
'It allows a user to Load the Shoe list from the Store and sort this list upon their specification.
'Finally it lets the User navigate to the searching Forms for the style of Shoe they wish to Purchase.

Private Sub cmdAll_Click()      'This cmdClick subroutine Shows the All Around shoes form and Hides the Main Form
    frmAll.Show
    frmMain.Hide
End Sub

Private Sub cmdBig_Click()      'This cmdClick subroutine Shows the Big Wall shoes form and Hides the Main Form
    frmBig.Show
    frmMain.Hide
End Sub

Private Sub cmdBoulder_Click()  'This cmdClick subroutine Shows the Bouldering shoes form and Hides the Main Form
    frmBoulder.Show
    frmMain.Hide
End Sub

Private Sub cmdExit_Click()     'This cmdClick subroutine Exits the Program
    End
End Sub

Private Sub cmdLoad_Click()     'This cmdClick subroutine Loads the data of the txt file provided by the store owner into 5 parallel arrays.
    Dim pos As Integer
    pos = 0
    Open App.Path & "\ClimbingShoes.txt" For Input As #1
    Do Until EOF(1)
        pos = pos + 1
        Input #1, Names(pos), Price(pos), Rating(pos), Agr(pos), Here(pos)
    Loop
    Close #1
    Size = pos
    MsgBox "Shoe list Loaded", , "Confirm Load"
End Sub

Private Sub cmdSortPrice_Click()        'This cmdClick sub routine sorts the arrays loaded by the price array.
    Dim pos As Integer
    Dim Pass As Integer
    Dim tempPr, temprate As Single
    Dim tempNames As String
    Dim TempAgr As Integer
    For Pass = 1 To (Size - 1)
        For pos = 1 To (Size - Pass)
            If Price(pos) < Price(pos + 1) Then
                tempPr = Price(pos)
                Price(pos) = Price(pos + 1)
                Price(pos + 1) = tempPr
                tempNames = Names(pos)
                Names(pos) = Names(pos + 1)
                Names(pos + 1) = tempNames
                temprate = Rating(pos)
                Rating(pos) = Rating(pos + 1)
                Rating(pos + 1) = temprate
                TempAgr = Agr(pos)
                Agr(pos) = Agr(pos + 1)
                Agr(pos + 1) = TempAgr
            End If
        Next pos
    Next Pass
    MsgBox "Results will be sorted by Price", , "Confirm Sort"
End Sub

Private Sub cmdSport_Click()    'This cmdClick subroutine Shows the Sport shoes form and Hides the Main Form
    frmSport.Show
    frmMain.Hide
End Sub

Private Sub cmdYDSSort_Click()  'This cmdClick sub routine sorts the arrays loaded by the YDS or Rating array.
    Dim pos As Integer
    Dim Pass As Integer
    Dim tempPr, temprate As Single
    Dim tempNames As String
    Dim TempAgr As Integer
    For Pass = 1 To (Size - 1)
        For pos = 1 To (Size - Pass)
            If Rating(pos) < Rating(pos + 1) Then
                temprate = Rating(pos)
                Rating(pos) = Rating(pos + 1)
                Rating(pos + 1) = temprate
                tempNames = Names(pos)
                Names(pos) = Names(pos + 1)
                Names(pos + 1) = tempNames
                tempPr = Price(pos)
                Price(pos) = Price(pos + 1)
                Price(pos + 1) = tempPr
                TempAgr = Agr(pos)
                Agr(pos) = Agr(pos + 1)
                Agr(pos + 1) = TempAgr
            End If
        Next pos
    Next Pass
    MsgBox "Results will be sorted by Rating", , "Confirm Sort"
End Sub
