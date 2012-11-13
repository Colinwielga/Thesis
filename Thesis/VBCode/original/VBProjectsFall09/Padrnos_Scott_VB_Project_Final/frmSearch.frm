VERSION 5.00
Begin VB.Form frmSearch 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   11400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18135
   LinkTopic       =   "Form1"
   ScaleHeight     =   11400
   ScaleWidth      =   18135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Return to home page"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2160
      TabIndex        =   10
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   9
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton cmdSearchWeight 
      Caption         =   "Click to See the Wrestlers at That Weight"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4080
      TabIndex        =   3
      Top             =   6120
      Width           =   1815
   End
   Begin VB.TextBox txtWeight 
      Height          =   1215
      Left            =   4440
      TabIndex        =   2
      Top             =   4320
      Width           =   1455
   End
   Begin VB.PictureBox picResults 
      Height          =   6615
      Left            =   6000
      ScaleHeight     =   6555
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   600
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "141     184"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "149     197"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "157    285 (HWT)"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "125     165"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "133     174"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblWeight 
      Caption         =   "Enter the NUMBER Weight Class you Would Like to View:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   1
      Top             =   4440
      Width           =   4335
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdback_Click()
    frmSearch.Hide 'Hides the search form
    frmHome.Show 'show's home form
End Sub

'Quits the program
Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSearchWeight_Click()
    Dim Weights As String, CTR As Integer, Wrestlers As String
    picResults.Cls 'clears the picture box
    Weights = txtWeight.Text 'stores the input from the text box into the "Weights" variable
    Select Case Weights
    Case Is = 125
        Wrestlers = "Chad Henle, Scott Padrnos, Trent Seck" 'stores the wrestlers at 125 into the "Wrestlers" variable
    Case Is = 133
        Wrestlers = "Mogi Baatar, Alex Boosalis, Matt Laine, Manny Livingstone" 'stores the wrestlers at 133 into the "Wrestlers" variable
    Case Is = 141
        Wrestlers = "Derik Gertken, Sam Morse, Minga Batsukh, Charlie Kirscht, Cody Goldschmidt" 'stores the wrestlers at 141 into the "Wrestlers" variable
    Case Is = 149
        Wrestlers = "Minga Batsukh, Kyle Glynn, Cody Goldschmidt, Drew Larson, John Paul Vaith" 'stores the wrestlers at 149 into the "Wrestlers" variable
    Case Is = 157
        Wrestlers = "Matt Baarson, Drew Larson, John Paul Vaith, Chris Stevermer, Kyle Glynn" 'stores the wrestlers at 157 into the "Wrestlers" variable
    Case Is = 165
        Wrestlers = "Matt Baarson, Grant Lydon" 'stores the wrestlers at 165 into the "Wrestlers" variable
    Case Is = 174
        Wrestlers = "Mitch Hagen, Matt Pfarr, Dustin Raygor" 'stores the wrestlers at 174 into the "Wrestlers" variable
    Case Is = 184
        Wrestlers = "Dustin Baxter, Dustin Raygor" 'stores the wrestlers at 184 into the "Wrestlers" variable
    Case Is = 197
        Wrestlers = "Tony Willaert, James Carlson" 'stores the wrestlers at 197 into the "Wrestlers" variable
    Case Is = 285
        Wrestlers = "Cody Socher, Jake Evenson" 'stores the wrestlers at HWT into the "Wrestlers" variable
    Case Else
    MsgBox ("Invalid Entry. Please try again by typing in the NUMBER weight class displayed to the left.") 'Let's the user know that they made an invalid entry
    End Select
    picResults.Print "The Wrestlers at "; Weights; " are:"
    picResults.Print Wrestlers
    'displays the wrestlers at the requested weight

End Sub


