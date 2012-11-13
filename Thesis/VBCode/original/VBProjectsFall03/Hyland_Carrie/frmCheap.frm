VERSION 5.00
Begin VB.Form frmCheap 
   BackColor       =   &H00FF0000&
   Caption         =   "Cheapest Burrito"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF80FF&
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8040
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   1815
   End
   Begin VB.PictureBox picCheapest 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   4815
      Left            =   3720
      ScaleHeight     =   4755
      ScaleWidth      =   5955
      TabIndex        =   2
      Top             =   120
      Width           =   6015
   End
   Begin VB.CommandButton cmdCheapBurrito 
      BackColor       =   &H00FF80FF&
      Caption         =   "Low on Cash?  See the items in order from least to most expensive."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FF80FF&
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5880
      Width           =   2895
   End
   Begin VB.Label lblName 
      BackColor       =   &H00FF0000&
      Caption         =   "Designed by Carrie Hyland"
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      Top             =   6600
      Width           =   2535
   End
End
Attribute VB_Name = "frmCheap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: ProjChipotleOrder (Carrie Hyland's VB Project.vbp)
'Form Name: frmCheap (Cheap_form.frm)
'Author: Carrie Hyland
'Date Written: October 19, 2003
'Purpose of Form: To have a user click a command button to
                 ' arrange the products (from a file) into
                 ' order from least to most expensive.
                 
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
'Dimensioning variables that are global
Dim Types(1 To 15) As String, Prices(1 To 15) As Single
Private Sub cmdCheapBurrito_Click()
'Dimensioning the variables
Dim Pass As Integer, Comp As Integer, Temp As String, J As Integer, CTR As Single
'Initializing CTR to Zero
CTR = 0
'Clears anything in the picCheapest screen to allow for repeated use
picCheapest.Cls
'Prints a title/heading for the picCheapest screen
picCheapest.Print "Name of Product", , "Price"
picCheapest.Print "-------------------------------------------------------------------------------"
'Opens the file for Prices to be used in an array
Open Path & "Prices.txt" For Input As #1
'Starts the loop, putting the information into arrays
Do While Not EOF(1)
  CTR = CTR + 1
  Input #1, Types(CTR), Prices(CTR)
Loop
'Closes file
Close
'Starts a comparison for putting the information in order from least
'to most expensive
For Pass = 1 To 14
    For Comp = 1 To 15 - Pass
    'Sorts the information by price from least to most expensive
       If Prices(Comp) > Prices(Comp + 1) Then
         Temp = Prices(Comp)
         Prices(Comp) = Prices(Comp + 1)
         Prices(Comp + 1) = Temp
         Temp = Types(Comp)
         Types(Comp) = Types(Comp + 1)
         Types(Comp + 1) = Temp
        End If
    Next Comp
Next Pass
'Prints the results of the sorting
For J = 1 To 15
  picCheapest.Print Types(J); Tab(43); FormatCurrency(Prices(J))
Next J
  
    
End Sub

Private Sub cmdQuit_Click()
'Ends program
End
End Sub

Private Sub cmdReturn_Click()
'Hiding the frmCheap and showing the frmMainMenu (switches
' between the two forms)
frmCheap.Hide
frmMainMenu.Show
End Sub

