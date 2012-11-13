VERSION 5.00
Begin VB.Form frmStores 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Store Nearest You"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9105
   FillColor       =   &H00FFFFC0&
   ForeColor       =   &H00FFFFC0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00FF0000&
      Caption         =   "Try another city!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF0000&
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
      Height          =   975
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton cmdMainMenu 
      BackColor       =   &H00FF0000&
      Caption         =   "Return Back To Main Menu"
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
      Left            =   720
      MaskColor       =   &H00C000C0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   2535
   End
   Begin VB.CommandButton cmdFindStore 
      BackColor       =   &H00FF0000&
      Caption         =   "Click here to see if there is a Chipotle Store in your City."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   480
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   3255
   End
   Begin VB.PictureBox picStore 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1215
      Left            =   4080
      ScaleHeight     =   1155
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   1200
      Width           =   4935
   End
   Begin VB.Label lblName 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Designed by Carrie Hyland"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmStores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: ProjChipotleOrder (Carrie Hyland's VB Project.vbp)
'Form Name: frmStores (Stores_form.frm)
'Author: Carrie Hyland
'Date Written: October 19, 2003
'Purpose of Form: To allow the user to see if there is a Chipotle
                  ' store in their own city.
                 
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
'Declaring the variables that are global
Dim Found As Boolean, EnterStore As String
Dim CTR As Integer, Cities(1 To 12) As String
Dim Store As String, NotFound As Boolean
Private Sub cmdFindStore_Click()
'An input box pops up to allow the user to type in their city of choice
EnterStore = InputBox("Enter your Minnesota city to see if there is a Chipotle located there!  (Must capitalize the city and no abbreviations i.e. Saint Cloud.)", "Store Near You")
'Clears anything in the picStore screen to allow for repeated use
picStore.Cls
'Initializes the variable CTR to one
CTR = 1
'Open the file Cities.txt and allow the information to be put into an array
Open Path & "Cities.txt" For Input As #1
'Start loop to put information into an array
Do While Not EOF(1)
  Input #1, Cities(CTR)
  CTR = CTR + 1
Loop
'Reinitializing the variable CTR to one
CTR = 1
'Initializing the variable NotFound to True
NotFound = True
'Starts the search to find (or not find) the city entered
Do While NotFound And CTR <= 12
   If EnterStore = Cities(CTR) Then NotFound = False
   CTR = CTR + 1
Loop
If NotFound = False Then
    'If the city is found, it will print that there is a chipotle in that city
    picStore.Print EnterStore; Tab(13); "There is a Chipotle in this city."
   Else
    'If the city isn't found, it will print that there isn't a chipotle in that city
    picStore.Print EnterStore; Tab(13); "There is no Chipotle in this city."
End If
'Close file
Close
End Sub

Private Sub cmdMainMenu_Click()
'Hide the frmStores and show frmMainMenu
frmStores.Hide
frmMainMenu.Show
End Sub

Private Sub cmdQuit_Click()
'End the program
End
End Sub

Private Sub cmdClear_Click()
'An inputbox will appear to enter another city
EnterStore = InputBox("Enter your city to see if there is a Chipotle located there!", "Store Near You")
'Clears what was previously in the picStore screen to allow
'for repeated use
picStore.Cls
'Initializes the variable CTR to one
CTR = 1
'Opens the file Cities.txt to be used and then put into an array
Open Path & "Cities.txt" For Input As #1
'Starts the loop and puts the information into an array
Do While Not EOF(1)
  Input #1, Cities(CTR)
  CTR = CTR + 1
Loop
'Reinitializes the variable CTR to one
CTR = 1
'Initializes the variable NotFound to True
NotFound = True
'Starts the search to find (or not find) the city entered
Do While NotFound And CTR <= 12
   If EnterStore = Cities(CTR) Then NotFound = False
   CTR = CTR + 1
Loop
If NotFound = False Then
    'If the city is found, it will print that there is a chipotle in that city
    picStore.Print EnterStore, "There is a Chipotle in this city."
   Else
    'If the city isn't found, it will print that there isn't a chipotle in that city
    picStore.Print EnterStore, "There is no Chipotle in this city."
End If
Close
End Sub


