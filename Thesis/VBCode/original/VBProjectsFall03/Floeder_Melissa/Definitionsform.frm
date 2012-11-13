VERSION 5.00
Begin VB.Form Definitionsform 
   BackColor       =   &H00800000&
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResults 
      Height          =   3615
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   120
      Width           =   5175
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Clear"
      Height          =   735
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Quit"
      Height          =   735
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Test your understanding."
      Enabled         =   0   'False
      Height          =   735
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Enter an Element of Design."
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   1095
   End
End
Attribute VB_Name = "Definitionsform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: DesignElements
'Project file name: Elements of design project.vbp
'Form name: Definitionsform
'Form file name: Definitionsform.frm
'Author's name: Melissa Floeder
'Date: 10/20/03
'Purpose: To define the elements of design and test the user to see if they can recognize them in an image
Option Explicit
Dim Definitions(1 To 20) As String, CTR As Integer, Chosen As String, J As Integer, NotFound As Boolean, CTR2 As Integer
Dim strMess As String
Dim PATH As String

'clear the results
Private Sub cmdClear_Click()
txtResults.Text = ""
strMess = ""
End Sub
Private Sub cmdFind_Click()

CTR = 0
CTR2 = 0
J = 0
NotFound = True

'creating the arrays
Open (PATH & "Elements.txt") For Input As #1
Open (PATH & "pictureElements.txt") For Input As #2
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Elements(CTR), Definitions(CTR)
Loop
Do While Not EOF(2)
    CTR2 = CTR2 + 1
    Input #2, One(CTR2), Two(CTR2), Three(CTR2), Four(CTR2), Five(CTR2)
Loop
Close #1
Close #2
'entering the element
Chosen = InputBox("Enter an element of design: texture, color, line, shape, size, direction, value.")
NotFound = True
'searching the arrays for the informantion to print
Do While NotFound And J <= CTR
    J = J + 1
    If Chosen = Elements(J) Then NotFound = False
Loop
'Printing results in a textbox as a string message
If NotFound = False Then
    strMess = strMess & Chosen
    strMess = strMess & ": "
    strMess = strMess & vbNewLine & vbNewLine
    strMess = strMess & Definitions(J)
    strMess = strMess & vbNewLine & vbNewLine & vbNewLine
    txtResults.Text = strMess
    Else
    strMess = strMess & "Your entry was not found.  Please try again."
    strMess = strMess & vbNewLine & vbNewLine & vbNewLine
    txtResults.Text = strMess
End If
'enabling other button
cmdGo.Enabled = True
End Sub
'going to other form (testform)
Private Sub cmdGo_Click()
Testform.Show
Definitionsform.Hide
End Sub
Private Sub cmdQuit_Click()
End
End Sub

Private Sub Form_Load()
PATH = "N:\CS130\Handin\Floeder_Melissa\"
End Sub
