VERSION 5.00
Begin VB.Form frmNameTwin 
   BackColor       =   &H00400000&
   Caption         =   "Name That Twin"
   ClientHeight    =   7935
   ClientLeft      =   2670
   ClientTop       =   1020
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   9885
   Begin VB.CommandButton cmdAgain 
      BackColor       =   &H000000C0&
      Caption         =   "Try Again"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H000000C0&
      Caption         =   "End"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6360
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      ScaleHeight     =   435
      ScaleWidth      =   3435
      TabIndex        =   29
      Top             =   7200
      Width           =   3495
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000C0&
      Caption         =   "Exit Twins Territory"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000C0&
      Caption         =   "Back to Twins Territory"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdBegin 
      BackColor       =   &H000000C0&
      Caption         =   "Begin"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   120
      Width           =   2295
   End
   Begin VB.PictureBox picName 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   7440
      ScaleHeight     =   5475
      ScaleWidth      =   2235
      TabIndex        =   24
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton cmdKubel 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   6240
      Picture         =   "frmNameTwin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdNeshek 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   5040
      Picture         =   "frmNameTwin.frx":7154
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdMorneau 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   3840
      Picture         =   "frmNameTwin.frx":E1EB
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdTyner 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   2640
      Picture         =   "frmNameTwin.frx":15189
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdHunter 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   1440
      Picture         =   "frmNameTwin.frx":1C2DD
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdTC 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   6240
      Picture         =   "frmNameTwin.frx":233C0
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdSilva 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   5040
      Picture         =   "frmNameTwin.frx":23BBA
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdMauer 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   3840
      Picture         =   "frmNameTwin.frx":2ADD0
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdPunto 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   2640
      Picture         =   "frmNameTwin.frx":31EB1
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdBaker 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   1440
      Picture         =   "frmNameTwin.frx":39152
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdLiriano 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   6240
      Picture         =   "frmNameTwin.frx":40294
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdCuddyer 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   5040
      Picture         =   "frmNameTwin.frx":472FE
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdGardenhire 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   3840
      Picture         =   "frmNameTwin.frx":4E455
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdBonsor 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   2640
      Picture         =   "frmNameTwin.frx":55595
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdSlowey 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   1440
      Picture         =   "frmNameTwin.frx":5C6E7
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdGuerrier 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   240
      Picture         =   "frmNameTwin.frx":63763
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdReyes 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   240
      Picture         =   "frmNameTwin.frx":6A81C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdNathan 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   240
      Picture         =   "frmNameTwin.frx":71B7F
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdSantana 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   6240
      Picture         =   "frmNameTwin.frx":78B94
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdRedmond 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   5040
      Picture         =   "frmNameTwin.frx":7FD20
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdBartlett 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   3840
      Picture         =   "frmNameTwin.frx":86D5E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdWhite 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   2640
      Picture         =   "frmNameTwin.frx":8E075
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdCrain 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   1440
      Picture         =   "frmNameTwin.frx":9510A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdCasilla 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   240
      Picture         =   "frmNameTwin.frx":9C2FE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblResults 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Results:"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   31
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H000000C0&
      Caption         =   $"frmNameTwin.frx":A33D9
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   25
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "frmNameTwin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim names(1 To 30) As String, answer As String
Dim correct As Integer, Ctr As Integer, total As Integer

Private Sub cmdAgain_Click()
'enable player buttons
    cmdBaker.Enabled = True
    cmdBonsor.Enabled = True
    cmdCrain.Enabled = True
    cmdGuerrier.Enabled = True
    cmdLiriano.Enabled = True
    cmdNathan.Enabled = True
    cmdNeshek.Enabled = True
    cmdSantana.Enabled = True
    cmdReyes.Enabled = True
    cmdSilva.Enabled = True
    cmdSlowey.Enabled = True
    cmdMauer.Enabled = True
    cmdRedmond.Enabled = True
    cmdBartlett.Enabled = True
    cmdCasilla.Enabled = True
    cmdMorneau.Enabled = True
    cmdPunto.Enabled = True
    cmdCuddyer.Enabled = True
    cmdHunter.Enabled = True
    cmdKubel.Enabled = True
    cmdTyner.Enabled = True
    cmdWhite.Enabled = True
    cmdGardenhire.Enabled = True
    cmdTC.Enabled = True

'reset correct and total
    correct = 0
    total = 0

'clear results box
    picResults.Cls

End Sub

Private Sub cmdBack_Click()
    frmNameTwin.Hide 'hides NameTwin form
    frmMain.Show 'shows main form
End Sub

Private Sub cmdBaker_Click()
    answer = InputBox("Name That Twin", "Name") 'set answer to read from input box
    total = total + 1 'adds one to the total Names answered
    If answer = "Scott Baker" Then correct = correct + 1 'counts if name answered was correct
    cmdBaker.Enabled = False 'disables player option

'enable End button
    cmdEnd.Enabled = True
End Sub

Private Sub cmdBartlett_Click()
    answer = InputBox("Name That Twin", "Name") 'set answer to read from input box
    total = total + 1 'adds one to the total Names answered
    If answer = "Jason Bartlett" Then correct = correct + 1 'counts if name answered was correct
    cmdBartlett.Enabled = False 'disables player option

'enable End button
    cmdEnd.Enabled = True
End Sub

Private Sub cmdBegin_Click()
'loads array of names from text file and prints

Ctr = 0 'initilize Ctr

Open App.Path & "\Names.txt" For Input As #1

Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, names(Ctr)
    picName.Print names(Ctr)
Loop

Close #1 'close file

cmdBegin.Enabled = False 'disable Begin button

'enable player buttons
    cmdBaker.Enabled = True
    cmdBonsor.Enabled = True
    cmdCrain.Enabled = True
    cmdGuerrier.Enabled = True
    cmdLiriano.Enabled = True
    cmdNathan.Enabled = True
    cmdNeshek.Enabled = True
    cmdSantana.Enabled = True
    cmdReyes.Enabled = True
    cmdSilva.Enabled = True
    cmdSlowey.Enabled = True
    cmdMauer.Enabled = True
    cmdRedmond.Enabled = True
    cmdBartlett.Enabled = True
    cmdCasilla.Enabled = True
    cmdMorneau.Enabled = True
    cmdPunto.Enabled = True
    cmdCuddyer.Enabled = True
    cmdHunter.Enabled = True
    cmdKubel.Enabled = True
    cmdTyner.Enabled = True
    cmdWhite.Enabled = True
    cmdGardenhire.Enabled = True
    cmdTC.Enabled = True

'enable Try Again button
    cmdAgain.Enabled = True
End Sub

Private Sub cmdBonsor_Click()
    answer = InputBox("Name That Twin", "Name") 'set answer to read from input box
    total = total + 1 'adds one to the total Names answered
    If answer = "Boof Bonsor" Then correct = correct + 1 'counts if name answered was correct
    cmdBonsor.Enabled = False 'disables player option

'enable End button
    cmdEnd.Enabled = True
End Sub

Private Sub cmdCasilla_Click()
    answer = InputBox("Name That Twin", "Name") 'set answer to read from input box
    total = total + 1 'adds one to the total Names answered
    If answer = "Alexi Casilla" Then correct = correct + 1 'counts if name answered was correct
    cmdCasilla.Enabled = False 'disables player option

'enable End button
    cmdEnd.Enabled = True
End Sub

Private Sub cmdCrain_Click()
    answer = InputBox("Name That Twin", "Name") 'set answer to read from input box
    total = total + 1 'adds one to the total Names answered
    If answer = "Jesse Crain" Then correct = correct + 1 'counts if name answered was correct
    cmdCrain.Enabled = False 'disables player option

'enable End button
    cmdEnd.Enabled = True
End Sub

Private Sub cmdCuddyer_Click()
    answer = InputBox("Name That Twin", "Name") 'set answer to read from input box
    total = total + 1 'adds one to the total Names answered
    If answer = "Michael Cuddyer" Then correct = correct + 1 'counts if name answered was correct
    cmdCuddyer.Enabled = False 'disables player option

'enable End button
    cmdEnd.Enabled = True
End Sub

Private Sub cmdEnd_Click()
'prints how many Twins the user named correctly
    picResults.Print "You named"; correct; "out of"; total; "Twins correctly"
'disable End button
    cmdEnd.Enabled = False
End Sub

Private Sub cmdExit_Click()
    End 'exits program
End Sub

Private Sub cmdGardenhire_Click()
    answer = InputBox("Name That Twin", "Name") 'set answer to read from input box
    total = total + 1 'adds one to the total Names answered
    If answer = "Ron Gardenhire" Then correct = correct + 1 'counts if name answered was correct
    cmdGardenhire.Enabled = False 'disables player option

'enable End button
    cmdEnd.Enabled = True
End Sub

Private Sub cmdGuerrier_Click()
    answer = InputBox("Name That Twin", "Name") 'set answer to read from input box
    total = total + 1 'adds one to the total Names answered
    If answer = "Matt Guerrier" Then correct = correct + 1 'counts if name answered was correct
    cmdGuerrier.Enabled = False 'disables player option

'enable End button
    cmdEnd.Enabled = True
End Sub

Private Sub cmdHunter_Click()
    answer = InputBox("Name That Twin", "Name") 'set answer to read from input box
    total = total + 1 'adds one to the total Names answered
    If answer = "Torii Hunter" Then correct = correct + 1 'counts if name answered was correct
    cmdHunter.Enabled = False 'disables player option

'enable End button
    cmdEnd.Enabled = True
End Sub

Private Sub cmdKubel_Click()
    answer = InputBox("Name That Twin", "Name") 'set answer to read from input box
    total = total + 1 'adds one to the total Names answered
    If answer = "Jason Kubel" Then correct = correct + 1 'counts if name answered was correct
    cmdKubel.Enabled = False 'disables player option

'enable End button
    cmdEnd.Enabled = True
End Sub

Private Sub cmdLiriano_Click()
    answer = InputBox("Name That Twin", "Name") 'set answer to read from input box
    total = total + 1 'adds one to the total Names answered
    If answer = "Francisco Liriano" Then correct = correct + 1 'counts if name answered was correct
    cmdLiriano.Enabled = False 'disables player option

'enable End button
    cmdEnd.Enabled = True
End Sub

Private Sub cmdMauer_Click()
    answer = InputBox("Name That Twin", "Name") 'set answer to read from input box
    total = total + 1 'adds one to the total Names answered
    If answer = "Joe Mauer" Then correct = correct + 1 'counts if name answered was correct
    cmdMauer.Enabled = False 'disables player option
    
'enable End button
    cmdEnd.Enabled = True
End Sub

Private Sub cmdMorneau_Click()
    answer = InputBox("Name That Twin", "Name") 'set answer to read from input box
    total = total + 1 'adds one to the total Names answered
    If answer = "Justin Morneau" Then correct = correct + 1 'counts if name answered was correct
    cmdMorneau.Enabled = False 'disables player option

'enable End button
    cmdEnd.Enabled = True
End Sub

Private Sub cmdNathan_Click()
    answer = InputBox("Name That Twin", "Name") 'set answer to read from input box
    total = total + 1 'adds one to the total Names answered
    If answer = "Joe Nathan" Then correct = correct + 1 'counts if name answered was correct
    cmdNathan.Enabled = False 'disables player option

'enable End button
    cmdEnd.Enabled = True
End Sub

Private Sub cmdNeshek_Click()
    answer = InputBox("Name That Twin", "Name") 'set answer to read from input box
    total = total + 1 'adds one to the total Names answered
    If answer = "Pat Neshek" Then correct = correct + 1 'counts if name answered was correct
    cmdNeshek.Enabled = False 'disables player option

'enable End button
    cmdEnd.Enabled = True
End Sub

Private Sub cmdPunto_Click()
    answer = InputBox("Name That Twin", "Name") 'set answer to read from input box
    total = total + 1 'adds one to the total Names answered
    If answer = "Nick Punto" Then correct = correct + 1 'counts if name answered was correct
    cmdPunto.Enabled = False 'disables player option

'enable End button
    cmdEnd.Enabled = True
End Sub

Private Sub cmdRedmond_Click()
    answer = InputBox("Name That Twin", "Name") 'set answer to read from input box
    total = total + 1 'adds one to the total Names answered
    If answer = "Mike Redmond" Then correct = correct + 1 'counts if name answered was correct
    cmdRedmond.Enabled = False 'disables player option

'enable End button
    cmdEnd.Enabled = True
End Sub

Private Sub cmdReyes_Click()
    answer = InputBox("Name That Twin", "Name") 'set answer to read from input box
    total = total + 1 'adds one to the total Names answered
    If answer = "Dennys Reyes" Then correct = correct + 1 'counts if name answered was correct
    cmdReyes.Enabled = False 'disables player option

'enable End button
    cmdEnd.Enabled = True
End Sub

Private Sub cmdSantana_Click()
    answer = InputBox("Name That Twin", "Name") 'set answer to read from input box
    total = total + 1 'adds one to the total Names answered
    If answer = "Johan Santana" Then correct = correct + 1 'counts if name answered was correct
    cmdSantana.Enabled = False 'disables player option

'enable End button
    cmdEnd.Enabled = True
End Sub

Private Sub cmdSilva_Click()
    answer = InputBox("Name That Twin", "Name") 'set answer to read from input box
    total = total + 1 'adds one to the total Names answered
    If answer = "Carlos Silva" Then correct = correct + 1 'counts if name answered was correct
    cmdSilva.Enabled = False 'disables player option

'enable End button
    cmdEnd.Enabled = True
End Sub

Private Sub cmdSlowey_Click()
    answer = InputBox("Name That Twin", "Name") 'set answer to read from input box
    total = total + 1 'adds one to the total Names answered
    If answer = "Kevin Slowey" Then correct = correct + 1 'counts if name answered was correct
    cmdSlowey.Enabled = False 'disables player option

'enable End button
    cmdEnd.Enabled = True
End Sub

Private Sub cmdTC_Click()
    answer = InputBox("Name That Twin", "Name") 'set answer to read from input box
    total = total + 1 'adds one to the total Names answered
    If answer = "T.C. Bear" Then correct = correct + 1 'counts if name answered was correct
    cmdTC.Enabled = False 'disables player option

'enable End button
    cmdEnd.Enabled = True
End Sub

Private Sub cmdTyner_Click()
    answer = InputBox("Name That Twin", "Name") 'set answer to read from input box
    total = total + 1 'adds one to the total Names answered
    If answer = "Jason Tyner" Then correct = correct + 1 'counts if name answered was correct
    cmdTyner.Enabled = False 'disables player option

'enable End button
    cmdEnd.Enabled = True
End Sub

Private Sub cmdWhite_Click()
    answer = InputBox("Name That Twin", "Name") 'set answer to read from input box
    total = total + 1 'adds one to the total Names answered
    If answer = "Rondell White" Then correct = correct + 1 'counts if name answered was correct
    cmdWhite.Enabled = False 'disables player option

'enable End button
    cmdEnd.Enabled = True
End Sub

