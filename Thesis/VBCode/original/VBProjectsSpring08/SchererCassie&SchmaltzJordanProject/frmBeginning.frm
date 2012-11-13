VERSION 5.00
Begin VB.Form frmBeginning 
   BackColor       =   &H00FFFF00&
   Caption         =   "Form1"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12990
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   12990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Quit"
      Height          =   855
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton cmdmoderate 
      BackColor       =   &H0000C000&
      Caption         =   "Moderate"
      Height          =   855
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmdhot 
      BackColor       =   &H000000FF&
      Caption         =   "Hot"
      Height          =   855
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   3990
      Left            =   6720
      Picture         =   "frmBeginning.frx":0000
      Top             =   1440
      Width           =   5970
   End
   Begin VB.Image Image1 
      Height          =   4620
      Left            =   360
      Picture         =   "frmBeginning.frx":A2EA
      Top             =   1320
      Width           =   6150
   End
   Begin VB.Label lbldirections 
      BackColor       =   &H0000FF00&
      Caption         =   "Please select which climate you prefer for your dream vacation!!!"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   12495
   End
End
Attribute VB_Name = "frmBeginning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Travel Agency
'Form Name: frmBeginning
'Authors: Cassie Scherer and Jordan Schmaltz
'Date Written: 3/4/08
'Objective:This is our starting start up form for the project
'Here the user can select which climate they prefer to vacation in, and what atmosphere
'The overall objective for our project is centered around a travel agency
'We guide the user along in choosing their vacation destination, their hotel preference, booking their hotel,
'selecting their activites, viewing possible flight plans, and selecting a rental car.


Private Sub cmdhot_Click()

'Here we are declaring our variables

Dim choice As Integer

'Here we are giving the user two options for hot climate vacations.
'We use an input box to allow them to decide what type of atmosphere they prefer.

choice = InputBox("Type a 1 if you want a more lively, city vacation, or a 2 for a more scenic, realaxing vacation.")

'Here we used an else if then statement to direct the user to the correct form of their dream location.

If choice = 1 Then
        frmBeginning.Hide
        frmVegasStart.Show
    ElseIf choice = 2 Then
        frmBeginning.Hide
        frmJamaicaStart.Show
    Else: MsgBox ("Sorry, you entered an invalid option plese enter either a 1 or a 2.")
End If

End Sub

Private Sub cmdmoderate_Click()

'Here we declared our variables for this command button

Dim choice As Integer

'Here we are giving the user two options for moderate climate vacations.
'We use an input box to allow them to decide what type of language they prefer the natives to speak within the country they are traveling to.

choice = InputBox("Type in a 1 if you prefer an English speaking country, or type in a 2 if a foreign speaking country interests you.")

'Here we used an else if then statement to direct the user to the correct form of their dream location.

If choice = 1 Then
        frmBeginning.Hide
        frmBostonStart.Show
    ElseIf choice = 2 Then
        frmBeginning.Hide
        frmItalyStart.Show
    Else: MsgBox ("Sorry, you entered an invalid option please enter either a 1 or a 2.")
End If

End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub Form_Load()

'This code centers the form on computer screen upon loading

Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

End Sub
