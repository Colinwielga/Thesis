VERSION 5.00
Begin VB.Form frmMiller 
   BackColor       =   &H00000080&
   Caption         =   "Miller Brewing Co. Form"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14280
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   9345
   ScaleWidth      =   14280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMainMenu 
      BackColor       =   &H0000FFFF&
      Caption         =   "Main Menu"
      Height          =   375
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8880
      Width           =   3255
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8880
      Width           =   3015
   End
   Begin VB.CommandButton cmdMGD 
      BackColor       =   &H0000FFFF&
      Caption         =   "Miller Genuine Draft"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7920
      Width           =   2895
   End
   Begin VB.CommandButton cmdMillerHigh 
      BackColor       =   &H0000FFFF&
      Caption         =   "Miller High Life"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6960
      Width           =   2895
   End
   Begin VB.CommandButton cmdLeinenkugel 
      BackColor       =   &H0000FFFF&
      Caption         =   "Leinenkugel's"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6000
      Width           =   2895
   End
   Begin VB.CommandButton cmdMillerLite 
      BackColor       =   &H0000FFFF&
      Caption         =   "Miller Lite"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   2895
   End
   Begin VB.PictureBox picResults 
      ForeColor       =   &H0000FFFF&
      Height          =   8655
      Left            =   3120
      Picture         =   "frmMiller.frx":0000
      ScaleHeight     =   8595
      ScaleWidth      =   10755
      TabIndex        =   2
      Top             =   120
      Width           =   10815
   End
   Begin VB.CommandButton cmdLength 
      BackColor       =   &H0000FFFF&
      Caption         =   "Click Here to List the Beers in Order of Longest Name to Shortest Name"
      Height          =   1575
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton cmdList 
      BackColor       =   &H0000FFFF&
      Caption         =   "(Click First): List All of the Beers that Miller Brewing Co. Offers"
      Height          =   1575
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label lblLearnMiller 
      BackColor       =   &H00000000&
      Caption         =   "Click on a Beer to Learn More About It"
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   2775
   End
End
Attribute VB_Name = "frmMiller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Beer Experience
'frm Miller
'Lauren Gooley and Tim Janssen
'March 21, 2008
'This form allows the user to learn more about the specific brands that the Miller Brewing company provides to consumers.
Option Explicit
Dim Beers(1 To 100) As String
Dim CTR As Single


Private Sub cmdLeinenkugel_Click()
MsgBox ("Leinenkugel's has been brewed in Chippewa Falls, Wisconson, since 1867. It is the leading craft brewer in the upper Midwest.")
End Sub
'This subroutine searches and prints the names of the brands of Miller beer from the longest name to the shortest name.
Private Sub cmdLength_Click()
Dim Pass As Single, POS As Single, TempLength As String, J As Integer
picResults.Cls
For Pass = 1 To CTR - 1
    For POS = 1 To CTR - Pass
        If Len(Beers(POS)) < Len(Beers(POS + 1)) Then
            TempLength = Beers(POS)
            Beers(POS) = Beers(POS + 1)
            Beers(POS + 1) = TempLength
        End If
    Next POS
Next Pass
For J = 1 To CTR
    picResults.Print Beers(J)
Next J
            
End Sub
'This subroutine reads the list of brands into an array.
Private Sub cmdList_Click()
Dim J As Integer
Open App.Path & "\MillerBeers.txt" For Input As #1
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Beers(CTR)
Loop
For J = 1 To CTR
    picResults.Print Beers(J)
Next J
Close #1
End Sub

Private Sub cmdMainMenu_Click()
frmMiller.Hide
Companies.Show
End Sub

Private Sub cmdMGD_Click()
MsgBox ("Miller Genuine Draft debuted in 1985 with fresh, smooth flavor that's a result of being cold-filtered four times.")
End Sub

Private Sub cmdMillerHigh_Click()
MsgBox ("The champagne of beers dates back to 1903, and is a classic American-style lager recognized for its consistently crisp, smooth taste, and classic clear bottle.")
End Sub

Private Sub cmdMillerLite_Click()
MsgBox ("Miller Lite is the great tasting, less filling beer that defined the American light beer category in 1975.")
End Sub

Private Sub cmdQuit_Click()
End
End Sub
