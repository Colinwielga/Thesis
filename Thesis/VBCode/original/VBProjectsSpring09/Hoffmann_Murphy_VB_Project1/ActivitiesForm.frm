VERSION 5.00
Begin VB.Form frmActivitiesForm 
   BackColor       =   &H00FFFF00&
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000080FF&
      Caption         =   "Back"
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdDates 
      BackColor       =   &H000080FF&
      Caption         =   "Check lodging availability"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4680
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   840
      Picture         =   "ActivitiesForm.frx":0000
      ScaleHeight     =   1335
      ScaleWidth      =   1935
      TabIndex        =   6
      Top             =   3000
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   5175
      Left            =   3480
      ScaleHeight     =   5115
      ScaleWidth      =   4155
      TabIndex        =   5
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton cmdActivitiesPrice 
      BackColor       =   &H000080FF&
      Caption         =   "Estimate your activities fees"
      Height          =   1095
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "4. Friday Fishing Contest"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   855
      Left            =   1800
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "3. Horseback riding"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "2. Golfing"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "1. Rent a jetski"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "frmActivitiesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Big Sky Resort
'frmActivitiesForm
'Ryan Hoffmann and Jamison Murphy
'Written on March 19, 2009
'This form was created so the user may enter what activities people
'may want to participate in and how much they cost
Option Explicit

Private Sub cmdActivitiesPrice_Click()

'Declare all variables
Dim People As Integer, Activity As Integer, Golf As Integer
Dim Jetski As Integer, Horse As Integer, Fish As Integer
Dim ActivitiesSubTotal As Single, Sum As Single

'Setting the variables numbers
Sum = 0
Golf = 75
Jetski = 50
Horse = 50
Fish = 20

'This clears anything that may be in the results box
picResults.Cls

'This is do while statement where the user must enter a corresponding
'number as seen on the form based on what activity they want to do.
'The user is able to choose all four activities if they wish and are
'asked how many people want to participate.  At the end it is printed
'out how much the desired activities will cost.
Activity = InputBox("Enter the number of the activity you would like to participate in.  Enter '0' to find total", , "Choose an activity")
Do While Activity <> 0
    If Activity = 1 Then
        People = InputBox("How many want to participate in Jetskiing?", , "People")
        ActivitiesSubTotal = People * Jetski
        Sum = Sum + ActivitiesSubTotal
        picResults.Print "Jetskiing will cost you "; FormatCurrency(ActivitiesSubTotal); "."
        Activity = InputBox("Enter the number of the activity you would like to participate in.  Enter '0' to find total", , "Choose an activity")
    ElseIf Activity = 2 Then
        People = InputBox("How many want to participate in Golfing?", , "People")
        ActivitiesSubTotal = People * Golf
        Sum = Sum + ActivitiesSubTotal
        picResults.Print "Golfing will cost you "; FormatCurrency(ActivitiesSubTotal); "."
        Activity = InputBox("Enter the number of the activity you would like to participate in.  Enter '0' to find total", , "Choose an activity")
    ElseIf Activity = 3 Then
        People = InputBox("How many want to participate in Horseback Riding?", , "People")
        ActivitiesSubTotal = People * Horse
        Sum = Sum + ActivitiesSubTotal
        picResults.Print "Horseback Riding will cost you "; FormatCurrency(ActivitiesSubTotal); "."
        Activity = InputBox("Enter the number of the activity you would like to participate in.  Enter '0' to find total", , "Choose an activity")
    ElseIf Activity = 4 Then
        People = InputBox("How many want to participate in Fishing?", , "People")
        ActivitiesSubTotal = People * Fish
        Sum = Sum + ActivitiesSubTotal
        picResults.Print "Fishing will cost you "; FormatCurrency(ActivitiesSubTotal); "."
        Activity = InputBox("Enter the number of the activity you would like to participate in.  Enter '0' to find total", , "Choose an activity")
    Else
        MsgBox "Please enter a valid number.", , "Error."
        Activity = InputBox("Enter the number of the activity you would like to participate in.  Enter '0' to find total", , "Choose an activity")
    End If
Loop

'This prints out the total of how much all the desired activities will cost."
ActivitiesTotal = Sum
picResults.Print "______________________________"
picResults.Print "Activities Total="; FormatCurrency(ActivitiesTotal); "."
End Sub

'This command goes back to previous form
Private Sub cmdBack_Click()
    frmPricingForm1.Show
    frmActivitiesForm.Hide
End Sub

'This command moves onto the next form
Private Sub cmdDates_Click()
    frmAvailabilityForm.Show
    frmActivitiesForm.Hide
End Sub

Private Sub Form_Load()

'This code centers the form on computer screen upon loading

Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

End Sub

