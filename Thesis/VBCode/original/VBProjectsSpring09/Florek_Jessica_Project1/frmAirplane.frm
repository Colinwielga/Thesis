VERSION 5.00
Begin VB.Form frmAirplane 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8790
   FillColor       =   &H00C0FFC0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   8790
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtflightinput 
      Height          =   495
      Left            =   7080
      TabIndex        =   6
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Continue"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdBook 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Book Flight"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6240
      Width           =   1455
   End
   Begin VB.PictureBox picResults2 
      Height          =   2655
      Left            =   5640
      ScaleHeight     =   2595
      ScaleWidth      =   2835
      TabIndex        =   2
      Top             =   3960
      Width           =   2895
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Show Flight Prices in decending order"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   5415
      Left            =   120
      Picture         =   "frmAirplane.frx":0000
      ScaleHeight     =   5355
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   600
      Width           =   5415
      Begin VB.Shape Shape5 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   960
         Shape           =   3  'Circle
         Top             =   3840
         Width           =   135
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   3360
         Shape           =   3  'Circle
         Top             =   3720
         Width           =   135
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   2040
         Shape           =   3  'Circle
         Top             =   2520
         Width           =   135
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   2040
         Shape           =   3  'Circle
         Top             =   2040
         Width           =   135
      End
   End
   Begin VB.Label lblinfo 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Please enter the name of the city you wish to fly to below and click ""Book Fllight"""
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      TabIndex        =   7
      Top             =   1560
      Width           =   2775
   End
End
Attribute VB_Name = "frmAirplane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FinalProject:Travel Europe
'frmAirplane
'Jessica Florek
'Written: 3/4/09
'Objective: have user select a city to fly into and have the amount of
'that plane ticket subtracted from their total budget.


Option Explicit

Dim city(1 To 10) As String, price(1 To 10) As Single, ctr As Integer


Private Sub cmdBook_Click()
Dim flightname As String, found As Boolean, I As Integer
found = False

'This takes the users inputed city name and matches it to the city name in the array and uses the price in the array that is associated with that city
flightname = txtflightinput
Do While ((Not found) And (I < ctr))
    I = I + 1
    If flightname = city(I) Then
        budget = budget - price(I)
        flightcost = price(I)
        MsgBox ("You have chosen to fly to " & city(I) & " which will cost " & FormatCurrency(price(I)) & " which has been subtracted from your budget.")
        found = True
        cmdBook.Enabled = False
    End If
Loop

'If city mistyped or not an option this is the default message
If (Not found) Then
    MsgBox ("You have not entered a valid city. Please check your spelling and try again.")
End If

cmdContinue.Enabled = True


End Sub



Private Sub cmdContinue_Click()
'take user to next step in the project, choosing the cities
frmAirplane.Hide
frmMapCities.Show

End Sub


Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdSort_Click()
Dim pass As Integer, pos As Integer, temp As Integer, temp2 As String, I As Integer


'inputs information from file into an array so that the information is available when the user chooses a flight
Open App.Path & "\planeprices.txt" For Input As #1

Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, city(ctr), price(ctr)
Loop
Close #1

'sorts the plane tickets in order of cheapest to most expensive. Done with an array so that they prices can be changed and modified as plane ticket prices change often.
For pass = 1 To (ctr - 1)
    For pos = 1 To (ctr - pass)
        If price(pos) > price(pos + 1) Then
            temp = price(pos)
            price(pos) = price(pos + 1)
            price(pos + 1) = temp
            temp2 = city(pos)
            city(pos) = city(pos + 1)
            city(pos + 1) = temp2
        End If
    Next pos
Next pass

picResults2.Print "Prices of tickets in decending order"
picResults2.Print "*******************************************"
picResults2.Print

For I = 1 To ctr
    picResults2.Print city(I), price(I)
Next I
            
            
cmdBook.Enabled = True
cmdSort.Enabled = False

End Sub


