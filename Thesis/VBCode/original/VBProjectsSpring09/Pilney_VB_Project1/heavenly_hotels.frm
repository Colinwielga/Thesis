VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form8"
   ClientHeight    =   10860
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form8"
   ScaleHeight     =   10860
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdrent 
      Caption         =   "Continue to ski rentals"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   11640
      TabIndex        =   6
      Top             =   5400
      Width           =   3615
   End
   Begin VB.CommandButton cmdorder 
      Caption         =   "List hotels in order of price"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   11640
      TabIndex        =   5
      Top             =   2880
      Width           =   3735
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   360
      TabIndex        =   3
      Top             =   7200
      Width           =   2775
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Choose a different resort"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   11640
      TabIndex        =   2
      Top             =   7680
      Width           =   3615
   End
   Begin VB.PictureBox picResults 
      Height          =   5055
      Left            =   360
      ScaleHeight     =   4995
      ScaleWidth      =   10395
      TabIndex        =   1
      Top             =   840
      Width           =   10455
   End
   Begin VB.CommandButton cmdread 
      Caption         =   "View hotel listings and find best deal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   11640
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "The Squaw Valley Lodge"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   4
      Top             =   10080
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   4740
      Left            =   3720
      Picture         =   "heavenly_hotels.frx":0000
      Top             =   6000
      Width           =   7500
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Ski Trip
'Form Name: heavenly_hotels
'Author: Sam Pilney
'Written: March 17,2009
'this page shows the user what hotels are in the area the chose using a file input

Option Explicit
Dim Hotel(1 To 10) As String
Dim Price(1 To 10) As Single
Dim CTR As Integer
Dim J As Integer
'this subroutine reads the file that has the hotels stored for this particular resort and then
'displays them in the picture box
'it also uses an exhautive search to find the best deal
Private Sub cmdread_Click()

Dim Cheap As Single

CTR = 0

Open App.Path & "\heavenly.txt" For Input As #1
picResults.Cls
picResults.Print "Hotel"; Tab(40); "Price per person/per night"
picResults.Print "---------------------------------------------------------------------------------------------------------------"
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Hotel(CTR), Price(CTR)
    picResults.Print Hotel(CTR); Tab(40); FormatCurrency(Price(CTR))
    picResults.Print
Loop
Close #1

Cheap = 999999999
For J = 1 To CTR
    If Price(J) < Cheap Then
    Cheap = Price(J)
    End If
Next J
picResults.Print "The best deal is "; FormatCurrency(Cheap); " per person, per night."

cmdorder.Enabled = True
End Sub
'this subroutine is a bubble sort that displays the hotels in ascending order of price
Private Sub cmdorder_Click()

picResults.Cls

Dim Pass As Integer
Dim Pos As Integer
Dim Temp As Single
Dim TempTwo As String

For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If Price(Pos) > Price(Pos + 1) Then
            Temp = Price(Pos)
            Price(Pos) = Price(Pos + 1)
            Price(Pos + 1) = Temp
            TempTwo = Hotel(Pos)
            Hotel(Pos) = Hotel(Pos + 1)
            Hotel(Pos + 1) = TempTwo
        End If
    Next Pos
Next Pass

picResults.Print "Hotel"; Tab(40); "Price per person/per night"
picResults.Print "-------------------------------------------------------------------------------------------------------------"

'this function prints the hotel names and prices in order
For J = 1 To CTR
    picResults.Print Hotel(J); Tab(40); FormatCurrency(Price(J))
    picResults.Print
Next J
End Sub
'this subroutine bring the user to the ski rental form
Private Sub cmdrent_Click()
Form8.Hide
Form10.Show
End Sub
'this subroutine brings the user back to the beginning form
Private Sub cmdreturn_Click()

Form1.Show
Form6.Hide

End Sub


Private Sub cmdquit_Click()
End
End Sub

