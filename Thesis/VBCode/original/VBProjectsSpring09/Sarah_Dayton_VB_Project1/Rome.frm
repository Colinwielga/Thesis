VERSION 5.00
Begin VB.Form Rome 
   BackColor       =   &H00C000C0&
   Caption         =   "Form5"
   ClientHeight    =   10575
   ClientLeft      =   1860
   ClientTop       =   450
   ClientWidth     =   11880
   FillColor       =   &H00C00000&
   LinkTopic       =   "Form5"
   ScaleHeight     =   10575
   ScaleWidth      =   11880
   Begin VB.CommandButton cmdprice 
      Caption         =   "I Only Have So Much Money"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      TabIndex        =   5
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdalphabetical 
      Caption         =   "Alphabetical"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Pick Another City To Visit Instead"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8520
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00C000C0&
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   7935
      Left            =   1320
      ScaleHeight     =   7875
      ScaleWidth      =   9315
      TabIndex        =   2
      Top             =   2520
      Width           =   9375
   End
   Begin VB.CommandButton cmdread 
      Caption         =   "Read the File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C000C0&
      Caption         =   "What To Do In Rome?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "Rome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title: Where to Travel in Italy
'Form Name: Rome
'Author: Sarah Dayton
'This form is to show what you can see in Rome in alphabetical order as well as by price
Option Explicit
Dim places(1 To 100) As String, prices(1 To 100) As Single, CTR As Integer, I As Integer

Private Sub cmdalphabetical_Click()
picresults.Cls
Dim pass As Integer, pos As Integer, temp As String
For pass = 1 To CTR - 1
    For pos = 1 To CTR - pass
        If places(pos) > places(pos + 1) Then
            temp = places(pos)
            places(pos) = places(pos + 1)
            places(pos + 1) = temp
            temp = prices(pos)
            prices(pos) = prices(pos + 1)
            prices(pos + 1) = temp
        End If
    Next pos
Next pass
picresults.Print "Places to see"; Tab(30); "Price"
picresults.Print "***********************************************"

For I = 1 To CTR
    picresults.Print places(I); Tab(30); FormatCurrency(prices(I))
Next I
End Sub

Private Sub cmdprice_Click()
Dim money As String, Found As Boolean
picresults.Cls
money = InputBox("How much money would you like to spend per place to visit?")
Found = False

For I = 1 To CTR
    If money >= prices(I) Then
        picresults.Print places(I); Tab(35); FormatCurrency(prices(I))
        Found = True
    End If
Next I

If Not Found Then
    picresults.Print "Sorry, there is nothing for you to do."
End If


        
End Sub

Private Sub cmdread_Click()
picresults.Print "Places to see"; Tab(30); "Price"
picresults.Print "**************************************************************"
Open App.Path & "\rome.txt" For Input As #1
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, places(CTR), prices(CTR)
    picresults.Print places(CTR); Tab(35); FormatCurrency(prices(CTR))
Loop

End Sub

Private Sub cmdreturn_Click()
Close #1
OpeningPage.Show
Milan.Hide
Venice.Hide
Florence.Hide
Rome.Hide
Naples.Hide
SlideShowItaly.Hide
End Sub
