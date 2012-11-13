VERSION 5.00
Begin VB.Form Milan 
   BackColor       =   &H00C00000&
   Caption         =   "Form2"
   ClientHeight    =   10575
   ClientLeft      =   1860
   ClientTop       =   360
   ClientWidth     =   11880
   FillColor       =   &H00FF0000&
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form2"
   ScaleHeight     =   10575
   ScaleWidth      =   11880
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Look for More Places to Visit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4440
      TabIndex        =   4
      Top             =   1440
      Width           =   2655
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   7815
      Left            =   1680
      ScaleHeight     =   7755
      ScaleWidth      =   8355
      TabIndex        =   3
      Top             =   2760
      Width           =   8415
   End
   Begin VB.CommandButton cmdrestaurants 
      Caption         =   "Restaurants in Milan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7680
      TabIndex        =   2
      Top             =   1440
      Width           =   2655
   End
   Begin VB.CommandButton cmdshopping 
      Caption         =   "Shopping in Milan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1320
      TabIndex        =   1
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label lblMilan 
      BackColor       =   &H00C00000&
      Caption         =   "What Can You Do In Milan?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "Milan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Where to Travel in Italy
'Form Name: Milan
'Author: Sarah Dayton
'Date Written: March 23, 2009
'This form is to show the user what stores they can shop at in Milan and what restaurants are available in Milan
Option Explicit
Dim CTR As Integer

Private Sub cmdrestaurants_Click()
Dim restaurants(1 To 100) As String
picresults.Cls
CTR = 0
Open App.Path & "\milanrestaurants.txt" For Input As #1
    Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, restaurants(CTR)
    picresults.Print restaurants(CTR)
Loop
Close #1
End Sub

Private Sub cmdreturn_Click()
OpeningPage.Show
Milan.Hide
Venice.Hide
Florence.Hide
Rome.Hide
Naples.Hide
SlideShowItaly.Hide
End Sub

Private Sub cmdshopping_Click()
Dim shopping(1 To 100) As String
picresults.Cls
CTR = 0
Open App.Path & "\milanshopping.txt" For Input As #1
    Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, shopping(CTR)
    picresults.Print shopping(CTR)
Loop
Close #1
End Sub
