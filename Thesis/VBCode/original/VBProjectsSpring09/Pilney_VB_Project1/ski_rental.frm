VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form10"
   ClientHeight    =   9015
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form10"
   ScaleHeight     =   9015
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdyouth 
      Caption         =   "Select your price range for youth ski rentals"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   3135
   End
   Begin VB.CommandButton cmdback 
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
      Height          =   2655
      Left            =   7080
      TabIndex        =   3
      Top             =   5160
      Width           =   2895
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
      Height          =   2655
      Left            =   10560
      TabIndex        =   2
      Top             =   5160
      Width           =   2895
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   3360
      ScaleHeight     =   3795
      ScaleWidth      =   10515
      TabIndex        =   1
      Top             =   240
      Width           =   10575
   End
   Begin VB.CommandButton cmdrent 
      Caption         =   "Select your price range for adult ski rentals"
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
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin VB.Image Image5 
      Height          =   1200
      Left            =   1080
      Picture         =   "ski_rental.frx":0000
      Top             =   7320
      Width           =   2250
   End
   Begin VB.Image Image4 
      Height          =   720
      Left            =   4560
      Picture         =   "ski_rental.frx":8D82
      Top             =   6000
      Width           =   1920
   End
   Begin VB.Image Image3 
      Height          =   2820
      Left            =   240
      Picture         =   "ski_rental.frx":D5C4
      Top             =   4320
      Width           =   3750
   End
   Begin VB.Image Image2 
      Height          =   1485
      Left            =   4440
      Picture         =   "ski_rental.frx":2FE46
      Top             =   4320
      Width           =   2250
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   4680
      Picture         =   "ski_rental.frx":3AD54
      Top             =   6960
      Width           =   1755
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Ski Trip
'Form Name: ski_rental
'Author: Sam Pilney
'Written: March 19,2009
'this page allows the user to find out what skis they should rent

Option Explicit
Dim Brand As String
Dim Price As Single
Dim Range As Single
'this subroutine uses a select case function to find which brand best fits the users price range for adult skis
'the user inputs their own price range via input box
Private Sub cmdrent_Click()

Range = InputBox("Please enter how much you want to spend per adult rental.", "Price Range")

Select Case Range
     Case Is < 65
        Brand = "Fischer"
        Price = 59
     Case 65 To 75
        Brand = "Rossingnol"
        Price = 75
    Case 76 To 84
        Brand = "Atomic"
        Price = 80
    Case 85 To 94
        Brand = "K2"
        Price = 85
    Case Is > 94
        Brand = "Volkl"
        Price = 95
    Case Else
        MsgBox "Please enter a valid price.", , "Error"
End Select

picResults.Cls
picResults.Print Brand; " skis best fit your price range at "; FormatCurrency(Price); " per person, per day."

End Sub

'this subroutine brings the user back to the beginning form
Private Sub cmdback_Click()
Form1.Show
Form10.Hide
End Sub

Private Sub cmdquit_Click()
End
End Sub
'this subroutine uses a select case function to find which brand best fits the users price range for youth skis
'the user inputs their own price range via input box
Private Sub cmdyouth_Click()
Range = InputBox("Please enter how much you want to spend per youth rental.", "Price Range")

Select Case Range
     Case Is < 40
        Brand = "Fischer"
        Price = 35
     Case 40 To 45
        Brand = "Rossingnol"
        Price = 45
     Case 46 To 60
        Brand = "Atomic"
        Price = 55
     Case 61 To 74
        Brand = "K2"
        Price = 70
     Case Is > 74
        Brand = "Volkl"
        Price = 80
     Case Else
        MsgBox "Please enter a valid price.", , "Error"
End Select

picResults.Cls
picResults.Print Brand; " skis best fit your price range at "; FormatCurrency(Price); " per person, per day."
End Sub

