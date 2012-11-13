VERSION 5.00
Begin VB.Form FrmTotal 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdStore 
      Caption         =   "Shopping Store"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2160
      TabIndex        =   3
      Top             =   4680
      Width           =   2895
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2160
      TabIndex        =   2
      Top             =   6480
      Width           =   2895
   End
   Begin VB.CommandButton cmdScore 
      Caption         =   "Show Score"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2160
      TabIndex        =   1
      Top             =   2880
      Width           =   2895
   End
   Begin VB.PictureBox PicResults 
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   5760
      ScaleHeight     =   4957.282
      ScaleMode       =   0  'User
      ScaleWidth      =   10740
      TabIndex        =   0
      Top             =   2880
      Width           =   10800
   End
   Begin VB.Image Image1 
      Height          =   10275
      Left            =   240
      Picture         =   "FrmTotal.frx":0000
      Top             =   120
      Width           =   17820
   End
End
Attribute VB_Name = "FrmTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Total As Integer

    'This form prints the total correct answers the user had gotten throughout the program as well as how many the user got correct in each subject.
    'Since its the end of the game, it allows the user to eithe quit or go to the store form to shop.


Private Sub cmdQuit_Click()
    'Ends the program.
End
End Sub

Private Sub cmdScore_Click()
    'It prints out the total correct answers the user had throughout the program.
    'It prints out the amount of correct answers per subject.
    'It uses the if statements to tell the user different levels of accomplishment due to how many answers the user got correct.
picResults.Cls
Total = 0
Total = Total + Sum

If Total < 5 Then
    
    picResults.Print UserName; ", you had "; Total; " questions correct"
    picResults.Print "You had "; SumHoc; " hockey questions correct"
    picResults.Print "You had "; SumLax; " lacrosse questions correct"
    picResults.Print "You had "; SumBase; " baseball questions correct"
    picResults.Print "That was not very good, Try again."

ElseIf Total < 10 Then

    picResults.Print UserName; ", you had "; Total; " questions correct"
    picResults.Print "You had "; SumHoc; " hockey questions correct"
    picResults.Print "You had "; SumLax; " lacrosse questions correct"
    picResults.Print "You had "; SumBase; " baseball questions correct"
    picResults.Print "You are getting pretty good, Try again."
    
ElseIf Total <= 15 Then

    picResults.Print UserName; ", you had "; Total; " questions correct"
    picResults.Print "You had "; SumHoc; " hockey questions correct"
    picResults.Print "You had "; SumLax; " lacrosse questions correct"
    picResults.Print "You had "; SumBase; " baseball questions correct"
    picResults.Print "That was amazing!"
    
End If
End Sub


Private Sub CmdStore_Click()
    'This transfers forms from the total form to the store form.
FrmTotal.Hide
FrmStore.Show
End Sub
