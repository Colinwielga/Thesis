VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H8000000D&
   Caption         =   "The Record -- Choosing a classified ad"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9990
   LinkTopic       =   "Form7"
   ScaleHeight     =   6180
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   7920
      Picture         =   "classified.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   480
      Picture         =   "classified.frx":7572
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   2760
      Picture         =   "classified.frx":EAE4
      ScaleHeight     =   1515
      ScaleWidth      =   4395
      TabIndex        =   3
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton classbutton 
      Caption         =   "Click here to calculate the total cost of a classified ad."
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   2
      Top             =   1920
      Width           =   2655
   End
   Begin VB.CommandButton Quitbutton 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   1
      Top             =   5640
      Width           =   1575
   End
   Begin VB.PictureBox Results 
      BackColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   3480
      ScaleHeight     =   3315
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   1920
      Width           =   6015
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project1 (Record_Advertising), Form 7(classified.frm), Kristen Nowak, 3-14-04, The purpose of this form is to allow the user to calculate the cost of a classified ad.

Private Sub classbutton_Click()
Dim Words As Integer, Addwords As Integer, Fifteen As String
Words = InputBox("How many words is the classified ad?") 'ask the user how many words their classified ad is
If Words <= 25 Then 'a classified ad of 25 words or less costs $5
    Total = 5
    Results.Print "A classified ad of"; Words; "words will cost "; FormatCurrency(Total)
ElseIf Words > 25 Then 'a classified ad of more than 25 words costs $5 plus $.15 per extra word
    Addwords = Words - 25
    Total = (Addwords * 0.15) + 5
    Results.Print "A classified ad of"; Words; "words will cost "; FormatCurrency(Total)
End If

Fifteen = InputBox("If you are placing a classified ad that is selling a personal item, you have the option of running the ad until the item sells for an additional $15. Would you like to choose this option? Type 1 for 'yes' or 2 for 'no'.")
If Fifteen = 1 Then
    Total = Total + 15 'If yes, add $15
    Results.Print "The cost of your classified ad with the run-until-it-sells option is "; FormatCurrency(Total); "."
    Results.Print "You have now completed the program. Thank you for your business."
ElseIf Fifteen = 2 Then
    Weeks = InputBox("How many weeks do you want your ad to run?")
    Total = Total * Weeks 'If no, multiply the cost of one ad by the number of weeks the user wants to run the ad
    Results.Print "To run your ad for "; Weeks; " weeks, the cost of your classified ad is "; FormatCurrency(Total); "."
    Results.Print "You have now completed the program. Thank you for your business."
End If
classbutton.Enabled = False
End Sub

Private Sub Quitbutton_Click()
End
End Sub
