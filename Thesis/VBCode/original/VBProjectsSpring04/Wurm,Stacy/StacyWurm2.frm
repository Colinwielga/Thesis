VERSION 5.00
Begin VB.Form GiftToGive 
   BackColor       =   &H00FF8080&
   Caption         =   "FirstChoice"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optBook 
      BackColor       =   &H00FF8080&
      Height          =   255
      Left            =   8040
      TabIndex        =   16
      Top             =   1680
      Width           =   255
   End
   Begin VB.PictureBox picBooks 
      BackColor       =   &H00FF8080&
      Height          =   1935
      Left            =   7320
      Picture         =   "StacyWurm2.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   15
      Top             =   2040
      Width           =   1575
   End
   Begin VB.OptionButton optChocolate 
      BackColor       =   &H00FF8080&
      Height          =   255
      Left            =   5880
      TabIndex        =   14
      Top             =   1680
      Width           =   255
   End
   Begin VB.PictureBox picChocolate 
      BackColor       =   &H00FF8080&
      Height          =   1935
      Left            =   5160
      Picture         =   "StacyWurm2.frx":0DD9
      ScaleHeight     =   1875
      ScaleWidth      =   1635
      TabIndex        =   13
      Top             =   2040
      Width           =   1695
   End
   Begin VB.OptionButton optDVDCD 
      BackColor       =   &H00FF8080&
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   1680
      Width           =   255
   End
   Begin VB.PictureBox picDVDCD 
      BackColor       =   &H00FF8080&
      Height          =   1935
      Left            =   720
      Picture         =   "StacyWurm2.frx":1C2C
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   11
      Top             =   2040
      Width           =   1575
   End
   Begin VB.OptionButton optFlower 
      BackColor       =   &H00FF8080&
      Height          =   255
      Left            =   3720
      TabIndex        =   10
      Top             =   1680
      Width           =   255
   End
   Begin VB.PictureBox picFlowers 
      BackColor       =   &H00FF8080&
      Height          =   1695
      Left            =   3360
      Picture         =   "StacyWurm2.frx":28CC
      ScaleHeight     =   1635
      ScaleWidth      =   915
      TabIndex        =   9
      Top             =   2160
      Width           =   975
   End
   Begin VB.PictureBox picResults1 
      Height          =   975
      Left            =   600
      ScaleHeight     =   915
      ScaleWidth      =   8715
      TabIndex        =   4
      Top             =   4080
      Width           =   8775
   End
   Begin VB.CommandButton cmdNext1 
      Caption         =   "Next choice to be made is:"
      Enabled         =   0   'False
      Height          =   735
      Left            =   4440
      TabIndex        =   3
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton cmdChoice1 
      Caption         =   "I have made my choice of gift!!!"
      Enabled         =   0   'False
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   8400
      TabIndex        =   1
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Other 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "A New Book"
      Height          =   255
      Left            =   7560
      TabIndex        =   8
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Chocolates 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Chocolates"
      Height          =   255
      Left            =   5400
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Flowers 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Flowers"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label DVDCD 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "A New DVD or CD"
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Gift 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "First you must choose the gift you will either bring your date or you would like to recieve from your date."
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Width           =   5415
   End
End
Attribute VB_Name = "GiftToGive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Project Name: Date Chooser (Wurm, Stacy - VB Project)
' Form Name: DateIntro (StacyWurmIntro.frm)
' Author: Stacy Wurm
' Date Written: Sunday, March 7th, 2004
' Purpose of this Form: ' Allows the user to choose a gift
                        ' it asks your choice and totals how much spent
                        ' Then also displays amount spent so far

Private Sub cmdChoice1_Click()
' First option and display total and choice
    If optDVDCD = True Then
        Cost = 15
        Choice = "a brand new DVD or CD!!  Hope she likes it!!"
        TotalCost = TotalCost + Cost
        Decision1 = "DVD\CD"
    ElseIf optFlower = True Then
        Cost = 20
        Choice = "flowers!!  Do you know her favorite??"
        TotalCost = TotalCost + Cost
        Decision1 = "Flower"
    ElseIf optChocolate = True Then
        Cost = 10
        Choice = "chocolates!!  MMM mmm...always a yummy choice!!"
        TotalCost = TotalCost + Cost
        Decision1 = "Chocolates"
    ElseIf optBook = True Then
        Cost = 8
        Choice = "a new book!!  Focusing on the intellectual side."
        TotalCost = TotalCost + Cost
        Decision1 = "a book"
    End If

' Tells what should be printed
picResults1.Print "You have decided to give "; Choice
picResults1.Print "This is going to take "; FormatCurrency(Cost); " away from your total!!"
picResults1.Print "So far you have spent "; FormatCurrency(TotalCost); " on your date."
cmdNext1.Enabled = True
End Sub

Private Sub cmdNext1_Click()
' moves the user to the next form in the program
GiftToGive.Hide
DateEvent.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub
Private Sub optBook_Click()
' Allows this option to be choosen
picResults1.Cls
cmdChoice1.Enabled = True
End Sub

Private Sub optChocolate_Click()
' Allows this option to be choosen
picResults1.Cls
cmdChoice1.Enabled = True
End Sub

Private Sub optDVDCD_Click()
' Allows this option to be choosen
picResults1.Cls
cmdChoice1.Enabled = True
End Sub

Private Sub optFlower_Click()
' Allows this option to be choosen
picResults1.Cls
cmdChoice1.Enabled = True
End Sub
