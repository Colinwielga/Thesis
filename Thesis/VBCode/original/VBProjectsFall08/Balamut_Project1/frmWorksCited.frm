VERSION 5.00
Begin VB.Form frmWorksCited 
   BackColor       =   &H00404000&
   Caption         =   "Works Cited"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit This Rad Program"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear List"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   5760
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   240
      ScaleHeight     =   3435
      ScaleWidth      =   9435
      TabIndex        =   2
      Top             =   120
      Width           =   9495
   End
   Begin VB.CommandButton cmdSee 
      Caption         =   "See the list"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   1
      Top             =   3960
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   3975
      Left            =   2400
      Picture         =   "frmWorksCited.frx":0000
      ScaleHeight     =   3915
      ScaleWidth      =   7275
      TabIndex        =   0
      Top             =   3960
      Width           =   7335
   End
End
Attribute VB_Name = "frmWorksCited"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Weezer
'Form Name: frmWorksCited.frm
'Author: Emily Balamut
'Date Written: 10/30/08
'Objective: This form shows the resources I used to compile the project when you
'click the button.
Option Explicit

Private Sub cmdClear_Click()
    picResults.Cls
End Sub

Private Sub cmdQuit_Click()
MsgBox "Thanks for rocking out with Weezer, " & UserName & "! See you later!", , "Bye!"
End
End Sub

Private Sub cmdSee_Click()
    picResults.Cls
    picResults.Print "For this project, I used many different resources."
    picResults.Print "I got all my pictures from images.google.com."
    picResults.Print "For the song lyrics, I used metrolyrics.com."
    picResults.Print "I modeled my project after the Virtual Hogwarts project which is located on the N: drive "
    picResults.Print "at N:\Classes\CS130\Trutwin_VB_Examples\Project Stuff\Sample Projects\A Virtual Hogwarts"
    picResults.Print "For all of the band member's information, I used Wikipedia."
    picResults.Print "All of the concert pictures are my own."
    picResults.Print "The track lists I got from my own iPod and the release dates"
    picResults.Print "are from weezer.com."
End Sub
