VERSION 5.00
Begin VB.Form Introduction 
   Caption         =   "Introduction"
   ClientHeight    =   12915
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12885
   BeginProperty Font 
      Name            =   "Old English Text MT"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   12915
   ScaleWidth      =   12885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdIntro 
      Caption         =   "Click For Explanation"
      Height          =   1095
      Left            =   9720
      TabIndex        =   2
      Top             =   3600
      Width           =   2655
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "Begin Program "
      Height          =   1335
      Left            =   9840
      TabIndex        =   1
      Top             =   5760
      Width           =   2295
   End
   Begin VB.PictureBox PicResults 
      Height          =   12015
      Left            =   600
      ScaleHeight     =   11955
      ScaleWidth      =   7995
      TabIndex        =   0
      Top             =   360
      Width           =   8055
   End
End
Attribute VB_Name = "Introduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdIntro_Click()
PicResults.Print "This program allows the user to interactively view"
PicResults.Print "and explore the different aspects of Real Madrid Club De Futbol"
PicResults.Print "Each button clicked will allow the user to discover interesting"
PicResults.Print "facts about the club.This program consist of the club's history"
PicResults.Print "It also show the statistics and information about the starters"
PicResults.Print "that plays for Real Madrid. It also has an accessories button"
PicResults.Print "that allows you to buy some of the gears and equipments used "
PicResults.Print "by the players. The program also consist of a trivia page that"
PicResults.Print "asks you simple questions about the club You keep guessing till"
PicResults.Print "you get the right answer"
PicResults.Print ""
PicResults.Print "Hope you enjoy our work"
PicResults.Print "This program was written by Seyi Alabi and Maxi Berger"
End Sub

Private Sub cmdProceed_Click()
OpenPage.Show
Form1.Hide
Gallery.Hide
Information.Hide
PlayersStat.Hide
Statistics.Hide
Trivia.Hide
End Sub

