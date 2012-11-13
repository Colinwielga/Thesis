VERSION 5.00
Begin VB.Form frmcited 
   BackColor       =   &H00800000&
   Caption         =   "Works Cited"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picresults2 
      Height          =   2535
      Left            =   7200
      Picture         =   "frmcited.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   2835
      TabIndex        =   4
      Top             =   3720
      Width           =   2895
   End
   Begin VB.CommandButton CmdBack 
      BackColor       =   &H000000FF&
      Caption         =   "Go Back"
      Height          =   3015
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton cmdthanks 
      BackColor       =   &H000000FF&
      Caption         =   "Special Thanks"
      Height          =   3015
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   3255
   End
   Begin VB.CommandButton cmdsource 
      BackColor       =   &H000000FF&
      Caption         =   "Sources Used"
      Height          =   3015
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
   Begin VB.PictureBox picresults 
      Height          =   2535
      Left            =   240
      ScaleHeight     =   2475
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   3720
      Width           =   6855
   End
End
Attribute VB_Name = "frmcited"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Super Tecmo Database
'Form name: frmcited
'Author: Nate Johnson & Kevin Klein
'Date Written: October 11th, 2006
'Objective of project: This project will allow its users to learn more about the game of football
'and will also allow them the oppurtunity to learn how to play the game of football with the Nintendo
'video game, Tecmo Super Bowl.
'Objective of form: This form allows the user to see the sources we used in creating
'this project, we also allow the user to see a section where we send out special thanks.




Private Sub CmdBack_Click()
frmcited.Hide
frmtutorial.Show 'shows the new form
End Sub

Private Sub cmdsource_Click() 'displays sources used in pic box
picresults.Cls
picresults.Print "Extra Tutorials and Source Code Examples Found At:"
picresults.Print "www.FreeVBCode.com"
picresults.Print "www.VBcode.com"
picresults.Print "Visual Basic Developer Center @ http://msdn2.microsoft.com/en-us/vbasic/default.aspx"
picresults.Print "http://www.programmersheaven.com/"


End Sub

Private Sub cmdthanks_Click() 'displays special thanks in pic box
picresults.Cls
picresults.Print "Thanks to Tecmo for making an awesome game."
picresults.Print "Thanks to GameFAQS.com for tons of extra information."
picresults.Print "Thanks to Romnation.net for all of our emulation needs."
picresults.Print "Thanks to Noreen Herzfeld for extra help and advice."
End Sub
