VERSION 5.00
Begin VB.Form frmEndScreen 
   BackColor       =   &H0000C000&
   Caption         =   "Congratulations!"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H000000FF&
      Caption         =   "         Congratulations! You have earned the title of Yon Dungeonier!"
      Height          =   2295
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Label lblEnd 
      Caption         =   "You have successfully escaped your imprisonment in the dungeon!                    You shall be remembered for all eternity!"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   4815
   End
End
Attribute VB_Name = "frmEndScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this is the end screen of the game if it is successfully completed.
'It congratulates the player on their accomplishment.

Private Sub cmdEnd_Click()
'when the player is done, they can hit this button to exit the program
End
End Sub
