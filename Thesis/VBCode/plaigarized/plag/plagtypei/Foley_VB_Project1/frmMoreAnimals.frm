VERSION 5.00
Begin VB.Form frmMoreAnimals
   BackColor       =   &H0000C000&
   Caption         =   "Form1"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12915
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   12915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit
      BackColor       =   &H008080FF&
      Caption         =   "Leave the Store"
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdPurchase
      BackColor       =   &H0080FF80&
      Caption         =   "Go to Checkout"
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdPreviousAnimals
      Caption         =   "View Previous Animals"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdDog
      BackColor       =   &H00FF80FF&
      Caption         =   "Soft Coated Wheaten Terrier"
      Height          =   735
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton cmdDeepSea
      BackColor       =   &H00FF80FF&
      Caption         =   "Deep Sea Predator"
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdFrog
      BackColor       =   &H00FF00FF&
      Caption         =   "Poisonous and Deadly Frog"
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdScorpion
      BackColor       =   &H00FF00FF&
      Caption         =   "Poisonous and Deadly Scorpion"
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Image imgDog
      Height          =   4080
      Left            =   9360
      Picture         =   "frmMoreAnimals.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2940
   End
   Begin VB.Image imgDeepSea
      Height          =   3300
      Left            =   6000
      Picture         =   "frmMoreAnimals.frx":A5A6
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3195
   End
   Begin VB.Image imgFrog
      Height          =   3285
      Left            =   3000
      Picture         =   "frmMoreAnimals.frx":16B8D
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2955
   End
   Begin VB.Image imgScorpion
      Height          =   3315
      Left            =   240
      Picture         =   "frmMoreAnimals.frx":235B7
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2595
   End
End
Attribute VB_Name = "frmMoreAnimals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Jimmy's Pet Store

'frmMoreAnimals

'Jimmy Foley

'February 24, 2010

'this form is the second lobby of animals where cusomers are able to see the availible animals and read a quick bit of information on each one



Private Sub cmdDeepSea_Click() ' there are several bottons like this one which tell about the animals

MsgBox ("This Deep Sea Predator eats other fish like Goldfish"), vbInformation

End Sub



Private Sub cmdDog_Click()

MsgBox ("The Soft Coated Wheaten Terrier is a freindly and fun loving dog. I, Jimmy the store owner, own one of these dogs. They are a little dim witted but a great addition to the family"), vbInformation

End Sub



Private Sub cmdFrog_Click()

MsgBox ("This Poisonous Frog is 2 inches long, it should be kept in an aquarium because it is poisonous"), vbInformation

End Sub



Private Sub cmdPreviousAnimals_Click() ' these bottoms allow the customer to move around the store

frmMoreAnimals.Hide

frmAnimals.Show

End Sub



Private Sub cmdPurchase_Click()

frmCheckout.Show

frmMoreAnimals.Hide

End Sub



Private Sub cmdQuit_Click() ' Quit the program

End

End Sub



Private Sub cmdScorpion_Click()

MsgBox ("This Deadly Scorpion is 7 and a half inches long and can lunge at its victims from 2 feet away"), vbInformation

End Sub







