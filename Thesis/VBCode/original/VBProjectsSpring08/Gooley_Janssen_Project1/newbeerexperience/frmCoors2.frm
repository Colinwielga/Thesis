VERSION 5.00
Begin VB.Form frmCoors 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Coors Brewing Co. Form"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13650
   LinkTopic       =   "Form2"
   ScaleHeight     =   9420
   ScaleWidth      =   13650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMainMenu 
      BackColor       =   &H00FF0000&
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9600
      Width           =   2775
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF0000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9600
      Width           =   2895
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Left            =   3240
      Picture         =   "frmCoors2.frx":0000
      ScaleHeight     =   8595
      ScaleWidth      =   13515
      TabIndex        =   7
      Top             =   480
      Width           =   13575
   End
   Begin VB.CommandButton cmdBlueMoon 
      BackColor       =   &H00FF0000&
      Caption         =   "Blue Moon"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7560
      Width           =   3015
   End
   Begin VB.CommandButton cmdKillians 
      BackColor       =   &H00FF0000&
      Caption         =   "Killian's Irish Red"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6600
      Width           =   3015
   End
   Begin VB.CommandButton cmdKeystone 
      BackColor       =   &H00FF0000&
      Caption         =   "Keystone Premium"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   3015
   End
   Begin VB.CommandButton cmdCoors 
      BackColor       =   &H00FF0000&
      Caption         =   "Coors"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   3015
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FF0000&
      Caption         =   "Search For and Count  All of the Beers Introduced Before 1990"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton cmdBrands 
      BackColor       =   &H00FF0000&
      Caption         =   "(Click First): List all of the Brands that Coors Offers "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblLearnCoors 
      BackColor       =   &H00FF0000&
      Caption         =   "Click on a Beer to Learn More About It"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   3840
      Width           =   2895
   End
End
Attribute VB_Name = "frmCoors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Beer Experience
'frm Coors
'Lauren Gooley and Tim Janssen
'March 20, 2008
'This form allows the user to learn more about the beers brewed at the Coors Brewery in Golden, Colorado.

Option Explicit
Dim Beers(1 To 100) As String, Dates(1 To 100) As Single, CTR As Single


Private Sub cmdBlueMoon_Click()
MsgBox ("A refreshing, medium-bodied, unfiltered Belgian-style wheat ale spiced with fresh coriander and orange peel for a uniquely complex taste and an uncommonly smooth finish.")
End Sub
'This subroutine opens the Coors beer array.
Private Sub cmdBrands_Click()
Dim J As Integer
Open App.Path & "\CoorsBeers.txt" For Input As #1
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Beers(CTR), Dates(CTR)
Loop
For J = 1 To CTR
    picResults.Print Beers(J)
Next J
Close #1
End Sub

Private Sub cmdCoors_Click()
MsgBox ("Coors beer, first introduced by Adolph Coors in April, 1874, is brewed in the Rockies for a uniquely crisp, clean, and drinkable Mile High Taste.")
End Sub

Private Sub cmdKeystone_Click()
MsgBox ("Introduced in 1989, Keystone is a popular-priced beer with exceptional market appeal.")
End Sub

Private Sub cmdKillians_Click()
MsgBox ("A traditional lager with an authentic Irish heritage, based on the Killian family's recipe created for the Killian's brewery in Enniscorthy Ireland in 1864.")
End Sub

Private Sub cmdMainMenu_Click()
frmCoors.Hide
Companies.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub

'This subroutine searches the Coors array and prints any brands that began production prior to 1990.
Private Sub cmdSearch_Click()
Dim YCTR As Single, J As Integer, Found As Boolean
picResults.Cls
Found = False
For J = 1 To CTR
    If Dates(J) < 1990 Then
        YCTR = YCTR + 1
        picResults.Print Beers(J)
        Found = True
    End If
Next J
If Found = True Then
    picResults.Print "There are "; YCTR; " beers that were introduced by Coors before 1990."
ElseIf Found = False Then
    picResults.Print "Sorry, Coors did not introduce any beers before 1990."
End If


End Sub

