VERSION 5.00
Begin VB.Form frmStayPrice 
   Caption         =   "Stay Price"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   Picture         =   "frmStayPrice.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   9720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Room Size Selection"
      Height          =   1815
      Left            =   720
      TabIndex        =   3
      Top             =   6000
      Width           =   2295
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Height          =   1815
      Left            =   720
      TabIndex        =   2
      Top             =   3720
      Width           =   2295
   End
   Begin VB.PictureBox picPrice 
      Height          =   1935
      Left            =   4200
      ScaleHeight     =   1875
      ScaleWidth      =   4755
      TabIndex        =   1
      Top             =   2640
      Width           =   4815
   End
   Begin VB.CommandButton cmdGetPrice 
      Caption         =   "Get price of Stay"
      Height          =   1815
      Left            =   720
      TabIndex        =   0
      Top             =   1200
      Width           =   2295
   End
End
Attribute VB_Name = "frmStayPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: Hotel Checkin
'Form: Stay Price
'Authors: Ellen Jansen & Stuart Van Ess
'Date: March 28, 2008
'Purpose:   On this page, we make the customer confirm that they want to
'           purchase the room. We also allow them to input how many nights
'           they would like to stay to show how much they will be paying.
    
Option Explicit

Private Sub cmdAccept_Click()
    
'sets accepted as true originally.
    Accepted = True
'if you choose smoking, then you go to the smoking page
    If Smoking = "yes" Then
        frmSmokingDoubleDiagram.Show
        frmStayPrice.Hide
'if you dont choose smoking, then you go to the non smoking page
    Else
        frmNonSmokingDoubleDiagram.Show
        frmStayPrice.Hide
    End If
        
End Sub

Private Sub cmdGetPrice_Click()
'Clear the picture box named Price
    picPrice.Cls
    
'asks the user to input how many nights they would like to stay
    NumberOfNights = InputBox("How many nights would you like to stay?", "Nights", "")

'calculates the cost for someone to stay "NunmberOfNights" nights according to
'the room they chose(SelectedPrice)
    CostOfStay = NumberOfNights * SelectedPrice
    
'Prints in the picture box the cost before taxes
    picPrice.Print "Your price for "; NumberOfNights; " nights before taxes is "; FormatCurrency(CostOfStay)
    
    picPrice.Print ""
    picPrice.Print ""
    picPrice.Print ""
    picPrice.Print ""
    picPrice.Print ""
    picPrice.Print ""
    picPrice.Print "*************************************************************"
    picPrice.Print "Please select accept if you would like to purchase this room."
End Sub

Private Sub cmdReturn_Click()
'if they do not want to purchase the room, they can return to the room size menu
    picPrice.Cls
    frmStayPrice.Hide
    frmRoomSize.Show
    
    
    
End Sub

