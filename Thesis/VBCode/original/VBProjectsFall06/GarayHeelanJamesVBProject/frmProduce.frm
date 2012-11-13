VERSION 5.00
Begin VB.Form frmProduce 
   BackColor       =   &H000080FF&
   Caption         =   "Produce"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picNumSurprise 
      Height          =   495
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   675
      TabIndex        =   9
      Top             =   3600
      Width           =   735
   End
   Begin VB.PictureBox picNumSalad 
      Height          =   495
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   675
      TabIndex        =   8
      Top             =   2520
      Width           =   735
   End
   Begin VB.PictureBox picNumStirFry 
      Height          =   495
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   675
      TabIndex        =   7
      Top             =   1440
      Width           =   735
   End
   Begin VB.PictureBox picNumFruit 
      Height          =   495
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   675
      TabIndex        =   6
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton cmdLogOut 
      Caption         =   "Log Out"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go to Another Aisle"
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   4560
      Width           =   2655
   End
   Begin VB.CommandButton cmdSurprise 
      Caption         =   "Produce Surprise!  $12.99"
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   3480
      Width           =   2655
   End
   Begin VB.CommandButton cmdSalad 
      Caption         =   "Salad Fixings $4.99"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   2655
   End
   Begin VB.CommandButton cmdStirFry 
      Caption         =   "Stir Fry Vegetables $5.99"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   2655
   End
   Begin VB.CommandButton cmdFruit 
      Caption         =   "Fruit Basket $9.99"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Image imgProduce 
      Height          =   5760
      Left            =   3960
      Picture         =   "frmProduce.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   5880
   End
End
Attribute VB_Name = "frmProduce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'James Delivers, CSCI 130 Visual Basic Project
'frmProduce
'Written by James Garay Heelan
'on 11-2-06
'the purpose of this form is to take the user's order for produce

Option Explicit
Dim NumFruit As Integer, NumStirFry As Integer, NumSalad As Integer, NumSurprise As Integer 'loading the variables to keep track of how many of each product the user would like

Private Sub cmdBack_Click()
    frmProduce.Hide 'hides the produce form
    frmGroceryStore.Show 'displays the main grocery store form with all the departments listed on it
End Sub

Private Sub cmdFruit_Click()
    
    picNumFruit.Cls 'clears the picturebox with the quantity of fruit baskets
    NumFruit = NumFruit + 1 'increases the quantity of fruit baskets ordered by 1
    picNumFruit.Print NumFruit ' displays in a picturebox the number of fruit baskets in the order

    Sum = Sum + 9.99 'adds the cost of the fruit basket to the total order amount
    Open App.Path & "/PurchasedItems.txt" For Append As #4 'opens the shopping cart file to be written in
        Write #4, "Fruit Basket", 9.99 'records in the shopping cart file what product is in it and how much it costs
    Close #4 'closes the shopping cart file

End Sub

Private Sub cmdLogOut_Click()
    End 'Exits the program
End Sub

Private Sub cmdSalad_Click()

    picNumSalad.Cls 'clears the picrutrebox with the quantity of salad fixings
    NumSalad = NumSalad + 1 'increases the quantity of salad fixings ordered by 1
    picNumSalad.Print NumSalad 'displays in a picturebox the total number of salad fixings in the shopping cart

    Sum = Sum + 4.99 'adds the costs of the salad fixings to the total order amount
    Open App.Path & "/PurchasedItems.txt" For Append As #4 'opens shopping cart file to be written into
        Write #4, "Salad Fixings", 4.99 'records in the shopping cart file what is being placed in it and how much it costs
    Close #4

End Sub

Private Sub cmdStirFry_Click()

    picNumStirFry.Cls 'clears the picturebox with the quantiy of stirfry displayed in it
    NumStirFry = NumStirFry + 1 ' increases the quantity of stirfry vegetables in the order by 1
    picNumStirFry.Print NumStirFry 'displays in a picturebox the total number of stirfry in the shopping cart
    
    Sum = Sum + 5.99 'adds the costs of the stirfry to the total order amount
    Open App.Path & "/PurchasedItems.txt" For Append As #4 'opens the shopping cart file to be written into
        Write #4, "Stir Fry Vegetables", 5.99 'records in the shopping cart file what is being placed in it and how much it costs
    Close #4

End Sub

Private Sub cmdSurprise_Click()

    picNumSurprise.Cls 'clears the picturebox with the quantity of produce surprise! displayed in it
    NumSurprise = NumSurprise + 1 'increases the quantity of produce surprise! in the order by 1
    picNumSurprise.Print NumSurprise 'displays in a picturebox the total number of produce surprise! in the shopping cart
    
    Sum = Sum + 12.99 'adds the cost of the produce surprise! to the total order amount
    Open App.Path & "/PurchasedItems.txt" For Append As #4 'opens the shopping cart file to be written into
        Write #4, "Produce Surprise!", 12.99 'records in the shopping cart file what is being placed into it and how much it costs
    Close #4

End Sub
