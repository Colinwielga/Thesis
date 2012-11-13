VERSION 5.00
Begin VB.Form frmDessert 
   BackColor       =   &H00800000&
   Caption         =   "Form1"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9765
   FillColor       =   &H0080FFFF&
   ForeColor       =   &H00FF00FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back!!"
      Height          =   495
      Left            =   8160
      TabIndex        =   12
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Click to go to the sorting program!!"
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   3000
      Width           =   3375
   End
   Begin VB.PictureBox picDisplay 
      BackColor       =   &H00FF00FF&
      Height          =   2895
      Left            =   1680
      ScaleHeight     =   2835
      ScaleWidth      =   3555
      TabIndex        =   10
      Top             =   4320
      Width           =   3615
   End
   Begin VB.CommandButton cmdExitt 
      Caption         =   "Exit"
      Height          =   495
      Left            =   8160
      TabIndex        =   9
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdRootBeer 
      Caption         =   "Root Beer Float"
      Height          =   735
      Left            =   5400
      TabIndex        =   8
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdBanana 
      Caption         =   "Banana Split"
      Height          =   735
      Left            =   5400
      TabIndex        =   7
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdCheesecake 
      Caption         =   "Cheesecake"
      Height          =   735
      Left            =   5400
      TabIndex        =   6
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdDessertMenu 
      Caption         =   "Click here to see the dessert menu."
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   3375
   End
   Begin VB.PictureBox picResultsss 
      BackColor       =   &H00FF0000&
      FillColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2895
      Left            =   4440
      ScaleHeight     =   2835
      ScaleWidth      =   4995
      TabIndex        =   4
      Top             =   1320
      Width           =   5055
   End
   Begin VB.CommandButton cmdMalt 
      Caption         =   "Malt/Shake"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdCake 
      Caption         =   "Chocolate Cake"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdIceCream 
      Caption         =   "Ice Cream"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label lblDessertImages 
      BackColor       =   &H00FFFF00&
      Caption         =   "Click buttons to see images of Desserts!"
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Label lblWelcome2 
      Caption         =   $"Dessert.frx":0000
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   9255
   End
End
Attribute VB_Name = "frmDessert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'What's on the menu?
'frmDessert
'Michael Murakami and Kevin Schnese
'30 October 2006
'This form will load the names and prices of dessert options from a text file. It will also display pictures of the options when a button is pressed.
Private Sub cmdBack_Click()
'This button allows the user to go back between the dessert menu and the regular menu.
    frmMenu.Visible = True
    frmDessert.Visible = False
    frmSorting.Visible = False
End Sub
Private Sub cmdBanana_Click()
    'This will clear the picture box of any pervious entry and load the new picture.
    picDisplay.Cls
    picDisplay.Picture = LoadPicture(App.Path + "\banana-split.jpg")
End Sub
Private Sub cmdCake_Click()
    'This will clear the picture box of any pervious entry and load the new picture.
    picDisplay.Cls
    picDisplay.Picture = LoadPicture(App.Path + "\cake.jpg")
End Sub
Private Sub cmdCheesecake_Click()
    'This will clear the picture box of any pervious entry and load the new picture.
    picDisplay.Cls
    picDisplay.Picture = LoadPicture(App.Path + "\cheesecake.jpg")
End Sub
Private Sub cmdDessertMenu_Click()
    'When this button is pressed, it will load the dessert items and their prices from a text file. It does this via a loop that runs until the end of the file.
    counter = 0
    picResultsss.Cls
    picResultsss.Print "The items on the dessert menu are:"
    picResultsss.Print "************************************"
    Open App.Path & "\dessert.txt" For Input As #1
        Do Until EOF(1)
        Input #1, dessertitem, dessertprice
        counter = counter + 1
        dessertitems(counter) = dessertitem
        dessertprices(counter) = dessertprice
        picResultsss.Print counter; ".) "; dessertitem; " for "; FormatCurrency(dessertprice)
        picResultsss.Print " "
    Loop
    Close #1
End Sub
Private Sub cmdExitt_Click()
    'This will end the program.
    MsgBox "Thank you!", , "THANK YOU!!!"
    End
End Sub
Private Sub cmdIceCream_Click()
    'This will clear the picture box of any pervious entry and load the new picture.
    picDisplay.Cls
    picDisplay.Picture = LoadPicture(App.Path + "\icecream.jpg")
End Sub
Private Sub cmdMalt_Click()
    'This will clear the picture box of any pervious entry and load the new picture.
    picDisplay.Cls
    picDisplay.Picture = LoadPicture(App.Path + "\milkshake.jpg")
End Sub
Private Sub cmdRootBeer_Click()
    'This will clear the picture box of any pervious entry and load the new picture.
    picDisplay.Cls
    picDisplay.Picture = LoadPicture(App.Path + "\rootbeer.jpg")
End Sub
Private Sub cmdSort_Click()
'This allows the user to go from the dessert menu to the sorting menu.
    frmMenu.Visible = False
    frmDessert.Visible = False
    frmSorting.Visible = True
End Sub
