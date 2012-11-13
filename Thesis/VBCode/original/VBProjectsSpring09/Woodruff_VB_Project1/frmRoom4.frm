VERSION 5.00
Begin VB.Form frmRoom4 
   BackColor       =   &H80000017&
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H80000015&
      Caption         =   "Search for Items You Can Afford"
      Height          =   800
      Left            =   12240
      MaskColor       =   &H80000012&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6360
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H80000015&
      Caption         =   "Sort by Price"
      Height          =   800
      Left            =   12240
      MaskColor       =   &H80000012&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7440
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.CommandButton cmdHealth 
      BackColor       =   &H80000015&
      Caption         =   "Buy Potion"
      Height          =   800
      Left            =   12240
      MaskColor       =   &H80000012&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4560
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.CommandButton cmdTorch 
      BackColor       =   &H80000015&
      Caption         =   "Buy Torch"
      Height          =   800
      Left            =   12240
      MaskColor       =   &H80000012&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3600
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.CommandButton cmdSword 
      BackColor       =   &H80000015&
      Caption         =   "Buy Sword"
      Height          =   800
      Left            =   12240
      MaskColor       =   &H80000012&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.CommandButton cmdLook 
      BackColor       =   &H80000015&
      Caption         =   "Look at the Mechandise"
      Height          =   800
      Left            =   12240
      MaskColor       =   &H80000012&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      UseMaskColor    =   -1  'True
      Width           =   2500
   End
   Begin VB.PictureBox picRoom4 
      Height          =   6255
      Left            =   3120
      ScaleHeight     =   6195
      ScaleWidth      =   4275
      TabIndex        =   4
      Top             =   480
      Width           =   4335
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H80000015&
      Caption         =   "Go Back"
      Height          =   800
      Left            =   120
      MaskColor       =   &H80000012&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   2500
   End
   Begin VB.CommandButton cmdFoward 
      BackColor       =   &H80000015&
      Caption         =   "Next Room"
      Height          =   800
      Left            =   120
      MaskColor       =   &H80000012&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   2500
   End
   Begin VB.PictureBox picStore 
      BackColor       =   &H80000017&
      ForeColor       =   &H8000000F&
      Height          =   6255
      Left            =   7920
      ScaleHeight     =   6195
      ScaleWidth      =   3915
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "Action"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   735
      Left            =   12240
      TabIndex        =   12
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Movement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label lblStore 
      BackColor       =   &H80000017&
      Caption         =   $"frmRoom4.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   2295
      Left            =   3240
      TabIndex        =   0
      Top             =   6960
      Width           =   8415
   End
End
Attribute VB_Name = "frmRoom4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name:  Super Awesome Cave Adventure Game
'Form Name:  frmRoom4
'Author:  Peter Woodruff
'Date Written:  3-15-09
'Purpose:  This is a 'room' in the game.  It is where the user meets the merchant and can purhase
'a sword, a torch, and life.

Option Explicit
Dim I As Integer
Dim Item(1 To 23) As String
Dim Price(1 To 23) As Integer
Dim F As Integer
Dim tempPrice As Integer
Dim tempItem As String
Dim CTR As Integer
Dim Pos As Integer
Dim Pass As Integer
Dim T As Integer
Dim Sum As Single

Private Sub cmdBack_Click()
    
    'User leaves to room 3
    frmRoom4.Visible = False
    frmRoom3.Visible = True
    
End Sub

Private Sub cmdFoward_Click()
    
    'User leaves to room 5
    frmRoom4.Visible = False
    frmRoom5.Visible = True
    
End Sub

Private Sub cmdHealth_Click()
    
    'User can buy 1 life for 1 coin
    If Coins >= 1 Then
        Coins = Coins - 1
        Life = Life + 1
        MsgBox "You gained 1 life.", , ""
    Else
        MsgBox "You don't have enough coins...", , ""
    End If
    
End Sub

Private Sub cmdLook_Click()

    I = 0
    Sum = 0
    
    picStore.Cls
        
        'Loads file into array
        Open App.Path & "\Store.txt" For Input As #1
        
            Do While Not EOF(1)
                I = I + 1
                Input #1, Item(I), Price(I)
            Loop
    
        Close 1
        
        'Prints user's coins
        picStore.Print "Your Coins:  "; Coins
        picStore.Print ""
        picStore.Print "Item"; Tab(30); "Price"
        picStore.Print ""
    
    F = 0
        'Prints store list
        Do Until F = 23
            F = F + 1
            picStore.Print Item(F); Tab(30); Price(F)
            Sum = Sum + Price(F)
        Loop

    picStore.Print ""
    picStore.Print "Buy All"; Tab(30); Sum
    
    'Reveals now-useful actions
    cmdSword.Visible = True
    cmdTorch.Visible = True
    cmdHealth.Visible = True
    cmdSort.Visible = True
    cmdSearch.Visible = True
    
    
End Sub

Private Sub cmdSearch_Click()

    CTR = 23
    
    'Searchs for and selects items costing less or equal to than user's coins
    picStore.Cls
    picStore.Print "Your Coins:  "; Coins
    picStore.Print ""
    picStore.Print "Item"; Tab(30); "Price"
    picStore.Print ""
    
    For T = 1 To CTR
        If Price(T) <= Coins Then
            picStore.Print Item(T); Tab(30); Price(T)
        End If
    Next T

    
End Sub

Private Sub cmdSort_Click()
    
    'Sorts items by price using bubble sort
    CTR = 23
    
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If Price(Pos) > Price(Pos + 1) Then
            
                tempItem = Item(Pos)
                Item(Pos) = Item(Pos + 1)
                Item(Pos + 1) = tempItem
                
                tempPrice = Price(Pos)
                Price(Pos) = Price(Pos + 1)
                Price(Pos + 1) = tempPrice
                
            End If
        Next Pos
    Next Pass
    
    T = 0
    
    picStore.Cls
    picStore.Print "Your Coins:  "; Coins
    picStore.Print ""
    picStore.Print "Item"; Tab(30); "Price"
    picStore.Print ""
    
    For T = 1 To CTR
        picStore.Print Item(T); Tab(30); Price(T)
    Next
    
    picStore.Print ""
    picStore.Print "Buy All"; Tab(30); Sum
    
End Sub

Private Sub cmdSword_Click()
    
    'User can by sword for 10 coins
    If Coins >= 10 Then
        Coins = Coins - 10
        Sword = True
        MsgBox "You bought the sword.  It's pretty cool.", , ""
        cmdSword.Visible = False
    Else
        MsgBox "You don't have enough coins...", , ""
        
    End If
    
End Sub

Private Sub cmdTorch_Click()

    'User can by torch for 10 coins
    If Coins >= 10 Then
        Coins = Coins - 10
        Light = True
        MsgBox "You bought the torch.  It's pretty bright.", , ""
        cmdTorch.Visible = False
    Else
        MsgBox "You don't have enough coins...", , ""
        
    End If
    
End Sub

Private Sub Form_Load()

picRoom4.Picture = LoadPicture(App.Path & "\merchant.jpg")

End Sub

